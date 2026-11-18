"""Excel download: the original workbook with the scenario's inputs written in.

The workbook is patched at the XML level, so formulas, formatting, charts, names
and data validation are untouched:
  * Master Inputs column E gets the edited input values;
  * every formula cell the Python model calculates gets its new cached value, so
    the file shows the scenario's numbers even before Excel recalculates
    (Protected View does not recalculate);
  * the workbook already asks Excel for a full recalculation on open
    (calcPr fullCalcOnLoad), with iterative calculation on, as in the source.
"""
from __future__ import annotations

import datetime as dt
import io
import math
import re
import zipfile
from html import escape
from pathlib import Path

from uw.engine import ROOT, Inputs, ModelResult

SOURCE = ROOT / "reference" / "site8_source.xlsx"
EPOCH = dt.datetime(1899, 12, 30)
CELL = re.compile(r'<c r="([A-Z]{1,3}\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def is_calculated(row: dict) -> bool:
    """A Master Inputs row whose value is a real formula (not just "=32100")."""
    f = row.get("excel_formula")
    return bool(f) and not re.fullmatch(r"=\s*-?[\d.]+", f)


def _sheet_paths(z: zipfile.ZipFile) -> dict[str, str]:
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target = {m.group(1): m.group(2) for m in re.finditer(
        r'<Relationship[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels)}
    target.update({m.group(2): m.group(1) for m in re.finditer(
        r'<Relationship[^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
    out = {}
    for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb):
        name = (m.group(1).replace("&amp;", "&").replace("&apos;", "'")
                .replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">"))
        t = target[m.group(2)].lstrip("/")
        out[name] = t if t.startswith("xl/") else "xl/" + t
    return out


def _attrs_without_t(attrs: str) -> str:
    return re.sub(r'\s+t="[^"]*"', "", attrs)


def _value_xml(v) -> tuple[str, str] | None:
    """(t attribute, <v> payload) for a cached value; None to leave the cell alone."""
    if isinstance(v, bool):
        return ' t="b"', f"<v>{int(v)}</v>"
    if isinstance(v, (dt.datetime, dt.date)):
        if not isinstance(v, dt.datetime):
            v = dt.datetime(v.year, v.month, v.day)
        return "", f"<v>{(v - EPOCH).total_seconds() / 86400!r}</v>"
    if isinstance(v, (int, float)):
        if isinstance(v, float) and not math.isfinite(v):
            return None
        return "", f"<v>{v!r}</v>"
    if isinstance(v, str):
        return ' t="str"', f"<v>{escape(v, quote=False)}</v>"
    return None


def _input_xml(v) -> tuple[str, str] | None:
    if isinstance(v, str):
        return ' t="inlineStr"', f'<is><t xml:space="preserve">{escape(v, quote=False)}</t></is>'
    if v is None:
        return "", ""
    return _value_xml(v)


def _patch_sheet(xml: str, cached: dict, inputs: dict) -> str:
    def sub(m):
        addr, attrs, inner = m.group(1), m.group(2), m.group(3) or ""
        if addr in inputs:  # an edited input replaces whatever the cell held
            x = _input_xml(inputs[addr])
            if x is None:
                return m.group(0)
            t, body = x
            return f'<c r="{addr}"{_attrs_without_t(attrs)}{t}>{body}</c>'
        if addr in cached and "<f" in inner:
            x = _value_xml(cached[addr])
            if x is None:
                return m.group(0)
            t, body = x
            f = re.search(r"<f\b[^>]*/>|<f\b[^>]*>.*?</f>", inner, re.S).group(0)
            return f'<c r="{addr}"{_attrs_without_t(attrs)}{t}>{f}{body}</c>'
        return m.group(0)
    return CELL.sub(sub, xml)


def build_xlsx(res: ModelResult, base: Inputs, changed: set[str]) -> bytes:
    if not changed:
        return SOURCE.read_bytes()
    edits = {f"E{r['row']}": res.inputs[r["key"]] for r in res.inputs.rows
             if r["key"] in changed and not is_calculated(r)}
    src = zipfile.ZipFile(SOURCE)
    paths = _sheet_paths(src)
    by_path = {p: s for s, p in paths.items()}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            sheet = by_path.get(info.filename)
            if sheet is not None and (sheet in res.sheets or sheet == "Master Inputs"):
                xml = data.decode("utf-8")
                xml = _patch_sheet(xml, res.sheets.get(sheet, {}),
                                   edits if sheet == "Master Inputs" else {})
                data = xml.encode("utf-8")
            out.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    return buf.getvalue()
