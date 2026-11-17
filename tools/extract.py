"""Extract a deal from the Excel underwriting model.

Writes, for one workbook:
  data/<deal>/inputs.json          Master Inputs (column E), keyed by readable names
  data/<deal>/excel_snapshot.json.gz  every cached cell value (the verification oracle)
  reference/sheets/<sheet>.txt     formula + cached value dump of every sheet (porting reference)
  reference/interfaces.json        per sheet: cells that OTHER sheets read from it

Usage:  python tools/extract.py "<path to .xlsx>" site8
"""
from __future__ import annotations

import datetime as dt
import gzip
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter

ROOT = Path(__file__).resolve().parents[1]


def jsonable(v):
    if isinstance(v, dt.datetime):
        return {"__date__": v.isoformat()}
    if isinstance(v, dt.date):
        return {"__date__": dt.datetime(v.year, v.month, v.day).isoformat()}
    if isinstance(v, dt.time):
        return v.isoformat()
    return v


def slug(s: str) -> str:
    s = re.sub(r"\(.*?\)", "", s or "")
    s = re.sub(r"[^A-Za-z0-9%]+", "_", s).replace("%", "pct").strip("_").lower()
    return re.sub(r"_+", "_", s)[:60].strip("_")


def camel_to_snake(s: str) -> str:
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", "_", s)
    return s.lower()


# Hand-picked keys where the label makes a poor name (row -> key).
KEY_OVERRIDES = {
    303: "bond_ratio",
    381: "residential_sf_adjustment",
}


def build_inputs(wf, wv):
    """Master Inputs column E -> ordered list of {row, key, section, label, segment, value}."""
    ws, vs = wf["Master Inputs"], wv["Master Inputs"]
    name_by_row = {}
    for n, d in wf.defined_names.items():
        m = re.fullmatch(r"'?Master Inputs'?!\$E\$(\d+)", d.attr_text)
        if m and n.startswith("MI_"):
            name_by_row[int(m.group(1))] = camel_to_snake(n[3:]).replace("__", "_")
    rows, used, section = [], set(), None
    for r in range(9, ws.max_row + 1):
        b, c, d, e = (ws.cell(r, k).value for k in (2, 3, 4, 5))
        if b:
            section = b
        if c is None and d is None:
            continue
        if b is None and e is None:  # section banner row
            continue
        key = KEY_OVERRIDES.get(r) or name_by_row.get(r) or slug(c + (" " + d if d else ""))
        if key in used:
            key = f"{key}_r{r}"
        used.add(key)
        formula = e if isinstance(e, str) and e.startswith("=") else None
        rows.append({
            "row": r, "key": key, "section": section, "label": c, "segment": d,
            "value": jsonable(vs.cell(r, 5).value),
            **({"excel_formula": formula} if formula else {}),
        })
    return rows


REF = re.compile(
    r"(?:'([^']+)'|([A-Za-z0-9_]+))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)")


def expand(rng: str):
    rng = rng.replace("$", "")
    c1, r1, c2, r2 = range_boundaries(rng if ":" in rng else f"{rng}:{rng}")
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            yield f"{get_column_letter(c)}{r}"


def main(xlsx: str, deal: str):
    wf = openpyxl.load_workbook(xlsx)
    wv = openpyxl.load_workbook(xlsx, data_only=True)
    out = ROOT / "data" / deal
    out.mkdir(parents=True, exist_ok=True)

    inputs = build_inputs(wf, wv)
    (out / "inputs.json").write_text(json.dumps(
        {"source_workbook": Path(xlsx).name, "inputs": inputs}, indent=1, default=str), encoding="utf-8")

    snap = {}
    for ws in wv.worksheets:
        snap[ws.title] = {c.coordinate: jsonable(c.value)
                          for row in ws.iter_rows() for c in row if c.value is not None}
    with gzip.open(out / "excel_snapshot.json.gz", "wt", encoding="utf-8") as f:
        json.dump(snap, f, default=str)

    refdir = ROOT / "reference" / "sheets"
    refdir.mkdir(parents=True, exist_ok=True)
    readers = defaultdict(lambda: defaultdict(set))  # source sheet -> cell -> {reader sheets}
    for ws in wf.worksheets:
        vs = wv[ws.title]
        lines = []
        for row in ws.iter_rows():
            parts = []
            for c in row:
                if c.value is None:
                    continue
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    parts.append(f"{c.coordinate}: {v} -> {vs[c.coordinate].value!r}")
                    for m in REF.finditer(v):
                        src = m.group(1) or m.group(2)
                        if src != ws.title and src in wf.sheetnames:
                            cells = list(expand(m.group(3)))
                            if len(cells) <= 20000:
                                for cell in cells:
                                    readers[src][cell].add(ws.title)
                else:
                    parts.append(f"{c.coordinate}= {v!r}")
            if parts:
                lines.append(" | ".join(parts))
        fn = re.sub(r"[^A-Za-z0-9 &_-]", "", ws.title).strip() + ".txt"
        (refdir / fn).write_text("\n".join(lines) + "\n", encoding="utf-8")

    names = {n: d.attr_text for n, d in wf.defined_names.items()}
    (ROOT / "reference" / "defined_names.json").write_text(json.dumps(names, indent=1), encoding="utf-8")
    iface = {s: {cell: sorted(r) for cell, r in sorted(cells.items())} for s, cells in readers.items()}
    (ROOT / "reference" / "interfaces.json").write_text(json.dumps(iface, indent=0), encoding="utf-8")
    print(f"inputs: {len(inputs)} rows; snapshot sheets: {len(snap)}; interfaces: {len(iface)} sheets")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "site8")
