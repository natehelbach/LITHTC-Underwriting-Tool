"""Pull display layout out of the workbook for the web pages.

    python tools/extract_layout.py reference/site8_source.xlsx site8

Writes data/<deal>/layout.json:
  inputs:  {row: {"fmt": number format of Master Inputs E<row>, "options": [dropdown list]}}
  sheets:  {sheet: {"cells": {addr: [constant-or-null, number format, bold]},
                    "hidden_rows": [...], "formulas": [addr, ...]}}
Constants (labels) come from the workbook; formula cells are filled from the model run.
"""
from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_SHEETS = ["Returns & Summary", "Operating Cash Flow"]


def jsonable(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return {"__date__": v.isoformat()}
    return v


def cells_of(sqref: str):
    for part in str(sqref).split():
        c1, r1, c2, r2 = range_boundaries(part)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                yield c, r


def list_options(formula: str, wv) -> list:
    f = formula.strip()
    if f.startswith('"'):
        return [s.strip() for s in f.strip('"').split(",")]
    sheet, rng = f.rsplit("!", 1)
    ws = wv[sheet.strip("'")]
    out = []
    for row in ws[rng.replace("$", "")]:
        for c in row:
            if c.value not in (None, "") and c.value not in out:
                out.append(c.value)
    return out


def main(xlsx: str, deal: str):
    wf = openpyxl.load_workbook(xlsx)
    wv = openpyxl.load_workbook(xlsx, data_only=True)

    mi = wf["Master Inputs"]
    inputs = {}
    for r in range(9, mi.max_row + 1):
        c = mi.cell(r, 5)
        if c.value is not None or mi.cell(r, 3).value is not None:
            inputs[r] = {"fmt": c.number_format}
    for dv in mi.data_validations.dataValidation:
        if dv.type != "list" or not dv.formula1:
            continue
        opts = list_options(dv.formula1, wv)
        for col, r in cells_of(dv.sqref):
            if col == 5:
                inputs.setdefault(r, {"fmt": "General"})["options"] = opts

    sheets = {}
    for name in OUTPUT_SHEETS:
        ws = wf[name]
        cells, formulas = {}, []
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                is_f = isinstance(c.value, str) and c.value.startswith("=")
                if is_f:
                    formulas.append(c.coordinate)
                cells[c.coordinate] = [None if is_f else jsonable(c.value),
                                       c.number_format, bool(c.font and c.font.b)]
        sheets[name] = {
            "cells": cells,
            "formulas": formulas,
            "hidden_rows": sorted(r for r, d in ws.row_dimensions.items() if d.hidden),
        }

    out = ROOT / "data" / deal / "layout.json"
    out.write_text(json.dumps({"inputs": inputs, "sheets": sheets}, separators=(",", ":")),
                   encoding="utf-8")
    print(f"wrote {out} ({out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main(*sys.argv[1:3])
