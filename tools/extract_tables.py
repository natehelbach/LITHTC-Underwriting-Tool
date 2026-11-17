"""Extract the constant (typed) data of the unit-mix and lookup-library tabs.

Writes data/<deal>/tables/:
  unit_takeoff.json          every typed cell on "Unit Takeoff" (variant rows, plan takeoff, labels)
  unit_mix_summary.json      every typed cell on "Unit Mix Summary" (AMI x type mix, scenarios, TCAC grid)
  utility_allowances.json    utility allowance schedule by city (B11:M204, "Actual" columns I:M)
  lihtc_tbl_tables.json      county TBL / 100% AMI rent / FMR tables (rows 11-68)
  federal_minimum_opex.json  TCAC operating-cost minimums (typed cells)

Only typed values are extracted; formula cells are computed in uw/sheets/unit_mix.py and
uw/sheets/tables.py.

Usage:  python tools/extract_tables.py [path to .xlsx] [deal]
"""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XLSX = ROOT / "reference" / "site8_source.xlsx"


def _jsonable(v):
    if isinstance(v, dt.datetime):
        return {"__date__": v.isoformat()}
    if isinstance(v, dt.date):
        return {"__date__": dt.datetime(v.year, v.month, v.day).isoformat()}
    return v


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=") or hasattr(v, "text")


def typed_cells(ws) -> dict:
    """Every non-formula, non-empty cell of a sheet, keyed by address."""
    out = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and not _is_formula(c.value):
                out[c.coordinate] = _jsonable(c.value)
    return out


def utility_allowances(ws) -> dict:
    """Schedule rows 11..204: city (B), utility (C), actual monthly allowance by bedroom (I:M).

    A few "actual" cells are formulas of the form IF(Sensitivities!$E$10="yes",0,N): they are
    stored as N and listed in `zero_in_application_case`.
    """
    rows, zeroed = [], []
    pat = re.compile(r'^=IF\(Sensitivities!\$E\$10="yes",0,([-\d.]+)\)$')
    for r in range(11, 205):
        actual = []
        for col in "IJKLM":
            v = ws[f"{col}{r}"].value
            if _is_formula(v):
                m = pat.match(v)
                if not m:
                    raise ValueError(f"Unexpected formula in Utility Allowances!{col}{r}: {v}")
                zeroed.append(f"{col}{r}")
                v = float(m.group(1))
            actual.append(v)
        rows.append({"row": r, "city": ws[f"B{r}"].value, "utility": ws[f"C{r}"].value,
                     "actual": actual})
    return {"source": "Utility Allowances!B11:M204", "bedrooms": ["Studio", "1 BR", "2 BR", "3 BR", "4 BR"],
            "rows": rows, "zero_in_application_case": zeroed}


def lihtc_tbl(ws) -> dict:
    def vals(r, cols):
        return [ws[f"{c}{r}"].value for c in cols]
    counties = []
    for r in range(11, 69):
        counties.append({
            "row": r,
            "county": ws[f"B{r}"].value,
            "tbl": vals(r, "CDEFG"),                 # threshold basis limits Studio..4BR
            "ami_county": ws[f"J{r}"].value,
            "ami_100": ws[f"K{r}"].value,            # 100% AMI income
            "max_rent_100": vals(r, "LMNOP"),        # 100% AMI max rents Studio..4BR
            "fmr_county": ws[f"R{r}"].value,
            "fmr": vals(r, "STUVW"),                 # county FMR Studio..4BR
        })
    return {"source": "LIHTC-TBL Tables!B11:W68", "counties": counties,
            "header_row9": {c: ws[f"{c}9"].value for c in "JKLMNOP"}}


def main(argv):
    xlsx = Path(argv[0]) if argv else DEFAULT_XLSX
    deal = argv[1] if len(argv) > 1 else "site8"
    wb = openpyxl.load_workbook(xlsx, data_only=False)
    out_dir = ROOT / "data" / deal / "tables"
    out_dir.mkdir(parents=True, exist_ok=True)
    files = {
        "unit_takeoff.json": {"source": "Unit Takeoff (typed cells)", "cells": typed_cells(wb["Unit Takeoff"])},
        "unit_mix_summary.json": {"source": "Unit Mix Summary (typed cells)",
                                  "cells": typed_cells(wb["Unit Mix Summary"])},
        "utility_allowances.json": utility_allowances(wb["Utility Allowances"]),
        "lihtc_tbl_tables.json": lihtc_tbl(wb["LIHTC-TBL Tables"]),
        "federal_minimum_opex.json": {"source": "Federal Minimum Opex (typed cells)",
                                      "cells": typed_cells(wb["Federal Minimum Opex"])},
    }
    for name, payload in files.items():
        (out_dir / name).write_text(json.dumps(payload, indent=1, ensure_ascii=False), encoding="utf-8")
        print("wrote", out_dir / name)


if __name__ == "__main__":
    main(sys.argv[1:])
