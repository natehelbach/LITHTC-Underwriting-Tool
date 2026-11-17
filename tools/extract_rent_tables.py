"""Extract the typed data of the market-rent tabs (comp datasets and amenity library).

Writes data/<deal>/tables/:
  studio_comps.json        "Studio Comps"  one record per comp row (building, rent, SF, concessions, links)
  1br_comps.json           "1BR Comps"
  2br_comps.json           "2BR Comps"     (adds the 2nd-bath flag, column Y)
  amenity_premiums.json    "Amenity & Property Premiums": amenity matrix (rows 11-70) and
                           building coordinates / neighborhoods (rows 87-146)
  market_rent_survey.json  "Market Rent Survey" listing rows (9-71)

Each file has "rows" (typed values of the data rows, by field name) and "labels" (other typed
text on the tab: headers and notes, by cell address). Formula cells are not extracted; they are
computed in uw/sheets/comps.py. HYPERLINK formulas are kept as {"url", "label"} since they carry
data (the link) rather than logic.

Numeric constants outside the data rows (amenity premiums, subject amenity flags, class
thresholds) are small and live as named constants in uw/sheets/comps.py; this script prints
them so they can be checked against the module.

Usage:  python tools/extract_rent_tables.py [path to .xlsx] [deal]
"""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XLSX = ROOT / "reference" / "site8_source.xlsx"

COMP_FIELDS = {
    "B": "building", "C": "address", "E": "plan", "F": "include", "G": "sf", "H": "rent",
    "I": "special", "J": "free_months", "K": "one_time_credit", "L": "lease_term",
    "U": "photo", "V": "listing", "W": "phone",
}

AMENITY_COLS = [get_column_letter(c) for c in range(column_index_from_string("D"),
                                                      column_index_from_string("Q") + 1)]

# tab -> output file, list of (first_row, last_row or None = last typed row, {col: field})
SPECS = {
    "Studio Comps": ("studio_comps.json", [(9, None, COMP_FIELDS)]),
    "1BR Comps": ("1br_comps.json", [(9, None, COMP_FIELDS)]),
    "2BR Comps": ("2br_comps.json", [(9, None, {**COMP_FIELDS, "Y": "second_bath"})]),
    "Amenity & Property Premiums": ("amenity_premiums.json", [
        (11, 70, {"B": "building", **{c: f"amenity_{c}" for c in AMENITY_COLS},
                  "T": "vintage", "V": "class", "W": "class_rationale", "X": "amenity_listing"}),
        (87, 146, {"B": "building", "C": "address", "D": "lat", "E": "lon",
                   "F": "coord_source", "I": "neighborhood"}),
    ]),
    "Market Rent Survey": ("market_rent_survey.json", [
        (9, 71, {"B": "property", "C": "address", "D": "distance", "E": "unit_type", "F": "sf",
                 "G": "rent", "H": "listings", "I": "note", "K": "source", "L": "observed"}),
    ]),
}

_HYPERLINK = re.compile(r'^=HYPERLINK\("([^"]*)","([^"]*)"\)$')


def _jsonable(v):
    if isinstance(v, dt.datetime):
        return {"__date__": v.isoformat()}
    if isinstance(v, dt.date):
        return {"__date__": dt.datetime(v.year, v.month, v.day).isoformat()}
    return v


def _typed(v):
    """(is_data, value): typed constants and HYPERLINK formulas are data; other formulas are not."""
    if hasattr(v, "text"):          # array formula
        return False, None
    if isinstance(v, str) and v.startswith("="):
        m = _HYPERLINK.match(v)
        return (True, {"url": m.group(1), "label": m.group(2)}) if m else (False, None)
    return v is not None, _jsonable(v)


def extract_sheet(ws, blocks):
    rows_out, in_block, numeric_left = [], set(), {}
    for first, last, fields in blocks:
        if last is None:  # data rows end at the last row with a typed value in a mapped column
            last = max(r for r in range(first, ws.max_row + 1)
                       for c in fields if _typed(ws[f"{c}{r}"].value)[0])
        recs = []
        for r in range(first, last + 1):
            rec = {"row": r}
            for c, name in fields.items():
                ok, v = _typed(ws[f"{c}{r}"].value)
                if ok:
                    rec[name] = v
                in_block.add((c, r))
            recs.append(rec)
        rows_out.append(recs)
    labels = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or (get_column_letter(cell.column), cell.row) in in_block:
                continue
            ok, v = _typed(cell.value)
            if not ok or isinstance(v, dict):
                continue
            if isinstance(v, str):
                labels[cell.coordinate] = v
            else:
                numeric_left[cell.coordinate] = v
    return rows_out, labels, numeric_left


def main(argv):
    xlsx = Path(argv[0]) if argv else DEFAULT_XLSX
    deal = argv[1] if len(argv) > 1 else "site8"
    wb = openpyxl.load_workbook(xlsx)
    out_dir = ROOT / "data" / deal / "tables"
    out_dir.mkdir(parents=True, exist_ok=True)
    for tab, (fname, blocks) in SPECS.items():
        rows, labels, numeric = extract_sheet(wb[tab], blocks)
        doc = {"source": f"{tab} (typed cells)"}
        if len(rows) == 1:
            doc["rows"] = rows[0]
        else:  # amenity tab: matrix + locations
            matrix, doc["locations"] = rows
            for rec in matrix:  # amenity flags D:Q -> one list, in column order
                rec["amenities"] = [rec.pop(f"amenity_{c}", None) for c in AMENITY_COLS]
            doc["amenity_columns"] = AMENITY_COLS
            doc["amenity_matrix"] = matrix
        doc["labels"] = labels
        (out_dir / fname).write_text(json.dumps(doc, indent=1, ensure_ascii=False), encoding="utf-8")
        print(f"{fname}: {sum(len(r) for r in rows)} rows, {len(labels)} labels")
        if numeric:
            print(f"   numeric typed cells outside data rows (module constants): {numeric}")


if __name__ == "__main__":
    main(sys.argv[1:])
