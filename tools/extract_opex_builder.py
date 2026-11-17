"""Extract the typed inputs and grid layout of the "OpEx Builder" tab.

    python tools/extract_opex_builder.py

Writes data/site8/tables/opex_builder.json with:

  constants     every typed (non-formula) cell on the tab: comp-property data,
                comp line-item amounts, comments, the builder's selections/overrides.
  live          for each column, the rows (56-175) that hold a formula. The comp
                grid is ragged (a comp that has no figure for a line simply has no
                cell there), and blank vs. zero matters to MIN/AVERAGE/MAX, so the
                Python port uses this to know which cells exist.
  age_adjust    comp adjustment factors of the form 1 - age/10*0.1 (+ bump), by cell.
  size_adjust   comp adjustment factors of the form 1 - (units - comp units)/100*0.1.
  passthrough   "Adjust gross" cells that take the raw gross times a factor instead
                of the standard gross * size-scale * adjustment.

Formulas that fit none of the patterns are printed; uw/sheets/opex.py hand-codes them.
"""
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "reference" / "site8_source.xlsx"
OUT = ROOT / "data" / "site8" / "tables" / "opex_builder.json"
TAB = "OpEx Builder"

BLOCKS = ["P", "W", "AD", "AK", "AR", "AY", "BF", "BM", "BT"]   # comp name columns (row 65)
HEADS = {73, 88, 134, 143, 161, 164, 168}                          # section subtotal rows


def _json(v):
    if isinstance(v, dt.datetime):
        return {"__date__": v.isoformat()}
    return v


def main():
    wb = openpyxl.load_workbook(SRC, data_only=False)
    ws = wb[TAB]
    constants, formulas = {}, {}
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                formulas[c.coordinate] = v
            elif hasattr(v, "text"):            # array formula
                formulas[c.coordinate] = v.text
            else:
                constants[c.coordinate] = _json(v)

    live: dict[str, list[int]] = {}
    for addr in formulas:
        m = re.fullmatch(r"([A-Z]+)(\d+)", addr)
        col, r = m.group(1), int(m.group(2))
        if 56 <= r <= 175:
            live.setdefault(col, []).append(r)
    live = {k: sorted(v) for k, v in sorted(live.items(), key=lambda kv: ci(kv[0]))}

    age, size, passthrough, other = {}, [], {}, {}
    for addr, f in formulas.items():
        m = re.fullmatch(r"([A-Z]+)(\d+)", addr)
        c, r = ci(m.group(1)), int(m.group(2))
        if not (ci("P") <= c <= ci("BZ") and 71 <= r <= 172):
            continue
        base = max(ci(b) for b in BLOCKS if ci(b) <= c)
        off, B = c - base, gl(base)
        g = lambda k: gl(base + k)  # noqa: E731
        sf = f.replace("$", "")
        standard = {2: f"={B}68", 4: f"={g(3)}{r}/{g(2)}{r}",
                    5: f"={g(3)}{r}*{g(6)}69*{g(1)}{r}", 6: f"={g(5)}{r}/{g(6)}68"}
        if standard.get(off) == sf or sf.startswith("=SUBTOTAL(9,"):
            continue
        ma = re.fullmatch(r"=100%-\(YEAR\(TODAY\(\)\)-" + g(4) + r"8\)/10\*0\.1([+-]0\.1)?", sf)
        if off == 1 and ma:
            age[addr] = float(ma.group(1) or 0)
            continue
        if off == 1 and sf == f"=1-({g(6)}68-{B}68)/100*0.1":
            size.append(addr)
            continue
        mp = re.fullmatch("=" + g(3) + str(r) + r"(?:([*/])([\d.]+))?", sf)
        if off == 5 and mp:
            k = float(mp.group(2) or 1)
            passthrough[addr] = 1 / k if mp.group(1) == "/" else k
            continue
        other[addr] = f

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({
        "source": f"{SRC.name} / {TAB}",
        "constants": constants,
        "live": live,
        "age_adjust": age,
        "size_adjust": sorted(size),
        "passthrough": passthrough,
    }, indent=1), encoding="utf-8")
    print(f"wrote {OUT}: {len(constants)} constants, {sum(map(len, live.values()))} live cells")
    print("hand-coded formulas (not pattern-classified):")
    for a, f in other.items():
        print(f"  {a}: {f}")


if __name__ == "__main__":
    main()
