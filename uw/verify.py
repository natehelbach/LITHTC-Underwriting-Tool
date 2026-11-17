"""Verify ported tabs against the Excel model.

    python -m uw.verify                 # every ported tab, isolated
    python -m uw.verify "Dev Budget"    # one tab, full mismatch list
    python -m uw.verify --model         # full Python model vs Excel (integration)
"""
from __future__ import annotations

import sys

from uw.engine import Inputs, Snapshot, registry, run_model, same, verify_sheet


def main(argv: list[str]) -> int:
    snap = Snapshot.load()
    if "--model" in argv:
        res = run_model(Inputs.load(), fallback=snap)
        print(f"model: {res.iterations} iterations, converged={res.converged}")
        for sheet, err in res.errors.items():
            print(f"  ERROR in {sheet}: {err}")
        bad = 0
        for sheet, cells in res.sheets.items():
            mism = [(a, v, snap.get(sheet, a)) for a, v in cells.items()
                    if not same(v, snap.get(sheet, a), rel=1e-5, abs_=1e-3)]
            bad += len(mism)
            print(f"  {sheet:32s} {len(cells):6d} cells  {len(mism):5d} differ from Excel")
            for a, v, e in mism[:5]:
                print(f"      {a}: python={v!r} excel={e!r}")
        fb = sorted({s for s, _ in res.fallback_reads})
        if fb:
            print("still read from Excel snapshot:", ", ".join(fb))
        return 1 if bad else 0

    names = [a for a in argv if not a.startswith("-")] or sorted(registry())
    failed = 0
    for name in names:
        rep = verify_sheet(name, snap=snap, core_only="--strict" not in argv)
        print(rep.summary(limit=200 if len(names) == 1 else 15))
        failed += not rep.ok
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
