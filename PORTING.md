# Porting an Excel tab to Python

The Excel underwriting model (`reference/site8_source.xlsx`) is being replaced by
this Python model, one tab at a time. Each tab becomes one module in `uw/sheets/`.
Examples to copy: `uw/sheets/program.py`, `uw/sheets/support.py`, `uw/sheets/master_inputs.py`.

## The contract

```python
from uw import xl
from uw.engine import Sheet, register

@register
class DevBudget(Sheet):
    name = "Dev Budget"            # exact Excel tab name

    def compute(self, ctx) -> dict:
        o = {}
        ...
        return o                   # {"H37": 224828849.99, "C13": "Acquisition Costs", ...}
```

`compute` returns the cells of the tab it calculates, keyed by Excel address.
It **must** include every cell another core tab reads (the tab's *exports*,
listed in `reference/interfaces.json`), including text labels and header cells
that fall inside ranges other tabs read. Returning more cells is better: every
returned cell is checked against Excel.

## Reading values

| Need | Use |
|---|---|
| Master Inputs value | `ctx.mi["credit_type"]` (keys in `data/site8/inputs.json`; find a key by its `row`) |
| One cell of another tab | `ctx.ref("Dev Budget", "H37")` |
| 1-D range of another tab | `ctx.col("Rent Builder", "F12:F39")` |
| 2-D range | `ctx.grid("Unit Takeoff", "B6:S94")` (list of rows) |
| Defined names like `MI_ConstrMonths` | they point at Master Inputs: `ctx.mi["constr_months"]` (see `reference/defined_names.json`) |

Never read your own tab through `ctx`; compute its cells in Python. Never import
another tab's module; go through `ctx` so the isolated check and the full model
both work.

## Style

* Write it the way an analyst would explain it: named intermediate values,
  loops over rows/months/years, short comments on the business meaning.
  Not a transliteration of cell soup. Store results into `o[...]` as you go.
* Use `uw.xl` for anything where Excel differs from Python: `ROUND` (half away
  from zero), `ROUNDUP/ROUNDDOWN`, `PMT/IPMT/PPMT/NPV/IRR/XIRR`, `EDATE/EOMONTH`,
  `SUMIF(S)/SUMPRODUCT/COUNTIF`, `MATCH/INDEX/XLOOKUP`, `eq` (case-insensitive
  compare), `n` (blank -> 0), `div(a, b, default)`, `IFERROR(lambda: ..., fallback)`.
* Blank Excel cells arrive as `None`; wrap with `xl.n()` before arithmetic.
* Text compares in Excel are case-insensitive: use `xl.eq(a, "Yes")`.
* Dates arrive as `datetime.datetime`; use `xl.EDATE` etc. Return datetimes for date cells.
* Hardcoded numbers typed on the tab (not on Master Inputs) become named class
  constants or module-level tables with a comment saying where they came from.
  Large lookup tables (comp rows, utility allowance tables) go in
  `data/site8/tables/<name>.json`, loaded once; write a small extractor in
  `tools/` if you need one.
* Keep parity with Excel even where Excel looks wrong. Mark such places with
  `# EXCEL-QUIRK:` and a one-line explanation, and list them in your final report.

## Verify

```
python -m uw.verify "Dev Budget"      # isolated check: other tabs served from Excel
python -m uw.verify                   # every ported tab
```

Target: 0 mismatches and 0 missing exports. Tolerance is 1e-6 relative.
Formula + cached value dumps for every tab are in `reference/sheets/<tab>.txt`
(`A1: =formula -> cached value`, `A1= constant`). Read the original workbook
with openpyxl if you need more (e.g. merged cells, number formats).

## Circular references

Dev Budget, Flow of Funds, Financing, LIHTC and Sources & Uses reference each
other. Do not resolve the circularity inside a module: just read the other
tabs' cells through `ctx`. `run_model` iterates all modules to a fixed point.
