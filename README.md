# Heartwood underwriting model (Python)

A Python replacement for the Excel 4% LIHTC underwriting model, rendered as HTML.
The starting point and reference is the Site 8 model
(`reference/site8_source.xlsx`, 21101 Victory Blvd, Canoga Park).

## Run

```
pip install openpyxl
python -m uw.serve          # local app at http://localhost:8765: edit inputs, recalculate
python -m uw.report --open  # static HTML report -> out/site8.html
python -m uw.verify         # check every ported tab against the Excel model
python -m uw.verify --model # run the full Python model and compare to Excel
```

## Layout

| Path | What |
|---|---|
| `data/<deal>/inputs.json` | Master Inputs for a deal (the only thing you edit per deal) |
| `uw/sheets/*.py` | One module per Excel tab, same names and calculation order as the workbook |
| `uw/engine.py` | Runs the modules to a fixed point (handles the Excel circular references) |
| `uw/xl.py` | Excel-compatible functions (ROUND, PMT, IRR, EDATE, SUMIFS, ...) |
| `uw/report.py`, `uw/serve.py` | HTML rendering and the local web app |
| `tools/extract.py` | Pulls inputs and a verification snapshot out of an Excel model |
| `reference/` | Source workbook, per-tab formula dumps, cross-tab interface map |
| `PORTING.md` | How a tab is ported and verified |

## How parity is guaranteed

`tools/extract.py` stores every cached value in the workbook. Each tab's module is
checked on its own (other tabs served from Excel), then the whole model is run in
Python and compared to Excel. Places where the Excel logic looks wrong are kept
for parity and marked `# EXCEL-QUIRK:` in the code, so they can be fixed
deliberately once the port is complete.

## Hosting (Vercel)

`api/index.py` serves the same app as `python -m uw.serve` as a Vercel Python
function; `vercel.json` routes every path to it. The app is stateless: each
Recalculate posts the full input set, so nothing is stored between requests.
Connect the GitHub repo in Vercel (Add New → Project → Import) with the default
settings (Framework Preset: Other, no build command). Every push to `main` deploys.
