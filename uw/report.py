"""Render a model run as a self-contained HTML page.

    python -m uw.report                 # writes out/site8.html
    python -m uw.report --open          # ...and opens it
"""
from __future__ import annotations

import datetime as dt
import html
import sys
import webbrowser
from pathlib import Path

from uw.engine import ROOT, Inputs, ModelResult, Snapshot, run_model, same

SU = "Sources & Uses"


def money(v, digits=0):
    if not isinstance(v, (int, float)):
        return "–"
    s = f"{abs(v):,.{digits}f}"
    return f"(${s})" if v < 0 else f"${s}"


def pct(v, digits=1):
    return f"{v * 100:.{digits}f}%" if isinstance(v, (int, float)) else "–"


def num(v, digits=0):
    return f"{v:,.{digits}f}" if isinstance(v, (int, float)) else "–"


def esc(v):
    return html.escape("" if v is None else str(v))


def _rows(res: ModelResult, sheet, rows, label_col, value_col, pct_col=None, per_unit_col=None):
    out = []
    for r in rows:
        v = res.get(sheet, f"{value_col}{r}")
        lab = res.get(sheet, f"{label_col}{r}")
        if lab in (None, "") or not isinstance(v, (int, float)):
            continue
        out.append((lab, v, res.get(sheet, f"{pct_col}{r}") if pct_col else None,
                    res.get(sheet, f"{per_unit_col}{r}") if per_unit_col else None))
    return out


def _table(rows, total=None, cols=("", "Amount", "% of total", "Per unit")):
    head = "".join(f"<th>{c}</th>" for c in cols)
    body = "".join(
        f"<tr><td>{esc(l)}</td><td class=n>{money(v)}</td><td class=n>{pct(p)}</td>"
        f"<td class=n>{money(u)}</td></tr>" for l, v, p, u in rows)
    if total:
        l, v, p, u = total
        body += (f"<tr class=total><td>{esc(l)}</td><td class=n>{money(v)}</td>"
                 f"<td class=n>{pct(p)}</td><td class=n>{money(u)}</td></tr>")
    return f"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


def render(res: ModelResult, snap: Snapshot | None = None, editable: bool = False,
           changed: set[str] | None = None) -> str:
    changed = changed or set()
    g = res.get
    project = res.inputs["project_name"]
    gap = g(SU, "E58")
    units = g("Program", "E64")
    tdc = g(SU, "E26")

    kpis = [
        ("Surplus / (gap)", money(gap), "neg" if isinstance(gap, (int, float)) and gap < 0 else "pos"),
        ("Total development cost", money(tdc), ""),
        ("Units", num(units), ""),
        ("Cost per unit", money(tdc / units if units else None), ""),
        ("Permanent loan", money(g(SU, "E47")), ""),
        ("Tax credit equity", money(g(SU, "E46")), ""),
        ("Tax-exempt bonds", money(g(SU, "E30")), ""),
        ("Deferred developer fee", money(g(SU, "E52")), ""),
    ]
    kpi_html = "".join(f"<div class='kpi {c}'><div class=k>{esc(k)}</div><div class=v>{v}</div></div>"
                       for k, v, c in kpis)

    uses = _rows(res, SU, range(12, 26), "D", "E", "F", "G")
    cons = _rows(res, SU, range(29, 43), "D", "E", "F", "G")
    perm = _rows(res, SU, range(46, 57), "D", "E", "F", None)

    # Unit mix from Program.
    mix = []
    for r, label in zip(range(60, 64), ["Studio", "1 BR", "2 BR", "3 BR"]):
        n_, sf = g("Program", f"E{r}"), g("Program", f"D{r}")
        if isinstance(n_, (int, float)) and n_:
            mix.append(f"<tr><td>{label}</td><td class=n>{num(n_)}</td><td class=n>{num(sf)}</td></tr>")
    mix.append(f"<tr class=total><td>Total</td><td class=n>{num(units)}</td>"
               f"<td class=n>{num(g('Program', 'D64'))} SF</td></tr>")

    # Key inputs, grouped by Master Inputs section.
    sections: dict[str, list] = {}
    for r in res.inputs.rows:
        v = res.inputs[r["key"]]
        if v in (None, ""):
            continue
        label = r["label"] + (f" — {r['segment']}" if r.get("segment") else "")
        if editable and not isinstance(v, dt.datetime):
            cls = " class=changed" if r["key"] in changed else ""
            cell = f"<input name='{esc(r['key'])}' value='{esc(v)}'{cls}>"
        else:
            if isinstance(v, dt.datetime):
                v = v.date().isoformat()
            elif isinstance(v, float):
                v = f"{v:,.4g}" if abs(v) < 1 else f"{v:,.2f}"
            cell = esc(v)
        sections.setdefault(r["section"], []).append(f"<tr><td>{esc(label)}</td><td class=n>{cell}</td></tr>")
    inputs_html = ("<form method=post action='/'><button>Recalculate</button> "
                   "<a class=muted href='/'>reset to base case</a>" if editable else "") + "".join(
        f"<details><summary>{esc(s)} <span class=muted>({len(rows)})</span></summary>"
        f"<table class=inputs><tbody>{''.join(rows)}</tbody></table></details>"
        for s, rows in sections.items()) + ("</form>" if editable else "")

    # Parity vs Excel (only meaningful for the source deal).
    parity = ""
    if snap is not None:
        checks = [(s, a) for s, cells in res.sheets.items() for a in cells]
        diff = [(s, a) for s, a in checks if not same(res.get(s, a), snap.get(s, a), rel=1e-5, abs_=1e-3)]
        fb = sorted({s for s, _ in res.fallback_reads})
        parity = (f"<p class=muted>Python vs Excel: {len(checks) - len(diff):,} of {len(checks):,} "
                  f"cells match across {len(res.sheets)} tabs. Model converged in {res.iterations} "
                  f"iterations{'' if res.converged else ' (NOT converged)'}."
                  + (f" Still read from the Excel snapshot: {esc(', '.join(fb))}." if fb else "")
                  + "</p>")

    return f"""<!doctype html>
<html lang=en><head><meta charset=utf-8>
<meta name=viewport content="width=device-width, initial-scale=1">
<title>{esc(project)} | Underwriting</title>
<style>
:root{{--bg:#fafaf8;--fg:#1d1d1b;--muted:#6b6b66;--line:#e3e2dc;--card:#fff;--neg:#b42318;--pos:#1a7f4b;--accent:#2f5d50}}
@media (prefers-color-scheme:dark){{:root{{--bg:#151614;--fg:#ecebe6;--muted:#9b9a93;--line:#2c2d2a;--card:#1d1e1b;--neg:#f2867b;--pos:#6fcf97;--accent:#8fc1ae}}}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--fg);font:15px/1.45 system-ui,-apple-system,Segoe UI,sans-serif}}
main{{max-width:1100px;margin:0 auto;padding:24px 16px 64px}}
h1{{font-size:22px;margin:0 0 4px}}h2{{font-size:16px;margin:32px 0 10px;color:var(--accent)}}
.muted{{color:var(--muted);font-size:13px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px;margin-top:18px}}
.kpi{{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px}}
.kpi .k{{font-size:12px;color:var(--muted)}}.kpi .v{{font-size:20px;font-weight:600;font-variant-numeric:tabular-nums}}
.kpi.neg .v{{color:var(--neg)}}.kpi.pos .v{{color:var(--pos)}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:24px}}@media (max-width:800px){{.grid2{{grid-template-columns:1fr}}}}
.scroll{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:8px}}
th,td{{padding:6px 10px;border-bottom:1px solid var(--line);text-align:left}}th{{font-size:12px;color:var(--muted);font-weight:500}}
td.n,th:not(:first-child){{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}}
tr.total td{{font-weight:600;border-top:2px solid var(--line)}}
details{{background:var(--card);border:1px solid var(--line);border-radius:8px;margin:6px 0;padding:6px 12px}}
summary{{cursor:pointer;font-weight:500}}
input{{width:9em;text-align:right;font:inherit;background:var(--bg);color:var(--fg);border:1px solid var(--line);border-radius:4px;padding:2px 6px}}
input.changed{{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent)}}
button{{font:inherit;background:var(--accent);color:var(--bg);border:0;border-radius:6px;padding:6px 14px;cursor:pointer;margin:4px 0 8px}}table.inputs{{border:0;margin-top:6px}}
</style></head><body><main>
<h1>{esc(project)}</h1>
<div class=muted>Underwriting model (Python) · run {dt.date.today().isoformat()}</div>
<div class=kpis>{kpi_html}</div>
<h2>Uses of funds</h2><div class=scroll>{_table(uses, ("Total uses", tdc, 1.0, g(SU, "G26")))}</div>
<div class=grid2>
<div><h2>Construction sources</h2><div class=scroll>{_table(cons, ("Total construction sources", g(SU, "E43"), 1.0, g(SU, "G43")))}</div></div>
<div><h2>Permanent sources</h2><div class=scroll>{_table(perm, ("Total permanent sources", g(SU, "E57"), 1.0, None), cols=("", "Amount", "%", ""))}</div></div>
</div>
<div class=grid2>
<div><h2>Unit mix</h2><table><thead><tr><th>Type</th><th>Units</th><th>Avg SF</th></tr></thead><tbody>{''.join(mix)}</tbody></table></div>
<div><h2>Inputs</h2>{inputs_html}</div>
</div>
<h2>Model check</h2>{parity or '<p class=muted>No Excel reference for this deal.</p>'}
</main></body></html>"""


def main(argv):
    deal = "site8"
    snap = Snapshot.load(deal)
    res = run_model(Inputs.load(deal), deal, fallback=snap)
    out = ROOT / "out"
    out.mkdir(exist_ok=True)
    path = out / f"{deal}.html"
    path.write_text(render(res, snap), encoding="utf-8")
    print(f"wrote {path}  (iterations={res.iterations}, converged={res.converged})")
    if "--open" in argv:
        webbrowser.open(path.as_uri())


if __name__ == "__main__":
    main(sys.argv[1:])
