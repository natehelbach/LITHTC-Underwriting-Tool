"""HTML pages: an Inputs page (Master Inputs) and an Outputs page (summary + cash flows).

    python -m uw.report                 # writes out/site8.html (outputs, base case)
    python -m uw.report --open          # ...and opens it
"""
from __future__ import annotations

import datetime as dt
import html
import json
import re
import sys
import webbrowser
from functools import lru_cache

from openpyxl.utils import column_index_from_string, get_column_letter

from uw.engine import DATA, ROOT, Inputs, ModelResult, Snapshot, run_model, same
from uw.excel_export import is_calculated
from uw.numfmt import fmt, input_text

SU = "Sources & Uses"
RS = "Returns & Summary"
OCF = "Operating Cash Flow"


def esc(v):
    return html.escape("" if v is None else str(v))


@lru_cache(maxsize=None)
def layout(deal: str = "site8") -> dict:
    d = json.loads((DATA / deal / "layout.json").read_text(encoding="utf-8"))
    for s in d["sheets"].values():
        s["formulas"] = set(s["formulas"])
        s["hidden_rows"] = set(s["hidden_rows"])
    return d


def split(addr: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", addr)
    return m.group(1), int(m.group(2))


def cols(a: str, b: str) -> list[str]:
    return [get_column_letter(i) for i in range(column_index_from_string(a),
                                                column_index_from_string(b) + 1)]


# ---------------------------------------------------------------- sheet grids
class Grid:
    """Cell access for an output tab: labels from the workbook, numbers from the model."""

    def __init__(self, res: ModelResult, sheet: str, deal="site8"):
        self.res, self.sheet = res, sheet
        self.lay = layout(deal)["sheets"][sheet]

    def raw(self, addr):
        if addr in self.lay["formulas"]:
            v = self.res.get(self.sheet, addr)
            return None if v == "" else v
        c = self.lay["cells"].get(addr)
        if c is None:
            return None
        v = c[0]
        if isinstance(v, dict) and "__date__" in v:
            v = dt.datetime.fromisoformat(v["__date__"])
        return v

    def text(self, addr):
        c = self.lay["cells"].get(addr)
        return fmt(self.raw(addr), c[1] if c else "General")

    def bold(self, addr):
        c = self.lay["cells"].get(addr)
        return bool(c and c[2])

    def hidden(self, row):
        return row in self.lay["hidden_rows"]


def _cls(v):
    return " class='n neg'" if isinstance(v, (int, float)) and not isinstance(v, bool) and v < 0 \
        else " class=n"


def nice_title(t: str) -> str:
    """'LIHTC & DEVELOPER FEE' -> 'LIHTC & developer fee' (keeps acronyms and brackets)."""
    head, _, rest = t.partition("(")
    keep = {"LIHTC", "TDC", "NOI", "AMI", "DSCR"}
    s = " ".join(w if w in keep else w.lower() for w in head.split())
    return s[:1].upper() + s[1:] + (f" ({rest}" if rest else "")


def summary_sections(res: ModelResult) -> str:
    """Returns & Summary tab, one card per section (rows whose only cell is an ALL-CAPS title)."""
    g = Grid(res, RS)
    cs = cols("B", "G")
    sections, cur = [], None
    for r in range(4, 127):
        vals = {c: g.raw(f"{c}{r}") for c in cs}
        filled = [c for c in cs if vals[c] not in (None, "")]
        if not filled:
            continue
        b = vals["B"]
        if filled == ["B"] and isinstance(b, str) and b[:4].isupper():
            cur = [b, []]
            sections.append(cur)
            continue
        if cur is None:
            cur = ["", []]
            sections.append(cur)
        cur[1].append(r)

    cards = []
    for title, rows in sections:
        used = [c for c in cs if any(g.raw(f"{c}{r}") not in (None, "") for r in rows)]
        body = []
        for r in rows:
            vals = [g.raw(f"{c}{r}") for c in used]
            n_filled = sum(v not in (None, "") for v in vals)
            header = all(isinstance(v, str) or v is None for v in vals) and n_filled > 1 and \
                (vals[0] in (None, "", "AMI", "Use") or str(vals[0]).endswith(" sources"))
            bold = g.bold(f"B{r}") or str(vals[0] or "").lower().startswith(("total", "surplus"))
            tds = []
            for i, c in enumerate(used):
                a = f"{c}{r}"
                if header:
                    tds.append(f"<th>{esc(g.text(a))}</th>")
                elif i == 0:
                    tds.append(f"<td>{esc(g.text(a))}</td>")
                else:
                    tds.append(f"<td{_cls(g.raw(a))}>{esc(g.text(a))}</td>")
            body.append(f"<tr{' class=total' if bold and not header else ''}>{''.join(tds)}</tr>")
        if title:
            cards.append(f"<section class=card><h3>{esc(nice_title(title))}</h3>"
                         f"<div class=scroll><table>{''.join(body)}</table></div></section>")
        else:
            cards.append(f"<p class=muted>{''.join(body)}</p>")
    return "".join(cards)


def cash_flow(res: ModelResult) -> str:
    g = Grid(res, OCF)
    years = cols("J", "BQ")

    # Annual operating cash flow, rows 12-86.
    head = "".join(f"<th>Yr {esc(g.text(f'{c}12'))}</th>" for c in years)
    rows = []
    for r in range(13, 87):
        if g.hidden(r):
            continue
        d, e = g.raw(f"D{r}"), g.raw(f"E{r}")
        label, indent = (e, True) if e not in (None, "") else (d, False)
        vals = [g.raw(f"{c}{r}") for c in years]
        if label in (None, "") and all(v in (None, "") for v in vals):
            continue
        if label in (None, "") and not any(isinstance(v, (int, float)) for v in vals):
            continue
        rate = g.text(f"G{r}")
        sub = indent and not str(label).lower().startswith("total")
        tot = (not indent and any(v not in (None, "") for v in vals)) or \
            str(label).lower().startswith("total")
        section = not indent and all(v in (None, "") for v in vals)
        cells = "".join(f"<td{_cls(v)}>{esc(g.text(f'{c}{r}'))}</td>" for c, v in zip(years, vals))
        cls = "sec" if section else ("total" if tot else "")
        rows.append(f"<tr class='{cls}'><th class='lab{' ind' if sub else ''}'>{esc(label)}</th>"
                    f"<td class=n>{esc(rate)}</td>{cells}</tr>")
    annual = (f"<div class='scroll wide'><table class=cf><thead><tr><th class=lab></th>"
              f"<th>Rate</th>{head}</tr></thead><tbody>{''.join(rows)}</tbody></table></div>")

    # Investor cash flow summary and returns, rows 91-134. Labels sit in column C
    # (cash flow summary) or column E (returns blocks); the rest are values.
    data_cols = [c for c in cols("D", "X")
                 if any(g.raw(f"{c}{r}") not in (None, "") for r in range(91, 135))]
    rows = []
    for r in range(91, 135):
        lab_col = "C" if g.raw(f"C{r}") not in (None, "") else "E"
        label = g.raw(f"{lab_col}{r}")
        vals = {c: (None if c == lab_col else g.raw(f"{c}{r}")) for c in data_cols}
        if label in (None, "") and all(v in (None, "") for v in vals.values()):
            continue
        header = r == 91
        tds = [f"<th class=lab>{esc(g.text(f'{lab_col}{r}'))}</th>"]
        for c in data_cols:
            t = "" if c == lab_col else esc(g.text(f"{c}{r}"))
            tds.append(f"<th>{t}</th>" if header else f"<td{_cls(vals[c])}>{t}</td>")
        has_vals = any(v not in (None, "") for v in vals.values())
        lab = str(label or "").lower()
        cls = ("sec" if not has_vals else
               "total" if lab.startswith(("total", "irr", "npv of")) else "")
        rows.append(f"<tr class='{cls}'>{''.join(tds)}</tr>")
    investor = f"<div class='scroll wide'><table class=cf><tbody>{''.join(rows)}</tbody></table></div>"

    yields = "".join(f"<tr><td>{esc(g.text(f'B{r}'))}</td><td class=n>{esc(g.text(f'E{r}'))}</td></tr>"
                     for r in (138, 139))
    return (f"<h2 id=cash-flow>Operating cash flow</h2>{annual}"
            f"<h2>Investor cash flow &amp; returns</h2>{investor}"
            f"<h2>Project yield</h2><table class=narrow>{yields}</table>")


# ---------------------------------------------------------------- page shell
CSS = """
:root{--bg:#f6f5f1;--fg:#1d1d1b;--muted:#6b6b66;--line:#e3e2dc;--card:#fff;--neg:#b42318;--pos:#1a7f4b;--accent:#2f5d50;--accent-fg:#fff;--hl:#fff7d6;--calc:#f0efea}
@media (prefers-color-scheme:dark){:root{--bg:#141513;--fg:#ecebe6;--muted:#9b9a93;--line:#2c2d2a;--card:#1c1d1a;--neg:#f2867b;--pos:#6fcf97;--accent:#8fc1ae;--accent-fg:#10201b;--hl:#3a3420;--calc:#232420}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif}
header.top{position:sticky;top:0;z-index:20;background:var(--card);border-bottom:1px solid var(--line)}
.bar{max-width:1400px;margin:0 auto;padding:10px 16px;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
.bar .title{font-weight:600;flex:1 1 280px;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tabs{display:flex;gap:4px}
.tabs a,.tabs button{font:inherit;padding:6px 12px;border-radius:6px;color:var(--fg);text-decoration:none;border:1px solid transparent;background:none;cursor:pointer}
.tabs .on{background:var(--bg);border-color:var(--line);font-weight:600}
.actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.btn{font:inherit;border-radius:6px;padding:7px 14px;cursor:pointer;border:1px solid var(--line);background:var(--card);color:var(--fg);text-decoration:none;white-space:nowrap}
.btn.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-fg);font-weight:600}
main{max-width:1400px;margin:0 auto;padding:20px 16px 80px}
h2{font-size:17px;margin:32px 0 10px;color:var(--accent)}
h3{font-size:14px;margin:0 0 8px;color:var(--accent)}
.muted{color:var(--muted);font-size:13px}
.kpis{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:10px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px}
.kpi .k{font-size:12px;color:var(--muted)}.kpi .v{font-size:20px;font-weight:600;font-variant-numeric:tabular-nums}
.kpi.neg .v{color:var(--neg)}.kpi.pos .v{color:var(--pos)}
.cards{columns:2 520px;column-gap:16px}
.card{break-inside:avoid;background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px;margin:0 0 16px}
.scroll{overflow-x:auto}
table{width:100%;border-collapse:collapse}
th,td{padding:4px 8px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{font-size:12px;color:var(--muted);font-weight:500}
td.n,.cf th:not(.lab){text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
td.neg{color:var(--neg)}
tr.total td,tr.total th{font-weight:600;color:var(--fg);border-top:1px solid var(--muted)}
.wide{background:var(--card);border:1px solid var(--line);border-radius:8px;max-height:75vh}
.cf{width:auto;min-width:100%}
.cf thead th{position:sticky;top:0;background:var(--card);z-index:2}
.cf .lab{position:sticky;left:0;background:var(--card);z-index:1;text-align:left;color:var(--fg);font-weight:500;min-width:250px;max-width:300px;font-size:13px}
.cf thead .lab{z-index:3}
.cf .lab.ind{padding-left:22px;font-weight:400}
.cf tr.sec th{color:var(--accent);font-weight:600;padding-top:14px}
table.narrow{width:auto;min-width:420px;background:var(--card);border:1px solid var(--line);border-radius:8px}
/* inputs */
.layout{display:grid;grid-template-columns:220px 1fr;gap:24px}
@media (max-width:900px){.layout{grid-template-columns:1fr}nav.toc{display:none}}
nav.toc{position:sticky;top:70px;align-self:start;max-height:calc(100vh - 90px);overflow:auto;font-size:13px}
nav.toc a{display:block;padding:3px 8px;border-radius:4px;color:var(--fg);text-decoration:none}
nav.toc a:hover{background:var(--card)}
nav.toc .cnt{color:var(--muted);float:right}
.inputs td{vertical-align:middle}
.inputs td.v{width:220px}
.inputs .seg{color:var(--muted)}
.inputs input,.inputs select{width:100%;font:inherit;text-align:right;background:var(--hl);color:var(--fg);border:1px solid var(--line);border-radius:4px;padding:4px 8px}
.inputs select{text-align:left}
.inputs input.txt{text-align:left}
.inputs .changed{border-color:var(--accent);box-shadow:0 0 0 2px var(--accent)}
.inputs .calc{display:block;text-align:right;background:var(--calc);border-radius:4px;padding:4px 8px;color:var(--muted)}
.tag{font-size:11px;color:var(--muted);margin-left:6px}
#changed-count{font-size:13px;color:var(--accent)}
.req{color:var(--neg);margin-left:3px;font-weight:600}
.alert{background:var(--card);border:1px solid var(--neg);border-left:4px solid var(--neg);border-radius:8px;padding:10px 14px;margin:0 0 14px}
.alert strong{color:var(--neg)}
.alert ul{margin:6px 0 0;padding-left:20px;columns:2 320px}
.alert a{color:var(--fg)}
.inputs [aria-invalid=true]{border-color:var(--neg);box-shadow:0 0 0 2px var(--neg);background:var(--card)}
.errmsg{color:var(--neg);font-size:12px;margin-top:2px;text-align:right}
button:disabled{opacity:.45;cursor:not-allowed}
"""

JS = r"""
const base=new Map();
document.querySelectorAll('[data-base]').forEach(el=>base.set(el,el.dataset.base));
function mark(){let n=0;base.forEach((b,el)=>{const c=el.value!==b;el.classList.toggle('changed',c);if(c)n++;});
 const s=document.getElementById('changed-count');if(s)s.textContent=n?n+' input'+(n>1?'s':'')+' changed from base case':'';}
document.addEventListener('input',mark);document.addEventListener('change',mark);mark();
const NUM=/^\(?-?\$?-?[\d,]*\.?\d+(e-?\d+)?%?\)?$/i;
function check(){const bad=[];
 document.querySelectorAll('[data-label]').forEach(el=>{
  const v=el.value.trim();let m=null;
  if(el.dataset.req&&v==='')m='required';
  else if(el.dataset.num&&v!==''&&!NUM.test(v.replace(/\s/g,'')))m='must be a number';
  else if(el.dataset.srv)m=el.dataset.srv;
  if(m){el.setAttribute('aria-invalid','true');bad.push([el,m]);}else el.removeAttribute('aria-invalid');});
 document.querySelectorAll('.needs-valid').forEach(b=>{b.disabled=bad.length>0;
  b.title=bad.length?'Fix the inputs listed at the top first':'';});
 const srv=document.getElementById('srv-alert');if(srv)srv.hidden=true;  // the live list below replaces it
 const box=document.getElementById('client-alert');if(!box)return;
 if(!bad.length){box.hidden=true;box.innerHTML='';return;}
 const e=s=>s.replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
 box.hidden=false;
 box.innerHTML='<strong>'+bad.length+' input'+(bad.length>1?'s need':' needs')+
  ' attention before outputs can be calculated:</strong><ul>'+
  bad.map(([el,m])=>'<li><a href="#'+el.id+'">'+e(el.dataset.label)+'</a> — '+e(m)+'</li>').join('')+'</ul>';}
document.querySelectorAll('.errmsg').forEach(d=>{const el=d.parentNode.querySelector('[data-label]');if(el)el.dataset.srv=d.textContent;});
document.addEventListener('input',ev=>{const t=ev.target;if(t.dataset&&t.dataset.srv){delete t.dataset.srv;
 const d=t.parentNode.querySelector('.errmsg');if(d)d.remove();}check();});
document.addEventListener('change',check);check();
document.addEventListener('submit',ev=>{const b=ev.submitter;
 if(b&&b.value!=='inputs'&&document.querySelector('.needs-valid:disabled')){ev.preventDefault();window.scrollTo(0,0);return;}
 if(b&&b.value==='outputs'&&!b.hidden){b.textContent='Calculating…';}});
"""


def shell(title: str, active: str, body: str, form_open: str, actions: str, script=False) -> str:
    tabs = (f"<button class='{'on' if active == 'inputs' else ''}' name=_view value=inputs "
            f"formaction='/inputs'>1 · Inputs</button>"
            f"<button class='{'on' if active == 'outputs' else ''} needs-valid' name=_view value=outputs "
            f"formaction='/outputs'>2 · Outputs</button>")
    return f"""<!doctype html>
<html lang=en><head><meta charset=utf-8>
<meta name=viewport content="width=device-width, initial-scale=1">
<title>{esc(title)}</title><style>{CSS}</style></head><body>
{form_open}{DEFAULT_BTN}
<header class=top><div class=bar>
<div class=title>{esc(title)}</div>
<div class=tabs>{tabs}</div>
<div class=actions>{actions}</div>
</div></header>
<main>{body}</main>
</form>
{'<script>' + JS + '</script>' if script else ''}
</body></html>"""


EXCEL_BTN = ("<button class='btn needs-valid' name=_view value=excel formaction='/excel'>"
             "Download Excel</button>")
# First submit button in the form, so Enter in a field calculates outputs.
DEFAULT_BTN = ("<button name=_view value=outputs formaction='/outputs' hidden "
               "class=needs-valid tabindex=-1 aria-hidden=true></button>")


# ---------------------------------------------------------------- inputs page
def input_rows(inputs: Inputs, deal="site8"):
    lay = layout(deal)["inputs"]
    for r in inputs.rows:
        meta = lay.get(str(r["row"]), {"fmt": "General"})
        yield r, meta, inputs[r["key"]], is_calculated(r)


def required(base_value) -> bool:
    """An input is required when the base case has a value for it."""
    return base_value not in (None, "")


def render_inputs(inputs: Inputs, base: Inputs, res: ModelResult | None = None,
                  changed: set[str] | None = None, problems: dict[str, str] | None = None,
                  error: str | None = None) -> str:
    changed = changed or set()
    problems = problems or {}
    sections: dict[str, list[str]] = {}
    for r, meta, v, calc in input_rows(inputs):
        key, f = r["key"], meta["fmt"]
        bv = base[key]
        label = r["label"] + (f" ({r['segment']})" if r.get("segment") else "")
        attrs = (f" id='f-{esc(key)}' data-label='{esc(label)}'"
                 + (" data-req=1" if required(bv) else "")
                 + (" data-num=1" if isinstance(bv, (int, float)) and not isinstance(bv, bool)
                    and not meta.get("options") else "")
                 + (" aria-invalid=true" if key in problems else ""))
        if calc:
            shown = res.get("Master Inputs", f"E{r['row']}") if res else v
            cell = f"<span class=calc title='Calculated in the model'>{esc(fmt(shown, f))}</span>"
        elif isinstance(bv, dt.datetime) or isinstance(v, dt.datetime):
            val = v.date().isoformat() if isinstance(v, dt.datetime) else ""
            bval = bv.date().isoformat() if isinstance(bv, dt.datetime) else ""
            cell = (f"<input type=date name='{esc(key)}' value='{val}' data-base='{bval}'{attrs}>")
        elif meta.get("options"):
            opts = list(meta["options"])
            if v not in opts and v not in (None, ""):
                opts.insert(0, v)
            if v in (None, ""):
                opts.insert(0, "")
            o = "".join(f"<option value='{esc(input_text(x, f))}'"
                        f"{' selected' if x == v else ''}>"
                        f"{esc(fmt(x, f))}</option>" for x in opts)
            cell = (f"<select name='{esc(key)}' data-base='{esc(input_text(bv, f))}'{attrs}>{o}</select>")
        else:
            txt = isinstance(bv, str) and not re.fullmatch(r"[-$\d.,%]+", bv or "")
            cell = (f"<input name='{esc(key)}' value='{esc(input_text(v, f))}' "
                    f"data-base='{esc(input_text(bv, f))}'{' class=txt' if txt else ''}{attrs}>")
        seg = f"<span class=seg>{esc(r['segment'])}</span>" if r.get("segment") else ""
        req = "<span class=req title='Required'>*</span>" if required(bv) and not calc else ""
        msg = f"<div class=errmsg>{esc(problems[key])}</div>" if key in problems else ""
        sections.setdefault(r["section"], []).append(
            f"<tr><td>{esc(r['label'])}{req}</td><td>{seg}</td><td class=v>{cell}{msg}</td></tr>")

    def anchor(s):
        return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")

    toc = "".join(f"<a href='#{anchor(s)}'>{esc(s)}<span class=cnt>{len(rows)}</span></a>"
                  for s, rows in sections.items())
    cards = "".join(
        f"<section class=card id='{anchor(s)}'><h3>{esc(s)}</h3>"
        f"<table class=inputs><thead><tr><th>Input</th><th>Detail</th><th>Value</th></tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table></section>"
        for s, rows in sections.items())
    banner = ""
    if error or problems:
        labels = {r["key"]: r["label"] + (f" ({r['segment']})" if r.get("segment") else "")
                  for r in inputs.rows}
        items = "".join(f"<li><a href='#f-{esc(k)}'>{esc(labels.get(k, k))}</a> — {esc(m)}</li>"
                        for k, m in problems.items())
        banner = (f"<div class=alert role=alert{' id=srv-alert' if problems else ''}><strong>{esc(error or 'Outputs could not be calculated.')}"
                  f"</strong>{f'<ul>{items}</ul>' if items else ''}</div>")
    body = (f"{banner}<div id=client-alert class=alert role=alert hidden></div>"
            f"<p class=muted>Every input from the Master Inputs tab. Yellow fields are editable; "
            f"grey values are calculated; <span class=req>*</span> marks inputs the model needs. "
            f"Percent fields accept 5 or 5%. <span id=changed-count></span></p>"
            f"<div class=layout><nav class=toc>{toc}</nav><div>{cards}</div></div>")
    actions = (f"<a class=btn href='/'>Reset to base case</a>{EXCEL_BTN}"
               f"<button class='btn primary needs-valid' name=_view value=outputs "
               f"formaction='/outputs'>Calculate outputs →</button>")
    return shell(f"{inputs['project_name']} · Inputs", "inputs", body,
                 "<form method=post action='/outputs'>", actions, script=True)


# ---------------------------------------------------------------- outputs page
def hidden_inputs(inputs: Inputs) -> str:
    out = []
    for r, meta, v, calc in input_rows(inputs):
        if calc:
            continue
        if isinstance(v, dt.datetime):
            t = v.date().isoformat()
        else:
            t = input_text(v, meta["fmt"])
        out.append(f"<input type=hidden name='{esc(r['key'])}' value='{esc(t)}'>")
    return "".join(out)


def render_outputs(res: ModelResult, snap: Snapshot | None = None,
                   changed: set[str] | None = None) -> str:
    changed = changed or set()
    g = res.get
    gap = g(SU, "E58")
    units = g("Program", "E64")
    tdc = g(SU, "E26")
    irr = g(OCF, "F112")
    kpis = [
        ("Surplus / (gap)", fmt(gap, "\\$#,##0;\"($\"#,##0\\)"), "neg" if isinstance(gap, (int, float)) and gap < 0 else "pos"),
        ("Total development cost", fmt(tdc, "\\$#,##0"), ""),
        ("Units", fmt(units, "#,##0"), ""),
        ("TDC per unit", fmt(g(SU, "G26"), "\\$#,##0"), ""),
        ("Tax credit equity", fmt(g(SU, "E46"), "\\$#,##0"), ""),
        ("Permanent loan", fmt(g(SU, "E47"), "\\$#,##0"), ""),
        ("Year 1 NOI", fmt(g(OCF, "J47"), "\\$#,##0;\"($\"#,##0\\)"), ""),
        ("Investor IRR (15-yr)", fmt(irr, "0.0%") if isinstance(irr, (int, float)) else "n/a", ""),
    ]
    kpi_html = "".join(f"<div class='kpi {c}'><div class=k>{esc(k)}</div><div class=v>{esc(v)}</div></div>"
                       for k, v, c in kpis)

    if changed:
        labels = [r["label"] + (f" ({r['segment']})" if r.get("segment") else "")
                  for r in res.inputs.rows if r["key"] in changed]
        scen = (f"<p class=muted>Scenario: {len(changed)} input{'s' if len(changed) > 1 else ''} "
                f"changed from the base case: {esc('; '.join(labels[:12]))}"
                f"{' …' if len(labels) > 12 else ''}.</p>")
    else:
        scen = "<p class=muted>Base case (inputs as in the Excel model).</p>"

    parity = ""
    if snap is not None:
        checks = [(s, a) for s, cells in res.sheets.items() for a in cells]
        diff = [(s, a) for s, a in checks if not same(res.get(s, a), snap.get(s, a), rel=1e-5, abs_=1e-3)]
        parity = (f"<p class=muted>Model check: Python matches Excel on {len(checks) - len(diff):,} of "
                  f"{len(checks):,} calculated cells across {len(res.sheets)} tabs.</p>")
    elif not res.converged:
        parity = "<p class=muted>Warning: the model did not fully converge for these inputs.</p>"

    body = (f"{scen}<div class=kpis>{kpi_html}</div>"
            f"<p class=muted><a href='#summary'>Summary</a> · <a href='#cash-flow'>Cash flow</a></p>"
            f"<h2 id=summary>Summary</h2><div class=cards>{summary_sections(res)}</div>"
            f"{cash_flow(res)}{parity}")
    actions = (f"<button class=btn formaction='/inputs'>← Edit inputs</button>{EXCEL_BTN}")
    return shell(f"{res.inputs['project_name']} · Outputs", "outputs", body,
                 f"<form method=post action='/outputs'>{hidden_inputs(res.inputs)}", actions)


def main(argv):
    deal = "site8"
    snap = Snapshot.load(deal)
    res = run_model(Inputs.load(deal), deal, fallback=snap)
    out = ROOT / "out"
    out.mkdir(exist_ok=True)
    path = out / f"{deal}.html"
    path.write_text(render_outputs(res, snap), encoding="utf-8")
    print(f"wrote {path}  (iterations={res.iterations}, converged={res.converged})")
    if "--open" in argv:
        webbrowser.open(path.as_uri())


if __name__ == "__main__":
    main(sys.argv[1:])
