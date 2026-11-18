"""Web app: an Inputs page, an Outputs page, and an Excel download.

Stateless: every form post carries the full set of inputs, so the same handler
runs locally and as a Vercel Python function (api/index.py).

    GET  /  or /inputs     inputs page, base case
    POST /inputs           inputs page with the posted values
    GET  /outputs          outputs (summary + cash flows), base case
    POST /outputs          outputs for the posted inputs
    GET  /excel            the Excel model, base case
    POST /excel            the Excel model with the posted inputs written in

    python -m uw.serve            # http://localhost:8765
"""
from __future__ import annotations

import datetime as dt
import re
import sys
import urllib.parse
import webbrowser
from functools import lru_cache
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from uw.engine import Inputs, Snapshot, run_model, same
from uw.excel_export import build_xlsx, is_calculated
from uw.report import layout, render_inputs, render_outputs, required

DEAL = "site8"
BASE = Inputs.load(DEAL)
SNAP = Snapshot.load(DEAL)
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def coerce(text: str, base, numfmt: str = "General"):
    """Parse a form value back to the type of the base-case input."""
    t = text.strip()
    if isinstance(base, dt.datetime):
        if not t:
            return None
        d = dt.date.fromisoformat(t)
        return base if d == base.date() else dt.datetime(d.year, d.month, d.day)
    if isinstance(base, bool):
        return t.lower() in ("true", "yes", "1")
    if isinstance(base, (int, float)) or base is None:
        if t == "":
            return None
        n = t.replace(",", "").replace("$", "").replace(" ", "")
        neg = n.startswith("(") and n.endswith(")")
        n = n.strip("()")
        if not re.fullmatch(r"-?\d*\.?\d+(e-?\d+)?%?", n, re.I):
            if base is None:
                return text
            raise ValueError(f"{text!r} is not a number")
        # Like Excel, typing 5 into a percent-formatted cell means 5%.
        pct = n.endswith("%") or ("%" in numfmt and isinstance(base, (int, float)))
        f = float(n.rstrip("%")) / (100 if pct else 1)
        f = -f if neg else f
        if isinstance(base, (int, float)) and same(f, base, rel=1e-9, abs_=1e-12):
            return base  # unchanged; only the display formatting round-tripped
        return int(f) if isinstance(base, int) and f.is_integer() else f
    return text


def inputs_from_form(form: dict[str, list[str]]) -> tuple[Inputs, dict[str, str]]:
    """Inputs from a posted form, plus {key: problem} for inputs that are missing or invalid.

    A field missing from the post keeps its base-case value; a field posted blank
    is "missing" when the base case has a value for it (the model needs it).
    """
    new = BASE.copy()
    fmts = layout(DEAL)["inputs"]
    problems: dict[str, str] = {}
    for r in BASE.rows:
        key = r["key"]
        if key not in form or is_calculated(r):
            continue
        text = form[key][0]
        if text.strip() == "" and required(BASE[key]):
            new[key] = None
            problems[key] = "required"
            continue
        try:
            new[key] = coerce(text, BASE[key], fmts.get(str(r["row"]), {}).get("fmt", "General"))
        except ValueError:
            new[key] = text  # keep what was typed so it can be corrected
            problems[key] = "must be a number" if isinstance(BASE[key], (int, float)) else "not valid"
    return new, problems


def changed_keys(inputs: Inputs) -> set[str]:
    return {k for k, v in inputs.values.items() if v != BASE.values[k]}


@lru_cache(maxsize=1)
def base_run():
    return run_model(BASE, DEAL, fallback=SNAP)


def run(inputs: Inputs):
    changed = changed_keys(inputs)
    return (base_run() if not changed else run_model(inputs, DEAL, fallback=SNAP)), changed


def inputs_page(inputs: Inputs, problems: dict[str, str] | None = None,
                error: str | None = None) -> bytes:
    if problems or error:  # don't run the model on inputs it can't use
        return render_inputs(inputs, BASE, None, changed_keys(inputs), problems, error).encode()
    res, changed = run(inputs)
    return render_inputs(inputs, BASE, res, changed).encode()


def outputs_page(inputs: Inputs) -> bytes:
    res, changed = run(inputs)
    # The Excel cross-check only means something for the unchanged base case.
    return render_outputs(res, None if changed else SNAP, changed).encode()


def apply_lookup(inputs: Inputs, lk) -> tuple[Inputs, dict]:
    """Write lookup findings into a copy of the inputs; returns (inputs, {key: Found} applied).

    Dropdown inputs only take values in their list; anything else is reported
    as not filled so the model never gets a value its lookups can't handle.
    """
    new = inputs.copy()
    fmts = layout(DEAL)["inputs"]
    applied = {}
    for key, f in lk.found.items():
        if key not in BASE.values:
            continue
        row = next(r for r in BASE.rows if r["key"] == key)
        if is_calculated(row):
            continue
        opts = fmts.get(str(row["row"]), {}).get("options")
        v = f.value
        if opts is not None:
            match = next((o for o in opts if str(o).strip().lower() == str(v).strip().lower()), None)
            if match is None:
                lk.missing.append(f"{row['label']}: found \"{v}\" ({f.source}), which is not in the "
                                  f"model's list; pick the closest option manually")
                continue
            v = match
        base = BASE[key]
        if isinstance(base, str) and not isinstance(v, str):
            v = str(v)
        elif isinstance(base, (int, float)) and isinstance(v, str):
            try:
                v = float(v.replace(",", ""))
            except ValueError:
                continue
        new[key] = v
        applied[key] = f
    return new, applied


def lookup_page(inputs: Inputs, problems: dict[str, str], address: str) -> bytes:
    from uw.lookup import lookup
    lk = lookup(address)
    new, applied = apply_lookup(inputs, lk)
    stale = {}
    if lk.matched:
        from uw.lookup import SITE_KEYS
        stale = {k: "not found by the lookup; still the previous site's value"
                 for k in SITE_KEYS if k in BASE.values and k not in applied
                 and new[k] not in (None, "")}
    problems = {k: m for k, m in problems.items() if k not in applied}
    if problems:
        return render_inputs(new, BASE, None, changed_keys(new), problems, None,
                             lookup=lk, filled=applied, previous=inputs, stale=stale).encode()
    try:
        res, changed = run(new)
    except Exception:  # noqa: BLE001 - still show what was found
        res, changed = None, changed_keys(new)
    return render_inputs(new, BASE, res, changed, lookup=lk, filled=applied,
                         previous=inputs, stale=stale).encode()


def excel_file(inputs: Inputs) -> tuple[bytes, str]:
    res, changed = run(inputs)
    name = f"Site 8 - 21101 Victory Blvd - Underwriting{' (scenario)' if changed else ''}.xlsx"
    return build_xlsx(res, BASE, changed), name


class Handler(BaseHTTPRequestHandler):
    def _send(self, status: int, body: bytes, ctype="text/html; charset=utf-8", filename=None):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if filename:
            q = urllib.parse.quote(filename)
            self.send_header("Content-Disposition",
                             f"attachment; filename=\"{filename}\"; filename*=UTF-8''{q}")
        self.end_headers()
        self.wfile.write(body)

    def _view(self, form: dict) -> str:
        """Which page: the button's _view field, else ?view=, else the URL path.

        On Vercel every path is rewritten to /api/index before it reaches this
        handler, so the path alone can't be relied on.
        """
        url = urllib.parse.urlparse(self.path)
        v = (form.get("_view") or urllib.parse.parse_qs(url.query).get("view") or [""])[0]
        if v in ("inputs", "outputs", "excel", "lookup"):
            return v
        for header in ("x-vercel-original-path", "x-forwarded-uri", "x-original-url"):
            if self.headers.get(header):
                url = urllib.parse.urlparse(self.headers[header])
                break
        path = url.path.rstrip("/")
        for v in ("outputs", "excel", "favicon.ico"):
            if path.endswith("/" + v):
                return v
        return "inputs"

    def _route(self, view: str, inputs: Inputs, problems: dict[str, str] | None = None,
               form: dict | None = None):
        if view == "favicon.ico":
            self._send(204, b"", "text/plain")
            return
        if view == "lookup":
            address = ((form or {}).get("lookup_address") or [""])[0]
            self._send(200, lookup_page(inputs, problems or {}, address))
            return
        if view != "inputs" and problems:
            n = len(problems)
            self._send(200, inputs_page(inputs, problems,
                       f"Outputs can't be calculated yet: {n} input{'s need' if n > 1 else ' needs'} "
                       f"attention."))
            return
        try:
            if view == "outputs":
                self._send(200, outputs_page(inputs))
            elif view == "excel":
                body, name = excel_file(inputs)
                self._send(200, body, XLSX, name)
            else:
                self._send(200, inputs_page(inputs, problems))
        except Exception as e:  # noqa: BLE001 - tell the user, keep their inputs
            self._send(200, inputs_page(inputs, None,
                       f"The model could not calculate with these inputs ({type(e).__name__}: {e}). "
                       f"Check the inputs you changed."))

    def do_GET(self):
        self._route(self._view({}), BASE)

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        form = urllib.parse.parse_qs(self.rfile.read(length).decode(), keep_blank_values=True)
        inputs, problems = inputs_from_form(form)
        self._route(self._view(form), inputs, problems, form)

    def log_message(self, *a):
        pass


def main(argv):
    port = int(argv[0]) if argv else 8765
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://localhost:{port}/"
    print(f"Underwriting model running at {url}  (Ctrl+C to stop)")
    webbrowser.open(url)
    srv.serve_forever()


if __name__ == "__main__":
    main(sys.argv[1:])
