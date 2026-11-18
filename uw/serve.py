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
from uw.report import layout, render_inputs, render_outputs

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


def inputs_from_form(form: dict[str, list[str]]) -> Inputs:
    new = BASE.copy()
    fmts = layout(DEAL)["inputs"]
    errors = []
    for r in BASE.rows:
        key = r["key"]
        if key not in form or is_calculated(r):
            continue
        try:
            new[key] = coerce(form[key][0], BASE[key],
                              fmts.get(str(r["row"]), {}).get("fmt", "General"))
        except ValueError as e:
            errors.append(f"{r['label']}: {e}")
    if errors:
        raise ValueError("; ".join(errors))
    return new


def changed_keys(inputs: Inputs) -> set[str]:
    return {k for k, v in inputs.values.items() if v != BASE.values[k]}


@lru_cache(maxsize=1)
def base_run():
    return run_model(BASE, DEAL, fallback=SNAP)


def run(inputs: Inputs):
    changed = changed_keys(inputs)
    return (base_run() if not changed else run_model(inputs, DEAL, fallback=SNAP)), changed


def inputs_page(inputs: Inputs) -> bytes:
    res, changed = run(inputs)
    return render_inputs(inputs, BASE, res, changed).encode()


def outputs_page(inputs: Inputs) -> bytes:
    res, changed = run(inputs)
    # The Excel cross-check only means something for the unchanged base case.
    return render_outputs(res, None if changed else SNAP, changed).encode()


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

    def _route(self, inputs: Inputs):
        # Vercel rewrites every path to api/index; unknown paths get the inputs page.
        path = urllib.parse.urlparse(self.path).path.rstrip("/")
        if path.endswith("/favicon.ico"):
            self._send(204, b"", "text/plain")
        elif path.endswith("/outputs"):
            self._send(200, outputs_page(inputs))
        elif path.endswith("/excel"):
            body, name = excel_file(inputs)
            self._send(200, body, XLSX, name)
        else:
            self._send(200, inputs_page(inputs))

    def do_GET(self):
        self._route(BASE)

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        form = urllib.parse.parse_qs(self.rfile.read(length).decode(), keep_blank_values=True)
        try:
            inputs = inputs_from_form(form)
        except ValueError as e:
            self._send(400, f"Could not read these inputs: {e}".encode(), "text/plain; charset=utf-8")
            return
        try:
            self._route(inputs)
        except Exception as e:  # noqa: BLE001 - show the user what failed
            self._send(500, f"Could not calculate with these inputs: {e}".encode(),
                       "text/plain; charset=utf-8")

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
