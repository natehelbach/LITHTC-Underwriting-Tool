"""Local web app: edit inputs in the browser, recalculate, see results.

    python -m uw.serve            # http://localhost:8765
"""
from __future__ import annotations

import datetime as dt
import sys
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from uw.engine import Inputs, Snapshot, run_model
from uw.report import render

DEAL = "site8"
BASE = Inputs.load(DEAL)
SNAP = Snapshot.load(DEAL)
state = {"inputs": BASE.copy()}


def coerce(text: str, base):
    """Parse a form value back to the type of the base-case input."""
    t = text.strip().replace(",", "").replace("$", "")
    if isinstance(base, bool):
        return t.lower() in ("true", "yes", "1")
    if isinstance(base, (int, float)) or base is None:
        if t == "":
            return None
        try:
            f = float(t.rstrip("%")) / (100 if t.endswith("%") else 1)
            return int(f) if isinstance(base, int) and f.is_integer() else f
        except ValueError:
            return text
    return text


def changed_keys(inputs: Inputs) -> set[str]:
    return {k for k, v in inputs.values.items() if v != BASE.values[k]}


class Handler(BaseHTTPRequestHandler):
    def _page(self):
        inputs = state["inputs"]
        res = run_model(inputs, DEAL, fallback=SNAP)
        changed = changed_keys(inputs)
        body = render(res, SNAP if not changed else None, editable=True, changed=changed).encode()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        if self.path.startswith("/favicon"):
            self.send_response(404)
            self.end_headers()
            return
        if "reset" in q:
            state["inputs"] = BASE.copy()
        self._page()

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        form = urllib.parse.parse_qs(self.rfile.read(length).decode(), keep_blank_values=True)
        new = state["inputs"].copy()
        for key, (val, *_) in form.items():
            if key in new.values and not isinstance(BASE.values[key], dt.datetime):
                new[key] = coerce(val, BASE.values[key])
        state["inputs"] = new
        self._page()

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
