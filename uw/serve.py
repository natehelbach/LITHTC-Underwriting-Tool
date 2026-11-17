"""Web app: edit inputs in the browser, recalculate, see results.

Stateless: every POST carries the full set of inputs from the form, so the same
handler runs locally and as a Vercel Python function (api/index.py).

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


def inputs_from_form(form: dict[str, list[str]]) -> Inputs:
    new = BASE.copy()
    for key, (val, *_) in form.items():
        if key in new.values and not isinstance(BASE.values[key], dt.datetime):
            new[key] = coerce(val, BASE.values[key])
    return new


def page(inputs: Inputs) -> bytes:
    res = run_model(inputs, DEAL, fallback=SNAP)
    changed = {k for k, v in inputs.values.items() if v != BASE.values[k]}
    # The Excel cross-check only means something for the unchanged base case.
    return render(res, SNAP if not changed else None, editable=True, changed=changed).encode()


class Handler(BaseHTTPRequestHandler):
    def _send(self, status: int, body: bytes, ctype="text/html; charset=utf-8"):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if urllib.parse.urlparse(self.path).path != "/":
            self._send(404, b"Not found", "text/plain")
            return
        self._send(200, page(BASE))

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        form = urllib.parse.parse_qs(self.rfile.read(length).decode(), keep_blank_values=True)
        try:
            self._send(200, page(inputs_from_form(form)))
        except Exception as e:  # noqa: BLE001 - show the user what failed
            self._send(400, f"Could not calculate with these inputs: {e}".encode(), "text/plain")

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
