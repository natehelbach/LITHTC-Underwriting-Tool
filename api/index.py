"""Vercel entrypoint: every route is rewritten here (see vercel.json)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from uw.serve import Handler  # noqa: E402


class handler(Handler):  # Vercel's Python runtime looks for a class named `handler`
    pass
