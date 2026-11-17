"""Calculation engine.

The model is a set of *sheet modules*, one per Excel tab (uw/sheets/*.py). Each
module subclasses `Sheet` and implements `compute(ctx) -> dict[cell, value]`,
returning the cells of its tab that it calculates (keyed by Excel address, e.g.
"H37"). Inside a module the logic is plain, readable Python; only the boundary
between tabs uses Excel addresses, which is what lets every tab be verified
against the Excel model in isolation and keeps cross-tab wiring mechanical.

Reading values:
    ctx.mi["credit_type"]           a Master Inputs value by key (data/<deal>/inputs.json)
    ctx.ref("Dev Budget", "H37")    one cell of another tab
    ctx.col("Rent Builder", "F12:F39")   1-D list from a single row/column range
    ctx.grid("Unit Takeoff", "B6:S94")   2-D list (rows of columns)

Two run modes:
  * isolated (verify):  other tabs are served from the Excel snapshot, so a module
                        is checked against Excel on its own (python -m uw.verify).
  * model:              other tabs are served from the other modules' latest
                        results; the whole model is iterated to a fixed point,
                        the same way Excel's iterative calculation resolves the
                        Dev Budget <-> Flow of Funds <-> Financing circularity.
"""
from __future__ import annotations

import datetime as dt
import gzip
import importlib
import json
import math
import pkgutil
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"


# ---------------------------------------------------------------- inputs
def _from_json(v):
    if isinstance(v, dict) and "__date__" in v:
        return dt.datetime.fromisoformat(v["__date__"])
    return v


class Inputs:
    """Master Inputs, keyed by readable names. Mutable, so scenarios are just copies with edits."""

    def __init__(self, rows: list[dict]):
        self.rows = rows
        self._by_key = {r["key"]: r for r in rows}
        self._by_row = {r["row"]: r for r in rows}
        self.values: dict[str, Any] = {r["key"]: _from_json(r["value"]) for r in rows}

    @classmethod
    def load(cls, deal: str = "site8") -> "Inputs":
        d = json.loads((DATA / deal / "inputs.json").read_text(encoding="utf-8"))
        return cls(d["inputs"])

    def __getitem__(self, key: str):
        if key not in self.values:
            raise KeyError(f"Unknown input {key!r}; see data/<deal>/inputs.json")
        return self.values[key]

    def __setitem__(self, key: str, value):
        if key not in self.values:
            raise KeyError(f"Unknown input {key!r}")
        self.values[key] = value

    def key_for_row(self, row: int) -> str:
        return self._by_row[row]["key"]

    def by_row(self, row: int):
        return self.values[self._by_row[row]["key"]]

    def copy(self, **overrides) -> "Inputs":
        new = Inputs(self.rows)
        new.values = dict(self.values)
        for k, v in overrides.items():
            new[k] = v
        return new

    def meta(self, key: str) -> dict:
        return self._by_key[key]


# ---------------------------------------------------------------- excel snapshot (oracle)
class Snapshot:
    """Cached values of every cell in the source workbook."""

    def __init__(self, data: dict[str, dict[str, Any]]):
        self.data = data

    @classmethod
    def load(cls, deal: str = "site8") -> "Snapshot":
        with gzip.open(DATA / deal / "excel_snapshot.json.gz", "rt", encoding="utf-8") as f:
            raw = json.load(f)
        return cls({s: {a: _from_json(v) for a, v in cells.items()} for s, cells in raw.items()})

    def get(self, sheet: str, addr: str):
        return self.data.get(sheet, {}).get(addr.replace("$", ""))


# ---------------------------------------------------------------- sheet modules
class Sheet:
    """Base class for one Excel tab."""
    name: str = ""          # exact Excel tab name

    def compute(self, ctx: "Ctx") -> dict[str, Any]:
        raise NotImplementedError


_REGISTRY: dict[str, type[Sheet]] = {}


def register(cls: type[Sheet]) -> type[Sheet]:
    _REGISTRY[cls.name] = cls
    return cls


def registry() -> dict[str, type[Sheet]]:
    import uw.sheets as pkg
    for m in pkgutil.iter_modules(pkg.__path__):
        importlib.import_module(f"uw.sheets.{m.name}")
    return dict(_REGISTRY)


# ---------------------------------------------------------------- context
def _cells(rng: str):
    rng = rng.replace("$", "")
    c1, r1, c2, r2 = range_boundaries(rng if ":" in rng else f"{rng}:{rng}")
    return c1, r1, c2, r2


class Ctx:
    def __init__(self, inputs: Inputs, resolver: Callable[[str, str], Any]):
        self.mi = inputs
        self._resolve = resolver
        self.reads: set[tuple[str, str]] = set()

    def ref(self, sheet: str, addr: str):
        addr = addr.replace("$", "")
        self.reads.add((sheet, addr))
        return self._resolve(sheet, addr)

    def grid(self, sheet: str, rng: str) -> list[list]:
        c1, r1, c2, r2 = _cells(rng)
        return [[self.ref(sheet, f"{get_column_letter(c)}{r}") for c in range(c1, c2 + 1)]
                for r in range(r1, r2 + 1)]

    def col(self, sheet: str, rng: str) -> list:
        g = self.grid(sheet, rng)
        if len(g) == 1:
            return g[0]
        if all(len(r) == 1 for r in g):
            return [r[0] for r in g]
        raise ValueError(f"{rng} is 2-D; use grid()")


# ---------------------------------------------------------------- comparison
def same(a, b, rel=1e-6, abs_=1e-6) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    blank = (None, "")
    if a in blank and b in blank:
        return True
    if (a in blank and b == 0) or (b in blank and a == 0):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        if math.isnan(a) or math.isnan(b):
            return False
        return math.isclose(a, b, rel_tol=rel, abs_tol=abs_)
    if isinstance(a, dt.datetime) and isinstance(b, dt.datetime):
        return abs((a - b).total_seconds()) < 1
    if isinstance(a, str) and isinstance(b, str):
        return a.strip() == b.strip()
    return a == b


@dataclass
class VerifyReport:
    sheet: str
    checked: int = 0
    mismatches: list[tuple[str, Any, Any]] = field(default_factory=list)
    missing_exports: list[str] = field(default_factory=list)
    error: str | None = None

    @property
    def ok(self) -> bool:
        return not self.mismatches and not self.missing_exports and self.error is None

    def summary(self, limit=40) -> str:
        lines = [f"[{self.sheet}] checked {self.checked} cells, "
                 f"{len(self.mismatches)} mismatches, {len(self.missing_exports)} exported cells not computed"]
        if self.error:
            lines.append("ERROR: " + self.error)
        for addr, got, exp in self.mismatches[:limit]:
            lines.append(f"  {addr}: python={got!r}  excel={exp!r}")
        if self.missing_exports:
            lines.append("  not computed (read by other tabs): " + " ".join(self.missing_exports[:60]))
        return "\n".join(lines)


def interfaces() -> dict[str, dict[str, list[str]]]:
    return json.loads((ROOT / "reference" / "interfaces.json").read_text(encoding="utf-8"))


def _addr_key(a: str):
    m = re.fullmatch(r"([A-Z]+)(\d+)", a)
    return (int(m.group(2)), column_index_from_string(m.group(1))) if m else (0, 0)


def verify_sheet(name: str, deal: str = "site8", snap: Snapshot | None = None,
                 rel=1e-6, abs_=1e-6, core_only: bool = True) -> VerifyReport:
    """Run one module with every other tab served from Excel, and compare its cells to Excel.

    Exported cells (read by other tabs) that the module does not return are reported,
    so cross-tab wiring can't silently break. With core_only, only cells read by
    tabs in the core calculation chain (ORDER) count as required exports.
    """
    snap = snap or Snapshot.load(deal)
    reg = registry()
    rep = VerifyReport(name)
    inputs = Inputs.load(deal)
    ctx = Ctx(inputs, lambda s, a: snap.get(s, a))
    try:
        out = reg[name]().compute(ctx)
    except Exception as e:  # noqa: BLE001
        import traceback
        rep.error = "".join(traceback.format_exception(e))[-3000:]
        return rep
    for addr in sorted(out, key=_addr_key):
        exp = snap.get(name, addr)
        rep.checked += 1
        if not same(out[addr], exp, rel, abs_):
            rep.mismatches.append((addr, out[addr], exp))
    exports = interfaces().get(name, {})
    for addr, readers in exports.items():
        if core_only and not any(r in ORDER for r in readers):
            continue
        if addr not in out and snap.get(name, addr) is not None:
            rep.missing_exports.append(addr)
    rep.missing_exports.sort(key=_addr_key)
    return rep


# ---------------------------------------------------------------- full model
@dataclass
class ModelResult:
    inputs: Inputs
    sheets: dict[str, dict[str, Any]]
    iterations: int
    converged: bool
    fallback_reads: set[tuple[str, str]]

    def get(self, sheet: str, addr: str):
        return self.sheets.get(sheet, {}).get(addr.replace("$", ""))


ORDER = [  # calculation order; the loop iterates until values stop changing
    "Master Inputs", "Sensitivities", "Initial Screen", "Program", "Lists",
    "LIHTC-TBL Tables", "Utility Allowances", "Federal Minimum Opex",
    "Location Premiums", "Amenity & Property Premiums",
    "Studio Comps", "1BR Comps", "2BR Comps", "Market Rent Survey", "Comps Summary",
    "Unit Takeoff", "Unit Mix Summary", "Rent Builder",
    "OpEx Builder", "OpEx & Other Income", "Leasing",
    "Dev Budget", "LIHTC & Dev Fee", "Financing & Debt Sizing", "Flow of Funds",
    "Operating Cash Flow", "Sources & Uses", "Returns & Summary",
]


def run_model(inputs: Inputs, deal: str = "site8", max_iter: int = 200, tol: float = 0.005,
              fallback: Snapshot | None = None) -> ModelResult:
    """Iterate every ported tab to a fixed point.

    Tabs without a module yet are read from `fallback` (the Excel snapshot) and
    recorded in `fallback_reads`, so the report can show what is still Excel-sourced.
    """
    reg = registry()
    order = [s for s in ORDER if s in reg] + sorted(s for s in reg if s not in ORDER)
    results: dict[str, dict[str, Any]] = {s: {} for s in order}
    fallback_reads: set[tuple[str, str]] = set()

    def resolve(sheet, addr):
        if sheet in results:
            return results[sheet].get(addr)
        fallback_reads.add((sheet, addr))
        return fallback.get(sheet, addr) if fallback else None

    converged, it = False, 0
    for it in range(1, max_iter + 1):
        delta = 0.0
        for s in order:
            new = reg[s]().compute(Ctx(inputs, resolve))
            old = results[s]
            for a, v in new.items():
                o = old.get(a)
                if isinstance(v, (int, float)) and isinstance(o, (int, float)) and not isinstance(v, bool):
                    delta = max(delta, abs(v - o))
                elif v != o:
                    delta = math.inf
            results[s] = new
        if delta <= tol:
            converged = True
            break
    return ModelResult(inputs, results, it, converged, fallback_reads)
