"""Excel-compatible helper functions.

Ported sheet code should use these instead of Python built-ins wherever Excel's
behaviour differs (rounding, blanks, financial functions, dates).
"""
from __future__ import annotations

import calendar
import datetime as dt
import math
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP, ROUND_UP
from typing import Iterable, Sequence

EXCEL_EPOCH = dt.datetime(1899, 12, 30)


# ---------------------------------------------------------------- blanks / coercion
def n(v) -> float:
    """Excel numeric coercion used inside arithmetic: blank/None/'' -> 0, bool -> 0/1."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (dt.date, dt.datetime)):
        return date_to_serial(v)
    if isinstance(v, str):
        try:
            return float(v)
        except ValueError:
            raise TypeError(f"#VALUE! non-numeric text in arithmetic: {v!r}")
    return float(v)


def is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def nums(values: Iterable) -> list[float]:
    """Values SUM/AVERAGE would count: numbers only (text, blanks, bools ignored)."""
    out = []
    for v in values:
        if isinstance(v, (list, tuple)):
            out.extend(nums(v))
        elif is_number(v):
            out.append(float(v))
    return out


def SUM(*args) -> float:
    return math.fsum(nums(args))


def AVERAGE(*args) -> float:
    v = nums(args)
    if not v:
        raise ZeroDivisionError("#DIV/0!")
    return math.fsum(v) / len(v)


def MAX(*args) -> float:
    v = nums(args)
    return max(v) if v else 0.0


def MIN(*args) -> float:
    v = nums(args)
    return min(v) if v else 0.0


def SUMPRODUCT(*arrays: Sequence) -> float:
    return math.fsum(math.prod(n(x) if is_number(x) or x in (None, "") else 0.0 for x in t)
                     for t in zip(*arrays))


def SUMIF(crit_range: Sequence, crit, sum_range: Sequence | None = None) -> float:
    sum_range = crit_range if sum_range is None else sum_range
    test = criteria(crit)
    return math.fsum(n(s) for c, s in zip(crit_range, sum_range) if test(c) and is_number(s))


def SUMIFS(sum_range: Sequence, *pairs) -> float:
    tests = [(pairs[i], criteria(pairs[i + 1])) for i in range(0, len(pairs), 2)]
    return math.fsum(n(s) for k, s in enumerate(sum_range)
                     if is_number(s) and all(t(rng[k]) for rng, t in tests))


def COUNTIF(rng: Sequence, crit) -> int:
    t = criteria(crit)
    return sum(1 for v in rng if t(v))


def COUNTIFS(*pairs) -> int:
    tests = [(pairs[i], criteria(pairs[i + 1])) for i in range(0, len(pairs), 2)]
    return sum(1 for k in range(len(pairs[0])) if all(t(rng[k]) for rng, t in tests))


def criteria(crit):
    """Excel SUMIF-style criteria: exact (case-insensitive text), or '>5', '<>x', '=x'."""
    if isinstance(crit, str):
        for op in (">=", "<=", "<>", ">", "<", "="):
            if crit.startswith(op):
                rhs = crit[len(op):]
                try:  # Excel accepts percent criteria such as ">30%"
                    num = float(rhs[:-1]) / 100 if rhs.endswith("%") else float(rhs)
                except ValueError:
                    num = None
                def f(v, op=op, rhs=rhs, num=num):
                    if num is not None and is_number(v):
                        a, b = float(v), num
                    else:
                        a, b = ("" if v is None else str(v)).lower(), rhs.lower()
                        if num is not None and op not in ("=", "<>"):
                            return False
                    return {"=": a == b, "<>": a != b, ">": a > b, "<": a < b,
                            ">=": a >= b, "<=": a <= b}[op]
                return f
        low = crit.lower()
        return lambda v: isinstance(v, str) and v.lower() == low
    if is_number(crit):
        return lambda v: is_number(v) and float(v) == float(crit)
    return lambda v: v == crit


def eq(a, b) -> bool:
    """Excel '=' comparison: case-insensitive for text, blank == 0 == ''."""
    if isinstance(a, str) and isinstance(b, str):
        return a.lower() == b.lower()
    if a in (None, "") and b in (None, ""):
        return True
    if a is None and is_number(b):
        return b == 0
    if b is None and is_number(a):
        return a == 0
    return a == b


# ---------------------------------------------------------------- rounding
def _dec(x) -> Decimal:
    # Excel works at 15 significant digits, so 2.675 rounds to 2.68 as users expect.
    return Decimal(f"{n(x):.15g}")


def _round(x, digits, mode) -> float:
    q = Decimal(1).scaleb(-int(digits))
    return float(_dec(x).quantize(q, rounding=mode))


def ROUND(x, digits=0) -> float:
    """Excel ROUND: half away from zero (Python's round() is banker's)."""
    return _round(x, digits, ROUND_HALF_UP)


def ROUNDUP(x, digits=0) -> float:
    """Away from zero."""
    return _round(x, digits, ROUND_UP)


def ROUNDDOWN(x, digits=0) -> float:
    """Toward zero."""
    return _round(x, digits, ROUND_DOWN)


def MROUND(x, multiple) -> float:
    return ROUND(n(x) / multiple, 0) * multiple if multiple else 0.0


def CEILING(x, significance=1) -> float:
    return math.ceil(n(x) / significance - 1e-12) * significance if significance else 0.0


def FLOOR(x, significance=1) -> float:
    return math.floor(n(x) / significance + 1e-12) * significance if significance else 0.0


def INT(x) -> float:
    return float(math.floor(n(x)))


# ---------------------------------------------------------------- financial
def PMT(rate, nper, pv, fv=0.0, typ=0) -> float:
    rate, nper, pv, fv = n(rate), n(nper), n(pv), n(fv)
    if nper == 0:
        raise ZeroDivisionError("#NUM! nper = 0")
    if rate == 0:
        return -(pv + fv) / nper
    f = (1 + rate) ** nper
    return -(rate * (pv * f + fv)) / ((1 + rate * typ) * (f - 1))


def PV(rate, nper, pmt, fv=0.0, typ=0) -> float:
    rate, nper, pmt, fv = n(rate), n(nper), n(pmt), n(fv)
    if rate == 0:
        return -(fv + pmt * nper)
    f = (1 + rate) ** nper
    return -(fv + pmt * (1 + rate * typ) * (f - 1) / rate) / f


def FV(rate, nper, pmt, pv=0.0, typ=0) -> float:
    rate, nper, pmt, pv = n(rate), n(nper), n(pmt), n(pv)
    if rate == 0:
        return -(pv + pmt * nper)
    f = (1 + rate) ** nper
    return -(pv * f + pmt * (1 + rate * typ) * (f - 1) / rate)


def IPMT(rate, per, nper, pv, fv=0.0, typ=0) -> float:
    pmt = PMT(rate, nper, pv, fv, typ)
    if typ == 1 and per == 1:
        return 0.0
    bal = FV(rate, per - 1, pmt, pv, typ)
    i = bal * rate
    return i / (1 + rate) if typ == 1 else i


def PPMT(rate, per, nper, pv, fv=0.0, typ=0) -> float:
    return PMT(rate, nper, pv, fv, typ) - IPMT(rate, per, nper, pv, fv, typ)


def NPV(rate, *values) -> float:
    vals = nums(values)
    return math.fsum(v / (1 + rate) ** (i + 1) for i, v in enumerate(vals))


def IRR(values: Sequence, guess=0.1) -> float:
    vals = nums(values)
    def f(r):
        return math.fsum(v / (1 + r) ** i for i, v in enumerate(vals))
    def df(r):
        return math.fsum(-i * v / (1 + r) ** (i + 1) for i, v in enumerate(vals))
    r = guess
    for _ in range(100):
        fr, d = f(r), df(r)
        if d == 0:
            break
        nr = r - fr / d
        if nr <= -1:
            nr = (r - 1) / 2
        if abs(nr - r) < 1e-12:
            return nr
        r = nr
    # fallback: bisection
    lo, hi = -0.9999, 10.0
    if f(lo) * f(hi) > 0:
        raise ValueError("#NUM! IRR did not converge")
    for _ in range(300):
        mid = (lo + hi) / 2
        if f(lo) * f(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def XNPV(rate, values, dates) -> float:
    d0 = n(dates[0])
    return math.fsum(n(v) / (1 + rate) ** ((n(d) - d0) / 365.0) for v, d in zip(values, dates))


def XIRR(values, dates, guess=0.1) -> float:
    lo, hi = -0.9999, 10.0
    f = lambda r: XNPV(r, values, dates)
    if f(lo) * f(hi) > 0:
        raise ValueError("#NUM! XIRR did not converge")
    for _ in range(300):
        mid = (lo + hi) / 2
        if f(lo) * f(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


# ---------------------------------------------------------------- dates
def date_to_serial(d) -> float:
    if isinstance(d, (int, float)):
        return float(d)
    if not isinstance(d, dt.datetime):
        d = dt.datetime(d.year, d.month, d.day)
    delta = d - EXCEL_EPOCH
    return delta.days + delta.seconds / 86400.0


def serial_to_date(s) -> dt.datetime:
    return EXCEL_EPOCH + dt.timedelta(days=float(s))


def as_date(v) -> dt.datetime:
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    return serial_to_date(n(v))


def EDATE(start, months) -> dt.datetime:
    d = as_date(start)
    m = d.month - 1 + int(n(months))
    y, m = d.year + m // 12, m % 12 + 1
    day = min(d.day, calendar.monthrange(y, m)[1])
    return dt.datetime(y, m, day)


def EOMONTH(start, months) -> dt.datetime:
    d = EDATE(dt.datetime(as_date(start).year, as_date(start).month, 1), months)
    return dt.datetime(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def DATE(y, m, d) -> dt.datetime:
    y, m, d = int(n(y)), int(n(m)), int(n(d))
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return dt.datetime(y, m, 1) + dt.timedelta(days=d - 1)


def YEAR(d) -> int:
    return as_date(d).year


def MONTH(d) -> int:
    return as_date(d).month


def DAY(d) -> int:
    return as_date(d).day


def DATEDIF_M(start, end) -> int:
    """DATEDIF(start, end, "m")."""
    a, b = as_date(start), as_date(end)
    months = (b.year - a.year) * 12 + b.month - a.month
    if b.day < a.day:
        months -= 1
    return months


# ---------------------------------------------------------------- lookup
def INDEX(rng, row, col=None):
    if col is None:
        return rng[int(row) - 1]
    return rng[int(row) - 1][int(col) - 1]


def MATCH(value, rng, match_type=0):
    if match_type == 0:
        for i, v in enumerate(rng):
            if eq(v, value):
                return i + 1
        raise LookupError(f"#N/A MATCH({value!r})")
    if match_type == 1:  # largest <= value, ascending
        best = None
        for i, v in enumerate(rng):
            if v is None:
                continue
            if (is_number(v) and is_number(value) and v <= value) or (
                    isinstance(v, str) and isinstance(value, str) and v.lower() <= value.lower()):
                best = i + 1
            else:
                break
        if best is None:
            raise LookupError("#N/A")
        return best
    best = None  # -1: smallest >= value, descending
    for i, v in enumerate(rng):
        if v is not None and v >= value:
            best = i + 1
        else:
            break
    if best is None:
        raise LookupError("#N/A")
    return best


def VLOOKUP(value, table, col, approx=False):
    keys = [r[0] for r in table]
    i = MATCH(value, keys, 1 if approx else 0)
    return table[i - 1][int(col) - 1]


def XLOOKUP(value, lookup, ret, if_not_found=LookupError):
    for k, v in zip(lookup, ret):
        if eq(k, value):
            return v
    if if_not_found is LookupError:
        raise LookupError(f"#N/A XLOOKUP({value!r})")
    return if_not_found


def IFERROR(fn, fallback):
    """IFERROR(fn(), fallback) where fn is a zero-arg callable."""
    try:
        v = fn()
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return fallback
        return v
    except (ZeroDivisionError, LookupError, ValueError, TypeError, IndexError, OverflowError):
        return fallback


def div(a, b, default=None):
    """a/b with Excel #DIV/0! semantics (raises) unless default is given."""
    b = n(b)
    if b == 0:
        if default is None:
            raise ZeroDivisionError("#DIV/0!")
        return default
    return n(a) / b
