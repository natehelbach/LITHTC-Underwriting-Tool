"""Display values the way an Excel number format would (the common cases in the model)."""
from __future__ import annotations

import datetime as dt
import re
from decimal import ROUND_HALF_UP, Decimal

EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def _sections(fmt: str) -> list[str]:
    out, cur, q = [], "", False
    for ch in fmt:
        if ch == '"':
            q = not q
        if ch == ";" and not q:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    return out + [cur]


def _clean(sec: str) -> tuple[str, str, str]:
    """Split a format section into (prefix literal, numeric pattern, suffix literal)."""
    sec = re.sub(r"\[[^\]]*\]", "", sec)          # [RED], [$-409]
    sec = re.sub(r"_.", "", sec)                   # padding
    sec = re.sub(r"\*.", "", sec)                  # fill
    pre, pat, suf, seen_digit = "", "", "", False
    i = 0
    while i < len(sec):
        ch = sec[i]
        if ch == '"':
            j = sec.index('"', i + 1)
            lit = sec[i + 1:j]
            i = j + 1
        elif ch == "\\":
            lit = sec[i + 1:i + 2]
            i += 2
        elif ch in "0#?,.%":
            pat += ch
            seen_digit = True
            i += 1
            continue
        else:
            lit = ch
            i += 1
        if seen_digit:
            suf += lit
        else:
            pre += lit
    return pre, pat, suf


def is_date_format(fmt: str) -> bool:
    f = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", fmt).lower()
    return bool(re.search(r"[dy]|m{3,}", f)) or f.strip() in ("m/d", "mm/dd")


def _date(v: dt.datetime, fmt: str) -> str:
    f = fmt.lower()
    if "mmmm" in f:
        return f"{v:%B} {v.day}, {v.year}"
    if "mmm" in f and "d" in f:
        return f"{v:%b} {v.day}, {v.year}"
    if "mmm" in f:
        return f"{v:%b-%y}"
    return f"{v.month}/{v.day}/{v.year}"


def _general(v: float) -> str:
    if float(v).is_integer():
        return f"{int(v):,}" if abs(v) >= 10000 else str(int(v))
    if abs(v) >= 1000:
        return f"{v:,.0f}"
    if abs(v) >= 1:
        return f"{v:,.2f}".rstrip("0").rstrip(".")
    return f"{v:.4g}"


def fmt(v, numfmt: str | None = "General") -> str:
    numfmt = numfmt or "General"
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (dt.datetime, dt.date)):
        if not isinstance(v, dt.datetime):
            v = dt.datetime(v.year, v.month, v.day)
        return _date(v, numfmt if is_date_format(numfmt) else "m/d/yyyy")
    if isinstance(v, str):
        return v
    if not isinstance(v, (int, float)):
        return str(v)
    if is_date_format(numfmt) and 0 < v < 200000:
        return _date(EXCEL_EPOCH + dt.timedelta(days=v), numfmt)
    if numfmt in ("General", "@"):
        return _general(v)

    secs = _sections(numfmt)
    if v != 0:  # a value that displays as zero uses the zero section, like Excel
        pre0, pat0, _ = _clean(secs[0])
        dec0 = len(pat0.split(".", 1)[1].replace("%", "").replace(",", "")) if "." in pat0 else 0
        if round(abs(v) * 100 ** pat0.count("%"), dec0) == 0:
            v = 0
    if v < 0 and len(secs) > 1:
        sec, val, neg = secs[1], -v, False
    elif v == 0 and len(secs) > 2:
        sec, val, neg = secs[2], 0, False
    else:
        sec, val, neg = secs[0], abs(v), v < 0
    pre, pat, suf = _clean(sec)
    if pat.strip("?") == "":  # literal-only section, e.g. "-" for zero
        return (pre + suf).strip() or "–"
    if "$" in pre and pre.rstrip().endswith("("):
        pre = "($"
    if not pat:
        return (pre + suf).strip() or "–"
    pct = pat.count("%")
    val = val * (100 ** pct)
    dec = len(pat.split(".", 1)[1].replace("%", "").replace(",", "")) if "." in pat else 0
    val = Decimal(repr(float(val))).quantize(Decimal(1).scaleb(-dec), ROUND_HALF_UP)  # Excel rounds half up
    body = f"{val:,.{dec}f}" if "," in pat else f"{val:.{dec}f}"
    out = pre + body + "%" * pct + suf
    return ("-" + out if neg else out).strip()


def input_text(v, numfmt: str | None) -> str:
    """Editable text for an input: formatted like Excel but without losing precision."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, str):
        return v
    if isinstance(v, (int, float)):
        f = numfmt or "General"
        if "%" in f and not is_date_format(f):
            return _plain(v * 100) + "%"
        s = _plain(v)
        if "$" in f:
            s = ("-$" + s[1:]) if s.startswith("-") else "$" + s
        return s
    return str(v)


def _plain(v: float) -> str:
    """Comma-grouped number with only the decimals it needs (up to 10)."""
    if float(v).is_integer():
        return f"{int(round(v)):,}"
    s = f"{v:,.10f}".rstrip("0").rstrip(".")
    return s
