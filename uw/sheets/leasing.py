"""Leasing: month-by-month lease-up from placed-in-service to conversion.

Units are grouped into segments (30/50/60/70/80% AMI, market, manager). Each AMI /
market segment fills along a power curve: cumulative leased units in month m are
    units * stab_occ * min(1, (m - lag) / months_to_stabilize) ^ exponent,
and those units are handed out to the unit types top-down (3 BR first, then 2 BR,
1 BR, Studio). Manager units lease one per month. Rent, concessions (first N
months of each lease free), other income and operating costs then roll up to
NOI during lease-up, which funds part of the development budget.

Columns L:M are the two pre-leasing months (-2, -1); N:AC are lease-up months 1-16.
"""
from __future__ import annotations

import datetime as dt

from openpyxl.utils import get_column_letter as gl

from uw import xl
from uw.engine import Sheet, register

PRE = ["L", "M"]                                   # pre-leasing months -2, -1
MONTHS = [gl(c) for c in range(14, 30)]            # N..AC = lease-up months 1..16
ALL = PRE + MONTHS
UNIT_TYPES = ["3 BR", "2 BR", "1 BR", "Studio"]

PRELEASE_MONTHS = 2      # col F, typed on the tab for every row
PRELEASE_PCT = 0         # col G, typed on the tab for every row


class Segment:
    """One AMI (or market) segment: unit rows, then a cumulative row and a rent row."""

    def __init__(self, key, ami, first, cum, rows, studio_ami="same"):
        self.key, self.ami, self.first, self.cum, self.rent = key, ami, first, cum, cum + 1
        self.rows = list(range(first, cum))          # rows carrying a unit count (col D)
        self.active = rows                           # rows that actually lease up
        self.last = rows[-1]
        # (unit type, AMI) per row, used for the Rent Builder lookups.
        self.keys = {}
        types = iter(UNIT_TYPES)
        for r in self.rows:
            if r in rows:
                self.keys[r] = (next(types), ami)
        studio_row = next((r for r in self.rows if r not in rows and r > self.last), None)
        if len(rows) == 3 and studio_row:
            # EXCEL-QUIRK: in the 70%/80%/market blocks the Studio row has a unit count
            # (counted in the segment total) but no lease-up formulas, so studios at
            # those AMIs never lease. The 70% Studio row also has no AMI in col C.
            self.keys[studio_row] = ("Studio", None if studio_ami is None else ami)


SEGMENTS = [
    Segment("30", 0.3, 8, 13, [8, 9, 10, 12]),
    Segment("50", 0.5, 16, 21, [16, 17, 18, 20]),
    Segment("60", 0.6, 24, 29, [24, 25, 26, 28]),
    Segment("70", 0.7, 32, 37, [32, 33, 34], studio_ami=None),
    Segment("80", 0.8, 40, 45, [40, 41, 42]),
    Segment("mkt", "MKT", 48, 53, [48, 49, 50]),
]
MGR_ROWS = [56, 57, 58, 59]                        # manager units, one per unit type
# Other income rows: row -> (Operating Cash Flow label row)
OTHER_INCOME = {69: "E15", 70: "E16", 71: "E17", 72: "E18", 73: "E19"}
# Operating cost rows: row -> (Operating Cash Flow annual cell, sign) ; monthly = annual/12
OPEX_ROWS = {78: ("J45", -1), 79: ("J49", 1), 80: ("J50", 1), 81: ("J51", 1),
             82: ("J53", 1), 83: ("J67", -1), 85: ("J71", -1)}
AVG_CONCESSION_WEEKS = 6                            # C93
TCO_SHARE = {98: 0.2, 99: 0.5, 100: 0}              # K98:K100 phased TCO shares (typed)


@register
class Leasing(Sheet):
    name = "Leasing"

    def compute(self, ctx):
        mi = ctx.mi
        o: dict = {}

        def v(a):
            return o.get(a)

        rb = {c: ctx.col("Rent Builder", f"{c}12:{c}38") for c in "FGHJOUV"}
        fof = lambda a: ctx.ref("Flow of Funds", a)  # noqa: E731
        o["B2"] = f"{mi['project_name']}  |  LEASING"

        # ---- schedule: model months and the phase table (rows 98-102)
        o["C98"] = xl.n(fof("F13")) - 1          # pre-leasing starts
        o["C99"] = xl.n(fof("F13")) + 1          # leasing (placed in service)
        o["C100"] = xl.n(fof("F15")) + 1         # stabilization
        o["C101"] = xl.n(fof("F17"))             # conversion
        o["C102"] = o["C101"] + 1
        o["D102"] = ""
        phases = [[o["C98"], "Pre-leasing"], [o["C99"], "Leasing"], [o["C100"], "Stabilization"],
                  [o["C101"], "Conversion"], [o["C102"], ""]]
        leasing_window = o["C100"] - o["C99"] + 1
        o["N5"] = xl.n(fof("F13")) + 1
        o["M5"], o["L5"] = o["N5"] - 1, o["N5"] - 2
        month_no = {"L": -2, "M": -1}
        for i, c in enumerate(MONTHS):
            o[f"{c}5"] = o["N5"] + i
            month_no[c] = o[f"{c}7"] = i + 1
        for c in ALL:
            o[f"{c}6"] = xl.VLOOKUP(o[f"{c}5"], phases, 2, True)
        after_conversion = {c: o[f"{c}5"] > o["C101"] for c in MONTHS}

        per_unit_rent = xl.eq(mi["lease_rent_basis"], "Per-unit")
        stab = xl.n(mi["stab_occ"])

        # ---- AMI / market segments
        for s in SEGMENTS:
            lag, rate = xl.n(mi[f"lease_lag_{s.key}"]), xl.n(mi[f"lease_rate_{s.key}"])
            exp = xl.n(mi[f"lease_exp_{s.key}"])
            for r in s.rows:
                b, ami = s.keys.get(r, (None, None))
                o[f"D{r}"] = xl.SUMIFS(rb["G"], rb["H"], b, rb["J"], ami)
                # monthly rent roll for the type: units x rent
                o[f"K{r}"] = sum(xl.n(g) * xl.n(rent) for g, h, j, rent in
                                 zip(rb["G"], rb["H"], rb["J"], rb["O"])
                                 if xl.eq(j, ami) and xl.eq(h, b))
            o[f"K{s.cum}"] = 0   # same lookup on the "Cumlative" label row: never matches
            for r in s.active:
                o[f"H{r}"] = 1 + lag                                 # placed-in-service month
                o[f"E{r}"] = o[f"H{r}"] - PRELEASE_MONTHS - 1        # pre-leasing start
                o[f"L{r}"] = o[f"M{r}"] = xl.ROUND(PRELEASE_PCT * o[f"D{r}"] / PRELEASE_MONTHS, 0)

            block = range(s.first, s.last + 1)
            total = xl.SUM([o.get(f"D{r}") for r in block])
            preleased = xl.SUM([o.get(f"{c}{r}") for r in block for c in PRE])
            to_lease = total - preleased
            months_to_stab = max(1, xl.ROUNDUP(leasing_window * rate, 0))
            o[f"D{s.cum}"] = xl.SUM([o[f"D{r}"] for r in s.rows])
            o[f"G{s.cum}"] = exp
            o[f"J{s.cum}"] = months_to_stab
            o[f"I{s.cum}"] = 0 if months_to_stab <= 0 else xl.ROUND(to_lease / months_to_stab, 0)
            o[f"C{s.rent}"] = mi[f"conc_{s.key}"]
            o[f"K{s.rent}"] = xl.SUM([o.get(f"K{r}") for r in range(s.first, s.cum + 1)])

            # Cumulative leased units by month (power curve, capped at stabilized occupancy).
            cap = xl.ROUND(total * stab, 0)
            for c in MONTHS:
                progress = min(1, max(0, month_no[c] - lag) / months_to_stab) ** exp
                o[f"{c}{s.cum}"] = min(cap, xl.ROUND(total * stab * progress, 0))

            # Allocate each month's new leases to unit types, top row first.
            for r in s.active:
                cum_d = xl.SUM([o.get(f"D{k}") for k in range(s.first, r + 1)])
                cum_pre = xl.SUM([o.get(f"{c}{k}") for k in range(s.first, r + 1) for c in PRE])
                if to_lease <= 0:
                    o[f"J{r}"] = o[f"H{r}"]
                else:  # month this type is fully leased, from the inverse curve
                    o[f"J{r}"] = max(o[f"H{r}"], xl.ROUNDUP(
                        months_to_stab * ((cum_d - cum_pre) / to_lease) ** (1 / exp), 0))
                leased = o[f"L{r}"] + o[f"M{r}"]
                prev_cum = preleased
                for c in MONTHS:
                    new = o[f"{c}{s.cum}"] - prev_cum
                    prev_cum = o[f"{c}{s.cum}"]
                    if month_no[c] < o[f"H{r}"]:
                        x = 0
                    else:
                        above = xl.SUM([o.get(f"{c}{k}") for k in range(s.first, r)])
                        x = min(o[f"D{r}"] - leased, max(0, new - above))
                    o[f"{c}{r}"] = x
                    leased += x
            for r in block:
                if r in s.active or f"D{r}" in o:
                    o[f"I{r}"] = xl.MAX([o.get(f"{c}{r}") for c in MONTHS])

            # Rent collected: cumulative leased units x rent per unit.
            def unit_rent(r):
                if per_unit_rent:
                    return xl.IFERROR(lambda: xl.div(o.get(f"K{r}"), o.get(f"D{r}")), 0)
                return xl.n(o.get(f"K{r}"))
            rents = {r: unit_rent(r) for r in s.rows}
            for i, c in enumerate(MONTHS):
                o[f"{c}{s.rent}"] = sum(xl.SUM([o.get(f"{m}{r}") for m in ALL[:i + 3]]) * rents[r]
                                        for r in s.rows)

        # ---- manager units: one per month from placed-in-service until stabilization
        for r, b in zip(MGR_ROWS, UNIT_TYPES):
            o[f"D{r}"] = xl.SUMIFS(rb["F"], rb["H"], b)
            o[f"K{r}"] = sum(xl.n(u) * xl.n(w) for h, u, w in zip(rb["H"], rb["U"], rb["V"])
                             if xl.eq(h, b))
            start, per_month = 1, 1                       # H, I typed on the tab
            o[f"E{r}"] = start - PRELEASE_MONTHS - 1
            o[f"J{r}"] = max(start, leasing_window)
            # EXCEL-QUIRK: pre-leasing for manager units sits in month -2 (col L) for
            # 3 BR / 2 BR but month -1 (col M) for 1 BR / Studio.
            pre = xl.ROUND(PRELEASE_PCT * o[f"D{r}"] / PRELEASE_MONTHS, 0)
            o[f"L{r}"], o[f"M{r}"] = (pre, 0) if r in (56, 57) else (0, pre)
            leased = o[f"L{r}"] + o[f"M{r}"]
            for c in MONTHS:
                x = min(per_month if start <= month_no[c] <= o[f"J{r}"] else 0, o[f"D{r}"] - leased)
                o[f"{c}{r}"] = x
                leased += x
        # EXCEL-QUIRK: C59 (manager Studio "AMI") links to Rent Builder D63, not "MGR".
        o["C59"] = ctx.ref("Rent Builder", "D63")
        for i, c in enumerate(MONTHS):
            o[f"{c}60"] = sum(xl.SUM([o[f"{m}{r}"] for m in ALL[:i + 3]]) * o[f"K{r}"] for r in MGR_ROWS)
        o["D61"] = sum(o[f"D{r}"] for r in MGR_ROWS)
        o["K62"] = sum(o[f"K{r}"] for r in MGR_ROWS)

        # ---- totals and occupancy
        o["D63"] = o["D61"] + sum(o[f"D{s.cum}"] for s in SEGMENTS)
        o["K63"] = o["K62"] + sum(o[f"K{s.rent}"] for s in SEGMENTS)
        # EXCEL-QUIRK: L63/M63 sum rows 8-58, so they miss the manager Studio row 59.
        for c in PRE:
            o[f"{c}63"] = xl.SUM([o.get(f"{c}{r}") for r in range(8, 59)])
        restricted_units = o["D63"] - sum(o[f"D{r}"] for r in MGR_ROWS)
        for c in MONTHS:
            o[f"{c}63"] = sum(o[f"{c}{s.cum}"] for s in SEGMENTS)       # cumulative leased
            o[f"{c}64"] = sum(o[f"{c}{r}"] for s in SEGMENTS for r in s.active)  # leased this month
            o[f"{c}65"] = o[f"{c}63"] / restricted_units                # % occupied
        o["I63"] = max(o[f"{c}64"] for c in MONTHS)

        # ---- revenue
        for c in MONTHS:
            gross = o[f"{c}60"] + sum(o[f"{c}{s.rent}"] for s in SEGMENTS)
            o[f"{c}67"] = 0 if after_conversion[c] else gross
            conc = 0.0
            for s in SEGMENTS:  # first N months of each lease are free
                free = xl.n(o[f"C{s.rent}"])
                lagged = 0 if month_no[c] <= free else \
                    xl.INDEX([o[f"{m}{s.rent}"] for m in MONTHS], int(month_no[c] - free))
                conc += o[f"{c}{s.rent}"] - lagged
            o[f"{c}68"] = 0 if after_conversion[c] else -conc

        ums = lambda a: ctx.ref("Unit Mix Summary", a)  # noqa: E731
        sens = lambda a: ctx.ref("Sensitivities", a)  # noqa: E731
        for r, lbl in OTHER_INCOME.items():
            o[f"B{r}"] = ctx.ref("Operating Cash Flow", lbl)
        o["D69"], o["E69"], o["K69"] = ums("F36"), sens("C33"), mi["parking_rent"]
        o["D70"], o["E70"], o["K70"] = ums("F37"), mi["antenna_vacancy"], ctx.ref("OpEx & Other Income", "E90")
        o["D71"], o["E71"], o["K71"] = ums("F38"), sens("C31"), sens("C24")
        o["D72"], o["E72"], o["K72"] = o["D63"], mi["laundry_vacancy"], mi["laundry_income"]
        o["D73"], o["E73"], o["K73"] = o["D63"], sens("C30"), sens("C25")
        for c in MONTHS:
            occ = o[f"{c}63"]
            if after_conversion[c]:
                for r in OTHER_INCOME:
                    o[f"{c}{r}"] = 0
                continue
            # EXCEL-QUIRK: parking is driven by the 70% AMI cumulative row (N37), not all units.
            o[f"{c}69"] = xl.n(o["K69"]) * min(o[f"{c}37"], xl.n(o["D69"])) * (1 - xl.n(o["E69"]))
            o[f"{c}70"] = xl.n(o["K70"]) * 1
            for r in (71, 72, 73):
                o[f"{c}{r}"] = xl.n(o[f"K{r}"]) * occ * (1 - xl.n(o[f"E{r}"]))
        for c in MONTHS:
            o[f"{c}75"] = 0 if after_conversion[c] else sum(o[f"{c}{r}"] for r in range(67, 74))
        for r in list(range(67, 74)):
            o[f"C{r}"] = sum(o[f"{c}{r}"] for c in MONTHS)
        o["C75"] = sum(o[f"C{r}"] for r in range(67, 74))

        # ---- operating costs during lease-up (monthly from the stabilized budget)
        ocf = lambda a: xl.n(ctx.ref("Operating Cash Flow", a))  # noqa: E731
        for r, (cell, sign) in OPEX_ROWS.items():
            o[f"D{r}"] = o["D63"]
            o[f"K{r}"] = sign * ocf(cell) / 12
        o["D84"] = o["D63"]
        o["L84"] = mi["pre_opening_start_up_cost_month_2"]
        o["M84"] = mi["pre_opening_start_up_cost_month_1"]
        for c in MONTHS:
            if after_conversion[c]:
                for r in OPEX_ROWS:
                    o[f"{c}{r}"] = 0
                continue
            o[f"{c}78"] = -o["K78"] / o["D78"] * o[f"{c}63"]   # scales with occupancy (E78 blank)
            for r in OPEX_ROWS:
                if r != 78:
                    o[f"{c}{r}"] = -o[f"K{r}"]
        for r in OPEX_ROWS:
            o[f"C{r}"] = sum(o[f"{c}{r}"] for c in MONTHS)
        for c in PRE:
            o[f"{c}87"] = xl.n(o[f"{c}84"])
            o[f"{c}89"] = o[f"{c}87"]
        for c in MONTHS:
            o[f"{c}87"] = 0 if after_conversion[c] else sum(o.get(f"{c}{r}", 0) for r in range(78, 86))
            o[f"{c}89"] = 0 if after_conversion[c] else o[f"{c}87"] + o[f"{c}75"]
        o["C87"] = sum(o[f"C{r}"] for r in OPEX_ROWS)
        o["D87"] = -o["C87"] / o["C75"]

        # ---- NOI during lease-up (plus the Dev Budget's lease-up line)
        o["C89"] = sum(o[f"{c}89"] for c in ALL) + xl.n(ctx.ref("Dev Budget", "H82"))
        o["D91"] = 1
        o["C91"] = o["C89"] * o["D91"]
        o["C93"] = AVG_CONCESSION_WEEKS
        o["C95"] = o["C91"] * ((52 - o["C93"]) / 52)

        # ---- phased-TCO interest estimate (rows 94-101)
        o["M94"] = xl.n(fof("D216")) + xl.n(fof("D223")) + xl.n(fof("D230"))
        o["M95"] = xl.n(fof("D232")) - o["M94"]
        o["M98"] = o["M94"]
        for r, share in TCO_SHARE.items():
            o[f"L{r}"] = share * o["D63"]
        conv_months = o["C101"] - o["C99"] + 1
        count = lambda rows, cols: sum(xl.is_number(o.get(f"{c}{r}")) for r in rows for c in cols)  # noqa: E731
        # Interest spread over the months to each TCO date (Q = month 4, U = month 8).
        o["Q99"] = o["M95"] * TCO_SHARE[99] * count([7], MONTHS[:4]) / conv_months
        # EXCEL-QUIRK: U100's COUNT range drifted to N7:U8 (two rows), counting 16, not 8.
        o["U100"] = o["M95"] * TCO_SHARE[100] * count([7, 8], MONTHS[:8]) / conv_months
        o["AB101"] = xl.SUM([o.get(f"{gl(c)}{r}") for r in (98, 99, 100) for c in range(13, 29)])
        o["AC101"] = xl.IFERROR(lambda: xl.div(o["AB101"], o["M94"] + o["M95"]), 0)

        # ---- project schedule
        close = mi["constr_close"]
        constr_months = xl.n(sens("C26"))
        o["F111"] = xl.as_date(close) + dt.timedelta(days=30 * constr_months)
        o["D113"] = xl.n(sens("C27")) + constr_months
        o["D114"] = xl.DATE(xl.YEAR(close), xl.MONTH(close) + o["D113"], xl.DAY(close))
        o["D119"] = constr_months / (o["D113"] + 12 * 15)
        return o
