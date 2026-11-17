"""Rent Builder: restricted (LIHTC) and market rent for every unit line, and the final rent.

Each unit line is a Unit Takeoff plan at one AMI level (or a manager unit). For each line:
  max rent      LIHTC max rent net of utility allowance (Unit Mix Summary grid, by AMI x type)
  market rent   Comps Summary market rent for the unit type, less the target discount
  adjusted rent the max rent, unless the max is above market-less-discount, then market-less-discount
  near-max rent max rent less a per-AMI haircut (Master Inputs)
  final rent    picked by the Master Inputs rent mode (Adjusted / Near Max / Max)
Totals feed annual gross potential rent and the AMI-weighted unit mix used downstream.
"""
from __future__ import annotations

import math

from uw import xl
from uw.engine import Sheet, register

FIRST, LAST = 12, 39          # unit-line rows
NEAR_MAX_AMIS = [0.3, 0.5, 0.6, 0.7, 0.8]
TYPE_LABELS = ["STUDIO", "1 BR", "2 BR", "3 BR"]


def _blank(v) -> bool:
    return v is None or v == ""


def _num(v) -> float:
    """SUMPRODUCT coercion: text and blanks count as 0."""
    return float(v) if xl.is_number(v) else 0.0


def _sumproduct(*cols) -> float:
    return math.fsum(math.prod(_num(c) for c in vals) for vals in zip(*cols))


@register
class RentBuilder(Sheet):
    name = "Rent Builder"

    # Unit lines typed on the tab: Unit ID (plan + "-" + AMI tag) and AMI / status, rows 12-21.
    UNIT_LINES = [
        ("S400-.30", 0.3), ("S400-.50", 0.5),
        ("1B560-.30", 0.3), ("1B560-.50", 0.5), ("1B560-.60", 0.6), ("1B560-.70", 0.7),
        ("1B560-MGR", "Manager"),
        ("2B720-.30", 0.3), ("2B720-.70", 0.7), ("2B720-MGR", "Manager"),
    ]
    # Unit Takeoff count column by AMI / status.
    TAKEOFF_COUNT_COL = {"Manager": "M", 0.3: "O", 0.5: "P", 0.6: "Q", 0.7: "R", 0.8: "S"}
    LABELS = {
        "C5": "UNIT MIX CHECK vs Unit Mix Summary target", "L7": "Project shortfall",
        "C10": "Target discount to market (per unit - input)", "L10": "Max Rents",
        "O10": "Adjusted Rents", "R10": "Market Rents", "U10": "Manager Units",
        "X10": "Discount on Adjusted", "AB10": "Discount on Max", "AD10": "Rent strategy",
        "B11": "Unit ID", "C11": "Target Disc.", "D11": "Unit Size", "E11": "Total Units",
        "F11": "MGR Units", "G11": "RES. Units", "H11": "Unit Type (AMI Calcs)",
        "I11": "Unit Type (Rent Comps)", "J11": "AMI / Status", "L11": "Unit", "M11": "SF",
        "O11": "Unit", "P11": "SF", "R11": "Unit", "S11": "SF", "U11": "Charged Units",
        "V11": "Rent", "X11": "Unit", "Y11": "SF", "AB11": "Unit", "AC11": "SF",
        "AD11": "Eff. discount", "AE11": "Near-Max rent", "AF11": "Final rent",
        "AG11": "Mgr host", "AN11": "Bedroom (Takeoff)",
        "C46": "MAX RENTS", "C47": "ADJUSTED RENTS",
        "C48": "FINAL RENTS (per Master Inputs rent mode)", "C50": "AMI BASED AVG",
        "C51": "Unit Type", "D51": "Unit Size", "E51": "Unit Count", "F51": "AMI",
        "G51": "Actual AMI", "C52": "Restricted", "C59": "Restricted", "C67": "Total",
        "AH19": "RECONCILIATION CHECKS", "AH20": "Unit mix check", "AI20": "Result",
        "AK20": "Market rent check", "AL20": "Result", "AH21": "Affordable units (AMI detail)",
        "AH22": "Total units incl. managers", "AK21": "Studio comp count",
        "AK22": "1 BR comp count", "AH23": "Takeoff units not mapped",
    }

    def compute(self, ctx):
        mi = ctx.mi
        o = {"B2": f"{mi['project_name']}  |  RENT BUILDER", **self.LABELS}
        o["D6"] = ctx.ref("Unit Mix Summary", "T26")
        o["E6"] = ("Unit Takeoff and Rent Builder match the target program" if o["D6"] == "PASS"
                   else "Mismatch - see Unit Mix Summary variance grid (cols AX:BB)")
        o["O7"] = ctx.ref("Sources & Uses", "E58")

        # Lookups.
        takeoff_ids = ctx.col("Unit Takeoff", "C6:C94")
        takeoff = {c: ctx.col("Unit Takeoff", f"{c}6:{c}94") for c in "BHMOPQRS"}
        max_rent_amis = ctx.col("Unit Mix Summary", "L44:L54")
        max_rent_types = ctx.col("Unit Mix Summary", "M43:Q43")
        max_rent_grid = ctx.grid("Unit Mix Summary", "M44:Q54")
        market_types = ctx.col("Comps Summary", "C12:F12")
        market_rents = ctx.col("Comps Summary", "C13:F13")
        discount_by_type = {"STUDIO": mi.by_row(386), "1 BR": mi.by_row(387), "2 BR": mi.by_row(388)}
        discount_3br = mi.by_row(389)
        near_max = [mi["near_max_30"], mi["near_max_50"], mi["near_max_60"],
                    mi["near_max_70"], mi["near_max_80"]]
        rent_mode = mi["rent_mode"]

        self._lk = dict(mi=mi, takeoff_ids=takeoff_ids, takeoff=takeoff, max_rent_amis=max_rent_amis,
                        max_rent_types=max_rent_types, max_rent_grid=max_rent_grid,
                        market_types=market_types, market_rents=market_rents,
                        discount_by_type=discount_by_type, discount_3br=discount_3br,
                        near_max=near_max, rent_mode=rent_mode)

        for i, r in enumerate(range(FIRST, LAST + 1)):
            uid, ami = self.UNIT_LINES[i] if i < len(self.UNIT_LINES) else (None, None)
            o[f"B{r}"], o[f"J{r}"] = uid, ami
            line = self.unit_line(uid, ami)
            for col, v in line.items():
                o[f"{col}{r}"] = v

        rows = range(FIRST, LAST + 1)
        rows38 = range(FIRST, 39)   # the SUMPRODUCT totals stop at row 38

        def col(c, rr=rows):
            return [o.get(f"{c}{r}") for r in rr]

        # Totals and annual rent roll.
        o["E44"], o["F44"], o["G44"] = xl.SUM(col("E")), xl.SUM(col("F")), xl.SUM(col("G"))
        res_units = col("G", rows38)
        o["L44"] = _sumproduct(col("L", rows38), res_units)
        o["O44"] = _sumproduct(col("O", rows38), res_units)
        o["AF44"] = _sumproduct(col("AF", rows38), res_units)
        o["V44"] = 0  # manager-unit rent column V12:V43 is empty
        o["D46"] = (o["L44"] + o["V44"]) * 12   # max rents, annual
        o["D47"] = (o["O44"] + o["V44"]) * 12   # adjusted rents, annual
        o["I47"] = o["O44"] + o["V44"]
        o["D48"] = (o["AF44"] + o["V44"]) * 12  # final rents, annual

        # Reconciliation checks.
        o["AI21"] = xl.SUM(col("G"))
        o["AI22"] = o["E44"]
        o["AI23"] = xl.n(ctx.ref("Unit Takeoff", "L95")) - o["E44"]
        o["AL21"] = self.comp_count(ctx, "Studio Comps", 94)
        o["AL22"] = self.comp_count(ctx, "1BR Comps", 107)

        # AMI-based averages by unit type.
        comp_type, limit_type = col("I", rows38), col("H", rows38)
        sf, total, mgr, ami = col("D", rows38), col("E", rows38), col("F", rows38), col("J", rows38)

        def wavg(types, label, values, weights):
            is_type = [1.0 if isinstance(t, str) and xl.eq(t, label) else 0.0 for t in types]
            den = xl.SUMIF(types, label, weights)
            return _sumproduct(is_type, weights, values) / den if den else 0

        for k, label in enumerate(TYPE_LABELS):
            r1, r2, r3 = 53 + k, 60 + k, 68 + k
            o[f"C{r1}"] = o[f"C{r2}"] = o[f"C{r3}"] = label
            # Restricted units (row 53+): unit size and AMI by comp basis.
            o[f"D{r1}"] = wavg(comp_type, label, sf, res_units)
            o[f"E{r1}"] = xl.SUMIF(comp_type, label, res_units)
            # EXCEL-QUIRK: the AMI average weights by total units (E) incl. manager units,
            # whose "Manager" status counts as 0% AMI, so it understates restricted AMI.
            o[f"F{r1}"] = wavg(comp_type, label, ami, total)
            # Manager units (row 60+) and all units (row 68+), by limit basis.
            o[f"D{r2}"] = wavg(limit_type, label, sf, mgr)
            o[f"E{r2}"] = xl.SUMIF(limit_type, label, mgr)
            o[f"D{r3}"] = wavg(limit_type, label, sf, total)
            o[f"E{r3}"] = xl.SUMIF(limit_type, label, total)
        for first, last in ((53, 57), (60, 64), (68, 72)):
            rr = range(first, first + 4)
            o[f"D{last}"] = xl.SUMPRODUCT([o[f"D{r}"] for r in rr], [o[f"E{r}"] for r in rr])
            o[f"E{last}"] = xl.SUM([o[f"E{r}"] for r in rr])
        o["F57"] = xl.SUMPRODUCT([o[f"E{r}"] for r in range(53, 57)],
                                 [o[f"F{r}"] for r in range(53, 57)]) / o["E57"]

        # Actual AMI: AMI weighted by restricted units and by adjusted/max rent.
        # EXCEL-QUIRK: the three SUMPRODUCT blocks cover rows 12-20, 22-29 and 31-38,
        # skipping rows 21 and 30 (row 21 is a manager line today, so no effect).
        def actual(rr):
            return math.fsum(_num(o.get(f"G{r}")) * _num(o.get(f"J{r}"))
                             * xl.IFERROR(lambda: o[f"O{r}"] / o[f"L{r}"], 0) for r in rr)
        o["G57"] = (actual(range(12, 21)) + actual(range(22, 30)) + actual(range(31, 39))) / o["G44"]
        return o

    def unit_line(self, uid, ami):
        """All formula columns of one unit line."""
        lk = self._lk
        takeoff_ids, takeoff = lk["takeoff_ids"], lk["takeoff"]

        def takeoff_lookup(plan, col, default):
            for k, v in zip(takeoff_ids, takeoff[col]):
                if not _blank(k) and xl.eq(k, plan):
                    return 0 if v is None else v
            return default

        c = {}
        blank = _blank(uid)
        is_mgr = xl.eq(ami, "Manager")
        plan = None if blank else uid[:len(uid) - 4]   # "1B560-.30" -> "1B560"
        c["AN"] = "" if blank else takeoff_lookup(plan, "B", "ID not in Takeoff")   # bedroom type
        # Target discount to market by bedroom type (blank / unknown -> the 3 BR input).
        c["C"] = next((d for t, d in lk["discount_by_type"].items() if xl.eq(c["AN"], t)),
                      lk["discount_3br"])
        c["D"] = "" if blank else takeoff_lookup(plan, "H", 0)                     # unit SF
        if blank:
            c["E"] = c["F"] = c["G"] = c["H"] = c["I"] = ""
        else:
            count_col = next((col for k, col in self.TAKEOFF_COUNT_COL.items() if xl.eq(ami, k)), None)
            c["E"] = 0 if count_col is None else xl.SUMIFS(takeoff[count_col], takeoff_ids, plan)
            c["F"] = c["E"] if is_mgr else 0
            c["G"] = c["E"] - c["F"]
            studio = xl.eq(c["AN"], "STUDIO")
            c["H"] = lk["mi"]["studio_limit_basis"] if studio else c["AN"]   # type for LIHTC limits
            c["I"] = lk["mi"]["studio_comp_basis"] if studio else c["AN"]    # type for market comps
        no_rent = blank or is_mgr

        # Max (LIHTC) rent and per-SF.
        if no_rent:
            c["L"] = 0
        else:
            c["L"] = xl.INDEX(lk["max_rent_grid"], xl.MATCH(ami, lk["max_rent_amis"]),
                              xl.MATCH(c["H"], lk["max_rent_types"]))
        c["M"] = xl.IFERROR(lambda: c["L"] / c["D"], 0)
        # Market rent and per-SF.
        c["R"] = 0 if blank else xl.INDEX(lk["market_rents"], xl.MATCH(c["I"], lk["market_types"]))
        c["S"] = 0 if blank or c["D"] == 0 else c["R"] / c["D"]
        c["AD"] = 0 if blank else c["C"]
        c["AC"] = xl.IFERROR(lambda: (c["S"] - c["M"]) / c["S"], 0)   # discount of max vs market
        # Adjusted rent: max rent, capped at market less the target discount.
        if no_rent:
            c["P"] = 0
        else:
            c["P"] = c["M"] if c["AC"] > c["AD"] else c["S"] - c["S"] * c["AD"]
        c["O"] = 0 if blank else c["P"] * c["D"]
        c["X"] = xl.IFERROR(lambda: (c["R"] - c["O"]) / c["R"], 0)
        c["Y"] = xl.IFERROR(lambda: (c["S"] - c["P"]) / c["S"], 0)
        c["Z"] = "" if blank else ("CHECK" if c["P"] > c["M"] else "OK")
        c["AB"] = xl.IFERROR(lambda: (c["R"] - c["L"]) / c["R"], 0)
        # Near-max rent: max rent less the per-AMI haircut.
        c["AE"] = 0 if no_rent else c["L"] * (1 - lk["near_max"][xl.MATCH(ami, NEAR_MAX_AMIS) - 1])
        # Final rent per rent mode.
        if no_rent or c["E"] == 0:
            c["AF"] = 0
        elif xl.eq(lk["rent_mode"], "Adjusted"):
            c["AF"] = c["O"]
        elif xl.eq(lk["rent_mode"], "Near Max"):
            c["AF"] = c["AE"]
        else:
            c["AF"] = c["L"]
        c["AG"] = "Y" if is_mgr else ""
        return c

    @staticmethod
    def comp_count(ctx, tab, last):
        include = ctx.col(tab, f"F9:F{last}")
        rent = ctx.col(tab, f"H9:H{last}")
        return xl.COUNTIFS(include, "Y", rent, ">0")
