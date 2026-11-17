"""Operating expenses.

OpEx Builder        - line-item operating budget ($/unit/yr). Each line either takes a
                      typed override (col G), the average of six comparable LIHTC
                      properties' size/age-adjusted costs, or one named comp. The comp
                      figures are typed data on the tab, kept in
                      data/site8/tables/opex_builder.json (tools/extract_opex_builder.py).
OpEx & Other Income - OpEx summary (builder vs. federal minimum), other-income detail
                      (internet, laundry, storage, parking, antenna) and antenna lease
                      tracker / escalation.
"""
from __future__ import annotations

import datetime as dt
import json
from functools import lru_cache

from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl

from uw import xl
from uw.engine import DATA, Sheet, register


@lru_cache(maxsize=None)
def _table():
    d = json.loads((DATA / "site8" / "tables" / "opex_builder.json").read_text(encoding="utf-8"))
    d["live"] = {col: set(rows) for col, rows in d["live"].items()}
    return d


def _col(c: str, k: int) -> str:
    """Column letter k columns to the right of c."""
    return gl(ci(c) + k)


def _pct(v) -> float:
    """Excel coercion of a value such as '5%' (text typed on Master Inputs) to a number."""
    if isinstance(v, str) and v.strip().endswith("%"):
        return float(v.strip()[:-1]) / 100
    return xl.n(v)


# ---------------------------------------------------------------- OpEx Builder layout
# Comparable properties: name column on row 65; each comp is a 7-column block
# (comment, adjustment, unit count, gross $, $/unit, adjusted gross $, adjusted $/unit).
COMPS = ["P", "W", "AD", "AK", "AR", "AY", "BF", "BM", "BT"]
# Comps with an age/size-adjusted $/unit (the ones averaged into the builder).
ADJUSTED = ["P", "W", "AD", "AY", "BF", "BM"]
# Comps with a profile (unit mix, AMI mix, year built) in rows 3-55, at block col + 4.
PROFILED = ["P", "W", "AD", "AY", "BF", "BM"]
# Comps whose block totals are split by bid category (rows 59-63), at block col + 4.
CATEGORISED = ["P", "W", "AD", "AK", "AR", "BT"]
# Section subtotal row -> line-item rows.
SECTIONS = {73: range(74, 88), 88: range(89, 134), 134: range(135, 143), 143: range(144, 161),
            161: range(162, 164), 164: range(165, 168), 168: range(169, 173)}
LINES = range(73, 173)
TARGET_PUPA = "D56"          # $/unit/yr benchmark the comp totals are compared to (row 175)
CATEGORIES = [(60, "Bid"), (61, "PLUG"), (62, "Allowance"), (63, "Alternates")]
AMIS = 6                     # 30/40/50/60/70/80% AMI rows in the comp profiles


@register
class OpExBuilder(Sheet):
    name = "OpEx Builder"

    def compute(self, ctx):
        t = _table()
        const, live = t["constants"], t["live"]
        o: dict = {}

        def v(addr):
            return o[addr] if addr in o else const.get(addr)

        def is_live(col, r):
            return r in live.get(col, ())

        this_year = dt.date.today().year            # Excel: YEAR(TODAY())
        units = ctx.ref("Program", "E64")
        o["E67"] = v("P65")
        o["E68"] = units

        # ---- comp property profiles (unit mix and AMI mix), rows 3-55
        o["BC17"] = 51 + 24 + 1                        # Mesa Commons 2-beds, typed as a sum
        for comp in PROFILED:
            c = _col(comp, 4)
            x = lambda r: xl.n(v(f"{c}{r}"))  # noqa: E731
            o[f"{c}10"] = this_year - x(8)   # age of building
            o[f"{c}19"] = x(15) + x(16) + x(17) + x(18)
            o[f"{c}20"] = x(15) + x(16) + x(17) * 2 + x(18) * 3
        o["BQ38"] = xl.n(v("BQ16")) * (27 / 180)      # Granger: 27 of 180 units at 50% AMI
        o["BQ39"] = xl.n(v("BQ16")) - o["BQ38"]
        o["BQ44"] = xl.n(v("BQ17")) * (27 / o["BQ19"])
        o["BQ45"] = xl.n(v("BQ17")) - o["BQ44"]
        for comp in PROFILED:
            c = _col(comp, 4)
            x = lambda r: xl.n(v(f"{c}{r}"))  # noqa: E731
            for i in range(AMIS):  # AMI totals = studio + 1bd + 2bd + 3bd rows
                o[f"{c}{23 + i}"] = x(30 + i) + x(36 + i) + x(42 + i) + x(48 + i)
            o[f"{c}29"] = sum(x(r) for r in range(23, 29)) + x(54)

        # ---- comp headers: unit counts and scale to this project
        for k, comp in enumerate(COMPS):
            o[f"CN{67 + k}"] = v(f"{comp}65")      # selection list
            if is_live(comp, 68):
                o[f"{comp}68"] = o[f"{_col(comp, 4)}19"]
            if comp in ADJUSTED:
                adj = _col(comp, 6)
                o[f"{adj}68"] = units
                o[f"{adj}69"] = units / xl.n(v(f"{comp}68"))

        # ---- comp line items
        for comp in COMPS:
            self._comp_block(comp, o, v, is_live, units, this_year, t)

        # ---- builder: chosen $/unit per line
        o["G131"] = 9270 / units     # laundry costs, per proforma budget $/yr
        o["G132"] = 14832 / units    # laundry (other expenses), per proforma budget $/yr
        o["G141"] = 3708 / units     # fuel, per proforma budget $/yr
        o["G169"] = ctx.mi["management_fee_pct_of_egi"] * xl.n(ctx.ref("Operating Cash Flow", "J33")) / units
        names = {v(f"{comp}65"): comp for comp in COMPS if ci(comp) >= ci("T")}  # MATCH range T65:CL65
        adj_cols = [_col(c, 6) for c in ADJUSTED]
        for r in list(LINES) + [173, 174]:   # min / avg / max of the adjusted comps
            row_vals = [v(f"{c}{r}") for c in adj_cols]
            if is_live("L", r):
                o[f"L{r}"] = xl.MIN(row_vals)
            if is_live("M", r):
                o[f"M{r}"] = xl.AVERAGE(row_vals)
            if is_live("N", r):
                o[f"N{r}"] = xl.MAX(row_vals)
        for r in LINES:
            if not is_live("H", r):
                continue
            o[f"H{r}"] = units
            if r in SECTIONS:
                continue
            override, pick = v(f"G{r}"), v(f"F{r}")
            if override is not None:
                per_unit = override
            elif xl.eq(pick, "Average"):
                per_unit = o[f"M{r}"]
            elif pick in names:
                # EXCEL-QUIRK: "Hyder 1" (col P) sits outside the T65:CL65 lookup range,
                # so picking it without an override is #N/A in Excel.
                per_unit = xl.n(v(f"{_col(names[pick], 6)}{r}"))
            else:
                raise LookupError(f"#N/A OpEx Builder J{r}: no comp named {pick!r}")
            o[f"J{r}"] = per_unit
            o[f"I{r}"] = per_unit * units
        self._subtotals(["I", "J"], o, v, is_live)

        # ---- bid-category split of the totals (all zero unless lines are tagged)
        for comp in CATEGORISED + ["E"]:
            tag_col = comp
            tot = "J" if comp == "E" else _col(comp, 4)
            for r, cat in CATEGORIES:
                rows = range(73, 176) if cat == "Alternates" else range(71, 173)
                # EXCEL-QUIRK: AH63 sums rows 71-172 for "Alternates"; the other comps use 73-175.
                if tot == "AH" and cat == "Alternates":
                    rows = range(71, 173)
                o[f"{tot}{r}"] = xl.SUMIF([v(f"{tag_col}{k}") for k in rows], f"={cat}",
                                          [v(f"{tot}{k}") for k in rows])
            o[f"{tot}59"] = sum(o[f"{tot}{r}"] for r, _ in CATEGORIES)
        # Payroll taxes as a share of salaries, per comp.
        for comp in ["P", "W", "AD", "AK", "AR"]:
            g = _col(comp, 3)
            o[f"{comp}60"] = xl.n(v(f"{g}82")) / sum(xl.n(v(f"{g}{r}")) for r in range(74, 82))

        o["J175"] = o["J174"] - xl.n(v(TARGET_PUPA))
        for r in LINES:                      # line labels (read by the summary tabs)
            if const.get(f"D{r}") is not None:
                o[f"D{r}"] = const[f"D{r}"]
        return o

    @staticmethod
    def _comp_block(comp, o, v, is_live, units, this_year, t):
        cmt, adj, cnt, gross, pu, agross, apu = (_col(comp, k) for k in range(7))
        comp_units = xl.n(v(f"{comp}68"))
        year_built = xl.n(v(f"{pu}8"))
        scale = v(f"{apu}69")
        for r in LINES:
            if is_live(cnt, r):
                o[f"{cnt}{r}"] = v(f"{comp}68")
            a = f"{adj}{r}"
            if a in t["age_adjust"]:   # older buildings cost more: -1% per year of age
                o[a] = 1 - (this_year - year_built) / 10 * 0.1 + t["age_adjust"][a]
            elif a in t["size_adjust"]:
                o[a] = 1 - (units - comp_units) / 100 * 0.1
            if r in SECTIONS:
                continue
            g = f"{gross}{r}"
            if g in COMP_GROSS:
                o[g] = COMP_GROSS[g](v)
            if is_live(pu, r):
                o[f"{pu}{r}"] = xl.n(v(g)) / xl.n(v(f"{cnt}{r}"))
            ag = f"{agross}{r}"
            if ag in t["passthrough"]:
                o[ag] = xl.n(v(g)) * t["passthrough"][ag]
            elif ag in COMP_ADJ_GROSS:
                o[ag] = COMP_ADJ_GROSS[ag](v)
            elif is_live(agross, r):
                o[ag] = xl.n(v(g)) * xl.n(scale) * xl.n(v(a))
            if is_live(apu, r):
                o[f"{apu}{r}"] = xl.n(v(ag)) / xl.n(v(f"{apu}68"))
        OpExBuilder._subtotals([gross, pu, agross, apu], o, v, is_live)

    @staticmethod
    def _subtotals(cols, o, v, is_live):
        """Section subtotals, grand total (row 174) and variance to the benchmark (row 175)."""
        for c in cols:
            for head, rows in SECTIONS.items():
                if is_live(c, head):
                    o[f"{c}{head}"] = xl.SUM([v(f"{c}{r}") for r in rows])
            if is_live(c, 174):
                o[f"{c}174"] = xl.SUM([v(f"{c}{r}") for r in range(72, 173) if r not in SECTIONS])
            if is_live(c, 175) and c != "J":
                o[f"{c}175"] = (o[f"{c}174"] - xl.n(v(TARGET_PUPA))) / xl.n(v(TARGET_PUPA))


# One-off typed arithmetic in the comp grid (gross $ column).
COMP_GROSS = {
    "S122": lambda v: 48777.56 - xl.n(v("S115")),     # Hyder 1 patrol: total less fire contract
    "S165": lambda v: 100 * xl.n(v("R165")),           # Hyder 1 advertising: $100/unit
    "AN165": lambda v: 25 * xl.n(v("AM165")),          # Hyder 4 advertising: $25/unit
    "AU165": lambda v: 25 * xl.n(v("AT165")),          # Hyder 5 advertising: $25/unit
    "S169": lambda v: 2719578.32 * 0.03,               # mgmt fee: 3% of EGI
    "Z169": lambda v: 1319510 * 0.03,
    "AG169": lambda v: 1850146.93 * 0.03,
    "AN169": lambda v: xl.n(v("AL169")) * xl.n(v("C169")),  # EGI x 3% fee
    "AU169": lambda v: xl.n(v("AS169")) * xl.n(v("C169")),
}
# One-off typed amounts in the "adjusted gross" column (replace the scaled comp figure).
COMP_ADJ_GROSS = {
    "BK135": lambda v: 300 * xl.n(v("BL68")),          # electricity $300/unit
    "BR135": lambda v: 300 * xl.n(v("BS68")),
    "BD136": lambda v: 600 * 250,
    "BD149": lambda v: 3 * 250,
    "BK149": lambda v: 3 * 250,
    "AB163": lambda v: 250 * 1000,                     # insurance
    "BD163": lambda v: 250 * 1000,
    "BK163": lambda v: 1000 * xl.n(v("BL68")),         # insurance $1,000/unit
    "BR163": lambda v: 1000 * xl.n(v("BS68")),
}


# ---------------------------------------------------------------- OpEx & Other Income
SUMMARY_ROWS = {12: 73, 13: 88, 14: 134, 15: 143, 16: 161, 17: 164, 18: 168}  # row -> builder row
# Antenna lease-sale offer tracker (typed on the tab): row -> (monthly rent, cap rate)
ANTENNA_OFFER = {79: (0, 0.058), 80: (0, 0)}           # Dish, Spec.; ATT (row 78) links to MI
ATT_CAP_RATE = 0.06
ANTENNA_LESS = {81: -50000, 82: -6665.51724137925}     # ATT fee, transaction costs
# 2026 vs 2025 income-limit reference table (typed): col -> (2026, 2025)
INCOME_LIMITS = {"E": (122500, 115800), "F": (140000, 132350), "G": (157500, 148900)}


def _yearfrac_30360(start, end) -> float:
    """YEARFRAC(start, end, 0): US (NASD) 30/360."""
    a, b = xl.as_date(start), xl.as_date(end)
    if a > b:
        a, b = b, a
    d1, d2 = a.day, b.day
    last_feb = lambda d: d.month == 2 and (d + dt.timedelta(days=1)).month == 3  # noqa: E731
    if last_feb(a) and last_feb(b):
        d2 = 30
    if last_feb(a):
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30
    if d1 == 31:
        d1 = 30
    return ((b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1)) / 360


@register
class OpExOtherIncome(Sheet):
    name = "OpEx & Other Income"

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        ob = lambda a: ctx.ref("OpEx Builder", a)  # noqa: E731
        units = xl.n(ctx.ref("Program", "E64"))
        o["B2"] = f"{mi['project_name']}  |  OPEX & OTHER INCOME"
        o["I9"] = ob("E68")

        # 1. OpEx budget from the builder, plus contingency.
        for r, b in SUMMARY_ROWS.items():
            o[f"C{r}"], o[f"H{r}"], o[f"I{r}"] = ob(f"D{b}"), ob(f"I{b}"), ob(f"J{b}")
        o["G19"] = mi["contingency"]
        cont = _pct(o["G19"])
        o["H19"] = xl.SUM([o[f"H{r}"] for r in range(12, 19)]) * cont
        o["I19"] = xl.SUM([o[f"I{r}"] for r in range(12, 19)]) * cont
        # EXCEL-QUIRK: H20 is the builder total without contingency, while I20 includes it.
        o["H20"] = ob("I174")
        o["I20"] = xl.SUM([o[f"I{r}"] for r in range(12, 20)])

        # Federal minimum $/unit: region x building type x housing type.
        fm = ctx.grid("Federal Minimum Opex", "B6:G30")
        housing_types = fm[0][2:]
        o["I23"] = sum(xl.n(val)
                       for row in fm[1:]
                       for ht, val in zip(housing_types, row[2:])
                       if xl.eq(mi["building_type_for_federal_minimum_opex_lookup"], row[1])
                       and xl.eq(mi["housing_type"], ht) and xl.eq(mi["opex_region"], row[0]))
        o["H23"] = o["I23"] * units
        o["I24"] = mi["replacement_reserve_per_unit_per_year"]
        o["H24"] = o["I24"] * units
        o["H26"] = mi["tenant_resident_services"]
        o["I26"] = o["H26"] / units
        basis = mi["opex_basis"]
        if xl.eq(basis, "Federal Minimum"):
            o["I27"] = o["I23"]
        elif xl.eq(basis, "Calculated (OpEx Builder)"):
            o["I27"] = o["I20"]
        else:
            o["I27"] = max(o["I20"], o["I23"])
        o["H27"] = o["I27"] * units
        o["I28"] = o["I27"] + o["I24"] + o["I26"]
        o["H28"] = o["I28"] * units

        # 3. Other income.
        o.update({"D38": "Parking", "D39": "AT&T antenna", "D40": "Storage", "D41": "Laundry",
                  "D46": "Vending machines", "D47": "Manager unit(s) RUBS reimbursement",
                  "D48": "Master Service - Internet", "D49": "Annual Interest Income",
                  "D50": "Misc (applications, late, termination)"})
        pg = lambda a: xl.n(ctx.ref("Program", a))  # noqa: E731
        if xl.n(mi["parking_spaces"]) > 0:
            o["E38"] = mi["parking_spaces"]
        else:  # 350 SF per stall across the parking floors
            o["E38"] = xl.ROUNDUP(pg("C22") / 350 + pg("C20") / 350 + pg("C18") / 350 + pg("C16") / 350, 0)
        o["H38"] = o["E38"] / xl.SUM(ctx.col("Program", "E46:E56"))
        if xl.n(mi["storage_units"]) > 0:
            o["E40"] = mi["storage_units"]
        else:  # one storage unit per 200 SF of site
            o["E40"] = xl.ROUNDUP(xl.n(ctx.ref("Initial Screen", "G18")) / 200, 0)
        o["H53"] = xl.SUM(ctx.col("Program", "E46:E50"))

        ums = lambda a: ctx.ref("Unit Mix Summary", a)  # noqa: E731
        sens = lambda a: ctx.ref("Sensitivities", a)  # noqa: E731
        o["F57"], o["G57"] = ums("G46"), sens("C30")                   # master-service internet
        o["F58"], o["G58"] = ums("H39"), mi["laundry_vacancy"]         # common laundry
        o["F61"] = o["E40"] * xl.n(sens("C24")) * 12                    # storage
        o["G61"] = sens("C31")
        o["F62"], o["G62"] = ums("H37"), ums("I37")                     # cell tower
        o["F63"], o["G63"] = ums("H36"), sens("C33")                    # podium parking
        for r in (57, 58, 61, 62, 63):
            o[f"H{r}"] = xl.n(o[f"F{r}"]) * (1 - xl.n(o[f"G{r}"]))
        o["F65"] = xl.SUM([o.get(f"F{r}") for r in range(57, 64)])
        o["H65"] = xl.SUM([o.get(f"H{r}") for r in range(57, 64)])

        # 4. Antenna lease: sale-offer value and rent escalation.
        offers = {78: (mi["antenna_base"], ATT_CAP_RATE), **ANTENNA_OFFER}
        for r, (monthly, cap) in offers.items():
            o[f"E{r}"] = monthly
            o[f"F{r}"] = xl.n(monthly) * 12
            o[f"H{r}"] = xl.IFERROR(lambda: xl.div(o[f"F{r}"], cap), 0)
        for col in "EFH":
            o[f"{col}83"] = xl.SUM([o.get(f"{col}{r}") for r in range(78, 81)]
                                   + ([ANTENNA_LESS[81], ANTENNA_LESS[82]] if col == "H" else []))
        o["G83"] = xl.div(o["F83"], o["H83"])

        base, close = mi["antenna_base"], mi["constr_close"]
        o["E87"] = base
        o["E88"] = xl.EDATE(close, mi["constr_months"])
        o["E89"] = xl.EDATE(close, xl.n(mi["constr_months"]) + xl.n(mi["lease_up_months"]))

        def escalated(on):  # stepped escalation every N years from the base date
            steps = xl.INT(_yearfrac_30360(mi["antenna_base_date"], on) / xl.n(mi["antenna_esc_years"]))
            return xl.n(base) * (1 + xl.n(mi["antenna_esc"])) ** steps

        o["E90"] = escalated(o["E88"])
        o["E91"] = escalated(o["E89"])

        # 5. Reference: 2026 income-limit increase, weighted by unit count.
        for col, (new, old) in INCOME_LIMITS.items():
            o[f"{col}100"] = (new - old) / old
        for col, rb in zip("EFG", ("E53", "E54", "E55")):
            o[f"{col}102"] = ctx.ref("Rent Builder", rb)
        cnt = [xl.n(o[f"{c}102"]) for c in "EFG"]
        o["E104"] = xl.SUMPRODUCT([o[f"{c}100"] for c in "EFG"], cnt) / sum(cnt)
        return o
