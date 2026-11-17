"""Unit Takeoff and Unit Mix Summary: what gets built, at which AMI, and the rent caps.

Unit Takeoff     - one row per physical unit variant (rows 6-94): bedroom type, SF, and the
                   split of its units across manager / 30-80% AMI. Some variants take their
                   unit count from the architect's plan takeoff (rows 102-117). Rent Builder
                   reads the variant grid.
Unit Mix Summary - the AMI x bedroom allocation (typed), checks against Unit Takeoff and Rent
                   Builder, max rents by AMI (100% AMI rents from LIHTC-TBL Tables) net of
                   utility allowances, commercial / other income inputs, the quick-analysis
                   scenario helper and the TCAC affordability-points calculation.

Typed cells of both tabs live in data/site8/tables/unit_takeoff.json and
unit_mix_summary.json (tools/extract_tables.py).
"""
from __future__ import annotations

import json
from functools import lru_cache

from uw import xl
from uw.engine import DATA, Sheet, register

BEDROOMS = ["Studio", "1 BR", "2 BR", "3 BR", "4 BR"]
AMI_COLS = "OPQRS"                 # Unit Takeoff 30/50/60/70/80% AMI unit columns


@lru_cache(maxsize=None)
def _typed(name: str, deal: str = "site8") -> dict:
    d = json.loads((DATA / deal / "tables" / f"{name}.json").read_text(encoding="utf-8"))
    return d["cells"]


def _num(x) -> str:
    """Number as Excel's & operator renders it (no trailing .0)."""
    x = float(x)
    return str(int(x)) if x.is_integer() else repr(x)


def _crit_num(x) -> float:
    """A number after a round trip through text, as in ">="&L13-0.0001 (15 significant digits)."""
    return float(f"{x:.15g}")


# ======================================================================== Unit Takeoff
@register
class UnitTakeoff(Sheet):
    name = "Unit Takeoff"

    FIRST, LAST = 6, 94            # variant rows
    TOTAL_ROW = 95
    # Variant rows whose total units come from the plan takeoff (by Model ID), not typed.
    PLAN_LINKED_ROWS = set(range(36, 42)) | set(range(66, 69))
    # Row 6 is the one variant whose total is simply manager + restricted.
    SIMPLE_TOTAL_ROWS = {6}
    # Rows that also show unrestricted units (N) and the unallocated remainder (T).
    REMAINDER_ROWS = SIMPLE_TOTAL_ROWS | PLAN_LINKED_ROWS
    # Group checks in column W: restricted units of rows [a..b] equal the target typed in V.
    GROUP_CHECKS = {6: (6, 6), 41: (36, 41), 67: (66, 67)}
    CHECK_RANGE_LAST = 68          # rows 97/98 reconcile only rows 6-68 against Unit Mix Summary
    PLAN_FIRST, PLAN_LAST, PLAN_TOTAL = 102, 117, 118

    def compute(self, ctx):
        typed = _typed("unit_takeoff")
        o = dict(typed)                                       # labels and unit inputs
        val = lambda col, r: typed.get(f"{col}{r}")          # noqa: E731
        o["B2"] = f"{ctx.mi['project_name']}  |  UNIT TAKEOFF"
        o["AA5"] = ctx.ref("Unit Mix Summary", "T26")

        # ---- plan takeoff (architect's floor plans): units per plan type
        lvl5_mult, typ_mult = xl.n(val("G", 101)), xl.n(val("H", 101))   # 1 level-5 floor, 11 typical floors
        plan_rows = range(self.PLAN_FIRST, self.PLAN_LAST + 1)
        for r in plan_rows:
            o[f"I{r}"] = xl.n(val("G", r)) * lvl5_mult + xl.n(val("H", r)) * typ_mult
            o[f"J{r}"] = o[f"I{r}"] * xl.n(val("F", r))                 # total NRSF
        for col in "GHI":
            o[f"{col}{self.PLAN_TOTAL}"] = xl.SUM([o.get(f"{col}{r}") for r in plan_rows])
        plan_model_ids = [val("C", r) for r in plan_rows]
        plan_units = [o[f"I{r}"] for r in plan_rows]

        # ---- variant rows
        rows = range(self.FIRST, self.LAST + 1)
        for r in rows:
            floors = xl.SUM([val(c, r) for c in "IJK"])               # optional floor split
            mgr = xl.n(val("M", r))
            restricted = xl.SUM([val(c, r) for c in AMI_COLS])
            o[f"U{r}"] = restricted
            if r in self.SIMPLE_TOTAL_ROWS:
                total = mgr + restricted
            elif r in self.PLAN_LINKED_ROWS:
                # Criterion is the variant's Model ID; a blank ID behaves like 0 (matches nothing).
                model_id = val("C", r)
                total = xl.SUMIF(plan_model_ids, model_id if model_id is not None else 0.0, plan_units)
            else:
                total = floors if floors > 0 else mgr + restricted
            o[f"L{r}"] = total
            if r in self.REMAINDER_ROWS:
                o[f"N{r}"] = total - mgr
                o[f"T{r}"] = o[f"N{r}"] - restricted
            if total == 0:
                o[f"X{r}"] = ""
            elif mgr + restricted != total:
                o[f"X{r}"] = "AMI split <> total"
            else:
                o[f"X{r}"] = "OK (no floor split)" if floors == 0 else "OK"
        for r, (a, b) in self.GROUP_CHECKS.items():
            o[f"W{r}"] = xl.SUM([o[f"U{k}"] for k in range(a, b + 1)]) == xl.n(val("V", r))

        # ---- totals
        t = self.TOTAL_ROW
        for col in "IJKLMOPQRSU":
            o[f"{col}{t}"] = xl.SUM([o.get(f"{col}{r}") for r in rows])
        o[f"H{t}"] = xl.IFERROR(
            lambda: xl.SUMPRODUCT([o.get(f"H{r}") for r in rows], [o[f"L{r}"] for r in rows]) / o[f"L{t}"], 0)
        rb_units = xl.n(ctx.ref("Rent Builder", "E44"))
        if o[f"L{t}"] == rb_units:
            o[f"X{t}"] = "Ties to Rent Builder"
        else:
            o[f"X{t}"] = f"CHECK: {xl.ROUND(o[f'L{t}'] - rb_units, 0):.0f} units not mapped to a Rent Builder row"

        # ---- reconciliation of AMI columns vs Unit Mix Summary (rows 6-68 only)
        # EXCEL-QUIRK: rows 97/98 sum only rows 6-68 although variants run to row 94.
        check_rows = range(self.FIRST, self.CHECK_RANGE_LAST + 1)
        o["N98"] = xl.SUM([o.get(f"N{r}") for r in check_rows])
        o["O98"] = ctx.ref("Unit Mix Summary", "R14")     # 30% AMI
        # EXCEL-QUIRK: P98 (50% AMI target) is a typed 31, not linked to Unit Mix Summary R16.
        o["Q98"] = ctx.ref("Unit Mix Summary", "R17")     # 60%
        o["R98"] = ctx.ref("Unit Mix Summary", "R18")     # 70%
        o["S98"] = ctx.ref("Unit Mix Summary", "R19")     # 80%
        for col in AMI_COLS:
            o[f"{col}97"] = xl.n(o.get(f"{col}98")) - xl.SUM([o.get(f"{col}{r}") for r in check_rows])

        i_tot = o[f"I{self.PLAN_TOTAL}"]
        o[f"K{self.PLAN_TOTAL}"] = (f"Ties to takeoff above ({_num(i_tot)} units)"
                                    if i_tot == o[f"L{t}"] else "CHECK vs K96")
        return o


# ======================================================================== Unit Mix Summary
@register
class UnitMixSummary(Sheet):
    name = "Unit Mix Summary"

    MIX_ROWS = range(13, 24)       # AMI rows of the allocation grid (AMI in L, units in M:Q)
    MIX_COLS = "MNOPQ"             # Studio..4BR
    # Allocation rows checked against Unit Takeoff (others are checked against zero).
    TAKEOFF_COL_BY_ROW = {14: "O", 16: "P", 17: "Q", 18: "R", 19: "S"}   # 30/50/60/70/80% AMI
    TAKEOFF_TYPES = ["STUDIO", "1 BR", "2 BR", "3 BR", "4 BR"]
    UA_COLS = "EFGHI"              # utility allowance by bedroom (Studio..4BR)
    MAX_RENT_ROWS = range(29, 40)  # max gross rent by AMI (AMI in L)
    NET_RENT_ROWS = range(44, 55)  # max rent net of utility allowance
    SCEN_ROWS = range(32, 43)      # quick-analysis AMI rows
    SCEN_COLS = ["W", "X", "Y", "Z", "AA"]
    SCEN_BLOCKS = [range(67, 78), range(79, 90), range(91, 102), range(103, 114)]
    RB_LINKED_SCEN_ROWS = [*range(68, 74), *range(79, 87), *range(92, 98), *range(104, 110)]
    SCEN_GRID_COLS = ["Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG"]
    TCAC_ROWS = {126: 14, 127: 15, 128: 16, 129: 17, 130: 19, 131: 20}    # TCAC row -> mix row
    TCAC_GRID_COLS = ["AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH"]

    def compute(self, ctx):
        mi = ctx.mi
        typed = _typed("unit_mix_summary")
        o = dict(typed)
        g = lambda a: o.get(a)                                # noqa: E731
        o["B2"] = f"{mi['project_name']}  |  UNIT MIX SUMMARY"
        o["D7"] = mi["project_name"]

        # ---- unit counts and SF by type (from Rent Builder)
        for i, r in enumerate(range(13, 17)):
            o[f"E{r}"] = ctx.ref("Rent Builder", f"D{53 + i}")
            o[f"F{r}"] = ctx.ref("Rent Builder", f"E{53 + i}")
        o["E17"] = xl.n(ctx.ref("Rent Builder", "D64")) / xl.n(ctx.ref("Rent Builder", "E64"))
        o["F17"] = ctx.ref("Rent Builder", "F44")                       # manager units
        o["F18"] = total_units = xl.SUM([o[f"F{r}"] for r in range(13, 18)])
        for r in range(13, 18):
            o[f"G{r}"] = xl.n(o[f"F{r}"]) / total_units
        o["G18"] = xl.SUM([o[f"G{r}"] for r in range(13, 18)])

        # ---- AMI x bedroom allocation grid
        for r in self.MIX_ROWS:
            o[f"L{r}"] = g(f"V{r + 19}")                               # AMI list typed in V32:V42
        grid_total = xl.SUM([g(f"{c}{r}") for r in self.MIX_ROWS for c in self.MIX_COLS])
        for r in self.MIX_ROWS:
            o[f"R{r}"] = xl.SUM([g(f"{c}{r}") for c in self.MIX_COLS])
            o[f"S{r}"] = o[f"R{r}"] / grid_total
        for c in self.MIX_COLS:
            o[f"{c}24"] = xl.SUM([g(f"{c}{r}") for r in self.MIX_ROWS])
        o["R24"] = xl.SUM([o[f"{c}24"] for c in self.MIX_COLS])
        for c in [*self.MIX_COLS, "R"]:
            o[f"{c}25"] = o[f"{c}24"] / grid_total

        # Check vs Unit Takeoff: units by type at this AMI (rows with no takeoff column vs 0).
        ut_type = ctx.col("Unit Takeoff", "B6:B94")
        for r in self.MIX_ROWS:
            col = self.TAKEOFF_COL_BY_ROW.get(r)
            ut_units = ctx.col("Unit Takeoff", f"{col}6:{col}94") if col else None
            diff = 0.0
            for t, c in zip(self.TAKEOFF_TYPES, self.MIX_COLS):
                takeoff = xl.SUMIFS(ut_units, ut_type, t) if col else 0
                diff += abs(takeoff - xl.n(g(f"{c}{r}")))
            o[f"T{r}"] = diff

        # Check vs Rent Builder: units by type at this AMI (AMI matched within +/-0.0001).
        rb_units = ctx.col("Rent Builder", "G12:G39")
        rb_type = ctx.col("Rent Builder", "I12:I39")
        rb_ami = ctx.col("Rent Builder", "J12:J39")
        for r in self.MIX_ROWS:
            ami = xl.n(o[f"L{r}"])
            lo, hi = _crit_num(ami - 0.0001), _crit_num(ami + 0.0001)
            diff = 0.0
            for t, c in zip(BEDROOMS, self.MIX_COLS):
                rb = xl.SUMIFS(rb_units, rb_type, t, rb_ami, f">={lo!r}", rb_ami, f"<={hi!r}")
                diff += abs(rb - xl.n(g(f"{c}{r}")))
            o[f"U{r}"] = diff
        o["T24"] = xl.SUM([o[f"T{r}"] for r in self.MIX_ROWS])
        o["U24"] = xl.SUM([o[f"U{r}"] for r in self.MIX_ROWS])
        o["T25"] = "PASS" if o["T24"] == 0 else "FAIL"
        o["U25"] = "PASS" if o["U24"] == 0 else "FAIL"
        mgr_takeoff = xl.SUM(ctx.col("Unit Takeoff", "M6:M94"))
        o["T26"] = ("PASS" if o["T25"] == "PASS" and o["U25"] == "PASS"
                    and mgr_takeoff == xl.n(ctx.ref("Rent Builder", "F44")) else "FAIL")
        avg_ami = xl.SUMPRODUCT([o[f"L{r}"] for r in self.MIX_ROWS],
                                [o[f"R{r}"] for r in self.MIX_ROWS]) / o["R24"]
        o["P27"] = o["S28"] = avg_ami                                   # unit-weighted average AMI

        # ---- utility allowance (project city, from Utility Allowances)
        ua_src = ["Z", "AA", "AB", "AC", "AD", "AE"]
        for k, r in enumerate(range(24, 31)):
            for dst, src in zip("DEFGHI", ua_src):
                o[f"{dst}{r}"] = ctx.ref("Utility Allowances", f"{src}{8 + k}")
        for c in self.UA_COLS:
            o[f"{c}31"] = xl.SUM([o[f"{c}{r}"] for r in range(24, 31)])

        # ---- max rents: gross = AMI% x 100% AMI rent (rounded down); net of utility allowance
        rent_100 = [ctx.ref("LIHTC-TBL Tables", f"AJ{7 + i}") for i in range(5)]
        for r in self.MAX_RENT_ROWS:
            for c, base in zip(self.MIX_COLS, rent_100):
                o[f"{c}{r}"] = xl.ROUNDDOWN(xl.n(g(f"L{r}")) * xl.n(base), 0)
        for r in self.NET_RENT_ROWS:
            for c, ua in zip(self.MIX_COLS, self.UA_COLS):
                o[f"{c}{r}"] = o[f"{c}{r - 15}"] - o[f"{ua}31"]

        # ---- commercial income (parking, antenna, storage, laundry)
        for r in range(36, 40):
            o[f"D{r}"] = ctx.ref("OpEx & Other Income", f"D{r + 2}")
        o["E36"] = xl.n(ctx.ref("Program", "C20")) + xl.n(ctx.ref("Program", "C22"))   # parking SF
        o["F36"] = ctx.ref("OpEx & Other Income", "E38")                 # parking spaces
        o["F37"] = mi["antenna_leases"]
        o["F38"] = ctx.ref("OpEx & Other Income", "E40")                 # storage units
        o["F39"] = total_units                                            # laundry: every unit
        o["G36"] = mi["parking_rent"]
        o["G37"] = ctx.ref("OpEx & Other Income", "E91")
        o["G38"] = ctx.ref("Sensitivities", "C24")                       # storage rent
        o["G39"] = mi["laundry_income"]
        o["I36"] = ctx.ref("Sensitivities", "C33")                       # parking vacancy
        o["I37"] = mi["antenna_vacancy"]
        o["I38"] = ctx.ref("Sensitivities", "C31")                       # storage vacancy
        o["I39"] = mi["laundry_vacancy"]
        for r in range(36, 40):
            o[f"H{r}"] = xl.n(o[f"G{r}"]) * xl.n(o[f"F{r}"]) * 12
        o["E40"] = xl.SUM([g(f"E{r}") for r in range(36, 40)])
        o["H40"] = xl.SUM([o[f"H{r}"] for r in range(36, 40)])

        # ---- other income (per unit per month) and vacancy
        for r in range(44, 49):
            o[f"D{r}"] = ctx.ref("OpEx & Other Income", f"D{r + 2}")
        o["F44"], o["H44"] = mi["vending_machines"], mi["vending_machines_vacancy"]
        o["F45"], o["H45"] = mi["manager_unit_rubs_reimbursement"], mi["manager_unit_rubs_vacancy"]
        o["F46"], o["H46"] = ctx.ref("Sensitivities", "C25"), ctx.ref("Sensitivities", "C30")  # internet
        o["F47"], o["H47"] = mi["annual_interest_income"], mi["interest_income_vacancy"]
        o["F48"], o["H48"] = mi["misc_income_applications_late_termination"], mi["misc_income_vacancy"]
        for r in range(44, 49):
            o[f"G{r}"] = xl.n(o[f"F{r}"]) * 12 * total_units
        o["F49"] = xl.SUM([o[f"F{r}"] for r in range(44, 49)])
        o["G49"] = xl.SUM([o[f"G{r}"] for r in range(44, 49)])
        o["Z49"] = xl.n(g("Z48")) * total_units                          # units with vouchers

        # ---- quick analysis: units = bedroom target % x (units less manager) x AMI target %
        o["X44"] = 1 - xl.SUM(g("W44"), [g(f"{c}44") for c in ["Y", "Z", "AA"]])
        o["AB44"] = xl.SUM([o.get(f"{c}44") for c in self.SCEN_COLS])
        o["AC36"] = 1 - xl.SUM([g(f"AC{r}") for r in self.SCEN_ROWS if r != 36])
        units_ex_mgr = xl.n(ctx.ref("Program", "E64"))
        if xl.eq(mi["manager_unit"], "Yes"):
            units_ex_mgr -= 1
        for r in self.SCEN_ROWS:
            for c in self.SCEN_COLS:
                o[f"{c}{r}"] = xl.n(o.get(f"{c}44")) * units_ex_mgr * xl.n(g(f"AC{r}"))
            o[f"AB{r}"] = xl.SUM([o[f"{c}{r}"] for c in self.SCEN_COLS])
        for c in [*self.SCEN_COLS, "AB", "AC"]:
            o[f"{c}43"] = xl.SUM([o.get(f"{c}{r}") for r in self.SCEN_ROWS])

        # ---- scenarios A-D (units / vouchers by AMI row)
        for r in range(66, 102):
            if r not in (78, 90):                                          # typed type headers
                o[f"V{r}"] = o[f"V{r - 1}"]
        for r in self.RB_LINKED_SCEN_ROWS:
            # EXCEL-QUIRK: W and Z read D/G of the same row, which are blank on this tab,
            # so these scenario rows are always 0.
            o[f"W{r}"] = xl.n(g(f"G{r}"))
            crit_type = g(f"D{r}") if g(f"D{r}") is not None else 0.0
            crit_ami = g(f"G{r}") if g(f"G{r}") is not None else 0.0
            o[f"Z{r}"] = xl.SUMIFS(rb_units, rb_type, crit_type, rb_ami, crit_ami)
        hdr_kind = [g(f"{c}66") for c in self.SCEN_GRID_COLS]           # Units / Vouchers
        hdr_scen = [g(f"{c}65") for c in self.SCEN_GRID_COLS]           # A..D
        scenario = g("X65")                                              # selected scenario
        for block in self.SCEN_BLOCKS:
            for r in block:
                vals = [o.get(f"{c}{r}") for c in self.SCEN_GRID_COLS]
                o[f"X{r}"] = xl.SUMIFS(vals, hdr_kind, g("X66"), hdr_scen, scenario)
                o[f"Y{r}"] = xl.SUMIFS(vals, hdr_kind, g("Y66"), hdr_scen, scenario)
        all_rows = range(67, 114)
        for c in ["X", "Y", *self.SCEN_GRID_COLS]:
            o[f"{c}114"] = xl.SUM([o.get(f"{c}{r}") for r in all_rows])

        # ---- AJ:AP quick mix (typed) totals
        quick_total = xl.SUM([g(f"{c}{r}") for r in range(13, 18) for c in ["AJ", "AK", "AL", "AM", "AN"]])
        for r in range(13, 16):
            o[f"AO{r}"] = xl.SUM([g(f"{c}{r}") for c in ["AJ", "AK", "AL", "AM", "AN"]])
            o[f"AP{r}"] = o[f"AO{r}"] / quick_total

        # ---- TCAC affordability points
        rural = xl.eq(mi["rural_set_aside"], "Yes")
        rural_pts = {128: 2.5, 129: 3.8, 130: 5, 131: 6.3, 132: 7.5, 133: 8.8}
        for r in range(128, 143):
            o[f"AB{r}"] = rural_pts.get(r, 10) if rural else 0           # 55% AMI column (rural only)
        o["AC135"] = 22.5 if rural else 20
        o["AC136"] = 25 if rural else 20
        rb_total = xl.n(ctx.ref("Rent Builder", "G44"))
        pct_keys = [g(f"AA{r}") for r in range(126, 143)]
        ami_hdr = [g(f"{c}125") for c in self.TCAC_GRID_COLS]
        for r, mix_row in self.TCAC_ROWS.items():
            o[f"W{r}"] = o[f"R{mix_row}"] / rb_total
            o[f"X{r}"] = xl.IFERROR(lambda r=r: self._points(o, g(f"V{r}"), o[f"W{r}"], pct_keys, ami_hdr), 0)
        o["X132"] = xl.SUM([o[f"X{r}"] for r in self.TCAC_ROWS])
        o["X134"] = min(xl.n(g("X133")), o["X132"])
        o["X136"] = 2 if xl.n(g("V126")) <= 0.3 and o["W126"] >= 0.1 else 0
        o["X137"] = o["X136"] + o["X134"]
        return o

    def _points(self, o, ami, share, pct_keys, ami_hdr):
        """VLOOKUP(share, AA126:AH142, MATCH(ami, AA125:AH125, 0), TRUE)."""
        col = self.TCAC_GRID_COLS[xl.MATCH(ami, ami_hdr, 0) - 1]
        # Approximate match: last numeric '% of units' key <= share (header text is skipped).
        hit = None
        for k, r in zip(pct_keys, range(126, 143)):
            if xl.is_number(k):
                if k <= share:
                    hit = r
                else:
                    break
        if hit is None:
            raise LookupError("#N/A")
        return o.get(f"{col}{hit}")
