"""Returns & Summary: export-ready one-page project summary.

Pure presentation: every value is a link to another tab (Master Inputs, Unit Takeoff,
Sources & Uses, Operating Cash Flow, Financing, LIHTC). Nothing else reads this tab.
"""
from uw import xl
from uw.engine import Sheet, register

UNIT_TYPES = (("C", "STUDIO"), ("D", "1 BR"), ("E", "2 BR"))
# AMI band rows -> Unit Takeoff count column
AMI_ROWS = ((22, "O"), (23, "P"), (24, "Q"), (25, "R"), (26, "S"), (27, "M"))


@register
class ReturnsSummary(Sheet):
    name = "Returns & Summary"

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        su = lambda a: ctx.ref("Sources & Uses", a)
        ocf = lambda a: ctx.ref("Operating Cash Flow", a)
        fin = lambda a: ctx.ref("Financing & Debt Sizing", a)
        lihtc = lambda a: ctx.ref("LIHTC & Dev Fee", a)

        # ---- project overview ----
        o["B2"] = f"{mi['project_name']} - Project Summary"
        o["C5"] = mi["project_name"]
        o["C6"] = f"{mi['city']}, {mi['county']}"
        o["C7"] = mi["tcac_geographic_area"]
        o["C8"] = mi["credit_type"]
        o["C9"] = mi["project_type"]
        o["C10"] = mi["tcac_housing_type"]
        o["C11"] = ctx.ref("Program", "C9")
        o["C12"] = ctx.ref("Program", "C31")
        o["C14"] = ctx.ref("Unit Mix Summary", "F36")
        o["C15"] = mi["constr_close"]
        o["C16"] = ctx.ref("Sensitivities", "C26")
        o["C17"] = mi["lease_up_months"]
        o["C18"] = ctx.ref("Leasing", "D114")

        # ---- unit mix by AMI band and bedroom type (from Unit Takeoff) ----
        ut = lambda c: ctx.col("Unit Takeoff", f"{c}6:{c}94")
        types = ut("B")
        for row, cnt_col in AMI_ROWS:
            counts = ut(cnt_col)
            for c, t in UNIT_TYPES:
                o[f"{c}{row}"] = xl.SUMIF(types, t, counts)
            o[f"F{row}"] = xl.SUM(counts)
        for c in "CDEF":
            o[f"{c}28"] = sum(o[f"{c}{r}"] for r in range(22, 28))
        o["C13"] = o["F28"]
        o["G28"] = ctx.ref("Unit Mix Summary", "T26")
        sf, units = ut("H"), ut("L")
        for c, t in UNIT_TYPES:
            is_t = [1 if xl.eq(b, t) else 0 for b in types]
            o[f"{c}29"] = xl.IFERROR(
                lambda: xl.SUMPRODUCT(is_t, sf, units) / xl.SUMIF(types, t, units), 0)
        o["F29"] = ctx.ref("Unit Takeoff", "H95")
        o["C30"] = ctx.ref("Rent Builder", "F57")

        # ---- development budget (Sources & Uses rows 12-24) ----
        for i in range(13):
            for c_out, c_src in zip("BCDE", "DEFG"):
                o[f"{c_out}{34 + i}"] = su(f"{c_src}{12 + i}")
        for c_out, c_src in zip("CDE", "EFG"):
            o[f"{c_out}47"] = su(f"{c_src}26")
        o["C48"] = ctx.ref("Dev Budget", "E178")

        # ---- sources: construction (S&U 29-41, skipping 34 and 38) and permanent (46-56) ----
        constr_rows = (29, 30, 31, 32, 33, 35, 36, 37, 39, 40, 41)
        for i, src in enumerate(constr_rows):
            for c_out, c_src in zip("BCD", "DEF"):
                o[f"{c_out}{52 + i}"] = su(f"{c_src}{src}")
        o["C63"], o["D63"] = su("E43"), su("F43")
        for i, src in enumerate(range(46, 57)):
            for c_out, c_src in zip("BCD", "DEF"):
                o[f"{c_out}{65 + i}"] = su(f"{c_src}{src}")
        o["C76"], o["D76"] = su("E57"), su("F57")
        o["C78"] = su("E58")

        # ---- operations: year 1 / 5 / 15 ----
        for c_out, yc in (("C", "J"), ("D", "N"), ("E", "X")):
            v = lambda r: ocf(f"{yc}{r}")
            o[f"{c_out}82"] = v(13)
            o[f"{c_out}83"] = xl.n(v(21)) - xl.n(v(13))
            o[f"{c_out}84"] = v(21)
            o[f"{c_out}85"] = v(31)
            o[f"{c_out}86"] = v(33)
            o[f"{c_out}87"] = v(45)
            o[f"{c_out}88"] = v(47)
            o[f"{c_out}89"] = xl.n(v(55)) - xl.n(v(47))
            o[f"{c_out}90"] = v(62)
            o[f"{c_out}91"] = v(64)
            o[f"{c_out}92"] = v(57)

        # ---- financing: Debt 1 / 2 / 3 ----
        for c_out, c_src in zip("CDE", "DEF"):
            f = lambda r: fin(f"{c_src}{r}")
            o[f"{c_out}96"] = f(12)
            o[f"{c_out}97"] = f(13)
            o[f"{c_out}98"] = f(51)
            o[f"{c_out}99"] = f(33)
            o[f"{c_out}100"] = f"{_num(f(38))} / {_num(f(39))}"
            o[f"{c_out}101"] = f(55)
            o[f"{c_out}102"] = f(66)
        o["C103"], o["D103"] = fin("D72"), fin("E72")
        o["C104"], o["D104"] = fin("D75"), fin("D76")

        # ---- LIHTC & developer fee ----
        o["C107"] = lihtc("G42")
        o["C108"] = lihtc("G52")
        o["C109"] = lihtc("G56")
        o["C110"] = lihtc("F58")
        o["C111"], o["D111"] = lihtc("G58"), lihtc("G64")
        o["C112"] = lihtc("G66")
        o["C113"], o["D113"], o["E113"] = lihtc("E76"), lihtc("E77"), lihtc("E78")

        # ---- key metrics & returns ----
        o["C116"] = su("G26")
        o["C117"], o["D117"] = ocf("E138"), ocf("E139")
        o["C118"] = mi["cdlac_points"]
        o["D118"] = ctx.ref("CDLAC Score and Tiebreaker", "E145")
        o["C119"] = ocf("E103")
        o["C121"], o["D121"] = ocf("F112"), ocf("F118")
        o["C122"], o["D122"] = ocf("F115"), ocf("F121")
        o["C123"] = _iferror_cell(ocf("F125"))
        o["D123"] = _iferror_cell(ocf("F131"))
        o["C124"], o["D124"] = ocf("F127"), ocf("F133")
        return o


def _num(x):
    """Format a number the way Excel's & operator does (no trailing .0)."""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x)


def _iferror_cell(v):
    """IFERROR(link, "n/a"): an error value in the source cell shows as n/a."""
    if v is None:
        return 0
    return "n/a" if isinstance(v, str) and v.startswith("#") else v
