"""Small support tabs that the core chain reads from.

Sensitivities  - scenario switch. In Excel a tab of "pov / application / previous"
                 cases feeds a handful of knobs into the model. Here the base case
                 (column D) links to Master Inputs; the alternate "application" case
                 values are kept as constants so the switch still works.
Initial Screen - only the site massing block (envelope -> residential SF) is ported;
                 the rest of the tab is a screening dashboard.
Lists          - two Yes/No flags derived from project type.
"""
from uw import xl
from uw.engine import Sheet, register


@register
class Sensitivities(Sheet):
    name = "Sensitivities"

    # Column E ("application" case) constants, used only when use_base_case is False.
    APPLICATION_CASE = {
        16: 0.0555, 17: 0.0644, 19: 0.82, 20: 0.88, 21: 0.1225, 22: 0.809349157499182,
        24: 40, 25: 60, 28: 0.15, 29: 100, 30: 0.25, 31: 0.1, 32: "no", 33: 0.05, 34: "no",
        36: 0, 37: 250, 39: 0.895, 12: 306, 13: 150, 14: 0,
    }
    E10 = "no"   # Excel: E10 "no" -> D10 "yes" -> use column D (base case)

    def compute(self, ctx):
        mi = ctx.mi
        d = {  # column D: base case, mostly links to Master Inputs
            12: 306, 13: 150, 14: 0,
            15: ctx.ref("OpEx Builder", "J174"),
            16: mi["d1_c_rate"], 17: mi["d1_c_rate_taxable"],
            18: ctx.ref("Financing & Debt Sizing", "E173"),
            19: mi["federal_credit_price"], 20: mi["state_credit_price"],
            21: 0.1225, 22: 0.809349157499182, 23: mi["cuac_reduction_pct"],
            24: mi["storage_rent"], 25: mi["master_service_internet_income"],
            26: mi["constr_months"], 27: mi["lease_up_months"],
            28: 0.15, 29: 100, 30: mi["internet_vacancy"], 31: mi["storage_vacancy"],
            32: mi["prevailing_wage"], 33: mi["parking_vacancy"],
            34: mi["parking_excluded_from_eligible_basis"],
            36: 0, 37: 250, 38: ctx.ref("OpEx & Other Income", "E104"), 39: 0.9395,
        }
        e = dict(self.APPLICATION_CASE)
        for r in (15, 23, 26, 27, 38):          # E = D for these rows
            e[r] = d[r]
        e[18] = ctx.ref("Financing & Debt Sizing", "D173")
        d10 = "yes" if self.E10 == "no" else "no"
        out = {"E10": self.E10, "D10": d10, "C5": ctx.ref("Sources & Uses", "E58")}
        for r in d:
            out[f"D{r}"] = d[r]
            out[f"E{r}"] = e[r]
            out[f"C{r}"] = d[r] if d10 == "yes" else e[r]
        out["C35"] = mi["bond_ratio"]
        return out


@register
class InitialScreen(Sheet):
    name = "Initial Screen"

    # Screening constants typed on the tab.
    HEIGHT_LIMIT_FT = 165        # J28
    GROUND_FLOOR_HT = 12         # J29
    GROUND_FLOOR_HT_ALT = 12     # J30
    PARKING_FLOOR_HT = 10        # J31
    DENSITY_DU_PER_ACRE = 1000   # J32
    FAR = 10                     # J33
    DENSITY_BONUS = 0.8          # P21
    EFFICIENCY_LOSS = 0.1225     # S31 (common-area share)

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        lot = mi["lot_size"]
        o["G18"] = lot
        o["P17"] = lot
        o["P18"] = lot / 43560
        o["P19"] = self.DENSITY_DU_PER_ACRE
        o["P20"] = o["P18"] * o["P19"]
        o["P22"] = xl.ROUNDUP((1 + self.DENSITY_BONUS) * o["P20"], 0)
        o["P24"] = self.FAR
        o["P25"] = o["P24"] * lot
        o["P26"] = max(o["P25"], 0)
        o["P27"] = self.EFFICIENCY_LOSS
        o["P28"] = (1 - o["P27"]) * o["P26"]
        # Average unit SF; note the Excel range includes the "Total Restricted" row 50.
        e = [xl.n(ctx.ref("Program", f"E{r}")) for r in range(46, 51)]
        d = [xl.n(ctx.ref("Program", f"D{r}")) for r in range(46, 51)]
        o["P29"] = xl.div(xl.SUMPRODUCT(e, d), sum(e), 0.0)
        o["P30"] = xl.div(o["P28"], o["P29"], 0.0)
        p22, p30 = o["P22"], o["P30"]
        if p22 > 0 and p30 == 0 or p22 == 0 and p30 > 0:
            o["P32"] = max(p30, p22)
        else:
            o["P32"] = min(p30, p22)
        o["S32"] = mi["lot_coverage_pct"]
        o["S36"] = mi["number_of_parking_floors"]
        o["S38"] = mi["number_of_residential_floors"]
        o["P35"] = min(self.GROUND_FLOOR_HT, self.GROUND_FLOOR_HT_ALT) if self.GROUND_FLOOR_HT_ALT > 0 \
            else self.GROUND_FLOOR_HT
        o["P36"] = o["S36"]
        j29 = self.GROUND_FLOOR_HT
        bump = 34 if j29 == 11 else 35 if j29 == 12 else 36 if j29 >= 13 else 33
        o["P37"] = xl.ROUND(self.HEIGHT_LIMIT_FT - o["P35"]
                            - self.PARKING_FLOOR_HT * (xl.ROUNDUP(o["P36"], 0) - 1) + bump, 0)
        o["P38"] = o["S38"] if xl.n(o["S38"]) > 0 else xl.ROUND(o["P37"] / 10, 0)
        o["P39"] = o["S32"]
        o["P40"] = 0 if o["P32"] == 0 else o["P39"] * lot * o["P38"]
        p38, p36 = o["P38"], o["P36"]
        o["O81"] = (f"{_num(p38 + p36)}-story residential mixed-use buildling "
                    f"({_num(p38)}-stories of residential over {_num(p36)}-stories of parking)")
        return o


def _num(x):
    """Format a number the way Excel's CONCATENATE does (no trailing .0)."""
    return str(int(x)) if float(x).is_integer() else repr(x)


@register
class Lists(Sheet):
    name = "Lists"

    def compute(self, ctx):
        return {
            "Q7": "Yes" if xl.eq(ctx.mi["project_type"], "Acquisition/Rehabilitation") else "No",
            "Q8": "Yes" if xl.eq(ctx.mi["tcac_housing_type"], "Special Needs") else "No",
        }
