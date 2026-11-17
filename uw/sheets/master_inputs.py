"""Master Inputs: the single source of project assumptions.

Every input lives in data/<deal>/inputs.json and is read by key (ctx.mi["credit_type"]).
This module also exposes column E by address so tabs that read
'Master Inputs'!E175 resolve, and computes the two derived cells on the tab.
"""
from uw import xl
from uw.engine import Sheet, register


@register
class MasterInputs(Sheet):
    name = "Master Inputs"

    def compute(self, ctx):
        mi = ctx.mi
        out = {f"E{r['row']}": mi[r["key"]] for r in mi.rows}

        # E291 tax-exempt bond amount: sized off aggregate basis, or typed.
        if mi["tax_exempt_bond_sizing_method"] == "% of aggregate basis":
            aggregate_basis = xl.n(ctx.ref("Financing & Debt Sizing", "D181"))
            out["E291"] = xl.ROUND(mi["bond_ratio"] * aggregate_basis, -3)
        else:
            out["E291"] = mi["tax_exempt_bond_amount_typed"]

        out["E2"] = ctx.ref("Sources & Uses", "E58")       # gap, displayed at top of tab
        out["J230"] = xl.n(mi["d1_p_rate"]) + xl.n(mi["d1_p_spread"])
        return out
