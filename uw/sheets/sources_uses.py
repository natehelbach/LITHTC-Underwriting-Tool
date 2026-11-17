"""Sources & Uses: the capital stack against total development cost.

Uses come straight from Dev Budget subtotals. Construction sources: the
construction loans, recycled bonds, deferred items and other fixed sources are
taken as given and tax-credit equity is the plug. Permanent sources: total LIHTC
equity, perm loan, soft debt and deferred developer fee. E58 = permanent sources
minus total cost, the model's headline surplus / (gap).
"""
from uw import xl
from uw.engine import Sheet, register

DB = "Dev Budget"
FIN = "Financing & Debt Sizing"

# Uses rows: row -> (Dev Budget label cell or literal label, Dev Budget total cell).
USES = {
    12: ("C13", "H19"), 13: ("C21", "H37"), 14: ("C39", "H39"), 15: ("C41", "H41"),
    16: ("C42", "H42"), 17: ("C43", "H43"), 18: ("C54", "H64"), 19: ("C66", "H70"),
    20: (None, "H52"), 21: ("C72", "H97"), 22: ("C100", "H100"), 23: ("C103", "H109"),
    24: ("C112", "H115"),
}


@register
class SourcesUses(Sheet):
    name = "Sources & Uses"

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        units = xl.n(ctx.ref("Program", "E64"))
        four_pct = abs(xl.n(mi["credit_type"]) - 0.04) < 1e-12
        o["B2"] = f"{mi['project_name']}  |  SOURCES & USES"
        o["D7"] = mi["project_name"]

        # ---------------- uses
        for r, (label_cell, cost_cell) in USES.items():
            o[f"D{r}"] = ctx.ref(DB, label_cell) if label_cell else "Legal Costs"
            o[f"E{r}"] = xl.n(ctx.ref(DB, cost_cell))
        uses = o["E26"] = sum(o[f"E{r}"] for r in USES)
        o["D26"] = "Total Uses"
        for r in list(USES) + [26]:
            o[f"F{r}"] = o[f"E{r}"] / uses
            o[f"G{r}"] = o[f"E{r}"] / units
        o["G26"] = sum(o[f"G{r}"] for r in USES)

        # ---------------- permanent-source flags & side items (needed by construction stack)
        o["G53"] = mi["rent_from_previous_commercial_ops_available_during_construct"]
        o["G54"] = mi["cell_tower_noi_available_during_construction"]
        o["G55"] = "No"
        o["G56"] = mi["application_adjustment_factor_available_during_construction"]
        o["H53"] = mi["hcd_0_42pct_debt_service_flag_for_perm_source_row_49"]
        o["H54"] = mi["hcd_0_42pct_debt_service_flag_for_perm_source_row_50"]
        o["H55"] = mi["hcd_0_42pct_debt_service_flag_for_perm_source_row_51"]
        o["H56"] = mi["hcd_0_42pct_debt_service_flag_for_perm_source_row_52"]
        o["K56"] = mi["noi_residual_from_commercial_space_already_received"]
        o["L56"] = ctx.ref("Leasing", "C89")
        e53 = 0.0     # rent from previous commercial operations: no amount typed
        e56 = 0.0     # application adjustment factor: no amount typed
        # Cell tower NOI over construction + lease-up months.
        o["E54"] = xl.n(ctx.ref("OpEx & Other Income", "E90")) * (
            xl.n(ctx.ref("Sensitivities", "C26")) + xl.n(ctx.ref("Sensitivities", "C27")))

        # ---------------- construction sources
        o["D29"] = "Tax Credit Equity"
        o["D30"] = "Tax Exempt Construction Loan" if four_pct else "Construction Loan"
        o["E30"] = xl.n(ctx.ref(FIN, "D180") if four_pct else ctx.ref(FIN, "C195"))
        o["D31"] = "Taxable Construction Loan"
        o["E31"] = xl.n(ctx.ref("Flow of Funds", "D191"))
        o["D32"] = "Recycled Bonds"
        o["E32"] = xl.n(ctx.ref(FIN, "C187"))
        o["D33"] = "Deferred Fees until Conversion (Developer fee reserves, etc)"
        o["E33"] = xl.n(ctx.ref(DB, "T119"))
        o["D34"] = "GP Contributions"
        o["E34"] = 0.0
        o["D35"] = f"Debt 2 construction - {ctx.ref(FIN, 'E12')}"
        o["E35"] = xl.n(ctx.ref(FIN, "E66"))
        o["D36"] = "Deferred Developer Fee"
        o["E36"] = xl.n(ctx.ref(DB, "H114"))
        o["D37"] = "Rent from commercial operations"
        o["E37"] = e53 if xl.eq(o["G53"], "Yes") else 0
        o["E38"] = o["E54"] if xl.eq(o["G54"], "Yes") else 0
        o["D39"] = "GP Loan Note"
        o["E39"] = xl.n(mi["gp_loan_note_construction_source"])
        o["D40"] = "Leasehold"
        o["E40"] = xl.n(mi["leasehold_seller_carryback_note"])
        o["D41"] = "application adjustment factor"
        o["E41"] = e56 if xl.eq(o["G56"], "Yes") else 0
        o["E42"] = 0.0
        # Equity is the plug that balances construction sources to total uses.
        # EXCEL-QUIRK: the plug does not net out E34, E38 or E41, so those would overfund E43.
        o["E29"] = uses - sum(o[f"E{r}"] for r in (30, 31, 32, 33, 35, 36, 37, 39, 40))
        constr = o["E43"] = sum(o[f"E{r}"] for r in range(29, 43))
        o["D43"] = "Total Construction Sources"
        for r in range(29, 43):
            o[f"F{r}"] = o[f"E{r}"] / constr
            if r != 42:
                o[f"G{r}"] = o[f"E{r}"] / units
        o["F43"] = sum(o[f"F{r}"] for r in range(29, 43))
        o["G43"] = sum(o[f"G{r}"] for r in range(29, 42))

        # ---------------- permanent sources
        o["D46"] = "Tax Credit Equity"
        o["E46"] = xl.n(ctx.ref("LIHTC & Dev Fee", "G66"))
        o["D47"] = "Perm Loan"
        o["E47"] = xl.n(ctx.ref(FIN, "D166"))
        o["D48"] = "Leasehold"
        o["E48"] = o["E40"]
        o["D49"] = "GP Loan Note"
        o["E49"] = xl.n(mi["gp_loan_note_permanent_source"])
        o["D50"] = f"Debt 3 - {ctx.ref(FIN, 'F12')}"
        o["E50"] = xl.n(ctx.ref(FIN, "F51"))
        o["D51"] = f"Debt 2 perm - {ctx.ref(FIN, 'E12')}"
        o["E51"] = xl.n(ctx.ref(FIN, "E51"))
        o["D52"] = "Deferred Developer Fee"
        o["E52"] = xl.n(ctx.ref(DB, "H114"))
        o["D53"] = "Rent from previous commercial operations"
        o["E53"] = e53
        o["D54"] = "NOI during construction and lease-up cell tower"
        o["D55"] = "NOI residual from commercial space already received"
        o["E55"] = xl.n(o["K56"])
        o["D56"] = "application adjustment factor"
        o["E56"] = e56
        perm = o["E57"] = sum(o[f"E{r}"] for r in range(46, 57))
        o["D57"] = "Total  Permanent Sources"
        for r in range(46, 57):
            o[f"F{r}"] = o[f"E{r}"] / perm
        o["F57"] = sum(o[f"F{r}"] for r in range(46, 57))

        o["D58"] = "Surplus (Gap)"
        o["E58"] = perm - constr

        # Voluntary exclusions of eligible basis (feed LIHTC).
        o["E62"] = mi["voluntary_exclusion_of_eligible_basis_federal"]
        o["E63"] = mi["voluntary_exclusion_state_credits"]
        return o
