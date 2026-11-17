"""Operating Cash Flow: 60-year annual pro forma and the investor returns block.

Columns J..BQ are operating years 1..60. Rows:
  13-21   potential gross income (rent, subsidy, other income lines)
  24-31   vacancy, 33 EGI
  36-45   operating expenses (line items only when OpEx tab is in itemized mode), 47 NOI
  49-55   reserves / taxes / monitoring / trustee / services -> income before debt service
  57-64   debt service by instrument (from Financing), DSCR, net cash flow
  67-86   fees, deferred developer fee paydown, cash available for distribution
Investor block (rows 90-134, columns G..X = 18 periods): cash flow summary, NPV,
seller carryback note, profit share, 15-yr and 5-yr IRR / NPV / multiple, oWOW returns.
"""
from openpyxl.utils import column_index_from_string as col_index, get_column_letter as col_letter

from uw import xl
from uw.engine import Sheet, register

YEARS = 60


def yc(year):
    """Pro forma column for operating year `year` (J = year 1)."""
    return col_letter(9 + year)


def fin_col(year):
    """Financing & Debt Sizing schedule column (I = year 1)."""
    return col_letter(8 + year)


INV_COLS = [col_letter(c) for c in range(col_index("G"), col_index("X") + 1)]  # 18 periods

LEASE_UP_INCOME_SHARE = 0.75        # I96: share of lease-up income paid to the investor
OWOW_YEAR6_DEDUCTION = 85000        # L125: one-off deduction in oWOW year-6 cash flow
SELLER_CARRYBACK_NOTE = 2000000     # F106: SD $2.0M seller carryback note (per source model)


@register
class OperatingCashFlow(Sheet):
    name = "Operating Cash Flow"

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        ums = lambda a: ctx.ref("Unit Mix Summary", a)
        oe = lambda a: ctx.ref("OpEx & Other Income", a)
        fin = lambda a: ctx.ref("Financing & Debt Sizing", a)
        su = lambda a: ctx.ref("Sources & Uses", a)
        yrs = range(1, YEARS + 1)

        o["B2"] = f"{mi['project_name']}  |  OPERATING CASH FLOW"
        o["E7"] = mi["project_name"]
        for y in yrs:
            o[f"{yc(y)}12"] = y

        def grow(row, first, rate):
            """Year-1 value escalated at `rate` per year across the 60 columns."""
            o[f"J{row}"] = xl.n(first)          # a link to a blank cell reads as 0
            for y in range(2, YEARS + 1):
                o[f"{yc(y)}{row}"] = o[f"{yc(y - 1)}{row}"] * (1 + rate)

        # ---- potential gross income ------------------------------------------------
        rent_esc = mi["residential_rent_escalation"]
        o["G13"] = rent_esc
        grow(13, ctx.ref("Rent Builder", "D48"), rent_esc)
        grow(14, mi["rent_subsidy"], rent_esc)
        # other income lines from Unit Mix Summary: (row, label/amount row, amount col)
        other = [(15, 36, "H"), (16, 37, "H"), (17, 38, "H"), (18, 39, "H"), (19, 46, "G")]
        for row, src, amt_col in other:
            o[f"E{row}"] = ums(f"D{src}")
            o[f"G{row}"] = mi["commercial_rent_escalation"] if row == 15 else rent_esc
            grow(row, ums(f"{amt_col}{src}"), o[f"G{row}"])
        o["G20"] = rent_esc                     # "Other Income" row is empty
        for y in yrs:
            x = yc(y)
            o[f"{x}21"] = xl.SUM([o.get(f"{x}{r}") for r in range(13, 21)])

        # ---- vacancy ----------------------------------------------------------------
        o["G24"] = mi["residential_vacancy"]
        vac_rates = {25: ums("I36"), 26: ums("I37"), 27: ums("I38"), 28: ums("I39"), 29: ums("H46")}
        for row, src in zip(range(25, 30), (36, 37, 38, 39, 46)):
            o[f"E{row}"] = ums(f"D{src}")
            o[f"G{row}"] = vac_rates[row]
        o["G30"] = mi["commercial_other_income_vacancy"]
        for y in yrs:
            x = yc(y)
            o[f"{x}24"] = o["G24"] * -o[f"{x}13"]
            for row in range(25, 30):
                o[f"{x}{row}"] = xl.n(o[f"G{row}"]) * -o[f"{x}{row - 10}"]
            o[f"{x}30"] = o["G30"] * -xl.n(o.get(f"{x}20"))
            o[f"{x}31"] = sum(o[f"{x}{r}"] for r in range(24, 31))
            o[f"{x}33"] = o[f"{x}21"] + o[f"{x}31"]

        # ---- operating expenses -----------------------------------------------------
        # Itemized lines are used only when the OpEx tab's mode cell I27 equals I20;
        # otherwise the total (H27) is escalated as one line.
        opex_esc = mi["operating_expense_escalation"]
        o["G35"] = opex_esc
        o["G44"] = mi["real_estate_tax_escalation"]
        itemized = xl.IFERROR(lambda: oe("I27") == oe("I20"), False)
        for i, row in enumerate(range(36, 43)):
            o[f"E{row}"] = oe(f"C{12 + i}")
            o[f"J{row}"] = xl.IFERROR(lambda: -xl.n(oe(f"H{12 + i}")) if itemized else 0, 0)
        o["J43"] = 0                         # resident services are charged below NOI (row 53)
        o["J44"] = 0                         # real estate taxes are charged below NOI (row 50)
        o["J45"] = xl.IFERROR(lambda: xl.SUM([o[f"J{r}"] for r in range(36, 45)]) if itemized
                              else -xl.n(oe("H27")), 0)
        for y in range(2, YEARS + 1):
            x, p = yc(y), yc(y - 1)
            for row in range(36, 45):
                esc = o["G44"] if row == 44 else opex_esc
                o[f"{x}{row}"] = o[f"{p}{row}"] * (1 + esc) if itemized else 0
            o[f"{x}45"] = (xl.SUM([o[f"{x}{r}"] for r in range(36, 45)]) if itemized
                           else o[f"{p}45"] * (1 + opex_esc))
        for y in yrs:
            x = yc(y)
            o[f"{x}47"] = o[f"{x}33"] + o[f"{x}45"]

        # ---- below-NOI items -> income before debt service --------------------------
        o["G49"] = mi["replacement_reserve_escalation"]
        grow(49, oe("H24"), o["G49"])
        o["G50"] = mi["real_estate_tax_escalation"]
        grow(50, mi["real_estate_taxes"], o["G50"])
        # bond monitoring fee steps down with the Debt 1 balance
        o["J51"] = xl.n(fin("O144"))
        for y in range(2, YEARS + 1):
            prev = o[f"{yc(y - 1)}51"]
            o[f"{yc(y)}51"] = xl.IFERROR(
                lambda: prev * xl.n(fin(f"{fin_col(y)}101")) / xl.n(fin(f"{fin_col(y - 1)}101")), 0)
        grow(52, mi["trustee_fees"], 0)
        o["G53"] = mi["service_amenities_escalation"]
        grow(53, mi["tenant_resident_services"], o["G53"])
        for y in yrs:
            x = yc(y)
            o[f"{x}55"] = o[f"{x}47"] - sum(o[f"{x}{r}"] for r in range(49, 54))

        # ---- debt service --------------------------------------------------------------
        o["E58"] = "Debt 1 - " + str(fin("D12"))
        o["E59"] = "Debt 2 - " + str(fin("E12")) + " (incl. monitoring fee)"
        o["E60"] = "Debt 3 - " + str(fin("F12"))
        for y in yrs:
            x, f = yc(y), fin_col(y)
            o[f"{x}58"] = -xl.n(fin(f"{f}102"))
            o[f"{x}59"] = -(xl.n(fin(f"{f}114")) + xl.n(fin(f"{f}119")))
            o[f"{x}60"] = -xl.n(fin(f"{f}127"))
            o[f"{x}62"] = o[f"{x}58"] + o[f"{x}59"] + o[f"{x}60"]
            o[f"{x}57"] = xl.IFERROR(lambda: o[f"{x}55"] / -o[f"{x}62"], "n/a")
            o[f"{x}64"] = o[f"{x}55"] + o[f"{x}62"]

        # ---- fees, deferred developer fee, distributions ----------------------------
        o["G67"] = mi["investor_asset_management_fee_escalation"]
        grow(67, -mi["investor_asset_management_fee"], o["G67"])
        o["G71"] = mi["partnership_management_fee_escalation"]
        grow(71, -mi["partnership_management_fee"], o["G71"])
        o["H77"] = 1
        o["J76"] = xl.n(su("E52"))
        lease_up_to_investor = xl.n(ctx.ref("Leasing", "C89")) * LEASE_UP_INCOME_SHARE
        o["J78"] = lease_up_to_investor
        o["G81"] = 1.0                          # GP/LP share = 1 - SUM(G82:G85), all empty
        o["G86"] = o["G81"]
        for y in yrs:
            x = yc(y)
            ncf = o[f"{x}64"]
            # asset management fee paid to the extent cash flow covers it
            o[f"{x}69"] = o[f"{x}67"] if ncf > -o[f"{x}67"] else ncf
            # partnership management fee
            if y == 1:
                # EXCEL-QUIRK: year 1 pays the full fee (J72+J71) without the MIN cap later years use.
                o[f"{x}73"] = o[f"{x}71"] if ncf + o[f"{x}69"] > 0 else 0
            elif y <= 58:
                # EXCEL-QUIRK: IF has no else branch, so it returns FALSE when cash is short;
                # the formula also stops at BO (years 59-60 are blank).
                o[f"{x}73"] = (min(o[f"{x}71"], ncf + o[f"{x}69"]) if ncf + o[f"{x}69"] > 0
                               else False)
            pm = xl.n(o.get(f"{x}73"))
            if y > 1:
                p = yc(y - 1)
                left = o[f"{p}76"] - o[f"{p}77"] - xl.n(o.get(f"{p}78"))
                o[f"{x}76"] = left if left > 0 else 0
            avail = ncf + o[f"{x}69"] + pm
            o[f"{x}77"] = avail if o[f"{x}76"] > avail else o[f"{x}76"]
            o[f"{x}79"] = avail - o[f"{x}77"]
            o[f"{x}81"] = o["G81"] * o[f"{x}79"]
            o[f"{x}86"] = o[f"{x}81"]           # row 83 is empty

        self._investor_block(ctx, o)

        # ---- project yield ------------------------------------------------------------
        cost_ex_deferred = xl.n(su("E43")) - xl.n(su("E36")) - xl.n(su("E33"))
        # (div default 0 only guards the model's cold start, before Sources & Uses has run)
        o["E138"] = xl.div(o["J47"], cost_ex_deferred, 0.0)
        o["E139"] = xl.div(o["J47"], cost_ex_deferred - xl.n(su("E46")), 0.0)
        return o

    def _investor_block(self, ctx, o):
        mi = ctx.mi
        su = lambda a: ctx.ref("Sources & Uses", a)
        cols = INV_COLS
        rate = mi["discount_rate_owow_wacc"]
        o["D103"] = rate

        # cash flow summary rows 92-100 (all zero unless set)
        for r in range(92, 101):
            for c in cols:
                o[f"{c}{r}"] = 0
        o["G92"] = mi["equity_investment"]
        o["H93"] = -o["G92"]
        o["H94"] = xl.n(ctx.ref("Dev Budget", "H18"))
        lihtc = lambda a: ctx.ref("LIHTC & Dev Fee", a)
        o["H95"], o["J95"], o["K95"] = (xl.n(lihtc(a)) for a in ("L176", "L177", "L178"))
        o["I96"] = o["J78"]
        for c in cols[cols.index("J"):]:
            o[f"{c}97"] = o[f"{c}77"]
        o["X98"] = xl.n(su("E23"))
        o["B99"] = mi["residual_cash_flow_share_to_investor"]
        o["J99"] = o["J86"] + xl.n(su("E58"))
        for c in "KLMN":
            o[f"{c}99"] = o[f"{c}86"]
        for c in cols[cols.index("O"):]:
            o[f"{c}99"] = o[f"{c}86"] * o["B99"]
        # sale at the end of year 15 on year-16 income
        # EXCEL-QUIRK: sale NOI = income before DS (Y55) plus monitoring, trustee and services
        # (Y51:Y53) but not reserves / property tax; X85 (empty row) is subtracted.
        o["B100"] = mi["net_sale_proceeds_factor"]
        sale_income = sum(o[f"Y{r}"] for r in (55, 53, 52, 51))
        o["X100"] = ((sale_income / mi["exit_cap_rate"]
                      - xl.n(ctx.ref("Financing & Debt Sizing", "D167")) - 0) * o["B100"])  # X85 empty

        for r in range(92, 101):
            row = [o[f"{c}{r}"] for c in cols]
            o[f"E{r}"] = xl.NPV(rate, row)
            o[f"F{r}"] = xl.SUM(row)
        for c in ["E", "F"] + cols:
            o[f"{c}101"] = xl.SUM([o[f"{c}{r}"] for r in range(92, 101)])
        for r in range(94, 101):
            o[f"D{r}"] = xl.div(o[f"F{r}"], o["F101"], 0.0)   # default: cold-start guard
        o["E103"] = xl.NPV(rate, [o[f"{c}101"] for c in cols])
        o["F103"] = ("no error" if xl.ROUNDUP(o["E103"], 0) == xl.ROUNDUP(o["E101"], 0)
                     else "ERROR")

        # seller carryback note: accrues interest, paid down from lease-up income,
        # half of the deferred fee and (from year 11) half of the residual cash flow
        o["E105"] = "Seller Carryback Deduction (SD: $2.0M note at 8%, per source model)"
        o["F106"] = SELLER_CARRYBACK_NOTE
        o["D107"] = mi["seller_carryback_interest_rate"]
        o["G106"] = o["F106"]
        o["G107"] = 0
        for i, c in enumerate(cols):
            if i > 0:
                p = cols[i - 1]
                o[f"{c}106"] = o[f"{p}108"]
                o[f"{c}107"] = o[f"{p}108"] * o["D107"]
            beg, intr = o[f"{c}106"], o[f"{c}107"]
            if c in "GH":
                o[f"{c}108"] = intr + beg
            elif c == "I":
                o[f"{c}108"] = intr + beg - o["I96"]
            elif c in "JKLMNOP":
                o[f"{c}108"] = intr + beg - o[f"{c}97"] / 2
            elif c != "X":
                o[f"{c}108"] = intr + beg - o[f"{c}97"] / 2 - o[f"{c}99"] / 2
            else:
                o[f"{c}108"] = intr + beg - beg - intr      # note repaid at sale
        o["F109"] = mi["profit_share_pct"]
        # EXCEL-QUIRK: years J..P exclude lease-up income (row 96) and the carryback balance,
        # Q..X include row 96 and subtract the carryback balance (row 108).
        for c in cols[cols.index("J"):]:
            if c in "JKLMNOP":
                o[f"{c}109"] = (o[f"{c}100"] + o[f"{c}99"] + o[f"{c}98"] + o[f"{c}97"]) * o["F109"]
            else:
                o[f"{c}109"] = (o[f"{c}96"] - o[f"{c}108"] + o[f"{c}97"] + o[f"{c}99"]
                                + o[f"{c}98"] + o[f"{c}100"]) * o["F109"]

        # ---- 15-year investor returns (after carryback) -------------------------------
        o["G112"] = mi["equity_investment_in_15_yr_irr_row"]
        o["H112"] = o["H101"]
        o["I112"] = 0
        for c in "JKLMNOP":
            o[f"{c}112"] = o[f"{c}101"] - o[f"{c}97"] / 2
        for c in "QRSTUVW":
            o[f"{c}112"] = o[f"{c}101"] - o[f"{c}97"] / 2 - o[f"{c}99"] / 2
        o["X112"] = o["X100"] + o["X98"] + o["X99"] - o["X107"] - o["X106"]
        self._returns(o, 112, rate, o["F92"])

        # 5-year: years 1-5 then the NPV of the remaining cash flows at a higher rate
        o["D120"] = mi["discount_rate_for_5_yr_npv_block"]
        for c in "GHIJK":
            o[f"{c}118"] = o[f"{c}112"]
        o["L118"] = xl.NPV(o["D120"], [o[f"{c}101"] for c in cols[cols.index("L"):]])
        self._returns(o, 118, rate, o["F92"])

        # ---- oWOW returns -------------------------------------------------------------
        o["G125"] = mi["owow_equity"]
        o["H125"] = o["H112"] * 0.5 + (o["H96"] + o["H97"] + o["H99"]) / 2
        for c in "IJK":
            o[f"{c}125"] = o[f"{c}112"] * 0.5
        o["L125"] = (o["L112"] * 0.5 + (o["L96"] + o["L97"] + o["L99"]) / 2
                     - OWOW_YEAR6_DEDUCTION)
        for c in "MNOPQRSTUVW":
            o[f"{c}125"] = o[f"{c}112"] * 0.5 + (o[f"{c}96"] + o[f"{c}97"] + o[f"{c}99"]) / 2
        o["X125"] = o["X112"] * 0.5 + o["X107"] + o["X106"]
        self._returns(o, 125, rate, o["G125"])

        o["D133"] = mi["discount_rate_for_owow_5_yr_npv_block"]
        for c in "GHIK":
            o[f"{c}131"] = o[f"{c}125"]
        # EXCEL-QUIRK: J131 subtracts J111/2, which is an empty label row (so nothing).
        o["J131"] = o["J125"] - 0 / 2
        o["L131"] = xl.NPV(o["D133"], [o[f"{c}125"] for c in cols[cols.index("L"):]])
        self._returns(o, 131, rate, o["G131"])

    @staticmethod
    def _returns(o, row, rate, equity):
        """IRR, total distributions, NPV and equity multiple for one cash-flow row.

        Rows below `row`: +1 total distributions, +2 NPV, +3 equity multiple.
        """
        flows = [o.get(f"{c}{row}") for c in INV_COLS]
        o[f"F{row}"] = xl.IFERROR(lambda: xl.IRR(flows), "#NUM!")   # Excel shows #NUM!
        o[f"F{row + 1}"] = xl.SUM(flows) - flows[0]
        o[f"F{row + 2}"] = xl.NPV(rate, flows)
        o[f"F{row + 3}"] = o[f"F{row + 1}"] / -equity
