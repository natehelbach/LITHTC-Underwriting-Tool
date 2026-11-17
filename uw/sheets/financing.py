"""Financing & Debt Sizing: three debt instruments, perm-loan sizing and annual schedules.

Debt 1 (column D) - primary construction-to-perm loan (tax-exempt bonds + taxable tail),
                    permanent amount sized off DSCR ("Auto") or hardcoded.
Debt 2 (column E) - secondary C-to-P loan (e.g. Apple AHF); perm amount = drawn
                    construction balance at conversion, capped at commitment.
Debt 3 (column F) - soft / HCD loan, interest-only or deferred.

Sections, top to bottom:
  1. instrument terms (links to Master Inputs)        rows 10-41
  2. permanent loan sizing + loan metrics               rows 45-62
  3. construction financing (Flow of Funds links)       rows 66-89
  4. annual debt schedules, 60 years (I = year 1)       rows 93-134
  5. debt fee schedule (feeds Dev Budget by tag)        rows 138-157
  6. supporting calcs: rate blend, bond test, 9% loan   rows 162-195

Circular with Operating Cash Flow, Sources & Uses, Flow of Funds and Dev Budget;
those are read through ctx and the model loop resolves the fixed point.
"""
from openpyxl.utils import get_column_letter as col_letter

from uw import xl
from uw.engine import Sheet, register

YEARS = 60
DEBTS = ("D", "E", "F")                      # Debt 1 / 2 / 3 columns
DEBT_LABELS = ("Debt 1 - Primary", "Debt 2 - Secondary C-to-P", "Debt 3 - Soft / HCD")


def fin_col(year):
    """Annual schedule column on this tab: I = operating year 1."""
    return col_letter(8 + year)


def ocf_col(year):
    """Operating Cash Flow column: J = operating year 1."""
    return col_letter(9 + year)


# Debt fee schedule typed on the tab (rows 139-156, columns B..M).
# (row, instrument, slot, name, basis, rate/$, flat add-on, threshold, timing, frequency,
#  #periods, capitalized?, Dev Budget tag). None = cell left for a formula below.
FEES = [
    (139, 0, "Fee 1", "Bond issuance & monitoring (per SD budget)", "Flat", 145234.445163614, 0, 0,
     "Closing", "One-time", 1, "Y", "Bond Issuance"),
    (140, 0, "Fee 2", "Bond annual admin fee (construction)", "Constr Commitment", 0, 0, 0,
     "Annual", "Annual", 3, "Y", "Bond Issuance"),
    (141, 0, "Fee 3", "Construction loan origination", "Constr Commitment", None, 0, 0,
     "Closing", "One-time", 1, "Y", "Construction Origination"),
    (142, 0, "Fee 4", "Lender inspections", "Flat", 75000, 0, 0,
     "Draw", "One-time", 1, "Y", "Lender Inspections"),
    (143, 0, "Fee 5", "Permanent loan origination", "Perm Commitment", None, 0, 0,
     "Conversion", "One-time", 1, "Y", "Perm Origination"),
    (144, 0, "Fee 6", "Bond monitoring fee (operating)", "Perm Commitment", 0.0005, 0, 0,
     "Annual", "Annual", None, "N", "OpEx Monitoring"),
    (145, 1, "Fee 1", "Application fee (credited vs. origination)", "Flat", None, 0, 0,
     "Closing", "One-time", 1, "Y", "Construction Origination"),
    (146, 1, "Fee 2", "Origination fee at closing (net of app. credit)", "Flat", None, 0, 0,
     "Closing", "One-time", 1, "Y", "Construction Origination"),
    (147, 1, "Fee 3", "Conversion fee", "Flat", None, 0, 0,
     "Conversion", "One-time", 1, "Y", "Perm Origination"),
    (148, 1, "Fee 4", "Third-party costs", "Flat", 15000, 0, 0,
     "Closing", "One-time", 1, "Y", "Construction Origination"),
    (149, 1, "Fee 5", "Annual monitoring fee", "Flat", None, 0, 0,
     "Annual", "Annual", None, "N", "Debt 2 Monitoring"),
    (150, 1, "Fee 6") + (None,) * 10,
    (151, 2, "Fee 1") + (None,) * 10,
    (152, 2, "Fee 2") + (None,) * 10,
    (153, 2, "Fee 3") + (None,) * 10,
    (154, 2, "Fee 4") + (None,) * 10,
    (155, 2, "Fee 5") + (None,) * 10,
    (156, 2, "Fee 6") + (None,) * 10,
]

TE_RATE_APPLICATION = 0.0616          # D171: tax-exempt perm rate at application
RECYCLED_BONDS_HARDCODE = 8836818.61410386   # C187: per source model (Financing K33)
NINE_PCT_LOAN = {"rate": 0.04, "avg_out_constr": 0.5, "avg_out_leaseup": 1, "ltc": 0.8}


@register
class Financing(Sheet):
    name = "Financing & Debt Sizing"

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        sens = lambda r: ctx.ref("Sensitivities", f"C{r}")
        su = lambda a: ctx.ref("Sources & Uses", a)
        fof = lambda a: ctx.ref("Flow of Funds", a)
        ocf = lambda a: ctx.ref("Operating Cash Flow", a)

        o["B2"] = f"{mi['project_name']}  |  FINANCING & DEBT SIZING"
        o["D10"], o["E10"], o["F10"] = DEBT_LABELS

        # ---- 1. instrument terms ------------------------------------------------
        o["D12"], o["E12"], o["F12"] = mi["d1_name"], mi["d2_name"], mi["d3_name"]
        o["D13"], o["E13"], o["F13"] = mi["d1_include"], mi["d2_include"], mi["d3_include"]
        o["D14"], o["E14"], o["F14"] = mi["d1_structure"], mi["d2_structure"], "Custom"
        o["D15"] = "Hardcode" if xl.eq(mi["d1_sizing"], "Hard Code") else "Auto"
        o["E15"], o["F15"] = mi["d2_sizing"], "Hardcode"
        o["D16"], o["E16"], o["F16"] = mi["d1_amount"], mi["d2_amount"], mi["d3_amount"]
        o["D17"], o["E17"], o["F17"] = mi["d1_lien"], mi["d2_lien"], "Subordinate"

        # construction phase
        o["D20"], o["E20"], o["F20"] = mi["d1_c_rate_type"], mi["d2_c_rate_type"], "n/a"
        o["D21"], o["E21"], o["F21"] = sens(16), mi["d2_c_rate"], 0
        o["D22"], o["E22"], o["F22"] = mi["d1_c_spread"], mi["d2_c_spread"], 0
        for c in "DE":
            o[f"{c}23"] = (xl.n(o[f"{c}21"]) + xl.n(o[f"{c}22"]) if xl.eq(o[f"{c}20"], "Floating")
                           else o[f"{c}21"])
        o["F23"] = o["F21"] + o["F22"]
        o["D24"], o["E24"], o["F24"] = mi["d1_c_int_treat"], mi["d2_c_int_treat"], "n/a"
        o["D25"] = sens(17)
        o["D26"] = mi["te_bond_amount"]
        o["D27"] = mi["recycled_pct"]

        # permanent phase
        o["D30"], o["E30"], o["F30"] = mi["d1_p_rate_type"], mi["d2_p_rate_type"], "Fixed"
        o["D31"], o["E31"], o["F31"] = mi["d1_p_rate"], mi["d2_p_rate"], mi["d3_rate"]
        o["D32"], o["E32"], o["F32"] = mi["d1_p_spread"], mi["d2_p_spread"], 0
        o["D34"], o["E34"], o["F34"] = mi["d1_dscr"], mi["d2_dscr"], "n/a"
        o["D35"] = o["E35"] = o["F35"] = mi["d1_test_year"]
        o["D36"], o["E36"], o["F36"] = mi["d1_io"], mi["d2_io"], 0
        o["D37"], o["E37"], o["F37"] = mi["d1_amort_type"], mi["d2_amort_type"], "Interest Only"
        o["D38"], o["E38"], o["F38"] = mi["d1_amort"], mi["d2_amort"], 0
        o["D39"], o["E39"], o["F39"] = mi["d1_term"], mi["d2_term"], mi["d3_term"]
        o["D40"], o["E40"], o["F40"] = mi["d1_ppy"], mi["d2_ppy"], 1
        o["F41"] = mi["d3_repay"]

        # ---- 6a. permanent rate blend (TE / taxable tranches) -------------------
        o["C174"] = mi["recycled_yn"]
        o["B186"] = mi["recycled_pct"]
        o["B187"] = xl.IFERROR(lambda: xl.n(su("E26")) * xl.n(o["B186"]), 0)
        # EXCEL-QUIRK: C187 (used in the blend when recycling = yes) is a hardcode, not B187.
        o["C187"] = RECYCLED_BONDS_HARDCODE
        te_tranche = xl.n(su("E30"))
        o["C171"] = te_tranche + o["C187"] if xl.eq(o["C174"], "yes") else te_tranche
        o["C172"] = xl.n(su("E47")) - o["C171"]
        o["D171"] = TE_RATE_APPLICATION
        o["E171"] = mi["d1_p_rate"]
        o["D172"] = o["D171"] + mi["d1_p_spread"]
        o["E172"] = o["E171"] + mi["d1_p_spread"]
        amounts = [o["C171"], o["C172"]]
        o["C173"] = sum(amounts)
        # (div default 0 only guards the model's cold start, before Sources & Uses has run)
        o["D173"] = xl.div(xl.SUMPRODUCT([o["D171"], o["D172"]], amounts), o["C173"], 0.0)
        o["E173"] = xl.div(xl.SUMPRODUCT([o["E171"], o["E172"]], amounts), o["C173"], 0.0)

        # permanent rate: Debt 1 uses the blended TE/taxable rate
        if xl.eq(o["D30"], "Blended TE/Taxable"):
            o["D33"] = o["E173"]
        else:
            o["D33"] = o["D31"] + o["D32"] if xl.eq(o["D30"], "Floating") else o["D31"]
        for c in "EF":
            o[f"{c}33"] = o[f"{c}31"] + o[f"{c}32"] if xl.eq(o[f"{c}30"], "Floating") else o[f"{c}31"]

        # ---- 3 (part). Debt 2 construction-to-perm terms -------------------------
        o["E79"] = sens(26)
        o["E80"] = max(0, o["E39"] * 12 - o["E79"])
        if not xl.eq(o["E13"], "Yes"):
            o["E81"] = 0
        else:
            draws = ctx.col("Flow of Funds", "L236:BH236")
            phases = ctx.col("Flow of Funds", "L23:BH23")
            o["E81"] = (xl.SUMIFS(draws, phases, fof("F17"))
                        + xl.SUM(ctx.col("Flow of Funds", "L195:BH195")))
        o["E82"] = xl.MAX(ctx.col("Flow of Funds", "L236:BH236"))

        # ---- 2. permanent loan sizing --------------------------------------------
        # Loan metrics that do not depend on the loan amount.
        for c in DEBTS:
            io_only = xl.eq(o[f"{c}37"], "Interest Only")
            o[f"{c}58"] = 0 if io_only else o[f"{c}38"] * o[f"{c}40"]
        o["D59"] = o["D39"] * o["D40"]
        o["E59"] = o["E80"] / 12 * o["E40"]
        o["F59"] = o["F39"] * o["F40"]
        # EXCEL-QUIRK: rows 57-59 are labelled "months" but are payment periods; Debt 3 pays
        # annually (F40 = 1), so its 55 "IO months" are really 55 years.
        for c in DEBTS:
            io_only = xl.eq(o[f"{c}37"], "Interest Only")
            o[f"{c}57"] = o[f"{c}59"] if io_only else min(o[f"{c}36"], o[f"{c}59"])

        def test_year_income(test_year):
            # EXCEL-QUIRK: OFFSET(I55, 0, MATCH(year, I12:BP12)) lands one column past the
            # test year, i.e. uses income of the year AFTER the DSCR test year (flagged on the tab).
            years_row = ctx.col("Operating Cash Flow", "I12:BP12")
            if not any(years_row):          # cold start: Operating Cash Flow not run yet
                return 0.0
            xl.MATCH(test_year, years_row)  # #N/A if the test year is outside years 1-59
            return xl.n(ocf(f"{ocf_col(int(test_year) + 1)}55"))

        yr1_income = xl.n(ocf("J55"))
        o["D45"] = o["E45"] = yr1_income
        o["D46"] = test_year_income(o["D35"])
        o["E46"] = test_year_income(o["E35"])
        o["D47"] = -(xl.n(ocf("J60")) + xl.n(ocf("J59")) + xl.n(ocf("J61")))

        def dscr_loan(c):
            o[f"{c}48"] = max(0, min(o[f"{c}46"], o[f"{c}45"]) / o[f"{c}34"] - o[f"{c}47"])
            rate, ppy, ds = o[f"{c}33"], o[f"{c}40"], o[f"{c}48"]
            if xl.eq(o[f"{c}37"], "Interest Only"):
                o[f"{c}49"] = xl.IFERROR(lambda: ds / rate, 0)
            else:
                o[f"{c}49"] = xl.IFERROR(lambda: -xl.PV(rate / ppy, o[f"{c}38"] * ppy, ds / ppy), 0)

        def loan_metrics(c):
            amt, rate, ppy = o[f"{c}51"], o[f"{c}33"], o[f"{c}40"]
            io_only = xl.eq(o[f"{c}37"], "Interest Only")
            if c == "F":
                o["F54"] = 0 if amt == 0 or xl.eq(o["F41"], "Deferred") else amt * rate / ppy
            elif amt == 0:
                o[f"{c}54"] = 0
            else:
                o[f"{c}54"] = amt * rate / ppy if io_only else -xl.PMT(rate / ppy, o[f"{c}58"], amt)
            o[f"{c}55"] = o[f"{c}54"] * ppy

        # Debt 1: DSCR-sized (Auto) or hardcoded.
        dscr_loan("D")
        o["D50"] = max(0, min(o["D49"], xl.n(su("E26")) - xl.SUM(
            su("E46"), ctx.col("Sources & Uses", "E48:E56"))))
        if not xl.eq(o["D13"], "Yes"):
            o["D51"] = 0
        else:
            o["D51"] = o["D16"] if xl.eq(o["D15"], "Hardcode") else o["D49"]
        # Debt 3: hardcoded soft loan.
        o["F51"] = o["F16"] if xl.eq(o["F13"], "Yes") else 0
        loan_metrics("D")
        loan_metrics("F")
        # Debt 2 sizes behind Debt 1 + Debt 3 debt service.
        o["E47"] = o["D55"] + o["F55"]
        dscr_loan("E")
        if not xl.eq(o["E13"], "Yes"):
            o["E51"] = 0
        else:
            sized = o["E16"] if xl.eq(o["E15"], "Hardcode") else o["E49"]
            o["E51"] = min(sized, o["E81"]) if xl.eq(o["E14"], "Construction-to-Perm") else sized
        loan_metrics("E")
        o["D56"] = o["C171"]

        for c in DEBTS:
            amt, rate, ppy = o[f"{c}51"], o[f"{c}33"], o[f"{c}40"]
            io_m, term, pmt = o[f"{c}57"], o[f"{c}59"], o[f"{c}54"]
            # balloon at maturity
            if c == "F":
                o["F60"] = amt
            elif amt == 0:
                o[f"{c}60"] = 0
            elif xl.eq(o[f"{c}37"], "Interest Only"):
                o[f"{c}60"] = amt
            else:
                o[f"{c}60"] = max(0, xl.ROUND(-xl.FV(rate / ppy, term - io_m, -pmt, amt), 2))
            o[f"{c}61"] = 0 if amt == 0 else 1 - o[f"{c}60"] / amt
            # balance at the test / hold year (feeds sale proceeds)
            test_m = o[f"{c}35"] * ppy
            if amt == 0 or test_m >= term:
                o[f"{c}62"] = 0
            elif c == "F" or test_m <= io_m:
                o[f"{c}62"] = amt
            else:
                o[f"{c}62"] = -xl.FV(rate / ppy, test_m - io_m, -pmt, amt)

        # ---- 3. construction financing --------------------------------------------
        o["D66"] = (xl.n(su("E30")) + xl.n(su("E31")) + xl.n(su("E32"))
                    if xl.eq(o["D14"], "Construction-to-Perm") else 0)
        o["E66"] = (o["E16"] if xl.eq(o["E13"], "Yes") and xl.eq(o["E14"], "Construction-to-Perm")
                    else 0)
        o["F66"] = 0
        o["D67"], o["D68"], o["D69"] = su("E30"), su("E31"), su("E32")
        o["D70"] = o["D66"] - o["D51"]
        o["E70"] = o["E66"] - o["E51"]
        o["D72"] = fof("D231")
        o["E72"] = fof("D237")
        o["D73"] = xl.n(fof("D215")) + xl.n(fof("D222")) + xl.n(fof("D229"))

        # 6b. construction rates and the 25% bond test
        o["D178"] = sens(16)
        o["E178"] = o["D178"]
        o["D179"] = sens(17)
        o["E179"] = o["D179"]
        o["D180"] = mi["te_bond_amount"]
        db = lambda a: ctx.ref("Dev Budget", a)

        def agg_basis():
            land = xl.n(db("H14")) - xl.n(db("P14"))
            return (xl.n(db("Q119")) + (land if xl.eq(mi["ground_lease"], "No") else land * 0)
                    + xl.n(db("V119")))
        o["D181"] = xl.IFERROR(agg_basis, 0)
        four_pct = mi["credit_type"] == 0.04
        o["D182"] = xl.IFERROR(lambda: o["D180"] / o["D181"] if four_pct else "N/A", 0)
        o["D183"] = ("Pass" if xl.div(o["D180"], o["D181"], 0.0) > 0.25 else "Fail") if four_pct else "N/A"
        o["D74"], o["D75"], o["D76"] = o["D181"], o["D182"], o["D183"]

        o["E83"] = "OK" if xl.ROUND(o["E82"] - o["E16"], 0) <= 0 else "OVER COMMITMENT"
        o["E84"] = o["E72"]
        d2_on = xl.eq(o["E13"], "Yes")
        o["E85"] = max(mi["d2_orig_rate"] * o["E16"], mi["d2_orig_min"]) if d2_on else 0
        o["E86"] = min(mi["d2_app_fee"], o["E85"]) if d2_on else 0
        o["E87"] = o["E85"] - o["E86"]
        o["E88"] = mi["d2_conv_fee"] if d2_on else 0
        o["E89"] = (mi["d2_monitor_rate"] * o["E16"]
                    if d2_on and not xl.eq(mi["d2_monitor_waive"], "Yes") else 0)

        # ---- 5. debt fee schedule ------------------------------------------------
        formula_inputs = {141: ("F", mi["fee_constr_orig"]), 143: ("F", mi["fee_perm_orig"]),
                          145: ("F", mi["d2_app_fee"]), 146: ("F", o["E87"]),
                          147: ("F", o["E88"]), 149: ("F", o["E89"]),
                          144: ("K", o["D39"])}
        o["K149"] = 0 if o["E40"] == 0 else xl.ROUNDUP(o["E59"] / o["E40"], 0)
        for row, inst, *vals in FEES:
            o[f"B{row}"] = DEBT_LABELS[inst]
            for col, v in zip("CDEFGHIJKLM", vals):
                if v is not None:
                    o[f"{col}{row}"] = v
            if row in formula_inputs:
                col, v = formula_inputs[row]
                o[f"{col}{row}"] = v
            basis = o.get(f"E{row}")
            c = DEBTS[inst]
            if xl.eq(basis, "Flat"):
                o[f"N{row}"] = 1
            else:
                if basis in ("Constr Commitment", "Draw"):
                    base = o[f"{c}66"]
                elif basis in ("Perm Commitment", "Balance"):
                    base = o[f"{c}51"]
                else:
                    base = 0
                o[f"N{row}"] = max(0, base - xl.n(o.get(f"H{row}")))
            on = xl.eq(o[f"{c}13"], "Yes")
            o[f"O{row}"] = (xl.n(o.get(f"F{row}")) * o[f"N{row}"] + xl.n(o.get(f"G{row}"))) if on else 0
            o[f"P{row}"] = o[f"O{row}"] * xl.n(o.get(f"K{row}"))
        o["P157"] = sum(o[f"P{r}"] for r in range(139, 157))

        # ---- 4. annual debt schedules ----------------------------------------------
        years = ctx.col("Operating Cash Flow", "J12:BQ12")
        if None in years:                   # cold start: Operating Cash Flow not run yet
            years = list(range(1, YEARS + 1))
        for i, yr in enumerate(years):
            o[f"{fin_col(i + 1)}93"] = yr
        o["B96"], o["B108"], o["B121"] = o["D12"], o["E12"], o["F12"]
        blocks = {"D": 97, "E": 109, "F": 122}
        for c, top in blocks.items():
            self._schedule(o, c, top, years)
        # Debt 2 monitoring fee while the loan is outstanding
        o["H119"] = 0.0
        for i in range(1, YEARS + 1):
            x = fin_col(i)
            o[f"{x}119"] = o["O149"] if o[f"{x}113"] > 0 else 0
            o["H119"] += o[f"{x}119"]
        for i in range(1, YEARS + 1):
            x = fin_col(i)
            o[f"{x}133"] = o[f"{x}102"] + o[f"{x}114"] + o[f"{x}127"] + o[f"{x}119"]
            o[f"{x}134"] = xl.IFERROR(lambda: xl.n(ocf(f"{ocf_col(i)}55")) / o[f"{x}133"], "n/a")

        # ---- 6. values read by other tabs ------------------------------------------
        o["D162"] = sens(18)
        o["D163"] = mi["d1_test_year"]
        o["D164"] = mi["d1_amort"]
        o["D165"] = o["D48"]
        o["D166"] = o["D51"]
        o["D167"] = o["D62"] + o["E62"] + o["F62"]

        # 9% construction loan (only used when credit type = 9%)
        nine = NINE_PCT_LOAN
        o["C191"] = nine["rate"]
        o["D191"] = o["C191"]
        o["C192"], o["D192"] = nine["avg_out_constr"], nine["avg_out_leaseup"]
        o["C193"], o["D193"] = sens(26), sens(27)
        o["C194"] = nine["ltc"]

        def nine_pct_loan():
            if mi["credit_type"] != 0.09:
                return 0
            tdc = xl.n(su("E26"))
            other = xl.SUM(su("E29"), ctx.col("Sources & Uses", "E31:E40"))
            return min(tdc - other, o["C194"] * tdc)
        o["C195"] = xl.IFERROR(nine_pct_loan, 0)
        return o

    @staticmethod
    def _schedule(o, c, top, years):
        """Annual schedule for one instrument (rows top .. top+9), by month counts.

        Rows: start month, end month, IO months, amortizing months, beginning balance,
        debt service, principal, interest, balloon, ending balance.
        """
        amt, rate, ppy = o[f"{c}51"], o[f"{c}33"], o[f"{c}40"]
        io_m, term, pmt = o[f"{c}57"], o[f"{c}59"], o[f"{c}54"]
        r = rate / ppy
        deferred = c == "F" and xl.eq(o["F41"], "Deferred")
        tot = {k: 0.0 for k in (5, 6, 7, 8)}
        for i, yr in enumerate(years):
            x = fin_col(i + 1)
            start = (yr - 1) * ppy
            end = min(yr * ppy, term)
            io_in_year = max(0, min(end, io_m) - start)
            amort_in_year = max(0, end - max(start, io_m))
            if amt == 0 or start >= term:
                beg = 0
            elif start <= io_m:
                beg = amt
            else:
                beg = -xl.FV(r, start - io_m, -pmt, amt)
            ds = 0 if deferred else io_in_year * amt * rate / ppy + amort_in_year * pmt
            prin = 0 if amort_in_year == 0 else beg + xl.FV(r, end - io_m, -pmt, amt)
            interest = ds - prin
            balloon = xl.ROUND(beg - prin, 2) if amt > 0 and start < term <= end else 0
            vals = [start, end, io_in_year, amort_in_year, beg, ds, prin, interest, balloon,
                    beg - prin - balloon]
            for k, v in enumerate(vals):
                o[f"{x}{top + k}"] = v
            for k in tot:
                tot[k] += vals[k]
        for k, v in tot.items():
            o[f"H{top + k}"] = v
