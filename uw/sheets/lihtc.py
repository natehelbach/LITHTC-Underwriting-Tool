"""LIHTC & Dev Fee: tax credit basis, credits, equity and the developer fee.

Flow: TCAC threshold basis limits (units x per-unit limit, plus boosts) cap the
eligible basis from Dev Budget; the high-cost (DDA/QCT) boost and applicable
fraction give qualified basis; credits x pricing give equity, paid in by
milestone. The developer fee is 15% of eligible construction basis; the TCAC
rules for the selected deal type split it into a paid and a deferred portion.
"""
from uw import xl
from uw.engine import Sheet, register

DB = "Dev Budget"
SU = "Sources & Uses"
FOF = "Flow of Funds"
RB = "Rent Builder"

# 2024 TCAC threshold basis limits per unit (typed on the tab, column F).
TBL_2024 = {13: ("Studio", 387110), 14: ("1 BR Units", 446334), 15: ("2 BR Units", 538400),
            16: ("3 BR Units", 689152), 17: ("4 BR Units", 0)}

# TBL boosts: row -> (boost %, Master Inputs Yes/No flag).
TBL_BOOSTS = {
    22: (0.2, "tbl_boost_applies_state_or_federal_prevailing_wage"),
    23: (0.05, "tbl_boost_applies_labor_agreement"),
    24: (0.1, "tbl_boost_applies_new_construction_with_structured_parking"),
    25: (0.02, "tbl_boost_applies_day_care_center"),
    26: (0.02, "tbl_boost_applies_100pct_special_needs"),
    27: (0.1, "tbl_boost_applies_95pct_upper_floor_units_served_by_elevator"),
    28: (0.15, "tbl_boost_applies_type_i_building"),
    29: (0.1, "tbl_boost_applies_type_iii_building"),
}
HIGH_OPP_BOOST = 0.1           # E33
HIGH_COST_BOOST = 1.3          # G47: DDA / QCT
FEDERAL_9PCT_CAP = 25_000_000  # G57
STATE_CREDIT_CAP = 0.3         # F61

DEV_FEE_PCT = 0.15             # typed 15% (9% and acq/rehab sections)
ACQ_BASIS_PCT = 0.05
CAP_9PCT, CAP_9PCT_SN = 2_500_000, 2_800_000
BASIS_THRESHOLD = 6_666_667
PLUS_1M = 1_000_000
CAP_6M = 6_000_000
SN_EXCESS_PCT = 0.07           # special-needs: 7% of basis above threshold

# EXCEL-QUIRK: 4% NC paid developer fee is typed as $3.0M (H174) instead of the
# "lesser of $6M or (c)" test the row label describes.
PAID_DEV_FEE_4PCT_NC = 3_000_000

# Equity pay-in splits typed in J68:J70 (memo).
PAYIN_MEMO = {68: 50, 69: 48, 70: 2}


def _is(rate, target):
    return abs(xl.n(rate) - target) < 1e-12


@register
class LIHTC(Sheet):
    name = "LIHTC & Dev Fee"

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        credit = mi["credit_type"]
        o["B2"] = f"{mi['project_name']}  |  LIHTC & DEV FEE"

        # ---------------- unit counts & applicable fraction
        o["C217"] = ctx.ref(RB, "G44")
        o["C218"] = ctx.ref(RB, "F44")
        o["C216"] = xl.n(o["C217"]) + xl.n(o["C218"])
        o["C219"] = xl.n(ctx.ref("Program", "E64")) - o["C216"]
        if o["C219"] == 0:
            o["C221"] = 1
        else:
            o["C221"] = (xl.SUMPRODUCT(ctx.col(RB, "G12:G39"), ctx.col(RB, "D12:D39"))
                         / xl.n(ctx.ref("Program", "C29")))

        # ---------------- threshold basis limits
        o["E11"] = 2024
        units = {13: ctx.ref(RB, "E53"), 14: ctx.ref(RB, "E54"), 15: ctx.ref(RB, "E55"),
                 16: ctx.ref(RB, "E56"), 17: 0, 18: ctx.ref(RB, "F44")}
        for r, (label, limit) in TBL_2024.items():
            o[f"D{r}"] = label
            o[f"F{r}"] = limit
        o["D18"] = "Manager"
        mgr_type = ctx.ref("Program", "H105")
        o["F18"] = o["F15"] if xl.eq(mgr_type, "2 bedroom") else o["F14"] if xl.eq(mgr_type, "1 Bedroom") else 0
        for r in range(13, 19):
            o[f"E{r}"] = units[r]
            o[f"G{r}"] = xl.n(units[r]) * o[f"F{r}"]
        o["E19"] = xl.SUM([o[f"E{r}"] for r in range(13, 19)])
        tbl = o["G19"] = sum(o[f"G{r}"] for r in range(13, 19))

        for r, (pct, key) in TBL_BOOSTS.items():
            o[f"E{r}"] = pct
            o[f"F{r}"] = mi[key]
            o[f"G{r}"] = pct * tbl if xl.eq(o[f"F{r}"], "Yes") else 0
        o["E30"] = mi["tbl_boost_energy_efficiency_resource_conservation"]
        o["G30"] = xl.n(o["E30"]) * tbl
        o["G31"] = mi["tbl_boost_seismic_environmental_mitigation"]
        o["G32"] = mi["tbl_boost_development_impact_fees"]
        o["E33"] = HIGH_OPP_BOOST
        o["F33"] = "Yes" if xl.eq(mi["high_highest_resource_area"], "Yes") else "No"
        o["G33"] = HIGH_OPP_BOOST * tbl if o["F33"] == "Yes" else 0

        # Affordability boosts (4% deals): share of restricted units by AMI band x TBL.
        rb_units = ctx.col(RB, "G12:G39")
        rb_ami = ctx.col(RB, "J12:J39")
        restricted = xl.n(ctx.ref(RB, "G44"))
        if _is(credit, 0.04):
            # EXCEL-QUIRK: labels say 35-50% / <=35% AMI, but criteria are >30%,<51% and <31%.
            o["G35"] = xl.SUMIFS(rb_units, rb_ami, ">0.3", rb_ami, "<0.51") / restricted * tbl
            o["G36"] = xl.SUMIF(rb_ami, "<0.31", rb_units) / restricted * tbl
        else:
            o["G35"] = o["G36"] = 0
        o["G37"] = xl.SUM([o.get(f"G{r}") for r in range(22, 37)]) + tbl

        # ---------------- eligible & qualified basis
        o["G40"] = xl.IFERROR(lambda: xl.n(ctx.ref(DB, "Q119")), 0)
        o["G41"] = xl.n(ctx.ref(DB, "P119"))
        o["G42"] = o["G40"] + o["G41"]
        # EXCEL-QUIRK: when TBL binds, allowable basis is TBL minus acquisition basis.
        o["G44"] = o["G42"] if o["G42"] < o["G37"] else o["G37"] - o["G41"]
        o["G45"] = xl.n(ctx.ref(SU, "E62"))
        o["G46"] = o["G44"] - o["G45"]
        o["G47"] = HIGH_COST_BOOST if (xl.eq(mi["qct_dda_status"], "DDA")
                                       or xl.eq(mi["qct_dda_status"], "QCT")) else 1
        o["G48"] = o["G47"] * o["G46"]
        o["G49"] = 1 if o["C219"] == 0 else o["C221"]
        o["G50"] = o["G49"] * o["G48"]
        o["F51"] = mi["credit_reduction_pct"]
        o["G51"] = -xl.n(o["F51"]) * o["G50"]
        o["G52"] = o["G51"] + o["G50"]

        # ---------------- credits & equity
        o["G55"] = o["G52"]
        o["F56"] = credit
        o["G56"] = credit * o["G55"]
        o["G57"] = min(o["G56"] * 10, FEDERAL_9PCT_CAP) if _is(credit, 0.09) else o["G56"] * 10
        o["F58"] = ctx.ref("Sensitivities", "C19")
        o["G58"] = xl.n(o["F58"]) * o["G57"]
        o["F60"] = mi["state_lihtc_available"]
        o["F61"] = min(xl.n(mi["state_credit_pct_of_basis"]), STATE_CREDIT_CAP)
        o["G61"] = o["F61"] * o["G46"] if xl.eq(o["F60"], "yes") else 0
        o["G62"] = xl.n(ctx.ref(SU, "E63"))
        o["G63"] = o["G61"] - o["G62"]
        o["F64"] = ctx.ref("Sensitivities", "C20")
        o["G64"] = xl.n(o["F64"]) * o["G63"]
        equity = o["G66"] = o["G64"] + o["G58"]

        # Equity pay-in timing (Flow of Funds milestones).
        o["G210"] = ctx.ref(FOF, "I10")
        o["G211"] = ctx.ref(FOF, "I14")
        o["G209"] = xl.n(o["G210"]) + xl.n(o["G211"])
        o["G212"] = 1 - o["G209"]
        o["F207"] = "30% max" if xl.n(mi["state_credit_pct_of_basis"]) > STATE_CREDIT_CAP else ""
        o["F68"] = xl.n(ctx.ref(FOF, "R11")) + xl.n(ctx.ref(FOF, "R12"))
        o["G68"] = equity * o["F68"]
        o["F69"] = o["G212"]
        o["G69"] = o["F69"] * equity
        o["F70"] = 1 - (o["F68"] + o["F69"])
        o["G70"] = equity - o["G68"] - o["G69"]
        o["G71"] = o["G68"] + o["G69"] + o["G70"]
        for r, v in PAYIN_MEMO.items():
            o[f"J{r}"] = v

        # ---------------- developer fee: basis inputs from Dev Budget
        q119, q115 = xl.n(ctx.ref(DB, "Q119")), xl.n(ctx.ref(DB, "Q115"))
        p119, p115 = xl.n(ctx.ref(DB, "P119")), xl.n(ctx.ref(DB, "P115"))
        p112 = xl.n(ctx.ref(DB, "P112"))
        nonres = xl.n(ctx.ref(DB, "V119"))
        constr_basis = q119 - q115                      # eligible construction basis ex dev fee
        acq_basis = p119 - p112
        basis_ex_fee = q119 + p119 - p115 - q115         # unadjusted eligible basis ex dev fee

        def fee_on_basis(top, with_acq):
            """15% of construction basis (+5% of acquisition basis) + 15% of non-res costs."""
            o[f"F{top}"] = DEV_FEE_PCT
            o[f"G{top}"] = constr_basis
            o[f"H{top}"] = constr_basis * DEV_FEE_PCT
            r = top + 1
            if with_acq:
                o[f"F{r}"] = ACQ_BASIS_PCT
                o[f"G{r}"] = acq_basis
                o[f"H{r}"] = acq_basis * ACQ_BASIS_PCT
                r += 1
            o[f"F{r}"] = DEV_FEE_PCT
            o[f"G{r}"] = nonres
            o[f"H{r}"] = nonres * DEV_FEE_PCT
            o[f"H{r + 1}"] = sum(o[f"H{x}"] for x in range(top, r + 1))
            return o[f"H{r + 1}"]

        # 9% new construction / rehab
        h88 = fee_on_basis(86, False)
        o["G90"], o["G91"] = CAP_9PCT, CAP_9PCT_SN
        o["H90"], o["H91"] = min(CAP_9PCT, h88), min(CAP_9PCT_SN, h88)
        # 9% acquisition / rehab
        h99 = fee_on_basis(96, True)
        o["G101"], o["G102"] = CAP_9PCT, CAP_9PCT_SN
        o["H101"], o["H102"] = min(CAP_9PCT, h99), min(CAP_9PCT_SN, h99)

        # 4% acquisition / rehab (and special needs variant).
        def acq_rehab_4pct(top, cap_a, excess_pct):
            total = fee_on_basis(top, True)          # rows top..top+3
            a = min(o[f"H{top + 1}"] + o[f"H{top}"], cap_a)
            excess = basis_ex_fee - BASIS_THRESHOLD
            b = excess * excess_pct + PLUS_1M
            c = max(b, a)
            paid = min(c, CAP_6M)
            return total, a, excess, excess * excess_pct, b, c, paid

        total, a, ex, ex_pct, b, c, paid = acq_rehab_4pct(110, CAP_9PCT, ACQ_BASIS_PCT)
        o.update({"H117": o["H110"], "H118": o["H111"], "E119": CAP_9PCT, "H119": a,
                  "E122": BASIS_THRESHOLD, "H122": ex, "E123": BASIS_THRESHOLD, "F123": ACQ_BASIS_PCT,
                  "H123": ex_pct, "E124": PLUS_1M, "H124": b, "H126": c, "E127": CAP_6M, "H127": paid,
                  "H129": total, "H130": paid, "H131": total - paid})
        total, a, ex, ex_pct, b, c, paid = acq_rehab_4pct(136, CAP_9PCT_SN, SN_EXCESS_PCT)
        o.update({"H143": o["H136"], "H144": o["H137"], "F145": CAP_9PCT_SN, "H145": a,
                  "E148": BASIS_THRESHOLD, "H148": ex, "E149": BASIS_THRESHOLD, "F149": SN_EXCESS_PCT,
                  "H149": ex_pct, "E150": PLUS_1M, "H150": b, "H152": c, "E153": CAP_6M, "H153": paid,
                  "H155": total, "H156": paid, "H157": total - paid})

        # 4% new construction (current project type) - percentages from Master Inputs.
        pct = mi["developer_fee_pct_of_eligible_basis"]
        o["F162"] = o["F163"] = pct
        o["G162"] = constr_basis
        o["H162"] = constr_basis * pct
        o["G163"] = nonres
        o["H163"] = nonres * pct
        o["H164"] = o["H162"] + o["H163"]
        o["G167"] = 2500000
        o["H167"] = mi["paid_fee_test_2_5m"]
        o["E169"] = mi["paid_fee_test_basis_threshold_6_666_667"]
        o["H169"] = basis_ex_fee - o["E169"]
        o["E170"] = BASIS_THRESHOLD
        o["F170"] = mi["paid_fee_test_pct_of_basis_above_threshold"]
        o["H170"] = o["H169"] * o["F170"]
        o["E171"] = mi["paid_fee_test_plus_1_000_000"]
        o["H171"] = o["H170"] + o["E171"]
        perm_loan = xl.n(ctx.ref(SU, "E47"))
        o["J171"] = perm_loan * 0.05
        o["J173"] = mi["paid_dev_fee_cap_as_pct_of_perm_loan"]
        o["E174"] = mi["paid_fee_cap_lesser_of_6m"]
        o["H174"] = PAID_DEV_FEE_4PCT_NC
        # Pay-out schedule of the paid fee (Flow of Funds milestone shares).
        o["K176"] = ctx.ref(FOF, "X11")
        o["K177"] = ctx.ref(FOF, "X13")
        o["K178"] = ctx.ref(FOF, "X15")
        o["K179"] = sum(xl.n(o[f"K{r}"]) for r in (176, 177, 178))
        o["H173"] = min(max(o["H171"], o["H167"]), perm_loan * o["J173"] / xl.n(o["K177"]))
        o["J174"] = o["H174"] * xl.n(o["K177"]) / perm_loan
        o["H176"] = o["H164"]
        o["H177"] = o["H174"]
        o["H178"] = o["H176"] - o["H177"]
        for r in (176, 177, 178):
            o[f"L{r}"] = xl.n(o[f"K{r}"]) * o["H177"]
        o["L179"] = o["L176"] + o["L177"] + o["L178"]
        o["J170"] = xl.n(ctx.ref(DB, "H39")) + o["L176"]
        o["L181"] = ctx.ref(SU, "E58")

        # 4% new construction, special needs.
        fee_on_basis(183, False)
        o["G188"] = CAP_9PCT_SN
        o["H188"] = min(CAP_9PCT_SN, o["H183"])
        o["H190"] = basis_ex_fee - BASIS_THRESHOLD
        o["H191"] = o["H190"] * SN_EXCESS_PCT
        o["H192"] = o["H191"] + PLUS_1M
        o["H194"] = max(o["H192"], o["H188"])
        o["G195"] = CAP_6M
        o["H195"] = min(o["H194"], CAP_6M)
        o["H197"] = o["H164"]     # EXCEL-QUIRK: uses the 4% NC total (H164), not this section's H185
        o["H198"] = o["H195"]
        o["H199"] = o["H197"] - o["H198"]

        o["G203"] = xl.n(ctx.ref(DB, "H113")) / min(xl.n(o["E174"]), max(o["H171"], xl.n(o["H167"])))

        # ---------------- selected developer fee (by credit type and project type)
        acq = xl.eq(ctx.ref("Lists", "Q7"), "Yes")
        sn = xl.eq(ctx.ref("Lists", "Q8"), "Yes")
        if _is(credit, 0.09):
            total_sel = paid_sel = {(True, True): "H91", (True, False): "H90",
                                    (False, True): "H102", (False, False): "H101"}[(acq, sn)]
            deferred = 0
        elif _is(credit, 0.04):
            total_sel, paid_sel, deferred_cell = {
                (False, True): ("H197", "H198", "H199"), (False, False): ("H176", "H177", "H178"),
                (True, True): ("H155", "H156", "H157"), (True, False): ("H129", "H130", "H131"),
            }[(acq, sn)]
            deferred = o[deferred_cell]
        else:
            total_sel = paid_sel = None
            deferred = 0      # nested IF falls through to FALSE
        # EXCEL-QUIRK: E76/E78 add H97 (9% acq/rehab 5%-of-acquisition-basis line) for every deal type.
        o["E76"] = (o[total_sel] if total_sel else 0) + o["H97"]
        o["E77"] = o[paid_sel] if paid_sel else False
        o["E78"] = deferred + o["H97"]
        return o
