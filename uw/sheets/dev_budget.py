"""Dev Budget: total development cost and the tax-credit (eligible basis) split.

Each budget line has a total cost (column H). Most soft/hard costs come from a
$/SF rate on Master Inputs times either the lot area or the gross building area
(Program C24). Every line is then apportioned:

    K / L   residential vs non-residential share  (Master Inputs non-res %)
    N       eligible factor: share of the residential cost that is LIHTC basis
    P / Q   acquisition basis / construction ("rehab") basis
    R       non-basis cost  (H - P - Q)
    T       cost paid at conversion (funded by perm sources, not construction)
    Y / Z   check columns (should be 0)

Construction interest, financing fees, reserves and developer fee come from
Flow of Funds, Financing, Operating Cash Flow and LIHTC (circular; read via ctx).
"""
from uw import xl
from uw.engine import Sheet, register

FIN = "Financing & Debt Sizing"
FOF = "Flow of Funds"
OCF = "Operating Cash Flow"
LIH = "LIHTC & Dev Fee"
SU = "Sources & Uses"

# Line items: row -> (label, $/SF Master Inputs key, area basis). "lot" = lot SF,
# "gba" = gross building area (Program C24).
PSF_LINES = {
    17: ("Title/Escrow/Closing", "title_escrow_closing_sf_of_lot", "lot"),
    22: ("Off-Site Improvements", "off_site_improvements_sf_of_lot", "lot"),
    23: ("Demolition", "demolition_sf_of_lot", "lot"),
    24: ("Site Work/ Landscape / Surface Parking", "site_work_landscape_surface_parking_sf_of_lot", "lot"),
    25: ("Structures - Residential", "structures_residential_sf_of_gross_building_area", "gba"),
    28: ("GC - General Conditions & Profit", "gc_general_conditions_profit_sf_of_gross_building_area", "gba"),
    29: ("GC - Overhead & Profit", "gc_overhead_profit_sf_of_gross_building_area", "gba"),
    31: ("GC Contingency", "gc_contingency_sf_of_gross_building_area", "gba"),
    41: ("Architect Fees", "architect_fees_sf_of_gross_building_area", "gba"),
    42: ("Survey and Engineering", "survey_and_engineering_sf_of_gross_building_area", "gba"),
    46: ("Construction Lender Legal", "construction_lender_legal_sf_of_gross_building_area", "gba"),
    47: ("Perm Related Legal and misc fees ", "perm_related_legal_and_misc_fees_sf_of_gross_building_area", "gba"),
    48: ("Construction third party and due dilligence fees",
         "construction_third_party_and_due_dilligence_fees_sf_of_gross", "gba"),
    49: ("Syndicator legal fees", "syndicator_legal_fees_sf_of_gross_building_area", "gba"),
    50: ("Construction applicant legal", "construction_applicant_legal_sf_of_gross_building_area", "gba"),
    51: ("Perm applicant legal", "perm_applicant_legal_sf_of_gross_building_area", "gba"),
    58: ("Taxes During Construction", "taxes_during_construction_sf_of_gross_building_area", "gba"),
    59: ("Performance Bond Premium", "performance_bond_premium_sf_of_gross_building_area", "gba"),
    61: ("Title and Recording", "title_and_recording_sf_of_gross_building_area", "gba"),
    62: ("Builders Risk Insurance", "builders_risk_insurance_sf_of_gross_building_area", "gba"),
    63: ("Contractor Insurance", "contractor_insurance_sf_of_gross_building_area", "gba"),
    68: ("Title and Recording", "title_and_recording_sf_of_gross_building_area_r367", "gba"),
    73: ("School Impact Fees", "school_impact_fees_sf_of_gross_building_area", "gba"),
    74: ("Other Impact Fees", "other_impact_fees_sf_of_gross_building_area", "gba"),
    75: ("Building Permit, Plan Check, Other Fees",
         "building_permit_plan_check_other_fees_sf_of_gross_building_a", "gba"),
    76: ("Capital Fees (utility fee & hook-up charges)", "capital_fees_sf_of_gross_building_area", "gba"),
    77: ("Environmental Reports", "environmental_reports_sf_of_gross_building_area", "gba"),
    79: ("Market Study", "market_study_sf_of_gross_building_area", "gba"),
    80: ("TCAC App/Alloc/Monitor Fees", "tcac_app_alloc_monitor_fees_sf_of_gross_building_area", "gba"),
    81: ("CDLAC Fees", "cdlac_fees_sf_of_gross_building_area", "gba"),
    85: ("Financial Consultant", "financial_consultant_sf_of_gross_building_area", "gba"),
    86: ("Cost Cert / Accounting", "cost_cert_accounting_sf_of_gross_building_area", "gba"),
    87: ("Construction management services", "construction_management_services_sf_of_gross_building_area", "gba"),
    88: ("Partnership Start-up Fee", "partnership_start_up_fee_sf_of_gross_building_area", "gba"),
    100: ("Soft Cost Contingency", "soft_cost_contingency_sf_of_gross_building_area", "gba"),
}

# Other line labels (column D; section headers are in column C).
LABELS = {
    14: "Total Land Cost or Value", 15: "Total Building Cost or Value", 16: "Land profit",
    18: "Holding costs / predevelopment interest", 19: "Total Acquisition",
    26: "Structures - Commercial", 27: "GC - General Requirements", 30: "Public Art onsite",
    32: "Interior amenities (furnishings)", 33: "Solar", 34: "Construction - Other #5",
    35: "Construction - Other #6", 37: "Total Construction", 52: "Total Legal",
    55: "Construction Loan Interest", 56: "Bond Issuance", 57: "Origination Fee",
    60: "Lender Inspections", 64: "Total Construction Interest and Fees",
    67: "Loan Origination Fee", 69: "Other", 70: "Total Permanent Financing",
    78: "Appraisal", 82: "Marketing/Leasing", 83: "Relocation Expenses",
    84: "Site and Interior Furnishings", 89: "Entitlement Services, Buildilng Permit Expediting",
    90: "Other #1", 91: "Other #2", 92: "Other #3", 93: "Other #4", 94: "Other #5", 95: "Other #6",
    97: "Total Other Soft Costs", 104: "Capitalized Operating Reserve",
    105: "Capitalized Lease Up Reserve", 106: "Capitalized Section 8 Transition Reserve",
    107: "Capitalized Commercial / Lease Up Reserve", 108: "Capitalized Replacement Reserve",
    109: "Total Reserves", 113: "Paid Developer Fee", 114: "Deferred Dev Fee",
    115: "Total Developer Fee", 118: "Contingent accrued interest",
}
SECTIONS = {
    13: "Acquisition Costs", 21: "Hard Costs", 39: "Hard Cost Contingency", 41: "Architect Fees",
    42: "Survey and Engineering", 43: "Environmental Mitigation", 45: "Legal",
    54: "Construction Interest and Fees", 66: "Permanent Financing", 72: "Other Soft Costs",
    100: "Soft Cost Contingency", 103: "Reserves", 112: "Developer Fee",
    119: "Total Development Costs",
}

# Hardcoded line totals typed on the tab.
FIXED_COST = {16: 0, 18: 1e-06, 26: 0, 27: 0, 30: 0, 32: 1450000 * 0.5 * 0, 69: 0, 78: 0,
              83: 0, 84: 0, 89: 0, 105: 0, 106: 0, 107: 0, 108: 0}

# Typed eligible factors (column N); "1-V8" rows are handled in compute.
FIXED_N = {15: 0, 22: 0, 31: 1, 39: 1, 40: 1, 41: 1, 42: 1, 43: 1, 46: 0.5, 47: 0, 48: 0.5,
           49: 0, 50: 1, 51: 0, 57: 1, 58: 1, 59: 1, 60: 1, 61: 1, 62: 1, 63: 1,
           73: 1, 74: 1, 75: 1, 76: 1, 77: 1, 78: 1, 79: 1, 81: 0, 82: 0, 85: 1, 86: 0,
           87: 1, 88: 0, 89: 1, 100: 1}
NONRES_SCALED_N = (24, 25, 26, 27, 28, 29, 30, 32, 33, 34, 35)   # N = 1 - commercial %

# Predevelopment interest: the tab books a near-zero cost (1e-6) on row 18 but a
# hardcoded basis amount, written as N18 = 770,170.98 / H18.
PREDEV_INTEREST_BASIS = (452000 + 166395) * 0 + 770170.98

# Utility fee detail (EBMUD / PG&E worksheet typed on the tab, rows 127-152).
UTILITY_FEES = {127: 37365 * 2 * 1.05, 128: 0, 130: 0, 131: 0, 132: 63 * 2, 134: 156218 + 585818,
                135: 0, 136: 0, 137: 7739 * 0, 138: 30144 + 133039, 139: 17200 * 2, 140: 0,
                142: 30946 * 4, 143: 0, 144: 0, 145: 2300 * 2, 146: -17200 * 2, 147: -4710 * 2}
PGE_FEES = 100000
LAND_VALUE_PLUG = 2180000          # F169: "+2180000" typed into the land value formula

HARD_COST_PSF_KEYS = [  # Master Inputs rows 150-160, aligned with Program C13:C23
    "hard_cost_sf_rooftops", "hard_cost_sf_amenities", "hard_cost_sf_residential",
    "hard_cost_sf_4th_floor_parking", "hard_cost_sf_4th_floor_non_parking",
    "hard_cost_sf_3rd_floor_parking", "hard_cost_sf_3rd_floor_non_parking",
    "hard_cost_sf_2nd_floor_parking", "hard_cost_sf_2nd_floor_non_parking",
    "hard_cost_sf_ground_floor_parking", "hard_cost_sf_ground_floor",
]

ITEM_ROWS = ([14, 15, 16, 17, 18] + list(range(22, 36)) + [39, 41, 42, 43] + list(range(46, 52))
             + list(range(55, 64)) + [67, 68, 69] + list(range(73, 95)) + [100]
             + list(range(104, 109)) + [113, 114])


@register
class DevBudget(Sheet):
    name = "Dev Budget"

    def compute(self, ctx):
        mi = ctx.mi
        o = {}
        units = xl.n(ctx.ref("Program", "E64"))
        gba = xl.n(ctx.ref("Program", "C24"))
        lot = xl.n(mi["lot_size"])
        o["B2"] = f"{mi['project_name']}  |  DEV BUDGET"
        o["D7"] = mi["project_name"]
        for r, t in SECTIONS.items():
            o[f"C{r}"] = t
        for r, t in LABELS.items():
            o[f"D{r}"] = t

        # --- Residential / non-residential split.
        o["L11"] = xl.n(mi["non_residential_share_of_cost"])
        o["K11"] = 1 - o["L11"]
        k11, l11 = o["K11"], o["L11"]

        # --- Column H: total cost per line.
        H = {r: 0.0 for r in ITEM_ROWS}
        H.update(FIXED_COST)
        H[14] = xl.n(mi["land_cost_value"])
        H[15] = xl.n(mi["improvements_value"])
        for r, (label, key, basis) in PSF_LINES.items():
            if r not in SECTIONS:          # rows 41/42/100 carry their label in column C
                o[f"D{r}"] = label
            H[r] = xl.n(mi[key]) * (lot if basis == "lot" else gba)
        hard = sum(H[r] for r in range(22, 36))
        o["E39"] = mi["hard_cost_contingency_pct"]
        # Contingency excludes furnishings and demolition.
        H[39] = o["E39"] * (hard - H[32] - H[23])

        # Construction loan interest & fees (from Flow of Funds / Financing).
        H[55] = sum(xl.n(ctx.ref(FOF, a)) for a in ("D231", "D283", "D286"))
        fee_amt = ctx.col(FIN, "P139:P156")
        fee_type = ctx.col(FIN, "M139:M156")
        fee_cap = ctx.col(FIN, "L139:L156")

        def capitalized_fee(kind):
            return xl.SUMIFS(fee_amt, fee_type, kind, fee_cap, "Y")

        H[56] = capitalized_fee("Bond Issuance")
        H[57] = capitalized_fee("Construction Origination")
        H[60] = capitalized_fee("Lender Inspections")
        H[67] = capitalized_fee("Perm Origination")
        o["E82"] = mi["marketing_leasing_cost_per_unit"]
        H[82] = o["E82"] * units

        # Operating reserve: N months of opex + debt service (OCF year-1 figures, negative).
        o["E104"] = mi["capitalized_operating_reserve"]
        H[104] = o["E104"] / 12 * (-xl.n(ctx.ref(OCF, "J45")) - xl.n(ctx.ref(OCF, "J62")))

        # Developer fee: paid portion from LIHTC; deferred portion limited to what cash flow repays.
        H[113] = xl.n(ctx.ref(LIH, "E77"))
        H[114] = xl.IFERROR(lambda: min(
            xl.n(ctx.ref(LIH, "E78")),
            xl.SUM(ctx.col(OCF, "J64:X64")) + xl.n(ctx.ref(OCF, "M98")) + xl.n(ctx.ref(OCF, "I96"))), 0)

        # --- Program-based cost calcs (rows 155-178).
        prog = ctx.col("Program", "C13:C23")
        o["F164"] = xl.SUMPRODUCT([mi[k] for k in HARD_COST_PSF_KEYS], prog) / gba
        o["F160"] = xl.ROUND(o["F164"], 2)
        h37 = hard
        if xl.eq(ctx.ref("Sensitivities", "C34"), "yes"):
            # Parking SF costs carved out of eligible basis.
            pk = (mi["hard_cost_sf_4th_floor_parking"] * xl.n(ctx.ref("Program", "C16"))
                  + xl.n(ctx.ref("Program", "C18")) * mi["hard_cost_sf_3rd_floor_parking"]
                  + xl.n(ctx.ref("Program", "C20")) * mi["hard_cost_sf_2nd_floor_parking"]
                  + mi["hard_cost_sf_ground_floor_parking"] * xl.n(ctx.ref("Program", "C22")))
            o["D157"] = pk / (h37 + H[39])
        else:
            o["D157"] = 0
        o["V8"] = o["D157"]
        commercial_pct = o["V8"]

        # --- Column N: eligible factor.
        N = {r: 0.0 for r in ITEM_ROWS}
        N.update(FIXED_N)
        for r in NONRES_SCALED_N:
            N[r] = 1 - commercial_pct
        N[18] = PREDEV_INTEREST_BASIS / H[18]
        fof = {a: xl.n(ctx.ref(FOF, a)) for a in ("D215", "D222", "D229", "D231")}
        # Share of construction interest that accrues before placed-in-service (basis eligible).
        N[55] = (fof["D215"] + fof["D222"] + fof["D229"]) / fof["D231"]
        N[56] = xl.n(ctx.ref("Leasing", "D119"))
        lih_h163 = xl.n(ctx.ref(LIH, "H163"))
        dev_fee_n = ((xl.n(ctx.ref(LIH, "H162")) + xl.n(ctx.ref(LIH, "H111"))) / xl.n(ctx.ref(LIH, "E76"))
                     if lih_h163 > 0 else 1)
        N[113] = N[114] = dev_fee_n

        # --- Apportion each line.
        K, L, P, Q, R = {}, {}, {}, {}, {}
        for r in ITEM_ROWS:
            h = H[r]
            K[r] = h * k11
            L[r] = h * l11
            P[r] = N[r] * K[r] if r == 15 else 0.0
            if r in (49, 50):
                Q[r] = h * N[r]            # these two rows apply N to H rather than K
            elif r in range(90, 95):
                Q[r] = K[r]
            else:
                Q[r] = K[r] * N[r]
            if r == 80:
                R[r] = K[r]
            elif r in (113, 114):
                R[r] = xl.IFERROR(lambda: h - P[r] - Q[r], 0)
            else:
                R[r] = h - (P[r] + Q[r])

        # --- Subtotals.
        def total(rows, col):
            return xl.SUM([col[r] for r in rows])

        groups = {19: range(14, 19), 37: range(22, 36), 52: range(46, 52), 64: range(55, 64),
                  70: range(67, 70), 97: range(73, 95), 109: range(104, 109), 115: range(113, 115)}
        for t, rows in groups.items():
            H[t] = total(rows, H)
            K[t] = total(rows, K) if t in (52, 115) else H[t] * k11
            L[t] = total(rows, L) if t in (52, 115) else H[t] * l11
            P[t] = total(rows, P)
            Q[t] = total(rows, Q)
            R[t] = H[t] - (P[t] + Q[t]) if t in (70, 97, 109) else total(rows, R)
        # Q63..Q101 references in empty rows (Q104 = Q101*Q98 etc.) are all blank -> 0.

        top = (115, 109, 100, 97, 70, 64, 37, 19, 52, 41, 42, 43, 39)
        H[119] = sum(H[r] for r in top)
        K[119] = sum(K[r] for r in top)
        L[119] = sum(L[r] for r in top)
        Q[119] = sum(Q[r] for r in top)
        R[119] = sum(R[r] for r in top)
        # EXCEL-QUIRK: P119 sums P101 (an empty row) instead of P100 and omits P39/P52.
        P[119] = sum(P.get(r, 0.0) for r in (115, 109, 97, 70, 64, 37, 19, 41, 42, 43))

        # --- Column T: costs paid at conversion (deferred until perm).
        T = {15: 0, 18: xl.n(ctx.ref(SU, "E55")), 47: H[47], 49: R[49], 51: R[51], 67: H[67],
             68: H[68], 69: H[69], 80: H[80], 104: H[104], 105: H[105], 106: H[106],
             107: H[107], 108: H[108],
             113: xl.n(ctx.ref(FOF, "Y13")) + xl.n(ctx.ref(FOF, "Y15"))}
        T[119] = sum(T.values())

        # --- Non-residential costs (commercial share of hard costs, ex GC contingency).
        o["V23"] = 0
        o["V37"] = commercial_pct * (H[37] - H[31])
        o["V119"] = o["V23"] + o["V37"]

        # --- Write columns.
        all_rows = ITEM_ROWS + list(groups) + [119]
        for r in all_rows:
            o[f"H{r}"] = H[r]
            o[f"K{r}"] = K[r]
            o[f"L{r}"] = L[r]
            o[f"Q{r}"] = Q[r]
            o[f"R{r}"] = R[r]
            if r not in (113, 114, 119):
                o[f"I{r}"] = H[r] / units
        for r in [15] + list(groups) + [119]:
            o[f"P{r}"] = P[r]
        for r in ITEM_ROWS:
            if r in N and (r in FIXED_N or r in NONRES_SCALED_N or r in (18, 55, 56, 113, 114)):
                o[f"N{r}"] = N[r]
        o["N40"] = FIXED_N[40]
        for r, v in T.items():
            o[f"T{r}"] = v
        o["I70"] = sum(o[f"I{r}"] for r in (67, 68, 69))
        # EXCEL-QUIRK: per-unit total I119 omits the Legal subtotal (I52).
        o["I119"] = sum(o[f"I{r}"] for r in (115, 109, 100, 97, 70, 64, 37, 19, 41, 42, 43, 39))
        o["I115"] = H[115] / units
        # Check columns (Excel shows these as zero / rounding noise).
        for r in range(14, 119):
            if r in H:
                o[f"Y{r}"] = K[r] + L[r] - H[r]
                o[f"Z{r}"] = P.get(r, 0.0) + Q[r] + R[r] - K[r]

        # --- Column E/G: unit costs and rates shown beside the lines.
        o["E14"] = H[14] / units
        o["E15"] = H[15] / units
        o["E16"] = 0
        o["E18"] = mi["preferred_return_rate"]
        o["E25"] = o["F160"]
        o["G26"] = H[26] / 25075          # typed commercial SF
        o["E31"] = mi["gc_contingency_pct"]
        o["G31"] = H[31] / gba
        o["E57"] = mi["fee_constr_orig"]
        o["E58"] = mi["property_tax_rate_during_construction"]
        o["E59"] = mi["performance_bond_premium_pct_of_hard_cost"]
        o["E62"] = mi["builder_s_risk_insurance_pct"]
        o["E63"] = mi["contractor_insurance_pct_of_hard_cost"]
        o["E67"] = mi["fee_perm_orig"]
        unit_count = xl.SUM(ctx.col("Program", "E46:E56"))
        for r in (73, 74, 75, 76):
            o[f"E{r}"] = H[r] / unit_count
        o["E80"] = mi["tcac_app_allocation_monitoring_fee_per_restricted_unit"]
        o["E100"] = mi["soft_cost_contingency_pct"]
        o["E113"] = xl.IFERROR(lambda: (H[113] - T[113]) / H[113], 0)

        # --- Utility fee detail.
        for r, v in UTILITY_FEES.items():
            o[f"D{r}"] = v
        o["D150"] = sum(UTILITY_FEES.values())
        o["D152"] = PGE_FEES

        # --- Land value.
        o["E169"] = xl.n(mi["land_value_sf_of_lot"])   # no override typed in D169
        # EXCEL-QUIRK: land value has a hardcoded +$2,180,000 plug in the formula.
        o["F169"] = (xl.n(ctx.ref("Initial Screen", "G18")) * o["E169"]
                     - xl.n(mi["improvements_value"]) + LAND_VALUE_PLUG)
        o["F170"] = mi["improvements_value"]
        o["F171"] = o["F169"] + xl.n(o["F170"])

        # --- Cost metrics (from Sources & Uses).
        su = {a: xl.n(ctx.ref(SU, a)) for a in ("E43", "E36", "E23", "E13", "E14")}
        o["E176"] = su["E43"] / units
        o["E177"] = (su["E43"] - su["E36"] - su["E23"] - H[70] - H[51] - H[49] - H[47]) / units
        o["E178"] = (su["E13"] + su["E14"]) / xl.n(ctx.ref("Program", "C31"))
        return o
