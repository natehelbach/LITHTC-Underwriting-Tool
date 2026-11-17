"""Flow of Funds: the monthly development funding engine.

Months run across columns L..BH (month 0 = closing .. month 48).

  rows 10-19    milestone schedule (closing, construction start, 50% completion, completion,
                C of O, lease-up, stabilization, conversion, 8609) with the LIHTC equity and
                paid developer fee installment due at each milestone; cols Q..AB hold the
                installment summary tables
  rows 21-23    phase label, calendar date and month number per column
  rows 25-48    construction sources (from Sources & Uses / Financing)
  rows 56-166   monthly uses by Dev Budget line. Each line is drawn one of three ways:
                  even   - spread evenly over its start..end months (E/F)
                  curve  - "S-Curve" (truncated normal) or "Straightline" over start..end,
                           optionally with a deposit at closing and a retention released
                           at C of O + equity lag
                  CA     - fixed monthly construction-administration fee, rest at closing
                plus construction interest (row 97, from the interest block below) and the
                paid developer fee (row 164, paid at the equity funding months)
  row 167       total monthly uses
  rows 170-176  equity sources and the construction account (a surplus carries forward)
  rows 179-187  soft sources fund the shortfall first, leaving the remainder to fund
  rows 189-191  construction loans in order: tax-exempt, recycled bonds, taxable. At
                C of O + lag the taxable loan pre-funds all remaining uses to conversion
                (net of the C of O equity installment)
  rows 193-207  balance check and the conversion take-out
  rows 209-231  construction loan balances and interest (interest is a use: row 97)
  rows 233-287  soft source balances, accrued / must-pay interest

Circularity: the taxable pre-funding draw looks ahead at future months' uses, which
include future interest on that draw. Excel resolves this with iterative calculation;
here the month loop is re-run until the uses stop changing. The cross-tab circularity
(Dev Budget, Sources & Uses, Financing, LIHTC) is left to the engine (read via ctx).
"""
import datetime as dt
import math

from openpyxl.utils import column_index_from_string, get_column_letter

from uw import xl
from uw.engine import Sheet, register

DB = "Dev Budget"
SU = "Sources & Uses"
FIN = "Financing & Debt Sizing"
LIH = "LIHTC & Dev Fee"

FIRST_COL = 12          # column L holds month 0
N_MONTHS = 49           # L..BH = months 0..48
MONTHS = range(N_MONTHS)
MILESTONE_ROWS = range(10, 19)


def col(m: int) -> str:
    """Excel column of month m (0 -> L)."""
    return get_column_letter(FIRST_COL + m)


def month_of(letter: str) -> int:
    """Month shown in an Excel column of this tab (L -> 0)."""
    return column_index_from_string(letter) - FIRST_COL


def _norm_cdf(x, mu, sd):
    return 0.5 * (1.0 + math.erf((x - mu) / (sd * math.sqrt(2.0))))


def _same(a, b):
    """Excel '=' between two float totals in the check cells (equal to ~15 digits)."""
    return math.isclose(xl.n(a), xl.n(b), rel_tol=1e-12, abs_tol=1e-9)


def _curve(kind, m, start, end, length, sd, amount, minus=0.0, straight_amount=None):
    """Draw in month m of a line scheduled on an S-curve or a straight line over start..end.

    The S-curve gives month k (0-based from start) the normal-distribution mass between
    k and k+1, centred on the middle of the window and rescaled so the truncated tails
    still sum to 100%.
    """
    if not (start <= m <= end):
        return 0.0
    if kind == "S-Curve":
        k, mu = m - start, length / 2
        share = (_norm_cdf(k + 1, mu, sd) - _norm_cdf(k, mu, sd)) / (1 - 2 * _norm_cdf(0, mu, sd))
        return share * amount - minus
    if kind == "Straightline":
        return (amount if straight_amount is None else straight_amount) / length
    return 0.0


# ---------------------------------------------------------------- budget line layout
# Tab row -> Dev Budget row (label in DB column D, amount in DB column H).
DB_ROW = {
    56: 14, 57: 15, 58: 16, 59: 17, 60: 18,
    **{64 + i: 22 + i for i in range(11)},           # 64..74 <- DB 22..32
    **{97 + i: 55 + i for i in range(9)},            # 97..105 <- DB 55..63
    107: 67, 108: 68, 109: 69,
    **{116 + i: 46 + i for i in range(6)},           # 116..121 <- DB 46..51
    **{136 + i: 73 + i for i in range(17)},          # 136..152 <- DB 73..89
    **{157 + i: 104 + i for i in range(5)},          # 157..161 <- DB 104..108
}
# Lines whose label is typed on this tab but whose amount comes from Dev Budget column H.
DB_AMOUNT_ONLY = {79: 39, 87: 41, 88: 42, 89: 43, 154: 100}

EVEN_ROWS = (list(range(56, 63)) + list(range(81, 87)) + list(range(91, 96))
             + [98, 99, 101, 102, 103, 104, 105, 108, 109, 110, 112, 113, 114]
             + list(range(116, 127)) + list(range(128, 135))
             + list(range(136, 156)) + list(range(157, 163)) + [165, 166])
CURVE_ROWS = range(64, 80)
CA_ROWS = (87, 88, 89)
# Lines with no start/end typed (blank E/F).
NO_TIMING = {61, 62, *range(75, 79), *range(81, 87), *range(91, 96), 110, 112, 113, 114,
             153, 155, 159, 160, 161, 162, 166}

SOFT_ROWS = range(179, 186)                      # soft source 1..7 draws
SOFT_BLOCKS = (233, 240, 247, 254, 261, 268, 276)  # header row of each soft-source interest block
LOAN_BLOCKS = {189: 210, 190: 217, 191: 224}    # draw row -> header row of its interest block


@register
class FlowOfFunds(Sheet):
    name = "Flow of Funds"

    # Typed on the tab ----------------------------------------------------------
    STRUCTURES_CLOSE_DEPOSIT = 4_500_000        # K67: residential structures paid at closing
    STRUCTURES_RETENTION = 0.05                 # row 67: 5% held to C of O + lag
    GC_CONTINGENCY_RETENTION = 0.5              # row 73: 50% held to C of O + lag
    CA_MONTHLY_FEE = {87: 26000, 88: 10000, 89: 0}   # H87:H89
    FED_8609_INSTALLMENT = 250000               # S15
    BOND_FEE_SPLIT = {11: 300000, 12: 1550728, 13: 300000}   # AA11:AA13
    TAXES_EARLY = {"N": 0, "R": 0}              # N100 / R100 typed; rest split Z100 / AD100
    PERMIT_FEE_DEFERRED = 0                     # AD138 (N138 nets it out)
    SOFT_SOURCE_SLOTS = {35: ("Soft Source 1", 1)}   # F35 / G35: soft source # and % drawn
    MARKETING_MONTHS = (18, 21)                 # E145:F145
    COST_CERT_MONTHS = (18, 19)                 # E149:F149
    CURVES = {  # row: (start, end, distribution); None = tied to a milestone (see _curve_timing)
        64: (16, 18, "Straightline"), 65: (1, 1, "Straightline"), 66: (15, None, "S-Curve"),
        67: (1, None, "S-Curve"), 68: (1, None, "S-Curve"), 69: (1, None, "S-Curve"),
        70: (1, None, "Straightline"), 71: (1, None, "Straightline"),
        72: (14, None, "Straightline"), 73: (None, None, "Straightline"),
        74: (None, None, "Straightline"), 75: (0, 0, "Straightline"), 76: (0, 0, "Straightline"),
        77: (0, 0, "Straightline"), 78: (0, 0, "Straightline"), 79: (None, None, "Straightline"),
    }
    INTEREST_IDC_SHARE = ((1, 1.0), (3, 0.8), (5, 0.6), (7, 0.4))  # months after completion -> share

    def compute(self, ctx):
        o = {"C2": f"{ctx.mi['project_name']}  |  FLOW OF FUNDS"}
        self._milestones(ctx, o)
        self._equity_schedule(ctx, o)
        self._timeline(ctx, o)
        self._sources(ctx, o)
        draws = self._budget_lines(ctx, o)
        self._waterfall(o, draws)
        o["F7"] = "** SOURCES SHORT **" if xl.ROUND(o["D48"] - o["D167"], 0) < 0 else "BALANCED"
        o["D168"] = o["D167"] - xl.n(ctx.ref(DB, "I113"))
        return o

    # ------------------------------------------------------------ milestones
    def _milestones(self, ctx, o):
        mi = ctx.mi
        # Months between milestones (E) and the cumulative month of each (F).
        o["E10"] = 0
        o["E11"] = mi["construction_start"]
        o["E13"] = xl.n(ctx.ref("Sensitivities", "C26"))               # construction months
        o["E12"] = xl.ROUND((o["E13"] - o["E11"]) / 2, 0)         # to 50% completion
        o["E14"] = mi["c_of_o_cost_cert"]
        o["E15"] = xl.n(ctx.ref("Sensitivities", "C27")) - 4      # lease-up
        o["E16"] = mi["stabilization_period"]
        o["E17"] = mi["conversion"]
        o["E18"] = mi["8609"]
        o["F10"] = o["E10"]
        o["F11"] = o["E11"]
        o["F12"] = o["F11"] - 1 + o["E12"]
        o["F13"] = o["E13"]
        o["F14"] = o["F13"] + o["E14"]
        o["F15"] = o["F13"] + o["E15"]
        o["F16"] = o["F15"] + o["E16"]
        o["F17"] = o["F16"] + o["E17"]
        o["F18"] = o["F17"] + o["E18"]
        o["E19"] = xl.SUM([o[f"E{r}"] for r in range(13, 19)]) - o["E14"]
        o["F19"] = o["F18"] - o["E19"]
        # Equity funding lag after each installment milestone (K).
        o["K10"] = mi["equity_funding_lag_after_closing"]
        o["K14"] = mi["equity_funding_lag_after_c_of_o_cost_cert"]
        o["K17"] = mi["equity_funding_lag_after_conversion"]
        o["K18"] = mi["equity_funding_lag_after_8609"]

        self.close, self.start, self.complete = o["F10"], o["F11"], o["F13"]
        self.cofo, self.leaseup, self.stab = o["F14"], o["F15"], o["F16"]
        self.conv, self.f8609 = o["F17"], o["F18"]
        self.half_months = o["E12"]
        self.cofo_funding = o["F13"] + o["K14"]     # retention release / taxable pre-fund month
        self.conv_funding = o["F17"] + o["K17"]     # conversion take-out month

    def _equity_schedule(self, ctx, o):
        mi = ctx.mi
        rows = MILESTONE_ROWS
        # Paid developer fee installments (X/Y).
        o["X10"] = xl.n(ctx.ref(SU, "E24")) - xl.n(ctx.ref(SU, "E36"))   # total fee - deferred
        o["X11"] = mi["paid_developer_fee_pct_at_closing_award"]
        o["X12"] = mi["paid_developer_fee_pct_at_tco"]
        o["O18"] = mi["federal_equity_at_8609"]
        o["X15"] = xl.div(o["O18"], o["X10"], 0.0)   # (0 only before the model has settled)
        o["X13"] = 1 - o["X11"] - o["X12"] - o["X15"]
        inst = (11, 12, 13, 15)
        for r in inst:
            o[f"Y{r}"] = o["X10"] * o[f"X{r}"]
        o["X16"] = xl.SUM([o[f"X{r}"] for r in inst])
        o["Y10"] = "no error" if _same(o["X10"], xl.SUM([o[f"Y{r}"] for r in inst])) else "error"

        # Federal (N/O) and state (L/M) LIHTC equity pay-in by milestone; J = total.
        o["R10"] = xl.n(ctx.ref(LIH, "G58"))      # total federal equity
        o["T10"] = xl.n(ctx.ref(LIH, "G64"))      # total state equity
        o["O19"], o["M19"] = o["R10"], o["T10"]
        state = {r: 0 for r in rows}
        state[14] = mi["state_equity_pay_in_pct_at_c_of_o"]
        state[17] = 1 - xl.SUM([state[r] for r in rows if r != 17])
        fed = {r: 0 for r in rows}
        fed[10] = mi["federal_equity_pay_in_pct_at_closing"]
        fed[14] = mi["federal_equity_pay_in_pct_at_c_of_o_cost_cert"]
        fed[18] = xl.div(o["O18"], o["O19"], 0.0)
        fed[17] = 1 - fed[18] - xl.SUM([fed[r] for r in range(10, 17)])
        for r in rows:
            o[f"L{r}"], o[f"N{r}"] = state[r], fed[r]
            # EXCEL-QUIRK: M18 is a typed 0, so a state 8609 installment would never be funded.
            o[f"M{r}"] = 0 if r == 18 else state[r] * o["M19"]
            if r != 18:
                o[f"O{r}"] = fed[r] * o["O19"]
            o[f"J{r}"] = o[f"M{r}"] + o[f"O{r}"]
        o["L19"] = 0
        o["N19"] = xl.SUM([fed[r] for r in rows])
        o["J19"] = xl.SUM([o[f"J{r}"] for r in rows])
        for r in rows:
            o[f"I{r}"] = xl.div(o[f"J{r}"], o["J19"], 0.0)
        o["I19"] = xl.SUM([o[f"I{r}"] for r in rows])

        # Paid developer fee by milestone (G = share of paid fee, H = amount).
        o["D164"] = o["X10"]
        g = {r: 0 for r in rows}
        g[10], g[17] = o["X11"], o["X13"]
        o["H18"] = o["Y15"]
        g[18] = xl.div(o["H18"], o["D164"], 0.0)
        for r in rows:
            o[f"G{r}"] = g[r]
            if r != 18:
                o[f"H{r}"] = g[r] * o["D164"]
        o["G19"] = xl.SUM([o[f"G{r}"] for r in rows])
        o["H19"] = xl.SUM([o[f"H{r}"] for r in rows])

        # Installment summary (Q..W): federal R/S, state T/U, total V/W.
        o["R11"], o["R12"], o["R15"] = o["N10"], o["N14"], o["N18"]
        o["R13"] = 1 - o["R11"] - o["R12"] - o["R15"]
        o["T11"], o["T12"], o["T13"], o["T15"] = o["L10"], o["L14"], o["L17"], o["L18"]
        for r in (11, 12, 13):
            o[f"S{r}"] = o["R10"] * o[f"R{r}"]
            o[f"U{r}"] = o["T10"] * o[f"T{r}"]
        o["S15"], o["U15"] = self.FED_8609_INSTALLMENT, 0
        for r in inst:
            o[f"V{r}"] = o[f"S{r}"] + o[f"U{r}"]
        o["V10"] = xl.SUM([o[f"V{r}"] for r in inst])
        for r in inst:
            o[f"W{r}"] = xl.div(o[f"V{r}"], o["V10"], 0.0)
        o["W10"] = xl.SUM([o[f"W{r}"] for r in inst])
        for c in "RTW":
            o[f"{c}16"] = xl.SUM([o[f"{c}{r}"] for r in inst])
        o["S10"] = "no error" if _same(o["R10"], xl.SUM([o[f"S{r}"] for r in inst])) else "error"
        o["U10"] = "no error" if _same(o["T10"], xl.SUM([o[f"U{r}"] for r in inst])) else "error"
        # Bond fee split (AA/AB).
        o["AA15"] = xl.SUM(list(self.BOND_FEE_SPLIT.values()))
        for r, v in self.BOND_FEE_SPLIT.items():
            o[f"AA{r}"] = v
            o[f"AB{r}"] = v / o["AA15"]

    # ------------------------------------------------------------ month headers
    def _timeline(self, ctx, o):
        date = None
        for m in MONTHS:
            c = col(m)
            for r in (23, 54, 169):              # month number (three copies on the tab)
                o[f"{c}{r}"] = m
            if m == 0:
                continue
            # Month 1 is the construction loan closing date; then the 1st of each month.
            date = ctx.mi["constr_close"] if m == 1 else xl.EOMONTH(date, 0) + dt.timedelta(days=1)
            o[f"{c}22"] = date
            if m <= self.complete:
                phase = "Construction"
            elif m <= self.leaseup:
                phase = "Lease-up"
            elif m <= self.stab:
                phase = "Stabilization"
            elif m == self.conv:
                phase = "Conversion"
            elif m == self.f8609:
                phase = "8609"
            else:
                phase = ""
            o[f"{c}21"] = phase

    # ------------------------------------------------------------ sources
    def _sources(self, ctx, o):
        o["D25"] = ctx.ref(SU, "E29")     # tax credit equity during construction
        o["D26"] = ctx.ref(SU, "E33")     # expenses paid at conversion
        o["D27"] = ctx.ref(SU, "E40")     # GP contributions
        o["D28"] = ctx.ref(SU, "E36")     # deferred developer fee
        o["D30"] = ctx.ref(SU, "E30")     # tax-exempt construction loan
        o["E30"] = ctx.ref(FIN, "D178")
        o["D31"] = ctx.ref(SU, "E32")     # recycled bonds loan
        o["E31"] = o["E30"]
        o["E32"] = ctx.ref(FIN, "D179")
        # Soft source slot (row 35 is the only one wired up).
        o["C35"] = ctx.ref(FIN, "E12")
        o["D35"] = ctx.ref(SU, "E35")
        o["E35"] = ctx.ref(FIN, "E23")
        o["H35"] = "Must Pay" if xl.eq(ctx.ref(FIN, "E24"), "Current Pay") else "Accrued"
        o["D172"] = ctx.ref(SU, "E55")    # other equity
        o["D194"] = ctx.ref(SU, "E47")    # permanent loan
        # The source table: soft source # -> name, amount, rate, % drawn, pay type.
        self.soft_sources = {num: {"name": o[f"C{r}"], "amount": o[f"D{r}"], "rate": o[f"E{r}"],
                                   "pct": pct, "pay": o[f"H{r}"]}
                             for r, (num, pct) in self.SOFT_SOURCE_SLOTS.items()}

    # ------------------------------------------------------------ uses by budget line
    def _budget_lines(self, ctx, o):
        """Monthly draw of every budget line except construction interest (row 97)."""
        draws = {}

        def amount(r):
            if r in DB_ROW:
                o[f"C{r}"] = ctx.ref(DB, f"D{DB_ROW[r]}")
                o[f"D{r}"] = ctx.ref(DB, f"H{DB_ROW[r]}")
            elif r in DB_AMOUNT_ONLY:
                o[f"D{r}"] = ctx.ref(DB, f"H{DB_AMOUNT_ONLY[r]}")
            elif r == 165:
                o["D165"] = o["D28"]                 # deferred developer fee
            return xl.n(o.get(f"D{r}"))

        def timing(r, start, end):
            if r not in NO_TIMING:
                o[f"E{r}"], o[f"F{r}"] = start, end
            return start, end

        close, comp, conv = self.close, self.complete, self.conv
        cofo_last = comp + o["K14"] - 1             # last month before the C of O funding
        after_conv = conv + 1

        # --- even spread over start..end
        even_timing = {
            102: (1, cofo_last),                                   # lender inspections
            108: (after_conv, after_conv), 109: (after_conv, conv),  # (109 is never drawn)
            117: (after_conv, after_conv), 119: (after_conv, after_conv),
            121: (after_conv, after_conv), 143: (after_conv, after_conv),
            157: (after_conv, after_conv), 158: (after_conv, after_conv),
            139: (close + 4, close + 4),                           # utility capital fees
            145: self.MARKETING_MONTHS, 149: self.COST_CERT_MONTHS,
            147: (comp - 5, comp), 150: (self.start, comp),
            **{r: (close, close) for r in (116, 120, *range(122, 127), *range(128, 135),
                                           136, 137, 140, 141, 142, 144)},
        }
        for r in EVEN_ROWS:
            d = amount(r)
            s, e = timing(r, *even_timing.get(r, (0, 0)))
            draws[r] = [d / (e - s + 1) if s <= m <= e else 0.0 for m in MONTHS]
        # Building permit fees: month 2 nets out the amount moved to month 18 (typed 0).
        draws[138][month_of("N")] -= self.PERMIT_FEE_DEFERRED
        draws[138][month_of("AD")] = self.PERMIT_FEE_DEFERRED

        # --- perm loan origination: half spread over E..F, half at conversion funding
        d = amount(107)
        s, e = timing(107, 0, 0)
        draws[107] = [((d / 2) / (e - s + 1) if s <= m <= e else 0.0)
                      + (d / 2 if m == self.conv_funding else 0.0) for m in MONTHS]

        # --- property taxes during construction: typed layout. Two typed early payments,
        # the rest split between months 14 and 18; later months use the even-spread formula
        # (zeroed in month 19). Months not listed are blank on the tab.
        d = amount(100)
        s, e = timing(100, 1, comp)
        tx = [None] * N_MONTHS
        for c, v in self.TAXES_EARLY.items():
            tx[month_of(c)] = v
        tx[month_of("Z")] = tx[month_of("AD")] = (d - sum(self.TAXES_EARLY.values())) / 2
        for m in range(month_of("AE"), N_MONTHS):
            tx[m] = d / (e - s + 1) if s <= m <= e else 0.0
        tx[month_of("AE")] *= 0
        draws[100] = tx

        # --- construction administration: fixed monthly fee, remainder at closing (month 0)
        ca_timing = {87: (self.start, cofo_last), 88: (1, cofo_last), 89: (self.start, 5)}
        for r in CA_ROWS:
            d = amount(r)
            s, e = timing(r, *ca_timing[r])
            fee = o[f"H{r}"] = self.CA_MONTHLY_FEE[r]
            # EXCEL-QUIRK: row 89 sizes its closing remainder on $F$13 rather than its own F89.
            billed_through = comp if r == 89 else e
            draws[r] = [d - fee * billed_through if m == 0 else (fee if s <= m <= e else 0.0)
                        for m in MONTHS]

        # --- hard costs on an S-curve / straight line
        steep = o["H51"] = ctx.mi["s_curve_steepness"]
        auto = {66: (None, comp), 67: (None, comp + 6), 68: (None, comp), 69: (None, comp),
                70: (None, comp), 71: (None, comp), 72: (None, comp),
                74: (comp - 3, comp), 79: (self.half_months, comp)}
        auto[73] = (self.half_months, auto[67][1])     # GC contingency: 50% completion .. F67
        for r in CURVE_ROWS:
            s, e, kind = self.CURVES[r]
            s = s if s is not None else auto[r][0]
            e = e if e is not None else auto[r][1]
            d = amount(r)
            timing(r, s, e)
            o[f"G{r}"] = kind
            length = o[f"H{r}"] = e - s + 1
            sd = o[f"I{r}"] = length / steep
            if r == 67:      # deposit at closing; 95% on the curve, 5% retention at C of O + lag
                dep = o["K67"] = self.STRUCTURES_CLOSE_DEPOSIT
                net, ret = d - dep, self.STRUCTURES_RETENTION
                draws[r] = [dep if m == 0 else
                            _curve(kind, m, s, e, length, sd, net) * (1 - ret)
                            + (net * ret if m == self.cofo_funding else 0.0) for m in MONTHS]
            elif r == 73:    # 50% on the curve, 50% retention at C of O + lag
                # EXCEL-QUIRK: L73 uses itself as the closing "deposit" (a self-reference that
                # resolves to 0), and its S-curve branch subtracts that deposit every month
                # instead of scaling (D-L). Kept as written; this line is Straightline here.
                dep, ret = 0.0, self.GC_CONTINGENCY_RETENTION
                net = d - dep
                draws[r] = [_curve(kind, m, s, e, length, sd, d, minus=dep, straight_amount=net)
                            * (1 - ret) + (net * ret if m == self.cofo_funding else 0.0)
                            for m in MONTHS]
            else:
                draws[r] = [_curve(kind, m, s, e, length, sd, d) for m in MONTHS]

        amount(97)   # construction interest: amount from Dev Budget, monthly from the interest block

        # --- paid developer fee: each milestone's installment in its equity funding month
        sched = [(o[f"F{r}"] + xl.n(o.get(f"K{r}")), o[f"H{r}"]) for r in MILESTONE_ROWS]
        draws[164] = [xl.SUM([h for f, h in sched if f == m]) for m in MONTHS]
        return draws

    # ------------------------------------------------------------ funding waterfall
    def _waterfall(self, o, draws):
        D = lambda r: xl.n(o.get(f"D{r}"))

        # Equity in by month: each installment lands in its milestone month + lag.
        installments = [(self.close + o["K10"], o["J10"]), (self.cofo_funding, o["J14"]),
                        (self.conv_funding, o["J17"]), (self.f8609 + o["K18"], o["J18"])]
        equity = [xl.SUM([j for f, j in installments if f == m]) for m in MONTHS]
        o["D29"] = 0
        o["D170"], o["D171"], o["D173"], o["D174"] = o["J19"], o["D27"], o["D28"], o["D29"]
        draws[97] = [0.0] * N_MONTHS              # placeholder; interest is added per pass
        upfront = {171: xl.n(o["D171"]), 172: xl.n(o.get("D172")), 173: xl.n(o["D173"]),
                   174: xl.n(o["D174"])}           # funded at closing (month 0)

        # Soft source draw limits (rows 179-185), looked up by soft source #.
        for k, r in enumerate(SOFT_ROWS, start=1):
            src = self.soft_sources.get(f"Soft Source {k}")
            o[f"B{r}"] = f"Soft Source {k}"
            if r != 179:
                o[f"C{r}"] = src["name"] if src else ""
            o[f"D{r}"] = src["amount"] if src else ""
            o[f"E{r}"] = src["pct"] if src else ""
            o[f"F{r}"] = xl.n(src["pct"]) * xl.n(src["amount"]) if src else 0

        # Loan names are typed in C30:C32 and linked down to C189:C191.
        o["C189"], o["D189"] = "Tax Exempt Construction Loan", o["D30"]
        o["C190"], o["D190"] = "Recycled Bonds Loan", o["D31"]
        o["C191"] = "Taxable Construction Loan"
        rates = {189: xl.n(o["E30"]), 190: xl.n(o["E31"]), 191: xl.n(o["E32"])}
        caps = {189: D(30), 190: D(31)}

        # Iterate the month loop until the look-ahead uses settle.
        guess = [xl.SUM([v[m] for v in draws.values() if v[m] is not None]) for m in MONTHS]
        for _ in range(200):
            g = self._run_months(o, draws, equity, upfront, caps, rates, guess)
            change = max(abs(a - b) for a, b in zip(g["uses"], guess))
            guess = g["uses"]
            if change < 1e-8:
                break
        g = self._run_months(o, draws, equity, upfront, caps, rates, guess)
        draws[97] = g["interest"]

        # ---- uses grid and checks
        o["E97"] = o["F97"] = "n/a"
        for r, vals in draws.items():
            for m in MONTHS:
                if vals[m] is not None:
                    o[f"{col(m)}{r}"] = vals[m]
            total = xl.SUM([v for v in vals if v is not None])
            o[f"J{r}"] = "OK" if _same(total, D(r)) else "Error"
        o["D167"] = xl.SUM([o.get(f"D{r}") for r in range(55, 167)])
        o["J167"] = "OK" if _same(xl.SUM(g["uses"]), o["D167"]) else "Error"

        # ---- sources grid
        for m in MONTHS:
            c = col(m)
            o[f"{c}167"] = g["uses"][m]
            o[f"{c}170"] = equity[m]
            o[f"{c}176"] = g["acct"][m]
            o[f"{c}187"] = g["remain"][m]
            for r in (*SOFT_ROWS, 189, 190, 191):
                o[f"{c}{r}"] = g[r][m]
            if m > 0:
                o[f"{c}48"] = 0                     # no monthly sources typed in rows 25-47
        for r, v in upfront.items():
            o[f"L{r}"] = v
        o["J170"] = "OK" if _same(xl.SUM(equity[:month_of("AZ") + 1]), o["D170"]) else "Error"
        for r in (171, 172, 173, 174):
            o[f"J{r}"] = "OK" if _same(upfront[r], D(r)) else "Error"
        for r in SOFT_ROWS:
            o[f"J{r}"] = "OK" if _same(xl.SUM(g[r]), o[f"F{r}"]) else "Error"
        # EXCEL-QUIRK: the J170 / J189:J191 checks only sum through column AZ (month 40).
        for r in (189, 190, 191):
            o[f"J{r}"] = "OK" if _same(xl.SUM(g[r][:month_of("AZ") + 1]), o.get(f"D{r}")) else "Error"

        # Taxable loan size = draws before the C of O pre-funding month; E = all draws.
        # (J191 therefore reads "Error" in Excel too.)
        o["D191"] = xl.SUM([g[191][m] for m in MONTHS if m < self.cofo_funding])
        o["J191"] = "OK" if _same(xl.SUM(g[191][:month_of("AZ") + 1]), o["D191"]) else "Error"
        o["E191"] = xl.SUM(g[191])
        o["D32"] = o["D191"]
        o["D48"] = xl.SUM([D(r) for r in range(25, 48)])

        # ---- balance remaining and conversion take-out
        o["D193"] = o["D48"] - o["D167"]
        o["E193"] = "SHORT - fund the gap" if xl.ROUND(o["D193"], 0) < 0 else "OK: Sources cover Uses"
        repaid = xl.SUMIF(g[191], "<0")
        for m in MONTHS:
            c = col(m)
            funded = g["remain"][m] + g[189][m] + g[190][m] + g[191][m]
            o[f"{c}193"] = funded + repaid if m > self.cofo else funded
            o[f"{c}194"] = xl.n(o["D194"]) if m == self.conv_funding else 0.0
            at_conv = m == self.conv_funding
            for k, r in enumerate(range(195, 202)):
                sr = 179 + k
                if at_conv:
                    o[f"{c}{r}"] = xl.IFERROR(lambda sr=sr: xl.n(_num(o[f"D{sr}"])) - xl.n(o[f"F{sr}"]), 0)
                else:
                    o[f"{c}{r}"] = ""
            for r, loan in ((202, 189), (203, 190), (204, 191)):
                o[f"{c}{r}"] = -g["bal"][loan][m] if at_conv else ""
            if at_conv:
                o[f"{c}205"] = xl.SUM([o[f"{c}{r}"] for r in range(193, 205)])
                # EXCEL-QUIRK: L205-(R167-R170) has relative R refs, so it compares against the
                # uses/equity six months later rather than a fixed column.
                ahead = m + 6
                u = g["uses"][ahead] if ahead < N_MONTHS else 0.0
                e = equity[ahead] if ahead < N_MONTHS else 0.0
                o[f"{c}207"] = o[f"{c}205"] - (u - e)
            else:
                o[f"{c}205"] = o[f"{c}207"] = ""
        for r in range(195, 202):
            o[f"B{r}"] = f"Soft Source {r - 194}"
        o["C202"], o["C203"], o["C204"] = o["C189"], o["C190"], o["C191"]

        # ---- construction loan interest blocks
        idc_total = [0.0] * N_MONTHS
        for loan, h in LOAN_BLOCKS.items():
            o[f"C{h}"] = o[f"C{loan}"]
            o[f"D{h}"] = o[f"E{30 + (loan - 189)}"]
            for m in MONTHS:
                c = col(m)
                if m > 0:
                    o[f"{c}{h + 1}"] = g["bal"][loan][m - 1]
                o[f"{c}{h + 2}"] = g[loan][m]
                o[f"{c}{h + 3}"] = g["bal"][loan][m]
                o[f"{c}{h + 4}"] = g["int"][loan][m]
                o[f"{c}{h + 5}"] = g["int"][loan][m] * self._idc_share(m)
            o[f"D{h + 4}"] = xl.SUM(g["int"][loan])
            o[f"D{h + 5}"] = xl.SUM([o[f"{col(m)}{h + 5}"] for m in MONTHS])
        for m in MONTHS:
            o[f"{col(m)}231"] = g["interest"][m]
        o["D231"] = xl.SUM(g["interest"])

        # ---- soft source balances and interest
        self._soft_interest(o, g)

    def _idc_share(self, m):
        """Share of a month's construction interest counted as interest during construction."""
        if m < self.start or m > self.complete + 7:
            return 0
        for after, share in self.INTEREST_IDC_SHARE:
            if m <= self.complete + after:
                return share
        return 0.4

    def _run_months(self, o, draws, equity, upfront, caps, rates, uses_ahead):
        """One pass of the monthly funding waterfall."""
        z = lambda: [0.0] * N_MONTHS
        g = {"uses": z(), "interest": z(), "acct": z(), "remain": z(),
             **{r: z() for r in (*SOFT_ROWS, 189, 190, 191)},
             "bal": {r: z() for r in (189, 190, 191)}, "int": {r: z() for r in (189, 190, 191)}}
        bal = {189: 0.0, 190: 0.0, 191: 0.0}
        soft_drawn = {r: 0.0 for r in SOFT_ROWS}
        carry = 0.0
        for m in MONTHS:
            # Interest accrues monthly on last month's balance, construction start..conversion+1.
            accrues = self.start <= m <= self.conv + 1
            for r in bal:
                g["int"][r][m] = bal[r] * rates[r] / 12 if accrues else 0.0
            g["interest"][m] = g["int"][189][m] + g["int"][190][m] + g["int"][191][m]
            uses = xl.SUM([v[m] for v in draws.values() if v[m] is not None]) + g["interest"][m]
            g["uses"][m] = uses

            # Construction account: equity in, uses out, any surplus carried forward.
            sources = equity[m] + (xl.SUM(list(upfront.values())) if m == 0 else 0.0)
            acct = g["acct"][m] = sources - uses + max(0.0, carry)
            carry = acct

            # Soft sources fund the shortfall first.
            for r in SOFT_ROWS:
                drawn_before = xl.SUM(g[r][:m])
                need = -acct - xl.SUM([g[x][m] for x in SOFT_ROWS if x < r])
                g[r][m] = max(min(xl.n(o[f"F{r}"]) - drawn_before, need), 0.0)
            remain = g["remain"][m] = acct + xl.SUM([g[r][m] for r in SOFT_ROWS])

            # Construction loans, in order.
            te = max(min(caps[189] - xl.SUM(g[189][:m]), -remain), 0.0)
            rb = max(min(caps[190] - xl.SUM(g[190][:m]), -remain - te), 0.0)
            if m == self.cofo_funding:
                # Pre-fund all uses through conversion, net of the C of O equity installment.
                # EXCEL-QUIRK: the conversion month's uses are counted twice
                # (SUMIFS ">month, <=F17" plus SUMIFS "=F17").
                tx = (-o["J14"] + xl.SUM([uses_ahead[k] for k in MONTHS if m < k <= self.conv])
                      + xl.SUM([uses_ahead[k] for k in MONTHS if k == self.conv]))
            elif m == self.cofo:
                # EXCEL-QUIRK: in column L the "prior draws" range K191:$L191 includes the cell itself.
                tx = max(-xl.SUM(g[191][:m]), -remain)
            else:
                tx = max(-(remain + te + rb), 0.0)
            for r, v in ((189, te), (190, rb), (191, tx)):
                g[r][m] = v
                bal[r] += v
                g["bal"][r][m] = bal[r]
        return g

    def _soft_interest(self, o, g):
        pay_totals = {}
        for k, h in enumerate(SOFT_BLOCKS, start=1):
            name = f"Soft Source {k}"
            src = self.soft_sources.get(name)
            o[f"B{h}"] = name
            o[f"C{h}"] = src["name"] if src else ("" if k == 1 else "N/A")
            rate = o[f"D{h}"] = xl.n(src["rate"]) if src else 0
            pay = src["pay"] if src else "N/A"
            o[f"B{h + 4}"] = o[f"B{h + 5}"] = pay
            draw_row = 178 + k
            bal, ints, idc = 0.0, [], []
            for m in MONTHS:
                c = col(m)
                if m > 0:
                    o[f"{c}{h + 1}"] = bal
                draw = g[draw_row][m]
                o[f"{c}{h + 2}"] = draw
                bal += draw
                o[f"{c}{h + 3}"] = bal
                # Soft interest accrues on this month's ending balance, construction..conversion.
                i = bal * rate / 12 if self.start <= m <= self.conv else 0.0
                ints.append(i)
                idc.append(i if self.start <= m <= self.complete else 0.0)
                o[f"{c}{h + 4}"] = i
                o[f"{c}{h + 5}"] = idc[-1]
            o[f"D{h + 4}"] = xl.SUM(ints)
            o[f"D{h + 5}"] = xl.SUM(idc)
            for kind, series in (("Interest", ints), ("IDC", idc)):
                tot = pay_totals.setdefault((pay.lower() if isinstance(pay, str) else pay, kind),
                                            [0.0] * N_MONTHS)
                for m in MONTHS:
                    tot[m] += series[m]
        for r, pay, kind in ((283, "Must Pay", "Interest"), (284, "Must Pay", "IDC"),
                             (286, "Accrued", "Interest"), (287, "Accrued", "IDC")):
            series = pay_totals.get((pay.lower(), kind), [0.0] * N_MONTHS)
            for m in MONTHS:
                o[f"{col(m)}{r}"] = series[m]
            o[f"D{r}"] = xl.SUM(series)


def _num(v):
    """Excel arithmetic on a cell: "" raises #VALUE! (caught by IFERROR)."""
    if v == "":
        raise TypeError("#VALUE!")
    return v
