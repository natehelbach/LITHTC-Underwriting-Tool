"""Lookup-library tabs: utility allowances, LIHTC/TBL county tables, federal minimum opex.

These tabs are mostly typed reference data (extracted to data/site8/tables/*.json by
tools/extract_tables.py). Only the formula cells are computed here:

Utility Allowances   - CUAC-reduced allowance schedule (actual x (1 - CUAC reduction)) and the
                       project's 5-row allowance block looked up by Master Inputs city.
LIHTC-TBL Tables     - county TBL / AMI / FMR analysis, and the project county's 100% AMI max
                       rents (trended) that Unit Mix Summary turns into max rents by AMI.
Federal Minimum Opex - TCAC operating-cost minimums; no formulas besides the title.
"""
from __future__ import annotations

import json
import statistics
from functools import lru_cache

from uw import xl
from uw.engine import DATA, Sheet, register

BEDROOM_COLS = ["Studio", "1 BR", "2 BR", "3 BR", "4 BR"]


@lru_cache(maxsize=None)
def _table(name: str, deal: str = "site8") -> dict:
    return json.loads((DATA / deal / "tables" / f"{name}.json").read_text(encoding="utf-8"))


def _title(ctx, suffix: str) -> str:
    return f"{ctx.mi['project_name']}  |  {suffix}"


@register
class UtilityAllowances(Sheet):
    name = "Utility Allowances"

    SCHEDULE_COLS = "DEFGH"          # CUAC-reduced allowance, Studio..4BR
    LOOKUP_COLS = ["Z", "AA", "AB", "AC", "AD", "AE"]   # utility label + Studio..4BR
    LOOKUP_ROWS = 5                  # Z8:AE12 = the city's five utility rows
    INDEX_LAST_ROW = 200             # Excel INDEX/MATCH range is rows 1..200

    def compute(self, ctx):
        o = {"B2": _title(ctx, "UTILITY ALLOWANCES")}
        tbl = _table("utility_allowances")
        application_case = xl.eq(ctx.ref("Sensitivities", "E10"), "yes")

        # CUAC (California Utility Allowance Calculator) reduction from the active scenario.
        o["AB6"] = ctx.ref("Sensitivities", "E23" if application_case else "D23")
        o["B9"] = keep = 1 - xl.n(o["AB6"])

        # Schedule: allowance actually charged = actual allowance x (1 - CUAC reduction).
        zeroed = set(tbl["zero_in_application_case"])
        by_row = {}
        for rec in tbl["rows"]:
            r = rec["row"]
            actual = []
            for col, v in zip("IJKLM", rec["actual"]):
                if f"{col}{r}" in zeroed:
                    # Oakland water heating: typed as IF(Sensitivities!E10="yes",0,N).
                    v = 0 if application_case else v
                    o[f"{col}{r}"] = v
                actual.append(v)
            reduced = [xl.n(v) * keep for v in actual]
            for col, v in zip(self.SCHEDULE_COLS, reduced):
                o[f"{col}{r}"] = v
            by_row[r] = (rec["utility"], reduced)

        # Project block: first schedule row whose city matches Master Inputs city, then the
        # next four rows (one per utility). Rows 1-10 of column B hold no city names.
        city = ctx.mi["city"]
        first = next((rec["row"] for rec in tbl["rows"]
                      if rec["row"] <= self.INDEX_LAST_ROW and xl.eq(rec["city"], city)), None)
        for k in range(self.LOOKUP_ROWS):
            out_row = 8 + k
            src = None if first is None else first + k
            if src is None or src > self.INDEX_LAST_ROW:
                vals = [None] * 6          # Excel: #N/A / #REF!
            else:
                label, reduced = by_row[src]
                vals = [label if label is not None else 0] + reduced
            for col, v in zip(self.LOOKUP_COLS, vals):
                o[f"{col}{out_row}"] = v
        # Two spare typed rows under the lookup block (labels only).
        o["Z13"] = o["Z14"] = "Other"
        return o


@register
class LihtcTblTables(Sheet):
    name = "LIHTC-TBL Tables"

    RENT_COLS = ["AJ7", "AJ8", "AJ9", "AJ10", "AJ11"]    # Studio..4BR 100% AMI max rent

    def compute(self, ctx):
        mi = ctx.mi
        o = {"B2": _title(ctx, "LIHTC-TBL TABLES")}
        counties = _table("lihtc_tbl_tables")["counties"]

        # Statewide basis delta: each county's 2BR threshold basis limit vs the median county.
        median_2br = statistics.median(xl.nums([c["tbl"][2] for c in counties]))
        o["E6"] = median_2br
        adjust = {c["row"]: xl.n(c["tbl"][2]) / median_2br - 1 for c in counties}
        lo, hi = min(adjust.values()), max(adjust.values())

        # FMR relativities: average FMR of each bedroom size vs the 1BR average.
        fmr_avg = [xl.AVERAGE([c["fmr"][i] for c in counties]) for i in range(5)]
        for col, avg in zip("STUVW", fmr_avg):
            o[f"{col}9"] = avg / fmr_avg[1]

        ami_by_county = {}
        for c in counties:
            ami_by_county.setdefault(str(c["ami_county"]).lower(), c["ami_100"])

        for c in counties:
            r = c["row"]
            o[f"H{r}"] = adjust[r]
            ami = ami_by_county[str(c["fmr_county"]).lower()]       # VLOOKUP(R, J:K, 2, FALSE)
            o[f"X{r}"] = ami
            # 1BR FMR as a share of the 30%-of-income affordable rent at 100% AMI.
            o[f"Y{r}"] = xl.n(c["fmr"][1]) / (xl.n(ami) / 12 * 0.3)
            o[f"AB{r}"] = c["county"]
            o[f"AC{r}"] = adjust[r]
            o[f"AD{r}"] = o[f"Y{r}"]
            o[f"AE{r}"] = (adjust[r] - lo) / (hi - lo)               # normalised basis boost
            o[f"AF{r}"] = o[f"AD{r}"]
            o[f"AG{r}"] = o[f"AE{r}"] * o[f"AF{r}"]                  # composite score

        studio_tbl = [c["tbl"][0] for c in counties]
        o["C69"] = xl.SUM(studio_tbl)
        o["C71"] = xl.AVERAGE(studio_tbl)
        o["C72"] = xl.MIN(studio_tbl)

        # Project county's 100% AMI max rents, trended pre-stabilization and post-award.
        trend = ((1 + xl.n(mi["pre_stabilization_trending_rate"])) ** xl.n(mi["pre_stabilization_trending_years"])
                 * (1 + xl.n(mi["post_award_trending_rate"])) ** xl.n(mi["post_award_trending_years"]))
        county_col = [c["ami_county"] for c in counties]
        for i, addr in enumerate(self.RENT_COLS):
            rents = [c["max_rent_100"][i] for c in counties]
            o[addr] = xl.SUMIF(county_col, mi["county"], rents) * trend
        return o


@register
class FederalMinimumOpex(Sheet):
    name = "Federal Minimum Opex"

    def compute(self, ctx):
        # Pure reference table (TCAC per-unit operating cost minimums by region / type).
        o = {a: v for a, v in _table("federal_minimum_opex")["cells"].items()}
        o["B2"] = _title(ctx, "FEDERAL MINIMUM OPEX")
        return o
