"""Market-rent chain: comp datasets -> amenity / location adjustments -> market rent by unit type.

Location Premiums           neighborhood scores; a comp's rent is scaled by
                            (1 + rate per point) ^ (subject score - comp score).
Amenity & Property Premiums premium library ($/mo per advertised amenity) applied to each comp
                            building, a rule-based class cross-check, comp coordinates
                            (haversine distance to the site) and the subject's own amenities.
Studio / 1BR / 2BR Comps    one row per comp unit: concession-adjusted effective rent, the
                            building's amenity premium, and the rent adjusted to the subject
                            (strip comp amenities, add subject amenities, apply location).
Market Rent Survey          5-mile listing survey; its average asking rent per unit type is
                            the market rent the Rent Builder uses (via Comps Summary row 13).
Comps Summary               comp statistics filtered by the Master Inputs class toggles.

Typed comp data lives in data/site8/tables/*.json (tools/extract_rent_tables.py).
"""
from __future__ import annotations

import json
import math
from functools import lru_cache

from uw import xl
from uw.engine import DATA, Sheet, register

AMENITY = "Amenity & Property Premiums"
LOCATION = "Location Premiums"


# ---------------------------------------------------------------- helpers
@lru_cache(maxsize=None)
def _table(name: str, deal: str = "site8") -> dict:
    return json.loads((DATA / deal / "tables" / name).read_text(encoding="utf-8"))


def _blank(v) -> bool:
    return v is None or v == ""


def _lookup(key, keys, values, default, blank_cell=0):
    """IFERROR(INDEX(values, MATCH(key, keys, 0)), default).

    A blank lookup value never matches (Excel MATCH returns #N/A), and INDEX of an
    empty *typed* cell returns 0 (pass blank_cell=None for formula columns, where a
    blank in the snapshot stands for a "" result).
    """
    if _blank(key):
        return default
    for k, v in zip(keys, values):
        if not _blank(k) and xl.eq(k, key):
            return blank_cell if v is None else v
    return default


def _label(v):
    """Displayed value of a typed cell; HYPERLINK cells show their label."""
    return v["label"] if isinstance(v, dict) else v


def _gt0(v) -> bool:
    """Excel `v > 0` in an array expression: text (even "") sorts above every number."""
    if isinstance(v, str):
        return True
    return xl.is_number(v) and v > 0


def _num(v) -> float:
    """SUMPRODUCT coercion: text and blanks count as 0."""
    return float(v) if xl.is_number(v) else 0.0


# ---------------------------------------------------------------- Location Premiums
@register
class LocationPremiums(Sheet):
    name = LOCATION

    SUBJECT_SITE = "21101 Victory Blvd, Canoga Park"   # C8
    SUBJECT_NEIGHBORHOOD = "ZIP 91303"                  # C9
    SUBJECT_SCORE = 6                                   # C10 (1-10)
    PCT_PER_POINT = 0.05                                # C11 rent adjustment per score point
    APPLY = "Yes"                                       # C12
    # B19:C32 neighborhood -> score (rows 21-32 are empty slots on the tab).
    NEIGHBORHOODS = [("ZIP 91303", 6), ("ZIP 91307", 6)] + [(None, None)] * 12
    FIRST_ROW = 19

    def compute(self, ctx):
        o = {"C8": self.SUBJECT_SITE, "C9": self.SUBJECT_NEIGHBORHOOD, "C10": self.SUBJECT_SCORE,
             "C11": self.PCT_PER_POINT, "C12": self.APPLY}
        comp_names = ctx.col(AMENITY, "B87:B146")
        comp_hoods = ctx.col(AMENITY, "I87:I146")
        comp_class = ctx.col(AMENITY, "H87:H146")
        tagged = tagged_a = 0
        for i, (hood, score) in enumerate(self.NEIGHBORHOODS):
            r = self.FIRST_ROW + i
            o[f"B{r}"], o[f"C{r}"] = hood, score
            # EXCEL-QUIRK: empty slot rows still get a multiplier, (1.05)^(6-0) = 1.34.
            o[f"D{r}"] = ((1 + self.PCT_PER_POINT) ** (self.SUBJECT_SCORE - xl.n(score))
                          if xl.eq(self.APPLY, "Yes") else 1)
            # COUNTIF with a blank criterion cell matches nothing here.
            in_hood = [not _blank(hood) and xl.eq(h, hood) for h in comp_hoods]
            o[f"E{r}"] = sum(in_hood)
            o[f"F{r}"] = sum(1 for m, c in zip(in_hood, comp_class) if m and xl.eq(c, "A"))
            tagged += o[f"E{r}"]
            tagged_a += o[f"F{r}"]
        o["E33"], o["F33"] = tagged, tagged_a
        n_buildings = sum(1 for b in comp_names if not _blank(b))
        o["G33"] = ("All comp buildings tagged" if tagged == n_buildings
                    else f"CHECK: {n_buildings - tagged} comp building(s) untagged")
        return o


# ---------------------------------------------------------------- Amenity & Property Premiums
@register
class AmenityPremiums(Sheet):
    name = AMENITY

    # Row 9: assumed monthly premium per advertised amenity, columns D:Q
    # (W/D, pool, fitness, rooftop, clubhouse, pet, EV, balcony, package, access,
    #  elevator, A/C, smart home, premium finishes).
    PREMIUMS = [100, 30, 50, 25, 20, 15, 15, 30, 15, 10, 10, 50, 25, 75]
    CLASS_A_MIN_SCORE = 8      # T7: A if amenity count + vintage points >= 8
    CLASS_C_MAX_SCORE = 3      # V7: C if score <= 3
    VINTAGE_POINTS = {"2020": 4, "2010": 3, "2000": 2, "1990": 1}
    SITE_LAT, SITE_LON = 34.1887214, -118.5932357           # D84, E84
    SITE_LABEL, SITE_ADDRESS, SITE_COORD_SOURCE = "Site", "21101 Victory Blvd, Canoga Park", "Google Places"
    EARTH_RADIUS_MI = 3958.8
    # Rows 151/153: subject amenities, columns B:M, then N = 2nd bath in 2BR.
    # EXCEL-QUIRK: the subject list prices W/D at $75 and premium finishes at $25 (comp
    # library: $100 / $75) and omits A/C and smart home, so comp and subject premiums are
    # not measured on the same scale. Kept as typed.
    SUBJECT_PREMIUMS = [75, 30, 50, 25, 20, 15, 15, 30, 15, 10, 10, 25]
    SUBJECT_FLAGS = [0.5, 0, 1, 1, 1, 1, 0, 0, 0, 1, 1, 1]
    SECOND_BATH_PREMIUM = 150  # N151
    SUBJECT_SECOND_BATH = 0    # N153

    def compute(self, ctx):
        o = {}
        t = _table("amenity_premiums.json")
        cols = t["amenity_columns"]
        for c, p in zip(cols, self.PREMIUMS):
            o[f"{c}9"] = p
        o["T7"], o["V7"] = self.CLASS_A_MIN_SCORE, self.CLASS_C_MAX_SCORE

        # Amenity matrix: premium, amenity count and rule-based class per comp building.
        for rec in t["amenity_matrix"]:
            r = rec["row"]
            flags = rec["amenities"]
            o[f"B{r}"] = rec.get("building")
            for c, f in zip(cols, flags):
                o[f"{c}{r}"] = f
            o[f"R{r}"] = xl.SUMPRODUCT([xl.n(f) for f in flags], self.PREMIUMS)
            o[f"S{r}"] = xl.SUM(flags)
            vintage = rec.get("vintage")
            o[f"T{r}"] = vintage
            # EXCEL-QUIRK: LEFT(vintage, 4) is compared with "2020"/"2010"/"2000"/"1990", so
            # only years ending in 0 earn vintage points (2024 or 2009 score 0). LEFT(,3)
            # against "202"/"201"/... was presumably intended.
            score = o[f"S{r}"] + self.VINTAGE_POINTS.get(("" if vintage is None else str(vintage))[:4], 0)
            o[f"U{r}"] = ("A" if score >= self.CLASS_A_MIN_SCORE
                          else "C" if score <= self.CLASS_C_MAX_SCORE else "B")
            o[f"V{r}"] = rec.get("class")
            o[f"W{r}"] = rec.get("class_rationale")
            o[f"X{r}"] = _label(rec.get("amenity_listing"))
        for r in range(11, 39):
            o[f"AU{r}"] = 0  # =AQ-AT on empty helper columns

        # Site and comp locations.
        o.update({"B84": self.SITE_LABEL, "C84": self.SITE_ADDRESS, "D84": self.SITE_LAT,
                  "E84": self.SITE_LON, "F84": self.SITE_COORD_SOURCE,
                  "I84": ctx.ref(LOCATION, "C9"), "J84": ctx.ref(LOCATION, "C10")})
        subject_score = ctx.ref(LOCATION, "C10")
        pct_per_pt = xl.n(ctx.ref(LOCATION, "C11"))
        loc_on = xl.eq(ctx.ref(LOCATION, "C12"), "Yes")
        hoods = ctx.col(LOCATION, "B19:B32")
        scores = ctx.col(LOCATION, "C19:C32")
        names = [o[f"B{r}"] for r in range(11, 71)]
        classes = [o[f"V{r}"] for r in range(11, 71)]
        for rec in t["locations"]:
            r = rec["row"]
            lat, lon = xl.n(rec.get("lat")), xl.n(rec.get("lon"))
            o[f"B{r}"], o[f"C{r}"] = rec.get("building"), rec.get("address")
            o[f"D{r}"], o[f"E{r}"] = rec.get("lat"), rec.get("lon")
            o[f"F{r}"], o[f"I{r}"] = rec.get("coord_source"), rec.get("neighborhood")
            o[f"G{r}"] = self.haversine_mi(self.SITE_LAT, self.SITE_LON, lat, lon)
            o[f"H{r}"] = _lookup(o[f"B{r}"], names, classes, "")
            o[f"J{r}"] = _lookup(o[f"I{r}"], hoods, scores, "")
            o[f"K{r}"] = (1 if o[f"J{r}"] == "" or not loc_on
                          else (1 + pct_per_pt) ** (xl.n(subject_score) - o[f"J{r}"]))

        # Subject property amenity premium (studio/1BR; 2BR adds the 2nd-bath premium).
        letters = "BCDEFGHIJKLM"
        for c, p, f in zip(letters, self.SUBJECT_PREMIUMS, self.SUBJECT_FLAGS):
            o[f"{c}151"], o[f"{c}153"] = p, f
        o["N151"], o["N153"] = self.SECOND_BATH_PREMIUM, self.SUBJECT_SECOND_BATH
        o["O153"] = xl.SUMPRODUCT(self.SUBJECT_PREMIUMS, self.SUBJECT_FLAGS)
        o["P153"] = o["O153"] + self.SECOND_BATH_PREMIUM * self.SUBJECT_SECOND_BATH
        return o

    def haversine_mi(self, lat0, lon0, lat, lon):
        rad = math.radians
        h = (math.sin(rad(lat - lat0) / 2) ** 2
             + math.cos(rad(lat0)) * math.cos(rad(lat)) * math.sin(rad(lon - lon0) / 2) ** 2)
        return self.EARTH_RADIUS_MI * 2 * math.asin(math.sqrt(h))


# ---------------------------------------------------------------- comp tabs
class _AmenityLibrary:
    """The Amenity & Property Premiums ranges a comp tab looks buildings up in."""

    def __init__(self, ctx):
        self.names = ctx.col(AMENITY, "B11:B70")
        self.premium = ctx.col(AMENITY, "R11:R70")
        self.cls = ctx.col(AMENITY, "V11:V70")
        self.loc_names = ctx.col(AMENITY, "B87:B146")
        self.distance = ctx.col(AMENITY, "G87:G146")
        self.hood = ctx.col(AMENITY, "I87:I146")
        self.score = ctx.col(AMENITY, "J87:J146")
        self.multiplier = ctx.col(AMENITY, "K87:K146")
        self.subject_premium = ctx.ref(AMENITY, "O153")
        self.subject_premium_2br = ctx.ref(AMENITY, "P153")
        self.second_bath_premium = ctx.ref(AMENITY, "N151")


def effective_rent(rec, lib):
    """Columns D, M:T, X of a comp row: concessions, amenity premium, $/SF, class."""
    b, rent, sf = rec.get("building"), rec.get("rent"), rec.get("sf")
    c = {"D": _lookup(b, lib.loc_names, lib.distance, "", blank_cell=None)}
    if _blank(rent):
        c.update(M="", N="", O="", Q="")
    else:
        # Concession spread over the lease: (rent x free months + one-time credit) / term.
        c["M"] = (rent * xl.n(rec.get("free_months")) + xl.n(rec.get("one_time_credit"))) \
            / xl.n(rec.get("lease_term"))
        c["N"] = rent - c["M"]
        c["O"] = c["M"] / rent
    c["P"] = _lookup(b, lib.names, lib.premium, 0, blank_cell=None)   # comp amenity premium
    if not _blank(rent):
        c["Q"] = c["N"] - xl.n(c["P"])                                  # base unit rent
    per_sf = not (_blank(rent) or _blank(sf))
    c["R"] = rent / sf if per_sf else ""
    c["S"] = c["N"] / sf if per_sf else ""
    c["T"] = c["Q"] / sf if per_sf else ""
    c["X"] = _lookup(b, lib.names, lib.cls, "")
    return c


def adjusted_rent(rec, c, lib, cols, second_bath=False):
    """Rent adjusted to the subject: effective rent - comp premium + subject premium, x location.

    `cols` names the output columns (they shift right by three on the 2BR tab, which
    adds a 2nd-bath premium to the comp side and uses the 2BR subject premium).
    """
    b, sf, eff = rec.get("building"), rec.get("sf"), c["N"]
    a = {}
    if second_bath:
        a["bath"] = "" if eff == "" else (
            lib.second_bath_premium if xl.eq(rec.get("second_bath"), "Y") else 0)
        a["comp_premium"] = "" if eff == "" else xl.n(c["P"]) + xl.n(a["bath"])
        subject = lib.subject_premium_2br
    else:
        a["comp_premium"] = c["P"]
        subject = lib.subject_premium
    a["subject"] = "" if eff == "" else subject
    a["pre_location"] = "" if eff == "" else eff - xl.n(a["comp_premium"]) + xl.n(subject)
    a["hood"] = _lookup(b, lib.loc_names, lib.hood, "")
    a["score"] = _lookup(b, lib.loc_names, lib.score, "", blank_cell=None)
    a["multiplier"] = "" if eff == "" else _lookup(b, lib.loc_names, lib.multiplier, 1, blank_cell=None)
    a["adjusted"] = "" if eff == "" else a["pre_location"] * xl.n(a["multiplier"])
    a["adjusted_psf"] = "" if (a["adjusted"] == "" or _blank(sf)) else a["adjusted"] / sf
    return {cols[k]: v for k, v in a.items() if k in cols}


class _CompTab(Sheet):
    TABLE = ""
    CORE_LAST = 0       # last row of the D, M:T formula columns
    CLASS_LAST = 0      # last row of the class (X) formula column
    ADJ_LAST = 0        # last row of the adjusted-rent formula columns
    SECOND_BATH = False
    ADJ_COLS = {"subject": "Y", "adjusted": "Z", "adjusted_psf": "AA", "hood": "AB",
                "score": "AC", "multiplier": "AD", "pre_location": "AE"}
    TYPED = {"B": "building", "C": "address", "E": "plan", "F": "include", "G": "sf", "H": "rent",
             "I": "special", "J": "free_months", "K": "one_time_credit", "L": "lease_term",
             "U": "photo", "V": "listing", "W": "phone"}
    TITLE = ""

    def compute(self, ctx):
        o = {"B2": f"{ctx.mi['project_name']}  |  {self.TITLE}"}
        t = _table(self.TABLE)
        o.update(t["labels"])
        lib = _AmenityLibrary(ctx)
        recs = {rec["row"]: rec for rec in t["rows"]}
        for r in range(9, self.ADJ_LAST + 1):
            rec = recs.get(r, {"row": r})
            for col, field in self.TYPED.items():
                if field in rec:
                    o[f"{col}{r}"] = _label(rec[field])
            c = effective_rent(rec, lib)
            if r <= self.CORE_LAST:
                for col in ("D", "M", "N", "O", "P", "Q", "R", "S", "T"):
                    o[f"{col}{r}"] = c[col]
            if r <= self.CLASS_LAST:
                o[f"X{r}"] = c["X"]
            o.update({f"{col}{r}": v for col, v in
                      adjusted_rent(rec, c, lib, self.ADJ_COLS, self.SECOND_BATH).items()})
        return o


@register
class StudioComps(_CompTab):
    name = "Studio Comps"
    TITLE = "STUDIO COMPS"
    TABLE = "studio_comps.json"
    CORE_LAST, CLASS_LAST, ADJ_LAST = 47, 47, 292


@register
class OneBRComps(_CompTab):
    name = "1BR Comps"
    TITLE = "1BR COMPS"
    TABLE = "1br_comps.json"
    CORE_LAST, CLASS_LAST, ADJ_LAST = 51, 294, 294

    def compute(self, ctx):
        o = super().compute(ctx)
        o["G52"] = xl.AVERAGE([o.get(f"G{r}") for r in range(9, 52)])
        return o


@register
class TwoBRComps(_CompTab):
    name = "2BR Comps"
    TITLE = "2BR COMPS"
    TABLE = "2br_comps.json"
    CORE_LAST, CLASS_LAST, ADJ_LAST = 42, 42, 294
    SECOND_BATH = True
    TYPED = {**_CompTab.TYPED, "Y": "second_bath"}
    ADJ_COLS = {"bath": "Z", "comp_premium": "AA", "subject": "AB", "adjusted": "AC",
                "adjusted_psf": "AD", "hood": "AE", "score": "AF", "multiplier": "AG",
                "pre_location": "AH"}


# ---------------------------------------------------------------- Market Rent Survey
@register
class MarketRentSurvey(Sheet):
    name = "Market Rent Survey"

    RADIUS_MI = 5  # E6
    TYPED = {"B": "property", "C": "address", "D": "distance", "E": "unit_type", "F": "sf",
             "G": "rent", "H": "listings", "I": "note", "K": "source", "L": "observed"}
    # summary column, unit type, market-discount cell (row 5)
    SUMMARY = [("C", "Studio", "E"), ("D", "1 BR", "F"), ("E", "2 BR", "G")]

    def compute(self, ctx):
        mi = ctx.mi
        t = _table("market_rent_survey.json")
        o = dict(t["labels"])
        o["E6"] = self.RADIUS_MI
        o["E5"] = mi.by_row(386)   # market discount: Studio / 1 BR / 2 BR / 3 BR
        o["F5"] = mi.by_row(387)
        o["G5"] = mi.by_row(388)
        o["H5"] = mi.by_row(389)
        rows = t["rows"]
        for rec in rows:
            r = rec["row"]
            for col, field in self.TYPED.items():
                if field in rec:
                    o[f"{col}{r}"] = rec[field]
            if "sf" in rec:
                o[f"J{r}"] = rec["rent"] / rec["sf"]
        types = [rec.get("unit_type") for rec in rows]
        for col, unit_type, disc in self.SUMMARY:
            mine = [rec for rec, ut in zip(rows, types) if xl.eq(ut, unit_type)]
            o[f"{col}75"] = len(mine)                                     # properties
            o[f"{col}76"] = xl.SUM([rec.get("listings") for rec in mine])  # listings
            o[f"{col}77"] = xl.AVERAGE([rec.get("rent") for rec in mine])  # avg rent per unit
            o[f"{col}78"] = xl.AVERAGE([rec.get("sf") for rec in mine])    # avg SF
            o[f"{col}79"] = o[f"{col}77"] / o[f"{col}78"]
            o[f"{col}80"] = o[f"{col}77"] * (1 - xl.n(o[f"{disc}5"]))
            n = o[f"{col}75"]
            o[f"{col}81"] = "OK" if n >= 5 else "Warning" if n >= 3 else "Fallback"
        return o


# ---------------------------------------------------------------- Comps Summary
class _CompColumns:
    """The comp-tab columns Comps Summary filters and averages, for one unit type."""

    def __init__(self, ctx, tab, last, score_col):
        def col(c):
            return ctx.col(tab, f"{c}9:{c}{last}")
        self.tab = tab
        self.include, self.cls, self.rent = col("F"), col("X"), col("H")
        self.cols = {c: col(c) for c in ("D", "G", "H", "M", "N", "P", "Q", "R", "S", "T")}
        self.score = col(score_col)
        if tab == "2BR Comps":
            self.cols["Y"], self.cols["Z"] = col("Y"), col("Z")

    def __len__(self):
        return len(self.rent)


@register
class CompsSummary(Sheet):
    name = "Comps Summary"

    TYPES = [("C", "Studio Comps", 292, "AC"), ("D", "1BR Comps", 294, "AC"),
             ("E", "2BR Comps", 294, "AF")]
    # Rows 21-28 (filtered averages): row -> (comp column, extra "field > 0" condition)
    AVERAGES = {21: ("D", True), 22: ("G", True), 23: ("H", False), 24: ("M", False),
                25: ("N", False), 26: ("P", False), 28: ("Q", False)}
    # Class blocks (rows 54+): metric offset -> (comp column, condition column must be > 0)
    CLASS_METRICS = [("D", "D"), ("G", "G"), ("H", "H"), ("M", "H"), ("N", "H"), ("P", "H"),
                     ("Q", "H"), ("R", "R"), ("S", "S"), ("T", "T")]
    CLASS_COLS = [("C", "A"), ("D", "B"), ("E", "C"), ("F", "*")]

    def compute(self, ctx):
        mi = ctx.mi
        o = {"C8": mi["inc_class_a"], "D8": mi["inc_class_b"], "E8": mi["inc_class_c"],
             "C9": mi["comps_to_include"]}
        o.update({"C7": "Class A", "D7": "Class B", "E7": "Class C"})
        o.update({"C12": "Studio", "D12": "1 BR", "E12": "2 BR", "F12": "3 BR"})
        o.update({"C19": "Studio", "D19": "1 BR", "E19": "2 BR"})
        o.update({"C35": "Studio", "D35": "1 BR", "E35": "2 BR"})
        # Row 13: market rent per unit type = Market Rent Survey average (3 BR uses 2 BR).
        for col, src in (("C", "C"), ("D", "D"), ("E", "E"), ("F", "E")):
            o[f"{col}13"] = ctx.ref("Market Rent Survey", f"{src}77")
            o[f"{col}14"] = ctx.ref("Market Rent Survey", f"{src}79")

        all_comps = xl.eq(o["C9"], "All comps")
        class_on = {"A": xl.eq(o["C8"], "Yes"), "B": xl.eq(o["D8"], "Yes"), "C": xl.eq(o["E8"], "Yes")}
        subject_premium = ctx.ref(AMENITY, "O153")
        subject_score = ctx.ref(LOCATION, "C10")
        tabs = {}
        for col, tab, last, score_col in self.TYPES:
            comps = tabs[col] = _CompColumns(ctx, tab, last, score_col)
            # Comp passes the filter: included (or all comps), class switched on, has a rent.
            keep = [(all_comps or xl.eq(f, "Y"))
                    and any(xl.eq(x, k) and on for k, on in class_on.items())
                    and _gt0(h)
                    for f, x, h in zip(comps.include, comps.cls, comps.rent)]

            def avg(values, extra=None):
                w = [k and (extra is None or e) for k, e in zip(keep, extra or keep)]
                den = sum(w)
                return math.fsum(_num(v) for v, x in zip(values, w) if x) / den if den else "-"

            o[f"{col}15"] = avg(comps.score, [not _blank(s) for s in comps.score])
            o[f"{col}16"] = subject_score
            o[f"{col}20"] = sum(keep)
            for r, (c, positive_only) in self.AVERAGES.items():
                vals = comps.cols[c]
                o[f"{col}{r}"] = avg(vals, [_gt0(v) for v in vals] if positive_only else None)
            for r, num in ((29, 23), (30, 25), (31, 28)):
                o[f"{col}{r}"] = xl.n(o[f"{col}{num}"]) / xl.n(o[f"{col}22"])
            o[f"{col}36"] = subject_premium
            o[f"{col}39"] = xl.n(o[f"{col}26"]) - xl.n(subject_premium)
            if col == "E":
                o["E27"] = avg(comps.cols["Z"], [not _blank(y) for y in comps.cols["Y"]])
                o["E37"] = xl.n(ctx.ref(AMENITY, "N153")) * xl.n(ctx.ref(AMENITY, "N151"))
                o["E40"] = xl.n(o["E27"]) - o["E37"]
                o["E41"] = o["E39"] + o["E40"]
            else:
                o[f"{col}27"] = o[f"{col}40"] = "-"
                o[f"{col}41"] = o[f"{col}39"]
            # Location adjustment embedded in row 13 vs the comps' amenity-adjusted rent.
            base = xl.IFERROR(lambda: o[f"{col}25"] - o[f"{col}41"], None)
            o[f"{col}17"] = 0 if base is None else o[f"{col}13"] - base
            o[f"{col}18"] = xl.IFERROR(lambda: o[f"{col}17"] / base, 0) if base is not None else 0

        # Display-only blocks by building class (Include = Y, rent > 0).
        blocks = [(54, ["C"]), (68, ["D"]), (82, ["E"]), (96, ["C", "D", "E"])]
        for first, members in blocks:
            for out_col, cls in self.CLASS_COLS:
                def match(comps, k, cond_col):
                    x = comps.cls[k]
                    cls_ok = isinstance(x, str) if cls == "*" else xl.eq(x, cls)
                    return (xl.eq(comps.include[k], "Y") and cls_ok and _pos(comps.rent[k])
                            and (cond_col is None or _pos(comps.cols[cond_col][k])))

                o[f"{out_col}{first}"] = sum(match(tabs[m], k, None)
                                             for m in members for k in range(len(tabs[m])))
                for i, (c, cond) in enumerate(self.CLASS_METRICS, start=1):
                    tot = cnt = 0.0
                    for m in members:
                        comps = tabs[m]
                        for k in range(len(comps)):
                            if match(comps, k, cond):
                                cnt += 1
                                v = comps.cols[c][k]
                                tot += v if xl.is_number(v) else 0
                    o[f"{out_col}{first + i}"] = tot / cnt if cnt else "-"
                if first == 82:  # 2BR: average 2nd-bath premium over comps with the flag filled
                    comps = tabs["E"]
                    sel = [k for k in range(len(comps)) if match(comps, k, None)
                           and not _blank(comps.cols["Y"][k])]
                    z = [comps.cols["Z"][k] for k in sel]
                    o[f"{out_col}93"] = (math.fsum(v for v in z if xl.is_number(v)) / len(sel)
                                         if sel else "-")
        return o


def _pos(v) -> bool:
    """SUMIFS/COUNTIFS criterion ">0": numbers only."""
    return xl.is_number(v) and v > 0
