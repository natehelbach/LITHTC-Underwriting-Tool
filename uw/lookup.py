"""Site lookup: fill location inputs from a street address using free public data.

    python -m uw.lookup "21101 Victory Blvd, Canoga Park, CA 91303"

Sources (no API keys):
  * US Census geocoder        address match, city, county, tract, ZIP, congressional
                              and state legislative districts, urban area
  * HUD (ArcGIS, official)    Qualified Census Tracts and Difficult Development Areas
  * TCAC/HCD opportunity map  resource category of the tract
  * LA County Assessor        APN, lot size, existing building, assessed values (LA County)
  * LA City ZIMAS             zoning (City of Los Angeles)
  * OpenStreetMap (Overpass)  amenities within 1/2 mile (best effort; public server)

Each finding records where it came from. Nothing here is authoritative; the
inputs page shows every filled value with its source for review.
"""
from __future__ import annotations

import json
import math
import re
import sys
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field

UA = "underwriting-tool/1.0"
TIMEOUT = 20

CENSUS = "https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress"
HUD = "https://services.arcgis.com/VTyQ9soqVukalItT/arcgis/rest/services"
HUD_QCT = [f"{HUD}/QUALIFIED_CENSUS_TRACTS_2026/FeatureServer/0",   # newest designations first
           f"{HUD}/QUALIFIED_CENSUS_TRACTS/FeatureServer/0"]
HUD_DDA = [f"{HUD}/Difficult_Development_Areas_2026/FeatureServer/0",
           f"{HUD}/Difficult_Development_Areas/FeatureServer/0"]
OPP_MAPS = [  # (label, layer); statewide TCAC/HCD opportunity map data
    ("2026 TCAC/HCD Opportunity Map",
     "https://services7.arcgis.com/3gcw4ks3WKzHbOWv/arcgis/rest/services/TCAC_HCD_Opportunity_Area_Map_2026/FeatureServer/0"),
    ("2025 TCAC/HCD Opportunity Map",
     "https://services8.arcgis.com/a4GMqC2tQHvYiVtK/arcgis/rest/services/main_final_opp_2025_public/FeatureServer/0"),
]
LA_PARCELS = "https://public.gis.lacounty.gov/public/rest/services/LACounty_Cache/LACounty_Parcel/MapServer/0"
ZIMAS = "https://zimas.lacity.org/arcgis/rest/services/zma/zoning/MapServer"
OVERPASS = ["https://overpass-api.de/api/interpreter",
            "https://overpass.kumi.systems/api/interpreter"]

# County -> TCAC geographic region (CTCAC regulations). Counties not listed are
# left for the user to pick. Los Angeles splits on City vs balance of county.
TCAC_REGION = {
    **dict.fromkeys(["El Dorado", "Placer", "Sacramento", "Sutter", "Yolo", "Yuba"], "Capital Region"),
    **dict.fromkeys(["Monterey", "San Benito", "San Luis Obispo", "Santa Barbara", "Santa Cruz",
                     "Ventura"], "Central Coast Region"),
    **dict.fromkeys(["Fresno", "Kern", "Kings", "Madera", "Merced", "San Joaquin", "Stanislaus",
                     "Tulare"], "Central Valley Region"),
    **dict.fromkeys(["Alameda", "Contra Costa"], "East Bay Region"),
    **dict.fromkeys(["Imperial", "Riverside", "San Bernardino"], "Inland Empire Region"),
    **dict.fromkeys(["Marin", "San Mateo", "Santa Clara"], "South and West Bay Region"),
    **dict.fromkeys(["Butte", "Del Norte", "Humboldt", "Lake", "Mendocino", "Napa", "Nevada",
                     "Shasta", "Solano", "Sonoma", "Tehama"], "Northern Region"),
    "Orange": "Orange County", "San Diego": "San Diego County",
    "San Francisco": "San Francisco County",
}
CDLAC_REGION = {
    **dict.fromkeys(["Alameda", "Contra Costa", "Marin", "Napa", "San Francisco", "San Mateo",
                     "Santa Clara", "Solano", "Sonoma"], "Bay Area"),
    **dict.fromkeys(["Monterey", "Orange", "San Benito", "San Diego", "San Luis Obispo",
                     "Santa Barbara", "Santa Cruz", "Ventura"], "Coastal"),
    **dict.fromkeys(["Fresno", "Imperial", "Kern", "Kings", "Madera", "Merced", "Riverside",
                     "Sacramento", "San Bernardino", "San Joaquin", "Stanislaus", "Tulare"],
                    "Inland"),
}
OPEX_REGION_NAME = {"South and West Bay Region": "South & West Bay Region"}

AMENITIES = {  # input key -> (label, Overpass selectors)
    "within_1_2_mile_of_grocery": ("grocery store", ['["shop"="supermarket"]']),
    "within_1_2_mile_of_library": ("library", ['["amenity"="library"]']),
    "within_1_2_mile_of_school": ("school", ['["amenity"="school"]']),
    "within_1_2_mile_of_hospital_clinic": ("hospital / clinic",
                                           ['["amenity"="hospital"]', '["amenity"="clinic"]',
                                            '["amenity"="doctors"]']),
    "within_1_2_mile_of_pharmacy": ("pharmacy", ['["amenity"="pharmacy"]', '["shop"="chemist"]']),
    "within_1_2_mile_of_public_park": ("public park", ['["leisure"="park"]']),
}
HALF_MILE_M = 805

# Inputs that describe the site itself. After a lookup, any of these that were not
# filled still hold the previous site's value and are flagged for review.
SITE_KEYS = ["project_name", "city", "county", "tcac_geographic_area", "cdlac_region",
             "opex_region", "census_tract", "apn_1", "apn_2", "zoning_district",
             "federal_congressional_district", "state_assembly_district", "state_senate_district",
             "land_cost_value", "improvements_value", "lot_size", "demo_required",
             "qct_dda_status", "neighborhood_change_area", "opportunity_resource_area_type",
             "rural_set_aside", *AMENITIES]


@dataclass
class Found:
    value: object
    source: str
    note: str = ""


@dataclass
class Lookup:
    address: str
    matched: str | None = None
    lon: float | None = None
    lat: float | None = None
    found: dict[str, Found] = field(default_factory=dict)     # input key -> finding
    facts: list[tuple[str, str, str]] = field(default_factory=list)  # (label, value, source) shown only
    missing: list[str] = field(default_factory=list)          # what could not be determined
    errors: list[str] = field(default_factory=list)           # sources that failed

    def put(self, key, value, source, note=""):
        if value not in (None, ""):
            self.found[key] = Found(value, source, note)

    def fact(self, label, value, source):
        if value not in (None, "", 0):
            self.facts.append((label, str(value), source))


# ---------------------------------------------------------------- http helpers
def get_json(url: str, params: dict | None = None, data: dict | None = None,
             timeout: float = TIMEOUT) -> dict:
    if params:
        url = f"{url}?{urllib.parse.urlencode(params)}"
    body = urllib.parse.urlencode(data).encode() if data else None
    req = urllib.request.Request(url, data=body, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def arcgis_query(layer: str, **params) -> list[dict]:
    p = {"outFields": "*", "returnGeometry": "false", "f": "json", **params}
    d = get_json(f"{layer}/query", p)
    if "error" in d:
        raise RuntimeError(d["error"].get("message", "ArcGIS error"))
    return d.get("features", [])


def point_params(lon, lat, **extra):
    return {"geometry": f"{lon},{lat}", "geometryType": "esriGeometryPoint", "inSR": "4326",
            "spatialRel": "esriSpatialRelIntersects", **extra}


def layer_year(layer: str) -> str:
    """Designation year of a HUD layer: HUD publishes in the fall for the next year."""
    try:
        d = get_json(layer, {"f": "json"})
        ms = (d.get("editingInfo") or {}).get("dataLastEditDate")
        if ms:
            import datetime as dt
            t = dt.datetime.fromtimestamp(ms / 1000, dt.timezone.utc)
            return str(t.year + 1 if t.month >= 7 else t.year)
    except Exception:  # noqa: BLE001
        pass
    return ""


# ---------------------------------------------------------------- sources
def census(lk: Lookup) -> dict:
    d = get_json(CENSUS, {"address": lk.address, "benchmark": "Public_AR_Current",
                          "vintage": "Current_Current", "layers": "all", "format": "json"})
    matches = d.get("result", {}).get("addressMatches", [])
    if not matches:
        raise LookupError("The Census geocoder could not match this address. "
                          "Try the format: number street, city, state ZIP.")
    m = matches[0]
    lk.matched = m["matchedAddress"]
    lk.lon, lk.lat = m["coordinates"]["x"], m["coordinates"]["y"]
    g = {k: (v[0] if v else {}) for k, v in m["geographies"].items()}

    def layer(prefix):
        for k, v in g.items():
            if prefix.lower() in k.lower():
                return v
        return {}

    src = "US Census geocoder"
    place = layer("Incorporated Places")
    county = layer("Counties").get("BASENAME")
    tract = layer("Census Tracts")
    cd = layer("Congressional Districts")
    cd_name = next((k for k in g if "Congressional Districts" in k), "")
    upper, lower = layer("Legislative Districts - Upper"), layer("Legislative Districts - Lower")
    zcta = layer("ZIP Code Tabulation Areas").get("ZCTA5") or m["addressComponents"].get("zip")
    urban = layer("Urban Areas").get("NAME")

    city = (place.get("BASENAME") or "").strip()
    lk.put("project_name", _title_address(lk.matched), src, "matched address")
    lk.put("city", city, src, f"incorporated place: {place.get('NAME', 'none')}"
           if city else "")
    if not city:
        lk.missing.append("City: the address is outside an incorporated city (unincorporated "
                          f"{county or ''} County)")
    lk.put("county", county, src)
    lk.put("census_tract", tract.get("GEOID"), src, f"Census Tract {tract.get('BASENAME', '')}")
    lk.put("federal_congressional_district", _num(cd.get("BASENAME")), src,
           cd_name.replace(" Congressional Districts", " Congress"))
    lk.put("state_senate_district", _num(upper.get("BASENAME")), src,
           next((k for k in g if "Upper" in k), ""))
    lk.put("state_assembly_district", _num(lower.get("BASENAME")), src,
           next((k for k in g if "Lower" in k), ""))
    lk.fact("ZIP", zcta, src)
    lk.fact("Urban area", urban or "none (outside Census urban areas)", src)
    lk.fact("Coordinates", f"{lk.lat:.6f}, {lk.lon:.6f}", src)
    return {"city": city, "county": county, "tract": tract.get("GEOID"), "zcta": zcta,
            "urban": bool(urban), "state": layer("States").get("BASENAME")}


def hud_qct(lk: Lookup, geoid: str):
    out = []
    for layer in HUD_QCT:
        year = layer_year(layer)
        hit = arcgis_query(layer, where=f"GEOID='{geoid}'", outFields="GEOID")
        out.append((year, bool(hit)))
    return out


def hud_dda(lk: Lookup, lon, lat):
    out = []
    for layer in HUD_DDA:
        year = layer_year(layer)
        hit = arcgis_query(layer, **point_params(lon, lat), outFields="DDA_NAME,DDA_TYPE,ZCTA5")
        out.append((year, hit[0]["attributes"] if hit else None))
    return out


def opportunity(lk: Lookup, lon, lat):
    for label, layer in OPP_MAPS:
        try:
            hit = arcgis_query(layer, **point_params(lon, lat), outFields="*")
        except Exception as e:  # noqa: BLE001
            lk.errors.append(f"{label}: {e}")
            continue
        if hit:
            return label, hit[0]["attributes"]
    return None, None


def la_parcel(lk: Lookup) -> dict | None:
    """LA County Assessor parcel for the address (situs match, else the parcel at the point)."""
    m = re.match(r"\s*(\d+)\s+(?:[NSEW]\.?\s+)?([A-Za-z0-9]+)", lk.address)
    feats = []
    if m:
        where = f"SitusHouseNo='{m.group(1)}' AND SitusStreet LIKE '{m.group(2).upper()}%'"
        feats = arcgis_query(LA_PARCELS, where=where, returnGeometry="true", outSR="4326")
        if len(feats) > 1:  # same number/street in several cities: keep the one nearest the match
            feats.sort(key=lambda f: _dist(_centroid(f), (lk.lon, lk.lat)))
    if not feats:
        feats = arcgis_query(LA_PARCELS, **point_params(lk.lon, lk.lat, distance=40,
                                                         units="esriSRUnit_Meter"),
                             returnGeometry="true", outSR="4326")
        feats.sort(key=lambda f: _dist(_centroid(f), (lk.lon, lk.lat)))
    if not feats:
        return None
    f = feats[0]
    a = f["attributes"]
    a["_centroid"] = _centroid(f)
    return a


def zimas_zoning(lon, lat) -> str | None:
    d = get_json(f"{ZIMAS}/identify", {
        "geometry": f"{lon},{lat}", "geometryType": "esriGeometryPoint", "sr": "4326",
        "layers": "all", "tolerance": "0", "mapExtent": f"{lon - .001},{lat - .001},{lon + .001},{lat + .001}",
        "imageDisplay": "400,400,96", "returnGeometry": "false", "f": "json"})
    zones = []
    for r in d.get("results", []):
        z = r.get("attributes", {}).get("ZONE_CMPLT") or r.get("attributes", {}).get("ZONE_CLASS")
        if z and z not in zones:
            zones.append(z)
    return " / ".join(zones) or None


def amenities(lon, lat) -> dict[str, list[str]]:
    parts = []
    for key, (_, sels) in AMENITIES.items():
        for s in sels:
            parts.append(f"nwr(around:{HALF_MILE_M},{lat},{lon}){s};")
    q = f"[out:json][timeout:20];({''.join(parts)});out tags center;"
    last = None
    for url in OVERPASS:
        try:
            d = get_json(url, data={"data": q}, timeout=12)
            break
        except Exception as e:  # noqa: BLE001
            last = e
    else:
        raise RuntimeError(f"OpenStreetMap server busy ({last})")
    found: dict[str, list[str]] = {k: [] for k in AMENITIES}
    for el in d.get("elements", []):
        t = el.get("tags", {})
        name = t.get("name") or "(unnamed)"
        for key, (_, sels) in AMENITIES.items():
            for s in sels:
                k, v = re.findall(r'"([^"]+)"', s)
                if t.get(k) == v and name not in found[key]:
                    found[key].append(name)
    return found


# ---------------------------------------------------------------- orchestration
def lookup(address: str) -> Lookup:
    lk = Lookup(address.strip())
    if not lk.address:
        lk.errors.append("Enter an address.")
        return lk
    try:
        geo = census(lk)
    except Exception as e:  # noqa: BLE001
        lk.errors.append(f"Address lookup: {e}")
        return lk

    county, city = geo["county"], geo["city"]
    in_ca = geo.get("state") == "California"
    in_la_county = county == "Los Angeles"

    # Parcel first (its interior point is better than the street-centerline match).
    parcel = None
    if in_la_county:
        try:
            parcel = la_parcel(lk)
        except Exception as e:  # noqa: BLE001
            lk.errors.append(f"LA County parcel lookup: {e}")
    lon, lat = (parcel["_centroid"] if parcel else (lk.lon, lk.lat))

    jobs = {
        "qct": lambda: hud_qct(lk, geo["tract"]) if geo["tract"] else [],
        "dda": lambda: hud_dda(lk, lon, lat),
        "opp": lambda: opportunity(lk, lon, lat) if in_ca else (None, None),
        "zoning": lambda: zimas_zoning(lon, lat) if city == "Los Angeles" and in_la_county else None,
        "amen": lambda: amenities(lon, lat),
    }
    results, failures = {}, {}
    with ThreadPoolExecutor(max_workers=len(jobs)) as ex:
        futs = {k: ex.submit(f) for k, f in jobs.items()}
        for k, fut in futs.items():
            try:
                results[k] = fut.result(timeout=TIMEOUT + 15)
            except Exception as e:  # noqa: BLE001
                failures[k] = e

    # --- regions (derived from county) ---
    src_regions = "Derived from county"
    if in_la_county:
        region = "City of Los Angeles" if city == "Los Angeles" else "Balance of Los Angeles County"
        lk.put("tcac_geographic_area", region, src_regions)
        lk.put("cdlac_region", region, src_regions)
        lk.put("opex_region", "Los Angeles County", src_regions)
    elif county in TCAC_REGION:
        r = TCAC_REGION[county]
        lk.put("tcac_geographic_area", r, src_regions, "check against current CTCAC regulations")
        lk.put("opex_region", OPEX_REGION_NAME.get(r, r), src_regions, "check")
    else:
        lk.missing.append(f"TCAC geographic area and opex region for {county} County: pick manually")
    if not in_la_county:
        if county in CDLAC_REGION:
            lk.put("cdlac_region", CDLAC_REGION[county], src_regions,
                   "check against current CDLAC regulations")
        elif county:
            lk.put("cdlac_region", "Northern ", src_regions, "check against current CDLAC regulations")

    # --- parcel ---
    if parcel:
        src = "LA County Assessor parcel data"
        lk.put("apn_1", parcel.get("APN"), src, parcel.get("SitusFullAddress", ""))
        area = parcel.get("Shape.STArea()") or parcel.get("Shape_Area")
        if area:
            lk.put("lot_size", int(round(area)), src, "GIS parcel area; confirm with title / survey")
        bldg = sum(parcel.get(f"SQFTmain{i}") or 0 for i in range(1, 6))
        if bldg:
            lk.put("demo_required", int(bldg), src,
                   "existing building SF, assuming it is all demolished")
        lk.fact("Current use", parcel.get("UseDescription"), src)
        lk.fact("Year built", parcel.get("YearBuilt1"), src)
        lk.fact("Existing building SF", f"{bldg:,}" if bldg else "", src)
        for label, k in (("Assessed land value", "Roll_LandValue"),
                         ("Assessed improvement value", "Roll_ImpValue")):
            if parcel.get(k):
                lk.fact(label, f"${parcel[k]:,.0f} (roll {parcel.get('Roll_Year', '')})".replace(" (roll )", ""), src)
        if parcel.get("AIN"):
            lk.fact("Assessor portal", f"https://portal.assessor.lacounty.gov/parceldetail/{parcel['AIN']}", src)
    elif in_la_county:
        lk.missing.append("APN and lot size: no LA County parcel matched this address")
    else:
        lk.missing.append(f"APN, lot size, existing building: parcel data is only wired up for "
                          f"LA County so far ({county} County)")

    # --- zoning ---
    if results.get("zoning"):
        lk.put("zoning_district", results["zoning"], "LA City ZIMAS",
               "base zone; check ZIMAS for specific plans and overlays")
        if parcel and parcel.get("AIN"):
            lk.fact("ZIMAS", "https://zimas.lacity.org/", "LA City Planning")
    elif "zoning" in failures:
        lk.errors.append(f"ZIMAS zoning: {failures['zoning']}")
    else:
        lk.missing.append("Zoning: only wired up for the City of Los Angeles so far")

    # --- QCT / DDA ---
    qct, dda = results.get("qct"), results.get("dda")
    if "qct" in failures:
        lk.errors.append(f"HUD QCT: {failures['qct']}")
    if "dda" in failures:
        lk.errors.append(f"HUD DDA: {failures['dda']}")
    if qct is not None and dda is not None:
        (qy, is_qct), (dy, dda_hit) = qct[0], dda[0]
        status = "QCT" if is_qct else ("DDA" if dda_hit else "N/A")
        detail = []
        for (y, q) in qct:
            detail.append(f"{y or '?'} QCT: {'yes' if q else 'no'}")
        for (y, h) in dda:
            detail.append(f"{y or '?'} DDA: {'yes (' + h['DDA_NAME'] + ')' if h else 'no'}")
        both = is_qct and dda_hit
        lk.put("qct_dda_status", status, "HUD QCT/DDA designations",
               "; ".join(detail) + ("; site is in both a QCT and a DDA" if both else ""))
        lk.fact("HUD QCT/DDA detail", "; ".join(detail), "HUD")

    # --- opportunity map ---
    if "opp" in failures:
        lk.errors.append(f"Opportunity map: {failures['opp']}")
    else:
        label, opp = results.get("opp") or (None, None)
        if opp:
            cat = opp.get("oppcat")
            if opp.get("pov_seg_flag") == 1:
                cat = "High Segregation & Poverty"
            lk.put("opportunity_resource_area_type", cat, label,
                   f"tract {opp.get('fips', '')}, region {opp.get('region', '')}".strip())
        elif in_ca:
            lk.missing.append("Opportunity / resource area: tract not found on the opportunity map")

    # --- amenities ---
    if "amen" in failures:
        lk.errors.append(f"Amenities within 1/2 mile (OpenStreetMap): {failures['amen']}. "
                         "Try the lookup again in a minute.")
    elif results.get("amen") is not None:
        for key, names in results["amen"].items():
            label = AMENITIES[key][0]
            lk.put(key, 1 if names else 0, "OpenStreetMap",
                   (f"{len(names)} within 1/2 mile: " + ", ".join(names[:3])
                    + (" …" if len(names) > 3 else "")) if names
                   else f"no {label} mapped within 1/2 mile; TCAC rules (size, type, walking "
                        f"distance) still need checking")

    lk.missing.append("Neighborhood change area: no public data service found; check the HCD map")
    return lk


# ---------------------------------------------------------------- helpers
def _num(v):
    if v in (None, ""):
        return None
    try:
        return str(int(v))
    except ValueError:
        return v


def _title_address(a: str | None) -> str | None:
    if not a:
        return a
    parts = [p.strip() for p in a.split(",")]
    if len(parts) >= 4:
        street, city, state, zip_ = parts[0], parts[1], parts[2], parts[3]
        return f"{street.title()}, {city.title()}, {state} {zip_}"
    return a.title()


def _centroid(feature) -> tuple[float, float]:
    rings = (feature.get("geometry") or {}).get("rings") or []
    pts = rings[0][:-1] if rings and len(rings[0]) > 1 else []
    if not pts:
        return (math.nan, math.nan)
    return (sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts))


def _dist(a, b) -> float:
    if any(math.isnan(x) for x in (*a, *b)):
        return math.inf
    return math.hypot(a[0] - b[0], a[1] - b[1])


if __name__ == "__main__":
    lk = lookup(" ".join(sys.argv[1:]) or "21101 Victory Blvd, Canoga Park, CA 91303")
    print("matched:", lk.matched)
    for k, f in lk.found.items():
        print(f"  {k:42} {f.value!s:30} [{f.source}] {f.note}")
    for l, v, s in lk.facts:
        print(f"  fact: {l}: {v} [{s}]")
    for m in lk.missing:
        print("  missing:", m)
    for e in lk.errors:
        print("  error:", e)
