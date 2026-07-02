"""
tools/ura_zone.py
=================
URA Master Plan land use zone lookup for Singapore properties.

Two strategies (tried in order):
  1. Keyword mapping  — maps raw land-zoning text from input data to URA codes.
                        No API call, instant, covers most clear-cut cases.
  2. OneMap coordinate lookup — queries the OneMap Themes API at the property's
                        geocoded lat/lon to retrieve the actual URA Master Plan
                        zone for that parcel. Requires an OneMap API token.

Usage
-----
    from tools.ura_zone import resolve_ura_zone

    zone = resolve_ura_zone(
        raw_zoning="Industrial",          # text from input data
        lon=103.8398, lat=1.3020,         # geocoded coordinates
        onemap_token="eyJ...",            # from shared_settings.json (optional)
    )
    # → "B1" / "B2" / "C" / "R" / None
"""

import json
import math
import re
import ssl
import time
import urllib.parse
import urllib.request

# ── SSL context (same as generate_comps_map_base) ────────────────────────────
try:
    _SSL_CTX = ssl.create_default_context()
    _SSL_CTX.check_hostname = False
    _SSL_CTX.verify_mode    = ssl.CERT_NONE
except Exception:
    _SSL_CTX = None


# ── URA zone abbreviation table ───────────────────────────────────────────────
# Maps the full zone name returned by OneMap → standard abbreviation.
_ZONE_NAME_TO_CODE = {
    "commercial":                       "C",
    "commercial/residential":           "CR",
    "commercial & residential":         "CR",
    "residential":                      "R",
    "residential with commercial at 1st storey": "RCO",
    "business 1":                       "B1",
    "business 1 - white":               "B1-W",
    "business 2":                       "B2",
    "business 2 - white":               "B2-W",
    "business park":                    "BP",
    "business park - white":            "BP-W",
    "hotel":                            "H",
    "white":                            "W",
    "mixed use":                        "MU",
    "civic & community institution":    "CI",
    "educational institution":          "ED",
    "health & medical care":            "HMC",
    "place of worship":                 "PW",
    "sports & recreation":              "SR",
    "open space":                       "OS",
    "park":                             "PK",
    "transport facilities":             "TP",
    "utility":                          "U",
    "agriculture":                      "AG",
    "cemetery":                         "CE",
    "reserve site":                     "RS",
    "special use":                      "SU",
    "port / airport":                   "PA",
    "beach area":                       "BA",
}

# ── Keyword mapping rules ─────────────────────────────────────────────────────
# Each entry: (keywords_that_must_ALL_appear, zone_code)
# Checked in order — first match wins.
_KEYWORD_RULES = [
    # Industrial — B1/B2 must be checked before generic "industrial"
    ({"b1", "business 1"},                          "B1"),
    ({"b2", "business 2"},                          "B2"),
    ({"light industrial"},                          "B1"),
    ({"clean industrial"},                          "B1"),
    ({"general industrial"},                        "B2"),
    ({"heavy industrial"},                          "B2"),
    ({"business park", "white"},                    "BP-W"),
    ({"business park"},                             "BP"),
    # Commercial
    ({"commercial", "residential"},                 "CR"),
    ({"commercial"},                                "C"),
    # Residential
    ({"residential", "commercial"},                 "RCO"),
    ({"residential"},                               "R"),
    # Other
    ({"hotel"},                                     "H"),
    ({"white"},                                     "W"),
    ({"mixed use"},                                 "MU"),
    ({"civic"},                                     "CI"),
    ({"educational", "institution"},                "ED"),
    ({"health", "medical"},                         "HMC"),
    ({"worship"},                                   "PW"),
    ({"sports", "recreation"},                      "SR"),
    ({"open space"},                                "OS"),
    ({"park"},                                      "PK"),
    ({"transport"},                                 "TP"),
    ({"utility"},                                   "U"),
    ({"agriculture"},                               "AG"),
    ({"cemetery"},                                  "CE"),
]


def _keyword_map(raw: str) -> str | None:
    """Map raw zoning text to URA code via keyword rules. Returns None if ambiguous."""
    norm = raw.lower().strip()
    words = set(re.sub(r"[^\w\s]", " ", norm).split())
    for required_keywords, code in _KEYWORD_RULES:
        if all(
            any(kw in w for w in words)   # keyword appears in any word
            for kw in required_keywords
        ):
            return code
    return None


def _onemap_zone_lookup(lon: float, lat: float, token: str) -> str | None:
    """
    Query OneMap Themes API for the URA Master Plan zone at (lon, lat).
    Returns URA zone code string, or None if lookup fails.
    """
    geojson = json.dumps({"type": "Point", "coordinates": [lon, lat]})
    params  = urllib.parse.urlencode({
        "queryPointGeojson": geojson,
        "token":             token,
        "themeName":         "Urban Planning",
    })
    url = f"https://www.onemap.gov.sg/api/public/themefinder/queryByPoint?{params}"
    req = urllib.request.Request(url, headers={"User-Agent": "pgim-comps-map/1.0"})

    try:
        open_kwargs = {"timeout": 10}
        if _SSL_CTX is not None:
            open_kwargs["context"] = _SSL_CTX
        with urllib.request.urlopen(req, **open_kwargs) as resp:
            data = json.loads(resp.read())
    except Exception as e:
        print(f"  [ura_zone] OneMap lookup failed: {e}")
        return None

    # Parse response — look for land use / zone field
    features = data.get("SrchResults", []) or data.get("results", [])
    for feat in features:
        # Try common field names for land use zone
        for field in ("LU_DESC", "LANDUSE", "ZONE", "DESCRIPTION", "lu_desc", "zone"):
            val = feat.get(field, "")
            if val:
                norm = str(val).lower().strip()
                code = _ZONE_NAME_TO_CODE.get(norm)
                if code:
                    return code
                # Partial match
                for zone_name, code in _ZONE_NAME_TO_CODE.items():
                    if zone_name in norm:
                        return code
    return None


def resolve_ura_zone(raw_zoning: str, lon: float = None, lat: float = None,
                     onemap_token: str = "") -> str | None:
    """
    Resolve raw zoning text to a URA zone code.

    1. Tries keyword mapping on raw_zoning.
    2. If ambiguous (e.g. plain "Industrial") and coordinates + token are
       available, queries OneMap for the actual parcel zone.
    3. Returns None if resolution fails.
    """
    if not raw_zoning:
        return None

    code = _keyword_map(raw_zoning)
    if code:
        return code

    # Ambiguous — try coordinate lookup
    if lon is not None and lat is not None and onemap_token:
        code = _onemap_zone_lookup(lon, lat, onemap_token)
        if code:
            return code

    return None


def _lu_desc_to_code(lu: str) -> str | None:
    """Map a URA LU_DESC ('COMMERCIAL', 'BUSINESS 1') to a zone code, else a
    title-cased fallback of the raw land use."""
    if not lu:
        return None
    norm = lu.lower().strip()
    code = _ZONE_NAME_TO_CODE.get(norm)
    if code:
        return code
    for zone_name, c in _ZONE_NAME_TO_CODE.items():
        if zone_name in norm:
            return c
    return lu.title()   # readable fallback (e.g. 'Open Space')


def zone_from_coords(lon: float = None, lat: float = None) -> str | None:
    """Derive a URA zone code from the *local* Master Plan land use at a point.

    Token-free point-in-polygon lookup (works offline) — used to fill a comp's
    Land Zoning when the source data didn't provide one. Returns None if the
    GeoJSON is absent or the point isn't inside any parcel.
    """
    if lon is None or lat is None:
        return None
    try:
        from . import ura_landuse as _U
    except Exception:
        return None
    if not _U.available():
        return None
    return _lu_desc_to_code(_U.land_use_at(lon, lat))
