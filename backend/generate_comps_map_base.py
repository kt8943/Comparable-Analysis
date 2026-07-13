#!/usr/bin/env python3
"""
generate_comps_map_base.py
==========================
Shared geocoding and map-rendering engine used by all three comp-type
map modules (sales, rent, land).

    generate_sales_comps_map.py  ─┐
    generate_rent_comps_map.py   ─┼─ all import from here
    generate_land_comps_map.py   ─┘

No comp-type-specific logic lives here — just the Mapbox API calls
and Pillow pin-drawing code.
"""

import io
import json
import math
import re
import ssl
import time
import urllib.parse
import urllib.request
from pathlib import Path

import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── Windows SSL fix ──────────────────────────────────────────────────────────
# urllib.request on Windows can raise WinError 10054 (connection reset) when
# a corporate proxy intercepts HTTPS traffic and presents its own certificate.
# Python's default SSL verification rejects the proxy cert and the connection
# is reset. Disabling certificate verification resolves this on corporate
# Windows machines with SSL-inspecting proxies.
try:
    _SSL_CTX = ssl.create_default_context()
    _SSL_CTX.check_hostname = False
    _SSL_CTX.verify_mode    = ssl.CERT_NONE
except Exception:
    _SSL_CTX = None   # fallback: let urllib use its own context


# ═══════════════════════════════════════════════════════════════════════════════
# GEOCODING  — Mapbox Geocoding API v5  OR  Google Maps Geocoding API
#
# Which provider is used is controlled by shared_settings.json:
#   "geocoding_provider": "mapbox"   (default, no change needed)
#   "geocoding_provider": "google"   (requires "google_maps_key" in same file)
#
# All scan files call geocode_with_fallbacks() — the provider switch is
# transparent to callers.
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_property_text(raw: str) -> tuple:
    """
    Split a property cell into (clean_name, address_line).
    e.g. "CapitaSpring (Remaining 45% stake)\\n88 Market St"
         → ("CapitaSpring", "88 Market St")
    """
    lines = [ln.strip() for ln in raw.split("\n") if ln.strip()]
    name  = re.sub(r"\s*\([^)]*(?:stake|%)[^)]*\)", "", lines[0], flags=re.I).strip()
    addr  = lines[1] if len(lines) >= 2 else ""
    return name, addr


def geocode(query: str, token: str, country_code: str = "",
            bounds: tuple = None) -> tuple:
    """
    Return (lon, lat) for *query* using Mapbox Geocoding API.

    bounds : optional (lon_min, lat_min, lon_max, lat_max).
             When provided the result is checked against it.

    Retries up to 3 times with back-off to handle transient WinError 10054
    (connection reset) and other network hiccups on Windows.
    """
    encoded = urllib.parse.quote(query)
    url = (f"https://api.mapbox.com/geocoding/v5/mapbox.places/{encoded}.json"
           f"?limit=1&access_token={token}"
           + (f"&country={country_code}" if country_code else ""))
    req = urllib.request.Request(url, headers={"User-Agent": "pgim-comps-map/1.0"})

    last_err = None
    for attempt in range(3):
        try:
            open_kwargs = {"timeout": 15}
            if _SSL_CTX is not None:
                open_kwargs["context"] = _SSL_CTX
            with urllib.request.urlopen(req, **open_kwargs) as resp:
                data = json.loads(resp.read())
            break   # success — exit retry loop
        except Exception as e:
            last_err = e
            if attempt < 2:
                time.sleep(1.5 * (attempt + 1))   # 1.5 s, 3 s
            else:
                raise ValueError(
                    f"Geocoding failed after 3 attempts for {query!r}: {e}"
                ) from e

    features = data.get("features", [])
    if not features:
        raise ValueError(f"No geocoding result for: {query!r}")

    feat = features[0]

    lon, lat = feat["geometry"]["coordinates"]
    lon, lat = round(lon, 6), round(lat, 6)
    if bounds:
        lo, la, hi, ha = bounds
        if not (lo <= lon <= hi and la <= lat <= ha):
            raise ValueError(
                f"Geocoding result ({lon}, {lat}) is outside config bounds for '{query}'"
            )
    return lon, lat


def geocode_with_fallbacks(queries: list, token: str, country_code: str = "",
                           bounds: tuple = None) -> tuple:
    """Try each query in order; return (lon, lat, note) for first that succeeds."""
    last_err = None
    for q in queries:
        try:
            lon, lat = geocode(q, token, country_code, bounds=bounds)
            return lon, lat, "mapbox"
        except Exception as e:
            last_err = e
    raise ValueError(f"All geocoding attempts failed. Last: {last_err}")


# ── Google Maps Geocoding API ─────────────────────────────────────────────────

def _geocode_google(query: str, api_key: str, country_code: str = "",
                    bounds: tuple = None) -> tuple:
    """Return (lon, lat) using Google Maps Geocoding API."""
    params = {"address": query, "key": api_key}
    if country_code:
        params["components"] = f"country:{country_code.upper()}"
    url = ("https://maps.googleapis.com/maps/api/geocode/json?"
           + urllib.parse.urlencode(params))
    req = urllib.request.Request(url, headers={"User-Agent": "pgim-comps-map/1.0"})

    last_err = None
    for attempt in range(3):
        try:
            open_kwargs = {"timeout": 15}
            if _SSL_CTX is not None:
                open_kwargs["context"] = _SSL_CTX
            with urllib.request.urlopen(req, **open_kwargs) as resp:
                data = json.loads(resp.read())
            break
        except Exception as e:
            last_err = e
            if attempt < 2:
                time.sleep(1.5 * (attempt + 1))
            else:
                raise ValueError(
                    f"Google geocoding failed after 3 attempts for {query!r}: {e}"
                ) from e

    if data.get("status") != "OK":
        raise ValueError(
            f"Google geocoding: status={data.get('status')} for {query!r}")
    loc = data["results"][0]["geometry"]["location"]
    lon, lat = round(loc["lng"], 6), round(loc["lat"], 6)
    if bounds:
        lo, la, hi, ha = bounds
        if not (lo <= lon <= hi and la <= lat <= ha):
            raise ValueError(
                f"Google geocoding result ({lon},{lat}) outside bounds for '{query}'")
    return lon, lat


def _geocode_google_with_fallbacks(queries: list, api_key: str,
                                   country_code: str = "",
                                   bounds: tuple = None) -> tuple:
    """Try each query in order using Google Maps; return first that succeeds."""
    last_err = None
    for q in queries:
        try:
            return _geocode_google(q, api_key, country_code, bounds=bounds)
        except Exception as e:
            last_err = e
    raise ValueError(f"All Google geocoding attempts failed. Last: {last_err}")


# ── Kakao Geocoding API (Korea) ──────────────────────────────────────────────
# Requires REST API key from developers.kakao.com (free tier: 300k req/day).
# Tries address search first, then keyword/building-name search as fallback.
# Docs: https://developers.kakao.com/docs/latest/en/local/dev-guide

_KAKAO_BOUNDS = (124.6, 33.1, 131.9, 38.9)   # Korea bounding box


def _geocode_kakao(query: str, api_key: str, bounds: tuple = None) -> tuple:
    """Return (lon, lat) using Kakao Local API (Korea only)."""
    encoded = urllib.parse.quote(query)
    headers = {
        "User-Agent":    "pgim-comps-map/1.0",
        "Authorization": f"KakaoAK {api_key}",
    }

    last_err = None
    # Try address search first, then keyword search (catches building names)
    for endpoint in (
        f"https://dapi.kakao.com/v2/local/search/address.json?query={encoded}&size=1",
        f"https://dapi.kakao.com/v2/local/search/keyword.json?query={encoded}&size=1",
    ):
        data = None
        req  = urllib.request.Request(endpoint, headers=headers)
        for attempt in range(3):
            try:
                open_kwargs = {"timeout": 15}
                if _SSL_CTX is not None:
                    open_kwargs["context"] = _SSL_CTX
                with urllib.request.urlopen(req, **open_kwargs) as resp:
                    data = json.loads(resp.read())
                break
            except Exception as e:
                last_err = e
                if attempt < 2:
                    time.sleep(1.5 * (attempt + 1))
                # else: break inner loop, try next endpoint

        if data is None:
            continue
        docs = data.get("documents", [])
        if docs:
            doc = docs[0]
            lon = round(float(doc.get("x", 0)), 6)
            lat = round(float(doc.get("y", 0)), 6)
            check_bounds = bounds or _KAKAO_BOUNDS
            lo, la, hi, ha = check_bounds
            if not (lo <= lon <= hi and la <= lat <= ha):
                raise ValueError(
                    f"Kakao result ({lon}, {lat}) is outside Korea bounds for '{query}'"
                )
            return lon, lat

    raise ValueError(
        f"No Kakao result for: {query!r}"
        + (f" (last error: {last_err})" if last_err else "")
    )


def _geocode_kakao_with_fallbacks(queries: list, api_key: str,
                                   bounds: tuple = None) -> tuple:
    """Try each query in order using Kakao; return first that succeeds."""
    last_err = None
    for q in queries:
        try:
            return _geocode_kakao(q, api_key, bounds=bounds)
        except Exception as e:
            last_err = e
    raise ValueError(f"All Kakao geocoding attempts failed. Last: {last_err}")


# ── OneMap Geocoding API (Singapore) ─────────────────────────────────────────
# Free, no API key required. Best building-level accuracy for Singapore.
# Docs: https://www.onemap.gov.sg/apidocs/

_ONEMAP_BOUNDS = (103.6, 1.15, 104.05, 1.48)   # Singapore bounding box


def _geocode_onemap(query: str, bounds: tuple = None) -> tuple:
    """Return (lon, lat) using OneMap Search API (Singapore only)."""
    encoded = urllib.parse.quote(query)
    url = (f"https://www.onemap.gov.sg/api/common/elastic/search"
           f"?searchVal={encoded}&returnGeom=Y&getAddrDetails=Y&pageNum=1")
    last_err = None
    for attempt in range(3):
        try:
            resp = requests.get(url, headers={"User-Agent": "pgim-comps-map/1.0"},
                                timeout=15, verify=False)
            data = resp.json()
            break
        except Exception as e:
            last_err = e
            if attempt < 2:
                time.sleep(1.5 * (attempt + 1))
            else:
                raise ValueError(
                    f"OneMap geocoding failed after 3 attempts for {query!r}: {e}"
                ) from e

    results = data.get("results", [])
    if not results:
        raise ValueError(f"No OneMap result for: {query!r}")

    # Prefer result where BUILDING matches the query exactly (over tenant/unit matches).
    query_upper = query.strip().upper()
    best = next(
        (r for r in results
         if str(r.get("BUILDING", "")).strip().upper() == query_upper),
        results[0],   # fallback: first result
    )

    # Relevance check: ALL meaningful words from the query must appear in the
    # result's BUILDING or ADDRESS. Prevents OneMap from matching "Bukit Panjang Plaza"
    # to a CBD plaza just because "plaza" appears.
    _STOPWORDS = {"the", "a", "an", "of", "in", "at", "on", "by", "to",
                  "for", "and", "or", "with", "lot", "lots", "strata",
                  "interest", "portfolio", "comprising", "three", "one"}
    query_words = [
        w for w in re.sub(r"[^\w\s]", " ", query_upper.lower()).split()
        if len(w) > 2 and w not in _STOPWORDS
    ]
    result_text = (
        str(best.get("BUILDING", "")) + " " + str(best.get("ADDRESS", ""))
    ).upper()
    result_words = set(re.sub(r"[^\w\s]", " ", result_text.lower()).split())
    if query_words and not all(w in result_words for w in query_words):
        raise ValueError(
            f"OneMap result '{best.get('BUILDING')}' does not fully match query '{query}'"
        )

    lon = round(float(best["LONGITUDE"]), 6)
    lat = round(float(best["LATITUDE"]), 6)

    check_bounds = bounds or _ONEMAP_BOUNDS
    lo, la, hi, ha = check_bounds
    if not (lo <= lon <= hi and la <= lat <= ha):
        raise ValueError(
            f"OneMap result ({lon}, {lat}) is outside Singapore bounds for '{query}'"
        )
    return lon, lat


def _geocode_onemap_with_fallbacks(queries: list, bounds: tuple = None) -> tuple:
    """Try each query in order using OneMap; return first that succeeds."""
    last_err = None
    for q in queries:
        try:
            return _geocode_onemap(q, bounds=bounds)
        except Exception as e:
            last_err = e
    raise ValueError(f"All OneMap geocoding attempts failed. Last: {last_err}")


# ── Unified entry point ───────────────────────────────────────────────────────

def _load_shared_settings() -> dict:
    """Load shared_settings.json relative to this file."""
    p = Path(__file__).parent.parent / "configs" / "shared_settings.json"
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def geocode_any(queries: list, mapbox_token: str, country_code: str = "",
                bounds: tuple = None) -> tuple:
    """
    Geocode using whichever provider is configured in shared_settings.json.

    geocoding_provider = "mapbox"   → Mapbox (default, global)
    geocoding_provider = "google"   → Google Maps (requires google_maps_key)
    geocoding_provider = "onemap"   → OneMap Singapore (free, best SG accuracy)
    geocoding_provider = "kakao"    → Kakao Maps (Korea, requires kakao_api_key)

    Falls back to Mapbox if the selected provider fails or is misconfigured.
    """
    settings = _load_shared_settings()
    provider = settings.get("geocoding_provider", "mapbox").lower()
    _fallback_reason = ""

    if provider == "onemap":
        try:
            lon, lat = _geocode_onemap_with_fallbacks(queries, bounds=bounds)
            return lon, lat, "onemap"
        except Exception as e:
            _fallback_reason = str(e)
            print(f"  [geocode] OneMap failed ({e}) — falling back to Mapbox")

    elif provider == "kakao":
        api_key = settings.get("kakao_api_key", "")
        if api_key:
            try:
                lon, lat = _geocode_kakao_with_fallbacks(queries, api_key, bounds=bounds)
                return lon, lat, "kakao"
            except Exception as e:
                _fallback_reason = str(e)
                print(f"  [geocode] Kakao failed ({e}) — falling back to Mapbox")
        else:
            _fallback_reason = "kakao_api_key not set"
            print("  [geocode] kakao_api_key not set — falling back to Mapbox")

    elif provider == "google":
        api_key = settings.get("google_maps_key", "")
        if api_key:
            try:
                lon, lat = _geocode_google_with_fallbacks(
                    queries, api_key, country_code, bounds=bounds)
                return lon, lat, "google"
            except Exception as e:
                _fallback_reason = str(e)
                print(f"  [geocode] Google Maps failed ({e}) — falling back to Mapbox")
        else:
            _fallback_reason = "google_maps_key not set"
            print("  [geocode] google_maps_key not set — falling back to Mapbox")

    lon, lat, _ = geocode_with_fallbacks(queries, mapbox_token, country_code, bounds=bounds)
    note = f"mapbox ({_fallback_reason})" if _fallback_reason else "mapbox"
    return lon, lat, note


# Approximate geographic centroids of supported markets. When a geocoder can't find
# a specific place it often falls back to the country centroid — a strong signal the
# coordinate is INVALID (the comp shouldn't be trusted / plotted there).
_COUNTRY_CENTROIDS = {
    "sg": (103.8198, 1.3521),
    "kr": (127.7669, 35.9078),
    "jp": (138.2529, 36.2048),
    "hk": (114.1095, 22.3964),
    "cn": (104.1954, 35.8617),
    "au": (133.7751, -25.2744),
}

# Country NAME → ISO code. A deal that sets only country_name still geocodes with the
# correct country restriction — a missing/empty country_code is a common cause of comps
# resolving to a same-named place in the WRONG country (e.g. "Capital Square" → USA).
_NAME_TO_CC = {
    "singapore": "sg", "south korea": "kr", "korea": "kr", "republic of korea": "kr",
    "japan": "jp", "hong kong": "hk", "hongkong": "hk", "china": "cn",
    "people's republic of china": "cn", "taiwan": "tw", "australia": "au",
    "malaysia": "my", "indonesia": "id", "thailand": "th", "vietnam": "vn",
    "philippines": "ph", "india": "in", "new zealand": "nz",
    "united states": "us", "usa": "us", "united kingdom": "gb", "uk": "gb",
}


def country_code_from_name(country_name: str) -> str:
    """ISO country code for a country name (e.g. 'Singapore' → 'sg'); '' if unknown."""
    return _NAME_TO_CC.get((country_name or "").strip().lower(), "")


def clean_property_name(name: str) -> str:
    """Deterministic safety net for a source cell that stacks 'Building Name⏎Submarket'
    (e.g. 'Shaw Tower⏎Bugis'). Keep the first line as the property name; re-append a
    later line ONLY if it looks like a genuine qualifier (has a digit or a parenthesis,
    e.g. '(one-third partial interest)') rather than a bare place/submarket name. Keeps
    the Property column clean and stops a submarket hijacking the geocoder."""
    segs = [s.strip() for s in str(name or "").replace("\r", "\n").split("\n") if s.strip()]
    if len(segs) <= 1:
        return segs[0] if segs else ""
    keep = [segs[0]]
    for s in segs[1:]:
        if any(ch.isdigit() for ch in s) or "(" in s or ")" in s:
            keep.append(s)
    return " ".join(keep).strip()


def near_country_centroid(lon, lat, country_code: str, tol_km: float = 1.5) -> bool:
    """True if (lon,lat) sits ~on the country centroid → a likely failed geocode."""
    if lon is None or lat is None:
        return False
    c = _COUNTRY_CENTROIDS.get((country_code or "").lower())
    if not c:
        return False
    dlon = math.radians(lon - c[0]) * math.cos(math.radians(lat))
    dlat = math.radians(lat - c[1])
    return math.hypot(dlon, dlat) * 6371.0 <= tol_km


# Connector phrases that signal "[descriptor] <connector> [Property]" naming.
# Longer alternatives must appear before shorter ones (forming part of > forming).
_CONNECTORS = re.compile(
    r'\s+(?:'
    r'of'
    r'|comprising'
    r'|within'
    r'|forming\s+part\s+of'
    r'|being\s+part\s+of'
    r'|forming'
    r'|making\s+up(?:\s+of)?'
    r'|known\s+as'
    r')\s+',
    re.IGNORECASE,
)

# Real-estate part-descriptor words that guard the connector match — the words
# before the connector must include at least one of these to avoid false positives.
_PART_WORDS = frozenset({
    "component", "portion", "interest", "wing", "floor", "unit",
    "element", "part", "commercial", "office", "retail",
    "residential", "hotel", "industrial", "strata", "tower",
    "block", "section", "tranche",
    "property", "land", "site", "premises", "lots", "units",
})


def build_geocode_queries(name: str, addr: str, suffix: str) -> list:
    """
    Build an ordered list of geocoding query strings for a property.

    Tries progressively simpler variants so that:
    - "The Centrepoint rear block, Singapore" → OneMap may miss it
    - "The Centrepoint, Singapore"            → OneMap finds it correctly
    - "Commercial Component of CapitaSpring"  → tries "CapitaSpring" as well

    suffix is typically ", Singapore" or "".
    """
    seen = set()
    queries = []

    def _add(q):
        q = q.strip().strip(",").strip()
        if q and q not in seen:
            seen.add(q)
            queries.append(q)

    _sfx = suffix if suffix else ""

    # 1. Address (most precise)
    if addr:
        _add(f"{addr}{_sfx}" if _sfx.lower() not in addr.lower() else addr)

    if name:
        # 2. Remove parenthetical DEAL-STRUCTURE annotations — "(remaining 45%
        # stake)", "(50% interest)", "(Q3 2024)" — from EVERY query. These have no
        # locational value; a geocoder can only ignore them at best (and may match
        # worse). Unlike "- Tower A"/"- Gangnam", which a geocoder can actually use.
        _core = (re.sub(r"\s*\(.*$", "", name).strip().strip(",").strip("-").strip()
                 or name)

        # 3. Core name (parenthetical removed) FIRST — a good geocoder (e.g. Google)
        # resolves "Marina Bay - Tower A" / "Weave Place - Gangnam Station" directly
        # and more precisely than a further-stripped version.
        base = f"{_core}{_sfx}" if _sfx.lower() not in _core.lower() else _core
        _add(base)

        # 4. Building-part-stripped core as a FALLBACK for weaker geocoders that
        # miss the full string. Only a " - X" suffix that is a building-part
        # descriptor (Tower/Block/Phase/Wing/…) is removed, so distinct BRANCH
        # names like "Weave Place - Gangnam Station" / " - Hoegi" are preserved.
        # Korean hyphenated words ('Samseong-dong', '-gu', '-ro') are kept.
        _stripped = re.sub(
            r"\s+[-–]\s*(?:tower|block|phase|wing|unit|lot|level|floor|fl|"
            r"north|south|east|west|rear|front|main)\b.*$"
            r"|\b(?:rear|front|main)\s+block\b"
            r"|\btower\s+[a-z0-9]+\b"
            r"|\bblock\s+[a-z0-9]+\b"
            r"|\bphase\s+[0-9]+\b"
            r"|\bwing\s+[a-z]\b"
            r"|\b(?:north|south|east|west)\s+tower\b",
            "", _core, flags=re.I,
        ).strip().strip(",").strip("-").strip()
        if _stripped and _stripped.lower() != _core.lower():
            _add(f"{_stripped}{_sfx}" if _sfx.lower() not in _stripped.lower() else _stripped)

        # 3. "[descriptor] <connector> [Property]" — extract just the property name
        _search = _stripped if _stripped else name
        _m = _CONNECTORS.search(_search)
        if _m:
            _pre_words = set(re.sub(r"[^\w]", " ", _search[:_m.start()].lower()).split())
            if _pre_words & _PART_WORDS:
                _core = re.sub(r"\s*\(.*$", "", _search[_m.end():]).strip()
                if _core and _core.lower() != _search.lower():
                    _add(f"{_core}{_sfx}" if _sfx.lower() not in _core.lower() else _core)

    return queries


# ═══════════════════════════════════════════════════════════════════════════════
# MAP RENDERING  — Mapbox Static Images API + Pillow pin drawing
# ═══════════════════════════════════════════════════════════════════════════════

_SUBJECT_COLOR   = "c0392b"   # bold red
_COMP_COLOR      = "1a5276"   # dark navy — matches Excel table headers
_CUSTOM_PIN_RADII = {
    "xl":  22,   # ~50 % bigger than Mapbox pin-l
    "xxl": 34,   # ~130 % bigger than Mapbox pin-l
}


def _marker(size: str, label: str, color: str, lon: float, lat: float) -> str:
    return f"pin-{size}-{label}+{color}({lon},{lat})"


def _haversine_km(lon1: float, lat1: float, lon2: float, lat2: float) -> float:
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2))
         * math.sin(dlon / 2) ** 2)
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def _fit_zoom(lons: list, lats: list,
              width: int, height: int, padding: int) -> tuple:
    """Return (center_lon, center_lat, zoom) fitting all coords in the image."""
    min_lon, max_lon = min(lons), max(lons)
    min_lat, max_lat = min(lats), max(lats)
    c_lon = (min_lon + max_lon) / 2
    c_lat = (min_lat + max_lat) / 2
    d_lon = max_lon - min_lon
    eff_w = max(width  - 2 * padding, 64)
    eff_h = max(height - 2 * padding, 64)

    def _merc(lat):
        lr = math.radians(max(-85.0, min(85.0, lat)))
        return math.log(math.tan(math.pi / 4 + lr / 2))

    d_merc = _merc(max_lat) - _merc(min_lat)
    z_lon  = math.log2(eff_w * 360 / (256 * d_lon))          if d_lon  > 1e-5 else 20.0
    z_lat  = math.log2(eff_h * 2 * math.pi / (256 * d_merc)) if d_merc > 1e-5 else 20.0
    zoom   = max(1.0, min(min(z_lon, z_lat), 18.0))
    return c_lon, c_lat, zoom


def _lonlat_to_px(lon: float, lat: float,
                  c_lon: float, c_lat: float,
                  zoom: float, width: int, height: int) -> tuple:
    """Convert (lon, lat) → image pixel (x, y) in logical (1×) pixels."""
    scale = 256 * (2 ** zoom)

    def _world(lo, la):
        la_r = math.radians(max(-85.0, min(85.0, la)))
        wx = (lo + 180) / 360 * scale
        wy = (1 - math.log(math.tan(math.pi / 4 + la_r / 2)) / math.pi) / 2 * scale
        return wx, wy

    cx, cy = _world(c_lon, c_lat)
    px, py = _world(lon, lat)
    return (px - cx) + width / 2, (py - cy) + height / 2


def _load_font(size: int):
    """Load a system font; fall back to PIL built-in."""
    try:
        from PIL import ImageFont
        for path in [
            "/System/Library/Fonts/HelveticaNeue.ttc",
            "/System/Library/Fonts/Helvetica.ttc",
            "/System/Library/Fonts/ArialHB.ttc",
            "/System/Library/Fonts/SFNS.ttf",
        ]:
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
        return ImageFont.load_default()
    except Exception:
        return None


def _draw_pin(draw, x: float, y: float, label: str,
              color_hex: str, radius: int):
    """Draw a Mapbox-style teardrop pin onto a Pillow ImageDraw surface."""
    r  = radius
    cx = round(x)
    cy = round(y) - r
    fill  = "#" + color_hex
    white = "#ffffff"

    half_base = round(r * 0.52)
    join_y    = cy + round(r * 0.65)
    tip_y     = round(y) + round(r * 0.12)
    draw.polygon(
        [(cx - half_base, join_y), (cx + half_base, join_y), (cx, tip_y)],
        fill=fill,
    )
    ring = max(3, round(r * 0.13))
    draw.ellipse([cx - r - ring, cy - r - ring,
                  cx + r + ring, cy + r + ring], fill=white)
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill)
    font_size = max(10, round(r * 0.88))
    font = _load_font(font_size)
    draw.text((cx, cy), label, fill=white, anchor="mm", font=font)


def _annotate_map(img, geocoding_provider: str):
    """Draw a small credit label at the bottom-left of the map image."""
    try:
        from PIL import ImageDraw
    except ImportError:
        return img

    _geo_labels = {
        "mapbox":  "Mapbox",
        "onemap":  "OneMap (Singapore)",
        "google":  "Google Maps",
        "kakao":   "Kakao Maps (Korea)",
    }
    geo_label = _geo_labels.get(geocoding_provider.lower(), geocoding_provider.title())
    text = f"Geocoding: {geo_label}   |   Map: Mapbox"

    draw  = ImageDraw.Draw(img, "RGBA")
    w, h  = img.size
    font  = _load_font(max(24, h // 50))   # scale with image size
    pad   = 20

    try:
        bbox  = draw.textbbox((0, 0), text, font=font)
        tw    = bbox[2] - bbox[0]
        th    = bbox[3] - bbox[1]
    except AttributeError:
        tw, th = draw.textsize(text, font=font)

    x = pad
    y = h - th - pad * 2
    draw.rectangle([x - 6, y - 6, x + tw + 6, y + th + 6],
                   fill=(0, 0, 0, 140))
    draw.text((x, y), text, fill=(255, 255, 255, 230), font=font)
    return img


def render_map(subject_lonlat: tuple,
               comps: list,           # [(marker_label, lon, lat), ...]
               token: str,
               output_path: str,
               style: str    = "streets-v12",
               width: int    = 1200,
               height: int   = 900,
               padding: int  = 100,
               pin_size: str = "xl",
               plot_subject: bool = True,
               **_kwargs):
    """
    Download a Mapbox static map PNG with pin markers.

    pin_size = "l"   → Mapbox built-in pin-l (no Pillow required)
    pin_size = "xl"/"xxl" → custom oversized pins drawn with Pillow
    """
    slon, slat = subject_lonlat
    geocoding_provider = _load_shared_settings().get("geocoding_provider", "mapbox")
    # When the subject star is hidden, render comps in red (no subject to contrast).
    _comp_clr = _COMP_COLOR if plot_subject else _SUBJECT_COLOR

    # Normalise comps to (label, lon, lat, color_hex). A comp may carry an explicit
    # per-pin colour as an optional 4th element (from the geo sidecar's "color");
    # otherwise it falls back to the default comp colour.
    def _resolve_clr(v):
        if not v:
            return _comp_clr
        v = str(v).lower()
        if v in ("red", "subject"):
            return _SUBJECT_COLOR
        if v in ("navy", "blue", "comp"):
            return _COMP_COLOR
        return v.lstrip("#")   # already a hex string
    comps = [(it[0], it[1], it[2], _resolve_clr(it[3] if len(it) >= 4 else None))
             for it in comps]

    if pin_size not in _CUSTOM_PIN_RADII:
        sz = "l"
        overlays = []
        for lbl, lon, lat, chex in comps:
            overlays.append(_marker(sz, lbl, chex, lon, lat))
        if plot_subject:
            overlays.append(_marker(sz, "star", _SUBJECT_COLOR, slon, slat))
        overlay_str = ",".join(overlays)
        url = (
            f"https://api.mapbox.com/styles/v1/mapbox/{style}/static/"
            f"{overlay_str}/auto/{width}x{height}@2x"
            f"?padding={padding}&access_token={token}"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "pgim-comps-map/1.0"})
        _open_kwargs = {"timeout": 30}
        if _SSL_CTX is not None:
            _open_kwargs["context"] = _SSL_CTX
        with urllib.request.urlopen(req, **_open_kwargs) as resp:
            img_bytes = resp.read()
        content_type = resp.headers.get("Content-Type", "")
        if "image" not in content_type:
            raise RuntimeError(f"Unexpected response ({content_type}): {img_bytes[:200]}")
        with open(output_path, "wb") as f:
            f.write(img_bytes)
        print(f"  Saved → {output_path}  ({len(img_bytes)//1024} KB, {width*2}×{height*2}px @2x)")
        return

    # Custom oversized pins via Pillow
    try:
        from PIL import Image as PILImage, ImageDraw
    except ImportError:
        raise ImportError("Pillow required for pin_size 'xl'/'xxl'.\n  pip install Pillow")

    radius_logical = _CUSTOM_PIN_RADII[pin_size]
    all_lons = ([slon] if plot_subject else []) + [lo for _, lo, _, _ in comps]
    all_lats = ([slat] if plot_subject else []) + [la for _, _, la, _ in comps]
    if not all_lons:  # subject hidden and no comps — fall back to subject for a valid view
        all_lons, all_lats = [slon], [slat]
    c_lon, c_lat, zoom = _fit_zoom(all_lons, all_lats, width, height, padding)
    zoom = round(zoom, 3)
    print(f"  Viewport: centre ({c_lon:.5f}, {c_lat:.5f})  zoom {zoom:.2f}")

    url = (
        f"https://api.mapbox.com/styles/v1/mapbox/{style}/static/"
        f"{c_lon},{c_lat},{zoom}/{width}x{height}@2x"
        f"?access_token={token}"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "pgim-comps-map/1.0"})
    _open_kwargs = {"timeout": 30}
    if _SSL_CTX is not None:
        _open_kwargs["context"] = _SSL_CTX
    with urllib.request.urlopen(req, **_open_kwargs) as resp:
        img_bytes = resp.read()
    content_type = resp.headers.get("Content-Type", "")
    if "image" not in content_type:
        raise RuntimeError(f"Unexpected response ({content_type}): {img_bytes[:200]}")

    img   = PILImage.open(io.BytesIO(img_bytes)).convert("RGBA")
    draw  = ImageDraw.Draw(img)
    scale2 = 2
    r_px   = radius_logical * scale2

    for lbl, lon, lat, chex in comps:
        px, py = _lonlat_to_px(lon, lat, c_lon, c_lat, zoom, width, height)
        _draw_pin(draw, px * scale2, py * scale2, lbl, chex, r_px)

    if plot_subject:
        sx, sy = _lonlat_to_px(slon, slat, c_lon, c_lat, zoom, width, height)
        _draw_pin(draw, sx * scale2, sy * scale2, "★", _SUBJECT_COLOR, r_px)

    img = img.convert("RGB")
    img.save(output_path, format="PNG")
    fsize = Path(output_path).stat().st_size
    print(f"  Saved → {output_path}  ({fsize//1024} KB, {width*2}×{height*2}px @2x)")
