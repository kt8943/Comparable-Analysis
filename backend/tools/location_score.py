"""Location competitiveness score (Singapore) from URA Master Plan proximity.

Scores a comp's location vs the subject (subject = 0) on sector-specific
proximity factors, normalised fixed-vs-subject to -1..1, then labelled:
    |s| <= 0.3 -> "Comparable"  ;  s > 0.3 -> "Superior"  ;  s < -0.3 -> "Inferior"

Only used for SG comps geocoded by OneMap; others get a blank Location (the
caller passes _geo_provider).  All proximity data is local (URA GeoJSON) — no
network/token at runtime.
"""

import json
import ssl
import urllib.parse
import urllib.request

from . import ura_landuse as U

try:
    _OM_CTX = ssl.create_default_context()
    _OM_CTX.check_hostname = False
    _OM_CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _OM_CTX = None


def _onemap_geocode(name: str, addr: str):
    """Geocode an SG property via OneMap (public search; address first, then name).

    Proximity is computed against OneMap-resolved coordinates for SG precision/
    consistency — separate from the map pin, which uses the sidebar provider.
    Returns (lon, lat) or None (None => not an SG/OneMap-resolvable location).
    """
    for q in (addr, name):
        q = (q or "").strip()
        if not q:
            continue
        try:
            url = ("https://www.onemap.gov.sg/api/common/elastic/search?searchVal="
                   + urllib.parse.quote(q) + "&returnGeom=Y&getAddrDetails=N&pageNum=1")
            req = urllib.request.Request(url, headers={"User-Agent": "pgim-comps/1.0"})
            kw = {"timeout": 15}
            if _OM_CTX is not None:
                kw["context"] = _OM_CTX
            res = json.loads(urllib.request.urlopen(req, **kw).read()).get("results", [])
            if res:
                return (float(res[0]["LONGITUDE"]), float(res[0]["LATITUDE"]))
        except Exception:
            continue
    return None


def _sector_match(subject_sector: str, comp_sector: str) -> bool:
    """Comps are scored only against same-sector subjects. 'mixed' matches anything
    (a mixed deal/comp overlaps if any one use matches)."""
    if not comp_sector:
        return True  # unknown comp sector — don't exclude
    return (comp_sector == subject_sector
            or subject_sector == "mixed" or comp_sector == "mixed")


# ── OneMap "Tourist Attractions" theme (hospitality proximity) ────────────────
from pathlib import Path as _Path
_TOURISM_CACHE = _Path(__file__).resolve().parents[1] / "data" / "_onemap_tourism.json"
_tourism_pts = None


def _tourism_points() -> list:
    """Tourist-attraction points (lon, lat) from OneMap, cached to disk."""
    global _tourism_pts
    if _tourism_pts is not None:
        return _tourism_pts
    if _TOURISM_CACHE.exists():
        try:
            _tourism_pts = [tuple(p) for p in json.loads(_TOURISM_CACHE.read_text(encoding="utf-8"))]
            return _tourism_pts
        except Exception:
            pass
    _tourism_pts = []
    try:
        from .onemap_auth import get_onemap_token
        tok = get_onemap_token()
        if tok:
            url = ("https://www.onemap.gov.sg/api/public/themesvc/"
                   "retrieveTheme?queryName=tourism")
            req = urllib.request.Request(url, headers={"Authorization": tok})
            kw = {"timeout": 25}
            if _OM_CTX is not None:
                kw["context"] = _OM_CTX
            res = json.loads(urllib.request.urlopen(req, **kw).read()).get("SrchResults", [])
            for item in (res[1:] if res else []):   # res[0] is theme metadata
                ll = str(item.get("LatLng", ""))
                if "," in ll:
                    try:
                        lat, lon = ll.split(",")[:2]
                        _tourism_pts.append((float(lon), float(lat)))
                    except Exception:
                        pass
            try:
                _TOURISM_CACHE.write_text(json.dumps(_tourism_pts), encoding="utf-8")
            except Exception:
                pass
    except Exception:
        pass
    return _tourism_pts


def _tourism_within(lon: float, lat: float, radius_km: float = 1.0) -> int:
    return sum(1 for (lo, la) in _tourism_points()
               if U._hav(lon, lat, lo, la) <= radius_km)

# CBD nodes — office competitiveness rises with closeness to a CBD.
_CBD_NODES = [
    (103.8519, 1.2830),  # CBD — Raffles Place / Marina Bay
    (103.7220, 1.3330),  # Jurong Lake District (2nd CBD)
]

# Regional / sub-regional centres (malls cluster here) — retail competitiveness
# rises with closeness to one of these.
_REGIONAL_NODES = [
    (103.8330, 1.3040),  # Orchard
    (103.7420, 1.3330),  # Jurong East / JLD
    (103.9450, 1.3540),  # Tampines
    (103.7860, 1.4370),  # Woodlands
    (103.8920, 1.3180),  # Paya Lebar
    (103.9010, 1.4050),  # Punggol
    (103.8480, 1.3500),  # Bishan / Ang Mo Kio
]


def _sector_key(asset_class: str) -> str:
    a = (asset_class or "").lower()
    if any(k in a for k in ("logistic", "industrial", "warehouse")):
        return "industrial"
    if "data" in a:
        return "data_centre"
    if any(k in a for k in ("retail", "mall", "shop")):
        return "retail"
    if any(k in a for k in ("hospitality", "hotel")):
        return "hospitality"
    if "office" in a:
        return "office"
    if "mixed" in a:
        return "mixed"
    return "office"  # sensible default


def _factors(lon: float, lat: float, sector: str) -> list:
    """Return [(value, higher_is_better), ...] for the sector."""
    if sector in ("industrial", "data_centre"):
        return [(U.count_within(lon, lat, "business", 1.0), True),     # industrial network
                (U.nearest_km(lon, lat, "port_airport"),    False)]    # transport access
    if sector == "office":
        cbd = min(U._hav(lon, lat, n[0], n[1]) for n in _CBD_NODES)
        return [(cbd,                                          False),  # closer to CBD
                (U.count_within(lon, lat, "commercial", 1.0),  True)]   # commercial cluster
    if sector == "retail":
        reg = min(U._hav(lon, lat, n[0], n[1]) for n in _REGIONAL_NODES)
        return [(U.count_within(lon, lat, "residential", 1.0), True),   # residential catchment
                (reg,                                          False)]  # closer to regional centre
    if sector == "hospitality":
        return [(_tourism_within(lon, lat, 1.0),               True),   # tourist draw
                (U.count_within(lon, lat, "commercial", 1.0),  True)]   # commercial cluster
    if sector == "mixed":
        cbd = min(U._hav(lon, lat, n[0], n[1]) for n in _CBD_NODES)
        combined = (U.count_within(lon, lat, "residential", 1.0)
                    + U.count_within(lon, lat, "commercial", 1.0))
        return [(cbd, False), (combined, True)]
    return []


_D_REF = 5.0   # km — reference scale for distance factors
_K     = 10    # smoothing for count factors (dampens small-number noise)


def _pair(comp_v: float, subj_v: float, higher_better: bool) -> float:
    """Fixed comp-vs-subject score in [-1, 1]; 0 when equal.

    Counts (higher-better): smoothed relative difference. Distances
    (lower-better): difference scaled by a fixed 5km reference (so a subject that
    sits ~on a landmark, distance~0, doesn't force every comp to -1).
    """
    if higher_better:
        return (comp_v - subj_v) / (comp_v + subj_v + _K)
    return max(-1.0, min(1.0, (subj_v - comp_v) / _D_REF))


def score(comp_lonlat: tuple, subj_lonlat: tuple, asset_class: str):
    """Return a -1..1 location score (subject = 0), or None if unavailable."""
    if not U.available() or comp_lonlat[0] is None or subj_lonlat[0] is None:
        return None
    sector = _sector_key(asset_class)
    cf = _factors(comp_lonlat[0], comp_lonlat[1], sector)
    sf = _factors(subj_lonlat[0], subj_lonlat[1], sector)
    if not cf or not sf:
        return None
    pairs = [_pair(c[0], s[0], c[1]) for c, s in zip(cf, sf)]
    return round(sum(pairs) / len(pairs), 3)


def label(s) -> str:
    if s is None:
        return ""
    if s > 0.3:
        return "Superior"
    if s < -0.3:
        return "Inferior"
    return "Comparable"


def _coords(r: dict):
    """Coordinates for proximity: reuse the map-resolved lon/lat if present
    (Google/Mapbox — reliable & consistent with the plotted pin), else fall back
    to a fresh OneMap geocode on name/address."""
    lon, lat = r.get("lon"), r.get("lat")
    if lon is not None and lat is not None:
        try:
            return (float(lon), float(lat))
        except (TypeError, ValueError):
            pass
    name = (r.get("property_name")
            or str(r.get("raw_description", "")).split("\n")[0])
    return _onemap_geocode(name, r.get("address"))


def apply_location(records: list, subject_name: str, subject_addr: str,
                   asset_class: str, subj_lonlat: tuple = None) -> list:
    """Set each record's 'location' to the proximity label.

    Uses the map-resolved lon/lat already on the subject and each comp (the same
    coordinates plotted on the map — Google/Mapbox), falling back to a OneMap
    geocode only when coordinates are missing.  A comp is scored only if it (a)
    resolves to coordinates and (b) is the same sector as the subject; otherwise
    its Location is left blank.  SG-only (URA Master Plan proximity).
    """
    if not U.available():
        return records
    subj = None
    if subj_lonlat and subj_lonlat[0] is not None and subj_lonlat[1] is not None:
        try:
            subj = (float(subj_lonlat[0]), float(subj_lonlat[1]))
        except (TypeError, ValueError):
            subj = None
    if subj is None:
        subj = _onemap_geocode(subject_name, subject_addr)
    if subj is None:
        print("  [location] subject not geolocatable — Location left blank for all comps.")
        for r in records:
            r["location"] = ""
        return records
    subj_sector = _sector_key(asset_class)
    print(f"  [location] competitiveness vs subject (sector={subj_sector}, subject=0.000):")
    for r in records:
        name = str(r.get("property_name") or r.get("property")
                   or str(r.get("raw_description", "")).split("\n")[0])[:46]
        comp_sector = _sector_key(str(r.get("land_zoning") or r.get("asset_type")
                                      or r.get("sale_type") or asset_class))
        if not _sector_match(subj_sector, comp_sector):
            r["location"] = ""
            print(f"      {name:<46}   —      (blank: {comp_sector} ≠ {subj_sector})")
            continue
        c = _coords(r)
        if c is None or c[0] is None:
            r["location"] = ""
            print(f"      {name:<46}   —      (blank: not geolocatable)")
            continue
        s = score(c, subj, asset_class)
        r["location"] = label(s)
        _s = f"{s:+.3f}" if s is not None else "  n/a"
        print(f"      {name:<46}  {_s}  → {r['location'] or '—'}")
    return records
