"""URA Master Plan land-use proximity (Singapore, fully local / on-prem).

Reads the URA Master Plan GeoJSON (backend/data/MasterPlan2025.geojson),
buckets each zoning polygon by land-use category, caches polygon centroids to a
pickle, and answers proximity questions used by the Location score:
    count_within(lon, lat, bucket, radius_km)  — # parcels of a category within r
    nearest_km(lon, lat, bucket)               — distance to the nearest such parcel

No network/token at runtime — the GeoJSON is a static local file.
"""

import json
import math
import pickle
from pathlib import Path

_DATA    = Path(__file__).resolve().parents[1] / "data"
_GEOJSON = _DATA / "MasterPlan2025.geojson"
_CACHE   = _DATA / "_landuse_buckets.pkl"

# Map URA LU_DESC values into the buckets the Location score uses.
# A parcel may belong to more than one bucket (e.g. COMMERCIAL & RESIDENTIAL).
_BUCKETS = {
    "residential":  {"RESIDENTIAL", "RESIDENTIAL WITH COMMERCIAL AT 1ST STOREY",
                     "RESIDENTIAL / INSTITUTION", "COMMERCIAL & RESIDENTIAL"},
    "commercial":   {"COMMERCIAL", "COMMERCIAL & RESIDENTIAL", "COMMERCIAL / INSTITUTION",
                     "RESIDENTIAL WITH COMMERCIAL AT 1ST STOREY"},
    "business":     {"BUSINESS 1", "BUSINESS 2", "BUSINESS PARK",
                     "BUSINESS 1 - WHITE", "BUSINESS 2 - WHITE", "BUSINESS PARK - WHITE"},
    "hotel":        {"HOTEL"},
    "port_airport": {"PORT / AIRPORT"},
}

_cache = None  # {bucket: [(lon, lat), ...]}


def available() -> bool:
    return _GEOJSON.exists() or _CACHE.exists()


def _centroid(geom: dict):
    t = geom.get("type"); c = geom.get("coordinates")
    try:
        if t == "Polygon":
            ring = c[0]
        elif t == "MultiPolygon":
            ring = c[0][0]
        else:
            return None
        if not ring:
            return None
        xs = [p[0] for p in ring]; ys = [p[1] for p in ring]
        return (sum(xs) / len(xs), sum(ys) / len(ys))
    except Exception:
        return None


def _poly_area_km2(geom: dict, ref_lat: float = None) -> float:
    """Approx polygon area in km² (exterior rings only; holes ignored).
    Shoelace on lon/lat degrees → km² via a local equirectangular scale."""
    t = geom.get("type"); c = geom.get("coordinates")
    try:
        if t == "Polygon":
            rings = [c[0]]
        elif t == "MultiPolygon":
            rings = [poly[0] for poly in c]
        else:
            return 0.0
    except Exception:
        return 0.0
    area_deg2 = 0.0
    for ring in rings:
        if not ring or len(ring) < 3:
            continue
        s = 0.0
        n = len(ring)
        for i in range(n):
            x1, y1 = ring[i][0], ring[i][1]
            x2, y2 = ring[(i + 1) % n][0], ring[(i + 1) % n][1]
            s += x1 * y2 - x2 * y1
        area_deg2 += abs(s) / 2.0
        if ref_lat is None:
            ref_lat = sum(p[1] for p in ring) / n
    km_lat = 110.574
    km_lon = 111.320 * math.cos(math.radians(ref_lat if ref_lat is not None else 1.35))
    return area_deg2 * km_lat * km_lon


def _build() -> dict:
    data = json.loads(_GEOJSON.read_text(encoding="utf-8"))
    buckets = {k: [] for k in _BUCKETS}
    for f in data.get("features", []):
        lu   = str(f.get("properties", {}).get("LU_DESC", "")).strip().upper()
        geom = f.get("geometry", {})
        cen  = _centroid(geom)
        if not cen:
            continue
        area = _poly_area_km2(geom, cen[1])            # km² (for coverage weighting)
        for b, descs in _BUCKETS.items():
            if lu not in descs:
                continue
            # PORT / AIRPORT: keep only parcels inside mainland Singapore's bounds —
            # this drops offshore stray parcels (e.g. ~104.08°E, ~104.41°E / Pedra
            # Branca) so the industrial "nearest freight node" distance reflects only
            # the real hubs (Tuas, Jurong, PSA/Keppel, Changi, Seletar).
            if b == "port_airport" and not (103.60 <= cen[0] <= 104.05
                                            and 1.15 <= cen[1] <= 1.48):
                continue
            buckets[b].append((cen[0], cen[1], area))
    try:
        with open(_CACHE, "wb") as fh:
            pickle.dump(buckets, fh)
    except Exception:
        pass
    return buckets


def _get() -> dict:
    global _cache
    if _cache is not None:
        return _cache
    if _CACHE.exists():
        try:
            with open(_CACHE, "rb") as fh:
                _cache = pickle.load(fh)
                return _cache
        except Exception:
            pass
    _cache = _build() if _GEOJSON.exists() else {k: [] for k in _BUCKETS}
    return _cache


def _hav(lon1, lat1, lon2, lat2) -> float:
    R = 6371.0
    p1 = math.radians(lat1); p2 = math.radians(lat2)
    dlat = math.radians(lat2 - lat1); dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlon / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def count_within(lon: float, lat: float, bucket: str, radius_km: float = 1.0) -> int:
    return sum(1 for p in _get().get(bucket, [])
               if _hav(lon, lat, p[0], p[1]) <= radius_km)


def coverage_within(lon: float, lat: float, bucket: str, radius_km: float = 1.0) -> float:
    """Approx fraction of the ``radius_km`` circle covered by this land-use.

    Sums the AREA (km²) of every parcel whose CENTROID is within the circle, divided
    by the circle's area. Approximation (a whole parcel counts if its centre is inside),
    so it can exceed 1.0 — that's fine: the Location score uses it relatively (comp vs
    subject), not as a literal capped percentage. Needs the area-enriched cache; if the
    cache has no areas it returns 0.0 (callers fall back gracefully)."""
    circle = math.pi * radius_km * radius_km
    if not circle:
        return 0.0
    total = 0.0
    for p in _get().get(bucket, []):
        if len(p) > 2 and _hav(lon, lat, p[0], p[1]) <= radius_km:
            total += p[2]
    return total / circle


def nearest_km(lon: float, lat: float, bucket: str) -> float:
    pts = _get().get(bucket, [])
    if not pts:
        return 99.0
    return min(_hav(lon, lat, p[0], p[1]) for p in pts)


# ── Land-use AT a point (point-in-polygon over the raw parcels) ───────────────
# Used to fill in a comp's land zoning when the source PDF didn't provide it.
# Local + token-free (works offline on Laptop B); needs the GeoJSON present.
_polys = None  # [(min_lon, min_lat, max_lon, max_lat, LU_DESC, ring), ...]


def _load_polys() -> list:
    """Lazily parse the GeoJSON into (bbox, LU_DESC, exterior-ring) tuples.

    Parsed once per process (the 189 MB file takes a few seconds); only invoked
    when a comp is missing its zoning, so runs that don't need it pay nothing.
    """
    global _polys
    if _polys is not None:
        return _polys
    _polys = []
    if not _GEOJSON.exists():
        return _polys
    data = json.loads(_GEOJSON.read_text(encoding="utf-8"))
    for f in data.get("features", []):
        lu = str(f.get("properties", {}).get("LU_DESC", "")).strip()
        if not lu:
            continue
        geom = f.get("geometry", {})
        t, c = geom.get("type"), geom.get("coordinates")
        rings = []
        if t == "Polygon" and c:
            rings = [c[0]]
        elif t == "MultiPolygon" and c:
            rings = [poly[0] for poly in c if poly]
        for ring in rings:
            if len(ring) < 3:
                continue
            xs = [p[0] for p in ring]
            ys = [p[1] for p in ring]
            _polys.append((min(xs), min(ys), max(xs), max(ys), lu, ring))
    return _polys


def _in_ring(x: float, y: float, ring: list) -> bool:
    """Ray-casting point-in-polygon test (holes ignored — fine for land use)."""
    inside = False
    n = len(ring)
    j = n - 1
    for i in range(n):
        xi, yi = ring[i][0], ring[i][1]
        xj, yj = ring[j][0], ring[j][1]
        if (yi > y) != (yj > y):
            if x < (xj - xi) * (y - yi) / (yj - yi) + xi:
                inside = not inside
        j = i
    return inside


def land_use_at(lon: float, lat: float) -> str:
    """Return the URA LU_DESC of the parcel containing (lon, lat), or '' if none."""
    if lon is None or lat is None:
        return ""
    for (minx, miny, maxx, maxy, lu, ring) in _load_polys():
        if minx <= lon <= maxx and miny <= lat <= maxy and _in_ring(lon, lat, ring):
            return lu
    return ""
