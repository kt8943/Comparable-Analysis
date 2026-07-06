"""
backend/sources/sg/ura_gls.py
=============================
URA **Government Land Sales (GLS)** — LAND comps from the LIVE **"URA Sale Sites
(GEOJSON)"** dataset on data.gov.sg (the same public source behind URA's GLS pages).

FREE / KEYLESS and always up to date. Each feature is an **awarded** GLS tender with the
successful tenderer + tender price + site area + plot ratio + lease + award date, and a
**polygon geometry** (WGS84) — so we take the centroid for the map pin directly, no
address geocoding needed.

Fields used (pinned to the dataset schema):
    LOCATION (site) · DEVT_CODE / DEVT_ALLOW (use) · DATE_AWARD/CLOSG/LNCH (yyyymmdd) ·
    LEASE_YR · SA_SQM (site area) · GPR (plot ratio) · GFA · SUCCESS_TP (tender price $) ·
    SUCCESS_TR (tenderer) · NO_OF_BIDS · geometry (Polygon/MultiPolygon).

Filtered to the subject's sector (office/retail/hotel/industrial/…); recency + distance
are applied downstream by the land pipeline. Fails soft (returns [], []) on any error.
The dataset id can be overridden via ``gls_resource_id`` in the search config.
"""
from __future__ import annotations

import json
import ssl
import urllib.request

from ..base import SourceConnector, num
from ..registry import register

_DEFAULT_DATASET = "d_0e2b42f98535686282031a42c9c7b05a"   # URA Sale Sites (GEOJSON)
_POLL = "https://api-open.data.gov.sg/v1/public/api/datasets/{id}/poll-download"
_UA = "Mozilla/5.0 (compatible; pgim-comps/1.0)"
_SQM_TO_SF = 10.7639
_MONTHS = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

try:
    _CTX = ssl.create_default_context()
    _CTX.check_hostname = False
    _CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _CTX = None

# subject sector → GLS DEVT_CODE values (lowercased) that count as the same market.
_SECTOR_CODES = {
    "office":     {"office", "commercial", "commercial and residential", "white site",
                   "transitional office"},
    "retail":     {"commercial", "commercial and residential", "white site",
                   "residential with commercial at 1st sty"},
    "hotel":      {"hotel"},
    "hospitality": {"hotel"},
    "industrial": {"industrial", "warehousing", "industrial - white",
                   "heavy vehicle park"},
    "logistics":  {"industrial", "warehousing"},
    "residential": {"residential (landed)", "residential (non-landed)",
                    "residential with commercial at 1st sty", "commercial and residential"},
    "mixed":      {"commercial and residential", "white site",
                   "residential with commercial at 1st sty"},
}


def _sector(asset_class: str) -> str:
    a = (asset_class or "").lower()
    if any(k in a for k in ("logistic", "warehouse")):       return "logistics"
    if any(k in a for k in ("industrial", "factory")):        return "industrial"
    if any(k in a for k in ("hotel", "hospitality")):         return "hotel"
    if any(k in a for k in ("retail", "mall", "shop")):       return "retail"
    if "office" in a:                                         return "office"
    if "resid" in a:                                          return "residential"
    if "mixed" in a:                                          return "mixed"
    return "office"


def _get(url: str, timeout: int = 60) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    kw = {"timeout": timeout}
    if _CTX is not None:
        kw["context"] = _CTX
    return urllib.request.urlopen(req, **kw).read()


def _centroid(geom: dict):
    """Average of the (first) exterior ring's vertices → (lon, lat)."""
    try:
        c = geom.get("coordinates")
        while isinstance(c, list) and c and isinstance(c[0], list) and \
                isinstance(c[0][0], list):
            c = c[0]                      # descend to a ring of [lon,lat] pairs
        pts = [(p[0], p[1]) for p in c if isinstance(p, list) and len(p) >= 2]
        if not pts:
            return None, None
        n = len(pts)
        return round(sum(p[0] for p in pts) / n, 6), round(sum(p[1] for p in pts) / n, 6)
    except Exception:
        return None, None


def _award_date(props: dict) -> str:
    for k in ("DATE_AWARD", "DATE_CLOSG", "DATE_LNCH"):
        v = str(props.get(k) or "").strip()
        if len(v) == 8 and v.isdigit():
            y, m = int(v[:4]), int(v[4:6])
            return f"{_MONTHS[m]} {y}" if 1 <= m <= 12 else str(y)
    return ""


class URAGLSConnector(SourceConnector):
    name = "ura_gls"
    market = "sg"
    comp_types = {"land"}
    label = "URA GLS (data.gov.sg, live)"

    def fetch(self, subject_cfg: dict, params: dict) -> tuple:
        did = (params.get("gls_resource_id") or "").strip() or _DEFAULT_DATASET
        try:
            poll = json.loads(_get(_POLL.format(id=did)).decode("utf-8", "replace"))
            url = poll.get("data", {}).get("url")
            if not url:
                print(f"    [ura_gls] no download url for {did[:12]}…: {poll}")
                return [], []
            gj = json.loads(_get(url).decode("utf-8", "replace"))
        except Exception as e:
            print(f"    [ura_gls] fetch failed: {e}")
            return [], []

        feats = gj.get("features") or []
        sector = _sector(subject_cfg.get("asset_class", ""))
        allowed = _SECTOR_CODES.get(sector)     # None → keep all sectors
        out = []
        for f in feats:
            p = f.get("properties", {}) or {}
            name = str(p.get("LOCATION") or "").strip()
            tp = num(p.get("SUCCESS_TP"))
            if not name or not tp:
                continue
            code = str(p.get("DEVT_CODE") or "").strip()
            if allowed is not None and code.lower() not in allowed:
                continue
            lon, lat = _centroid(f.get("geometry") or {})
            sa_sf  = round(num(p.get("SA_SQM")) * _SQM_TO_SF) if num(p.get("SA_SQM")) else None
            gfa    = num(p.get("GFA"))
            gpr    = num(p.get("GPR"))
            gfa_sf = (round(gfa * _SQM_TO_SF) if gfa else
                      (round(num(p.get("SA_SQM")) * gpr * _SQM_TO_SF)
                       if (num(p.get("SA_SQM")) and gpr) else None))
            lease  = num(p.get("LEASE_YR"))
            tenure = ("Freehold" if (lease and lease >= 999)
                      else (f"{int(lease)} yrs" if lease else ""))
            rec = {
                "site_name":     name,
                "property_name": name,
                "address":       f"{name}, Singapore",
                "launch_date":   _award_date(p),
                "land_zoning":   code or str(p.get("DEVT_ALLOW") or ""),
                "tenure":        tenure,
                "site_area_sf":  sa_sf,
                "max_gfa_sf":    gfa_sf,
                "price_sgd_m":   round(tp / 1_000_000, 3),
                "price_psf_ppr": round(tp / gfa_sf, 2) if gfa_sf else None,
                "sale_type":     "GLS Tender",
                "asset_type":    code or "GLS",
                "no_of_bids":    int(num(p.get("NO_OF_BIDS")) or 0) or None,
                "buyer":         str(p.get("SUCCESS_TR") or "").strip(),
                "seller":        "State (URA GLS)",
                "country":       "Singapore",
            }
            if lon is not None:
                rec["lon"], rec["lat"] = lon, lat
            out.append(rec)
        print(f"    [ura_gls] {len(feats)} awarded sites → {len(out)} matched "
              f"sector '{sector}'")
        return out, [{"title": self.label,
                      "url": "https://www.ura.gov.sg/land-sales/current-ura-gls-sites/"}]


register(URAGLSConnector())
