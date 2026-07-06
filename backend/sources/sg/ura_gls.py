"""
backend/sources/sg/ura_gls.py
=============================
URA **Government Land Sales (GLS)** — LAND comps from the LIVE data.gov.sg dataset that
powers https://www.ura.gov.sg/land-sales/current-ura-gls-sites/ .

This is the same free, keyless **data.gov.sg `datastore_search`** API the URA site itself
uses, so the data is always up to date (not a static snapshot). No URA Access Key needed.

The only thing this connector needs is the dataset's **resource id** (looks like
``d_xxxxxxxxxxxxxxxx``). URA stores it in their CMS, so it can't be auto-scraped — set it
once in Shared Settings as ``gls_resource_id`` (or pass ``gls_resource_id`` in the search
config). Without it the connector prints guidance and fails soft.

Because the dataset's exact column names aren't known ahead of time, the connector
**prints the live schema on first run** and **fuzzy-maps** columns → the LAND record shape
(site name, use/zoning, tenure, site area, GFA, tender price, $psm PPR, tenderer, date).
Once we see the real field names we can pin the mapping exactly.
"""
from __future__ import annotations

import json
import re
import ssl
import urllib.parse
import urllib.request

from ..base import SourceConnector, num
from ..registry import register
from .ura_api import _ROOT   # project-root path helper

_DGS = "https://data.gov.sg/api/action/datastore_search"
_UA = "Mozilla/5.0 (compatible; pgim-comps/1.0)"
_SQM_TO_SF = 10.7639

try:
    _CTX = ssl.create_default_context()
    _CTX.check_hostname = False
    _CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _CTX = None


def _get_resource_id(params: dict) -> str:
    rid = (params.get("gls_resource_id") or "").strip()
    if rid:
        return rid
    try:
        p = _ROOT / "configs" / "shared_settings.json"
        if p.exists():
            rid = (json.loads(p.read_text(encoding="utf-8")) or {}).get("gls_resource_id", "")
    except Exception:
        rid = ""
    return (rid or "").strip()


def _dgs(resource_id: str, limit: int = 500) -> dict:
    url = f"{_DGS}?{urllib.parse.urlencode({'resource_id': resource_id, 'limit': limit})}"
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    kw = {"timeout": 45}
    if _CTX is not None:
        kw["context"] = _CTX
    return json.loads(urllib.request.urlopen(req, **kw).read().decode("utf-8", "replace"))


# Fuzzy column matching: target field → keyword list (first field whose name contains
# any keyword wins). Kept broad because the dataset column names aren't documented.
_MATCH = {
    "site_name":     ["site", "land parcel", "parcel", "project", "development"],
    "address":       ["location", "address", "road", "street", "locality"],
    "launch_date":   ["date", "launch", "closing", "tender clos", "award", "gazette"],
    "land_zoning":   ["use", "zoning", "permissible", "allowable"],
    "tenure":        ["lease", "tenure"],
    "site_area":     ["site area", "land area", "area of site"],
    "max_gfa":       ["gfa", "gross floor", "max. permissible", "maximum permissible"],
    "price_total":   ["tender price", "sale price", "successful tender", "awarded price",
                      "price ($)", "purchase price", "tendered"],
    "price_psm":     ["psm", "per sq m", "$ per", "ppr", "per plot ratio", "psf"],
    "tenderer":      ["tenderer", "developer", "successful", "awarded to", "purchaser",
                      "bidder", "company"],
}


def _pick(fields: list, keys: list) -> str:
    """Return the dataset field id whose (lowercased) name matches any keyword, else ''."""
    for kw in keys:
        for f in fields:
            if kw in f.lower():
                return f
    return ""


def _area_sf(raw, field_name: str):
    v = num(raw)
    if v is None:
        return None
    fn = field_name.lower()
    if any(k in fn for k in ("sqm", "sq m", "sq.m", "m2", "m²", "square met")):
        return round(v * _SQM_TO_SF)
    return round(v)


class URAGLSConnector(SourceConnector):
    name = "ura_gls"
    market = "sg"
    comp_types = {"land"}
    label = "URA GLS (data.gov.sg, live)"

    def fetch(self, subject_cfg: dict, params: dict) -> tuple:
        rid = _get_resource_id(params)
        if not rid:
            print("    [ura_gls] no dataset id — set 'gls_resource_id' in Shared Settings "
                  "(the d_… id of the URA GLS dataset on data.gov.sg). Skipping.")
            return [], []
        try:
            limit = int(params.get("ura_max_rows", 500) or 500)
            resp = _dgs(rid, limit)
        except Exception as e:
            print(f"    [ura_gls] datastore_search failed: {e}")
            return [], []
        result = resp.get("result", {}) if isinstance(resp, dict) else {}
        recs = result.get("records") or []
        field_ids = [f.get("id") for f in result.get("fields", []) if f.get("id")]
        if not recs:
            print(f"    [ura_gls] dataset {rid[:12]}… returned 0 rows "
                  f"(check the resource id). Fields: {field_ids}")
            return [], []
        # Surface the live schema once so the mapping can be pinned precisely.
        print(f"    [ura_gls] {len(recs)} rows | columns: {field_ids}")

        col = {t: _pick(field_ids, kws) for t, kws in _MATCH.items()}
        out = []
        for r in recs:
            name = str(r.get(col["site_name"], "") or "").strip()
            if not name:
                continue
            total = num(r.get(col["price_total"], "")) if col["price_total"] else None
            price_m = round(total / 1_000_000, 3) if (total and total > 100_000) else num(
                r.get(col["price_total"], "")) if col["price_total"] else None
            out.append({
                "site_name":     name,
                "property_name": name,
                "address":       (str(r.get(col["address"], "") or "").strip()
                                  + (", Singapore" if col["address"] else "")),
                "launch_date":   str(r.get(col["launch_date"], "") or "").strip(),
                "land_zoning":   str(r.get(col["land_zoning"], "") or "").strip(),
                "tenure":        str(r.get(col["tenure"], "") or "").strip(),
                "site_area_sf":  _area_sf(r.get(col["site_area"], ""), col["site_area"])
                                 if col["site_area"] else None,
                "max_gfa_sf":    _area_sf(r.get(col["max_gfa"], ""), col["max_gfa"])
                                 if col["max_gfa"] else None,
                "price_sgd_m":   price_m,
                "price_psf_ppr": num(r.get(col["price_psm"], "")) if col["price_psm"] else None,
                "sale_type":     "GLS Tender",
                "asset_type":    str(r.get(col["land_zoning"], "") or "GLS").strip(),
                "buyer":         str(r.get(col["tenderer"], "") or "").strip(),
                "seller":        "State (URA GLS)",
                "country":       "Singapore",
            })
        print(f"    [ura_gls] mapped {len(out)} GLS site(s)")
        return out, [{"title": self.label,
                      "url": "https://www.ura.gov.sg/land-sales/current-ura-gls-sites/"}]


register(URAGLSConnector())
