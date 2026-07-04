"""
backend/sources/base.py
========================
Market-agnostic source-connector abstraction for Online Search.

A connector fetches raw comparable records for a subject + params and returns them
in the SAME shape the existing ``search_and_extract*`` functions return, so they flow
through the shared pipeline (dedup → geocode → classify → Excel/map) unchanged.

Connectors do NOT geocode, dedup across sources, or classify — that is the shared
pipeline's job. They only produce raw records + citations, and MUST FAIL SOFT (return
``([], [])`` on a missing key / HTTP error) so the pipeline continues on other sources.

Record shapes (per comp_type) the connectors must emit — same as the extractors today:
  sales: property_name, address, sale_date, price_sgd_m, gfa_sf, remaining_yrs,
         cap_rate_pct, stake_pct, sale_type, asset_type, land_zoning, buyer, seller, country
  land : site_name, address, launch_date, land_zoning, tenure, site_area_sf,
         max_gfa_sf, price_sgd_m, price_psf_ppr, sale_type, asset_type, country
  rent : property_name, address, lease_date, nla_sf, asking_rent, eff_rent,
         lease_term_yrs, rent_free_mths, lease_type, currency
"""
from __future__ import annotations

import re


class SourceConnector:
    """Base class for a grounded data source. Subclasses set the class attrs and
    implement ``fetch``."""

    name: str = ""            # unique key, e.g. "ura_gls"
    market: str = ""          # "sg" / "kr" / … ; "" = market-agnostic
    comp_types: set = set()   # subset of {"sales", "land", "rent"}
    label: str = ""           # human label for the Sources sheet (defaults to name)

    def fetch(self, subject_cfg: dict, params: dict) -> tuple:
        """Return ``(records, sources)``.

        records : list[dict] in the comp_type's raw record shape.
        sources : list[{"title","url"}] (optional; records may also carry per-record
                  "sources").
        Implementations MUST fail soft — return ``([], [])`` on any error / missing key.
        """
        raise NotImplementedError


# ── Shared coercers (market-agnostic; reused by connectors) ────────────────────
def num(v):
    """Best-effort numeric parse → float or None (strips $ , spaces, units)."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    if s in ("", "-", ".", "--"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def clean_str(v) -> str:
    return "" if v is None else str(v).strip()


def sqm_to_sf(v) -> float | None:
    """Convert square metres → square feet (source areas are often sqm)."""
    n = num(v)
    return round(n * 10.7639, 1) if n is not None else None
