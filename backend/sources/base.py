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
from datetime import datetime

_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def months_ago(date_str: str):
    """Approx. months between today and a sale date given in varied formats:
    'Jun-26', 'Jun 2025', 'Jun-2025', 'Q2 2025', '2025', '01 Jun 2026'.
    Returns int months (>=0) or None when unparseable (caller should KEEP on None)."""
    s = (date_str or "").strip().lower()
    if not s:
        return None
    now = datetime.now()
    yr = mo = None
    mq = re.search(r"q([1-4])\D*(\d{4})", s)              # Q2 2025
    if mq:
        yr, mo = int(mq.group(2)), (int(mq.group(1)) - 1) * 3 + 2
    if yr is None:                                        # Jun-26 / Jun 2025
        mm = re.search(r"([a-z]{3})[\s\-/]*(\d{2,4})", s)
        if mm and mm.group(1) in _MONTHS:
            mo = _MONTHS[mm.group(1)]
            y = int(mm.group(2))
            yr = 2000 + y if y < 100 else y
    if yr is None:                                        # bare year
        my = re.search(r"((?:19|20)\d{2})", s)
        if my:
            yr, mo = int(my.group(1)), 6                  # mid-year assumption
    if yr is None:
        return None
    return max(0, (now.year - yr) * 12 + (now.month - (mo or 6)))


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
