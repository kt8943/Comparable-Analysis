#!/usr/bin/env python3
"""
house_rules.py
==============
House comp-search rules — the single source of truth, applied to every deal
(existing and new, local and cloud) by search_online_sales/rent/land_comps.py.

These are policy, not per-deal settings, so they live here rather than being copied
into each deal config. Change a number below and every deal picks it up on the next
run: no config edits, no migration, and local and cloud cannot drift apart.

Precedence
----------
    HOUSE_RULES  →  BY_ASSET_CLASS  →  the deal's own search block

A deal that genuinely needs different numbers can still set the key in its own
online_search / rent_search / land_search block, and that wins. Config always beats
code here — nothing in this module overrides a value a deal explicitly states.

Two settings that interact
--------------------------
years_back shapes the QUERY (which years the web search is asked for).
recency_months filters the RESULTS (which comps are kept). They are independent, so
a years_back_max reaching past recency_months only buys rows that are then dropped.
warn_window_vs_recency reports that conflict rather than silently "fixing" it.
"""

# ── House rules ──────────────────────────────────────────────────────────────
# Location ladder: proximity → city → country. The country tier has no radius.
# The city tier is a RADIUS, not a municipal boundary — the geocoder returns lon/lat
# with no locality field to test containment against. Widen it for a larger metro.
HOUSE_RULES = {
    "proximity_km":    3.0,   # tier 1
    "city_km":         25.0,  # tier 2
    "min_results":     3,     # below this, escalate to the next tier
    "max_results":     15,    # hard cap on comps per category
    "max_queries":     5,     # web queries per run; 1 query = 1 search + 1 extract call
    "years_back":      2,     # initial query window
    "years_back_max":  5,     # never ask beyond what recency keeps
    "years_back_step": 2,
    "max_level":       3,     # 1=proximity, 2=+city, 3=+country
}

# Recency differs by evidence type: rental evidence dates faster than capital
# evidence, so a lease from the last cycle says little about today's achievable rent.
RECENCY_MONTHS = {
    "sales": 60,   # 5 years
    "land":  60,   # 5 years
    "rent":  36,   # 3 years
}

# Radius overrides by asset class — logistics and industrial trade over a wider
# catchment than office. That is a property of the asset class, not of one deal.
BY_ASSET_CLASS = {
    "logistics":  {"proximity_km": 5.0, "city_km": 50.0},
    "industrial": {"proximity_km": 5.0, "city_km": 50.0},
}


def search_rules(comp_type: str, deal_block: dict | None = None,
                 asset_class: str = "") -> dict:
    """Effective search settings for one comp type ("sales" | "rent" | "land").

    Returns a plain dict the caller reads with .get(key, fallback) as before.
    """
    out = dict(HOUSE_RULES)
    out["recency_months"] = RECENCY_MONTHS.get(comp_type, 60)

    ac = (asset_class or "").strip().lower()
    out.update(BY_ASSET_CLASS.get(ac, {}))

    out.update(deal_block or {})   # a deal's explicit setting always wins
    return out


def warn_window_vs_recency(years_back_max, recency_months) -> str:
    """Message when the query window reaches past what recency will keep, else "".

    Reports the conflict; does not change either number. A run that quietly ignores
    its own settings is worse than one that says what it is about to waste.
    """
    try:
        ybm     = float(years_back_max)
        rec_yrs = float(recency_months) / 12.0
    except (TypeError, ValueError):
        return ""
    if ybm > rec_yrs:
        return (f"  ⚠  years_back_max={ybm:g}yr reaches past recency_months="
                f"{recency_months:g}mo (={rec_yrs:g}yr). Queries will ask for sales "
                f"the recency filter then drops. Lower years_back_max to ≤ {rec_yrs:g} "
                f"in backend/tools/house_rules.py.")
    return ""
