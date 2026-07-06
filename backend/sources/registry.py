"""
backend/sources/registry.py
===========================
Registry for GROUNDED source connectors. Connectors self-register by ``name``; the
search scripts pick them by market + comp_type + user-enabled names.

Note: the existing OpenAI **web search** is query/level-driven and stays handled inline
in each ``search_online_*.py`` (sentinel name ``"web_search"``). This registry manages
only the one-shot grounded connectors (URA PMI, URA GLS, broker PDFs, …).

Adding a new market = create ``backend/sources/<market>/`` with connectors that call
``register(...)`` at import; then extend ``_ensure_loaded``. Nothing else changes.
"""
from __future__ import annotations

WEB_SEARCH = "web_search"

_REGISTRY: dict = {}   # name -> SourceConnector instance
_LOADED: set = set()   # markets whose connector modules have been imported


def register(conn):
    """Called by each connector module at import time."""
    _REGISTRY[conn.name] = conn
    return conn


def _ensure_loaded(market: str) -> None:
    """Import a market's connector modules once (they self-register on import)."""
    if market in _LOADED:
        return
    _LOADED.add(market)
    if market == "sg":
        for _mod in ("ura_pmi", "ura_pmi_rental", "broker_reports_sg"):   # add "ura_gls", … as built
            try:
                __import__(f"sources.sg.{_mod}")   # side-effect: register()
            except Exception as e:                 # pragma: no cover — fail soft
                print(f"  [sources] sg.{_mod} not loaded: {e}")
    # future markets:
    # elif market == "kr":
    #     from .kr import ...


def available(market: str, comp_type: str) -> list:
    """Names of selectable sources for this market+comp_type (web_search first)."""
    _ensure_loaded(market)
    names = [WEB_SEARCH]
    for name, c in _REGISTRY.items():
        if name == WEB_SEARCH:
            continue
        if c.market in ("", market) and comp_type in c.comp_types:
            names.append(name)
    return names


def get_grounded(market: str, comp_type: str, enabled: list, params: dict) -> list:
    """Return the enabled GROUNDED connector instances valid for market+comp_type.
    (Excludes the ``web_search`` sentinel, which the caller handles inline.)"""
    _ensure_loaded(market)
    out = []
    for name in (enabled or []):
        if name == WEB_SEARCH:
            continue
        c = _REGISTRY.get(name)
        if c is not None and c.market in ("", market) and comp_type in c.comp_types:
            out.append(c)
    return out
