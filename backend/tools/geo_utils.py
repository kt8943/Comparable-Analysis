"""
tools/geo_utils.py
==================
Geocoding sidecar utilities shared across all comp pipelines.

Public API
----------
write_geo_sidecar(path, s_lon, s_lat, comps, mb_cfg) -> None
    Write a _geo.json sidecar file used by the interactive pydeck map.
"""

import json
from pathlib import Path


def write_geo_sidecar(
    path: str,
    s_lon: float,
    s_lat: float,
    comps: list,
    mb_cfg: dict,
) -> None:
    """Write the _geo.json sidecar that drives the interactive pydeck map.

    comps is a list of dicts, each with:
        map_marker  str
        property    str   display name
        address     str
        lon         float | None
        lat         float | None
    Any key named 'token' in mb_cfg is excluded from the sidecar.
    """
    geo_data = {
        "subject": {"lon": s_lon, "lat": s_lat},
        "mapbox":  {k: v for k, v in mb_cfg.items() if k != "token"},
        "comps": [
            {
                "map_marker": c["map_marker"],
                "property":   str(c.get("property") or ""),
                "address":    str(c.get("address")  or ""),
                "lon":        c.get("lon"),
                "lat":        c.get("lat"),
                "hidden":     False,
                "color":      c.get("color"),   # per-pin override; None = default
            }
            for c in comps
        ],
    }
    Path(path).write_text(json.dumps(geo_data, indent=2), encoding="utf-8")
