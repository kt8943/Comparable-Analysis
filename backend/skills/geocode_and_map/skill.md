---
name: geocode_and_map
description: Geocode comp records via Mapbox, sort by haversine distance from subject, assign map markers, and optionally render a static map PNG
type: atomic
requires:
  config_keys:
    - subject_property.lon
    - subject_property.lat
    - subject_property.country_code
    - subject_property.country_name
    - mapbox_token     # optional — geocoding and map are skipped if absent
  skills: []
allowed_tools:
  - tools.calculations.haversine_km
  - tools.geo_utils.write_geo_sidecar
---

## When to use

Called in Stage 4 of all three comp pipelines (`analyse_sales_comps`, `analyse_rent_comps`, `analyse_land_comps`) after classification. Also called in the online comp pipelines after web search results are parsed. Requires a Mapbox API token to geocode; map rendering is also skipped without a token.

## Instructions

1. For each comp record, call `generate_*_comps_map.geocode_with_fallbacks([address_query], token, country_code)` → returns `(lon, lat)`
   - Tries address first; falls back to property name if address geocoding fails
   - Records already having `lon`/`lat` (from a `--from-records` re-run) are skipped — distance is recalculated but no API call is made
   - If geocoding fails, set `lon=None`, `lat=None`, `distance_km=9999.0`
2. Compute `distance_km = tools.calculations.haversine_km(lon, lat, s_lon, s_lat)` for each geocoded record
3. Sort records by `distance_km` ascending
4. Assign sequential `map_marker` strings (`"1"`, `"2"`, `"3"`, …)
5. Call `tools.geo_utils.write_geo_sidecar(out_geo, s_lon, s_lat, comps, mb_cfg)` to write the `_geo.json` sidecar
6. If `--map` flag is set and `mapbox_token` is present: call `render_map(subject_lonlat, geo_records, token, output_path)` → saves PNG

## Output format

| Output | Description |
| --- | --- |
| Records list (in-place mutation) | Each record gains `lon`, `lat`, `distance_km`, `map_marker` |
| `*_geo.json` | Geocoded records sidecar — drives the interactive pydeck map and PNG regeneration |
| `*_map.png` | Mapbox static map image (only when `--map` flag is passed) |

## Examples

```bash
python3 scan_input_sales_comps.py --config configs/deal_config_88_Cecil.json --map
```

```python
# Internal call sequence in run()
classified = _geocode_comps(classified, mapbox_tok, country_code, country_name, s_lon, s_lat)
classified.sort(key=lambda r: r.get("distance_km", 9999))
for i, r in enumerate(classified, 1):
    r["map_marker"] = str(i)
write_geo_sidecar(out_geo, s_lon, s_lat, geo_comps, mb_cfg)
if generate_map:
    render_map((s_lon, s_lat), classified, mapbox_tok, map_output_path)
```

## Notes

- `geocode_with_fallbacks` is defined in each `generate_*_comps_map.py` — there is one per comp type but they share the same Mapbox Geocoding API call pattern
- Land comps only geocode records where `address` was resolved by `classify_land_comps` — records with blank address are skipped to avoid wrong coordinates from site-name queries
- `_geo.json` is written unconditionally (no `--map` flag needed) so the interactive dashboard map always works even without a static PNG
- Mapbox free tier: 100,000 geocoding requests/month; each comp is 1–2 API calls
