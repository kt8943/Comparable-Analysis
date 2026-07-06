#!/usr/bin/env python3
"""
new_deal.py
===========
Interactive wizard that creates a new deal config JSON in seconds.

You provide the essentials (name, address, asset class, GFA, deal numbers).
The LLM auto-fills everything derived: country, currency, GFA unit, location
descriptor, submarket keywords, broader market query, land zoning, and the
asset search keyword used in online comp searches.

Usage
-----
    python3 new_deal.py
"""

import json
import os
import re
import sys
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

# ── Windows UTF-8 fix ─────────────────────────────────────────────────────────
for _stream in (sys.stdout, sys.stderr):
    try:
        if getattr(_stream, "encoding", "utf-8").lower().replace("-", "") != "utf8":
            _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError):
        pass

# ─────────────────────────────────────────────────────────────────────────────
# DEFAULTS  (read from any existing config so shared settings stay in sync)
# ─────────────────────────────────────────────────────────────────────────────

_SEARCH_DEFAULTS = {
    "proximity_km":    1.0,
    "submarket_km":    5.0,
    "min_results":     3,
    "max_results":     10,
    "years_back":      2,
    "years_back_max":  8,
    "years_back_step": 2,
    "max_level":       3,
}

# Typical proximity / submarket radii by asset class
_RADIUS_BY_CLASS = {
    "office":      {"proximity_km": 1.0, "submarket_km": 3.0},
    "retail":      {"proximity_km": 1.0, "submarket_km": 3.0},
    "logistics":   {"proximity_km": 5.0, "submarket_km": 10.0},
    "industrial":  {"proximity_km": 5.0, "submarket_km": 10.0},
    "mixed-use":   {"proximity_km": 1.0, "submarket_km": 5.0},
}


def _load_shared_settings() -> dict:
    """Read mapbox token + ollama config from the first existing deal config (any deal)."""
    configs_dir = Path("configs")
    candidates = sorted(configs_dir.glob("deal_config*.json")) if configs_dir.exists() else []
    for p in candidates:
        try:
            cfg = json.loads(p.read_text(encoding="utf-8"))
            return {
                "mapbox_token": cfg.get("mapbox", {}).get("token", ""),
                "ollama_base":  cfg.get("llm", {}).get("ollama", {}).get("base_url",
                                                                          "http://localhost:11434"),
                "ollama_model": cfg.get("llm", {}).get("ollama", {}).get("model", "qwen2.5:3b"),
            }
        except Exception:
            continue
    return {
        "mapbox_token": "",
        "ollama_base":  "http://localhost:11434",
        "ollama_model": "qwen2.5:3b",
    }


# ─────────────────────────────────────────────────────────────────────────────
# LLM HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _ollama_post(base_url: str, model: str, messages: list, timeout: int = 60) -> str:
    payload = json.dumps({
        "model": model, "messages": messages,
        "stream": False, "format": "json", "options": {"temperature": 0},
    }).encode()
    req = urllib.request.Request(
        f"{base_url}/api/chat", data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read())["message"]["content"]


def _openai_post(model: str, api_key: str, messages: list) -> str:
    """Call OpenAI chat completions with JSON response format enforced."""
    try:
        import openai as _openai
    except ImportError:
        raise ImportError("openai package required.  pip install openai")
    client = _openai.OpenAI(api_key=api_key)
    resp = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=0,
        response_format={"type": "json_object"},
    )
    return resp.choices[0].message.content.strip()


def _llm_post(messages: list, llm_cfg: dict, openai_key: str = "",
              timeout: int = 60) -> str:
    """Route to Ollama or OpenAI based on llm_cfg provider."""
    provider = llm_cfg.get("provider", "ollama")
    if provider == "openai":
        key   = openai_key or llm_cfg.get("openai_api_key", "")
        model = llm_cfg.get("openai_model", "gpt-4o")
        if not key:
            raise ValueError("OpenAI model selected but no API key found.")
        return _openai_post(model, key, messages)
    else:
        ollama = llm_cfg.get("ollama", {})
        return _ollama_post(
            ollama.get("base_url", "http://localhost:11434"),
            ollama.get("model", "qwen2.5:3b"),
            messages,
            timeout=timeout,
        )


def _mapbox_geocode(address: str, token: str) -> dict:
    """
    Forward geocode an address using the Mapbox Geocoding API.
    Returns a dict with whatever place hierarchy Mapbox finds:
      neighborhood, locality, district, place, region, country
    Uses the same Mapbox token already stored in the deal config.
    """
    if not token:
        return {}
    try:
        encoded = urllib.parse.quote(address)
        url = (
            f"https://api.mapbox.com/geocoding/v5/mapbox.places/{encoded}.json"
            f"?access_token={token}&limit=1"
        )
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())

        if not data.get("features"):
            return {}

        feature = data["features"][0]
        result  = {}

        # Walk the context array — each entry is a level of the place hierarchy
        for ctx in feature.get("context", []):
            ctx_id = ctx.get("id", "")
            text   = ctx.get("text", "").strip()
            if not text:
                continue
            if ctx_id.startswith("neighborhood"):
                result["neighborhood"] = text
            elif ctx_id.startswith("locality"):
                result["locality"] = text
            elif ctx_id.startswith("district"):
                result["district"] = text
            elif ctx_id.startswith("place"):
                result["place"] = text
            elif ctx_id.startswith("region"):
                result["region"] = text
            elif ctx_id.startswith("country"):
                result["country"] = text

        # The feature itself may be a neighborhood-level result
        if "neighborhood" in feature.get("place_type", []):
            result.setdefault("neighborhood", feature.get("text", "").strip())

        return result
    except Exception as e:
        print(f"  [mapbox geocode] failed: {e}")
        return {}


# Deterministic market lookups for the rule-based (no-LLM) config generator.
# Keyed by a lowercase substring found in the address; first match wins.
_COUNTRY_RULES = {
    "singapore":      dict(country_name="Singapore",      country_code="sg", currency="SGD", currency_symbol="S$",  gfa_unit="sf"),
    "south korea":    dict(country_name="South Korea",    country_code="kr", currency="KRW", currency_symbol="₩",   gfa_unit="sqm"),
    "korea":          dict(country_name="South Korea",    country_code="kr", currency="KRW", currency_symbol="₩",   gfa_unit="sqm"),
    "seoul":          dict(country_name="South Korea",    country_code="kr", currency="KRW", currency_symbol="₩",   gfa_unit="sqm"),
    "japan":          dict(country_name="Japan",          country_code="jp", currency="JPY", currency_symbol="¥",   gfa_unit="sqm"),
    "tokyo":          dict(country_name="Japan",          country_code="jp", currency="JPY", currency_symbol="¥",   gfa_unit="sqm"),
    "osaka":          dict(country_name="Japan",          country_code="jp", currency="JPY", currency_symbol="¥",   gfa_unit="sqm"),
    "hong kong":      dict(country_name="Hong Kong",      country_code="hk", currency="HKD", currency_symbol="HK$", gfa_unit="sf"),
    "shanghai":       dict(country_name="China",          country_code="cn", currency="CNY", currency_symbol="RMB", gfa_unit="sqm"),
    "beijing":        dict(country_name="China",          country_code="cn", currency="CNY", currency_symbol="RMB", gfa_unit="sqm"),
    "china":          dict(country_name="China",          country_code="cn", currency="CNY", currency_symbol="RMB", gfa_unit="sqm"),
    "sydney":         dict(country_name="Australia",      country_code="au", currency="AUD", currency_symbol="A$",  gfa_unit="sqm"),
    "melbourne":      dict(country_name="Australia",      country_code="au", currency="AUD", currency_symbol="A$",  gfa_unit="sqm"),
    "australia":      dict(country_name="Australia",      country_code="au", currency="AUD", currency_symbol="A$",  gfa_unit="sqm"),
    "kuala lumpur":   dict(country_name="Malaysia",       country_code="my", currency="MYR", currency_symbol="RM",  gfa_unit="sf"),
    "malaysia":       dict(country_name="Malaysia",       country_code="my", currency="MYR", currency_symbol="RM",  gfa_unit="sf"),
    "london":         dict(country_name="United Kingdom", country_code="gb", currency="GBP", currency_symbol="£",   gfa_unit="sf"),
    "united kingdom": dict(country_name="United Kingdom", country_code="gb", currency="GBP", currency_symbol="£",   gfa_unit="sf"),
    "new york":       dict(country_name="United States",  country_code="us", currency="USD", currency_symbol="US$", gfa_unit="sf"),
    "united states":  dict(country_name="United States",  country_code="us", currency="USD", currency_symbol="US$", gfa_unit="sf"),
    "usa":            dict(country_name="United States",  country_code="us", currency="USD", currency_symbol="US$", gfa_unit="sf"),
}
_DEFAULT_COUNTRY = dict(country_name="Singapore", country_code="sg",
                        currency="SGD", currency_symbol="S$", gfa_unit="sf")
# keyword-in-asset_class → (asset_search_keyword, land_zoning)
_ASSET_RULES = [
    ("logistic",    ("logistics warehouse",  "Logistics / Industrial")),
    ("warehouse",   ("logistics warehouse",  "Logistics / Industrial")),
    ("data",        ("data centre",          "Business Park / Industrial")),
    ("industrial",  ("industrial",           "Industrial")),
    ("office",      ("office building",      "Commercial")),
    ("retail",      ("retail mall",          "Commercial")),
    ("mall",        ("retail mall",          "Commercial")),
    ("shop",        ("retail",               "Commercial")),
    ("hotel",       ("hotel",                "Hospitality")),
    ("hospitality", ("hotel",                "Hospitality")),
    ("resid",       ("residential",          "Residential")),
    ("mixed",       ("mixed use development", "Commercial")),
]


def _derive_fields_rules(address: str, asset_class: str) -> dict:
    """Deterministic (no-LLM) counterpart of _derive_fields_with_llm.

    Infers the market/config fields from the address string + asset class using
    lookup tables — no model call, nothing leaves the machine. Mapbox geocoding
    still refines ``location`` / ``submarket_keywords`` back in derive_market_fields.
    """
    a  = (address or "").lower()
    ac = (asset_class or "").lower()
    country = next((v for k, v in _COUNTRY_RULES.items() if k in a), _DEFAULT_COUNTRY)
    kw, zoning = "commercial property", "Commercial"
    for k, (kwv, zv) in _ASSET_RULES:
        if k in ac:
            kw, zoning = kwv, zv
            break
    # submarket fallback: the part before the country/postcode is usually the city/area
    parts = [p.strip() for p in (address or "").split(",") if p.strip()]
    loc   = parts[-2] if len(parts) >= 2 else (parts[0] if parts else country["country_name"])
    return {
        **country,
        "land_zoning":          zoning,
        "location":             loc,
        "asset_search_keyword": kw,
        "submarket_keywords":   [loc] if loc else [country["country_name"]],
        "broader_market_query": f"{country['country_name']} {kw} investment sale",
    }


def _derive_fields_with_llm(address: str, asset_class: str,
                              llm_cfg: dict, openai_key: str = "") -> dict:
    """Ask the LLM to derive all config fields from address + asset class."""
    system = (
        "You are a real estate data assistant. Given a property address and asset class, "
        "return ONLY a JSON object with the exact keys listed. No commentary, no markdown.\n\n"
        "CRITICAL RULE: Every location name you produce — in 'location', 'submarket_keywords', "
        "and 'broader_market_query' — must be within the SAME city and country as the address. "
        "Never include names from other cities, regions, or countries.\n\n"
        "Keys to return:\n"
        "  country_name        — full English country name (e.g. 'Singapore', 'South Korea')\n"
        "  country_code        — ISO 3166-1 alpha-2 (e.g. 'sg', 'kr', 'jp', 'au')\n"
        "  currency            — 3-letter ISO code (e.g. 'SGD', 'USD', 'KRW', 'AUD')\n"
        "  currency_symbol     — display symbol (e.g. 'S$', 'US$', '₩', 'A$')\n"
        "  gfa_unit            — 'sf' for Singapore/US/UK/HK/AU; 'sqm' for Korea/Japan/Europe/SE Asia\n"
        "  land_zoning         — typical zoning for this asset class in this country "
                                "(e.g. 'Commercial', 'Logistics / Industrial', 'B1 Industrial')\n"
        "  location            — the specific district or submarket the street address sits in, "
                                "within that same city only. "
                                "Derive this from the actual street name and city in the address — "
                                "do not copy from any example. "
                                "Never name a district from a different city or country.\n"
        "  asset_search_keyword — 2-4 word phrase for web search, e.g. 'office building', "
                                 "'logistics warehouse', 'retail mall', 'industrial park'\n"
        "  submarket_keywords  — JSON array of 3-5 neighbourhood, district, or precinct names "
                                "that are geographically close to the address provided. "
                                "Derive these from the actual street and city in the address — "
                                "do not copy from any example. "
                                "Never include names from other cities or countries.\n"
        "  broader_market_query — one sentence web search scoped to the same city and country, "
                                 "e.g. 'Singapore office investment sale 2024'\n"
    )
    user = (
        f"Address: {address}\n"
        f"Asset class: {asset_class}\n\n"
        f"All location names must be within the same city and country as this address. "
        f"Return the JSON object."
    )
    raw = _llm_post(
        [{"role": "system", "content": system}, {"role": "user", "content": user}],
        llm_cfg, openai_key, timeout=300,   # 5 min — 8b+ models on CPU can be slow
    )
    raw = re.sub(r"^```[a-z]*\n?", "", raw.strip())
    raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw)


# ─────────────────────────────────────────────────────────────────────────────
# INPUT HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _ask(prompt: str, default=None, cast=str) -> object:
    """Prompt the user; return default on empty input."""
    hint = f" [{default}]" if default is not None else ""
    while True:
        raw = input(f"  {prompt}{hint}: ").strip()
        if raw == "" and default is not None:
            return default
        if raw == "" and default is None:
            print("    (required — please enter a value)")
            continue
        try:
            return cast(raw)
        except (ValueError, TypeError):
            print(f"    (invalid value, expected {cast.__name__})")


def _ask_optional(prompt: str, cast=float) -> object:
    """Prompt for an optional number; return None on empty input."""
    raw = input(f"  {prompt} [leave blank if unknown]: ").strip()
    if raw == "":
        return None
    try:
        return cast(raw)
    except (ValueError, TypeError):
        print("    (invalid — skipping, set to null)")
        return None


def _slugify(s: str) -> str:
    s = re.sub(r"[^\w\s-]", "", s).strip().lower()
    return re.sub(r"[\s-]+", "_", s)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    shared = _load_shared_settings()

    print("\n╔══════════════════════════════════════════════╗")
    print("║          PGIM  —  New Deal Setup Wizard      ║")
    print("╚══════════════════════════════════════════════╝")
    print("  Fill in the deal essentials below.")
    print("  Press Enter to accept [defaults]. Leave blank for optional fields.\n")

    # ── 1. Core identity ──────────────────────────────────────────────────────
    print("── Property ──────────────────────────────────────────────")
    deal_name     = _ask("Deal / asset name (short)", cast=str)
    property_name = _ask("Full property name", default=deal_name, cast=str)
    address       = _ask("Full street address (include city & country)", cast=str)

    # ── 2. Asset details ──────────────────────────────────────────────────────
    print("\n── Asset Details ─────────────────────────────────────────")
    asset_class_raw = _ask(
        "Asset class  (office / logistics / retail / industrial / mixed-use)",
        default="office", cast=str
    ).lower().strip()
    asset_class = asset_class_raw if asset_class_raw in _RADIUS_BY_CLASS else "office"

    gfa_raw  = _ask("GFA (enter number, unit will be set by country)", cast=float)
    quality  = _ask("Quality / grade", default="Grade A", cast=str)
    asset_type = _ask(
        "Asset type / transaction structure",
        default=f"Whole Block ({asset_class.title()})", cast=str
    )

    # ── 3. Deal numbers (all optional) ────────────────────────────────────────
    print("\n── Deal Numbers (press Enter to skip) ────────────────────")
    cur_year   = datetime.now().year
    sale_date  = _ask("Sale date label", default=f"{cur_year}E (Mktg)", cast=str)
    leasehold  = _ask("Remaining leasehold years (0 = freehold)", default=0, cast=int)
    price      = _ask_optional("Price (deal currency, millions) e.g. 775.0", cast=float)
    cap_rate   = _ask_optional("FTM NOI cap rate  e.g. 0.040 for 4.00%", cast=float)

    # ── 4. LLM derivation ─────────────────────────────────────────────────────
    _llm_cfg = {"provider": "ollama",
                "ollama": {"base_url": shared["ollama_base"],
                           "model":    shared["ollama_model"]}}
    print(f"\n  Deriving market fields from address via Ollama ({shared['ollama_model']}) …")
    try:
        derived = _derive_fields_with_llm(address, asset_class, _llm_cfg)
        country_name         = derived.get("country_name", "")
        country_code         = derived.get("country_code", "").lower()
        currency             = derived.get("currency", "SGD")
        currency_symbol      = derived.get("currency_symbol", currency)
        gfa_unit             = derived.get("gfa_unit", "sf").lower()
        land_zoning          = derived.get("land_zoning", "Commercial")
        location             = derived.get("location", "")
        asset_search_keyword = derived.get("asset_search_keyword", asset_class)
        submarket_keywords   = derived.get("submarket_keywords", [])
        broader_market_query = derived.get("broader_market_query", "")

        print(f"  ✓ Country: {country_name} ({country_code.upper()})  |  "
              f"Currency: {currency_symbol}  |  GFA unit: {gfa_unit}  |  "
              f"Location: {location}")
        print(f"  ✓ Submarket keywords: {submarket_keywords}")
        print(f"  ✓ Broader query: {broader_market_query}")

    except Exception as e:
        print(f"  ✗ LLM failed ({e}) — you'll need to fill in market fields manually.")
        country_name = country_code = currency = currency_symbol = ""
        gfa_unit = "sf"; land_zoning = "Commercial"; location = ""
        asset_search_keyword = asset_class; submarket_keywords = []; broader_market_query = ""

    # ── 5. Confirmation / override ────────────────────────────────────────────
    print("\n── Confirm or Override LLM suggestions ───────────────────")
    country_name         = _ask("Country name",         default=country_name)
    country_code         = _ask("Country code (2-letter ISO)", default=country_code).lower()
    currency             = _ask("Currency code",        default=currency)
    currency_symbol      = _ask("Currency symbol",      default=currency_symbol)
    gfa_unit             = _ask("GFA unit (sf / sqm)",  default=gfa_unit).lower()
    land_zoning          = _ask("Land zoning",          default=land_zoning)
    location             = _ask("Location descriptor",  default=location)
    asset_search_keyword = _ask("Asset search keyword", default=asset_search_keyword)

    print(f"  Submarket keywords {submarket_keywords}")
    override = input("  Override submarket keywords? (Enter to keep, or type comma-separated list): ").strip()
    if override:
        submarket_keywords = [k.strip() for k in override.split(",") if k.strip()]

    broader_market_query = _ask("Broader market query", default=broader_market_query)

    # ── 6. Files ──────────────────────────────────────────────────────────────
    slug         = _slugify(deal_name)
    folder_slug  = slug.replace("_", " ").title().replace(" ", "_")
    config_file  = _ask("Config file name",
                        default=f"configs/deal_config_{slug}.json")
    output_file  = _ask("Output Excel name",
                        default=f"output/{folder_slug}/Transaction_Comparables_{folder_slug}.xlsx")
    input_file   = _ask("Input Excel (curated comps, leave blank if online-only)",
                        default="", cast=str)

    # ── 7. Build config dict ──────────────────────────────────────────────────
    radii = _RADIUS_BY_CLASS.get(asset_class, _RADIUS_BY_CLASS["office"])
    search_cfg = {**_SEARCH_DEFAULTS, **radii}

    gfa_int = int(gfa_raw)

    config = {
        "subject_property": {
            "deal_name":               deal_name,
            "property_name":           property_name,
            "address":                 address,
            "sale_date":               sale_date,
            "land_zoning":             land_zoning,
            "remaining_leasehold_yrs": leasehold,
            "gfa_sf":                  gfa_int,
            "price_sgd_m":             price,
            "ftm_noi_cap_rate":        cap_rate,
            "location":                location,
            "quality":                 quality,
            "asset_type":              asset_type,

            "country_name":            country_name,
            "currency":                currency,
            "currency_symbol":         currency_symbol,
            "gfa_unit":                gfa_unit,
            "asset_class":             asset_class,
            "asset_search_keyword":    asset_search_keyword,
            "submarket_keywords":      submarket_keywords,
            "broader_market_query":    broader_market_query,
        },

        "country_code": country_code,

        "input_file":  input_file or None,
        "output_file": output_file,

        "parameters": {
            "bala_yield": 0.06,
            "max_comps":  10,
        },

        "openai": {
            "api_key":       None,
            "search_model":  "gpt-4o-mini-search-preview",
            "extract_model": "gpt-4o-mini",
        },

        "online_search": search_cfg,

        "mapbox": {
            "token":    shared["mapbox_token"],
            "style":    "streets-v12",
            "width":    1200,
            "height":   900,
            "padding":  100,
            "pin_size": "l",
        },

        "llm": {
            "provider": "ollama",
            "ollama": {
                "base_url": shared["ollama_base"],
                "model":    shared["ollama_model"],
            },
        },
    }

    # Remove null input_file key if blank
    if config["input_file"] is None:
        del config["input_file"]

    # ── 8. Save ───────────────────────────────────────────────────────────────
    out_path = Path(config_file)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\n╔══════════════════════════════════════════════╗")
    print(f"║  ✓  Config saved → {config_file:<26}║")
    print(f"╚══════════════════════════════════════════════╝")
    print(f"\n  Next steps:")
    print(f"    python3 search_online_comps.py --config {config_file} --map")
    if input_file:
        print(f"    python3 generate_comps_table.py --config {config_file}")
    print()


# ─────────────────────────────────────────────────────────────────────────────
# PUBLIC API  (used by the Streamlit dashboard — no CLI prompts)
# ─────────────────────────────────────────────────────────────────────────────

# Keywords that typically label or precede an address in deal documents
_ADDR_LABEL_KWS = [
    "address", "location", "property address", "site address",
    "property location", "situated at", "located at", "premises",
]

# Street-type words — a line containing one of these is likely an address line
_STREET_KWS = [
    "street", "road", "avenue", "boulevard", "drive", "lane", "place",
    "way", "crescent", "close", "court", "terrace", "jalan", "lorong",
    "tanjong", "bukit", "tower", "centre", "center", "plaza", "park",
    "building", "house", "block",
]

# Country/city names common in APAC deal documents
_COUNTRY_KWS = [
    "singapore", "south korea", "korea", "japan", "hong kong", "australia",
    "malaysia", "indonesia", "thailand", "vietnam", "india", "china",
    "taiwan", "philippines",
]


def _extract_address_candidates(text: str, max_candidates: int = 10) -> str:
    """
    Scan the FULL document text for lines that look like an address label or
    contain address keywords.  Returns a deduplicated block of candidate lines
    so the LLM can focus on them regardless of where they appear in the document.

    Strategy (in order of confidence):
      1. Lines immediately following a label keyword (e.g. "Address: 88 Cecil St")
      2. Lines that contain a street-type word AND a digit (e.g. "88 Cecil Street")
      3. Lines that contain a country/city name (e.g. "Singapore 069538")
    """
    lines     = [ln.strip() for ln in text.splitlines() if ln.strip()]
    candidates = []
    seen       = set()

    def _add(line: str):
        key = line.lower()
        if key not in seen and len(line) > 4:
            seen.add(key)
            candidates.append(line)

    for i, line in enumerate(lines):
        low = line.lower()

        # Strategy 1: line starts with or contains an address label keyword
        for kw in _ADDR_LABEL_KWS:
            if kw in low:
                _add(line)
                # Also grab the next line — label and value are often on separate lines
                if i + 1 < len(lines):
                    _add(lines[i + 1])
                break

        # Strategy 2: contains a street-type word AND a digit
        if any(sw in low for sw in _STREET_KWS) and re.search(r'\d', line):
            _add(line)

        # Strategy 3: contains a country/city name
        if any(ck in low for ck in _COUNTRY_KWS):
            _add(line)

        if len(candidates) >= max_candidates:
            break

    return "\n".join(candidates)


def extract_from_document(text: str, llm_cfg: dict, openai_key: str = "") -> dict:
    """
    Extract basic deal fields from raw document text (PDF / Excel / deal brief).

    Called from the Streamlit dashboard when the user uploads or pastes a deal
    brief in the New Deal wizard.  The LLM reads up to 4000 characters of the
    document body, PLUS a focused block of address-candidate lines extracted
    from the entire document — so the address is found even if it appears deep
    in the file.

    Any field not found in the document is returned as None so the wizard can
    fall back to the user's manually entered values or LLM-derived defaults.

    Returns a dict with the same keys as subject_property in the deal config.
    """
    # Pre-scan the FULL text for address candidate lines (keyword-driven)
    addr_candidates = _extract_address_candidates(text)

    system = (
        "You are a real estate data extractor. "
        "Read the document and extract deal information. "
        "Return ONLY a JSON object with these exact keys — use null if not found:\n"
        "  deal_name               : short asset / deal name\n"
        "  property_name           : full official property name\n"
        "  address                 : complete street address including city and country.\n"
        "                            Look for labels like 'Address:', 'Location:', 'Property Address:',\n"
        "                            'Situated at:', or any line with a street number + street name.\n"
        "  asset_class             : one of office / retail / logistics / industrial / mixed-use\n"
        "  gfa_sf                  : gross floor area as a plain number (no units, no commas)\n"
        "  quality                 : building grade e.g. Grade A, Grade B+\n"
        "  asset_type              : e.g. Whole Block (Office), Strata (Retail)\n"
        "  sale_date               : marketing date label e.g. 2025E (Mktg)\n"
        "  remaining_leasehold_yrs : remaining leasehold years as integer; 0 if freehold\n"
        "  price_sgd_m             : asking price in millions (deal currency); null if unknown\n"
        "  ftm_noi_cap_rate        : forward NOI cap rate as a decimal e.g. 0.04; null if unknown\n"
    )

    # Build user message: focused address candidates first, then document body
    addr_block = (
        f"ADDRESS CANDIDATE LINES (extracted from full document):\n{addr_candidates}\n\n"
        if addr_candidates else ""
    )
    user = (
        f"{addr_block}"
        f"Extract deal information from this document:\n\n{text[:4000]}"
    )

    raw = _llm_post(
        [{"role": "system", "content": system}, {"role": "user", "content": user}],
        llm_cfg, openai_key, timeout=300,   # 5 min — 8b+ models on CPU can be slow
    )
    raw = re.sub(r"^```[a-z]*\n?", "", raw.strip())
    raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw)


def derive_market_fields(address: str, asset_class: str,
                          llm_cfg: dict, openai_key: str = "",
                          mapbox_token: str = "") -> dict:
    """
    Use the LLM to infer all market / location config fields from an address,
    then override location and submarket_keywords with real geodata from the
    Mapbox Geocoding API if a token is provided.

    Given just the property address and asset class, the LLM returns:
      country_name, country_code, currency, currency_symbol, gfa_unit,
      land_zoning, location (submarket descriptor), asset_search_keyword,
      submarket_keywords (list), broader_market_query

    If mapbox_token is supplied, the Mapbox Geocoding API is called to look up
    the real neighbourhood / district names for the address. These replace the
    LLM's guessed submarket_keywords, which are often inaccurate on small models.

    This is what drives the "Generate Config Preview" button in the dashboard —
    the user types an address and the LLM fills in all the remaining config
    fields automatically.  The user can then review and correct any value in
    Step 2 of the wizard before saving.
    """
    provider = (llm_cfg or {}).get("provider", "ollama")
    if provider in ("none", "rules"):
        print("  Deriving market fields via RULES (no LLM) …")
        fields = _derive_fields_rules(address, asset_class)
    else:
        fields = _derive_fields_with_llm(address, asset_class, llm_cfg, openai_key)

    # Override submarket with real geodata from Mapbox if token is available
    if mapbox_token:
        geo = _mapbox_geocode(address, mapbox_token)
        if geo:
            # Build submarket keywords from the place hierarchy Mapbox returned
            # Priority: neighborhood > locality > district (most granular first)
            geo_keywords = []
            for key in ("neighborhood", "locality", "district"):
                val = geo.get(key)
                if val and val not in geo_keywords:
                    geo_keywords.append(val)

            if geo_keywords:
                fields["submarket_keywords"] = geo_keywords
                # Update location descriptor to the most granular name found
                fields["location"] = geo_keywords[0]
                print(f"  [mapbox geocode] location: {geo_keywords[0]}  |  "
                      f"submarket keywords: {geo_keywords}")

    return fields


def build_config(fields: dict, shared: dict = None) -> tuple:
    """
    Assemble the full deal config dict from a flat fields dict.

    Takes the fields collected by the New Deal wizard (a mix of user-typed,
    LLM-extracted, and LLM-derived values) and wraps them into the complete
    config structure that all backend scripts expect:
      subject_property, country_code, input_file, output_file, parameters,
      llm, mapbox, openai, online_search

    Shared settings (Mapbox token, Ollama URL) are read from any existing deal
    config via _load_shared_settings() so they stay in sync without the user
    having to re-enter them for each new deal.

    Type coercion is applied: GFA and leasehold years become ints; price and
    cap rate become floats (or None if blank).

    Returns (config_dict, suggested_config_file_path).
    """
    if shared is None:
        shared = _load_shared_settings()

    asset_class = (fields.get("asset_class") or "office").lower().strip()
    if asset_class not in _RADIUS_BY_CLASS:
        asset_class = "office"

    radii      = _RADIUS_BY_CLASS[asset_class]
    search_cfg = {**_SEARCH_DEFAULTS, **radii}

    deal_name   = (fields.get("deal_name") or fields.get("property_name") or "New Deal").strip()
    slug        = _slugify(deal_name)
    folder_slug = slug.replace("_", " ").title().replace(" ", "_")

    def _safe_int(v, default=0):
        try:    return int(float(v))
        except: return default

    def _safe_float(v):
        try:    return float(v)
        except: return None

    config = {
        "subject_property": {
            "deal_name":               deal_name,
            "property_name":           (fields.get("property_name") or deal_name).strip(),
            "address":                 fields.get("address", ""),
            "sale_date":               fields.get("sale_date") or f"{datetime.now().year}E (Mktg)",
            "land_zoning":             fields.get("land_zoning", "Commercial"),
            "remaining_leasehold_yrs": _safe_int(fields.get("remaining_leasehold_yrs"), 0),
            "gfa_sf":                  _safe_int(fields.get("gfa_sf"), 0),
            "price_sgd_m":             _safe_float(fields.get("price_sgd_m")),
            "ftm_noi_cap_rate":        _safe_float(fields.get("ftm_noi_cap_rate")),
            "location":                fields.get("location", ""),
            "quality":                 fields.get("quality", ""),
            "asset_type":              fields.get("asset_type") or f"Whole Block ({asset_class.title()})",
            "country_name":            fields.get("country_name", ""),
            "currency":                fields.get("currency", "SGD"),
            "currency_symbol":         fields.get("currency_symbol", "S$"),
            "gfa_unit":                (fields.get("gfa_unit") or "sf").lower(),
            "asset_class":             asset_class,
            "asset_search_keyword":    fields.get("asset_search_keyword", asset_class),
            "submarket_keywords":      fields.get("submarket_keywords") or [],
            "broader_market_query":    fields.get("broader_market_query", ""),
        },
        "country_code": (fields.get("country_code") or "").lower(),
        "output_file":  f"output/{folder_slug}/Transaction_Comparables_{folder_slug}.xlsx",
        "parameters":   {"bala_yield": 0.06, "max_comps": 10},
        "openai": {
            "api_key":       None,
            "search_model":  "gpt-4o-mini-search-preview",
            "extract_model": "gpt-4o-mini",
        },
        "online_search": search_cfg,
        "rent_search": {
            "min_results":  5,
            "max_level":    3,
            "proximity_km": 0.5,
            "submarket_km": 2.0,
            "market_km":    3.0,
        },
        "mapbox": {
            "token":    shared.get("mapbox_token", ""),
            "style":    "streets-v12",
            "width":    1200,
            "height":   900,
            "padding":  100,
            "pin_size": "l",
        },
        "llm": {
            "provider": "ollama",
            "ollama": {
                "base_url": shared.get("ollama_base", "http://localhost:11434"),
                "model":    shared.get("ollama_model", "qwen2.5:3b"),
            },
        },
    }

    config_file = f"configs/deal_config_{slug}.json"
    return config, config_file


def save_config(config: dict, config_file: str) -> str:
    """Save config dict to JSON. Returns the saved file path."""
    out = Path(config_file)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
    return str(out)


if __name__ == "__main__":
    main()
