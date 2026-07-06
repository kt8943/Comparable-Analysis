"""
tools/column_mapper.py
======================
Tiered column-header mapper.

  Tier 1 – Exact match     normalised header == known synonym        fast, zero false-positives
  Tier 2 – Embedding       cosine similarity vs field synonym corpus  offline, deterministic
  Tier 3 – LLM             GPT / Ollama with full context            called only for remaining unknowns

Also detects units from header text and returns a per-field multiplier so
callers can normalise values (sqm→SF, S$000→S$M, psm→psf) without extra logic.

Public API
----------
map_columns(
    headers, sample_rows, output_fields, col_to_key,
    base_url, model,
    extra_fields=None,
    passthrough_units=False,
) -> (col_map, unit_map)

    col_map  : {internal_key: int | None}   — 0-based column index
    unit_map : {internal_key: float}        — multiply raw cell value by this

detect_unit_multiplier(header, field_key) -> float

map_columns_ollama(...)  — legacy alias, returns only col_map
"""

import json
import re

try:
    from fastembed import TextEmbedding as _TE
    import numpy as _np
    _EMBED_AVAILABLE = True
except ImportError:
    _EMBED_AVAILABLE = False

from tools.llm_client import ollama_post, openai_chat

# Minimum quality bar for LLM-tier column candidates:
#   header must contain ≥4 consecutive letters (rules out '', ')', 'T', truncated fragments)
_USABLE_RE = re.compile(r'[a-zA-Z]{4,}')


# ═══════════════════════════════════════════════════════════════════════════════
# SYNONYM DICTIONARY
#
# Derived from the actual output column names in generate_*_comps_table.py,
# plus common input variations seen in real estate comp Excel files.
#
# Rules:
#  • Each entry anchors to the real output column name (normalised, no units)
#  • Only multi-word phrases or highly specific single words are listed
#  • Generic single words ("area", "date", "rent") are NOT used alone —
#    they must appear as part of a phrase to avoid false-positive matches
#  • Only keys that appear in the active output_fields are scored,
#    so cross-comp-type overlaps do not cause conflicts
# ═══════════════════════════════════════════════════════════════════════════════

_SYNONYMS: dict[str, list[str]] = {

    # ── SALES ─────────────────────────────────────────────────────────────────
    # Output: "Property"
    "property_name": [
        "property name", "property",
        "building name", "building",
        "asset name", "asset",
        "development name",
        "location", "site",
        # "location" catches GLS / government land sales tables where the
        # site identifier column is labelled "LOCATION".
        # "development" removed — too ambiguous (matches "TYPE OF DEVELOPMENT ALLOWED").
    ],
    # Output: "Sale Date"
    "sale_date_raw": [
        "sale date", "transaction date", "transacted date",
        "completion date", "date of transaction",
        "date of award", "award date",  # GLS / public-tender tables
        "period", "quarter", "sale period", "transaction period",
    ],
    # Output: "Land Zoning"
    "land_zoning": [
        "land zoning", "land use", "zoning",
        "planning zone", "permitted use",
        "development type", "type of development", "type of development allowed",
        "sector",
        # Asset/property type columns (e.g. Cushman KR marketbeat "PROPERTY TYPE"
        # with values Office/Logistics/Hospitality). These belong in zoning/use,
        # NOT address — exact match here stops embedding sending them to address.
        "property type", "asset type", "asset class",
        "building type", "property use", "use",
    ],
    # Output: "Remaining Leasehold (Y)"
    "remaining_yrs": [
        "remaining leasehold", "remaining lease", "unexpired lease",
        "leasehold years", "leasehold remaining",
        "tenure", "leasehold",
    ],
    # Output: "GFA (SF)"
    "gfa_sf": [
        "gfa", "gross floor area",
        "nla", "net lettable area",
        "floor area", "lettable area",
        "transacted area", "transaction area",
        "net floor area", "total area", "building area",
    ],
    # Output: "Price (SGD M)"
    "price_sgd_m": [
        "sale price", "transaction price", "transacted price",
        "purchase price", "consideration", "land price",
        # GLS / land tender headers (total tendered amount, not psf)
        "successful tender price", "tendered price", "tender price",
        "winning bid", "winning tender", "bid price", "tender amount",
        "successful bid", "awarded price", "land cost",
        # Plain 'Price' header (e.g. Colliers investment report)
        "price",
        # "PRICE (<currency/unit>)" headers (PRICE (S$M), PRICE (KRW), Sale Price
        # (S$ million), …) are resolved by the embedding tier — no need to
        # enumerate currencies here (doing so pollutes the field's embedding).
    ],
    # Output: "Price (SGD psf GFA)" — unit price per SF of GFA
    "price_psf_gfa": [
        "price psf gfa", "psf gfa", "price per sf gfa",
        "price sgd psf gfa", "price psf", "unit price psf",
        "price per sq ft", "price per sqft",
    ],
    # Output: "FTM NOI Capitalisation Rate"
    "npi_yield": [
        "ftm noi capitalisation rate", "ftm noi cap rate",
        "npi yield", "cap rate", "capitalisation rate",
        "net yield", "net initial yield",
    ],
    # Output: "Adj. Capitalisation Rate"
    "adj_npi_yield": [
        "adj capitalisation rate", "adj cap rate",
        "adjusted capitalisation rate", "adjusted cap rate",
        "adj npi yield", "bala adjusted cap rate",
    ],
    # Output: "Sale Type"  (input extraction — not a standalone output column)
    "sale_type": [
        "sale type", "type of sale", "transaction type",
        "deal structure", "deal type",
    ],
    # Output: "Buyer" (global sales comps)
    "buyer": [
        "buyer", "purchaser", "acquirer", "acquiring entity",
        "buyer name", "purchaser name",
    ],

    # ── RENT ──────────────────────────────────────────────────────────────────
    # Output: "Property"
    "building_name": [
        "property", "property name",
        "building", "building name",
        "development", "asset",
    ],
    # Output: (address — embedded, not standalone output column)
    "address": [
        "property address", "address", "street address", "street",
    ],
    # Output: "Location"
    "district": [
        "location", "district", "planning area", "submarket",
    ],
    # Output: "Quality"
    "quality": [
        "quality", "building grade", "grade", "building quality",
    ],
    # Output: "Leased GLA (SF)"
    "nla_sf": [
        "leased gla", "nla", "net lettable area",
        "gfa", "gross floor area", "gla", "leased area",
        "floor area",
        # Bare area-unit headers used in lease-transaction tables (e.g. Cushman
        # office marketbeat "SF" column = space leased).
        "sf", "area (sf)", "leased sf", "area sf", "sqm", "area (sqm)",
    ],
    # Output: "Gross Face Rents (SGD psf pm)"
    "asking_rent": [
        "gross face rents", "asking rent", "gross rent",
        "face rent", "headline rent", "passing rent",
        "rent psf pm",
    ],
    # Output: "Effective Rents (SGD psf pm)"
    "eff_rent": [
        "effective rents", "effective rent", "net effective rent",
        "net rent", "eff rent",
    ],
    # Output: "Date of Lease Start"
    "lease_date": [
        "date of lease start", "lease date", "lease start",
        "commencement date", "lease commencement", "start date",
    ],
    # Output: "Lease Tenure (Yrs)"
    "lease_term_yrs": [
        "lease tenure", "lease term", "term yrs",
        "lease duration", "lease period",
    ],
    # Output: "Rent-Free (Mths)"
    "rent_free_mths": [
        "rent-free", "rent free", "rent free period",
        "rf period", "incentive period",
    ],
    # Output: "Tenant"  (rent/lease comps — the occupier in a lease deal)
    "tenant": [
        "tenant", "occupier", "lessee", "occupant", "tenant name",
    ],
    # Output: "Type of Lease Area / Comments"
    "lease_type": [
        "type of lease area", "lease type", "space type",
        "asset type", "type", "remarks", "comments",
    ],

    # ── LAND ──────────────────────────────────────────────────────────────────
    # Output: "Property"
    "site_name": [
        "property", "site", "site name",
        "project", "project name",
        "asset", "land parcel",
        # "location" / "site location" — in GLS / tender tables the site
        # identifier column is labelled "LOCATION" (same as sales property_name).
        "location", "site location",
        # Note: bare "development" removed — too ambiguous.
    ],
    # Output: "Date of Launch"
    "launch_date": [
        "date of launch", "launch date", "award date",
        "tender award date", "transaction date", "date of award",
        "tender closing date", "tender date", "date of tender",
        "successful tender date", "date awarded",
    ],
    # Output: "Land Tenure (Y)"
    "tenure": [
        "land tenure", "tenure", "lease tenure",
        "leasehold", "land leasehold",
    ],
    # Output: "Site Area (SF)"
    "site_area_sf": [
        "site area", "land area", "plot area",
        "lot area", "site size",
    ],
    # Output: "Max GFA (SF)"
    "max_gfa_sf": [
        "max gfa", "maximum gfa", "permissible gfa",
        "allowable gfa", "developable gfa",
    ],
    # Output: "Price (SGD psf ppr)"
    "price_psf_ppr": [
        "price sgd psf ppr", "psf ppr", "land price psf ppr",
        "unit land price", "price per sqft per pr",
        "psf ppr gfa", "psm ppr", "per plot ratio", "psf gpr",
        "tendered price psf ppr", "successful tender psf ppr",
        "land rate psf ppr", "land rate", "unit price psf ppr",
    ],
    # Output: "Comment"
    "remarks": [
        "comment", "comments", "remarks", "notes",
    ],
}


# ═══════════════════════════════════════════════════════════════════════════════
# EMBEDDING MODEL (Tier 2)
# ═══════════════════════════════════════════════════════════════════════════════

_embed_model = None
_field_embed_cache: dict = {}  # field_key → np.ndarray (cached across calls)


def _get_embed_model():
    global _embed_model
    if _embed_model is None:
        _embed_model = _TE("sentence-transformers/all-MiniLM-L6-v2")
    return _embed_model


def _embed(model, text: str) -> "_np.ndarray":
    """Single-text embed via fastembed (already L2-normalised)."""
    return _np.array(next(model.embed([text])))


def _build_field_embeddings(field_keys: list, output_fields: list,
                             extra_fields: list) -> dict:
    """Embed synonym corpus for each field key. Results cached in _field_embed_cache."""
    model = _get_embed_model()
    desc_map = {key: desc for _, key, desc in output_fields}
    if extra_fields:
        desc_map.update({key: desc for key, desc in extra_fields})

    result = {}
    for fk in field_keys:
        if fk in _field_embed_cache:
            result[fk] = _field_embed_cache[fk]
            continue
        synonyms = _SYNONYMS.get(fk, [])
        desc = desc_map.get(fk, "")
        corpus = ", ".join(synonyms)
        if desc:
            corpus = f"{corpus}. {desc}"
        vec = _embed(model, corpus)
        _field_embed_cache[fk] = vec
        result[fk] = vec
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# UNIT DETECTION
# ═══════════════════════════════════════════════════════════════════════════════

_AREA_KEYS     = {"gfa_sf", "nla_sf", "site_area_sf", "max_gfa_sf"}
_PRICE_M_KEYS  = {"price_sgd_m"}
_RENT_PSF_KEYS = {"asking_rent", "eff_rent"}


def detect_unit_multiplier(header: str, field_key: str) -> float:
    """
    Detect the unit from a column header and return the multiplier needed to
    convert raw cell values to the canonical unit for that field.

    Area fields      → canonical: SF.          sqm/m² header  → × 10.7639
    Price M fields   → canonical: S$M.         S$'000 header  → × 0.001
    Rent PSF fields  → canonical: S$ PSF/mth.  PSM header     → × 10.7639
    All other fields → 1.0 (no conversion)
    """
    h = header.lower()

    if field_key in _AREA_KEYS:
        # sqm / m² / sq m → convert to SF
        if re.search(r"sq\.?\s*m(?!i|ft|f\b)|sqm\b|m²|㎡|m2\b|\(m2\)|\(sqm\)", h):
            return 10.7639
        return 1.0  # already SF (or unspecified — default)

    if field_key in _PRICE_M_KEYS:
        # Billions → multiply by 1000 to get millions
        if re.search(r"\bbillion\b|\bbn\b|\bbil\b|\bb\b", h):
            return 1000.0
        # S$’000 / (000) / thousands → convert to millions
        if re.search(r"[‘’’]000|\b000s?\b|\(000\)|\(‘000\)|thousands?", h):
            return 0.001
        # Raw S$ (actual dollars, not millions — uncommon in comp tables)
        if re.search(r"\(\s*s\$\s*\)|\bsgd\s*\)|\bsin\s*\)", h) and not re.search(r"[mk]", h):
            return 1e-6
        return 1.0  # assume already in millions

    if field_key in _RENT_PSF_KEYS:
        # PSM / per sqm → convert to PSF/month
        if re.search(r"\bpsm\b|per\s*sq\.?\s*m(?!i|ft)|per\s*sqm\b", h):
            return 10.7639
        return 1.0  # already PSF/month

    return 1.0


# ═══════════════════════════════════════════════════════════════════════════════
# NORMALISATION + SCORING HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _norm(text: str) -> str:
    """Lowercase, strip punctuation, collapse whitespace."""
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", text.lower())).strip()


# General price-header detection (used by the Tier 1.5 rule in map_columns).
# Operate on _norm()'d headers, so brackets/symbols are already stripped to spaces
# (e.g. "PRICE (S$M)" -> "price s m", "PRICE / UNIT (Mn. KRW/3.3㎡)" -> "price unit mn krw 3 3").
_PRICE_RE      = re.compile(r"\b(price|consideration)\b")
_UNIT_PRICE_RE = re.compile(r"\b(psf|psm|ppr|per|unit|sqft|sqm)\b|/\s*sq")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

def map_columns(
    headers: list,
    sample_rows: list,
    output_fields: list,
    col_to_key: dict,  # noqa: ARG001 — kept for backward-compat callers
    base_url: str = "",
    model: str = "",
    extra_fields: list = None,
    llm_cfg: dict = None,
    passthrough_units: bool = False,
) -> tuple[dict, dict]:
    """
    Map input Excel/PDF headers to schema field keys.

    Tier 1 — Exact match  : normalised header == known synonym (fast, zero false-positives)
    Tier 2 — Embedding    : cosine similarity vs per-field synonym corpus (all-MiniLM-L6-v2)
                            Skipped gracefully if sentence-transformers is not installed.
    Tier 3 — LLM          : GPT / Ollama for columns still unresolved after embedding.
                            Only called when Tier 1+2 leave gaps.

    Returns
    -------
    col_map  : {internal_key: int | None}  — 0-based column index, None if not found
    unit_map : {internal_key: float}       — multiply raw cell value by this factor
    """
    all_keys = [key for _, key, _ in output_fields] + [k for k, _ in (extra_fields or [])]
    n = len(headers)
    headers_norm = [_norm(h) for h in headers]

    col_map:  dict[str, int | None] = {k: None for k in all_keys}
    unit_map: dict[str, float]      = {k: 1.0  for k in all_keys}
    claimed:  dict[int, str]        = {}  # col_idx → field_key that claimed it

    def _assign(field_key: str, col_idx: int, method: str) -> None:
        col_map[field_key]  = col_idx
        unit_map[field_key] = (1.0 if passthrough_units
                               else detect_unit_multiplier(headers[col_idx], field_key))
        claimed[col_idx]    = field_key
        mult_tag = f" ×{unit_map[field_key]:.4g}" if unit_map[field_key] != 1.0 else ""
        print(f"    {field_key:<22} → col {col_idx} ({headers[col_idx]!r}){mult_tag}  [{method}]")

    def _build_llm_inputs(keys_to_resolve: list) -> tuple[dict, dict, str]:
        """Build inputs shared by both GPT and Ollama LLM calls."""
        all_cols = {str(i): h for i, h in enumerate(headers) if h}
        all_descs = {}
        for _, key, desc in output_fields:
            if key in keys_to_resolve:
                all_descs[key] = desc
        for key, desc in (extra_fields or []):
            if key in keys_to_resolve:
                all_descs[key] = desc
        sample_str = "\n".join(
            str({str(j): v for j, v in enumerate(r) if v is not None})
            for r in (sample_rows or [])[:3]
        )
        return all_cols, all_descs, sample_str

    def _apply_llm_result(llm_result: dict, keys_to_resolve: list, method: str) -> None:
        for field_key, val in llm_result.items():
            if field_key not in keys_to_resolve or col_map.get(field_key) is not None:
                continue
            if val is None or val == "null":
                continue
            try:
                col_idx = int(val)
            except (TypeError, ValueError):
                continue
            if 0 <= col_idx < n and col_idx not in claimed:
                _assign(field_key, col_idx, method)

    # ── Tier 1: Exact match ────────────────────────────────────────────────────
    # Normalize synonyms the same way headers are normalized so that headers
    # like 'PRICE (S$ Million)' → 'price s million' match the synonym
    # "price s million" even though the raw synonym has no special chars.
    syns_norm = {fk: {_norm(s) for s in _SYNONYMS.get(fk, [])} for fk in all_keys}

    # ── Tier 0: value-based disambiguation (runs before header matching) ───────
    # A "Location" header is ambiguous: in GLS/land tables it's the SITE identifier,
    # but in sales/rent comp tables users often put the location-competitiveness
    # LABEL (Superior / Comparable / Inferior) there. Decide by the column's VALUES:
    # if they are competitiveness labels it is the Location label — never a property
    # name or street address (mapping it there geocodes every comp to the country
    # centroid). Map it to the "location" field if the schema has one, else claim it
    # so it is simply excluded from geocoding.
    _COMPET = ("superior", "comparable", "inferior")
    _label_key = ("location" if "location" in col_map
                  else "district" if "district" in col_map else None)
    for i in range(n):
        if i in claimed:
            continue
        vals = [str(r[i]).strip().lower() for r in (sample_rows or [])
                if i < len(r) and r[i] not in (None, "")]
        if not vals:
            continue
        hits = sum(1 for v in vals if any(w in v for w in _COMPET))
        if hits >= max(1, (len(vals) + 1) // 2):
            if _label_key:
                _assign(_label_key, i, "value:location-label")
            else:
                claimed[i] = "_location_label"
                print(f"    location(label)        → col {i} ({headers[i]!r})  "
                      f"[value:competitiveness — excluded from geocoding]")

    for field_key in all_keys:
        for i, hn in enumerate(headers_norm):
            if i in claimed or not hn:
                continue
            if hn in syns_norm.get(field_key, set()):
                _assign(field_key, i, "exact")
                break

    # ── Tier 1.5: General price rule (deterministic, no threshold) ─────────────
    # Any header containing "price" (or "consideration") maps to the *total* price
    # field — regardless of the currency/unit in brackets (PRICE (S$M), PRICE (KRW),
    # Sale Price (S$ million)). A *per-unit* price (psf/psm/ppr/per/unit) maps to the
    # unit-price field instead. Runs before the embedding/LLM tiers so an obvious
    # price column is never lost to a weak score (e.g. PRICE (S$M) embeds at only
    # 0.26 < 0.30) or stolen by a fuzzy near-match — and it works in rule-based mode.
    for i, hn in enumerate(headers_norm):
        if i in claimed or not hn or not _PRICE_RE.search(hn):
            continue
        tgt = "price_psf_gfa" if _UNIT_PRICE_RE.search(hn) else "price_sgd_m"
        if col_map.get(tgt) is None and tgt in col_map:
            _assign(tgt, i, "price-rule")

    unresolved = [k for k in all_keys if col_map[k] is None]
    if not unresolved:
        return col_map, unit_map

    # ── Tier 2: Embedding similarity ──────────────────────────────────────────
    if unresolved and _EMBED_AVAILABLE:
        emb_model  = _get_embed_model()
        field_vecs = _build_field_embeddings(unresolved, output_fields, extra_fields)
        _EMBED_THRESHOLD = 0.30
        for i, h in enumerate(headers):
            if i in claimed or not h:
                continue
            hn = _norm(h)
            q_vec = _embed(emb_model, hn)
            best_score, best_key = 0.0, None
            for fk in unresolved:
                if col_map[fk] is not None:
                    continue
                score = float(_np.dot(q_vec, field_vecs[fk]))
                if score > best_score:
                    best_score, best_key = score, fk
            if best_key and best_score >= _EMBED_THRESHOLD:
                _assign(best_key, i, f"embed({best_score:.2f})")

    unresolved = [k for k in all_keys if col_map[k] is None]
    if not unresolved:
        return col_map, unit_map

    # ── Tier 3: LLM — only for columns still unresolved after embedding ────────
    provider = (llm_cfg or {}).get("provider", "ollama")

    # Rule-based mode: no LLM tier. Any columns Tier 1/2 couldn't resolve stay None.
    if provider in ("none", "rules"):
        for field_key in all_keys:
            if col_map[field_key] is None:
                print(f"    {field_key:<22} → not found")
        return col_map, unit_map

    all_cols, all_descs, sample_str = _build_llm_inputs(unresolved)
    synonym_hints = "\n".join(
        f'  "{k}": {", ".join(_SYNONYMS[k])}'
        for k in unresolved
        if k in _SYNONYMS and _SYNONYMS[k]
    )
    system = (
        "You are a real estate data extraction assistant. "
        "Map column headers to schema field keys using SEMANTIC understanding. "
        "Use the column headers AND sample row values to infer the true meaning — "
        "do not rely on surface word overlap. "
        "Critical rules: "
        "'Transacted Area (sqm)' / 'Transaction Area' are AREA fields (gfa_sf / nla_sf / site_area_sf), NOT price. "
        "'Sale Price (KRW B)' / 'Transaction Price' are PRICE fields. "
        "The word 'transacted' alone does not mean price — look at whether the column says 'area' or 'price'. "
        "Return ONLY a valid JSON object: {\"field_key\": column_index_or_null}. "
        "Use null when no column matches a field. One column per field."
    )
    user = (
        f"All column headers (index: name):\n{json.dumps(all_cols, indent=2)}\n\n"
        f"Sample data rows:\n{sample_str}\n\n"
        "Fields to map (key: description):\n"
        + "\n".join(f'  "{k}": {d}' for k, d in all_descs.items())
        + ("\n\nSynonym hints (not strict rules):\n" + synonym_hints if synonym_hints else "")
    )
    messages = [
        {"role": "system", "content": system},
        {"role": "user",   "content": user},
    ]

    if provider == "openai":
        try:
            raw = openai_chat(llm_cfg, messages, json_mode=True)
            raw = re.sub(r"^```[a-z]*\n?", "", raw.strip())
            raw = re.sub(r"\n?```$", "", raw)
            _apply_llm_result(json.loads(raw), unresolved, "gpt")
        except Exception as e:
            print(f"  [column_mapper] GPT mapping failed: {e}")

    elif base_url:
        try:
            raw = ollama_post(base_url, model, messages)
            raw = re.sub(r"^```[a-z]*\n?", "", raw.strip())
            raw = re.sub(r"\n?```$", "", raw)
            _apply_llm_result(json.loads(raw), unresolved, "ollama")
        except Exception as e:
            print(f"  [column_mapper] Ollama mapping failed: {e}")

    # Log fields that couldn't be resolved by any tier
    for field_key in all_keys:
        if col_map[field_key] is None:
            print(f"    {field_key:<22} → not found")

    return col_map, unit_map


# ── Backward-compat alias ──────────────────────────────────────────────────────

def map_columns_ollama(
    headers: list,
    sample_rows: list,
    output_fields: list,
    col_to_key: dict,
    base_url: str,
    model: str,
    system_preamble: str = "",
    extra_fields: list = None,
) -> dict:
    """Legacy alias — returns only col_map (drops unit_map). Use map_columns() for new code."""
    col_map, _ = map_columns(
        headers, sample_rows, output_fields, col_to_key,
        base_url, model, extra_fields=extra_fields,
    )
    return col_map
