"""
generate_investment_rationale.py
=================================
Two-stage LLM pipeline that reads market report PDFs and writes an
institutional-quality investment rationale for a subject property.

─── HOW IT WORKS ────────────────────────────────────────────────────────────

Stage 1 — extract_report_insights()
    Each market report PDF is read and passed to the LLM, which extracts
    structured facts (vacancy, rents, yields, outlook, key statistics) into
    a JSON object.  The result is saved to a cache file keyed by the PDF's
    filename + size + modified-time, so the same unchanged file is never
    re-extracted.  Only new or force-refreshed PDFs trigger an LLM call.

    PDF / txt report  →  LLM (temp=0.1)  →  structured JSON
    Cached at:  Input_files/market_reports/cache/<stem>_<hash>.json

Stage 2 — generate_rationale()
    Uses the cached insights from Stage 1 plus the deal config (subject
    property details) to write a 3–5 section investment committee memo.

    This stage runs TWO separate LLM calls:

    Call 1 — Prose writing (sources anonymised)
        The research data is passed with source labels replaced by generic
        names ("Research Report 1", "Research Report 2"…) so the LLM cannot
        echo real PDF filenames into the client-facing memo.  The LLM follows
        a STEP 1→2→3→4 chain-of-thought: first retrieving all evidence,
        then planning sections, then writing prose, then self-verifying.

    Call 2 — Source audit JSON (real filenames visible)
        The finished prose + research data with real filenames are sent to a
        second LLM call whose only job is to match every specific claim back
        to its source PDF and page number.  Output populates Source_Audit.xlsx.

    Cached insights + deal config  →  LLM (temp=0.3)  →  markdown rationale
    Output saved to: output/<deal_dir>/Investment_Rationale.md
                     output/<deal_dir>/Source_Audit.xlsx

─── CLI USAGE ───────────────────────────────────────────────────────────────
python generate_investment_rationale.py \\
    --config  configs/deal_config_xxx.json \\
   [--reports Input_files/market_reports/report1.pdf report2.pdf ...] \\
   [--refresh]          # ignore cache and re-extract every report \\
   [--notes-file /tmp/analyst_notes.txt]

─── KEY DESIGN DECISIONS ────────────────────────────────────────────────────
• File-hash caching      — same PDF is never re-extracted unless it changes
• Two separate LLM calls — keeps filenames out of prose; tracks sources cleanly
• Anonymised sources     — prevents model echoing internal PDF names into output
• STEP 1 data retrieval  — forces model to list evidence before writing prose,
                           reducing hallucination
• CJK stripping          — removes Korean/Japanese/Chinese text before Stage 2
                           so non-English characters never appear in the output
• Smart truncation       — keeps front 75% + tail 25% of PDF text to capture
                           both executive summary and conclusions/outlook
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

# ── Windows UTF-8 fix ────────────────────────────────────────────────────────
# Windows consoles default to cp1252 which cannot encode many Unicode characters
# produced by LLMs (figure dashes, box-drawing chars, CJK text, etc.).
# Reconfigure stdout/stderr to UTF-8 so print() never raises UnicodeEncodeError.
for _stream in (sys.stdout, sys.stderr):
    try:
        if getattr(_stream, "encoding", "utf-8").lower().replace("-", "") != "utf8":
            _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError):
        pass  # reconfigure not available (Python < 3.7) or stream is not a TTY

# Project root (two levels up from this file: backend/ → PGIM/)
ROOT      = Path(__file__).parent.parent
# Where Stage 1 cached JSON files are stored — one file per PDF
CACHE_DIR = ROOT / "Input_files" / "market_reports" / "cache"
# Where market report PDFs are uploaded by the user
REPORTS_DIR = ROOT / "Input_files" / "market_reports"


# ═══════════════════════════════════════════════════════════════════════════════
# PDF / FILE READING
# ═══════════════════════════════════════════════════════════════════════════════

def _read_pdf(pdf_path: Path, max_pages: int = 50) -> str:
    """
    Extract plain text from a PDF using pypdf (up to max_pages pages).

    Each page is prefixed with a [PAGE N] tag so the LLM knows which page
    a statistic came from — this feeds the page references in the source audit.

    Limitation: pypdf only reads text layer; charts, images, and scanned PDFs
    produce empty or garbled output.  If a report is image-only, Stage 1
    extraction will return little or no data.
    """
    try:
        import pypdf
    except ImportError:
        raise ImportError("pypdf is required for PDF reading.  pip install pypdf")

    reader = pypdf.PdfReader(str(pdf_path))
    pages  = reader.pages[:max_pages]
    parts  = []
    for i, pg in enumerate(pages, 1):
        t = pg.extract_text() or ""
        if t.strip():
            parts.append(f"[PAGE {i}]\n{t.strip()}")
    return "\n\n".join(parts)


def _read_report(path: Path) -> str:
    """Read a report file — PDF or plain text."""
    if path.suffix.lower() == ".pdf":
        return _read_pdf(path)
    return path.read_text(encoding="utf-8", errors="ignore")


def _smart_truncate(text: str, max_chars: int = 14000) -> str:
    """
    Trim text to fit within the LLM's context limit (14,000 chars by default).

    Strategy: keep the first 75% + the last 25% of the text.
    - First 75% captures the executive summary and main findings (front of report)
    - Last 25% captures the outlook, forecasts, and conclusions (tail of report)
    - Middle sections (detailed tables, appendices) are dropped with a marker

    This is better than a simple head-truncation which would cut off the outlook.
    """
    if len(text) <= max_chars:
        return text
    front = int(max_chars * 0.75)
    back  = max_chars - front
    return text[:front] + "\n\n[…middle omitted…]\n\n" + text[-back:]


# ═══════════════════════════════════════════════════════════════════════════════
# LLM HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _strip_thinking_tags(text: str) -> str:
    """
    Remove <think>...</think> blocks that reasoning models emit before their answer.

    Models like deepseek-r1 and qwen3 output their internal chain-of-thought
    wrapped in <think> tags before giving the final answer.  These tags must be
    stripped so the pipeline only processes the actual output, not the reasoning.
    Safe to call on any text — has no effect if no <think> tags are present.
    """
    return re.sub(r'<think>[\s\S]*?</think>', '', text, flags=re.IGNORECASE).strip()


# CJK Unicode ranges: Hangul, CJK Unified Ideographs, Hiragana, Katakana,
# CJK Compatibility Ideographs, CJK Extension A/B, Hangul Jamo, etc.
_CJK_RE = re.compile(
    r'[ᄀ-ᇿ'   # Hangul Jamo
    r'぀-ヿ'    # Hiragana + Katakana
    r'㄰-㆏'    # Hangul Compatibility Jamo
    r'가-힣'    # Hangul Syllables
    r'㐀-䶿'    # CJK Extension A
    r'一-鿿'    # CJK Unified Ideographs
    r'豈-﫿'    # CJK Compatibility Ideographs
    r'\U00020000-\U0002A6DF]+'  # CJK Extension B
)


def _strip_cjk(text: str) -> str:
    """
    Remove Korean, Japanese, and Chinese characters from a string.

    Why this is needed:
    - Market reports for Asian deals (Korea, Japan, China) often contain
      text in the local language mixed with English.
    - pypdf extracts this text as-is, so Stage 1 cache may contain CJK characters.
    - If Stage 2 sees CJK text in the research summary, small LLMs sometimes
      output non-English characters in the final rationale.
    - This function strips all CJK before the text reaches Stage 2, ensuring
      the rationale is always written in English only.

    Consecutive CJK sequences are replaced with a single space so surrounding
    Latin text (numbers, punctuation) stays readable.
    """
    return _CJK_RE.sub(" ", text).strip()


def _ollama_chat(messages: list[dict], base_url: str, model: str,
                 temperature: float = 0.2) -> str:
    """
    Send a chat request to a locally running Ollama instance.

    Uses urllib (no extra dependencies) to POST to the /api/chat endpoint.
    stream=False means Ollama waits for the full response before returning
    (simpler than streaming but holds the connection open for longer runs).
    timeout=600 allows up to 10 minutes for large models on slow hardware.
    """
    import urllib.request
    payload = json.dumps({
        "model":   model,
        "messages": messages,
        "stream":  False,
        "options": {"temperature": temperature},
    }).encode()
    req = urllib.request.Request(
        f"{base_url.rstrip('/')}/api/chat",
        data=payload,
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=600) as resp:
        data = json.loads(resp.read())
    return _strip_thinking_tags(data["message"]["content"].strip())


def _openai_chat(messages: list[dict], model: str, api_key: str,
                 temperature: float = 0.3) -> str:
    """
    Send a chat request to OpenAI's API (gpt-4o, gpt-4o-mini, etc.).

    Requires the `openai` Python package (pip install openai).
    Used when the user selects a cloud model from the sidebar instead of
    a local Ollama model.  Much faster and higher quality than local 3b/7b models
    for investment rationale writing, but costs money per call.
    """
    try:
        import openai as _openai
    except ImportError:
        raise ImportError("openai package required for GPT models.  pip install openai")
    client = _openai.OpenAI(api_key=api_key)
    resp   = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=temperature,
    )
    return _strip_thinking_tags(resp.choices[0].message.content.strip())


def _llm_chat(messages: list[dict], llm_cfg: dict, openai_key: str = "",
              temperature: float = 0.3) -> str:
    """
    Central router — sends a chat request to either Ollama or OpenAI.

    The frontend sidebar injects the user's model choice into llm_cfg before
    calling the pipeline, so this function always receives the correct provider.

    llm_cfg expected shape (from deal config + frontend overrides):
      Ollama:  {"provider": "ollama",  "ollama": {"base_url": ..., "model": ...}}
      OpenAI:  {"provider": "openai",  "openai_model": "gpt-4o"}

    Falls back to Ollama qwen2.5:3b if llm_cfg is empty or missing.
    Prints the provider and model name to the console so the user can see
    which model is running during each pipeline step.
    """
    provider = llm_cfg.get("provider", "ollama")
    if provider == "openai":
        key   = openai_key or llm_cfg.get("openai_api_key", "")
        model = llm_cfg.get("openai_model", "gpt-4o")
        if not key:
            raise ValueError(
                "OpenAI model selected but no API key found. "
                "Set it in the deal config under openai.api_key or via OPENAI_API_KEY env var."
            )
        print(f"  [LLM] OpenAI / {model}")
        return _openai_chat(messages, model, key, temperature)
    else:
        ollama = llm_cfg.get("ollama", {})
        model  = ollama.get("model", "qwen2.5:3b")
        url    = ollama.get("base_url", "http://localhost:11434")
        print(f"  [LLM] Ollama / {model}")
        return _ollama_chat(messages, url, model, temperature)


def _parse_json_from_llm(text: str) -> dict:
    """
    Extract and parse a JSON object from an LLM response.

    LLMs often wrap their JSON in markdown code fences (```json ... ```)
    or add preamble text before the actual JSON.  This function handles both:
    1. Strips ```json ... ``` fences if present
    2. Falls back to finding the first { ... } block if no fences found
    3. Raises json.JSONDecodeError if no valid JSON can be found
    """
    # Remove ```json ... ``` fences
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if m:
        text = m.group(1)
    # Fallback: find first { ... }
    m2 = re.search(r"\{[\s\S]*\}", text)
    if m2:
        text = m2.group(0)
    return json.loads(text.strip())


def _fill(template: str, **kwargs) -> str:
    """
    Safe template fill.  Replaces {key} placeholders without choking on
    curly braces that may appear in the substituted values.
    """
    for key, val in kwargs.items():
        template = template.replace("{" + key + "}", str(val))
    return template


# ═══════════════════════════════════════════════════════════════════════════════
# STAGE 1 — EXTRACT REPORT INSIGHTS
# ═══════════════════════════════════════════════════════════════════════════════

_EXTRACT_SYSTEM = (
    "You are a senior real estate research analyst at an institutional investment firm. "
    "Your job is to extract factual, quantitative market intelligence from property research reports. "
    "Always return valid JSON only — no markdown, no preamble, no explanation. "
    "LANGUAGE RULE: All JSON values must be written in English only. "
    "If the report contains text, names, or statistics in any non-English language "
    "(Korean, Japanese, Chinese, etc.), translate or transliterate them into English "
    "before writing them into the JSON. Never include non-English characters in any field."
)

_EXTRACT_PROMPT = """\
Extract the key real estate market facts from the report text below.

Focus on:
- Geography and property sectors covered
- Supply and demand dynamics (vacancy rates, net absorption, new supply)
- Rental levels and growth trends
- Capital values, investment yields, and transaction volumes
- Key demand drivers (economic, structural, occupier trends)
- Development pipeline and completions schedule
- Market outlook, forecasts, and risks

Write ALL JSON values in English only. Translate any Korean, Japanese, Chinese, or other
non-English text into English before writing it. Never include non-English characters.

Return ONLY a JSON object with these exact keys (use null if data not available):
{
  "country_region": "country or region covered",
  "sectors_covered": ["e.g. logistics, office, industrial"],
  "report_period": "time period e.g. 2025-2026",
  "market_overview": "2-3 sentences summarising the overall market context",
  "supply_demand": "key supply/demand facts — vacancy rates, absorption, completions",
  "rental_trends": "rental rate levels, recent changes, forecast growth",
  "capital_values_transactions": "yield levels, capital value trends, investment volumes",
  "demand_drivers": "main factors driving occupier and investor demand",
  "supply_pipeline": "under-construction and planned development pipeline",
  "market_outlook": "key forecast statements and risks",
  "key_statistics": [
    "Prefix EVERY item with its page number using [p.N] — e.g. '[p.5] Vacancy rate: 3.2% (Q1 2026)'. If you cannot determine the page, write [p.?]."
  ]
}

Report text:
---
REPORT_TEXT_PLACEHOLDER
"""


def extract_report_insights(
    report_path: str | Path,
    llm_cfg: dict,
    openai_key: str = "",
    force_refresh: bool = False,
) -> dict:
    """
    STAGE 1 — Extract structured market facts from a single report PDF.

    Reads the PDF, sends the text to the LLM with _EXTRACT_PROMPT, and
    parses the response into a structured JSON dict.  Results are cached
    to disk so the same unchanged PDF is never re-processed.

    Cache invalidation:  cache key = MD5(filename + file_size + mtime).
    If the PDF changes (re-uploaded, updated), the hash changes automatically
    and the file is re-extracted on the next run.

    force_refresh=True skips the cache check and always re-runs the LLM.
    Use this if the extraction model was recently changed or the cache is bad.

    Returns a dict:
    {
        "source_file":  "filename.pdf",
        "source_path":  "/full/path/to/file.pdf",
        "extracted_at": "2026-06-04T12:00:00",
        "insights": {
            "market_overview": "...",
            "supply_demand": "...",
            "rental_trends": "...",
            "capital_values_transactions": "...",
            "demand_drivers": "...",
            "supply_pipeline": "...",
            "market_outlook": "...",
            "key_statistics": ["[p.5] Vacancy: 3.2%", ...]
        }
    }
    Returns {} if the file cannot be read or the LLM call fails.
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    p = Path(report_path)

    # Cache key = filename + file size + mtime — re-extracts automatically if file changes
    try:
        stat = p.stat()
        cache_seed = f"{p.name}:{stat.st_size}:{stat.st_mtime}"
    except OSError:
        cache_seed = p.name
    cache_key  = hashlib.md5(cache_seed.encode()).hexdigest()[:10]
    cache_file = CACHE_DIR / f"{p.stem}_{cache_key}.json"

    # Serve from cache if available and not force-refreshing
    if cache_file.exists() and not force_refresh:
        print(f"  [cache]     {p.name}")
        return json.loads(cache_file.read_text(encoding="utf-8"))

    print(f"  [extracting] {p.name} ...")
    try:
        raw_text = _read_report(p)
    except Exception as e:
        print(f"  [warning]   Could not read {p.name}: {e}")
        return {}

    if not raw_text.strip():
        print(f"  [warning]   No extractable text in {p.name}")
        return {}

    # Trim to fit within model context: first 75% (exec summary) + last 25% (outlook)
    truncated = _smart_truncate(raw_text, max_chars=14000)

    prompt = _EXTRACT_PROMPT.replace("REPORT_TEXT_PLACEHOLDER", truncated)

    try:
        response = _llm_chat(
            [
                {"role": "system", "content": _EXTRACT_SYSTEM},
                {"role": "user",   "content": prompt},
            ],
            llm_cfg=llm_cfg,
            openai_key=openai_key,
            temperature=0.1,
        )
        insights = _parse_json_from_llm(response)
    except Exception as e:
        print(f"  [warning]   LLM extraction failed for {p.name}: {e}")
        insights = {"extraction_error": str(e), "raw_excerpt": raw_text[:1500]}

    result = {
        "source_file":  p.name,
        "source_path":  str(p),
        "extracted_at": datetime.now().isoformat(),
        "insights":     insights,
    }
    cache_file.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"  [cached]    -> {cache_file.name}")
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# STAGE 2 — GENERATE INVESTMENT RATIONALE
# ═══════════════════════════════════════════════════════════════════════════════


# ── Rationale generation prompts (Call 1) ────────────────────────────────────
# _RATIONALE_SYSTEM sets the persona, writing rules, data integrity constraints,
# style guidelines, and section title quality standards.  It is sent as the
# system message so its rules take highest precedence over the user prompt.
#
# _RATIONALE_PROMPT is the user message.  It contains placeholder tokens
# (ALL_CAPS) that are replaced at runtime with the real deal config values and
# the anonymised market research summary.  The anonymisation is critical — the
# LLM sees source data labelled "Research Report 1 / 2 / …", never real PDF
# filenames, so it cannot echo them into the prose.
_RATIONALE_SYSTEM = """\
You are a senior investment professional at a global institutional real estate fund.
You write investment committee memos — tight, authoritative, and wholly grounded in evidence.

━━ DATA INTEGRITY (non-negotiable) ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Only write statistics and facts that appear explicitly in the Market Research Summary.
  If a figure is not there, omit it entirely — never estimate, round, or extrapolate.
• PDF charts and images do not extract to text reliably.
  Ignore any isolated numbers that lack clear written context in the research.
• If the research does not cover a specific sub-point, skip that sub-point and develop
  another angle that IS supported by the data. Never signal the gap to the reader —
  sentences like "while vacancy data is not available", "specific figures are not provided",
  or "data is limited" must never appear in the output. Omit silently, move on.
• Never use general market knowledge. Every claim, statistic, and assertion must come
  exclusively from the Market Research Summary or the Subject Property configuration.
  If a data point cannot be traced to one of those two sources, omit it entirely.
  Do not substitute general knowledge when research data is absent — simply omit the point.
• Every number you write (percentage, rate, area, price, volume, yield, index value)
  must appear word-for-word or digit-for-digit in the Market Research Summary or Deal Config.
  Do not round, adjust, combine, or derive figures — use only what is explicitly stated.
• NEVER write bracketed placeholders such as [X%], [X.X%], [Y], [Z units], [B]%, [D] USD billion,
  [p.?], [p.N], or any similar template variable. If you do not have the exact figure from the
  research, omit that data point entirely — do not substitute a placeholder under any circumstances.
• Every policy name, regulation, government initiative, or named programme you reference
  must be explicitly named in the Market Research Summary. Do not cite policies from memory.

━━ VOICE AND STYLE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Write declarative statements. State the conclusion first, then the evidence.
• No attribution hedges or source references anywhere in the body text — do not write
  "according to", "the report states", "data shows", "it is noted that", or any similar
  phrase. Do not mention any report filename, report title, or source document by name
  anywhere in the three sections. Just state the fact. Sources are recorded in the
  audit JSON only.
• Banned transitions: "additionally", "furthermore", "moreover", "in addition",
  "it is worth noting", "it should be noted", "lastly", "to summarise".
• No bullet lists. Continuous prose only.
• Active voice. Present tense for market conditions; past tense for completed transactions.
• Each section runs 150–250 words in total across 2–3 paragraphs. Every paragraph makes
  one clear point — state the market condition, explain the mechanism, connect it to the
  investment case. No padding, no vague generalisations. Be concise and evidence-dense.
• You MUST complete every section you start in full. Do not stop early. Do not add any
  heading, JSON block, or separator text after finishing your final section.
• Write between 3 and 5 sections. Add a 4th or 5th only when there is a genuinely
  distinct, data-supported investment angle. A tight 3-section memo beats a padded 5.

━━ LANGUAGE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• The entire output must be written in English only.
• If the research contains statistics, place names, property names, or any text in a
  non-English language (Korean, Japanese, Chinese, etc.), translate or transliterate it
  into English before using it. Never include non-English characters in the prose.

━━ EVIDENCE RULES ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Every opinion, forecast, or investment conclusion must be directly supported by a
  specific figure, percentage, named trend, or fact from the Market Research Summary.
  Plain assertions with no data anchor are not permitted — if you cannot cite a number
  or named fact, do not make the claim.
• No single statistic, figure, or named fact may appear more than 3 times across the
  entire rationale. Each key data point has a primary section where it is introduced and
  developed; do not repeat it mechanically across sections. Spread the evidence.

━━ SECTION TITLES ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Write the title directly after the section number (## 1. / ## 2. / ## 3. / ## 4. / ## 5.)
— no brackets, no angle brackets, no quotes, no period.  Length: 6–9 words.

A title must do two things at once:
  (A) name a specific, evidenced market condition or asset characteristic
  (B) imply the investment consequence — why it matters to returns or capital value

RULES:
  • Do NOT include the property name, deal name, or location name in the title.
  • Do NOT use a colon to separate a category from a location (e.g. "Market Cycle: Property in City" is wrong).
  • The title must stand alone as an investment thesis statement.

STRONG titles — specific mechanism + investment implication (DO NOT copy these):
  "Structurally Under-Supplied Market Entering a Landlord-Favoured Cycle"
  "Prime Last-Mile Node Commanding Seoul's Fastest-Growing Demand Catchment"
  "Defensive Income Profile Anchored by Long-WALE Institutional Covenants"
  "Tight Development Pipeline Sustaining Above-Inflation Rental Reversion"

WEAK titles — never write anything that follows these patterns:
  ✗ Generic:     "Strong Fundamentals Supporting Logistics Investment Returns"
  ✗ Categorical: "Market Overview" / "Location Analysis" / "Deal Highlights"
  ✗ Vague:       "Positive Outlook for Industrial Property in the Region"
  ✗ Hedged:      "Potential Upside Amid Evolving Market Conditions"
  ✗ Property+Location: "Market Cycle Position: Property Name in City"

Vocabulary to draw from (use what fits the data):
  yield compression, structural undersupply, rental reversion, last-mile logistics,
  e-commerce penetration, constrained land release, institutional-grade covenant,
  WALE, net absorption, anchor tenant, leasehold profile, supply pipeline, cap rate,
  occupier demand, submarket tightness, development moratorium, value-add angle

Derive every title from the specific data in this prompt — country, submarket,
asset class, and the actual numbers from the research.  Original titles only."""

_RATIONALE_PROMPT = """\
Write an investment committee rationale for the subject property below.
Use only the Market Research Summary and Subject Property details as your evidence base.

═══ SUBJECT PROPERTY ════════════════════════════════════════════
Deal Name     : DEAL_NAME
Property Name : PROPERTY_NAME
Address       : ADDRESS
Asset Class   : ASSET_CLASS
Asset Type    : ASSET_TYPE
Location      : LOCATION
Quality/Grade : QUALITY
GFA           : GFA_SF GFA_UNIT
Country       : COUNTRY_NAME
EXTRA_FIELDS
═════════════════════════════════════════════════════════════════

═══ MARKET RESEARCH SUMMARY ════════════════════════════════════
COMBINED_INSIGHTS
════════════════════════════════════════════════════════════════

═══ ANALYST NOTES ══════════════════════════════════════════════
ANALYST_NOTES
════════════════════════════════════════════════════════════════
REFINEMENT_BLOCK

STEP 1 — RETRIEVE DATA BEFORE WRITING ANYTHING
Do not write any prose yet.

Read through the entire Market Research Summary and Subject Property details.
Decide what distinct categories of investment-relevant information are actually present
in the research — do not use a fixed list of categories. Name the categories yourself
based on what the research actually contains.

GUIDANCE (examples of typical categories — not an exhaustive list, not all required):
  • Market cycle position — e.g. vacancy rates, net absorption, rental growth rate,
    cap rate or yield levels, investment volume, capital value trends, market outlook
  • Submarket or location dynamics — e.g. submarket vacancy vs city average,
    infrastructure projects, location cluster or precinct performance, rental premiums
    by district, migration of occupiers across submarkets
  • Demand drivers — e.g. occupier sector trends (e-commerce, 3PL, tech, financial
    services), structural demand shifts, ESG or green-building requirements, lease
    demand and pre-commitment activity, floor-plate or specification preferences
  • Supply pipeline — e.g. completions schedule, under-construction volume, land
    scarcity or development moratorium, GLS or planning approval pipeline,
    construction cost pressures
  • Deal-specific details — e.g. tenancy profile, WALE, lease expiry schedule,
    anchor tenant, occupancy rate, passing rent vs market rent, pricing, GFA,
    remaining leasehold, reversionary potential
  • Other categories the research supports — e.g. investment market activity,
    ESG premium evidence, structural sector shifts with their own data set

Aim for 3–6 categories. Merge closely related data into one rather than splitting
into many thin categories. You may use fewer or more depending on what the research
actually contains.

For each category you identify:
  — Name the category clearly so you can reference it in STEP 2.
  — Quote every specific figure, percentage, named fact, and direct data point found.
  — Only list data explicitly stated in the Market Research Summary or Deal Config.
    No inference, no general knowledge, no paraphrasing that changes the numbers.
  — Write "— none found —" if a category has no data in either source.

STEP 2 — PLAN SECTIONS
Using only the data retrieved in STEP 1:
  a) Decide how many sections to write (3 default; 4–5 only if a distinct evidenced
     angle exists that cannot fit in the first 3).
  b) Assign each category to the section where it will serve as primary evidence.
     Categories with "— none found —" must not drive any section — omit that angle entirely.
  c) State in one sentence the investment thesis each section will argue.

STEP 3 — WRITE THE SECTIONS
For each section, replace the [write a title here...] instruction on the ## line with your
actual title — 6–9 words, no property name, no location name, no colon separating a category
from a location. Just the title itself, nothing else on that line.
Do NOT mention any report name, file name, or source document anywhere in the sections.
Do NOT use phrases like "according to", "the report indicates", or any attribution.
Write every fact as a direct, unattributed statement. All source tracking goes in the audit JSON only.

Use only the data points you listed in STEP 1. If a figure or named fact was not listed
in STEP 1, it must not appear in the prose — do not add new data at the writing stage.

Each section: 150–250 words total across 2–3 paragraphs. Do not exceed 250 words per section.
Pack in specific figures and data points — concise and evidence-dense beats long and vague.

CORE SECTIONS (always include these three unless the research clearly supports more):

## 1. [write a 6–9 word investment thesis title here — derived from your STEP 1 market-cycle data]
Draw from your market-cycle and supply-pipeline categories (whatever you named them in STEP 1).
Cover: supply/demand imbalance; vacancy trajectory and landlord pricing power; rental growth
trend and momentum; capital value and yield context; development pipeline constraints;
market outlook and entry timing thesis.

## 2. [write a 6–9 word investment thesis title here — derived from your STEP 1 location/demand data]
Draw from your location/submarket and demand-driver categories. Cover: why this submarket
commands premium occupier demand; proximity to key infrastructure; catchment area and tenant
base depth; competitive set positioning; regulatory or land-scarcity factors that entrench
defensibility.

## 3. [write a 6–9 word investment thesis title here — derived from your STEP 1 deal-specific data]
Draw from your deal-specific category and supporting data from the other categories retrieved
in STEP 1. Cover: tenancy profile, lease structure, rent reversionary potential; deal pricing
vs market; capital appreciation thesis; key risks and mitigants; why compelling at this point
in the cycle.

OPTIONAL ADDITIONAL SECTIONS (add only if your STEP 1 categories strongly support a distinct angle):

## 4. [write a 6–9 word investment thesis title here — only if a distinct data-supported angle exists]
Examples: a specific ESG or green-building premium quantified in the research; a supply-side
development moratorium with named policy evidence; a demand-side structural shift (e-commerce
penetration, 3PL consolidation) with its own data set.

## 5. [write a 6–9 word investment thesis title here — rare; most deals do not need a 5th section]

STEP 4 — VERIFY BEFORE RETURNING
Before returning your response, check every section against this list.
Fix any failure before returning — do not return output that fails a check.

  [ ] RETRIEVAL    — Every data point in the prose was listed in STEP 1. Any figure or
                     named fact that was NOT in your STEP 1 list must be deleted from the prose.
  [ ] LANGUAGE     — No non-English characters anywhere. All foreign terms translated.
  [ ] DATA ANCHOR  — Every opinion, forecast, or investment conclusion contains or directly
                     follows a specific figure, percentage, or named fact. No plain assertions.
  [ ] SOURCE CHECK — Every number (vacancy rate, cap rate, rental rate, yield, price, volume,
                     GFA, growth %, index value) must appear explicitly in the Market Research
                     Summary or Deal Config. Delete any figure you cannot locate there.
  [ ] POLICY CHECK — Every policy name, regulation, government initiative, or named programme
                     must be explicitly named in the Market Research Summary. Delete any policy
                     reference you cannot locate there.
  [ ] REUSE LIMIT  — No single statistic or named fact appears more than 3 times in total
                     across all sections. If it does, remove the duplicate and replace with
                     a different supporting data point or omit that sentence.
  [ ] WORD COUNT   — Each section is 150–250 words. Trim any section that exceeds 250 words.
  [ ] ATTRIBUTION  — No sentence contains "according to", "the report", "data shows",
                     or any source reference. No report filename anywhere.
  [ ] TRANSITIONS  — No banned transition words anywhere in the output.

No fabrication. No data gaps flagged in prose. Use only research and deal config data.
Write ALL sections completely. Stop immediately after the last paragraph of your final section.
Do not add any headings, JSON, commentary, or separator after the final section.
"""


def _merge_insights(extracted_reports: list[dict], anonymize: bool = False) -> str:
    """
    Combine cached insights from multiple reports into one text block for the prompt.

    Called twice in Stage 2 with different anonymize settings:

    Call 1 (prose writing)  — anonymize=True
        Source labels become "Research Report 1", "Research Report 2", etc.
        The LLM physically cannot echo real PDF filenames into the prose
        because it never sees them.  This keeps client-facing memos clean.

    Call 2 (source audit)   — anonymize=False
        Source labels show the real PDF filename (e.g. "AEWResearch_Korea.pdf").
        The audit LLM needs the real names to populate Source_Audit.xlsx correctly.

    CJK characters are stripped from every field before the text is included,
    ensuring no Korean/Japanese/Chinese text reaches Stage 2.
    """
    parts = []
    for idx, rpt in enumerate(extracted_reports):
        src = rpt.get("source_file", "Unknown")
        ins = rpt.get("insights", {})
        if not ins:
            continue

        label = f"Research Report {idx + 1}" if anonymize else f"Source: {src}"
        lines = [label]

        field_labels = [
            ("market_overview",             "Market Overview"),
            ("supply_demand",               "Supply & Demand"),
            ("rental_trends",               "Rental Trends"),
            ("capital_values_transactions", "Capital Values & Transactions"),
            ("demand_drivers",              "Demand Drivers"),
            ("supply_pipeline",             "Supply Pipeline"),
            ("market_outlook",              "Market Outlook"),
        ]
        for key, label in field_labels:
            val = ins.get(key)
            if val:
                lines.append(f"{label}: {_strip_cjk(str(val))}")

        stats = ins.get("key_statistics", [])
        if stats:
            lines.append("Key Data Points: " + " | ".join(
                _strip_cjk(str(s)) for s in stats[:12]))

        parts.append("\n".join(lines))

    if not parts:
        return "(No market research provided — rationale will be based on property specifics only.)"
    return "\n\n─────\n\n".join(parts)


def _clean_rationale_body(text: str) -> str:
    """
    Remove any trailing audit/JSON sections the LLM sometimes appends after the prose.

    Some models (especially smaller ones) ignore the instruction to stop after the
    last section and append headings like "## Source Audit", "## Citations", or
    a raw JSON block.  This function detects those headings and truncates the text
    at the first one, returning only the clean prose sections.
    """
    audit_hdr = re.compile(
        r'^#{1,4}\s*(source\s*audit|audit\s*json|audit|citations?|references?|json)\s*$',
        re.IGNORECASE,
    )
    lines = text.splitlines()
    clean = []
    for line in lines:
        if audit_hdr.match(line.strip()):
            break          # stop collecting at first audit heading
        clean.append(line)
    # Drop trailing blank lines
    while clean and not clean[-1].strip():
        clean.pop()
    return "\n".join(clean)


# ── Audit-call prompts (separate second LLM call) ─────────────────────────────

_AUDIT_SYSTEM = (
    "You are a citation auditor for an institutional real estate investment memo. "
    "You identify specific claims and match each one to its source. "
    "Return ONLY a valid JSON array — no markdown, no preamble, no explanation."
)

_AUDIT_PROMPT_TEMPLATE = """\
Audit the investment rationale below against the market research sources provided.
You must capture a row for EVERY auditable item. Do not skip anything that contains a
specific claim, figure, official statement, policy reference, or named data point.

Audit ALL of the following — every instance must appear as a separate row:

DATA QUOTES — every quantitative figure cited in the rationale:
  - Vacancy rate, occupancy rate, take-up, net absorption, pre-commitment rate
  - Cap rate, NPI yield, net yield, passing rent, market rent, rental reversion
  - Price, price psf, GFA, site area, transaction volume, deal count
  - Completions, pipeline supply, under-construction quantum
  - Growth rate, CAGR, index level, percentage change, basis points move
  - Any other number, ratio, or metric used to support an argument

NAMED QUALITATIVE CLAIMS — every specific non-numeric assertion tied to a named fact:
  - Named trend, structural shift, or outlook statement tied to a direction or timeframe
  - Named location, precinct, infrastructure node, or submarket used to justify positioning
  - Named occupier sector, tenant type, or demand driver (e.g. "e-commerce", "3PL", "data centres")
  - Named building, development, or asset used as a comparable or benchmark
  - Any ESG certification, green building standard, or sustainability requirement cited

Only skip sentences that are purely structural connective tissue with zero data content
(e.g. "The asset benefits from strong fundamentals" with no figures, names, or specific claims).
Every number and every named fact in the rationale MUST appear as a row.

INVESTMENT RATIONALE:
────────────────────────────────────────────────────────────────
RATIONALE_BODY_PLACEHOLDER
────────────────────────────────────────────────────────────────

RESEARCH SOURCES (read these to match claims):
────────────────────────────────────────────────────────────────
SOURCE_DATA_PLACEHOLDER
────────────────────────────────────────────────────────────────

VALID source_file VALUES — copy one of these EXACTLY (character-for-character):
VALID_FILENAMES_PLACEHOLDER
  - Deal Config

Do NOT invent filenames, abbreviate them, or write "Research Report N".
Do NOT use "General Knowledge" — every claim must trace to one of the research reports or Deal Config.
Every source_file entry must be one of the exact strings listed above.

Return ONLY a JSON array, one entry per specific claim:
[
  {
    "section_num": 1,
    "section_title": "exact title of that section from the rationale",
    "claim": "specific assertion, data point, or official statement from the rationale text",
    "source_file": "one of the exact valid values listed above",
    "page_ref": "page number where this data appears, e.g. 'p.5' — use 'p.?' if uncertain, null if Deal Config",
    "supporting_text": "verbatim or near-verbatim passage from the research — null if Deal Config",
    "citation_type": "Verbatim | Paraphrased | Deal Config"
  }
]

citation_type guide:
  Verbatim    — the rationale quotes the source almost word-for-word
  Paraphrased — the rationale restates a figure or finding in different words
  Deal Config — the claim originates from the deal configuration / subject property data
"""


# ── Source audit Excel writer ──────────────────────────────────────────────────

def _cross_check_claim(claim_entry: dict, extracted_reports: list[dict]) -> str:
    """
    Python-level verification: check whether a citation's supporting text
    actually appears in the cached Stage 1 extract for that source file.

    This is a fast automated check — it does NOT re-read the original PDF.
    It searches the cached extract text (which is the LLM-summarised version,
    not the full PDF text), so results are approximate:

    ✓  Found in cached extract         — 4-word sliding window matched
    ⚠  Not found in cached extract     — no match; reviewer must check original PDF
    ⚠  No supporting text provided     — audit LLM left supporting_text blank
    ⚠  Source file not in selected reports — filename not found in any loaded report

    Matching strategy: 4-word sliding window — any 4 consecutive words from the
    supporting text must appear somewhere in the cached extract text.
    This is loose enough to handle minor paraphrasing but strict enough to
    catch complete fabrications.

    Two passes:
    1. Exact filename match — looks for source_file == rpt["source_file"]
    2. Fuzzy match — handles truncation differences (e.g. "AEWResearch..." vs full name)
    """
    src_file = claim_entry.get("source_file", "")
    support  = claim_entry.get("supporting_text") or ""
    ctype    = claim_entry.get("citation_type", "")

    # Deal Config citations come from the config file, not a research report — skip
    if ctype in ("Deal Config", "General Knowledge") or src_file in ("Deal Config", "General Knowledge"):
        return ctype

    if not support.strip():
        return "⚠  No supporting text provided"

    def _flatten(rpt: dict) -> str:
        """Flatten all cached insight fields into a single lowercase string for matching."""
        ins = rpt.get("insights", {})
        txt = " ".join(str(v) for v in ins.values() if v and not isinstance(v, list))
        txt += " " + " ".join(str(s) for s in (ins.get("key_statistics") or []))
        return txt.lower()

    def _text_match(support: str, full_text: str) -> bool:
        """Return True if any 4-word window from support appears in full_text."""
        words  = support.lower().split()
        chunks = [" ".join(words[i:i+4]) for i in range(max(len(words) - 3, 1))]
        return any(c in full_text for c in chunks)

    # Pass 1 — exact filename match
    for rpt in extracted_reports:
        if rpt.get("source_file", "") == src_file:
            full_text = _flatten(rpt)
            if _text_match(support, full_text):
                return "✓  Found in cached extract"
            return "⚠  Not found in cached extract — verify against original PDF"

    # Pass 2 — fuzzy filename match (handles truncation / casing differences)
    src_lower = src_file.lower()
    for rpt in extracted_reports:
        rpt_src = rpt.get("source_file", "").lower()
        if rpt_src and (src_lower in rpt_src or rpt_src in src_lower):
            full_text = _flatten(rpt)
            if _text_match(support, full_text):
                return "✓  Found in cached extract (fuzzy filename match)"
            return "⚠  Not found in cached extract — verify against original PDF"

    return "⚠  Source file not in selected reports"


def _write_source_audit_excel(
    audit_entries: list[dict],
    extracted_reports: list[dict],
    out_path: Path,
    subject_cfg: dict,
) -> None:
    """
    Write Source_Audit.xlsx — a formatted Excel table for human citation review.

    Each row = one specific claim from the investment rationale, with:
    - Which section and section title it came from
    - The exact claim text
    - The source PDF filename and page reference
    - The LLM's verbatim supporting quote from the research
    - Citation type (Verbatim / Paraphrased / Deal Config)
    - Backend cross-check status (auto-verified against the cached extract)
    - Blank columns for the human reviewer to mark CONFIRMED / INCORRECT / CANNOT VERIFY

    Row colour coding:
    - Red    = claim not found in cached extract → must verify against original PDF
    - Orange = no supporting text provided by audit LLM → review needed
    - White/Grey = alternating rows, auto-verified or Deal Config

    Columns
    -------
    #  |  Section  |  Section Title  |  Claim / Data Point  |  Source Report  |
    Supporting Text (LLM quote)  |  Citation Type  |  Backend Cross-Check  |
    Human Validation  |  Reviewer Notes
    """
    import openpyxl
    from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side)
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Source Audit"

    # ── Palette ───────────────────────────────────────────────────────────────
    NAVY    = "FF1A3A5C"
    WHITE   = "FFFFFFFF"
    LGREY   = "FFF2F5F8"
    DGREY   = "FFE0E7EF"
    GREEN   = "FFD6F0DC"
    ORANGE  = "FFFFF3CD"
    RED_BG  = "FFFDE8E8"
    AMBER   = "FFFF8C00"
    DARK    = "FF1A1A1A"

    def _hdr_font(bold=True):
        return Font(name="Calibri", bold=bold, color=WHITE, size=10)
    def _body_font(bold=False, color=DARK):
        return Font(name="Calibri", bold=bold, color=color, size=9)
    def _fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)
    def _align(h="left", v="top", wrap=True):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    def _thin_border():
        s = Side(style="thin", color="FFD0D0D0")
        return Border(left=s, right=s, top=s, bottom=s)

    # ── Title row ─────────────────────────────────────────────────────────────
    prop_name = subject_cfg.get("property_name", subject_cfg.get("deal_name", "Property"))
    ws.merge_cells("A1:K1")
    title_cell = ws["A1"]
    title_cell.value = (f"Source Audit — {prop_name}  "
                        f"({subject_cfg.get('asset_class','').title()}, "
                        f"{subject_cfg.get('country_name','')})  |  "
                        f"Generated {datetime.now().strftime('%d %b %Y')}")
    title_cell.font      = Font(name="Calibri", bold=True, color=WHITE, size=11)
    title_cell.fill      = _fill(NAVY)
    title_cell.alignment = _align("left", "center", wrap=False)
    ws.row_dimensions[1].height = 22

    # Instruction row
    ws.merge_cells("A2:K2")
    note_cell = ws["A2"]
    note_cell.value = (
        "VALIDATION INSTRUCTIONS:  Review every row where Backend Cross-Check = "
        "⚠  Not found in cached extract.  Open the original PDF and verify the claim.  "
        "Mark Human Validation as CONFIRMED / INCORRECT / CANNOT VERIFY, and add notes."
    )
    note_cell.font      = Font(name="Calibri", italic=True, color="FF555555", size=8)
    note_cell.fill      = _fill("FFFFF8DC")
    note_cell.alignment = _align("left", "center", wrap=True)
    ws.row_dimensions[2].height = 28

    # ── Column headers ─────────────────────────────────────────────────────────
    HEADERS = [
        "#",
        "Section",
        "Section Title",
        "Claim / Data Point Used",
        "Source Report",
        "Page\nRef",
        "LLM Supporting Text\n(verbatim quote provided by LLM)",
        "Citation Type",
        "Backend Cross-Check\n(auto-verified against cached extract)",
        "Human Validation\n(CONFIRMED / INCORRECT / CANNOT VERIFY)",
        "Reviewer Notes",
    ]
    COL_WIDTHS = [4, 9, 38, 45, 32, 8, 45, 14, 34, 20, 28]

    for col_idx, (hdr, width) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
        cell = ws.cell(row=3, column=col_idx, value=hdr)
        cell.font      = _hdr_font()
        cell.fill      = _fill(NAVY)
        cell.alignment = _align("center", "center", wrap=True)
        cell.border    = _thin_border()
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[3].height = 36

    # ── Data rows ──────────────────────────────────────────────────────────────
    for row_num, entry in enumerate(audit_entries, 1):
        row_idx    = row_num + 3   # rows 4+
        cross_chk  = _cross_check_claim(entry, extracted_reports)
        ctype      = entry.get("citation_type", "")

        # Row background
        if "⚠  Not found" in cross_chk:
            row_fill = _fill(RED_BG)
        elif "⚠  No supporting" in cross_chk:
            row_fill = _fill(ORANGE)
        elif row_num % 2 == 0:
            row_fill = _fill(LGREY)
        else:
            row_fill = _fill(WHITE)

        row_data = [
            row_num,
            f"Section {entry.get('section_num', '?')}",
            entry.get("section_title", ""),
            entry.get("claim", ""),
            entry.get("source_file", ""),
            entry.get("page_ref") or "—",
            entry.get("supporting_text") or "",
            ctype,
            cross_chk,
            "",   # Human Validation — blank for reviewer
            "",   # Reviewer Notes — blank for reviewer
        ]

        for col_idx, value in enumerate(row_data, 1):
            cell            = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill       = row_fill
            cell.border     = _thin_border()
            cell.alignment  = _align(wrap=True)

            # Special formatting  (cols shifted +1 after adding Page Ref at col 6)
            if col_idx == 1:  # row number
                cell.alignment = _align("center", "top", wrap=False)
                cell.font      = _body_font(color="FF888888")
            elif col_idx == 6:  # page reference
                cell.alignment = _align("center", "top", wrap=False)
                cell.font      = _body_font(bold=True, color="FF1A3A5C")
            elif col_idx == 8:  # citation type
                color = AMBER if ctype in ("Inferred", "General Knowledge") else DARK
                cell.font = _body_font(bold=(ctype == "Inferred"), color=color)
            elif col_idx == 9:  # cross-check status
                if "✓" in cross_chk:
                    cell.font = _body_font(color="FF1A7A2A")
                elif "⚠" in cross_chk:
                    cell.font = _body_font(bold=True, color="FFC00000")
                else:
                    cell.font = _body_font(color="FF555555")
            elif col_idx == 10:  # human validation (blank but highlighted)
                cell.fill = _fill(GREEN)
                cell.font = _body_font(color="FF1A7A2A")
            else:
                cell.font = _body_font()

        ws.row_dimensions[row_idx].height = 52

    # ── Summary stats at the bottom ───────────────────────────────────────────
    total      = len(audit_entries)
    n_verified = sum(1 for e in audit_entries
                     if "✓" in _cross_check_claim(e, extracted_reports))
    n_warn     = sum(1 for e in audit_entries
                     if "⚠" in _cross_check_claim(e, extracted_reports))
    n_inferred = sum(1 for e in audit_entries
                     if e.get("citation_type") in ("Inferred", "General Knowledge"))

    sum_row = total + 4 + 2
    ws.merge_cells(f"A{sum_row}:K{sum_row}")
    sum_cell = ws[f"A{sum_row}"]
    sum_cell.value = (
        f"SUMMARY:  {total} citations total  |  "
        f"{n_verified} verified in cached extract  |  "
        f"{n_warn} require manual PDF verification  |  "
        f"{n_inferred} inferred / general knowledge (review carefully)"
    )
    sum_cell.font      = Font(name="Calibri", bold=True, color=WHITE, size=9)
    sum_cell.fill      = _fill(NAVY)
    sum_cell.alignment = _align("left", "center", wrap=False)
    ws.row_dimensions[sum_row].height = 18

    # Freeze panes below header
    ws.freeze_panes = "A4"

    wb.save(str(out_path))
    print(f"  [source audit] -> {out_path.name}")


# ── Parsing helper ─────────────────────────────────────────────────────────────

def _parse_audit_json(raw_audit: str) -> list[dict]:
    """
    Extract and parse the source audit JSON array from the LLM's raw response.

    The audit LLM is instructed to return a bare JSON array, but may still wrap
    it in markdown fences or add preamble text.  This function handles both cases:
    1. Strips ```json ... ``` fences if present
    2. Finds the first [ ... ] array block
    3. Returns [] if no valid JSON array can be parsed (triggers fallback)
    """
    text = raw_audit.strip()
    # Strip markdown fences
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if m:
        text = m.group(1).strip()
    # Find JSON array
    m2 = re.search(r"\[[\s\S]*\]", text)
    if m2:
        text = m2.group(0)
    try:
        result = json.loads(text)
        if isinstance(result, list):
            return result
    except Exception as e:
        print(f"  [warning] Could not parse source audit JSON: {e}")
    return []


def _fallback_citations_from_text(rationale_body: str) -> list[dict]:
    """
    Last-resort fallback if the audit LLM fails to return valid JSON.

    Scans the rationale prose for inline [SOURCE: filename] markers and
    builds a minimal citation list from them.  These markers are not normally
    present in the output (the main pipeline uses a separate audit call), but
    some model configurations may produce them as a side effect.
    Returns [] if no markers are found.
    """
    entries = []
    # Find [SOURCE: something] markers and the preceding sentence
    pattern = re.compile(
        r'([^.\n]{10,200}?)\s*\[SOURCE:\s*([^\]]+)\]',
        re.DOTALL,
    )
    for m in pattern.finditer(rationale_body):
        claim     = m.group(1).strip().rstrip(",;:")
        src_raw   = m.group(2).strip()
        entries.append({
            "section_num":    "?",
            "section_title":  "",
            "claim":          claim,
            "source_file":    src_raw,
            "supporting_text": None,
            "citation_type":  "Paraphrased",
        })
    return entries


def _build_input_summary(extracted_reports: list) -> str:
    """Build the 'Input Summary' section from the Stage-1 extracted report facts.

    Summarises, per input PDF, the key points and what they convey for this
    investment — the evidence base that shapes the rationale that follows. Built
    directly from the already-extracted structured facts (no extra LLM call).
    """
    reports = [r for r in (extracted_reports or []) if r.get("source_file")]
    if not reports:
        return ""

    out = [
        "## Input Summary\n",
        "_Key points drawn from the market report(s) provided, and what they convey "
        "for this investment — the evidence base for the rationale that follows._\n",
    ]
    for r in reports:
        src     = r.get("source_file", "report")
        sectors = r.get("sectors_covered")
        sectors = ", ".join(sectors) if isinstance(sectors, list) else (sectors or "")
        meta    = " · ".join(x for x in (r.get("country_region") or "", sectors,
                                         r.get("report_period") or "") if x)
        out.append(f"### {src}" + (f"  \n_{meta}_" if meta else ""))

        overview = (r.get("market_overview") or "").strip()
        if overview:
            out.append(overview)

        # "What it conveys" — the decision-relevant angles, when present.
        for label, key in (
            ("Demand drivers",            "demand_drivers"),
            ("Supply & demand",           "supply_demand"),
            ("Rents",                     "rental_trends"),
            ("Capital values / yields",   "capital_values_transactions"),
            ("Development pipeline",       "supply_pipeline"),
            ("Outlook & risks",           "market_outlook"),
        ):
            val = (r.get(key) or "").strip()
            if val:
                out.append(f"- **{label}:** {val}")

        stats = [s.strip() for s in (r.get("key_statistics") or [])
                 if isinstance(s, str) and s.strip()][:3]
        if stats:
            out.append("- **Notable figures:** " + "; ".join(stats))
        out.append("")   # blank line between reports

    out.append("---\n")
    return "\n".join(out) + "\n"


def generate_rationale(
    extracted_reports: list[dict],
    subject_cfg: dict,
    llm_cfg: dict,
    openai_key: str = "",
    analyst_notes: str = "",
    refinement_notes: str = "",
) -> tuple[str, list]:
    """
    Two-call LLM pipeline that produces the investment rationale and its
    source audit.

    WHY TWO SEPARATE CALLS?
    ─────────────────────────────────────────────────────────────────────────
    A single call that sees the real PDF filenames will naturally echo them
    into the prose — producing sentences like "According to AEWResearch_
    AsiaPacific_South-Korea-Logistics_WEBSITE.pdf, vacancy is 3%…", which is
    unprofessional in a client-facing investment memo and exposes internal
    source filenames.

    Splitting into two calls solves this cleanly:

    Call 1 — Prose writing (anonymised sources)
        The research data is passed with source labels replaced by generic
        names ("Research Report 1", "Research Report 2", …) via
        _merge_insights(anonymize=True).  The LLM physically cannot echo a
        real filename because it never sees one.  Output: clean 3-section
        investment rationale markdown.

    Call 2 — Source audit JSON (real filenames visible)
        The finished prose and the research data WITH real filenames are sent
        to a separate LLM call whose only job is to match each specific claim
        in the prose back to its real source PDF and page number.  Output: a
        JSON array of citation entries that populates Source_Audit.xlsx.
        This JSON is never shown in the front-end; it stays in the Excel only.

    Returns (rationale_markdown: str, audit_entries: list[dict])
    """
    # Build optional extra deal fields
    sym   = subject_cfg.get("currency_symbol", "")
    extra = []
    if subject_cfg.get("price_sgd_m"):
        extra.append(f"Asking Price  : {sym}{subject_cfg['price_sgd_m']:.1f}M")
    if subject_cfg.get("ftm_noi_cap_rate"):
        extra.append(f"Cap Rate (FTM): {subject_cfg['ftm_noi_cap_rate']*100:.2f}%")
    if subject_cfg.get("remaining_leasehold_yrs"):
        extra.append(f"Leasehold     : {subject_cfg['remaining_leasehold_yrs']} yrs remaining")
    if subject_cfg.get("land_zoning"):
        extra.append(f"Zoning        : {subject_cfg['land_zoning']}")

    gfa_display = (f"{int(subject_cfg.get('gfa_sf', 0)):,}"
                   if subject_cfg.get("gfa_sf") else "—")

    # ── CALL 1: prose rationale (anonymised sources so filenames never appear) ──
    combined_anon = _merge_insights(extracted_reports, anonymize=True)

    prose_prompt = (
        _RATIONALE_PROMPT
        .replace("DEAL_NAME",         subject_cfg.get("deal_name",     "—"))
        .replace("PROPERTY_NAME",     subject_cfg.get("property_name", "—"))
        .replace("ADDRESS",           subject_cfg.get("address",       "—"))
        .replace("ASSET_CLASS",       subject_cfg.get("asset_class",   "—").title())
        .replace("ASSET_TYPE",        subject_cfg.get("asset_type",    "—"))
        .replace("LOCATION",          subject_cfg.get("location",      "—"))
        .replace("QUALITY",           subject_cfg.get("quality",       "—"))
        .replace("GFA_SF",            gfa_display)
        .replace("GFA_UNIT",          subject_cfg.get("gfa_unit",      "sf").upper())
        .replace("COUNTRY_NAME",      subject_cfg.get("country_name",  "—"))
        .replace("EXTRA_FIELDS",      "\n".join(extra) if extra else "")
        .replace("COMBINED_INSIGHTS", combined_anon)
        .replace("ANALYST_NOTES",     analyst_notes.strip() if analyst_notes else "(none)")
        .replace("REFINEMENT_BLOCK",  (
            "═══ REFINEMENT INSTRUCTIONS (highest priority) ══════════════════\n"
            "The analyst has reviewed the previous draft and requests these specific changes.\n"
            "Address every point below before writing:\n"
            f"{refinement_notes.strip()}\n"
            "════════════════════════════════════════════════════════════════"
        ) if refinement_notes.strip() else "")
    )

    print("  [generating] Investment rationale (3 sections) ...")
    t0_rationale = time.perf_counter()
    try:
        raw_rationale = _llm_chat(
            [
                {"role": "system", "content": _RATIONALE_SYSTEM},
                {"role": "user",   "content": prose_prompt},
            ],
            llm_cfg=llm_cfg,
            openai_key=openai_key,
            temperature=0.3,
        )
    except Exception as e:
        return f"**Error generating rationale:** {e}", []
    t1_rationale = time.perf_counter()
    print(f"  [timing]    Rationale prose  : {t1_rationale - t0_rationale:.1f}s")

    # Clean up any audit headings or stray separator the model appended
    rationale_body = _clean_rationale_body(raw_rationale)

    # ── CALL 2: source audit JSON (real filenames visible to auditor LLM) ───────
    combined_with_src = _merge_insights(extracted_reports, anonymize=False)

    # Build the exact list of valid filenames the audit LLM must choose from
    valid_filenames = "\n".join(
        f"  - {rpt['source_file']}"
        for rpt in extracted_reports if rpt.get("source_file")
    )

    audit_prompt = (
        _AUDIT_PROMPT_TEMPLATE
        .replace("RATIONALE_BODY_PLACEHOLDER",  rationale_body)
        .replace("SOURCE_DATA_PLACEHOLDER",     combined_with_src)
        .replace("VALID_FILENAMES_PLACEHOLDER", valid_filenames)
    )

    print("  [generating] Source audit JSON ...")
    t0_audit = time.perf_counter()
    audit_entries: list[dict] = []
    try:
        raw_audit = _llm_chat(
            [
                {"role": "system", "content": _AUDIT_SYSTEM},
                {"role": "user",   "content": audit_prompt},
            ],
            llm_cfg=llm_cfg,
            openai_key=openai_key,
            temperature=0.1,
        )
        audit_entries = _parse_audit_json(raw_audit)
        if not audit_entries:
            print("  [fallback] Audit JSON parse failed -- trying inline markers")
            audit_entries = _fallback_citations_from_text(rationale_body)
    except Exception as e:
        print(f"  [warning] Audit call failed: {e} -- continuing without audit")
    t1_audit = time.perf_counter()
    print(f"  [timing]    Source audit JSON: {t1_audit - t0_audit:.1f}s")
    print(f"  [timing]    Total generation : {t1_audit - t0_rationale:.1f}s")

    # ── Build document header ──────────────────────────────────────────────────
    prop_name   = subject_cfg.get("property_name", subject_cfg.get("deal_name", "Property"))
    report_srcs = ", ".join(r["source_file"] for r in extracted_reports if r.get("source_file"))
    header = (
        f"# Investment Rationale — {prop_name}\n\n"
        f"**Address:** {subject_cfg.get('address', '—')}  \n"
        f"**Asset Class:** {subject_cfg.get('asset_class','—').title()}  \n"
        f"**Country:** {subject_cfg.get('country_name', '—')}  \n"
        f"**Generated:** {datetime.now().strftime('%d %b %Y')}  \n"
    )
    if report_srcs:
        header += f"**Market Reports Used:** {report_srcs}  \n"
    header += ("\n> _LLM-generated from the cited market reports — verify all figures "
               "against the source PDFs before use._\n")
    header += "\n---\n\n"

    rationale_md = header + _build_input_summary(extracted_reports) + rationale_body

    if audit_entries:
        rationale_md += (
            "\n\n---\n\n"
            f"> **Source Audit:** {len(audit_entries)} citations logged.  "
            "See `Source_Audit.xlsx` for verification details.\n"
        )

    return rationale_md, audit_entries


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT  (called from CLI or via run.py)
# ═══════════════════════════════════════════════════════════════════════════════

def run(
    config_path: str,
    report_paths: list[str] | None = None,
    force_refresh: bool = False,
    analyst_notes: str = "",
    refinement_notes: str = "",
) -> str:
    """
    Full pipeline orchestrator: extract insights → generate rationale → save outputs.

    Called by the Streamlit frontend (via subprocess) and directly from the CLI.
    The frontend always injects the sidebar model choice into the config before
    calling this function, so llm_cfg reflects the user's current model selection.

    Pipeline steps
    --------------
    1. Load config  — read the deal config JSON; build llm_cfg and resolve the
                      OpenAI API key (config field → env var fallback).

    2. Resolve reports  — if report_paths is empty, auto-discover all PDFs in
                          Input_files/market_reports/.

    3. Stage 1: Extract  — for each PDF, call extract_report_insights().
                           Already-extracted reports are served from the cache
                           in under a second.  Only new or force-refreshed reports
                           go through an LLM call.

    4. Stage 2: Generate  — call generate_rationale() which runs two LLM calls:
                            prose writing (anonymised sources) then audit JSON
                            (real filenames).  See generate_rationale() docstring
                            for the full explanation of why two calls are needed.

    5. Save outputs  — write Investment_Rationale.md and Source_Audit.xlsx into
                       the deal's output directory (output/<DealName>/).

    Returns the rationale markdown string.
    """
    import os
    cfg        = json.loads(Path(config_path).read_text(encoding="utf-8"))
    subj       = cfg["subject_property"]
    llm_cfg    = cfg.get("llm") or {}
    # Ensure a default Ollama block exists if the config has no llm section
    # (e.g. a freshly created config before the user has run any analysis).
    if not llm_cfg:
        llm_cfg = {"provider": "ollama",
                   "ollama": {"base_url": "http://localhost:11434", "model": "qwen2.5:3b"}}
    openai_key = (cfg.get("openai", {}).get("api_key")
                  or os.environ.get("OPENAI_API_KEY", ""))

    # Resolve report paths — auto-discover if caller did not specify any
    if not report_paths:
        report_paths = sorted(
            str(p) for p in REPORTS_DIR.glob("*.pdf")
            if not p.name.startswith(".")
        )
        if not report_paths:
            print("[investment_rationale] No market reports found in "
                  f"{REPORTS_DIR}  — generating from property info only.")

    print(f"\n[Investment Rationale] {subj.get('deal_name', '')}")
    print(f"  Asset class : {subj.get('asset_class','-')}")
    print(f"  Country     : {subj.get('country_name','-')}")
    print(f"  Reports     : {len(report_paths)}")

    # ── Stage 1: Extract insights from each market report ────────────────────
    # extract_report_insights() caches results by (filename + size + mtime).
    # If the file hasn't changed since the last run the LLM is not called at all
    # and the cached JSON is returned in milliseconds.
    t_pipeline_start = time.perf_counter()
    extracted = []
    for rp in report_paths:
        t0 = time.perf_counter()
        result = extract_report_insights(rp, llm_cfg, openai_key, force_refresh=force_refresh)
        if result:
            extracted.append(result)
            cached = "(cached)" if (time.perf_counter() - t0) < 0.5 else f"{time.perf_counter() - t0:.1f}s"
            print(f"  [timing]    Extract {Path(rp).name}: {cached}")

    # ── Stage 2: Generate rationale prose + source audit JSON ────────────────
    # Two separate LLM calls — see generate_rationale() docstring for why.
    rationale, audit_entries = generate_rationale(
        extracted, subj, llm_cfg, openai_key,
        analyst_notes=analyst_notes,
        refinement_notes=refinement_notes,
    )

    # ── Save outputs ──────────────────────────────────────────────────────────
    out_dir = ROOT / Path(cfg.get("output_file", "output/deal/deal.xlsx")).parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # Investment_Rationale.md — the prose shown in the dashboard and downloaded
    out_file = out_dir / "Investment_Rationale.md"
    out_file.write_text(rationale, encoding="utf-8")
    print(f"  [OK] Saved -> {out_file.relative_to(ROOT)}")

    # Source_Audit.xlsx — citation table for human verification (never shown inline)
    if audit_entries:
        audit_xlsx = out_dir / "Source_Audit.xlsx"
        try:
            _write_source_audit_excel(audit_entries, extracted, audit_xlsx, subj)
            print(f"  [OK] Saved -> {audit_xlsx.relative_to(ROOT)}")
        except Exception as e:
            print(f"  [warning] Could not write source audit Excel: {e}")
    else:
        print("  [note] No citations captured -- source audit Excel not generated.")

    print(f"\n  [timer] Total pipeline time: {time.perf_counter() - t_pipeline_start:.1f}s")
    return rationale


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="Generate investment rationale from market reports + deal config"
    )
    ap.add_argument("--config",     required=True,
                    help="Path to deal config JSON")
    ap.add_argument("--reports",    nargs="*", default=[],
                    help="PDF/txt report paths.  Defaults to all PDFs in "
                         "Input_files/market_reports/")
    ap.add_argument("--refresh",    action="store_true",
                    help="Re-extract all reports even if cached")
    ap.add_argument("--notes-file", default="",
                    help="Path to a plain-text file with analyst notes")
    ap.add_argument("--refinement-file", default="",
                    help="Path to a plain-text file with refinement/feedback instructions")
    args = ap.parse_args()

    notes = ""
    if args.notes_file:
        try:
            notes = Path(args.notes_file).read_text(encoding="utf-8")
        except Exception as e:
            print(f"[warning] Could not read notes file: {e}")

    refinement = ""
    if args.refinement_file:
        try:
            refinement = Path(args.refinement_file).read_text(encoding="utf-8")
        except Exception as e:
            print(f"[warning] Could not read refinement file: {e}")

    result = run(
        config_path=args.config,
        report_paths=args.reports or None,
        force_refresh=args.refresh,
        analyst_notes=notes,
        refinement_notes=refinement,
    )

    print("\n" + "=" * 70)
    # Print rationale safely — replace any remaining unencodable chars
    safe_result = result.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(
        sys.stdout.encoding or "utf-8"
    )
    print(safe_result)
