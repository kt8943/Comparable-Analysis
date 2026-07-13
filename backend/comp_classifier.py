#!/usr/bin/env python3
"""
comp_classifier.py
==================
Deterministic-first classifier that detects which comparable-table types an uploaded
file contains — **Asset Sales**, **Leasing (Rent)**, and/or **Land Sales** — so the
user can drop every PDF/Excel into one box and the orchestrator routes each file to
the right scan tool(s). Classification is **multi-label**: one file may hold more than
one type (a broker PDF with a land table AND a sales table), and it is routed to EACH
matching scan (their reject-markers keep each type's tables separate). A file that
reads like a market research/outlook report (not a comp table) is flagged is_report so
the UI can nudge it to the Market reports box.

Design (mirrors the project's principle: deterministic where the path is known):
  • Primary signal — keyword scoring on the file's text. The scan scripts already
    encode the discriminating vocabulary (see their reject-marker / column lists);
    we reuse it here. Strong markers weigh more than weak ones.
  • Tie-break — ONE bounded LLM call, only when the text gives no clear winner and
    an LLM is configured. Fixed output enum; never invents a fourth class.
  • Never fabricates: an unreadable / signal-free file returns type "unknown" so
    the UI asks the analyst to assign it.

Pure functions + a small CLI for spot-checking.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# Discriminating vocabulary (lifted from the scan scripts' column + reject lists)
#   strong → 3 pts (near-unique to the type), weak → 1 pt (supportive)
# ─────────────────────────────────────────────────────────────────────────────
_SIGNALS = {
    "land": {
        "strong": [
            "successful tenderer", "psf ppr", "per plot ratio", "psf per plot ratio",
            "date of award", "government land sales", "gls site", "provisional permission",
            "tendered price", "land rate", "plot ratio",
        ],
        "weak": [
            "tenderer", "tender", "site area", "land parcel", "state land",
            "land tender", "gpr", "zoning", "land use", "gross floor area allowable",
        ],
    },
    "sales": {
        "strong": [
            "cap rate", "npi yield", "capitalisation rate", "capitalization rate",
            "ftm noi", "net property income", "en bloc", "vendor", "purchaser",
            # Plain investment-sales TABLE headers (broker reports don't always give
            # cap rates — a "Seller/Buyer + Price" table is still a sales comp table).
            "seller/buyer", "seller / buyer", "buyer/seller", "sales transaction",
        ],
        "weak": [
            "buyer", "acquirer", "sale price", "transacted price", "investment sales",
            "net yield", "psf gfa", "psf on gfa", "acquisition", "seller",
            "capital markets", "key sales", "price (s$",
        ],
    },
    "rent": {
        "strong": [
            "asking rent", "gross rent", "face rent", "passing rent", "headline rent",
            "effective rent", "psf pm", "psf/month", "psf per month", "monthly rent",
            "rent-free", "rent free", "wale",
        ],
        "weak": [
            "tenant", "occupier", "lessee", "lease term", "lease commencement",
            "vacancy", "occupancy", "leasing", "rental", "take-up", "net lettable",
            "lease transaction",
        ],
    },
}

_STRONG_PTS, _WEAK_PTS = 3, 1
_MARGIN = 3            # winner must beat runner-up by ≥ this to be "confident"
_MIN_SIGNAL = 3        # below this total the text is treated as signal-free

# Multi-label inclusion: a type's table is judged PRESENT when the file has ≥1 strong
# marker for it, or enough weak points. Lean toward INCLUDING — a scan for an absent
# type just returns nothing (harmless), while missing one loses comps.
_PRESENT_STRONG = 1    # ≥1 strong marker ⇒ present
_PRESENT_WEAK   = 5    # …or ≥5 weak-only points

# A market research / outlook report discusses every sector, so it trips comp keywords;
# these flag "prose, not a comp table" so the UI can nudge it to the Market reports box.
_REPORT_MARKERS = [
    "outlook", "market report", "research report", "marketbeat", "market commentary",
    "executive summary", "forecast", "our view", "market overview", "market pulse",
    "research & forecast", "quarterly report", "market insights", "economic overview",
]


# ─────────────────────────────────────────────────────────────────────────────
# Text extraction
# ─────────────────────────────────────────────────────────────────────────────
def _pdf_text(path: Path, max_pages: int = 8) -> str:
    """First pages of a PDF as text. Tries pdfplumber, then pypdf, then PyMuPDF."""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return "\n".join((pg.extract_text() or "") for pg in pdf.pages[:max_pages])
    except Exception:
        pass
    try:
        import pypdf
        reader = pypdf.PdfReader(str(path))
        return "\n".join((pg.extract_text() or "") for pg in reader.pages[:max_pages])
    except Exception:
        pass
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(path)
        return "\n".join(doc[i].get_text() for i in range(min(max_pages, doc.page_count)))
    except Exception:
        return ""


def _excel_text(path: Path, max_rows: int = 200) -> str:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        out = []
        for ws in wb.worksheets:
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i >= max_rows:
                    break
                out.append(" ".join(str(c) for c in row if c not in (None, "")))
        return "\n".join(out)
    except Exception:
        return ""


def extract_text(path) -> str:
    """Best-effort text for classification. Images return '' (no cheap text layer)."""
    p = Path(path)
    suf = p.suffix.lower()
    if suf == ".pdf":
        return _pdf_text(p)
    if suf in (".xlsx", ".xls"):
        return _excel_text(p)
    return ""   # png/jpg/… — handled as "unknown" unless an LLM vision pass is added


# ─────────────────────────────────────────────────────────────────────────────
# Scoring
# ─────────────────────────────────────────────────────────────────────────────
def score_text(text: str) -> dict:
    """{type: score} keyword score over the three comp types (case-insensitive,
    each marker counted once so a repeated word can't dominate)."""
    norm = re.sub(r"\s+", " ", (text or "")).lower()
    scores = {}
    hits: dict = {}
    for ctype, groups in _SIGNALS.items():
        s, matched = 0, []
        for kw in groups["strong"]:
            if kw in norm:
                s += _STRONG_PTS
                matched.append(kw)
        for kw in groups["weak"]:
            if kw in norm:
                s += _WEAK_PTS
                matched.append(kw)
        scores[ctype] = s
        hits[ctype] = matched
    return {"scores": scores, "hits": hits}


def _present_types(scores: dict, hits: dict) -> list:
    """Every comp type whose table looks PRESENT (multi-label), best-scored first."""
    present = []
    for t in ("sales", "rent", "land"):
        strong = sum(1 for h in hits.get(t, []) if h in _SIGNALS[t]["strong"])
        if strong >= _PRESENT_STRONG or scores.get(t, 0) >= _PRESENT_WEAK:
            present.append(t)
    present.sort(key=lambda t: scores.get(t, 0), reverse=True)
    return present


def _looks_like_report(text: str, name: str) -> bool:
    """True when the file reads like a market research / outlook report (prose,
    forecasts) rather than a table of named transactions."""
    blob = (name + " " + (text or "")).lower()
    return sum(1 for m in _REPORT_MARKERS if m in blob) >= 2


_LLM_SYSTEM = (
    "You label a real-estate document by WHICH comparable-transaction table types it "
    "contains — a single document may contain MORE THAN ONE. Types: "
    "\"sales\" (asset/building sale transactions — buyers, cap rates, NPI yields), "
    "\"rent\" (leasing deals — tenants, asking/effective rents), "
    "\"land\" (land/GLS site sales — tenderers, price psf ppr). "
    "Also set is_report=true if it is a market research / outlook REPORT (prose, "
    "forecasts) rather than a table of named transactions. "
    "Reply ONLY as JSON: {\"types\": [\"sales\"|\"rent\"|\"land\", ...], "
    "\"is_report\": true|false, \"reason\": \"<10 words\"}. "
    "Use an empty types list when no comparable table is present."
)


def _llm_classify(text: str, llm_cfg: dict, openai_key: str = "") -> dict | None:
    """One bounded, multi-label LLM pass. Returns {types:[...], is_report:bool, reason}
    or None on any failure."""
    excerpt = re.sub(r"\s+", " ", text or "")[:4000]
    if not excerpt.strip():
        return None
    provider = (llm_cfg or {}).get("provider", "")
    msgs = [{"role": "system", "content": _LLM_SYSTEM},
            {"role": "user", "content": excerpt}]
    try:
        raw = None
        if provider == "openai" or openai_key:
            from tools.llm_client import openai_chat
            _cfg = dict(llm_cfg or {})
            _cfg.setdefault("provider", "openai")
            if openai_key:
                _cfg["openai_api_key"] = openai_key
            raw = openai_chat(_cfg, msgs, json_mode=True)
        elif provider == "ollama":
            from tools.llm_client import ollama_post
            oll = (llm_cfg or {}).get("ollama", {})
            raw = ollama_post(oll.get("base_url", "http://localhost:11434"),
                              oll.get("model", "qwen2.5:3b"), msgs, timeout=60)
        if not raw:
            return None
        m = re.search(r"\{.*\}", raw, re.S)
        obj = json.loads(m.group(0) if m else raw)
        types = [t for t in (obj.get("types") or []) if t in ("sales", "rent", "land")]
        return {"types": types, "is_report": bool(obj.get("is_report")),
                "reason": str(obj.get("reason", "")).strip()}
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────
_LABEL = {"sales": "Asset Sales", "rent": "Rent", "land": "Land", "unknown": "Unknown"}


def _result(p, types, is_report, confidence, scores, reason, method) -> dict:
    """Assemble the return. Keeps single-label keys (type/label) for backward-compat
    while adding `types` (multi-label list) and `is_report`."""
    return {
        "path": str(p), "name": p.name,
        "types": list(types),                       # multi-label (may be [], 1, or many)
        "type": types[0] if types else "unknown",   # primary — back-compat
        "label": " + ".join(_LABEL[t] for t in types) if types else _LABEL["unknown"],
        "is_report": bool(is_report),
        "confidence": confidence, "scores": scores,
        "reason": reason, "method": method,
    }


def classify_file(path, llm_cfg: dict | None = None, openai_key: str = "",
                  allow_llm: bool = True) -> dict:
    """Classify one file (MULTI-LABEL). A file may contain more than one comp-table
    type, or be a market report.

    Returns {path, name, types:[…], type (primary), label, is_report, confidence
    ∈ {high,low,none}, scores, reason, method}. Keyword-first; the multi-label LLM pass
    only runs when the keyword signal is inconclusive and an LLM is configured."""
    p = Path(path)
    text = extract_text(p)
    sc = score_text(text)
    scores, hits = sc["scores"], sc["hits"]
    present = _present_types(scores, hits)
    is_report = _looks_like_report(text, p.name)

    # Clear keyword signal — one or more types present
    if present:
        conf = "high" if (len(present) == 1 and not is_report) else "low"
        why = ", ".join(hits[present[0]][:4]) or "keyword match"
        reason = ("matched: " + why
                  + ("; also reads like a market report" if is_report else ""))
        return _result(p, present, is_report, conf, scores, reason, "keywords")

    # Inconclusive keywords → one bounded multi-label LLM pass if available
    if allow_llm and (llm_cfg or openai_key):
        llm = _llm_classify(text, llm_cfg or {}, openai_key)
        if llm and (llm["types"] or llm["is_report"]):
            return _result(p, llm["types"], llm["is_report"] or is_report, "low",
                           scores, llm.get("reason") or "LLM", "llm")

    # Weak single guess if there was any signal; else report / unknown
    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    if ranked[0][1] >= _MIN_SIGNAL:
        t = ranked[0][0]
        why = ", ".join(hits[t][:4]) or "weak keyword match"
        return _result(p, [t], is_report, "low", scores, f"best guess: {why}", "keywords")
    if is_report:
        return _result(p, [], True, "none", scores,
                       "looks like a market report — use the Market reports box", "report")
    return _result(p, [], False, "none", scores,
                   "no comp-type signal found — please assign", "none")


def classify_files(paths, llm_cfg: dict | None = None, openai_key: str = "",
                   allow_llm: bool = True) -> list:
    """Classify many files; returns a list of classify_file() dicts."""
    return [classify_file(p, llm_cfg, openai_key, allow_llm) for p in paths]


# ─────────────────────────────────────────────────────────────────────────────
# CLI (spot-check): python comp_classifier.py file1.pdf file2.xlsx ...
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    for f in sys.argv[1:]:
        r = classify_file(f, allow_llm=False)
        tag = r["label"] + ("  +report" if r["is_report"] else "")
        print(f"{tag:<24} [{r['confidence']:<4}] {Path(f).name}"
              f"  scores={r['scores']}  — {r['reason']}")
