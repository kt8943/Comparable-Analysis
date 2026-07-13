#!/usr/bin/env python3
"""
comp_classifier.py
==================
Deterministic-first classifier that decides whether an uploaded file is a set of
**Asset Sales**, **Leasing (Rent)**, or **Land Sales** comparables — so the user
can drop every PDF/Excel into one box and the orchestrator routes each to the
right scan tool (scan_input_sales_comps.py / _rent_ / _land_).

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
        ],
        "weak": [
            "buyer", "acquirer", "sale price", "transacted price", "investment sales",
            "net yield", "psf gfa", "psf on gfa", "acquisition", "seller",
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
        ],
    },
}

_STRONG_PTS, _WEAK_PTS = 3, 1
_MARGIN = 3            # winner must beat runner-up by ≥ this to be "confident"
_MIN_SIGNAL = 3        # below this total the text is treated as signal-free


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


_LLM_SYSTEM = (
    "You classify a real-estate comparables document into exactly one of: "
    "\"sales\" (asset / building sale transactions, with buyers, cap rates, NPI yields), "
    "\"rent\" (leasing deals, tenants, asking/effective rents), or "
    "\"land\" (land / GLS site sales, tenderers, price psf ppr). "
    "Reply ONLY as JSON: {\"type\": \"sales|rent|land|unknown\", \"reason\": \"<8 words\"}. "
    "Use \"unknown\" if the excerpt is not a comparables table."
)


def _llm_classify(text: str, llm_cfg: dict, openai_key: str = "") -> dict | None:
    """One bounded LLM tie-break. Returns {type, reason} or None on any failure."""
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
        t = str(obj.get("type", "")).lower().strip()
        if t in ("sales", "rent", "land", "unknown"):
            return {"type": t, "reason": str(obj.get("reason", "")).strip()}
    except Exception:
        return None
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────
_LABEL = {"sales": "Asset Sales", "rent": "Rent", "land": "Land", "unknown": "Unknown"}


def classify_file(path, llm_cfg: dict | None = None, openai_key: str = "",
                  allow_llm: bool = True) -> dict:
    """Classify one file.

    Returns {path, name, type ∈ {sales,rent,land,unknown}, label, confidence
    ∈ {high,low,none}, scores, reason, method}. Keyword-first; LLM only breaks a
    genuine tie / signal-free text when an LLM is configured."""
    p = Path(path)
    text = extract_text(p)
    sc = score_text(text)
    scores, hits = sc["scores"], sc["hits"]
    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    top_type, top_score = ranked[0]
    second_score = ranked[1][1] if len(ranked) > 1 else 0
    margin = top_score - second_score

    # Clear keyword winner
    if top_score >= _MIN_SIGNAL and margin >= _MARGIN:
        why = ", ".join(hits[top_type][:4]) or "keyword match"
        return {"path": str(p), "name": p.name, "type": top_type,
                "label": _LABEL[top_type], "confidence": "high",
                "scores": scores, "reason": f"matched: {why}", "method": "keywords"}

    # Ambiguous or signal-free → optional single LLM tie-break
    if allow_llm and (llm_cfg or openai_key):
        llm = _llm_classify(text, llm_cfg or {}, openai_key)
        if llm and llm["type"] != "unknown":
            return {"path": str(p), "name": p.name, "type": llm["type"],
                    "label": _LABEL[llm["type"]], "confidence": "low",
                    "scores": scores,
                    "reason": llm.get("reason") or "LLM tie-break", "method": "llm"}

    # Fall back to the best keyword guess if there was *any* signal, else unknown
    if top_score >= _MIN_SIGNAL:
        why = ", ".join(hits[top_type][:4]) or "weak keyword match"
        return {"path": str(p), "name": p.name, "type": top_type,
                "label": _LABEL[top_type], "confidence": "low",
                "scores": scores, "reason": f"best guess: {why}", "method": "keywords"}
    return {"path": str(p), "name": p.name, "type": "unknown",
            "label": _LABEL["unknown"], "confidence": "none",
            "scores": scores,
            "reason": "no comp-type signal found — please assign", "method": "none"}


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
        print(f"{r['label']:<12} [{r['confidence']:<4}] {Path(f).name}"
              f"  scores={r['scores']}  — {r['reason']}")
