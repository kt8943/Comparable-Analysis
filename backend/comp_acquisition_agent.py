#!/usr/bin/env python3
"""
comp_acquisition_agent.py
=========================
The agentic layer over the deterministic comp scripts (see
docs/comp_acquisition_agent.md). Pure functions — the frontend runs the actual
scan / online-search scripts (as tools) and calls these to VERIFY the result,
EVALUATE quality, and REFLECT on failure to pick a fallback.

Design rules:
  • verify/flag only — never invent or "correct" a number
  • deterministic scoring; the LLM is used only for the reflection step
  • bounded, auditable — every decision returns a small typed dict
"""

from __future__ import annotations

import glob
import json
import re
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# Per-type configuration
# ─────────────────────────────────────────────────────────────────────────────
# comp_type → {min_comps, valid_any (≥1 present ⇒ a valid comp),
#              ground (key figures to check against the source, first present wins),
#              input_keys (config keys holding the uploaded files)}
_TYPE_CFG = {
    "sales": {
        "min_comps": 3,
        "valid_any": ["price_sgd_m", "price_psf_gfa"],
        "ground":    ["price_sgd_m", "price_psf_gfa", "gfa_sf"],
        "input_keys": ("input_file", "input_pdf_file", "input_image_file"),
    },
    "rent": {
        "min_comps": 5,
        "valid_any": ["nla_sf", "asking_rent", "eff_rent"],
        "ground":    ["asking_rent", "eff_rent", "nla_sf"],
        "input_keys": ("rent_input_file", "rent_input_pdf_file", "rent_input_image_file"),
    },
    "land": {
        "min_comps": 3,
        "valid_any": ["price_sgd_m", "price_psf_ppr"],
        "ground":    ["price_sgd_m", "price_psf_ppr", "max_gfa_sf"],
        "input_keys": ("land_input_file", "land_input_pdf_file", "land_input_image_file"),
    },
}

GROUNDED_MIN = 0.5      # share of comps grounded to accept the result
LOW_CONF     = 0.5      # overall confidence below this → flag for the analyst

# Run-log signatures that mean "extraction produced nothing usable"
_HARD_FLAG_PATTERNS = {
    "no comps found":       r"No (?:qualifying|eligible)",
    "zero valid records":   r"0 (?:valid|eligible)",
    "centroid geocode":     r"ON COUNTRY CENTROID",
    "looks like a report":  r"skipping this file|extraction failed",
}


# ─────────────────────────────────────────────────────────────────────────────
# Records + source text
# ─────────────────────────────────────────────────────────────────────────────
def read_records(out_dir, prefix: str) -> list:
    """Load comp dicts for a prefix. Uses {prefix}*_records.json (scan output) if
    present, else falls back to {prefix}*_search_cache.json's 'records' list
    (online-search output)."""
    files = glob.glob(str(Path(out_dir) / f"{prefix}*_records.json"))
    if files:
        latest = max(files, key=lambda f: Path(f).stat().st_mtime)
        try:
            data = json.load(open(latest, encoding="utf-8"))
            return data if isinstance(data, list) else []
        except Exception:
            return []
    caches = glob.glob(str(Path(out_dir) / f"{prefix}*_search_cache.json"))
    if caches:
        latest = max(caches, key=lambda f: Path(f).stat().st_mtime)
        try:
            data = json.load(open(latest, encoding="utf-8"))
            recs = data.get("records") if isinstance(data, dict) else data
            return recs if isinstance(recs, list) else []
        except Exception:
            return []
    return []


def _read_excel_text(path: Path) -> str:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        out = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                out.append(" ".join(str(c) for c in row if c not in (None, "")))
        return "\n".join(out)
    except Exception:
        return ""


def _read_pdf_text(path: Path) -> str:
    try:
        import pdfplumber
        out = []
        with pdfplumber.open(path) as pdf:
            for pg in pdf.pages[:40]:
                out.append(pg.extract_text() or "")
        return "\n".join(out)
    except Exception:
        return ""


def source_text(cfg: dict, comp_type: str, source_kind: str, out_dir) -> str:
    """Concatenated text of the source the comps came from, for grounding.
    source_kind: 'files' → the configured input files; 'online' → the search cache."""
    root = Path(cfg.get("_root", "."))
    if source_kind == "online":
        parts = []
        for f in glob.glob(str(Path(out_dir) / "Online_*_search_cache.json")):
            try:
                parts.append(Path(f).read_text(encoding="utf-8"))
            except Exception:
                pass
        return "\n".join(parts)

    # files
    keys = _TYPE_CFG.get(comp_type, {}).get("input_keys", ())
    paths = []
    for k in keys:
        v = cfg.get(k)
        if not v:
            continue
        paths += (v if isinstance(v, list) else [v])
    text = []
    for p in paths:
        fp = (root / p) if not Path(p).is_absolute() else Path(p)
        if not fp.exists():
            continue
        suf = fp.suffix.lower()
        if suf in (".xlsx", ".xls"):
            text.append(_read_excel_text(fp))
        elif suf == ".pdf":
            text.append(_read_pdf_text(fp))
    return "\n".join(text)


# ─────────────────────────────────────────────────────────────────────────────
# verify_comps — grounding check (exact number match against the source)
# ─────────────────────────────────────────────────────────────────────────────
def _figure_candidates(val) -> list:
    """String forms a numeric figure might take in the source text."""
    try:
        f = float(str(val).replace(",", ""))
    except (TypeError, ValueError):
        return []
    cands = set()
    if f == int(f):
        i = int(f)
        cands.add(str(i))
        cands.add(f"{i:,}")            # 88,500
    # 1 dp and raw
    cands.add(f"{f:g}")               # 91.8
    cands.add(f"{f:.1f}")             # 91.8
    return [c for c in cands if c]


def verify_comps(records: list, src_text: str, comp_type: str) -> list:
    """For each comp, check whether its key figure appears in the source text.
    Returns [{marker, grounded, figure, field, evidence}]. Grounding is exact
    number matching only — it flags, it never edits a value."""
    ground_fields = _TYPE_CFG.get(comp_type, {}).get("ground", [])
    norm = re.sub(r"\s+", " ", (src_text or "")).lower()
    norm_nc = norm.replace(",", "")
    checks = []
    for r in records:
        marker = str(r.get("map_marker", "") or "")
        field, figure, grounded, evidence = None, None, None, ""
        for fld in ground_fields:
            if r.get(fld) in (None, "", 0):
                continue
            field, figure = fld, r.get(fld)
            cands = _figure_candidates(figure)
            hit = next((c for c in cands
                        if c.lower() in norm or c.replace(",", "").lower() in norm_nc),
                       None)
            grounded = hit is not None
            if grounded:
                # short evidence window around the match
                idx = norm_nc.find(hit.replace(",", "").lower())
                if idx >= 0:
                    evidence = norm_nc[max(0, idx - 40): idx + 40]
            break
        checks.append({
            "marker": marker, "grounded": bool(grounded),
            "field": field, "figure": figure, "evidence": evidence,
        })
    return checks


# ─────────────────────────────────────────────────────────────────────────────
# evaluate — deterministic quality score + red flags
# ─────────────────────────────────────────────────────────────────────────────
def _is_valid(rec: dict, comp_type: str) -> bool:
    return any(rec.get(k) not in (None, "", 0)
               for k in _TYPE_CFG.get(comp_type, {}).get("valid_any", []))


def evaluate(records: list, checks: list, run_log: str, comp_type: str) -> dict:
    """Score the acquisition. Returns {n_records, n_valid, pct_grounded,
    confidence, flags, ok}."""
    cfg = _TYPE_CFG.get(comp_type, {})
    n_records = len(records)
    n_valid = sum(1 for r in records if _is_valid(r, comp_type))
    n_ground = sum(1 for c in checks if c.get("grounded"))
    pct_grounded = (n_ground / n_records) if n_records else 0.0

    flags = []
    log = run_log or ""
    for label, pat in _HARD_FLAG_PATTERNS.items():
        if re.search(pat, log):
            flags.append(label)
    if n_valid == 0:
        flags.append("no valid comps extracted")

    min_comps = cfg.get("min_comps", 3)
    hard = any(f in ("no comps found", "zero valid records", "no valid comps extracted",
                     "looks like a report") for f in flags)
    ok = (n_valid >= min_comps) and (pct_grounded >= GROUNDED_MIN) and not hard

    # confidence: blend of coverage vs target + grounding, minus a hit for flags
    coverage = min(1.0, n_valid / min_comps) if min_comps else 0.0
    confidence = round(max(0.0, 0.5 * coverage + 0.5 * pct_grounded
                           - (0.2 if hard else 0.0)), 2)
    return {
        "n_records": n_records, "n_valid": n_valid,
        "pct_grounded": round(pct_grounded, 2),
        "confidence": confidence, "flags": flags, "ok": ok,
        "min_comps": min_comps,
    }


# ─────────────────────────────────────────────────────────────────────────────
# reflect — one bounded LLM step (diagnosis + next action from a fixed enum)
# ─────────────────────────────────────────────────────────────────────────────
_REFLECT_SYSTEM = (
    "You are a QA step in a real-estate comparables pipeline. Extraction under-"
    "performed. In ONE short sentence say WHY, then choose the single best next "
    "source to try. Reply ONLY as JSON: {\"diagnosis\": str, \"next_action\": str}. "
    "next_action MUST be one of the 'remaining' source names, or \"stop\" if "
    "remaining is empty or nothing else is worth trying."
)


def _rule_reflect(evaluation: dict, tried: list, remaining: list) -> dict:
    """Deterministic fallback when no LLM is available. next_action is a source
    name from `remaining` (preferring 'online'), or 'stop'."""
    if not remaining:
        action = "stop"
    elif "online" in remaining:
        action = "online"
    else:
        action = remaining[0]
    why = "no valid comps" if evaluation.get("n_valid", 0) == 0 else \
          "too few / weakly-grounded comps"
    if "looks like a report" in evaluation.get("flags", []):
        why = "the input looks like a market report, not a comp table"
    tail = "stop" if action == "stop" else f"try the {action} source"
    return {"diagnosis": f"{why}; will {tail}.", "next_action": action}


def reflect(evaluation: dict, tried: list, remaining: list,
            llm_cfg: dict | None = None, openai_key: str | None = None) -> dict:
    """Diagnose the failure and pick the next action (validated against
    `remaining`). Falls back to rule-based logic without an LLM."""
    provider = (llm_cfg or {}).get("provider", "")
    try:
        payload = {
            "n_valid": evaluation.get("n_valid"),
            "target_min": evaluation.get("min_comps"),
            "pct_grounded": evaluation.get("pct_grounded"),
            "flags": evaluation.get("flags"),
            "tried": tried, "remaining": remaining,
        }
        msgs = [{"role": "system", "content": _REFLECT_SYSTEM},
                {"role": "user", "content": json.dumps(payload)}]
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
        if raw:
            m = re.search(r"\{.*\}", raw, re.S)
            obj = json.loads(m.group(0)) if m else json.loads(raw)
            action = obj.get("next_action")
            if action != "stop" and action not in remaining:
                # keep the diagnosis, but repair an invalid/exhausted source choice
                action = (remaining[0] if remaining else "stop")
            return {"diagnosis": str(obj.get("diagnosis", "")).strip()
                    or _rule_reflect(evaluation, tried, remaining)["diagnosis"],
                    "next_action": action}
    except Exception:
        pass
    return _rule_reflect(evaluation, tried, remaining)
