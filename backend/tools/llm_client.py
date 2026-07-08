"""
tools/llm_client.py
===================
Thin wrappers for Ollama and OpenAI LLM calls, plus a general agent loop.

Public API
----------
ollama_post(base_url, model, messages, timeout=90) -> str
    Returns the content string from the LLM response.

apply_refinement(records, instructions, llm_cfg, base_url=None, model=None) -> list
    Apply analyst free-text instructions to a list of comparable records.
    Routes to run_agent_loop_gpt (OpenAI) or run_agent_loop (Ollama) based on
    llm_cfg["provider"].

run_agent_loop_gpt(instruction, records, llm_cfg) -> list
    GPT-4o function-calling agent: sees all records, decides in one call.
    Supports filter, keep/reorder, and no-change.

run_agent_loop(instruction, context, tools, base_url, model, max_turns=5) -> dict
    Ollama multi-turn agent loop (query → inspect → action).
"""

import json
import math
import re
import urllib.request


# ── Low-level HTTP helpers ────────────────────────────────────────────────────

def ollama_post(base_url: str, model: str, messages: list,
                timeout: int = 90) -> str:
    """POST a chat request to Ollama (JSON mode) and return content string."""
    payload = json.dumps({
        "model": model, "messages": messages,
        "stream": False, "format": "json",
        "options": {"temperature": 0},
    }).encode()
    req = urllib.request.Request(
        f"{base_url}/api/chat", data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read())["message"]["content"]


def _ollama_raw(base_url: str, model: str, messages: list,
                timeout: int = 60) -> str:
    """POST to Ollama without JSON mode — returns raw content string."""
    payload = json.dumps({
        "model": model, "messages": messages,
        "stream": False, "options": {"temperature": 0},
    }).encode()
    req = urllib.request.Request(
        f"{base_url.rstrip('/')}/api/chat", data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read())["message"]["content"]


def openai_chat(llm_cfg: dict, messages: list,
                json_mode: bool = False, timeout: int = 120) -> str:
    """OpenAI chat completion. Reads provider/openai_model/openai_api_key from llm_cfg."""
    try:
        import openai as _openai
    except ImportError:
        raise ImportError("pip install openai")
    key    = llm_cfg.get("openai_api_key", "") or None  # None triggers env-var fallback
    model  = llm_cfg.get("openai_model", "gpt-4o-mini")
    # max_retries=2: auto-retry on transient connection drops (503, timeout, mid-stream close)
    client = _openai.OpenAI(api_key=key, max_retries=2)
    kwargs: dict = {"model": model, "messages": messages, "temperature": 0, "timeout": timeout}
    if json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    resp = client.chat.completions.create(**kwargs)
    return resp.choices[0].message.content


def _extract_json_obj(text: str) -> dict | None:
    """Extract the first {...} JSON object from a string."""
    text = re.sub(r"^```[a-z]*\n?", "", text.strip())
    text = re.sub(r"\n?```$", "", text)
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        return None
    try:
        return json.loads(m.group(0))
    except json.JSONDecodeError:
        return None


# ── Query tools  (return data to the LLM, loop continues) ────────────────────

def _qt_compute_stats(records: list, field: str) -> dict:
    """Return mean, std, min, max, median for a numeric field."""
    values = []
    for r in records:
        v = r.get(field)
        if v is None:
            continue
        try:
            values.append(float(v))
        except (TypeError, ValueError):
            pass
    if not values:
        return {"error": f"No numeric values found for field '{field}'"}
    n  = len(values)
    sv = sorted(values)
    mu = sum(values) / n
    sd = math.sqrt(sum((x - mu) ** 2 for x in values) / n)
    return {
        "count":  n,
        "mean":   round(mu, 3),
        "std":    round(sd, 3),
        "min":    round(sv[0], 3),
        "p25":    round(sv[n // 4], 3),
        "median": round(sv[n // 2], 3),
        "p75":    round(sv[3 * n // 4], 3),
        "max":    round(sv[-1], 3),
    }


def _qt_get_values(records: list, field: str) -> list:
    """Return each record's value for a field (position + name + value)."""
    out = []
    for i, r in enumerate(records, 1):
        name = str(r.get("property_name") or r.get("site_name") or "")
        out.append({"pos": i, "name": name, "value": r.get(field)})
    return out


_QUERY_TOOLS: dict[str, callable] = {
    "compute_stats": _qt_compute_stats,
    "get_values":    _qt_get_values,
}


# ── Action tools  (filter records, loop ends) ─────────────────────────────────

_CMP_OPS = {
    ">":  lambda a, b: a > b,
    ">=": lambda a, b: a >= b,
    "<":  lambda a, b: a < b,
    "<=": lambda a, b: a <= b,
    "==": lambda a, b: a == b,
    "!=": lambda a, b: a != b,
}


def _at_filter_numeric(records, field, op, value, **_) -> set[int]:
    # Normalise op: extract the first valid operator token if LLM returns garbage
    op_str = str(op).strip()
    resolved = op_str if op_str in _CMP_OPS else next(
        (k for k in (">=", "<=", "!=", ">", "<", "==") if k in op_str), None
    )
    if resolved is None:
        raise ValueError(f"Unknown operator: {op!r}")
    cmp = _CMP_OPS[resolved]
    thr = float(value)
    return {i for i, r in enumerate(records, 1)
            if r.get(field) is not None
            and _safe_cmp(cmp, r[field], thr)}


def _safe_cmp(cmp, raw, thr) -> bool:
    try:
        return cmp(float(raw), thr)
    except (TypeError, ValueError):
        return False


def _at_filter_by_marker(records, markers, **_) -> set[int]:
    targets = {str(m) for m in markers}
    return {i for i, r in enumerate(records, 1)
            if str(r.get("_map_marker", "")) in targets}


def _at_filter_by_name(records, names, **_) -> set[int]:
    lowered = [n.lower() for n in names]
    out: set[int] = set()
    for i, r in enumerate(records, 1):
        prop = str(r.get("property_name") or r.get("site_name") or "").lower()
        if any(sub in prop for sub in lowered):
            out.add(i)
    return out


def _at_filter_by_position(records, positions, **_) -> set[int]:
    return {int(p) for p in positions if 1 <= int(p) <= len(records)}


def _at_filter_last_n(records, n, **_) -> set[int]:
    n = max(0, min(int(n), len(records)))
    return set(range(len(records) - n + 1, len(records) + 1))


def _at_filter_first_n(records, n, **_) -> set[int]:
    n = max(0, min(int(n), len(records)))
    return set(range(1, n + 1))


def _at_filter_by_field_value(records, field, value, **_) -> set[int]:
    """Remove records where field contains value as a substring (case-insensitive)."""
    needle = str(value).lower().strip()
    return {i for i, r in enumerate(records, 1)
            if needle in str(r.get(field, "") or "").lower()}


def _at_no_change(records, **_) -> set[int]:
    return set()


_ACTION_TOOLS: dict[str, callable] = {
    "filter_numeric":        _at_filter_numeric,
    "filter_by_field_value": _at_filter_by_field_value,
    "filter_by_marker":      _at_filter_by_marker,
    "filter_by_name":        _at_filter_by_name,
    "filter_by_position":    _at_filter_by_position,
    "filter_last_n":         _at_filter_last_n,
    "filter_first_n":        _at_filter_first_n,
    "no_change":             _at_no_change,
}


# ── Agent system prompt ───────────────────────────────────────────────────────

_AGENT_SYSTEM = """You are a data-filtering agent for real estate comparable records.
Each response must be a single JSON object — no explanation outside the JSON.

QUERY tools  (gather information; loop continues):
  {"type":"query","tool":"compute_stats","field":"<field_name>"}
      → returns count, mean, std, min, p25, median, p75, max for that field
  {"type":"query","tool":"get_values","field":"<field_name>"}
      → returns every record's value for that field

ACTION tools  (apply the filter; loop ends):
  {"type":"action","calls":[
    {"tool":"filter_numeric","field":"price_sgd_m","op":">","value":1000},
    ...
  ]}

  Valid tools inside "calls":
    filter_numeric        — {"tool":"filter_numeric","field":"<field_name>","op":"<op>","value":<number>}
                            op must be one of: >  >=  <  <=  ==  !=
    filter_by_field_value — {"tool":"filter_by_field_value","field":"<field_name>","value":"<substring>"}
                            removes records where field contains the substring (case-insensitive)
                            use for string fields: land_zoning, sale_type, asset_type, location, quality
                            example: remove land_zoning containing "Mixed" →
                              {"tool":"filter_by_field_value","field":"land_zoning","value":"Mixed"}
    filter_by_marker      — {"tool":"filter_by_marker","markers":[<int>, ...]}
    filter_by_name        — {"tool":"filter_by_name","names":["<substring>", ...]}
    filter_by_position    — {"tool":"filter_by_position","positions":[<int>, ...]}
    filter_last_n         — {"tool":"filter_last_n","n":<int>}
    filter_first_n        — {"tool":"filter_first_n","n":<int>}
    no_change             — {"tool":"no_change"}

Rules:
- Use query tools first when the instruction requires knowing statistics or
  specific values (e.g. "remove outliers", "keep only above average").
- Use action tools when you have enough information to decide.
- The calls array may contain multiple action tools if the instruction requires
  removing records on more than one criterion.
- Do not remove records that do not meet the stated criteria.
- IMPORTANT: Your ONLY job is to REMOVE records. You cannot read files, extract
  metadata, add fields, look up quarter numbers, or perform any action other than
  filtering. If the instruction is not asking you to remove specific records,
  respond with no_change: {"type":"action","calls":[{"tool":"no_change"}]}
""".strip()


# ── General agent loop ────────────────────────────────────────────────────────

def run_agent_loop(
    instruction: str,
    records: list,
    base_url: str,
    model: str,
    max_turns: int = 5,
) -> list:
    """
    General multi-turn agent loop for filtering comparable records.

    The LLM may call query tools (compute_stats, get_values) to inspect the
    data before deciding on action tools.  All numeric comparisons and index
    lookups are executed by Python — the LLM only selects tools and provides
    parameters.

    Returns the filtered list of records.
    """
    _SKIP      = {"_map_marker", "_source", "raw_description", "lon", "lat",
                  "distance_km", "map_marker"}
    _NAME_KEYS = {"property_name", "site_name"}
    # String fields worth showing in the index (filterable via filter_by_field_value)
    _STR_FIELDS = {"land_zoning", "sale_type", "asset_type", "location", "quality",
                   "tenure", "lease_type", "district"}

    # Collect numeric field names present in the data
    numeric_fields: set[str] = set()
    string_fields:  set[str] = set()
    for r in records:
        for k, v in r.items():
            if k in _SKIP or k in _NAME_KEYS or v is None or v == "":
                continue
            if k in _STR_FIELDS and str(v).strip():
                string_fields.add(k)
                continue
            try:
                float(v)
                numeric_fields.add(k)
            except (TypeError, ValueError):
                pass

    # Build compact record index (name + marker + string fields + numeric values)
    index_lines = []
    for i, r in enumerate(records, 1):
        name   = str(r.get("property_name") or r.get("site_name") or "")
        marker = r.get("_map_marker", "")
        parts  = [f'{i}. "{name}"']
        if marker:
            parts.append(f"marker={marker}")
        for k in sorted(string_fields):
            v = r.get(k)
            if v is not None and str(v).strip():
                parts.append(f'{k}="{v}"')
        for k in sorted(numeric_fields):
            v = r.get(k)
            if v is not None and v != "":
                parts.append(f"{k}={v}")
        index_lines.append("  " + ", ".join(parts))

    user_init = (
        f"Available numeric fields: {sorted(numeric_fields)}\n"
        f"Available string fields: {sorted(string_fields)}\n\n"
        f"Records ({len(records)} total):\n" + "\n".join(index_lines) + "\n\n"
        f"Instruction: {instruction}"
    )

    messages = [
        {"role": "system", "content": _AGENT_SYSTEM},
        {"role": "user",   "content": user_init},
    ]

    for turn in range(1, max_turns + 1):
        raw  = _ollama_raw(base_url, model, messages)
        resp = _extract_json_obj(raw)
        if resp is None:
            print(f"  [Agent] Turn {turn}: no JSON in response — stopping")
            break

        resp_type = resp.get("type")

        # ── Query turn ───────────────────────────────────────────────────────
        if resp_type == "query":
            tool_name = resp.get("tool", "")
            field     = resp.get("field", "")
            print(f"  [Agent] Turn {turn}: query → {tool_name}({field})")
            qt = _QUERY_TOOLS.get(tool_name)
            if qt is None:
                result = {"error": f"Unknown query tool '{tool_name}'"}
            else:
                try:
                    result = qt(records, field)
                except Exception as e:
                    result = {"error": str(e)}
            print(f"  [Agent] Query result: {result}")
            messages.append({"role": "assistant", "content": raw})
            messages.append({
                "role": "user",
                "content": (
                    f"Query result for {tool_name}({field!r}):\n"
                    f"{json.dumps(result)}\n\n"
                    "Now return your action (or another query if needed)."
                ),
            })
            continue

        # ── Action turn ──────────────────────────────────────────────────────
        if resp_type == "action":
            calls = resp.get("calls", [])
            # Also accept a flat single-call form {"type":"action","tool":"..."}
            if not calls and resp.get("tool"):
                calls = [resp]

            to_remove: set[int] = set()
            for call in calls:
                if not isinstance(call, dict):
                    continue
                tool_name = call.get("tool", "")
                at = _ACTION_TOOLS.get(tool_name)
                if at is None:
                    print(f"  [Agent] Unknown action tool '{tool_name}' — skipped")
                    continue
                try:
                    positions = at(records, **{k: v for k, v in call.items()
                                               if k != "tool"})
                    if positions:
                        print(f"  [Agent] Turn {turn}: {tool_name} → "
                              f"remove positions {sorted(positions)}")
                    to_remove |= positions
                except Exception as e:
                    print(f"  [Agent] Tool {tool_name} failed ({e}) — skipped")

            filtered = [r for i, r in enumerate(records, 1) if i not in to_remove]
            removed_names = [
                str(records[i - 1].get("property_name")
                    or records[i - 1].get("site_name") or "?")
                for i in sorted(to_remove) if 1 <= i <= len(records)
            ]
            print(f"  [Refinement] Removing {len(to_remove)} record(s): {removed_names}")
            print(f"  [Refinement] {len(records)} → {len(filtered)} records after refinement")
            return filtered

        print(f"  [Agent] Unexpected response type '{resp_type}' — stopping")
        break

    print("  [Agent] No action reached — keeping all records")
    return records


# ── GPT-4o function-calling refinement agent ──────────────────────────────────

_GPT_REFINE_SYSTEM = """You are an analyst assistant helping filter and reorder real estate comparable transaction records.

You will receive a numbered list of records and an instruction from the analyst.
Interpret the instruction carefully and call exactly ONE of the available tools.

You can:
- Remove records that don't meet the stated criteria (filter_records)
- Keep only specific records in a specified order — use this for "sort by", "top N", or "most recent" (keep_records)
- Do nothing if the instruction is unclear or no change is needed (no_change)

Be conservative: if a record's data is missing or ambiguous, keep it rather than removing it.
For date comparisons, parse dates from any format (e.g. "Jan 2026", "2026-01-15", "Q1 2026").
Always provide a clear reason explaining what was done and why."""

_GPT_REFINE_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "filter_records",
            "description": (
                "Remove records that don't meet the analyst's criteria. "
                "Use when you want to exclude specific records based on any field value, "
                "date range, price threshold, distance, tenure type, etc."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "indices_to_remove": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "1-based position numbers of records to REMOVE"
                    },
                    "reason": {
                        "type": "string",
                        "description": "Explain which records were removed and the criteria applied"
                    }
                },
                "required": ["indices_to_remove", "reason"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "keep_records",
            "description": (
                "Keep only specific records, in the given order. "
                "Use for 'top N', 'most recent', 'sort by price', or any instruction "
                "that selects a subset and/or changes the order."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "indices_to_keep": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": (
                            "1-based position numbers of records to KEEP, "
                            "in the desired output order"
                        )
                    },
                    "reason": {
                        "type": "string",
                        "description": "Explain what was selected and the ordering applied"
                    }
                },
                "required": ["indices_to_keep", "reason"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "no_change",
            "description": "Keep all records unchanged when no filtering or reordering is needed.",
            "parameters": {
                "type": "object",
                "properties": {
                    "reason": {
                        "type": "string",
                        "description": "Explain why no change was made"
                    }
                },
                "required": ["reason"]
            }
        }
    }
]


def run_agent_loop_gpt(instruction: str, records: list, llm_cfg: dict) -> list:
    """
    GPT-4o function-calling refinement agent.

    Sends all records + analyst instruction to GPT in a single call.
    GPT decides whether to filter (remove), keep/reorder (select subset in order),
    or leave unchanged — using native function calling for reliable structured output.
    """
    try:
        import openai as _openai
    except ImportError:
        raise ImportError("pip install openai")

    api_key = (
        llm_cfg.get("openai_api_key")
        or __import__("os").environ.get("OPENAI_API_KEY", "")
    )
    if not api_key:
        raise ValueError("OpenAI API key not found in llm_cfg or OPENAI_API_KEY env var")

    gpt_model = llm_cfg.get("openai_model", "gpt-4o")
    client = _openai.OpenAI(api_key=api_key)

    _SKIP = {"_map_marker", "_source", "raw_description"}

    # Build compact numbered record list for GPT
    rows = []
    for i, r in enumerate(records, 1):
        parts: list[str] = [f"#{i}"]
        name = r.get("property_name") or r.get("site_name") or ""
        if name:
            parts.append(f'"{name}"')
        for k, v in r.items():
            if k in _SKIP or k in {"property_name", "site_name"}:
                continue
            if v is None or v == "":
                continue
            parts.append(f"{k}={v!r}")
        rows.append("  " + " | ".join(parts))

    records_text = "\n".join(rows)
    user_msg = (
        f"Records ({len(records)} total):\n{records_text}\n\n"
        f"Analyst instruction: {instruction}"
    )

    print(f"  [GPT Refine] Sending {len(records)} records to {gpt_model} ...")
    resp = client.chat.completions.create(
        model=gpt_model,
        messages=[
            {"role": "system", "content": _GPT_REFINE_SYSTEM},
            {"role": "user",   "content": user_msg},
        ],
        tools=_GPT_REFINE_TOOLS,
        tool_choice="required",
        temperature=0,
    )

    tool_calls = resp.choices[0].message.tool_calls
    if not tool_calls:
        print("  [GPT Refine] No tool call returned — keeping all records")
        return records

    tc      = tool_calls[0]
    fn_name = tc.function.name
    fn_args = json.loads(tc.function.arguments)
    reason  = fn_args.get("reason", "")

    if fn_name == "filter_records":
        to_remove = {int(x) for x in fn_args.get("indices_to_remove", [])}
        removed_names = [
            str(records[i - 1].get("property_name")
                or records[i - 1].get("site_name") or "?")
            for i in sorted(to_remove) if 1 <= i <= len(records)
        ]
        result = [r for i, r in enumerate(records, 1) if i not in to_remove]
        print(f"  [GPT Refine] Removed {len(to_remove)}: {removed_names}")
        print(f"  [GPT Refine] {reason}")
        print(f"  [GPT Refine] {len(records)} → {len(result)} records")
        return result

    if fn_name == "keep_records":
        to_keep = [int(x) for x in fn_args.get("indices_to_keep", [])]
        result  = [records[i - 1] for i in to_keep if 1 <= i <= len(records)]
        print(f"  [GPT Refine] Kept {len(result)} records in specified order")
        print(f"  [GPT Refine] {reason}")
        print(f"  [GPT Refine] {len(records)} → {len(result)} records")
        return result

    # no_change
    print(f"  [GPT Refine] No change. {reason}")
    return records


# ── apply_refinement (router used by scan pipeline) ───────────────────────────

def apply_refinement(records: list, instructions: str,
                     llm_cfg: dict,
                     base_url: str = None, model: str = None) -> list:
    """
    Apply analyst free-text instructions to filter/reorder comparable records.

    Routes to GPT function-calling agent when llm_cfg["provider"] == "openai",
    otherwise falls back to the Ollama multi-turn agent loop.
    """
    provider = (llm_cfg or {}).get("provider", "ollama")
    if provider == "openai":
        return run_agent_loop_gpt(instructions, records, llm_cfg)
    # Ollama path
    _base  = base_url  or (llm_cfg or {}).get("ollama", {}).get("base_url", "http://localhost:11434")
    _model = model     or (llm_cfg or {}).get("ollama", {}).get("model",    "qwen2.5:3b")
    return run_agent_loop(instructions, records, _base, _model)
