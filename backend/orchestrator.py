#!/usr/bin/env python3
"""
orchestrator.py
===============
The deal-analysis **orchestrator**: given a deal config and what inputs are
available, it decides *which agent performs which task with which deterministic
tool*, in what order — and returns that as an explicit, auditable plan.

Why deterministic (not an LLM agent picking steps)?
  The routing here is a *known* path: file-type → scan tool, reports → rationale.
  Per the project principle ("agent where the path is uncertain; deterministic
  tool where the path is known"), the orchestrator is a rule-based coordinator.
  The *uncertain* judgment is delegated downward:
    • comp_classifier      — is this file sales / rent / land?      (keyword+LLM)
    • comp_acquisition_agent — did extraction work; try which source? (verify+reflect)
    • rationale_writer     — the LLM write-up itself.
  So the orchestrator names the agents and tools; the agents own the judgment.

This module is pure/plannable (no Streamlit, no subprocess). The frontend calls
build_plan() to SHOW the plan and to DRIVE execution (it owns the st.* context and
the _run_script / _run_comp_agent tool wrappers).
"""

from __future__ import annotations

# Display order for comp types
COMP_TYPES = ["rent", "sales", "land"]

_TYPE_META = {
    "rent":  {"title": "Rent (leasing) comparables",
              "tool": "scan_input_rent_comps.py",
              "online_tool": "search_online_rent_comps.py",
              "prefix": "Rent_Comps"},
    "sales": {"title": "Asset Sales comparables",
              "tool": "scan_input_sales_comps.py",
              "online_tool": "search_online_sales_comps.py",
              "prefix": "Transaction_Comparables"},
    "land":  {"title": "Land Sales comparables",
              "tool": "scan_input_land_comps.py",
              "online_tool": "search_online_land_comps.py",
              "prefix": "Land_Sale_Comps"},
}

# task → responsible agent (for the "who does what" display)
AGENT_OF = {
    "classify":        "comp_classifier",
    "acquire":         "comp_acquisition_agent",
    "write_rationale": "rationale_writer",
}


def build_plan(*, unclassified: list | None = None,
               comp_inputs: dict | None = None,
               online_flags: dict | None = None,
               has_reports: bool = False) -> list[dict]:
    """Produce the ordered list of steps.

    Parameters
    ----------
    unclassified : list of filenames dropped in the single auto-sort box that still
                   need a type decision. If non-empty → a leading `classify` step.
    comp_inputs  : {type: bool}   — that comp type has saved input files.
    online_flags : {type: bool}   — web fallback allowed for that type.
    has_reports  : market-report PDFs exist → run the rationale writer.

    Each step is a typed dict:
      {task, agent, title, tool, type?, online_tool?, uses_online?, uses_comps?, reason}
    """
    unclassified = unclassified or []
    comp_inputs  = comp_inputs or {}
    online_flags = online_flags or {}
    steps: list[dict] = []

    # 0 · Classify anything dropped in the single comparable box
    if unclassified:
        steps.append({
            "task": "classify", "agent": AGENT_OF["classify"],
            "title": f"Sort {len(unclassified)} uploaded file(s) by comp type",
            "tool": "comp_classifier.py",
            "reason": "one comparable box → route each file to sales / rent / land",
        })

    # 1 · Acquire each comp type that has an input source (files and/or web)
    any_comps = False
    for t in COMP_TYPES:
        has_files = bool(comp_inputs.get(t))
        online    = bool(online_flags.get(t))
        if not (has_files or online):
            continue
        any_comps = True
        meta = _TYPE_META[t]
        if has_files and online:
            reason = "extract from files; web fallback if grounding is weak"
        elif has_files:
            reason = "extract comparables from the uploaded files"
        else:
            reason = "no files — acquire comparables via web search"
        steps.append({
            "task": "acquire", "type": t, "agent": AGENT_OF["acquire"],
            "title": meta["title"], "tool": meta["tool"],
            "online_tool": meta["online_tool"] if online else None,
            "uses_online": online, "reason": reason,
        })

    # 2 · Investment rationale (LLM write-up), enriched with the comps just produced
    if has_reports:
        steps.append({
            "task": "write_rationale", "agent": AGENT_OF["write_rationale"],
            "title": "Investment rationale write-up",
            "tool": "generate_investment_rationale.py",
            "uses_comps": any_comps,
            "reason": ("write the IC memo from the market reports"
                       + (" + comparable pricing evidence" if any_comps else "")),
        })
    return steps


def describe_plan(steps: list[dict]) -> list[str]:
    """Human-readable lines: 'Task — agent · tool (why)'. Drives the UI plan panel."""
    lines = []
    for i, s in enumerate(steps, 1):
        tool = s.get("tool", "")
        if s.get("online_tool"):
            tool += f" (+ {s['online_tool']})"
        lines.append(f"{i}. {s['title']} — {s['agent']} · {tool}  ·  {s['reason']}")
    return lines


def route_files(classifications: list[dict]) -> dict:
    """Group comp_classifier output by type.
    Returns {'sales':[paths], 'rent':[...], 'land':[...], 'unknown':[...]}."""
    out = {"sales": [], "rent": [], "land": [], "unknown": []}
    for c in classifications or []:
        out.setdefault(c.get("type", "unknown"), out["unknown"]).append(
            c.get("path") or c.get("name"))
    return out


if __name__ == "__main__":
    # Illustrative plan
    demo = build_plan(
        unclassified=["a.pdf", "b.pdf"],
        comp_inputs={"sales": True, "land": True},
        online_flags={"rent": True},
        has_reports=True,
    )
    for line in describe_plan(demo):
        print(line)
