#!/usr/bin/env python3
"""
run.py
======
Interactive launcher for all PGIM deal scripts.

Usage
-----
    python3 run.py
"""

import json
import os
import subprocess
import sys
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────────────
# MENU CONFIG
# ─────────────────────────────────────────────────────────────────────────────

ACTIONS = [
    {
        "label": "Asset Sales Comps — Online Search",
        "script": "backend/search_online_sales_comps.py",
        "extra_flags": ["--map"],
        "description": "Search the web for asset transaction comparables + map",
    },
    {
        "label": "Asset Sales Comps — From Input Excel",
        "script": "backend/scan_input_sales_comps.py",
        "extra_flags": [],
        "description": "Generate asset sales comps table from a provided Excel file",
    },
    {
        "label": "Rent Comps — Online Search",
        "script": "backend/search_online_rent_comps.py",
        "extra_flags": ["--map"],
        "description": "Search the web for rental comparables + map",
    },
    {
        "label": "Rent Comps — From Input Excel",
        "script": "backend/scan_input_rent_comps.py",
        "extra_flags": ["--map"],
        "description": "Read a provided rent comps Excel, classify via Ollama + map",
    },
    {
        "label": "Land Sales Comps — From Input Excel",
        "script": "backend/scan_input_land_comps.py",
        "extra_flags": ["--map"],
        "description": "Read a provided land sales Excel, classify via Ollama + map",
    },
    {
        "label": "Land Sales Comps — Online Search",
        "script": "backend/search_online_land_comps.py",
        "extra_flags": ["--map"],
        "description": "Search the web for land sale comparables (GLS, en bloc) + map",
    },
    {
        "label": "Investment Rationale — Generate from Market Reports",
        "script": "backend/generate_investment_rationale.py",
        "extra_flags": [],
        "description": "Read market report PDFs + deal config → write 3-section investment rationale",
        "investment_rationale": True,   # triggers special report-selection flow
    },
    {
        "label": "New Deal Setup",
        "script": "backend/new_deal.py",
        "extra_flags": [],
        "description": "Wizard to create a new deal config",
        "no_config": True,
    },
]


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _load_deal_name(config_path: str) -> str:
    try:
        with open(config_path) as f:
            cfg = json.load(f)
        return cfg.get("subject_property", {}).get("deal_name", config_path)
    except Exception:
        return config_path


def _list_configs() -> list:
    """Return sorted list of (display_name, config_path) from configs/ folder."""
    configs_dir = Path("configs")
    if not configs_dir.exists():
        return []
    files = sorted(configs_dir.glob("deal_config*.json"))
    result = []
    for f in files:
        name = _load_deal_name(str(f))
        result.append((name, str(f)))
    return result


def _pick(prompt: str, options: list, key=lambda x: x) -> object:
    """Show a numbered menu and return the chosen item."""
    print(f"\n  {prompt}")
    print("  " + "─" * 50)
    for i, opt in enumerate(options, 1):
        print(f"  {i:>2}.  {key(opt)}")
    print()
    while True:
        raw = input("  Enter number: ").strip()
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(options):
                return options[idx]
        except ValueError:
            pass
        print("  (invalid — please enter a number from the list)")


def _confirm(prompt: str, default: bool = True) -> bool:
    hint = "[Y/n]" if default else "[y/N]"
    raw = input(f"  {prompt} {hint}: ").strip().lower()
    if raw == "":
        return default
    return raw in ("y", "yes")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("\n╔══════════════════════════════════════════════╗")
    print("║        PGIM  —  Deal Analysis Runner         ║")
    print("╚══════════════════════════════════════════════╝")

    # ── 1. Pick action ────────────────────────────────────────────────────────
    action = _pick("What would you like to run?",
                   ACTIONS, key=lambda a: f"{a['label']:<40}  {a['description']}")

    # ── 2. New deal wizard: no config needed ──────────────────────────────────
    if action.get("no_config"):
        print(f"\n  Running: python3 {action['script']}\n")
        subprocess.run([sys.executable, action["script"]], check=False)
        return

    # ── 3. Pick deal ──────────────────────────────────────────────────────────
    configs = _list_configs()
    if not configs:
        print("\n  No deal configs found in configs/. Run 'New Deal Setup' first.")
        return

    deal_name, config_path = _pick("Which deal?",
                                    configs, key=lambda x: x[0])

    # ── 4. Optional flags ─────────────────────────────────────────────────────
    flags = list(action["extra_flags"])

    # --refresh for online search scripts
    if "search_online" in action["script"]:
        if _confirm("Force refresh (ignore cache)?", default=False):
            flags.append("--refresh")

    # ── 4b. Investment rationale: pick which reports to include ───────────────
    if action.get("investment_rationale"):
        reports_dir = Path("Input_files/market_reports")
        available   = sorted(reports_dir.glob("*.pdf")) if reports_dir.exists() else []
        if available:
            print(f"\n  Found {len(available)} market report(s) in {reports_dir}/")
            use_all = _confirm("Use ALL available reports?", default=True)
            if use_all:
                chosen_reports = [str(p) for p in available]
            else:
                chosen_reports = []
                for i, rpt in enumerate(available, 1):
                    if _confirm(f"  Include  {rpt.name}?", default=True):
                        chosen_reports.append(str(rpt))
            if chosen_reports:
                flags += ["--reports"] + chosen_reports
        else:
            print(f"\n  (No PDF reports found in {reports_dir}/ — "
                  "rationale will use deal config only)")

        if _confirm("Force re-extract reports (ignore cache)?", default=False):
            flags.append("--refresh")

    # ── 5. Build + show command ───────────────────────────────────────────────
    cmd = [sys.executable, action["script"], "--config", config_path] + flags
    print(f"\n  ┌─ Running ──────────────────────────────────────────────")
    print(f"  │  {' '.join(cmd)}")
    print(f"  └────────────────────────────────────────────────────────\n")

    subprocess.run(cmd, check=False)


if __name__ == "__main__":
    main()
