"""
backend/sources/sg/ura_api.py
=============================
URA **Data Service API** auth helper (Access-Key → daily Token → invoke a service).

This is the keyed URA API (register at https://eservice.ura.gov.sg/maps/api/reg.html).
The Access Key is emailed on activation; it is used to mint a **daily token**, which is
then sent with every data call:

    token = get_ura_token(access_key)                      # once per day (cached)
    rows  = ura_invoke("PMI_Resi_Transaction", access_key, {"batch": 1})

The Access Key is read (in priority order) from:
    1. an explicit argument
    2. shared_settings.json  →  "ura_access_key"      (on-prem, plaintext, gitignored)
    3. env  URA_ACCESS_KEY                             (cloud: Streamlit secret)

NOTE — the classic Data Service API currently exposes RESIDENTIAL PMI, car parks and
planning decisions only (services: PMI_Resi_Transaction, PMI_Resi_Rental,
PMI_Resi_Rental_Median, PMI_Resi_Developer_Sales, PMI_Resi_Pipeline, EAU_Appr_Resi_Use,
Planning_Decision, Car_Park_*). It does NOT include GLS or commercial transactions.

Fails soft: every function returns None / [] and prints a note on any error.

Quick check (key stays local):
    python backend/sources/sg/ura_api.py YOUR_ACCESS_KEY
"""
from __future__ import annotations

import json
import os
import ssl
import time
import urllib.parse
import urllib.request
from pathlib import Path

_TOKEN_URL  = "https://eservice.ura.gov.sg/uraDataService/insertNewToken/v1"
_INVOKE_URL = "https://eservice.ura.gov.sg/uraDataService/invokeUraDS/v1"
_UA = "Mozilla/5.0 (compatible; pgim-comps/1.0)"
# backend/sources/sg/ura_api.py → parents[3] is the project root (…/PGIM)
_ROOT = Path(__file__).resolve().parents[3]
_TOKEN_CACHE = _ROOT / "output" / "_ura_token.json"

try:
    _CTX = ssl.create_default_context()
    _CTX.check_hostname = False
    _CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _CTX = None


def get_access_key(access_key: str = "") -> str:
    if access_key:
        return access_key.strip()
    try:
        p = _ROOT / "configs" / "shared_settings.json"
        if p.exists():
            k = (json.loads(p.read_text(encoding="utf-8")) or {}).get("ura_access_key", "")
            if k:
                return k.strip()
    except Exception:
        pass
    return (os.environ.get("URA_ACCESS_KEY", "") or "").strip()


def _get(url: str, headers: dict, timeout: int = 30) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": _UA, **headers})
    kw = {"timeout": timeout}
    if _CTX is not None:
        kw["context"] = _CTX
    raw = urllib.request.urlopen(req, **kw).read().decode("utf-8", "replace")
    return json.loads(raw)


def get_ura_token(access_key: str = "", force: bool = False) -> str | None:
    """Return today's token (cached to disk per day). None if the key is missing/bad."""
    key = get_access_key(access_key)
    if not key:
        print("    [ura_api] no Access Key (set ura_access_key in Shared Settings / "
              "URA_ACCESS_KEY secret)")
        return None
    today = time.strftime("%Y-%m-%d")
    if not force and _TOKEN_CACHE.exists():
        try:
            c = json.loads(_TOKEN_CACHE.read_text(encoding="utf-8"))
            if c.get("date") == today and c.get("token"):
                return c["token"]
        except Exception:
            pass
    try:
        resp = _get(_TOKEN_URL, {"AccessKey": key})
        token = resp.get("Result") or ""
        if not token:
            print(f"    [ura_api] token request returned no Result: {resp}")
            return None
        try:
            _TOKEN_CACHE.parent.mkdir(parents=True, exist_ok=True)
            _TOKEN_CACHE.write_text(json.dumps({"date": today, "token": token}),
                                    encoding="utf-8")
        except Exception:
            pass
        return token
    except Exception as e:
        print(f"    [ura_api] token request failed: {e}")
        return None


def ura_invoke(service: str, access_key: str = "", params: dict | None = None) -> list:
    """Call a URA data service; returns its `Result` list (or [] on any error)."""
    key = get_access_key(access_key)
    token = get_ura_token(key)
    if not (key and token):
        return []
    qs = {"service": service, **(params or {})}
    url = f"{_INVOKE_URL}?{urllib.parse.urlencode(qs)}"
    try:
        resp = _get(url, {"AccessKey": key, "Token": token})
        res = resp.get("Result")
        if isinstance(res, list):
            return res
        return [res] if isinstance(res, dict) else []
    except Exception as e:
        print(f"    [ura_api] invoke {service} failed: {e}")
        return []


if __name__ == "__main__":   # pragma: no cover — manual key check
    import sys
    key = sys.argv[1] if len(sys.argv) > 1 else get_access_key()
    print("Access key:", (key[:6] + "…") if key else "(none)")
    tok = get_ura_token(key, force=True)
    print("Token:", (tok[:12] + "…") if tok else "FAILED")
    if tok:
        rows = ura_invoke("PMI_Resi_Transaction", key, {"batch": 1})
        print(f"PMI_Resi_Transaction batch 1: {len(rows)} project group(s)")
        if rows:
            r0 = rows[0]
            print("  sample keys:", list(r0.keys())[:12])
            print("  project:", r0.get("project"), "| street:", r0.get("street"),
                  "| #txns:", len(r0.get("transaction", []) or []))
