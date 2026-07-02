"""OneMap token auth — fetch/refresh the OneMap API token from email+password.

OneMap API tokens expire (~3 days). The public geocoding/search endpoint needs
no token, but the Theme service (amenity POIs used for the Location proximity
score) and URA zone lookups do.

This reads ``onemap_email`` / ``onemap_password`` from configs/shared_settings.json,
fetches a fresh token via the getToken endpoint, and caches it back into
shared_settings (``onemap_token`` + ``onemap_token_expiry``) so repeated calls
reuse it until it is close to expiry. Falls back to any manually-pasted
``onemap_token`` when no email/password is set.
"""

import json
import ssl
import time
import urllib.request
from pathlib import Path

_TOKEN_URL = "https://www.onemap.gov.sg/api/auth/post/getToken"
_SS_PATH = Path(__file__).resolve().parents[2] / "configs" / "shared_settings.json"

try:
    _CTX = ssl.create_default_context()
    _CTX.check_hostname = False
    _CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _CTX = None


def _load() -> dict:
    try:
        return json.loads(_SS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save(d: dict) -> None:
    try:
        _SS_PATH.write_text(json.dumps(d, indent=2), encoding="utf-8")
    except Exception:
        pass


def get_onemap_token() -> str:
    """Return a valid OneMap token (auto-refreshing via email/password if needed)."""
    ss = _load()
    tok = str(ss.get("onemap_token", "")).strip()
    exp = ss.get("onemap_token_expiry", 0)
    now = int(time.time())

    # Reuse the cached token while it still has comfortable life left (>1h).
    try:
        if tok and float(exp) - now > 3600:
            return tok
    except (TypeError, ValueError):
        pass

    email = str(ss.get("onemap_email", "")).strip()
    pw = str(ss.get("onemap_password", "")).strip()
    if not (email and pw):
        return tok  # no credentials — fall back to any manual token (may be empty/expired)

    try:
        body = json.dumps({"email": email, "password": pw}).encode()
        req = urllib.request.Request(
            _TOKEN_URL, data=body,
            headers={"Content-Type": "application/json"}, method="POST",
        )
        kw = {"timeout": 20}
        if _CTX is not None:
            kw["context"] = _CTX
        resp = json.loads(urllib.request.urlopen(req, **kw).read())
        new_tok = resp.get("access_token", "")
        new_exp = int(resp.get("expiry_timestamp", now + 2 * 24 * 3600))
        if new_tok:
            ss["onemap_token"] = new_tok
            ss["onemap_token_expiry"] = new_exp
            _save(ss)
            return new_tok
    except Exception as e:
        print(f"  [onemap_auth] token fetch failed: {e}")
    return tok
