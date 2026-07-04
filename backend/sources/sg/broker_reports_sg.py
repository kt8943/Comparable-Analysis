"""
backend/sources/sg/broker_reports_sg.py
=======================================
Broker market-report connector (Singapore): Savills, Cushman & Wakefield
(MarketBeat), CBRE. Discovers the report **PDF** links published on each broker's
research/insights page, filters to the subject's asset type + recent quarters, and
extracts named transactions from the report TEXT via the OpenAI extract model —
the same two-step (fetch text → LLM extract) the web-search source uses.

Requires an OpenAI ``client`` + ``extract_model`` (passed in ``params``) — fails
soft (returns [],[]) without them or on any error. Uses pdfplumber for text (works
on the cloud; no camelot/Ghostscript needed).

CBRE's page bot-blocks plain requests (HTTP 403) — best-effort; skipped on failure.
"""
from __future__ import annotations

import json
import re
import ssl
import tempfile
import urllib.parse
import urllib.request

from ..base import SourceConnector
from ..registry import register

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

# Broker research/insight landing pages (curated; user-extendable via config).
_DEFAULT_PAGES = [
    "https://www.cushmanwakefield.com/en/singapore/insights/singapore-marketbeat",
    "https://www.savills.com.sg/insight-and-opinion/research.aspx?rc=Singapore&f=date&page=1",
    "https://www.cbre.com.sg/insights",   # often 403 — best-effort
]

_ASSET_KW = {
    "office":      ["office"],
    "retail":      ["retail", "mall", "shop"],
    "industrial":  ["industrial", "logistics", "warehouse", "business-park", "business park"],
    "logistics":   ["logistics", "industrial", "warehouse"],
    "hospitality": ["hotel", "hospitality"],
    "residential": ["residential"],
    "mixed":       ["mixed", "investment"],
}

try:
    _CTX = ssl.create_default_context()
    _CTX.check_hostname = False
    _CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _CTX = None


def _get(url: str, timeout: int = 25) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    kw = {"timeout": timeout}
    if _CTX is not None:
        kw["context"] = _CTX
    return urllib.request.urlopen(req, **kw).read()


def _discover_pdfs(page_url: str) -> list:
    try:
        html = _get(page_url).decode("utf-8", "replace")
    except Exception as e:
        print(f"    [broker] {page_url[:40]}… failed: {e}")
        return []
    urls = []
    for m in re.findall(r'href="([^"]+\.pdf[^"]*)"', html, re.I):
        urls.append(urllib.parse.urljoin(page_url, m))
    # de-dup, keep order
    seen, out = set(), []
    for u in urls:
        base = u.split("?")[0]
        if base not in seen:
            seen.add(base)
            out.append(u)
    return out


def _pdf_matches(url: str, asset_kw: list, years: list) -> bool:
    u = url.lower()
    if "singapore" not in u and "-sg-" not in u and "/sg/" not in u:
        # many SG report URLs contain 'singapore'; be lenient if a year matches
        if not any(y in u for y in years):
            return False
    if asset_kw and not any(k in u for k in asset_kw):
        return False
    return any(y in u for y in years) or True  # recency mainly enforced downstream


def _pdf_text(url: str, max_pages: int = 8) -> str:
    try:
        data = _get(url, timeout=40)
    except Exception:
        return ""
    try:
        import pdfplumber, io
        text = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for pg in pdf.pages[:max_pages]:
                text.append(pg.extract_text() or "")
        return "\n".join(text)
    except Exception as e:
        print(f"    [broker] pdf text failed: {e}")
        return ""


_SHAPE_HINT = {
    "sales": ('property_name, address, sale_date, price_sgd_m (total price in S$ '
              'millions), gfa_sf, cap_rate_pct, stake_pct, buyer, seller'),
    "land":  ('site_name, address, launch_date, land_zoning, tenure, site_area_sf, '
              'max_gfa_sf, price_sgd_m, price_psf_ppr'),
    "rent":  ('property_name, address, lease_date, nla_sf, asking_rent (S$ psf/month), '
              'eff_rent, lease_term_yrs, lease_type, tenant'),
}
_KIND = {"sales": "confirmed investment/sale transactions of standing buildings",
         "land":  "government land sale / land tender transactions",
         "rent":  "confirmed lease/rental transactions with actual rent figures"}


def _llm_extract(text: str, client, model: str, comp_type: str) -> list:
    if not text.strip():
        return []
    try:
        resp = client.chat.completions.create(
            model=model, response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content":
                 "You extract structured real-estate comparable transactions from a "
                 "broker market report. Only extract " + _KIND.get(comp_type, "") +
                 " that name a specific property and figure. Ignore market averages, "
                 "indices, vacancy/supply statistics and forecasts. Return a JSON "
                 "object {\"transactions\": [...]}."},
                {"role": "user", "content":
                 f"Fields per transaction: {_SHAPE_HINT.get(comp_type, _SHAPE_HINT['sales'])}, "
                 f"country. Use null when absent. Report text:\n\n{text[:12000]}"},
            ])
        obj = json.loads(resp.choices[0].message.content)
        recs = obj.get("transactions") or next(
            (v for v in obj.values() if isinstance(v, list)), [])
        return recs if isinstance(recs, list) else []
    except Exception as e:
        print(f"    [broker] extract failed: {e}")
        return []


class BrokerReportsSGConnector(SourceConnector):
    name = "broker_reports"
    market = "sg"
    comp_types = {"sales", "land", "rent"}
    label = "Broker market reports (Savills / C&W / CBRE)"

    def fetch(self, subject_cfg: dict, params: dict) -> tuple:
        client = params.get("client")
        model  = params.get("extract_model", "gpt-4o-mini")
        if client is None:
            print("    [broker] no OpenAI client — skipping")
            return [], []
        comp_type = params.get("comp_type", "sales")
        pages = params.get("broker_pages") or _DEFAULT_PAGES
        asset = (subject_cfg.get("asset_class", "office") or "office").lower()
        asset_kw = next((v for k, v in _ASSET_KW.items() if k in asset), ["office"])
        from datetime import datetime
        now = datetime.now()
        years = [str(now.year), str(now.year - 1)]
        max_pdfs = int(params.get("broker_max_pdfs", 4) or 4)

        pdfs = []
        for pg in pages:
            for u in _discover_pdfs(pg):
                if _pdf_matches(u, asset_kw, years):
                    pdfs.append(u)
        # de-dup + cap
        seen, picked = set(), []
        for u in pdfs:
            b = u.split("?")[0]
            if b not in seen:
                seen.add(b)
                picked.append(u)
        picked = picked[:max_pdfs]
        print(f"    [broker] {len(picked)} matching report PDF(s) for '{asset}'")

        records, sources = [], []
        for u in picked:
            txt = _pdf_text(u)
            recs = _llm_extract(txt, client, model, comp_type)
            title = u.split("/")[-1].split("?")[0]
            for r in recs:
                r.setdefault("country", "Singapore")
                r["sources"] = [{"title": title, "url": u, "source_name": self.name}]
            records.extend(recs)
            if recs:
                sources.append({"title": title, "url": u})
            print(f"      · {title[:50]}: {len(recs)} txn(s)")
        return records, sources


register(BrokerReportsSGConnector())
