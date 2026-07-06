"""
backend/sources/sg/broker_reports_sg.py
=======================================
Broker market-report connector (Singapore): Savills, Cushman & Wakefield
(MarketBeat), CBRE. Reads each broker's research/insights page, lists every report
(**title + first-paragraph snippet + link**), lets the LLM judge which are relevant
to the subject (asset type / location / transaction type / recency), then downloads
the chosen reports (PDF or article page), extracts the text, and LLM-extracts named
transactions — flowing through the same dedup → geocode → comparability → Excel/map
pipeline, citing the source report.

Requires an OpenAI ``client`` + ``extract_model`` (passed in ``params``). Fails soft
(returns [],[]) without them or on any error. Uses pdfplumber for PDF text (works on
the cloud; no camelot/Ghostscript needed).
"""
from __future__ import annotations

import io
import json
import re
import ssl
import urllib.parse
import urllib.request
from datetime import datetime

from ..base import SourceConnector
from ..registry import register

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

_DEFAULT_PAGES = [
    "https://www.cushmanwakefield.com/en/singapore/insights/singapore-marketbeat",
    "https://www.savills.com.sg/insight-and-opinion/research.aspx?rc=Singapore&f=date&page=1",
    "https://www.cbre.com.sg/insights",
]

# Only anchors whose URL looks like a report/insight are considered candidates.
_REPORT_HINTS = ("research", "insight", "marketbeat", "market-report",
                 "/report", "outlook", "briefing")

try:
    _CTX = ssl.create_default_context()
    _CTX.check_hostname = False
    _CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _CTX = None


def _get(url: str, timeout: int = 30) -> bytes:
    req = urllib.request.Request(url, headers={
        "User-Agent": _UA,
        "Accept": "text/html,application/pdf,*/*",
        "Accept-Language": "en-SG,en;q=0.9",
    })
    kw = {"timeout": timeout}
    if _CTX is not None:
        kw["context"] = _CTX
    return urllib.request.urlopen(req, **kw).read()


def _strip(html: str) -> str:
    html = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html)).strip()


def _discover_report_links(page_url: str) -> list:
    """Return [{title, url, snippet}] for report-like links on a broker page."""
    try:
        html = _get(page_url).decode("utf-8", "replace")
    except Exception as e:
        print(f"    [broker] {page_url.split('//')[-1][:35]}… fetch failed: {e}")
        return []
    items, seen = [], set()
    for m in re.finditer(r'<a\b[^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.S | re.I):
        href, inner = m.group(1), m.group(2)
        url = urllib.parse.urljoin(page_url, href)
        low = url.lower().split("?")[0]
        if not (low.endswith(".pdf") or any(k in low for k in _REPORT_HINTS)):
            continue
        title = _strip(inner)
        if len(title) < 6:
            continue
        base = url.split("?")[0]
        if base in seen:
            continue
        seen.add(base)
        snippet = _strip(html[m.end():m.end() + 900])[:280]   # ~first paragraph after link
        items.append({"title": title[:140], "url": url, "snippet": snippet})
    return items[:40]   # cap candidates for the relevance prompt


_KIND = {"sales": "named investment / sale transactions of buildings",
         "land":  "government land sale / land tender transactions",
         "rent":  "named lease / rental transactions with rent figures"}

# What report TYPES actually list the named deals we want (vs. macro commentary).
# Used to steer the relevance picker toward transaction-bearing reports.
_PREFER = {
    "sales": ("Strongly PREFER 'Capital Markets', 'Investment', 'Investment Sales', "
              "'Transactions' or 'Deals' reports — those list named building sales with "
              "prices. AVOID pure market-outlook / forecast / vacancy / rent-index / "
              "leasing / occupier reports (they only give aggregate statistics, no named "
              "deals)."),
    "land":  ("Strongly PREFER 'Government Land Sales', 'GLS', 'land tender', "
              "'development site' or 'Capital Markets' reports. AVOID generic outlook / "
              "occupier / leasing reports."),
    "rent":  ("Strongly PREFER 'Leasing', 'Occupier', 'Office/Retail Leasing' reports "
              "that cite named lease deals. AVOID capital-markets / investment-sale and "
              "pure rent-index outlook reports."),
}


def _llm_pick_relevant(items, subject_cfg, comp_type, client, model, max_n):
    asset = subject_cfg.get("asset_class", "office")
    country = subject_cfg.get("country_name", "Singapore")
    listing = "\n".join(f"{i}. {it['title']} — {it['snippet'][:160]}"
                        for i, it in enumerate(items))
    try:
        resp = client.chat.completions.create(
            model=model, response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content":
                 f"You pick {country} {asset} market reports most likely to contain "
                 f"{_KIND.get(comp_type, '')}. {_PREFER.get(comp_type, '')} "
                 f"REJECT reports for a different asset type than '{asset}' (e.g. skip "
                 f"residential/hotel/industrial reports when the subject is {asset}). "
                 "Prefer the most RECENT quarterly reports. "
                 f"Return JSON {{\"relevant\": [indices]}} — at most {max_n} indices, "
                 "most relevant first. If NONE look like they contain named deals, return "
                 "an empty list. Judge from the title AND snippet."},
                {"role": "user", "content": f"Reports:\n{listing}"}])
        idxs = json.loads(resp.choices[0].message.content).get("relevant", [])
        picked = [items[i] for i in idxs if isinstance(i, int) and 0 <= i < len(items)]
        return picked[:max_n]   # honour an empty pick — don't fall back to noise
    except Exception as e:
        print(f"    [broker] relevance pick failed ({e}); using first {max_n}")
        return items[:max_n]


def _fetch_text(url: str, max_pages: int = 8) -> str:
    try:
        data = _get(url, timeout=45)
    except Exception as e:
        print(f"    [broker] fetch text failed for {url.split('/')[-1][:40]}: {e}")
        return ""
    if url.lower().split("?")[0].endswith(".pdf"):
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(data)) as pdf:
                return "\n".join((p.extract_text() or "") for p in pdf.pages[:max_pages])
        except Exception as e:
            print(f"    [broker] pdf parse failed: {e}")
            return ""
    return _strip(data.decode("utf-8", "replace"))[:12000]


def _llm_extract(text: str, client, model: str, comp_type: str) -> list:
    if not text.strip():
        return []
    shape = {
        "sales": "property_name, address, sale_date, price_sgd_m, gfa_sf, cap_rate_pct, "
                 "stake_pct, buyer, seller",
        "land":  "site_name, address, launch_date, land_zoning, tenure, site_area_sf, "
                 "max_gfa_sf, price_sgd_m, price_psf_ppr",
        "rent":  "property_name, address, lease_date, nla_sf, asking_rent, eff_rent, "
                 "lease_term_yrs, lease_type, tenant",
    }.get(comp_type, "property_name, address, sale_date, price_sgd_m")
    try:
        resp = client.chat.completions.create(
            model=model, response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content":
                 "Extract structured comparable transactions from a broker market "
                 "report. Only " + _KIND.get(comp_type, "") + " that name a specific "
                 "property and figure. Ignore market averages/indices/vacancy/supply/"
                 "forecasts. Return JSON {\"transactions\": [...]}."},
                {"role": "user", "content":
                 f"Fields: {shape}, country. Use null when absent. Text:\n\n{text[:12000]}"}])
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
        model = params.get("extract_model", "gpt-4o-mini")
        if client is None:
            print("    [broker] no OpenAI client — skipping")
            return [], []
        comp_type = params.get("comp_type", "sales")
        pages = params.get("broker_pages") or _DEFAULT_PAGES
        max_reports = int(params.get("broker_max_pdfs", 4) or 4)

        items = []
        for pg in pages:
            items += _discover_report_links(pg)
        if not items:
            return [], []
        picked = _llm_pick_relevant(items, subject_cfg, comp_type, client, model, max_reports)
        print(f"    [broker] {len(items)} candidate reports → {len(picked)} selected")
        if not picked:
            print("      (none looked like they contain named deals — nothing to open)")
            return [], []

        records, sources = [], []
        for it in picked:
            txt = _fetch_text(it["url"])
            n_chars = len(txt.strip())
            recs = _llm_extract(txt, client, model, comp_type)
            for r in recs:
                r.setdefault("country", "Singapore")
                r["sources"] = [{"title": it["title"], "url": it["url"],
                                 "source_name": self.name}]
            records.extend(recs)
            if recs:
                sources.append({"title": it["title"], "url": it["url"]})
            # Diagnostic: distinguish "couldn't fetch/gated" from "read it, no named deals".
            if n_chars < 200:
                _why = "  (no text — fetch failed / gated / not a PDF)"
            elif not recs:
                _why = "  (text read, no named deals)"
            else:
                _why = ""
            print(f"      · {it['title'][:52]}: {n_chars:,} chars → {len(recs)} txn(s){_why}")
        return records, sources


register(BrokerReportsSGConnector())
