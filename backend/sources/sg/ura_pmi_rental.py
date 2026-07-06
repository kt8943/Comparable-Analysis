"""
backend/sources/sg/ura_pmi_rental.py
====================================
URA Property Market Information (PMI) — **Commercial Rental Statistics by Street**.

FREE / PUBLIC (no REALIS subscription). Like the commercial *transaction* search, the
rental eService is a server-rendered Spring form scraped with plain HTTP — **no browser
needed**, works on the cloud too:
    1. GET the page  → session cookie (JSESSIONID) + CSRF token + quarter list
    2. POST the form → HTML table of median rents ($ PSM/month) per street
    3. parse each street's Retail / Office 25th·Median·75th percentiles → RENT records

IMPORTANT — this is an **aggregate benchmark by street**, NOT individual named lease
deals: each row is a street with the median asking rent for that quarter. It's surfaced
as a grounded local rent benchmark near the subject (one pseudo-comp per nearby street),
which the shared pipeline geocodes (by street) and distance-filters to the subject.

Columns (per data row): Reference Quarter | Street | Retail 25th | Retail Median |
Retail 75th | Office 25th | Office Median | Office 75th.  Values may be "-".

Returns records in the RENT raw shape. Fails soft (returns [], []) on any error.
"""
from __future__ import annotations

import http.cookiejar
import re
import ssl
import urllib.parse
import urllib.request

from ..base import SourceConnector, num
from ..registry import register

_BASE = ("https://eservice.ura.gov.sg/property-market-information/"
         "pmiCommercialRentalStatsByStreet")
_UA = "Mozilla/5.0 (compatible; pgim-comps/1.0)"
_SQM_PER_SF = 10.7639   # 1 sqm = 10.7639 sf → $PSM ÷ this = $PSF

try:
    _CTX = ssl.create_default_context()
    _CTX.check_hostname = False
    _CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _CTX = None


def _is_retail(asset_class: str) -> bool:
    a = (asset_class or "").lower()
    return any(k in a for k in ("retail", "mall", "shop"))


def _opener():
    cj = http.cookiejar.CookieJar()
    handlers = [urllib.request.HTTPCookieProcessor(cj)]
    if _CTX is not None:
        handlers.append(urllib.request.HTTPSHandler(context=_CTX))
    op = urllib.request.build_opener(*handlers)
    op.addheaders = [("User-Agent", _UA)]
    return op


def _title(s: str) -> str:
    return " ".join(w.capitalize() for w in (s or "").split())


class URAPMIRentalConnector(SourceConnector):
    name = "ura_pmi_rental"
    market = "sg"
    comp_types = {"rent"}
    label = "URA PMI (commercial rents by street)"

    def fetch(self, subject_cfg: dict, params: dict) -> tuple:
        try:
            op = _opener()
            page = op.open(_BASE, timeout=30).read().decode("utf-8", "replace")
            m = re.search(r'name="_csrf"\s+value="([^"]+)"', page)
            if not m:
                print("    [ura_pmi_rental] no CSRF token — skipping")
                return [], []
            csrf = m.group(1)
            quarters = re.findall(r'<option value="(\d{4}Q\d)"', page)
            if not quarters:
                print("    [ura_pmi_rental] no quarters listed — skipping")
                return [], []
            quarter = quarters[0]   # most recent quarter (current benchmark)
            max_rows = int(params.get("ura_max_rows", 300) or 300)

            form = {
                "_csrf": csrf, "refQuarter": quarter, "selectedQuarter": quarter,
                "resultPerPage": str(max_rows), "displayResult": "true",
            }
            req = urllib.request.Request(
                _BASE, data=urllib.parse.urlencode(form).encode(),
                headers={"User-Agent": _UA,
                         "Content-Type": "application/x-www-form-urlencoded"})
            html = op.open(req, timeout=45).read().decode("utf-8", "replace")
        except Exception as e:
            print(f"    [ura_pmi_rental] fetch failed: {e}")
            return [], []

        retail = _is_retail(subject_cfg.get("asset_class", ""))
        sqm_unit = (subject_cfg.get("gfa_unit", "sf").lower() == "sqm")
        seg = "retail" if retail else "office"
        # median column index: [q, street, r25, rMed, r75, o25, oMed, o75]
        med_idx = 3 if retail else 6

        records = []
        for row in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
            tds = [re.sub(r"<[^>]+>", "", c).strip()
                   for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
            if len(tds) < 8 or not re.match(r"\d{4}Q\d", tds[0]):
                continue
            street = tds[1]
            med_psm = num(tds[med_idx])
            if not street or med_psm is None:
                continue   # no data for this street/segment this quarter
            rent = med_psm if sqm_unit else round(med_psm / _SQM_PER_SF, 2)
            records.append({
                "property_name":  _title(street),
                "address":        f"{street}, Singapore",
                "lease_date":     tds[0],
                "asking_rent":    rent,
                "eff_rent":       None,
                "nla_sf":         None,
                "lease_term_yrs": None,
                "lease_type":     f"URA median {seg} rent",
                "asset_type":     seg.title(),
                "tenant":         "",
                "country":        "Singapore",
            })
        unit = "PSM" if sqm_unit else "PSF"
        print(f"    [ura_pmi_rental] {len(records)} street medians "
              f"({seg} {quarter}, ${unit}/mo)")
        return records, [{"title": f"{self.label} — {quarter}", "url": _BASE}]


register(URAPMIRentalConnector())
