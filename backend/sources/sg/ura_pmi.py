"""
backend/sources/sg/ura_pmi.py
=============================
URA Property Market Information (PMI) — **Commercial Transaction Search**.

FREE / PUBLIC (no REALIS subscription). The commercial search is a server-rendered
Spring form, so it's scraped with plain HTTP — **no browser needed**, works on the
cloud server too:
    1. GET the page   → session cookie (JSESSIONID) + CSRF token
    2. POST the form  → HTML results table (propertyTypeNo, sale year/month range)
    3. parse the 13-column table into SALES records

Property types: 1=Retail, 2=Office, 3=Shophouses.
Columns: Project Name | Street Name | Property Type | Transacted Price ($) |
Area (SQFT) | Unit Price ($ PSF) | Sale Date | Type of Area | Area (SQM) |
Unit Price ($ PSM) | Tenure | Postal District | Floor Level.

Returns records in the SALES raw shape; the shared pipeline geocodes (by street) and
distance-filters to the subject. Fails soft (returns [], []) on any error.
"""
from __future__ import annotations

import http.cookiejar
import re
import ssl
import urllib.parse
import urllib.request
from datetime import datetime

from ..base import SourceConnector, num
from ..registry import register

_BASE = ("https://eservice.ura.gov.sg/property-market-information/"
         "pmiCommercialTransactionSearch")
_UA = "Mozilla/5.0 (compatible; pgim-comps/1.0)"

try:
    _CTX = ssl.create_default_context()
    _CTX.check_hostname = False
    _CTX.verify_mode = ssl.CERT_NONE
except Exception:
    _CTX = None

_TYPE_NO = {"retail": "1", "mall": "1", "shop": "1",
            "office": "2",
            "shophouse": "3", "shophouses": "3"}


def _asset_type_no(asset_class: str) -> str:
    a = (asset_class or "").lower()
    for k, v in _TYPE_NO.items():
        if k in a:
            return v
    return "2"   # default → Office


def _tenure_years(tenure: str):
    t = (tenure or "").lower()
    if "freehold" in t:
        return None
    m = re.search(r"(\d{2,4})\s*(?:yr|year)", t)
    return int(m.group(1)) if m else None


def _opener():
    cj = http.cookiejar.CookieJar()
    handlers = [urllib.request.HTTPCookieProcessor(cj)]
    if _CTX is not None:
        handlers.append(urllib.request.HTTPSHandler(context=_CTX))
    op = urllib.request.build_opener(*handlers)
    op.addheaders = [("User-Agent", _UA)]
    return op


class URAPMICommercialConnector(SourceConnector):
    name = "ura_pmi"
    market = "sg"
    comp_types = {"sales"}
    label = "URA PMI (commercial transactions)"

    def fetch(self, subject_cfg: dict, params: dict) -> tuple:
        try:
            op = _opener()
            page = op.open(_BASE, timeout=30).read().decode("utf-8", "replace")
            m = re.search(r'name="_csrf"\s+value="([^"]+)"', page)
            if not m:
                print("    [ura_pmi] no CSRF token — skipping")
                return [], []
            csrf = m.group(1)

            yrs = int(params.get("years_back", 2) or 2)
            now = datetime.now()
            y_from = now.year - yrs
            type_no = _asset_type_no(subject_cfg.get("asset_class", ""))
            max_rows = int(params.get("ura_max_rows", 60) or 60)

            form = {
                "_csrf": csrf,
                "propertyTypeNo": type_no,
                "saleYearFrom": str(y_from), "saleMonthFrom": "1",
                "saleYearTo": str(now.year), "saleMonthTo": str(now.month),
                "resultPerPage": str(max_rows), "displayResult": "true",
            }
            req = urllib.request.Request(
                _BASE, data=urllib.parse.urlencode(form).encode(),
                headers={"User-Agent": _UA,
                         "Content-Type": "application/x-www-form-urlencoded"})
            html = op.open(req, timeout=45).read().decode("utf-8", "replace")
        except Exception as e:
            print(f"    [ura_pmi] fetch failed: {e}")
            return [], []

        records = []
        for row in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
            tds = [re.sub(r"<[^>]+>", "", c).strip()
                   for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
            if len(tds) < 13:
                continue
            proj, street, ptype, price, area_sf, _psf, sdate, area_type, \
                _area_sm, _psm, tenure, _district, _floor = tds[:13]
            price_v = num(price)
            if not proj or price_v is None:
                continue
            records.append({
                "property_name": proj,
                "address":       f"{street}, Singapore" if street else "",
                "sale_date":     (sdate or "").strip(),
                "price_sgd_m":   round(price_v / 1_000_000, 4),
                "gfa_sf":        num(area_sf),
                "remaining_yrs": _tenure_years(tenure),
                "cap_rate_pct":  None,
                "stake_pct":     100,
                "sale_type":     (area_type or "Strata"),
                "asset_type":    ptype,
                "land_zoning":   "",
                "buyer":         "",
                "seller":        "",
                "country":       "Singapore",
            })
        print(f"    [ura_pmi] {len(records)} commercial transactions "
              f"(type {type_no}, {y_from}–{now.year})")
        return records, [{"title": self.label, "url": _BASE}]


register(URAPMICommercialConnector())
