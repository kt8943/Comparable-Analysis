"""
tools/calculations.py
=====================
Pure math and parsing utilities shared across all comp pipelines.
No I/O, no LLM calls — safe to import anywhere.

Public API
----------
haversine_km(lon1, lat1, lon2, lat2) -> float
parse_num(value) -> float | None
parse_remaining_yrs(val) -> float | None
parse_sale_date(val) -> str
bala_factor(n) -> float
bala_expr(x_ref) -> str
"""

import datetime
import math
import re
import sys
from pathlib import Path

import openpyxl

# Allow importing from the backend root when this module is run directly
sys.path.insert(0, str(Path(__file__).parent.parent))

# ─────────────────────────────────────────────────────────────────────────────
# BALA TABLE  (loaded once at import time)
# ─────────────────────────────────────────────────────────────────────────────

_BALA_TABLE_PATH = Path(__file__).parent.parent.parent / "Input_files" / "bala_table.xlsx"


def _load_bala_table() -> dict:
    """Load Singapore Bala Table → {remaining_years: pct_of_freehold}."""
    path = _BALA_TABLE_PATH
    if not path.exists():
        raise FileNotFoundError(
            f"Bala Table Excel not found: {path}\n"
            "Please ensure Input_files/bala_table.xlsx exists."
        )
    wb = openpyxl.load_workbook(str(path), data_only=True)
    ws = wb.active
    tbl: dict[int, float] = {}
    for row in ws.iter_rows(min_row=2):
        yrs, pct = row[0].value, row[1].value
        if yrs is None:
            break
        try:
            tbl[int(round(float(yrs)))] = float(pct)
        except (TypeError, ValueError):
            continue
    wb.close()
    if not tbl:
        raise ValueError("bala_table.xlsx appears empty — no data rows found.")
    return tbl


_BALA_TABLE: dict[int, float] = _load_bala_table()


# ─────────────────────────────────────────────────────────────────────────────
# DISTANCE
# ─────────────────────────────────────────────────────────────────────────────

def haversine_km(lon1: float, lat1: float, lon2: float, lat2: float) -> float:
    """Great-circle distance in km between two lon/lat points."""
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(math.radians(lat1))
         * math.cos(math.radians(lat2))
         * math.sin(dlon / 2) ** 2)
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


# ─────────────────────────────────────────────────────────────────────────────
# NUMBER PARSING
# ─────────────────────────────────────────────────────────────────────────────

def parse_num(val) -> "float | None":
    """Convert a cell value to float; return None on failure.

    Handles: plain numbers, comma-separated strings, currency symbols,
    percentage signs, trailing asterisks/letters, and range strings like
    '600-630*' or '600.0–630.0' — the midpoint is returned for ranges.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace("$", "").replace("%", "").strip()
    m = re.match(r"^(\d+(?:\.\d+)?)\s*[-–]\s*(\d+(?:\.\d+)?)", s)
    if m:
        return (float(m.group(1)) + float(m.group(2))) / 2
    s = re.sub(r"[^0-9.]", "", s)
    try:
        return float(s) if s else None
    except Exception:
        return None


def parse_cap_rate(val) -> "float | None":
    """Cap rate / yield as a FRACTION (0.045), or None when not reported.

    The internal convention is a fraction: every cap-rate cell is written with
    Excel's "0.00%" number format, which multiplies by 100 to display it, and
    the online-search path already divides its percent-quoted cap_rate_pct by
    100 before storing. Sources quote it both ways ("4.5%", "4.5", "0.045") and
    parse_num strips the % sign WITHOUT rescaling, so "4.5%" would otherwise
    reach the report as 450.00%.

    A real cap rate is never >= 1 as a fraction (that is a 100%+ yield), so a
    value at or above 1 is percent-quoted and gets scaled down; anything below
    is already a fraction and is left alone.
    """
    v = parse_num(val)
    if v is None:
        return None
    return v / 100.0 if v >= 1 else v


# ─────────────────────────────────────────────────────────────────────────────
# TENURE / DATE PARSING
# ─────────────────────────────────────────────────────────────────────────────

def parse_remaining_yrs(val) -> "float | None":
    """Parse a remaining-leasehold value expressed as a number or tenure string.

    - Plain number         → returned as-is  (e.g. 77 → 77.0)
    - Tenure string        → total − (current_year − start_year)
                             e.g. "99 years from 2004" → 77.0 (in 2026)
    - 'Freehold' / 'FH'   → 999.0
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val) if val else None

    s = str(val).strip()

    if re.search(r"\bfreehold\b|\bfh\b", s, re.I):
        return 999.0

    m = re.search(
        r"(\d{2,3})\s*[-\s]?(?:years?|yrs?)"
        r".*?"
        r"\b(\d{4})\b",
        s, re.I,
    )
    if m:
        total_yrs = int(m.group(1))
        start_year = int(m.group(2))
        current_year = datetime.date.today().year
        if 1800 <= start_year <= current_year:
            remaining = total_yrs - (current_year - start_year)
            return float(max(remaining, 0))

    return parse_num(val)


def parse_sale_date(val, fallback_year: str = None) -> str:
    """Normalise any date-like value to 'Qn YYYY' or plain 'YYYY'.

    Handles: Python date/datetime, 'Q1 2024', 'Jan 2024', 'YYYY-MM-DD',
    'DD/MM/YYYY', 'MM/YYYY', bare year.

    fallback_year: year string extracted from the column header (e.g. "2025")
    — appended when the cell value has no 4-digit year so that bare month
    names like "Jan" or quarter labels like "Q1" get a year attached.
    """
    if val is None:
        return ""
    if isinstance(val, (datetime.date, datetime.datetime)):
        q = (val.month - 1) // 3 + 1
        return f"Q{q} {val.year}"
    s = str(val).strip()
    if not s or s in ("None", ""):
        return ""
    # Append header year when the cell value contains no 4-digit year.
    if fallback_year and not re.search(r"\b(?:19|20)\d{2}\b", s):
        s = f"{s} {fallback_year}"
    m = re.match(r"(Q[1-4])\s*(\d{4})", s, re.I)
    if m:
        return f"{m.group(1).upper()} {m.group(2)}"
    _MON = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
            "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
    m2 = re.match(r"([a-z]{3})[a-z]*[.\s\-]*(\d{4})", s, re.I)
    if m2:
        mon = _MON.get(m2.group(1).lower(), 0)
        if mon:
            return f"Q{(mon - 1) // 3 + 1} {m2.group(2)}"
    m3 = re.search(r"(\d{4})", s)
    if m3:
        year = m3.group(1)
        for m4 in re.finditer(r"(?<!\d)(\d{1,2})(?=[/\-])", s):
            try:
                mon = int(m4.group(1))
                if 1 <= mon <= 12:
                    return f"Q{(mon - 1) // 3 + 1} {year}"
            except Exception:
                pass
        return year
    return s


# ─────────────────────────────────────────────────────────────────────────────
# BALA TABLE LOOKUPS
# ─────────────────────────────────────────────────────────────────────────────

def bala_factor(n, y: float = 0.06) -> float:  # noqa: y kept for backward compat
    """Singapore Bala Table factor for n remaining years → fraction 0.0–1.0.

    n ≤ 0 or n ≥ 999 → freehold → 1.0
    n = 1 … 99       → exact table lookup
    n = 100 … 998    → linear interpolation (96 % at 99 yrs → 100 % at 999 yrs)
    """
    if n is None:
        return 1.0
    n = int(round(float(n)))
    if n <= 0 or n >= 999:
        return 1.0
    if n in _BALA_TABLE:
        return _BALA_TABLE[n] / 100.0
    if n > 99:
        frac = (n - 99) / (999 - 99)
        return (96.0 + frac * 4.0) / 100.0
    return _BALA_TABLE.get(max(1, n), 3.8) / 100.0


def bala_expr(x_ref: str) -> str:
    """Excel formula string for the Bala Table factor given cell reference x_ref.

    x_ref ≤ 0 or ≥ 999 or = "—" → 1  (freehold)
    x_ref = 1 … 99               → VLOOKUP in 'Bala Tbl'!$A$2:$B$100
    x_ref = 100 … 998            → linear interpolation (0.96 → 1.0)
    """
    return (
        f"IF(OR({x_ref}<=0,{x_ref}=\"—\",{x_ref}>=999),1,"
        f"IF({x_ref}<=99,"
        f"VLOOKUP(ROUND({x_ref},0),'Bala Tbl'!$A$2:$B$100,2,FALSE),"
        f"0.96+({x_ref}-99)*0.04/900))"
    )


# ─────────────────────────────────────────────────────────────────────────────
# ONLINE-SEARCH CROSS-SOURCE DEDUP
# ─────────────────────────────────────────────────────────────────────────────

def haversine_km(lon1, lat1, lon2, lat2) -> float:
    """Great-circle distance in km."""
    r    = 6371.0
    p1, p2 = math.radians(float(lat1)), math.radians(float(lat2))
    dp   = p2 - p1
    dl   = math.radians(float(lon2) - float(lon1))
    a    = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def find_same_building(records, lon, lat, value, value_of,
                       max_km: float = 0.075, tol: float = 0.05):
    """Index of an existing record at the same building AND the same money, else None.

    This is the cross-source dedup that merges one deal reported by two sources
    (e.g. in two languages, or under a translated building name) without fusing two
    genuinely different buildings.

    Both conditions must hold: within `max_km` (~75m) AND within `tol` on the money.

    75m is deliberately tight. Two sources quoting the same building normally quote
    the same canonical street address, so they geocode to nearly the same point; the
    tolerance only has to absorb provider jitter. Widening it merges NEIGHBOURS —
    two distinct CBD towers ~130m apart whose prices land within 5% of each other is
    entirely plausible, and fusing them loses a real comp invisibly. The two failure
    modes are not symmetric: a missed duplicate shows up as two similar rows the
    analyst can see and merge, whereas a false merge silently deletes evidence. When
    in doubt this keeps both.

    Replaces a hashed key of the form
        (round(lon, 2), round(lat, 2), round(v / max(v * 0.05, floor)))
    which was broken twice over: the value term is algebraically round(1 / 0.05) = 20
    for every v above the floor — a CONSTANT, contributing nothing — leaving a bare
    2dp coordinate cell of ~1.1km. In a CBD that is dozens of distinct towers, all
    silently merged into whichever one was found first.

    A direct comparison is used rather than a hash bucket because "within 5%" has no
    exact bucket representation: any fixed grid splits near-identical figures that
    straddle a boundary (500 vs 505). Comp lists are small, so the scan is cheap.
    """
    if lon is None or lat is None:
        return None
    try:
        v = float(value or 0)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    for i, r in enumerate(records):
        rlon, rlat = r.get("lon"), r.get("lat")
        if rlon is None or rlat is None:
            continue
        if haversine_km(rlon, rlat, lon, lat) > max_km:
            continue
        try:
            rv = float(value_of(r) or 0)
        except (TypeError, ValueError):
            continue
        if rv <= 0:
            continue
        if abs(rv - v) <= tol * max(rv, v):
            return i
    return None


# ─────────────────────────────────────────────────────────────────────────────
# UPLOADED-REPORT CROSS-SOURCE DEDUP  (used by scan_input_{sales,rent,land}_comps.py)
#
# Different from find_same_building() above: that one is a single lookup used
# while building the ONLINE-search list one record at a time. These two merge/
# flag an already-assembled list of UPLOADED-file records against each other —
# same idea (same real-world thing reported by two different sources), applied
# after all of them are geocoded.
# ─────────────────────────────────────────────────────────────────────────────

_COMP_NAME_KEYS = ("property_name", "site_name", "building_name", "property")


def _comp_name(r: dict, name_keys=_COMP_NAME_KEYS) -> str:
    return str(next((r.get(k) for k in name_keys if r.get(k)), "") or "")


def record_completeness(r: dict) -> int:
    """Count of non-blank canonical fields — used to pick the most complete
    record as the base when merging cross-source duplicates."""
    _skip = {"lon", "lat", "distance_km", "map_marker"}
    return sum(1 for k, v in r.items()
              if k not in _skip and not str(k).startswith("_") and str(v or "").strip())


def name_overlap(a: str, b: str) -> float:
    """Token Jaccard of two property/building names (0..1). Used to confirm two
    records from different sources really are the SAME building before treating
    a value mismatch as a conflict (vs. two genuinely different adjacent ones)."""
    ta = set(re.sub(r"\W+", " ", str(a or "").lower()).split())
    tb = set(re.sub(r"\W+", " ", str(b or "").lower()).split())
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


def dedup_cross_source(records: list, value_fields=("price_sgd_m",),
                       name_keys=_COMP_NAME_KEYS, name_min_overlap: float = 0.0,
                       threshold_km: float = 0.05, price_tol_frac: float = 0.15,
                       exact_km: float = 0.0) -> list:
    """
    Merge records that represent the SAME transaction/lease reported by DIFFERENT
    input sources (e.g. the same deal appears in both an uploaded Colliers PDF
    and a CBRE PDF, under different property-name spellings — 'MBFC (T3)' vs
    'Marina Bay Financial Centre Tower 3').

    Two records are treated as the same deal when:
      - they geocoded within `threshold_km` of each other, AND
      - if `name_min_overlap` > 0, their names token-overlap at least that much
        (guards against merging two genuinely different, merely-adjacent
        buildings that happen to land within `threshold_km`) — set to 0 (the
        default) to skip this check entirely, AND
      - EITHER the first of `value_fields` present as a positive number on BOTH
        records agrees within `price_tol_frac`, OR (`exact_km` > 0) they geocoded
        within `exact_km` of each other — near-identical coordinates are treated
        as sufficient evidence of the same building on their own, overriding a
        value disagreement (two sources rounding/reporting a price differently
        is more likely than two unrelated deals landing on the same rooftop
        pin). `exact_km` is disabled (0) by default; leave it off for comp types
        where the SAME building legitimately carries multiple, genuinely
        different transactions (e.g. rent: many leases, many different rents).
    Records sharing the same `_source` tag are never merged here — same-file
    duplicates are a different bug, handled earlier by the extraction pipeline's
    own name-based dedup.

    Why `name_min_overlap` defaults to off (sales/land turn it on; rent doesn't):
    for sales/land, a building is normally sold/tendered once in a reporting
    window, so requiring the names to also match is a cheap extra safety check
    against a coincidental proximity+price match between two different buildings.
    For rent, the SAME building legitimately produces many different leases at
    many different rents, so its own name will always "match" trivially — the
    `value_fields` (rent) agreement is what actually establishes "same lease
    reported twice", and requiring name overlap on top of that adds nothing.

    For each matched pair/group, keeps the MOST COMPLETE record (most non-blank
    fields) as the base and fills any of its still-blank fields from the other
    record(s) — so complementary data from both sources survives (e.g. one
    source has psf, the other has the cap rate). Note the `exact_km` path can
    therefore merge away a genuine price/rent DISAGREEMENT between two records at
    the same coordinates — for a comp type where the same building can carry more
    than one distinct sale, that would silently discard one instead of flagging
    it, which is exactly why it must stay off for those comp types.

    value_fields: ordered field names to compare — the first one present as a
    positive number on both records decides agreement (e.g. rent comps may only
    have eff_rent OR asking_rent filled in, not both).
    """
    n = len(records)
    used = [False] * n
    groups: list = []
    for i in range(n):
        if used[i] or records[i].get("lon") is None or records[i].get("lat") is None:
            continue
        group = [i]
        for j in range(i + 1, n):
            if used[j] or records[j].get("lon") is None or records[j].get("lat") is None:
                continue
            if records[i].get("_source") and records[i].get("_source") == records[j].get("_source"):
                continue  # same source — not a cross-source duplicate
            dist = haversine_km(records[i]["lon"], records[i]["lat"],
                               records[j]["lon"], records[j]["lat"])
            if dist > threshold_km:
                continue
            if name_min_overlap and name_overlap(
                    _comp_name(records[i], name_keys),
                    _comp_name(records[j], name_keys)) < name_min_overlap:
                continue  # nearby but a different building — not a duplicate
            if exact_km and dist <= exact_km:
                group.append(j)
                continue  # near-identical coordinates — duplicate regardless of value agreement
            agree = False
            for vf in value_fields:
                p1, p2 = parse_num(records[i].get(vf)), parse_num(records[j].get(vf))
                if p1 and p2 and p1 > 0 and p2 > 0:
                    agree = abs(p1 - p2) / max(p1, p2) <= price_tol_frac
                    break
            if not agree:
                continue
            group.append(j)
        if len(group) > 1:
            groups.append(group)
            for idx in group:
                used[idx] = True

    if not groups:
        return records

    grouped_idx = {idx for g in groups for idx in g}
    result = [r for i, r in enumerate(records) if i not in grouped_idx]

    for group in groups:
        group_recs = sorted((records[i] for i in group),
                            key=record_completeness, reverse=True)
        base = dict(group_recs[0])
        sources = [s for s in (r.get("_source") for r in group_recs) if s]
        filled = []
        for other in group_recs[1:]:
            for k, v in other.items():
                if k in ("lon", "lat", "distance_km", "map_marker") or str(k).startswith("_"):
                    continue
                if not str(base.get(k) or "").strip() and str(v or "").strip():
                    base[k] = v
                    filled.append(k)
        if sources:
            base["_source"] = "+".join(dict.fromkeys(sources))  # de-dup, keep order
        if filled:
            base["_dedup_note"] = (f"merged {len(group_recs)} cross-source record(s); "
                                   f"filled from other source: {', '.join(sorted(set(filled)))}")
        print(f"      [cross-source dedup] merged {len(group_recs)} record(s) → "
              f"{_comp_name(base, name_keys)[:40]!r} (sources: {base.get('_source','')})")
        result.append(base)
    return result


def flag_cross_source_conflicts(records: list, value_fields=(("price_sgd_m", "price"),),
                                name_keys=_COMP_NAME_KEYS,
                                threshold_km: float = 0.2, price_tol_frac: float = 0.15,
                                name_min_overlap: float = 0.5) -> list:
    """Flag (do NOT merge, do NOT drop) records that are the SAME building/unit
    reported by DIFFERENT sources but whose key numeric value DISAGREES.

    Complements dedup_cross_source: that function MERGES cross-source records
    whose value agrees; this one catches the opposite case — same building,
    sources report DIFFERENT numbers (e.g. Colliers vs Cushman MarketBeat both
    list 'Northpoint City South Wing' but at different prices). Silently keeping
    both rows, or silently picking one, hides a real data question, so instead
    we attach a ``_review_flag`` to both records and print a run-log warning
    (same surfacing channel as the geocode-centroid warning) so the analyst can
    verify which source is right.

    'Same building' requires BOTH geocode proximity (≤ threshold_km) AND a name
    token overlap ≥ name_min_overlap, so two genuinely different buildings that
    merely sit within threshold_km of each other in a dense CBD are not flagged.

    value_fields: ordered (field_key, label) pairs — every pair present as a
    positive number on both records is checked (unlike dedup_cross_source,
    which only needs the first mutually-present field to agree).
    """
    n = len(records)
    n_flagged = 0
    for i in range(n):
        ri = records[i]
        if ri.get("lon") is None or ri.get("lat") is None:
            continue
        for j in range(i + 1, n):
            rj = records[j]
            if rj.get("lon") is None or rj.get("lat") is None:
                continue
            if ri.get("_source") and ri.get("_source") == rj.get("_source"):
                continue  # same source — not cross-source
            if haversine_km(ri["lon"], ri["lat"], rj["lon"], rj["lat"]) > threshold_km:
                continue
            if name_overlap(_comp_name(ri, name_keys),
                            _comp_name(rj, name_keys)) < name_min_overlap:
                continue  # nearby but a different building — not a conflict

            conflicts = []
            for fk, label in value_fields:
                a, b = parse_num(ri.get(fk)), parse_num(rj.get(fk))
                if a and b and a > 0 and b > 0 \
                        and abs(a - b) / max(a, b) > price_tol_frac:
                    conflicts.append(
                        f"{label}: {ri.get('_source','?')}={a:g} vs "
                        f"{rj.get('_source','?')}={b:g}")
            if not conflicts:
                continue

            note = ("Cross-source disagreement for the same building — "
                    + "; ".join(conflicts) + ". Verify which source is correct.")
            for r in (ri, rj):
                r["_review_flag"] = note
            n_flagged += 1
            print(f"      [cross-source CONFLICT] "
                  f"{_comp_name(ri, name_keys)[:40]!r}: {'; '.join(conflicts)} "
                  f"— flagged for review")
    if n_flagged:
        print(f"      Cross-source conflict check: {n_flagged} disagreement(s) "
              f"flagged for review (kept both rows; not merged)")
    return records
