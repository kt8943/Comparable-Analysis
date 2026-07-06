# Location Competitiveness Score — Methodology (Singapore)

**Purpose.** Give each comparable a location label — **Superior / Comparable / Inferior** —
relative to the subject property, based on how well its immediate surroundings suit its
asset class. Singapore-only; foreign comps are left blank.

**Data source.** URA Master Plan land-use parcels (every zoned plot in Singapore),
pre-processed into a lightweight cache of each parcel's **centre point + area (km²)**,
bucketed into five land uses: **residential, commercial, business, hotel, port/airport**.
Hotel/tourist draw also uses OneMap "Tourist Attractions". No live network calls at runtime.

---

## 1. The idea in one line
> For the subject and each comp we measure **2 location factors** chosen for that asset
> class, compare each comp to the subject, average the two into a score from **−1 to +1**,
> and label it.

Subject is always the benchmark (score = 0). A comp scores **positive** when its
surroundings are *better* than the subject's, **negative** when *worse*.

---

## 2. The two kinds of factor
Every sector's score is built from two factors, each of one of two types:

**A. Land-use coverage (a "density" factor).**
What share of the **1 km circle** around the property is covered by a given land use.
```
coverage = (total area of parcels of that land use whose CENTRE is within 1 km) ÷ (area of the 1 km circle)
```
- Approximation: a whole parcel is counted if its **centre** falls inside the circle.
- Bigger estates therefore weigh more than tiny lots (an improvement over a raw count).
- Because it is used **relatively** (comp vs subject), it doesn't need to be a capped %.

**B. Distance to a landmark (a "proximity" factor).**
Straight-line distance (km) to the **nearest** relevant node:
- **CBD** = Raffles Place MRT `(1.28348° N, 103.85176° E)`.
- **Retail-centre attractiveness** (retail): a 0–1 score = `tier_weight × proximity`, where
  `proximity = max(0, 1 − distance/3 km)` (distance still drives it) and `tier_weight` is
  **1.0 for prime** centres (Orchard, CBD) and **0.7 for URA Regional Centres** (Jurong Lake
  District, Tampines, Woodlands, Seletar). So being *at* a regional centre tops out at 0.7
  while a prime centre reaches 1.0 — a **0.3 score penalty** for regional (higher = better).
- **Freight nodes** (industrial): nearest of the 5 real port/airport hubs —
  **Tuas, Jurong, PSA/Keppel, Changi, Seletar** (offshore stray parcels removed).

---

## 3. Factors by sector
Each row is one asset class. "↑ better" = more is better; "↓ better" = closer/less is better.

| Subject sector | Factor 1 | Factor 2 |
|---|---|---|
| **Office** | Distance to **CBD** ↓ better | **Commercial** land coverage (1 km) ↑ better |
| **Retail** | **Commercial/retail** land coverage (1 km) ↑ better *(mall cluster)* | Distance to nearest **retail centre — tiered** (Orchard/CBD prime; regional centres +2 km) ↓ better |
| **Industrial / Logistics / Data centre** | **Business** land coverage (1 km) ↑ better *(industrial cluster)* | Distance to nearest **port/airport hub** ↓ better |
| **Hotel / Hospitality** | **Tourist attractions** within 1 km ↑ better *(count)* | **Commercial** land coverage (1 km) ↑ better |
| **Mixed** | Distance to **CBD** ↓ better | **Residential + Commercial** coverage (1 km) ↑ better |

*Rationale.* Office value tracks CBD proximity + surrounding commercial density; retail
tracks its residential catchment + being near a mall precinct; industrial tracks the
industrial cluster + freight access; hotels track tourist draw + a lively commercial area.

---

## 4. Comparing a comp to the subject (per factor)
Each factor produces a sub-score in **[−1, +1]** (0 when comp equals subject):

**Coverage / count factors (↑ better):** smoothed relative difference
```
sub_score = (comp_value − subject_value) ÷ (comp_value + subject_value + k)
```
- `k` = smoothing constant that dampens noise when values are small.
  `k = 0.3` for coverage fractions, `k = 10` for the tourist-attraction count.

**Distance factors (↓ better):** difference scaled by a fixed 5 km reference, clamped to ±1
```
sub_score = clamp( (subject_distance − comp_distance) ÷ 5 km , −1, +1 )
```
- Comp **closer** than the subject → **positive**. A 5 km gap = the full ±1.
- The 5 km reference stops a subject that sits *on* a landmark (distance ≈ 0) from forcing
  every comp to −1.

---

## 5. Final score & label
```
location_score = average(sub_score_factor1, sub_score_factor2)      # range −1 … +1
```
| Score | Label |
|---|---|
| **> +0.3** | **Superior** |
| **−0.3 … +0.3** | **Comparable** |
| **< −0.3** | **Inferior** |

---

## 6. Worked example (office comp vs a Raffles Place subject)
| | Subject (Raffles Place) | Comp (CapitaSpring) |
|---|---|---|
| Distance to CBD | 0.0 km | 0.3 km |
| Commercial coverage (1 km) | 0.20 | 0.19 |

- Factor 1 (CBD distance, ↓): `(0.0 − 0.3) / 5 = −0.06`
- Factor 2 (commercial coverage, ↑): `(0.19 − 0.20) / (0.19 + 0.20 + 0.3) = −0.014`
- **Score** = average(−0.06, −0.014) = **−0.04 → Comparable**

A far suburban office (large CBD distance, low commercial coverage) would score near −0.7 → **Inferior**.

---

## 7. Guardrails & caveats (state these on the slide)
- **Same sector only.** A comp is scored only if it's the **same asset class** as the
  subject (Mixed matches anything); otherwise Location is blank.
- **Singapore-only.** Both subject and comp must sit inside Singapore; the model is built
  on Singapore geography, so non-SG comps are left blank.
- **Analyst override.** If the input file already provides a Superior/Comparable/Inferior
  label, that is **kept** — never overwritten by the computed one.
- **Approximation, not valuation.** Coverage counts a whole parcel when its centre is in
  the circle; nodes (CBD, regional centres, freight hubs) are fixed reference points. The
  output is a **coarse, directional** 3-bucket signal — not a precise valuation input.
- **Coordinates.** Uses the same map-resolved lat/long plotted on the map (Google/Mapbox),
  falling back to a OneMap geocode when coordinates are missing.
