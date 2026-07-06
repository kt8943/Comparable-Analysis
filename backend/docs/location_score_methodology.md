# Location Competitiveness Score — Methodology (Singapore)

## 1. What it is and why
Every comparable-transactions table needs a consistent read on **how good each comp's
location is relative to the subject property**. Historically that was an analyst's manual
judgement — slow, subjective, and inconsistent between people and deals. This model replaces
that with a **repeatable, data-grounded signal**: it labels each comp **Superior**,
**Comparable**, or **Inferior** on location, adjusted for the **asset class** (what counts as
"prime" for an office is different from retail, logistics, or a hotel). It is deliberately a
**directional 3-way label, not a valuation input** — a quick, defensible, auditable read.

## 2. The data behind it
The model is built on the **URA Master Plan** — the official land-use map covering every
zoned plot in Singapore. Because the full map is large (181 MB), it is pre-processed once into
a small (~3 MB) cache that keeps, for each parcel, only its **centre point and its area
(km²)**, sorted into five land-use groups: **residential, commercial, business, hotel, and
port/airport**. Hotels additionally use OneMap's **"Tourist Attractions"** layer. This cache
is enough to answer the only two questions the model asks — *"how much of the surroundings is
a given land use?"* and *"how far is the nearest relevant centre or hub?"* — and it is small
enough to run anywhere, including the cloud, with no live network calls.

## 3. How the score is built (three steps)
For the **subject** and for **each comp**, the model measures **two location factors** chosen
for that asset class. It then **compares each comp to the subject**, factor by factor, turning
each into a sub-score between **−1 and +1** (0 means the comp matches the subject). Finally it
**averages the two sub-scores** into the location score and maps it to a label. The subject is
always the benchmark (score 0); a comp scores **positive** when its surroundings are *better*
than the subject's and **negative** when they are *worse*.

## 4. The two kinds of factor
Every sector's score combines one factor of each type below.

**Land-use coverage (a "density" factor).** This measures what share of the **1 km circle**
around a property is covered by a given land use:
```
coverage = (total area of parcels of that use whose CENTRE is within 1 km) ÷ (area of the 1 km circle)
```
Because it is area-weighted, large estates count for more than tiny lots. It is an
approximation — a whole parcel is counted if its *centre* falls inside the circle — but it is
used **relatively** (comp vs subject), so it never needs to be a capped percentage.

**Proximity to a node (a "distance" or "attractiveness" factor).** This measures how close a
property is to the right kind of centre: either the straight-line distance (km) to the nearest
node (the CBD, or a freight hub), or — for retail — a tier-weighted **attractiveness** score
that blends distance with how *prime* the nearest centre is.

## 5. Which factors each sector uses
The two factors are tailored to what drives location value for each asset class.

| Sector | Factor 1 | Factor 2 |
|---|---|---|
| **Office** | Distance to the **CBD** (Raffles Place) — closer is better | **Commercial** land coverage within 1 km — more is better |
| **Retail** | **Commercial / retail** land coverage within 1 km — more is better | **Retail-centre attractiveness** (tiered; prime vs regional) — higher is better |
| **Industrial / Logistics / Data centre** | **Business** land coverage within 1 km — more is better | Distance to the nearest **freight hub** — closer is better |
| **Hotel** | **Tourist attractions** within 1 km (count) — more is better | **Commercial** land coverage within 1 km — more is better |
| **Mixed** | Distance to the **CBD** — closer is better | **Residential + Commercial** coverage within 1 km — more is better |

In words: office value tracks CBD proximity and the surrounding commercial density; retail
tracks its retail-cluster density and being near a *prime* shopping centre; industrial tracks
the surrounding industrial-land cluster and access to a major sea/air freight node; a hotel
tracks its tourist draw and how lively the surrounding commercial area is; and a mixed-use
asset blends CBD proximity with combined residential-plus-commercial vibrancy.

## 6. How a comp is compared to the subject
Each factor produces a sub-score in **[−1, +1]**, equal to 0 when the comp and subject are
equal. There are two formulas, one per factor type.

**Coverage or count factors (higher is better)** use a smoothed relative difference:
```
sub_score = (comp_value − subject_value) ÷ (comp_value + subject_value + k)
```
The constant `k` dampens noise when the numbers are small — it is **0.3** for coverage
fractions and **10** for the tourist-attraction count.

**Distance factors (lower is better)** scale the difference by a fixed 5 km reference and clamp
to ±1:
```
sub_score = clamp( (subject_distance − comp_distance) ÷ 5 km , −1, +1 )
```
A comp that is **closer** than the subject scores **positive**, and a 5 km gap corresponds to
the full ±1. The 5 km reference matters: it scales the *difference* rather than the raw
distance, so a subject that sits right on a landmark (distance ≈ 0) does not force every comp
to −1.

## 7. Final score and label
The location score is the **average of the two factor sub-scores**, ranging from −1 to +1,
mapped to three labels:

| Location score | Label |
|---|---|
| **greater than +0.3** | **Superior** |
| **−0.3 to +0.3** | **Comparable** |
| **less than −0.3** | **Inferior** |

The ±0.3 dead-band is intentional: a comp must be *clearly* better or worse to move off
"Comparable," which keeps the label stable and honest about the model's coarseness.

## 8. Worked example — office
Take a subject office at **Raffles Place** and a comp at **CapitaSpring**, about 300 m away.
The subject is at the CBD (distance 0.0 km) with commercial coverage 0.20; the comp is 0.3 km
from the CBD with commercial coverage 0.19. Factor 1 (CBD distance) gives
`(0.0 − 0.3) / 5 = −0.06`; factor 2 (commercial coverage) gives
`(0.19 − 0.20) / (0.19 + 0.20 + 0.3) = −0.01`. The average is **−0.04 → Comparable**. By
contrast, a far suburban office — a large CBD distance and low commercial coverage — would
score around **−0.7 → Inferior**.

## 9. Deep dive — retail centre tiering
Retail is the one sector where *which* centre you are near matters, not just how close. So its
second factor is a **0–1 attractiveness score** rather than a raw distance:
```
proximity      = max(0, 1 − distance ÷ 3 km)        # distance still drives this
attractiveness = tier_weight × proximity            # the tier is a SCORE penalty
```
Distance still counts — being closer to a centre raises the score, and beyond **3 km** from
every centre the score is **0** ("not near a retail centre"). What the tier adds is a **score
penalty** for less-prime centres:

| Tier | Centres | Weight | Meaning |
|---|---|---|---|
| **Prime** | Orchard, CBD (Raffles Place) | **1.0** | Reaches 1.0 at the centre |
| **Regional** | Jurong Lake District, Tampines, Woodlands, Seletar | **0.6** | Tops out at 0.6 — a 0.4 penalty |

So a property *at* a regional centre can only reach 0.6, while one *at* a prime centre reaches
1.0. Example: two properties each 2 km from their nearest centre have the same proximity
(`1 − 2/3 = 0.33`), but the subject near **Orchard** gets `1.0 × 0.33 = 0.33` and the comp
near **Jurong** gets `0.6 × 0.33 = 0.20`; the factor is
`(0.20 − 0.33) / (0.20 + 0.33 + 0.3) = −0.16`, correctly marking the comp inferior *because it
is near a regional rather than a prime centre.* The 0.6 weight was chosen as a balance — 0.7
made a regional centre look almost as good as prime, while 0.5 let the tier dominate the
label; 0.6 makes prime clearly better while staying balanced against the coverage factor.

## 10. Deep dive — industrial freight access
For industrial, logistics, and data-centre subjects, the second factor is the distance to the
nearest **freight hub**, with **equal treatment** — being near *any* major sea or air hub is
good, and the model does not try to rank them. The hubs were **derived from the URA
port/airport parcels** (with offshore stray parcels removed), giving five: **Tuas, Jurong,
PSA/Keppel, Changi, and Seletar**. Distance is measured to the **nearest parcel** of a hub, so
a sprawling port like Tuas is measured to its nearest edge rather than a distant centre point.
The first factor (business-land coverage) then differentiates comps that are all "near a hub."

## 11. Guardrails
- **Same sector only.** A comp is scored only against a same-class subject (a Mixed subject
  matches anything); otherwise its Location is left blank.
- **Singapore-only.** Both the subject and the comp must sit inside Singapore, since the model
  is built on Singapore geography; overseas comps are left blank.
- **Analyst override.** If the input file already provides a Superior/Comparable/Inferior
  label, that label is **kept** and never overwritten by the computed one.
- **Coordinates.** The score uses the same latitude/longitude plotted on the map
  (Google/Mapbox), falling back to a OneMap geocode only when coordinates are missing.

## 12. Limitations (state these plainly)
- **Approximation, not valuation.** Coverage counts a whole parcel if its centre is in the
  circle, and the CBD, retail centres, and freight hubs are fixed reference points. The output
  is a **coarse, directional** three-bucket signal, not a precise valuation input.
- **Some nodes are hand-set.** The CBD point, the retail-centre coordinates, and the tier
  weights are modelling choices — defensible but subjective; only the freight hubs are
  data-derived.
- **Tunable.** The key knobs are all adjustable: retail tier weight (0.6), retail influence
  radius (3 km), coverage smoothing (0.3), the 5 km distance reference, and the ±0.3 label
  thresholds.

## 13. One-line summary
For each comp the model measures **two asset-class-tailored location factors**, compares them
to the subject, **averages** them into a **−1…+1 score**, and labels it **Superior /
Comparable / Inferior** — grounded in URA's Master Plan, adjusted per sector, and fully
repeatable.

---

## Appendix A — Exact reference points, definitions & sources
This is what every term in the model concretely resolves to.

**CBD (used by Office & Mixed).** A single **hand-set** point at **Raffles Place MRT —
1.28348° N, 103.85176° E**. "Distance to CBD" is the straight-line distance to this one point.
A second node (e.g. Jurong Lake District) is deliberately *not* used, so "CBD" means the
central core only.

**Retail centres (used by Retail).** Six fixed points in two tiers:

| Centre | Lat, Lon | Tier | Weight |
|---|---|---|---|
| CBD (Raffles Place) | 1.28348, 103.85176 | Prime | 1.0 |
| Orchard | 1.3040, 103.8330 | Prime | 1.0 |
| Jurong Lake District | 1.3330, 103.7420 | Regional | 0.6 |
| Tampines | 1.3540, 103.9450 | Regional | 0.6 |
| Woodlands | 1.4370, 103.7860 | Regional | 0.6 |
| Seletar | 1.4050, 103.8850 | Regional | 0.6 |

The four regional centres are **URA's designated Regional Centres**; Orchard and the CBD are
the prime retail belts layered on top. Coordinates are **hand-set** approximations (Seletar is
a *future* centre, so its point is approximate).

**Freight hubs (used by Industrial / Logistics / Data centre).** Five hubs — **Tuas, Jurong,
PSA/Keppel, Changi, Seletar** — which are **not** hand-typed: they are the **47 URA
"PORT / AIRPORT"-zoned parcels**, after removing 2 offshore stray parcels outside mainland
Singapore. "Distance to freight hub" is the distance to the **nearest such parcel** (so a
sprawling port is measured to its nearest edge).

**Tourist attractions (used by Hotel).** From OneMap's official **"Tourist Attractions" theme**
(`retrieveTheme`, `queryName=tourism`), cached locally. "Attractions within 1 km" counts these
points.

**Land-use buckets (used by every coverage factor).** Each URA parcel's `LU_DESC` maps to:

| Bucket | URA `LU_DESC` values |
|---|---|
| **Residential** | RESIDENTIAL · RESIDENTIAL / INSTITUTION · RESIDENTIAL WITH COMMERCIAL AT 1ST STOREY · COMMERCIAL & RESIDENTIAL |
| **Commercial** | COMMERCIAL · COMMERCIAL / INSTITUTION · COMMERCIAL & RESIDENTIAL · RESIDENTIAL WITH COMMERCIAL AT 1ST STOREY |
| **Business** | BUSINESS 1 · BUSINESS 2 · BUSINESS PARK (incl. their "- WHITE" variants) |
| **Hotel** | HOTEL |
| **Port / Airport** | PORT / AIRPORT |

A parcel with a dual use (e.g. "COMMERCIAL & RESIDENTIAL") is counted in **both** relevant
buckets.

**Fixed constants.** Coverage radius **1 km**; retail-centre influence radius **3 km**;
distance reference **5 km**; smoothing **k = 0.3** (coverage) / **10** (attraction count);
label thresholds **±0.3**. All are tunable in one place.

## Appendix B — Anticipated Q&A
**Q. Why only Raffles Place as the CBD — not Marina Bay or a second CBD like JLD?**
The CBD is defined as the single heart at Raffles Place MRT (deal-team preference). Marina Bay
is ~0.5 km away, so it still reads as CBD-adjacent; a second node was dropped so "CBD" means
the central core.

**Q. Where did the retail-centre list come from?**
URA's four official Regional Centres (JLD, Tampines, Woodlands, Seletar) plus the two prime
retail belts (Orchard, CBD). Coordinates are hand-set; Seletar is approximate (a future centre).

**Q. Why treat all ports/airports equally — isn't Tuas bigger than Seletar?**
For a coarse 3-bucket label we reward proximity to *any* major freight node; ranking them adds
subjectivity for little benefit. The business-land coverage factor then separates comps that
are all "near a hub."

**Q. Why 1 km / 3 km / 5 km?**
1 km = the immediate catchment for land-use density; 3 km = how far a retail centre's pull
reaches before fading to 0; 5 km = the scale over which a distance *difference* between comp
and subject maps to the full ±1. All tunable.

**Q. Why 0.6 for regional retail centres?**
It's a 0.4 score penalty that puts a regional centre clearly below prime. 0.7 was too soft
(≈ prime), 0.5 too dominant; 0.6 balances against the coverage factor.

**Q. How accurate is the "parcel centre in circle" approximation?**
Tested: moving a comp ±100 m barely changes the score and never flips the label (the data is
granular — ~84k residential, ~10k commercial parcels). It's a coarse label, so the
approximation is immaterial; it is not a precise valuation input.

**Q. What about a comp outside Singapore, or far from any centre?**
Non-SG comps are left blank (SG-only model). A retail comp more than 3 km from every centre
gets 0 on the centre factor ("not near a retail centre").

**Q. Why area coverage instead of a simple parcel count?**
Area-weighting lets a large estate count more than a tiny lot — a better density proxy than
treating every parcel equally.

**Q. Can an analyst override the label?**
Yes — if the input file already has a Superior/Comparable/Inferior label, it is kept and never
overwritten.

**Q. Does it work on the cloud?**
Yes for Singapore — the ~3 MB cache is deployed; the full 181 MB map isn't needed at runtime.
