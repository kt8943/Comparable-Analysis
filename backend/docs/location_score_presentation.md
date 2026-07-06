# Location Competitiveness Score — Presentation Deck (Singapore)

A slide-by-slide script. Each slide has **on-slide bullets** and **speaker notes** (the
detail you say out loud). Reflects the current production logic.

---

## Slide 1 — Title
**Location Competitiveness Score**
*How each comparable is labelled Superior / Comparable / Inferior by location*

**Speaker notes.** Every comp table needs a consistent, objective read on *location quality*
relative to the subject. This model produces that in one word — Superior, Comparable, or
Inferior — using Singapore's official land-use map, so it's repeatable and defensible rather
than a manual judgement call.

---

## Slide 2 — The problem we're solving
- Analysts previously judged each comp's location **by hand** — slow, subjective, inconsistent.
- We need a **repeatable, data-grounded** signal that adjusts for **asset class**
  (what's "prime" for an office differs from retail or logistics).
- Output must be **simple** (a 3-way label) and **auditable**.

**Speaker notes.** The score isn't trying to be a valuation — it's a *directional* location
read that's consistent across deals and analysts, and that respects the fact that "good
location" means different things for an office vs a warehouse vs a mall.

---

## Slide 3 — Data foundation
- **URA Master Plan** land-use parcels — every zoned plot in Singapore.
- Pre-processed into a lightweight cache: each parcel's **centre point + area (km²)**,
  bucketed into 5 uses — **residential, commercial, business, hotel, port/airport**.
- Hotels also use **OneMap "Tourist Attractions."**
- **No live network** at runtime; **181 MB map → ~3 MB cache** (deploys to the cloud).

**Speaker notes.** We shrink the full 181 MB Master Plan to a ~3 MB cache by keeping each
parcel's centre and area and dropping the polygon shape. That's enough for "how much of the
surroundings is X" and "how far to the nearest Y," and it's small enough to run anywhere,
including the cloud. The trade-off is a small approximation we'll flag later.

---

## Slide 4 — The method in three steps
1. **Measure 2 location factors** for the subject and for each comp — factors chosen for the
   asset class.
2. **Compare** each comp to the subject, factor by factor → a sub-score in **−1…+1**.
3. **Average** the two sub-scores → the location score, then map to a label.

**Speaker notes.** Subject is always the benchmark (score 0). A comp scores positive when its
surroundings are better than the subject's, negative when worse. Two factors per sector keeps
it interpretable — one "density of the right land use nearby" and one "proximity to the right
kind of centre/node."

---

## Slide 5 — Two kinds of factor
**A. Land-use coverage (density).**
```
coverage = (area of parcels of a land use whose CENTRE is within 1 km) ÷ (area of the 1 km circle)
```
- Share of the 1 km circle covered by that use. Bigger estates weigh more than tiny lots.

**B. Proximity to a node (distance or attractiveness).**
- Distance (km) to the nearest relevant node (CBD, freight hub), **or**
- A tier-weighted **attractiveness** score (retail centres).

**Speaker notes.** Coverage answers "how much of the neighbourhood is the *right* kind of land
for this asset." Proximity answers "how close is it to the *right* kind of centre." Every
sector combines one of each.

---

## Slide 6 — Factors by sector  *(the core slide)*

| Sector | Factor 1 | Factor 2 |
|---|---|---|
| **Office** | Distance to **CBD** (Raffles Place) ↓ | **Commercial** coverage (1 km) ↑ |
| **Retail** | **Commercial/retail** coverage (1 km) ↑ | **Retail-centre attractiveness** (tiered) ↑ |
| **Industrial / Logistics / Data centre** | **Business** coverage (1 km) ↑ | Distance to nearest **freight hub** ↓ |
| **Hotel** | **Tourist attractions** within 1 km ↑ | **Commercial** coverage (1 km) ↑ |
| **Mixed** | Distance to **CBD** ↓ | **Residential + Commercial** coverage ↑ |

**Speaker notes.** Read across each row: office value = CBD proximity + surrounding commercial
density; retail = retail-cluster density + being near a *prime* shopping centre; industrial =
industrial-land cluster + freight access; hotel = tourist draw + a lively commercial area;
mixed = CBD + combined residential/commercial vibrancy.

---

## Slide 7 — Comparing a comp to the subject (the math)
Each factor → a sub-score in **[−1, +1]** (0 when comp = subject).

**Coverage / count (higher = better):**
```
sub_score = (comp − subject) ÷ (comp + subject + k)
```
- `k` dampens small-number noise: **0.3** for coverage, **10** for the tourist count.

**Distance (lower = better):**
```
sub_score = clamp( (subject_distance − comp_distance) ÷ 5 km , −1, +1 )
```
- Comp closer than subject → positive; a **5 km** gap = the full ±1.

**Speaker notes.** The smoothing `k` stops tiny differences from producing big swings. The 5 km
reference on distances stops a subject that sits *on* a landmark (distance ≈ 0) from forcing
every comp to −1 — it scales the *difference*, not the raw distance.

---

## Slide 8 — Final score & label
```
location_score = average(factor1_subscore, factor2_subscore)     # −1 … +1
```
| Score | Label |
|---|---|
| **> +0.3** | **Superior** |
| **−0.3 … +0.3** | **Comparable** |
| **< −0.3** | **Inferior** |

**Speaker notes.** Three buckets on purpose — this is a directional read, not a decimal
valuation input. The ±0.3 dead-band means a comp has to be *clearly* better or worse to move
off "Comparable."

---

## Slide 9 — Worked example: OFFICE
Subject = Raffles Place office. Comp = CapitaSpring (≈300 m away).

| | Subject | Comp |
|---|---|---|
| Distance to CBD | 0.0 km | 0.3 km |
| Commercial coverage (1 km) | 0.20 | 0.19 |

- Factor 1 (CBD distance ↓): `(0.0 − 0.3)/5 = −0.06`
- Factor 2 (commercial coverage ↑): `(0.19 − 0.20)/(0.19+0.20+0.3) = −0.01`
- **Score = −0.04 → Comparable.** A far suburban office ≈ **−0.7 → Inferior.**

---

## Slide 10 — Deep dive: RETAIL centre tiering
- Retail-centre proximity is a **0–1 attractiveness** score, not a raw distance:
```
proximity      = max(0, 1 − distance / 3 km)          # distance still drives it
attractiveness = tier_weight × proximity              # the tier is a SCORE penalty
```
- **Prime centres = 1.0:** Orchard, CBD (Raffles Place).
- **URA Regional Centres = 0.6:** Jurong Lake District, Tampines, Woodlands, Seletar.
- At a regional centre you top out at **0.6**; at a prime centre you reach **1.0** — a **0.4
  score penalty**. Beyond **3 km** from every centre → **0** ("not near a retail centre").

**Worked comparison (both 2 km from their nearest centre):**
| | proximity | × tier | attractiveness |
|---|---|---|---|
| Subject — 2 km from **Orchard** (prime) | 0.33 | ×1.0 | 0.33 |
| Comp — 2 km from **Jurong** (regional) | 0.33 | ×0.6 | 0.20 |

Retail-centre factor = `(0.20 − 0.33)/(0.20 + 0.33 + 0.3) = −0.16` → comp is **inferior** on
this factor because it's near a *regional*, not *prime*, centre.

**Speaker notes.** We deliberately put the penalty on the *score*, not the distance — so
distance still matters (closer = better) but a suburban centre can never look as good as
Orchard/CBD. The "why 0.6": 0.7 was too soft (regional ≈ prime), 0.5 let the tier dominate;
0.6 makes prime clearly better while staying balanced against the coverage factor.

---

## Slide 11 — Deep dive: INDUSTRIAL freight access
- Factor 2 = distance to the **nearest freight hub**, **equal treatment** (near any hub is good).
- **5 hubs**, data-derived from the URA port/airport parcels (offshore strays removed):
  **Tuas, Jurong, PSA/Keppel, Changi, Seletar.**
- Measured to the **nearest parcel** of a hub (so a sprawling port is measured to its nearest
  edge, not a far centre).

**Speaker notes.** We chose equal treatment — a coarse label shouldn't try to rank Tuas vs
Changi vs Seletar. Being near *a* major sea/air freight node is the signal. Factor 1 (business-
land coverage) then differentiates comps that are all "near a hub."

---

## Slide 12 — Guardrails
- **Same sector only.** A comp is scored only against a same-class subject (Mixed matches any);
  otherwise Location is blank.
- **Singapore-only.** Both subject and comp must be inside Singapore; the model is built on SG
  geography, so overseas comps are blank.
- **Analyst override.** An input file's own Superior/Comparable/Inferior label is **kept** —
  never overwritten by the computed one.
- **Coordinates.** Uses the same lat/long plotted on the map (Google/Mapbox), else a OneMap
  geocode.

---

## Slide 13 — Limitations (be upfront)
- **Approximation, not valuation.** Coverage counts a whole parcel if its *centre* is in the
  circle; nodes (CBD, retail centres, hubs) are fixed reference points. Output is a coarse,
  **directional** 3-bucket signal.
- **Hand-set nodes.** CBD and retail-centre coordinates + tier weights are modelling choices
  (defensible, but subjective). Freight hubs *are* data-derived.
- **Cloud vs local.** SG location computes on the cloud (via the cache); overseas markets need
  their own equivalents.

---

## Slide 14 — Roadmap / possible extensions
- **Other markets** (e.g. Seoul/Tokyo) with their own CBD/centre/hub definitions.
- **True area intersection** (polygon-clipped coverage) if we run locally with the full map.
- **Tunable knobs** already exposed: retail tier weight (0.6), influence radius (3 km),
  smoothing (0.3), distance reference (5 km), label thresholds (±0.3).

---

## Slide 15 — One-line summary
> For each comp we measure **two location factors** tailored to the asset class, compare them
> to the subject, average into a **−1…+1 score**, and label it **Superior / Comparable /
> Inferior** — grounded in URA's Master Plan, adjusted per sector, and fully repeatable.
