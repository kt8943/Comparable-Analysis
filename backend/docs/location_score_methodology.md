# Location Competitiveness Score — Methodology & Justification

How the **Location** column (Superior / Comparable / Inferior) is calculated for
each comparable, and why the formula is built the way it is.

Source code: `backend/tools/location_score.py`

---

## 1. The idea in one sentence

For each comp we ask: **"Is this comp's location better or worse than the subject
property's location?"** — measured on the things that actually drive value for that
asset type, and expressed as a single number where the subject is always **zero**.

- Score **> 0**  → comp location is **better** than the subject
- Score **= 0**  → comp location is **the same** as the subject
- Score **< 0**  → comp location is **worse** than the subject

The score runs from **−1 to +1**, then gets a plain-English label.

---

## 2. The four steps

### Step 1 — Pick the location "factors" for the asset type

Each sector is judged on **two** location drivers. Some are *"more is better"*
(counts of nearby parcels), some are *"closer is better"* (distances).

| Asset type            | Factor 1                          | Factor 2                              |
|-----------------------|-----------------------------------|---------------------------------------|
| **Office**            | Closer to the CBD                 | More commercial parcels within 1 km   |
| **Industrial / Data centre** | More business parcels within 1 km | Closer to a port / airport      |
| **Retail**            | More residential parcels within 1 km | Closer to a regional centre / mall hub |
| **Hospitality**       | More tourist attractions within 1 km | More commercial parcels within 1 km |
| **Mixed use**         | Closer to the CBD                 | More residential + commercial within 1 km |

The exact same two factors are measured for **both** the subject and the comp,
using the local URA Master Plan land use + OneMap data.

> **Why these factors?** They reflect what a real-estate team already knows drives
> rent/value: offices trade on CBD access and business clustering; industrial on
> logistics connectivity; retail on catchment population and mall gravity;
> hotels on tourist draw and amenities.

---

### Step 2 — Score each factor, comp vs subject

Each factor is turned into a number between **−1 and +1**. There are two simple
recipes depending on whether the factor is a *count* or a *distance*.

#### Recipe A — "More is better" factors (counts)

> Example: number of commercial parcels within 1 km.

```
                  comp_count − subject_count
factor score  =  ──────────────────────────────
                  comp_count + subject_count + 10
```

Plain English:
- If the comp has **more** than the subject → **positive**.
- If the comp has **fewer** → **negative**.
- If they're equal → **zero**.
- The **"+ 10"** stops tiny numbers from looking dramatic.

| Subject | Comp | Raw difference | Factor score | Reading |
|--------:|-----:|---------------:|-------------:|---------|
| 1       | 2    | "double!"      | +0.08        | barely different (correct) |
| 45      | 30   | −15            | −0.18        | meaningfully fewer |
| 20      | 20   | 0              | 0.00         | identical |

> **Why "+ 10"?** Without it, going from 1 to 2 parcels would score +0.33 — treating
> a one-parcel difference as huge. That's noise, not signal. Adding 10 keeps small
> counts calm while still letting large, real differences show through. (As the
> numbers get big, the "+ 10" becomes irrelevant.)

#### Recipe B — "Closer is better" factors (distances)

> Example: distance to the CBD, in km.

```
                  subject_distance − comp_distance
factor score  =  ────────────────────────────────── , then capped to the −1…+1 range
                            5 km
```

Plain English:
- If the comp is **closer** than the subject → **positive**.
- If the comp is **farther** → **negative**.
- Being **5 km** closer/farther than the subject reaches the ±1 edge.

| Subject dist | Comp dist | Factor score | Reading |
|-------------:|----------:|-------------:|---------|
| 0.2 km       | 1.5 km    | −0.26        | comp is a bit farther |
| 3.0 km       | 1.0 km    | +0.40        | comp is notably closer |
| 6.0 km       | 0.5 km    | +1.00 (capped) | comp much closer |

> **Why divide by a fixed 5 km (not by the subject's distance)?** If a subject sits
> right on top of the CBD (distance ≈ 0), dividing by its distance would blow up and
> unfairly push **every** comp to −1. A fixed 5 km scale keeps things sensible: "5 km
> difference" is treated as a large but not extreme gap for Singapore. The cap stops
> a far-away outlier from dominating.

---

### Step 3 — Average the two factors

```
location score = (factor 1 score + factor 2 score) / 2
```

Both factors get **equal weight**.

> **Why equal weight?** Each sector's two drivers are roughly co-equal (an office
> needs *both* CBD access *and* a business cluster). Without a large labelled dataset
> to calibrate custom weights, equal weighting is the honest, transparent choice — and
> both factors are already on the same −1…+1 scale, so averaging is meaningful.

---

### Step 4 — Turn the number into a label

```
score > +0.3   →  Superior
score < −0.3   →  Inferior
otherwise      →  Comparable   (i.e. between −0.3 and +0.3)
```

> **Why the ±0.3 bands?** Analysts want a clear 3-way verdict, not a raw decimal.
> ±0.3 is a deliberately **conservative materiality threshold**: within that band the
> two locations are close enough to call "Comparable"; only a clear advantage or
> disadvantage gets flagged Superior or Inferior.

---

## 3. Full worked example

**Subject:** CapitaSpring (office, CBD). **Comp:** South Beach.

Office factors = (distance to CBD ↓, commercial parcels within 1 km ↑).

Suppose:

| Measure                       | Subject | Comp (South Beach) |
|-------------------------------|--------:|-------------------:|
| Distance to CBD               | 0.2 km  | 1.5 km             |
| Commercial parcels within 1 km| 45      | 30                 |

**Factor 1 (distance, closer is better):**
```
(0.2 − 1.5) / 5  =  −0.26
```

**Factor 2 (count, more is better):**
```
(30 − 45) / (30 + 45 + 10)  =  −0.18
```

**Average:**
```
(−0.26 + −0.18) / 2  =  −0.22   →  |−0.22| ≤ 0.3  →  Comparable (slightly weaker)
```

*(The live run used exact coordinates and returned −0.317 → **Inferior**, just past
the band — showing how the label reacts near the boundary.)*

---

## 4. Where the numbers come from (data & precision)

- **Coordinates**: the same map-resolved latitude/longitude used to plot the pin
  (Google / Mapbox), so the score is consistent with what you see on the map.
  OneMap is used only as a fallback when coordinates are missing.
- **Land-use counts & distances**: the **local URA Master Plan** GeoJSON
  (`backend/data/MasterPlan2025.geojson`) — fully on-premise, no network needed.
- **Tourist attractions**: OneMap "Tourist Attractions" theme, cached locally.
- A comp is scored **only** if (a) it resolves to coordinates and (b) it is the
  **same sector** as the subject (mixed matches anything). Otherwise its Location is
  left **blank** — we never guess across different asset types.

---

## 5. The tunable "knobs" (all one-line changes)

Every judgement call is a single constant, so the model can be tuned against real
deals without touching the logic:

| Knob            | Current | Controls |
|-----------------|--------:|----------|
| Count smoothing | `10`    | How much small counts are damped |
| Distance scale  | `5 km`  | What counts as a "large" distance gap |
| Label bands     | `±0.3`  | How different before Superior / Inferior |
| Factor weights  | `50/50` | Relative importance of the two factors |

---

### Summary

> **raw geography → sector-specific factors → each factor scored −1…+1 against the
> subject (subject = 0) → average the two → apply ±0.3 bands →
> Superior / Comparable / Inferior.**

Every step is deterministic, explainable, and on-premise — no black box, no cloud model.
