---
title: PGIM Comparable Analysis
emoji: 🏢
colorFrom: blue
colorTo: indigo
sdk: docker
app_port: 8501
pinned: false
---

# PGIM Comparable Analysis — Demo

Agentic AI document-automation pipeline for real-estate comparable transactions:
extracts comps from PDF / Excel / image market reports, maps columns to a schema,
geocodes and scores them, and renders Excel tables + maps.

> ⚠️ **Public demo — use sample / non-confidential data only.**
> This deployment uses the OpenAI API (not on-prem Ollama). Do not upload
> confidential deal data to a public Space.

## Configuration
API keys are provided via **Space Secrets** (Settings → Variables and secrets) and
injected at startup — nothing is stored in the repo:

| Secret | Required | Purpose |
|--------|----------|---------|
| `OPENAI_API_KEY` | ✅ | LLM (classification, mapping fallback, rationale) |
| `MAPBOX_TOKEN` | recommended | static map rendering + geocoding fallback |
| `GOOGLE_MAPS_KEY` | optional | Google geocoding |
| `KAKAO_API_KEY` | optional | Korean-address geocoding |
| `GEOCODING_PROVIDER` | optional | `google` \| `onemap` \| `kakao` \| `mapbox` (default `mapbox`) |

In the sidebar model selector, choose a **GPT** model (Ollama is not available on
the cloud). "🚫 Rule-based (no LLM)" also works for testing geocoding.
