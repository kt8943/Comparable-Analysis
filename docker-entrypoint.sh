#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
# Container entrypoint for Hugging Face Spaces.
#
# Generates configs/shared_settings.json from Space *Secrets* (env vars) at
# startup, so no API keys are ever committed to the repo. Then launches Streamlit
# on the port HF expects.
#
# Set these as Space Secrets (Settings → Variables and secrets):
#   OPENAI_API_KEY      (required — the LLM)
#   MAPBOX_TOKEN        (maps + geocoding fallback)
#   GOOGLE_MAPS_KEY     (optional — Google geocoding)
#   KAKAO_API_KEY       (optional — Korean geocoding)
#   GEOCODING_PROVIDER  (optional — google | onemap | kakao | mapbox; default mapbox)
# ─────────────────────────────────────────────────────────────────────────────
set -e

mkdir -p /app/configs
cat > /app/configs/shared_settings.json <<EOF
{
  "geocoding_provider": "${GEOCODING_PROVIDER:-mapbox}",
  "mapbox_token": "${MAPBOX_TOKEN:-}",
  "google_maps_key": "${GOOGLE_MAPS_KEY:-}",
  "kakao_api_key": "${KAKAO_API_KEY:-}",
  "openai_api_key": "${OPENAI_API_KEY:-}"
}
EOF

exec streamlit run frontend/app.py \
    --server.port=8501 \
    --server.address=0.0.0.0 \
    --server.headless=true \
    --server.enableCORS=false \
    --server.enableXsrfProtection=false
