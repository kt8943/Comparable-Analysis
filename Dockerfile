# ─────────────────────────────────────────────────────────────────────────────
# PGIM Comparable Analysis — Hugging Face Spaces (Docker SDK) image
#
# Public DEMO deployment. Uses OpenAI (cloud) instead of Ollama.
# Do NOT deploy with confidential deal data — this is a public URL.
# ─────────────────────────────────────────────────────────────────────────────
FROM python:3.11-slim

# System dependencies:
#   ghostscript      → camelot table extraction
#   libgl1/libglib   → OpenCV (camelot-py[cv])
#   python3-tk       → camelot
RUN apt-get update && apt-get install -y --no-install-recommends \
        ghostscript \
        python3-tk \
        libgl1 \
        libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Run as the non-root user Hugging Face Spaces expects (uid 1000).
RUN useradd -m -u 1000 appuser

ENV HOME=/home/appuser \
    PYTHONUNBUFFERED=1 \
    PYTHONUTF8=1 \
    PYTHONIOENCODING=utf-8 \
    FASTEMBED_CACHE_PATH=/home/appuser/.fastembed_cache \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

WORKDIR /app

# Python deps first so the layer caches across code changes.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code (see .dockerignore — real configs/, Input_files/, output/ are excluded).
COPY . .

# Writable dirs + warm the embedding model at build (ignore failure if no network).
RUN mkdir -p /app/configs /app/output /app/Input_files /home/appuser/.fastembed_cache \
    && python -c "from fastembed import TextEmbedding; TextEmbedding('sentence-transformers/all-MiniLM-L6-v2')" || true
RUN chmod +x /app/docker-entrypoint.sh \
    && chown -R appuser:appuser /app /home/appuser

USER appuser
EXPOSE 8501
ENTRYPOINT ["/app/docker-entrypoint.sh"]
