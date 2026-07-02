# Deploying to Hugging Face Spaces (Docker) — Step by Step

A public demo of the PGIM comps app, using **OpenAI** instead of Ollama.

> ⚠️ Public URL → **sample / non-confidential data only**.

## Files in this repo that make it work
| File | Role |
|------|------|
| `Dockerfile` | Python 3.11 + Ghostscript/OpenCV (for camelot) + deps; launches Streamlit on :8501 |
| `docker-entrypoint.sh` | Writes `configs/shared_settings.json` from Space **Secrets** at startup, then starts Streamlit |
| `.dockerignore` | Keeps real `configs/`, `Input_files/`, `output/`, keys OUT of the image |
| `README_HF.md` | Becomes the Space's `README.md` (has the `sdk: docker`, `app_port: 8501` header HF needs) |

## Prerequisites
- A Hugging Face account.
- An OpenAI API key (and a Mapbox token if you want maps).
- `git` and `git-lfs` installed locally (`git lfs install`).

## Steps

### 1. Create the Space
huggingface.co → **New Space** → give it a name → **SDK: Docker** → **Blank** → Create.

### 2. Clone the (empty) Space repo
```bash
git clone https://huggingface.co/spaces/<your-username>/<space-name>
cd <space-name>
git lfs install
git lfs track "*.geojson"          # the 181 MB URA Master Plan
```

### 3. Copy the app in
Copy these from `Desktop/PGIM` into the Space folder (NOT `configs/`, `Input_files/`,
`output/`, or `offline_packages/` — `.dockerignore` skips them anyway):
```
backend/            frontend/           requirements.txt
Dockerfile          docker-entrypoint.sh   .dockerignore
backend/data/MasterPlan2025.geojson    (via git-lfs)
```
Then use the deploy README as the Space's README:
```bash
mv README_HF.md README.md
```

### 4. Commit & push
```bash
git add .gitattributes README.md Dockerfile docker-entrypoint.sh .dockerignore \
        requirements.txt backend frontend
git commit -m "Deploy PGIM comps demo (Docker, OpenAI)"
git push
```
The Space will start building automatically (watch the **Logs** tab).

### 5. Add Secrets
Space → **Settings → Variables and secrets → New secret**:
| Name | Value |
|------|-------|
| `OPENAI_API_KEY` | `sk-...` |
| `MAPBOX_TOKEN` | `pk...` (optional but recommended) |
| `GOOGLE_MAPS_KEY` | optional |
| `GEOCODING_PROVIDER` | e.g. `mapbox` or `google` (optional) |

After adding secrets, **Restart** the Space so the entrypoint regenerates
`shared_settings.json`.

### 6. Use it
Open the Space URL. In the **sidebar model selector pick a GPT model** (Ollama isn't
available in the cloud). Upload a **sample** Excel/PDF and run.

## Notes & troubleshooting
- **First run is slow** — the embedding model + URA land-use cache build on first use
  (RAM is ample on HF free CPU tier).
- **Build too big / slow?** The GeoJSON is 181 MB; make sure it's tracked by git-lfs.
  To skip the Location/Zoning feature entirely for a lighter demo, just omit
  `backend/data/MasterPlan2025.geojson` — the code degrades gracefully.
- **Maps blank?** Set `MAPBOX_TOKEN` and restart.
- **LLM errors?** Confirm `OPENAI_API_KEY` is set and you picked a GPT model.
- **camelot errors?** The Dockerfile installs Ghostscript/OpenCV libs; rebuild if you
  changed the Dockerfile.
