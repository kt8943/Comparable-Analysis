#!/usr/bin/env python3
"""
warm_embed.py — pre-download + cache the embedding model used by the comp pipeline.
====================================================================================
The column mapper (backend/tools/column_mapper.py) uses fastembed's
``sentence-transformers/all-MiniLM-L6-v2`` for Tier-2 semantic column matching.
fastembed downloads that model the first time it is used, then loads it from a
local on-disk cache on every run after that.

Run this ONCE on a new machine to do that download deliberately and verify it
works — afterwards the pipeline loads the model from cache with no network call:

    python warm_embed.py

IMPORTANT: run it with the SAME Python interpreter you use to launch the app
(the app runs backend steps as subprocesses via sys.executable). For example, if
you start the app with `python3.14 -m streamlit run frontend/app.py`, then run
`python3.14 warm_embed.py` so the cache matches.
"""

import sys

MODEL = "sentence-transformers/all-MiniLM-L6-v2"  # must match column_mapper.py


def main() -> int:
    try:
        from fastembed import TextEmbedding
    except ImportError:
        print("ERROR: fastembed is not installed for this Python "
              f"({sys.executable}).\n"
              "Install it (e.g. from your offline wheels) and try again.")
        return 1

    print(f"Python      : {sys.executable}")
    print(f"Caching     : {MODEL}")
    print("Downloading (first run only) ...")
    try:
        model = TextEmbedding(MODEL)               # triggers the one-time download
        vec = list(model.embed(["hello world"]))   # force a real embed to confirm
    except Exception as e:
        print(f"\nFAILED to download/load the model: {e.__class__.__name__}: {e}")
        print("If this is an SSL / certificate error, the corporate proxy is "
              "blocking the download — let me know and we'll point fastembed at "
              "the corporate CA.")
        return 1

    print(f"\nOK — model cached and working. Embedding dim = {len(vec[0])}")
    print("The pipeline will now load it from cache with no network call.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
