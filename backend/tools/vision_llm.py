"""
tools/vision_llm.py
===================
Vision LLM transport layer — send an image + prompt to OpenAI or Ollama
and return the raw text response.

Public API
----------
call_vision_llm(image_path, prompt, llm_cfg, openai_key="",
                vision_models=None) -> str
    Returns the raw LLM response string.
    Raises ValueError if the model does not support vision.
"""

import base64
import json
import urllib.request
from pathlib import Path

# Default set of Ollama model prefixes that accept image input.
_DEFAULT_VISION_MODELS: set = {
    "llava", "llava-llama3", "llava-phi3", "bakllava",
    "moondream", "minicpm-v", "llama3.2-vision",
}


def call_vision_llm(image_path: str, prompt: str, llm_cfg: dict,
                    openai_key: str = "",
                    vision_models: set = None) -> str:
    """Send an image + prompt to a vision-capable LLM and return the raw response.

    Supports:
    - OpenAI  (gpt-4o, gpt-4o-mini) — best quality, requires API key
    - Ollama  (llava, minicpm-v, etc.) — local, model must support vision

    Raises ValueError if the configured model does not support vision.

    vision_models: set of Ollama model name prefixes that support vision.
                   Defaults to _DEFAULT_VISION_MODELS if not provided.
    """
    _vmodels = vision_models if vision_models is not None else _DEFAULT_VISION_MODELS

    suffix = Path(image_path).suffix.lower()
    mime   = "image/png" if suffix == ".png" else "image/jpeg"
    with open(image_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    provider = llm_cfg.get("provider", "ollama")

    if provider == "openai":
        key   = openai_key or llm_cfg.get("openai_api_key", "")
        model = llm_cfg.get("openai_model", "gpt-4o-mini")
        if not key:
            raise ValueError("OpenAI API key required for image input.")
        try:
            import openai as _openai
        except ImportError:
            raise ImportError("pip install openai")
        client = _openai.OpenAI(api_key=key)
        resp   = client.chat.completions.create(
            model=model,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url",
                     "image_url": {"url": f"data:{mime};base64,{b64}"}},
                ],
            }],
            temperature=0.1,
        )
        return resp.choices[0].message.content.strip()

    else:  # Ollama
        ollama = llm_cfg.get("ollama", {})
        # Use dedicated vision_model if set (from sidebar Vision model selector),
        # otherwise fall back to the main analysis model.
        model  = ollama.get("vision_model") or ollama.get("model", "")
        url    = ollama.get("base_url", "http://localhost:11434")
        model_base = model.split(":")[0].lower()
        if model_base not in _vmodels:
            raise ValueError(
                f"Model '{model}' does not support vision (image input).\n"
                f"Select a vision model (e.g. minicpm-v, llava) in the sidebar "
                f"👁️ Vision model selector, or switch provider to OpenAI."
            )
        payload = json.dumps({
            "model":    model,
            "messages": [{
                "role":    "user",
                "content": prompt,
                "images":  [b64],
            }],
            "stream":  False,
            # num_predict: allow a generous output budget so the model can
            # produce a full JSON array for tables with many rows.
            # Without this some models stop mid-array after ~256 tokens.
            "options": {"temperature": 0.1, "num_predict": 4096},
        }).encode()
        req = urllib.request.Request(
            f"{url.rstrip('/')}/api/chat",
            data=payload,
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=300) as resp:
            return json.loads(resp.read())["message"]["content"].strip()
