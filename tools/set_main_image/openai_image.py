# -*- coding: utf-8 -*-
"""OpenAI Image（ChatGPT画像）呼び出し。キーは環境変数 or secrets。"""
from __future__ import annotations

import base64
import logging
import os
from pathlib import Path
from typing import List, Optional, Tuple

LOG = logging.getLogger("set_main_image.openai")

SECRET_CANDIDATES = (
    Path(__file__).resolve().parent / "secrets" / "openai_api_key.txt",
    Path(__file__).resolve().parent / "secrets" / "OPENAI_API_KEY.txt",
)

# 精度比較用。2026-08 時点の最新は gpt-image-2（旧 gpt-image-1 は後方互換）。
DEFAULT_OPENAI_IMAGE_MODEL = "gpt-image-2"
OPENAI_IMAGE_FALLBACKS = ("gpt-image-2", "gpt-image-1.5", "gpt-image-1")
ENV_MODEL_OVERRIDE = "SET_MAIN_OPENAI_IMAGE_MODEL"


def load_openai_api_key() -> str:
    v = (os.environ.get("OPENAI_API_KEY") or "").strip()
    if v:
        LOG.info("API key from env OPENAI_API_KEY")
        return v
    for p in SECRET_CANDIDATES:
        if p.is_file():
            text = p.read_text(encoding="utf-8").strip()
            if text and not text.startswith("#"):
                LOG.info("API key from file %s", p.name)
                return text
    raise SystemExit(
        "OPENAI_API_KEY がありません。"
        "環境変数 OPENAI_API_KEY か "
        "tools/set_main_image/secrets/openai_api_key.txt を用意してください。"
    )


def resolve_openai_image_model(explicit: Optional[str] = None) -> str:
    if explicit and explicit.strip():
        return explicit.strip()
    env = (os.environ.get(ENV_MODEL_OVERRIDE) or "").strip()
    if env:
        return env
    return DEFAULT_OPENAI_IMAGE_MODEL


def generate_with_references(
    *,
    prompt: str,
    image_paths: List[Path],
    model: Optional[str] = None,
    size: str = "1024x1024",
    quality: str = "high",
    api_key: Optional[str] = None,
) -> Tuple[bytes, str]:
    """戻り: (画像バイト, 実際に使った model_id)。"""
    try:
        from openai import OpenAI
    except ImportError as e:
        raise SystemExit(
            "openai 未インストール。 pip install openai"
        ) from e

    client = OpenAI(api_key=api_key or load_openai_api_key())
    primary = resolve_openai_image_model(model)
    candidates = [primary] + [m for m in OPENAI_IMAGE_FALLBACKS if m != primary]
    last_err: Optional[Exception] = None
    for model_id in candidates:
        files = []
        try:
            for p in image_paths:
                files.append(open(p, "rb"))
            LOG.info("openai images.edit model=%s n_images=%s", model_id, len(files))
            kwargs = {
                "model": model_id,
                "image": files if len(files) != 1 else files[0],
                "prompt": prompt,
                "size": size,
            }
            # quality はモデルにより非対応のことがある
            try:
                result = client.images.edit(**kwargs, quality=quality)
            except TypeError:
                result = client.images.edit(**kwargs)
            except Exception as e:
                # quality 非対応など
                msg = str(e).lower()
                if "quality" in msg:
                    result = client.images.edit(**kwargs)
                else:
                    raise
            data0 = result.data[0]
            b64 = getattr(data0, "b64_json", None)
            if b64:
                return base64.b64decode(b64), model_id
            url = getattr(data0, "url", None)
            if url:
                import urllib.request
                with urllib.request.urlopen(url) as resp:
                    return resp.read(), model_id
            raise RuntimeError(f"OpenAI画像が返りませんでした model={model_id}")
        except Exception as e:
            last_err = e
            LOG.warning("openai model=%s failed: %s", model_id, e)
            msg = str(e).lower()
            # クレジット枯渇はフォールバックしても同じ → 即打ち切り
            if (
                "insufficient_quota" in msg
                or "credit_balance_exhausted" in msg
                or "no credits remaining" in msg
            ):
                raise RuntimeError(f"OpenAI画像生成に失敗（クレジット不足）: {e}") from e
        finally:
            for f in files:
                try:
                    f.close()
                except Exception:
                    pass
    raise RuntimeError(f"OpenAI画像生成に失敗: {last_err}")
