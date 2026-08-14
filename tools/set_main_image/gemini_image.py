# -*- coding: utf-8 -*-
"""Gemini Image（Nano Banana）呼び出し。APIキーは環境変数 or secrets ファイル。"""
from __future__ import annotations

import base64
import logging
import os
from pathlib import Path
from typing import List, Optional, Tuple, Any, Dict

LOG = logging.getLogger("set_main_image.gemini")

SECRET_CANDIDATES = (
    Path(__file__).resolve().parent / "secrets" / "gemini_api_key.txt",
    Path(__file__).resolve().parent / "secrets" / "GEMINI_API_KEY.txt",
)


def load_api_key() -> str:
    for env_name in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
        v = (os.environ.get(env_name) or "").strip()
        if v:
            LOG.info("API key from env %s", env_name)
            return v
    for p in SECRET_CANDIDATES:
        if p.is_file():
            text = p.read_text(encoding="utf-8").strip()
            if text and not text.startswith("#"):
                LOG.info("API key from file %s", p.name)
                return text
    raise SystemExit(
        "GEMINI_API_KEY がありません。"
        "環境変数 GEMINI_API_KEY を設定するか、"
        "tools/set_main_image/secrets/gemini_api_key.txt に1行で置いてください（gitignore済）。"
    )


def mime_for(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in (".jpg", ".jpeg"):
        return "image/jpeg"
    if ext == ".png":
        return "image/png"
    if ext == ".webp":
        return "image/webp"
    return "image/jpeg"


def _b64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def make_client(api_key: Optional[str] = None):
    try:
        from google import genai
    except ImportError as e:
        raise SystemExit(
            "google-genai 未インストール。"
            " pip install -r tools/set_main_image/requirements-ai.txt"
        ) from e
    key = api_key or load_api_key()
    return genai.Client(api_key=key)


def generate_with_references(
    *,
    client,
    model_id: str,
    prompt: str,
    image_paths: List[Path],
    image_roles: Optional[List[str]] = None,
    aspect_ratio: str = "1:1",
    image_size: str = "1K",
) -> Tuple[bytes, Dict[str, Any]]:
    """
    テキスト＋参照画像 → (画像バイト, meta)。
    各画像の直前に role ラベルを挿入し、見本がどのスロットか追跡可能にする。
    """
    from typing import Any, Dict

    roles = image_roles or [f"IMAGE_{i+1}" for i in range(len(image_paths))]
    if len(roles) != len(image_paths):
        raise ValueError("image_roles と image_paths の件数が一致しません")

    inputs: List[dict] = [{"type": "text", "text": prompt}]
    labeled: List[Dict[str, Any]] = []
    for role, p in zip(roles, image_paths):
        label = f"<<<{role} filename={p.name}>>> Follow the prompt rules for this role."
        inputs.append({"type": "text", "text": label})
        inputs.append(
            {
                "type": "image",
                "data": _b64(p),
                "mime_type": mime_for(p),
            }
        )
        labeled.append({"role": role, "path": str(p), "name": p.name})

    api_path = "interactions"
    try:
        interaction = client.interactions.create(
            model=model_id,
            input=inputs,
            response_format={
                "type": "image",
                "aspect_ratio": aspect_ratio,
                "image_size": image_size,
            },
        )
        img = getattr(interaction, "output_image", None)
        if img is not None and getattr(img, "data", None):
            raw = img.data
            data = base64.b64decode(raw) if isinstance(raw, str) else bytes(raw)
            return data, {"apiPath": api_path, "labeledInputs": labeled}
        outputs = getattr(interaction, "outputs", None) or getattr(interaction, "output", None)
        data = _extract_image_bytes_from_unknown(outputs) or _extract_image_bytes_from_unknown(interaction)
        if data:
            return data, {"apiPath": api_path, "labeledInputs": labeled}
        LOG.warning("interactions: output_image なし → generateContent へ")
    except Exception as e:
        LOG.warning("interactions 失敗 (%s) → generateContent へ", e)

    api_path = "generateContent"
    data = _generate_content_fallback(
        client=client,
        model_id=model_id,
        prompt=prompt,
        image_paths=image_paths,
        image_roles=roles,
    )
    return data, {"apiPath": api_path, "labeledInputs": labeled}


def _extract_image_bytes_from_unknown(obj) -> Optional[bytes]:
    if obj is None:
        return None
    if isinstance(obj, (bytes, bytearray)):
        return bytes(obj)
    if isinstance(obj, str) and len(obj) > 100:
        try:
            return base64.b64decode(obj)
        except Exception:
            return None
    if isinstance(obj, dict):
        for k in ("data", "image_bytes", "b64_json"):
            if k in obj and obj[k]:
                v = obj[k]
                return base64.b64decode(v) if isinstance(v, str) else bytes(v)
        for v in obj.values():
            got = _extract_image_bytes_from_unknown(v)
            if got:
                return got
    if isinstance(obj, (list, tuple)):
        for v in obj:
            got = _extract_image_bytes_from_unknown(v)
            if got:
                return got
    # SDKオブジェクト
    for attr in ("data", "inline_data", "image", "parts", "content", "outputs", "output"):
        if hasattr(obj, attr):
            got = _extract_image_bytes_from_unknown(getattr(obj, attr))
            if got:
                return got
    if hasattr(obj, "model_dump"):
        try:
            return _extract_image_bytes_from_unknown(obj.model_dump())
        except Exception:
            pass
    return None


def _generate_content_fallback(
    *,
    client,
    model_id: str,
    prompt: str,
    image_paths: List[Path],
    image_roles: Optional[List[str]] = None,
) -> bytes:
    from google.genai import types
    from PIL import Image as PILImage

    roles = image_roles or [f"IMAGE_{i+1}" for i in range(len(image_paths))]
    parts: list = [prompt]
    for role, p in zip(roles, image_paths):
        parts.append(f"<<<{role} filename={p.name}>>>")
        parts.append(PILImage.open(p))

    config = None
    try:
        config = types.GenerateContentConfig(
            response_modalities=["TEXT", "IMAGE"],
        )
    except Exception:
        config = None

    kwargs = {"model": model_id, "contents": parts}
    if config is not None:
        kwargs["config"] = config

    resp = client.models.generate_content(**kwargs)
    data = _extract_from_generate_content(resp)
    if not data:
        raise RuntimeError(f"画像が返りませんでした model={model_id}")
    return data


def _extract_from_generate_content(resp) -> Optional[bytes]:
    cands = getattr(resp, "candidates", None) or []
    for cand in cands:
        content = getattr(cand, "content", None)
        parts = getattr(content, "parts", None) or []
        for part in parts:
            inline = getattr(part, "inline_data", None)
            if inline is not None:
                raw = getattr(inline, "data", None)
                if raw is None:
                    continue
                if isinstance(raw, str):
                    return base64.b64decode(raw)
                return bytes(raw)
            # 新しいフィールド名
            if getattr(part, "as_image", None):
                try:
                    im = part.as_image()
                    import io
                    buf = io.BytesIO()
                    im.save(buf, format="PNG")
                    return buf.getvalue()
                except Exception:
                    pass
    return _extract_image_bytes_from_unknown(resp)


def save_as_jpeg(raw: bytes, out_path: Path, quality: int = 85) -> Tuple[Path, dict]:
    from io import BytesIO
    from PIL import Image

    out_path.parent.mkdir(parents=True, exist_ok=True)
    im = Image.open(BytesIO(raw))
    if im.mode in ("RGBA", "P"):
        bg = Image.new("RGB", im.size, (255, 255, 255))
        if im.mode == "P":
            im = im.convert("RGBA")
        bg.paste(im, mask=im.split()[-1] if im.mode == "RGBA" else None)
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")
    im.save(out_path, format="JPEG", quality=quality, optimize=True)
    meta = {"width": im.size[0], "height": im.size[1], "quality": quality}
    return out_path, meta
