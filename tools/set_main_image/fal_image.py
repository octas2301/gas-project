# -*- coding: utf-8 -*-
"""
fal.ai 画像生成（PoC）

- 既定（参照画像あり）: fal-ai/flux-kontext/dev … 編集向け・格安帯
- 参照なし / フォールバック: fal-ai/flux/schnell … 最安テキスト→画像
- キー: 環境変数 FAL_KEY または secrets/fal_api_key.txt
"""
from __future__ import annotations

import logging
import os
import urllib.request
from pathlib import Path
from typing import List, Optional, Tuple

LOG = logging.getLogger("set_main_image.fal")

SECRET_CANDIDATES = (
    Path(__file__).resolve().parent / "secrets" / "fal_api_key.txt",
    Path(__file__).resolve().parent / "secrets" / "FAL_KEY.txt",
)

# 参照画像あり（競合パーツ流用向け）
DEFAULT_EDIT_ENDPOINT = "fal-ai/flux-kontext/dev"
# テキストのみ（最安）
DEFAULT_TXT_ENDPOINT = "fal-ai/flux/schnell"
# 高品質編集（高い）
PRO_EDIT_ENDPOINT = "fal-ai/flux-pro/kontext"
# FLUX.2 pro 編集（image_urls 形式）
FLUX2_PRO_EDIT_ENDPOINT = "fal-ai/flux-2-pro/edit"

ENV_KEY = "FAL_KEY"
ENV_EDIT = "SET_MAIN_FAL_EDIT_ENDPOINT"
ENV_TXT = "SET_MAIN_FAL_TXT_ENDPOINT"


def load_fal_api_key() -> str:
    v = (os.environ.get(ENV_KEY) or os.environ.get("FAL_API_KEY") or "").strip()
    if v:
        LOG.info("fal key from env")
        return v
    for p in SECRET_CANDIDATES:
        if p.is_file():
            text = p.read_text(encoding="utf-8").strip()
            if text and not text.startswith("#"):
                LOG.info("fal key from file %s", p.name)
                return text
    raise SystemExit(
        "FAL_KEY がありません。"
        "環境変数 FAL_KEY か tools/set_main_image/secrets/fal_api_key.txt を用意してください。"
        "キー発行: https://fal.ai/dashboard/keys"
    )


def resolve_edit_endpoint(explicit: Optional[str] = None) -> str:
    if explicit and explicit.strip():
        return explicit.strip()
    return (os.environ.get(ENV_EDIT) or DEFAULT_EDIT_ENDPOINT).strip()


def resolve_txt_endpoint(explicit: Optional[str] = None) -> str:
    if explicit and explicit.strip():
        return explicit.strip()
    return (os.environ.get(ENV_TXT) or DEFAULT_TXT_ENDPOINT).strip()


def _ensure_client(api_key: Optional[str] = None):
    try:
        import fal_client
    except ImportError as e:
        raise SystemExit(
            "fal-client 未インストール。 "
            "pip install fal-client"
        ) from e
    key = api_key or load_fal_api_key()
    os.environ[ENV_KEY] = key
    return fal_client


def upload_image(path: Path, *, api_key: Optional[str] = None) -> str:
    """ローカル画像を fal storage に上げて URL を返す。"""
    fal_client = _ensure_client(api_key)
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(str(p))
    url = fal_client.upload_file(str(p))
    LOG.info("fal upload %s → %s", p.name, url[:80])
    return str(url)


def _download(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "gas-project-fal-poc/1.0"})
    with urllib.request.urlopen(req, timeout=120) as resp:
        return resp.read()


def _extract_image_bytes(result) -> bytes:
    if not isinstance(result, dict):
        result = dict(result) if result is not None else {}
    images = result.get("images") or []
    if not images:
        raise RuntimeError(f"fal: images が空です keys={list(result.keys())}")
    first = images[0]
    if isinstance(first, dict):
        url = first.get("url") or ""
        data_uri = first.get("file_data") or first.get("data") or ""
        if url:
            return _download(url)
        if isinstance(data_uri, str) and data_uri.startswith("data:"):
            import base64

            b64 = data_uri.split(",", 1)[-1]
            return base64.b64decode(b64)
    raise RuntimeError(f"fal: 画像URLを取得できません: {first!r}")


def generate_text_to_image(
    *,
    prompt: str,
    endpoint: Optional[str] = None,
    image_size: str = "square_hd",
    num_inference_steps: int = 4,
    api_key: Optional[str] = None,
) -> Tuple[bytes, str]:
    """戻り: (jpeg/png bytes, endpoint_id)"""
    fal_client = _ensure_client(api_key)
    ep = resolve_txt_endpoint(endpoint)
    LOG.info("fal txt2img endpoint=%s", ep)
    result = fal_client.subscribe(
        ep,
        arguments={
            "prompt": prompt,
            "image_size": image_size,
            "num_images": 1,
            "num_inference_steps": num_inference_steps,
            "enable_safety_checker": True,
            "output_format": "jpeg",
        },
        with_logs=False,
    )
    return _extract_image_bytes(result), ep


def _is_flux2_edit_endpoint(endpoint: str) -> bool:
    ep = (endpoint or "").lower()
    return "flux-2" in ep and "/edit" in ep


def generate_with_references(
    *,
    prompt: str,
    image_paths: List[Path],
    endpoint: Optional[str] = None,
    api_key: Optional[str] = None,
    guidance_scale: float = 2.5,
    num_inference_steps: Optional[int] = None,
    allow_txt_fallback: bool = True,
) -> Tuple[bytes, str]:
    """
    参照画像つき編集。
    - Kontext 系: image_url + guidance_scale
    - FLUX.2 */edit: image_urls（複数可）
    参照0枚なら schnell テキスト生成にフォールバック。
    戻り: (bytes, 使用 endpoint)
    """
    paths = [Path(p) for p in image_paths if Path(p).is_file()]
    if not paths:
        return generate_text_to_image(prompt=prompt, api_key=api_key)

    fal_client = _ensure_client(api_key)
    ep = resolve_edit_endpoint(endpoint)
    urls = [upload_image(p, api_key=api_key) for p in paths]
    extra = ""
    if len(paths) > 1 and not _is_flux2_edit_endpoint(ep):
        extra = (
            f"\nAdditional reference images were considered conceptually "
            f"({len(paths)-1} more competitor frames). Keep product identity "
            f"consistent with the primary reference."
        )

    if _is_flux2_edit_endpoint(ep):
        args = {
            "prompt": prompt + extra,
            "image_urls": urls,
            "output_format": "jpeg",
            "safety_tolerance": "2",
            "enable_safety_checker": True,
            "image_size": "square_hd",
        }
    else:
        args = {
            "prompt": prompt + extra,
            "image_url": urls[0],
            "num_images": 1,
            "guidance_scale": guidance_scale,
            "enable_safety_checker": True,
            "output_format": "jpeg",
            "resolution_mode": "1:1",
        }
        if num_inference_steps is not None:
            args["num_inference_steps"] = int(num_inference_steps)

    LOG.info("fal edit endpoint=%s refs=%s", ep, len(paths))
    try:
        result = fal_client.subscribe(ep, arguments=args, with_logs=False)
        return _extract_image_bytes(result), ep
    except Exception as e:
        if not allow_txt_fallback:
            raise
        LOG.warning("fal edit failed (%s) → schnell txt2img fallback: %s", ep, e)
        return generate_text_to_image(prompt=prompt, api_key=api_key)
