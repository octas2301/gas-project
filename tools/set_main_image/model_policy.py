# -*- coding: utf-8 -*-
"""
Nano Banana（Gemini Image）モデル選定。

方針（社長ロック 2026-08-02）:
- 常に「通常（Flash）」系の最新を使う
- PRO（gemini-*-pro-image）は使わない
- Lite（flash-lite-image）も既定では使わない（通常 Flash Image）
"""
from __future__ import annotations

import logging
import os
import re
from typing import Iterable, List, Optional

LOG = logging.getLogger("set_main_image.model")

# 解決失敗時のフォールバック（2026-08 時点の通常 Flash = Nano Banana 2）
FALLBACK_FLASH_IMAGE = "gemini-3.1-flash-image"

ENV_MODEL_OVERRIDE = "SET_MAIN_IMAGE_MODEL"


def _norm(name: str) -> str:
    return (name or "").strip().removeprefix("models/")


def is_pro_image(model_id: str) -> bool:
    n = _norm(model_id).lower()
    return "pro-image" in n or n.endswith("-pro") and "image" in n


def is_lite_image(model_id: str) -> bool:
    return "flash-lite-image" in _norm(model_id).lower() or "lite-image" in _norm(model_id).lower()


def is_standard_flash_image(model_id: str) -> bool:
    """通常 Flash Image（非PRO・非Lite）。例: gemini-3.1-flash-image"""
    n = _norm(model_id).lower()
    if "image" not in n:
        return False
    if is_pro_image(n) or is_lite_image(n):
        return False
    # 通常: *flash-image* （lite を含まない）
    if "flash-image" in n and "flash-lite" not in n:
        return True
    return False


def _version_key(model_id: str) -> tuple:
    """より新しい版を大きく。gemini-3.1-flash-image > gemini-2.5-flash-image"""
    n = _norm(model_id).lower()
    nums = [int(x) for x in re.findall(r"\d+", n)]
    # プレビューは下げたい場合は減点
    preview_penalty = 0 if "preview" in n else 1
    return (preview_penalty, nums)


def pick_latest_standard_flash(model_ids: Iterable[str]) -> Optional[str]:
    cands = [ _norm(m) for m in model_ids if is_standard_flash_image(m) ]
    if not cands:
        return None
    cands.sort(key=_version_key, reverse=True)
    return cands[0]


def list_image_model_ids(client) -> List[str]:
    """google.genai Client からモデル一覧を取得。"""
    out: List[str] = []
    try:
        for m in client.models.list():
            name = _norm(getattr(m, "name", "") or "")
            if name:
                out.append(name)
    except Exception as e:
        LOG.warning("models.list 失敗: %s", e)
    return out


def resolve_model_id(client=None, explicit: Optional[str] = None) -> str:
    """
    解決順:
    1. CLI --model / 引数 explicit
    2. 環境変数 SET_MAIN_IMAGE_MODEL
    3. API models.list から通常 Flash Image の最新
    4. FALLBACK_FLASH_IMAGE
    """
    if explicit and explicit.strip():
        mid = _norm(explicit)
        if is_pro_image(mid):
            raise SystemExit(
                f"PROモデルは方針で禁止です: {mid}（通常 Flash Image を指定してください）"
            )
        LOG.info("model=explicit %s", mid)
        return mid

    env = (os.environ.get(ENV_MODEL_OVERRIDE) or "").strip()
    if env:
        mid = _norm(env)
        if is_pro_image(mid):
            raise SystemExit(
                f"環境変数 {ENV_MODEL_OVERRIDE} に PRO は使えません: {mid}"
            )
        LOG.info("model=env %s", mid)
        return mid

    if client is not None:
        latest = pick_latest_standard_flash(list_image_model_ids(client))
        if latest:
            LOG.info("model=latest-flash-from-list %s", latest)
            return latest
        LOG.warning("list から通常 Flash Image が見つからず fallback=%s", FALLBACK_FLASH_IMAGE)

    LOG.info("model=fallback %s", FALLBACK_FLASH_IMAGE)
    return FALLBACK_FLASH_IMAGE
