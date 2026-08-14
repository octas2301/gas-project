# -*- coding: utf-8 -*-
"""楽天マトリクス向けファイル命名契約（SKU 自動紐付け）。

本線:
  MAIN … {childSku}_rakuten.jpg（compose_set_main と同契約）
  サブ … {childSku}_{pattern}_subN.jpg（N=1..10）
       pattern 省略時は従来互換の {childSku}_subN.jpg

GAS 側はファイル名に子SKUを含み、_subN / _pN があればサブ枠へ投入する。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Optional


def normalize_for_match(s: object) -> str:
    """Amazon U2-ε / 楽天 autobind と同型（小文字＋英数字以外除去）。"""
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def sanitize_pattern_slug(pattern: object) -> str:
    """ファイル名用パターン（themeSlug 等）。空なら空文字。"""
    raw = str(pattern or "").strip()
    if not raw:
        return ""
    # パス区切り・空白を安全化（SKU照合用の英数字は残す）
    s = re.sub(r"[\\/:\*\?\"<>\|]+", "", raw)
    s = re.sub(r"\s+", "_", s)
    s = s.strip("._")
    return s[:80]


def main_filename(child_sku: str, *, mall: str = "rakuten") -> str:
    sku = str(child_sku or "").strip()
    if not sku:
        raise ValueError("child_sku is required")
    m = str(mall or "rakuten").strip().lower() or "rakuten"
    return f"{sku}_{m}.jpg"


def sub_filename(
    child_sku: str,
    sub_index: int,
    *,
    pattern: object = "",
) -> str:
    """
    サブ画像ファイル名。
    パターンあり: {sku}_{pattern}_subN.jpg
    なし: {sku}_subN.jpg（互換）
    """
    sku = str(child_sku or "").strip()
    if not sku:
        raise ValueError("child_sku is required")
    n = int(sub_index)
    if n < 1 or n > 10:
        raise ValueError(f"sub_index must be 1..10, got {sub_index!r}")
    slug = sanitize_pattern_slug(pattern)
    if slug:
        return f"{sku}_{slug}_sub{n}.jpg"
    return f"{sku}_sub{n}.jpg"


def parse_sub_index(file_name: str) -> Optional[int]:
    """ファイル名からサブ枠 1..10。無ければ None（MAIN扱い）。"""
    name = Path(str(file_name or "")).name
    m = re.search(r"_sub(\d+)", name, flags=re.I)
    if not m:
        m = re.search(r"_p(\d+)", name, flags=re.I)
    if not m:
        return None
    n = int(m.group(1))
    if n < 1 or n > 10:
        return None
    return n
