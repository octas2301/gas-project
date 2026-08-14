# -*- coding: utf-8 -*-
"""
サブ画像ベース色（トンマナ）パレット — B-④で人が選ぶ2〜3択 PoC。

- 変えてよいのは背景・カード・余白のみ。PACKAGE_TRUTH の商品色は禁止。
- compose は --base-color / --tonmana で受け取る。
"""
from __future__ import annotations

from typing import Any, Dict, List

VERSION = "2026-08-10.1"
DEFAULT_ID = "beige"

PALETTE: List[Dict[str, str]] = [
    {
        "id": "beige",
        "label": "ベージュ（既定）",
        "ja": (
            "背景・カード・余白はすべてベージュ／サンド系で統一する。"
            "白ベタ全面・黒ベース・強い赤黄の全面背景は使わない。"
        ),
        "en": "soft warm beige / sand studio background and cards; no full white, black, or loud red/yellow wash.",
    },
    {
        "id": "warm_white",
        "label": "ウォームホワイト",
        "ja": (
            "背景・カード・余白はウォームホワイト／アイボリー系で統一する。"
            "冷たい純白ベタ・黒ベース・強い彩度の全面背景は使わない。"
        ),
        "en": "warm white / ivory studio background and cards; avoid cold pure-white wash, black base, or saturated full-frame colors.",
    },
    {
        "id": "soft_gray",
        "label": "ソフトグレー",
        "ja": (
            "背景・カード・余白は明るいソフトグレー／ニュートラルグレーで統一する。"
            "真っ黒・純白ベタ・強い彩度の全面背景は使わない。"
        ),
        "en": "light soft gray / neutral studio background and cards; no pure black, pure white wash, or saturated full-frame colors.",
    },
]

_BY_ID = {p["id"]: p for p in PALETTE}


def normalize_base_color(raw: object) -> str:
    s = str(raw or "").strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "beige": "beige",
        "sand": "beige",
        "default": "beige",
        "warm_white": "warm_white",
        "warmwhite": "warm_white",
        "ivory": "warm_white",
        "soft_gray": "soft_gray",
        "softgrey": "soft_gray",
        "gray": "soft_gray",
        "grey": "soft_gray",
    }
    return aliases.get(s, DEFAULT_ID if s not in _BY_ID else s)


def resolve_tonmana(base_color: object = None) -> Dict[str, str]:
    cid = normalize_base_color(base_color)
    p = _BY_ID.get(cid) or _BY_ID[DEFAULT_ID]
    return dict(p)


def tonmana_block_ja(base_color: object = None) -> str:
    p = resolve_tonmana(base_color)
    return (
        f"【トンマナ／ベース色={p['id']}（{p['label']}）】\n"
        f"- {p['ja']}\n"
        "- ベース色の変更は背景・カード・余白のみ。"
        "IMAGE_PACKAGE_TRUTH の商品パッケージ色相・ラベル色は絶対に変えない。"
    )


def tonmana_block_en(base_color: object = None) -> str:
    p = resolve_tonmana(base_color)
    return (
        f"Base color / tonmana id={p['id']}: {p['en']} "
        "Change background/cards/margins only; never recolor PACKAGE_TRUTH packaging."
    )


def palette_for_ui() -> List[Dict[str, str]]:
    return [{"id": p["id"], "label": p["label"]} for p in PALETTE]


def palette_meta() -> Dict[str, Any]:
    return {
        "version": VERSION,
        "default": DEFAULT_ID,
        "ids": [p["id"] for p in PALETTE],
    }
