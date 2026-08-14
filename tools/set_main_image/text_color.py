# -*- coding: utf-8 -*-
"""楽天バッジ文字色（番号／日本語名 → RGB）。見本は 94.文字色見本。"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image

LOG = logging.getLogger("set_main_image.text_color")

# フォールバック（layout_rules が読めないとき）
_FALLBACK_PALETTE: List[Dict[str, Any]] = [
    {"id": "1", "nameJa": "えんじ", "hex": "#800020", "rgb": [128, 0, 32], "sampleFile": "01_えんじ.png"},
    {"id": "2", "nameJa": "青", "hex": "#001b44", "rgb": [0, 27, 68], "sampleFile": "02_青.png"},
    {"id": "3", "nameJa": "黒", "hex": "#0d1321", "rgb": [13, 19, 33], "sampleFile": "03_黒.png"},
    {"id": "4", "nameJa": "緑", "hex": "#1b4332", "rgb": [27, 67, 50], "sampleFile": "04_緑.png"},
    {"id": "5", "nameJa": "茶", "hex": "#3d2b1f", "rgb": [61, 43, 31], "sampleFile": "05_茶.png"},
]


def _palette_from_typo(typo: Optional[dict] = None) -> List[Dict[str, Any]]:
    if typo and isinstance(typo.get("textColors"), dict):
        pal = typo["textColors"].get("palette")
        if isinstance(pal, list) and pal:
            return pal
    try:
        from rakuten_layer import load_layout_rules

        rules = load_layout_rules()
        typo2 = (rules.get("rakuten") or {}).get("badgeTypography") or {}
        pal = ((typo2.get("textColors") or {}).get("palette")) or []
        if pal:
            return pal
    except Exception:
        pass
    return list(_FALLBACK_PALETTE)


def list_text_colors(typo: Optional[dict] = None) -> List[Dict[str, Any]]:
    return _palette_from_typo(typo)


def resolve_text_color(
    key: Optional[str],
    typo: Optional[dict] = None,
) -> Dict[str, Any]:
    """
    key: '1'|'2'|… または 日本語名（えんじ/青/黒/緑/茶）または hex。
    戻り値: palette エントリ + rgba (r,g,b,255)
    """
    pal = _palette_from_typo(typo)
    default_id = "1"
    if typo and isinstance(typo.get("textColors"), dict):
        default_id = str(typo["textColors"].get("defaultId") or "1")

    raw = (key or "").strip()
    if not raw:
        raw = default_id

    # hex
    if raw.startswith("#") and len(raw) == 7:
        r = int(raw[1:3], 16)
        g = int(raw[3:5], 16)
        b = int(raw[5:7], 16)
        return {
            "id": "custom",
            "nameJa": "カスタム",
            "hex": raw.lower(),
            "rgb": [r, g, b],
            "rgba": (r, g, b, 255),
            "sampleFile": "",
        }

    low = raw.lower()
    for ent in pal:
        if str(ent.get("id")) == raw:
            rgb = list(ent["rgb"])
            return {**ent, "rgba": (rgb[0], rgb[1], rgb[2], 255)}
        if str(ent.get("nameJa")) == raw:
            rgb = list(ent["rgb"])
            return {**ent, "rgba": (rgb[0], rgb[1], rgb[2], 255)}
        if str(ent.get("hex") or "").lower() == low:
            rgb = list(ent["rgb"])
            return {**ent, "rgba": (rgb[0], rgb[1], rgb[2], 255)}

    # 英語エイリアス
    aliases = {
        "enji": "1",
        "burgundy": "1",
        "blue": "2",
        "black": "3",
        "green": "4",
        "brown": "5",
        "cha": "5",
    }
    if low in aliases:
        return resolve_text_color(aliases[low], typo)

    LOG.warning("unknown text color %r — fallback id=%s", key, default_id)
    return resolve_text_color(default_id, typo)


def recolor_glyph_rgba(
    im: Image.Image,
    rgba: Tuple[int, int, int, int],
    *,
    alpha_threshold: int = 8,
) -> Image.Image:
    """不透明画素のRGBを差し替え（アルファ維持）。縦横比・形状はそのまま。"""
    src = im.convert("RGBA")
    px = list(src.getdata())
    r, g, b, _a = rgba
    out = []
    for pr, pg, pb, pa in px:
        if pa <= alpha_threshold:
            out.append((pr, pg, pb, pa))
        else:
            # 元のアルファを維持（アンチエイリアス縁も同じ色で薄く）
            out.append((r, g, b, pa))
    dst = Image.new("RGBA", src.size)
    dst.putdata(out)
    return dst


def find_color_sample_dir(work_root: Optional[Path]) -> Optional[Path]:
    if not work_root or not work_root.is_dir():
        return None
    for p in work_root.iterdir():
        if p.is_dir() and p.name.startswith("94"):
            return p
    return None
