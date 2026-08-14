# -*- coding: utf-8 -*-
"""楽天／Yahoo: 数字なし土台＋書体カタログ（または数字レイヤ位置）で N を描画。"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from PIL import Image, ImageDraw, ImageFont

LOG = logging.getLogger("set_main_image.rakuten")

_CATALOG_PATH = Path(__file__).resolve().parent / "font_catalog.json"


def load_font_catalog() -> Dict[str, Any]:
    return json.loads(_CATALOG_PATH.read_text(encoding="utf-8"))


def list_font_ids() -> str:
    cat = load_font_catalog()
    lines = []
    for f in cat.get("fonts", []):
        lines.append(f"{f['id']}: {f.get('label', '')} ≒ {', '.join(f.get('canva_like') or [])}")
    return "\n".join(lines)


def _opaque_bbox(im: Image.Image, alpha_min: int = 16) -> Tuple[int, int, int, int]:
    im = im.convert("RGBA")
    alpha = im.split()[-1]
    bbox = alpha.point(lambda a: 255 if a >= alpha_min else 0).getbbox()
    if not bbox:
        raise ValueError("数字レイヤに不透明ピクセルがありません")
    return bbox


def _sample_fill_color(im: Image.Image, bbox: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
    crop = im.convert("RGBA").crop(bbox)
    pixels = [p for p in crop.getdata() if p[3] >= 16]
    if not pixels:
        return (120, 20, 20, 255)
    r = sum(p[0] for p in pixels) // len(pixels)
    g = sum(p[1] for p in pixels) // len(pixels)
    b = sum(p[2] for p in pixels) // len(pixels)
    return (r, g, b, 255)


def _font_entry(font_id: str) -> Dict[str, Any]:
    cat = load_font_catalog()
    fid = (font_id or cat.get("default_font_id") or "helvetica_bold").strip()
    for f in cat.get("fonts", []):
        if f.get("id") == fid:
            return f
    raise ValueError(f"未知の font_id={fid}。候補:\n{list_font_ids()}")


def _load_font(font_id: str, pixel_h: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    entry = _font_entry(font_id)
    size = max(12, int(pixel_h * 0.9))
    for path in entry.get("windows_paths") or []:
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    # 最終フォールバック
    for path in (
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/msgothic.ttc",
    ):
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _resize_font(font: ImageFont.ImageFont, size: int, font_id: str) -> ImageFont.ImageFont:
    entry = _font_entry(font_id)
    for path in entry.get("windows_paths") or []:
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return font


def compose_rakuten_badge(
    base: Image.Image,
    set_count: int,
    *,
    digit_layer: Image.Image | None = None,
    digit_box: Dict[str, int] | None = None,
    font_id: str = "helvetica_bold",
    fill: Tuple[int, int, int, int] | None = None,
    canvas_size: int = 1200,
) -> Image.Image:
    """
    base: 数字なし（金丸込み推奨）
    digit_layer: 任意。不透明領域＝数字の位置・色サンプル
    digit_box: 任意。{x,y,w,h}。レイヤも箱も無ければ catalog の default_digit_box_1200
    font_id: font_catalog.json の id
    """
    if set_count < 1:
        raise ValueError("rakuten/yahoo compose requires set_count >= 1")

    base = base.convert("RGBA")
    if base.size != (canvas_size, canvas_size):
        base = base.resize((canvas_size, canvas_size), Image.Resampling.LANCZOS)

    bbox: Tuple[int, int, int, int]
    if digit_layer is not None:
        dl = digit_layer.convert("RGBA")
        if dl.size != base.size:
            dl = dl.resize(base.size, Image.Resampling.LANCZOS)
        bbox = _opaque_bbox(dl)
        if fill is None:
            fill = _sample_fill_color(dl, bbox)
    elif digit_box:
        x, y, w, h = (
            int(digit_box["x"]),
            int(digit_box["y"]),
            int(digit_box["w"]),
            int(digit_box["h"]),
        )
        bbox = (x, y, x + w, y + h)
    else:
        cat = load_font_catalog()
        box = cat.get("default_digit_box_1200") or {"x": 955, "y": 70, "w": 140, "h": 110}
        x, y, w, h = int(box["x"]), int(box["y"]), int(box["w"]), int(box["h"])
        bbox = (x, y, x + w, y + h)

    if fill is None:
        fill = (120, 20, 20, 255)

    out = base.copy()
    draw = ImageDraw.Draw(out)
    text = str(int(set_count))
    box_h = max(12, bbox[3] - bbox[1])
    box_w = max(12, bbox[2] - bbox[0])
    font = _load_font(font_id, box_h)

    for _ in range(16):
        tb = draw.textbbox((0, 0), text, font=font)
        tw, th = tb[2] - tb[0], tb[3] - tb[1]
        if tw <= box_w * 1.08 and th <= box_h * 1.15:
            break
        cur = getattr(font, "size", None)
        if not isinstance(cur, int) or cur <= 10:
            break
        font = _resize_font(font, max(10, cur - 4), font_id)

    tb = draw.textbbox((0, 0), text, font=font)
    tw, th = tb[2] - tb[0], tb[3] - tb[1]
    cx = (bbox[0] + bbox[2]) // 2
    cy = (bbox[1] + bbox[3]) // 2
    x = cx - tw // 2 - tb[0]
    y = cy - th // 2 - tb[1]
    draw.text((x, y), text, font=font, fill=fill)
    LOG.info("rakuten badge n=%s font_id=%s bbox=%s", set_count, font_id, bbox)
    return out.convert("RGB")
