# -*- coding: utf-8 -*-
"""
楽天／Yahoo: ベースは変更せず、金丸内だけ3段組版で重ねる（Python本線）。

  大数字 ＋ 右に小単位（マスタ「バリエーション単位」）＋ 下に「セット」
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, ImageDraw

from glyph_assets import (
    compose_badge_with_glyphs,
    default_glyph_dirs,
    ensure_generated_glyphs,
    local_auto_glyph_dir,
)
from rakuten_badge import _load_font, _resize_font, load_font_catalog
from trace_log import file_fingerprint

LOG = logging.getLogger("set_main_image.rakuten_layer")

_RULES_PATH = Path(__file__).resolve().parent / "layout_rules.json"
CanvasXY = Tuple[int, int]


def load_layout_rules() -> Dict[str, Any]:
    if _RULES_PATH.is_file():
        import json

        return json.loads(_RULES_PATH.read_text(encoding="utf-8"))
    return {}


def _rakuten_typo() -> Dict[str, Any]:
    return (load_layout_rules().get("rakuten") or {}).get("badgeTypography") or {}


def _gold() -> Dict[str, Any]:
    return (load_layout_rules().get("rakuten") or {}).get("goldCircle1200") or {
        "cx": 1048,
        "cy": 152,
        "diameter": 256,
    }


def sample_gold_fill(im: Image.Image, cx: int, cy: int, diameter: int) -> Tuple[int, int, int, 255]:
    """金丸リング付近の平均色（文字クリア塗り用）。"""
    rgb = im.convert("RGB")
    px = rgb.load()
    r0 = max(8, int(diameter * 0.38))
    r1 = max(r0 + 2, int(diameter * 0.48))
    acc = [0, 0, 0]
    n = 0
    w, h = rgb.size
    for y in range(max(0, cy - r1), min(h, cy + r1 + 1)):
        for x in range(max(0, cx - r1), min(w, cx + r1 + 1)):
            dx, dy = x - cx, y - cy
            d2 = dx * dx + dy * dy
            if r0 * r0 <= d2 <= r1 * r1:
                r, g, b = px[x, y]
                if r + g > 280 and b < 200:
                    acc[0] += r
                    acc[1] += g
                    acc[2] += b
                    n += 1
    if n < 20:
        return (232, 198, 110, 255)
    return (acc[0] // n, acc[1] // n, acc[2] // n, 255)


def _fit_font(draw: ImageDraw.ImageDraw, text: str, font_id: str, target_h: int, max_w: int):
    size = max(10, int(target_h))
    font = _load_font(font_id, size)
    for _ in range(28):
        font = _resize_font(font, size, font_id)
        tb = draw.textbbox((0, 0), text, font=font)
        tw, th = tb[2] - tb[0], tb[3] - tb[1]
        if th <= target_h * 1.05 and tw <= max_w:
            return font, tb, tw, th
        size = max(8, int(size * 0.92))
    tb = draw.textbbox((0, 0), text, font=font)
    return font, tb, tb[2] - tb[0], tb[3] - tb[1]


def compose_rakuten_layered(
    *,
    base_path: Path,
    set_count: int,
    unit: str = "",
    font_id: str = "gothic",
    digit_layer_path: Optional[Path] = None,
    digit_box: Optional[Dict[str, int]] = None,
    style_ref_path: Optional[Path] = None,
    draw_set_label: bool = True,
    work_root: Optional[Path] = None,
    render_mode: Optional[str] = None,
    text_color: Optional[str] = None,
) -> Tuple[Image.Image, Dict[str, Any]]:
    """
    金丸デザインは触らず、透過文字（グリフPNG）またはFont描画のみ重ねる。
    ベースに数字は入っていない前提 → clearOldText 既定 false。
    text_color: 1|えんじ|青|黒|緑|茶|#rrggbb（既定=layout_rules defaultId）
    """
    from text_color import resolve_text_color

    del digit_layer_path, digit_box
    canvas = int(load_layout_rules().get("canvas") or 1200)
    base = Image.open(base_path).convert("RGBA")
    if base.size != (canvas, canvas):
        base = base.resize((canvas, canvas), Image.Resampling.LANCZOS)

    work = base.copy()
    typo = _rakuten_typo()
    gold = _gold()
    cx, cy = int(gold["cx"]), int(gold["cy"])
    diameter = int(gold["diameter"])
    radius = diameter / 2.0

    color_ent = resolve_text_color(text_color, typo)
    tint_rgba = color_ent["rgba"]

    mode = (render_mode or typo.get("renderMode") or "glyph_alpha").strip()
    # 旧仕様の楕円クリアは、ベースに文字が無い運用では使わない
    if typo.get("clearOldText", False):
        inset = float(typo.get("clearInsetRatioOfDiameter", 0.10))
        clear_r = radius * (1.0 - inset)
        fill_gold = sample_gold_fill(work, cx, cy, diameter)
        draw = ImageDraw.Draw(work)
        draw.ellipse(
            [cx - clear_r, cy - clear_r, cx + clear_r, cy + clear_r],
            fill=fill_gold,
        )
        LOG.warning("clearOldText=true: gold fill applied (hides texture)")

    glyph_meta: Dict[str, Any] = {}
    if mode == "glyph_alpha":
        dirs = default_glyph_dirs(work_root)
        # Canva 95 は上書き禁止。ローカル自動生成フォルダだけ補完。
        fill = tuple(tint_rgba[:4])
        fid = font_id or "tsukushi_mincho_like"
        ensure_generated_glyphs(local_auto_glyph_dir(), font_id=fid, fill=fill, force=False)
        work, glyph_meta = compose_badge_with_glyphs(
            work,
            set_count=int(set_count),
            unit=(unit or "").strip(),
            cx=cx,
            cy=cy,
            diameter=diameter,
            typo=typo,
            glyph_dirs=dirs,
            draw_set_label=draw_set_label,
            canvas=int(work.size[0]),
            tint_rgba=tint_rgba,
        )
        layout = glyph_meta
    else:
        layout = _draw_three_tier(
            work,
            set_count=int(set_count),
            unit=(unit or "").strip(),
            font_id=font_id,
            cx=cx,
            cy=cy,
            diameter=diameter,
            typo=typo,
            draw_set_label=draw_set_label,
            fill_rgba=tint_rgba,
        )

    badge = f"{set_count}{unit}" if unit else str(set_count)
    if draw_set_label:
        badge = f"{badge}/セット"

    meta: Dict[str, Any] = {
        "mode": layout.get("mode") if isinstance(layout, dict) else "rakuten_layer",
        "renderMode": mode,
        "base": file_fingerprint(base_path),
        "styleRef": file_fingerprint(style_ref_path) if style_ref_path and style_ref_path.is_file() else None,
        "fontId": font_id,
        "digitLen": 1 if int(set_count) < 10 else 2,
        "badgeText": badge,
        "setCount": int(set_count),
        "unit": unit,
        "textColor": {
            "id": color_ent.get("id"),
            "nameJa": color_ent.get("nameJa"),
            "hex": color_ent.get("hex"),
            "rgb": color_ent.get("rgb"),
        },
        "goldCircle": {"cx": cx, "cy": cy, "diameter": diameter},
        "typographyLayout": layout,
        "rule": "base_locked_alpha_text_overlay_only",
        "clearOldText": bool(typo.get("clearOldText", False)),
        "layoutRules": _RULES_PATH.name,
        "masterUnitHeader": (load_layout_rules().get("rakuten") or {}).get(
            "masterUnitHeader", "バリエーション単位"
        ),
    }
    LOG.info(
        "rakuten overlay base=%s mode=%s n=%s unit=%r clear=%s",
        base_path.name,
        mode,
        set_count,
        unit,
        meta["clearOldText"],
    )
    return work.convert("RGB"), meta


def _draw_three_tier(
    work: Image.Image,
    *,
    set_count: int,
    unit: str,
    font_id: str,
    cx: int,
    cy: int,
    diameter: int,
    typo: Dict[str, Any],
    draw_set_label: bool,
    fill_rgba: Optional[Tuple[int, int, int, int]] = None,
) -> Dict[str, Any]:
    draw = ImageDraw.Draw(work)
    content_inset = float(typo.get("contentInsetRatioOfDiameter", 0.14))
    content_r = (diameter / 2.0) * (1.0 - content_inset)
    content_d = content_r * 2

    digit_len = "1digit" if set_count < 10 else "2digit"
    num_cfg = (typo.get("number") or {}).get(digit_len) or {}
    unit_cfg = typo.get("unit") or {}
    set_cfg = typo.get("setLabel") or {}
    if fill_rgba:
        fill = tuple(fill_rgba[:4])
    else:
        fill = tuple((typo.get("fillColorDefault") or [90, 18, 18, 255])[:4])

    number_h = int(diameter * float(num_cfg.get("heightRatioOfDiameter", 0.48)))
    number_max_w = int(content_d * float(num_cfg.get("maxWidthRatioOfContent", 0.5)))
    num_text = str(int(set_count))
    num_font, num_tb, num_w, num_h = _fit_font(draw, num_text, font_id, number_h, number_max_w)

    unit_h = int(num_h * float(unit_cfg.get("heightRatioOfNumber", 0.40)))
    gap_nu = int(num_h * float(unit_cfg.get("gapAfterNumberRatioOfNumberHeight", 0.08)))
    unit_text = unit
    unit_w = unit_h_draw = 0
    unit_font = unit_tb = None
    if unit_text:
        unit_font, unit_tb, unit_w, unit_h_draw = _fit_font(
            draw, unit_text, font_id, unit_h, int(content_d * 0.35)
        )

    set_text = str(set_cfg.get("text") or "セット")
    set_h = int(num_h * float(set_cfg.get("heightRatioOfNumber", 0.24)))
    gap_ns = int(diameter * float(set_cfg.get("gapBelowNumberRowRatioOfDiameter", 0.045)))
    set_font = set_tb = None
    set_w = set_h_draw = 0
    if draw_set_label:
        set_font, set_tb, set_w, set_h_draw = _fit_font(
            draw, set_text, font_id, set_h, int(content_d * 0.7)
        )

    row1_w = num_w + (gap_nu + unit_w if unit_text else 0)
    row1_h = max(num_h, unit_h_draw or 0)
    total_h = row1_h + (gap_ns + set_h_draw if draw_set_label else 0)

    bias = float(typo.get("blockVerticalBias", -0.06))
    block_top = int(cy - total_h / 2 + bias * diameter)
    # 余白下限
    margins = typo.get("margins") or {}
    min_top = int(cy - content_r + diameter * float(margins.get("minTopPaddingRatioOfDiameter", 0.10)))
    min_bottom = int(cy + content_r - diameter * float(margins.get("minBottomPaddingRatioOfDiameter", 0.14)))
    if block_top < min_top:
        block_top = min_top
    if block_top + total_h > min_bottom:
        block_top = max(min_top, min_bottom - total_h)

    row1_left = int(cx - row1_w / 2)
    num_x = row1_left - num_tb[0]
    num_y = block_top + (row1_h - num_h) // 2 - num_tb[1]
    draw.text((num_x, num_y), num_text, font=num_font, fill=fill)

    unit_pos = None
    if unit_text and unit_font is not None and unit_tb is not None:
        unit_x = row1_left + num_w + gap_nu - unit_tb[0]
        unit_y = block_top + (row1_h - unit_h_draw) // 2 - unit_tb[1]
        draw.text((unit_x, unit_y), unit_text, font=unit_font, fill=fill)
        unit_pos = {"x": unit_x, "y": unit_y, "w": unit_w, "h": unit_h_draw}

    set_pos = None
    if draw_set_label and set_font is not None and set_tb is not None:
        set_x = int(cx - set_w / 2) - set_tb[0]
        set_y = block_top + row1_h + gap_ns - set_tb[1]
        draw.text((set_x, set_y), set_text, font=set_font, fill=fill)
        set_pos = {"x": set_x, "y": set_y, "w": set_w, "h": set_h_draw}

    return {
        "numberPx": {"x": num_x, "y": num_y, "w": num_w, "h": num_h, "fontPx": getattr(num_font, "size", None)},
        "unitPx": unit_pos,
        "setPx": set_pos,
        "groupBox": {
            "left": row1_left,
            "top": block_top,
            "width": max(row1_w, set_w),
            "height": total_h,
        },
        "ratios": {
            "numberHeightRatioOfDiameter": number_h / diameter,
            "unitHeightRatioOfNumber": (unit_h_draw / num_h) if num_h else None,
            "setHeightRatioOfNumber": (set_h_draw / num_h) if num_h and set_h_draw else None,
            "contentInsetRatio": content_inset,
            "gapNumberUnitPx": gap_nu,
            "gapRowToSetPx": gap_ns,
        },
    }


# 後方互換エイリアス
def digit_box_for_count(set_count: int) -> Dict[str, int]:
    gold = _gold()
    d = int(gold["diameter"])
    cx, cy = int(gold["cx"]), int(gold["cy"])
    return {"x": cx - d // 2, "y": cy - d // 2, "w": d, "h": d}


def font_scale_for_count(set_count: int) -> float:
    return 0.90 if int(set_count) < 10 else 0.78


def format_badge_text(set_count: int, unit: str) -> str:
    u = (unit or "").strip()
    return f"{set_count}{u}" if u else str(set_count)
