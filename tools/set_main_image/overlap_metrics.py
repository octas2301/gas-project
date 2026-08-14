# -*- coding: utf-8 -*-
"""貼付後の重なり率。V3は『単体がヒーローに隠れる率』を主指標にする。"""
from __future__ import annotations

from typing import Any, Dict, List, Tuple

from PIL import Image, ImageChops


def _alpha_mask(im: Image.Image, thresh: int = 12) -> Image.Image:
    a = im.convert("RGBA").getchannel("A")
    return a.point(lambda p: 255 if p > thresh else 0)


def _ink_count(m: Image.Image) -> int:
    h = m.histogram()
    return int(sum(h[1:]))


def _place_mask(
    layer: Image.Image, xy: Tuple[int, int], canvas_size: int
) -> Image.Image:
    canvas = Image.new("L", (canvas_size, canvas_size), 0)
    canvas.paste(_alpha_mask(layer), xy)
    return canvas


def measure_plans_overlap(
    *,
    hero: Image.Image,
    unit: Image.Image,
    plans: List[Any],
    canvas_size: int = 1200,
) -> Dict[str, Any]:
    """
    pairMaxBackCovered: 各単体が『ヒーロー』に隠れる最大割合
    pairMaxAnyCovered: 任意レイヤが手前レイヤに隠れる最大割合（単体同士含む）
    heroMinVisible: ヒーローの可視割合（手前に何かある場合）
    """
    from amazon_paste import _scale_rgba
    from portrait_layout import rotate_rgba_cw

    placed = []
    for p in sorted(plans, key=lambda x: x.z):
        src = hero if p.role == "hero" else unit
        im = _scale_rgba(src, p.scale)
        tilt = float(getattr(p, "rotation_deg", 0) or 0)
        if abs(tilt) > 1e-9:
            im = rotate_rgba_cw(im, tilt)
        mask = _place_mask(im, (p.x, p.y), canvas_size)
        placed.append({"plan": p, "mask": mask, "pixels": max(1, _ink_count(mask))})

    hero_masks = [x for x in placed if x["plan"].role == "hero"]
    pair_rows = []
    max_unit_by_hero = 0.0
    for back in placed:
        if back["plan"].role != "unit":
            continue
        hero_union = Image.new("L", (canvas_size, canvas_size), 0)
        for h in hero_masks:
            if h["plan"].z <= back["plan"].z:
                continue
            hero_union = ImageChops.lighter(hero_union, h["mask"])
        inter = ImageChops.multiply(back["mask"], hero_union)
        covered = _ink_count(inter) / back["pixels"]
        max_unit_by_hero = max(max_unit_by_hero, covered)
        pair_rows.append(
            {
                "unitZ": back["plan"].z,
                "coveredByHero": round(covered, 4),
            }
        )

    max_any = 0.0
    unit_front_rows: List[Dict[str, Any]] = []
    for back in placed:
        front_union = Image.new("L", (canvas_size, canvas_size), 0)
        has_front = False
        for fr in placed:
            if fr["plan"].z <= back["plan"].z:
                continue
            has_front = True
            front_union = ImageChops.lighter(front_union, fr["mask"])
        if not has_front:
            if back["plan"].role == "unit":
                unit_front_rows.append(
                    {
                        "z": int(back["plan"].z),
                        "coveredByFront": 0.0,
                        "visible": 1.0,
                    }
                )
            continue
        inter = ImageChops.multiply(back["mask"], front_union)
        covered = _ink_count(inter) / back["pixels"]
        max_any = max(max_any, covered)
        if back["plan"].role == "unit":
            unit_front_rows.append(
                {
                    "z": int(back["plan"].z),
                    "coveredByFront": round(covered, 4),
                    "visible": round(max(0.0, 1.0 - covered), 4),
                }
            )
    unit_front_rows.sort(key=lambda r: r["z"])

    hero_vis = 1.0
    for h in hero_masks:
        front_union = Image.new("L", (canvas_size, canvas_size), 0)
        has_front = False
        for fr in placed:
            if fr["plan"].z <= h["plan"].z:
                continue
            has_front = True
            front_union = ImageChops.lighter(front_union, fr["mask"])
        if has_front:
            inter = ImageChops.multiply(h["mask"], front_union)
            hero_vis = min(hero_vis, 1.0 - _ink_count(inter) / h["pixels"])

    return {
        "pairMaxBackCovered": round(max_unit_by_hero, 4),
        "pairMaxAnyCovered": round(max_any, 4),
        "unitMinVisible": round(max(0.0, 1.0 - max_unit_by_hero), 4),
        "heroMinVisible": round(hero_vis, 4),
        "pairs": pair_rows,
        "unitFrontVisibility": unit_front_rows,
        "noteJa": "pairMaxBackCovered=単体←ヒーロー／unitMinVisible=1-covered／unitFrontVisibility=手前全体に対する可視",
    }
