# -*- coding: utf-8 -*-
"""Amazon MAIN 合成（compose_set_main 用の旧エンジン）。"""
from __future__ import annotations

import logging
from typing import Literal, Tuple

from PIL import Image

LOG = logging.getLogger("set_main_image.amazon_compose")

Preset = Literal["hero_pyramid", "hero_row", "hero_stack", "hero_grid"]


def choose_preset(n: int) -> Preset:
    n = int(n)
    if n <= 0:
        raise ValueError("set_count must be >= 1")
    if n == 3:
        return "hero_pyramid"
    if 2 <= n <= 6:
        return "hero_row"
    if 7 <= n <= 15:
        return "hero_stack"
    return "hero_grid"


def _fit_square_rgba(im: Image.Image, box: int) -> Image.Image:
    im = im.convert("RGBA")
    im.thumbnail((box, box), Image.Resampling.LANCZOS)
    canvas = Image.new("RGBA", (box, box), (0, 0, 0, 0))
    x = (box - im.width) // 2
    y = (box - im.height) // 2
    canvas.paste(im, (x, y), im)
    return canvas


def _paste_octas(
    canvas: Image.Image,
    hero_box: Tuple[int, int, int, int],
    octas: Image.Image,
    *,
    hero_layer: Image.Image | None = None,
) -> None:
    """ヒーロー本体（不透明外接）右下に Octas を載せる。"""
    from amazon_paste import measure_body_box, place_octas_on_hero_body

    hx0, hy0, hx1, hy1 = hero_box
    if hero_layer is not None:
        body = measure_body_box(hero_layer)
        bx0, by0, bx1, by1 = [int(v) for v in body["bodyBoxInFull"]]
        body_box = {
            "x0": hx0 + bx0,
            "y0": hy0 + by0,
            "x1": hx0 + bx1,
            "y1": hy0 + by1,
            "bodyW": int(body["bodyW"]),
            "bodyH": int(body["bodyH"]),
            "fullBox": [hx0, hy0, hx1, hy1],
        }
    else:
        body_box = {
            "x0": hx0,
            "y0": hy0,
            "x1": hx1,
            "y1": hy1,
            "bodyW": max(1, hx1 - hx0),
            "bodyH": max(1, hy1 - hy0),
            "fullBox": [hx0, hy0, hx1, hy1],
        }
    place_octas_on_hero_body(
        canvas,
        octas,
        body_box=body_box,
        canvas_size=canvas.width,
        overlap_frac=0.10,
        hero_rgba=hero_layer,
        hero_x=hx0,
        hero_y=hy0,
        hero_scale=1.0,
    )


def compose_amazon(
    base: Image.Image,
    set_count: int,
    *,
    canvas_size: int | None = None,
    octas: Image.Image | None = None,
    require_octas: bool = False,
    preset: Preset | None = None,
) -> Image.Image:
    """
    set_count >= 2。ヒーロー1＋ミニ set_count-1（または見た目上 set_count 個相当）。
    小は base の完全コピー。
    """
    if set_count < 2:
        raise ValueError("Amazon compose requires set_count >= 2")
    if require_octas and octas is None:
        raise ValueError("食品のため Octas PNG が必須です")

    base = base.convert("RGBA")
    side = canvas_size or min(2000, max(base.width, base.height, 1200))
    if max(base.width, base.height) <= 1200 and canvas_size is None:
        side = 1200

    preset = preset or choose_preset(set_count)
    LOG.info("amazon compose n=%s preset=%s canvas=%s", set_count, preset, side)

    out = Image.new("RGBA", (side, side), (255, 255, 255, 255))
    margin = int(side * 0.04)

    hero_max = int(side * 0.62)
    hero = _fit_square_rgba(base, hero_max)
    hx = side - margin - hero.width
    hy = (side - hero.height) // 2
    out.alpha_composite(hero, (hx, hy))
    hero_box = (hx, hy, hx + hero.width, hy + hero.height)

    n_mini = max(1, set_count - 1)
    mini_box = int(side * (0.28 if preset != "hero_grid" else 0.18))
    mini = _fit_square_rgba(base, mini_box)

    left_w = hx - margin * 2
    left_h = side - margin * 2

    if preset == "hero_row":
        gap = max(4, int(mini.width * 0.08))
        total_w = n_mini * mini.width + (n_mini - 1) * gap
        scale = min(1.0, left_w / max(1, total_w))
        if scale < 1.0:
            nw = max(16, int(mini.width * scale))
            mini = _fit_square_rgba(base, nw)
            gap = max(2, int(mini.width * 0.08))
            total_w = n_mini * mini.width + (n_mini - 1) * gap
        x0 = margin + max(0, (left_w - total_w) // 2)
        y0 = margin + (left_h - mini.height) // 2
        for i in range(n_mini):
            out.alpha_composite(mini, (x0 + i * (mini.width + gap), y0))

    elif preset == "hero_stack":
        gap = max(2, int(mini.height * 0.06))
        total_h = n_mini * mini.height + (n_mini - 1) * gap
        scale = min(1.0, left_h / max(1, total_h), left_w / max(1, mini.width))
        if scale < 1.0:
            nw = max(16, int(mini.width * scale))
            mini = _fit_square_rgba(base, nw)
            gap = max(2, int(mini.height * 0.06))
            total_h = n_mini * mini.height + (n_mini - 1) * gap
        x0 = margin + (left_w - mini.width) // 2
        y0 = margin + max(0, (left_h - total_h) // 2)
        for i in range(n_mini):
            out.alpha_composite(mini, (x0, y0 + i * (mini.height + gap)))

    elif preset == "hero_pyramid":
        back = min(2, n_mini)
        back_box = int(side * 0.34)
        bimg = _fit_square_rgba(base, back_box)
        positions = [
            (margin, side - margin - bimg.height - int(side * 0.08)),
            (
                margin + int(bimg.width * 0.55),
                side - margin - bimg.height,
            ),
        ]
        for i in range(back):
            out.alpha_composite(bimg, positions[i % len(positions)])

    else:  # hero_grid
        cols = 3 if n_mini > 9 else 2
        gap = max(2, int(side * 0.008))
        cell = min((left_w - gap * (cols - 1)) // cols, int(side * 0.16))
        cell = max(20, cell)
        mini = _fit_square_rgba(base, cell)
        rows_needed = (n_mini + cols - 1) // cols
        total_h = rows_needed * mini.height + (rows_needed - 1) * gap
        y0 = margin + max(0, (left_h - total_h) // 2)
        for i in range(n_mini):
            r, c = divmod(i, cols)
            x = margin + c * (mini.width + gap)
            y = y0 + r * (mini.height + gap)
            if y + mini.height > side - margin:
                break
            out.alpha_composite(mini, (x, y))

    if octas is not None:
        _paste_octas(out, hero_box, octas, hero_layer=hero)

    return out.convert("RGB")
