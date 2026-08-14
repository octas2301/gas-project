# -*- coding: utf-8 -*-
"""Amazon MAIN のインク占有率など（チューニング用計測）。"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Tuple

from PIL import Image


def ink_fill_ratio(
    path: Path, *, white_thresh: int = 245
) -> Dict[str, Any]:
    """
    純白に近い画素以外をインクとみなし、占有率を返す。
    white_thresh: RGB全チャネルがこれ以上なら「余白」。
    """
    im = Image.open(path).convert("RGB")
    w, h = im.size
    pix = im.load()
    ink_count = 0
    x0, y0, x1, y1 = w, h, -1, -1
    for y in range(h):
        for x in range(w):
            r, g, b = pix[x, y]
            if r < white_thresh or g < white_thresh or b < white_thresh:
                ink_count += 1
                if x < x0:
                    x0 = x
                if y < y0:
                    y0 = y
                if x > x1:
                    x1 = x
                if y > y1:
                    y1 = y
    total = h * w
    ratio = ink_count / total if total else 0.0
    if x1 < 0:
        margins = {"top": 1.0, "bottom": 1.0, "left": 1.0, "right": 1.0}
    else:
        margins = {
            "top": y0 / h,
            "bottom": (h - 1 - y1) / h,
            "left": x0 / w,
            "right": (w - 1 - x1) / w,
        }

    return {
        "path": str(path),
        "width": w,
        "height": h,
        "inkFillRatio": round(ratio, 4),
        "inkPixels": ink_count,
        "whiteThresh": white_thresh,
        "margins": {k: round(v, 4) for k, v in margins.items()},
        "maxMargin": round(max(margins.values()), 4),
    }


def fill_target_for_n(set_count: int) -> Tuple[float, str]:
    """layout_rules の draft 閾値。未設定時は帯別デフォルト。"""
    try:
        import json

        rules = json.loads(
            (Path(__file__).resolve().parent / "layout_rules.json").read_text(
                encoding="utf-8"
            )
        )
        tun = (rules.get("amazon") or {}).get("tuningDraft") or {}
        fill = tun.get("inkFillMinByN") or {}
        n = int(set_count)
        if n <= 1:
            return float(fill.get("n1", 0.30)), "n1"
        if n <= 3:
            return float(fill.get("n2_3", 0.48)), "n2_3"
        if n <= 6:
            return float(fill.get("n4_6", 0.52)), "n4_6"
        return float(fill.get("n8_plus", 0.58)), "n8_plus"
    except Exception:
        n = int(set_count)
        if n <= 1:
            return 0.30, "n1"
        if n <= 3:
            return 0.48, "n2_3"
        if n <= 6:
            return 0.52, "n4_6"
        return 0.58, "n8_plus"
