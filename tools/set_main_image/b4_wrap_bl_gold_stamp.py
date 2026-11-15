# -*- coding: utf-8 -*-
"""MAIN左下・個包装金丸の Pillow 後載せ。

専用素材: assets/kinmaru_wrap_bl/kinmaru_wrap_bl_ring_1200.png
（二重リング。中心とリング間は透過）

正:
- 写真円は内リングに密着（隙間なし。内リング＝写真の縁取り）
- 内リング〜外リングの透過すき間に MAIN 背景が見える
- 金丸の幾何中心と写真円の中心を一致させる
切戻し: B4_WRAP_BL_GOLD=0
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from PIL import Image, ImageDraw

SCRIPT_DIR = Path(__file__).resolve().parent
ASSET_DIR = SCRIPT_DIR / "assets" / "kinmaru_wrap_bl"
GOLD_RING = ASSET_DIR / "kinmaru_wrap_bl_ring_1200.png"

CANVAS = 1200
TR_CX, TR_CY, DIAM = 1102, 103, 343
# 写真を内リングの下にわずかに食い込ませ、アンチエイリアスの白隙間を消す（直径比）
DEFAULT_INNER_RING_OVERLAP_RATIO = 0.012


def wrap_bl_gold_enabled() -> bool:
    raw = str(os.environ.get("B4_WRAP_BL_GOLD") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def bl_gold_params(
    *,
    canvas: int = CANVAS,
    diam: int = DIAM,
) -> Tuple[float, float, int]:
    cx = float(canvas - 1) - float(TR_CX) * (canvas / CANVAS)
    cy = float(canvas - 1) - float(TR_CY) * (canvas / CANVAS)
    d = int(round(diam * (canvas / CANVAS)))
    return cx, cy, d


def _punch_circle(im: Image.Image) -> Image.Image:
    out = im.convert("RGBA")
    w, h = out.size
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, w - 1, h - 1), fill=255)
    out.putalpha(mask)
    return out


def measure_double_ring(im: Image.Image) -> Dict[str, float]:
    """
    中心から外へ見た不透明帯。
    戻り: holeR（内リング内側）, innerRingOuterR, outerRingInnerR, outerRingOuterR
    """
    import numpy as np

    a = np.array(im.convert("RGBA"))
    h, w = a.shape[:2]
    cx, cy = w / 2.0, h / 2.0
    alpha = a[:, :, 3]
    yy, xx = np.ogrid[:h, :w]
    dist = np.sqrt((xx - cx) ** 2 + (yy - cy) ** 2)
    max_r = int(min(cx, cy, w - 1 - cx, h - 1 - cy))
    means = []
    for r0 in range(0, max_r):
        m = (dist >= r0) & (dist < r0 + 1)
        means.append(float(alpha[m].mean()) if m.any() else 0.0)

    def segs(thr: float = 80.0):
        out = []
        start = None
        for r, ma in enumerate(means):
            on = ma > thr
            if on and start is None:
                start = r
            if (not on) and start is not None:
                out.append((start, r - 1))
                start = None
        if start is not None:
            out.append((start, len(means) - 1))
        return out

    opaque = segs(80.0)
    if len(opaque) < 1:
        raise ValueError("金丸リングの不透明帯が見つかりません")
    # 二重なら内側帯＝内リング、外側帯＝外リング
    if len(opaque) >= 2:
        inner = opaque[0]
        outer = opaque[-1]
    else:
        # 単帯だけなら外枠のみ扱い（穴半径＝帯の内側）
        inner = opaque[0]
        outer = opaque[0]
    return {
        "holeR": float(inner[0]),
        "innerRingOuterR": float(inner[1]),
        "outerRingInnerR": float(outer[0]),
        "outerRingOuterR": float(outer[1]),
        "imageSize": float(min(w, h)),
    }


def load_wrap_gold_ring(diameter: int) -> Tuple[Image.Image, Dict[str, float]]:
    """二重リングPNGを直径にリサイズ。幾何はスケール後の値で返す。"""
    if not GOLD_RING.is_file():
        raise FileNotFoundError(f"個包装専用リングがありません: {GOLD_RING}")
    im = Image.open(GOLD_RING).convert("RGBA")
    geo0 = measure_double_ring(im)
    # 不透明外接で切り出してから diameter へ（余白透過を落とす）
    a = im.split()[-1]
    w, h = im.size
    xs = [x for x in range(w) if any(a.getpixel((x, y)) > 20 for y in range(0, h, max(1, h // 64)))]
    ys = [y for y in range(h) if any(a.getpixel((x, y)) > 20 for x in range(0, w, max(1, w // 64)))]
    if not xs or not ys:
        raise ValueError(f"リング不透明領域なし: {GOLD_RING}")
    # 円として中央寄せ正方形の方が安定
    side = min(w, h)
    x0 = (w - side) // 2
    y0 = (h - side) // 2
    crop = im.crop((x0, y0, x0 + side, y0 + side))
    scale = float(diameter) / float(side)
    ring = crop.resize((diameter, diameter), Image.Resampling.LANCZOS)
    geo = {
        "holeR": geo0["holeR"] * scale,
        "innerRingOuterR": geo0["innerRingOuterR"] * scale,
        "outerRingInnerR": geo0["outerRingInnerR"] * scale,
        "outerRingOuterR": geo0["outerRingOuterR"] * scale,
        "imageSize": float(diameter),
        "scale": scale,
    }
    return ring, geo


def subject_square_box(src: Image.Image) -> Tuple[int, int, int, int]:
    im = src.convert("RGB")
    w, h = im.size
    px = im.load()
    corners = [px[2, 2], px[w - 3, 2], px[2, h - 3], px[w - 3, h - 3]]
    br = sum(c[0] for c in corners) / 4.0
    bg = sum(c[1] for c in corners) / 4.0
    bb = sum(c[2] for c in corners) / 4.0
    step = max(1, min(w, h) // 256)
    xs: list[int] = []
    ys: list[int] = []
    thr = 28.0
    for y in range(0, h, step):
        for x in range(0, w, step):
            r, g, b = px[x, y]
            if abs(r - br) + abs(g - bg) + abs(b - bb) >= thr:
                xs.append(x)
                ys.append(y)
    if len(xs) < 20:
        side = min(w, h)
        x0 = (w - side) // 2
        y0 = (h - side) // 2
        return x0, y0, x0 + side, y0 + side
    cx = sum(xs) / len(xs)
    cy = sum(ys) / len(ys)
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    bw, bh = x1 - x0 + 1, y1 - y0 + 1
    side = int(round(max(bw, bh) * 0.72))
    side = max(32, min(side, min(w, h)))
    half = side / 2.0
    sx0 = int(round(cx - half))
    sy0 = int(round(cy - half))
    sx1 = sx0 + side
    sy1 = sy0 + side
    if sx0 < 0:
        sx1 -= sx0
        sx0 = 0
    if sy0 < 0:
        sy1 -= sy0
        sy0 = 0
    if sx1 > w:
        sx0 -= sx1 - w
        sx1 = w
    if sy1 > h:
        sy0 -= sy1 - h
        sy1 = h
    return max(0, sx0), max(0, sy0), sx1, sy1


def circular_crop_subject(src: Image.Image, diameter: int) -> Image.Image:
    box = subject_square_box(src)
    crop = src.convert("RGBA").crop(box).resize((diameter, diameter), Image.Resampling.LANCZOS)
    return _punch_circle(crop)


def stamp_wrap_bl_gold(
    base: Image.Image,
    hand_hold: Image.Image,
    *,
    canvas: int = CANVAS,
    diam: Optional[int] = None,
    inner_ring_overlap_ratio: float = DEFAULT_INNER_RING_OVERLAP_RATIO,
) -> Tuple[Image.Image, Dict[str, Any]]:
    """
    左下に二重金リングを後載せし、内リングに密着するよう個包装円をはめる。
    写真→リングの順で、内リング下にわずかに食い込ませて白隙間を出さない。
    リング間の透過はそのまま（MAINが見える）。
    """
    if not wrap_bl_gold_enabled():
        return base.convert("RGBA"), {"skipped": True, "reason": "B4_WRAP_BL_GOLD=0"}

    out = base.convert("RGBA")
    if out.size != (canvas, canvas):
        out = out.resize((canvas, canvas), Image.Resampling.LANCZOS)

    cx, cy, d = bl_gold_params(canvas=canvas, diam=diam or DIAM)
    ring, geo = load_wrap_gold_ring(d)

    hole_r = float(geo["holeR"])
    # 内リング内側に接しつつ、リング金の下へわずかに重ねる（隙間＝不合格）
    overlap = max(1.0, d * float(inner_ring_overlap_ratio))
    # 内リング外側を超えない上限
    max_r = float(geo.get("innerRingOuterR") or hole_r) - 0.5
    content_r = min(max_r, hole_r + overlap)
    content_d = max(8, int(round(content_r * 2.0)))

    subject = circular_crop_subject(hand_hold, content_d)
    # 金丸と同じ幾何中心（内リングと一つの円）
    ix = int(round(cx - content_d / 2.0))
    iy = int(round(cy - content_d / 2.0))
    gx = int(round(cx - d / 2.0))
    gy = int(round(cy - d / 2.0))

    layer = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
    layer.alpha_composite(subject, (ix, iy))
    # 写真 → リング（内リングが縁取り。リング間透過は base）
    out.alpha_composite(layer, (0, 0))
    out.alpha_composite(ring, (gx, gy))

    meta = {
        "ok": True,
        "asset": str(GOLD_RING),
        "geomCx": cx,
        "geomCy": cy,
        "diameter": d,
        "holeR": hole_r,
        "contentD": content_d,
        "overlap": overlap,
        "geo": geo,
        "note": "double_ring; content_flush_no_gap; inter_ring_shows_base",
    }
    return out, meta


def stamp_wrap_bl_gold_to_file(
    base_path: Path,
    hand_hold_path: Path,
    out_path: Path,
    **kwargs: Any,
) -> Dict[str, Any]:
    base = Image.open(base_path)
    hh = Image.open(hand_hold_path)
    out, meta = stamp_wrap_bl_gold(base, hh, **kwargs)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out.convert("RGB").save(out_path, quality=92)
    meta["out"] = str(out_path)
    return meta


def render_hand_hold_ring_zoom_jpeg(
    hand_hold_path: Path,
    *,
    bg_rgb: Tuple[int, int, int] = (245, 242, 236),
    zoom_size: int = 440,
    quality: int = 88,
) -> bytes:
    """
    見本選びUI用: 左下金リングにはめたアップJPEG（課金なし）。
    hand_hold の mtime でキャッシュキーを切る想定は呼び出し側。
    """
    import io

    hand_hold_path = Path(hand_hold_path)
    bg = Image.new("RGB", (CANVAS, CANVAS), bg_rgb)
    stamped, _meta = stamp_wrap_bl_gold(bg, Image.open(hand_hold_path))
    z0 = CANVAS - zoom_size
    zoom = stamped.crop((0, z0, zoom_size, CANVAS)).convert("RGB")
    buf = io.BytesIO()
    zoom.save(buf, format="JPEG", quality=int(quality))
    return buf.getvalue()
