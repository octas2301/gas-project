# -*- coding: utf-8 -*-
"""Octasシール前処理: 商品貼付に見せない／左右に軽い回転。"""
from __future__ import annotations

import hashlib
import logging
import random
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from PIL import Image

LOG = logging.getLogger("set_main_image.octas")

# draft → 2026-08-07 合格固定: |θ|=8°（上部を右へ）
DEFAULT_TILT_ABS_MIN = 8.0
DEFAULT_TILT_ABS_MAX = 8.0
DEFAULT_TILT_ABS = 8.0
# 表示サイズ: 商品の不透明ピクセル数 × この比率 ＝ シールの不透明ピクセル数
# 2026-08-07: 標準0.05の1.5倍でロック
OCTAS_PRODUCT_AREA_RATIO = 0.075
# 极端サイズのガード（キャンバス幅比）
OCTAS_WIDTH_MIN_RATIO = 0.03
OCTAS_WIDTH_MAX_RATIO = 0.35
OPAQUE_ALPHA_THRESH = 20

# 後方互換名（線形スケールの目安 √0.05）
OCTAS_HERO_WIDTH_RATIO = OCTAS_PRODUCT_AREA_RATIO ** 0.5


def count_opaque_pixels(im: Image.Image, *, threshold: int = OPAQUE_ALPHA_THRESH) -> int:
    """RGBA の不透明ピクセル数（alpha > threshold）。"""
    if im.mode != "RGBA":
        im = im.convert("RGBA")
    alpha = im.split()[-1]
    thr = int(threshold)
    return sum(1 for p in alpha.getdata() if p > thr)


def octas_display_size_for_opaque_ratio(
    *,
    canvas_size: int,
    product_opaque_px: int,
    seal_rgba: Image.Image,
    area_ratio: float = OCTAS_PRODUCT_AREA_RATIO,
) -> Tuple[int, int, Dict[str, Any]]:
    """
    シール不透明ピクセル ≈ 商品不透明ピクセル × area_ratio になる (w,h)。
    縦横比は seal を維持。リサイズ後にバイナリ探索で合わせる。
    """
    canvas_size = max(1, int(canvas_size))
    target = max(1.0, float(product_opaque_px) * float(area_ratio))
    src = seal_rgba.convert("RGBA") if seal_rgba.mode != "RGBA" else seal_rgba
    src_w, src_h = max(1, src.width), max(1, src.height)
    src_opaque = max(1, count_opaque_pixels(src))

    s0 = (target / float(src_opaque)) ** 0.5
    tw0 = int(round(src_w * s0))
    tw_min = max(24, int(round(canvas_size * OCTAS_WIDTH_MIN_RATIO)))
    tw_max = max(tw_min, int(round(canvas_size * OCTAS_WIDTH_MAX_RATIO)))
    tw0 = max(tw_min, min(tw0, tw_max))

    def _size_for_w(w: int) -> Tuple[int, int, int]:
        w = max(1, int(w))
        h = max(1, int(round(src_h * (w / src_w))))
        scaled = src.resize((w, h), Image.Resampling.LANCZOS)
        return w, h, count_opaque_pixels(scaled)

    best_w, best_h, best_op = _size_for_w(tw0)
    best_diff = abs(best_op - target)
    lo = tw_min
    hi = tw_max
    for _ in range(18):
        mid = (lo + hi) // 2
        w, h, op = _size_for_w(mid)
        diff = abs(op - target)
        if diff < best_diff:
            best_diff = diff
            best_w, best_h, best_op = w, h, op
        if op < target:
            lo = mid + 1
        else:
            hi = mid - 1
        if lo > hi:
            break

    meta: Dict[str, Any] = {
        "productOpaquePx": int(product_opaque_px),
        "sealOpaquePx": int(best_op),
        "targetOpaquePx": int(round(target)),
        "sealOpaqueRatioActual": round(best_op / float(max(1, product_opaque_px)), 4),
        "sizeMode": "opaque_pixel_ratio",
        "areaRatio": float(area_ratio),
    }
    return best_w, best_h, meta


def octas_display_size_for_hero(
    *,
    canvas_size: int,
    hero_w: int,
    src_w: int,
    src_h: int,
    hero_ratio: float | None = None,
    area_ratio: float = OCTAS_PRODUCT_AREA_RATIO,
    hero_h: int | None = None,
) -> Tuple[int, int]:
    """後方互換: 外接矩形面積比。本線は octas_display_size_for_opaque_ratio。"""
    canvas_size = max(1, int(canvas_size))
    hero_w = max(1, int(hero_w))
    src_w = max(1, int(src_w))
    src_h = max(1, int(src_h))
    aspect = src_h / src_w

    if hero_ratio is not None:
        tw = int(round(hero_w * float(hero_ratio)))
    else:
        hh = max(1, int(hero_h if hero_h is not None else hero_w))
        target_area = float(hero_w) * float(hh) * float(area_ratio)
        tw = int(round((target_area / aspect) ** 0.5))

    tw_min = max(24, int(round(canvas_size * OCTAS_WIDTH_MIN_RATIO)))
    tw_max = max(tw_min, int(round(canvas_size * OCTAS_WIDTH_MAX_RATIO)))
    tw = max(tw_min, min(tw, tw_max))
    th = max(1, int(round(src_h * (tw / src_w))))
    return tw, th


# 後方互換エイリアス
OCTAS_FIXED_WIDTH_RATIO = OCTAS_HERO_WIDTH_RATIO


def octas_fixed_display_size(
    canvas_size: int, src_w: int, src_h: int
) -> Tuple[int, int]:
    """非推奨: 外接概算。"""
    approx = int(round(canvas_size * 0.80))
    return octas_display_size_for_hero(
        canvas_size=canvas_size,
        hero_w=approx,
        hero_h=approx,
        src_w=src_w,
        src_h=src_h,
        area_ratio=OCTAS_PRODUCT_AREA_RATIO,
    )


def choose_octas_tilt_deg(
    *,
    tilt_abs: Optional[float] = None,
    seed: Optional[int] = None,
    direction: str = "top_to_right",
) -> float:
    """
    符号付き角度（Pillow: 正=反時計、負=時計）。

    direction:
      - top_to_right: 上部を右へ（時計回り・負）…本線
      - top_to_left: 上部を左へ（反時計・正）
      - left_or_right: 左右ランダム（旧）
    """
    rng = random.Random(seed)
    mag = float(tilt_abs) if tilt_abs is not None else DEFAULT_TILT_ABS
    mag = max(DEFAULT_TILT_ABS_MIN, min(DEFAULT_TILT_ABS_MAX, mag))
    d = (direction or "top_to_right").strip().lower()
    if d in ("top_to_left", "ccw", "left"):
        return mag
    if d in ("left_or_right", "either", "random"):
        sign = -1.0 if rng.random() < 0.5 else 1.0
        return sign * mag
    # default: 上部を右側へ
    return -mag


def tilt_octas_rgba(
    octas_rgba: Image.Image,
    *,
    tilt_deg: Optional[float] = None,
    seed: Optional[int] = None,
    direction: str = "top_to_right",
) -> Tuple[Image.Image, Dict[str, Any]]:
    """
    メモリ上のシールに本線傾きを適用（ファイルキャッシュなし）。
    compose 直呼びで未傾きシールが来たときの再発防止用。
    """
    im = octas_rgba.convert("RGBA")
    if tilt_deg is not None and abs(float(tilt_deg)) > 1e-9:
        deg = float(tilt_deg)
    else:
        tilt_abs = DEFAULT_TILT_ABS
        tilt_dir = direction
        try:
            from rakuten_layer import load_layout_rules

            oc = (
                (load_layout_rules().get("amazon") or {})
                .get("tuningDraft")
                or {}
            ).get("octasSeal") or {}
            if oc.get("tiltDegAbsDefault") is not None:
                tilt_abs = float(oc["tiltDegAbsDefault"])
            if oc.get("tiltDirection"):
                tilt_dir = str(oc["tiltDirection"])
        except Exception:
            pass
        deg = choose_octas_tilt_deg(
            tilt_abs=tilt_abs, seed=seed, direction=tilt_dir
        )
    # 白背景を透過（未処理素材向け）
    px = im.load()
    w, h = im.size
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if a > 0 and r >= 245 and g >= 245 and b >= 245:
                px[x, y] = (255, 255, 255, 0)
    bbox = im.getbbox()
    if bbox:
        im = im.crop(bbox)
    rotated = im.rotate(deg, expand=True, resample=Image.Resampling.BICUBIC)
    meta = {
        "tiltDeg": round(deg, 2),
        "tiltAbs": round(abs(deg), 2),
        "direction": "ccw" if deg > 0 else "cw",
        "directionPolicy": direction,
        "source": "memory_tilt_octas_rgba",
        "noteJa": "compose直呼び用。上部を右へ（時計回り・負角）。",
    }
    return rotated, meta


def prepare_octas_seal(
    src: Path,
    cache_dir: Path,
    *,
    tilt_deg: Optional[float] = None,
    seed: Optional[int] = None,
    direction: str = "top_to_right",
) -> Tuple[Path, Dict[str, Any]]:
    """
    シール画像を透過化し、回転してキャッシュ保存。
    既定は上部を右へ少し倒す（時計回り）。
    """
    src = Path(src)
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    if tilt_deg is not None and tilt_deg != 0:
        deg = float(tilt_deg)
    else:
        tilt_abs = DEFAULT_TILT_ABS
        tilt_dir = direction
        try:
            from rakuten_layer import load_layout_rules

            oc = (
                (load_layout_rules().get("amazon") or {})
                .get("tuningDraft")
                or {}
            ).get("octasSeal") or {}
            if oc.get("tiltDegAbsDefault") is not None:
                tilt_abs = float(oc["tiltDegAbsDefault"])
            if oc.get("tiltDirection"):
                tilt_dir = str(oc["tiltDirection"])
        except Exception:
            pass
        deg = choose_octas_tilt_deg(
            tilt_abs=tilt_abs, seed=seed, direction=tilt_dir
        )

    key = hashlib.sha1(f"{src.resolve()}|{deg:.2f}".encode("utf-8")).hexdigest()[:12]
    out = cache_dir / f"{src.stem}_tilt{deg:+.1f}_{key}.png"
    meta: Dict[str, Any] = {
        "source": str(src),
        "tiltDeg": round(deg, 2),
        "tiltAbs": round(abs(deg), 2),
        "direction": "ccw" if deg > 0 else "cw",
        "directionPolicy": direction,
        "noteJa": "上部を右へ（時計回り・負角）が本線。ヒーロー右下に軽く掛ける配置は貼付側。",
        "output": str(out),
        "policy": {
            "absMin": DEFAULT_TILT_ABS_MIN,
            "absMax": DEFAULT_TILT_ABS_MAX,
            "absDefault": DEFAULT_TILT_ABS,
            "defaultDirection": "top_to_right",
        },
    }
    if out.is_file() and out.stat().st_size > 500:
        LOG.info("octas cache hit tilt=%+.1f %s", deg, out.name)
        return out, meta

    im = Image.open(src).convert("RGBA")
    px = im.load()
    w, h = im.size
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if r >= 245 and g >= 245 and b >= 245:
                px[x, y] = (255, 255, 255, 0)
    bbox = im.getbbox()
    if bbox:
        im = im.crop(bbox)
    rotated = im.rotate(deg, expand=True, resample=Image.Resampling.BICUBIC)
    rotated.save(out, format="PNG")
    LOG.info("octas prepared tilt=%+.1f -> %s", deg, out.name)
    return out, meta
