# -*- coding: utf-8 -*-
"""
Amazon MAIN — Pillow 等倍貼付（AI描き直しなし）

- 縦横比ロック（uniform scale のみ）
- ヒーロー／単体の使い分け
- N≤4: 全個体同一スケール＋強め重なり
- Octas: 浮遊合成（商品に貼らない）
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, ImageChops, ImageFilter

LOG = logging.getLogger("set_main_image.amazon_paste")


@dataclass
class PastePlan:
    """1個体の貼付計画（キャンバス座標）。"""

    role: str  # hero | unit
    x: int
    y: int
    scale: float
    z: int  # 大きいほど手前
    rotation_deg: float = 0.0  # 時計回り度（縦長パターン）
    anchor: str = "topleft"  # topleft | foot | top
    foot_x: Optional[int] = None
    foot_y: Optional[int] = None
    top_x: Optional[int] = None
    top_y: Optional[int] = None


def _trim_rgba(im: Image.Image) -> Image.Image:
    im = im.convert("RGBA")
    bbox = im.getbbox()
    if bbox:
        im = im.crop(bbox)
    return im


def _scale_rgba(im: Image.Image, scale: float) -> Image.Image:
    """等倍スケール（縦横比維持）。"""
    if scale <= 0:
        raise ValueError("scale must be > 0")
    w = max(1, int(round(im.width * scale)))
    h = max(1, int(round(im.height * scale)))
    return im.resize((w, h), Image.Resampling.LANCZOS)


def _soft_shadow(
    size: Tuple[int, int], *, blur: int = 18, opacity: int = 70
) -> Image.Image:
    """旧: 矩形影（斜めAABBだと白っぽい板に見えるので原則使わない）。"""
    w, h = size
    sh = Image.new("RGBA", (w + blur * 4, h + blur * 4), (0, 0, 0, 0))
    blob = Image.new("RGBA", (w, h), (0, 0, 0, opacity))
    sh.paste(blob, (blur * 2, blur * 2), blob)
    return sh.filter(ImageFilter.GaussianBlur(radius=blur))


def _soft_shadow_from_alpha(
    layer: Image.Image, *, blur: int = 14, opacity: int = 55
) -> Image.Image:
    """
    レイヤのアルファ輪郭だけで影を作る（回転後の外接矩形を塗らない）。
    斜め配置で「白い板状の背景」に見えないようにする。
    """
    layer = layer.convert("RGBA")
    alpha = layer.getchannel("A")
    # 不透明部だけ薄い黒
    sh_a = alpha.point(lambda p: int(p * (opacity / 255.0)) if p > 10 else 0)
    sh = Image.new("RGBA", layer.size, (0, 0, 0, 0))
    sh.putalpha(sh_a)
    pad = max(8, int(blur) * 2)
    canvas = Image.new(
        "RGBA", (layer.width + pad * 2, layer.height + pad * 2), (0, 0, 0, 0)
    )
    canvas.paste(sh, (pad, pad), sh)
    return canvas.filter(ImageFilter.GaussianBlur(radius=blur))


def measure_body_box(
    im: Image.Image, *, top_skip_ratio: float = 0.18
) -> Dict[str, Any]:
    """
    商品『本体』の縦横サイズ。

    スプーン等の上部突起を除くため、上端 top_skip_ratio を無視した
    不透明領域の外接矩形を本体とみなす。
    返り値: bodyW/bodyH（判定用）と fullW/fullH（貼付レイアウト用）。
    """
    im = im.convert("RGBA")
    full_w, full_h = im.size
    alpha = im.getchannel("A")
    # 上をスキップした帯で bbox
    y0 = min(full_h - 1, max(0, int(full_h * top_skip_ratio)))
    band = alpha.crop((0, y0, full_w, full_h))
    bb = band.getbbox()
    if not bb:
        bb = alpha.getbbox() or (0, 0, full_w, full_h)
        body = {
            "bodyW": bb[2] - bb[0],
            "bodyH": bb[3] - bb[1],
            "bodyBoxInFull": list(bb),
            "topSkipRatio": 0.0,
        }
    else:
        # band 座標 → full 座標
        x0, by0, x1, by1 = bb
        body = {
            "bodyW": x1 - x0,
            "bodyH": by1 - by0,
            "bodyBoxInFull": [x0, y0 + by0, x1, y0 + by1],
            "topSkipRatio": top_skip_ratio,
        }
    body["fullW"] = full_w
    body["fullH"] = full_h
    return body


def hero_body_box_on_canvas(
    hero_rgba: Image.Image,
    *,
    x: int,
    y: int,
    scale: float,
    top_skip_ratio: float = 0.18,
) -> Dict[str, Any]:
    """
    貼付後キャンバス座標でのヒーロー『本体』外接（上部スキップ込み）。
    Octas 配置は full 外接ではなくこちらを正とする。
    """
    layer = _scale_rgba(_trim_rgba(hero_rgba), scale)
    body = measure_body_box(layer, top_skip_ratio=top_skip_ratio)
    bx0, by0, bx1, by1 = [int(v) for v in body["bodyBoxInFull"]]
    return {
        "x0": x + bx0,
        "y0": y + by0,
        "x1": x + bx1,
        "y1": y + by1,
        "bodyW": int(body["bodyW"]),
        "bodyH": int(body["bodyH"]),
        "fullW": layer.width,
        "fullH": layer.height,
        "fullBox": [x, y, x + layer.width, y + layer.height],
        "topSkipRatio": body.get("topSkipRatio"),
    }


def _hero_opaque_right_edge_(
    hero_rgba: Image.Image,
    *,
    x: int,
    y: int,
    scale: float,
    canvas_size: int,
    y0: int,
    y1: int,
) -> int:
    """指定バンド内でのヒーロー不透明の最大 x（キャンバス座標）。"""
    layer = _scale_rgba(_trim_rgba(hero_rgba), scale)
    alpha = layer.split()[-1]
    a = alpha.load()
    lw, lh = layer.size
    edge = x
    yy0 = max(0, y0 - y)
    yy1 = min(lh, y1 - y)
    for yy in range(yy0, max(yy0 + 1, yy1)):
        row_max = -1
        for xx in range(lw):
            if a[xx, yy] > 20:
                row_max = xx
        if row_max >= 0:
            edge = max(edge, x + row_max)
    return min(canvas_size, max(0, edge))


def _hero_alpha_mask_(
    hero_rgba: Image.Image,
    *,
    x: int,
    y: int,
    scale: float,
    canvas_size: int,
) -> Image.Image:
    layer = _scale_rgba(_trim_rgba(hero_rgba), scale)
    mask = Image.new("L", (canvas_size, canvas_size), 0)
    mask.paste(layer.split()[-1], (x, y))
    return mask


def _seal_mask_overlap_frac_(
    seal_alpha: Image.Image,
    other_mask: Image.Image,
    *,
    ox: int,
    oy: int,
    canvas_size: int,
) -> float:
    """シール不透明に対する other_mask 被覆率（ImageChops）。"""
    placed = Image.new("L", (canvas_size, canvas_size), 0)
    placed.paste(seal_alpha, (ox, oy))
    # しきい値で二値化
    seal_bin = placed.point(lambda p: 255 if p >= 30 else 0)
    other_bin = other_mask.point(lambda p: 255 if p > 20 else 0)
    ink = sum(seal_bin.histogram()[1:])
    if ink <= 0:
        return 0.0
    inter = ImageChops.multiply(seal_bin, other_bin)
    on = sum(inter.histogram()[1:])
    return on / float(ink)


def _seal_hero_overlap_frac_(
    seal_alpha: Image.Image,
    hero_mask: Image.Image,
    *,
    ox: int,
    oy: int,
    canvas_size: int,
) -> float:
    return _seal_mask_overlap_frac_(
        seal_alpha, hero_mask, ox=ox, oy=oy, canvas_size=canvas_size
    )


def _seal_hero_cover_of_hero_frac_(
    seal_alpha: Image.Image,
    hero_mask: Image.Image,
    *,
    ox: int,
    oy: int,
    canvas_size: int,
    hero_ink: Optional[int] = None,
) -> float:
    """シール∩hero / hero不透明。ヒーローを隠す面積の比率。"""
    placed = Image.new("L", (canvas_size, canvas_size), 0)
    placed.paste(seal_alpha, (ox, oy))
    seal_bin = placed.point(lambda p: 255 if p >= 30 else 0)
    other_bin = hero_mask.point(lambda p: 255 if p > 20 else 0)
    if hero_ink is None:
        ink_h = sum(other_bin.histogram()[1:])
    else:
        ink_h = int(hero_ink)
    if ink_h <= 0:
        return 0.0
    inter = ImageChops.multiply(seal_bin, other_bin)
    on = sum(inter.histogram()[1:])
    return on / float(ink_h)


def place_octas_on_hero_body(
    canvas: Image.Image,
    octas_rgba: Image.Image,
    *,
    body_box: Dict[str, Any],
    canvas_size: int,
    overlap_frac: float = 0.10,
    area_ratio: Optional[float] = None,
    hero_rgba: Optional[Image.Image] = None,
    hero_x: int = 0,
    hero_y: int = 0,
    hero_scale: float = 1.0,
    avoid_mask: Optional[Image.Image] = None,
    max_avoid_cover: float = 0.005,
    min_hero_overlap: float = 0.05,
    max_hero_cover: Optional[float] = None,
    area_ratio_fallbacks: Optional[List[float]] = None,
    pin_frame_bottom: bool = False,
) -> Dict[str, Any]:
    """
    メイン商品（レイアウト上の hero）右下に Octas をキス配置。

    - サイズ: 商品不透明ピクセル × area_ratio ＝ シール不透明ピクセル
    - 位置: シルエット右下。シール不透明∩商品 ≈ overlap_frac（既定10%）
    - avoid_mask 指定時（N≥5）: unit 非被覆を優先し、左へスライド／面積比フォールバック
    - pin_frame_bottom: シール不透明下端をキャンバス外枠下辺に接触（N≥5）
    - max_hero_cover: シールが隠す hero 不透明比率の上限。超過時は面積比を下げて再配置
    """
    from octas_prep import (
        OCTAS_PRODUCT_AREA_RATIO,
        count_opaque_pixels,
        octas_display_size_for_opaque_ratio,
        octas_display_size_for_hero,
    )

    area_r0 = float(
        OCTAS_PRODUCT_AREA_RATIO if area_ratio is None else area_ratio
    )
    area_r0 = max(0.02, min(0.12, area_r0))

    oc0 = _trim_rgba(octas_rgba)
    bx0 = max(0, int(body_box["x0"]))
    by0 = max(0, int(body_box["y0"]))
    bx1 = min(canvas_size, int(body_box["x1"]))
    by1 = min(canvas_size, int(body_box["y1"]))
    body_w = max(1, int(body_box.get("bodyW") or (bx1 - bx0)))
    body_h = max(1, int(body_box.get("bodyH") or (by1 - by0)))

    full = body_box.get("fullBox") or [bx0, by0, bx1, by1]
    fx1 = min(canvas_size, int(full[2]))
    fy1 = min(canvas_size, int(full[3]))

    hero_mask = None
    if hero_rgba is not None and hero_scale > 0:
        hero_mask = _hero_alpha_mask_(
            hero_rgba,
            x=hero_x,
            y=hero_y,
            scale=hero_scale,
            canvas_size=canvas_size,
        )

    ratio_list: List[float] = []
    for r in [area_r0] + list(area_ratio_fallbacks or []):
        rr = max(0.02, min(0.12, float(r)))
        if rr not in ratio_list:
            ratio_list.append(rr)
    if avoid_mask is not None:
        for r in (0.055, 0.04):
            if r not in ratio_list:
                ratio_list.append(r)

    def _size_for(ar: float) -> Tuple[Image.Image, Dict[str, Any]]:
        size_meta_local: Dict[str, Any] = {}
        if hero_rgba is not None and hero_scale > 0:
            layer = _scale_rgba(_trim_rgba(hero_rgba), hero_scale)
            product_opaque = count_opaque_pixels(layer)
            tw, th, size_meta_local = octas_display_size_for_opaque_ratio(
                canvas_size=canvas_size,
                product_opaque_px=product_opaque,
                seal_rgba=oc0,
                area_ratio=ar,
            )
        else:
            tw, th = octas_display_size_for_hero(
                canvas_size=canvas_size,
                hero_w=body_w,
                hero_h=body_h,
                src_w=oc0.width,
                src_h=oc0.height,
                area_ratio=ar,
            )
            size_meta_local = {
                "sizeMode": "bbox_fallback",
                "areaRatio": ar,
                "noteJa": "hero_rgba無しのため外接面積比フォールバック",
            }
        return oc0.resize((tw, th), Image.Resampling.LANCZOS), size_meta_local

    def _kiss_place(oc: Image.Image, oy_hint: Optional[int] = None) -> Tuple[int, int, float, int, int]:
        alpha = oc.split()[-1]
        sb = alpha.getbbox() or (0, 0, oc.width, oc.height)
        sx0, sy0, sx1, sy1 = (int(v) for v in sb)
        sw = max(1, sx1 - sx0)
        sh = max(1, sy1 - sy0)
        kiss0 = max(3, int(min(sw, sh) * overlap_frac))
        kiss_cy = by0 + int(body_h * 0.86)
        oy = (
            int(oy_hint)
            if oy_hint is not None
            else max(0, min(kiss_cy - int(sh * 0.45) - sy0, canvas_size - oc.height))
        )
        oy = max(0, min(oy, canvas_size - oc.height))

        if hero_mask is not None and hero_rgba is not None:
            mid = oy + sy0 + sh // 2
            band_y0 = max(0, mid - max(8, sh // 6))
            band_y1 = min(canvas_size, mid + max(8, sh // 6))
            edge_x = _hero_opaque_right_edge_(
                hero_rgba,
                x=hero_x,
                y=hero_y,
                scale=hero_scale,
                canvas_size=canvas_size,
                y0=band_y0,
                y1=band_y1,
            )
            lo = max(0, edge_x - oc.width)
            hi = min(canvas_size - oc.width, edge_x + max(4, oc.width // 5))
            best_ox = max(0, min(edge_x - kiss0 - sx0, canvas_size - oc.width))
            best_diff = 1.0
            measured = 0.0
            for _ in range(16):
                mid_x = (lo + hi) // 2
                frac = _seal_hero_overlap_frac_(
                    alpha, hero_mask, ox=mid_x, oy=oy, canvas_size=canvas_size
                )
                diff = abs(frac - overlap_frac)
                if diff < best_diff:
                    best_diff = diff
                    best_ox = mid_x
                    measured = frac
                if frac > overlap_frac:
                    lo = mid_x + 1
                else:
                    hi = mid_x - 1
                if lo > hi:
                    break
            if measured < 0.05:
                for step in range(1, oc.width):
                    cand = max(0, best_ox - step)
                    frac = _seal_hero_overlap_frac_(
                        alpha, hero_mask, ox=cand, oy=oy, canvas_size=canvas_size
                    )
                    if frac >= min(0.08, overlap_frac):
                        best_ox = cand
                        measured = frac
                        break
            return best_ox, oy, measured, edge_x, kiss0

        edge_x = fx1
        ox = max(0, min(edge_x - kiss0 - sx0, canvas_size - oc.width))
        return ox, oy, float(overlap_frac), edge_x, kiss0

    chosen_oc: Optional[Image.Image] = None
    size_meta: Dict[str, Any] = {}
    ox = oy = 0
    measured = 0.0
    edge_x = fx1
    kiss0 = 3
    area_r = area_r0
    unit_cover = 0.0
    hero_cover = 0.0
    avoid_mode = avoid_mask is not None
    placed_ok = False
    # 下枠接触時はシールの大半が hero 上に乗る → seal基準キス率ではなく
    # 「hero を隠す面積比」で一回り縮小を判定する
    max_hc: Optional[float]
    if max_hero_cover is not None:
        max_hc = max(0.0, float(max_hero_cover))
    elif avoid_mode and pin_frame_bottom:
        max_hc = 0.05
    else:
        max_hc = None
    hero_ink_total: Optional[int] = None
    if hero_mask is not None and max_hc is not None:
        hero_bin0 = hero_mask.point(lambda p: 255 if p > 20 else 0)
        hero_ink_total = int(sum(hero_bin0.histogram()[1:]))

    if not avoid_mode:
        chosen_oc, size_meta = _size_for(area_r0)
        ox, oy, measured, edge_x, kiss0 = _kiss_place(chosen_oc)
        area_r = area_r0
        placed_ok = True
        if hero_mask is not None:
            hero_cover = _seal_hero_cover_of_hero_frac_(
                chosen_oc.split()[-1],
                hero_mask,
                ox=ox,
                oy=oy,
                canvas_size=canvas_size,
                hero_ink=hero_ink_total,
            )
    else:
        max_uc = max(0.0, float(max_avoid_cover))
        min_ho = max(0.0, float(min_hero_overlap))
        # cand: (score_tuple..., ox, oy, h_ov, u_cov, h_cover, ar, oc, sm, edge, kiss)
        best: Optional[
            Tuple[
                Tuple[float, ...],
                int,
                int,
                float,
                float,
                float,
                float,
                Image.Image,
                Dict[str, Any],
                int,
                int,
            ]
        ] = None
        for ar in ratio_list:
            oc_try, sm = _size_for(ar)
            alpha = oc_try.split()[-1]
            sb_try = alpha.getbbox() or (0, 0, oc_try.width, oc_try.height)
            _sx0, _sy0, _sx1, sy1_try = (int(v) for v in sb_try)
            ox0, oy0, _m0, edge0, kiss_try = _kiss_place(oc_try)
            if pin_frame_bottom:
                # シール不透明下端 = キャンバス外枠下辺 (canvas-1)
                # PIL getbbox の lower は排他的 → oy + sy1_try - 1 = canvas-1
                oy_pin = int(canvas_size) - int(sy1_try)
                oy_list = [max(0, min(oy_pin, canvas_size - oc_try.height))]
            else:
                oy_candidates = [oy0]
                for dy in (40, 80, 120, 160, 200):
                    oy_candidates.append(max(0, oy0 - dy))
                seen_y = set()
                oy_list = []
                for yv in oy_candidates:
                    if yv not in seen_y:
                        seen_y.add(yv)
                        oy_list.append(yv)
            found_hard_zero_unit = False
            for oy_try in oy_list:
                # 下辺固定時は edge をその高さ帯で再取得
                if pin_frame_bottom and hero_mask is not None and hero_rgba is not None:
                    mid = oy_try + (_sy0 + sy1_try) // 2
                    band_y0 = max(0, mid - max(8, (sy1_try - _sy0) // 6))
                    band_y1 = min(canvas_size, mid + max(8, (sy1_try - _sy0) // 6))
                    edge0 = _hero_opaque_right_edge_(
                        hero_rgba,
                        x=hero_x,
                        y=hero_y,
                        scale=hero_scale,
                        canvas_size=canvas_size,
                        y0=band_y0,
                        y1=band_y1,
                    )
                ox_start = min(canvas_size - oc_try.width, max(ox0, edge0 - kiss_try))
                # 右→左へ走査。unit被覆を最小化し、同率なら右側を優先
                for ox_try in range(ox_start, -1, -4):
                    u_cov = _seal_mask_overlap_frac_(
                        alpha, avoid_mask, ox=ox_try, oy=oy_try, canvas_size=canvas_size
                    )
                    if hero_mask is not None:
                        h_ov = _seal_hero_overlap_frac_(
                            alpha, hero_mask, ox=ox_try, oy=oy_try, canvas_size=canvas_size
                        )
                        h_cov = _seal_hero_cover_of_hero_frac_(
                            alpha,
                            hero_mask,
                            ox=ox_try,
                            oy=oy_try,
                            canvas_size=canvas_size,
                            hero_ink=hero_ink_total,
                        )
                    else:
                        h_ov = float(overlap_frac)
                        h_cov = 0.0
                    cover_ok = max_hc is None or h_cov <= max_hc + 1e-9
                    if u_cov <= max_uc + 1e-9 and h_ov + 1e-9 >= min_ho and cover_ok:
                        # score: hero隠し超過なし前提で unit被覆→目標キス差→右寄せ
                        score = (u_cov, abs(h_ov - overlap_frac), h_cov, -float(ox_try))
                        cand = (
                            score,
                            ox_try,
                            oy_try,
                            h_ov,
                            u_cov,
                            h_cov,
                            ar,
                            oc_try,
                            sm,
                            edge0,
                            kiss_try,
                        )
                        if best is None or score < best[0]:
                            best = cand
                        if u_cov <= 1e-6:
                            found_hard_zero_unit = True
                            break  # この oy で完全非被覆の最右
                if found_hard_zero_unit:
                    break
            # 完全非被覆かつ hero 隠し上限内なら面積比ループ終了。
            # 隠しすぎなら一回り小さい面積比へ続行。
            if (
                best is not None
                and best[4] <= 1e-6
                and (max_hc is None or best[5] <= max_hc + 1e-9)
            ):
                break

        if best is not None:
            (
                _sc,
                ox,
                oy,
                measured,
                unit_cover,
                hero_cover,
                area_r,
                chosen_oc,
                size_meta,
                edge_x,
                kiss0,
            ) = best
            placed_ok = True
        else:
            # フォールバック: 下辺接触を優先。hero隠し最小の面積比を選ぶ
            pick_ar = ratio_list[-1] if ratio_list else area_r0
            chosen_oc, size_meta = _size_for(pick_ar)
            alpha = chosen_oc.split()[-1]
            sb_fb = alpha.getbbox() or (0, 0, chosen_oc.width, chosen_oc.height)
            _a, _b, _c, sy1_fb = (int(v) for v in sb_fb)
            if pin_frame_bottom:
                oy = max(0, min(int(canvas_size) - int(sy1_fb), canvas_size - chosen_oc.height))
                ox0, _oy0, measured, edge_x, kiss0 = _kiss_place(chosen_oc, oy_hint=oy)
                ox = ox0
                best_fb: Optional[Tuple[Tuple[float, ...], int, float, float, float]] = None
                for ox_try in range(min(ox0, canvas_size - chosen_oc.width), -1, -4):
                    h_ov = (
                        _seal_hero_overlap_frac_(
                            alpha, hero_mask, ox=ox_try, oy=oy, canvas_size=canvas_size
                        )
                        if hero_mask is not None
                        else float(overlap_frac)
                    )
                    u_cov = _seal_mask_overlap_frac_(
                        alpha, avoid_mask, ox=ox_try, oy=oy, canvas_size=canvas_size
                    )
                    h_cov = (
                        _seal_hero_cover_of_hero_frac_(
                            alpha,
                            hero_mask,
                            ox=ox_try,
                            oy=oy,
                            canvas_size=canvas_size,
                            hero_ink=hero_ink_total,
                        )
                        if hero_mask is not None
                        else 0.0
                    )
                    if h_ov + 1e-9 >= min_ho * 0.5:
                        score_fb = (u_cov, h_cov, -float(ox_try))
                        if best_fb is None or score_fb < best_fb[0]:
                            best_fb = (score_fb, ox_try, h_ov, u_cov, h_cov)
                        if u_cov <= max_uc + 1e-9 and (
                            max_hc is None or h_cov <= max_hc + 1e-9
                        ):
                            ox = ox_try
                            measured = h_ov
                            unit_cover = u_cov
                            hero_cover = h_cov
                            placed_ok = True
                            break
                if not placed_ok and best_fb is not None:
                    ox = best_fb[1]
                    measured = best_fb[2]
                    unit_cover = best_fb[3]
                    hero_cover = best_fb[4]
                area_r = pick_ar
            else:
                ox, oy, measured, edge_x, kiss0 = _kiss_place(chosen_oc)
                area_r = pick_ar
                unit_cover = _seal_mask_overlap_frac_(
                    alpha, avoid_mask, ox=ox, oy=oy, canvas_size=canvas_size
                )
                if hero_mask is not None:
                    hero_cover = _seal_hero_cover_of_hero_frac_(
                        alpha,
                        hero_mask,
                        ox=ox,
                        oy=oy,
                        canvas_size=canvas_size,
                        hero_ink=hero_ink_total,
                    )
                placed_ok = False

    assert chosen_oc is not None
    alpha = chosen_oc.split()[-1]
    sb = alpha.getbbox() or (0, 0, chosen_oc.width, chosen_oc.height)
    sx0, sy0, sx1, sy1 = (int(v) for v in sb)
    seal_bottom_y = int(oy) + int(sy1) - 1
    frame_bottom_contact = abs(seal_bottom_y - (canvas_size - 1)) <= 1
    if hero_mask is not None and max_hc is not None and avoid_mode:
        hero_cover = _seal_hero_cover_of_hero_frac_(
            alpha,
            hero_mask,
            ox=ox,
            oy=oy,
            canvas_size=canvas_size,
            hero_ink=hero_ink_total,
        )

    canvas.alpha_composite(chosen_oc, (ox, oy))
    seal_op = int(size_meta.get("sealOpaquePx") or count_opaque_pixels(chosen_oc))
    prod_op = int(size_meta.get("productOpaquePx") or 0)
    cover_ok_final = max_hc is None or hero_cover <= max_hc + 1e-9
    anchor = (
        "frame_bottom_avoid_units"
        if (avoid_mode and pin_frame_bottom)
        else ("hero_kiss_avoid_units" if avoid_mode else "hero_bottom_right_area_kiss")
    )
    note = (
        f"商品不透明ピクセル×{area_r:.3f}＝シール不透明。"
        f"右下キス（重なり≈{overlap_frac:.0%}）。"
    )
    if avoid_mode:
        note += (
            f" N≥5 unit非被覆（cover={unit_cover:.3f}"
            f"{' OK' if placed_ok else ' FALLBACK'}）。"
        )
    if pin_frame_bottom:
        note += " シール下端＝キャンバス外枠下辺接触。"
    if max_hc is not None:
        note += (
            f" hero隠し≤{max_hc:.0%}（実測{hero_cover:.1%}"
            f"{' OK' if cover_ok_final else ' OVER'}）。"
        )
    return {
        "x": ox,
        "y": oy,
        "w": chosen_oc.width,
        "h": chosen_oc.height,
        "floating": True,
        "anchor": anchor,
        "overlapFrac": overlap_frac,
        "kissPx": kiss0,
        "overlapMeasured": round(float(measured), 4),
        "unitCoverMeasured": round(float(unit_cover), 4) if avoid_mode else None,
        "heroCoverMeasured": round(float(hero_cover), 4) if max_hc is not None else None,
        "heroCoverOk": bool(cover_ok_final) if max_hc is not None else None,
        "avoidUnits": bool(avoid_mode),
        "avoidUnitsOk": bool(placed_ok) if avoid_mode else None,
        "pinFrameBottom": bool(pin_frame_bottom),
        "frameBottomContact": bool(frame_bottom_contact) if pin_frame_bottom else None,
        "sealBottomY": seal_bottom_y if pin_frame_bottom else None,
        "maxAvoidCover": float(max_avoid_cover) if avoid_mode else None,
        "minHeroOverlap": float(min_hero_overlap) if avoid_mode else None,
        "maxHeroCover": float(max_hc) if max_hc is not None else None,
        "heroEdgeX": edge_x,
        "sealOpaqueBox": [sx0, sy0, sx1, sy1],
        "sizeMode": size_meta.get("sizeMode") or "opaque_pixel_ratio",
        "productAreaRatio": area_r,
        "productOpaquePx": prod_op,
        "sealOpaquePx": seal_op,
        "targetOpaquePx": size_meta.get("targetOpaquePx"),
        "sealOpaqueRatioActual": size_meta.get("sealOpaqueRatioActual")
        or (round(seal_op / float(prod_op), 4) if prod_op else None),
        "heroWidthRatio": area_r ** 0.5,
        "heroBodyW": body_w,
        "heroBodyH": body_h,
        "heroBodyBox": [bx0, by0, bx1, by1],
        "heroFullBox": list(full),
        "anchorPoint": [edge_x, fy1],
        "noteJa": note,
    }




def scales_match_body_wh(
    hero_body: Dict[str, Any],
    unit_body: Dict[str, Any],
    *,
    target_body_w: float,
) -> Tuple[float, float, Dict[str, Any]]:
    """
    本体の縦横を同じターゲット枠 (Tw×Th) に contain するスケール。

    - Tw = target_body_w
    - Th = Tw × 平均アスペクト（本体H/W）
    - scale = min(Tw/bodyW, Th/bodyH)  … 各画像ごと（縦横比は維持）

    これにより『幅だけ揃え／高さ無視』で単体が大きく見える問題を避ける。
    """
    hb_w = max(1, int(hero_body["bodyW"]))
    hb_h = max(1, int(hero_body["bodyH"]))
    ub_w = max(1, int(unit_body["bodyW"]))
    ub_h = max(1, int(unit_body["bodyH"]))
    aspect_h = hb_h / hb_w
    aspect_u = ub_h / ub_w
    aspect = (aspect_h + aspect_u) / 2.0
    tw = float(target_body_w)
    th = tw * aspect
    hs = min(tw / hb_w, th / hb_h)
    us = min(tw / ub_w, th / ub_h)
    meta = {
        "metric": "body_contain_WH",
        "noteJa": (
            "本体外接（上部18%除外）の縦横を、同一ターゲット枠Tw×Thへcontain。"
            "画像全体の外接幅揃えではない（スプーン込み幅だと缶が小さく見える）。"
        ),
        "targetBodyW": round(tw, 2),
        "targetBodyH": round(th, 2),
        "heroBody": {"w": hb_w, "h": hb_h, "scale": round(hs, 4)},
        "unitBody": {"w": ub_w, "h": ub_h, "scale": round(us, 4)},
        "displayHeroBody": {
            "w": round(hb_w * hs, 1),
            "h": round(hb_h * hs, 1),
        },
        "displayUnitBody": {
            "w": round(ub_w * us, 1),
            "h": round(ub_h * us, 1),
        },
    }
    return hs, us, meta


def plan_small_n_same_scale(
    *,
    n: int,
    canvas: int,
    hero_full_w: int,
    hero_full_h: int,
    unit_full_w: int,
    unit_full_h: int,
    hero_body: Dict[str, Any],
    unit_body: Dict[str, Any],
    fill_min: float = 0.48,
    overlap: Optional[float] = None,
) -> Tuple[float, List[PastePlan], Dict[str, Any]]:
    """
    N≤4: 本体縦横を同一枠にcontain。奥行きは重なりとわずかなYオフセットのみ。
    """
    if n < 2 or n > 4:
        raise ValueError("plan_small_n_same_scale supports 2..4")

    margin = int(canvas * 0.02)
    if overlap is None:
        overlap = 0.36 if n == 2 else (0.32 if n == 3 else 0.28)
    size_meta: Dict[str, Any] = {}

    def layouts(target_body_w: float) -> List[PastePlan]:
        nonlocal size_meta
        hs, us, size_meta = scales_match_body_wh(
            hero_body, unit_body, target_body_w=target_body_w
        )
        hw = max(1, int(round(hero_full_w * hs)))
        hh = max(1, int(round(hero_full_h * hs)))
        uw = max(1, int(round(unit_full_w * us)))
        uh = max(1, int(round(unit_full_h * us)))
        plans: List[PastePlan] = []
        if n == 2:
            cluster_w = int(uw + hw * (1 - overlap))
            cluster_h = max(uh, hh) + int(hh * 0.04)
            x0 = max(margin, (canvas - cluster_w) // 2)
            # 上余白を減らす: やや下寄せではなく垂直センター寄り＋わずかに上
            y0 = max(margin, int((canvas - cluster_h) * 0.42))
            ux = x0
            uy = y0 + max(0, (cluster_h - uh) // 2) - int(hh * 0.02)
            hx = x0 + int(uw * (1 - overlap))
            hy = y0 + max(0, cluster_h - hh)
            plans.append(PastePlan("unit", ux, uy, us, z=0))
            plans.append(PastePlan("hero", hx, hy, hs, z=1))
            return plans

        # N=3/4: 横一列だとスケールが落ちる → 密な扇／段に詰める
        if n == 3:
            # 密な扇: 単体同士は重ね可。ヒーロー被りは overlap で制御
            u_ov = max(overlap, 0.34)
            u_step = int(uw * (1 - u_ov))
            h_inset = int(uw * (1 - overlap))
            cluster_w = u_step + h_inset + int(hw * 0.90)
            cluster_h = max(uh, hh) + int(hh * 0.05)
            x0 = max(margin, (canvas - cluster_w) // 2)
            y0 = max(margin, int((canvas - cluster_h) * 0.38))
            plans.append(PastePlan("unit", x0, y0 + int(hh * 0.04), us, z=0))
            plans.append(PastePlan("unit", x0 + u_step, y0, us, z=1))
            hx = x0 + u_step + h_inset - int(uw * 0.06)
            hy = y0 + max(0, cluster_h - hh)
            plans.append(PastePlan("hero", hx, hy, hs, z=100))
            return plans

        # n == 4: 背面3を強く重ねた弧 + 手前ヒーロー（占有確保）
        ov4 = max(overlap, 0.40)
        step_x = int(uw * (1 - ov4))
        cluster_w = step_x * 2 + int(hw * 0.85)
        cluster_h = max(uh, hh) + int(hh * 0.06)
        x0 = max(margin, (canvas - cluster_w) // 2)
        y0 = max(margin, int((canvas - cluster_h) * 0.36))
        for i in range(3):
            plans.append(
                PastePlan(
                    "unit",
                    x0 + i * step_x,
                    y0 + abs(i - 1) * int(uh * 0.04),
                    us,
                    z=i,
                )
            )
        hx = x0 + int(step_x * 0.9)
        hy = y0 + max(0, cluster_h - hh)
        plans.append(PastePlan("hero", hx, hy, hs, z=100))
        return plans

    def fits(plans: List[PastePlan]) -> bool:
        for p in plans:
            sw = hero_full_w if p.role == "hero" else unit_full_w
            sh = hero_full_h if p.role == "hero" else unit_full_h
            w = int(round(sw * p.scale))
            h = int(round(sh * p.scale))
            if p.x < 0 or p.y < 0:
                return False
            if p.x + w > canvas - margin // 2:
                return False
            if p.y + h > canvas - margin // 2:
                return False
        return True

    # target_body_w を大きい方から探索（本体幅の目標ピクセル）
    best_tw = canvas * 0.25
    best_plans: List[PastePlan] = layouts(best_tw)
    for tw in [canvas * x / 100 for x in range(75, 20, -1)]:
        plans = layouts(tw)
        if fits(plans):
            best_tw = tw
            best_plans = plans
            break

    LOG.info(
        "paste plan n=%s targetBodyW=%.1f heroScale=%.3f unitScale=%.3f overlap≈%.2f",
        n,
        best_tw,
        best_plans[-1].scale if best_plans else 0,
        next((p.scale for p in best_plans if p.role == "unit"), 0),
        overlap,
    )
    return best_tw, best_plans, size_meta


def compose_amazon_paste(
    *,
    hero_rgba: Image.Image,
    unit_rgba: Image.Image,
    set_count: int,
    canvas_size: int = 1200,
    octas_rgba: Optional[Image.Image] = None,
    fill_min: float = 0.48,
    layout_mode: str = "edge_fill",
    aspect: str = "square",
    octas_tilt_applied: bool = False,
    n3_cfg_override: Optional[Dict[str, Any]] = None,
) -> Tuple[Image.Image, Dict[str, Any]]:
    """
    戻り値: (RGB画像, meta)
    layout_mode:
      - edge_fill: 縁際・余白最小化（正方形本線）
      - legacy_body: 本体WH＋重なり探索（旧）
    aspect:
      - square: edge_fill／legacy
      - portrait: 縦長商品向け斜めファン（キャンバスは正方形のまま）
      - landscape: 横長商品向けセット組み（キャンバスは正方形のまま）
    octas_tilt_applied:
      - True: 呼び出し側で prepare_octas_seal 済み（二重回転しない）
      - False: 未傾きなら本線8°（上部を右）をここで適用
    n3_cfg_override:
      - 指定時は portraitTilt.n3 に上書きマージ（プレビュー用ノブ等）
    """
    n = int(set_count)
    if n < 1:
        raise ValueError("set_count >= 1 required")

    hero = _trim_rgba(hero_rgba)
    unit = _trim_rgba(unit_rgba)
    hero_body = measure_body_box(hero)
    unit_body = measure_body_box(unit)
    size_meta: Dict[str, Any] = {}
    overlap_meta: Dict[str, Any] = {"deferred": True, "noteJa": "余白配置を先に確定中"}
    layout_meta: Dict[str, Any] = {}
    aspect_id = str(aspect or "square").strip().lower() or "square"

    # 閾値は layout_rules.json の tuningDraft を正とする
    pair_overlap_max = 0.35
    hero_visible_min = 0.70
    portrait_cfg: Dict[str, Any] = {}
    landscape_cfg: Dict[str, Any] = {}
    octas_cfg: Dict[str, Any] = {}
    try:
        from rakuten_layer import load_layout_rules

        tun = (load_layout_rules().get("amazon") or {}).get("tuningDraft") or {}
        if tun.get("pairOverlapMax") is not None:
            pair_overlap_max = float(tun["pairOverlapMax"])
        if tun.get("heroVisibleMin") is not None:
            hero_visible_min = float(tun["heroVisibleMin"])
        portrait_cfg = tun.get("portraitTilt") or {}
        landscape_cfg = tun.get("landscapeTilt") or {}
        octas_cfg = tun.get("octasSeal") or {}
    except Exception:
        pass

    octas_overlap_frac = float(octas_cfg.get("overlapFrac", 0.10))
    octas_area_ratio = float(octas_cfg.get("productAreaRatio", 0.075))
    # N=2/3/4: 共通 1.5倍（0.075）。個別ノブがあれば上書き
    if n == 2:
        octas_area_ratio = float(octas_cfg.get("n2ProductAreaRatio", octas_area_ratio))
        n2_oct = (portrait_cfg.get("n2") or {}) if aspect_id == "portrait" else {}
        if n2_oct.get("octasProductAreaRatio") is not None:
            octas_area_ratio = float(n2_oct["octasProductAreaRatio"])
        if n2_oct.get("octasOverlapFrac") is not None:
            octas_overlap_frac = float(n2_oct["octasOverlapFrac"])
    elif n in (3, 4):
        if octas_cfg.get("n3n4ProductAreaRatio") is not None:
            octas_area_ratio = float(octas_cfg["n3n4ProductAreaRatio"])
        if aspect_id == "portrait":
            n3_oct = portrait_cfg.get("n3") or {}
            if n == 3 and n3_oct.get("octasProductAreaRatio") is not None:
                octas_area_ratio = float(n3_oct["octasProductAreaRatio"])
            n4_oct = portrait_cfg.get("n4") or {}
            if n == 4 and n4_oct.get("octasProductAreaRatio") is not None:
                octas_area_ratio = float(n4_oct["octasProductAreaRatio"])

    if aspect_id == "portrait":
        from portrait_layout import fit_portrait_layout_under_overlap

        pw = int(unit.width)
        ph = int(unit.height)
        hero_tilt = float(
            portrait_cfg.get("heroTiltDegCw", portrait_cfg.get("tiltDegCw", -12))
        )
        # 旧キー tiltDegCw が正だった頃の互換: hero は左＝負を強制したい場合
        if "heroTiltDegCw" not in portrait_cfg and hero_tilt > 0:
            hero_tilt = -abs(hero_tilt)
        unit_tilt = float(portrait_cfg.get("unitTiltDegCw", 16))
        unit_step = float(portrait_cfg.get("unitTiltStepDegCw", 3))
        n1_tilt = float(portrait_cfg.get("n1TiltDegCw") or 0)
        h_ov = float(portrait_cfg.get("hOverlap") or 0.50)
        v_ov = float(portrait_cfg.get("vOverlap") or 0.55)
        n2_cfg = portrait_cfg.get("n2") or {}
        n3_cfg = dict(portrait_cfg.get("n3") or {})
        if isinstance(n3_cfg_override, dict) and n3_cfg_override:
            n3_cfg.update(n3_cfg_override)
        n4_cfg = portrait_cfg.get("n4") or {}
        n5plus_cfg = portrait_cfg.get("n5plus") or {}
        # 縦長は密なファンのため重なり上限を少し緩める（見本準拠）
        p_ov_max = float(portrait_cfg.get("pairOverlapMax") or max(pair_overlap_max, 0.45))
        h_vis_min = float(portrait_cfg.get("heroVisibleMin") or min(hero_visible_min, 0.65))
        port_plans, layout_meta, overlap_meta = fit_portrait_layout_under_overlap(
            n=n,
            canvas=canvas_size,
            product_w=pw,
            product_h=ph,
            hero_rgba=hero,
            unit_rgba=unit,
            pair_overlap_max=p_ov_max,
            hero_visible_min=h_vis_min,
            hero_tilt_deg_cw=hero_tilt,
            unit_tilt_deg_cw=unit_tilt,
            unit_tilt_step_deg_cw=unit_step,
            n1_tilt_deg_cw=n1_tilt,
            h_overlap=h_ov,
            v_overlap=v_ov,
            n2_cfg=n2_cfg,
            n3_cfg=n3_cfg,
            n4_cfg=n4_cfg,
            n5plus_cfg=n5plus_cfg,
        )
        plans = [
            PastePlan(
                p.role,
                p.x,
                p.y,
                p.scale,
                p.z,
                rotation_deg=p.rotation_deg,
                anchor=getattr(p, "anchor", "topleft") or "topleft",
                foot_x=getattr(p, "foot_x", None),
                foot_y=getattr(p, "foot_y", None),
                top_x=getattr(p, "top_x", None),
                top_y=getattr(p, "top_y", None),
            )
            for p in port_plans
        ]
        scale = layout_meta.get("scale") or (plans[0].scale if plans else 0)
        size_meta = {
            "metric": "portrait_tilt_fan",
            "noteJa": (
                "縦長 N≥5: hero左＋右傾け積み。"
                if n >= 5
                else "縦長商品: 斜めファン＋余白埋め。キャンバスは正方形。"
            ),
            "productWh": [pw, ph],
        }
        layout_mode = "portrait_tilt_fan"

    elif aspect_id == "landscape":
        from landscape_layout import fit_landscape_layout_under_overlap

        pw = int(unit.width)
        ph = int(unit.height)
        land_plans, layout_meta, overlap_meta = fit_landscape_layout_under_overlap(
            n=n,
            canvas=canvas_size,
            product_w=pw,
            product_h=ph,
            hero_rgba=hero,
            unit_rgba=unit,
            landscape_cfg=landscape_cfg,
        )
        plans = [
            PastePlan(
                p.role,
                p.x,
                p.y,
                p.scale,
                p.z,
                rotation_deg=getattr(p, "rotation_deg", 0.0) or 0.0,
                anchor=getattr(p, "anchor", "topleft") or "topleft",
            )
            for p in land_plans
        ]
        scale = layout_meta.get("scale") or layout_meta.get("scaleHero") or (
            plans[0].scale if plans else 0
        )
        size_meta = {
            "metric": "landscape_set_geometry",
            "noteJa": "横長セット組み（N1中央／N2縦二段／N3–4階段／N≥5上hero+下グリッド）。",
            "productWh": [pw, ph],
        }
        layout_mode = "landscape_set"

    if aspect_id not in ("portrait", "landscape") and layout_mode == "edge_fill":
        from edge_layout import fit_edge_layout_under_overlap

        # 単体テスト時は hero=unit 同一。スケール基準は unit 外接
        pw = int(unit.width)
        ph = int(unit.height)
        edge_plans, layout_meta, overlap_meta = fit_edge_layout_under_overlap(
            n=n,
            canvas=canvas_size,
            product_w=pw,
            product_h=ph,
            hero_rgba=hero,
            unit_rgba=unit,
            pair_overlap_max=pair_overlap_max,
            hero_visible_min=hero_visible_min,
            overlap=0.28,
        )
        plans = [
            PastePlan(p.role, p.x, p.y, p.scale, p.z) for p in edge_plans
        ]
        scale = layout_meta.get("scale") or (plans[0].scale if plans else 0)
        size_meta = {
            "metric": "edge_fill_uniform_overlap_capped",
            "noteJa": "縁際パターン固定→重なり硬制約下でスケール最大化。",
            "productWh": [pw, ph],
        }
    elif aspect_id not in ("portrait", "landscape"):
        from overlap_metrics import measure_plans_overlap


        def _plan_with_ov(ov: float, plan_n: int):
            return plan_small_n_same_scale(
                n=plan_n,
                canvas=canvas_size,
                hero_full_w=hero.width,
                hero_full_h=hero.height,
                unit_full_w=unit.width,
                unit_full_h=unit.height,
                hero_body=hero_body,
                unit_body=unit_body,
                fill_min=fill_min,
                overlap=ov,
            )

        ov_candidates = [0.40, 0.36, 0.32, 0.28, 0.24, 0.20]
        scale, plans, size_meta = _plan_with_ov(0.32, min(n, 4))
        for ov in ov_candidates:
            sc, pl, sm = _plan_with_ov(ov, min(n, 4))
            om = measure_plans_overlap(
                hero=hero, unit=unit, plans=pl, canvas_size=canvas_size
            )
            ok = (
                om["pairMaxBackCovered"] <= pair_overlap_max
                and om["heroMinVisible"] >= hero_visible_min
            )
            if ok:
                scale, plans, size_meta = sc, pl, sm
                overlap_meta = {
                    "layoutOverlapParam": ov,
                    "measured": om,
                    "pass": True,
                    "deferred": False,
                }
                break
        layout_meta = {"layoutFamily": "legacy_body"}

    canvas = Image.new("RGBA", (canvas_size, canvas_size), (255, 255, 255, 255))
    # z昇順で貼る
    from portrait_layout import measure_foot_xy, measure_top_xy, rotate_rgba_cw

    # foot/top anchor: スケール＋回転後レイヤの足/頂を目標に合わせ、topleft を確定
    resolved: List[Tuple[PastePlan, Image.Image]] = []
    for p in sorted(plans, key=lambda x: x.z):
        src = hero if p.role == "hero" else unit
        layer = _scale_rgba(src, p.scale)
        if abs(float(getattr(p, "rotation_deg", 0) or 0)) > 1e-9:
            layer = rotate_rgba_cw(layer, float(p.rotation_deg))
        anc = getattr(p, "anchor", "topleft") or "topleft"
        if anc == "foot":
            fx_t = getattr(p, "foot_x", None)
            fy_t = getattr(p, "foot_y", None)
            if fx_t is not None and fy_t is not None:
                fx_l, fy_l = measure_foot_xy(layer)
                p.x = int(round(float(fx_t) - fx_l))
                p.y = int(round(float(fy_t) - fy_l))
        elif anc == "top":
            tx_t = getattr(p, "top_x", None)
            ty_t = getattr(p, "top_y", None)
            if tx_t is not None and ty_t is not None:
                tx_l, ty_l = measure_top_xy(layer)
                p.x = int(round(float(tx_t) - tx_l))
                p.y = int(round(float(ty_t) - ty_l))
        resolved.append((p, layer))

    for p, layer in resolved:
        # アルファ輪郭の影（矩形影は斜めで白背景に見える）
        shadow = _soft_shadow_from_alpha(
            layer, blur=max(10, int(min(layer.size) * 0.035)), opacity=50
        )
        sx = p.x - (shadow.width - layer.width) // 2 + 4
        sy = p.y - (shadow.height - layer.height) // 2 + 8
        canvas.alpha_composite(shadow, (max(0, sx), max(0, sy)))
        canvas.alpha_composite(layer, (p.x, p.y))

    octas_meta = None
    octas_tilt_meta = None
    if octas_rgba is not None:
        if not octas_tilt_applied:
            from octas_prep import tilt_octas_rgba

            octas_rgba, octas_tilt_meta = tilt_octas_rgba(
                octas_rgba,
                seed=hash(f"compose:{n}:{aspect_id}") % (2**31),
                direction="top_to_right",
            )
        hero_plans = [p for p in plans if p.role == "hero"]
        hero_layer = None
        if hero_plans:
            hp = hero_plans[0]
            hero_layer = _scale_rgba(hero, hp.scale)
            if abs(float(getattr(hp, "rotation_deg", 0) or 0)) > 1e-9:
                hero_layer = rotate_rgba_cw(hero_layer, float(hp.rotation_deg))
            # 回転後レイヤを仮キャンバスに置き、本体外接を取る
            body_box = hero_body_box_on_canvas(
                hero_layer, x=hp.x, y=hp.y, scale=1.0
            )
        else:
            body_box = {
                "x0": 0,
                "y0": 0,
                "x1": canvas_size,
                "y1": canvas_size,
                "bodyW": canvas_size,
                "bodyH": canvas_size,
                "fullBox": [0, 0, canvas_size, canvas_size],
            }
            hp = None

        # N≥5 portrait: unit 非被覆＋キャンバス下辺接触＋hero隠し上限で縮小
        # N≥5 landscape: unit 非被覆＋下枠ピンなし（heroが上のため）＋hero隠し上限
        avoid_mask = None
        pin_bottom = False
        max_avoid = 0.005
        min_hero_ov = 0.05
        max_hero_cov: Optional[float] = None
        area_fallbacks: List[float] = [0.055, 0.04, 0.03]
        if aspect_id in ("portrait", "landscape") and n >= 5:
            n5_cfg = (
                (portrait_cfg.get("n5plus") or {})
                if aspect_id == "portrait"
                else (landscape_cfg.get("n5plus") or {})
            )
            avoid_on = bool(octas_cfg.get("n5plusAvoidUnits", True))
            if n5_cfg.get("octasAvoidUnits") is not None:
                avoid_on = bool(n5_cfg["octasAvoidUnits"])
            if aspect_id == "landscape":
                pin_bottom = bool(n5_cfg.get("octasPinFrameBottom", False))
            else:
                pin_bottom = bool(octas_cfg.get("n5plusPinFrameBottom", True))
                if n5_cfg.get("octasPinFrameBottom") is not None:
                    pin_bottom = bool(n5_cfg["octasPinFrameBottom"])
            if avoid_on:
                avoid_mask = Image.new("L", (canvas_size, canvas_size), 0)
                for p, layer in resolved:
                    if p.role != "unit":
                        continue
                    m = Image.new("L", (canvas_size, canvas_size), 0)
                    m.paste(layer.split()[-1], (p.x, p.y))
                    avoid_mask = ImageChops.lighter(avoid_mask, m)
                max_avoid = float(
                    n5_cfg.get(
                        "octasMaxUnitCover",
                        octas_cfg.get("n5plusMaxUnitCover", 0.005),
                    )
                )
                min_hero_ov = float(
                    n5_cfg.get(
                        "octasMinHeroOverlap",
                        octas_cfg.get("n5plusMinHeroOverlap", 0.05),
                    )
                )
                max_hero_cov = float(
                    n5_cfg.get(
                        "octasMaxHeroCover",
                        octas_cfg.get("n5plusMaxHeroCover", 0.05),
                    )
                )
                fb = n5_cfg.get("octasAreaRatioFallbacks") or octas_cfg.get(
                    "n5plusAreaRatioFallbacks"
                )
                if isinstance(fb, list) and fb:
                    area_fallbacks = [float(x) for x in fb]

        octas_meta = place_octas_on_hero_body(
            canvas,
            octas_rgba,
            body_box=body_box,
            canvas_size=canvas_size,
            overlap_frac=octas_overlap_frac,
            area_ratio=octas_area_ratio,
            hero_rgba=hero_layer if hero_plans else None,
            hero_x=hp.x if hero_plans else 0,
            hero_y=hp.y if hero_plans else 0,
            hero_scale=1.0 if hero_plans else 1.0,
            avoid_mask=avoid_mask,
            max_avoid_cover=max_avoid,
            min_hero_overlap=min_hero_ov,
            max_hero_cover=max_hero_cov if avoid_mask is not None else None,
            area_ratio_fallbacks=area_fallbacks if avoid_mask is not None else None,
            pin_frame_bottom=bool(pin_bottom and avoid_mask is not None),
        )
        if octas_meta is not None:
            octas_meta["setCount"] = n
            octas_meta["sizePolicyJa"] = "N共通1.5倍ロック(0.075)/傾き8°"
            if avoid_mask is not None:
                if aspect_id == "landscape":
                    octas_meta["sizePolicyJa"] += "／N≥5 unit非被覆＋hero隠し上限（下枠ピンなし）"
                else:
                    octas_meta["sizePolicyJa"] += "／N≥5 unit非被覆＋下枠接触＋hero隠し上限で縮小"
            if octas_tilt_meta is not None:
                octas_meta["tiltPrep"] = octas_tilt_meta
                octas_meta["tiltAppliedInCompose"] = True
            else:
                octas_meta["tiltAppliedInCompose"] = False
                octas_meta["tiltAppliedUpstream"] = bool(octas_tilt_applied)

    # 最終クランプ（はみ出し防止は plan 側。念のため）
    out = canvas.convert("RGB")
    meta: Dict[str, Any] = {
        "mode": "amazon_pillow_paste",
        "layoutMode": layout_mode,
        "aspect": aspect_id,
        "setCount": n,
        "canvas": canvas_size,
        "scale": scale,
        "sameScaleAll": True,
        "aspectLock": True,
        "sizeMetric": size_meta,
        "overlapMetric": overlap_meta,
        "layoutMeta": layout_meta,
        "heroBody": hero_body,
        "unitBody": unit_body,
        "plans": [
            {
                "role": p.role,
                "x": p.x,
                "y": p.y,
                "scale": p.scale,
                "z": p.z,
                "rotationDeg": float(getattr(p, "rotation_deg", 0) or 0),
                "anchor": getattr(p, "anchor", "topleft") or "topleft",
                "footX": getattr(p, "foot_x", None),
                "footY": getattr(p, "foot_y", None),
                "topX": getattr(p, "top_x", None),
                "topY": getattr(p, "top_y", None),
            }
            for p in plans
        ],
        "octas": octas_meta,
        "heroSizeSrc": list(hero.size),
        "unitSizeSrc": list(unit.size),
    }
    return out, meta


def run_paste_file(
    *,
    hero_path: Path,
    unit_path: Path,
    set_count: int,
    out_path: Path,
    octas_path: Optional[Path] = None,
    canvas_size: int = 1200,
    fill_min: float = 0.48,
    layout_mode: str = "edge_fill",
    aspect: str = "square",
) -> Dict[str, Any]:
    from transparent_bg import ensure_transparent_product, inspect_alpha
    from octas_prep import prepare_octas_seal
    from fill_metrics import ink_fill_ratio

    cache = hero_path.parent / "_transparent_cache"
    # ensure 内で透過チェックメッセージを1回出す。meta 用は再検査（announce無し）
    hero_a = ensure_transparent_product(hero_path, cache)
    if Path(hero_path).resolve() == Path(unit_path).resolve():
        unit_a = hero_a
        hero_check = inspect_alpha(Path(hero_path), announce=False)
        unit_check = hero_check
    else:
        unit_a = ensure_transparent_product(unit_path, cache)
        hero_check = inspect_alpha(Path(hero_path), announce=False)
        unit_check = inspect_alpha(Path(unit_path), announce=False)
    hero_im = Image.open(hero_a)
    unit_im = Image.open(unit_a)

    octas_im = None
    octas_prep = None
    if octas_path and Path(octas_path).is_file():
        tilted, octas_prep = prepare_octas_seal(
            Path(octas_path),
            Path(octas_path).parent / "_tilt_cache",
            seed=hash(f"{hero_path.name}:{set_count}") % (2**31),
            direction="top_to_right",
        )
        octas_im = Image.open(tilted)

    img, meta = compose_amazon_paste(
        hero_rgba=hero_im,
        unit_rgba=unit_im,
        set_count=set_count,
        canvas_size=canvas_size,
        octas_rgba=octas_im,
        fill_min=fill_min,
        layout_mode=layout_mode,
        aspect=aspect,
        octas_tilt_applied=bool(octas_im is not None),
    )
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, format="JPEG", quality=90, optimize=True)
    meas = ink_fill_ratio(out_path)
    meta["output"] = str(out_path)
    meta["fill"] = meas
    meta["fillPass"] = meas["inkFillRatio"] >= fill_min
    meta["octasPrep"] = octas_prep
    meta["heroPath"] = str(hero_path)
    meta["unitPath"] = str(unit_path)
    meta["alphaCheck"] = {
        "hero": hero_check,
        "unit": unit_check,
        "policyJa": "透過PNGを本線。アルファ無しは自動白抜きフォールバック。",
    }

    # portrait: 直立1個四隅注釈。素材(unit)単位で代表1枚だけ _meta へ保存
    if str(aspect).strip().lower() == "portrait":
        try:
            from rakuten_layer import load_layout_rules
            from portrait_layout import render_upright_quad_annot

            pt = (
                (load_layout_rules().get("amazon") or {})
                .get("tuningDraft", {})
                .get("portraitTilt")
                or {}
            )
            export_annot = bool(pt.get("exportUprightQuadAnnot", True))
            if export_annot:
                once = bool(pt.get("exportUprightQuadAnnotOncePerUnit", True))
                disp = float(pt.get("exportUprightQuadAnnotScale", 0.90))
                unit_trim = _trim_rgba(unit_im.convert("RGBA"))
                unit_key = Path(unit_path).resolve().stem
                safe = "".join(
                    ch if (ch.isalnum() or ch in "-_") else "_" for ch in unit_key
                )[:80]
                meta_dir = out_path.parent / "_meta"
                if once:
                    annot_path = meta_dir / f"upright_quad__{safe}.jpg"
                    meta["uprightQuad"] = (
                        (meta.get("layoutMeta") or {}).get("productQuadUpright")
                    )
                    if annot_path.is_file():
                        meta["uprightQuadAnnot"] = str(annot_path)
                        meta["uprightQuadAnnotSkipped"] = "already_exists_same_unit"
                        LOG.info(
                            "paste skip upright-quad annot (same unit) %s",
                            annot_path.name,
                        )
                    else:
                        annot = render_upright_quad_annot(
                            unit_trim,
                            display_scale=disp,
                            label=f"{safe} upright-quad",
                        )
                        meta_dir.mkdir(parents=True, exist_ok=True)
                        annot.save(annot_path, format="JPEG", quality=92, optimize=True)
                        meta["uprightQuadAnnot"] = str(annot_path)
                        LOG.info("paste wrote upright-quad annot %s", annot_path.name)
                else:
                    annot = render_upright_quad_annot(
                        unit_trim,
                        display_scale=disp,
                        label=f"{out_path.stem} upright-quad",
                    )
                    annot_path = out_path.with_name(f"{out_path.stem}_upright_quad.jpg")
                    annot.save(annot_path, format="JPEG", quality=92, optimize=True)
                    meta["uprightQuadAnnot"] = str(annot_path)
                    meta["uprightQuad"] = (
                        (meta.get("layoutMeta") or {}).get("productQuadUpright")
                    )
                    LOG.info("paste wrote upright-quad annot %s", annot_path.name)
        except Exception as exc:
            LOG.warning("upright-quad annot failed: %s", exc)
            meta["uprightQuadAnnotError"] = str(exc)

    LOG.info(
        "paste wrote %s fill=%.3f pass=%s scale=%.3f alphaHero=%s alphaUnit=%s",
        out_path.name,
        meas["inkFillRatio"],
        meta["fillPass"],
        meta["scale"],
        hero_check.get("hasTransparency"),
        unit_check.get("hasTransparency"),
    )
    return meta
