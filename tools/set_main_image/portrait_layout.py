# -*- coding: utf-8 -*-
"""
縦長パターン（portrait）— 正方形キャンバス＋逆向き斜め

正方形 edge_fill との共通:
- 余白を埋める（スケール最大化）
- hero はレイヤー最前面
- 縦横比ロック（等倍＋回転のみ。ストレッチ禁止）
- 素材は 01.amazon白抜きベース（透過PNG）

差分（見本: 03.amazon見本/縦型レイアウト基本）:
- N=1: 直立・中央最大化
- N=2: の字型。hero上下目一杯・左傾き。unitは逆傾き・やや下・株見切れ可
- N=3: 頂円弧扇（合格固定）
- N=4: 高さ／扇ロジック（合格固定・propose_n4_fan_plans）
- N≥5: 正方形型と同様 hero左＋右unit列積み。縦長のため unit を約30°傾けて積む
"""
from __future__ import annotations

import logging
import math
from dataclasses import dataclass, replace
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image

LOG = logging.getLogger("set_main_image.portrait_layout")

EDGE_MARGIN_RATIO = 0.02
DEFAULT_HERO_TILT_DEG_CW = -12.0
DEFAULT_UNIT_TILT_DEG_CW = 16.0
DEFAULT_UNIT_TILT_STEP_DEG_CW = 3.0
DEFAULT_N1_TILT_DEG_CW = 0.0
DEFAULT_H_OVERLAP = 0.50
DEFAULT_V_OVERLAP = 0.55


@dataclass
class PortraitPlan:
    role: str
    x: int
    y: int
    scale: float
    z: int
    rotation_deg: float = 0.0  # 時計回り度（貼付時は PIL へ -deg）
    anchor: str = "topleft"  # topleft | foot | top
    foot_x: Optional[int] = None
    foot_y: Optional[int] = None
    top_x: Optional[int] = None
    top_y: Optional[int] = None


def aabb_after_rotate(w: int, h: int, tilt_deg_cw: float) -> Tuple[int, int]:
    if abs(tilt_deg_cw) < 1e-9:
        return max(1, int(w)), max(1, int(h))
    rad = math.radians(float(tilt_deg_cw))
    c, s = abs(math.cos(rad)), abs(math.sin(rad))
    rw = int(math.ceil(w * c + h * s))
    rh = int(math.ceil(w * s + h * c))
    return max(1, rw), max(1, rh)


def rotate_rgba_cw(im: Image.Image, tilt_deg_cw: float) -> Image.Image:
    im = im.convert("RGBA")
    if abs(tilt_deg_cw) < 1e-9:
        return im
    return im.rotate(
        -float(tilt_deg_cw),
        expand=True,
        resample=Image.Resampling.BICUBIC,
        fillcolor=(0, 0, 0, 0),
    )


def measure_foot_xy(im: Image.Image, *, alpha_thresh: int = 96) -> Tuple[float, float]:
    """
    不透明画素の『足』＝最下帯の中心。傾きの起点合わせに使う。
    alpha_thresh は高めにしてソフト縁を足に含めない（本体底を基準）。
    """
    im = im.convert("RGBA")
    a = im.getchannel("A")
    w, h = im.size
    ap = a.load()
    for y in range(h - 1, -1, -1):
        row = False
        for x in range(w):
            if ap[x, y] > alpha_thresh:
                row = True
                break
        if row:
            # 最下から数px帯
            band = min(28, y + 1)
            xs = []
            y0 = max(0, y - band + 1)
            for yy in range(y0, y + 1):
                for x in range(w):
                    if ap[x, yy] > alpha_thresh:
                        xs.append(x)
            if not xs:
                continue
            return (sum(xs) / len(xs), float(y))
    bb = a.getbbox() or (0, 0, w, h)
    return ((bb[0] + bb[2]) / 2.0, float(bb[3] - 1))


def measure_bottom_left_xy(
    im: Image.Image, *, alpha_thresh: int = 96, band: int = 28
) -> Tuple[float, float]:
    """
    商品底辺の左端＝最下帯の最左不透明点。N=4扇の『起点』定義。
    """
    im = im.convert("RGBA")
    a = im.getchannel("A")
    w, h = im.size
    ap = a.load()
    bottom_y = None
    for y in range(h - 1, -1, -1):
        if any(ap[x, y] > alpha_thresh for x in range(w)):
            bottom_y = y
            break
    if bottom_y is None:
        bb = a.getbbox() or (0, 0, w, h)
        return (float(bb[0]), float(bb[3] - 1))
    y0 = max(0, bottom_y - max(4, int(band)) + 1)
    best_x = w
    best_y = float(bottom_y)
    for yy in range(y0, bottom_y + 1):
        for x in range(w):
            if ap[x, yy] > alpha_thresh:
                if x < best_x:
                    best_x = x
                    best_y = float(yy)
                break
    if best_x >= w:
        bb = a.getbbox() or (0, 0, w, h)
        return (float(bb[0]), float(bb[3] - 1))
    return (float(best_x), float(best_y))


def measure_bottom_right_xy(
    im: Image.Image, *, alpha_thresh: int = 96, band: int = 28
) -> Tuple[float, float]:
    """
    商品底辺の右縁＝最下帯の最右不透明点。N=4の4本目・枠下接触に使う。
    """
    im = im.convert("RGBA")
    a = im.getchannel("A")
    w, h = im.size
    ap = a.load()
    bottom_y = None
    for y in range(h - 1, -1, -1):
        if any(ap[x, y] > alpha_thresh for x in range(w)):
            bottom_y = y
            break
    if bottom_y is None:
        bb = a.getbbox() or (0, 0, w, h)
        return (float(bb[2] - 1), float(bb[3] - 1))
    y0 = max(0, bottom_y - max(4, int(band)) + 1)
    best_x = -1
    best_y = float(bottom_y)
    for yy in range(y0, bottom_y + 1):
        for x in range(w - 1, -1, -1):
            if ap[x, yy] > alpha_thresh:
                if x > best_x:
                    best_x = x
                    best_y = float(yy)
                break
    if best_x < 0:
        bb = a.getbbox() or (0, 0, w, h)
        return (float(bb[2] - 1), float(bb[3] - 1))
    return (float(best_x), float(best_y))


def measure_top_xy(im: Image.Image, *, alpha_thresh: int = 96) -> Tuple[float, float]:
    """
    不透明画素の『頂』＝最上帯の中心。N=3上部円弧合わせに使う。
    """
    im = im.convert("RGBA")
    a = im.getchannel("A")
    w, h = im.size
    ap = a.load()
    for y in range(0, h):
        xs = [x for x in range(w) if ap[x, y] > alpha_thresh]
        if xs:
            band = min(28, h - y)
            xs2: List[int] = []
            for yy in range(y, min(h, y + band)):
                for x in range(w):
                    if ap[x, yy] > alpha_thresh:
                        xs2.append(x)
            if not xs2:
                continue
            return (sum(xs2) / len(xs2), float(y))
    bb = a.getbbox() or (0, 0, w, h)
    return ((bb[0] + bb[2]) / 2.0, float(bb[1]))


def measure_top_right_xy(im: Image.Image, *, alpha_thresh: int = 96, band: int = 36) -> Tuple[float, float]:
    """
    最上帯の最右不透明点。扇の右端接触（キャップ側）に使う。
    """
    im = im.convert("RGBA")
    a = im.getchannel("A")
    w, h = im.size
    ap = a.load()
    top_y = None
    for y in range(0, h):
        if any(ap[x, y] > alpha_thresh for x in range(w)):
            top_y = y
            break
    if top_y is None:
        bb = a.getbbox() or (0, 0, w, h)
        return (float(bb[2] - 1), float(bb[1]))
    y1 = min(h, top_y + max(8, int(band)))
    best = (-1, float(top_y))
    for yy in range(top_y, y1):
        for x in range(w - 1, -1, -1):
            if ap[x, yy] > alpha_thresh:
                if x > best[0]:
                    best = (x, float(yy))
                break
    if best[0] < 0:
        return measure_top_xy(im, alpha_thresh=alpha_thresh)
    return (float(best[0]), float(best[1]))


def measure_top_left_xy(im: Image.Image, *, alpha_thresh: int = 96, band: int = 36) -> Tuple[float, float]:
    """
    最上帯の最左不透明点。3本目の左右均等配置に使う。
    """
    im = im.convert("RGBA")
    a = im.getchannel("A")
    w, h = im.size
    ap = a.load()
    top_y = None
    for y in range(0, h):
        if any(ap[x, y] > alpha_thresh for x in range(w)):
            top_y = y
            break
    if top_y is None:
        bb = a.getbbox() or (0, 0, w, h)
        return (float(bb[0]), float(bb[1]))
    y1 = min(h, top_y + max(8, int(band)))
    best = (w, float(top_y))
    for yy in range(top_y, y1):
        for x in range(w):
            if ap[x, yy] > alpha_thresh:
                if x < best[0]:
                    best = (x, float(yy))
                break
    if best[0] >= w:
        return measure_top_xy(im, alpha_thresh=alpha_thresh)
    return (float(best[0]), float(best[1]))


def tilts_for_n(
    n: int,
    *,
    n1_tilt_deg_cw: float = DEFAULT_N1_TILT_DEG_CW,
    hero_tilt_deg_cw: float = DEFAULT_HERO_TILT_DEG_CW,
    unit_tilt_deg_cw: float = DEFAULT_UNIT_TILT_DEG_CW,
    unit_tilt_step_deg_cw: float = DEFAULT_UNIT_TILT_STEP_DEG_CW,
) -> List[float]:
    n = max(1, int(n))
    if n == 1:
        return [float(n1_tilt_deg_cw)]
    out = [float(hero_tilt_deg_cw)]
    for i in range(1, n):
        out.append(float(unit_tilt_deg_cw) + (i - 1) * float(unit_tilt_step_deg_cw))
    return out


def _cluster_rel_offsets(
    n: int, rw: int, rh: int, h_ov: float, v_ov: float
) -> List[Tuple[float, float]]:
    n = max(1, int(n))
    h_ov = max(0.15, min(0.75, float(h_ov)))
    v_ov = max(0.15, min(0.85, float(v_ov)))
    step_x = rw * (1.0 - h_ov)
    step_y = -rh * (1.0 - v_ov) * 0.35
    return [(i * step_x, i * step_y) for i in range(n)]


def _fit_scale_for_cluster(
    *,
    n: int,
    canvas: int,
    sizes: List[Tuple[int, int]],
    margin: int,
    h_ov: float,
    v_ov: float,
) -> float:
    usable = canvas - 2 * margin
    if usable <= 0 or not sizes:
        return 0.05
    rw_ref = max(s[0] for s in sizes)
    rh_ref = max(s[1] for s in sizes)
    offs = _cluster_rel_offsets(n, rw_ref, rh_ref, h_ov, v_ov)
    xs: List[float] = []
    ys: List[float] = []
    for i, (ox, oy) in enumerate(offs):
        sw, sh = sizes[i]
        xs.extend([ox, ox + sw])
        ys.extend([oy, oy + sh])
    cw = max(xs) - min(xs)
    ch = max(ys) - min(ys)
    if cw <= 0 or ch <= 0:
        return 0.05
    return max(0.05, min(usable / cw, usable / ch))


def propose_n2_noji_plans(
    *,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_tilt_deg_cw: float = -12.0,
    unit_tilt_deg_cw: float = 23.0,
    hero_height_fill: float = 1.10,
    hero_max_width_ratio: float = 0.92,
    unit_scale_ratio: float = 0.97,
    shared_foot_y_ratio: float = 0.9992,
    hero_foot_x_ratio: float = 0.4167,
    unit_foot_x_ratio: float = 0.5500,
    unit_base_crop_ratio: float = 0.03,
    scale: Optional[float] = None,
    hero_scale: Optional[float] = None,
    unit_overlap_x_ratio: float = 0.42,
    unit_down_shift_ratio: float = 0.05,
    hero_left_bias_ratio: float = 0.08,
    hero_rgba: Optional[Any] = None,
    unit_rgba: Optional[Any] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    2本・の字型（直立四隅モデル → PIL一致変換 → 枠接触）:
    1) 直立PNGで ProductQuad（四隅+4辺）を定義
    2) 同尺・傾きを四隅に適用（PIL rotate と一致）
    3) hero: TL→枠左 / BL→枠下を優先。ただし四隅は必ず青枠内（はみ出し禁止）
       unit: 既定 TR→枠右 / BR→枠下。縮小等で TL が枠上から離れたら
             TL→枠上 / TR→枠右へ切替（BR下は解除。接触以外はみ出し可）
    4) scale 指定時はその同尺（hero枠内 cap でクランプ）。未指定時は枠内最大同尺。
       unit可視≥閾値での最大化は fit_portrait_layout_under_overlap 側。
    """
    canvas = int(canvas)
    ht = float(hero_tilt_deg_cw)
    if ht > 0:
        ht = -abs(ht)
    ut = float(unit_tilt_deg_cw)
    if ut < 0:
        ut = abs(ut)

    src = None
    if isinstance(unit_rgba, Image.Image):
        src = unit_rgba.convert("RGBA")
    elif isinstance(hero_rgba, Image.Image):
        src = hero_rgba.convert("RGBA")

    if src is not None:
        quad0 = ProductQuad.from_upright_rgba(src)

        def _quad_fits(q: ProductQuad, lim: float) -> bool:
            min_x, min_y, max_x, max_y = q.aabb()
            return (max_x - min_x) <= lim and (max_y - min_y) <= lim

        # hero が枠内に収まる最大同尺（scale_cap）。指定 scale があれば min(scale, scale_cap)
        lim = float(canvas - 1)
        lo_c, hi_c = 0.05, 1.60
        for _ in range(28):
            mid = (lo_c + hi_c) * 0.5
            qh = quad0.transformed(scale=mid, tilt_deg_cw=ht)
            if _quad_fits(qh, lim):
                lo_c = mid
            else:
                hi_c = mid
        scale_cap = lo_c
        if scale is not None:
            scale_h = max(0.05, min(float(scale), scale_cap))
        else:
            scale_h = scale_cap
        scale_u = scale_h

        h_quad = quad0.transformed(scale=scale_h, tilt_deg_cw=ht)
        u_quad = quad0.transformed(scale=scale_u, tilt_deg_cw=ut)

        # hero: 接触優先 → はみ出しなら内側へクランプ（完全枠内）
        h_paste_x = 0.0 - float(h_quad.tl[0])
        h_paste_y = float(canvas - 1) - float(h_quad.bl[1])
        h_min_x, h_min_y, h_max_x, h_max_y = h_quad.aabb()
        w_min_x = h_min_x + h_paste_x
        w_max_x = h_max_x + h_paste_x
        w_min_y = h_min_y + h_paste_y
        w_max_y = h_max_y + h_paste_y
        if w_min_x < 0:
            h_paste_x += -w_min_x
            w_min_x = 0.0
            w_max_x = h_max_x + h_paste_x
        if w_max_x > lim:
            h_paste_x -= w_max_x - lim
            w_min_x = h_min_x + h_paste_x
            w_max_x = h_max_x + h_paste_x
        if w_min_y < 0:
            h_paste_y += -w_min_y
            w_min_y = 0.0
            w_max_y = h_max_y + h_paste_y
        if w_max_y > lim:
            h_paste_y -= w_max_y - lim
            w_min_y = h_min_y + h_paste_y
            w_max_y = h_max_y + h_paste_y
        if w_min_x < -1e-6 or w_max_x > lim + 1e-6:
            h_paste_x = -h_min_x + (lim - (h_max_x - h_min_x)) * 0.5
        if w_min_y < -1e-6 or w_max_y > lim + 1e-6:
            h_paste_y = -h_min_y + (lim - (h_max_y - h_min_y)) * 0.5

        # unit: 既定 TR=枠右 / BR=枠下。
        # 同尺縮小などで TL が枠上から離れたら TL=枠上 / TR=枠右へ切替（BR下は解除）
        u_paste_x = float(canvas - 1) - float(u_quad.tr[0])
        u_paste_y = float(canvas - 1) - float(u_quad.br[1])
        unit_anchor_mode = "tr_right_br_bottom"
        tl_y_after_br = u_paste_y + float(u_quad.tl[1])
        top_gap_eps = 2.0
        if tl_y_after_br > top_gap_eps:
            u_paste_y = 0.0 - float(u_quad.tl[1])
            unit_anchor_mode = "tr_right_tl_top"

        plans = [
            PortraitPlan(
                "unit",
                int(round(u_paste_x)),
                int(round(u_paste_y)),
                scale_u,
                z=0,
                rotation_deg=ut,
                anchor="topleft",
            ),
            PortraitPlan(
                "hero",
                int(round(h_paste_x)),
                int(round(h_paste_y)),
                scale_h,
                z=1000,
                rotation_deg=ht,
                anchor="topleft",
            ),
        ]

        def _world(q: ProductQuad, px: float, py: float) -> Dict[str, Any]:
            return {
                "tl": [round(px + q.tl[0], 1), round(py + q.tl[1], 1)],
                "tr": [round(px + q.tr[0], 1), round(py + q.tr[1], 1)],
                "bl": [round(px + q.bl[0], 1), round(py + q.bl[1], 1)],
                "br": [round(px + q.br[0], 1), round(py + q.br[1], 1)],
            }

        hw = _world(h_quad, h_paste_x, h_paste_y)
        uw = _world(u_quad, u_paste_x, u_paste_y)

        def _inside_frame(pt: List[float], eps: float = 1.5) -> bool:
            return (
                -eps <= float(pt[0]) <= lim + eps
                and -eps <= float(pt[1]) <= lim + eps
            )

        hero_all_inside = all(_inside_frame(hw[k]) for k in ("tl", "tr", "bl", "br"))
        logic_ja = [
            "①直立PNGで四隅TL/TR/BL/BRと4辺を定義",
            "②同尺・傾きを四隅に変換（PIL rotateと一致）",
            "③hero: TL→枠左/BL→枠下を優先、四隅は必ず青枠内",
            "④unit: 既定TR右+BR下。縮小で上隙間が出たらTL上+TR右へ切替",
            "⑤大きさ: hero枠内上限の下で unit可視≥unitVisibleMin となる最大同尺（fit側）",
            f"傾き hero{ht:.0f} / unit+{ut:.0f}",
        ]
        meta: Dict[str, Any] = {
            "layoutFamily": "portrait_n2_upright_quad",
            "pattern": "n2_upright_quad_corner_contact",
            "status": "locked_pass",
            "lockedAt": "2026-08-07",
            "n": 2,
            "scale": scale_h,
            "scaleMax": scale_cap,
            "scaleCapHeroInside": scale_cap,
            "scaleHero": scale_h,
            "scaleUnit": scale_u,
            "sameScale": True,
            "tiltsDegCw": [round(ht, 2), round(ut, 2)],
            "heroTiltDegCw": ht,
            "unitTiltDegCw": ut,
            "productQuadUpright": quad0.as_dict(),
            "heroQuadLayer": h_quad.as_dict(),
            "unitQuadLayer": u_quad.as_dict(),
            "heroQuadWorld": hw,
            "unitQuadWorld": uw,
            "heroContacts": {
                "topLeftToFrameLeft": abs(float(hw["tl"][0]) - 0.0) < 2.0,
                "bottomLeftToFrameBottom": abs(float(hw["bl"][1]) - lim) < 2.0,
                "allCornersInsideFrame": hero_all_inside,
                "overflowPolicy": "forbidden",
            },
            "unitContacts": {
                "topRightToFrameRight": abs(float(uw["tr"][0]) - lim) < 2.0,
                "bottomRightToFrameBottom": abs(float(uw["br"][1]) - lim) < 2.0,
                "topLeftToFrameTop": abs(float(uw["tl"][1]) - 0.0) < 2.0,
                "anchorMode": unit_anchor_mode,
                "overflowPolicy": "non_contact_corners_allowed",
            },
            "logicJa": logic_ja,
            "ruleStepsJa": logic_ja,
            "anchor": "topleft",
            "zOrderJa": "hero最前面。PIL一致四隅→hero完全枠内／unit指定接触",
            "proposalJa": " / ".join(logic_ja),
            "canvasStillSquare": True,
            "sampleAlignJa": "直立で四隅定義。変換はPIL一致。大きさはunit可視でfit",
        }
        return plans, meta

    # rgba 無し時の旧フォールバック（足比）
    rw_h1, rh_h1 = aabb_after_rotate(product_w, product_h, ht)
    fill = max(0.92, min(1.20, float(hero_height_fill)))
    target_h = canvas * fill
    scale_h0 = target_h / float(rh_h1)
    scale_w = (canvas * float(hero_max_width_ratio)) / float(rw_h1)
    scale_max = max(0.05, min(scale_h0, scale_w))
    if hero_scale is not None:
        scale_hero = max(0.05, min(1.15, float(hero_scale)))
    elif scale is None:
        scale_hero = scale_max
    else:
        scale_hero = max(0.05, min(float(scale), max(scale_max * 1.08, 1.05)))
    scale_unit = scale_hero * max(0.85, min(1.05, float(unit_scale_ratio)))

    foot_y_i = int(round(canvas * float(shared_foot_y_ratio)))
    foot_y_i = max(int(canvas * 0.92), min(canvas - 1, foot_y_i))
    hero_fx = int(round(canvas * float(hero_foot_x_ratio)))
    unit_fx = int(round(canvas * float(unit_foot_x_ratio)))

    plans = [
        PortraitPlan(
            "unit",
            0,
            0,
            scale_unit,
            z=0,
            rotation_deg=ut,
            anchor="foot",
            foot_x=unit_fx,
            foot_y=foot_y_i,
        ),
        PortraitPlan(
            "hero",
            0,
            0,
            scale_hero,
            z=1000,
            rotation_deg=ht,
            anchor="foot",
            foot_x=hero_fx,
            foot_y=foot_y_i,
        ),
    ]
    meta = {
        "layoutFamily": "portrait_n2_noji_foot_pivot",
        "pattern": "n2_noji_shared_foot_baseline",
        "n": 2,
        "scale": scale_hero,
        "scaleMax": scale_max,
        "scaleHero": scale_hero,
        "scaleUnit": scale_unit,
        "tiltsDegCw": [round(ht, 2), round(ut, 2)],
        "heroTiltDegCw": ht,
        "unitTiltDegCw": ut,
        "heroHeightFill": fill,
        "sharedFootY": foot_y_i,
        "heroFootX": hero_fx,
        "unitFootX": unit_fx,
        "sharedFootYRatio": float(shared_foot_y_ratio),
        "heroFootXRatio": float(hero_foot_x_ratio),
        "unitFootXRatio": float(unit_foot_x_ratio),
        "unitBaseCropRatio": float(unit_base_crop_ratio),
        "anchor": "foot",
        "zOrderJa": "hero最前面。傾き起点=共有の足Y",
        "proposalJa": "N=2: rgba無しフォールバック（見本126足比）",
        "canvasStillSquare": True,
        "sampleAlignJa": "fallback foot ratios",
    }
    return plans, meta


def propose_n3_upright_quad_fan_plans(
    *,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_tilt_deg_cw: float = -14.0,
    unit0_tilt_deg_cw: float = 18.0,
    unit1_tilt_deg_cw: float = 50.0,
    unit1_right_overflow_ratio: float = 0.10,
    scale: Optional[float] = None,
    hero_rgba: Optional[Any] = None,
    unit_rgba: Optional[Any] = None,
    **_compat: Any,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    N=3（N=2踏襲＋扇）:
    - 同尺。hero: TL左/BL下・四隅青枠内
    - unit0（中・ユーザー呼称unit1）: TL=枠上、Xは hero と unit1 の中間
    - unit1（右奥・ユーザー呼称unit2）: TRを枠右へ overflow 比だけはみ出し / BR=枠下
    - 傾き: hero左 / unit0 / unit1 右へ段階（既定 -14/+18/+45）
    可視閾値でのスケール最大化は fit 側。
    """
    canvas = int(canvas)
    ht = float(hero_tilt_deg_cw)
    if ht > 0:
        ht = -abs(ht)
    u0t = float(unit0_tilt_deg_cw)
    if u0t < 0:
        u0t = abs(u0t)
    u1t = float(unit1_tilt_deg_cw)
    if u1t < 0:
        u1t = abs(u1t)
    u1_overflow = max(0.0, min(0.20, float(unit1_right_overflow_ratio)))

    src = None
    if isinstance(unit_rgba, Image.Image):
        src = unit_rgba.convert("RGBA")
    elif isinstance(hero_rgba, Image.Image):
        src = hero_rgba.convert("RGBA")
    if src is None:
        # rgba無し: 旧頂円弧へフォールバック
        return propose_n3_fan_plans(
            canvas=canvas,
            product_w=product_w,
            product_h=product_h,
            hero_tilt_deg_cw=ht,
            unit0_tilt_deg_cw=u0t,
            unit1_tilt_deg_cw=u1t,
        )

    quad0 = ProductQuad.from_upright_rgba(src)
    lim = float(canvas - 1)

    def _quad_fits(q: ProductQuad, lim_v: float) -> bool:
        min_x, min_y, max_x, max_y = q.aabb()
        return (max_x - min_x) <= lim_v and (max_y - min_y) <= lim_v

    lo_c, hi_c = 0.05, 1.60
    for _ in range(28):
        mid = (lo_c + hi_c) * 0.5
        if _quad_fits(quad0.transformed(scale=mid, tilt_deg_cw=ht), lim):
            lo_c = mid
        else:
            hi_c = mid
    scale_cap = lo_c
    if scale is not None:
        scale_u = max(0.05, min(float(scale), scale_cap))
    else:
        scale_u = scale_cap

    h_quad = quad0.transformed(scale=scale_u, tilt_deg_cw=ht)
    u0_quad = quad0.transformed(scale=scale_u, tilt_deg_cw=u0t)
    u1_quad = quad0.transformed(scale=scale_u, tilt_deg_cw=u1t)

    # hero: TL左 / BL下 → 枠内クランプ
    h_paste_x = 0.0 - float(h_quad.tl[0])
    h_paste_y = float(canvas - 1) - float(h_quad.bl[1])
    h_min_x, h_min_y, h_max_x, h_max_y = h_quad.aabb()
    w_min_x = h_min_x + h_paste_x
    w_max_x = h_max_x + h_paste_x
    w_min_y = h_min_y + h_paste_y
    w_max_y = h_max_y + h_paste_y
    if w_min_x < 0:
        h_paste_x += -w_min_x
        w_min_x = 0.0
        w_max_x = h_max_x + h_paste_x
    if w_max_x > lim:
        h_paste_x -= w_max_x - lim
        w_min_x = h_min_x + h_paste_x
        w_max_x = h_max_x + h_paste_x
    if w_min_y < 0:
        h_paste_y += -w_min_y
        w_min_y = 0.0
        w_max_y = h_max_y + h_paste_y
    if w_max_y > lim:
        h_paste_y -= w_max_y - lim
        w_min_y = h_min_y + h_paste_y
        w_max_y = h_max_y + h_paste_y
    if w_min_x < -1e-6 or w_max_x > lim + 1e-6:
        h_paste_x = -h_min_x + (lim - (h_max_x - h_min_x)) * 0.5
    if w_min_y < -1e-6 or w_max_y > lim + 1e-6:
        h_paste_y = -h_min_y + (lim - (h_max_y - h_min_y)) * 0.5

    # unit1（右奥）: TRを枠右へ overflow 比だけはみ出し / BR下
    u1_tr_target_x = float(canvas - 1) + float(canvas) * u1_overflow
    u1_paste_x = u1_tr_target_x - float(u1_quad.tr[0])
    u1_paste_y = float(canvas - 1) - float(u1_quad.br[1])

    # unit0（中）: TL上、Xは hero と unit1 の中間
    h_cx = h_paste_x + 0.5 * (h_min_x + h_max_x)
    u1_min_x, u1_min_y, u1_max_x, u1_max_y = u1_quad.aabb()
    u1_cx = u1_paste_x + 0.5 * (u1_min_x + u1_max_x)
    target_cx = 0.5 * (h_cx + u1_cx)
    u0_min_x, u0_min_y, u0_max_x, u0_max_y = u0_quad.aabb()
    u0_cx_local = 0.5 * (u0_min_x + u0_max_x)
    u0_paste_x = target_cx - u0_cx_local
    u0_paste_y = 0.0 - float(u0_quad.tl[1])

    def _world(q: ProductQuad, px: float, py: float) -> Dict[str, Any]:
        return {
            "tl": [round(px + q.tl[0], 1), round(py + q.tl[1], 1)],
            "tr": [round(px + q.tr[0], 1), round(py + q.tr[1], 1)],
            "bl": [round(px + q.bl[0], 1), round(py + q.bl[1], 1)],
            "br": [round(px + q.br[0], 1), round(py + q.br[1], 1)],
        }

    hw = _world(h_quad, h_paste_x, h_paste_y)
    u0w = _world(u0_quad, u0_paste_x, u0_paste_y)
    u1w = _world(u1_quad, u1_paste_x, u1_paste_y)

    plans = [
        PortraitPlan(
            "unit",
            int(round(u1_paste_x)),
            int(round(u1_paste_y)),
            scale_u,
            z=0,
            rotation_deg=u1t,
            anchor="topleft",
        ),
        PortraitPlan(
            "unit",
            int(round(u0_paste_x)),
            int(round(u0_paste_y)),
            scale_u,
            z=10,
            rotation_deg=u0t,
            anchor="topleft",
        ),
        PortraitPlan(
            "hero",
            int(round(h_paste_x)),
            int(round(h_paste_y)),
            scale_u,
            z=1000,
            rotation_deg=ht,
            anchor="topleft",
        ),
    ]

    def _inside(pt: List[float], eps: float = 1.5) -> bool:
        return -eps <= float(pt[0]) <= lim + eps and -eps <= float(pt[1]) <= lim + eps

    hero_all_inside = all(_inside(hw[k]) for k in ("tl", "tr", "bl", "br"))
    logic_ja = [
        "①直立PNGで四隅定義→PIL一致変換",
        "②同尺。hero: TL左/BL下・青枠内",
        "③unit0(中/unit1): TL上・heroとunit1(右)の中間X",
        "④unit1(右奥/unit2): TRを右へoverflowはみ出し/BR下",
        f"⑤傾き扇 hero{ht:.0f} / unit0+{u0t:.0f} / unit1+{u1t:.0f}",
    ]
    meta: Dict[str, Any] = {
        "layoutFamily": "portrait_n3_upright_quad_fan",
        "pattern": "n3_upright_quad_fan",
        "status": "locked_pass",
        "lockedAt": "2026-08-07",
        "n": 3,
        "scale": scale_u,
        "scaleMax": scale_cap,
        "scaleCapHeroInside": scale_cap,
        "scaleHero": scale_u,
        "scaleUnit": scale_u,
        "scales": [scale_u, scale_u, scale_u],
        "sameScale": True,
        "tiltsDegCw": [round(ht, 2), round(u0t, 2), round(u1t, 2)],
        "heroTiltDegCw": ht,
        "unit0TiltDegCw": u0t,
        "unit1TiltDegCw": u1t,
        "unit1RightOverflowRatio": u1_overflow,
        "productQuadUpright": quad0.as_dict(),
        "heroQuadWorld": hw,
        "unit0QuadWorld": u0w,
        "unit1QuadWorld": u1w,
        "heroContacts": {
            "topLeftToFrameLeft": abs(float(hw["tl"][0]) - 0.0) < 2.0,
            "bottomLeftToFrameBottom": abs(float(hw["bl"][1]) - lim) < 2.0,
            "allCornersInsideFrame": hero_all_inside,
            "overflowPolicy": "forbidden",
        },
        "unit0Contacts": {
            "topLeftToFrameTop": abs(float(u0w["tl"][1]) - 0.0) < 2.0,
            "xMode": "mid_between_hero_unit1",
            "midTargetX": round(target_cx, 1),
            "overflowPolicy": "non_contact_corners_allowed",
        },
        "unit1Contacts": {
            "topRightOverflowRatio": u1_overflow,
            "topRightTargetX": round(u1_tr_target_x, 1),
            "topRightToFrameRight": abs(float(u1w["tr"][0]) - lim) < 2.0,
            "topRightPastFrameRight": float(u1w["tr"][0]) > lim + 1.0,
            "bottomRightToFrameBottom": abs(float(u1w["br"][1]) - lim) < 2.0,
            "overflowPolicy": "tr_right_overflow_allowed",
        },
        "roleMapJa": {
            "hero": "最前",
            "unit0": "中（ユーザーunit1）z=10",
            "unit1": "右奥（ユーザーunit2）z=0",
        },
        "logicJa": logic_ja,
        "ruleStepsJa": logic_ja,
        "anchor": "topleft",
        "zOrderJa": "unit1右奥→unit0中→hero最前",
        "proposalJa": " / ".join(logic_ja),
        "canvasStillSquare": True,
        "sampleAlignJa": "N2踏襲+扇傾き。unit2はTR右10%はみ出し",
        "productWhRef": [product_w, product_h],
    }
    return plans, meta


def propose_n4_upright_quad_fan_plans(
    *,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_tilt_deg_cw: float = -14.0,
    unit0_tilt_deg_cw: float = 18.0,
    unit1_tilt_deg_cw: float = 36.0,
    unit2_tilt_deg_cw: float = 60.0,
    unit2_right_overflow_ratio: float = 0.15,
    unit2_bottom_overflow_ratio: float = 0.0,
    slender_min_hw: float = 1.75,
    hero_height_fill_max_slender: float = 0.92,
    slender_unit1_top_mode: str = "perp_mid",
    scale: Optional[float] = None,
    hero_rgba: Optional[Any] = None,
    unit_rgba: Optional[Any] = None,
    **_compat: Any,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    N=4（N=3踏襲＋扇）:
    - 同尺。hero: TL左/BL下・四隅青枠内
    - unit0（ユーザーunit1）: TL上、X=hero〜unit2 の 1/3
    - unit1（ユーザーunit2）: 幅広=直角線スライドでTL枠上／
      細長(H/W≥slender_min_hw)=unit0〜unit2上辺中点弦の中点Mの直角線上（左右均等）
    - unit2（ユーザーunit3）: TR右へ overflow / BRは同縦ラインで枠下辺に接触（下はみ出しなし）
    - 傾き: -14 / +18 / +36 / +60
    - 細長時は hero 高さ上限で同尺を抑えつつ画面埋め（既定0.92）
    """
    canvas = int(canvas)
    ht = float(hero_tilt_deg_cw)
    if ht > 0:
        ht = -abs(ht)
    tilts = []
    for t in (unit0_tilt_deg_cw, unit1_tilt_deg_cw, unit2_tilt_deg_cw):
        tt = float(t)
        tilts.append(abs(tt) if tt < 0 else tt)
    u0t, u1t, u2t = tilts
    u2_overflow = max(0.0, min(0.25, float(unit2_right_overflow_ratio)))
    u2_bottom_overflow = max(0.0, min(0.25, float(unit2_bottom_overflow_ratio)))
    slender_hw = max(1.0, float(slender_min_hw))
    slender_h_fill = max(0.45, min(0.98, float(hero_height_fill_max_slender)))
    u1_mode_cfg = str(slender_unit1_top_mode or "perp_mid").strip().lower()
    # 旧名 top_arc も左右均等の直角線配置へ統合
    if u1_mode_cfg in ("top_arc", "arc", "perp", "perp_midpoint"):
        u1_mode_cfg = "perp_mid"

    src = None
    if isinstance(unit_rgba, Image.Image):
        src = unit_rgba.convert("RGBA")
    elif isinstance(hero_rgba, Image.Image):
        src = hero_rgba.convert("RGBA")
    if src is None:
        return propose_n4_fan_plans(
            canvas=canvas,
            product_w=product_w,
            product_h=product_h,
            hero_tilt_deg_cw=ht,
            unit0_tilt_deg_cw=u0t,
            unit1_tilt_deg_cw=u1t,
            unit2_tilt_deg_cw=u2t,
            hero_rgba=None,
            unit_rgba=None,
        )

    quad0 = ProductQuad.from_upright_rgba(src)
    lim = float(canvas - 1)
    src_bb = src.split()[-1].getbbox() or (0, 0, src.width, src.height)
    src_w = max(1, int(src_bb[2] - src_bb[0]))
    src_h = max(1, int(src_bb[3] - src_bb[1]))
    product_hw = float(src_h) / float(src_w)
    is_slender = product_hw + 1e-9 >= slender_hw
    unit1_top_mode = u1_mode_cfg if is_slender else "frame_top"
    if unit1_top_mode not in ("perp_mid", "frame_top"):
        unit1_top_mode = "perp_mid" if is_slender else "frame_top"

    def _quad_fits(q: ProductQuad, lim_v: float) -> bool:
        min_x, min_y, max_x, max_y = q.aabb()
        return (max_x - min_x) <= lim_v and (max_y - min_y) <= lim_v

    lo_c, hi_c = 0.05, 1.60
    for _ in range(28):
        mid = (lo_c + hi_c) * 0.5
        if _quad_fits(quad0.transformed(scale=mid, tilt_deg_cw=ht), lim):
            lo_c = mid
        else:
            hi_c = mid
    scale_cap = lo_c
    # 細長: 高さ埋め上限で同尺を抑え、扇の逃げ場を確保
    scale_cap_height = scale_cap
    if is_slender:
        target_h = float(canvas) * slender_h_fill
        lo_h, hi_h = 0.05, scale_cap
        for _ in range(28):
            mid = (lo_h + hi_h) * 0.5
            _mn_x, mn_y, _mx_x, mx_y = quad0.transformed(
                scale=mid, tilt_deg_cw=ht
            ).aabb()
            if (mx_y - mn_y) <= target_h + 1e-6:
                lo_h = mid
            else:
                hi_h = mid
        scale_cap_height = lo_h
        scale_cap = min(scale_cap, scale_cap_height)
    if scale is not None:
        scale_u = max(0.05, min(float(scale), scale_cap))
    else:
        scale_u = scale_cap

    h_quad = quad0.transformed(scale=scale_u, tilt_deg_cw=ht)
    u0_quad = quad0.transformed(scale=scale_u, tilt_deg_cw=u0t)
    u1_quad = quad0.transformed(scale=scale_u, tilt_deg_cw=u1t)
    u2_quad = quad0.transformed(scale=scale_u, tilt_deg_cw=u2t)

    # hero
    h_paste_x = 0.0 - float(h_quad.tl[0])
    h_paste_y = float(canvas - 1) - float(h_quad.bl[1])
    h_min_x, h_min_y, h_max_x, h_max_y = h_quad.aabb()
    w_min_x = h_min_x + h_paste_x
    w_max_x = h_max_x + h_paste_x
    w_min_y = h_min_y + h_paste_y
    w_max_y = h_max_y + h_paste_y
    if w_min_x < 0:
        h_paste_x += -w_min_x
        w_min_x = 0.0
        w_max_x = h_max_x + h_paste_x
    if w_max_x > lim:
        h_paste_x -= w_max_x - lim
        w_min_x = h_min_x + h_paste_x
        w_max_x = h_max_x + h_paste_x
    if w_min_y < 0:
        h_paste_y += -w_min_y
        w_min_y = 0.0
        w_max_y = h_max_y + h_paste_y
    if w_max_y > lim:
        h_paste_y -= w_max_y - lim
        w_min_y = h_min_y + h_paste_y
        w_max_y = h_max_y + h_paste_y
    if w_min_x < -1e-6 or w_max_x > lim + 1e-6:
        h_paste_x = -h_min_x + (lim - (h_max_x - h_min_x)) * 0.5
    if w_min_y < -1e-6 or w_max_y > lim + 1e-6:
        h_paste_y = -h_min_y + (lim - (h_max_y - h_min_y)) * 0.5

    # unit2（右奥＝ユーザーunit3）:
    # TR右overflow（X決定）／BRは同じ縦ラインのまま枠下辺に接触（bottomOverflow=0で確定）
    u2_tr_target_x = float(canvas - 1) + float(canvas) * u2_overflow
    u2_br_target_y = float(canvas - 1) + float(canvas) * u2_bottom_overflow
    u2_paste_x = u2_tr_target_x - float(u2_quad.tr[0])
    u2_paste_y = u2_br_target_y - float(u2_quad.br[1])
    u2_br_x_line = u2_paste_x + float(u2_quad.br[0])

    h_cx = h_paste_x + 0.5 * (h_min_x + h_max_x)
    u2_min_x, _, u2_max_x, _ = u2_quad.aabb()
    u2_cx = u2_paste_x + 0.5 * (u2_min_x + u2_max_x)

    def _place_tl_top_at_frac(q: ProductQuad, frac: float) -> Tuple[float, float, float]:
        target_cx = h_cx + (u2_cx - h_cx) * float(frac)
        mn_x, _, mx_x, _ = q.aabb()
        cx_local = 0.5 * (mn_x + mx_x)
        px = target_cx - cx_local
        py = 0.0 - float(q.tl[1])
        return px, py, target_cx

    u0_paste_x, u0_paste_y, u0_target = _place_tl_top_at_frac(u0_quad, 1.0 / 3.0)

    # unit1（ユーザーunit2）
    def _top_mid_local(q: ProductQuad) -> Tuple[float, float]:
        return (
            0.5 * (float(q.tl[0]) + float(q.tr[0])),
            0.5 * (float(q.tl[1]) + float(q.tr[1])),
        )

    def _top_mid_world(q: ProductQuad, px: float, py: float) -> Tuple[float, float]:
        lx, ly = _top_mid_local(q)
        return (px + lx, py + ly)

    def _bottom_mid_local(q: ProductQuad) -> Tuple[float, float]:
        return (
            0.5 * (float(q.bl[0]) + float(q.br[0])),
            0.5 * (float(q.bl[1]) + float(q.br[1])),
        )

    u0_top_mid_w = _top_mid_world(u0_quad, u0_paste_x, u0_paste_y)
    u2_top_mid_w = _top_mid_world(u2_quad, u2_paste_x, u2_paste_y)
    ax, ay = float(u0_top_mid_w[0]), float(u0_top_mid_w[1])
    cx, cy = float(u2_top_mid_w[0]), float(u2_top_mid_w[1])
    mx, my = 0.5 * (ax + cx), 0.5 * (ay + cy)
    vx, vy = cx - ax, cy - ay
    vlen = (vx * vx + vy * vy) ** 0.5
    if vlen < 1e-6:
        nx, ny = 0.0, 1.0
    else:
        # A→C に対する直角（右回り法線）。キャンバス内側（概ね下方向）へ向ける
        nx, ny = -vy / vlen, vx / vlen
        canvas_cx = 0.5 * float(canvas - 1)
        canvas_cy = 0.5 * float(canvas - 1)
        if nx * (canvas_cx - mx) + ny * (canvas_cy - my) < 0:
            nx, ny = -nx, -ny

    u1_top_mid_local = _top_mid_local(u1_quad)
    u1_bot_mid_local = _bottom_mid_local(u1_quad)
    tl_lx, tl_ly = float(u1_quad.tl[0]), float(u1_quad.tl[1])
    tmx, tmy = float(u1_top_mid_local[0]), float(u1_top_mid_local[1])
    u1_tl_local = (float(u1_quad.tl[0]), float(u1_quad.tl[1]))
    u1_tr_local = (float(u1_quad.tr[0]), float(u1_quad.tr[1]))
    u0_tr_local = (float(u0_quad.tr[0]), float(u0_quad.tr[1]))
    u2_tl_local = (float(u2_quad.tl[0]), float(u2_quad.tl[1]))

    tl_frame_top = False
    t_slide = 0.0
    arc_meta: Dict[str, Any] = {"enabled": False}
    if unit1_top_mode == "perp_mid":
        # 細長: A=unit0上辺中点 / C=unit2上辺中点 の弦。
        # 中点Mの直角線上に unit1 上辺中点を置く → 左右の白い隙間が均等。
        # スライドは画面上方向（Y減）＝弦より上へ。下へ沈む谷を防ぐ（扇の上凸）。
        chord_len = max(vlen, 1.0)
        # 法線を「上方向優先」に揃える（キャンバス内側＝下向きだと谷になる）
        if ny > 0:
            nx, ny = -nx, -ny
        elif abs(ny) < 1e-6 and nx * (0.5 * float(canvas - 1) - mx) > 0:
            # ほぼ水平: 枠中央から離れる側へ
            nx, ny = -nx, -ny
        # 上辺扇: B を A/C の高い方付近まで上げる（垂線上・左右均等維持）
        # わずかに枠上へ余裕（完全貼付は避け、谷＝弦より下は禁止）
        high_y = min(float(ay), float(cy))
        target_y = max(8.0, high_y + 6.0)  # 高い端よりわずかに下＝なめらかな上凸
        # 弦中点よりは必ず上
        target_y = min(target_y, float(my) - 4.0)
        target_y = max(8.0, target_y)
        if abs(ny) >= 1e-6:
            t_slide = (target_y - my) / ny
            # 上方向のみ（ny<0 前提）。下向きなら符号反転済みのはず
            if t_slide < 0:
                t_slide = abs(t_slide)
                nx, ny = -nx, -ny
                t_slide = (target_y - my) / ny
            bx = mx + t_slide * nx
            by = my + t_slide * ny
        else:
            t_slide = 0.0
            bx, by = mx, target_y
        if by < 8.0:
            if abs(ny) >= 1e-6:
                t_slide = (8.0 - my) / ny
                bx = mx + t_slide * nx
                by = my + t_slide * ny
            else:
                by = 8.0
                bx = mx
        # 谷禁止: 弦中点より下なら M まで戻す
        if by > my + 1e-6:
            t_slide = 0.0
            bx, by = mx, my
        u1_paste_x = bx - tmx
        u1_paste_y = by - tmy
        u1_top_mid_target = (bx, by)
        tl_frame_top = abs((u1_paste_y + tl_ly) - 0.0) < 2.0
        # 左右均等の検算（弦方向への射影距離）
        def _proj_along_chord(px: float, py: float) -> float:
            if vlen < 1e-6:
                return 0.0
            return ((px - ax) * vx + (py - ay) * vy) / (vlen * vlen)

        b_s = _proj_along_chord(bx, by)
        gap_l_s = b_s - 0.0
        gap_r_s = 1.0 - b_s
        arc_meta = {
            "enabled": True,
            "mode": "perp_mid_up_bulge",
            "chordA": [round(ax, 1), round(ay, 1)],
            "chordC": [round(cx, 1), round(cy, 1)],
            "midM": [round(mx, 1), round(my, 1)],
            "perpN": [round(nx, 4), round(ny, 4)],
            "perpSlideT": round(float(t_slide), 3),
            "targetTopMid": [round(bx, 1), round(by, 1)],
            "chordParamB": round(float(b_s), 4),
            "bulgeUp": True,
            "sideGapBalance": {
                "left": round(float(gap_l_s), 4),
                "right": round(float(gap_r_s), 4),
                "absDiff": round(abs(float(gap_l_s) - float(gap_r_s)), 4),
            },
            "noteJa": "A–C中点Mの直角線上・上方向bulge（扇の上凸）。左右均等",
        }
    else:
        # 幅広（豚汁）: 直角線スライドで TL を枠上に接触
        tl_frame_top = True
        if abs(ny) >= 1e-6:
            t_slide = (tmy - tl_ly - my) / ny
            bx = mx + t_slide * nx
            by = my + t_slide * ny
        else:
            t_slide = 0.0
            bx, by = mx, my
            tl_frame_top = False
        u1_top_mid_target = (bx, by)
        u1_paste_x = bx - tmx
        u1_paste_y = by - tmy
        if abs(ny) < 1e-6:
            u1_paste_y = 0.0 - tl_ly
            u1_paste_x = mx - tmx
            bx = u1_paste_x + tmx
            by = u1_paste_y + tmy
            u1_top_mid_target = (bx, by)
            tl_frame_top = True

    u1_tl_w = (u1_paste_x + tl_lx, u1_paste_y + tl_ly)
    u1_bot_mid_w = (
        u1_paste_x + u1_bot_mid_local[0],
        u1_paste_y + u1_bot_mid_local[1],
    )
    # 直角線の表示用端点（上下に伸ばす）
    perp_half = max(180.0, 0.35 * float(canvas))
    perp_p0 = (mx - nx * perp_half, my - ny * perp_half)
    perp_p1 = (mx + nx * perp_half, my + ny * perp_half)

    def _world(q: ProductQuad, px: float, py: float) -> Dict[str, Any]:
        return {
            "tl": [round(px + q.tl[0], 1), round(py + q.tl[1], 1)],
            "tr": [round(px + q.tr[0], 1), round(py + q.tr[1], 1)],
            "bl": [round(px + q.bl[0], 1), round(py + q.bl[1], 1)],
            "br": [round(px + q.br[0], 1), round(py + q.br[1], 1)],
        }

    hw = _world(h_quad, h_paste_x, h_paste_y)
    u0w = _world(u0_quad, u0_paste_x, u0_paste_y)
    u1w = _world(u1_quad, u1_paste_x, u1_paste_y)
    u2w = _world(u2_quad, u2_paste_x, u2_paste_y)

    plans = [
        PortraitPlan(
            "unit", int(round(u2_paste_x)), int(round(u2_paste_y)),
            scale_u, z=0, rotation_deg=u2t, anchor="topleft",
        ),
        PortraitPlan(
            "unit", int(round(u1_paste_x)), int(round(u1_paste_y)),
            scale_u, z=10, rotation_deg=u1t, anchor="topleft",
        ),
        PortraitPlan(
            "unit", int(round(u0_paste_x)), int(round(u0_paste_y)),
            scale_u, z=20, rotation_deg=u0t, anchor="topleft",
        ),
        PortraitPlan(
            "hero", int(round(h_paste_x)), int(round(h_paste_y)),
            scale_u, z=1000, rotation_deg=ht, anchor="topleft",
        ),
    ]

    def _inside(pt: List[float], eps: float = 1.5) -> bool:
        return -eps <= float(pt[0]) <= lim + eps and -eps <= float(pt[1]) <= lim + eps

    hero_all_inside = all(_inside(hw[k]) for k in ("tl", "tr", "bl", "br"))
    if unit1_top_mode == "perp_mid":
        u1_logic = "④unit1(unit2): A↔C中点Mの直角線上・上方向bulge（細長・扇の上凸・左右均等）"
    else:
        u1_logic = "④unit1(unit2): A↔C中点Mの直角線上をスライド・TLを枠上に接触"
    logic_ja = [
        "①直立PNG四隅→PIL一致変換・同尺",
        "②hero: TL左/BL下・青枠内"
        + (f"・細長高さ≤{slender_h_fill:.0%}" if is_slender else ""),
        "③unit0(unit1): TL上・hero〜unit2の1/3X",
        u1_logic,
        "⑤unit2(unit3): TR右overflow／BRは同縦ラインで枠下辺接触（下はみ出しなし）",
        f"⑥傾き {ht:.0f}/+{u0t:.0f}/+{u1t:.0f}/+{u2t:.0f}",
        f"⑦形状判定 H/W={product_hw:.3f} "
        f"{'細長' if is_slender else '幅広'}(閾値{slender_hw:.2f})",
    ]
    meta: Dict[str, Any] = {
        "layoutFamily": "portrait_n4_upright_quad_fan",
        "pattern": "n4_upright_quad_fan_slender"
        if is_slender
        else "n4_upright_quad_fan",
        "status": "locked_pass",
        "n": 4,
        "scale": scale_u,
        "scaleMax": scale_cap,
        "scaleCapHeroInside": scale_cap,
        "scaleCapHeightSlender": round(float(scale_cap_height), 5)
        if is_slender
        else None,
        "scaleHero": scale_u,
        "scaleUnit": scale_u,
        "scales": [scale_u, scale_u, scale_u, scale_u],
        "sameScale": True,
        "productHw": round(product_hw, 4),
        "slenderMinHw": slender_hw,
        "isSlender": bool(is_slender),
        "heroHeightFillMaxSlender": slender_h_fill if is_slender else None,
        "unit1TopMode": unit1_top_mode,
        "tiltsDegCw": [round(ht, 2), round(u0t, 2), round(u1t, 2), round(u2t, 2)],
        "heroTiltDegCw": ht,
        "unit0TiltDegCw": u0t,
        "unit1TiltDegCw": u1t,
        "unit2TiltDegCw": u2t,
        "unit2RightOverflowRatio": u2_overflow,
        "unit2BottomOverflowRatio": u2_bottom_overflow,
        "heroQuadWorld": hw,
        "unit0QuadWorld": u0w,
        "unit1QuadWorld": u1w,
        "unit2QuadWorld": u2w,
        "heroContacts": {
            "topLeftToFrameLeft": abs(float(hw["tl"][0]) - 0.0) < 2.0,
            "bottomLeftToFrameBottom": abs(float(hw["bl"][1]) - lim) < 2.0,
            "allCornersInsideFrame": hero_all_inside,
            "overflowPolicy": "forbidden",
        },
        "unit0Contacts": {
            "topLeftToFrameTop": abs(float(u0w["tl"][1]) - 0.0) < 2.0,
            "xMode": "frac_1_3_hero_to_unit2",
            "targetX": round(u0_target, 1),
        },
        "unit1Contacts": {
            "xMode": (
                "perp_mid_chord_ac"
                if unit1_top_mode == "perp_mid"
                else "perp_slide_tl_to_frame_top"
            ),
            "chordA": [round(ax, 1), round(ay, 1)],
            "chordC": [round(cx, 1), round(cy, 1)],
            "midM": [round(mx, 1), round(my, 1)],
            "perpN": [round(nx, 4), round(ny, 4)],
            "perpSlideT": round(float(t_slide), 3),
            "perpSeg": [
                [round(perp_p0[0], 1), round(perp_p0[1], 1)],
                [round(perp_p1[0], 1), round(perp_p1[1], 1)],
            ],
            "topArc": arc_meta,
            "targetTopMid": [round(u1_top_mid_target[0], 1), round(u1_top_mid_target[1], 1)],
            "unit0TopMidWorld": [round(ax, 1), round(ay, 1)],
            "unit2TopMidWorld": [round(cx, 1), round(cy, 1)],
            "actualTopMidWorld": [
                round(u1_paste_x + tmx, 1),
                round(u1_paste_y + tmy, 1),
            ],
            "actualBottomMidWorld": [round(u1_bot_mid_w[0], 1), round(u1_bot_mid_w[1], 1)],
            "topLeftWorld": [round(u1_tl_w[0], 1), round(u1_tl_w[1], 1)],
            "topLeftToFrameTop": abs(float(u1_tl_w[1]) - 0.0) < 2.0 and tl_frame_top,
            "noteJa": (
                "細長: A–C中点Mの直角線上にunit1上辺中点を上方向bulge（扇の上凸・左右均等）"
                if unit1_top_mode == "perp_mid"
                else "A↔C中点Mの直角線上をスライドし、unit2.TLを枠上へ"
            ),
        },
        "unit2Contacts": {
            "topRightOverflowRatio": u2_overflow,
            "topRightTargetX": round(u2_tr_target_x, 1),
            "topRightPastFrameRight": float(u2w["tr"][0]) > lim + 1.0,
            "bottomRightOverflowRatio": u2_bottom_overflow,
            "bottomRightTargetY": round(u2_br_target_y, 1),
            "bottomRightToFrameBottom": abs(float(u2w["br"][1]) - lim) < 2.0,
            "bottomRightPastFrameBottom": float(u2w["br"][1]) > lim + 1.0,
            "bottomRightVerticalLineX": round(u2_br_x_line, 1),
            "bottomRightWorld": [
                round(float(u2w["br"][0]), 1),
                round(float(u2w["br"][1]), 1),
            ],
            "noteJa": "BRはTR由来の縦ラインを維持し、枠下辺に接触（下はみ出しなし・確定）",
        },
        "roleMapJa": {
            "hero": "最前",
            "unit0": "ユーザーunit1 z=20",
            "unit1": "ユーザーunit2 z=10",
            "unit2": "ユーザーunit3（旧N3 unit2）z=0",
        },
        "logicJa": logic_ja,
        "ruleStepsJa": logic_ja,
        "anchor": "topleft",
        "zOrderJa": "unit2→unit1→unit0→hero",
        "proposalJa": " / ".join(logic_ja),
        "canvasStillSquare": True,
        "sampleAlignJa": "N3踏襲。可視閾値はfit側",
        "productWhRef": [product_w, product_h],
    }
    return plans, meta


def _shift_portrait_plans_xy_(
    plans: List[PortraitPlan], *, dx: float, dy: float
) -> List[PortraitPlan]:
    """全プランを剛体平行移動（相対配置・円弧関係を維持）。"""
    if abs(dx) < 1e-9 and abs(dy) < 1e-9:
        return list(plans)
    out: List[PortraitPlan] = []
    for p in plans:
        kw: Dict[str, Any] = {
            "x": int(round(float(p.x) + dx)),
            "y": int(round(float(p.y) + dy)),
        }
        if p.top_x is not None:
            kw["top_x"] = int(round(float(p.top_x) + dx))
        if p.top_y is not None:
            kw["top_y"] = int(round(float(p.top_y) + dy))
        if p.foot_x is not None:
            kw["foot_x"] = int(round(float(p.foot_x) + dx))
        if p.foot_y is not None:
            kw["foot_y"] = int(round(float(p.foot_y) + dy))
        out.append(replace(p, **kw))
    return out


def pin_plans_hero_aabb_to_frame_bottom(
    plans: List[PortraitPlan],
    *,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_rgba: Optional[Image.Image] = None,
    unit_rgba: Optional[Image.Image] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    hero の不透明 AABB 下辺をキャンバス外枠下辺（y=canvas-1）に接地。
    扇全体を剛体平行移動するため、頂円弧の相対関係は崩れない。
    N=4 細長（n4_legacy）の hero 下辺接触と同型。
    """
    canvas = int(canvas)
    hero_p = next((p for p in plans if str(p.role) == "hero"), None)
    if hero_p is None:
        return list(plans), {"pinned": False, "reason": "no_hero"}
    src = hero_rgba if isinstance(hero_rgba, Image.Image) else unit_rgba
    layer = _prepare_scaled_rotated_layer(
        product_w=product_w,
        product_h=product_h,
        scale=float(hero_p.scale),
        tilt_deg_cw=float(hero_p.rotation_deg),
        rgba=src if isinstance(src, Image.Image) else None,
    )
    _l, _t, _r, hb = measure_aabb4(layer)
    anc = str(hero_p.anchor or "topleft")
    if anc == "top" and hero_p.top_x is not None and hero_p.top_y is not None:
        tx_l, ty_l = measure_top_xy(layer)
        oy = float(hero_p.top_y) - float(ty_l)
    elif anc == "foot" and hero_p.foot_x is not None and hero_p.foot_y is not None:
        _fx, fy_l = measure_foot_xy(layer)
        oy = float(hero_p.foot_y) - float(fy_l)
    else:
        oy = float(hero_p.y)
    world_b = float(oy) + float(hb)
    target_b = float(canvas - 1)
    dy = target_b - world_b
    if abs(dy) < 0.5:
        return list(plans), {
            "pinned": True,
            "dy": 0.0,
            "heroWorldBottom": round(world_b, 2),
            "targetBottom": target_b,
            "already": True,
        }
    shifted = _shift_portrait_plans_xy_(plans, dx=0.0, dy=dy)
    return shifted, {
        "pinned": True,
        "dy": round(dy, 2),
        "heroWorldBottomBefore": round(world_b, 2),
        "heroWorldBottomAfter": round(world_b + dy, 2),
        "targetBottom": target_b,
        "ruleJa": "hero AABB下辺＝枠下辺（扇全体を剛体平行移動／N=4細長と同型）",
    }


def _plan_paste_xy_(
    p: PortraitPlan,
    layer: Image.Image,
) -> Tuple[float, float]:
    anc = str(p.anchor or "topleft")
    if anc == "top" and p.top_x is not None and p.top_y is not None:
        tx_l, ty_l = measure_top_xy(layer)
        return float(p.top_x) - float(tx_l), float(p.top_y) - float(ty_l)
    if anc == "foot" and p.foot_x is not None and p.foot_y is not None:
        fx_l, fy_l = measure_foot_xy(layer)
        return float(p.foot_x) - float(fx_l), float(p.foot_y) - float(fy_l)
    return float(p.x), float(p.y)


def plan_world_aabb(
    p: PortraitPlan,
    *,
    product_w: int,
    product_h: int,
    rgba: Optional[Image.Image] = None,
) -> Dict[str, float]:
    """プラン貼付後の不透明 AABB（ワールド座標）。"""
    layer = _prepare_scaled_rotated_layer(
        product_w=product_w,
        product_h=product_h,
        scale=float(p.scale),
        tilt_deg_cw=float(p.rotation_deg),
        rgba=rgba,
    )
    l, t, r, b = measure_aabb4(layer)
    ox, oy = _plan_paste_xy_(p, layer)
    left = ox + float(l)
    top = oy + float(t)
    right = ox + float(r)
    bottom = oy + float(b)
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "cx": 0.5 * (left + right),
        "cy": 0.5 * (top + bottom),
        "w": right - left + 1.0,
        "h": bottom - top + 1.0,
    }


def _n3_hero_u0_u1_(
    plans: List[PortraitPlan],
) -> Tuple[PortraitPlan, PortraitPlan, PortraitPlan]:
    hero = next(p for p in plans if str(p.role) == "hero")
    units = sorted(
        [p for p in plans if str(p.role) == "unit"],
        key=lambda p: float(p.rotation_deg),
    )
    if len(units) < 2:
        raise ValueError("N=3 plans require hero + 2 units")
    return hero, units[0], units[1]


def fit_plans_hero_inside_frame(
    plans: List[PortraitPlan],
    *,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_rgba: Optional[Image.Image] = None,
    unit_rgba: Optional[Image.Image] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    hero AABB を枠内へ（unitのはみ出しは許容）。
    1) 左はみ出し → 扇全体を +X
    2) なお右/上はみ出し → inside=False（呼び出し側で縮小再提案）
    下辺接地は維持（Yは動かさない）。
    """
    canvas = int(canvas)
    src = hero_rgba if isinstance(hero_rgba, Image.Image) else unit_rgba
    hero_p = next((p for p in plans if str(p.role) == "hero"), None)
    if hero_p is None:
        return list(plans), {"inside": False, "reason": "no_hero"}
    box = plan_world_aabb(
        hero_p, product_w=product_w, product_h=product_h, rgba=src
    )
    dx = 0.0
    if box["left"] < -1e-6:
        # 整数座標シフトなので切り上げ（左辺≥0を保証）
        dx = float(math.ceil(-box["left"]))
    plans2 = _shift_portrait_plans_xy_(plans, dx=dx, dy=0.0) if abs(dx) >= 0.5 else list(plans)
    box2 = plan_world_aabb(
        next(p for p in plans2 if str(p.role) == "hero"),
        product_w=product_w,
        product_h=product_h,
        rgba=src,
    )
    inside = (
        box2["left"] >= -0.5
        and box2["top"] >= -0.5
        and box2["right"] <= float(canvas - 1) + 0.5
        and box2["bottom"] <= float(canvas - 1) + 0.5
    )
    return plans2, {
        "inside": bool(inside),
        "dx": round(dx, 2),
        "heroAabb": {k: round(float(v), 1) for k, v in box2.items()},
        "overflowLeft": round(min(0.0, box2["left"]), 1),
        "overflowTop": round(min(0.0, box2["top"]), 1),
        "overflowRight": round(max(0.0, box2["right"] - (canvas - 1)), 1),
        "overflowBottom": round(max(0.0, box2["bottom"] - (canvas - 1)), 1),
        "ruleJa": "hero AABB枠内（左寄せ剛体平行移動／不足時は縮小再提案）",
    }


def nudge_n3_unit0_aabb_center_mid_x(
    plans: List[PortraitPlan],
    *,
    product_w: int,
    product_h: int,
    hero_rgba: Optional[Image.Image] = None,
    unit_rgba: Optional[Image.Image] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    N=3: unit0（ユーザーunit1）の AABB 中心Xを
    hero と unit1（ユーザーunit2）の AABB 中心Xの中点へ合わせる。
    頂X中点だと傾き差で胴体が unit2 寄りに見えるため。
    円弧の頂Yは維持（Xのみ調整＝円周から外れるが「見た目中点」優先）。
    """
    src_h = hero_rgba if isinstance(hero_rgba, Image.Image) else unit_rgba
    src_u = unit_rgba if isinstance(unit_rgba, Image.Image) else hero_rgba
    hero_p, u0_p, u1_p = _n3_hero_u0_u1_(plans)
    hb = plan_world_aabb(hero_p, product_w=product_w, product_h=product_h, rgba=src_h)
    u1b = plan_world_aabb(u1_p, product_w=product_w, product_h=product_h, rgba=src_u)
    target_cx = 0.5 * (hb["cx"] + u1b["cx"])

    layer = _prepare_scaled_rotated_layer(
        product_w=product_w,
        product_h=product_h,
        scale=float(u0_p.scale),
        tilt_deg_cw=float(u0_p.rotation_deg),
        rgba=src_u,
    )
    l, t, r, b = measure_aabb4(layer)
    tx_l, ty_l = measure_top_xy(layer)
    # center_x = (top_x - tx_l) + (l+r)/2  → top_x = target_cx - (l+r)/2 + tx_l
    new_top_x = target_cx - 0.5 * (float(l) + float(r)) + float(tx_l)
    old_top_x = float(u0_p.top_x if u0_p.top_x is not None else u0_p.x)
    u0_new = replace(
        u0_p,
        top_x=int(round(new_top_x)),
        x=int(round(float(u0_p.x) + (new_top_x - old_top_x))),
    )
    out: List[PortraitPlan] = []
    for p in plans:
        if (
            str(p.role) == "unit"
            and abs(float(p.rotation_deg) - float(u0_p.rotation_deg)) < 1e-6
            and int(p.z) == int(u0_p.z)
        ):
            out.append(u0_new)
        else:
            out.append(p)
    u0b = plan_world_aabb(u0_new, product_w=product_w, product_h=product_h, rgba=src_u)
    span = max(1e-6, u1b["cx"] - hb["cx"])
    frac = (u0b["cx"] - hb["cx"]) / span
    return out, {
        "applied": True,
        "targetCx": round(target_cx, 1),
        "unit0Cx": round(u0b["cx"], 1),
        "heroCx": round(hb["cx"], 1),
        "unit1Cx": round(u1b["cx"], 1),
        "fracAlongHeroUnit1": round(frac, 3),
        "topXBefore": round(old_top_x, 1),
        "topXAfter": int(round(new_top_x)),
        "ruleJa": "unit0 AABB中心X＝heroとunit1のAABB中心中点（見た目バランス）",
    }


def _peak_world_of_plan_(
    p: PortraitPlan,
    *,
    product_w: int,
    product_h: int,
    rgba: Optional[Image.Image],
) -> Tuple[float, float]:
    layer = _prepare_scaled_rotated_layer(
        product_w=product_w,
        product_h=product_h,
        scale=float(p.scale),
        tilt_deg_cw=float(p.rotation_deg),
        rgba=rgba,
    )
    tx, ty = measure_top_xy(layer)
    ox, oy = _plan_paste_xy_(p, layer)
    return ox + float(tx), oy + float(ty)


def _top_edge_lr_mid_local_(
    layer: Image.Image,
    *,
    band_ratio: float = 0.035,
    band_px_min: int = 24,
    band_px_max: int = 56,
    alpha_thresh: int = 96,
) -> Tuple[Tuple[float, float], Tuple[float, float], Tuple[float, float]]:
    """
    上辺の左端・右端・左右中点（ローカル座標）。

    定義（正）: 回転後シルエット最上端から薄い上部帯（キャップ上辺）内の
      不透明画素の最左=TL / 最右=TR / MID=((TLx+TRx)/2,(TLy+TRy)/2)

    帯は高さの約3.5%（24〜56px）。広すぎると胴体左右を掴み、
    『上辺』にならない。AABB全幅も使わない。
    """
    im = layer.convert("RGBA")
    a = im.getchannel("A")
    bb = a.getbbox()
    if not bb:
        return (0.0, 0.0), (0.0, 0.0), (0.0, 0.0)
    l0, t0, r0, b0 = bb
    h = max(1, b0 - t0)
    band = int(round(h * float(band_ratio)))
    band = max(int(band_px_min), min(int(band_px_max), band))
    y1 = min(im.height, t0 + band)
    ap = a.load()
    xs: List[Tuple[int, int]] = []
    for y in range(t0, y1):
        for x in range(l0, r0):
            if ap[x, y] > alpha_thresh:
                xs.append((x, y))
    if not xs:
        mid = (0.5 * (l0 + r0 - 1), float(t0))
        return (float(l0), float(t0)), (float(r0 - 1), float(t0)), mid

    # 上部帯の最左・最右（同xならより上）
    min_x = min(p[0] for p in xs)
    max_x = max(p[0] for p in xs)
    tl_y = min(y for x, y in xs if x == min_x)
    tr_y = min(y for x, y in xs if x == max_x)
    tl_p = (float(min_x), float(tl_y))
    tr_p = (float(max_x), float(tr_y))
    mid = (0.5 * (tl_p[0] + tr_p[0]), 0.5 * (tl_p[1] + tr_p[1]))
    return tl_p, tr_p, mid


def _circle_from_three_xy_(
    a: Tuple[float, float],
    b: Tuple[float, float],
    c: Tuple[float, float],
) -> Optional[Tuple[float, float, float]]:
    ax, ay = a
    bx, by = b
    cx, cy = c
    d = 2.0 * (ax * (by - cy) + bx * (cy - ay) + cx * (ay - by))
    if abs(d) < 1e-6:
        return None
    a2 = ax * ax + ay * ay
    b2 = bx * bx + by * by
    c2 = cx * cx + cy * cy
    ox = (a2 * (by - cy) + b2 * (cy - ay) + c2 * (ay - by)) / d
    oy = (a2 * (cx - bx) + b2 * (ax - cx) + c2 * (bx - ax)) / d
    r = math.hypot(ax - ox, ay - oy)
    if r < 10:
        return None
    return float(ox), float(oy), float(r)


def _product_quad_top_lr_mid_world_(
    p: PortraitPlan,
    *,
    product_w: int,
    product_h: int,
    rgba: Image.Image,
) -> Tuple[Tuple[float, float], Tuple[float, float], Tuple[float, float]]:
    """
    unit1 四隅（ProductQuad）を結ぶ上辺 TL–TR の両端と中点 Q。

    Q = ((TLx+TRx)/2, (TLy+TRy)/2) … AABBの水平上辺ではない。
    """
    q0 = ProductQuad.from_upright_rgba(rgba)
    qt = q0.transformed(scale=float(p.scale), tilt_deg_cw=float(p.rotation_deg))
    layer = _prepare_scaled_rotated_layer(
        product_w=product_w,
        product_h=product_h,
        scale=float(p.scale),
        tilt_deg_cw=float(p.rotation_deg),
        rgba=rgba,
    )
    ox, oy = _plan_paste_xy_(p, layer)
    tl = (ox + float(qt.tl[0]), oy + float(qt.tl[1]))
    tr = (ox + float(qt.tr[0]), oy + float(qt.tr[1]))
    q = (0.5 * (tl[0] + tr[0]), 0.5 * (tl[1] + tr[1]))
    return tl, tr, q


def nudge_n3_unit0_top_edge_mid_on_arc_normal(
    plans: List[PortraitPlan],
    *,
    product_w: int,
    product_h: int,
    hero_rgba: Optional[Image.Image] = None,
    unit_rgba: Optional[Image.Image] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    N=3: unit0（ユーザーunit1）の四隅上辺中点 Q を、
    hero頂〜unit2頂の円弧中点 M に一致させる。

    Q 定義: ProductQuad の TL–TR 辺の中点（4隅を結ぶ線の上辺のみ）。
    AABB 水平上辺は使わない。
    M は円弧上かつ法線（緑＝半径 O→M）上なので、
    「緑線上かつ円弧に沿った位置」＝ M そのもの。
    """
    src_h = hero_rgba if isinstance(hero_rgba, Image.Image) else unit_rgba
    src_u = unit_rgba if isinstance(unit_rgba, Image.Image) else hero_rgba
    if not isinstance(src_u, Image.Image):
        return list(plans), {"applied": False, "reason": "no_rgba"}
    hero_p, u0_p, u2_p = _n3_hero_u0_u1_(plans)

    ph = _peak_world_of_plan_(
        hero_p, product_w=product_w, product_h=product_h, rgba=src_h
    )
    p0 = _peak_world_of_plan_(
        u0_p, product_w=product_w, product_h=product_h, rgba=src_u
    )
    p2 = _peak_world_of_plan_(
        u2_p, product_w=product_w, product_h=product_h, rgba=src_u
    )
    circ = _circle_from_three_xy_(ph, p0, p2)
    if circ is None:
        circ = _circle_from_three_xy_(
            ph, (0.5 * (ph[0] + p2[0]), min(ph[1], p2[1]) - 40.0), p2
        )
    if circ is None:
        return list(plans), {"applied": False, "reason": "no_circle"}

    ox, oy, radius = circ

    def _ang(pt: Tuple[float, float]) -> float:
        return math.degrees(math.atan2(pt[0] - ox, oy - pt[1]))

    ah, a2 = _ang(ph), _ang(p2)
    da = a2 - ah
    while da > 180.0:
        da -= 360.0
    while da < -180.0:
        da += 360.0
    a_mid = ah + 0.5 * da
    rad = math.radians(a_mid)
    mx = ox + radius * math.sin(rad)
    my = oy - radius * math.cos(rad)

    layer = _prepare_scaled_rotated_layer(
        product_w=product_w,
        product_h=product_h,
        scale=float(u0_p.scale),
        tilt_deg_cw=float(u0_p.rotation_deg),
        rgba=src_u,
    )
    tx_l, ty_l = measure_top_xy(layer)
    q0 = ProductQuad.from_upright_rgba(src_u)
    qt = q0.transformed(
        scale=float(u0_p.scale), tilt_deg_cw=float(u0_p.rotation_deg)
    )
    # Q_local = mid(TL,TR); Q_world = (top_x - tx_l + qx_l, top_y - ty_l + qy_l)
    qx_l = 0.5 * (float(qt.tl[0]) + float(qt.tr[0]))
    qy_l = 0.5 * (float(qt.tl[1]) + float(qt.tr[1]))
    new_top_x = float(mx) - qx_l + float(tx_l)
    new_top_y = float(my) - qy_l + float(ty_l)
    old_top_x = float(u0_p.top_x if u0_p.top_x is not None else u0_p.x)
    old_top_y = float(u0_p.top_y if u0_p.top_y is not None else u0_p.y)
    u0_new = replace(
        u0_p,
        top_x=int(round(new_top_x)),
        top_y=int(round(new_top_y)),
        x=int(round(float(u0_p.x) + (new_top_x - old_top_x))),
        y=int(round(float(u0_p.y) + (new_top_y - old_top_y))),
    )
    out: List[PortraitPlan] = []
    for p in plans:
        if (
            str(p.role) == "unit"
            and abs(float(p.rotation_deg) - float(u0_p.rotation_deg)) < 1e-6
            and int(p.z) == int(u0_p.z)
        ):
            out.append(u0_new)
        else:
            out.append(p)

    tl_w, tr_w, q_w = _product_quad_top_lr_mid_world_(
        u0_new, product_w=product_w, product_h=product_h, rgba=src_u
    )
    dx, dy = mx - ox, my - oy
    ln = math.hypot(dx, dy) or 1.0
    dist_normal = abs((q_w[0] - ox) * dy - (q_w[1] - oy) * dx) / ln
    dist_to_arc_mid = math.hypot(q_w[0] - mx, q_w[1] - my)
    dist_to_circle = abs(math.hypot(q_w[0] - ox, q_w[1] - oy) - radius)

    def _quad_top_mid_w(p: PortraitPlan, rgba: Image.Image) -> Tuple[float, float]:
        _tl, _tr, qm = _product_quad_top_lr_mid_world_(
            p, product_w=product_w, product_h=product_h, rgba=rgba
        )
        return qm

    qh = _quad_top_mid_w(hero_p, src_h if isinstance(src_h, Image.Image) else src_u)
    q_u2 = _quad_top_mid_w(u2_p, src_u)
    span = max(1e-6, q_u2[0] - qh[0])
    edge_frac = (q_w[0] - qh[0]) / span

    return out, {
        "applied": True,
        "ruleJa": "unit1四隅上辺(TL-TR)中点Qを円弧中点Mへ（緑法線∩円弧）",
        "qSource": "productQuadTopEdgeMid",
        "circle": {"cx": round(ox, 1), "cy": round(oy, 1), "r": round(radius, 1)},
        "arcMid": {"x": round(mx, 1), "y": round(my, 1)},
        "unit0TopEdgeLeft": {"x": round(tl_w[0], 1), "y": round(tl_w[1], 1)},
        "unit0TopEdgeRight": {"x": round(tr_w[0], 1), "y": round(tr_w[1], 1)},
        "unit0TopEdgeMid": {"x": round(q_w[0], 1), "y": round(q_w[1], 1)},
        "qPoint": {"x": round(q_w[0], 1), "y": round(q_w[1], 1)},
        "distToNormalPx": round(dist_normal, 2),
        "distToArcMidPx": round(dist_to_arc_mid, 2),
        "distToCirclePx": round(dist_to_circle, 2),
        "topEdgeFracAlongHeroUnit2": round(edge_frac, 3),
        "topXBefore": round(old_top_x, 1),
        "topYBefore": round(old_top_y, 1),
        "topXAfter": int(round(new_top_x)),
        "topYAfter": int(round(new_top_y)),
    }


def nudge_n3_unit0_up_along_arc_normal(
    plans: List[PortraitPlan],
    *,
    product_w: int,
    product_h: int,
    nudge_up_px: float,
    hero_rgba: Optional[Image.Image] = None,
    unit_rgba: Optional[Image.Image] = None,
    edge_meta: Optional[Dict[str, Any]] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    Q=M のあと、緑法線（円心→ARC-MID）に沿って unit0 を「上」へ平行移動。

    注意（幾何）:
      - 緑線上は維持される（角度＝弧中点）。
      - 円弧上条件は崩れる（中心距離 ≠ r）。nudge>0 で外側（だいたい上）。
    """
    amt = float(nudge_up_px or 0.0)
    if abs(amt) < 1e-9:
        return list(plans), {"applied": False, "reason": "nudge_zero"}

    src_h = hero_rgba if isinstance(hero_rgba, Image.Image) else unit_rgba
    src_u = unit_rgba if isinstance(unit_rgba, Image.Image) else hero_rgba
    if not isinstance(src_u, Image.Image):
        return list(plans), {"applied": False, "reason": "no_rgba"}

    em = dict(edge_meta or {})
    circ = em.get("circle") or {}
    am = em.get("arcMid") or {}
    if circ.get("cx") is None or am.get("x") is None:
        # edge_meta が無い場合は再計算（Q配置と同じ円）
        hero_p, u0_p, u2_p = _n3_hero_u0_u1_(plans)
        ph = _peak_world_of_plan_(
            hero_p, product_w=product_w, product_h=product_h, rgba=src_h
        )
        p0 = _peak_world_of_plan_(
            u0_p, product_w=product_w, product_h=product_h, rgba=src_u
        )
        p2 = _peak_world_of_plan_(
            u2_p, product_w=product_w, product_h=product_h, rgba=src_u
        )
        c3 = _circle_from_three_xy_(ph, p0, p2)
        if c3 is None:
            return list(plans), {"applied": False, "reason": "no_circle"}
        ox, oy, radius = c3

        def _ang(pt: Tuple[float, float]) -> float:
            return math.degrees(math.atan2(pt[0] - ox, oy - pt[1]))

        ah, a2 = _ang(ph), _ang(p2)
        da = a2 - ah
        while da > 180.0:
            da -= 360.0
        while da < -180.0:
            da += 360.0
        rad = math.radians(ah + 0.5 * da)
        mx = ox + radius * math.sin(rad)
        my = oy - radius * math.cos(rad)
    else:
        ox = float(circ["cx"])
        oy = float(circ["cy"])
        radius = float(circ["r"])
        mx = float(am["x"])
        my = float(am["y"])

    dx, dy = mx - ox, my - oy
    ln = math.hypot(dx, dy) or 1.0
    ux, uy = dx / ln, dy / ln
    # キャンバス上「上」＝Y減。法線方向のうち Y が減る向きへ移動
    if uy > 0:
        ux, uy = -ux, -uy
    # amt>0 → 上へ（円の外側寄りになりやすい）
    shift_x = ux * amt
    shift_y = uy * amt

    hero_p, u0_p, u2_p = _n3_hero_u0_u1_(plans)
    old_top_x = float(u0_p.top_x if u0_p.top_x is not None else u0_p.x)
    old_top_y = float(u0_p.top_y if u0_p.top_y is not None else u0_p.y)
    new_top_x = old_top_x + shift_x
    new_top_y = old_top_y + shift_y
    u0_new = replace(
        u0_p,
        top_x=int(round(new_top_x)),
        top_y=int(round(new_top_y)),
        x=int(round(float(u0_p.x) + shift_x)),
        y=int(round(float(u0_p.y) + shift_y)),
    )
    out: List[PortraitPlan] = []
    for p in plans:
        if (
            str(p.role) == "unit"
            and abs(float(p.rotation_deg) - float(u0_p.rotation_deg)) < 1e-6
            and int(p.z) == int(u0_p.z)
        ):
            out.append(u0_new)
        else:
            out.append(p)

    _tl, _tr, q_w = _product_quad_top_lr_mid_world_(
        u0_new, product_w=product_w, product_h=product_h, rgba=src_u
    )
    dist_normal = abs((q_w[0] - ox) * dy - (q_w[1] - oy) * dx) / ln
    dist_to_arc_mid = math.hypot(q_w[0] - mx, q_w[1] - my)
    dist_to_circle = abs(math.hypot(q_w[0] - ox, q_w[1] - oy) - radius)

    return out, {
        "applied": True,
        "ruleJa": "Q=M後に緑法線上で上へ平行移動（円弧上は意図的に外す）",
        "nudgeUpPx": round(amt, 2),
        "shift": {"x": round(shift_x, 2), "y": round(shift_y, 2)},
        "dirUnit": {"x": round(ux, 4), "y": round(uy, 4)},
        "qPointAfter": {"x": round(q_w[0], 1), "y": round(q_w[1], 1)},
        "arcMid": {"x": round(mx, 1), "y": round(my, 1)},
        "distToNormalPx": round(dist_normal, 2),
        "distToArcMidPx": round(dist_to_arc_mid, 2),
        "distToCirclePx": round(dist_to_circle, 2),
        "topXAfter": int(round(new_top_x)),
        "topYAfter": int(round(new_top_y)),
        "note": "緑線は維持／円弧上は破綻（半径が変わる）",
    }


def propose_top_arc_fan_plans(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    tilts_deg_cw: List[float],
    scales: List[float],
    top_arc_cx_ratio: float = 0.1304,
    top_arc_cy_ratio: float = 2.777,
    top_arc_radius_ratio: float = 2.794,
    top_arc_angles_deg: Optional[List[float]] = None,
    bottom_protrude_px: float = 36.0,
    top_down_px: Optional[List[float]] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    N本・上部円弧扇（N=3合格ルールの一般形）:
    - 各瓶の『頂』を同一円周上に置く
    - 円弧角と本体傾きは分離
    - bottomProtrude で下端を枠下へ少し出す
    - topDownPx[i] で奥の瓶を段階的に下げ（右隣の右上より左上が下）
    - z: 右奥が最下、hero最前
    """
    n = int(n)
    if n < 3:
        raise ValueError("top_arc_fan requires n>=3")
    canvas = int(canvas)
    if len(tilts_deg_cw) < n or len(scales) < n:
        raise ValueError("tilts/scales length must be >= n")

    tilts: List[float] = []
    for i, t in enumerate(tilts_deg_cw[:n]):
        tt = float(t)
        if i == 0 and tt > 0:
            tt = -abs(tt)
        if i > 0 and tt < 0:
            tt = abs(tt)
        tilts.append(tt)

    scs = [max(0.05, min(1.25, float(s))) for s in scales[:n]]
    cx = float(canvas) * float(top_arc_cx_ratio)
    cy = float(canvas) * float(top_arc_cy_ratio)
    radius = float(canvas) * float(top_arc_radius_ratio)
    default_angles = [3.65, 8.8, 14.02, 20.8, 27.0][:n]
    angles = list(top_arc_angles_deg) if top_arc_angles_deg else default_angles
    if len(angles) < n:
        # extend by ~6.5deg steps
        while len(angles) < n:
            angles.append(float(angles[-1]) + 6.5)
    downs = list(top_down_px) if top_down_px is not None else ([0.0] * n)
    while len(downs) < n:
        downs.append(0.0)
    protrude = float(bottom_protrude_px)

    def _top_on_arc(angle_deg_cw_from_up: float) -> Tuple[int, int]:
        rad = math.radians(float(angle_deg_cw_from_up))
        tx = cx + radius * math.sin(rad)
        ty = cy - radius * math.cos(rad) + protrude
        return int(round(tx)), int(round(ty))

    tops: List[Tuple[int, int]] = []
    for i in range(n):
        tx, ty = _top_on_arc(angles[i])
        ty = int(round(ty + float(downs[i])))
        tops.append((tx, ty))

    plans: List[PortraitPlan] = []
    # paint back-to-front: last unit first, hero last
    for i in range(n - 1, -1, -1):
        role = "hero" if i == 0 else "unit"
        z = 1000 if i == 0 else (10 * (n - 1 - i))
        tx, ty = tops[i]
        plans.append(
            PortraitPlan(
                role, 0, 0, scs[i], z=z, rotation_deg=tilts[i],
                anchor="top", top_x=tx, top_y=ty,
            )
        )

    meta: Dict[str, Any] = {
        "layoutFamily": "portrait_top_arc_fan",
        "pattern": "n%d_top_arc_fan" % n,
        "n": n,
        "scale": scs[0],
        "scaleHero": scs[0],
        "scales": scs,
        "tiltsDegCw": [round(t, 2) for t in tilts],
        "topArc": {
            "cx": round(cx, 1),
            "cy": round(cy, 1),
            "radius": round(radius, 1),
            "cxRatio": float(top_arc_cx_ratio),
            "cyRatio": float(top_arc_cy_ratio),
            "radiusRatio": float(top_arc_radius_ratio),
            "anglesDegCwFromUp": [round(float(a), 3) for a in angles[:n]],
            "bottomProtrudePx": protrude,
            "topDownPx": [float(d) for d in downs[:n]],
            "ruleJa": (
                "N頂を同一円周上。"
                "円弧角≠本体傾き。"
                "bottomProtrudeで下端はみ出し。"
                "topDownPxで奥瓶を下げ右隣右上より左上を下へ。"
            ),
        },
        "tops": [
            {"role": ("hero" if i == 0 else "unit%d" % (i - 1)), "x": tops[i][0], "y": tops[i][1]}
            for i in range(n)
        ],
        "anchor": "top",
        "allowEdgeCrop": True,
        "allowBottomProtrude": True,
        "zOrderJa": "右奥→…→hero最前。上部は円弧",
        "proposalJa": "N=%d: 上部円弧扇（N=3合格ルールの拡張）。" % n,
        "canvasStillSquare": True,
        "sampleAlignJa": "見本縦型レイアウト基本 / N=3合格パラメータ拡張",
        "productWhRef": [product_w, product_h],
    }
    return plans, meta


def propose_n3_fan_plans(
    *,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_tilt_deg_cw: float = -14.0,
    unit0_tilt_deg_cw: float = 18.0,
    unit1_tilt_deg_cw: float = 38.0,
    hero_scale: float = 1.02,
    unit0_scale_ratio: float = 0.98,
    unit1_scale_ratio: float = 0.96,
    top_arc_cx_ratio: float = 0.1304,
    top_arc_cy_ratio: float = 2.777,
    top_arc_radius_ratio: float = 2.794,
    top_arc_angles_deg: Optional[List[float]] = None,
    bottom_protrude_px: float = 36.0,
    unit1_top_down_px: float = 55.0,
    **_compat: Any,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """N=3合格ラッパ（2026-08-05 合格固定）。"""
    sh = float(hero_scale)
    plans, meta = propose_top_arc_fan_plans(
        n=3,
        canvas=canvas,
        product_w=product_w,
        product_h=product_h,
        tilts_deg_cw=[hero_tilt_deg_cw, unit0_tilt_deg_cw, unit1_tilt_deg_cw],
        scales=[sh, sh * float(unit0_scale_ratio), sh * float(unit1_scale_ratio)],
        top_arc_cx_ratio=top_arc_cx_ratio,
        top_arc_cy_ratio=top_arc_cy_ratio,
        top_arc_radius_ratio=top_arc_radius_ratio,
        top_arc_angles_deg=top_arc_angles_deg or [3.65, 8.8, 14.02],
        bottom_protrude_px=bottom_protrude_px,
        top_down_px=[0.0, 0.0, float(unit1_top_down_px)],
    )
    meta["pattern"] = "n3_top_arc_fan_locked"
    meta["layoutFamily"] = "portrait_n3_top_arc_fan"
    meta["sampleAlignJa"] = "見本127 N=3: 頂円弧+下はみ出し+unit1下・中角8.8°(頂X中点)+傾き-14/18/38"
    return plans, meta


def _synth_scaled_rotated_rect(
    product_w: int,
    product_h: int,
    scale: float,
    tilt_deg_cw: float,
) -> Image.Image:
    """計測用の矩形シルエット（実画像なしで頂/足/左端を近似）。"""
    pw = max(1, int(round(float(product_w) * float(scale))))
    ph = max(1, int(round(float(product_h) * float(scale))))
    im = Image.new("RGBA", (pw, ph), (0, 0, 0, 0))
    im.putalpha(Image.new("L", (pw, ph), 255))
    return rotate_rgba_cw(im, float(tilt_deg_cw))


def _prepare_scaled_rotated_layer(
    *,
    product_w: int,
    product_h: int,
    scale: float,
    tilt_deg_cw: float,
    rgba: Optional[Image.Image] = None,
) -> Image.Image:
    """頂/足計測用レイヤ。rgbaがあれば実画像、なければ矩形近似。"""
    if rgba is not None:
        im = rgba.convert("RGBA")
        nw = max(1, int(round(im.width * float(scale))))
        nh = max(1, int(round(im.height * float(scale))))
        if (nw, nh) != im.size:
            im = im.resize((nw, nh), resample=Image.Resampling.BICUBIC)
        return rotate_rgba_cw(im, float(tilt_deg_cw))
    return _synth_scaled_rotated_rect(product_w, product_h, scale, tilt_deg_cw)


def _leftmost_near_top_x(
    im: Image.Image,
    *,
    alpha_thresh: int = 96,
    band_below_top: int = 50,
) -> float:
    a = im.convert("RGBA").getchannel("A")
    w, h = im.size
    ap = a.load()
    top_y = None
    for y in range(h):
        if any(ap[x, y] > alpha_thresh for x in range(w)):
            top_y = y
            break
    if top_y is None:
        return 0.0
    y1 = min(h, int(top_y) + int(band_below_top))
    left = None
    for y in range(int(top_y), y1):
        for x in range(w):
            if ap[x, y] > alpha_thresh:
                left = x if left is None else min(left, x)
                break
    return float(0 if left is None else left)


def measure_aabb4(
    im: Image.Image, *, alpha_thresh: int = 96
) -> Tuple[int, int, int, int]:
    """
    商品の4点（最も左・右・上・下）を結んだ矩形。
    戻り値: (left, top, right, bottom) いずれも inclusive。
    """
    a = im.convert("RGBA").getchannel("A")
    bb = a.point(lambda v: 255 if v > int(alpha_thresh) else 0).getbbox()
    if not bb:
        w, h = im.size
        return (0, 0, max(0, w - 1), max(0, h - 1))
    # PIL getbbox の right/bottom は exclusive
    return (int(bb[0]), int(bb[1]), int(bb[2] - 1), int(bb[3] - 1))


@dataclass
class ProductQuad:
    """
    直立1個画像（透過PNG）上で先に決める商品モデル。
    四隅 TL/TR/BL/BR と、それを結ぶ4辺。
    """

    tl: Tuple[float, float]
    tr: Tuple[float, float]
    bl: Tuple[float, float]
    br: Tuple[float, float]
    src_w: int
    src_h: int

    @classmethod
    def from_upright_rgba(
        cls, im: Image.Image, *, alpha_thresh: int = 96
    ) -> "ProductQuad":
        im = im.convert("RGBA")
        l, t, r, b = measure_aabb4(im, alpha_thresh=alpha_thresh)
        w, h = im.size
        return cls(
            tl=(float(l), float(t)),
            tr=(float(r), float(t)),
            bl=(float(l), float(b)),
            br=(float(r), float(b)),
            src_w=int(w),
            src_h=int(h),
        )

    def edges(self) -> Dict[str, Tuple[Tuple[float, float], Tuple[float, float]]]:
        return {
            "top": (self.tl, self.tr),
            "bottom": (self.bl, self.br),
            "left": (self.tl, self.bl),
            "right": (self.tr, self.br),
        }

    def as_dict(self) -> Dict[str, Any]:
        return {
            "tl": [round(self.tl[0], 2), round(self.tl[1], 2)],
            "tr": [round(self.tr[0], 2), round(self.tr[1], 2)],
            "bl": [round(self.bl[0], 2), round(self.bl[1], 2)],
            "br": [round(self.br[0], 2), round(self.br[1], 2)],
            "srcWh": [self.src_w, self.src_h],
            "edges": {
                k: [[round(a[0], 2), round(a[1], 2)], [round(b[0], 2), round(b[1], 2)]]
                for k, (a, b) in self.edges().items()
            },
        }

    def transformed(
        self, *, scale: float, tilt_deg_cw: float
    ) -> "ProductQuad":
        """
        直立四隅に等倍スケール＋CW回転（expand）を適用し、
        rotate_rgba_cw 後レイヤ座標系の四隅を返す。
        リサイズ寸法は paste と同じ int(round(src*scale))。
        """
        s = float(scale)
        sw = max(1, int(round(float(self.src_w) * s)))
        sh = max(1, int(round(float(self.src_h) * s)))
        sx = sw / float(max(1, self.src_w))
        sy = sh / float(max(1, self.src_h))
        pts = [
            (self.tl[0] * sx, self.tl[1] * sy),
            (self.tr[0] * sx, self.tr[1] * sy),
            (self.bl[0] * sx, self.bl[1] * sy),
            (self.br[0] * sx, self.br[1] * sy),
        ]
        mapped, rw, rh = _map_points_through_pil_rotate_cw(
            pts, width=float(sw), height=float(sh), tilt_deg_cw=float(tilt_deg_cw)
        )
        return ProductQuad(
            tl=mapped[0],
            tr=mapped[1],
            bl=mapped[2],
            br=mapped[3],
            src_w=int(rw),
            src_h=int(rh),
        )

    def aabb(self) -> Tuple[float, float, float, float]:
        """レイヤ座標の外接AABB (min_x, min_y, max_x, max_y)。"""
        xs = [self.tl[0], self.tr[0], self.bl[0], self.br[0]]
        ys = [self.tl[1], self.tr[1], self.bl[1], self.br[1]]
        return (min(xs), min(ys), max(xs), max(ys))


def render_upright_quad_annot(
    unit_rgba: Image.Image,
    *,
    display_scale: float = 0.90,
    pad: int = 40,
    label: str = "",
) -> Image.Image:
    """
    直立1個画像に四隅★と赤辺を描画した検証用RGB。
    本番 portrait 出力時に併せて保存する正本。
    """
    from PIL import ImageDraw

    unit = unit_rgba.convert("RGBA")
    quad = ProductQuad.from_upright_rgba(unit)
    s = max(0.05, float(display_scale))
    w = max(1, int(round(unit.width * s)))
    h = max(1, int(round(unit.height * s)))
    canvas = Image.new("RGB", (w + pad * 2, h + pad * 2), (255, 255, 255))
    scaled = unit.resize((w, h), Image.Resampling.LANCZOS)
    canvas.paste(scaled, (pad, pad), scaled)
    draw = ImageDraw.Draw(canvas)

    def _map(pt: Tuple[float, float]) -> Tuple[float, float]:
        return (pt[0] * s + pad, pt[1] * s + pad)

    corners = {
        "TL": _map(quad.tl),
        "TR": _map(quad.tr),
        "BR": _map(quad.br),
        "BL": _map(quad.bl),
    }
    edge_pairs = [
        (corners["TL"], corners["TR"]),
        (corners["TR"], corners["BR"]),
        (corners["BR"], corners["BL"]),
        (corners["BL"], corners["TL"]),
    ]
    for a, b in edge_pairs:
        draw.line([a, b], fill=(255, 30, 30), width=5)
    for name, (x, y) in corners.items():
        r = 16
        draw.ellipse((x - r, y - r, x + r, y + r), outline=(255, 200, 0), width=4)
        draw.line((x - r, y, x + r, y), fill=(255, 215, 0), width=3)
        draw.line((x, y - r, x, y + r), fill=(255, 215, 0), width=3)
        draw.line(
            (x - r * 0.7, y - r * 0.7, x + r * 0.7, y + r * 0.7),
            fill=(255, 215, 0),
            width=2,
        )
        draw.line(
            (x - r * 0.7, y + r * 0.7, x + r * 0.7, y - r * 0.7),
            fill=(255, 215, 0),
            width=2,
        )
        draw.text((x + 18, y - 20), name, fill=(200, 0, 0))

    title = label or "upright 1-unit quad"
    draw.text(
        (8, 8),
        f"{title} @{s:.0%} size={unit.size} -> {(w, h)}",
        fill=(0, 0, 0),
    )
    draw.text(
        (8, 28),
        f"TL{tuple(round(v, 1) for v in quad.tl)} "
        f"TR{tuple(round(v, 1) for v in quad.tr)} "
        f"BL{tuple(round(v, 1) for v in quad.bl)} "
        f"BR{tuple(round(v, 1) for v in quad.br)}",
        fill=(120, 0, 0),
    )
    return canvas


def _map_points_through_pil_rotate_cw(
    points: List[Tuple[float, float]],
    *,
    width: float,
    height: float,
    tilt_deg_cw: float,
) -> Tuple[List[Tuple[float, float]], int, int]:
    """
    rotate_rgba_cw（PIL Image.rotate(-tilt, expand=True)）と一致する点変換。
    Pillow は dest→src の逆アフィンを使うため、同じ行列を組み立てて逆写像する。
    戻り値: (expand後レイヤ座標の点列, nw, nh)
    """
    w = float(width)
    h = float(height)
    if abs(float(tilt_deg_cw)) < 1e-9:
        return (list(points), max(1, int(round(w))), max(1, int(round(h))))

    # Pillow Image.rotate と同じ dest→src 行列
    angle_arg = -float(tilt_deg_cw)
    angle = -math.radians(angle_arg)
    matrix = [
        round(math.cos(angle), 15),
        round(math.sin(angle), 15),
        0.0,
        round(-math.sin(angle), 15),
        round(math.cos(angle), 15),
        0.0,
    ]

    def _transform(x: float, y: float, m: List[float]) -> Tuple[float, float]:
        a, b, c, d, e, f = m
        return a * x + b * y + c, d * x + e * y + f

    center = (w / 2.0, h / 2.0)
    matrix[2], matrix[5] = _transform(-center[0], -center[1], matrix)
    matrix[2] += center[0]
    matrix[5] += center[1]

    xx: List[float] = []
    yy: List[float] = []
    for x, y in ((0.0, 0.0), (w, 0.0), (w, h), (0.0, h)):
        tx, ty = _transform(x, y, matrix)
        xx.append(tx)
        yy.append(ty)
    nw = int(math.ceil(max(xx)) - math.floor(min(xx)))
    nh = int(math.ceil(max(yy)) - math.floor(min(yy)))
    matrix[2], matrix[5] = _transform(-(nw - w) / 2.0, -(nh - h) / 2.0, matrix)

    a, b, c, d, e, f = matrix
    det = a * e - b * d
    if abs(det) < 1e-12:
        return (list(points), max(1, nw), max(1, nh))
    ia, ib = e / det, -b / det
    id_, ie = -d / det, a / det

    def _src_to_dest(xs: float, ys: float) -> Tuple[float, float]:
        x = xs - c
        y = ys - f
        return (ia * x + ib * y, id_ * x + ie * y)

    return ([_src_to_dest(x, y) for x, y in points], max(1, nw), max(1, nh))


def _circle_from_two_points_and_cy(
    ax: float, ay: float, bx: float, by: float, cy: float
) -> Tuple[float, float, float]:
    """2点と円中心Yから円 (cx, cy, R) を求める。"""
    den = 2.0 * (bx - ax)
    if abs(den) < 1e-6:
        cx = float(ax)
    else:
        cx = (
            (bx * bx - ax * ax)
            + (by - cy) * (by - cy)
            - (ay - cy) * (ay - cy)
        ) / den
    r = math.hypot(ax - cx, ay - cy)
    return float(cx), float(cy), float(max(1.0, r))


def _angle_cw_from_up(cx: float, cy: float, px: float, py: float) -> float:
    """円中心から点への角度（真上=0、時計回り正・度）。"""
    return math.degrees(math.atan2(px - cx, cy - py))


def _point_on_circle(cx: float, cy: float, r: float, angle_deg_cw_from_up: float) -> Tuple[float, float]:
    rad = math.radians(float(angle_deg_cw_from_up))
    return (cx + r * math.sin(rad), cy - r * math.cos(rad))


def _circle_from_three_points(
    ax: float, ay: float, bx: float, by: float, cx: float, cy: float
) -> Optional[Tuple[float, float, float]]:
    """3点から円 (ox, oy, R)。同一直線なら None。"""
    d = 2.0 * (ax * (by - cy) + bx * (cy - ay) + cx * (ay - by))
    if abs(d) < 1e-6:
        return None
    a2 = ax * ax + ay * ay
    b2 = bx * bx + by * by
    c2 = cx * cx + cy * cy
    ox = (a2 * (by - cy) + b2 * (cy - ay) + c2 * (ay - by)) / d
    oy = (a2 * (cx - bx) + b2 * (ax - cx) + c2 * (bx - ax)) / d
    r = math.hypot(ax - ox, ay - oy)
    if r < 10:
        return None
    return (float(ox), float(oy), float(r))


def propose_n4_fan_plans(
    *,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_tilt_deg_cw: float = -15.0,
    unit0_tilt_deg_cw: float = 15.0,
    unit1_tilt_deg_cw: float = 35.0,
    unit2_tilt_deg_cw: float = 55.0,
    hero_height_fill: float = 0.90,
    unit0_foot_left_of_hero_center_px: float = 40.0,
    unit1_tilt_deg_cw_override: Optional[float] = None,
    unit2_tilt_deg_cw_override: Optional[float] = None,
    hero_rgba: Optional[Image.Image] = None,
    unit_rgba: Optional[Image.Image] = None,
    # legacy / unused knobs (cfg compat)
    hero_scale: float = 0.85,
    unit0_scale_ratio: float = 1.0,
    unit1_scale_ratio: float = 1.0,
    unit2_scale_ratio: float = 1.0,
    **_compat: Any,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    N=4 高さ・扇ロジック（正）:

    【ヒーロー】
      枠縦×hero_height_fill・縦横比固定・−15°・下辺+左辺接触・はみ出し禁止

    【2本目 unit0 — 高さ】
      トップ（シルエット頂）が枠の上辺に接する位置に置く。
      起点（商品底辺の左端）は hero 縦中心より若干左（ヒーロー背面）。

    【4本目 unit2 — 扇の端】
      トップ（頂帯の最右）が枠の右辺に接する。
      底辺の右縁が枠の下辺に接する。
      起点（底辺左端）はヒーロー背面の下寄り（2本目より低い）。

    【3本目 unit1 — 高さ・横】
      2本目トップと4本目トップを通る円弧上にトップを載せる。
      上部辺の左右端が、2本目上部右端・4本目上部左端と等距離になるよう横位置を決める。

    【起点】
      定義＝商品底辺の左端（最下帯の最左点）。
      2→3→4本目の起点はヒーロー背面に密集し、上→中→下へ段差。
    """
    canvas = int(canvas)
    ht = -abs(float(hero_tilt_deg_cw))
    u0t = abs(float(unit0_tilt_deg_cw))
    u1t = abs(
        float(
            unit1_tilt_deg_cw_override
            if unit1_tilt_deg_cw_override is not None
            else unit1_tilt_deg_cw
        )
    )
    u2t = abs(
        float(
            unit2_tilt_deg_cw_override
            if unit2_tilt_deg_cw_override is not None
            else unit2_tilt_deg_cw
        )
    )
    tilts = [ht, u0t, u1t, u2t]

    src_h = hero_rgba if hero_rgba is not None else unit_rgba
    src_u = unit_rgba if unit_rgba is not None else hero_rgba

    hero_ref = _prepare_scaled_rotated_layer(
        product_w=product_w, product_h=product_h, scale=1.0, tilt_deg_cw=ht, rgba=src_h
    )
    _l1, _t1, _r1, b1 = measure_aabb4(hero_ref)
    h1 = max(1, b1 - _t1 + 1)
    target_h = float(canvas) * float(hero_height_fill)
    sh = max(0.05, min(1.25, target_h / float(h1)))

    def _rebuild(scale: float) -> None:
        nonlocal sh, hero_layer, unit0_layer, unit1_layer, unit2_layer
        nonlocal hl, ht0, hr, hb, u0l, u0t0, u0r, u0b
        nonlocal h_top, u0_top, u0_origin, h_origin, u1_top_local, u2_top_local, u1_aabb, u2_aabb
        nonlocal u2_top_right_local, u1_origin, u2_origin, u2_bottom_right
        nonlocal u0_top_right_local, u1_top_left_local, u1_top_right_local, u2_top_left_local
        sh = float(scale)
        hero_layer = _prepare_scaled_rotated_layer(
            product_w=product_w, product_h=product_h, scale=sh, tilt_deg_cw=ht, rgba=src_h
        )
        unit0_layer = _prepare_scaled_rotated_layer(
            product_w=product_w, product_h=product_h, scale=sh, tilt_deg_cw=u0t, rgba=src_u
        )
        unit1_layer = _prepare_scaled_rotated_layer(
            product_w=product_w, product_h=product_h, scale=sh, tilt_deg_cw=u1t, rgba=src_u
        )
        unit2_layer = _prepare_scaled_rotated_layer(
            product_w=product_w, product_h=product_h, scale=sh, tilt_deg_cw=u2t, rgba=src_u
        )
        hl, ht0, hr, hb = measure_aabb4(hero_layer)
        u0l, u0t0, u0r, u0b = measure_aabb4(unit0_layer)
        h_top = measure_top_xy(hero_layer)
        u0_top = measure_top_xy(unit0_layer)
        u0_origin = measure_bottom_left_xy(unit0_layer)
        h_origin = measure_bottom_left_xy(hero_layer)
        u1_top_local = measure_top_xy(unit1_layer)
        u2_top_local = measure_top_xy(unit2_layer)
        u0_top_right_local = measure_top_right_xy(unit0_layer)
        u1_top_left_local = measure_top_left_xy(unit1_layer)
        u1_top_right_local = measure_top_right_xy(unit1_layer)
        u2_top_left_local = measure_top_left_xy(unit2_layer)
        u2_top_right_local = measure_top_right_xy(unit2_layer)
        u1_origin = measure_bottom_left_xy(unit1_layer)
        u2_origin = measure_bottom_left_xy(unit2_layer)
        u2_bottom_right = measure_bottom_right_xy(unit2_layer)
        u1_aabb = measure_aabb4(unit1_layer)
        u2_aabb = measure_aabb4(unit2_layer)

    hero_layer = unit0_layer = unit1_layer = unit2_layer = None  # type: ignore
    hl = ht0 = hr = hb = u0l = u0t0 = u0r = u0b = 0
    h_top = u0_top = u0_origin = h_origin = (0.0, 0.0)
    u1_top_local = u2_top_local = u2_top_right_local = (0.0, 0.0)
    u0_top_right_local = u1_top_left_local = u1_top_right_local = u2_top_left_local = (0.0, 0.0)
    u1_origin = u2_origin = u2_bottom_right = (0.0, 0.0)
    u1_aabb = u2_aabb = (0, 0, 0, 0)
    _rebuild(sh)

    hero_paste_x = 0 - hl
    hero_paste_y = (canvas - 1) - hb
    hero_world = {
        "left": hero_paste_x + hl,
        "top": hero_paste_y + ht0,
        "right": hero_paste_x + hr,
        "bottom": hero_paste_y + hb,
    }
    if hero_world["right"] > canvas - 1:
        need_w = float(hr - hl + 1)
        if need_w > canvas:
            _rebuild(sh * (float(canvas) / need_w) * 0.999)
            hero_paste_x = 0 - hl
            hero_paste_y = (canvas - 1) - hb
            hero_world = {
                "left": hero_paste_x + hl,
                "top": hero_paste_y + ht0,
                "right": hero_paste_x + hr,
                "bottom": hero_paste_y + hb,
            }
    hero_overflow = (
        hero_world["left"] < -0.5
        or hero_world["top"] < -0.5
        or hero_world["right"] > canvas - 0.5
        or hero_world["bottom"] > canvas - 0.5
    )
    hero_cx = 0.5 * (hero_world["left"] + hero_world["right"])
    hero_mid_y = 0.5 * (hero_world["top"] + hero_world["bottom"])
    hero_low_y = hero_world["top"] + 0.72 * (hero_world["bottom"] - hero_world["top"])

    # 2本目高さ: トップが枠上に接する / 起点(底辺左端): hero中心やや左
    unit0_paste_y = 0.0 - float(u0_top[1])
    origin_target_x = float(hero_cx) - float(unit0_foot_left_of_hero_center_px)
    unit0_paste_x = origin_target_x - u0_origin[0]
    if unit0_paste_x + u0l < -80:
        unit0_paste_x += -(unit0_paste_x + u0l + 80)
    unit0_top_w = (unit0_paste_x + u0_top[0], unit0_paste_y + u0_top[1])
    unit0_origin_w = (unit0_paste_x + u0_origin[0], unit0_paste_y + u0_origin[1])
    unit0_world = {
        "left": unit0_paste_x + u0l,
        "top": unit0_paste_y + u0t0,
        "right": unit0_paste_x + u0r,
        "bottom": unit0_paste_y + u0b,
    }

    ax = float(hero_paste_x + h_top[0])
    ay = float(hero_paste_y + h_top[1])
    bx = float(unit0_top_w[0])
    by = float(unit0_top_w[1])
    hero_top_w = (ax, ay)
    right_x = float(canvas - 1)

    def _paste_by_top(
        target: Tuple[float, float], top_local: Tuple[float, float]
    ) -> Tuple[float, float]:
        return (target[0] - top_local[0], target[1] - top_local[1])

    def _aabb_overlap_x(
        a: Tuple[float, float, float, float], b: Tuple[float, float, float, float]
    ) -> float:
        return max(0.0, min(a[2], b[2]) - max(a[0], b[0]))

    best = None
    u0_box = (
        float(unit0_world["left"]),
        float(unit0_world["top"]),
        float(unit0_world["right"]),
        float(unit0_world["bottom"]),
    )
    origin1_target_y = float(hero_mid_y)
    origin_cluster_x = float(hero_cx) - 10.0

    # 4本目固定: 頂帯最右=枠右 かつ 底辺右縁=枠下
    paste2 = (
        right_x - u2_top_right_local[0],
        float(canvas - 1) - u2_bottom_right[1],
    )
    tr = (paste2[0] + u2_top_right_local[0], paste2[1] + u2_top_right_local[1])
    br = (paste2[0] + u2_bottom_right[0], paste2[1] + u2_bottom_right[1])
    p2 = (float(tr[0]), float(tr[1]))
    top2_sil = (paste2[0] + u2_top_local[0], paste2[1] + u2_top_local[1])
    origin2_w = (paste2[0] + u2_origin[0], paste2[1] + u2_origin[1])
    box2_fixed = (
        paste2[0] + u2_aabb[0],
        paste2[1] + u2_aabb[1],
        paste2[0] + u2_aabb[2],
        paste2[1] + u2_aabb[3],
    )

    # 2本目上部右端・4本目上部左端（3本目の左右均等用）
    u0_top_right_w = (
        unit0_paste_x + u0_top_right_local[0],
        unit0_paste_y + u0_top_right_local[1],
    )
    u2_top_left_w = (
        paste2[0] + u2_top_left_local[0],
        paste2[1] + u2_top_left_local[1],
    )

    dy = float(p2[1] - by)
    dx = float(p2[0] - bx)
    if dy > 10.0:
        rr = (dx * dx + dy * dy) / (2.0 * dy)
        ox = float(bx)
        oy = float(by + rr)
        ang_0 = _angle_cw_from_up(ox, oy, bx, by)
        ang_4 = _angle_cw_from_up(ox, oy, p2[0], p2[1])
        span = float(ang_4 - ang_0)
        ang_h = _angle_cw_from_up(ox, oy, ax, ay)
        if (
            rr >= canvas * 0.15
            and rr <= canvas * 3.5
            and ang_4 > 5.0
            and 12.0 <= span <= 170.0
        ):
            # 3本目: 上部辺の左右端が、2本目右端・4本目左端と等距離
            # paste_x + u1_tl.x - u0_tr.x = u2_tl.x - (paste_x + u1_tr.x)
            paste1_x = 0.5 * (
                u0_top_right_w[0]
                + u2_top_left_w[0]
                - u1_top_left_local[0]
                - u1_top_right_local[0]
            )
            # 高さは扇円上（シルエット頂のXに対応する円上Y）
            top1_x = paste1_x + u1_top_local[0]
            t = (top1_x - ox) / max(1e-6, rr)
            t = max(-1.0, min(1.0, t))
            ang_1 = math.degrees(math.asin(t))
            if not (ang_0 - 2.0 <= ang_1 <= ang_4 + 2.0):
                # 円弧範囲外なら弧の中点角を使い、X均等は維持
                ang_1 = ang_0 + span * 0.5
            top1_y = oy - rr * math.cos(math.radians(ang_1))
            # asin分岐でYが上側に飛ぶ場合は円の下側（降り側）を取る
            y_alt = oy + rr * math.cos(math.radians(ang_1))
            if abs(top1_y - by) > abs(y_alt - by) and by <= y_alt <= p2[1] + 50:
                top1_y = y_alt
            # 扇の単調下降を優先
            if top1_y < by:
                top1_y = by + max(8.0, (p2[1] - by) * 0.35)
            paste1_y = top1_y - u1_top_local[1]
            paste1 = (paste1_x, paste1_y)
            p1 = (paste1[0] + u1_top_local[0], paste1[1] + u1_top_local[1])
            origin1_w = (paste1[0] + u1_origin[0], paste1[1] + u1_origin[1])
            box1 = (
                paste1[0] + u1_aabb[0],
                paste1[1] + u1_aabb[1],
                paste1[0] + u1_aabb[2],
                paste1[1] + u1_aabb[3],
            )
            u1_tl_w = (
                paste1[0] + u1_top_left_local[0],
                paste1[1] + u1_top_left_local[1],
            )
            u1_tr_w = (
                paste1[0] + u1_top_right_local[0],
                paste1[1] + u1_top_right_local[1],
            )
            gap_l = u1_tl_w[0] - u0_top_right_w[0]
            gap_r = u2_top_left_w[0] - u1_tr_w[0]
            gap_err = abs(gap_l - gap_r) / float(canvas)
            circ_err_1 = abs(math.hypot(p1[0] - ox, p1[1] - oy) - rr) / float(canvas)
            ov01 = _aabb_overlap_x(u0_box, box1) / max(1.0, u0_box[2] - u0_box[0])
            ov12 = _aabb_overlap_x(box1, box2_fixed) / max(1.0, box1[2] - box1[0])
            origin_order_pen = 0.0
            if not (unit0_origin_w[1] < origin1_w[1] < origin2_w[1] + 120):
                origin_order_pen += 0.2
            origin1_pen = abs(origin1_w[1] - origin1_target_y) / float(canvas)
            right_touch_err = abs(tr[0] - right_x) / float(canvas)
            bottom_touch_err = abs(br[1] - float(canvas - 1)) / float(canvas)
            L = min(u0_box[0], box1[0], box2_fixed[0], hero_world["left"])
            R = max(u0_box[2], box1[2], box2_fixed[2], hero_world["right"])
            over_pen = max(0.0, R - (canvas + 100)) / float(canvas)
            left_pen = max(0.0, -220 - L) / float(canvas)
            # 負ギャップ（食い込み過ぎ）は軽く減点
            nest_pen = max(0.0, -gap_l) / float(canvas) + max(0.0, -gap_r) / float(
                canvas
            )
            score = (
                1.0 * ov01
                + 1.0 * ov12
                - 2.5 * gap_err
                - 1.2 * circ_err_1
                - 0.45 * origin1_pen
                - origin_order_pen
                - 0.35 * over_pen
                - 0.10 * left_pen
                - 2.0 * right_touch_err
                - 2.0 * bottom_touch_err
                - 0.5 * nest_pen
            )
            f1_used = (ang_1 - ang_0) / span if span > 1e-6 else 0.5
            best = (
                score,
                (p1, p2),
                (paste1, paste2),
                (ox, oy, rr),
                (ang_h, ang_0, ang_1, ang_4),
                float(f1_used),
                top2_sil,
                tr,
                br,
                origin1_w,
                origin2_w,
                (gap_l, gap_r, u1_tl_w, u1_tr_w),
            )

    if best is None:
        paste1_x = 0.5 * (
            u0_top_right_w[0]
            + u2_top_left_w[0]
            - u1_top_left_local[0]
            - u1_top_right_local[0]
        )
        paste1 = _paste_by_top(
            (
                paste1_x + u1_top_local[0],
                by + max(40.0, max(0.0, p2[1] - by) * 0.45),
            ),
            u1_top_local,
        )
        # 上式は top 指定なので paste を直指定に補正
        paste1 = (
            paste1_x,
            (by + max(40.0, max(0.0, p2[1] - by) * 0.45)) - u1_top_local[1],
        )
        origin1_w = (paste1[0] + u1_origin[0], paste1[1] + u1_origin[1])
        p1 = (paste1[0] + u1_top_local[0], paste1[1] + u1_top_local[1])
        u1_tl_w = (
            paste1[0] + u1_top_left_local[0],
            paste1[1] + u1_top_left_local[1],
        )
        u1_tr_w = (
            paste1[0] + u1_top_right_local[0],
            paste1[1] + u1_top_right_local[1],
        )
        gap_l = u1_tl_w[0] - u0_top_right_w[0]
        gap_r = u2_top_left_w[0] - u1_tr_w[0]
        arc_meta: Dict[str, Any] = {
            "fallback": True,
            "unit0TopFrameTop": True,
            "unit1TopOnFanCircle": True,
            "unit1TopEdgeEqualGap": True,
            "unit2TopRightEdge": True,
            "unit2BottomRightFrameBottom": True,
            "originDefJa": "商品底辺の左端",
            "unit2BottomRight": {"x": round(br[0], 1), "y": round(br[1], 1)},
            "unit2TopBandRight": {"x": round(tr[0], 1), "y": round(tr[1], 1)},
            "unit1TopGaps": {"left": round(gap_l, 1), "right": round(gap_r, 1)},
        }
    else:
        (
            _score,
            (p1, p2),
            (paste1, paste2),
            cycrr,
            angs,
            f1_used,
            top2_sil,
            tr,
            br,
            origin1_w,
            origin2_w,
            gap_pack,
        ) = best
        gap_l, gap_r, u1_tl_w, u1_tr_w = gap_pack
        arc_meta = {
            "method": "u0_top_u1_equal_gap_circle_u2_right_bottom",
            "cx": round(cycrr[0], 1),
            "cy": round(cycrr[1], 1),
            "radius": round(cycrr[2], 1),
            "unit1FracOnArc": round(float(f1_used), 3),
            "unit0TopFrameTop": True,
            "unit1TopOnFanCircle": True,
            "unit1TopEdgeEqualGap": True,
            "unit2TopRightEdge": True,
            "unit2BottomRightFrameBottom": True,
            "originDefJa": "商品底辺の左端",
            "unit2TopBandRight": {"x": round(tr[0], 1), "y": round(tr[1], 1)},
            "unit2BottomRight": {"x": round(br[0], 1), "y": round(br[1], 1)},
            "unit2SilhouetteTop": {
                "x": round(top2_sil[0], 1),
                "y": round(top2_sil[1], 1),
            },
            "unit1TopGaps": {
                "toUnit0Right": round(float(gap_l), 1),
                "toUnit2Left": round(float(gap_r), 1),
                "equal": abs(float(gap_l) - float(gap_r)) < 2.0,
            },
            "unit1TopEdge": {
                "left": {"x": round(u1_tl_w[0], 1), "y": round(u1_tl_w[1], 1)},
                "right": {"x": round(u1_tr_w[0], 1), "y": round(u1_tr_w[1], 1)},
            },
            "unit0TopRight": {
                "x": round(u0_top_right_w[0], 1),
                "y": round(u0_top_right_w[1], 1),
            },
            "unit2TopLeft": {
                "x": round(u2_top_left_w[0], 1),
                "y": round(u2_top_left_w[1], 1),
            },
            "anglesDegCwFromUp": [round(a, 3) for a in angs],
            "score": round(float(_score), 4),
            "fanTops": [
                {
                    "role": "unit0",
                    "x": round(bx, 1),
                    "y": round(by, 1),
                    "rule": "frame_top",
                },
                {
                    "role": "unit1",
                    "x": round(p1[0], 1),
                    "y": round(p1[1], 1),
                    "rule": "on_circle_equal_gap_x",
                },
                {
                    "role": "unit2",
                    "x": round(p2[0], 1),
                    "y": round(p2[1], 1),
                    "rule": "frame_right",
                },
            ],
            "fanOrigins": [
                {
                    "role": "unit0",
                    "x": round(unit0_origin_w[0], 1),
                    "y": round(unit0_origin_w[1], 1),
                    "def": "bottom_left",
                },
                {
                    "role": "unit1",
                    "x": round(origin1_w[0], 1),
                    "y": round(origin1_w[1], 1),
                    "def": "bottom_left",
                },
                {
                    "role": "unit2",
                    "x": round(origin2_w[0], 1),
                    "y": round(origin2_w[1], 1),
                    "def": "bottom_left",
                },
            ],
        }

    pastes = [
        (hero_paste_x, hero_paste_y),
        (unit0_paste_x, unit0_paste_y),
        paste1,
        paste2,
    ]
    layers_aabb = [
        (hl, ht0, hr, hb),
        (u0l, u0t0, u0r, u0b),
        u1_aabb,
        u2_aabb,
    ]
    tops_local = [h_top, u0_top, u1_top_local, u2_top_local]
    scs = [sh, sh, sh, sh]

    plans: List[PortraitPlan] = []
    for i in range(3, -1, -1):
        role = "hero" if i == 0 else "unit"
        z = 1000 if i == 0 else (10 * (3 - i))
        px, py = pastes[i]
        plans.append(
            PortraitPlan(
                role,
                int(round(px)),
                int(round(py)),
                scs[i],
                z=z,
                rotation_deg=tilts[i],
                anchor="topleft",
            )
        )

    tops_ordered = []
    aabb_ordered = []
    for i in range(4):
        px, py = pastes[i]
        tops_ordered.append(
            {
                "role": ("hero" if i == 0 else "unit%d" % (i - 1)),
                "x": int(round(px + tops_local[i][0])),
                "y": int(round(py + tops_local[i][1])),
            }
        )
        L, T, R, B = layers_aabb[i]
        aabb_ordered.append(
            {
                "role": ("hero" if i == 0 else "unit%d" % (i - 1)),
                "left": round(px + L, 1),
                "top": round(py + T, 1),
                "right": round(px + R, 1),
                "bottom": round(py + B, 1),
            }
        )

    logic_ja = [
        "hero: 枠縦90%・-15°・下辺+左辺接触・はみ出し禁止",
        "起点定義: 商品底辺の左端（最下帯の最左点）",
        "2本目高さ: トップが枠上辺に接する",
        "2本目起点: hero縦中心よりやや左（背面）",
        "4本目: 頂帯最右が枠右辺に接する",
        "4本目: 底辺の右縁が枠下辺に接する",
        "扇円: 2本目トップ（頂点）と4本目トップを通る円",
        "3本目高さ: その円弧上にトップを載せて決める",
        "3本目横: 上部辺の左右端が、2本目右端・4本目左端と等距離",
        "起点: 2→3→4でヒーロー背面に上→中→下の段差",
    ]

    meta: Dict[str, Any] = {
        "layoutFamily": "portrait_n4_aabb4_fan",
        "pattern": "n4_u0_top_u1_circle_u2_right",
        "n": 4,
        "scale": scs[0],
        "scaleHero": scs[0],
        "scales": scs,
        "heroHeightFill": float(hero_height_fill),
        "tiltsDegCw": [round(t, 2) for t in tilts],
        "tops": tops_ordered,
        "aabb4": aabb_ordered,
        "heroWorld": {k: round(float(v), 1) for k, v in hero_world.items()},
        "heroOverflowForbidden": True,
        "heroOverflow": bool(hero_overflow),
        "heroLeftEdgeTouch": True,
        "heroBottomEdgeTouch": True,
        "unit0TopEdgeTouch": True,
        "originDefJa": "商品底辺の左端",
        "unit0OriginLeftOfHeroCenterPx": float(unit0_foot_left_of_hero_center_px),
        "unit0OriginTargetX": round(float(origin_target_x), 1),
        "unit0OriginWorld": {
            "x": round(unit0_origin_w[0], 1),
            "y": round(unit0_origin_w[1], 1),
        },
        "heroCenterX": round(float(hero_cx), 1),
        "topArc": arc_meta,
        "logicJa": logic_ja,
        "anchor": "topleft",
        "allowEdgeCrop": False,
        "allowBottomProtrude": True,
        "zOrderJa": "右奥→…→hero最前",
        "proposalJa": " / ".join(logic_ja),
        "canvasStillSquare": True,
        "sampleAlignJa": "見本128合格固定: 2本目枠上・3本目等距離+円弧・4本目右+下",
        "status": "locked_pass",
        "lockedAt": "2026-08-06",
        "lockedIter": "v3_iter31",
        "productWhRef": [product_w, product_h],
        "ruleStepsJa": logic_ja,
    }
    return plans, meta


def propose_n5plus_tilted_stack_plans(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_tilt_deg_cw: float = 0.0,
    unit_tilt_deg_cw: float = 30.0,
    overlap: float = 0.32,
    scale: Optional[float] = None,
    hero_rgba: Optional[Any] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    """
    N≥5 縦長: 正方形 edge_fill と同型（hero左・unit右列積み）。
    unit は約30°傾けて積む。
    hero は直立・上下枠接触・左枠〜unit左端との左右余白均等。
    """
    from edge_layout import (
        N5_UNIT_MIN_WIDTH_RATIO,
        max_unit_scale_right_col,
        unit_column_counts,
    )

    n = int(n)
    if n < 5:
        raise ValueError("propose_n5plus_tilted_stack_plans requires n>=5")

    margin = max(8, int(canvas * EDGE_MARGIN_RATIO))
    # hero は直立固定（設定値があっても 0 に強制）
    ht = 0.0
    ut = abs(float(unit_tilt_deg_cw))

    # 傾け後 AABB を「仮想 product 寸法」として正方形と同ロジックで unit スケール決定
    u_aabb_w, u_aabb_h = aabb_after_rotate(product_w, product_h, ut)

    scale_cap, ov_col, counts, h_ov = max_unit_scale_right_col(
        n=n,
        canvas=canvas,
        product_w=u_aabb_w,
        product_h=u_aabb_h,
        margin=margin,
        overlap=overlap,
    )
    if scale is None:
        scale_u = float(scale_cap)
    else:
        scale_u = max(0.05, min(float(scale), float(scale_cap) * 1.02))

    num_u = n - 1
    if not counts:
        counts = unit_column_counts(num_u)

    # hero: 直立・高さ＝枠いっぱい（上辺・下辺が枠に接する）
    # 不透明AABBがあればそれで高さ合わせ、なければ product_h
    if isinstance(hero_rgba, Image.Image):
        # 不透明AABB高さを枠高に合わせる
        probe = hero_rgba.convert("RGBA")
        a = probe.getchannel("A")
        bb = a.getbbox() or (0, 0, probe.size[0], probe.size[1])
        opaque_h = max(1, bb[3] - bb[1])
        opaque_w = max(1, bb[2] - bb[0])
        scale_h = float(canvas) / float(opaque_h)
        hw = max(1, int(round(opaque_w * scale_h)))
        hh = canvas
        # 貼付後の不透明上端が y=0 になるようオフ
        paste_oy = -int(round(bb[1] * scale_h))
        paste_ox_off = -int(round(bb[0] * scale_h))
    else:
        scale_h = float(canvas) / float(max(1, product_h))
        hw = max(1, int(round(product_w * scale_h)))
        hh = canvas
        paste_oy = 0
        paste_ox_off = 0

    min_side_gap = 4.0

    def _place_units(su: float) -> Tuple[List[PortraitPlan], float, float, int, int]:
        uw = max(1, int(round(u_aabb_w * su)))
        uh = max(1, int(round(u_aabb_h * su)))
        step_x = max(1, int(round(uw * (1.0 - h_ov))))
        max_rows = max(counts) if counts else 1
        y_top = margin
        y_bot = canvas - margin - uh
        step_y = 0.0 if max_rows <= 1 else max(0.0, (y_bot - y_top) / (max_rows - 1))
        unit_plans: List[PortraitPlan] = []
        leftmost = canvas
        for col_i, nrows in enumerate(counts):
            x_u = canvas - margin - uw - col_i * step_x
            x_u = max(margin, x_u)
            leftmost = min(leftmost, x_u)
            for r in range(nrows):
                y = int(round(y_bot - r * step_y))
                y = max(margin, min(y, canvas - margin - uh))
                z = (len(counts) - 1 - col_i) * 100 + (nrows - 1 - r)
                unit_plans.append(
                    PortraitPlan(
                        "unit",
                        x_u,
                        y,
                        su,
                        z=z,
                        rotation_deg=ut,
                        anchor="topleft",
                    )
                )
        return unit_plans, float(leftmost), step_y, uw, uh

    # unit 列が hero 全高幅＋左右均等余白を残すまで縮小
    unit_plans, leftmost_x, step_y, uw, uh = _place_units(scale_u)
    need_left = float(hw) + 2.0 * min_side_gap
    shrink_guard = 0
    while leftmost_x < need_left and scale_u > 0.06 and shrink_guard < 24:
        scale_u *= 0.92
        unit_plans, leftmost_x, step_y, uw, uh = _place_units(scale_u)
        shrink_guard += 1

    # 左右余白均等: gap_L = hero不透明左 - 0, gap_R = unit左 - hero不透明右
    # paste_x + paste_ox_off = gap, paste_x + paste_ox_off + hw = leftmost - gap
    # => gap = (leftmost_x - hw) / 2
    gap_side = 0.5 * (float(leftmost_x) - float(hw))
    opaque_left = gap_side
    paste_x = opaque_left - float(paste_ox_off)
    paste_y = float(paste_oy)
    # 枠上・下に不透明が接するよう y は paste_oy（不透明上=0）。下端は scale で合わせ済み。

    plans: List[PortraitPlan] = list(unit_plans)
    plans.append(
        PortraitPlan(
            "hero",
            int(round(paste_x)),
            int(round(paste_y)),
            scale_h,
            z=1000,
            rotation_deg=ht,
            anchor="topleft",
        )
    )

    for p in plans:
        p.x = max(-canvas, min(p.x, canvas - 1))
        p.y = max(-canvas, min(p.y, canvas - 1))

    gap_l = float(opaque_left)
    gap_r = float(leftmost_x) - (float(opaque_left) + float(hw))
    logic_ja = [
        "N>=5: 正方形型と同じ hero左＋右unit列積み",
        f"unit傾き ≈{ut:.0f}°（縦長のため傾けて積む）",
        "hero: 直立・上辺=枠上・下辺=枠下",
        "hero横: 左枠〜hero左端 と hero右端〜unit左端 が等距離",
        "列分割: <10=1列 / 10-19=2列 / >=20=10+10+端数",
    ]
    meta: Dict[str, Any] = {
        "layoutFamily": "portrait_tilted_right_stack",
        "pattern": "n5plus_units_right_cols_hero_left_tilted",
        "status": "locked_pass",
        "lockedAt": "2026-08-06",
        "lockedIter": "v2_stack",
        "n": n,
        "edgeMarginRatio": EDGE_MARGIN_RATIO,
        "scale": scale_u,
        "scaleMax": float(scale_cap),
        "scaleUnit": scale_u,
        "scaleHero": scale_h,
        "tiltsDegCw": [round(ht, 2)] + [round(ut, 2)] * num_u,
        "heroTiltDegCw": round(ht, 2),
        "unitTiltDegCw": round(ut, 2),
        "unitColOverlapUsed": ov_col,
        "unitColHOverlapUsed": h_ov,
        "unitColumnCounts": counts,
        "unitNumColumns": len(counts),
        "unitStepYPx": round(step_y, 2),
        "unitStackFrom": "bottom",
        "unitWidthRatio": round(uw / float(canvas), 4),
        "unitMinWidthRatio": N5_UNIT_MIN_WIDTH_RATIO,
        "leftmostUnitX": round(float(leftmost_x), 1),
        "heroOpaqueWh": [hw, hh],
        "heroSideGaps": {
            "left": round(gap_l, 1),
            "right": round(gap_r, 1),
            "equal": abs(gap_l - gap_r) < 2.0,
        },
        "heroTopBottomTouch": True,
        "heroUpright": True,
        "unitAabb1": [u_aabb_w, u_aabb_h],
        "logicJa": logic_ja,
        "zOrderJa": "右奥unit→…→下・右寄り手前、hero左=最前",
        "proposalJa": " / ".join(logic_ja),
        "canvasStillSquare": True,
        "anchor": "topleft",
        "allowEdgeCrop": False,
        "basePolicyJa": "01.amazon白抜きベース（透過PNG）を hero=unit に使用",
    }
    return plans, meta


def propose_portrait_plans(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_tilt_deg_cw: float = DEFAULT_HERO_TILT_DEG_CW,
    unit_tilt_deg_cw: float = DEFAULT_UNIT_TILT_DEG_CW,
    unit_tilt_step_deg_cw: float = DEFAULT_UNIT_TILT_STEP_DEG_CW,
    n1_tilt_deg_cw: float = DEFAULT_N1_TILT_DEG_CW,
    h_overlap: float = DEFAULT_H_OVERLAP,
    v_overlap: float = DEFAULT_V_OVERLAP,
    scale: Optional[float] = None,
    tilt_deg_cw: Optional[float] = None,
    n2_cfg: Optional[Dict[str, Any]] = None,
    n3_cfg: Optional[Dict[str, Any]] = None,
    n4_cfg: Optional[Dict[str, Any]] = None,
    n5plus_cfg: Optional[Dict[str, Any]] = None,
    hero_rgba: Optional[Any] = None,
    unit_rgba: Optional[Any] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any]]:
    n = int(n)
    if n < 1:
        raise ValueError("n>=1")

    if n >= 5:
        cfg = dict(n5plus_cfg or {})
        return propose_n5plus_tilted_stack_plans(
            n=n,
            canvas=canvas,
            product_w=product_w,
            product_h=product_h,
            hero_tilt_deg_cw=0.0,
            unit_tilt_deg_cw=float(cfg.get("unitTiltDegCw", 30)),
            overlap=float(cfg.get("overlap", 0.32)),
            scale=scale,
            hero_rgba=hero_rgba if isinstance(hero_rgba, Image.Image) else None,
        )

    if n == 4:
        cfg = dict(n4_cfg or {})
        engine = str(cfg.get("layoutEngine") or "upright_quad_fan").strip().lower()
        slender_hw = max(1.0, float(cfg.get("slenderMinHw", 1.75)))
        # 直立素材の H/W で細長判定（缶・瓶 vs パウチ）
        src_hw = None
        src_im = None
        if isinstance(unit_rgba, Image.Image):
            src_im = unit_rgba.convert("RGBA")
        elif isinstance(hero_rgba, Image.Image):
            src_im = hero_rgba.convert("RGBA")
        if src_im is not None:
            bb = src_im.split()[-1].getbbox()
            if bb:
                sw = max(1, int(bb[2] - bb[0]))
                sh = max(1, int(bb[3] - bb[1]))
                src_hw = float(sh) / float(sw)
        is_slender = src_hw is not None and src_hw + 1e-9 >= slender_hw
        slender_engine = str(
            cfg.get("slenderLayoutEngine") or "n4_legacy"
        ).strip().lower()
        # 細長: 添付相当の頂円弧・高さ扇（n4_legacy）。幅広: upright_quad のまま
        use_legacy = engine in ("legacy", "n4_legacy", "height_fan", "fan") or (
            is_slender and slender_engine in ("legacy", "n4_legacy", "height_fan", "fan", "top_arc")
        )
        if use_legacy and (engine in ("legacy", "n4_legacy", "height_fan", "fan") or is_slender):
            # 細長時の高さ埋め（既定0.92）。幅広で明示 legacy 指定時は heroHeightFill
            if is_slender:
                h_fill = float(
                    cfg.get(
                        "heroHeightFillMaxSlender",
                        cfg.get("heroHeightFill", 0.92),
                    )
                )
            else:
                h_fill = float(cfg.get("heroHeightFill", 0.90))
            plans, meta = propose_n4_fan_plans(
                canvas=canvas,
                product_w=product_w,
                product_h=product_h,
                hero_tilt_deg_cw=float(
                    cfg.get(
                        "slenderHeroTiltDegCw",
                        cfg.get("heroTiltDegCw", -15),
                    )
                ),
                unit0_tilt_deg_cw=float(
                    cfg.get(
                        "slenderUnit0TiltDegCw",
                        cfg.get("unit0TiltDegCw", 15),
                    )
                ),
                unit1_tilt_deg_cw=float(
                    cfg.get(
                        "slenderUnit1TiltDegCw",
                        cfg.get("unit1TiltDegCw", 35),
                    )
                ),
                unit2_tilt_deg_cw=float(
                    cfg.get(
                        "slenderUnit2TiltDegCw",
                        cfg.get("unit2TiltDegCw", 55),
                    )
                ),
                hero_height_fill=h_fill,
                unit0_foot_left_of_hero_center_px=float(
                    cfg.get("unit0FootLeftOfHeroCenterPx", 40)
                ),
                hero_rgba=hero_rgba if isinstance(hero_rgba, Image.Image) else None,
                unit_rgba=unit_rgba if isinstance(unit_rgba, Image.Image) else None,
                hero_scale=float(cfg.get("heroScale", 0.85)),
                unit0_scale_ratio=float(cfg.get("unit0ScaleRatio", 1.0)),
                unit1_scale_ratio=float(cfg.get("unit1ScaleRatio", 1.0)),
                unit2_scale_ratio=float(cfg.get("unit2ScaleRatio", 1.0)),
            )
            meta = dict(meta)
            meta["productHw"] = round(float(src_hw), 4) if src_hw is not None else None
            meta["slenderMinHw"] = slender_hw
            meta["isSlender"] = bool(is_slender)
            meta["layoutEngineResolved"] = "n4_legacy"
            meta["branchJa"] = (
                "細長→n4_legacy（頂円弧・高さ扇／添付相当）"
                if is_slender
                else "明示legacy"
            )
            if is_slender:
                meta["pattern"] = "n4_legacy_slender_top_arc"
                meta["heroHeightFillMaxSlender"] = h_fill
            return plans, meta
        return propose_n4_upright_quad_fan_plans(
            canvas=canvas,
            product_w=product_w,
            product_h=product_h,
            hero_tilt_deg_cw=float(cfg.get("heroTiltDegCw", -14)),
            unit0_tilt_deg_cw=float(cfg.get("unit0TiltDegCw", 18)),
            unit1_tilt_deg_cw=float(cfg.get("unit1TiltDegCw", 36)),
            unit2_tilt_deg_cw=float(cfg.get("unit2TiltDegCw", 60)),
            unit2_right_overflow_ratio=float(cfg.get("unit2RightOverflowRatio", 0.15)),
            unit2_bottom_overflow_ratio=float(cfg.get("unit2BottomOverflowRatio", 0.0)),
            slender_min_hw=slender_hw,
            hero_height_fill_max_slender=float(
                cfg.get("heroHeightFillMaxSlender", 0.92)
            ),
            slender_unit1_top_mode=str(
                cfg.get("slenderUnit1TopMode", "perp_mid")
            ),
            scale=scale,
            hero_rgba=hero_rgba if isinstance(hero_rgba, Image.Image) else None,
            unit_rgba=unit_rgba if isinstance(unit_rgba, Image.Image) else None,
        )

    if n == 3:
        cfg = dict(n3_cfg or {})
        engine = str(cfg.get("layoutEngine") or "upright_quad_fan").strip().lower()
        slender_hw = max(1.0, float(cfg.get("slenderMinHw", 1.75)))
        src_hw = None
        src_im = None
        if isinstance(unit_rgba, Image.Image):
            src_im = unit_rgba.convert("RGBA")
        elif isinstance(hero_rgba, Image.Image):
            src_im = hero_rgba.convert("RGBA")
        if src_im is not None:
            bb = src_im.split()[-1].getbbox()
            if bb:
                sw = max(1, int(bb[2] - bb[0]))
                sh0 = max(1, int(bb[3] - bb[1]))
                src_hw = float(sh0) / float(sw)
        is_slender = src_hw is not None and src_hw + 1e-9 >= slender_hw
        slender_engine = str(
            cfg.get("slenderLayoutEngine") or "top_arc_legacy"
        ).strip().lower()
        use_top_arc = engine in ("top_arc", "top_arc_legacy", "legacy") or (
            is_slender
            and slender_engine in ("top_arc", "top_arc_legacy", "legacy", "n3_legacy")
        )
        if use_top_arc and (
            engine in ("top_arc", "top_arc_legacy", "legacy") or is_slender
        ):
            ht = float(
                cfg.get(
                    "slenderHeroTiltDegCw",
                    cfg.get("heroTiltDegCw", -14),
                )
            )
            if ht > 0:
                ht = -abs(ht)
            u0t = abs(
                float(
                    cfg.get(
                        "slenderUnit0TiltDegCw",
                        cfg.get("unit0TiltDegCw", 18),
                    )
                )
            )
            u1t = abs(
                float(
                    cfg.get(
                        "slenderUnit1TiltDegCw",
                        cfg.get("unit1TiltDegCw", 38),
                    )
                )
            )
            # N=4細長と同型: 高さ埋め上限で少し小さく（既定0.92）
            h_fill0 = float(
                cfg.get(
                    "heroHeightFillMaxSlender",
                    cfg.get("heroHeightFill", 0.92),
                )
            )
            h_fill0 = max(0.45, min(0.98, h_fill0))
            u0_ratio = float(
                cfg.get("slenderUnit0ScaleRatio", cfg.get("unit0ScaleRatio", 1.0))
            )
            u1_ratio = float(
                cfg.get("slenderUnit1ScaleRatio", cfg.get("unit1ScaleRatio", 1.0))
            )
            if is_slender:
                u0_ratio = float(cfg.get("slenderUnit0ScaleRatio", 1.0))
                u1_ratio = float(cfg.get("slenderUnit1ScaleRatio", 1.0))
            pin_bottom = bool(cfg.get("pinHeroFrameBottom", True))
            hero_inside = bool(cfg.get("heroInsideFrame", True))
            # 既定: 上辺左右中点を円弧中点の法線上へ。旧 AABB 中点は明示時のみ
            top_edge_normal = bool(cfg.get("unit0TopEdgeMidOnArcNormal", True))
            mid_center = bool(cfg.get("unit0AabbCenterMidX", False)) and (
                not top_edge_normal
            )
            src_h = hero_rgba if isinstance(hero_rgba, Image.Image) else None
            src_u = unit_rgba if isinstance(unit_rgba, Image.Image) else None

            h_fill = h_fill0
            hero_scale = float(cfg.get("heroScale", 0.92))
            plans: List[PortraitPlan] = []
            meta: Dict[str, Any] = {}
            pin_meta: Dict[str, Any] = {}
            inside_meta: Dict[str, Any] = {}
            mid_meta: Dict[str, Any] = {}
            scale_tries = 0
            for _try in range(14):
                scale_tries = _try + 1
                src_for_scale = src_im if src_im is not None else None
                if src_for_scale is not None:
                    probe = _prepare_scaled_rotated_layer(
                        product_w=product_w,
                        product_h=product_h,
                        scale=1.0,
                        tilt_deg_cw=ht,
                        rgba=src_for_scale,
                    )
                    _l, _t, _r, b1 = measure_aabb4(probe)
                    h1 = max(1, b1 - _t + 1)
                    hero_scale = max(
                        0.05, min(1.25, (float(canvas) * h_fill) / float(h1))
                    )
                plans, meta = propose_n3_fan_plans(
                    canvas=canvas,
                    product_w=product_w,
                    product_h=product_h,
                    hero_tilt_deg_cw=ht,
                    unit0_tilt_deg_cw=u0t,
                    unit1_tilt_deg_cw=u1t,
                    hero_scale=hero_scale,
                    unit0_scale_ratio=u0_ratio,
                    unit1_scale_ratio=u1_ratio,
                    top_arc_cx_ratio=float(cfg.get("topArcCxRatio", 0.1304)),
                    top_arc_cy_ratio=float(cfg.get("topArcCyRatio", 2.777)),
                    top_arc_radius_ratio=float(cfg.get("topArcRadiusRatio", 2.794)),
                    top_arc_angles_deg=(
                        list(cfg["topArcAnglesDeg"])
                        if isinstance(cfg.get("topArcAnglesDeg"), list)
                        else None
                    ),
                    bottom_protrude_px=float(cfg.get("bottomProtrudePx", 36)),
                    unit1_top_down_px=float(cfg.get("unit1TopDownPx", 55)),
                )
                meta = dict(meta)
                if pin_bottom:
                    plans, pin_meta = pin_plans_hero_aabb_to_frame_bottom(
                        plans,
                        canvas=canvas,
                        product_w=product_w,
                        product_h=product_h,
                        hero_rgba=src_h,
                        unit_rgba=src_u,
                    )
                    meta["heroFrameBottomPin"] = pin_meta
                if hero_inside:
                    plans, inside_meta = fit_plans_hero_inside_frame(
                        plans,
                        canvas=canvas,
                        product_w=product_w,
                        product_h=product_h,
                        hero_rgba=src_h,
                        unit_rgba=src_u,
                    )
                    meta["heroInsideFrameFit"] = inside_meta
                    if inside_meta.get("inside"):
                        break
                    # 縮小して再提案（下辺接地＋枠内を両立）
                    h_fill = max(0.45, h_fill * 0.96)
                    continue
                break

            # tops を最終プランの頂へ同期
            tops_sync: List[Dict[str, Any]] = []
            for p in sorted(
                plans,
                key=lambda q: (
                    0 if str(q.role) == "hero" else 1,
                    float(q.rotation_deg),
                ),
            ):
                role = "hero" if str(p.role) == "hero" else (
                    "unit0" if abs(float(p.rotation_deg) - u0t) < 1e-6 else "unit1"
                )
                tops_sync.append(
                    {
                        "role": role,
                        "x": int(p.top_x if p.top_x is not None else p.x),
                        "y": int(p.top_y if p.top_y is not None else p.y),
                    }
                )
            meta["tops"] = tops_sync
            if pin_meta:
                ta = dict(meta.get("topArc") or {})
                if ta and pin_meta.get("dy") is not None:
                    ta["pinHeroFrameBottomDy"] = pin_meta.get("dy")
                    meta["topArc"] = ta

            if top_edge_normal:
                plans, edge_meta = nudge_n3_unit0_top_edge_mid_on_arc_normal(
                    plans,
                    product_w=product_w,
                    product_h=product_h,
                    hero_rgba=src_h,
                    unit_rgba=src_u,
                )
                meta["unit0TopEdgeMidOnArcNormal"] = edge_meta
                tops2 = list(meta.get("tops") or [])
                for t in tops2:
                    if t.get("role") == "unit0":
                        t["x"] = int(edge_meta.get("topXAfter") or t.get("x") or 0)
                        if edge_meta.get("topYAfter") is not None:
                            t["y"] = int(edge_meta["topYAfter"])
                meta["tops"] = tops2
                # Q=M の後、緑法線上で上へ（円弧上は外れる。preview/knob）
                nudge_up = float(cfg.get("unit0NudgeUpAlongNormalPx") or 0.0)
                if abs(nudge_up) > 1e-9:
                    plans, up_meta = nudge_n3_unit0_up_along_arc_normal(
                        plans,
                        product_w=product_w,
                        product_h=product_h,
                        nudge_up_px=nudge_up,
                        hero_rgba=src_h,
                        unit_rgba=src_u,
                        edge_meta=edge_meta,
                    )
                    meta["unit0NudgeUpAlongNormal"] = up_meta
                    if up_meta.get("applied"):
                        tops3 = list(meta.get("tops") or [])
                        for t in tops3:
                            if t.get("role") == "unit0":
                                t["x"] = int(up_meta.get("topXAfter") or t.get("x") or 0)
                                t["y"] = int(up_meta.get("topYAfter") or t.get("y") or 0)
                        meta["tops"] = tops3
            elif mid_center:
                plans, mid_meta = nudge_n3_unit0_aabb_center_mid_x(
                    plans,
                    product_w=product_w,
                    product_h=product_h,
                    hero_rgba=src_h,
                    unit_rgba=src_u,
                )
                meta["unit0AabbCenterMidX"] = mid_meta
                # unit0 頂Xだけ更新
                tops2 = list(meta.get("tops") or [])
                for t in tops2:
                    if t.get("role") == "unit0":
                        t["x"] = int(mid_meta.get("topXAfter") or t.get("x") or 0)
                meta["tops"] = tops2

            meta["productHw"] = round(float(src_hw), 4) if src_hw is not None else None
            meta["slenderMinHw"] = slender_hw
            meta["isSlender"] = bool(is_slender)
            meta["layoutEngineResolved"] = "top_arc_legacy"
            meta["scale"] = hero_scale
            meta["scaleHero"] = hero_scale
            meta["heroHeightFillMaxSlender"] = h_fill if is_slender else None
            meta["heroHeightFillRequested"] = h_fill0
            meta["heroInsideScaleTries"] = scale_tries
            meta["branchJa"] = (
                "細長→top_arc_legacy（hero枠内＋下辺接地・unit1上辺左右中点＝円弧法線）"
                if is_slender
                else "明示top_arc_legacy（hero枠内＋下辺接地・unit1上辺左右中点＝円弧法線）"
            )
            if is_slender:
                meta["pattern"] = "n3_top_arc_fan_slender"
            return plans, meta
        return propose_n3_upright_quad_fan_plans(
            canvas=canvas,
            product_w=product_w,
            product_h=product_h,
            hero_tilt_deg_cw=float(cfg.get("heroTiltDegCw", -14)),
            unit0_tilt_deg_cw=float(cfg.get("unit0TiltDegCw", 18)),
            unit1_tilt_deg_cw=float(cfg.get("unit1TiltDegCw", 50)),
            unit1_right_overflow_ratio=float(cfg.get("unit1RightOverflowRatio", 0.10)),
            scale=scale,
            hero_rgba=hero_rgba if isinstance(hero_rgba, Image.Image) else None,
            unit_rgba=unit_rgba if isinstance(unit_rgba, Image.Image) else None,
        )

    if n == 2:
        cfg = dict(n2_cfg or {})
        ht = float(cfg.get("heroTiltDegCw", hero_tilt_deg_cw))
        if ht > 0:
            ht = -abs(ht)
        ut = float(cfg.get("unitTiltDegCw", unit_tilt_deg_cw))
        if ut < 0:
            ut = abs(ut)
        return propose_n2_noji_plans(
            canvas=canvas,
            product_w=product_w,
            product_h=product_h,
            hero_tilt_deg_cw=ht,
            unit_tilt_deg_cw=ut,
            hero_height_fill=float(cfg.get("heroHeightFill", 1.10)),
            hero_max_width_ratio=float(cfg.get("heroMaxWidthRatio", 0.92)),
            unit_scale_ratio=float(cfg.get("unitScaleRatio", 0.97)),
            shared_foot_y_ratio=float(cfg.get("sharedFootYRatio", 0.9992)),
            hero_foot_x_ratio=float(cfg.get("heroFootXRatio", 0.4167)),
            unit_foot_x_ratio=float(cfg.get("unitFootXRatio", 0.5500)),
            unit_base_crop_ratio=float(cfg.get("unitBaseCropRatio", 0.03)),
            unit_overlap_x_ratio=float(cfg.get("unitOverlapXRatio", 0.42)),
            unit_down_shift_ratio=float(cfg.get("unitDownShiftRatio", 0.05)),
            hero_left_bias_ratio=float(cfg.get("heroLeftBiasRatio", 0.08)),
            hero_scale=None,
            scale=scale,
            hero_rgba=hero_rgba if isinstance(hero_rgba, Image.Image) else None,
            unit_rgba=unit_rgba if isinstance(unit_rgba, Image.Image) else None,
        )

    margin = max(8, int(canvas * EDGE_MARGIN_RATIO))
    tilts = tilts_for_n(
        n,
        n1_tilt_deg_cw=n1_tilt_deg_cw,
        hero_tilt_deg_cw=hero_tilt_deg_cw,
        unit_tilt_deg_cw=unit_tilt_deg_cw,
        unit_tilt_step_deg_cw=unit_tilt_step_deg_cw,
    )
    sizes1 = [aabb_after_rotate(product_w, product_h, t) for t in tilts]
    scale_max = _fit_scale_for_cluster(
        n=n,
        canvas=canvas,
        sizes=sizes1,
        margin=margin,
        h_ov=h_overlap,
        v_ov=v_overlap,
    )
    if scale is None:
        scale = scale_max
    else:
        scale = max(0.05, min(float(scale), scale_max * 1.02))

    sizes = [
        (max(1, int(round(w * scale))), max(1, int(round(h * scale))))
        for w, h in sizes1
    ]
    rw_ref = max(s[0] for s in sizes)
    rh_ref = max(s[1] for s in sizes)
    offs = _cluster_rel_offsets(n, rw_ref, rh_ref, h_overlap, v_overlap)

    xs: List[float] = []
    ys: List[float] = []
    for i, (ox, oy) in enumerate(offs):
        sw, sh = sizes[i]
        xs.extend([ox, ox + sw])
        ys.extend([oy, oy + sh])
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    cluster_w = max_x - min_x
    cluster_h = max_y - min_y

    origin_x = margin + max(0, (canvas - 2 * margin - cluster_w) // 2) - min_x
    origin_y = margin + max(0, int((canvas - 2 * margin - cluster_h) * 0.55)) - min_y

    plans: List[PortraitPlan] = []
    for i, (ox, oy) in enumerate(offs):
        role = "hero" if i == 0 else "unit"
        z = 1000 if i == 0 else (100 - i)
        x = int(round(origin_x + ox))
        y = int(round(origin_y + oy))
        x = max(0, min(x, canvas - 1))
        y = max(0, min(y, canvas - 1))
        plans.append(
            PortraitPlan(role, x, y, scale, z=z, rotation_deg=tilts[i])
        )

    meta: Dict[str, Any] = {
        "layoutFamily": "portrait_opposing_tilt_fan",
        "pattern": "n1_center_upright" if n == 1 else "n%d_hero_left_unit_opposite" % n,
        "edgeMarginRatio": EDGE_MARGIN_RATIO,
        "n": n,
        "scale": scale,
        "scaleMax": scale_max,
        "tiltsDegCw": [round(t, 2) for t in tilts],
        "heroTiltDegCw": tilts[0] if tilts else 0,
        "unitTiltDegCw": tilts[1] if len(tilts) > 1 else None,
        "hOverlap": h_overlap,
        "vOverlap": v_overlap,
        "aabbScaled": [[s[0], s[1]] for s in sizes],
        "clusterWh": [round(cluster_w, 1), round(cluster_h, 1)],
        "zOrderJa": "左前=hero最前面、右へunit",
        "proposalJa": (
            "N=1: 直立中央最大化。"
            if n == 1
            else "N>=3: hero左傾き・unitは逆向き。扇状。"
        ),
        "canvasStillSquare": True,
        "basePolicyJa": "01.amazon白抜きベース（透過PNG）を hero=unit に使用",
    }
    return plans, meta


def fit_portrait_layout_under_overlap(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_rgba: Any,
    unit_rgba: Any,
    pair_overlap_max: float = 0.45,
    hero_visible_min: float = 0.65,
    hero_tilt_deg_cw: float = DEFAULT_HERO_TILT_DEG_CW,
    unit_tilt_deg_cw: float = DEFAULT_UNIT_TILT_DEG_CW,
    unit_tilt_step_deg_cw: float = DEFAULT_UNIT_TILT_STEP_DEG_CW,
    n1_tilt_deg_cw: float = DEFAULT_N1_TILT_DEG_CW,
    h_overlap: float = DEFAULT_H_OVERLAP,
    v_overlap: float = DEFAULT_V_OVERLAP,
    iters: int = 16,
    tilt_deg_cw: Optional[float] = None,
    n2_cfg: Optional[Dict[str, Any]] = None,
    n3_cfg: Optional[Dict[str, Any]] = None,
    n4_cfg: Optional[Dict[str, Any]] = None,
    n5plus_cfg: Optional[Dict[str, Any]] = None,
) -> Tuple[List[PortraitPlan], Dict[str, Any], Dict[str, Any]]:
    """
    重なり硬制約を満たす最大スケールを二分探索。
    N=2: hero枠内上限の下で unit可視≥unitVisibleMin となる最大同尺。
    N=3: hero枠内上限の下で unit0可視≥unit0VisibleMin かつ unit1可視≥unit1VisibleMin。
    N=4/≥5: 合格配置を優先し overlap で縮めない。
    """
    from overlap_metrics import measure_plans_overlap

    kw: Dict[str, Any] = dict(
        n=n,
        canvas=canvas,
        product_w=product_w,
        product_h=product_h,
        hero_tilt_deg_cw=hero_tilt_deg_cw,
        unit_tilt_deg_cw=unit_tilt_deg_cw,
        unit_tilt_step_deg_cw=unit_tilt_step_deg_cw,
        n1_tilt_deg_cw=n1_tilt_deg_cw,
        h_overlap=h_overlap,
        v_overlap=v_overlap,
        n2_cfg=n2_cfg,
        n3_cfg=n3_cfg,
        n4_cfg=n4_cfg,
        n5plus_cfg=n5plus_cfg,
        hero_rgba=hero_rgba,
        unit_rgba=unit_rgba,
    )

    def _unit_vis_ok(om: Dict[str, Any], *, mid_min: float, back_min: float) -> Tuple[bool, float, float]:
        rows = list(om.get("unitFrontVisibility") or [])
        rows = sorted(rows, key=lambda r: int(r.get("z") or 0))
        if len(rows) < 2:
            vis = float(om.get("unitMinVisible") or 1.0)
            return vis >= mid_min and vis >= back_min, vis, vis
        back_vis = float(rows[0].get("visible") or 0.0)
        mid_vis = float(rows[1].get("visible") or 0.0)
        return (mid_vis >= mid_min - 1e-6 and back_vis >= back_min - 1e-6), mid_vis, back_vis

    def _unit_vis_ok_n4(
        om: Dict[str, Any], *, u0_min: float, u1_min: float, u2_min: float
    ) -> Tuple[bool, float, float, float]:
        """z昇順: unit2(back), unit1, unit0 → 可視は u0/u1/u2 の順で返す。"""
        rows = list(om.get("unitFrontVisibility") or [])
        rows = sorted(rows, key=lambda r: int(r.get("z") or 0))
        if len(rows) < 3:
            return False, 0.0, 0.0, 0.0
        u2_vis = float(rows[0].get("visible") or 0.0)
        u1_vis = float(rows[1].get("visible") or 0.0)
        u0_vis = float(rows[2].get("visible") or 0.0)
        ok = (
            u0_vis >= u0_min - 1e-6
            and u1_vis >= u1_min - 1e-6
            and u2_vis >= u2_min - 1e-6
        )
        return ok, u0_vis, u1_vis, u2_vis

    # N=2: unit可視面積で同尺最大化
    if int(n) == 2:
        cfg = dict(n2_cfg or {})
        unit_vis_min = float(cfg.get("unitVisibleMin", 0.50))
        covered_max = float(cfg.get("unitCoveredMax", 1.0 - unit_vis_min))
        # 両方指定時は厳しい方（隠れ上限が小さい方）を採用
        covered_limit = min(covered_max, 1.0 - unit_vis_min)
        unit_vis_min_eff = 1.0 - covered_limit
        search_iters = int(cfg.get("scaleSearchIters", max(16, iters)))
        scale_floor = float(cfg.get("scaleFloor", 0.15))

        plans_hi, meta_hi = propose_portrait_plans(**kw, scale=None)
        om_hi = measure_plans_overlap(
            hero=hero_rgba,
            unit=unit_rgba,
            plans=plans_hi,
            canvas_size=canvas,
        )
        scale_cap = float(
            meta_hi.get("scaleCapHeroInside")
            or meta_hi.get("scaleHero")
            or meta_hi.get("scale")
            or 0.5
        )
        covered_hi = float(om_hi.get("pairMaxBackCovered") or 0.0)
        vis_hi = float(om_hi.get("unitMinVisible") or (1.0 - covered_hi))

        if covered_hi <= covered_limit + 1e-6:
            meta_hi["overlapPass"] = True
            meta_hi["bindingConstraint"] = "hero_inside"
            meta_hi["unitVisibleMin"] = unit_vis_min_eff
            meta_hi["unitCoveredMax"] = covered_limit
            meta_hi["overlapNoteJa"] = (
                f"N=2: 枠内最大同尺で unit可視={vis_hi:.3f}≥{unit_vis_min_eff:.3f}"
                "（unit可視制約は非拘束）"
            )
            LOG.info(
                "portrait_fit n=2 binding=hero_inside scale=%.3f unitVis=%.3f",
                scale_cap,
                vis_hi,
            )
            return plans_hi, meta_hi, om_hi

        lo = max(0.05, scale_floor)
        hi = scale_cap
        best_plans, best_meta, best_om = plans_hi, meta_hi, om_hi
        for _ in range(max(4, search_iters)):
            mid = (lo + hi) * 0.5
            plans, meta = propose_portrait_plans(**kw, scale=mid)
            om = measure_plans_overlap(
                hero=hero_rgba,
                unit=unit_rgba,
                plans=plans,
                canvas_size=canvas,
            )
            covered = float(om.get("pairMaxBackCovered") or 0.0)
            if covered <= covered_limit + 1e-6:
                lo = mid
                best_plans, best_meta, best_om = plans, meta, om
            else:
                hi = mid

        best_plans, best_meta = propose_portrait_plans(**kw, scale=lo)
        best_om = measure_plans_overlap(
            hero=hero_rgba,
            unit=unit_rgba,
            plans=best_plans,
            canvas_size=canvas,
        )
        vis = float(
            best_om.get("unitMinVisible")
            or (1.0 - float(best_om.get("pairMaxBackCovered") or 0.0))
        )
        covered = float(best_om.get("pairMaxBackCovered") or 0.0)
        best_meta["overlapPass"] = covered <= covered_limit + 1e-6
        best_meta["bindingConstraint"] = "unit_visible"
        best_meta["unitVisibleMin"] = unit_vis_min_eff
        best_meta["unitCoveredMax"] = covered_limit
        best_meta["overlapNoteJa"] = (
            f"N=2: unit可視≥{unit_vis_min_eff:.3f}で最大同尺"
            f"（scale={float(best_meta.get('scaleHero') or lo):.3f}"
            f" / unitVis={vis:.3f} / covered={covered:.3f}）"
        )
        LOG.info(
            "portrait_fit n=2 binding=unit_visible scale=%.3f unitVis=%.3f covered=%.3f pass=%s",
            best_meta.get("scaleHero") or lo,
            vis,
            covered,
            best_meta.get("overlapPass"),
        )
        return best_plans, best_meta, best_om

    # N=3: unit0(中)≥unit0VisibleMin / unit1(右奥)≥unit1VisibleMin で同尺最大化
    if int(n) == 3:
        cfg = dict(n3_cfg or {})
        engine = str(cfg.get("layoutEngine") or "upright_quad_fan").strip().lower()
        if engine in ("top_arc", "top_arc_legacy", "legacy"):
            plans, meta = propose_portrait_plans(**kw, scale=None)
            om = measure_plans_overlap(
                hero=hero_rgba,
                unit=unit_rgba,
                plans=plans,
                canvas_size=canvas,
            )
            meta["overlapPass"] = True
            meta["overlapNoteJa"] = "N=3 legacy top_arc: overlap探索スキップ"
            return plans, meta, om

        # 細長は propose 側で top_arc_legacy に分岐 → 探索スキップ
        plans_probe, meta_probe = propose_portrait_plans(**kw, scale=None)
        if (
            meta_probe.get("layoutEngineResolved") == "top_arc_legacy"
            or bool(meta_probe.get("isSlender"))
        ):
            om = measure_plans_overlap(
                hero=hero_rgba,
                unit=unit_rgba,
                plans=plans_probe,
                canvas_size=canvas,
            )
            meta_probe["overlapPass"] = True
            meta_probe["overlapNoteJa"] = (
                "N=3 細長→top_arc_legacy（頂円弧）: overlap探索スキップ"
            )
            return plans_probe, meta_probe, om

        mid_min = float(cfg.get("unit0VisibleMin", 0.40))
        back_min = float(cfg.get("unit1VisibleMin", 0.30))
        search_iters = int(cfg.get("scaleSearchIters", max(16, iters)))
        scale_floor = float(cfg.get("scaleFloor", 0.15))

        plans_hi, meta_hi = plans_probe, meta_probe
        om_hi = measure_plans_overlap(
            hero=hero_rgba,
            unit=unit_rgba,
            plans=plans_hi,
            canvas_size=canvas,
        )
        scale_cap = float(
            meta_hi.get("scaleCapHeroInside")
            or meta_hi.get("scaleHero")
            or meta_hi.get("scale")
            or 0.5
        )
        ok_hi, mid_vis_hi, back_vis_hi = _unit_vis_ok(
            om_hi, mid_min=mid_min, back_min=back_min
        )
        if ok_hi:
            meta_hi["overlapPass"] = True
            meta_hi["bindingConstraint"] = "hero_inside"
            meta_hi["unit0VisibleMin"] = mid_min
            meta_hi["unit1VisibleMin"] = back_min
            meta_hi["unit0Visible"] = mid_vis_hi
            meta_hi["unit1Visible"] = back_vis_hi
            meta_hi["overlapNoteJa"] = (
                f"N=3: 枠内最大同尺で unit0可視={mid_vis_hi:.3f}≥{mid_min:.3f} / "
                f"unit1可視={back_vis_hi:.3f}≥{back_min:.3f}（可視制約は非拘束）"
            )
            LOG.info(
                "portrait_fit n=3 binding=hero_inside scale=%.3f u0=%.3f u1=%.3f",
                scale_cap,
                mid_vis_hi,
                back_vis_hi,
            )
            return plans_hi, meta_hi, om_hi

        lo = max(0.05, scale_floor)
        hi = scale_cap
        best_plans, best_meta, best_om = plans_hi, meta_hi, om_hi
        for _ in range(max(4, search_iters)):
            mid = (lo + hi) * 0.5
            plans, meta = propose_portrait_plans(**kw, scale=mid)
            om = measure_plans_overlap(
                hero=hero_rgba,
                unit=unit_rgba,
                plans=plans,
                canvas_size=canvas,
            )
            ok, _, _ = _unit_vis_ok(om, mid_min=mid_min, back_min=back_min)
            if ok:
                lo = mid
                best_plans, best_meta, best_om = plans, meta, om
            else:
                hi = mid

        best_plans, best_meta = propose_portrait_plans(**kw, scale=lo)
        best_om = measure_plans_overlap(
            hero=hero_rgba,
            unit=unit_rgba,
            plans=best_plans,
            canvas_size=canvas,
        )
        ok, mid_vis, back_vis = _unit_vis_ok(
            best_om, mid_min=mid_min, back_min=back_min
        )
        best_meta["overlapPass"] = ok
        best_meta["bindingConstraint"] = "unit_visible"
        best_meta["unit0VisibleMin"] = mid_min
        best_meta["unit1VisibleMin"] = back_min
        best_meta["unit0Visible"] = mid_vis
        best_meta["unit1Visible"] = back_vis
        best_meta["overlapNoteJa"] = (
            f"N=3: unit0≥{mid_min:.3f}/unit1≥{back_min:.3f}で最大同尺"
            f"（scale={float(best_meta.get('scaleHero') or lo):.3f}"
            f" / u0={mid_vis:.3f} / u1={back_vis:.3f}）"
        )
        LOG.info(
            "portrait_fit n=3 binding=unit_visible scale=%.3f u0=%.3f u1=%.3f pass=%s",
            best_meta.get("scaleHero") or lo,
            mid_vis,
            back_vis,
            best_meta.get("overlapPass"),
        )
        return best_plans, best_meta, best_om

    # N=4: unit0≥30% / unit1≥25% / unit2≥20% で同尺最大化
    if int(n) == 4:
        cfg = dict(n4_cfg or {})
        engine = str(cfg.get("layoutEngine") or "upright_quad_fan").strip().lower()
        if engine in ("legacy", "n4_legacy", "height_fan", "fan"):
            plans, meta = propose_portrait_plans(**kw, scale=None)
            om = measure_plans_overlap(
                hero=hero_rgba, unit=unit_rgba, plans=plans, canvas_size=canvas
            )
            meta["overlapPass"] = True
            meta["overlapNoteJa"] = "N=4 legacy: overlap探索スキップ"
            return plans, meta, om

        # 細長は propose 側で n4_legacy に分岐 → 探索スキップ
        plans_probe, meta_probe = propose_portrait_plans(**kw, scale=None)
        if (
            meta_probe.get("layoutEngineResolved") == "n4_legacy"
            or bool(meta_probe.get("isSlender"))
        ):
            om = measure_plans_overlap(
                hero=hero_rgba, unit=unit_rgba, plans=plans_probe, canvas_size=canvas
            )
            meta_probe["overlapPass"] = True
            meta_probe["overlapNoteJa"] = (
                "N=4 細長→n4_legacy（頂円弧・高さ扇）: overlap探索スキップ"
            )
            return plans_probe, meta_probe, om

        u0_min = float(cfg.get("unit0VisibleMin", 0.30))
        u1_min = float(cfg.get("unit1VisibleMin", 0.25))
        u2_min = float(cfg.get("unit2VisibleMin", 0.20))
        search_iters = int(cfg.get("scaleSearchIters", max(16, iters)))
        scale_floor = float(cfg.get("scaleFloor", 0.15))

        plans_hi, meta_hi = plans_probe, meta_probe
        om_hi = measure_plans_overlap(
            hero=hero_rgba, unit=unit_rgba, plans=plans_hi, canvas_size=canvas
        )
        scale_cap = float(
            meta_hi.get("scaleCapHeroInside")
            or meta_hi.get("scaleHero")
            or meta_hi.get("scale")
            or 0.5
        )
        ok_hi, v0h, v1h, v2h = _unit_vis_ok_n4(
            om_hi, u0_min=u0_min, u1_min=u1_min, u2_min=u2_min
        )
        if ok_hi:
            meta_hi["overlapPass"] = True
            meta_hi["bindingConstraint"] = "hero_inside"
            meta_hi["unit0VisibleMin"] = u0_min
            meta_hi["unit1VisibleMin"] = u1_min
            meta_hi["unit2VisibleMin"] = u2_min
            meta_hi["unit0Visible"] = v0h
            meta_hi["unit1Visible"] = v1h
            meta_hi["unit2Visible"] = v2h
            meta_hi["overlapNoteJa"] = (
                f"N=4: 枠内最大同尺で u0={v0h:.3f}≥{u0_min:.3f} / "
                f"u1={v1h:.3f}≥{u1_min:.3f} / u2={v2h:.3f}≥{u2_min:.3f}"
            )
            return plans_hi, meta_hi, om_hi

        lo = max(0.05, scale_floor)
        hi = scale_cap
        best_plans, best_meta, best_om = plans_hi, meta_hi, om_hi
        for _ in range(max(4, search_iters)):
            mid = (lo + hi) * 0.5
            plans, meta = propose_portrait_plans(**kw, scale=mid)
            om = measure_plans_overlap(
                hero=hero_rgba, unit=unit_rgba, plans=plans, canvas_size=canvas
            )
            ok, _, _, _ = _unit_vis_ok_n4(
                om, u0_min=u0_min, u1_min=u1_min, u2_min=u2_min
            )
            if ok:
                lo = mid
                best_plans, best_meta, best_om = plans, meta, om
            else:
                hi = mid

        best_plans, best_meta = propose_portrait_plans(**kw, scale=lo)
        best_om = measure_plans_overlap(
            hero=hero_rgba, unit=unit_rgba, plans=best_plans, canvas_size=canvas
        )
        ok, v0, v1, v2 = _unit_vis_ok_n4(
            best_om, u0_min=u0_min, u1_min=u1_min, u2_min=u2_min
        )
        best_meta["overlapPass"] = ok
        best_meta["bindingConstraint"] = "unit_visible"
        best_meta["unit0VisibleMin"] = u0_min
        best_meta["unit1VisibleMin"] = u1_min
        best_meta["unit2VisibleMin"] = u2_min
        best_meta["unit0Visible"] = v0
        best_meta["unit1Visible"] = v1
        best_meta["unit2Visible"] = v2
        best_meta["overlapNoteJa"] = (
            f"N=4: u0≥{u0_min:.3f}/u1≥{u1_min:.3f}/u2≥{u2_min:.3f}で最大同尺"
            f"（scale={float(best_meta.get('scaleHero') or lo):.3f}"
            f" / u0={v0:.3f} / u1={v1:.3f} / u2={v2:.3f}）"
        )
        LOG.info(
            "portrait_fit n=4 binding=unit_visible scale=%.3f u0=%.3f u1=%.3f u2=%.3f pass=%s",
            best_meta.get("scaleHero") or lo,
            v0,
            v1,
            v2,
            best_meta.get("overlapPass"),
        )
        return best_plans, best_meta, best_om

    # N≥5: 列積み優先
    if int(n) >= 5:
        plans, meta = propose_portrait_plans(**kw, scale=None)
        om = measure_plans_overlap(
            hero=hero_rgba,
            unit=unit_rgba,
            plans=plans,
            canvas_size=canvas,
        )
        meta["overlapPass"] = True
        meta["overlapNoteJa"] = "N≥5傾け積みは列スケール優先のため overlap 二分探索スキップ"
        LOG.info(
            "portrait_fit n=%s pattern=%s scaleH=%.3f tilts=%s",
            n,
            meta.get("pattern"),
            meta.get("scaleHero") or meta.get("scale"),
            meta.get("tiltsDegCw"),
        )
        return plans, meta, om

    _, meta0 = propose_portrait_plans(**kw, scale=None)
    hi = float(meta0["scaleMax"])
    lo = 0.08
    best_plans, best_meta = propose_portrait_plans(**kw, scale=hi)
    best_om: Dict[str, Any] = {"deferred": True}

    for _ in range(max(4, iters)):
        mid = (lo + hi) / 2.0
        plans, meta = propose_portrait_plans(**kw, scale=mid)
        om = measure_plans_overlap(
            hero=hero_rgba,
            unit=unit_rgba,
            plans=plans,
            canvas_size=canvas,
        )
        ok = (
            float(om.get("pairMaxAnyCovered") or 0) <= pair_overlap_max + 1e-6
            and float(om.get("heroMinVisible") or 1) >= hero_visible_min - 1e-6
        )
        if ok:
            lo = mid
            best_plans, best_meta, best_om = plans, meta, om
        else:
            hi = mid

    best_plans, best_meta = propose_portrait_plans(**kw, scale=lo)
    best_om = measure_plans_overlap(
        hero=hero_rgba,
        unit=unit_rgba,
        plans=best_plans,
        canvas_size=canvas,
    )
    best_meta["overlapPass"] = (
        float(best_om.get("pairMaxAnyCovered") or 0) <= pair_overlap_max + 1e-6
        and float(best_om.get("heroMinVisible") or 1) >= hero_visible_min - 1e-6
    )
    LOG.info(
        "portrait_fit n=%s pattern=%s scale=%.3f tilts=%s ovAny=%.3f pass=%s",
        n,
        best_meta.get("pattern"),
        best_meta.get("scale"),
        best_meta.get("tiltsDegCw"),
        best_om.get("pairMaxAnyCovered"),
        best_meta.get("overlapPass"),
    )
    return best_plans, best_meta, best_om
