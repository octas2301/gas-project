# -*- coding: utf-8 -*-
"""
余白最小化レイアウト（縁際配置）

方針:
- 配置パターンを固定し、縁にピン留め
- 重なり硬制約（pairOverlapMax / heroVisibleMin）の下でスケール最大化
- N≤4: 全個体同一スケール（単体画像テスト時は hero=unit）
- N≥5: **先に unit（heroの逆側＝右）縦列を最大化** → 残り左余白で hero を最大化

N の提案配置（本モジュールの正）:
  N=1  単体最大化・中央配置（hero のみ）
  N=2  対角: 左下=hero（手前）/ 右上=unit
  N=3  三角形: 左下=hero（手前）/ 右下=unit / 上中央=unit
  N=4  四隅 2×2: 左下=hero（手前）/ 右下・左上・右上=unit
  N≥5  unitを右に配置（<10:1列 / 10–19:2列÷2 / ≥20:10+10+端数3列目〜）→ 左余白で hero 最大
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

LOG = logging.getLogger("set_main_image.edge_layout")

# 縁までの目標余白（キャンバス比）
EDGE_MARGIN_RATIO = 0.02
# N≥5 右列 unit: 横幅がキャンバス比この値未満なら縦重なりを強めて拡大
N5_UNIT_MIN_WIDTH_RATIO = 0.20
# 右列 unit 同士の縦重なり上限（幅確保用）
N5_UNIT_OV_MAX = 0.62
# 1列あたりの縦最大（達したら列を増やす）
N5_UNIT_MAX_PER_COL = 10
# 列同士の水平重なり（右から左へ積む）の下限〜上限
N5_UNIT_COL_H_OVERLAP_MIN = 0.22
N5_UNIT_COL_H_OVERLAP_MAX = 0.48


@dataclass
class EdgePlan:
    role: str
    x: int
    y: int
    scale: float
    z: int


def _max_scale_for_wh(
    pw: int, ph: int, need_w: float, need_h: float, canvas: int, margin: int
) -> float:
    """need_w/h はスケール1のときの必要幅高さ（重なり込み）。"""
    usable = canvas - 2 * margin
    if need_w <= 0 or need_h <= 0:
        return 0.1
    return min(usable / need_w, usable / need_h)


def unit_column_counts(num_u: int) -> List[int]:
    """
    右ブロックの列ごとの個数（右端列が index0）。

    - <10: 1列
    - 10〜19: ÷2 して縦2列
    - ≥20: 10+10 を固定し、端数は3列目以降（各列最大10）
    """
    num_u = max(0, int(num_u))
    if num_u <= 0:
        return []
    if num_u < N5_UNIT_MAX_PER_COL:
        return [num_u]
    if num_u < 2 * N5_UNIT_MAX_PER_COL:
        a = (num_u + 1) // 2
        b = num_u // 2
        return [a, b]
    counts = [N5_UNIT_MAX_PER_COL, N5_UNIT_MAX_PER_COL]
    rem = num_u - 2 * N5_UNIT_MAX_PER_COL
    while rem > 0:
        take = min(N5_UNIT_MAX_PER_COL, rem)
        counts.append(take)
        rem -= take
    return counts


def max_unit_scale_right_col(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    margin: int,
    overlap: float,
) -> Tuple[float, float, List[int], float]:
    """
    戻り値: (scale, ov_v, counts, h_ov)
    高さは最多列基準。幅は列数＋水平重なり。
    単体横幅 <20% なら縦ov→水平ovの順で押し上げ。
    """
    num_u = max(1, int(n) - 1)
    counts = unit_column_counts(num_u)
    max_rows = max(counts) if counts else 1
    n_cols = max(1, len(counts))
    usable = canvas - 2 * margin
    ov_v = max(0.18, min(N5_UNIT_OV_MAX, float(overlap)))
    h_ov = N5_UNIT_COL_H_OVERLAP_MIN

    def _scale_at(ov: float, hov: float) -> float:
        need_h = product_h * (1.0 + (max_rows - 1) * (1.0 - ov))
        need_w = product_w * (1.0 + (n_cols - 1) * (1.0 - hov))
        s_h = usable / need_h if need_h > 0 else 0.1
        s_w = usable / need_w if need_w > 0 else 0.1
        return max(0.05, min(s_h, s_w))

    scale = _scale_at(ov_v, h_ov)
    min_w = canvas * N5_UNIT_MIN_WIDTH_RATIO
    if product_w * scale < min_w and max_rows >= 2:
        target_scale = min_w / float(product_w)
        need_factor = usable / (product_h * target_scale)
        if need_factor >= 1.0:
            ov_needed = 1.0 - (need_factor - 1.0) / max(1, max_rows - 1)
            ov_v = max(ov_v, min(N5_UNIT_OV_MAX, ov_needed))
        else:
            ov_v = N5_UNIT_OV_MAX
        scale = _scale_at(ov_v, h_ov)
    # まだ幅不足（多列で横が詰まる）→ 列の水平重なりを強化
    if product_w * scale < min_w and n_cols >= 2:
        target_scale = min_w / float(product_w)
        # usable >= pw*s * (1+(ncols-1)*(1-hov)) → hov を逆算
        need_factor_w = usable / (product_w * target_scale)
        if need_factor_w >= 1.0 and n_cols > 1:
            hov_needed = 1.0 - (need_factor_w - 1.0) / (n_cols - 1)
            h_ov = max(h_ov, min(N5_UNIT_COL_H_OVERLAP_MAX, hov_needed))
        else:
            h_ov = N5_UNIT_COL_H_OVERLAP_MAX
        scale = _scale_at(ov_v, h_ov)
        if product_w * scale < min_w:
            ov_v = N5_UNIT_OV_MAX
            h_ov = N5_UNIT_COL_H_OVERLAP_MAX
            scale = _scale_at(ov_v, h_ov)
    # それでも単体幅20%未満なら、縦ovを上げて幅優先（列は左へ張り出し可）
    if product_w * scale < min_w:
        target_scale = min_w / float(product_w)
        scale = min(target_scale, usable / float(product_w))
        h_ov = N5_UNIT_COL_H_OVERLAP_MAX
        if max_rows >= 2:
            need_factor = usable / (product_h * scale)
            if need_factor >= 1.0:
                ov_v = max(ov_v, min(0.72, 1.0 - (need_factor - 1.0) / (max_rows - 1)))
            else:
                ov_v = 0.72
                need_h = product_h * (1.0 + (max_rows - 1) * (1.0 - ov_v))
                scale = min(scale, usable / need_h)
        scale = min(scale, _scale_at(ov_v, h_ov) if n_cols == 1 else scale)
    return scale, ov_v, counts, h_ov


def max_edge_scale(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    overlap: float = 0.28,
) -> Tuple[float, int]:
    """
    パターン上の探索用最大スケールと margin px。
    N≥5 は **unit スケール**（サイズ決定の主変数）を返す。
    """
    n = int(n)
    margin = max(8, int(canvas * EDGE_MARGIN_RATIO))
    pw, ph = product_w, product_h
    ov = max(0.12, min(0.45, float(overlap)))
    usable = canvas - 2 * margin
    if n == 1:
        return min(usable / float(pw), usable / float(ph)), margin
    if n == 2:
        return min(usable / float(pw), usable / float(ph)), margin
    if n == 3:
        return min(usable / float(pw), usable / float(ph)) * 0.92, margin
    if n == 4:
        need_w = pw * (2.0 - ov)
        need_h = ph * (2.0 - ov * 0.85)
        return _max_scale_for_wh(pw, ph, need_w, need_h, canvas, margin), margin
    s_u, _ov, _counts, _hov = max_unit_scale_right_col(
        n=n,
        canvas=canvas,
        product_w=pw,
        product_h=ph,
        margin=margin,
        overlap=ov,
    )
    return s_u, margin


def plan_edge_layout(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    overlap: float = 0.28,
    scale: Optional[float] = None,
) -> Tuple[List[EdgePlan], Dict[str, Any]]:
    """
    scale:
      - N=1: 単体スケール（中央最大化）
      - N≤4: 全個体共通スケール
      - N≥5: **unit スケール**（hero は左余白から自動最大）
    """
    n = int(n)
    if n < 1:
        raise ValueError("n>=1")
    margin = max(8, int(canvas * EDGE_MARGIN_RATIO))
    pw, ph = product_w, product_h
    ov = max(0.12, min(0.45, float(overlap)))
    scale_max, _ = max_edge_scale(
        n=n, canvas=canvas, product_w=pw, product_h=ph, overlap=ov
    )
    if scale is None:
        scale = scale_max
    else:
        scale = max(0.05, min(float(scale), scale_max))

    meta: Dict[str, Any] = {
        "layoutFamily": "edge_fill",
        "edgeMarginRatio": EDGE_MARGIN_RATIO,
        "overlapParam": ov,
        "n": n,
        "scaleMax": scale_max,
    }

    w = max(1, int(round(pw * scale)))
    h = max(1, int(round(ph * scale)))

    if n == 1:
        plans = [
            EdgePlan("hero", (canvas - w) // 2, (canvas - h) // 2, scale, z=1),
        ]
        meta["pattern"] = "n1_center_max"
        meta["zOrderJa"] = "単体heroのみ"
        meta["proposalJa"] = "N=1: 余白内で最大化し中央配置。"

    elif n == 2:
        plans = [
            EdgePlan("unit", canvas - margin - w, margin, scale, z=0),
            EdgePlan("hero", margin, canvas - margin - h, scale, z=1),
        ]
        meta["pattern"] = "n2_diag_bl_tr"
        meta["zOrderJa"] = "左下=手前（重なり上側）"
        meta["proposalJa"] = "N=2: 左下hero・右上unit。左下が手前。"

    elif n == 3:
        # 左下=hero（最前面）。unit=上中央・右下
        plans = [
            EdgePlan("unit", (canvas - w) // 2, margin, scale, z=0),
            EdgePlan("unit", canvas - margin - w, canvas - margin - h, scale, z=1),
            EdgePlan("hero", margin, canvas - margin - h, scale, z=100),
        ]
        meta["pattern"] = "n3_triangle_hero_bl"
        meta["zOrderJa"] = "hero左下=最前面（unitの上）"
        meta["proposalJa"] = "N=3: 左下hero（手前）＋上中央unit＋右下unit。"

    elif n == 4:
        # 左下=hero（最前面）。他三隅=unit
        plans = [
            EdgePlan("unit", canvas - margin - w, margin, scale, z=0),
            EdgePlan("unit", margin, margin, scale, z=1),
            EdgePlan("unit", canvas - margin - w, canvas - margin - h, scale, z=2),
            EdgePlan("hero", margin, canvas - margin - h, scale, z=100),
        ]
        meta["pattern"] = "n4_corners_hero_bl"
        meta["zOrderJa"] = "hero左下=最前面（unitの上）"
        meta["proposalJa"] = "N=4: 四隅。左下=hero（手前）、他三隅=unit。"

    else:
        # --- N≥5: ① 右ブロック unit 先決（列分割＋幅20%）② 左余白で hero 最大 ---
        _s_cap, ov_col, counts, h_ov = max_unit_scale_right_col(
            n=n,
            canvas=canvas,
            product_w=pw,
            product_h=ph,
            margin=margin,
            overlap=ov,
        )
        scale_u = scale
        num_u = n - 1
        if not counts:
            counts = unit_column_counts(num_u)
        uw = max(1, int(round(pw * scale_u)))
        uh = max(1, int(round(ph * scale_u)))
        usable = canvas - 2 * margin
        width_ratio = uw / float(canvas)
        step_x = max(1, int(round(uw * (1.0 - h_ov))))

        plans = []
        z_i = 0
        # 右端列から左へ。縦間隔は最多列基準で統一。端数・短い列は下から積む
        max_rows = max(counts) if counts else 1
        y_top = margin
        y_bot = canvas - margin - uh
        if max_rows <= 1:
            step_y = 0.0
        else:
            step_y = max(0.0, (y_bot - y_top) / (max_rows - 1))
        leftmost_x = canvas
        for col_i, nrows in enumerate(counts):
            x_u = canvas - margin - uw - col_i * step_x
            x_u = max(margin, x_u)
            leftmost_x = min(leftmost_x, x_u)
            # 下から nrows 個（満杯列は上端まで届く）。間隔は最多列と同一
            for r in range(nrows):
                y = int(round(y_bot - r * step_y))
                y = max(margin, min(y, canvas - margin - uh))
                # r=0 が最下＝手前寄り
                z = (len(counts) - 1 - col_i) * 100 + (nrows - 1 - r)
                plans.append(EdgePlan("unit", x_u, y, scale_u, z=z))

        meta["unitStepYPx"] = round(step_y, 2)
        meta["unitStackFrom"] = "bottom"

        # ヒーローは左余白内に収める（unit列への食い込み禁止）。
        # 以前の bite 食い込みで、ヒーロー透過域の下に unit が見え
        # 「ヒーロー右端が切れた／投下された」ように見えていた。
        gap = max(12, int(round(uw * 0.06)))
        left_w = max(1, leftmost_x - margin - gap)
        scale_h = min(left_w / float(pw), usable / float(ph))
        scale_h = max(0.05, float(scale_h))
        hw = max(1, int(round(pw * scale_h)))
        hh = max(1, int(round(ph * scale_h)))
        if hw > left_w:
            scale_h = left_w / float(pw)
            hw = max(1, int(round(pw * scale_h)))
            hh = max(1, int(round(ph * scale_h)))
        # 右端が unit に食い込まないよう再クランプ
        if margin + hw > leftmost_x - gap:
            scale_h = max(0.05, (leftmost_x - gap - margin) / float(pw))
            hw = max(1, int(round(pw * scale_h)))
            hh = max(1, int(round(ph * scale_h)))
        x_h = margin
        y_h = canvas - margin - hh
        plans.append(EdgePlan("hero", x_h, y_h, scale_h, z=1000))

        meta["pattern"] = "n5plus_units_right_cols_hero_left"
        meta["zOrderJa"] = "右ブロックunit（下・右寄り手前）、hero左下=最手前"
        meta["proposalJa"] = (
            "N≥5: 右unit列分割→幅20%→左余白でhero最大（unitへ食い込ませない）。"
        )
        meta["sizeOrderJa"] = "unit先決→heroは左余白最大（gap確保）"
        meta["scaleUnit"] = scale_u
        meta["scaleHero"] = scale_h
        meta["leftGapPx"] = leftmost_x - margin
        meta["heroUnitGapPx"] = gap
        meta["heroBitePx"] = 0
        meta["unitColOverlapUsed"] = ov_col
        meta["unitColHOverlapUsed"] = h_ov
        meta["unitColumnCounts"] = counts
        meta["unitNumColumns"] = len(counts)
        meta["unitWidthRatio"] = round(width_ratio, 4)
        meta["unitMinWidthRatio"] = N5_UNIT_MIN_WIDTH_RATIO
        meta["unitWidthBoost"] = bool(width_ratio >= N5_UNIT_MIN_WIDTH_RATIO - 1e-6)

    for p in plans:
        p.x = max(0, min(p.x, canvas - 1))
        p.y = max(0, min(p.y, canvas - 1))

    # N≥5 の meta["scale"] は探索主変数＝unit
    meta["scale"] = scale
    meta["marginPx"] = margin
    return plans, meta


def fit_edge_layout_under_overlap(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_rgba: Any,
    unit_rgba: Any,
    pair_overlap_max: float = 0.35,
    hero_visible_min: float = 0.70,
    overlap: float = 0.28,
    iters: int = 18,
) -> Tuple[List[EdgePlan], Dict[str, Any], Dict[str, Any]]:
    """
    重なり硬制約を満たす最大スケールを二分探索。
    N≥5 は unit スケールを探索（hero は都度・左余白最大）。
    """
    from overlap_metrics import measure_plans_overlap

    scale_hi, _ = max_edge_scale(
        n=n,
        canvas=canvas,
        product_w=product_w,
        product_h=product_h,
        overlap=overlap,
    )
    min_scale_w = (N5_UNIT_MIN_WIDTH_RATIO * canvas) / float(product_w)
    if n >= 5:
        # 横幅20%未満へは落とさない（unit列の優先）
        scale_lo = max(0.05, min(scale_hi, max(scale_hi * 0.08, min_scale_w)))
    else:
        scale_lo = max(0.05, scale_hi * 0.08)

    def _ok(om: Dict[str, Any], meta: Optional[Dict[str, Any]] = None) -> bool:
        any_cap = pair_overlap_max
        if n >= 5 and meta is not None:
            # 右ブロックは多段・多列でアルファ上の被りが大きくなりやすい。意図的密度として緩めに許容。
            # hero が単体を隠す率（pairMaxBackCovered）は従来 0.35 のまま。
            any_cap = max(
                0.85,
                pair_overlap_max,
                float(meta.get("unitColOverlapUsed") or 0.0),
                float(meta.get("unitColHOverlapUsed") or 0.0),
                N5_UNIT_OV_MAX,
            )
        return (
            float(om.get("pairMaxAnyCovered", 1.0)) <= any_cap
            and float(om.get("pairMaxBackCovered", 1.0)) <= pair_overlap_max
            and float(om.get("heroMinVisible", 0.0)) >= hero_visible_min
        )

    def _eval(sc: float):
        plans, meta = plan_edge_layout(
            n=n,
            canvas=canvas,
            product_w=product_w,
            product_h=product_h,
            overlap=overlap,
            scale=sc,
        )
        om = measure_plans_overlap(
            hero=hero_rgba,
            unit=unit_rgba,
            plans=plans,
            canvas_size=canvas,
        )
        return plans, meta, om

    hi_plans, hi_meta, hi_om = _eval(scale_hi)
    if _ok(hi_om, hi_meta):
        best_plans, best_meta, best_om = hi_plans, hi_meta, hi_om
        best_scale = scale_hi
    else:
        lo, hi = scale_lo, scale_hi
        best_scale = scale_lo
        best_plans, best_meta, best_om = _eval(scale_lo)
        for _ in range(iters):
            mid = (lo + hi) / 2.0
            plans, meta, om = _eval(mid)
            if _ok(om, meta):
                best_plans, best_meta, best_om = plans, meta, om
                best_scale = mid
                lo = mid
            else:
                hi = mid
        best_plans, best_meta, best_om = _eval(best_scale)

    passed = _ok(best_om, best_meta)
    note = (
        "N≥5: unit右縦列を先に最大化→左余白でhero最大化。重なり硬制約下でunitスケール探索"
        if n >= 5
        else "縁際パターン固定→重なり硬制約下でスケール最大化（二分探索）"
    )
    overlap_meta = {
        "deferred": False,
        "pass": passed,
        "pairOverlapMax": pair_overlap_max,
        "heroVisibleMin": hero_visible_min,
        "scaleChosen": round(float(best_scale), 4),
        "scaleMax": round(scale_hi, 4),
        "scaleRole": "unit" if n >= 5 else "uniform",
        "measured": best_om,
        "noteJa": note,
    }
    best_meta["scaleChosen"] = best_scale
    best_meta["overlapPass"] = passed
    LOG.info(
        "edge_fit n=%s pattern=%s scaleU=%.3f/%.3f scaleH=%s ovAny=%.3f pass=%s",
        n,
        best_meta.get("pattern"),
        best_scale,
        scale_hi,
        best_meta.get("scaleHero"),
        best_om.get("pairMaxAnyCovered"),
        passed,
    )
    return best_plans, best_meta, overlap_meta
