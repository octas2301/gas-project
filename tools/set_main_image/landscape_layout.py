# -*- coding: utf-8 -*-
"""
横長パターン（landscape）— 正方形キャンバス＋横長商品向けセット組み

縦長との共通:
- キャンバスは正方形
- 縦横比ロック（等倍＋平行移動。ストレッチ禁止）
- 素材は 01.amazon白抜きベース（透過PNG）
- Octas は compose 側（hero右下キス）

セット組み（合格固定・RULES §1.2）:
- N=1: 枠内最大化・高さ中央揃え＋左右中央
- N=2: 縦二段。下=hero左辺接触／上=unit右辺接触
- N=3/4: 左下→右上階段＋枠ピン。最小横進み≥幅22%を優先
- N≥5: 上hero（左+上接触・右はOctas帯）＋下グリッド。半端行は中央寄せ
"""
from __future__ import annotations

import logging
import math
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

LOG = logging.getLogger("set_main_image.landscape_layout")

DEFAULT_MARGIN = 36
DEFAULT_GAP = 14
DEFAULT_STAIR_STEP_DX_FRAC = 0.22
DEFAULT_N5_OCTAS_RIGHT_BAND = 140


@dataclass
class LandscapePlan:
    role: str
    x: int
    y: int
    scale: float
    z: int
    rotation_deg: float = 0.0
    anchor: str = "topleft"
    foot_x: Optional[int] = None
    foot_y: Optional[int] = None
    top_x: Optional[int] = None
    top_y: Optional[int] = None


def _choose_cols(rem: int) -> int:
    if rem <= 0:
        return 1
    if rem <= 3:
        return rem
    target = math.sqrt(rem)
    c_min = max(2, int(math.floor(target - 1)))
    c_max = min(rem - 1, int(math.ceil(target + 2)))
    if c_max < c_min:
        c_max = c_min
    best = max(2, int(round(target)))
    best_score = None
    for c in range(c_min, c_max + 1):
        rows = int(math.ceil(rem / float(c)))
        top = rem % c
        partial = 0 if top == 0 else 1
        aspect = abs(rows - c)
        dist = abs(c - target)
        flat_pen = 10 if rows == 1 else 0
        score = (flat_pen, aspect, -c, partial, dist)
        if best_score is None or score < best_score:
            best_score = score
            best = c
    return best


def _grid_rows(rem: int, cols: int) -> List[int]:
    if rem <= 0:
        return []
    full = rem // cols
    top = rem % cols
    rows: List[int] = []
    if top:
        rows.append(top)
    for _ in range(full):
        rows.append(cols)
    return rows


def propose_landscape_plans(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    margin: int = DEFAULT_MARGIN,
    gap: int = DEFAULT_GAP,
    stair_step_dx_frac: float = DEFAULT_STAIR_STEP_DX_FRAC,
    n5_octas_right_band: int = DEFAULT_N5_OCTAS_RIGHT_BAND,
) -> Tuple[List[LandscapePlan], Dict[str, Any]]:
    """N別の貼付計画（scale は product_w/h 基準）。"""
    n = int(n)
    if n < 1:
        raise ValueError("n>=1")
    pw = max(1, int(product_w))
    ph = max(1, int(product_h))
    inner = canvas - 2 * margin

    if n == 1:
        s = min(inner / float(pw), inner / float(ph))
        nw = max(1, int(round(pw * s)))
        nh = max(1, int(round(ph * s)))
        x = margin + (inner - nw) // 2
        y = margin + (inner - nh) // 2
        plans = [LandscapePlan("hero", x, y, s, z=1)]
        meta = {
            "layoutFamily": "landscape_set",
            "pattern": "n1_center_vcenter",
            "n": n,
            "scale": s,
            "proposalJa": "N=1: 枠内最大・高さ中央揃え＋左右中央。",
            "status": "locked_pass",
        }
        return plans, meta

    if n == 2:
        avail_h = canvas - 2 * margin - gap
        uh = avail_h // 2
        s = uh / float(ph)
        nw = max(1, int(round(pw * s)))
        nh = max(1, int(round(ph * s)))
        # 幅が枠を超える場合は縮小
        if nw > inner:
            s = inner / float(pw)
            nw = max(1, int(round(pw * s)))
            nh = max(1, int(round(ph * s)))
        hero_x = margin
        hero_y = canvas - margin - nh
        unit_x = canvas - margin - nw
        unit_y = margin
        if hero_y < unit_y + nh + gap:
            uh2 = (canvas - 2 * margin - gap) // 2
            s = min(uh2 / float(ph), inner / float(pw))
            nw = max(1, int(round(pw * s)))
            nh = max(1, int(round(ph * s)))
            hero_x = margin
            hero_y = canvas - margin - nh
            unit_x = canvas - margin - nw
            unit_y = margin
        plans = [
            LandscapePlan("unit", unit_x, unit_y, s, z=0),
            LandscapePlan("hero", hero_x, hero_y, s, z=1),
        ]
        meta = {
            "layoutFamily": "landscape_set",
            "pattern": "n2_vstack_lr_anchor",
            "n": n,
            "scale": s,
            "proposalJa": "N=2: 下hero左辺／上unit右辺。",
            "status": "locked_pass",
        }
        return plans, meta

    if n in (3, 4):
        steps = n - 1
        max_uw = int(inner / (1.0 + steps * stair_step_dx_frac))
        lo, hi = 0.05, min(1.2, max_uw / float(pw) + 0.02)
        best = None
        for _ in range(32):
            mid = (lo + hi) / 2.0
            nw = max(1, int(pw * mid))
            nh = max(1, int(round(ph * (nw / float(pw)))))
            s = nw / float(pw)
            hero_x = margin
            hero_y = canvas - margin - nh
            last_x = canvas - margin - nw
            last_y = margin
            travel_x = last_x - hero_x
            min_travel = int(round(steps * stair_step_dx_frac * nw))
            ok = (
                nw <= inner
                and nh <= inner
                and last_x >= margin
                and hero_y >= margin
                and travel_x >= min_travel
                and nw <= max_uw + 2
            )
            if ok:
                for i in range(1, n - 1):
                    t = i / float(n - 1)
                    x = int(round(hero_x + (last_x - hero_x) * t))
                    y = int(round(hero_y + (last_y - hero_y) * t))
                    if x < 0 or y < 0 or x + nw > canvas or y + nh > canvas:
                        ok = False
                        break
            if ok:
                best = (s, nw, nh, hero_x, hero_y, last_x, last_y, travel_x)
                lo = mid
            else:
                hi = mid
        if not best:
            raise RuntimeError("landscape stair fit failed")
        s, nw, nh, hero_x, hero_y, last_x, last_y, travel_x = best
        plans: List[LandscapePlan] = []
        for i in range(n - 1, -1, -1):
            t = i / float(n - 1)
            x = int(round(hero_x + (last_x - hero_x) * t))
            y = int(round(hero_y + (last_y - hero_y) * t))
            role = "hero" if i == 0 else "unit"
            z = 1000 if i == 0 else (100 - i)
            plans.append(LandscapePlan(role, x, y, s, z=z))
        meta = {
            "layoutFamily": "landscape_set",
            "pattern": "n%d_stair_frame_pin" % n,
            "n": n,
            "scale": s,
            "travelX": travel_x,
            "stepDxFrac": stair_step_dx_frac,
            "proposalJa": "N=%d: 階段＋枠ピン（横進み≥%.0f%%幅）。" % (n, stair_step_dx_frac * 100),
            "status": "locked_pass",
        }
        return plans, meta

    # N≥5
    rem = n - 1
    cols = _choose_cols(rem)
    rows = _grid_rows(rem, cols)
    n_rows = len(rows)
    band = min(int(n5_octas_right_band), max(80, inner // 8))
    hero_w = max(64, inner - band)
    s_h = hero_w / float(pw)
    hw = max(1, int(round(pw * s_h)))
    hh = max(1, int(round(ph * s_h)))
    hero_x = margin
    hero_y = margin

    grid_top = hero_y + hh + gap
    grid_bottom = canvas - margin
    grid_h_budget = max(20, grid_bottom - grid_top)
    cell_h = max(12, (grid_h_budget - gap * (n_rows - 1)) // max(1, n_rows))
    s_u = cell_h / float(ph)
    cell_w = max(1, int(round(pw * s_u)))
    cell_h = max(1, int(round(ph * s_u)))
    grid_w = cols * cell_w + gap * (cols - 1)
    if grid_w > inner:
        cell_w = max(1, (inner - gap * (cols - 1)) // cols)
        s_u = cell_w / float(pw)
        cell_w = max(1, int(round(pw * s_u)))
        cell_h = max(1, int(round(ph * s_u)))
        grid_w = cols * cell_w + gap * (cols - 1)
        used = n_rows * cell_h + gap * (n_rows - 1)
        if used > grid_h_budget and used > 0:
            s_u *= grid_h_budget / float(used)
            cell_w = max(1, int(round(pw * s_u)))
            cell_h = max(1, int(round(ph * s_u)))
            grid_w = cols * cell_w + gap * (cols - 1)

    full_x0 = margin + (inner - grid_w) // 2
    plans = [LandscapePlan("hero", hero_x, hero_y, s_h, z=1000)]
    used_h = n_rows * cell_h + gap * (n_rows - 1)
    gy = grid_top
    if used_h < grid_h_budget:
        gy = grid_top + (grid_h_budget - used_h) // 2
    idx = 0
    for count in rows:
        if count == cols:
            xs = [full_x0 + i * (cell_w + gap) for i in range(count)]
        else:
            row_w = count * cell_w + gap * max(0, count - 1)
            x0 = margin + (inner - row_w) // 2
            xs = [x0 + i * (cell_w + gap) for i in range(count)]
        for x in xs:
            idx += 1
            plans.append(
                LandscapePlan("unit", int(x), int(gy), s_u, z=100 - idx)
            )
        gy += cell_h + gap

    meta = {
        "layoutFamily": "landscape_set",
        "pattern": "n5plus_hero_top_grid",
        "n": n,
        "scaleHero": s_h,
        "scaleUnit": s_u,
        "cols": cols,
        "rows": rows,
        "octasRightBand": band,
        "proposalJa": "N≥5: hero左+上（右Octas帯）＋下グリッド・半端行中央寄せ。",
        "status": "locked_pass",
    }
    return plans, meta


def fit_landscape_layout_under_overlap(
    *,
    n: int,
    canvas: int,
    product_w: int,
    product_h: int,
    hero_rgba=None,
    unit_rgba=None,
    pair_overlap_max: float = 0.45,
    hero_visible_min: float = 0.65,
    landscape_cfg: Optional[Dict[str, Any]] = None,
    **_kwargs,
) -> Tuple[List[LandscapePlan], Dict[str, Any], Dict[str, Any]]:
    """本線入口。横長は幾何優先のため overlap 二分探索はスキップ。"""
    del hero_rgba, unit_rgba, pair_overlap_max, hero_visible_min
    cfg = landscape_cfg or {}
    margin = int(cfg.get("marginPx") or DEFAULT_MARGIN)
    gap = int(cfg.get("gapPx") or DEFAULT_GAP)
    step_frac = float(cfg.get("stairStepDxFrac") or DEFAULT_STAIR_STEP_DX_FRAC)
    band = int(cfg.get("n5OctasRightBand") or DEFAULT_N5_OCTAS_RIGHT_BAND)

    plans, meta = propose_landscape_plans(
        n=n,
        canvas=canvas,
        product_w=product_w,
        product_h=product_h,
        margin=margin,
        gap=gap,
        stair_step_dx_frac=step_frac,
        n5_octas_right_band=band,
    )
    meta["canvasStillSquare"] = True
    meta["aspectId"] = "landscape"
    overlap_meta = {
        "deferred": True,
        "pass": True,
        "overlapNoteJa": "landscape: 幾何ルール優先（overlap探索スキップ）",
    }
    LOG.info(
        "landscape n=%s pattern=%s plans=%s",
        n,
        meta.get("pattern"),
        len(plans),
    )
    return plans, meta, overlap_meta
