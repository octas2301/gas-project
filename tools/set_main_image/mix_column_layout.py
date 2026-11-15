# -*- coding: utf-8 -*-
"""
ミックス列積みレイアウト（aspect 共通）

種＝列、数量＝列内縦（下前・上奥）。全列右傾き。
探索: 見かけ袋サイズ最大 → 辺余白 → 列間。
後段: 袋サイズ固定のまま縦送りだけ広げ上下空きを埋める。
左端逆回転は不採用。
"""
from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PIL import Image

from portrait_layout import rotate_rgba_cw

LOG = logging.getLogger("set_main_image.mix_column")

Slot = Tuple[Image.Image, int, int]
SlotK = Tuple[Image.Image, int, int, int, int]


def _cfg_defaults() -> Dict[str, Any]:
    return {
        "enabled": True,
        "tiltDegCw": 6.0,
        "hOverlapMax": 0.25,
        "pitchMax": 0.92,
        "vStepMin": 0.10,
        "vStepMax": 1.12,
        "vStepPrefLo": 0.28,
        "vStepPrefHi": 1.05,
        "sideFlush": True,
        "pitchStep": 0.03,
        "vStepStep": 0.08,
    }


def load_mix_column_cfg() -> Dict[str, Any]:
    cfg = _cfg_defaults()
    try:
        from rakuten_layer import load_layout_rules

        tun = (load_layout_rules().get("amazon") or {}).get("tuningDraft") or {}
        raw = tun.get("mixColumnLayout") or {}
        if isinstance(raw, dict):
            cfg.update(raw)
    except Exception as e:
        LOG.warning("mixColumnLayout cfg fallback: %s", e)
    return cfg


def _trim(im: Image.Image) -> Image.Image:
    bb = im.split()[-1].getbbox()
    return im.crop(bb) if bb else im


def _fit_h(im: Image.Image, h: int) -> Image.Image:
    w0, h0 = im.size
    s = h / max(1, h0)
    return im.resize((max(1, int(round(w0 * s))), h), Image.Resampling.LANCZOS)


_PREP_CACHE: Dict[Tuple[str, int, float], Image.Image] = {}


def _unit_cache_token(unit: Image.Image) -> str:
    """id() はGC後に再利用され別種を誤ヒットするため、内容指紋を使う。"""
    import hashlib

    raw = unit.tobytes()
    return hashlib.md5(raw).hexdigest()


def clear_prep_cache() -> None:
    _PREP_CACHE.clear()


def _prep(unit: Image.Image, uh: int, tilt: float) -> Image.Image:
    key = (_unit_cache_token(unit), int(uh), float(tilt))
    hit = _PREP_CACHE.get(key)
    if hit is not None:
        return hit
    out = _trim(rotate_rgba_cw(_fit_h(_trim(unit), uh), tilt))
    _PREP_CACHE[key] = out
    return out


def _slots_wh(slots: Sequence[SlotK]) -> Tuple[int, int]:
    xs = [x for _i, x, _y, _a, _b in slots]
    ys = [y for _i, x, y, _a, _b in slots]
    x2 = [x + im.width for im, x, _y, _a, _b in slots]
    y2 = [y + im.height for im, x, y, _a, _b in slots]
    return max(x2) - min(xs), max(y2) - min(ys)


def _edge_pads(gw: int, gh: int, usable: int, canvas: int) -> Tuple[float, float, float]:
    s = min(usable / max(1, gw), usable / max(1, gh))
    nw, nh = gw * s, gh * s
    h_pad = (usable - nw) / 2 / canvas
    v_pad = (usable - nh) / 2 / canvas
    return max(h_pad, v_pad), h_pad, v_pad


def _v_pref_penalty(v_step: float, lo: float, hi: float) -> float:
    if lo <= v_step <= hi:
        return 0.0
    if v_step < lo:
        return lo - v_step
    return v_step - hi


def align_side_extremes(
    slots_kij: List[SlotK],
    *,
    k: int,
) -> List[SlotK]:
    if not slots_kij:
        return []
    items = [
        {"im": im, "x": float(x), "y": float(y), "ki": ki, "j": j}
        for im, x, y, ki, j in slots_kij
    ]

    def box(it):
        x, y = int(round(it["x"])), int(round(it["y"]))
        return (x, y, x + it["im"].width, y + it["im"].height)

    col_tb: Dict[int, Tuple[float, float]] = {}
    for ki in range(k):
        col = [it for it in items if it["ki"] == ki]
        if not col:
            continue
        col_tb[ki] = (
            float(min(box(it)[1] for it in col)),
            float(max(box(it)[3] for it in col)),
        )
    t_star = min(t for t, _b in col_tb.values())
    b_star = max(b for _t, b in col_tb.values())
    span_star = max(1.0, b_star - t_star)

    for ki, (t_i, b_i) in col_tb.items():
        span_i = max(1.0, b_i - t_i)
        scale = span_star / span_i
        for it in items:
            if it["ki"] != ki:
                continue
            _l, t, _r, _b = box(it)
            it["y"] = t_star + (t - t_i) * scale

    for ki in range(k):
        col = [it for it in items if it["ki"] == ki]
        if not col:
            continue
        target_l = min(box(it)[0] for it in col)
        for it in col:
            it["x"] += target_l - box(it)[0]

    return [
        (it["im"], int(round(it["x"])), int(round(it["y"])), it["ki"], it["j"])
        for it in items
    ]


def _ink_col_gap_px(
    slots: List[SlotK],
    *,
    k: int,
    alpha_thresh: int = 40,
) -> int:
    if k <= 1:
        return 0
    gw, gh = _slots_wh(slots)
    if gw <= 0 or gh <= 0:
        return 0
    masks: List[Image.Image] = []
    for ki in range(k):
        m = Image.new("L", (gw, gh), 0)
        for im, x, y, cki, _j in slots:
            if cki != ki:
                continue
            a = im.split()[-1]
            m.paste(a, (x, y), a)
        masks.append(m)

    y_step = 2
    max_gap = 0
    for ki in range(k - 1):
        left = masks[ki]
        right = masks[ki + 1]
        ld = left.load()
        rd = right.load()
        best = None
        for y in range(0, gh, y_step):
            lx = -1
            for x in range(gw - 1, -1, -1):
                if ld[x, y] > alpha_thresh:
                    lx = x
                    break
            if lx < 0:
                continue
            rx = -1
            for x in range(gw):
                if rd[x, y] > alpha_thresh:
                    rx = x
                    break
            if rx < 0:
                continue
            gap = rx - lx - 1
            if best is None or gap < best:
                best = gap
        if best is None:
            max_gap = max(max_gap, gw)
        else:
            max_gap = max(max_gap, max(0, best))
    return int(max_gap)


def _internal_col_gap_ratio(
    slots: List[SlotK],
    *,
    k: int,
    usable: int,
) -> float:
    if k <= 1:
        return 0.0
    by_col: Dict[int, List[Tuple[Image.Image, int, int]]] = {i: [] for i in range(k)}
    for im, x, y, ki, _j in slots:
        by_col[ki].append((im, x, y))
    gap_sum = 0.0
    for ki in range(k - 1):
        left = by_col[ki]
        right = by_col[ki + 1]
        if not left or not right:
            continue
        right_of_left = max(x + im.width for im, x, _y in left)
        left_of_right = min(x for _im, x, _y in right)
        gap_sum += max(0, left_of_right - right_of_left)
    return gap_sum / max(1, usable)


def _build_cols(
    kinds: List[Image.Image],
    each: int,
    *,
    pitch_ratio: float,
    v_step_ratio: float,
    usable: int,
    tilt: float,
    pitch_min: float,
    pitch_max: float,
    uh_fixed: Optional[int] = None,
) -> Tuple[List[SlotK], int]:
    assert pitch_ratio + 1e-9 >= pitch_min
    assert pitch_ratio <= pitch_max + 1e-9
    k = len(kinds)
    probes = [_prep(u, 400, tilt) for u in kinds]
    pitch_probe = [max(1, int(probes[i].width * pitch_ratio)) for i in range(max(0, k - 1))]
    gw_probe = (sum(pitch_probe) + probes[-1].width) if k else 1
    gw_probe += int(probes[0].width * 0.04) * max(0, each - 1)
    uh0 = max(48, int(400 * (usable / max(1, gw_probe)) * 0.995))

    def place(uh: int) -> List[SlotK]:
        imgs = [_prep(u, uh, tilt) for u in kinds]
        widths = [im.width for im in imgs]
        pitches = [max(1, int(widths[i] * pitch_ratio)) for i in range(k - 1)]
        out: List[SlotK] = []
        x_cursor = 0
        for ki, base in enumerate(imgs):
            step = max(1, int(round(base.height * v_step_ratio)))
            for j in range(each):
                y = 0 if each == 1 else (each - 1 - j) * step
                out.append((base, x_cursor, y, ki, j))
            x_cursor += pitches[ki] if ki < k - 1 else base.width
        return out

    if uh_fixed is not None:
        uh0 = max(48, int(uh_fixed))
        return place(uh0), uh0

    slots = place(uh0)
    for _ in range(10):
        gw, gh = _slots_wh(slots)
        if gw <= usable + 1 and gh <= usable + 1:
            s = min(usable / max(1, gw), usable / max(1, gh))
            if s > 1.01:
                uh0 = max(48, int(uh0 * min(s, 1.06)))
                slots = place(uh0)
                continue
            break
        shrink = min(usable / max(1, gw), usable / max(1, gh))
        uh0 = max(48, int(uh0 * shrink * 0.995))
        slots = place(uh0)
    return slots, uh0


def _paste_group(slots: List[Slot]) -> Image.Image:
    max_x = max(x + im.width for im, x, y in slots)
    max_y = max(y + im.height for im, x, y in slots)
    min_x = min(x for im, x, y in slots)
    min_y = min(y for im, x, y in slots)
    layer = Image.new("RGBA", (max_x - min_x + 2, max_y - min_y + 2), (0, 0, 0, 0))
    for im, x, y in slots:
        layer.alpha_composite(im, (x - min_x, y - min_y))
    return _trim(layer)


def layout_mix_columns(
    kinds: List[Image.Image],
    each: int,
    *,
    canvas: int = 1200,
    cfg: Optional[Dict[str, Any]] = None,
) -> Tuple[List[Slot], Dict[str, Any]]:
    """
    戻り: (slots as (im,x,y) back-to-front), meta
    """
    c = dict(_cfg_defaults())
    if cfg:
        c.update(cfg)
    k = len(kinds)
    if k < 2 or each < 1:
        raise ValueError("mix columns require k>=2 and each>=1")
    # 親SKUが変わるたびに呼ぶ前提。内容指紋でも念のため親境界で消す
    clear_prep_cache()

    tilt = float(c["tiltDegCw"])
    h_ov_max = float(c["hOverlapMax"])
    pitch_min = 1.0 - h_ov_max
    pitch_max = float(c["pitchMax"])
    v_min = float(c["vStepMin"])
    v_max = float(c["vStepMax"])
    v_pref_lo = float(c["vStepPrefLo"])
    v_pref_hi = float(c["vStepPrefHi"])
    side_flush = bool(c.get("sideFlush", True))
    pitch_step = float(c.get("pitchStep", 0.03))
    v_step_step = float(c.get("vStepStep", 0.08))
    usable = canvas

    pitches: List[float] = []
    p = pitch_max
    while p + 1e-9 >= pitch_min:
        pitches.append(round(p, 3))
        p -= pitch_step
    if not pitches or abs(pitches[-1] - pitch_min) > 1e-9:
        pitches.append(round(pitch_min, 3))

    v_steps: List[float] = []
    v = v_min
    while v <= v_max + 1e-9:
        v_steps.append(round(v, 3))
        v += v_step_step

    def _prepare(
        pitch: float,
        v_step: float,
        *,
        do_flush: bool,
        uh_fixed: Optional[int] = None,
    ) -> Tuple[List[SlotK], int]:
        raw, uh = _build_cols(
            kinds,
            each,
            pitch_ratio=pitch,
            v_step_ratio=v_step,
            usable=usable,
            tilt=tilt,
            pitch_min=pitch_min,
            pitch_max=pitch_max,
            uh_fixed=uh_fixed,
        )
        min_x = min(s[1] for s in raw)
        min_y = min(s[2] for s in raw)
        norm = [(im, x - min_x, y - min_y, ki, j) for im, x, y, ki, j in raw]
        if do_flush and side_flush:
            norm = align_side_extremes(norm, k=k)
            min_x2 = min(s[1] for s in norm)
            min_y2 = min(s[2] for s in norm)
            norm = [
                (im, x - min_x2, y - min_y2, ki, j) for im, x, y, ki, j in norm
            ]
        return norm, uh

    best_key: Optional[tuple] = None
    best_pitch = pitch_max
    best_v = v_min
    best_meta: Dict[str, Any] = {}

    for pitch in pitches:
        for v_step in v_steps:
            norm, _uh = _prepare(pitch, v_step, do_flush=True)
            gw, gh = _slots_wh(norm)
            max_pad, h_pad, v_pad = _edge_pads(gw, gh, usable, canvas)
            aabb_gap = _internal_col_gap_ratio(norm, k=k, usable=usable)
            h_overlap = 1.0 - pitch
            s = min(usable / max(1, gw), usable / max(1, gh))
            bag_h = next(im.height for im, _x, _y, ki, _j in norm if ki == 0)
            disp_h = bag_h * s
            key = (
                -round(disp_h, 2),
                round(max_pad, 5),
                round(aabb_gap, 5),
                round(_v_pref_penalty(v_step, v_pref_lo, v_pref_hi), 5),
            )
            if best_key is None or key < best_key:
                best_key = key
                best_pitch = pitch
                best_v = v_step
                best_meta = {
                    "pitch": pitch,
                    "vStep": v_step,
                    "hOverlap": round(h_overlap, 4),
                    "dispBagH": round(disp_h, 2),
                    "predMaxEdge": round(max_pad, 4),
                    "predHPad": round(h_pad, 4),
                    "predVPad": round(v_pad, 4),
                    "group": [gw, gh],
                }

    pitch_use = best_pitch
    final, uh_lock = _prepare(pitch_use, best_v, do_flush=True)
    ink_gap = _ink_col_gap_px(final, k=k)
    closed = False
    if ink_gap > 0:
        for pitch in pitches:
            if pitch > pitch_use:
                continue
            cand, _ = _prepare(pitch, best_v, do_flush=True, uh_fixed=uh_lock)
            gw_c, gh_c = _slots_wh(cand)
            if gw_c > usable + 1 or gh_c > usable + 1:
                continue
            g = _ink_col_gap_px(cand, k=k)
            if g <= 0:
                final = cand
                pitch_use = pitch
                ink_gap = 0
                closed = True
                break
            if g < ink_gap:
                final = cand
                pitch_use = pitch
                ink_gap = g

    v_use = float(best_v)
    v_opened = False
    lo, hi = v_use, float(v_max)
    while hi - lo > 0.004:
        mid = round((lo + hi) / 2.0, 4)
        cand, _ = _prepare(pitch_use, mid, do_flush=True, uh_fixed=uh_lock)
        gw_c, gh_c = _slots_wh(cand)
        if gw_c <= usable + 1 and gh_c <= usable + 1:
            final = cand
            v_use = mid
            v_opened = mid > best_v + 1e-9
            lo = mid
        else:
            hi = mid

    gw, gh = _slots_wh(final)
    max_pad, h_pad, v_pad = _edge_pads(gw, gh, usable, canvas)
    bag_h = next(im.height for im, _x, _y, ki, _j in final if ki == 0)
    s = min(usable / max(1, gw), usable / max(1, gh))
    meta = {
        "engine": "mix_column_layout",
        "k": k,
        "each": each,
        "totalBags": k * each,
        "tiltDegCw": tilt,
        "hOverlapMaxHard": h_ov_max,
        "sideFlush": side_flush,
        "sideFlushAllColumns": True,
        "pitch": pitch_use,
        "vStep": round(v_use, 4),
        "vStepBeforeOpen": best_v,
        "vStepOpened": v_opened,
        "uhLock": uh_lock,
        "hOverlap": round(1.0 - pitch_use, 4),
        "inkColGapPx": int(_ink_col_gap_px(final, k=k)),
        "inkGapClosedByPitch": closed,
        "dispBagH": round(bag_h * s, 2),
        "predMaxEdge": round(max_pad, 4),
        "predHPad": round(h_pad, 4),
        "predVPad": round(v_pad, 4),
        "group": [gw, gh],
        "objective": "maxDispBagH then minEdgePad then uhLock openV",
    }
    assert meta["hOverlap"] <= h_ov_max + 1e-6

    # 合意ロック（v3b）と同じ: reverse(j, ki)。
    # 下前・上奥、かつ左列が右列の手前（左が重なりの上）。
    order = sorted(final, key=lambda t: (t[4], t[3]), reverse=True)
    slots = [(im, x, y) for im, x, y, _ki, _j in order]
    return slots, meta


def _octas_area_ratio_locked() -> float:
    """単品本線と同じ productAreaRatio（既定 0.075）。"""
    try:
        from work_paths import load_layout_rules

        tun = (load_layout_rules().get("amazon") or {}).get("tuningDraft") or {}
        oct = tun.get("octasSeal") or {}
        return float(oct.get("productAreaRatio", 0.075))
    except Exception:
        return 0.075


def compose_mix_columns(
    kinds: List[Image.Image],
    each: int,
    *,
    canvas: int = 1200,
    octas_rgba: Optional[Image.Image] = None,
    cfg: Optional[Dict[str, Any]] = None,
) -> Tuple[Image.Image, Dict[str, Any]]:
    """白キャンバスへ列積み。Octas は単品と同じ不透明面積比 0.075。"""
    slots, layout_meta = layout_mix_columns(kinds, each, canvas=canvas, cfg=cfg)
    group = _paste_group(slots)
    usable = canvas
    s = min(usable / max(1, group.width), usable / max(1, group.height))
    nw = max(1, int(round(group.width * s)))
    nh = max(1, int(round(group.height * s)))
    scaled = group.resize((nw, nh), Image.Resampling.LANCZOS)
    x = max(0, (usable - nw) // 2)
    y = max(0, (usable - nh) // 2)
    out = Image.new("RGBA", (canvas, canvas), (255, 255, 255, 255))
    out.alpha_composite(scaled, (x, y))

    octas_meta = None
    if octas_rgba is not None:
        from octas_prep import (
            count_opaque_pixels,
            octas_display_size_for_opaque_ratio,
            tilt_octas_rgba,
        )

        oc, tilt_m = tilt_octas_rgba(
            octas_rgba.convert("RGBA"),
            seed=hash(f"mix:{len(kinds)}:{each}") % (2**31),
            direction="top_to_right",
        )
        oc = _trim(oc)
        # 列全体の合算だと袋数分だけシールが膨らむ。
        # 単品N=1相当＝表示後の代表1袋の不透明 × 0.075。
        im0 = slots[0][0] if slots else scaled
        bw = max(1, int(round(im0.width * s)))
        bh = max(1, int(round(im0.height * s)))
        one_bag = im0.resize((bw, bh), Image.Resampling.LANCZOS)
        prod_op = count_opaque_pixels(one_bag)
        area_r = _octas_area_ratio_locked()
        tw, th, size_meta = octas_display_size_for_opaque_ratio(
            canvas_size=canvas,
            product_opaque_px=prod_op,
            seal_rgba=oc,
            area_ratio=area_r,
        )
        oc = oc.resize((max(8, tw), max(8, th)), Image.Resampling.LANCZOS)
        ox = canvas - oc.width
        oy = canvas - oc.height
        out.alpha_composite(oc, (ox, oy))
        octas_meta = {
            "placed": True,
            "tilt": tilt_m,
            "anchor": "bottom_right",
            "x": ox,
            "y": oy,
            "w": oc.width,
            "h": oc.height,
            "productAreaRatio": area_r,
            "sizeMode": size_meta.get("sizeMode"),
            "productOpaquePx": size_meta.get("productOpaquePx"),
            "productOpaqueBasis": "one_bag_display",
            "sealOpaquePx": size_meta.get("sealOpaquePx"),
            "sealOpaqueRatioActual": size_meta.get("sealOpaqueRatioActual"),
            "sizePolicyJa": (
                f"単品見た目基準: 表示後1袋の不透明×{area_r:.3f}"
                "（列全体合算しない）"
            ),
        }

    rgb = out.convert("RGB")
    meta = {
        "layoutMode": "mix_column",
        "aspectNoteJa": "ミックス列積み（正方形／縦長／横長メニュー共通）",
        "layout": layout_meta,
        "octas": octas_meta,
        "place": {"x": x, "y": y, "nw": nw, "nh": nh, "scale": round(s, 5)},
    }
    return rgb, meta
