# -*- coding: utf-8 -*-
"""
単位＋セットだけを見本外接に固定配置し、大きさ・位置IoUを最大化（数字は採点外）。
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PIL import Image, ImageDraw

from glyph_assets import (
    _crop_to_alpha,
    default_glyph_dirs,
    load_unitset_glyph,
    paste_glyph,
)
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"
OUT_QA = Path(__file__).resolve().parent / "tmp_unitset_lock"
OUT_QA.mkdir(exist_ok=True)


def red_mask(im1000: Image.Image, x_min: int = 0):
    pix = im1000.load()
    m = set()
    cx, cy, rad = 880, 120, 122
    for y in range(max(0, cy - rad), min(1000, cy + rad + 1)):
        for x in range(max(x_min, cx - rad), min(1000, cx + rad + 1)):
            if (x - cx) ** 2 + (y - cy) ** 2 > rad * rad:
                continue
            r, g, b = pix[x, y]
            if r > 70 and r > g + 25 and r > b + 25 and g < 110 and b < 110:
                m.add((x, y))
    return m


def iou(a, b) -> float:
    u = len(a | b)
    return (len(a & b) / u) if u else 0.0


def bbox_of(mask, region=None):
    pts = [p for p in mask if (region is None or region(p))]
    if not pts:
        return None
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys), max(xs) - min(xs) + 1, max(ys) - min(ys) + 1


def size_pos_score(sample_bb, gen_bb) -> Tuple[float, float, float]:
    """大きさ一致率・位置一致率・総合(平均)。"""
    if not sample_bb or not gen_bb:
        return 0.0, 0.0, 0.0
    sw, sh = sample_bb[4], sample_bb[5]
    gw, gh = gen_bb[4], gen_bb[5]
    size = 1.0 - min(1.0, (abs(sw - gw) / max(1, sw) + abs(sh - gh) / max(1, sh)) / 2)
    scx, scy = (sample_bb[0] + sample_bb[2]) / 2, (sample_bb[1] + sample_bb[3]) / 2
    gcx, gcy = (gen_bb[0] + gen_bb[2]) / 2, (gen_bb[1] + gen_bb[3]) / 2
    # 許容ズレ: 対角の相対
    diag = (sw**2 + sh**2) ** 0.5
    dist = ((scx - gcx) ** 2 + (scy - gcy) ** 2) ** 0.5
    pos = 1.0 - min(1.0, dist / max(1.0, diag * 0.5))
    return size, pos, (size + pos) / 2


def split_unit_set(unitset: Image.Image) -> Tuple[Image.Image, Image.Image]:
    """unitset から単位(右上)とセット(下)を分離。"""
    full = unitset.convert("RGBA")
    cropped = _crop_to_alpha(full)
    w, h = cropped.size
    pix = cropped.load()
    # unit: upper half, right 55%
    unit_pts = []
    set_pts = []
    for y in range(h):
        for x in range(w):
            if pix[x, y][3] <= 20:
                continue
            if y < h * 0.55 and x >= w * 0.45:
                unit_pts.append((x, y))
            if y >= h * 0.55:
                set_pts.append((x, y))

    def crop_pts(pts):
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        box = (min(xs), min(ys), max(xs) + 1, max(ys) + 1)
        return cropped.crop(box)

    return crop_pts(unit_pts), crop_pts(set_pts)


def paste_box(base: Image.Image, glyph: Image.Image, box1000, canvas=1200, grid=1000):
    """box=[x,y,w,h] on 1000 → exact stretch paste on canvas."""
    scale = canvas / grid
    x = int(box1000[0] * scale)
    y = int(box1000[1] * scale)
    w = max(1, int(box1000[2] * scale))
    h = max(1, int(box1000[3] * scale))
    g = _crop_to_alpha(glyph.convert("RGBA")).resize((w, h), Image.Resampling.LANCZOS)
    paste_glyph(base, g, (x, y))
    return {"x": x, "y": y, "w": w, "h": h}


def compose_locked(
    base: Image.Image,
    *,
    unit_im: Image.Image,
    set_im: Image.Image,
    unit_box,
    set_box,
    canvas=1200,
) -> Image.Image:
    work = base.convert("RGBA")
    paste_box(work, unit_im, unit_box, canvas=canvas)
    paste_box(work, set_im, set_box, canvas=canvas)
    return work


def overlay_unit_set(sample, gen, path, title, x_min):
    sm = red_mask(sample, x_min=x_min)
    # also include set which may start left of unit
    sm |= red_mask(sample, x_min=0)  # full then filter
    # rebuild: unit region x>=unit_xmin OR y>=140
    sm = set()
    pix = sample.load()
    cx, cy, rad = 880, 120, 122
    for y in range(cy - rad, cy + rad + 1):
        for x in range(max(0, cx - rad), min(1000, cx + rad + 1)):
            if (x - cx) ** 2 + (y - cy) ** 2 > rad * rad:
                continue
            r, g, b = pix[x, y]
            if not (r > 70 and r > g + 25 and r > b + 25 and g < 110 and b < 110):
                continue
            if x >= x_min or y >= 140:  # unit or set band
                # exclude tall number stem: left of x_min and y<140
                if y < 140 and x < x_min:
                    continue
                sm.add((x, y))
    gm = set()
    pix = gen.load()
    for y in range(cy - rad, cy + rad + 1):
        for x in range(max(0, cx - rad), min(1000, cx + rad + 1)):
            if (x - cx) ** 2 + (y - cy) ** 2 > rad * rad:
                continue
            r, g, b = pix[x, y]
            if not (r > 70 and r > g + 25 and r > b + 25 and g < 110 and b < 110):
                continue
            if y < 140 and x < x_min:
                continue
            if x >= x_min or y >= 140:
                gm.add((x, y))

    sheet = Image.new("RGB", (540, 560), (30, 30, 30))
    sheet.paste(sample.crop((740, 0, 1000, 260)), (10, 30))
    sheet.paste(gen.crop((740, 0, 1000, 260)), (280, 30))
    ov = Image.new("RGB", (260, 260), (25, 25, 25))
    op = ov.load()
    for y in range(260):
        for x in range(260):
            sx, sy = x + 740, y
            a, b = (sx, sy) in sm, (sx, sy) in gm
            if a and b:
                op[x, y] = (255, 220, 40)
            elif a:
                op[x, y] = (255, 70, 70)
            elif b:
                op[x, y] = (70, 220, 90)
    sheet.paste(ov, (140, 300))
    score = iou(sm, gm)
    draw = ImageDraw.Draw(sheet)
    draw.text((10, 8), f"{title} SAMPLE", fill=(255, 200, 200))
    draw.text((280, 8), f"{title} GEN (unit+set only)", fill=(200, 255, 200))
    draw.text((120, 290), f"OVERLAY IoU={score:.3f} (number ignored)", fill=(220, 220, 220))
    sheet.save(path, quality=92)
    return score, sm, gm


def main():
    work_root = default_work_root()
    ref = next(p for p in work_root.iterdir() if p.is_dir() and p.name.startswith("04"))
    base_path = next(
        (next(p for p in work_root.iterdir() if p.is_dir() and p.name.startswith("02"))).glob(
            "*.jpg"
        )
    )
    sample_1 = (
        Image.open(ref / "sanky-4906283045614-oya.jpg")
        .convert("RGB")
        .resize((1000, 1000), Image.Resampling.LANCZOS)
    )
    sample_10 = (
        Image.open(ref / "sanky-4906283045614-oya_6.jpg")
        .convert("RGB")
        .resize((1000, 1000), Image.Resampling.LANCZOS)
    )

    # locked boxes from sample measure
    boxes = {
        "1digit": {
            "unitBox": [900, 58, 72, 87],
            "setBox": [834, 140, 136, 28],
            "numExcludeX": 900,
        },
        "2digit": {
            "unitBox": [923, 64, 62, 81],
            "setBox": [848, 140, 135, 34],
            "numExcludeX": 920,
        },
    }

    dirs = default_glyph_dirs(work_root)
    unitset = load_unitset_glyph("袋", dirs)
    assert unitset is not None
    unit_im, set_im = split_unit_set(unitset)
    unit_im.save(OUT_QA / "split_unit.png")
    set_im.save(OUT_QA / "split_set.png")

    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)

    def search(key, sample, max_iter=20):
        ub0 = boxes[key]["unitBox"][:]
        sb0 = boxes[key]["setBox"][:]
        xmin = boxes[key]["numExcludeX"]
        best = (-1.0, None, None, None, None)
        # generate candidates: nudge unit/set boxes
        cands = []
        for du in range(-4, 5):
            for dvs in range(-3, 4):
                for dsw in (-4, 0, 4, 8, -8):
                    for dsh in (-2, 0, 2, 4, -4):
                        for dx in (-3, 0, 3):
                            for dy in (-2, 0, 2):
                                ub = [
                                    ub0[0] + dx,
                                    ub0[1] + dy,
                                    max(40, ub0[2] + du),
                                    max(40, ub0[3] + dvs),
                                ]
                                sb = [
                                    sb0[0] + dx,
                                    sb0[1] + dy,
                                    max(80, sb0[2] + dsw),
                                    max(16, sb0[3] + dsh),
                                ]
                                cands.append((ub, sb))
        # subsample to ~20 best-spaced + always include exact
        cands = [(ub0, sb0)] + cands[:: max(1, len(cands) // 19)]
        cands = cands[:20]
        print(f"=== {key} trials {len(cands)} ===")
        for i, (ub, sb) in enumerate(cands, 1):
            out = compose_locked(base0.copy(), unit_im=unit_im, set_im=set_im, unit_box=ub, set_box=sb)
            g = out.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
            score, sm, gm = overlay_unit_set(
                sample, g, OUT_QA / f"{key}_iter{i:02d}.jpg", f"{key} i{i}", xmin
            )
            # size/pos vs target boxes (expected)
            gu = bbox_of(gm, lambda p: p[1] < 145 and p[0] >= xmin - 5)
            gs = bbox_of(gm, lambda p: p[1] >= 140)
            su = bbox_of(sm, lambda p: p[1] < 145 and p[0] >= xmin - 5)
            ss = bbox_of(sm, lambda p: p[1] >= 140)
            sz_u, pos_u, _ = size_pos_score(su, gu)
            sz_s, pos_s, _ = size_pos_score(ss, gs)
            size_m = (sz_u + sz_s) / 2
            pos_m = (pos_u + pos_s) / 2
            combo = (score + size_m + pos_m) / 3
            print(
                f"i{i:02d} IoU={score:.3f} size={size_m:.3f} pos={pos_m:.3f} combo={combo:.3f} ub={ub} sb={sb}"
            )
            if combo > best[0]:
                best = (combo, ub, sb, out, (score, size_m, pos_m))
            if size_m >= 0.98 and pos_m >= 0.98 and score >= 0.98:
                print("*** REACHED 98% on size+pos+IoU ***")
                break
        return best

    best1 = search("1digit", sample_1, 20)
    best10 = search("2digit", sample_10, 20)

    # persist
    rules = json.loads(RULES.read_text(encoding="utf-8"))
    typo = rules["rakuten"]["badgeTypography"]
    typo["lockedUnitSet1000"] = {
        "enabled": True,
        "1digit": {
            "unitBox": best1[1],
            "setBox": best1[2],
            "note": f"combo={best1[0]:.3f} metrics={best1[4]}",
        },
        "2digit": {
            "unitBox": best10[1],
            "setBox": best10[2],
            "note": f"combo={best10[0]:.3f} metrics={best10[4]}",
        },
    }
    rules["rakuten"]["badgeTypography"] = typo
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    out_dir = work_root / "00.テスト出力"
    from glyph_assets import load_digit_glyph, load_pair_glyph

    def export_with_number(path, set_count, ub, sb, nb):
        work = compose_locked(base0.copy(), unit_im=unit_im, set_im=set_im, unit_box=ub, set_box=sb)
        if set_count < 10:
            num = _crop_to_alpha(load_digit_glyph(str(set_count), dirs))
        else:
            num = load_pair_glyph(set_count, dirs)
            if num is None:
                num = _crop_to_alpha(load_digit_glyph("1", dirs))
            else:
                num = _crop_to_alpha(num)
        if num is not None and nb:
            paste_box(work, num, nb)
        work.convert("RGB").save(path, quality=95)
        return work

    n1 = [845, 26, 54, 119]
    n10 = [806, 43, 113, 102]
    p1 = out_dir / "POC12_1fukuro_unitset_locked.jpg"
    p10 = out_dir / "POC12_10fukuro_unitset_locked.jpg"
    export_with_number(p1, 1, best1[1], best1[2], n1)
    export_with_number(p10, 10, best10[1], best10[2], n10)

    # overlays of best unit+set-only
    g1 = best1[3].convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
    g10 = best10[3].convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
    overlay_unit_set(sample_1, g1, out_dir / "POC12_1_OVERLAY_unitset.jpg", "1袋 unit+set", 900)
    overlay_unit_set(sample_10, g10, out_dir / "POC12_10_OVERLAY_unitset.jpg", "10袋 unit+set", 920)

    print("BEST1 combo", round(best1[0], 4), "metrics", best1[4], "unit", best1[1], "set", best1[2])
    print("BEST10 combo", round(best10[0], 4), "metrics", best10[4], "unit", best10[1], "set", best10[2])
    print("OUT", p1)
    print("OUT", p10)
    s1, p1m, c1 = best1[4]
    s10, p10m, c10 = best10[4]
    for label, size_m, pos_m, sc in (
        ("1袋", best1[4][1], best1[4][2], best1[4][0]),
        ("10袋", best10[4][1], best10[4][2], best10[4][0]),
    ):
        ok = size_m >= 0.98 and pos_m >= 0.98
        print(f"{label} size={size_m:.1%} pos={pos_m:.1%} IoU={sc:.1%} 98%={'YES' if ok else 'NO'}")


if __name__ == "__main__":
    main()
