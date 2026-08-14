# -*- coding: utf-8 -*-
"""1袋・10袋の本物見本と重ね合わせしながら最大5回調整。"""
from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw

from glyph_assets import compose_badge_with_glyphs, default_glyph_dirs
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"
OUT_QA = Path(__file__).resolve().parent / "tmp_overlay_qa"
OUT_QA.mkdir(exist_ok=True)


def red_mask(im1000: Image.Image):
    pix = im1000.load()
    m = set()
    cx, cy, rad = 880, 120, 122
    for y in range(max(0, cy - rad), min(1000, cy + rad + 1)):
        for x in range(max(0, cx - rad), min(1000, cx + rad + 1)):
            if (x - cx) ** 2 + (y - cy) ** 2 > rad * rad:
                continue
            r, g, b = pix[x, y]
            if r > 70 and r > g + 25 and r > b + 25 and g < 110 and b < 110:
                m.add((x, y))
    return m


def bbox(pts):
    if not pts:
        return None
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys), max(xs) - min(xs) + 1, max(ys) - min(ys) + 1


def measure_parts(im1000: Image.Image):
    pts = red_mask(im1000)
    upper = [p for p in pts if p[1] < 145]
    lower = [p for p in pts if p[1] >= 145]
    best = None
    for gx in range(860, 960):
        L = [p for p in upper if p[0] < gx]
        R = [p for p in upper if p[0] >= gx]
        if len(L) < 20 or len(R) < 15:
            continue
        lh = max(p[1] for p in L) - min(p[1] for p in L) + 1
        rh = max(p[1] for p in R) - min(p[1] for p in R) + 1
        if lh <= rh:
            continue
        gap = min(p[0] for p in R) - max(p[0] for p in L)
        if gap < 1:
            continue
        score = lh * 2 + gap
        if best is None or score > best[0]:
            best = (score, bbox(L), bbox(R))
    return {
        "all": bbox(pts),
        "num": best[1] if best else None,
        "unit": best[2] if best else None,
        "set": bbox(lower) if lower else None,
        "mask": pts,
    }


def iou(a, b):
    u = len(a | b)
    return (len(a & b) / u) if u else 0.0


def overlay(sample, gen, path: Path, title: str):
    """見本=赤、生成=緑、重なり=黄。"""
    sm = red_mask(sample)
    gm = red_mask(gen)
    canvas = Image.new("RGB", (260, 280), (35, 35, 35))
    crop_s = sample.crop((740, 0, 1000, 260)).convert("RGBA")
    crop_g = gen.crop((740, 0, 1000, 260)).convert("RGBA")
    # side by side top: sample | gen
    sheet = Image.new("RGB", (540, 560), (30, 30, 30))
    sheet.paste(crop_s.convert("RGB"), (10, 30))
    sheet.paste(crop_g.convert("RGB"), (280, 30))
    # overlay bottom
    ov = Image.new("RGB", (260, 260), (25, 25, 25))
    op = ov.load()
    for y in range(260):
        for x in range(260):
            sx, sy = x + 740, y
            a = (sx, sy) in sm
            b = (sx, sy) in gm
            if a and b:
                op[x, y] = (255, 220, 40)
            elif a:
                op[x, y] = (255, 70, 70)
            elif b:
                op[x, y] = (70, 220, 90)
    sheet.paste(ov, (140, 300))
    draw = ImageDraw.Draw(sheet)
    draw.text((10, 8), f"{title}  SAMPLE", fill=(255, 200, 200))
    draw.text((280, 8), f"{title}  GEN", fill=(200, 255, 200))
    score = iou(sm, gm)
    draw.text((140, 290), f"OVERLAY yellow=overlap IoU={score:.3f}  red=sample green=gen", fill=(220, 220, 220))
    sheet.save(path, quality=92)
    return score


def main():
    work_root = default_work_root()
    ref = next(p for p in work_root.iterdir() if p.is_dir() and p.name.startswith("04"))
    base_path = next(
        (next(p for p in work_root.iterdir() if p.is_dir() and p.name.startswith("02"))).glob(
            "*.jpg"
        )
    )
    sample_1 = Image.open(ref / "sanky-4906283045614-oya.jpg").convert("RGB").resize(
        (1000, 1000), Image.Resampling.LANCZOS
    )
    sample_10 = Image.open(ref / "sanky-4906283045614-oya_6.jpg").convert("RGB").resize(
        (1000, 1000), Image.Resampling.LANCZOS
    )
    m1 = measure_parts(sample_1)
    m10 = measure_parts(sample_10)
    print("SAMPLE 1袋", {k: v for k, v in m1.items() if k != "mask"})
    print("SAMPLE 10袋", {k: v for k, v in m10.items() if k != "mask"})

    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)

    rules = json.loads(RULES.read_text(encoding="utf-8"))
    typo = rules["rakuten"]["badgeTypography"]
    gold = rules["rakuten"]["goldCircle1200"]
    cx, cy, d = int(gold["cx"]), int(gold["cy"]), int(gold["diameter"])
    dirs = default_glyph_dirs(work_root)

    # 見本外接から初期 rect
    # 1袋 all
    a1 = m1["all"]
    a10 = m10["all"]
    # trials: progressively nudge from measured overall frame
    trials_1 = []
    if a1:
        x, y, _, _, w, h = a1
        trials_1 = [
            ([x, y, w, h], True, 0.05, [0.02, 0.12, 0.48, 0.42]),
            ([x - 4, y - 2, w + 8, h + 6], True, 0.04, [0.02, 0.10, 0.50, 0.44]),
            ([x + 2, y, w - 2, h + 4], True, 0.06, [0.03, 0.11, 0.46, 0.45]),
            ([x - 2, y + 2, w + 4, h + 2], True, 0.05, [0.01, 0.10, 0.52, 0.46]),
            ([x, y - 4, w + 6, h + 8], True, 0.04, [0.02, 0.08, 0.50, 0.48]),
        ]
    trials_10 = []
    if a10:
        x, y, _, _, w, h = a10
        trials_10 = [
            ([x, y, w, h], True, 0.02, [0.02, 0.12, 0.55, 0.42]),
            ([x - 4, y - 2, w + 8, h + 6], True, 0.02, [0.02, 0.10, 0.58, 0.44]),
            ([x + 2, y, w - 2, h + 4], True, 0.01, [0.03, 0.11, 0.54, 0.45]),
            ([x - 2, y + 2, w + 4, h + 2], True, 0.02, [0.01, 0.10, 0.56, 0.46]),
            ([x, y - 4, w + 6, h + 8], True, 0.02, [0.02, 0.09, 0.57, 0.47]),
        ]

    def run(set_count, key, rect, stretch, pad, hole):
        t = json.loads(json.dumps(typo))
        t["canvaUnitset"][key]["unitsetRect1000"] = list(rect)
        t["canvaUnitset"][key]["numberHoleNorm"] = list(hole)
        t["canvaUnitset"][key]["numberPadLeftRatio"] = pad
        t["canvaUnitset"][key]["allowStretch"] = stretch
        t["canvaUnitset"][key]["align"] = "top_left"
        out, meta = compose_badge_with_glyphs(
            base0.copy(),
            set_count=set_count,
            unit="袋",
            cx=cx,
            cy=cy,
            diameter=d,
            typo=t,
            glyph_dirs=dirs,
            canvas=1200,
        )
        g1000 = out.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
        return out, g1000, meta, t

    best1 = (-1, None)
    print("=== 1袋 × sample oya.jpg (max 5) ===")
    for i, (rect, stretch, pad, hole) in enumerate(trials_1, 1):
        out, g1000, meta, t = run(1, "1digit", rect, stretch, pad, hole)
        score = overlay(sample_1, g1000, OUT_QA / f"iter1_{i:02d}.jpg", f"1袋 iter{i}")
        print(f"iter{i}", round(score, 4), rect, hole)
        if score > best1[0]:
            best1 = (score, (rect, stretch, pad, hole, out, meta, t, g1000))

    best10 = (-1, None)
    print("=== 10袋 × sample oya_6.jpg (max 5) ===")
    for i, (rect, stretch, pad, hole) in enumerate(trials_10, 1):
        out, g1000, meta, t = run(10, "2digit", rect, stretch, pad, hole)
        score = overlay(sample_10, g1000, OUT_QA / f"iter10_{i:02d}.jpg", f"10袋 iter{i}")
        print(f"iter{i}", round(score, 4), rect, hole)
        if score > best10[0]:
            best10 = (score, (rect, stretch, pad, hole, out, meta, t, g1000))

    # save best rules + finals
    b1 = best1[1]
    b10 = best10[1]
    typo["canvaUnitset"]["1digit"] = {
        "unitsetRect1000": b1[0],
        "numberHoleNorm": [round(x, 3) for x in b1[3]],
        "numberPadLeftRatio": b1[2],
        "allowStretch": b1[1],
        "align": "top_left",
        "note": f"IoU={best1[0]:.3f} vs sanky-...-oya.jpg (1袋)",
    }
    typo["canvaUnitset"]["2digit"] = {
        "unitsetRect1000": b10[0],
        "numberHoleNorm": [round(x, 3) for x in b10[3]],
        "numberPadLeftRatio": b10[2],
        "allowStretch": b10[1],
        "align": "top_left",
        "note": f"IoU={best10[0]:.3f} vs sanky-...-oya_6.jpg (10袋)",
    }
    typo["canvaUnitset"]["enabled"] = True
    rules["rakuten"]["badgeTypography"] = typo
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    out_dir = work_root / "00.テスト出力"
    p1 = out_dir / "POC11_1fukuro_set.jpg"
    p10 = out_dir / "POC11_10fukuro_set.jpg"
    b1[4].convert("RGB").save(p1, quality=95)
    b10[4].convert("RGB").save(p10, quality=95)
    # best overlays
    overlay(sample_1, b1[7], out_dir / "POC11_1fukuro_OVERLAY.jpg", "1袋 BEST")
    overlay(sample_10, b10[7], out_dir / "POC11_10fukuro_OVERLAY.jpg", "10袋 BEST")

    print("BEST1", round(best1[0], 4), b1[0], b1[3])
    print("BEST10", round(best10[0], 4), b10[0], b10[3])
    print("OUT", p1)
    print("OUT", p10)
    print("OVERLAY", out_dir / "POC11_1fukuro_OVERLAY.jpg")
    print("OVERLAY", out_dir / "POC11_10fukuro_OVERLAY.jpg")
    print("QA iters", OUT_QA)


if __name__ == "__main__":
    main()
