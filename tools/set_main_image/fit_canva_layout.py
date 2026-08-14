# -*- coding: utf-8 -*-
"""最大10回、見本とのIoUで canvaUnitset を調整して1袋/10袋を出力。"""
from __future__ import annotations

import json
from pathlib import Path

from PIL import Image

from glyph_assets import compose_badge_with_glyphs, default_glyph_dirs
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"


def red_mask(im1000: Image.Image):
    pix = im1000.load()
    m = set()
    cx, cy, rad = 880, 120, 120
    for y in range(cy - rad, cy + rad + 1):
        for x in range(max(0, cx - rad), min(1000, cx + rad + 1)):
            if (x - cx) ** 2 + (y - cy) ** 2 > rad * rad:
                continue
            r, g, b = pix[x, y]
            if r > 70 and r > g + 25 and r > b + 25 and g < 110 and b < 110:
                m.add((x, y))
    return m


def iou(a, b):
    u = len(a | b)
    return (len(a & b) / u) if u else 0.0


def main():
    work_root = default_work_root()
    base_path = next(
        (next(p for p in work_root.iterdir() if p.is_dir() and p.name.startswith("02"))).glob(
            "*.jpg"
        )
    )
    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)

    assets = Path(
        r"C:/Users/takuy/.cursor/projects/c-Users-takuy-Desktop-gas-project/assets"
    )
    s1 = (
        Image.open(list(assets.glob("*oya_5*"))[0])
        .convert("RGB")
        .resize((1000, 1000), Image.Resampling.LANCZOS)
    )
    s10 = (
        Image.open(list(assets.glob("*oya_6*"))[0])
        .convert("RGB")
        .resize((1000, 1000), Image.Resampling.LANCZOS)
    )
    m1, m10 = red_mask(s1), red_mask(s10)

    rules = json.loads(RULES.read_text(encoding="utf-8"))
    typo = rules["rakuten"]["badgeTypography"]
    gold = rules["rakuten"]["goldCircle1200"]
    cx, cy, d = int(gold["cx"]), int(gold["cy"]), int(gold["diameter"])
    dirs = default_glyph_dirs(work_root)

    # Canva unitset 元幾何（1200）に基づく空洞
    hole_canva = [0.04, 0.145, 0.48, 0.40]

    # 見本外接（1000）
    # 1桁(6袋): overall (834,26)-(970,167)
    # 2桁(10袋): overall (806,43)-(983,172)
    trials_1 = [
        # iter: rect, allowStretch, pad, hole override
        ([834, 26, 136, 141], True, 0.05, hole_canva),
        ([830, 22, 145, 148], True, 0.04, hole_canva),
        ([828, 24, 142, 145], True, 0.06, [0.02, 0.12, 0.50, 0.42]),
        ([832, 20, 140, 150], True, 0.05, [0.03, 0.10, 0.48, 0.45]),
        ([826, 22, 150, 148], True, 0.03, [0.02, 0.12, 0.52, 0.43]),
        ([834, 26, 136, 141], False, 0.05, hole_canva),
        ([820, 18, 160, 155], False, 0.04, hole_canva),
        ([838, 28, 132, 138], True, 0.08, [0.05, 0.14, 0.45, 0.40]),
        ([830, 24, 148, 146], True, 0.05, [0.01, 0.11, 0.55, 0.44]),
        ([832, 22, 144, 150], True, 0.04, [0.02, 0.10, 0.50, 0.46]),
    ]
    trials_10 = [
        ([806, 43, 177, 130], True, 0.02, hole_canva),
        ([800, 38, 185, 138], True, 0.02, hole_canva),
        ([798, 40, 190, 136], True, 0.01, [0.02, 0.12, 0.58, 0.42]),
        ([804, 42, 180, 134], True, 0.02, [0.01, 0.11, 0.60, 0.43]),
        ([810, 40, 175, 136], True, 0.02, [0.03, 0.12, 0.55, 0.42]),
        ([806, 43, 177, 130], False, 0.02, hole_canva),
        ([790, 36, 200, 142], False, 0.01, hole_canva),
        ([802, 38, 185, 140], True, 0.00, [0.00, 0.10, 0.62, 0.45]),
        ([808, 44, 172, 128], True, 0.03, [0.02, 0.13, 0.56, 0.40]),
        ([800, 40, 188, 138], True, 0.02, [0.01, 0.11, 0.58, 0.44]),
    ]

    def run_trial(set_count, key, rect, stretch, pad, hole):
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
    print("=== fit 1digit using count=6 vs sample 6袋 ===")
    for i, (rect, stretch, pad, hole) in enumerate(trials_1, 1):
        out, g1000, meta, t = run_trial(6, "1digit", rect, stretch, pad, hole)
        score = iou(m1, red_mask(g1000))
        print(f"i{i:02d}", round(score, 4), "rect", rect, "stretch", stretch)
        if score > best1[0]:
            best1 = (score, (rect, stretch, pad, hole, out, meta, t))

    best10 = (-1, None)
    print("=== fit 2digit using count=10 vs sample 10袋 ===")
    for i, (rect, stretch, pad, hole) in enumerate(trials_10, 1):
        out, g1000, meta, t = run_trial(10, "2digit", rect, stretch, pad, hole)
        score = iou(m10, red_mask(g1000))
        print(f"i{i:02d}", round(score, 4), "rect", rect, "stretch", stretch)
        if score > best10[0]:
            best10 = (score, (rect, stretch, pad, hole, out, meta, t))

    # second pass: local refine around best (±3 px / hole tweaks)
    def refine(set_count, key, sample_m, seed, pad0):
        rect0, stretch0, _, hole0, _, _, _ = seed
        best = (-1, None)
        for dx in (-3, 0, 3):
            for dy in (-3, 0, 3):
                for dw in (-4, 0, 4):
                    for dh in (-4, 0, 4):
                        rect = [rect0[0] + dx, rect0[1] + dy, rect0[2] + dw, rect0[3] + dh]
                        for dhw in (-0.04, 0, 0.04):
                            for dhh in (-0.04, 0, 0.04):
                                hole = [
                                    max(0, hole0[0]),
                                    max(0, hole0[1]),
                                    max(0.35, hole0[2] + dhw),
                                    max(0.30, hole0[3] + dhh),
                                ]
                                if hole[0] + hole[2] > 0.75:
                                    continue
                                out, g1000, meta, t = run_trial(
                                    set_count, key, rect, stretch0, pad0, hole
                                )
                                score = iou(sample_m, red_mask(g1000))
                                if score > best[0]:
                                    best = (score, (rect, stretch0, pad0, hole, out, meta, t))
        return best

    print("refine1...")
    r1 = refine(6, "1digit", m1, best1[1], best1[1][2])
    if r1[0] > best1[0]:
        best1 = r1
        print("refined1", round(best1[0], 4), best1[1][0])
    print("refine10...")
    r10 = refine(10, "2digit", m10, best10[1], best10[1][2])
    if r10[0] > best10[0]:
        best10 = r10
        print("refined10", round(best10[0], 4), best10[1][0])

    # persist best into rules
    _, b1 = best1
    _, b10 = best10
    typo["canvaUnitset"]["1digit"] = {
        "unitsetRect1000": b1[0],
        "numberHoleNorm": b1[3],
        "numberPadLeftRatio": b1[2],
        "allowStretch": b1[1],
        "align": "top_left",
        "note": f"IoU={best1[0]:.3f} fit count=6 vs sample 6袋 → apply to 1袋",
    }
    typo["canvaUnitset"]["2digit"] = {
        "unitsetRect1000": b10[0],
        "numberHoleNorm": b10[3],
        "numberPadLeftRatio": b10[2],
        "allowStretch": b10[1],
        "align": "top_left",
        "note": f"IoU={best10[0]:.3f} fit count=10 vs sample 10袋",
    }
    typo["canvaUnitset"]["enabled"] = True
    rules["rakuten"]["badgeTypography"] = typo
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    # final export: 1袋 + 10袋 (+ reference 6袋 for QA)
    out_dir = work_root / "00.テスト出力"

    def export(name, set_count, key, conf):
        rect, stretch, pad, hole, _, _, _ = conf
        out, g1000, meta, _ = run_trial(set_count, key, rect, stretch, pad, hole)
        path = out_dir / name
        out.convert("RGB").save(path, quality=95)
        return path, meta, g1000

    p1, meta1, g1 = export("POC10_1fukuro_set.jpg", 1, "1digit", b1)
    p6, meta6, g6 = export("POC10_6fukuro_fitcheck.jpg", 6, "1digit", b1)
    p10, meta10, g10 = export("POC10_10fukuro_set.jpg", 10, "2digit", b10)

    g1.crop((740, 0, 1000, 260)).save("tmp_out1.png")
    g6.crop((740, 0, 1000, 260)).save("tmp_out6.png")
    g10.crop((740, 0, 1000, 260)).save("tmp_out10.png")
    s1.crop((740, 0, 1000, 260)).save("tmp_samp1.png")
    s10.crop((740, 0, 1000, 260)).save("tmp_samp10.png")

    print("BEST6fit IoU", round(best1[0], 4), "→ used for 1袋", b1[0], b1[3])
    print("BEST10 IoU", round(best10[0], 4), b10[0], b10[3])
    print("SAVED", p1)
    print("SAVED", p10)
    print("CHECK", p6)
    print("meta1", meta1.get("numberPos"), meta1.get("unitsetRect"))
    print("meta10", meta10.get("numberPos"), meta10.get("unitsetRect"))
    print("IoU6", round(iou(m1, red_mask(g6)), 4), "IoU10", round(iou(m10, red_mask(g10)), 4))


if __name__ == "__main__":
    main()
