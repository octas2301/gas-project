# -*- coding: utf-8 -*-
"""2桁pair位置を固定。digit組合せはpair numberBoxの右縁・下縁に合わせる。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional, Tuple

from PIL import Image, ImageDraw

from glyph_assets import compose_badge_with_glyphs, default_glyph_dirs
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"
BBox = Tuple[int, int, int, int, int, int]


def alpha_bbox_in_region(rgb1000: Image.Image, region: List[int], thr_sum: int = 280) -> Optional[BBox]:
    """RGB画像の暗いインク外接（数字帯）。"""
    x, y, w, h = [int(v) for v in region]
    pix = rgb1000.load()
    xs: List[int] = []
    ys: List[int] = []
    for yy in range(max(0, y - 2), min(1000, y + h + 2)):
        for xx in range(max(0, x - 2), min(1000, x + w + 40)):
            r, g, b = pix[xx, yy]
            if r + g + b < thr_sum and r > 40:
                xs.append(xx)
                ys.append(yy)
    if not xs:
        return None
    return min(xs), min(ys), max(xs), max(ys), max(xs) - min(xs) + 1, max(ys) - min(ys) + 1


def main() -> None:
    root = default_work_root()
    base_path = next(
        (next(p for p in root.iterdir() if p.is_dir() and p.name.startswith("02"))).glob("*.jpg")
    )
    out_dir = root / "00.テスト出力"
    dirs = default_glyph_dirs(root)

    rules = json.loads(RULES.read_text(encoding="utf-8"))
    typo = rules["rakuten"]["badgeTypography"]
    locked = typo["lockedUnitSet1000"]
    d2 = locked["2digit"]

    # pair位置を完全固定
    d2["numberBox"] = [818, 30, 103, 85]
    d2["freezeUnitSet"] = True
    d2["freezeNumber"] = True
    d2["preferPair"] = True
    d2["numberFitMode"] = "contain"
    d2["numberSource"] = "pair_10"
    d2["noteNumber"] = (
        "FROZEN pair_10 @ [818,30,103,85] — digit compose aligns right/bottom to this box"
    )
    locked["2digit"] = d2

    typo["digitCompose2"] = {
        "preferPair": True,
        "alignRightBottomToNumberBox": True,
        "gapRatioOfHeight": 0.04,
        "note": (
            "pair無し時は digit_X+digit_Y を等倍で横並び。"
            "右桁の右縁＝numberBox右縁、下端＝numberBox下端（pair版と同基準）。"
        ),
    }
    typo["lockedUnitSet1000"] = locked
    rules["rakuten"]["badgeTypography"] = typo
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)
    gold = rules["rakuten"]["goldCircle1200"]
    box = d2["numberBox"]
    box_right = box[0] + box[2]
    box_bottom = box[1] + box[3]

    def run(force_digits: bool):
        t = json.loads(json.dumps(typo))
        t["lockedUnitSet1000"]["2digit"]["forceDigits"] = force_digits
        t["lockedUnitSet1000"]["2digit"]["preferPair"] = not force_digits
        out, meta = compose_badge_with_glyphs(
            base0.copy(),
            set_count=10,
            unit="缶",
            cx=int(gold["cx"]),
            cy=int(gold["cy"]),
            diameter=int(gold["diameter"]),
            typo=t,
            glyph_dirs=dirs,
            canvas=1200,
        )
        return out, meta

    pair_im, pair_meta = run(False)
    dig_im, dig_meta = run(True)
    print("pair", pair_meta.get("numSource"), pair_meta.get("numberPos"), pair_meta.get("numberFitMode"))
    print("digits", dig_meta.get("numSource"), dig_meta.get("numberPos"), dig_meta.get("numberFitMode"))

    pair_im.convert("RGB").save(out_dir / "POC19_10kan_PAIR_locked.jpg", quality=95)
    dig_im.convert("RGB").save(out_dir / "POC19_10kan_DIGITS.jpg", quality=95)

    # 右縁・下端の一致確認（1200→1000）
    def pos1000(p):
        if not p:
            return None
        s = 1000 / 1200
        x0 = int(p["x"] * s)
        y0 = int(p["y"] * s)
        x1 = int((p["x"] + p["w"] - 1) * s)
        y1 = int((p["y"] + p["h"] - 1) * s)
        return x0, y0, x1, y1

    pp = pos1000(pair_meta.get("numberPos"))
    dd = pos1000(dig_meta.get("numberPos"))
    print("pair ink1000", pp, "right", pp[2] if pp else None, "bottom", pp[3] if pp else None)
    print("dig  ink1000", dd, "right", dd[2] if dd else None, "bottom", dd[3] if dd else None)
    print("target box right/bottom", box_right, box_bottom)
    if pp and dd:
        print("right delta (dig-pair)", dd[2] - pp[2], "bottom delta", dd[3] - pp[3])

    # 並べてオーバーレイ
    pr = pair_im.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
    dr = dig_im.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
    sheet = Image.new("RGB", (560, 360), (28, 28, 28))
    left = pr.crop((740, 0, 1000, 260)).copy()
    right = dr.crop((740, 0, 1000, 260)).copy()
    if pp:
        ImageDraw.Draw(left).rectangle(
            [pp[0] - 740, pp[1], pp[2] - 740, pp[3]], outline=(255, 80, 80), width=2
        )
        # 右縁ガイド
        ImageDraw.Draw(left).line(
            [pp[2] - 740, 0, pp[2] - 740, 259], fill=(255, 120, 120), width=1
        )
    if dd:
        ImageDraw.Draw(right).rectangle(
            [dd[0] - 740, dd[1], dd[2] - 740, dd[3]], outline=(80, 255, 80), width=2
        )
        ImageDraw.Draw(right).line(
            [dd[2] - 740, 0, dd[2] - 740, 259], fill=(120, 255, 120), width=1
        )
    # pair右縁を右パネルにも点線相当
    if pp:
        ImageDraw.Draw(right).line(
            [pp[2] - 740, 0, pp[2] - 740, 259], fill=(255, 180, 80), width=1
        )
    sheet.paste(left, (10, 40))
    sheet.paste(right, (290, 40))
    d = ImageDraw.Draw(sheet)
    d.text((10, 8), "PAIR locked (ref)", fill=(255, 180, 180))
    d.text((290, 8), "DIGITS 1+0 right-edge", fill=(180, 255, 180))
    if pp and dd:
        d.text(
            (10, 310),
            f"rightΔ={dd[2]-pp[2]}px bottomΔ={dd[3]-pp[3]}px  numberBox FROZEN",
            fill=(220, 220, 220),
        )
    ov = out_dir / "POC19_10kan_OVERLAY_pair_vs_digits.jpg"
    sheet.save(ov, quality=92)
    print("OVERLAY", ov)

    # pair無しの例: 21（pair無し想定）— 同じ numberBox 右縁合わせ
    t21 = json.loads(json.dumps(typo))
    # 21にpairが無いことを確認しつつ
    from glyph_assets import load_pair_glyph

    has21 = load_pair_glyph(21, dirs) is not None
    print("pair_21 exists?", has21)
    out21, meta21 = compose_badge_with_glyphs(
        base0.copy(),
        set_count=21,
        unit="缶",
        cx=int(gold["cx"]),
        cy=int(gold["cy"]),
        diameter=int(gold["diameter"]),
        typo=t21,
        glyph_dirs=dirs,
        canvas=1200,
    )
    out21.convert("RGB").save(out_dir / "POC19_21kan_DIGITS.jpg", quality=95)
    print("21", meta21.get("numSource"), meta21.get("numberPos"), meta21.get("numberFitMode"))


if __name__ == "__main__":
    main()
