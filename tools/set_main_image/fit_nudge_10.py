# -*- coding: utf-8 -*-
"""10を缶に寄せる（くっつない）＋下端を缶下端に合わせる。unitset固定・縦横比維持。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional, Tuple

from PIL import Image, ImageDraw

from glyph_assets import (
    _crop_to_alpha,
    _paste_box1000,
    _paste_unitset_uniform1000,
    compose_badge_with_glyphs,
    default_glyph_dirs,
    load_pair_glyph,
    load_unitset_glyph,
    split_unitset_parts,
)
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"
BBox = Tuple[int, int, int, int, int, int]


def alpha_bbox1000(layer1200: Image.Image, thr: int = 40) -> Optional[BBox]:
    a = layer1200.convert("RGBA").resize((1000, 1000), Image.Resampling.NEAREST)
    pix = a.load()
    xs: List[int] = []
    ys: List[int] = []
    for y in range(1000):
        for x in range(1000):
            if pix[x, y][3] > thr:
                xs.append(x)
                ys.append(y)
    if not xs:
        return None
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    return x0, y0, x1, y1, x1 - x0 + 1, y1 - y0 + 1


def main() -> None:
    root = default_work_root()
    ref = next(p for p in root.iterdir() if p.is_dir() and p.name.startswith("04"))
    base_path = next(
        (next(p for p in root.iterdir() if p.is_dir() and p.name.startswith("02"))).glob("*.jpg")
    )
    out_dir = root / "00.テスト出力"
    dirs = default_glyph_dirs(root)

    rules = json.loads(RULES.read_text(encoding="utf-8"))
    typo = rules["rakuten"]["badgeTypography"]
    d2 = typo["lockedUnitSet1000"]["2digit"]
    unitset_rect = list(d2["unitsetRect1000"])
    align = str(d2.get("align") or "top_left")
    num_box0 = list(d2["numberBox"])
    print("LOCKED unitset", unitset_rect)
    print("current numberBox", num_box0)

    unitset = load_unitset_glyph("缶", dirs)
    assert unitset is not None
    pair = _crop_to_alpha(load_pair_glyph(10, dirs))

    # 缶だけ（unitsetから単位パート）の外接を測る
    unit_im, _set_im = split_unitset_parts(unitset)
    # unitset一体配置後の「単位」位置を推定: 全体を貼ってから右上領域のインク
    us_layer = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
    _paste_unitset_uniform1000(us_layer, unitset, unitset_rect, canvas=1200, align=align)
    # 単位のみ: unitset内の単位を同じfitで、unitset配置に合わせて貼るのは複雑なので
    # 全体インクの右上（セット帯より上）を缶とする
    full_bb = alpha_bbox1000(us_layer)
    assert full_bb is not None
    # セットは下側。単位は y < set_top 付近。setは crop 下55%
    # 簡易: y < full_bb[1] + int(full_bb[5] * 0.55) かつ x > full中心
    a = us_layer.resize((1000, 1000), Image.Resampling.NEAREST)
    pix = a.load()
    xs: List[int] = []
    ys: List[int] = []
    y_cut = full_bb[1] + int(full_bb[5] * 0.52)
    x_cut = full_bb[0] + int(full_bb[4] * 0.40)
    for y in range(full_bb[1], y_cut + 1):
        for x in range(x_cut, full_bb[2] + 1):
            if pix[x, y][3] > 40:
                xs.append(x)
                ys.append(y)
    unit_bb = (min(xs), min(ys), max(xs), max(ys), max(xs) - min(xs) + 1, max(ys) - min(ys) + 1)
    print("unit(缶) bb", unit_bb)

    # 現在の数字外接
    num_layer = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
    _paste_box1000(num_layer, pair, num_box0, canvas=1200, fit="contain")
    num_bb0 = alpha_bbox1000(num_layer)
    print("num bb now", num_bb0)

    # 目標: 数字下端 = 缶下端、数字右端 = 缶左端 - gap
    gap = 3  # くっつかない隙間（1000座標）
    tw, th = num_box0[2], num_box0[3]  # サイズ維持（等倍）

    # contain貼り後の実インクサイズはほぼ tw x th（アスペクト一致時）
    # numberBox の左上を決める: 実インクが box 内中央に来る場合あり
    # → 繰り返し: boxを動かして実インクの bottom/right を合わせる
    target_bottom = unit_bb[3]  # y1
    target_right = unit_bb[0] - gap  # x0 - gap

    best = None
    # まず概算: 現在の num の w/h と box のオフセット
    assert num_bb0 is not None
    # contain中央寄せのオフセット
    off_x = num_bb0[0] - num_box0[0]
    off_y = num_bb0[1] - num_box0[1]
    # 目標インク左上
    want_x1 = target_right  # ink right
    want_y1 = target_bottom
    want_x0 = want_x1 - num_bb0[4] + 1
    want_y0 = want_y1 - num_bb0[5] + 1
    # box 左上 ≈ ink左上 - offset
    guess = [want_x0 - off_x, want_y0 - off_y, tw, th]
    print("guess numberBox", guess, "want ink", want_x0, want_y0, num_bb0[4], num_bb0[5])

    # 微調整探索（位置のみ、サイズ固定）
    for dx in range(-6, 7):
        for dy in range(-6, 7):
            box = [guess[0] + dx, guess[1] + dy, tw, th]
            layer = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
            _paste_box1000(layer, pair, box, canvas=1200, fit="contain")
            nb = alpha_bbox1000(layer)
            if not nb:
                continue
            # 下端差・右端差（缶との）
            err_b = abs(nb[3] - target_bottom)
            err_r = abs(nb[2] - target_right)
            # 重なり禁止: nb右端 < 缶左端
            overlap = max(0, nb[2] - (unit_bb[0] - 1))
            score = err_b + err_r + overlap * 5
            gap_actual = unit_bb[0] - nb[2] - 1
            if best is None or score < best[0]:
                best = (score, box, nb, err_b, err_r, gap_actual, layer)

    assert best is not None
    print(
        f"BEST score={best[0]} box={best[1]} num={best[2]} "
        f"err_b={best[3]} err_r={best[4]} gap={best[5]}"
    )

    d2["numberBox"] = best[1]
    d2["numberFitMode"] = "contain"
    d2["numberSource"] = "pair_10"
    d2["unitsetRect1000"] = unitset_rect
    d2["noteNumber"] = (
        f"nudged to 缶: bottom_align err={best[3]}px right_gap={best[5]}px box={best[1]}"
    )
    typo["lockedUnitSet1000"]["2digit"] = d2
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)
    gold = rules["rakuten"]["goldCircle1200"]
    out, meta = compose_badge_with_glyphs(
        base0.copy(),
        set_count=10,
        unit="缶",
        cx=int(gold["cx"]),
        cy=int(gold["cy"]),
        diameter=int(gold["diameter"]),
        typo=typo,
        glyph_dirs=dirs,
        canvas=1200,
    )
    path = out_dir / "POC18_10kan_NUDGE.jpg"
    out.convert("RGB").save(path, quality=95)

    sample = (
        Image.open(ref / "sanky-4906283045614-oya_6.jpg")
        .convert("RGB")
        .resize((1000, 1000), Image.Resampling.LANCZOS)
    )
    view = out.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
    nb = best[2]
    sheet = Image.new("RGB", (560, 360), (28, 28, 28))
    left = sample.crop((740, 0, 1000, 260)).copy()
    right = view.crop((740, 0, 1000, 260)).copy()
    ImageDraw.Draw(left).rectangle([806 - 740, 43, 909 - 740, 127], outline=(255, 80, 80), width=2)
    dr = ImageDraw.Draw(right)
    dr.rectangle([nb[0] - 740, nb[1], nb[2] - 740, nb[3]], outline=(80, 255, 80), width=2)
    dr.rectangle(
        [unit_bb[0] - 740, unit_bb[1], unit_bb[2] - 740, unit_bb[3]],
        outline=(255, 200, 80),
        width=2,
    )
    dr.rectangle(
        [
            unitset_rect[0] - 740,
            unitset_rect[1],
            unitset_rect[0] + unitset_rect[2] - 740,
            unitset_rect[1] + unitset_rect[3],
        ],
        outline=(80, 200, 255),
        width=1,
    )
    sheet.paste(left, (10, 40))
    sheet.paste(right, (290, 40))
    d = ImageDraw.Draw(sheet)
    d.text((10, 8), "SAMPLE 10袋", fill=(255, 180, 180))
    d.text((290, 8), "10缶 nudged", fill=(180, 255, 180))
    d.text(
        (10, 310),
        f"bottom_err={best[3]}px gap_to_缶={best[5]}px unitset LOCKED",
        fill=(220, 220, 220),
    )
    ov = out_dir / "POC18_10kan_OVERLAY_nudge.jpg"
    sheet.save(ov, quality=92)
    print("EXPORT", path)
    print("OVERLAY", ov)
    print("numberPos", meta.get("numberPos"), "unitsetPos", meta.get("unitsetPos"))


if __name__ == "__main__":
    main()
