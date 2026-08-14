# -*- coding: utf-8 -*-
"""1桁数字だけ見本位置へ合わせる。単位・セット(unitsetRect)は一切動かさない。"""
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
    load_digit_glyph,
    load_unitset_glyph,
)
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"
BBox = Tuple[int, int, int, int, int, int]


def sample_num_bbox(rgb1000: Image.Image) -> BBox:
    """1袋見本: 数字領域（単位左・セット上）の暗いインク外接。"""
    pix = rgb1000.load()
    cx, cy, rad = 880, 120, 122
    xs: List[int] = []
    ys: List[int] = []
    for y in range(max(0, cy - rad), min(1000, cy + rad + 1)):
        for x in range(max(0, cx - rad), min(1000, cx + rad + 1)):
            if (x - cx) ** 2 + (y - cy) ** 2 > rad * rad:
                continue
            if x >= 890 or y >= 125:
                continue
            r, g, b = pix[x, y]
            if r + g + b < 280 and r > 40:
                xs.append(x)
                ys.append(y)
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    return x0, y0, x1, y1, x1 - x0 + 1, y1 - y0 + 1


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


def size_pos(sample_bb: BBox, gen_bb: Optional[BBox]) -> Tuple[float, float, float]:
    if not gen_bb:
        return 0.0, 0.0, 0.0
    sw, sh = sample_bb[4], sample_bb[5]
    gw, gh = gen_bb[4], gen_bb[5]
    size = 1.0 - min(1.0, (abs(sw - gw) / max(1, sw) + abs(sh - gh) / max(1, sh)) / 2)
    scx = (sample_bb[0] + sample_bb[2]) / 2
    scy = (sample_bb[1] + sample_bb[3]) / 2
    gcx = (gen_bb[0] + gen_bb[2]) / 2
    gcy = (gen_bb[1] + gen_bb[3]) / 2
    diag = (sw**2 + sh**2) ** 0.5
    dist = ((scx - gcx) ** 2 + (scy - gcy) ** 2) ** 0.5
    pos = 1.0 - min(1.0, dist / max(1.0, diag * 0.5))
    return size, pos, (size + pos) / 2


def paste_num(
    base: Image.Image,
    digit: Image.Image,
    box: List[int],
    mode: str,
) -> None:
    """stretch=見本枠に引き伸ばし / contain=素材比維持で枠内中央。"""
    if mode == "stretch":
        _paste_box1000(base, digit, box, canvas=1200)
        return
    scale = 1200 / 1000
    x = int(round(box[0] * scale))
    y = int(round(box[1] * scale))
    w = max(1, int(round(box[2] * scale)))
    h = max(1, int(round(box[3] * scale)))
    g = _crop_to_alpha(digit.convert("RGBA"))
    gw, gh = g.size
    s = min(w / gw, h / gh)
    nw, nh = max(1, int(round(gw * s))), max(1, int(round(gh * s)))
    g = g.resize((nw, nh), Image.Resampling.LANCZOS)
    ox = x + (w - nw) // 2
    oy = y + (h - nh) // 2
    base.alpha_composite(g, (ox, oy))


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
    locked = typo["lockedUnitSet1000"]
    cfg1 = locked["1digit"]
    # 単位・セット固定（変更禁止）
    unitset_rect = list(cfg1["unitsetRect1000"])
    align = str(cfg1.get("align") or "top_left")
    locked["freezeUnitSet"] = True
    cfg1["freezeUnitSet"] = True
    cfg1["noteUnitSet"] = "LOCKED — do not change unitsetRect/align"

    sample = (
        Image.open(ref / "sanky-4906283045614-oya.jpg")
        .convert("RGB")
        .resize((1000, 1000), Image.Resampling.LANCZOS)
    )
    sn = sample_num_bbox(sample)
    print(f"sample number ink={sn} box=[{sn[0]},{sn[1]},{sn[4]},{sn[5]}]")

    unitset = load_unitset_glyph("袋", dirs)
    digit = _crop_to_alpha(load_digit_glyph("1", dirs))
    print("digit native", digit.size)

    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)

    target = [sn[0], sn[1], sn[4], sn[5]]
    # 候補: 見本枠±微調整、stretch/contain（単位セットは常に同じ）
    cands = []
    for mode in ("stretch", "contain"):
        for dw in (-4, -2, 0, 2, 4):
            for dh in (-4, -2, 0, 2, 4):
                for dx in (-3, -1, 0, 1, 3):
                    for dy in (-3, -1, 0, 1, 3):
                        box = [
                            target[0] + dx,
                            target[1] + dy,
                            max(16, target[2] + dw),
                            max(40, target[3] + dh),
                        ]
                        cands.append((box, mode))
    seen = set()
    uniq = []
    prefer = [(target, "stretch"), (target, "contain")] + cands
    for box, mode in prefer:
        k = tuple(box + [mode])
        if k in seen:
            continue
        seen.add(k)
        uniq.append((box, mode))
        if len(uniq) >= 20:
            break

    best = (-1.0, None, None, None, None)
    for i, (box, mode) in enumerate(uniq, 1):
        # 数字レイヤのみ採点（単位セットは載せない）
        layer = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
        paste_num(layer, digit, box, mode)
        gn = alpha_bbox1000(layer)
        size_m, pos_m, combo = size_pos(sn, gn)
        print(f"i{i:02d} mode={mode} size={size_m:.4f} pos={pos_m:.4f} box={box} gen={gn}")
        if combo > best[0]:
            best = (combo, box, mode, (size_m, pos_m), layer)
        if size_m >= 0.98 and pos_m >= 0.98:
            print("*** REACHED 98% number size+pos ***")

    print(
        f"BEST combo={best[0]:.4f} size={best[3][0]:.4f} pos={best[3][1]:.4f} "
        f"mode={best[2]} box={best[1]}"
    )

    # numberFitMode を layout に保存（compose が stretch 固定なら contain 時は box を実インクに合わせる）
    num_box = list(best[1])
    num_mode = best[2]
    if num_mode == "contain":
        # compose は stretch 貼りなので、実効インク枠に合わせて numberBox を更新
        gn = alpha_bbox1000(best[4])
        assert gn is not None
        num_box = [gn[0], gn[1], gn[4], gn[5]]
        num_mode = "stretch"  # 実効枠へ stretch=ほぼ等倍
        print(f"contain→effective numberBox={num_box}")

    cfg1["numberBox"] = num_box
    cfg1["numberFitMode"] = "stretch"
    cfg1["noteNumber"] = f"size={best[3][0]:.4f} pos={best[3][1]:.4f} (unitset frozen)"
    cfg1["unitsetRect1000"] = unitset_rect  # 明示再固定
    cfg1["align"] = align
    locked["1digit"] = cfg1
    typo["lockedUnitSet1000"] = locked
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    # 最終合成（単位セット固定＋数字）
    gold = rules["rakuten"]["goldCircle1200"]
    out, meta = compose_badge_with_glyphs(
        base0.copy(),
        set_count=1,
        unit="袋",
        cx=int(gold["cx"]),
        cy=int(gold["cy"]),
        diameter=int(gold["diameter"]),
        typo=typo,
        glyph_dirs=dirs,
        canvas=1200,
    )
    path = out_dir / "POC15_1fukuro_NUM.jpg"
    out.convert("RGB").save(path, quality=95)

    # 重ね: 左見本 / 右生成 / 数字枠
    view = out.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
    # 数字のみレイヤを再生成して枠確認
    num_only = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
    paste_num(num_only, digit, num_box, "stretch")
    gn = alpha_bbox1000(num_only)
    sheet = Image.new("RGB", (560, 360), (28, 28, 28))
    left = sample.crop((740, 0, 1000, 260)).copy()
    right = view.crop((740, 0, 1000, 260)).copy()
    d1 = ImageDraw.Draw(left)
    d1.rectangle([sn[0] - 740, sn[1], sn[2] - 740, sn[3]], outline=(255, 80, 80), width=2)
    d2 = ImageDraw.Draw(right)
    if gn:
        d2.rectangle([gn[0] - 740, gn[1], gn[2] - 740, gn[3]], outline=(80, 255, 80), width=2)
    # unitset rect (cyan) to show frozen
    ur = unitset_rect
    d2.rectangle(
        [ur[0] - 740, ur[1], ur[0] + ur[2] - 740, ur[1] + ur[3]],
        outline=(80, 200, 255),
        width=1,
    )
    sheet.paste(left, (10, 40))
    sheet.paste(right, (290, 40))
    d = ImageDraw.Draw(sheet)
    d.text((10, 8), "1 SAMPLE num", fill=(255, 180, 180))
    d.text((290, 8), "1 GEN (unitset LOCKED)", fill=(180, 255, 180))
    size_m, pos_m = best[3]
    d.text(
        (10, 310),
        f"NUM size={size_m:.1%} pos={pos_m:.1%}  unitset frozen {unitset_rect}",
        fill=(220, 220, 220),
    )
    ov = out_dir / "POC15_1_OVERLAY_num.jpg"
    sheet.save(ov, quality=92)

    ok = size_m >= 0.98 and pos_m >= 0.98
    print(f"EXPORT {path}")
    print(f"OVERLAY {ov}")
    print(f"unitset LOCKED {unitset_rect}")
    print(f"numberBox {num_box}")
    print(f"meta numberPos={meta.get('numberPos')} unitsetPos={meta.get('unitsetPos')}")
    print(f"98% number size+pos: {'YES' if ok else 'NO'} ({size_m:.1%} / {pos_m:.1%})")


if __name__ == "__main__":
    main()
