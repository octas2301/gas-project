# -*- coding: utf-8 -*-
"""単位・セットの大きさ+位置を見本に合わせる（数字は採点外）。最大20試行。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional, Tuple

from PIL import Image, ImageDraw

from glyph_assets import (
    _crop_to_alpha,
    default_glyph_dirs,
    load_digit_glyph,
    load_pair_glyph,
    load_unitset_glyph,
    paste_glyph,
    split_unitset_parts,
)
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"
OUT_QA = Path(__file__).resolve().parent / "tmp_unitset_lock"
OUT_QA.mkdir(exist_ok=True)

BBox = Tuple[int, int, int, int, int, int]  # x0,y0,x1,y1,w,h


def paste_box(
    base: Image.Image,
    glyph: Image.Image,
    box1000: List[int],
    canvas: int = 1200,
    grid: float = 1000.0,
    mode: str = "stretch",
) -> Tuple[int, int, int, int]:
    scale = canvas / grid
    x = int(round(box1000[0] * scale))
    y = int(round(box1000[1] * scale))
    w = max(1, int(round(box1000[2] * scale)))
    h = max(1, int(round(box1000[3] * scale)))
    g = _crop_to_alpha(glyph.convert("RGBA"))
    if mode == "stretch":
        g = g.resize((w, h), Image.Resampling.LANCZOS)
        paste_glyph(base, g, (x, y))
        return x, y, w, h
    gw, gh = g.size
    s = min(w / gw, h / gh)
    nw = max(1, int(round(gw * s)))
    nh = max(1, int(round(gh * s)))
    g = g.resize((nw, nh), Image.Resampling.LANCZOS)
    ox = x + (w - nw) // 2
    oy = y + (h - nh) // 2
    paste_glyph(base, g, (ox, oy))
    return ox, oy, nw, nh


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


def dark_bbox_in_region(rgb1000: Image.Image, region: List[int], thr_sum: int = 280) -> Optional[BBox]:
    x, y, w, h = [int(v) for v in region]
    pad = 12
    x0, y0 = max(0, x - pad), max(0, y - pad)
    x1, y1 = min(1000, x + w + pad), min(1000, y + h + pad)
    pix = rgb1000.load()
    xs: List[int] = []
    ys: List[int] = []
    for yy in range(y0, y1):
        for xx in range(x0, x1):
            r, g, b = pix[xx, yy]
            if r + g + b < thr_sum and r > 40:
                xs.append(xx)
                ys.append(yy)
    if not xs:
        return None
    xa, xb = min(xs), max(xs)
    ya, yb = min(ys), max(ys)
    return xa, ya, xb, yb, xb - xa + 1, yb - ya + 1


def size_pos(sample_bb: Optional[BBox], gen_bb: Optional[BBox]) -> Tuple[float, float, float]:
    if not sample_bb or not gen_bb:
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


def make_overlay(
    sample: Image.Image,
    gen_rgb: Image.Image,
    path: Path,
    title: str,
    su: Optional[BBox],
    ss: Optional[BBox],
    gu: Optional[BBox],
    gs: Optional[BBox],
    size_m: float,
    pos_m: float,
) -> None:
    sheet = Image.new("RGB", (560, 360), (28, 28, 28))

    def draw_bb(im: Image.Image, bb: Optional[BBox], color: tuple) -> None:
        if not bb:
            return
        d = ImageDraw.Draw(im)
        d.rectangle([bb[0] - 740, bb[1], bb[2] - 740, bb[3]], outline=color, width=2)

    left = sample.crop((740, 0, 1000, 260)).copy()
    right = gen_rgb.crop((740, 0, 1000, 260)).copy()
    draw_bb(left, su, (255, 80, 80))
    draw_bb(left, ss, (255, 160, 80))
    draw_bb(right, gu, (80, 255, 80))
    draw_bb(right, gs, (80, 255, 160))
    sheet.paste(left, (10, 40))
    sheet.paste(right, (290, 40))
    d = ImageDraw.Draw(sheet)
    d.text((10, 8), f"{title} SAMPLE", fill=(255, 180, 180))
    d.text((290, 8), f"{title} GEN", fill=(180, 255, 180))
    d.text((10, 310), f"size={size_m:.1%}  pos={pos_m:.1%}  (number excluded)", fill=(220, 220, 220))
    sheet.save(path, quality=92)


def main() -> None:
    root = default_work_root()
    ref = next(p for p in root.iterdir() if p.is_dir() and p.name.startswith("04"))
    base_path = next(
        (next(p for p in root.iterdir() if p.is_dir() and p.name.startswith("02"))).glob("*.jpg")
    )
    out_dir = root / "00.テスト出力"
    dirs = default_glyph_dirs(root)
    unitset = load_unitset_glyph("袋", dirs)
    assert unitset is not None
    unit_im, set_im = split_unitset_parts(unitset)
    unit_im = _crop_to_alpha(unit_im)
    set_im = _crop_to_alpha(set_im)
    unit_im.save(OUT_QA / "split_unit.png")
    set_im.save(OUT_QA / "split_set.png")

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
    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)

    configs = [
        ("1digit", 1, sample_1, [900, 58, 72, 87], [834, 140, 136, 28], [845, 26, 54, 119]),
        ("2digit", 10, sample_10, [923, 64, 62, 81], [848, 140, 135, 34], [806, 43, 113, 102]),
    ]

    results = {}
    for key, n, sample, ub0, sb0, nb in configs:
        su0 = dark_bbox_in_region(sample, ub0)
        ss0 = dark_bbox_in_region(sample, sb0)
        print(f"=== {key} sample ink unit={su0} set={ss0} ===")
        target_u = [su0[0], su0[1], su0[4], su0[5]] if su0 else ub0
        target_s = [ss0[0], ss0[1], ss0[4], ss0[5]] if ss0 else sb0

        best = (-1.0, None, None, None, None, None)
        cands = []
        for mode in ("stretch", "contain"):
            for dw in (-2, -1, 0, 1, 2):
                for dh in (-2, -1, 0, 1, 2):
                    for dx in (-1, 0, 1):
                        for dy in (-1, 0, 1):
                            ub = [
                                target_u[0] + dx,
                                target_u[1] + dy,
                                max(20, target_u[2] + dw),
                                max(20, target_u[3] + dh),
                            ]
                            sb = [
                                target_s[0] + dx,
                                target_s[1] + dy,
                                max(40, target_s[2] + dw),
                                max(12, target_s[3] + dh),
                            ]
                            cands.append((ub, sb, mode))
        seen = set()
        uniq = []
        for c in [(target_u, target_s, "stretch")] + cands:
            k = tuple(c[0] + c[1] + [c[2]])
            if k in seen:
                continue
            seen.add(k)
            uniq.append(c)
            if len(uniq) >= 20:
                break

        reached = False
        for i, (ub, sb, mode) in enumerate(uniq, 1):
            lu = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
            ls = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
            paste_box(lu, unit_im, ub, mode=mode)
            paste_box(ls, set_im, sb, mode=mode)
            gu = alpha_bbox1000(lu)
            gs = alpha_bbox1000(ls)
            sz_u, pos_u, _ = size_pos(su0, gu)
            sz_s, pos_s, _ = size_pos(ss0, gs)
            size_m = (sz_u + sz_s) / 2
            pos_m = (pos_u + pos_s) / 2
            combo = (size_m + pos_m) / 2
            print(
                f"i{i:02d} mode={mode} size={size_m:.4f} pos={pos_m:.4f} "
                f"u({sz_u:.3f}/{pos_u:.3f}) s({sz_s:.3f}/{pos_s:.3f}) ub={ub} sb={sb}"
            )
            layer = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
            paste_box(layer, unit_im, ub, mode=mode)
            paste_box(layer, set_im, sb, mode=mode)
            if combo > best[0]:
                best = (combo, ub, sb, mode, (size_m, pos_m, sz_u, pos_u, sz_s, pos_s), layer)
            if size_m >= 0.98 and pos_m >= 0.98 and not reached:
                print("*** REACHED 98% size+pos ***")
                reached = True
        results[key] = best
        print(
            f"BEST {key} combo={best[0]:.4f} size={best[4][0]:.4f} pos={best[4][1]:.4f} "
            f"mode={best[3]} ub={best[1]} sb={best[2]}"
        )

    rules = json.loads(RULES.read_text(encoding="utf-8"))
    typo = rules["rakuten"]["badgeTypography"]
    locked = {"enabled": True}
    for key, n, sample, ub0, sb0, nb in configs:
        b = results[key]
        locked[key] = {
            "unitBox": b[1],
            "setBox": b[2],
            "numberBox": nb,
            "fitMode": b[3],
            "note": f"size={b[4][0]:.4f} pos={b[4][1]:.4f} (unit/set only; number excluded)",
        }
    typo["lockedUnitSet1000"] = locked
    if "canvaUnitset" in typo:
        typo["canvaUnitset"]["enabled"] = False
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    for key, n, sample, ub0, sb0, nb in configs:
        b = results[key]
        work = base0.copy()
        paste_box(work, unit_im, b[1], mode=b[3])
        paste_box(work, set_im, b[2], mode=b[3])
        if n < 10:
            num = _crop_to_alpha(load_digit_glyph(str(n), dirs))
        else:
            pair = load_pair_glyph(n, dirs)
            num = _crop_to_alpha(pair) if pair is not None else None
        if num is not None:
            paste_box(work, num, nb, mode="stretch")
        name = "POC13_1fukuro_LOCKED.jpg" if n == 1 else "POC13_10fukuro_LOCKED.jpg"
        path = out_dir / name
        work.convert("RGB").save(path, quality=95)

        view = base0.copy()
        view.alpha_composite(b[5])
        view_rgb = view.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
        lu = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
        ls = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
        paste_box(lu, unit_im, b[1], mode=b[3])
        paste_box(ls, set_im, b[2], mode=b[3])
        gu = alpha_bbox1000(lu)
        gs = alpha_bbox1000(ls)
        su0 = dark_bbox_in_region(sample, b[1])
        ss0 = dark_bbox_in_region(sample, b[2])
        ov = out_dir / ("POC13_1_OVERLAY_unitset.jpg" if n == 1 else "POC13_10_OVERLAY_unitset.jpg")
        size_m, pos_m = b[4][0], b[4][1]
        make_overlay(sample, view_rgb, ov, f"{n}袋", su0, ss0, gu, gs, size_m, pos_m)
        ok = size_m >= 0.98 and pos_m >= 0.98
        print(f"EXPORT {path.name} size={size_m:.1%} pos={pos_m:.1%} 98%={'YES' if ok else 'NO'}")
        print("OVERLAY", ov)


if __name__ == "__main__":
    main()
