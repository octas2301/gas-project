# -*- coding: utf-8 -*-
"""素材 unitset を一体・等倍縮尺のまま見本の単位+セット外接に合わせる（数字は採点外）。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional, Tuple

from PIL import Image, ImageDraw

from glyph_assets import (
    _crop_to_alpha,
    _paste_unitset_uniform1000,
    compose_badge_with_glyphs,
    default_glyph_dirs,
    load_unitset_glyph,
)
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"
OUT_QA = Path(__file__).resolve().parent / "tmp_unitset_uniform"
OUT_QA.mkdir(exist_ok=True)

BBox = Tuple[int, int, int, int, int, int]


def dark_bbox_union(rgb1000: Image.Image, regions: List[List[int]], thr_sum: int = 280) -> Optional[BBox]:
    pix = rgb1000.load()
    xs: List[int] = []
    ys: List[int] = []
    for region in regions:
        x, y, w, h = [int(v) for v in region]
        pad = 10
        x0, y0 = max(0, x - pad), max(0, y - pad)
        x1, y1 = min(1000, x + w + pad), min(1000, y + h + pad)
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
    gu: Optional[BBox],
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
    draw_bb(right, gu, (80, 255, 80))
    sheet.paste(left, (10, 40))
    sheet.paste(right, (290, 40))
    d = ImageDraw.Draw(sheet)
    d.text((10, 8), f"{title} SAMPLE unit+set", fill=(255, 180, 180))
    d.text((290, 8), f"{title} GEN uniform", fill=(180, 255, 180))
    d.text((10, 310), f"size={size_m:.1%}  pos={pos_m:.1%}  (material scale kept)", fill=(220, 220, 220))
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
    _crop_to_alpha(unitset).save(OUT_QA / "unitset_crop.png")

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

    # 見本の単位・セット推定枠（前回実測）→ 結合外接
    configs = [
        (
            "1digit",
            1,
            sample_1,
            [[888, 58, 83, 99], [834, 128, 136, 39]],
            [845, 26, 54, 119],
        ),
        (
            "2digit",
            10,
            sample_10,
            [[911, 55, 73, 102], [848, 131, 135, 42]],
            [806, 43, 113, 102],
        ),
    ]

    results = {}
    for key, n, sample, regions, nb in configs:
        su = dark_bbox_union(sample, regions)
        print(f"=== {key} sample unit+set ink={su} ===")
        assert su is not None
        target = [su[0], su[1], su[4], su[5]]

        best = (-1.0, None, None, None, None)
        # 20候補: 等倍のみ。枠の微調整＋align
        cands = []
        for align in ("top_left", "top_right", "center"):
            for dw in (-4, -2, 0, 2, 4, 6):
                for dh in (-4, -2, 0, 2, 4, 6):
                    for dx in (-2, 0, 2):
                        for dy in (-2, 0, 2):
                            rect = [
                                target[0] + dx,
                                target[1] + dy,
                                max(40, target[2] + dw),
                                max(40, target[3] + dh),
                            ]
                            cands.append((rect, align))
        seen = set()
        uniq = []
        for c in [(target, "top_left")] + cands:
            k = tuple(c[0] + [c[1]])
            if k in seen:
                continue
            seen.add(k)
            uniq.append(c)
            if len(uniq) >= 20:
                break

        reached = False
        for i, (rect, align) in enumerate(uniq, 1):
            layer = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
            pos = _paste_unitset_uniform1000(layer, unitset, rect, canvas=1200, align=align)
            gu = alpha_bbox1000(layer)
            size_m, pos_m, combo = size_pos(su, gu)
            print(
                f"i{i:02d} align={align} size={size_m:.4f} pos={pos_m:.4f} "
                f"rect={rect} placed={pos} gen={gu}"
            )
            if combo > best[0]:
                best = (combo, rect, align, (size_m, pos_m), layer)
            if size_m >= 0.98 and pos_m >= 0.98 and not reached:
                print("*** REACHED 98% size+pos (uniform material) ***")
                reached = True
        results[key] = best
        print(
            f"BEST {key} combo={best[0]:.4f} size={best[3][0]:.4f} pos={best[3][1]:.4f} "
            f"align={best[2]} rect={best[1]}"
        )

    rules = json.loads(RULES.read_text(encoding="utf-8"))
    typo = rules["rakuten"]["badgeTypography"]
    locked = {"enabled": True, "note": "unitset一体・等倍縮尺。袋/セットを別ストレッチしない。"}
    for key, n, sample, regions, nb in configs:
        b = results[key]
        locked[key] = {
            "unitsetRect1000": b[1],
            "align": b[2],
            "numberBox": nb,
            "note": f"uniform size={b[3][0]:.4f} pos={b[3][1]:.4f}",
        }
    typo["lockedUnitSet1000"] = locked
    if "canvaUnitset" in typo:
        typo["canvaUnitset"]["enabled"] = False
        typo["canvaUnitset"]["allowStretch"] = False
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    gold = rules["rakuten"]["goldCircle1200"]
    for key, n, sample, regions, nb in configs:
        b = results[key]
        typo_run = json.loads(json.dumps(typo))
        out, meta = compose_badge_with_glyphs(
            base0.copy(),
            set_count=n,
            unit="袋",
            cx=int(gold["cx"]),
            cy=int(gold["cy"]),
            diameter=int(gold["diameter"]),
            typo=typo_run,
            glyph_dirs=dirs,
            canvas=1200,
        )
        name = "POC14_1fukuro_UNIFORM.jpg" if n == 1 else "POC14_10fukuro_UNIFORM.jpg"
        path = out_dir / name
        out.convert("RGB").save(path, quality=95)

        view = base0.copy()
        view.alpha_composite(b[4])
        view_rgb = view.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
        gu = alpha_bbox1000(b[4])
        su = dark_bbox_union(sample, regions)
        ov = out_dir / ("POC14_1_OVERLAY_uniform.jpg" if n == 1 else "POC14_10_OVERLAY_uniform.jpg")
        size_m, pos_m = b[3]
        make_overlay(sample, view_rgb, ov, f"{n}袋", su, gu, size_m, pos_m)
        ok = size_m >= 0.98 and pos_m >= 0.98
        print(f"EXPORT {path.name} mode={meta.get('mode')} size={size_m:.1%} pos={pos_m:.1%} 98%={'YES' if ok else 'NO'}")
        print("OVERLAY", ov)
        print("unitsetPos", meta.get("unitsetPos"))


if __name__ == "__main__":
    main()
