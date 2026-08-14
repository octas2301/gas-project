# -*- coding: utf-8 -*-
"""
2桁: 用意済み pair_* を素材アスペクト維持で見本数字位置へ。
単位・セット(unitsetRect)と1桁設定は固定。10缶で検証。
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional, Tuple

from PIL import Image, ImageDraw

from glyph_assets import (
    _crop_to_alpha,
    _paste_box1000,
    compose_badge_with_glyphs,
    default_glyph_dirs,
    load_pair_glyph,
)
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"
BBox = Tuple[int, int, int, int, int, int]


def sample_num_bbox_2digit(rgb1000: Image.Image) -> BBox:
    """10袋見本: 数字領域（単位左・セット上）。"""
    pix = rgb1000.load()
    cx, cy, rad = 880, 120, 122
    xs: List[int] = []
    ys: List[int] = []
    for y in range(max(0, cy - rad), min(1000, cy + rad + 1)):
        for x in range(max(0, cx - rad), min(1000, cx + rad + 1)):
            if (x - cx) ** 2 + (y - cy) ** 2 > rad * rad:
                continue
            # 2桁: 単位は右寄り。数字は x<910 かつセット帯より上
            if x >= 910 or y >= 128:
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

    # --- Canva 縦横比ルール（正本） ---
    typo["canvaAspectLock"] = {
        "enabled": True,
        "rule": "Canva素材（digit / pair / unitset）の縦横比を変えない。stretch禁止。contain（等倍縮小）のみ。",
        "fitModeAllowed": ["contain"],
        "fitModeForbidden": ["stretch"],
    }

    locked = typo["lockedUnitSet1000"]
    # 1桁完全固定
    d1 = locked["1digit"]
    d1["freezeUnitSet"] = True
    d1["freezeNumber"] = True
    d1["numberFitMode"] = "contain"
    d1["noteNumber"] = (d1.get("noteNumber") or "") + " | FROZEN 2026-08-02"
    locked["1digit"] = d1

    d2 = locked["2digit"]
    unitset_rect = list(d2["unitsetRect1000"])  # 単位・セット固定
    align = str(d2.get("align") or "top_left")
    d2["freezeUnitSet"] = True
    d2["noteUnitSet"] = "LOCKED — do not change unitsetRect/align"
    d2["numberFitMode"] = "contain"
    print("FROZEN unitset 2digit", unitset_rect)
    print("FROZEN 1digit numberBox", d1["numberBox"])

    sample = (
        Image.open(ref / "sanky-4906283045614-oya_6.jpg")
        .convert("RGB")
        .resize((1000, 1000), Image.Resampling.LANCZOS)
    )
    sn = sample_num_bbox_2digit(sample)
    print("sample 2digit num", sn)

    pair = load_pair_glyph(10, dirs)
    assert pair is not None, "pair_10.png missing in Canva 95"
    pair = _crop_to_alpha(pair)
    pw, ph = pair.size
    aspect = pw / ph
    print("pair_10 crop", pw, ph, "aspect", round(aspect, 4))

    # 等倍: 高さ合わせ / 幅合わせ
    th = sn[5]
    tw = max(1, round(th * aspect))
    tw2 = sn[4]
    th2 = max(1, round(tw2 / aspect))
    print("uniform match_h", tw, th, "match_w", tw2, th2)

    cands = []
    for tw0, th0, tag in ((tw, th, "match_h"), (tw2, th2, "match_w")):
        for dx in range(-4, 5):
            for dy in range(-4, 5):
                cands.append(([sn[0] + dx, sn[1] + dy, tw0, th0], tag))
                cands.append(
                    (
                        [
                            sn[0] + (sn[4] - tw0) // 2 + dx,
                            sn[1] + (sn[5] - th0) // 2 + dy,
                            tw0,
                            th0,
                        ],
                        tag + "_c",
                    )
                )
    seen = set()
    uniq = []
    for box, tag in cands:
        k = tuple(box + [tag])
        if k in seen:
            continue
        seen.add(k)
        uniq.append((box, tag))
        if len(uniq) >= 20:
            break

    best = (-1.0, None, None, None, None)
    for i, (box, tag) in enumerate(uniq, 1):
        layer = Image.new("RGBA", (1200, 1200), (0, 0, 0, 0))
        _paste_box1000(layer, pair, box, canvas=1200, fit="contain")
        gn = alpha_bbox1000(layer)
        assert gn is not None
        sw, sh = sn[4], sn[5]
        gw, gh = gn[4], gn[5]
        size_h = 1 - min(1.0, abs(sh - gh) / max(1, sh))
        size_w = 1 - min(1.0, abs(sw - gw) / max(1, sw))
        size = (size_h + size_w) / 2
        scx = (sn[0] + sn[2]) / 2
        scy = (sn[1] + sn[3]) / 2
        gcx = (gn[0] + gn[2]) / 2
        gcy = (gn[1] + gn[3]) / 2
        diag = (sw**2 + sh**2) ** 0.5
        dist = ((scx - gcx) ** 2 + (scy - gcy) ** 2) ** 0.5
        pos = 1 - min(1.0, dist / max(1.0, diag * 0.5))
        combo = 0.55 * pos + 0.45 * size
        print(
            f"i{i:02d} {tag} size={size:.3f} (h={size_h:.3f} w={size_w:.3f}) "
            f"pos={pos:.3f} box={box} gen={gn}"
        )
        if combo > best[0]:
            best = (combo, box, tag, (size, pos, size_h, size_w), layer)

    print("BEST", best[0], best[2], best[3], best[1])

    d2["numberBox"] = best[1]
    d2["numberFitMode"] = "contain"
    d2["numberSource"] = "pair_10"
    d2["unitsetRect1000"] = unitset_rect
    d2["align"] = align
    d2["noteNumber"] = (
        f"pair_10 contain size={best[3][0]:.4f} pos={best[3][1]:.4f} "
        f"h={best[3][2]:.4f} w={best[3][3]:.4f}"
    )
    locked["2digit"] = d2
    locked["freezeUnitSet"] = True
    typo["lockedUnitSet1000"] = locked
    rules["rakuten"]["badgeTypography"] = typo
    RULES.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)
    gold = rules["rakuten"]["goldCircle1200"]

    # 10缶（単位=缶 U+7F36）
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
    assert meta.get("unitsetRect1000") == unitset_rect
    assert meta.get("numberFitMode") == "contain"
    path = out_dir / "POC17_10kan_PAIR.jpg"
    out.convert("RGB").save(path, quality=95)

    view = out.convert("RGB").resize((1000, 1000), Image.Resampling.LANCZOS)
    gn = alpha_bbox1000(best[4])
    sheet = Image.new("RGB", (560, 360), (28, 28, 28))
    left = sample.crop((740, 0, 1000, 260)).copy()
    right = view.crop((740, 0, 1000, 260)).copy()
    d = ImageDraw.Draw(left)
    d.rectangle([sn[0] - 740, sn[1], sn[2] - 740, sn[3]], outline=(255, 80, 80), width=2)
    d2draw = ImageDraw.Draw(right)
    if gn:
        d2draw.rectangle([gn[0] - 740, gn[1], gn[2] - 740, gn[3]], outline=(80, 255, 80), width=2)
    d2draw.rectangle(
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
    dd = ImageDraw.Draw(sheet)
    dd.text((10, 8), "10 SAMPLE num (oya_6)", fill=(255, 180, 180))
    dd.text((290, 8), "10缶 GEN pair_10 contain", fill=(180, 255, 180))
    dd.text(
        (10, 310),
        f"pos={best[3][1]:.1%} h={best[3][2]:.1%} w={best[3][3]:.1%} unitset LOCKED",
        fill=(220, 220, 220),
    )
    ov = out_dir / "POC17_10kan_OVERLAY_pair.jpg"
    sheet.save(ov, quality=92)

    print("EXPORT", path)
    print("OVERLAY", ov)
    print("meta", {k: meta.get(k) for k in ("mode", "numSource", "numberPos", "unitsetPos", "numberFitMode")})
    print("next: digit+digit compose for missing pairs")


if __name__ == "__main__":
    main()
