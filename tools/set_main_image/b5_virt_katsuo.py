# -*- coding: utf-8 -*-
"""
B-⑤ 仮想テスト（案A）: たべるおだし・かつお（JAN 4906283045119）

- 98.楽天金丸素材 の実ファイルを SampleOverflow 座標で後貼り
- テーマ対ドラフト（材料があるもの最大4）。アピールは Pillow
- AI画像は使わず、正本＋構図差で仮想（02.楽天ベースには書かない）

出力: 00.テスト出力/b5_main_base_virt/{JAN}_{timestamp}/
"""
from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, ImageDraw, ImageFilter

from rakuten_badge import _load_font
from rakuten_layer import compose_rakuten_layered, load_layout_rules
from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT, resolve_kinmaru

LOG = logging.getLogger("set_main_image.b5_virt_katsuo")

JAN = "4906283045119"
TRUTH_NAME = "たべるおだしかつお.png"
UNIT = "袋"
NS = (1, 3, 7)
CANVAS = 1200
BASE_COLOR = (250, 247, 240)  # warm_white 寄り

# 仮想コピー（材料想定。T1は賞・累計が無いので出さない＝§3.7）
PROPOSALS: List[Dict[str, Any]] = [
    {
        "themeId": "N1",
        "themeName": "商品名明示",
        "mode": "name",
        "upper": "かつお節のおだし",
        "lower": "たべるおだし",
    },
    {
        "themeId": "T2",
        "themeName": "産地・素材",
        "mode": "talk",
        "upper": "かつお節をていねいに",
        "lower": "素材を味わう",
    },
    {
        "themeId": "T3",
        "themeName": "使い方・シーン",
        "mode": "talk",
        "upper": "お茶漬け・おつまみに",
        "lower": "さっと一振り",
    },
    {
        "themeId": "T4",
        "themeName": "ベネフィット",
        "mode": "talk",
        "upper": "口にとろける食感",
        "lower": "だしの旨み",
    },
]


def _truth_path(work: Path) -> Path:
    p = work / "01.amazon白抜きベース" / TRUTH_NAME
    if not p.is_file():
        raise SystemExit(f"正本なし: {p}")
    return p


def _sample_path(work: Path) -> Path:
    p = work / "04.楽天見本" / "sanky-4906283045614-oya.jpg"
    if not p.is_file():
        raise SystemExit(f"見本なし: {p}")
    return p


def _gold_params() -> Tuple[int, int, int]:
    rakuten = load_layout_rules().get("rakuten") or {}
    g = rakuten.get("goldCircle1200SampleOverflow") or rakuten.get("goldCircle1200") or {}
    return int(g.get("cx", 1102)), int(g.get("cy", 103)), int(g.get("diameter", 343))


def _appeal_box() -> Dict[str, Any]:
    rakuten = load_layout_rules().get("rakuten") or {}
    return rakuten.get("appealHeadline1200") or {
        "box": {"x": 40, "y": 24, "w": 840, "h": 200},
        "rightLimitX": 880,
        "upper": {"maxZenkaku": 18},
        "lower": {"maxZenkaku": 10},
    }


def _badge_bbox_1200() -> Tuple[int, int, int, int]:
    rakuten = load_layout_rules().get("rakuten") or {}
    g = rakuten.get("goldCircle1200SampleOverflow") or {}
    if g.get("visibleLeftX") is not None:
        return (
            int(g["visibleLeftX"]),
            0,
            1199,
            int(g.get("visibleBottomY") or 274),
        )
    return (931, 0, 1199, 274)


def _zenkaku_len(s: str) -> int:
    """全角=1・半角=0.5・切り上げ。"""
    import math

    n = 0.0
    for ch in s:
        n += 0.5 if ord(ch) < 128 else 1.0
    return int(math.ceil(n))


def _load_gold_disk(work: Path, diameter: int) -> Image.Image:
    """98素材（中央の円）を切り出し、直径にリサイズ。円の外は透過のまま。"""
    path = resolve_kinmaru(work)
    if path is None or not path.is_file():
        raise SystemExit(f"98.楽天金丸素材 にPNGなし: {work / '98.楽天金丸素材'}")
    im = Image.open(path).convert("RGBA")
    alpha = im.split()[-1]
    w, h = im.size
    mid_y = h // 2
    mid_x = w // 2
    xs = [x for x in range(w) if alpha.getpixel((x, mid_y)) > 20]
    ys = [y for y in range(h) if alpha.getpixel((mid_x, y)) > 20]
    if not xs or not ys:
        raise SystemExit(f"金丸の不透明領域なし: {path}")
    crop = im.crop((xs[0], ys[0], xs[-1] + 1, ys[-1] + 1))
    d = max(8, int(diameter))
    disk = crop.resize((d, d), Image.Resampling.LANCZOS)
    return _punch_circle_alpha(disk)


def _punch_circle_alpha(im: Image.Image) -> Image.Image:
    """正方形の四隅を必ず透過にする（円の外にベージュを残さない）。"""
    out = im.convert("RGBA")
    w, h = out.size
    cx, cy = (w - 1) / 2.0, (h - 1) / 2.0
    r2 = (min(w, h) / 2.0) ** 2
    px = out.load()
    for y in range(h):
        dy = y - cy
        for x in range(w):
            dx = x - cx
            if dx * dx + dy * dy > r2:
                r, g, b, _a = px[x, y]
                px[x, y] = (r, g, b, 0)
    return out


def _stamp_gold(base: Image.Image, gold: Image.Image) -> Image.Image:
    cx, cy, diameter = _gold_params()
    out = base.convert("RGBA")
    gx = int(round(cx - diameter / 2.0))
    gy = int(round(cy - diameter / 2.0))
    out.alpha_composite(gold, (gx, gy))
    return out


def _bg() -> Image.Image:
    canvas = CANVAS
    bg = Image.new("RGB", (canvas, canvas), BASE_COLOR)
    overlay = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    for y in range(canvas):
        a = int(18 * (y / canvas))
        od.line([(0, y), (canvas, y)], fill=(220, 200, 170, a))
    return Image.alpha_composite(bg.convert("RGBA"), overlay)


def _place_pkg(
    base: Image.Image,
    pkg: Image.Image,
    *,
    max_w: int,
    max_h: int,
    px: int,
    py_from_bottom: int,
) -> Tuple[Image.Image, Tuple[int, int, int, int]]:
    p = pkg.copy()
    p.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
    py = CANVAS - p.height - py_from_bottom
    out = base.copy()
    out.alpha_composite(p, (px, py))
    return out, (px, py, px + p.width, py + p.height)


def _sizzle_from_sample(sample: Image.Image, box: Tuple[int, int, int, int]) -> Image.Image:
    """見本右下のシズル帯を切り出し（仮想）。"""
    return sample.crop(box).convert("RGBA")


def _build_photo_layer(theme_id: str, truth: Path, sample: Image.Image) -> Image.Image:
    """レイヤ1（仮想）。テーマで袋の大きさ比とシズル枠を変える。"""
    bx0, _, _, _ = _badge_bbox_1200()
    base = _bg()
    pkg = Image.open(truth).convert("RGBA")
    sample_rgb = sample.convert("RGB").resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)

    if theme_id == "N1":
        # 袋主役・シズル小さめ／無し
        out, _ = _place_pkg(base, pkg, max_w=min(720, bx0 - 60), max_h=980, px=80, py_from_bottom=50)
        return out

    if theme_id == "T2":
        # 袋左＋原料寄りの静物帯（見本右下を弱く）
        out, _ = _place_pkg(base, pkg, max_w=480, max_h=820, px=40, py_from_bottom=60)
        sz = _sizzle_from_sample(sample_rgb, (620, 420, 1100, 1100))
        sz = sz.resize((420, 420), Image.Resampling.LANCZOS)
        # 円マスクで静物っぽく
        mask = Image.new("L", sz.size, 0)
        ImageDraw.Draw(mask).ellipse([0, 0, sz.size[0] - 1, sz.size[1] - 1], fill=255)
        sz.putalpha(mask)
        out.alpha_composite(sz, (700, 620))
        return out

    if theme_id == "T3":
        # 食べ方シーン：皿風楕円＋袋
        out, _ = _place_pkg(base, pkg, max_w=500, max_h=860, px=50, py_from_bottom=40)
        plate = Image.new("RGBA", (520, 360), (0, 0, 0, 0))
        pd = ImageDraw.Draw(plate)
        pd.ellipse([10, 40, 510, 340], fill=(245, 240, 230, 255), outline=(210, 200, 180, 255), width=3)
        sz = _sizzle_from_sample(sample_rgb, (650, 500, 1150, 1050)).resize((420, 260), Image.Resampling.LANCZOS)
        plate.alpha_composite(sz, (50, 70))
        out.alpha_composite(plate, (620, 680))
        return out

    # T4 ベネフィット: 食感マクロ円形インセット
    out, pkg_box = _place_pkg(base, pkg, max_w=520, max_h=880, px=60, py_from_bottom=50)
    # 正本の中央を拡大
    pw, ph = pkg.size
    cx0, cy0 = pw // 2, ph // 2
    r = min(pw, ph) // 5
    macro = pkg.crop((cx0 - r, cy0 - r, cx0 + r, cy0 + r)).resize((340, 340), Image.Resampling.LANCZOS)
    macro = macro.filter(ImageFilter.SMOOTH_MORE)
    mask = Image.new("L", (340, 340), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, 339, 339], fill=255)
    macro.putalpha(mask)
    # 白縁
    ring = Image.new("RGBA", (360, 360), (0, 0, 0, 0))
    ImageDraw.Draw(ring).ellipse([0, 0, 359, 359], fill=(255, 255, 255, 255))
    ring.alpha_composite(macro, (10, 10))
    out.alpha_composite(ring, (740, 640))
    return out


def _draw_tracked_line(
    canvas: Image.Image,
    text: str,
    *,
    x: int,
    y_ink: int,
    font,
    fill: Tuple[int, int, int, int],
    tracking_px: int,
    x_scale: float,
    shadow: Optional[Tuple[int, int, int, int]] = None,
    shadow_off: int = 3,
    stroke_width: int = 0,
    stroke_fill: Optional[Tuple[int, int, int, int]] = None,
) -> int:
    """インク上端 (x, y_ink) から1字ずつ。戻り値は右端x。"""
    cursor = float(x)
    probe = ImageDraw.Draw(Image.new("RGBA", (8, 8), (0, 0, 0, 0)))
    sw = max(0, int(stroke_width))
    sf = stroke_fill if stroke_fill is not None else fill
    for ch in text:
        if ch == " ":
            cursor += max(8.0, font.getlength(" ") * x_scale)
            continue
        bb = probe.textbbox((0, 0), ch, font=font, stroke_width=sw)
        gw, gh = bb[2] - bb[0], bb[3] - bb[1]
        if gw <= 0 or gh <= 0:
            cursor += tracking_px
            continue
        pad = shadow_off + sw + 2
        glyph = Image.new("RGBA", (gw + pad * 2, gh + pad * 2), (0, 0, 0, 0))
        gd = ImageDraw.Draw(glyph)
        ox, oy = pad - bb[0], pad - bb[1]
        if shadow:
            gd.text(
                (ox + shadow_off, oy + shadow_off),
                ch,
                font=font,
                fill=shadow,
                stroke_width=sw,
                stroke_fill=shadow,
            )
        gd.text((ox, oy), ch, font=font, fill=fill, stroke_width=sw, stroke_fill=sf)
        crop = glyph.split()[-1].getbbox()
        if not crop:
            cursor += tracking_px
            continue
        glyph = glyph.crop(crop)
        tw = max(1, int(round(glyph.size[0] * x_scale)))
        th = glyph.size[1]
        if x_scale != 1.0:
            glyph = glyph.resize((tw, th), Image.Resampling.LANCZOS)
        else:
            tw = glyph.size[0]
        canvas.alpha_composite(glyph, (int(round(cursor)), int(y_ink)))
        cursor += tw + tracking_px
    return int(round(cursor))


def _stamp_appeal(base: Image.Image, upper: str, lower: str) -> Image.Image:
    """レイヤ2: Pillow アピール2段（見本実測の大きさ・位置・字間）。"""
    cfg = _appeal_box()
    right = int(cfg.get("rightLimitX", 880))
    font_id = str(cfg.get("fontId") or "genjyuu_gothic_like")
    u = cfg.get("upper") or {}
    l = cfg.get("lower") or {}
    out = base.convert("RGBA")
    fill = (35, 28, 22, 255)
    shadow = (40, 30, 20, 80)
    font_u = _load_font(font_id, int(u.get("fontPx") or 59))
    font_l = _load_font(font_id, int(l.get("fontPx") or 134))
    ru = _draw_tracked_line(
        out,
        upper,
        x=int(u.get("x") or 48),
        y_ink=int(u.get("y") or 31),
        font=font_u,
        fill=fill,
        tracking_px=int(u.get("trackingPx") or 3),
        x_scale=float(u.get("xScale") or 1.0),
        shadow=shadow,
        shadow_off=2,
    )
    rl = _draw_tracked_line(
        out,
        lower,
        x=int(l.get("x") or 40),
        y_ink=int(l.get("y") or 97),
        font=font_l,
        fill=fill,
        tracking_px=int(l.get("trackingPx") or 8),
        x_scale=float(l.get("xScale") or 0.858),
        shadow=shadow,
        shadow_off=5,
    )
    if max(ru, rl) > right:
        LOG.warning("appeal overflow right=%s ru=%s rl=%s", right, ru, rl)
    return out


def annotate_v1(im: Image.Image) -> Image.Image:
    out = im.convert("RGBA")
    draw = ImageDraw.Draw(out)
    cx, cy, diameter = _gold_params()
    r = diameter / 2.0
    bx0, by0, bx1, by1 = _badge_bbox_1200()
    draw.rectangle([bx0, by0, bx1, by1], outline=(0, 180, 0, 255), width=4)
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], outline=(220, 0, 0, 255), width=5)
    draw.ellipse([cx - 4, cy - 4, cx + 4, cy + 4], fill=(220, 0, 0, 255))
    draw.text((bx0 + 8, by1 + 8), "badgeBBox / goldCircle SampleOverflow", fill=(20, 20, 20, 255))
    return out.convert("RGB")


def _contact_sheet(paths: List[Path], labels: List[str], out_path: Path) -> None:
    """選定UI風の横並び。"""
    n = len(paths)
    thumb = 280
    pad = 16
    header = 48
    w = pad + n * (thumb + pad)
    h = pad + header + thumb + 72 + pad
    sheet = Image.new("RGB", (w, h), (245, 245, 245))
    draw = ImageDraw.Draw(sheet)
    font = _load_font("genjyuu_gothic_like", 22)
    for i, (p, lab) in enumerate(zip(paths, labels)):
        im = Image.open(p).convert("RGB")
        im.thumbnail((thumb, thumb), Image.Resampling.LANCZOS)
        x = pad + i * (thumb + pad)
        y = pad + header
        sheet.paste(im, (x, y))
        draw.text((x, pad + 8), lab, fill=(20, 20, 20), font=font)
        draw.text((x, y + thumb + 8), "テーマ名先出し→画像", fill=(80, 80, 80), font=font)
    sheet.save(out_path, quality=92)


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    work = Path(DEFAULT_WORK_ROOT)
    out_root = work / SUB_TEST_OUT / "b5_main_base_virt" / f"{JAN}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    out_root.mkdir(parents=True, exist_ok=True)
    meta_dir = out_root / "_meta"
    meta_dir.mkdir(exist_ok=True)

    truth = _truth_path(work)
    sample = Image.open(_sample_path(work))
    gold_path = resolve_kinmaru(work)
    cx, cy, diameter = _gold_params()
    gold = _load_gold_disk(work, diameter)
    LOG.info("truth=%s gold=%s out=%s", truth.name, gold_path, out_root)

    # 金丸単体プレビュー（スタンプ位置確認）
    preview = Image.new("RGBA", (CANVAS, CANVAS), (40, 40, 40, 255))
    preview = _stamp_gold(preview, gold)
    annotate_v1(preview.convert("RGB")).save(out_root / "V0_gold_stamp_annot.jpg", quality=92)

    draft_paths: List[Path] = []
    labels: List[str] = []
    proposal_out: List[Dict[str, Any]] = []

    for prop in PROPOSALS:
        tid = prop["themeId"]
        upper, lower = prop["upper"], prop["lower"]
        u_len, l_len = _zenkaku_len(upper), _zenkaku_len(lower)
        if u_len > 18 or l_len > 10:
            raise SystemExit(f"字数NG {tid}: upper={u_len} lower={l_len}")

        photo = _build_photo_layer(tid, truth, sample)
        with_appeal = _stamp_appeal(photo, upper, lower)
        final = _stamp_gold(with_appeal, gold).convert("RGB")
        dest = out_root / f"{tid}.jpg"
        final.save(dest, quality=92)
        draft_paths.append(dest)
        labels.append(f"{tid} {prop['themeName']}")
        proposal_out.append(
            {
                **prop,
                "upperZenkaku": u_len,
                "lowerZenkaku": l_len,
                "file": dest.name,
            }
        )
        LOG.info("draft %s → %s (%s/%s)", tid, dest.name, u_len, l_len)

    # 選定UI風
    _contact_sheet(draft_paths, labels, out_root / "select_contact_sheet.jpg")

    # 見本左上 vs 仮想 N1 左上（900x220）
    sample_rgb = Image.open(_sample_path(work)).convert("RGB").resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
    virt_n1 = Image.open(draft_paths[0]).convert("RGB")
    crop_box = (0, 0, 900, 230)
    cmp = Image.new("RGB", (900, 470), (30, 30, 30))
    cmp.paste(sample_rgb.crop(crop_box), (0, 0))
    cmp.paste(virt_n1.crop(crop_box), (0, 240))
    ImageDraw.Draw(cmp).text((8, 8), "sample 5614-oya top", fill=(255, 220, 80))
    ImageDraw.Draw(cmp).text((8, 248), "virt N1 top (appeal measured)", fill=(255, 220, 80))
    cmp.save(out_root / "appeal_vs_sample_top.jpg", quality=92)

    # 仮想採用＝N1（人が選ぶ想定）。後続数字載せ
    adopted = out_root / "N1.jpg"
    adopted_png = out_root / f"{JAN}_単体_virt_adopted_N1.png"
    Image.open(adopted).convert("RGBA").save(adopted_png)
    v3_paths = []
    for n in NS:
        img, meta = compose_rakuten_layered(
            base_path=adopted_png,
            set_count=n,
            unit=UNIT,
            work_root=work,
            text_color="1",
        )
        p = out_root / f"{JAN}_V3_adopted_N1_n{n}_{UNIT}.jpg"
        img.save(p, quality=92)
        (meta_dir / f"V3_n{n}_meta.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        v3_paths.append(p.name)
        LOG.info("V3 n=%s → %s", n, p.name)

    # 注釈（N1）
    annotate_v1(Image.open(adopted)).save(out_root / "N1_V1_annot.jpg", quality=92)

    proposals_json = {
        "jan": JAN,
        "parentSkuHint": "sanky-4906283045119",
        "product": "たべるおだし・かつお",
        "baseColor": "warm_white",
        "plan": "A",
        "goldSource": str(gold_path),
        "goldCircle1200SampleOverflow": {"cx": cx, "cy": cy, "diameter": diameter},
        "skippedThemes": [{"themeId": "T1", "reason": "材料に賞・累計なし（§3.7）"}],
        "proposals": proposal_out,
        "adoptedForVirt": "N1",
        "note": "仮想。AI未使用。02.楽天ベース未書込。金丸は98実素材。",
    }
    (out_root / "b5_proposals.json").write_text(
        json.dumps(proposals_json, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    summary = {
        **proposals_json,
        "outRoot": str(out_root),
        "files": {
            "drafts": [p.name for p in draft_paths],
            "contact": "select_contact_sheet.jpg",
            "V0_gold": "V0_gold_stamp_annot.jpg",
            "V3": v3_paths,
        },
    }
    (out_root / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
