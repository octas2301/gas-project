# -*- coding: utf-8 -*-
"""
B-⑤ 仮想 v7: 上段起点固定→間隔→下段→下線。位置は％。20マス内。
書体 UDモトヤマルベリ。下段白。下線半透明。02には書かない。
"""
from __future__ import annotations

import argparse
import json
import logging
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List

from PIL import Image, ImageDraw

from b5_virt_katsuo import (
    CANVAS,
    JAN,
    _gold_params,
    _load_gold_disk,
    _stamp_gold,
    _truth_path,
    _zenkaku_len,
)
from gemini_image import generate_with_references, make_client
from model_policy import resolve_model_id
from photo_realism_rules import photo_realism_block_ja
from rakuten_layer import load_layout_rules
from text_color import resolve_text_color
from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT, resolve_underline_sample

LOG = logging.getLogger("set_main_image.b5_virt_ai_headline")

# 字数ロック後の仮想コピー（上段12-18／下段4-8。T1は材料なし）
PROPOSALS: List[Dict[str, Any]] = [
    {
        "themeId": "N1",
        "themeName": "商品名明示",
        "mode": "name",
        "upper": "だしをとらずにそのまま食べる",
        "lower": "たべるおだし",
        "accentWord": "たべるおだし",
    },
    {
        "themeId": "T2",
        "themeName": "産地・素材",
        "mode": "talk",
        "upper": "国産かつお節をていねいに",
        "lower": "素材を味わう",
        "accentWord": "味わう",
    },
    {
        "themeId": "T3",
        "themeName": "使い方・シーン",
        "mode": "talk",
        "upper": "お茶漬けにもおつまみにも",
        "lower": "さっと一振り",
        "accentWord": "さっと一振り",
    },
    {
        "themeId": "T4",
        "themeName": "ベネフィット",
        "mode": "talk",
        "upper": "口にとろけるやわらかい食感",
        "lower": "だしの旨み",
        "accentWord": "旨み",
    },
]

THEME_SHOT: Dict[str, str] = {
    "N1": "パッケージ正面を大きく左〜中央。シズルは小さくするか無し。袋の商品名が読める。",
    "T2": "袋は左。右にかつお節・原料の静物。料理シーンは主役にしない。",
    "T3": "袋は左。右にお茶漬けやおつまみなど使い方のシーン（材料にある用途のみ）。",
    "T4": "袋は左。右に食感マクロ（とろける・湯気）。円形インセット可。",
}

BASE_TEXT_LIGHT = "#0d1321"
BASE_TEXT_DARK = "#ffffff"

# 98色見本の上から: 青 / えんじ / 緑 / 金
_UNDERLINE_BANDS = (
    {"id": "2", "y0": 161, "y1": 237},
    {"id": "1", "y0": 349, "y1": 425},
    {"id": "4", "y0": 570, "y1": 646},
    {"id": "5", "y0": 801, "y1": 878},
)


def _crop_underline_stroke(sheet: Image.Image, color_id: str) -> Image.Image:
    """色見本シートから、指定色に近い1本だけ切り出す。"""
    cid = str(color_id or "1")
    if cid == "3":
        cid = "1"
    band = next((b for b in _UNDERLINE_BANDS if b["id"] == cid), _UNDERLINE_BANDS[1])
    pad = 24
    y0 = max(0, int(band["y0"]) - pad)
    y1 = min(sheet.size[1], int(band["y1"]) + pad)
    crop = sheet.convert("RGBA").crop((0, y0, sheet.size[0], y1))
    return crop


def _trim_ink(im: Image.Image, *, alpha_min: int = 40) -> Image.Image:
    """透明余白を落としてインク外接にする。"""
    rgba = im.convert("RGBA")
    w, h = rgba.size
    minx, miny, maxx, maxy = w, h, -1, -1
    px = rgba.load()
    for y in range(h):
        for x in range(w):
            if px[x, y][3] > alpha_min:
                if x < minx:
                    minx = x
                if y < miny:
                    miny = y
                if x > maxx:
                    maxx = x
                if y > maxy:
                    maxy = y
    if maxx < 0:
        return rgba
    return rgba.crop((minx, miny, maxx + 1, maxy + 1))


def _appeal_cfg() -> Dict[str, Any]:
    rakuten = load_layout_rules().get("rakuten") or {}
    return rakuten.get("appealHeadline1200") or {}


def _base_text_hex(base_color: str) -> str:
    if (base_color or "").strip().lower() in ("black", "dark", "navy"):
        return BASE_TEXT_DARK
    return BASE_TEXT_LIGHT


def _assert_counts(prop: Dict[str, Any]) -> None:
    cfg = _appeal_cfg()
    u = cfg.get("upper") or {}
    lo = cfg.get("lower") or {}
    nu = _zenkaku_len(prop["upper"])
    nl = _zenkaku_len(prop["lower"])
    umin, umax = int(u.get("minZenkaku", 12)), int(u.get("maxZenkaku", 18))
    lmin, lmax = int(lo.get("minZenkaku", 4)), int(lo.get("maxZenkaku", 8))
    if not (umin <= nu <= umax):
        raise SystemExit(f"{prop['themeId']} 上段字数 {nu} が {umin}-{umax} 外: {prop['upper']}")
    if not (lmin <= nl <= lmax):
        raise SystemExit(f"{prop['themeId']} 下段字数 {nl} が {lmin}-{lmax} 外: {prop['lower']}")


def _gold_placed(work: Path) -> Image.Image:
    """ベージュ1200に98を見本座標で置いた参照（AIに位置と素材を渡す）。"""
    bg = Image.new("RGBA", (CANVAS, CANVAS), (250, 247, 240, 255))
    gold = _load_gold_disk(work, _gold_params()[2])
    return _stamp_gold(bg, gold)


def _headline_stack() -> Dict[str, Any]:
    """上段起点固定 → 間隔 → 下段 → 下線。1200換算と％の両方。"""
    cfg = _appeal_cfg()
    u = cfg.get("upper") or {}
    lo = cfg.get("lower") or {}
    ul = cfg.get("underline") or {}
    ch = cfg.get("headlineChain") or {}
    canvas = float(CANVAS)
    ux = int(ch.get("upperOriginX", u.get("x", 48)))
    uy = int(ch.get("upperOriginY", u.get("y", 31)))
    uh = int(u.get("targetInkH", 46))
    gap = int(ch.get("gapAfterUpperPx", 20))
    lx = int(lo.get("x", 40))
    ly = uy + uh + gap
    lh = int(lo.get("targetInkH", 88))
    overlap = int(ul.get("overlapLowerPx", 40))
    ul_h = int(ul.get("h", 76))
    ul_x = int(ul.get("x", 52))
    ul_w = int(ul.get("w", 720))
    lower_bottom = ly + lh
    ul_y = lower_bottom - overlap
    ul_bottom = ul_y + ul_h
    cap = int(ch.get("yHardCap", 228))
    right_pct = float(ch.get("rightLimitPct", 70))

    def pct(v: float) -> float:
        return round(100.0 * v / canvas, 2)

    return {
        "upper": {
            "x": ux,
            "y": uy,
            "h": uh,
            "yBottom": uy + uh,
            "xPct": pct(ux),
            "yPct": pct(uy),
            "hPct": pct(uh),
        },
        "gapPx": gap,
        "gapPct": pct(gap),
        "lower": {
            "x": lx,
            "y": ly,
            "h": lh,
            "yBottom": lower_bottom,
            "xPct": pct(lx),
            "yPct": pct(ly),
            "hPct": pct(lh),
        },
        "underline": {
            "x": ul_x,
            "y": ul_y,
            "h": ul_h,
            "w": ul_w,
            "yBottom": ul_bottom,
            "overlap": overlap,
            "xPct": pct(ul_x),
            "yPct": pct(ul_y),
            "hPct": pct(ul_h),
            "yBottomPct": pct(ul_bottom),
            "wPct": pct(ul_w),
        },
        "yHardCap": cap,
        "yHardCapPct": pct(cap),
        "yBandPct": 20.0,
        "rightLimitPct": right_pct,
    }


def _ul_place() -> Dict[str, int]:
    ul = _headline_stack()["underline"]
    return {
        "x": int(ul["x"]),
        "y": int(ul["y"]),
        "h": int(ul["h"]),
        "w": int(ul["w"]),
        "yBottom": int(ul["yBottom"]),
        "overlap": int(ul["overlap"]),
    }


def _layout_placed(work: Path, ul_ink: Image.Image) -> Image.Image:
    """金丸＋下線を見本番地に置いた地図（完成のベージュ固定ではない）。"""
    canvas = _gold_placed(work)
    pos = _ul_place()
    ink = ul_ink.convert("RGBA")
    if ink.size[1] != pos["h"]:
        ratio = pos["h"] / float(ink.size[1])
        nw = max(1, int(round(ink.size[0] * ratio)))
        ink = ink.resize((nw, pos["h"]), Image.Resampling.LANCZOS)
    canvas.alpha_composite(ink, (pos["x"], pos["y"]))
    return canvas


def _prompt(prop: Dict[str, Any], color: Dict[str, Any]) -> str:
    tid = prop["themeId"]
    upper = prop["upper"]
    lower = prop["lower"]
    hex_c = color.get("hex") or "#800020"
    name_ja = color.get("nameJa") or "えんじ"
    cx, cy, diameter = _gold_params()
    nu = _zenkaku_len(upper)
    nl = _zenkaku_len(lower)
    cfg = _appeal_cfg()
    font_ja = cfg.get("fontNameJa") or "UDモトヤマルベリ"
    font_en = cfg.get("fontNameEn") or "UD Motoya Maru Berry"
    st = _headline_stack()
    u, lo, ul = st["upper"], st["lower"], st["underline"]
    return f"""あなたは楽天商品MAIN（正方形）のデザイナーです。出力解像度が1200でも1024でも、位置はキャンバス比（小数）で置く。1200のピクセル数を1024にそのまま使わない。

【IMAGE_PACKAGE_TRUTH】商品パッケージ正本。色・ラベル・印字・縦横比を変えない。表面のOCR・描き直し禁止。
【IMAGE_LAYOUT_PLACED】金丸と下線の位置地図（1200基準）。完成では同じ比の位置に写す。ベージュは位置確認用で完成背景にしない。
- 金丸: 98素材。中心は右上（1200換算 x={cx}, y={cy}、直径 {diameter}）。模様・形を変えない。数字・「袋」「セット」は書かない。
- 下線: 地図どおりの割合。ずらすな。描き直すな。
【IMAGE_UNDERLINE】下線の筆見本（半透明）。

【キャンバス】1:1。見出しの後ろは写真（白が読める）。商品は見出し帯の下。
【テーマ】{tid} {prop['themeName']}（mode={prop['mode']}）
【構図】{THEME_SHOT.get(tid, '')} 金丸で文字や袋の顔を隠さない。

【配置の決め方（この順。中央寄せ禁止）】
比は小数。完成に数値・単位・ガイドを描かない。
1. 上段の左上を固定。左 = 幅×{u['xPct']/100:.4f} 、上 = 高さ×{u['yPct']/100:.4f}（1200換算 x={u['x']}, y={u['y']}）。起点。動かさない。
2. 上段インク高さ = 高さ×{u['hPct']/100:.4f}。
3. 上段下端から 高さ×{st['gapPct']/100:.4f} だけ空けて下段開始。下段左 = 幅×{lo['xPct']/100:.4f}。下段の縦位置は計算で決める。目測で下げない。
4. 下段インク高さ = 高さ×{lo['hPct']/100:.4f}。これ以上大きくしない。
5. 下線は下段の下側に重ねる（重なりは下段高さの約半分弱）。下線上端＝下段下端マイナスその重なり。下線高さ = 高さ×{ul['hPct']/100:.4f}。下線下端は 高さ×{ul['yBottomPct']/100:.4f}。背面に置き、文字は前面。
6. 見出しと下線の右端は 幅×{st['rightLimitPct']/100:.2f} 未満（金丸の左。列0から6）。列7から9は金丸専用。

【20マス（最重要）】
画像を縦10・横10に分ける。上2行だけ（高さの0から0.20まで）に見出しと下線を収める。
高さ0.20より下に文字も下線も1ピクセルも出すな。
後工程のセット数字も同じ上2行の右。今は金丸を空にする。

【書体】{font_ja}（{font_en}）。丸ゴシックUD。源柔・游ゴシック・明朝にしない。

【見出し・一字一句これだけ】
上段（小さめ・1行・{nu}字）: {upper}
下段（大きめ・1行・{nl}字）: {lower}
役割ラベルや鉤括弧は描かない。パッケージ以外の文字はこの2行のみ。

【色】
- 上段・下段は白 #FFFFFF
- 下線は {name_ja} {hex_c}、半透明 70〜85パーセント
- 文字を {hex_c} にしない。白縁・黒縁禁止

【アンダーライン】IMAGE_UNDERLINE の筆。左垂直・右尖り・上弧。直線バー・横断帯にしない。1本だけ。
【棒線禁止】文字を貫く線、飾り棒、二重線、矩形枠禁止。

【絶対禁止】
- 完成に％・px・座標・小数・ガイド線・寸法数字を描く
- 緑の線・緑の枠・案内格子
- 白縁・多重縁・グロー
- 金丸の描き直し・位置ずらし・金丸内の数字
- 下線を高さ0.20より下へずらす
- 袋を複数並べる
- 指定以外の累計・No.1
- 正本パッケージの改変

{photo_realism_block_ja()}
"""


def _to_1200(raw: bytes) -> Image.Image:
    im = Image.open(BytesIO(raw))
    if im.mode != "RGBA":
        im = im.convert("RGBA")
    return im.resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)


def _write_grid20(img_path: Path, dest: Path) -> Dict[str, Any]:
    """10x10仮線。見出し帯（上3行・金丸左）の白/えんじだけ数える。"""
    im = Image.open(img_path).convert("RGB").resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
    w, h = im.size
    cw, ch = w / 10.0, h / 10.0
    px = im.load()
    gold_left = 931
    y_scan = int(round(3 * ch))
    white_c = [[0] * 10 for _ in range(10)]
    ul_c = [[0] * 10 for _ in range(10)]
    max_y = -1
    for y in range(y_scan):
        cy = min(9, int(y / ch))
        for x in range(min(w, gold_left)):
            cx = min(9, int(x / cw))
            r, g, b = px[x, y]
            white = r >= 225 and g >= 225 and b >= 225
            ul = r >= 90 and r <= 170 and g <= 55 and b <= 55 and (r - g) >= 45
            if white:
                white_c[cy][cx] += 1
                max_y = max(max_y, y)
            elif ul:
                ul_c[cy][cx] += 1
                max_y = max(max_y, y)
    th = 40
    hit = set()
    for cy in range(10):
        for cx in range(10):
            if white_c[cy][cx] >= th or ul_c[cy][cx] >= th:
                hit.add((cx, cy))
    overflow = {(cx, cy) for (cx, cy) in hit if cy >= 2}
    overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    d.rectangle([0, 0, w - 1, 239], fill=(255, 220, 0, 36))
    for cy in range(10):
        for cx in range(10):
            x0, y0 = int(cx * cw), int(cy * ch)
            x1, y1 = int((cx + 1) * cw) - 1, int((cy + 1) * ch) - 1
            if (cx, cy) in hit:
                col = (255, 40, 40, 100) if cy >= 2 else (0, 180, 255, 85)
                d.rectangle([x0, y0, x1, y1], fill=col)
            d.rectangle([x0, y0, x1, y1], outline=(0, 255, 80, 170), width=1)
    d.line([(0, 240), (w, 240)], fill=(255, 60, 0, 255), width=3)
    mixed = Image.alpha_composite(im.convert("RGBA"), overlay).convert("RGB")
    dest.parent.mkdir(exist_ok=True)
    mixed.save(dest, quality=90)
    return {
        "file": img_path.name,
        "hitN": len(hit),
        "in20N": len(hit) - len(overflow),
        "out20N": len(overflow),
        "overflow": [list(p) for p in sorted(overflow)],
        "maxInkY": max_y,
        "ok_band": len(overflow) == 0,
    }


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser()
    ap.add_argument("--text-color", default="1", help="Cと同じ: 1|えんじ|青|黒|緑|茶（下線と金丸数字）")
    ap.add_argument("--base-color", default="beige", help="beige|warm_white|soft_gray")
    args = ap.parse_args()
    color = resolve_text_color(args.text_color)

    for prop in PROPOSALS:
        _assert_counts(prop)

    work = Path(DEFAULT_WORK_ROOT)
    out_root = (
        work
        / SUB_TEST_OUT
        / "b5_main_base_virt"
        / f"{JAN}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_ai_headline"
    )
    out_root.mkdir(parents=True, exist_ok=True)
    meta_dir = out_root / "_meta"
    meta_dir.mkdir(exist_ok=True)

    truth = _truth_path(work)
    ul_src = resolve_underline_sample(work)
    if ul_src is None or not ul_src.is_file():
        raise SystemExit("98フォルダにセールストークアンダーライン素材がありません")
    ul_ink = _trim_ink(_crop_underline_stroke(Image.open(ul_src), str(color.get("id") or "1")))
    ul_path = meta_dir / "underline_stroke.png"
    ul_ink.save(ul_path)
    layout_ref = _layout_placed(work, ul_ink)
    layout_path = meta_dir / "layout_placed.png"
    layout_ref.save(layout_path)
    stack = _headline_stack()
    (meta_dir / "headline_coords.json").write_text(
        json.dumps(stack, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    client = make_client()
    model_id = resolve_model_id(client)
    LOG.info(
        "model=%s underline=%s %s font=UDモトヤマルベリ layoutPassToAi out=%s",
        model_id,
        color.get("id"),
        color.get("hex"),
        out_root,
    )

    results: List[Dict[str, Any]] = []
    draft_paths: List[Path] = []
    labels: List[str] = []

    for prop in PROPOSALS:
        tid = prop["themeId"]
        prompt = _prompt(prop, color)
        (meta_dir / f"{tid}_prompt.txt").write_text(prompt, encoding="utf-8")
        row: Dict[str, Any] = {
            **prop,
            "upperZenkaku": _zenkaku_len(prop["upper"]),
            "lowerZenkaku": _zenkaku_len(prop["lower"]),
            "ok": False,
        }
        try:
            raw, gmeta = generate_with_references(
                client=client,
                model_id=model_id,
                prompt=prompt,
                image_paths=[truth, layout_path, ul_path],
                image_roles=["IMAGE_PACKAGE_TRUTH", "IMAGE_LAYOUT_PLACED", "IMAGE_UNDERLINE"],
                aspect_ratio="1:1",
                image_size="1K",
            )
            dest = out_root / f"{tid}.jpg"
            rgb1200 = _to_1200(raw).convert("RGB")
            rgb1200.save(dest, quality=92)
            rgb1200.save(meta_dir / f"{tid}_ai_raw.jpg", quality=90)
            grid_row = _write_grid20(dest, out_root / "_grid20" / f"{tid}_grid20.jpg")
            draft_paths.append(dest)
            labels.append(f"{tid} {prop['themeName']}")
            row["ok"] = True
            row["file"] = dest.name
            row["apiPath"] = gmeta.get("apiPath")
            row["grid20"] = grid_row
            LOG.info(
                "OK %s -> %s grid hit=%s overflow=%s maxY=%s",
                tid,
                dest.name,
                grid_row["hitN"],
                grid_row["out20N"],
                grid_row["maxInkY"],
            )
        except Exception as e:
            LOG.exception("FAIL %s", tid)
            row["error"] = str(e)[:400]
        results.append(row)

    if draft_paths:
        from b5_virt_katsuo import _contact_sheet, _sample_path

        _contact_sheet(draft_paths, labels, out_root / "select_contact_sheet.jpg")
        sample_p = _sample_path(work)
        if sample_p.is_file() and draft_paths:
            sample_rgb = Image.open(sample_p).convert("RGB").resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
            virt_rgb = Image.open(draft_paths[0]).convert("RGB")
            crop_box = (0, 0, 900, 250)
            cmp = Image.new("RGB", (900, 520), (30, 30, 30))
            cmp.paste(sample_rgb.crop(crop_box), (0, 0))
            cmp.paste(virt_rgb.crop(crop_box), (0, 260))
            ImageDraw.Draw(cmp).text((8, 8), "sample 5614-oya top", fill=(255, 220, 80))
            ImageDraw.Draw(cmp).text((8, 268), "virt N1 top (chain v7)", fill=(255, 220, 80))
            cmp.save(out_root / "appeal_vs_sample_top.jpg", quality=92)

    pos = _ul_place()
    summary = {
        "jan": JAN,
        "plan": "A_ai_headline_v7",
        "model": model_id,
        "goldPassToAi": True,
        "layoutPassToAi": True,
        "font": "UDモトヤマルベリ",
        "headlineFill": "#ffffff",
        "headlineChain": stack,
        "coords": {
            "upper": {"x": stack["upper"]["x"], "y": stack["upper"]["y"]},
            "lower": {"x": stack["lower"]["x"], "y": stack["lower"]["y"]},
            "underline": pos,
        },
        "textBand": {"y0": 0, "y1": 240, "cells": 20, "grid": "10x10"},
        "zenkaku": {"upper": "12-18", "lower": "4-8"},
        "textColorAccent": {
            "id": color.get("id"),
            "nameJa": color.get("nameJa"),
            "hex": color.get("hex"),
        },
        "note": "上段起点固定→間隔20→下段→下線。位置は％。20マス内。02未書込。",
        "outRoot": str(out_root),
        "proposals": results,
    }
    (out_root / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_root / "b5_proposals.json").write_text(
        json.dumps(
            {
                "jan": JAN,
                "headlineRender": "ai_in_image",
                "goldPassToAi": True,
                "textColorAccent": summary["textColorAccent"],
                "proposals": results,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if sum(1 for r in results if r.get("ok")) else 1


if __name__ == "__main__":
    raise SystemExit(main())
