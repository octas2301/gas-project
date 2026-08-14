# -*- coding: utf-8 -*-
"""
Amazon サブ画像 PoC — 10パターン（写真背景＋競合掛け合わせ）

① 実写写真背景バージョン
② 競合パーツ掛け合わせ
③ サブ画像パターン10種
④ MAIN相当（商品だけの白抜き系）は作らない

テスト出力のみ（マスタ／R2／GAS 書込なし）。
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, ImageDraw, ImageFilter, ImageFont

from gemini_image import generate_with_references, make_client, save_as_jpeg
from model_policy import resolve_model_id
from sub_image_prompts import (
    SUB_PATTERNS,
    logic_steps_for_humans,
    prompt_comp_mash_ai,
    prompt_scene_ai,
)
from work_paths import default_work_root, meta_dir_for

LOG = logging.getLogger("set_main_image.sub_image_poc")

BOARD_W = 2400
PANEL_H = 900
HEADER_H = 120
FOOTER_H = 160
SQUARE = 1200

BG_DIR = Path(__file__).resolve().parent / "_sub_bg_photos"
SAMPLES = Path(__file__).resolve().parent / "_portrait_samples"


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _font(size: int) -> ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\YuGothB.ttc"),
        Path(r"C:\Windows\Fonts\meiryo.ttc"),
        Path(r"C:\Windows\Fonts\msgothic.ttc"),
        Path(r"C:\Windows\Fonts\arial.ttf"),
    ]
    for p in candidates:
        if p.is_file():
            try:
                return ImageFont.truetype(str(p), size=size)
            except Exception:
                continue
    return ImageFont.load_default()


def _fit(im: Image.Image, box: Tuple[int, int, int, int], pad: int = 12) -> Image.Image:
    x0, y0, x1, y1 = box
    bw, bh = max(1, x1 - x0 - pad * 2), max(1, y1 - y0 - pad * 2)
    src = im.convert("RGBA") if im.mode != "RGBA" else im
    sw, sh = src.size
    scale = min(bw / sw, bh / sh)
    nw, nh = max(1, int(sw * scale)), max(1, int(sh * scale))
    resized = src.resize((nw, nh), Image.Resampling.LANCZOS)
    canvas = Image.new("RGBA", (bw, bh), (245, 245, 245, 255))
    ox, oy = (bw - nw) // 2, (bh - nh) // 2
    canvas.alpha_composite(resized, (ox, oy))
    return canvas


def _wrap(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_w: int) -> List[str]:
    lines: List[str] = []
    for para in (text or "").split("\n"):
        cur = ""
        for ch in para:
            trial = cur + ch
            if draw.textlength(trial, font=font) <= max_w:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = ch
        lines.append(cur)
    return lines or [""]


def build_annot_board(
    *,
    mode_ja: str,
    mode_id: str,
    slot: Dict[str, str],
    step_index: int,
    step_total: int,
    prompt_summary_ja: str,
    left_img: Image.Image,
    left_caption: str,
    mid_img: Optional[Image.Image],
    mid_caption: str,
    right_img: Image.Image,
    right_caption: str,
) -> Image.Image:
    cols = 3
    panel_w = BOARD_W // cols
    H = HEADER_H + PANEL_H + FOOTER_H
    board = Image.new("RGB", (BOARD_W, H), (32, 36, 44))
    draw = ImageDraw.Draw(board)
    font_h = _font(34)
    font_s = _font(24)
    font_t = _font(20)

    title = f"[{step_index}/{step_total}] {mode_ja}（{mode_id}）｜{slot['ja_title']}"
    draw.text((24, 28), title, fill=(255, 220, 120), font=font_h)
    draw.text((24, 78), f"目的: {slot['ja_goal']}", fill=(200, 210, 220), font=font_s)

    panels = [
        (left_img, left_caption, (70, 130, 180)),
        (mid_img, mid_caption, (180, 120, 60)),
        (right_img, right_caption, (80, 160, 100)),
    ]
    for i, (pim, cap, color) in enumerate(panels):
        x0 = i * panel_w
        y0 = HEADER_H
        draw.rectangle([x0, y0, x0 + panel_w - 2, y0 + PANEL_H], outline=color, width=4)
        label_h = 48
        draw.rectangle([x0, y0, x0 + panel_w - 2, y0 + label_h], fill=color)
        draw.text((x0 + 12, y0 + 10), cap, fill=(255, 255, 255), font=font_s)
        if pim is not None:
            fitted = _fit(pim, (0, 0, panel_w - 8, PANEL_H - label_h - 8))
            board.paste(
                fitted.convert("RGB"),
                (x0 + 4, y0 + label_h + 4),
                fitted.split()[-1] if fitted.mode == "RGBA" else None,
            )
        else:
            box = [x0 + 16, y0 + label_h + 20, x0 + panel_w - 16, y0 + PANEL_H - 20]
            draw.rectangle(box, fill=(250, 250, 250))
            lines = _wrap(draw, prompt_summary_ja, font_t, panel_w - 48)
            ty = y0 + label_h + 36
            for ln in lines[:18]:
                draw.text((x0 + 28, ty), ln, fill=(30, 30, 30), font=font_t)
                ty += 28

    foot = "サブ画像のみ／MAIN白抜き禁止。写真背景＋競合掛け合わせ。テスト出力（マスタ未書込）。"
    draw.text((24, HEADER_H + PANEL_H + 40), foot, fill=(180, 190, 200), font=font_s)
    draw.text(
        (24, HEADER_H + PANEL_H + 90),
        f"slot={slot['id']}  mode={mode_id}",
        fill=(140, 150, 160),
        font=font_t,
    )
    return board


def _default_own(work: Path) -> Path:
    base = work / "01.amazon白抜きベース"
    if base.is_dir():
        pngs = sorted(base.glob("*.png")) + sorted(base.glob("*.jpg"))
        pngs = [p for p in pngs if p.is_file() and not p.name.startswith("_")]
        if pngs:
            return pngs[0]
    return SAMPLES / "126.jpg"


def _default_comps() -> List[Path]:
    names = ["125.jpg", "126.jpg", "127.jpg", "128.jpg"]
    out = [SAMPLES / n for n in names if (SAMPLES / n).is_file()]
    return out or [SAMPLES / "125.jpg"]


def _load_rgba(path: Path) -> Image.Image:
    return Image.open(path).convert("RGBA")


def _bg_path(bg_key: str) -> Path:
    p = BG_DIR / f"{bg_key}.jpg"
    if not p.is_file():
        # fallback
        alts = sorted(BG_DIR.glob("bg_*.jpg"))
        if not alts:
            raise FileNotFoundError(f"背景写真なし: {BG_DIR}")
        return alts[0]
    return p


def _square_cover(im: Image.Image, size: int = SQUARE) -> Image.Image:
    src = im.convert("RGBA")
    sw, sh = src.size
    scale = max(size / sw, size / sh)
    nw, nh = max(1, int(sw * scale)), max(1, int(sh * scale))
    resized = src.resize((nw, nh), Image.Resampling.LANCZOS)
    x0 = (nw - size) // 2
    y0 = (nh - size) // 2
    return resized.crop((x0, y0, x0 + size, y0 + size))


def _trim_alpha(im: Image.Image, pad: int = 8) -> Image.Image:
    src = im.convert("RGBA")
    bbox = src.getchannel("A").getbbox()
    if not bbox:
        return src
    l, t, r, b = bbox
    l = max(0, l - pad)
    t = max(0, t - pad)
    r = min(src.width, r + pad)
    b = min(src.height, b + pad)
    return src.crop((l, t, r, b))


def _soft_shadow(layer: Image.Image, opacity: int = 90) -> Image.Image:
    alpha = layer.getchannel("A")
    sh = Image.new("RGBA", layer.size, (0, 0, 0, 0))
    sh_a = alpha.point(lambda p: int(p * (opacity / 255.0)) if p > 8 else 0)
    sh.putalpha(sh_a)
    sh = sh.filter(ImageFilter.GaussianBlur(radius=12))
    return sh


def _place_product_on_bg(
    bg: Image.Image,
    product: Image.Image,
    *,
    scale: float = 0.42,
    anchor: str = "center_bottom",
) -> Image.Image:
    canvas = _square_cover(bg, SQUARE)
    prod = _trim_alpha(product)
    pw = max(1, int(SQUARE * scale))
    ratio = pw / prod.width
    ph = max(1, int(prod.height * ratio))
    prod_r = prod.resize((pw, ph), Image.Resampling.LANCZOS)
    if anchor == "center_bottom":
        ox = (SQUARE - pw) // 2
        oy = SQUARE - ph - int(SQUARE * 0.08)
    elif anchor == "left":
        ox = int(SQUARE * 0.08)
        oy = SQUARE - ph - int(SQUARE * 0.1)
    else:
        ox = (SQUARE - pw) // 2
        oy = (SQUARE - ph) // 2
    shadow = _soft_shadow(prod_r)
    canvas.alpha_composite(shadow, (ox + 8, oy + 10))
    canvas.alpha_composite(prod_r, (ox, oy))
    return canvas


def _crop_part(im: Image.Image, box_frac: Tuple[float, float, float, float]) -> Image.Image:
    src = im.convert("RGBA")
    w, h = src.size
    l, t, r, b = box_frac
    return src.crop((int(w * l), int(h * t), int(w * r), int(h * b)))


def compose_comp_mash(
    bg: Image.Image,
    own: Image.Image,
    comps: List[Image.Image],
) -> Image.Image:
    """競合クロップを背景に散らし、中央に自社。"""
    canvas = _square_cover(bg, SQUARE)
    # 競合パーツ（角・帯）
    layouts = [
        ((0.05, 0.05), 0.32, (0.15, 0.10, 0.85, 0.75)),
        ((0.63, 0.08), 0.30, (0.10, 0.20, 0.90, 0.80)),
        ((0.08, 0.62), 0.28, (0.20, 0.15, 0.80, 0.85)),
        ((0.66, 0.64), 0.26, (0.15, 0.15, 0.85, 0.85)),
    ]
    for i, (pos, sc, frac) in enumerate(layouts):
        if not comps:
            break
        src = comps[i % len(comps)]
        part = _crop_part(src, frac)
        pw = max(1, int(SQUARE * sc))
        ph = max(1, int(part.height * (pw / part.width)))
        part_r = part.resize((pw, ph), Image.Resampling.LANCZOS)
        # 少し透明にして「パーツ」感
        a = part_r.getchannel("A").point(lambda p: int(p * 0.85) if p > 0 else 0)
        part_r.putalpha(a)
        ox = int(SQUARE * pos[0])
        oy = int(SQUARE * pos[1])
        canvas.alpha_composite(part_r, (ox, oy))

    # 自社を大きく中央やや下
    return _place_product_on_bg(canvas, own, scale=0.48, anchor="center_bottom")


def write_instruction_logic_md(out_dir: Path, product_name: str) -> None:
    steps = logic_steps_for_humans()
    lines = [
        "# サブ画像 PoC — 10パターン（写真背景＋競合掛け合わせ）",
        "",
        f"自社商品ヒント: **{product_name}**",
        "",
        "## 方針",
        "",
        "- **サブ画像のみ**（白背景の商品だけ＝MAIN相当は作らない）",
        "- **写真背景**: `_sub_bg_photos/` の実写JPGを使用",
        "- **競合掛け合わせ**: 複数競合のクロップ合成（P05）＋AI再構成（P06）",
        "",
        "## 実行順",
        "",
    ]
    for s in steps:
        lines.append(f"### Step {s['step']}. {s['title']}")
        lines.append(s["detail"])
        lines.append("")
    lines.append("## 10パターン")
    lines.append("")
    for sl in SUB_PATTERNS:
        lines.append(
            f"- **{sl['id']}** — {sl['ja_title']}（{sl['kind']}）: {sl['ja_goal']}"
        )
    lines.append("")
    (out_dir / "INSTRUCTION_LOGIC.md").write_text("\n".join(lines), encoding="utf-8")


def _summary_for(slot: Dict[str, str], engine: str) -> str:
    return (
        f"【種別】サブ画像のみ（MAIN白抜き禁止）\n"
        f"【パターン】{slot['ja_title']}\n"
        f"【目的】{slot['ja_goal']}\n"
        f"【エンジン】{engine}\n"
        f"【背景】実写写真キー={slot.get('bg_key', '-')}\n"
        f"【禁止】商品だけのシンプルMAIN／OCR再描画\n"
    )


def run_poc(
    *,
    work_root: Path,
    own_path: Path,
    comp_paths: List[Path],
    product_name: str,
    model: Optional[str],
    skip_ai: bool,
) -> Path:
    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_root = work_root / "00.テスト出力" / "sub_image_poc" / run_id
    dir_in = out_root / "01_inputs"
    dir_raw = out_root / "02_raw"
    dir_annot = out_root / "03_annot"
    dir_photo = dir_annot / "写真背景"
    dir_mash = dir_annot / "競合掛け合わせ"
    dir_scene = dir_annot / "シーンAI"
    for d in (dir_in, dir_raw, dir_photo, dir_mash, dir_scene, meta_dir_for(out_root)):
        d.mkdir(parents=True, exist_ok=True)

    own_im = _load_rgba(own_path)
    comp_ims = [_load_rgba(p) for p in comp_paths]
    own_im.convert("RGB").save(dir_in / f"OWN_{own_path.name}")
    for i, (p, im) in enumerate(zip(comp_paths, comp_ims)):
        im.convert("RGB").save(dir_in / f"COMP{i+1}_{p.name}")

    # 背景コピー
    bg_cache: Dict[str, Path] = {}
    for sl in SUB_PATTERNS:
        key = sl.get("bg_key") or "bg_kitchen"
        bp = _bg_path(key)
        bg_cache[key] = bp
        dest = dir_in / f"BG_{key}{bp.suffix}"
        if not dest.is_file():
            Image.open(bp).convert("RGB").save(dest, quality=92)

    write_instruction_logic_md(out_root, product_name)

    client = None
    model_id = None
    if not skip_ai:
        client = make_client()
        model_id = resolve_model_id(client, explicit=model)
        LOG.info("model=%s", model_id)

    jobs: List[Dict[str, Any]] = []
    total = len(SUB_PATTERNS)

    for step, slot in enumerate(SUB_PATTERNS, start=1):
        kind = slot["kind"]
        bg_key = slot.get("bg_key") or "bg_kitchen"
        bg_path = bg_cache[bg_key]
        bg_im = _load_rgba(bg_path)
        stem = f"{step:02d}_{slot['id']}"
        raw_path = dir_raw / f"{stem}.jpg"

        if kind == "photo_pillow":
            engine = "Pillow実写背景合成"
            mode_id = "PHOTO_PILLOW"
            mode_ja = "写真背景"
            annot_dir = dir_photo
            result_im = _place_product_on_bg(bg_im, own_im, scale=0.44)
            summary = _summary_for(slot, engine)
            api_meta: Dict[str, Any] = {"skipped": True, "note": "pillow photo bg"}
            left_img = bg_im
            left_cap = "①実写背景"

        elif kind == "mash_pillow":
            engine = "Pillow競合パーツ掛け合わせ"
            mode_id = "COMP_MASH_PILLOW"
            mode_ja = "競合掛け合わせ"
            annot_dir = dir_mash
            result_im = compose_comp_mash(bg_im, own_im, comp_ims)
            summary = _summary_for(slot, engine)
            api_meta = {"skipped": True, "note": "pillow mash"}
            left_img = Image.new("RGBA", (SQUARE, SQUARE), (40, 40, 40, 255))
            # 入力プレビュー: 競合を並べた簡易板
            for i, cim in enumerate(comp_ims[:4]):
                th = cim.copy()
                th.thumbnail((280, 280), Image.Resampling.LANCZOS)
                left_img.alpha_composite(th, (20 + (i % 2) * 300, 20 + (i // 2) * 300))
            left_cap = "①競合パーツ元"

        elif kind == "mash_ai":
            engine = "AI競合掛け合わせ"
            mode_id = "COMP_MASH_AI"
            mode_ja = "競合掛け合わせAI"
            annot_dir = dir_mash
            summary = _summary_for(slot, engine)
            prompt = prompt_comp_mash_ai(product_name_hint=product_name)
            paths = [bg_path, own_path] + list(comp_paths[:2])
            roles = ["IMAGE_BG", "IMAGE_OWN", "IMAGE_COMP_A", "IMAGE_COMP_B"][: len(paths)]
            if skip_ai or client is None:
                result_im = compose_comp_mash(bg_im, own_im, comp_ims)
                api_meta = {"skipped": True, "note": "AIスキップ→pillow fallback"}
            else:
                data, api_meta = generate_with_references(
                    client=client,
                    model_id=model_id or "",
                    prompt=prompt,
                    image_paths=paths,
                    image_roles=roles,
                    aspect_ratio="1:1",
                    image_size="1K",
                )
                save_as_jpeg(data, raw_path)
                result_im = Image.open(raw_path).convert("RGBA")
            left_img = own_im
            left_cap = "①自社＋競合ヒント"

        else:  # scene_ai
            engine = "AIシーン（写真背景誘導）"
            mode_id = "SCENE_AI"
            mode_ja = "シーンAI"
            annot_dir = dir_scene
            summary = _summary_for(slot, engine)
            prompt = prompt_scene_ai(slot_id=slot["id"], product_name_hint=product_name)
            paths = [bg_path, own_path]
            roles = ["IMAGE_BG", "IMAGE_OWN"]
            if comp_paths:
                paths.append(comp_paths[0])
                roles.append("IMAGE_COMP_A")
            if skip_ai or client is None:
                result_im = _place_product_on_bg(bg_im, own_im, scale=0.4)
                api_meta = {"skipped": True, "note": "AIスキップ→pillow fallback"}
            else:
                data, api_meta = generate_with_references(
                    client=client,
                    model_id=model_id or "",
                    prompt=prompt,
                    image_paths=paths,
                    image_roles=roles,
                    aspect_ratio="1:1",
                    image_size="1K",
                )
                save_as_jpeg(data, raw_path)
                result_im = Image.open(raw_path).convert("RGBA")
            left_img = bg_im
            left_cap = "①実写背景"

        if not raw_path.is_file():
            result_im.convert("RGB").save(raw_path, quality=92)
            LOG.info("saved %s (%s)", raw_path.name, mode_id)
        else:
            LOG.info("saved %s", raw_path.name)

        annot_path = annot_dir / f"{stem}_annot.jpg"
        board = build_annot_board(
            mode_ja=mode_ja,
            mode_id=mode_id,
            slot=slot,
            step_index=step,
            step_total=total,
            prompt_summary_ja=summary,
            left_img=left_img,
            left_caption=left_cap,
            mid_img=None,
            mid_caption="②指示",
            right_img=result_im,
            right_caption="③出力",
        )
        board.save(annot_path, quality=92)

        # 入力比較: 背景|自社|出力
        trio = _build_trio(
            title=f"[{step}/{total}] {slot['ja_title']}",
            bg=bg_im,
            own=own_im,
            result=result_im,
            summary=summary,
        )
        trio.save(annot_dir / f"{stem}_trio.jpg", quality=92)

        jobs.append(
            {
                "step": step,
                "slot": slot["id"],
                "kind": kind,
                "mode": mode_id,
                "bgKey": bg_key,
                "raw": str(raw_path),
                "annot": str(annot_path),
                "api": {
                    "apiPath": api_meta.get("apiPath"),
                    "skipped": api_meta.get("skipped"),
                    "note": api_meta.get("note"),
                },
            }
        )

    _build_contact_sheet(dir_photo, out_root / "04_contact_写真背景.jpg")
    _build_contact_sheet(dir_mash, out_root / "04_contact_競合掛け合わせ.jpg")
    _build_contact_sheet(dir_scene, out_root / "04_contact_シーンAI.jpg")
    _build_contact_all_raw(dir_raw, out_root / "04_contact_all_raw.jpg")

    meta = {
        "runId": run_id,
        "productName": product_name,
        "own": str(own_path),
        "comps": [str(p) for p in comp_paths],
        "model": model_id,
        "skipAi": skip_ai,
        "jobs": jobs,
        "logic": logic_steps_for_humans(),
        "outRoot": str(out_root),
        "patterns": [s["id"] for s in SUB_PATTERNS],
        "notes": [
            "sub-images only (no MAIN packshot)",
            "real photo backgrounds in _sub_bg_photos",
            "competitor part mashup P05/P06",
        ],
    }
    (meta_dir_for(out_root) / "run_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    LOG.info("PoC done → %s", out_root)
    return out_root


def _build_trio(
    *,
    title: str,
    bg: Image.Image,
    own: Image.Image,
    result: Image.Image,
    summary: str,
) -> Image.Image:
    w, h = 2400, 1100
    board = Image.new("RGB", (w, h), (28, 30, 36))
    draw = ImageDraw.Draw(board)
    font_h = _font(32)
    font_s = _font(22)
    draw.text((20, 24), title, fill=(255, 210, 100), font=font_h)
    pw = w // 3
    items = [
        (bg, "実写背景", (70, 130, 180)),
        (own, "自社", (100, 100, 160)),
        (result, "出力サブ画像", (80, 170, 110)),
    ]
    for i, (im, cap, col) in enumerate(items):
        x0 = i * pw
        draw.rectangle([x0 + 8, 90, x0 + pw - 8, 90 + 40], fill=col)
        draw.text((x0 + 20, 96), cap, fill=(255, 255, 255), font=font_s)
        fitted = _fit(im, (0, 0, pw - 24, 720))
        board.paste(fitted.convert("RGB"), (x0 + 12, 140), fitted.split()[-1])
    lines = _wrap(draw, summary, font_s, w - 40)
    ty = 900
    for ln in lines[:5]:
        draw.text((20, ty), ln, fill=(210, 210, 220), font=font_s)
        ty += 28
    return board


def _build_contact_sheet(annot_dir: Path, out_path: Path) -> None:
    files = sorted(annot_dir.glob("*_annot.jpg"))
    if not files:
        return
    thumbs: List[Image.Image] = []
    for f in files:
        im = Image.open(f).convert("RGB")
        im.thumbnail((900, 360), Image.Resampling.LANCZOS)
        thumbs.append(im)
    cols = 2
    rows = (len(thumbs) + cols - 1) // cols
    tw, th = thumbs[0].size
    sheet = Image.new("RGB", (cols * tw + 20, rows * th + 20), (40, 40, 40))
    for i, im in enumerate(thumbs):
        r, c = divmod(i, cols)
        sheet.paste(im, (10 + c * tw, 10 + r * th))
    sheet.save(out_path, quality=90)


def _build_contact_all_raw(raw_dir: Path, out_path: Path) -> None:
    files = sorted(raw_dir.glob("*.jpg"))
    if not files:
        return
    thumbs = []
    for f in files:
        im = Image.open(f).convert("RGB")
        im.thumbnail((360, 360), Image.Resampling.LANCZOS)
        thumbs.append(im)
    cols = 5
    rows = (len(thumbs) + cols - 1) // cols
    cell = 370
    sheet = Image.new("RGB", (cols * cell + 20, rows * cell + 20), (30, 30, 30))
    for i, im in enumerate(thumbs):
        r, c = divmod(i, cols)
        sheet.paste(im, (10 + c * cell, 10 + r * cell))
    sheet.save(out_path, quality=90)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Amazon サブ画像 PoC（10パターン・写真背景＋掛け合わせ）")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--own", type=Path, default=None, help="自社白抜き")
    ap.add_argument("--comp", type=Path, action="append", default=None, help="競合参照（複数可）")
    ap.add_argument("--name", default="自社商品（缶飯系サンプル）", help="商品名ヒント")
    ap.add_argument("--model", default=None)
    ap.add_argument("--skip-ai", action="store_true", help="API無し（Pillowのみ＋注釈）")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    work = args.work_root or default_work_root()
    own = args.own or _default_own(work)
    comps = list(args.comp) if args.comp else _default_comps()
    if not own.is_file():
        LOG.error("自社画像なし: %s", own)
        return 2
    for c in comps:
        if not c.is_file():
            LOG.error("競合画像なし: %s", c)
            return 2
    if not BG_DIR.is_dir() or not list(BG_DIR.glob("bg_*.jpg")):
        LOG.error("実写背景フォルダが空です: %s", BG_DIR)
        return 2

    out = run_poc(
        work_root=work,
        own_path=own,
        comp_paths=comps,
        product_name=args.name,
        model=args.model,
        skip_ai=args.skip_ai,
    )
    print(str(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
