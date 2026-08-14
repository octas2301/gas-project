# -*- coding: utf-8 -*-
"""
日本語指示一覧と、各指示へのAI対処が視覚的に分かるボード画像を生成する。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFont

LOG = logging.getLogger("set_main_image.instruction_report")

# ユーザー向け・AI向けの日本語指示（提出用の正本）
JA_INSTRUCTIONS: List[Dict[str, str]] = [
    {
        "id": "I01",
        "title": "① 文字量ハード上限（スマホ可読）",
        "body": (
            "見出し全角12文字以内／本文最大3行／吹き出し・コールアウト最大2／"
            "栄養は3要点または別枠。微小文字の全面埋め禁止。"
        ),
    },
    {
        "id": "I02",
        "title": "② 心理プロセス順スロット5〜6×A/B（ページ分類）",
        "body": (
            "競合サブ画像をページ全体でテーマ①〜⑳に分類し、"
            "フック→プロダクト理解→ライフスタイル→信頼→クロージングの順でスロットを埋める。"
            "各スロット2案（A/B）。不足時は競合のphase被り、それでも不足なら想像（最大2・事実数字禁止）。"
            "商品名SEO語で軽微ブースト。1枚1主テーマ。正本PACKAGE_TRUTHは分類・OCRしない。"
        ),
    },
    {
        "id": "I03",
        "title": "③ トンマナはベージュ固定",
        "body": (
            "全画像でベージュ／サンド系の背景とカード色に統一。"
            "白・黒・グレー・ブランド色の切替はテスト段階では行わない。"
        ),
    },
    {
        "id": "I04",
        "title": "④ PACKAGE_LOCK＋写真実写＋外側のみAI≤50%",
        "body": (
            "IMAGE_PACKAGE_TRUTHの色相・ラベル・縦横比は改変不可（正本OCR禁止）。"
            "光・反射・距離に応じたピンボケ（複数被写体で必須）は photo_realism_rules に従う。"
            "外側パネルのみ競合OCR再配置可。面積の約半分まで新規レイアウト可。"
            "終了前チェックリスト必須。人間レビューでチェック＋要望→再生成。"
        ),
    },
]


def _find_jp_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\YuGothM.ttc"),
        Path(r"C:\Windows\Fonts\YuGothR.ttc"),
        Path(r"C:\Windows\Fonts\meiryo.ttc"),
        Path(r"C:\Windows\Fonts\msgothic.ttc"),
        Path("/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"),
        Path("/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"),
    ]
    for p in candidates:
        if p.is_file():
            try:
                return ImageFont.truetype(str(p), size=size)
            except Exception:
                continue
    return ImageFont.load_default()


def _wrap(draw: ImageDraw.ImageDraw, text: str, font, max_w: int) -> List[str]:
    lines: List[str] = []
    for para in text.split("\n"):
        if not para:
            lines.append("")
            continue
        cur = ""
        for ch in para:
            trial = cur + ch
            if draw.textlength(trial, font=font) <= max_w:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = ch
        if cur:
            lines.append(cur)
    return lines


def _thumb(path: Optional[Path], size: Tuple[int, int]) -> Image.Image:
    w, h = size
    if path and Path(path).is_file():
        im = Image.open(path).convert("RGB")
        im.thumbnail((w, h), Image.Resampling.LANCZOS)
        canvas = Image.new("RGB", (w, h), (245, 240, 230))
        canvas.paste(im, ((w - im.width) // 2, (h - im.height) // 2))
        return canvas
    blank = Image.new("RGB", (w, h), (60, 60, 60))
    d = ImageDraw.Draw(blank)
    f = _find_jp_font(22)
    d.text((20, h // 2 - 12), "画像なし", fill=(200, 200, 200), font=f)
    return blank


def build_instruction_list_board(
    *,
    out_path: Path,
    product_name: str,
    jan: str,
    theme_lines: Sequence[str],
) -> Path:
    """指示①〜④の一覧＋選択テーマを1枚にまとめる。"""
    W, H = 1600, 2200
    bg = (250, 245, 235)
    im = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(im)
    title_f = _find_jp_font(40)
    h_f = _find_jp_font(28)
    b_f = _find_jp_font(22)
    s_f = _find_jp_font(18)

    y = 40
    d.text((48, y), "サブ画像AI — 日本語指示一覧（テスト）", fill=(40, 30, 20), font=title_f)
    y += 56
    d.text((48, y), f"{product_name}  /  JAN {jan}", fill=(80, 60, 40), font=h_f)
    y += 48
    d.line((48, y, W - 48, y), fill=(180, 160, 130), width=2)
    y += 24

    for inst in JA_INSTRUCTIONS:
        d.rectangle((40, y, W - 40, y + 8), fill=(210, 180, 140))
        y += 20
        d.text((56, y), inst["title"], fill=(50, 35, 20), font=h_f)
        y += 40
        for line in _wrap(d, inst["body"], b_f, W - 120):
            d.text((56, y), line, fill=(55, 45, 35), font=b_f)
            y += 30
        y += 16

    y += 8
    d.text((48, y), "選択テーマ（競合に実在したもののみ）", fill=(40, 30, 20), font=h_f)
    y += 40
    for line in theme_lines:
        for wl in _wrap(d, line, s_f, W - 120):
            d.text((56, y), wl, fill=(60, 50, 40), font=s_f)
            y += 26
        if y > H - 80:
            break

    y = max(y + 20, H - 70)
    d.text(
        (48, H - 48),
        "トンマナ: ベージュ固定 ／ モデル比較: Gemini + OpenAI",
        fill=(100, 80, 60),
        font=s_f,
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    im.save(out_path, quality=92)
    return out_path


def build_instruction_response_board(
    *,
    out_path: Path,
    product_name: str,
    jan: str,
    rows: List[Dict[str, Any]],
    provider_cols: Optional[Sequence[Tuple[str, str]]] = None,
) -> Path:
    """
    各指示に対するAI対処の視覚ボード。
    rows: {title, note, ref, gemini?, fal?, openai?}
    provider_cols: [(key, label), ...] 既定は gemini / fal / openai のうち行に存在する列
    """
    if provider_cols is None:
        candidates = [("gemini", "Gemini"), ("fal", "fal.ai"), ("openai", "OpenAI")]
        provider_cols = []
        for key, lab in candidates:
            if any(r.get(key) for r in rows):
                provider_cols.append((key, lab))
        if not provider_cols:
            provider_cols = [("gemini", "Gemini"), ("openai", "OpenAI")]

    thumb_w, thumb_h = 320, 320
    pad = 24
    n_prov = len(provider_cols)
    row_h = thumb_h + 120
    W = pad * 2 + 280 + (thumb_w + 16) * (1 + n_prov)
    H = 100 + row_h * max(1, len(rows)) + 40
    bg = (36, 34, 32)
    im = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(im)
    title_f = _find_jp_font(32)
    h_f = _find_jp_font(20)
    s_f = _find_jp_font(16)

    col_names = " / ".join(lab for _, lab in provider_cols)
    d.text((pad, 24), f"指示への対処ビジュアル — {product_name}", fill=(245, 235, 220), font=title_f)
    d.text((pad, 64), f"JAN {jan} ｜ 競合参照 + {col_names}", fill=(180, 170, 150), font=s_f)

    y = 100
    labels = [("ref", "競合参照")] + list(provider_cols)
    labels_x = [pad + 280 + i * (thumb_w + 16) for i in range(len(labels))]
    for (key, lab), x in zip(labels, labels_x):
        d.text((x, y - 28), lab, fill=(200, 190, 170), font=s_f)

    for row in rows:
        d.rectangle((pad, y, pad + 260, y + thumb_h), fill=(55, 48, 40))
        ty = y + 16
        for line in _wrap(d, str(row.get("title") or ""), h_f, 230):
            d.text((pad + 12, ty), line, fill=(250, 240, 220), font=h_f)
            ty += 26
        ty += 8
        for line in _wrap(d, str(row.get("note") or ""), s_f, 230):
            d.text((pad + 12, ty), line, fill=(200, 185, 160), font=s_f)
            ty += 22

        for i, (key, _lab) in enumerate(labels):
            path = row.get(key)
            tip = _thumb(Path(path) if path else None, (thumb_w, thumb_h))
            im.paste(tip, (labels_x[i], y))
        y += row_h

    out_path.parent.mkdir(parents=True, exist_ok=True)
    im.save(out_path, quality=90)
    LOG.info("wrote response board %s", out_path)
    return out_path


def build_theme_contact(
    *,
    out_path: Path,
    jobs: List[Dict[str, Any]],
    provider_dirs: Optional[Sequence[Tuple[str, Path]]] = None,
    gemini_dir: Optional[Path] = None,
    openai_dir: Optional[Path] = None,
) -> Path:
    """テーマラベル付き連絡帳。provider_dirs=[(label, dir), ...]。"""
    if provider_dirs is None:
        provider_dirs = []
        if gemini_dir is not None:
            provider_dirs.append(("Gemini", Path(gemini_dir)))
        if openai_dir is not None:
            provider_dirs.append(("OpenAI", Path(openai_dir)))
    if not provider_dirs:
        raise ValueError("provider_dirs が空です")

    cell = 280
    cols = 5
    band_h = 2 * (cell + 36) + 28
    W = cols * cell + 40
    H = 60 + band_h * len(provider_dirs) + 20
    im = Image.new("RGB", (W, H), (30, 28, 26))
    d = ImageDraw.Draw(im)
    f = _find_jp_font(18)
    tf = _find_jp_font(26)
    names = " / ".join(lab for lab, _ in provider_dirs)
    d.text((20, 16), f"テーマ別10枚 — {names}", fill=(240, 230, 210), font=tf)

    def paste_row(start_y: int, folder: Path, label: str) -> None:
        d.text((20, start_y - 22), label, fill=(200, 190, 170), font=f)
        for i, job in enumerate(jobs[:12]):
            r, c = divmod(i, cols)
            if r > 1:
                break
            p = folder / f"{job['stem']}.jpg"
            thumb = _thumb(p if p.is_file() else None, (cell - 10, cell - 10))
            x = 20 + c * cell
            y = start_y + r * (cell + 36)
            im.paste(thumb, (x, y))
            caption = f"{job['stem']} {job['themeName'][:10]}"
            d.text((x, y + cell - 8), caption, fill=(210, 200, 180), font=f)

    y0 = 70
    for lab, folder in provider_dirs:
        paste_row(y0, Path(folder), lab)
        y0 += band_h
    out_path.parent.mkdir(parents=True, exist_ok=True)
    im.save(out_path, quality=90)
    return out_path
