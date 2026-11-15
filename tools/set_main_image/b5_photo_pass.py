# -*- coding: utf-8 -*-
"""金丸前の写真感2回目。none / package / whole。"""
from __future__ import annotations

import logging
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional

from PIL import Image

from openai_image import generate_with_references
from photo_realism_rules import photo_realism_block_ja

LOG = logging.getLogger("set_main_image.b5_photo_pass")

CANVAS = 1200
MODES = ("none", "package", "whole")
COMPARE_MODES = ("none", "package", "whole")
MODE_JA = {"none": "なし", "package": "袋だけ", "whole": "全体"}


def resolve_photo_modes(flag: Optional[str]) -> List[str]:
    raw = str(flag or "none").strip().lower()
    if raw in ("compare", "3", "ab"):
        return list(COMPARE_MODES)
    if raw in MODES:
        return [raw]
    return ["none"]


def photo_pass_prompt(mode: str) -> str:
    lock = (
        "袋の絵柄・印字・個品の形と色・後載せ文字は変えるな。"
        "構図・皿の数・文字の文言は維持。"
    )
    if mode == "package":
        return (
            "完成画の袋だけを、写真に置いたように整えてください。\n"
            f"{lock}\n"
            "対象は袋の縁・接地影・袋まわりの粒感だけ。\n"
            "皿・食品・文字・背景の明るさや形は触るな。\n"
            "切り抜きパスのような硬い袋縁をやめ、接地させる。\n"
            f"{photo_realism_block_ja()}\n"
            "ただし袋以外は加工するな。"
        )
    if mode == "whole":
        return (
            "完成画を、撮った写真のように整えてください。\n"
            f"{lock}\n"
            "縁の硬さ・接地・粒感・距離のボケだけ整えろ。\n"
            f"{photo_realism_block_ja()}\n"
        )
    return ""


def apply_photo_pass(
    im: Image.Image,
    *,
    mode: str,
    package_truth: Optional[Path] = None,
    openai_quality: str = "medium",
    work_tmp: Optional[Path] = None,
) -> Dict[str, Any]:
    """戻り: {im, mode, skipped, error}。none は入力をそのまま。"""
    m = str(mode or "none").strip().lower()
    base = im.convert("RGBA")
    if m == "none" or m not in ("package", "whole"):
        return {"im": base, "mode": "none", "skipped": True, "error": ""}
    tmp = Path(work_tmp) if work_tmp else Path(".")
    tmp.mkdir(parents=True, exist_ok=True)
    src = tmp / f"_photo_pass_src_{m}.png"
    base.save(src)
    paths = [src]
    if m == "package" and package_truth and Path(package_truth).is_file():
        paths.append(Path(package_truth))
    prompt = photo_pass_prompt(m)
    try:
        raw, model = generate_with_references(
            prompt=prompt,
            image_paths=paths,
            quality=openai_quality,
        )
        out = Image.open(BytesIO(raw))
        if out.mode != "RGBA":
            out = out.convert("RGBA")
        if out.size != (CANVAS, CANVAS):
            out = out.resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
        LOG.info("photo pass mode=%s model=%s", m, model)
        return {"im": out, "mode": m, "skipped": False, "error": "", "model": model}
    except Exception as e:
        LOG.warning("photo pass fail mode=%s: %s", m, e)
        return {"im": base, "mode": m, "skipped": True, "error": str(e)}


def write_compare_html(
    out_dir: Path,
    *,
    rows: List[Dict[str, Any]],
    title: str = "写真感 3枚比較",
) -> Path:
    cards = []
    for r in rows:
        src = str(r.get("rel") or r.get("file") or "")
        lab = str(r.get("label") or r.get("mode") or "")
        cards.append(
            f"<figure><img src='{src}' alt='{lab}'>"
            f"<figcaption>{lab}</figcaption></figure>"
        )
    html = (
        "<!doctype html><html lang=ja><meta charset=utf-8>"
        f"<title>{title}</title>"
        "<style>body{font-family:sans-serif;margin:24px;background:#f4f1ea}"
        "h1{font-size:20px}.row{display:flex;gap:20px;flex-wrap:wrap}"
        "figure{margin:0;background:#fff;padding:10px;border-radius:8px}"
        "img{max-width:min(92vw,420px);height:auto;display:block}"
        "figcaption{margin-top:8px;font-size:14px}</style>"
        f"<h1>{title}</h1><div class=row>{''.join(cards)}</div>"
    )
    dest = Path(out_dir) / "photo_pass_compare.html"
    dest.write_text(html, encoding="utf-8")
    return dest
