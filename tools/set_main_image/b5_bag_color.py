# -*- coding: utf-8 -*-
"""包材印刷面の色（色相・彩度・明度）を正本ファイルから画素で測り、指示と合否に使う。

細部（ロゴ一画素・文字詰め・窓内の中身）は数値化しない。印刷ベタ面だけ。
APIは使わない（指示用）。合否は bag_aspect の Gemini 呼び出しに printHueOk を載せる。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from PIL import Image

LOG = logging.getLogger("set_main_image.b5_bag_color")

_ROLE_OK = __import__("re").compile(r"^IMAGE_PACKAGE_(TRUTH|PART_\d+)$")


def _sat(r: float, g: float, b: float) -> float:
    mx, mn = max(r, g, b), min(r, g, b)
    if mx < 1.0:
        return 0.0
    return (mx - mn) / mx


def _hue_ja(r: float, g: float, b: float, sat: float) -> str:
    if sat < 0.12:
        return "無彩色に近い暗色"
    if b >= r + 15 and b >= g + 10:
        if (r + g + b) / 3 < 55:
            return "青系のごく濃い紺"
        return "青系の濃紺"
    if r >= g + 10 and r >= b + 10:
        return "赤〜茶系"
    if g >= r + 10 and g >= b + 10:
        return "緑系"
    if abs(r - g) < 15 and b < r - 10:
        return "黄〜茶系"
    return "有彩色（正本の色相を保て）"


def measure_print_face(path: Path) -> Optional[Dict[str, Any]]:
    """余白を除いた上側（印刷ベタ）の平均色。窓や白文字はできるだけ外す。"""
    p = Path(path)
    if not p.is_file():
        return None
    im = Image.open(p).convert("RGBA")
    w, h = im.size
    alpha = im.getchannel("A")
    hist = alpha.histogram()
    transparent = sum(hist[:17])
    if transparent >= w * h * 0.02:
        bb = alpha.point(lambda px: 255 if px > 16 else 0).getbbox()
    else:
        bb = (0, 0, w, h)
    if not bb:
        return None
    x0, y0, x1, y1 = bb
    bw, bh = x1 - x0, y1 - y0
    if bw < 16 or bh < 16:
        return None
    # 上 12%〜48%: 印刷面。下の窓を避け、上端ロゴ帯も少し避ける
    yy0 = y0 + int(bh * 0.12)
    yy1 = y0 + int(bh * 0.48)
    xx0 = x0 + int(bw * 0.20)
    xx1 = x0 + int(bw * 0.80)
    if yy1 <= yy0 + 4 or xx1 <= xx0 + 4:
        return None
    crop = im.crop((xx0, yy0, xx1, yy1))
    pix = list(crop.getdata())
    kept: List[tuple] = []
    for r, g, b, a in pix:
        if a < 32:
            continue
        L = (r + g + b) / 3.0
        if L > 200:  # 白文字・ハイライト
            continue
        if L < 8:  # 完全黒ノイズ
            continue
        kept.append((r, g, b))
    if len(kept) < 80:
        return None
    n = float(len(kept))
    mr = sum(t[0] for t in kept) / n
    mg = sum(t[1] for t in kept) / n
    mb = sum(t[2] for t in kept) / n
    sat = _sat(mr, mg, mb)
    L = (mr + mg + mb) / 3.0
    return {
        "path": str(p),
        "r": round(mr, 1),
        "g": round(mg, 1),
        "b": round(mb, 1),
        "sat": round(sat, 3),
        "L": round(L, 1),
        "blueDom": round(mb - mr, 1),
        "hueJa": _hue_ja(mr, mg, mb, sat),
        "sampleN": len(kept),
    }


def measure_role_colors(
    image_paths: Sequence[Path], image_roles: Sequence[str]
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for path, role in zip(image_paths, image_roles):
        if not _ROLE_OK.match(str(role or "")):
            continue
        met = measure_print_face(Path(path))
        if met is None:
            LOG.info("bag color skip role=%s", role)
            continue
        rec = {"role": str(role), **met}
        rows.append(rec)
        LOG.info(
            "bag color role=%s hue=%s sat=%s L=%s RGB=(%s,%s,%s)",
            role,
            met["hueJa"],
            met["sat"],
            met["L"],
            met["r"],
            met["g"],
            met["b"],
        )
    return rows


def color_prompt_block(rows: Sequence[Dict[str, Any]]) -> str:
    """指示用。数値は印刷ベタのみ。細部は正本画像を見ろ、と明記。"""
    if not rows:
        return ""
    lines = [
        "【包材の印刷色・ファイル実測（ベタ面のみ）】",
        "数字は正本ファイルの印刷ベタ（上側）の平均。ロゴ一画素・文字詰め・窓の中身は数値化しない。細部は添付の正本画像を見ろ。",
        "印刷面の色相と彩度はファイルのまま。黒・無彩色に落とすな。影の向きだけ場面に合わせる。白飛びさせるな。",
    ]
    for row in rows:
        lines.append(
            f"- {row['role']}: 色相={row['hueJa']}／彩度約{float(row['sat']):.2f}／"
            f"明度L約{float(row['L']):.0f}／RGB約("
            f"{float(row['r']):.0f},{float(row['g']):.0f},{float(row['b']):.0f})"
        )
        if float(row.get("blueDom") or 0) >= 25:
            lines.append(
                f"  ※青み（B-R約{float(row['blueDom']):.0f}）を残せ。黒やグレーに寄せるな。"
            )
    return "\n".join(lines) + "\n"
