# -*- coding: utf-8 -*-
"""Amazon素材: 白っぽい背景を透過（アルファ）にしてから合成・AI入力に使う。"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, Tuple

from PIL import Image

LOG = logging.getLogger("set_main_image.transparent")


def remove_near_white_bg(
    im: Image.Image,
    *,
    threshold: int = 245,
    soft: int = 12,
) -> Image.Image:
    """RGB/RGBA → RGBA。明るい背景を透明化（JPGフォールバック用）。"""
    base = im.convert("RGBA")
    lo = max(0, threshold - soft)
    span = max(1, threshold - lo)
    out_data = []
    for r, g, b, a in base.getdata():
        bright = r if r < g else g
        if b < bright:
            bright = b
        if bright >= threshold:
            out_data.append((r, g, b, 0))
        elif bright >= lo:
            t = (bright - lo) / span
            out_data.append((r, g, b, int(a * (1.0 - t))))
        else:
            out_data.append((r, g, b, a))
    out = Image.new("RGBA", base.size)
    out.putdata(out_data)
    return out


def inspect_alpha(path: Path, *, announce: bool = True) -> Dict[str, Any]:
    """
    保存画像の透過を検査する（2点）:
      ① PNG等にアルファ／透明画素があるか
      ② Canva「背景リムーバー」相当か（透明比率が十分か。キャンバス透過のみは不合格）
    """
    p = Path(path)
    info: Dict[str, Any] = {
        "path": str(p),
        "name": p.name,
        "suffix": p.suffix.lower(),
        "exists": p.is_file(),
        "hasAlphaChannel": False,
        "hasTransparency": False,
        "clearRatio": None,
        "nearWhiteOpaqueRatio": None,
        "canvaBgRemovedLikely": False,
        "matteOk": False,
        "alphaMin": None,
        "alphaMax": None,
        "clearPx": None,
        "opaquePx": None,
        "midAlphaPx": None,
        "mode": None,
        "check1Ja": "",
        "check2Ja": "",
        "messageJa": "",
    }
    if not p.is_file():
        info["messageJa"] = f"ファイルがありません: {p.name}"
        info["check1Ja"] = "① NG: ファイル無し"
        info["check2Ja"] = "② 判定不可"
        if announce:
            LOG.warning("%s", info["messageJa"])
            print(info["messageJa"])
        return info

    im = Image.open(p)
    info["mode"] = im.mode
    has_channel = im.mode in ("RGBA", "LA") or (
        im.mode == "P" and "transparency" in im.info
    )
    info["hasAlphaChannel"] = bool(has_channel)
    if not has_channel:
        info["check1Ja"] = (
            f"① NG: アルファ無し（mode={im.mode}）。透過PNGを置いてください。"
        )
        info["check2Ja"] = (
            "② NG: Canva背景リムーバー未確認（アルファが無いため判定不可）。"
            "Canvaで背景削除→PNG背景透過で書き出してください。"
        )
        info["messageJa"] = info["check1Ja"] + " / " + info["check2Ja"]
        if announce:
            LOG.info("%s", info["messageJa"])
            print(info["messageJa"])
        return info

    rgba = im.convert("RGBA")
    w, h = rgba.size
    total = max(1, w * h)
    hist = rgba.getchannel("A").histogram()
    a_min = next((i for i, c in enumerate(hist) if c), 0)
    a_max = next((i for i in range(255, -1, -1) if hist[i]), 0)
    clear_px = int(hist[0])
    opaque_px = int(hist[255])
    mid_px = int(sum(hist[1:255]))
    clear_ratio = clear_px / float(total)
    info["alphaMin"] = a_min
    info["alphaMax"] = a_max
    info["clearPx"] = clear_px
    info["opaquePx"] = opaque_px
    info["midAlphaPx"] = mid_px
    info["clearRatio"] = round(clear_ratio, 4)
    has_t = clear_px > 0 or a_min < 250
    info["hasTransparency"] = bool(has_t)

    # 不透明かつほぼ白の画素比率（白背景が焼き付いている指標）
    near_white_opaque = 0
    for r, g, b, a in rgba.getdata():
        if a < 250:
            continue
        if min(r, g, b) >= 245 and (max(r, g, b) - min(r, g, b)) <= 12:
            near_white_opaque += 1
    nw_ratio = near_white_opaque / float(total)
    info["nearWhiteOpaqueRatio"] = round(nw_ratio, 4)

    # ① 透過PNGか
    if has_t and p.suffix.lower() == ".png":
        info["check1Ja"] = (
            f"① OK: 透過PNG（透明px={clear_px}, 比率={clear_ratio:.1%}）"
        )
        check1_ok = True
    elif has_t:
        info["check1Ja"] = (
            f"① 注意: 透過はあるが拡張子がPNG以外（{p.suffix}）。PNG推奨。"
        )
        check1_ok = True
    else:
        info["check1Ja"] = (
            f"① NG: アルファはあるが実質不透明（alphaMin={a_min}）"
        )
        check1_ok = False

    # ② Canva背景リムーバー相当か
    # キャンバス透過のみだと透明比率がごく小さい（数%未満）ことが多い
    MIN_CLEAR_RATIO = 0.10
    MAX_NEAR_WHITE_OPAQUE = 0.08
    canva_ok = (
        check1_ok
        and clear_ratio >= MIN_CLEAR_RATIO
        and nw_ratio <= MAX_NEAR_WHITE_OPAQUE
    )
    info["canvaBgRemovedLikely"] = bool(canva_ok)
    if canva_ok:
        info["check2Ja"] = (
            f"② OK: Canva背景リムーバー実施の見込み"
            f"（透明比率 {clear_ratio:.1%} >= {MIN_CLEAR_RATIO:.0%} / "
            f"白不透明 {nw_ratio:.1%} <= {MAX_NEAR_WHITE_OPAQUE:.0%}）"
        )
    else:
        reasons = []
        if clear_ratio < MIN_CLEAR_RATIO:
            reasons.append(
                f"透明比率が低い {clear_ratio:.1%} < {MIN_CLEAR_RATIO:.0%}"
                "（キャンバス透過のみの可能性）"
            )
        if nw_ratio > MAX_NEAR_WHITE_OPAQUE:
            reasons.append(
                f"白い不透明画素が多い {nw_ratio:.1%} > {MAX_NEAR_WHITE_OPAQUE:.0%}"
                "（写真の白背景が残存）"
            )
        if not check1_ok:
            reasons.append("①がNG")
        info["check2Ja"] = (
            "② NG: Canva背景リムーバー未実施の可能性。理由: "
            + " / ".join(reasons)
            + "。Canvaで対象写真に背景削除→市松確認→PNG背景透過で再保存。"
        )

    info["matteOk"] = bool(check1_ok and canva_ok)
    info["messageJa"] = info["check1Ja"] + " | " + info["check2Ja"]
    if announce:
        LOG.info("%s", info["messageJa"])
        print(info["messageJa"])
    return info


def ensure_transparent_product(path: Path, cache_dir: Path | None = None) -> Path:
    """
    本線: 透過PNGがあればそのまま使う。
    JPG等でアルファが無いときだけ明るい背景を抜いてキャッシュする。
    """
    src = Path(path)
    check = inspect_alpha(src, announce=True)

    im = Image.open(src)
    use_existing = bool(check.get("hasTransparency"))
    if use_existing and src.suffix.lower() == ".png":
        return src

    cache = cache_dir or (src.parent / "_transparent_cache")
    cache.mkdir(parents=True, exist_ok=True)
    out = cache / (src.stem + "_alpha.png")
    if use_existing:
        im.convert("RGBA").save(out, format="PNG")
        LOG.info("[透過チェック] 既存アルファをキャッシュへ: %s", out.name)
        return out

    rgba = remove_near_white_bg(im)
    rgba.save(out, format="PNG")
    LOG.info(
        "wrote transparent %s from %s（JPG等フォールバック白抜き）",
        out.name,
        src.name,
    )
    print(
        f"[透過チェック] {src.name} → 自動白抜きで {out.name} を生成"
        "（本線は透過PNG推奨）"
    )
    return out


def paste_on_white(rgba: Image.Image, size: Tuple[int, int] | None = None) -> Image.Image:
    if size:
        rgba = rgba.resize(size, Image.Resampling.LANCZOS)
    bg = Image.new("RGB", rgba.size, (255, 255, 255))
    bg.paste(rgba, mask=rgba.split()[-1])
    return bg
