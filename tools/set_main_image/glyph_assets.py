# -*- coding: utf-8 -*-
"""
楽天金丸用・背景透過グリフ（Canva優先）。

Canva（95.CANVA数字・セット）構成:
  - digit_0..9.png … 1桁数字
  - pair_10.png 等 … 2桁まとまり
  - unitset_*.png … 単位(右上)+「セット」(下) 一体。左上は数字スペース
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PIL import Image, ImageDraw

from rakuten_badge import _load_font, _resize_font

LOG = logging.getLogger("set_main_image.glyphs")

# 単位文字 → unitset / 旧unit ファイル名候補
UNITSET_FILE_ALIASES: Dict[str, List[str]] = {
    "点": ["unitset_ten.png"],
    "個": ["unitset_ko.png"],
    "袋": ["unitset_fukuro.png", "unitset_fukuro_alt.png"],
    "缶": ["unitset_kan.png"],  # U+7F36
    "箱": ["unitset_hako.png"],
    "包": ["unitset_tsutsumi.png"],
    "本": ["unitset_hon.png"],
    "枚": ["unitset_mai.png"],
    "瓶": ["unitset_bin.png"],
    "食": ["unitset_shoku.png"],
    "粒": ["unitset_tsubu.png"],
}
# 見た目が似る別コードポイントも同一素材へ
UNITSET_FILE_ALIASES["\u7f50"] = list(UNITSET_FILE_ALIASES["缶"])  # U+7F50

DEFAULT_UNITSET_UNIT = "個"

UNIT_FILE_ALIASES: Dict[str, List[str]] = {
    "缶": ["unit_03_can_kan.png", "unit_kan.png", "text_kan.png", "缶.png"],
    "袋": ["unit_01_bag_fukuro.png", "text_fukuro.png", "袋.png"],
    "個": ["unit_02_piece_ko.png", "個.png"],
    "点": ["unit_04_item_ten.png", "点.png"],
    "包": ["unit_05_package_tsutsumi.png", "包.png"],
    "箱": ["unit_06_box_hako.png", "unit_07_carton_karton.png", "箱.png"],
    "束": ["unit_08_bundle_tabane.png", "束.png"],
    "ケース": ["unit_09_case_keesu.png", "ケース.png"],
    "粒": ["unit_10_grain_tsubu.png", "粒.png"],
    "本": ["本.png"],
    "枚": ["枚.png"],
    "瓶": ["瓶.png"],
    "食": ["食.png"],
}

SET_FILE_ALIASES = ["text_set.png", "セット.png", "set.png"]

DEFAULT_FILL = (38, 62, 37, 255)


def default_glyph_dirs(work_root: Optional[Path] = None) -> List[Path]:
    """Canva 95 → 96 → ローカル自動生成 の順で探す。"""
    here = Path(__file__).resolve().parent / "glyphs" / "rakuten"
    out: List[Path] = []
    if work_root and work_root.is_dir():
        for p in sorted(work_root.iterdir()):
            if p.is_dir() and p.name.startswith("95"):
                out.append(p)
        out.append(work_root / "96.楽天透過文字")
    out.append(here)
    return out


def local_auto_glyph_dir() -> Path:
    return Path(__file__).resolve().parent / "glyphs" / "rakuten"


def _glyph_avg_rgb(path: Path) -> Optional[Tuple[float, float, float]]:
    if not path.is_file():
        return None
    im = Image.open(path).convert("RGBA")
    opaque = [(r, g, b) for r, g, b, a in im.getdata() if a > 200]
    if not opaque:
        return None
    n = len(opaque)
    return (
        sum(c[0] for c in opaque) / n,
        sum(c[1] for c in opaque) / n,
        sum(c[2] for c in opaque) / n,
    )


def _fill_mismatch(path: Path, fill: Tuple[int, int, int, int], tol: int = 40) -> bool:
    avg = _glyph_avg_rgb(path)
    if avg is None:
        return True
    return any(abs(avg[i] - fill[i]) > tol for i in range(3))


def ensure_generated_glyphs(
    dest: Path,
    *,
    font_id: str = "tsukushi_mincho_like",
    fill: Tuple[int, int, int, int] = DEFAULT_FILL,
    canvas: int = 256,
    force: bool = False,
) -> Path:
    """ローカル自動生成用。Canvaフォルダには呼ばないこと。"""
    dest.mkdir(parents=True, exist_ok=True)
    sample = dest / "digit_1.png"
    # Canvaっぽい大サイズPNGは触らない
    if sample.is_file():
        try:
            with Image.open(sample) as im:
                if max(im.size) >= 800:
                    LOG.info("skip ensure (looks like Canva) dir=%s", dest)
                    return dest
        except Exception:
            pass
    need = force or (not sample.is_file()) or _fill_mismatch(sample, fill)
    for d in range(10):
        path = dest / f"digit_{d}.png"
        if need or not path.is_file():
            _render_glyph(str(d), path, font_id=font_id, fill=fill, canvas=canvas, stroke=3)
    for unit, names in UNIT_FILE_ALIASES.items():
        path = dest / names[0]
        if need or not path.is_file():
            _render_glyph(unit, path, font_id=font_id, fill=fill, canvas=canvas, stroke=3)
    set_path = dest / SET_FILE_ALIASES[0]
    if need or not set_path.is_file():
        _render_glyph("セット", set_path, font_id=font_id, fill=fill, canvas=canvas, stroke=3)
    LOG.info("glyphs ready dir=%s refreshed=%s", dest, need)
    return dest


def _render_glyph(
    text: str,
    out: Path,
    *,
    font_id: str,
    fill: Tuple[int, int, int, int],
    canvas: int,
    stroke: int,
) -> None:
    im = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
    draw = ImageDraw.Draw(im)
    target_h = int(canvas * 0.72)
    size = target_h
    font = _load_font(font_id, size)
    for _ in range(24):
        font = _resize_font(font, size, font_id)
        tb = draw.textbbox((0, 0), text, font=font)
        tw, th = tb[2] - tb[0], tb[3] - tb[1]
        if tw <= canvas * 0.88 and th <= canvas * 0.88:
            break
        size = max(10, int(size * 0.92))
    tb = draw.textbbox((0, 0), text, font=font)
    tw, th = tb[2] - tb[0], tb[3] - tb[1]
    x = (canvas - tw) // 2 - tb[0]
    y = (canvas - th) // 2 - tb[1]
    if stroke > 0:
        draw.text((x, y), text, font=font, fill=fill, stroke_width=stroke, stroke_fill=fill)
    else:
        draw.text((x, y), text, font=font, fill=fill)
    im.save(out, format="PNG")


def find_glyph(name_candidates: List[str], dirs: List[Path]) -> Optional[Path]:
    for d in dirs:
        if not d.is_dir():
            continue
        for name in name_candidates:
            p = d / name
            if p.is_file():
                return p
        lower_map = {p.name.lower(): p for p in d.glob("*.png")}
        for name in name_candidates:
            hit = lower_map.get(name.lower())
            if hit:
                return hit
    return None


def load_digit_glyph(digit: str, dirs: List[Path]) -> Image.Image:
    p = find_glyph([f"digit_{digit}.png", f"{digit}.png"], dirs)
    if not p:
        raise FileNotFoundError(f"digit glyph missing: {digit} in {dirs}")
    return Image.open(p).convert("RGBA")


def load_pair_glyph(n: int, dirs: List[Path]) -> Optional[Image.Image]:
    p = find_glyph([f"pair_{int(n)}.png"], dirs)
    if not p:
        return None
    return Image.open(p).convert("RGBA")


def load_unitset_glyph(unit: str, dirs: List[Path]) -> Optional[Image.Image]:
    u = (unit or "").strip()
    if not u:
        return None
    names = UNITSET_FILE_ALIASES.get(u, [f"unitset_{u}.png"])
    p = find_glyph(names, dirs)
    if not p:
        return None
    return Image.open(p).convert("RGBA")


def resolve_unit_for_canva(
    unit: str,
    dirs: List[Path],
    *,
    fallback: str = DEFAULT_UNITSET_UNIT,
) -> Tuple[str, str]:
    """
    マスタ単位 → Canva unitset がある単位。
    空・未知・素材なしは fallback（既定=個）。
    戻り値: (resolved_unit, reason)
    """
    raw = (unit or "").strip()
    if raw and load_unitset_glyph(raw, dirs) is not None:
        return raw, "master"
    if not raw:
        if load_unitset_glyph(fallback, dirs) is not None:
            return fallback, "empty_to_default"
        return fallback, "empty_to_default_missing_glyph"
    # 素材なし
    if load_unitset_glyph(fallback, dirs) is not None:
        LOG.warning(
            "unitset missing for unit=%r — fallback to %r", raw, fallback
        )
        return fallback, "missing_glyph_to_default"
    LOG.warning(
        "unitset missing for unit=%r and fallback=%r", raw, fallback
    )
    return fallback, "missing_both"


def load_unit_glyph(unit: str, dirs: List[Path]) -> Optional[Image.Image]:
    u = (unit or "").strip()
    if not u:
        return None
    names = UNIT_FILE_ALIASES.get(u, [f"{u}.png", f"unit_{u}.png"])
    p = find_glyph(names, dirs)
    if not p:
        LOG.warning("unit glyph missing unit=%r — skip unit paste", u)
        return None
    return _crop_to_alpha(Image.open(p).convert("RGBA"))


def load_set_glyph(dirs: List[Path]) -> Optional[Image.Image]:
    p = find_glyph(SET_FILE_ALIASES, dirs)
    if not p:
        return None
    return _crop_to_alpha(Image.open(p).convert("RGBA"))


def _crop_to_alpha(im: Image.Image, alpha_min: int = 16) -> Image.Image:
    im = im.convert("RGBA")
    alpha = im.split()[-1]
    bbox = alpha.point(lambda a: 255 if a >= alpha_min else 0).getbbox()
    if not bbox:
        return im
    return im.crop(bbox)


def _resize_to_height(im: Image.Image, height: int) -> Image.Image:
    if height < 1:
        height = 1
    w, h = im.size
    nw = max(1, int(w * (height / max(1, h))))
    return im.resize((nw, height), Image.Resampling.LANCZOS)


def _fit_in_box(im: Image.Image, box_w: int, box_h: int) -> Image.Image:
    """アスペクト維持で box に収める。"""
    im = _crop_to_alpha(im)
    w, h = im.size
    if w < 1 or h < 1:
        return im
    scale = min(box_w / w, box_h / h)
    nw = max(1, int(w * scale))
    nh = max(1, int(h * scale))
    return im.resize((nw, nh), Image.Resampling.LANCZOS)


def paste_glyph(base: Image.Image, glyph: Image.Image, xy: Tuple[int, int]) -> None:
    g = glyph.convert("RGBA")
    base.alpha_composite(g, dest=xy)


def split_unitset_parts(unitset: Image.Image) -> Tuple[Image.Image, Image.Image]:
    """unitset から単位(右上)とセット(下)を分離して返す。"""
    cropped = _crop_to_alpha(unitset.convert("RGBA"))
    w, h = cropped.size
    pix = cropped.load()
    unit_pts: List[Tuple[int, int]] = []
    set_pts: List[Tuple[int, int]] = []
    for y in range(h):
        for x in range(w):
            if pix[x, y][3] <= 20:
                continue
            if y < h * 0.55 and x >= w * 0.45:
                unit_pts.append((x, y))
            if y >= h * 0.55:
                set_pts.append((x, y))

    def crop_pts(pts: List[Tuple[int, int]]) -> Image.Image:
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        return cropped.crop((min(xs), min(ys), max(xs) + 1, max(ys) + 1))

    if not unit_pts or not set_pts:
        raise ValueError("failed to split unitset into unit/set parts")
    return crop_pts(unit_pts), crop_pts(set_pts)


def _paste_box1000(
    base: Image.Image,
    glyph: Image.Image,
    box: List[int],
    *,
    canvas: int,
    grid: float = 1000,
    fit: str = "stretch",
) -> Dict[str, int]:
    """
    box=[x,y,w,h] on grid.
    fit=stretch … 枠に合わせて歪ませる（非推奨・数字は使わない）
    fit=contain … 素材アスペクト維持で枠内に収める（中央）
    """
    scale = canvas / grid
    x = int(round(box[0] * scale))
    y = int(round(box[1] * scale))
    w = max(1, int(round(box[2] * scale)))
    h = max(1, int(round(box[3] * scale)))
    g = _crop_to_alpha(glyph.convert("RGBA"))
    mode = (fit or "stretch").lower()
    if mode == "contain":
        gw, gh = g.size
        s = min(w / max(1, gw), h / max(1, gh))
        nw = max(1, int(round(gw * s)))
        nh = max(1, int(round(gh * s)))
        g = g.resize((nw, nh), Image.Resampling.LANCZOS)
        ox = x + (w - nw) // 2
        oy = y + (h - nh) // 2
        paste_glyph(base, g, (ox, oy))
        return {"x": ox, "y": oy, "w": nw, "h": nh}
    g = g.resize((w, h), Image.Resampling.LANCZOS)
    paste_glyph(base, g, (x, y))
    return {"x": x, "y": y, "w": w, "h": h}


def compose_badge_with_glyphs(
    base: Image.Image,
    *,
    set_count: int,
    unit: str,
    cx: int,
    cy: int,
    diameter: int,
    typo: dict,
    glyph_dirs: List[Path],
    draw_set_label: bool = True,
    canvas: int = 1200,
    tint_rgba: Optional[Tuple[int, int, int, int]] = None,
) -> Tuple[Image.Image, dict]:
    """金丸はそのまま。lockedUnitSet があれば単位+セット固定→数字。"""
    from text_color import recolor_glyph_rgba

    resolved_unit, unit_reason = resolve_unit_for_canva(unit, glyph_dirs)
    if unit_reason != "master":
        LOG.info("unit resolve %r → %r (%s)", unit, resolved_unit, unit_reason)
    unit = resolved_unit

    locked = typo.get("lockedUnitSet1000") or {}
    digit_key = "1digit" if int(set_count) < 10 else "2digit"
    if locked.get("enabled"):
        unitset = load_unitset_glyph(unit, glyph_dirs)
        if unitset is not None:
            if tint_rgba:
                unitset = recolor_glyph_rgba(unitset, tint_rgba)
            work, meta = _compose_locked_unit_set(
                base,
                set_count=int(set_count),
                unit=(unit or "").strip(),
                unitset=unitset,
                typo=typo,
                glyph_dirs=glyph_dirs,
                canvas=canvas,
                digit_key=digit_key,
                tint_rgba=tint_rgba,
            )
            meta["unitResolved"] = unit
            meta["unitResolveReason"] = unit_reason
            meta["tintRgba"] = list(tint_rgba) if tint_rgba else None
            return work, meta
        LOG.warning("lockedUnitSet enabled but unitset missing unit=%r", unit)

    canva_cfg = typo.get("canvaUnitset") or {}
    if canva_cfg.get("enabled", True):
        unitset = load_unitset_glyph(unit, glyph_dirs)
        if unitset is not None:
            if tint_rgba:
                unitset = recolor_glyph_rgba(unitset, tint_rgba)
            work, meta = _compose_canva_unitset(
                base,
                set_count=int(set_count),
                unit=(unit or "").strip(),
                unitset=unitset,
                typo=typo,
                glyph_dirs=glyph_dirs,
                canvas=canvas,
            )
            meta["unitResolved"] = unit
            meta["unitResolveReason"] = unit_reason
            return work, meta
        LOG.warning("unitset missing for unit=%r — fallback to legacy glyphs", unit)

    work, meta = _compose_legacy_absolute(
        base,
        set_count=int(set_count),
        unit=(unit or "").strip(),
        cx=cx,
        cy=cy,
        diameter=diameter,
        typo=typo,
        glyph_dirs=glyph_dirs,
        draw_set_label=draw_set_label,
        canvas=canvas,
    )
    meta["unitResolved"] = unit
    meta["unitResolveReason"] = unit_reason
    return work, meta


def _paste_unitset_uniform1000(
    base: Image.Image,
    unitset: Image.Image,
    rect1000: List[int],
    *,
    canvas: int,
    grid: float = 1000,
    align: str = "top_left",
) -> Dict[str, int]:
    """単位+セット一体を素材アスペクトのまま rect に収めて貼る（歪めない）。"""
    scale = canvas / grid
    tx = int(round(rect1000[0] * scale))
    ty = int(round(rect1000[1] * scale))
    tw = max(1, int(round(rect1000[2] * scale)))
    th = max(1, int(round(rect1000[3] * scale)))
    cropped = _crop_to_alpha(unitset.convert("RGBA"))
    cw, ch = cropped.size
    fit = min(tw / max(1, cw), th / max(1, ch))
    nw = max(1, int(round(cw * fit)))
    nh = max(1, int(round(ch * fit)))
    a = (align or "top_left").lower()
    if a == "top_right":
        ux, uy = tx + tw - nw, ty
    elif a == "center":
        ux, uy = tx + (tw - nw) // 2, ty + (th - nh) // 2
    elif a == "bottom_right":
        ux, uy = tx + tw - nw, ty + th - nh
    else:
        ux, uy = tx, ty
    paste_glyph(base, cropped.resize((nw, nh), Image.Resampling.LANCZOS), (ux, uy))
    return {"x": ux, "y": uy, "w": nw, "h": nh}


def _compose_two_digits(
    glyph_dirs: List[Path],
    n: int,
    *,
    target_h_px: int,
    gap_ratio: float = 0.04,
) -> Tuple[Image.Image, str]:
    """
    1桁グリフ2つを等倍（高さ合わせ）で横並び。縦横比は変えない。
    target_h_px はキャンバス座標の高さ。
    """
    parts = [_crop_to_alpha(load_digit_glyph(ch, glyph_dirs)) for ch in str(int(n))]
    if len(parts) != 2:
        raise ValueError(f"expected 2 digits, got {n}")
    scaled: List[Image.Image] = []
    for p in parts:
        s = target_h_px / max(1, p.size[1])
        nw = max(1, int(round(p.size[0] * s)))
        nh = max(1, int(round(p.size[1] * s)))
        scaled.append(p.resize((nw, nh), Image.Resampling.LANCZOS))
    gap = max(1, int(round(target_h_px * gap_ratio)))
    tw = scaled[0].size[0] + gap + scaled[1].size[0]
    th = max(scaled[0].size[1], scaled[1].size[1])
    canvas_im = Image.new("RGBA", (tw, th), (0, 0, 0, 0))
    x = 0
    for i, p in enumerate(scaled):
        canvas_im.alpha_composite(p, (x, th - p.size[1]))  # 下端揃え
        x += p.size[0] + (gap if i == 0 else 0)
    return canvas_im, f"digits:{n}"


def _paste_right_bottom1000(
    base: Image.Image,
    glyph: Image.Image,
    box: List[int],
    *,
    canvas: int,
    grid: float = 1000,
) -> Dict[str, int]:
    """
    数字グループを numberBox の右縁・下縁に合わせて貼る（等倍・歪めない）。
    右側の桁の右縁＝box右縁、下端＝box下端。
    """
    scale = canvas / grid
    bx = int(round(box[0] * scale))
    by = int(round(box[1] * scale))
    bw = max(1, int(round(box[2] * scale)))
    bh = max(1, int(round(box[3] * scale)))
    g = _crop_to_alpha(glyph.convert("RGBA"))
    # 高さに合わせて等倍（幅は素材比）
    s = bh / max(1, g.size[1])
    nw = max(1, int(round(g.size[0] * s)))
    nh = max(1, int(round(g.size[1] * s)))
    g = g.resize((nw, nh), Image.Resampling.LANCZOS)
    # 右縁・下縁合わせ
    ox = bx + bw - nw
    oy = by + bh - nh
    paste_glyph(base, g, (ox, oy))
    return {"x": ox, "y": oy, "w": nw, "h": nh}


def _load_number_glyph(
    n: int,
    glyph_dirs: List[Path],
    *,
    gap_h: int,
    prefer_pair: bool = True,
    digit_gap_ratio: float = 0.04,
    target_h_canvas: Optional[int] = None,
) -> Tuple[Image.Image, str]:
    num_im: Optional[Image.Image] = None
    num_source = ""
    if prefer_pair and n >= 10:
        num_im = load_pair_glyph(n, glyph_dirs)
        if num_im is not None:
            return _crop_to_alpha(num_im), f"pair_{n}"
    if n >= 10:
        th = int(target_h_canvas or gap_h)
        return _compose_two_digits(
            glyph_dirs, n, target_h_px=max(1, th), gap_ratio=digit_gap_ratio
        )
    num_im = _crop_to_alpha(load_digit_glyph(str(n), glyph_dirs))
    return num_im, f"digits:{n}"


def _compose_locked_unit_set(
    base: Image.Image,
    *,
    set_count: int,
    unit: str,
    unitset: Image.Image,
    typo: dict,
    glyph_dirs: List[Path],
    canvas: int,
    digit_key: str,
    tint_rgba: Optional[Tuple[int, int, int, int]] = None,
) -> Tuple[Image.Image, dict]:
    """
    素材の単位+セットを一体のまま等倍縮尺で配置（袋だけ別縮尺にしない）。
    数字は numberBox に別置き。
    2桁 digit 組合せ時は numberBox の右縁・下縁に合わせる（pair版と同基準）。
    """
    from text_color import recolor_glyph_rgba

    work = base.convert("RGBA")
    cfg = (typo.get("lockedUnitSet1000") or {}).get(digit_key) or {}
    num_box = cfg.get("numberBox")
    rect = cfg.get("unitsetRect1000")

    if not rect:
        ub = cfg.get("unitBox") or [900, 58, 72, 87]
        sb = cfg.get("setBox") or [834, 140, 136, 28]
        x0 = min(ub[0], sb[0])
        y0 = min(ub[1], sb[1])
        x1 = max(ub[0] + ub[2], sb[0] + sb[2])
        y1 = max(ub[1] + ub[3], sb[1] + sb[3])
        rect = [x0, y0, x1 - x0, y1 - y0]

    align = str(cfg.get("align") or "top_left")
    unitset_pos = _paste_unitset_uniform1000(
        work, unitset, rect, canvas=canvas, align=align
    )

    n = int(set_count)
    num_pos = None
    num_source = ""
    aspect_lock = (typo.get("canvaAspectLock") or {}).get("enabled", True)
    num_fit = str(cfg.get("numberFitMode") or "contain").lower()
    if aspect_lock or num_fit == "stretch":
        num_fit = "contain"

    digit_cfg = typo.get("digitCompose2") or {}
    prefer_pair = bool(cfg.get("preferPair", digit_cfg.get("preferPair", True)))
    force_digits = bool(cfg.get("forceDigits", False))
    if force_digits:
        prefer_pair = False
    gap_ratio = float(digit_cfg.get("gapRatioOfHeight") or 0.04)
    align_rb = bool(digit_cfg.get("alignRightBottomToNumberBox", True))

    if num_box:
        scale = canvas / float(typo.get("addressGrid") or 1000)
        target_h = max(1, int(round(num_box[3] * scale)))
        num_im, num_source = _load_number_glyph(
            n,
            glyph_dirs,
            gap_h=int(num_box[3]),
            prefer_pair=prefer_pair,
            digit_gap_ratio=gap_ratio,
            target_h_canvas=target_h,
        )
        if tint_rgba:
            num_im = recolor_glyph_rgba(num_im, tint_rgba)
        # digit組合せ（pair無し）は右縁・下縁合わせ。pairは従来contain。
        use_rb = align_rb and num_source.startswith("digits:") and n >= 10
        if use_rb:
            num_pos = _paste_right_bottom1000(work, num_im, num_box, canvas=canvas)
            num_fit = "right_bottom"
        else:
            num_pos = _paste_box1000(work, num_im, num_box, canvas=canvas, fit=num_fit)

    return work, {
        "mode": "locked_unit_set_uniform",
        "digitKey": digit_key,
        "unit": unit,
        "numSource": num_source,
        "unitsetPos": unitset_pos,
        "numberPos": num_pos,
        "unitsetRect1000": list(rect),
        "numberBox1000": list(num_box) if num_box else None,
        "numberFitMode": num_fit,
        "align": align,
        "glyphDirs": [str(d) for d in glyph_dirs],
        "clearOldText": False,
    }


def _compose_canva_unitset(
    base: Image.Image,
    *,
    set_count: int,
    unit: str,
    unitset: Image.Image,
    typo: dict,
    glyph_dirs: List[Path],
    canvas: int,
) -> Tuple[Image.Image, dict]:
    """
    Canva設計:
      1) unitset（単位+セット一体・左上空洞）を金丸内に配置
      2) 空洞に digit_* または pair_* をはめ込む
    1桁 / 2桁で unitsetRect と numberHoleNorm が異なる。
    """
    work = base.convert("RGBA")
    digit_key = "1digit" if int(set_count) < 10 else "2digit"
    cfg = (typo.get("canvaUnitset") or {}).get(digit_key) or {}
    grid = float(typo.get("addressGrid") or 1000)
    scale = canvas / grid

    # 元1200キャンバス上の空洞（norm）→ クロップ後に再計算
    hole_src = cfg.get("numberHoleNorm") or [0.04, 0.14, 0.48, 0.42]
    src_w, src_h = unitset.size
    # インク外接でトリム（余白除去）
    cropped = _crop_to_alpha(unitset.convert("RGBA"))
    # 元画像に対するクロップ原点を推定（再オープンしてbbox）
    full = unitset.convert("RGBA")
    alpha = full.split()[-1]
    bb = alpha.point(lambda a: 255 if a >= 16 else 0).getbbox() or (0, 0, src_w, src_h)
    ox, oy = bb[0], bb[1]
    cw, ch = cropped.size

    # hole in cropped pixel space
    hx0 = hole_src[0] * src_w - ox
    hy0 = hole_src[1] * src_h - oy
    hw0 = hole_src[2] * src_w
    hh0 = hole_src[3] * src_h

    # unitsetRect1000: [x, y, w, h] 目標枠。アスペクト維持で収める。
    rect = cfg.get("unitsetRect1000") or [820, 18, 170, 160]
    tx = int(rect[0] * scale)
    ty = int(rect[1] * scale)
    tw = int(rect[2] * scale)
    th = int(rect[3] * scale)

    fit = min(tw / max(1, cw), th / max(1, ch))
    if cfg.get("allowStretch", False):
        # 見本外接に強制フィット（わずかな歪み許容）
        nw, nh = tw, th
        ux, uy = tx, ty
        sx_img, sy_img = tw / cw, th / ch
    else:
        nw = max(1, int(cw * fit))
        nh = max(1, int(ch * fit))
        align = (cfg.get("align") or "top_right").lower()
        if align == "top_right":
            ux = tx + tw - nw
            uy = ty
        elif align == "center":
            ux = tx + (tw - nw) // 2
            uy = ty + (th - nh) // 2
        else:
            ux, uy = tx, ty
        sx_img = sy_img = fit

    unitset_r = cropped.resize((nw, nh), Image.Resampling.LANCZOS)
    paste_glyph(work, unitset_r, (ux, uy))

    hx = ux + int(hx0 * sx_img)
    hy = uy + int(hy0 * sy_img)
    hw = max(1, int(hw0 * sx_img))
    hh = max(1, int(hh0 * sy_img))

    n = int(set_count)
    num_im: Optional[Image.Image] = None
    num_source = ""
    if n >= 10:
        num_im = load_pair_glyph(n, glyph_dirs)
        if num_im is not None:
            num_source = f"pair_{n}"
    if num_im is None:
        digits = str(n)
        parts = [_crop_to_alpha(load_digit_glyph(ch, glyph_dirs)) for ch in digits]
        if len(parts) == 1:
            num_im = parts[0]
        else:
            gap = max(1, int(hh * 0.02))
            total_w = sum(p.size[0] for p in parts) + gap * (len(parts) - 1)
            max_h = max(p.size[1] for p in parts)
            canvas_im = Image.new("RGBA", (total_w, max_h), (0, 0, 0, 0))
            x = 0
            for i, p in enumerate(parts):
                canvas_im.alpha_composite(p, (x, (max_h - p.size[1]) // 2))
                x += p.size[0] + (gap if i < len(parts) - 1 else 0)
            num_im = canvas_im
        num_source = f"digits:{n}"

    fitted = _fit_in_box(num_im, hw, hh)
    pad_x = int(cfg.get("numberPadLeftRatio", 0.06) * hw)
    nx = hx + pad_x
    ny = hy + (hh - fitted.size[1]) // 2
    nx = min(nx, hx + hw - fitted.size[0])
    nx = max(hx, nx)
    paste_glyph(work, fitted, (nx, ny))

    meta = {
        "mode": "canva_unitset",
        "digitKey": digit_key,
        "unit": unit,
        "numSource": num_source,
        "unitsetRect": {"x": ux, "y": uy, "w": nw, "h": nh},
        "targetRect": {"x": tx, "y": ty, "w": tw, "h": th},
        "numberHole": {"x": hx, "y": hy, "w": hw, "h": hh},
        "numberPos": {"x": nx, "y": ny, "w": fitted.size[0], "h": fitted.size[1]},
        "glyphDirs": [str(d) for d in glyph_dirs],
        "clearOldText": False,
    }
    return work, meta


def _compose_legacy_absolute(
    base: Image.Image,
    *,
    set_count: int,
    unit: str,
    cx: int,
    cy: int,
    diameter: int,
    typo: dict,
    glyph_dirs: List[Path],
    draw_set_label: bool,
    canvas: int,
) -> Tuple[Image.Image, dict]:
    work = base.convert("RGBA")
    digit_key = "1digit" if int(set_count) < 10 else "2digit"
    digits = str(int(set_count))

    addr = None
    if typo.get("useAbsoluteAddress"):
        grid = float(typo.get("addressGrid") or 1000)
        scale = canvas / grid
        raw = (typo.get("addresses1000") or {}).get(digit_key) or {}
        if raw.get("numberStart"):
            addr = {
                "numberStart": (
                    int(raw["numberStart"][0] * scale),
                    int(raw["numberStart"][1] * scale),
                ),
                "numberHeight": int(raw.get("numberHeight", 125) * scale),
                "unitStart": (
                    int(raw["unitStart"][0] * scale),
                    int(raw["unitStart"][1] * scale),
                )
                if raw.get("unitStart")
                else None,
                "unitHeight": int(raw.get("unitHeight", 55) * scale),
                "setStart": (
                    int(raw["setStart"][0] * scale),
                    int(raw["setStart"][1] * scale),
                )
                if raw.get("setStart")
                else None,
                "setHeight": int(raw.get("setHeight", 42) * scale),
                "setMaxWidth": int(raw["setMaxWidth"] * scale) if raw.get("setMaxWidth") else None,
            }

    num_cfg = (typo.get("number") or {}).get(digit_key) or {}
    unit_cfg = typo.get("unit") or {}
    set_cfg = typo.get("setLabel") or {}

    if addr:
        number_h = addr["numberHeight"]
    else:
        number_h = int(diameter * float(num_cfg.get("heightRatioOfDiameter", 0.50)))

    digit_ims = [
        _resize_to_height(_crop_to_alpha(load_digit_glyph(ch, glyph_dirs)), number_h)
        for ch in digits
    ]
    digit_gap = max(0, int(number_h * 0.02))
    num_w = sum(im.size[0] for im in digit_ims) + digit_gap * (len(digit_ims) - 1)
    num_h = max(im.size[1] for im in digit_ims)

    unit_im = None
    unit_w = unit_h = 0
    gap_nu = int(num_h * float(unit_cfg.get("gapAfterNumberRatioOfNumberHeight", 0.08)))
    if (unit or "").strip():
        raw_u = load_unit_glyph(unit, glyph_dirs)
        if raw_u is not None:
            unit_h = addr["unitHeight"] if addr else int(num_h * float(unit_cfg.get("heightRatioOfNumber", 0.40)))
            unit_im = _resize_to_height(raw_u, unit_h)
            unit_w = unit_im.size[0]

    set_im = None
    set_w = set_h = 0
    if draw_set_label:
        raw_set = load_set_glyph(glyph_dirs)
        if raw_set is not None:
            set_h = addr["setHeight"] if addr else int(num_h * float(set_cfg.get("heightRatioOfNumber", 0.24)))
            set_im = _resize_to_height(raw_set, set_h)
            set_w = set_im.size[0]
            max_sw = (addr or {}).get("setMaxWidth") if addr else None
            if max_sw and set_w > max_sw:
                set_im = set_im.resize((max_sw, set_h), Image.Resampling.LANCZOS)
                set_w = max_sw

    if addr:
        num_x, num_y = addr["numberStart"]
        x = num_x
        for i, dim in enumerate(digit_ims):
            paste_glyph(work, dim, (x, num_y))
            x += dim.size[0] + (digit_gap if i < len(digit_ims) - 1 else 0)
        unit_pos = None
        if unit_im is not None and addr.get("unitStart"):
            ux, uy = addr["unitStart"]
            ux = max(ux, num_x + num_w + max(2, gap_nu // 2))
            paste_glyph(work, unit_im, (ux, uy))
            unit_pos = {"x": ux, "y": uy, "w": unit_w, "h": unit_h}
        set_pos = None
        if set_im is not None and addr.get("setStart"):
            _sx, sy = addr["setStart"]
            if set_cfg.get("centerUnderNumberUnitGroup", True):
                group_l = num_x
                group_r = num_x + num_w
                if unit_pos:
                    group_r = max(group_r, unit_pos["x"] + unit_pos["w"])
                sx = int((group_l + group_r) / 2 - set_w / 2)
            else:
                sx = _sx
            sx = max(0, min(sx, work.size[0] - set_w))
            paste_glyph(work, set_im, (sx, sy))
            set_pos = {"x": sx, "y": sy, "w": set_w, "h": set_h}
        return work, {
            "mode": "glyph_alpha_overlay_absolute",
            "address1200": addr,
            "numberPos": {"x": num_x, "y": num_y, "w": num_w, "h": num_h},
            "unitPos": unit_pos,
            "setPos": set_pos,
            "glyphDirs": [str(d) for d in glyph_dirs],
            "clearOldText": False,
        }

    # center-block fallback
    gap_ns = int(diameter * float(set_cfg.get("gapBelowNumberRowRatioOfDiameter", 0.045)))
    row1_w = num_w + (gap_nu + unit_w if unit_im else 0)
    row1_h = max(num_h, unit_h or 0)
    total_h = row1_h + (gap_ns + set_h if set_im else 0)
    bias = float(typo.get("blockVerticalBias", -0.06))
    block_top = int(cy - total_h / 2 + bias * diameter)
    row1_left = int(cx - row1_w / 2)
    x = row1_left
    for i, dim in enumerate(digit_ims):
        paste_glyph(work, dim, (x, block_top + (row1_h - dim.size[1]) // 2))
        x += dim.size[0] + (digit_gap if i < len(digit_ims) - 1 else 0)
    if unit_im is not None:
        paste_glyph(
            work,
            unit_im,
            (row1_left + num_w + gap_nu, block_top + (row1_h - unit_im.size[1]) // 2),
        )
    if set_im is not None:
        paste_glyph(
            work,
            set_im,
            (int(cx - set_w / 2), block_top + row1_h + gap_ns),
        )
    return work, {"mode": "glyph_alpha_overlay_centered", "glyphDirs": [str(d) for d in glyph_dirs]}
