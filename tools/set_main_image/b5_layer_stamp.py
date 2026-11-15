# -*- coding: utf-8 -*-
"""B-④ MAIN: 袋・個品は正本画素を後載せ（描き直ししない）。

- 縦横比は正本のまま（stretch 禁止）
- 同じ段の複数袋／複数個品は同一縮尺（片方だけ縮小しない）
- 組み合わせ袋があればそれを1枚載せる（構成袋を別縮尺で並べない）
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFilter

from transparent_bg import inspect_alpha, remove_near_white_bg

LOG = logging.getLogger("set_main_image.b5_layer_stamp")

CANVAS = 1200
# 上＝キャッチ、右上＝金丸、右下＝Octas
CATCH_BOX = (20, 20, 820, 250)
BAG_BOX = (24, 640, 560, 1168)
FOOD_BOX = (360, 270, 1080, 820)
GOLD_BOX = (900, 0, 1200, 300)
OCTAS_BOX = (980, 980, 1200, 1200)
N1_BAG_BOX = (20, 520, 700, 1168)
N1_FOOD_BOX = (520, 280, 1100, 760)


def stamp_layers_enabled(flag: Optional[bool] = None) -> bool:
    if flag is not None:
        return bool(flag)
    import os

    raw = str(os.environ.get("B5_STAMP_TRUTH_LAYERS", "0") or "0").strip().lower()
    return raw not in ("0", "false", "no", "off")


def _as_box(box: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
    x0, y0, x1, y1 = (int(v) for v in box)
    return min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)


def _box_wh(box: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
    x0, y0, x1, y1 = _as_box(box)
    return x0, y0, max(1, x1 - x0), max(1, y1 - y0)


def cutout_truth(path: Path) -> Image.Image:
    """正本／見本を透過切り抜き。縦横比は変えない。"""
    src = Path(path)
    info = inspect_alpha(src, announce=False)
    im = Image.open(src).convert("RGBA")
    if not info.get("canvaBgRemovedLikely"):
        im = remove_near_white_bg(im)
    bb = im.split()[-1].getbbox()
    if bb:
        im = im.crop(bb)
    return im


def uniform_scale_factor(
    sizes: Sequence[Tuple[int, int]],
    box_w: int,
    box_h: int,
    *,
    gap: int = 20,
) -> float:
    """複数枚を横一列に収める共通縮尺。枚ごとに別 thumbnail しない。"""
    clean = [(max(1, int(w)), max(1, int(h))) for w, h in sizes if int(w) > 0 and int(h) > 0]
    if not clean:
        return 1.0
    n = len(clean)
    sum_w = sum(w for w, _h in clean)
    max_h = max(h for _w, h in clean)
    avail_w = max(1, int(box_w) - int(gap) * max(0, n - 1))
    return min(float(avail_w) / float(sum_w), float(box_h) / float(max_h))


def _scale_keep_aspect(im: Image.Image, scale: float) -> Image.Image:
    nw = max(1, int(round(im.size[0] * scale)))
    nh = max(1, int(round(im.size[1] * scale)))
    return im.resize((nw, nh), Image.Resampling.LANCZOS)


def _rotate2d(im: Image.Image, deg: float) -> Image.Image:
    if abs(float(deg)) < 0.05:
        return im
    return im.rotate(float(deg), expand=True, resample=Image.Resampling.BICUBIC)


def _paste_shadow(
    canvas: Image.Image,
    layer: Image.Image,
    xy: Tuple[int, int],
    *,
    blur: int = 10,
    alpha: int = 80,
    ox: int = 8,
    oy: int = 10,
) -> Image.Image:
    x, y = xy
    rw, rh = layer.size
    shadow = Image.new("RGBA", (CANVAS, CANVAS), (0, 0, 0, 0))
    sh = Image.new("RGBA", (rw, rh), (0, 0, 0, alpha))
    sh.putalpha(layer.split()[-1].point(lambda a: min(alpha, a)))
    shadow.paste(sh, (x + ox, y + oy), sh)
    shadow = shadow.filter(ImageFilter.GaussianBlur(blur))
    out = Image.alpha_composite(canvas.convert("RGBA"), shadow)
    out.alpha_composite(layer, (x, y))
    return out


def _place_row(
    canvas: Image.Image,
    cutouts: Sequence[Image.Image],
    box: Tuple[int, int, int, int],
    *,
    tilts: Sequence[float],
    gap: int,
    role: str,
) -> Tuple[Image.Image, Dict[str, Any]]:
    x0, y0, bw, bh = _box_wh(box)
    sizes = [im.size for im in cutouts]
    scale = uniform_scale_factor(sizes, bw, bh, gap=gap)
    layers: List[Image.Image] = []
    for i, im in enumerate(cutouts):
        deg = float(tilts[i]) if i < len(tilts) else 0.0
        layers.append(_rotate2d(_scale_keep_aspect(im, scale), deg))
    total_w = sum(im.size[0] for im in layers) + gap * max(0, len(layers) - 1)
    max_h = max((im.size[1] for im in layers), default=1)
    x = x0 + max(0, (bw - total_w) // 2)
    if role == "bag":
        x = x0  # 袋は左寄せ
    y_base = y0 + max(0, bh - max_h)  # 下端揃え
    out = canvas.convert("RGBA")
    placed: List[Dict[str, Any]] = []
    for im in layers:
        y = y_base + (max_h - im.size[1])
        x_i = max(0, min(CANVAS - im.size[0], x))
        y_i = max(0, min(CANVAS - im.size[1], y))
        out = _paste_shadow(out, im, (x_i, y_i))
        placed.append(
            {
                "x": x_i,
                "y": y_i,
                "w": im.size[0],
                "h": im.size[1],
                "srcW": im.size[0],
                "srcH": im.size[1],
            }
        )
        x += im.size[0] + gap
    return out, {"scale": scale, "count": len(layers), "placed": placed, "role": role}


def write_zone_guide(
    dest: Path,
    *,
    bag_box: Tuple[int, int, int, int] = BAG_BOX,
    food_box: Tuple[int, int, int, int] = FOOD_BOX,
) -> Path:
    """AI用。袋と個品の空きを示す。商品は描かない。"""
    im = Image.new("RGB", (CANVAS, CANVAS), (232, 220, 204))
    d = ImageDraw.Draw(im)
    for box, label in (
        (CATCH_BOX, "キャッチ"),
        (food_box, "皿は空ける（個品は後載せ）"),
        (bag_box, "袋は空ける（後載せ）"),
        (GOLD_BOX, ""),
        (OCTAS_BOX, ""),
    ):
        x0, y0, x1, y1 = _as_box(box)
        d.rectangle((x0, y0, x1, y1), outline=(176, 156, 136), width=4)
        if label:
            d.text((x0 + 12, y0 + 12), label, fill=(90, 70, 55))
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    im.save(dest, quality=90)
    return dest


def stamp_truth_layers(
    base: Image.Image,
    *,
    bag_paths: Sequence[Path],
    food_paths: Sequence[Path],
    theme_id: str = "",
    bag_mode: str = "single",
) -> Tuple[Image.Image, Dict[str, Any]]:
    """完成画へ袋・個品を画素のまま載せる。"""
    tid = str(theme_id or "").strip()
    bag_box = N1_BAG_BOX if tid == "N1" else BAG_BOX
    food_box = N1_FOOD_BOX if tid == "N1" else FOOD_BOX
    bags = [Path(p) for p in bag_paths if p and Path(p).is_file()]
    foods = [Path(p) for p in food_paths if p and Path(p).is_file()]
    # 組み合わせ袋が1枚なら、構成袋を並べない（実物の大小はその写真が正）
    if bag_mode == "mix_parent" and bags:
        bags = bags[:1]
    meta: Dict[str, Any] = {
        "themeId": tid,
        "bagMode": bag_mode,
        "bagPaths": [str(p) for p in bags],
        "foodPaths": [str(p) for p in foods],
        "bagBox": list(bag_box),
        "foodBox": list(food_box),
    }
    canvas = base.convert("RGBA")
    if canvas.size != (CANVAS, CANVAS):
        canvas = canvas.resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
    if foods:
        cuts = [cutout_truth(p) for p in foods]
        n = len(cuts)
        tilts = [(-4.0 if i % 2 == 0 else 4.0) for i in range(n)]
        if n == 1:
            tilts = [0.0]
        canvas, food_meta = _place_row(
            canvas, cuts, food_box, tilts=tilts, gap=28, role="food"
        )
        meta["food"] = food_meta
        LOG.info(
            "stamp food n=%s scale=%.4f paths=%s",
            food_meta.get("count"),
            float(food_meta.get("scale") or 0),
            [p.name for p in foods],
        )
    if bags:
        cuts = [cutout_truth(p) for p in bags]
        n = len(cuts)
        tilts = [(-8.0 if i == 0 else 8.0) for i in range(n)]
        if n == 1:
            tilts = [-8.0]
        canvas, bag_meta = _place_row(
            canvas, cuts, bag_box, tilts=tilts, gap=16, role="bag"
        )
        meta["bag"] = bag_meta
        LOG.info(
            "stamp bag n=%s scale=%.4f mode=%s paths=%s",
            bag_meta.get("count"),
            float(bag_meta.get("scale") or 0),
            bag_mode,
            [p.name for p in bags],
        )
    return canvas, meta


def collect_stamp_assets(
    *,
    parent_bag: Path,
    mix_parts: Optional[Sequence[Dict[str, Any]]] = None,
    identity: Optional[Path] = None,
) -> Dict[str, Any]:
    """後載せ対象。判定は商品名ベースの構成リスト。"""
    from food_unit_canon import load_part_bundle

    truth = Path(parent_bag) if parent_bag else Path()
    parts = [
        p
        for p in (mix_parts or [])
        if isinstance(p, dict) and str(p.get("nameBase") or p.get("name") or "").strip()
    ]
    bags: List[Path] = []
    foods: List[Path] = []
    bag_mode = "single"
    if len(parts) >= 2:
        if truth.is_file():
            bags = [truth]
            bag_mode = "mix_parent"
        else:
            bag_mode = "mix_parts"
        for raw in parts:
            bundle = load_part_bundle(
                name_base=str(raw.get("nameBase") or raw.get("name") or "").strip(),
                jan_for_folder=str(raw.get("jan") or ""),
                parent_sku=str(raw.get("parentSku") or ""),
                bag_path=str(raw.get("packageTruth") or raw.get("bag") or ""),
            )
            if bag_mode == "mix_parts":
                bag = bundle.get("bag")
                if bag and Path(bag).is_file():
                    bags.append(Path(bag))
            ident = bundle.get("identity")
            if ident and Path(ident).is_file():
                foods.append(Path(ident))
    else:
        if truth.is_file():
            bags = [truth]
        if identity and Path(identity).is_file():
            if not truth.is_file() or Path(identity).resolve() != truth.resolve():
                foods.append(Path(identity))
    plan = {
        "bagMode": bag_mode,
        "bags": bags,
        "foods": foods,
        "bagN": len(bags),
        "foodN": len(foods),
    }
    LOG.info(
        "stamp plan mode=%s bags=%s foods=%s",
        bag_mode,
        [p.name for p in bags],
        [p.name for p in foods],
    )
    return plan


def dump_stamp_meta(path: Path, meta: Dict[str, Any]) -> None:
    Path(path).write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
