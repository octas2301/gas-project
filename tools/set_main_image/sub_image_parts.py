# -*- coding: utf-8 -*-
"""
競合サブ画像のパーツ自動提案＋安全背景への組み合わせ（文字は改変しない）。

part_mode:
  auto_propose … Gemini が box 提案 → テストでは auto_accept 可
  human_box    … 人が JSON の box を渡す（後続）
"""
from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFilter, ImageFont

LOG = logging.getLogger("set_main_image.sub_image_parts")

SAFE_BG_RGB = [
    (245, 245, 243),  # off-white
    (232, 232, 230),  # light gray
    (237, 230, 220),  # beige
    (232, 224, 212),  # sand
    (217, 212, 206),  # warm gray
]

TEXT_MODELS = (
    "gemini-flash-latest",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
)

SQUARE = 1200


def _extract_json_obj(text: str) -> Optional[dict]:
    if not text:
        return None
    t = text.strip()
    if t.startswith("```"):
        t = re.sub(r"^```(?:json)?\s*", "", t)
        t = re.sub(r"\s*```$", "", t)
    try:
        obj = json.loads(t)
        if isinstance(obj, dict):
            return obj
    except json.JSONDecodeError:
        pass
    m = re.search(r"\{[\s\S]*\}", t)
    if not m:
        return None
    try:
        obj = json.loads(m.group(0))
        return obj if isinstance(obj, dict) else None
    except json.JSONDecodeError:
        return None


def propose_parts(
    path: Path,
    *,
    client=None,
    model_id: Optional[str] = None,
) -> Dict[str, Any]:
    """吹き出し・囲み・見出し帯などを正規化 box [x0,y0,x1,y1] (0-1) で提案。"""
    from gemini_image import load_api_key, make_client, mime_for
    import base64

    if client is None:
        client = make_client(load_api_key())

    b64 = base64.b64encode(path.read_bytes()).decode("ascii")
    mime = mime_for(path)
    prompt = """Amazon商品の説明・アピール系サブ画像です。
文字入りの吹き出し・囲み枠・見出し帯・バッジ・アイコン＋短文ブロックを「画像パーツ」として検出してください。
文字は読み取り用のメモだけで、書き直し案は不要です。
商品本体のパックショット領域は type=product_cut とし、説明パーツとは分けてください。

各パーツ:
- type: callout | frame_text | title_band | badge | icon_row | product_cut | other
- box: [x0,y0,x1,y1] 画像左上原点・0〜1の正規化座標
- textNote: 中の文字の要約（改変せず、読めたまま短く）
- usefulForSub: 自社サブ画像の説明パーツとして使えそうなら true（他社固有の配送・店舗・問合せは false）

JSONのみ:
{
  "parts": [
    {"id":"p1","type":"callout","box":[0.1,0.1,0.5,0.35],"textNote":"...","usefulForSub":true}
  ],
  "bgIsPhoto": false,
  "bgToneHint": "beige|gray|white|strong_color|photo"
}
"""
    models = [model_id] if model_id else list(TEXT_MODELS)
    last_err: Optional[Exception] = None
    for mid in models:
        if not mid:
            continue
        try:
            resp = client.models.generate_content(
                model=mid,
                contents=[
                    {
                        "role": "user",
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": mime, "data": b64}},
                        ],
                    }
                ],
            )
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                try:
                    text = resp.candidates[0].content.parts[0].text
                except Exception:
                    text = ""
            obj = _extract_json_obj(text)
            if not obj:
                raise RuntimeError(f"parts JSON failed: {text[:200]}")
            parts = obj.get("parts") or []
            cleaned: List[Dict[str, Any]] = []
            for i, p in enumerate(parts):
                if not isinstance(p, dict):
                    continue
                box = p.get("box")
                if not isinstance(box, (list, tuple)) or len(box) != 4:
                    continue
                try:
                    x0, y0, x1, y1 = [float(v) for v in box]
                except (TypeError, ValueError):
                    continue
                x0, y0 = max(0.0, min(1.0, x0)), max(0.0, min(1.0, y0))
                x1, y1 = max(0.0, min(1.0, x1)), max(0.0, min(1.0, y1))
                if x1 - x0 < 0.05 or y1 - y0 < 0.05:
                    continue
                cleaned.append(
                    {
                        "id": str(p.get("id") or f"p{i+1}"),
                        "type": str(p.get("type") or "other"),
                        "box": [x0, y0, x1, y1],
                        "textNote": str(p.get("textNote") or "")[:200],
                        "usefulForSub": bool(p.get("usefulForSub", True)),
                        "accepted": False,
                    }
                )
            result = {
                "path": str(path),
                "model": mid,
                "parts": cleaned,
                "bgIsPhoto": bool(obj.get("bgIsPhoto")),
                "bgToneHint": str(obj.get("bgToneHint") or ""),
                "partMode": "auto_propose",
            }
            LOG.info("parts %s n=%d", path.name, len(cleaned))
            return result
        except Exception as e:
            last_err = e
            LOG.warning("propose_parts model=%s failed: %s", mid, e)
    return {
        "path": str(path),
        "model": None,
        "parts": [],
        "bgIsPhoto": False,
        "bgToneHint": "",
        "partMode": "auto_propose",
        "error": str(last_err),
    }


def auto_accept_useful(
    parts_doc: Dict[str, Any],
    *,
    max_parts: int = 4,
    allow_full_frame_fallback: bool = True,
) -> Dict[str, Any]:
    """
    テスト用採用:
    - product_cut 以外で usefulForSub のパーツを優先
    - 説明系なのに採用0なら、最大の非product枠 or 全面フォールバック
    """
    doc = json.loads(json.dumps(parts_doc))  # copy
    explain_types = {"callout", "frame_text", "title_band", "badge", "icon_row", "other"}
    accepted = 0
    for p in doc.get("parts") or []:
        if accepted >= max_parts:
            p["accepted"] = False
            continue
        ptype = str(p.get("type") or "")
        if ptype == "product_cut":
            p["accepted"] = False
            continue
        # useful 明示 or 説明タイプなら採用候補
        if p.get("usefulForSub") or ptype in explain_types:
            p["accepted"] = True
            accepted += 1
        else:
            p["accepted"] = False

    if accepted == 0 and allow_full_frame_fallback:
        # 最大面積の非 product_cut
        candidates = [
            p
            for p in (doc.get("parts") or [])
            if p.get("type") != "product_cut" and isinstance(p.get("box"), list)
        ]
        if candidates:
            def area(p: dict) -> float:
                b = p["box"]
                return max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])

            best = max(candidates, key=area)
            best["accepted"] = True
            best["fallbackAccept"] = True
            accepted = 1
        else:
            # 全面を説明フレームとして追加（文字は画像のまま）
            full = {
                "id": "full_frame",
                "type": "frame_text",
                "box": [0.04, 0.04, 0.96, 0.96],
                "textNote": "full-frame fallback (no explain parts detected)",
                "usefulForSub": True,
                "accepted": True,
                "fallbackAccept": True,
            }
            doc.setdefault("parts", []).append(full)
            accepted = 1
            doc["usedFullFrameFallback"] = True

    doc["acceptedCount"] = accepted
    return doc


def crop_part(im: Image.Image, box01: Sequence[float]) -> Image.Image:
    w, h = im.size
    x0, y0, x1, y1 = box01
    return im.crop(
        (
            int(w * x0),
            int(h * y0),
            max(int(w * x0) + 1, int(w * x1)),
            max(int(h * y0) + 1, int(h * y1)),
        )
    ).convert("RGBA")


def clean_part_fringe(
    im: Image.Image,
    *,
    tol: int = 32,
    erode: int = 1,
) -> Image.Image:
    """
    切り抜き縁に残る元画像の背景（白・単色など）を透明化してから合成する。
    文字・図は改変せず、縁の背景画素だけ落とす。
    """
    from collections import deque

    rgba = im.convert("RGBA")
    w, h = rgba.size
    if w < 4 or h < 4:
        return rgba
    px = rgba.load()

    def sample(x: int, y: int) -> Tuple[int, int, int]:
        r, g, b, _a = px[x, y]
        return (r, g, b)

    samples = [
        sample(0, 0),
        sample(w - 1, 0),
        sample(0, h - 1),
        sample(w - 1, h - 1),
        sample(w // 2, 0),
        sample(w // 2, h - 1),
        sample(0, h // 2),
        sample(w - 1, h // 2),
    ]
    br = sum(s[0] for s in samples) // len(samples)
    bg = sum(s[1] for s in samples) // len(samples)
    bb = sum(s[2] for s in samples) // len(samples)

    def near_bg(x: int, y: int) -> bool:
        r, g, b, a = px[x, y]
        if a < 8:
            return True
        return abs(r - br) <= tol and abs(g - bg) <= tol and abs(b - bb) <= tol

    # 縁からの洪水で背景連結成分を透明化
    seen = [[False] * w for _ in range(h)]
    q: deque = deque()
    for x in range(w):
        for y in (0, h - 1):
            if near_bg(x, y) and not seen[y][x]:
                seen[y][x] = True
                q.append((x, y))
    for y in range(h):
        for x in (0, w - 1):
            if near_bg(x, y) and not seen[y][x]:
                seen[y][x] = True
                q.append((x, y))

    while q:
        x, y = q.popleft()
        r, g, b, _a = px[x, y]
        px[x, y] = (r, g, b, 0)
        for nx, ny in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
            if nx < 0 or ny < 0 or nx >= w or ny >= h:
                continue
            if seen[ny][nx]:
                continue
            if near_bg(nx, ny):
                seen[ny][nx] = True
                q.append((nx, ny))

    # わずかに内側も削って縁の食い残りを減らす
    if erode > 0:
        alpha = rgba.getchannel("A")
        for _ in range(erode):
            alpha = alpha.filter(ImageFilter.MinFilter(3))
        rgba.putalpha(alpha)

    # アルファを少しぼかして合成時の硬さを抑える
    a = rgba.getchannel("A").filter(ImageFilter.GaussianBlur(radius=0.8))
    rgba.putalpha(a)

    bbox = rgba.getchannel("A").getbbox()
    if bbox:
        rgba = rgba.crop(bbox)
    return rgba


def pick_safe_bg(index: int = 0) -> Tuple[int, int, int]:
    return SAFE_BG_RGB[index % len(SAFE_BG_RGB)]


def _font(size: int) -> ImageFont.ImageFont:
    for p in (
        Path(r"C:\Windows\Fonts\YuGothB.ttc"),
        Path(r"C:\Windows\Fonts\meiryo.ttc"),
        Path(r"C:\Windows\Fonts\arial.ttf"),
    ):
        if p.is_file():
            try:
                return ImageFont.truetype(str(p), size=size)
            except Exception:
                continue
    return ImageFont.load_default()


# サブ画像10パターン（商品ごと）
SUB_PATTERNS_10: List[Dict[str, Any]] = [
    {
        "id": "P01_text_beige",
        "ja": "説明パーツ集合・ベージュ",
        "bg": 2,
        "include_product": False,
        "slots": [
            (0.05, 0.08, 0.55, 0.38),
            (0.52, 0.08, 0.95, 0.38),
            (0.05, 0.42, 0.48, 0.72),
            (0.52, 0.42, 0.95, 0.72),
            (0.15, 0.76, 0.85, 0.96),
        ],
        "max_parts": 5,
        "source_mode": "all",
    },
    {
        "id": "P02_text_gray",
        "ja": "説明パーツ集合・グレー",
        "bg": 1,
        "include_product": False,
        "slots": [
            (0.05, 0.08, 0.55, 0.38),
            (0.52, 0.08, 0.95, 0.38),
            (0.05, 0.42, 0.48, 0.72),
            (0.52, 0.42, 0.95, 0.72),
            (0.15, 0.76, 0.85, 0.96),
        ],
        "max_parts": 5,
        "source_mode": "all",
    },
    {
        "id": "P03_text_offwhite",
        "ja": "説明パーツ集合・オフ白",
        "bg": 0,
        "include_product": False,
        "slots": [
            (0.06, 0.08, 0.94, 0.36),
            (0.06, 0.40, 0.48, 0.72),
            (0.52, 0.40, 0.94, 0.72),
            (0.10, 0.76, 0.90, 0.96),
        ],
        "max_parts": 4,
        "source_mode": "all",
    },
    {
        "id": "P04_two_column",
        "ja": "二カラム説明",
        "bg": 2,
        "include_product": False,
        "slots": [
            (0.04, 0.08, 0.48, 0.48),
            (0.52, 0.08, 0.96, 0.48),
            (0.04, 0.52, 0.48, 0.94),
            (0.52, 0.52, 0.96, 0.94),
        ],
        "max_parts": 4,
        "source_mode": "all",
    },
    {
        "id": "P05_hero_band",
        "ja": "大枠＋下部帯",
        "bg": 3,
        "include_product": False,
        "slots": [
            (0.08, 0.08, 0.92, 0.62),
            (0.08, 0.68, 0.92, 0.94),
        ],
        "max_parts": 2,
        "source_mode": "first",
    },
    {
        "id": "P06_grid_2x2",
        "ja": "2×2グリッド",
        "bg": 1,
        "include_product": False,
        "slots": [
            (0.04, 0.08, 0.48, 0.48),
            (0.52, 0.08, 0.96, 0.48),
            (0.04, 0.52, 0.48, 0.92),
            (0.52, 0.52, 0.96, 0.92),
        ],
        "max_parts": 4,
        "source_mode": "rotate",
    },
    {
        "id": "P07_product_beige",
        "ja": "パーツ＋自社商品・ベージュ",
        "bg": 2,
        "include_product": True,
        "slots": [
            (0.04, 0.08, 0.58, 0.40),
            (0.04, 0.44, 0.58, 0.72),
            (0.04, 0.76, 0.58, 0.96),
        ],
        "max_parts": 3,
        "source_mode": "all",
        "product_anchor": "right",
    },
    {
        "id": "P08_product_corner",
        "ja": "パーツ＋自社・右下寄せ",
        "bg": 0,
        "include_product": True,
        "slots": [
            (0.05, 0.08, 0.70, 0.45),
            (0.05, 0.50, 0.45, 0.90),
            (0.48, 0.50, 0.70, 0.75),
        ],
        "max_parts": 3,
        "source_mode": "all",
        "product_anchor": "br",
    },
    {
        "id": "P09_single_focus",
        "ja": "単一フォーカス＋小パーツ",
        "bg": 2,
        "include_product": False,
        "slots": [
            (0.10, 0.10, 0.90, 0.68),
            (0.08, 0.74, 0.36, 0.96),
            (0.40, 0.74, 0.68, 0.96),
            (0.72, 0.74, 0.96, 0.96),
        ],
        "max_parts": 4,
        "source_mode": "first",
    },
    {
        "id": "P10_badge_strip",
        "ja": "上部バッジ列＋本文",
        "bg": 4,
        "include_product": False,
        "slots": [
            (0.04, 0.08, 0.32, 0.28),
            (0.36, 0.08, 0.64, 0.28),
            (0.68, 0.08, 0.96, 0.28),
            (0.06, 0.34, 0.94, 0.94),
        ],
        "max_parts": 4,
        "source_mode": "all",
    },
]


def _collect_accepted_parts(
    source_images: List[Tuple[Image.Image, Dict[str, Any]]],
    *,
    source_mode: str,
    max_parts: int,
    pattern_index: int = 0,
) -> List[Tuple[Image.Image, Dict[str, Any]]]:
    """(cleaned_part_im, part_meta) を返す。縁背景除去済み。"""
    sources = list(source_images)
    if source_mode == "first":
        sources = sources[:1]
    elif source_mode == "rotate" and sources:
        rot = pattern_index % len(sources)
        sources = sources[rot:] + sources[:rot]

    out: List[Tuple[Image.Image, Dict[str, Any]]] = []
    for src_im, parts_doc in sources:
        for p in parts_doc.get("parts") or []:
            if not p.get("accepted"):
                continue
            part_im = clean_part_fringe(crop_part(src_im, p["box"]))
            out.append((part_im, p))
            if len(out) >= max_parts:
                return out
    return out


def compose_parts_board(
    *,
    source_images: List[Tuple[Image.Image, Dict[str, Any]]],
    own_product: Optional[Image.Image] = None,
    include_product: bool = False,
    bg_index: int = 2,
    title: str = "",
    slots: Optional[List[Tuple[float, float, float, float]]] = None,
    max_parts: int = 5,
    source_mode: str = "all",
    product_anchor: str = "br",
    pattern_index: int = 0,
    clean_edges: bool = True,
) -> Image.Image:
    """
    採用パーツを安全BG上に配置。文字はクロップ画像のまま（改変なし）。
    縁の元背景は clean_part_fringe で除去してから合成。
    """
    bg = Image.new("RGBA", (SQUARE, SQUARE), pick_safe_bg(bg_index) + (255,))
    if slots is None:
        slots = [
            (0.06, 0.06, 0.55, 0.32),
            (0.52, 0.08, 0.94, 0.36),
            (0.06, 0.38, 0.48, 0.72),
            (0.50, 0.40, 0.94, 0.74),
            (0.10, 0.76, 0.90, 0.94),
        ]

    parts = _collect_accepted_parts(
        source_images,
        source_mode=source_mode,
        max_parts=min(max_parts, len(slots)),
        pattern_index=pattern_index,
    )
    if not clean_edges:
        # 既に clean 済みだがフラグ互換
        pass

    for placed, (part_im, _meta) in enumerate(parts):
        if placed >= len(slots):
            break
        x0, y0, x1, y1 = slots[placed]
        bw = int(SQUARE * (x1 - x0))
        bh = int(SQUARE * (y1 - y0))
        pw, ph = part_im.size
        scale = min(bw / max(1, pw), bh / max(1, ph))
        nw, nh = max(1, int(pw * scale)), max(1, int(ph * scale))
        resized = part_im.resize((nw, nh), Image.Resampling.LANCZOS)
        ox = int(SQUARE * x0) + (bw - nw) // 2
        oy = int(SQUARE * y0) + (bh - nh) // 2
        sh = Image.new("RGBA", resized.size, (0, 0, 0, 0))
        sh.putalpha(resized.getchannel("A").point(lambda a: int(a * 0.30) if a > 8 else 0))
        sh = sh.filter(ImageFilter.GaussianBlur(7))
        bg.alpha_composite(sh, (ox + 3, oy + 4))
        bg.alpha_composite(resized, (ox, oy))

    if include_product and own_product is not None:
        prod = own_product.convert("RGBA")
        bbox = prod.getchannel("A").getbbox() if "A" in prod.getbands() else None
        if bbox:
            prod = prod.crop(bbox)
        target_w = int(SQUARE * (0.36 if product_anchor == "right" else 0.32))
        scale = target_w / max(1, prod.width)
        prod_r = prod.resize(
            (target_w, max(1, int(prod.height * scale))),
            Image.Resampling.LANCZOS,
        )
        if product_anchor == "right":
            px = SQUARE - prod_r.width - int(SQUARE * 0.04)
            py = (SQUARE - prod_r.height) // 2
        else:
            px = SQUARE - prod_r.width - int(SQUARE * 0.04)
            py = SQUARE - prod_r.height - int(SQUARE * 0.04)
        bg.alpha_composite(prod_r, (px, py))

    if title:
        draw = ImageDraw.Draw(bg)
        draw.rectangle([0, 0, SQUARE, 42], fill=(40, 44, 52, 200))
        draw.text((12, 8), title[:70], fill=(255, 230, 140), font=_font(22))

    return bg


def annotate_proposals(
    src: Image.Image,
    parts_doc: Dict[str, Any],
) -> Image.Image:
    """提案枠の可視化（人手レビュー用）。"""
    im = src.convert("RGBA").copy()
    draw = ImageDraw.Draw(im)
    w, h = im.size
    for p in parts_doc.get("parts") or []:
        box = p.get("box") or [0, 0, 0, 0]
        x0, y0, x1, y1 = box
        rect = [int(w * x0), int(h * y0), int(w * x1), int(h * y1)]
        accepted = p.get("accepted")
        color = (40, 180, 90, 255) if accepted else (220, 120, 40, 255)
        draw.rectangle(rect, outline=color, width=4)
        label = f"{p.get('id')}:{p.get('type')}"
        draw.rectangle([rect[0], max(0, rect[1] - 28), rect[0] + 8 * len(label), rect[1]], fill=color)
        draw.text((rect[0] + 4, max(0, rect[1] - 26)), label, fill=(255, 255, 255), font=_font(18))
    return im
