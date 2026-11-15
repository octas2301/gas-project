# -*- coding: utf-8 -*-
"""
B-⑤ 仮想 5セット提案（T3×3＋T4×2）。02には書かない。

① シズル＋正本袋（左下・左傾き）。金丸なし。Octasなし。
② ①＋Pillow空金丸＋Pillow確定文。
③a 構図固定／③b/c スプシ材料。AIは右上金丸聖域を空け、金丸＋Octas(右下1.3マス)はPillow後載せ。
本線は OpenAI（gpt-image-2 medium）。仮想セットは T3×3＋T4×2。
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from PIL import Image, ImageDraw, ImageFilter, ImageFont

from b5_virt_ai_headline import (
    _appeal_cfg,
    _crop_underline_stroke,
    _headline_stack,
    _to_1200,
    _trim_ink,
    _write_grid20,
)
from b5_virt_katsuo import (
    CANVAS,
    JAN,
    _badge_bbox_1200,
    _draw_tracked_line,
    _gold_params,
    _load_gold_disk,
    _stamp_gold,
    _truth_path,
    _zenkaku_len,
)
from gemini_image import generate_with_references as gemini_generate
from gemini_image import make_client
from model_policy import resolve_model_id
from openai_image import generate_with_references as openai_generate
from openai_image import resolve_openai_image_model
from octas_prep import tilt_octas_rgba
from photo_realism_rules import photo_realism_block_ja
from rakuten_badge import _font_entry, _load_font
from text_color import resolve_text_color
from transparent_bg import inspect_alpha, remove_near_white_bg
from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT, resolve_octas, resolve_underline_sample
from master_sets import _find_header_row
from sheets_master import _get_credentials, fetch_master_rows, load_sheets_settings

LOG = logging.getLogger("set_main_image.b5_virt_abc")

# サブ画像本線と同じ（sub_image_ai_compose_poc.DEFAULT_OPENAI_QUALITY）
OPENAI_QUALITY_DEFAULT = "medium"

LOCKED: Dict[str, Any] = {
    "themeId": "N1",
    "themeName": "商品名明示",
    "mode": "name",
    "upper": "だしをとらずにそのまま食べる",
    "lower": "たべるおだし",
    "facts": (
        "食べるおだし。だしをとらずにそのまま食べる。国産かつお節。"
        "お茶漬け・おつまみ。さっと一振り。やわらかい食感・だしの旨み。"
        "累計・受賞は材料に無いので書かない。"
    ),
}

P3_TEXT_MAX_CELLS = 20
P3_TEXT_RETRY_MAX_CELLS = 25
P3_FREE_MAX_TRIES = 3

# ③比較本線: free のみ（①②・③aは既定オフ。コスト削減）。
P3_VARIANTS: List[Dict[str, Any]] = [
    {"id": "b", "hook": "free", "mode": "free"},
    {"id": "c", "hook": "free", "mode": "free"},
]

# 仮想比較: T3使い方×3 + T4成果ベネフィット×2。
# scene はシズル／未来の場面だけ書く。袋の禁止は書かない（③は正本袋必須）。
SETS: List[Dict[str, Any]] = [
    {
        "id": "S1",
        "themeId": "T3",
        "nameJa": "お茶漬けシズル",
        "scene": (
            "使い方シーン。湯気のあるお茶漬け茶碗と、器に盛ったかつお節のシズルが伝わる。"
            "『これを振る／かける』が一瞬で分かる構図。照明は食欲が出る自然光。"
            "金丸・ランキングバッジ・人物の顔は描くな。"
        ),
    },
    {
        "id": "S2",
        "themeId": "T3",
        "nameJa": "おつまみ皿",
        "scene": (
            "使い方シーン。おつまみ皿に盛られたかつお節が伝わる。"
            "酒器は遠くにボケ。『今夜のおつまみ』が写真だけで分かる。"
            "金丸・ランキング・人物の顔は描くな。"
        ),
    },
    {
        "id": "S3",
        "themeId": "T3",
        "nameJa": "振る瞬間",
        "scene": (
            "使い方シーン。お茶漬けの湯気に、かつお節が今振られている手元（手首まで。顔は出すな）。"
            "降りかかる削り節が大きく見える。金丸・ランキングは描くな。"
        ),
    },
    {
        "id": "S4",
        "themeId": "T4",
        "nameJa": "すぐ一品",
        "scene": (
            "成果ベネフィット。忙しい日でも、だしをとらずにすぐお茶漬けや一品が並ぶ食卓の未来。"
            "手軽さ・時短が伝わる（時計や『時短』文字は描くな。写真の空気で示す）。"
            "金丸・ランキング・人物の顔は描くな。"
        ),
    },
    {
        "id": "S5",
        "themeId": "T4",
        "nameJa": "晩酌が華やぐ",
        "scene": (
            "成果ベネフィット。晩酌・おつまみで食卓やお酒の時間が少し贅沢になる未来。"
            "酒器は添えでよく、主役は商品とおつまみシズル。金丸・ランキング・人物の顔は描くな。"
        ),
    },
]

# ③Pillow後載せ Octas（右下・1.3マス・少し傾き）。AIには渡さない。
OCTAS_BR_CELL = CANVAS // 10
OCTAS_BR_SIZE_MAX = int(round(OCTAS_BR_CELL * 1.3))  # 156px @1200
OCTAS_BR_MARGIN = 16
OCTAS_BR_TILT_DEG = -8.0
# Octas聖域バッファ（シール外＋約1マス）。AIは文字・袋を入れない
OCTAS_BR_KEEPOUT_PAD = OCTAS_BR_CELL  # 120px


def _hex_rgb(hex_s: str) -> Tuple[int, int, int]:
    h = (hex_s or "#800020").lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _region_light_frac(im: Image.Image, box: Tuple[int, int, int, int], *, light_thresh: float = 160.0) -> Tuple[float, float]:
    rgb = im.convert("RGB").crop(box)
    if rgb.size[0] < 1 or rgb.size[1] < 1:
        return 200.0, 1.0
    small = rgb.resize(
        (max(1, rgb.size[0] // 8), max(1, rgb.size[1] // 8)),
        Image.Resampling.BOX,
    )
    pix = list(small.getdata())
    if not pix:
        return 200.0, 1.0
    lumas = [0.299 * r + 0.587 * g + 0.114 * b for r, g, b in pix]
    mean = sum(lumas) / len(lumas)
    frac = sum(1 for v in lumas if v >= light_thresh) / len(lumas)
    return mean, frac


def _headline_fill_rgba(base: Image.Image, cfg: Dict[str, Any]) -> Tuple[Tuple[int, int, int, int], str, float]:
    # 文字が乗る左上。平均だけでなく「明るい画素の割合」で白文字事故を防ぐ
    mean, frac = _region_light_frac(base, (40, 16, 640, 210), light_thresh=130)
    frac_thresh = float(cfg.get("headlineLightFracThresh") or 0.25)
    if frac >= frac_thresh or mean >= float(cfg.get("headlineFillLumaThresh") or 140):
        hex_s = str(cfg.get("baseTextOnLight") or "#0d1321")
        mode = "light_bg"
    else:
        hex_s = str(cfg.get("baseTextOnDark") or "#ffffff")
        mode = "dark_bg"
    r, g, b = _hex_rgb(hex_s)
    return (r, g, b, 255), mode, mean


def _p3_style_lock_path() -> Path:
    return Path(DEFAULT_WORK_ROOT) / "99.P3おしゃれトンマナ" / "human_best.png"


def _tint_stroke(im: Image.Image, rgb: Tuple[int, int, int], *, alpha_scale: float = 0.85) -> Image.Image:
    out = im.convert("RGBA")
    px = out.load()
    r, g, b = rgb
    w, h = out.size
    for y in range(h):
        for x in range(w):
            _, _, _, a = px[x, y]
            if a > 8:
                px[x, y] = (r, g, b, min(255, int(a * alpha_scale)))
    return out


def _load_font_px(font_id: str, pixel_h: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    size = max(12, int(pixel_h))
    entry = _font_entry(font_id)
    for path in entry.get("windows_paths") or []:
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _hook_ref_dir() -> Path:
    return Path(DEFAULT_WORK_ROOT) / "99.MAIN売れる構図参照"


def _hook_ref_paths() -> List[Path]:
    d = _hook_ref_dir()
    if not d.is_dir():
        return []
    return sorted(p for p in d.iterdir() if p.suffix.lower() in (".png", ".jpg", ".jpeg") and p.is_file())


def _oshare_typo_dir() -> Path:
    """女性・主婦向けの文字＋装飾見本（固定デザイン用04本体とは別）。"""
    return Path(DEFAULT_WORK_ROOT) / "04.楽天見本" / "99.おしゃれ文字装飾見本"


def _oshare_typo_paths(limit: int = 6) -> List[Path]:
    """スタイルシート用。退避フォルダは除外。見出し＋柔らかいT3を優先。"""
    d = _oshare_typo_dir()
    if not d.is_dir():
        return []
    soft_ids = {
        552,
        554,
        584,
        585,
        481,
        586,
        597,
        569,
        549,
        512,
        523,
        525,
        531,
        362,
        369,
        376,
        397,
        438,
        447,
        454,
        464,
        467,
        570,
        572,
        573,
        577,
        579,
        550,
        551,
        560,
    }
    midashi: List[Path] = []
    soft_t3: List[Path] = []
    other_t3: List[Path] = []
    for p in d.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
            continue
        if "_退避" in p.name:
            continue
        name = p.name
        if "midashi" in name:
            midashi.append(p)
            continue
        if "t3_" in name:
            import re as _re

            m = _re.search(r"t3_(\d+)", name)
            wid = int(m.group(1)) if m else 0
            if wid in soft_ids:
                soft_t3.append(p)
            else:
                other_t3.append(p)
    soft_t3.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    other_t3.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    midashi.sort(key=lambda p: p.name)
    picked: List[Path] = []
    for src in (midashi[:3], soft_t3, other_t3):
        for p in src:
            if p not in picked:
                picked.append(p)
            if len(picked) >= limit:
                return picked
    return picked


def _write_style_sheet(paths: List[Path], dest: Path) -> Path:
    imgs = [Image.open(p).convert("RGB") for p in paths[:6]]
    if not imgs:
        raise SystemExit("スタイルシート用の参照画像が空")
    cell_w, cell_h = 320, 320
    n = len(imgs)
    cols = min(3, n)
    rows = (n + cols - 1) // cols
    sheet = Image.new("RGB", (cell_w * cols, cell_h * rows), (245, 240, 232))
    for i, im in enumerate(imgs):
        im.thumbnail((cell_w - 12, cell_h - 12), Image.Resampling.LANCZOS)
        x = (i % cols) * cell_w + (cell_w - im.size[0]) // 2
        y = (i // cols) * cell_h + (cell_h - im.size[1]) // 2
        sheet.paste(im, (x, y))
    dest.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(dest, quality=92)
    return dest


def _keep_out() -> Tuple[int, int, int, int]:
    cfg = _appeal_cfg()
    ko = cfg.get("textKeepOut1200") or {}
    x = int(ko.get("x", 0))
    y = int(ko.get("y", 0))
    w = int(ko.get("w", 840))
    h = int(ko.get("h", 248))
    return x, y, x + w, y + h


def _load_pkg(truth: Path) -> Image.Image:
    info = inspect_alpha(truth, announce=True)
    im = Image.open(truth).convert("RGBA")
    if not info.get("canvaBgRemovedLikely"):
        im = remove_near_white_bg(im)
    bb = im.split()[-1].getbbox()
    if bb:
        im = im.crop(bb)
    return im


def _intersects(a: Tuple[int, int, int, int], b: Tuple[int, int, int, int]) -> bool:
    return not (a[2] <= b[0] or a[0] >= b[2] or a[3] <= b[1] or a[1] >= b[3])


def _place_package(
    bg: Image.Image,
    pkg: Image.Image,
) -> Tuple[Image.Image, Dict[str, Any]]:
    """正本袋を左下寄せ・左傾き。文字帯と非交差。Octasは③で右下Pillow後載せ。"""
    cfg = _appeal_cfg()
    place = cfg.get("packagePlace") or {}
    margin = int(place.get("margin", 28))
    tilt = float(place.get("tiltDegCcw", 10))
    max_h = int(round(CANVAS * float(place.get("maxHRatio", 0.58))))
    max_w = int(round(CANVAS * float(place.get("maxWRatio", 0.46))))
    ko = _keep_out()
    scaled = pkg.copy()
    scaled.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
    rotated = scaled.rotate(tilt, expand=True, resample=Image.Resampling.BICUBIC)
    rw, rh = rotated.size
    x = margin
    y = CANVAS - margin - rh
    for _ in range(10):
        box = (x, y, x + rw, y + rh)
        if y >= ko[3] - 4 and not _intersects(box, ko) and x >= 0 and y >= 0:
            break
        nw = max(80, int(rw * 0.92))
        nh = max(80, int(rh * 0.92))
        rotated = rotated.resize((nw, nh), Image.Resampling.LANCZOS)
        rw, rh = rotated.size
        x = margin
        y = CANVAS - margin - rh
    box = (x, y, min(CANVAS, x + rw), min(CANVAS, y + rh))

    canvas = bg.convert("RGBA")
    shadow = Image.new("RGBA", (CANVAS, CANVAS), (0, 0, 0, 0))
    shad = Image.new("RGBA", (rw, rh), (0, 0, 0, 90))
    shad.putalpha(rotated.split()[-1].point(lambda a: min(90, a)))
    shadow.paste(shad, (x + 10, y + 14), shad)
    shadow = shadow.filter(ImageFilter.GaussianBlur(12))
    out = Image.alpha_composite(canvas, shadow)
    out.alpha_composite(rotated, (x, y))

    meta = {
        "x": x,
        "y": y,
        "w": rw,
        "h": rh,
        "tiltDegCcw": tilt,
        "keepOut": list(ko),
        "intersectsKeepOut": _intersects(box, ko),
    }
    return out, meta


def _stamp_octas_br(base: Image.Image, octas_rgba: Image.Image) -> Image.Image:
    """画像枠の右下に Octas。外接≤1.3マス・少し傾き。AIには渡さない。

    影キャンバスは base と同じ大きさ（1024生成でも CANVAS=1200 固定で落とさない）。
    """
    oc, _ = tilt_octas_rgba(octas_rgba.convert("RGBA"), tilt_deg=OCTAS_BR_TILT_DEG)
    bb = oc.split()[-1].getbbox()
    if bb:
        oc = oc.crop(bb)
    w, h = oc.size
    out = base.convert("RGBA")
    cw, ch = out.size
    scale_c = min(cw, ch) / float(CANVAS)
    size_max = max(24, int(round(OCTAS_BR_SIZE_MAX * scale_c)))
    margin = max(4, int(round(OCTAS_BR_MARGIN * scale_c)))
    scale = min(size_max / max(1, w), size_max / max(1, h), 1.0)
    nw = max(8, int(round(w * scale)))
    nh = max(8, int(round(h * scale)))
    seal = oc.resize((nw, nh), Image.Resampling.LANCZOS)
    x = cw - margin - nw
    y = ch - margin - nh
    shadow = Image.new("RGBA", (cw, ch), (0, 0, 0, 0))
    sh = Image.new("RGBA", (nw, nh), (0, 0, 0, 70))
    sh.putalpha(seal.split()[-1].point(lambda a: min(70, a)))
    shadow.paste(sh, (x + max(1, int(4 * scale_c)), y + max(1, int(5 * scale_c))), sh)
    shadow = shadow.filter(ImageFilter.GaussianBlur(max(1, int(round(5 * scale_c)))))
    out = Image.alpha_composite(out, shadow)
    out.alpha_composite(seal, (x, y))
    return out


def _finish_p3_overlays(im: Image.Image, gold: Image.Image, octas_rgba: Image.Image) -> Image.Image:
    """③最終: 金丸→Octas（どちらもPillow）。"""
    return _stamp_octas_br(_stamp_gold(im, gold), octas_rgba)


def _prompt_scene(scene: str) -> str:
    return f"""あなたは楽天で売れる食品メインのカメラマンです。正方形。スマホの一覧で手が止まる写真。

{scene}

【IMAGE_HOOK_SHEET】止まり方・シズルの置き方の参考。商品・文字・金丸・ランキングはコピーするな。
この写真には袋を描くな（袋は後で本物を載せる）。描くのは中身・盛り付け・使い方だけ。
文字・ロゴ・透かし・ガイド・金メダル禁止。

{photo_realism_block_ja()}
"""


def _octas_keepout_bbox() -> Tuple[int, int, int, int]:
    """右下 Octas 後載せ聖域（外接上限＋バッファ）。"""
    x0 = CANVAS - OCTAS_BR_MARGIN - OCTAS_BR_SIZE_MAX - OCTAS_BR_KEEPOUT_PAD
    y0 = CANVAS - OCTAS_BR_MARGIN - OCTAS_BR_SIZE_MAX - OCTAS_BR_KEEPOUT_PAD
    return max(0, x0), max(0, y0), CANVAS - 1, CANVAS - 1


def _gold_keepout_line() -> str:
    """Pillow後載せ用。AIには金丸を描かせず、visible bbox を空ける。"""
    x0, y0, x1, y1 = _badge_bbox_1200()
    return f"""【右上・金丸聖域・空け必須】
後工程で空金丸PNGを透過合成します。完成画には金丸・金バッジ・ランキング丸・メダルを一切描かないでください。
見本（IMAGE_COMPOSE_LOCK）に金丸があってもコピーするな。
次の矩形の内側には、商品（袋）・シズル・文字・ロゴ・アイコン・吹き出し・帯を置かないでください。
- キャンバス 1200×1200
- 聖域: 左上 ({x0}, {y0}) 〜 右下 ({x1}, {y1})
  （言い換え: xが{x0}以上 かつ yが{y1}以下）
右上は背景の続き（机・壁・ボケ）のままでよい。白い四角やグレーの穴・市松を描くな。
見出しやキャッチの右端は x=880 未満に収めること。
"""


def _wrap_gold_bl_keepout_bbox() -> Tuple[int, int, int, int]:
    """左下・個包装金丸後載せの空け（右上金丸の見えと同程度〜やや大・はみ出し想定）。"""
    # 右上 visible は概ね幅~269×高~275。左下は同程度〜やや大きく端からはみ出す。
    w, h = 300, 300
    return (0, CANVAS - h, w - 1, CANVAS - 1)


def _wrap_gold_bl_keepout_line() -> str:
    """個包装アピール用・左下金丸（Pillow後載せ予定）。袋・主役個品が裏に来て消えないよう空ける。"""
    x0, y0, x1, y1 = _wrap_gold_bl_keepout_bbox()
    return f"""【左下・個包装金丸聖域・空け必須】
後工程で金枠円（右上金丸と同系・同程度〜やや大・左下はみ出し）を透過合成し、その中に個包装素材を入れる。
完成画には金丸・金枠を描くな。白い四角や市松の穴を描くな。
開封前の個包装をこの聖域（画面左下）に描くな。個包装は皿の横か後ろのみ。
次の矩形の内側には、開封前の個包装・メインの個品（開封後の主役）・正本袋・大きな包材・キャッチ帯を置かないでください（金丸の後ろに隠れて消えない／後載せ金丸に潰されないようにする）。
袋・主役シズルは聖域の外（中央〜右寄りなど）に置く。
- キャンバス 1200×1200
- 聖域: 左上 ({x0}, {y0}) 〜 右下 ({x1}, {y1})
  （言い換え: xが{x1}以下 かつ yが{y0}以上）
左下は背景の続き（机・ボケ）のままでよい。
"""


def _octas_keepout_line() -> str:
    """Pillow後載せ用。右下 Octas 聖域を空ける。"""
    x0, y0, x1, y1 = _octas_keepout_bbox()
    return f"""【右下・Octas聖域・空け必須】
後工程で期限管理シール（Octas）を右下に透過合成します。完成画にはシール・バーコード風・角バッジを描くな。
次の矩形の内側には、文字・ロゴ・吹き出し・帯・キャッチ・袋を置かないでください。
- キャンバス 1200×1200
- 聖域: 左上 ({x0}, {y0}) 〜 右下 ({x1}, {y1})
  （言い換え: xが{x0}以上 かつ yが{y0}以上）
右下は背景の続きのままでよい。白い四角や市松を描くな。
"""


def _catch_placement_line() -> str:
    return """【後載せキャッチ・配置・必須】
スマホ一覧で左上から見る前提。手を止めさせる主役は上部・左寄り。
- キャッチ1本（吹き出し・短い帯含む）は **上半分・左寄り**（目安 y≤480、右端は金丸聖域の左＝x<880）。
- **長い横長帯は上部のみ**。画面下半分の横長帯・下帯は禁止。
- 右下・Octas聖域への文字禁止。右上・金丸聖域への文字禁止。WRAPあり時は左下・個包装金丸聖域への文字・袋・主役個品も禁止。
- 袋の印字を大きく塗り潰すな。袋は読める大きさを残せ。
"""


def _prompt_p3_lock() -> str:
    return f"""【IMAGE_COMPOSE_LOCK】の構図で、楽天メイン画像を1枚作る。正方形。構図は固定（変えるな）。

固定するもの:
- 袋は木のテーブルに立って置いて撮ったように見える。右に中身が少し散らばる
- 左上の手書き見出しと、右の白い吹き出し。両方の近くに同じ三本の目立たせ棒

{_gold_keepout_line()}

{_octas_keepout_line()}

【IMAGE_PACKAGE_TRUTH】が本物の袋。袋の絵・印字はこれだけを使え。
真っ白・無地の白背景のみは不可。
文字は10×10マスで20マス以下（パッケージ印刷は数えない）。短くおしゃれに。
虚偽禁止: No.1・ランキング・受賞・累計・送料無料・無添加など材料に無い売りを足すな。
完成に％・座標・ガイド枠を描くな。

{photo_realism_block_ja()}
"""


def _catch_hero_size_line(
    *,
    h_min: Optional[int] = None,
    h_max: Optional[int] = None,
) -> str:
    """③free用。キャッチ高さ帯（既定: 下限なし・上限155。人間合格 2026-09-03）。

    環境変数:
      B5_CATCH_H_MIN … 下限 px（空＝0。0 / none ＝下限なし）
      B5_CATCH_H_MAX … 上限 px（空＝155）
    """
    import os

    if h_min is None:
        raw = (os.environ.get("B5_CATCH_H_MIN") or "").strip().lower()
        if raw in ("0", "none", "-"):
            h_min = 0
        elif raw.isdigit():
            h_min = int(raw)
        else:
            # 既定: 下限なし（レイアウト自由度）
            h_min = 0
    if h_max is None:
        raw = (os.environ.get("B5_CATCH_H_MAX") or "").strip()
        # 既定: 上限 155px（かつお／まぐろ／ぶり MAX155 合格）
        h_max = int(raw) if raw.isdigit() else 155

    h_min = int(h_min or 0)
    if h_min > 0:
        h_min = max(80, h_min)
    h_max = int(h_max or 0)
    if h_max and h_min and h_max < h_min:
        h_max = h_min

    elegant = (
        "- **おしゃれさ最優先**: 無意味に小さくするな。横幅に収めるための縮小はしてよい。"
        "細い上質な筆致・余白・1ブロック集約で上品に引き締めよ。"
        "安い極太縁・赤帯ゴシック連打・余白ゼロ詰めは禁止。\n"
        "- **書字方向**: 縦書き＝おしゃれ、ではない。指定が横書きなら横でおしゃれに。"
        "縦書きは和風トーン商品向けの選択肢に過ぎない。\n"
    )

    if h_max and h_min <= 0:
        # 上限のみ＝レイアウト自由度を上げる（下限で強制巨大化しない）
        catch_h = (
            f"文字高さ **約{h_max}px以下**（上限超で巨大化・多マス占有するな。"
            "下限は設けない——余白・字間・筆致・配置でおしゃれに整えてよい。"
            "極小化・存在感ゼロは禁止。スマホ一覧で主役として読める大きさは保て）"
        )
    elif h_max and h_min > 0:
        catch_h = (
            f"文字高さ **約{h_min}〜{h_max}px**（下限未満で存在感を消すな。"
            f"上限超で巨大化・多マス占有するな）"
        )
    else:
        # 下限のみ（明示指定時）
        catch_h = (
            f"文字高さ **約{h_min}px以上**（金丸直径343pxの約半分を目安）。"
            "スマホ一覧で目を引く主役"
        )
        elegant = (
            "- **書字方向**: 縦書き＝おしゃれ、ではない。指定が横書きなら横でおしゃれに。\n"
        )

    return f"""【後載せ・サイズ階層・必須】
- **キャッチ（大）**: 指定の文字列を一字も落とすな。基本は1行。横の枠に入りきらないときは、行を増やす前に字を小さくして収めよ。助詞（が、の、に、で、を、へ、と、は、も）の直後では改行するな。キーワード1語以上は{catch_h}。
{elegant}- **セールストーク（小）**: ある場合のみ。1行・全角6〜12字。文字高さ **約57px前後**。キャッチより明確に小さく。
- **10×10マス**: 後載せ文字＋装飾の努力目標 **20マス以内**（理想16）。21–24はバッファ。25以上は避ける。**おしゃれさ・目を引く力を最優先**（コンパクト化で地味にしない）。
- 和風大掛け軸・白矩形パネル・太帯・多段囲み・下帯横長は禁止。上部左にコンパクト1ブロック。
- キャッチ無し・極小文字だけ・補足だけの構成は不可。
おしゃれに。袋・シズル・金丸・Octasとバランス。背景と同色にしない。
"""


def _style_typo_line() -> str:
    return """【IMAGE_STYLE_TYPO・文字と装飾の見本】
このコラージュは文字＋装飾の見本です。フォント感・装飾（帯・縁・吹き出し・余白・丸み）だけ借りる。
文字色・強調色は上の textColor 固定。見本の色は使うな。
見本の商品名・ロゴ・パッケージ絵はコピーするな。
赤帯ゴシック連打・極太チープ縁・少年誌ポップ・ギルト系は禁止。
"""


def _prompt_p3_free(
    facts: str,
    *,
    theme_id: str,
    name_ja: str,
    scene: str,
    retry_note: str = "",
    catch: str = "",
    text_color: str = "1",
    has_food_unit_canon: bool = False,
    has_food_unit_sizzle: bool = False,
    has_food_unit_secondary: bool = False,
    has_food_unit_wrap: bool = False,
    has_food_unit_piece_set: bool = False,
    food_mode: str = "identity",
    layout_reference: bool = False,
    food_tune_note: str = "",
    grid_budget_block: str = "",
    mix_parts: Optional[List[str]] = None,
    mix_part_bags: bool = False,
    stamp_layers: bool = False,
    food_look_block: str = "",
    wrap_look_block: str = "",
    bag_aspect_block: str = "",
    size_arm_block: str = "",
    photo_arm: str = "",
    catch_head: Optional[str] = None,
    sales_given: Optional[str] = None,
    dish_part_name: str = "",
) -> str:
    from theme_compose_rules import accent_text_block
    from food_unit_canon import food_unit_prompt_block_main

    accent = accent_text_block(text_color)
    retry = f"\n【再作成】{retry_note}\n" if retry_note else ""
    theme = theme_id.strip() or "T3"
    if theme == "N1" or theme.startswith("N1"):
        theme_ja = "商品名明示（袋が大きく読める）"
    elif theme == "T1":
        theme_ja = "実績・受賞（根拠のある数字のみ。トロフィーは作らない）"
    elif theme == "T2":
        theme_ja = "産地・素材（原料の静物。袋は残す）"
    elif theme == "T3":
        theme_ja = "使い方・シーン（どう使うか）"
    elif theme == "T4":
        theme_ja = "成果・未来ベネフィット（買うとどんな良い状態になるか）。食感マクロだけの説明にしない"
    else:
        theme_ja = theme
    from b5_main_base_names import catch_sales_pair_ja

    # compose は catch_head / sales_given。旧 virt free は catch= のみ。
    # catch_head を無視すると固定ブロックが空になり、言い回し候補から自作文される。
    if catch_head is not None or sales_given is not None:
        catch_image = (catch_head or "").strip()
        sales_talk = (sales_given or "").strip()
    else:
        catch_full = (catch or "").strip()
        catch_image, sales_talk = (
            catch_sales_pair_ja(catch_full) if catch_full else ("", "")
        )
    scene_text = scrub_old_wrap_mix_ja((scene or "").strip()) or "（テーマに合うシズル・場面）"
    dish_part = str(dish_part_name or "").strip()
    if (
        dish_part
        and dish_part not in ("（構成未設定）", "構成未設定")
        and "載せるのは" not in scene_text
        and "載せる具は" not in scene_text
    ):
        scene_text = (
            scene_text.rstrip("。")
            + "。"
            + f"添える料理の上に載せる具は構成「{dish_part}」の見本だけ。"
            "他構成の個品をその料理の上に出すな。他の構成は小皿の見本のみ。"
        )
    if has_food_unit_wrap:
        # 場面文の「個包装を描くな」は WRAP 添付と衝突するので落とす（汎用）。
        # 「手前に出すな」等の置き位置は残す（出すな、だけでは落とさない）。
        scene_text = "。".join(
            s
            for s in re.split(r"[。．]", scene_text)
            if s.strip()
            and "個包装を新たに描くな" not in s
            and not (
                "個包装" in s
                and ("描くな" in s or "創作" in s)
                and "見本がある" not in s
                and "手前" not in s
                and "より前" not in s
            )
        )
        if scene_text and not scene_text.endswith("。"):
            scene_text += "。"
        scene_text = scrub_old_wrap_mix_ja(scene_text) or "（テーマに合うシズル・場面）"
        if "小皿に混ぜるな" not in scene_text or "カメラ最近傍" not in scene_text:
            scene_text = scene_text.rstrip("。") + "。" + _SCENE_NO_INVENT_WRAP
    scene_block = f"""【構図・演出＝絵の指示のみ】
{scene_text}
※上記は構図・シズル・場面の指示。概要の文章そのものを画像内に書くな。"""
    catch_block = ""
    draw_head = catch_image or sales_talk
    draw_sales = sales_talk if catch_image else ""
    if draw_head:
        sales_line = (
            f"セールストーク（小・添える）: {draw_sales}\n"
            if draw_sales
            else "セールストーク（小）: （無し。キャッチのみでよい）\n"
        )
        catch_block = f"""【後載せ文字＝キャッチ大＋セールストーク小（この通り・固定）】
キャッチ（大・目を引く）: {draw_head}
{sales_line}※上記の文字は、一字も落とさず、この順で書け。言い換え・長文化・第3の文句追加禁止。語の途中切れ禁止。字数の目安で切るな。
※材料の「言い回し候補」や特徴文からキャッチを作るな。後載せは上の指定文だけ。
"""
    color_priority = """【色の優先順位】
強調色・核語の色は textColor（上記 accent）が IMAGE_STYLE_TYPO より優先。見本シートの色に寄せて別色相にしない。
"""
    food_block = food_unit_prompt_block_main(
        has_identity=bool(has_food_unit_canon),
        has_sizzle=bool(has_food_unit_sizzle),
        has_secondary=bool(has_food_unit_secondary),
        has_wrap=bool(has_food_unit_wrap),
        has_piece_set=bool(has_food_unit_piece_set),
        food_mode=str(food_mode or "identity"),
        mix_parts=mix_parts,
        mix_part_bags=bool(mix_part_bags),
        stamp_layers=bool(stamp_layers),
    )
    layout_block = ""
    if layout_reference:
        if stamp_layers:
            layout_block = """
【IMAGE_LAYOUT_REFERENCE・構図維持（最優先）】
添付の IMAGE_LAYOUT_REFERENCE は直前の完成見本。背景・照明・カメラ角度・金丸/Octas の空き方を維持。
袋と皿の上の個品は描くな（後載せ）。変更してよいのは後載せ文字（キャッチ大＋セールストーク小）だけ。
"""
        else:
            layout_block = """
【IMAGE_LAYOUT_REFERENCE・構図維持（最優先）】
添付の IMAGE_LAYOUT_REFERENCE は直前の完成見本。背景・盛り付け・袋の位置と大きさ・
照明・カメラ角度・金丸/Octas の空き方はこの見本を維持。シズル配置を大きく変えるな。
変更してよいのは後載せ文字（キャッチ大＋セールストーク小）だけ。
"""
    tune_block = ""
    tune_note = str(food_tune_note or "").strip()
    if tune_note and not stamp_layers:
        tune_block = f"""
【個品見本の微調整】
{tune_note}
IMAGE_FOOD_UNIT_CANON の色・魚種・種類は維持。厚み・1片サイズ・盛り量だけ上記に合わせる。
正本に無い厚切り・ステーキ断面・別色の塗り替え禁止。
"""
    mix_names = [str(x).strip() for x in (mix_parts or []) if str(x).strip()]
    if len(mix_names) >= 2 and not stamp_layers:
        pkg_block = """【袋の役割・組み合わせ】
IMAGE_PACKAGE_SCALE は、縦横比と、袋どうしの大きさの見本だけ。この写真自体を完成画に置くな。並びもコピーするな。
IMAGE_PACKAGE_PART_i が、完成画に描く袋。各袋の縦横比は、その単品のまま。潰すな。
袋どうしの大きさの対比だけ、IMAGE_PACKAGE_SCALE と同じにしろ。片方だけ小さくするな。
袋の主軸（通常は左→右）は PART_1 → PART_2 → … の順（構成の順）。入れ替えるな。
個品を小皿などに分けるときは、個品の並びも袋と同じ順・同じ向きに連動させろ（PART_i の近く／同列に IDENTITY_i）。
奥行きは盛り付けの後ろ。重なったところは、前の器や食品が包材を隠す。器を透明にして後ろの包材を透かすな。
包材の印刷面は、ファイルにある窓以外は透明にするな。
窓の中に見える中身も正本ファイルどおり。別物の中身・別包装に描き変えるな。
袋は場面の中に描け。上から貼った形にするな。置いた面の光・接地・その面と同じ距離のピントに乗せろ。包材全体はぼかすな。外側の縁だけ、隣の物の縁と同じ柔らかさにしろ。影は下端の形に沿わせろ。白抜き写真の照明を写すな。
光は画面で1つ。包材の影の方向と濃さは、同じ面の食品・皿と同じ。包材を白飛びさせるな。包材だけ影を薄くしない。
印刷面の色相・彩度・明度帯は正本ファイルのまま（例: 濃紺は濃紺。ハイライトの青みを残せ）。場面光で黒・無彩色に落とすな。影の向きだけ合わせる。
IMAGE_FOOD_UNIT_WRAP は個包装専用。WRAP写真の大袋・背景で PACKAGE を描き変えるな。個包装は1〜2個まで。皿の横か後ろのみ（小皿に混ぜるな）。カメラ最近傍は開封後の個品。個包装はそれより手前に出すな。画面左下（個包装金丸聖域）に個包装を置くな。見本に無い空気だまり禁止。主役は開封後の個品。
開封後の個品の色・縁・筋・血合いは、IMAGE_FOOD_UNIT_PIECE（一片セット）が添付されているときはその写真＋一片言語化が正。IDENTITY／CANON は盛り配置・量感のみ。一片が無いときは IDENTITY／CANON。袋窓の中身は PACKAGE 正本のまま（一片で窓を描き変えるな）。場面の暖色・夜景で下層を黒く落とすな。影の向きだけ場面に合わせる。
包材の輪郭はファイルの縁のまま。波・しわ・シール縁を直線の四角に伸ばすな。影を合わせるために輪郭を平らにするな。
印字は単品の絵柄を保て。別デザインに描き直すな。見本写真そのものを完成画に残すな。描く包材は盛り付けの後ろ。袋を並べて主役にするな。
金丸聖域（右上）・Octas聖域（右下）には袋・文字を入れない。WRAPあり時は左下・個包装金丸聖域にも袋・主役個品・開封前個包装を入れない。
"""
    else:
        pkg_block = """【IMAGE_PACKAGE_TRUTH・袋必須】
添付の袋が本物です。完成画に必ず置け。省略・別デザイン差し替えは不可。
① 縦横比を変えるな（潰すな）。
② 袋が複数種類あるとき、正本どうしの縮尺の対比を変えるな。片方だけ小さくするな。
印字・絵柄を描き直して別物にするな。左右の位置はテーマで変えてよい。奥行きは盛り付けの後ろ。重なったところは、前の器や食品が包材を隠す。器を透明にして後ろの包材を透かすな。
包材の印刷面は、ファイルにある窓以外は透明にするな。
窓の中に見える中身も正本ファイルどおり。別物の中身・別包装に描き変えるな。
袋は置いた面の光・接地・その面と同じ距離のピントに乗せろ。包材全体はぼかすな。外側の縁だけ、隣の物の縁と同じ柔らかさにしろ。影は下端の形に沿わせろ。白抜き写真の照明を写すな。
光は画面で1つ。包材の影の方向と濃さは、同じ面の食品・皿と同じ。包材を白飛びさせるな。包材だけ影を薄くしない。
印刷面の色相・彩度・明度帯は正本ファイルのまま（例: 濃紺は濃紺。ハイライトの青みを残せ）。場面光で黒・無彩色に落とすな。影の向きだけ合わせる。
IMAGE_FOOD_UNIT_WRAP は個包装専用。WRAP写真の大袋・背景で PACKAGE を描き変えるな。個包装は1〜2個まで。皿の横か後ろのみ（小皿に混ぜるな）。カメラ最近傍は開封後の個品。個包装はそれより手前に出すな。画面左下（個包装金丸聖域）に個包装を置くな。見本に無い空気だまり禁止。主役は開封後の個品。
開封後の個品の色・縁・筋・血合いは、IMAGE_FOOD_UNIT_PIECE（一片セット）が添付されているときはその写真＋一片言語化が正。IDENTITY／CANON は盛り配置・量感のみ。一片が無いときは IDENTITY／CANON。袋窓の中身は PACKAGE 正本のまま（一片で窓を描き変えるな）。場面の暖色・夜景で下層を黒く落とすな。影の向きだけ場面に合わせる。
包材の輪郭はファイルの縁のまま。波・しわ・シール縁を直線の四角に伸ばすな。影を合わせるために輪郭を平らにするな。
金丸聖域（右上）・Octas聖域（右下）には包材・文字を入れない。WRAPあり時は左下・個包装金丸聖域にも袋・主役個品・開封前個包装を入れない。袋を並べて主役にするな。
"""
    if stamp_layers:
        pkg_block = """【袋・個品は後載せ・描くな】
IMAGE_LAYOUT_ZONES の空きを守れ。袋を描くな。皿の上に個品を描くな。
原料の雰囲気は背景に限り可。見本に無い形は原料イメージにも出すな。
金丸聖域（右上）・Octas聖域（右下）には袋・文字を入れない。WRAPあり時は左下・個包装金丸聖域にも袋・主役個品・開封前個包装を入れない。
"""
    theme_follow = (
        "袋と個品は後載せ。シズル・場面（皿・背景・小物）はこのテーマに合わせること（他テーマに寄せるな）:"
        if stamp_layers
        else "袋を置いたうえで、シズル・場面はこのテーマに合わせること（他テーマに寄せるな）:"
    )
    bg_line = (
        "【背景】真っ白・無地の白背景のみは不可。テーマのシズル／場面が伝わる背景にせよ。袋は描くな。"
        if stamp_layers
        else "【背景】真っ白・無地の白背景のみは不可。袋＋材料にある場面。原料の雰囲気は背景に限り可。見本に無い形は原料イメージにも出すな。"
    )
    budget_block = (grid_budget_block or "").strip()
    if budget_block:
        budget_block = budget_block.rstrip() + "\n\n"
    return f"""楽天のメイン画像の規約に則って、売れる商品ページのメイン画像を作成してください。正方形。スマホ一覧でお客さまが思わず手を止めるメイン。
トーンは材料にある用途・場面に合わせる。材料に無い生活シーンを決め打ちするな。
{retry}
{pkg_block}
{(bag_aspect_block or "").strip()}
{food_block}
{(food_look_block or "").strip()}
{(wrap_look_block or "").strip()}
{layout_block}
{tune_block}
【テーマ】themeId={theme}（{theme_ja}）／案名={name_ja}
{theme_follow}
{scene_block}
{catch_block}
{budget_block}{accent.strip()}
{color_priority}
{_style_typo_line()}

{_gold_keepout_line()}

{_wrap_gold_bl_keepout_line() if has_food_unit_wrap else ""}

{_octas_keepout_line()}

{_catch_placement_line()}

{bg_line}

【後載せ文字・必須】パッケージ印刷以外の文言は **キャッチ大＋（任意）セールストーク小の1セット** だけ（0セット不可。別帯・第3文句・多段帯禁止）。
キャッチは全角おおよそ6〜18字。セールストークは全角おおよそ8〜16字・小さく。全文の長文パネル禁止。
矢印・吹き出し・帯は上記セットの装飾として、見本（IMAGE_STYLE_TYPO）に合うおしゃれなセットで。核語の色は上の textColor を厳守（見本の色に寄せて別色相にしない）。セールストークは濃茶／ダーク系で可読性優先（アクセント一色にしない）。

{_catch_hero_size_line()}
上に書いたキャッチとセールストークを、一字も落とさず書け。字数で切るな。セールストークはキャッチより小さく。全文の長文パネル禁止。
【文字20%・厳守】画像を10×10＝100マスに等分し、後載せの文字・ロゴ・アイコン・バッジ・文字装飾（帯・縁・影・囲み）が1pxでも掛かるマスは合計 **目標16以下**（金丸に後載せ4マス予約）。**上限24・30超禁止**。パッケージ印刷は数えない。四辺囲み禁止。キャッチ＋セールストーク＋装飾も含めコンパクト1ブロック。

【売り文句・虚偽禁止】次の事実・調査シート材料の言い換えだけ使ってください。材料に無い売りは一切禁止（送料無料・メール便・無添加・創業・累計・受賞・No.1・ランキングなど）。
使える事実（マスタ＋AI情報取得data）:
{facts}

完成に％・座標・ガイド枠を描かないでください。

{photo_realism_block_ja(photo_arm)}
{(size_arm_block or "").strip()}
"""


def _gen_ai(
    engine: str,
    *,
    prompt: str,
    image_paths: List[Path],
    image_roles: List[str],
    gemini_client: Any = None,
    gemini_model: str = "",
    openai_model: str = "",
    openai_quality: str = OPENAI_QUALITY_DEFAULT,
) -> Tuple[bytes, Dict[str, Any]]:
    """同じ prompt / 参照画像で Gemini または OpenAI を呼ぶ。"""
    if engine == "gemini":
        if gemini_client is None:
            raise RuntimeError("gemini_client が無い")
        raw, meta = gemini_generate(
            client=gemini_client,
            model_id=gemini_model,
            prompt=prompt,
            image_paths=image_paths,
            image_roles=image_roles,
            aspect_ratio="1:1",
            image_size="1K",
        )
        out = dict(meta or {})
        out.update({"engine": "gemini", "model": gemini_model})
        return raw, out
    if engine == "openai":
        role_header = "\n".join(
            f"- File order {i + 1}: {image_roles[i]} = {image_paths[i].name}"
            for i in range(len(image_paths))
        )
        full = prompt + "\n\nFILE ORDER:\n" + role_header
        mid = openai_model or resolve_openai_image_model()
        raw, used = openai_generate(
            prompt=full,
            image_paths=image_paths,
            model=mid,
            size="1024x1024",
            quality=openai_quality,
        )
        return raw, {
            "engine": "openai",
            "model": used,
            "apiPath": "images.edit",
            "quality": openai_quality,
        }
    raise ValueError(f"unknown engine={engine}")


def _jan_digits(s: str) -> str:
    return "".join(ch for ch in str(s) if ch.isdigit())


def _name_base_key(s: str) -> str:
    return " ".join(str(s or "").split())


def _sheet_cell(row: List[Any], index: Optional[int]) -> str:
    if index is None or index >= len(row):
        return ""
    return str(row[index] or "").strip()


def _sheet_checked(value: str) -> bool:
    return str(value or "").strip().upper() in ("TRUE", "1")


def _ai_field(ai: Dict[str, str], *names: str) -> str:
    for n in names:
        v = (ai.get(n) or "").strip()
        if v and v != "不明":
            return v
    return ""


def _compact_facts_from_sheets(
    jan: str, name_base: str = "", parent_sku: str = ""
) -> Dict[str, Any]:
    """マスタ＋AI情報取得data（正解優先）から③用の短い事実リストを作る。

    name_base / parent_sku があるときは、その親行だけを材料にする。
    同じJANの先に出る単品行では埋めない。
    """
    from googleapiclient.discovery import build

    master_bits: Dict[str, str] = {}
    ai_bits: Dict[str, str] = {}
    try:
        rows, info = fetch_master_rows()
        hi, idx = _find_header_row(rows)
        jan_i = idx.get("JANコード") or idx.get("JAN") or idx.get("商品コード(JANコード等)")
        parent_i = idx.get("親SKU")
        child_i = idx.get("子SKU")
        name_i = idx.get("商品名ベース")
        ck_i = idx.get("出品CK")
        want = _jan_digits(jan)
        want_nb = _name_base_key(name_base)
        want_sku = str(parent_sku or "").strip()
        candidates: List[List[Any]] = []
        for r in rows[hi + 1 :]:
            if jan_i is None or jan_i >= len(r):
                continue
            got = _jan_digits(r[jan_i])
            if got and (got == want or want in got):
                candidates.append(r)
        hit = None
        if want_nb or want_sku:
            best_score = -10**9
            for r in candidates:
                sku = _sheet_cell(r, parent_i)
                child = _sheet_cell(r, child_i)
                nb = _name_base_key(_sheet_cell(r, name_i))
                score = 0
                if want_sku:
                    score += 100 if sku == want_sku else -80
                if want_nb:
                    score += 40 if nb == want_nb else -40
                if not child:
                    score += 10
                if _sheet_checked(_sheet_cell(r, ck_i)):
                    score += 5
                if score > best_score:
                    best_score = score
                    hit = r
            if hit is not None:
                if want_sku and _sheet_cell(hit, parent_i) != want_sku:
                    hit = None
                elif want_nb and _name_base_key(_sheet_cell(hit, name_i)) != want_nb:
                    hit = None
        elif candidates:
            hit = candidates[0]
        if hit:
            for name in (
                "商品名ベース",
                "メーカー名ベース",
                "メーカー名",
                "商品説明の箇条書き①",
                "商品説明の箇条書き②",
                "商品説明の箇条書き③",
                "楽天キャッチコピーAIから取得",
                "商品説明文AIから取得",
            ):
                j = idx.get(name)
                if j is not None and j < len(hit) and str(hit[j]).strip():
                    master_bits[name] = str(hit[j]).strip()
        sid = str(info.get("spreadsheetId") or "")
        settings = load_sheets_settings()
        creds = _get_credentials(settings["credentials_path"], settings["token_path"])
        sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
        data = (
            sheets.spreadsheets()
            .values()
            .get(
                spreadsheetId=sid or settings["spreadsheet_id"],
                range="'AI情報取得data'",
                majorDimension="ROWS",
            )
            .execute()
        )
        araw = data.get("values") or []
        if araw:
            ah = [str(h) for h in araw[0]]
            ai_idx = {h: i for i, h in enumerate(ah)}
            ji = next((ai_idx[k] for k in ah if "JAN" in k.upper()), None)
            name_ai_i = ai_idx.get("商品名")
            want = _jan_digits(jan)
            ai_hits: List[List[Any]] = []
            for r in araw[1:]:
                ok = False
                if ji is not None and ji < len(r) and want and want == _jan_digits(r[ji]):
                    ok = True
                elif want and any(want == _jan_digits(c) for c in r[:12]):
                    ok = True
                if ok:
                    ai_hits.append(r)
            chosen_ai = None
            if want_nb:
                for r in ai_hits:
                    if _name_base_key(_sheet_cell(r, name_ai_i)) == want_nb:
                        chosen_ai = r
                        break
            elif ai_hits:
                chosen_ai = ai_hits[0]
            if chosen_ai is not None:
                for h, i in ai_idx.items():
                    if i < len(chosen_ai) and str(chosen_ai[i]).strip():
                        ai_bits[h] = str(chosen_ai[i]).strip()
    except Exception as e:
        LOG.exception("材料読取失敗: %s", e)
        return {
            "factsLine": "材料を読め。材料に無い売りは書くな。",
            "master": master_bits,
            "ai": ai_bits,
            "error": str(e),
            "materials": {},
        }

    master_name = master_bits.get("商品名ベース") or ""
    if want_nb and _name_base_key(master_name) == want_nb:
        name = master_name
    else:
        name = _ai_field(ai_bits, "商品名") or master_name
    maker = _ai_field(ai_bits, "メーカー名") or master_bits.get("メーカー名ベース") or master_bits.get("メーカー名") or ""
    feats = _ai_field(
        ai_bits,
        "特徴キーワード[正解]",
        "特徴キーワード[GPT]",
        "▼マスタ(特徴キーワード)",
    ).replace("\n", "／")
    kw = _ai_field(
        ai_bits,
        "検索KW(商品名用)[正解]",
        "検索KW(商品名用)[GPT]",
        "▼マスタ(検索KW(商品名用))",
    )
    yahoo = _ai_field(
        ai_bits,
        "Yahooキャッチ(20-30字)[正解]",
        "Yahooキャッチ(20-30字)[GPT]",
        "▼マスタ(Yahooキャッチ(20-30字))",
    )
    rakuten = _ai_field(
        ai_bits,
        "楽天キャッチ(60-80字)[正解]",
        "楽天キャッチ(60-80字)[GPT]",
        "▼マスタ(楽天キャッチ(60-80字))",
    )
    bullets = []
    for i in range(1, 6):
        b = _ai_field(ai_bits, f"箇条書き{i}[正解]", f"箇条書き{i}[GPT]", f"▼マスタ(箇条書き{i})")
        if not b:
            b = master_bits.get(f"商品説明の箇条書き{'①②③④⑤'[i - 1]}") or ""
        if b:
            bullets.append(b)
    # プロンプト用は短く（長文を渡すと文字20%超えを誘発する）
    lines = [
        f"商品名: {name}",
        f"メーカー: {maker}" if maker else "",
        f"特徴: {feats}" if feats else "",
        f"用途KW: {kw}" if kw else "",
        "言い回し候補（場面のニュアンス用。後載せキャッチの種にするな。全文を画像に載せるな）:",
    ]
    # 候補は短く少数（長い材料を渡すと文字20%超えを誘発）
    short_bits = []
    for y in yahoo.splitlines():
        y = y.strip()
        if y:
            short_bits.append(y[:28])
        if len(short_bits) >= 2:
            break
    for b in bullets[:2]:
        short_bits.append(b[:28])
    for bit in short_bits[:3]:
        lines.append(f"- {bit}")
    lines.append("禁止: 材料に無い送料無料・無添加・創業・累計・受賞・No.1")
    facts_line = "\n".join(x for x in lines if x)
    return {
        "factsLine": facts_line or "材料を読め。材料に無い売りは書くな。",
        "master": master_bits,
        "ai": {k: v[:200] for k, v in ai_bits.items() if "[正解]" in k or k in ("商品名", "メーカー名")},
        "error": "",
        "materials": {
            "name": name,
            "maker": maker,
            "feats": feats,
            "kw": kw,
            "yahoo": yahoo,
            "rakuten": rakuten,
            "bullets": bullets,
            "desc": _ai_field(
                ai_bits,
                "商品説明[正解]",
                "商品説明[GPT]",
                "▼マスタ(商品説明)",
            )
            or master_bits.get("商品説明文AIから取得")
            or "",
        },
    }


# 売れる順。成果は材料の【見出し】・箇条書き。食品固定の晩酌表は使わない。
_SELL_BAN = ("健康", "低糖質", "ダイエット", "糖質オフ", "カロリーオフ", "血圧", "免疫")
# 根拠の薄い安心訴求は見出しにしない（景表・モール誤認の入口）。
_WEAK_HEAD_BAN = ("安心感", "安心安全", "安心・安全", "安心", "安全への", "安全にこだわ")
_T1_MARK = ("受賞", "金賞", "累計", "突破", "億粒", "万枚", "万袋")
_T2_MARK = ("国産", "産地", "原料", "原産", "国内産")


def _mat_blob(materials: Dict[str, Any]) -> str:
    bits = [
        str(materials.get("name") or ""),
        str(materials.get("feats") or ""),
        str(materials.get("kw") or ""),
        str(materials.get("yahoo") or ""),
        str(materials.get("rakuten") or ""),
        str(materials.get("desc") or ""),
        " ".join(str(x) for x in (materials.get("bullets") or [])),
    ]
    return "\n".join(bits)


def _snip_around(text: str, needles: Tuple[str, ...], limit: int = 80) -> str:
    """根拠引用。語の途中開始禁止。【…】優先。ヒット語から前に食い込まない。"""
    t = " ".join(str(text or "").split())
    if not t:
        return ""
    for n in needles:
        i = t.find(n)
        if i < 0:
            continue
        # ヒット語が【見出し】の中にあるときだけ、その見出しを使う（直後の】外は含めない）
        for m in re.finditer(r"【[^】]{1,40}】", t):
            frag = m.group(0)
            if n in frag:
                return frag
        # ヒット語自体から始める（前に食い込まない）。直後の短い続きだけ
        end = min(len(t), i + max(len(n), 24))
        for sep in ("。", "！", "？", "\n", "【"):
            j = t.find(sep, i + len(n))
            if j >= 0:
                end = min(end, j)
        return t[i:end].strip()
    return t[:limit]


def _prominence(materials: Dict[str, Any], needles: Tuple[str, ...]) -> Tuple[int, str]:
    """材料の目立ち度。戻り: (加点, 根拠引用)。"""
    checks: List[Tuple[int, str]] = [
        (12, str(materials.get("rakuten") or "")),
        (8, str(materials.get("yahoo") or "")),
    ]
    bullets = list(materials.get("bullets") or [])
    for i, b in enumerate(bullets):
        checks.append((6 if i < 2 else 3, str(b)))
    checks.append(
        (
            1,
            str(materials.get("desc") or "")
            + " "
            + str(materials.get("kw") or "")
            + " "
            + str(materials.get("feats") or ""),
        )
    )
    best = 0
    ev = ""
    for pts, text in checks:
        if any(n in text for n in needles) and pts > best:
            best = pts
            ev = _snip_around(text, needles)
    return best, ev


def _has_any(text: str, marks: Tuple[str, ...]) -> bool:
    return any(m in text for m in marks)


CATCH_LEN_MIN = 6
CATCH_LEN_MAX = 18


def _zenkaku_slice(s: str, max_n: float) -> str:
    out: List[str] = []
    n = 0.0
    for ch in s:
        add = 0.5 if re.match(r"[ -~]", ch) else 1.0
        if n + add > max_n:
            break
        out.append(ch)
        n += add
    return "".join(out)


def _trim_catch_len(t: str) -> str:
    """上限内に収める。語の途中で切らず、長くなりすぎない。"""
    t = (t or "").strip()
    if not t:
        return t
    if t.startswith("【") and "】" not in t:
        t = t[1:]
    if _zenkaku_len(t) <= CATCH_LEN_MAX:
        return t
    cut = _zenkaku_slice(t, CATCH_LEN_MAX)
    # 句読点・括弧で戻る
    for i in range(len(cut) - 1, 0, -1):
        if cut[i] in "。！？、・／】」』":
            cand = cut[: i + (0 if cut[i] in "、・／" else 1)].rstrip("、・／")
            if _zenkaku_len(cand) >= CATCH_LEN_MIN:
                return cand
    # 助詞終わり（一口に／お供に 等）まで戻る
    for i in range(len(cut) - 1, 0, -1):
        if cut[i] in "にをがはでともへのや" and _zenkaku_len(cut[: i + 1]) >= CATCH_LEN_MIN:
            return cut[: i + 1]
    # ひらがな終わりまで戻る（漢字の途中＝「仕」等を避ける）
    for i in range(len(cut) - 1, 0, -1):
        ch = cut[i]
        if "\u3040" <= ch <= "\u309f" and _zenkaku_len(cut[: i + 1]) >= CATCH_LEN_MIN:
            return cut[: i + 1]
    return _zenkaku_slice(t, 14)


def _short_catch(text: str, fallback: str) -> str:
    """キャッチ案。【…】優先。6〜18字。語の途中開始・途中終了を避ける。"""
    raw = " ".join(str(text or "").split()).strip()
    if not raw:
        raw = fallback
    m = re.search(r"【([^】]{2,24})】", raw)
    if m:
        inner = m.group(1).strip()
        t = f"【{inner}】" if _zenkaku_len(inner) <= 16 else inner
    else:
        t = re.sub(r"\s+", "", raw)
        t = re.sub(r"^[にもをはがでとにのへや]、*", "", t)
        t = re.sub(r"^】+", "", t)
    if not t:
        t = re.sub(r"\s+", "", fallback or "")
    t = _trim_catch_len(t)
    if _zenkaku_len(t) < CATCH_LEN_MIN and fallback:
        t = _trim_catch_len(re.sub(r"\s+", "", fallback))
    return t or fallback


PROPOSE_MAX = 6


def _assign_unique_theme_id(row: Dict[str, Any], used: Set[str]) -> Dict[str, Any]:
    out = dict(row)
    base = str(out.get("themeId") or "T3").strip()
    if base not in used:
        used.add(base)
        return out
    for ch in "bcdefghijk":
        cand = f"{base}{ch}"
        if cand not in used:
            out["themeId"] = cand
            used.add(cand)
            return out
    cand = f"{base}_{len(used)}"
    used.add(cand)
    out["themeId"] = cand
    return out


def expand_proposals(
    materials: Dict[str, Any],
    existing: List[Dict[str, Any]],
    *,
    target: int = PROPOSE_MAX,
) -> List[Dict[str, Any]]:
    """既存案を維持しつつ target 件まで材料から追加案を補う。"""
    out = [dict(p) for p in (existing or []) if isinstance(p, dict)]
    if len(out) >= target:
        return out[:target]
    used = {str(p.get("themeId") or "") for p in out}
    catches = {str(p.get("catch") or "").strip() for p in out}
    for row in propose_from_materials(materials):
        catch = str(row.get("catch") or "").strip()
        if not catch or catch in catches:
            continue
        row = _assign_unique_theme_id(row, used)
        row["sellRank"] = len(out) + 1
        out.append(row)
        catches.add(catch)
        if len(out) >= target:
            break
    for i, p in enumerate(out):
        p["sellRank"] = i + 1
    return out[:target]


def _origin_only(s: str) -> bool:
    t = re.sub(r"[の・\s　]", "", str(s or ""))
    return bool(re.fullmatch(r"(国産|産地|原料|原産|国内産)+", t))


# 場面の語（描く条件用）。個品の厚み・食感はここに入れない。
_SCENE_MARKS: Tuple[Tuple[str, str, str], ...] = (
    ("drink", "晩酌の卓。袋は必須。手順をキャッチに書くな。", ("晩酌", "家飲み", "お酒好き", "肴", "酒のつまみ", "酒が進む")),
    ("snack", "おやつの卓・間食。袋は必須。手順をキャッチに書くな。", ("おやつ", "間食", "小腹", "罪悪感")),
    ("daily", "毎日の食卓。袋は必須。手順をキャッチに書くな。", ("毎日", "食卓を彩")),
    (
        "cook",
        "料理に足した場面。袋は必須。手順をキャッチに書くな。",
        ("料理", "トッピング", "アクセント", "化ける", "アレンジ", "パスタ", "サラダ", "冷奴", "リゾット"),
    ),
    ("rice", "ご飯のお供。袋は必須。手順をキャッチに書くな。", ("ご飯", "おにぎり", "お茶漬け", "常備菜", "ふりかけ")),
    ("quick", "すぐ一品の食卓。袋は必須。手順をキャッチに書くな。", ("忙しい", "一品", "時短", "開けてそのまま", "手間なし")),
    ("origin", "原料の静物。袋は残す。偽認証は描くな。", ("国産", "産地", "原料", "原産", "国内産", "素材")),
)
# 食品と分かったときだけ、場面ごとに材料から1つ拾う（長い語を先に）。
_SCENE_DISHES: Dict[str, Tuple[str, ...]] = {
    "drink": ("おつまみ", "肴"),
    "snack": ("おやつ", "間食"),
    "daily": ("食卓",),
    "cook": ("パスタ", "サラダ", "冷奴", "リゾット", "卵かけご飯", "卵"),
    "rice": ("おにぎり", "お茶漬け", "ふりかけ", "ご飯"),
    "quick": ("パスタ", "サラダ", "冷奴", "一品"),
}
_FOOD_SIGNAL = (
    "食品",
    "食材",
    "料理",
    "ご飯",
    "おつまみ",
    "おやつ",
    "おかず",
    "調味",
    "飲料",
    "菓子",
    "水産",
    "惣菜",
    "乾物",
    "ジャーキー",
    "トッピング",
)
_FOOD_UNIT_SCENE_BAN = (
    "厚切り",
    "厚削り",
    "しっとり",
    "食感",
    "生ハム",
    "柔らか",
    "硬い",
    "色味",
    "質感",
    "1片",
    "一片",
)
_SCENE_GENERIC = "場面はこの見出しに合わせる。他の見出しと同じ卓・同じ盛りにするな。"
_SCENE_MIX = (
    "組み合わせ: 袋と個品の並び順をそろえる（PART_i と IDENTITY_i）。包材は見本の後ろ。"
    "袋を並べて主役にしない。文に片方の種類があっても、個品は両方出す。"
)
_SCENE_NO_INVENT_WRAP = (
    "小皿の主役は開封後の中身（IDENTITY）。"
    "個包装の見本があるときだけ、皿の横か後ろに1〜2個置く（小皿に混ぜるな。見本どおり。"
    "カメラ最近傍は開封後の個品。個包装はそれより手前に出すな。"
    "画面左下（個包装金丸聖域）に個包装を置くな。見本に無い空気だまり禁止）。"
    "見本が無い個包装・小袋・包み紙は描くな。"
)

# 旧WRAP置き（小皿混ぜ／周り）が保存案に残ると新指示と同居し手前に出る。汎用で除去。
_OLD_WRAP_MIX_RES = (
    re.compile(
        r"個包装の見本があるときだけ皿の横[・･／/]?周りに[1１][〜～\-][2２]個置くか小皿に少し混ぜて出せ[。．]?"
    ),
    re.compile(
        r"個包装の見本があるときだけ皿の横か後ろに[1１][〜～\-][2２]個置くか小皿に少し混ぜて出せ[。．]?"
    ),
    re.compile(
        r"個包装の見本があるときだけ皿の横[・･／/]?周りに置くか小皿に少し混ぜて出せ[。．]?"
    ),
    re.compile(r"小皿に少し混ぜて出せ[。．]?"),
    re.compile(r"皿の横[・･／/]?周りに[1１]?[〜～\-]?[2２]?個?置くか[。．]?"),
)


def scrub_old_wrap_mix_ja(scene: str) -> str:
    """保存済み場面文から旧『混ぜて出せ／横・周り』を落とす（案件固定なし）。"""
    s = str(scene or "")
    if not s:
        return ""
    for pat in _OLD_WRAP_MIX_RES:
        s = pat.sub("", s)
    s = re.sub(r"[。．]{2,}", "。", s)
    s = re.sub(r"[ \t]{2,}", " ", s)
    return s.strip()


def _proposal_text(row: Dict[str, Any]) -> str:
    return f"{row.get('catch') or ''} {row.get('evidence') or ''} {row.get('sceneHint') or ''}"


def _is_food_materials(materials: Dict[str, Any]) -> bool:
    return _has_any(_mat_blob(materials), _FOOD_SIGNAL)


def _scene_key_from_parts(head: str, text: str) -> str:
    """見出しの場面を先に見る。本文の国産・ご飯で別場面に吸わせない。"""
    head_s = str(head or "")
    text_s = str(text or "")
    for key, _hint, marks in _SCENE_MARKS:
        if any(m in head_s for m in marks):
            return key
    for key, _hint, marks in _SCENE_MARKS:
        if key == "origin":
            continue
        if any(m in text_s for m in marks):
            return key
    if any(m in text_s for m in _SCENE_MARKS[-1][2]):
        return "origin"
    return "benefit"


def _scene_hint_for_key(key: str, *, fallback: str = "") -> str:
    for k, hint, _marks in _SCENE_MARKS:
        if k == key:
            return hint
    fb = str(fallback or "").strip()
    if fb and not any(b in fb for b in _FOOD_UNIT_SCENE_BAN):
        return fb
    return "買うと良い状態（材料の語）。手順をキャッチに書くな。袋は必須。"


def _pick_dish_for_scene(
    key: str,
    materials: Dict[str, Any],
    used: set[str],
) -> str:
    blob = _mat_blob(materials)
    for dish in _SCENE_DISHES.get(key) or ():
        if dish in used:
            continue
        if dish in blob:
            return dish
    return ""


def _apply_scene_hint(row: Dict[str, Any]) -> Dict[str, Any]:
    """見出しの場面の語だけを描く条件にする。個品の語は足さない。"""
    out = dict(row)
    tid = str(out.get("themeId") or "")
    cid = str(out.get("classId") or "")
    if cid == "name" or tid == "N1":
        out["sceneHint"] = "袋が大きく読める。シズルは小さくてよい。"
        out["sceneKey"] = "name"
        return out
    if cid == "proof" or tid == "T1":
        out["sceneHint"] = "袋＋シズル。トロフィーは作らない。根拠のある数字のみ。"
        out["sceneKey"] = "proof"
        return out
    if cid == "origin" or tid == "T2":
        out["sceneKey"] = "origin"
        out["sceneHint"] = _scene_hint_for_key("origin")
        return out
    head = str(out.get("evidence") or "")
    m = re.search(r"【([^】]+)】", str(out.get("catch") or ""))
    if m:
        head = m.group(1) or head
    key = _scene_key_from_parts(head, _proposal_text(out))
    out["sceneKey"] = key
    out["sceneHint"] = _scene_hint_for_key(key, fallback=str(out.get("sceneHint") or ""))
    return out


def _enrich_scene_hints(
    rows: List[Dict[str, Any]],
    materials: Dict[str, Any],
    *,
    mix: bool = False,
) -> List[Dict[str, Any]]:
    """共通: 場面を分ける。食品のときだけ材料の食べ方を1つ。組み合わせは両方出す。"""
    food = _is_food_materials(materials)
    used_dishes: set[str] = set()
    out: List[Dict[str, Any]] = []
    for row in rows:
        r = dict(row)
        key = str(r.get("sceneKey") or "")
        bits = [str(r.get("sceneHint") or "").strip(), _SCENE_GENERIC]
        if food and key not in ("name", "proof", "origin", ""):
            dish = _pick_dish_for_scene(key, materials, used_dishes)
            if dish:
                used_dishes.add(dish)
                bits.append(f"材料にある食べ方（{dish}）を出せ。材料に無い料理は出すな。")
        dish_part = str(r.get("dishPartName") or "").strip()
        if mix and food and dish_part:
            bits.append(
                f"添える料理に載せるのは構成「{dish_part}」だけ。他の構成は小皿の見本のみ。"
            )
        if mix:
            bits.append(_SCENE_MIX)
        if food:
            bits.append(_SCENE_NO_INVENT_WRAP)
        r["sceneHint"] = "。".join(x.rstrip("。") for x in bits if x)
        out.append(r)
    return out


def reflect_checked_scene_hints(
    rows: List[Dict[str, Any]],
    materials: Dict[str, Any],
    *,
    mix: bool = False,
    food: Optional[bool] = None,
) -> List[Dict[str, Any]]:
    """確定した文案は残す。食品の食べ方と、組み合わせの両方出しを場面へ足す。"""
    if food is None:
        food = _is_food_materials(materials)
    used: set[str] = set()
    out: List[Dict[str, Any]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        applied = _apply_scene_hint(row)
        key = str(applied.get("sceneKey") or "")
        hint = scrub_old_wrap_mix_ja(str(row.get("sceneHint") or "").strip())
        bits = [hint] if hint else []
        if _SCENE_GENERIC not in hint:
            bits.append(_SCENE_GENERIC)
        already = ("食べ方（" in hint) or ("材料にある食べ方" in hint)
        if food and not already and key not in ("name", "proof", "origin", ""):
            dish = _pick_dish_for_scene(key, materials, used)
            if dish:
                used.add(dish)
                bits.append(f"材料にある食べ方（{dish}）を出せ。材料に無い料理は出すな。")
        dish_part = str(row.get("dishPartName") or "").strip()
        if mix and food and dish_part and "載せるのは" not in hint:
            bits.append(
                f"添える料理に載せるのは構成「{dish_part}」だけ。他の構成は小皿の見本のみ。"
            )
        if mix and "両方の包材" not in hint:
            bits.append(_SCENE_MIX)
        if food and (
            ("小皿に混ぜるな" not in hint or "カメラ最近傍" not in hint)
            and "個包装を新たに描くな" not in hint
        ):
            bits.append(_SCENE_NO_INVENT_WRAP)
        merged = dict(row)
        merged["sceneKey"] = key
        merged["sceneHint"] = "。".join(x.rstrip("。") for x in bits if x)
        out.append(merged)
    return out


def _pick_scene_slots(rows: List[Dict[str, Any]], target: int) -> List[Dict[str, Any]]:
    """同じ場面は1枚。商品名・実績・素材を先に残し、あとは点数の高い別場面。"""
    def _rank(row: Dict[str, Any], key: str) -> Tuple[int, int]:
        dedicated = 0
        if key == "origin" and (
            str(row.get("classId") or "") == "origin" or str(row.get("themeId") or "") == "T2"
        ):
            dedicated = 1
        return (dedicated, int(row.get("sellScore") or 0))

    by_key: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        keyed = _apply_scene_hint(row)
        key = str(keyed.get("sceneKey") or "benefit")
        prev = by_key.get(key)
        if prev is None or _rank(keyed, key) > _rank(prev, key):
            by_key[key] = keyed
    prefer = ("name", "proof", "origin", "drink", "cook", "rice", "snack", "quick", "daily", "benefit")
    picked: List[Dict[str, Any]] = []
    seen_catch: set[str] = set()

    def take_key(key: str) -> bool:
        row = by_key.pop(key, None)
        if not row:
            return False
        ck = re.sub(r"\s+", "", str(row.get("catch") or ""))
        if ck and ck in seen_catch:
            return False
        if ck:
            seen_catch.add(ck)
        picked.append(row)
        return True

    for key in prefer:
        if len(picked) >= target:
            break
        take_key(key)
    rest = sorted(by_key.values(), key=lambda r: -int(r.get("sellScore") or 0))
    for row in rest:
        if len(picked) >= target:
            break
        ck = re.sub(r"\s+", "", str(row.get("catch") or ""))
        if ck and ck in seen_catch:
            continue
        if ck:
            seen_catch.add(ck)
        picked.append(row)
    return picked


_SALES_WEAK_END = set("にをがはでともへのやっょゅゃァィゥェォッャュョー、，")


def _trim_clause_body(body: str, max_n: float = 16) -> str:
    """セールストーク用。句読点まで。助詞・小書きで終わる途中文は捨てる。"""
    t = re.sub(r"\s+", "", str(body or "")).strip("。．.、， ")
    if not t:
        return ""

    def _clean(cand: str) -> str:
        c = str(cand or "").strip("。．.、， ")
        while c and c[-1] in _SALES_WEAK_END:
            c = c[:-1]
        if _zenkaku_len(c) < 4:
            return ""
        if c[-1] in _SALES_WEAK_END:
            return ""
        return c

    # 短い句読点切れを優先（長い「。」より先の「！」を使う）
    candidates: List[str] = []
    for sep in ("！", "？", "。", "、"):
        i = t.find(sep)
        if i < 4:
            continue
        cand = _clean(t[:i])
        if cand:
            candidates.append(cand)
    for cand in sorted(candidates, key=_zenkaku_len):
        if _zenkaku_len(cand) <= max_n:
            return cand
    if _zenkaku_len(t) <= max_n:
        return _clean(t)
    # 上限内で意味の切れ目が無ければ載せない（途中切れ禁止）
    return ""


def _catch_with_body(head: str, materials: Dict[str, Any]) -> Tuple[str, str]:
    """材料の【見出し】とその直後の文。文は変えない。無いときは見出しだけ。"""
    h = str(head or "").strip()
    if not h:
        return "", ""
    blob_parts = [
        str(materials.get("rakuten") or ""),
        str(materials.get("yahoo") or ""),
        str(materials.get("desc") or ""),
        "\n".join(str(x) for x in (materials.get("bullets") or [])),
    ]
    body = ""
    for src in blob_parts:
        m = re.search(r"【\s*" + re.escape(h) + r"\s*】([^【]{0,120})", src)
        if not m:
            continue
        body = re.sub(r"\s+", "", m.group(1)).strip("。．.、， ")
        break
    catch_head = _short_catch(f"【{h}】", h)
    if not catch_head.startswith("【"):
        catch_head = f"【{re.sub(r'^【|】$', '', catch_head)}】"
    body = _trim_clause_body(body, max_n=16)
    if body:
        return f"{catch_head}{body}", h
    return catch_head, h


def _benefit_snippets(materials: Dict[str, Any]) -> List[Tuple[str, int]]:
    out: List[Tuple[str, int]] = []
    seen: set[str] = set()

    def add(text: str, pts: int) -> None:
        t = str(text or "").strip()
        if not t or len(t) < 2:
            return
        if any(b in t for b in _SELL_BAN) or _origin_only(t):
            return
        if any(b in t for b in _WEAK_HEAD_BAN):
            return
        name = str(materials.get("name") or "").strip()
        if name and re.sub(r"\s+", "", t) == re.sub(r"\s+", "", name):
            return
        key = re.sub(r"\s+", "", t)
        if key in seen:
            return
        seen.add(key)
        out.append((t, pts))

    def add_heads(src: str, pts: int) -> None:
        for m in re.finditer(r"【([^】]{2,20})】", str(src or "")):
            add(m.group(1), pts)

    add_heads(str(materials.get("rakuten") or ""), 12)
    add_heads(str(materials.get("yahoo") or ""), 8)
    for i, b in enumerate(materials.get("bullets") or []):
        pts = 6 if i < 2 else 3
        add_heads(str(b), pts)
        bs = str(b or "").strip()
        if bs and "【" not in bs:
            add(_zenkaku_slice(bs, 18), pts)
    add_heads(str(materials.get("desc") or ""), 1)
    out.sort(key=lambda x: x[1], reverse=True)
    return out


def _benefit_rows_mixed(
    materials: Dict[str, Any],
    part_materials: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """組み合わせ: 構成ごとの見出しを交互に取る（片方の材料だけで埋めない）。"""
    lists = [_benefit_snippets(p) for p in part_materials if p]
    if not lists:
        lists = [_benefit_snippets(materials)]
    rows: List[Dict[str, Any]] = []
    bi = 0
    max_n = max((len(x) for x in lists), default=0)
    for i in range(max_n):
        for pi, lst in enumerate(lists):
            if i >= len(lst):
                continue
            ev4, pts = lst[i]
            src = part_materials[pi] if pi < len(part_materials) else materials
            catch_full, ev_head = _catch_with_body(ev4, src)
            rows.append(
                {
                    "themeId": "T4",
                    "themeName": "成果",
                    "catch": catch_full or _short_catch(ev4, ev4),
                    "sceneHint": "買うと良い状態（材料の語）。手順をキャッチに書くな。袋は必須。",
                    "evidence": ev_head or ev4,
                    "classId": f"ben_{bi}",
                    "sellScore": 40 + pts,
                    "partIndex": pi,
                }
            )
            bi += 1
    return rows


def propose_from_materials(
    materials: Dict[str, Any],
    *,
    part_materials: Optional[List[Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    """1商品の材料から門＋点数。件数は childCount（2未満なら4、上限12）。

    キャッチは材料の見出しのまま。同じ場面は1枚。描く条件に個品の語は入れない。
    part_materials が2件以上なら組み合わせ扱い（見出しは交互、両方を描く）。
    """
    parts = [p for p in (part_materials or []) if isinstance(p, dict) and p]
    mix = len(parts) >= 2
    blob = _mat_blob(materials)
    name = str(materials.get("name") or "商品")
    raw: List[Dict[str, Any]] = []
    raw.append(
        {
            "themeId": "N1",
            "themeName": "商品名明示",
            "catch": _short_catch(name, name),
            "sceneHint": "袋が大きく読める。シズルは小さくてよい。",
            "evidence": name,
            "classId": "name",
            "sellScore": 10,
        }
    )
    if _has_any(blob, _T1_MARK):
        ev = next((m for m in _T1_MARK if m in blob), "")
        raw.append(
            {
                "themeId": "T1",
                "themeName": "実績・受賞",
                "catch": _short_catch(ev, ev),
                "sceneHint": "袋＋シズル。トロフィーは作らない。根拠のある数字のみ。",
                "evidence": ev,
                "classId": "proof",
                "sellScore": 70,
            }
        )
    if _has_any(blob, _T2_MARK):
        pr, ev = _prominence(materials, _T2_MARK)
        # 「安心」は材料に無い修飾なので付けない。事実の引用だけ。
        catch_t2 = _short_catch(ev or "国産", "国産")
        raw.append(
            {
                "themeId": "T2",
                "themeName": "産地・素材",
                "catch": catch_t2,
                "sceneHint": "原料の静物。袋は残す。偽認証は描くな。",
                "evidence": ev or next((m for m in _T2_MARK if m in blob), ""),
                "classId": "origin",
                "sellScore": 30 + pr,
            }
        )
    if mix:
        raw.extend(_benefit_rows_mixed(materials, parts))
    else:
        for bi, (ev4, pts) in enumerate(_benefit_snippets(materials)):
            catch_full, ev_head = _catch_with_body(ev4, materials)
            raw.append(
                {
                    "themeId": "T4",
                    "themeName": "成果",
                    "catch": catch_full or _short_catch(ev4, ev4),
                    "sceneHint": "買うと良い状態（材料の語）。手順をキャッチに書くな。袋は必須。",
                    "evidence": ev_head or ev4,
                    "classId": f"ben_{bi}",
                    "sellScore": 40 + pts,
                }
            )

    child = int(materials.get("childCount") or 0)
    target = child if child >= 2 else 4
    target = min(max(target, 1), 12)
    used_theme: set[str] = set()
    out: List[Dict[str, Any]] = []
    for row in _enrich_scene_hints(_pick_scene_slots(raw, target), materials, mix=mix):
        row = _assign_unique_theme_id(dict(row), used_theme)
        row.pop("sceneKey", None)
        row.pop("partIndex", None)
        row["sellRank"] = len(out) + 1
        out.append(row)
    return out


_SKU_IN_PROMPT = re.compile(
    r"(?i)(ishi-|sanky-|sku|oya\d+|親SKU|子SKU|[A-Za-z]+-\d{10,})"
)


def _name_base_for_prompt(s: str) -> str:
    """絵の指示に品番を入れない。商品名ベースだけ残す。"""
    t = str(s or "").strip()
    if not t or _SKU_IN_PROMPT.search(t):
        return ""
    return t


def sets_from_proposals(proposals: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """案F（最大6案）を仮想SETS形にする。"""
    out: List[Dict[str, Any]] = []
    for i, row in enumerate(proposals, start=1):
        catch = str(row.get("catch") or "")
        ev = str(row.get("evidence") or "")
        hint = str(row.get("sceneHint") or "")
        out.append(
            {
                "id": f"R{i}",
                "themeId": str(row.get("themeId") or "T3"),
                "nameJa": str(row.get("themeName") or row.get("themeId") or f"案{i}"),
                "catch": catch,
                "catchBurn": str(row.get("catchBurn") or ""),
                "burnSkip": bool(row.get("burnSkip")),
                "burnSkipReason": str(row.get("burnSkipReason") or ""),
                "salesTalk": str(row.get("salesTalk") or ""),
                "salesBurn": str(row.get("salesBurn") or ""),
                "evidence": ev,
                "sellScore": row.get("sellScore"),
                "sellRank": row.get("sellRank"),
                "dishName": str(row.get("dishName") or row.get("dish") or ""),
                "dishPartName": str(row.get("dishPartName") or ""),
                "dishPartSku": str(row.get("dishPartSku") or ""),
                "dishPartJan": str(row.get("dishPartJan") or ""),
                "scene": (
                    f"{hint}\n"
                    + (
                        ""
                        if "見本を前" in hint
                        else "商品の見本を前に置く。包材は見本の後ろ。袋を並べて主役にしない。"
                        "組み合わせは袋と個品の並び順をそろえる。顔と手渡す手元は出さない。"
                        "描く物はその案の状況の場だけ。選んだ食べ方と別の食べ物は足さない。"
                        "状況の文は画像の文字にしない。\n"
                    )
                    + "金丸・ランキング・人物の顔は描くな。見本に無い形・厚みを描くな。"
                ),
            }
        )
    return out


def _text_cells_over_cap(grid: Dict[str, Any], *, max_cells: int = P3_TEXT_MAX_CELLS) -> bool:
    """楽天文字20%: 掛かるマス合計が max_cells 超なら再作成。"""
    hit = int(grid.get("hitN") or 0)
    return hit > max_cells


def _checkerboard_suspected(im: Image.Image) -> bool:
    """金丸周辺に透過プレビューの市松が焼き込まれていないか。"""
    rgb = im.convert("RGB")
    px = rgb.load()
    gx, gy, gd = _gold_params()
    r2 = (gd / 2.0 + 8.0) ** 2
    n = 0
    for y in range(0, 240, 2):
        for x in range(960, CANVAS, 2):
            if (x - gx) ** 2 + (y - gy) ** 2 <= r2:
                continue
            if x + 2 >= CANVAS:
                continue
            r, g, b = px[x, y]
            if abs(r - g) >= 14 or abs(g - b) >= 14 or not (70 <= r <= 220):
                continue
            r2p, g2, b2 = px[x + 2, y]
            if abs(r2p - g2) >= 14 or abs(g2 - b2) >= 14:
                continue
            if abs(r - r2p) >= 35:
                n += 1
                if n >= 40:
                    return True
    return False


P3_CATCH_Y_SCAN = 500
P3_CATCH_X_MAX = 880
P3_CATCH_ROI_ROWS = 4
P3_CATCH_ROI_COLS = 7
P3_VERT_ROI_ROWS = 10
P3_VERT_ROI_COLS = 3
P3_CATCH_TARGET_CELLS = 16
P3_INK_SEED_MIN = 10
P3_INK_EXPAND_MIN = 4
P3_INK_DILATED_MIN = 10
P3_INK_DILATED_MIN_INK = 3
P3_DECO_CELL_MIN = 40
P3_ANCHOR_INK_MIN = 4
P3_ANCHOR_DIL_MIN = 8
GRID_MEASURE_ALGO = "catch_v8b"


def _p3_lum(r: int, g: int, b: int) -> float:
    return 0.299 * r + 0.587 * g + 0.114 * b


def _p3_text_zones(
    catch_head: str,
    sales_talk: str,
) -> Tuple[int, Tuple[int, ...], int, Tuple[int, ...]]:
    """キャッチ／セール文字長から ROI 内の期待帯域（行・列）を推定。"""
    catch_head = (catch_head or "").strip()
    sales_talk = (sales_talk or "").strip()
    catch_cols = (
        min(P3_CATCH_ROI_COLS, max(3, round(len(catch_head) * 1.3) + 1))
        if catch_head
        else 0
    )
    sales_cols = (
        min(P3_CATCH_ROI_COLS, max(3, round(len(sales_talk) * 0.55) + 1))
        if sales_talk
        else 0
    )
    catch_rows: Tuple[int, ...] = (0, 1, 2) if catch_head else ()
    if sales_talk:
        sales_rows = (2, 3) if len(sales_talk) <= 10 else (1, 2, 3)
    else:
        sales_rows = ()
    return catch_cols, catch_rows, sales_cols, sales_rows


def _p3_expected_zone_cells(
    catch_cols: int,
    catch_rows: Tuple[int, ...],
    sales_cols: int,
    sales_rows: Tuple[int, ...],
) -> set[Tuple[int, int]]:
    """文字長から推定した期待帯域（UI 点線オレンジ用）。"""
    cells: set[Tuple[int, int]] = set()
    for cy in catch_rows:
        for cx in range(catch_cols):
            cells.add((cx, cy))
    for cy in sales_rows:
        for cx in range(sales_cols):
            cells.add((cx, cy))
    return cells


def _p3_combine_ink_grids(
    a: List[List[int]],
    b: List[List[int]],
) -> List[List[int]]:
    return [[a[cy][cx] + b[cy][cx] for cx in range(10)] for cy in range(10)]


def _p3_anchor_fill(
    hit: set[Tuple[int, int]],
    ink: List[List[int]],
    dilated: List[List[int]],
    deco: List[List[int]],
    *,
    zone_fn,
) -> set[Tuple[int, int]]:
    """BFS 結果のアンカー列幅内で、連結を保ったまま ink/deco を補完。"""
    if not hit:
        return hit
    xs = [cx for cx, _ in hit]
    ys = [cy for _, cy in hit]
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    out = set(hit)
    changed = True
    while changed:
        changed = False
        for cy in range(y0, y1 + 1):
            for cx in range(x0, x1 + 1):
                if (cx, cy) in out:
                    continue
                if not zone_fn(cx, cy):
                    continue
                if not any((cx + dx, cy + dy) in out for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1))):
                    continue
                ink_n = ink[cy][cx]
                dil_n = dilated[cy][cx]
                deco_n = deco[cy][cx]
                if ink_n >= P3_ANCHOR_INK_MIN or (
                    dil_n >= P3_ANCHOR_DIL_MIN and ink_n >= 2
                ):
                    out.add((cx, cy))
                    changed = True
                elif deco_n >= P3_DECO_CELL_MIN and any(
                    (cx + dx, cy + dy) in out and ink[cy + dy][cx + dx] >= P3_ANCHOR_INK_MIN
                    for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1))
                ):
                    out.add((cx, cy))
                    changed = True
    return out


def _p3_in_text_zone(
    cx: int,
    cy: int,
    catch_cols: int,
    catch_rows: Tuple[int, ...],
    sales_cols: int,
    sales_rows: Tuple[int, ...],
) -> bool:
    if catch_rows and cy in catch_rows and cx < catch_cols:
        return True
    if sales_rows and cy in sales_rows and cx < sales_cols:
        return True
    return False


def _p3_is_catch_accent(r: int, g: int, b: int) -> bool:
    """後載せ吹き出し・帯・下線の accent 色（木目・食品オレンジは除外）。"""
    lum = _p3_lum(r, g, b)
    chroma = max(r, g, b) - min(r, g, b)
    if chroma < 45 or lum < 50 or lum > 245:
        return False
    if r > g > b and chroma < 80 and lum < 200:
        return False
    if r > 180 and g > 100 and b < 80:
        return False
    if b > r + 15 and b > g:
        return True
    if g > r + 10 and g > b:
        return True
    if r > 120 and g < 80 and b < 80:
        return True
    return chroma >= 70 and not (r > g > b)


def _p3_is_catch_ink(r: int, g: int, b: int, nmax: float) -> bool:
    """えんじ系キャッチ stroke。"""
    lum = _p3_lum(r, g, b)
    contrast = nmax - lum
    if contrast < 28:
        return False
    return r > 105 and g < 92 and b < 92 and r > max(g, b) + 20 and contrast >= 32


def _p3_is_sales_ink(r: int, g: int, b: int, nmax: float) -> bool:
    """黒／紫系セールストーク stroke。"""
    lum = _p3_lum(r, g, b)
    contrast = nmax - lum
    if contrast < 28:
        return False
    if lum <= 48 and contrast >= 30:
        return True
    return lum <= 58 and b > r + 8 and contrast >= 32


def _p3_is_strict_ink(r: int, g: int, b: int, nmax: float) -> bool:
    """文字情報なし時のフォールバック。"""
    return _p3_is_catch_ink(r, g, b, nmax) or _p3_is_sales_ink(r, g, b, nmax)


def _p3_is_cream_deco(r: int, g: int, b: int) -> bool:
    """吹き出し・帯のクリーム地（インク隣接時のみセル加算）。"""
    lum = _p3_lum(r, g, b)
    chroma = max(r, g, b) - min(r, g, b)
    return lum >= 192 and chroma <= 38 and min(r, g, b) >= 168


def _p3_dilate_ink(
    mask: List[List[bool]],
    *,
    w: int,
    h: int,
    radius: int = 1,
) -> List[List[bool]]:
    out = [[False] * w for _ in range(h)]
    for y in range(h):
        for x in range(w):
            if not mask[y][x]:
                continue
            for dy in range(-radius, radius + 1):
                ny = y + dy
                if ny < 0 or ny >= h:
                    continue
                for dx in range(-radius, radius + 1):
                    nx = x + dx
                    if nx < 0 or nx >= w:
                        continue
                    out[ny][nx] = True
    return out


def _p3_text_check(
    hit_catch: set[Tuple[int, int]],
    hit_sales: set[Tuple[int, int]],
    *,
    catch_head: str,
    sales_talk: str,
    catch_cols: int,
    catch_rows: Tuple[int, ...],
    sales_cols: int,
    sales_rows: Tuple[int, ...],
    writing_mode: str = "",
) -> Dict[str, Any]:
    """ proposals のキャッチ／セールと検出ブロック（アンカー幅）を照合。"""
    catch_head = (catch_head or "").strip()
    sales_talk = (sales_talk or "").strip()

    def _span(cells: set[Tuple[int, int]]) -> Tuple[int, int]:
        if not cells:
            return 0, 0
        xs = [cx for cx, _ in cells]
        ys = [cy for _, cy in cells]
        return max(xs) - min(xs) + 1, max(ys) - min(ys) + 1

    catch_w, catch_h = _span(hit_catch)
    sales_w, sales_h = _span(hit_sales)
    is_vertical = (writing_mode or "").strip() == "vertical-rl"
    col_cap = P3_VERT_ROI_COLS if is_vertical else P3_CATCH_ROI_COLS
    catch_ok = True
    sales_ok = True
    if catch_head:
        if not hit_catch:
            catch_ok = False
        elif is_vertical:
            catch_ok = 1 <= catch_w <= col_cap and catch_h >= 1
        else:
            min_w = max(1, (len(catch_head) + 2) // 3)
            max_w = min(P3_CATCH_ROI_COLS, len(catch_head) + 1)
            catch_ok = min_w <= catch_w <= max_w and len(hit_catch) >= 1
    if sales_talk:
        if not hit_sales:
            sales_ok = False
        elif is_vertical:
            sales_ok = 1 <= sales_w <= col_cap and sales_h >= 1
        else:
            min_w = max(1, (len(sales_talk) + 3) // 4)
            max_w = min(P3_CATCH_ROI_COLS, (len(sales_talk) + 1) // 2 + 1)
            sales_ok = min_w <= sales_w <= max_w and len(hit_sales) >= 1
    return {
        "catchHead": catch_head,
        "salesTalk": sales_talk,
        "catchCells": len(hit_catch),
        "salesCells": len(hit_sales),
        "catchSpanW": catch_w,
        "salesSpanW": sales_w,
        "catchOk": catch_ok,
        "salesOk": sales_ok,
        "textOk": catch_ok and sales_ok,
        "roiRows": P3_CATCH_ROI_ROWS,
        "roiCols": P3_CATCH_ROI_COLS,
        "catchZoneCols": catch_cols,
        "salesZoneCols": sales_cols,
    }


def _p3_cell_has_ink(
    ink_n: int,
    dil_n: int,
    *,
    seed: bool = False,
) -> bool:
    thresh = P3_INK_SEED_MIN if seed else P3_INK_EXPAND_MIN
    if ink_n >= thresh:
        return True
    return dil_n >= P3_INK_DILATED_MIN and ink_n >= P3_INK_DILATED_MIN_INK


def _p3_pass_bfs(
    ink: List[List[int]],
    dilated: List[List[int]],
    deco: List[List[int]],
    seeds: set[Tuple[int, int]],
    *,
    zone_fn,
    max_cx: int = P3_CATCH_ROI_COLS,
    max_cy: int = P3_CATCH_ROI_ROWS,
) -> set[Tuple[int, int]]:
    """1 pass: seed から 4 近傍 BFS。ink 実体＋隣接 cream のみ。矩形塗りなし。"""
    hit: set[Tuple[int, int]] = set()
    queue = list(seeds)
    while queue:
        cx, cy = queue.pop()
        if (cx, cy) in hit:
            continue
        if cx < 0 or cy < 0 or cx >= max_cx or cy >= max_cy:
            continue
        if not zone_fn(cx, cy):
            continue
        ink_n = ink[cy][cx]
        dil_n = dilated[cy][cx]
        deco_n = deco[cy][cx]
        is_seed = (cx, cy) in seeds
        if is_seed:
            if not _p3_cell_has_ink(ink_n, dil_n, seed=True):
                continue
        elif _p3_cell_has_ink(ink_n, dil_n):
            pass
        elif deco_n >= P3_DECO_CELL_MIN:
            ink_adj = False
            for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
                nx, ny = cx + dx, cy + dy
                if (nx, ny) in hit and ink[ny][nx] >= P3_INK_EXPAND_MIN:
                    ink_adj = True
                    break
            if not ink_adj:
                continue
        else:
            continue
        hit.add((cx, cy))
        for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
            nx, ny = cx + dx, cy + dy
            if (nx, ny) not in hit:
                queue.append((nx, ny))
    return hit


def _p3_vertical_column_seeds(
    ink: List[List[int]],
    dilated: List[List[int]],
    zone_fn,
    *,
    max_cy: int = P3_VERT_ROI_ROWS,
    cols: Optional[Iterable[int]] = None,
) -> set[Tuple[int, int]]:
    col_list = sorted(set(cols)) if cols is not None else list(range(P3_VERT_ROI_COLS))
    seeds: set[Tuple[int, int]] = set()
    for cx in col_list:
        if cx < 0 or cx >= 10:
            continue
        for cy in range(max_cy):
            if not zone_fn(cx, cy):
                continue
            if _p3_cell_has_ink(ink[cy][cx], dilated[cy][cx], seed=True):
                seeds.add((cx, cy))
                break
    return seeds


def _p3_rect_seeds(
    ink: List[List[int]],
    dilated: List[List[int]],
    cells: Set[Tuple[int, int]],
) -> set[Tuple[int, int]]:
    return {
        (cx, cy)
        for cx, cy in cells
        if _p3_cell_has_ink(ink[cy][cx], dilated[cy][cx], seed=True)
    }


def measure_grid20_p3(
    img_path: Path,
    *,
    catch_head: str = "",
    sales_talk: str = "",
    writing_mode: str = "",
    guide_cells: Optional[Iterable[Tuple[int, int]]] = None,
    guide_pattern: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """10×10 後載せキャッチ＋セールストーク＋吹き出し装飾のマス計測。

    catch_v8b: writingMode 対応 ROI 走査。検出（赤/紫）は走査 ROI 全体、緑ガイドはクリップしない。
    """
    is_vertical = (writing_mode or "").strip() == "vertical-rl"
    scan_cols = P3_VERT_ROI_COLS if is_vertical else P3_CATCH_ROI_COLS
    scan_rows = P3_VERT_ROI_ROWS if is_vertical else P3_CATCH_ROI_ROWS
    bfs_max_cx = scan_cols
    bfs_max_cy = scan_rows

    def scan_roi(cx: int, cy: int) -> bool:
        return 0 <= cx < scan_cols and 0 <= cy < scan_rows

    def roi_seeds(
        ink: List[List[int]],
        dilated: List[List[int]],
        zone_fn,
    ) -> set[Tuple[int, int]]:
        if is_vertical:
            return _p3_vertical_column_seeds(
                ink,
                dilated,
                zone_fn,
                max_cy=bfs_max_cy,
                cols=list(range(scan_cols)),
            )
        return {
            (cx, cy)
            for cy in range(scan_rows)
            for cx in range(scan_cols)
            if zone_fn(cx, cy)
            and _p3_cell_has_ink(ink[cy][cx], dilated[cy][cx], seed=True)
        }

    im = Image.open(img_path).convert("RGB").resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
    w, h = im.size
    cw, ch = w / 10.0, h / 10.0
    px = im.load()
    catch_ink = [[0] * 10 for _ in range(10)]
    sales_ink = [[0] * 10 for _ in range(10)]
    catch_dil = [[0] * 10 for _ in range(10)]
    sales_dil = [[0] * 10 for _ in range(10)]
    deco = [[0] * 10 for _ in range(10)]
    roi_x1 = int(scan_cols * cw)
    roi_y1 = int(scan_rows * ch)
    gx, gy, gd = _gold_params()
    gold_r2 = (gd / 2.0 + 4.0) ** 2
    ox0, oy0, ox1, oy1 = _octas_keepout_bbox()
    bx0, by0, bx1, by1 = _badge_bbox_1200()
    catch_head = (catch_head or "").strip()
    sales_talk = (sales_talk or "").strip()
    catch_cols, catch_rows, sales_cols, sales_rows = _p3_text_zones(
        catch_head, sales_talk
    )

    catch_mask = [[False] * roi_x1 for _ in range(roi_y1)]
    sales_mask = [[False] * roi_x1 for _ in range(roi_y1)]

    def lum_at(x: int, y: int) -> float:
        r, g, b = px[x, y]
        return _p3_lum(r, g, b)

    for y in range(0, roi_y1):
        for x in range(0, roi_x1):
            if (x - gx) ** 2 + (y - gy) ** 2 <= gold_r2:
                continue
            if ox0 <= x <= ox1 and oy0 <= y <= oy1:
                continue
            if bx0 <= x <= bx1 and by0 <= y <= by1:
                continue
            cx = min(9, int(x / cw))
            cy = min(9, int(y / ch))
            if cx >= scan_cols or cy >= scan_rows:
                continue
            r, g, b = px[x, y]
            nmax = max(
                lum_at(max(0, x - 1), y),
                lum_at(min(w - 1, x + 1), y),
                lum_at(x, max(0, y - 1)),
                lum_at(x, min(h - 1, y + 1)),
            )
            if _p3_is_catch_ink(r, g, b, nmax):
                catch_ink[cy][cx] += 1
                catch_mask[y][x] = True
            elif _p3_is_sales_ink(r, g, b, nmax):
                sales_ink[cy][cx] += 1
                sales_mask[y][x] = True
            elif _p3_is_cream_deco(r, g, b):
                deco[cy][cx] += 1

    for mask, dil_grid in (
        (catch_mask, catch_dil),
        (sales_mask, sales_dil),
    ):
        dil = _p3_dilate_ink(mask, w=roi_x1, h=roi_y1, radius=1)
        for y in range(roi_y1):
            for x in range(roi_x1):
                if not dil[y][x]:
                    continue
                cx = min(9, int(x / cw))
                cy = min(9, int(y / ch))
                if cx < scan_cols and cy < scan_rows:
                    dil_grid[cy][cx] += 1

    def catch_zone(cx: int, cy: int) -> bool:
        return scan_roi(cx, cy)

    def sales_zone(cx: int, cy: int) -> bool:
        return scan_roi(cx, cy)

    hit_catch: set[Tuple[int, int]] = set()
    hit_sales: set[Tuple[int, int]] = set()

    if catch_head:
        catch_seeds = roi_seeds(catch_ink, catch_dil, scan_roi)
        hit_catch = _p3_pass_bfs(
            catch_ink,
            catch_dil,
            deco,
            catch_seeds,
            zone_fn=catch_zone,
            max_cx=bfs_max_cx,
            max_cy=bfs_max_cy,
        )
        hit_catch = _p3_anchor_fill(
            hit_catch, catch_ink, catch_dil, deco, zone_fn=catch_zone
        )

    if sales_talk:
        sales_seeds = roi_seeds(sales_ink, sales_dil, scan_roi)
        hit_sales = _p3_pass_bfs(
            sales_ink,
            sales_dil,
            deco,
            sales_seeds,
            zone_fn=sales_zone,
            max_cx=bfs_max_cx,
            max_cy=bfs_max_cy,
        )
        hit_sales = _p3_anchor_fill(
            hit_sales, sales_ink, sales_dil, deco, zone_fn=sales_zone
        )

    if not catch_head and not sales_talk:
        combined = _p3_combine_ink_grids(catch_ink, sales_ink)
        combined_dil = _p3_combine_ink_grids(catch_dil, sales_dil)
        seeds = roi_seeds(combined, combined_dil, scan_roi)
        hit_all = _p3_pass_bfs(
            combined,
            combined_dil,
            deco,
            seeds,
            zone_fn=scan_roi,
            max_cx=bfs_max_cx,
            max_cy=bfs_max_cy,
        )
        hit_all = _p3_anchor_fill(
            hit_all, combined, combined_dil, deco, zone_fn=scan_roi
        )
        hit_catch = hit_all

    expected = _p3_expected_zone_cells(
        catch_cols, catch_rows, sales_cols, sales_rows
    )
    hit = hit_catch | hit_sales
    hit_n = len(hit)
    over16 = max(0, hit_n - P3_CATCH_TARGET_CELLS)
    over24 = max(0, hit_n - P3_TEXT_MAX_CELLS)
    text_check = _p3_text_check(
        hit_catch,
        hit_sales,
        catch_head=catch_head,
        sales_talk=sales_talk,
        catch_cols=catch_cols,
        catch_rows=catch_rows,
        sales_cols=sales_cols,
        sales_rows=sales_rows,
        writing_mode=writing_mode,
    )
    ink_roi = [
        [
            catch_ink[cy][cx] + sales_ink[cy][cx]
            for cx in range(scan_cols)
        ]
        for cy in range(scan_rows)
    ]
    return {
        "file": Path(img_path).name,
        "algo": GRID_MEASURE_ALGO,
        "writingMode": writing_mode or None,
        "hitN": hit_n,
        "over16": over16,
        "over24": over24,
        "ok20": hit_n <= P3_TEXT_MAX_CELLS,
        "ok16": hit_n <= P3_CATCH_TARGET_CELLS,
        "ok25": hit_n <= 25,
        "cells": [list(p) for p in sorted(hit)],
        "catchCells": [list(p) for p in sorted(hit_catch)],
        "salesCells": [list(p) for p in sorted(hit_sales)],
        "expectedZone": [list(p) for p in sorted(expected)],
        "textCheck": text_check,
        "inkGrid": ink_roi,
        "scanCols": scan_cols,
        "scanRows": scan_rows,
        "detectClip": "scan_roi",
        "promptMode": "adopted_only" if guide_pattern else "soft_roi",
    }


def _write_grid20_p3(img_path: Path, dest: Path) -> Dict[str, Any]:
    """③用 10×10。金丸円内は除外。後載せキャッチ（極白/極黒の強コントラスト）だけ数える。

    パッケージ印刷・木目・食品の中間トーンは対象外（公式も後載せ文字が対象）。
    """
    measured = measure_grid20_p3(img_path)
    hit = {tuple(p) for p in measured.get("cells") or []}
    im = Image.open(img_path).convert("RGB").resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
    w, h = im.size
    cw, ch = w / 10.0, h / 10.0
    overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    for cy in range(10):
        for cx in range(10):
            x0, y0 = int(cx * cw), int(cy * ch)
            x1, y1 = int((cx + 1) * cw) - 1, int((cy + 1) * ch) - 1
            if (cx, cy) in hit:
                d.rectangle([x0, y0, x1, y1], fill=(255, 40, 40, 90))
            d.rectangle([x0, y0, x1, y1], outline=(0, 255, 80, 170), width=1)
    mixed = Image.alpha_composite(im.convert("RGBA"), overlay).convert("RGB")
    dest.parent.mkdir(exist_ok=True)
    mixed.save(dest, quality=90)
    return {
        "file": measured.get("file") or img_path.name,
        "hitN": measured.get("hitN"),
        "ok20": measured.get("ok20"),
        "cells": measured.get("cells") or [],
    }


def _stamp_p2(
    base: Image.Image,
    *,
    upper: str,
    lower: str,
    stroke: Image.Image,
    color: Dict[str, Any],
) -> Image.Image:
    cfg = _appeal_cfg()
    u = cfg.get("upper") or {}
    lo = cfg.get("lower") or {}
    ul = cfg.get("underline") or {}
    sw = cfg.get("pillowSameColorStroke") or {}
    sh = cfg.get("pillowSoftShadow") or {}
    out = base.convert("RGBA")
    fill, fill_mode, luma = _headline_fill_rgba(out, cfg)
    LOG.info("P2 headline mode=%s rgb=%s luma=%.1f", fill_mode, fill[:3], luma)
    rgb = _hex_rgb(str(color.get("hex") or "#800020"))
    ink = _tint_stroke(stroke, rgb, alpha_scale=0.85)
    ul_h = int(ul.get("h", 76))
    ul_w = int(ul.get("w", 720))
    ul_x = int(ul.get("x", 52))
    ul_y = int(ul.get("y", 161))
    if ink.size[1] != ul_h:
        ratio = ul_h / float(ink.size[1])
        ink = ink.resize(
            (max(1, int(round(ink.size[0] * ratio))), ul_h),
            Image.Resampling.LANCZOS,
        )
    if ink.size[0] > ul_w:
        ink = ink.resize((ul_w, ul_h), Image.Resampling.LANCZOS)
    out.alpha_composite(ink, (ul_x, ul_y))
    font_id = str(cfg.get("fontId") or "ud_motoya_maru_berry")
    font_u = _load_font_px(font_id, int(u.get("fontPx") or 59))
    font_l = _load_font_px(font_id, int(lo.get("fontPx") or 134))
    shadow_rgba = sh.get("rgba") or [40, 30, 20, 70]
    if fill_mode == "light_bg":
        shadow = (40, 30, 20, 40)
    else:
        shadow = (int(shadow_rgba[0]), int(shadow_rgba[1]), int(shadow_rgba[2]), int(shadow_rgba[3]))
    shadow_off = int(sh.get("off") or 2)
    _draw_tracked_line(
        out,
        upper,
        x=int(u.get("x") or 48),
        y_ink=int(u.get("y") or 31),
        font=font_u,
        fill=fill,
        tracking_px=int(u.get("trackingPx") or 3),
        x_scale=float(u.get("xScale") or 1.0),
        shadow=shadow,
        shadow_off=shadow_off,
        stroke_width=int(sw.get("upper") or 1),
        stroke_fill=fill,
    )
    _draw_tracked_line(
        out,
        lower,
        x=int(lo.get("x") or 40),
        y_ink=int(lo.get("y") or 97),
        font=font_l,
        fill=fill,
        tracking_px=int(lo.get("trackingPx") or 8),
        x_scale=float(lo.get("xScale") or 0.858),
        shadow=shadow,
        shadow_off=shadow_off,
        stroke_width=int(sw.get("lower") or 2),
        stroke_fill=fill,
    )
    return out


def _sheet(paths: List[Path], labels: List[str], out_path: Path, *, thumb: int = 240) -> None:
    n = len(paths)
    pad = 14
    header = 44
    cols = n
    w = pad + cols * (thumb + pad)
    h = pad + header + thumb + 56 + pad
    sheet = Image.new("RGB", (w, h), (245, 245, 245))
    draw = ImageDraw.Draw(sheet)
    font = _load_font("genjyuu_gothic_like", 20)
    for i, (p, lab) in enumerate(zip(paths, labels)):
        im = Image.open(p).convert("RGB")
        im.thumbnail((thumb, thumb), Image.Resampling.LANCZOS)
        x = pad + i * (thumb + pad)
        y = pad + header
        sheet.paste(im, (x, y))
        draw.text((x, pad + 8), lab, fill=(20, 20, 20), font=font)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out_path, quality=92)


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser()
    ap.add_argument("--text-color", default="1")
    ap.add_argument(
        "--engines",
        default="openai",
        help="画像は openai のみ。gemini は拒否（画像生成しない）",
    )
    ap.add_argument(
        "--openai-quality",
        default=OPENAI_QUALITY_DEFAULT,
        choices=("low", "medium", "high"),
        help="サブ画像本線と同じ既定=medium",
    )
    ap.add_argument(
        "--reuse-p1",
        default="",
        help="①②を作るときだけ。既存ランの P1 流用先",
    )
    ap.add_argument(
        "--with-p1p2",
        action="store_true",
        help="①②も作る（APIコストあり）。既定は③のみ",
    )
    ap.add_argument(
        "--propose-only",
        action="store_true",
        help="画像を作らず、1商品の材料から最大4案＋点数だけ出す（案F試行）",
    )
    ap.add_argument(
        "--legacy-sets",
        action="store_true",
        help="旧・固定S1〜S5（かつお検証用）。既定は材料から最大4案",
    )
    args = ap.parse_args()
    if args.propose_only:
        mat = _compact_facts_from_sheets(JAN)
        materials = mat.get("materials") or {}
        if not materials:
            materials = {
                "name": "食べるおだし",
                "maker": "",
                "feats": LOCKED["facts"],
                "kw": "お茶漬け おつまみ",
                "yahoo": "だしをとらずにそのまま食べる",
                "rakuten": "だしをとらずさっと一振り。お茶漬け・おつまみ。",
                "bullets": ["だしをとらずにそのまま食べる", "お茶漬け・おつまみ"],
                "desc": LOCKED["facts"],
            }
        proposals = propose_from_materials(materials)
        work = Path(DEFAULT_WORK_ROOT)
        out_root = (
            work
            / SUB_TEST_OUT
            / "b5_main_base_virt"
            / f"{JAN}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_propose"
        )
        out_root.mkdir(parents=True, exist_ok=True)
        payload = {
            "jan": JAN,
            "error": mat.get("error") or "",
            "materials": materials,
            "proposals": proposals,
        }
        (out_root / "b5_proposals.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        LOG.info("propose-only out=%s error=%s", out_root, payload["error"])
        for row in proposals:
            LOG.info(
                "  #%s %s %s score=%s catch=%s ev=%s",
                row.get("sellRank"),
                row.get("themeId"),
                row.get("themeName"),
                row.get("sellScore"),
                row.get("catch"),
                row.get("evidence"),
            )
        return 0
    color = resolve_text_color(args.text_color)
    engines = [e.strip().lower() for e in str(args.engines or "").split(",") if e.strip()]
    for e in engines:
        if e == "gemini":
            raise SystemExit("Geminiで画像は作らない。--engines openai のみ。")
        if e not in ("openai",):
            raise SystemExit(f"不明な engine: {e}")
    if not engines:
        raise SystemExit("--engines が空")
    with_p1p2 = bool(args.with_p1p2)
    reuse_p1 = Path(args.reuse_p1) if str(args.reuse_p1 or "").strip() else None
    if reuse_p1 is not None and not reuse_p1.is_dir():
        raise SystemExit(f"--reuse-p1 が無い: {reuse_p1}")
    if with_p1p2:
        nu, nl = _zenkaku_len(LOCKED["upper"]), _zenkaku_len(LOCKED["lower"])
        if not (12 <= nu <= 18 and 4 <= nl <= 8):
            raise SystemExit(f"LOCKED 字数NG upper={nu} lower={nl}")

    work = Path(DEFAULT_WORK_ROOT)
    out_root = (
        work
        / SUB_TEST_OUT
        / "b5_main_base_virt"
        / f"{JAN}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_abc16"
    )
    out_root.mkdir(parents=True, exist_ok=True)
    meta = out_root / "_meta"
    meta.mkdir(exist_ok=True)
    # 人が開いて比較する代表フォルダ（金丸＋Octas済みの個別ファイルのみ）
    pick_dir = out_root / "比較"
    pick_dir.mkdir(exist_ok=True)

    truth = _truth_path(work)
    pkg = _load_pkg(truth)
    pkg.save(meta / "pkg_cutout.png")

    octas_path = resolve_octas(work)
    if octas_path is None or not octas_path.is_file():
        raise SystemExit("99.octas期限管理シール素材 なし")
    octas_src = Image.open(octas_path).convert("RGBA")

    gold = _load_gold_disk(work, _gold_params()[2])
    gold.save(meta / "gold_disk.png")

    stroke: Image.Image | None = None
    hook_sheet = None
    style_lock = _p3_style_lock_path()
    if with_p1p2:
        ul_src = resolve_underline_sample(work)
        if ul_src is None or not ul_src.is_file():
            raise SystemExit("98 下線素材なし")
        stroke = _trim_ink(_crop_underline_stroke(Image.open(ul_src), str(color.get("id") or "1")))
        stroke.save(meta / "underline_stroke.png")
        hooks = _hook_ref_paths()
        if hooks:
            hook_sheet = _write_style_sheet(hooks, meta / "hook_sheet.jpg")
        if any(str(v.get("mode") or "") == "lock" for v in P3_VARIANTS) and not style_lock.is_file():
            raise SystemExit(f"P3トンマナ見本なし: {style_lock}")

    typo_refs = _oshare_typo_paths(6)
    typo_sheet: Optional[Path] = None
    if typo_refs:
        typo_sheet = _write_style_sheet(typo_refs, meta / "style_typo_sheet.jpg")
        (meta / "style_typo_sources.json").write_text(
            json.dumps([str(p) for p in typo_refs], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        LOG.info("style_typo_sheet n=%s -> %s", len(typo_refs), typo_sheet)
    else:
        LOG.warning("おしゃれ文字装飾見本が空。IMAGE_STYLE_TYPOなしで続行")

    gemini_client = None
    gemini_model = ""
    openai_model = ""
    if "gemini" in engines:
        gemini_client = make_client()
        gemini_model = resolve_model_id(gemini_client)
    if "openai" in engines:
        openai_model = resolve_openai_image_model()
    LOG.info(
        "abc16 engines=%s with_p1p2=%s gemini=%s openai=%s quality=%s out=%s pick=%s",
        engines,
        with_p1p2,
        gemini_model or "-",
        openai_model or "-",
        args.openai_quality,
        out_root,
        pick_dir,
    )

    mat = _compact_facts_from_sheets(JAN)
    (meta / "facts.json").write_text(
        json.dumps(
            {
                "factsLine": mat.get("factsLine"),
                "masterKeys": list((mat.get("master") or {}).keys()),
                "aiKeys": list((mat.get("ai") or {}).keys()),
                "error": mat.get("error") or "",
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (meta / "facts_prompt.txt").write_text(str(mat.get("factsLine") or ""), encoding="utf-8")
    facts_line = str(mat.get("factsLine") or LOCKED["facts"])
    materials = mat.get("materials") or {}
    if args.legacy_sets:
        run_sets = SETS
        proposals: List[Dict[str, Any]] = []
    else:
        if not materials:
            materials = {
                "name": "食べるおだし",
                "maker": "",
                "feats": LOCKED["facts"],
                "kw": "お茶漬け おつまみ",
                "yahoo": "だしをとらずにそのまま食べる",
                "rakuten": "だしをとらずさっと一振り。お茶漬け・おつまみ。",
                "bullets": ["だしをとらずにそのまま食べる", "お茶漬け・おつまみ"],
                "desc": LOCKED["facts"],
            }
        proposals = propose_from_materials(materials)
        run_sets = sets_from_proposals(proposals)
        if not run_sets:
            raise SystemExit("材料から案が0件（門で全部落ちた）")
    (meta / "b5_proposals.json").write_text(
        json.dumps(
            {
                "jan": JAN,
                "error": mat.get("error") or "",
                "legacySets": bool(args.legacy_sets),
                "materials": materials,
                "proposals": proposals,
                "sets": [
                    {
                        "id": s["id"],
                        "themeId": s.get("themeId"),
                        "nameJa": s.get("nameJa"),
                        "catch": s.get("catch"),
                        "sellRank": s.get("sellRank"),
                    }
                    for s in run_sets
                ],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    LOG.info(
        "abc16 propose n=%s legacy=%s %s",
        len(run_sets),
        bool(args.legacy_sets),
        [(s.get("id"), s.get("themeId"), s.get("nameJa")) for s in run_sets],
    )
    (meta / "gold_keepout.txt").write_text(_gold_keepout_line(), encoding="utf-8")
    (meta / "octas_keepout.txt").write_text(_octas_keepout_line(), encoding="utf-8")
    (meta / "catch_placement.txt").write_text(_catch_placement_line(), encoding="utf-8")
    (meta / "catch_hero_size.txt").write_text(_catch_hero_size_line(), encoding="utf-8")
    (meta / "style_typo_prompt.txt").write_text(_style_typo_line(), encoding="utf-8")

    def _reuse_p1_path(engine: str, sid: str) -> Optional[Path]:
        if reuse_p1 is None:
            return None
        for cand in (
            reuse_p1 / engine / sid / "P1_base.jpg",
            reuse_p1 / sid / "P1_base.jpg",
        ):
            if cand.is_file():
                return cand
        return None

    def _pick_name(spec: Dict[str, Any], vid: str) -> str:
        raw = f"{spec['id']}_{spec.get('themeId')}_{spec['nameJa']}_{vid}"
        return re.sub(r'[\\/:*?"<>|\s]+', "_", raw) + ".jpg"

    summary_by_engine: Dict[str, Any] = {}
    paths_by_engine: Dict[str, Dict[str, Dict[str, Path]]] = {}

    for engine in engines:
        eng_root = out_root / engine
        eng_root.mkdir(exist_ok=True)
        grid_dir = eng_root / "_grid20"
        grid_dir.mkdir(exist_ok=True)
        eng_meta = eng_root / "_meta"
        eng_meta.mkdir(exist_ok=True)
        summary_sets: List[Dict[str, Any]] = []
        p3_lock_im: Image.Image | None = None
        paths_by_engine[engine] = {}

        for spec in run_sets:
            sid = spec["id"]
            sdir = eng_root / sid
            sdir.mkdir(exist_ok=True)
            LOG.info("[%s] set %s %s theme=%s", engine, sid, spec["nameJa"], spec.get("themeId"))

            g_bg: Dict[str, Any] = {}
            pkg_meta: Dict[str, Any] = {}
            p1_path: Optional[Path] = None
            p2_path: Optional[Path] = None
            g1s: Dict[str, Any] = {}
            g2s: Dict[str, Any] = {}
            if with_p1p2:
                src_p1 = _reuse_p1_path(engine, sid)
                if src_p1 is not None:
                    p1 = Image.open(src_p1).convert("RGBA")
                    if p1.size != (CANVAS, CANVAS):
                        p1 = p1.resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
                    p1_path = sdir / "P1_base.jpg"
                    p1.convert("RGB").save(p1_path, quality=92)
                    g_bg = {"apiPath": "reuse_p1", "engine": engine}
                    LOG.info("[%s] %s P1 reuse %s", engine, sid, src_p1)
                else:
                    if hook_sheet is None:
                        raise SystemExit("99.MAIN売れる構図参照 が無い（P1新規時は必須）")
                    scene_prompt = _prompt_scene(spec["scene"])
                    (eng_meta / f"{sid}_scene_prompt.txt").write_text(scene_prompt, encoding="utf-8")
                    raw_bg, g_bg = _gen_ai(
                        engine,
                        prompt=scene_prompt,
                        image_paths=[hook_sheet],
                        image_roles=["IMAGE_HOOK_SHEET"],
                        gemini_client=gemini_client,
                        gemini_model=gemini_model,
                        openai_model=openai_model,
                        openai_quality=str(args.openai_quality),
                    )
                    bg = _to_1200(raw_bg)
                    bg.convert("RGB").save(eng_meta / f"{sid}_scene.jpg", quality=90)
                    p1, pkg_meta = _place_package(bg, pkg)
                    p1_path = sdir / "P1_base.jpg"
                    p1.convert("RGB").save(p1_path, quality=92)
                    LOG.info("[%s] %s P1 ok", engine, sid)
                assert p1_path is not None and stroke is not None
                g1s = _write_grid20(p1_path, grid_dir / f"{sid}_P1_grid20.jpg")
                p1_gold = _stamp_gold(Image.open(p1_path).convert("RGBA"), gold)
                p2 = _stamp_p2(
                    p1_gold,
                    upper=LOCKED["upper"],
                    lower=LOCKED["lower"],
                    stroke=stroke,
                    color=color,
                )
                p2_path = sdir / "P2_locked.jpg"
                p2.convert("RGB").save(p2_path, quality=92)
                g2s = _write_grid20(p2_path, grid_dir / f"{sid}_P2_grid20.jpg")

            p3_slots: List[Dict[str, Any]] = []
            set_paths: Dict[str, Path] = {}
            if p1_path is not None:
                set_paths["P1"] = p1_path
            if p2_path is not None:
                set_paths["P2"] = p2_path

            for var in P3_VARIANTS:
                vid = var["id"]
                mode = str(var.get("mode") or "free")
                if mode == "lock":
                    p3_prompt = _prompt_p3_lock()
                    img_paths = [style_lock, truth]
                    img_roles = ["IMAGE_COMPOSE_LOCK", "IMAGE_PACKAGE_TRUTH"]
                else:
                    img_paths = [truth]
                    img_roles = ["IMAGE_PACKAGE_TRUTH"]
                    if typo_sheet is not None:
                        img_paths.append(typo_sheet)
                        img_roles.append("IMAGE_STYLE_TYPO")
                try:
                    if mode == "lock" and p3_lock_im is not None:
                        p3 = p3_lock_im.copy()
                        g3 = {"apiPath": "reuse_p3a_lock", "engine": engine}
                        p3_path = sdir / f"P3{vid}_{var['hook']}.jpg"
                        p3.convert("RGB").save(p3_path, quality=92)
                        g3s = _write_grid20_p3(p3_path, grid_dir / f"{sid}_P3{vid}_grid20.jpg")
                        (eng_meta / f"{sid}_P3{vid}_{mode}_prompt.txt").write_text(
                            p3_prompt, encoding="utf-8"
                        )
                        p3_slots.append(
                            {
                                "file": p3_path.name,
                                "hook": var["hook"],
                                "mode": mode,
                                "grid20": g3s,
                                "api": g3.get("apiPath"),
                                "gold": "pillow_stamp",
                            }
                        )
                        set_paths[f"P3{vid}"] = p3_path
                        pick_dest = pick_dir / _pick_name(spec, vid)
                        shutil.copy2(p3_path, pick_dest)
                        continue

                    retry_note = ""
                    tries: List[Dict[str, Any]] = []
                    best: Optional[Dict[str, Any]] = None
                    p3_path = sdir / f"P3{vid}_{var['hook']}.jpg"
                    max_tries = 1 if mode == "lock" else P3_FREE_MAX_TRIES
                    for attempt in range(1, max_tries + 1):
                        if mode == "free":
                            p3_prompt = _prompt_p3_free(
                                facts_line,
                                theme_id=str(spec.get("themeId") or ""),
                                name_ja=str(spec.get("nameJa") or sid),
                                scene=str(spec.get("scene") or ""),
                                retry_note=retry_note,
                                catch=str(spec.get("catch") or ""),
                                text_color=str(color.get("id") or args.text_color or "1"),
                            )
                        (eng_meta / f"{sid}_P3{vid}_{mode}_try{attempt}_prompt.txt").write_text(
                            p3_prompt, encoding="utf-8"
                        )
                        raw3, g3 = _gen_ai(
                            engine,
                            prompt=p3_prompt,
                            image_paths=img_paths,
                            image_roles=img_roles,
                            gemini_client=gemini_client,
                            gemini_model=gemini_model,
                            openai_model=openai_model,
                            openai_quality=str(args.openai_quality),
                        )
                        cand_im = _to_1200(raw3)
                        try_path = sdir / f"P3{vid}_{var['hook']}_try{attempt}.jpg"
                        cand_im.convert("RGB").save(try_path, quality=92)
                        grid_try = _write_grid20_p3(
                            try_path, grid_dir / f"{sid}_P3{vid}_try{attempt}_grid20.jpg"
                        )
                        hit_n = int(grid_try.get("hitN") or 0)
                        over = mode == "free" and _text_cells_over_cap(grid_try)
                        no_catch = mode == "free" and hit_n < 1
                        checker = _checkerboard_suspected(cand_im)
                        # 採用優先度: 市松・文字過多・キャッチ無しを避け、その中で hitN が小さい方
                        rank = (
                            (1 if checker else 0) * 1000
                            + (1 if over else 0) * 500
                            + (1 if no_catch else 0) * 200
                            + hit_n
                        )
                        tries.append(
                            {
                                "attempt": attempt,
                                "hitN": hit_n,
                                "ok20": not over,
                                "noCatch": no_catch,
                                "checker": checker,
                                "rank": rank,
                                "api": g3.get("apiPath"),
                                "model": g3.get("model"),
                                "file": try_path.name,
                            }
                        )
                        LOG.info(
                            "[%s] %s P3%s try=%s hitN=%s over=%s noCatch=%s checker=%s",
                            engine,
                            sid,
                            vid,
                            attempt,
                            hit_n,
                            over,
                            no_catch,
                            checker,
                        )
                        if best is None or rank < int(best["rank"]):
                            best = {
                                "im": cand_im,
                                "grid": grid_try,
                                "api": g3.get("apiPath"),
                                "model": g3.get("model"),
                                "hitN": hit_n,
                                "rank": rank,
                                "file": try_path.name,
                            }
                        if not over and not checker and not no_catch:
                            break
                        hero_half = max(1, int(_gold_params()[2]) // 2)
                        if no_catch:
                            retry_note = (
                                "前回は後載せキャッチが無かった。必ずちょうど1本載せろ"
                                "（全角8〜18字。材料の言い回しを短く）。0本は不可。"
                                "配置は上部・左寄り。下帯の長い横長は禁止。"
                                f"キーワード1語以上は文字高さ約{hero_half}px以上（金丸の約半分）。"
                                "右上金丸聖域・右下Octas聖域に文字を入れるな。"
                                "IMAGE_PACKAGE_TRUTHの袋は必ず画面に残せ（印字改変禁止）。"
                            )
                        else:
                            retry_note = (
                                f"前回は後載せ文字が{hit_n}マス（上限{P3_TEXT_MAX_CELLS}）。"
                                "短いキャッチ1つだけ。上部・左寄り。下帯の長い横長・多段・囲み禁止。"
                                f"主役キーワード1語以上は約{hero_half}px以上の高さ（金丸の約半分）。"
                                "右上金丸聖域・右下Octas聖域に文字を入れるな。"
                                "IMAGE_PACKAGE_TRUTHの袋は必ず画面に残せ（印字改変禁止）。"
                                + (" 市松・チェッカー模様を描くな。" if checker else "")
                            )
                    if best is None:
                        if not tries:
                            raise RuntimeError("P3 generate empty")
                        last_path = sdir / str(tries[-1]["file"])
                        best = {
                            "im": Image.open(last_path).convert("RGBA"),
                            "grid": _write_grid20_p3(
                                last_path, grid_dir / f"{sid}_P3{vid}_grid20.jpg"
                            ),
                            "api": tries[-1].get("api"),
                            "model": tries[-1].get("model"),
                            "hitN": tries[-1].get("hitN"),
                            "file": tries[-1]["file"],
                        }
                    # 金丸＋OctasはPillowで固定載せ（AIには渡さない）
                    p3 = _finish_p3_overlays(best["im"], gold, octas_src)
                    p3.convert("RGB").save(p3_path, quality=92)
                    g3s = _write_grid20_p3(p3_path, grid_dir / f"{sid}_P3{vid}_grid20.jpg")
                    if mode == "lock":
                        p3_lock_im = p3.copy()
                    p3_slots.append(
                        {
                            "file": p3_path.name,
                            "hook": var["hook"],
                            "mode": mode,
                            "grid20": g3s,
                            "api": best.get("api"),
                            "model": best.get("model"),
                            "gold": "pillow_stamp",
                            "octas": "pillow_br_1p3cell",
                            "tries": tries,
                            "pickedTry": best.get("file"),
                        }
                    )
                    set_paths[f"P3{vid}"] = p3_path
                    pick_dest = pick_dir / _pick_name(spec, vid)
                    shutil.copy2(p3_path, pick_dest)
                    LOG.info(
                        "[%s] %s P3%s ok picked=%s -> %s",
                        engine,
                        sid,
                        vid,
                        best.get("file"),
                        pick_dest.name,
                    )
                except Exception as e:
                    LOG.exception("[%s] %s P3%s fail", engine, sid, vid)
                    p3_slots.append({"error": str(e), "hook": var["hook"], "mode": mode})

            paths_by_engine[engine][sid] = set_paths
            summary_sets.append(
                {
                    "id": sid,
                    "themeId": spec.get("themeId"),
                    "nameJa": spec["nameJa"],
                    "export02": False,
                    "withP1P2": with_p1p2,
                    "gold": {"P3": "pillow_stamp"},
                    "octas": {"P3": "pillow_br_1p3cell"},
                    "P1": (
                        {
                            "file": str(p1_path.relative_to(out_root)),
                            "grid20": g1s,
                            "api": g_bg.get("apiPath"),
                        }
                        if p1_path is not None
                        else None
                    ),
                    "P2": (
                        {"file": str(p2_path.relative_to(out_root)), "grid20": g2s}
                        if p2_path is not None
                        else None
                    ),
                    "P3": p3_slots,
                    "pickFiles": [
                        _pick_name(spec, str(v["id"]))
                        for v in P3_VARIANTS
                        if (pick_dir / _pick_name(spec, str(v["id"]))).is_file()
                    ],
                }
            )

        summary_by_engine[engine] = {
            "engine": engine,
            "model": gemini_model if engine == "gemini" else openai_model,
            "openaiQuality": str(args.openai_quality) if engine == "openai" else None,
            "sets": summary_sets,
        }

    summary = {
        "jan": JAN,
        "plan": "A_abc_v14_theme_in_p3_pick_folder",
        "export02": False,
        "withP1P2": with_p1p2,
        "engines": engines,
        "sets": [
            {"id": s["id"], "themeId": s.get("themeId"), "nameJa": s["nameJa"]} for s in run_sets
        ],
        "models": {
            "gemini": gemini_model or None,
            "openai": openai_model or None,
            "openaiQuality": str(args.openai_quality),
        },
        "gold": {
            "P3": "keepout_then_pillow",
            "keepout": list(_badge_bbox_1200()),
            "circle": {"cx": _gold_params()[0], "cy": _gold_params()[1], "diameter": _gold_params()[2]},
        },
        "octas": {
            "P3": "keepout_then_pillow_br",
            "keepout": list(_octas_keepout_bbox()),
            "sizeMaxPx": OCTAS_BR_SIZE_MAX,
            "cellPx": OCTAS_BR_CELL,
            "scaleVsCell": 1.3,
            "marginPx": OCTAS_BR_MARGIN,
            "padPx": OCTAS_BR_KEEPOUT_PAD,
            "tiltDeg": OCTAS_BR_TILT_DEG,
        },
        "catchPlacement": "upper_left",
        "factsError": mat.get("error") or "",
        "byEngine": summary_by_engine,
        "note": "abc16。材料提案＋キャッチ【】優先。金丸・Octas聖域。キャッチ上部左寄り。02未書込。",
        "outRoot": str(out_root),
        "pickDir": str(pick_dir),
    }
    (out_root / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
