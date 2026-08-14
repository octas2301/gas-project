# -*- coding: utf-8 -*-
"""
Amazon サブ画像 PoC 用プロンプト（テスト出力専用・サブのみ）。

10パターン:
  PHOTO_*     … 実写写真を背景に使う（AI合成背景だけにしない）
  COMP_MASH_* … 競合画像のパーツ掛け合わせ
  SCENE_*     … 利用シーン／食事／サイズ／保管（白抜きメイン相当は禁止）

日本語OCR→再描画禁止。シンプル商品のみ（白背景パックショットMAIN相当）は出さない。
"""
from __future__ import annotations

from typing import Dict, List

PHOTO_REALISM_RULES = """
PHOTOREALISM (mandatory — reduce "AI look"):
- Look like a real camera photo, NOT CGI / illustration / AI showcase.
- Natural lighting; soft shadows; mild sensor noise OK.
- Avoid: plastic materials, oversaturated HDR, fake creamy bokeh, glowing edges,
  watermarks, overly smooth gradients, regenerated packaging text.
- Neutral-commercial Amazon JP secondary-image look.
"""

NO_MAIN_RULE = """
SUB-IMAGE ONLY (mandatory):
- This is an Amazon SECONDARY / PT-style image, NOT the main packshot.
- FORBIDDEN: plain white/gray studio background with only the product centered (MAIN style).
- REQUIRED: visible scene, photo background, props context, or multi-part mashup.
"""

# 10種サブ画像パターン（MAIN相当なし）
SUB_PATTERNS: List[Dict[str, str]] = [
    {
        "id": "P01_photo_table",
        "ja_title": "P01 写真背景・テーブル",
        "ja_goal": "実写の木テーブル写真の上に自社商品を置く（AI背景ではなく写真）。",
        "role_en": "own product on real wood-table photo",
        "kind": "photo_pillow",
        "bg_key": "bg_table_wood",
    },
    {
        "id": "P02_photo_kitchen",
        "ja_title": "P02 写真背景・キッチン",
        "ja_goal": "実写キッチン写真の上に自社商品を置く。",
        "role_en": "own product on real kitchen photo",
        "kind": "photo_pillow",
        "bg_key": "bg_kitchen",
    },
    {
        "id": "P03_photo_dining",
        "ja_title": "P03 写真背景・ダイニング",
        "ja_goal": "実写ダイニング写真の上に自社商品を置く。",
        "role_en": "own product on real dining photo",
        "kind": "photo_pillow",
        "bg_key": "bg_dining",
    },
    {
        "id": "P04_photo_pantry",
        "ja_title": "P04 写真背景・パントリー",
        "ja_goal": "実写パントリー／棚写真の上に自社商品を置く。",
        "role_en": "own product on real pantry photo",
        "kind": "photo_pillow",
        "bg_key": "bg_pantry",
    },
    {
        "id": "P05_comp_mash_pillow",
        "ja_title": "P05 競合パーツ掛け合わせ（合成）",
        "ja_goal": "複数競合写真のパーツを切り貼りし、中央に自社商品を置いて新画像にする。",
        "role_en": "pillow collage of competitor crops + own product",
        "kind": "mash_pillow",
        "bg_key": "bg_counter",
    },
    {
        "id": "P06_comp_mash_ai",
        "ja_title": "P06 競合パーツ掛け合わせ（AI再構成）",
        "ja_goal": "競合A/Bの構図・小物ヒントを掛け合わせ、自社商品のサブ画像に再構成。",
        "role_en": "AI mashup from competitor parts/hints + own",
        "kind": "mash_ai",
        "bg_key": "bg_counter",
    },
    {
        "id": "P07_usage_photo",
        "ja_title": "P07 利用シーン（写真背景）",
        "ja_goal": "実写キッチン写真を背景に、自社商品の利用シーン風サブ画像。",
        "role_en": "usage scene guided by real kitchen photo",
        "kind": "scene_ai",
        "bg_key": "bg_kitchen",
    },
    {
        "id": "P08_meal_photo",
        "ja_title": "P08 食事ペアリング（写真背景）",
        "ja_goal": "実写ダイニング写真を背景に、食事との組み合わせサブ画像。",
        "role_en": "meal pairing guided by real dining photo",
        "kind": "scene_ai",
        "bg_key": "bg_dining",
    },
    {
        "id": "P09_size_photo",
        "ja_title": "P09 サイズ感（写真背景）",
        "ja_goal": "実写テーブル写真上で、手や食器とのサイズ感が分かるサブ画像。",
        "role_en": "size context on real table photo",
        "kind": "scene_ai",
        "bg_key": "bg_table_wood",
    },
    {
        "id": "P10_storage_photo",
        "ja_title": "P10 保管・棚（写真背景）",
        "ja_goal": "実写パントリー写真を背景に、保管イメージのサブ画像。",
        "role_en": "storage/shelf mood on real pantry photo",
        "kind": "scene_ai",
        "bg_key": "bg_pantry",
    },
]

# 後方互換（旧軸は使わないが import 切れ防止）
POC_SLOTS: List[Dict[str, str]] = []
OWN_SLOTS: List[Dict[str, str]] = []


def prompt_comp_mash_ai(*, product_name_hint: str = "") -> str:
    name = (product_name_hint or "our product").strip()
    return f"""You create ONE Amazon SECONDARY listing image by MASHING competitor visual parts.

MODE = COMP_MASH_AI
Product name hint: {name}

INPUT ROLES:
- IMAGE_COMP_A / IMAGE_COMP_B = competitor listing photos. Use them as PART SOURCES only
  (props, table vibe, lighting, crop ideas). Do NOT keep competitor brand/SKU as the hero.
- IMAGE_OWN = OUR product. This is the ONLY allowed hero product identity.
- IMAGE_BG = a REAL photograph used as the environment base. Prefer keeping this photo's
  surfaces/lighting; do not invent a pure white studio.

TASK:
- Combine layout/prop ideas from COMP_A and COMP_B into a NEW image.
- Place OUR product from IMAGE_OWN as the clear hero.
- Keep a lived-in secondary-image feel (not MAIN packshot).

{NO_MAIN_RULE}
{PHOTO_REALISM_RULES}

FORBIDDEN: OCR/redraw of Japanese packaging text; competitor SKU as hero; plain white MAIN shot.
Return one square photoreal secondary image.
"""


def prompt_scene_ai(*, slot_id: str, product_name_hint: str = "") -> str:
    slot = next((s for s in SUB_PATTERNS if s["id"] == slot_id), SUB_PATTERNS[6])
    name = (product_name_hint or "our product").strip()
    return f"""You create ONE Amazon SECONDARY (PT) lifestyle-style image.

MODE = SCENE_AI
GOAL = {slot['id']} ({slot['role_en']})
Human intent (JA): {slot['ja_goal']}
Product: {name}

INPUT ROLES:
- IMAGE_BG = REAL photograph. Use this photo as the scene/background base (do not replace with
  empty white AI studio). Keep recognizable photo surfaces from IMAGE_BG.
- IMAGE_OWN = OUR product (hero). Faithful shape/label; no OCR redraw of text.
- IMAGE_COMP_A = optional composition hint only (ignore competitor brand identity).

SLOT HINTS:
- P07_usage_photo: product in use / ready-to-use on kitchen surfaces from IMAGE_BG.
- P08_meal_photo: product near food/plates mood using IMAGE_BG dining ambience.
- P09_size_photo: product scale clear vs tableware/hand-scale cues; keep IMAGE_BG table.
- P10_storage_photo: product on shelf/pantry vibe from IMAGE_BG.

{NO_MAIN_RULE}
{PHOTO_REALISM_RULES}

Return one square photoreal secondary image (not MAIN packshot).
"""


def prompt_comp_micro(*, slot_id: str, product_name_hint: str = "") -> str:
    """旧API互換スタブ（現行PoCでは未使用）。"""
    _ = (slot_id, product_name_hint)
    return "unused"


def prompt_own_remake(*, slot_id: str, product_name_hint: str = "", use_comp_hint: bool = True) -> str:
    """旧API互換スタブ（現行PoCでは未使用）。"""
    _ = (slot_id, product_name_hint, use_comp_hint)
    return "unused"


def logic_steps_for_humans() -> List[Dict[str, str]]:
    return [
        {
            "step": "1",
            "title": "サブ画像のみ",
            "detail": "白背景の商品だけMAIN相当は作らない。シーン／写真背景／掛け合わせのみ。",
        },
        {
            "step": "2",
            "title": "写真背景",
            "detail": "P01–P04は実写JPGを背景にPillow合成。P07–P10は同写真をIMAGE_BGとしてAI誘導。",
        },
        {
            "step": "3",
            "title": "競合掛け合わせ",
            "detail": "P05は複数競合クロップのPillowコラージュ＋自社。P06はAIでパーツ再構成。",
        },
        {
            "step": "4",
            "title": "10パターン出力",
            "detail": "各パターン1枚＋注釈ボード。テスト出力のみ（マスタ未書込）。",
        },
    ]
