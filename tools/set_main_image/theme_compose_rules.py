# -*- coding: utf-8 -*-
"""
テーマ別新規／競合本線で共有するプロンプト条項（袋・スケール・食材・強調色）。

- テーマ別新規: dry_sub_overview_copy_test.py
- 競合本線: sub_image_ai_compose_poc.py
"""
from __future__ import annotations

from typing import Any, Dict, Optional

# 既定はパッケージ無し。ギフト・素材証明・実績など袋が訴求に効くときだけ True。
PACKAGE_DEFAULT_ON_THEME_IDS = frozenset({6, 13, 14, 16})

SCALE_LAYOUT_JA = """
【スケール／人物と物の比率（必須・全商品共通）】
目的: 人物と同画面の器・盛りが「巨大な晩餐皿」に見える違和感をなくす。

■ 体をものさしにする（比較写真が無くても守る）
- 小皿・盛り皿の見た目直径は、だいたい「手のひら〜手のひらよりやや大きい」程度。顔の幅より明らかに大きい皿は禁止。
- 大皿・石皿・ランチプレート級・顔ほどもあるトレーは、人物と同画面では使わない（おつまみ・トッピング・少量珍味では特に）。
- 箸・手・コップが写るときは、それらと盛り皿のサイズが矛盾しないこと（コップより皿が極端に大きい等は禁止）。
- 人物の手前に置く盛りは、上半身に対して控えめ。盛り山が高く・広く人物を圧迫しない。

■ 比較参照があるとき（IMAGE_COMP 等）
- 競合に「人物＋茶碗／小皿／手元」が写っているカットがあれば、器と人の相対サイズの参考にしてよい（レイアウト丸コピーは禁止）。
- ご飯茶碗・酒器・手が写っているカットはスケールの正本候補。食品の色・種類の正本にはしない（FOOD_IDENTITY とは別）。

■ どうしても合わないとき
- (A) 人物ゾーンと料理／商品ゾーンを分ける、または (B) 器をさらに小さくする。巨大皿のまま合成しない。
"""

# 商品非依存。吉野家・石原他SKU・缶飯等でも同一文面。
SIZZLE_FOOD_IDENTITY_JA = """
【T03シズル・食材見た目ロック（全商品共通／必須）】
目的: 「美味しそう」は照り・質感・寄りの食欲で出す。ただし食材の種類・形・色・厚みは偽らない。

■ 食材の正本（FOOD_IDENTITY）
- 食材の見た目の唯一の正は、次のいずれか（優先順）:
  1) このSKUの正本パッケージに写っている中身（窓・抜き絵・中身カット）
  2) このSKU向けに明示した FOOD_PLATING 参照（その商品の盛り／中身写真）
- IMAGE_COMP_*（競合サブ）は、原則レイアウト・文字カード・フレームの参考。
  食材の色・厚み・形状の正本にしてはいけない（別商品の盛りが混ざるため）。
- 別SKU・別味・別シリーズ・別ジャンルの食品が写った参照を、この商品の中身として真似するな。

■ 形・厚み
- 正本に写っている中身の形・厚み・大きさ感に合わせる（薄い／厚い／粒／粉／塊など）。
- 正本が薄い片・フレーク・細い削り状なのに、刺身の厚切り・ステーキ断面・大きな塊に盛るな。
- 正本が塊・ダイス・粒なのに、薄い削り節状に変えるな。
- 「断面アップ」「寄り」はよいが、正本に無い厚み・形状へ誇張しない。

■ 色・質感
- 色相は FOOD_IDENTITY（正本窓／FOOD_PLATING）に合わせる。別商品の色に寄せるな。
- 照り・しっとりは「質感の強調」まで。色そのものを別物に塗り替えるな。
- 正本パッケージを画面に載せない場合でも、中身色だけは正本窓／FOOD_PLATING を厳守する。

■ 湯気
- 常温流通・そのまま食べる前提の商品なら、商品単体から湯気を立てるな。
- 湯気は、画面内に明らかに温かい相手（熱いご飯・スープ等）があるときだけ可。
"""

COMPETITOR_PARTS_NO_PACKAGE_JA = """
④ AI独自・改変は最大約50%／残りは競合パーツ（パッケージは含めない）
- 面積のおよそ半分まで: レイアウト再構成・余白・柔らかい影・ベース色背景・軽い装飾を生成してよい。
- 残りは競合の説明カード・吹き出し・バッジ・フレーム等を流用する。
- 商品パッケージ（袋・箱・缶・パウチ）は競合パーツに含めない。競合写真に袋が大きく写っていても切り貼りするな。
- 袋を画面に出すときの唯一の正は IMAGE_PACKAGE_TRUTH（このテーマで載せる指定のときだけ）。
- 流用は可。構図・人物・文字配置の完全複製は禁止（全く同じはNG）。
- 下記「競合外側パネルOCR」を再配置して描く場合も競合パーツ扱いとする（袋面の文字は除く）。
"""

PERSON_REUSE_JA = """
【人物（競合本線）】
- 競合参照に人物がいれば、構図・雰囲気の流用が正。顔のトレース・同一人物の複製は禁止。
- 詳細な新規人物プロンプトは使わない。
"""

PACKAGE_LOCK_ON = """
【最重要・絶対厳禁／PACKAGE_LOCK（商品パッケージ改変不可）】
- IMAGE_PACKAGE_TRUTH が商品パッケージの唯一の正本である（単体／N=1相当。缶・瓶・ボトル・箱・パウチ可）。
- 正本の色相・彩度・印刷柄・ラベル絵柄・ラベル文字・ロゴ・形状・質感を一切変更しない。
- 正本パッケージ表面へのOCR・再描字・ラベル差し替え・リブランド・新パッケージ創作は禁止。
- N≥2のセットMAINコラージュを色・縦横比の正本にしてはならない。
- 変えてよいのは背景・レイアウト・外側の説明パネル・吹き出し・バッジのみ。
- 競合写真の袋を正本の代わりに貼るな。
- HARD BAN: Do NOT recolor, redraw, OCR-rewrite, or alter PACKAGE_TRUTH packaging. Keep hue, label art, printed text, logo, shape, and aspect ratio identical.

【終了前チェックリスト（必須）】
□ パッケージ色相が正本と同一か
□ ラベル柄・印刷文字が正本のままか（読み替え・創作なし）
□ パッケージ縦横比が PACKAGE_ASPECT_LOCK どおりか（潰しなし）
□ 袋は正本からで、競合写真の袋切り貼りではないか
□ 外側パネルの文字だけが新規／再配置対象になっているか
"""

PACKAGE_LOCK_OFF = """
【パッケージ＝このテーマでは載せない】
- 商品パッケージ（袋・箱）を画面に入れるな。カタログ商品カットにしない。
- IMAGE_PACKAGE_TRUTH は色・質感の参照のみ。正本袋の切り貼り禁止。
- 競合写真に袋が写っていても切り貼りするな（パッケージは競合パーツではない）。
- HARD BAN: Do not composite a product bag/box from PACKAGE_TRUTH or competitor refs.

【終了前チェックリスト（必須）】
□ 画面に商品袋・箱が無いか
□ 正本は色参照だけか（袋の切り貼りなし）
□ 競合写真の袋を貼っていないか
□ 外側パネルの文字だけが新規／再配置対象になっているか
"""


def default_need_package(theme_id: int, explicit: Any = None) -> bool:
    if explicit is None or explicit == "":
        return int(theme_id) in PACKAGE_DEFAULT_ON_THEME_IDS
    if isinstance(explicit, bool):
        return explicit
    s = str(explicit).strip().lower()
    if s in ("1", "true", "yes", "y", "on"):
        return True
    if s in ("0", "false", "no", "n", "off"):
        return False
    return int(theme_id) in PACKAGE_DEFAULT_ON_THEME_IDS


def package_policy_block(need_package: bool, theme_id: int = 0) -> str:
    if int(theme_id) == 13 and need_package:
        return (
            "【パッケージ＝ギフト（T13）】\n"
            "- 正本袋は画面に載せてよいが、人物が袋を直接手渡しする構図は禁止（スケール違和感）。\n"
            "- 望ましい: ラッピング箱／のし袋の手渡し、または箱と正本袋をテーブル横並び展示。袋は人物の手元サイズに合わせる。\n"
            "- 人物の手・腕・肩が袋から生える／袋を貫通する／袋の天面から出るのは禁止。手前の袋の上に人物部位を重ねない。\n"
            "- 色・柄・ラベル・縦横比は正本のまま（PACKAGE_LOCK）。競合写真の袋は貼るな。"
        )
    if need_package:
        return (
            "【パッケージ＝このテーマでは載せる】\n"
            "- IMAGE_PACKAGE_TRUTH の袋を画面に載せてよい（訴求に必要）。\n"
            "- 色・柄・ラベル・縦横比は正本のまま（PACKAGE_LOCK）。再描字・色替え禁止。\n"
            "- 競合写真の袋は切り貼りするな。正本だけが袋の正。\n"
            "- 人物と同画面のとき、袋は実物比率（頭より大きい袋禁止）。難しければ人物と袋を別ゾーンに分ける。\n"
            "- 人物の手・腕・肩が袋から生える／袋を貫通する／袋の天面から出るのは禁止。手前の袋の上に人物部位を重ねない。"
        )
    return (
        "【パッケージ＝このテーマでは載せない（基本）】\n"
        "- 商品パッケージ（袋・箱）を画面に入れるな。カタログ商品カットにしない。\n"
        "- IMAGE_PACKAGE_TRUTH は色・質感の参照のみ。袋の切り貼り禁止。\n"
        "- 競合写真の袋も切り貼りするな。\n"
        "- 主役は料理・食材・人物・小物・訴求文字。"
    )


def package_lock_block(need_package: bool) -> str:
    return PACKAGE_LOCK_ON.strip() if need_package else PACKAGE_LOCK_OFF.strip()


def theme_extra_block(
    theme_id: int,
    *,
    include_steam: bool = True,
    has_food_unit_canon: bool = False,
    has_food_unit_secondary: bool = False,
) -> str:
    """スケール＋食材ユニット正本（全テーマ）＋T03湯気＋T13。"""
    from food_unit_canon import food_unit_prompt_block

    tid = int(theme_id)
    parts = [
        SCALE_LAYOUT_JA,
        "\n"
        + food_unit_prompt_block(
            has_canon_image=has_food_unit_canon,
            has_secondary=has_food_unit_secondary,
        )
        + "\n",
    ]
    # T03: 旧 FOOD_IDENTITY の湯気条項のみ追加（ユニット正本は上で全テーマ適用済み）
    if tid == 3 and include_steam:
        steam = SIZZLE_FOOD_IDENTITY_JA.split("■ 湯気")[-1]
        parts.append("\n【T03 湯気】\n■ 湯気" + steam)
    if tid == 13:
        parts.append(
            "\n【T13ギフト構図】袋の直接手渡し禁止。箱手渡し＋袋は別置き／横並びが正。"
            "手・腕が袋から生える／袋を貫通する構図は禁止（人物と袋は奥行きで分ける）。\n"
        )
    return "".join(parts)


def accent_text_block(text_color: str = "1", *, competitor_ocr: bool = False) -> str:
    from text_color import resolve_text_color

    c = resolve_text_color(text_color)
    name = str(c.get("nameJa") or text_color)
    hexv = str(c.get("hex") or "")
    cid = str(c.get("id") or text_color)
    ocr = ""
    if competitor_ocr:
        ocr = (
            "- 競合OCRのパネル文言・配置を優先。核語の色だけこのアクセント。"
            "OCR本線を別のキャッチ文に差し替えない。\n"
        )
    return f"""
【訴求文字の強調色（MAIN textColor 継承）】
- アクセントの正: textColor={cid} {name}（{hexv}）。親の set_base_map.json / MAIN と同じ番号。
- 主キャッチ（大きな訴求）の核となる語・フレーズの一部だけをこの色系で強調する（全文をアクセント一色にしない）。
- 色相ファミリーは上記を固定。その範囲でグラデ・やや薄め・縁取りなど多少の変化は可（読みやすさ優先）。
- マップと無関係な別色相（例: えんじ指定なのに青）は禁止。
{ocr}- 補足文・小さめ文字は濃茶／ダーク系で可読性優先。強調はキャッチの核語に限定。
"""


def person_reuse_block(theme_id: int) -> str:
    if int(theme_id) in (1, 2, 12):
        return PERSON_REUSE_JA
    return ""


def hit_has_competitor_photo(hit: Optional[Dict[str, Any]]) -> bool:
    if not hit:
        return False
    paths = hit.get("paths") or []
    return any(str(p).strip() for p in paths)
