# -*- coding: utf-8 -*-
"""
サブ画像レビュー／再生成向け・人間コメントのテンプレ集。

- review_loop UI・B-④目視チェック・Agent指示から共通利用する。
- 偽物感（CAMERA_LOOK）対策の定型句をここに足していく。
"""
from __future__ import annotations

from typing import Any, Dict, List

VERSION = "2026-10-19.1"

# id, label（UI表示）, text（textareaへ挿入する本文）
TEMPLATES: List[Dict[str, str]] = [
    {
        "id": "piece_vs_specimen",
        "label": "個品が見本と違う",
        "text": (
            "個品の色・厚み・1片の大きさが、確認済みの見本と違う。"
            "種類ごとに、その見本の色・厚み・1片のまま描き直す。"
            "見本に無い炙り、別の切り方、別の色は足さない。"
            "文言、配置、包材は変えない。"
        ),
    },
    {
        "id": "scale_plate_person",
        "label": "器・盛りが人物に対して大きい",
        "text": (
            "人物と同画面の器・盛りが大きすぎる。皿は手のひら〜やや大きめにし、"
            "顔幅を超える大皿・石皿・ランチプレートは禁止。顔・手・コップをものさしに合わせる。"
            "合わなければ人物と料理を別ゾーンに分ける。"
        ),
    },
    {
        "id": "food_unit_mismatch",
        "label": "食材の見た目がテーマ間でバラバラ",
        "text": (
            "テーマごとに食材の色・厚み・1片サイズが別商品に見える。"
            "IMAGE_FOOD_UNIT_CANON（または正本パッケージの中身窓）に合わせ、"
            "全テーマで同じSKUの同じ1片に統一する。競合COMPの盛りを食材正本にしない。"
        ),
    },
    {
        "id": "accent_text_weak",
        "label": "強調色が弱い／無い",
        "text": (
            "主キャッチの核となる語に強調色が無い、または弱い。"
            "親の共有 textColor（1..5）をアクセントの正とし、キャッチの一部語だけ色付けする。"
            "色相ファミリー内でグラデ・やや薄め・縁取りは可。マップと無関係な別色相は禁止。"
            "補足文はダーク系のまま可読性優先。"
        ),
    },
    {
        "id": "cutout_edges",
        "label": "切り抜き輪郭が硬い",
        "text": (
            "商品・背景・小物の輪郭が切り抜きパスのように硬く偽物。 "
            "Canon R5 50mm f/1.8 の光学ソフトネスとフォーカスフォールオフを入れ、"
            "非主役の縁を少しぼかす。全面ピンシャープのコラージュ禁止。"
        ),
    },
    {
        "id": "bg_too_sharp",
        "label": "背景・小物がシャープすぎ",
        "text": (
            "背景・布・テーブル縁・遠い小物が主役と同じ鋭さで偽物。 "
            "浅い被写界深度と creamy/smooth bokeh を強く。主役パッケージ主面だけシャープ。"
        ),
    },
    {
        "id": "cgi_lighting",
        "label": "均一CGI照明",
        "text": (
            "照明が均一で影が無くCGIっぽい。PACKAGE_TRUTHの光源方向・ハイライト・接地影を残し、"
            "フラットスタジオライト禁止。"
        ),
    },
    {
        "id": "plastic_surface",
        "label": "プラスチック質感",
        "text": (
            "表面がのっぺりプラスチック。わずかなセンサー粒感（ISO粒）を残し、"
            "過剰スムージング禁止。缶/瓶の反射は消さない。"
        ),
    },
    {
        "id": "steam_vector",
        "label": "湯気・粉が線画",
        "text": (
            "湯気・ふりかけ・粒子がベクター線のように整いすぎ。写真的なボケ粒子・半透明の塊で描く。"
        ),
    },
    {
        "id": "package_lock",
        "label": "パッケージ色・ラベルずれ",
        "text": (
            "PACKAGE_LOCK違反。正本と色・柄・ラベル文字・ロゴ・縦横比を一致させる。"
            "パッケージ表面の再描字・色替え禁止。"
        ),
    },
    {
        "id": "too_much_text",
        "label": "文字が多すぎ",
        "text": (
            "文字量上限超過。見出し≤12文字、本文≤3行、コールアウト≤2。"
            "栄養は3要点まで。余分な文言を削る。"
        ),
    },
]


def templates_for_ui() -> List[Dict[str, str]]:
    return [dict(t) for t in TEMPLATES]


def templates_block_ja() -> str:
    lines = [
        "【再生成コメント・テンプレ（コピーして要望欄へ）】",
        f"（templatesVersion={VERSION}）",
        "",
    ]
    for i, t in enumerate(TEMPLATES, 1):
        lines.append(f"{i}. {t['label']}")
        lines.append(f"   {t['text']}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def templates_meta() -> Dict[str, Any]:
    return {"version": VERSION, "count": len(TEMPLATES), "ids": [t["id"] for t in TEMPLATES]}
