# -*- coding: utf-8 -*-
"""
サブ画像レビュー／再生成向け・人間コメントのテンプレ集。

- review_loop UI・B-④目視チェック・Agent指示から共通利用する。
- 偽物感（CAMERA_LOOK）対策の定型句をここに足していく。
"""
from __future__ import annotations

from typing import Any, Dict, List

VERSION = "2026-08-10.1"

# id, label（UI表示）, text（textareaへ挿入する本文）
TEMPLATES: List[Dict[str, str]] = [
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
