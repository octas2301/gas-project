# -*- coding: utf-8 -*-
"""Generate YahooCaptionRich.js TMPL constant from verified paste HTML.

注意（2026-09-15）:
  本スクリプトが書き出すのは骨格HTMLスニペットのみ。
  P5変数埋込・API反映リトライ・toast 等は手メンテの YahooCaptionRich.js 側が正。
  骨格HTMLだけ差し替えるときは、生成結果の TMPL 定数だけを現行ファイルへ手動マージすること。
"""
from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parents[2]
PASTE = ROOT / "tools" / "yahoo_page_mock" / "paste"
OUT_TMPL_ONLY = ROOT / "tools" / "yahoo_page_mock" / "out" / "YAHOO_CAPTION_RICH_TMPL_snippet.js"

FILES = {
    "caption": "katsuo_1bag_caption_PASTE.html",
    "additional1": "katsuo_1bag_fs1_PASTE.html",
    "additional2": "katsuo_1bag_fs2_PASTE.html",
    "sp_additional": "katsuo_1bag_caption_PASTE_sp.html",
}


def js_string(s: str) -> str:
    return json.dumps(s, ensure_ascii=False)


parts = {}
for key, name in FILES.items():
    parts[key] = (PASTE / name).read_text(encoding="utf-8").replace("\r\n", "\n").strip() + "\n"

snippet = f'''/** 検証済みかつおテンプレ＝おだし系共通骨格（シリーズ全袋同一HTML＋P5変数埋込） */
var YAHOO_CAPTION_RICH_TMPL_KATSUO_ODASHI_ = {{
  caption: {js_string(parts["caption"])},
  additional1: {js_string(parts["additional1"])},
  additional2: {js_string(parts["additional2"])},
  sp_additional: {js_string(parts["sp_additional"])}
}};
'''

OUT_TMPL_ONLY.parent.mkdir(parents=True, exist_ok=True)
OUT_TMPL_ONLY.write_text(snippet, encoding="utf-8", newline="\n")
print("wrote TMPL snippet only:", OUT_TMPL_ONLY)
print("Do NOT overwrite YahooCaptionRich.js wholesale (P5 logic lives there).", file=sys.stderr)
