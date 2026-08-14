#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MAP 行を attrKey 指定で部分更新（sheet を最新の実行ルールに揃える）。
正本フロー: 結論を sheet に書いてから sync_map_sheet_to_column_json。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
TOKEN_RW = SCRIPT_DIR / "secrets" / "token_sheets_rw.json"
SHEET_TITLE = "▼設定(Amazonマッピング)"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 2026-08-02 缶飯 GROCERY 対策（承認済み実装と一致）
PATCHES: Dict[str, Dict[str, str]] = {
    "var_theme": {
        "valueSource": "固定",
        "inherit": "継承なし（固定／別経路）",
        "transform": "固定値を使う",
        "defaultValue": "サイズ",
        "masterColPrimary": "",
        "doNotUse": "",
        "notes": "GROCERY純正プルダウンのみ（サイズ／フレーバー／色／味/サイズ）。SET_NAME不可",
    },
    "color": {
        "required": "OPT",
        "defaultValue": "",
        "valueSource": "マスタのみ",
        "notes": "GROCERYテーマ=サイズ時は未出力（その他固定禁止）",
    },
    "temperature_rating": {
        "required": "OPT",
        "valueSource": "固定",
        "inherit": "継承なし（固定／別経路）",
        "defaultValue": "常温：室温",
        "masterColPrimary": "温度の定格",
        "masterColFallback": "",
        "doNotUse": "保存方法(食品)|(保存方法(食品))|▼マスタ(保存方法(食品))",
        "notes": "GROCERY黒セル時は未出力。保存方法長文を定格に載せない",
    },
    "item_form": {
        "required": "OPT",
        "defaultValue": "",
        "valueSource": "マスタのみ",
        "inherit": "継承なし（固定／別経路）",
        "notes": "GROCERY黒セル時は未出力（ホール固定禁止）",
    },
    "size": {
        "scHeaderJa": "サイズ",
        "scHeaderAlias": "サイズ|size[marketplace_id=A1VC38T7YXB528][language_tag=ja_JP]#1.value",
        "doNotUse": "パッケージサイズ名",
        "notes": "テーマ=サイズ時はATのsize属性。package_size_nameではない",
    },
}


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    cfg = json.loads((SCRIPT_DIR / "config.local.json").read_text(encoding="utf-8"))
    creds = Credentials.from_authorized_user_file(str(TOKEN_RW), SCOPES)
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=cfg["spreadsheet_id"], range="'%s'!A1:Q400" % SHEET_TITLE)
        .execute()
        .get("values")
        or []
    )
    header_i = None
    for i, r in enumerate(vals):
        if r and str(r[0]) == "productCategory" and len(r) > 2 and str(r[2]) == "attrKey":
            header_i = i
            break
    if header_i is None:
        print("header not found", file=sys.stderr)
        return 2
    headers = [str(h) for h in vals[header_i]]
    idx = {h: i for i, h in enumerate(headers)}
    updates = []
    for ri, r in enumerate(vals):
        if ri <= header_i + 2:
            continue
        if len(r) < 3:
            continue
        ak = str(r[idx["attrKey"]]) if idx["attrKey"] < len(r) else ""
        if ak not in PATCHES:
            continue
        patch = PATCHES[ak]
        new_r = list(r) + [""] * (len(headers) - len(r))
        new_r = new_r[: len(headers)]
        for k, v in patch.items():
            if k in idx:
                new_r[idx[k]] = v
        sheet_row = ri + 1  # 1-based
        # write A..Q for that row
        updates.append(
            {
                "range": "'%s'!A%d:Q%d" % (SHEET_TITLE, sheet_row, sheet_row),
                "values": [new_r],
            }
        )
        print("patch", ak, "row", sheet_row)

    if not updates:
        print("no rows patched", file=sys.stderr)
        return 1
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=cfg["spreadsheet_id"],
        body={"valueInputOption": "RAW", "data": updates},
    ).execute()
    print("ok", len(updates))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
