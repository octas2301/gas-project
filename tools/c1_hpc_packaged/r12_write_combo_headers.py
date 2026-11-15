# -*- coding: utf-8 -*-
"""マスタ8行目に組み合わせ列名を書く。字母固定ではなく、今の項目名で列を探す。

AR〜AY は 対象ASIN／商品名／セール等のため触らない。
空3列（商品の公開日の右）＋出品者SKU画像URL6列を改名し、足りない5列は
「⇒AI作業エリア」の直後に挿入する（式コピー範囲の外）。
AI情報取得data の1行目末尾に構成1〜5を追加する。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

APPLY = "--apply" in sys.argv
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MASTER = "▼商品マスタ(人間作業用)"
AI = "AI情報取得data"
HEADER_ROW = 8

COMBO_ON_REUSE = [
    "セット構成1_親SKU",
    "セット構成1_数量",
    "セット構成2_親SKU",
    "セット構成2_数量",
    "セット構成3_親SKU",
    "セット構成3_数量",
    "セット構成4_親SKU",
    "セット構成4_数量",
    "セット構成_状態",
]
COMBO_AFTER_AI_AREA = [
    "セット構成5_親SKU",
    "セット構成5_数量",
    "セット構成_幅_cm",
    "セット構成_奥_cm",
    "セット構成_高_cm",
]
AI_COMBO = []
for i in range(1, 6):
    AI_COMBO.append("構成%d_商品名" % i)
    AI_COMBO.append("構成%d_数量" % i)

IMAGE_HEADERS = [
    "出品者SKUのメイン画像URL",
    "出品者SKUのサブ画像URL1",
    "出品者SKUのサブ画像URL2",
    "出品者SKUのサブ画像URL3",
    "出品者SKUのサブ画像URL4",
    "出品者SKUのサブ画像URL5",
]


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    base = Path(__file__).resolve().parent
    sid = json.loads((base / "config.local.json").read_text(encoding="utf-8"))["spreadsheet_id"]
    creds = Credentials.from_authorized_user_file(str(base / "secrets" / "token_sheets_rw.json"), SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(sheetId,title))").execute()
    sheet_ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta.get("sheets") or []}

    row = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!%d:%d" % (MASTER, HEADER_ROW, HEADER_ROW))
        .execute()
        .get("values")
        or [[]]
    )[0]
    name_to_i = {}
    for i, h in enumerate(row):
        key = str(h).strip()
        if key and key not in name_to_i:
            name_to_i[key] = i

    if "セット構成1_親SKU" in name_to_i:
        print("マスタは既にセット構成1_親SKUあり。スキップ")
    else:
        reuse_idx = []
        pub = name_to_i.get("商品の公開日")
        if pub is None:
            print("商品の公開日 が見つかりません")
            return 1
        for j in range(pub + 1, min(pub + 8, len(row) + 8)):
            h = str(row[j]).strip() if j < len(row) else ""
            if h == "":
                reuse_idx.append(j)
            if len(reuse_idx) >= 3:
                break
        for hn in IMAGE_HEADERS:
            if hn not in name_to_i:
                print("画像列なし", hn)
                return 1
            reuse_idx.append(name_to_i[hn])
        if len(reuse_idx) != 9:
            print("再利用列が9でない", reuse_idx)
            return 1
        print("reuse by name:")
        for i, idx in enumerate(reuse_idx):
            old = row[idx] if idx < len(row) else ""
            print(" ", col_letter(idx + 1), "was", repr(old), "->", COMBO_ON_REUSE[i])

        ai_i = name_to_i.get("⇒AI作業エリア")
        if ai_i is None:
            print("⇒AI作業エリア なし")
            return 1
        print("insert 5 cols after", col_letter(ai_i + 1), COMBO_AFTER_AI_AREA)

        if APPLY:
            data = []
            for i, idx in enumerate(reuse_idx):
                a1 = "'%s'!%s%d" % (MASTER, col_letter(idx + 1), HEADER_ROW)
                data.append({"range": a1, "values": [[COMBO_ON_REUSE[i]]]})
            svc.spreadsheets().values().batchUpdate(
                spreadsheetId=sid,
                body={"valueInputOption": "RAW", "data": data},
            ).execute()
            gid = sheet_ids[MASTER]
            svc.spreadsheets().batchUpdate(
                spreadsheetId=sid,
                body={
                    "requests": [
                        {
                            "insertDimension": {
                                "range": {
                                    "sheetId": gid,
                                    "dimension": "COLUMNS",
                                    "startIndex": ai_i + 1,
                                    "endIndex": ai_i + 1 + len(COMBO_AFTER_AI_AREA),
                                },
                                "inheritFromBefore": False,
                            }
                        }
                    ]
                },
            ).execute()
            start_letter = col_letter(ai_i + 2)
            end_letter = col_letter(ai_i + 1 + len(COMBO_AFTER_AI_AREA))
            rng = "'%s'!%s%d:%s%d" % (MASTER, start_letter, HEADER_ROW, end_letter, HEADER_ROW)
            svc.spreadsheets().values().update(
                spreadsheetId=sid,
                range=rng,
                valueInputOption="RAW",
                body={"values": [COMBO_AFTER_AI_AREA]},
            ).execute()
            print("master headers written")
        else:
            print("dry_run (pass --apply to write)")

    ai_row = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!1:1" % AI)
        .execute()
        .get("values")
        or [[]]
    )[0]
    if "構成1_商品名" in [str(x).strip() for x in ai_row]:
        print("AI 構成1_商品名 既存。スキップ")
        return 0
    start = len(ai_row) + 1
    end = start + len(AI_COMBO) - 1
    print("AI append", col_letter(start), "to", col_letter(end), AI_COMBO)
    if APPLY:
        gid = sheet_ids[AI]
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={
                "requests": [
                    {
                        "insertDimension": {
                            "range": {
                                "sheetId": gid,
                                "dimension": "COLUMNS",
                            "startIndex": len(ai_row),
                            "endIndex": len(ai_row) + len(AI_COMBO),
                            },
                            "inheritFromBefore": True,
                        }
                    }
                ]
            },
        ).execute()
        rng = "'%s'!%s1:%s1" % (AI, col_letter(start), col_letter(end))
        svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range=rng,
            valueInputOption="RAW",
            body={"values": [AI_COMBO]},
        ).execute()
        print("AI headers written")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
