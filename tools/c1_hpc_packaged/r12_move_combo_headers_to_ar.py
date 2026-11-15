# -*- coding: utf-8 -*-
"""組み合わせ列を AR〜BE（14列）へ移す。AR〜BJ は19列あるので足りる。

前回: 空列＋画像URL列＋作業エリア右挿入。
今回: 作業エリア右の5列を削除し、AR8:BE8 に14項目を書く。
BF〜BI は旧「出品者SKUのサブ画像URL2〜5」に戻す。
AI の構成列はそのまま。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding="utf-8")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MASTER = "▼商品マスタ(人間作業用)"
HEADER_ROW = 8
AR_INDEX0 = 43  # AR
N_COMBO = 14
COMBO_NAMES = [
    "セット構成1_親SKU",
    "セット構成1_数量",
    "セット構成2_親SKU",
    "セット構成2_数量",
    "セット構成3_親SKU",
    "セット構成3_数量",
    "セット構成4_親SKU",
    "セット構成4_数量",
    "セット構成5_親SKU",
    "セット構成5_数量",
    "セット構成_状態",
    "セット構成_幅_cm",
    "セット構成_奥_cm",
    "セット構成_高_cm",
]
RESTORE_BF_BI = [
    "出品者SKUのサブ画像URL2",
    "出品者SKUのサブ画像URL3",
    "出品者SKUのサブ画像URL4",
    "出品者SKUのサブ画像URL5",
]
AFTER_AI = [
    "セット構成5_親SKU",
    "セット構成5_数量",
    "セット構成_幅_cm",
    "セット構成_奥_cm",
    "セット構成_高_cm",
]


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main() -> int:
    base = Path(__file__).resolve().parent
    sid = json.loads((base / "config.local.json").read_text(encoding="utf-8"))["spreadsheet_id"]
    creds = Credentials.from_authorized_user_file(str(base / "secrets" / "token_sheets_rw.json"), SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(sheetId,title))").execute()
    gid = None
    for s in meta.get("sheets") or []:
        if s["properties"]["title"] == MASTER:
            gid = s["properties"]["sheetId"]
            break
    if gid is None:
        print("master sheet not found")
        return 1

    row = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!%d:%d" % (MASTER, HEADER_ROW, HEADER_ROW))
        .execute()
        .get("values")
        or [[]]
    )[0]
    print("AR-BJ now:")
    for i in range(AR_INDEX0, AR_INDEX0 + 19):
        h = row[i] if i < len(row) else ""
        print(" ", col_letter(i + 1), repr(h)[:60])

    ai_i = None
    for i, h in enumerate(row):
        if str(h).strip() == "⇒AI作業エリア":
            ai_i = i
            break
    if ai_i is None:
        print("⇒AI作業エリア not found")
        return 1

    extra_start = None
    if ai_i + 1 < len(row) and str(row[ai_i + 1]).strip() == AFTER_AI[0]:
        extra_start = ai_i + 1
        ok = True
        for k, name in enumerate(AFTER_AI):
            if ai_i + 1 + k >= len(row) or str(row[ai_i + 1 + k]).strip() != name:
                ok = False
                break
        if not ok:
            print("作業エリア右の5列が想定と違う。中止")
            for k in range(6):
                j = ai_i + 1 + k
                print(" ", col_letter(j + 1), repr(row[j] if j < len(row) else ""))
            return 1

    print("delete 5 cols after AI area:", extra_start is not None)
    print("write AR:BE", COMBO_NAMES)
    print("restore BF:BI", RESTORE_BF_BI)

    if extra_start is not None:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={
                "requests": [
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": gid,
                                "dimension": "COLUMNS",
                                "startIndex": extra_start,
                                "endIndex": extra_start + 5,
                            }
                        }
                    }
                ]
            },
        ).execute()
        print("deleted extra 5")

    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range="'%s'!AR%d:BE%d" % (MASTER, HEADER_ROW, HEADER_ROW),
        valueInputOption="RAW",
        body={"values": [COMBO_NAMES]},
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range="'%s'!BF%d:BI%d" % (MASTER, HEADER_ROW, HEADER_ROW),
        valueInputOption="RAW",
        body={"values": [RESTORE_BF_BI]},
    ).execute()
    row2 = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!AR%d:BJ%d" % (MASTER, HEADER_ROW, HEADER_ROW))
        .execute()
        .get("values")
        or [[]]
    )[0]
    print("AR-BJ after:")
    letters = ["AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ"]
    for i, let in enumerate(letters):
        h = row2[i] if i < len(row2) else ""
        print(" ", let, h)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
