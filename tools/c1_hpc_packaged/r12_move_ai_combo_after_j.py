# -*- coding: utf-8 -*-
"""AI情報取得data: 参考情報(画像URL)の右に構成10列を挿入し、右端の重複構成*を削除。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding="utf-8")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
AI = "AI情報取得data"
AI_COMBO = []
for i in range(1, 6):
    AI_COMBO.append("構成%d_商品名" % i)
    AI_COMBO.append("構成%d_数量" % i)
REF = "参考情報(画像URL)"


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
        if s["properties"]["title"] == AI:
            gid = s["properties"]["sheetId"]
            break
    if gid is None:
        print("AI sheet not found")
        return 1

    row = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!1:1" % AI)
        .execute()
        .get("values")
        or [[]]
    )[0]
    print("ncols before", len(row))
    ref_i = None
    for i, h in enumerate(row):
        if str(h).strip() == REF:
            ref_i = i
            break
    if ref_i is None:
        print(REF, "not found")
        return 1
    print("ref col", col_letter(ref_i + 1), ref_i + 1)

    next_h = str(row[ref_i + 1]).strip() if ref_i + 1 < len(row) else ""
    if next_h == AI_COMBO[0]:
        print("already inserted after ref. skip insert")
        inserted = False
    else:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={
                "requests": [
                    {
                        "insertDimension": {
                            "range": {
                                "sheetId": gid,
                                "dimension": "COLUMNS",
                                "startIndex": ref_i + 1,
                                "endIndex": ref_i + 1 + len(AI_COMBO),
                            },
                            "inheritFromBefore": True,
                        }
                    }
                ]
            },
        ).execute()
        start_l = col_letter(ref_i + 2)
        end_l = col_letter(ref_i + 1 + len(AI_COMBO))
        svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range="'%s'!%s1:%s1" % (AI, start_l, end_l),
            valueInputOption="RAW",
            body={"values": [AI_COMBO]},
        ).execute()
        print("inserted", start_l, "to", end_l)
        inserted = True

    row2 = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!1:1" % AI)
        .execute()
        .get("values")
        or [[]]
    )[0]
    first = {}
    dup_idx = []
    for i, h in enumerate(row2):
        key = str(h).strip()
        if not key.startswith("構成"):
            continue
        if key not in first:
            first[key] = i
        else:
            dup_idx.append(i)
            print("dup", col_letter(i + 1), key, "keep", col_letter(first[key] + 1))
    if not dup_idx:
        print("no duplicate 構成* to delete")
    else:
        dup_idx.sort(reverse=True)
        reqs = []
        for i in dup_idx:
            reqs.append(
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": gid,
                            "dimension": "COLUMNS",
                            "startIndex": i,
                            "endIndex": i + 1,
                        }
                    }
                }
            )
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": reqs}).execute()
        print("deleted", len(dup_idx), "dup cols")

    row3 = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!1:1" % AI)
        .execute()
        .get("values")
        or [[]]
    )[0]
    print("ncols after", len(row3))
    for i, h in enumerate(row3):
        if i <= ref_i + len(AI_COMBO) + 4 or str(h).startswith("構成"):
            if i >= ref_i:
                print(col_letter(i + 1), h)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
