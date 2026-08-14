#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""非定型ダンボール行を exclude にする。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
EXCLUDE_NAMES = {"商品入荷箱", "Nekopos封筒（他）"}
SHEET = "00_設定マスタ"


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    cfg = json.loads((SCRIPT_DIR / "config.local.json").read_text(encoding="utf-8"))
    creds = Credentials.from_authorized_user_file(
        str(SCRIPT_DIR / "secrets" / "token_sheets_rw.json"),
        ["https://www.googleapis.com/auth/spreadsheets"],
    )
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    sid = cfg["spreadsheet_id"]
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!B58:H80" % SHEET)
        .execute()
        .get("values")
        or []
    )
    data = []
    for i, row in enumerate(vals):
        if i == 0:
            continue
        name = str(row[0] if row else "").strip()
        if name not in EXCLUDE_NAMES:
            continue
        row1 = 58 + i
        data.append(
            {
                "range": "'%s'!E%d:H%d" % (SHEET, row1, row1),
                "values": [["", "", "", "exclude"]],
            }
        )
        print("exclude", row1, name)
    if not data:
        print("no rows")
        return 0
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=sid,
        body={"valueInputOption": "USER_ENTERED", "data": data},
    ).execute()
    print("done", len(data))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
