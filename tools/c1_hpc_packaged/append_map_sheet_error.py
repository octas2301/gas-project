#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ERRORS ブロックへ1行追記（SC結果の結論を sheet に残す）。"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import List

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
TOKEN_RW = SCRIPT_DIR / "secrets" / "token_sheets_rw.json"
SHEET_TITLE = "▼設定(Amazonマッピング)"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
# ERRORS starts at column S (19) — keep in sync with sheet layout
ERR_COL = "S"


def main(argv: List[str] | None = None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--code", default="(human)")
    ap.add_argument("--product-type", default="FOOD")
    ap.add_argument("--count", default="-")
    ap.add_argument("--symptom", required=True)
    ap.add_argument("--cause", default="")
    ap.add_argument("--master-fix", default="")
    ap.add_argument("--map-fix", default="")
    ap.add_argument("--status", default="OPEN")
    ap.add_argument("--summary-file", default="")
    args = ap.parse_args(argv)

    sys.stdout.reconfigure(encoding="utf-8")
    cfg = json.loads((SCRIPT_DIR / "config.local.json").read_text(encoding="utf-8"))
    creds = Credentials.from_authorized_user_file(str(TOKEN_RW), SCOPES)
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    # find ERRORS header row and next empty
    rng = "'%s'!%s1:%s200" % (SHEET_TITLE, ERR_COL, "AA")
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=cfg["spreadsheet_id"], range=rng)
        .execute()
        .get("values")
        or []
    )
    start = None
    for i, r in enumerate(vals):
        if r and str(r[0]).startswith("=== ERRORS"):
            start = i
            break
    if start is None:
        print("ERRORS block not found", file=sys.stderr)
        return 2
    # data after EN header (look for エラーコード)
    data_start = start + 1
    for j in range(start, min(start + 8, len(vals))):
        if vals[j] and "エラーコード" in str(vals[j][0]):
            data_start = j + 1
            break
    row_num = data_start + 1
    for j in range(data_start, len(vals)):
        if not vals[j] or not str(vals[j][0]).strip():
            row_num = j + 1
            break
        row_num = j + 2

    row = [
        args.code,
        args.product_type,
        args.count,
        args.symptom,
        args.cause,
        args.master_fix,
        args.map_fix,
        args.status,
        args.summary_file,
    ]
    target = "'%s'!%s%d" % (SHEET_TITLE, ERR_COL, row_num)
    svc.spreadsheets().values().update(
        spreadsheetId=cfg["spreadsheet_id"],
        range=target,
        valueInputOption="RAW",
        body={"values": [row]},
    ).execute()
    print("appended", target, row)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
