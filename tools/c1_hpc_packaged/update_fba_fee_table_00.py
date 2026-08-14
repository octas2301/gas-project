#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""00_設定マスタ の FBA手数料ブロックだけを現行公式表に更新する。

安全方針:
- A列が「FBA手数料」の連続ブロックだけを対象
- その直後（販売手数料等）の内容は書き換えない
- 行不足時は FBAブロック末尾（次カテゴリの直前）に行を挿入してから書く
- 行余剰時は余剰FBA行の A/B/D/F のみクリア（他カテゴリ行は削除しない）
- 書く列は A/B/D/F のみ（C/E は触らない）
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
SHEET = "00_設定マスタ"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 公式 sell.amazon.co.jp/pricing（1,000円超）2026-08-13 承認
# docs/org/LV4_FBA_FEE_TABLE_00_UPDATE_APPROVAL.md
TARGET: List[Tuple[str, int, str]] = [
    ("小型", 288, "25x18x2.0cm/250g"),
    ("標準1", 318, "35x30x3.3cm/1kg"),
    ("標準2a", 410, "20cm/2kg"),
    ("標準2b", 415, "30cm/2kg"),
    ("標準2c", 420, "40cm/2kg"),
    ("標準2d", 425, "50cm/2kg"),
    ("標準2e", 430, "60cm/2kg"),
    ("標準3", 472, "80cm/5kg"),
    ("標準4", 532, "100cm/9kg"),
    ("大型1", 589, "60cm/2kg"),
    ("大型2", 624, "80cm/5kg"),
    ("大型3", 675, "100cm/10kg"),
    ("大型4", 781, "120cm/15kg"),
    ("大型5", 1020, "140cm/20kg"),
    ("大型6", 1100, "160cm/25kg"),
    ("大型7", 1532, "180cm/30kg"),
    ("大型8", 1756, "200cm/40kg"),
    ("特大1", 2755, "200cm/50kg"),
    ("特大2", 3573, "220cm/50kg"),
    ("特大3", 4496, "240cm/50kg"),
    ("特大4", 5625, "260cm/50kg"),
]


def cell(row: List[Any], idx: int) -> str:
    if idx < len(row) and row[idx] is not None:
        return str(row[idx]).strip()
    return ""


def find_fba_block(vals: List[List[Any]]) -> Tuple[int, int]:
    """1-based inclusive [start, end] of contiguous FBA手数料 rows."""
    starts = [i for i, r in enumerate(vals) if cell(r, 0) == "FBA手数料"]
    if not starts:
        raise RuntimeError("A列に FBA手数料 行がありません")
    start0 = starts[0]
    end0 = start0
    for i in range(start0 + 1, len(vals)):
        if cell(vals[i], 0) == "FBA手数料":
            end0 = i
        else:
            break
    # contiguous from first — if gaps, still take first contiguous run only
    return start0 + 1, end0 + 1


def sheet_id_for_title(svc, sid: str, title: str) -> int:
    meta = (
        svc.spreadsheets()
        .get(spreadsheetId=sid, fields="sheets(properties(sheetId,title))")
        .execute()
    )
    for sh in meta.get("sheets", []):
        props = sh.get("properties") or {}
        if props.get("title") == title:
            return int(props["sheetId"])
    raise RuntimeError("sheet not found: " + title)


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    dry = "--dry-run" in sys.argv
    cfg = json.loads((SCRIPT_DIR / "config.local.json").read_text(encoding="utf-8"))
    sid = cfg["spreadsheet_id"]
    creds = Credentials.from_authorized_user_file(
        str(SCRIPT_DIR / "secrets" / "token_sheets_rw.json"), SCOPES
    )
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!A1:F120" % SHEET)
        .execute()
        .get("values")
        or []
    )

    start1, end1 = find_fba_block(vals)
    cur_n = end1 - start1 + 1
    need_n = len(TARGET)
    print("FBA block rows %d..%d (n=%d) need=%d dry=%s" % (start1, end1, cur_n, need_n, dry))

    # show after-block category
    after_i = end1  # 1-based next
    after_a = ""
    if after_i < len(vals):
        after_a = cell(vals[after_i], 0)  # vals is 0-based; after_i is 1-based = index after_i
    # end1 is last FBA 1-based → next row index in vals is end1 (0-based end1)
    if end1 < len(vals):
        after_a = cell(vals[end1], 0)
        after_b = cell(vals[end1], 1)
        print("next row after FBA: r%d A=%s B=%s" % (end1 + 1, after_a, after_b))
    else:
        print("next row after FBA: (EOF)")

    print("--- current FBA ---")
    for r in range(start1, end1 + 1):
        row = vals[r - 1] if r - 1 < len(vals) else []
        print(
            "%d\t%s\t%s\t%s\t%s"
            % (r, cell(row, 0), cell(row, 1), cell(row, 3), cell(row, 5))
        )

    if dry:
        print("dry-run: no write")
        return 0

    gid = sheet_id_for_title(svc, sid, SHEET)
    requests: List[Dict[str, Any]] = []

    if need_n > cur_n:
        insert_at = end1  # insert before row end1+1 (0-based startIndex = end1)
        n_ins = need_n - cur_n
        print("insert %d rows at 1-based position %d (before next category)" % (n_ins, end1 + 1))
        requests.append(
            {
                "insertDimension": {
                    "range": {
                        "sheetId": gid,
                        "dimension": "ROWS",
                        "startIndex": end1,  # 0-based: inserts before current end1+1
                        "endIndex": end1 + n_ins,
                    },
                    "inheritFromBefore": True,
                }
            }
        )
        end1 = start1 + need_n - 1
        cur_n = need_n
    elif need_n < cur_n:
        # clear surplus FBA rows only (do not delete — safer for formulas/refs below)
        print("surplus FBA rows %d..%d will clear A/B/D/F only" % (start1 + need_n, end1))

    if requests:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid, body={"requests": requests}
        ).execute()

    # Re-read after insert
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!A1:F120" % SHEET)
        .execute()
        .get("values")
        or []
    )
    start1, end1 = find_fba_block(vals)
    # After insert, block may have grown with blank A — find_fba_block only contiguous FBA.
    # So if we inserted blank rows, they might break continuity if A is empty.
    # Fix: write all target rows starting at original start1 for need_n rows.
    # Re-detect start from first FBA only.
    first_fba = next(i for i, r in enumerate(vals) if cell(r, 0) == "FBA手数料")
    write_start = first_fba + 1  # 1-based

    data_a: List[List[Any]] = []
    data_b: List[List[Any]] = []
    data_d: List[List[Any]] = []
    data_f: List[List[Any]] = []
    for name, fee, remark in TARGET:
        data_a.append(["FBA手数料"])
        data_b.append([name])
        data_d.append([fee])
        data_f.append([remark])

    def upd(rng: str, values: List[List[Any]]) -> None:
        svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range=rng,
            valueInputOption="USER_ENTERED",
            body={"values": values},
        ).execute()

    end_write = write_start + need_n - 1
    upd("'%s'!A%d:A%d" % (SHEET, write_start, end_write), data_a)
    upd("'%s'!B%d:B%d" % (SHEET, write_start, end_write), data_b)
    upd("'%s'!D%d:D%d" % (SHEET, write_start, end_write), data_d)
    upd("'%s'!F%d:F%d" % (SHEET, write_start, end_write), data_f)

    # Clear surplus old FBA rows if any remain below with A=FBA手数料 beyond target
    vals2 = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!A1:F120" % SHEET)
        .execute()
        .get("values")
        or []
    )
    # any FBA手数料 after end_write
    cleared = 0
    for i, row in enumerate(vals2):
        r1 = i + 1
        if r1 <= end_write:
            continue
        if cell(row, 0) != "FBA手数料":
            # stop at first non-FBA after block (販売手数料 etc.)
            if cell(row, 0):
                break
            continue
        upd("'%s'!A%d:B%d" % (SHEET, r1, r1), [["", ""]])
        upd("'%s'!D%d:D%d" % (SHEET, r1, r1), [[""]])
        upd("'%s'!F%d:F%d" % (SHEET, r1, r1), [[""]])
        cleared += 1
    print("wrote FBA rows %d..%d cleared_surplus=%d" % (write_start, end_write, cleared))

    # verify
    vals3 = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=sid,
            range="'%s'!A%d:F%d" % (SHEET, write_start, end_write + 3),
        )
        .execute()
        .get("values")
        or []
    )
    print("--- after ---")
    for i, row in enumerate(vals3):
        print(
            "%d\t%s\t%s\t%s\t%s"
            % (
                write_start + i,
                cell(row, 0),
                cell(row, 1),
                cell(row, 3),
                cell(row, 5),
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
