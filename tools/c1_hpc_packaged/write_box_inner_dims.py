#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""00_設定マスタ ダンボールサイズ表に内寸A/B/C_mm（E/F/G）を書く。"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
SHEET = "00_設定マスタ"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def parse_dims_mm(text: str) -> Optional[Tuple[int, int, int]]:
    """C列テキストから3辺(mm)を推定。失敗時 None。"""
    if not text:
        return None
    s = str(text).strip()
    # 口折・マチ付き紙袋: 幅×マチ×高さ
    m = re.search(
        r"幅\s*(\d+)\s*[×xX]\s*マチ\s*(\d+)\s*[×xX]\s*高さ\s*(\d+)",
        s,
    )
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    # 長さ/幅/深さ mm
    m = re.search(
        r"長さ\s*(\d+)\s*[×xX]\s*幅\s*(\d+)\s*[×xX]\s*深さ\s*(\d+)\s*mm",
        s,
        re.I,
    )
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    # 縦/横/厚さ cm
    m = re.search(
        r"縦\s*([\d.]+)\s*cm\s*[×xX]\s*横\s*([\d.]+)\s*cm\s*[×xX]\s*厚さ\s*([\d.]+)\s*cm",
        s,
        re.I,
    )
    if m:
        return (
            int(round(float(m.group(1)) * 10)),
            int(round(float(m.group(2)) * 10)),
            int(round(float(m.group(3)) * 10)),
        )
    # a×b×c mm（単位なし数字3つ + mm）
    m = re.search(r"(\d+)\s*[×xX]\s*(\d+)\s*[×xX]\s*(\d+)\s*mm", s, re.I)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    # 封筒系: 口幅×長さ／高さ（2辺）＋折返し。第3辺は薄さ既定10mm（手修正可）
    m = re.search(
        r"口幅\s*(\d+)\s*(?:mm)?\s*[×xX]\s*(?:長さ|高さ)\s*(\d+)\s*(?:mm)?",
        s,
    )
    if m:
        return int(m.group(1)), int(m.group(2)), 10
    return None


EXCLUDE_NAMES = {"商品入荷箱", "Nekopos封筒（他）"}


def pack_type(name: str, text: str) -> str:
    n = (name or "") + " " + (text or "")
    if (name or "").strip() in EXCLUDE_NAMES:
        return "exclude"
    if re.search(r"紙袋|封筒|クッション|ビニール|袋", n):
        return "soft"
    return "rigid"


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    cfg = json.loads((SCRIPT_DIR / "config.local.json").read_text(encoding="utf-8"))
    sid = cfg["spreadsheet_id"]
    creds = Credentials.from_authorized_user_file(
        str(SCRIPT_DIR / "secrets" / "token_sheets_rw.json"), SCOPES
    )
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!A50:I90" % SHEET)
        .execute()
        .get("values")
        or []
    )

    header_row_1based = None
    data_start_i = None
    for i, row in enumerate(vals):
        a = str(row[0] if len(row) > 0 else "").strip()
        b = str(row[1] if len(row) > 1 else "").strip()
        if a == "ダンボールサイズ" or b == "ダンボールサイズ":
            header_row_1based = 50 + i
            data_start_i = i + 1
            break
    if header_row_1based is None:
        print("ダンボールサイズ ヘッダが見つかりません")
        return 1

    # ヘッダ E/F/G/H
    hdr_range = "'%s'!E%d:H%d" % (SHEET, header_row_1based, header_row_1based)
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range=hdr_range,
        valueInputOption="USER_ENTERED",
        body={"values": [["内寸A_mm", "内寸B_mm", "内寸C_mm", "梱包種別"]]},
    ).execute()
    print("header written", hdr_range)

    updates: List[List[object]] = []
    first_data_row = header_row_1based + 1
    last_data_row = first_data_row - 1
    for j in range(data_start_i, len(vals)):
        row = vals[j]
        b = str(row[1] if len(row) > 1 else "").strip()
        if not b:
            break
        c = str(row[2] if len(row) > 2 else "").strip()
        dims = parse_dims_mm(c)
        ptype = pack_type(b, c)
        if dims:
            updates.append([dims[0], dims[1], dims[2], ptype])
            print("OK", b, dims, ptype)
        else:
            updates.append(["", "", "", ptype])
            print("NEED_HAND", b, "|", c[:60], "|", ptype)
        last_data_row = 50 + j

    if not updates:
        print("データ行なし")
        return 1

    body_range = "'%s'!E%d:H%d" % (SHEET, first_data_row, last_data_row)
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range=body_range,
        valueInputOption="USER_ENTERED",
        body={"values": updates},
    ).execute()
    print("data written", body_range, "rows", len(updates))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
