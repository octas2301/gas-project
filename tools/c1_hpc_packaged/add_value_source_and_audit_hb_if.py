#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""① MAPに「値の取り方」列追加 ② マスタHB〜IFと突合。"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
TOKEN_RW = SCRIPT_DIR / "secrets" / "token_sheets_rw.json"
SHEET_TITLE = "▼設定(Amazonマッピング)"
MASTER_TITLE = "▼商品マスタ(人間作業用)"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# After insert: valueSource becomes col after masterColFallback (index 8), shifting inherit+
# Current headers (0-based):
# 0 productCategory, 1 productType, 2 attrKey, 3 scHeaderJa, 4 scHeaderAlias,
# 5 required, 6 masterColPrimary, 7 masterColFallback, 8 inherit, 9 transform,
# 10 defaultValue, 11 doNotUse, 12 sourceNote, 13 notes, 14 enabled, 15 sampleFilledThisRun

INSERT_AT = 8  # before inherit
OLD_MAP_COLS = 16
NEW_MAP_COLS = 17


def col_letter(n: int) -> str:
    """1-based column number to letters."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def pad(row: List[Any], n: int) -> List[Any]:
    return (list(row) + [""] * n)[:n]


def decide_value_source(row: List[Any]) -> str:
    """attrKey / inherit / transform / sourceNote / masters から初期値を決める。"""
    # old indexes before insert
    attr = str(row[2] if len(row) > 2 else "").strip()
    inherit = str(row[8] if len(row) > 8 else "").strip()
    transform = str(row[9] if len(row) > 9 else "").strip()
    primary = str(row[6] if len(row) > 6 else "").strip()
    fallback = str(row[7] if len(row) > 7 else "").strip()
    default = str(row[10] if len(row) > 10 else "").strip()
    source = str(row[12] if len(row) > 12 else "").strip()
    notes = str(row[13] if len(row) > 13 else "").strip()

    # fixed / role / generated-ish
    if transform == "固定値を使う" or inherit.startswith("継承なし"):
        if "GENERATED" in source.upper() or "generated" in source.lower():
            return "GENERATED"
        if attr in ("sku", "parentage", "inventory", "price", "shipping"):
            if attr == "sku" or attr == "parentage":
                return "GENERATED" if attr != "parentage" else "GENERATED"
            return "GENERATED"
        if attr in ("action", "var_theme", "brand", "id_type", "import_type", "exclusive", "condition", "fulfillment"):
            return "固定"
        if transform == "親／子供を役割で付ける":
            return "GENERATED"
        return "固定"

    if attr in ("sku", "parentage", "price", "inventory", "shipping"):
        return "GENERATED"

    # dropdown-prone food attributes
    dropdown_attrs = {
        "item_form",
        "temperature_rating",
        "unit_uom",
        "item_weight_unit",
        "heat",
        "liquid",
        "hazmat",
        "color",
        "condition",
        "id_type",
        "product_type",
        "browse",
        "expiration_dated",
        "expiration_type",
        "shelf_life_unit",
        "grind_type",
        "flavor",
    }
    if attr in dropdown_attrs:
        if primary or fallback:
            return "マスタ→バルク選択肢"
        return "バルク選択肢のみ"

    if "バルク" in notes or "選択肢" in notes:
        return "マスタ→バルク選択肢"

    if primary or fallback or source == "マスタ":
        return "マスタのみ"

    if default:
        return "固定"

    return "マスタのみ"


def main() -> int:
    cfg = json.loads((SCRIPT_DIR / "config.local.json").read_text(encoding="utf-8"))
    sid = cfg["spreadsheet_id"]
    creds = Credentials.from_authorized_user_file(str(TOKEN_RW), SCOPES)
    svc = build("sheets", "v4", credentials=creds)

    meta = (
        svc.spreadsheets()
        .get(spreadsheetId=sid, fields="sheets(properties(sheetId,title))")
        .execute()
    )
    sheet_id = None
    for sh in meta.get("sheets") or []:
        p = sh.get("properties") or {}
        if p.get("title") == SHEET_TITLE:
            sheet_id = int(p["sheetId"])
            break
    if sheet_id is None:
        print("mapping sheet not found", file=sys.stderr)
        return 2

    # ----- ② HB:IF headers from master -----
    hb_if = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'!HB8:IF8" % MASTER_TITLE)
        .execute()
        .get("values")
        or [[]]
    )
    headers = hb_if[0] if hb_if else []
    hb_if_items: List[Tuple[str, str]] = []
    for i, h in enumerate(headers):
        name = str(h or "").strip()
        col = col_letter(210 + i)  # HB = 210
        hb_if_items.append((col, name))

    # ----- read full mapping sheet -----
    old = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'" % SHEET_TITLE)
        .execute()
        .get("values")
        or []
    )

    # Detect layout starts from row4
    # MAP A-P (0-15), ERRORS R(17), RULES AC(28), COMMON AF(31) from previous localize
    ERR_START_OLD = 17
    RULE_START_OLD = 28
    COMMON_START_OLD = 31

    # After inserting 1 col into MAP, MAP becomes 0-16, shift right blocks by +1
    ERR_START = 18  # S
    RULE_START = 29  # AD
    COMMON_START = 32  # AG
    WIDTH = 35

    def blank() -> List[Any]:
        return [""] * WIDTH

    def place(row: List[Any], start: int, cells: List[Any]) -> List[Any]:
        out = pad(row, WIDTH)
        for i, v in enumerate(cells):
            out[start + i] = v
        return out

    # Parse MAP data rows (row index 8+)
    map_rows: List[List[Any]] = []
    errors: List[List[Any]] = []
    rules: List[List[Any]] = []
    commons: List[List[Any]] = []

    for r in old[8:]:
        rr = pad(r, 40)
        cat = str(rr[0]).strip()
        attr = str(rr[2]).strip()
        if cat and attr and re.match(r"^[a-z][a-z0-9_]*$", attr):
            map_rows.append(pad(rr[:16], 16))
        ec = str(rr[ERR_START_OLD]).strip()
        if ec and (ec[0].isdigit() or ec.startswith("(")):
            errors.append(pad(rr[ERR_START_OLD : ERR_START_OLD + 9], 9))
        rid = str(rr[RULE_START_OLD]).strip()
        if rid and rid not in ("ルール名", "継承・変換で使う名前", "日本語で統一（コード接続時に英ID列を追加可）"):
            if not rid.startswith("===") and rid != "意味":
                rules.append([rid, str(rr[RULE_START_OLD + 1]).strip()])
        c0 = str(rr[COMMON_START_OLD]).strip()
        if c0.isdigit():
            commons.append([c0, str(rr[COMMON_START_OLD + 1]).strip()])

    if not map_rows:
        print("no map rows parsed", file=sys.stderr)
        return 2

    # Insert value source + translate sourceNote optionally
    new_map: List[List[Any]] = []
    for r in map_rows:
        vs = decide_value_source(r)
        new_row = r[:INSERT_AT] + [vs] + r[INSERT_AT:]
        new_map.append(pad(new_row, NEW_MAP_COLS))
        # sync sourceNote lightly if fixed
        # keep as-is

    # Add RULES entries for value source
    extra_rules = [
        ["マスタのみ", "マスタ列だけから取る（バルク選択肢は見ない）"],
        ["バルク選択肢のみ", "純正xlsmの推奨値／Dropdownから選ぶ（マスタは見ない）"],
        ["マスタ→バルク選択肢", "まずマスタ。空または選択肢外ならバルクのプルダウンから選ぶ"],
        ["バルク選択肢→マスタ", "まずバルク選択肢の既定。無ければマスタ"],
        ["固定", "既定値をそのまま使う"],
        ["GENERATED", "GENERATED CSV（または役割行）から取る"],
    ]
    have = {a for a, _ in rules}
    for a, b in extra_rules:
        if a not in have:
            rules.append([a, b])
            have.add(a)

    # Rebuild header block
    map_mnemonic = [
        "いつ効くか",
        "いつ効くか",
        "何を埋めるか",
        "何を埋めるか",
        "何を埋めるか",
        "必須度",
        "どこから取るか",
        "どこから取るか",
        "どこから取るか",  # 値の取り方
        "どう取るか",
        "どう取るか",
        "どこから取るか",
        "どこから取るか",
        "人間向けメモ",
        "人間向けメモ",
        "いつ効くか",
        "人間向けメモ",
    ]
    map_headers = [
        "productCategory",
        "productType",
        "attrKey",
        "scHeaderJa",
        "scHeaderAlias",
        "required",
        "masterColPrimary",
        "masterColFallback",
        "valueSource",
        "inherit",
        "transform",
        "defaultValue",
        "doNotUse",
        "sourceNote",
        "notes",
        "enabled",
        "sampleFilledThisRun",
    ]
    map_meaning = [
        "商品カテゴリー（大きな括り）",
        "商品タイプ（技術値）",
        "内部キー",
        "SCの日本語項目名",
        "SC項目名の別名一覧",
        "必須度",
        "マスタ第1候補列",
        "マスタ第2候補列",
        "値の取り方",
        "継承",
        "変換",
        "既定値",
        "使用禁止のマスタ列",
        "値の出所メモ",
        "補足・注意",
        "有効フラグ",
        "今回の缶飯で埋まったか",
    ]
    map_purpose = [
        "SCの基本カテゴリー（例:食品＆飲料）。人間の台帳の章立て・絞り込み用",
        "xlsmに書くProduct Type（FOOD/SEASONING等）",
        "コード・jsonが使う安定名",
        "Seller Central／xlsm上の見た目の列名",
        "テンプレ更新で揺れても当てる候補（|区切り）",
        "必須／なるべく／任意／参考のみ",
        "まず読むマスタ列名（値の取り方にマスタが含まれるとき）",
        "第1が空のときに読む列",
        "マスタのみ／バルク選択肢のみ／マスタ→バルク選択肢／固定／GENERATED 等",
        "子優先（空なら親）／子のみ／親のみ／継承なし",
        "RULESの日本語ルール名と対応",
        "子も親も空のときの固定値",
        "誤って使ってはいけないマスタ列",
        "マスタ／GENERATED／固定のメモ",
        "運用上の注意",
        "TRUE=使う／FALSE=無効",
        "YES/NO/REF",
    ]

    err_headers = [
        "errorCode",
        "productType",
        "count",
        "symptom",
        "rootCause",
        "fixMaster",
        "fixMap",
        "status",
        "sampleFile",
    ]
    # keep previous meaning rows short
    err_meaning = ["エラーコード", "商品タイプ", "件数", "症状", "原因", "マスタの直し方", "MAPの直し方", "状態", "サマリファイル"]
    err_purpose = [
        "SCコード",
        "対象型",
        "件数",
        "症状",
        "原因",
        "マスタ修正",
        "MAP修正",
        "未対応／対応済／待機",
        "ファイル名",
    ]

    values: List[List[Any]] = []
    r1 = blank()
    r1[0] = "▼設定(Amazonマッピング) — 試行版"
    r1[1] = "食品＆飲料 × FOOD"
    r1[2] = "値の取り方列あり"
    r1[3] = "C1未接続"
    r1[ERR_START] = "左→右: MAP → ERRORS → RULES → 共通認識"
    values.append(r1)

    r2 = blank()
    r2[0] = "使い方"
    r2[1] = "値の取り方でマスタ／バルク選択肢／固定／GENERATEDを切り替える。MAP・ERRORSは下に伸ばす。"
    values.append(r2)

    r3 = blank()
    r3[0] = "色分け"
    r3[1] = "灰=設計キー／青=人間メンテ／緑=メモ"
    r3[2] = "値の取り方は青（人間メンテ）"
    values.append(r3)

    r4 = blank()
    r4[0] = "=== MAP ==="
    r4[ERR_START] = "=== ERRORS（今回缶飯 CK_5beb0cbf67ea_B1 初期） ==="
    r4[RULE_START] = "=== RULES（継承・変換・値の取り方） ==="
    r4[COMMON_START] = "=== 共通認識（触らない正本） ==="
    values.append(r4)

    values.append(
        place(
            place(
                place(place(blank(), 0, map_mnemonic), ERR_START, ["エラー実績（下に伸ばす）"] + [""] * 8),
                RULE_START,
                ["ルール名", "意味"],
            ),
            COMMON_START,
            ["#", "内容"],
        )
    )
    values.append(
        place(
            place(
                place(place(blank(), 0, map_headers), ERR_START, err_headers),
                RULE_START,
                ["ルール名", "意味"],
            ),
            COMMON_START,
            ["番号", "内容"],
        )
    )
    values.append(
        place(
            place(
                place(place(blank(), 0, map_meaning), ERR_START, err_meaning),
                RULE_START,
                ["継承・変換・値の取り方", "人間向け説明"],
            ),
            COMMON_START,
            ["番号", "運用の約束"],
        )
    )
    values.append(
        place(
            place(
                place(place(blank(), 0, map_purpose), ERR_START, err_purpose),
                RULE_START,
                ["日本語で統一", "MAPと一致させる"],
            ),
            COMMON_START,
            ["", "最右。原則変えない"],
        )
    )

    commons = [
        ["1", "マスタで子に値がある項目は子を使う。子が空なら親を使う（子優先（空なら親））"],
        ["2", "バリエ数量・重量・サイズは「子のみ」"],
        ["3", "章立て＝商品カテゴリー、xlsm値＝商品タイプを併記"],
        ["4", "値の取り方でマスタ／バルク選択肢／固定／GENERATEDを明示する"],
        ["5", "本シートは試行用。C1は当面jsonが正"],
    ]

    max_data = max(len(new_map), len(errors), len(rules), len(commons))
    for i in range(max_data):
        row = blank()
        if i < len(new_map):
            row = place(row, 0, new_map[i])
        if i < len(errors):
            row = place(row, ERR_START, errors[i])
        if i < len(rules):
            row = place(row, RULE_START, rules[i])
        if i < len(commons):
            row = place(row, COMMON_START, commons[i])
        values.append(row)

    # write sheet
    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "rowCount": max(220, len(values) + 40),
                                "columnCount": max(45, WIDTH + 2),
                            },
                        },
                        "fields": "gridProperties(rowCount,columnCount)",
                    }
                }
            ]
        },
    ).execute()
    try:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={
                "requests": [
                    {
                        "unmergeCells": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 0,
                                "endRowIndex": 12,
                                "startColumnIndex": 0,
                                "endColumnIndex": WIDTH,
                            }
                        }
                    }
                ]
            },
        ).execute()
    except Exception:
        pass

    svc.spreadsheets().values().clear(
        spreadsheetId=sid, range="'%s'" % SHEET_TITLE
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range="'%s'!A1" % SHEET_TITLE,
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    # formatting + CF (shifted columns)
    gray = {"red": 0.85, "green": 0.85, "blue": 0.85}
    blue = {"red": 0.79, "green": 0.89, "blue": 0.98}
    green = {"red": 0.82, "green": 0.93, "blue": 0.82}
    yellow = {"red": 1.0, "green": 0.95, "blue": 0.8}
    section = {"red": 1.0, "green": 0.90, "blue": 0.70}

    def color_cells(row0: int, cols: List[int], color: dict) -> List[dict]:
        out = []
        for c in cols:
            out.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": row0,
                            "endRowIndex": row0 + 1,
                            "startColumnIndex": c,
                            "endColumnIndex": c + 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": color,
                                "textFormat": {"bold": True},
                                "wrapStrategy": "WRAP",
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,wrapStrategy)",
                    }
                }
            )
        return out

    # Clear old CF then re-add with new column indexes
    meta2 = (
        svc.spreadsheets()
        .get(spreadsheetId=sid, fields="sheets(properties(title),conditionalFormats)")
        .execute()
    )
    n_cf = 0
    for sh in meta2.get("sheets") or []:
        if sh.get("properties", {}).get("title") == SHEET_TITLE:
            n_cf = len(sh.get("conditionalFormats") or [])
    del_reqs = [
        {"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": 0}}
        for _ in range(n_cf)
    ]

    reqs: List[dict] = del_reqs + [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 8},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 4,
                    "endRowIndex": 5,
                    "startColumnIndex": 0,
                    "endColumnIndex": NEW_MAP_COLS,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": yellow,
                        "textFormat": {"bold": True},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        },
    ]
    # human blue: 0 category, 5 required, 6-8 masters+valueSource, 9-10 inherit transform? 
    # 9 inherit, 10 transform, 11 default, 12 doNotUse, 14 notes, 15 enabled
    # gray: 1-4 type/keys
    # green: 13 sourceNote, 16 sample
    reqs.extend(color_cells(5, [0, 5, 6, 7, 8, 9, 10, 11, 12, 14, 15], blue))
    reqs.extend(color_cells(5, [1, 2, 3, 4], gray))
    reqs.extend(color_cells(5, [13, 16], green))

    for start, end in (
        (0, NEW_MAP_COLS),
        (ERR_START, ERR_START + 9),
        (RULE_START, RULE_START + 2),
        (COMMON_START, COMMON_START + 2),
    ):
        reqs.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 3,
                        "endRowIndex": 4,
                        "startColumnIndex": start,
                        "endColumnIndex": end,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": section,
                            "textFormat": {"bold": True},
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat)",
                }
            }
        )
    reqs.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 6,
                    "endRowIndex": 8,
                    "startColumnIndex": 0,
                    "endColumnIndex": WIDTH,
                },
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP",
                        "verticalAlignment": "TOP",
                    }
                },
                "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)",
            }
        }
    )
    for start, end in ((0, 2), (2, 5), (6, 9), (9, 11), (11, 13), (13, 15)):
        reqs.append(
            {
                "mergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 4,
                        "endRowIndex": 5,
                        "startColumnIndex": start,
                        "endColumnIndex": end,
                    },
                    "mergeType": "MERGE_ALL",
                }
            }
        )

    def cf(col: int, text: str, rgb: Tuple[float, float, float], index: int) -> dict:
        return {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": sheet_id,
                            "startRowIndex": 8,
                            "endRowIndex": 500,
                            "startColumnIndex": col,
                            "endColumnIndex": col + 1,
                        }
                    ],
                    "booleanRule": {
                        "condition": {
                            "type": "TEXT_EQ",
                            "values": [{"userEnteredValue": text}],
                        },
                        "format": {
                            "backgroundColor": {
                                "red": rgb[0],
                                "green": rgb[1],
                                "blue": rgb[2],
                            },
                            "textFormat": {"bold": True},
                        },
                    },
                },
                "index": index,
            }
        }

    # column indexes after insert: required=5, valueSource=8, inherit=9, transform=10, sourceNote=13
    cf_specs = [
        (5, "MUST", (0.96, 0.80, 0.80)),
        (9, "継承なし（固定／別経路）", (0.88, 0.82, 0.96)),
        (9, "子のみ", (1.00, 0.90, 0.70)),
        (9, "親のみ", (0.72, 0.91, 0.92)),
        (10, "固定値を使う", (0.78, 0.88, 0.99)),
        (13, "マスタ", (0.78, 0.93, 0.80)),
        # valueSource colors
        (8, "マスタのみ", (0.85, 0.95, 0.85)),
        (8, "マスタ→バルク選択肢", (1.00, 0.94, 0.80)),
        (8, "バルク選択肢のみ", (1.00, 0.90, 0.75)),
        (8, "固定", (0.80, 0.88, 0.98)),
        (8, "GENERATED", (0.90, 0.85, 0.95)),
    ]
    for i, (c, t, rgb) in enumerate(cf_specs):
        reqs.append(cf(c, t, rgb, i))

    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": reqs}
    ).execute()

    # ----- audit HB-IF vs MAP -----
    mapped_master_names = set()
    for r in new_map:
        for idx in (6, 7):  # primary, fallback
            raw = str(r[idx] if idx < len(r) else "")
            for part in re.split(r"[|／/]", raw):
                p = part.strip()
                if p:
                    mapped_master_names.add(p)
        # also scHeaderJa and aliases sometimes match master
        for idx in (3, 4):
            raw = str(r[idx] if idx < len(r) else "")
            for part in re.split(r"[|]", raw):
                p = part.strip()
                if p:
                    mapped_master_names.add(p)

    def is_charcount(name: str) -> bool:
        n = name.replace(" ", "")
        if not n:
            return True
        keys = ["文字数", "文字カウント", "文字ｶｳﾝﾄ", "len(", "カウント", "字数"]
        return any(k in n for k in keys)

    def is_mapped(name: str) -> bool:
        if not name or is_charcount(name):
            return True  # excluded
        if name in mapped_master_names:
            return True
        # soft match
        for m in mapped_master_names:
            if name == m or name in m or m in name:
                return True
        # strip ▼マスタ()
        m2 = re.sub(r"^▼マスタ\((.+)\)$", r"\1", name)
        if m2 != name and is_mapped(m2):
            return True
        return False

    missing = []
    excluded = []
    present = []
    for col, name in hb_if_items:
        if not name:
            continue
        if is_charcount(name):
            excluded.append((col, name))
            continue
        if is_mapped(name):
            present.append((col, name))
        else:
            missing.append((col, name))

    report = {
        "hb_if_total_named": len([1 for _, n in hb_if_items if n]),
        "excluded_charcount": excluded,
        "mapped_ok": present,
        "missing": missing,
        "value_source_counts": {},
    }
    from collections import Counter

    report["value_source_counts"] = dict(Counter(r[8] for r in new_map))

    out = SCRIPT_DIR / "_hb_if_audit.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("=== ① 値の取り方 初期値件数 ===")
    for k, v in sorted(report["value_source_counts"].items(), key=lambda x: -x[1]):
        print(" ", v, k)
    print("=== ② HB〜IF 監査 ===")
    print("named", report["hb_if_total_named"], "excluded_charcount", len(excluded), "mapped", len(present), "missing", len(missing))
    print("--- 除外（文字数系）---")
    for c, n in excluded:
        print(c, n)
    print("--- マッピング漏れ ---")
    for c, n in missing:
        print(c, n)
    print("DONE sheet updated")
    print(
        "URL: https://docs.google.com/spreadsheets/d/%s/edit#gid=%s" % (sid, sheet_id)
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
