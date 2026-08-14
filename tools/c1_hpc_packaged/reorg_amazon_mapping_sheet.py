#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""▼設定(Amazonマッピング) を MAP | ERRORS | RULES | 共通認識 の横並びに再配置する。"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, List

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
TOKEN_RW = SCRIPT_DIR / "secrets" / "token_sheets_rw.json"
SHEET_TITLE = "▼設定(Amazonマッピング)"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 0-based column indexes
MAP_START = 0  # A
MAP_END = 14  # O
ERR_START = 16  # Q (P=15 blank)
RULE_START = 27  # AB (Z=25 blank, AA unused gap -> use AA=26 blank, AB=27)
# Actually: Q=16 ... Y=24 (9 cols), Z=25 blank, AA=26 blank, AB=27 RULES, AC=28, AD=29 blank, AE=30 共通認識
RULE_START = 26  # AA
COMMON_START = 29  # AD


def col_letter(idx: int) -> str:
    n = idx + 1
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def pad(row: List[Any], n: int) -> List[Any]:
    out = list(row) + [""] * max(0, n - len(row))
    return out[:n]


def main() -> int:
    cfg = json.loads((SCRIPT_DIR / "config.local.json").read_text(encoding="utf-8"))
    sid = cfg["spreadsheet_id"]
    creds = Credentials.from_authorized_user_file(str(TOKEN_RW), SCOPES)
    svc = build("sheets", "v4", credentials=creds)

    meta = (
        svc.spreadsheets()
        .get(spreadsheetId=sid, fields="sheets(properties(sheetId,title,gridProperties))")
        .execute()
    )
    sheet_id = None
    for sh in meta.get("sheets") or []:
        p = sh.get("properties") or {}
        if p.get("title") == SHEET_TITLE:
            sheet_id = int(p["sheetId"])
            break
    if sheet_id is None:
        print("sheet not found", file=sys.stderr)
        return 2

    snap_path = SCRIPT_DIR / "_amazon_map_snapshot.json"
    if snap_path.is_file():
        old = json.loads(snap_path.read_text(encoding="utf-8"))
    else:
        old = (
            svc.spreadsheets()
            .values()
            .get(spreadsheetId=sid, range="'%s'" % SHEET_TITLE)
            .execute()
            .get("values")
            or []
        )

    # Extract MAP data rows (attrKey starts after header block)
    map_data: List[List[Any]] = []
    rules: List[List[Any]] = []
    commons: List[List[Any]] = []
    errors: List[List[Any]] = []
    mode = None
    for r in old:
        if not r:
            continue
        a0 = str(r[0]) if r else ""
        if a0.startswith("=== MAP"):
            mode = "map_header"
            continue
        if a0.startswith("=== RULES"):
            mode = "rules"
            continue
        if a0.startswith("=== 共通認識"):
            mode = "common"
            continue
        if a0.startswith("=== ERRORS"):
            mode = "errors"
            continue
        if mode == "map_header":
            # skip until first FOOD data row
            if len(r) >= 2 and str(r[0]) == "FOOD":
                mode = "map"
                map_data.append(pad(r, 15))
            continue
        if mode == "map":
            if str(r[0]) == "FOOD":
                map_data.append(pad(r, 15))
            continue
        if mode == "rules":
            if a0 == "ruleId":
                continue
            if a0:
                rules.append(pad(r, 2))
            continue
        if mode == "common":
            if a0:
                commons.append(pad(r, 2))
            continue
        if mode == "errors":
            if a0 == "errorCode":
                continue
            if a0:
                errors.append(pad(r, 9))
            continue

    print("map_data=%d errors=%d rules=%d commons=%d" % (len(map_data), len(errors), len(rules), len(commons)))

    map_headers = [
        "productType",
        "attrKey",
        "scHeaderJa",
        "scHeaderAlias",
        "required",
        "masterColPrimary",
        "masterColFallback",
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
        "商品タイプ（適用範囲）",
        "内部キー",
        "SCの日本語項目名",
        "SC項目名の別名一覧",
        "必須度",
        "マスタ第1候補列",
        "マスタ第2候補列",
        "継承ルール",
        "変換ルールID",
        "既定値",
        "使用禁止のマスタ列",
        "値の出所メモ",
        "補足・注意",
        "有効フラグ",
        "今回の缶飯で埋まったか",
    ]
    map_purpose = [
        "この行が効くAmazon Product Type（FOOD/SEASONING/*）。PTごとに行を分ける",
        "コード・jsonが使う安定名。SC日本語名が変わってもここは変えない",
        "Seller Central／xlsm上の見た目の列名。人間が照合するための表示名",
        "テンプレ更新で列名が揺れても当てる候補（|区切り）。項目名解決用",
        "MUST=必須／SHOULD=なるべく／OPT=任意／REF=参考のみ",
        "まず読むマスタ列名",
        "第1が空のときに読む列（|区切り可）。GENERATED項目名も可",
        "CHILD_THEN_PARENT=子→親／CHILD_ONLY=子のみ／NO_INHERIT=継承なし 等",
        "値の加工方法（RULESのruleIdと対応）",
        "子も親も空のときに入れる固定値。必須で空ならエラーになり得る",
        "誤って使ってはいけないマスタ列（再発防止）",
        "マスタ／GENERATED／固定のどれから来るか（読み取り用）",
        "運用上の注意書き",
        "TRUE=使う／FALSE=無効（試しに切るとき）",
        "YES=今回PACKAGEDに値あり／NO=未出力だが紐付け残す／REF=参考行",
    ]
    # row5 mnemonic per column (will merge visually via values + formatting)
    map_mnemonic = [
        "何を埋めるか",
        "何を埋めるか",
        "何を埋めるか",
        "何を埋めるか",
        "必須度",
        "どこから取るか",
        "どこから取るか",
        "どう取るか",
        "どう取るか",
        "どこから取るか",
        "どこから取るか",
        "人間向けメモ",
        "人間向けメモ",
        "いつ効くか",
        "人間向けメモ",
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
    err_meaning = [
        "エラーコード",
        "商品タイプ",
        "件数",
        "症状",
        "原因",
        "マスタの直し方",
        "MAPの直し方",
        "状態",
        "サマリファイル",
    ]
    err_purpose = [
        "SC processing-summary のコード",
        "対象PT",
        "何件で出たか",
        "何が起きたか一言",
        "マスタ／マップ／SCの切り分け",
        "人間がマスタで直すこと",
        "MAP行の直し方",
        "OPEN／FIXED／WAIT",
        "根拠ファイル名",
    ]

    # Build wide grid: 31 columns (A-AE)
    width = 31
    values: List[List[Any]] = []

    def blank_row() -> List[Any]:
        return [""] * width

    def place(base_row: List[Any], start: int, cells: List[Any]) -> List[Any]:
        row = pad(base_row, width)
        for i, v in enumerate(cells):
            row[start + i] = v
        return row

    r1 = blank_row()
    r1[0] = "▼設定(Amazonマッピング) — 試行版"
    r1[1] = "缶飯 FOOD 初期投入"
    r1[2] = "subBatchId=CK_5beb0cbf67ea_B1"
    r1[3] = "継承: 子→親"
    r1[4] = "C1未接続"
    r1[ERR_START] = "左→右: MAP → ERRORS → RULES → 共通認識"
    values.append(r1)

    r2 = blank_row()
    r2[0] = "使い方"
    r2[1] = "MAPを下に足しながら試す。enabled=FALSEで無効。C1は当面 json が正。シート確定後に接続。"
    r2[ERR_START] = "ERRORSも下に追記する"
    values.append(r2)

    r3 = blank_row()
    r3[0] = "色分け"
    r3[1] = "灰=設計キー（相談して変更）"
    r3[2] = "青=人間がメンテする列"
    r3[3] = "緑=メモ・参考（自由に直してよい）"
    values.append(r3)

    r4 = blank_row()
    r4[0] = "=== MAP ==="
    r4[ERR_START] = "=== ERRORS（今回缶飯 CK_5beb0cbf67ea_B1 初期） ==="
    r4[RULE_START] = "=== RULES（継承・変換） ==="
    r4[COMMON_START] = "=== 共通認識（触らない正本） ==="
    values.append(r4)

    r5 = place(blank_row(), MAP_START, map_mnemonic)
    r5 = place(r5, ERR_START, ["エラー実績（下に伸ばす）"] + [""] * 8)
    r5 = place(r5, RULE_START, ["ルールID", "意味"])
    r5 = place(r5, COMMON_START, ["#", "内容"])
    values.append(r5)

    r6 = place(blank_row(), MAP_START, map_headers)
    r6 = place(r6, ERR_START, err_headers)
    r6 = place(r6, RULE_START, ["ruleId", "意味"])
    r6 = place(r6, COMMON_START, ["no", "content"])
    values.append(r6)

    r7 = place(blank_row(), MAP_START, map_meaning)
    r7 = place(r7, ERR_START, err_meaning)
    r7 = place(r7, RULE_START, ["ルール識別子", "人間向け説明"])
    r7 = place(r7, COMMON_START, ["番号", "運用の約束"])
    values.append(r7)

    r8 = place(blank_row(), MAP_START, map_purpose)
    r8 = place(r8, ERR_START, err_purpose)
    r8 = place(r8, RULE_START, ["transform/inheritで参照", "変更時はMAPと整合を取る"])
    r8 = place(r8, COMMON_START, ["", "最右。原則ここは変えない"])
    values.append(r8)

    max_data = max(len(map_data), len(errors), len(rules), len(commons))
    for i in range(max_data):
        row = blank_row()
        if i < len(map_data):
            row = place(row, MAP_START, map_data[i])
        if i < len(errors):
            row = place(row, ERR_START, errors[i])
        if i < len(rules):
            row = place(row, RULE_START, rules[i])
        if i < len(commons):
            row = place(row, COMMON_START, commons[i])
        values.append(row)

    # Ensure grid size
    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "rowCount": max(200, len(values) + 50),
                                "columnCount": max(40, width + 2),
                            },
                        },
                        "fields": "gridProperties(rowCount,columnCount)",
                    }
                }
            ]
        },
    ).execute()

    # Clear then write
    svc.spreadsheets().values().clear(
        spreadsheetId=sid, range="'%s'" % SHEET_TITLE
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range="'%s'!A1" % SHEET_TITLE,
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    # Formatting
    # Colors
    gray = {"red": 0.85, "green": 0.85, "blue": 0.85}  # design
    blue = {"red": 0.79, "green": 0.89, "blue": 0.98}  # human
    green = {"red": 0.82, "green": 0.93, "blue": 0.82}  # memo
    yellow = {"red": 1.0, "green": 0.95, "blue": 0.8}  # mnemonic row
    section = {"red": 1.0, "green": 0.90, "blue": 0.70}

    def color_cols(row0: int, start_col: int, col_indexes: List[int], color: dict) -> List[dict]:
        reqs = []
        for c in col_indexes:
            reqs.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": row0,
                            "endRowIndex": row0 + 1,
                            "startColumnIndex": start_col + c,
                            "endColumnIndex": start_col + c + 1,
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
        return reqs

    requests: List[dict] = []
    # freeze rows 1-8
    requests.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 8},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        }
    )
    # mnemonic row (row5 = index 4) yellow
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 4,
                    "endRowIndex": 5,
                    "startColumnIndex": 0,
                    "endColumnIndex": 15,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": yellow,
                        "textFormat": {"bold": True},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        }
    )
    # header row6 colors
    # A-E gray (0-4), F-K blue except L, M-N blue, L and O green
    # human blue: F G H I J K M N = 5,6,7,8,9,10,12,13
    # gray: A B C D E = 0-4
    # green: L O = 11,14
    requests.extend(color_cols(5, 0, [0, 1, 2, 3, 4], gray))
    requests.extend(color_cols(5, 0, [5, 6, 7, 8, 9, 10, 12, 13], blue))
    requests.extend(color_cols(5, 0, [11, 14], green))
    # meaning/purpose wrap
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 6,
                    "endRowIndex": 8,
                    "startColumnIndex": 0,
                    "endColumnIndex": width,
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
    # section headers row4
    for start, end in (
        (MAP_START, MAP_END + 1),
        (ERR_START, ERR_START + 9),
        (RULE_START, RULE_START + 2),
        (COMMON_START, COMMON_START + 2),
    ):
        requests.append(
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
    # ERRORS header row6 light red/pink for visibility
    err_header_bg = {"red": 0.98, "green": 0.85, "blue": 0.85}
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 5,
                    "endRowIndex": 6,
                    "startColumnIndex": ERR_START,
                    "endColumnIndex": ERR_START + 9,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": err_header_bg,
                        "textFormat": {"bold": True},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        }
    )
    # RULES / 共通認識 headers
    rule_bg = {"red": 0.93, "green": 0.90, "blue": 0.98}
    common_bg = {"red": 0.90, "green": 0.90, "blue": 0.90}
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 5,
                    "endRowIndex": 6,
                    "startColumnIndex": RULE_START,
                    "endColumnIndex": RULE_START + 2,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": rule_bg,
                        "textFormat": {"bold": True},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        }
    )
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 5,
                    "endRowIndex": 6,
                    "startColumnIndex": COMMON_START,
                    "endColumnIndex": COMMON_START + 2,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": common_bg,
                        "textFormat": {"bold": True},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        }
    )
    # merge mnemonic groups on row5
    for start, end in ((0, 4), (5, 7), (7, 9), (9, 11), (11, 13)):
        requests.append(
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
    # row height for purpose
    requests.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 7,
                    "endIndex": 8,
                },
                "properties": {"pixelSize": 72},
                "fields": "pixelSize",
            }
        }
    )
    requests.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 6,
                    "endIndex": 7,
                },
                "properties": {"pixelSize": 36},
                "fields": "pixelSize",
            }
        }
    )

    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": requests}
    ).execute()

    print(
        "DONE layout MAP=A-O ERRORS=%s-%s RULES=%s-%s COMMON=%s-%s"
        % (
            col_letter(ERR_START),
            col_letter(ERR_START + 8),
            col_letter(RULE_START),
            col_letter(RULE_START + 1),
            col_letter(COMMON_START),
            col_letter(COMMON_START + 1),
        )
    )
    print(
        "URL: https://docs.google.com/spreadsheets/d/%s/edit#gid=%s" % (sid, sheet_id)
    )
    # cleanup snapshot
    try:
        snap_path.unlink(missing_ok=True)
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
