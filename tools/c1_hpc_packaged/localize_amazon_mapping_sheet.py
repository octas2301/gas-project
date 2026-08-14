#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Amazonマッピング: 商品カテゴリー追加＋継承／RULES日本語化。"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
TOKEN_RW = SCRIPT_DIR / "secrets" / "token_sheets_rw.json"
SHEET_TITLE = "▼設定(Amazonマッピング)"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

INHERIT_JA = {
    "CHILD_THEN_PARENT": "子優先（空なら親）",
    "CHILD_ONLY": "子のみ",
    "PARENT_ONLY": "親のみ",
    "NO_INHERIT": "継承なし（固定／別経路）",
}

TRANSFORM_JA = {
    "PARSE_SET_COUNT": "セット数を数値化",
    "PARSE_WEIGHT_FROM_SIZE": "サイズ名から重量を取る",
    "USE_SET_COUNT": "ユニット数＝セット数",
    "TITLE_DEDUP_MAX75": "商品名の重複削除・75字以内",
    "HIGHLIGHT_IF_TITLE_LE75": "タイトル75字以内のときだけハイライト",
    "KW_JOIN_1SLOT": "検索KWを1枠に結合",
    "FALLBACK_CHILD_SKU": "空なら子SKU",
    "YES_NO_JP": "はい／いいえに正規化",
    "FIXED": "固定値を使う",
    "ROLE": "親／子供を役割で付ける",
    "DIGITS_ONLY": "数字のみ",
}

# MAP layout after change: A-P (16 cols)
MAP_COLS = 16
ERR_START = 17  # R
RULE_START = 28  # AC
COMMON_START = 31  # AF
WIDTH = 33

CATEGORY_FOR_FOOD = "食品＆飲料"


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


def transform_ja(raw: str) -> str:
    s = str(raw or "").strip()
    if not s:
        return ""
    if s in TRANSFORM_JA:
        return TRANSFORM_JA[s]
    m = re.match(r"SPLIT_PT_URL_(\d+)$", s)
    if m:
        return "その他画像URLを分割（%s枚目）" % m.group(1)
    # already Japanese
    return s


def inherit_ja(raw: str) -> str:
    s = str(raw or "").strip()
    return INHERIT_JA.get(s, s)


def place(base_row: List[Any], start: int, cells: List[Any]) -> List[Any]:
    row = pad(base_row, WIDTH)
    for i, v in enumerate(cells):
        row[start + i] = v
    return row


def blank() -> List[Any]:
    return [""] * WIDTH


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
        print("sheet not found", file=sys.stderr)
        return 2

    old = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'%s'" % SHEET_TITLE)
        .execute()
        .get("values")
        or []
    )

    # Parse current wide layout: MAP A-O (15), ERRORS Q(16), RULES AA(26), COMMON AD(29)
    map_data: List[List[Any]] = []
    errors: List[List[Any]] = []
    rules: List[List[Any]] = []
    commons: List[List[Any]] = []

    for r in old[8:]:  # data from row9
        rr = pad(r, 31)
        if str(rr[0]).strip() == "FOOD" or (
            str(rr[0]).strip() and str(rr[1]).strip() and str(rr[0]) not in ("",)
            and str(rr[1])
            not in (
                "errorCode",
                "ruleId",
                "no",
                "ルールID",
                "#",
            )
            and not str(rr[0]).startswith("===")
        ):
            # MAP row if col B looks like attrKey (sku, product_type, ...)
            attr = str(rr[1]).strip()
            if attr and re.match(r"^[a-z][a-z0-9_]*$", attr):
                map_data.append(pad(rr[:15], 15))
        # ERRORS at Q=16
        if str(rr[16]).strip() and str(rr[16]) not in (
            "errorCode",
            "エラーコード",
            "エラー実績（下に伸ばす）",
        ):
            if str(rr[16]) not in ("SC processing-summary のコード",):
                # skip meaning/purpose rows if any leaked - data rows have codes
                code = str(rr[16]).strip()
                if code and code not in ("エラーコード",) and not code.startswith("==="):
                    if code[0].isdigit() or code.startswith("(") or code in (
                        "18367",
                        "13013",
                        "100730",
                        "100470",
                        "100476",
                        "20014",
                        "8007",
                        "(human)",
                    ):
                        errors.append(pad(rr[16:25], 9))
        # RULES at AA=26
        rid = str(rr[26]).strip()
        if rid and rid not in (
            "ruleId",
            "ルールID",
            "ルール識別子",
            "transform/inheritで参照",
        ):
            if not rid.startswith("==="):
                rules.append([rid, str(rr[27]).strip()])
        # COMMON at AD=29
        c0 = str(rr[29]).strip()
        if c0 and c0 not in ("no", "#", "番号", ""):
            if c0.isdigit() or c0 in ("1", "2", "3", "4", "5"):
                commons.append([c0, str(rr[30]).strip()])

    # Fallback parse if map_data empty: scan A column FOOD
    if not map_data:
        for r in old:
            rr = pad(r, 15)
            if str(rr[0]) == "FOOD" and re.match(r"^[a-z]", str(rr[1])):
                map_data.append(pad(rr, 15))

    # Rebuild rules in Japanese (canonical list + any extras)
    rules_ja = [
        ["子優先（空なら親）", "子SKU行に値あり→子。空→親。両方空→既定値。必須で既定も空ならエラー"],
        ["子のみ", "子のみ参照。親の総量などをバリエ子へ流さない（入数・重量・サイズ）"],
        ["親のみ", "親のみ（参考・禁止列の説明用）"],
        ["継承なし（固定／別経路）", "固定値やGENERATED役割。マスタ継承しない"],
        ["セット数を数値化", "A.セット商品数「3個で1セット」→3。なければGENERATED setCount"],
        ["サイズ名から重量を取る", "バリエーション値「3缶/480g」→480"],
        ["ユニット数＝セット数", "ユニット数にセット数（缶数）を使う"],
        ["商品名の重複削除・75字以内", "商品名の重複語削除。ハイライト使用時は75文字以内"],
        ["タイトル75字以内のときだけハイライト", "タイトルが75超ならハイライトを出さない（100476）"],
        ["検索KWを1枠に結合", "検索KWは1枠に空白結合（99016）"],
        ["空なら子SKU", "メーカー品番が空なら子SKUを入れる"],
        ["はい／いいえに正規化", "はい／いいえへそろえる"],
        ["固定値を使う", "既定値をそのまま書く"],
        ["親／子供を役割で付ける", "親行=親、子行=子供"],
        ["数字のみ", "価格などから数字だけ抜く"],
        ["その他画像URLを分割（1〜8枚目）", "Amazon PT URLを|分割して各画像列へ"],
    ]
    # keep unknown old rules translated
    known = {a for a, _ in rules_ja}
    for rid, meaning in rules:
        ja = inherit_ja(rid)
        if ja == rid:
            ja = transform_ja(rid)
        if ja not in known and ja:
            rules_ja.append([ja, meaning or ja])
            known.add(ja)

    commons_ja = [
        ["1", "マスタで子に値がある項目は子を使う。子が空なら親を使う（継承＝子優先（空なら親））"],
        ["2", "バリエ数量・重量・サイズは「子のみ」（親総個数／総重量を流用しない）"],
        ["3", "管理の見出しは商品カテゴリー（例:食品＆飲料）。実際にxlsmへ書く型は商品タイプ（FOOD等）を併記する"],
        ["4", "本シートは試行用。C1実行エンジンはまだjson優先。シート確定後にコード接続予定"],
        ["5", "列番号は持たない。SC項目名の別名で項目名解決する。Yahooマッピングと違い縦型・カテゴリー＋型単位"],
    ]

    # Convert map rows: insert category, translate inherit/transform
    new_map: List[List[Any]] = []
    for r in map_data:
        pt = str(r[0]).strip()
        cat = CATEGORY_FOR_FOOD if pt in ("FOOD", "SEASONING", "HERB", "FISH", "VEGETABLE", "SAUCE") else ""
        if pt == "SEASONING":
            cat = "食品＆飲料"
        inherit = inherit_ja(str(r[7]))
        transform = transform_ja(str(r[8]))
        # defaultValue for product_type row: keep FOOD; category is separate
        new_map.append(
            [
                cat,
                r[0],  # productType
                r[1],
                r[2],
                r[3],
                r[4],
                r[5],
                r[6],
                inherit,
                transform,
                r[9],
                r[10],
                r[11],
                r[12],
                r[13],
                r[14],
            ]
        )

    print(
        "parsed map=%d errors=%d rules_in=%d -> map_out=%d"
        % (len(map_data), len(errors), len(rules), len(new_map))
    )
    if not new_map:
        print("no map data", file=sys.stderr)
        return 2

    map_mnemonic = [
        "いつ効くか",
        "いつ効くか",
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
    map_headers = [
        "productCategory",
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
        "商品カテゴリー（大きな括り）",
        "商品タイプ（技術値）",
        "内部キー",
        "SCの日本語項目名",
        "SC項目名の別名一覧",
        "必須度",
        "マスタ第1候補列",
        "マスタ第2候補列",
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
        "xlsmに書くProduct Type（FOOD/SEASONING等）。カテゴリーの下で型別に行を分ける",
        "コード・jsonが使う安定名。SC日本語名が変わってもここは変えない",
        "Seller Central／xlsm上の見た目の列名。人間が照合するための表示名",
        "テンプレ更新で列名が揺れても当てる候補（|区切り）。項目名解決用",
        "必須／なるべく／任意／参考のみ",
        "まず読むマスタ列名",
        "第1が空のときに読む列（|区切り可）",
        "子優先（空なら親）／子のみ／親のみ／継承なし（固定／別経路）",
        "RULESの日本語ルール名と対応する加工方法",
        "子も親も空のときに入れる固定値",
        "誤って使ってはいけないマスタ列",
        "マスタ／GENERATED／固定のどれから来るか",
        "運用上の注意書き",
        "TRUE=使う／FALSE=無効",
        "YES=今回値あり／NO=未出力だが紐付け／REF=参考",
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
        "対象の商品タイプ",
        "何件で出たか",
        "何が起きたか一言",
        "マスタ／マップ／SCの切り分け",
        "人間がマスタで直すこと",
        "MAP行の直し方",
        "未対応／対応済／待機",
        "根拠ファイル名",
    ]

    values: List[List[Any]] = []
    r1 = blank()
    r1[0] = "▼設定(Amazonマッピング) — 試行版"
    r1[1] = "缶飯：食品＆飲料 × FOOD"
    r1[2] = "subBatchId=CK_5beb0cbf67ea_B1"
    r1[3] = "継承・変換は日本語"
    r1[4] = "C1未接続"
    r1[ERR_START] = "左→右: MAP → ERRORS → RULES → 共通認識"
    values.append(r1)

    r2 = blank()
    r2[0] = "使い方"
    r2[1] = "商品カテゴリーで章立てし、商品タイプで型別行を置く。MAP/ERRORSは下に伸ばす。"
    r2[ERR_START] = "ERRORSも下に追記"
    values.append(r2)

    r3 = blank()
    r3[0] = "色分け"
    r3[1] = "灰=設計キー（相談して変更）"
    r3[2] = "青=人間がメンテする列"
    r3[3] = "緑=メモ・参考"
    values.append(r3)

    r4 = blank()
    r4[0] = "=== MAP ==="
    r4[ERR_START] = "=== ERRORS（今回缶飯 CK_5beb0cbf67ea_B1 初期） ==="
    r4[RULE_START] = "=== RULES（継承・変換・日本語） ==="
    r4[COMMON_START] = "=== 共通認識（触らない正本） ==="
    values.append(r4)

    values.append(place(place(place(place(blank(), 0, map_mnemonic), ERR_START, ["エラー実績（下に伸ばす）"] + [""] * 8), RULE_START, ["ルール名", "意味"]), COMMON_START, ["#", "内容"]))
    values.append(place(place(place(place(blank(), 0, map_headers), ERR_START, err_headers), RULE_START, ["ルール名", "意味"]), COMMON_START, ["番号", "内容"]))
    values.append(place(place(place(place(blank(), 0, map_meaning), ERR_START, err_meaning), RULE_START, ["継承・変換で使う名前", "人間向け説明"]), COMMON_START, ["番号", "運用の約束"]))
    values.append(place(place(place(place(blank(), 0, map_purpose), ERR_START, err_purpose), RULE_START, ["日本語で統一（コード接続時に英ID列を追加可）", "MAPの継承／変換と一致させる"]), COMMON_START, ["", "最右。原則ここは変えない"]))

    # normalize error status to Japanese lightly
    status_ja = {"OPEN": "未対応", "FIXED": "対応済", "WAIT": "待機"}
    err_out = []
    for e in errors:
        ee = pad(e, 9)
        ee[7] = status_ja.get(str(ee[7]), ee[7])
        err_out.append(ee)

    max_data = max(len(new_map), len(err_out), len(rules_ja), len(commons_ja))
    for i in range(max_data):
        row = blank()
        if i < len(new_map):
            row = place(row, 0, new_map[i])
        if i < len(err_out):
            row = place(row, ERR_START, err_out[i])
        if i < len(rules_ja):
            row = place(row, RULE_START, rules_ja[i])
        if i < len(commons_ja):
            row = place(row, COMMON_START, commons_ja[i])
        values.append(row)

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
                                "columnCount": max(40, WIDTH + 2),
                            },
                        },
                        "fields": "gridProperties(rowCount,columnCount)",
                    }
                }
            ]
        },
    ).execute()

    svc.spreadsheets().values().clear(
        spreadsheetId=sid, range="'%s'" % SHEET_TITLE
    ).execute()
    # unmerge all first
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
                                "endRowIndex": 20,
                                "startColumnIndex": 0,
                                "endColumnIndex": WIDTH,
                            }
                        }
                    }
                ]
            },
        ).execute()
    except Exception as e:
        print("unmerge skip:", e)

    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range="'%s'!A1" % SHEET_TITLE,
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    gray = {"red": 0.85, "green": 0.85, "blue": 0.85}
    blue = {"red": 0.79, "green": 0.89, "blue": 0.98}
    green = {"red": 0.82, "green": 0.93, "blue": 0.82}
    yellow = {"red": 1.0, "green": 0.95, "blue": 0.8}
    section = {"red": 1.0, "green": 0.90, "blue": 0.70}
    err_bg = {"red": 0.98, "green": 0.85, "blue": 0.85}
    rule_bg = {"red": 0.93, "green": 0.90, "blue": 0.98}
    common_bg = {"red": 0.90, "green": 0.90, "blue": 0.90}

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

    reqs: List[dict] = [
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
                    "endColumnIndex": MAP_COLS,
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
    # header colors: A-B gray (category+type design), C-E gray keys, F-K+N blue human, L M? 
    # A productCategory human-ish but design chapter -> blue for human filter? User maintains category assignment -> blue
    # B productType design -> gray
    # C D E gray, F G H I J K blue, L green, M blue, N blue, O green
    reqs.extend(color_cells(5, [0, 5, 6, 7, 8, 9, 10, 12, 13], blue))  # human: category + mapping cols
    reqs.extend(color_cells(5, [1, 2, 3, 4], gray))  # type + keys
    reqs.extend(color_cells(5, [11, 14], green))

    for start, end, bg in (
        (0, MAP_COLS, section),
        (ERR_START, ERR_START + 9, section),
        (RULE_START, RULE_START + 2, section),
        (COMMON_START, COMMON_START + 2, section),
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
                            "backgroundColor": bg,
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
                    "startRowIndex": 5,
                    "endRowIndex": 6,
                    "startColumnIndex": ERR_START,
                    "endColumnIndex": ERR_START + 9,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": err_bg,
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
    reqs.append(
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
    # merge mnemonic: A-B いつ効くか, C-E 何を埋めるか, G-H どこから, I-J どう取るか, K-L どこから, M-N メモ
    for start, end in ((0, 2), (2, 5), (6, 8), (8, 10), (10, 12), (12, 14)):
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
    reqs.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 7,
                    "endIndex": 8,
                },
                "properties": {"pixelSize": 78},
                "fields": "pixelSize",
            }
        }
    )

    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": reqs}
    ).execute()

    print(
        "DONE MAP=A-%s ERRORS=%s-%s RULES=%s-%s COMMON=%s-%s"
        % (
            col_letter(MAP_COLS - 1),
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
    for p in (
        SCRIPT_DIR / "_amazon_map_snapshot.json",
        SCRIPT_DIR / "_amazon_map_snapshot2.json",
    ):
        try:
            p.unlink(missing_ok=True)
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
