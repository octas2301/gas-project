#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RULES／共通認識を復元し、その右に SHELF（Browse網羅）を配置する。
sync_shelf_browse_to_map_sheet が AA 列から消してしまった被害の修復。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_RW = SCRIPT_DIR / "secrets" / "token_sheets_rw.json"
SHEET_TITLE = "▼設定(Amazonマッピング)"

# add_value_source 後の横並び（0-based）
# MAP A–Q(0-16), gap R(17), ERRORS S(18), … RULES AD(29), COMMON AG(32)
RULE_START = 29  # AD
COMMON_START = 32  # AG
SHELF_START = 36  # AK（共通認識の右に1列空け）
# クリア帯: RULES〜SHELF右端
CLEAR_FROM = RULE_START  # AD
CLEAR_TO = SHELF_START + 12  # AV 付近


def load_cfg() -> dict:
    for name in ("config.local.json", "config.json"):
        p = SCRIPT_DIR / name
        if p.is_file():
            return json.loads(p.read_text(encoding="utf-8"))
    return {"spreadsheet_id": "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"}


def get_creds() -> Credentials:
    cred_path = SCRIPT_DIR / "secrets" / "credentials.json"
    creds = None
    if TOKEN_RW.is_file():
        creds = Credentials.from_authorized_user_file(str(TOKEN_RW), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_RW.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_RW.write_text(creds.to_json(), encoding="utf-8")
    return creds


def col_letter(n1: int) -> str:
    """1-based column number → A1 letter."""
    s = ""
    n = n1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def rules_rows() -> List[List[Any]]:
    """RULES データ（見出し行は別途）。日本語正本＋値の取り方。"""
    return [
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
        ["空なら子SKU", "メーカー型番／品番が空ならGENERATED→子SKU"],
        ["ハイライト優先B", "タイトル≤75: 楽天キャッチ→Yahooキャッチ→箇条書き①。超なら空"],
        ["はい／いいえに正規化", "はい／いいえへそろえる"],
        ["固定値を使う", "既定値をそのまま書く"],
        ["親／子供を役割で付ける", "親行=親、子行=子供"],
        ["数字のみ", "価格などから数字だけ抜く"],
        ["その他画像URLを分割（1〜8枚目）", "Amazon PT URLを|分割して各画像列へ"],
        ["マスタのみ", "マスタ列だけから取る（バルク選択肢は見ない）"],
        ["バルク選択肢のみ", "純正xlsmの推奨値／Dropdownから選ぶ（マスタは見ない）"],
        ["マスタ→バルク選択肢", "まずマスタ。空または選択肢外ならバルクのプルダウンから選ぶ"],
        ["バルク選択肢→マスタ", "まずバルク選択肢の既定。無ければマスタ"],
        ["固定", "既定値をそのまま使う"],
        ["GENERATED", "GENERATED CSV（または役割行）から取る"],
    ]


def common_rows() -> List[List[Any]]:
    return [
        ["1", "マスタで子に値がある項目は子を使う。子が空なら親を使う（子優先（空なら親））"],
        ["2", "バリエ数量・重量・サイズは「子のみ」"],
        ["3", "章立て＝商品カテゴリー、xlsm値＝商品タイプを併記"],
        ["4", "値の取り方でマスタ／バルク選択肢／固定／GENERATEDを明示する"],
        ["5", "本シートは試行用。C1は当面jsonが正。SHELF網羅はbrowseNodeIdでテンプレを確定する"],
    ]


def shelf_block(catalog: Dict[str, Any]) -> List[List[Any]]:
    """行6英語・行7日本語・行8説明（header固定）・行9〜データ。呼び出し側で行1–3をパディング。"""
    from sync_shelf_browse_to_map_sheet import build_shelf_values

    # build_shelf_values は行1–3空＋行4=== 込みのフルブロック
    full = build_shelf_values(catalog)
    # restore 側は行1–3を別途足すため、行4以降だけ返す
    return full[3:]


def main() -> int:
    cfg = load_cfg()
    sid = cfg["spreadsheet_id"]
    catalog_path = SCRIPT_DIR / "shelf_browse_catalog.json"
    if not catalog_path.is_file():
        print("missing shelf_browse_catalog.json", file=sys.stderr)
        return 2
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))

    svc = build("sheets", "v4", credentials=get_creds(), cache_discovery=False)

    # Expand grid
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

    shelf = shelf_block(catalog)
    n_rows = max(220, len(shelf) + 20)
    n_cols = max(55, CLEAR_TO + 2)
    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {"rowCount": n_rows, "columnCount": n_cols},
                        },
                        "fields": "gridProperties(rowCount,columnCount)",
                    }
                }
            ]
        },
    ).execute()

    # Clear from RULES through SHELF (do not touch MAP/ERRORS)
    c_from = col_letter(CLEAR_FROM + 1)
    c_to = col_letter(CLEAR_TO + 1)
    clear_range = "'%s'!%s1:%s%d" % (SHEET_TITLE, c_from, c_to, n_rows)
    svc.spreadsheets().values().clear(spreadsheetId=sid, range=clear_range).execute()

    # Also clear leftover SHELF fragments left of RULES (AA–AC) that sync polluted
    # AA=27, AB=28, AC=29 → indices 26-28; RULE_START=29 so clear 26-28
    pollute_from = 26  # AA
    if pollute_from < RULE_START:
        pf = col_letter(pollute_from + 1)
        pt = col_letter(RULE_START)  # up to AC (before AD)
        svc.spreadsheets().values().clear(
            spreadsheetId=sid,
            range="'%s'!%s1:%s%d" % (SHEET_TITLE, pf, pt, n_rows),
        ).execute()

    rules = rules_rows()
    commons = common_rows()

    # Build wide rows aligned with MAP header rows 1–8 (1-based)
    # Row4 === markers, row5 mnemonic, row6 headers, row7 meaning, row8 purpose, row9+ data
    # Match add_value_source: r1 title, r2 usage, r3 colors, r4 ===, r5 labels, r6 headers, r7 meaning, r8 purpose, r9+ data
    max_len = max(len(rules), len(commons), len(shelf) - 6)  # shelf has 6 header rows then data
    width = CLEAR_TO + 1
    values: List[List[Any]] = []

    def blank() -> List[Any]:
        return [""] * width

    def put(row: List[Any], start: int, cells: List[Any]) -> List[Any]:
        out = list(row) + [""] * max(0, width - len(row))
        out = out[:width]
        for i, v in enumerate(cells):
            if start + i < width:
                out[start + i] = v
        return out

    # We only write columns from RULE_START — but update API needs full rows if we use A1.
    # Write three separate ranges instead (safer for MAP/ERRORS).

    # --- RULES block (rows 1-8 header structure + data from row 9) ---
    rules_values: List[List[Any]] = [
        ["", ""],  # r1 — layout note filled below in batch with row1 note
        ["", ""],
        ["", ""],
        ["=== RULES（継承・変換・値の取り方） ===", ""],
        ["ルール名", "意味"],
        ["ルール名", "意味"],
        ["継承・変換・値の取り方", "人間向け説明"],
        ["日本語で統一", "MAPの加工方法・値の取り方と対応"],
    ]
    # Align with MAP: row1-3 are sheet intro (leave empty in RULES cols except row1 note)
    rules_values[0] = ["（RULES）", "MAPと同じ行位置。右が共通認識→SHELF"]
    for a, b in rules:
        rules_values.append([a, b])

    common_values: List[List[Any]] = [
        ["", ""],
        ["", ""],
        ["", ""],
        ["=== 共通認識（触らない正本） ===", ""],
        ["#", "内容"],
        ["番号", "内容"],
        ["番号", "運用の約束"],
        ["", "最右付近。原則変えない。SHELFはその右"],
    ]
    for a, b in commons:
        common_values.append([a, b])

    # SHELF already has 6 meta/header rows; prepend 2 blank to align === with MAP row4
    # MAP: row1 title, row2 usage, row3 color, row4 ===
    # shelf_block starts with === at [0] — pad 3 rows so === lands on sheet row 4
    shelf_values = [["", ""], ["", ""], ["", ""]] + shelf

    # Pad lengths
    def pad_block(block: List[List[Any]], cols: int) -> List[List[Any]]:
        out = []
        for r in block:
            rr = list(r) + [""] * max(0, cols - len(r))
            out.append(rr[:cols])
        return out

    rules_values = pad_block(rules_values, 2)
    common_values = pad_block(common_values, 2)
    shelf_values = pad_block(shelf_values, 10)

    data = [
        {
            "range": "'%s'!%s1" % (SHEET_TITLE, col_letter(RULE_START + 1)),
            "values": rules_values,
        },
        {
            "range": "'%s'!%s1" % (SHEET_TITLE, col_letter(COMMON_START + 1)),
            "values": common_values,
        },
        {
            "range": "'%s'!%s1" % (SHEET_TITLE, col_letter(SHELF_START + 1)),
            "values": shelf_values,
        },
    ]
    # Update row1 layout hint on ERRORS area if present — skip to avoid overwrite

    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=sid,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()

    # Patch sheet row1 layout string in a safe cell (ERRORS header area often col S=19)
    try:
        svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range="'%s'!S1" % SHEET_TITLE,
            valueInputOption="RAW",
            body={
                "values": [
                    ["左→右: MAP → ERRORS → RULES → 共通認識 → SHELF（Browse網羅）"]
                ]
            },
        ).execute()
    except Exception as e:
        print("warn S1:", e, file=sys.stderr)

    print(
        json.dumps(
            {
                "ok": True,
                "rulesCol": col_letter(RULE_START + 1),
                "commonCol": col_letter(COMMON_START + 1),
                "shelfCol": col_letter(SHELF_START + 1),
                "shelfDataRows": len(catalog.get("rows") or []),
                "rulesRows": len(rules),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
