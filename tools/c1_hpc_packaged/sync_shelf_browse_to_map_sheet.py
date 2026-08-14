#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
shelf_browse_catalog.json → ▼設定(Amazonマッピング) の SHELF ブロックのみ差分更新。

配置（0-based）: RULES=AD(29)／共通認識=AG(32)／SHELF=AK(36)
MAP・ERRORS・RULES・共通認識は消さない。
=== SHELF === は MAP の === と同じくシート行4。
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
SHELF_START = 36  # AK (1-based 37) — 共通認識の右
SHELF_COLS = 10


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
    s = ""
    n = n1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build_shelf_values(catalog: Dict[str, Any]) -> List[List[Any]]:
    """
    行4: ===（MAPと揃える）
    行5: （空）
    行6: 英語項目名
    行7: 日本語訳
    行8: 項目の説明（リストheader固定）
    行9〜: データ
    """
    rows = catalog.get("rows") or []
    en = [
        "templateFile",
        "templateUrl",
        "allowedProductTypes",
        "preferredProductType",
        "browseNodeId",
        "browsePath",
        "fingerprintSha",
        "columnMapPath",
        "extractedAt",
        "sourceSheet",
    ]
    ja = [
        "テンプレファイル名",
        "Drive URL（任意）",
        "選べる商品タイプ",
        "採用する商品タイプ",
        "Browse Node ID",
        "Browse Path（表示名）",
        "指紋sha",
        "C1列マップ",
        "抽出日時",
        "抽出元シート",
    ]
    desc = [
        "06に置いた純正xlsmのファイル名",
        "テンプレのDriveリンク（任意）",
        "そのxlsmの商品タイプ候補（カンマ区切り）",
        "このBrowseで採用するPT（候補のうち1つ）",
        "照合キー。Catalog/マスタのNode IDと突合",
        "データを閲覧するの BrowsePath 全文",
        "B-T0テンプレ指紋（充填検証用）",
        "C1の column_map json パス",
        "SHELF抽出を実行したUTC日時",
        "抽出元シート名（データを閲覧する）",
    ]
    out: List[List[Any]] = [
        [""] * SHELF_COLS,
        [""] * SHELF_COLS,
        [""] * SHELF_COLS,
        ["=== SHELF（テンプレ×Browse網羅） ==="] + [""] * (SHELF_COLS - 1),
        [""] * SHELF_COLS,
        en,
        ja,
        desc,
    ]
    for r in rows:
        allowed = r.get("allowedProductTypes") or []
        allowed_s = ",".join(allowed) if isinstance(allowed, list) else str(allowed)
        out.append(
            [
                r.get("templateFile") or "",
                r.get("templateUrl") or "",
                allowed_s,
                r.get("preferredProductType") or "",
                r.get("browseNodeId") or "",
                r.get("browsePath") or "",
                r.get("fingerprintSha") or "",
                r.get("columnMapPath") or "",
                r.get("extractedAt") or "",
                r.get("sourceSheet") or "",
            ]
        )
    padded = []
    for row in out:
        rr = list(row) + [""] * max(0, SHELF_COLS - len(row))
        padded.append(rr[:SHELF_COLS])
    return padded


def main() -> int:
    cfg = load_cfg()
    sid = cfg["spreadsheet_id"]
    start = int(cfg.get("shelf_sheet_start_col") or (SHELF_START + 1))  # 1-based
    # allow 0-based override via shelf_sheet_start_col_0
    if cfg.get("shelf_sheet_start_col_0") is not None:
        start = int(cfg["shelf_sheet_start_col_0"]) + 1

    catalog_path = SCRIPT_DIR / "shelf_browse_catalog.json"
    if not catalog_path.is_file():
        print("missing catalog — run c1_shelf_browse_extract.py", file=sys.stderr)
        return 2
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    values = build_shelf_values(catalog)

    c0 = col_letter(start)
    c1 = col_letter(start + SHELF_COLS - 1)
    svc = build("sheets", "v4", credentials=get_creds(), cache_discovery=False)
    # sheetId for freeze
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

    clear_range = "'%s'!%s1:%s5000" % (SHEET_TITLE, c0, c1)
    svc.spreadsheets().values().clear(spreadsheetId=sid, range=clear_range).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range="'%s'!%s1" % (SHEET_TITLE, c0),
        valueInputOption="RAW",
        body={"values": values},
    ).execute()
    # リストheader=8行目で固定（データは9行目〜）
    if sheet_id is not None:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": sheet_id,
                                "gridProperties": {"frozenRowCount": 8},
                            },
                            "fields": "gridProperties.frozenRowCount",
                        }
                    }
                ]
            },
        ).execute()
    print(
        json.dumps(
            {
                "ok": True,
                "shelfCol": c0,
                "endCol": c1,
                "rows": len(values),
                "browseRows": len(catalog.get("rows") or []),
                "headerRows": "6=en 7=ja 8=desc frozen",
                "dataStartRow": 9,
                "note": "RULES/COMMON untouched",
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
