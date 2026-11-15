# -*- coding: utf-8 -*-
"""Sheets 読取専用。common_ss_view / competitor_store の認証を流用。書込はしない。"""
from __future__ import annotations

import sys
from pathlib import Path

TOOLS = Path(__file__).resolve().parents[1]
INV_SS_ID = "1_kqBcQTcL0Q5--KsAdxMYbPMsaToLEXxOMuDw_z-kwE"
LISTING_ALIAS = "自動出品260221_セラセン出力"


def sheets_service_readonly():
    """write=False の Sheets API クライアント。失敗時は None。"""
    # competitor_store.client はリポジトリに存在する
    sys.path.insert(0, str(TOOLS / "competitor_store"))
    sys.path.insert(0, str(TOOLS / "common_ss_view"))
    try:
        from client import sheets_service  # type: ignore

        return sheets_service(write=False, interactive=False)
    except Exception:
        return None


def listing_spreadsheet_id() -> str:
    sys.path.insert(0, str(TOOLS / "common_ss_view"))
    from registry import file_id  # type: ignore

    return file_id(LISTING_ALIAS)


def inventory_spreadsheet_id() -> str:
    sys.path.insert(0, str(TOOLS / "common_ss_view"))
    try:
        from registry import file_id  # type: ignore

        return file_id("在庫管理アプリ_Octas")
    except Exception:
        return INV_SS_ID


def read_sheet_values(svc, spreadsheet_id: str, title: str) -> list[list]:
    rng = "'%s'" % title.replace("'", "''")
    # SERIAL_NUMBER + UNFORMATTED で日時を数値／生値に近づける
    resp = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=rng,
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="SERIAL_NUMBER",
        )
        .execute()
    )
    return resp.get("values") or []


def list_sheet_titles(svc, spreadsheet_id: str) -> list[str]:
    meta = (
        svc.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties.title")
        .execute()
    )
    return [s["properties"]["title"] for s in meta.get("sheets") or []]
