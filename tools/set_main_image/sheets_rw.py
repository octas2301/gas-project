# -*- coding: utf-8 -*-
"""
Google Sheets 読み書き（C1 の token_sheets_rw.json を利用）。

サブ画像キュレーションシート作成用。マスタ本体は書かない。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from sheets_master import load_sheets_settings, resolve_c1_config

LOG = logging.getLogger("set_main_image.sheets_rw")

C1_DIR = Path(__file__).resolve().parent.parent / "c1_hpc_packaged"
SCOPES_RW = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_RW = C1_DIR / "secrets" / "token_sheets_rw.json"


def get_rw_credentials():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    settings = load_sheets_settings()
    cred_path = settings["credentials_path"]
    creds = None
    if TOKEN_RW.is_file():
        creds = Credentials.from_authorized_user_file(str(TOKEN_RW), SCOPES_RW)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not Path(cred_path).is_file():
                raise FileNotFoundError(f"OAuth credentials がありません: {cred_path}")
            flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), SCOPES_RW)
            creds = flow.run_local_server(port=0)
        TOKEN_RW.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_RW.write_text(creds.to_json(), encoding="utf-8")
        LOG.info("RW token saved: %s", TOKEN_RW)
    return creds


def build_sheets_rw():
    from googleapiclient.discovery import build

    return build("sheets", "v4", credentials=get_rw_credentials(), cache_discovery=False)


def spreadsheet_id(explicit: str = "") -> str:
    if explicit.strip():
        return explicit.strip()
    return load_sheets_settings()["spreadsheet_id"]


def list_sheet_titles(svc, sid: str) -> List[str]:
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(title,sheetId))").execute()
    return [(sh.get("properties") or {}).get("title") or "" for sh in (meta.get("sheets") or [])]


def ensure_sheet(svc, sid: str, title: str) -> int:
    """シートが無ければ作成。sheetId を返す。"""
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(title,sheetId))").execute()
    for sh in meta.get("sheets") or []:
        props = sh.get("properties") or {}
        if props.get("title") == title:
            return int(props.get("sheetId"))
    req = {"requests": [{"addSheet": {"properties": {"title": title}}}]}
    resp = svc.spreadsheets().batchUpdate(spreadsheetId=sid, body=req).execute()
    replies = resp.get("replies") or []
    props = ((replies[0] or {}).get("addSheet") or {}).get("properties") or {}
    return int(props.get("sheetId") or 0)


def read_sheet_values(svc, sid: str, title: str) -> List[List[str]]:
    rng = "'%s'" % title.replace("'", "''")
    data = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=sid,
            range=rng,
            majorDimension="ROWS",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
    )
    raw = data.get("values") or []
    if not raw:
        return []
    width = max(len(r) for r in raw)
    out: List[List[str]] = []
    for r in raw:
        padded = list(r) + [""] * (width - len(r))
        out.append(["" if c is None else str(c) for c in padded])
    return out


def write_sheet_values(svc, sid: str, title: str, rows: Sequence[Sequence[Any]]) -> None:
    ensure_sheet(svc, sid, title)
    rng = "'%s'" % title.replace("'", "''")
    svc.spreadsheets().values().clear(spreadsheetId=sid, range=rng).execute()
    if not rows:
        return
    # normalize width
    width = max(len(r) for r in rows)
    body_rows = []
    for r in rows:
        body_rows.append(list(r) + [""] * (width - len(r)))
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": body_rows},
    ).execute()


def delete_sheet_by_title(svc, sid: str, title: str) -> bool:
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(title,sheetId))").execute()
    for sh in meta.get("sheets") or []:
        props = sh.get("properties") or {}
        if props.get("title") == title:
            sid_sheet = int(props.get("sheetId"))
            svc.spreadsheets().batchUpdate(
                spreadsheetId=sid,
                body={"requests": [{"deleteSheet": {"sheetId": sid_sheet}}]},
            ).execute()
            LOG.info("deleted sheet %r", title)
            return True
    return False


def apply_checkbox_column(
    svc,
    sid: str,
    sheet_id: int,
    *,
    start_row_0: int,
    end_row_0: int,
    col_0: int,
) -> None:
    """0-based exclusive end rows/cols. BOOLEAN checkbox."""
    if end_row_0 <= start_row_0:
        return
    req = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row_0,
                        "endRowIndex": end_row_0,
                        "startColumnIndex": col_0,
                        "endColumnIndex": col_0 + 1,
                    },
                    "cell": {
                        "dataValidation": {
                            "condition": {"type": "BOOLEAN"},
                            "showCustomUi": True,
                        }
                    },
                    "fields": "dataValidation",
                }
            }
        ]
    }
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body=req).execute()


def freeze_header_and_autosize(svc, sid: str, sheet_id: int, n_cols: int) -> None:
    reqs = [
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
        {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": max(1, n_cols),
                }
            }
        },
    ]
    try:
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": reqs}).execute()
    except Exception as e:
        LOG.warning("freeze/autosize skipped: %s", e)

