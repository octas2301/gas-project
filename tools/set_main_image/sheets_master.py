# -*- coding: utf-8 -*-
"""
商品マスタを Google Sheets から直読する。

認証は C1 と同じ OAuth（tools/c1_hpc_packaged/secrets + config.local.json）。
CSVダウンロード不要。
"""
from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

LOG = logging.getLogger("set_main_image.sheets_master")

C1_DIR = Path(__file__).resolve().parent.parent / "c1_hpc_packaged"
DEFAULT_MASTER_SHEET = "▼商品マスタ(人間作業用)"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def _load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_c1_config(explicit: Optional[Path] = None) -> Path:
    if explicit is not None:
        p = Path(explicit)
        if not p.is_file():
            raise FileNotFoundError(f"config がありません: {p}")
        return p
    for name in ("config.local.json", "config.json"):
        p = C1_DIR / name
        if p.is_file():
            return p
    raise FileNotFoundError(
        f"C1 config がありません（{C1_DIR / 'config.local.json'}）。"
        "C1 と同じ spreadsheet_id / fetch を置いてください。"
    )


def load_sheets_settings(
    config_path: Optional[Path] = None,
    *,
    spreadsheet_id: str = "",
    master_sheet: str = "",
) -> Dict[str, Any]:
    cfg_path = resolve_c1_config(config_path)
    cfg = _load_json(cfg_path)
    base = cfg_path.parent
    fetch = cfg.get("fetch") if isinstance(cfg.get("fetch"), dict) else {}
    sid = (spreadsheet_id or cfg.get("spreadsheet_id") or fetch.get("spreadsheet_id") or "").strip()
    if not sid:
        raise ValueError("spreadsheet_id が config にありません")
    sheet = (
        master_sheet
        or fetch.get("master_sheet_name")
        or DEFAULT_MASTER_SHEET
    ).strip()
    cred = base / str(fetch.get("credentials_path") or "secrets/credentials.json")
    if not cred.is_absolute():
        cred = (base / cred).resolve()
    token = base / str(fetch.get("token_path") or "secrets/token.json")
    if not token.is_absolute():
        token = (base / token).resolve()
    return {
        "config_path": cfg_path,
        "spreadsheet_id": sid,
        "master_sheet": sheet,
        "credentials_path": cred,
        "token_path": token,
    }


def _get_credentials(credentials_path: Path, token_path: Path):
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError as e:
        raise RuntimeError(
            "Google API ライブラリが必要です: pip install google-api-python-client "
            "google-auth-httplib2 google-auth-oauthlib"
        ) from e

    creds = None
    if token_path.is_file():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not credentials_path.is_file():
                raise FileNotFoundError(
                    f"OAuth credentials がありません: {credentials_path}"
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path), SCOPES
            )
            creds = flow.run_local_server(port=0)
        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")
        LOG.info("token を保存しました: %s", token_path)
    return creds


def _cell_to_str(c: Any) -> str:
    if c is True:
        return "TRUE"
    if c is False:
        return "FALSE"
    if c is None:
        return ""
    return str(c)


def fetch_master_rows(
    *,
    config_path: Optional[Path] = None,
    spreadsheet_id: str = "",
    master_sheet: str = "",
    value_render_option: str = "UNFORMATTED_VALUE",
) -> Tuple[List[List[str]], Dict[str, Any]]:
    """
    スプシのマスタシート全行を取得。
    チェックボックスは UNFORMATTED_VALUE → TRUE/FALSE 文字列へ正規化。
    画像URLが IMAGE() 数式のときは value_render_option='FORMULA' も併用可。
    """
    from googleapiclient.discovery import build

    settings = load_sheets_settings(
        config_path,
        spreadsheet_id=spreadsheet_id,
        master_sheet=master_sheet,
    )
    creds = _get_credentials(settings["credentials_path"], settings["token_path"])
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    sid = settings["spreadsheet_id"]
    title = settings["master_sheet"]

    meta = sheets.spreadsheets().get(spreadsheetId=sid).execute()
    titles = [
        (sh.get("properties") or {}).get("title") or ""
        for sh in (meta.get("sheets") or [])
    ]
    if title not in titles:
        raise RuntimeError(
            f"マスタシートが見つかりません: {title!r} / 候補={titles[:12]}"
        )

    rng = "'%s'" % title.replace("'", "''")
    # UNFORMATTED: チェックボックスが bool で返る（レ点の取りこぼし防止）
    data = (
        sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=sid,
            range=rng,
            majorDimension="ROWS",
            valueRenderOption=str(value_render_option or "UNFORMATTED_VALUE"),
        )
        .execute()
    )
    raw: List[List[Any]] = data.get("values") or []
    if not raw:
        raise RuntimeError("マスタシートが空です")

    width = max(len(r) for r in raw)
    rows: List[List[str]] = []
    true_count = 0
    for r in raw:
        padded = list(r) + [""] * (width - len(r))
        cells = [_cell_to_str(c) for c in padded]
        rows.append(cells)
    # 出品CK列の TRUE 件数は呼び出し側で数える

    info = {
        "spreadsheetId": sid,
        "masterSheet": title,
        "rowCount": len(rows),
        "colCount": width,
        "configPath": str(settings["config_path"]),
        "source": "google_sheets",
    }
    LOG.info(
        "sheets master loaded sheet=%r rows=%d cols=%d",
        title,
        len(rows),
        width,
    )
    return rows, info


def count_true_in_column(rows: List[List[str]], col_name: str = "出品CK") -> int:
    from master_sets import _find_header_row, checkbox_is_true

    header_i, idx = _find_header_row(rows)
    if col_name not in idx:
        return -1
    j = idx[col_name]
    n = 0
    for row in rows[header_i + 1 :]:
        if j < len(row) and checkbox_is_true(row[j]):
            n += 1
    return n


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    rs, info = fetch_master_rows()
    print(json.dumps({**info, "trueCk": count_true_in_column(rs)}, ensure_ascii=False, indent=2))
