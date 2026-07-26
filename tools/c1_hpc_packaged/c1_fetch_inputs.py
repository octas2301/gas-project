#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
C1 入力自動取得（案A）: Drive GENERATED + マスタCSV → ローカル input/

要件: docs/org/D_MENU_C1_FETCH_INPUTS_REQUIREMENTS.md
認証: ユーザ OAuth（drive.readonly + spreadsheets.readonly）
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import parse_qs, urlparse

SCRIPT_DIR = Path(__file__).resolve().parent
LOG = logging.getLogger("c1_fetch_inputs")

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]

FILE_ID_RE = re.compile(r"/d/([a-zA-Z0-9_-]{10,})")
OPEN_ID_RE = re.compile(r"[?&]id=([a-zA-Z0-9_-]{10,})")


def _load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def resolve_path(p: str, base: Path) -> Path:
    path = Path(p)
    if not path.is_absolute():
        path = (base / path).resolve()
    return path


def extract_drive_file_id(url_or_id: str) -> str:
    s = (url_or_id or "").strip()
    if not s:
        return ""
    if re.fullmatch(r"[a-zA-Z0-9_-]{25,}", s) and "/" not in s and "http" not in s:
        return s
    m = FILE_ID_RE.search(s)
    if m:
        return m.group(1)
    m = OPEN_ID_RE.search(s)
    if m:
        return m.group(1)
    parsed = urlparse(s)
    qs = parse_qs(parsed.query)
    if "id" in qs and qs["id"]:
        return qs["id"][0]
    return ""


def get_credentials(credentials_path: Path, token_path: Path):
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        LOG.error(
            "Google API ライブラリが必要です: pip install -r requirements.txt"
        )
        sys.exit(2)

    creds = None
    if token_path.is_file():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not credentials_path.is_file():
                LOG.error(
                    "OAuth クライアント JSON がありません: %s\n"
                    "GCP でデスクトップアプリの OAuth クライアントを作成し、"
                    "JSON をこのパスへ置いてください。",
                    credentials_path,
                )
                sys.exit(2)
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path), SCOPES
            )
            creds = flow.run_local_server(port=0)
        token_path.parent.mkdir(parents=True, exist_ok=True)
        with token_path.open("w", encoding="utf-8") as f:
            f.write(creds.to_json())
        LOG.info("token を保存しました: %s", token_path)
    return creds


def build_services(creds):
    from googleapiclient.discovery import build

    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return drive, sheets


def download_drive_file(drive, file_id: str, dest: Path) -> Path:
    from googleapiclient.http import MediaIoBaseDownload
    import io

    dest.parent.mkdir(parents=True, exist_ok=True)
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    dest.write_bytes(buf.getvalue())
    LOG.info("GENERATED 保存: %s (%d bytes) fileId=%s", dest, dest.stat().st_size, file_id)
    return dest


def find_generated_in_folder(drive, folder_id: str, file_name: str) -> Optional[str]:
    folder_id = (folder_id or "").strip()
    if folder_id:
        q = "name = {name} and '{fid}' in parents and trashed = false".format(
            name=json.dumps(file_name), fid=folder_id.replace("'", "\\'")
        )
    else:
        q = "name = {name} and trashed = false".format(name=json.dumps(file_name))
    res = (
        drive.files()
        .list(q=q, spaces="drive", fields="files(id, name)", pageSize=5)
        .execute()
    )
    files = res.get("files") or []
    if not files and folder_id:
        q2 = "name = {name} and trashed = false".format(name=json.dumps(file_name))
        res2 = (
            drive.files()
            .list(q=q2, spaces="drive", fields="files(id, name)", pageSize=5)
            .execute()
        )
        files = res2.get("files") or []
    if not files:
        return None
    if len(files) > 1:
        LOG.warning("同名 GENERATED が複数: %s → 先頭を使用", [f["id"] for f in files])
    return files[0]["id"]


def find_latest_generated_from_log(
    sheets, spreadsheet_id: str, log_sheet_name: str
) -> Tuple[str, str, str]:
    """
    ログシートを末尾から探し、status=GENERATED かつ fileUrl ありの最新を返す。
    @return (sub_batch_id, file_id, file_url)
    """
    meta = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = None
    for sh in meta.get("sheets") or []:
        props = sh.get("properties") or {}
        if props.get("title") == log_sheet_name:
            sheet_id = props.get("sheetId")
            title = props.get("title")
            break
    if sheet_id is None:
        raise RuntimeError("ログシートが見つかりません: %s" % log_sheet_name)

    # ヘッダー込みで広めに取得
    rng = "'%s'!A1:N5000" % log_sheet_name.replace("'", "''")
    data = (
        sheets.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=rng)
        .execute()
    )
    rows = data.get("values") or []
    if len(rows) < 2:
        raise RuntimeError("ログシートが空です")

    header = [str(c).strip() for c in rows[0]]
    # 1行目が recordType でない場合（説明行）を許容
    header_row_idx = 0
    for i, row in enumerate(rows[:5]):
        cells = [str(c).strip() for c in row]
        if "recordType" in cells or (cells and cells[0] == "recordType"):
            header = cells
            header_row_idx = i
            break

    def col(*names: str) -> int:
        for n in names:
            if n in header:
                return header.index(n)
        return -1

    i_type = col("recordType")
    i_status = col("status")
    i_url = col("fileUrl")
    i_sub = col("subBatchId")
    i_name = col("fileName")
    if i_status < 0 or i_url < 0:
        raise RuntimeError(
            "ログヘッダーに status / fileUrl がありません: %s" % header[:20]
        )

    for row in reversed(rows[header_row_idx + 1 :]):
        def cell(idx: int) -> str:
            if idx < 0 or idx >= len(row):
                return ""
            return str(row[idx]).strip()

        if i_type >= 0 and cell(i_type) and cell(i_type) != "RUN":
            continue
        if cell(i_status) != "GENERATED":
            continue
        url = cell(i_url)
        fid = extract_drive_file_id(url)
        if not fid:
            continue
        sub = cell(i_sub)
        if not sub and i_name >= 0:
            fn = cell(i_name)
            if fn.endswith("_GENERATED.csv"):
                sub = fn[: -len("_GENERATED.csv")]
        return sub, fid, url

    raise RuntimeError(
        "ログに fileUrl 付き GENERATED 行がありません（%s）" % log_sheet_name
    )


def export_master_csv(
    sheets, spreadsheet_id: str, sheet_name: str, dest: Path
) -> Path:
    meta = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    found = False
    for sh in meta.get("sheets") or []:
        title = (sh.get("properties") or {}).get("title") or ""
        if title == sheet_name:
            found = True
            break
    if not found:
        raise RuntimeError("マスタシートが見つかりません: %s" % sheet_name)

    rng = "'%s'" % sheet_name.replace("'", "''")
    data = (
        sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=rng,
            majorDimension="ROWS",
            valueRenderOption="FORMATTED_VALUE",
        )
        .execute()
    )
    rows: List[List[Any]] = data.get("values") or []
    if not rows:
        raise RuntimeError("マスタシートが空です")

    # 行ごとに列数を揃える
    width = max(len(r) for r in rows)
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        for r in rows:
            padded = list(r) + [""] * (width - len(r))
            w.writerow(["" if c is None else c for c in padded])
    LOG.info("マスタCSV保存: %s rows=%d cols=%d", dest, len(rows), width)
    return dest


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    ap = argparse.ArgumentParser(description="C1 入力取得（GENERATED + マスタCSV）")
    ap.add_argument("--config", required=True, help="config.local.json 等")
    ap.add_argument("--sub-batch", default="", help="例: A1_…_B2")
    ap.add_argument("--generated-file-id", default="", help="Drive ファイル ID")
    ap.add_argument(
        "--latest",
        action="store_true",
        help="Lv4ログの最新 GENERATED（fileUrl）を使用",
    )
    ap.add_argument("--skip-generated", action="store_true")
    ap.add_argument("--skip-master", action="store_true")
    args = ap.parse_args(argv)

    cfg_path = Path(args.config).resolve()
    if not cfg_path.is_file():
        LOG.error("config がありません: %s", cfg_path)
        return 2
    cfg = _load_json(cfg_path)
    base = cfg_path.parent
    fetch = cfg.get("fetch") or {}
    if not isinstance(fetch, dict) or not fetch:
        LOG.error("config に fetch {} ブロックが必要です（config.example.json 参照）")
        return 2

    spreadsheet_id = str(cfg.get("spreadsheet_id") or fetch.get("spreadsheet_id") or "").strip()
    if not spreadsheet_id and not args.skip_master and not args.latest:
        # generated only by file id / folder なら spreadsheet 不要の場合あり
        pass
    if (not args.skip_master or args.latest) and not spreadsheet_id:
        LOG.error("spreadsheet_id が必要です（マスタ取得または --latest）")
        return 2

    cred_path = resolve_path(
        str(fetch.get("credentials_path") or "secrets/credentials.json"), base
    )
    token_path = resolve_path(
        str(fetch.get("token_path") or "secrets/token.json"), base
    )
    input_dir = resolve_path(str(fetch.get("input_dir") or "input"), base)
    master_sheet = str(
        fetch.get("master_sheet_name") or "▼商品マスタ(人間作業用)"
    ).strip()
    log_sheet = str(
        fetch.get("log_sheet_name") or "▼Lv4実行ログ(Amazon)"
    ).strip()
    folder_id = str(fetch.get("generated_folder_id") or "").strip()

    creds = get_credentials(cred_path, token_path)
    drive, sheets = build_services(creds)

    out_generated = None
    out_master = None
    sub_batch = str(args.sub_batch or "").strip()

    if not args.skip_generated:
        file_id = str(args.generated_file_id or "").strip()
        file_url = ""
        if not file_id and args.latest:
            if not spreadsheet_id:
                LOG.error("--latest には spreadsheet_id が必要です")
                return 2
            sub_batch, file_id, file_url = find_latest_generated_from_log(
                sheets, spreadsheet_id, log_sheet
            )
            LOG.info(
                "ログ最新 GENERATED subBatchId=%s fileId=%s", sub_batch, file_id
            )
        if not file_id and sub_batch:
            name = sub_batch + "_GENERATED.csv"
            file_id = find_generated_in_folder(drive, folder_id, name)
            if not file_id:
                LOG.error(
                    "GENERATED が見つかりません: %s（folder_id=%s）。"
                    "--generated-file-id または --latest を使ってください。",
                    name,
                    folder_id or "(Drive全体名検索)",
                )
                return 1
        if not file_id:
            LOG.error(
                "GENERATED の指定がありません。"
                "--sub-batch / --generated-file-id / --latest のいずれかを指定してください。"
            )
            return 2
        if not sub_batch:
            # メタから名前取得
            meta = drive.files().get(fileId=file_id, fields="name").execute()
            fn = str(meta.get("name") or "")
            if fn.endswith("_GENERATED.csv"):
                sub_batch = fn[: -len("_GENERATED.csv")]
            else:
                sub_batch = file_id[:12]
        dest_name = sub_batch + "_GENERATED.csv"
        # config の generated_csv があればそのパスを優先
        cfg_gen = str(cfg.get("generated_csv") or "").strip()
        if cfg_gen and "{subBatchId}" not in cfg_gen:
            # 固定パスが書いてある場合は input_dir 配下の標準名も併記用に書く
            out_generated = input_dir / dest_name
        else:
            out_generated = input_dir / dest_name
        download_drive_file(drive, file_id, out_generated)
        # config 記載パスへもコピー（相対パス解決）
        if cfg_gen:
            target = resolve_path(
                cfg_gen.replace("{subBatchId}", sub_batch), base
            )
            if target.resolve() != out_generated.resolve():
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(out_generated.read_bytes())
                LOG.info("config.generated_csv へコピー: %s", target)
                out_generated = target

    if not args.skip_master:
        cfg_master = str(cfg.get("master_csv") or "").strip()
        if cfg_master:
            out_master = resolve_path(cfg_master, base)
        else:
            out_master = input_dir / "master_export.csv"
        export_master_csv(sheets, spreadsheet_id, master_sheet, out_master)

    LOG.info("完了")
    if out_generated:
        LOG.info("  generated=%s", out_generated)
    if out_master:
        LOG.info("  master=%s", out_master)
    if sub_batch:
        LOG.info("  subBatchId=%s", sub_batch)
    LOG.info("次: python c1_packaged.py --config %s --mode dry_run", cfg_path.name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
