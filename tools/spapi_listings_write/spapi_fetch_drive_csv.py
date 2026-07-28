#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SP-API v1.3: Drive 上の最新 *_SPAPI_ITEMS.csv をローカルへ取得

- 認証: C1 と同じユーザ OAuth（drive.readonly）
- 既定: フォルダ内で name contains SPAPI_ITEMS / 最新 modifiedTime
- 出力: tools/spapi_listings_write/items.csv（上書き）

正本: docs/org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md
承認: docs/org/LV4_SPAPI_DRIVE_FETCH_APPROVAL.md
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import parse_qs, urlparse

SCRIPT_DIR = Path(__file__).resolve().parent
LOG = logging.getLogger("spapi_fetch_drive_csv")

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
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
        LOG.error("Google API が必要です: pip install -r requirements.txt")
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
                    "GCP デスクトップ OAuth クライアントを作成し配置してください。"
                    "（C1 の secrets/credentials.json 流用可）",
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
        LOG.info("token 保存: %s", token_path)
    return creds


def build_drive(creds):
    from googleapiclient.discovery import build

    return build("drive", "v3", credentials=creds, cache_discovery=False)


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
    LOG.info("保存: %s (%d bytes) fileId=%s", dest, dest.stat().st_size, file_id)
    return dest


def find_latest_spapi_items(
    drive, folder_id: str, name_contains: str = "SPAPI_ITEMS"
) -> Optional[Dict[str, Any]]:
    """フォルダ内（または全体）で名前に name_contains を含む最新 CSV。"""
    name_contains = (name_contains or "SPAPI_ITEMS").strip()
    folder_id = (folder_id or "").strip()
    q_parts = [
        "mimeType != 'application/vnd.google-apps.folder'",
        "trashed = false",
        "name contains %s" % json.dumps(name_contains),
    ]
    if folder_id:
        q_parts.append("'%s' in parents" % folder_id.replace("'", "\\'"))
    q = " and ".join(q_parts)
    res = (
        drive.files()
        .list(
            q=q,
            spaces="drive",
            fields="files(id, name, modifiedTime, webViewLink)",
            orderBy="modifiedTime desc",
            pageSize=10,
        )
        .execute()
    )
    files: List[Dict[str, Any]] = res.get("files") or []
    # csv 優先
    csv_files = [f for f in files if str(f.get("name") or "").lower().endswith(".csv")]
    pick = csv_files[0] if csv_files else (files[0] if files else None)
    if not pick:
        return None
    LOG.info(
        "Drive 最新: name=%s id=%s modified=%s",
        pick.get("name"),
        pick.get("id"),
        pick.get("modifiedTime"),
    )
    return pick


def fetch_to_local(cfg: Dict[str, Any], base: Path) -> Path:
    drive_cfg = cfg.get("drive") or {}
    if not isinstance(drive_cfg, dict):
        drive_cfg = {}

    cred_rel = str(
        drive_cfg.get("credentials_path")
        or "../c1_hpc_packaged/secrets/credentials.json"
    )
    token_rel = str(drive_cfg.get("token_path") or "secrets/token_drive.json")
    folder_id = str(
        drive_cfg.get("folder_id")
        or cfg.get("drive_folder_id")
        or ""
    ).strip()
    file_id = extract_drive_file_id(
        str(drive_cfg.get("file_id") or drive_cfg.get("file_url") or "")
    )
    name_contains = str(drive_cfg.get("name_contains") or "SPAPI_ITEMS").strip()
    dest_rel = str(
        drive_cfg.get("dest_csv") or cfg.get("items_csv") or "items.csv"
    )
    dest = resolve_path(dest_rel, base)
    cred_path = resolve_path(cred_rel, base)
    token_path = resolve_path(token_rel, base)

    creds = get_credentials(cred_path, token_path)
    drive = build_drive(creds)

    if not file_id:
        found = find_latest_spapi_items(drive, folder_id, name_contains)
        if not found:
            raise SystemExit(
                "Drive に SPAPI_ITEMS CSV が見つかりません。"
                "21-⑧ 実行後か folder_id / file_id を確認してください。"
                " folder_id=%r" % (folder_id or "(未設定・全体検索)")
            )
        file_id = str(found["id"])

    download_drive_file(drive, file_id, dest)
    return dest


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Drive から最新 SPAPI_ITEMS.csv を取得（v1.3）"
    )
    parser.add_argument(
        "--config",
        default=str(SCRIPT_DIR / "config.local.json"),
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    config_path = Path(args.config).expanduser().resolve()
    if not config_path.is_file():
        raise SystemExit(
            "config がありません: %s\ncopy config.example.json config.local.json"
            % config_path
        )
    cfg = _load_json(config_path)
    dest = fetch_to_local(cfg, config_path.parent)
    LOG.info("完了: %s", dest)
    return 0


if __name__ == "__main__":
    sys.exit(main())
