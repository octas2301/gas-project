# -*- coding: utf-8 -*-
"""Sheets or local CSV store. Never writes to listing Keepa cache."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
LOCAL = ROOT / "local_store"
SECRETS = ROOT.parents[0] / "c1_hpc_packaged" / "secrets"
CONFIG = ROOT.parents[0] / "purchase_research_path3" / "config.local.json"

READ_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]
WRITE_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_RW = SECRETS / "token_sheets_rw.json"
CREDENTIALS = SECRETS / "credentials.json"


def load_config() -> dict:
    if not CONFIG.exists():
        return {}
    return json.loads(CONFIG.read_text(encoding="utf-8"))


def _creds(token_path: Path, scopes: list[str]):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    if not token_path.exists():
        return None
    creds = Credentials.from_authorized_user_file(str(token_path), scopes)
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            return None
    if not creds or not creds.valid:
        return None
    return creds


def ensure_write_creds(interactive: bool = False):
    """Prefer existing Amazon-mapping write token. Optional browser OAuth."""
    creds = _creds(TOKEN_RW, WRITE_SCOPES)
    if creds:
        return creds
    if not interactive:
        return None
    if not CREDENTIALS.is_file():
        return None
    from google_auth_oauthlib.flow import InstalledAppFlow

    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS), WRITE_SCOPES)
    creds = flow.run_local_server(port=0)
    TOKEN_RW.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_RW.write_text(creds.to_json(), encoding="utf-8")
    return creds


def sheets_service(write: bool = False, interactive: bool = False):
    from googleapiclient.discovery import build

    if write:
        creds = ensure_write_creds(interactive=interactive)
    else:
        creds = _creds(SECRETS / "token.json", READ_SCOPES)
        if not creds:
            creds = ensure_write_creds(interactive=False)
    if not creds:
        return None
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def drive_service():
    from googleapiclient.discovery import build

    creds = ensure_write_creds(interactive=False) or _creds(SECRETS / "token.json", READ_SCOPES)
    if not creds:
        return None
    return build("drive", "v3", credentials=creds, cache_discovery=False)
