# -*- coding: utf-8 -*-
"""販促通知メール（Gmail API 優先）。"""
from __future__ import annotations

import base64
import logging
import smtplib
from email.mime.text import MIMEText
from typing import Any, Dict

LOG = logging.getLogger("amazon_deals_bulk.notify_mail")


def send_smtp(cfg: dict, subject: str, body: str) -> bool:
    to_addr = str(cfg.get("notify_email_to") or "").strip()
    host = str(cfg.get("smtp_host") or "").strip()
    if not to_addr or not host:
        return False
    port = int(cfg.get("smtp_port") or 587)
    user = str(cfg.get("smtp_user") or "").strip()
    password = str(cfg.get("smtp_password") or "").strip()
    from_addr = str(cfg.get("notify_email_from") or user or to_addr).strip()
    msg = MIMEText(body, _charset="utf-8")
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = to_addr
    with smtplib.SMTP(host, port, timeout=60) as s:
        s.starttls()
        if user and password:
            s.login(user, password)
        s.sendmail(from_addr, [to_addr], msg.as_string())
    return True


def send_gmail_api(cfg: dict, subject: str, body: str) -> bool:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    from sheets_io import _c1_paths

    to_addr = str(cfg.get("notify_email_to") or "contact@octas2301.com").strip()
    if not to_addr:
        return False
    scopes = ["https://www.googleapis.com/auth/gmail.send"]
    cred_path, _ = _c1_paths()
    token_path = cred_path.parent / "token_gmail_send.json"
    creds = None
    if token_path.is_file():
        creds = Credentials.from_authorized_user_file(str(token_path), scopes)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    if not creds or not creds.valid:
        if not cred_path.is_file():
            LOG.warning("Gmail: credentials.json なし")
            return False
        flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), scopes)
        creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    msg = MIMEText(body, _charset="utf-8")
    msg["To"] = to_addr
    msg["Subject"] = subject
    from_addr = str(cfg.get("notify_email_from") or "").strip()
    if from_addr:
        msg["From"] = from_addr
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
    svc = build("gmail", "v1", credentials=creds, cache_discovery=False)
    svc.users().messages().send(userId="me", body={"raw": raw}).execute()
    LOG.info("Gmail API 送信OK → %s", to_addr)
    return True


def send_notify(cfg: dict, subject: str, body: str) -> bool:
    """SMTP → Gmail API。どちらも失敗なら False。"""
    try:
        if send_smtp(cfg, subject, body):
            return True
    except Exception as e:
        LOG.warning("SMTP失敗: %s", e)
    try:
        return send_gmail_api(cfg, subject, body)
    except Exception as e:
        LOG.warning("Gmail失敗: %s", e)
        return False


def build_custom_clip_alert_body(events: list) -> str:
    lines = [
        "【amazonタイムセール】独自セール日時を公式Sale優先で調整しました",
        "",
        "名付き公式と期間が重なったため、名前なし（カスタム／月）側の開始・終了を短縮または見送りにしました。",
        "すでにSC登録済みの枠は、必要なら Seller Central の画面編集で日付を合わせてください（バルク再ULでは直らない場合があります）。",
        "",
    ]
    for i, ev in enumerate(events, 1):
        lines.append("--- %s ---" % i)
        for k in ("SKU", "ASIN", "スケジュール", "旧開始", "旧終了", "新開始", "新終了", "状態", "理由"):
            if k in ev and ev.get(k) not in (None, ""):
                lines.append("%s: %s" % (k, ev.get(k)))
        lines.append("")
    return "\n".join(lines)
