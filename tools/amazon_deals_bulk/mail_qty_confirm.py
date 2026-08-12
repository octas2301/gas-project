# -*- coding: utf-8 -*-
"""
§9.7 公式Sale数量確認メール下書き／送信。

- 開始 T-21（第1報）／T-14（最終）
- 改定は SC画面編集（バルク再UL不可）と明記
- リンクは広告スプシの「タイムセール」行（スマホはSheetsアプリで開く）
  ※ SCモバイルアプリではタイムセール確認不可のためスプシへ誘導

例:
  python mail_qty_confirm.py
  python mail_qty_confirm.py --days 21
  python mail_qty_confirm.py --days 14 --send
"""
from __future__ import annotations

import argparse
import json
import logging
import smtplib
import sys
from datetime import date, datetime, timezone
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from paths import load_config  # noqa: E402
from qty_logic import compute_q_deal, deal_day_count  # noqa: E402
from schedule_class import format_ymd, parse_ymd  # noqa: E402
from sheet_schema import LANE_B, MASTER_SHEET, SALE_SHEET  # noqa: E402
from sheets_io import read_sheet_rows, sheet_id_by_title, sheets_service  # noqa: E402
from v30_source import resolve_v30_map  # noqa: E402

LOG = logging.getLogger("amazon_deals_bulk.mail_qty_confirm")


def _truthy(v: Any) -> bool:
    return str(v or "").strip().upper() in ("TRUE", "はい", "YES", "Y", "1", "○")


def _float(v: Any) -> Optional[float]:
    if v is None or str(v).strip() == "":
        return None
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None


def spreadsheet_row_url(ssid: str, gid: int, row_1based: int) -> str:
    """スマホのGoogleスプレッドシートアプリで開ける行リンク。"""
    # range は A{row} で当該行付近へ
    return (
        "https://docs.google.com/spreadsheets/d/%s/edit?usp=drivesdk#gid=%s&range=A%s"
        % (ssid, gid, row_1based)
    )


def sheet_gid(svc, ssid: str, title: str) -> int:
    sid = sheet_id_by_title(svc, ssid, title)
    return int(sid) if sid is not None else 0


def select_b_at_t_minus(
    sales: List[Dict[str, Any]],
    *,
    today: date,
    days: int,
    tol: int = 0,
) -> List[Dict[str, Any]]:
    """
    開始まで残り days±tol 日のレーンB。
    登録済（UL済等）は提出対象=いいえでも対象（数量確認のため）。
    """
    out: List[Dict[str, Any]] = []
    tol = max(0, int(tol))
    keep_states = {
        "",
        "予定",
        "要確認",
        "数量改定済",
        "UL済",
        "アップロード済",
        "実施中",
    }
    for r in sales:
        if str(r.get("レーン") or "").strip() != LANE_B:
            continue
        st = str(r.get("状態") or "").strip()
        if st in ("見送り", "終了", "失敗", "停止", "延期"):
            continue
        if st and st not in keep_states:
            continue
        start = parse_ymd(r.get("開始日"))
        if not start or start < today:
            continue
        rem = (start - today).days
        if abs(rem - int(days)) <= tol:
            out.append(r)
    return out


def build_body(
    *,
    cfg: dict,
    rows: List[Dict[str, Any]],
    row_numbers: Dict[str, int],
    today: date,
    days: int,
    ssid: str,
    gid: int,
    revise_map: Dict[str, Dict[str, Any]],
) -> str:
    label = "第1報（T-21）" if days == 21 else ("最終確認（T-14）" if days == 14 else "T-%s" % days)
    lines = [
        "【タイムセール】数量確認メール %s" % label,
        "今日: %s" % today.isoformat(),
        "",
        "目的: 事前登録済み公式Saleの数量を、開始前に見直すか確認する。",
        "重要: 数量改定は Seller Central の Deal 編集画面で行う（バルク再ULでは直らない）。",
        "確認・記録は下のスプレッドシート（スマホはSheetsアプリで開く）。",
        "※ SCモバイルアプリではタイムセール詳細が確認できないため、スプシを正とする。",
        "",
        "選択肢: 改定する／このまま／見送り",
        "",
    ]
    for r in rows:
        sku = str(r.get("SKU") or "").strip()
        asin = str(r.get("ASIN") or "").strip()
        key = sku or asin
        row_n = row_numbers.get(key) or row_numbers.get(asin) or 2
        link = spreadsheet_row_url(ssid, gid, row_n)
        rev = revise_map.get(asin) or revise_map.get(sku) or {}
        lines.extend(
            [
                "----",
                "商品: %s" % (r.get("商品名") or ""),
                "ASIN: %s  SKU: %s" % (asin, sku),
                "スケジュール: %s" % (r.get("スケジュール") or ""),
                "期間: %s .. %s"
                % (format_ymd(r.get("開始日")) or r.get("開始日"), format_ymd(r.get("終了日")) or r.get("終了日")),
                "現状 販売商品数_確定: %s" % (r.get("販売商品数_確定") or ""),
                "再計算案 Q_deal: %s  (%s)" % (rev.get("q_deal", ""), rev.get("note", "")),
                "V30: %s  Q_fba: %s" % (rev.get("v30", ""), rev.get("q_fba", "")),
                "",
                "▶ スプシで開く（スマホ可）:",
                link,
                "",
                "改定する場合の手順:",
                "1. 上のリンクで施策行を確認",
                "2. PCのSCで当該タイムセールを開き数量を修正→「編集内容を送信」",
                "3. 画面で反映を確認し、スプシのメモ／状態を更新",
                "",
            ]
        )
    lines.append("（生成: mail_qty_confirm.py §9.7）")
    return "\n".join(lines)


def write_draft_sheet(
    cfg: dict, *, to_addr: str, subject: str, body: str, web_token: str = ""
) -> str:
    """広告スプシへ下書きを書き、GAS MailApp 送信の入力にする。web_token を返す。"""
    import secrets

    from sheets_io import ensure_sheet, sheets_service

    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    if not sid:
        raise RuntimeError("ads_spreadsheet_id がありません")
    title = "⏱数量確認メール下書き"
    token = (web_token or secrets.token_urlsafe(24)).strip()
    svc = sheets_service(write=True)
    ensure_sheet(svc, sid, title)
    values = [
        ["to", to_addr],
        ["subject", subject],
        ["body", body],
        ["sent_at", ""],
        ["web_token", token],
    ]
    q = title.replace("'", "''")
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range=f"'{q}'!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()
    LOG.info("下書きシート更新: %s (to=%s)", title, to_addr)
    return token


def try_webapp_send_(cfg: dict, *, token: str) -> bool:
    """GAS WebApp doGet で MailApp 送信。"""
    import urllib.error
    import urllib.parse
    import urllib.request

    base = str(cfg.get("qty_mail_web_url") or "").strip()
    ssid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    if not base or not token:
        return False
    q = urllib.parse.urlencode({"ssid": ssid, "token": token})
    url = base.split("?")[0] + "?" + q
    try:
        req = urllib.request.Request(url, method="GET")
        with urllib.request.urlopen(req, timeout=90) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
        LOG.info("WebApp 応答: %s", raw[:300])
        data = json.loads(raw) if raw.strip().startswith("{") else {}
        return bool(data.get("ok"))
    except Exception as e:
        LOG.warning("WebApp 送信失敗: %s", e)
        return False


def try_clasp_send_gas_() -> bool:
    """広告運用GASで sendTimeSaleQtyConfirmMail を実行。"""
    import subprocess

    gas_dir = HERE.parent.parent / "広告運用GAS"
    if not gas_dir.is_dir():
        LOG.warning("広告運用GAS フォルダなし: %s", gas_dir)
        return False
    cmd = ["npx", "--yes", "@google/clasp", "run", "sendTimeSaleQtyConfirmMail"]
    try:
        r = subprocess.run(
            cmd,
            cwd=str(gas_dir),
            capture_output=True,
            text=True,
            timeout=120,
            shell=True,
        )
        out = ((r.stdout or "") + "\n" + (r.stderr or "")).strip()
        LOG.info("clasp run rc=%s out: %s", r.returncode, out[:500])
        if r.returncode != 0:
            return False
        low = out.lower()
        if "not found" in low or "error" in low or "exception" in low:
            return False
        return True
    except Exception as e:
        LOG.warning("clasp run 失敗: %s", e)
        return False


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
    """
    Gmail API で送信（初回のみブラウザ同意 → secrets/token_gmail_send.json）。
    From は認可アカウント。To は notify_email_to。
    """
    import base64

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
            LOG.warning("Gmail: credentials.json なし (%s)", cred_path)
            return False
        flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), scopes)
        creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
        LOG.info("Gmail token 保存: %s", token_path.name)

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
    try:
        # 監査用: 下書きシート sent_at を更新
        from sheets_io import sheets_service as _ss

        sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
        title = "⏱数量確認メール下書き"
        q = title.replace("'", "''")
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " GmailAPI → " + to_addr
        _ss(write=True).spreadsheets().values().update(
            spreadsheetId=sid,
            range=f"'{q}'!B4",
            valueInputOption="RAW",
            body={"values": [[stamp]]},
        ).execute()
    except Exception as e:
        LOG.warning("sent_at 更新失敗（送信自体は成功）: %s", e)
    return True


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="§9.7 数量確認メール")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--days", type=int, default=None, help="21 or 14。省略時は両方（該当があれば）")
    ap.add_argument(
        "--tol",
        type=int,
        default=0,
        help="残り日数の許容±日（例: --days 14 --tol 2 で 12〜16日前も拾う）",
    )
    ap.add_argument("--today", type=str, default=None)
    ap.add_argument(
        "--send",
        action="store_true",
        help="送信: SMTP設定があればSMTP。無ければ下書きシート→GAS MailApp（clasp run）",
    )
    ap.add_argument(
        "--send-gas",
        action="store_true",
        help="SMTPを使わず下書きシート＋GAS MailAppのみ",
    )
    args = ap.parse_args(argv)

    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))
    today = parse_ymd(args.today) if args.today else date.today()
    assert today

    svc = sheets_service(write=False)
    ssid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    gid = sheet_gid(svc, ssid, SALE_SHEET)
    _h, sales = read_sheet_rows(svc, ssid, SALE_SHEET)
    _mh, master = read_sheet_rows(svc, ssid, MASTER_SHEET)
    asins = []
    for r in sales:
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            asins.append(a)
    v30_map = resolve_v30_map(asins, master_rows=master, use_spapi=False)
    qfba = {}
    for m in master:
        a = str(m.get("ASIN") or "").strip().upper()
        if a:
            qfba[a] = _float(m.get("Q_fba"))

    # 行番号（ヘッダ=1）
    row_numbers: Dict[str, int] = {}
    for i, r in enumerate(sales):
        n = i + 2
        sku = str(r.get("SKU") or "").strip()
        asin = str(r.get("ASIN") or "").strip()
        if sku:
            row_numbers[sku] = n
        if asin:
            row_numbers[asin] = n

    day_list = [args.days] if args.days else [21, 14]
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    written = []

    for days in day_list:
        if days is None:
            continue
        rows = select_b_at_t_minus(
            sales, today=today, days=int(days), tol=int(args.tol or 0)
        )
        if not rows:
            LOG.info("T-%s 該当なし", days)
            continue
        revise_map: Dict[str, Dict[str, Any]] = {}
        for r in rows:
            asin = str(r.get("ASIN") or "").strip().upper()
            start = parse_ymd(r.get("開始日"))
            end = parse_ymd(r.get("終了日"))
            dcount = deal_day_count(start, end) if start and end else 0
            qd = compute_q_deal(
                v30=v30_map.get(asin),
                q_fba=qfba.get(asin),
                d_days=dcount,
                schedule=str(r.get("スケジュール") or ""),
            )
            revise_map[asin] = {
                "q_deal": qd.get("Q_deal"),
                "note": qd.get("note"),
                "v30": qd.get("V30"),
                "q_fba": qfba.get(asin),
            }
        body = build_body(
            cfg=cfg,
            rows=rows,
            row_numbers=row_numbers,
            today=today,
            days=int(days),
            ssid=ssid,
            gid=gid,
            revise_map=revise_map,
        )
        subject = "[amazonタイムセール] 数量確認 （スプシで確認・改定はSC画面）"
        path = out_dir / ("qty_confirm_T%s_%s.txt" % (days, stamp))
        path.write_text(body, encoding="utf-8")
        written.append(str(path))
        LOG.info("下書き %s (%s件)", path, len(rows))
        do_send = bool(args.send or args.send_gas)
        if do_send:
            to_addr = str(cfg.get("notify_email_to") or "contact@octas2301.com").strip()
            host = str(cfg.get("smtp_host") or "").strip()
            sent = False
            web_token = ""
            try:
                web_token = write_draft_sheet(
                    cfg, to_addr=to_addr, subject=subject, body=body
                )
            except Exception as e:
                LOG.error("下書きシート更新失敗: %s", e)
            if args.send and host and not args.send_gas:
                try:
                    sent = send_smtp(cfg, subject, body)
                    if sent:
                        LOG.info("SMTP送信OK → %s", to_addr)
                except Exception as e:
                    LOG.warning("SMTP失敗: %s", e)
            if not sent and not args.send_gas:
                try:
                    sent = send_gmail_api(cfg, subject, body)
                except Exception as e:
                    LOG.warning("Gmail API 失敗: %s", e)
            if not sent and web_token:
                if try_webapp_send_(cfg, token=web_token):
                    LOG.info("GAS WebApp 送信OK → %s", to_addr)
                    sent = True
            if not sent:
                if try_clasp_send_gas_():
                    LOG.info("GAS clasp run 送信OK → %s", to_addr)
                    sent = True
                else:
                    LOG.warning(
                        "自動送信未完了。スプシメニュー「⏱ 数量確認メール送信（下書きシート）」で送信可（To=%s）",
                        to_addr,
                    )

    print(json.dumps({"today": today.isoformat(), "drafts": written}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
