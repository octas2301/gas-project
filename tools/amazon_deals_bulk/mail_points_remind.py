# -*- coding: utf-8 -*-
"""
§10.10 / P0-G2・G3 ポイントリマインドメール（apply／restore）。

- apply: 開始の T-N 日前（既定 N=1）に「期間中%を送る」催促
- restore: 終了の N 日後（既定 N=1）に「減衰中%へ戻す」催促
- 下書き: _work/ + スプシ「⏱ポイントリマインド下書き」
- --send: Gmail API／SMTP（数量確認と同系統）

例:
  python mail_points_remind.py --kind apply --days 1
  python mail_points_remind.py --kind restore --days 1
  python mail_points_remind.py --kind both --today 2026-08-27 --tol 1
  python mail_points_remind.py --kind apply --days 1 --send
"""
from __future__ import annotations

import argparse
import json
import logging
import secrets
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from notify_mail import send_notify  # noqa: E402
from paths import load_config  # noqa: E402
from points_logic import (  # noqa: E402
    POINT_STATUS_APPLIED,
    POINT_STATUS_RESTORED,
    before_percent,
    current_percent,
    needs_sync,
    normalize_point_status,
    period_percent,
    restore_percent,
    MODE_RESTORE,
)
from schedule_class import format_ymd, parse_ymd  # noqa: E402
from sheet_schema import (  # noqa: E402
    LANE_B,
    MASTER_SHEET,
    POINT_STATUS_COL,
    SALE_SHEET,
)
from sheets_io import (  # noqa: E402
    ensure_sheet,
    read_sheet_rows,
    sheet_id_by_title,
    sheets_service,
)

LOG = logging.getLogger("amazon_deals_bulk.mail_points_remind")

DRAFT_SHEET = "⏱ポイントリマインド下書き"
SUBJECT_APPLY = "[amazonタイムセール] ポイント apply 確認 （期間中%送信）"
SUBJECT_RESTORE = "[amazonタイムセール] ポイント restore 確認 （減衰中%へ戻す）"

KEEP_STATES = {
    "",
    "予定",
    "要確認",
    "数量改定済",
    "UL済",
    "アップロード済",
    "実施中",
}


def sheet_gid(svc, ssid: str, title: str) -> int:
    sid = sheet_id_by_title(svc, ssid, title)
    return int(sid) if sid is not None else 0


def spreadsheet_row_url(ssid: str, gid: int, row_1based: int) -> str:
    return (
        "https://docs.google.com/spreadsheets/d/%s/edit?usp=drivesdk#gid=%s&range=A%s"
        % (ssid, gid, row_1based)
    )


def master_by_sku(master: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for r in master:
        sku = str(r.get("SKU") or "").strip()
        if sku:
            out[sku] = r
    return out


def select_b_rows(
    sales: List[Dict[str, Any]],
    *,
    today: date,
    kind: str,
    days: int,
    tol: int = 0,
) -> List[Tuple[Dict[str, Any], int]]:
    """
    kind=apply: 開始まで残り days±tol
    kind=restore: 終了から経過 days±tol（終了当日は days=0）
    戻り: (施策行, 行番号1-based)
    """
    out: List[Tuple[Dict[str, Any], int]] = []
    tol = max(0, int(tol))
    for i, r in enumerate(sales):
        if str(r.get("レーン") or "").strip() != LANE_B:
            continue
        st = str(r.get("状態") or "").strip()
        if st in ("見送り", "終了", "失敗", "停止", "延期"):
            continue
        if st and st not in KEEP_STATES:
            continue
        row_n = i + 2
        if kind == "apply":
            start = parse_ymd(r.get("開始日"))
            if not start or start < today:
                continue
            rem = (start - today).days
            if abs(rem - int(days)) <= tol:
                out.append((r, row_n))
        elif kind == "restore":
            end = parse_ymd(r.get("終了日"))
            if not end or end > today:
                continue
            after = (today - end).days
            if abs(after - int(days)) <= tol:
                out.append((r, row_n))
    return out


def needs_apply_remind(mrow: Optional[Dict[str, Any]]) -> bool:
    """
    日程ヒット時の apply 催促。
    期間中適用済／セール前復元済は通常スキップ。
    現在%が既に期間中%でも未適用なら催促する（確認・差分ゼロでも送る）。
    """
    if not mrow:
        return True
    st = normalize_point_status(mrow.get(POINT_STATUS_COL))
    if st == POINT_STATUS_APPLIED:
        return False
    if st == POINT_STATUS_RESTORED:
        return False
    return True


def needs_restore_remind(mrow: Optional[Dict[str, Any]]) -> bool:
    """日程ヒット時の restore 催促。期間中適用済なら減衰中%空でも催促。"""
    if not mrow:
        return False
    st = normalize_point_status(mrow.get(POINT_STATUS_COL))
    if st == POINT_STATUS_RESTORED:
        return False
    if st == POINT_STATUS_APPLIED:
        return True
    return needs_sync(mrow, MODE_RESTORE)


def build_body(
    *,
    kind: str,
    rows: List[Tuple[Dict[str, Any], int, Dict[str, Any]]],
    today: date,
    days: int,
    ssid: str,
    gid_sale: int,
    gid_master: int,
) -> str:
    if kind == "apply":
        title = "【タイムセール】ポイント apply リマインド（開始 T-%s）" % days
        action = (
            "やること: 期間中ポイント% を Pointsフィードで送信（既定1%）\n"
            "  cd tools/amazon_deals_bulk\n"
            "  python points_send.py --mode apply --sku \"…\"   # dry_run\n"
            "  python points_send.py --mode apply --sku \"…\" --prod --i-confirm-prod --wait --update-sheet\n"
            "事前に出品者現在%を見るなら: python points_fetch.py --sku \"…\" --write\n"
            "（最終終着%は fetch しない。空なら減衰フロア1%）"
        )
    else:
        title = "【タイムセール】ポイント restore リマインド（終了＋%s日）" % days
        action = (
            "やること: 減衰中ポイント%（カレンダー位置）へ戻す\n"
            "  cd tools/amazon_deals_bulk\n"
            "  python taper_send.py --poll --mail   # 先にカレンダー同期\n"
            "  python points_send.py --mode restore --sku \"…\"\n"
            "  python points_send.py --mode restore --sku \"…\" --prod --i-confirm-prod --wait --update-sheet\n"
            "※ B終了後は店頭1%→減衰中%。最終終着%へ一気に戻さない"
        )
    lines = [
        title,
        "今日: %s" % today.isoformat(),
        "",
        action,
        "",
        "--- 対象 ---",
    ]
    for sale, row_n, mrow in rows:
        sku = str(sale.get("SKU") or "").strip()
        asin = str(sale.get("ASIN") or "").strip()
        sch = str(sale.get("スケジュール") or "").strip()
        start = format_ymd(sale.get("開始日")) or ""
        end = format_ymd(sale.get("終了日")) or ""
        st_sale = str(sale.get("状態") or "").strip()
        pst = normalize_point_status(mrow.get(POINT_STATUS_COL) if mrow else "")
        cur = current_percent(mrow) if mrow else None
        period = period_percent(mrow) if mrow else 1
        before = before_percent(mrow) if mrow else None
        try:
            restore_pct = restore_percent(mrow) if mrow else None
        except Exception:
            restore_pct = None
        lines.append(
            "- SKU=%s ASIN=%s | %s | %s〜%s | 施策状態=%s"
            % (sku, asin, sch, start, end, st_sale or "-")
        )
        lines.append(
            "  ポイント状態=%s 現在%%=%s 期間中%%=%s 減衰中(restore)%%=%s 終着%%=%s"
            % (
                pst,
                cur if cur is not None else "-",
                period,
                restore_pct if restore_pct is not None else "-",
                before if before is not None else "-",
            )
        )
        if kind == "apply" and cur is not None and cur == period:
            lines.append("  注: 現在%＝期間中%（差分なし。SC確認のみ／送信スキップ可）")
        if kind == "restore" and restore_pct is None:
            lines.append("  警告: 減衰中%／販促%が空 → マスタを埋めてから restore")
        if kind == "apply" and before is None:
            lines.append("  注: 最終終着%空 → 減衰フロアは1%既定（fetchで埋めない）")
        lines.append("  施策行: %s" % spreadsheet_row_url(ssid, gid_sale, row_n))
        if mrow:
            # マスタ行番号は呼び出し側で埋めない場合あり → SKU検索用にマスタシートリンク
            lines.append(
                "  マスタ: %s"
                % spreadsheet_row_url(ssid, gid_master, 1)
            )
        lines.append("")
    lines.append("件数: %s" % len(rows))
    return "\n".join(lines)


def write_draft_sheet(cfg: dict, *, to_addr: str, subject: str, body: str) -> str:
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    if not sid:
        raise RuntimeError("ads_spreadsheet_id がありません")
    token = secrets.token_urlsafe(24)
    svc = sheets_service(write=True)
    ensure_sheet(svc, sid, DRAFT_SHEET)
    values = [
        ["to", to_addr],
        ["subject", subject],
        ["body", body],
        ["sent_at", ""],
        ["web_token", token],
    ]
    q = DRAFT_SHEET.replace("'", "''")
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range=f"'{q}'!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()
    LOG.info("下書きシート更新: %s", DRAFT_SHEET)
    return token


def run_kind(
    *,
    cfg: dict,
    sales: List[Dict[str, Any]],
    master_map: Dict[str, Dict[str, Any]],
    kind: str,
    days: int,
    tol: int,
    today: date,
    ssid: str,
    gid_sale: int,
    gid_master: int,
    do_send: bool,
    stamp: str,
    include_done: bool = False,
) -> Optional[str]:
    picked = select_b_rows(sales, today=today, kind=kind, days=days, tol=tol)
    filtered: List[Tuple[Dict[str, Any], int, Dict[str, Any]]] = []
    for sale, row_n in picked:
        sku = str(sale.get("SKU") or "").strip()
        mrow = master_map.get(sku)
        if not include_done:
            if kind == "apply" and not needs_apply_remind(mrow):
                continue
            if kind == "restore" and not needs_restore_remind(mrow):
                continue
        filtered.append((sale, row_n, mrow or {}))
    if not filtered:
        LOG.info("%s days=%s 該当なし", kind, days)
        return None
    body = build_body(
        kind=kind,
        rows=filtered,
        today=today,
        days=days,
        ssid=ssid,
        gid_sale=gid_sale,
        gid_master=gid_master,
    )
    subject = SUBJECT_APPLY if kind == "apply" else SUBJECT_RESTORE
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    path = out_dir / ("points_remind_%s_D%s_%s.txt" % (kind, days, stamp))
    path.write_text(body, encoding="utf-8")
    LOG.info("下書き %s (%s件)", path, len(filtered))
    to_addr = str(cfg.get("notify_email_to") or "contact@octas2301.com").strip()
    try:
        write_draft_sheet(cfg, to_addr=to_addr, subject=subject, body=body)
    except Exception as e:
        LOG.warning("下書きシート失敗: %s", e)
    if do_send:
        try:
            if send_notify(cfg, subject, body):
                LOG.info("送信OK → %s", to_addr)
            else:
                LOG.warning(
                    "自動送信未完了。スプシ「%s」またはメニュー「⏱ ポイントリマインド送信」で送信可",
                    DRAFT_SHEET,
                )
        except Exception as e:
            LOG.warning("送信失敗: %s", e)
    return str(path)

def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="ポイントリマインド（apply/restore）")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument(
        "--kind",
        choices=("apply", "restore", "both"),
        default="both",
        help="both=apply(T-days)+restore(終了+days)",
    )
    ap.add_argument(
        "--days",
        type=int,
        default=1,
        help="apply=開始まで残り日／restore=終了からの経過日（既定1）",
    )
    ap.add_argument("--tol", type=int, default=0, help="±日の許容")
    ap.add_argument("--today", type=str, default=None)
    ap.add_argument(
        "--send",
        action="store_true",
        help="notify_email へ送信（Gmail API／SMTP）",
    )
    ap.add_argument(
        "--include-done",
        action="store_true",
        help="期間中適用済／セール前復元済でも対象に含める（予行・再送用）",
    )
    args = ap.parse_args(argv)

    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))
    today = parse_ymd(args.today) if args.today else date.today()
    assert today

    svc = sheets_service(write=False)
    ssid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    gid_sale = sheet_gid(svc, ssid, SALE_SHEET)
    gid_master = sheet_gid(svc, ssid, MASTER_SHEET)
    _h, sales = read_sheet_rows(svc, ssid, SALE_SHEET)
    _mh, master = read_sheet_rows(svc, ssid, MASTER_SHEET)
    mmap = master_by_sku(master)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    kinds = ["apply", "restore"] if args.kind == "both" else [args.kind]
    drafts = []
    for kind in kinds:
        p = run_kind(
            cfg=cfg,
            sales=sales,
            master_map=mmap,
            kind=kind,
            days=int(args.days),
            tol=int(args.tol or 0),
            today=today,
            ssid=ssid,
            gid_sale=gid_sale,
            gid_master=gid_master,
            do_send=bool(args.send),
            stamp=stamp,
            include_done=bool(args.include_done),
        )
        if p:
            drafts.append(p)
    print(
        json.dumps(
            {
                "today": today.isoformat(),
                "kind": args.kind,
                "days": args.days,
                "drafts": drafts,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
