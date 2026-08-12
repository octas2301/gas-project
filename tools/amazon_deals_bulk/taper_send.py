# -*- coding: utf-8 -*-
"""
ポイント減衰1段（P1c-1／E）。

- 既定 dry_run: 計画JSON＋任意メール
- --prod --i-confirm-prod: Pointsフィード（既存 points_send と同一 FEED_TYPE）
- --poll: 次回減衰日≦今日 または 減衰実行依頼=TRUE
- --mail: 結果メール（商品URL付き）。リカバリURLは今回スコープ外（手動）

例:
  python taper_send.py --poll --mail
  python taper_send.py --sku "…" --start --mail
  python taper_send.py --sku "…" --prod --i-confirm-prod --update-sheet --mail
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

LOG = logging.getLogger("amazon_deals_bulk.taper_send")

from notify_mail import send_notify  # noqa: E402
from paths import load_config  # noqa: E402
from points_logic import POINT_STATUS_FEED_PREFIX  # noqa: E402
from points_send import (  # noqa: E402
    MARKETPLACE_JP,
    create_feed,
    create_feed_document,
    upload_feed_body,
    wait_feed,
)
from price_recovery_logic import (  # noqa: E402
    apply_plan_to_row,
    last_run_date,
    plan_calendar_sync,
    plan_one_points_step,
    points_blocks_recovery,
    select_due_recovery_rows,
    skus_in_active_b,
    taper_requested,
    taper_start_date,
)
from schedule_class import parse_ymd  # noqa: E402
from sheet_schema import (  # noqa: E402
    MASTER_SHEET,
    POINT_CURRENT_COL,
    POINT_STATUS_COL,
    PRICE_RECOVERY_NEXT_COL,
    PRICE_RECOVERY_PROGRESS_COL,
    PRICE_RECOVERY_STATUS_COL,
    SALE_SHEET,
    TAPER_ACTIVE_COL,
    TAPER_LAST_RUN_COL,
    TAPER_REQUEST_COL,
)
from sheets_io import read_sheet_rows, sheets_service, update_row_fields  # noqa: E402
from lane_a_send import lwa_token, resolve_spapi_cfg  # noqa: E402

AMAZON_DP = "https://www.amazon.co.jp/dp/"


def amazon_jp_url(asin: str) -> str:
    a = str(asin or "").strip().upper()
    if a.startswith("B0") and len(a) >= 10:
        return AMAZON_DP + a
    return ""


def build_points_tsv_explicit(items: List[Dict[str, Any]]) -> str:
    lines = ["sku\tpoints_percent"]
    for it in items:
        sku = str(it.get("sku") or "").strip()
        if not sku:
            continue
        pct = int(round(float(it["to_pct"])))
        lines.append("%s\t%s" % (sku, max(0, pct)))
    return "\n".join(lines) + "\n"


def build_taper_mail(
    *,
    today: date,
    prod: bool,
    ok: List[Dict[str, Any]],
    skipped: List[Dict[str, str]],
    failed: List[Dict[str, str]],
    feed_id: str = "",
    sheet_only: Optional[List[Dict[str, Any]]] = None,
) -> tuple[str, str]:
    n_ok = len(ok)
    n_fail = len(failed)
    sheet_only = sheet_only or []
    kind = "本番" if prod else "dry_run"
    subject = "[Amazonポイント減衰] %s %s 成功%s／失敗%s" % (
        today.isoformat(),
        kind,
        n_ok,
        n_fail,
    )
    lines = [
        "ポイント減衰の結果です（細かい事前承認はありません。異常なら店頭URLで確認し、手動で%を直してください）。",
        "",
        "モード: %s" % kind,
        "成功: %s / 失敗: %s / シートのみ: %s / スキップ: %s"
        % (n_ok, n_fail, len(sheet_only), len(skipped)),
    ]
    if feed_id:
        lines.append("feedId: %s" % feed_id)
    lines.append("")
    if ok:
        lines.append("【成功】")
        for it in ok:
            url = amazon_jp_url(str(it.get("asin") or ""))
            lines.append(
                "- %s / %s / %s : %s%% → %s%%"
                % (
                    it.get("sku"),
                    it.get("asin") or "-",
                    (it.get("name") or "")[:40],
                    it.get("from_pct"),
                    it.get("to_pct"),
                )
            )
            if url:
                lines.append("  %s" % url)
        lines.append("")
    if sheet_only:
        lines.append("【シートのみ（B中カレンダー／期間中適用済）】")
        for it in sheet_only:
            lines.append(
                "- %s : %s%% → %s%%（%s）"
                % (
                    it.get("sku"),
                    it.get("from_pct"),
                    it.get("to_pct"),
                    it.get("sheet_only_reason") or "sheet_only",
                )
            )
        lines.append("")
    if failed:
        lines.append("【失敗】")
        for it in failed:
            lines.append("- %s : %s" % (it.get("sku"), it.get("reason")))
        lines.append("")
    if skipped:
        lines.append("【スキップ】")
        for it in skipped:
            lines.append("- %s : %s" % (it.get("sku"), it.get("reason")))
        lines.append("")
    lines.extend(
        [
            "手動リカバリ:",
            "1. 上の商品URLで店頭ポイントを確認",
            "2. B終了後は先に taper_send --poll でカレンダー同期 → points_send --mode restore（減衰中%へ）",
            "3. 間違っていれば SC 画面、または",
            '   python points_send.py --sku "…" --all --prod --i-confirm-prod --wait --update-sheet',
            "   （先にマスタの期間中ポイント% or 減衰中ポイント%をあるべき値にしてから）",
            "",
            "要件§10.14: docs/org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md",
        ]
    )
    return subject, "\n".join(lines)


def _sku_row_index(master: List[Dict[str, Any]]) -> Dict[str, int]:
    """read_sheet_rows の順序 → シート行番号（ヘッダ=1、先頭データ=2）。空行穴は非対応。"""
    out: Dict[str, int] = {}
    for i, r in enumerate(master):
        sku = str(r.get("SKU") or "").strip()
        if sku and sku not in out:
            out[sku] = i + 2
    return out


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="ポイント減衰1段（taper）")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--sku", type=str, default="")
    ap.add_argument("--today", type=str, default=None)
    ap.add_argument("--start", action="store_true", help="次回減衰日が空の未開始も対象")
    ap.add_argument("--poll", action="store_true", help="期日到来＋減衰実行依頼を対象")
    ap.add_argument("--prod", action="store_true")
    ap.add_argument("--i-confirm-prod", action="store_true")
    ap.add_argument("--update-sheet", action="store_true")
    ap.add_argument("--wait", action="store_true")
    ap.add_argument("--mail", action="store_true", help="結果メール")
    ap.add_argument("--allow-during-b", action="store_true")
    ap.add_argument("--allow-before-restore", action="store_true")
    args = ap.parse_args(argv)

    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))
    today = parse_ymd(args.today) if args.today else date.today()
    assert today

    svc_ro = sheets_service(write=False)
    ssid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    _mh, master = read_sheet_rows(svc_ro, ssid, MASTER_SHEET)
    _sh, sales = read_sheet_rows(svc_ro, ssid, SALE_SHEET)
    active_b = skus_in_active_b(sales, today=today)
    sku_f = str(args.sku or "").strip() or None
    include_start = bool(args.start) or bool(args.poll)

    due = select_due_recovery_rows(
        master,
        today=today,
        sku_filter=sku_f,
        include_start=include_start,
    )
    row_index = _sku_row_index(master)

    plans: List[Dict[str, Any]] = []
    skipped: List[Dict[str, str]] = []
    for row in due:
        sku = str(row.get("SKU") or "").strip()
        lr = last_run_date(row)
        if lr == today and not taper_requested(row):
            skipped.append({"sku": sku, "reason": "本日実行済"})
            continue
        try:
            if taper_start_date(row):
                plan = plan_calendar_sync(row, today=today)
            else:
                plan = plan_one_points_step(row, today=today)
        except ValueError as e:
            skipped.append({"sku": sku, "reason": str(e)})
            continue
        in_b = sku in active_b and not args.allow_during_b
        applied_block = points_blocks_recovery(row) and not args.allow_before_restore
        if in_b:
            plan["sheet_only"] = True
            plan["sheet_only_reason"] = "B期間中カレンダー"
        elif applied_block:
            plan["sheet_only"] = True
            plan["sheet_only_reason"] = "期間中適用済→先にrestore"
        plan["asin"] = str(row.get("ASIN") or "").strip()
        plan["name"] = str(row.get("商品名") or "").strip()
        plan["_row"] = row
        plans.append(plan)

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)

    sheet_only_plans = [p for p in plans if p.get("sheet_only")]
    sendable = [p for p in plans if not p.get("skip_api") and not p.get("sheet_only")]
    done_only = [p for p in plans if p.get("skip_api") and not p.get("sheet_only")]
    summary = [
        {
            "sku": p["sku"],
            "asin": p.get("asin"),
            "from_pct": p["from_pct"],
            "to_pct": p["to_pct"],
            "status": p["status"],
            "next": p["next_date"].isoformat() if p.get("next_date") else None,
            "skip_api": bool(p.get("skip_api")),
            "sheet_only": bool(p.get("sheet_only")),
            "sheet_only_reason": p.get("sheet_only_reason") or "",
        }
        for p in plans
    ]
    meta: Dict[str, Any] = {
        "stamp": stamp,
        "today": today.isoformat(),
        "mode": "taper",
        "prod": bool(args.prod),
        "poll": bool(args.poll),
        "count": len(plans),
        "sendable": len(sendable),
        "sheet_only": len(sheet_only_plans),
        "skipped": skipped,
        "summary": summary,
    }

    for s in skipped:
        LOG.info("skip %s: %s", s["sku"], s["reason"])
    for p in plans:
        LOG.info(
            "plan %s: %s%% → %s%% (%s) next=%s skip_api=%s sheet_only=%s",
            p["sku"],
            p["from_pct"],
            p["to_pct"],
            p["status"],
            p.get("next_date"),
            p.get("skip_api"),
            p.get("sheet_only"),
        )

    ok_mail: List[Dict[str, Any]] = []
    failed: List[Dict[str, str]] = []
    feed_id = ""

    if not args.prod:
        meta_path = out_dir / ("taper_send_%s.json" % stamp)
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        print(
            json.dumps(
                {
                    "dry_run": True,
                    "mode": "taper",
                    "count": len(plans),
                    "sendable": len(sendable),
                    "sheet_only": len(sheet_only_plans),
                    "skipped": skipped,
                    "summary": summary,
                    "meta": str(meta_path),
                },
                ensure_ascii=False,
            )
        )
        ok_mail = [
            {
                "sku": p["sku"],
                "asin": p.get("asin"),
                "name": p.get("name"),
                "from_pct": p["from_pct"],
                "to_pct": p["to_pct"],
            }
            for p in sendable
        ]
        if args.mail:
            subj, body = build_taper_mail(
                today=today,
                prod=False,
                ok=ok_mail,
                skipped=skipped,
                failed=failed,
                sheet_only=sheet_only_plans,
            )
            sent = send_notify(cfg, subj, body)
            LOG.info("mail sent=%s", sent)
        return 0

    if not args.i_confirm_prod:
        LOG.error("--prod には --i-confirm-prod も必要です")
        return 1
    if not sendable and not done_only and not sheet_only_plans:
        LOG.info("送信対象なし")
        print(json.dumps({"prod": True, "count": 0, "skipped": skipped}, ensure_ascii=False))
        if args.mail:
            subj, body = build_taper_mail(
                today=today,
                prod=True,
                ok=[],
                skipped=skipped,
                failed=failed,
                sheet_only=sheet_only_plans,
            )
            send_notify(cfg, subj, body)
        return 0

    token = ""
    endpoint = ""
    ua = ""
    marketplace_id = ""
    tsv = ""
    if sendable:
        spapi = resolve_spapi_cfg(cfg)
        if not bool(spapi.get("allow_prod")):
            LOG.error("spapi config の allow_prod=true が必要です")
            return 1
        tsv = build_points_tsv_explicit(sendable)
        tsv_path = out_dir / ("points_feed_taper_%s.tsv" % stamp)
        tsv_path.write_text(tsv, encoding="utf-8")
        meta["tsv"] = str(tsv_path)
        token = lwa_token(spapi)
        endpoint = str(spapi.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
        ua = str(spapi.get("user_agent") or "OctasAmazonDealsTaper/0.1")
        marketplace_id = str(
            spapi.get("marketplace_id") or cfg.get("marketplace_id") or MARKETPLACE_JP
        )

    if sendable:
        LOG.warning("PROD Points taper count=%s", len(sendable))
        try:
            doc = create_feed_document(endpoint=endpoint, access_token=token, user_agent=ua)
            doc_id = str(doc.get("feedDocumentId") or "")
            upload_url = str(doc.get("url") or "")
            if not doc_id or not upload_url:
                raise RuntimeError("createFeedDocument 応答不正: %s" % doc)
            upload_feed_body(upload_url, tsv.encode("utf-8"))
            created = create_feed(
                endpoint=endpoint,
                access_token=token,
                user_agent=ua,
                document_id=doc_id,
                marketplace_id=marketplace_id,
            )
            feed_id = str(created.get("feedId") or "")
            meta["feedDocumentId"] = doc_id
            meta["feedId"] = feed_id
            meta["createFeed"] = created
            feed_st = ""
            if args.wait and feed_id:
                meta["feed"] = wait_feed(
                    endpoint=endpoint,
                    access_token=token,
                    user_agent=ua,
                    feed_id=feed_id,
                )
                feed_st = str(meta["feed"].get("processingStatus") or "")
            if feed_st in ("CANCELLED", "FATAL"):
                failed = [{"sku": p["sku"], "reason": "feed %s" % feed_st} for p in sendable]
            else:
                ok_mail = [
                    {
                        "sku": p["sku"],
                        "asin": p.get("asin"),
                        "name": p.get("name"),
                        "from_pct": p["from_pct"],
                        "to_pct": p["to_pct"],
                    }
                    for p in sendable
                ]
        except Exception as e:
            LOG.exception("taper feed failed: %s", e)
            failed = [{"sku": p["sku"], "reason": str(e)[:200]} for p in sendable]

    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    if args.update_sheet and not failed:
        svc_w = sheets_service(write=True)
        to_apply = sendable + done_only + sheet_only_plans
        sendable_skus = {x["sku"] for x in sendable}
        for p in to_apply:
            sku = p["sku"]
            ridx = row_index.get(sku)
            if not ridx:
                LOG.warning("row not found for %s", sku)
                continue
            apply_plan_to_row(p["_row"], p)
            fields = {
                TAPER_ACTIVE_COL: p["to_pct"],
                PRICE_RECOVERY_STATUS_COL: p["status"],
                PRICE_RECOVERY_PROGRESS_COL: p["progress"],
                PRICE_RECOVERY_NEXT_COL: p["next_date"].isoformat() if p.get("next_date") else "",
                TAPER_LAST_RUN_COL: now_iso,
                TAPER_REQUEST_COL: "",
            }
            if not p.get("sheet_only"):
                fields[POINT_CURRENT_COL] = p["to_pct"]
            if sku in sendable_skus:
                fields[POINT_STATUS_COL] = POINT_STATUS_FEED_PREFIX + (
                    str((meta.get("feed") or {}).get("processingStatus") or "IN_PROGRESS")
                )
            update_row_fields(svc_w, ssid, MASTER_SHEET, ridx, fields)
            LOG.info("sheet updated row=%s sku=%s sheet_only=%s", ridx, sku, bool(p.get("sheet_only")))

    meta_path = out_dir / ("taper_send_%s.json" % stamp)
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(
        json.dumps(
            {
                "prod": True,
                "mode": "taper",
                "count": len(sendable),
                "sheet_only": len(sheet_only_plans),
                "feedId": feed_id,
                "skipped": skipped,
                "failed": failed,
                "meta": str(meta_path),
            },
            ensure_ascii=False,
        )
    )
    if args.mail:
        subj, body = build_taper_mail(
            today=today,
            prod=True,
            ok=ok_mail,
            skipped=skipped,
            failed=failed,
            feed_id=feed_id,
            sheet_only=sheet_only_plans,
        )
        sent = send_notify(cfg, subj, body)
        LOG.info("mail sent=%s", sent)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
