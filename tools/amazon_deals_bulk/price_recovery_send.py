# -*- coding: utf-8 -*-
"""
実質戻し CLI（§10.12 改訂）。

- 既定 dry_run: ポイント減衰1段の計画JSON（from%/to%）
- --snap: our_price を目標売価へ一発（VALIDATION_PREVIEW／--prod）
- ポイント減衰の本番フィード送信は後続（points_send 拡張）。本CLIの --prod は --snap 時のみ可

人必須入力: 目標売価円＋販促ポイント%。提案で減衰期間／減衰段%／減衰間隔等を埋める。
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

LOG = logging.getLogger("amazon_deals_bulk.price_recovery_send")

from lane_a_send import lwa_token, patch_listings_item, resolve_spapi_cfg  # noqa: E402
from paths import load_config  # noqa: E402
from price_recovery_logic import (  # noqa: E402
    _num,
    build_our_price_patch,
    plan_one_points_step,
    points_blocks_recovery,
    select_due_recovery_rows,
    skus_in_active_b,
)
from schedule_class import parse_ymd  # noqa: E402
from sheet_schema import (  # noqa: E402
    MASTER_SHEET,
    PRICE_CURRENT_SELL_COL,
    PRICE_TARGET_COL,
    SALE_SHEET,
)
from sheets_io import read_sheet_rows, sheets_service, update_row_fields  # noqa: E402

MARKETPLACE_JP = "A1VC38T7YXB528"


def _enabled(row: Dict[str, Any]) -> bool:
    v = row.get("有効")
    if v is True:
        return True
    return str(v or "").strip().upper() in ("TRUE", "はい", "YES", "Y", "1", "○")


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="実質戻し（ポイント減衰計画／売価スナップ）")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--sku", type=str, default="", help="1SKUに限定")
    ap.add_argument("--start", action="store_true", help="次回減衰日が空の未開始行も初回対象")
    ap.add_argument("--today", type=str, default=None)
    ap.add_argument("--snap", action="store_true", help="our_price を目標売価円へ一発")
    ap.add_argument("--prod", action="store_true", help="API本番（--snap 時のみ）")
    ap.add_argument("--i-confirm-prod", action="store_true")
    ap.add_argument("--update-sheet", action="store_true")
    ap.add_argument("--allow-during-b", action="store_true")
    ap.add_argument("--allow-before-restore", action="store_true")
    ap.add_argument("--dry-run-preview", action="store_true", help="--snap 時 VALIDATION_PREVIEW")
    args = ap.parse_args(argv)

    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))
    today = parse_ymd(args.today) if args.today else date.today()
    assert today

    svc = sheets_service(write=False)
    ssid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    _mh, master = read_sheet_rows(svc, ssid, MASTER_SHEET)
    _sh, sales = read_sheet_rows(svc, ssid, SALE_SHEET)
    active_b = skus_in_active_b(sales, today=today)
    sku_f = str(args.sku or "").strip() or None

    if args.snap:
        return run_snap(
            cfg=cfg,
            ssid=ssid,
            master=master,
            active_b=active_b,
            today=today,
            sku_filter=sku_f,
            args=args,
        )

    due = select_due_recovery_rows(
        master,
        today=today,
        sku_filter=sku_f,
        include_start=bool(args.start),
    )
    plans: List[Dict[str, Any]] = []
    skipped: List[Dict[str, str]] = []
    for row in due:
        sku = str(row.get("SKU") or "").strip()
        if sku in active_b and not args.allow_during_b:
            skipped.append({"sku": sku, "reason": "B期間中"})
            continue
        if points_blocks_recovery(row) and not args.allow_before_restore:
            skipped.append({"sku": sku, "reason": "ポイント期間中適用済"})
            continue
        try:
            plan = plan_one_points_step(row, today=today)
        except ValueError as e:
            skipped.append({"sku": sku, "reason": str(e)})
            continue
        plans.append(plan)

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    summary = [
        {
            "sku": p["sku"],
            "from_pct": p["from_pct"],
            "to_pct": p["to_pct"],
            "status": p["status"],
            "next": p["next_date"].isoformat() if p.get("next_date") else None,
            "skip_api": bool(p.get("skip_api")),
        }
        for p in plans
    ]
    meta: Dict[str, Any] = {
        "stamp": stamp,
        "today": today.isoformat(),
        "mode": "points_taper_plan",
        "count": len(plans),
        "skipped": skipped,
        "summary": summary,
        "note": "減衰の本番送信は後続（points_send拡張）。本出力は計画のみ。",
    }
    for s in skipped:
        LOG.info("skip %s: %s", s["sku"], s["reason"])
    for p in plans:
        LOG.info(
            "plan %s: %s%% → %s%% (%s) next=%s",
            p["sku"],
            p["from_pct"],
            p["to_pct"],
            p["status"],
            p["next_date"],
        )
    meta_path = out_dir / ("price_recovery_taper_%s.json" % stamp)
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    if args.prod:
        LOG.error("--prod は --snap と併用してください（ポイント減衰フィードは未接続）")
        print(json.dumps({"error": "use --snap for prod", "meta": str(meta_path)}, ensure_ascii=False))
        return 1
    print(
        json.dumps(
            {
                "dry_run": True,
                "mode": "points_taper_plan",
                "count": len(plans),
                "skipped": skipped,
                "summary": summary,
                "meta": str(meta_path),
            },
            ensure_ascii=False,
        )
    )
    return 0


def run_snap(*, cfg, ssid, master, active_b, today, sku_filter, args) -> int:
    targets = []
    skipped = []
    for row in master:
        sku = str(row.get("SKU") or "").strip()
        if not sku:
            continue
        if sku_filter and sku != sku_filter:
            continue
        if not _enabled(row):
            continue
        price = _num(row.get(PRICE_TARGET_COL))
        if not price or price <= 0:
            skipped.append({"sku": sku, "reason": "目標売価円なし"})
            continue
        if sku in active_b and not args.allow_during_b:
            skipped.append({"sku": sku, "reason": "B期間中"})
            continue
        targets.append({"sku": sku, "price": price, "row": row})

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    summary = [{"sku": t["sku"], "our_price": t["price"]} for t in targets]
    meta: Dict[str, Any] = {
        "stamp": stamp,
        "mode": "snap",
        "today": today.isoformat(),
        "prod": bool(args.prod),
        "count": len(targets),
        "skipped": skipped,
        "summary": summary,
    }
    for s in skipped:
        LOG.info("skip %s: %s", s["sku"], s["reason"])
    for t in targets:
        LOG.info("snap %s → %s", t["sku"], t["price"])

    if not targets:
        meta_path = out_dir / ("price_recovery_snap_%s.json" % stamp)
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"count": 0, "skipped": skipped, "meta": str(meta_path)}, ensure_ascii=False))
        return 0

    do_api = bool(args.prod) or bool(args.dry_run_preview)
    if not do_api:
        meta_path = out_dir / ("price_recovery_snap_%s.json" % stamp)
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        print(
            json.dumps(
                {
                    "dry_run": True,
                    "mode": "snap",
                    "count": len(targets),
                    "summary": summary,
                    "meta": str(meta_path),
                },
                ensure_ascii=False,
            )
        )
        return 0

    if args.prod and not args.i_confirm_prod:
        LOG.error("--prod には --i-confirm-prod も必要です")
        return 1
    spapi = resolve_spapi_cfg(cfg)
    if args.prod and not bool(spapi.get("allow_prod")):
        LOG.error("spapi config の allow_prod=true が必要です")
        return 1
    token = lwa_token(spapi)
    endpoint = str(spapi.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    ua = str(spapi.get("user_agent") or "OctasAmazonDealsPriceSnap/0.1")
    seller_id = str(spapi.get("seller_id") or "").strip()
    marketplace_id = str(spapi.get("marketplace_id") or cfg.get("marketplace_id") or MARKETPLACE_JP)
    validation_preview = not bool(args.prod)
    results = []
    ok_skus = []
    for t in targets:
        body = build_our_price_patch(
            marketplace_id=marketplace_id,
            currency=str(t["row"].get("通貨") or "JPY"),
            our_price=float(t["price"]),
        )
        resp = patch_listings_item(
            endpoint=endpoint,
            access_token=token,
            seller_id=seller_id,
            sku=t["sku"],
            marketplace_id=marketplace_id,
            body=body,
            user_agent=ua,
            validation_preview=validation_preview,
        )
        item: Dict[str, Any] = {"sku": t["sku"], "http_status": resp.status_code, "to": t["price"]}
        try:
            item["response"] = resp.json()
        except Exception:
            item["response_text"] = resp.text[:2000]
        results.append(item)
        if 200 <= resp.status_code < 300:
            ok_skus.append(t)
            LOG.info("OK snap %s HTTP %s", t["sku"], resp.status_code)
        else:
            LOG.error("FAIL snap %s HTTP %s", t["sku"], resp.status_code)
    meta["results"] = results
    meta_path = out_dir / ("price_recovery_snap_%s.json" % stamp)
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    if args.prod and args.update_sheet and ok_skus:
        svc_w = sheets_service(write=True)
        _h, rows = read_sheet_rows(svc_w, ssid, MASTER_SHEET)
        by = {t["sku"]: t for t in ok_skus}
        for i, r in enumerate(rows):
            sku = str(r.get("SKU") or "").strip()
            if sku not in by:
                continue
            update_row_fields(
                svc_w,
                ssid,
                MASTER_SHEET,
                i + 2,
                {PRICE_CURRENT_SELL_COL: by[sku]["price"]},
            )
        LOG.info("マスタ現在売価更新 %s件", len(by))
    failed = [x for x in results if not (200 <= int(x.get("http_status") or 0) < 300)]
    print(
        json.dumps(
            {
                "mode": "snap",
                "prod": bool(args.prod),
                "ok": len(ok_skus),
                "failed": len(failed),
                "meta": str(meta_path),
            },
            ensure_ascii=False,
        )
    )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
