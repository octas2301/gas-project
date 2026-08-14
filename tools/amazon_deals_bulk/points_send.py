# -*- coding: utf-8 -*-
"""
タイムセール_マスタ → Amazon Points フィード（§10.10 Phase0）。

--mode apply   : 期間中ポイント% を送信（既定1%）
--mode restore : 減衰中ポイント%（カレンダー位置）に戻す

既定: TSV生成のみ（dry_run）
本番: --prod --i-confirm-prod かつ spapi allow_prod=true

例:
  python points_send.py
  python points_send.py --mode restore
  python points_send.py --all-master          # 施策連動オフ（有効マスタ全体）
  python points_send.py --prod --i-confirm-prod --wait --update-sheet
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

LOG = logging.getLogger("amazon_deals_bulk.points_send")

try:
    import requests
except ImportError:
    raise SystemExit("requests が必要です") from None

from lane_a_send import lwa_token, resolve_spapi_cfg  # noqa: E402
from paths import load_config  # noqa: E402
from points_logic import (  # noqa: E402
    MODE_APPLY,
    MODE_RESTORE,
    build_points_tsv,
    diff_summary,
    period_percent,
    sale_skus_for_points,
    select_diff_rows,
    send_percent,
    skus_missing_before,
    status_after_send,
)
from sheet_schema import (  # noqa: E402
    MASTER_HEADERS,
    MASTER_SHEET,
    POINT_CURRENT_COL,
    POINT_STATUS_COL,
    SALE_SHEET,
)
from sheets_io import read_sheet_rows, sheets_service, write_headers_and_rows  # noqa: E402

FEED_TYPE = "POST_FLAT_FILE_OFFER_POINTS_PREFERENCE_DATA"
CONTENT_TYPE = "text/tab-separated-values; charset=UTF-8"
MARKETPLACE_JP = "A1VC38T7YXB528"


def _spapi_headers(endpoint: str, access_token: str, user_agent: str) -> Dict[str, str]:
    return {
        "host": endpoint.replace("https://", "").replace("http://", "").split("/")[0],
        "x-amz-access-token": access_token,
        "x-amz-date": datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ"),
        "user-agent": user_agent,
        "accept": "application/json",
    }


def create_feed_document(
    *,
    endpoint: str,
    access_token: str,
    user_agent: str,
) -> dict:
    url = "%s/feeds/2021-06-30/documents" % endpoint.rstrip("/")
    resp = requests.post(
        url,
        headers={**_spapi_headers(endpoint, access_token, user_agent), "content-type": "application/json"},
        data=json.dumps({"contentType": CONTENT_TYPE}),
        timeout=60,
    )
    if resp.status_code >= 300:
        raise RuntimeError("createFeedDocument HTTP %s %s" % (resp.status_code, resp.text[:500]))
    return resp.json()


def upload_feed_body(upload_url: str, body: bytes) -> None:
    resp = requests.put(
        upload_url,
        data=body,
        headers={"Content-Type": CONTENT_TYPE},
        timeout=120,
    )
    if resp.status_code >= 300:
        raise RuntimeError("feed upload HTTP %s %s" % (resp.status_code, resp.text[:500]))


def create_feed(
    *,
    endpoint: str,
    access_token: str,
    user_agent: str,
    document_id: str,
    marketplace_id: str,
) -> dict:
    url = "%s/feeds/2021-06-30/feeds" % endpoint.rstrip("/")
    payload = {
        "feedType": FEED_TYPE,
        "marketplaceIds": [marketplace_id],
        "inputFeedDocumentId": document_id,
    }
    resp = requests.post(
        url,
        headers={**_spapi_headers(endpoint, access_token, user_agent), "content-type": "application/json"},
        data=json.dumps(payload),
        timeout=60,
    )
    if resp.status_code >= 300:
        raise RuntimeError("createFeed HTTP %s %s" % (resp.status_code, resp.text[:500]))
    return resp.json()


def get_feed(
    *,
    endpoint: str,
    access_token: str,
    user_agent: str,
    feed_id: str,
) -> dict:
    url = "%s/feeds/2021-06-30/feeds/%s" % (endpoint.rstrip("/"), feed_id)
    resp = requests.get(
        url,
        headers=_spapi_headers(endpoint, access_token, user_agent),
        timeout=60,
    )
    if resp.status_code >= 300:
        raise RuntimeError("getFeed HTTP %s %s" % (resp.status_code, resp.text[:500]))
    return resp.json()


def wait_feed(
    *,
    endpoint: str,
    access_token: str,
    user_agent: str,
    feed_id: str,
    timeout_sec: int = 180,
    poll_sec: int = 8,
) -> dict:
    deadline = time.time() + timeout_sec
    last: dict = {}
    while time.time() < deadline:
        last = get_feed(
            endpoint=endpoint,
            access_token=access_token,
            user_agent=user_agent,
            feed_id=feed_id,
        )
        status = str(last.get("processingStatus") or "")
        LOG.info("feed %s status=%s", feed_id, status)
        if status in ("DONE", "CANCELLED", "FATAL"):
            return last
        time.sleep(poll_sec)
    return last


def load_master_rows(cfg: dict) -> List[Dict[str, Any]]:
    svc = sheets_service(write=False)
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    _h, rows = read_sheet_rows(svc, sid, MASTER_SHEET)
    return rows


def apply_sheet_after_send(
    cfg: dict,
    sent_rows: List[Dict[str, Any]],
    *,
    mode: str,
    status: str,
    today=None,
) -> None:
    """送信成功SKUの 出品者ポイント現在%・状態 を更新。"""
    svc = sheets_service(write=True)
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    _h, rows = read_sheet_rows(svc, sid, MASTER_SHEET)
    want = {str(r.get("SKU") or "").strip(): r for r in sent_rows}
    values: List[List[Any]] = []
    for r in rows:
        sku = str(r.get("SKU") or "").strip()
        d = {h: r.get(h, "") for h in MASTER_HEADERS}
        if sku in want:
            d[POINT_CURRENT_COL] = str(send_percent(want[sku], mode, today=today))
            d[POINT_STATUS_COL] = status
        values.append([d.get(h, "") for h in MASTER_HEADERS])
    write_headers_and_rows(svc, sid, MASTER_SHEET, MASTER_HEADERS, values, clear=True)
    LOG.info("マスタ更新: %s SKU → %s", len(want), status)


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="タイムセール_マスタ Pointsフィード")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--sku", type=str, default="", help="1SKUのみ")
    ap.add_argument("--all", action="store_true", help="差分無視（強制送信用。対象集合は下記）")
    ap.add_argument(
        "--all-master",
        action="store_true",
        help="施策連動オフ: マスタ有効SKU全体（従来）。既定は施策Bの接近/直後のみ",
    )
    ap.add_argument(
        "--within-days",
        type=int,
        default=1,
        help="施策連動: apply=開始まで0..N日or実施中 / restore=終了から0..N日（既定1）",
    )
    ap.add_argument(
        "--today",
        type=str,
        default=None,
        help="base date YYYY-MM-DD (sale window and restore calendar pct)",
    )
    ap.add_argument(
        "--mode",
        choices=(MODE_APPLY, MODE_RESTORE),
        default=MODE_APPLY,
        help="apply=period percent / restore=calendar taper percent",
    )
    ap.add_argument(
        "--prod",
        action="store_true",
        help="Points feed prod write. Only after explicit approval",
    )
    ap.add_argument("--i-confirm-prod", action="store_true")
    ap.add_argument(
        "--update-sheet",
        action="store_true",
        help="After prod, update current percent and status on master",
    )
    ap.add_argument("--wait", action="store_true", help="getFeed で DONE まで待機")
    ap.add_argument(
        "--backup-before",
        action="store_true",
        help="apply前に points_fetch でセール前が空なら退避（--write相当）",
    )
    ap.add_argument(
        "--allow-missing-before",
        action="store_true",
        help="P0-G8 override: continue apply even if before-percent empty or backup failed",
    )
    args = ap.parse_args(argv)
    mode = args.mode
    from schedule_class import parse_ymd as _parse_ymd_today

    today_d = _parse_ymd_today(args.today) if args.today else date.today()
    assert today_d

    local = HERE / "config.local.json"
    deals_cfg = load_config(
        args.config or (local if local.is_file() else HERE / "config.example.json")
    )

    if mode == MODE_APPLY and args.backup_before:
        from points_fetch import main as fetch_main  # noqa: WPS433

        fetch_argv = ["--write"]
        if args.sku:
            fetch_argv.extend(["--sku", args.sku])
        if args.config:
            fetch_argv.extend(["--config", str(args.config)])
        LOG.info("セール前退避を実行: points_fetch %s", " ".join(fetch_argv))
        rc = fetch_main(fetch_argv)
        if rc != 0:
            if args.allow_missing_before:
                LOG.warning("points_fetch rc=%s（--allow-missing-before で続行）", rc)
            else:
                LOG.error(
                    "points_fetch 失敗 rc=%s。apply中止（強行は --allow-missing-before）",
                    rc,
                )
                return 1
    rows = load_master_rows(deals_cfg)
    sku_allow = None
    sale_linked = not bool(args.all_master) and not str(args.sku or "").strip()
    if sale_linked:
        svc = sheets_service(write=False)
        sid = str(deals_cfg.get("ads_spreadsheet_id") or "").strip()
        _sh, sales = read_sheet_rows(svc, sid, SALE_SHEET)
        sku_allow = sale_skus_for_points(
            sales,
            mode=mode,
            today=today_d,
            within_days=int(args.within_days),
        )
        LOG.info(
            "施策連動 mode=%s within_days=%s today=%s SKU=%s",
            mode,
            args.within_days,
            today_d.isoformat(),
            len(sku_allow),
        )
        if not sku_allow:
            LOG.info("施策連動の対象SKUなし（--all-master または --sku で解除可）")
            print(
                json.dumps(
                    {
                        "diff": 0,
                        "mode": mode,
                        "sale_linked": True,
                        "sale_skus": 0,
                    },
                    ensure_ascii=False,
                )
            )
            return 0
    try:
        targets = select_diff_rows(
            rows,
            mode=mode,
            sku_filter=args.sku or None,
            force_all=bool(args.all),
            enabled_only=True,
            sku_allow=sku_allow,
            today=today_d,
        )
    except ValueError as e:
        LOG.error("%s", e)
        return 1

    if not targets:
        LOG.info("差分なし mode=%s sale_linked=%s", mode, sale_linked)
        print(
            json.dumps(
                {"diff": 0, "mode": mode, "sale_linked": sale_linked},
                ensure_ascii=False,
            )
        )
        return 0

    # 最終終着%空は減衰フロア1%既定。apply を止めない（restore先は減衰中%）
    if mode == MODE_APPLY:
        missing = skus_missing_before(targets)
        if missing:
            LOG.warning(
                "最終終着%%（セール前列）空→減衰フロア1%%既定: %s",
                ", ".join(missing[:10]),
            )

    for r in targets:
        try:
            send_percent(r, mode, today=today_d)
        except ValueError as e:
            LOG.error("SKU=%s: %s", r.get("SKU"), e)
            return 1

    tsv = build_points_tsv(targets, mode=mode, today=today_d)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    tsv_path = out_dir / ("points_feed_%s_%s.tsv" % (mode, stamp))
    tsv_path.write_text(tsv, encoding="utf-8")
    summary = [
        {"sku": s, "current": c, "send": t, "before": b}
        for s, c, t, b in diff_summary(targets, mode, today=today_d)
    ]
    meta: Dict[str, Any] = {
        "stamp": stamp,
        "mode": mode,
        "feed_type": FEED_TYPE,
        "prod": bool(args.prod),
        "count": len(targets),
        "sale_linked": sale_linked,
        "sale_sku_count": len(sku_allow) if sku_allow is not None else None,
        "within_days": int(args.within_days) if sale_linked else None,
        "tsv": str(tsv_path),
        "summary": summary,
    }
    LOG.info("mode=%s 差分 %s件 → %s", mode, len(targets), tsv_path)
    for row in summary:
        LOG.info(
            "  %s: current=%s send=%s before=%s period_default_hint=%s",
            row["sku"],
            row["current"],
            row["send"],
            row["before"],
            period_percent(next(x for x in targets if x.get("SKU") == row["sku"])),
        )

    if not args.prod:
        meta_path = out_dir / ("points_send_%s_%s.json" % (mode, stamp))
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        print(
            json.dumps(
                {
                    "dry_run": True,
                    "mode": mode,
                    "count": len(targets),
                    "sale_linked": sale_linked,
                    "tsv": str(tsv_path),
                    "meta": str(meta_path),
                },
                ensure_ascii=False,
            )
        )
        return 0

    if not args.i_confirm_prod:
        LOG.error("--prod には --i-confirm-prod も必要です")
        return 1
    spapi = resolve_spapi_cfg(deals_cfg)
    if not bool(spapi.get("allow_prod")):
        LOG.error("spapi config の allow_prod=true が必要です")
        return 1

    token = lwa_token(spapi)
    endpoint = str(spapi.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    ua = str(spapi.get("user_agent") or "OctasAmazonDealsPoints/0.1")
    marketplace_id = str(
        spapi.get("marketplace_id") or deals_cfg.get("marketplace_id") or MARKETPLACE_JP
    )

    LOG.warning("PROD Pointsフィード mode=%s count=%s", mode, len(targets))
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

    if args.wait and feed_id:
        meta["feed"] = wait_feed(
            endpoint=endpoint,
            access_token=token,
            user_agent=ua,
            feed_id=feed_id,
        )

    meta_path = out_dir / ("points_send_%s_%s.json" % (mode, stamp))
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.update_sheet:
        feed_st = None
        if meta.get("feed"):
            feed_st = str(meta["feed"].get("processingStatus") or "") or None
        status = status_after_send(mode, feed_st)
        apply_sheet_after_send(
            deals_cfg, targets, mode=mode, status=status, today=today_d
        )

    print(
        json.dumps(
            {
                "prod": True,
                "mode": mode,
                "count": len(targets),
                "feedId": feed_id,
                "tsv": str(tsv_path),
                "meta": str(meta_path),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
