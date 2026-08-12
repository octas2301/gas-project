# -*- coding: utf-8 -*-
"""
レーンA discounted_price を SP-API へ送る（§10.8）。

既定: VALIDATION_PREVIEW（永続化しない dry_run）
本番: --prod かつ config/allow と社長「やってよい」後

例:
  python lane_a_send.py --sku "originalM-1803--KOUS--Cinderellas b" --include-unapproved --before-b-only
  python lane_a_send.py --sku "..." --include-unapproved --before-b-only --prod
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import quote, urlencode

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

LOG = logging.getLogger("amazon_deals_bulk.lane_a_send")

try:
    import requests
except ImportError:
    raise SystemExit("requests が必要です") from None

from lane_a_patch import (  # noqa: E402
    MARKETPLACE_JP,
    build_discounted_price_patch,
    load_a_rows,
    _float,
    _truthy,
)
from paths import load_config  # noqa: E402
from schedule_class import parse_ymd  # noqa: E402
from sheet_schema import SALE_SHEET  # noqa: E402
from sheets_io import read_sheet_rows, sheets_service  # noqa: E402

LWA_TOKEN_URL = "https://api.amazon.com/auth/o2/token"


def _load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_spapi_cfg(deals_cfg: dict) -> dict:
    candidates = [
        HERE.parent / "spapi_listings_write" / "config.local.json",
        HERE.parent / "spapi_smoke" / "config.local.json",
    ]
    auth_rel = str(deals_cfg.get("spapi_auth_config") or "").strip()
    if auth_rel:
        p = Path(auth_rel)
        if not p.is_absolute():
            p = (HERE / p).resolve()
        candidates.insert(0, p)
    for p in candidates:
        if p.is_file():
            cfg = _load_json(p)
            # auth_config_path マージ
            rel = str(cfg.get("auth_config_path") or "").strip()
            if rel:
                ap = Path(rel)
                if not ap.is_absolute():
                    ap = (p.parent / ap).resolve()
                if ap.is_file():
                    auth = _load_json(ap)
                    for k in ("lwa_client_id", "lwa_client_secret", "refresh_token"):
                        if auth.get(k):
                            cfg[k] = auth[k]
            return cfg
    raise SystemExit("SP-API config.local.json が見つかりません（spapi_listings_write / spapi_smoke）")


def lwa_token(cfg: dict) -> str:
    data = {
        "grant_type": "refresh_token",
        "refresh_token": cfg["refresh_token"],
        "client_id": cfg["lwa_client_id"],
        "client_secret": cfg["lwa_client_secret"],
    }
    resp = requests.post(
        LWA_TOKEN_URL,
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"},
        timeout=60,
    )
    if resp.status_code != 200:
        raise RuntimeError("LWA失敗 %s %s" % (resp.status_code, resp.text[:400]))
    return resp.json()["access_token"]


def patch_listings_item(
    *,
    endpoint: str,
    access_token: str,
    seller_id: str,
    sku: str,
    marketplace_id: str,
    body: dict,
    user_agent: str,
    validation_preview: bool,
) -> requests.Response:
    path = "/listings/2021-08-01/items/%s/%s" % (
        quote(seller_id, safe="-_.~"),
        quote(sku, safe="-_.~"),
    )
    q: Dict[str, str] = {"marketplaceIds": marketplace_id}
    if validation_preview:
        q["mode"] = "VALIDATION_PREVIEW"
    url = "%s%s?%s" % (endpoint.rstrip("/"), path, urlencode(q))
    headers = {
        "host": endpoint.replace("https://", "").replace("http://", "").split("/")[0],
        "x-amz-access-token": access_token,
        "x-amz-date": datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ"),
        "user-agent": user_agent,
        "accept": "application/json",
        "content-type": "application/json",
    }
    return requests.patch(
        url,
        headers=headers,
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        timeout=90,
    )


def filter_before_b(cfg: dict, rows: list) -> list:
    today = __import__("datetime").date.today()
    svc = sheets_service(write=False)
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    _h, all_rows = read_sheet_rows(svc, sid, SALE_SHEET)
    b_starts = []
    for r in all_rows:
        if str(r.get("レーン") or "").startswith("B"):
            d = parse_ymd(r.get("開始日"))
            if d and d >= today:
                b_starts.append(d)
    b0 = min(b_starts) if b_starts else None
    out = []
    for r in rows:
        s, e = parse_ymd(r.get("開始日")), parse_ymd(r.get("終了日"))
        if not s or not e or e < today:
            continue
        if b0 and e >= b0:
            continue
        out.append(r)
    return out


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="レーンA discounted_price 送信")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--sku", type=str, default="", help="検証は1SKU必須（--backfill-a-logs 時は不要）")
    ap.add_argument("--include-unapproved", action="store_true")
    ap.add_argument("--before-b-only", action="store_true")
    ap.add_argument(
        "--prod",
        action="store_true",
        help="本番書き込み（永続化）。社長『やってよい』後のみ",
    )
    ap.add_argument(
        "--i-confirm-prod",
        action="store_true",
        help="--prod 時の二重確認",
    )
    ap.add_argument(
        "--backfill-a-logs",
        action="store_true",
        help="_work/lane_a_send_*.json の本番成功をマスタ／施策へ反映",
    )
    args = ap.parse_args(argv)

    local = HERE / "config.local.json"
    deals_cfg = load_config(
        args.config or (local if local.is_file() else HERE / "config.example.json")
    )

    if args.backfill_a_logs:
        from a_track import backfill_from_work

        sid = str(deals_cfg.get("ads_spreadsheet_id") or "").strip()
        svc = sheets_service(write=True)
        results = backfill_from_work(svc, sid)
        print(json.dumps({"backfill": len(results), "items": results}, ensure_ascii=False))
        return 0

    if not str(args.sku or "").strip():
        LOG.error("--sku が必要です（または --backfill-a-logs）")
        return 1
    rows = load_a_rows(
        deals_cfg,
        sku_filter=args.sku,
        include_unapproved=bool(args.include_unapproved),
    )
    if args.before_b_only:
        rows = filter_before_b(deals_cfg, rows)
    if not rows:
        LOG.error("対象A行なし（sku / 期間 / before-b を確認）")
        return 1
    if len(rows) > 1:
        LOG.error("§10.8は1SKU・1期間。複数ヒット: %s", len(rows))
        return 1

    r = rows[0]
    start = parse_ymd(r.get("開始日"))
    end = parse_ymd(r.get("終了日"))
    sale = _float(r.get("セール価格")) or _float(r.get("タイムセール価格_確定"))
    our = _float(r.get("通常価格")) or _float(r.get("出品者価格_SC"))
    assert start and end and sale

    spapi = resolve_spapi_cfg(deals_cfg)
    marketplace_id = str(spapi.get("marketplace_id") or MARKETPLACE_JP)
    body = build_discounted_price_patch(
        marketplace_id=marketplace_id,
        currency=str(r.get("通貨") or "JPY"),
        sale_price=sale,
        start=start,
        end=end,
        our_price=our,
    )

    validation_preview = not bool(args.prod)
    if args.prod:
        if not args.i_confirm_prod:
            LOG.error("--prod には --i-confirm-prod も必要です")
            return 1
        if not bool(spapi.get("allow_prod")):
            LOG.error("spapi config の allow_prod=true が必要です")
            return 1
        LOG.warning("PROD 書き込みを実行します sku=%s %s..%s", args.sku, start, end)

    token = lwa_token(spapi)
    seller_id = str(spapi.get("seller_id") or "").strip()
    endpoint = str(spapi.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    ua = str(spapi.get("user_agent") or "OctasAmazonDealsLaneA/0.1")
    sku = str(r.get("SKU") or "").strip()

    LOG.info(
        "PATCH sku=%s preview=%s period=%s..%s price=%s",
        sku,
        validation_preview,
        start,
        end,
        sale,
    )
    resp = patch_listings_item(
        endpoint=endpoint,
        access_token=token,
        seller_id=seller_id,
        sku=sku,
        marketplace_id=marketplace_id,
        body=body,
        user_agent=ua,
        validation_preview=validation_preview,
    )
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    out = {
        "stamp": stamp,
        "sku": sku,
        "asin": r.get("ASIN"),
        "validation_preview": validation_preview,
        "http_status": resp.status_code,
        "request": body,
        "response": None,
    }
    try:
        out["response"] = resp.json()
    except Exception:
        out["response_text"] = resp.text[:4000]
    path = out_dir / f"lane_a_send_{stamp}.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info("HTTP %s → %s", resp.status_code, path)

    # マスタ／施策へ A実施フラグ＋ログ参照
    try:
        from a_track import record_lane_a_send  # noqa: WPS433

        sid = str(deals_cfg.get("ads_spreadsheet_id") or "").strip()
        if sid and resp.status_code < 400:
            svc_w = sheets_service(write=True)
            info = record_lane_a_send(
                svc_w,
                sid,
                sku=sku,
                asin=str(r.get("ASIN") or ""),
                start=start.isoformat(),
                end=end.isoformat(),
                price=sale,
                log_name=path.name,
                prod=bool(args.prod),
                http_status=resp.status_code,
            )
            out["sheet_track"] = info
            path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
            LOG.info("sheet track %s", info)
    except Exception as e:
        LOG.warning("A実施フラグ書き込みスキップ: %s", e)

    print(json.dumps({"http": resp.status_code, "preview": validation_preview, "out": str(path)}, ensure_ascii=False))
    if resp.status_code >= 400:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
