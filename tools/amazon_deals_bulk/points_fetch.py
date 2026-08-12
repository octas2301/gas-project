# -*- coding: utf-8 -*-
"""
出品者ポイント付与率の取得 → 出品者ポイント現在% スナップ（§10.10 / P0-G9）。

手段:
  1) Listings Items GET の offers[].points.pointsNumber と price から % を復元（本線）
  2) だめなら Feeds GET_FLAT_FILE_OFFER_POINTS_PREFERENCE_DATA（環境によっては 400）

--write は 出品者ポイント現在%／円のみ更新。最終終着%（セール前列）は埋めない。

例:
  python points_fetch.py
  python points_fetch.py --write
  python points_fetch.py --sku "..." --write
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import logging
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote, urlencode

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

LOG = logging.getLogger("amazon_deals_bulk.points_fetch")

try:
    import requests
except ImportError:
    raise SystemExit("requests が必要です") from None

from lane_a_send import lwa_token, resolve_spapi_cfg  # noqa: E402
from paths import load_config  # noqa: E402
from points_logic import parse_percent  # noqa: E402
from sheet_schema import (  # noqa: E402
    MASTER_SHEET,
    POINT_CURRENT_COL,
    POINT_CURRENT_YEN_COL,
)
from sheets_io import read_sheet_rows, sheets_service, update_row_fields  # noqa: E402

GET_FEED_TYPE = "GET_FLAT_FILE_OFFER_POINTS_PREFERENCE_DATA"
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


def get_listings_item(
    *,
    endpoint: str,
    access_token: str,
    user_agent: str,
    seller_id: str,
    sku: str,
    marketplace_id: str,
) -> dict:
    path = "/listings/2021-08-01/items/%s/%s" % (
        quote(seller_id, safe="-_.~"),
        quote(sku, safe="-_.~"),
    )
    q = urlencode(
        {
            "marketplaceIds": marketplace_id,
            "includedData": "attributes,offers",
        }
    )
    url = "%s%s?%s" % (endpoint.rstrip("/"), path, q)
    resp = requests.get(url, headers=_spapi_headers(endpoint, access_token, user_agent), timeout=90)
    if resp.status_code >= 300:
        raise RuntimeError("listings GET HTTP %s %s" % (resp.status_code, resp.text[:400]))
    return resp.json()


def _walk_for_points_percent(obj: Any, found: List[int]) -> None:
    if isinstance(obj, dict):
        for k, v in obj.items():
            lk = str(k).lower()
            if "point" in lk and ("percent" in lk or "percentage" in lk or lk.endswith("pct")):
                p = parse_percent(v, default=None)
                if p is not None:
                    found.append(p)
            _walk_for_points_percent(v, found)
    elif isinstance(obj, list):
        for x in obj:
            _walk_for_points_percent(x, found)


def points_from_listings_offer(body: dict) -> Tuple[Optional[int], Optional[int]]:
    """
    Listings GET から (出品者ポイント%, ポイント数) を返す。

    JPでは offers[].points.pointsNumber のみ返ることが多く、
    % = round(pointsNumber * 100 / price) で復元する（P0-G9）。
    B2C を優先。
    """
    offers = body.get("offers")
    if not isinstance(offers, list):
        return None, None
    ordered = sorted(
        offers,
        key=lambda o: 0 if str((o or {}).get("offerType") or "") == "B2C" else 1,
    )
    for o in ordered:
        if not isinstance(o, dict):
            continue
        pts = o.get("points") if isinstance(o.get("points"), dict) else {}
        raw_n = pts.get("pointsNumber")
        try:
            n = int(float(raw_n)) if raw_n is not None and str(raw_n).strip() != "" else None
        except (TypeError, ValueError):
            n = None
        price_obj = o.get("price") if isinstance(o.get("price"), dict) else {}
        raw_p = price_obj.get("amount")
        try:
            price = float(raw_p) if raw_p is not None and str(raw_p).strip() != "" else None
        except (TypeError, ValueError):
            price = None
        if n is not None and price and price > 0:
            pct = int(round(n * 100.0 / price))
            return max(0, pct), n
        if n is not None and (price is None or price <= 0):
            # 価格なしでも円は取れる
            return None, n
    found: List[int] = []
    _walk_for_points_percent(body, found)
    if found:
        return found[0], None
    return None, None


def percent_from_listings(body: dict) -> Optional[int]:
    pct, _yen = points_from_listings_offer(body)
    return pct


def create_feed_document(*, endpoint: str, access_token: str, user_agent: str) -> dict:
    url = "%s/feeds/2021-06-30/documents" % endpoint.rstrip("/")
    resp = requests.post(
        url,
        headers={**_spapi_headers(endpoint, access_token, user_agent), "content-type": "application/json"},
        data=json.dumps({"contentType": CONTENT_TYPE}),
        timeout=60,
    )
    if resp.status_code >= 300:
        raise RuntimeError("createFeedDocument HTTP %s %s" % (resp.status_code, resp.text[:400]))
    return resp.json()


def upload_empty_tsv(upload_url: str) -> None:
    # GET 系は空／ヘッダのみで要求する実装が多い
    body = b"sku\tpoints_percent\n"
    resp = requests.put(upload_url, data=body, headers={"Content-Type": CONTENT_TYPE}, timeout=120)
    if resp.status_code >= 300:
        raise RuntimeError("upload HTTP %s %s" % (resp.status_code, resp.text[:400]))


def create_feed(
    *,
    endpoint: str,
    access_token: str,
    user_agent: str,
    document_id: str,
    marketplace_id: str,
    feed_type: str,
) -> dict:
    url = "%s/feeds/2021-06-30/feeds" % endpoint.rstrip("/")
    payload = {
        "feedType": feed_type,
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
        raise RuntimeError("createFeed HTTP %s %s" % (resp.status_code, resp.text[:400]))
    return resp.json()


def get_feed(*, endpoint: str, access_token: str, user_agent: str, feed_id: str) -> dict:
    url = "%s/feeds/2021-06-30/feeds/%s" % (endpoint.rstrip("/"), feed_id)
    resp = requests.get(url, headers=_spapi_headers(endpoint, access_token, user_agent), timeout=60)
    if resp.status_code >= 300:
        raise RuntimeError("getFeed HTTP %s %s" % (resp.status_code, resp.text[:400]))
    return resp.json()


def wait_feed(*, endpoint: str, access_token: str, user_agent: str, feed_id: str, timeout_sec: int = 240) -> dict:
    deadline = time.time() + timeout_sec
    last: dict = {}
    while time.time() < deadline:
        last = get_feed(endpoint=endpoint, access_token=access_token, user_agent=user_agent, feed_id=feed_id)
        st = str(last.get("processingStatus") or "")
        LOG.info("feed %s status=%s", feed_id, st)
        if st in ("DONE", "CANCELLED", "FATAL"):
            return last
        time.sleep(8)
    return last


def get_feed_document(*, endpoint: str, access_token: str, user_agent: str, document_id: str) -> dict:
    url = "%s/feeds/2021-06-30/documents/%s" % (endpoint.rstrip("/"), document_id)
    resp = requests.get(url, headers=_spapi_headers(endpoint, access_token, user_agent), timeout=60)
    if resp.status_code >= 300:
        raise RuntimeError("getFeedDocument HTTP %s %s" % (resp.status_code, resp.text[:400]))
    return resp.json()


def download_text(url: str) -> str:
    resp = requests.get(url, timeout=120)
    if resp.status_code >= 300:
        raise RuntimeError("download HTTP %s" % resp.status_code)
    # JP TSV は Shift_JIS のことがある
    for enc in ("utf-8-sig", "utf-8", "cp932"):
        try:
            return resp.content.decode(enc)
        except UnicodeDecodeError:
            continue
    return resp.content.decode("utf-8", errors="replace")


def parse_points_tsv(text: str) -> Dict[str, int]:
    out: Dict[str, int] = {}
    reader = csv.DictReader(io.StringIO(text), delimiter="\t")
    if not reader.fieldnames:
        return out
    fields = [str(f or "").strip().lower() for f in reader.fieldnames]
    # normalize keys
    for row in reader:
        raw = {str(k or "").strip().lower(): v for k, v in row.items()}
        sku = str(raw.get("sku") or raw.get("seller-sku") or "").strip()
        pct = parse_percent(
            raw.get("points_percent")
            or raw.get("point_percent")
            or raw.get("points-percent"),
            default=None,
        )
        if sku and pct is not None:
            out[sku] = pct
    return out


def fetch_map_via_get_feed(
    *,
    endpoint: str,
    access_token: str,
    user_agent: str,
    marketplace_id: str,
) -> Dict[str, int]:
    doc = create_feed_document(endpoint=endpoint, access_token=access_token, user_agent=user_agent)
    doc_id = str(doc.get("feedDocumentId") or "")
    upload_url = str(doc.get("url") or "")
    upload_empty_tsv(upload_url)
    created = create_feed(
        endpoint=endpoint,
        access_token=access_token,
        user_agent=user_agent,
        document_id=doc_id,
        marketplace_id=marketplace_id,
        feed_type=GET_FEED_TYPE,
    )
    feed_id = str(created.get("feedId") or "")
    fed = wait_feed(endpoint=endpoint, access_token=access_token, user_agent=user_agent, feed_id=feed_id)
    result_id = str(fed.get("resultFeedDocumentId") or "")
    if not result_id:
        LOG.warning("GET points feed に resultFeedDocumentId なし: %s", fed)
        return {}
    meta = get_feed_document(
        endpoint=endpoint, access_token=access_token, user_agent=user_agent, document_id=result_id
    )
    text = download_text(str(meta.get("url") or ""))
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    (out_dir / ("points_get_%s.tsv" % stamp)).write_text(text, encoding="utf-8")
    return parse_points_tsv(text)


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="出品者ポイント取得→現在%スナップ")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--sku", type=str, default="")
    ap.add_argument("--write", action="store_true", help="出品者ポイント現在%／円だけシートへ書く")
    ap.add_argument(
        "--method",
        choices=("auto", "listings", "feed"),
        default="auto",
        help="auto=listings→feed",
    )
    args = ap.parse_args(argv)

    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))
    spapi = resolve_spapi_cfg(cfg)
    token = lwa_token(spapi)
    endpoint = str(spapi.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    ua = str(spapi.get("user_agent") or "OctasAmazonDealsPointsFetch/0.1")
    seller_id = str(spapi.get("seller_id") or "").strip()
    marketplace_id = str(spapi.get("marketplace_id") or cfg.get("marketplace_id") or MARKETPLACE_JP)

    svc = sheets_service(write=bool(args.write))
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    _h, rows = read_sheet_rows(svc, sid, MASTER_SHEET)
    want_sku = (args.sku or "").strip()
    targets = []
    for r in rows:
        sku = str(r.get("SKU") or "").strip()
        if not sku:
            continue
        if want_sku and sku != want_sku:
            continue
        targets.append(r)
    if not targets:
        LOG.info("対象SKUなし")
        return 0

    fetched: Dict[str, int] = {}
    fetched_yen: Dict[str, int] = {}
    method_used = args.method

    if args.method in ("auto", "listings") and seller_id:
        for r in targets:
            sku = str(r.get("SKU") or "").strip()
            try:
                body = get_listings_item(
                    endpoint=endpoint,
                    access_token=token,
                    user_agent=ua,
                    seller_id=seller_id,
                    sku=sku,
                    marketplace_id=marketplace_id,
                )
                pct, yen = points_from_listings_offer(body)
                if pct is not None:
                    fetched[sku] = pct
                    if yen is not None:
                        fetched_yen[sku] = yen
                    LOG.info("listings %s → %s%% (pointsNumber=%s)", sku, pct, yen)
                else:
                    LOG.info("listings %s: points から%%復元できず yen=%s", sku, yen)
            except Exception as e:
                LOG.warning("listings失敗 %s: %s", sku, e)

    if args.method == "feed" or (args.method == "auto" and len(fetched) < len(targets)):
        try:
            feed_map = fetch_map_via_get_feed(
                endpoint=endpoint,
                access_token=token,
                user_agent=ua,
                marketplace_id=marketplace_id,
            )
            method_used = "feed" if feed_map else method_used
            for sku, pct in feed_map.items():
                if want_sku and sku != want_sku:
                    continue
                fetched.setdefault(sku, pct)
            LOG.info("GET feed map size=%s", len(feed_map))
        except Exception as e:
            LOG.warning(
                "GET points feed 失敗（listings復元を本線とする）: %s",
                e,
            )

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    meta = {
        "stamp": stamp,
        "method": method_used,
        "fetched": fetched,
        "fetched_yen": fetched_yen,
        "targets": [str(r.get("SKU") or "") for r in targets],
        "note": "listings: pointsNumber/price→%%。GET feed は環境により不可",
    }
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    meta_path = out_dir / ("points_fetch_%s.json" % stamp)
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    filled = 0
    if args.write:
        row_index: Dict[str, int] = {}
        for i, r in enumerate(rows):
            sku = str(r.get("SKU") or "").strip()
            if sku and sku not in row_index:
                row_index[sku] = i + 2
        for sku, pct in fetched.items():
            ridx = row_index.get(sku)
            if not ridx:
                continue
            fields: Dict[str, Any] = {POINT_CURRENT_COL: pct}
            if sku in fetched_yen:
                fields[POINT_CURRENT_YEN_COL] = fetched_yen[sku]
            update_row_fields(svc, sid, MASTER_SHEET, ridx, fields)
            filled += 1
        LOG.info("出品者現在%%を同期した件数=%s（最終終着%%はfetchしない）", filled)

    print(
        json.dumps(
            {
                "fetched": len(fetched),
                "filled": filled,
                "write": bool(args.write),
                "meta": str(meta_path),
                "sample": dict(list(fetched.items())[:5]),
            },
            ensure_ascii=False,
        )
    )
    if args.write:
        return 0 if filled or not fetched else 1
    return 0 if fetched or not targets else 1


if __name__ == "__main__":
    raise SystemExit(main())
