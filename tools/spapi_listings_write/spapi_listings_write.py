#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SP-API Listings Items 書込 v1.1（既存 ASIN・offer only・1行 or CSV複数行）

- dry_run: putListingsItem ?mode=VALIDATION_PREVIEW（永続化しない）
- prod: 実送信（config.allow_prod=true 必須）
- items_csv があれば複数行（max_items 上限）。無ければ config の sku/asin 1件

正本: docs/org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md
承認: docs/org/LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md
公式: https://developer-docs.amazon.com/sp-api/docs/submit-listings-data
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import quote, urlencode

LOG = logging.getLogger("spapi_listings_write")
SCRIPT_DIR = Path(__file__).resolve().parent

try:
    import requests
except ImportError as e:  # pragma: no cover
    raise SystemExit("requests が必要です: pip install -r requirements.txt") from e

LWA_TOKEN_URL = "https://api.amazon.com/auth/o2/token"
ASIN_RE = re.compile(r"^B0[A-Z0-9]{8}$", re.I)


def _utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")


def _load_json(path: Path) -> Dict[str, Any]:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def _save_json(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
        f.write("\n")


def resolve_cred(cfg: Dict[str, Any], key: str, env_name: str) -> str:
    v = str(os.environ.get(env_name) or cfg.get(key) or "").strip()
    if not v or v.startswith("REPLACE"):
        raise SystemExit("未設定: config.%s または環境変数 %s" % (key, env_name))
    return v


def merge_auth_from_path(cfg: Dict[str, Any], base: Path) -> Dict[str, Any]:
    """auth_config_path があれば LWA 3点を上書きマージ（seller_id 等は本 config 優先）。"""
    rel = str(cfg.get("auth_config_path") or "").strip()
    if not rel:
        return cfg
    auth_path = Path(rel)
    if not auth_path.is_absolute():
        auth_path = (base / auth_path).resolve()
    if not auth_path.is_file():
        LOG.warning("auth_config_path がありません（無視）: %s", auth_path)
        return cfg
    auth = _load_json(auth_path)
    out = dict(cfg)
    for k in ("lwa_client_id", "lwa_client_secret", "refresh_token"):
        if auth.get(k) and not str(out.get(k) or "").startswith("REPLACE"):
            continue
        if auth.get(k) and not str(auth.get(k) or "").startswith("REPLACE"):
            out[k] = auth[k]
    for k in ("lwa_client_id", "lwa_client_secret", "refresh_token"):
        if str(out.get(k) or "").startswith("REPLACE") and auth.get(k):
            out[k] = auth[k]
    LOG.info("認証を読み込み: %s", auth_path)
    return out


def request_lwa_access_token(client_id: str, client_secret: str, refresh_token: str) -> Dict[str, Any]:
    resp = requests.post(
        LWA_TOKEN_URL,
        data={
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "client_id": client_id,
            "client_secret": client_secret,
        },
        headers={"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"},
        timeout=60,
    )
    if resp.status_code != 200:
        raise RuntimeError("LWA 失敗 HTTP %s: %s" % (resp.status_code, resp.text[:500]))
    body = resp.json()
    if not body.get("access_token"):
        raise RuntimeError("access_token なし")
    return body


def build_offer_body(
    marketplace_id: str,
    asin: str,
    price: float,
    quantity: int,
    condition_type: str,
    fulfillment_channel_code: str,
    currency: str,
) -> Dict[str, Any]:
    return {
        "productType": "PRODUCT",
        "requirements": "LISTING_OFFER_ONLY",
        "attributes": {
            "condition_type": [
                {"value": condition_type, "marketplace_id": marketplace_id}
            ],
            "merchant_suggested_asin": [
                {"value": asin, "marketplace_id": marketplace_id}
            ],
            "purchasable_offer": [
                {
                    "currency": currency,
                    "our_price": [{"schedule": [{"value_with_tax": price}]}],
                    "marketplace_id": marketplace_id,
                }
            ],
            "fulfillment_availability": [
                {
                    "fulfillment_channel_code": fulfillment_channel_code,
                    "quantity": quantity,
                    "marketplace_id": marketplace_id,
                }
            ],
        },
    }


def validate_item(sku: str, asin: str, price: float, quantity: int) -> Optional[str]:
    if not sku:
        return "sku が空です"
    if not ASIN_RE.match(asin):
        return "ASIN 不正: %r" % asin
    if price <= 0:
        return "price は正の数が必要です"
    if quantity < 0:
        return "quantity は 0 以上"
    return None


def spapi_headers(endpoint: str, access_token: str, user_agent: str) -> Dict[str, str]:
    host = endpoint.replace("https://", "").replace("http://", "").split("/")[0]
    return {
        "host": host,
        "x-amz-access-token": access_token,
        "x-amz-date": datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ"),
        "user-agent": user_agent,
        "accept": "application/json",
        "content-type": "application/json",
    }


def get_listings_item(
    endpoint: str,
    access_token: str,
    seller_id: str,
    sku: str,
    marketplace_id: str,
    user_agent: str,
) -> requests.Response:
    path = "/listings/2021-08-01/items/%s/%s" % (
        quote(seller_id, safe="-_.~"),
        quote(sku, safe="-_.~"),
    )
    qs = urlencode(
        {
            "marketplaceIds": marketplace_id,
            "includedData": "summaries,attributes,issues",
        }
    )
    url = "%s%s?%s" % (endpoint.rstrip("/"), path, qs)
    return requests.get(
        url, headers=spapi_headers(endpoint, access_token, user_agent), timeout=60
    )


def put_listings_item(
    endpoint: str,
    access_token: str,
    seller_id: str,
    sku: str,
    marketplace_id: str,
    body: Dict[str, Any],
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
    return requests.put(
        url,
        headers=spapi_headers(endpoint, access_token, user_agent),
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        timeout=90,
    )


def load_items(cfg: Dict[str, Any], base: Path, max_items: int) -> List[Dict[str, Any]]:
    """items_csv 優先。無ければ config の単一行。"""
    rel = str(cfg.get("items_csv") or "").strip()
    if rel:
        csv_path = Path(rel)
        if not csv_path.is_absolute():
            csv_path = (base / csv_path).resolve()
        if not csv_path.is_file():
            raise SystemExit("items_csv がありません: %s" % csv_path)
        rows: List[Dict[str, Any]] = []
        with csv_path.open(encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            if not reader.fieldnames:
                raise SystemExit("items_csv にヘッダがありません")
            for i, raw in enumerate(reader, start=2):
                sku = str(raw.get("sku") or "").strip()
                asin = str(raw.get("asin") or "").strip()
                if not sku and not asin:
                    continue
                try:
                    price = float(str(raw.get("price") or "").strip())
                    quantity = int(float(str(raw.get("quantity") or "0").strip() or "0"))
                except ValueError:
                    raise SystemExit("items_csv 行%d: price/quantity 不正" % i)
                note = str(raw.get("note") or "").strip()
                rows.append(
                    {
                        "sku": sku,
                        "asin": asin,
                        "price": price,
                        "quantity": quantity,
                        "note": note,
                        "sourceLine": i,
                    }
                )
        if not rows:
            raise SystemExit("items_csv に有効行がありません: %s" % csv_path)
        if len(rows) > max_items:
            raise SystemExit(
                "items_csv が max_items=%d を超えています（%d行）。分割するか max_items を見直す"
                % (max_items, len(rows))
            )
        LOG.info("items_csv 読込: %s rows=%d", csv_path, len(rows))
        return rows

    sku = str(cfg.get("sku") or "").strip()
    asin = str(cfg.get("asin") or "").strip()
    price = float(cfg.get("price"))
    quantity = int(cfg.get("quantity"))
    return [
        {
            "sku": sku,
            "asin": asin,
            "price": price,
            "quantity": quantity,
            "note": "config.single",
            "sourceLine": 0,
        }
    ]


def process_one_item(
    *,
    item: Dict[str, Any],
    access_token: str,
    seller_id: str,
    marketplace_id: str,
    endpoint: str,
    user_agent: str,
    condition_type: str,
    fulfillment_channel_code: str,
    currency: str,
    mode: str,
) -> Dict[str, Any]:
    sku = str(item["sku"])
    asin = str(item["asin"])
    price = float(item["price"])
    quantity = int(item["quantity"])
    note = str(item.get("note") or "")

    result: Dict[str, Any] = {
        "sku": sku,
        "asin": asin,
        "price": price,
        "quantity": quantity,
        "note": note,
        "ok": False,
        "exitHint": "",
    }

    err = validate_item(sku, asin, price, quantity)
    if err:
        result["exitHint"] = err
        LOG.error("[%s] 入力不正: %s", sku, err)
        return result

    body = build_offer_body(
        marketplace_id,
        asin,
        price,
        quantity,
        condition_type,
        fulfillment_channel_code,
        currency,
    )
    result["requestBody"] = body

    LOG.info("GET sku=%s asin=%s…", sku, asin)
    get_resp = get_listings_item(
        endpoint, access_token, seller_id, sku, marketplace_id, user_agent
    )
    result["getHttpStatus"] = get_resp.status_code
    result["getBodyPreview"] = get_resp.text[:400]
    LOG.info("GET HTTP %s sku=%s", get_resp.status_code, sku)

    validation_preview = mode == "dry_run"
    LOG.info(
        "PUT sku=%s mode=%s validationPreview=%s…",
        sku,
        mode,
        validation_preview,
    )
    put_resp = put_listings_item(
        endpoint,
        access_token,
        seller_id,
        sku,
        marketplace_id,
        body,
        user_agent,
        validation_preview=validation_preview,
    )
    result["putHttpStatus"] = put_resp.status_code
    result["validationPreview"] = validation_preview
    result["putBodyPreview"] = put_resp.text[:800]
    LOG.info("PUT HTTP %s sku=%s", put_resp.status_code, sku)

    if put_resp.status_code not in (200, 202):
        result["exitHint"] = "PUT HTTP %s" % put_resp.status_code
        LOG.error("[%s] 書込失敗: %s", sku, put_resp.text[:400])
        return result

    try:
        put_json = put_resp.json()
        status = put_json.get("status") or put_json.get("submissionId")
        issues = put_json.get("issues") or []
        result["status"] = status
        result["issueCount"] = len(issues)
        result["issuesPreview"] = [
            {
                "code": i.get("code"),
                "severity": i.get("severity"),
                "message": (i.get("message") or "")[:200],
            }
            for i in issues[:5]
        ]
        LOG.info(
            "[%s] status=%s issues=%d",
            sku,
            status,
            len(issues),
        )
        for iss in issues[:5]:
            LOG.warning(
                "[%s] issue code=%s severity=%s message=%s",
                sku,
                iss.get("code"),
                iss.get("severity"),
                (iss.get("message") or "")[:200],
            )
        if any(str(i.get("severity") or "").upper() == "ERROR" for i in issues):
            result["exitHint"] = "ERROR issues"
            return result
    except Exception as e:
        result["exitHint"] = "JSON parse: %s" % e
        return result

    result["ok"] = True
    result["exitHint"] = "ok"
    return result


def run(config_path: Path, mode_override: Optional[str]) -> int:
    cfg = _load_json(config_path)
    cfg = merge_auth_from_path(cfg, config_path.parent)

    mode = (mode_override or cfg.get("mode") or "dry_run").strip().lower()
    if mode not in ("dry_run", "prod"):
        raise SystemExit("mode は dry_run または prod")

    allow_prod = bool(cfg.get("allow_prod"))
    if mode == "prod" and not allow_prod:
        raise SystemExit(
            "prod には config.allow_prod=true が必要です（誤送信防止）"
        )

    max_items = int(cfg.get("max_items") or 5)
    if max_items < 1 or max_items > 50:
        raise SystemExit("max_items は 1〜50")

    client_id = resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID")
    client_secret = resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET")
    refresh_token = resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN")
    seller_id = resolve_cred(cfg, "seller_id", "SPAPI_SELLER_ID")

    marketplace_id = str(
        os.environ.get("SPAPI_MARKETPLACE_ID")
        or cfg.get("marketplace_id")
        or "A1VC38T7YXB528"
    ).strip()
    endpoint = str(
        os.environ.get("SPAPI_ENDPOINT")
        or cfg.get("endpoint")
        or "https://sellingpartnerapi-fe.amazon.com"
    ).strip()
    condition_type = str(cfg.get("condition_type") or "new_new").strip()
    fulfillment_channel_code = str(
        cfg.get("fulfillment_channel_code") or "DEFAULT"
    ).strip()
    currency = str(cfg.get("currency") or "JPY").strip()
    user_agent = str(
        cfg.get("user_agent") or "OctasSpapiListingsWrite/1.1 (Language=Python)"
    ).strip()

    items = load_items(cfg, config_path.parent, max_items)

    run_id = "SPAPI_LISTINGS_WRITE_%s" % _utc_stamp()
    report: Dict[str, Any] = {
        "runId": run_id,
        "version": "spapi-listings-write-v1.1",
        "mode": mode,
        "itemCount": len(items),
        "maxItems": max_items,
        "marketplaceId": marketplace_id,
        "endpoint": endpoint,
        "sellerIdSuffix": seller_id[-4:] if len(seller_id) >= 4 else "****",
        "steps": {},
        "results": [],
    }

    LOG.info("Step1 LWA…")
    token_body = request_lwa_access_token(client_id, client_secret, refresh_token)
    access_token = str(token_body["access_token"])
    report["steps"]["lwa"] = {
        "ok": True,
        "expires_in": token_body.get("expires_in"),
    }
    LOG.info("LWA OK items=%d mode=%s", len(items), mode)

    ok_n = 0
    fail_n = 0
    for idx, item in enumerate(items, start=1):
        LOG.info("--- item %d/%d ---", idx, len(items))
        one = process_one_item(
            item=item,
            access_token=access_token,
            seller_id=seller_id,
            marketplace_id=marketplace_id,
            endpoint=endpoint,
            user_agent=user_agent,
            condition_type=condition_type,
            fulfillment_channel_code=fulfillment_channel_code,
            currency=currency,
            mode=mode,
        )
        report["results"].append(one)
        if one.get("ok"):
            ok_n += 1
        else:
            fail_n += 1

    report["summary"] = {"ok": ok_n, "fail": fail_n, "total": len(items)}
    out_dir = SCRIPT_DIR / "out"
    report_path = out_dir / ("%s_REPORT.json" % run_id)
    _save_json(report_path, report)
    LOG.info("レポート: %s", report_path)
    LOG.info("%s 完了 ok=%d fail=%d total=%d", mode, ok_n, fail_n, len(items))

    if fail_n:
        return 1 if ok_n else 2
    return 0


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(
        description="SP-API Listings 書込 v1.1（CSV複数行可・offer only・dry_run/prod）"
    )
    parser.add_argument(
        "--config",
        default=str(SCRIPT_DIR / "config.local.json"),
    )
    parser.add_argument("--mode", choices=["dry_run", "prod"], default=None)
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
    return run(config_path, args.mode)


if __name__ == "__main__":
    sys.exit(main())
