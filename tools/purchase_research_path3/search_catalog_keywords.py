#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""経路3 PoC: SP-API Catalog Items キーワード検索（読取のみ）。出品書込なし。

認証は tools/spapi_smoke/config.local.json（gitignore）または環境変数。
PA-API は使わない。HTML スクレイプはしない。
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlencode

SCRIPT_DIR = Path(__file__).resolve().parent
SMOKE_DIR = SCRIPT_DIR.parent / "spapi_smoke"
sys.path.insert(0, str(SMOKE_DIR))

try:
    import requests
except ImportError as e:  # pragma: no cover
    raise SystemExit("requests が必要です: pip install -r ../spapi_smoke/requirements.txt") from e

from spapi_smoke import (  # noqa: E402
    _spapi_headers,
    resolve_cred,
    request_lwa_access_token,
    _load_config,
)


def search_catalog_items(
    endpoint: str,
    access_token: str,
    marketplace_id: str,
    user_agent: str,
    keywords: str,
    page_size: int = 20,
    page_token: str = "",
    brand_names: str = "",
) -> requests.Response:
    path = "/catalog/2022-04-01/items"
    params = {
        "marketplaceIds": marketplace_id,
        "keywords": keywords,
        "includedData": "summaries",
        "pageSize": str(max(1, min(20, page_size))),
    }
    if page_token:
        params["pageToken"] = page_token
    if brand_names:
        params["brandNames"] = brand_names
    url = "%s%s?%s" % (endpoint.rstrip("/"), path, urlencode(params))
    return requests.get(
        url,
        headers=_spapi_headers(endpoint, access_token, user_agent),
        timeout=60,
    )


def asins_from_body(body: Any) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    if not isinstance(body, dict):
        return rows
    items = body.get("items") or []
    if not isinstance(items, list):
        return rows
    for it in items:
        if not isinstance(it, dict):
            continue
        asin = str(it.get("asin") or "").strip().upper()
        title = ""
        brand = ""
        sums = it.get("summaries") or []
        if isinstance(sums, list) and sums and isinstance(sums[0], dict):
            title = str(sums[0].get("itemName") or "").strip()
            brand = str(sums[0].get("brand") or "").strip()
            mfr = str(sums[0].get("manufacturer") or "").strip()
        else:
            mfr = ""
        if asin:
            rows.append({"asin": asin, "title": title, "brand": brand, "manufacturer": mfr})
    return rows


def main() -> int:
    p = argparse.ArgumentParser(description="Path3 SP-API keyword search (read-only)")
    p.add_argument("--keywords", required=True)
    p.add_argument("--brand-names", default="", help="Catalog の brandNames（任意）")
    p.add_argument("--max-pages", type=int, default=10, help="pageSize20×この回数")
    p.add_argument("--config", default="", help="省略時 spapi_smoke/config.local.json")
    p.add_argument("--out-dir", default="")
    args = p.parse_args()

    cfg_path = Path(args.config) if args.config else SMOKE_DIR / "config.local.json"
    if not cfg_path.is_file():
        print("config がありません: %s" % cfg_path)
        print("人手の amazon_asins.txt で diff_keepa_vs_amazon.py を先に回してください。")
        return 2

    cfg = _load_config(cfg_path)
    client_id = resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID")
    client_secret = resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET")
    refresh = resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN")
    marketplace_id = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    user_agent = str(cfg.get("user_agent") or "OctasPath3Poc/1.0 (Language=Python)")

    token_body = request_lwa_access_token(client_id, client_secret, refresh)
    access = str(token_body.get("access_token") or "")

    rows: List[Dict[str, str]] = []
    seen: set[str] = set()
    page_token = ""
    last_http = 0
    pages = 0
    last_err = ""
    max_pages = max(1, min(int(args.max_pages), 25))
    while pages < max_pages:
        resp = search_catalog_items(
            endpoint,
            access,
            marketplace_id,
            user_agent,
            args.keywords,
            page_token=page_token,
            brand_names=str(args.brand_names or "").strip(),
        )
        last_http = resp.status_code
        pages += 1
        try:
            body = resp.json()
        except Exception:
            body = {"_raw": (resp.text or "")[:400]}
        if resp.status_code != 200:
            last_err = str(body)[:400]
            break
        for r in asins_from_body(body):
            if r["asin"] in seen:
                continue
            seen.add(r["asin"])
            rows.append(r)
        nxt = ""
        if isinstance(body, dict):
            pag = body.get("pagination") or {}
            if isinstance(pag, dict):
                nxt = str(pag.get("nextToken") or "").strip()
        if not nxt:
            break
        page_token = nxt

    needle = str(args.keywords or "")
    brand_hit = 0
    for r in rows:
        blob = (r.get("brand") or "") + (r.get("manufacturer") or "") + (r.get("title") or "")
        if needle and needle in blob:
            brand_hit += 1

    report = {
        "at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "keywords": args.keywords,
        "brandNames": str(args.brand_names or ""),
        "http": last_http,
        "pages": pages,
        "count": len(rows),
        "title_or_brand_contains_keywords": brand_hit,
        "items": rows,
        "note": "読取のみ。トークンは書かない。pageSize最大20。Keepa集合Aとの差分はこのJSONだけでは出ない。",
    }
    out_dir = Path(args.out_dir) if args.out_dir else SCRIPT_DIR / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    dest = out_dir / ("PATH3_SPAPI_%s.json" % stamp)
    dest.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    asins_txt = out_dir / ("PATH3_SPAPI_%s_asins.txt" % stamp)
    asins_txt.write_text("\n".join(r["asin"] for r in rows) + ("\n" if rows else ""), encoding="utf-8")

    print("HTTP %s pages=%s count=%s keyword_in_title_brand=%s" % (
        last_http, pages, len(rows), brand_hit
    ))
    print("report=%s" % dest)
    print("asins=%s" % asins_txt)
    if last_http != 200:
        print("error_preview=%s" % last_err)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
