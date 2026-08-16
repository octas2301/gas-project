# -*- coding: utf-8 -*-
"""Amazon inventory GET for exactly one sellerSku. Never writes listing master. Default is no HTTP."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from urllib.parse import quote

SCRIPT_DIR = Path(__file__).resolve().parent
SMOKE = SCRIPT_DIR.parent / "spapi_smoke"
sys.path.insert(0, str(SMOKE))

from inventory import qty_by_seller_sku, qty_from_listings  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--sku", required=True, help="sellerSku exactly one")
    ap.add_argument("--live", action="store_true", help="Actually GET SP-API")
    ap.add_argument("--config", default="", help="spapi_smoke/config.local.json")
    args = ap.parse_args()
    sku = str(args.sku).strip().split()[0]
    if not sku:
        print("empty sku")
        return 2
    if not args.live:
        print("dry-run sku=%s no HTTP (pass --live to GET)" % sku)
        return 0
    from spapi_smoke import (  # noqa: E402
        _load_config,
        _spapi_headers,
        request_lwa_access_token,
        resolve_cred,
    )
    import requests

    cfg_path = Path(args.config) if args.config else SMOKE / "config.local.json"
    cfg = _load_config(cfg_path)
    token = request_lwa_access_token(
        resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID"),
        resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET"),
        resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN"),
    )["access_token"]
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com").rstrip("/")
    seller = resolve_cred(cfg, "seller_id", "SPAPI_SELLER_ID")
    mid = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    ua = str(cfg.get("user_agent") or "OctasCompetitorInventory/0.1")
    headers = _spapi_headers(endpoint, token, ua)
    list_url = (
        "%s/listings/2021-08-01/items/%s/%s?marketplaceIds=%s&includedData=fulfillmentAvailability,summaries"
        % (endpoint, quote(seller, safe="-_.~"), quote(sku, safe="-_.~"), quote(mid))
    )
    r = requests.get(list_url, headers=headers, timeout=60)
    listings_qty = None
    if r.status_code == 200:
        listings_qty = qty_from_listings(r.json())
    print("listings http=%s qty=%s" % (r.status_code, listings_qty))
    if listings_qty is not None:
        print("source=listings sku=%s qty=%s master_write=no" % (sku, listings_qty))
        return 0
    fba_url = (
        "%s/fba/inventory/v1/summaries?details=true&granularityType=Marketplace"
        "&granularityId=%s&marketplaceIds=%s&sellerSkus=%s"
        % (endpoint, quote(mid), quote(mid), quote(sku))
    )
    r2 = requests.get(fba_url, headers=headers, timeout=60)
    fba_qty = None
    if r2.status_code == 200:
        fba_qty = qty_by_seller_sku(r2.json()).get(sku)
    print("fba http=%s qty=%s" % (r2.status_code, fba_qty))
    print("source=fba_or_none sku=%s qty=%s master_write=no" % (sku, fba_qty))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
