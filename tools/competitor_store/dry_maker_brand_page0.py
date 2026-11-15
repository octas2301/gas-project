# -*- coding: utf-8 -*-
"""倉庫ブランド上位N語の Catalog page0。既定20。結果を maker_alias.BRAND_PAGE0_HITS に書く。"""
from __future__ import annotations

import argparse
import re
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SMOKE = ROOT.parent / "spapi_smoke"
PATH3 = ROOT.parent / "purchase_research_path3"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(SMOKE))
sys.path.insert(0, str(PATH3))

from client import sheets_service  # noqa: E402
from dry_maker_sum import as_dicts, is_empty_brand, read_all  # noqa: E402
from maker_alias import BRAND_PAGE0_HITS, resolve_alias  # noqa: E402
from schema import SHEET_KEEPA_FULL  # noqa: E402
from search_catalog_keywords import asins_from_body, search_catalog_items  # noqa: E402
from spapi_smoke import _load_config, request_lwa_access_token, resolve_cred  # noqa: E402

COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
ALIAS_PATH = ROOT / "maker_alias.py"


def top_brands(full: list[dict], limit: int, offset: int) -> list[tuple[str, int]]:
    c = Counter()
    for r in full:
        b = str(r.get("ブランド") or "").strip()
        if not b or is_empty_brand(b):
            continue
        c[b] += 1
    ranked = sorted(c.items(), key=lambda x: (-x[1], x[0]))
    return ranked[offset : offset + limit]


def write_hits(new_hits: dict[str, str]) -> None:
    text = ALIAS_PATH.read_text(encoding="utf-8")
    merged = dict(BRAND_PAGE0_HITS)
    merged.update(new_hits)
    body_lines = ["BRAND_PAGE0_HITS: dict[str, str] = {"]
    for k in sorted(merged.keys()):
        body_lines.append("    %s: %s," % (repr(k), repr(merged[k])))
    body_lines.append("}")
    block = "\n".join(body_lines)
    pat = r"BRAND_PAGE0_HITS: dict\[str, str\] = \{.*?\}"
    if not re.search(pat, text, flags=re.S):
        raise SystemExit("BRAND_PAGE0_HITS block not found")
    text2 = re.sub(pat, block, text, count=1, flags=re.S)
    ALIAS_PATH.write_text(text2, encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("--offset", type=int, default=0, help="次の20は --offset 20")
    ap.add_argument("--apply-seed", action="store_true", help="maker_alias.py に結果を書く")
    args = ap.parse_args()

    svc = sheets_service(write=False, interactive=True)
    if not svc:
        print("no_creds")
        return 1
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    brands = top_brands(full, args.limit, args.offset)
    print("brand_slice offset=%d limit=%d n=%d" % (args.offset, args.limit, len(brands)))
    for b, n in brands:
        print("  n=%d parent=%s brand=%s" % (n, resolve_alias(b), b))

    cfg = _load_config(SMOKE / "config.local.json")
    client_id = resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID")
    client_secret = resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET")
    refresh = resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN")
    marketplace_id = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    user_agent = str(cfg.get("user_agent") or "OctasBrandPage0/1.0")
    token_body = request_lwa_access_token(client_id, client_secret, refresh)
    access = str(token_body.get("access_token") or "")

    day = time.strftime("%Y-%m-%d")
    new_hits = {}
    print("=== Catalog page0 ===")
    for brand, nwh in brands:
        resp = search_catalog_items(
            endpoint, access, marketplace_id, user_agent, brand, page_size=20, page_token=""
        )
        code = resp.status_code
        body = {}
        try:
            body = resp.json() if resp.content else {}
        except Exception:
            body = {}
        rows = asins_from_body(body)
        next_tok = bool(
            isinstance(body, dict) and (body.get("pagination") or {}).get("nextToken")
        )
        full_page = len(rows) >= 20 and next_tok
        label = "page0=%d%s(%s)" % (
            len(rows),
            "満杯" if full_page else "",
            day,
        )
        if code != 200:
            label = "http=%s(%s)" % (code, day)
        new_hits[brand] = label
        print(
            "wh=%d\tbrand=%s\tparent=%s\thttp=%s\t%s\t次頁=%s"
            % (nwh, brand, resolve_alias(brand), code, label, "有" if next_tok else "無")
        )
        time.sleep(1.1)

    if args.apply_seed:
        write_hits(new_hits)
        print("wrote BRAND_PAGE0_HITS n=%d" % len(new_hits))
    else:
        print("dry (no seed write). use --apply-seed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
