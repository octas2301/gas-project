# -*- coding: utf-8 -*-
"""入口ヒット名が製造メーカーか: Catalog page0属性＋Keepa既存。シート非書。"""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SMOKE = ROOT.parent / "spapi_smoke"
PATH3 = ROOT.parent / "purchase_research_path3"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(SMOKE))
sys.path.insert(0, str(PATH3))

from client import sheets_service  # noqa: E402
from dry_maker_sum import as_dicts, is_empty_brand, norm_maker, read_all  # noqa: E402
from maker_alias import resolve_alias  # noqa: E402
from schema import SHEET_KEEPA_FULL  # noqa: E402
from search_catalog_keywords import asins_from_body, search_catalog_items  # noqa: E402
from spapi_smoke import _load_config, request_lwa_access_token, resolve_cred  # noqa: E402

COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"

# 入口ヒットした語のサンプル（製造／ブランド／店を混ぜる）
SAMPLES = [
    "AGF(エージーエフ)",
    "Nagatanien",
    "石原水産",
    "エバラ",
    "五木食品",
    "わが街とくさんネット",
    "扶桑社",
    "リフココ",
    "天然生活",
    "Sato",
]


def main() -> int:
    svc = sheets_service(write=False, interactive=True)
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    by_brand = {}
    for r in full:
        b = str(r.get("ブランド") or "").strip()
        if not b or is_empty_brand(b):
            continue
        by_brand.setdefault(b, []).append(r)

    cfg = _load_config(SMOKE / "config.local.json")
    access = str(
        request_lwa_access_token(
            resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID"),
            resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET"),
            resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN"),
        ).get("access_token")
        or ""
    )
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    mp = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    ua = "OctasMfrProbe/1.0"

    print("name\tparent\tKeepa_n\tKeepa_mfr≈brand\tKeepa_mfr上位\tCat_n\tCat_mfr≈brand\tCat_brand上位\tCat_mfr上位")
    for name in SAMPLES:
        rows_k = by_brand.get(name) or []
        same_k = 0
        mfr_c = {}
        for r in rows_k:
            mfr = str(r.get("製造者") or "").strip() or "(空)"
            mfr_c[mfr] = mfr_c.get(mfr, 0) + 1
            if norm_maker(mfr) and (
                norm_maker(mfr) == norm_maker(name)
                or norm_maker(name) in norm_maker(mfr)
                or resolve_alias(name) == norm_maker(mfr)
            ):
                same_k += 1
        top_mfr_k = sorted(mfr_c.items(), key=lambda x: -x[1])[:2]

        resp = search_catalog_items(endpoint, access, mp, ua, name, page_size=20)
        body = {}
        try:
            body = resp.json() if resp.content else {}
        except Exception:
            body = {}
        rows_c = asins_from_body(body)
        same_c = 0
        bc, mc = {}, {}
        for r in rows_c:
            b = (r.get("brand") or "").strip() or "(空)"
            m = (r.get("manufacturer") or "").strip() or "(空)"
            bc[b] = bc.get(b, 0) + 1
            mc[m] = mc.get(m, 0) + 1
            if m != "(空)" and (
                norm_maker(m) == norm_maker(b)
                or norm_maker(name) in norm_maker(m)
                or resolve_alias(name) == norm_maker(m)
            ):
                same_c += 1
        print(
            "%s\t%s\t%d\t%d/%d\t%s\t%d\t%d/%d\t%s\t%s"
            % (
                name,
                resolve_alias(name),
                len(rows_k),
                same_k,
                len(rows_k),
                top_mfr_k,
                len(rows_c),
                same_c,
                len(rows_c),
                sorted(bc.items(), key=lambda x: -x[1])[:2],
                sorted(mc.items(), key=lambda x: -x[1])[:2],
            )
        )
        time.sleep(1.1)
    print("probe_mfr_ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
