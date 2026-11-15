# -*- coding: utf-8 -*-
"""小テスト実行: Catalog page0件数＋Keepa既存列突合。再GETなし。シート非書。"""
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
from schema import SHEET_KEEPA_FULL  # noqa: E402
from search_catalog_keywords import asins_from_body, search_catalog_items  # noqa: E402
from spapi_smoke import _load_config, request_lwa_access_token, resolve_cred  # noqa: E402

RESEARCH_SS = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
T_CAND = "①候補"

WORDS = [
    ("ブランド→親", "AGF", "味の素"),
    ("ブランド→親", "エージーエフ", "味の素"),
    ("有名メーカー", "エバラ", "エバラ"),
    ("有名メーカー", "エースコック", "エースコック"),
    ("有名メーカー", "S&B", "エスビー食品"),
    ("有名メーカー", "エスビー食品", "エスビー食品"),
    ("表記ゆれ", "いしもと", "いし本食品工業"),
    ("表記ゆれ", "いし本食品工業", "いし本食品工業"),
]

ASINS = [
    ("候補と一致", "B0DF7P45TL"),
    ("候補と一致", "B084MYZWXJ"),
    ("候補と一致", "B09CP5Y7GS"),
    ("ブランド≠候補", "B08NJSWSBH"),
    ("ブランド≠候補", "B0GF6X8NGG"),
    ("ブランド≠候補", "B0FV5RC1X4"),
    ("ノーブランド", "B0DV46Z3D1"),
    ("ノーブランド", "B0F9QZT8WW"),
    ("ノーブランド", "B0DVB9VK8Q"),
    ("親未確定", "B0CXMFHB8M"),
    ("親未確定", "B005J8SALO"),
    ("親未確定", "B07CRNDJGT"),
    ("製造者ゴミ", "B0F4PVZY2N"),
    ("製造者ゴミ", "B0DJGLDYJK"),
    ("製造者ゴミ", "B0GY42TRTQ"),
]


def title_hit(title: str, parent: str, word: str) -> bool:
    t = title or ""
    return (parent and parent in t) or (word and word in t)


def run_catalog_page0() -> None:
    cfg = _load_config(SMOKE / "config.local.json")
    client_id = resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID")
    client_secret = resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET")
    refresh = resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN")
    marketplace_id = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    user_agent = str(cfg.get("user_agent") or "OctasMakerProbe/1.0")
    token_body = request_lwa_access_token(client_id, client_secret, refresh)
    access = str(token_body.get("access_token") or "")
    print("=== Catalog page0（keywordsのみ・brandNamesなし）===")
    for kind, word, parent in WORDS:
        resp = search_catalog_items(
            endpoint, access, marketplace_id, user_agent, word, page_size=20, page_token=""
        )
        code = resp.status_code
        body = {}
        try:
            body = resp.json() if resp.content else {}
        except Exception:
            body = {}
        rows = asins_from_body(body)
        n = len(rows)
        next_tok = bool(body.get("pagination", {}).get("nextToken") if isinstance(body, dict) else False)
        # page0満杯かつ next あり → 打ち切り疑い（件数は20+）
        full_page = n >= 20 and next_tok
        th = sum(1 for r in rows if title_hit(r.get("title") or "", parent, word))
        brands = {}
        for r in rows:
            b = (r.get("brand") or "").strip() or "(空)"
            brands[b] = brands.get(b, 0) + 1
        top_b = sorted(brands.items(), key=lambda x: -x[1])[:3]
        print(
            "%s\t語=%s\t期待親=%s\thttp=%s\tpage0=%d\t次頁=%s\t満杯疑い=%s\tタイトルに親or語=%d/%d\tbrand上位=%s"
            % (
                kind,
                word,
                parent,
                code,
                n,
                "有" if next_tok else "無",
                "Y" if full_page else "N",
                th,
                n,
                top_b,
            )
        )
        time.sleep(1.1)


def run_keepa_sheet() -> None:
    svc = sheets_service(write=False, interactive=True)
    if not svc:
        print("no_creds keepa")
        return
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    by_full = {}
    for r in full:
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            by_full[a] = r
    by_cand = {}
    for r in cand:
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            by_cand[a] = r
    print("=== Keepa既存列突合（再GETなし）===")
    miss = 0
    for kind, asin in ASINS:
        fr = by_full.get(asin)
        cr = by_cand.get(asin)
        if not fr:
            miss += 1
            print("%s\t%s\t倉庫なし" % (kind, asin))
            continue
        brand = str(fr.get("ブランド") or "").strip()
        mfr = str(fr.get("製造者") or "").strip()
        title = str(fr.get("商品名") or "").strip()
        cand_m = str((cr or {}).get("メーカー") or "").strip()
        # 判定メモ
        brand_empty = is_empty_brand(brand)
        mfr_n = norm_maker(mfr)
        cand_n = norm_maker(cand_m)
        mfr_ok = bool(cand_n and mfr_n and (mfr_n == cand_n or cand_n in mfr or mfr_n in cand_n))
        title_ok = bool(cand_n and cand_n in title)
        print(
            "%s\t%s\tcand=%s\tbrand=%s\tmfr=%s\tbrand空=%s\tmfr≈cand=%s\ttitle含cand=%s\t| %s"
            % (
                kind,
                asin,
                cand_m or "(空)",
                brand or "(空)",
                mfr or "(空)",
                "Y" if brand_empty else "N",
                "Y" if mfr_ok else "N",
                "Y" if title_ok else "N",
                title[:50],
            )
        )
    print("倉庫miss=%d / %d" % (miss, len(ASINS)))


def main() -> int:
    run_keepa_sheet()
    run_catalog_page0()
    print("probe_run_ok (no sheet write)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
