# -*- coding: utf-8 -*-
"""接続テスト T1: 製造メーカー1社・少量。サマリ実態→Catalog→Keepa→門→①候補。

既定 dry。--apply --live で書込。モールヒット非書。CAP既定5。
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SMOKE = ROOT.parent / "spapi_smoke"
PATH3 = ROOT.parent / "purchase_research_path3"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(SMOKE))
sys.path.insert(0, str(PATH3))

from apply_keepa_full import (  # noqa: E402
    COMPETITOR_SS,
    RESEARCH_SS,
    T_CAND,
    append_log,
    append_rows,
    as_dicts,
    prepend_rows,
    fetch_stats90,
    keepa_key,
    read_all,
    sheets_service,
    utc_now,
)
from dry_w3_gate import gate_product, load_profiles  # noqa: E402
from keepa_full import keepa_get_needed, latest_row_for_asin, product_to_full_row  # noqa: E402
from maker_alias import JITTAI_MAKER, jitta_for_maker  # noqa: E402
from schema import KEEPA_FULL_HEADERS, PURPOSE_RESEARCH, SHEET_HITS, SHEET_KEEPA_FULL  # noqa: E402
from search_catalog_keywords import asins_from_body, search_catalog_items  # noqa: E402
from spapi_smoke import _load_config, request_lwa_access_token, resolve_cred  # noqa: E402

# T1 既定: 製造メーカー・倉庫少なめ
DEFAULT_MAKER = "尾西食品"
DEFAULT_CAP = 5


def catalog_page0_title_hits(maker: str, access: str, endpoint: str, mp: str, ua: str) -> list[dict]:
    resp = search_catalog_items(endpoint, access, mp, ua, maker, page_size=20, page_token="")
    body = resp.json() if resp.content else {}
    rows = asins_from_body(body)
    out = []
    for r in rows:
        title = str(r.get("title") or "")
        brand = str(r.get("brand") or "")
        blob = title + " " + brand
        if maker in blob:
            out.append(r)
    return out, resp.status_code, bool((body.get("pagination") or {}).get("nextToken"))


def cand_map(rows: list[dict]) -> dict[str, int]:
    m = {}
    for i, r in enumerate(rows):
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            m[a] = i
    return m


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--maker", default=DEFAULT_MAKER)
    ap.add_argument("--cap", type=int, default=DEFAULT_CAP)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--run-id", default="")
    args = ap.parse_args()
    maker = str(args.maker).strip()
    cap = max(1, min(20, int(args.cap)))
    run_id = args.run_id or ("pr_t1_" + utc_now()[:10].replace("-", "") + "_" + maker[:6])

    j = jitta_for_maker(maker)
    print("T1 maker=%s 実態=%s (expect %s) cap=%d" % (maker, j, JITTAI_MAKER, cap))
    if j != JITTAI_MAKER:
        print("FAIL: 実態が製造メーカーではない。接続T1の対象外")
        return 2

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
    ua = "OctasT1Connect/1.0"

    hits, code, next_tok = catalog_page0_title_hits(maker, access, endpoint, mp, ua)
    print(
        "catalog http=%s page0_title_hit=%d next=%s head=%s"
        % (code, len(hits), "Y" if next_tok else "N", ",".join(r["asin"] for r in hits[:3]))
    )
    want = [str(r["asin"]).upper() for r in hits[:cap]]
    if not want:
        print("FAIL: タイトル含有0。入口語ではつながらない")
        return 3

    svc = sheets_service(write=True, interactive=True)
    if not svc:
        print("no_creds")
        return 1
    hits_n = max(0, len(read_all(svc, COMPETITOR_SS, SHEET_HITS)) - 1)
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    need, skip = [], []
    for a in want:
        if keepa_get_needed(latest_row_for_asin(full, a)):
            need.append(a)
        else:
            skip.append(a)
    print("want=%d need_get=%d skip_fresh=%d hits_sheet=%d" % (len(want), len(need), len(skip), hits_n))

    if not args.apply:
        print("dry_ok (no write). next: --apply --live")
        return 0
    if not args.live:
        print("need --live")
        return 2

    key = keepa_key()
    if not key:
        print("NO_KEEPA_KEY")
        return 2
    fetched = utc_now()
    products_by = {}
    recs = []
    data = {"products": [], "tokensConsumed": 0, "tokensLeft": ""}
    if need:
        data = fetch_stats90(key, need)
    consumed = data.get("tokensConsumed") or 0
    left = data.get("tokensLeft")
    for p in data.get("products") or []:
        rec = product_to_full_row(p, fetched, purpose=PURPOSE_RESEARCH)
        products_by[rec["ASIN"]] = p
        if not rec["ASIN"]:
            continue
        latest = latest_row_for_asin(full, rec["ASIN"])
        if latest and str(latest.get("価格指紋") or "") == rec["価格指紋"]:
            continue
        rawj = json.loads(rec["生JSON"] or "{}")
        if "csv" in rawj:
            print("FAIL csv", rec["ASIN"])
            return 4
        recs.append(rec)
        full.append(rec)
    if recs:
        live_h = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))[0] or list(KEEPA_FULL_HEADERS)
        prepend_rows(svc, COMPETITOR_SS, SHEET_KEEPA_FULL, live_h, recs)

    _, full2 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    for a in skip:
        row = latest_row_for_asin(full2, a)
        if not row:
            continue
        try:
            products_by[a] = json.loads(row.get("生JSON") or "{}")
        except json.JSONDecodeError:
            pass

    profs = load_profiles(svc)
    ch, crow = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    idx = {h: i for i, h in enumerate(ch)}
    by = cand_map(crow)
    new_rows = []
    pass_n = drop_n = 0
    for a in want:
        p = products_by.get(a) or {}
        if not p:
            print("WARN no_product", a)
            continue
        st, why = gate_product(p, profs) if profs else ("落ち", "プロファイルなし")
        if st == "通過":
            pass_n += 1
        else:
            drop_n += 1
        eans = p.get("eanList") or []
        rec = {
            "メーカー": maker,
            "ASIN": a,
            "商品名": str(p.get("title") or ""),
            "税込価格": "",
            "順位90": "",
            "JAN": str(eans[0] if eans else p.get("ean") or ""),
            "発見経路": "t1_connect_catalog",
            "sellerId": "",
            "門結果": st,
            "門理由": why,
            "runId": run_id,
            "取得日時": fetched,
        }
        if a in by:
            r = crow[by[a]]
            if r.get("sellerId"):
                rec["sellerId"] = r.get("sellerId")
            line = [r.get(h, "") for h in ch]
            for h, v in rec.items():
                if h in idx:
                    line[idx[h]] = v
            crow[by[a]] = {ch[i]: line[i] for i in range(len(ch))}
            rng = "'%s'!A%d" % (T_CAND.replace("'", "''"), by[a] + 2)
            svc.spreadsheets().values().update(
                spreadsheetId=RESEARCH_SS,
                range=rng,
                valueInputOption="RAW",
                body={"values": [line]},
            ).execute()
        else:
            new_rows.append(rec)
    if new_rows:
        append_rows(svc, RESEARCH_SS, T_CAND, ch, new_rows)

    hits_after = max(0, len(read_all(svc, COMPETITOR_SS, SHEET_HITS)) - 1)
    msg = (
        "T1 ok maker=%s runId=%s full+=%s consumed=%s left=%s pass=%s drop=%s new_cand=%s"
        % (maker, run_id, len(recs), consumed, left, pass_n, drop_n, len(new_rows))
    )
    append_log(svc, "T1", msg)
    print(msg)
    if hits_after != hits_n:
        print("FAIL hits changed")
        return 5
    print("connect_T1_PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
