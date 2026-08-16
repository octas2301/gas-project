# -*- coding: utf-8 -*-
"""A①: 門通過のカート店（倉庫列）＋新品オファー（SP-API）をセラー台帳へ。新規100/回。Keepa GETなし。構成比なし。"""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from apply_keepa_full import (
    COMPETITOR_SS,
    RESEARCH_SS,
    T_CAND,
    append_log,
    append_rows,
    as_dicts,
    read_all,
    sheets_service,
)
from dry_o2_offers import seller_id_of, spapi_item_offers_cap, spapi_session
from dry_t1_list_paste import ORIG_SS
from schema import SELLER_HEADERS, SHEET_KEEPA_FULL, SHEET_SELLER

CAP_NEW = 100
SKIP_IDS = {"ATVPDKIKX0DER", "A1VC38T7YXB528"}


def ok_sid(s: str) -> bool:
    s = (s or "").strip()
    if len(s) < 10 or s in SKIP_IDS:
        return False
    return True


def pass_asins(cand: list[dict]) -> list[str]:
    out, seen = [], set()
    for r in cand:
        if str(r.get("門結果") or "") != "通過":
            continue
        a = str(r.get("ASIN") or "").strip().upper()
        if len(a) == 10 and a not in seen:
            seen.add(a)
            out.append(a)
    return out


def bb_map(full: list[dict]) -> dict[str, str]:
    m = {}
    for r in full:
        a = str(r.get("ASIN") or "").strip().upper()
        s = str(r.get("BuyBoxセラー") or "").strip()
        if a and ok_sid(s):
            m[a] = s
    return m


def spapi_item_offers_batch(asins: list[str], sess: dict) -> dict:
    import requests
    from spapi_smoke import _spapi_headers

    chunk = asins[:20]
    reqs = []
    for a in chunk:
        reqs.append(
            {
                "uri": "/products/pricing/v0/items/%s/offers" % a,
                "method": "GET",
                "MarketplaceId": sess["marketplace_id"],
                "ItemCondition": "New",
            }
        )
    url = sess["endpoint"].rstrip("/") + "/batches/products/pricing/v0/itemOffers"
    hdr = _spapi_headers(sess["endpoint"], sess["access"], sess["ua"])
    hdr["content-type"] = "application/json"
    resp = requests.post(url, headers=hdr, json={"requests": reqs}, timeout=90)
    by: dict[str, list[str]] = {}
    codes = []
    err = ""
    try:
        body = resp.json()
    except Exception:
        return {"http": resp.status_code, "err": (resp.text or "")[:200], "by": {}, "codes": []}
    if resp.status_code != 200:
        return {"http": resp.status_code, "err": str(body.get("errors") or body)[:240], "by": {}, "codes": [resp.status_code]}
    for row in body.get("responses") or []:
        st = (row.get("status") or {}).get("statusCode")
        codes.append(st)
        b = row.get("body") or {}
        payload = b.get("payload") or b
        ident = payload.get("Identifier") or payload.get("identifier") or {}
        a = str(ident.get("ASIN") or ident.get("asin") or payload.get("ASIN") or "").upper()
        offers = payload.get("Offers") or payload.get("offers") or []
        ids = []
        for o in offers:
            sid = seller_id_of(o)
            if sid:
                ids.append(sid)
        if a:
            by[a] = ids
        if st and int(st) >= 400 and not err:
            err = str((b.get("errors") or st))[:120]
    return {"http": resp.status_code, "err": err, "by": by, "codes": codes, "n": len(chunk)}


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--smoke", action="store_true", help="dry時に先頭1 ASINだけ SP-API")
    ap.add_argument("--after", default="", help="このASINの次から（429再開）")
    ap.add_argument("--sleep", type=float, default=1.1, help="getItemOffers 間隔秒（公式1/秒）")
    ap.add_argument("--max-new", type=int, default=0, help="0=CAP_NEW。再開時は残り件数")
    ap.add_argument("--batch", action="store_true", help="getItemOffersBatch 最大20")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    if RESEARCH_SS == ORIG_SS:
        print("refuse original")
        return 2
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    _, srows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    asins = pass_asins(cand)
    after = str(args.after or "").strip().upper()
    if after:
        try:
            i = asins.index(after)
            asins = asins[i + 1 :]
        except ValueError:
            print("after not in pass list", after)
            return 2
    existing = {str(r.get("sellerId") or "").strip() for r in srows if str(r.get("sellerId") or "").strip()}
    bb = bb_map(full)
    bb_new = [s for s in set(bb.values()) if s not in existing]
    cap = args.max_new if args.max_new > 0 else CAP_NEW
    line = (
        "runId=pr_20260816_a1dry pass=%d after=%s remain_asin=%d seller_n=%d bb_new=%d cap=%d sleep=%s KeepaGETなし"
        % (len(pass_asins(cand)), after or "-", len(asins), len(existing), len(bb_new), cap, args.sleep)
    )
    print(line)
    ok = RESEARCH_SS != ORIG_SS and len(asins) >= 1 and cap <= CAP_NEW
    print("PLAN", "PASS" if ok else "FAIL")
    if not args.apply:
        append_log(svc, "A1", line)
        if args.smoke and asins:
            time.sleep(max(0.0, args.sleep))
            sess = spapi_session()
            if args.batch:
                r = spapi_item_offers_batch(asins[: min(20, len(asins))], sess)
                print("smoke_batch http", r.get("http"), "n", r.get("n"), "codes", r.get("codes"), "err", r.get("err") or "")
                if r.get("http") != 200:
                    print("VERIFY FAIL smoke")
                    return 1
            else:
                r = spapi_item_offers_cap(asins[0], sess, cap=20)
                print("smoke", asins[0], "http", r.get("http"), "uniq", r.get("uniq"), "sec", r.get("sec"), "err", r.get("err") or "")
                if 200 not in (r.get("http") or []):
                    print("VERIFY FAIL smoke")
                    return 1
        print("VERIFY", "PASS" if ok else "FAIL")
        return 0 if ok else 1
    if not ok:
        return 1
    sess = spapi_session()
    new_ids: list[str] = []
    have = set(existing)
    n_api = 0
    stopped = ""
    last_a = ""
    if args.batch:
        i = 0
        while i < len(asins) and len(new_ids) < cap:
            chunk = asins[i : i + 20]
            i += 20
            time.sleep(max(2.0, args.sleep))
            n_api += 1
            last_a = chunk[-1]
            r = spapi_item_offers_batch(chunk, sess)
            print("batch http", r.get("http"), "n", r.get("n"), "codes", r.get("codes"))
            if r.get("http") == 429 or 429 in (r.get("codes") or []):
                stopped = "429"
                break
            if r.get("http") != 200:
                stopped = "http"
                print("STOP batch", r.get("err"))
                break
            bymap = r.get("by") or {}
            for a in chunk:
                ids = list(bymap.get(a) or [])
                if a in bb:
                    ids = [bb[a]] + ids
                for s in ids:
                    if not ok_sid(s) or s in have:
                        continue
                    have.add(s)
                    new_ids.append(s)
                    if len(new_ids) >= cap:
                        stopped = "cap"
                        break
            if stopped:
                break
    else:
        for a in asins:
            if len(new_ids) >= cap:
                stopped = "cap"
                break
            ids = []
            if a in bb:
                ids.append(bb[a])
            time.sleep(max(0.0, args.sleep))
            n_api += 1
            last_a = a
            r = spapi_item_offers_cap(a, sess, cap=20)
            if 429 in (r.get("http") or []):
                stopped = "429"
                print("STOP 429 after", n_api, a)
                break
            if 200 not in (r.get("http") or []):
                stopped = "http"
                print("STOP http", r.get("http"), r.get("err"), a)
                break
            ids.extend(r.get("ids") or r.get("head") or [])
            for s in ids:
                if not ok_sid(s) or s in have:
                    continue
                have.add(s)
                new_ids.append(s)
                if len(new_ids) >= cap:
                    stopped = "cap"
                    break
    print("new_plan", len(new_ids), "api", n_api, "last", last_a, "stopped", stopped or "end")
    recs = []
    for s in new_ids:
        rec = {k: "" for k in SELLER_HEADERS}
        rec["sellerId"] = s
        rec["採取元"] = "bb+spapi_new"
        rec["メモ"] = "A1貯め 構成比は店洗い時"
        recs.append(rec)
    live_h = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))[0] or list(SELLER_HEADERS)
    if recs:
        use = [h for h in live_h if h] or list(SELLER_HEADERS)
        rows = [{h: rec.get(h, "") for h in use} for rec in recs]
        append_rows(svc, COMPETITOR_SS, SHEET_SELLER, use, rows)
    _, after_rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    ids2 = {str(r.get("sellerId") or "").strip() for r in after_rows}
    hit = sum(1 for s in new_ids if s in ids2)
    line2 = (
        "runId=pr_20260816_a1bcol new=%d hit=%d api=%d last=%s stop=%s sleep=%s KeepaGETなし 品番非書"
        % (len(new_ids), hit, n_api, last_a, stopped or "ok", args.sleep)
    )
    append_log(svc, "A1", line2)
    print(line2)
    vok = hit == len(new_ids) and stopped != "http"
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
