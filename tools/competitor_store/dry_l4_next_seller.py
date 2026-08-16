# -*- coding: utf-8 -*-
"""L4: ①候補からモリタ以外1件。/seller storefront=0。asinList禁止。"""
from __future__ import annotations

import gzip
import json
import sys
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "purchase_research_path3"))

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, T_CAND, append_log, append_rows, as_dicts, read_all, sheets_service
from dry_s4_seller0 import FOOD, parse_seller, seller_url
from init_store import ensure_seller_google
from keepa_csv_vs_api import keepa_key
from schema import AMAZON_SELLER_CAT_COLS, SELLER_HEADERS, SHEET_KEEPA_FULL, SHEET_SELLER

MORITA = "AYC4Z8PML8T30"


def pick_next(cand: list[dict], existing: set[str], full: list[dict]) -> str:
    morita_asins = {
        str(r.get("ASIN") or "").strip().upper()
        for r in cand
        if str(r.get("sellerId") or "").strip().upper() == MORITA
    }
    for r in cand:
        s = str(r.get("sellerId") or "").strip()
        if len(s) >= 10 and s.upper() != MORITA and s not in existing:
            return s
    from collections import Counter

    bb = Counter()
    for r in full:
        a = str(r.get("ASIN") or "").strip().upper()
        if morita_asins and a not in morita_asins:
            continue
        s = str(r.get("BuyBoxセラー") or "").strip()
        if len(s) >= 10 and s.upper() != MORITA and s not in existing:
            bb[s] += 1
    if bb:
        return bb.most_common(1)[0][0]
    return ""


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--seller", default="", help="指定sellerId。空なら自動ピック")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    h, srows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    existing = {str(r.get("sellerId") or "").strip() for r in srows}
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    nxt = str(args.seller or "").strip() or pick_next(cand, existing, full)
    url = seller_url(nxt or "X", "REDACTED")
    ok = bool(nxt) and "storefront=0" in url and "storefront=1" not in url
    line = "runId=pr_20260816_l4dry next=%s GET形 %s" % (nxt or "-", "PASS" if ok else "FAIL")
    print(line)
    if not args.apply:
        append_log(svc, "L4", line)
        return 0 if ok else 1
    if not ok:
        return 1
    key = keepa_key()
    if not key:
        return 2
    req = Request(seller_url(nxt, key), method="GET")
    with urlopen(req, timeout=60) as resp:
        raw = resp.read()
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    body = json.loads(raw.decode("utf-8"))
    parsed = parse_seller(body, nxt)
    consumed = body.get("tokensConsumed")
    if parsed["_has_asinlist"] or parsed["_n_cat"] < 1:
        append_log(svc, "L4", "runId=pr_20260816_l4col FAIL asinList=%s cats=%s" % (parsed["_asin_n"], parsed["_n_cat"]))
        print("FAIL", parsed["_n_cat"], parsed["_asin_n"])
        return 1
    print("ensure", ensure_seller_google(COMPETITOR_SS))
    # map mix parts "cid:n:pct%" to amazon cols via カテゴリ構成 + names later L4 only IDs
    # fill 28 cols 0, 不明 leftover; names need /category — L4b in same chunk: GET category if we have ids
    mix_raw = parsed["カテゴリ構成"]
    ids = [p.split(":")[0] for p in mix_raw.split("|") if p.split(":")[0].isdigit()]
    names = {}
    if ids:
        from dry_s4b_catcols import cat_url, decode_body, names_from_body

        req2 = Request(cat_url(ids, key), method="GET")
        with urlopen(req2, timeout=60) as resp2:
            b2 = decode_body(resp2.read())
        names = names_from_body(b2)
        print("named", len(names), "cat_consumed", b2.get("tokensConsumed"))
    pct_by_name = {}
    leftover = 0.0
    for part in mix_raw.split("|"):
        bits = part.split(":")
        if len(bits) < 3:
            continue
        cid, pct_s = bits[0], bits[2].replace("%", "")
        try:
            pct = float(pct_s)
        except ValueError:
            continue
        nm = names.get(cid) or ""
        if nm in AMAZON_SELLER_CAT_COLS:
            pct_by_name[nm] = pct_by_name.get(nm, 0) + pct
        else:
            leftover += pct
    row = {k: "" for k in SELLER_HEADERS}
    row["sellerId"] = nxt
    row["店名"] = parsed["店名"]
    row["対象カテゴリ"] = FOOD
    row["巡回日"] = ""
    row["asinList件数"] = ""
    row["抽出メーカー"] = ""
    row["ピック"] = ""
    row["採取元"] = "storefront0"
    row["メモ"] = "L4 構成比。queryはL5"
    row["カテゴリ構成"] = mix_raw
    row["食品比率"] = parsed["食品比率"]
    row["ストアASIN数"] = parsed["ストアASIN数"]
    row["卸仮説"] = parsed["卸仮説"]
    main_nm = names.get(mix_raw.split(":")[0] if mix_raw else "") or mix_raw.split(":")[0]
    try:
        main_pct = float(str(parsed["食品比率"] or 0))
    except ValueError:
        main_pct = 0
    if mix_raw:
        b0 = mix_raw.split("|")[0].split(":")
        main_nm = names.get(b0[0], b0[0])
        try:
            main_pct = float(b0[2].replace("%", ""))
        except (IndexError, ValueError):
            pass
    row["メインカテゴリ"] = "%s %.1f%%" % (main_nm, main_pct)
    for c in AMAZON_SELLER_CAT_COLS:
        if c == "不明":
            row[c] = leftover
        else:
            row[c] = pct_by_name.get(c, 0)
    live_h = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))[0]
    append_rows(svc, COMPETITOR_SS, SHEET_SELLER, live_h, [row])
    h2, rows2 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    ids_now = [str(r.get("sellerId") or "") for r in rows2]
    vok = nxt in ids_now and len(rows2) >= 2 and not parsed["_has_asinlist"]
    line2 = "runId=pr_20260816_l4col seller=%s cats=%s food=%s consumed=%s n=%d %s" % (
        nxt,
        parsed["_n_cat"],
        parsed["食品比率"],
        consumed,
        len(rows2),
        "PASS" if vok else "FAIL",
    )
    append_log(svc, "L4", line2)
    print(line2)
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
