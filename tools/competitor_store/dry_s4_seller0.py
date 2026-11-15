# -*- coding: utf-8 -*-
"""S4: GET /seller storefront=0。構成比。asinList・storefront=1禁止。既定 dry。"""
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

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, T_CAND, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from init_store import ensure_seller_google
from keepa_csv_vs_api import keepa_key
from schema import SHEET_SELLER

FOOD = "57239051"
SELLER_API = "https://api.keepa.com/seller"
COLS = ("カテゴリ構成", "食品比率", "ストアASIN数", "卸仮説")


def seller_url(seller_id: str, key: str) -> str:
    q = urlencode({"key": key, "domain": "5", "seller": seller_id, "storefront": "0"})
    return SELLER_API + "?" + q


def parse_seller(body: dict, seller_id: str) -> dict:
    sellers = body.get("sellers")
    rec = {}
    if isinstance(sellers, dict):
        rec = sellers.get(seller_id) or sellers.get(seller_id.upper()) or {}
        if not rec and sellers:
            rec = next(iter(sellers.values()))
    if isinstance(sellers, list) and sellers:
        rec = sellers[0] if isinstance(sellers[0], dict) else {}
    stats = rec.get("sellerCategoryStatistics") or []
    rows = []
    for it in stats:
        if not isinstance(it, dict):
            continue
        cid = str(it.get("catId") or "").strip()
        try:
            n = int(it.get("productCount") or 0)
        except (TypeError, ValueError):
            n = 0
        if cid and n > 0:
            rows.append((n, cid, it.get("avg30SalesRank")))
    tot = sum(x[0] for x in rows) or 0
    rows.sort(reverse=True)
    parts = []
    food_n = 0
    for n, cid, avg in rows:
        pct = (100.0 * n / tot) if tot else 0.0
        if cid == FOOD:
            food_n += n
        avg_s = "" if avg in (None, "", -1) else str(avg)
        parts.append("%s:%d:%.1f%%%s" % (cid, n, pct, (":r" + avg_s) if avg_s else ""))
    food_pct = round(100.0 * food_n / tot, 1) if tot else 0.0
    if tot and len(rows) == 1 and rows[0][1] == FOOD:
        hypo = "食品のみ（卸の可能性・要人判断）"
    elif food_pct >= 90:
        hypo = "食品偏重（卸の可能性・要人判断）"
    elif tot:
        hypo = "複数カテゴリ"
    else:
        hypo = "構成比なし"
    tsf = rec.get("totalStorefrontAsins")
    store_n = ""
    if isinstance(tsf, list) and len(tsf) >= 2:
        store_n = str(tsf[1])
    elif tsf not in (None, ""):
        store_n = str(tsf)
    asin_list = rec.get("asinList") or []
    return {
        "店名": str(rec.get("sellerName") or rec.get("shopName") or "").strip(),
        "カテゴリ構成": "|".join(parts[:20]),
        "食品比率": str(food_pct) if tot else "",
        "ストアASIN数": store_n,
        "卸仮説": hypo,
        "_n_cat": len(rows),
        "_has_asinlist": bool(asin_list),
        "_asin_n": len(asin_list) if isinstance(asin_list, list) else 0,
        "_name_keys": list(rec.keys())[:12] if isinstance(rec, dict) else [],
    }


def fill_cand_names(svc, cap: int, apply: bool) -> int:
    """①候補の欠ける sellerId を /seller storefront=0 で名前付け、台帳へ残す。"""
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    _, srows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    nmap = {}
    for r in srows:
        sid = str(r.get("sellerId") or "").strip().upper()
        if not sid:
            continue
        nm = str(r.get("店名") or "").strip()
        if nm:
            nmap[sid] = nm
    need = []
    seen = set()
    for r in cand:
        for p in str(r.get("sellerId") or "").replace(",", "|").split("|"):
            sid = p.strip().upper()
            if len(sid) < 10 or sid in nmap or sid in seen:
                continue
            seen.add(sid)
            need.append(sid)
    leftover = max(0, len(need) - cap)
    print("missing_ids", len(need), "cap", cap, "leftover", leftover)
    print("sample", need[:8])
    if not apply:
        print("dry fill-cand-names (pass --apply)")
        return 0
    key = keepa_key()
    if not key:
        print("no_key")
        return 2
    got = {}
    for sid in need[:cap]:
        req = Request(seller_url(sid, key), method="GET")
        with urlopen(req, timeout=60) as resp:
            raw = resp.read()
        if raw[:2] == b"\x1f\x8b":
            raw = gzip.decompress(raw)
        body = json.loads(raw.decode("utf-8"))
        parsed = parse_seller(body, sid)
        name = parsed["店名"]
        print("GET", sid, "name=", name or "-", "consumed", body.get("tokensConsumed"), "asinList", parsed["_asin_n"])
        if parsed["_has_asinlist"]:
            print("WARN asinList unexpected", sid)
        if name:
            got[sid] = name
            nmap[sid] = name
    if got:
        print("ensure", ensure_seller_google(COMPETITOR_SS))
        h2, srows2 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
        idx_nm = h2.index("店名") if "店名" in h2 else -1
        live_ids = {str(r.get("sellerId") or "").strip().upper(): i for i, r in enumerate(srows2)}
        new_rows = []
        for sid, name in got.items():
            if sid in live_ids:
                i = live_ids[sid]
                if not str(srows2[i].get("店名") or "").strip() and idx_nm >= 0:
                    col = col_letter(idx_nm + 1)
                    svc.spreadsheets().values().update(
                        spreadsheetId=COMPETITOR_SS,
                        range="'%s'!%s%d" % (SHEET_SELLER.replace("'", "''"), col, i + 2),
                        valueInputOption="RAW",
                        body={"values": [[name]]},
                    ).execute()
            else:
                rec = {k: "" for k in h2}
                rec["sellerId"] = sid
                rec["店名"] = name
                if "調査CK" in rec:
                    rec["調査CK"] = False
                new_rows.append([rec.get(k, "") for k in h2])
        if new_rows:
            svc.spreadsheets().values().append(
                spreadsheetId=COMPETITOR_SS,
                range="'" + SHEET_SELLER + "'!A1",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": new_rows},
            ).execute()
    ch, crow = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    if "店名" not in ch:
        print("no 店名 col")
        return 1
    ni = ch.index("店名")
    col = col_letter(ni + 1)
    filled = 0
    for i, r in enumerate(crow):
        ids = [p.strip().upper() for p in str(r.get("sellerId") or "").replace(",", "|").split("|") if p.strip()]
        names = []
        seen_n = set()
        for sid in ids:
            n = nmap.get(sid) or ""
            if n and n not in seen_n:
                seen_n.add(n)
                names.append(n)
        val = "|".join(names)
        old = str(r.get("店名") or "").strip()
        if not val or val == old:
            continue
        svc.spreadsheets().values().update(
            spreadsheetId=RESEARCH_SS,
            range="'%s'!%s%d" % (T_CAND.replace("'", "''"), col, i + 2),
            valueInputOption="RAW",
            body={"values": [[val]]},
        ).execute()
        filled += 1
    line = "fill_names got=%d leftover=%d cand_upd=%d" % (len(got), leftover, filled)
    append_log(svc, "S4", line)
    print(line)
    return 0


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--fill-cand-names", action="store_true")
    ap.add_argument("--cap", type=int, default=40)
    args = ap.parse_args()
    if args.fill_cand_names:
        svc = sheets_service(write=args.apply, interactive=args.apply)
        if not svc:
            print("no_sheets")
            return 2
        return fill_cand_names(svc, max(1, min(args.cap, 80)), apply=args.apply)
    svc = sheets_service(write=True)
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    sid = str(rows[0].get("sellerId") or "").strip() if rows else ""
    url_red = seller_url(sid or "X", "REDACTED")
    ok = (
        sid == "AYC4Z8PML8T30"
        and "storefront=0" in url_red
        and "storefront=1" not in url_red
    )
    line = "runId=pr_20260816_s4dry seller=%s storefront=0 GET形 %s" % (sid, "PASS" if ok else "FAIL")
    print(line)
    if not args.apply:
        append_log(svc, "S4", line)
        return 0 if ok else 1
    if not ok:
        return 1
    key = keepa_key()
    if not key:
        append_log(svc, "S4", "runId=pr_20260816_s4col FAIL no_key")
        return 2
    req = Request(seller_url(sid, key), method="GET")
    with urlopen(req, timeout=60) as resp:
        raw = resp.read()
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    body = json.loads(raw.decode("utf-8"))
    parsed = parse_seller(body, sid)
    consumed = body.get("tokensConsumed")
    left = body.get("tokensLeft")
    vok = (not parsed["_has_asinlist"]) and parsed["_n_cat"] >= 1
    print("cats", parsed["_n_cat"], "food_pct", parsed["食品比率"], "hypo", parsed["卸仮説"], "store", parsed["ストアASIN数"])
    print("consumed", consumed, "left", left, "asinList", parsed["_asin_n"])
    if not vok:
        line2 = "runId=pr_20260816_s4col FAIL cats=%s asinList=%s consumed=%s" % (
            parsed["_n_cat"],
            parsed["_asin_n"],
            consumed,
        )
        append_log(svc, "S4", line2)
        print(line2)
        return 1
    print("ensure", ensure_seller_google(COMPETITOR_SS))
    h2, _ = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    miss = [c for c in COLS if c not in h2]
    if miss:
        print("missing", miss)
        return 1
    for c in COLS:
        idx = h2.index(c)
        svc.spreadsheets().values().update(
            spreadsheetId=COMPETITOR_SS,
            range="'%s'!%s2" % (SHEET_SELLER.replace("'", "''"), col_letter(idx + 1)),
            valueInputOption="RAW",
            body={"values": [[parsed[c]]]},
        ).execute()
    if parsed["店名"]:
        idx = h2.index("店名")
        svc.spreadsheets().values().update(
            spreadsheetId=COMPETITOR_SS,
            range="'%s'!%s2" % (SHEET_SELLER.replace("'", "''"), col_letter(idx + 1)),
            valueInputOption="RAW",
            body={"values": [[parsed["店名"]]]},
        ).execute()
    h3, rows3 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    r = rows3[0]
    ok2 = bool(str(r.get("カテゴリ構成") or "").strip()) and str(r.get("卸仮説") or "").strip()
    line2 = (
        "runId=pr_20260816_s4col cats=%s food=%s store=%s hypo=%s consumed=%s asinList=0 GET1 %s"
        % (
            parsed["_n_cat"],
            r.get("食品比率"),
            r.get("ストアASIN数"),
            r.get("卸仮説"),
            consumed,
            "PASS" if ok2 else "FAIL",
        )
    )
    append_log(svc, "S4", line2)
    print(line2)
    print("VERIFY", "PASS" if ok2 else "FAIL")
    return 0 if ok2 else 1


if __name__ == "__main__":
    raise SystemExit(main())
