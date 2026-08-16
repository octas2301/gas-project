# -*- coding: utf-8 -*-
"""永谷園あさげ 先頭20件: Keepaフル＋門＋①候補。モールヒット非書。既定 dry。"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from apply_keepa_full import (
    COMPETITOR_SS,
    RESEARCH_SS,
    T_CAND,
    append_log,
    append_rows,
    as_dicts,
    fetch_stats90,
    keepa_key,
    read_all,
    sheets_service,
    utc_now,
)
from dry_w3_gate import gate, read_profile
from keepa_full import keepa_get_needed, latest_row_for_asin, product_to_full_row
from schema import KEEPA_FULL_HEADERS, PURPOSE_RESEARCH, SHEET_HITS, SHEET_KEEPA_FULL

CAP = 20


def slice_asins(catalog: Path, needles: list[str]) -> list[str]:
    items = json.loads(catalog.read_text(encoding="utf-8")).get("items") or []
    out = []
    seen = set()
    for r in items:
        title = str(r.get("title") or "") + " " + str(r.get("brand") or "")
        a = str(r.get("asin") or "").strip().upper()
        if not a or a in seen:
            continue
        if needles and all(n in title for n in needles):
            seen.add(a)
            out.append(a)
        if len(out) >= CAP:
            break
    return out


def cand_map(rows: list[dict]) -> dict[str, int]:
    m = {}
    for i, r in enumerate(rows):
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            m[a] = i
    return m


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--json", default="")
    ap.add_argument("--needles", default="永谷園,あさげ")
    ap.add_argument("--maker", default="永谷園")
    ap.add_argument("--run-id", default="pr_20260815_naga20")
    args = ap.parse_args()
    catalog = Path(args.json) if args.json else (
        Path(__file__).resolve().parents[1]
        / "purchase_research_path3"
        / "out"
        / "PATH3_SPAPI_20260815_115012.json"
    )
    needles = [x for x in str(args.needles).split(",") if x]
    want = slice_asins(catalog, needles)
    svc = sheets_service(write=True)
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    hits_n = max(0, len(read_all(svc, COMPETITOR_SS, SHEET_HITS)) - 1)
    need, skip = [], []
    for a in want:
        if keepa_get_needed(latest_row_for_asin(full, a)):
            need.append(a)
        else:
            skip.append(a)
    print("want=%d need_get=%d skip=%d hits=%d" % (len(want), len(need), len(skip), hits_n))
    print("head", ",".join(want[:5]))
    if not args.apply:
        print("dry no write")
        return 0
    if not args.live:
        print("need --live")
        return 2
    key = keepa_key()
    fetched = utc_now()
    recs = []
    products_by = {}
    data = fetch_stats90(key, need) if need else {"products": [], "tokensConsumed": 0, "tokensLeft": ""}
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
        recs.append(rec)
        full.append(rec)
    if recs:
        live_h = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))[0] or list(KEEPA_FULL_HEADERS)
        append_rows(svc, COMPETITOR_SS, SHEET_KEEPA_FULL, live_h, recs)
    # skip_fresh products from existing json
    _, full2 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    for a in skip:
        row = latest_row_for_asin(full2, a)
        if not row:
            continue
        try:
            products_by[a] = json.loads(row.get("生JSON") or "{}")
        except json.JSONDecodeError:
            pass
    price_min, rank_max = read_profile(svc)
    ch, crow = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    idx = {h: i for i, h in enumerate(ch)}
    by = cand_map(crow)
    new_rows = []
    pass_n = drop_n = 0
    for a in want:
        p = products_by.get(a) or {}
        st, why = gate(p, price_min, rank_max)
        if st == "通過":
            pass_n += 1
        else:
            drop_n += 1
        title = str(p.get("title") or "")
        eans = p.get("eanList") or []
        rec = {
            "メーカー": args.maker,
            "ASIN": a,
            "商品名": title,
            "税込価格": "",
            "順位90": "",
            "JAN": str(eans[0] if eans else p.get("ean") or ""),
            "発見経路": "catalog_title",
            "sellerId": "",
            "門結果": st,
            "門理由": why,
            "runId": args.run_id,
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
        "runId=%s cap=20 append_full=%s consumed=%s left=%s pass=%s drop=%s new_cand=%s hits=%s"
        % (args.run_id, len(recs), consumed, left, pass_n, drop_n, len(new_rows), hits_after)
    )
    append_log(svc, "N20", msg)
    print(msg)
    if hits_after != hits_n:
        print("FAIL hits changed")
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())
