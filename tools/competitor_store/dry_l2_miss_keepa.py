# -*- coding: utf-8 -*-
"""L2: Keepaフルに無い ASIN を stats=90 で倉庫へ。最大20/回。品番リスト非書。
既定は品番リストの miss。--query-seller ならその店の /query 1ページの miss。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "purchase_research_path3"))

from apply_keepa_full import (
    COMPETITOR_SS,
    RESEARCH_SS,
    TOKEN_RESERVE,
    append_log,
    append_rows,
    as_dicts,
    fetch_stats90,
    read_all,
    sheets_service,
    utc_now,
)
from dry_s3_query import query_url, selection_for
from dry_t1_list_paste import DATA_START, HEADER_ROW, LIST_TITLE, ORIG_SS
from keepa_csv_vs_api import keepa_key
from keepa_full import keepa_get_needed, latest_row_for_asin, product_to_full_row
from schema import KEEPA_FULL_HEADERS, PURPOSE_RESEARCH, SHEET_KEEPA_FULL

CAP = 20


def list_asins(svc) -> list[str]:
    rng = "'%s'!%d:%d" % (LIST_TITLE.replace("'", "''"), HEADER_ROW, HEADER_ROW)
    headers = [
        str(x).replace("\n", " ").strip()
        for x in (svc.spreadsheets().values().get(spreadsheetId=RESEARCH_SS, range=rng).execute().get("values") or [[]])[0]
    ]
    asin_i = headers.index("ASIN")
    raw = read_all(svc, RESEARCH_SS, LIST_TITLE)
    out, seen = [], set()
    for r in raw[DATA_START - 1 :]:
        a = str(r[asin_i] if asin_i < len(r) else "").strip().upper()
        if len(a) == 10 and a not in seen:
            seen.add(a)
            out.append(a)
    return out


def query_page_asins(seller: str, page: int) -> tuple[list[str], int]:
    import gzip
    from urllib.request import Request, urlopen

    key = keepa_key()
    sel = selection_for(seller, page)
    req = Request(query_url(sel, key), method="GET")
    with urlopen(req, timeout=60) as resp:
        raw = resp.read()
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    body = json.loads(raw.decode("utf-8"))
    asins = [str(x).upper() for x in (body.get("asinList") or []) if str(x).strip()]
    total = int(body.get("totalResults") or 0)
    return asins, total


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--query-seller", default="")
    ap.add_argument("--page", type=int, default=0)
    ap.add_argument("--once", action="store_true", help="Keepa は先頭20件だけ。残りは次フェーズ")
    ap.add_argument("--run-id", default="pr_20260816_l2")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    if RESEARCH_SS == ORIG_SS:
        print("refuse original")
        return 2
    if args.query_seller:
        asins, total = query_page_asins(args.query_seller, args.page)
        print("query total=%s n=%s seller=%s page=%s" % (total, len(asins), args.query_seller, args.page))
        if total > 1000:
            print("STOP huge")
            return 3
    else:
        asins = list_asins(svc)
    fh, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    have = {str(r.get("ASIN") or "").strip().upper() for r in full}
    miss = [a for a in asins if a not in have]
    need = [a for a in miss if keepa_get_needed(latest_row_for_asin(full, a))]
    chunk = need[:CAP]
    ok = RESEARCH_SS != ORIG_SS and len(chunk) <= CAP
    rid = args.run_id + ("dry" if not args.apply else "col")
    line = "runId=%s src=%s list=%d miss=%d need=%d chunk=%d %s %s" % (
        rid,
        ("query:%s p%s" % (args.query_seller, args.page)) if args.query_seller else "list",
        len(asins),
        len(miss),
        len(need),
        len(chunk),
        "productGETなし" if not args.apply else "once" if args.once else "loop",
        "PASS" if ok else "FAIL",
    )
    print(line)
    print("chunk_head", ",".join(chunk[:8]))
    if not args.apply:
        append_log(svc, "L2", line)
        return 0 if ok else 1
    if not ok:
        return 1
    if not chunk:
        append_log(svc, "L2", "runId=%scol append=0 miss=0" % args.run_id)
        print("nothing")
        return 0
    if not args.live:
        print("need --live")
        return 2
    key = keepa_key()
    if not key:
        return 2
    fetched = utc_now()
    remaining = list(chunk if args.once else need)
    appended = 0
    consumed_sum = 0
    left = ""
    while remaining:
        batch = remaining[:CAP]
        remaining = remaining[CAP:]
        data = fetch_stats90(key, batch)
        left = str(data.get("tokensLeft", ""))
        cons = data.get("tokensConsumed") or 0
        try:
            consumed_sum += int(cons)
        except (TypeError, ValueError):
            pass
        recs = []
        live_h = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))[0] or list(KEEPA_FULL_HEADERS)
        for p in data.get("products") or []:
            rec = product_to_full_row(p, fetched, purpose=PURPOSE_RESEARCH)
            if not rec["ASIN"]:
                continue
            rawj = json.loads(rec["生JSON"] or "{}")
            if "csv" in rawj:
                print("FAIL csv")
                return 3
            recs.append(rec)
        if recs:
            append_rows(svc, COMPETITOR_SS, SHEET_KEEPA_FULL, live_h, recs)
            appended += len(recs)
        try:
            left_n = int(left)
        except (TypeError, ValueError):
            left_n = TOKEN_RESERVE + 1
        if left_n < TOKEN_RESERVE:
            print("stop_token", left)
            break
        if args.once or not remaining:
            break
    _, after = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    have2 = {str(r.get("ASIN") or "").strip().upper() for r in after}
    vok = all(a in have2 for a in chunk) if chunk else True
    rest = sum(1 for a in need if a not in have2)
    line2 = "runId=%scol append=%d consumed=%s left=%s chunk=%d/%d rest_need=%d 品番非書 %s" % (
        args.run_id,
        appended,
        consumed_sum,
        left,
        sum(1 for a in chunk if a in have2),
        len(chunk),
        rest,
        "PASS" if vok else "FAIL",
    )
    append_log(svc, "L2", line2)
    print(line2)
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
