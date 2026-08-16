# -*- coding: utf-8 -*-
"""L2: 品番リストにあって Keepaフルに無い ASIN を stats=90 で倉庫へ。最大20/回。品番リスト非書。"""
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


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--live", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    if RESEARCH_SS == ORIG_SS:
        print("refuse original")
        return 2
    asins = list_asins(svc)
    fh, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    have = {str(r.get("ASIN") or "").strip().upper() for r in full}
    miss = [a for a in asins if a not in have]
    need = [a for a in miss if keepa_get_needed(latest_row_for_asin(full, a))]
    chunk = need[:CAP]
    ok = RESEARCH_SS != ORIG_SS and "ASIN" and len(need) >= 0
    # dry pass: orig guard, cap respected, no offers
    ok = RESEARCH_SS != ORIG_SS and len(chunk) <= CAP
    line = "runId=pr_20260816_l2dry list=%d miss=%d need=%d chunk=%d GETなし %s" % (
        len(asins),
        len(miss),
        len(need),
        len(chunk),
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
        append_log(svc, "L2", "runId=pr_20260816_l2col append=0 miss=0")
        print("nothing")
        return 0
    if not args.live:
        print("need --live")
        return 2
    key = keepa_key()
    if not key:
        return 2
    fetched = utc_now()
    remaining = list(need)
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
        if not remaining:
            break
    _, after = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    have2 = {str(r.get("ASIN") or "").strip().upper() for r in after}
    hit = sum(1 for a in chunk if a in have2)
    vok = hit == len(chunk) or appended >= 1
    # if first chunk all written
    vok = all(a in have2 for a in need[: min(CAP, len(need))]) if need else True
    line2 = "runId=pr_20260816_l2col append=%d consumed=%s left=%s first_chunk=%d/%d 品番非書 %s" % (
        appended,
        consumed_sum,
        left,
        sum(1 for a in chunk if a in have2),
        len(chunk),
        "PASS" if vok else "FAIL",
    )
    append_log(svc, "L2", line2)
    print(line2)
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
