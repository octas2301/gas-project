# -*- coding: utf-8 -*-
"""L5: 台帳のモリタ以外1件 /query page0。productなし。"""
from __future__ import annotations

import gzip
import json
import sys
from datetime import date
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "purchase_research_path3"))

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, T_CAND, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from dry_s3_query import query_url, selection_for
from keepa_csv_vs_api import keepa_key
from schema import SHEET_SELLER

MORITA = "AYC4Z8PML8T30"


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--seller", default="")
    ap.add_argument("--page", type=int, default=0)
    args = ap.parse_args()
    svc = sheets_service(write=True)
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    tgt_i = None
    sid = ""
    pending = []
    want = str(args.seller or "").strip()
    for i, r in enumerate(rows):
        s = str(r.get("sellerId") or "").strip()
        memo = str(r.get("メモ") or "")
        if s and s != MORITA:
            if "洗わない" in memo:
                continue
            pending.append((i, s, str(r.get("巡回日") or "").strip()))
    if want:
        for i, s, day in pending:
            if s == want:
                tgt_i, sid = i, s
                break
        if not sid:
            for i, r in enumerate(rows):
                if str(r.get("sellerId") or "").strip() == want:
                    tgt_i, sid = i, want
                    break
    else:
        for i, s, day in pending:
            if not day:
                tgt_i, sid = i, s
                break
        if not sid and pending:
            tgt_i, sid, _ = pending[-1]
    sel = selection_for(sid, int(args.page)) if sid else {}
    ok = bool(sid) and sel.get("page") == int(args.page) and "storefront" not in json.dumps(sel)
    line = "runId=pr_20260816_l5dry seller=%s page=%s GET形 %s" % (sid or "-", args.page, "PASS" if ok else "FAIL")
    print(line)
    if not args.apply:
        append_log(svc, "L5", line)
        return 0 if ok else 1
    if not ok:
        return 1
    key = keepa_key()
    req = Request(query_url(sel, key), method="GET")
    with urlopen(req, timeout=60) as resp:
        raw = resp.read()
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    body = json.loads(raw.decode("utf-8"))
    asins = [str(x).upper() for x in (body.get("asinList") or []) if str(x).strip()]
    total = body.get("totalResults")
    consumed = body.get("tokensConsumed")
    left = body.get("tokensLeft")
    leak = isinstance(total, int) and total > 100000
    huge = isinstance(total, int) and total > 1000
    vok = (not leak) and isinstance(total, int) and len(asins) <= 100 and not body.get("products")
    if huge:
        line2 = "runId=pr_20260816_l5col STOP total=%s n=%s consumed=%s left=%s 1000超 人判断 page続行しない" % (
            total,
            len(asins),
            consumed,
            left,
        )
        if tgt_i is not None and "メモ" in h:
            svc.spreadsheets().values().update(
                spreadsheetId=COMPETITOR_SS,
                range="'%s'!%s%d" % (SHEET_SELLER.replace("'", "''"), col_letter(h.index("メモ") + 1), tgt_i + 2),
                valueInputOption="RAW",
                body={"values": [["洗わない total=%s 1000超" % total]]},
            ).execute()
        append_log(svc, "L5", line2)
        print(line2)
        print("VERIFY STOP huge")
        return 3
    row1 = tgt_i + 2
    today = date.today().isoformat()
    if "巡回日" in h:
        svc.spreadsheets().values().update(
            spreadsheetId=COMPETITOR_SS,
            range="'%s'!%s%d" % (SHEET_SELLER.replace("'", "''"), col_letter(h.index("巡回日") + 1), row1),
            valueInputOption="RAW",
            body={"values": [[today]]},
        ).execute()
    if "asinList件数" in h:
        svc.spreadsheets().values().update(
            spreadsheetId=COMPETITOR_SS,
            range="'%s'!%s%d" % (SHEET_SELLER.replace("'", "''"), col_letter(h.index("asinList件数") + 1), row1),
            valueInputOption="RAW",
            body={"values": [[str(total if isinstance(total, int) else len(asins))]]},
        ).execute()
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    cand_a = {str(r.get("ASIN") or "").strip().upper() for r in cand}
    in_cand = sum(1 for a in asins if a in cand_a)
    line2 = "runId=pr_20260816_l5col total=%s n=%s consumed=%s left=%s in_cand=%s productなし %s" % (
        total,
        len(asins),
        consumed,
        left,
        in_cand,
        "PASS" if vok else "FAIL",
    )
    append_log(svc, "L5", line2)
    print(line2)
    print("asins_n", len(asins))
    print("asins", ",".join(asins))
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
