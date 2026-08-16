# -*- coding: utf-8 -*-
"""W1: Keepaフルへリサーチ追記。既定は dry。出品マスタ・モールヒットは書かない。"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "purchase_research_path3"))

from client import sheets_service  # noqa: E402
from keepa_csv_vs_api import keepa_key  # noqa: E402
from keepa_full import keepa_get_needed, latest_row_for_asin, product_to_full_row, warehouse_get_needed  # noqa: E402
from schema import KEEPA_FULL_HEADERS, PURPOSE_RESEARCH, SHEET_KEEPA_FULL  # noqa: E402

RESEARCH_SS = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
T_CAND = "①候補"
T_LOG = "①ログ"
SLICE = ["B0DF7P45TL", "B084MYZWXJ", "B09CP5Y7GS", "B0H2CDC526", "B0B9GC8DTN"]
BATCH = 20
TOKEN_RESERVE = 50


def utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def read_all(svc, sid: str, title: str) -> list[list[str]]:
    rng = "'%s'" % title.replace("'", "''")
    return svc.spreadsheets().values().get(spreadsheetId=sid, range=rng).execute().get("values") or []


def as_dicts(raw: list[list[str]]) -> tuple[list[str], list[dict]]:
    if not raw:
        return [], []
    h = [str(x) for x in raw[0]]
    rows = []
    for r in raw[1:]:
        rows.append({h[i]: (str(r[i]) if i < len(r) else "") for i in range(len(h))})
    return h, rows


def cand_asins(svc) -> list[str]:
    _, rows = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    out = []
    seen = set()
    for r in rows:
        a = str(r.get("ASIN") or "").strip().upper()
        if a and a not in seen:
            seen.add(a)
            out.append(a)
    return out


def fetch_stats90(key: str, asins: list[str]) -> dict:
    import gzip

    q = urlencode({"key": key, "domain": "5", "asin": ",".join(asins), "stats": "90", "history": "0"})
    url = "https://api.keepa.com/product?" + q
    req = Request(url, headers={"User-Agent": "OctasKeepaW1/1.0", "Accept-Encoding": "identity"})
    with urlopen(req, timeout=90) as resp:
        raw = resp.read()
        if len(raw) >= 2 and raw[0] == 0x1F and raw[1] == 0x8B:
            raw = gzip.decompress(raw)
        return json.loads(raw.decode("utf-8"))


def append_rows(svc, sid: str, title: str, headers: list[str], recs: list[dict]) -> None:
    values = [[rec.get(h, "") for h in headers] for rec in recs]
    svc.spreadsheets().values().append(
        spreadsheetId=sid,
        range="'" + title.replace("'", "''") + "'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": values},
    ).execute()


def append_log(svc, step: str, line: str) -> None:
    svc.spreadsheets().values().append(
        spreadsheetId=RESEARCH_SS,
        range="'" + T_LOG + "'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [[utc_now(), step, "Keepaフル", line]]},
    ).execute()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--live", action="store_true", help="Keepa GET stats=90 history=0")
    ap.add_argument("--from-cand", action="store_true", help="①候補全ASIN（W2）")
    ap.add_argument("--limit", type=int, default=0, help="0=上限なし（スライス時は5）")
    ap.add_argument("--batch", type=int, default=BATCH)
    ap.add_argument("--run-id", default="")
    args = ap.parse_args()
    svc = sheets_service(write=True, interactive=False)
    if not svc:
        print("NO_CREDS")
        return 2
    if args.apply:
        from init_store import ensure_keepa_full_google

        print("ensure", ensure_keepa_full_google(COMPETITOR_SS))
    raw_full = read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL)
    headers, existing = as_dicts(raw_full)
    if not headers:
        headers = list(KEEPA_FULL_HEADERS)
    cands = cand_asins(svc)
    if args.from_cand:
        want = list(cands)
        step = "W2"
        run_id = args.run_id or "pr_20260815_w2"
    else:
        want = [a for a in SLICE if a in set(cands)]
        step = "W1"
        run_id = args.run_id or "pr_20260815_w1d"
    lim = args.limit if args.limit > 0 else (0 if args.from_cand else 5)
    if lim > 0:
        want = want[:lim]
    need_get = []
    skip_fresh = []
    for a in want:
        latest = latest_row_for_asin(existing, a)
        if keepa_get_needed(latest):
            need_get.append(a)
        else:
            skip_fresh.append(a)
    print(
        "cand=%d want=%d need_get=%d skip_fresh=%d apply=%s live=%s"
        % (len(cands), len(want), len(need_get), len(skip_fresh), args.apply, args.live)
    )
    print("need_get_head", ",".join(need_get[:8]))
    print("skip_fresh_n", len(skip_fresh))
    if not args.apply:
        print("dry-run no write")
        return 0
    if not need_get:
        print("nothing_to_write")
        append_log(svc, step, "runId=%s append=0 skip_fresh=%d" % (run_id, len(skip_fresh)))
        return 0
    if not args.live:
        print("refuse: --apply needs --live for Keepa GET")
        return 2
    key = keepa_key()
    if not key:
        print("NO_KEEPA_KEY")
        return 2
    fetched = utc_now()
    recs = []
    got_asins = []
    left = ""
    consumed_sum = 0
    bsz = max(1, args.batch)
    for i in range(0, len(need_get), bsz):
        chunk = need_get[i : i + bsz]
        data = fetch_stats90(key, chunk)
        left = str(data.get("tokensLeft", ""))
        cons = data.get("tokensConsumed") or 0
        try:
            consumed_sum += int(cons)
        except (TypeError, ValueError):
            pass
        products = data.get("products") or []
        print("batch %d-%d n=%d tokensLeft=%s consumed=%s" % (i + 1, i + len(chunk), len(products), left, cons))
        try:
            left_n = int(left)
        except (TypeError, ValueError):
            left_n = TOKEN_RESERVE + 1
        for p in products:
            rec = product_to_full_row(p, fetched, purpose=PURPOSE_RESEARCH)
            if not rec["ASIN"]:
                continue
            latest = latest_row_for_asin(existing, rec["ASIN"])
            if latest and str(latest.get("価格指紋") or "") == rec["価格指紋"]:
                continue
            rawj = json.loads(rec["生JSON"] or "{}")
            if "csv" in rawj:
                print("FAIL csv_in_json", rec["ASIN"])
                return 3
            recs.append(rec)
            existing.append(rec)
            got_asins.append(rec["ASIN"])
        if left_n < TOKEN_RESERVE:
            print("stop_token_reserve left=%s" % left)
            break
    if recs:
        use_h = headers if set(KEEPA_FULL_HEADERS).issubset(set(headers)) else list(KEEPA_FULL_HEADERS)
        if "目的" not in use_h:
            use_h = list(use_h) + ["目的"]
        append_rows(svc, COMPETITOR_SS, SHEET_KEEPA_FULL, use_h, recs)
    msg = "runId=%s append=%d skip_fresh=%d consumed=%s left=%s" % (
        run_id,
        len(recs),
        len(skip_fresh),
        consumed_sum,
        left,
    )
    append_log(svc, step, msg)
    print(msg)
    _, after = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    check = got_asins
    ok = 0
    for a in check:
        row = latest_row_for_asin(after, a)
        if not row:
            print("VERIFY_MISS", a)
            continue
        try:
            rawj = json.loads(row.get("生JSON") or "{}")
        except json.JSONDecodeError:
            print("VERIFY_JSON", a)
            continue
        if row.get("目的") != PURPOSE_RESEARCH:
            print("VERIFY_PURPOSE", a, row.get("目的"))
            continue
        if "csv" in rawj:
            print("VERIFY_CSV", a)
            continue
        ok += 1
    print("verify %d/%d" % (ok, len(check)))
    return 0 if check and ok == len(check) else (0 if not check else 1)


if __name__ == "__main__":
    raise SystemExit(main())
