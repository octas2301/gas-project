# -*- coding: utf-8 -*-
"""O1: Keepaフル生JSONのオファー数。GETしない。"""
from __future__ import annotations

import json
import sys

from apply_keepa_full import COMPETITOR_SS, as_dicts, read_all, append_log, sheets_service
from schema import SHEET_KEEPA_FULL

IDX_COUNT_NEW = 11


def ok(v) -> bool:
    if v is None or v == -1:
        return False
    try:
        return float(v) >= 0
    except (TypeError, ValueError):
        return False


def main() -> int:
    svc = sheets_service(write=True)
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    n = len(full)
    tot = fba = fbm = ret = cur11 = cur12 = avg11 = offers = usable = both = 0
    disagree = []
    vals_tot = []
    vals_c11 = []
    for r in full:
        try:
            p = json.loads(r.get("生JSON") or "{}")
        except json.JSONDecodeError:
            continue
        if p.get("offers") not in (None, [], {}):
            offers += 1
        st = p.get("stats") if isinstance(p.get("stats"), dict) else {}
        cur = st.get("current") or []
        avg = st.get("avg90") or []
        t, a, b, rt = st.get("totalOfferCount"), st.get("offerCountFBA"), st.get("offerCountFBM"), st.get(
            "retrievedOfferCount"
        )
        c11 = cur[IDX_COUNT_NEW] if len(cur) > IDX_COUNT_NEW else None
        c12 = cur[12] if len(cur) > 12 else None
        a11 = avg[IDX_COUNT_NEW] if len(avg) > IDX_COUNT_NEW else None
        if ok(t):
            tot += 1
            vals_tot.append(int(float(t)))
        if ok(a):
            fba += 1
        if ok(b):
            fbm += 1
        if ok(rt):
            ret += 1
        if ok(c11):
            cur11 += 1
            vals_c11.append(int(float(c11)))
        if ok(c12):
            cur12 += 1
        if ok(a11):
            avg11 += 1
        if ok(t) or ok(c11) or ok(a):
            usable += 1
        if ok(t) and ok(c11):
            both += 1
            if int(float(t)) != int(float(c11)):
                disagree.append((r.get("ASIN"), int(float(t)), int(float(c11)), a, b))
    print("n", n)
    print("offers_array", offers)
    print("totalOfferCount", tot, "/", n)
    print("offerCountFBA", fba, "FBM", fbm, "retrieved", ret)
    print("current11", cur11, "current12", cur12, "avg90_11", avg11)
    print("usable_any", usable, "/", n)
    print("both", both, "disagree_tot_vs_c11", len(disagree))
    print("disagree_head", disagree[:6])
    if vals_tot:
        s = sorted(vals_tot)
        print("tot min/med/max/zero", s[0], s[len(s) // 2], s[-1], vals_tot.count(0))
    if vals_c11:
        s = sorted(vals_c11)
        print("c11 min/med/max/zero", s[0], s[len(s) // 2], s[-1], vals_c11.count(0))
    line = (
        "runId=pr_20260815_o1dry n=%d totalOfferCount=%d/%d current11=%d usable=%d offers[]=%d disagree=%d GETなし"
        % (n, tot, n, cur11, usable, offers, len(disagree))
    )
    append_log(svc, "O1", line)
    print(line)
    return 0


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def backfill_col(svc) -> str:
    from init_store import ensure_keepa_full_google
    from keepa_full import IDX_COUNT_NEW, _keepa_current

    print("ensure", ensure_keepa_full_google(COMPETITOR_SS))
    raw = read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL)
    h, rows = as_dicts(raw)
    if "出品者数" not in h:
        return "no_col"
    idx = h.index("出品者数")
    filled = 0
    vals = []
    for r in rows:
        try:
            p = json.loads(r.get("生JSON") or "{}")
        except json.JSONDecodeError:
            vals.append([""])
            continue
        st = p.get("stats") if isinstance(p.get("stats"), dict) else {}
        v = _keepa_current(st.get("current") or [], IDX_COUNT_NEW)
        if v:
            filled += 1
        vals.append([v])
    rng = "'%s'!%s2" % (SHEET_KEEPA_FULL.replace("'", "''"), col_letter(idx + 1))
    svc.spreadsheets().values().update(
        spreadsheetId=COMPETITOR_SS,
        range=rng,
        valueInputOption="RAW",
        body={"values": vals},
    ).execute()
    line = "runId=pr_20260815_o3col filled=%d blank=%d GETなし COUNT_NEW正" % (filled, len(vals) - filled)
    append_log(svc, "O3", line)
    return line


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply-col", action="store_true")
    args = ap.parse_args()
    if args.apply_col:
        svc = sheets_service(write=True)
        print(backfill_col(svc))
        raise SystemExit(0)
    raise SystemExit(main())
