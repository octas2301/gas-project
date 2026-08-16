# -*- coding: utf-8 -*-
"""F0: Keepaフル生JSON flatten。GETしない。既定 dry。"""
from __future__ import annotations

import json

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from keepa_full import flatten_from_product
from schema import SHEET_KEEPA_FULL

COLS = ("Amazon直販", "新品: 現在価格", "売れ筋ランキング: 現在", "Amazon: 180日在庫切れ%")


def extract_rows(full: list[dict]) -> list[dict]:
    out = []
    for r in full:
        try:
            p = json.loads(r.get("生JSON") or "{}")
        except json.JSONDecodeError:
            p = {}
        rec = flatten_from_product(p)
        rec["ASIN"] = str(r.get("ASIN") or "")
        out.append(rec)
    return out


def summarize(recs: list[dict]) -> dict:
    n = len(recs)
    iru = sum(1 for x in recs if x.get("Amazon直販") == "いる")
    inai = sum(1 for x in recs if x.get("Amazon直販") == "いない")
    blank_d = sum(1 for x in recs if not x.get("Amazon直販"))
    newp = sum(1 for x in recs if x.get("新品: 現在価格"))
    rank = sum(1 for x in recs if x.get("売れ筋ランキング: 現在"))
    oos = sum(1 for x in recs if x.get("Amazon: 180日在庫切れ%"))
    return {
        "n": n,
        "direct_iru": iru,
        "direct_inai": inai,
        "direct_blank": blank_d,
        "new_price": newp,
        "rank_now": rank,
        "oos180": oos,
    }


def dry_ok(s: dict) -> bool:
    if s["n"] < 1:
        return False
    if s["direct_iru"] + s["direct_inai"] + s["direct_blank"] != s["n"]:
        return False
    if s["direct_blank"] > 0:
        return False
    if s["new_price"] < 1:
        return False
    if s["rank_now"] < 1:
        return False
    if s["oos180"] < 1:
        return False
    return True


def apply_cols(svc) -> str:
    from init_store import ensure_keepa_full_google

    print("ensure", ensure_keepa_full_google(COMPETITOR_SS))
    raw = read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL)
    h, rows = as_dicts(raw)
    recs = extract_rows(rows)
    missing = [c for c in COLS if c not in h]
    if missing:
        return "missing_cols " + ",".join(missing)
    n = len(rows)
    for c in COLS:
        idx = h.index(c)
        vals = [[r.get(c) or ""] for r in recs]
        rng = "'%s'!%s2" % (SHEET_KEEPA_FULL.replace("'", "''"), col_letter(idx + 1))
        svc.spreadsheets().values().update(
            spreadsheetId=COMPETITOR_SS,
            range=rng,
            valueInputOption="RAW",
            body={"values": vals},
        ).execute()
    s = summarize(recs)
    line = (
        "runId=pr_20260815_f0col n=%d いる=%d いない=%d 新品現在=%d 順位現在=%d oos180=%d GETなし"
        % (s["n"], s["direct_iru"], s["direct_inai"], s["new_price"], s["rank_now"], s["oos180"])
    )
    append_log(svc, "F0", line)
    return line


def verify(svc) -> dict:
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    miss = [c for c in COLS if c not in h]
    s = summarize(extract_rows(rows))
    filled = {}
    for c in COLS:
        filled[c] = sum(1 for r in rows if str(r.get(c) or "").strip())
    return {"missing": miss, "sheet_filled": filled, "from_json": s, "nrows": len(rows)}


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply-col", action="store_true")
    ap.add_argument("--verify", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    if args.verify:
        v = verify(svc)
        print(json.dumps(v, ensure_ascii=False))
        ok = not v["missing"] and v["sheet_filled"].get("Amazon直販") == v["from_json"]["n"]
        print("VERIFY", "PASS" if ok else "FAIL")
        return 0 if ok else 1
    h, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    recs = extract_rows(full)
    s = summarize(recs)
    print("headers_have", {c: c in h for c in COLS})
    print("summary", s)
    print("sample", recs[0] if recs else None)
    ok = dry_ok(s)
    line = (
        "runId=pr_20260815_f0dry n=%d いる=%d いない=%d blank=%d 新品=%d 順位=%d oos180=%d GETなし %s"
        % (
            s["n"],
            s["direct_iru"],
            s["direct_inai"],
            s["direct_blank"],
            s["new_price"],
            s["rank_now"],
            s["oos180"],
            "PASS" if ok else "FAIL",
        )
    )
    print(line)
    if not args.apply_col:
        append_log(svc, "F0", line)
        return 0 if ok else 1
    if not ok:
        print("skip apply: dry FAIL")
        return 1
    print(apply_cols(svc))
    v = verify(svc)
    print("verify", json.dumps(v, ensure_ascii=False))
    vok = not v["missing"] and v["sheet_filled"].get("Amazon直販") == v["from_json"]["n"]
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
