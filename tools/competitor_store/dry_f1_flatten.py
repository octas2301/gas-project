# -*- coding: utf-8 -*-
"""F1: Keepaフル既存列を生JSONから展開。GETしない。既定 dry。出品FBA列は触らない。"""
from __future__ import annotations

import json

from apply_keepa_full import COMPETITOR_SS, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from keepa_full import flatten_keepa_display
from schema import SHEET_KEEPA_FULL

COLS = (
    "Buy Box: 現在価格",
    "Buy Box: 30 日平均",
    "Buy Box: 90 日平均",
    "レビュー: 評価",
    "レビュー: 評価件数",
    "カテゴリ: ルート",
    "カテゴリ: ツリー",
    "梱包_L_cm",
    "梱包_W_cm",
    "梱包_H_cm",
    "梱包_重量_g",
    "FBA手数料",
    "BuyBoxセラー",
    "BuyBox_FBA",
)


def extract(full: list[dict]) -> list[dict]:
    out = []
    for r in full:
        try:
            p = json.loads(r.get("生JSON") or "{}")
        except json.JSONDecodeError:
            p = {}
        rec = flatten_keepa_display(p)
        rec["ASIN"] = str(r.get("ASIN") or "")
        out.append(rec)
    return out


def summarize(recs: list[dict]) -> dict:
    n = len(recs)
    filled = {c: sum(1 for x in recs if str(x.get(c) or "").strip()) for c in COLS}
    filled["n"] = n
    return filled


def dry_ok(s: dict) -> bool:
    if s.get("n", 0) < 1:
        return False
    if s.get("カテゴリ: ルート", 0) < 1:
        return False
    if s.get("梱包_L_cm", 0) < 1:
        return False
    return True


def apply_cols(svc, recs: list[dict]) -> str:
    raw = read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL)
    h, rows = as_dicts(raw)
    missing = [c for c in COLS if c not in h]
    if missing:
        return "missing_cols " + ",".join(missing)
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
    line = "runId=pr_20260815_f1col n=%d cat=%d pack=%d fba=%d bb90=%d GETなし" % (
        s["n"],
        s["カテゴリ: ルート"],
        s["梱包_L_cm"],
        s["FBA手数料"],
        s["Buy Box: 90 日平均"],
    )
    append_log(svc, "F1", line)
    return line


def verify(svc, recs: list[dict]) -> dict:
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    miss = [c for c in COLS if c not in h]
    filled = {c: sum(1 for r in rows if str(r.get(c) or "").strip()) for c in COLS}
    expect = summarize(recs)
    return {"missing": miss, "filled": filled, "expect": expect, "nrows": len(rows)}


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply-col", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    h, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    recs = extract(full)
    s = summarize(recs)
    print("headers_have", {c: c in h for c in COLS})
    print("summary", s)
    print("sample", recs[0] if recs else None)
    ok = dry_ok(s)
    line = "runId=pr_20260815_f1dry n=%d cat=%d packL=%d fbaFee=%d bbNow=%d bb90=%d seller=%d GETなし %s" % (
        s["n"],
        s["カテゴリ: ルート"],
        s["梱包_L_cm"],
        s["FBA手数料"],
        s["Buy Box: 現在価格"],
        s["Buy Box: 90 日平均"],
        s["BuyBoxセラー"],
        "PASS" if ok else "FAIL",
    )
    print(line)
    if not args.apply_col:
        append_log(svc, "F1", line)
        return 0 if ok else 1
    if not ok:
        print("skip apply dry FAIL")
        return 1
    print(apply_cols(svc, recs))
    v = verify(svc, recs)
    print("verify filled", v["filled"])
    vok = not v["missing"] and v["filled"].get("カテゴリ: ルート") == s["カテゴリ: ルート"]
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
