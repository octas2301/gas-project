# -*- coding: utf-8 -*-
"""L1: Keepa梱包 → 出品設定マスタ FBA／自己発送 first-fit。GETなし。既定 dry。"""
from __future__ import annotations

import json
import math

from apply_keepa_full import COMPETITOR_SS, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from listing_fees import keepa_pack_cm_g, parse_fba_table, parse_ship_table, pick_fba_tier, pick_self_ship
from schema import MASTER_SS_ID, SHEET_KEEPA_FULL

COLS = ("出品FBAティア", "出品FBA手数料", "自己発送サイズ", "自己発送送料", "梱包3辺合計_cm")
SETTINGS = "00_設定マスタ"


def load_settings(svc) -> tuple[list[dict], list[dict]]:
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + SETTINGS + "'!A:F")
        .execute()
        .get("values")
        or []
    )
    return parse_fba_table(vals), parse_ship_table(vals)


def one_row(p: dict, fba_tbl: list, ship_tbl: list) -> dict:
    l, w, h, g = keepa_pack_cm_g(p)
    fba = pick_fba_tier(fba_tbl, l, w, h, g)
    ship = pick_self_ship(ship_tbl, l, w, h)
    return {
        "出品FBAティア": fba["tier"],
        "出品FBA手数料": fba["fee"],
        "自己発送サイズ": ship["size"],
        "自己発送送料": ship["price"],
        "梱包3辺合計_cm": fba["sumCm"] or ship["sumCm"],
        "_fba_src": fba["source"],
        "_fba_reason": fba["reason"],
        "_ship_reason": ship["reason"],
        "_dims": all(math.isfinite(x) and x > 0 for x in (l, w, h)),
    }


def extract(full: list[dict], fba_tbl, ship_tbl) -> list[dict]:
    out = []
    for r in full:
        try:
            p = json.loads(r.get("生JSON") or "{}")
        except json.JSONDecodeError:
            p = {}
        rec = one_row(p, fba_tbl, ship_tbl)
        rec["ASIN"] = str(r.get("ASIN") or "")
        out.append(rec)
    return out


def summarize(recs: list[dict], n_fba_tbl: int, n_ship_tbl: int) -> dict:
    n = len(recs)
    dims = sum(1 for x in recs if x.get("_dims"))
    fba = sum(1 for x in recs if x.get("出品FBAティア"))
    fba_set = sum(1 for x in recs if x.get("_fba_src") == "settings")
    ship = sum(1 for x in recs if x.get("自己発送送料"))
    return {
        "n": n,
        "fba_table": n_fba_tbl,
        "ship_table": n_ship_tbl,
        "dims": dims,
        "fba_named": fba,
        "fba_settings": fba_set,
        "ship_named": ship,
    }


def dry_ok(s: dict) -> bool:
    if s["fba_table"] < 10 or s["ship_table"] < 3:
        return False
    if s["dims"] < 10:
        return False
    if s["fba_settings"] < 1 or s["ship_named"] < 1:
        return False
    if s["fba_settings"] < s["dims"] * 0.5:
        return False
    return True


def apply_cols(svc, recs: list[dict]) -> str:
    from init_store import ensure_keepa_full_google

    print("ensure", ensure_keepa_full_google(COMPETITOR_SS))
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
    s = summarize(recs, 0, 0)
    line = (
        "runId=pr_20260815_l1col n=%d dims=%d fba=%d ship=%d GETなし"
        % (len(recs), s["dims"], s["fba_named"], s["ship_named"])
    )
    append_log(svc, "L1", line)
    return line


def verify(svc) -> dict:
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    miss = [c for c in COLS if c not in h]
    filled = {c: sum(1 for r in rows if str(r.get(c) or "").strip()) for c in COLS}
    return {"missing": miss, "filled": filled, "nrows": len(rows)}


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply-col", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    fba_tbl, ship_tbl = load_settings(svc)
    h, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    recs = extract(full, fba_tbl, ship_tbl)
    s = summarize(recs, len(fba_tbl), len(ship_tbl))
    print("headers_have", {c: c in h for c in COLS})
    print("fba_names", [t["name"] for t in fba_tbl[:8]], "...", len(fba_tbl))
    print("ship_names", [t["size"] for t in ship_tbl[:8]], "...", len(ship_tbl))
    print("summary", s)
    sample = next((x for x in recs if x.get("出品FBAティア")), recs[0] if recs else None)
    print("sample", {k: sample.get(k) for k in ("ASIN",) + COLS + ("_fba_reason",)} if sample else None)
    ok = dry_ok(s)
    line = (
        "runId=pr_20260815_l1dry n=%d dims=%d fbaSet=%d/%d ship=%d tblFba=%d tblShip=%d GETなし %s"
        % (
            s["n"],
            s["dims"],
            s["fba_settings"],
            s["dims"],
            s["ship_named"],
            s["fba_table"],
            s["ship_table"],
            "PASS" if ok else "FAIL",
        )
    )
    print(line)
    if not args.apply_col:
        append_log(svc, "L1", line)
        return 0 if ok else 1
    if not ok:
        print("skip apply: dry FAIL")
        return 1
    print(apply_cols(svc, recs))
    v = verify(svc)
    print("verify", json.dumps(v, ensure_ascii=False))
    vok = not v["missing"] and v["filled"].get("出品FBAティア", 0) == s["fba_named"]
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
