# -*- coding: utf-8 -*-
"""C0 仮想: 門通過 ASIN → Keepaフル BB 列 + SP-API 新品モック → セラー台帳。

fixture 駆動。Keepa GET / SP-API / スプシは呼ばない。
正本: docs/org/B_PURCHASE_RESEARCH_CYCLE_DEV_MEMO.md §5

  python tools/competitor_store/dry_cycle_seller_accum.py
  python tools/competitor_store/dry_cycle_seller_accum.py --case T-CSA-03
  python tools/competitor_store/dry_cycle_seller_accum.py --list
"""
from __future__ import annotations

import argparse
import json
import sys
from copy import deepcopy
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "cycle_seller_accum"
sys.path.insert(0, str(ROOT))

CAP_NEW_DEFAULT = 100
SKIP_IDS = {"ATVPDKIKX0DER", "A1VC38T7YXB528"}
ACCUM_SRC = "bb+spapi_new"
ACCUM_MEMO = "C0貯め 構成比は店洗い時"
CAND_SPAPI_DONE = "セラー調査済みSP-API"
CAND_SPAPI_AT = "セラー調査日時"


def ok_sid(s: str) -> bool:
    s = (s or "").strip()
    if len(s) < 10 or s in SKIP_IDS:
        return False
    return True


def cand_spapi_done(r: dict) -> bool:
    v = r.get(CAND_SPAPI_DONE)
    if v is True or v == 1:
        return True
    s = str(v or "").strip().upper()
    return s in ("TRUE", "1", "YES", "ON", "済")


def pass_asin_list(cand: list[dict], *, skip_spapi_done: bool = True) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for r in cand:
        if str(r.get("門結果") or "") != "通過":
            continue
        if skip_spapi_done and cand_spapi_done(r):
            continue
        a = str(r.get("ASIN") or "").strip().upper()
        if len(a) == 10 and a not in seen:
            seen.add(a)
            out.append(a)
    return out


def mark_cand_spapi_done(cand: list[dict], asin: str) -> bool:
    a = str(asin or "").strip().upper()
    if len(a) != 10:
        return False
    for r in cand:
        if str(r.get("ASIN") or "").strip().upper() != a:
            continue
        r[CAND_SPAPI_DONE] = True
        r[CAND_SPAPI_AT] = r.get(CAND_SPAPI_AT) or "2026-08-22T00:00:00"
        return True
    return False


def bb_map_for_pass(keepa_full: list[dict], pass_asins: list[str]) -> dict[str, str]:
    want = set(pass_asins)
    out: dict[str, str] = {}
    for r in keepa_full:
        a = str(r.get("ASIN") or "").strip().upper()
        if a not in want:
            continue
        s = str(r.get("BuyBoxセラー") or "").strip()
        if ok_sid(s):
            out[a] = s
    return out


def run_seller_accum(
    sellers_before: list[dict],
    cand: list[dict],
    keepa_full: list[dict],
    *,
    cap: int = CAP_NEW_DEFAULT,
    after_asin: str = "",
    offers_by_asin: dict[str, list[str]] | None = None,
    stop_at_asin: str = "",
    stop_reason: str = "429",
    side_effect: dict | None = None,
) -> tuple[list[dict], list[dict], dict]:
    """GAS prRunSellerAccum_ / dry_a1 と同型（単品 getItemOffers モード）。"""
    side = side_effect or {}
    stats = {
        "keepa_get": 0,
        "cand_rows": int(side.get("cand_rows") or 0),
        "keepa_full_rows": int(side.get("keepa_full_rows") or 0),
        "api_calls": 0,
        "stopped": "",
        "after_asin_out": "",
        "new_ids": [],
        "marked_asins": [],
        "skip_http": 0,
        "skip_http_asins": [],
    }
    offers_by_asin = offers_by_asin or {}
    cand_after = deepcopy(cand)
    pass_all = pass_asin_list(cand_after, skip_spapi_done=False)
    asins = pass_asin_list(cand_after, skip_spapi_done=True)
    pass_pending = len(asins)
    after = str(after_asin or "").strip().upper()
    if after:
        try:
            idx = asins.index(after)
            asins = asins[idx + 1 :]
        except ValueError:
            pass

    have = {
        str(r.get("sellerId") or "").strip().upper()
        for r in sellers_before
        if str(r.get("sellerId") or "").strip()
    }
    bb = bb_map_for_pass(keepa_full, pass_all)
    new_ids: list[str] = []
    last_a = ""

    for a in asins:
        if len(new_ids) >= cap:
            stats["stopped"] = "cap"
            break
        last_a = a
        stats["api_calls"] += 1

        if stop_at_asin and a.upper() == str(stop_at_asin).strip().upper() and stop_reason == "429":
            stats["stopped"] = "429"
            stats["after_asin_out"] = a
            break

        ids: list[str] = []
        if a in bb:
            ids.append(bb[a])
        http_skip = (
            bool(stop_at_asin)
            and a.upper() == str(stop_at_asin).strip().upper()
            and stop_reason == "http"
        )
        if http_skip:
            stats["skip_http"] += 1
            stats["skip_http_asins"].append(a)
        else:
            ids.extend(offers_by_asin.get(a) or offers_by_asin.get(a.upper()) or [])

        for s in ids:
            if not ok_sid(s):
                continue
            key = s.strip().upper()
            if key in have:
                continue
            have.add(key)
            new_ids.append(s.strip())
            if len(new_ids) >= cap:
                stats["stopped"] = "cap"
                break
        if mark_cand_spapi_done(cand_after, a):
            stats["marked_asins"].append(a)
        if stats["stopped"] == "cap":
            break

    if not stats["stopped"]:
        stats["stopped"] = "ok"
        stats["after_asin_out"] = ""
    stats["new_ids"] = new_ids
    stats["last_asin"] = last_a
    stats["pass_all"] = len(pass_all)
    stats["pass_pending"] = pass_pending
    stats["remain_asin"] = max(0, pass_pending - len(stats["marked_asins"]))

    sellers_after = deepcopy(sellers_before)
    for s in new_ids:
        sellers_after.append(
            {
                "調査CK": False,
                "sellerId": s,
                "採取元": ACCUM_SRC,
                "メモ": ACCUM_MEMO,
            }
        )
    return sellers_after, cand_after, stats


def load_case(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def list_cases() -> list[Path]:
    return sorted(FIX_DIR.glob("T-CSA-*.json"))


def run_case(case: dict) -> list[str]:
    cid = str(case.get("id") or "?")
    run = case.get("run") or {}
    expect = case.get("expect") or {}
    side = case.get("side_effect") or {}
    errs: list[str] = []

    sellers_after, cand_after, stats = run_seller_accum(
        case.get("sellers_before") or [],
        case.get("cand") or [],
        case.get("keepa_full") or [],
        cap=int(run.get("cap") or CAP_NEW_DEFAULT),
        after_asin=str(run.get("after_asin") or ""),
        offers_by_asin=run.get("offers_by_asin") or {},
        stop_at_asin=str(run.get("stop_at_asin") or ""),
        stop_reason=str(run.get("stop_reason") or "429"),
        side_effect=side,
    )

    exp_new = expect.get("new_seller_ids") or []
    if list(stats["new_ids"]) != exp_new:
        errs.append("%s new_seller_ids want=%s got=%s" % (cid, exp_new, stats["new_ids"]))

    if expect.get("stopped") is not None and stats["stopped"] != expect["stopped"]:
        errs.append("%s stopped want=%s got=%s" % (cid, expect["stopped"], stats["stopped"]))

    if "after_asin" in expect and stats["after_asin_out"] != expect["after_asin"]:
        errs.append(
            "%s after_asin want=%r got=%r" % (cid, expect["after_asin"], stats["after_asin_out"])
        )

    if "api_calls" in expect and stats["api_calls"] != expect["api_calls"]:
        errs.append("%s api_calls want=%s got=%s" % (cid, expect["api_calls"], stats["api_calls"]))

    if expect.get("keepa_get") is not None and stats["keepa_get"] != expect["keepa_get"]:
        errs.append("%s keepa_get want=%s got=%s" % (cid, expect["keepa_get"], stats["keepa_get"]))

    if "marked_asins" in expect and list(stats["marked_asins"]) != expect["marked_asins"]:
        errs.append(
            "%s marked_asins want=%s got=%s"
            % (cid, expect["marked_asins"], stats["marked_asins"])
        )

    if "remain_asin" in expect and stats["remain_asin"] != expect["remain_asin"]:
        errs.append(
            "%s remain_asin want=%s got=%s" % (cid, expect["remain_asin"], stats["remain_asin"])
        )

    if "skip_http" in expect and stats["skip_http"] != expect["skip_http"]:
        errs.append("%s skip_http want=%s got=%s" % (cid, expect["skip_http"], stats["skip_http"]))
    if "skip_http_asins" in expect and list(stats["skip_http_asins"]) != expect["skip_http_asins"]:
        errs.append(
            "%s skip_http_asins want=%s got=%s"
            % (cid, expect["skip_http_asins"], stats["skip_http_asins"])
        )

    for a in expect.get("must_stay_unmarked") or []:
        row = next(
            (
                r
                for r in cand_after
                if str(r.get("ASIN") or "").strip().upper() == str(a).strip().upper()
            ),
            None,
        )
        if row and cand_spapi_done(row):
            errs.append("%s must_stay_unmarked but marked %s" % (cid, a))

    by_id = {str(r.get("sellerId") or ""): r for r in sellers_after}
    for sid, want in (expect.get("new_rows") or {}).items():
        got = by_id.get(sid)
        if not got:
            errs.append("%s new row missing %s" % (cid, sid))
            continue
        for key, val in want.items():
            if key == "メモ_contains":
                cell = str(got.get("メモ") or "")
                if val not in cell:
                    errs.append("%s %s メモ missing %r in %r" % (cid, sid, val, cell))
            elif got.get(key) != val:
                errs.append("%s %s.%s want=%r got=%r" % (cid, sid, key, val, got.get(key)))

    for sid in expect.get("must_not_add") or []:
        if sid not in {str(r.get("sellerId") or "") for r in case.get("sellers_before") or []} and sid in by_id:
            if sid in exp_new:
                continue
            errs.append("%s must_not_add but present %s" % (cid, sid))
        elif sid in exp_new:
            errs.append("%s must_not_add listed in new_seller_ids %s" % (cid, sid))

    before_n = len(case.get("sellers_before") or [])
    after_n = len(sellers_after)
    if "seller_row_delta" in expect and after_n - before_n != expect["seller_row_delta"]:
        errs.append(
            "%s seller_row_delta want=%s got=%s"
            % (cid, expect["seller_row_delta"], after_n - before_n)
        )

    if expect.get("side_effect_unchanged"):
        if stats["cand_rows"] != int(side.get("cand_rows") or 0):
            errs.append("%s cand_rows changed" % cid)
        if stats["keepa_full_rows"] != int(side.get("keepa_full_rows") or 0):
            errs.append("%s keepa_full_rows changed" % cid)

    return errs


def main() -> int:
    ap = argparse.ArgumentParser(description="C0 セラー貯め dry（fixture）")
    ap.add_argument("--case", help="T-CSA-01 等（部分一致）")
    ap.add_argument("--list", action="store_true", help="ケース一覧")
    args = ap.parse_args()

    if args.list:
        for p in list_cases():
            c = load_case(p)
            print("%s\t%s\t%s" % (c.get("id"), p.name, c.get("desc", "")))
        return 0

    paths = list_cases()
    if args.case:
        needle = args.case.strip().upper()
        paths = [
            p
            for p in paths
            if needle in load_case(p).get("id", "").upper() or needle in p.name.upper()
        ]
        if not paths:
            print("no case match:", args.case, file=sys.stderr)
            return 2

    all_errs: list[str] = []
    passed = 0
    for path in paths:
        case = load_case(path)
        errs = run_case(case)
        if errs:
            print("FAIL", case.get("id"), path.name)
            for e in errs:
                print(" ", e)
                all_errs.append(e)
        else:
            print("PASS", case.get("id"), case.get("desc", ""))
            passed += 1

    total = len(paths)
    print("---")
    print("PASS %s/%s" % (passed, total))
    return 0 if not all_errs else 1


if __name__ == "__main__":
    raise SystemExit(main())
