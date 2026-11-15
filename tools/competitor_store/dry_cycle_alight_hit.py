# -*- coding: utf-8 -*-
"""C1+C2 仮想: A-light（店）→ セラーヒット全店再計算。

fixture 駆動。Keepa/Catalog/SP-API は呼ばない。
正本: docs/org/B_PURCHASE_RESEARCH_CYCLE_DEV_MEMO.md §5

  python tools/competitor_store/dry_cycle_alight_hit.py
  python tools/competitor_store/dry_cycle_alight_hit.py --case T-CAH-03
  python tools/competitor_store/dry_cycle_alight_hit.py --list
"""
from __future__ import annotations

import argparse
import json
import sys
from copy import deepcopy
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "cycle_alight_hit"
sys.path.insert(0, str(ROOT))

from maker_alias import JITTAI_EXCLUDE, jitta_for_maker, resolve_alias  # noqa: E402

QUERY_TOTAL_HOLD = 901
DEFAULT_PRODUCT_CAP = 20
DEFAULT_SHOP_CAP = 1

# ヒット集計から除外（CYCLE §5 Generic 等）
HIT_EXCLUDE_NAMES = {
    "Generic",
    "generic",
    "ノーブランド品",
    "",
}


def makers_in_cell(cell: str) -> list[str]:
    out = []
    for part in str(cell or "").split("|"):
        p = resolve_alias(part.strip())
        if p:
            out.append(p)
    return out


def product_parent_name(p: dict) -> str:
    mfr = str(p.get("manufacturer") or "").strip()
    brand = str(p.get("brand") or "").strip()
    raw = mfr or brand
    return resolve_alias(raw)


def extract_makers_from_products(products: list[dict], cap: int) -> list[str]:
    seen = set()
    ordered: list[str] = []
    n = 0
    for p in products:
        if n >= cap:
            break
        parent = product_parent_name(p)
        if not parent or parent in seen:
            n += 1
            continue
        seen.add(parent)
        ordered.append(parent)
        n += 1
    return sorted(ordered)


def merge_extract_cell(old: str, new_makers: list[str]) -> str:
    seen = set()
    merged: list[str] = []
    for name in makers_in_cell(old) + new_makers:
        if not name or name in seen:
            continue
        if name in HIT_EXCLUDE_NAMES or jitta_for_maker(name) == JITTAI_EXCLUDE:
            continue
        seen.add(name)
        merged.append(name)
    return "|".join(sorted(merged))


def hit_eligible_maker(name: str) -> bool:
    if not name or name in HIT_EXCLUDE_NAMES:
        return False
    if jitta_for_maker(name) == JITTAI_EXCLUDE:
        return False
    return True


def recompute_seller_hits(sellers: list[dict]) -> dict[str, int]:
    hits: dict[str, int] = {}
    for row in sellers:
        sid = str(row.get("sellerId") or "").strip()
        if not sid:
            continue
        cell = str(row.get("抽出メーカー") or "")
        seen_in_shop: set[str] = set()
        for parent in makers_in_cell(cell):
            if not hit_eligible_maker(parent):
                continue
            if parent in seen_in_shop:
                continue
            seen_in_shop.add(parent)
            hits[parent] = hits.get(parent, 0) + 1
    return hits


def sync_maker_sum(maker_sum: list[dict], hits: dict[str, int]) -> list[dict]:
    by_name = {str(r.get("メーカー") or "").strip(): deepcopy(r) for r in maker_sum if r.get("メーカー")}
    for name, count in hits.items():
        if name not in by_name:
            by_name[name] = {"メーカー": name, "実態": jitta_for_maker(name), "セラーヒット": count}
        else:
            by_name[name]["セラーヒット"] = count
            if not str(by_name[name].get("実態") or "").strip():
                by_name[name]["実態"] = jitta_for_maker(name)
    return list(by_name.values())


def apply_alight_shop(
    seller: dict,
    mock: dict,
    product_cap: int,
    stats: dict,
) -> dict:
    out = deepcopy(seller)
    total = int(mock.get("total") or 0)
    products = mock.get("products") or []
    if not isinstance(products, list):
        products = []
    capped = products[: max(product_cap, 0)]
    new_makers = extract_makers_from_products(capped, product_cap)
    out["抽出メーカー"] = merge_extract_cell(str(out.get("抽出メーカー") or ""), new_makers)
    out["調査CK"] = False
    if total > QUERY_TOTAL_HOLD:
        out["status"] = "hold"
        stats["hold_shops"] = stats.get("hold_shops", 0) + 1
    else:
        out["status"] = "done"
        stats["done_shops"] = stats.get("done_shops", 0) + 1
    stats["product_gets"] = stats.get("product_gets", 0) + len(capped)
    stats["gate_called"] = stats.get("gate_called", False)
    stats["keepa_full_append"] = stats.get("keepa_full_append", 0)
    return out


def apply_alight_batch(
    sellers: list[dict],
    batch: dict | None,
    side_effect: dict,
) -> tuple[list[dict], dict]:
    stats = {
        "gate_called": False,
        "keepa_full_append": 0,
        "cand_rows": int(side_effect.get("cand_rows") or 0),
        "keepa_full_rows": int(side_effect.get("keepa_full_rows") or 0),
        "done_shops": 0,
        "hold_shops": 0,
        "product_gets": 0,
    }
    rows = deepcopy(sellers)
    if not batch:
        return rows, stats
    product_cap = int(batch.get("product_cap") or DEFAULT_PRODUCT_CAP)
    shop_cap = int(batch.get("shop_cap") or DEFAULT_SHOP_CAP)
    mocks = batch.get("mocks") or {}
    ids = list(batch.get("shop_seller_ids") or [])[:shop_cap]
    by_id = {str(r.get("sellerId") or "").strip(): r for r in rows}
    for sid in ids:
        sid = str(sid or "").strip()
        if not sid or sid not in by_id:
            continue
        mock = mocks.get(sid) or {"total": 0, "products": []}
        by_id[sid] = apply_alight_shop(by_id[sid], mock, product_cap, stats)
    return list(by_id.values()), stats


def load_case(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def list_cases() -> list[Path]:
    return sorted(FIX_DIR.glob("T-CAH-*.json"))


def assert_hits(actual: dict[str, int], expected: dict[str, int], case_id: str) -> list[str]:
    errs = []
    for k, v in expected.items():
        if actual.get(k) != v:
            errs.append("%s hit[%s] want=%s got=%s" % (case_id, k, v, actual.get(k)))
    for k, v in actual.items():
        if k not in expected:
            errs.append("%s hit[%s] unexpected=%s" % (case_id, k, v))
    return errs


def run_case(case: dict) -> list[str]:
    cid = str(case.get("id") or "?")
    sellers = deepcopy(case.get("sellers_before") or [])
    maker_sum = deepcopy(case.get("maker_sum_before") or [])
    side = case.get("side_effect") or {}
    batch = case.get("batch")
    expect = case.get("expect") or {}
    errs: list[str] = []

    sellers, stats = apply_alight_batch(sellers, batch, side)
    hits = recompute_seller_hits(sellers)
    maker_sum = sync_maker_sum(maker_sum, hits)

    exp_hits = expect.get("seller_hits") or {}
    errs.extend(assert_hits(hits, exp_hits, cid))

    for banned in expect.get("hits_must_not_include") or []:
        if banned in hits:
            errs.append("%s hit must not include %s" % (cid, banned))

    sa = expect.get("sellers_after") or {}
    by_id = {str(r.get("sellerId") or ""): r for r in sellers}
    for sid, want in sa.items():
        got = by_id.get(sid)
        if not got:
            errs.append("%s seller missing %s" % (cid, sid))
            continue
        for key, val in want.items():
            if key == "抽出メーカー_contains":
                cell = str(got.get("抽出メーカー") or "")
                for m in val:
                    if m not in cell:
                        errs.append("%s %s 抽出メーカー missing %s in %s" % (cid, sid, m, cell))
            elif got.get(key) != val:
                errs.append("%s %s.%s want=%r got=%r" % (cid, sid, key, val, got.get(key)))

    mj = expect.get("maker_sum_jitta") or {}
    jmap = {str(r.get("メーカー") or ""): str(r.get("実態") or "") for r in maker_sum}
    for name, jitta in mj.items():
        if jmap.get(name) != jitta:
            errs.append("%s jitta[%s] want=%s got=%s" % (cid, name, jitta, jmap.get(name)))

    if expect.get("side_effect_unchanged"):
        if stats["cand_rows"] != int(side.get("cand_rows") or 0):
            errs.append("%s cand_rows changed" % cid)
        if stats["keepa_full_rows"] != int(side.get("keepa_full_rows") or 0):
            errs.append("%s keepa_full_rows changed" % cid)

    if expect.get("gate_called") is False and stats.get("gate_called"):
        errs.append("%s gate_called must be false" % cid)
    if "keepa_full_append" in expect:
        if stats.get("keepa_full_append") != expect["keepa_full_append"]:
            errs.append(
                "%s keepa_full_append want=%s got=%s"
                % (cid, expect["keepa_full_append"], stats.get("keepa_full_append"))
            )

    return errs


def main() -> int:
    ap = argparse.ArgumentParser(description="C1+C2 A-light / セラーヒット dry")
    ap.add_argument("--case", help="T-CAH-01 等（部分一致）")
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
        paths = [p for p in paths if needle in load_case(p).get("id", "").upper() or needle in p.name.upper()]
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
    sys.exit(main())
