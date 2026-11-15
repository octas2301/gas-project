# -*- coding: utf-8 -*-
"""
F1 読取スパイク: 在庫・出品を読み、現在庫を再計算して突合用TSVを出す。
書込なし。適正在庫の本計算は F2。

例:
  python read_spike.py --sample 5
  python read_spike.py --jans 4901234567890,4901234567891
"""
from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "out"
sys.path.insert(0, str(ROOT))

from sheets_io import (  # noqa: E402
    inventory_spreadsheet_id,
    list_sheet_titles,
    listing_spreadsheet_id,
    read_sheet_values,
    sheets_service_readonly,
)
from stock_recalc import (  # noqa: E402
    REQUIRED_INV_SHEETS,
    InventoryCache,
    get_current_stock,
    get_last_count_date,
    load_products_from_sheet,
    norm_key,
)

JST = timezone(timedelta(hours=9))


def build_cache(svc, inv_id: str) -> InventoryCache:
    titles = set(list_sheet_titles(svc, inv_id))
    missing = [t for t in REQUIRED_INV_SHEETS if t not in titles]
    cache = InventoryCache(missing_sheets=missing)
    if missing:
        return cache
    cache.session_data = read_sheet_values(svc, inv_id, "棚卸セッション")
    cache.count_data = read_sheet_values(svc, inv_id, "棚卸スキャンリスト")
    cache.receiving_data = read_sheet_values(svc, inv_id, "入荷リスト")
    cache.shipping_data = read_sheet_values(svc, inv_id, "出荷リスト")
    products = read_sheet_values(svc, inv_id, "商品マスタ")
    cache.product_jans, cache.product_names = load_products_from_sheet(products)
    return cache


def peek_listing(svc, listing_id: str) -> dict:
    """出品マスタのヘッダーと子SKU件数だけ（算出はしない）。"""
    title = "▼商品マスタ(人間作業用)"
    rows = read_sheet_values(svc, listing_id, title)
    if not rows:
        return {"sheet": title, "headers": [], "child_rows": 0, "error": "empty"}
    headers = [str(c or "").strip() for c in rows[0]]
    i_child = next((i for i, h in enumerate(headers) if h == "子SKU"), -1)
    child_rows = 0
    if i_child >= 0:
        for r in rows[1:]:
            if i_child < len(r) and norm_key(r[i_child]):
                child_rows += 1
    return {
        "sheet": title,
        "headers": headers,
        "child_rows": child_rows,
        "has_set_col": any("セット" in h for h in headers),
        "has_jan_col": any("JAN" in h for h in headers),
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Domain4 F1 read spike (read-only)")
    ap.add_argument("--sample", type=int, default=5, help="商品マスタ先頭から何JANを計算するか")
    ap.add_argument("--jans", type=str, default="", help="カンマ区切りの単品JAN/社内キー")
    ap.add_argument("--skip-listing", action="store_true", help="出品マスタpeekを省略")
    args = ap.parse_args()

    print("mode=read_only write=never phase=F1")
    svc = sheets_service_readonly()
    if not svc:
        print("FAIL: NO_CREDS (tools/c1_hpc_packaged/secrets の token を確認)")
        return 2

    inv_id = inventory_spreadsheet_id()
    print("inventory_ss=%s" % inv_id)
    cache = build_cache(svc, inv_id)
    if cache.missing_sheets:
        print("FAIL: missing_sheets=%s" % ",".join(cache.missing_sheets))
        return 3

    last = get_last_count_date(cache)
    if last is None:
        print("FAIL: no_last_count_date (棚卸未確定)")
        return 4
    print("last_count_date=%s product_jans=%d" % (last.isoformat(), len(cache.product_jans)))

    if not args.skip_listing:
        try:
            lid = listing_spreadsheet_id()
            peek = peek_listing(svc, lid)
            print(
                "listing_peek sheet=%s child_rows=%s has_set=%s has_jan=%s"
                % (peek.get("sheet"), peek.get("child_rows"), peek.get("has_set_col"), peek.get("has_jan_col"))
            )
        except Exception as e:
            print("listing_peek_error=%s" % e)

    if args.jans.strip():
        keys = [norm_key(x) for x in args.jans.split(",") if norm_key(x)]
    else:
        keys = cache.product_jans[: max(1, args.sample)]

    if len(keys) < 1:
        print("FAIL: no_jans_to_sample")
        return 5

    ran = datetime.now(JST).strftime("%Y%m%d_%H%M%S")
    OUT.mkdir(parents=True, exist_ok=True)
    out_path = OUT / ("f1_stock_compare_%s.tsv" % ran)

    fields = [
        "ran_at",
        "stock_key",
        "product_name",
        "stock_qty",
        "last_count_date",
        "count_on_last",
        "adjust_total",
        "receiving_after",
        "shipping_after",
        "error",
        "screen_qty_blank_for_human",
        "diff_blank_for_human",
    ]
    rows_out = []
    errors = 0
    for k in keys:
        b = get_current_stock(k, cache, last)
        if b.error:
            errors += 1
        rows_out.append(
            {
                "ran_at": ran,
                "stock_key": b.stock_key,
                "product_name": cache.product_names.get(k, ""),
                "stock_qty": "" if b.stock_qty is None else b.stock_qty,
                "last_count_date": b.last_count_date,
                "count_on_last": b.count_on_last,
                "adjust_total": b.adjust_total,
                "receiving_after": b.receiving_after,
                "shipping_after": b.shipping_after,
                "error": b.error,
                "screen_qty_blank_for_human": "",
                "diff_blank_for_human": "",
            }
        )
        print(
            "JAN=%s qty=%s err=%s (count=%s adj=%s in=%s out=%s)"
            % (
                b.stock_key,
                b.stock_qty,
                b.error or "-",
                b.count_on_last,
                b.adjust_total,
                b.receiving_after,
                b.shipping_after,
            )
        )

    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter="\t")
        w.writeheader()
        w.writerows(rows_out)

    print("out=%s rows=%d errors=%d" % (out_path, len(rows_out), errors))
    print("G1b: inventory-app画面の現在庫を screen_qty 列に書き、diff=screen-stock_qty が0か確認（HUMAN_RUN）")
    print("G1c: 本コマンドは values.get のみ。更新APIは呼ばない")
    return 0 if errors == 0 else 6


if __name__ == "__main__":
    raise SystemExit(main())
