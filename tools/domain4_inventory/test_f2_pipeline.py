# -*- coding: utf-8 -*-
"""F2パイプラインのオフライン検証（Sheets不要）。"""
from __future__ import annotations

from formulas import available_qty, parse_set_count, proper_qty, safety_stock
from listing_read import iter_child_rows
from stock_recalc import InventoryCache, get_current_stock, norm_key


def test_end_to_end_single_ok_and_skips():
    rows = [
        ["親SKU", "子SKU", "JANコード", "A.セット商品数", "商品名ベース", "セット構成1_親SKU"],
        ["P1", "", "4901", "", "親", ""],
        ["P1", "C-8", "4901", 8, "子8", ""],
        ["P2", "C-M", "4902", 2, "ミックス", "OTHER"],
        ["P3", "C-X", "9999", 3, "名簿なし", ""],
    ]
    children, _ = iter_child_rows(rows)
    assert len(children) == 3

    # 在庫: 4901 が 100
    from datetime import datetime, timedelta, timezone

    JST = timezone(timedelta(hours=9))

    def serial(dt):
        epoch = datetime(1899, 12, 30, tzinfo=JST)
        return (dt - epoch).total_seconds() / 86400.0

    cache = InventoryCache(
        session_data=[
            ["セッションID", "完了日時"],
            ["S1", serial(datetime(2026, 11, 10, 18, 0, tzinfo=JST))],
        ],
        count_data=[
            ["日時", "単品JAN", "数量", "種別", "セッションID"],
            [serial(datetime(2026, 11, 10, 12, 0, tzinfo=JST)), "4901", 100, "", "S1"],
        ],
        receiving_data=[["日時", "単品JAN", "数量"]],
        shipping_data=[["日時", "単品JAN", "数量"]],
        product_jans=["4901"],
    )
    inv = set(cache.product_jans)
    results = []
    for ch in children:
        set_n, set_err = parse_set_count(ch.set_count_raw)
        skip = ""
        pq = ""
        if ch.is_mix_candidate:
            skip = "mix_deferred_p6"
        elif set_err:
            skip = set_err
        elif not ch.stock_key:
            skip = "no_jan"
        elif ch.stock_key not in inv:
            skip = "stock_key_not_in_inventory"
        else:
            br = get_current_stock(ch.stock_key, cache)
            monthly = 0.0
            saf = safety_stock(monthly)
            avail = available_qty(float(br.stock_qty or 0), saf)
            pq = proper_qty(avail, set_n)  # type: ignore[arg-type]
        results.append((ch.child_sku, skip, pq))

    assert results[0] == ("C-8", "", 12)  # floor(100/8)=12
    assert results[1][0] == "C-M" and results[1][1] == "mix_deferred_p6"
    assert results[2][0] == "C-X" and results[2][1] == "stock_key_not_in_inventory"
    # 名簿なしは 0 埋めしていない
    assert results[2][2] == ""


if __name__ == "__main__":
    test_end_to_end_single_ok_and_skips()
    print("OK")
