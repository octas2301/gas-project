# -*- coding: utf-8 -*-
"""現在庫再計算の最小ユニットテスト（API不要）。"""
from __future__ import annotations

from datetime import date, datetime, timedelta, timezone

from stock_recalc import InventoryCache, get_current_stock, get_last_count_date

JST = timezone(timedelta(hours=9))


def _serial(dt: datetime) -> float:
    epoch = datetime(1899, 12, 30, tzinfo=JST)
    return (dt - epoch).total_seconds() / 86400.0


def test_basic_count_plus_receive_minus_ship():
    last = date(2026, 11, 10)
    # 棚卸日に 100、翌々日入荷 10、出荷 3 → 107
    count_day = datetime(2026, 11, 10, 12, 0, tzinfo=JST)
    recv = datetime(2026, 11, 12, 9, 0, tzinfo=JST)
    ship = datetime(2026, 11, 12, 15, 0, tzinfo=JST)
    cache = InventoryCache(
        session_data=[
            ["セッションID", "完了日時"],
            ["S1", _serial(datetime(2026, 11, 10, 18, 0, tzinfo=JST))],
        ],
        count_data=[
            ["日時", "単品JAN", "数量", "種別", "セッションID"],
            [_serial(count_day), "JAN1", 100, "", "S1"],
        ],
        receiving_data=[
            ["日時", "単品JAN", "数量"],
            [_serial(recv), "JAN1", 10],
        ],
        shipping_data=[
            ["日時", "単品JAN", "数量"],
            [_serial(ship), "JAN1", 3],
        ],
    )
    assert get_last_count_date(cache) == last
    b = get_current_stock("JAN1", cache)
    assert b.error == ""
    assert b.stock_qty == 107


def test_same_day_receive_excluded():
    # 棚卸当日の入荷は含めない
    count_day = datetime(2026, 11, 10, 12, 0, tzinfo=JST)
    same_day_recv = datetime(2026, 11, 10, 20, 0, tzinfo=JST)
    cache = InventoryCache(
        session_data=[
            ["セッションID", "完了日時"],
            ["S1", _serial(datetime(2026, 11, 10, 18, 0, tzinfo=JST))],
        ],
        count_data=[
            ["日時", "単品JAN", "数量", "種別", "セッションID"],
            [_serial(count_day), "JAN2", 50, "", "S1"],
        ],
        receiving_data=[
            ["日時", "単品JAN", "数量"],
            [_serial(same_day_recv), "JAN2", 9],
        ],
        shipping_data=[["日時", "単品JAN", "数量"]],
    )
    b = get_current_stock("JAN2", cache)
    assert b.stock_qty == 50


def test_adjust_included():
    count_day = datetime(2026, 11, 10, 12, 0, tzinfo=JST)
    adj = datetime(2026, 11, 11, 10, 0, tzinfo=JST)
    cache = InventoryCache(
        session_data=[
            ["セッションID", "完了日時"],
            ["S1", _serial(datetime(2026, 11, 10, 18, 0, tzinfo=JST))],
        ],
        count_data=[
            ["日時", "単品JAN", "数量", "種別", "セッションID"],
            [_serial(count_day), "JAN3", 40, "", "S1"],
            [_serial(adj), "JAN3", 5, "確定後修正", ""],
        ],
        receiving_data=[["日時", "単品JAN", "数量"]],
        shipping_data=[["日時", "単品JAN", "数量"]],
    )
    b = get_current_stock("JAN3", cache)
    assert b.stock_qty == 45


def test_no_last_count_fail_closed():
    cache = InventoryCache(
        session_data=[["セッションID", "完了日時"]],
        count_data=[["日時", "単品JAN", "数量", "種別", "セッションID"]],
        receiving_data=[["日時", "単品JAN", "数量"]],
        shipping_data=[["日時", "単品JAN", "数量"]],
    )
    b = get_current_stock("JANX", cache)
    assert b.error == "no_last_count_date"
    assert b.stock_qty is None


if __name__ == "__main__":
    test_basic_count_plus_receive_minus_ship()
    test_same_day_receive_excluded()
    test_adjust_included()
    test_no_last_count_fail_closed()
    print("OK")
