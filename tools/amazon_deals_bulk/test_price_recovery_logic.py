# -*- coding: utf-8 -*-
from datetime import date

from price_recovery_logic import (
    build_our_price_patch,
    calendar_active_pct,
    calendar_step_count,
    effective_price,
    next_points_percent,
    plan_calendar_sync,
    plan_one_points_step,
    points_blocks_recovery,
    promo_points_yen,
    propose_points_taper,
    select_due_recovery_rows,
    skus_in_active_b,
)


def test_cinderella_display():
    assert promo_points_yen(4480, 22) == 986
    assert effective_price(4480, 22) == 3494


def test_cinderella_taper_propose():
    p = propose_points_taper(4480, 22, end_pct=1)
    assert p["販促ポイント円"] == 986
    assert p["実質価格円"] == 3494
    assert p["減衰間隔"] == "2週間"
    assert p["減衰期間"] == "3か月"
    assert p["減衰段%"] >= 1
    assert p["開始%"] == 22
    assert p["終着%"] == 1


def test_next_points():
    assert next_points_percent(22, 3, 1) == 19
    assert next_points_percent(4, 3, 1) == 1


def test_plan_one_points_step():
    row = {
        "SKU": "s1",
        "目標売価円": 4480,
        "販促ポイント%": 22,
        "セール前ポイント%": 1,
        "減衰段%": 3,
        "減衰間隔": "2週間",
        "減衰状態": "進行中",
        "減衰中ポイント%": 19,
        "出品者ポイント現在%": 22,
    }
    p = plan_one_points_step(row, today=date(2026, 9, 4))
    assert p["from_pct"] == 19
    assert p["to_pct"] == 16
    assert p["status"] == "進行中"
    assert p["next_date"] == date(2026, 9, 18)


def test_g10_active_b_and_points():
    sales = [
        {
            "レーン": "B_公式",
            "SKU": "s1",
            "状態": "UL済",
            "開始日": "2026-08-28",
            "終了日": "2026-09-03",
        }
    ]
    assert "s1" in skus_in_active_b(sales, today=date(2026, 8, 30))
    assert "s1" not in skus_in_active_b(sales, today=date(2026, 9, 4))
    assert points_blocks_recovery({"ポイント状態": "期間中適用済"})


def test_select_due():
    rows = [
        {
            "SKU": "a",
            "有効": True,
            "目標売価円": 4480,
            "販促ポイント%": 22,
            "減衰段%": 3,
            "減衰間隔": "2週間",
            "減衰状態": "進行中",
            "次回減衰日": "2026-09-04",
        },
        {
            "SKU": "b",
            "有効": True,
            "目標売価円": 4480,
            "販促ポイント%": 22,
            "減衰段%": 3,
            "減衰間隔": "2週間",
            "減衰状態": "未開始",
        },
    ]
    due = select_due_recovery_rows(rows, today=date(2026, 9, 4), include_start=False)
    assert [r["SKU"] for r in due] == ["a"]
    due2 = select_due_recovery_rows(rows, today=date(2026, 9, 4), include_start=True)
    assert {r["SKU"] for r in due2} == {"a", "b"}
    rows[1]["減衰実行依頼"] = "TRUE"
    due3 = select_due_recovery_rows(rows, today=date(2026, 9, 1), include_start=False)
    assert {r["SKU"] for r in due3} == {"b"}


def test_calendar_cinderella():
    row = {
        "SKU": "b",
        "販促ポイント%": 22,
        "セール前ポイント%": 1,
        "減衰段%": 4,
        "減衰間隔": "2週間",
        "減衰開始日": "2026-08-14",
        "減衰中ポイント%": 22,
        "目標売価円": 4480,
    }
    assert calendar_step_count(row, date(2026, 8, 12)) == 0
    assert calendar_active_pct(row, date(2026, 8, 12)) == 22
    assert calendar_step_count(row, date(2026, 8, 14)) == 1
    assert calendar_active_pct(row, date(2026, 8, 14)) == 18
    assert calendar_step_count(row, date(2026, 8, 28)) == 2
    assert calendar_active_pct(row, date(2026, 8, 28)) == 14
    p = plan_calendar_sync(row, today=date(2026, 8, 14))
    assert p["from_pct"] == 22
    assert p["to_pct"] == 18
    assert p["next_date"] == date(2026, 8, 28)
    assert p["skip_api"] is False
    due = select_due_recovery_rows(
        [dict(row, 有効=True, 減衰状態="未開始")],
        today=date(2026, 8, 12),
        include_start=True,
    )
    assert due == []
    due2 = select_due_recovery_rows(
        [dict(row, 有効=True, 減衰状態="進行中", 次回減衰日="2026-08-14")],
        today=date(2026, 8, 14),
        include_start=True,
    )
    assert [r["SKU"] for r in due2] == ["b"]
    row_s = dict(row, SKU="s", 有効=True, 減衰状態="未開始", 次回減衰日="2026-08-14")
    due3 = select_due_recovery_rows(
        [dict(row, 有効=True, 減衰状態="未開始", 次回減衰日="2026-08-14"), row_s],
        today=date(2026, 8, 14),
        include_start=True,
    )
    assert {r["SKU"] for r in due3} == {"b", "s"}


def test_patch_shape():
    body = build_our_price_patch(
        marketplace_id="A1VC38T7YXB528", currency="JPY", our_price=4480
    )
    assert body["patches"][0]["path"] == "/attributes/purchasable_offer"
    assert body["patches"][0]["value"][0]["our_price"][0]["schedule"][0]["value_with_tax"] == 4480.0


if __name__ == "__main__":
    test_cinderella_display()
    test_cinderella_taper_propose()
    test_next_points()
    test_plan_one_points_step()
    test_g10_active_b_and_points()
    test_select_due()
    test_calendar_cinderella()
    test_patch_shape()
    print("ALL OK")
