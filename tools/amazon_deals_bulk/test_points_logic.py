# -*- coding: utf-8 -*-
from points_logic import (
    MODE_APPLY,
    MODE_RESTORE,
    before_percent,
    build_points_tsv,
    needs_sync,
    period_percent,
    restore_percent,
    select_diff_rows,
    send_percent,
)


def test_period_default_one():
    assert period_percent({}) == 1
    assert period_percent({"期間中ポイント%": ""}) == 1
    assert period_percent({"ポイント目標%": "3"}) == 3  # legacy


def test_explicit_zero():
    assert period_percent({"期間中ポイント%": "0"}) == 0


def test_before_and_restore():
    r = {
        "SKU": "a",
        "有効": "TRUE",
        "期間中ポイント%": "1",
        "セール前ポイント%": "1",
        "減衰中ポイント%": "18",
        "出品者ポイント現在%": "1",
    }
    assert before_percent(r) == 1
    assert restore_percent(r) == 18
    assert needs_sync(r, MODE_APPLY) is False
    assert needs_sync(r, MODE_RESTORE) is True
    tsv = build_points_tsv([r], MODE_RESTORE)
    assert "a\t18\n" in tsv


def test_select_restore_skips_empty_before():
    rows = [
        {"SKU": "s1", "有効": "TRUE", "減衰中ポイント%": "18", "出品者ポイント現在%": "1"},
        {"SKU": "s2", "有効": "TRUE", "出品者ポイント現在%": "1"},
    ]
    assert [r["SKU"] for r in select_diff_rows(rows, mode=MODE_RESTORE)] == ["s1"]


def test_restore_calendar_catchup():
    from datetime import date

    r = {
        "SKU": "c",
        "販促ポイント%": "22",
        "セール前ポイント%": "1",
        "減衰段%": "4",
        "減衰間隔": "2週間",
        "減衰開始日": "2026-08-14",
        "減衰中ポイント%": "22",
        "出品者ポイント現在%": "1",
    }
    assert restore_percent(r, today=date(2026, 8, 12)) == 22
    assert restore_percent(r, today=date(2026, 8, 14)) == 18
    assert restore_percent(r, today=date(2026, 8, 28)) == 14
    r["出品者ポイント現在%"] = "18"
    got = select_diff_rows(
        [r], mode=MODE_RESTORE, today=date(2026, 9, 4)
    )
    assert got and send_percent(got[0], MODE_RESTORE, today=date(2026, 9, 4)) == 14


def test_point_status_vocab():
    from points_logic import (
        POINT_STATUS_APPLIED,
        POINT_STATUS_BACKED_UP,
        POINT_STATUS_RESTORED,
        POINT_STATUS_UNSET,
        normalize_point_status,
        status_after_backup,
        status_after_send,
    )

    assert normalize_point_status("") == POINT_STATUS_UNSET
    assert normalize_point_status("セール前へ復元済") == POINT_STATUS_RESTORED
    assert normalize_point_status("期間中適用済") == POINT_STATUS_APPLIED
    assert status_after_backup() == POINT_STATUS_BACKED_UP
    assert status_after_send(MODE_APPLY, "DONE") == POINT_STATUS_APPLIED
    assert status_after_send(MODE_RESTORE, None) == POINT_STATUS_RESTORED
    assert status_after_send(MODE_APPLY, "IN_PROGRESS") == "フィードIN_PROGRESS"


def test_sale_skus_for_points():
    from datetime import date

    from points_logic import sale_skus_for_points

    today = date(2026, 8, 11)
    sales = [
        {
            "レーン": "B_公式",
            "状態": "UL済",
            "SKU": "near",
            "開始日": "2026-08-12",
            "終了日": "2026-08-13",
        },
        {
            "レーン": "B_公式",
            "状態": "UL済",
            "SKU": "far",
            "開始日": "2026-09-18",
            "終了日": "2026-09-24",
        },
        {
            "レーン": "B_公式",
            "状態": "見送り",
            "SKU": "skip",
            "開始日": "2026-08-12",
            "終了日": "2026-08-13",
        },
    ]
    apply = sale_skus_for_points(sales, mode=MODE_APPLY, today=today, within_days=1)
    assert apply == {"near"}
    restore = sale_skus_for_points(
        sales, mode=MODE_RESTORE, today=date(2026, 8, 14), within_days=1
    )
    assert restore == {"near"}
    # 実施中
    mid = sale_skus_for_points(
        [
            {
                "レーン": "B_公式",
                "状態": "実施中",
                "SKU": "run",
                "開始日": "2026-08-10",
                "終了日": "2026-08-13",
            }
        ],
        mode=MODE_APPLY,
        today=today,
        within_days=1,
    )
    assert mid == {"run"}


def test_select_diff_sku_allow():
    rows = [
        {"SKU": "a", "有効": "TRUE", "期間中ポイント%": "1", "出品者ポイント現在%": "2"},
        {"SKU": "b", "有効": "TRUE", "期間中ポイント%": "1", "出品者ポイント現在%": "2"},
    ]
    got = select_diff_rows(rows, mode=MODE_APPLY, sku_allow={"a"})
    assert [r["SKU"] for r in got] == ["a"]


def test_skus_missing_before():
    from points_logic import skus_missing_before

    rows = [
        {"SKU": "a", "セール前ポイント%": "20"},
        {"SKU": "b"},
        {"SKU": "c", "セール前ポイント%": ""},
    ]
    assert skus_missing_before(rows) == ["b", "c"]
