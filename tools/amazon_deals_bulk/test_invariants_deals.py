# -*- coding: utf-8 -*-
"""不変条件の単体テスト（Sheets不要）。"""
from __future__ import annotations

from schedule_class import format_ymd, parse_ymd
from invariants import assert_sale_row_order, check_sale_sheet_invariants, sale_sort_key


def expect(cond: bool, msg: str, fails: list) -> None:
    if not cond:
        fails.append(msg)
        print("FAIL:", msg)
    else:
        print("OK  :", msg)


def main() -> int:
    fails: list = []
    # シリアル日付
    d = parse_ymd(46262)
    expect(d is not None and d.isoformat() == "2026-08-28", "serial→2026-08-28", fails)
    d2 = parse_ymd(46246)
    expect(d2 is not None and d2.isoformat() == "2026-08-12", "serial→2026-08-12", fails)
    expect(format_ymd(46262) == "2026-08-28", "format_ymd serial", fails)
    expect(parse_ymd("2026-08-28").isoformat() == "2026-08-28", "ymd string", fails)

    rows = [
        {
            "sale_id": "1",
            "レーン": "B_公式",
            "開始日": 46246,
            "商品名": "B商品",
            "スケジュール": "カスタム",
            "画像": '=IMAGE("https://example.com/b.jpg")',
        },
        {
            "sale_id": "2",
            "レーン": "B_公式",
            "開始日": 46262,
            "商品名": "A商品",
            "スケジュール": "Smile",
            "画像": '=IMAGE("https://example.com/a.jpg")',
        },
        {
            "sale_id": "3",
            "レーン": "B_公式",
            "開始日": 46262,
            "商品名": "B商品",
            "スケジュール": "Smile",
            "画像": '=IMAGE("https://example.com/b2.jpg")',
        },
    ]
    ordered = sorted(rows, key=sale_sort_key)
    # 降順: 8/28 A商品, 8/28 B商品, 8/12 B商品
    expect(
        [r["sale_id"] for r in ordered] == ["2", "3", "1"],
        "sort 開始日降順→商品名昇順",
        fails,
    )
    try:
        assert_sale_row_order(ordered)
        expect(True, "assert_sale_row_order OK", fails)
    except AssertionError as e:
        expect(False, f"assert_sale_row_order: {e}", fails)

    bad_order = list(reversed(ordered))
    errs = check_sale_sheet_invariants(bad_order)
    expect(any("並び不正" in e for e in errs), "並び違反を検出", fails)

    print("---")
    if fails:
        print("FAILED", len(fails))
        return 1
    print("ALL PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
