# -*- coding: utf-8 -*-
"""
タイムセール施策の不変条件（要件§2 と同期）。

sync / build のたびに壊しやすいルールをここに集約し、
テストと実行後検証の両方から呼ぶ。
"""
from __future__ import annotations

from typing import Any, Dict, List, Tuple

from schedule_class import parse_ymd
from sheet_schema import LANE_A, LANE_B


def sale_sort_key(r: Dict[str, Any]) -> Tuple:
    """開始日降順 → 商品名昇順 → sale_id（安定化）。"""
    d = parse_ymd(r.get("開始日"))
    date_rank = -(d.toordinal()) if d else float("inf")
    return (date_rank, str(r.get("商品名") or ""), str(r.get("sale_id") or ""))


def assert_sale_row_order(rows: List[Dict[str, Any]]) -> None:
    """並びが開始日降順→商品名昇順であること。"""
    got = list(rows)
    want = sorted(got, key=sale_sort_key)
    if got != want:
        def brief(r):
            return f"{r.get('開始日')}|{str(r.get('商品名') or '')[:20]}|{r.get('スケジュール')}"

        raise AssertionError(
            "並び不正（開始日降順→商品名昇順）\n"
            + "got:  "
            + " / ".join(brief(r) for r in got)
            + "\nwant: "
            + " / ".join(brief(r) for r in want)
        )


def assert_no_lane_a(rows: List[Dict[str, Any]]) -> None:
    bad = [r for r in rows if str(r.get("レーン") or "").strip() == LANE_A]
    if bad:
        raise AssertionError(f"レーンAが残存（未運用なのに {len(bad)} 行）")


def assert_b_images(rows: List[Dict[str, Any]]) -> None:
    for r in rows:
        if str(r.get("レーン") or "").strip() != LANE_B:
            continue
        img = str(r.get("画像") or "").strip()
        if not img.upper().startswith("=IMAGE("):
            raise AssertionError(
                f"B行にIMAGE式なし: {r.get('SKU')} {r.get('スケジュール')} img={img!r}"
            )


def assert_dates_parseable(rows: List[Dict[str, Any]]) -> None:
    for r in rows:
        for col in ("開始日", "終了日"):
            v = r.get(col)
            if v is None or str(v).strip() == "":
                continue
            if parse_ymd(v) is None:
                raise AssertionError(f"日付解析不能: {col}={v!r} sku={r.get('SKU')}")


def check_sale_sheet_invariants(rows: List[Dict[str, Any]]) -> List[str]:
    """違反メッセージ一覧（空ならOK）。"""
    errs: List[str] = []
    for fn in (assert_sale_row_order, assert_no_lane_a, assert_b_images, assert_dates_parseable):
        try:
            fn(rows)
        except AssertionError as e:
            errs.append(str(e))
    return errs
