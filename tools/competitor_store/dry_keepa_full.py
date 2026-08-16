# -*- coding: utf-8 -*-
"""Keepaフル upsert の dry（専用スプシ非書）。列名対応。"""
from keepa_full import (
    headers_live_like,
    plan_upsert_actions,
    product_to_full_row,
    row_values_for_headers,
)
from schema import PURPOSE_LISTING


def main() -> None:
    p1 = {
        "asin": "B0LIST0001",
        "title": "出品A",
        "csv": [[1]],
        "stats": {"current": [10] + [-1] * 17 + [11]},
    }
    p_bad = {"title": "noasin"}
    p2 = {
        "asin": "B0LIST0001",
        "title": "出品A",
        "csv": [[2]],
        "stats": {"current": [99] + [-1] * 17 + [11]},
    }
    acts = plan_upsert_actions([], [p1, p1, p_bad, p2], "dry", purpose=PURPOSE_LISTING)
    print("write=false purpose=出品 plan=" + ",".join(acts))
    assert acts == ["append", "skip_same_fp", "skip_no_asin", "append"]
    live = headers_live_like()
    rec = product_to_full_row(p1, "dry", purpose=PURPOSE_LISTING)
    row = row_values_for_headers(rec, live)
    print("live_asin_col", live.index("ASIN"), "val", row[live.index("ASIN")])
    print("live_purpose_col", live.index("目的"), "val", row[live.index("目的")])
    assert row[1] == "B0LIST0001"
    assert row[-1] == PURPOSE_LISTING


if __name__ == "__main__":
    main()
