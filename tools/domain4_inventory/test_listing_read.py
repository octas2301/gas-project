# -*- coding: utf-8 -*-
from listing_read import iter_child_rows


def test_parent_excluded_and_mix():
    rows = [
        ["説明行", "", "", "", "", ""],
        ["親SKU", "子SKU", "JANコード", "A.セット商品数", "商品名", "セット構成1_親SKU"],
        ["P1", "", "4901", "", "親", ""],
        ["P1", "C1-8", "4901", 8, "子8", ""],
        ["P2", "C2-M", "4902", 2, "ミックス", "PX"],
    ]
    children, meta = iter_child_rows(rows)
    assert meta["header_row"] == 2
    assert len(children) == 2
    assert children[0].child_sku == "C1-8"
    assert children[0].is_mix_candidate is False
    assert children[1].is_mix_candidate is True
    assert meta["set_col"] == "A.セット商品数"


if __name__ == "__main__":
    test_parent_excluded_and_mix()
    print("OK")
