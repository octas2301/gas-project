# -*- coding: utf-8 -*-
from formulas import available_qty, parse_set_count, proper_qty, safety_stock


def test_formula_basic():
    # 月販30 → 安全在庫7、実在庫100 → 可用93、N=8 → floor(93/8)=11
    saf = safety_stock(30)
    assert abs(saf - 7.0) < 1e-9
    avail = available_qty(100, saf)
    assert abs(avail - 93.0) < 1e-9
    assert proper_qty(avail, 8) == 11


def test_parse_set_numeric_and_text():
    n, err = parse_set_count(8)
    assert n == 8 and err == ""
    n, err = parse_set_count("8個で1セット")
    assert n == 8 and err == ""
    n, err = parse_set_count("")
    assert n is None and err == "set_count_invalid"


def test_available_floor_at_zero():
    assert available_qty(3, 10) == 0.0


if __name__ == "__main__":
    test_formula_basic()
    test_parse_set_numeric_and_text()
    test_available_floor_at_zero()
    print("OK")
