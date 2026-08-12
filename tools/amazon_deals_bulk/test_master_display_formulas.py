# -*- coding: utf-8 -*-
from sheets_io import master_next_pct_formula


def test_next_pct_formula_interpolated():
    s = master_next_pct_formula(
        2, c_step="Z", c_active="AE", c_pct="X", c_before="R"
    )
    assert "{c_" not in s
    assert "{r}" not in s
    assert 'AE2=""' in s
    assert "X2" in s and "Z2" in s and "R2" in s
    assert s.startswith("=")


if __name__ == "__main__":
    test_next_pct_formula_interpolated()
    print("ALL OK")
