# -*- coding: utf-8 -*-
from datetime import date

from taper_send import amazon_jp_url, build_points_tsv_explicit, build_taper_mail


def test_url():
    assert amazon_jp_url("B0DB664V55").endswith("B0DB664V55")
    assert amazon_jp_url("") == ""


def test_tsv():
    tsv = build_points_tsv_explicit(
        [{"sku": "s1", "to_pct": 19}, {"sku": "s2", "to_pct": 1}]
    )
    assert "sku\tpoints_percent" in tsv
    assert "s1\t19" in tsv
    assert "s2\t1" in tsv


def test_mail_body():
    subj, body = build_taper_mail(
        today=date(2026, 8, 12),
        prod=False,
        ok=[
            {
                "sku": "KOUSO-x",
                "asin": "B0DB664V55",
                "name": "酵素",
                "from_pct": 22,
                "to_pct": 19,
            }
        ],
        skipped=[{"sku": "z", "reason": "本日実行済"}],
        failed=[],
        sheet_only=[{"sku": "b", "from_pct": 22, "to_pct": 22, "sheet_only_reason": "B期間中カレンダー"}],
    )
    assert "dry_run" in subj or "dry_run" in body
    assert "22%" in body and "19%" in body
    assert "amazon.co.jp/dp/B0DB664V55" in body
    assert "KOUSO-x" in body
    assert "手動リカバリ" in body
    assert "シートのみ" in body
    assert "B期間中カレンダー" in body


if __name__ == "__main__":
    test_url()
    test_tsv()
    test_mail_body()
    print("ALL OK")
