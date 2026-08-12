# -*- coding: utf-8 -*-
"""mail_points_remind の選定ロジック単体。"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from mail_points_remind import (  # noqa: E402
    needs_apply_remind,
    needs_restore_remind,
    select_b_rows,
)


def main() -> int:
    fails = []
    today = date(2026, 8, 11)
    sales = [
        {
            "レーン": "B_公式",
            "状態": "UL済",
            "SKU": "s1",
            "開始日": "2026-08-12",
            "終了日": "2026-08-13",
        },
        {
            "レーン": "B_公式",
            "状態": "見送り",
            "SKU": "skip",
            "開始日": "2026-08-12",
            "終了日": "2026-08-13",
        },
    ]
    got = select_b_rows(sales, today=today, kind="apply", days=1, tol=0)
    if len(got) != 1 or got[0][0]["SKU"] != "s1":
        fails.append("apply T-1")
    else:
        print("OK  : apply T-1")

    got_r = select_b_rows(
        sales, today=date(2026, 8, 14), kind="restore", days=1, tol=0
    )
    if len(got_r) != 1:
        fails.append("restore end+1")
    else:
        print("OK  : restore end+1")

    m_applied = {"ポイント状態": "期間中適用済", "期間中ポイント%": "1", "出品者ポイント現在%": "1"}
    if needs_apply_remind(m_applied):
        fails.append("applied should skip apply remind")
    else:
        print("OK  : skip apply when 期間中適用済")

    # 既に期間中%でも未適用なら催促（日程リマインド）
    m_backed = {
        "ポイント状態": "セール前退避済",
        "期間中ポイント%": "1",
        "出品者ポイント現在%": "1",
        "セール前ポイント%": "1",
    }
    if not needs_apply_remind(m_backed):
        fails.append("backed-up should remind apply even if cur==period")
    else:
        print("OK  : apply remind when セール前退避済 (cur==period)")

    m_need_restore = {
        "ポイント状態": "期間中適用済",
        "期間中ポイント%": "1",
        "出品者ポイント現在%": "1",
        "セール前ポイント%": "20",
    }
    if not needs_restore_remind(m_need_restore):
        fails.append("need restore")
    else:
        print("OK  : restore remind when 期間中適用済+before")

    m_applied_no_before = {
        "ポイント状態": "期間中適用済",
        "期間中ポイント%": "1",
        "出品者ポイント現在%": "1",
    }
    if not needs_restore_remind(m_applied_no_before):
        fails.append("applied without before should still remind restore")
    else:
        print("OK  : restore remind when 期間中適用済 and before empty")

    m_done = dict(m_need_restore)
    m_done["ポイント状態"] = "セール前復元済"
    if needs_restore_remind(m_done):
        fails.append("restored should skip")
    else:
        print("OK  : skip restore when セール前復元済")

    if fails:
        for f in fails:
            print("FAIL:", f)
        return 1
    print("ALL OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
