# -*- coding: utf-8 -*-
"""§10.2 数量＋§10.6 B先行A埋めのユニット検算。"""
from __future__ import annotations

import math
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from qty_logic import Q_MIN, R_MAX, compute_q_deal, deal_day_count, event_multiplier
from schedule_class import (
    a_end_before_b,
    a_start_after_b,
    cooldown_ok,
    fill_a_windows,
    pick_schedules_split,
)
from lane_a_patch import build_discounted_price_patch
from revise_b_qty import select_b_rows


def expect(cond: bool, msg: str, fails: list) -> None:
    if not cond:
        fails.append(msg)
        print("FAIL:", msg)
    else:
        print("OK  :", msg)


def main() -> int:
    fails: list = []

    expect(deal_day_count("2026-08-28", "2026-09-03") == 7, "Smile 7日", fails)
    expect(deal_day_count("2026-11-27", "2026-12-07") == 11, "BF 11日", fails)

    expect(event_multiplier("Amazon Smileセール (2026-08-28 - 2026-09-03)") == 3.0, "Smile M=3", fails)
    expect(event_multiplier("月 (2026-08-17 - 2026-08-23)") == 2.0, "月 M=2", fails)

    r = compute_q_deal(v30=108, d_days=7, schedule="Amazon Smileセール", q_fba=130)
    expect(r["Q_deal"] == 76, f"108/7日 →76 got={r['Q_deal']}", fails)

    r = compute_q_deal(v30=25, d_days=7, schedule="Smile", q_fba=92)
    expect(r["Q_deal"] == 30, f"25/7日 →30 got={r['Q_deal']}", fails)

    r = compute_q_deal(v30=108, d_days=11, schedule="ブラックフライデー", q_fba=130)
    expect(r["Q_deal"] == 110, f"108/11日 FBA上限→110 got={r['Q_deal']}", fails)

    r = compute_q_deal(v30=None, d_days=7, schedule="Smile", q_fba=100)
    expect(r["Q_deal"] == Q_MIN and r["need_v30"], "V30なし→30要確認", fails)

    r = compute_q_deal(
        v30=10, d_days=7, schedule="Smile", q_fba=500, prev_allocation=100, prev_sellthrough=0.96
    )
    expect(r["Q_deal"] == 150, f"売切×1.5→150 got={r['Q_deal']}", fails)

    r = compute_q_deal(v30=100, d_days=14, schedule="Smile", q_fba=20)
    expect(r["deferred"] and r["Q_deal"] == Q_MIN, "FBA不足→延期候補", fails)

    expect(math.floor(130 * R_MAX) == 110, "R_max cap 130", fails)

    expect(
        not cooldown_ok(date(2026, 8, 23), date(2026, 8, 28), days=14),
        "月終了→Smile開始5日はおすすめ同士なら不可",
        fails,
    )
    expect(
        cooldown_ok(date(2026, 8, 23), date(2026, 9, 6), days=14),
        "終了+14日は可",
        fails,
    )
    expect(
        cooldown_ok(date(2026, 9, 3), date(2026, 11, 27), days=14),
        "Smile→BFは可",
        fails,
    )

    expect(a_end_before_b(date(2026, 8, 28), 0) == date(2026, 8, 27), "gap0 A終了=B前日", fails)
    expect(a_end_before_b(date(2026, 8, 28), 5) == date(2026, 8, 22), "gap5 A終了", fails)
    expect(a_start_after_b(date(2026, 9, 3), 0) == date(2026, 9, 4), "gap0 A開始=B翌日", fails)

    cands = [
        {
            "schedule": "Amazon Smileセール (2026-08-28 - 2026-09-03)",
            "start": "2026-08-28",
            "end": "2026-09-03",
        },
        {
            "schedule": "ブラックフライデー (2026-11-27 - 2026-12-07)",
            "start": "2026-11-27",
            "end": "2026-12-07",
        },
        {"schedule": "月 (2026-08-17 - 2026-08-23)", "start": "2026-08-17", "end": "2026-08-23"},
    ]
    today = date(2026, 8, 11)

    # 2026-08-11: BFは早期割引締切(9/30)前のため horizon90でもSmile+BF
    b, a = pick_schedules_split(
        cands, limit_b=2, limit_a=2, today=today, ab_gap_days=0, b_horizon_days=90
    )
    expect(len(b) == 2 and any("Smile" in x["schedule"] for x in b), "早期締切前はSmile+BF", fails)
    expect(len(a) == 2, f"A2本 got={len(a)}", fails)
    expect(a[0]["end"] == "2026-08-27", f"前A終了8/27 got={a[0].get('end')}", fails)
    expect(a[0]["start"] == "2026-08-14", f"前A開始8/14(14日) got={a[0].get('start')}", fails)
    expect(a[1]["start"] == "2026-09-04", f"後A開始9/04 got={a[1].get('start')}", fails)
    expect(a[1]["end"] == "2026-09-17", f"後A終了9/17 got={a[1].get('end')}", fails)

    b5, a5 = pick_schedules_split(
        cands, limit_b=1, limit_a=1, today=today, ab_gap_days=5, b_horizon_days=90
    )
    expect(len(b5) == 1 and "Smile" in b5[0]["schedule"], "limit_b=1はSmile優先", fails)
    expect(a5[0]["end"] == "2026-08-22", f"gap5前A終了8/22 got={a5[0].get('end')}", fails)

    # 早期締切後＋短いhorizonならBF除外
    b_late, _ = pick_schedules_split(
        cands, limit_b=2, limit_a=1, today=date(2026, 10, 1), ab_gap_days=0, b_horizon_days=30
    )
    expect(
        not any("ブラック" in x["schedule"] for x in b_late),
        "9/30過ぎ＋horizon外ならBF除外",
        fails,
    )

    # 月枠はAに使わない
    expect(not any("月 (" in str(x.get("schedule")) for x in a), "Aは合成のみ（月枠不使用）", fails)

    smile_only = [
        {
            "schedule": "Amazon Smileセール (2026-08-28 - 2026-09-03)",
            "start": "2026-08-28",
            "end": "2026-09-03",
        }
    ]
    filled = fill_a_windows(smile_only, today=today, limit_a=2, gap_days=0)
    expect(len(filled) == 2, "fill_a_windows 2本", fails)

    # SKUローカル: カタログBFは無く、カスタム8/12-13がバルク提出
    from schedule_class import pick_schedules_sku_local

    local_only = [
        {
            "schedule": "カスタム - (2026-08-12 - 2026-08-13)",
            "start": "2026-08-12",
            "end": "2026-08-13",
            "deal_type": "おすすめタイムセール",
        },
        {
            "schedule": "月 (2026-08-17 - 2026-08-23)",
            "start": "2026-08-17",
            "end": "2026-08-23",
            "deal_type": "数量限定タイムセール",
        },
    ]
    b_loc, cust, a_loc = pick_schedules_sku_local(
        local_only, limit_b=2, limit_a=2, limit_bulk_custom=2, today=today
    )
    expect(len(b_loc) == 0, "ローカルに名付き無し", fails)
    expect(len(cust) == 1 and "08-12" in cust[0]["start"], "カスタム提出", fails)
    expect(len(a_loc) >= 1, "合成Aあり", fails)

    # ドロップダウン相当: 日付付きカスタム2本は両方（非重なり）
    two_customs = [
        {
            "schedule": "カスタム - (2026-08-12 - 2026-08-13)",
            "start": "2026-08-12",
            "end": "2026-08-13",
            "deal_type": "おすすめタイムセール",
            "source": "dropdown",
        },
        {
            "schedule": "カスタム - (2026-09-18 - 2026-09-24)",
            "start": "2026-09-18",
            "end": "2026-09-24",
            "deal_type": "おすすめタイムセール",
            "source": "dropdown",
        },
    ]
    _b2, cust2, _a2 = pick_schedules_sku_local(
        two_customs, limit_b=2, limit_a=0, limit_bulk_custom=20, today=today
    )
    expect(len(cust2) == 2, "ドロップダウン日付付きカスタム2本", fails)

    # lane A patch
    patch = build_discounted_price_patch(
        marketplace_id="A1VC38T7YXB528",
        currency="JPY",
        sale_price=3479,
        start=date(2026, 8, 14),
        end=date(2026, 8, 27),
        our_price=4480,
    )
    expect(patch["patches"][0]["path"] == "/attributes/purchasable_offer", "A patch path", fails)
    sch = patch["patches"][0]["value"][0]["discounted_price"][0]["schedule"][0]
    expect(sch["value_with_tax"] == 3479, "A patch price", fails)
    expect("2026-08-14" in sch["start_at"], "A patch start", fails)

    # revise selector
    sample = [
        {
            "レーン": "B_公式",
            "提出対象": "はい",
            "状態": "予定",
            "開始日": "2026-08-28",
            "ASIN": "B0DB664V55",
        },
        {
            "レーン": "B_公式",
            "提出対象": "はい",
            "状態": "予定",
            "開始日": "2026-11-27",
            "ASIN": "BF",
        },
    ]
    near = select_b_rows(sample, today=today, within_days=21, all_upcoming=False)
    expect(len(near) == 1 and near[0]["ASIN"] == "B0DB664V55", "21日内はSmileのみ", fails)
    allu = select_b_rows(sample, today=today, within_days=21, all_upcoming=True)
    expect(len(allu) == 2, "all-upcomingは2", fails)

    print("---")
    if fails:
        print(f"FAILED {len(fails)}")
        return 1
    print("ALL PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
