# -*- coding: utf-8 -*-
"""要件§10.2 販売商品数（割当）ロジック。"""
from __future__ import annotations

import math
from datetime import date, timedelta
from typing import Any, Dict, Optional, Tuple

from schedule_class import is_official_b_schedule, parse_ymd

Q_MIN = 30
R_MAX = 0.85
M_DEFAULT = 2.0
M_EVENT = 3.0
L_LEAD_DAYS = 14


def deal_day_count(start: Any, end: Any) -> int:
    """開始〜終了（両端含む）。不正時は0。"""
    a = parse_ymd(start) if not isinstance(start, date) else start
    b = parse_ymd(end) if not isinstance(end, date) else end
    if not a or not b or b < a:
        return 0
    return (b - a).days + 1


def event_multiplier(schedule: str) -> float:
    """大型イベント名があれば M=3.0、それ以外 2.0。"""
    if is_official_b_schedule(schedule):
        return M_EVENT
    return M_DEFAULT


def compute_q_deal(
    *,
    v30: Optional[float],
    d_days: int,
    schedule: str = "",
    q_fba: Optional[float] = None,
    prev_allocation: Optional[float] = None,
    prev_sellthrough: Optional[float] = None,
    sellthrough_threshold: float = 0.95,
) -> Dict[str, Any]:
    """
    Q_deal_推奨 を返す。

    V30 が無い場合は Q_min のみ（要確認フラグ）。
    Q_fba が無い場合は上限クランプなし。
    """
    m = event_multiplier(schedule)
    v30_n = float(v30) if v30 is not None and str(v30).strip() != "" else None
    if v30_n is not None and v30_n < 0:
        v30_n = 0.0

    notes = []
    if d_days <= 0:
        notes.append("日数不明→Q_min")
        d_days = 0

    if v30_n is None:
        q_base = float(Q_MIN)
        notes.append("V30未取得→Q_min")
        need_v30 = True
    else:
        daily = v30_n / 30.0
        q_base = math.ceil(daily * max(d_days, 1) * m) if d_days > 0 else float(Q_MIN)
        need_v30 = False
        notes.append(f"V30={v30_n:g} D={d_days} M={m:g}")

    # 前回売切 ×1.5
    if (
        prev_allocation is not None
        and prev_sellthrough is not None
        and prev_allocation > 0
        and prev_sellthrough >= sellthrough_threshold
    ):
        bumped = math.ceil(float(prev_allocation) * 1.5)
        if bumped > q_base:
            q_base = float(bumped)
            notes.append(f"売切×1.5→{bumped}")

    q_low = max(float(Q_MIN), q_base)
    q_cap = None
    if q_fba is not None and str(q_fba).strip() != "":
        try:
            qf = float(q_fba)
            if qf > 0:
                q_cap = math.floor(qf * R_MAX)
                notes.append(f"Q_fba={qf:g} cap={q_cap}")
        except (TypeError, ValueError):
            pass

    if q_cap is not None:
        q_deal = max(float(Q_MIN), min(q_low, float(q_cap)))
        if q_cap < Q_MIN:
            notes.append("FBA上限がQ_min未満→延期候補")
            deferred = True
        else:
            deferred = False
    else:
        q_deal = q_low
        deferred = False
        notes.append("Q_fbaなし→上限クランプなし")

    return {
        "V30": v30_n,
        "D_days": d_days,
        "M": m,
        "Q_base": q_base,
        "Q_fba": float(q_fba) if q_fba not in (None, "") else None,
        "Q_deal": int(q_deal),
        "need_v30": need_v30,
        "deferred": deferred,
        "note": "; ".join(notes),
    }


def v30_windows(today: Optional[date] = None) -> Tuple[date, date, date, date]:
    """直近30日（両端含む）と前年同期間。"""
    today = today or date.today()
    end = today
    start = today - timedelta(days=29)

    def shift_year(d: date, years: int) -> date:
        try:
            return d.replace(year=d.year + years)
        except ValueError:
            # 2/29 → 2/28
            return d.replace(month=2, day=28, year=d.year + years)

    return start, end, shift_year(start, -1), shift_year(end, -1)
