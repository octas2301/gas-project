# -*- coding: utf-8 -*-
"""適正在庫の計算式（DOMAIN4_P0_REQUIREMENTS §3・§4）。"""
from __future__ import annotations

import math
from typing import Any


SAFETY_DAYS = 7.0
DAYS_PER_MONTH = 30.0


def to_float(v: Any) -> float | None:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def safety_stock(monthly_sales: float, safety_days: float = SAFETY_DAYS) -> float:
    return float(monthly_sales) * (safety_days / DAYS_PER_MONTH)


def available_qty(stock_qty: float, safety: float) -> float:
    return max(0.0, float(stock_qty) - float(safety))


def proper_qty(available: float, set_count: float) -> int | None:
    """floor(可用 / N)。N<=0 は None。"""
    n = float(set_count)
    if n <= 0:
        return None
    return int(math.floor(float(available) / n))


def parse_set_count(v: Any) -> tuple[float | None, str]:
    """
    数値セル前提。文字なら「数字」を抽出するフォールバック。
    戻り: (値, エラーコード空文字 or set_count_invalid)
    """
    if v is None or v == "":
        return None, "set_count_invalid"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if float(v) <= 0:
            return None, "set_count_invalid"
        return float(v), ""
    s = str(v).strip().replace(",", "")
    try:
        n = float(s)
        if n <= 0:
            return None, "set_count_invalid"
        return n, ""
    except ValueError:
        pass
    import re

    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if not m:
        return None, "set_count_invalid"
    n = float(m.group(1))
    if n <= 0:
        return None, "set_count_invalid"
    return n, ""
