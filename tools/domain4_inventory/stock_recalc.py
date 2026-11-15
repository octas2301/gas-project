# -*- coding: utf-8 -*-
"""inventory-app InventoryLogic.gs 現在庫ロジックの Python 移植（読取・算出のみ）。"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from typing import Any

JST = timezone(timedelta(hours=9))
KIND_ADJUST = "確定後修正"

# Excel/Sheets シリアル日付の原点（1899-12-30、Sheets 互換）
_SHEETS_EPOCH = datetime(1899, 12, 30, tzinfo=JST)


@dataclass
class InventoryCache:
    session_data: list[list] = field(default_factory=list)
    count_data: list[list] = field(default_factory=list)
    receiving_data: list[list] = field(default_factory=list)
    shipping_data: list[list] = field(default_factory=list)
    product_jans: list[str] = field(default_factory=list)
    product_names: dict[str, str] = field(default_factory=dict)
    missing_sheets: list[str] = field(default_factory=list)


@dataclass
class StockBreakdown:
    stock_key: str
    stock_qty: float | None
    last_count_date: str
    count_on_last: float
    adjust_total: float
    receiving_after: float
    shipping_after: float
    error: str = ""


def norm_key(v: Any) -> str:
    s = str(v if v is not None else "").strip()
    if not s:
        return ""
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    # 数値セル由来の科学表記を避ける
    try:
        if isinstance(v, float) and v == int(v):
            s = str(int(v))
        elif isinstance(v, int):
            s = str(v)
    except Exception:
        pass
    return s.strip()


def _header_index(row: list, name: str) -> int:
    for i, c in enumerate(row):
        if str(c or "").strip() == name:
            return i
    return -1


def parse_dt(v: Any) -> datetime | None:
    """Sheets セル値 → aware datetime (JST)。"""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        if v.tzinfo is None:
            return v.replace(tzinfo=JST)
        return v.astimezone(JST)
    if isinstance(v, date) and not isinstance(v, datetime):
        return datetime(v.year, v.month, v.day, tzinfo=JST)
    if isinstance(v, (int, float)):
        # serial day (+ fraction)
        base = _SHEETS_EPOCH + timedelta(days=float(v))
        return base
    s = str(v).strip()
    if not s:
        return None
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
        "%Y/%m/%d",
    ):
        try:
            d = datetime.strptime(s[:19], fmt) if "T" not in fmt or " " in fmt else datetime.strptime(s[:19], fmt)
            return d.replace(tzinfo=JST)
        except ValueError:
            continue
    try:
        # 2026-11-15T12:34:56.000Z 等
        raw = s.replace("Z", "+00:00")
        d = datetime.fromisoformat(raw)
        if d.tzinfo is None:
            d = d.replace(tzinfo=JST)
        return d.astimezone(JST)
    except Exception:
        return None


def day_only(dt: datetime) -> date:
    return dt.astimezone(JST).date()


def day_start_ts(d: date) -> float:
    return datetime(d.year, d.month, d.day, tzinfo=JST).timestamp()


def day_end_ts(d: date) -> float:
    return day_start_ts(d + timedelta(days=1))


def get_last_count_date(cache: InventoryCache) -> date | None:
    """完了セッションの最大完了日を優先。無ければスキャン日時の最大日。"""
    if cache.session_data and len(cache.session_data) >= 2:
        sh = cache.session_data[0]
        i_end = _header_index(sh, "完了日時")
        if i_end >= 0:
            max_d: date | None = None
            for r in cache.session_data[1:]:
                if i_end >= len(r):
                    continue
                dt = parse_dt(r[i_end])
                if not dt:
                    continue
                d = day_only(dt)
                if max_d is None or d > max_d:
                    max_d = d
            if max_d:
                return max_d
    if not cache.count_data or len(cache.count_data) < 2:
        return None
    h = cache.count_data[0]
    i_date = _header_index(h, "日時")
    if i_date < 0:
        return None
    max_d = None
    for r in cache.count_data[1:]:
        if i_date >= len(r):
            continue
        dt = parse_dt(r[i_date])
        if not dt:
            continue
        d = day_only(dt)
        if max_d is None or d > max_d:
            max_d = d
    return max_d


def _completed_sessions_on_date(cache: InventoryCache, d: date) -> dict[str, bool] | None:
    if not cache.session_data or len(cache.session_data) < 2:
        return None
    sh = cache.session_data[0]
    si_id = _header_index(sh, "セッションID")
    si_end = _header_index(sh, "完了日時")
    if si_id < 0 or si_end < 0:
        return None
    out: dict[str, bool] = {}
    start = day_start_ts(d)
    end = day_end_ts(d)
    for r in cache.session_data[1:]:
        if si_end >= len(r) or si_id >= len(r):
            continue
        dt = parse_dt(r[si_end])
        if not dt:
            continue
        t = dt.timestamp()
        if start <= t < end:
            sid = norm_key(r[si_id])
            if sid:
                out[sid] = True
    return out


def get_count_total_on_date(jan: str, d: date, cache: InventoryCache) -> float:
    data = cache.count_data
    if not data or len(data) < 2:
        return 0.0
    h = data[0]
    i_date = _header_index(h, "日時")
    i_jan = _header_index(h, "単品JAN")
    i_qty = _header_index(h, "数量")
    if i_date < 0 or i_jan < 0 or i_qty < 0:
        return 0.0
    i_kind = _header_index(h, "種別")
    i_session = _header_index(h, "セッションID")
    completed = None
    if i_session >= 0:
        completed = _completed_sessions_on_date(cache, d)
    start = day_start_ts(d)
    end = day_end_ts(d)
    total = 0.0
    for r in data[1:]:
        if i_jan >= len(r) or norm_key(r[i_jan]) != jan:
            continue
        if i_kind >= 0 and i_kind < len(r) and norm_key(r[i_kind]) == KIND_ADJUST:
            continue
        if i_date >= len(r):
            continue
        dt = parse_dt(r[i_date])
        if not dt:
            continue
        t = dt.timestamp()
        if t < start or t >= end:
            continue
        if i_session >= 0 and completed is not None:
            sid = norm_key(r[i_session]) if i_session < len(r) else ""
            if sid and sid not in completed:
                continue
        qty = r[i_qty] if i_qty < len(r) else 0
        try:
            total += float(qty or 0)
        except (TypeError, ValueError):
            pass
    return total


def get_count_adjust_total(jan: str, from_d: date, cache: InventoryCache) -> float:
    data = cache.count_data
    if not data or len(data) < 2:
        return 0.0
    h = data[0]
    i_date = _header_index(h, "日時")
    i_jan = _header_index(h, "単品JAN")
    i_qty = _header_index(h, "数量")
    i_kind = _header_index(h, "種別")
    if min(i_date, i_jan, i_qty, i_kind) < 0:
        return 0.0
    from_time = day_start_ts(from_d)
    total = 0.0
    for r in data[1:]:
        if i_jan >= len(r) or norm_key(r[i_jan]) != jan:
            continue
        if i_kind >= len(r) or norm_key(r[i_kind]) != KIND_ADJUST:
            continue
        if i_date >= len(r):
            continue
        dt = parse_dt(r[i_date])
        if not dt:
            continue
        t = dt.timestamp()
        if t < from_time:
            continue
        try:
            total += float(r[i_qty] if i_qty < len(r) else 0)
        except (TypeError, ValueError):
            pass
    return total


def _sum_after(data: list[list], jan: str, after_dt: datetime) -> float:
    if not data or len(data) < 2:
        return 0.0
    h = data[0]
    i_date = _header_index(h, "日時")
    i_jan = _header_index(h, "単品JAN")
    i_qty = _header_index(h, "数量")
    if i_date < 0 or i_jan < 0 or i_qty < 0:
        return 0.0
    after_time = after_dt.timestamp()
    total = 0.0
    for r in data[1:]:
        if i_jan >= len(r) or norm_key(r[i_jan]) != jan:
            continue
        if i_date >= len(r):
            continue
        dt = parse_dt(r[i_date])
        if not dt:
            continue
        if dt.timestamp() > after_time:
            try:
                total += float(r[i_qty] if i_qty < len(r) else 0)
            except (TypeError, ValueError):
                pass
    return total


def get_current_stock(jan: str, cache: InventoryCache, last: date | None = None) -> StockBreakdown:
    key = norm_key(jan)
    if not key:
        return StockBreakdown(key, None, "", 0, 0, 0, 0, error="empty_key")
    last = last or get_last_count_date(cache)
    if last is None:
        return StockBreakdown(key, None, "", 0, 0, 0, 0, error="no_last_count_date")
    count_on = get_count_total_on_date(key, last, cache)
    adjust = get_count_adjust_total(key, last, cache)
    next_day = datetime(last.year, last.month, last.day, tzinfo=JST) + timedelta(days=1)
    # GAS: nextDayStart の getTime() より後 → after midnight of next day... 
    # nextDayStart = Date(y,m,d+1) at local midnight; receiving if t > afterTime
    # So after_dt = next day 00:00:00, include t > that.
    in_qty = _sum_after(cache.receiving_data, key, next_day)
    out_qty = _sum_after(cache.shipping_data, key, next_day)
    qty = count_on + adjust + in_qty - out_qty
    return StockBreakdown(
        stock_key=key,
        stock_qty=qty,
        last_count_date=last.isoformat(),
        count_on_last=count_on,
        adjust_total=adjust,
        receiving_after=in_qty,
        shipping_after=out_qty,
    )


def load_products_from_sheet(rows: list[list]) -> tuple[list[str], dict[str, str]]:
    if not rows:
        return [], {}
    h = rows[0]
    i_jan = _header_index(h, "単品JAN")
    i_name = _header_index(h, "商品名")
    jans: list[str] = []
    names: dict[str, str] = {}
    seen: set[str] = set()
    for r in rows[1:]:
        if i_jan < 0 or i_jan >= len(r):
            continue
        k = norm_key(r[i_jan])
        if not k or k in seen:
            continue
        seen.add(k)
        jans.append(k)
        if i_name >= 0 and i_name < len(r):
            names[k] = str(r[i_name] or "")
    return jans, names


REQUIRED_INV_SHEETS = (
    "商品マスタ",
    "棚卸セッション",
    "棚卸スキャンリスト",
    "入荷リスト",
    "出荷リスト",
)
