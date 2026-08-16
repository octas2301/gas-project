# -*- coding: utf-8 -*-
"""Select JANs for scheduled (定時) refresh. Stock is input-only; never write listing master."""
from __future__ import annotations

from datetime import datetime, timedelta, timezone

from schema import PURPOSE_SCHEDULED, SHEET_HITS

INTERVAL_DAYS = 2


def child_stock_positive(rows: list[dict], stock_key: str = "在庫数", child_key: str = "子SKU", jan_key: str = "JANコード") -> set[str]:
    """API/マスタ読取結果を模した行。子SKUがあり在庫>0の JAN。"""
    jans = set()
    for r in rows:
        child = str(r.get(child_key) or "").strip()
        if not child:
            continue
        try:
            stock = float(str(r.get(stock_key) or "0").replace(",", ""))
        except ValueError:
            continue
        jan = str(r.get(jan_key) or "").strip()
        if stock > 0 and len(jan) >= 8:
            jans.add(jan)
    return jans


def last_scheduled(hits: list[dict], jan: str) -> datetime | None:
    latest = None
    for r in hits:
        if str(r.get("検索JAN") or "").strip() != jan:
            continue
        if str(r.get("目的") or "") != PURPOSE_SCHEDULED:
            continue
        raw = str(r.get("取得日時") or "").strip()
        if not raw:
            continue
        try:
            dt = datetime.strptime(raw[:19].replace("Z", ""), "%Y-%m-%dT%H:%M:%S")
            dt = dt.replace(tzinfo=timezone.utc)
        except ValueError:
            continue
        if latest is None or dt > latest:
            latest = dt
    return latest


def due_jans(stock_jans: set[str], hits: list[dict], now: datetime | None = None, interval_days: int = INTERVAL_DAYS) -> list[str]:
    now = now or datetime.now(timezone.utc)
    cut = now - timedelta(days=interval_days)
    out = []
    for jan in sorted(stock_jans):
        last = last_scheduled(hits, jan)
        if last is None or last <= cut:
            out.append(jan)
    return out
