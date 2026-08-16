# -*- coding: utf-8 -*-
"""Merge inventory API payloads. Never write listing master. No live HTTP."""
from __future__ import annotations


def qty_by_seller_sku(payload: dict) -> dict[str, float]:
    """Amazon GetInventorySummaries-like JSON (fixture)."""
    out = {}
    for row in payload.get("inventorySummaries") or []:
        sku = str(row.get("sellerSku") or "").strip()
        if not sku:
            continue
        q = row.get("totalQuantity")
        try:
            out[sku] = float(q)
        except (TypeError, ValueError):
            continue
    return out


def qty_from_listings(payload: dict) -> float | None:
    arr = payload.get("fulfillmentAvailability") or payload.get("fulfillment_availability") or []
    found = None
    for row in arr:
        try:
            n = float(row.get("quantity"))
        except (TypeError, ValueError):
            continue
        if found is None or n > found:
            found = n
    return found


def jans_with_api_stock_gt0(child_rows: list[dict], api_qty: dict[str, float], sku_key: str = "子SKU", jan_key: str = "JANコード") -> set[str]:
    """API qty only. Sheet 在庫数 is ignored. Rows without API sku are excluded."""
    jans = set()
    for r in child_rows:
        sku = str(r.get(sku_key) or "").strip()
        jan = str(r.get(jan_key) or "").strip()
        if not sku or len(jan) < 8:
            continue
        if sku not in api_qty:
            continue
        if api_qty[sku] > 0:
            jans.add(jan)
    return jans
