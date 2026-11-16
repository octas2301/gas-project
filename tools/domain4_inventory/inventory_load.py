# -*- coding: utf-8 -*-
"""在庫スプシ読取（F1/F2共通）。書込なし。"""
from __future__ import annotations

from stock_recalc import REQUIRED_INV_SHEETS, InventoryCache, load_products_from_sheet
from sheets_io import list_sheet_titles, read_sheet_values


def build_inventory_cache(svc, inv_id: str) -> InventoryCache:
    titles = set(list_sheet_titles(svc, inv_id))
    missing = [t for t in REQUIRED_INV_SHEETS if t not in titles]
    cache = InventoryCache(missing_sheets=missing)
    if missing:
        return cache
    cache.session_data = read_sheet_values(svc, inv_id, "棚卸セッション")
    cache.count_data = read_sheet_values(svc, inv_id, "棚卸スキャンリスト")
    cache.receiving_data = read_sheet_values(svc, inv_id, "入荷リスト")
    cache.shipping_data = read_sheet_values(svc, inv_id, "出荷リスト")
    products = read_sheet_values(svc, inv_id, "商品マスタ")
    cache.product_jans, cache.product_names = load_products_from_sheet(products)
    return cache
