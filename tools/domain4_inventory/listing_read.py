# -*- coding: utf-8 -*-
"""出品マスタ読取（子SKU行）。書込なし。"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from stock_recalc import norm_key


LISTING_SHEET = "▼商品マスタ(人間作業用)"

# セット数列の候補（先勝ち）
SET_COUNT_HEADERS = (
    "A.セット商品数",
    "セット商品数",
    "セット数",
)

# ミックス構成らしき列（値があれば P6 延期）
MIX_HEADER_HINTS = (
    "セット構成",
    "構成1",
    "構成2",
    "構成品",
    "combo",
    "COMBO",
    "ミックス構成",
)


@dataclass
class ListingChildRow:
    row_num: int  # 1-based sheet row
    child_sku: str
    parent_sku: str
    stock_key: str  # 代表JAN / 社内キー候補
    set_count_raw: Any
    product_name: str
    is_mix_candidate: bool
    mix_hint: str


def _header_map(header_row: list) -> dict[str, int]:
    out: dict[str, int] = {}
    for i, c in enumerate(header_row):
        name = str(c or "").strip()
        if name and name not in out:
            out[name] = i
    return out


def _find_col(hmap: dict[str, int], candidates: tuple[str, ...]) -> int:
    for name in candidates:
        if name in hmap:
            return hmap[name]
    # 部分一致フォールバック
    for name, idx in hmap.items():
        for cand in candidates:
            if cand in name:
                return idx
    return -1


def _cell(row: list, idx: int) -> Any:
    if idx < 0 or idx >= len(row):
        return ""
    return row[idx]


def _is_mix_row(hmap: dict[str, int], row: list) -> tuple[bool, str]:
    for name, idx in hmap.items():
        if any(h in name for h in MIX_HEADER_HINTS):
            v = _cell(row, idx)
            if v is not None and str(v).strip() != "":
                return True, name
    return False, ""


def find_header_row_index(rows: list[list], scan_max: int = 30) -> int:
    """
    出品マスタは先頭に説明行・Amazon見出しがあり、
    「子SKU」「親SKU」が揃う行が本物のヘッダー（実測: 8行目）。
    """
    limit = min(len(rows), scan_max)
    for i in range(limit):
        names = {str(c or "").strip() for c in rows[i]}
        if "子SKU" in names and "親SKU" in names:
            return i
    for i in range(limit):
        names = {str(c or "").strip() for c in rows[i]}
        if "子SKU" in names:
            return i
    return -1


def iter_child_rows(rows: list[list]) -> tuple[list[ListingChildRow], dict]:
    """
    ヘッダー行を自動検出。子SKUが空の親行は除外。
    """
    meta: dict = {"headers": [], "set_col": "", "jan_col": "", "child_col": "", "header_row": 0}
    if not rows:
        return [], meta
    h_i = find_header_row_index(rows)
    if h_i < 0:
        meta["error"] = "child_sku_column_missing"
        return [], meta
    meta["header_row"] = h_i + 1  # 1-based
    hmap = _header_map(rows[h_i])
    meta["headers"] = list(hmap.keys())
    i_child = _find_col(hmap, ("子SKU",))
    i_parent = _find_col(hmap, ("親SKU",))
    i_jan = _find_col(hmap, ("JANコード", "JAN", "単品JAN"))
    i_set = _find_col(hmap, SET_COUNT_HEADERS)
    i_name = _find_col(hmap, ("商品名ベース", "商品名"))
    meta["child_col"] = "子SKU" if i_child >= 0 else ""
    meta["jan_col"] = next((k for k, v in hmap.items() if v == i_jan), "") if i_jan >= 0 else ""
    meta["set_col"] = next((k for k, v in hmap.items() if v == i_set), "") if i_set >= 0 else ""

    out: list[ListingChildRow] = []
    if i_child < 0:
        meta["error"] = "child_sku_column_missing"
        return out, meta

    for r_i, row in enumerate(rows[h_i + 1 :], start=h_i + 2):
        child = norm_key(_cell(row, i_child))
        if not child:
            continue  # 親行除外
        parent = norm_key(_cell(row, i_parent)) if i_parent >= 0 else ""
        stock_key = norm_key(_cell(row, i_jan)) if i_jan >= 0 else ""
        set_raw = _cell(row, i_set) if i_set >= 0 else ""
        name = str(_cell(row, i_name) or "") if i_name >= 0 else ""
        is_mix, mix_hint = _is_mix_row(hmap, row)
        out.append(
            ListingChildRow(
                row_num=r_i,
                child_sku=child,
                parent_sku=parent,
                stock_key=stock_key,
                set_count_raw=set_raw,
                product_name=name,
                is_mix_candidate=is_mix,
                mix_hint=mix_hint,
            )
        )
    return out, meta
