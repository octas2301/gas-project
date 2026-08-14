#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""B-T1: xlsm ヘッダ項目名から論理キー→列番号を解決する。"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple


def _norm(s: Any) -> str:
    if s is None:
        return ""
    t = str(s).replace("\r", " ").replace("\n", " ").strip()
    while "  " in t:
        t = t.replace("  ", " ")
    return t


def build_header_index(
    ws: Any,
    rows: List[int],
    max_col: int,
) -> Dict[str, List[int]]:
    """正規化ヘッダ文字列 → 列番号リスト（1始まり）。"""
    index: Dict[str, List[int]] = {}
    for c in range(1, max_col + 1):
        for r in rows:
            key = _norm(ws.cell(r, c).value)
            if not key:
                continue
            index.setdefault(key, [])
            if c not in index[key]:
                index[key].append(c)
    return index


def resolve_cols_by_aliases(
    ws: Any,
    aliases: Dict[str, List[str]],
    header_rows: Optional[List[int]] = None,
    max_col: Optional[int] = None,
    legacy_cols: Optional[Dict[str, int]] = None,
) -> Tuple[Dict[str, int], List[dict]]:
    """
    戻り値: (resolved_cols, gap_rows)
    gap_rows の各要素: logicalKey / searched / hitCol / status / legacyCol
    """
    rows = list(header_rows or [4, 5])
    mc = int(max_col or ws.max_column or 300)
    index = build_header_index(ws, rows, mc)
    legacy = {str(k): int(v) for k, v in (legacy_cols or {}).items()}

    resolved: Dict[str, int] = {}
    gaps: List[dict] = []

    for logical, names in (aliases or {}).items():
        searched = [_norm(n) for n in (names or []) if _norm(n)]
        hit: Optional[int] = None
        matched_alias = ""
        for alias in searched:
            cols = index.get(alias) or []
            if not cols:
                continue
            # 同一表示名が複数列にある場合は先頭（属性IDで一意になる想定）
            hit = cols[0]
            matched_alias = alias
            break

        leg = legacy.get(logical)
        if hit is None:
            status = "MISS"
        elif leg is not None and hit != leg:
            status = "HIT_DIFF_LEGACY"
        else:
            status = "HIT"

        gaps.append(
            {
                "logicalKey": logical,
                "searched": searched,
                "matchedAlias": matched_alias,
                "hitCol": hit,
                "legacyCol": leg,
                "status": status,
                "required": logical
                in (
                    "sku",
                    "product_type",
                    "parentage",
                    "title",
                    "browse",
                    "mfr_name",
                    "main_image_url",
                    "tax_code",
                    "price",
                    "ingredients",
                    "unit_count",
                    "unit_uom",
                    "origin",
                ),
            }
        )
        if hit is not None:
            resolved[logical] = hit

    return resolved, gaps
