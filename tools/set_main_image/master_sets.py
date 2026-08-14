# -*- coding: utf-8 -*-
"""マスタCSVからセット行（子SKU・総個数・単位・出品CK）を読む。"""
from __future__ import annotations

import csv
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

LOG = logging.getLogger("set_main_image.master")

FOOD_COL_HINTS = (
    "カテゴリー ※食品のみ入力でOK",
    "カテゴリー※食品のみ入力でOK",
    "カテゴリー",
)
SET_COL_HINTS = (
    "新マスタ(総個数)",  # 正本を優先（「総個数」部分一致より先）
    "総個数",
    "楽天セット数",
    "セット数",
    "セット個数",
    "A.セット商品数",  # 値が「10個で1セット」形式のことがある
)
UNIT_COL_HINTS = (
    "バリエーション単位",  # マスタ正（シートGM列・ヘッダ名で解決）
    "単位",
    "セット単位",
    "表示単位",
    "個数単位",
)
CHECK_COL_HINTS = ("出品CK", "出品チェック", "出力CK", "レ点")
PARENT_HINTS = ("親SKU",)
CHILD_HINTS = ("子SKU",)

DEFAULT_UNIT = "個"


@dataclass
class SetChildRow:
    parent_sku: str
    child_sku: str
    set_count: int
    unit: str
    is_food: bool
    checked: bool = True


def _norm(s: object) -> str:
    return str(s or "").strip()


def checkbox_is_true(cell: object) -> bool:
    """出品CK: boolean true / 1 / 'TRUE'（Yahoo・楽天と同契約）。"""
    if cell is True or cell == 1:
        return True
    return str(cell or "").strip().upper() == "TRUE"


def _find_header_row(rows: List[List[str]]) -> Tuple[int, Dict[str, int]]:
    for i, row in enumerate(rows[:30]):
        cells = [_norm(c) for c in row]
        idx = {cells[j]: j for j in range(len(cells)) if cells[j]}
        if any(h in idx for h in PARENT_HINTS) and any(h in idx for h in CHILD_HINTS):
            return i, idx
    raise ValueError("マスタCSVに親SKU/子SKUの列名行が見つかりません")


def _col(idx: Dict[str, int], hints: Tuple[str, ...]) -> Optional[int]:
    for h in hints:
        if h in idx:
            return idx[h]
    for name, j in idx.items():
        for h in hints:
            if h in name:
                return j
    return None


def _parse_set_count(raw: str) -> Optional[int]:
    """数値、または『10個で1セット』形式の先頭個数をセット数とみなす。"""
    import re

    s = _norm(raw).replace(",", "")
    if not s:
        return None
    try:
        n = int(float(s))
        return n if n >= 1 else None
    except ValueError:
        pass
    m = re.match(r"^(\d+)\s*[個コつ]", s)
    if m:
        n = int(m.group(1))
        return n if n >= 1 else None
    return None


def _read_table(master_csv: Path) -> Tuple[List[List[str]], int, Dict[str, int]]:
    text = master_csv.read_text(encoding="utf-8-sig", errors="replace")
    rows = list(csv.reader(text.splitlines()))
    header_i, idx = _find_header_row(rows)
    return rows, header_i, idx


def load_set_children_for_parent(
    master_csv: Path,
    parent_sku: str,
    *,
    checked_only: bool = False,
) -> Tuple[List[SetChildRow], bool]:
    """
    親SKUに紐づく子行を返す。
    checked_only=True のとき出品CKが真の子のみ（親のみレ点→その親の全子）。
    """
    rows, header_i, idx = _read_table(master_csv)
    i_parent = _col(idx, PARENT_HINTS)
    i_child = _col(idx, CHILD_HINTS)
    i_set = _col(idx, SET_COL_HINTS)
    i_unit = _col(idx, UNIT_COL_HINTS)
    i_food = _col(idx, FOOD_COL_HINTS)
    i_ck = _col(idx, CHECK_COL_HINTS)
    if i_parent is None or i_child is None:
        raise ValueError("親SKU/子SKU列が解決できません")
    if i_set is None:
        raise ValueError("総個数（またはセット数）列が解決できません")
    if i_unit is None:
        LOG.warning(
            "単位列が見つかりません（空→個）。ヘッダ『バリエーション単位』推奨。候補=%s",
            UNIT_COL_HINTS,
        )
    else:
        resolved_name = next((h for h, j in idx.items() if j == i_unit), "unknown")
        LOG.info("unit column resolved header=%r index=%s", resolved_name, i_unit)
    if checked_only and i_ck is None:
        raise ValueError("出品CK列が見つかりません（--checked-only）")

    want = _norm(parent_sku)
    # 親レ点判定用
    parent_checked = False
    parent_food = False
    any_food = False
    children_raw: List[Tuple[str, str, int, str, bool, bool]] = []

    for row in rows[header_i + 1 :]:
        if max(i_parent, i_child, i_set) >= len(row):
            continue
        p = _norm(row[i_parent])
        c = _norm(row[i_child])
        if p != want:
            continue
        food_val = _norm(row[i_food]) if i_food is not None and i_food < len(row) else ""
        is_food = "食品" in food_val
        unit = ""
        if i_unit is not None and i_unit < len(row):
            unit = _norm(row[i_unit])
        if not unit:
            unit = DEFAULT_UNIT
        ck = False
        if i_ck is not None and i_ck < len(row):
            ck = checkbox_is_true(row[i_ck])
        if not c:
            if is_food:
                parent_food = True
            if ck:
                parent_checked = True
            continue
        n = _parse_set_count(row[i_set])
        if n is None:
            LOG.warning("skip child=%s set_count invalid=%r", c, row[i_set])
            continue
        if is_food:
            any_food = True
        children_raw.append((p, c, n, unit, is_food, ck))

    food = parent_food or any_food
    out: List[SetChildRow] = []
    for p, c, n, unit, is_food, ck in children_raw:
        selected = True
        if checked_only:
            # 子に1つ以上レ点 → レ点子のみ / 親のみレ点 → 全子（C画像コースと同契約）
            child_any = any(t[5] for t in children_raw)
            if child_any:
                selected = ck
            else:
                selected = parent_checked
        if not selected:
            continue
        out.append(
            SetChildRow(
                parent_sku=p,
                child_sku=c,
                set_count=n,
                unit=unit,
                is_food=food or is_food,
                checked=ck,
            )
        )

    LOG.info(
        "master parent=%s children=%d food=%s checked_only=%s unit_col=%s ck_col=%s",
        want,
        len(out),
        food,
        checked_only,
        i_unit is not None,
        i_ck is not None,
    )
    return out, food


def load_checked_set_children_from_rows(
    rows: List[List[str]],
    *,
    parent_sku: str = "",
) -> Tuple[List[SetChildRow], Dict[str, bool]]:
    """
    行配列（CSVまたは Sheets 直読）からレ点対象の子行を返す。
    親SKU省略時は全親。戻り値: (rows, parent_food_map)
    """
    header_i, idx = _find_header_row(rows)
    i_parent = _col(idx, PARENT_HINTS)
    i_child = _col(idx, CHILD_HINTS)
    i_set = _col(idx, SET_COL_HINTS)
    i_unit = _col(idx, UNIT_COL_HINTS)
    i_food = _col(idx, FOOD_COL_HINTS)
    i_ck = _col(idx, CHECK_COL_HINTS)
    if i_parent is None or i_child is None or i_set is None:
        raise ValueError("親SKU/子SKU/総個数列が解決できません")
    if i_ck is None:
        raise ValueError("出品CK列が見つかりません")

    set_header = next((h for h, j in idx.items() if j == i_set), "?")
    LOG.info("set column resolved header=%r index=%s", set_header, i_set)

    want = _norm(parent_sku)
    parent_checked: Dict[str, bool] = {}
    parent_food: Dict[str, bool] = {}
    children_by_parent: Dict[str, List[Tuple[str, int, str, bool, bool]]] = {}

    for row in rows[header_i + 1 :]:
        if max(i_parent, i_child, i_set, i_ck) >= len(row):
            continue
        p = _norm(row[i_parent])
        c = _norm(row[i_child])
        if not p:
            continue
        if want and p != want:
            continue
        food_val = _norm(row[i_food]) if i_food is not None and i_food < len(row) else ""
        is_food = "食品" in food_val
        unit = _norm(row[i_unit]) if i_unit is not None and i_unit < len(row) else ""
        if not unit:
            unit = DEFAULT_UNIT
        ck = checkbox_is_true(row[i_ck])
        if not c:
            if ck:
                parent_checked[p] = True
            if is_food:
                parent_food[p] = True
            continue
        n = _parse_set_count(row[i_set])
        if n is None:
            continue
        children_by_parent.setdefault(p, []).append((c, n, unit, is_food, ck))
        if is_food:
            parent_food[p] = True

    out: List[SetChildRow] = []
    for p, kids in children_by_parent.items():
        child_any = any(k[4] for k in kids)
        p_ck = parent_checked.get(p, False)
        food = parent_food.get(p, False) or any(k[3] for k in kids)
        for c, n, unit, is_food, ck in kids:
            if child_any:
                if not ck:
                    continue
            else:
                if not p_ck:
                    continue
            out.append(
                SetChildRow(
                    parent_sku=p,
                    child_sku=c,
                    set_count=n,
                    unit=unit,
                    is_food=food,
                    checked=ck,
                )
            )

    LOG.info(
        "checked children=%d parents=%d filter_parent=%r",
        len(out),
        len({r.parent_sku for r in out}),
        want or "(all)",
    )
    return out, parent_food


def load_checked_set_children(
    master_csv: Path,
    *,
    parent_sku: str = "",
) -> Tuple[List[SetChildRow], Dict[str, bool]]:
    """CSVファイルからレ点対象の子行を返す（親SKU省略時は全親）。"""
    rows, _header_i, _idx = _read_table(master_csv)
    return load_checked_set_children_from_rows(rows, parent_sku=parent_sku)


def resolve_child_by_set_count(
    children: List[SetChildRow],
    set_count: int,
) -> SetChildRow:
    """
    マスタ上のセット数 N に一致する子SKUを1件返す。
    0件・複数件は推測せず ValueError（Vision推定はしない）。
    """
    n = int(set_count)
    if n < 1:
        raise ValueError(f"set_count must be >= 1, got {set_count!r}")
    matches = [c for c in children if int(c.set_count) == n]
    if not matches:
        raise ValueError(f"no child row for set_count={n}")
    if len(matches) > 1:
        skus = ", ".join(c.child_sku for c in matches)
        raise ValueError(
            f"ambiguous set_count={n}: multiple children ({skus}); resolve manually"
        )
    LOG.info(
        "resolve_child_by_set_count n=%s -> child=%s parent=%s",
        n,
        matches[0].child_sku,
        matches[0].parent_sku,
    )
    return matches[0]


JAN_COL_HINTS = ("JANコード", "JAN", "jan")


def parents_for_jan_from_rows(rows: List[List[str]], jan: str) -> List[str]:
    """マスタ行から JAN に紐づく親SKU一覧（重複除去・安定順）。"""
    want = _norm(jan)
    if not want:
        return []
    header_i, idx = _find_header_row(rows)
    i_jan = _col(idx, JAN_COL_HINTS)
    i_parent = _col(idx, PARENT_HINTS)
    if i_jan is None or i_parent is None:
        raise ValueError("マスタに JAN / 親SKU 列がありません")
    found: List[str] = []
    seen = set()
    for row in rows[header_i + 1 :]:
        if max(i_jan, i_parent) >= len(row):
            continue
        if _norm(row[i_jan]) != want:
            continue
        p = _norm(row[i_parent])
        if not p or p in seen:
            continue
        seen.add(p)
        found.append(p)
    return found


def resolve_checked_children_for_jan(
    jan: str,
    *,
    master_csv: Optional[Path] = None,
    rows: Optional[List[List[str]]] = None,
) -> List[SetChildRow]:
    """
    JAN → 親SKU → 出品CK対象の子SKU一覧。
    本番の `--to-checked-children` 用。目視 auto-export は品番キー1件（全子FO禁止）。
    """
    table = rows
    if table is None:
        if not master_csv:
            raise ValueError("master_csv or rows required")
        table, _h, _i = _read_table(Path(master_csv))
    parents = parents_for_jan_from_rows(table, jan)
    if not parents:
        LOG.warning("JAN=%s に紐づく親SKUがマスタにありません", jan)
        return []
    out: List[SetChildRow] = []
    seen_child = set()
    for p in parents:
        kids, _ = load_checked_set_children_from_rows(table, parent_sku=p)
        for k in kids:
            if k.child_sku in seen_child:
                continue
            seen_child.add(k.child_sku)
            out.append(k)
    LOG.info(
        "JAN=%s parents=%s checked_children=%s",
        jan,
        parents,
        [c.child_sku for c in out],
    )
    return out
