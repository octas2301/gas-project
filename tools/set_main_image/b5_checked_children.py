# -*- coding: utf-8 -*-
"""B-⑤: 親配下の出品CKレ点子SKU一覧（レビュー割当用）。"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from master_sets import SetChildRow, load_checked_set_children_from_rows
from work_paths import DEFAULT_MASTER_CSV

LOG = logging.getLogger("set_main_image.b5_checked_children")


def _from_csv(parent_sku: str, master_csv: Path) -> List[SetChildRow]:
    import csv

    text = Path(master_csv).read_text(encoding="utf-8-sig", errors="replace")
    rows = list(csv.reader(text.splitlines()))
    kids, _ = load_checked_set_children_from_rows(rows, parent_sku=parent_sku)
    return kids


def _from_sheets(parent_sku: str) -> List[SetChildRow]:
    from sheets_master import fetch_master_rows

    rows, _info = fetch_master_rows()
    kids, _ = load_checked_set_children_from_rows(rows, parent_sku=parent_sku)
    return kids


def load_checked_children_for_parent(
    parent_sku: str,
    *,
    master_csv: Optional[Path] = None,
    prefer_sheets: bool = True,
) -> Dict[str, Any]:
    """
    戻り値: { ok, parentSku, children:[{childSku,setCount,unit}], source, error? }
    """
    sku = str(parent_sku or "").strip()
    if not sku:
        return {"ok": False, "parentSku": "", "children": [], "error": "parentSku が空"}

    kids: List[SetChildRow] = []
    source = ""
    err = ""

    if prefer_sheets:
        try:
            kids = _from_sheets(sku)
            source = "sheets"
        except Exception as e:
            LOG.warning("sheets children fail: %s", e)
            err = f"sheets: {e}"

    if not kids:
        csv_path = Path(master_csv or DEFAULT_MASTER_CSV)
        try:
            if csv_path.is_file():
                kids = _from_csv(sku, csv_path)
                source = "csv"
            elif not source:
                err = f"master csv 無し: {csv_path}"
        except Exception as e:
            LOG.warning("csv children fail: %s", e)
            err = (err + " / " if err else "") + f"csv: {e}"

    children = [
        {
            "childSku": r.child_sku,
            "setCount": int(r.set_count),
            "unit": r.unit,
            "label": f"{r.child_sku}（{r.set_count}セット）",
        }
        for r in sorted(kids, key=lambda x: (x.set_count, x.child_sku))
    ]
    ok = bool(children)
    return {
        "ok": ok,
        "parentSku": sku,
        "children": children,
        "source": source,
        "error": "" if ok else (err or "出品CKレ点子が0件"),
    }
