# -*- coding: utf-8 -*-
"""親SKU単位の共有（文字色・正本・個品見本サマリ）。ALIGN §1／§2／§3／§7.14。

Drive／ローカル正:
  05.画像生成（セットMAIN）/_b5_parent_shared.json

文字色の単体ストア `_b5_text_color_by_parent.json` とも相互同期可。
"""
from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from work_paths import DEFAULT_WORK_ROOT, default_work_root

LOG = logging.getLogger("set_main_image.b5_parent_shared")

SHARED_NAME = "_b5_parent_shared.json"


def shared_path(work_root: Optional[Path] = None) -> Path:
    root = Path(work_root) if work_root is not None else default_work_root()
    return root / SHARED_NAME


def load_parent_shared(work_root: Optional[Path] = None) -> Dict[str, Any]:
    p = shared_path(work_root)
    if not p.is_file():
        return {"schema": "b5_parent_shared_v1", "updatedAt": "", "byParentSku": {}}
    try:
        data = json.loads(p.read_text(encoding="utf-8-sig"))
        if isinstance(data, dict) and isinstance(data.get("byParentSku"), dict):
            return data
    except Exception as e:
        LOG.warning("load_parent_shared fail: %s", e)
    return {"schema": "b5_parent_shared_v1", "updatedAt": "", "byParentSku": {}}


def _norm_rel(rel: str) -> str:
    return str(rel or "").strip().replace("/", "\\")


def upsert_parent_shared(
    *,
    parent_sku: str,
    work_root: Optional[Path] = None,
    jan: str = "",
    name_base: str = "",
    text_color: str = "",
    package_truth_rel: str = "",
    food_canon: Optional[Dict[str, Any]] = None,
) -> Path:
    sku = str(parent_sku or "").strip()
    if not sku:
        raise ValueError("parent_sku required")
    root = Path(work_root) if work_root is not None else default_work_root()
    doc = load_parent_shared(root)
    by = dict(doc.get("byParentSku") or {})
    row = dict(by.get(sku) or {})
    if jan:
        row["jan"] = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
    if name_base:
        row["nameBase"] = str(name_base).strip()
    if text_color and str(text_color).strip() in ("1", "2", "3", "4", "5"):
        row["textColor"] = str(text_color).strip()
    rel = _norm_rel(package_truth_rel)
    if rel:
        row["packageTruthRel"] = rel
    if food_canon and isinstance(food_canon, dict):
        row["foodCanon"] = food_canon
    row["updatedAt"] = datetime.now(timezone.utc).isoformat()
    by[sku] = row
    doc["schema"] = "b5_parent_shared_v1"
    doc["updatedAt"] = row["updatedAt"]
    doc["byParentSku"] = by
    path = shared_path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    # 文字色単体ストアとも同期
    if row.get("textColor"):
        try:
            from b5_set_base_map import save_shared_text_color

            save_shared_text_color(sku, row["textColor"], work_root=root)
        except Exception as e:
            LOG.warning("sync textColor store fail: %s", e)
    LOG.info("parent_shared upsert %s keys=%s", sku, list(row.keys()))
    return path


def seed_from_existing(work_root: Optional[Path] = None) -> Dict[str, Any]:
    """文字色ストア・処理済みマップ・個品 status から親共有を埋める。"""
    from b5_food_canon_status import load_drive_status
    from b5_set_base_map import MAP_SUFFIX, load_shared_text_colors, normalize_text_color

    root = Path(work_root) if work_root is not None else default_work_root()
    rak = root / "02.楽天ベース"
    colors = load_shared_text_colors(root)
    food = load_drive_status(root)
    by_food = food.get("byJan") or {}

    # maps
    for d in (rak, rak / "処理済み"):
        if not d.is_dir():
            continue
        for p in d.glob(f"*{MAP_SUFFIX}"):
            try:
                data = json.loads(p.read_text(encoding="utf-8-sig"))
            except Exception:
                continue
            if not isinstance(data, dict):
                continue
            sku = str(data.get("parentSku") or p.name.replace(MAP_SUFFIX, "")).strip()
            if not sku:
                continue
            jan = ""
            # ishi-{jan}-oya
            if sku.startswith("ishi-") and sku.endswith("-oya"):
                jan = sku[len("ishi-") : -len("-oya")]
            tc = ""
            if data.get("textColor") not in (None, ""):
                tc = normalize_text_color(data.get("textColor"))
            elif sku in colors:
                tc = colors[sku]
            fc = by_food.get(jan) if jan else None
            upsert_parent_shared(
                parent_sku=sku,
                work_root=root,
                jan=jan,
                text_color=tc,
                food_canon=fc if isinstance(fc, dict) else None,
            )

    for sku, tc in colors.items():
        jan = ""
        if sku.startswith("ishi-") and sku.endswith("-oya"):
            jan = sku[len("ishi-") : -len("-oya")]
        fc = by_food.get(jan) if jan else None
        upsert_parent_shared(
            parent_sku=sku,
            work_root=root,
            jan=jan,
            text_color=tc,
            food_canon=fc if isinstance(fc, dict) else None,
        )

    return load_parent_shared(root)


def package_truth_rel_for_parent(
    parent_sku: str, *, work_root: Optional[Path] = None
) -> str:
    sku = str(parent_sku or "").strip()
    if not sku:
        return ""
    row = (load_parent_shared(work_root).get("byParentSku") or {}).get(sku) or {}
    return _norm_rel(str(row.get("packageTruthRel") or ""))
