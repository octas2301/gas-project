# -*- coding: utf-8 -*-
"""B-⑤ キャッチ置き場 — ピッカー採用 JSON の読み書き。"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT

LOG = logging.getLogger("set_main_image.b5_catch_zone_adopted")

ADOPTED_FILENAME = "b5_catch_zone_adopted.json"

DEFAULT_PICKER_DIR = (
    Path(DEFAULT_WORK_ROOT)
    / SUB_TEST_OUT
    / "b5_main_base_compose"
    / "_catch_zone_picker"
)


def adopted_path(picker_dir: Optional[Path] = None) -> Path:
    return (picker_dir or DEFAULT_PICKER_DIR) / ADOPTED_FILENAME


def load_adopted(picker_dir: Optional[Path] = None) -> Optional[Dict[str, Any]]:
    path = adopted_path(picker_dir)
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        LOG.warning("adopted read failed path=%s err=%s", path, exc)
        return None
    if not isinstance(data, dict):
        return None
    return data


def save_adopted(
    data: Dict[str, Any],
    picker_dir: Optional[Path] = None,
) -> Path:
    path = adopted_path(picker_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    LOG.info("adopted saved path=%s patterns=%s", path, len(data.get("patterns") or []))
    return path


def adopted_patterns(data: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    payload = data if data is not None else load_adopted()
    if not payload:
        return []
    raw = payload.get("patterns") or []
    return [p for p in raw if isinstance(p, dict)]


def normalize_vertical_pattern(p: Dict[str, Any]) -> Dict[str, Any]:
    """縦書き採用を範囲のみ表現に正規化（旧版の赤/青 zones は捨てる）。"""
    cells = p.get("cells") or []
    if not cells:
        return p
    min_c = min(int(c) for c, _ in cells)
    max_c = max(int(c) for c, _ in cells)
    min_r = min(int(r) for _, r in cells)
    max_r = max(int(r) for _, r in cells)
    out = dict(p)
    out["zoneMode"] = "range-only"
    out["catchN"] = None
    out["salesN"] = None
    out["zones"] = [
        {
            "role": "range",
            "c0": min_c,
            "r0": min_r,
            "w": max_c - min_c + 1,
            "h": max_r - min_r + 1,
            "label": f"col {min_c}–{max_c}, row {min_r}–{max_r}",
        }
    ]
    return out


def adopted_vertical_ranges(
    data: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for p in adopted_patterns(data):
        if p.get("writingMode") != "vertical-rl":
            continue
        if p.get("zoneMode") == "range-only":
            out.append(p)
        else:
            out.append(normalize_vertical_pattern(p))
    return out


def adopted_horizontal_splits(
    data: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    return [
        p
        for p in adopted_patterns(data)
        if p.get("zoneMode") != "range-only"
        and p.get("writingMode") != "vertical-rl"
    ]


def adopted_all_for_compose(
    data: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """compose 用: 縦=範囲正規化済み + 横=split の順。"""
    payload = data if data is not None else load_adopted()
    if not payload:
        return []
    vertical = adopted_vertical_ranges(payload)
    horizontal = adopted_horizontal_splits(payload)
    return vertical + horizontal


def pick_adopted_for_theme(
    theme_index: int,
    data: Optional[Dict[str, Any]] = None,
) -> Optional[Dict[str, Any]]:
    pool = adopted_all_for_compose(data)
    if not pool:
        return None
    return pool[int(theme_index) % len(pool)]


# 強い和風シグナルのみ（国産・晩酌・燻製等は和風書字の根拠にしない）
_WAFU_MARKERS = (
    "和風",
    "和食",
    "料亭",
    "懐石",
    "割烹",
    "出汁",
    "おだし",
    "お茶漬け",
    "茶漬け",
    "冷奴",
    "掛け軸",
    "昆布",
    "味噌",
    "佃煮",
    "醤油",
    "おせち",
)


def infer_wafu_score(materials: Optional[Dict[str, Any]]) -> float:
    """0..1。和風トーンの商品かどうか（縦書き採用の可否判定用）。

    縦書き＝おしゃれ、ではない。スコアは書字方向の候補判定のみに使う。
    """
    if not materials:
        return 0.0
    parts = [
        str(materials.get("name") or ""),
        str(materials.get("feats") or ""),
        str(materials.get("kw") or ""),
        str(materials.get("rakuten") or ""),
        str(materials.get("yahoo") or ""),
        str(materials.get("desc") or ""),
        " ".join(str(b) for b in (materials.get("bullets") or [])),
    ]
    blob = " ".join(parts)
    hits = sum(1 for m in _WAFU_MARKERS if m in blob)
    if hits <= 0:
        return 0.0
    # マーカー1つ（例: おだし）で和風扱いの下限 0.45
    return min(1.0, 0.45 + (hits - 1) * 0.15)


def plan_writing_modes(n_themes: int, wafu_score: float) -> List[str]:
    """書字方向プラン。

    - 非和風: すべて横書き
    - 和風: 縦書きは少数のみ（最大 n//3・上限2）。横書きを既定のおしゃれ側とする
    """
    n = max(0, int(n_themes))
    if n <= 0:
        return []
    if wafu_score < 0.4:
        return ["horizontal-tb"] * n
    if n == 1:
        # 1案だけなら横でおしゃれに寄せる（縦＝必須ではない）
        return ["horizontal-tb"]
    if n == 2:
        v_target = 1
    else:
        v_target = min(2, max(1, n // 3))
    v_target = min(v_target, n - 1)
    h_target = n - v_target
    modes: List[str] = []
    v_left, h_left = v_target, h_target
    for i in range(n):
        # 横を先行。縦は 1,4,7… 番目に間引き配置
        pick_v = v_left > 0 and (h_left <= 0 or (i % 3 == 1))
        if pick_v:
            modes.append("vertical-rl")
            v_left -= 1
        elif h_left > 0:
            modes.append("horizontal-tb")
            h_left -= 1
        else:
            modes.append("vertical-rl")
            v_left -= 1
    return modes


def pick_adopted_for_theme_balanced(
    theme_index: int,
    n_themes: int,
    materials: Optional[Dict[str, Any]] = None,
    data: Optional[Dict[str, Any]] = None,
) -> Optional[Dict[str, Any]]:
    """商品特徴に応じた縦横バランスで採用パターンを選ぶ。"""
    payload = data if data is not None else load_adopted()
    if not payload:
        return None
    modes = plan_writing_modes(n_themes, infer_wafu_score(materials))
    idx = int(theme_index) % max(1, len(modes))
    mode = modes[idx] if modes else "horizontal-tb"
    vertical = adopted_vertical_ranges(payload)
    horizontal = adopted_horizontal_splits(payload)
    pool = vertical if mode == "vertical-rl" else horizontal
    if not pool:
        pool = adopted_all_for_compose(payload)
    if not pool:
        return None
    same_before = sum(1 for j in range(idx) if j < len(modes) and modes[j] == mode)
    return pool[same_before % len(pool)]


def cell_key(cells: List[List[int]]) -> str:
    return ";".join(f"{int(c)},{int(r)}" for c, r in sorted(cells, key=lambda x: (x[1], x[0])))


def pattern_cells(p: Dict[str, Any]) -> Set[Tuple[int, int]]:
    out: Set[Tuple[int, int]] = set()
    for item in p.get("cells") or []:
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            continue
        out.add((int(item[0]), int(item[1])))
    return out


def pattern_zone_sets(
    p: Dict[str, Any],
    *,
    has_sales: bool = True,
) -> Tuple[Set[Tuple[int, int]], Set[Tuple[int, int]], Set[Tuple[int, int]]]:
    """Return (catch_cells, sales_cells, all_cells) from an adopted pattern."""
    all_cells = pattern_cells(p)
    catch: Set[Tuple[int, int]] = set()
    sales: Set[Tuple[int, int]] = set()
    for z in p.get("zones") or []:
        if not isinstance(z, dict):
            continue
        c0 = int(z.get("c0", 0))
        r0 = int(z.get("r0", 0))
        w = int(z.get("w", 0))
        h = int(z.get("h", 0))
        role = str(z.get("role") or "")
        rect = {
            (c0 + dc, r0 + dr)
            for dr in range(max(0, h))
            for dc in range(max(0, w))
        }
        if role == "catch":
            catch |= rect
        elif role == "sales":
            sales |= rect
    wm = str(p.get("writingMode") or "")
    if wm == "vertical-rl" or p.get("zoneMode") == "range-only":
        catch = catch or all_cells
        if has_sales:
            sales = sales or all_cells
    else:
        if not catch:
            catch = all_cells
        if has_sales and not sales:
            sales = all_cells - catch if catch else all_cells
    return catch, sales, all_cells


def pattern_by_id(
    pattern_id: str,
    data: Optional[Dict[str, Any]] = None,
) -> Optional[Dict[str, Any]]:
    needle = str(pattern_id or "").strip()
    if not needle:
        return None
    for p in adopted_patterns(data):
        if str(p.get("id") or "") == needle:
            return p
    return None


def find_nearest_adopted_pattern(
    reference_cells: Iterable[Tuple[int, int]],
    *,
    data: Optional[Dict[str, Any]] = None,
    prefer_writing_mode: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    """期待帯域（文字長ベース等）に最も近い採用マスタパターンを返す。"""
    ref = frozenset((int(c), int(r)) for c, r in reference_cells)
    if not ref:
        return None
    pool = adopted_patterns(data)
    if not pool:
        return None
    ref_n = len(ref)
    best: Optional[Dict[str, Any]] = None
    best_score = float("-inf")
    for p in pool:
        cells = pattern_cells(p)
        if not cells:
            continue
        inter = len(ref & cells)
        union = len(ref | cells)
        jaccard = inter / union if union else 0.0
        count_penalty = abs(len(cells) - ref_n) * 0.05
        mode_bonus = (
            0.02
            if prefer_writing_mode
            and p.get("writingMode") == prefer_writing_mode
            else 0.0
        )
        score = jaccard - count_penalty + mode_bonus
        if score > best_score:
            best_score = score
            best = p
    return best


def _guide_meta_from_pattern(
    p: Dict[str, Any],
    *,
    source: str,
) -> Dict[str, Any]:
    cells = sorted(pattern_cells(p), key=lambda x: (x[1], x[0]))
    return {
        "id": p.get("id"),
        "cells": [list(c) for c in cells],
        "writingMode": p.get("writingMode"),
        "cellN": p.get("cellN"),
        "source": source,
    }


def adopted_guide_for_theme(
    *,
    theme_id: str,
    summary: Optional[Dict[str, Any]] = None,
    expected_cells: Optional[Iterable[Tuple[int, int]]] = None,
    data: Optional[Dict[str, Any]] = None,
) -> Optional[Dict[str, Any]]:
    """Review UI 緑枠用。summary.adoptedZone.cells 優先、無ければ compose id → 最近傍。"""
    payload = data if data is not None else load_adopted()
    adopted_id = ""
    prefer_mode: Optional[str] = None
    summary_cells: Optional[List[List[int]]] = None
    summary_cell_n: Optional[Any] = None
    if summary:
        for row in summary.get("results") or []:
            if not isinstance(row, dict):
                continue
            if str(row.get("themeId") or "") != str(theme_id):
                continue
            az = row.get("adoptedZone") or {}
            if isinstance(az, dict):
                adopted_id = str(az.get("id") or "").strip()
                prefer_mode = az.get("writingMode")
                raw_cells = az.get("cells")
                if isinstance(raw_cells, list) and raw_cells:
                    summary_cells = raw_cells
                    summary_cell_n = az.get("cellN")
            break
    if summary_cells:
        cells = sorted(
            (
                (int(item[0]), int(item[1]))
                for item in summary_cells
                if isinstance(item, (list, tuple)) and len(item) >= 2
            ),
            key=lambda x: (x[1], x[0]),
        )
        return {
            "id": adopted_id or None,
            "cells": [list(c) for c in cells],
            "writingMode": prefer_mode,
            "cellN": summary_cell_n,
            "source": "compose",
        }
    if adopted_id:
        p = pattern_by_id(adopted_id, payload)
        if p:
            return _guide_meta_from_pattern(p, source="compose")
    ref = list(expected_cells or [])
    if not ref:
        return None
    p = find_nearest_adopted_pattern(
        ref,
        data=payload,
        prefer_writing_mode=prefer_mode,
    )
    if not p:
        return None
    return _guide_meta_from_pattern(p, source="nearest")
