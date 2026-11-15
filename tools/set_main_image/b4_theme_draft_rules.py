# -*- coding: utf-8 -*-
"""B-⑤ テーマ案下書きの契約（P1）。GASと同趣旨。画像・Geminiなし。

正本: docs/org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md §4.5.7
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence, Set


SCHEMA = "b4_theme_draft_v1"
MAX_THEMES = 10  # サブ通常上限（子レ点件数連動しない）
# WRAPあり時は個包装訴求(+1)で最大11。正本: B45 §4.5.11 / b4_wrap_asset_rules.py


def draft_file_stem(parent_sku: str = "", jan: str = "", name_base: str = "") -> str:
    sku = str(parent_sku or "").strip()
    if sku:
        return _safe_name(sku)
    jan_s = str(jan or "").strip()
    nb = str(name_base or "").strip()
    if jan_s and nb:
        return _safe_name(f"{jan_s}__{nb}")
    if jan_s:
        return _safe_name(jan_s)
    if nb:
        return _safe_name(nb)
    return ""


def _safe_name(s: str) -> str:
    out = []
    for ch in s:
        if ch in '\\/:*?"<>|':
            out.append("_")
        else:
            out.append(ch)
    return "".join(out).strip() or ""


def is_proposal_full(pr: Optional[Dict[str, Any]]) -> bool:
    """丸ごと再利用に足りるか（catchだけ等の薄い行は不可）。"""
    if not isinstance(pr, dict):
        return False
    tid = pr.get("themeId")
    if tid is None or str(tid).strip() == "":
        return False
    ov = str(pr.get("overview") or "").strip()
    catch = str(pr.get("catch") or "").strip()
    # overview か catch のどちらか必須＋両方あるのが望ましいが、最低 catch
    if not catch and not ov:
        return False
    # 部分 overlay 疑い: catch だけあって overview キー自体が無い薄い行は不可
    if catch and "overview" not in pr and "subs" not in pr and "needPackage" not in pr:
        return False
    return True


def is_draft_reusable(draft: Optional[Dict[str, Any]]) -> bool:
    if not isinstance(draft, dict):
        return False
    props = draft.get("proposals") or []
    if not isinstance(props, list) or not props:
        return False
    full = [p for p in props if is_proposal_full(p)]
    return len(full) >= 1


def reject_partial_overlay(incoming: Sequence[Dict[str, Any]]) -> bool:
    """True＝拒否（部分 overlay）。丸ごと行だけ許可。"""
    if not incoming:
        return True
    for pr in incoming:
        if not is_proposal_full(pr):
            return True
    return False


def normalize_selected(proposals: Sequence[Dict[str, Any]], *, force_all_on: bool = False) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for pr in proposals or []:
        row = dict(pr)
        if force_all_on:
            row["selected"] = True
            row["checkedDefault"] = True
        else:
            sel = row.get("selected")
            if sel is None:
                sel = row.get("checkedDefault")
            if sel is None:
                sel = True
            row["selected"] = bool(sel)
            row["checkedDefault"] = bool(sel)
        out.append(row)
    return out


def remake_indexes_unchecked(proposals: Sequence[Dict[str, Any]]) -> List[int]:
    """レ点が外れている（selected/checkedDefault が false）案の index。"""
    idxs: List[int] = []
    for i, pr in enumerate(proposals or []):
        sel = pr.get("selected")
        if sel is None:
            sel = pr.get("checkedDefault", True)
        if not bool(sel):
            idxs.append(i)
    return idxs


def apply_remake(
    existing: Sequence[Dict[str, Any]],
    remade_by_index: Dict[int, Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """外した index だけ差し替え。他は保持。差し替え後は selected=true。"""
    out: List[Dict[str, Any]] = []
    for i, pr in enumerate(existing or []):
        if i in remade_by_index:
            row = dict(remade_by_index[i])
            # themeId は既存を維持（差し替え側が空なら）
            if row.get("themeId") in (None, ""):
                row["themeId"] = pr.get("themeId")
            if not row.get("name"):
                row["name"] = pr.get("name")
            row["selected"] = True
            row["checkedDefault"] = True
            out.append(row)
        else:
            out.append(dict(pr))
    return normalize_selected(out, force_all_on=False)


def merge_by_theme_id(
    existing: Sequence[Dict[str, Any]],
    incoming: Sequence[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """themeId が正。同 themeId は上書き。無い themeId は末尾追加。件数は MAX_THEMES まで切らない（呼び出し側）。"""
    by: Dict[str, Dict[str, Any]] = {}
    order: List[str] = []
    for pr in existing or []:
        tid = str((pr or {}).get("themeId") or "").strip()
        if not tid:
            continue
        if tid not in by:
            order.append(tid)
        by[tid] = dict(pr)
    for pr in incoming or []:
        tid = str((pr or {}).get("themeId") or "").strip()
        if not tid:
            continue
        if tid not in by:
            order.append(tid)
        by[tid] = dict(pr)
    return [by[k] for k in order if k in by]


def norm_name_base(s: str) -> str:
    return "".join(str(s or "").split())


def draft_matches_binding(
    draft: Optional[Dict[str, Any]],
    *,
    parent_sku: str = "",
    jan: str = "",
    name_base: str = "",
) -> bool:
    """親SKU・JAN・商品名ベースが一致しない下書きは使わない。"""
    if not isinstance(draft, dict):
        return False
    b_sku = str(parent_sku or "").strip()
    d_sku = str(draft.get("parentSku") or "").strip()
    if b_sku:
        if not d_sku or d_sku != b_sku:
            return False
    b_jan = str(jan or "").strip()
    d_jan = str(draft.get("jan") or "").strip()
    if b_jan and d_jan and d_jan != b_jan:
        return False
    b_nb = norm_name_base(name_base)
    d_nb = norm_name_base(str(draft.get("nameBase") or ""))
    if b_nb:
        if not d_nb or d_nb != b_nb:
            return False
    return True


def decide_open_mode(
    draft: Optional[Dict[str, Any]],
    *,
    force_full: bool = False,
    parent_sku: str = "",
    jan: str = "",
    name_base: str = "",
) -> str:
    """reuse | generate"""
    if force_full:
        return "generate"
    if not is_draft_reusable(draft):
        return "generate"
    if parent_sku or name_base or jan:
        if not draft_matches_binding(
            draft, parent_sku=parent_sku, jan=jan, name_base=name_base
        ):
            return "generate"
    return "reuse"


def duplicate_catch_indexes(proposals: Sequence[Dict[str, Any]]) -> List[int]:
    """先勝ち。後から同じ catch の index を返す（1回直し対象）。"""
    seen: Set[str] = set()
    dups: List[int] = []
    for i, pr in enumerate(proposals or []):
        c = str((pr or {}).get("catch") or "").strip()
        if not c:
            continue
        if c in seen:
            dups.append(i)
        else:
            seen.add(c)
    return dups


def clip_theme_count(proposals: Sequence[Dict[str, Any]], limit: int = MAX_THEMES) -> List[Dict[str, Any]]:
    """子レ点には連動しない。テーマ上限のみ。"""
    return list(proposals or [])[: max(0, int(limit))]
