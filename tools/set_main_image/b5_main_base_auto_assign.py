# -*- coding: utf-8 -*-
"""B-⑤ MAIN レビュー — 訴求順のレ点子自動割当。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from b5_grid_budget import classify_grid_tier
from b5_set_base_map import picture_family_key


def _proposal_map(meta: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for row in meta.get("proposals") or []:
        if isinstance(row, dict):
            tid = str(row.get("themeId") or "").strip()
            if tid:
                out[tid] = row
    return out


def _summary_by_theme(compose_dir: Path) -> Dict[str, Dict[str, Any]]:
    p = Path(compose_dir) / "summary.json"
    if not p.is_file():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    out: Dict[str, Dict[str, Any]] = {}
    for row in data.get("results") or []:
        if isinstance(row, dict):
            tid = str(row.get("themeId") or "").strip()
            if tid:
                out[tid] = row
    return out


def appeal_score_for_card(
    row: Dict[str, Any],
    *,
    proposal: Optional[Dict[str, Any]] = None,
    summary_row: Optional[Dict[str, Any]] = None,
) -> int:
    """訴求スコア（大きいほど上位）。文案 sellRank + 生成品質。"""
    pr = proposal or {}
    rank = int(pr.get("sellRank") or row.get("sellRank") or 99)
    sell = int(pr.get("sellScore") or 0)
    score = 1000 - rank * 100 + min(sell, 120)

    catch = str(row.get("catchHead") or row.get("catch") or pr.get("catch") or "")
    if catch.startswith("【") and "】" in catch:
        score += 8
    if str(row.get("salesTalk") or pr.get("salesTalk") or "").strip():
        score += 4

    sr = summary_row or {}
    grid = sr.get("grid20") or {}
    hit = int(grid.get("hitN") or sr.get("hitN") or 0)
    tier = classify_grid_tier(hit)
    if tier == "ok":
        score += 25
    elif tier == "buffer":
        score += 8
    elif tier == "retry":
        score -= 40
    if sr.get("ok") is False:
        score -= 200
    return score


def _load_adopted_by_child(compose_dir: Optional[Path]) -> Dict[str, str]:
    """adopted.json の byChildSku（過去に02へ出した採用）。無ければ空。"""
    if compose_dir is None:
        return {}
    p = Path(compose_dir) / "adopted.json"
    if not p.is_file():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return {}
    raw = data.get("byChildSku") if isinstance(data, dict) else None
    if not isinstance(raw, dict):
        return {}
    out: Dict[str, str] = {}
    for sku, akey in raw.items():
        s = str(sku or "").strip()
        a = str(akey or "").strip()
        if s and a:
            out[s] = a
    return out


def _pick_card_in_family(
    group: List[Dict[str, Any]],
    *,
    preferred_akeys: set,
) -> Dict[str, Any]:
    """
    同じ見出しの候補から1枚。
    1) 過去採用（adopted.json）にあればそれ
    2) 作り直し（__rN）があれば番号が新しい方
    3) なければ先頭（初回）
    ※旧挙動の「常に初回を優先」は、過去レ点が初回に戻る原因だったので廃止。
    """
    if not group:
        raise ValueError("empty group")
    hit = [c for c in group if str(c.get("adoptKey") or "") in preferred_akeys]
    if hit:
        return hit[0]
    remakes = [c for c in group if "__r" in str(c.get("adoptKey") or "")]
    if remakes:
        def _rnum(c: Dict[str, Any]) -> int:
            a = str(c.get("adoptKey") or "")
            try:
                return int(a.rsplit("__r", 1)[-1])
            except ValueError:
                return 0

        remakes.sort(key=_rnum, reverse=True)
        return remakes[0]
    return group[0]


def build_auto_assign_plan(
    rows: List[Dict[str, Any]],
    meta: Dict[str, Any],
    children: List[Dict[str, Any]],
    *,
    compose_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """全レ点子を訴求順テーマへ割当。同じ見出しの再作成前後は、残す1枚だけ。"""
    props = _proposal_map(meta)
    summary = _summary_by_theme(compose_dir) if compose_dir else {}
    adopted_by_child = _load_adopted_by_child(compose_dir)
    preferred_akeys = set(adopted_by_child.values())

    scored: List[Dict[str, Any]] = []
    for row in rows:
        tid = str(row.get("themeId") or "").strip()
        akey = str(row.get("adoptKey") or tid or "?")
        appeal = appeal_score_for_card(
            row,
            proposal=props.get(tid),
            summary_row=summary.get(tid),
        )
        scored.append(
            {
                "adoptKey": akey,
                "themeId": tid,
                "fileName": str(row.get("name") or ""),
                "themeName": str(
                    row.get("themeName") or props.get(tid, {}).get("themeName") or ""
                ),
                "appeal": appeal,
                "sellRank": int(props.get(tid, {}).get("sellRank") or 99),
            }
        )
    scored.sort(key=lambda x: (-int(x["appeal"]), int(x["sellRank"]), x["adoptKey"]))
    akeys_present = {str(c["adoptKey"]) for c in scored}

    # テーマ番号ではなく見出し（family）単位。T4が複数あっても別絵として残す。
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for card in scored:
        fam = picture_family_key(str(card.get("adoptKey") or ""))
        key = fam or str(card.get("adoptKey") or card.get("themeId") or "?")
        grouped.setdefault(key, []).append(card)
    chosen: List[Dict[str, Any]] = []
    for group in grouped.values():
        chosen.append(_pick_card_in_family(group, preferred_akeys=preferred_akeys))
    chosen.sort(key=lambda x: (-int(x["appeal"]), int(x["sellRank"]), x["adoptKey"]))

    kids = sorted(
        [c for c in children if str(c.get("childSku") or "").strip()],
        key=lambda c: (-int(c.get("setCount") or 0), str(c.get("childSku") or "")),
    )
    kid_skus = [str(c.get("childSku") or "").strip() for c in kids]
    kid_set = set(kid_skus)

    buckets: Dict[str, List[str]] = {c["adoptKey"]: [] for c in scored}
    used_skus: set = set()

    # 過去の02採用を最優先（画面を開いたとき初回に戻さない）
    for sku, akey in adopted_by_child.items():
        if sku not in kid_set or akey not in akeys_present:
            continue
        if sku in used_skus:
            continue
        if buckets.get(akey):
            continue
        buckets[akey] = [sku]
        used_skus.add(sku)

    if chosen and kids:
        # 画像1枚につき子SKUは最大1。余った子は未割当のまま。
        free_kids = [s for s in kid_skus if s not in used_skus]
        free_cards = [c for c in chosen if not buckets.get(c["adoptKey"])]
        n = min(len(free_cards), len(free_kids))
        for i in range(n):
            buckets[free_cards[i]["adoptKey"]].append(free_kids[i])
            used_skus.add(free_kids[i])

    assignments: List[Dict[str, Any]] = []
    for card in scored:
        skus = list(buckets.get(card["adoptKey"]) or [])[:1]
        assignments.append(
            {
                **card,
                "adopt": bool(skus),
                "childSku": skus[0] if skus else "",
                "childSkus": skus,
            }
        )

    assigned_n = sum(len(a["childSkus"]) for a in assignments)
    return {
        "assignments": assignments,
        "childOrder": [c["childSku"] for c in kids],
        "assignedChildN": assigned_n,
        "childN": len(kids),
        "fromAdoptedN": sum(
            1
            for sku, akey in adopted_by_child.items()
            if sku in used_skus and akey in akeys_present
        ),
        "note": (
            "画像1枚につき子SKUは1つ。"
            "テーマ番号が同じでも見出しが違えば別の絵。"
            "再作成前後は同じ見出しの1枚だけ。"
            "adopted.json があればそれを優先（初回案へ戻さない）。"
        ),
    }
