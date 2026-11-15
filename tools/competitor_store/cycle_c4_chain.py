# -*- coding: utf-8 -*-
"""C4 完走連鎖ロジック（副作用なし）。正本: B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md §6.4。"""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from typing import Any

from cycle_c3_select import pick_cycle_maker_select

PR_WASH_ST_DONE = "完了"


def d_step_success(step: dict | None) -> bool:
    """D1/D2 1段が機械的成功か（state=DONE）。"""
    if not step:
        return False
    return str(step.get("state") or "").upper() == "DONE"


def d_chain_success(d1: dict | None, d2: dict | None) -> bool:
    """D1→D2 連鎖成功: 両方 DONE かつ D2 leftover=0。"""
    if not d_step_success(d1) or not d_step_success(d2):
        return False
    leftover = int(d2.get("leftover") or 0)
    return leftover <= 0


def apply_wash_complete_summary(
    maker_sum: list[dict],
    parent: str,
    *,
    survey_day: str = "2026-08-23",
) -> list[dict]:
    """完走時サマリ更新（§6.4-1）。"""
    out = deepcopy(maker_sum)
    for row in out:
        if str(row.get("メーカー") or "").strip() != parent:
            continue
        row["洗い状態"] = PR_WASH_ST_DONE
        row["洗い進捗"] = "done"
        row["最終調査日"] = survey_day
        ck = row.get("調査CK")
        if ck is True or str(ck or "").upper() in ("TRUE", "1"):
            row["調査CK"] = False
        break
    return out


def run_cycle_c4_chain(
    *,
    parent: str,
    wash_complete: bool,
    maker_sum: list[dict],
    d1: dict | None = None,
    d2: dict | None = None,
    sellers: list[dict] | None = None,
    wash_split: list[dict] | None = None,
    wash_queue: list[dict] | None = None,
    catalog_estimate: dict[str, int] | None = None,
    c4_enabled: bool = True,
    cap: int = 1,
    now_ms: int | None = None,
    survey_day: str = "2026-08-23",
) -> dict[str, Any]:
    """完走後 C4 状態機械。

    Returns:
        maker_sum_after, ran_d, d_ok, arm_next, next_maker, chain_state, summary
    """
    parent = str(parent or "").strip()
    if not wash_complete:
        return {
            "maker_sum_after": deepcopy(maker_sum),
            "ran_d": False,
            "d_ok": False,
            "arm_next": False,
            "next_maker": None,
            "chain_state": "skipped_incomplete",
            "summary": "wash_incomplete parent=" + parent,
        }

    sum_after = apply_wash_complete_summary(maker_sum, parent, survey_day=survey_day)
    d_ok = d_chain_success(d1, d2)
    if not d_ok:
        reason = "d_fail"
        if d1 and not d_step_success(d1):
            reason = "d1_fail"
        elif d2 and not d_step_success(d2):
            reason = "d2_fail"
        elif d2 and int(d2.get("leftover") or 0) > 0:
            reason = "d2_leftover"
        return {
            "maker_sum_after": sum_after,
            "ran_d": True,
            "d_ok": False,
            "arm_next": False,
            "next_maker": None,
            "chain_state": "stopped",
            "summary": reason + " parent=" + parent,
        }

    if not c4_enabled:
        return {
            "maker_sum_after": sum_after,
            "ran_d": True,
            "d_ok": True,
            "arm_next": False,
            "next_maker": None,
            "chain_state": "idle_disabled",
            "summary": "d_ok c4_off parent=" + parent,
        }

    now_ms = now_ms if now_ms is not None else int(datetime(2026, 8, 23).timestamp() * 1000)
    picked = pick_cycle_maker_select(
        sum_after,
        sellers=sellers or [],
        wash_split=wash_split or [],
        wash_queue=wash_queue or [],
        catalog_estimate=catalog_estimate or {},
        cap=cap,
        now_ms=now_ms,
    )
    next_maker = picked[0] if picked else None
    if not next_maker:
        return {
            "maker_sum_after": sum_after,
            "ran_d": True,
            "d_ok": True,
            "arm_next": False,
            "next_maker": None,
            "chain_state": "idle_no_next",
            "summary": "d_ok pick=0 parent=" + parent,
        }

    return {
        "maker_sum_after": sum_after,
        "ran_d": True,
        "d_ok": True,
        "arm_next": True,
        "next_maker": next_maker,
        "chain_state": "armed",
        "summary": "d_ok arm=" + next_maker + " parent=" + parent,
    }
