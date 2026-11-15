# -*- coding: utf-8 -*-
"""title0_skip と C4 needs/D1 recover の仮想（GAS PurchaseResearch.js 同期）。

正本: B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md §7.5（2026-08-28）
"""
from __future__ import annotations

from copy import deepcopy
from typing import Any

from cycle_heartbeat import next_heartbeat_action

WASH_HOLD = "保留"
WASH_DONE = "完了"
WASH_KEEPA = "Keepa中"
WASH_ENUM = "列挙中"
TITLE0_SKIP_AFTER = 3
PROP_NEEDS = "PR_CYCLE_C4_NEEDS_D1"
PROP_D1_DONE = "PR_CYCLE_C4_D1_DONE"
PROP_TITLE0 = "PR_WASH_TITLE0_FAIL_JSON"
PROP_TITLE0_SKIP_SET = "PR_WASH_TITLE0_SKIP_SET"


def _pipe_get(props: dict[str, str], key: str) -> list[str]:
    raw = str(props.get(key) or "").strip()
    if not raw:
        return []
    return [p.strip() for p in raw.split("|") if p.strip()]


def _pipe_add(props: dict[str, str], key: str, parent: str) -> None:
    parent = str(parent or "").strip()
    if not parent:
        return
    lst = _pipe_get(props, key)
    if parent in lst:
        return
    lst.append(parent)
    props[key] = "|".join(lst)


def _pipe_remove(props: dict[str, str], key: str, parent: str) -> None:
    parent = str(parent or "").strip()
    if not parent:
        return
    lst = [p for p in _pipe_get(props, key) if p != parent]
    if lst:
        props[key] = "|".join(lst)
    else:
        props.pop(key, None)


def title0_map_get(props: dict[str, str]) -> dict[str, int]:
    import json

    raw = props.get(PROP_TITLE0, "")
    if not raw:
        return {}
    try:
        data = json.loads(raw)
        return {str(k): int(v) for k, v in (data or {}).items()}
    except (TypeError, ValueError, json.JSONDecodeError):
        return {}


def title0_map_save(props: dict[str, str], mp: dict[str, int]) -> None:
    import json

    props[PROP_TITLE0] = json.dumps(mp or {})


def title0_skip_set_has(props: dict[str, str], parent: str) -> bool:
    return str(parent or "").strip() in _pipe_get(props, PROP_TITLE0_SKIP_SET)


def title0_skip_set_add(props: dict[str, str], parent: str) -> None:
    _pipe_add(props, PROP_TITLE0_SKIP_SET, parent)


def title0_bump(props: dict[str, str], parent: str) -> int:
    parent = str(parent or "").strip()
    if title0_skip_set_has(props, parent):
        return TITLE0_SKIP_AFTER
    mp = title0_map_get(props)
    mp[parent] = int(mp.get(parent, 0)) + 1
    title0_map_save(props, mp)
    return mp[parent]


def title0_clear(props: dict[str, str], parent: str) -> None:
    parent = str(parent or "").strip()
    mp = title0_map_get(props)
    if parent in mp:
        del mp[parent]
        title0_map_save(props, mp)


def cand_gate_pass_count(parent: str, candidates: list[dict]) -> int:
    parent = str(parent or "").strip()
    n = 0
    for row in candidates:
        if str(row.get("メーカー") or "").strip() != parent:
            continue
        if str(row.get("門結果") or "").strip() != "通過":
            continue
        n += 1
    return n


def is_title0_blocked(row: dict, *, props: dict[str, str] | None = None) -> bool:
    parent = str(row.get("メーカー") or "").strip()
    if props and title0_skip_set_has(props, parent):
        return True
    st = str(row.get("洗い状態") or "").strip()
    prog = str(row.get("洗い進捗") or "")
    return st == WASH_HOLD and "title0_skip" in prog


def clear_false_complete(
    maker_sum: list[dict],
    props: dict[str, str],
    parent: str,
    *,
    queue_n: int = 0,
) -> None:
    _pipe_remove(props, PROP_NEEDS, parent)
    for row in maker_sum:
        if str(row.get("メーカー") or "").strip() != parent:
            continue
        if queue_n > 0:
            row["洗い状態"] = WASH_KEEPA
            row["洗い進捗"] = f"c4_no_candidates queue={queue_n}"
        else:
            row["洗い状態"] = WASH_ENUM
            row["洗い進捗"] = "c4_no_candidates rewash"
        break


def set_maker_status(
    maker_sum: list[dict],
    parent: str,
    status: str,
    progress: str = "",
    *,
    mark_c4_needs: bool = True,
    props: dict[str, str] | None = None,
    c4_enabled: bool = True,
    candidates: list[dict] | None = None,
) -> None:
    parent = str(parent or "").strip()
    if status == WASH_DONE and candidates is not None:
        if cand_gate_pass_count(parent, candidates) <= 0:
            status = WASH_KEEPA
            progress = "done_no_candidates"
    for row in maker_sum:
        if str(row.get("メーカー") or "").strip() != parent:
            continue
        row["洗い状態"] = status
        if progress != "":
            row["洗い進捗"] = progress
        break
    if (
        status == WASH_DONE
        and mark_c4_needs
        and props is not None
        and c4_enabled
        and (candidates is None or cand_gate_pass_count(parent, candidates) > 0)
    ):
        mark_c4_needs_d1(props, parent, candidates or [])


def mark_c4_needs_d1(props: dict[str, str], parent: str, candidates: list[dict]) -> None:
    parent = str(parent or "").strip()
    if not parent:
        return
    if parent in _pipe_get(props, PROP_D1_DONE):
        return
    if cand_gate_pass_count(parent, candidates) <= 0:
        return
    _pipe_add(props, PROP_NEEDS, parent)


def mark_c4_d1_done(props: dict[str, str], parent: str) -> None:
    parent = str(parent or "").strip()
    if not parent:
        return
    _pipe_remove(props, PROP_NEEDS, parent)
    _pipe_add(props, PROP_D1_DONE, parent)


def list_needs_d1(
    props: dict[str, str],
    maker_sum: list[dict],
    candidates: list[dict],
) -> list[str]:
    merged: dict[str, bool] = {}
    for p in _pipe_get(props, PROP_NEEDS):
        if cand_gate_pass_count(p, candidates) > 0:
            merged[p] = True
        else:
            clear_false_complete(maker_sum, props, p)
    for row in maker_sum:
        if str(row.get("洗い状態") or "").strip() != WASH_DONE:
            continue
        p = str(row.get("メーカー") or "").strip()
        if not p:
            continue
        if cand_gate_pass_count(p, candidates) <= 0:
            clear_false_complete(maker_sum, props, p)
            continue
        merged[p] = True
    d1done = set(_pipe_get(props, PROP_D1_DONE))
    return [p for p in merged if p not in d1done]


def arm_c4_recover(
    props: dict[str, str],
    maker_sum: list[dict],
    candidates: list[dict],
    *,
    c4_enabled: bool = True,
    c4_pending: str = "",
    c4_trigger_n: int = 0,
) -> str:
    if not c4_enabled:
        return ""
    if str(c4_pending or "").strip():
        return ""
    if int(c4_trigger_n or 0) > 0:
        return ""
    needs = list_needs_d1(props, maker_sum, candidates)
    with_cand = [p for p in needs if cand_gate_pass_count(p, candidates) > 0]
    if not with_cand:
        return ""
    last = str(props.get("PR_CYCLE_C4_LAST_PARENT") or "").strip()
    parent = last if last in with_cand else with_cand[0]
    row = next((r for r in maker_sum if str(r.get("メーカー") or "").strip() == parent), None)
    if not row or str(row.get("洗い状態") or "").strip() != WASH_DONE:
        return ""
    return parent


def handle_wash_title_zero(
    props: dict[str, str],
    maker_sum: list[dict],
    parent: str,
    *,
    arm_trigger: bool = True,
    heartbeat_enabled: bool = True,
    skip_after: int = TITLE0_SKIP_AFTER,
) -> dict[str, Any]:
    n = title0_bump(props, parent)
    if title0_skip_set_has(props, parent):
        set_maker_status(maker_sum, parent, WASH_HOLD, "title0_skip permanent", mark_c4_needs=False)
        return {
            "skipped": True,
            "count": n,
            "detail": f"title0_skip permanent parent={parent}",
            "wash_armed": False,
            "heartbeat_armed": False,
            "status": WASH_HOLD,
        }
    if n < skip_after:
        armed = False
        if arm_trigger:
            props["PR_WASH_FULL_TRIGGER"] = "true"
            props["_wash_arm_parent"] = parent
            armed = True
        return {
            "skipped": False,
            "count": n,
            "detail": f"title0_retry n={n} parent={parent}",
            "wash_armed": armed,
            "heartbeat_armed": False,
            "status": None,
        }
    title0_skip_set_add(props, parent)
    mp = title0_map_get(props)
    mp[parent] = TITLE0_SKIP_AFTER
    title0_map_save(props, mp)
    set_maker_status(
        maker_sum,
        parent,
        WASH_HOLD,
        f"title0_skip n={n}",
        mark_c4_needs=False,
    )
    props["PR_WASH_FULL_TRIGGER"] = "false"
    props.pop("_wash_arm_parent", None)
    detail = f"title0_skip parent={parent} n={n}"
    hb = heartbeat_enabled
    if hb:
        detail += " heartbeat_armed afterMin=1"
    return {
        "skipped": True,
        "count": n,
        "detail": detail,
        "wash_armed": False,
        "heartbeat_armed": hb,
        "status": WASH_HOLD,
    }


def maker_sum_for_pick(
    maker_sum: list[dict],
    props: dict[str, str],
) -> list[dict]:
    out: list[dict] = []
    for row in maker_sum:
        if is_title0_blocked(row, props=props):
            continue
        out.append(row)
    return out


def heartbeat_pick(
    maker_sum: list[dict],
    props: dict[str, str],
    *,
    wash_queue: list[dict] | None = None,
    wash_split: list[dict] | None = None,
    catalog_estimate: dict[str, int] | None = None,
) -> dict[str, Any]:
    ms = maker_sum_for_pick(maker_sum, props)
    q = wash_queue or []
    filtered_q = [
        r
        for r in q
        if not is_title0_blocked({"メーカー": r.get("メーカー")}, props=props)
    ]
    return next_heartbeat_action(
        maker_sum=ms,
        wash_queue=filtered_q,
        wash_split=wash_split or [],
        catalog_estimate=catalog_estimate or {},
    )


def simulate_d1_append(
    *,
    candidates: list[dict],
    list_asins: set[str],
    maker: str,
) -> dict[str, int]:
    pass_n = 0
    exist = 0
    append = 0
    for row in candidates:
        if str(row.get("メーカー") or "").strip() != maker:
            continue
        if str(row.get("門結果") or "").strip() != "通過":
            continue
        pass_n += 1
        asin = str(row.get("ASIN") or "").strip().upper()
        if not asin:
            continue
        if asin in list_asins:
            exist += 1
        else:
            append += 1
            list_asins.add(asin)
    return {"pass": pass_n, "exist": exist, "append": append, "new": append}


def run_script(steps: list[dict], start: dict[str, Any] | None = None) -> dict[str, Any]:
    state: dict[str, Any] = deepcopy(start or {})
    props: dict[str, str] = dict(state.get("props") or {})
    maker_sum: list[dict] = deepcopy(state.get("maker_sum") or [])
    candidates: list[dict] = deepcopy(state.get("candidates") or [])
    list_asins: set[str] = set(str(a).upper() for a in (state.get("list_asins") or []))
    logs: list[str] = list(state.get("logs") or [])

    for step in steps:
        op = step.get("op") or ""
        if op == "title0_fail":
            r = handle_wash_title_zero(
                props,
                maker_sum,
                step["parent"],
                arm_trigger=bool(step.get("arm_trigger", True)),
                heartbeat_enabled=bool(step.get("heartbeat_enabled", True)),
            )
            logs.append(r["detail"])
            state["last_title0"] = r
        elif op == "title0_clear":
            title0_clear(props, step["parent"])
            logs.append(f"title0_clear {step['parent']}")
        elif op == "mark_done":
            set_maker_status(
                maker_sum,
                step["parent"],
                WASH_DONE,
                step.get("progress") or "done",
                props=props,
                c4_enabled=bool(step.get("c4_enabled", True)),
                candidates=candidates,
            )
            logs.append(f"mark_done {step['parent']}")
        elif op == "mark_d1_done":
            mark_c4_d1_done(props, step["parent"])
            logs.append(f"d1_done {step['parent']}")
        elif op == "arm_c4":
            p = arm_c4_recover(
                props,
                maker_sum,
                candidates,
                c4_enabled=bool(step.get("c4_enabled", True)),
                c4_pending=str(state.get("c4_pending") or ""),
                c4_trigger_n=int(state.get("c4_n") or 0),
            )
            state["c4_arm"] = p
            logs.append(f"c4_arm {p or 'none'}")
        elif op == "d1":
            p = step["parent"]
            r = simulate_d1_append(candidates=candidates, list_asins=list_asins, maker=p)
            state["last_d1"] = r
            logs.append(f"d1 {p} append={r['append']}")
        elif op == "heartbeat":
            r = heartbeat_pick(
                maker_sum,
                props,
                wash_queue=step.get("wash_queue") or state.get("wash_queue") or [],
            )
            state["last_hb"] = r
            logs.append(f"hb action={r.get('action')} maker={r.get('maker')}")
        elif op == "set_last_parent":
            props["PR_CYCLE_C4_LAST_PARENT"] = step["parent"]
        else:
            logs.append(f"unknown_op {op}")

    state["props"] = props
    state["maker_sum"] = maker_sum
    state["candidates"] = candidates
    state["list_asins"] = sorted(list_asins)
    state["needs_d1"] = list_needs_d1(props, maker_sum, candidates)
    state["title0_map"] = title0_map_get(props)
    state["logs"] = logs
    return state
