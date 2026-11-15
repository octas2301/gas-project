# -*- coding: utf-8 -*-
"""心拍が止まらないこと（CYCLE 全体）。

洗いトリガー発火だけでなく、心拍分岐・C4・pause・生存番犬まで含む。
生きている = 停止中なら 01:00 がある。稼働中なら 洗いorC4 の「これから来る」時計、
または実行中(inflight)かつ 20 分未満。
予定時刻が Keepa最大回復+α（既定 61+5=66分）以上過ぎている／inflight がその長さ腐っている
→ 沈黙扱い（リセット対象）。
"""
from __future__ import annotations

from typing import Any

INFLIGHT_NO_CLOCK_MS = 20 * 60 * 1000
LIVENESS_EVERY_MS = 15 * 60 * 1000
# ceil(300/5)+1 + alpha5
STALE_LIMIT_MS = 66 * 60 * 1000


def is_alive(state: dict[str, Any], *, now_ms: int) -> dict[str, Any]:
    if not state.get("heartbeat_enabled", True):
        return {"alive": True, "reason": "hb_off"}
    if state.get("paused"):
        ok = int(state.get("resume01_n") or 0) > 0
        return {"alive": ok, "reason": "paused_01" if ok else "paused_no_01"}
    wash = int(state.get("wash_n") or 0)
    c4 = int(state.get("c4_n") or 0)
    expect_ms = int(state.get("expect_ms") or 0)
    owner = str(state.get("inflight_run_id") or "")
    since = int(state.get("inflight_ms") or 0)
    age = now_ms - since if owner and since > 0 else 0

    if (wash > 0 or c4 > 0) and expect_ms > 0 and now_ms - expect_ms >= STALE_LIMIT_MS:
        return {"alive": False, "reason": "expect_overdue"}
    if owner and age >= STALE_LIMIT_MS:
        return {"alive": False, "reason": "inflight_overdue"}
    if wash > 0:
        return {"alive": True, "reason": "wash_trigger"}
    if c4 > 0:
        return {"alive": True, "reason": "c4_trigger"}
    if owner and since > 0 and now_ms - since < INFLIGHT_NO_CLOCK_MS:
        return {"alive": True, "reason": "inflight_fresh"}
    return {"alive": False, "reason": "silence"}


def _arm_wash(state: dict[str, Any], tag: str, *, now_ms: int = 0) -> None:
    state["wash_n"] = 1
    state["c4_n"] = 0
    state["last_arm"] = tag
    if now_ms:
        state["expect_ms"] = now_ms + 60 * 1000


def _arm_c4(state: dict[str, Any], *, now_ms: int = 0) -> None:
    state["c4_n"] = 1
    state["wash_n"] = 0
    state["last_arm"] = "c4"
    if now_ms:
        state["expect_ms"] = now_ms + 60 * 1000


def _arm_01(state: dict[str, Any]) -> None:
    state["resume01_n"] = 1
    state["wash_n"] = 0
    state["c4_n"] = 0
    state["last_arm"] = "resume01"


def _clear_inflight(state: dict[str, Any]) -> None:
    state["inflight_run_id"] = ""
    state["inflight_ms"] = 0


def _hard_reset(state: dict[str, Any]) -> None:
    state["wash_n"] = 0
    state["c4_n"] = 0
    state["expect_ms"] = 0
    _clear_inflight(state)


def apply_action(state: dict[str, Any], action: str, *, now_ms: int, legacy: bool = False) -> dict[str, Any]:
    """1アクション後の状態。legacy=True は改修前の抜け。"""
    logs = state.setdefault("logs", [])
    hb = state.get("heartbeat_enabled", True)

    def log(msg: str) -> None:
        logs.append(msg)

    if action == "pause_other":
        state["paused"] = True
        state["wash_n"] = 0
        state["c4_n"] = 0
        state["expect_ms"] = 0
        _clear_inflight(state)
        if not legacy:
            _arm_01(state)
        log("pause_other")
        return state

    if action == "manual_resume":
        state["paused"] = False
        state["resume01_n"] = 0
        _hard_reset(state)
        _arm_wash(state, "A1_resume", now_ms=now_ms)
        log("manual_resume")
        return state

    if action == "hb_paused":
        if not legacy:
            if int(state.get("resume01_n") or 0) < 1:
                _arm_01(state)
        log("hb_paused")
        return state

    if action == "hb_no_maker":
        _clear_inflight(state)
        if legacy:
            state["wash_n"] = 0
            state["c4_n"] = 0
            state["expect_ms"] = 0
        elif hb:
            _arm_wash(state, "hb_no_maker", now_ms=now_ms)
        log("hb_no_maker")
        return state

    if action in ("hb_enum", "hb_alight", "hb_auto_ck", "hb_c0", "hb_defer_xl", "hb_retry"):
        _clear_inflight(state)
        if hb:
            _arm_wash(state, action, now_ms=now_ms)
        log(action)
        return state

    if action == "hb_idle":
        _clear_inflight(state)
        if hb:
            _arm_wash(state, "hb_idle", now_ms=now_ms)
        log("hb_idle")
        return state

    if action == "hb_wash_incomplete":
        _clear_inflight(state)
        _arm_wash(state, "full_trigger", now_ms=now_ms)
        log("hb_wash_incomplete")
        return state

    if action == "hb_wash_complete":
        _clear_inflight(state)
        if hb:
            _arm_c4(state, now_ms=now_ms)
        else:
            state["wash_n"] = 0
            state["c4_n"] = 0
            state["expect_ms"] = 0
        log("hb_wash_complete")
        return state

    if action == "c4_ok_arm_next":
        _clear_inflight(state)
        _arm_wash(state, "C4_arm", now_ms=now_ms)
        log("c4_ok_arm_next")
        return state

    if action == "c4_pick0":
        _clear_inflight(state)
        if hb:
            _arm_wash(state, "C4_pick0_heartbeat", now_ms=now_ms)
        log("c4_pick0")
        return state

    if action == "c4_d_fail":
        _clear_inflight(state)
        if legacy:
            state["wash_n"] = 0
            state["c4_n"] = 0
            state["expect_ms"] = 0
        elif hb:
            _arm_wash(state, "C4_d_fail_heartbeat", now_ms=now_ms)
        log("c4_d_fail")
        return state

    if action == "fire_start":
        state["inflight_run_id"] = "pr_run"
        state["inflight_ms"] = now_ms
        state["wash_n"] = 0
        state["expect_ms"] = 0
        log("fire_start")
        return state

    if action == "fire_die_hard":
        state["inflight_run_id"] = state.get("inflight_run_id") or "pr_run"
        if not state.get("inflight_ms"):
            state["inflight_ms"] = now_ms
        state["wash_n"] = 0
        state["c4_n"] = 0
        state["expect_ms"] = 0
        log("fire_die_hard")
        return state

    if action == "arm_wash_past":
        # 腐った予約: washN=1 だが expect は古い
        state["wash_n"] = 1
        state["c4_n"] = 0
        state["expect_ms"] = int(state.get("expect_ms") or now_ms)
        log("arm_wash_past")
        return state

    if action == "liveness_tick":
        snap = is_alive(state, now_ms=now_ms)
        if snap["alive"]:
            log("liveness_ok " + snap["reason"])
            return state
        if state.get("paused"):
            _arm_01(state)
            log("liveness_rearm resume01")
            return state
        if hb:
            _hard_reset(state)
            _arm_wash(state, "liveness_reset", now_ms=now_ms)
            log("liveness_reset " + snap["reason"])
        return state

    log("unknown " + action)
    return state


def run_script(steps: list[dict[str, Any]], *, start: dict[str, Any] | None = None) -> dict[str, Any]:
    state: dict[str, Any] = dict(start or {})
    state.setdefault("heartbeat_enabled", True)
    state.setdefault("paused", False)
    state.setdefault("wash_n", 0)
    state.setdefault("c4_n", 0)
    state.setdefault("resume01_n", 0)
    state.setdefault("expect_ms", 0)
    state.setdefault("logs", [])
    last_now = 0
    for step in steps:
        now_ms = int(step.get("now_ms") or last_now)
        last_now = now_ms
        apply_action(
            state,
            str(step.get("action") or ""),
            now_ms=now_ms,
            legacy=bool(step.get("legacy")),
        )
    alive = is_alive(state, now_ms=last_now)
    return {"state": state, "alive": alive, "logs": state.get("logs") or []}
