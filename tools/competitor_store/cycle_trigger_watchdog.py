# -*- coding: utf-8 -*-
"""洗いトリガー発火ライフサイクル（副作用なし）。

16:29 沈黙の仮説を仮想で潰す正本。GAS `prRunWashFullFromTrigger` の穴:

- 発火直後に expect/arm を消す → missed 番犬が無効
- 外側 try/finally なし → 選定中の例外・30分殺しで FAILED も再武装も無い
- 選定がメーカーごとに pool 再計算 → 重い
- 実行中ロックなし → 二重発火が空 arm で入る

正: inflight を expect より先に立てる。終了まで inflight を残す。
番犬は expect だけでなく inflight 期限切れを見る。二重発火は skip。
"""
from __future__ import annotations

from typing import Any

PHASE_FIRE = "fire"
PHASE_INFLIGHT = "inflight"
PHASE_EXPECT_CLEAR = "expect_clear"
PHASE_PICK = "pick"
PHASE_BRANCH = "branch"
PHASE_ARM = "arm"
PHASE_DONE = "done"
PHASE_FAILED = "failed"
PHASE_HARD_KILL = "hard_kill"
PHASE_OVERLAP_SKIP = "overlap_skip"

# 実行中とみなす上限（GAS 30分壁＋余裕）
INFLIGHT_STALE_MS = 35 * 60 * 1000
WATCHDOG_GRACE_MS = 3 * 60 * 1000
OVERLAP_WAIT_BASE_MIN = 8
EFF_ARM_FULL_LEFT = 280
OVERLAP_TAKEOVER_MIN = 8


def wait_shortened_allowed(tag: str, *, ckpt_no_shorten: bool = True) -> bool:
    """GAS prWashTriggerTagNoShorten_ + prWashCkptNoShorten_ 相当。"""
    if not ckpt_no_shorten:
        return True
    t = str(tag or "")
    if "_ckpt" in t:
        return False
    if t == "overlap_wait":
        return False
    return True


def apply_wait_shortened(
    mins: int,
    *,
    tag: str,
    calc_missing: bool = True,
    live_ok: bool = True,
    ckpt_no_shorten: bool = True,
) -> tuple[int, bool]:
    """GAS prScheduleWashFullTrigger_ の wait_shortened 分岐。"""
    if not wait_shortened_allowed(tag, ckpt_no_shorten=ckpt_no_shorten):
        return mins, False
    if live_ok and mins > 2 and calc_missing:
        return 2, True
    return mins, False


def apply_eff_loop_cap(
    mins: int,
    *,
    live_left: float | int | None,
    eff_loop: bool = True,
    full_left: int = EFF_ARM_FULL_LEFT,
) -> tuple[int, bool]:
    """GAS P-eff-1: 満タンなら afterMin≤2。"""
    if not eff_loop:
        return mins, False
    try:
        left = float(live_left)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return mins, False
    if left >= full_left and mins > 2:
        return 2, True
    return mins, False


def checkpoint_armed_mins(
    *,
    slice_n: int,
    wait_min: int = 1,
    gas_remain_min: int = 20,
    eff_loop: bool = True,
) -> int:
    """GAS prCheckpointWashFullTrigger_ 相当。eff_loop ON なら常に 0。"""
    if eff_loop:
        return 0
    after = max(2, int(gas_remain_min) + 2)
    from_tok = max(1, int(wait_min))
    return max(after, from_tok)


def overlap_wait_after_min(
    *,
    inflight_age_ms: int,
    stale_aware: bool = True,
    max_wait: int = 61,
    eff_loop: bool = False,
) -> int:
    """GAS prOverlapWaitAfterMin_ 相当。eff_loop ON は短再試行2分。"""
    if eff_loop:
        return 2
    base = OVERLAP_WAIT_BASE_MIN
    if not stale_aware:
        return base
    if inflight_age_ms <= 0:
        return base
    remain_ms = max(0, INFLIGHT_STALE_MS - inflight_age_ms)
    from_stale = max(base, (remain_ms + 59_999) // 60_000 + 1)
    if from_stale > max_wait:
        from_stale = max_wait
    return int(from_stale)


def overlap_should_takeover(
    *,
    inflight_age_ms: int,
    eff_loop: bool = True,
    thresh_min: int = OVERLAP_TAKEOVER_MIN,
) -> bool:
    """GAS P-eff-2: age≥thresh なら即 takeover。"""
    if not eff_loop:
        return False
    return int(inflight_age_ms) >= int(thresh_min) * 60 * 1000


def _log(state: dict[str, Any], run_id: str, step: str, st: str, detail: str) -> None:
    logs = state.setdefault("logs", [])
    logs.append({"runId": run_id, "step": step, "state": st, "detail": detail})


def schedule_trigger(
    state: dict[str, Any],
    *,
    now_ms: int,
    after_min: int = 1,
    tag: str = "hb_after_enum",
    parent: str = "",
    run_id: str = "",
) -> dict[str, Any]:
    """既存洗いトリガーを消して1本予約（prScheduleWashFullTrigger_）。"""
    mins = max(1, int(after_min))
    state["trigger_n"] = 1
    state["expect_ms"] = now_ms + mins * 60 * 1000
    state["arm_tag"] = tag
    state["arm_parent"] = parent
    state["arm_run_id"] = run_id
    state["prop_on"] = True
    _log(
        state,
        run_id,
        tag + " トリガー予約",
        "DONE",
        "trigger_armed afterMin=" + str(mins) + " tag=" + tag,
    )
    return state


def begin_fire(state: dict[str, Any], *, now_ms: int, run_id: str) -> dict[str, Any]:
    """発火先頭。inflight を立ててから fire ログ。expect はまだ消さない。"""
    out: dict[str, Any] = {
        "overlap_skip": False,
        "phase": PHASE_FIRE,
        "arm_tag_at_fire": str(state.get("arm_tag") or ""),
        "expect_at_fire": int(state.get("expect_ms") or 0),
        "trigger_n_at_fire": int(state.get("trigger_n") or 0),
        "inflight_owner": str(state.get("inflight_run_id") or ""),
        "overlap_rearmed": False,
        "overlap_after_min": 0,
    }
    owner = str(state.get("inflight_run_id") or "")
    since = int(state.get("inflight_ms") or 0)
    if owner and since > 0 and now_ms - since < INFLIGHT_STALE_MS:
        out["overlap_skip"] = True
        out["phase"] = PHASE_OVERLAP_SKIP
        _log(
            state,
            run_id,
            "full_trigger overlap",
            "DONE",
            "overlap_skip owner=" + owner + " ageMs=" + str(now_ms - since),
        )
        if bool(state.get("prop_on")):
            stale_aware = bool(state.get("overlap_stale_aware", True))
            after_min = overlap_wait_after_min(
                inflight_age_ms=now_ms - since,
                stale_aware=stale_aware,
            )
            schedule_trigger(
                state,
                now_ms=now_ms,
                after_min=after_min,
                tag="overlap_wait",
                run_id=run_id,
            )
            out["overlap_rearmed"] = True
            out["overlap_after_min"] = after_min
        return out
    state["inflight_run_id"] = run_id
    state["inflight_ms"] = now_ms
    state["inflight_phase"] = PHASE_INFLIGHT
    lag = ""
    exp = int(state.get("expect_ms") or 0)
    if exp > 0:
        lag = str(round((now_ms - exp) / 60000))
    _log(
        state,
        run_id,
        "full_trigger 洗い完走",
        "RUNNING",
        "fire lagMin="
        + str(lag)
        + " armTag="
        + str(state.get("arm_tag") or "")
        + " triggerN="
        + str(state.get("trigger_n") or 0)
        + " inflight="
        + run_id,
    )
    return out


def clear_expect_after_claim(state: dict[str, Any], *, run_id: str) -> None:
    """inflight 確保後にだけ expect を消す（missed 誤検知防止）。"""
    state["expect_ms"] = 0
    state["arm_tag"] = ""
    state["arm_parent"] = ""
    state["arm_run_id"] = ""
    state["last_clear"] = "fired:" + run_id
    state["inflight_phase"] = PHASE_EXPECT_CLEAR
    # one-shot 発火後、実行中は trigger が残ることがある（GAS）。消費は終了時。
    _log(state, run_id, "full_trigger phase", "RUNNING", "phase=expect_clear inflight=" + run_id)


def pick_pool_scans(candidates: int, cached: bool) -> int:
    """選定中の pool 再計算回数。cached=True なら 1 回。"""
    n = max(0, int(candidates))
    if n == 0:
        return 1 if cached else 1
    return 1 if cached else n


def finish_ok(
    state: dict[str, Any],
    *,
    run_id: str,
    branch: str,
    arm_after: bool = True,
    after_min: int = 1,
    now_ms: int = 0,
) -> dict[str, Any]:
    state["inflight_phase"] = PHASE_DONE
    state["inflight_run_id"] = ""
    state["inflight_ms"] = 0
    if arm_after:
        schedule_trigger(state, now_ms=now_ms, after_min=after_min, tag=branch, run_id=run_id)
    else:
        state["trigger_n"] = 0
        state["prop_on"] = False
    _log(state, run_id, "full_trigger 洗い完走", "DONE", "branch=" + branch)
    return {"ok": True, "branch": branch}


def finish_fail_rearm(
    state: dict[str, Any],
    *,
    run_id: str,
    reason: str,
    now_ms: int,
    after_min: int = 2,
) -> dict[str, Any]:
    state["inflight_phase"] = PHASE_FAILED
    state["inflight_run_id"] = ""
    state["inflight_ms"] = 0
    _log(state, run_id, "full_trigger 洗い完走", "FAILED", reason)
    schedule_trigger(
        state,
        now_ms=now_ms,
        after_min=after_min,
        tag="watchdog_rearm",
        run_id=run_id,
    )
    return {"ok": False, "rearmed": True, "reason": reason}


def hard_kill(state: dict[str, Any], *, run_id: str, phase: str) -> dict[str, Any]:
    """catch 未到達。inflight は残る（番犬の手がかり）。expect は既に空の想定。"""
    state["inflight_phase"] = PHASE_HARD_KILL
    # inflight_run_id / inflight_ms は残す
    _log(state, run_id, "full_trigger phase", "RUNNING", "hard_kill phase=" + phase)
    return {"ok": False, "hard_kill": True, "phase": phase}


def watchdog_tick(
    state: dict[str, Any],
    *,
    now_ms: int,
    run_id: str = "wd",
    defer_inflight: bool = True,
) -> dict[str, Any]:
    """沈黙検知。expect 期限切れ OR inflight 期限切れ。"""
    hint = ""
    inflight_id = str(state.get("inflight_run_id") or "")
    inflight_ms = int(state.get("inflight_ms") or 0)
    inflight_age = now_ms - inflight_ms if inflight_ms > 0 else 0
    expect_ms = int(state.get("expect_ms") or 0)
    trigger_n = int(state.get("trigger_n") or 0)
    prop_on = bool(state.get("prop_on"))

    if inflight_id and inflight_ms > 0 and inflight_age >= INFLIGHT_STALE_MS:
        hint = "inflight_stale"
    elif (
        defer_inflight
        and inflight_id
        and inflight_ms > 0
        and 0 < inflight_age < INFLIGHT_STALE_MS
        and expect_ms > 0
        and now_ms >= expect_ms + WATCHDOG_GRACE_MS
    ):
        return {"hint": "", "deferred_inflight": True, "rearmed": False}
    elif expect_ms > 0 and now_ms >= expect_ms + WATCHDOG_GRACE_MS:
        if trigger_n == 0 and prop_on:
            hint = "prop_on_but_no_trigger"
        elif trigger_n == 0 and not prop_on:
            hint = "no_trigger_prop_off"
        elif trigger_n > 0 and prop_on:
            hint = "trigger_still_queued_or_stuck"
        else:
            hint = "orphan_trigger_prop_off"
    # 旧穴: expect を先に消すと inflight 無しなら沈黙を見逃す
    elif (
        not inflight_id
        and expect_ms <= 0
        and trigger_n == 0
        and prop_on
        and state.get("legacy_expect_cleared_no_inflight")
    ):
        hint = "blind_silence"

    out = {"hint": hint, "rearmed": False}
    if hint in ("inflight_stale", "prop_on_but_no_trigger", "blind_silence"):
        schedule_trigger(
            state,
            now_ms=now_ms,
            after_min=1,
            tag="watchdog_" + hint,
            run_id=run_id,
        )
        state["inflight_run_id"] = ""
        state["inflight_ms"] = 0
        out["rearmed"] = True
        _log(state, run_id, "洗いトリガー診断", "DONE", "trigger_missed hint=" + hint)
    return out


def simulate_legacy_fire_then_die(
    *,
    now_ms: int = 1_000_000,
    candidates: int = 8,
) -> dict[str, Any]:
    """旧実装: fire ログの直後に expect を消し、inflight を立てないまま選定で死ぬ。"""
    state: dict[str, Any] = {}
    schedule_trigger(state, now_ms=now_ms - 60_000, after_min=1, tag="hb_after_enum", run_id="prior")
    run_id = "pr_legacy"
    _log(state, run_id, "full_trigger 洗い完走", "RUNNING", "fire armTag=hb_after_enum")
    state["expect_ms"] = 0
    state["arm_tag"] = ""
    state["arm_parent"] = ""
    state["arm_run_id"] = ""
    state["last_clear"] = "fired:" + run_id
    state["legacy_expect_cleared_no_inflight"] = True
    state["trigger_n"] = 0
    state["prop_on"] = True
    scans = pick_pool_scans(candidates, False)
    hard_kill(state, run_id=run_id, phase=PHASE_PICK)
    state["inflight_run_id"] = ""
    state["inflight_ms"] = 0
    return {
        "outcome": PHASE_HARD_KILL,
        "pool_scans": scans,
        "state": state,
        "logs": state.get("logs") or [],
        "legacy": True,
    }


def simulate_overlap_during_inflight(*, now_ms: int = 1_000_000) -> dict[str, Any]:
    """1本目が inflight のまま 2本目が発火。"""
    state: dict[str, Any] = {}
    schedule_trigger(state, now_ms=now_ms - 60_000, after_min=1, tag="hb_after_enum", run_id="prior")
    first = begin_fire(state, now_ms=now_ms, run_id="pr_162345")
    clear_expect_after_claim(state, run_id="pr_162345")
    second = begin_fire(state, now_ms=now_ms + 6 * 60 * 1000, run_id="pr_162942")
    return {
        "first": first,
        "second": second,
        "state": state,
        "logs": state.get("logs") or [],
        "second_arm_empty": not str(state.get("arm_tag") or ""),
    }


def simulate_fire(
    *,
    die_at: str | None = None,
    throw_at: str | None = None,
    overlap: bool = False,
    candidates: int = 8,
    pool_cached: bool = True,
    heartbeat_on: bool = True,
    now_ms: int = 1_000_000,
    prior: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """1回の発火を模擬。die_at / throw_at は phase 名。"""
    state = prior or {}
    if not state.get("expect_ms") and not state.get("inflight_run_id"):
        schedule_trigger(state, now_ms=now_ms - 60_000, after_min=1, tag="hb_after_enum", run_id="prior")
    run_id = "pr_fire"
    fire = begin_fire(state, now_ms=now_ms, run_id=run_id)
    scans = pick_pool_scans(candidates, pool_cached)
    result: dict[str, Any] = {
        "fire": fire,
        "pool_scans": scans,
        "logs": state.get("logs") or [],
        "state": state,
    }
    if fire.get("overlap_skip"):
        result["outcome"] = PHASE_OVERLAP_SKIP
        return result
    if die_at == PHASE_FIRE:
        hard_kill(state, run_id=run_id, phase=PHASE_FIRE)
        result["outcome"] = PHASE_HARD_KILL
        return result
    clear_expect_after_claim(state, run_id=run_id)
    if die_at == PHASE_EXPECT_CLEAR:
        hard_kill(state, run_id=run_id, phase=PHASE_EXPECT_CLEAR)
        result["outcome"] = PHASE_HARD_KILL
        return result
    _log(state, run_id, "full_trigger phase", "RUNNING", "phase=pick_begin candidates=" + str(candidates))
    if throw_at == PHASE_PICK:
        finish_fail_rearm(state, run_id=run_id, reason="pick_throw", now_ms=now_ms)
        result["outcome"] = PHASE_FAILED
        return result
    if die_at == PHASE_PICK:
        hard_kill(state, run_id=run_id, phase=PHASE_PICK)
        result["outcome"] = PHASE_HARD_KILL
        return result
    _log(
        state,
        run_id,
        "full_trigger phase",
        "DONE",
        "phase=pick_end scans=" + str(scans) + " elapsedMs=1",
    )
    if not heartbeat_on:
        finish_ok(state, run_id=run_id, branch="no_target", arm_after=False, now_ms=now_ms)
        result["outcome"] = PHASE_DONE
        return result
    if die_at == PHASE_BRANCH:
        hard_kill(state, run_id=run_id, phase=PHASE_BRANCH)
        result["outcome"] = PHASE_HARD_KILL
        return result
    finish_ok(state, run_id=run_id, branch="heartbeat", arm_after=True, now_ms=now_ms)
    result["outcome"] = PHASE_DONE
    if overlap:
        fire2 = begin_fire(state, now_ms=now_ms + 1000, run_id="pr_fire2")
        result["overlap_second"] = fire2
    return result


def simulate_ckpt_schedule(
    *,
    after_min: int = 22,
    tag: str = "full_trigger_ckpt0",
    ckpt_no_shorten: bool = True,
    now_ms: int = 1_000_000,
) -> dict[str, Any]:
    """ckpt 予約で wait_shortened が無効なら長待ちのまま。"""
    state: dict[str, Any] = {"prop_on": True}
    mins, shortened = apply_wait_shortened(
        after_min,
        tag=tag,
        ckpt_no_shorten=ckpt_no_shorten,
    )
    schedule_trigger(state, now_ms=now_ms, after_min=mins, tag=tag, run_id="pr_ckpt")
    return {
        "after_min": mins,
        "wait_shortened": shortened,
        "state": state,
        "logs": state.get("logs") or [],
    }


def simulate_ckpt_overlap_missed(
    *,
    now_ms: int = 1_000_000,
    ckpt_no_shorten: bool = True,
    defer_inflight: bool = True,
) -> dict[str, Any]:
    """ckpt 予約→overlap→watchdog。inflight 中は trigger_missed を出さない。"""
    state: dict[str, Any] = {"prop_on": True, "overlap_stale_aware": True}
    t0 = now_ms
    schedule_trigger(state, now_ms=t0 - 60_000, after_min=1, tag="hb", run_id="prior")
    begin_fire(state, now_ms=t0, run_id="pr_main")
    clear_expect_after_claim(state, run_id="pr_main")
    ck_mins, shortened = apply_wait_shortened(
        22,
        tag="full_trigger_ckpt0",
        ckpt_no_shorten=ckpt_no_shorten,
    )
    schedule_trigger(
        state,
        now_ms=t0 + 30_000,
        after_min=ck_mins,
        tag="full_trigger_ckpt0",
        run_id="pr_main",
    )
    fire_at = t0 + 30_000 + ck_mins * 60_000
    overlap = begin_fire(state, now_ms=fire_at, run_id="pr_overlap")
    wd = watchdog_tick(state, now_ms=fire_at + 5 * 60_000, defer_inflight=defer_inflight)
    missed = any(
        "trigger_missed" in str(r.get("detail") or "")
        for r in (state.get("logs") or [])
    )
    return {
        "ckpt_after_min": ck_mins,
        "wait_shortened": shortened,
        "overlap": overlap,
        "wd": wd,
        "trigger_missed_logged": missed,
        "state": state,
        "logs": state.get("logs") or [],
    }


def simulate_overlap_stale_after_min(
    *,
    inflight_age_ms: int = 6 * 60_000,
    stale_aware: bool = True,
    eff_loop: bool = False,
) -> dict[str, Any]:
    after = overlap_wait_after_min(
        inflight_age_ms=inflight_age_ms, stale_aware=stale_aware, eff_loop=eff_loop
    )
    return {"after_min": after, "stale_aware": stale_aware, "eff_loop": eff_loop}


def simulate_eff_ckpt_skip(*, slice_n: int = 0, wait_min: int = 1) -> dict[str, Any]:
    armed = checkpoint_armed_mins(slice_n=slice_n, wait_min=wait_min, eff_loop=True)
    legacy = checkpoint_armed_mins(slice_n=slice_n, wait_min=wait_min, eff_loop=False)
    return {"armed_min": armed, "legacy_min": legacy, "slice_n": slice_n}


def simulate_eff_full_cap(*, after_min: int = 22, live_left: int = 300) -> dict[str, Any]:
    mins, capped = apply_eff_loop_cap(after_min, live_left=live_left, eff_loop=True)
    return {"after_min": mins, "eff_capped": capped, "live_left": live_left}


def simulate_eff_overlap_takeover(
    *,
    inflight_age_ms: int = 23 * 60_000,
    eff_loop: bool = True,
) -> dict[str, Any]:
    take = overlap_should_takeover(inflight_age_ms=inflight_age_ms, eff_loop=eff_loop)
    rearm = 0 if take else overlap_wait_after_min(
        inflight_age_ms=inflight_age_ms, eff_loop=eff_loop
    )
    return {
        "takeover": take,
        "rearm_min": rearm,
        "age_min": int(inflight_age_ms // 60_000),
    }


def simulate_shop_yield_clears_inflight(
    *,
    now_ms: int = 1_000_000,
    yield_arm_min: int = 3,
) -> dict[str, Any]:
    """洗い→店譲り: inflight を捨て短予約（GAS prWashYieldToHeartbeat_）。"""
    state: dict[str, Any] = {"prop_on": True}
    schedule_trigger(state, now_ms=now_ms - 60_000, after_min=1, tag="prior", run_id="prior")
    begin_fire(state, now_ms=now_ms, run_id="pr_wash")
    clear_expect_after_claim(state, run_id="pr_wash")
    # yield
    state["inflight_run_id"] = ""
    state["inflight_ms"] = 0
    state["inflight_phase"] = ""
    state["shop_busy"] = True
    schedule_trigger(
        state, now_ms=now_ms + 1000, after_min=yield_arm_min, tag="hb_yield_deferred", run_id="pr_wash"
    )
    return {
        "inflight_cleared": not bool(state.get("inflight_run_id")),
        "shop_busy": bool(state.get("shop_busy")),
        "arm_tag": state.get("arm_tag"),
        "after_min": yield_arm_min,
        "state": state,
    }


def simulate_shop_busy_skip(*, busy: bool = True, ttl_ok: bool = True) -> dict[str, Any]:
    """店 busy 中の洗い発火は skip＋短再試行。"""
    if busy and ttl_ok:
        return {"skip": True, "rearm_min": 3, "reason": "shop_hb_busy"}
    return {"skip": False, "rearm_min": 0, "reason": ""}


def simulate_inflight_defer_missed(*, now_ms: int = 1_000_000) -> dict[str, Any]:
    """inflight 生存中に expect だけ期限切れ→trigger_missed にしない。"""
    state: dict[str, Any] = {"prop_on": True}
    t0 = now_ms
    schedule_trigger(state, now_ms=t0 - 60_000, after_min=1, tag="hb", run_id="prior")
    begin_fire(state, now_ms=t0, run_id="pr_main")
    clear_expect_after_claim(state, run_id="pr_main")
    schedule_trigger(
        state,
        now_ms=t0,
        after_min=2,
        tag="full_trigger_ckpt0",
        run_id="pr_main",
    )
    wd = watchdog_tick(state, now_ms=t0 + 5 * 60_000, defer_inflight=True)
    missed = any(
        "trigger_missed" in str(r.get("detail") or "")
        for r in (state.get("logs") or [])
    )
    return {
        "wd": wd,
        "trigger_missed_logged": missed,
        "state": state,
        "logs": state.get("logs") or [],
    }
