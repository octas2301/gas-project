# -*- coding: utf-8 -*-
"""洗いトリガー沈黙の仮想: 仮説つぶし。

  python tools/competitor_store/dry_cycle_trigger_watchdog.py
  python tools/competitor_store/dry_cycle_trigger_watchdog.py --case T-CTW-01
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "cycle_trigger_watchdog"
sys.path.insert(0, str(ROOT))

from cycle_trigger_watchdog import (  # noqa: E402
    INFLIGHT_STALE_MS,
    simulate_ckpt_overlap_missed,
    simulate_ckpt_schedule,
    simulate_eff_ckpt_skip,
    simulate_eff_full_cap,
    simulate_eff_overlap_takeover,
    simulate_fire,
    simulate_legacy_fire_then_die,
    simulate_overlap_during_inflight,
    simulate_overlap_stale_after_min,
    simulate_inflight_defer_missed,
    simulate_shop_busy_skip,
    simulate_shop_yield_clears_inflight,
    watchdog_tick,
)


def _get(data: dict, result: dict, key: str):
    if key.startswith("state."):
        return (result.get("state") or {}).get(key[6:])
    if key.startswith("fire."):
        return (result.get("fire") or {}).get(key[5:])
    if key.startswith("wd."):
        return (result.get("wd") or {}).get(key[3:])
    if key.startswith("overlap."):
        return (result.get("overlap") or {}).get(key[8:])
    if key.startswith("first."):
        return (result.get("first") or {}).get(key[6:])
    if key.startswith("second."):
        return (result.get("second") or {}).get(key[7:])
    if key == "log_has":
        return None
    return result.get(key)


def _log_has(result: dict, needle: str) -> bool:
    for row in result.get("logs") or []:
        blob = " ".join(str(row.get(k) or "") for k in ("step", "state", "detail"))
        if needle in blob:
            return True
    return False


def run_case(path: Path) -> tuple[bool, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    case_id = data.get("id") or path.stem
    desc = data.get("desc") or ""
    kind = data.get("kind") or "fire"
    now_ms = int(data.get("now_ms") or 1_000_000)

    if kind == "legacy_die":
        result = simulate_legacy_fire_then_die(
            now_ms=now_ms,
            candidates=int(data.get("candidates") or 8),
        )
        wd_at = now_ms + int(data.get("watchdog_after_ms") or 0)
        result["wd"] = watchdog_tick(result["state"], now_ms=wd_at, run_id="wd")
        result["logs"] = result["state"].get("logs") or []
    elif kind == "overlap":
        result = simulate_overlap_during_inflight(now_ms=now_ms)
    elif kind == "ckpt_no_shorten":
        result = simulate_ckpt_schedule(
            after_min=int(data.get("after_min") or 22),
            tag=str(data.get("tag") or "full_trigger_ckpt0"),
            ckpt_no_shorten=bool(data.get("ckpt_no_shorten", True)),
            now_ms=now_ms,
        )
    elif kind == "ckpt_overlap_missed":
        result = simulate_ckpt_overlap_missed(
            now_ms=now_ms,
            ckpt_no_shorten=bool(data.get("ckpt_no_shorten", True)),
            defer_inflight=bool(data.get("defer_inflight", True)),
        )
    elif kind == "overlap_stale_min":
        result = simulate_overlap_stale_after_min(
            inflight_age_ms=int(data.get("inflight_age_ms") or 6 * 60_000),
            stale_aware=bool(data.get("stale_aware", True)),
            eff_loop=bool(data.get("eff_loop", False)),
        )
    elif kind == "eff_ckpt_skip":
        result = simulate_eff_ckpt_skip(
            slice_n=int(data.get("slice_n") or 0),
            wait_min=int(data.get("wait_min") or 1),
        )
    elif kind == "eff_full_cap":
        result = simulate_eff_full_cap(
            after_min=int(data.get("after_min") or 22),
            live_left=int(data.get("live_left") or 300),
        )
    elif kind == "eff_overlap_takeover":
        result = simulate_eff_overlap_takeover(
            inflight_age_ms=int(data.get("inflight_age_ms") or 23 * 60_000),
            eff_loop=bool(data.get("eff_loop", True)),
        )
    elif kind == "shop_yield":
        result = simulate_shop_yield_clears_inflight(
            now_ms=now_ms,
            yield_arm_min=int(data.get("yield_arm_min") or 3),
        )
    elif kind == "shop_busy_skip":
        result = simulate_shop_busy_skip(
            busy=bool(data.get("busy", True)),
            ttl_ok=bool(data.get("ttl_ok", True)),
        )
    elif kind == "inflight_defer_missed":
        result = simulate_inflight_defer_missed(now_ms=now_ms)
    elif kind == "fire_then_watchdog":
        result = simulate_fire(
            die_at=data.get("die_at"),
            throw_at=data.get("throw_at"),
            candidates=int(data.get("candidates") or 8),
            pool_cached=bool(data.get("pool_cached", True)),
            heartbeat_on=bool(data.get("heartbeat_on", True)),
            now_ms=now_ms,
        )
        wd_at = now_ms + int(data.get("watchdog_after_ms") or INFLIGHT_STALE_MS)
        result["wd"] = watchdog_tick(result["state"], now_ms=wd_at, run_id="wd")
        result["logs"] = result["state"].get("logs") or []
    else:
        result = simulate_fire(
            die_at=data.get("die_at"),
            throw_at=data.get("throw_at"),
            overlap=bool(data.get("overlap")),
            candidates=int(data.get("candidates") or 8),
            pool_cached=bool(data.get("pool_cached", True)),
            heartbeat_on=bool(data.get("heartbeat_on", True)),
            now_ms=now_ms,
        )
        result["logs"] = result["state"].get("logs") or []

    expect = data.get("expect") or {}
    diffs: list[str] = []
    ok = True
    for key, exp in expect.items():
        if key == "log_has":
            for needle in exp or []:
                if not _log_has(result, str(needle)):
                    ok = False
                    diffs.append("  log_has missing " + repr(needle))
            continue
        if key == "log_not":
            for needle in exp or []:
                if _log_has(result, str(needle)):
                    ok = False
                    diffs.append("  log_not present " + repr(needle))
            continue
        act = _get(data, result, key)
        if exp != act:
            ok = False
            diffs.append(f"  {key}: expect={exp!r} actual={act!r}")
    msg = f"{'PASS' if ok else 'FAIL'} {case_id} {desc}"
    if diffs:
        msg += "\n" + "\n".join(diffs)
    return ok, msg


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--case")
    args = ap.parse_args()
    files = sorted(FIX_DIR.glob("*.json"))
    if args.case:
        key = args.case.upper()
        files = [f for f in files if key in f.stem.upper()]
    if not files:
        print("no fixtures", file=sys.stderr)
        return 1
    ok_n = 0
    for f in files:
        ok, msg = run_case(f)
        print(msg)
        if ok:
            ok_n += 1
    print("---")
    print(f"{ok_n}/{len(files)} PASS")
    return 0 if ok_n == len(files) else 1


if __name__ == "__main__":
    raise SystemExit(main())
