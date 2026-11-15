# -*- coding: utf-8 -*-
"""title0_skip / C4 recover 仮想テスト。

  python tools/competitor_store/dry_cycle_title0_c4_recover.py
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "cycle_title0_c4_recover"
sys.path.insert(0, str(ROOT))

from cycle_title0_c4_recover import (  # noqa: E402
    TITLE0_SKIP_AFTER,
    arm_c4_recover,
    handle_wash_title_zero,
    heartbeat_pick,
    list_needs_d1,
    run_script,
    title0_bump,
    title0_clear,
    title0_map_get,
    title0_skip_set_has,
)


def _check_expect(result: dict, expect: dict, prefix: str = "") -> list[str]:
    diffs: list[str] = []
    for key, want in expect.items():
        if key == "log_has":
            blob = " ".join(result.get("logs") or [])
            for needle in want or []:
                if str(needle) not in blob:
                    diffs.append(f"{prefix}log_has missing {needle!r}")
            continue
        if key == "maker_status":
            ms = {str(r.get("メーカー") or ""): r for r in result.get("maker_sum") or []}
            for maker, st_want in (want or {}).items():
                row = ms.get(maker) or {}
                act_st = str(row.get("洗い状態") or "")
                act_prog = str(row.get("洗い進捗") or "")
                if isinstance(st_want, dict):
                    if st_want.get("洗い状態") and act_st != st_want["洗い状態"]:
                        diffs.append(f"{prefix}{maker} status expect={st_want['洗い状態']!r} actual={act_st!r}")
                    if st_want.get("洗い進捗") and act_prog != st_want["洗い進捗"]:
                        diffs.append(f"{prefix}{maker} prog expect={st_want['洗い進捗']!r} actual={act_prog!r}")
                elif act_st != st_want:
                    diffs.append(f"{prefix}{maker} status expect={st_want!r} actual={act_st!r}")
            continue
        act = result
        for part in key.split("."):
            if isinstance(act, dict):
                act = act.get(part)
            else:
                act = None
                break
        if act != want:
            diffs.append(f"{prefix}{key}: expect={want!r} actual={act!r}")
    return diffs


def run_case(path: Path) -> tuple[bool, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    case_id = data.get("id") or path.stem
    desc = data.get("desc") or ""
    mode = data.get("mode") or "script"
    candidates = list(data.get("candidates") or [])

    if mode == "unit_title0":
        props = dict(data.get("props") or {})
        maker_sum = list(data.get("maker_sum") or [])
        parent = data["parent"]
        result: dict = {}
        for _ in range(int(data.get("fail_times") or 1)):
            result = handle_wash_title_zero(props, maker_sum, parent, arm_trigger=bool(data.get("arm_trigger", True)))
        result["title0_map"] = title0_map_get(props)
        result["maker_sum"] = maker_sum
        result["logs"] = [result.get("detail", "")]
    elif mode == "unit_c4":
        props = dict(data.get("props") or {})
        maker_sum = list(data.get("maker_sum") or [])
        result = {
            "needs_d1": list_needs_d1(props, maker_sum, candidates),
            "c4_arm": arm_c4_recover(
                props,
                maker_sum,
                candidates,
                c4_enabled=bool(data.get("c4_enabled", True)),
                c4_pending=str(data.get("c4_pending") or ""),
                c4_trigger_n=int(data.get("c4_n") or 0),
            ),
        }
    elif mode == "unit_hb":
        props = dict(data.get("props") or {})
        result = {"current": heartbeat_pick(data.get("maker_sum") or [], props)}
    elif mode == "unit_counter_reset":
        props: dict[str, str] = {}
        parent = data["parent"]
        for _ in range(int(data.get("skip_rounds") or 1)):
            for _ in range(3):
                title0_bump(props, parent)
            handle_wash_title_zero(props, [], parent)
        result = {
            "after_repick_count": title0_bump(props, parent),
            "permanent": title0_skip_set_has(props, parent),
        }
    else:
        result = run_script(data.get("steps") or [], start=data.get("start") or {})

    diffs = _check_expect(result, data.get("expect") or {}, "  ")
    ok = not diffs
    msg = f"{'PASS' if ok else 'FAIL'} {case_id} {desc}"
    if diffs:
        msg += "\n" + "\n".join(diffs)
    if data.get("note"):
        msg += f"\n  note: {data['note']}"
    return ok, msg


def check_gas_cand_count_sync() -> tuple[bool, str]:
    """本番 GAS が D1 と同じ読取を使うこと（存在しない関数呼び出しを検出）。"""
    gas = ROOT.parent.parent / "purchase_research" / "PurchaseResearch.js"
    if not gas.is_file():
        return False, "FAIL T-GAS-01 PurchaseResearch.js なし"
    text = gas.read_text(encoding="utf-8")
    start = text.find("function prCandGatePassCountForMaker_")
    if start < 0:
        return False, "FAIL T-GAS-01 prCandGatePassCountForMaker_ なし"
    end = text.find("\nfunction ", start + 1)
    body = text[start:end] if end > start else text[start : start + 400]
    if "prCandRowsForListAppend_" in body:
        return False, "FAIL T-GAS-01 存在しない prCandRowsForListAppend_ を呼んでいる"
    if "prLoadCandPassByAsin_" not in body:
        return False, "FAIL T-GAS-01 prLoadCandPassByAsin_ を呼んでいない"
    return True, "PASS T-GAS-01 門通過件数は D1 と同じ prLoadCandPassByAsin_"


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
        print("---")
        if ok:
            ok_n += 1
    gas_ok, gas_msg = check_gas_cand_count_sync()
    print(gas_msg)
    print("---")
    total = len(files) + 1
    if gas_ok:
        ok_n += 1
    print(f"{ok_n}/{total} PASS")
    return 0 if ok_n == total else 1


if __name__ == "__main__":
    raise SystemExit(main())
