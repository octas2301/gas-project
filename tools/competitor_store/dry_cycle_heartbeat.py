# -*- coding: utf-8 -*-
"""心拍仮想: CYCLE §7 完成形。

  python tools/competitor_store/dry_cycle_heartbeat.py
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "cycle_heartbeat"
sys.path.insert(0, str(ROOT))

from cycle_heartbeat import (  # noqa: E402
    build_daily_mail_body,
    next_heartbeat_action,
    on_01_fire,
    pick_auto_ck_sellers,
    recompute_small_maker_hits,
    should_arm_01_resume,
)


def run_case(path: Path) -> tuple[bool, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    case_id = data.get("id") or path.stem
    kind = data.get("kind") or "next"
    desc = data.get("desc") or ""

    if kind == "next":
        got = next_heartbeat_action(**(data.get("input") or {}))
        exp = data.get("expect") or {}
        ok = got.get("action") == exp.get("action") and (
            exp.get("maker") is None or got.get("maker") == exp.get("maker")
        )
        detail_has = exp.get("detail_has")
        if ok and detail_has:
            ok = detail_has in str(got.get("detail") or "")
        msg = f"{'PASS' if ok else 'FAIL'} {case_id} {desc}"
        if not ok:
            msg += f"\n  expect={exp!r}\n  actual={got!r}"
        return ok, msg

    if kind == "01_arm":
        got = should_arm_01_resume(**(data.get("input") or {}))
        exp = bool(data.get("expect"))
        ok = got == exp
        return ok, f"{'PASS' if ok else 'FAIL'} {case_id} arm={got} expect={exp}"

    if kind == "01_fire":
        got = on_01_fire(**(data.get("input") or {}))
        exp = data.get("expect")
        ok = got == exp
        return ok, f"{'PASS' if ok else 'FAIL'} {case_id} fire={got} expect={exp}"

    if kind == "auto_ck":
        got = pick_auto_ck_sellers(
            data.get("sellers") or [],
            daily_cap=int(data.get("daily_cap") or 3),
            already_cked_today=int(data.get("already_cked_today") or 0),
        )
        exp = list(data.get("expect") or [])
        ok = got == exp
        msg = f"{'PASS' if ok else 'FAIL'} {case_id} {desc}"
        if not ok:
            msg += f"\n  expect={exp!r}\n  actual={got!r}"
        return ok, msg

    if kind == "small_hits":
        got = recompute_small_maker_hits(
            gate_asins=data.get("gate_asins") or [],
            maker_bands=data.get("maker_bands") or {},
        )
        exp = data.get("expect") or {}
        ok = got == exp
        msg = f"{'PASS' if ok else 'FAIL'} {case_id} {desc}"
        if not ok:
            msg += f"\n  expect={exp!r}\n  actual={got!r}"
        return ok, msg

    if kind == "mail":
        got = build_daily_mail_body(**(data.get("input") or {}))
        exp_sub = data.get("expect_contains") or []
        ok = all(s in got for s in exp_sub)
        if data.get("expect_not_contains"):
            ok = ok and all(s not in got for s in data["expect_not_contains"])
        msg = f"{'PASS' if ok else 'FAIL'} {case_id} {desc}"
        if not ok:
            msg += f"\n  body={got!r}"
        return ok, msg

    return False, f"FAIL {case_id} unknown kind={kind}"


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
