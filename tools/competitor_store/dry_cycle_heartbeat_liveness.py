# -*- coding: utf-8 -*-
"""心拍が止まらない全体フローの仮想。

  python tools/competitor_store/dry_cycle_heartbeat_liveness.py
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "cycle_heartbeat_liveness"
sys.path.insert(0, str(ROOT))

from cycle_heartbeat_liveness import run_script  # noqa: E402


def run_case(path: Path) -> tuple[bool, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    case_id = data.get("id") or path.stem
    desc = data.get("desc") or ""
    result = run_script(data.get("steps") or [], start=data.get("start") or {})
    expect = data.get("expect") or {}
    diffs: list[str] = []
    ok = True
    if "alive" in expect:
        act = bool((result.get("alive") or {}).get("alive"))
        if act != bool(expect["alive"]):
            ok = False
            diffs.append(f"  alive: expect={expect['alive']!r} actual={act!r} reason={(result.get('alive') or {}).get('reason')}")
    if "reason" in expect:
        act_r = (result.get("alive") or {}).get("reason")
        if act_r != expect["reason"]:
            ok = False
            diffs.append(f"  reason: expect={expect['reason']!r} actual={act_r!r}")
    st = result.get("state") or {}
    for key in ("wash_n", "c4_n", "resume01_n", "paused"):
        if key in expect:
            if st.get(key) != expect[key]:
                ok = False
                diffs.append(f"  {key}: expect={expect[key]!r} actual={st.get(key)!r}")
    if "log_has" in expect:
        blob = " ".join(result.get("logs") or [])
        for needle in expect["log_has"] or []:
            if str(needle) not in blob:
                ok = False
                diffs.append(f"  log_has missing {needle!r} in {blob!r}")
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
