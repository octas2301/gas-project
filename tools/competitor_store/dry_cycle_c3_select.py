# -*- coding: utf-8 -*-
"""C3 仮想: メーカー自動選定（§6.3）。

  python tools/competitor_store/dry_cycle_c3_select.py
  python tools/competitor_store/dry_cycle_c3_select.py --case T-CCS-01
  python tools/competitor_store/dry_cycle_c3_select.py --list
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "cycle_c3_select"
sys.path.insert(0, str(ROOT))

from cycle_c3_select import pick_cycle_maker_select, pick_enum_unknown  # noqa: E402


def _parse_now_ms(raw: str | None) -> int:
    if not raw:
        return int(datetime(2026, 8, 23).timestamp() * 1000)
    d = datetime.strptime(raw[:10], "%Y-%m-%d")
    return int(d.timestamp() * 1000)


def run_case(path: Path) -> tuple[bool, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    case_id = data.get("id") or path.stem
    desc = data.get("desc") or ""
    cap = int(data.get("cap") or data.get("expect", {}).get("cap") or 1)
    now_ms = _parse_now_ms(data.get("now"))
    mode = data.get("mode") or "wash"
    kwargs = dict(
        sellers=data.get("sellers") or [],
        wash_queue=data.get("wash_queue") or [],
        catalog_estimate=data.get("catalog_estimate") or {},
        cap=cap,
        now_ms=now_ms,
    )
    if mode == "enum_unknown":
        picked = pick_enum_unknown(data.get("maker_sum") or [], **kwargs)
    else:
        picked = pick_cycle_maker_select(
            data.get("maker_sum") or [],
            wash_split=data.get("wash_split") or [],
            **kwargs,
        )
    expect = list(data.get("expect", {}).get("picked") or [])
    ok = picked == expect
    msg = f"{'PASS' if ok else 'FAIL'} {case_id} {desc}"
    if not ok:
        msg += f"\n  expect={expect!r}\n  actual={picked!r}"
    return ok, msg


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", help="T-CCS-01 等（部分一致）")
    ap.add_argument("--list", action="store_true")
    args = ap.parse_args()
    files = sorted(FIX_DIR.glob("*.json"))
    if args.list:
        for f in files:
            print(f.stem)
        return 0
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
    total = len(files)
    print("---")
    print(f"{ok_n}/{total} PASS")
    return 0 if ok_n == total else 1


if __name__ == "__main__":
    raise SystemExit(main())
