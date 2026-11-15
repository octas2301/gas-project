# -*- coding: utf-8 -*-
"""C4 仮想: 完走連鎖（§6.4）。

  python tools/competitor_store/dry_cycle_c4_chain.py
  python tools/competitor_store/dry_cycle_c4_chain.py --case T-CCC-01
  python tools/competitor_store/dry_cycle_c4_chain.py --list
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "cycle_c4_chain"
sys.path.insert(0, str(ROOT))

from cycle_c4_chain import run_cycle_c4_chain  # noqa: E402


def _parse_now_ms(raw: str | None) -> int:
    if not raw:
        return int(datetime(2026, 8, 23).timestamp() * 1000)
    d = datetime.strptime(raw[:10], "%Y-%m-%d")
    return int(d.timestamp() * 1000)


def _row_by_maker(maker_sum: list, name: str) -> dict | None:
    for row in maker_sum:
        if str(row.get("メーカー") or "").strip() == name:
            return row
    return None


def run_case(path: Path) -> tuple[bool, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    case_id = data.get("id") or path.stem
    desc = data.get("desc") or ""
    expect = data.get("expect") or {}
    now_ms = _parse_now_ms(data.get("now"))
    result = run_cycle_c4_chain(
        parent=data.get("parent") or "",
        wash_complete=bool(data.get("wash_complete")),
        maker_sum=data.get("maker_sum") or [],
        d1=data.get("d1"),
        d2=data.get("d2"),
        sellers=data.get("sellers") or [],
        wash_split=data.get("wash_split") or [],
        wash_queue=data.get("wash_queue") or [],
        catalog_estimate=data.get("catalog_estimate") or {},
        c4_enabled=data.get("c4_enabled", True),
        cap=int(data.get("cap") or 1),
        now_ms=now_ms,
        survey_day=str(data.get("survey_day") or "2026-08-23"),
    )
    ok = True
    diffs: list[str] = []
    for key in ("chain_state", "arm_next", "d_ok", "next_maker", "ran_d"):
        if key not in expect:
            continue
        exp = expect[key]
        act = result.get(key)
        if exp != act:
            ok = False
            diffs.append(f"  {key}: expect={exp!r} actual={act!r}")
    if "summary_contains" in expect:
        sub = str(expect["summary_contains"])
        if sub not in str(result.get("summary") or ""):
            ok = False
            diffs.append(f"  summary: missing {sub!r} in {result.get('summary')!r}")
    if "summary_row" in expect:
        want = expect["summary_row"]
        maker = str(want.get("メーカー") or data.get("parent") or "")
        row = _row_by_maker(result.get("maker_sum_after") or [], maker)
        if not row:
            ok = False
            diffs.append(f"  summary_row: maker {maker!r} not found")
        else:
            for k, v in want.items():
                if row.get(k) != v:
                    ok = False
                    diffs.append(f"  summary_row.{k}: expect={v!r} actual={row.get(k)!r}")
    msg = f"{'PASS' if ok else 'FAIL'} {case_id} {desc}"
    if diffs:
        msg += "\n" + "\n".join(diffs)
    return ok, msg


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", help="T-CCC-01 等（部分一致）")
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
