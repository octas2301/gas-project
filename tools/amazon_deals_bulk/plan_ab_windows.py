# -*- coding: utf-8 -*-
"""
B固定後のA期間案を表示（書き込みなし）。

例:
  python plan_ab_windows.py
  python plan_ab_windows.py --gap 5 --today 2026-08-11
  python plan_ab_windows.py --gap 0 --horizon 90
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from paths import folder_path, latest_xlsx, load_config  # noqa: E402
from schedule_class import parse_ymd, pick_schedules_split  # noqa: E402
from template_parse import (  # noqa: E402
    collect_schedule_catalog,
    collect_schedules,
    merge_schedules,
    read_template_rows,
)


def main() -> int:
    ap = argparse.ArgumentParser(description="B先行→A空き期間の案出し（dry）")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--gap", type=int, default=None, help="A↔B空き日数（未指定時はconfig）")
    ap.add_argument("--horizon", type=int, default=None, help="B取り込み水平線日数")
    ap.add_argument("--today", type=str, default=None, help="YYYY-MM-DD（検証用）")
    ap.add_argument("--limit-a", type=int, default=None)
    ap.add_argument("--limit-b", type=int, default=None)
    args = ap.parse_args()

    local = HERE / "config.local.json"
    cfg_path = args.config or (local if local.is_file() else HERE / "config.example.json")
    cfg = load_config(cfg_path)

    today = parse_ymd(args.today) if args.today else date.today()
    if today is None:
        raise SystemExit("--today は YYYY-MM-DD")
    gap = args.gap if args.gap is not None else int(cfg.get("ab_gap_days") or 0)
    horizon = args.horizon if args.horizon is not None else int(cfg.get("b_horizon_days") or 90)
    limit_a = args.limit_a if args.limit_a is not None else int(cfg.get("limit_a") or 2)
    limit_b = args.limit_b if args.limit_b is not None else int(cfg.get("limit_b") or 2)

    xlsx = latest_xlsx(folder_path(cfg, "02")) or latest_xlsx(folder_path(cfg, "01"))
    all_sched = []
    if xlsx and xlsx.is_file():
        t_rows, _, wb = read_template_rows(xlsx)
        all_sched = merge_schedules(collect_schedules(t_rows), collect_schedule_catalog(wb))

    if not all_sched:
        all_sched = [
            {
                "schedule": "Amazon Smileセール (2026-08-28 - 2026-09-03)",
                "start": "2026-08-28",
                "end": "2026-09-03",
            },
            {
                "schedule": "ブラックフライデー (2026-11-27 - 2026-12-07)",
                "start": "2026-11-27",
                "end": "2026-12-07",
            },
        ]
        print("(info) テンプレ無し→デモ候補を使用")

    b, a = pick_schedules_split(
        all_sched,
        limit_b=limit_b,
        limit_a=limit_a,
        today=today,
        ab_gap_days=gap,
        a_max_days=int(cfg.get("a_max_days") or 14),
        b_horizon_days=horizon,
        early_fee_deadlines=cfg.get("early_fee_deadlines"),
    )

    print(f"today={today.isoformat()} gap={gap} horizon={horizon} xlsx={xlsx}")
    print("--- B (公式・提出対象候補) ---")
    for x in b:
        print(f"  {x.get('schedule')} | {x.get('start')} .. {x.get('end')}")
    if not b:
        print("  (なし)")
    print("--- A (合成・API用・バルク外) ---")
    for x in a:
        print(f"  {x.get('schedule')} | gap={x.get('ab_gap_days')}")
    if not a:
        print("  (なし)")
    print("---")
    print("検証: BをSC登録確認後、AをAPIへ。NGなら --gap 5 / 10 / 14 で再案出し")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
