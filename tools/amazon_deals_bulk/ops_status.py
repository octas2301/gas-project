# -*- coding: utf-8 -*-
"""
運用ダッシュボード（書き込みなし）。

- 早期申請割引の締切カウントダウン
- 今後のB／数量改定ウィンドウ
- A行の有無
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, List

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from paths import load_config  # noqa: E402
from schedule_class import (  # noqa: E402
    EARLY_FEE_DEADLINES_DEFAULT,
    parse_ymd,
)
from sheet_schema import LANE_A, LANE_B, MASTER_SHEET, SALE_SHEET  # noqa: E402
from sheets_io import read_sheet_rows, sheets_service  # noqa: E402


def _truthy(v: Any) -> bool:
    return str(v or "").strip().upper() in ("TRUE", "はい", "YES", "Y", "1", "○")


def main() -> int:
    ap = argparse.ArgumentParser(description="販促タイムセール運用状況")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--today", type=str, default=None)
    args = ap.parse_args()
    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))
    today = parse_ymd(args.today) if args.today else date.today()
    assert today

    dlmap = cfg.get("early_fee_deadlines") or EARLY_FEE_DEADLINES_DEFAULT
    yen = cfg.get("early_fee_discount_yen") or 500

    print(f"=== 今日 {today.isoformat()} ===")
    print(f"--- 早期申請割引（{yen}円）締切 ---")
    for name, ymd in dlmap.items():
        d = parse_ymd(str(ymd))
        if not d:
            continue
        left = (d - today).days
        if left < 0:
            print(f"  {name}: {ymd} 期限切れ（{abs(left)}日前）")
        elif left == 0:
            print(f"  {name}: {ymd} ★本日締切")
        elif left <= 7:
            print(f"  {name}: {ymd} ★残り{left}日")
        else:
            print(f"  {name}: {ymd} 残り{left}日")

    svc = sheets_service(write=False)
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    try:
        _h, sales = read_sheet_rows(svc, sid, SALE_SHEET)
    except Exception as e:
        print(f"（シート未読込: {e}）")
        return 0

    print("--- B公式 ---")
    b_n = 0
    for r in sales:
        if str(r.get("レーン") or "").strip() != LANE_B:
            continue
        b_n += 1
        start = parse_ymd(r.get("開始日"))
        days = (start - today).days if start else None
        flag = ""
        if start and 0 <= days <= 21:
            flag = " [数量改定ウィンドウ]"
        print(
            f"  {r.get('ASIN')} {r.get('開始日')}..{r.get('終了日')} "
            f"qty={r.get('販売商品数_確定')} 状態={r.get('状態')} 提出={r.get('提出対象')}{flag}"
        )
    if not b_n:
        print("  (なし)")

    print("--- A期間値下げ ---")
    a_n = 0
    for r in sales:
        if str(r.get("レーン") or "").strip() != LANE_A:
            continue
        a_n += 1
        print(
            f"  {r.get('ASIN')} {r.get('開始日')}..{r.get('終了日')} "
            f"price={r.get('セール価格')} 承認={r.get('承認済')} 状態={r.get('状態')} "
            f"A実施={r.get('A実施') or '（空）'} log={r.get('Aログ参照') or '-'}"
        )
    if not a_n:
        print("  (なし) → sync 後に合成Aが付く")

    print("--- マスタ A実施フラグ ---")
    try:
        sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
        _mh, masters = read_sheet_rows(svc, sid, MASTER_SHEET)
        shown = 0
        for r in masters:
            done = str(r.get("A実施") or "").strip()
            log = str(r.get("Aログ参照") or "").strip()
            if not done and not log:
                continue
            shown += 1
            print(
                f"  {r.get('SKU') or r.get('ASIN')} A実施={done or '（空）'} "
                f"期間={r.get('A期間') or '-'} 価格={r.get('A価格円') or '-'} log={log or '-'}"
            )
        if not shown:
            print("  (未記録) → lane_a_send 後 or python lane_a_send.py --backfill-a-logs")
    except Exception as e:
        print(f"  (読取スキップ: {e})")

    print("--- 数量確認メール窓（§9.7） ---")
    for r in sales:
        if str(r.get("レーン") or "").strip() != LANE_B:
            continue
        start = parse_ymd(r.get("開始日"))
        if not start or start < today:
            continue
        days = (start - today).days
        if days in (21, 14) or (14 <= days <= 21):
            print(
                f"  ★T-{days} {r.get('ASIN')} {start} qty={r.get('販売商品数_確定')} "
                f"{r.get('スケジュール')}"
            )
    print("  （既定: 開始21日前＝第1報／14日前＝最終。改定はSC画面編集）")
    print("  下書き: python mail_qty_confirm.py   送信: … --send（SMTP設定時）")
    print("  リンク先は広告スプシ（SCモバイルではタイムセール確認不可）")

    print("--- 新Sale検知 ---")
    print("  python detect_new_schedules.py          # 差分表示")
    print("  python detect_new_schedules.py --save   # 基準保存（月次DL後）")

    print("--- 次アクション ---")
    print("1. 月次バルクDL後: detect_new_schedules → 新規Saleを確認 → 必要なら参加")
    print("2. Smile接近: mail_qty_confirm / revise_b_qty → スプシ確認 → SC画面で数量編集")
    print("3. A: 別期間は新行＋lane_a_send（延長ヘルパなし）")
    print("4. Points: points_fetch --write → points_send --backup-before …")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
