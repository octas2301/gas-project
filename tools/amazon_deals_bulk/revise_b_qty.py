# -*- coding: utf-8 -*-
"""
§10.9.1 接近時のB数量改定。

対象: レーンB・提出対象・開始日が today〜within_days 以内（または --all-upcoming）
V30/Q_fba を再取得し 販売商品数_確定 を再計算。--write でシート更新。

例:
  python revise_b_qty.py
  python revise_b_qty.py --within-days 21 --write
  python revise_b_qty.py --all-upcoming --write
"""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from paths import load_config  # noqa: E402
from qty_logic import compute_q_deal, deal_day_count  # noqa: E402
from schedule_class import parse_ymd  # noqa: E402
from sheet_schema import LANE_B, MASTER_SHEET, SALE_HEADERS, SALE_SHEET  # noqa: E402
from sheets_io import read_sheet_rows, sheets_service, write_headers_and_rows  # noqa: E402
from v30_source import resolve_v30_map  # noqa: E402

LOG = logging.getLogger("amazon_deals_bulk.revise_b_qty")


def _truthy(v: Any) -> bool:
    return str(v or "").strip().upper() in ("TRUE", "はい", "YES", "Y", "1", "○")


def _float(v: Any) -> Optional[float]:
    if v is None or str(v).strip() == "":
        return None
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None


def master_qfba_map(master: List[Dict[str, Any]]) -> Dict[str, Optional[float]]:
    out: Dict[str, Optional[float]] = {}
    for r in master:
        asin = str(r.get("ASIN") or "").strip().upper()
        if asin:
            out[asin] = _float(r.get("Q_fba"))
    return out


def select_b_rows(
    sales: List[Dict[str, Any]],
    *,
    today: date,
    within_days: int,
    all_upcoming: bool,
) -> List[Dict[str, Any]]:
    end = today + timedelta(days=max(0, within_days))
    out = []
    for r in sales:
        if str(r.get("レーン") or "").strip() != LANE_B:
            continue
        if str(r.get("提出対象") or "").strip() != "" and not _truthy(r.get("提出対象")):
            continue
        st = str(r.get("状態") or "").strip()
        if st in ("見送り", "終了", "失敗", "停止", "延期"):
            continue
        start = parse_ymd(r.get("開始日"))
        if not start:
            continue
        if start < today:
            continue
        if all_upcoming or start <= end:
            out.append(r)
    return out


def revise(
    cfg: dict,
    *,
    within_days: int,
    all_upcoming: bool,
    write: bool,
    today: Optional[date] = None,
) -> int:
    today = today or date.today()
    svc = sheets_service(write=True)
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    _mh, master = read_sheet_rows(svc, sid, MASTER_SHEET)
    _sh, sales = read_sheet_rows(svc, sid, SALE_SHEET)

    targets = select_b_rows(
        sales, today=today, within_days=within_days, all_upcoming=all_upcoming
    )
    LOG.info("改定候補B=%s within=%s all_upcoming=%s", len(targets), within_days, all_upcoming)
    if not targets:
        print("改定対象なし")
        return 0

    asins = [str(r.get("ASIN") or "").strip().upper() for r in targets]
    v30_map = resolve_v30_map(asins, master_rows=master, use_spapi=True)
    qfba = master_qfba_map(master)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_id = datetime.now().strftime("QTYREV_%Y%m%d_%H%M%S")
    by_id = {str(r.get("sale_id") or ""): r for r in sales}
    changes = []

    for r in targets:
        sid_row = str(r.get("sale_id") or "")
        asin = str(r.get("ASIN") or "").strip().upper()
        old_q = _float(r.get("販売商品数_確定"))
        d = deal_day_count(r.get("開始日"), r.get("終了日"))
        qd = compute_q_deal(
            v30=v30_map.get(asin),
            d_days=d,
            schedule=str(r.get("スケジュール") or ""),
            q_fba=qfba.get(asin),
        )
        new_q = qd["Q_deal"]
        row = by_id.get(sid_row) or r
        row["V30"] = qd["V30"] if qd["V30"] is not None else row.get("V30") or ""
        row["販売商品数_確定"] = new_q
        row["更新日時"] = now
        row["runId"] = run_id
        note = f"数量改定 {old_q}→{new_q} | {qd['note']}"
        prev = str(row.get("メッセージ") or "")
        row["メッセージ"] = note if not prev else f"{note} || {prev[:120]}"
        if qd.get("deferred"):
            row["状態"] = "延期"
            row["提出対象"] = "いいえ"
            row["有効"] = "FALSE"
        elif str(row.get("状態") or "") in ("", "下書き", "予定", "要確認"):
            row["状態"] = "数量改定済"
        changes.append(
            {
                "sale_id": sid_row,
                "ASIN": asin,
                "SKU": row.get("SKU"),
                "start": row.get("開始日"),
                "old": old_q,
                "new": new_q,
                "deferred": bool(qd.get("deferred")),
                "note": qd["note"],
            }
        )
        print(
            f"{'CHG' if old_q != new_q else 'SAME'} {asin} {row.get('開始日')} "
            f"{old_q}→{new_q} deferred={qd.get('deferred')}"
        )

    if not write:
        print(f"dry-run（{len(changes)}件）。シート反映は --write")
        return 0

    values = [[r.get(h, "") for h in SALE_HEADERS] for r in sales]
    write_headers_and_rows(svc, sid, SALE_SHEET, SALE_HEADERS, values, clear=True)
    LOG.info("シート更新完了 runId=%s changes=%s", run_id, len(changes))
    print(f"更新完了 runId={run_id}")
    print("次: 必要なら python build_submit_xlsx.py --write → SCで数量編集、または再UL")
    return 0


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="B公式の数量接近改定 §10.9.1")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--within-days", type=int, default=21, help="開始まで何日以内を対象（既定21）")
    ap.add_argument("--all-upcoming", action="store_true", help="未来のBすべて")
    ap.add_argument("--today", type=str, default=None)
    ap.add_argument("--write", action="store_true")
    args = ap.parse_args(argv)
    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))
    today = parse_ymd(args.today) if args.today else date.today()
    if today is None:
        raise SystemExit("--today は YYYY-MM-DD")
    return revise(
        cfg,
        within_days=args.within_days,
        all_upcoming=bool(args.all_upcoming),
        write=bool(args.write),
        today=today,
    )


if __name__ == "__main__":
    raise SystemExit(main())
