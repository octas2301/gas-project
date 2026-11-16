# -*- coding: utf-8 -*-
"""
F2 適正在庫算出（単品）。モール・マスタへ書かない。出口はローカルTSVのみ。

例:
  python calc_proper_stock.py --limit 20
  python calc_proper_stock.py --stub-monthly 0
  python calc_proper_stock.py --monthly-csv path/to/stub.csv
"""
from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "out"
sys.path.insert(0, str(ROOT))

from formulas import (  # noqa: E402
    available_qty,
    parse_set_count,
    proper_qty,
    safety_stock,
)
from inventory_load import build_inventory_cache  # noqa: E402
from listing_read import LISTING_SHEET, iter_child_rows  # noqa: E402
from sheets_io import (  # noqa: E402
    inventory_spreadsheet_id,
    listing_spreadsheet_id,
    read_sheet_values,
    sheets_service_readonly,
)
from stock_recalc import get_current_stock, get_last_count_date, norm_key  # noqa: E402

JST = timezone(timedelta(hours=9))

OUT_FIELDS = [
    "run_id",
    "ran_at",
    "child_sku",
    "parent_sku",
    "set_count",
    "stock_key",
    "stock_qty",
    "monthly_sales",
    "monthly_sales_source",
    "safety_stock",
    "available",
    "proper_qty",
    "skip_reason",
    "note",
    "listing_row",
]


def load_monthly_stub(path: str | None, default: float) -> tuple[dict[str, float], str]:
    """stock_key → 月販。ファイル無ければ全キー default。"""
    if not path:
        return {}, "stub"
    p = Path(path)
    if not p.is_file():
        raise SystemExit("monthly csv not found: %s" % path)
    m: dict[str, float] = {}
    with p.open(encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        for row in r:
            k = norm_key(row.get("stock_key") or row.get("single_jan") or row.get("JAN") or "")
            if not k:
                continue
            try:
                m[k] = float(row.get("monthly_sales") or 0)
            except ValueError:
                m[k] = default
    return m, "stub"


def main() -> int:
    ap = argparse.ArgumentParser(description="Domain4 F2 proper stock calc (TSV only, no writes)")
    ap.add_argument("--limit", type=int, default=0, help="子SKUの最大件数（0=全件）")
    ap.add_argument("--stub-monthly", type=float, default=0.0, help="月販スタブ既定値")
    ap.add_argument("--monthly-csv", type=str, default="", help="stub月販CSV (stock_key,monthly_sales)")
    ap.add_argument("--dry-headers", action="store_true", help="出品ヘッダーだけ表示して終了")
    args = ap.parse_args()

    print("mode=read_only write=never phase=F2 monthly_sales_source=stub")
    svc = sheets_service_readonly()
    if not svc:
        print("FAIL: NO_CREDS")
        return 2

    inv_id = inventory_spreadsheet_id()
    listing_id = listing_spreadsheet_id()
    cache = build_inventory_cache(svc, inv_id)
    if cache.missing_sheets:
        print("FAIL: missing_sheets=%s" % ",".join(cache.missing_sheets))
        return 3
    last = get_last_count_date(cache)
    if last is None:
        print("FAIL: no_last_count_date")
        return 4

    listing_rows = read_sheet_values(svc, listing_id, LISTING_SHEET)
    children, meta = iter_child_rows(listing_rows)
    print(
        "listing header_row=%s set_col=%s jan_col=%s child_rows=%d inv_keys=%d last_count=%s"
        % (
            meta.get("header_row"),
            meta.get("set_col"),
            meta.get("jan_col"),
            len(children),
            len(cache.product_jans),
            last.isoformat(),
        )
    )
    if args.dry_headers:
        print("headers=%s" % ",".join(meta.get("headers") or [])[:500])
        return 0
    if meta.get("error"):
        print("FAIL: %s" % meta["error"])
        return 5

    monthly_map, monthly_source = load_monthly_stub(args.monthly_csv or None, args.stub_monthly)
    inv_keys = set(cache.product_jans)

    if args.limit and args.limit > 0:
        children = children[: args.limit]

    ran = datetime.now(JST)
    run_id = "f2_" + ran.strftime("%Y%m%d_%H%M%S")
    # stub が一目で分かるファイル名
    out_path = OUT / ("proper_stock_STUB_%s.tsv" % ran.strftime("%Y%m%d_%H%M%S"))
    OUT.mkdir(parents=True, exist_ok=True)

    # stock_key ごとの現在庫をキャッシュ
    stock_cache: dict[str, object] = {}

    out_rows: list[dict] = []
    counts = {"ok": 0, "skip": 0}

    for ch in children:
        note_parts: list[str] = []
        skip = ""
        set_n, set_err = parse_set_count(ch.set_count_raw)
        stock_key = ch.stock_key
        stock_qty_out = ""
        safety_out = ""
        avail_out = ""
        proper_out = ""
        monthly = monthly_map.get(stock_key, args.stub_monthly) if stock_key else args.stub_monthly

        if ch.is_mix_candidate:
            skip = "mix_deferred_p6"
            note_parts.append("mix_col=" + (ch.mix_hint or ""))
        elif set_err:
            skip = set_err
        elif not stock_key:
            skip = "no_jan"
        elif stock_key not in inv_keys:
            skip = "stock_key_not_in_inventory"
        else:
            if stock_key not in stock_cache:
                stock_cache[stock_key] = get_current_stock(stock_key, cache, last)
            br = stock_cache[stock_key]
            if br.error:  # type: ignore[attr-defined]
                skip = "stock_calc_error"
                note_parts.append(br.error)  # type: ignore[attr-defined]
            elif br.stock_qty is None:  # type: ignore[attr-defined]
                skip = "stock_calc_error"
            else:
                sq = float(br.stock_qty)  # type: ignore[attr-defined]
                if sq < 0:
                    note_parts.append("negative_stock")
                saf = safety_stock(monthly)
                avail = available_qty(sq, saf)
                pq = proper_qty(avail, set_n)  # type: ignore[arg-type]
                if pq is None:
                    skip = "set_count_invalid"
                else:
                    stock_qty_out = sq
                    safety_out = saf
                    avail_out = avail
                    proper_out = pq
                    counts["ok"] += 1

        if skip:
            counts["skip"] += 1

        out_rows.append(
            {
                "run_id": run_id,
                "ran_at": ran.isoformat(timespec="seconds"),
                "child_sku": ch.child_sku,
                "parent_sku": ch.parent_sku,
                "set_count": "" if set_n is None else set_n,
                "stock_key": stock_key,
                "stock_qty": stock_qty_out,
                "monthly_sales": monthly,
                "monthly_sales_source": monthly_source,
                "safety_stock": safety_out,
                "available": avail_out,
                "proper_qty": proper_out,
                "skip_reason": skip,
                "note": ";".join(note_parts),
                "listing_row": ch.row_num,
            }
        )

    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=OUT_FIELDS, delimiter="\t")
        w.writeheader()
        w.writerows(out_rows)

    # サマリ
    skip_hist: dict[str, int] = {}
    for r in out_rows:
        sr = r["skip_reason"] or "(ok)"
        skip_hist[str(sr)] = skip_hist.get(str(sr), 0) + 1

    print("out=%s rows=%d ok=%d skip=%d" % (out_path, len(out_rows), counts["ok"], counts["skip"]))
    print("skip_breakdown=%s" % skip_hist)
    print("G2: parent excluded / numeric set / proper_or_skip / stub / mix_deferred_p6 / no fake zero")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
