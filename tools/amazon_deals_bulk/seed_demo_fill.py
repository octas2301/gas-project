# -*- coding: utf-8 -*-
"""今できる範囲で 期間値下げ／実行／マスタ原価U を埋める（デモ埋め）。"""
from __future__ import annotations

import hashlib
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

from paths import folder_path, latest_xlsx, load_config
from sheet_schema import EXEC_HEADERS, EXEC_SHEET, LANE_A_HEADERS, LANE_A_SHEET, MASTER_HEADERS, MASTER_SHEET
from sheets_io import read_sheet_rows, sheets_service, write_headers_and_rows
from template_parse import collect_schedules, read_template_rows

LOG = logging.getLogger("amazon_deals_bulk.seed")
HERE = Path(__file__).resolve().parent
PRODUCT_SS = "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"
PRODUCT_SHEET = "▼商品マスタ(人間作業用)"


def _sale_id(sku: str, schedule: str) -> str:
    return hashlib.sha1(f"{sku}|{schedule}".encode("utf-8")).hexdigest()[:12]


def lookup_cost_by_asin(svc) -> Dict[str, Dict[str, Any]]:
    """U列=セット卸値（税込み）。SKU列はマスタ表記が違うことがあるのでASINで突合。"""
    N = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=PRODUCT_SS, range=f"'{PRODUCT_SHEET}'!N8:N8000")
        .execute()
        .get("values")
        or []
    )
    U = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=PRODUCT_SS, range=f"'{PRODUCT_SHEET}'!U8:U8000")
        .execute()
        .get("values")
        or []
    )
    AK = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=PRODUCT_SS, range=f"'{PRODUCT_SHEET}'!AK8:AK8000")
        .execute()
        .get("values")
        or []
    )

    def g(arr, i):
        return str(arr[i][0]).strip() if i < len(arr) and arr[i] else ""

    out: Dict[str, Dict[str, Any]] = {}
    for i in range(len(N)):
        asin = g(N, i).upper()
        if not asin.startswith("B0"):
            continue
        out[asin] = {"原価U": g(U, i), "商品マスタSKU": g(AK, i)}
    return out


def pick_named_and_monthly(schedules: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """公式名付きを優先し、ついでに直近の月枠。カスタムは後回し。"""
    named = []
    monthly = []
    custom = []
    for s in schedules:
        name = s["schedule"]
        if "カスタム" in name:
            custom.append(s)
        elif any(k in name for k in ("スマイル", "Smile", "プライム", "ブラック")):
            named.append(s)
        else:
            monthly.append(s)
    picked = []
    for pool in (named, monthly, custom):
        for s in pool:
            if s not in picked:
                picked.append(s)
            if len(picked) >= 2:
                return picked
    return picked[:2]


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    cfg = load_config(HERE / "config.local.json")
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    svc = sheets_service(write=True)

    costs = lookup_cost_by_asin(svc)
    LOG.info("cost map size=%s", len(costs))

    _h, master = read_sheet_rows(svc, sid, MASTER_SHEET)
    src = latest_xlsx(folder_path(cfg, "02"))
    t_rows, _, _ = read_template_rows(src)
    schedules = collect_schedules(t_rows)
    picks = pick_named_and_monthly(schedules)
    LOG.info("schedule picks=%s", [p["schedule"] for p in picks])

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_id = datetime.now().strftime("SEED_%Y%m%d_%H%M%S")
    today = datetime.now().date()
    a_start = today + timedelta(days=1)
    a_end = a_start + timedelta(days=13)  # 14日イメージ

    # --- マスタ: 原価U埋め ---
    new_master: List[List[Any]] = []
    for r in master:
        asin = str(r.get("ASIN") or "").strip().upper()
        cost = costs.get(asin, {}).get("原価U", "")
        row = [r.get(h, "") for h in MASTER_HEADERS]
        # indices
        idx = {h: i for i, h in enumerate(MASTER_HEADERS)}
        if cost:
            row[idx["原価U"]] = cost
        if costs.get(asin, {}).get("商品マスタSKU") and not str(row[idx["メモ"]] or "").strip():
            row[idx["メモ"]] = f"商品マスタSKU={costs[asin]['商品マスタSKU']}"
        new_master.append(row)

    # --- 期間値下げ（レーンA）デモ ---
    lane_a: List[List[Any]] = []
    for r in master:
        sku = str(r.get("SKU") or "").strip()
        asin = str(r.get("ASIN") or "").strip().upper()
        if not sku and not asin:
            continue
        try:
            normal = float(str(r.get("出品者価格_SC") or "0").replace(",", ""))
        except ValueError:
            normal = 0.0
        try:
            cost = float(str(costs.get(asin, {}).get("原価U") or "0").replace(",", ""))
        except ValueError:
            cost = 0.0
        # API期間値下げの想定: 通常の約8%引き（公式SC上限とは別系統）
        sale = round(normal * 0.92) if normal else ""
        qty = 30
        try:
            qty = max(30, int(float(str(r.get("販売商品数_SC") or "30"))))
        except ValueError:
            pass
        profit = ""
        if sale and cost:
            profit = int((float(sale) - cost) * qty)
        lane_a.append(
            [
                sku,
                asin,
                "TRUE",
                "FALSE",  # 承認済（安定まで要承認）
                normal if normal else "",
                sale,
                a_start.isoformat(),
                a_end.isoformat(),
                qty,
                profit,
                "下書き",
                "",
                "P1a未接続のデモ埋め。承認後にAPI送信想定",
            ]
        )

    # --- 実行（B）: 月＋Smile優先。カスタムは見送り ---
    new_exec: List[List[Any]] = []
    for r in master:
        sku = str(r.get("SKU") or "").strip()
        asin = str(r.get("ASIN") or "").strip().upper()
        if not sku:
            continue
        try:
            price = float(str(r.get("タイムセール価格_SC") or "0").replace(",", ""))
        except ValueError:
            price = ""
        try:
            qty = float(str(r.get("販売商品数_SC") or "0").replace(",", ""))
        except ValueError:
            qty = ""
        for p in picks:
            is_custom = "カスタム" in p["schedule"]
            is_official_name = any(
                k in p["schedule"] for k in ("スマイル", "Smile", "プライム", "ブラック")
            )
            submit = "いいえ" if is_custom else "はい"
            note = "デモ: 次々回候補"
            if is_official_name:
                note = "公式名付きイベント枠"
            elif "月 (" in p["schedule"]:
                note = "おすすめ週次枠（月）"
            if is_custom:
                note = "カスタムは手動選定用・当面提出対象外"
            new_exec.append(
                [
                    _sale_id(sku, p["schedule"]),
                    sku,
                    asin,
                    "おすすめタイムセール",
                    p["schedule"],
                    p.get("start") or "",
                    p.get("end") or "",
                    price if price != "" else "",
                    int(qty) if qty != "" else "",
                    "予定",
                    submit,
                    now,
                    run_id,
                    note,
                ]
            )

    write_headers_and_rows(svc, sid, MASTER_SHEET, MASTER_HEADERS, new_master, clear=True)
    write_headers_and_rows(svc, sid, LANE_A_SHEET, LANE_A_HEADERS, lane_a, clear=True)
    write_headers_and_rows(svc, sid, EXEC_SHEET, EXEC_HEADERS, new_exec, clear=True)
    LOG.info(
        "seed done master=%s lane_a=%s exec=%s runId=%s",
        len(new_master),
        len(lane_a),
        len(new_exec),
        run_id,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
