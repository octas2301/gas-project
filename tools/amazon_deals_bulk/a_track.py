# -*- coding: utf-8 -*-
"""レーンA実施フラグをマスタ／施策へ書き戻す。"""
from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from sheet_schema import (
    A_DONE_COL,
    A_LOG_COL,
    A_PERIOD_COL,
    A_PRICE_COL,
    A_SENT_AT_COL,
    A_TRACK_COLS,
    LANE_A,
    MASTER_HEADERS,
    MASTER_SHEET,
    SALE_A_TRACK_COLS,
    SALE_HEADERS,
    SALE_SHEET,
)
from sheets_io import (
    ensure_headers_append,
    read_sheet_rows,
    update_row_fields,
)

LOG = logging.getLogger("amazon_deals_bulk.a_track")
HERE = Path(__file__).resolve().parent


def _jst_now_str() -> str:
    # 簡易: ローカル表記（Windowsも可）。厳密TZ不要。
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_a_track_headers(svc, sid: str) -> None:
    ensure_headers_append(svc, sid, MASTER_SHEET, list(MASTER_HEADERS))
    ensure_headers_append(svc, sid, SALE_SHEET, list(SALE_HEADERS))


def find_row_1based(headers: List[str], rows: List[Dict[str, Any]], pred) -> Optional[int]:
    for i, r in enumerate(rows):
        if pred(r):
            return i + 2  # header=1
    return None


def record_lane_a_send(
    svc,
    sid: str,
    *,
    sku: str,
    asin: str,
    start: str,
    end: str,
    price: Any,
    log_name: str,
    prod: bool,
    http_status: int,
) -> Dict[str, Any]:
    """
    lane_a_send 成功後に呼ぶ。
    prod かつ HTTP<400 のとき A実施=はい。dry_run は A実施を変えずログのみメモ可。
    """
    ensure_a_track_headers(svc, sid)
    sku = str(sku or "").strip()
    period = f"{start}～{end}"
    log_ref = str(log_name or "").strip()
    sent_at = _jst_now_str()
    done = "はい" if (prod and http_status < 400) else "dry_run"
    if not prod:
        # dry_runは実施扱いにしない（空のまま or dry_run 表記のみ更新したい場合）
        master_fields = {
            A_SENT_AT_COL: sent_at,
            A_PERIOD_COL: period,
            A_PRICE_COL: price,
            A_LOG_COL: log_ref + " (dry_run)",
        }
        # A実施は既存「はい」を消さない
    else:
        master_fields = {
            A_DONE_COL: "はい" if http_status < 400 else "失敗",
            A_SENT_AT_COL: sent_at,
            A_PERIOD_COL: period,
            A_PRICE_COL: price,
            A_LOG_COL: log_ref,
        }

    mh, mrows = read_sheet_rows(svc, sid, MASTER_SHEET)
    mrow = find_row_1based(
        mh,
        mrows,
        lambda r: str(r.get("SKU") or "").strip() == sku
        or (asin and str(r.get("ASIN") or "").strip().upper() == asin.upper()),
    )
    n_m = 0
    if mrow:
        # dry_run時は A実施キーを送らない
        fields = dict(master_fields)
        if not prod and A_DONE_COL in fields:
            del fields[A_DONE_COL]
        n_m = update_row_fields(svc, sid, MASTER_SHEET, mrow, fields)
    else:
        LOG.warning("マスタにSKUなし: %s", sku)

    sh, srows = read_sheet_rows(svc, sid, SALE_SHEET)
    srow = find_row_1based(
        sh,
        srows,
        lambda r: str(r.get("SKU") or "").strip() == sku
        and str(r.get("レーン") or "").startswith("A")
        and (
            str(r.get("開始日") or "").startswith(str(start)[:10])
            or str(r.get("開始日") or "") == str(start)
        ),
    )
    if srow is None:
        srow = find_row_1based(
            sh,
            srows,
            lambda r: str(r.get("SKU") or "").strip() == sku
            and str(r.get("レーン") or "").startswith("A"),
        )
    n_s = 0
    if srow:
        sale_fields = {
            A_SENT_AT_COL: sent_at,
            A_LOG_COL: log_ref + (" (dry_run)" if not prod else ""),
        }
        if prod:
            sale_fields[A_DONE_COL] = "はい" if http_status < 400 else "失敗"
        n_s = update_row_fields(svc, sid, SALE_SHEET, srow, sale_fields)
    else:
        LOG.warning("施策A行なし（マスタのみ更新）: %s", sku)

    return {"master_cells": n_m, "sale_cells": n_s, "done": done, "log": log_ref}


def backfill_from_work(svc, sid: str, work_dir: Optional[Path] = None) -> List[Dict[str, Any]]:
    """_work/lane_a_send_*.json から本番成功分をシートへ反映。"""
    ensure_a_track_headers(svc, sid)
    d = work_dir or (HERE / "_work")
    results = []
    if not d.is_dir():
        return results
    files = sorted(d.glob("lane_a_send_*.json"))
    for path in files:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception as e:
            LOG.warning("skip %s: %s", path.name, e)
            continue
        if data.get("validation_preview"):
            continue
        http = int(data.get("http_status") or 0)
        if http >= 400:
            continue
        sku = str(data.get("sku") or "").strip()
        asin = str(data.get("asin") or "").strip()
        req = data.get("request") or {}
        # patch body から期間・価格を拾う（構造は lane_a_patch 依存）
        start, end, price = _extract_period_price(req)
        info = record_lane_a_send(
            svc,
            sid,
            sku=sku,
            asin=asin,
            start=start or "",
            end=end or "",
            price=price or "",
            log_name=path.name,
            prod=True,
            http_status=http,
        )
        info["file"] = path.name
        info["sku"] = sku
        results.append(info)
        LOG.info("backfill %s → %s", path.name, info)
    return results


def _extract_period_price(req: Dict[str, Any]) -> Tuple[str, str, Any]:
    """Listings patch JSON から start/end/price をベストエフォート抽出。"""
    start = end = ""
    price: Any = ""
    try:
        attrs = (req.get("productType") and req) or req
        # 一般形: attributes.purchasable_offer[0].discounted_price...
        po = None
        if isinstance(req.get("attributes"), dict):
            po = req["attributes"].get("purchasable_offer")
        if isinstance(po, list) and po:
            offer = po[0] if isinstance(po[0], dict) else {}
            dp = offer.get("discounted_price") or offer.get("our_price")
            if isinstance(dp, list) and dp:
                sch = dp[0].get("schedule") if isinstance(dp[0], dict) else None
                if isinstance(sch, list) and sch:
                    start = str(sch[0].get("start_at") or sch[0].get("value_with_tax") or "")[:10]
                    # price often in value_with_tax
                    price = sch[0].get("value_with_tax") or sch[0].get("value") or ""
                # alternate shapes
                if isinstance(dp[0], dict) and "value_with_tax" in dp[0]:
                    price = price or dp[0].get("value_with_tax")
            # start/end sometimes sibling
            for key in ("discounted_price",):
                block = offer.get(key)
                if isinstance(block, list):
                    for b in block:
                        if not isinstance(b, dict):
                            continue
                        for s in b.get("schedule") or []:
                            if isinstance(s, dict):
                                if s.get("start_at"):
                                    start = str(s.get("start_at"))[:10]
                                if s.get("end_at"):
                                    end = str(s.get("end_at"))[:10]
                                if s.get("value_with_tax") is not None:
                                    price = s.get("value_with_tax")
    except Exception:
        pass
    return start, end, price
