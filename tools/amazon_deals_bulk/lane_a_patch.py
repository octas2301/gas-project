# -*- coding: utf-8 -*-
"""
レーンA（discounted_price）パッチ生成。既定はJSONのみ（API送信しない）。

§10.8: B登録確認後、A期間で dry_run→prod。本番APIは社長「やってよい」必須。

例:
  python lane_a_patch.py
  python lane_a_patch.py --gap 0
  python lane_a_patch.py --sku originalM-1803--KOUS--Cinderellas-b
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from paths import load_config  # noqa: E402
from schedule_class import parse_ymd  # noqa: E402
from sheet_schema import LANE_A, SALE_SHEET  # noqa: E402
from sheets_io import read_sheet_rows, sheets_service  # noqa: E402

LOG = logging.getLogger("amazon_deals_bulk.lane_a_patch")
MARKETPLACE_JP = "A1VC38T7YXB528"


def _truthy(v: Any) -> bool:
    return str(v or "").strip().upper() in ("TRUE", "はい", "YES", "Y", "1", "○")


def _float(v: Any) -> Optional[float]:
    if v is None or str(v).strip() == "":
        return None
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None


def iso_start(d: date) -> str:
    # JST 0:00 相当を UTC 表現（SP-APIは ISO-8601）
    return f"{d.isoformat()}T00:00:00+09:00"


def iso_end(d: date) -> str:
    return f"{d.isoformat()}T23:59:59+09:00"


def build_discounted_price_patch(
    *,
    marketplace_id: str,
    currency: str,
    sale_price: float,
    start: date,
    end: date,
    our_price: Optional[float] = None,
) -> Dict[str, Any]:
    """Listings Items PATCH 用（purchasable_offer.discounted_price）。"""
    offer: Dict[str, Any] = {
        "marketplace_id": marketplace_id,
        "currency": currency,
        "discounted_price": [
            {
                "schedule": [
                    {
                        "value_with_tax": sale_price,
                        "start_at": iso_start(start),
                        "end_at": iso_end(end),
                    }
                ]
            }
        ],
    }
    if our_price is not None and our_price > 0:
        offer["our_price"] = [{"schedule": [{"value_with_tax": our_price}]}]
    return {
        "productType": "PRODUCT",
        "patches": [
            {
                "op": "replace",
                "path": "/attributes/purchasable_offer",
                "value": [offer],
            }
        ],
    }


def load_a_rows(cfg: dict, *, sku_filter: Optional[str], include_unapproved: bool) -> List[Dict[str, Any]]:
    svc = sheets_service(write=False)
    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    _h, rows = read_sheet_rows(svc, sid, SALE_SHEET)
    out = []
    for r in rows:
        if str(r.get("レーン") or "").strip() != LANE_A:
            continue
        if not _truthy(r.get("有効")):
            continue
        st = str(r.get("状態") or "").strip()
        if st in ("停止", "見送り", "終了", "失敗"):
            continue
        if not include_unapproved and not _truthy(r.get("承認済")):
            continue
        if sku_filter and str(r.get("SKU") or "").strip() != sku_filter:
            continue
        start = parse_ymd(r.get("開始日"))
        end = parse_ymd(r.get("終了日"))
        if not start or not end:
            continue
        price = _float(r.get("セール価格")) or _float(r.get("タイムセール価格_確定"))
        if not price or price <= 0:
            continue
        out.append(r)
    return out


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="レーンA discounted_price パッチJSON生成")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--sku", type=str, default=None, help="1SKUだけ（§10.8検証用）")
    ap.add_argument(
        "--include-unapproved",
        action="store_true",
        help="承認済でなくても下書きAを含める（検証用）",
    )
    ap.add_argument(
        "--before-b-only",
        action="store_true",
        help="今日以降・直近B開始より前のAだけ（Smile前隙）",
    )
    args = ap.parse_args(argv)
    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))

    rows = load_a_rows(
        cfg, sku_filter=args.sku, include_unapproved=bool(args.include_unapproved)
    )
    if args.before_b_only:
        today = date.today()
        # シート上のB開始の最早
        svc = sheets_service(write=False)
        sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
        _h, all_rows = read_sheet_rows(svc, sid, SALE_SHEET)
        b_starts = []
        for r in all_rows:
            if str(r.get("レーン") or "").startswith("B"):
                d = parse_ymd(r.get("開始日"))
                if d and d >= today:
                    b_starts.append(d)
        b0 = min(b_starts) if b_starts else None
        filtered = []
        for r in rows:
            s, e = parse_ymd(r.get("開始日")), parse_ymd(r.get("終了日"))
            if not s or not e or e < today:
                continue
            if b0 and e >= b0:
                continue
            filtered.append(r)
        rows = filtered

    items = []
    for r in rows:
        start = parse_ymd(r.get("開始日"))
        end = parse_ymd(r.get("終了日"))
        assert start and end
        sale = _float(r.get("セール価格")) or _float(r.get("タイムセール価格_確定"))
        our = _float(r.get("通常価格")) or _float(r.get("出品者価格_SC"))
        assert sale
        body = build_discounted_price_patch(
            marketplace_id=str(cfg.get("marketplace_id") or MARKETPLACE_JP),
            currency=str(r.get("通貨") or "JPY"),
            sale_price=sale,
            start=start,
            end=end,
            our_price=our,
        )
        items.append(
            {
                "sku": str(r.get("SKU") or "").strip(),
                "asin": str(r.get("ASIN") or "").strip().upper(),
                "sale_id": r.get("sale_id"),
                "start": start.isoformat(),
                "end": end.isoformat(),
                "sale_price": sale,
                "our_price": our,
                "approved": _truthy(r.get("承認済")),
                "patch": body,
            }
        )
        print(
            f"A {r.get('SKU')} {start}..{end} price={sale} approved={_truthy(r.get('承認済'))}"
        )

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_dir = HERE / "_work"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"lane_a_patches_{stamp}.json"
    payload = {
        "generated_at": stamp,
        "count": len(items),
        "note": "API送信は未実施。P1aは社長『やってよい』後。§10.8は1SKUから。",
        "items": items,
    }
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info("出力: %s count=%s", out_path, len(items))
    if not items:
        print("対象Aなし。syncでA行を作り、検証時は --include-unapproved を検討")
        return 1
    print(f"JSON: {out_path}")
    print("次: 承認後に SP-API Listings PATCH（先に VALIDATION_PREVIEW）。本番は都度承認。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
