# -*- coding: utf-8 -*-
"""
V30 / Q_fba 取得。

優先順:
1. マスタ列 V30 / Q_fba
2. data/v30.csv（ASIN,V30[,V30_yoy]）
3. SP-API Sales orderMetrics（直近30日＋前年同期間の max）
4. テンプレ seller_quantity → Q_fba のみ
"""
from __future__ import annotations

import csv
import json
import logging
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import urlencode

HERE = Path(__file__).resolve().parent
LOG = logging.getLogger("amazon_deals_bulk.v30")

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None  # type: ignore


def load_v30_csv(path: Optional[Path] = None) -> Dict[str, float]:
    path = path or (HERE / "data" / "v30.csv")
    if not path.is_file():
        return {}
    out: Dict[str, float] = {}
    with path.open(encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        for row in r:
            asin = str(row.get("ASIN") or row.get("asin") or "").strip().upper()
            if not asin:
                continue
            raw = row.get("V30") or row.get("v30") or row.get("units") or ""
            yoy = row.get("V30_yoy") or row.get("v30_yoy") or ""
            try:
                a = float(str(raw).replace(",", "")) if str(raw).strip() else None
            except ValueError:
                a = None
            try:
                b = float(str(yoy).replace(",", "")) if str(yoy).strip() else None
            except ValueError:
                b = None
            if a is None and b is None:
                continue
            if a is None:
                out[asin] = b  # type: ignore
            elif b is None:
                out[asin] = a
            else:
                out[asin] = max(a, b)
    return out


def _load_spapi_cfg() -> Optional[dict]:
    for p in (
        HERE.parent / "spapi_smoke" / "config.local.json",
        HERE.parent / "spapi_listings_write" / "config.local.json",
    ):
        if p.is_file():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                continue
    return None


def _lwa_token(cfg: dict) -> str:
    assert requests is not None
    resp = requests.post(
        "https://api.amazon.com/auth/o2/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": cfg["refresh_token"],
            "client_id": cfg["lwa_client_id"],
            "client_secret": cfg["lwa_client_secret"],
        },
        headers={"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"},
        timeout=60,
    )
    resp.raise_for_status()
    return str(resp.json()["access_token"])


def _order_metrics_units(
    *,
    endpoint: str,
    token: str,
    marketplace_id: str,
    asin: str,
    start: date,
    end: date,
    user_agent: str,
) -> Optional[float]:
    """Sales API orderMetrics の unitCount 合計。失敗時 None。"""
    if requests is None:
        return None
    # interval: start--endExclusive（終了翌日 00:00 JST）
    end_excl = end + timedelta(days=1)
    interval = (
        f"{start.isoformat()}T00:00:00+09:00--"
        f"{end_excl.isoformat()}T00:00:00+09:00"
    )
    qs = urlencode(
        {
            "marketplaceIds": marketplace_id,
            "interval": interval,
            "granularity": "Total",
            "asin": asin,
        }
    )
    url = f"{endpoint.rstrip('/')}/sales/v1/orderMetrics?{qs}"
    headers = {
        "x-amz-access-token": token,
        "user-agent": user_agent,
        "accept": "application/json",
    }
    try:
        resp = requests.get(url, headers=headers, timeout=60)
    except Exception as e:
        LOG.warning("orderMetrics network %s: %s", asin, e)
        return None
    if resp.status_code != 200:
        LOG.warning("orderMetrics HTTP %s asin=%s body=%s", resp.status_code, asin, resp.text[:200])
        return None
    body = resp.json()
    payload = body.get("payload") if isinstance(body, dict) else None
    if not isinstance(payload, list):
        return None
    total = 0.0
    for row in payload:
        if not isinstance(row, dict):
            continue
        u = row.get("unitCount")
        if u is None:
            continue
        try:
            total += float(u)
        except (TypeError, ValueError):
            pass
    return total


def fetch_v30_spapi(asins: list[str], *, today: Optional[date] = None) -> Dict[str, float]:
    """直近30日と前年同期間の max(V30)。権限・失敗時は空。"""
    from qty_logic import v30_windows

    cfg = _load_spapi_cfg()
    if not cfg or requests is None:
        LOG.info("SP-API設定なし → V30 APIスキップ")
        return {}
    if str(cfg.get("refresh_token") or "").startswith("REPLACE"):
        return {}

    today = today or date.today()
    start, end, y_start, y_end = v30_windows(today)
    try:
        token = _lwa_token(cfg)
    except Exception as e:
        LOG.warning("LWA失敗: %s", e)
        return {}

    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    mid = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    ua = str(cfg.get("user_agent") or "OctasDealsQty/1.0")
    out: Dict[str, float] = {}
    for asin in asins:
        a = str(asin or "").strip().upper()
        if not a.startswith("B0"):
            continue
        cur = _order_metrics_units(
            endpoint=endpoint,
            token=token,
            marketplace_id=mid,
            asin=a,
            start=start,
            end=end,
            user_agent=ua,
        )
        yoy = _order_metrics_units(
            endpoint=endpoint,
            token=token,
            marketplace_id=mid,
            asin=a,
            start=y_start,
            end=y_end,
            user_agent=ua,
        )
        vals = [v for v in (cur, yoy) if v is not None]
        if not vals:
            continue
        out[a] = max(vals)
        LOG.info("V30 asin=%s cur=%s yoy=%s → %s", a, cur, yoy, out[a])
    return out


def resolve_v30_map(
    asins: list[str],
    *,
    master_rows: Optional[list] = None,
    csv_path: Optional[Path] = None,
    use_spapi: bool = True,
) -> Dict[str, float]:
    """マスタ → CSV → SP-API の順でマージ（後勝ちではなく先勝ち）。"""
    out: Dict[str, float] = {}
    if master_rows:
        for r in master_rows:
            asin = str(r.get("ASIN") or "").strip().upper()
            raw = r.get("V30")
            if not asin or raw is None or str(raw).strip() == "":
                continue
            try:
                out[asin] = float(str(raw).replace(",", ""))
            except ValueError:
                pass
    for asin, v in load_v30_csv(csv_path).items():
        out.setdefault(asin, v)
    missing = [a for a in asins if a and a not in out]
    if use_spapi and missing:
        for asin, v in fetch_v30_spapi(missing).items():
            out.setdefault(asin, v)
    return out
