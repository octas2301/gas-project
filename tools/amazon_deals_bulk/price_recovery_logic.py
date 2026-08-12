# -*- coding: utf-8 -*-
"""
実質戻し（最終売価固定＋ポイント減衰）の計算・提案。

人必須入力: 目標売価円（最終売価）／販促ポイント%
表示: 販促ポイント円・実質価格円
提案: 減衰期間・間隔・段%（減衰段%列）

GAS TimeSalePriceRecovery.js と同式。
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Set, Tuple

from schedule_class import parse_ymd
from sheet_schema import (
    EFFECTIVE_PRICE_COL,
    LANE_B,
    POINT_BEFORE_COL,
    POINT_STATUS_COL,
    PRICE_CURRENT_SELL_COL,
    PRICE_RECOVERY_INTERVAL_COL,
    PRICE_RECOVERY_NEXT_COL,
    PRICE_RECOVERY_PERIOD_COL,
    PRICE_RECOVERY_PROGRESS_COL,
    PRICE_RECOVERY_STATUS_COL,
    PRICE_RECOVERY_STEP_COL,
    PRICE_TARGET_COL,
    PROMO_POINT_PCT_COL,
    PROMO_POINT_YEN_COL,
    TAPER_ACTIVE_COL,
    TAPER_LAST_RUN_COL,
    TAPER_REQUEST_COL,
    TAPER_START_COL,
)

PERIOD_OPTIONS = ("1か月", "2か月", "3か月", "4か月", "5か月", "6か月")
INTERVAL_OPTIONS = ("1週間", "2週間", "1か月", "2か月")

STATUS_NOT_STARTED = "未開始"
STATUS_IN_PROGRESS = "進行中"
STATUS_DONE = "完了"
STATUS_STOPPED = "停止"
STATUS_SKIP = "見送り"

TERMINAL_STATUSES = frozenset({STATUS_DONE, STATUS_STOPPED, STATUS_SKIP})

POINT_STATUS_APPLIED = "期間中適用済"

# 販促ポイント%の上限（Amazonポイントセントラルは概ね1〜50）
PROMO_PCT_MAX = 50
DEFAULT_END_PCT = 1


def _num(v: Any) -> Optional[float]:
    if v is None or str(v).strip() == "":
        return None
    try:
        return float(str(v).replace(",", "").replace("円", "").replace("%", "").strip())
    except ValueError:
        return None


def _truthy(v: Any) -> bool:
    return str(v or "").strip().upper() in ("TRUE", "はい", "YES", "Y", "1", "○")


def _interval_weeks(label: str) -> int:
    return {"1週間": 1, "2週間": 2, "1か月": 4, "2か月": 8}.get(label, 2)


def interval_timedelta(label: str) -> timedelta:
    return timedelta(days=7 * _interval_weeks(str(label or "").strip() or "2週間"))


def promo_points_yen(final_price: float, promo_pct: float) -> int:
    """販促ポイント円 = round(最終売価 × 販促%/100)。"""
    return int(round(float(final_price) * float(promo_pct) / 100.0))


def effective_price(final_price: float, promo_pct: float) -> int:
    """実質価格円 = 最終売価 − 販促ポイント円。"""
    return int(round(float(final_price) - promo_points_yen(final_price, promo_pct)))


def _period_months_from_promo_pct(pct: float) -> int:
    """販促ポイント%の大きさ → 推奨減衰期間（月）。"""
    if pct <= 10:
        return 2
    if pct <= 20:
        return 3
    if pct <= 35:
        return 3
    if pct <= 45:
        return 4
    return 5


def fill_display_fields(row: Dict[str, Any]) -> Dict[str, Any]:
    """目標売価・販促ポイント%から円・実質を埋めた dict（シート書込用）。"""
    p = _num(row.get(PRICE_TARGET_COL))
    a = _num(row.get(PROMO_POINT_PCT_COL))
    out = dict(row)
    if p is None or a is None:
        return out
    yen = promo_points_yen(p, a)
    out[PROMO_POINT_YEN_COL] = yen
    out[EFFECTIVE_PRICE_COL] = int(round(p - yen))
    return out


def propose_points_taper(
    final_price: float,
    promo_pct: float,
    *,
    end_pct: Optional[float] = None,
) -> Dict[str, Any]:
    """
    人入力の最終売価＋販促ポイント%から減衰スケジュールを提案。

    戻りキーはシート列名寄せ（減衰期間／減衰段%／減衰間隔 等）。
    """
    if final_price <= 0:
        raise ValueError("目標売価円（最終売価）は正の数が必要です")
    if promo_pct < 0 or promo_pct > PROMO_PCT_MAX:
        raise ValueError("販促ポイント%%は 0〜%s です" % PROMO_PCT_MAX)

    e = DEFAULT_END_PCT if end_pct is None else float(end_pct)
    if e < 0:
        e = 0.0
    if promo_pct < e:
        raise ValueError("販促ポイント%%は最終終着%%以上にしてください")

    yen = promo_points_yen(final_price, promo_pct)
    eff = int(round(final_price - yen))

    interval = "2週間"
    iw = _interval_weeks(interval)
    months = _period_months_from_promo_pct(promo_pct)
    weeks_budget = max(iw, int(round(months * 4.345)))
    n_steps = max(2, weeks_budget // iw)
    delta = promo_pct - e
    step_pct = max(1, int(round(delta / float(n_steps))))
    # 段数が膨らみすぎたら期間を延ばして再計算（最大6か月）
    guard = 0
    while step_pct * n_steps < delta - 0.5 and months < 6 and guard < 4:
        months += 1
        weeks_budget = max(iw, int(round(months * 4.345)))
        n_steps = max(2, weeks_budget // iw)
        step_pct = max(1, int(round(delta / float(n_steps))))
        guard += 1

    period = "%sか月" % months
    progress = (
        "未開始｜最終売価%g／販促ポイント%s%%（%s円）／実質%g｜"
        "終着%s%%｜段−%s%%×%s｜目安%s段・期間%s"
        % (
            final_price,
            int(promo_pct) if float(promo_pct).is_integer() else promo_pct,
            yen,
            eff,
            int(e) if float(e).is_integer() else e,
            step_pct,
            interval,
            n_steps,
            period,
        )
    )
    return {
        PRICE_TARGET_COL: int(final_price) if float(final_price).is_integer() else final_price,
        PROMO_POINT_PCT_COL: int(promo_pct) if float(promo_pct).is_integer() else promo_pct,
        PROMO_POINT_YEN_COL: yen,
        EFFECTIVE_PRICE_COL: eff,
        PRICE_RECOVERY_PERIOD_COL: period,
        PRICE_RECOVERY_STEP_COL: step_pct,
        PRICE_RECOVERY_INTERVAL_COL: interval,
        PRICE_RECOVERY_PROGRESS_COL: progress,
        PRICE_RECOVERY_STATUS_COL: STATUS_NOT_STARTED,
        PRICE_CURRENT_SELL_COL: int(final_price) if float(final_price).is_integer() else final_price,
        "段数": n_steps,
        "所要週": int(weeks_budget),
        "開始%": float(promo_pct),
        "終着%": float(e),
        "段%": step_pct,
        "販促ポイント円": yen,
        "実質価格円": eff,
    }


def propose_recovery(final_price: float, promo_pct: float, **kwargs: Any) -> Dict[str, Any]:
    """互換エイリアス。第2引数は販促ポイント%（旧: 販促売価は廃止）。"""
    return propose_points_taper(final_price, promo_pct, **kwargs)


def end_pct_from_row(row: Dict[str, Any]) -> float:
    v = _num(row.get(POINT_BEFORE_COL))
    return float(v) if v is not None else float(DEFAULT_END_PCT)


def propose_from_row(row: Dict[str, Any]) -> Dict[str, Any]:
    p = _num(row.get(PRICE_TARGET_COL))
    a = _num(row.get(PROMO_POINT_PCT_COL))
    if p is None or a is None:
        raise ValueError("目標売価円と販促ポイント%を入力してください")
    return propose_points_taper(p, a, end_pct=end_pct_from_row(row))


def next_points_percent(current: float, step_pct: float, end_pct: float) -> float:
    """1段下げ後のポイント%。終着を下回らない。"""
    if step_pct <= 0:
        raise ValueError("減衰段%は正の数が必要です")
    nxt = current - step_pct
    if nxt <= end_pct + 1e-9:
        return float(int(end_pct)) if float(end_pct).is_integer() else float(end_pct)
    return float(int(nxt)) if float(nxt).is_integer() else float(nxt)


def build_our_price_patch(
    *,
    marketplace_id: str,
    currency: str,
    our_price: float,
    product_type: str = "PRODUCT",
) -> Dict[str, Any]:
    """売価スナップ用 Listings PATCH（段階上げはしない）。"""
    offer: Dict[str, Any] = {
        "marketplace_id": marketplace_id,
        "currency": currency or "JPY",
        "our_price": [{"schedule": [{"value_with_tax": float(our_price)}]}],
    }
    return {
        "productType": product_type or "PRODUCT",
        "patches": [
            {
                "op": "replace",
                "path": "/attributes/purchasable_offer",
                "value": [offer],
            }
        ],
    }


def skus_in_active_b(
    sales: List[Dict[str, Any]],
    *,
    today: date,
) -> Set[str]:
    out: Set[str] = set()
    skip_st = {"終了", "見送り", "失敗", "停止", "延期"}
    for r in sales:
        if str(r.get("レーン") or "").strip() != LANE_B:
            continue
        st = str(r.get("状態") or "").strip()
        if st in skip_st:
            continue
        start = parse_ymd(r.get("開始日"))
        end = parse_ymd(r.get("終了日"))
        if not start or not end:
            continue
        if start <= today <= end:
            sku = str(r.get("SKU") or "").strip()
            if sku:
                out.add(sku)
    return out


def points_blocks_recovery(row: Dict[str, Any]) -> bool:
    """期間中適用済のまま Amazon 減衰フィードを送るのは原則禁止（先に restore＝減衰中%）。カレンダー更新は可。"""
    return str(row.get(POINT_STATUS_COL) or "").strip() == POINT_STATUS_APPLIED


def _truthy_flag(v: Any) -> bool:
    if v is True:
        return True
    return str(v or "").strip().upper() in ("TRUE", "はい", "YES", "Y", "1", "○")


def taper_requested(row: Dict[str, Any]) -> bool:
    return _truthy_flag(row.get(TAPER_REQUEST_COL))


def taper_active_percent(row: Dict[str, Any]) -> Optional[float]:
    """運用目標% = 減衰中 → 出品者現在 → 販促%。"""
    v = _num(row.get(TAPER_ACTIVE_COL))
    if v is not None:
        return v
    from points_logic import current_percent  # 遅延 import

    cur = current_percent(row)
    if cur is not None:
        return float(cur)
    return _num(row.get(PROMO_POINT_PCT_COL))


def taper_start_date(row: Dict[str, Any]) -> Optional[date]:
    return parse_ymd(row.get(TAPER_START_COL))


def calendar_step_count(row: Dict[str, Any], today: date) -> int:
    """減衰開始日を1段目。開始日前は0。"""
    start = taper_start_date(row)
    if not start or today < start:
        return 0
    inv = interval_timedelta(
        str(row.get(PRICE_RECOVERY_INTERVAL_COL) or row.get("戻し間隔") or "").strip()
        or "2週間"
    )
    days = inv.days or 14
    return 1 + ((today - start).days // days)


def calendar_active_pct(row: Dict[str, Any], today: date) -> float:
    """カレンダー上のあるべき減衰中%。B中でも店頭と独立して進む。"""
    promo = _num(row.get(PROMO_POINT_PCT_COL))
    if promo is None:
        raise ValueError("販促ポイント%が空です")
    end = end_pct_from_row(row)
    step = _num(row.get(PRICE_RECOVERY_STEP_COL) or row.get("戻し価格円"))
    if not step or step <= 0:
        raise ValueError("減衰段%が不正です")
    n = calendar_step_count(row, today)
    if n <= 0:
        return float(promo)
    return next_points_percent(promo, n * step, end)


def calendar_next_date(row: Dict[str, Any], today: date) -> Optional[date]:
    start = taper_start_date(row)
    if not start:
        return None
    inv = interval_timedelta(
        str(row.get(PRICE_RECOVERY_INTERVAL_COL) or row.get("戻し間隔") or "").strip()
        or "2週間"
    )
    days = inv.days or 14
    end = end_pct_from_row(row)
    want = calendar_active_pct(row, today)
    if want <= end + 1e-9:
        return None
    if today < start:
        return start
    n = calendar_step_count(row, today)
    return start + timedelta(days=n * days)


def plan_calendar_sync(row: Dict[str, Any], *, today: date) -> Dict[str, Any]:
    """減衰開始日あり: シート減衰中%をカレンダー位置へ同期（1段ずつではなく追いつき可）。"""
    sku = str(row.get("SKU") or "").strip()
    want = calendar_active_pct(row, today)
    cur = taper_active_percent(row)
    if cur is None:
        cur = want
    end = end_pct_from_row(row)
    next_d = calendar_next_date(row, today)
    done = want <= end + 1e-9
    n = calendar_step_count(row, today)
    if done:
        status = STATUS_DONE
    elif n <= 0:
        status = STATUS_NOT_STARTED
    else:
        status = STATUS_IN_PROGRESS
    promo = _num(row.get(PROMO_POINT_PCT_COL)) or 0
    step = _num(row.get(PRICE_RECOVERY_STEP_COL) or row.get("戻し価格円")) or 0
    sheet_in_sync = abs(float(cur) - float(want)) < 1e-9
    progress = (
        "%s｜カレンダー減衰中%g%%（販促%g・終着%g・段%d×−%g）｜次回=%s"
        % (
            status,
            want,
            promo,
            end,
            n,
            step,
            next_d.isoformat() if next_d else "-",
        )
    )
    return {
        "sku": sku,
        "from_pct": cur,
        "to_pct": want,
        "done": done,
        "next_date": next_d,
        "status": status,
        "progress": progress,
        "skip_api": sheet_in_sync,
        "calendar": True,
    }


def last_run_date(row: Dict[str, Any]) -> Optional[date]:
    raw = row.get(TAPER_LAST_RUN_COL)
    d = parse_ymd(raw)
    if d:
        return d
    s = str(raw or "").strip()
    if len(s) >= 10:
        return parse_ymd(s[:10])
    return None


def has_taper_plan(row: Dict[str, Any]) -> bool:
    p = _num(row.get(PRICE_TARGET_COL) or row.get("最終売価円"))
    a = _num(row.get(PROMO_POINT_PCT_COL)) or _num(row.get(TAPER_ACTIVE_COL))
    step = _num(row.get(PRICE_RECOVERY_STEP_COL) or row.get("戻し価格円"))
    interval = str(
        row.get(PRICE_RECOVERY_INTERVAL_COL) or row.get("戻し間隔") or ""
    ).strip()
    return bool(p and a is not None and step and step > 0 and interval)


def plan_one_points_step(
    row: Dict[str, Any],
    *,
    today: date,
    current_pct: Optional[float] = None,
) -> Dict[str, Any]:
    """
    ポイント減衰1段。current_pct 省略時は 減衰中%→出品者現在%→販促%。
    """
    sku = str(row.get("SKU") or "").strip()
    end = end_pct_from_row(row)
    step = _num(row.get(PRICE_RECOVERY_STEP_COL) or row.get("戻し価格円"))
    interval = str(
        row.get(PRICE_RECOVERY_INTERVAL_COL) or row.get("戻し間隔") or ""
    ).strip()
    if not step or step <= 0 or not interval:
        raise ValueError("減衰計画（段%／間隔）が不足")
    cur = current_pct if current_pct is not None else taper_active_percent(row)
    if cur is None:
        raise ValueError("現在ポイント%が不明です")
    if cur <= end + 1e-9:
        return {
            "sku": sku,
            "from_pct": cur,
            "to_pct": cur,
            "done": True,
            "next_date": None,
            "status": STATUS_DONE,
            "progress": "完了｜既に終着%%以下（現在%g／終着%g）" % (cur, end),
            "skip_api": True,
        }
    to_pct = next_points_percent(cur, step, end)
    done = to_pct <= end + 1e-9
    next_d = None if done else (today + interval_timedelta(interval))
    status = STATUS_DONE if done else STATUS_IN_PROGRESS
    progress = (
        "%s｜ポイント%g%%→%g%%（終着%g・段−%g）｜次回=%s"
        % (
            status,
            cur,
            to_pct,
            end,
            step,
            next_d.isoformat() if next_d else "-",
        )
    )
    return {
        "sku": sku,
        "from_pct": cur,
        "to_pct": to_pct,
        "done": done,
        "next_date": next_d,
        "status": status,
        "progress": progress,
        "skip_api": False,
    }


# 旧名互換（売価段は廃止。呼び出し側は points 段へ）
def plan_one_step(row: Dict[str, Any], *, today: date) -> Dict[str, Any]:
    return plan_one_points_step(row, today=today)


def select_due_recovery_rows(
    master: List[Dict[str, Any]],
    *,
    today: date,
    sku_filter: Optional[str] = None,
    include_start: bool = False,
    enabled_only: bool = True,
) -> List[Dict[str, Any]]:
    want = (sku_filter or "").strip()
    out: List[Dict[str, Any]] = []
    for r in master:
        sku = str(r.get("SKU") or "").strip()
        if not sku:
            continue
        if want and sku != want:
            continue
        if enabled_only and not _truthy(r.get("有効")):
            continue
        st = str(
            r.get(PRICE_RECOVERY_STATUS_COL) or r.get("戻し状態") or ""
        ).strip()
        if st in TERMINAL_STATUSES:
            continue
        if not has_taper_plan(r):
            continue
        start = taper_start_date(r)
        if start:
            if today < start and not taper_requested(r):
                continue
            try:
                cal_pct = calendar_active_pct(r, today)
            except ValueError:
                continue
            cur = _num(r.get(TAPER_ACTIVE_COL))
            next_d = parse_ymd(
                r.get(PRICE_RECOVERY_NEXT_COL) or r.get("次回戻し日")
            )
            drift = cur is None or abs(float(cur) - cal_pct) > 1e-9
            if taper_requested(r) or drift or (next_d is not None and next_d <= today):
                out.append(r)
            continue
        if taper_requested(r):
            out.append(r)
            continue
        next_d = parse_ymd(
            r.get(PRICE_RECOVERY_NEXT_COL) or r.get("次回戻し日")
        )
        if next_d is not None:
            if next_d <= today:
                out.append(r)
            continue
        if include_start and (st in ("", STATUS_NOT_STARTED) or st == STATUS_IN_PROGRESS):
            out.append(r)
    return out


def apply_plan_to_row(row: Dict[str, Any], plan: Dict[str, Any]) -> None:
    """減衰1段成功後のシート更新用（状態・次回・進捗・減衰中%）。"""
    row[PRICE_RECOVERY_STATUS_COL] = plan["status"]
    row[PRICE_RECOVERY_PROGRESS_COL] = plan["progress"]
    nd = plan.get("next_date")
    row[PRICE_RECOVERY_NEXT_COL] = nd.isoformat() if nd else ""
    to_pct = plan.get("to_pct")
    if to_pct is not None:
        row[TAPER_ACTIVE_COL] = to_pct
    row[TAPER_REQUEST_COL] = ""


def recovery_fields_from_row(r: Dict[str, Any]) -> Dict[str, str]:
    """旧列名（最終売価円／戻し*）も現行名として拾う。"""
    target = r.get(PRICE_TARGET_COL)
    if target is None or str(target).strip() == "":
        target = r.get("最終売価円")
    period = r.get(PRICE_RECOVERY_PERIOD_COL) or r.get("戻し期間")
    step = r.get(PRICE_RECOVERY_STEP_COL) or r.get("戻し価格円")
    interval = r.get(PRICE_RECOVERY_INTERVAL_COL) or r.get("戻し間隔")
    progress = r.get(PRICE_RECOVERY_PROGRESS_COL) or r.get("戻し進捗")
    status = r.get(PRICE_RECOVERY_STATUS_COL) or r.get("戻し状態")
    nxt = r.get(PRICE_RECOVERY_NEXT_COL) or r.get("次回戻し日")
    return {
        PRICE_TARGET_COL: str(target or ""),
        PROMO_POINT_PCT_COL: str(r.get(PROMO_POINT_PCT_COL) or ""),
        PROMO_POINT_YEN_COL: str(r.get(PROMO_POINT_YEN_COL) or ""),
        EFFECTIVE_PRICE_COL: str(r.get(EFFECTIVE_PRICE_COL) or ""),
        PRICE_RECOVERY_PERIOD_COL: str(period or ""),
        PRICE_RECOVERY_STEP_COL: str(step or ""),
        PRICE_RECOVERY_INTERVAL_COL: str(interval or ""),
        PRICE_RECOVERY_PROGRESS_COL: str(progress or ""),
        PRICE_RECOVERY_STATUS_COL: str(status or ""),
        PRICE_RECOVERY_NEXT_COL: str(nxt or ""),
        PRICE_CURRENT_SELL_COL: str(r.get(PRICE_CURRENT_SELL_COL) or ""),
        TAPER_ACTIVE_COL: str(r.get(TAPER_ACTIVE_COL) or ""),
        TAPER_REQUEST_COL: str(r.get(TAPER_REQUEST_COL) or ""),
        TAPER_LAST_RUN_COL: str(r.get(TAPER_LAST_RUN_COL) or ""),
        TAPER_START_COL: str(r.get(TAPER_START_COL) or ""),
    }


# テスト互換の薄いラッパ（売価段は廃止）
def next_price(current: float, step: float, target: float) -> float:
    """互換: 売価段は使わない。目標を超えない加算（スナップ検証用）。"""
    nxt = current + step
    return min(nxt, target)
