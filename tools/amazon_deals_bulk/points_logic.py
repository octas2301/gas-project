# -*- coding: utf-8 -*-
"""タイムセール_マスタ向けポイント差分ロジック（§10.10 Phase0）。"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from sheet_schema import (
    LEGACY_POINT_CURRENT,
    LEGACY_POINT_TARGET,
    POINT_BEFORE_COL,
    POINT_BEFORE_YEN_COL,
    POINT_CURRENT_COL,
    POINT_CURRENT_YEN_COL,
    POINT_MEMO_COL,
    POINT_PERIOD_COL,
    POINT_PERIOD_YEN_COL,
    POINT_STATUS_COL,
    PROMO_POINT_PCT_COL,
    TAPER_ACTIVE_COL,
    TAPER_START_COL,
)

DEFAULT_PERIOD_PERCENT = 1
MODE_APPLY = "apply"
MODE_RESTORE = "restore"


def parse_percent(v: Any, default: Optional[int] = None) -> Optional[int]:
    s = str(v or "").strip().replace("%", "").replace(",", "")
    if not s:
        return default
    try:
        return int(float(s))
    except ValueError:
        return default


def period_percent(row: Dict[str, Any]) -> int:
    """タイム期間中の出品者付与%。未記入＝1。明示0＝なし。"""
    raw = row.get(POINT_PERIOD_COL)
    if raw is None or str(raw).strip() == "":
        raw = row.get(LEGACY_POINT_TARGET)
    parsed = parse_percent(raw, default=None)
    if parsed is None:
        return DEFAULT_PERIOD_PERCENT
    return max(0, parsed)


def before_percent(row: Dict[str, Any]) -> Optional[int]:
    """最終終着%（列名はセール前ポイント%。減衰フロア。restore先ではない）。"""
    return parse_percent(row.get(POINT_BEFORE_COL), default=None)


def skus_missing_before(rows: List[Dict[str, Any]]) -> List[str]:
    """最終終着%（セール前列）が空のSKU一覧。空なら減衰フロアは1%既定。"""
    out: List[str] = []
    for r in rows:
        sku = str(r.get("SKU") or "").strip()
        if sku and before_percent(r) is None:
            out.append(sku)
    return out


def restore_percent(row: Dict[str, Any], today=None) -> int:
    """B終了後に店頭へ戻す% = カレンダー減衰中（開始日あり）またはシート減衰中。"""
    from datetime import date as _date

    from schedule_class import parse_ymd

    d = today if isinstance(today, _date) else _date.today()
    start = parse_ymd(row.get(TAPER_START_COL))
    if start:
        from price_recovery_logic import calendar_active_pct

        return max(0, int(round(calendar_active_pct(row, d))))
    v = parse_percent(row.get(TAPER_ACTIVE_COL), default=None)
    if v is not None:
        return max(0, v)
    v = parse_percent(row.get(PROMO_POINT_PCT_COL), default=None)
    if v is not None:
        return max(0, v)
    raise ValueError("減衰中ポイント%（または販促ポイント%）が空です（復元不可）")


def current_percent(row: Dict[str, Any]) -> Optional[int]:
    raw = row.get(POINT_CURRENT_COL)
    if raw is None or str(raw).strip() == "":
        raw = row.get(LEGACY_POINT_CURRENT)
    return parse_percent(raw, default=None)


# 互換
def target_percent(row: Dict[str, Any]) -> int:
    return period_percent(row)


def send_percent(row: Dict[str, Any], mode: str) -> int:
    if mode == MODE_RESTORE:
        return restore_percent(row)
    return period_percent(row)


def needs_sync(row: Dict[str, Any], mode: str = MODE_APPLY) -> bool:
    try:
        want = send_percent(row, mode)
    except ValueError:
        return mode == MODE_RESTORE  # 空でも一覧に出してエラーにする
    cur = current_percent(row)
    if cur is None:
        return True
    return cur != want


def select_diff_rows(
    rows: List[Dict[str, Any]],
    *,
    mode: str = MODE_APPLY,
    sku_filter: Optional[str] = None,
    force_all: bool = False,
    enabled_only: bool = True,
    sku_allow: Optional[set] = None,
) -> List[Dict[str, Any]]:
    """
    sku_allow: 指定時はそのSKU集合に限定（施策連動・P0-G4）。
    None ならマスタ有効行全体（従来）。
    """
    out: List[Dict[str, Any]] = []
    want = (sku_filter or "").strip()
    allow = sku_allow
    for r in rows:
        sku = str(r.get("SKU") or "").strip()
        if not sku:
            continue
        if want and sku != want:
            continue
        if allow is not None and sku not in allow:
            continue
        if enabled_only:
            en = str(r.get("有効") or "TRUE").strip()
            if en and en.upper() not in ("TRUE", "はい", "YES", "Y", "1", "○"):
                continue
        if mode == MODE_RESTORE and not force_all:
            try:
                restore_percent(r)
            except ValueError:
                continue
        if force_all or needs_sync(r, mode):
            out.append(r)
    return out


def sale_skus_for_points(
    sales: List[Dict[str, Any]],
    *,
    mode: str,
    today,  # date
    within_days: int = 1,
) -> set:
    """
    施策シート（レーンB）からポイント対象SKUを抽出（P0-G4）。

    apply:
      - 開始まで残り 0..within_days 日、または実施中（開始≦今日≦終了）
    restore:
      - 終了から 0..within_days 日経過（終了当日=0）
    """
    from datetime import date as _date

    from schedule_class import parse_ymd
    from sheet_schema import LANE_B

    if not isinstance(today, _date):
        raise TypeError("today must be date")
    within = max(0, int(within_days))
    keep_states = {
        "",
        "予定",
        "要確認",
        "数量改定済",
        "UL済",
        "アップロード済",
        "実施中",
    }
    out: set = set()
    for r in sales:
        if str(r.get("レーン") or "").strip() != LANE_B:
            continue
        st = str(r.get("状態") or "").strip()
        if st in ("見送り", "終了", "失敗", "停止", "延期"):
            continue
        if st and st not in keep_states:
            continue
        sku = str(r.get("SKU") or "").strip()
        if not sku:
            continue
        start = parse_ymd(r.get("開始日"))
        end = parse_ymd(r.get("終了日"))
        if mode == MODE_APPLY:
            if start and end and start <= today <= end:
                out.add(sku)
                continue
            if start and start >= today:
                rem = (start - today).days
                if 0 <= rem <= within:
                    out.add(sku)
        elif mode == MODE_RESTORE:
            if end and end <= today:
                after = (today - end).days
                if 0 <= after <= within:
                    out.add(sku)
    return out


def build_points_tsv(rows: List[Dict[str, Any]], mode: str = MODE_APPLY) -> str:
    """SC/SP-API Pointsフィード用 TSV（sku / points_percent）。"""
    lines = ["sku\tpoints_percent"]
    for r in rows:
        sku = str(r.get("SKU") or "").strip()
        if not sku:
            continue
        lines.append("%s\t%s" % (sku, send_percent(r, mode)))
    return "\n".join(lines) + ("\n" if lines else "")


def diff_summary(
    rows: List[Dict[str, Any]], mode: str = MODE_APPLY
) -> List[Tuple[str, Optional[int], int, Optional[int]]]:
    """(sku, current, send, before) 一覧。"""
    out = []
    for r in rows:
        sku = str(r.get("SKU") or "").strip()
        if not sku:
            continue
        out.append((sku, current_percent(r), send_percent(r, mode), before_percent(r)))
    return out


def point_fields_from_row(r: Dict[str, Any]) -> Dict[str, str]:
    period = r.get(POINT_PERIOD_COL)
    if period is None or str(period).strip() == "":
        period = r.get(LEGACY_POINT_TARGET) or ""
    current = r.get(POINT_CURRENT_COL)
    if current is None or str(current).strip() == "":
        current = r.get(LEGACY_POINT_CURRENT) or ""
    before = r.get(POINT_BEFORE_COL) or ""
    return {
        POINT_PERIOD_COL: str(period or ""),
        POINT_PERIOD_YEN_COL: str(r.get(POINT_PERIOD_YEN_COL) or ""),
        POINT_BEFORE_COL: str(before or ""),
        POINT_BEFORE_YEN_COL: str(r.get(POINT_BEFORE_YEN_COL) or ""),
        POINT_CURRENT_COL: str(current or ""),
        POINT_CURRENT_YEN_COL: str(r.get(POINT_CURRENT_YEN_COL) or ""),
        POINT_STATUS_COL: str(r.get(POINT_STATUS_COL) or ""),
        POINT_MEMO_COL: str(r.get(POINT_MEMO_COL) or ""),
    }


# --- ポイント状態語彙（§10.10 / P0-G7）---
# 空セル = 未設定（シートに「未設定」と書かなくてもよい）
POINT_STATUS_UNSET = "未設定"
POINT_STATUS_BACKED_UP = "セール前退避済"
POINT_STATUS_APPLIED = "期間中適用済"
POINT_STATUS_RESTORED = "セール前復元済"
POINT_STATUS_FEED_PREFIX = "フィード"  # 例: フィードIN_PROGRESS / フィードFATAL

POINT_STATUS_CANONICAL = frozenset(
    {
        POINT_STATUS_UNSET,
        POINT_STATUS_BACKED_UP,
        POINT_STATUS_APPLIED,
        POINT_STATUS_RESTORED,
    }
)

# 旧表記 → 正
POINT_STATUS_ALIASES = {
    "セール前へ復元済": POINT_STATUS_RESTORED,
    "復元済": POINT_STATUS_RESTORED,
    "適用済": POINT_STATUS_APPLIED,
    "退避済": POINT_STATUS_BACKED_UP,
    "backup": POINT_STATUS_BACKED_UP,
    "applied": POINT_STATUS_APPLIED,
    "restored": POINT_STATUS_RESTORED,
}


def normalize_point_status(raw: Any) -> str:
    s = str(raw or "").strip()
    if not s:
        return POINT_STATUS_UNSET
    if s in POINT_STATUS_ALIASES:
        return POINT_STATUS_ALIASES[s]
    if s in POINT_STATUS_CANONICAL:
        return s
    if s.startswith(POINT_STATUS_FEED_PREFIX):
        return s
    return s


def is_feed_point_status(raw: Any) -> bool:
    return str(raw or "").strip().startswith(POINT_STATUS_FEED_PREFIX)


def status_after_send(mode: str, feed_processing_status: Optional[str] = None) -> str:
    """
    points_send --update-sheet 用。
    feed が DONE／未待機なら apply→期間中適用済 / restore→セール前復元済。
    それ以外は フィード{processingStatus}。
    """
    st = str(feed_processing_status or "").strip()
    if st and st != "DONE":
        return "%s%s" % (POINT_STATUS_FEED_PREFIX, st)
    if mode == MODE_RESTORE:
        return POINT_STATUS_RESTORED
    return POINT_STATUS_APPLIED


def status_after_backup() -> str:
    """points_fetch --write でセール前%を埋めたとき。"""
    return POINT_STATUS_BACKED_UP
