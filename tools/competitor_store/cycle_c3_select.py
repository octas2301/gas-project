# -*- coding: utf-8 -*-
"""C3 自動選定（規模帯）。正本: B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md §6.3（2026-08-25）。"""
from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime
from typing import Any

from maker_alias import JITTAI_BRAND, JITTAI_CHECK, JITTAI_EXCLUDE, JITTAI_MAKER, JITTAI_UNCHECKED, jitta_for_maker

BAND_S = "小"
BAND_M = "中"
BAND_L = "大"
BAND_XL = "特大"
BAND_UNK = "不明"
BAND_ORDER_WASH = (BAND_S, BAND_M, BAND_L, BAND_XL)

CATALOG_SPLIT_THRESHOLD = 500
BAND_S_MAX = 100
BAND_M_MAX = 300
BAND_L_MAX = 500
REWALK_DAYS_DEFAULT = 90
K_SMALL_MED = 3

WASH_IN_PROGRESS = frozenset({"列挙中", "分割推定中", "Keepa中", "要分割語"})
HIT_EXCLUDE_NAMES = frozenset({"Generic", "generic", "ノーブランド品", ""})


def _cell_date_ms(v: Any) -> int | None:
    if v is None or v == "":
        return None
    if isinstance(v, (date, datetime)):
        return int(datetime(v.year, v.month, v.day).timestamp() * 1000)
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
        try:
            d = datetime.strptime(s[:10], fmt)
            return int(d.timestamp() * 1000)
        except ValueError:
            pass
    return None


def makers_in_cell(cell: str) -> list[str]:
    out: list[str] = []
    for part in str(cell or "").split("|"):
        p = part.strip()
        if p:
            out.append(p)
    return out


def wash_in_progress_status(status: str) -> bool:
    return str(status or "").strip() in WASH_IN_PROGRESS


def pick_wash_continue_target(
    maker_sum: list[dict],
    *,
    wash_split: list[dict] | None = None,
    wash_queue: list[dict] | None = None,
    skip_queue: bool = False,
) -> str | None:
    for row in maker_sum:
        st = str(row.get("洗い状態") or "").strip()
        if wash_in_progress_status(st):
            m = str(row.get("メーカー") or "").strip()
            if m:
                return m
    for row in wash_split or []:
        if str(row.get("状態") or "").strip() != "済":
            m = str(row.get("メーカー") or "").strip()
            if m:
                return m
    if skip_queue:
        return None
    for row in wash_queue or []:
        if str(row.get("Keepa済") or "").upper() in ("TRUE", "1"):
            continue
        m = str(row.get("メーカー") or "").strip()
        if m:
            return m
    return None


def jitta_wash_eligible(jitta: str) -> bool:
    return str(jitta or "").strip() in (JITTAI_MAKER, JITTAI_BRAND)


def maker_food_zero_only(maker: str, sellers: list[dict]) -> bool:
    seen = False
    for row in sellers:
        if maker not in makers_in_cell(str(row.get("抽出メーカー") or "")):
            continue
        seen = True
        ratio_raw = str(row.get("食品比率") or "").replace("%", "").strip()
        try:
            ratio = float(ratio_raw) if ratio_raw else 0.0
        except ValueError:
            ratio = 0.0
        if ratio > 0:
            return False
    return seen


def catalog_estimate_for_maker(
    maker: str,
    *,
    catalog_estimate: dict[str, int] | None = None,
    wash_queue: list[dict] | None = None,
    wash_progress: str = "",
    row: dict | None = None,
) -> int | None:
    """列挙見積。不明は None。"""
    if row and row.get("Catalog見積") not in (None, ""):
        try:
            return int(row.get("Catalog見積"))
        except (TypeError, ValueError):
            pass
    catalog_estimate = catalog_estimate or {}
    if maker in catalog_estimate:
        return int(catalog_estimate[maker])
    asins = {
        str(r.get("ASIN") or "").strip().upper()
        for r in (wash_queue or [])
        if str(r.get("メーカー") or "").strip() == maker and str(r.get("ASIN") or "").strip()
    }
    if asins:
        return len(asins)
    prog = str(wash_progress or "")
    if "queue=" in prog:
        try:
            return int(prog.split("queue=")[1].split()[0].strip().rstrip(","))
        except (ValueError, IndexError):
            pass
    return None


def catalog_band(estimate: int | None) -> str:
    if estimate is None:
        return BAND_UNK
    n = int(estimate)
    if n <= BAND_S_MAX:
        return BAND_S
    if n <= BAND_M_MAX:
        return BAND_M
    if n <= BAND_L_MAX:
        return BAND_L
    return BAND_XL


def catalog_tier(estimate: int | None) -> str:
    """後方互換: 小・中・大 → A、特大・不明 → B。"""
    b = catalog_band(estimate)
    return "A" if b in (BAND_S, BAND_M, BAND_L) else "B"


def _eligible_row(
    row: dict,
    *,
    sellers: list[dict],
    now_ms: int,
    rewalk_days: int,
) -> bool:
    maker = str(row.get("メーカー") or "").strip()
    if not maker or maker in HIT_EXCLUDE_NAMES:
        return False
    if str(row.get("調査CK") or "").upper() in ("TRUE", "1"):
        return False
    jitta = str(row.get("実態") or "").strip() or jitta_for_maker(maker)
    if jitta in (JITTAI_EXCLUDE, JITTAI_CHECK, JITTAI_UNCHECKED):
        return False
    if not jitta_wash_eligible(jitta):
        return False
    last_ms = _cell_date_ms(row.get("最終調査日"))
    if last_ms is not None and last_ms > now_ms - rewalk_days * 86400 * 1000:
        return False
    if maker_food_zero_only(maker, sellers):
        return False
    return True


def _pool_row(
    row: dict,
    i: int,
    *,
    sellers: list[dict],
    catalog_estimate: dict[str, int] | None,
    wash_queue: list[dict] | None,
    now_ms: int,
    rewalk_days: int,
) -> dict | None:
    if not _eligible_row(row, sellers=sellers, now_ms=now_ms, rewalk_days=rewalk_days):
        return None
    maker = str(row["メーカー"]).strip()
    est = catalog_estimate_for_maker(
        maker,
        catalog_estimate=catalog_estimate,
        wash_queue=wash_queue,
        wash_progress=str(row.get("洗い進捗") or ""),
        row=row,
    )
    band = str(row.get("規模帯") or "").strip() or catalog_band(est)
    hit = int(row.get("セラーヒット") or 0)
    last_ms = _cell_date_ms(row.get("最終調査日"))
    never = last_ms is None
    return {
        "maker": maker,
        "band": band,
        "est": est,
        "hit": hit,
        "hit_ok": hit >= 2,
        "never": never,
        "last_ms": 0 if never else last_ms,
        "row_idx": i,
    }


def _sort_within_band(x: dict) -> tuple:
    # hit>=2 先、その中でヒット降順。0〜1は後。同点は未調査・古い順
    return (
        0 if x["hit_ok"] else 1,
        -x["hit"],
        0 if x["never"] else 1,
        x["last_ms"] if not x["never"] else 0,
        x["row_idx"],
    )


def count_washable_small_med(
    maker_sum: list[dict],
    *,
    sellers: list[dict] | None = None,
    wash_queue: list[dict] | None = None,
    catalog_estimate: dict[str, int] | None = None,
    rewalk_days: int = REWALK_DAYS_DEFAULT,
    now_ms: int | None = None,
) -> int:
    sellers = sellers or []
    now_ms = now_ms if now_ms is not None else int(datetime.now().timestamp() * 1000)
    n = 0
    for i, row in enumerate(maker_sum):
        item = _pool_row(
            row,
            i,
            sellers=sellers,
            catalog_estimate=catalog_estimate,
            wash_queue=wash_queue,
            now_ms=now_ms,
            rewalk_days=rewalk_days,
        )
        if item and item["band"] in (BAND_S, BAND_M):
            n += 1
    return n


def pick_enum_unknown(
    maker_sum: list[dict],
    *,
    sellers: list[dict] | None = None,
    wash_queue: list[dict] | None = None,
    catalog_estimate: dict[str, int] | None = None,
    cap: int = 1,
    rewalk_days: int = REWALK_DAYS_DEFAULT,
    now_ms: int | None = None,
) -> list[str]:
    """規模不明のみ（心拍段3）。"""
    sellers = sellers or []
    now_ms = now_ms if now_ms is not None else int(datetime.now().timestamp() * 1000)
    pool = []
    for i, row in enumerate(maker_sum):
        item = _pool_row(
            row,
            i,
            sellers=sellers,
            catalog_estimate=catalog_estimate,
            wash_queue=wash_queue,
            now_ms=now_ms,
            rewalk_days=rewalk_days,
        )
        if item and item["band"] == BAND_UNK:
            pool.append(item)
    pool.sort(key=_sort_within_band)
    return [x["maker"] for x in pool[:cap]]


def pick_cycle_maker_select(
    maker_sum: list[dict],
    *,
    sellers: list[dict] | None = None,
    wash_split: list[dict] | None = None,
    wash_queue: list[dict] | None = None,
    catalog_estimate: dict[str, int] | None = None,
    cap: int = 1,
    rewalk_days: int = REWALK_DAYS_DEFAULT,
    now_ms: int | None = None,
) -> list[str]:
    """洗い選定。小→中→（両方0なら）大→特大。不明は含めない。"""
    cont = pick_wash_continue_target(maker_sum, wash_split=wash_split, wash_queue=wash_queue)
    if cont:
        return [cont][:cap]

    sellers = sellers or []
    now_ms = now_ms if now_ms is not None else int(datetime.now().timestamp() * 1000)
    by_band: dict[str, list[dict]] = {b: [] for b in BAND_ORDER_WASH}
    for i, row in enumerate(maker_sum):
        item = _pool_row(
            row,
            i,
            sellers=sellers,
            catalog_estimate=catalog_estimate,
            wash_queue=wash_queue,
            now_ms=now_ms,
            rewalk_days=rewalk_days,
        )
        if not item or item["band"] == BAND_UNK:
            continue
        if item["band"] in by_band:
            by_band[item["band"]].append(item)

    for b in BAND_ORDER_WASH:
        by_band[b].sort(key=_sort_within_band)

    # 小・中が1件でもあれば大・特大は出さない（枯渇=0 まで）
    sm = by_band[BAND_S] + by_band[BAND_M]
    if sm:
        ordered = by_band[BAND_S] + by_band[BAND_M]
    elif by_band[BAND_L]:
        ordered = by_band[BAND_L]
    else:
        ordered = by_band[BAND_XL]
    return [x["maker"] for x in ordered[:cap]]


def apply_pick_to_state(maker_sum: list[dict], picked: list[str]) -> list[dict]:
    return deepcopy(maker_sum)
