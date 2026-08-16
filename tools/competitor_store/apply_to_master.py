# -*- coding: utf-8 -*-
"""Cluster mall hits by JAN + set count. Never use hit rank 1 as the price."""
from __future__ import annotations

import re

_SET_PATTERNS = [
    (re.compile(r"×\s*(\d+)\s*缶\s*セット", re.I), False),
    (re.compile(r"(\d+)\s*缶\s*セット", re.I), False),
    (re.compile(r"×\s*(\d+)\s*個\s*セット", re.I), False),
    (re.compile(r"×\s*(\d+)\s*個\s*入", re.I), False),
    (re.compile(r"(\d+)\s*個\s*セット", re.I), False),
    (re.compile(r"(\d+)\s*個\s*入", re.I), False),
    (re.compile(r"(\d+)\s*食\s*入", re.I), False),
    (re.compile(r"×\s*(\d+)\s*袋", re.I), False),
    (re.compile(r"(?<!各)(\d+)\s*袋(?:セット|入)?(?!\s*ずつ)", re.I), False),
    (re.compile(r"(\d+)\s*[Pp]\s*(?:セット)?"), True),
]
_ONE = re.compile(r"\[1袋\]|(?<!各)1\s*袋(?!\s*ずつ)|1\s*個(?:セット)?|1\s*食\s*入?", re.I)
_TOTAL_BAG = re.compile(r"(?:合計|計|全)\s*(\d+)\s*袋", re.I)
_KIND_COUNT = re.compile(r"(?:【)?(\d+)\s*種類?(?:】)?")
_EACH_BAG = re.compile(r"各\s*(\d+)\s*袋")
_EXCLUDE = re.compile(r"ふるさと納税|返礼品|よりどり|種類が選べ|選べるセット|中古品|\b中古\b")
UNIT_OUTLIER_RATIO = 2.0


_FW_DIGIT = str.maketrans("０１２３４５６７８９", "0123456789")


def normalize_fullwidth_digits(s: str) -> str:
    return str(s or "").translate(_FW_DIGIT)


def kind_times_each_bags(name: str) -> int | None:
    """3種類×各5袋 → 15。計が無いミックス用。全角数字も半角化。"""
    s = normalize_fullwidth_digits(name)
    km = _KIND_COUNT.search(s)
    em = _EACH_BAG.search(s)
    if not km or not em:
        return None
    tot = int(km.group(1)) * int(em.group(1))
    return tot if 1 <= tot <= 999 else None


def parse_set_count_from_title(name: str) -> tuple[int | None, bool]:
    s = normalize_fullwidth_digits(str(name or "")).strip()
    if not s:
        return None, False
    tm = _TOTAL_BAG.search(s)
    if tm:
        n = int(tm.group(1))
        if 1 <= n <= 999:
            return n, False
    ke = kind_times_each_bags(s)
    if ke:
        return ke, False
    for rx, from_p in _SET_PATTERNS:
        m = rx.search(s)
        if m:
            n = int(m.group(1))
            if 1 <= n <= 999:
                return n, from_p
    if _ONE.search(s):
        return 1, False
    return None, False


def exclude_competitor_title(name: str) -> bool:
    return bool(_EXCLUDE.search(str(name or "")))


def _num(v) -> float | None:
    if v is None or v == "":
        return None
    try:
        n = float(str(v).replace(",", "").strip())
    except ValueError:
        return None
    if n != n:
        return None
    return n


def effective_price(rec: dict) -> float | None:
    price = _num(rec.get("表示価格"))
    if price is None or price <= 0:
        return None
    mall = str(rec.get("モール") or "")
    if mall.startswith("楽天"):
        yen = _num(rec.get("楽天還元円"))
        if yen is None:
            rate = _num(rec.get("楽天ポイント％")) or 0
            yen = round(price * rate / 100)
        return price - (yen or 0)
    pts = _num(rec.get("Yahooポイント数")) or 0
    return price - pts


def ship_paid(rec: dict) -> bool:
    mall = str(rec.get("モール") or "")
    f = str(rec.get("送料フラグ") if rec.get("送料フラグ") is not None else "").strip()
    if "無料" in f or f in ("0",):
        return False
    if mall.startswith("楽天"):
        return f in ("1", "true", "TRUE")
    if f in ("1", "2"):
        return False
    if f == "3":
        return True
    return bool(f) and f != "0"


def latest_per_shop(rows: list[dict]) -> list[dict]:
    latest: dict[str, dict] = {}
    for rec in rows:
        code = str(rec.get("店商品コード") or "").strip() or str(rec.get("商品URL") or "").strip()
        k = "\t".join(
            [
                str(rec.get("モール") or "").strip(),
                str(rec.get("検索JAN") or "").strip(),
                code,
            ]
        )
        prev = latest.get(k)
        ts = str(rec.get("取得日時") or "")
        if not prev or ts >= str(prev.get("取得日時") or ""):
            latest[k] = rec
    return list(latest.values())


def filter_outlier_set_unit_prices(by_set: dict) -> dict:
    keys = [k for k in (by_set or {}) if str(k).isdigit()]
    if len(keys) < 2:
        return dict(by_set or {})
    units = {}
    for k in keys:
        rec = by_set[k]
        n = int(k)
        if rec and n >= 1 and rec.get("priceIncl", 0) > 0:
            units[k] = rec["priceIncl"] / n
    keep = {}
    uk = list(units)
    for k in uk:
        others = sorted(units[x] for x in uk if x != k)
        if not others:
            keep[k] = by_set[k]
            continue
        med = others[(len(others) - 1) // 2]
        if med > 0 and units[k] >= med * UNIT_OUTLIER_RATIO:
            continue
        keep[k] = by_set[k]
    return keep


def cluster_hits_by_set(rows: list[dict]) -> dict:
    """JAN -> {rakutenBySet, yahooBySet}. Rank-1 is never auto-picked. Cheapest (free-ship first) wins."""
    by_jan: dict[str, dict] = {}
    for rec in latest_per_shop(rows):
        jan = str(rec.get("検索JAN") or "").strip()
        name = str(rec.get("商品名") or "")
        if len(jan) < 8 or exclude_competitor_title(name):
            continue
        set_n, from_p = parse_set_count_from_title(name)
        if set_n is None or from_p:
            continue
        price = effective_price(rec)
        if price is None:
            continue
        paid = 1 if ship_paid(rec) else 0
        mall = str(rec.get("モール") or "")
        if mall.startswith("楽天"):
            bucket = "rakutenBySet"
        elif "Yahoo" in mall:
            bucket = "yahooBySet"
        else:
            continue
        slot = by_jan.setdefault(jan, {"rakutenBySet": {}, "yahooBySet": {}})
        key = str(set_n)
        cand = {"priceIncl": int(round(price)), "url": str(rec.get("商品URL") or ""), "assumedShipping": paid}
        cur = slot[bucket].get(key)
        if not cur or paid < cur["assumedShipping"] or (paid == cur["assumedShipping"] and cand["priceIncl"] < cur["priceIncl"]):
            slot[bucket][key] = cand
    for jan, slot in by_jan.items():
        slot["rakutenBySet"] = filter_outlier_set_unit_prices(slot["rakutenBySet"])
        slot["yahooBySet"] = filter_outlier_set_unit_prices(slot["yahooBySet"])
    return by_jan
