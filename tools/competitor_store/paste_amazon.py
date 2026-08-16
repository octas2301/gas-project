# -*- coding: utf-8 -*-
"""P3: 貼付の人間◎だけを JAN＋袋数で束ねる。1件目禁止。マスタ非書。"""
from __future__ import annotations

import re

from apply_to_master import exclude_competitor_title, normalize_fullwidth_digits, parse_set_count_from_title


def jan_digits(v) -> str:
    s = str(v or "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return "".join(ch for ch in s if ch.isdigit())


def _num(v) -> float | None:
    try:
        n = float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None
    return n if n == n and n > 0 else None


_EACH_BAG = re.compile(r"各\s*\d+\s*袋")
_KIND_COUNT = re.compile(r"(?:【)?(\d+)\s*種類?(?:】)?")


def amazon_circle_is_kind_mix(title: str) -> bool:
    """2種以上は単品JANのマスタに載せない（◎でも）。"""
    t = normalize_fullwidth_digits(title)
    m = _KIND_COUNT.search(t)
    return bool(m and int(m.group(1)) >= 2)


def bag_for_amazon_circle(title: str, set_count_cell) -> int | None:
    """タイトル優先。種類ミックスは袋数を付けない。"""
    t = normalize_fullwidth_digits(title)
    if amazon_circle_is_kind_mix(t):
        return None
    n, from_p = parse_set_count_from_title(t)
    sc = _num(set_count_cell)
    if n:
        return int(n)
    if _EACH_BAG.search(t) or from_p:
        return None
    if sc:
        return int(sc)
    return None


def cluster_circle_amazon(rows: list[dict], jan: str) -> dict:
    """rows: asin,title,eval,price,url,set_count_cell. ◎のみ。同袋は最安。"""
    by_set: dict[str, dict] = {}
    for rec in rows:
        if str(rec.get("eval") or "").strip() != "◎":
            continue
        if exclude_competitor_title(str(rec.get("title") or "")):
            continue
        asin = str(rec.get("asin") or "").strip().upper()
        price = _num(rec.get("price"))
        if not asin or not price:
            continue
        bag = bag_for_amazon_circle(str(rec.get("title") or ""), rec.get("set_count_cell"))
        if not bag:
            continue
        key = str(bag)
        prev = by_set.get(key)
        if prev and prev["priceIncl"] <= price:
            continue
        by_set[key] = {
            "priceIncl": int(price) if price == int(price) else price,
            "asin": asin,
            "url": str(rec.get("url") or "") or ("https://www.amazon.co.jp/dp/" + asin),
        }
    return {"jan": jan_digits(jan), "amazonBySet": by_set}


def pick_for_master_set(cluster: dict, set_qty: int) -> dict | None:
    if not set_qty or set_qty < 1:
        return None
    return (cluster or {}).get("amazonBySet", {}).get(str(int(set_qty)))


def parse_master_set_qty(v) -> int | None:
    """'1袋=1セット' → 1。先頭の個数。"""
    s = str(v or "").strip()
    m = re.match(r"(\d+)", s)
    if m:
        n = int(m.group(1))
        return n if n >= 1 else None
    try:
        n = int(float(s.replace(",", "")))
    except ValueError:
        return None
    return n if n >= 1 else None


def checkbox_is_true(v) -> bool:
    if v is True:
        return True
    return str(v).strip().upper() in ("TRUE", "1")


def plan_master_amazon_rows(master_rows: list[dict], clusters: dict) -> list[dict]:
    """master_rows: jan, set_qty, ck, current_amazon, row. 出品CKのみ。空上書きしない。"""
    keyed = {jan_digits(k): v for k, v in (clusters or {}).items()}
    out = []
    for rec in master_rows:
        if not checkbox_is_true(rec.get("ck")):
            continue
        jan = jan_digits(rec.get("jan"))
        set_qty = parse_master_set_qty(rec.get("set_qty"))
        if not set_qty:
            continue
        hit = pick_for_master_set(keyed.get(jan) or {}, set_qty)
        if not hit:
            continue
        out.append(
            {
                "jan": jan,
                "set_qty": set_qty,
                "row": rec.get("row"),
                "current": rec.get("current_amazon"),
                "new_price": hit["priceIncl"],
                "new_asin": hit["asin"],
                "new_url": hit["url"],
            }
        )
    return out
