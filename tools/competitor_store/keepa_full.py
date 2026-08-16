# -*- coding: utf-8 -*-
"""Keepaフル行の組み立て。csv[] は保存しない。stats=90 を列化。"""
from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from typing import Any

IDX_AMAZON = 0
IDX_NEW = 1
IDX_SALES = 3
IDX_LISTPRICE = 4
IDX_COUNT_NEW = 11
IDX_RATING = 16
IDX_COUNT_REVIEWS = 17
IDX_BUY_BOX_SHIPPING = 18
FRESH_DAYS = 90
IMAGE_SEP = "|"  # サブ画像URL。splitしやすい。URLに含まれない


def strip_keepa_csv(product: Any) -> Any:
    if isinstance(product, dict):
        return {k: strip_keepa_csv(v) for k, v in product.items() if k != "csv"}
    if isinstance(product, list):
        return [strip_keepa_csv(x) for x in product]
    return product


def price_fingerprint(product: dict) -> str:
    stats = product.get("stats") or {}
    cur = stats.get("current") or []
    amazon = cur[0] if len(cur) > 0 else ""
    buybox = cur[18] if len(cur) > 18 else ""
    title = product.get("title") or ""
    return "%s|%s|%s" % (buybox, amazon, title)


def raw_json_for_store(product: dict) -> str:
    return json.dumps(strip_keepa_csv(product), ensure_ascii=False, separators=(",", ":"))


def amazon_direct_label(availability: Any) -> str:
    """0＝本体いま売る→いる。-1＝なし→いない。他は空。"""
    if availability == 0 or availability == "0":
        return "いる"
    if availability == -1 or availability == "-1":
        return "いない"
    return ""


def _oos_amazon(stats: dict, key: str) -> str:
    arr = stats.get(key)
    if not isinstance(arr, list) or not arr:
        return ""
    v = arr[0]
    if v is None or v == -1:
        return ""
    return str(v)


def flatten_from_product(product: dict) -> dict:
    """GETなし。生JSON相当から品番リスト向け列。BBは使わない。"""
    stats = product.get("stats") if isinstance(product.get("stats"), dict) else {}
    cur = stats.get("current") or []
    out = {
        "Amazon直販": amazon_direct_label(product.get("availabilityAmazon")),
        "新品: 現在価格": _keepa_current(cur, IDX_NEW),
        "売れ筋ランキング: 現在": _keepa_current(cur, IDX_SALES),
        "Amazon: 180日在庫切れ%": _oos_amazon(stats, "outOfStockPercentage180"),
    }
    out.update(flatten_keepa_display(product))
    return out


def _rating_str(v: Any) -> str:
    if v is None or v == "" or v == -1:
        return ""
    try:
        n = float(v)
    except (TypeError, ValueError):
        return str(v)
    if n > 5 and n <= 50:
        n = n / 10.0
    if n <= 0:
        return ""
    s = ("%.1f" % n).rstrip("0").rstrip(".")
    return s


def _fba_fee_yen(product: dict) -> str:
    ff = product.get("fbaFees")
    if isinstance(ff, dict):
        v = ff.get("pickAndPackFee")
        if v is None:
            v = ff.get("pickAndPack")
        if v in (None, "", -1):
            return ""
        return str(v)
    if ff in (None, "", -1):
        return ""
    return str(ff)


def _cat_root_tree(product: dict) -> tuple[str, str]:
    tree = product.get("categoryTree") or []
    names = []
    if isinstance(tree, list):
        for n in tree:
            if isinstance(n, dict) and n.get("name"):
                names.append(str(n.get("name")))
            elif n:
                names.append(str(n))
    root = names[0] if names else ""
    joined = " > ".join(names) if names else ""
    return root, joined


def _pack_cm(v: Any) -> str:
    if v in (None, "", -1):
        return ""
    try:
        n = float(v)
    except (TypeError, ValueError):
        return ""
    if n <= 0:
        return ""
    cm = n / 10.0
    s = ("%.1f" % cm).rstrip("0").rstrip(".")
    return s


def _bool_ja(v: Any) -> str:
    if v is True or v == "true" or v == 1 or v == "1":
        return "はい"
    if v is False or v == "false" or v == 0 or v == "0":
        return "いいえ"
    return ""


def flatten_keepa_display(product: dict) -> dict:
    """Keepaフル既存列。offers[] が無いときは BB店名は sellerId のみ。"""
    stats = product.get("stats") if isinstance(product.get("stats"), dict) else {}
    cur = stats.get("current") or []
    root, tree = _cat_root_tree(product)
    rating = _rating_str(_keepa_current(cur, IDX_RATING) or product.get("rating"))
    reviews = _keepa_current(cur, IDX_COUNT_REVIEWS) or (
        "" if product.get("reviewCount") in (None, "", -1) else str(product.get("reviewCount"))
    )
    bb_now = _keepa_current(cur, IDX_BUY_BOX_SHIPPING)
    if not bb_now:
        bp = stats.get("buyBoxPrice")
        if bp not in (None, "", -1, -2):
            bb_now = str(bp)
    seller = str(stats.get("buyBoxSellerId") or "").strip()
    return {
        "Buy Box: 現在価格": bb_now,
        "Buy Box: 30 日平均": _stats_at(stats, "avg30", IDX_BUY_BOX_SHIPPING),
        "Buy Box: 90 日平均": _stats_at(stats, "avg90", IDX_BUY_BOX_SHIPPING),
        "レビュー: 評価": rating,
        "レビュー: 評価件数": reviews,
        "カテゴリ: ルート": root,
        "カテゴリ: ツリー": tree,
        "梱包_L_cm": _pack_cm(product.get("packageLength")),
        "梱包_W_cm": _pack_cm(product.get("packageWidth")),
        "梱包_H_cm": _pack_cm(product.get("packageHeight")),
        "梱包_重量_g": "" if product.get("packageWeight") in (None, "", -1) else str(product.get("packageWeight")),
        "FBA手数料": _fba_fee_yen(product),
        "BuyBoxセラー": seller,
        "BuyBox_FBA": _bool_ja(stats.get("buyBoxIsFBA")),
        "画像": image_main_formula(product),
        "サブ画像": image_subs_joined(product),
    }


def _keepa_current(cur: Any, idx: int) -> str:
    if not isinstance(cur, list) or len(cur) <= idx:
        return ""
    v = cur[idx]
    if v is None or v == -1:
        return ""
    return str(v)


def _stats_at(stats: dict, key: str, idx: int) -> str:
    return _keepa_current(stats.get(key) or [], idx)


def keepa_image_urls(product: dict) -> list[str]:
    """Keepa images[].l → Amazon CDN。imagesCSV は空でも images 配列はある。"""
    out = []
    seen = set()
    arr = product.get("images")
    if isinstance(arr, list):
        for it in arr:
            name = ""
            if isinstance(it, dict):
                name = str(it.get("l") or it.get("m") or "").strip()
            elif it:
                name = str(it).strip()
            if not name:
                continue
            url = name if name.startswith("http") else ("https://m.media-amazon.com/images/I/" + name)
            if url not in seen:
                seen.add(url)
                out.append(url)
    csv = str(product.get("imagesCSV") or product.get("image") or "").strip()
    if csv:
        for part in csv.split(","):
            name = part.strip()
            if not name:
                continue
            url = name if name.startswith("http") else ("https://m.media-amazon.com/images/I/" + name)
            if url not in seen:
                seen.add(url)
                out.append(url)
    return out


def _images(product: dict) -> str:
    urls = keepa_image_urls(product)
    return urls[0] if urls else ""


def image_main_formula(product: dict) -> str:
    """シートでクリックできるメイン。表示は「メイン」。"""
    url = _images(product)
    if not url:
        return ""
    return '=HYPERLINK("%s","メイン")' % url.replace('"', "")


def image_subs_joined(product: dict) -> str:
    urls = keepa_image_urls(product)
    return IMAGE_SEP.join(urls[1:])


def _product_has_gate_stats(product: dict) -> bool:
    if not product:
        return False
    stats = product.get("stats") if isinstance(product.get("stats"), dict) else {}
    for key in ("avg90", "avg30", "avg"):
        arr = stats.get(key)
        if not isinstance(arr, list):
            continue
        for idx in (18, 1, 3):
            if idx >= len(arr):
                continue
            v = arr[idx]
            if v is None or v == -1:
                continue
            try:
                if float(v) >= 0:
                    return True
            except (TypeError, ValueError):
                continue
    return False


def warehouse_get_needed(latest: dict | None, now: datetime | None = None, days: int = FRESH_DAYS) -> bool:
    """行なし・90日超・生JSONの門statsが空なら GET。"""
    if keepa_get_needed(latest, now, days):
        return True
    try:
        p = json.loads(str(latest.get("生JSON") or "{}"))
    except json.JSONDecodeError:
        return True
    return not _product_has_gate_stats(p)


def keepa_get_needed(latest: dict | None, now: datetime | None = None, days: int = FRESH_DAYS) -> bool:
    """90日内の Keepaフル行があれば GET しない（日付のみ）。"""
    from purge import parse_acquired

    if not latest:
        return True
    dt = parse_acquired(latest.get("取得日時"))
    if dt is None:
        return True
    now = now or datetime.now(timezone.utc)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return (now - dt) > timedelta(days=days)


def product_to_full_row(product: dict, fetched_at: str, purpose: str | None = None) -> dict:
    from schema import KEEPA_FULL_HEADERS, PURPOSE_RESEARCH

    stats = product.get("stats") if isinstance(product.get("stats"), dict) else {}
    cur = stats.get("current") or []
    eans = product.get("eanList") or []
    ean = eans[0] if eans else (product.get("ean") or "")
    asin = str(product.get("asin") or "").strip().upper()
    rec = {h: "" for h in KEEPA_FULL_HEADERS}
    rec["取得日時"] = fetched_at
    rec["目的"] = purpose or PURPOSE_RESEARCH
    rec["ASIN"] = asin
    rec["商品コード: EAN"] = str(ean or "")
    rec["商品名"] = str(product.get("title") or "")
    rec["画像"] = image_main_formula(product)
    rec["製造者"] = str(product.get("manufacturer") or "")
    rec["ブランド"] = str(product.get("brand") or "")
    rec["親ASIN"] = str(product.get("parentAsin") or "")
    rec["URL: Amazon"] = ("https://www.amazon.co.jp/dp/" + asin) if asin else ""
    rec["URL: Keepa"] = ("https://keepa.com/#!product/5-" + asin) if asin else ""
    rec["Buy Box: 現在価格"] = _keepa_current(cur, IDX_BUY_BOX_SHIPPING)
    rec["Buy Box: 30 日平均"] = _stats_at(stats, "avg30", IDX_BUY_BOX_SHIPPING)
    rec["Buy Box: 90 日平均"] = _stats_at(stats, "avg90", IDX_BUY_BOX_SHIPPING)
    rec["Amazon: 現在価格"] = _keepa_current(cur, IDX_AMAZON)
    rec["Amazon: 30 日平均"] = _stats_at(stats, "avg30", IDX_AMAZON)
    rec["Amazon: 90 日平均"] = _stats_at(stats, "avg90", IDX_AMAZON)
    rec["新品: 90 日平均"] = _stats_at(stats, "avg90", IDX_NEW)
    rec["参考価格: 90 日平均"] = _stats_at(stats, "avg90", IDX_LISTPRICE)
    rec["売れ筋ランキング: 90 日平均"] = _stats_at(stats, "avg90", IDX_SALES)
    rec["月間売上"] = "" if product.get("monthlySold") in (None, "") else str(product.get("monthlySold"))
    rec["出品者数"] = _keepa_current(cur, IDX_COUNT_NEW)
    rec.update(flatten_from_product(product))
    rec["発売日"] = str(product.get("releaseDate") or product.get("listedSince") or "")
    rec["アイテム数"] = str(product.get("numberOfItems") or "")
    rec["パッケージ数量"] = str(product.get("packageQuantity") or "")
    rec["価格指紋"] = price_fingerprint(product)
    rec["生JSON"] = raw_json_for_store(product)
    return rec


def latest_row_for_asin(rows: list[dict], asin: str) -> dict | None:
    key = str(asin or "").strip().upper()
    hits = [r for r in rows if str(r.get("ASIN") or "").strip().upper() == key]
    return hits[-1] if hits else None


def upsert_keepa_full(store: Any, product: dict, fetched_at: str, purpose: str | None = None) -> str:
    """最新1行の指紋が同じなら非書。変化時のみ追記。csv[]は落とす。目的既定=リサーチ。"""
    from schema import SHEET_KEEPA_FULL

    rec = product_to_full_row(product, fetched_at, purpose=purpose)
    if not rec["ASIN"]:
        return "skip_no_asin"
    latest = latest_row_for_asin(store.rows(SHEET_KEEPA_FULL), rec["ASIN"])
    if latest and str(latest.get("価格指紋") or "") == rec["価格指紋"]:
        return "skip_same_fp"
    store.append(SHEET_KEEPA_FULL, rec)
    return "append"


def plan_upsert_actions(
    existing_rows: list[dict],
    products: list[dict],
    fetched_at: str,
    purpose: str | None = None,
) -> list[str]:
    """シート非書。skip_no_asin / skip_same_fp / append。"""
    rows = list(existing_rows or [])
    actions: list[str] = []
    for product in products:
        rec = product_to_full_row(product, fetched_at, purpose=purpose)
        if not rec["ASIN"]:
            actions.append("skip_no_asin")
            continue
        latest = latest_row_for_asin(rows, rec["ASIN"])
        if latest and str(latest.get("価格指紋") or "") == rec["価格指紋"]:
            actions.append("skip_same_fp")
            continue
        actions.append("append")
        rows.append(rec)
    return actions


def headers_live_like() -> list[str]:
    """現行専用スプシ: 目的が末尾。辞書順（目的がASINの前）ではない。"""
    from schema import KEEPA_FULL_HEADERS

    h = [x for x in KEEPA_FULL_HEADERS if x != "目的"]
    h.append("目的")
    return h


def row_values_for_headers(rec: dict, headers: list[str]) -> list[str]:
    return ["" if rec.get(h) is None else str(rec.get(h, "")) for h in headers]


def classify_keepa_get(
    asins: list[str],
    full_rows: list[dict],
    now: datetime | None = None,
    days: int = FRESH_DAYS,
) -> dict[str, list[str]]:
    """90日内かつ門statsありなら skip_fresh。stats空は need_get。"""
    need: list[str] = []
    skip: list[str] = []
    seen: set[str] = set()
    for raw in asins:
        asin = str(raw or "").strip().upper()
        if len(asin) != 10 or asin in seen:
            continue
        seen.add(asin)
        latest = latest_row_for_asin(full_rows, asin)
        if warehouse_get_needed(latest, now, days=days):
            need.append(asin)
        else:
            skip.append(asin)
    return {"need_get": need, "skip_fresh": skip}


def plan_a_keepa_fetch(
    asins: list[str],
    full_rows: list[dict],
    cache_asins: list[str] | None = None,
    now: datetime | None = None,
    days: int = FRESH_DAYS,
) -> dict[str, list[str]]:
    """マスタキャッシュ優先。90日フルかつ門statsありなら hydrate。"""
    cached = {str(a or "").strip().upper() for a in (cache_asins or []) if str(a or "").strip()}
    hydrate: list[str] = []
    fetch: list[str] = []
    seen: set[str] = set()
    for raw in asins:
        asin = str(raw or "").strip().upper()
        if len(asin) != 10 or asin in seen:
            continue
        seen.add(asin)
        if asin in cached:
            continue
        latest = latest_row_for_asin(full_rows, asin)
        if warehouse_get_needed(latest, now, days=days):
            fetch.append(asin)
        else:
            hydrate.append(asin)
    return {"hydrate": hydrate, "fetch": fetch}
