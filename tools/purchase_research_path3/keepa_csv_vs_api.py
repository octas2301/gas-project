#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""人のKeepa DL（列ごと） vs product API。キーは出さない。日本価格は円の整数。"""

from __future__ import annotations

import gzip
import csv
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlencode
from urllib.request import Request, urlopen

SCRIPT_DIR = Path(__file__).resolve().parent
SMOKE_SECRETS = SCRIPT_DIR.parent / "set_main_image" / "secrets"
OUT_DIR = SCRIPT_DIR / "out"

IDX_AMAZON = 0
IDX_NEW = 1
IDX_SALES = 3
IDX_LISTPRICE = 4
IDX_NEW_FBM_SHIPPING = 7
IDX_NEW_FBA = 10
IDX_COUNT_NEW = 11
IDX_RATING = 16
IDX_COUNT_REVIEWS = 17
IDX_BUY_BOX_SHIPPING = 18

ASINS = ("B0FQB34W4B", "B08W56WVHD")


def keepa_key() -> str:
    for env in ("KEEPA_API_KEY", "KEEPA_KEY"):
        v = (os.environ.get(env) or "").strip()
        if v:
            return v
    for name in ("keepa_api_key.txt", "KEEPA_API_KEY.txt"):
        for folder in (SMOKE_SECRETS, SCRIPT_DIR / "secrets"):
            p = folder / name
            if p.is_file():
                t = p.read_text(encoding="utf-8").strip()
                if t and not t.startswith("#"):
                    return t
    return ""


def stats_at(stats: Dict[str, Any], key: str, idx: int) -> Any:
    arr = stats.get(key)
    if not isinstance(arr, list) or idx >= len(arr):
        return None
    return arr[idx]


def is_missing(v: Any) -> bool:
    if v is None or v == "" or v == -1:
        return True
    try:
        if float(v) < 0:
            return True
    except (TypeError, ValueError):
        pass
    return False


def as_int_str(v: Any) -> str:
    if is_missing(v):
        return ""
    try:
        n = float(v)
    except (TypeError, ValueError):
        return str(v).strip()
    return str(int(round(n)))


def as_num_str(v: Any) -> str:
    if is_missing(v):
        return ""
    try:
        n = float(v)
    except (TypeError, ValueError):
        return str(v).strip()
    if n == int(n):
        return str(int(n))
    return ("%s" % n).rstrip("0").rstrip(".")


def rating_str(v: Any) -> str:
    if is_missing(v):
        return ""
    n = float(v)
    if n > 5.5:
        n = n / 10.0
    s = ("%.1f" % n).rstrip("0").rstrip(".")
    return s


def mm_to_cm(v: Any) -> str:
    if is_missing(v):
        return ""
    n = float(v) / 10.0
    s = ("%.1f" % n).rstrip("0").rstrip(".")
    return s


def pct_str(v: Any) -> str:
    if is_missing(v):
        return ""
    n = float(v)
    s = as_num_str(n)
    return s + " %"


def drop_pct(avg_short: Any, avg_long: Any) -> str:
    """CSVの『90日間の下落 %』に近い: (avg90 - avg30) / avg30。"""
    if is_missing(avg_short) or is_missing(avg_long):
        return ""
    a = float(avg_short)
    b = float(avg_long)
    if a == 0:
        return ""
    n = (b - a) / a * 100.0
    return ("%s %%" % int(round(n))) if abs(n - round(n)) < 0.15 else ("%.0f %%" % n)


def parse_human(s: str) -> str:
    t = str(s or "").strip()
    if t in ("-", "–", "—"):
        return ""
    t = t.replace(",", "").replace("％", "%")
    t = re.sub(r"\s+", " ", t)
    return t


def parse_num(s: str) -> Optional[float]:
    t = parse_human(s).replace("%", "").replace(" ", "")
    if t == "":
        return None
    m = re.match(r"^-?\d+(\.\d+)?$", t)
    if not m:
        return None
    return float(t)


def image_ids(s: str) -> List[str]:
    ids = []
    for part in re.split(r"[;,]", str(s or "")):
        part = part.strip()
        if not part:
            continue
        m = re.search(r"/I/([^/.]+)", part)
        ids.append(m.group(1) if m else part.replace(".jpg", "").replace(".png", ""))
    return ids


def seller_id_from_csv(s: str) -> str:
    t = parse_human(s)
    m = re.search(r"/ ([A-Z0-9]{8,})$", t)
    return m.group(1) if m else ""


def yesno(v: Any) -> str:
    if v is True or v == 1 or str(v).lower() in ("true", "yes", "1"):
        return "yes"
    if v is False or v == 0 or str(v).lower() in ("false", "no", "0"):
        return "no"
    return ""


def cat_names(p: Dict[str, Any]) -> Tuple[str, str, str]:
    tree = p.get("categoryTree")
    if not isinstance(tree, list) or not tree:
        return "", "", ""
    names = [str(x.get("name") or "") for x in tree if isinstance(x, dict)]
    names = [n for n in names if n]
    root = names[0] if names else ""
    sub = names[-1] if names else ""
    path = " › ".join(names)
    return root, sub, path


def ean_str(p: Dict[str, Any]) -> str:
    el = p.get("eanList")
    if isinstance(el, list) and el:
        return str(el[0])
    e = p.get("ean")
    if isinstance(e, list) and e:
        return str(e[0])
    return str(e or "")


def images_csv_urls(p: Dict[str, Any]) -> str:
    raw = str(p.get("imagesCSV") or "")
    parts = [x.strip() for x in raw.split(",") if x.strip()]
    urls = []
    for x in parts:
        if x.startswith("http"):
            urls.append(x)
        else:
            if not x.lower().endswith((".jpg", ".png", ".jpeg")):
                x = x + ".jpg"
            urls.append("https://m.media-amazon.com/images/I/" + x)
    return ";".join(urls)


def api_map(p: Dict[str, Any]) -> Dict[str, str]:
    st = p.get("stats") if isinstance(p.get("stats"), dict) else {}
    asin = str(p.get("asin") or "")
    root, sub, path = cat_names(p)
    bb_id = str(st.get("buyBoxSellerId") or "")
    bb_stats = st.get("buyBoxStats") if isinstance(st.get("buyBoxStats"), dict) else {}
    won = ""
    if bb_id and isinstance(bb_stats.get(bb_id), dict):
        pw = bb_stats[bb_id].get("percentageWon")
        if not is_missing(pw):
            won = as_int_str(pw)
    bb_seller_disp = ""
    if bb_id:
        bb_seller_disp = ("%s%% / %s" % (won, bb_id)) if won else bb_id

    fba_ids = st.get("sellerIdsLowestFBA")
    fbm_ids = st.get("sellerIdsLowestFBM")
    fba0 = fba_ids[0] if isinstance(fba_ids, list) and fba_ids else ""
    fbm0 = fbm_ids[0] if isinstance(fbm_ids, list) and fbm_ids else ""

    a30 = stats_at(st, "avg30", IDX_SALES)
    a90 = stats_at(st, "avg90", IDX_SALES)
    a365 = stats_at(st, "avg365", IDX_SALES)
    bb30 = stats_at(st, "avg30", IDX_BUY_BOX_SHIPPING)
    bb90 = stats_at(st, "avg90", IDX_BUY_BOX_SHIPPING)
    n30 = stats_at(st, "avg30", IDX_NEW)
    n90 = stats_at(st, "avg90", IDX_NEW)
    fba30 = stats_at(st, "avg30", IDX_NEW_FBA)
    fba90 = stats_at(st, "avg90", IDX_NEW_FBA)
    fbm30 = stats_at(st, "avg30", IDX_NEW_FBM_SHIPPING)
    fbm90 = stats_at(st, "avg90", IDX_NEW_FBM_SHIPPING)
    lp30 = stats_at(st, "avg30", IDX_LISTPRICE)
    lp90 = stats_at(st, "avg90", IDX_LISTPRICE)
    cn30 = stats_at(st, "avg30", IDX_COUNT_NEW)
    cn90 = stats_at(st, "avg90", IDX_COUNT_NEW)
    am90 = stats_at(st, "avg90", IDX_AMAZON)

    oos90 = stats_at(st, "outOfStockPercentage90", IDX_AMAZON)
    uc = p.get("unitCount") if isinstance(p.get("unitCount"), dict) else {}

    return {
        "画像": images_csv_urls(p),
        "商品名": str(p.get("title") or ""),
        "商品ハイライト": "",  # APIに専用キーなし（今回空同士）
        "売れ筋ランキング: 90 日平均": as_int_str(a90),
        "売れ筋ランキング: 365 日平均": as_int_str(a365),
        "売れ筋ランキング: 90日間の下落 %": drop_pct(a30, a90),
        "売れ筋ランキング: 過去30日間の減少": as_int_str(st.get("salesRankDrops30")),
        "売れ筋ランキング: 過去90日間の減少": as_int_str(st.get("salesRankDrops90")),
        "売れ筋ランキング: 過去180日間の減少": as_int_str(st.get("salesRankDrops180")),
        "売れ筋ランキング: 過去365日間の減少": as_int_str(st.get("salesRankDrops365")),
        "月間売上トレンド: 先月の購入": as_int_str(p.get("monthlySold")),
        "レビュー: 評価": rating_str(p.get("rating") if p.get("rating") not in (None, -1) else stats_at(st, "current", IDX_RATING)),
        "レビュー: 評価件数": as_int_str(p.get("reviewCount") or p.get("reviews") or stats_at(st, "current", IDX_COUNT_REVIEWS)),
        "Buy Box: 30 日平均": as_int_str(bb30),
        "Buy Box: 90 日平均": as_int_str(bb90),
        "Buy Box: 90日間の下落 %": drop_pct(bb30, bb90),
        "Buy Box: Buy Box セラー": bb_seller_disp,
        "Buy Box: FBAです": yesno(st.get("buyBoxIsFBA")),
        "Amazon: 現在価格": as_int_str(stats_at(st, "current", IDX_AMAZON)),
        "Amazon: 90 日平均": as_int_str(am90),
        "Amazon: 365 日平均": as_int_str(stats_at(st, "avg365", IDX_AMAZON)),
        "Amazon: 90日間の下落 %": drop_pct(stats_at(st, "avg30", IDX_AMAZON), am90),
        "Amazon: 在庫": "",  # 画面専用。APIは availabilityAmazon
        "Amazon: 90日間在庫切れ": pct_str(oos90),
        "Amazon: 在庫切れカウント 30 日間": as_int_str(st.get("outOfStockCountAmazon30")),
        "Amazon: 在庫切れカウント 90 日間": as_int_str(st.get("outOfStockCountAmazon90")),
        "新品: 90 日平均": as_int_str(n90),
        "新品: 365 日平均": as_int_str(stats_at(st, "avg365", IDX_NEW)),
        "新品: 90日間の下落 %": drop_pct(n30, n90),
        "新しい、第三者FBA: 90 日平均": as_int_str(fba90),
        "新しい、第三者FBA: 365 日平均": as_int_str(stats_at(st, "avg365", IDX_NEW_FBA)),
        "新しい、第三者FBA: 90日間の下落 %": drop_pct(fba30, fba90),
        "最安の FBA セラー": str(fba0 or ""),
        "新品 第三者 FBM: 90 日平均": as_int_str(fbm90),
        "新品 第三者 FBM: 365 日平均": as_int_str(stats_at(st, "avg365", IDX_NEW_FBM_SHIPPING)),
        "新品 第三者 FBM: 90日間の下落 %": drop_pct(fbm30, fbm90),
        "最安の FBM セラー": str(fbm0 or ""),
        "参考価格: 90 日平均": as_int_str(lp90),
        "参考価格: 365 日平均": as_int_str(stats_at(st, "avg365", IDX_LISTPRICE)),
        "参考価格: 90日間の下落 %": drop_pct(lp30, lp90),
        "新品アイテム数: 90 日平均": as_int_str(cn90),
        "新品アイテム数: 365 日平均": as_int_str(stats_at(st, "avg365", IDX_COUNT_NEW)),
        "新品アイテム数: 90日間の下落 %": drop_pct(cn30, cn90),
        "URL: Amazon": "https://www.amazon.co.jp/dp/%s?psc=1" % asin,
        "URL: Keepa": "https://keepa.com/#!product/5-%s" % asin,
        "カテゴリ: ルート": root,
        "カテゴリ: サブ": sub,
        "カテゴリ: ツリー": path,
        "ASIN": asin,
        "商品コード: EAN": ean_str(p),
        "製造者": str(p.get("manufacturer") or ""),
        "ブランド": str(p.get("brand") or ""),
        "単位の詳細: 単位の価値": as_num_str(uc.get("unitValue")),
        "アイテム数": as_int_str(p.get("numberOfItems")),
        "発売日": as_int_str(p.get("releaseDate") or p.get("listedSince")),
        "パッケージ: 長さ (cm)": mm_to_cm(p.get("packageLength")),
        "パッケージ: 幅 (cm)": mm_to_cm(p.get("packageWidth")),
        "パッケージ: 高さ (cm)": mm_to_cm(p.get("packageHeight")),
        "パッケージ: 重さ (g)": as_int_str(p.get("packageWeight")),
        "パッケージ: 数量": as_int_str(p.get("packageQuantity")),
        "商品: 長さ (cm)": mm_to_cm(p.get("itemLength")),
        "商品: 幅 (cm)": mm_to_cm(p.get("itemWidth")),
        "商品: 高さ (cm)": mm_to_cm(p.get("itemHeight")),
        "商品: 重さ (g)": as_int_str(p.get("itemWeight")),
        "Buy Box: 定期おトク便": "",
        "ワンタイムクーポン: 定期おトク便 %": "",
        "ビジネス割引: パーセンテージ": "",
        "Buy Box: % Amazon 30 日": "",
        "Buy Box: % Amazon 90 日": "",
        "Buy Box: % Amazon 180 日": "",
        "Buy Box: % Amazon 365 日": "",
        "Buy Box: % トップセラー 30 日": "",
        "Buy Box: % トップセラー 90 日": "",
        "Buy Box: % トップセラー 180 日": "",
        "Buy Box: % トップセラー 365 日": "",
        "Buy Box: 勝者数 30 日": "",
        "Buy Box: 勝者数 90 日": "",
        "Buy Box: 勝者数 180 日": "",
        "Buy Box: 勝者数 365 日": "",
        "Buy Box: 標準偏差 30 日": "",
        "Buy Box: 標準偏差 90 日": "",
        "Buy Box: 標準偏差 365 日": "",
        "Buy Box: 変動性 30 日": "",
        "Buy Box: 変動性 90 日": "",
        "Buy Box: 変動性 365 日": "",
        "_buyBoxSellerId": bb_id,
        "_availabilityAmazon": as_int_str(p.get("availabilityAmazon")) if p.get("availabilityAmazon") is not None else "",
    }


def classify(col: str, csv_v: str, api_v: str, extra: Dict[str, str]) -> str:
    c = parse_human(csv_v)
    a = parse_human(api_v)
    if col == "画像":
        cid = image_ids(c)
        aid = image_ids(a)
        if not cid and not aid:
            return "両方空"
        if cid and aid and cid[0] == aid[0]:
            return "一致（先頭画像ID）" if cid == aid else "先頭一致・枚数または順不同差"
        if set(cid) == set(aid) and cid:
            return "一致（集合）"
        return "不一致"
    if col in ("Buy Box: Buy Box セラー", "最安の FBA セラー", "最安の FBM セラー"):
        sid_c = seller_id_from_csv(c) or (c if re.match(r"^[A-Z0-9]{8,}$", c) else "")
        sid_a = extra.get("_buyBoxSellerId", "") if "Buy Box" in col else a
        if col.startswith("最安"):
            sid_a = a
        if sid_c and sid_a and sid_c == sid_a:
            return "一致（セラーID）"
        if not c and not a:
            return "両方空"
        if sid_c and sid_a:
            return "不一致（セラーID）"
        return "不一致（表記。APIはIDのみ／CSVは店名付き）" if (c or a) else "両方空"
    if col.startswith("URL:"):
        if c.replace("https://", "http://") == a.replace("https://", "http://"):
            return "一致"
        if asin_in(c) and asin_in(c) == asin_in(a):
            return "一致（ASIN同じ）"
        return "不一致" if c or a else "両方空"

    cn = parse_num(c)
    an = parse_num(a)
    if c in ("",) and a in ("",):
        return "両方空"
    if cn is not None and an is not None:
        if cn == an:
            return "一致"
        if col.startswith("売れ筋ランキング: ") and "減少" not in col and "%" not in col:
            if abs(cn - an) <= max(20, abs(cn) * 0.01):
                return "ほぼ一致（取得時刻差）"
        if abs(cn - an) <= 1 and ("平均" in col or "価格" in col or "アイテム数" in col):
            return "ほぼ一致（丸め）"
        if "%" in col and abs(cn - an) <= 1:
            return "ほぼ一致（%丸め）"
        return "不一致"
    c_cmp = c.lower().replace(" ", "")
    a_cmp = a.lower().replace(" ", "")
    if c_cmp == a_cmp:
        return "一致"
    if not c and a:
        return "CSV空・APIあり"
    if c and not a:
        return "CSVあり・API空（期間統計は未マップまたは画面専用）"
    return "不一致"


def asin_in(s: str) -> str:
    m = re.search(r"(B0[A-Z0-9]{8})", s.upper())
    return m.group(1) if m else ""


def load_export(path: Path) -> Tuple[List[str], Dict[str, Dict[str, str]]]:
    out: Dict[str, Dict[str, str]] = {}
    headers: List[str] = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        headers = list(r.fieldnames or [])
        for rec in r:
            asin = (rec.get("ASIN") or "").strip().upper()
            if asin:
                out[asin] = rec
    return headers, out


def fetch_products(key: str, asins: List[str]) -> Dict[str, Any]:
    q = urlencode(
        {
            "key": key,
            "domain": "5",
            "asin": ",".join(asins),
            "stats": "365",
            "offers": "20",
        }
    )
    url = "https://api.keepa.com/product?" + q
    req = Request(url, headers={"User-Agent": "OctasKeepaPoc/1.0", "Accept-Encoding": "identity"})
    with urlopen(req, timeout=60) as resp:
        raw = resp.read()
        enc = (resp.headers.get("Content-Encoding") or "").lower()
        if enc == "gzip" or (len(raw) >= 2 and raw[0] == 0x1F and raw[1] == 0x8B):
            raw = gzip.decompress(raw)
    return json.loads(raw.decode("utf-8"))


def main() -> int:
    export_path = Path(sys.argv[1]) if len(sys.argv) > 1 else SCRIPT_DIR / "ishihara_keepa_2.csv"
    if not export_path.is_file():
        print("export missing: %s" % export_path)
        return 2
    headers, export = load_export(export_path)
    key = keepa_key()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    err = ""
    products: List[Dict[str, Any]] = []
    tokens_left = ""
    api_ok = False
    if not key:
        err = "KEEPA_API_KEY なし"
    else:
        try:
            data = fetch_products(key, list(ASINS))
            tokens_left = str(data.get("tokensLeft", ""))
            products = data.get("products") or []
            api_ok = True
        except Exception as e:
            err = str(e)[:300]

    by_asin = {str(p.get("asin") or "").upper(): api_map(p) for p in products}
    rows: List[Dict[str, str]] = []
    skip_internal = ("_buyBoxSellerId", "_availabilityAmazon")

    for asin in ASINS:
        csv_row = export.get(asin) or {}
        api_row = by_asin.get(asin) or {}
        extra = {k: api_row.get(k, "") for k in skip_internal}
        for col in headers:
            cv = csv_row.get(col, "")
            av = api_row.get(col, "")
            verdict = classify(col, cv, av, extra) if api_ok else "API_SKIP"
            rows.append({
                "asin": asin,
                "csv_column": col,
                "csv_value": parse_human(cv)[:200],
                "api_value": parse_human(str(av))[:200],
                "verdict": verdict,
            })

    dest = SCRIPT_DIR / "石原水産_Keepa値一致一覧.csv"
    with dest.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["asin", "csv_column", "csv_value", "api_value", "verdict"])
        w.writeheader()
        w.writerows(rows)

    # 列単位サマリ（2ASINとも同じ判定なら1行）
    summary: List[Dict[str, str]] = []
    for col in headers:
        vs = [r for r in rows if r["csv_column"] == col]
        vset = sorted(set(r["verdict"] for r in vs))
        summary.append({
            "csv_column": col,
            "verdict_asins": " | ".join("%s:%s" % (r["asin"], r["verdict"]) for r in vs),
            "all_same": "Y" if len(vset) == 1 else "N",
            "note": vset[0] if len(vset) == 1 else "ASINで判定が分かれた",
        })
    dest2 = SCRIPT_DIR / "石原水産_Keepa値一致_列サマリ.csv"
    with dest2.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["csv_column", "verdict_asins", "all_same", "note"])
        w.writeheader()
        w.writerows(summary)

    print("export=%s cols=%s asins=%s" % (export_path.name, len(headers), len(export)))
    print("api_ok=%s tokensLeft=%s err=%s" % (api_ok, tokens_left, err))
    print("detail=%s" % dest)
    print("summary=%s" % dest2)
    if api_ok:
        from collections import Counter
        c = Counter(r["verdict"] for r in rows)
        for k, n in c.most_common():
            print("  %s\t%s" % (n, k))
    return 0 if api_ok else 1


if __name__ == "__main__":
    sys.exit(main())
