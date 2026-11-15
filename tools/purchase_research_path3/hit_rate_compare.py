#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""JAN-first vs 多段フォールバックのヒット率比較（楽天・Yahoo・Amazon Catalog）。

読取のみ。マスタ非書込。
認証: config.local.json（Rakuten/Yahoo）, spapi_smoke/config.local.json（Amazon）
"""
from __future__ import annotations

import csv
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
SMOKE_DIR = SCRIPT_DIR.parent / "spapi_smoke"
OUT_DIR = SCRIPT_DIR / "out"
RAKUTEN_ICHIBA_API_VERSION = "20260701"
MAX_PRODUCTS = 100
HITS = 10
NAME_THRESHOLD = 50
SLEEP = 0.35


@dataclass
class Product:
    pid: str
    jan: str
    maker: str
    name: str
    is_combo: bool = False
    kind_n: int = 0
    source: str = ""


@dataclass
class QuerySpec:
    keyword: str
    is_jan: bool = False
    label: str = ""


def load_rakuten_cfg() -> dict:
    cfg = {}
    p = SCRIPT_DIR / "config.local.json"
    if p.exists():
        cfg.update(json.loads(p.read_text(encoding="utf-8")))
    for k in ("RAKUTEN_APP_ID", "RAKUTEN_ACCESS_KEY", "YAHOO_SHOPPING_CLIENT_ID"):
        v = __import__("os").environ.get(k, "").strip()
        if v:
            cfg[k] = v
    return cfg


def simplify_name(name: str) -> str:
    s = str(name or "")
    s = re.sub(r"[（）()]", " ", s)
    s = re.sub(r"[・·／/]", " ", s)
    s = re.sub(r"[\s\u3000]+", " ", s).strip()
    return s


def build_jan_first_queries(p: Product) -> list[QuerySpec]:
    if not p.jan:
        return []
    return [QuerySpec(keyword=p.jan, is_jan=True, label="JANのみ")]


def build_fallback_queries(p: Product) -> list[QuerySpec]:
    jan = p.jan.strip()
    name = p.name.strip()
    maker = p.maker.strip()
    simple = simplify_name(name)
    maker_simple = (maker + " " + simple).strip()
    maker_name = (maker + " " + name).strip()
    kind_tag = (str(p.kind_n) + "種") if p.kind_n >= 2 else ""
    seen: set[str] = set()
    out: list[QuerySpec] = []

    def push(kw: str, is_jan: bool, label: str) -> None:
        k = kw.strip()
        if len(k) < 2 or k in seen:
            return
        seen.add(k)
        out.append(QuerySpec(keyword=k, is_jan=is_jan, label=label))

    if p.is_combo:
        if kind_tag and maker_simple:
            push((maker_simple + " " + kind_tag).strip(), False, "maker+short+N種")
        if kind_tag and simple:
            push((simple + " " + kind_tag).strip(), False, "short+N種")
        if maker_simple:
            push(maker_simple, False, "maker+short")
        if maker_name and maker_name != name:
            push(maker_name, False, "maker+name")
        if name and name != jan:
            push(name, False, "name")
        if simple and simple != name:
            push(simple, False, "short")
        if jan:
            push(jan, True, "JAN")
    else:
        if jan:
            push(jan, True, "JAN")
        if maker_name and maker_name != jan:
            push(maker_name, False, "maker+name")
        if name and name != jan:
            push(name, False, "name")
        if simple and simple != name:
            push(simple, False, "short")
    return out


def normalize(s: str) -> str:
    return re.sub(r"[\s\u3000]+", "", str(s or "").lower())


def name_match_score(maker: str, expected: str, hit: str) -> int:
    """GAS evalRakutenItemNameMatchScore の簡略版（完全一致・単語・文字）。"""
    if not expected and not maker:
        return 100
    exp = normalize(expected)
    hit_n = normalize(hit)
    mk = normalize(maker)
    if not hit_n:
        return 0
    score = 0
    if exp and exp in hit_n:
        score = max(score, 100)
    if mk and mk in hit_n:
        score = max(score, 40)
    exp_words = [w for w in re.split(r"[\s\u3000]+", expected.strip()) if len(w) >= 2]
    if exp_words:
        word_hits = sum(1 for w in exp_words if normalize(w) in hit_n)
        score = max(score, int(50 * word_hits / len(exp_words)))
    if exp:
        exp_chars = set(c for c in exp if c)
        hit_chars = set(c for c in hit_n if c)
        if exp_chars:
            score = max(score, int(50 * len(exp_chars & hit_chars) / len(exp_chars)))
    return min(score, 100)


def set_listing_relevant(hit_name: str, kind_n: int) -> bool:
    t = str(hit_name or "")
    if re.search(r"セット|ミックス|詰合|詰め合|食べ比べ|アソート|詰め合わせ", t):
        return True
    if kind_n >= 2:
        if str(kind_n) + "種" in t:
            return True
        if kind_n == 2 and re.search(r"二種|２種", t):
            return True
        if kind_n == 3 and "三種" in t:
            return True
    return False


def relevant_hit(maker: str, name: str, hit_title: str, is_combo: bool, kind_n: int, jan_query: bool) -> bool:
    th = min(NAME_THRESHOLD, 30) if jan_query else NAME_THRESHOLD
    if is_combo and set_listing_relevant(hit_title, kind_n):
        return True
    return name_match_score(maker, name, hit_title) >= th


def http_get(url: str, headers: dict | None = None, timeout: int = 60) -> tuple[int, str]:
    req = urllib.request.Request(url, headers=headers or {"User-Agent": "gas-project-hit-rate/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as res:
            return res.status, res.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", errors="replace")


def rakuten_search(app_id: str, access_key: str, keyword: str) -> tuple[str, list[str]]:
    if not app_id or not access_key:
        return "RAKUTEN_APP_ID/ACCESS_KEY missing", []
    q = urllib.parse.urlencode(
        {
            "applicationId": app_id,
            "accessKey": access_key,
            "keyword": keyword,
            "format": "json",
            "formatVersion": "2",
            "hits": str(HITS),
            "page": "1",
            "field": "0",
            "availability": "0",
        }
    )
    url = (
        "https://openapi.rakuten.co.jp/ichibams/api/IchibaItem/Search/"
        + RAKUTEN_ICHIBA_API_VERSION
        + "?"
        + q
    )
    headers = {
        "Authorization": "Bearer " + access_key,
        "Origin": "https://script.google.com",
        "Referer": "https://script.google.com/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
    }
    code, text = http_get(url, headers)
    if code != 200:
        return "HTTP %s %s" % (code, text[:120]), []
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return "JSON error", []
    if data.get("error"):
        return str(data.get("error_description") or data.get("error")), []
    raw = data.get("Items") or data.get("items") or []
    titles = []
    for it in raw[:HITS]:
        if isinstance(it, dict) and "Item" in it:
            it = it["Item"]
        titles.append(str((it or {}).get("itemName") or ""))
    return "", titles


def yahoo_search(client_id: str, keyword: str, use_jan_param: bool) -> tuple[str, list[str]]:
    if not client_id:
        return "YAHOO_SHOPPING_CLIENT_ID missing", []
    qs: dict[str, str] = {"appid": client_id, "results": str(HITS)}
    if use_jan_param:
        qs["jan_code"] = keyword
    else:
        qs["query"] = keyword
    url = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch?" + urllib.parse.urlencode(qs)
    code, body = http_get(url)
    if code != 200:
        return "HTTP %s %s" % (code, body[:120]), []
    try:
        data = json.loads(body)
    except json.JSONDecodeError:
        return "JSON error", []
    hits = data.get("hits") or []
    return "", [str(h.get("name") or "") for h in hits[:HITS]]


def amazon_catalog_search(keywords: str, brand: str = "") -> tuple[str, list[str]]:
    sys.path.insert(0, str(SMOKE_DIR))
    try:
        import requests
        from spapi_smoke import _load_config, _spapi_headers, request_lwa_access_token, resolve_cred
    except ImportError:
        return "spapi_smoke missing", []

    cfg_path = SMOKE_DIR / "config.local.json"
    if not cfg_path.is_file():
        return "spapi config missing", []
    cfg = _load_config(cfg_path)
    token = request_lwa_access_token(
        resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID"),
        resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET"),
        resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN"),
    )
    access = str(token.get("access_token") or "")
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    mp = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    ua = str(cfg.get("user_agent") or "OctasHitRate/1.0")
    params = {
        "marketplaceIds": mp,
        "keywords": keywords,
        "includedData": "summaries",
        "pageSize": "20",
    }
    if brand:
        params["brandNames"] = brand
    url = endpoint.rstrip("/") + "/catalog/2022-04-01/items?" + urllib.parse.urlencode(params)
    try:
        resp = requests.get(url, headers=_spapi_headers(endpoint, access, ua), timeout=60)
    except Exception as e:
        return str(e), []
    if resp.status_code != 200:
        return "HTTP %s %s" % (resp.status_code, resp.text[:120]), []
    body = resp.json()
    titles = []
    for it in body.get("items") or []:
        sums = it.get("summaries") or []
        title = ""
        if sums and isinstance(sums[0], dict):
            title = str(sums[0].get("itemName") or "")
        titles.append(title)
    return "", titles


def load_products_from_sheets(max_n: int) -> list[Product]:
    sys.path.insert(0, str(SCRIPT_DIR.parent / "competitor_store"))
    from client import sheets_service

    svc = sheets_service(write=False, interactive=False)
    if not svc:
        return []
    sid = "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"
    master = svc.spreadsheets().values().get(
        spreadsheetId=sid, range="'▼商品マスタ(人間作業用)'"
    ).execute().get("values") or []
    ai = svc.spreadsheets().values().get(
        spreadsheetId=sid, range="'AI情報取得data'"
    ).execute().get("values") or []

    hdr_idx = -1
    for i, row in enumerate(master[:25]):
        if "ASINコード" in row:
            hdr_idx = i
            break
    if hdr_idx < 0:
        return []
    hdr = master[hdr_idx]
    col = {h: i for i, h in enumerate(hdr) if h}

    ai_hdr = ai[0] if ai else []
    ai_col = {h: i for i, h in enumerate(ai_hdr) if h}
    for i in range(1, 6):
        alt = "セット構成" + str(i) + "_商品名"
        canon = "構成" + str(i) + "_商品名"
        if canon not in ai_col and alt in ai_col:
            ai_col[canon] = ai_col[alt]

    def ai_kinds(ai_row: list) -> int:
        seen: set[str] = set()
        n = 0
        for i in range(1, 6):
            idx = ai_col.get("構成" + str(i) + "_商品名")
            if idx is None:
                continue
            nm = simplify_name(ai_row[idx] if idx < len(ai_row) else "")
            if nm and nm not in seen:
                seen.add(nm)
                n += 1
        return n

    out: list[Product] = []
    seen_jan: set[str] = set()
    for r in range(hdr_idx + 1, len(master)):
        row = master[r]
        child = str(row[col.get("子SKU", -1)] if col.get("子SKU") is not None and col.get("子SKU") < len(row) else "").strip()
        if child:
            continue
        jan = str(row[col["JANコード"]] if "JANコード" in col and col["JANコード"] < len(row) else "").strip()
        if len(jan) < 8 or jan in seen_jan:
            continue
        maker = str(row[col["メーカー名"]] if "メーカー名" in col and col["メーカー名"] < len(row) else "").strip()
        name = ""
        for nk in ("商品名ベース", "商品名", "商品名amazon"):
            if nk in col and col[nk] < len(row):
                name = str(row[col[nk]] or "").strip()
                if name:
                    break
        if not name:
            continue
        is_combo = bool(str(row[col.get("セット構成1_親SKU", -1)] if col.get("セット構成1_親SKU") is not None and col.get("セット構成1_親SKU") < len(row) else "").strip())
        kind_n = 0
        if is_combo and ai:
            m_asin = str(row[col.get("ASINコード", -1)] if col.get("ASINコード") is not None and col.get("ASINコード") < len(row) else "").strip()
            m_jan = jan
            for ar in ai[1:]:
                ai_j = str(ar[ai_col["JANコード"]] if "JANコード" in ai_col and ai_col["JANコード"] < len(ar) else "").strip()
                ai_a = str(ar[ai_col["ASINコード"]] if "ASINコード" in ai_col and ai_col["ASINコード"] < len(ar) else "").strip()
                if (m_asin and ai_a == m_asin) or (m_jan and ai_j == m_jan):
                    kind_n = ai_kinds(ar)
                    break
            if kind_n < 2:
                skus = []
                for s in range(1, 6):
                    sk = col.get("セット構成" + str(s) + "_親SKU")
                    if sk is None or sk >= len(row):
                        continue
                    v = str(row[sk] or "").strip()
                    if v:
                        skus.append(v)
                kind_n = len(set(skus))
        seen_jan.add(jan)
        out.append(Product(pid="m%d" % (r + 1), jan=jan, maker=maker, name=name, is_combo=is_combo, kind_n=kind_n, source="master"))
        if len(out) >= max_n:
            break
    return out


def load_products_from_keepa_csv(max_n: int, existing: set[str]) -> list[Product]:
    p = SCRIPT_DIR / "石原水産_Keepaリサーチ用_抜粋.csv"
    if not p.exists():
        return []
    out: list[Product] = []
    with p.open(encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        for row in rd:
            jan = str(row.get("商品コード: EAN") or row.get("EAN") or "").strip()
            if len(jan) < 8 or jan in existing:
                continue
            title = str(row.get("商品名") or "").strip()
            brand = str(row.get("ブランド") or row.get("製造者") or "石原水産").strip()
            if not title:
                continue
            existing.add(jan)
            out.append(Product(pid="k" + jan[-4:], jan=jan, maker=brand, name=title, source="keepa_csv"))
            if len(out) >= max_n:
                break
    return out


def eval_strategy(
    mall: str,
    product: Product,
    specs: list[QuerySpec],
    cfg: dict,
) -> dict[str, Any]:
    app_id = str(cfg.get("RAKUTEN_APP_ID") or "").strip()
    access_key = str(cfg.get("RAKUTEN_ACCESS_KEY") or "").strip()
    yahoo_id = str(cfg.get("YAHOO_SHOPPING_CLIENT_ID") or "").strip()

    raw_ok = False
    rel_ok = False
    used = ""
    err = ""
    hit_n = 0
    for spec in specs:
        time.sleep(SLEEP)
        if mall == "rakuten":
            err, titles = rakuten_search(app_id, access_key, spec.keyword)
        elif mall == "yahoo":
            use_jan = spec.is_jan and bool(re.fullmatch(r"\d{8,}", spec.keyword))
            err, titles = yahoo_search(yahoo_id, spec.keyword, use_jan)
        else:
            err, titles = amazon_catalog_search(spec.keyword, product.maker if spec.label == "maker+name" else "")
        if err and not titles:
            continue
        hit_n = len(titles)
        if hit_n > 0:
            raw_ok = True
            used = spec.label + ":" + spec.keyword[:40]
            for t in titles:
                if relevant_hit(product.maker, product.name, t, product.is_combo, product.kind_n, spec.is_jan):
                    rel_ok = True
                    break
            if rel_ok:
                break
            # JAN-first: 1件でも返れば止める（現行競合ロジック）
            if len(specs) == 1:
                break
            # fallback: raw hit でも relevant 無なら次クエリ
            if not rel_ok:
                continue
        if raw_ok and rel_ok:
            break
    return {"raw": raw_ok, "relevant": rel_ok, "query_used": used, "error": err, "hit_n": hit_n}


def main() -> int:
    cfg = load_rakuten_cfg()
    products = load_products_from_sheets(MAX_PRODUCTS)
    seen = {p.jan for p in products}
    extra = load_products_from_keepa_csv(MAX_PRODUCTS - len(products), seen)
    products.extend(extra)
    products = products[:MAX_PRODUCTS]

    if not products:
        print("商品サンプル0件。Sheets認証またはCSVを確認してください。")
        return 2

    stamp = datetime.now(timezone.utc).astimezone().strftime("%Y%m%d_%H%M%S")
    OUT_DIR.mkdir(exist_ok=True)
    detail_path = OUT_DIR / ("hit_rate_compare_%s.csv" % stamp)
    summary_path = OUT_DIR / ("hit_rate_summary_%s.csv" % stamp)

    detail_rows: list[dict] = []
    stats: dict[str, dict[str, dict[str, int]]] = {
        m: {"jan_first": {"raw": 0, "relevant": 0, "n": 0}, "fallback": {"raw": 0, "relevant": 0, "n": 0}}
        for m in ("rakuten", "yahoo", "amazon")
    }

    print("products=%d (master+keepa)" % len(products))
    for pi, p in enumerate(products):
        print("[%d/%d] %s jan=%s combo=%s kindN=%s" % (pi + 1, len(products), p.name[:40], p.jan, p.is_combo, p.kind_n))
        jan_specs = build_jan_first_queries(p)
        fb_specs = build_fallback_queries(p)
        for mall in ("rakuten", "yahoo", "amazon"):
            for strat, specs in (("jan_first", jan_specs), ("fallback", fb_specs)):
                if not specs:
                    res = {"raw": False, "relevant": False, "query_used": "", "error": "no query", "hit_n": 0}
                else:
                    res = eval_strategy(mall, p, specs, cfg)
                stats[mall][strat]["n"] += 1
                if res["raw"]:
                    stats[mall][strat]["raw"] += 1
                if res["relevant"]:
                    stats[mall][strat]["relevant"] += 1
                detail_rows.append(
                    {
                        "product_id": p.pid,
                        "jan": p.jan,
                        "maker": p.maker,
                        "name": p.name[:80],
                        "is_combo": int(p.is_combo),
                        "kind_n": p.kind_n,
                        "source": p.source,
                        "mall": mall,
                        "strategy": strat,
                        "raw_hit": int(res["raw"]),
                        "relevant_hit": int(res["relevant"]),
                        "hit_count": res["hit_n"],
                        "query_used": res["query_used"],
                        "error": res["error"],
                    }
                )

    fields = list(detail_rows[0].keys())
    with detail_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(detail_rows)

    summary_rows = []
    for mall in ("rakuten", "yahoo", "amazon"):
        n = stats[mall]["jan_first"]["n"] or 1
        for strat in ("jan_first", "fallback"):
            s = stats[mall][strat]
            summary_rows.append(
                {
                    "mall": mall,
                    "strategy": strat,
                    "products": s["n"],
                    "raw_hit_count": s["raw"],
                    "raw_hit_pct": round(100.0 * s["raw"] / n, 1),
                    "relevant_hit_count": s["relevant"],
                    "relevant_hit_pct": round(100.0 * s["relevant"] / n, 1),
                }
            )

    with summary_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(summary_rows[0].keys()))
        w.writeheader()
        w.writerows(summary_rows)

    print("\n=== SUMMARY (products=%d) ===" % len(products))
    for row in summary_rows:
        print(
            "%-8s %-10s raw=%s/%s (%.1f%%) relevant=%s/%s (%.1f%%)"
            % (
                row["mall"],
                row["strategy"],
                row["raw_hit_count"],
                row["products"],
                row["raw_hit_pct"],
                row["relevant_hit_count"],
                row["products"],
                row["relevant_hit_pct"],
            )
        )
    print("detail:", detail_path)
    print("summary:", summary_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
