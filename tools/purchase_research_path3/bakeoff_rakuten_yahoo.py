#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""楽天・Yahoo 競合検索ベイクオフ A〜E（読取のみ・マスタ非書込）。

認証: 同ディレクトリ config.local.json または環境変数
  RAKUTEN_APP_ID / RAKUTEN_ACCESS_KEY / YAHOO_SHOPPING_CLIENT_ID
"""
from __future__ import annotations

import csv
import json
import os
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
OUT_DIR = SCRIPT_DIR / "out"
HITS = 10

CASES = [
    {
        "id": "katsuobushi",
        "jan": "4906283045119",
        "maker": "石原水産",
        "ai_name": "食べるおだし（かつお）50ｇ",
        "core": "石原水産 食べるおだし かつお 50g",
    },
    {
        "id": "maguro",
        "jan": "4906283045317",
        "maker": "石原水産",
        "ai_name": "食べるおだし（まぐろ）35ｇ",
        "core": "石原水産 食べるおだし まぐろ 35g",
    },
    {
        "id": "buri",
        "jan": "4906283047410",
        "maker": "石原水産",
        "ai_name": "食べるおだし（ぶり）40ｇ",
        "core": "石原水産 食べるおだし ぶり 40g",
    },
    {
        "id": "set3",
        "jan": "4906283045119",
        "maker": "石原水産",
        "ai_name": "食べるおだし（かつお・まぐろ・ぶり）",
        "core": "石原水産 食べるおだし かつお まぐろ ぶり",
    },
]


def load_cfg() -> dict:
    cfg = {}
    p = SCRIPT_DIR / "config.local.json"
    if p.exists():
        cfg.update(json.loads(p.read_text(encoding="utf-8")))
    for k in ("RAKUTEN_APP_ID", "RAKUTEN_ACCESS_KEY", "YAHOO_SHOPPING_CLIENT_ID"):
        v = os.environ.get(k, "").strip()
        if v:
            cfg[k] = v
    return cfg


def queries(c: dict) -> list[dict]:
    jan = c["jan"]
    name = c["ai_name"]
    maker_name = (c["maker"] + " " + name).strip()
    return [
        {"id": "A", "label": "現行JAN", "text": jan},
        {"id": "B", "label": "JANをキーワード", "text": jan},
        {"id": "C", "label": "AI商品名", "text": name},
        {"id": "D", "label": "メーカー+商品名", "text": maker_name},
        {"id": "E", "label": "短い核", "text": c["core"]},
    ]


def http_get(url: str, headers: dict | None = None, timeout: int = 60) -> tuple[int, str]:
    req = urllib.request.Request(url, headers=headers or {"User-Agent": "gas-project-bakeoff/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as res:
            return res.status, res.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", errors="replace")


def rakuten_search(app_id: str, access_key: str, keyword: str) -> tuple[str, list[dict]]:
    if not app_id:
        return "RAKUTEN_APP_ID missing", []
    if access_key:
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
        url = "https://openapi.rakuten.co.jp/ichibams/api/IchibaItem/Search/20220601?" + q
        headers = {
            "Authorization": "Bearer " + access_key,
            "User-Agent": "Mozilla/5.0",
        }
    else:
        q = urllib.parse.urlencode(
            {
                "applicationId": app_id,
                "keyword": keyword,
                "format": "json",
                "formatVersion": "2",
                "hits": str(HITS),
                "availability": "0",
            }
        )
        url = "https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706?" + q
        headers = {"User-Agent": "Mozilla/5.0"}
    code, text = http_get(url, headers)
    if code != 200:
        return "HTTP %s %s" % (code, text[:180]), []
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return "JSON parse error", []
    if data.get("error"):
        return str(data.get("error_description") or data.get("error")), []
    raw = data.get("Items") or data.get("items") or []
    items = []
    for it in raw[:HITS]:
        if isinstance(it, dict) and "Item" in it and isinstance(it["Item"], dict):
            it = it["Item"]
        price = it.get("itemPrice")
        try:
            price = int(price) if price is not None else ""
        except (TypeError, ValueError):
            price = ""
        items.append(
            {
                "name": str(it.get("itemName") or ""),
                "price": price,
                "url": str(it.get("itemUrl") or ""),
                "extra": it.get("postageFlag", ""),
            }
        )
    return "", items


def yahoo_search(appid: str, text: str, param: str) -> tuple[str, list[dict]]:
    if not appid:
        return "YAHOO_SHOPPING_CLIENT_ID missing", []
    qs = {"appid": appid, "results": str(HITS)}
    if param == "jan_code":
        qs["jan_code"] = text
    else:
        qs["query"] = text
    url = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch?" + urllib.parse.urlencode(qs)
    code, body = http_get(url)
    if code != 200:
        return "HTTP %s %s" % (code, body[:180]), []
    try:
        data = json.loads(body)
    except json.JSONDecodeError:
        return "JSON parse error", []
    hits = data.get("hits") or []
    items = []
    for h in hits[:HITS]:
        price = h.get("price")
        try:
            price = int(price) if price is not None else ""
        except (TypeError, ValueError):
            price = ""
        img = ""
        image = h.get("image") or {}
        if isinstance(image, dict):
            img = str(image.get("medium") or image.get("small") or "")
        ship = ""
        shipping = h.get("shipping") or {}
        if isinstance(shipping, dict):
            ship = shipping.get("code", "")
        items.append(
            {
                "name": str(h.get("name") or ""),
                "price": price,
                "url": str(h.get("url") or ""),
                "extra": ship,
            }
        )
    return "", items


def row(mall, c, q, param, rank, name, price, url, extra, err, n):
    return {
        "mall": mall,
        "case_id": c["id"],
        "jan": c["jan"],
        "maker": c["maker"],
        "ai_name": c["ai_name"],
        "query_id": q["id"],
        "query_label": q["label"],
        "query_text": q["text"],
        "api_param": param,
        "hit_rank": rank,
        "hit_name": name,
        "hit_price": price,
        "hit_url": url,
        "postage_or_ship": extra,
        "error": err,
        "hit_count": n,
    }


def main() -> None:
    cfg = load_cfg()
    app_id = str(cfg.get("RAKUTEN_APP_ID") or "").strip()
    access_key = str(cfg.get("RAKUTEN_ACCESS_KEY") or "").strip()
    yahoo_id = str(cfg.get("YAHOO_SHOPPING_CLIENT_ID") or "").strip()
    rows = []
    for c in CASES:
        for q in queries(c):
            time.sleep(0.4)
            err, items = rakuten_search(app_id, access_key, q["text"])
            param = "keyword"
            if not items:
                rows.append(row("rakuten", c, q, param, 0, "", "", "", "", err, 0))
            else:
                for i, it in enumerate(items, 1):
                    rows.append(row("rakuten", c, q, param, i, it["name"], it["price"], it["url"], it["extra"], err, len(items)))
            time.sleep(0.25)
            if q["id"] == "A":
                yparam = "jan_code"
                yerr, yitems = yahoo_search(yahoo_id, q["text"], "jan_code")
            else:
                yparam = "query"
                yerr, yitems = yahoo_search(yahoo_id, q["text"], "query")
            if not yitems:
                rows.append(row("yahoo", c, q, yparam, 0, "", "", "", "", yerr, 0))
            else:
                for i, it in enumerate(yitems, 1):
                    rows.append(row("yahoo", c, q, yparam, i, it["name"], it["price"], it["url"], it["extra"], yerr, len(yitems)))
    OUT_DIR.mkdir(exist_ok=True)
    stamp = datetime.now(timezone.utc).astimezone().strftime("%Y%m%d_%H%M%S")
    path = SCRIPT_DIR / ("rakuten_yahoo_bakeoff_AE_%s.csv" % stamp)
    path_stable = SCRIPT_DIR / "rakuten_yahoo_bakeoff_AE.csv"
    fields = list(rows[0].keys()) if rows else []
    for dest in (path, path_stable):
        with dest.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(rows)
    print("wrote", path)
    print("wrote", path_stable)
    print("rows", len(rows), "rakuten_app", bool(app_id), "yahoo_app", bool(yahoo_id))


if __name__ == "__main__":
    main()
