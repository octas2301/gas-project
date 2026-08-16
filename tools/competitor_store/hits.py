# -*- coding: utf-8 -*-
"""Map mall search hits to Japanese モールヒット rows. Never fill 競合確定価格."""
from __future__ import annotations

import json
from datetime import datetime, timezone

from schema import HITS_HEADERS, PURPOSE_RESEARCH, SHEET_HITS

OWN_SHOP_ID = "octas"


def is_own_octas_hit(*, kind: str, shop_code: str = "", shop_name: str = "", url: str = "") -> bool:
    """自店。商品名に Octas があっても True にしない。"""
    own = OWN_SHOP_ID.lower()
    if str(shop_code or "").strip().lower() == own:
        return True
    n = "".join(str(shop_name or "").split()).lower()
    if "オンラインショップoctas" in n or n == "octas":
        return True
    u = str(url or "").lower()
    if ("/" + own + "/") not in u:
        return False
    return "rakuten" in u or "yahoo" in u or "paypaymall" in u


def hit_dedupe_key(rec: dict) -> str:
    code = str(rec.get("店商品コード") or "").strip()
    if not code:
        code = str(rec.get("商品URL") or "").strip()
    return "\t".join(
        [
            str(rec.get("目的") or "").strip(),
            str(rec.get("モール") or "").strip(),
            str(rec.get("検索JAN") or "").strip(),
            code,
        ]
    )


def hit_fingerprint(rec: dict) -> str:
    def v(k: str) -> str:
        return str(rec.get(k) if rec.get(k) is not None else "").strip()

    return "\t".join(
        [v("表示価格"), v("送料フラグ"), v("楽天ポイント％"), v("Yahooポイント数"), v("商品名")]
    )


def filter_unchanged_hits(existing: list[dict], incoming: list[dict]) -> tuple[list[dict], int]:
    latest: dict[str, str] = {}
    for rec in existing:
        k = hit_dedupe_key(rec)
        if not k.replace("\t", ""):
            continue
        latest[k] = hit_fingerprint(rec)
    out: list[dict] = []
    skipped = 0
    for rec in incoming:
        k = hit_dedupe_key(rec)
        fp = hit_fingerprint(rec)
        if k.replace("\t", "") and latest.get(k) == fp:
            skipped += 1
            continue
        if k.replace("\t", ""):
            latest[k] = fp
        out.append(rec)
    return out, skipped


def utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def hit_row(*, mall: str, jan: str, query: str, rank: int, name: str, price, ship, point, url: str, raw: dict, purpose: str | None = None) -> dict:
    rec = {h: "" for h in HITS_HEADERS}
    rec.update({
        "取得日時": utc_now(),
        "目的": purpose or PURPOSE_RESEARCH,
        "モール": mall,
        "検索JAN": jan or "",
        "商品名": name or "",
        "表示価格": "" if price is None else str(price),
        "送料フラグ": "" if ship is None else str(ship),
        "商品URL": url or "",
        "ヒット順位": str(rank),
        "クエリ": query or "",
        "マップ版": "2026-08-15c",
        "生JSON": json.dumps(raw, ensure_ascii=False)[:40000],
        "競合確定価格": "",
    })
    if str(mall).startswith("楽天"):
        rec["楽天ポイント％"] = "" if point is None else str(point)
    else:
        rec["Yahooポイント数"] = "" if point is None else str(point)
    return rec


def append_local(store, rows: list[dict]) -> int:
    existing = store.rows(SHEET_HITS)
    kept, _skipped = filter_unchanged_hits(existing, rows)
    n = 0
    for rec in kept:
        rec["競合確定価格"] = ""
        if not rec.get("目的"):
            rec["目的"] = PURPOSE_RESEARCH
        store.append(SHEET_HITS, rec)
        n += 1
    return n
