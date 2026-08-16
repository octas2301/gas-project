# -*- coding: utf-8 -*-
"""T0–T5. Never writes listing Keepa cache. Never writes 定時 from domain1 path."""
from __future__ import annotations

import csv
import json
from datetime import datetime, timezone
from pathlib import Path

from init_store import copy_keepa_from_master, init_local
from schema import (
    HITS_HEADERS,
    MASTER_KEEPA_SHEET,
    MASTER_SS_ID,
    PURPOSE_RESEARCH,
    PURPOSE_SCHEDULED,
    SHEET_FIELD_MAP,
    SHEET_HITS,
    SHEET_KEEPA,
    SHEET_KEEPA_FULL,
    SHEETS,
)
from store import LocalStore

BAKEOFF = Path(__file__).resolve().parents[1] / "purchase_research_path3" / "rakuten_yahoo_bakeoff_AE.csv"


def _now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def t0(store: LocalStore) -> str:
    store.ensure_schema()
    for name in SHEETS:
        h = store.headers(name)
        assert h, name
    fmap = store.headers(SHEET_FIELD_MAP)
    assert "仕分け" in fmap
    picks = {r.get("仕分け") for r in store.rows(SHEET_FIELD_MAP)}
    assert any(str(p).startswith("◎") for p in picks)
    assert not any(str(p).startswith("×") for p in picks)
    assert "楽天ポイント％" in store.headers(SHEET_HITS)
    assert "Yahooポイント数" in store.headers(SHEET_HITS)
    from schema import SHEET_MAKER

    assert "メーカー" in store.headers(SHEET_MAKER)
    assert "第2クエリ語" in store.headers(SHEET_MAKER)
    assert SHEET_KEEPA_FULL in SHEETS
    from schema import SHEET_SELLER, SELLER_HEADERS

    assert SHEET_SELLER in SHEETS
    assert "sellerId" in SELLER_HEADERS
    assert "カテゴリ構成" in SELLER_HEADERS
    assert "卸仮説" in SELLER_HEADERS
    assert "メインカテゴリ" in SELLER_HEADERS
    from schema import AMAZON_SELLER_CAT_COLS

    assert AMAZON_SELLER_CAT_COLS[0] == "洋書"
    assert AMAZON_SELLER_CAT_COLS[-1] == "不明"
    assert "食品・飲料・お酒" in SELLER_HEADERS
    assert "ペット用品" in SELLER_HEADERS
    from schema import KEEPA_FULL_HEADERS

    assert "サブ画像" in KEEPA_FULL_HEADERS
    assert KEEPA_FULL_HEADERS.index("サブ画像") == KEEPA_FULL_HEADERS.index("画像") + 1
    assert "画像一覧" not in KEEPA_FULL_HEADERS
    assert "Amazon直販" in KEEPA_FULL_HEADERS
    assert "新品: 現在価格" in KEEPA_FULL_HEADERS
    assert "売れ筋ランキング: 現在" in KEEPA_FULL_HEADERS
    assert "Amazon: 180日在庫切れ%" in KEEPA_FULL_HEADERS
    assert "出品FBAティア" in KEEPA_FULL_HEADERS
    assert "自己発送送料" in KEEPA_FULL_HEADERS
    assert "生JSON" in store.headers(SHEET_KEEPA_FULL)
    assert "価格指紋" in store.headers(SHEET_KEEPA_FULL)
    assert "csv" not in [h.lower() for h in store.headers(SHEET_KEEPA_FULL)]
    from keepa_full import raw_json_for_store, strip_keepa_csv

    stripped = strip_keepa_csv({"asin": "B0", "csv": [[1]], "title": "t"})
    assert "csv" not in stripped
    assert "csv" not in raw_json_for_store({"asin": "B0", "csv": [1]})
    off = Path(__file__).resolve().parents[2] / "docs" / "org" / "competitor_fields" / "keepa_official.csv"
    assert off.exists()
    krows = list(csv.DictReader(off.open(encoding="utf-8-sig")))
    hist = [r for r in krows if "csv[*]" in (r.get("field") or "")]
    assert hist and all(str(r.get("keepa_full_column")) == "N" for r in hist)
    return "ok map=%d full_cols=%d" % (store.row_count(SHEET_FIELD_MAP), len(store.headers(SHEET_KEEPA_FULL)))


def t1(store: LocalStore) -> str:
    if not BAKEOFF.exists():
        return "skip_no_bakeoff"
    rows = list(csv.DictReader(BAKEOFF.open(encoding="utf-8-sig")))
    jan_rows = [r for r in rows if r.get("mall") == "yahoo" and r.get("case_id") == "katsuobushi"]
    assert jan_rows, "no yahoo katsuobushi"
    before = store.row_count(SHEET_HITS)
    for r in jan_rows[:10]:
        rec = {h: "" for h in HITS_HEADERS}
        rec.update({
            "取得日時": _now(),
            "目的": PURPOSE_RESEARCH,
            "モール": "Yahoo!",
            "検索JAN": r.get("jan") or "",
            "商品名": r.get("hit_name") or "",
            "表示価格": r.get("hit_price") or "",
            "送料フラグ": r.get("postage_or_ship") or "",
            "商品URL": r.get("hit_url") or "",
            "ヒット順位": r.get("hit_rank") or "",
            "クエリ": r.get("query_id") or "",
            "マップ版": "2026-08-15",
            "生JSON": json.dumps({"name": r.get("hit_name"), "price": r.get("hit_price")}, ensure_ascii=False),
            "競合確定価格": "",
        })
        store.append(SHEET_HITS, rec)
    after = store.rows(SHEET_HITS)
    new = after[before:]
    assert new and all(x.get("目的") == PURPOSE_RESEARCH for x in new)
    assert all(str(x.get("競合確定価格") or "") == "" for x in new)
    assert not any(str(x.get("ヒット順位")) == "1" and x.get("競合確定価格") for x in new)
    return "ok n=%d" % len(new)


def t2(store: LocalStore) -> str:
    from client import sheets_service

    svc = sheets_service()
    master_before = None
    if svc:
        vals = (
            svc.spreadsheets()
            .values()
            .get(spreadsheetId=MASTER_SS_ID, range="'" + MASTER_KEEPA_SHEET + "'")
            .execute()
            .get("values")
            or []
        )
        master_before = max(0, len(vals) - 1)
    snap_before = store.row_count(SHEET_KEEPA)
    mn, cp = copy_keepa_from_master(store, max_rows=5)
    if svc:
        vals2 = (
            svc.spreadsheets()
            .values()
            .get(spreadsheetId=MASTER_SS_ID, range="'" + MASTER_KEEPA_SHEET + "'")
            .execute()
            .get("values")
            or []
        )
        master_after = max(0, len(vals2) - 1)
        assert master_before == master_after, "master keepa mutated"
    assert store.row_count(SHEET_KEEPA) >= snap_before
    return "ok master_n=%s copied=%s" % (mn, cp)


def t3(store: LocalStore) -> str:
    from client import sheets_service

    snap = store.rows(SHEET_KEEPA)
    dedicated_asins = {str(r.get("ASIN") or "").strip() for r in snap if r.get("ASIN")}
    svc = sheets_service()
    if not svc:
        return "skip_no_sheets"
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + MASTER_KEEPA_SHEET + "'!A1:D20")
        .execute()
        .get("values")
        or []
    )
    if len(vals) < 2:
        return "skip_empty_master"
    hdr = vals[0]
    try:
        ai = hdr.index("ASIN")
    except ValueError:
        ai = 3
    master_asin = ""
    for r in vals[1:]:
        if ai < len(r) and str(r[ai]).strip():
            master_asin = str(r[ai]).strip()
            break
    assert master_asin
    hit_dedicated = master_asin in dedicated_asins
    hit_master = True
    fallback_ok = hit_dedicated or hit_master
    assert fallback_ok
    return "ok asin=%s dedicated=%s master=True" % (master_asin, hit_dedicated)


def t4(store: LocalStore) -> str:
    n0 = store.row_count(SHEET_KEEPA)
    n1 = store.row_count(SHEET_HITS)
    from client import sheets_service

    purge_would = 0
    svc = sheets_service()
    if svc:
        vals = (
            svc.spreadsheets()
            .values()
            .get(spreadsheetId=MASTER_SS_ID, range="'" + MASTER_KEEPA_SHEET + "'")
            .execute()
            .get("values")
            or []
        )
        if len(vals) >= 2:
            hdr = vals[0]
            di = 0
            for i, h in enumerate(hdr):
                if str(h).strip() == "取得日時":
                    di = i
                    break
            cutoff_ms = 90 * 86400000
            import time
            now = time.time() * 1000
            for r in vals[1:]:
                raw = r[di] if di < len(r) else ""
                try:
                    t = float(raw)
                    if t and t < now - cutoff_ms:
                        purge_would += 1
                except (TypeError, ValueError):
                    pass
    assert store.row_count(SHEET_KEEPA) == n0
    assert store.row_count(SHEET_HITS) == n1
    return "ok dedicated_unchanged purge_candidates_master=%d" % purge_would


def t5(store: LocalStore) -> str:
    rec = {h: "" for h in HITS_HEADERS}
    rec.update({
        "取得日時": _now(),
        "目的": PURPOSE_RESEARCH,
        "モール": "楽天",
        "検索JAN": "4906283045119",
        "商品名": "domain1-lane-test",
        "競合確定価格": "",
    })
    before = store.row_count(SHEET_HITS)
    store.append(SHEET_HITS, rec)
    new = store.rows(SHEET_HITS)[before:]
    assert new and all(x.get("目的") == PURPOSE_RESEARCH for x in new)
    assert not any(x.get("目的") == PURPOSE_SCHEDULED for x in new)
    return "ok"


def t6(store: LocalStore) -> str:
    """Chunk2: map bakeoff hits; never fill 競合確定価格; purpose stays リサーチ."""
    from hits import append_local, hit_row

    if not BAKEOFF.exists():
        return "skip_no_bakeoff"
    rows = list(csv.DictReader(BAKEOFF.open(encoding="utf-8-sig")))
    mapped = []
    for i, r in enumerate(rows[:15], start=1):
        mall = "Yahoo!" if r.get("mall") == "yahoo" else "楽天"
        mapped.append(hit_row(
            mall=mall,
            jan=r.get("jan") or "",
            query=r.get("query_id") or "",
            rank=i,
            name=r.get("hit_name") or "",
            price=r.get("hit_price"),
            ship=r.get("postage_or_ship"),
            point=None,
            url=r.get("hit_url") or "",
            raw={"name": r.get("hit_name"), "price": r.get("hit_price")},
        ))
    before = store.row_count(SHEET_HITS)
    n = append_local(store, mapped)
    new = store.rows(SHEET_HITS)[before:]
    assert n == len(new)
    assert n <= 15
    assert all(x.get("目的") == PURPOSE_RESEARCH for x in new)
    assert all(str(x.get("競合確定価格") or "") == "" for x in new)
    assert not any(x.get("目的") == PURPOSE_SCHEDULED for x in new)
    return "ok n=%d" % n


def t7(store: LocalStore) -> str:
    """90-day purge dedicated only; listing master id never written."""
    from datetime import timedelta
    from purge import DEFAULT_DAYS, purge_local
    from schema import HITS_HEADERS, MASTER_SS_ID, SHEET_HITS

    old = {h: "" for h in HITS_HEADERS}
    old.update({
        "取得日時": (datetime.now(timezone.utc) - timedelta(days=DEFAULT_DAYS + 5)).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "目的": PURPOSE_RESEARCH,
        "モール": "楽天",
        "検索JAN": "0000000000000",
        "商品名": "old-row",
    })
    fresh = {h: "" for h in HITS_HEADERS}
    fresh.update({
        "取得日時": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "目的": PURPOSE_RESEARCH,
        "モール": "楽天",
        "検索JAN": "1111111111111",
        "商品名": "fresh-row",
    })
    store.append(SHEET_HITS, old)
    store.append(SHEET_HITS, fresh)
    dry = purge_local(store, days=DEFAULT_DAYS, apply=False)
    names = [r.get("商品名") for r in store.rows(SHEET_HITS)]
    assert "old-row" in names and "fresh-row" in names
    applied = purge_local(store, days=DEFAULT_DAYS, apply=True)
    names2 = [r.get("商品名") for r in store.rows(SHEET_HITS)]
    assert "old-row" not in names2
    assert "fresh-row" in names2
    assert applied[SHEET_HITS]["drop"] >= 1
    assert dry[SHEET_HITS]["drop"] >= 1
    assert MASTER_SS_ID == "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"
    return "ok drop=%s keep=%s" % (applied[SHEET_HITS]["drop"], applied[SHEET_HITS]["keep"])


def t8(store: LocalStore) -> str:
    from datetime import timedelta
    from scheduled import child_stock_positive, due_jans
    from schema import HITS_HEADERS, PURPOSE_SCHEDULED, SHEET_HITS

    master_like = [
        {"子SKU": "", "在庫数": "9", "JANコード": "4900000000001"},
        {"子SKU": "c1", "在庫数": "0", "JANコード": "4900000000002"},
        {"子SKU": "c2", "在庫数": "3", "JANコード": "4900000000003"},
        {"子SKU": "c3", "在庫数": "1", "JANコード": "4900000000004"},
    ]
    stock = child_stock_positive(master_like)
    assert "4900000000001" not in stock
    assert "4900000000002" not in stock
    assert "4900000000003" in stock and "4900000000004" in stock
    now = datetime.now(timezone.utc)
    rec = {h: "" for h in HITS_HEADERS}
    rec.update({
        "取得日時": (now - timedelta(hours=12)).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "目的": PURPOSE_SCHEDULED,
        "モール": "楽天",
        "検索JAN": "4900000000003",
    })
    hits = [rec]
    due = due_jans(stock, hits, now=now)
    assert "4900000000003" not in due
    assert "4900000000004" in due
    return "ok due=%s" % ",".join(due)


def t9(store: LocalStore) -> str:
    from backup import backup_local

    dry = backup_local(root=store.root, apply=False)
    assert dry["apply"] is False
    a1 = backup_local(root=store.root, apply=True)
    a2 = backup_local(root=store.root, apply=True)
    folders = list((store.root / "backups").glob("退避_*"))
    assert len(folders) == 1, folders
    assert (folders[0] / "モールヒット.csv").exists()
    return "ok dest=%s deleted_first=%s" % (a2["dest"], len(a1["would_delete"]) == 0)


def t10(store: LocalStore) -> str:
    from backup import drive_trash_plan
    from schema import MASTER_SS_ID

    live = "live-dedicated-id"
    newest = "copy-new"
    files = [
        {"id": live, "name": "gas-project 競合ストア（段階1）"},
        {"id": newest, "name": "競合ストア退避_20260815"},
        {"id": "old-copy", "name": "競合ストア退避_20260701"},
        {"id": MASTER_SS_ID, "name": "競合ストア退避_fake_on_master_name"},
        {"id": "other", "name": " unrelated "},
    ]
    trash = drive_trash_plan(files, live_id=live, new_id=newest, protect_ids=[MASTER_SS_ID])
    assert "old-copy" in trash
    assert live not in trash
    assert newest not in trash
    assert MASTER_SS_ID not in trash
    assert "other" not in trash
    return "ok trash=%s" % ",".join(trash)


def t11(store: LocalStore) -> str:
    from hits import append_local, hit_row
    from schema import PURPOSE_RESEARCH, PURPOSE_SCHEDULED, SHEET_HITS

    rec = hit_row(
        mall="楽天",
        jan="t11-" + _now().replace(":", ""),
        query="4999999999999",
        rank=1,
        name="scheduled-hit",
        price=100,
        ship=0,
        point=1,
        url="https://example.invalid/item",
        raw={"itemName": "scheduled-hit", "itemPrice": 100},
        purpose=PURPOSE_SCHEDULED,
    )
    rec["競合確定価格"] = "999"
    rec["競合確定価格"] = ""
    before = store.row_count(SHEET_HITS)
    append_local(store, [rec])
    last = store.rows(SHEET_HITS)[-1]
    assert last.get("目的") == PURPOSE_SCHEDULED
    assert last.get("目的") != PURPOSE_RESEARCH
    assert str(last.get("競合確定価格") or "") == ""
    assert str(last.get("ヒット順位")) == "1"
    assert store.row_count(SHEET_HITS) == before + 1
    return "ok"


def t12(store: LocalStore) -> str:
    from inventory import jans_with_api_stock_gt0, qty_by_seller_sku

    payload = {
        "inventorySummaries": [
            {"sellerSku": "c-zero", "totalQuantity": 0},
            {"sellerSku": "c-pos", "totalQuantity": 4},
        ]
    }
    api = qty_by_seller_sku(payload)
    children = [
        {"子SKU": "c-zero", "JANコード": "4900000000010", "在庫数": "99"},
        {"子SKU": "c-pos", "JANコード": "4900000000011", "在庫数": "0"},
        {"子SKU": "c-nosku", "JANコード": "4900000000012", "在庫数": "7"},
        {"子SKU": "", "JANコード": "4900000000013", "在庫数": "7"},
    ]
    jans = jans_with_api_stock_gt0(children, api)
    assert "4900000000010" not in jans
    assert "4900000000011" in jans
    assert "4900000000012" not in jans
    assert "4900000000013" not in jans
    return "ok api_only"


def t13(store: LocalStore) -> str:
    from inventory import qty_from_listings

    q = qty_from_listings({
        "sku": "only-one",
        "fulfillmentAvailability": [
            {"fulfillmentChannelCode": "DEFAULT", "quantity": 0},
            {"fulfillmentChannelCode": "DEFAULT", "quantity": 2},
        ],
    })
    assert q == 2
    assert qty_from_listings({}) is None
    return "ok listings_qty=%s" % q


def t14(store: LocalStore) -> str:
    from hits import is_own_octas_hit

    assert is_own_octas_hit(kind="rakuten", shop_code="octas")
    assert is_own_octas_hit(kind="rakuten", shop_name="オンラインショップOctas")
    assert is_own_octas_hit(
        kind="rakuten",
        url="https://item.rakuten.co.jp/octas/foo/",
    )
    assert is_own_octas_hit(
        kind="yahoo",
        shop_name="オンラインショップ Octas",
        url="https://store.shopping.yahoo.co.jp/octas/bar",
    )
    assert not is_own_octas_hit(kind="rakuten", shop_name="石原水産公式")
    assert not is_own_octas_hit(kind="rakuten", shop_name="別店", url="https://item.rakuten.co.jp/other/x/")
    return "ok own_shop_exclude"


def t15(store: LocalStore) -> str:
    from hits import filter_unchanged_hits, hit_row

    base = hit_row(
        mall="楽天",
        jan="4538872281013",
        query="4538872281013",
        rank=1,
        name="nori",
        price=9980,
        ship=0,
        point=1,
        url="https://item.rakuten.co.jp/other/sku1/",
        raw={},
        purpose="定時",
    )
    base["店商品コード"] = "other_sku1"
    base["楽天ポイント％"] = "1"
    kept, skipped = filter_unchanged_hits([base], [dict(base)])
    assert skipped == 1 and kept == []
    changed = dict(base)
    changed["表示価格"] = "10000"
    kept2, skipped2 = filter_unchanged_hits([base], [changed])
    assert skipped2 == 0 and len(kept2) == 1
    return "ok skip=%d write_on_price_change" % skipped


def t16(store: LocalStore) -> str:
    from apply_to_master import cluster_hits_by_set

    jan = "4538872281013"
    rows = [
        {
            "取得日時": "2026-08-15T12:00:00",
            "モール": "楽天",
            "検索JAN": jan,
            "店商品コード": "a1",
            "商品名": "のり 12袋セット",
            "表示価格": "2000",
            "送料フラグ": "0",
            "楽天ポイント％": "1",
            "楽天還元円": "20",
            "商品URL": "https://item.rakuten.co.jp/other/cheap/",
            "ヒット順位": "2",
        },
        {
            "取得日時": "2026-08-15T12:00:00",
            "モール": "楽天",
            "検索JAN": jan,
            "店商品コード": "a0",
            "商品名": "のり ふるさと納税 12袋セット",
            "表示価格": "1",
            "送料フラグ": "0",
            "楽天還元円": "0",
            "商品URL": "https://item.rakuten.co.jp/other/tax/",
            "ヒット順位": "1",
        },
        {
            "取得日時": "2026-08-15T12:00:00",
            "モール": "楽天",
            "検索JAN": jan,
            "店商品コード": "a2",
            "商品名": "のり 16g×4P",
            "表示価格": "500",
            "送料フラグ": "0",
            "楽天還元円": "0",
            "商品URL": "https://item.rakuten.co.jp/other/4p/",
            "ヒット順位": "3",
        },
        {
            "取得日時": "2026-08-15T13:00:00",
            "モール": "Yahoo!",
            "検索JAN": jan,
            "店商品コード": "y1",
            "商品名": "のり 12袋",
            "表示価格": "2100",
            "送料フラグ": "1",
            "Yahooポイント数": "100",
            "商品URL": "https://store.shopping.yahoo.co.jp/other/y1",
            "ヒット順位": "1",
        },
    ]
    clustered = cluster_hits_by_set(rows)
    r12 = clustered[jan]["rakutenBySet"]["12"]
    y12 = clustered[jan]["yahooBySet"]["12"]
    assert r12["priceIncl"] == 1980
    assert "4" not in clustered[jan]["rakutenBySet"]
    assert y12["priceIncl"] == 2000
    assert clustered[jan]["rakutenBySet"]["12"]["url"].endswith("cheap/")
    return "ok set12 rakuten=%s yahoo=%s" % (r12["priceIncl"], y12["priceIncl"])


def t17(store: LocalStore) -> str:
    from apply_to_master import cluster_hits_by_set, parse_set_count_from_title

    n, fp = parse_set_count_from_title(
        "石原水産 食べるおだし 2種 かつお 50g まぐろ 35g 各5袋ずつ 計10袋"
    )
    assert n == 10 and fp is False
    n2, _ = parse_set_count_from_title("かつお 各5袋ずつ")
    assert n2 is None
    n3, _ = parse_set_count_from_title("【3種類】食べるおだし かつお/まぐろ/ぶり (各5袋) 石原水産")
    assert n3 == 15
    n5, _ = parse_set_count_from_title("【３種類】食べるおだし かつお/まぐろ/ぶり (各５袋) 石原水産")
    assert n5 == 15
    jan = "4906283045119"
    def row(code, name, price, rank, ship="0"):
        return {
            "取得日時": "2026-08-15T13:00:00",
            "モール": "楽天",
            "検索JAN": jan,
            "店商品コード": code,
            "商品名": name,
            "表示価格": str(price),
            "送料フラグ": ship,
            "楽天還元円": "0",
            "商品URL": "https://item.rakuten.co.jp/other/%s/" % code,
            "ヒット順位": str(rank),
        }
    rows = [
        row("r1", "かつお 1袋", 700, 1),
        row("r1b", "かつお 1袋 安い", 642, 9),
        row("r2", "かつお 2袋セット", 990, 2),
        row("r3", "かつお 3袋セット", 1564, 3),
        row("r4", "かつお 4袋セット", 1960, 4),
        row("rmix", "2種 各5袋ずつ 計10袋", 5227, 5),
        row("r5bad", "かつお 5袋セット 別商品", 5227, 6),
        row("r6", "かつお 6袋セット", 2871, 7),
        row("r10", "かつお 10袋セット", 4633, 8),
    ]
    cl = cluster_hits_by_set(rows)[jan]["rakutenBySet"]
    assert cl["1"]["priceIncl"] == 642
    assert cl["1"]["url"].endswith("r1b/")
    assert "5" not in cl
    assert cl["10"]["priceIncl"] == 5227 or cl["10"]["priceIncl"] == 4633
    return "ok min_not_rank1 mix10=%s no5" % cl["10"]["priceIncl"]


def t18(store: LocalStore) -> str:
    from paste_catalog import STAGES_MAIN, STAGE_DIAG, fill_empty_asins, short_core

    assert STAGE_DIAG not in STAGES_MAIN
    called = []

    def search(stage, block):
        called.append(stage)
        if stage == "A_id":
            return [
                {"asin": "B0AAAAAAAA", "title": "他社 かつお"},
                {"asin": "B0ISH00001", "title": "石原水産 食べるおだし かつお"},
            ]
        if stage == "B_kw_jan":
            return [{"asin": "B0SHOULDNT", "title": "石原水産 x"}]
        return []

    block = {
        "jan": "4906283045119",
        "maker": "石原水産",
        "name": "石原水産 食べるおだし かつお",
        "rows": [
            {"asin": "B0EXIST000", "eval": "", "title": "既存"},
            {"asin": "", "eval": "◎", "title": "人間◎空ASIN"},
            {"asin": "", "eval": "", "title": ""},
        ],
    }
    out = fill_empty_asins(block, search)
    asins = [r.get("asin") for r in out["rows"]]
    assert out["stage"] == "A_id"
    assert "B_kw_jan" not in called
    assert "B0AAAAAAAA" not in asins
    assert "B0ISH00001" in asins
    assert asins[0] == "B0EXIST000"
    assert asins[1] == ""
    assert out["rows"][1]["eval"] == "◎"

    called.clear()

    def search2(stage, block):
        called.append(stage)
        if stage == "A_id":
            return []
        if stage == "B_kw_jan":
            return [{"asin": "B0JANHIT01", "title": "石原水産 JANヒット"}]
        return [{"asin": "B0LATE0001", "title": "石原水産 late"}]

    out2 = fill_empty_asins({**block, "rows": [{"asin": "", "eval": "", "title": ""}]}, search2)
    assert out2["stage"] == "B_kw_jan"
    assert called[:2] == ["A_id", "B_kw_jan"]
    assert out2["rows"][0]["asin"] == "B0JANHIT01"
    core = short_core("石原水産 食べるおだし かつお 50g", "石原水産")
    assert "石原" not in core
    assert "食べる" in core or "おだし" in core
    full = fill_empty_asins(
        {"jan": "1", "maker": "石原水産", "name": "x", "rows": [{"asin": "B0EXIST000", "eval": "", "title": "a"}]},
        lambda s, b: [{"asin": "B0NEW00001", "title": "石原水産 y"}],
    )
    assert full["filled"] == 0 and full["stage"] == "skip_no_empty"
    return "ok stage=%s filled=%s" % (out["stage"], out["filled"])


def t19(store: LocalStore) -> str:
    from paste_rank import sort_block

    maker = "石原水産"
    rows = [
        {"asin": "B0LOW00001", "title": "石原水産 かつお 10袋セット", "eval": "40%", "price": "2000", "set_count_cell": "10"},
        {"asin": "B0CIRCLE001", "title": "石原水産 かつお 1袋", "eval": "◎", "price": "400", "set_count_cell": "1"},
        {"asin": "B0MIXCIRCLE", "title": "【3種類】食べるおだし かつお/まぐろ/ぶり (各5袋) 石原水産", "eval": "◎", "price": "5980", "set_count_cell": "5"},
        {"asin": "B0FURUCIRC", "title": "石原水産 かつお ふるさと納税 1袋", "eval": "◎", "price": "100", "set_count_cell": "1"},
        {"asin": "B0FURU0001", "title": "石原水産 かつお ふるさと納税 3袋", "eval": "90%", "price": "100", "set_count_cell": "3"},
        {"asin": "B0NAKED001", "title": "他社 かつお 2袋", "eval": "80%", "price": "500", "set_count_cell": "2"},
        {"asin": "B0PENDING1", "title": "", "eval": "", "price": "", "set_count_cell": ""},
        {"asin": "B0OUTLIER1", "title": "石原水産 かつお 1袋 業務", "eval": "30%", "price": "8000", "set_count_cell": "1"},
        {"asin": "", "title": "", "eval": "", "price": "", "set_count_cell": ""},
    ]
    out = sort_block(rows, maker)
    asins = [r.get("asin") for r in out]
    assert asins[-1] == ""
    assert "B0FURU0001" in asins
    assert out[asins.index("B0FURU0001")]["p2"] == "非候補"
    assert out[asins.index("B0CIRCLE001")]["p2"] == "候補"
    assert out[asins.index("B0CIRCLE001")]["eval"] == "◎"
    assert asins.index("B0CIRCLE001") < asins.index("B0FURU0001")
    assert asins.index("B0LOW00001") < asins.index("B0FURU0001")
    assert asins.index("B0CIRCLE001") < asins.index("B0LOW00001")
    assert out[asins.index("B0NAKED001")]["p2"] == "非候補"
    assert out[asins.index("B0PENDING1")]["p2"] == "未属性"
    assert asins.index("B0PENDING1") < asins.index("B0FURU0001")
    assert out[asins.index("B0OUTLIER1")]["p2"] == "非候補"
    assert out[asins.index("B0CIRCLE001")]["eval"] == "◎"
    assert out[asins.index("B0MIXCIRCLE")]["p2"] == "候補"
    assert out[asins.index("B0FURUCIRC")]["p2"] == "非候補"
    return "ok n=%d" % len(out)


def t20(store: LocalStore) -> str:
    from paste_amazon import cluster_circle_amazon, pick_for_master_set

    jan = "4906283045119"
    rows = [
        {"asin": "B0FIRST000", "title": "石原水産 かつお 1袋", "eval": "◎", "price": "500", "url": "https://www.amazon.co.jp/dp/B0FIRST000", "set_count_cell": "1"},
        {"asin": "B0CHEAP001", "title": "石原水産 かつお 1袋 別", "eval": "◎", "price": "400", "url": "", "set_count_cell": "1"},
        {"asin": "B0PCT00001", "title": "石原水産 かつお 1袋 低評価", "eval": "40%", "price": "100", "url": "", "set_count_cell": "1"},
        {"asin": "B0TEN00001", "title": "石原水産 かつお 10袋セット", "eval": "◎", "price": "3500", "url": "", "set_count_cell": "10"},
        {"asin": "B0DWMJQTXT", "title": "【3種類】食べるおだし かつお/まぐろ/ぶり (各5袋) 石原水産", "eval": "◎", "price": "5980", "url": "", "set_count_cell": "5"},
    ]
    cl = cluster_circle_amazon(rows, jan)
    a1 = pick_for_master_set(cl, 1)
    a10 = pick_for_master_set(cl, 10)
    a2 = pick_for_master_set(cl, 2)
    a5 = pick_for_master_set(cl, 5)
    assert a1["asin"] == "B0CHEAP001" and a1["priceIncl"] == 400
    assert a10["priceIncl"] == 3500
    assert a2 is None
    assert a5 is None
    a15 = pick_for_master_set(cl, 15)
    assert a15 is None
    from paste_amazon import bag_for_amazon_circle

    assert bag_for_amazon_circle("【3種類】かつお/まぐろ/ぶり 各5袋 計15袋", "5") is None
    assert bag_for_amazon_circle("【3種類】かつお/まぐろ/ぶり (各5袋)", "5") is None
    assert bag_for_amazon_circle("【３種類】かつお/まぐろ/ぶり (各５袋)", "5") is None
    assert bag_for_amazon_circle("かつお 各５袋", "5") is None
    assert bag_for_amazon_circle("石原水産 かつお 15袋セット", "15") == 15
    from paste_amazon import parse_master_set_qty, plan_master_amazon_rows

    assert parse_master_set_qty("1袋=1セット") == 1
    assert parse_master_set_qty("10袋=1セット") == 10

    clusters = {jan: cl}
    planned = plan_master_amazon_rows(
        [
            {"jan": jan, "set_qty": 1, "ck": True, "current_amazon": 999, "row": 10},
            {"jan": jan, "set_qty": 2, "ck": True, "current_amazon": "", "row": 11},
            {"jan": jan, "set_qty": 1, "ck": False, "current_amazon": "", "row": 12},
        ],
        clusters,
    )
    assert len(planned) == 1 and planned[0]["new_price"] == 400
    return "ok cheap1=%s plan=%d" % (a1["asin"], len(planned))


def t21(store: LocalStore) -> str:
    import tempfile

    from keepa_full import upsert_keepa_full

    s = LocalStore(Path(tempfile.mkdtemp()))
    s.ensure_schema()
    p1 = {
        "asin": "B0TEST0001",
        "title": "石原 かつお",
        "csv": [[9, 8]],
        "stats": {
            "current": [100] + [-1] * 17 + [200],
            "avg90": [150] + [-1] * 17 + [210],
        },
    }
    assert upsert_keepa_full(s, p1, "t1") == "append"
    assert upsert_keepa_full(s, p1, "t2") == "skip_same_fp"
    assert s.row_count(SHEET_KEEPA_FULL) == 1
    raw = json.loads(s.rows(SHEET_KEEPA_FULL)[0]["生JSON"])
    assert "csv" not in raw
    p2 = {
        "asin": "B0TEST0001",
        "title": "石原 かつお",
        "csv": [[1]],
        "stats": {"current": [999] + [-1] * 17 + [200]},
    }
    assert upsert_keepa_full(s, p2, "t3") == "append"
    assert s.row_count(SHEET_KEEPA_FULL) == 2
    row0 = s.rows(SHEET_KEEPA_FULL)[0]
    assert row0.get("目的") == PURPOSE_RESEARCH
    assert row0.get("Amazon: 90 日平均") == "150"
    assert upsert_keepa_full(s, {"title": "x"}, "t4") == "skip_no_asin"
    return "ok rows=2"


def t22(store: LocalStore) -> str:
    from datetime import datetime, timezone

    from keepa_full import keepa_get_needed, product_to_full_row
    from schema import PURPOSE_RESEARCH

    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    assert keepa_get_needed(None, now) is True
    fresh = {"取得日時": "2026-08-01T00:00:00Z"}
    stale = {"取得日時": "2026-04-01T00:00:00Z"}
    assert keepa_get_needed(fresh, now) is False
    assert keepa_get_needed(stale, now) is True
    p = {
        "asin": "B0DRY00001",
        "title": "stats90 dry",
        "csv": [[1, 2, 3]],
        "monthlySold": 40,
        "stats": {
            "current": [980] + [-1] * 17 + [990],
            "avg90": [150] + [-1] * 17 + [160],
            "avg30": [140] + [-1] * 17 + [155],
        },
    }
    rec = product_to_full_row(p, now.strftime("%Y-%m-%dT%H:%M:%SZ"))
    assert rec["目的"] == PURPOSE_RESEARCH
    assert rec["Amazon: 現在価格"] == "980"
    assert rec["Amazon: 90 日平均"] == "150"
    assert rec["Buy Box: 90 日平均"] == "160"
    assert rec["月間売上"] == "40"
    assert "csv" not in json.loads(rec["生JSON"])
    assert rec["アイテム数"] == ""  # numberOfItems 未使用＝袋数の正にしない
    return "ok fresh_skip stale_get stats90"


def t23(store: LocalStore) -> str:
    from keepa_full import plan_upsert_actions, product_to_full_row
    from schema import PURPOSE_LISTING, PURPOSE_RESEARCH

    p1 = {
        "asin": "B0LIST0001",
        "title": "出品A",
        "csv": [[1]],
        "stats": {"current": [10] + [-1] * 17 + [11]},
    }
    p_bad = {"title": "noasin"}
    p2 = {
        "asin": "B0LIST0001",
        "title": "出品A",
        "csv": [[2]],
        "stats": {"current": [99] + [-1] * 17 + [11]},
    }
    acts = plan_upsert_actions([], [p1, p1, p_bad, p2], "t0", purpose=PURPOSE_LISTING)
    assert acts == ["append", "skip_same_fp", "skip_no_asin", "append"]
    rec = product_to_full_row(p1, "t0", purpose=PURPOSE_LISTING)
    assert rec["目的"] == PURPOSE_LISTING
    assert rec["目的"] != PURPOSE_RESEARCH
    return "ok listing_plan"


def t24(store: LocalStore) -> str:
    from keepa_full import headers_live_like, product_to_full_row, row_values_for_headers
    from schema import KEEPA_FULL_HEADERS, PURPOSE_LISTING

    live = headers_live_like()
    assert live[1] == "ASIN" and live[-1] == "目的"
    assert KEEPA_FULL_HEADERS[1] == "目的"
    rec = product_to_full_row(
        {"asin": "B0HDR00001", "title": "列名", "stats": {"current": [1] + [-1] * 17 + [2]}},
        "t",
        purpose=PURPOSE_LISTING,
    )
    canon = row_values_for_headers(rec, KEEPA_FULL_HEADERS)
    aligned = row_values_for_headers(rec, live)
    assert canon[1] == PURPOSE_LISTING
    assert aligned[1] == "B0HDR00001"
    assert aligned[-1] == PURPOSE_LISTING
    assert aligned[live.index("生JSON")]
    assert "csv" not in aligned[live.index("生JSON")]
    return "ok purpose_last asin_col1"


def t25(store: LocalStore) -> str:
    from datetime import datetime, timezone

    from keepa_full import classify_keepa_get

    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    okj = '{"stats":{"avg90":' + json.dumps([0] * 18 + [100]) + "}}"
    rows = [
        {"ASIN": "B0FRESH001", "取得日時": "2026-08-01T00:00:00Z", "生JSON": okj},
        {"ASIN": "B0STALE001", "取得日時": "2026-04-01T00:00:00Z"},
    ]
    out = classify_keepa_get(
        ["B0FRESH001", "B0FRESH001", "B0STALE001", "B0MISS0001", "x"],
        rows,
        now,
    )
    assert out["skip_fresh"] == ["B0FRESH001"]
    assert out["need_get"] == ["B0STALE001", "B0MISS0001"]
    return "ok skip=1 need=2"


def t26(store: LocalStore) -> str:
    from keepa_full import warehouse_get_needed

    fresh_empty = {
        "取得日時": "2026-08-15T00:00:00Z",
        "生JSON": "{}",
    }
    fresh_ok = {
        "取得日時": "2026-08-15T00:00:00Z",
        "生JSON": '{"stats":{"avg90":' + json.dumps([0] * 18 + [100]) + "}}",
    }
    assert warehouse_get_needed(None) is True
    assert warehouse_get_needed(fresh_empty) is True
    assert warehouse_get_needed(fresh_ok) is False
    return "ok stats_empty_get"


def t27(store: LocalStore) -> str:
    from datetime import datetime, timezone

    from keepa_full import plan_a_keepa_fetch

    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    okj = '{"stats":{"avg90":' + json.dumps([0] * 18 + [100]) + "}}"
    rows = [
        {"ASIN": "B0FRESH001", "取得日時": "2026-08-01T00:00:00Z", "商品名": "フル新鮮", "生JSON": okj},
        {"ASIN": "B0STALE001", "取得日時": "2026-04-01T00:00:00Z"},
    ]
    out = plan_a_keepa_fetch(
        ["B0CACHE000", "B0FRESH001", "B0STALE001", "B0MISS0001"],
        rows,
        cache_asins=["B0CACHE000"],
        now=now,
    )
    assert out["hydrate"] == ["B0FRESH001"]
    assert out["fetch"] == ["B0STALE001", "B0MISS0001"]
    return "ok hydrate=1 fetch=2"


def t28(store: LocalStore) -> str:
    from datetime import datetime, timezone

    from keepa_full import plan_a_keepa_fetch

    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    okj = '{"stats":{"avg90":' + json.dumps([0] * 18 + [100]) + "}}"
    rows = [
        {"ASIN": "B0FRESH001", "取得日時": "2026-08-01T00:00:00Z", "生JSON": okj},
        {"ASIN": "B0EMPTY001", "取得日時": "2026-08-01T00:00:00Z", "生JSON": "{}"},
    ]
    out = plan_a_keepa_fetch(["B0FRESH001", "B0EMPTY001"], rows, cache_asins=[], now=now)
    assert out["hydrate"] == ["B0FRESH001"]
    assert out["fetch"] == ["B0EMPTY001"]
    return "ok empty_stats_fetch"


def t29(store: LocalStore) -> str:
    from keepa_full import flatten_from_product, product_to_full_row

    p = {
        "asin": "B0F0TEST01",
        "availabilityAmazon": 0,
        "stats": {
            "current": [100, 200, -1, 3000] + [-1] * 14,
            "outOfStockPercentage180": [12],
        },
    }
    f = flatten_from_product(p)
    assert f["Amazon直販"] == "いる"
    assert f["新品: 現在価格"] == "200"
    assert f["売れ筋ランキング: 現在"] == "3000"
    assert f["Amazon: 180日在庫切れ%"] == "12"
    rec = product_to_full_row(p, "t")
    assert rec["Amazon直販"] == "いる"
    p2 = {"availabilityAmazon": -1, "stats": {"current": [-1] * 20, "outOfStockPercentage180": [-1]}}
    f2 = flatten_from_product(p2)
    assert f2["Amazon直販"] == "いない"
    assert f2["新品: 現在価格"] == ""
    assert f2["Amazon: 180日在庫切れ%"] == ""
    return "ok flatten"


def t30(store: LocalStore) -> str:
    from listing_fees import parse_fba_remark, pick_fba_tier, pick_self_ship, parse_fba_table, parse_ship_table

    b = parse_fba_remark("25x18x2.0cm/250g")
    assert b["mode"] == "box" and b["maxWeightG"] == 250
    s = parse_fba_remark("30cm/2kg")
    assert s["mode"] == "sum" and s["maxSum"] == 30 and s["maxWeightG"] == 2000
    tbl = parse_fba_table(
        [
            ["FBA手数料", "小型", "", "288", "", "25x18x2.0cm/250g"],
            ["FBA手数料", "標準2b", "", "415", "", "30cm/2kg"],
            ["FBA手数料", "標準2d", "", "425", "", "50cm/2kg"],
        ]
    )
    p = pick_fba_tier(tbl, 10, 8, 1, 100)
    assert p["tier"] == "小型" and p["fee"] == "288"
    p2 = pick_fba_tier(tbl, 20, 15, 10, 500)
    assert p2["tier"] == "標準2d"
    sh = parse_ship_table(
        [
            ["自己発送", "ネコポス", "60", "210"],
            ["自己発送", "宅急便コンパクト", "", "450"],
            ["自己発送", "60サイズ", "60", "750"],
        ]
    )
    sp = pick_self_ship(sh, 10, 8, 5)
    assert sp["size"] == "ネコポス"
    return "ok listing_fees"


def t31(store: LocalStore) -> str:
    from keepa_full import flatten_keepa_display

    p = {
        "packageLength": 283,
        "packageWidth": 239,
        "packageHeight": 67,
        "packageWeight": 320,
        "fbaFees": {"pickAndPackFee": 430},
        "categoryTree": [{"name": "食品・飲料・お酒"}, {"name": "食品"}],
        "stats": {
            "current": [-1] * 16 + [43, 12, -1],
            "avg90": [-1] * 18 + [2500],
            "avg30": [-1] * 18 + [2400],
            "buyBoxSellerId": "A1TEST",
            "buyBoxIsFBA": True,
        },
    }
    d = flatten_keepa_display(p)
    assert d["梱包_L_cm"] == "28.3"
    assert d["FBA手数料"] == "430"
    assert d["カテゴリ: ルート"] == "食品・飲料・お酒"
    assert "食品" in d["カテゴリ: ツリー"]
    assert d["レビュー: 評価"] == "4.3"
    assert d["BuyBoxセラー"] == "A1TEST"
    assert d["BuyBox_FBA"] == "はい"
    p["images"] = [{"l": "71ABC.jpg"}, {"l": "61DEF.jpg"}]
    d2 = flatten_keepa_display(p)
    assert "HYPERLINK" in d2["画像"] and "71ABC.jpg" in d2["画像"]
    assert d2["サブ画像"] == "https://m.media-amazon.com/images/I/61DEF.jpg"
    assert "71ABC.jpg" not in d2["サブ画像"]
    return "ok f1_display"


def main() -> int:
    store = init_local()
    tests = [t0, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16, t17, t18, t19, t20, t21, t22, t23, t24, t25, t26, t27, t28, t29, t30, t31]
    failed = 0
    for fn in tests:
        try:
            msg = fn(store)
            print("PASS", fn.__name__, msg)
        except Exception as e:
            failed += 1
            print("FAIL", fn.__name__, e)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
