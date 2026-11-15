# -*- coding: utf-8 -*-
"""P1 ドライ: メーカー洗い完走の Catalog 列挙（案A+C）。

既定 dry（シート非書）。--apply で調査複製 `①洗いキュー` へ追記。
Keepa／門／最終調査日は書かない（P2）。
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SMOKE = ROOT.parent / "spapi_smoke"
PATH3 = ROOT.parent / "purchase_research_path3"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(SMOKE))
sys.path.insert(0, str(PATH3))

from apply_keepa_full import (  # noqa: E402
    COMPETITOR_SS,
    RESEARCH_SS,
    as_dicts,
    read_all,
    sheets_service,
    utc_now,
)
from schema import SHEET_MAKER  # noqa: E402
from search_catalog_keywords import asins_from_body, search_catalog_items  # noqa: E402
from spapi_smoke import _load_config, request_lwa_access_token, resolve_cred  # noqa: E402

QUEUE_TITLE = "①洗いキュー"
QUEUE_HEADERS = ["メーカー", "入口語", "ASIN", "列挙日時", "Keepa済", "門結果", "メモ"]
SPLIT_THRESHOLD = 500
DEFAULT_MAKER = "尾西食品"


def title_hits(rows: list[dict], needles: list[str]) -> list[dict]:
    out = []
    for r in rows:
        blob = str(r.get("title") or "") + " " + str(r.get("brand") or "")
        if any(n and n in blob for n in needles):
            out.append(r)
    return out


def load_q2(svc, parent: str) -> list[str]:
    try:
        _, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_MAKER))
    except Exception as e:
        print("q2_read_skip", e)
        return []
    for r in rows:
        if str(r.get("メーカー") or "").strip() != parent:
            continue
        raw = str(r.get("第2クエリ語") or "").strip()
        if not raw:
            return []
        return [p.strip() for p in raw.split("|") if p.strip()]
    return []


def catalog_page(access: str, endpoint: str, mp: str, ua: str, keywords: str, token: str = ""):
    resp = search_catalog_items(endpoint, access, mp, ua, keywords, page_size=20, page_token=token or "")
    body = resp.json() if resp.content else {}
    rows = asins_from_body(body)
    nxt = str((body.get("pagination") or {}).get("nextToken") or "").strip()
    return resp.status_code, rows, nxt


def enumerate_query(
    access: str,
    endpoint: str,
    mp: str,
    ua: str,
    keywords: str,
    needles: list[str],
    start_token: str = "",
    max_pages: int = 40,
    stop_at: int = 0,
) -> tuple[list[str], int, bool, str]:
    """Returns (asins unique order, pages, done, last_token).
    stop_at>0 ならユニークがその数以上で打ち切り（done=False）。
    """
    seen: set[str] = set()
    ordered: list[str] = []
    token = start_token
    pages = 0
    while pages < max_pages:
        if stop_at > 0 and len(ordered) >= stop_at:
            return ordered, pages, False, token
        code, rows, nxt = catalog_page(access, endpoint, mp, ua, keywords, token)
        if code != 200:
            raise RuntimeError("Catalog HTTP %s keywords=%s" % (code, keywords))
        pages += 1
        for r in title_hits(rows, needles):
            a = str(r.get("asin") or "").strip().upper()
            if a and a not in seen:
                seen.add(a)
                ordered.append(a)
        token = nxt
        if not token:
            return ordered, pages, True, ""
        time.sleep(0.25)
    return ordered, pages, False, token


def ensure_queue_headers(svc, sid: str) -> None:
    meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
    titles = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if QUEUE_TITLE not in titles:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [{"addSheet": {"properties": {"title": QUEUE_TITLE}}}]},
        ).execute()
    got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'" + QUEUE_TITLE + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    if not got:
        svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range="'" + QUEUE_TITLE + "'!A1",
            valueInputOption="RAW",
            body={"values": [QUEUE_HEADERS]},
        ).execute()


def existing_asins(svc, sid: str, parent: str) -> set[str]:
    try:
        _, rows = as_dicts(read_all(svc, sid, QUEUE_TITLE))
    except Exception:
        return set()
    out = set()
    for r in rows:
        if str(r.get("メーカー") or "").strip() != parent:
            continue
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            out.add(a)
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--maker", default=DEFAULT_MAKER)
    ap.add_argument("--entry", default="")
    ap.add_argument("--apply", action="store_true", help="①洗いキューへ追記")
    ap.add_argument("--max-pages", type=int, default=40)
    args = ap.parse_args()
    parent = str(args.maker).strip()
    entry = str(args.entry or parent).strip()
    needles = [parent]

    cfg = _load_config(SMOKE / "config.local.json")
    access = str(
        request_lwa_access_token(
            resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID"),
            resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET"),
            resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN"),
        ).get("access_token")
        or ""
    )
    if not access:
        print("FAIL no access_token")
        return 1
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    mp = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    ua = str(cfg.get("user_agent") or "PurchaseResearchDry/1.0")

    code0, rows0, nxt0 = catalog_page(access, endpoint, mp, ua, entry, "")
    if code0 != 200:
        print("FAIL page0 HTTP", code0)
        return 2
    hit0 = title_hits(rows0, needles)
    print(
        "probe maker=%s entry=%s page0=%d hit0=%d next=%s"
        % (parent, entry, len(rows0), len(hit0), "Y" if nxt0 else "N")
    )

    svc = sheets_service(write=bool(args.apply), interactive=True)
    q2: list[str] = []
    if svc:
        q2 = load_q2(svc, parent)
    print("q2_words", len(q2), q2[:8])

    # 本クエリを先に列挙。500超で止めて分割語へ。未満なら完了まで（最大200頁）。
    got, pages, done, tok = enumerate_query(
        access,
        endpoint,
        mp,
        ua,
        entry,
        needles,
        "",
        max_pages=200,
        stop_at=SPLIT_THRESHOLD + 1,
    )
    print("enum_main n=%d pages=%d done=%s" % (len(got), pages, done))
    all_asins = list(got)
    pages_total = pages
    mode = "single"

    if len(all_asins) > SPLIT_THRESHOLD:
        mode = "split"
        if not q2:
            print("STATUS 要分割語 uniq=%d" % len(all_asins))
            print("P1_DRY FAIL need_q2")
            return 3
        # 500超分は本クエリを捨てず残しつつ、第2語で追加列挙
        seen = set(all_asins)
        for w in q2:
            q = parent + " " + w
            g2, p2, d2, _t = enumerate_query(
                access, endpoint, mp, ua, q, needles, "", args.max_pages, 0
            )
            pages_total += p2
            for a in g2:
                if a not in seen:
                    seen.add(a)
                    all_asins.append(a)
            print("q2_done word=%s n=%d pages=%d done=%s" % (w, len(g2), p2, d2))
            if not d2:
                print("STATUS 列挙中")
                print("P1_DRY PARTIAL")
                return 4
        print("RESULT mode=%s unique=%d pages=%d" % (mode, len(all_asins), pages_total))
    elif done:
        print("RESULT mode=%s unique=%d pages=%d" % (mode, len(all_asins), pages_total))
    else:
        print("STATUS 列挙中 main uniq=%d (page cap?)" % len(all_asins))
        print("P1_DRY PARTIAL")
        return 4

    if not all_asins:
        print("P1_DRY FAIL title_hit0")
        return 5

    if not args.apply:
        print("dry_ok (no sheet write). sample:", ",".join(all_asins[:5]))
        print("P1_DRY PASS")
        return 0

    if not svc:
        print("no_creds for --apply")
        return 1
    ensure_queue_headers(svc, RESEARCH_SS)
    exist = existing_asins(svc, RESEARCH_SS, parent)
    now = utc_now()
    rows = []
    for a in all_asins:
        if a in exist:
            continue
        rows.append([parent, entry if mode == "single" else "(q2)", a, now, "FALSE", "", "p1_dry"])
    if rows:
        svc.spreadsheets().values().append(
            spreadsheetId=RESEARCH_SS,
            range="'" + QUEUE_TITLE + "'!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": rows},
        ).execute()
    print("apply_ok new=%d skip_exist=%d" % (len(rows), len(all_asins) - len(rows)))
    print("P1_DRY PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
