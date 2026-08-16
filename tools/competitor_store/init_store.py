# -*- coding: utf-8 -*-
"""Seed 項目マップ and empty モールヒット; copy Keepa cache from listing SS (read-only)."""
from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

from schema import (
    FIELD_MAP_HEADERS,
    KEEP_PICKS,
    KEEPA_HEADER_ALIASES,
    KEEPA_SNAP_HEADERS,
    MASTER_KEEPA_SHEET,
    MASTER_SS_ID,
    META_HEADERS,
    SHEET_FIELD_MAP,
    SHEET_HITS,
    SHEET_KEEPA,
    SHEET_KEEPA_FULL,
    SHEET_MAKER,
    SHEET_META,
    SHEET_SELLER,
    SHEETS,
    MAKER_HEADERS,
    SELLER_HEADERS,
)
from store import LocalStore

FIELDS_DIR = Path(__file__).resolve().parents[2] / "docs" / "org" / "competitor_fields"
CROSSWALK = FIELDS_DIR / "logical_crosswalk.csv"
RAKUTEN_OFF = FIELDS_DIR / "rakuten_official.csv"
YAHOO_OFF = FIELDS_DIR / "yahoo_official.csv"
KEEPA_OFF = FIELDS_DIR / "keepa_official.csv"


def _keep_pick(pick: str) -> bool:
    p = (pick or "").strip()
    if p.startswith("×"):
        return False
    return any(p == k or p.startswith(k) for k in KEEP_PICKS)


def _dest_for(scope: str, pick: str, api: str = "") -> str:
    if (api or "").startswith("Keepa"):
        return "Keepaフル"
    if (pick or "").startswith("K"):
        return "Keepaフル"
    if "検索全体" in (scope or ""):
        return "検索メタ"
    return "モールヒット.生JSON"


def seed_field_map(store: LocalStore) -> None:
    today = date.today().isoformat()
    rows = []
    seen = set()

    def add_row(rec: dict) -> None:
        key = (rec.get("ソースAPI"), rec.get("フィールド"), rec.get("論理名"))
        if key in seen:
            return
        seen.add(key)
        rows.append(rec)

    for path, mall_col in ((RAKUTEN_OFF, "楽天"), (YAHOO_OFF, "Yahoo"), (KEEPA_OFF, "Amazon")):
        if not path.exists():
            continue
        for e in csv.DictReader(path.open(encoding="utf-8-sig")):
            pick = (e.get("agent_pick") or "").strip()
            if not _keep_pick(pick):
                continue
            rec = {h: "" for h in FIELD_MAP_HEADERS}
            rec["論理名"] = e.get("name_ja") or e.get("field") or ""
            rec[mall_col] = e.get("field") or ""
            rec["適用開始日"] = today
            rec["変換メモ"] = e.get("note") or ""
            rec["仕分け"] = pick
            rec["優先度"] = e.get("priority") or ""
            rec["取得先"] = _dest_for(e.get("scope") or "", pick, e.get("api") or "")
            rec["ソースAPI"] = e.get("api") or ""
            rec["フィールド"] = e.get("field") or ""
            add_row(rec)

    if CROSSWALK.exists():
        for e in csv.DictReader(CROSSWALK.open(encoding="utf-8-sig")):
            pick = (e.get("agent_pick") or "").strip()
            if not _keep_pick(pick):
                continue
            rec = {h: "" for h in FIELD_MAP_HEADERS}
            rec["論理名"] = e.get("logical_column_ja") or ""
            rec["Amazon"] = e.get("keepa_product") or ""
            rec["楽天"] = e.get("rakuten_ichiba") or ""
            rec["Yahoo"] = e.get("yahoo_itemsearch_v3") or ""
            rec["適用開始日"] = today
            rec["変換メモ"] = ((e.get("same_column") or "") + " " + (e.get("why") or "")).strip()
            rec["仕分け"] = pick
            rec["優先度"] = e.get("priority") or ""
            rec["取得先"] = "Keepaフル" if pick.startswith("K") else "論理（JOIN注意）"
            rec["ソースAPI"] = "logical_crosswalk"
            rec["フィールド"] = e.get("logical_column_ja") or ""
            add_row(rec)

    store.replace(SHEET_FIELD_MAP, FIELD_MAP_HEADERS, rows)


def _map_keepa_row(rec: dict) -> dict:
    out = {h: "" for h in KEEPA_SNAP_HEADERS}
    for src, val in rec.items():
        dst = KEEPA_HEADER_ALIASES.get(src, src)
        if dst in out:
            out[dst] = val
    return out


def copy_keepa_from_master(store: LocalStore, max_rows: int = 20) -> tuple[int, int]:
    from client import sheets_service

    svc = sheets_service()
    master_n = 0
    copied = 0
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
            master_n = len(vals) - 1
            snap = []
            for r in vals[1 : 1 + max_rows]:
                rec = {h: (r[i] if i < len(r) else "") for i, h in enumerate(hdr)}
                snap.append(_map_keepa_row(rec))
                copied += 1
            existing = store.rows(SHEET_KEEPA)
            store.replace(SHEET_KEEPA, KEEPA_SNAP_HEADERS, existing + snap)
    return master_n, copied


def init_local() -> LocalStore:
    store = LocalStore()
    store.ensure_schema()
    seed_field_map(store)
    if store.row_count(SHEET_META) == 0:
        store.append(SHEET_META, {"キー": "マップ版", "値": date.today().isoformat()})
        store.append(SHEET_META, {"キー": "保存先", "値": "local"})
    return store


def try_create_google_spreadsheet() -> str:
    from client import drive_service, load_config

    cfg = load_config()
    if cfg.get("COMPETITOR_SS_ID"):
        return str(cfg["COMPETITOR_SS_ID"])
    drv = drive_service()
    if not drv:
        return ""
    body = {
        "name": "gas-project 競合ストア（段階1）",
        "mimeType": "application/vnd.google-apps.spreadsheet",
    }
    try:
        f = drv.files().create(body=body, fields="id").execute()
        return f.get("id") or ""
    except Exception:
        return ""


def init_google_workbook(ss_id: str) -> str:
    from client import sheets_service
    from schema import HITS_HEADERS

    svc = sheets_service(write=True, interactive=True)
    if not svc:
        return "no_write_creds"
    meta = svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta.get("sheets", [])}
    reqs = []
    for title in SHEETS:
        if title not in existing:
            reqs.append({"addSheet": {"properties": {"title": title}}})
    if reqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=ss_id, body={"requests": reqs}).execute()

    store = LocalStore()
    store.ensure_schema()
    seed_field_map(store)
    if store.row_count(SHEET_META) == 0:
        store.append(SHEET_META, {"キー": "マップ版", "値": date.today().isoformat()})
        store.append(SHEET_META, {"キー": "保存先", "値": ss_id})

    def write_sheet(name: str, hdr: list, rows: list[list]):
        rng = "'" + name + "'!A1"
        values = [hdr] + rows
        svc.spreadsheets().values().update(
            spreadsheetId=ss_id,
            range=rng,
            valueInputOption="RAW",
            body={"values": values},
        ).execute()

    fmap = store.rows(SHEET_FIELD_MAP)
    write_sheet(SHEET_FIELD_MAP, FIELD_MAP_HEADERS, [[r.get(h, "") for h in FIELD_MAP_HEADERS] for r in fmap])
    write_sheet(SHEET_HITS, HITS_HEADERS, [])
    snap_rows = [[r.get(h, "") for h in KEEPA_SNAP_HEADERS] for r in store.rows(SHEET_KEEPA)]
    write_sheet(SHEET_KEEPA, KEEPA_SNAP_HEADERS, snap_rows)
    from schema import KEEPA_FULL_HEADERS

    full_got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range="'" + SHEET_KEEPA_FULL + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    if not full_got:
        write_sheet(SHEET_KEEPA_FULL, KEEPA_FULL_HEADERS, [])
    write_sheet(
        SHEET_META,
        META_HEADERS,
        [
            ["マップ版", date.today().isoformat()],
            ["保存先ID", ss_id],
            ["段階", "1"],
        ],
    )
    maker_rng = "'" + SHEET_MAKER + "'!A1:A2"
    maker_vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range=maker_rng)
        .execute()
        .get("values")
        or []
    )
    if len(maker_vals) < 2:
        write_sheet(SHEET_MAKER, MAKER_HEADERS, [])
    return "ok sheets=" + ",".join(SHEETS)


def update_field_map_google(ss_id: str) -> str:
    """項目マップだけ更新。モールヒット・Keepaは消さない。"""
    from client import sheets_service

    svc = sheets_service(write=True, interactive=True)
    if not svc:
        return "no_write_creds"
    meta = svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if SHEET_FIELD_MAP not in existing:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=ss_id,
            body={"requests": [{"addSheet": {"properties": {"title": SHEET_FIELD_MAP}}}]},
        ).execute()
    store = LocalStore()
    store.ensure_schema()
    seed_field_map(store)
    fmap = store.rows(SHEET_FIELD_MAP)
    values = [FIELD_MAP_HEADERS] + [[r.get(h, "") for h in FIELD_MAP_HEADERS] for r in fmap]
    svc.spreadsheets().values().clear(
        spreadsheetId=ss_id, range="'" + SHEET_FIELD_MAP + "'"
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=ss_id,
        range="'" + SHEET_FIELD_MAP + "'!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()
    return "ok map_rows=%d" % (len(values) - 1)


def ensure_hits_headers_google(ss_id: str) -> str:
    """モールヒットの不足ヘッダーだけ右に足す。既存行は消さない。"""
    from client import sheets_service
    from schema import HITS_HEADERS

    svc = sheets_service(write=True, interactive=True)
    if not svc:
        return "no_write_creds"
    got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range="'" + SHEET_HITS + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    current = [str(x).strip() for x in (got[0] if got else [])]
    if not current:
        svc.spreadsheets().values().update(
            spreadsheetId=ss_id,
            range="'" + SHEET_HITS + "'!A1",
            valueInputOption="RAW",
            body={"values": [HITS_HEADERS]},
        ).execute()
        return "ok wrote_full_headers n=%d" % len(HITS_HEADERS)
    missing = [h for h in HITS_HEADERS if h not in current]
    if not missing:
        return "ok no_missing n=%d" % len(current)

    def col_letter(n: int) -> str:
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    start_col = len(current) + 1
    a1 = "'" + SHEET_HITS + "'!" + col_letter(start_col) + "1"
    svc.spreadsheets().values().update(
        spreadsheetId=ss_id,
        range=a1,
        valueInputOption="RAW",
        body={"values": [missing]},
    ).execute()
    return "ok added=%d" % len(missing)


def ensure_keepa_full_google(ss_id: str) -> str:
    """Keepaフルタブを作り、空ならヘッダーだけ書く。既存データは消さない。"""
    from client import sheets_service
    from schema import KEEPA_FULL_HEADERS

    svc = sheets_service(write=True, interactive=True)
    if not svc:
        return "no_write_creds"
    meta = svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if SHEET_KEEPA_FULL not in existing:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=ss_id,
            body={"requests": [{"addSheet": {"properties": {"title": SHEET_KEEPA_FULL}}}]},
        ).execute()
    got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range="'" + SHEET_KEEPA_FULL + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    current = [str(x).strip() for x in (got[0] if got else [])]
    if not current:
        svc.spreadsheets().values().update(
            spreadsheetId=ss_id,
            range="'" + SHEET_KEEPA_FULL + "'!A1",
            valueInputOption="RAW",
            body={"values": [KEEPA_FULL_HEADERS]},
        ).execute()
        return "ok wrote_headers n=%d" % len(KEEPA_FULL_HEADERS)
    missing = [h for h in KEEPA_FULL_HEADERS if h not in current]
    if not missing:
        return "ok no_missing n=%d" % len(current)

    meta2 = svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    sid = None
    for s in meta2.get("sheets", []):
        if s["properties"]["title"] == SHEET_KEEPA_FULL:
            sid = s["properties"]["sheetId"]
            break
    need_cols = len(current) + len(missing)
    if sid is not None:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=ss_id,
            body={
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {"sheetId": sid, "gridProperties": {"columnCount": max(need_cols, 40)}},
                            "fields": "gridProperties.columnCount",
                        }
                    }
                ]
            },
        ).execute()

    def col_letter(n: int) -> str:
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    start_col = len(current) + 1
    a1 = "'" + SHEET_KEEPA_FULL + "'!" + col_letter(start_col) + "1"
    svc.spreadsheets().values().update(
        spreadsheetId=ss_id,
        range=a1,
        valueInputOption="RAW",
        body={"values": [missing]},
    ).execute()
    added = "ok added=%d" % len(missing)
    placed = place_keepa_full_col_after(ss_id, "サブ画像", "画像")
    return "%s %s" % (added, placed)


def place_keepa_full_col_after(ss_id: str, col: str, after: str) -> str:
    """指定列を after の右隣へ移す。データごと。"""
    from client import sheets_service

    svc = sheets_service(write=True, interactive=True)
    if not svc:
        return "no_write_creds"
    got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range="'" + SHEET_KEEPA_FULL + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    current = [str(x).strip() for x in (got[0] if got else [])]
    if col not in current or after not in current:
        return "skip no_%s" % (col if col not in current else after)
    i_col = current.index(col)
    i_after = current.index(after)
    dest = i_after + 1
    if i_col == dest:
        return "ok already_after"
    meta = svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    sid = None
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == SHEET_KEEPA_FULL:
            sid = s["properties"]["sheetId"]
            break
    if sid is None:
        return "no_sheet"
    svc.spreadsheets().batchUpdate(
        spreadsheetId=ss_id,
        body={
            "requests": [
                {
                    "moveDimension": {
                        "source": {
                            "sheetId": sid,
                            "dimension": "COLUMNS",
                            "startIndex": i_col,
                            "endIndex": i_col + 1,
                        },
                        "destinationIndex": dest,
                    }
                }
            ]
        },
    ).execute()
    return "ok moved %s after %s" % (col, after)


def ensure_seller_google(ss_id: str) -> str:
    """セラータブを作り、空ならヘッダーだけ書く。既存データは消さない。"""
    from client import sheets_service

    svc = sheets_service(write=True, interactive=True)
    if not svc:
        return "no_write_creds"
    meta = svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if SHEET_SELLER not in existing:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=ss_id,
            body={"requests": [{"addSheet": {"properties": {"title": SHEET_SELLER}}}]},
        ).execute()
    got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range="'" + SHEET_SELLER + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    current = [str(x).strip() for x in (got[0] if got else [])]
    if not current:
        svc.spreadsheets().values().update(
            spreadsheetId=ss_id,
            range="'" + SHEET_SELLER + "'!A1",
            valueInputOption="RAW",
            body={"values": [SELLER_HEADERS]},
        ).execute()
        return "ok wrote_headers n=%d" % len(SELLER_HEADERS)
    missing = [h for h in SELLER_HEADERS if h not in current]
    if not missing:
        return "ok exists n=%d" % len(current)

    def col_letter(n: int) -> str:
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    start_col = len(current) + 1
    a1 = "'" + SHEET_SELLER + "'!" + col_letter(start_col) + "1"
    svc.spreadsheets().values().update(
        spreadsheetId=ss_id,
        range=a1,
        valueInputOption="RAW",
        body={"values": [missing]},
    ).execute()
    return "ok added=%d" % len(missing)


if __name__ == "__main__":
    import sys

    st = init_local()
    from client import load_config
    gid = str(load_config().get("COMPETITOR_SS_ID") or "")
    print("local_store", st.root)
    print("field_map_rows", st.row_count(SHEET_FIELD_MAP))
    print("google_ss_id", gid or "(none)")
    if gid and "--full-init" in sys.argv:
        print("google_init", init_google_workbook(gid))
    elif gid:
        print("google_map_only", update_field_map_google(gid))
        print("google_hits_headers", ensure_hits_headers_google(gid))
        print("google_keepa_full", ensure_keepa_full_google(gid))
        print("google_seller", ensure_seller_google(gid))
