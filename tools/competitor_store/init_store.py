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
    SHEET_MAKER_SUM,
    SHEET_META,
    SHEET_SELLER,
    SHEETS,
    MAKER_HEADERS,
    MAKER_SUM_HEADERS,
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
    sum_vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range="'" + SHEET_MAKER_SUM + "'!A1:A2")
        .execute()
        .get("values")
        or []
    )
    if len(sum_vals) < 2:
        write_sheet(SHEET_MAKER_SUM, MAKER_SUM_HEADERS, [])
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


CK_HEADER = "調査CK"


def _sheet_gid(meta: dict, title: str):
    for s in meta.get("sheets", []):
        props = s.get("properties") or {}
        if props.get("title") == title:
            return int(props["sheetId"])
    return None


def ensure_leading_ck_google(svc, ss_id: str, sheet_title: str, current: list[str]) -> list[str]:
    """A列に調査CKが無ければ列挿入。既存データは右へ。重複列は作らない。"""
    cur = [str(x).strip() for x in (current or [])]
    if cur and cur[0] == CK_HEADER:
        return cur
    if CK_HEADER in cur:
        return cur
    meta = svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    gid = _sheet_gid(meta, sheet_title)
    if gid is None:
        return cur
    svc.spreadsheets().batchUpdate(
        spreadsheetId=ss_id,
        body={
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": gid,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 1,
                        },
                        "inheritFromBefore": False,
                    }
                }
            ]
        },
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=ss_id,
        range="'" + sheet_title + "'!A1",
        valueInputOption="RAW",
        body={"values": [[CK_HEADER]]},
    ).execute()
    return [CK_HEADER] + cur


def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def apply_ck_ui_google(svc, ss_id: str, sheet_title: str) -> str:
    """調査CK列にチェック枠を付け、レ点行の文字を青にする。既存のTRUEは消さない。"""
    meta = (
        svc.spreadsheets()
        .get(
            spreadsheetId=ss_id,
            fields="sheets(properties(sheetId,title,gridProperties),conditionalFormats)",
        )
        .execute()
    )
    sheet = None
    for s in meta.get("sheets") or []:
        if (s.get("properties") or {}).get("title") == sheet_title:
            sheet = s
            break
    if not sheet:
        return "no_sheet"
    props = sheet.get("properties") or {}
    gid = int(props["sheetId"])
    grid = props.get("gridProperties") or {}
    row_count = int(grid.get("rowCount") or 1000)
    col_count = int(grid.get("columnCount") or 26)
    hdr_got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range="'" + sheet_title + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    hdr = [str(x).strip() for x in (hdr_got[0] if hdr_got else [])]
    try:
        idx = hdr.index(CK_HEADER)
    except ValueError:
        return "no_ck_col"
    # 実データの最終行は B列（sellerId／メーカー）で見る。Aはまだ空が多い
    b_a1 = "'" + sheet_title + "'!B2:B"
    b_vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range=b_a1)
        .execute()
        .get("values")
        or []
    )
    last_data = 1
    for i, row in enumerate(b_vals):
        if row and str(row[0]).strip() != "":
            last_data = i + 2
    ck_letter = _col_letter(idx + 1)
    col_a1 = "'" + sheet_title + "'!" + ck_letter + "2:" + ck_letter + str(max(last_data, 2))
    col_vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=ss_id, range=col_a1)
        .execute()
        .get("values")
        or []
    )
    out = []
    need_false = False
    for r in range(2, last_data + 1):
        i = r - 2
        raw = col_vals[i][0] if i < len(col_vals) and col_vals[i] else ""
        s = str(raw).strip().upper()
        if s in ("TRUE", "1", "YES"):
            out.append([True])
        else:
            if s == "":
                need_false = True
            out.append([False])
    if need_false and out:
        pending_false = out
    else:
        pending_false = None
    end_row = min(row_count, max(last_data + 80, 120))
    formula = "=$" + ck_letter + "2=TRUE"
    requests = []
    cfs = sheet.get("conditionalFormats") or []
    for i in range(len(cfs) - 1, -1, -1):
        rule = cfs[i]
        vals = ((rule.get("booleanRule") or {}).get("condition") or {}).get("values") or []
        entered = " ".join(str(v.get("userEnteredValue") or "") for v in vals)
        if formula in entered:
            requests.append({"deleteConditionalFormatRule": {"sheetId": gid, "index": i}})
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": gid,
                    "startRowIndex": 1,
                    "endRowIndex": end_row,
                    "startColumnIndex": idx,
                    "endColumnIndex": idx + 1,
                },
                "cell": {
                    "dataValidation": {
                        "condition": {"type": "BOOLEAN"},
                        "showCustomUi": True,
                        "strict": False,
                    }
                },
                "fields": "dataValidation",
            }
        }
    )
    requests.append(
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": gid,
                            "startRowIndex": 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 0,
                            "endColumnIndex": col_count,
                        }
                    ],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": formula}],
                        },
                        "format": {
                            "textFormat": {
                                "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0}
                            }
                        },
                    },
                },
                "index": 0,
            }
        }
    )
    svc.spreadsheets().batchUpdate(spreadsheetId=ss_id, body={"requests": requests}).execute()
    if pending_false:
        svc.spreadsheets().values().update(
            spreadsheetId=ss_id,
            range="'" + sheet_title + "'!" + ck_letter + "2",
            valueInputOption="USER_ENTERED",
            body={"values": pending_false},
        ).execute()
    return "ok ck_col=%s end_row=%s last_data=%s" % (ck_letter, end_row, last_data)


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
        apply_ck_ui_google(svc, ss_id, SHEET_SELLER)
        return "ok wrote_headers n=%d" % len(SELLER_HEADERS)
    current = ensure_leading_ck_google(svc, ss_id, SHEET_SELLER, current)
    missing = [h for h in SELLER_HEADERS if h not in current]
    ui = apply_ck_ui_google(svc, ss_id, SHEET_SELLER)
    if not missing:
        return "ok exists n=%d ui=%s" % (len(current), ui)

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
    ui2 = apply_ck_ui_google(svc, ss_id, SHEET_SELLER)
    return "ok added=%d ui=%s" % (len(missing), ui2)


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
