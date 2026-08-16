# -*- coding: utf-8 -*-
"""90-day purge for dedicated store only. Never touches listing Keepa取得_キャッシュ."""
from __future__ import annotations

import argparse
from datetime import datetime, timedelta, timezone

from schema import MASTER_KEEPA_SHEET, MASTER_SS_ID, SHEET_HITS, SHEET_KEEPA
from store import LocalStore

PURGE_SHEETS = (SHEET_HITS, SHEET_KEEPA)
DEFAULT_DAYS = 90


def parse_acquired(raw) -> datetime | None:
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime):
        return raw if raw.tzinfo else raw.replace(tzinfo=timezone.utc)
    s = str(raw).strip()
    if not s:
        return None
    try:
        n = float(s)
        if n > 1e11:
            return datetime.fromtimestamp(n / 1000.0, tz=timezone.utc)
        if n > 1e9:
            return datetime.fromtimestamp(n, tz=timezone.utc)
    except (TypeError, ValueError):
        pass
    s2 = s.replace("Z", "")
    for fmt, size in (("%Y-%m-%dT%H:%M:%S", 19), ("%Y-%m-%d", 10)):
        try:
            dt = datetime.strptime(s2[:size], fmt)
            return dt.replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def cutoff(days: int = DEFAULT_DAYS) -> datetime:
    return datetime.now(timezone.utc) - timedelta(days=days)


def split_rows(rows: list[dict], days: int = DEFAULT_DAYS) -> tuple[list[dict], list[dict]]:
    cut = cutoff(days)
    keep, drop = [], []
    for r in rows:
        dt = parse_acquired(r.get("取得日時"))
        if dt is None or dt >= cut:
            keep.append(r)
        else:
            drop.append(r)
    return keep, drop


def purge_local(store: LocalStore, days: int = DEFAULT_DAYS, apply: bool = False) -> dict:
    out = {}
    for name in PURGE_SHEETS:
        headers = store.headers(name)
        keep, drop = split_rows(store.rows(name), days)
        out[name] = {"keep": len(keep), "drop": len(drop)}
        if apply:
            store.replace(name, headers, keep)
    return out


def assert_not_master(ss_id: str) -> None:
    if ss_id == MASTER_SS_ID:
        raise SystemExit("refuse: listing master id")


def main() -> int:
    ap = argparse.ArgumentParser(description="Purge dedicated competitor store rows older than N days. Default dry-run.")
    ap.add_argument("--days", type=int, default=DEFAULT_DAYS)
    ap.add_argument("--apply", action="store_true", help="Actually delete local CSV rows")
    ap.add_argument("--google", action="store_true", help="Also report/apply dedicated Google workbook")
    args = ap.parse_args()
    store = LocalStore()
    store.ensure_schema()
    stats = purge_local(store, days=args.days, apply=args.apply)
    print("local apply=%s" % args.apply, stats)
    print("master_keepa_untouched", MASTER_KEEPA_SHEET, MASTER_SS_ID)
    if args.google:
        from client import load_config, sheets_service
        from schema import KEEPA_SNAP_HEADERS, HITS_HEADERS

        gid = str(load_config().get("COMPETITOR_SS_ID") or "")
        assert_not_master(gid)
        if not gid:
            print("google skip no COMPETITOR_SS_ID")
            return 0
        svc = sheets_service(write=bool(args.apply), interactive=False)
        if not svc:
            print("google skip no creds")
            return 0
        for title, hdr in ((SHEET_HITS, HITS_HEADERS), (SHEET_KEEPA, KEEPA_SNAP_HEADERS)):
            vals = (
                svc.spreadsheets()
                .values()
                .get(spreadsheetId=gid, range="'" + title + "'")
                .execute()
                .get("values")
                or []
            )
            if not vals:
                print("google", title, "empty")
                continue
            headers = vals[0]
            rows = []
            for line in vals[1:]:
                rec = {headers[i]: (line[i] if i < len(line) else "") for i in range(len(headers))}
                rows.append(rec)
            keep, drop = split_rows(rows, args.days)
            print("google", title, "keep", len(keep), "drop", len(drop), "apply", args.apply)
            if args.apply:
                body = [headers] + [[r.get(h, "") for h in headers] for r in keep]
                svc.spreadsheets().values().clear(spreadsheetId=gid, range="'" + title + "'").execute()
                svc.spreadsheets().values().update(
                    spreadsheetId=gid,
                    range="'" + title + "'!A1",
                    valueInputOption="RAW",
                    body={"values": body},
                ).execute()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
