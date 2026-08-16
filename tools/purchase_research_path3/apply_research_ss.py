#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""調査複製へ門を書く。メーカーマスタは競合DB。出品マスタは触らない。"""
from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime, timezone
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR.parent / "competitor_store"))
from client import sheets_service  # noqa: E402
from schema import MAKER_HEADERS, SHEET_MAKER  # noqa: E402

RESEARCH_SS = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
OAUTH_NEED = "contact@octas2301.com"

T_CAND = "①候補"
T_SUM = "①サマリ"
T_PROF = "①プロファイル"
T_LOG = "①ログ"


def utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def seed_makers() -> list[dict]:
    now = utc_now()
    return [
        {
            "メーカー": "五木食品",
            "第2クエリ語": "ラーメン|うどん|そば|そうめん|スパゲティ|ちゃんぽん",
            "採取元": "official",
            "公式URL": "https://www.itsukifoods.jp/product_list.html",
            "取得日時": now,
            "再採取しない": "TRUE",
            "メモ": "公式ナビ6語。焼そばはちゃんぽんページ",
        },
        {
            "メーカー": "石原水産",
            "第2クエリ語": "",
            "採取元": "",
            "公式URL": "",
            "取得日時": "",
            "再採取しない": "FALSE",
            "メモ": "第2クエリ語が空なので次回採取対象。公式未確認",
        },
    ]


def load_gate_rows() -> list[list[str]]:
    path = SCRIPT_DIR / "out" / "ishihara_gate_paste.csv"
    rows: list[list[str]] = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for line in csv.reader(f):
            rows.append(line)
    return rows


def ensure_sheet(svc, sid: str, title: str) -> None:
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
    titles = [(s.get("properties") or {}).get("title") for s in (meta.get("sheets") or [])]
    if title in titles:
        return
    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={"requests": [{"addSheet": {"properties": {"title": title}}}]},
    ).execute()


def write_values(svc, sid: str, title: str, rows: list[list[str]]) -> None:
    ensure_sheet(svc, sid, title)
    rng = "'%s'" % title.replace("'", "''")
    svc.spreadsheets().values().clear(spreadsheetId=sid, range=rng).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": rows},
    ).execute()


def list_titles(svc, sid: str) -> list[str]:
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
    return [(s.get("properties") or {}).get("title") or "" for s in (meta.get("sheets") or [])]


def read_all(svc, sid: str, title: str) -> list[list[str]]:
    rng = "'%s'" % title.replace("'", "''")
    return (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range=rng)
        .execute()
        .get("values")
        or []
    )


def upsert_makers(svc, sid: str) -> int:
    ensure_sheet(svc, sid, SHEET_MAKER)
    raw = read_all(svc, sid, SHEET_MAKER)
    header = list(MAKER_HEADERS)
    by: dict[str, list[str]] = {}
    if raw:
        header = [str(h) for h in raw[0]]
        for h in MAKER_HEADERS:
            if h not in header:
                header.append(h)
        idx = {h: i for i, h in enumerate(header)}
        for r in raw[1:]:
            padded = list(r) + [""] * (len(header) - len(r))
            key = (padded[idx.get("メーカー", 0)] or "").strip()
            if key:
                by[key] = padded[: len(header)]
    idx = {h: i for i, h in enumerate(header)}

    def cell(row: list[str], name: str) -> str:
        i = idx.get(name)
        if i is None or i >= len(row):
            return ""
        return str(row[i] or "").strip()

    for rec in seed_makers():
        key = rec["メーカー"]
        old = by.get(key)
        if old and cell(old, "第2クエリ語") and not rec.get("第2クエリ語"):
            continue
        row = [""] * len(header)
        if old:
            row = list(old) + [""] * (len(header) - len(old))
        for h, v in rec.items():
            if h not in idx:
                continue
            if h == "第2クエリ語" and old and cell(old, "第2クエリ語") and not v:
                continue
            row[idx[h]] = v
        by[key] = row
    out = [header] + [by[k] for k in sorted(by.keys())]
    write_values(svc, sid, SHEET_MAKER, out)
    return len(out) - 1


def probe(svc, sid: str, label: str) -> list[str] | None:
    try:
        t = list_titles(svc, sid)
        print("tabs_%s" % label, t)
        return t
    except Exception as e:
        msg = str(e)
        if "403" in msg or "404" in msg or "permission" in msg.lower():
            print("NEED_SHARE_%s" % label, OAUTH_NEED, sid)
            return None
        raise


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--list-only", action="store_true")
    args = p.parse_args()

    svc = sheets_service(write=True, interactive=False)
    if not svc:
        print("NO_CREDS token_sheets_rw.json")
        return 2

    t_comp = probe(svc, COMPETITOR_SS, "competitor")
    t_res = probe(svc, RESEARCH_SS, "research")
    if args.list_only:
        return 0 if t_comp is not None and t_res is not None else 3
    if t_comp is None:
        return 3

    n_maker = upsert_makers(svc, COMPETITOR_SS)
    print("makers", n_maker, "sheet", SHEET_MAKER, COMPETITOR_SS)

    if t_res is None:
        print("research_skip")
        return 4

    gate = load_gate_rows()
    pass_n = sum(1 for r in gate[1:] if len(r) >= 8 and r[7] == "通過")
    drop_n = len(gate) - 1 - pass_n
    now = utc_now()
    write_values(svc, RESEARCH_SS, T_CAND, gate)
    write_values(
        svc,
        RESEARCH_SS,
        T_SUM,
        [
            ["メーカー", "通過数", "対象数", "落ち数", "更新"],
            ["石原水産", str(pass_n), str(len(gate) - 1), str(drop_n), now],
        ],
    )
    write_values(
        svc,
        RESEARCH_SS,
        T_PROF,
        [
            ["カテゴリ", "価格下限", "順位段階1", "順位段階2", "順位段階3"],
            ["食品", "2000", "150000", "60000", "30000"],
        ],
    )
    write_values(
        svc,
        RESEARCH_SS,
        T_LOG,
        [
            ["runId", "内容", "state", "at"],
            ["seed_ishihara_gate", "CSV門を①候補へ。メーカーマスタは競合DB", "DONE", now],
        ],
    )
    print("research_wrote", T_CAND, T_SUM, T_PROF, T_LOG, "n_cand", len(gate) - 1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
