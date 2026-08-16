# -*- coding: utf-8 -*-
"""⑧ 週次メーカー決定。Catalog GET しない。既定 dry。"""
from __future__ import annotations

import argparse
import sys

from apply_keepa_full import COMPETITOR_SS, append_log, append_rows, as_dicts, read_all, sheets_service, utc_now
from schema import MAKER_HEADERS, SHEET_META, SHEET_MAKER, SHEET_SELLER

PICK = "永谷園"
NEXT = "サトウ食品"
NAGA_Q2 = "お茶づけ|ふりかけ|みそ汁|お吸いもの|ちらし|チャーハン|そうざい|カレー"
NAGA_URL = "https://www.nagatanien.co.jp/product/"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--q2-master", action="store_true", help="永谷園の第2クエリ語をメーカーマスタへ")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    _, makers = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_MAKER))
    _, sellers = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    master = {str(m.get("メーカー") or "").strip() for m in makers}
    extracted = set()
    if sellers:
        extracted = {x for x in str(sellers[0].get("抽出メーカー") or "").split("|") if x}
    print("master", sorted(master))
    print("extracted", sorted(extracted))
    print("pick", PICK, "in_master", PICK in master, "next", NEXT)
    if args.q2_master:
        rec = {
            "メーカー": PICK,
            "第2クエリ語": NAGA_Q2,
            "採取元": "official",
            "公式URL": NAGA_URL,
            "取得日時": utc_now(),
            "再採取しない": "TRUE",
            "メモ": "公式商品ナビ。1語Catalogは320打ち切り。Keepa一括はしない",
        }
        print("q2", rec["第2クエリ語"])
        if not args.apply:
            print("dry no write")
            return 0
        if PICK in master:
            append_log(svc, "S8q2", "runId=pr_20260815_s8q2 skip exists")
            print("skip exists")
            return 0
        append_rows(svc, COMPETITOR_SS, SHEET_MAKER, MAKER_HEADERS, [rec])
        append_log(svc, "S8q2", "runId=pr_20260815_s8q2 append 永谷園 q2 official KeepaGETなし")
        print("ok master append")
        return 0
    if not args.apply:
        print("dry no write")
        return 0
    _, meta = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_META))
    keys = {str(r.get("キー") or "") for r in meta}
    action = "skip"
    if "週次メーカー" not in keys:
        rng = "'" + SHEET_META.replace("'", "''") + "'!A1"
        svc.spreadsheets().values().append(
            spreadsheetId=COMPETITOR_SS,
            range=rng,
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [["週次メーカー", PICK], ["週次メーカー候補", NEXT], ["週次メーカー日", "2026-08-15"]]},
        ).execute()
        action = "append"
    line = (
        "runId=pr_20260815_s8w %s pick=%s skip=石原|Generic next=%s CatalogGETなし"
        % (action, PICK, NEXT)
    )
    append_log(svc, "S8", line)
    print("ok", line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
