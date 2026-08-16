# -*- coding: utf-8 -*-
"""P1計画: 空ASINスロット数だけ。Catalog GETしない。貼付非書。"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from client import sheets_service
from schema import MASTER_SS_ID

PASTE = "ASIN貼り付け（Keepa用）"
AI = "AI情報取得data"
BLOCK = 10
ASIN_RE = re.compile(r"^[A-Z0-9]{10}$", re.I)


def main() -> None:
    svc = sheets_service()
    wide = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + PASTE + "'!A1:CV80")
        .execute()
        .get("values")
        or []
    )
    ai = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + AI + "'!1:12")
        .execute()
        .get("values")
        or []
    )
    ix = {str(h).strip(): i for i, h in enumerate(ai[0] or [])}
    print("write=false catalog_get=false")
    need_get = 0
    for b in range(10):
        jan = ""
        maker = ""
        if b + 1 < len(ai) and "JANコード" in ix:
            arow = ai[b + 1]
            jan = re.sub(r"\D", "", str(arow[ix["JANコード"]] if ix["JANコード"] < len(arow) else ""))
            if "メーカー名" in ix:
                maker = str(arow[ix["メーカー名"]] if ix["メーカー名"] < len(arow) else "").strip()
        start = b * BLOCK
        last = 1
        existing = 0
        tagged = 0
        for r, row in enumerate(wide):
            if r < 2:
                continue
            if any(str(row[start + c] if start + c < len(row) else "").strip() for c in range(BLOCK)):
                last = r
        empty = 0
        circle_empty = 0
        for r, row in enumerate(wide):
            if r < 2 or r > last:
                continue
            asin = str(row[start + 1] if start + 1 < len(row) else "").strip()
            ev = str(row[start + 3] if start + 3 < len(row) else "").strip()
            tag = str(row[start + 7] if start + 7 < len(row) else "")
            if ASIN_RE.match(asin):
                existing += 1
            if "[機械]" in tag:
                tagged += 1
            if not asin and ev != "◎":
                empty += 1
            if not asin and ev == "◎":
                circle_empty += 1
        if last < 2 and len(jan) < 8:
            continue
        stage = "skip_no_empty" if empty == 0 else "would_GET_A_id"
        if empty:
            need_get += 1
        print(
            "b%d jan=%s maker=%s last=%d asin=%d empty_fillable=%d circle_empty_skip=%d tagged=%d stage=%s"
            % (b, jan, maker, last + 1, existing, empty, circle_empty, tagged, stage)
        )
    print("blocks_would_catalog_get", need_get)
    print("ok no catalog GET")


if __name__ == "__main__":
    main()
