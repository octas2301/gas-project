# -*- coding: utf-8 -*-
"""P5計画: 貼付P2並べ替えを読取 dry。シート非書。Aへの組み込みはしない。"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from client import sheets_service
from paste_rank import sort_block
from schema import MASTER_SS_ID

PASTE = "ASIN貼り付け（Keepa用）"
AI = "AI情報取得data"
BLOCK = 10
ASIN_RE = __import__("re").compile(r"^[A-Z0-9]{10}$", __import__("re").I)


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
    print("write=false A_wired=false")
    for b in range(10):
        maker = ""
        if b + 1 < len(ai) and "メーカー名" in ix:
            arow = ai[b + 1]
            maker = str(arow[ix["メーカー名"]] if ix["メーカー名"] < len(arow) else "").strip()
        start = b * BLOCK
        last = 1
        for r, row in enumerate(wide):
            if r < 2:
                continue
            if any(str(row[start + c] if start + c < len(row) else "").strip() for c in range(BLOCK)):
                last = r
        if last < 2:
            continue
        rows = []
        n_circle = 0
        for r, row in enumerate(wide):
            if r < 2 or r > last:
                continue
            ev = row[start + 3] if start + 3 < len(row) else ""
            if str(ev).strip() == "◎":
                n_circle += 1
            rows.append(
                {
                    "asin": str(row[start + 1] if start + 1 < len(row) else "").strip(),
                    "title": row[start + 2] if start + 2 < len(row) else "",
                    "eval": ev,
                    "price": row[start + 4] if start + 4 < len(row) else "",
                    "set_count_cell": row[start + 5] if start + 5 < len(row) else "",
                    "tag": row[start + 7] if start + 7 < len(row) else "",
                }
            )
        out = sort_block(rows, maker)
        cand = sum(1 for x in out if x.get("p2") == "候補")
        unk = sum(1 for x in out if x.get("p2") == "未属性")
        bad = sum(1 for x in out if x.get("p2") == "非候補")
        blank = sum(1 for x in out if not str(x.get("asin") or "").strip())
        circle_after = sum(1 for x in out if str(x.get("eval") or "").strip() == "◎")
        reasons = {}
        for x in out:
            if x.get("p2") == "非候補":
                reasons[x.get("p2_reason") or "?"] = reasons.get(x.get("p2_reason") or "?", 0) + 1
        moved = 0
        for i, x in enumerate(out):
            if str(x.get("asin") or "") != str(rows[i].get("asin") or ""):
                moved += 1
        print(
            "b%d maker=%s n=%d cand=%d pending=%d reject=%d blank=%d circle=%d/%d order_diff=%d"
            % (b, maker, len(out), cand, unk, bad, blank, n_circle, circle_after, moved)
        )
        if reasons:
            print("  reject", reasons)
        if circle_after != n_circle:
            raise SystemExit("fail: ◎ count changed")
    print("ok ◎ preserved, no sheet write")


if __name__ == "__main__":
    main()
