# -*- coding: utf-8 -*-
"""S2: 競合DB セラータブへ第1版1行（モリタ）。storefront=1しない。既定は dry。"""
from __future__ import annotations

import argparse
import sys

from apply_keepa_full import (
    COMPETITOR_SS,
    RESEARCH_SS,
    T_CAND,
    append_log,
    append_rows,
    as_dicts,
    read_all,
    sheets_service,
)
from init_store import ensure_seller_google
from schema import SELLER_HEADERS, SHEET_SELLER

MORITA = "AYC4Z8PML8T30"
FOOD = "57239051"


def virtual_row(cand: list[dict]) -> dict:
    morita = [c for c in cand if str(c.get("sellerId") or "").upper() == MORITA]
    makers = sorted(
        {str(c.get("メーカー") or "").strip() for c in morita if str(c.get("メーカー") or "").strip()}
    )
    return {
        "sellerId": MORITA,
        "店名": "モリタストア",
        "対象カテゴリ": FOOD,
        "巡回日": "2026-08-15",
        "asinList件数": str(len(morita)),
        "抽出メーカー": "|".join(makers),
        "ピック": "通過",
        "採取元": "query+storefront0",
        "メモ": "asinListは/query件数。storefront=1しない",
        "_n": len(morita),
        "_makers": makers,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    if not svc:
        print("no_creds")
        return 2
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    rec = virtual_row(cand)
    print("virtual n=%s makers=%s" % (rec["_n"], rec["_makers"]))
    if not args.apply:
        print("dry no write")
        return 0
    print("ensure", ensure_seller_google(COMPETITOR_SS))
    _, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    ids = {str(r.get("sellerId") or "").upper() for r in rows}
    action = "skip"
    if MORITA not in ids:
        body = {k: rec.get(k, "") for k in SELLER_HEADERS}
        append_rows(svc, COMPETITOR_SS, SHEET_SELLER, SELLER_HEADERS, [body])
        action = "append"
    _, after = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    print("rows", len(after), after)
    append_log(
        svc,
        "S2",
        "runId=pr_20260815_s2w %s seller=%s n=%s makers=%s"
        % (action, MORITA, rec["_n"], rec["抽出メーカー"]),
    )
    print("ok", action)
    return 0


if __name__ == "__main__":
    sys.exit(main())
