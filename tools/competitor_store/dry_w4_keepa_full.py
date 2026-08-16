# -*- coding: utf-8 -*-
"""W4 dry: 貼付ASIN vs Keepaフル。GETしない。スプシ非書。"""
from __future__ import annotations

import re
from datetime import datetime, timezone

from client import sheets_service
from keepa_full import classify_keepa_get
from schema import MASTER_SS_ID, SHEET_KEEPA_FULL

COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
PASTE = "ASIN貼り付け（Keepa用）"
ASIN_RE = re.compile(r"^[A-Z0-9]{10}$", re.I)
BLOCK_COLS = 10


def as_dicts(raw):
    if not raw:
        return [], []
    h = [str(x).strip() for x in raw[0]]
    rows = []
    for r in raw[1:]:
        rows.append({h[i]: (str(r[i]) if i < len(r) else "") for i in range(len(h))})
    return h, rows


def paste_asins(vals):
    out = []
    seen = set()
    if len(vals) < 3:
        return out
    for b in range(10):
        col = b * BLOCK_COLS + 1
        for r in vals[2:]:
            if col >= len(r):
                continue
            a = str(r[col] or "").strip().upper()
            if ASIN_RE.match(a) and a not in seen:
                seen.add(a)
                out.append(a)
    return out


def main() -> None:
    svc = sheets_service()
    full_raw = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMPETITOR_SS, range="'" + SHEET_KEEPA_FULL + "'")
        .execute()
        .get("values")
        or []
    )
    _, full_rows = as_dicts(full_raw)
    paste_raw = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + PASTE + "'")
        .execute()
        .get("values")
        or []
    )
    asins = paste_asins(paste_raw)
    now = datetime.now(timezone.utc)
    plan = classify_keepa_get(asins, full_rows, now)
    from keepa_full import plan_a_keepa_fetch

    plan_a = plan_a_keepa_fetch(asins, full_rows, cache_asins=[], now=now)
    print("write=false get=false paste_n=%d full_n=%d skip_fresh=%d need_get=%d" % (
        len(asins),
        len(full_rows),
        len(plan["skip_fresh"]),
        len(plan["need_get"]),
    ))
    print("hydrate=%d fetch=%d (cache empty)" % (len(plan_a["hydrate"]), len(plan_a["fetch"])))
    print("skip_sample", ",".join(plan["skip_fresh"][:8]))
    print("need_sample", ",".join(plan["need_get"][:8]))


if __name__ == "__main__":
    main()
