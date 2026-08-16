# -*- coding: utf-8 -*-
"""WRITE計画 dry。Keepaフル非書。Aのrawがあるときだけ追記、という現行フックを数える。"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from client import sheets_service
from schema import MASTER_SS_ID, PURPOSE_LISTING, PURPOSE_RESEARCH, SHEET_KEEPA_FULL

COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
PASTE = "ASIN貼り付け（Keepa用）"
ASIN_RE = re.compile(r"^[A-Z0-9]{10}$", re.I)


def paste_asins(vals):
    out = []
    seen = set()
    for b in range(10):
        col = b * 10 + 1
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
    paste = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + PASTE + "'")
        .execute()
        .get("values")
        or []
    )
    full = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMPETITOR_SS, range="'" + SHEET_KEEPA_FULL + "'")
        .execute()
        .get("values")
        or []
    )
    h = [str(x).strip() for x in (full[0] if full else [])]
    ai = h.index("ASIN") if "ASIN" in h else 1
    pi = h.index("目的") if "目的" in h else None
    full_asins = set()
    listing = 0
    research = 0
    for r in full[1:]:
        a = str(r[ai] if ai < len(r) else "").strip().upper()
        if ASIN_RE.match(a):
            full_asins.add(a)
        p = str(r[pi] if pi is not None and pi < len(r) else "")
        if p == PURPOSE_LISTING:
            listing += 1
        if p == PURPOSE_RESEARCH:
            research += 1
    pas = paste_asins(paste)
    in_full = [a for a in pas if a in full_asins]
    not_full = [a for a in pas if a not in full_asins]
    print("write=false paste=%d full_n=%d research=%d listing=%d" % (len(pas), max(0, len(full) - 1), research, listing))
    print("paste_in_full=%d paste_not_in_full=%d" % (len(in_full), len(not_full)))
    print("not_in_full_sample", ",".join(not_full[:12]))
    print("A_api=0_then_write_append=0 (hook is raw-only)")
    print("would_append_if_A_GET", len(not_full))


if __name__ == "__main__":
    main()
