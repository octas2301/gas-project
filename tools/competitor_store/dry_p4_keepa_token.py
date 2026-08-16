# -*- coding: utf-8 -*-
"""Pick 1 listing ASIN not in master cache / Keepaフル, GET once, no sheet write."""
from __future__ import annotations

import gzip
import json
import re
import sys
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "purchase_research_path3"))

from client import sheets_service
from keepa_csv_vs_api import keepa_key
from schema import MASTER_KEEPA_SHEET, MASTER_SS_ID, SHEET_KEEPA_FULL

COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
PASTE = "ASIN貼り付け（Keepa用）"
ASIN_RE = re.compile(r"^[A-Z0-9]{10}$", re.I)


def col_asins(raw, col_name_or_idx):
    if not raw:
        return set()
    if isinstance(col_name_or_idx, int):
        idx = col_name_or_idx
    else:
        h = [str(x).strip() for x in raw[0]]
        if col_name_or_idx not in h:
            return set()
        idx = h.index(col_name_or_idx)
    out = set()
    for r in raw[1:]:
        if idx >= len(r):
            continue
        a = str(r[idx] or "").strip().upper()
        if ASIN_RE.match(a):
            out.add(a)
    return out


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


def main() -> int:
    svc = sheets_service()
    paste = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + PASTE + "'")
        .execute()
        .get("values")
        or []
    )
    cache = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + MASTER_KEEPA_SHEET + "'")
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
    p = paste_asins(paste)
    c = col_asins(cache, "ASIN")
    f = col_asins(full, "ASIN")
    miss_cache = [a for a in p if a not in c]
    miss_both = [a for a in miss_cache if a not in f]
    print("paste=%d cache=%d full=%d miss_cache=%d miss_both=%d" % (len(p), len(c), len(f), len(miss_cache), len(miss_both)))
    pick = (miss_both or miss_cache or p)[:1]
    print("pick", ",".join(pick), "in_cache", pick[0] in c if pick else None, "in_full", pick[0] in f if pick else None)
    if not pick:
        print("NO_ASIN")
        return 2
    key = keepa_key()
    if not key:
        print("NO_KEEPA_KEY")
        return 2
    q = urlencode({"key": key, "domain": "5", "asin": pick[0], "stats": "90", "history": "0"})
    req = Request("https://api.keepa.com/product?" + q, headers={"User-Agent": "OctasP4/1.0"})
    with urlopen(req, timeout=90) as resp:
        raw = resp.read()
        if len(raw) >= 2 and raw[0] == 0x1F and raw[1] == 0x8B:
            raw = gzip.decompress(raw)
        data = json.loads(raw.decode("utf-8"))
    n = len(data.get("products") or [])
    left = data.get("tokensLeft")
    consumed = data.get("tokensConsumed")
    print("write=false get=true asin=%s products=%s tokensLeft=%s tokensConsumed=%s" % (pick[0], n, left, consumed))
    per = None
    if consumed not in (None, "") and n:
        try:
            per = float(consumed) / n
        except (TypeError, ValueError):
            per = None
    print("per_asin", per)
    print("cap_default 20 remains (no code change)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
