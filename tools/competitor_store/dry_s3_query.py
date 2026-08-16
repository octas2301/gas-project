# -*- coding: utf-8 -*-
"""S3: 台帳1セラーの /query GET（page0のみ）。storefront=1・product・POST禁止。既定 dry。"""
from __future__ import annotations

import gzip
import json
import sys
from datetime import date
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "purchase_research_path3"))

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, T_CAND, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from keepa_csv_vs_api import keepa_key
from schema import SHEET_SELLER

FOOD = "57239051"
QUERY = "https://api.keepa.com/query"


def selection_for(seller_id: str, page: int = 0) -> dict:
    return {
        "rootCategory": [FOOD],
        "sellerIds": [seller_id],
        "productType": ["0"],
        "sort": [["current_SALES", "asc"], ["monthlySold", "desc"]],
        "page": page,
        "perPage": 100,
    }


def query_url(sel: dict, key: str) -> str:
    payload = json.dumps(sel, ensure_ascii=False, separators=(",", ":"))
    q = urlencode({"key": key, "domain": "5", "selection": payload})
    return QUERY + "?" + q


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    if not rows:
        print("no seller rows")
        return 1
    sid = str(rows[0].get("sellerId") or "").strip()
    sel = selection_for(sid, 0)
    ok_shape = (
        sel["rootCategory"] == [FOOD]
        and sel["productType"] == ["0"]
        and sel["page"] == 0
        and sel["perPage"] == 100
        and len(sel["sellerIds"]) == 1
        and sid == "AYC4Z8PML8T30"
    )
    url_nkey = query_url(sel, "REDACTED")
    post_forbidden = "POST" not in url_nkey and "selection=" in url_nkey
    ok = ok_shape and post_forbidden and "storefront" not in json.dumps(sel)
    line = "runId=pr_20260816_s3qdry seller=%s page=0 GET形 %s" % (sid, "PASS" if ok else "FAIL")
    print(line)
    print("selection", json.dumps(sel, ensure_ascii=False))
    if not args.apply:
        append_log(svc, "S3", line)
        return 0 if ok else 1
    if not ok:
        return 1
    key = keepa_key()
    if not key:
        print("no keepa key")
        append_log(svc, "S3", "runId=pr_20260816_s3qcol FAIL no_key")
        return 2
    req = Request(query_url(sel, key), method="GET")
    with urlopen(req, timeout=60) as resp:
        raw = resp.read()
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    body = json.loads(raw.decode("utf-8"))
    asins = [str(x).upper() for x in (body.get("asinList") or []) if str(x).strip()]
    total = body.get("totalResults")
    consumed = body.get("tokensConsumed")
    left = body.get("tokensLeft")
    first = asins[0] if asins else ""
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    cand_a = {str(r.get("ASIN") or "").strip().upper() for r in cand}
    in_cand = sum(1 for a in asins if a in cand_a)
    # update セラー 巡回日・件数
    idx_day = h.index("巡回日") if "巡回日" in h else -1
    idx_n = h.index("asinList件数") if "asinList件数" in h else -1
    today = date.today().isoformat()
    if idx_day >= 0:
        svc.spreadsheets().values().update(
            spreadsheetId=COMPETITOR_SS,
            range="'%s'!%s2" % (SHEET_SELLER.replace("'", "''"), col_letter(idx_day + 1)),
            valueInputOption="RAW",
            body={"values": [[today]]},
        ).execute()
    if idx_n >= 0:
        svc.spreadsheets().values().update(
            spreadsheetId=COMPETITOR_SS,
            range="'%s'!%s2" % (SHEET_SELLER.replace("'", "''"), col_letter(idx_n + 1)),
            valueInputOption="RAW",
            body={"values": [[str(len(asins))]]},
        ).execute()
    vok = isinstance(total, int) and total == len(asins) and len(asins) >= 1 and "products" not in body
    line2 = (
        "runId=pr_20260816_s3qcol total=%s n=%s consumed=%s left=%s first=%s in_cand=%s productGETなし %s"
        % (total, len(asins), consumed, left, first, in_cand, "PASS" if vok else "FAIL")
    )
    append_log(svc, "S3", line2)
    print(line2)
    print("VERIFY", "PASS" if vok else "FAIL")
    # キーを出さない
    return 0 if vok else 1


if __name__ == "__main__":
    sys.exit(main())
