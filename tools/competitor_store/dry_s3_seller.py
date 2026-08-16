# -*- coding: utf-8 -*-
"""S3 dry: セラー③巡回の次手。GETなし。storefront/query はしない。"""
from __future__ import annotations

from apply_keepa_full import COMPETITOR_SS, append_log, as_dicts, read_all, sheets_service
from schema import SELLER_HEADERS, SHEET_SELLER


def main() -> int:
    svc = sheets_service(write=True)
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    n = len(rows)
    ids = [str(r.get("sellerId") or "").strip() for r in rows]
    ids = [x for x in ids if x]
    miss_h = [x for x in SELLER_HEADERS if x not in h]
    ok = n >= 1 and not miss_h and "sellerId" in h
    line = (
        "runId=pr_20260816_s3dry n=%d sellerIds=%d miss_h=%s GETなし next=query要承認 %s"
        % (n, len(ids), ",".join(miss_h) or "-", "PASS" if ok else "FAIL")
    )
    print(line)
    print("ids", ids[:10])
    print("next", "③は /query GET。token消費。本線は承認後。フルから貯めない")
    append_log(svc, "S3", line)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
