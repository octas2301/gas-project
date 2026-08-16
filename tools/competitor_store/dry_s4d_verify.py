# -*- coding: utf-8 -*-
"""S4d: Amazon28列の合否再確認。GETなし。"""
from __future__ import annotations

from apply_keepa_full import COMPETITOR_SS, append_log, as_dicts, read_all, sheets_service
from schema import AMAZON_SELLER_CAT_COLS, SHEET_SELLER


def main() -> int:
    svc = sheets_service(write=True)
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    rec = rows[0] if rows else {}
    miss = [c for c in AMAZON_SELLER_CAT_COLS if c not in h]
    vals = []
    for c in AMAZON_SELLER_CAT_COLS:
        try:
            vals.append(float(rec.get(c) or 0))
        except (TypeError, ValueError):
            vals.append(-1)
    s = round(sum(v for v in vals if v >= 0), 1)
    food = rec.get("食品・飲料・お酒")
    dup = [x for x in h if str(x).startswith("構成% ")]
    ok = (
        not miss
        and len(rows) == 1
        and abs(s - 100) < 1.5
        and str(food) in ("46.9", "46.90")
        and rec.get("sellerId") == "AYC4Z8PML8T30"
    )
    line = (
        "runId=pr_20260816_s4dver n=%d sum=%s food=%s miss=%d dup構成%%=%d GETなし %s"
        % (len(rows), s, food, len(miss), len(dup), "PASS" if ok else "FAIL")
    )
    print(line)
    print("next_blocked", "見積後 週次G5後 次セラー未 永谷園全洗いCatalog要承認 T2miss43 GETせず")
    append_log(svc, "S4d", line)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
