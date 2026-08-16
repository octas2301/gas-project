# -*- coding: utf-8 -*-
"""L1: 旧 構成% 列を削除。Amazon28列は残す。GETなし。既定 dry。"""
from __future__ import annotations

from apply_keepa_full import COMPETITOR_SS, append_log, as_dicts, read_all, sheets_service
from schema import AMAZON_SELLER_CAT_COLS, SHEET_SELLER


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    meta = svc.spreadsheets().get(spreadsheetId=COMPETITOR_SS).execute()
    sid = None
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == SHEET_SELLER:
            sid = s["properties"]["sheetId"]
            break
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    dup_i = [i for i, x in enumerate(h) if str(x).startswith("構成% ")]
    miss28 = [c for c in AMAZON_SELLER_CAT_COLS if c not in h]
    rec = rows[0] if rows else {}
    food = rec.get("食品・飲料・お酒")
    ok = sid is not None and len(dup_i) >= 1 and not miss28 and str(food) in ("46.9", "46.90")
    line = "runId=pr_20260816_l1dry dup=%d miss28=%d food=%s GETなし %s" % (
        len(dup_i),
        len(miss28),
        food,
        "PASS" if ok else "FAIL",
    )
    print(line)
    print("dup_names", [h[i] for i in dup_i])
    if not args.apply:
        append_log(svc, "L1", line)
        return 0 if ok else 1
    if not ok:
        return 1
    reqs = []
    for i in sorted(dup_i, reverse=True):
        reqs.append(
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": sid,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i + 1,
                    }
                }
            }
        )
    svc.spreadsheets().batchUpdate(spreadsheetId=COMPETITOR_SS, body={"requests": reqs}).execute()
    h2, rows2 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    left = [x for x in h2 if str(x).startswith("構成% ")]
    food2 = rows2[0].get("食品・飲料・お酒") if rows2 else ""
    vok = not left and str(food2) in ("46.9", "46.90") and all(c in h2 for c in AMAZON_SELLER_CAT_COLS)
    line2 = "runId=pr_20260816_l1col del=%d left=%d food=%s GETなし %s" % (
        len(dup_i),
        len(left),
        food2,
        "PASS" if vok else "FAIL",
    )
    append_log(svc, "L1", line2)
    print(line2)
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
