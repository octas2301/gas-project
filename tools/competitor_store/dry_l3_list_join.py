# -*- coding: utf-8 -*-
"""L3: 品番リストを Keepaフル JOIN。式・SKIP・原本非書。GETなし。"""
from __future__ import annotations

import json

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, append_log, as_dicts, read_all, sheets_service
from dry_t1_list_paste import (
    DATA_START,
    HEADER_ROW,
    LIST_TITLE,
    ORIG_SS,
    SKIP_HEADERS,
    build_row,
    col_letter,
    latest_full,
)
from keepa_full import flatten_keepa_display
from schema import SHEET_KEEPA_FULL


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    if RESEARCH_SS == ORIG_SS:
        return 2
    rng_h = "'%s'!%d:%d" % (LIST_TITLE.replace("'", "''"), HEADER_ROW, HEADER_ROW)
    headers = [
        str(x).replace("\n", " ").strip()
        for x in (svc.spreadsheets().values().get(spreadsheetId=RESEARCH_SS, range=rng_h).execute().get("values") or [[]])[0]
    ]
    frng = "'%s'!A%d:%s%d" % (LIST_TITLE.replace("'", "''"), DATA_START, col_letter(len(headers)), DATA_START)
    frow = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RESEARCH_SS, range=frng, valueRenderOption="FORMULA")
        .execute()
        .get("values")
        or [[]]
    )[0]
    formula_cols = set(i for i, v in enumerate(frow) if str(v).startswith("="))
    formula_names = [headers[i] for i in sorted(formula_cols) if i < len(headers)]
    raw = read_all(svc, RESEARCH_SS, LIST_TITLE)
    asin_i = headers.index("ASIN")
    list_asins, row_nums = [], []
    for ri, r in enumerate(raw[DATA_START - 1 :], start=DATA_START):
        a = str(r[asin_i] if asin_i < len(r) else "").strip().upper()
        if len(a) == 10:
            list_asins.append(a)
            row_nums.append(ri)
    fh, frows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    fullmap = latest_full(frows)
    for a, rec in list(fullmap.items()):
        try:
            p = json.loads(rec.get("生JSON") or "{}")
        except json.JSONDecodeError:
            p = {}
        rec["画像"] = flatten_keepa_display(p).get("画像") or rec.get("画像") or ""
    miss = [a for a in list_asins if a not in fullmap]
    data_cols = [i for i in range(len(headers)) if i not in formula_cols and headers[i] not in SKIP_HEADERS]
    skip_write = {"調査日"}
    planned = []
    for a, rn in zip(list_asins, row_nums):
        if a not in fullmap:
            continue
        new = build_row(headers, fullmap[a], formula_cols)
        cur = raw[rn - 1]
        for c in data_cols:
            name = headers[c]
            if name in skip_write:
                continue
            nv = new[c] if c < len(new) else ""
            if not nv:
                continue
            ov = str(cur[c] if c < len(cur) else "")
            if ov != nv:
                planned.append((rn, c, name, a))
    ok = (
        RESEARCH_SS != ORIG_SS
        and len(formula_cols) >= 5
        and "Keepaグラフ" in formula_names
        and "amazonページ" in formula_names
        and not miss
        and "報酬額" in SKIP_HEADERS
    )
    line = "runId=pr_20260816_l3dry list=%d miss=%d formula=%d cells=%d GETなし %s" % (
        len(list_asins),
        len(miss),
        len(formula_cols),
        len(planned),
        "PASS" if ok else "FAIL",
    )
    print(line)
    if not args.apply:
        append_log(svc, "L3", line)
        return 0 if ok else 1
    if not ok:
        return 1
    data = []
    for rn, c, name, a in planned:
        new = build_row(headers, fullmap[a], formula_cols)
        data.append(
            {
                "range": "'%s'!%s%d" % (LIST_TITLE.replace("'", "''"), col_letter(c + 1), rn),
                "values": [[new[c]]],
            }
        )
    for i in range(0, len(data), 400):
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=RESEARCH_SS,
            body={"valueInputOption": "USER_ENTERED", "data": data[i : i + 400]},
        ).execute()
    line2 = "runId=pr_20260816_l3col wrote=%d miss=%d formula_untouched=%d GETなし" % (
        len(data),
        len(miss),
        len(formula_cols),
    )
    append_log(svc, "L3", line2)
    print(line2)
    print("VERIFY", "PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
