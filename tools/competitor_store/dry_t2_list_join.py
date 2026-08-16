# -*- coding: utf-8 -*-
"""T2: 既存品番リスト行を Keepaフルで埋め直し。式列・SKIP・原本は触らない。既定 dry。"""
from __future__ import annotations

from keepa_full import flatten_keepa_display
import json

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
from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, append_log, as_dicts, read_all, sheets_service
from schema import SHEET_KEEPA_FULL

# 式で入っている列は上書きしない（T1のコピー式）
# 画像URL は Keepaフルの HYPERLINK に合わせる（式列でなければ）


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    if RESEARCH_SS == ORIG_SS:
        print("refuse original SS")
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
    list_asins = []
    row_nums = []
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
    join = [a for a in list_asins if a in fullmap]
    miss = [a for a in list_asins if a not in fullmap]

    img_i = headers.index("画像URL") if "画像URL" in headers else -1
    if img_i >= 0 and row_nums:
        last = max(row_nums)
        rng = "'%s'!%s%d:%s%d" % (
            LIST_TITLE.replace("'", "''"),
            col_letter(img_i + 1),
            DATA_START,
            col_letter(img_i + 1),
            last,
        )
        fcol = (
            svc.spreadsheets()
            .values()
            .get(spreadsheetId=RESEARCH_SS, range=rng, valueRenderOption="FORMULA")
            .execute()
            .get("values")
            or []
        )
        for i in range(last - DATA_START + 1):
            idx = DATA_START - 1 + i
            while idx >= len(raw):
                raw.append([])
            while len(raw[idx]) <= img_i:
                raw[idx].append("")
            if i < len(fcol) and fcol[i]:
                raw[idx][img_i] = str(fcol[i][0])

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
                planned.append((rn, c, name, ov[:40], nv[:40]))

    img_hl = 0
    if img_i >= 0:
        for a, rn in zip(list_asins, row_nums):
            if a not in fullmap:
                continue
            cur = raw[rn - 1]
            if img_i < len(cur) and "HYPERLINK" in str(cur[img_i]):
                img_hl += 1

    keepa_url_formula = "Keepaグラフ" in formula_names and "amazonページ" in formula_names
    img_not_formula = img_i >= 0 and img_i not in formula_cols
    only_img = (not planned) or all(p[2] == "画像URL" for p in planned)
    hl_plan = all("HYPERLINK" in (fullmap[a].get("画像") or "") for a in join)
    already = img_hl == 82 and len(planned) == 0
    need_write = len(planned) == 82 and only_img
    ok = (
        RESEARCH_SS != ORIG_SS
        and len(formula_cols) >= 5
        and keepa_url_formula
        and img_not_formula
        and len(join) == 82
        and hl_plan
        and only_img
        and (already or need_write)
        and "報酬額" in SKIP_HEADERS
    )
    line = (
        "runId=pr_20260816_t2dry list=%d join=%d miss=%d formula=%d cells_diff=%d imgHL=%d GETなし %s"
        % (
            len(list_asins),
            len(join),
            len(miss),
            len(formula_cols),
            len(planned),
            img_hl,
            "PASS" if ok else "FAIL",
        )
    )
    print(line)
    print("formula_names", formula_names)
    chg = {}
    for rn, c, name, ov, nv in planned:
        chg[name] = chg.get(name, 0) + 1
    print("diff_by_col", sorted(chg.items(), key=lambda x: -x[1])[:20])
    if planned:
        print("sample_diff", planned[0])

    if not args.apply:
        append_log(svc, "T2", line)
        return 0 if ok else 1
    if not ok:
        return 1
    if already:
        line2 = "runId=pr_20260816_t2col already join=%d imgHL=%d GETなし" % (len(join), img_hl)
        append_log(svc, "T2", line2)
        print(line2)
        print("VERIFY PASS")
        return 0
    data = []
    for rn, c, name, ov, nv in planned:
        # rebuild full new value
        a = None
        for aa, rrn in zip(list_asins, row_nums):
            if rrn == rn:
                a = aa
                break
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
    raw2 = read_all(svc, RESEARCH_SS, LIST_TITLE)
    img_join_hl = 0
    if img_i >= 0 and row_nums:
        last = max(row_nums)
        rng = "'%s'!%s%d:%s%d" % (
            LIST_TITLE.replace("'", "''"),
            col_letter(img_i + 1),
            DATA_START,
            col_letter(img_i + 1),
            last,
        )
        fcol = (
            svc.spreadsheets()
            .values()
            .get(spreadsheetId=RESEARCH_SS, range=rng, valueRenderOption="FORMULA")
            .execute()
            .get("values")
            or []
        )
        for a, rn in zip(list_asins, row_nums):
            if a not in fullmap:
                continue
            i = rn - DATA_START
            if i < len(fcol) and fcol[i] and "HYPERLINK" in str(fcol[i][0]):
                img_join_hl += 1
    line2 = "runId=pr_20260816_t2col list=%d join=%d wrote=%d imgHL_join=%d formula_untouched=%d GETなし" % (
        len(list_asins),
        len(join),
        len(data),
        img_join_hl,
        len(formula_cols),
    )
    append_log(svc, "T2", line2)
    print(line2)
    vok = img_join_hl == 82
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
