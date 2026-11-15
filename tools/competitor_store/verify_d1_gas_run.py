# -*- coding: utf-8 -*-
"""D1 GAS 実行後の合否判定。①ログ＋品番リスト。"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, T_CAND, as_dicts, read_all, sheets_service
from dry_t1_list_paste import DATA_START, HEADER_ROW, LIST_TITLE, build_row, col_letter, latest_full
from schema import SHEET_KEEPA_FULL


def main() -> int:
    svc = sheets_service()
    logs = read_all(svc, RESEARCH_SS, "①ログ")
    print("=== ①ログ 直近10行 ===")
    for row in logs[-10:]:
        print("\t".join(str(c) for c in row))

    d1 = [r for r in logs[1:] if len(r) > 1 and "D1" in str(r[1])]
    print("\n=== D1ログ 全件 ===")
    for r in d1:
        print(r)

    latest_d1 = d1[-1] if d1 else None
    if not latest_d1:
        print("\nOVERALL FAIL: D1ログなし")
        return 1

    state = str(latest_d1[2] if len(latest_d1) > 2 else "")
    detail = str(latest_d1[7] if len(latest_d1) > 7 else "")
    run_id = str(latest_d1[0] if latest_d1 else "")
    print("\nlatest_d1 runId=", run_id, "state=", state, "detail=", detail)

    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    pass_rows = []
    seen = set()
    for r in cand:
        a = str(r.get("ASIN") or "").strip().upper()
        if a and str(r.get("門結果") or "") == "通過" and a not in seen:
            seen.add(a)
            pass_rows.append(r)
    print("\n=== ①候補 門通過 ===", len(pass_rows))

    rng = "'%s'!%d:%d" % (LIST_TITLE.replace("'", "''"), HEADER_ROW, HEADER_ROW)
    headers = [
        str(x).replace("\n", " ").strip()
        for x in (svc.spreadsheets().values().get(spreadsheetId=RESEARCH_SS, range=rng).execute().get("values") or [[]])[0]
    ]
    asin_i = headers.index("ASIN") if "ASIN" in headers else -1
    raw = read_all(svc, RESEARCH_SS, LIST_TITLE)
    exist = set()
    for r in raw[DATA_START - 1 :]:
        if asin_i >= 0 and asin_i < len(r):
            a = str(r[asin_i] or "").strip().upper()
            if len(a) == 10:
                exist.add(a)

    new_asins = [str(r.get("ASIN") or "").upper() for r in pass_rows if str(r.get("ASIN") or "").upper() not in exist]
    print("list_asins", len(exist), "pass_not_in_list", len(new_asins))

    import re

    append_m = re.search(r"append=(\d+)", detail)
    miss_m = re.search(r"miss_full=(\d+)", detail)
    append_n = int(append_m.group(1)) if append_m else 0
    miss_full = int(miss_m.group(1)) if miss_m else 0

    log_ok = state == "DONE" and "FAILED" not in detail.upper()
    if append_n > 0:
        outcome_ok = append_n > 0 and miss_full == 0
    else:
        outcome_ok = len(new_asins) == 0 or miss_full == 0

    # verify appended ASINs in list
    fh, frows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    fm = latest_full(frows)
    frng = "'%s'!A%d:%s%d" % (LIST_TITLE.replace("'", "''"), DATA_START, col_letter(len(headers)), DATA_START)
    frow = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RESEARCH_SS, range=frng, valueRenderOption="FORMULA")
        .execute()
        .get("values")
        or [[]]
    )[0]
    formula_cols = {i for i, v in enumerate(frow) if str(v).startswith("=")}

    verify_asins = []
    if append_n > 0:
        # last N data rows with 10-char ASIN
        data_rows = []
        for r in raw[DATA_START - 1 :]:
            if asin_i < len(r):
                a = str(r[asin_i] or "").strip().upper()
                if len(a) == 10:
                    data_rows.append(a)
        verify_asins = data_rows[-append_n:] if append_n <= len(data_rows) else data_rows
    elif new_asins:
        verify_asins = new_asins[:3]

    mapping_ok = True
    for asin in verify_asins:
        cand_row = next((r for r in pass_rows if str(r.get("ASIN") or "").upper() == asin), {})
        full = fm.get(asin)
        if not full:
            print("WARN no_keepa_full", asin)
            mapping_ok = False
            continue
        expected = build_row(headers, full, formula_cols, cand_row)
        idx = {headers[i]: i for i in range(len(headers))}
        for r in raw[DATA_START - 1 :]:
            if asin_i < len(r) and str(r[asin_i]).upper() == asin:
                print("\n=== 行確認", asin, "===")
                checks = ["メーカー名", "最安値", "順位", "ランキング", "カテゴリー", "Keepaグラフ", "amazonページ", "調査日"]
                for k in checks:
                    if k not in idx:
                        continue
                    got = str(r[idx[k]] if idx[k] < len(r) else "")
                    exp = str(expected[idx[k]] if idx[k] < len(expected) else "")
                    ok = got.replace(",", "") == exp.replace(",", "") or got == exp
                    if not ok and k in ("Keepaグラフ",):
                        ok = True  # optional
                    if not ok:
                        mapping_ok = False
                    mark = "OK" if ok else "NG"
                    print(" ", k, mark, "| got:", got[:70], "| exp:", exp[:70])
                break

    overall = log_ok and outcome_ok and mapping_ok
    print("\n=== 判定 ===")
    print("log_ok", log_ok, "outcome_ok", outcome_ok, "mapping_ok", mapping_ok)
    print("OVERALL", "PASS" if overall else "FAIL")
    return 0 if overall else 1


if __name__ == "__main__":
    raise SystemExit(main())
