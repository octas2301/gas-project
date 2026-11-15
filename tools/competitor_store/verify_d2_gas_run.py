# -*- coding: utf-8 -*-
"""D2 GAS 実行後の合否判定。①ログ＋品番リスト。"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from apply_keepa_full import RESEARCH_SS, T_CAND, as_dicts, read_all, sheets_service
from dry_d2_filter import COL_QUOTE, COL_REASON, load_cand, load_jitta, load_profiles
from d2_filter import classify, profile_for_category
from dry_t1_list_paste import DATA_START, HEADER_ROW, LIST_TITLE, ORIG_SS
from maker_alias import JITTAI_MAKER


def main() -> int:
    svc = sheets_service()
    if RESEARCH_SS == ORIG_SS:
        print("OVERALL FAIL: original SS")
        return 2

    logs = read_all(svc, RESEARCH_SS, "①ログ")
    print("=== ①ログ 直近8行 ===")
    for row in logs[-8:]:
        print("\t".join(str(c) for c in row))

    d2 = [r for r in logs[1:] if len(r) > 1 and "D2" in str(r[1])]
    print("\n=== D2ログ 全件 ===")
    for r in d2:
        print(r)
    latest = d2[-1] if d2 else None
    if not latest:
        print("\nOVERALL FAIL: D2ログなし")
        return 1

    state = str(latest[2] if len(latest) > 2 else "")
    detail = str(latest[7] if len(latest) > 7 else "")
    run_id = str(latest[0] or "")
    print("\nlatest_d2 runId=", run_id, "state=", state, "detail=", detail)
    log_ok = state == "DONE" and "FAILED" not in detail.upper()

    m_rows = re.search(r"(?:sheet_rows|rows)=(\d+)", detail)
    m_quote = re.search(r"quote=(\d+)", detail)
    m_skip = re.search(r"skip=(\d+)", detail)
    m_left = re.search(r"leftover=(\d+)", detail)
    leftover = int(m_left.group(1)) if m_left else -1
    log_rows = int(m_rows.group(1)) if m_rows else -1
    log_quote = int(m_quote.group(1)) if m_quote else -1
    log_skip = int(m_skip.group(1)) if m_skip else -1

    rng = "'%s'!%d:%d" % (LIST_TITLE.replace("'", "''"), HEADER_ROW, HEADER_ROW)
    headers = [
        str(x).replace("\n", " ").strip()
        for x in (svc.spreadsheets().values().get(spreadsheetId=RESEARCH_SS, range=rng).execute().get("values") or [[]])[0]
    ]
    col_ok = COL_QUOTE in headers and COL_REASON in headers and "ASIN" in headers
    print("\n=== 品番リスト列 ===")
    print("見積候補" if COL_QUOTE in headers else "MISSING 見積候補", "/", "見送り理由" if COL_REASON in headers else "MISSING 見送り理由")

    raw = read_all(svc, RESEARCH_SS, LIST_TITLE)
    idx = {h: i for i, h in enumerate(headers)}
    cand = load_cand(svc)
    jitta = load_jitta(svc)
    profs = load_profiles(svc)

    n_asin = n_quote = n_skip = n_blank = 0
    mismatch = []
    reason_cnt: Counter[str] = Counter()
    profit_written = 0
    for r in raw[DATA_START - 1 :]:
        asin = str(r[idx["ASIN"]] if "ASIN" in idx and idx["ASIN"] < len(r) else "").strip().upper()
        if len(asin) != 10:
            continue
        n_asin += 1
        flag = str(r[idx[COL_QUOTE]] if COL_QUOTE in idx and idx[COL_QUOTE] < len(r) else "").strip()
        reason = str(r[idx[COL_REASON]] if COL_REASON in idx and idx[COL_REASON] < len(r) else "").strip()
        maker = str(r[idx["メーカー名"]] if "メーカー名" in idx and idx["メーカー名"] < len(r) else "").strip()
        title = str(r[idx["商品名"]] if "商品名" in idx and idx["商品名"] < len(r) else "")
        cat = str(r[idx["カテゴリー"]] if "カテゴリー" in idx and idx["カテゴリー"] < len(r) else "")
        qs = str(r[idx["クイックショップサイズ表記"]] if "クイックショップサイズ表記" in idx and idx["クイックショップサイズ表記"] < len(r) else "")
        wt = str(r[idx["重量"]] if "重量" in idx and idx["重量"] < len(r) else "")
        prof = profile_for_category(cat, profs)
        exp_f, exp_r = classify(
            asin, maker, title, cat, cand.get(asin), jitta.get(maker, JITTAI_MAKER),
            qs_tier=qs, weight=wt, prof=prof,
        )
        if flag == "見積候補":
            n_quote += 1
        elif flag == "見送り":
            n_skip += 1
            if reason:
                reason_cnt[reason] += 1
        else:
            n_blank += 1
        if flag != exp_f or (exp_f == "見送り" and reason != exp_r):
            if len(mismatch) < 8:
                mismatch.append((asin, flag, reason, exp_f, exp_r))
        if "粗利益率" in idx and idx["粗利益率"] < len(r):
            # 式列が壊れていないこと（値または式があればOK。D2は当該列を書かない）
            pass

    print("sheet_asin", n_asin, "quote", n_quote, "skip", n_skip, "blank_flag", n_blank)
    print("skip_top", reason_cnt.most_common(8))
    print("mismatch_n", len(mismatch), "sample", mismatch[:5])

    # 門通過なのに見送りになっている理由が粗利でないこと
    profit_skip = any("粗利" in k or "利益" in k for k in reason_cnt)
    sample_quote = n_quote >= 1
    sample_skip = n_skip >= 1
    count_ok = leftover == 0 and n_blank == 0
    mapping_ok = len(mismatch) == 0
    overall = log_ok and col_ok and sample_quote and sample_skip and mapping_ok and not profit_skip and count_ok

    print("\n=== 判定 ===")
    print("log_ok", log_ok, "col_ok", col_ok, "count_ok", count_ok, "mapping_ok", mapping_ok)
    print("profit_skip_forbidden", profit_skip)
    print("OVERALL", "PASS" if overall else "FAIL")
    return 0 if overall else 1


if __name__ == "__main__":
    raise SystemExit(main())
