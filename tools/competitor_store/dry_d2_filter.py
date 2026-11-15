# -*- coding: utf-8 -*-
"""D2 機械絞込。GAS prMenuListFilterD2 と同ロジック。既定 dry（--apply で書込）。"""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, T_CAND, as_dicts, read_all, sheets_service
from d2_filter import classify, profile_for_category
from dry_t1_list_paste import DATA_START, HEADER_ROW, LIST_TITLE, ORIG_SS
from maker_alias import JITTAI_MAKER
from schema import SHEET_MAKER_SUM

T_PROF = "①プロファイル"

COL_QUOTE = "見積候補"
COL_REASON = "見送り理由"


def load_jitta(svc) -> dict[str, str]:
    _, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_MAKER_SUM))
    out = {}
    for r in rows:
        m = str(r.get("メーカー") or "").strip()
        if m:
            out[m] = str(r.get("実態") or JITTAI_MAKER).strip() or JITTAI_MAKER
    return out


def load_cand(svc) -> dict[str, dict]:
    _, rows = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    out = {}
    for r in rows:
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            out[a] = r
    return out


def load_profiles(svc) -> list[dict]:
    _, rows = as_dicts(read_all(svc, RESEARCH_SS, T_PROF))
    out = []
    for r in rows:
        en = str(r.get("有効") or "はい")
        if en in ("いいえ", "0"):
            continue
        out.append(r)
    return out


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--cap", type=int, default=500, help="ASIN件数上限（空行は数えない）")
    args = ap.parse_args()
    svc = sheets_service(write=args.apply)
    if RESEARCH_SS == ORIG_SS:
        print("FAIL refuse original SS")
        return 2

    rng = "'%s'!%d:%d" % (LIST_TITLE.replace("'", "''"), HEADER_ROW, HEADER_ROW)
    headers = [
        str(x).replace("\n", " ").strip()
        for x in (svc.spreadsheets().values().get(spreadsheetId=RESEARCH_SS, range=rng).execute().get("values") or [[]])[0]
    ]
    for col in (COL_QUOTE, COL_REASON):
        if col not in headers:
            headers.append(col)

    idx = {h: i for i, h in enumerate(headers)}
    need = ["ASIN", "メーカー名", "商品名", "カテゴリー"]
    if not all(k in idx for k in need):
        print("FAIL missing cols", [k for k in need if k not in idx])
        return 1

    raw = read_all(svc, RESEARCH_SS, LIST_TITLE)
    cand = load_cand(svc)
    jitta = load_jitta(svc)
    profs = load_profiles(svc)
    flags: list[list[str]] = []
    reasons: list[list[str]] = []
    reason_cnt: Counter[str] = Counter()
    n_quote = n_skip = n_blank = n_asin = leftover = 0
    cap = max(1, int(args.cap))

    last_asin_i = -1
    data = raw[DATA_START - 1 :]
    for i, r in enumerate(data):
        a = str(r[idx["ASIN"]] if idx["ASIN"] < len(r) else "").strip().upper()
        if len(a) == 10:
            last_asin_i = i
    if last_asin_i < 0:
        print("FAIL no ASIN rows")
        return 1
    data = data[: last_asin_i + 1]

    for r in data:
        while len(r) < len(headers):
            r.append("")
        asin = str(r[idx["ASIN"]] or "").strip().upper()
        prev_f = str(r[idx[COL_QUOTE]] if COL_QUOTE in idx and idx[COL_QUOTE] < len(r) else "")
        prev_r = str(r[idx[COL_REASON]] if COL_REASON in idx and idx[COL_REASON] < len(r) else "")
        if len(asin) != 10:
            flags.append([prev_f])
            reasons.append([prev_r])
            n_blank += 1
            continue
        n_asin += 1
        if n_asin > cap:
            flags.append([prev_f])
            reasons.append([prev_r])
            leftover += 1
            continue
        maker = str(r[idx["メーカー名"]] or "").strip()
        title = str(r[idx["商品名"]] or "")
        cat = str(r[idx["カテゴリー"]] or "")
        j = jitta.get(maker, JITTAI_MAKER)
        qs = str(r[idx["クイックショップサイズ表記"]] if "クイックショップサイズ表記" in idx else "")
        wt = str(r[idx["重量"]] if "重量" in idx else "")
        prof = profile_for_category(cat, profs)
        f, reason = classify(
            asin, maker, title, cat, cand.get(asin), j, qs_tier=qs, weight=wt, prof=prof
        )
        flags.append([f])
        reasons.append([reason])
        if f == "見積候補":
            n_quote += 1
        else:
            n_skip += 1
            if reason:
                reason_cnt[reason] += 1

    print("sheet_rows", len(flags), "asin", n_asin, "quote", n_quote, "skip", n_skip, "leftover", leftover, "blank_rows", n_blank, "cap", cap)
    print("skip_top", reason_cnt.most_common(8))
    ok = n_asin >= 1 and leftover == 0 and n_quote + n_skip == n_asin
    print("DRY", "PASS" if ok else "FAIL")

    if not args.apply:
        return 0 if ok else 1

    hdr_row = headers
    svc.spreadsheets().values().update(
        spreadsheetId=RESEARCH_SS,
        range="'%s'!A%d" % (LIST_TITLE.replace("'", "''"), HEADER_ROW),
        valueInputOption="RAW",
        body={"values": [hdr_row]},
    ).execute()
    idx = {h: i for i, h in enumerate(hdr_row)}
    q_col = idx[COL_QUOTE] + 1
    r_col = idx[COL_REASON] + 1
    n = len(flags)
    svc.spreadsheets().values().update(
        spreadsheetId=RESEARCH_SS,
        range="'%s'!%s%d" % (LIST_TITLE.replace("'", "''"), _col(q_col), DATA_START),
        valueInputOption="RAW",
        body={"values": flags},
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=RESEARCH_SS,
        range="'%s'!%s%d" % (LIST_TITLE.replace("'", "''"), _col(r_col), DATA_START),
        valueInputOption="RAW",
        body={"values": reasons},
    ).execute()
    print("APPLY ok rows", n)
    return 0


def _col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


if __name__ == "__main__":
    raise SystemExit(main())
