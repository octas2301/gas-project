# -*- coding: utf-8 -*-
"""品番リスト: 袋×ケースのQ修正＋希望式（販手・FBA・保管・目標利益込み）。

希望税込1個 = 1点販売 − 1点FBA − 最低利益1個 − 販手/Q − 保管/Q
手数料（販手・保管）が欠ける行は希望を空（事実が揃うまで出さない）。

使い方:
  python tools/common_ss_view/estimate_list_hope_q_fix.py --dry
  python tools/common_ss_view/estimate_list_hope_q_fix.py
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "competitor_store"))
from client import sheets_service

RES = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
SHEET = "仕入検討品番リスト"
BAG_CASE = re.compile(r"(\d+)\s*袋入?\s*[×xXｘ]\s*[\(（]?\s*(\d+)\s*ケース")


def col_letter(n1: int) -> str:
    n = n1
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def hope_formula(row: int) -> str:
    # 販手 BV・保管 BW が数値のときだけ。欠ける行は空＝手数料不明と揃える
    return (
        f'=IF(OR(NOT(ISNUMBER(Q{row})),Q{row}<=0,'
        f'NOT(ISNUMBER(W{row})),NOT(ISNUMBER(X{row})),'
        f'NOT(ISNUMBER(BV{row})),NOT(ISNUMBER(BW{row}))),'
        f'"",AO{row}-AP{row}-AN{row}-BV{row}/Q{row}-BW{row}/Q{row})'
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=not args.dry, interactive=False)

    hdr = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{SHEET}'!4:4")
        .execute()["values"][0]
    )
    idx = {str(h).strip(): i for i, h in enumerate(hdr)}
    title_c = idx["商品名"]
    q_c = idx["同商品のセット数量"]
    aq_c = idx["目標仕入額(税込み）"]
    asin_c = idx["ASIN"]

    titles = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{SHEET}'!{col_letter(title_c+1)}5:{col_letter(title_c+1)}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    qs = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{SHEET}'!{col_letter(q_c+1)}5:{col_letter(q_c+1)}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    asins = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{SHEET}'!{col_letter(asin_c+1)}5:{col_letter(asin_c+1)}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )

    last = 4 + len(titles)
    print("data rows", len(titles), "last", last)

    q_updates = []
    for i, row in enumerate(titles):
        t = str(row[0] if row else "")
        m = BAG_CASE.search(t)
        if not m:
            continue
        q_new = int(m.group(1)) * int(m.group(2))
        try:
            q_old = int(float(qs[i][0])) if i < len(qs) and qs[i] and qs[i][0] != "" else None
        except Exception:
            q_old = None
        r = i + 5
        a = str(asins[i][0]).strip().upper() if i < len(asins) and asins[i] else ""
        print(f"Q fix r{r} {a} {q_old}->{q_new} | {t[:50]}")
        if q_old != q_new:
            q_updates.append({"range": f"'{SHEET}'!{col_letter(q_c+1)}{r}", "values": [[q_new]]})

    # AQ template row2 + all data rows
    aq_letter = col_letter(aq_c + 1)
    aq_updates = [{"range": f"'{SHEET}'!{aq_letter}2", "values": [[hope_formula(2)]]}]
    # chunk formula column
    chunk = 400
    for start in range(5, last + 1, chunk):
        end = min(start + chunk - 1, last)
        vals = [[hope_formula(r)] for r in range(start, end + 1)]
        aq_updates.append(
            {
                "range": f"'{SHEET}'!{aq_letter}{start}:{aq_letter}{end}",
                "values": vals,
            }
        )
    print("Q updates", len(q_updates), "AQ chunks", len(aq_updates))
    print("sample AQ5", hope_formula(5))

    if args.dry:
        print("DRY ok")
        return

    # Q first
    if q_updates:
        for i in range(0, len(q_updates), 50):
            svc.spreadsheets().values().batchUpdate(
                spreadsheetId=RES,
                body={"valueInputOption": "USER_ENTERED", "data": q_updates[i : i + 50]},
            ).execute()
        print("WROTE Q", len(q_updates))

    for i in range(0, len(aq_updates), 5):
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=RES,
            body={"valueInputOption": "USER_ENTERED", "data": aq_updates[i : i + 5]},
        ).execute()
        print("WROTE AQ chunk", i, "/", len(aq_updates))

    # verify key ASINs
    for asin, row_hint in (("B0829GQVSW", 1786), ("B01J5FY1UA", 3787), ("B005FOBQOA", 1008)):
        pass
    verify = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{SHEET}'!A5:{col_letter(max(aq_c, q_c)+1)}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    # lighter verify by direct cells
    for row in (1008, 1786, 3787):
        vals = (
            svc.spreadsheets()
            .values()
            .get(
                spreadsheetId=RES,
                range=f"'{SHEET}'!Q{row}:Q{row}",
                valueRenderOption="UNFORMATTED_VALUE",
            )
            .execute()
            .get("values")
            or [[]]
        )
        aq = (
            svc.spreadsheets()
            .values()
            .get(
                spreadsheetId=RES,
                range=f"'{SHEET}'!AQ{row}:AR{row}",
                valueRenderOption="UNFORMATTED_VALUE",
            )
            .execute()
            .get("values")
            or [[]]
        )
        print(f"verify r{row} Q={vals[0][0] if vals and vals[0] else None} AQ/AR={aq[0] if aq else None}")

    print("DONE")


if __name__ == "__main__":
    main()
