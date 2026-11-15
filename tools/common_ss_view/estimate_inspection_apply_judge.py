# -*- coding: utf-8 -*-
"""見積検品: 判断列追加・手数料不明アラート・可否／分析の再計算。

1. 仕入可否の右隣に「判断」列を挿入（無ければ）
2. 各行の可否・判断・分析を品番リストから再計算
3. 見送り=行グレー／手数料不明=可否・判断セルをアラート色
4. 検討ASINの月販・順位を Keepa から転記

使い方:
  python tools/common_ss_view/estimate_inspection_apply_judge.py --dry
  python tools/common_ss_view/estimate_inspection_apply_judge.py
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "competitor_store"))
from client import sheets_service
import estimate_inspection_build_rows as b
import estimate_inspection_kentou_asin as kentou

RES = b.RES
SHEET = b.SHEET

# アラート色（手数料不明）
ALERT_BG = {"red": 1.0, "green": 0.85, "blue": 0.4}
GRAY_BG = {"red": 0.85, "green": 0.85, "blue": 0.85}
WHITE_BG = {"red": 1.0, "green": 1.0, "blue": 1.0}


def ensure_judge_col(svc, sheet_id: int, dry: bool) -> bool:
    hdr = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{SHEET}'!1:1")
        .execute()
        .get("values")
        or [[]]
    )[0]
    hdr = [str(h).strip() for h in hdr]
    if len(hdr) > 5 and hdr[5] == "判断":
        print("判断 col already at F")
        return False
    if any(h == "判断" for h in hdr):
        print("判断 col exists elsewhere", hdr)
        return False
    # insert at F (index 5) = 仕入可否の右隣
    print("insert 判断 at F")
    if dry:
        return True
    svc.spreadsheets().batchUpdate(
        spreadsheetId=RES,
        body={
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 5,
                            "endIndex": 6,
                        },
                        "inheritFromBefore": True,
                    }
                }
            ]
        },
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=RES,
        range=f"'{SHEET}'!F1",
        valueInputOption="USER_ENTERED",
        body={"values": [["判断"]]},
    ).execute()
    return True


def col_index(hdr: list[str], *needles: str) -> int | None:
    for i, h in enumerate(hdr):
        for n in needles:
            if n in h:
                return i
    return None


def apply_row_colors(svc, sheet_id: int, rows_meta: list[dict], n_cols: int):
    """rows_meta: {row0, decision} 0-based sheet row index for data."""
    reqs = []
    # reset data area bg
    if rows_meta:
        end = max(r["row0"] for r in rows_meta) + 1
        reqs.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": end,
                        "startColumnIndex": 0,
                        "endColumnIndex": n_cols,
                    },
                    "cell": {"userEnteredFormat": {"backgroundColor": WHITE_BG}},
                    "fields": "userEnteredFormat.backgroundColor",
                }
            }
        )
    for r in rows_meta:
        d = r["decision"]
        row0 = r["row0"]
        if d == "見送り":
            reqs.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": row0,
                            "endRowIndex": row0 + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": n_cols,
                        },
                        "cell": {"userEnteredFormat": {"backgroundColor": GRAY_BG}},
                        "fields": "userEnteredFormat.backgroundColor",
                    }
                }
            )
        elif d == "手数料不明":
            # 可否(E=4) と 判断(F=5) をアラート色
            for c0 in (4, 5):
                reqs.append(
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": row0,
                                "endRowIndex": row0 + 1,
                                "startColumnIndex": c0,
                                "endColumnIndex": c0 + 1,
                            },
                            "cell": {"userEnteredFormat": {"backgroundColor": ALERT_BG}},
                            "fields": "userEnteredFormat.backgroundColor",
                        }
                    }
                )
    for i in range(0, len(reqs), 40):
        svc.spreadsheets().batchUpdate(
            spreadsheetId=RES, body={"requests": reqs[i : i + 40]}
        ).execute()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=not args.dry, interactive=False)

    meta = (
        svc.spreadsheets()
        .get(spreadsheetId=RES, ranges=[f"'{SHEET}'"], fields="sheets(properties(sheetId))")
        .execute()
    )
    sid = meta["sheets"][0]["properties"]["sheetId"]
    ensure_judge_col(svc, sid, dry=args.dry)

    insp = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{SHEET}'!A1:Z")
        .execute()
        .get("values")
        or []
    )
    if len(insp) < 2:
        raise SystemExit("empty")
    hdr = [str(c).strip() for c in insp[0]]
    # dry insert 未反映なら仮想で F を判断扱い
    if args.dry and (len(hdr) <= 5 or hdr[5] != "判断"):
        hdr = hdr[:5] + ["判断"] + hdr[5:]
        for r in insp[1:]:
            while len(r) < 5:
                r.append("")
            r.insert(5, "")

    i_dec = 4
    i_judge = col_index(hdr, "判断")
    i_analysis = col_index(hdr, "分析結果")
    i_rec = col_index(hdr, "推奨")
    i_target = col_index(hdr, "対象セット")
    i_units = col_index(hdr, "仕入目安袋")
    i_mprofit = col_index(hdr, "月次利益")
    i_sales = col_index(hdr, "月販想定")
    if i_judge is None:
        raise SystemExit("判断列なし")
    if i_target is not None:
        b.COL_L = i_target + 1

    paste = b.load_paste_meta(svc)
    catalog = b.load_list(svc)

    data_writes = []  # (A1range, values)
    link_items = []
    color_meta = []
    report = []

    for i, row in enumerate(insp[1:], start=2):
        while len(row) < len(hdr):
            row.append("")
        maker = str(row[1]).strip()
        key = str(row[2]).strip()
        base, cost = b.parse_key(key)
        if not base or not cost:
            print("skip parse", i, key)
            continue
        rows = b.collect_rows_for_product(catalog, maker, base, cost)
        if not rows:
            print("skip no catalog", i, key)
            # 判断だけ空でも可否は触らない
            continue
        prod = b.build_product(rows, paste)
        decision = prod["decision"]
        judgment = prod.get("judgment") or ""
        report.append((i, key[:32], decision, judgment))

        # 値の更新（可否・判断・分析・推奨・対象・袋・月利・月販想定）
        row[i_dec] = decision
        row[i_judge] = judgment
        if i_analysis is not None:
            row[i_analysis] = prod["analysis"]
        if i_rec is not None:
            row[i_rec] = prod["rec"]
        if i_target is not None:
            row[i_target] = prod["targets"]
        if i_units is not None:
            row[i_units] = prod["units"]
        if i_mprofit is not None:
            row[i_mprofit] = prod["month_profit"]
        if i_sales is not None:
            row[i_sales] = prod["sales_block"]

        color_meta.append({"row0": i - 1, "decision": decision})
        if i_target is not None:
            link_items.append(
                {
                    "row": i,
                    "text": prod["targets"],
                    "asins": prod.get("target_asins") or [],
                    "gray": decision == "見送り",
                }
            )

        # batch cell updates for E/F and analysis block
        data_writes.append(
            {
                "range": f"'{SHEET}'!E{i}:F{i}",
                "values": [[decision, judgment]],
            }
        )
        if i_analysis is not None and i_rec is not None:
            # 分析・推奨は連続とは限らないが現レイアウトでは L/M
            letter_a = b.col_letter(i_analysis + 1)
            letter_r = b.col_letter(i_rec + 1)
            if i_rec == i_analysis + 1:
                data_writes.append(
                    {
                        "range": f"'{SHEET}'!{letter_a}{i}:{letter_r}{i}",
                        "values": [[prod["analysis"], prod["rec"]]],
                    }
                )
            else:
                data_writes.append(
                    {
                        "range": f"'{SHEET}'!{letter_a}{i}",
                        "values": [[prod["analysis"]]],
                    }
                )
                data_writes.append(
                    {
                        "range": f"'{SHEET}'!{letter_r}{i}",
                        "values": [[prod["rec"]]],
                    }
                )
        if i_units is not None and i_mprofit is not None and i_mprofit == i_units + 1:
            lu = b.col_letter(i_units + 1)
            lm = b.col_letter(i_mprofit + 1)
            data_writes.append(
                {
                    "range": f"'{SHEET}'!{lu}{i}:{lm}{i}",
                    "values": [[prod["units"], prod["month_profit"]]],
                }
            )
        if i_sales is not None:
            ls = b.col_letter(i_sales + 1)
            data_writes.append(
                {
                    "range": f"'{SHEET}'!{ls}{i}",
                    "values": [[prod["sales_block"]]],
                }
            )

    print("--- judgments ---")
    for t in report:
        print(t)

    if args.dry:
        print("DRY data_writes", len(data_writes), "links", len(link_items))
        print("kentou dry…")
        kentou.apply_kentou(svc, dry=True)
        return

    # values batch
    # Sheets API batchUpdate values
    for i in range(0, len(data_writes), 20):
        chunk = data_writes[i : i + 20]
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=RES,
            body={"valueInputOption": "USER_ENTERED", "data": chunk},
        ).execute()

    if link_items:
        b.apply_l_links_batch(svc, sid, link_items)
        print("LINKS", len(link_items))

    apply_row_colors(svc, sid, color_meta, max(22, len(hdr)))
    print("COLORS", len(color_meta))

    # 検討ASIN 月販転記（対象セットを上書き追記）
    n = kentou.apply_kentou(svc, dry=False)
    print("KENTOU", n)
    print("DONE")


if __name__ == "__main__":
    main()
