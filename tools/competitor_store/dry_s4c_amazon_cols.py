# -*- coding: utf-8 -*-
"""S4c: Amazonカテゴリ28列をヘッダにし、既存構成％を展開。GETなし。"""
from __future__ import annotations

from apply_keepa_full import COMPETITOR_SS, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from init_store import ensure_seller_google
from schema import AMAZON_SELLER_CAT_COLS, SHEET_SELLER

PREFIX = "構成% "


def pcts_from_row(rec: dict) -> dict[str, float]:
    out = {}
    for k, v in rec.items():
        if not str(k).startswith(PREFIX):
            continue
        name = str(k)[len(PREFIX) :].strip()
        try:
            out[name] = float(v)
        except (TypeError, ValueError):
            continue
    return out


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    rec = rows[0] if rows else {}
    got = pcts_from_row(rec)
    known = set(AMAZON_SELLER_CAT_COLS)
    mapped = {n: p for n, p in got.items() if n in known}
    leftover = round(sum(p for n, p in got.items() if n not in known), 1)
    ok = (
        "食品・飲料・お酒" in AMAZON_SELLER_CAT_COLS
        and len(AMAZON_SELLER_CAT_COLS) == 28
        and abs(sum(got.values()) - 100) < 1.5
        and "食品・飲料・お酒" in mapped
    )
    line = "runId=pr_20260816_s4cdry mapped=%d leftover=%s GETなし %s" % (
        len(mapped),
        leftover,
        "PASS" if ok else "FAIL",
    )
    print(line)
    print("mapped", mapped)
    if not args.apply:
        append_log(svc, "S4c", line)
        return 0 if ok else 1
    if not ok:
        return 1
    print("ensure", ensure_seller_google(COMPETITOR_SS))
    h2, _ = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    miss = [c for c in AMAZON_SELLER_CAT_COLS if c not in h2]
    if miss:
        print("missing", miss)
        return 1
    vals = []
    for name in AMAZON_SELLER_CAT_COLS:
        if name == "不明":
            vals.append(leftover if leftover else mapped.get(name, 0))
        else:
            vals.append(mapped.get(name, 0))
    data = []
    for name, pct in zip(AMAZON_SELLER_CAT_COLS, vals):
        data.append(
            {
                "range": "'%s'!%s2" % (SHEET_SELLER.replace("'", "''"), col_letter(h2.index(name) + 1)),
                "values": [[pct]],
            }
        )
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=COMPETITOR_SS,
        body={"valueInputOption": "USER_ENTERED", "data": data},
    ).execute()
    h3, rows3 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    r = rows3[0]
    food = float(r.get("食品・飲料・お酒") or 0)
    pet = float(r.get("ペット用品") or 0)
    nzero = sum(1 for c in AMAZON_SELLER_CAT_COLS if str(r.get(c) or "") != "")
    vok = abs(food - 46.9) < 0.05 and abs(pet - 35.9) < 0.05 and nzero == 28
    line2 = "runId=pr_20260816_s4ccol food=%s pet=%s cols=%d GETなし %s" % (
        r.get("食品・飲料・お酒"),
        r.get("ペット用品"),
        nzero,
        "PASS" if vok else "FAIL",
    )
    append_log(svc, "S4c", line2)
    print(line2)
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
