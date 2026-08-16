# -*- coding: utf-8 -*-
"""F2b: 画像=メインHYPERLINK、サブ画像=2枚目以降を | 区切り。GETなし。"""
from __future__ import annotations

import json

from apply_keepa_full import COMPETITOR_SS, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from init_store import ensure_keepa_full_google, place_keepa_full_col_after
from keepa_full import IMAGE_SEP, flatten_keepa_display
from schema import SHEET_KEEPA_FULL

COLS = ("画像", "サブ画像")


def recs_from_full(full: list[dict]) -> list[dict]:
    out = []
    for r in full:
        try:
            p = json.loads(r.get("生JSON") or "{}")
        except json.JSONDecodeError:
            p = {}
        d = flatten_keepa_display(p)
        out.append({"画像": d.get("画像") or "", "サブ画像": d.get("サブ画像") or ""})
    return out


def rename_list_header(svc, h: list[str]) -> str:
    if "サブ画像" in h:
        return "already"
    if "画像一覧" not in h:
        return "no_old"
    idx = h.index("画像一覧")
    rng = "'%s'!%s1" % (SHEET_KEEPA_FULL.replace("'", "''"), col_letter(idx + 1))
    svc.spreadsheets().values().update(
        spreadsheetId=COMPETITOR_SS,
        range=rng,
        valueInputOption="RAW",
        body={"values": [["サブ画像"]]},
    ).execute()
    return "renamed_col_%s" % col_letter(idx + 1)


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply-col", action="store_true")
    ap.add_argument("--move-next", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    h, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    if args.move_next:
        i_img = h.index("画像") if "画像" in h else -1
        i_sub = h.index("サブ画像") if "サブ画像" in h else -1
        print("before", "画像", col_letter(i_img + 1) if i_img >= 0 else "?", "サブ画像", col_letter(i_sub + 1) if i_sub >= 0 else "?")
        print(place_keepa_full_col_after(COMPETITOR_SS, "サブ画像", "画像"))
        h2, rows2 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
        i_img = h2.index("画像")
        i_sub = h2.index("サブ画像")
        okm = i_sub == i_img + 1
        line2 = "runId=pr_20260816_f2cmove 画像=%s サブ画像=%s adj=%s n=%d GETなし" % (
            col_letter(i_img + 1),
            col_letter(i_sub + 1),
            okm,
            len(rows2),
        )
        append_log(svc, "F2c", line2)
        print(line2)
        print("VERIFY", "PASS" if okm else "FAIL")
        return 0 if okm else 1
    recs = recs_from_full(full)
    n = len(recs)
    main_n = sum(1 for x in recs if "HYPERLINK" in x["画像"])
    sub_n = sum(1 for x in recs if x["サブ画像"])
    leak = sum(1 for x in recs if x["サブ画像"] and x["画像"] and x["画像"].split('"')[1] in x["サブ画像"] if '"' in x["画像"])
    ok = main_n == n and leak == 0 and IMAGE_SEP == "|"
    line = "runId=pr_20260816_f2bdry n=%d mainHL=%d sub=%d leak_main=%d sep=%s GETなし %s" % (
        n, main_n, sub_n, leak, IMAGE_SEP, "PASS" if ok else "FAIL"
    )
    print(line)
    print("sample_main", recs[0]["画像"] if recs else None)
    print("sample_sub", (recs[0]["サブ画像"][:120] if recs else None))
    if not args.apply_col:
        append_log(svc, "F2b", line)
        return 0 if ok else 1
    if not ok:
        return 1
    print("rename", rename_list_header(svc, h))
    print("ensure", ensure_keepa_full_google(COMPETITOR_SS))
    h, _ = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    miss = [c for c in COLS if c not in h]
    if miss:
        print("missing", miss)
        return 1
    idx_img = h.index("画像")
    vals_img = [[r["画像"]] for r in recs]
    svc.spreadsheets().values().update(
        spreadsheetId=COMPETITOR_SS,
        range="'%s'!%s2" % (SHEET_KEEPA_FULL.replace("'", "''"), col_letter(idx_img + 1)),
        valueInputOption="USER_ENTERED",
        body={"values": vals_img},
    ).execute()
    idx_sub = h.index("サブ画像")
    vals_sub = [[r["サブ画像"]] for r in recs]
    svc.spreadsheets().values().update(
        spreadsheetId=COMPETITOR_SS,
        range="'%s'!%s2" % (SHEET_KEEPA_FULL.replace("'", "''"), col_letter(idx_sub + 1)),
        valueInputOption="RAW",
        body={"values": vals_sub},
    ).execute()
    h2, rows2 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    # FORMULA check via values USER_ENTERED stored as formula
    filled = sum(1 for r in rows2 if str(r.get("画像") or "").strip())
    listed = sum(1 for r in rows2 if str(r.get("サブ画像") or "").strip())
    old = "画像一覧" in h2
    line2 = "runId=pr_20260816_f2bcol n=%d 画像=%d サブ画像=%d old一覧残=%s GETなし" % (
        len(rows2), filled, listed, old
    )
    append_log(svc, "F2b", line2)
    print(line2)
    vok = filled == n and "サブ画像" in h2 and not old
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
