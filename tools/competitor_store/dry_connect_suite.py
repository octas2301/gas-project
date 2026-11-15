# -*- coding: utf-8 -*-
"""接続テスト T1–T4。サマリ実態→次工程のつなぎ。

T1 製造メーカー: Catalog→Keepa→門→①候補（少量）
T2 入口ブランド: 入口語でCatalog、親名で①候補
T3 非仕入先: 洗いに進まない（拒否が合格）
T4 要確認: メール下書きのみ・洗いしない

既定 dry。書込は --apply --live（T1/T2のみ）。モールヒット非書。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SMOKE = ROOT.parent / "spapi_smoke"
PATH3 = ROOT.parent / "purchase_research_path3"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(SMOKE))
sys.path.insert(0, str(PATH3))

from apply_keepa_full import (  # noqa: E402
    COMPETITOR_SS,
    RESEARCH_SS,
    T_CAND,
    append_log,
    append_rows,
    as_dicts,
    prepend_rows,
    fetch_stats90,
    keepa_key,
    read_all,
    sheets_service,
    utc_now,
)
from dry_w3_gate import gate_product, load_profiles  # noqa: E402
from keepa_full import keepa_get_needed, latest_row_for_asin, product_to_full_row  # noqa: E402
from maker_alias import (  # noqa: E402
    JITTAI_BRAND,
    JITTAI_CHECK,
    JITTAI_EXCLUDE,
    JITTAI_MAKER,
    JITTAI_UNCHECKED,
    jitta_for_maker,
)
from schema import KEEPA_FULL_HEADERS, PURPOSE_RESEARCH, SHEET_HITS, SHEET_KEEPA_FULL  # noqa: E402
from search_catalog_keywords import asins_from_body, search_catalog_items  # noqa: E402
from spapi_smoke import _load_config, request_lwa_access_token, resolve_cred  # noqa: E402

SS_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs/edit"
)


def spapi_access():
    cfg = _load_config(SMOKE / "config.local.json")
    access = str(
        request_lwa_access_token(
            resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID"),
            resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET"),
            resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN"),
        ).get("access_token")
        or ""
    )
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    mp = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    return access, endpoint, mp, "OctasConnect/1.0"


def catalog_hits(keyword: str, needles: list[str], access, endpoint, mp, ua):
    resp = search_catalog_items(endpoint, access, mp, ua, keyword, page_size=20, page_token="")
    body = resp.json() if resp.content else {}
    rows = asins_from_body(body)
    out = []
    for r in rows:
        blob = str(r.get("title") or "") + " " + str(r.get("brand") or "")
        if any(n in blob for n in needles):
            out.append(r)
    next_tok = bool((body.get("pagination") or {}).get("nextToken"))
    return out, resp.status_code, next_tok


def cand_map(rows: list[dict]) -> dict[str, int]:
    m = {}
    for i, r in enumerate(rows):
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            m[a] = i
    return m


def run_wash(
    svc,
    parent: str,
    entry: str,
    needles: list[str],
    cap: int,
    run_id: str,
    apply: bool,
    live: bool,
    tag: str,
) -> int:
    access, endpoint, mp, ua = spapi_access()
    hits, code, next_tok = catalog_hits(entry, needles, access, endpoint, mp, ua)
    print(
        "%s catalog entry=%s http=%s title_hit=%d next=%s"
        % (tag, entry, code, len(hits), "Y" if next_tok else "N")
    )
    want = [str(r["asin"]).upper() for r in hits[:cap]]
    if not want:
        print("%s FAIL title_hit=0" % tag)
        return 3
    hits_n = max(0, len(read_all(svc, COMPETITOR_SS, SHEET_HITS)) - 1)
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    need, skip = [], []
    for a in want:
        if keepa_get_needed(latest_row_for_asin(full, a)):
            need.append(a)
        else:
            skip.append(a)
    print("%s want=%d need_get=%d skip=%d" % (tag, len(want), len(need), len(skip)))
    if not apply:
        print("%s dry_ok" % tag)
        return 0
    if not live:
        print("%s need --live" % tag)
        return 2
    key = keepa_key()
    if not key:
        print("NO_KEEPA_KEY")
        return 2
    fetched = utc_now()
    products_by = {}
    recs = []
    data = fetch_stats90(key, need) if need else {"products": [], "tokensConsumed": 0, "tokensLeft": ""}
    consumed = data.get("tokensConsumed") or 0
    left = data.get("tokensLeft")
    for p in data.get("products") or []:
        rec = product_to_full_row(p, fetched, purpose=PURPOSE_RESEARCH)
        products_by[rec["ASIN"]] = p
        if not rec["ASIN"]:
            continue
        latest = latest_row_for_asin(full, rec["ASIN"])
        if latest and str(latest.get("価格指紋") or "") == rec["価格指紋"]:
            continue
        rawj = json.loads(rec["生JSON"] or "{}")
        if "csv" in rawj:
            print("FAIL csv", rec["ASIN"])
            return 4
        recs.append(rec)
        full.append(rec)
    if recs:
        live_h = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))[0] or list(KEEPA_FULL_HEADERS)
        prepend_rows(svc, COMPETITOR_SS, SHEET_KEEPA_FULL, live_h, recs)
    _, full2 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    for a in skip:
        row = latest_row_for_asin(full2, a)
        if not row:
            continue
        try:
            products_by[a] = json.loads(row.get("生JSON") or "{}")
        except json.JSONDecodeError:
            pass
    profs = load_profiles(svc)
    ch, crow = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    idx = {h: i for i, h in enumerate(ch)}
    by = cand_map(crow)
    new_rows = []
    pass_n = drop_n = 0
    for a in want:
        p = products_by.get(a) or {}
        if not p:
            print("WARN no_product", a)
            continue
        st, why = gate_product(p, profs) if profs else ("落ち", "プロファイルなし")
        if st == "通過":
            pass_n += 1
        else:
            drop_n += 1
        eans = p.get("eanList") or []
        rec = {
            "メーカー": parent,
            "ASIN": a,
            "商品名": str(p.get("title") or ""),
            "税込価格": "",
            "順位90": "",
            "JAN": str(eans[0] if eans else p.get("ean") or ""),
            "発見経路": tag + "_catalog",
            "sellerId": "",
            "門結果": st,
            "門理由": why,
            "runId": run_id,
            "取得日時": fetched,
        }
        if a in by:
            r = crow[by[a]]
            if r.get("sellerId"):
                rec["sellerId"] = r.get("sellerId")
            line = [r.get(h, "") for h in ch]
            for h, v in rec.items():
                if h in idx:
                    line[idx[h]] = v
            rng = "'%s'!A%d" % (T_CAND.replace("'", "''"), by[a] + 2)
            svc.spreadsheets().values().update(
                spreadsheetId=RESEARCH_SS,
                range=rng,
                valueInputOption="RAW",
                body={"values": [line]},
            ).execute()
        else:
            new_rows.append(rec)
    if new_rows:
        append_rows(svc, RESEARCH_SS, T_CAND, ch, new_rows)
    hits_after = max(0, len(read_all(svc, COMPETITOR_SS, SHEET_HITS)) - 1)
    msg = (
        "%s ok parent=%s entry=%s full+=%s consumed=%s left=%s pass=%s drop=%s new_cand=%s"
        % (tag, parent, entry, len(recs), consumed, left, pass_n, drop_n, len(new_rows))
    )
    append_log(svc, tag, msg)
    print(msg)
    if hits_after != hits_n:
        print("FAIL hits changed")
        return 5
    print("%s_PASS" % tag)
    return 0


def run_t3(svc, name: str, apply: bool) -> int:
    j = jitta_for_maker(name)
    print("T3 name=%s 実態=%s (expect %s)" % (name, j, JITTAI_EXCLUDE))
    if j != JITTAI_EXCLUDE:
        print("T3 FAIL 実態が非仕入先ではない")
        return 2
    if apply:
        # 意図的に洗わない
        append_log(svc, "T3", "skip_wash name=%s reason=非仕入先" % name)
    print("T3 refuse_wash=OK (Catalog/Keepa/①候補へ進まない)")
    print("T3_PASS")
    return 0


def run_t4(svc, name: str, apply: bool) -> int:
    j = jitta_for_maker(name)
    print("T4 name=%s 実態=%s (expect %s)" % (name, j, JITTAI_UNCHECKED))
    if j != JITTAI_UNCHECKED:
        print("T4 FAIL 実態が未調査ではない")
        return 2
    _, rows = as_dicts(read_all(svc, COMPETITOR_SS, "メーカーサマリ"))
    row = next((r for r in rows if str(r.get("メーカー") or "") == name), {})
    print("To: （担当）")
    print("Subject: 【仕入れ検討①】要確認 接続T4 %s" % name)
    print()
    print("実態が要確認のため洗いを止めました。製造メーカー／入口ブランド／非仕入先へ確定してください。")
    print("シート: %s" % SS_URL)
    print(
        "メーカー=%s 倉庫SKU=%s 通過=%s 入口語ヒット=%s"
        % (name, row.get("倉庫SKU"), row.get("通過数"), (row.get("入口語ヒット") or "")[:50])
    )
    print("--- 自動送信なし ---")
    if apply:
        append_log(svc, "T4", "mail_draft_only name=%s no_wash" % name)
    print("T4_PASS")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--test", choices=["t1", "t2", "t3", "t4", "all"], default="all")
    ap.add_argument("--cap", type=int, default=5)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--skip-t1", action="store_true", help="T1済みのとき")
    args = ap.parse_args()
    cap = max(1, min(20, args.cap))
    svc = sheets_service(write=True, interactive=True)
    if not svc:
        print("no_creds")
        return 1

    tests = []
    if args.test == "all":
        tests = ["t1", "t2", "t3", "t4"]
        if args.skip_t1:
            tests = ["t2", "t3", "t4"]
    else:
        tests = [args.test]

    results = {}
    for t in tests:
        print("==== %s ====" % t.upper())
        if t == "t1":
            parent = "尾西食品"
            if jitta_for_maker(parent) != JITTAI_MAKER:
                results[t] = 2
                continue
            rid = "pr_t1_" + utc_now()[:10].replace("-", "")
            results[t] = run_wash(
                svc, parent, parent, [parent], cap, rid, args.apply, args.live, "T1"
            )
        elif t == "t2":
            parent = "味の素"
            entry = "AGF"
            if jitta_for_maker(parent) != JITTAI_BRAND:
                print("T2 FAIL 実態")
                results[t] = 2
                continue
            rid = "pr_t2_" + utc_now()[:10].replace("-", "")
            results[t] = run_wash(
                svc,
                parent,
                entry,
                ["AGF", "エージーエフ", "味の素AGF"],
                cap,
                rid,
                args.apply,
                args.live,
                "T2",
            )
        elif t == "t3":
            results[t] = run_t3(svc, "リフココ", args.apply)
        elif t == "t4":
            results[t] = run_t4(svc, "BBQ FOUNDATION", args.apply)

    print("==== SUMMARY ====")
    for t, code in results.items():
        print("%s %s" % (t.upper(), "PASS" if code == 0 else "FAIL:%s" % code))
    bad = [t for t, c in results.items() if c != 0]
    return 0 if not bad else 1


if __name__ == "__main__":
    raise SystemExit(main())
