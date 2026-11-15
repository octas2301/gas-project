# -*- coding: utf-8 -*-
"""メーカーサマリ: ①候補＋Keepaフルを集計し競合DBへ。GETなし。既定 dry。"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from client import sheets_service  # noqa: E402
from schema import (  # noqa: E402
    MAKER_SUM_HEADERS,
    SHEET_KEEPA_FULL,
    SHEET_MAKER,
    SHEET_MAKER_SUM,
)
from maker_alias import (  # noqa: E402
    JITTAI_BRAND,
    JITTAI_CHECK,
    JITTAI_UNCHECKED,
    JITTAI_EXCLUDE,
    JITTAI_MAKER,
    amazon_search_url,
    hit_text_for_maker,
    jitta_for_maker,
    resolve_alias,
)

RESEARCH_SS = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
T_CAND = "①候補"
T_LOG = "①ログ"


def utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def read_all(svc, sid: str, title: str) -> list[list[str]]:
    rng = "'%s'" % title.replace("'", "''")
    return svc.spreadsheets().values().get(spreadsheetId=sid, range=rng).execute().get("values") or []


def as_dicts(raw: list[list[str]]) -> tuple[list[str], list[dict]]:
    if not raw:
        return [], []
    h = [str(x) for x in raw[0]]
    rows = []
    for r in raw[1:]:
        rows.append({h[i]: (str(r[i]) if i < len(r) else "") for i in range(len(h))})
    return h, rows


def ensure_maker_sum(svc, sid: str) -> str:
    from init_store import apply_ck_ui_google, ensure_leading_ck_google
    meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if SHEET_MAKER_SUM not in existing:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [{"addSheet": {"properties": {"title": SHEET_MAKER_SUM}}}]},
        ).execute()
    got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'" + SHEET_MAKER_SUM + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    current = [str(x).strip() for x in (got[0] if got else [])]
    if not current:
        svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range="'" + SHEET_MAKER_SUM + "'!A1",
            valueInputOption="RAW",
            body={"values": [MAKER_SUM_HEADERS]},
        ).execute()
        apply_ck_ui_google(svc, sid, SHEET_MAKER_SUM)
        return "wrote_headers"
    current = ensure_leading_ck_google(svc, sid, SHEET_MAKER_SUM, current)
    missing = [h for h in MAKER_SUM_HEADERS if h not in current]
    if not missing:
        return "exists ui=%s" % apply_ck_ui_google(svc, sid, SHEET_MAKER_SUM)
    start = len(current)
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range="'" + SHEET_MAKER_SUM + "'!%s1" % chr(65 + start),
        valueInputOption="RAW",
        body={"values": [missing]},
    ).execute()
    return "added_cols ui=%s" % apply_ck_ui_google(svc, sid, SHEET_MAKER_SUM)


EMPTY_BRAND = {
    "",
    "ノーブランド品",
    "ノーブランド",
    "generic",
    "Generic",
    "GENERIC",
}


def is_empty_brand(s: str) -> bool:
    t = (s or "").strip()
    return t in EMPTY_BRAND or t.lower() == "generic"


def norm_maker(s: str) -> str:
    t = (s or "").strip()
    for suf in ("株式会社", "有限会社", "合同会社", "(株)", "（株）"):
        t = t.replace(suf, "")
    t = t.strip()
    return resolve_alias(t) if t else ""


def maker_dict(cand: list[dict], full: list[dict], master: list[dict]) -> list[str]:
    names: set[str] = set()
    for r in master:
        names.add(norm_maker(str(r.get("メーカー") or "")))
    for r in cand:
        names.add(norm_maker(str(r.get("メーカー") or "")))
    # 別名の親も辞書に入れる（タイトル最長一致用）
    for parent in ("味の素", "いし本食品工業", "エスビー食品", "エバラ", "エースコック"):
        names.add(parent)
    out = [n for n in names if n and not is_empty_brand(n) and len(n) >= 2]
    out.sort(key=len, reverse=True)
    return out


def title_pick(title: str, names: list[str]) -> tuple[str, list[str]]:
    found = [n for n in names if n and n in (title or "")]
    if not found:
        return "", []
    filtered = [n for n in found if not any(n != o and n in o for o in found)]
    if not filtered:
        return "", found
    mx = max(len(x) for x in filtered)
    top = [x for x in filtered if len(x) == mx]
    if len(top) != 1:
        return "", top
    return top[0], top


def resolve_nobrand(full_row: dict, cand_by_asin: dict[str, str], names: list[str]) -> dict:
    asin = str(full_row.get("ASIN") or "").strip().upper()
    title = str(full_row.get("商品名") or "").strip()
    mfr_raw = norm_maker(str(full_row.get("製造者") or ""))
    mfr = mfr_raw if mfr_raw and mfr_raw in names and not is_empty_brand(mfr_raw) else ""
    cand_m = norm_maker(cand_by_asin.get(asin, ""))
    if is_empty_brand(cand_m):
        cand_m = ""
    t_pick, t_hits = title_pick(title, names)
    if cand_m:
        return {
            "st": "振り分け",
            "why": "cand",
            "pick": cand_m,
            "src": "cand",
            "asin": asin,
            "title": title[:80],
            "hits": "",
            "cand": cand_m,
            "製造者": mfr,
        }
    if mfr and t_pick and mfr != t_pick:
        return {
            "st": "衝突",
            "why": "source_disagree",
            "pick": "",
            "src": "製造者=%s|title=%s" % (mfr, t_pick),
            "asin": asin,
            "title": title[:80],
            "hits": "",
            "cand": cand_m,
            "製造者": mfr,
        }
    if mfr:
        return {
            "st": "振り分け",
            "why": "製造者",
            "pick": mfr,
            "src": "製造者",
            "asin": asin,
            "title": title[:80],
            "hits": "",
            "cand": cand_m,
            "製造者": mfr,
        }
    if t_hits and not t_pick:
        return {
            "st": "衝突",
            "why": "title_multi",
            "pick": "",
            "src": "",
            "asin": asin,
            "title": title[:80],
            "hits": "|".join(t_hits),
            "cand": cand_m,
            "製造者": mfr,
        }
    if t_pick:
        return {
            "st": "振り分け",
            "why": "title",
            "pick": t_pick,
            "src": "title",
            "asin": asin,
            "title": title[:80],
            "hits": "",
            "cand": cand_m,
            "製造者": mfr,
        }
    return {
        "st": "未確定",
        "why": "no_source",
        "pick": "",
        "src": "",
        "asin": asin,
        "title": title[:80],
        "hits": "",
        "cand": cand_m,
        "製造者": mfr,
    }


def run_assign_dry(cand: list[dict], full: list[dict], master: list[dict]) -> int:
    names = maker_dict(cand, full, master)
    cand_by_asin = {}
    for r in cand:
        a = str(r.get("ASIN") or "").strip().upper()
        m = str(r.get("メーカー") or "").strip()
        if a and m and not is_empty_brand(m):
            cand_by_asin[a] = norm_maker(m)
    nobrand = [r for r in full if is_empty_brand(str(r.get("ブランド") or ""))]
    stats = {"振り分け": 0, "衝突": 0, "未確定": 0}
    by_why = defaultdict(int)
    by_pick = defaultdict(int)
    samples = {"衝突": [], "未確定": [], "振り分け": []}
    for r in nobrand:
        d = resolve_nobrand(r, cand_by_asin, names)
        stats[d["st"]] += 1
        by_why[d["why"]] += 1
        if d["pick"]:
            by_pick[d["pick"]] += 1
        bag = samples[d["st"]]
        if len(bag) < 8:
            bag.append(d)
    print("assign_dry dict=%d keepa=%d nobrand=%d" % (len(names), len(full), len(nobrand)))
    print("振り分け=%d 衝突=%d 未確定=%d" % (stats["振り分け"], stats["衝突"], stats["未確定"]))
    print("why " + " ".join("%s=%d" % (k, v) for k, v in sorted(by_why.items(), key=lambda x: -x[1])))
    print("pick_top")
    for m, n in sorted(by_pick.items(), key=lambda x: -x[1])[:12]:
        print("  %s %d" % (m, n))
    for st in ("衝突", "未確定", "振り分け"):
        print("sample_%s" % st)
        for d in samples[st]:
            print(
                "  %s why=%s pick=%s src=%s hits=%s cand=%s mfr=%s title=%s"
                % (d["asin"], d["why"], d["pick"], d["src"], d.get("hits") or "", d["cand"], d["製造者"], d["title"])
            )
    print("dry_assign_ok=%s" % (stats["衝突"] / max(len(nobrand), 1) < 0.05))
    print("no write")
    return 0


def append_log(svc, line: str) -> None:
    svc.spreadsheets().values().append(
        spreadsheetId=RESEARCH_SS,
        range="'" + T_LOG + "'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [[utc_now(), "メーカーサマリ", "競合DB", line]]},
    ).execute()


def parent_maker(full_row: dict, cand_by_asin: dict[str, str], names: list[str]) -> str:
    """倉庫SKU寄せ用。①候補メーカー優先。ブランド列はキーにしない。"""
    asin = str(full_row.get("ASIN") or "").strip().upper()
    cand_m = norm_maker(cand_by_asin.get(asin, ""))
    if cand_m and not is_empty_brand(cand_m):
        return cand_m
    b = str(full_row.get("ブランド") or "").strip()
    if not is_empty_brand(b):
        aliased = resolve_alias(b)
        if aliased and aliased != b:
            return aliased
        nb = norm_maker(b)
        if nb in names:
            return nb
        mfr = norm_maker(str(full_row.get("製造者") or ""))
        if mfr in names:
            return mfr
        return "未確定"
    d = resolve_nobrand(full_row, cand_by_asin, names)
    return d["pick"] if d["st"] == "振り分け" and d["pick"] else "未確定"


def empty_agg(maker: str) -> dict:
    return {
        "メーカー": maker,
        "最終調査日": "",
        "asins": set(),
        "wh": set(),
        "pass": 0,
        "drop": 0,
        "iru": 0,
        "inai": 0,
    }


def build_rows(
    cand: list[dict],
    full: list[dict],
    master: list[dict],
    memos: dict[str, str],
    jitta_overrides: dict[str, str] | None = None,
) -> list[dict]:
    """メーカー1行。門は①候補メーカー。倉庫SKUは parent_maker 寄せ。"""
    overrides = jitta_overrides or {}
    names = maker_dict(cand, full, master)
    cand_by_asin = {}
    for r in cand:
        a = str(r.get("ASIN") or "").strip().upper()
        m = norm_maker(str(r.get("メーカー") or ""))
        if a and m and not is_empty_brand(m):
            cand_by_asin[a] = m
    by_asin = {}
    for r in full:
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            by_asin[a] = r
    q2 = {}
    for r in master:
        m = norm_maker(str(r.get("メーカー") or ""))
        if m:
            q2[m] = bool(str(r.get("第2クエリ語") or "").strip())
    agg = {}
    for r in full:
        a = str(r.get("ASIN") or "").strip().upper()
        if not a:
            continue
        m = parent_maker(r, cand_by_asin, names)
        rec = agg.setdefault(m, empty_agg(m))
        rec["wh"].add(a)
    for r in cand:
        m = norm_maker(str(r.get("メーカー") or ""))
        if not m:
            continue
        a = str(r.get("ASIN") or "").strip().upper()
        fr = by_asin.get(a) or {}
        rec = agg.setdefault(m, empty_agg(m))
        st = str(r.get("門結果") or "").strip()
        dt = str(r.get("取得日時") or r.get("更新") or "").strip()
        if a:
            rec["asins"].add(a)
        if st == "通過":
            rec["pass"] += 1
            amz = str(fr.get("Amazon直販") or "").strip()
            if amz == "いる":
                rec["iru"] += 1
            elif amz == "いない":
                rec["inai"] += 1
        elif st == "落ち":
            rec["drop"] += 1
        if dt and dt > rec["最終調査日"]:
            rec["最終調査日"] = dt
    out = []
    now = utc_now()
    for m, rec in agg.items():
        j = jitta_for_maker(m, overrides.get(m, ""))
        if j == JITTAI_BRAND:
            pref = "ブランド"
        elif j == JITTAI_MAKER:
            pref = "メーカー"
        elif j == JITTAI_EXCLUDE:
            pref = "非仕入先"
        elif j == JITTAI_CHECK:
            pref = "要確認"
        elif j == JITTAI_UNCHECKED:
            pref = "未調査"
        else:
            pref = "未設定"
        out.append(
            {
                "メーカー": m,
                "最終調査日": rec["最終調査日"],
                "入口推奨": pref,
                "入口語ヒット": hit_text_for_maker(m),
                "実態": j,
                "実態ck": amazon_search_url(m) if j == JITTAI_CHECK else "",
                "メモ": memos.get(m, ""),
                "更新": now,
            }
        )
    out.sort(key=lambda x: x["メーカー"])
    return out


def main() -> int:
    ap = argparse.ArgumentParser(
        description="メーカーサマリ集計（dry 本線）。全消し書込は非本線。"
    )
    ap.add_argument("--apply", action="store_true", help="非本線。要 --force-full-clear")
    ap.add_argument(
        "--force-full-clear",
        action="store_true",
        help="緊急時のみ。シート全消し＋再構築（本線禁止）",
    )
    ap.add_argument("--assign", action="store_true", help="ノーブランド振り分け dry（非書）")
    args = ap.parse_args()
    svc = sheets_service(write=True, interactive=True)
    if not svc:
        print("no_creds")
        return 1
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    _, master = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_MAKER))
    if args.assign:
        if args.apply:
            print("assign is dry-only; refuse --apply")
            return 2
        return run_assign_dry(cand, full, master)
    memos = {}
    jitta_overrides = {}
    cks = {}
    try:
        _, old = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_MAKER_SUM))
        for r in old:
            m = str(r.get("メーカー") or "").strip()
            note = str(r.get("メモ") or "").strip()
            j = str(r.get("実態") or "").strip()
            ck = r.get("調査CK")
            if m and note:
                memos[m] = note
            if m and ck not in (None, ""):
                cks[m] = ck
            if m and j in (JITTAI_MAKER, JITTAI_BRAND, JITTAI_CHECK, JITTAI_UNCHECKED, JITTAI_EXCLUDE):
                # 人が直した実態を優先（シードよりシート）
                jitta_overrides[m] = j
    except Exception:
        pass
    recs = build_rows(cand, full, master, memos, jitta_overrides)
    for r in recs:
        m = str(r.get("メーカー") or "").strip()
        if m in cks:
            r["調査CK"] = cks[m]
    print(
        "cand=%d full=%d master=%d makers=%d apply=%s"
        % (len(cand), len(full), len(master), len(recs), args.apply)
    )
    by_j = {}
    for r in recs:
        by_j[r["実態"]] = by_j.get(r["実態"], 0) + 1
    print("実態 " + " ".join("%s=%d" % (k, v) for k, v in sorted(by_j.items())))
    for r in recs[:20]:
        print(
            "%s j=%s pref=%s"
            % (r["メーカー"], r["実態"], r["入口推奨"])
        )
    if not args.apply:
        print("dry_ok (no write)")
        print("note: mainline maker insert = GAS prEnsureMakerSumRowTop_ (new only at row2)")
        print("note: --apply full clear is NOT mainline (needs --force-full-clear)")
        return 0
    if not args.force_full_clear:
        print("refuse: dry_maker_sum --apply full clear is NOT mainline")
        print("mainline: GAS ①メニュー新規メーカー先頭挿入。緊急再構築のみ --apply --force-full-clear")
        return 2
    ens = ensure_maker_sum(svc, COMPETITOR_SS)
    values = [MAKER_SUM_HEADERS] + [[r.get(h, "") for h in MAKER_SUM_HEADERS] for r in recs]
    svc.spreadsheets().values().clear(spreadsheetId=COMPETITOR_SS, range="'" + SHEET_MAKER_SUM + "'").execute()
    svc.spreadsheets().values().update(
        spreadsheetId=COMPETITOR_SS,
        range="'" + SHEET_MAKER_SUM + "'!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()
    line = "FORCE_FULL_CLEAR ensure=%s n=%d" % (ens, len(recs))
    append_log(svc, line)
    print("apply_ok " + line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
