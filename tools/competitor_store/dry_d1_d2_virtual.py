# -*- coding: utf-8 -*-
"""D1/D2 仮想テスト。仕入検討品番リストへの転記マッピング検証。既定 dry。

--apply-mapping で一時シート D1マッピング_仮 を書く（人間チェックまで消さない）。
--cleanup で削除。
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "purchase_research_path3"))

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, T_CAND, as_dicts, read_all, sheets_service
from dry_t1_list_paste import (
    DATA_START,
    DIRECT_WRITE_HEADERS,
    HEADER_ROW,
    LIST_TITLE,
    MAP,
    ORIG_SS,
    SKIP_HEADERS,
    build_row,
    category_name,
    col_letter,
    latest_full,
    min_cart_new_price,
    list_maker_name,
    rank_display,
)
from schema import SHEET_KEEPA_FULL

MAP_SHEET = "D1マッピング_仮"

# リスト列 → (ソース, ソース列/式, 備考)。全列の正本。
FIELD_RULES: dict[str, tuple[str, str, str]] = {
    "メーカー名": ("①候補", "メーカー", "競合ストアで決めた名称（ブランド洗いの親名）"),
    "最安値": ("computed", "min(Buy Box:現在, 新品:現在)", "カート価格と新品現在の安い方"),
    "ランキング": ("computed", "90日平均順位+カテゴリ", "rank_display()"),
    "順位": ("Keepaフル", "売れ筋ランキング: 90 日平均", "90日平均ランキング"),
    "カテゴリー": ("Keepaフル", "カテゴリ: ルート", "ルートカテゴリ"),
    "Keepaグラフ": ("Keepaフル", "URL: Keepa", "任意（無くても可）"),
    "amazonページ": ("Keepaフル", "URL: Amazon", ""),
    "調査日": ("computed", "date.today()", "転記日"),
    "サイズ": ("computed", "梱包L/W/H", "pack_size()"),
    "重量": ("computed", "梱包_重量_g", "pack_weight()"),
    "過去1か月販売数（表示あれば）": ("Keepaフル", "月間売上", ""),
    "梱包サイズ": ("computed", "サイズと同系", "未使用なら空"),
}

USER_CRITICAL = (
    "メーカー名",
    "最安値",
    "ランキング",
    "順位",
    "カテゴリー",
    "amazonページ",
)

D2_FILTER_COLS = [
    "メーカー名",
    "ASIN",
    "商品名",
    "最安値",
    "順位",
    "ランキング",
    "amazon本体が出品者にいる",
    "出品者数",
    "粗利益率",
    "1点利益額",
    "目標仕入額(税込み）",
    "調査日",
    "画像URL",
    "Keepaグラフ",
    "amazonページ",
]

CAND_FALLBACK = {
    "メーカー名": "メーカー",
    "ASIN": "ASIN",
    "商品名": "商品名",
    "最安値": "税込価格",
    "順位": "順位90",
    "JANコード": "JAN",
}


def sheet_id_by_title(svc, ss: str, title: str) -> int | None:
    meta = svc.spreadsheets().get(spreadsheetId=ss).execute()
    for sh in meta.get("sheets", []):
        if sh["properties"]["title"] == title:
            return int(sh["properties"]["sheetId"])
    return None


def ensure_map_sheet(svc, rows: list[list[str]]) -> None:
    sid = sheet_id_by_title(svc, RESEARCH_SS, MAP_SHEET)
    if sid is None:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=RESEARCH_SS,
            body={"requests": [{"addSheet": {"properties": {"title": MAP_SHEET}}}]},
        ).execute()
    hdr = ["リスト列", "ソース", "ソース列/式", "状態", "備考", "サンプル値"]
    body = [hdr] + rows
    svc.spreadsheets().values().update(
        spreadsheetId=RESEARCH_SS,
        range="'%s'!A1" % MAP_SHEET.replace("'", "''"),
        valueInputOption="RAW",
        body={"values": body},
    ).execute()


def delete_map_sheet(svc) -> bool:
    sid = sheet_id_by_title(svc, RESEARCH_SS, MAP_SHEET)
    if sid is None:
        return False
    svc.spreadsheets().batchUpdate(
        spreadsheetId=RESEARCH_SS,
        body={"requests": [{"deleteSheet": {"sheetId": sid}}]},
    ).execute()
    return True


def classify_col(h: str, i: int, formula_cols: set[int]) -> tuple[str, str, str, str]:
    """(ソース, ソース列/式, 状態, 備考)"""
    if h in SKIP_HEADERS:
        return "—", "—", "SKIP", "式・人間・K1-K6"
    if h in FIELD_RULES:
        src, expr, note = FIELD_RULES[h]
        st = "OK"
        if h in DIRECT_WRITE_HEADERS and i in formula_cols:
            st = "OK_DIRECT"
            note = (note + "；" if note else "") + "行5は式だが直書き"
        elif h == "Keepaグラフ":
            st = "OPTIONAL"
        return src, expr, st, note
    if h in MAP:
        if i in formula_cols and h not in DIRECT_WRITE_HEADERS:
            return "Keepaフル→式", MAP[h], "FORMULA", "行5式コピー（直書きしない）"
        return "Keepaフル", MAP[h], "OK", ""
    if i in formula_cols:
        return "—", "式（行コピー）", "FORMULA", "上書きしない"
    if h in CAND_FALLBACK:
        return "①候補(fallback)", CAND_FALLBACK[h], "FALLBACK", "Keepaフル miss 時のみ"
    if not h:
        return "—", "—", "EMPTY", "ヘッダ空列"
    return "—", "—", "EMPTY", "転記しない（空のまま）"


def mapping_rows(
    list_headers: list[str],
    formula_cols: set[int],
    sample_rec: dict[str, str],
) -> list[list[str]]:
    out = []
    for i, h in enumerate(list_headers):
        src, expr, st, note = classify_col(h, i, formula_cols)
        samp = sample_rec.get(h, "")
        if len(samp) > 80:
            samp = samp[:77] + "..."
        out.append([h, src, expr, st, note, samp])
    return out


def build_sample_rec(
    list_headers: list[str],
    full: dict,
    cand_row: dict,
    formula_cols: set[int],
) -> dict[str, str]:
    row = build_row(list_headers, full, formula_cols, cand_row)
    rec = {list_headers[i]: row[i] if i < len(row) else "" for i in range(len(list_headers))}
    # 検証用にソース側も記録
    rec["_maker_cand"] = list_maker_name(cand_row, full)
    rec["_min_price"] = min_cart_new_price(full)
    rec["_rank90"] = str(full.get("売れ筋ランキング: 90 日平均") or "")
    rec["_rank_disp"] = rank_display(full)
    rec["_cat"] = category_name(full)
    rec["_keepa_url"] = str(full.get("URL: Keepa") or "")
    rec["_amz_url"] = str(full.get("URL: Amazon") or "")
    rec["_bb"] = str(full.get("Buy Box: 現在価格") or "")
    rec["_new_now"] = str(full.get("新品: 現在価格") or "")
    return rec


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply-mapping", action="store_true", help="マッピング仮シートを書く")
    ap.add_argument("--cleanup", action="store_true", help="マッピング仮シートを削除")
    args = ap.parse_args()

    if args.cleanup:
        svc = sheets_service(write=True)
        ok = delete_map_sheet(svc)
        print("cleanup", "deleted" if ok else "not_found")
        return 0

    svc = sheets_service(write=True)
    if RESEARCH_SS == ORIG_SS:
        print("FAIL refuse original SS")
        return 2

    rng_h = "'%s'!%d:%d" % (LIST_TITLE.replace("'", "''"), HEADER_ROW, HEADER_ROW)
    list_headers = [
        str(x).replace("\n", " ").strip()
        for x in (
            svc.spreadsheets().values().get(spreadsheetId=RESEARCH_SS, range=rng_h).execute().get("values") or [[]]
        )[0]
    ]
    frng = "'%s'!A%d:%s%d" % (LIST_TITLE.replace("'", "''"), DATA_START, col_letter(len(list_headers)), DATA_START)
    frow = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RESEARCH_SS, range=frng, valueRenderOption="FORMULA")
        .execute()
        .get("values")
        or [[]]
    )[0]
    formula_cols = {i for i, v in enumerate(frow) if str(v).startswith("=")}
    formula_names = [list_headers[i] for i in sorted(formula_cols) if i < len(list_headers)]

    ch, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    pass_rows = []
    seen = set()
    cand_by_asin: dict[str, dict] = {}
    for r in cand:
        a = str(r.get("ASIN") or "").strip().upper()
        if a and str(r.get("門結果") or "") == "通過" and a not in seen:
            seen.add(a)
            pass_rows.append(r)
            cand_by_asin[a] = r

    raw_list = read_all(svc, RESEARCH_SS, LIST_TITLE)
    asin_i = list_headers.index("ASIN") if "ASIN" in list_headers else -1
    exist = set()
    if asin_i >= 0:
        for r in raw_list[DATA_START - 1 :]:
            if asin_i < len(r):
                a = str(r[asin_i] or "").strip().upper()
                if len(a) == 10:
                    exist.add(a)

    fh, frows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    fullmap = latest_full(frows)
    new_asins = [str(r.get("ASIN") or "").upper() for r in pass_rows if str(r.get("ASIN") or "").upper() not in exist]
    have_full = [a for a in new_asins if a in fullmap]
    miss_full = [a for a in new_asins if a not in fullmap]

    sample_asin = have_full[0] if have_full else (pass_rows[0].get("ASIN") if pass_rows else None)
    sample_asin = str(sample_asin or "").strip().upper()
    sample_full = fullmap.get(sample_asin, {})
    sample_cand = cand_by_asin.get(sample_asin, pass_rows[0] if pass_rows else {})
    sample_rec = (
        build_sample_rec(list_headers, sample_full, sample_cand, formula_cols) if sample_asin and sample_full else {}
    )

    mrows = mapping_rows(list_headers, formula_cols, sample_rec)
    ok_states = frozenset({"OK", "OK_DIRECT", "OPTIONAL", "FORMULA", "SKIP", "FALLBACK"})
    ok_map = sum(1 for r in mrows if r[3] in ok_states)

    user_checks: dict[str, dict] = {}
    if sample_rec:
        user_checks = {
            "メーカー名": {
                "ok": bool(sample_rec.get("メーカー名")),
                "expect": "①候補.メーカー",
                "got": sample_rec.get("メーカー名", ""),
                "src": sample_rec.get("_maker_cand", ""),
            },
            "最安値": {
                "ok": bool(sample_rec.get("最安値")),
                "expect": "min(BB,新品)",
                "got": sample_rec.get("最安値", ""),
                "src": "BB=%s NEW=%s" % (sample_rec.get("_bb", ""), sample_rec.get("_new_now", "")),
            },
            "ランキング": {
                "ok": bool(sample_rec.get("ランキング")) and "位" in str(sample_rec.get("ランキング", "")),
                "expect": "90日平均+カテゴリ",
                "got": sample_rec.get("ランキング", ""),
                "src": "rank90=%s cat=%s" % (sample_rec.get("_rank90", ""), sample_rec.get("_cat", "")),
            },
            "順位": {
                "ok": sample_rec.get("順位") == sample_rec.get("_rank90"),
                "expect": "売れ筋ランキング: 90 日平均",
                "got": sample_rec.get("順位", ""),
                "src": sample_rec.get("_rank90", ""),
            },
            "カテゴリー": {
                "ok": bool(sample_rec.get("カテゴリー")),
                "expect": "カテゴリ: ルート",
                "got": sample_rec.get("カテゴリー", ""),
                "src": sample_rec.get("_cat", ""),
            },
            "Keepaグラフ": {
                "ok": True,
                "expect": "URL: Keepa（任意）",
                "got": sample_rec.get("Keepaグラフ", "") or "(空)",
                "src": sample_rec.get("_keepa_url", "") or "(空)",
            },
            "amazonページ": {
                "ok": bool(sample_rec.get("amazonページ")),
                "expect": "URL: Amazon",
                "got": sample_rec.get("amazonページ", ""),
                "src": sample_rec.get("_amz_url", ""),
            },
        }

    user_ok = all(user_checks[k]["ok"] for k in USER_CRITICAL if k in user_checks)
    critical_data_ok = all(
        any(r[0] == k and r[3] in ("OK", "OK_DIRECT", "OPTIONAL") for r in mrows)
        for k in ("メーカー名", "ASIN", "商品名", "最安値", "出品者数")
        if k in list_headers
    )

    d1_ok = (
        RESEARCH_SS != ORIG_SS
        and "ASIN" in list_headers
        and asin_i >= 0
        and len(formula_cols) >= 5
        and len(pass_rows) >= 1
        and critical_data_ok
        and user_ok
        and (not new_asins or len(have_full) == len(new_asins))
    )

    d2_present = [c for c in D2_FILTER_COLS if c in list_headers]
    d2_formula = [c for c in ("粗利益率", "1点利益額", "目標仕入額(税込み）") if c in formula_names]
    d2_ok = len(d2_present) == len(D2_FILTER_COLS) and len(d2_formula) >= 1

    print("==== D1 仮想 ====")
    print("list_headers", len(list_headers), "formula_cols", len(formula_cols))
    print("cand_pass", len(pass_rows), "already_in_list", len(exist & seen), "new", len(new_asins))
    print("new_have_full", len(have_full), "new_miss_full", miss_full[:5])
    print("map_rows resolved", ok_map, "/", len(mrows))
    print("critical_data_ok", critical_data_ok)
    print("sample_asin", sample_asin or "(none)")
    for k in USER_CRITICAL:
        if k in user_checks:
            c = user_checks[k]
            print("check_%s" % k, "OK" if c["ok"] else "FAIL", c["got"][:60], "|", c["src"][:60])
    print("user_spec_ok", user_ok)
    print("D1", "PASS" if d1_ok else "FAIL", "miss_full_block" if miss_full else "")

    print("==== D2 仮想 ====")
    print("filter_cols_present", len(d2_present), "/", len(D2_FILTER_COLS))
    print("profit_formula_cols", d2_formula)
    print("D2", "PASS" if d2_ok else "FAIL")

    empty_cols = [r[0] for r in mrows if r[3] == "EMPTY"]
    print("empty_cols", len(empty_cols), empty_cols[:5])

    overall = d1_ok and d2_ok
    print("==== OVERALL ====", "PASS" if overall else "FAIL")

    if args.apply_mapping:
        ensure_map_sheet(svc, mrows)
        print("mapping_sheet", MAP_SHEET, "rows", len(mrows) + 1)
        print("url", "https://docs.google.com/spreadsheets/d/%s/edit#gid=0" % RESEARCH_SS)
        print("NOTE sheet kept for human review. cleanup: python dry_d1_d2_virtual.py --cleanup")

    return 0 if overall else 1


if __name__ == "__main__":
    raise SystemExit(main())
