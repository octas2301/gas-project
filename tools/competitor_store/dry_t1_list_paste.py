# -*- coding: utf-8 -*-
"""T1 dry: 仕入検討品番リストへ Keepaフル転記計画。式列・原本は触らない。既定非書。"""
from __future__ import annotations

import json
import re
from datetime import date

from apply_keepa_full import COMPETITOR_SS, RESEARCH_SS, T_CAND, append_log, as_dicts, read_all, sheets_service
from schema import SHEET_KEEPA_FULL

LIST_TITLE = "仕入検討品番リスト"
HEADER_ROW = 4  # 1-based
DATA_START = 5
ORIG_SS = "1Uk-DEfoSTVqb5N8XsfSqciaZcIjzbbPpGvg4hQYgDl4"

# 品番リスト列名 → Keepaフル列。式・人間・K1-K6は載せない。
MAP = {
    "メーカー名": "製造者",
    "ASIN": "ASIN",
    "商品名": "商品名",
    "JANコード": "商品コード: EAN",
    "出品者数": "出品者数",
    "画像URL": "画像",
    "Keepaグラフ": "URL: Keepa",
    "amazonページ": "URL: Amazon",
    "最安値": "新品: 現在価格",
    "FBA手数料": "出品FBA手数料",
    "amazon本体が出品者にいる": "Amazon直販",
    "順位": "売れ筋ランキング: 現在",
    "カテゴリー": "カテゴリ: ルート",
    "クイックショップサイズ表記": "出品FBAティア",
    "自己配送 送料概算": "自己発送送料",
    "サイズ3辺合計": "梱包3辺合計_cm",
}

SKIP_HEADERS = {
    "ASIN重複チェック",
    "調査員名",
    "メーカーがマイナーか有名か",
    "商品画像",
    "販売形態",
    "同商品のセット数量",
    "単品商品名",
    "JAN調査プロンプト",
    "AI調査プロンプト",
    "FBA出品者",
    "メーカーが出品者にいる",
    "1点当たり メーカー定価(税込)",
    "価格が安定しているか",
    "卸商品リスト卸値（税抜き）",
    "税率",
    "卸商品リスト卸値（税込み）",
    "卸値一覧（あれば）sheet商品名VLOOK",
    "卸値一覧（あれば）sheet卸値税抜き",
    "卸値一覧（あれば）sheet卸値税込み",
    "最低利益額",
    "1点販売額",
    "1点FBA手数料",
    "目標仕入額(税込み）",
    "目標仕入額(税抜き）",
    "最低目標掛率",
    "定価と比べてama価格は",
    "割引率",
    "1点仕入れ額（税込み）",
    "1点損益分岐額",
    "1点利益額",
    "粗利益率",
    "順位10000単位",
    "順位対象外",
    "調査年月",
    "件数",
    "調査追加時",
    "合計調査件数",
    "報酬額",
    "単品商品名２ＰＶ用",
    "計算",
    "nn",
}


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def latest_full(rows: list[dict]) -> dict[str, dict]:
    out = {}
    for r in rows:
        a = str(r.get("ASIN") or "").strip().upper()
        if a:
            out[a] = r
    return out


def pack_size(full: dict) -> str:
    l, w, h = full.get("梱包_L_cm"), full.get("梱包_W_cm"), full.get("梱包_H_cm")
    if not (l and w and h):
        try:
            p = json.loads(full.get("生JSON") or "{}")
            def cm(v):
                try:
                    n = float(v)
                    return n / 10.0 if n > 0 else None
                except (TypeError, ValueError):
                    return None
            l, w, h = cm(p.get("packageLength")), cm(p.get("packageWidth")), cm(p.get("packageHeight"))
        except json.JSONDecodeError:
            return ""
    if not (l and w and h):
        return ""
    try:
        return "%s x %s x %scm" % (round(float(l), 1), round(float(w), 1), round(float(h), 1))
    except (TypeError, ValueError):
        return ""


def pack_weight(full: dict) -> str:
    g = full.get("梱包_重量_g")
    if g:
        return str(g) + "g" if not str(g).endswith("g") else str(g)
    try:
        p = json.loads(full.get("生JSON") or "{}")
        w = p.get("packageWeight")
        if w not in (None, "", -1):
            return str(w) + "g"
    except json.JSONDecodeError:
        return ""
    return ""


def category_name(full: dict) -> str:
    cat = str(full.get("カテゴリ: ルート") or "").strip()
    if cat:
        return cat
    try:
        p = json.loads(full.get("生JSON") or "{}")
        tree = p.get("categoryTree") or []
        if tree and isinstance(tree, list) and isinstance(tree[0], dict):
            return str(tree[0].get("name") or "")
    except json.JSONDecodeError:
        return ""
    return ""


def rank_display(full: dict) -> str:
    r = str(full.get("売れ筋ランキング: 現在") or "").strip()
    cat = category_name(full)
    if r and cat:
        try:
            return "{:,}位 | {}".format(int(float(r)), cat)
        except ValueError:
            return r + " | " + cat
    if r:
        try:
            return "{:,}位".format(int(float(r)))
        except ValueError:
            return r
    return r


def build_row(headers: list[str], full: dict, formula_cols: set[int]) -> list[str]:
    rec = {h: "" for h in headers}
    rec["調査日"] = date.today().isoformat()
    rec["メーカー名"] = full.get("製造者") or full.get("ブランド") or ""
    rec["ASIN"] = full.get("ASIN") or ""
    rec["ランキング"] = rank_display(full)
    rec["過去1か月販売数（表示あれば）"] = full.get("月間売上") or ""
    rec["商品名"] = full.get("商品名") or ""
    rec["JANコード"] = full.get("商品コード: EAN") or ""
    rec["サイズ"] = pack_size(full)
    rec["重量"] = pack_weight(full)
    rec["出品者数"] = full.get("出品者数") or ""
    rec["画像URL"] = full.get("画像") or ""
    rec["Keepaグラフ"] = full.get("URL: Keepa") or ""
    rec["amazonページ"] = full.get("URL: Amazon") or ""
    rec["最安値"] = full.get("新品: 現在価格") or ""
    rec["FBA手数料"] = full.get("出品FBA手数料") or ""
    rec["amazon本体が出品者にいる"] = full.get("Amazon直販") or ""
    rec["順位"] = full.get("売れ筋ランキング: 現在") or ""
    rec["カテゴリー"] = category_name(full)
    rec["クイックショップサイズ表記"] = full.get("出品FBAティア") or ""
    out = []
    for i, h in enumerate(headers):
        if i in formula_cols or h in SKIP_HEADERS:
            out.append("")
            continue
        out.append(str(rec.get(h) or ""))
    return out


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    if RESEARCH_SS == ORIG_SS:
        print("refuse original SS")
        return 2

    rng_h = "'%s'!%d:%d" % (LIST_TITLE.replace("'", "''"), HEADER_ROW, HEADER_ROW)
    headers = [str(x).replace("\n", " ").strip() for x in (
        svc.spreadsheets().values().get(spreadsheetId=RESEARCH_SS, range=rng_h).execute().get("values") or [[]]
    )[0]]
    # formulas on first data row
    frng = "'%s'!A%d:%s%d" % (LIST_TITLE.replace("'", "''"), DATA_START, col_letter(len(headers)), DATA_START)
    frow = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RESEARCH_SS, range=frng, valueRenderOption="FORMULA")
        .execute()
        .get("values")
        or [[]]
    )[0]
    formula_cols = set()
    for i, v in enumerate(frow):
        if str(v).startswith("="):
            formula_cols.add(i)
    mapped = [h for h in headers if h in MAP or h in ("調査日", "ランキング", "サイズ", "重量", "過去1か月販売数（表示あれば）", "梱包サイズ")]
    mapped = list(dict.fromkeys(mapped))

    ch, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    pass_asins = []
    seen = set()
    for r in cand:
        a = str(r.get("ASIN") or "").strip().upper()
        st = str(r.get("門結果") or "")
        if a and st == "通過" and a not in seen:
            seen.add(a)
            pass_asins.append(a)

    raw_list = read_all(svc, RESEARCH_SS, LIST_TITLE)
    exist = set()
    for r in raw_list[DATA_START - 1 :]:
        if len(r) > 5:
            a = str(r[5] if headers[5] == "ASIN" else "").strip().upper()
            # find ASIN col
    asin_i = headers.index("ASIN") if "ASIN" in headers else 5
    for r in raw_list[DATA_START - 1 :]:
        if asin_i < len(r):
            a = str(r[asin_i] or "").strip().upper()
            if len(a) == 10:
                exist.add(a)

    fh, frows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    fullmap = latest_full(frows)
    new_asins = [a for a in pass_asins if a not in exist]
    have_full = [a for a in new_asins if a in fullmap]
    miss_full = [a for a in new_asins if a not in fullmap]
    already = [a for a in pass_asins if a in exist]

    print("ss", RESEARCH_SS[:8], "orig_guard", RESEARCH_SS != ORIG_SS)
    print("headers", len(headers), "formula_cols", sorted(formula_cols)[:20], "n_formula", len(formula_cols))
    print("formula_names", [headers[i] for i in sorted(formula_cols) if i < len(headers)][:25])
    print("cand_pass", len(pass_asins), "already_in_list", len(already), "new", len(new_asins), "new_with_full", len(have_full), "new_miss_full", len(miss_full))
    print("map_headers", mapped)
    sample = have_full[0] if have_full else None
    if sample:
        row = build_row(headers, fullmap[sample], formula_cols)
        nonempty = [(headers[i], row[i][:40]) for i in range(len(headers)) if row[i]]
        print("sample", sample, nonempty[:12])

    ok = (
        RESEARCH_SS != ORIG_SS
        and len(formula_cols) >= 5
        and "ASIN" in headers
        and len(pass_asins) >= 1
        and "報酬額" in SKIP_HEADERS
        and "FBA出品者" in SKIP_HEADERS
    )
    line = (
        "runId=pr_20260815_t1dry pass=%d already=%d append_plan=%d miss_full=%d formula_cols=%d GETなし %s"
        % (len(pass_asins), len(already), len(have_full), len(miss_full), len(formula_cols), "PASS" if ok else "FAIL")
    )
    print(line)
    if not args.apply:
        append_log(svc, "T1", line)
        return 0 if ok else 1
    if not ok:
        print("skip apply dry FAIL")
        return 1
    if not have_full:
        print("nothing to append")
        append_log(svc, "T1", "runId=pr_20260815_t1col append=0")
        return 0
    meta = svc.spreadsheets().get(spreadsheetId=RESEARCH_SS).execute()
    sid = None
    for sh in meta.get("sheets", []):
        if sh["properties"]["title"] == LIST_TITLE:
            sid = sh["properties"]["sheetId"]
            break
    if sid is None:
        print("no sheet id")
        return 2
    last = len(raw_list)
    start = max(last + 1, DATA_START)
    n = len(have_full)
    props = None
    for sh in meta.get("sheets", []):
        if sh["properties"]["title"] == LIST_TITLE:
            props = sh["properties"]
            break
    cur_rows = int((props.get("gridProperties") or {}).get("rowCount") or last)
    need_rows = start - 1 + n
    reqs = []
    if need_rows > cur_rows:
        reqs.append(
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sid,
                        "gridProperties": {"rowCount": need_rows},
                    },
                    "fields": "gridProperties.rowCount",
                }
            }
        )
    reqs.append(
        {
            "copyPaste": {
                "source": {
                    "sheetId": sid,
                    "startRowIndex": DATA_START - 1,
                    "endRowIndex": DATA_START,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(headers),
                },
                "destination": {
                    "sheetId": sid,
                    "startRowIndex": start - 1,
                    "endRowIndex": start - 1 + n,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(headers),
                },
                "pasteType": "PASTE_FORMULA",
                "pasteOrientation": "NORMAL",
            }
        }
    )
    svc.spreadsheets().batchUpdate(spreadsheetId=RESEARCH_SS, body={"requests": reqs}).execute()
    values = [build_row(headers, fullmap[a], formula_cols) for a in have_full]
    data_cols = [i for i in range(len(headers)) if i not in formula_cols and headers[i] not in SKIP_HEADERS]
    data = []
    for ridx, row in enumerate(values):
        for c in data_cols:
            if not row[c]:
                continue
            data.append(
                {
                    "range": "'%s'!%s%d" % (LIST_TITLE.replace("'", "''"), col_letter(c + 1), start + ridx),
                    "values": [[row[c]]],
                }
            )
    for i in range(0, len(data), 400):
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=RESEARCH_SS,
            body={"valueInputOption": "USER_ENTERED", "data": data[i : i + 400]},
        ).execute()
    blanks = []
    skip_i = [i for i, name in enumerate(headers) if name in SKIP_HEADERS or name == "調査員名"]
    for ridx in range(n):
        for c in skip_i:
            blanks.append(
                {
                    "range": "'%s'!%s%d" % (LIST_TITLE.replace("'", "''"), col_letter(c + 1), start + ridx),
                    "values": [[""]],
                }
            )
    for i in range(0, len(blanks), 400):
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=RESEARCH_SS,
            body={"valueInputOption": "RAW", "data": blanks[i : i + 400]},
        ).execute()
    line2 = "runId=pr_20260815_t1col append=%d skip_exist=%d formula_untouched=%d" % (
        len(values),
        len(already),
        len(formula_cols),
    )
    append_log(svc, "T1", line2)
    print(line2)
    # verify last appended ASINs
    raw2 = read_all(svc, RESEARCH_SS, LIST_TITLE)
    got = set()
    for r in raw2[DATA_START - 1 :]:
        if asin_i < len(r) and len(str(r[asin_i])) == 10:
            got.add(str(r[asin_i]).strip().upper())
    hit = sum(1 for a in have_full if a in got)
    print("verify_appended", hit, "/", len(have_full))
    vok = hit == len(have_full)
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
