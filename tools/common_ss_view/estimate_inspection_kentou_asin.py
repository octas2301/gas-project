# -*- coding: utf-8 -*-
"""検討の別ASINを Keepaフルから対象セットへ追記／更新。

正:
- 月間売上・大カテゴリ順位を倉庫から転記（ハードコードの「-」禁止）
- 索引行の ASIN 列が一致するときだけ採用
- 既に（検討ASIN）行があるときは上書き

使い方:
  python tools/common_ss_view/estimate_inspection_kentou_asin.py --dry
  python tools/common_ss_view/estimate_inspection_kentou_asin.py
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "competitor_store"))
from client import sheets_service
import estimate_inspection_build_rows as b

RES = b.RES
SANK = "17nXSRjFp8xy-3EaBmp-LIu_tMc9ZOgfkhJoZjO3FLIc"
COMP = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
SHEET = b.SHEET


def asin_from_url(url: str) -> str:
    m = re.search(r"/(?:dp|gp/product)/([A-Z0-9]{10})", str(url or ""), re.I)
    return m.group(1).upper() if m else ""


def load_kentou_jan_asin(svc) -> dict[str, str]:
    rows = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=SANK, range="'検討'!A4:AZ3000")
        .execute()
        .get("values")
        or []
    )
    if not rows:
        return {}
    hdr = [str(c).replace("\n", "/") for c in rows[0]]
    idx = {h: i for i, h in enumerate(hdr)}
    out = {}
    for r in rows[1:]:
        while len(r) < len(hdr):
            r.append("")
        j = b.jan_norm(r[idx["JANコード"]] if "JANコード" in idx else "")
        a = asin_from_url(r[idx["URL"]] if "URL" in idx else "")
        if j and a:
            out[j] = a
    return out


def load_keepa_index(svc) -> dict[str, int]:
    out = {}
    for r in (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMP, range="'Keepaフル索引'!A:B")
        .execute()
        .get("values")
        or []
    )[1:]:
        if len(r) < 2:
            continue
        try:
            out[str(r[0]).strip().upper()] = int(float(r[1]))
        except Exception:
            continue
    return out


def keepa_cols(kh: list) -> dict:
    kidx = {str(h).strip(): i for i, h in enumerate(kh)}
    rank_c = kidx.get("売れ筋ランキング: 現在")
    if rank_c is None:
        for h, i in kidx.items():
            if "大カテゴリ" in h and "順位" in h:
                rank_c = i
                break
    return {
        "asin": kidx.get("ASIN"),
        "title": kidx.get("商品名") if kidx.get("商品名") is not None else kidx.get("タイトル"),
        "sold": kidx.get("月間売上"),
        "rank": rank_c,
        "price": kidx.get("新品: 現在価格")
        if kidx.get("新品: 現在価格") is not None
        else kidx.get("Amazon: 現在価格"),
    }


def read_keepa_asin(svc, amap: dict[str, int], cols: dict, asin: str) -> dict | None:
    """索引→行を開き、ASIN列が一致するときだけ返す。"""
    grow = amap.get(asin)
    if not grow or cols["asin"] is None:
        return None
    need = [c for c in cols.values() if c is not None]
    letter = b.col_letter(max(need) + 1)
    vals = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=COMP,
            range=f"'Keepaフル'!A{grow}:{letter}{grow}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or [[]]
    )[0]
    while len(vals) <= max(need):
        vals.append("")
    got = str(vals[cols["asin"]] or "").strip().upper()
    if got != asin:
        print(f"skip index mismatch {asin} row={grow} got={got}")
        return None
    title = str(vals[cols["title"]] or "") if cols["title"] is not None else ""
    sold = b.num(vals[cols["sold"]]) if cols["sold"] is not None else None
    rank = None
    if cols["rank"] is not None:
        try:
            rv = vals[cols["rank"]]
            if rv not in (None, ""):
                rank = int(float(rv))
        except Exception:
            rank = None
    price = b.num(vals[cols["price"]]) if cols["price"] is not None else None
    return {"title": title, "sold": sold, "rank": rank, "price": price}


def format_kentou_line(asin: str, facts: dict) -> str | None:
    rank = facts.get("rank")
    if rank and rank > b.RANK_CAP:
        print("skip rank cap", asin, rank)
        return None
    rank_s = f"{rank:,}位" if rank else "順位-"
    sold = facts.get("sold")
    month_s = f"月販{int(sold)}" if sold and sold > 0 else "月販-"
    price = facts.get("price")
    price_s = f"{b.yen(price)}円" if price else "-円"
    title = facts.get("title") or ""
    m = re.search(r"[×xXｘ](\d+)\s*袋", title)
    q = int(m.group(1)) if m else None
    kind = f"単品{q}袋" if q else "単品?袋"
    # 検討ASINは品番リスト外が多い → FBA利益は計算せず -
    return f"{kind}（{price_s}）／{rank_s}／FBA利益-／{month_s}／{asin}（検討ASIN）"


def upsert_kentou_in_text(cur: str, asin: str, line: str) -> str:
    lines = [ln for ln in str(cur or "").split("\n") if ln.strip()]
    out = []
    replaced = False
    for ln in lines:
        if asin in ln and "検討ASIN" in ln:
            out.append(line)
            replaced = True
        elif asin in ln and re.search(rf"\b{re.escape(asin)}\b", ln):
            # 同一ASINが本線側にあれば検討行は付けない（本線優先）
            return "\n".join(lines)
        else:
            out.append(ln)
    if not replaced:
        out.append(line)
    return "\n".join(out)


def asins_from_text(text: str) -> list[str]:
    out = []
    for ln in str(text or "").split("\n"):
        m = re.search(r"(B[0-9A-Z]{9})", ln)
        if m:
            out.append(m.group(1))
    return out


def apply_kentou(svc, dry: bool = False) -> int:
    jan_asin = load_kentou_jan_asin(svc)
    amap = load_keepa_index(svc)
    kh = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMP, range="'Keepaフル'!1:1")
        .execute()
        .get("values")
        or [[]]
    )[0]
    cols = keepa_cols(kh)
    print("keepa cols", cols)

    insp = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{SHEET}'!A1:Z")
        .execute()
        .get("values")
        or []
    )
    if len(insp) < 2:
        print("no rows")
        return 0
    hdr = [str(c).strip() for c in insp[0]]
    try:
        col_target = next(i for i, h in enumerate(hdr) if "対象セット" in h) + 1
    except StopIteration:
        print("no 対象セット col")
        return 0
    b.COL_L = col_target

    paste = b.load_paste_meta(svc)
    catalog = b.load_list(svc)
    key_jan = {}
    for r in catalog:
        if r.get("yokakunin") or not r.get("cost") or not r.get("jan"):
            continue
        key_jan[f"{r['name']}／仕入{int(round(r['cost']))}"] = r["jan"]

    updates = []
    for i, row in enumerate(insp[1:], start=2):
        while len(row) < col_target:
            row.append("")
        key = str(row[2])
        base, _cost = b.parse_key(key)
        j = key_jan.get(key)
        if not j:
            for kk, jj in key_jan.items():
                if base and base in kk:
                    j = jj
                    break
        if not j or j not in jan_asin:
            continue
        alt = jan_asin[j]
        facts = read_keepa_asin(svc, amap, cols, alt)
        if not facts:
            print("no keepa", key[:28], alt)
            continue
        line = format_kentou_line(alt, facts)
        if not line:
            continue
        cur = str(row[col_target - 1] or "")
        # 本線に同じASINがあるなら検討追記不要
        if alt in cur and "検討ASIN" not in cur:
            print("already main", key[:28], alt)
            continue
        new_text = upsert_kentou_in_text(cur, alt, line)
        if new_text == cur:
            print("unchanged", key[:28], alt)
            continue
        decision = str(row[4]).strip() if len(row) > 4 else ""
        updates.append(
            {
                "row": i,
                "text": new_text,
                "asins": asins_from_text(new_text),
                "gray": decision == "見送り",
                "key": key,
                "alt": alt,
                "line": line,
            }
        )
        print("upd", key[:28], line)

    print("updates", len(updates))
    if dry or not updates:
        return len(updates)

    meta = (
        svc.spreadsheets()
        .get(spreadsheetId=RES, ranges=[f"'{SHEET}'"], fields="sheets(properties(sheetId))")
        .execute()
    )
    sid = meta["sheets"][0]["properties"]["sheetId"]
    b.apply_l_links_batch(svc, sid, updates)
    print("WROTE", len(updates))
    return len(updates)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=not args.dry, interactive=False)
    n = apply_kentou(svc, dry=args.dry)
    print("DONE", n, "dry" if args.dry else "wrote")


if __name__ == "__main__":
    main()
