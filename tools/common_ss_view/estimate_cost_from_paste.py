# -*- coding: utf-8 -*-
"""貼付sheetの税抜卸値 → 品番リスト「1点仕入れ額（税込み）」へ戻す（H前段）。

正（FLOW §5.3.2 / 見積検品要件）:
- 鍵の優先: JAN一致（完全一致）→ 商品名部分一致（要確認・`--allow-name-partial`）
- 税込 = floor(税抜 × (1+税率))
- 既定は見積候補のみ。`--include-skip` で見送りも
- `--makers` でメーカー完全一致に限定
- 既に仕入がある行は上書きしない
- 貼付・依頼リスト・検討は触らない
- 部分一致が複数候補のときは転記しない（ambiguous）

使い方:
  python tools/common_ss_view/estimate_cost_from_paste.py --dry
  python tools/common_ss_view/estimate_cost_from_paste.py
  python tools/common_ss_view/estimate_cost_from_paste.py --makers オリエンタル,株式会社オリエンタル,オリエンタル洋行 --include-skip --allow-name-partial
"""
from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "competitor_store"))
sys.path.insert(0, str(ROOT))
from client import sheets_service  # noqa: E402
import estimate_paste_match as pm  # noqa: E402

RES = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
SANK = "17nXSRjFp8xy-3EaBmp-LIu_tMc9ZOgfkhJoZjO3FLIc"
LIST = "仕入検討品番リスト"
PASTE = "貼付sheet"


def num(v):
    return pm.num(v)


def jan_norm(v) -> str:
    return pm.jan_norm(v)


def col_letter(n1: int) -> str:
    n = n1
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def tax_incl_from_ex(ex: float, tax: float) -> int:
    """税抜×(1+税率)を切り下げ（FLOW §5.3.2）。"""
    return int(math.floor(ex * (1.0 + tax) + 1e-9))


def load_paste_rows(svc) -> list[dict]:
    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=SANK,
            range=f"'{PASTE}'!A3:G",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    return pm.paste_rows_from_values(rows)


def collect_targets(
    svc,
    paste_rows: list[dict],
    *,
    makers_filter: list[str] | None,
    include_skip: bool,
    allow_name_partial: bool,
):
    paste_map = pm.paste_by_jan(paste_rows)
    hdr = [
        str(c).strip()
        for c in (
            (
                svc.spreadsheets()
                .values()
                .get(spreadsheetId=RES, range=f"'{LIST}'!4:4")
                .execute()
                .get("values")
                or [[]]
            )[0]
        )
    ]
    idx = {h: i for i, h in enumerate(hdr)}
    need = [
        "メーカー名",
        "単品商品名",
        "商品名",
        "JANコード",
        "1点仕入れ額（税込み）",
        "見積候補",
        "税率",
    ]
    for n in need:
        if n not in idx:
            raise SystemExit(f"missing col: {n}")

    i_m, i_n, i_t = idx["メーカー名"], idx["単品商品名"], idx["商品名"]
    i_j = idx["JANコード"]
    i_c, i_cand, i_tax = idx["1点仕入れ額（税込み）"], idx["見積候補"], idx["税率"]
    last = max(i_m, i_n, i_t, i_j, i_c, i_cand, i_tax)
    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{LIST}'!A5:{col_letter(last + 1)}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )

    targets = []
    skip = {
        "not_flag": 0,
        "maker": 0,
        "bad_name": 0,
        "no_match": 0,
        "ambiguous": 0,
        "has_cost": 0,
        "no_tax": 0,
    }
    for i, r in enumerate(rows):
        while len(r) <= last:
            r.append("")
        flag = str(r[i_cand]).strip()
        if flag == "見積候補":
            pass
        elif include_skip and flag == "見送り":
            pass
        else:
            skip["not_flag"] += 1
            continue
        maker = str(r[i_m]).strip()
        if makers_filter and maker not in makers_filter:
            skip["maker"] += 1
            continue
        name = str(r[i_n]).strip()
        title = str(r[i_t]).strip()
        if (not name or "要確認" in name) and not title:
            skip["bad_name"] += 1
            continue
        if not name or "要確認" in name:
            name = title[:40]
        j = jan_norm(r[i_j])
        match = pm.resolve_paste_match(
            paste_rows=paste_rows,
            paste_map=paste_map,
            jan=j,
            maker=maker,
            name=name,
            title=title,
            makers_filter=makers_filter,
            allow_name_partial=allow_name_partial,
        )
        if match["kind"] == pm.MATCH_AMBIG:
            skip["ambiguous"] += 1
            continue
        if match["kind"] == pm.MATCH_NONE or match.get("ex") is None:
            skip["no_match"] += 1
            continue
        cost = num(r[i_c])
        if cost and cost > 0:
            skip["has_cost"] += 1
            continue
        tax = num(r[i_tax])
        if tax is None or tax < 0:
            skip["no_tax"] += 1
            continue
        sheet_row = i + 5
        incl = tax_incl_from_ex(float(match["ex"]), tax)
        targets.append(
            {
                "row": sheet_row,
                "maker": maker,
                "name": name,
                "jan": j,
                "paste_jan": match["jan"],
                "ex": match["ex"],
                "tax": tax,
                "incl": incl,
                "col": i_c + 1,
                "match_kind": match["kind"],
                "match_label": match["label"],
                "paste_name": match.get("paste_name") or "",
            }
        )
    return targets, skip, idx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    ap.add_argument(
        "--makers",
        default="",
        help="カンマ区切りメーカー名（完全一致）。指定時はそのメーカーだけ",
    )
    ap.add_argument(
        "--include-skip",
        action="store_true",
        help="見送り行も仕入転記の対象にする",
    )
    ap.add_argument(
        "--allow-name-partial",
        action="store_true",
        help="JAN不一致時に商品名部分一致で転記（完全一致ではない・要確認）",
    )
    args = ap.parse_args()
    makers_filter = [m.strip() for m in args.makers.split(",") if m.strip()] or None

    svc = sheets_service(write=not args.dry, interactive=False)
    paste_rows = load_paste_rows(svc)
    print(f"paste priced rows={len(paste_rows)}")
    targets, skip, _idx = collect_targets(
        svc,
        paste_rows,
        makers_filter=makers_filter,
        include_skip=args.include_skip,
        allow_name_partial=args.allow_name_partial,
    )
    print("skip", skip)
    print(f"targets={len(targets)}")
    by_kind = {}
    for t in targets:
        by_kind[t["match_kind"]] = by_kind.get(t["match_kind"], 0) + 1
        print(
            f"  r{t['row']} [{t['match_kind']}] {t['maker']} | {t['name'][:36]} | "
            f"listJAN={t['jan'] or '-'} pasteJAN={t['paste_jan']} "
            f"| ex={t['ex']} -> incl={t['incl']} | {t['match_label']}"
        )
    print("by_kind", by_kind)

    if args.dry:
        print("DRY: no write")
        return

    if not targets:
        print("nothing to write")
        return

    data = []
    for t in targets:
        letter = col_letter(t["col"])
        data.append(
            {
                "range": f"'{LIST}'!{letter}{t['row']}",
                "values": [[t["incl"]]],
            }
        )
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=RES,
        body={"valueInputOption": "USER_ENTERED", "data": data},
    ).execute()
    print(f"wrote {len(data)} cells to {LIST} 1点仕入れ額（税込み）")


if __name__ == "__main__":
    main()
