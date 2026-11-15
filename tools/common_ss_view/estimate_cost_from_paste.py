# -*- coding: utf-8 -*-
"""貼付sheetの税抜卸値 → 品番リスト「1点仕入れ額（税込み）」へ戻す（H前段）。

正（FLOW §5.3.2）:
- 鍵は JAN
- 税込 = floor(税抜 × (1+税率))
- 見積候補のみ。単品商品名が空／要確認は対象外
- 既に仕入がある行は上書きしない
- 貼付・依頼リスト・検討は触らない

使い方:
  python tools/common_ss_view/estimate_cost_from_paste.py --dry
  python tools/common_ss_view/estimate_cost_from_paste.py
"""
from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "competitor_store"))
from client import sheets_service  # noqa: E402

RES = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
SANK = "17nXSRjFp8xy-3EaBmp-LIu_tMc9ZOgfkhJoZjO3FLIc"
LIST = "仕入検討品番リスト"
PASTE = "貼付sheet"


def num(v):
    try:
        if v is None or v == "":
            return None
        x = float(str(v).replace(",", "").replace("円", "").strip())
        return x if x == x else None
    except Exception:
        return None


def jan_norm(v) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    # Sheets may return scientific; keep digits only when mostly numeric
    if s.replace(".", "", 1).isdigit():
        try:
            s = str(int(float(s)))
        except Exception:
            pass
    if s.isdigit() and len(s) < 13:
        s = s.zfill(13)
    return s


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


def load_paste_prices(svc) -> dict[str, float]:
    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=SANK,
            range=f"'{PASTE}'!A3:F",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    out: dict[str, float] = {}
    if not rows:
        return out
    # row0 = header
    for r in rows[1:]:
        while len(r) < 6:
            r.append("")
        j = jan_norm(r[2])  # C JANコード
        p = num(r[5])  # F 卸値税抜
        if len(j) >= 8 and p and p > 0:
            out[j] = p
    return out


def collect_targets(svc, paste: dict[str, float]):
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
    need = ["メーカー名", "単品商品名", "JANコード", "1点仕入れ額（税込み）", "見積候補", "税率"]
    for n in need:
        if n not in idx:
            raise SystemExit(f"missing col: {n}")

    i_m, i_n, i_j = idx["メーカー名"], idx["単品商品名"], idx["JANコード"]
    i_c, i_cand, i_tax = idx["1点仕入れ額（税込み）"], idx["見積候補"], idx["税率"]
    last = max(i_m, i_n, i_j, i_c, i_cand, i_tax)
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
        "not_cand": 0,
        "bad_name": 0,
        "no_jan": 0,
        "no_paste": 0,
        "has_cost": 0,
        "no_tax": 0,
    }
    for i, r in enumerate(rows):
        while len(r) <= last:
            r.append("")
        if str(r[i_cand]).strip() != "見積候補":
            skip["not_cand"] += 1
            continue
        name = str(r[i_n]).strip()
        if not name or "要確認" in name:
            skip["bad_name"] += 1
            continue
        j = jan_norm(r[i_j])
        if not j:
            skip["no_jan"] += 1
            continue
        p = paste.get(j)
        if not p:
            skip["no_paste"] += 1
            continue
        cost = num(r[i_c])
        if cost and cost > 0:
            skip["has_cost"] += 1
            continue
        tax = num(r[i_tax])
        if tax is None or tax < 0:
            skip["no_tax"] += 1
            continue
        sheet_row = i + 5  # data starts at row 5
        incl = tax_incl_from_ex(p, tax)
        targets.append(
            {
                "row": sheet_row,
                "maker": str(r[i_m]).strip(),
                "name": name,
                "jan": j,
                "ex": p,
                "tax": tax,
                "incl": incl,
                "col": i_c + 1,
            }
        )
    return targets, skip, idx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    args = ap.parse_args()

    svc = sheets_service(write=not args.dry, interactive=False)
    paste = load_paste_prices(svc)
    print(f"paste priced JANs={len(paste)}")
    targets, skip, _idx = collect_targets(svc, paste)
    print("skip", skip)
    print(f"targets={len(targets)}")
    for t in targets:
        print(
            f"  r{t['row']} {t['maker']} | {t['name'][:40]} | JAN={t['jan']} "
            f"| ex={t['ex']} tax={t['tax']} -> incl={t['incl']}"
        )

    if args.dry:
        print("DRY: no write")
        return

    if not targets:
        print("nothing to write")
        return

    # batch by contiguous? use values.batchUpdate with individual cells
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
