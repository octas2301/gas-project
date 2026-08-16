# -*- coding: utf-8 -*-
"""P3 dry-run: 貼付◎ → マスタ出品CK。書込なし。"""
from __future__ import annotations

from client import sheets_service
from paste_amazon import cluster_circle_amazon, jan_digits, plan_master_amazon_rows
from schema import MASTER_SS_ID

BLOCK = 10
MASTER_SHEET = "▼商品マスタ(人間作業用)"


def _clusters(svc) -> dict:
    wide = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'ASIN貼り付け（Keepa用）'!A1:CV80")
        .execute()
        .get("values")
        or []
    )
    ai = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'AI情報取得data'!1:12")
        .execute()
        .get("values")
        or []
    )
    ix = {str(h).strip(): i for i, h in enumerate(ai[0])}
    clusters = {}
    for b in range(10):
        if b + 1 >= len(ai) or "JANコード" not in ix:
            continue
        arow = ai[b + 1]
        jan = jan_digits(arow[ix["JANコード"]] if ix["JANコード"] < len(arow) else "")
        if len(jan) < 8:
            continue
        start = b * BLOCK
        last = 1
        for r, row in enumerate(wide):
            if r < 2:
                continue
            if any(str(row[start + c] if start + c < len(row) else "").strip() for c in range(BLOCK)):
                last = r
        rows = []
        for r, row in enumerate(wide):
            if r < 2 or r > last:
                continue
            rows.append(
                {
                    "asin": row[start + 1] if start + 1 < len(row) else "",
                    "title": row[start + 2] if start + 2 < len(row) else "",
                    "eval": row[start + 3] if start + 3 < len(row) else "",
                    "price": row[start + 4] if start + 4 < len(row) else "",
                    "url": row[start + 6] if start + 6 < len(row) else "",
                    "set_count_cell": row[start + 5] if start + 5 < len(row) else "",
                }
            )
        cl = cluster_circle_amazon(rows, jan)
        if jan not in clusters:
            clusters[jan] = cl
            continue
        merged = dict(clusters[jan]["amazonBySet"])
        for k, v in cl["amazonBySet"].items():
            old = merged.get(k)
            if not old or v["priceIncl"] < old["priceIncl"]:
                merged[k] = v
        clusters[jan] = {"jan": jan, "amazonBySet": merged}
    return clusters


def _master_rows(svc) -> list[dict]:
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range="'" + MASTER_SHEET + "'")
        .execute()
        .get("values")
        or []
    )
    hidx = -1
    for i, row in enumerate(vals[:20]):
        cells = [str(x).strip() for x in row]
        if "ASINコード" in cells and "JANコード" in cells:
            hidx = i
            break
    if hidx < 0:
        return []
    h = {str(x).strip(): j for j, x in enumerate(vals[hidx])}
    out = []
    for i, row in enumerate(vals[hidx + 1 :], start=hidx + 2):
        def g(name):
            j = h.get(name)
            if j is None or j >= len(row):
                return ""
            return row[j]

        out.append(
            {
                "row": i,
                "jan": g("JANコード"),
                "set_qty": g("A.セット商品数"),
                "ck": g("出品CK"),
                "current_amazon": g("競合価格amazon"),
            }
        )
    return out


def main() -> None:
    svc = sheets_service()
    clusters = _clusters(svc)
    print("P3 dry_run write=false")
    for jan, cl in clusters.items():
        print(" cluster", jan, sorted((k, v["priceIncl"], v["asin"]) for k, v in cl["amazonBySet"].items()))
    rows = _master_rows(svc)
    ck_n = sum(1 for r in rows if str(r.get("ck")).strip().upper() in ("TRUE", "1") or r.get("ck") is True)
    print("master_rows", len(rows), "ck_true", ck_n)
    paste_jans = set(clusters)
    for r in rows:
        if not (str(r.get("ck")).strip().upper() in ("TRUE", "1") or r.get("ck") is True):
            continue
        print("  ck row", r["row"], "jan", jan_digits(r.get("jan")), "set", r.get("set_qty"), "in_paste", jan_digits(r.get("jan")) in paste_jans)
    planned = plan_master_amazon_rows(rows, clusters)
    print("would_write", len(planned))
    for p in planned[:40]:
        print(
            "  row=%s JAN=%s set=%s now=%s -> %s %s"
            % (p["row"], p["jan"], p["set_qty"], p["current"], p["new_price"], p["new_asin"])
        )
    flagged = [p for p in planned if p["set_qty"] == 5 and p["jan"] == "4906283045119"]
    print("flag_5bag_5119", flagged)


if __name__ == "__main__":
    main()
