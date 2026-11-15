# -*- coding: utf-8 -*-
"""メーカー／ブランド精度の小テスト案。GETなし。倉庫＋①候補から語とASINを抜く。"""
from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from client import sheets_service  # noqa: E402
from dry_maker_sum import (  # noqa: E402
    as_dicts,
    is_empty_brand,
    maker_dict,
    norm_maker,
    parent_maker,
    read_all,
)
from schema import SHEET_KEEPA_FULL, SHEET_MAKER  # noqa: E402

RESEARCH_SS = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
T_CAND = "①候補"

# 人手で足す入口語（倉庫に無くても小テストしたい）
EXTRA_WORDS = [
    ("ブランド→親", "AGF", "味の素"),
    ("ブランド→親", "エージーエフ", "味の素"),
    ("有名メーカー", "エバラ", "エバラ"),
    ("有名メーカー", "エースコック", "エースコック"),
    ("有名メーカー", "S&B", "エスビー食品"),
    ("有名メーカー", "エスビー食品", "エスビー食品"),
    ("表記ゆれ", "いしもと", "いし本食品工業"),
    ("表記ゆれ", "いし本食品工業", "いし本食品工業"),
]


def main() -> int:
    svc = sheets_service(write=False, interactive=True)
    if not svc:
        print("no_creds")
        return 1
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    _, master = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_MAKER))
    names = maker_dict(cand, full, master)
    cand_by_asin = {}
    for r in cand:
        a = str(r.get("ASIN") or "").strip().upper()
        m = norm_maker(str(r.get("メーカー") or ""))
        if a and m and not is_empty_brand(m):
            cand_by_asin[a] = m

    buckets = defaultdict(list)
    for r in full:
        a = str(r.get("ASIN") or "").strip().upper()
        if not a:
            continue
        brand = str(r.get("ブランド") or "").strip()
        mfr = str(r.get("製造者") or "").strip()
        title = str(r.get("商品名") or "").strip()
        cand_m = cand_by_asin.get(a, "")
        parent = parent_maker(r, cand_by_asin, names)
        row = {
            "asin": a,
            "cand": cand_m,
            "brand": brand or "(空)",
            "mfr": mfr or "(空)",
            "parent": parent,
            "title": title[:60],
        }
        if is_empty_brand(brand):
            buckets["ノーブランド"].append(row)
        elif cand_m and norm_maker(brand) != cand_m and norm_maker(brand) not in (cand_m,):
            buckets["ブランド≠候補メーカー"].append(row)
        elif cand_m and cand_m == parent and not is_empty_brand(brand):
            buckets["候補と一致"].append(row)
        if parent == "未確定":
            buckets["親未確定"].append(row)
        # 店名っぽい製造者
        junk = ("通販", "直行", "セレクト", "007", "---", "パーク")
        if any(j in mfr for j in junk):
            buckets["製造者ゴミ"].append(row)

    print("=== 語の案（Catalogキーワード小テスト・GETは別承認）===")
    for kind, word, parent in EXTRA_WORDS:
        print("%s\t入口語=%s\t期待親=%s" % (kind, word, parent))
    # 倉庫に多いブランド上位（入口語候補）
    bc = defaultdict(int)
    for r in full:
        b = str(r.get("ブランド") or "").strip()
        if b and not is_empty_brand(b):
            bc[b] += 1
    print("=== 倉庫ブランド上位（入口語の実データ）===")
    for b, n in sorted(bc.items(), key=lambda x: -x[1])[:12]:
        print("n=%d\t%s" % (n, b))

    print("=== ASIN案（各バケツ最大3・GETなし）===")
    order = ["候補と一致", "ブランド≠候補メーカー", "ノーブランド", "親未確定", "製造者ゴミ"]
    for k in order:
        rows = buckets.get(k) or []
        print("[%s] n=%d" % (k, len(rows)))
        for row in rows[:3]:
            print(
                "  %s cand=%s brand=%s mfr=%s parent=%s | %s"
                % (row["asin"], row["cand"], row["brand"], row["mfr"], row["parent"], row["title"])
            )
    print("dry_probe_ok (no write, no GET)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
