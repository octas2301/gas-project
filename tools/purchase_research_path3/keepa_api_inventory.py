#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""1 ASIN の product API からキー一覧を作り、DL85列と突き合わせる。履歴の中身は保存しない。"""

from __future__ import annotations

import csv
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Set

from keepa_csv_vs_api import ASINS, fetch_products, keepa_key
from keepa_field_coverage import MAP

CSV_IDX = {
    0: "AMAZON",
    1: "NEW",
    2: "USED",
    3: "SALES",
    4: "LISTPRICE",
    5: "COLLECTIBLE",
    6: "REFURBISHED",
    7: "NEW_FBM_SHIPPING",
    8: "LIGHTNING_DEAL",
    9: "WAREHOUSE",
    10: "NEW_FBA",
    11: "COUNT_NEW",
    12: "COUNT_USED",
    13: "COUNT_REFURBISHED",
    14: "COUNT_COLLECTIBLE",
    15: "EXTRA_INFO_UPDATES",
    16: "RATING",
    17: "COUNT_REVIEWS",
    18: "BUY_BOX_SHIPPING",
    19: "USED_NEW_SHIPPING",
    20: "USED_VERY_GOOD_SHIPPING",
    21: "USED_GOOD_SHIPPING",
    22: "USED_ACCEPTABLE_SHIPPING",
    23: "COLLECTIBLE_NEW_SHIPPING",
    24: "COLLECTIBLE_VERY_GOOD_SHIPPING",
    25: "COLLECTIBLE_GOOD_SHIPPING",
    26: "COLLECTIBLE_ACCEPTABLE_SHIPPING",
    27: "REFURBISHED_SHIPPING",
    28: "EBAY_NEW_SHIPPING",
    29: "EBAY_USED_SHIPPING",
    30: "TRADE_IN",
    31: "RENTAL",
    32: "PRIME_EXCL",
    33: "NEW_FBA",
}


def walk(obj: Any, prefix: str, out: List[str], depth: int = 0) -> None:
    if depth > 6:
        return
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = "%s.%s" % (prefix, k) if prefix else str(k)
            if k == "csv" and isinstance(v, list):
                for i, series in enumerate(v):
                    lab = CSV_IDX.get(i, str(i))
                    present = "null" if series is None else ("n=%s" % (len(series) if isinstance(series, list) else "1"))
                    out.append("csv[%s](%s) %s" % (i, lab, present))
                continue
            if k == "offers" and isinstance(v, list):
                out.append("offers n=%s" % len(v))
                if v and isinstance(v[0], dict):
                    for ok in sorted(v[0].keys()):
                        out.append("offers[].%s" % ok)
                continue
            if isinstance(v, (dict, list)) and not _is_leaf_list(v):
                walk(v, p, out, depth + 1)
            else:
                out.append("%s = %s" % (p, _leaf(v)))
    elif isinstance(obj, list):
        if obj and isinstance(obj[0], dict):
            out.append("%s[] n=%s keys=%s" % (prefix, len(obj), ",".join(sorted(obj[0].keys())[:40])))
        else:
            out.append("%s list_len=%s" % (prefix, len(obj)))


def _is_leaf_list(v: Any) -> bool:
    if not isinstance(v, list):
        return False
    if not v:
        return True
    return not isinstance(v[0], (dict, list))


def _leaf(v: Any) -> str:
    if v is None:
        return "null"
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return "num"
    if isinstance(v, str):
        return "str:%s" % (len(v),)
    if isinstance(v, list):
        return "list:%s" % (len(v),)
    return type(v).__name__


def main() -> int:
    key = keepa_key()
    if not key:
        print("no key")
        return 2
    asin = ASINS[1]  # 石原水産の方
    data = fetch_products(key, [asin])
    products = data.get("products") or []
    print("tokensLeft=%s n=%s" % (data.get("tokensLeft"), len(products)))
    if not products:
        print("no product", list(data.keys()))
        return 1
    p = products[0]
    lines: List[str] = []
    walk(p, "product", lines)

    base = Path(__file__).resolve().parent
    inv = base / "石原水産_KeepaAPI全項目.csv"
    with inv.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["asin", "path"])
        for line in lines:
            w.writerow([asin, line])

    blob = "\n".join(lines)
    dest = base / "石原水産_KeepaDL85_vs_API全項目.csv"
    with dest.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["export_column", "api_class", "api_path", "seen_in_this_asin", "research_ok", "note"])
        for col, cls, path, note in MAP:
            seen = "Y" if _seen(path, blob, p) else "N"
            ok = "Y" if cls in ("yes", "derive", "build", "offers") else "review"
            if cls == "hard":
                ok = "optional"
            w.writerow([col, cls, path, seen, ok, note])
    print("inventory", len(lines), "wrote", inv.name, dest.name)
    return 0


def _seen(path: str, blob: str, p: Dict[str, Any]) -> bool:
    needles = []
    if "title" in path:
        needles.append("product.title")
    if "brand" in path:
        needles.append("product.brand")
    if "manufacturer" in path:
        needles.append("product.manufacturer")
    if "monthlySold" in path:
        needles.append("product.monthlySold")
    if "imagesCSV" in path or path.startswith("images"):
        needles.append("product.imagesCSV")
    if "features" in path:
        needles.append("product.features")
    if "categoryTree" in path:
        needles.append("product.categoryTree")
    if "packageLength" in path:
        needles.append("product.packageLength")
    if "ean" in path.lower():
        needles += ["product.ean", "product.eanList"]
    if "offers" in path:
        needles.append("offers")
    if "stats" in path or "avg90" in path or "avg365" in path:
        needles.append("product.stats")
    if "csv[" in path or "csv[3]" in path or "履歴" in path:
        needles.append("csv[")
    if "asin" in path.lower() and "amazon.co.jp" not in path:
        needles.append("product.asin")
    if "numberOfItems" in path:
        needles.append("product.numberOfItems")
    if "releaseDate" in path:
        needles.append("product.releaseDate")
    if "amazon.co.jp" in path or "keepa.com" in path:
        return True
    if "availabilityAmazon" in path:
        needles.append("product.availabilityAmazon")
    if "coupon" in path:
        needles.append("product.coupon")
    if "sns" in path.lower() or "Subscribe" in path:
        needles += ["product.isSNS", "csv["]
    if needles:
        return any(n in blob or n.replace("product.", "") in p for n in needles)
    if path in p:
        return True
    return "product.stats" in blob


if __name__ == "__main__":
    sys.exit(main())
