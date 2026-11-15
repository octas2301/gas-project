#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""既存 hit_rate_compare の楽天列だけ Referer 修正後に再計測。"""
from __future__ import annotations

import csv
import importlib.util
from collections import defaultdict
from pathlib import Path

SCRIPT = Path(__file__).resolve().parent / "hit_rate_compare.py"
spec = importlib.util.spec_from_file_location("hrc", SCRIPT)
hrc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(hrc)

DETAIL = Path(__file__).resolve().parent / "out" / "hit_rate_compare_20260820_074451.csv"
OUT = Path(__file__).resolve().parent / "out" / "hit_rate_rakuten_rerun.csv"


def main() -> None:
    cfg = hrc.load_rakuten_cfg()
    products: dict[str, hrc.Product] = {}
    with DETAIL.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row["mall"] != "rakuten" or row["strategy"] != "jan_first":
                continue
            pid = row["product_id"]
            if pid in products:
                continue
            products[pid] = hrc.Product(
                pid=pid,
                jan=row["jan"],
                maker=row["maker"],
                name=row["name"],
                is_combo=bool(int(row["is_combo"])),
                kind_n=int(row["kind_n"] or 0),
                source=row["source"],
            )

    rows = []
    stats = {
        "jan_first": {"raw": 0, "rel": 0, "n": 0},
        "fallback": {"raw": 0, "rel": 0, "n": 0},
    }
    for p in products.values():
        for strat, specs in (
            ("jan_first", hrc.build_jan_first_queries(p)),
            ("fallback", hrc.build_fallback_queries(p)),
        ):
            res = hrc.eval_strategy("rakuten", p, specs, cfg)
            stats[strat]["n"] += 1
            stats[strat]["raw"] += int(res["raw"])
            stats[strat]["rel"] += int(res["relevant"])
            rows.append(
                {
                    "product_id": p.pid,
                    "jan": p.jan,
                    "is_combo": int(p.is_combo),
                    "strategy": strat,
                    "raw_hit": int(res["raw"]),
                    "relevant_hit": int(res["relevant"]),
                    "query_used": res["query_used"],
                    "error": res["error"],
                }
            )

    with OUT.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    n = stats["jan_first"]["n"] or 1
    print("rakuten rerun products=%d" % n)
    for strat in ("jan_first", "fallback"):
        s = stats[strat]
        print(
            "%s raw=%d/%d (%.1f%%) relevant=%d/%d (%.1f%%)"
            % (strat, s["raw"], n, 100 * s["raw"] / n, s["rel"], n, 100 * s["rel"] / n)
        )
    print("wrote", OUT)


if __name__ == "__main__":
    main()
