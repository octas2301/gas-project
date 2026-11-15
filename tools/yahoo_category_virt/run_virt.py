# -*- coding: utf-8 -*-
"""Yahooカテゴリ仮想テスト実行。

  python tools/yahoo_category_virt/run_virt.py
  python tools/yahoo_category_virt/run_virt.py --fixture row20_coffee_noise.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from score import legacy_weight_only, pick_yahoo_category  # noqa: E402


def load_fixtures(one: str | None) -> list[Path]:
    d = ROOT / "fixtures"
    if one:
        p = d / one
        if not p.exists():
            raise SystemExit("fixture not found: %s" % p)
        return [p]
    return sorted(d.glob("*.json"))


def run_one(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    name = data.get("productName") or ""
    maker = data.get("makerName") or ""
    hits = data.get("hits") or []
    expected = str(data.get("expectedCategoryId") or "").strip()
    fallback = data.get("fallback") or {}
    top_n = int(data.get("topN") or 15)

    new = pick_yahoo_category(name, maker, hits, top_n=top_n, fallback=fallback)
    old = legacy_weight_only(hits, top_n=top_n)
    ok = (not expected) or (new.get("categoryId") == expected)
    return {
        "file": path.name,
        "ok": ok,
        "expected": expected,
        "newId": new.get("categoryId"),
        "newReason": new.get("reason"),
        "oldId": old.get("categoryId"),
        "sameCount": new.get("sameCount"),
        "seriesCount": new.get("seriesCount"),
        "seriesLooseCount": new.get("seriesLooseCount"),
        "vote": new.get("vote"),
        "log": new.get("log"),
        "needsReview": new.get("needsReview"),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--fixture", default="", help="単一 JSON 名")
    args = ap.parse_args()
    files = load_fixtures(args.fixture or None)
    if not files:
        print("NO fixtures")
        return 2
    failed = 0
    for f in files:
        r = run_one(f)
        mark = "PASS" if r["ok"] else "FAIL"
        if not r["ok"]:
            failed += 1
        print(
            "[%s] %s expected=%s new=%s (old_legacy=%s) same=%s series=%s loose=%s reason=%s"
            % (
                mark,
                r["file"],
                r["expected"],
                r["newId"],
                r["oldId"],
                r["sameCount"],
                r.get("seriesCount"),
                r.get("seriesLooseCount"),
                r["newReason"],
            )
        )
        for line in r.get("log") or []:
            print("  ", line)
        if r.get("vote"):
            print("   vote=", r["vote"])
    print("---")
    print("total=%d fail=%d" % (len(files), failed))
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
