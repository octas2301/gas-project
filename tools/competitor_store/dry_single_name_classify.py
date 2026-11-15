# -*- coding: utf-8 -*-
"""単品名・JAN パターン仮想テスト（20件）。

  python tools/competitor_store/dry_single_name_classify.py
  python tools/competitor_store/dry_single_name_classify.py --case T-SN-01
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "single_name"
sys.path.insert(0, str(ROOT))

from single_name_classify import classify_single_name  # noqa: E402

KEYS = (
    "pattern",
    "single_trust",
    "jan_trust",
    "for_estimate",
    "need_human",
    "write_single_name",
    "corpus_assist_ok",
    "auto_pick_jan",
)


def run_case(path: Path) -> tuple[bool, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    case_id = data.get("id") or path.stem
    expect = data.get("expect") or {}
    result = classify_single_name(data.get("input") or {})
    diffs: list[str] = []
    for k in KEYS:
        if k not in expect:
            continue
        if result.get(k) != expect[k]:
            diffs.append("%s: expect=%r actual=%r" % (k, expect[k], result.get(k)))
    if "jan_candidates_len" in expect:
        n = len(result.get("jan_candidates") or [])
        if n != int(expect["jan_candidates_len"]):
            diffs.append("jan_candidates_len: expect=%r actual=%r" % (expect["jan_candidates_len"], n))
    if diffs:
        return False, "%s %s | %s" % (case_id, data.get("desc") or "", "; ".join(diffs))
    return True, "%s %s" % (case_id, data.get("desc") or "")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", default="", help="T-SN-01 など（部分一致）")
    args = ap.parse_args()
    cases = sorted(FIX_DIR.glob("T-SN-*.json"))
    if args.case:
        cases = [p for p in cases if args.case in p.stem]
    if not cases:
        print("FAIL no fixtures in", FIX_DIR)
        return 1
    fail = 0
    for p in cases:
        ok, msg = run_case(p)
        print(("PASS" if ok else "FAIL"), msg)
        if not ok:
            fail += 1
    print("SUMMARY", len(cases) - fail, "/", len(cases))
    return 1 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
