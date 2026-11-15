# -*- coding: utf-8 -*-
"""D2 classify 単体仮想テスト。

  python tools/competitor_store/dry_d2_classify_unit.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "d2_classify"
sys.path.insert(0, str(ROOT))

from d2_filter import classify  # noqa: E402
from maker_alias import JITTAI_MAKER  # noqa: E402

PASS_CAND = {"門結果": "通過", "門理由": "門通過"}


def run_case(path: Path) -> tuple[bool, str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    case_id = data.get("id") or path.stem
    inp = data.get("input") or {}
    expect = data.get("expect") or {}
    flag, reason = classify(
        inp.get("asin") or "B000000000",
        inp.get("maker") or "テスト",
        inp.get("title") or "通常商品",
        inp.get("category") or "食品・飲料・お酒",
        inp.get("cand") if "cand" in inp else PASS_CAND,
        inp.get("jitta") or JITTAI_MAKER,
        qs_tier=str(inp.get("qs_tier") or ""),
        weight=str(inp.get("weight") or ""),
        prof=inp.get("prof"),
    )
    ok = flag == expect.get("flag") and reason == expect.get("reason", "")
    if ok:
        return True, case_id
    return False, "%s flag=%r reason=%r expect=%r" % (case_id, flag, reason, expect)


def main() -> int:
    cases = sorted(FIX_DIR.glob("*.json"))
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
