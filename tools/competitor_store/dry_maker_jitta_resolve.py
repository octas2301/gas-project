# -*- coding: utf-8 -*-
"""実態リゾルバ仮想テスト（永谷園・味の素・あさくま）。

  python tools/competitor_store/dry_maker_jitta_resolve.py
  python tools/competitor_store/dry_maker_jitta_resolve.py --case T-MJR-02
  python tools/competitor_store/dry_maker_jitta_resolve.py --list
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FIX_DIR = ROOT / "fixtures" / "maker_jitta_resolve"
sys.path.insert(0, str(ROOT))

from maker_alias import JITTAI_BRAND, JITTAI_CHECK, JITTAI_MAKER  # noqa: E402
from maker_jitta_resolve import resolve_maker_jitta, split_amazon_categories  # noqa: E402


def run_case(case: dict) -> tuple[bool, list[str]]:
    errs: list[str] = []
    maker = case["maker"]
    out = resolve_maker_jitta(
        maker,
        summary_row=case.get("summary") or {},
        master_row=case.get("master") or {},
        use_fixture=case.get("use_fixture", True),
    )

    exp = case.get("expect") or {}
    if exp.get("skipped") is not None and bool(out.get("skipped")) != bool(exp["skipped"]):
        errs.append("skipped want=%s got=%s" % (exp["skipped"], out.get("skipped")))

    if "jitta" in exp and str(out.get("jitta")) != str(exp["jitta"]):
        errs.append("jitta want=%s got=%s" % (exp["jitta"], out.get("jitta")))

    if "entry_preferred" in exp and str(out.get("entry_preferred")) != str(exp["entry_preferred"]):
        errs.append(
            "entry_preferred want=%s got=%s" % (exp["entry_preferred"], out.get("entry_preferred"))
        )

    if "auto_confirm" in exp and bool(out.get("auto_confirm")) != bool(exp["auto_confirm"]):
        errs.append("auto_confirm want=%s got=%s" % (exp["auto_confirm"], out.get("auto_confirm")))

    resolved = out.get("resolved") or {}
    master_patch = out.get("master_patch") or {}
    for key in ("gyoshu", "amazon_root_categories", "amazon_browse_shelves", "entry_needle", "confidence"):
        if key in exp:
            if key == "gyoshu":
                got = master_patch.get("業種", resolved.get(key, ""))
            elif key == "amazon_root_categories":
                got = master_patch.get("Amazonルートカテゴリ", resolved.get(key, ""))
            elif key == "amazon_browse_shelves":
                got = master_patch.get("AmazonBrowse棚", resolved.get(key, ""))
            else:
                got = resolved.get(key, "")
            if key in ("amazon_root_categories", "amazon_browse_shelves"):
                got_list = split_amazon_categories(got)
                want_list = split_amazon_categories(exp[key])
                if got_list != want_list:
                    errs.append("%s want=%s got=%s" % (key, want_list, got_list))
            elif str(got) != str(exp[key]):
                errs.append("%s want=%s got=%s" % (key, exp[key], got))

    mirror = out.get("summary_mirror") or {}
    for key in ("業種", "Amazonルートカテゴリ", "AmazonBrowse棚", "実態信頼度"):
        if key in exp:
            got = mirror.get(key, "")
            if key in ("Amazonルートカテゴリ", "AmazonBrowse棚"):
                if split_amazon_categories(got) != split_amazon_categories(exp[key]):
                    errs.append("mirror.%s mismatch got=%s" % (key, got))
            elif str(got) != str(exp[key]):
                errs.append("mirror.%s want=%s got=%s" % (key, exp[key], got))

    if "evidence_contains" in exp:
        url = str((out.get("master_patch") or {}).get("実態根拠URL") or "")
        for frag in exp["evidence_contains"]:
            if frag not in url:
                errs.append("evidence missing %s in %s" % (frag, url))

    return (len(errs) == 0, errs)


def load_cases() -> list[dict]:
    cases: list[dict] = []
    for path in sorted(FIX_DIR.glob("*.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        data["_id"] = path.stem
        cases.append(data)
    return cases


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", help="fixture stem e.g. T-MJR-01")
    ap.add_argument("--list", action="store_true")
    args = ap.parse_args()

    cases = load_cases()
    if args.list:
        for c in cases:
            print(c["_id"], c.get("maker"), (c.get("expect") or {}).get("jitta", ""))
        return 0

    if args.case:
        cases = [c for c in cases if c["_id"] == args.case]
        if not cases:
            print("unknown case", args.case)
            return 1

    ok_n = 0
    for c in cases:
        ok, errs = run_case(c)
        tag = "PASS" if ok else "FAIL"
        print("%s %s %s" % (tag, c["_id"], c.get("maker")))
        for e in errs:
            print("  ", e)
        if ok:
            ok_n += 1

    total = len(cases)
    print("---")
    print("%d/%d PASS" % (ok_n, total))
    return 0 if ok_n == total else 1


if __name__ == "__main__":
    raise SystemExit(main())
