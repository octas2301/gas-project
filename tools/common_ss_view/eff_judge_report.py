# -*- coding: utf-8 -*-
"""①ログの「効率合否」行を集計し、合否・原因・対策ヒントを出す。

  python tools/common_ss_view/dump.py dump --menu 仕入れ検討① --intent keepa --no-limit
  python tools/common_ss_view/eff_judge_report.py
  python tools/common_ss_view/eff_judge_report.py --hours 24
"""
from __future__ import annotations

import argparse
import csv
import re
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DEFAULT_LOG = (
    ROOT
    / "tools"
    / "common_ss_view"
    / "out"
    / "自動リサーチ260814_アマゾン調査"
    / "①ログ.tsv"
)


def parse_at(s: str) -> datetime | None:
    s = (s or "").strip()
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M"):
        try:
            return datetime.strptime(s[:19], fmt)
        except ValueError:
            pass
    return None


def parse_kv(detail: str) -> dict[str, str]:
    out: dict[str, str] = {}
    for part in str(detail or "").split():
        if "=" not in part:
            continue
        k, v = part.split("=", 1)
        out[k] = v
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="効率合否ログ集計")
    ap.add_argument("--log", type=Path, default=DEFAULT_LOG)
    ap.add_argument("--hours", type=float, default=36.0)
    args = ap.parse_args()
    if not args.log.is_file():
        print("MISSING", args.log)
        print("先に: python tools/common_ss_view/dump.py dump --menu 仕入れ検討① --intent keepa --no-limit")
        return 2

    rows = list(csv.DictReader(args.log.open(encoding="utf-8-sig"), delimiter="\t"))
    # detail is often last column; also named
    ats = [t for t in (parse_at(r.get("at") or "") for r in rows) if t]
    latest = max(ats, default=None)
    if not latest:
        print("no timestamps")
        return 1
    start = latest - timedelta(hours=args.hours)

    judges = []
    for r in rows:
        if (r.get("内容") or "") != "効率合否":
            continue
        at = parse_at(r.get("at") or "")
        if not at or at < start:
            continue
        # detail column
        detail = ""
        for k, v in r.items():
            if k in ("runId", "内容", "state", "at"):
                continue
            if v and len(str(v)) > len(detail):
                detail = str(v)
        kv = parse_kv(detail)
        judges.append(
            {
                "at": at,
                "verdict": (r.get("state") or kv.get("verdict") or "").upper(),
                "event": kv.get("event", ""),
                "cause": kv.get("cause", ""),
                "fixHint": kv.get("fixHint", ""),
                "parent": kv.get("parent", ""),
                "kv": kv,
            }
        )

    print("window", start, "→", latest, "n=", len(judges))
    if not judges:
        print("効率合否行なし（clasp push 後の実行が必要）")
        return 0

    by_v = Counter(j["verdict"] for j in judges)
    by_e = Counter(j["event"] for j in judges)
    by_c = Counter(j["cause"] for j in judges)
    print("verdict", dict(by_v))
    print("event", dict(by_e))
    print("cause top", by_c.most_common(12))

    print("\n=== FAIL / WARN（新しい順・最大40）===")
    shown = 0
    for j in reversed(judges):
        if j["verdict"] not in ("FAIL", "WARN"):
            continue
        print(
            j["at"].strftime("%m/%d %H:%M"),
            j["verdict"],
            j["event"],
            j["cause"],
            j["parent"],
            "fix=" + j["fixHint"] if j["fixHint"] else "",
        )
        shown += 1
        if shown >= 40:
            break

    # cause → fixHint majority
    print("\n=== 原因→対策ヒント（FAIL/WARN）===")
    pairs: dict[str, Counter] = defaultdict(Counter)
    for j in judges:
        if j["verdict"] in ("FAIL", "WARN") and j["cause"]:
            pairs[j["cause"]][j["fixHint"] or "(none)"] += 1
    for cause, hints in sorted(pairs.items(), key=lambda x: -sum(x[1].values())):
        tip, n = hints.most_common(1)[0]
        print("%s x%s → %s" % (cause, sum(hints.values()), tip))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
