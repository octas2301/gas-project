#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""経路3 PoC: Keepa集合 A と Amazon集合 B の差（B-A = Keepa漏れ）。認証不要。"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

ASIN_RE = re.compile(r"\b([A-Z0-9]{10})\b", re.I)


def load_asins(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    seen: set[str] = set()
    out: list[str] = []
    for line in text.splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        m = ASIN_RE.search(s.upper())
        if not m:
            continue
        a = m.group(1).upper()
        if a in seen:
            continue
        seen.add(a)
        out.append(a)
    return out


def main() -> int:
    p = argparse.ArgumentParser(description="Keepa A vs Amazon B (path3 rescue = B-A)")
    p.add_argument("--keepa", required=True, help="経路2 ASIN ファイル")
    p.add_argument("--amazon", required=True, help="経路3 ASIN ファイル")
    p.add_argument("--out-dir", default="", help="省略時 tools/.../out")
    args = p.parse_args()

    keepa_path = Path(args.keepa)
    amazon_path = Path(args.amazon)
    a = set(load_asins(keepa_path))
    b = set(load_asins(amazon_path))
    rescue = sorted(b - a)
    only_keepa = sorted(a - b)
    both = sorted(a & b)

    report = {
        "at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "keepa_file": str(keepa_path),
        "amazon_file": str(amazon_path),
        "count_keepa_A": len(a),
        "count_amazon_B": len(b),
        "count_both": len(both),
        "count_rescue_B_minus_A": len(rescue),
        "count_only_keepa_A_minus_B": len(only_keepa),
        "rescue_asins": rescue,
        "only_keepa_asins": only_keepa,
        "both_asins": both,
        "note": "品質CPは rescue（Keepa漏れ）。スクレイプ結果は入れない。",
    }

    out_dir = Path(args.out_dir) if args.out_dir else Path(__file__).resolve().parent / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    dest = out_dir / ("PATH3_DIFF_%s.json" % stamp)
    dest.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("A(Keepa)=%s B(Amazon)=%s both=%s rescue(B-A)=%s onlyKeepa=%s" % (
        len(a), len(b), len(both), len(rescue), len(only_keepa)
    ))
    print("report=%s" % dest)
    if rescue:
        print("rescue: %s" % ", ".join(rescue[:30]))
        if len(rescue) > 30:
            print("... +%s" % (len(rescue) - 30))
    return 0


if __name__ == "__main__":
    sys.exit(main())
