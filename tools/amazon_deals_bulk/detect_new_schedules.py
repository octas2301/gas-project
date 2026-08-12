# -*- coding: utf-8 -*-
"""
推奨バルク取込時の「新しい名付きSale」差分検知。

前回スナップショット（_work/schedule_catalog_snapshot.json）と
②最新xlsxの ValidationDataSheet／埋込schedules を比較し、新規スケジュール名を出す。

例:
  python detect_new_schedules.py
  python detect_new_schedules.py --save
  python detect_new_schedules.py --source path/to.xlsx --save
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Set

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from openpyxl import load_workbook  # noqa: E402

from paths import folder_path, latest_xlsx, load_config  # noqa: E402
from schedule_class import is_official_b_schedule  # noqa: E402
from template_parse import collect_schedule_catalog  # noqa: E402

LOG = logging.getLogger("amazon_deals_bulk.detect_new_schedules")
SNAP = HERE / "_work" / "schedule_catalog_snapshot.json"


def load_snapshot(path: Path) -> Dict[str, Any]:
    if not path.is_file():
        return {"schedules": [], "saved_at": None}
    return json.loads(path.read_text(encoding="utf-8"))


def save_snapshot(path: Path, catalog: List[Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "saved_at": datetime.now(timezone.utc).isoformat(),
        "schedules": catalog,
        "names": sorted({str(c.get("schedule") or "") for c in catalog if c.get("schedule")}),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def named_only(catalog: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = []
    for c in catalog:
        name = str(c.get("schedule") or "").strip()
        if not name:
            continue
        if is_official_b_schedule(name):
            out.append(c)
    return out


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="新Sale（名付き）差分検知")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--source", type=Path, default=None)
    ap.add_argument("--save", action="store_true", help="今回カタログをスナップショット保存")
    ap.add_argument("--snapshot", type=Path, default=SNAP)
    args = ap.parse_args(argv)

    local = HERE / "config.local.json"
    cfg = load_config(args.config or (local if local.is_file() else HERE / "config.example.json"))
    src = args.source
    if src is None:
        src = latest_xlsx(folder_path(cfg, "02")) or latest_xlsx(folder_path(cfg, "01"))
    if src is None or not Path(src).is_file():
        LOG.error("推奨xlsxが見つかりません（①/②フォルダ）")
        return 1

    wb = load_workbook(str(src), read_only=True, data_only=True)
    try:
        catalog = named_only(collect_schedule_catalog(wb))
    finally:
        wb.close()

    names_now: Set[str] = {str(c.get("schedule") or "") for c in catalog}
    prev = load_snapshot(args.snapshot)
    names_prev: Set[str] = set(prev.get("names") or [])
    if not names_prev and prev.get("schedules"):
        names_prev = {str(c.get("schedule") or "") for c in prev["schedules"]}

    added = sorted(names_now - names_prev)
    removed = sorted(names_prev - names_now)

    print("=== カタログ差分（名付き公式のみ） ===")
    print("source:", src)
    print("prev_saved_at:", prev.get("saved_at"))
    print("now_count:", len(names_now), "prev_count:", len(names_prev))
    if not names_prev:
        print("（初回: 差分なし扱い。--save で基準を保存）")
    print("--- 新規Sale ---")
    if added:
        for n in added:
            hit = next((c for c in catalog if c.get("schedule") == n), {})
            print("  +", n, hit.get("start"), hit.get("end"))
    else:
        print("  (なし)")
    print("--- 消えたSale（参考） ---")
    if removed:
        for n in removed[:20]:
            print("  -", n)
        if len(removed) > 20:
            print("  ...他", len(removed) - 20)
    else:
        print("  (なし)")

    if args.save:
        save_snapshot(args.snapshot, catalog)
        LOG.info("snapshot saved → %s", args.snapshot)

    out = {
        "added": added,
        "removed_count": len(removed),
        "saved": bool(args.save),
    }
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_path = HERE / "_work" / ("new_schedules_%s.json" % stamp)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"added": len(added), "out": str(out_path)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
