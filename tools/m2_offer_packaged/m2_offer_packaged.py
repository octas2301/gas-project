#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
M2（TRACK=A）案L: GENERATED offer 行 → Listing Loader 系 CSV

C1 HPC xlsm は使わない。公式テンプレ列は column_map.json で調整。
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

LOG = logging.getLogger("m2_offer")
SCRIPT_DIR = Path(__file__).resolve().parent


def _utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")


def resolve_path(p: str, base: Path) -> Path:
    path = Path(p)
    if not path.is_absolute():
        path = (base / path).resolve()
    return path


def expand_sub_batch_id_(path_str: str, sub_batch_id: str) -> str:
    s = str(path_str or "")
    if "{subBatchId}" not in s:
        return s
    sid = str(sub_batch_id or "").strip()
    if not sid:
        raise ValueError(
            "generated_csv に {subBatchId} があります。"
            "config.sub_batch_id または --sub-batch を指定してください。"
        )
    return s.replace("{subBatchId}", sid)


def _load_json(path: Path) -> Dict[str, Any]:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def _save_json(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
        f.write("\n")


def sub_batch_id_from_generated(path: Path) -> str:
    m = re.match(r"^(.+)_GENERATED\.csv$", path.name, re.I)
    return m.group(1) if m else path.stem


def load_generated_offer_rows(path: Path) -> List[Dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("GENERATED が空です: %s" % path)
        rows: List[Dict[str, str]] = []
        for raw in reader:
            row = {k: str(raw.get(k) or "").strip().replace("\r", "") for k in reader.fieldnames}
            track = (row.get("track") or "").upper()
            role = (row.get("variationRole") or "").lower()
            if track == "A" and role == "offer":
                rows.append(row)
            elif track == "A":
                LOG.warning("track=A だが variationRole!=offer をスキップ: %s", row.get("sellerSku"))
        return rows


def filter_rows(
    rows: List[Dict[str, str]],
    parent_filter: List[str],
    sku_filter: List[str],
) -> List[Dict[str, str]]:
    out = rows
    if parent_filter:
        ps = set(parent_filter)
        out = [r for r in out if (r.get("parentSku") or "") in ps]
    if sku_filter:
        ss = set(sku_filter)
        out = [r for r in out if (r.get("sellerSku") or "") in ss]
    return out


def build_output_row(
    src: Dict[str, str],
    headers: List[str],
    mapping: Dict[str, str],
    defaults: Dict[str, str],
) -> Dict[str, str]:
    row: Dict[str, str] = {}
    for h in headers:
        if h in mapping:
            row[h] = src.get(mapping[h]) or defaults.get(h) or ""
        else:
            row[h] = defaults.get(h) or ""
    return row


def validate_row(row: Dict[str, str], line_no: int) -> Optional[str]:
    sku = row.get("sku") or ""
    asin = row.get("product-id") or ""
    price = row.get("price") or ""
    qty = row.get("quantity") or ""
    if not sku:
        return "line %d: sku 空" % line_no
    if not re.match(r"^B0[A-Z0-9]{8}$", asin, re.I):
        return "line %d: ASIN不正 %r sku=%s" % (line_no, asin, sku)
    try:
        if float(price) <= 0:
            return "line %d: price 不正 %r" % (line_no, price)
    except ValueError:
        return "line %d: price 非数値 %r" % (line_no, price)
    if qty not in ("0", "1"):
        LOG.warning("line %d: quantity=%r （0/1推奨）sku=%s", line_no, qty, sku)
    return None


def run(config_path: Path, mode: str, sub_batch_id: Optional[str] = None) -> int:
    cfg = _load_json(config_path)
    base = config_path.parent
    mode = (mode or cfg.get("mode") or "dry_run").strip().lower()
    if mode not in ("dry_run", "prod"):
        raise SystemExit("mode は dry_run または prod")

    sid = str(sub_batch_id or cfg.get("sub_batch_id") or "").strip()
    gen_path = resolve_path(expand_sub_batch_id_(cfg["generated_csv"], sid), base)
    if not gen_path.is_file():
        raise FileNotFoundError("GENERATED がありません: %s" % gen_path)

    map_path = resolve_path(cfg.get("column_map_path") or "column_map.json", base)
    colmap = _load_json(map_path)
    headers: List[str] = list(colmap.get("output_headers") or [])
    defaults = dict(colmap.get("defaults") or {})
    mapping = dict(colmap.get("generated_to_output") or {})

    output_dir = resolve_path(cfg.get("output_dir") or "out", base)
    log_dir = resolve_path(cfg.get("log_dir") or str(output_dir), base)
    output_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    sub = sid or sub_batch_id_from_generated(gen_path)
    run_id = "M2_%s_%s" % (sub, _utc_stamp())

    offers = load_generated_offer_rows(gen_path)
    parent_filter = list(cfg.get("parent_sku_filter") or [])
    sku_filter = list(cfg.get("seller_sku_filter") or [])
    offers = filter_rows(offers, parent_filter, sku_filter)

    out_rows: List[Dict[str, str]] = []
    errors: List[str] = []
    for i, src in enumerate(offers, start=2):
        row = build_output_row(src, headers, mapping, defaults)
        err = validate_row(row, i)
        if err:
            errors.append(err)
            continue
        out_rows.append(row)

    report = {
        "runId": run_id,
        "mode": mode,
        "version": "M2-L-v1",
        "subBatchId": sub,
        "generatedCsv": str(gen_path),
        "counts": {
            "offerRowsIn": len(offers),
            "accepted": len(out_rows),
            "rejected": len(errors),
        },
        "errors": errors,
        "sample": out_rows[:3],
    }

    suffix = "_DRYRUN" if mode == "dry_run" else ""
    out_name = "%s_M2_OFFER_LOADER%s.csv" % (sub, suffix)
    out_path = output_dir / out_name
    if out_path.exists():
        out_path = output_dir / ("%s_M2_OFFER_LOADER_%s%s.csv" % (sub, _utc_stamp(), suffix))

    report_path = log_dir / ("%s_M2_REPORT.json" % run_id)
    report["outputPath"] = str(out_path)

    if mode == "prod" and not out_rows:
        _save_json(report_path, report)
        LOG.error("受理行0のため prod 出力なし。report=%s", report_path)
        return 2

    if out_rows or mode == "dry_run":
        with out_path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
            w.writeheader()
            for row in out_rows:
                w.writerow(row)
        LOG.info("出力: %s rows=%d", out_path, len(out_rows))
    else:
        report["outputPath"] = ""

    _save_json(report_path, report)
    LOG.info("レポート: %s accepted=%d rejected=%d", report_path, len(out_rows), len(errors))
    for e in errors[:10]:
        LOG.warning("%s", e)
    return 0 if not errors else (0 if out_rows else 1)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="M2 TRACK=A offer → Listing Loader CSV (案L)")
    parser.add_argument("--config", required=True)
    parser.add_argument("--mode", choices=["dry_run", "prod"], default=None)
    parser.add_argument("--sub-batch", default=None)
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    config_path = Path(args.config).resolve()
    cfg = _load_json(config_path)
    return run(config_path, args.mode or cfg.get("mode") or "dry_run", args.sub_batch)


if __name__ == "__main__":
    sys.exit(main())
