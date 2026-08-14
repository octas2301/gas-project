# -*- coding: utf-8 -*-
"""
Amazon MAIN — Pillow 等倍貼付 PoC

  python amazon_paste_poc.py ^
    --master-csv "...\\master_export.csv" ^
    --parent-sku sanky-4538872285127-oya ^
    --set-count 2 --food --stem TUNING_paste_n2
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

from amazon_paste import run_paste_file
from fill_metrics import fill_target_for_n
from work_paths import (
    SUB_TEST_OUT,
    default_work_root,
    resolve_amazon_product_bases,
    resolve_octas,
)

LOG = logging.getLogger("set_main_image.paste_poc")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Amazon MAIN Pillow paste PoC")
    ap.add_argument("--parent-sku", default="")
    ap.add_argument("--set-count", type=int, required=True)
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--out-dir", type=Path, default=None)
    ap.add_argument("--base", type=Path, default=None, help="明示1枚（両方に使用）")
    ap.add_argument("--hero", type=Path, default=None)
    ap.add_argument("--unit", type=Path, default=None)
    ap.add_argument("--octas", type=Path, default=None)
    ap.add_argument("--food", action="store_true")
    ap.add_argument(
        "--unit-only",
        action="store_true",
        help="ヒーロー原画を使わず単体画像だけで全個体を作る（余白テスト用）",
    )
    ap.add_argument(
        "--layout",
        choices=("edge_fill", "legacy_body"),
        default="edge_fill",
        help="edge_fill=縁際余白最小化（既定）",
    )
    ap.add_argument(
        "--aspect",
        choices=("square", "portrait", "landscape"),
        default="square",
        help="square=縁際 / portrait=斜めファン（縦長商品）",
    )
    ap.add_argument("--stem", default="PASTE")
    ap.add_argument("--canvas", type=int, default=1200)
    ap.add_argument("-v", "--verbose", action="store_true")
    # master-csv は将来バッチ用に受け口だけ残す
    ap.add_argument("--master-csv", type=Path, default=None)
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    work_root = args.work_root or default_work_root()
    out_dir = args.out_dir or (work_root / SUB_TEST_OUT)
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.hero and args.unit and not args.unit_only:
        hero, unit = args.hero, args.unit
        mode = "cli_hero_unit"
    else:
        bases = resolve_amazon_product_bases(work_root, args.parent_sku, args.base)
        hero, unit = bases.hero, bases.unit
        mode = bases.mode
        if args.unit_only:
            # 単体タグ優先。無ければ unit 側
            unit = bases.unit
            hero = unit
            mode = "unit_only_test"
            LOG.info("UNIT-ONLY test: all units from %s", unit.name)
        LOG.info("bases mode=%s hero=%s unit=%s", mode, hero.name, unit.name)

    octas = None
    if args.food or args.octas:
        octas = resolve_octas(work_root, args.octas)

    fill_min, band = fill_target_for_n(args.set_count)
    out_path = out_dir / f"{args.stem}_pillow_amazon_set{args.set_count}.jpg"
    meta = run_paste_file(
        hero_path=hero,
        unit_path=unit,
        set_count=args.set_count,
        out_path=out_path,
        octas_path=octas,
        canvas_size=args.canvas,
        fill_min=fill_min,
        layout_mode=args.layout,
        aspect=args.aspect,
    )
    meta["baseMode"] = mode
    meta["fillTarget"] = fill_min
    meta["fillBand"] = band
    from work_paths import meta_dir_for

    json_path = meta_dir_for(out_dir) / f"{out_path.stem}.json"
    json_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(out_path),
                "json": str(json_path),
                "fill": meta.get("fill"),
                "fillPass": meta.get("fillPass"),
                "scale": meta.get("scale"),
                "mode": meta.get("mode"),
                "baseMode": mode,
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
