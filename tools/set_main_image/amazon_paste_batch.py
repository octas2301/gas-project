# -*- coding: utf-8 -*-
"""
Amazon MAIN — Pillow 等倍貼付 一括（出品CKレ点）

  # 推奨: スプシ直読（CSVダウンロード不要）
  python amazon_paste_batch.py --from-sheets --checked-only

  # 互換: ローカルCSV
  python amazon_paste_batch.py --master-csv "...\\master.csv" --checked-only

現行ロジック: 単体のみ / Octas=不透明面積5% / 右下キス
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from amazon_paste import run_paste_file
from fill_metrics import fill_target_for_n
from master_sets import (
    load_checked_set_children,
    load_checked_set_children_from_rows,
    load_set_children_for_parent,
)
from work_paths import (
    ASPECT_CHOICES,
    ASPECT_LABEL_JA,
    ASPECT_LANDSCAPE,
    ASPECT_PORTRAIT,
    ASPECT_SQUARE,
    DEFAULT_AMAZON_07_OUT,
    SUB_AMAZON,
    SUB_META,
    SUB_TEST_OUT,
    TAG_UNIT,
    default_work_root,
    resolve_amazon_product_bases,
    resolve_octas,
)

LOG = logging.getLogger("set_main_image.paste_batch")


def _safe_stem(s: str) -> str:
    out = []
    for ch in s:
        if ch.isalnum() or ch in ("-", "_"):
            out.append(ch)
        else:
            out.append("_")
    return "".join(out)[:80] or "sku"


def _parent_match_tokens(parent_sku: str) -> List[str]:
    import re

    p = (parent_sku or "").strip().lower()
    toks = [p] if p else []
    for m in re.findall(r"\d{8,}", p):
        toks.append(m)
    return toks


def _unit_files(work_root: Path) -> List[Path]:
    """01直下の素材。『単体』タグ優先。無ければ画像全件（処理済み以外）。"""
    folder = work_root / SUB_AMAZON
    if not folder.is_dir():
        return []
    tagged: List[Path] = []
    all_imgs: List[Path] = []
    for p in folder.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in (".jpg", ".jpeg", ".png", ".webp"):
            continue
        all_imgs.append(p)
        if TAG_UNIT in p.stem:
            tagged.append(p)
    return tagged if tagged else all_imgs


def _unit_matches_parent(unit_path: Path, parent_sku: str) -> bool:
    """単体ファイル名に親SKU／JAN断片が含まれるか（誤商品貼付防止）。"""
    name = unit_path.name.lower()
    stem = unit_path.stem.lower()
    for t in _parent_match_tokens(parent_sku):
        if t and (t in name or t in stem):
            return True
    return False


def _allow_unit_for_parent(
    unit_path: Path,
    parent_sku: str,
    work_root: Path,
    *,
    loose_unit_match: bool,
) -> Tuple[bool, str]:
    if _unit_matches_parent(unit_path, parent_sku):
        return True, "sku_match"
    units = _unit_files(work_root)
    # 01に素材が1枚だけのとき、SKU名不一致でも許可（透過PNG検証用）
    if loose_unit_match and len(units) == 1 and units[0].resolve() == unit_path.resolve():
        return True, "single_unit_fallback"
    return False, f"unit_sku_mismatch:{unit_path.name}"


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Amazon Pillow paste batch (出品CK)")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument(
        "--from-sheets",
        action="store_true",
        help="Google Sheets からマスタ直読（C1 config.local.json の spreadsheet_id）",
    )
    src.add_argument("--master-csv", type=Path, help="ローカルCSV（互換）")
    ap.add_argument(
        "--c1-config",
        type=Path,
        default=None,
        help="C1 config パス（省略時 tools/c1_hpc_packaged/config.local.json）",
    )
    ap.add_argument("--spreadsheet-id", default="", help="スプシID上書き")
    ap.add_argument(
        "--master-sheet",
        default="",
        help="マスタシート名（省略時 ▼商品マスタ(人間作業用)）",
    )
    ap.add_argument("--parent-sku", default="", help="省略時は全親（--checked-only時）")
    ap.add_argument(
        "--checked-only",
        action="store_true",
        help="出品CKレ点のセット行のみ（親のみレ点→全子）",
    )
    ap.add_argument("--min-n", type=int, default=1, help="この個数未満はスキップ（既定1＝1個セット含む）")
    ap.add_argument(
        "--loose-unit-match",
        action="store_true",
        help="01に単体が1枚だけのとき、ファイル名にSKUが無くても許可",
    )
    ap.add_argument(
        "--auto-bind-bases",
        action="store_true",
        help="量産前に bind_amazon_base_to_parents（楽天メインVision照合→SKUリネーム）を実行",
    )
    ap.add_argument(
        "--aspect",
        choices=list(ASPECT_CHOICES),
        default=ASPECT_SQUARE,
        help="量産パターン: square=正方形(本線) / portrait=縦長 / landscape=横長",
    )
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="省略時は 07.白抜きの置き場（C→D本線）",
    )
    ap.add_argument(
        "--test-out",
        action="store_true",
        help="出力を 05…/00.テスト出力 にする（検証用）",
    )
    ap.add_argument(
        "--name-style",
        choices=("prod", "debug"),
        default="prod",
        help="prod={子SKU}_amazon.jpg / debug=BATCH_paste_…長名",
    )
    ap.add_argument("--stem-prefix", default="BATCH_paste")
    ap.add_argument("--canvas", type=int, default=1200)
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    aspect = str(args.aspect)
    # landscape / portrait / square いずれも本線（landscape=合格固定）

    work_root = args.work_root or default_work_root()
    if args.out_dir is not None:
        out_dir = args.out_dir
    elif args.test_out:
        out_dir = work_root / SUB_TEST_OUT
    else:
        out_dir = DEFAULT_AMAZON_07_OUT
    out_dir.mkdir(parents=True, exist_ok=True)
    meta_dir = out_dir / SUB_META
    meta_dir.mkdir(parents=True, exist_ok=True)
    LOG.info(
        "out_dir=%s meta_dir=%s aspect=%s(%s) nameStyle=%s",
        out_dir,
        meta_dir,
        aspect,
        ASPECT_LABEL_JA.get(aspect, aspect),
        args.name_style,
    )

    sheets_info: Dict[str, Any] = {}
    master_source = ""

    if args.from_sheets:
        from sheets_master import count_true_in_column, fetch_master_rows

        rows, sheets_info = fetch_master_rows(
            config_path=args.c1_config,
            spreadsheet_id=args.spreadsheet_id or "",
            master_sheet=args.master_sheet or "",
        )
        true_ck = count_true_in_column(rows, "出品CK")
        sheets_info["trueCk"] = true_ck
        master_source = f"sheets:{sheets_info.get('spreadsheetId')}:{sheets_info.get('masterSheet')}"
        LOG.info(
            "from-sheets rows=%s trueCk=%s sheet=%r",
            sheets_info.get("rowCount"),
            true_ck,
            sheets_info.get("masterSheet"),
        )
        if not args.checked_only:
            ap.error("--from-sheets 時は --checked-only を指定してください（全件は危険）")
        children, food_map = load_checked_set_children_from_rows(
            rows, parent_sku=args.parent_sku or ""
        )
        if args.auto_bind_bases:
            from bind_amazon_base_to_parents import (
                enrich_master_rows_rakuten_main_urls,
                load_rakuten_shop_id,
                run_bind,
            )
            from sheets_master import load_sheets_settings

            settings = load_sheets_settings(
                args.c1_config,
                spreadsheet_id=args.spreadsheet_id or "",
                master_sheet=args.master_sheet or "",
            )
            shop_id = load_rakuten_shop_id(
                config_path=args.c1_config,
                spreadsheet_id=args.spreadsheet_id or "",
            )
            rows = enrich_master_rows_rakuten_main_urls(
                rows,
                config_path=args.c1_config,
                spreadsheet_id=args.spreadsheet_id or "",
                master_sheet=args.master_sheet or "",
                shop_id=shop_id,
            )
            LOG.info("auto-bind-bases: Vision照合で01PNGを親SKUへリネーム shop=%s", shop_id)
            bind_summary = run_bind(
                work_root=work_root,
                rows=rows,
                dry_run=False,
                overall_min=70,
                shape_min=55,
                color_min=55,
                credentials_path=settings["credentials_path"],
                token_path=settings["token_path"],
                shop_id=shop_id,
            )
            bind_meta = meta_dir / (
                "BASE_BIND_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".json"
            )
            bind_meta.write_text(
                json.dumps(bind_summary, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            LOG.info(
                "auto-bind done renamed=%s skip=%s -> %s",
                (bind_summary.get("counts") or {}).get("renamed"),
                (bind_summary.get("counts") or {}).get("skip"),
                bind_meta,
            )
    else:
        master_source = str(args.master_csv)
        if args.checked_only:
            children, food_map = load_checked_set_children(
                args.master_csv, parent_sku=args.parent_sku or ""
            )
        else:
            if not args.parent_sku:
                ap.error("--checked-only なしのときは --parent-sku が必要です")
            children, is_food = load_set_children_for_parent(
                args.master_csv, args.parent_sku, checked_only=False
            )
            food_map = {args.parent_sku: is_food}

    targets = [c for c in children if int(c.set_count) >= int(args.min_n)]
    LOG.info(
        "batch targets=%s (from %s, min_n=%s checked_only=%s source=%s)",
        len(targets),
        len(children),
        args.min_n,
        args.checked_only,
        master_source,
    )

    octas_path = resolve_octas(work_root, None)
    results: List[Dict[str, Any]] = []
    ok = 0
    err = 0
    skip = 0

    for ch in sorted(targets, key=lambda x: (x.parent_sku, x.set_count)):
        is_food = bool(ch.is_food or food_map.get(ch.parent_sku, False))
        entry: Dict[str, Any] = {
            "parentSku": ch.parent_sku,
            "childSku": ch.child_sku,
            "setCount": ch.set_count,
            "isFood": is_food,
        }
        try:
            bases = resolve_amazon_product_bases(work_root, ch.parent_sku, None)
        except FileNotFoundError as e:
            entry["status"] = "skip"
            entry["reason"] = f"no_base:{e}"
            results.append(entry)
            skip += 1
            LOG.warning("SKIP %s n=%s no base", ch.child_sku, ch.set_count)
            continue

        allowed, match_reason = _allow_unit_for_parent(
            bases.unit,
            ch.parent_sku,
            work_root,
            loose_unit_match=bool(args.loose_unit_match),
        )
        if not allowed:
            entry["status"] = "skip"
            entry["reason"] = match_reason
            entry["unitFile"] = bases.unit.name
            results.append(entry)
            skip += 1
            LOG.warning(
                "SKIP %s n=%s unit file does not match parent (%s)",
                ch.child_sku,
                ch.set_count,
                bases.unit.name,
            )
            continue
        if match_reason == "single_unit_fallback":
            LOG.warning(
                "unit match fallback (single file in 01): %s for parent=%s",
                bases.unit.name,
                ch.parent_sku,
            )
            entry["unitMatch"] = match_reason

        if is_food and not octas_path:
            entry["status"] = "skip"
            entry["reason"] = "food_but_no_octas"
            results.append(entry)
            skip += 1
            LOG.warning("SKIP %s food but Octas missing", ch.child_sku)
            continue

        fill_min, band = fill_target_for_n(ch.set_count)
        if args.name_style == "prod":
            out_path = out_dir / f"{ch.child_sku}_amazon.jpg"
        else:
            stem = (
                f"{args.stem_prefix}_{_safe_stem(ch.parent_sku)}"
                f"_n{ch.set_count}_{_safe_stem(ch.child_sku)}"
            )
            out_path = out_dir / f"{stem}_pillow_amazon_set{ch.set_count}.jpg"
        LOG.info(
            "==== paste parent=%s n=%s child=%s food=%s base=%s aspect=%s ====",
            ch.parent_sku,
            ch.set_count,
            ch.child_sku,
            is_food,
            bases.mode,
            aspect,
        )
        try:
            meta = run_paste_file(
                hero_path=bases.hero,
                unit_path=bases.unit,
                set_count=ch.set_count,
                out_path=out_path,
                octas_path=octas_path if is_food else None,
                canvas_size=args.canvas,
                fill_min=fill_min,
                layout_mode="edge_fill",
                aspect=aspect,
            )
            meta["baseMode"] = bases.mode
            meta["fillTarget"] = fill_min
            meta["fillBand"] = band
            meta["childSku"] = ch.child_sku
            meta["parentSku"] = ch.parent_sku
            meta["aspect"] = aspect
            meta["aspectLabelJa"] = ASPECT_LABEL_JA.get(aspect, aspect)
            # 画像直下にはJPGのみ。メタは _meta/
            json_path = meta_dir / f"{out_path.stem}.json"
            json_path.write_text(
                json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            oct = meta.get("octas") or {}
            entry.update(
                {
                    "status": "ok",
                    "output": str(out_path),
                    "json": str(json_path),
                    "fillPass": meta.get("fillPass"),
                    "sealOpaqueRatioActual": oct.get("sealOpaqueRatioActual"),
                    "overlapMeasured": oct.get("overlapMeasured"),
                    "baseMode": bases.mode,
                    "aspect": aspect,
                }
            )
            results.append(entry)
            ok += 1
        except Exception as e:
            LOG.exception("FAIL %s n=%s", ch.child_sku, ch.set_count)
            entry["status"] = "error"
            entry["reason"] = str(e)
            results.append(entry)
            err += 1

    summary = {
        "mode": "amazon_paste_batch",
        "at": datetime.now(timezone.utc).isoformat(),
        "masterSource": master_source,
        "sheets": sheets_info or None,
        "checkedOnly": bool(args.checked_only),
        "minN": int(args.min_n),
        "looseUnitMatch": bool(args.loose_unit_match),
        "aspect": aspect,
        "aspectLabelJa": ASPECT_LABEL_JA.get(aspect, aspect),
        "outDir": str(out_dir),
        "metaDir": str(meta_dir),
        "nameStyle": args.name_style,
        "ok": ok,
        "skip": skip,
        "error": err,
        "results": results,
    }
    summary_path = (
        meta_dir
        / f"{args.stem_prefix}_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    )
    summary_path.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "summary": str(summary_path),
                "outDir": str(out_dir),
                "ok": ok,
                "skip": skip,
                "error": err,
                "aspect": aspect,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    LOG.info("batch done ok=%s skip=%s err=%s -> %s", ok, skip, err, summary_path.name)
    return 0 if err == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
