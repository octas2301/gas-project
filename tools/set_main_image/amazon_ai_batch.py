# -*- coding: utf-8 -*-
"""
Amazon セットMAIN — 見本＋AI 量産バッチ（ロジック／見本トレース付き）

  python amazon_ai_batch.py ^
    --master-csv "...\\master_export.csv" ^
    --parent-sku sanky-4538872285127-oya ^
    --engine gemini
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional, Set

from ai_compose_poc import run_amazon_ai
from competitor_fact import resolve_competitor_fact_image
from master_sets import load_set_children_for_parent
from work_paths import (
    default_work_root,
    move_amazon_base_to_processed,
    resolve_amazon_base,
)

LOG = logging.getLogger("set_main_image.amazon_batch")

DEFAULT_OUT_07 = Path(
    r"G:\マイドライブ\04.amazonカタログ作成（CSV一括UL）\07.白抜きの置き場（人間が入れる）"
)


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Amazon MAIN 見本+AI 量産")
    ap.add_argument("--master-csv", type=Path, required=True)
    ap.add_argument("--parent-sku", required=True)
    ap.add_argument("--engine", choices=("gemini", "openai"), default="gemini")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--out-dir", type=Path, default=None)
    ap.add_argument("--model", default=None)
    ap.add_argument("--base", type=Path, default=None)
    ap.add_argument("--checked-only", action="store_true", default=True)
    ap.add_argument("--no-checked-only", action="store_true")
    ap.add_argument("--move-base", action="store_true", default=True)
    ap.add_argument("--no-move-base", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    work_root = args.work_root or default_work_root()
    out_dir = args.out_dir or DEFAULT_OUT_07
    out_dir.mkdir(parents=True, exist_ok=True)
    checked_only = not args.no_checked_only
    move_base = not args.no_move_base

    children, is_food = load_set_children_for_parent(
        args.master_csv, args.parent_sku, checked_only=checked_only
    )
    targets = [ch for ch in children if ch.set_count >= 1]
    if not targets:
        raise SystemExit("N>=1 の対象がありません")

    base_path = resolve_amazon_base(work_root, args.parent_sku, args.base)
    LOG.info(
        "batch start parent=%s food=%s targets=%s base=%s out=%s",
        args.parent_sku,
        is_food,
        [(c.set_count, c.child_sku) for c in targets],
        base_path.name,
        out_dir,
    )

    # 競合事実参照はバッチ先頭で1回だけ解決（Nごとに再判定しない）
    fact = resolve_competitor_fact_image(
        master_csv=args.master_csv,
        parent_sku=args.parent_sku,
        base_product_path=base_path,
        work_root=work_root,
    )
    fact_dict = fact.to_dict()
    LOG.info(
        "competitor fact used=%s asin=%s reason=%s",
        fact.used,
        fact.asin,
        fact.skip_reason or "ok",
    )

    used_refs: Set[str] = set()
    reports = []
    errors = []
    logic_rows = []

    for ch in sorted(targets, key=lambda x: x.set_count):
        LOG.info("==== N=%s child=%s ====", ch.set_count, ch.child_sku)
        try:
            rep = run_amazon_ai(
                engine=args.engine,
                set_count=ch.set_count,
                work_root=work_root,
                out_dir=out_dir,
                parent_sku=args.parent_sku,
                is_food=is_food or ch.is_food,
                model=args.model,
                base=args.base or base_path,
                reference=None,
                octas=None,
                stem=ch.child_sku,
                prefer_unused_refs=used_refs,
                out_name=f"{ch.child_sku}_amazon.jpg",
                competitor_fact=fact_dict,
            )
            bd = rep.get("blueprintDecision") or {}
            used_refs.add(str(bd.get("file_name") or ""))
            reports.append(rep)
            logic_rows.append(
                {
                    "setCount": ch.set_count,
                    "childSku": ch.child_sku,
                    "blueprintFile": bd.get("file_name"),
                    "patternHint": bd.get("pattern_hint"),
                    "band": bd.get("band"),
                    "preferredPattern": bd.get("preferred_pattern"),
                    "score": bd.get("score"),
                    "reasonJa": bd.get("reason_ja") or rep.get("logicSummaryJa"),
                    "alternatives": bd.get("alternatives"),
                    "output": rep.get("output"),
                    "trace": rep.get("trace"),
                    "modelId": rep.get("modelId"),
                    "competitorFactUsed": bool((rep.get("competitorFact") or {}).get("used")),
                }
            )
            LOG.info(
                "OK N=%s blueprint=%s -> %s",
                ch.set_count,
                bd.get("file_name"),
                Path(rep["output"]).name,
            )
        except Exception as e:
            LOG.exception("FAIL N=%s child=%s", ch.set_count, ch.child_sku)
            errors.append(
                {
                    "setCount": ch.set_count,
                    "childSku": ch.child_sku,
                    "error": str(e),
                }
            )

    moved = None
    if move_base and reports and not errors:
        moved_path = move_amazon_base_to_processed(base_path, work_root)
        moved = str(moved_path) if moved_path else None
    elif move_base and errors:
        LOG.warning("エラーがあるためベースは処理済みへ移動しません")

    summary = {
        "runAt": datetime.now(timezone.utc).isoformat(),
        "mode": "amazon_ai_layout_transfer_batch",
        "parentSku": args.parent_sku,
        "isFood": is_food,
        "engine": args.engine,
        "productBase": str(base_path),
        "outDir": str(out_dir),
        "logic": logic_rows,
        "successCount": len(reports),
        "errorCount": len(errors),
        "errors": errors,
        "movedAmazonBase": moved,
        "competitorFact": fact_dict,
        "selectionPolicyJa": (
            "見本ファイル名にセット数が無いため、"
            "sample_scan_report の patternHint/leftShare と実測 ink 密度で "
            "N帯に合う LAYOUT_BLUEPRINT をスコア選定。"
            "希望構図は『左スタック＋右ヒーロー』（山積み centered は個数誤認リスクで低優先）。"
            "プロンプト硬制約: ヒーロー＝01ベース保持（縦横比・デザイン変更禁止）／"
            "個数完全一致・裏積み暗示禁止・列の規則性・動き＋画面埋め。"
            "競合事実参照: マスタ競合店ASIN／競合URLから画像取得→01とVision一致ゲート。"
            "不一致・取得失敗は参照スキップ（ASIN貼り付けKeepa用シートは使わない）。"
            "バッチ内は再利用をやや減点。"
        ),
    }
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    from work_paths import meta_dir_for

    meta_dir = meta_dir_for(out_dir)
    logic_path = meta_dir / f"AMAZON_AI_LOGIC_{args.parent_sku}_{stamp}.json"
    logic_path.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    # 作業ルートのテスト出力 _meta にもコピー
    try:
        from work_paths import SUB_META as _SM

        copy_dir = work_root / "00.テスト出力" / _SM
        copy_dir.mkdir(parents=True, exist_ok=True)
        (copy_dir / logic_path.name).write_text(
            logic_path.read_text(encoding="utf-8"), encoding="utf-8"
        )
    except OSError:
        pass

    LOG.info(
        "batch done ok=%s err=%s logic=%s",
        len(reports),
        len(errors),
        logic_path,
    )
    # Windows コンソールの cp932 対策: stdout は ASCII サマリのみ
    print(
        json.dumps(
            {
                "logicReport": str(logic_path),
                "successCount": len(reports),
                "errorCount": len(errors),
                "movedAmazonBase": moved,
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0 if reports and not errors else (0 if reports else 1)


if __name__ == "__main__":
    sys.exit(main())
