# -*- coding: utf-8 -*-
"""
サブ画像パーツ組み合わせ PoC（B-③連動）

ロジック:
  1) B-③シートから JAN の競合画像URLを取得・キャッシュ
  2) 文字列を読んで意図分類（選定専用。画像上の文字は改変しない）
  3) shipping/店舗/問合せ等は除外
  4) パーツ自動提案 → テストは --auto-accept（本番は人手確定へ切替可）
  5) 安全BG（ベージュ／モノトーン）上にパーツ組み合わせ
  6) 商品あり／なしの両型を出力（商品は必須ではない）

テスト出力のみ（マスタ／R2未書込）。
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from PIL import Image

from b3_comp_catalog import B3_SHEET_TITLE, cache_b3_images, fetch_b3_rows, parse_b3_image_rows
from gemini_image import load_api_key, make_client
from sub_image_intent import classify_competitor_image
from sub_image_parts import (
    SUB_PATTERNS_10,
    annotate_proposals,
    auto_accept_useful,
    clean_part_fringe,
    compose_parts_board,
    propose_parts,
)
from work_paths import default_work_root, meta_dir_for

LOG = logging.getLogger("set_main_image.sub_image_part_poc")


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _default_own(work: Path) -> Optional[Path]:
    base = work / "01.amazon白抜きベース"
    if base.is_dir():
        files = sorted(base.glob("*.png")) + sorted(base.glob("*.jpg"))
        files = [p for p in files if p.is_file() and not p.name.startswith("_")]
        if files:
            return files[0]
    return None


def _own_for_product(work: Path, product_name: str) -> Optional[Path]:
    """商品名ヒントで白抜きを探す（見つからなければ default）。"""
    base = work / "01.amazon白抜きベース"
    if not base.is_dir():
        return _default_own(work)
    keys = []
    name = product_name or ""
    if "唐辛子" in name or "七味" in name:
        keys = ["唐辛子", "七味", "トウガラシ"]
    elif "牛丼" in name:
        keys = ["牛丼", "缶飯"]
    elif "焼鳥" in name or "焼き鳥" in name:
        keys = ["焼鳥", "焼き鳥", "缶飯"]
    files = sorted(base.glob("*.png")) + sorted(base.glob("*.jpg"))
    files = [p for p in files if p.is_file() and not p.name.startswith("_")]
    for k in keys:
        for p in files:
            if k in p.name:
                return p
    return files[0] if files else None


def _build_contact_10(raw_dir: Path, out_path: Path) -> None:
    files = sorted(raw_dir.glob("P*.jpg"))
    if not files:
        files = sorted(raw_dir.glob("*.jpg"))
    if not files:
        return
    thumbs = []
    for f in files[:10]:
        im = Image.open(f).convert("RGB")
        im.thumbnail((360, 360), Image.Resampling.LANCZOS)
        thumbs.append(im)
    cols = 5
    rows = (len(thumbs) + cols - 1) // cols
    cell = 370
    sheet = Image.new("RGB", (cols * cell + 20, rows * cell + 20), (30, 30, 30))
    for i, im in enumerate(thumbs):
        r, c = divmod(i, cols)
        sheet.paste(im, (10 + c * cell, 10 + r * cell))
    sheet.save(out_path, quality=90)

def write_logic_md(out_dir: Path, jan: str, *, from_curate: bool) -> None:
    src = "人間確認シート（採用CK=TRUEのみ）" if from_curate else f"B-③ `{B3_SHEET_TITLE}`（JANフィルタ）"
    text = f"""# サブ画像パーツ組み合わせ PoC

JAN: **{jan}**（他JANの画像は混在させない）

## ロジック

1. 入力: {src}
2. **文字列読取→意図分類**（選定専用）
3. 配送・店舗・問合せ等は **除外**
4. パーツ自動提案（吹き出し・囲みは画像のまま）
5. 文字は **改変しない**
6. 背景はベージュ／モノトーン（強色禁止）
7. 商品実物は **必須ではない**（あり／なし両出力）

## part_mode

- 初期: `auto_propose` + テスト時 `--auto-accept`
- 精度が低ければ: `human_box`（人手で領域指定）へ切替
"""
    (out_dir / "INSTRUCTION_LOGIC.md").write_text(text, encoding="utf-8")


def run(
    *,
    jan: str,
    work_root: Path,
    own_path: Optional[Path],
    max_images: int,
    auto_accept: bool,
    skip_classify: bool,
    curated_items: Optional[List[Dict[str, str]]] = None,
) -> Path:
    """
    curated_items: 人間確認シート由来の採用行（同一JANのみ）。指定時はB-③生データに戻らない。
    """
    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_root = work_root / "00.テスト出力" / "sub_image_parts_poc" / f"{jan}_{run_id}"
    dir_cache = out_root / "01_b3_cache"
    dir_class = out_root / "02_classify"
    dir_parts = out_root / "03_parts_propose"
    dir_raw = out_root / "04_compose"
    dir_annot = out_root / "05_annot"
    for d in (dir_cache, dir_class, dir_parts, dir_raw, dir_annot, meta_dir_for(out_root)):
        d.mkdir(parents=True, exist_ok=True)

    from_curate = curated_items is not None
    write_logic_md(out_root, jan, from_curate=from_curate)

    if curated_items is not None:
        # 混在防止: 他JANが紛れていたら落とす
        items = [x for x in curated_items if str(x.get("jan") or "").strip() == jan]
        if not items:
            raise SystemExit(f"キュレーション採用0件 JAN={jan}（シートの採用CKを確認）")
        from b3_comp_catalog import download_url, url_cache_name

        cached = []
        for it in items[:max_images]:
            url = it["url"]
            path = dir_cache / url_cache_name(url)
            try:
                download_url(url, path)
            except Exception as e:
                LOG.warning("skip: %s", e)
                continue
            cached.append(
                {
                    "jan": jan,
                    "kind": it.get("kind") or "",
                    "listing_key": it.get("listingKey") or "",
                    "image_index": int(str(it.get("imageIndex") or "0") or 0),
                    "url": url,
                    "path": str(path),
                    "productName": it.get("productName") or "",
                    "fromCurate": True,
                }
            )
        LOG.info("curate accepted JAN=%s n=%d", jan, len(cached))
    else:
        LOG.info("fetch B-③ sheet（JAN=%s のみ）…", jan)
        rows = fetch_b3_rows()
        parsed = parse_b3_image_rows(rows, jan=jan)
        if not parsed:
            raise SystemExit(
                f"B-③に JAN={jan} の画像URLがありません。先に B-③ を実行してください。"
            )
        LOG.info("B-③ URL rows for JAN=%s: %d", jan, len(parsed))
        cached = cache_b3_images(parsed, dir_cache, limit=max_images)
    (meta_dir_for(out_root) / "b3_cached.json").write_text(
        json.dumps(cached, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if not cached:
        raise SystemExit("画像ダウンロードに失敗しました。")

    client = make_client(load_api_key())

    classifications: List[Dict[str, Any]] = []
    for item in cached:
        p = Path(item["path"])
        # 人間レ点済み（from_curate）は再分類で落とさない
        if skip_classify or from_curate:
            c = {
                "intentLabel": "product_benefit",
                "decision": "use",
                "ocrTextPreview": "(adopted-from-b3)" if from_curate else "(skip)",
                "reasonJa": "b3_sub_adopt" if from_curate else "skip_classify",
                "hasProductVisible": True,
                "layoutHint": "mixed",
                "path": str(p),
            }
        else:
            c = classify_competitor_image(p, client=client)
        c["b3"] = {
            "kind": item.get("kind"),
            "imageIndex": item.get("image_index"),
            "listingKey": item.get("listing_key"),
            "url": item.get("url"),
        }
        classifications.append(c)
        stem = p.stem
        (dir_class / f"{stem}_intent.json").write_text(
            json.dumps(c, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    used = [c for c in classifications if c.get("decision") == "use"]
    rejected = [c for c in classifications if c.get("decision") == "reject"]
    review = [c for c in classifications if c.get("decision") == "review"]
    LOG.info("classify use=%d reject=%d review=%d", len(used), len(rejected), len(review))

    propose_targets = used + review
    if not propose_targets:
        LOG.warning("採用候補0件。rejectのみの場合はレビュー用に全件提案します。")
        propose_targets = classifications

    # パーツ提案は最大8枚（速度）
    propose_targets = propose_targets[:8]

    parts_docs: List[Dict[str, Any]] = []
    source_for_compose: List = []
    for c in propose_targets:
        p = Path(c["path"])
        doc = propose_parts(p, client=client)
        doc["intent"] = {
            "intentLabel": c.get("intentLabel"),
            "decision": c.get("decision"),
            "ocrTextPreview": c.get("ocrTextPreview"),
        }
        if auto_accept:
            doc = auto_accept_useful(doc, max_parts=4)
        (dir_parts / f"{p.stem}_parts.json").write_text(
            json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        src = Image.open(p).convert("RGBA")
        # レビュー用: 縁除去サンプルも保存
        annot = annotate_proposals(src, doc)
        annot.convert("RGB").save(dir_annot / f"{p.stem}_parts_boxes.jpg", quality=90)
        for pi, part in enumerate(doc.get("parts") or []):
            if not part.get("accepted"):
                continue
            from sub_image_parts import crop_part

            cleaned = clean_part_fringe(crop_part(src, part["box"]))
            cleaned.convert("RGBA").save(dir_parts / f"{p.stem}_part{pi}_clean.png")
        parts_docs.append(doc)
        if (doc.get("acceptedCount") or 0) > 0:
            source_for_compose.append((src, doc))

    product_name = ""
    if curated_items:
        product_name = str(curated_items[0].get("productName") or "")
    own_path_resolved = own_path
    if own_path_resolved is None:
        own_path_resolved = _own_for_product(work_root, product_name)

    own_im = None
    if own_path_resolved and own_path_resolved.is_file():
        own_im = Image.open(own_path_resolved).convert("RGBA")
        own_im.convert("RGB").save(dir_cache / f"OWN_{own_path_resolved.name}")

    composed_meta: List[Dict[str, Any]] = []
    if not source_for_compose:
        LOG.error("組み合わせ可能なパーツがありません。")
    else:
        for pi, pat in enumerate(SUB_PATTERNS_10):
            need_prod = bool(pat.get("include_product"))
            if need_prod and own_im is None:
                # 自社が無い場合は商品なし版にフォールバック
                need_prod = False
            title = f"{pat['id']}｜{pat['ja']}"
            board = compose_parts_board(
                source_images=source_for_compose,
                own_product=own_im,
                include_product=need_prod,
                bg_index=int(pat.get("bg") or 2),
                title=title,
                slots=list(pat.get("slots") or []),
                max_parts=int(pat.get("max_parts") or 4),
                source_mode=str(pat.get("source_mode") or "all"),
                product_anchor=str(pat.get("product_anchor") or "br"),
                pattern_index=pi,
                clean_edges=True,
            )
            out_name = f"{pat['id']}.jpg"
            outp = dir_raw / out_name
            board.convert("RGB").save(outp, quality=92)
            composed_meta.append(
                {
                    "file": str(outp),
                    "patternId": pat["id"],
                    "ja": pat["ja"],
                    "includeProduct": need_prod,
                    "bgIndex": pat.get("bg"),
                }
            )
            LOG.info("pattern %s saved", pat["id"])
        _build_contact_10(dir_raw, out_root / "04_contact_10patterns.jpg")

    meta = {
        "runId": run_id,
        "jan": jan,
        "productName": product_name,
        "b3Sheet": B3_SHEET_TITLE,
        "cached": len(cached),
        "classify": {
            "use": len(used),
            "reject": len(rejected),
            "review": len(review),
        },
        "classifications": classifications,
        "partsDocs": parts_docs,
        "composed": composed_meta,
        "patterns": [p["id"] for p in SUB_PATTERNS_10],
        "autoAccept": auto_accept,
        "partMode": "auto_propose",
        "edgeClean": True,
        "own": str(own_path_resolved) if own_path_resolved else None,
        "outRoot": str(out_root),
        "fromCurate": from_curate,
        "notes": [
            "JAN isolation",
            "10 patterns per product",
            "part fringe cleaned before composite",
            "on-image text not rewritten",
        ],
    }
    (meta_dir_for(out_root) / "run_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    LOG.info("done → %s", out_root)
    return out_root


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="B-③連動サブ画像パーツ組み合わせ PoC（JAN分離）")
    ap.add_argument(
        "--jan",
        action="append",
        default=None,
        help="対象JAN（複数指定可。商品ごとに分離実行）",
    )
    ap.add_argument(
        "--from-curate-sheet",
        action="store_true",
        help="B-③『サブ採用CK』レ点のみ使う（旧名・互換。中身はB-③右列）",
    )
    ap.add_argument(
        "--from-b3-adopt",
        action="store_true",
        help="B-③のサブ採用CKレ点のみ使う",
    )
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--own", type=Path, default=None, help="自社白抜き（任意・全JAN共通）")
    ap.add_argument("--max-images", type=int, default=12, help="JANあたりDL・パーツ提案する最大枚数")
    ap.add_argument(
        "--auto-accept",
        action="store_true",
        help="パーツ提案を自動採用（人手レビュー前のテスト用）",
    )
    ap.add_argument("--skip-classify", action="store_true", help="分類スキップ（デバッグ）")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    work = args.work_root or default_work_root()
    # --own 未指定なら商品名から白抜きを解決（run内）
    own = args.own
    jans = [str(j).strip() for j in (args.jan or []) if str(j).strip()]
    if not jans:
        LOG.error("--jan を1つ以上指定してください")
        return 2

    curated_by_jan: Dict[str, List[Dict[str, str]]] = {}
    use_adopt = bool(args.from_curate_sheet or args.from_b3_adopt)
    if use_adopt:
        from sub_image_b3_curate import read_accepted_from_b3

        curated_by_jan = read_accepted_from_b3(jans=jans)
        LOG.info(
            "B-③ サブ採用CK loaded: %s",
            {k: len(v) for k, v in curated_by_jan.items()},
        )

    outs: List[str] = []
    for jan in jans:
        curated = curated_by_jan.get(jan) if use_adopt else None
        if use_adopt and not curated:
            LOG.warning("JAN=%s サブ採用レ点0件のためスキップ", jan)
            continue
        out = run(
            jan=jan,
            work_root=work,
            own_path=own,
            max_images=max(1, int(args.max_images)),
            auto_accept=bool(args.auto_accept),
            skip_classify=bool(args.skip_classify) or use_adopt,
            curated_items=curated,
        )
        outs.append(str(out))
        print(str(out))
    return 0 if outs else 3


if __name__ == "__main__":
    sys.exit(main())
