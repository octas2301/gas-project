# -*- coding: utf-8 -*-
"""
セットMAIN — AI / レイヤー PoC

Amazon: 見本レイアウト拘束の AI 生成（詳細トレース付き）
楽天:   ベース不変のレイヤー重ねのみ（フル生成禁止）

  python ai_compose_poc.py --engine gemini --mall amazon --set-count 4 --food --stem POC2
  python ai_compose_poc.py --mall rakuten --set-count 4 --unit 個 --font-id helvetica_bold --stem POC2
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from ai_prompts import prompt_amazon
from gemini_image import (
    generate_with_references as gemini_generate,
    make_client,
    save_as_jpeg,
)
from master_sets import load_set_children_for_parent
from model_policy import resolve_model_id
from openai_image import (
    generate_with_references as openai_generate,
    resolve_openai_image_model,
)
from rakuten_layer import compose_rakuten_layered
from trace_log import file_fingerprint, write_run_trace
from transparent_bg import ensure_transparent_product
from work_paths import (
    default_work_root,
    resolve_amazon_base,
    resolve_amazon_reference,
    resolve_digit_layer,
    resolve_octas,
    resolve_rakuten_base,
    resolve_rakuten_reference,
)

LOG = logging.getLogger("set_main_image.ai_poc")


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def default_test_out(work_root: Path) -> Path:
    return work_root / "00.テスト出力"


def run_amazon_ai(
    *,
    engine: str,
    set_count: int,
    work_root: Path,
    out_dir: Path,
    parent_sku: str,
    is_food: bool,
    model: Optional[str],
    base: Optional[Path],
    reference: Optional[Path],
    octas: Optional[Path],
    stem: str,
    prefer_unused_refs: Optional[set] = None,
    out_name: Optional[str] = None,
    master_csv: Optional[Path] = None,
    competitor_fact: Optional[dict] = None,
) -> dict:
    from amazon_blueprint import decision_to_dict, select_amazon_blueprint
    from competitor_fact import CompetitorFactResult, resolve_competitor_fact_image
    from octas_prep import prepare_octas_seal
    from work_paths import resolve_amazon_product_bases

    bases = resolve_amazon_product_bases(work_root, parent_sku, base)
    bp = bases.hero
    unit_path = bases.unit
    has_unit_base = bases.mode == "paired_hero_unit" and bases.hero.resolve() != bases.unit.resolve()
    cache_dir = work_root / "01.amazon白抜きベース" / "_transparent_cache"
    # 白背景JPGでも透過化してから渡す（見本ルール）
    bp_alpha = ensure_transparent_product(bp, cache_dir)
    unit_alpha = (
        ensure_transparent_product(unit_path, cache_dir) if has_unit_base else bp_alpha
    )
    decision = select_amazon_blueprint(
        set_count,
        work_root,
        prefer_unused=prefer_unused_refs,
        explicit=reference,
    )
    rp = decision.path
    paths: List[Path] = [bp_alpha]
    roles = ["IMAGE_1_HERO_PRODUCT" if has_unit_base else "IMAGE_1_PRODUCT_TRANSPARENT"]
    if has_unit_base:
        paths.append(unit_alpha)
        roles.append("IMAGE_1B_UNIT_PRODUCT")
    paths.append(rp)
    roles.append("IMAGE_2_LAYOUT_BLUEPRINT")

    octas_meta = None
    octas_tilt = None
    octas_raw = resolve_octas(work_root, octas) if is_food else None
    has_octas = bool(octas_raw)
    if has_octas and octas_raw is not None:
        octas_path, octas_meta = prepare_octas_seal(
            octas_raw,
            work_root / "99.octas期限管理シール素材" / "_tilt_cache",
            seed=hash(f"{parent_sku}:{set_count}") % (2**31),
        )
        octas_tilt = float(octas_meta.get("tiltDeg") or 0)
        paths.append(octas_path)
        roles.append("IMAGE_3_OCTAS_SEAL")

    # 競合事実参照（マスタASIN/URL → 同一商品ゲート）
    fact_res: Optional[CompetitorFactResult] = None
    if competitor_fact is not None:
        # バッチから渡された解決済み結果
        fact_res = CompetitorFactResult(
            used=bool(competitor_fact.get("used")),
            path=Path(competitor_fact["path"])
            if competitor_fact.get("path")
            else None,
            asin=str(competitor_fact.get("asin") or ""),
            source_url=str(competitor_fact.get("source_url") or ""),
            source=str(competitor_fact.get("source") or ""),
            skip_reason=str(competitor_fact.get("skip_reason") or ""),
            candidates_tried=list(competitor_fact.get("candidates_tried") or []),
        )
        if competitor_fact.get("match"):
            from competitor_fact import FactMatchScore

            m = competitor_fact["match"]
            fact_res.match = FactMatchScore(**m)
    elif master_csv is not None:
        fact_res = resolve_competitor_fact_image(
            master_csv=master_csv,
            parent_sku=parent_sku,
            base_product_path=bp,
            work_root=work_root,
        )

    has_fact = bool(fact_res and fact_res.used and fact_res.path and fact_res.path.is_file())
    if has_fact and fact_res and fact_res.path:
        paths.append(fact_res.path)
        roles.append("IMAGE_FACT_COMPETITOR_REALITY")
        LOG.info(
            "[amazon/fact] USE asin=%s source=%s overall=%s file=%s",
            fact_res.asin,
            fact_res.source,
            fact_res.match.overall if fact_res.match else None,
            fact_res.path.name,
        )
    elif fact_res is not None:
        LOG.info(
            "[amazon/fact] SKIP reason=%s asin=%s tried=%s",
            fact_res.skip_reason,
            fact_res.asin,
            len(fact_res.candidates_tried),
        )

    prompt = prompt_amazon(
        set_count=set_count,
        is_food=is_food,
        has_octas=has_octas,
        blueprint_file=decision.file_name,
        pattern_hint=decision.pattern_hint,
        layout_intent_ja=decision.reason_ja,
        has_fact_ref=has_fact,
        fact_asin=(fact_res.asin if fact_res else "") or "",
        has_unit_base=has_unit_base,
        octas_tilt_deg=octas_tilt,
    )
    input_meta = []
    for role, p in zip(roles, paths):
        fp = file_fingerprint(p)
        fp["role"] = role
        if role == "IMAGE_2_LAYOUT_BLUEPRINT":
            fp["blueprintDecision"] = decision_to_dict(decision)
        if role == "IMAGE_FACT_COMPETITOR_REALITY" and fact_res:
            fp["competitorFact"] = fact_res.to_dict()
        if role == "IMAGE_3_OCTAS_SEAL" and octas_meta:
            fp["octasPrep"] = octas_meta
        input_meta.append(fp)
        LOG.info("[amazon/%s] %s -> %s sha=%s", engine, role, p.name, fp["sha256_16"])

    LOG.info(
        "[amazon/bases] mode=%s hero=%s unit=%s",
        bases.mode,
        bases.hero.name,
        bases.unit.name,
    )
    LOG.info("[amazon/logic] %s", decision.reason_ja)

    notes = [
        "Amazon uses AI layout transfer; IMAGE_2 is LAYOUT_BLUEPRINT.",
        "HARD: aspect ratio lock — no stretch/squash; paste only.",
        "HARD: no invented lid/label text — only provided bases (+ fact check).",
        "HARD N<=4: same on-canvas scale for all units (no perspective shrink).",
        f"PRODUCT bases mode={bases.mode} hero={bp.name} unit={unit_path.name}",
        f"HERO transparent: {bp_alpha.name}",
        (
            f"UNIT transparent: {unit_alpha.name}"
            if has_unit_base
            else "UNIT same as HERO (no 単体 tag pair)"
        ),
        "Each image is preceded by a role label in the API payload.",
        f"Layout blueprint selected: {rp.name}",
        f"Blueprint reason: {decision.reason_ja}",
        f"band={decision.band} preferredPattern={decision.preferred_pattern} "
        f"patternHint={decision.pattern_hint} score={decision.score}",
        (
            f"Octas tilt={octas_tilt:+.1f}° floating seal (not stuck on label)"
            if has_octas and octas_tilt is not None
            else "Octas unused"
        ),
        (
            f"Competitor FACT used: asin={fact_res.asin} source={fact_res.source}"
            if has_fact and fact_res
            else f"Competitor FACT skipped: {(fact_res.skip_reason if fact_res else 'n/a')}"
        ),
    ]

    if engine == "gemini":
        client = make_client()
        model_id = resolve_model_id(client, explicit=model)
        raw, api_meta = gemini_generate(
            client=client,
            model_id=model_id,
            prompt=prompt,
            image_paths=paths,
            image_roles=roles,
        )
        engine_name = "nano_banana_gemini_image"
        api_path = api_meta.get("apiPath")
    elif engine == "openai":
        model_id = resolve_openai_image_model(model)
        # OpenAI edits: roles are embedded into prompt header
        role_header = "\n".join(
            f"- File order {i+1}: {roles[i]} = {paths[i].name}" for i in range(len(paths))
        )
        full_prompt = prompt + "\nFILE ORDER:\n" + role_header
        raw, used_openai_model = openai_generate(
            prompt=full_prompt,
            image_paths=paths,
            model=model_id,
        )
        model_id = used_openai_model
        prompt = full_prompt
        engine_name = "chatgpt_openai_image"
        api_path = "images.edits"
    else:
        raise SystemExit(f"amazon 非対応 engine={engine}")

    final_name = out_name or f"{stem}_{engine}_amazon_set{set_count}_ai.jpg"
    out_path = out_dir / final_name
    out_path, im_meta = save_as_jpeg(raw, out_path, quality=85)

    from work_paths import meta_dir_for

    meta_dir = meta_dir_for(out_dir)
    trace_path = meta_dir / f"{out_path.stem}_trace.json"
    write_run_trace(
        trace_path,
        mall="amazon",
        engine=engine_name,
        mode="amazon_ai_layout_transfer",
        model_id=model_id,
        set_count=set_count,
        unit="",
        prompt=prompt,
        inputs=input_meta,
        api_path=api_path,
        notes=notes,
        extra={
            "imageMeta": im_meta,
            "output": str(out_path),
            "blueprintDecision": decision_to_dict(decision),
            "productBase": str(bp),
            "parentSku": parent_sku,
            "competitorFact": fact_res.to_dict() if fact_res else None,
        },
    )

    from fill_metrics import fill_target_for_n, ink_fill_ratio

    fill_min, fill_band = fill_target_for_n(set_count)
    fill_meas = ink_fill_ratio(out_path)
    fill_ok = fill_meas["inkFillRatio"] >= fill_min
    LOG.info(
        "[amazon/fill] ratio=%.4f target=%.2f (%s) pass=%s maxMargin=%.3f",
        fill_meas["inkFillRatio"],
        fill_min,
        fill_band,
        fill_ok,
        fill_meas["maxMargin"],
    )

    report = {
        "runAt": datetime.now(timezone.utc).isoformat(),
        "engine": engine_name,
        "mode": "amazon_ai_layout_transfer",
        "modelId": model_id,
        "mall": "amazon",
        "setCount": set_count,
        "parentSku": parent_sku,
        "productBase": str(bp),
        "output": str(out_path),
        "trace": str(trace_path),
        "apiPath": api_path,
        "inputs": input_meta,
        "imageMeta": im_meta,
        "blueprintDecision": decision_to_dict(decision),
        "logicSummaryJa": decision.reason_ja,
        "competitorFact": fact_res.to_dict() if fact_res else None,
        "productBases": bases.to_dict(),
        "octasPrep": octas_meta,
        "tuning": {
            "activeKnob": "inkFillMin",
            "inkFillMin": fill_min,
            "band": fill_band,
            "measured": fill_meas,
            "pass": fill_ok,
            "smallN_sameSize": set_count <= 4,
            "smallN_strongerOverlap": set_count <= 4,
        },
    }
    (meta_dir / f"{out_path.stem}.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    LOG.info("wrote %s trace=%s blueprint=%s", out_path, trace_path.name, decision.file_name)
    return report


def run_rakuten_layer(
    *,
    set_count: int,
    unit: str,
    font_id: str,
    work_root: Path,
    out_dir: Path,
    parent_sku: str,
    base: Optional[Path],
    reference: Optional[Path],
    digit_layer: Optional[Path],
    stem: str,
) -> dict:
    bp = resolve_rakuten_base(work_root, parent_sku, base)
    style_ref = None
    try:
        style_ref = resolve_rakuten_reference(work_root, parent_sku, reference)
    except FileNotFoundError:
        LOG.warning("04.楽天見本が空です（書体親和の目視比較用。合成自体はベース＋レイヤ）")

    dig = resolve_digit_layer(work_root, parent_sku, digit_layer)
    im, meta = compose_rakuten_layered(
        base_path=bp,
        set_count=set_count,
        unit=unit,
        font_id=font_id,
        digit_layer_path=dig,
        style_ref_path=style_ref,
        work_root=work_root,
    )

    out_name = f"{stem}_layer_rakuten_set{set_count}.jpg"
    out_path = out_dir / out_name
    out_dir.mkdir(parents=True, exist_ok=True)
    if im.mode != "RGB":
        im = im.convert("RGB")
    im.save(out_path, format="JPEG", quality=85, optimize=True)

    from work_paths import meta_dir_for

    meta_dir = meta_dir_for(out_dir)

    input_meta = [{"role": "BASE_LOCKED", **file_fingerprint(bp)}]
    if style_ref:
        input_meta.append({"role": "STYLE_REF_FONT_BALANCE_EYEBALL", **file_fingerprint(style_ref)})
    if dig:
        input_meta.append({"role": "DIGIT_POSITION_LAYER", **file_fingerprint(dig)})

    notes = [
        "Rakuten = Python 3-tier layer only (number + unit + セット). No AI redraw.",
        "Unit from master header バリエーション単位 (or --unit override).",
        "Typography ratios in layout_rules.json badgeTypography.",
        f"badgeText={meta.get('badgeText')!r} fontId={font_id}",
    ]

    trace_path = meta_dir / f"{out_path.stem}_trace.json"
    write_run_trace(
        trace_path,
        mall="rakuten",
        engine="pillow_layer",
        mode="rakuten_layer_only",
        model_id=None,
        set_count=set_count,
        unit=unit,
        prompt="(no generative prompt — pixel layer compose)",
        inputs=input_meta,
        api_path=None,
        notes=notes,
        extra=meta,
    )

    report = {
        "runAt": datetime.now(timezone.utc).isoformat(),
        "engine": "pillow_layer",
        "mode": "rakuten_layer_only",
        "mall": "rakuten",
        "setCount": set_count,
        "unit": unit,
        "badgeText": meta.get("badgeText"),
        "fontId": font_id,
        "output": str(out_path),
        "trace": str(trace_path),
        "inputs": input_meta,
    }
    (meta_dir / f"{out_path.stem}.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    LOG.info("wrote %s trace=%s badge=%r", out_path, trace_path.name, meta.get("badgeText"))
    return report


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="セットMAIN PoC (Amazon AI / 楽天Pythonレイヤ)")
    ap.add_argument("--mall", choices=("amazon", "rakuten"), required=True)
    ap.add_argument("--set-count", type=int, default=None, help="省略時は --master-csv + --child-sku から取得可")
    ap.add_argument(
        "--engine",
        choices=("gemini", "openai", "both", "layer"),
        default=None,
        help="amazon: gemini/openai/both。rakuten: layer 固定（省略可）",
    )
    ap.add_argument(
        "--unit",
        default="",
        help="楽天金丸の単位。空ならマスタ『バリエーション単位』を使用（--master-csv時）",
    )
    ap.add_argument("--master-csv", type=Path, default=None, help="マスタCSV（ヘッダ名で列解決）")
    ap.add_argument("--child-sku", default="", help="マスタから単位・セット数を拾う子SKU")
    ap.add_argument("--font-id", default="gothic", help="グリフ生成時の書体（Canva PNGがあれば未使用）。太字推奨=gothic")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--out-dir", type=Path, default=None)
    ap.add_argument("--parent-sku", default="")
    ap.add_argument("--food", action="store_true")
    ap.add_argument("--model", default=None)
    ap.add_argument("--base", type=Path, default=None)
    ap.add_argument("--reference", type=Path, default=None)
    ap.add_argument("--octas", type=Path, default=None)
    ap.add_argument("--digit-layer", type=Path, default=None)
    ap.add_argument("--stem", default="POC")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    work_root = args.work_root or default_work_root()
    out_dir = args.out_dir or default_test_out(work_root)
    reports: List[dict] = []
    errors: List[dict] = []

    set_count = args.set_count
    unit = (args.unit or "").strip()
    unit_source = "cli" if unit else None

    if args.master_csv:
        if not args.parent_sku:
            raise SystemExit("--master-csv 使用時は --parent-sku が必要です")
        children, _food = load_set_children_for_parent(args.master_csv, args.parent_sku)
        if not children:
            raise SystemExit(f"マスタに子がありません parent={args.parent_sku}")
        row = None
        if args.child_sku:
            for ch in children:
                if ch.child_sku == args.child_sku:
                    row = ch
                    break
            if row is None:
                raise SystemExit(f"child-sku が見つかりません: {args.child_sku}")
        else:
            # set-count 指定があれば一致行、なければ先頭
            if set_count is not None:
                for ch in children:
                    if ch.set_count == set_count:
                        row = ch
                        break
            row = row or children[0]
        if set_count is None:
            set_count = row.set_count
        if not unit and row.unit:
            unit = row.unit
            unit_source = "master:バリエーション単位"
        LOG.info(
            "master resolved child=%s set=%s unit=%r source=%s",
            row.child_sku,
            set_count,
            unit,
            unit_source,
        )

    if set_count is None:
        raise SystemExit("--set-count か --master-csv が必要です")

    if args.mall == "rakuten":
        try:
            rep = run_rakuten_layer(
                set_count=set_count,
                unit=unit,
                font_id=args.font_id,
                work_root=work_root,
                out_dir=out_dir,
                parent_sku=args.parent_sku,
                base=args.base,
                reference=args.reference,
                digit_layer=args.digit_layer,
                stem=args.stem,
            )
            rep["unitSource"] = unit_source or ("cli-empty" if not unit else "cli")
            reports.append(rep)
        except Exception as e:
            LOG.exception("rakuten layer failed")
            errors.append({"engine": "layer", "error": str(e)})
    else:
        eng = args.engine or "both"
        engines = ["gemini", "openai"] if eng == "both" else [eng]
        if eng == "layer":
            raise SystemExit("amazon に --engine layer は使えません")
        for engine in engines:
            try:
                model = None if eng == "both" else args.model
                reports.append(
                    run_amazon_ai(
                        engine=engine,
                        set_count=set_count,
                        work_root=work_root,
                        out_dir=out_dir,
                        parent_sku=args.parent_sku,
                        is_food=bool(args.food),
                        model=model,
                        base=args.base,
                        reference=args.reference,
                        octas=args.octas,
                        stem=args.stem,
                        master_csv=args.master_csv,
                    )
                )
            except SystemExit as e:
                errors.append({"engine": engine, "error": str(e)})
                LOG.error("[%s] %s", engine, e)
            except Exception as e:
                errors.append({"engine": engine, "error": str(e)})
                LOG.exception("[%s] failed", engine)

    summary = {"reports": reports, "errors": errors}
    # Windows cp932 コンソール対策
    print(json.dumps(summary, ensure_ascii=True, indent=2))
    return 0 if reports and not errors else (0 if reports else 1)


if __name__ == "__main__":
    sys.exit(main())
