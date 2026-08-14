# -*- coding: utf-8 -*-
"""
サブ画像 AI 合成 PoC v2（Gemini本線 + falシズル）

- 競合に実在するLPテーマのみ選定 → テーマごとにバリエーション → 計10枚
- 段階1: route=hybrid なら T03(シズル)のみ fal（文字禁止）、他テーマは Gemini
- トンマナ: ベージュ固定 / 文字量抑制（成分表例外）
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

from PIL import Image

from b3_comp_catalog import B3_SHEET_TITLE, download_url, url_cache_name
from gemini_image import generate_with_references as gemini_generate
from gemini_image import make_client, save_as_jpeg
from model_policy import resolve_model_id
from fal_image import generate_with_references as fal_generate
from openai_image import generate_with_references as openai_generate
from photo_realism_rules import (
    photo_realism_block_en,
    photo_realism_block_ja,
    photo_realism_meta,
)
from tonmana_palette import (
    DEFAULT_ID as TONMANA_DEFAULT,
    normalize_base_color,
    palette_meta,
    resolve_tonmana,
    tonmana_block_en,
    tonmana_block_ja,
)
from package_truth import (
    format_aspect_lock_ja,
    measure_package_aspect,
    parse_truth_specs,
    resolve_package_truth,
)
from sub_image_b3_curate import read_accepted_from_b3
from sub_image_instruction_report import (
    JA_INSTRUCTIONS,
    build_instruction_list_board,
    build_instruction_response_board,
    build_theme_contact,
)
from sub_image_lp_themes import (
    aggregate_theme_hits,
    classify_image_themes,
    format_seo_hints_ja,
    jobs_summary_ja,
    seo_keywords_from_product_name,
)
from work_paths import (
    DEFAULT_MASTER_CSV,
    DEFAULT_RAKUTEN_UPLOAD_DIR,
    default_work_root,
    meta_dir_for,
)

LOG = logging.getLogger("set_main_image.sub_image_ai_compose")

# 後方互換: 既定ベージュ文言（tonmana_palette.beige と同趣旨）
TONMANA_BEIGE = resolve_tonmana("beige")["ja"]

PACKAGE_LOCK = """
【最重要・絶対厳禁／PACKAGE_LOCK（商品パッケージ改変不可）】
- IMAGE_PACKAGE_TRUTH が商品パッケージの唯一の正本である（単体／N=1相当。缶・瓶・ボトル・箱・パウチ可）。
- 正本の色相・彩度・印刷柄・ラベル絵柄・ラベル文字・ロゴ・形状・質感を一切変更しない。
- 正本パッケージ表面へのOCR・再描字・ラベル差し替え・リブランド・新パッケージ創作は禁止。
- N≥2のセットMAINコラージュを色・縦横比の正本にしてはならない。
- 変えてよいのは背景・レイアウト・外側の説明パネル・吹き出し・バッジのみ。
- HARD BAN: Do NOT recolor, redraw, OCR-rewrite, or alter PACKAGE_TRUTH packaging. Keep hue, label art, printed text, logo, shape, and aspect ratio identical.

【終了前チェックリスト（必須）】
□ パッケージ色相が正本と同一か
□ ラベル柄・印刷文字が正本のままか（読み替え・創作なし）
□ パッケージ縦横比が PACKAGE_ASPECT_LOCK どおりか（潰しなし）
□ 外側パネルの文字だけが新規／再配置対象になっているか
"""

TEXT_LIMITS_JA = """
① 文字量ハード上限（スマホ可読）
- 見出し: 全角12文字以内
- 本文: 最大3行
- 吹き出し／コールアウト: 最大2個
- 栄養・成分: 3要点まで、または別枠の成分表スロットに限定（密な微小文字の全面埋め禁止）
- 小さな文字の詰め込み禁止。遠くから読める大きさ。
"""

# fal に流すテーマ（視覚的シズルのみ）
FAL_SIZZLE_THEME_IDS: Set[int] = {3}
DEFAULT_OPENAI_QUALITY = "medium"


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def build_ja_prompt(
    *,
    product_name: str,
    job: Dict[str, Any],
    n_comp_refs: int,
    ocr_bundle: str,
    aspect_lock_ja: str = "",
    seo_hints_ja: str = "",
    human_feedback: str = "",
    base_color: str = TONMANA_DEFAULT,
) -> str:
    """画像生成モデルへの日本語指示（PACKAGE_TRUTH＋文字量上限込み）。"""
    sec = job.get("secondaryThemeId")
    sec_line = (
        f"- 補助テーマ（secondary）は T{int(sec):02d} のみ軽く触れてよい。主役にしない。第3テーマ禁止。"
        if sec
        else "- 補助テーマなし。主テーマ以外を前面に出さない。テーマ総数は1。"
    )
    src = job.get("source") or "competitor"
    invent_line = ""
    if src == "invented":
        invent_line = (
            "\n【想像スロット】競合に無いフェーズ補完。表現の型のみ。"
            "ランキング・累計販売・具体価格・特定産地の断定など事実数字の創作は禁止。\n"
        )
    prop = job.get("proposal") or ("a" if int(job.get("variant") or 1) == 1 else "b")
    feedback_block = ""
    if (human_feedback or "").strip():
        feedback_block = (
            "\n【人間フィードバック（再生成・最優先で反映）】\n"
            f"{human_feedback.strip()}\n"
            "上記以外のPACKAGE_LOCK・文字量上限・ベース色は維持すること。\n"
        )
    ton = resolve_tonmana(base_color)
    return f"""あなたは EC 商品ページ用のサブ画像を1枚作るデザイナーです（出口は楽天。Amazonは流用）。

【商品】{product_name}
【スロット】S{int(job.get('slotIndex') or job.get('jobIndex') or 0):02d} / phaseOrder={job.get('phaseOrder', '?')}
【主テーマ】T{job['themeId']:02d} [{job['phase']}] {job['themeName']}
【提案】AB{prop}（同じ主テーマで構図・パーツ配置を変える）
【テーマの狙い】{job['themeHint']}
【候補ソース】{src}
{invent_line}
【入力】
- IMAGE_PACKAGE_TRUTH: 商品パッケージ正本（色・柄・ラベル文字・縦横比の唯一の真実。表面は改変不可）
- IMAGE_COMP_*: 競合サブ参照 {n_comp_refs} 枚（外側パネル・レイアウト参考）
{feedback_block}
■ 日本語の必須指示（すべて守ること）

{TEXT_LIMITS_JA.strip()}

② 指定テーマのみ（1枚1主テーマ・被り最大2）
- この画像の主役は上記主テーマのみ。
{sec_line}
- 参照に無い実績数字・価格・誇大表現を新規に作らない。

③ トンマナ／ベース色（B-④選択）
{tonmana_block_ja(base_color)}

④ AI独自・改変は最大約50%／残りは競合パーツ（パッケージ外）
- 面積のおよそ半分まで: レイアウト再構成・余白・柔らかい影・ベース色背景・軽い装飾を生成してよい。
- 残りは競合の説明カード・吹き出し・バッジ・フレーム・シズル等を流用する。
- 下記「競合外側パネルOCR」を再配置して描く場合も競合パーツ扱いとする。
- パッケージ表面の文字は正本のまま。OCR対象外・再描字禁止。

{PACKAGE_LOCK.strip()}

{photo_realism_block_ja()}

{aspect_lock_ja}

【SEOヒント】
{seo_hints_ja or '（なし）'}

【競合外側パネルOCR（パッケージ表面は含めない・パネル／吹き出しのみ）】
{ocr_bundle or '（参照の外側パネル文字を短く優先）'}

【禁止】透かし、配送/店舗/問い合わせUI、他SKU商品の混入、カートボタン。架空の価格・ランキング。正本パッケージの色替え・潰し。

【出力】正方形の商用サブ画像。パーツ縁に元背景の白フチを残さない。終了前チェックリストを満たすこと。
（baseColor={ton['id']}）
"""


def build_fal_sizzle_prompt(
    *,
    product_name: str,
    job: Dict[str, Any],
    aspect_lock_ja: str = "",
    base_color: str = TONMANA_DEFAULT,
) -> str:
    """fal専用: シズル写真。文字・ロゴ・UIを描かせない（日本語描字禁止）。"""
    return f"""Create ONE photoreal EC secondary image focused on FOOD SIZZLE only.

PRODUCT: {product_name}
THEME: T{job['themeId']:02d} {job['themeName']} / variant V{job['variant']:02d}
GOAL: {job['themeHint']}

IMAGE_PACKAGE_TRUTH is the ONLY packaging identity. Keep package hue, label art, printed text, logo, shape, and aspect ratio identical. Do not recolor or squash the package.
{aspect_lock_ja}

{photo_realism_block_en()}

{tonmana_block_en(base_color)}

Show appetite appeal: steam, gloss, texture, pouring/sprinkling, or close-up ingredients.

HARD BAN — do NOT draw any of the following:
- Any text, letters, kanji, kana, numbers, captions, nutrition tables, labels as readable text
- Logos (Amazon, store marks), watermarks, price badges, UI, cart buttons
- Fake packaging typography or gibberish glyphs
- Shipping / store / contact graphics

If competitor refs contain text panels, IGNORE them; keep only the physical product / food from PACKAGE_TRUTH.
Output: square commercial photo. No borders from old backgrounds.
"""


def resolve_route_mode(route: str, providers: List[str]) -> str:
    """hybrid は gemini+fal が両方あるときだけ有効。それ以外は all。"""
    r = (route or "auto").strip().lower()
    if r == "all":
        return "all"
    if r == "hybrid":
        return "hybrid"
    # auto
    if "fal" in providers and "gemini" in providers:
        return "hybrid"
    return "all"


def providers_for_job(
    job: Dict[str, Any],
    *,
    providers: List[str],
    route_mode: str,
) -> List[str]:
    """
    hybrid: T03 → fal のみ / その他 → fal 以外（gemini, openai）
    all: 指定 providers すべて
    """
    if route_mode != "hybrid":
        return list(providers)
    tid = int(job.get("themeId") or 0)
    if tid in FAL_SIZZLE_THEME_IDS:
        return [p for p in providers if p == "fal"]
    return [p for p in providers if p != "fal"]


def _save_bytes_jpeg(data: bytes, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    im = Image.open(BytesIO(data))
    if im.mode in ("RGBA", "P"):
        bg = Image.new("RGB", im.size, (245, 240, 230))
        if im.mode == "P":
            im = im.convert("RGBA")
        bg.paste(im, mask=im.split()[-1] if im.mode == "RGBA" else None)
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")
    im.save(path, format="JPEG", quality=92, optimize=True)


def _ocr_bundle(job: Dict[str, Any], max_chars: int = 900) -> str:
    parts = []
    for i, s in enumerate(job.get("ocrSnippets") or [], 1):
        s = str(s).strip()
        if s:
            parts.append(f"[{i}] {s}")
    text = "\n".join(parts)
    return text[:max_chars]


def _auto_export_compose(
    out_root: Path,
    *,
    jan: str,
    auto_export: bool,
    export_out_dir: Optional[Path],
    export_pick: str,
    export_provider: str,
    master_csv: Optional[Path],
    child_sku: str = "",
) -> Optional[Dict[str, Any]]:
    if not auto_export:
        return None
    from export_sub_images_for_rakuten_matrix import (
        export_for_human_review,
        write_export_plan,
    )

    out_dir = Path(export_out_dir) if export_out_dir else DEFAULT_RAKUTEN_UPLOAD_DIR
    # 目視既定: 品番キー1件 × 最大10枚（pick=ab）。全子複製しない。
    pick = (export_pick or "ab").strip() or "ab"
    try:
        rows, key_info = export_for_human_review(
            compose_dir=out_root,
            out_dir=out_dir,
            jan=jan,
            child_sku=child_sku,
            master_csv=master_csv or DEFAULT_MASTER_CSV,
            pick=pick,
            provider=export_provider or "openai",
            max_subs=10,
        )
        key = str(key_info.get("key") or "")
        plan = write_export_plan(
            out_root,
            out_dir=out_dir,
            child_skus=[key] if key else [],
            pick=pick,
            provider=export_provider or "openai",
            jan=jan,
            product_key=key,
            mode="review",
            key_via=str(key_info.get("via") or ""),
        )
        LOG.info(
            "auto-export(review) done jan=%s key=%s via=%s files=%s out=%s plan=%s",
            jan,
            key,
            key_info.get("via"),
            len(rows),
            out_dir,
            plan,
        )
        return {
            "ok": True,
            "mode": "review",
            "outDir": str(out_dir),
            "productKey": key,
            "keyVia": key_info.get("via"),
            "childSkus": [key] if key else [],
            "exported": len(rows),
            "plan": str(plan),
        }
    except SystemExit as e:
        msg = str(e) if e.args else "SystemExit"
        LOG.warning("auto-export failed jan=%s: %s", jan, msg)
        return {"ok": False, "error": msg, "outDir": str(out_dir)}
    except Exception as e:
        LOG.warning("auto-export failed jan=%s: %s", jan, e)
        return {"ok": False, "error": str(e), "outDir": str(out_dir)}


def run_jan(
    *,
    jan: str,
    items: List[Dict[str, str]],
    work_root: Path,
    max_classify: int,
    max_refs_per_job: int,
    providers: List[str],
    gemini_model: Optional[str],
    openai_model: Optional[str],
    openai_quality: str,
    fal_endpoint: Optional[str],
    route: str,
    target_slots: int = 5,
    allow_invented: bool = True,
    package_truth_path: Optional[Path] = None,
    package_truth_dir: Optional[Path] = None,
    truth_map: Optional[Dict[str, Path]] = None,
    require_package_truth: bool = True,
    auto_export: bool = True,
    export_out_dir: Optional[Path] = None,
    export_pick: str = "ab",
    master_csv: Optional[Path] = None,
    export_child_sku: str = "",
    base_color: str = TONMANA_DEFAULT,
) -> Optional[Path]:
    _ = require_package_truth  # 正本無しは常にスキップ
    base_color = normalize_base_color(base_color)
    ton = resolve_tonmana(base_color)
    # PACKAGE_TRUTH（必須・無ければスキップ）
    truth = package_truth_path or resolve_package_truth(
        jan,
        work_root=work_root,
        truth_dir=package_truth_dir,
        truth_map=truth_map,
    )
    if not truth or not Path(truth).is_file():
        LOG.warning("JAN=%s: PACKAGE_TRUTH（単体／N=1正本）がありません。スキップ。", jan)
        return None

    truth_path = Path(truth)
    aspect_meta = measure_package_aspect(truth_path)
    aspect_lock_ja = format_aspect_lock_ja(aspect_meta)
    LOG.info(
        "PACKAGE_TRUTH jan=%s path=%s aspectWH=%s",
        jan,
        truth_path,
        aspect_meta.get("aspectWH"),
    )

    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    product_name = (items[0].get("productName") if items else "") or jan
    seo_kws = seo_keywords_from_product_name(product_name)
    seo_hints_ja = format_seo_hints_ja(product_name, seo_kws)
    out_root = work_root / "00.テスト出力" / "sub_image_ai_compose" / f"{jan}_{run_id}"
    dir_in = out_root / "01_refs"
    dir_g = out_root / "02_gemini"
    dir_o = out_root / "03_openai"
    dir_f = out_root / "04_fal"
    dir_rep = out_root / "05_instruction_report"
    for d in (dir_in, dir_g, dir_o, dir_f, dir_rep, meta_dir_for(out_root)):
        d.mkdir(parents=True, exist_ok=True)

    # 正本を出力ツリーにコピー（再生成用）
    truth_copy = dir_in / f"PACKAGE_TRUTH{truth_path.suffix.lower() or '.jpg'}"
    try:
        Image.open(truth_path).convert("RGB").save(truth_copy, quality=95)
    except Exception:
        import shutil

        shutil.copy2(truth_path, truth_copy)
    truth_path = truth_copy

    # 参照DL
    path_by_url: Dict[str, Path] = {}
    ref_paths: List[Path] = []
    for it in items[: max(1, max_classify)]:
        url = it["url"]
        dest = dir_in / url_cache_name(url)
        try:
            download_url(url, dest)
            path_by_url[url] = dest
            ref_paths.append(dest)
        except Exception as e:
            LOG.warning("skip download %s: %s", url[:80], e)
    if not ref_paths:
        raise SystemExit(f"JAN={jan} 参照画像DL失敗")

    route_mode = resolve_route_mode(route, providers)
    LOG.info("route_mode=%s providers=%s", route_mode, providers)

    # テーマ分類は常に Gemini テキストモデル（キー必要）— PACKAGE_TRUTH は分類しない
    gemini_client = make_client()
    gemini_mid = None
    if "gemini" in providers:
        gemini_mid = resolve_model_id(gemini_client, explicit=gemini_model)
        LOG.info("gemini image model=%s", gemini_mid)

    classifications: List[Dict[str, Any]] = []
    for p in ref_paths:
        classifications.append(classify_image_themes(p, client=gemini_client))
    hits = aggregate_theme_hits(classifications)
    from sub_image_lp_themes import select_lp_slots

    jobs = select_lp_slots(
        hits,
        target_slots=int(target_slots),
        proposals_per_slot=2,
        allow_invented=bool(allow_invented),
        product_name=product_name,
    )
    if not jobs:
        raise SystemExit(f"JAN={jan}: テーマスロットを1つも作れませんでした")

    # hybrid で fal があるのに T03 が無い場合、最終スロットを T03 に差し替え（A/Bとも）
    if (
        route_mode == "hybrid"
        and "fal" in providers
        and not any(int(j.get("themeId") or 0) in FAL_SIZZLE_THEME_IDS for j in jobs)
        and 3 in hits
    ):
        from sub_image_lp_themes import THEME_BY_ID, make_stem

        meta3 = THEME_BY_ID[3]
        bucket = hits[3]
        last_slot = int(jobs[-1].get("slotIndex") or 1)
        repl: List[Dict[str, Any]] = []
        for j in jobs:
            if int(j.get("slotIndex") or 0) != last_slot:
                repl.append(j)
                continue
            prop = j.get("proposal") or "a"
            repl.append(
                {
                    **j,
                    "stem": make_stem(slot=last_slot, theme_id=3, proposal=str(prop)),
                    "themeId": 3,
                    "phase": meta3["phase"],
                    "phaseOrder": meta3["phaseOrder"],
                    "phaseKey": meta3["phaseKey"],
                    "themeName": meta3["name"],
                    "themeSlug": meta3["nameSlug"],
                    "themeHint": meta3["hint"],
                    "contentCluster": meta3["contentCluster"],
                    "source": "competitor",
                    "refPaths": list(bucket["paths"]),
                    "ocrSnippets": list(bucket.get("ocrSnippets") or []),
                    "score": bucket["score"],
                    "count": bucket["count"],
                }
            )
        jobs = repl
        LOG.info("hybrid: forced T03 into last slot for fal sizzle lane")

    summary = jobs_summary_ja(jobs)
    route_note = (
        f"\n\n# 振り分け (route={route_mode})\n"
        f"- fal: テーマID {sorted(FAL_SIZZLE_THEME_IDS)} のみ（文字禁止シズル）\n"
        f"- gemini/openai: 上記以外\n"
        if route_mode == "hybrid"
        else f"\n\n# 振り分け (route={route_mode})\n- 全スロットで指定プロバイダを実行\n"
    )
    route_note += (
        f"\n# スロット設計\n"
        f"- target_slots={target_slots} × A/B → {len(jobs)}枚\n"
        f"- 分類: ページ全体のみ（パーツ単位は将来拡張）\n"
        f"- フォールバック: 競合 → phase被り → 想像(最大2)\n"
        f"- OpenAI quality={openai_quality}\n"
        f"- PACKAGE_TRUTH: {truth_path.name} aspectWH={aspect_meta.get('aspectWH')}\n"
        f"- SEO keywords: {', '.join(seo_kws) or '(none)'}\n"
    )
    (out_root / "THEMES.md").write_text(summary + route_note, encoding="utf-8")
    LOG.info("themes selected:\n%s", summary)

    theme_lines = [ln for ln in summary.splitlines() if ln.startswith("- ")]

    build_instruction_list_board(
        out_path=dir_rep / "00_日本語指示一覧.jpg",
        product_name=product_name,
        jan=jan,
        theme_lines=theme_lines or summary.splitlines()[:20],
    )

    # 指示本文をテキストでも提出
    instr_md = [
        f"# 日本語指示一覧 — {product_name} ({jan})",
        "",
        "## AIへの指示（本番プロンプト方針）",
        "",
    ]
    for inst in JA_INSTRUCTIONS:
        instr_md.append(f"### {inst['title']}")
        instr_md.append(inst["body"])
        instr_md.append("")
    instr_md.append("## 選択テーマ")
    instr_md.append(summary)
    (dir_rep / "INSTRUCTIONS_JA.md").write_text("\n".join(instr_md), encoding="utf-8")

    job_results: List[Dict[str, Any]] = []
    for job in jobs:
        refs = [Path(p) for p in job["refPaths"] if Path(p).is_file()][: max_refs_per_job]
        if not refs:
            refs = ref_paths[:max_refs_per_job]
        # 正本を先頭。OCRは競合のみ（正本は含めない）
        image_paths = [truth_path] + refs
        roles = ["IMAGE_PACKAGE_TRUTH"] + [f"IMAGE_COMP_{i+1}" for i in range(len(refs))]
        ocr = _ocr_bundle(job)
        prompt_gemini = build_ja_prompt(
            product_name=product_name,
            job=job,
            n_comp_refs=len(refs),
            ocr_bundle=ocr,
            aspect_lock_ja=aspect_lock_ja,
            seo_hints_ja=seo_hints_ja,
            base_color=base_color,
        )
        prompt_fal = build_fal_sizzle_prompt(
            product_name=product_name,
            job=job,
            aspect_lock_ja=aspect_lock_ja,
            base_color=base_color,
        )
        stem = job["stem"]
        use_providers = providers_for_job(job, providers=providers, route_mode=route_mode)
        rec: Dict[str, Any] = {
            **job,
            "promptJa": prompt_gemini,
            "promptFal": prompt_fal,
            "refsUsed": [str(p) for p in refs],
            "packageTruth": str(truth_path),
            "packageAspect": aspect_meta,
            "seoKeywords": list(seo_kws),
            "routeMode": route_mode,
            "assignedProviders": use_providers,
        }
        LOG.info(
            "job %s theme=%s → providers=%s",
            stem,
            job.get("themeId"),
            use_providers,
        )

        if "gemini" in use_providers and gemini_mid:
            outp = dir_g / f"{stem}.jpg"
            try:
                data, api_meta = gemini_generate(
                    client=gemini_client,
                    model_id=gemini_mid,
                    prompt=prompt_gemini,
                    image_paths=image_paths,
                    image_roles=roles,
                    aspect_ratio="1:1",
                    image_size="1K",
                )
                save_as_jpeg(data, outp)
                rec["gemini"] = {
                    "ok": True,
                    "path": str(outp),
                    "model": gemini_mid,
                    "api": api_meta.get("apiPath"),
                }
                LOG.info("gemini %s %s", jan, stem)
            except Exception as e:
                rec["gemini"] = {"ok": False, "error": str(e)}
                LOG.warning("gemini fail %s %s: %s", jan, stem, e)

        if "openai" in use_providers:
            outp = dir_o / f"{stem}.jpg"
            try:
                data, used_model = openai_generate(
                    prompt=prompt_gemini,
                    image_paths=image_paths,
                    model=openai_model,
                    size="1024x1024",
                    quality=openai_quality or DEFAULT_OPENAI_QUALITY,
                )
                _save_bytes_jpeg(data, outp)
                rec["openai"] = {
                    "ok": True,
                    "path": str(outp),
                    "model": used_model,
                    "quality": openai_quality or DEFAULT_OPENAI_QUALITY,
                }
                LOG.info(
                    "openai %s %s model=%s q=%s",
                    jan,
                    stem,
                    used_model,
                    openai_quality or DEFAULT_OPENAI_QUALITY,
                )
            except Exception as e:
                rec["openai"] = {"ok": False, "error": str(e)}
                LOG.warning("openai fail %s %s: %s", jan, stem, e)

        if "fal" in use_providers:
            outp = dir_f / f"{stem}.jpg"
            try:
                data, used_ep = fal_generate(
                    prompt=prompt_fal,
                    image_paths=image_paths,
                    endpoint=fal_endpoint,
                )
                _save_bytes_jpeg(data, outp)
                rec["fal"] = {
                    "ok": True,
                    "path": str(outp),
                    "endpoint": used_ep,
                    "promptMode": "sizzle_no_text",
                }
                LOG.info("fal %s %s endpoint=%s", jan, stem, used_ep)
            except Exception as e:
                rec["fal"] = {"ok": False, "error": str(e)}
                LOG.warning("fal fail %s %s: %s", jan, stem, e)

        job_results.append(rec)

    contact_dirs = []
    if any("gemini" in (r.get("assignedProviders") or []) for r in job_results):
        contact_dirs.append(("Gemini", dir_g))
    if any("fal" in (r.get("assignedProviders") or []) for r in job_results):
        contact_dirs.append(("fal.ai sizzle", dir_f))
    if any("openai" in (r.get("assignedProviders") or []) for r in job_results):
        contact_dirs.append(("OpenAI", dir_o))
    if contact_dirs:
        build_theme_contact(
            out_path=out_root / "04_contact_themes_both.jpg",
            jobs=jobs,
            provider_dirs=contact_dirs,
        )

    # 指示→対処ビジュアル（代表画像で①〜④を説明）
    def _first_ok(provider: str) -> Optional[str]:
        for r in job_results:
            block = r.get(provider) or {}
            if block.get("ok") and block.get("path"):
                return str(block["path"])
        return None

    def _theme_job_path(provider: str, theme_id: int) -> Optional[str]:
        for r in job_results:
            if int(r.get("themeId") or 0) != theme_id:
                continue
            block = r.get(provider) or {}
            if block.get("ok"):
                return str(block.get("path") or "")
        return _first_ok(provider)

    # 栄養テーマがあれば①の「成分表OK」例に使う
    nutrition_tid = 9 if 9 in hits else (jobs[0]["themeId"] if jobs else 0)
    sizzle_tid = 3 if 3 in hits else nutrition_tid
    response_rows = [
        {
            "title": JA_INSTRUCTIONS[0]["title"],
            "note": "見出し≤12／本文≤3行／コールアウト≤2",
            "ref": (hits.get(nutrition_tid) or {}).get("paths", [None])[0],
            "gemini": _theme_job_path("gemini", nutrition_tid),
            "fal": _theme_job_path("fal", nutrition_tid),
            "openai": _theme_job_path("openai", nutrition_tid),
        },
        {
            "title": JA_INSTRUCTIONS[1]["title"],
            "note": f"例: T{jobs[0]['themeId']:02d} {jobs[0]['themeName']}",
            "ref": jobs[0]["refPaths"][0] if jobs[0].get("refPaths") else None,
            "gemini": job_results[0].get("gemini", {}).get("path"),
            "fal": job_results[0].get("fal", {}).get("path"),
            "openai": job_results[0].get("openai", {}).get("path"),
        },
        {
            "title": JA_INSTRUCTIONS[2]["title"],
            "note": "全出力ベージュ基調（代表1枚）",
            "ref": ref_paths[0] if ref_paths else None,
            "gemini": _first_ok("gemini"),
            "fal": _first_ok("fal"),
            "openai": _first_ok("openai"),
        },
        {
            "title": JA_INSTRUCTIONS[3]["title"],
            "note": "競合パーツ＋AI再構成（シズル系があればそれを使用）",
            "ref": (hits.get(sizzle_tid) or {}).get("paths", [str(ref_paths[0])])[0],
            "gemini": _theme_job_path("gemini", sizzle_tid),
            "fal": _theme_job_path("fal", sizzle_tid),
            "openai": _theme_job_path("openai", sizzle_tid),
        },
    ]
    build_instruction_response_board(
        out_path=dir_rep / "01_指示への対処ビジュアル.jpg",
        product_name=product_name,
        jan=jan,
        rows=response_rows,
    )

    # テーマごとの対処行も追加ボード
    theme_rows = []
    for r in job_results:
        if str(r.get("proposal") or "") not in ("a", "1") and int(r.get("variant") or 0) != 1:
            continue
        stem = r["stem"]
        assigned = set(r.get("assignedProviders") or [])
        theme_rows.append(
            {
                "title": f"S{int(r.get('slotIndex') or 0):02d} T{r['themeId']:02d} {r['themeName']}"[:28],
                "note": f"P{r.get('phaseOrder','?')} [{r['phase']}] src={r.get('source')} → {','.join(r.get('assignedProviders') or [])}",
                "ref": (r.get("refPaths") or [None])[0],
                "gemini": str(dir_g / f"{stem}.jpg") if "gemini" in assigned else None,
                "fal": str(dir_f / f"{stem}.jpg") if "fal" in assigned else None,
                "openai": str(dir_o / f"{stem}.jpg") if "openai" in assigned else None,
            }
        )
    if theme_rows:
        build_instruction_response_board(
            out_path=dir_rep / "02_テーマ別_対処ビジュアル.jpg",
            product_name=product_name,
            jan=jan,
            rows=theme_rows,
        )

    (out_root / "INSTRUCTION.md").write_text(
        "\n".join(
            [
                f"# AI合成 v3 — {product_name} ({jan})",
                "",
                f"- スロット: {target_slots} × A/B（心理プロセス順）→ {len(jobs)}枚",
                "- 分類: ページ全体のみ（パーツ単位は将来拡張）",
                "- フォールバック: 競合 → phase被り → 想像(最大2・事実数字禁止)",
                f"- 振り分け: route={route_mode}",
                f"- OpenAI本線 quality={openai_quality}",
                f"- PACKAGE_TRUTH: {truth_path.name} aspectWH={aspect_meta.get('aspectWH')}",
                f"- トンマナ: ベース色={ton['id']}（{ton['label']}）／PACKAGE_LOCK厳守／正本OCR禁止",
                f"- 写真実写ルール: photo_realism_rules.py ({photo_realism_meta().get('version')})",
                "- 文字量: 見出し≤12／本文≤3行／コールアウト≤2／栄養3要点",
                "- AI改変: 最大約50%／外側パネルOCRのみ競合パーツ扱い",
                "- 提出ボード: `05_instruction_report/`",
                "- レビュー再生成: `python sub_image_review_loop.py --compose-dir …`（再生成後も楽天へ再export）",
                f"- 【人間目視の正】{DEFAULT_RAKUTEN_UPLOAD_DIR}",
                "",
                summary,
                route_note,
                "",
            ]
        ),
        encoding="utf-8",
    )

    export_info = None
    meta = {
        "jan": jan,
        "productName": product_name,
        "b3Sheet": B3_SHEET_TITLE,
        "version": 4,
        "tonmana": ton["id"],
        "baseColor": ton["id"],
        "baseColorLabel": ton["label"],
        "tonmanaPalette": palette_meta(),
        "packageLock": True,
        "packageTruth": str(truth_path),
        "packageAspect": aspect_meta,
        "seoKeywords": list(seo_kws),
        "aiGenerateMaxPct": 50,
        "targetSlots": target_slots,
        "allowInvented": allow_invented,
        "classifyScope": "page_only_competitor_no_package_truth_ocr",
        "routeMode": route_mode,
        "falSizzleThemeIds": sorted(FAL_SIZZLE_THEME_IDS),
        "classifications": classifications,
        "themeHits": {str(k): v for k, v in hits.items()},
        "jobs": job_results,
        "geminiModel": gemini_mid,
        "openaiModelRequested": openai_model or "gpt-image-2",
        "openaiQuality": openai_quality or DEFAULT_OPENAI_QUALITY,
        "falEndpointRequested": fal_endpoint or "fal-ai/flux-kontext/dev",
        "providers": providers,
        "jaInstructions": JA_INSTRUCTIONS,
        "outRoot": str(out_root),
        "autoExport": None,
        "humanReviewDir": str(DEFAULT_RAKUTEN_UPLOAD_DIR),
        "photoRealism": photo_realism_meta(),
    }
    # auto-export が run_meta を読むため、先に書く（後で autoExport を更新）
    meta_path = meta_dir_for(out_root) / "run_meta.json"
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    export_info = _auto_export_compose(
        out_root,
        jan=jan,
        auto_export=auto_export,
        export_out_dir=export_out_dir,
        export_pick=export_pick,
        export_provider=(providers[0] if providers else "openai"),
        master_csv=master_csv,
        child_sku=export_child_sku,
    )
    meta["autoExport"] = export_info
    meta["humanReviewDir"] = str(
        (export_info or {}).get("outDir") or DEFAULT_RAKUTEN_UPLOAD_DIR
    )
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info("done JAN=%s → %s humanReview=%s", jan, out_root, meta.get("humanReviewDir"))
    return out_root


def run_regen_from_queue(
    *,
    compose_dir: Path,
    queue_path: Path,
    providers: List[str],
    gemini_model: Optional[str],
    openai_model: Optional[str],
    openai_quality: str,
    fal_endpoint: Optional[str],
    max_refs_per_job: int = 4,
) -> Path:
    """
    レビューHTMLが書いた regen_queue.json に従い、チェック済み stem だけ再生成。
    """
    out_root = Path(compose_dir)
    meta_path = meta_dir_for(out_root) / "run_meta.json"
    if not meta_path.is_file():
        raise SystemExit(f"run_meta.json がありません: {meta_path}")
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    queue = json.loads(Path(queue_path).read_text(encoding="utf-8"))
    items = [x for x in (queue.get("items") or []) if x.get("checked")]
    if not items:
        LOG.info("再生成対象なし（チェック0）→ OK扱い")
        return out_root

    jan = str(meta.get("jan") or "")
    product_name = str(meta.get("productName") or jan)
    truth_path = Path(meta.get("packageTruth") or "")
    if not truth_path.is_file():
        # コピー先が相対化されていても 01_refs を探す
        cand = out_root / "01_refs"
        found = list(cand.glob("PACKAGE_TRUTH.*")) if cand.is_dir() else []
        if found:
            truth_path = found[0]
    if not truth_path.is_file():
        raise SystemExit("PACKAGE_TRUTH が再生成に必要ですが見つかりません")

    aspect_meta = meta.get("packageAspect") or measure_package_aspect(truth_path)
    aspect_lock_ja = format_aspect_lock_ja(aspect_meta)
    seo_kws = list(meta.get("seoKeywords") or seo_keywords_from_product_name(product_name))
    seo_hints_ja = format_seo_hints_ja(product_name, seo_kws)
    route_mode = str(meta.get("routeMode") or "all")
    jobs_by_stem = {str(j.get("stem")): j for j in (meta.get("jobs") or [])}

    dir_g = out_root / "02_gemini"
    dir_o = out_root / "03_openai"
    dir_f = out_root / "04_fal"
    for d in (dir_g, dir_o, dir_f):
        d.mkdir(parents=True, exist_ok=True)

    gemini_client = None
    gemini_mid = None
    if "gemini" in providers:
        gemini_client = make_client()
        gemini_mid = resolve_model_id(gemini_client, explicit=gemini_model)

    base_color = normalize_base_color(
        meta.get("baseColor") or meta.get("tonmana") or TONMANA_DEFAULT
    )
    LOG.info("regen base_color=%s", base_color)
    regen_log: List[Dict[str, Any]] = []
    for it in items:
        stem = str(it.get("stem") or "").strip()
        feedback = str(it.get("comment") or "").strip()
        provider = str(it.get("provider") or providers[0]).strip().lower()
        job = jobs_by_stem.get(stem)
        if not job:
            LOG.warning("stem不明スキップ: %s", stem)
            continue
        refs = [Path(p) for p in (job.get("refsUsed") or job.get("refPaths") or []) if Path(p).is_file()]
        refs = refs[: max(1, max_refs_per_job)]
        image_paths = [truth_path] + refs
        roles = ["IMAGE_PACKAGE_TRUTH"] + [f"IMAGE_COMP_{i+1}" for i in range(len(refs))]
        ocr = _ocr_bundle(job)
        prompt = build_ja_prompt(
            product_name=product_name,
            job=job,
            n_comp_refs=len(refs),
            ocr_bundle=ocr,
            aspect_lock_ja=aspect_lock_ja,
            seo_hints_ja=seo_hints_ja,
            human_feedback=feedback,
            base_color=base_color,
        )
        prompt_fal = build_fal_sizzle_prompt(
            product_name=product_name,
            job=job,
            aspect_lock_ja=aspect_lock_ja,
            base_color=base_color,
        )
        rec: Dict[str, Any] = {"stem": stem, "feedback": feedback, "provider": provider}
        try:
            if provider == "gemini" and gemini_client and gemini_mid:
                data, api_meta = gemini_generate(
                    client=gemini_client,
                    model_id=gemini_mid,
                    prompt=prompt,
                    image_paths=image_paths,
                    image_roles=roles,
                    aspect_ratio="1:1",
                    image_size="1K",
                )
                outp = dir_g / f"{stem}.jpg"
                save_as_jpeg(data, outp)
                rec.update({"ok": True, "path": str(outp), "api": api_meta.get("apiPath")})
            elif provider == "fal":
                data, used_ep = fal_generate(
                    prompt=prompt_fal + (f"\nHuman feedback: {feedback}" if feedback else ""),
                    image_paths=image_paths,
                    endpoint=fal_endpoint,
                )
                outp = dir_f / f"{stem}.jpg"
                _save_bytes_jpeg(data, outp)
                rec.update({"ok": True, "path": str(outp), "endpoint": used_ep})
            else:
                data, used_model = openai_generate(
                    prompt=prompt,
                    image_paths=image_paths,
                    model=openai_model,
                    size="1024x1024",
                    quality=openai_quality or DEFAULT_OPENAI_QUALITY,
                )
                outp = dir_o / f"{stem}.jpg"
                _save_bytes_jpeg(data, outp)
                rec.update({"ok": True, "path": str(outp), "model": used_model})
            # jobs メタ更新
            for j in meta.get("jobs") or []:
                if str(j.get("stem")) != stem:
                    continue
                j["promptJa"] = prompt
                j["humanFeedback"] = feedback
                j[provider] = {
                    "ok": True,
                    "path": rec.get("path"),
                    "regenerated": True,
                    "feedback": feedback,
                }
                break
            LOG.info("regen ok %s %s", stem, provider)
        except Exception as e:
            rec.update({"ok": False, "error": str(e)})
            LOG.warning("regen fail %s: %s", stem, e)
        regen_log.append(rec)

    meta["lastRegen"] = {
        "at": datetime.now(timezone.utc).isoformat(),
        "queue": str(queue_path),
        "results": regen_log,
    }
    # 再生成後は export_plan があれば楽天フォルダへ再投入（1話完結）
    try:
        from export_sub_images_for_rakuten_matrix import export_from_plan

        plan_p = meta_dir_for(out_root) / "export_plan.json"
        if plan_p.is_file():
            re_rows = export_from_plan(out_root)
            meta["lastRegen"]["reExport"] = {
                "ok": True,
                "files": len(re_rows),
                "plan": str(plan_p),
            }
            LOG.info("regen re-export files=%s", len(re_rows))
        else:
            LOG.warning("export_plan.json 無し — 楽天フォルダへの再exportをスキップ")
    except Exception as e:
        meta["lastRegen"]["reExport"] = {"ok": False, "error": str(e)}
        LOG.warning("regen re-export failed: %s", e)

    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (meta_dir_for(out_root) / "regen_log.json").write_text(
        json.dumps(regen_log, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    # キュー消化（チェッククリア）
    queue["items"] = []
    queue["clearedAt"] = datetime.now(timezone.utc).isoformat()
    Path(queue_path).write_text(json.dumps(queue, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info("regen done → %s", out_root)
    return out_root


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="サブ画像 AI合成 v4（PACKAGE_TRUTH＋文字量上限）")
    ap.add_argument("--jan", action="append", default=[])
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--max-classify", type=int, default=12, help="テーマ分類する採用画像の上限")
    ap.add_argument("--max-refs-per-job", type=int, default=4)
    ap.add_argument(
        "--providers",
        default="openai",
        help="gemini,openai,fal をカンマ区切り（本線: openai）",
    )
    ap.add_argument(
        "--route",
        default="all",
        choices=["auto", "hybrid", "all"],
        help="auto/hybrid= T03のみfal・他Gemini（gemini+fal時）。all=全プロバイダ全スロット",
    )
    ap.add_argument("--gemini-model", default=None)
    ap.add_argument("--openai-model", default=None)
    ap.add_argument(
        "--openai-quality",
        default=DEFAULT_OPENAI_QUALITY,
        choices=("low", "medium", "high"),
        help="本線は medium",
    )
    ap.add_argument("--target-slots", type=int, default=5, choices=(5, 6))
    ap.add_argument(
        "--allow-invented",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="不足フェーズの想像スロット（既定ON・最大2）",
    )
    ap.add_argument(
        "--fal-endpoint",
        default=None,
        help="既定 fal-ai/flux-kontext/dev（高品質は fal-ai/flux-pro/kontext）",
    )
    ap.add_argument(
        "--package-truth",
        action="append",
        default=[],
        help="正本指定 JAN=path または path（単一JAN時）",
    )
    ap.add_argument(
        "--package-truth-dir",
        type=Path,
        default=None,
        help="正本探索フォルダ（ファイル名にJAN推奨。SKU必須ではない）",
    )
    ap.add_argument(
        "--require-package-truth",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="正本が無いJANはスキップ（既定ON）",
    )
    ap.add_argument(
        "--regen-from",
        type=Path,
        default=None,
        help="compose出力ディレクトリ。regen_queue.json があるときチェック済みのみ再生成",
    )
    ap.add_argument(
        "--regen-queue",
        type=Path,
        default=None,
        help="再生成キューJSON（既定: <compose>/_meta/regen_queue.json）",
    )
    ap.add_argument(
        "--auto-export",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="合成後に楽天アップロードフォルダへ自動export（既定ON・1話完結）",
    )
    ap.add_argument(
        "--export-out-dir",
        type=Path,
        default=DEFAULT_RAKUTEN_UPLOAD_DIR,
        help="人間目視の正（既定: 02.楽天アップロード画像保存場所）",
    )
    ap.add_argument("--export-pick", default="ab", help="auto-export の A/B（a|b|ab）。目視既定=ab（最大10枚）")
    ap.add_argument(
        "--master-csv",
        type=Path,
        default=None,
        help="JAN→子SKU解決用（省略時はSheets→DEFAULT_MASTER_CSV）",
    )
    ap.add_argument(
        "--export-child-sku",
        default="",
        help="目視品番キーを1子SKUに固定（省略時はJAN→親→代表1子）",
    )
    ap.add_argument(
        "--base-color",
        "--tonmana",
        dest="base_color",
        default=TONMANA_DEFAULT,
        choices=("beige", "warm_white", "soft_gray"),
        help="B-④ベース色（beige|warm_white|soft_gray）。背景・カードのみ。既定=beige",
    )
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    providers = [p.strip().lower() for p in str(args.providers).split(",") if p.strip()]
    work = args.work_root or default_work_root()
    truth_map = parse_truth_specs(args.package_truth)

    if args.regen_from:
        q = args.regen_queue or (meta_dir_for(Path(args.regen_from)) / "regen_queue.json")
        out = run_regen_from_queue(
            compose_dir=Path(args.regen_from),
            queue_path=Path(q),
            providers=providers,
            gemini_model=args.gemini_model,
            openai_model=args.openai_model,
            openai_quality=str(args.openai_quality),
            fal_endpoint=args.fal_endpoint,
            max_refs_per_job=max(1, int(args.max_refs_per_job)),
        )
        print(str(out))
        print(f"HUMAN_REVIEW_DIR={DEFAULT_RAKUTEN_UPLOAD_DIR}")
        return 0

    jans = [str(j).strip() for j in (args.jan or []) if str(j).strip()]
    if not jans:
        ap.error("--jan が必要です（または --regen-from）")

    accepted = read_accepted_from_b3(jans=jans)
    LOG.info("adopted counts %s", {k: len(v) for k, v in accepted.items()})

    outs = []
    for jan in jans:
        items = accepted.get(jan) or []
        if not items:
            LOG.warning("JAN=%s レ点0件スキップ", jan)
            continue
        out = run_jan(
            jan=jan,
            items=items,
            work_root=work,
            max_classify=max(1, int(args.max_classify)),
            max_refs_per_job=max(1, int(args.max_refs_per_job)),
            providers=providers,
            gemini_model=args.gemini_model,
            openai_model=args.openai_model,
            openai_quality=str(args.openai_quality),
            fal_endpoint=args.fal_endpoint,
            route=str(args.route),
            target_slots=int(args.target_slots),
            allow_invented=bool(args.allow_invented),
            package_truth_dir=args.package_truth_dir,
            truth_map=truth_map,
            require_package_truth=bool(args.require_package_truth),
            auto_export=bool(args.auto_export),
            export_out_dir=Path(args.export_out_dir) if args.export_out_dir else DEFAULT_RAKUTEN_UPLOAD_DIR,
            export_pick=str(args.export_pick or "a"),
            master_csv=args.master_csv,
            export_child_sku=str(args.export_child_sku or ""),
            base_color=str(args.base_color or TONMANA_DEFAULT),
        )
        if out:
            outs.append(str(out))
            print(str(out))
    if outs:
        print(f"HUMAN_REVIEW_DIR={args.export_out_dir or DEFAULT_RAKUTEN_UPLOAD_DIR}")
    return 0 if outs else 3


if __name__ == "__main__":
    sys.exit(main())
