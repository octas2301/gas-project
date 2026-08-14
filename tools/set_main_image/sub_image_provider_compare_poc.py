# -*- coding: utf-8 -*-
"""
モデル比較 PoC（同一JAN・同一テーマ・同一日本語プロンプト）

例:
  # Gemini 3.1 Flash Image + OpenAI gpt-image-2 high（パッケージ厳禁）
  python sub_image_provider_compare_poc.py --jan 4538872281013 \\
    --providers gemini,openai \\
    --gemini-model gemini-3.1-flash-image \\
    --openai-quality high
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFont

from fal_image import FLUX2_PRO_EDIT_ENDPOINT
from fal_image import generate_with_references as fal_generate
from gemini_image import generate_with_references as gemini_generate
from gemini_image import make_client
from openai_image import generate_with_references as openai_generate
from b3_comp_catalog import download_url, url_cache_name
from sub_image_ai_compose_poc import build_ja_prompt, _ocr_bundle, _save_bytes_jpeg
from sub_image_b3_curate import read_accepted_from_b3
from sub_image_lp_themes import (
    aggregate_theme_hits,
    classify_image_themes,
    jobs_summary_ja,
    select_themes_for_jobs,
)
from work_paths import default_work_root, meta_dir_for

LOG = logging.getLogger("set_main_image.provider_compare")

DEFAULT_GEMINI_MODEL = "gemini-2.5-flash-image"
DEFAULT_OPENAI_MODEL = "gpt-image-2"
DEFAULT_OPENAI_QUALITY = "medium"
DEFAULT_FAL_ENDPOINT = FLUX2_PRO_EDIT_ENDPOINT

PACKAGE_LOCK_JA = """
【最重要・絶対厳禁／商品パッケージ・ロック】
- 商品本体のパッケージ（缶・瓶・ボトル・箱・蓋・ラベル印刷）の色・デザイン・ロゴ・文字組・配置・形状・質感を一切変更しないこと。
- 参照画像に写っている実物パッケージと同一の見た目を保つこと。再彩色・リブランド・ラベル差し替え・新しいパッケージの創作は禁止。
- 変えてよいのは背景・レイアウト・説明パネル・吹き出しなど「パッケージの外側」のみ。
- HARD BAN (EN): Do NOT change product packaging colors, label artwork, logo, print design, can/bottle/box shape, or lid. Keep the physical product package identity identical to the reference photos. Only surrounding layout/background/callouts may change.
"""


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _jp_font(size: int) -> ImageFont.ImageFont:
    for p in (
        Path(r"C:\Windows\Fonts\YuGothM.ttc"),
        Path(r"C:\Windows\Fonts\meiryo.ttc"),
        Path(r"C:\Windows\Fonts\msgothic.ttc"),
    ):
        if p.is_file():
            try:
                return ImageFont.truetype(str(p), size=size)
            except Exception:
                pass
    return ImageFont.load_default()


def _thumb(path: Path, size: Tuple[int, int]) -> Image.Image:
    w, h = size
    im = Image.open(path).convert("RGB")
    im.thumbnail((w, h), Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (w, h), (245, 240, 230))
    canvas.paste(im, ((w - im.width) // 2, (h - im.height) // 2))
    return canvas


def _slug(s: str) -> str:
    s = re.sub(r"[^\w.\-]+", "_", (s or "").strip())
    return s[:48] or "x"


def provider_labels(
    providers: Sequence[str],
    *,
    gemini_model: str,
    openai_quality: str,
) -> List[Tuple[str, str]]:
    out: List[Tuple[str, str]] = []
    for p in providers:
        if p == "fal":
            out.append(("fal", "FLUX.2 pro"))
        elif p == "gemini":
            short = gemini_model.replace("gemini-", "").replace("-flash-image", " Flash")
            out.append(("gemini", f"Gemini {short}"[:22]))
        elif p == "openai":
            out.append(("openai", f"gpt-image-2 {openai_quality}"[:22]))
    return out


def with_package_lock(prompt: str) -> str:
    return (prompt or "").rstrip() + "\n" + PACKAGE_LOCK_JA.strip() + "\n"


def build_grid(
    *,
    rows: List[Dict[str, Any]],
    providers: Sequence[Tuple[str, str]],
    out_path: Path,
    title: str,
) -> None:
    cell = 280
    cols = 1 + len(providers)
    n = len(rows)
    W = 16 + max(cols, 1) * cell
    H = 56 + max(n, 1) * (cell + 36)
    im = Image.new("RGB", (W, H), (28, 26, 24))
    d = ImageDraw.Draw(im)
    f = _jp_font(14)
    tf = _jp_font(18)
    d.text((12, 10), title[:90], fill=(240, 230, 210), font=tf)
    headers = ["テーマ"] + [ja for _, ja in providers]
    for ci, h in enumerate(headers):
        d.text((16 + ci * cell, 36), h[:18], fill=(200, 190, 170), font=f)

    for ri, row in enumerate(rows):
        y = 56 + ri * (cell + 36)
        d.text((16, y + 8), str(row.get("label") or "")[:22], fill=(220, 200, 170), font=f)
        for ci, (pid, _) in enumerate(providers):
            x = 16 + (ci + 1) * cell
            p = (row.get("paths") or {}).get(pid)
            if p and Path(p).is_file():
                im.paste(_thumb(Path(p), (cell - 12, cell - 12)), (x, y))
            else:
                box = Image.new("RGB", (cell - 12, cell - 12), (60, 50, 48))
                bd = ImageDraw.Draw(box)
                bd.text((20, cell // 2 - 20), "FAIL / SKIP", fill=(200, 120, 100), font=f)
                im.paste(box, (x, y))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    im.save(out_path, quality=90)


def run_compare(
    *,
    jan: str,
    items: List[Dict[str, str]],
    work_root: Path,
    max_classify: int,
    max_refs_per_job: int,
    total_jobs: int,
    providers: Sequence[str],
    gemini_model: str,
    openai_model: str,
    openai_quality: str,
    fal_endpoint: str,
    package_lock: bool,
) -> Path:
    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    product_name = (items[0].get("productName") if items else "") or jan
    out_root = work_root / "00.テスト出力" / "sub_image_provider_compare" / f"{jan}_{run_id}"
    dir_in = out_root / "01_refs"
    dir_fal = out_root / "02_fal_flux2_pro"
    dir_gem = out_root / f"03_gemini_{_slug(gemini_model)}"
    dir_oai = out_root / f"04_openai_{_slug(openai_model)}_{_slug(openai_quality)}"
    for d in (dir_in, dir_fal, dir_gem, dir_oai, meta_dir_for(out_root)):
        d.mkdir(parents=True, exist_ok=True)

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
        raise SystemExit(f"JAN={jan} 参照DL失敗")

    gemini_client = make_client()
    classifications: List[Dict[str, Any]] = []
    for p in ref_paths:
        classifications.append(classify_image_themes(p, client=gemini_client))
    hits = aggregate_theme_hits(classifications)
    jobs = select_themes_for_jobs(hits, max_themes=5, total_jobs=total_jobs)
    if not jobs:
        raise SystemExit(f"JAN={jan}: テーマ抽出0件")

    labels = provider_labels(
        providers, gemini_model=gemini_model, openai_quality=openai_quality
    )
    jobs_meta: List[Dict[str, Any]] = []
    grid_rows: List[Dict[str, Any]] = []

    for job in jobs:
        stem = f"T{job['themeId']:02d}_V{job['variant']:02d}"
        label = f"{stem} {job['themeName']}"[:24]
        urls = list(job.get("sourceUrls") or [])[:max_refs_per_job]
        refs = [path_by_url[u] for u in urls if u in path_by_url]
        if not refs:
            refs = ref_paths[:1]
        prompt = build_ja_prompt(
            product_name=product_name,
            job=job,
            n_refs=len(refs),
            ocr_bundle=_ocr_bundle(job),
        )
        if package_lock:
            prompt = with_package_lock(prompt)

        paths_out: Dict[str, str] = {}
        row_meta: Dict[str, Any] = {
            "stem": stem,
            "themeId": job["themeId"],
            "themeName": job["themeName"],
            "variant": job["variant"],
            "prompt": prompt,
            "sourceUrls": urls,
            "providers": {},
        }

        if "fal" in providers:
            try:
                data, ep = fal_generate(
                    prompt=prompt,
                    image_paths=refs,
                    endpoint=fal_endpoint,
                    allow_txt_fallback=False,
                )
                outp = dir_fal / f"{stem}.jpg"
                _save_bytes_jpeg(data, outp)
                paths_out["fal"] = str(outp)
                row_meta["providers"]["fal"] = {"ok": True, "endpoint": ep, "path": str(outp)}
                LOG.info("fal ok %s", stem)
            except Exception as e:
                row_meta["providers"]["fal"] = {"ok": False, "error": str(e)}
                LOG.warning("fal fail %s: %s", stem, e)

        if "gemini" in providers:
            try:
                data, gmeta = gemini_generate(
                    client=gemini_client,
                    model_id=gemini_model,
                    prompt=prompt,
                    image_paths=refs,
                    aspect_ratio="1:1",
                    image_size="1K",
                )
                outp = dir_gem / f"{stem}.jpg"
                _save_bytes_jpeg(data, outp)
                paths_out["gemini"] = str(outp)
                row_meta["providers"]["gemini"] = {
                    "ok": True,
                    "model": gemini_model,
                    "path": str(outp),
                    "meta": gmeta,
                }
                LOG.info("gemini ok %s model=%s", stem, gemini_model)
            except Exception as e:
                row_meta["providers"]["gemini"] = {"ok": False, "error": str(e)}
                LOG.warning("gemini fail %s: %s", stem, e)

        if "openai" in providers:
            try:
                data, used = openai_generate(
                    prompt=prompt,
                    image_paths=refs,
                    model=openai_model,
                    quality=openai_quality,
                    size="1024x1024",
                )
                outp = dir_oai / f"{stem}.jpg"
                _save_bytes_jpeg(data, outp)
                paths_out["openai"] = str(outp)
                row_meta["providers"]["openai"] = {
                    "ok": True,
                    "model": used,
                    "quality": openai_quality,
                    "path": str(outp),
                }
                LOG.info("openai ok %s model=%s q=%s", stem, used, openai_quality)
            except Exception as e:
                row_meta["providers"]["openai"] = {"ok": False, "error": str(e)}
                LOG.warning("openai fail %s: %s", stem, e)

        jobs_meta.append(row_meta)
        grid_rows.append({"stem": stem, "label": label, "paths": paths_out})

    contact = out_root / "05_contact_providers.jpg"
    title_bits = " / ".join(ja for _, ja in labels)
    build_grid(
        rows=grid_rows,
        providers=labels,
        out_path=contact,
        title=f"比較 {title_bits} — {product_name} ({jan})"
        + (" パッケージ厳禁" if package_lock else ""),
    )

    (out_root / "README.md").write_text(
        "\n".join(
            [
                f"# モデル比較 — {product_name} ({jan})",
                "",
                f"- providers: {', '.join(providers)}",
                f"- fal: `{fal_endpoint}`" if "fal" in providers else "",
                f"- Gemini: `{gemini_model}` 1K" if "gemini" in providers else "",
                f"- OpenAI: `{openai_model}` quality=`{openai_quality}`"
                if "openai" in providers
                else "",
                "- パッケージ色・デザイン変更: **厳禁**" if package_lock else "",
                "- 同一日本語プロンプト・同一テーマジョブ",
                "",
                "## テーマ",
                jobs_summary_ja(jobs),
                "",
                f"contact: `{contact.name}`",
                "",
            ]
        ),
        encoding="utf-8",
    )
    meta = {
        "jan": jan,
        "productName": product_name,
        "providers": list(providers),
        "packageLock": package_lock,
        "models": {
            "fal": fal_endpoint if "fal" in providers else None,
            "gemini": gemini_model if "gemini" in providers else None,
            "openai": (
                {"model": openai_model, "quality": openai_quality}
                if "openai" in providers
                else None
            ),
        },
        "jobs": jobs_meta,
        "outRoot": str(out_root),
        "contact": str(contact),
        "dirs": {
            "fal": str(dir_fal),
            "gemini": str(dir_gem),
            "openai": str(dir_oai),
        },
    }
    (meta_dir_for(out_root) / "run_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(str(out_root))
    return out_root


def resume_openai_only(
    *,
    out_root: Path,
    max_refs_per_job: int = 3,
    force: bool = False,
    openai_model: Optional[str] = None,
    openai_quality: Optional[str] = None,
) -> Path:
    """既存比較ランの OpenAI 列だけ再生成し、contact / meta を更新する。"""
    out_root = Path(out_root)
    meta_path = meta_dir_for(out_root) / "run_meta.json"
    if not meta_path.is_file():
        raise SystemExit(f"run_meta.json がありません: {meta_path}")
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    jobs = list(meta.get("jobs") or [])
    if not jobs:
        raise SystemExit("jobs が空です")

    oai_cfg = (meta.get("models") or {}).get("openai") or {}
    model = openai_model or oai_cfg.get("model") or DEFAULT_OPENAI_MODEL
    quality = openai_quality or oai_cfg.get("quality") or DEFAULT_OPENAI_QUALITY
    dirs = meta.get("dirs") or {}
    dir_in = out_root / "01_refs"
    dir_fal = Path(dirs.get("fal") or (out_root / "02_fal_flux2_pro"))
    dir_gem = Path(dirs.get("gemini") or (out_root / "03_gemini_25_flash"))
    dir_oai = Path(
        dirs.get("openai")
        or (out_root / f"04_openai_{_slug(model)}_{_slug(quality)}")
    )
    dir_oai.mkdir(parents=True, exist_ok=True)

    refs = sorted(
        [p for p in dir_in.iterdir() if p.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp")],
        key=lambda p: p.name,
    )
    if not refs:
        raise SystemExit(f"参照画像がありません: {dir_in}")
    refs_use = refs[: max(1, max_refs_per_job)]
    LOG.info("resume openai-only out=%s jobs=%s q=%s", out_root.name, len(jobs), quality)

    product_name = meta.get("productName") or meta.get("jan") or ""
    jan = meta.get("jan") or ""
    providers = list(meta.get("providers") or ["fal", "gemini", "openai"])
    if "openai" not in providers:
        providers.append("openai")
    gemini_model = (meta.get("models") or {}).get("gemini") or DEFAULT_GEMINI_MODEL
    labels = provider_labels(providers, gemini_model=str(gemini_model), openai_quality=quality)
    grid_rows: List[Dict[str, Any]] = []
    ok_n = 0

    for j in jobs:
        stem = str(j.get("stem") or "")
        if not stem:
            continue
        prompt = str(j.get("prompt") or "").strip()
        if not prompt:
            LOG.warning("skip %s: prompt 無し", stem)
            continue
        label = f"{stem} {j.get('themeName') or ''}"[:24]
        outp = dir_oai / f"{stem}.jpg"
        prov = j.setdefault("providers", {})
        already_ok = bool((prov.get("openai") or {}).get("ok")) and outp.is_file()
        if already_ok and not force:
            LOG.info("openai skip (exists) %s", stem)
            ok_n += 1
        else:
            try:
                data, used = openai_generate(
                    prompt=prompt,
                    image_paths=refs_use,
                    model=model,
                    quality=quality,
                    size="1024x1024",
                )
                _save_bytes_jpeg(data, outp)
                prov["openai"] = {
                    "ok": True,
                    "model": used,
                    "quality": quality,
                    "path": str(outp),
                    "resumed": True,
                }
                ok_n += 1
                LOG.info("openai ok %s model=%s", stem, used)
            except Exception as e:
                prov["openai"] = {"ok": False, "error": str(e), "resumed": True}
                LOG.warning("openai fail %s: %s", stem, e)

        paths_out: Dict[str, str] = {}
        fal_p = dir_fal / f"{stem}.jpg"
        gem_p = dir_gem / f"{stem}.jpg"
        if fal_p.is_file():
            paths_out["fal"] = str(fal_p)
        if gem_p.is_file():
            paths_out["gemini"] = str(gem_p)
        if outp.is_file() and (prov.get("openai") or {}).get("ok"):
            paths_out["openai"] = str(outp)
        grid_rows.append({"stem": stem, "label": label, "paths": paths_out})

    contact = out_root / "05_contact_providers.jpg"
    build_grid(
        rows=grid_rows,
        providers=labels,
        out_path=contact,
        title=f"比較 — {product_name} ({jan})",
    )
    meta["contact"] = str(contact)
    meta["openaiResumeAt"] = datetime.now(timezone.utc).isoformat()
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info("openai resume done ok=%s/%s → %s", ok_n, len(jobs), out_root)
    print(str(out_root))
    return out_root


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="モデル比較 PoC（fal / Gemini / OpenAI）")
    ap.add_argument("--jan", default=None, help="新規実行時のJAN（--resume 時は不要）")
    ap.add_argument("--resume", type=Path, default=None, help="既存比較フォルダを再開")
    ap.add_argument("--only", choices=("openai",), default=None)
    ap.add_argument("--force", action="store_true")
    ap.add_argument(
        "--providers",
        default="gemini,openai",
        help="カンマ区切り: fal,gemini,openai",
    )
    ap.add_argument("--gemini-model", default=DEFAULT_GEMINI_MODEL)
    ap.add_argument("--openai-model", default=DEFAULT_OPENAI_MODEL)
    ap.add_argument("--openai-quality", default=DEFAULT_OPENAI_QUALITY, choices=("low", "medium", "high"))
    ap.add_argument("--fal-endpoint", default=DEFAULT_FAL_ENDPOINT)
    ap.add_argument(
        "--package-lock",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="商品パッケージ改変厳禁（既定ON）",
    )
    ap.add_argument("--max-classify", type=int, default=12)
    ap.add_argument("--max-refs-per-job", type=int, default=3)
    ap.add_argument("--total-jobs", type=int, default=10)
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    if args.resume:
        if args.only != "openai":
            raise SystemExit("--resume 時は --only openai を指定してください")
        out = resume_openai_only(
            out_root=Path(args.resume),
            max_refs_per_job=int(args.max_refs_per_job),
            force=bool(args.force),
            openai_model=str(args.openai_model),
            openai_quality=str(args.openai_quality),
        )
        meta_path = meta_dir_for(out) / "run_meta.json"
        data = json.loads(meta_path.read_text(encoding="utf-8"))
        ok = sum(
            1
            for j in data.get("jobs") or []
            if (j.get("providers") or {}).get("openai", {}).get("ok")
        )
        return 0 if ok else 3

    jan = str(args.jan or "").strip()
    if not jan:
        raise SystemExit("--jan か --resume が必要です")
    providers = [p.strip().lower() for p in str(args.providers).split(",") if p.strip()]
    for p in providers:
        if p not in ("fal", "gemini", "openai"):
            raise SystemExit(f"不明な provider: {p}")
    accepted = read_accepted_from_b3(jans=[jan])
    items = accepted.get(jan) or []
    if not items:
        LOG.error("JAN=%s レ点0件", jan)
        return 2
    LOG.info(
        "JAN=%s providers=%s gemini=%s openai=%s/%s package_lock=%s",
        jan,
        providers,
        args.gemini_model,
        args.openai_model,
        args.openai_quality,
        args.package_lock,
    )
    out = run_compare(
        jan=jan,
        items=items,
        work_root=args.work_root or default_work_root(),
        max_classify=int(args.max_classify),
        max_refs_per_job=int(args.max_refs_per_job),
        total_jobs=int(args.total_jobs),
        providers=providers,
        gemini_model=str(args.gemini_model),
        openai_model=str(args.openai_model),
        openai_quality=str(args.openai_quality),
        fal_endpoint=str(args.fal_endpoint),
        package_lock=bool(args.package_lock),
    )
    ok = 0
    meta_path = meta_dir_for(out) / "run_meta.json"
    if meta_path.is_file():
        data = json.loads(meta_path.read_text(encoding="utf-8"))
        for j in data.get("jobs") or []:
            for p in (j.get("providers") or {}).values():
                if p.get("ok"):
                    ok += 1
    LOG.info("done ok_images=%s → %s", ok, out)
    return 0 if ok else 3


if __name__ == "__main__":
    sys.exit(main())
