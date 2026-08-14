# -*- coding: utf-8 -*-
"""
fal Kontext 競合改変 — 比較テスト

同じ競合1枚に対して:
  mode=4way:
    A 背景のみ（ベージュ）
    B リライトのみ（スタジオ光）
    C 湯気・シズル局所追加のみ
    D 文字色のみ（文言は変えない）
  mode=combo:
    E 背景変更 + 文字色変更 + レイアウト小変更（同一パスで同時適用）
  mode=both: 4way のあと combo も実行

商品本体の色変更はプロンプトで厳禁。
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFont

from b3_comp_catalog import download_url, url_cache_name
from fal_image import generate_with_references
from sub_image_b3_curate import read_accepted_from_b3
from work_paths import default_work_root, meta_dir_for

LOG = logging.getLogger("set_main_image.fal_edit_compare")

PRODUCT_LOCK = (
    "Keep the product bottle/can/package, lid/cap, label print colors, food/contents colors, "
    "shape, pose, scale, crop, and camera angle EXACTLY unchanged. "
    "Do NOT recolor the product. Do NOT invent new packaging artwork. Photoreal."
)

EDIT_CASES_4WAY: List[Dict[str, str]] = [
    {
        "id": "A_bg_beige",
        "ja": "A 背景のみ→ベージュ",
        "prompt": (
            "Replace ONLY the background with a soft sand-beige seamless studio backdrop "
            "(#F5F0E6), clean, no clutter. "
            f"{PRODUCT_LOCK} "
            "Only the background changes."
        ),
    },
    {
        "id": "B_relight",
        "ja": "B リライトのみ→スタジオ光",
        "prompt": (
            "Relight ONLY: soft even studio lighting from front-left, gentle natural shadows. "
            "Do not move objects. Do not change background content except lighting. "
            f"{PRODUCT_LOCK} "
            "Only lighting and subtle shadows change."
        ),
    },
    {
        "id": "C_steam",
        "ja": "C 湯気・シズル局所追加のみ",
        "prompt": (
            "Add ONLY subtle photoreal steam or appetite sizzle near food/product if natural; "
            "keep it minimal. Do not add text, logos, badges, or new panels. "
            f"{PRODUCT_LOCK} "
            "Only a small steam/sizzle accent may be added."
        ),
    },
    {
        "id": "D_text_color",
        "ja": "D 文字色のみ（文言固定）",
        "prompt": (
            "Recolor ONLY non-product text panels / callout headings / body text to dark warm brown "
            "(#3A2A1A). Keep the exact same characters, wording, layout, and positions. "
            "Do not rewrite text. Do not change product label print colors on the package. "
            f"{PRODUCT_LOCK} "
            "Only text ink color on graphic panels changes."
        ),
    },
]

EDIT_CASE_COMBO: Dict[str, str] = {
    "id": "E_combo_bg_text_layout",
    "ja": "E 背景+文字色+レイアウト小変更",
    "prompt": (
        "Apply ALL THREE of the following changes in a SINGLE edit (and nothing else): "
        "(1) BACKGROUND: replace only the backdrop with soft sand-beige seamless studio "
        "(#F5F0E6), clean, no clutter. "
        "(2) TEXT COLOR: recolor ONLY non-product graphic/callout text ink to dark warm brown "
        "(#3A2A1A). Keep the exact same Japanese/characters, wording, and panel count. "
        "Do not rewrite, translate, or invent text. Do not change printed colors on the "
        "physical product label. "
        "(3) LAYOUT (SMALL ONLY): make a subtle layout polish — slightly tighten spacing "
        "between existing callout panels OR nudge one existing text panel a few percent "
        "toward the visual center. Do NOT add/remove panels, badges, or logos. "
        "Do NOT redesign the composition. Do NOT significantly change product pose, scale, "
        "or crop. "
        f"{PRODUCT_LOCK} "
        "If the source has almost no graphic text panels, still do (1) and a tiny (3); "
        "skip inventing text for (2)."
    ),
}


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


def _save_jpeg(data: bytes, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    im = Image.open(BytesIO(data))
    if im.mode != "RGB":
        im = im.convert("RGB")
    im.save(path, format="JPEG", quality=92, optimize=True)


def build_contact(
    *,
    src: Path,
    results: List[Tuple[str, str, Path]],
    out_path: Path,
    title: str,
) -> None:
    """results: (id, ja, path)"""
    cell = 360
    cols = 1 + len(results)
    W = max(cols * cell + 40, 400)
    H = cell + 110
    im = Image.new("RGB", (W, H), (32, 30, 28))
    d = ImageDraw.Draw(im)
    f = _jp_font(15)
    tf = _jp_font(20)
    d.text((16, 10), title[:72], fill=(240, 230, 210), font=tf)
    items = [("SRC", "元・競合", src)] + list(results)
    for i, (cid, ja, p) in enumerate(items):
        x = 20 + i * cell
        thumb = _thumb(p, (cell - 16, cell - 40))
        im.paste(thumb, (x, 48))
        d.text((x, 48 + cell - 36), f"{cid}"[:22], fill=(220, 200, 170), font=f)
        d.text((x, 48 + cell - 18), ja[:20], fill=(180, 170, 150), font=f)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    im.save(out_path, quality=90)


def pick_source(
    items: List[Dict[str, str]],
    *,
    cache_dir: Path,
    index: int,
) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    if not items:
        raise SystemExit("採用レ点画像がありません")
    it = items[min(max(0, index), len(items) - 1)]
    url = it["url"]
    dest = cache_dir / url_cache_name(url)
    download_url(url, dest)
    return dest


def _parse_ref_indices(raw: str) -> List[int]:
    out: List[int] = []
    for part in (raw or "0").split(","):
        part = part.strip()
        if not part:
            continue
        out.append(int(part))
    return out or [0]


def _run_cases(
    *,
    cases: Sequence[Dict[str, str]],
    product: str,
    src: Path,
    dir_out: Path,
    endpoint: Optional[str],
) -> Tuple[List[Tuple[str, str, Path]], List[Dict[str, Any]]]:
    result_paths: List[Tuple[str, str, Path]] = []
    jobs_meta: List[Dict[str, Any]] = []
    for case in cases:
        prompt = f"Product hint: {product}. {case['prompt']}"
        outp = dir_out / f"{case['id']}.jpg"
        try:
            data, ep = generate_with_references(
                prompt=prompt,
                image_paths=[src],
                endpoint=endpoint,
            )
            _save_jpeg(data, outp)
            result_paths.append((case["id"], case["ja"], outp))
            jobs_meta.append(
                {
                    "id": case["id"],
                    "ja": case["ja"],
                    "ok": True,
                    "path": str(outp),
                    "endpoint": ep,
                    "prompt": prompt,
                }
            )
            LOG.info("ok %s → %s", case["id"], outp.name)
        except Exception as e:
            jobs_meta.append(
                {
                    "id": case["id"],
                    "ja": case["ja"],
                    "ok": False,
                    "error": str(e),
                    "prompt": prompt,
                }
            )
            LOG.warning("fail %s: %s", case["id"], e)
    return result_paths, jobs_meta


def run_one(
    *,
    jan: str,
    items: List[Dict[str, str]],
    ref_index: int,
    mode: str,
    work: Path,
    run_id: str,
    endpoint: Optional[str],
) -> Optional[Path]:
    product = (items[0].get("productName") if items else "") or jan
    tag = f"{jan}_r{ref_index}_{run_id}"
    out_root = work / "00.テスト出力" / "sub_image_fal_edit_compare" / tag
    dir_in = out_root / "01_src"
    dir_out = out_root / "02_edits"
    for d in (dir_in, dir_out, meta_dir_for(out_root)):
        d.mkdir(parents=True, exist_ok=True)

    src = pick_source(items, cache_dir=dir_in, index=ref_index)
    src_copy = dir_in / "source.jpg"
    if src.resolve() != src_copy.resolve():
        src_copy.write_bytes(src.read_bytes())

    cases: List[Dict[str, str]] = []
    if mode in ("4way", "both"):
        cases.extend(EDIT_CASES_4WAY)
    if mode in ("combo", "both"):
        cases.append(EDIT_CASE_COMBO)

    result_paths, jobs_meta = _run_cases(
        cases=cases,
        product=product,
        src=src,
        dir_out=dir_out,
        endpoint=endpoint,
    )

    if result_paths:
        if mode == "combo":
            title = "fal競合 3点同時（背景+文字色+レイアウト小）左=元 商品色禁止"
            contact_name = "03_contact_combo.jpg"
        elif mode == "4way":
            title = "fal競合改変 4本比較（左=元画像）商品色変更禁止"
            contact_name = "03_contact_4way.jpg"
        else:
            title = "fal競合 4way+3点同時（左=元）商品色変更禁止"
            contact_name = "03_contact_all.jpg"
        build_contact(
            src=src,
            results=result_paths,
            out_path=out_root / contact_name,
            title=title,
        )

    mode_line = {
        "4way": "- A背景 / Bリライト / C湯気 / D文字色",
        "combo": "- E: 背景変更 + 文字色変更 + レイアウト小変更（同一画像・同時）",
        "both": "- 4way + E 3点同時",
    }.get(mode, mode)

    (out_root / "README.md").write_text(
        "\n".join(
            [
                f"# fal競合改変 — {product} ({jan}) ref={ref_index}",
                "",
                "- 元画像: B-③サブ採用レ点",
                mode_line,
                "- 商品色変更はプロンプトで禁止（保証は人眼確認）",
                "",
                f"元: `{src.name}`",
                "",
            ]
        ),
        encoding="utf-8",
    )
    meta = {
        "jan": jan,
        "productName": product,
        "refIndex": ref_index,
        "mode": mode,
        "source": str(src),
        "jobs": jobs_meta,
        "outRoot": str(out_root),
    }
    (meta_dir_for(out_root) / "run_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(str(out_root))
    return out_root if result_paths else None


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="fal競合改変比較（4way / 3点同時combo）")
    ap.add_argument("--jan", action="append", required=True)
    ap.add_argument(
        "--ref-index",
        default="0",
        help="採用レ点インデックス。カンマ区切りで複数可（例: 0,2）",
    )
    ap.add_argument(
        "--mode",
        choices=("4way", "combo", "both"),
        default="combo",
        help="既定=combo（背景+文字色+レイアウト小を同時）",
    )
    ap.add_argument("--endpoint", default=None, help="既定 flux-kontext/dev")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    jans = [str(j).strip() for j in args.jan if str(j).strip()]
    ref_indices = _parse_ref_indices(str(args.ref_index))
    accepted = read_accepted_from_b3(jans=jans)
    work = args.work_root or default_work_root()
    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")

    ok_n = 0
    for jan in jans:
        items = accepted.get(jan) or []
        if not items:
            LOG.warning("JAN=%s レ点0件スキップ", jan)
            continue
        product = items[0].get("productName") or jan
        LOG.info("JAN=%s adopted=%s product=%s", jan, len(items), product)
        for ri in ref_indices:
            if ri < 0 or ri >= len(items):
                LOG.warning("JAN=%s ref-index=%s 範囲外(0..%s)スキップ", jan, ri, len(items) - 1)
                continue
            out = run_one(
                jan=jan,
                items=items,
                ref_index=ri,
                mode=str(args.mode),
                work=work,
                run_id=run_id,
                endpoint=args.endpoint,
            )
            if out:
                ok_n += 1
    return 0 if ok_n else 3


if __name__ == "__main__":
    sys.exit(main())
