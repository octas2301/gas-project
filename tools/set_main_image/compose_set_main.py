# -*- coding: utf-8 -*-
"""
セットMAIN画像 — CLI

楽天本線: 02ベース + Canva unitset/digit/pair（layout_rules ①②③固定）
マスタ: 出品CKレ点 × 総個数 × バリエーション単位（無ければ個）

  python compose_set_main.py --list-fonts
  python compose_set_main.py --master-csv … --checked-only --malls rakuten
  python compose_set_main.py --parent-sku … --master-csv … --checked-only --malls rakuten
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional

from PIL import Image

from amazon_compose import choose_preset, compose_amazon
from master_sets import load_checked_set_children, load_set_children_for_parent
from rakuten_badge import compose_rakuten_badge, list_font_ids, load_font_catalog
from rakuten_layer import compose_rakuten_layered
from work_paths import (
    default_work_root,
    move_rakuten_base_to_processed,
    resolve_amazon_base,
    resolve_digit_layer,
    resolve_octas,
    resolve_rakuten_base,
)

LOG = logging.getLogger("set_main_image")


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def default_out_dir(work_root: Optional[Path] = None) -> Path:
    root = work_root or default_work_root()
    # 作業ルート内 00.テスト出力 → なければ Drive 07
    cand = root / "00.テスト出力"
    if cand.is_dir() or root.is_dir():
        cand.mkdir(parents=True, exist_ok=True)
        return cand
    drive = Path(r"G:/マイドライブ")
    if drive.is_dir():
        for p in drive.rglob("*"):
            if p.is_dir() and p.name.startswith("07."):
                return p
    out = Path.cwd() / "_set_main_out"
    out.mkdir(parents=True, exist_ok=True)
    return out


def run_batch(
    *,
    parent_sku: str,
    master_csv: Path,
    work_root: Path,
    out_dir: Path,
    malls: List[str],
    octas_path: Optional[Path],
    font_id: str,
    amazon_base: Optional[Path],
    rakuten_base: Optional[Path],
    digit_layer: Optional[Path],
    checked_only: bool,
    rakuten_engine: str,
    text_color: str = "1",
) -> dict:
    if checked_only:
        children, _food_map = load_checked_set_children(
            master_csv, parent_sku=parent_sku or ""
        )
        is_food = any(ch.is_food for ch in children)
    else:
        if not parent_sku:
            raise SystemExit("--parent-sku が必要です（--checked-only なしのとき）")
        children, is_food = load_set_children_for_parent(
            master_csv, parent_sku, checked_only=False
        )

    if not children:
        raise SystemExit(
            "対象行がありません"
            + ("（出品CKレ点を確認）" if checked_only else f": parent={parent_sku}")
        )

    octas_file = resolve_octas(work_root, octas_path)
    octas = None
    if is_food and "amazon" in malls:
        if not octas_file:
            raise SystemExit(
                "食品のため 99.octas期限管理シール素材 に画像を置くか --octas を指定してください"
            )
        octas = Image.open(octas_file)

    out_dir.mkdir(parents=True, exist_ok=True)
    report = {
        "runAt": datetime.now(timezone.utc).isoformat(),
        "parentSku": parent_sku or "(checked-all)",
        "checkedOnly": checked_only,
        "isFood": is_food,
        "malls": malls,
        "fontId": font_id,
        "rakutenEngine": rakuten_engine,
        "textColor": text_color,
        "workRoot": str(work_root),
        "outputs": [],
        "skipped": [],
    }

    amazon_base_im = None
    if "amazon" in malls:
        # 親ごとにベースが違うのでループ内で解決
        report["amazonBase"] = "(per-parent)"

    rakuten_digit_im = None
    if "rakuten" in malls or "yahoo" in malls:
        dp = resolve_digit_layer(work_root, parent_sku or None, digit_layer)
        if dp:
            rakuten_digit_im = Image.open(dp)
            report["digitLayer"] = str(dp)
        else:
            report["digitLayer"] = None

    # 親SKUごとにベースをキャッシュ
    rakuten_bases: dict = {}
    amazon_bases: dict = {}

    for ch in children:
        if ch.set_count < 1:
            report["skipped"].append(
                {"childSku": ch.child_sku, "n": ch.set_count, "reason": "N<1"}
            )
            continue
        # N=1（1個セット）もレ点対象なら生成（楽天 layered／badge とも可）

        for mall in malls:
            if mall == "yahoo":
                eng, name_mall = "rakuten", "yahoo"
            elif mall == "rakuten":
                eng, name_mall = "rakuten", "rakuten"
            else:
                eng, name_mall = "amazon", "amazon"

            if eng == "amazon":
                if ch.parent_sku not in amazon_bases:
                    ap = resolve_amazon_base(work_root, ch.parent_sku, amazon_base)
                    amazon_bases[ch.parent_sku] = Image.open(ap)
                    LOG.info("amazon base parent=%s %s", ch.parent_sku, ap)
                base_im = amazon_bases[ch.parent_sku]
                preset = choose_preset(ch.set_count)
                img = compose_amazon(
                    base_im,
                    ch.set_count,
                    octas=octas,
                    require_octas=ch.is_food,
                    preset=preset,
                )
                mode = preset
                meta_extra = {}
            else:
                if ch.parent_sku not in rakuten_bases:
                    bp = resolve_rakuten_base(work_root, ch.parent_sku, rakuten_base)
                    rakuten_bases[ch.parent_sku] = bp
                    LOG.info("rakuten base parent=%s %s", ch.parent_sku, bp)
                bp = rakuten_bases[ch.parent_sku]
                if rakuten_engine == "layered":
                    img, layer_meta = compose_rakuten_layered(
                        base_path=bp,
                        set_count=ch.set_count,
                        unit=ch.unit,
                        font_id=font_id,
                        digit_layer_path=None,
                        work_root=work_root,
                        render_mode="glyph_alpha",
                        text_color=text_color,
                    )
                    mode = f"layered:{layer_meta.get('mode')}"
                    tc = layer_meta.get("textColor") or {}
                    meta_extra = {
                        "unit": layer_meta.get("unit"),
                        "unitResolved": (layer_meta.get("typographyLayout") or {}).get(
                            "unitResolved"
                        ),
                        "numSource": (layer_meta.get("typographyLayout") or {}).get(
                            "numSource"
                        ),
                        "textColorId": tc.get("id"),
                        "textColorNameJa": tc.get("nameJa"),
                        "textColorHex": tc.get("hex"),
                    }
                else:
                    img = compose_rakuten_badge(
                        Image.open(bp),
                        ch.set_count,
                        digit_layer=rakuten_digit_im,
                        font_id=font_id,
                        canvas_size=1200,
                    )
                    mode = f"font:{font_id}"
                    meta_extra = {}

            out_name = f"{ch.child_sku}_{name_mall}.jpg"
            out_path = out_dir / out_name
            img.convert("RGB").save(out_path, format="JPEG", quality=85, optimize=True)
            entry = {
                "parentSku": ch.parent_sku,
                "childSku": ch.child_sku,
                "n": ch.set_count,
                "unit": ch.unit,
                "mall": name_mall,
                "path": str(out_path),
                "mode": mode,
            }
            entry.update(meta_extra)
            report["outputs"].append(entry)
            LOG.info(
                "wrote %s n=%s unit=%s mode=%s",
                out_path.name,
                ch.set_count,
                ch.unit,
                mode,
            )

    # 楽天／Yahoo 量産が1件以上成功したベースだけ 02/処理済み へ移動
    moved_bases = []
    if report["outputs"]:
        used_paths = set()
        for ent in report["outputs"]:
            if ent.get("mall") in ("rakuten", "yahoo"):
                psku = ent.get("parentSku")
                bp = rakuten_bases.get(psku)
                if bp:
                    used_paths.add(Path(bp))
        for bp in sorted(used_paths, key=lambda p: str(p)):
            dest = move_rakuten_base_to_processed(bp, work_root)
            if dest:
                moved_bases.append({"from": str(bp), "to": str(dest)})
    report["movedRakutenBases"] = moved_bases

    rep_path = (
        out_dir
        / f"SET_MAIN_{datetime.now().strftime('%Y%m%d_%H%M%S')}_REPORT.json"
    )
    rep_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info(
        "report %s outputs=%s skipped=%s movedBases=%s",
        rep_path,
        len(report["outputs"]),
        len(report["skipped"]),
        len(moved_bases),
    )
    return report


def smoke_amazon(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    from PIL import ImageDraw

    base = Image.new("RGBA", (800, 800), (0, 0, 0, 0))
    d = ImageDraw.Draw(base)
    d.rounded_rectangle((150, 100, 650, 700), radius=40, fill=(230, 100, 40, 255))
    octas = Image.new("RGBA", (120, 80), (255, 255, 255, 255))
    od = ImageDraw.Draw(octas)
    od.rectangle((2, 2, 117, 77), outline=(255, 100, 0), width=3)
    od.text((20, 28), "Octas", fill=(0, 0, 0, 255))
    for n in (2, 3, 10, 30):
        img = compose_amazon(base, n, octas=octas, require_octas=True)
        p = out_dir / f"_smoke_amazon_set{n}.jpg"
        img.save(p, quality=85)
        LOG.info("smoke %s", p)


def smoke_rakuten(
    out_dir: Path, font_id: str, work_root: Path, text_color: str = "1"
) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    bp = resolve_rakuten_base(work_root, "", None)
    for n, unit in ((1, "袋"), (10, "缶"), (21, "個")):
        img, meta = compose_rakuten_layered(
            base_path=bp,
            set_count=n,
            unit=unit,
            font_id=font_id,
            work_root=work_root,
            text_color=text_color,
        )
        p = out_dir / f"_smoke_layered_{unit}_set{n}.jpg"
        img.convert("RGB").save(p, quality=85)
        LOG.info("smoke %s meta=%s color=%s", p, meta.get("mode"), meta.get("textColor"))


def main(argv: Optional[List[str]] = None) -> int:
    cat = load_font_catalog()
    default_font = cat.get("default_font_id") or "helvetica_bold"

    ap = argparse.ArgumentParser(description="セットMAIN画像（楽天 layered / Amazon）")
    ap.add_argument("--parent-sku", default="", help="省略可（--checked-only時は全親）")
    ap.add_argument("--master-csv", type=Path)
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--out-dir", type=Path)
    ap.add_argument("--octas", type=Path)
    ap.add_argument("--amazon-base", type=Path)
    ap.add_argument("--rakuten-base", type=Path)
    ap.add_argument("--digit-layer", type=Path)
    ap.add_argument("--font-id", default=default_font)
    ap.add_argument("--list-fonts", action="store_true")
    ap.add_argument("--malls", default="rakuten", help="既定: rakuten（amazon,rakuten可）")
    ap.add_argument(
        "--checked-only",
        action="store_true",
        help="マスタ出品CKレ点の行だけ生成（親のみレ点→全子）",
    )
    ap.add_argument(
        "--rakuten-engine",
        choices=("layered", "badge"),
        default="layered",
        help="layered=Canva unitset本線（既定） / badge=旧数字のみ",
    )
    ap.add_argument(
        "--text-color",
        default="1",
        help="文字色: 1|2|3|4|5 または えんじ|青|黒|緑|茶（既定=1えんじ）",
    )
    ap.add_argument("--smoke-amazon", action="store_true")
    ap.add_argument("--smoke-rakuten", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    ap.add_argument("--base-dir", type=Path, help=argparse.SUPPRESS)
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    if args.list_fonts:
        print(list_font_ids())
        return 0

    work_root = args.work_root or args.base_dir or default_work_root()
    out_dir = args.out_dir or default_out_dir(work_root)

    if args.smoke_amazon:
        smoke_amazon(out_dir)
        return 0
    if args.smoke_rakuten:
        smoke_rakuten(out_dir, args.font_id, work_root, text_color=args.text_color)
        return 0

    if not args.master_csv:
        ap.error("--master-csv が必要です")

    malls = [m.strip().lower() for m in args.malls.split(",") if m.strip()]
    for m in malls:
        if m not in ("amazon", "rakuten", "yahoo"):
            ap.error(f"unknown mall: {m}")

    run_batch(
        parent_sku=args.parent_sku,
        master_csv=args.master_csv,
        work_root=work_root,
        out_dir=out_dir,
        malls=malls,
        octas_path=args.octas,
        font_id=args.font_id,
        amazon_base=args.amazon_base,
        rakuten_base=args.rakuten_base,
        digit_layer=args.digit_layer,
        checked_only=args.checked_only,
        rakuten_engine=args.rakuten_engine,
        text_color=args.text_color,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
