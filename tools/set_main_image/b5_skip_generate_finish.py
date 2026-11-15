# -*- coding: utf-8 -*-
"""
B-⑤「画像の生成不要」完走ヘルパー。

既存のテーマ別新規 / 競合 compose を生成せず、レビュー起動まで進める。
確認フォルダへ出すのは、確認ページで採用したときだけ。

例:
  python b5_skip_generate_finish.py --jan 4906283045119 --path both --export --launch-review
  python b5_skip_generate_finish.py --jan A --jan B --path theme_new --export
  python b5_skip_generate_finish.py --jan A --path competitor --export-only
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from sub_image_review_launch import compose_root, latest_compose_dir
from work_paths import DEFAULT_RAKUTEN_UPLOAD_DIR, DEFAULT_WORK_ROOT

LOG = logging.getLogger("set_main_image.b5_skip_generate_finish")
SCRIPT_DIR = Path(__file__).resolve().parent
HELD_SUB_DIR_NAME = "外した画像"
_SUB_MARK = re.compile(r"_(?:sub|p)(\d+)", re.I)
_IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp"}


def _theme_new_dir(jan: str) -> Path:
    return SCRIPT_DIR / "_out" / f"theme_new_run_{jan}"


def count_theme_new_images(theme_new_dir: Path) -> int:
    root = Path(theme_new_dir)
    if not root.is_dir():
        return 0
    return sum(1 for p in root.glob("T*.jpg") if p.is_file())


def retire_product_subs(out_dir: Path, product_key: str) -> List[str]:
    """確認フォルダ直下の、その品番のサブ画像を「外した画像」へ移す。

    メイン（*_rakuten.jpg）と、サブフォルダの中は動かさない。
    紐付けは直下の画像だけ読む。
    """
    key = str(product_key or "").strip()
    root = Path(out_dir)
    if not key or not root.is_dir():
        LOG.info("retire sub skip key=%s dir=%s", key, root)
        return []
    prefix = f"{key}_"
    hold = root / HELD_SUB_DIR_NAME
    moved: List[str] = []
    for src in root.iterdir():
        if not src.is_file():
            continue
        name = src.name
        if not name.startswith(prefix):
            continue
        if name.lower().endswith("_rakuten.jpg"):
            continue
        if src.suffix.lower() not in _IMAGE_EXT:
            continue
        if not _SUB_MARK.search(name):
            continue
        hold.mkdir(parents=True, exist_ok=True)
        dest = hold / name
        if dest.exists():
            stamp = datetime.now().strftime("%H%M%S")
            dest = hold / f"{src.stem}__moved{stamp}{src.suffix}"
        shutil.move(str(src), str(dest))
        moved.append(name)
        LOG.info("retire sub key=%s file=%s held=%s", key, name, dest.name)
    LOG.info("retire sub key=%s moved=%s dir=%s", key, len(moved), root)
    return moved


def child_skus_for_image_bind(jan: str, out_dir: Optional[Path] = None) -> List[str]:
    """確認フォルダのサブ名に入れる子SKU。

    出品CKが付いた子を先に使う。
    付いていなければ、確認フォルダに既にあるメイン画像と同じ子SKUだけを使う。
    同じ商品番号の別の親には広げない。商品番号そのものは返さない。
    """
    from master_sets import (
        CHILD_HINTS,
        PARENT_HINTS,
        _col,
        _find_header_row,
        _norm,
        parents_for_jan_from_rows,
        resolve_checked_children_for_jan,
    )
    from sheets_master import fetch_master_rows

    rows, _info = fetch_master_rows()
    checked = resolve_checked_children_for_jan(jan, rows=rows)
    if checked:
        skus = [c.child_sku for c in checked if c.child_sku]
        LOG.info("bind skus jan=%s via=checked n=%s", jan, len(skus))
        return skus
    # 同じ商品番号の別親（単品とミックス）へは広げない。
    # 確認フォルダに既にあるメイン画像の子SKUだけを使う。
    parents = {_norm(p) for p in parents_for_jan_from_rows(rows, jan)}
    header_i, idx = _find_header_row(rows)
    i_parent = _col(idx, PARENT_HINTS)
    i_child = _col(idx, CHILD_HINTS)
    skus: List[str] = []
    seen = set()
    if i_parent is None or i_child is None:
        LOG.info("bind skus jan=%s via=none reason=no_columns", jan)
        return []
    for row in rows[header_i + 1 :]:
        if max(i_parent, i_child) >= len(row):
            continue
        parent = _norm(row[i_parent])
        child = _norm(row[i_child])
        if parent not in parents or not child or child in seen:
            continue
        seen.add(child)
        skus.append(child)
    if out_dir is not None:
        mains = {
            p.name[: -len("_rakuten.jpg")]
            for p in Path(out_dir).glob("*_rakuten.jpg")
            if p.is_file()
        }
        skus = [s for s in skus if s in mains]
    LOG.info("bind skus jan=%s via=main_files n=%s", jan, len(skus))
    return skus


def parent_sku_from_path_hint(hint: str) -> str:
    """正本パス／ファイル名から親SKU（ishi-…-oya3_ミックス.png → ishi-…-oya3）。"""
    stem = Path(str(hint or "").replace("\\", "/")).name
    stem = Path(stem).stem
    for suf in ("_単体", "_ミックス", "_mix", "_single"):
        if stem.endswith(suf):
            return stem[: -len(suf)].strip()
    s = stem.strip()
    if s.startswith("ishi-") or "-oya" in s:
        return s
    return ""


def _main_skus_in_dir(out_dir: Optional[Path]) -> set:
    if out_dir is None or not Path(out_dir).is_dir():
        return set()
    return {
        p.name[: -len("_rakuten.jpg")]
        for p in Path(out_dir).glob("*_rakuten.jpg")
        if p.is_file()
    }


def _pick_one_child_sku(kids: Sequence[Any], out_dir: Optional[Path] = None) -> str:
    """親1セット出口用。MAINがある子→総個数最小→SKU名。"""
    mains = _main_skus_in_dir(out_dir)
    rows = [k for k in kids if getattr(k, "child_sku", None)]
    if not rows:
        return ""
    with_main = [k for k in rows if k.child_sku in mains]
    pool = with_main or list(rows)
    pool.sort(key=lambda k: (int(getattr(k, "set_count", 0) or 10**9), str(k.child_sku)))
    return str(pool[0].child_sku)


def represent_child_skus_for_layout_export(
    jan: str = "",
    *,
    parent_sku: str = "",
    out_dir: Optional[Path] = None,
) -> List[str]:
    """
    Cレイアウト（サブ＝親枠1セット→L2で全子同一）向けの出口名。
    親ごとに子SKUを1件だけ返す。全セット子へ複製しない。
    """
    from master_sets import (
        load_checked_set_children_from_rows,
        load_set_children_for_parent,
        resolve_checked_children_for_jan,
    )
    from sheets_master import fetch_master_rows
    from work_paths import DEFAULT_MASTER_CSV

    want_parent = str(parent_sku or "").strip()
    rows, _info = fetch_master_rows()
    picked: List[str] = []
    seen_p = set()

    if want_parent:
        kids, _ = load_checked_set_children_from_rows(rows, parent_sku=want_parent)
        if not kids:
            try:
                kids, _ = load_set_children_for_parent(
                    Path(DEFAULT_MASTER_CSV), want_parent, checked_only=False
                )
            except Exception as e:
                LOG.warning("layout export fallback master_csv failed: %s", e)
                kids = []
        one = _pick_one_child_sku(kids, out_dir)
        if one:
            picked.append(one)
        LOG.info(
            "layout export sku parent=%s via=parent n=%s sku=%s",
            want_parent,
            len(picked),
            picked,
        )
        return picked

    checked = resolve_checked_children_for_jan(str(jan or "").strip(), rows=rows)
    if not checked:
        # fallback: bind list then 1名（親不明時は先頭1件のみ・増殖禁止）
        skus = child_skus_for_image_bind(str(jan or "").strip(), out_dir)
        if skus:
            picked = [skus[0]]
        LOG.info("layout export sku jan=%s via=bind_fallback n=%s sku=%s", jan, len(picked), picked)
        return picked

    by_parent: Dict[str, List[Any]] = {}
    for k in checked:
        by_parent.setdefault(str(k.parent_sku), []).append(k)
    for p, kids in by_parent.items():
        if p in seen_p:
            continue
        seen_p.add(p)
        one = _pick_one_child_sku(kids, out_dir)
        if one:
            picked.append(one)
    LOG.info("layout export sku jan=%s via=checked_one_per_parent n=%s sku=%s", jan, len(picked), picked)
    return picked


def rewrite_jan_subs_to_child_sku(jan: str, out_dir: Path) -> Dict[str, Any]:
    """直下の「商品番号_…_subN」を、親あたり1子SKU名へ複製してから外す（全子増殖しない）。"""
    key = str(jan or "").strip()
    root = Path(out_dir)
    skus = represent_child_skus_for_layout_export(key, out_dir=root)
    if not skus:
        raise RuntimeError(f"子SKUがありません: {key}")
    # 旧・全子コピー分も枠外の元になるので、bind対象だった子名のサブは先に外す
    all_bind = child_skus_for_image_bind(key, root)
    pat = re.compile(
        rf"^{re.escape(key)}_(.+)_sub(\d+)\.(jpg|jpeg|png|webp)$",
        re.I,
    )
    sources = []
    for src in root.iterdir():
        if not src.is_file():
            continue
        m = pat.match(src.name)
        if not m:
            continue
        sources.append((int(m.group(2)), m.group(1), src))
    sources.sort()
    if not sources:
        LOG.info("rewrite subs jan=%s sources=0", key)
        return {"jan": key, "skus": skus, "copied": 0, "retired": 0}
    retired: List[str] = []
    for sku in all_bind:
        retired.extend(retire_product_subs(root, sku))
    copied = 0
    for sku in skus:
        for n, slug, src in sources:
            dest = root / f"{sku}_{slug}_sub{n}{src.suffix.lower()}"
            shutil.copy2(src, dest)
            copied += 1
            LOG.info("rewrite sub %s -> %s", src.name, dest.name)
    retired.extend(retire_product_subs(root, key))
    LOG.info(
        "rewrite subs jan=%s skus=%s copied=%s retired=%s",
        key,
        len(skus),
        copied,
        len(retired),
    )
    return {
        "jan": key,
        "skus": skus,
        "copied": copied,
        "retired": len(retired),
    }


def export_theme_new_to_rakuten(
    *,
    jan: str,
    theme_new_dir: Path,
    out_dir: Path,
    max_subs: int = 10,
) -> Dict[str, Any]:
    """作成直後の全コピーはしない。確認フォルダへ出すのは採用時だけ。"""
    LOG.info(
        "theme_new export deferred jan=%s reason=adopt_only images=%s out=%s",
        jan,
        count_theme_new_images(Path(theme_new_dir)),
        out_dir,
    )
    return {
        "ok": True,
        "exported": 0,
        "deferred": True,
        "reason": "adopt_only",
        "outDir": str(out_dir),
        "maxSubs": max_subs,
    }


def export_compose_only(
    *,
    jan: str,
    compose_dir: Path,
    out_dir: Path,
    export_pick: str = "a",
) -> Dict[str, Any]:
    from sub_image_ai_compose_poc import _auto_export_compose

    info = _auto_export_compose(
        Path(compose_dir),
        jan=jan,
        auto_export=True,
        export_out_dir=Path(out_dir),
        export_pick=export_pick or "a",
        export_provider="openai",
        master_csv=None,
        child_sku="",
    )
    return info or {"ok": False, "error": "auto-export returned None"}


def launch_mix_review(jans: Sequence[str], *, no_wait: bool = True) -> int:
    cmd = [
        sys.executable,
        str(SCRIPT_DIR / "sub_image_mix_review_launch.py"),
    ]
    for j in jans:
        cmd.extend(["--jan", j])
    if no_wait:
        cmd.append("--no-wait")
    LOG.info("launch mix review: %s", " ".join(cmd))
    return subprocess.call(cmd, cwd=str(SCRIPT_DIR))


def launch_compose_review(jans: Sequence[str], *, no_wait: bool = True) -> int:
    cmd = [
        sys.executable,
        str(SCRIPT_DIR / "sub_image_review_launch.py"),
    ]
    for j in jans:
        cmd.extend(["--jan", j])
    if no_wait:
        cmd.append("--no-wait")
    LOG.info("launch compose review: %s", " ".join(cmd))
    return subprocess.call(cmd, cwd=str(SCRIPT_DIR))


def process_jan(
    jan: str,
    *,
    path: str,
    do_export: bool,
    out_dir: Path,
    work_root: Path,
    export_pick: str,
) -> Dict[str, Any]:
    jan = str(jan).strip()
    result: Dict[str, Any] = {"jan": jan, "path": path, "ok": True, "steps": []}
    root = compose_root(work_root)

    if path in ("theme_new", "both"):
        tn = _theme_new_dir(jan)
        n = count_theme_new_images(tn)
        step: Dict[str, Any] = {"kind": "theme_new", "dir": str(tn), "imageN": n}
        if n <= 0:
            step["ok"] = False
            step["error"] = "theme_new 画像がありません（T*.jpg）"
            result["ok"] = False
        elif do_export:
            exp = export_theme_new_to_rakuten(jan=jan, theme_new_dir=tn, out_dir=out_dir)
            step["export"] = exp
            step["ok"] = bool(exp.get("ok"))
            if not step["ok"]:
                result["ok"] = False
        else:
            step["ok"] = True
            step["export"] = {"skipped": True}
        result["steps"].append(step)

    if path in ("competitor", "both"):
        cp = latest_compose_dir(root, jan) if root.is_dir() else None
        step = {"kind": "competitor", "dir": str(cp) if cp else "", "imageN": 0}
        if not cp or not cp.is_dir():
            step["ok"] = False
            step["error"] = "最新 compose-dir がありません"
            # both のとき競合が無くても theme_new だけで続行可
            if path == "competitor":
                result["ok"] = False
            else:
                step["warn"] = "competitor skip"
                step["ok"] = True
        elif do_export:
            exp = export_compose_only(
                jan=jan, compose_dir=cp, out_dir=out_dir, export_pick=export_pick
            )
            step["export"] = exp
            step["ok"] = bool(exp.get("ok"))
            if not step["ok"] and path == "competitor":
                result["ok"] = False
        else:
            step["ok"] = True
            step["export"] = {"skipped": True}
        result["steps"].append(step)

    return result


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="B-⑤ 画像生成スキップ完走（export＋レビュー）")
    ap.add_argument("--jan", action="append", default=[], help="対象JAN（複数可）")
    ap.add_argument(
        "--path",
        choices=("theme_new", "competitor", "both"),
        default="both",
        help="対象経路（既定 both）",
    )
    ap.add_argument(
        "--export",
        dest="do_export",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="楽天目視フォルダへ export（既定ON）",
    )
    ap.add_argument(
        "--export-only",
        action="store_true",
        help="exportのみ（レビュー起動しない）。--export と同義＋--no-launch-review",
    )
    ap.add_argument(
        "--launch-review",
        dest="launch_review",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="export後にレビューUI起動（既定ON）",
    )
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=Path(DEFAULT_RAKUTEN_UPLOAD_DIR),
        help="楽天目視フォルダ",
    )
    ap.add_argument("--work-root", type=Path, default=Path(DEFAULT_WORK_ROOT))
    ap.add_argument("--export-pick", default="a", help="競合 export の A/B（既定 a）")
    ap.add_argument("--no-wait", action="store_true", default=True)
    args = ap.parse_args()

    jans = [str(j).strip() for j in (args.jan or []) if str(j).strip()]
    if not jans:
        raise SystemExit("--jan が必要です")

    do_export = bool(args.do_export)
    launch = bool(args.launch_review)
    if args.export_only:
        do_export = True
        launch = False

    reports: List[Dict[str, Any]] = []
    all_ok = True
    for jan in jans:
        r = process_jan(
            jan,
            path=str(args.path),
            do_export=do_export,
            out_dir=Path(args.out_dir),
            work_root=Path(args.work_root),
            export_pick=str(args.export_pick or "a"),
        )
        reports.append(r)
        if not r.get("ok"):
            all_ok = False
        LOG.info("JAN %s ok=%s steps=%s", jan, r.get("ok"), json.dumps(r.get("steps"), ensure_ascii=False))

    if launch and all_ok:
        if args.path in ("theme_new", "both"):
            launch_mix_review(jans, no_wait=True)
        if args.path in ("competitor", "both"):
            # both のときは mix で競合も見られるが、競合専用 review も起動可
            if args.path == "competitor":
                launch_compose_review(jans, no_wait=True)

    summary = {
        "ok": all_ok,
        "path": args.path,
        "export": do_export,
        "launchReview": launch,
        "outDir": str(args.out_dir),
        "reports": reports,
        "at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }
    out_sum = SCRIPT_DIR / "_out" / f"b5_skip_generate_{time.strftime('%Y%m%d_%H%M%S')}.json"
    out_sum.parent.mkdir(parents=True, exist_ok=True)
    out_sum.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"SUMMARY={out_sum}")
    print(f"HUMAN_REVIEW_DIR={args.out_dir}")
    return 0 if all_ok else 2


if __name__ == "__main__":
    sys.exit(main())
