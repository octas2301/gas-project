# -*- coding: utf-8 -*-
"""AI サブ出力を楽天マトリクス用ファイルへコピーする。

目視（auto-export 既定）:
  品番キー（JAN→親SKU→代表1子→compose名）あたり最大10枚。全セット子への複製はしない。
  `{品番キー}_{pattern}_subN.jpg`（既定 pick=ab）

本番投入（明示 `--to-checked-children`）:
  出品CK付き子SKUへだけ複製。出品CK0件時の全セット子フォールバックは禁止。
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from rakuten_image_names import sub_filename
from work_paths import DEFAULT_MASTER_CSV, DEFAULT_RAKUTEN_UPLOAD_DIR

LOG = logging.getLogger("export_sub_rakuten")

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp"}
PROVIDER_PREFERENCE = ("openai", "gemini", "fal")
REVIEW_MAX_SUBS = 10


def _slot_key(path: Path) -> Tuple[int, int, str]:
    name = path.name
    m = re.search(r"S(\d{2})", name, flags=re.I)
    slot = int(m.group(1)) if m else 999
    m2 = re.search(r"_P(\d+)", name, flags=re.I) or re.search(
        r"phase(?:Order)?[=_]?(\d+)", name, flags=re.I
    )
    phase = int(m2.group(1)) if m2 else slot
    return (phase, slot, name.lower())


def list_source_images(src_dir: Path) -> List[Path]:
    files = [
        p
        for p in src_dir.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_EXT and not p.name.startswith(".")
    ]
    files.sort(key=_slot_key)
    return files


def _load_run_meta(compose_dir: Path) -> Dict[str, Any]:
    compose_dir = Path(compose_dir)
    candidates = [
        compose_dir / "_meta" / "run_meta.json",
        compose_dir / "run_meta.json",
    ]
    for p in candidates:
        if p.is_file():
            return json.loads(p.read_text(encoding="utf-8"))
    raise SystemExit(f"run_meta.json not found under {compose_dir}")


def _load_pick_manifest(path: Optional[Path]) -> Dict[int, str]:
    if not path:
        return {}
    path = Path(path)
    if not path.is_file():
        raise SystemExit(f"pick manifest not found: {path}")
    raw = json.loads(path.read_text(encoding="utf-8"))
    out: Dict[int, str] = {}
    if isinstance(raw, dict):
        # {"1":"a","2":"b"} or {"slots":[{"slotIndex":1,"pick":"a"}]}
        if "slots" in raw and isinstance(raw["slots"], list):
            for s in raw["slots"]:
                si = int(s.get("slotIndex") or s.get("slot") or 0)
                pk = str(s.get("pick") or s.get("proposal") or "a").lower()
                if si >= 1:
                    out[si] = "b" if pk in ("b", "2", "abb") else "a"
        else:
            for k, v in raw.items():
                if str(k).isdigit():
                    pk = str(v).lower()
                    out[int(k)] = "b" if pk in ("b", "2", "abb") else "a"
    return out


def _proposal_key(job: Dict[str, Any]) -> str:
    prop = str(job.get("proposal") or "").lower()
    if prop in ("b", "2", "abb"):
        return "b"
    if prop in ("a", "1", "aba"):
        return "a"
    var = int(job.get("variant") or 0)
    return "b" if var == 2 else "a"


def _job_image_path(job: Dict[str, Any], provider: str) -> Optional[Path]:
    order = [provider] + [p for p in PROVIDER_PREFERENCE if p != provider]
    for p in order:
        block = job.get(p) or {}
        if not isinstance(block, dict):
            continue
        if not block.get("ok"):
            continue
        path = Path(str(block.get("path") or ""))
        if path.is_file():
            return path
    # stem fallback under provider dirs
    stem = str(job.get("stem") or "")
    if not stem:
        return None
    root = Path(str(job.get("outRoot") or ""))
    if not root.is_dir():
        # try parent of meta
        return None
    for p in order:
        folder = {
            "openai": "03_openai",
            "gemini": "02_gemini",
            "fal": "04_fal",
        }.get(p, p)
        cand = root / folder / f"{stem}.jpg"
        if cand.is_file():
            return cand
    return None


def select_jobs_from_meta(
    meta: Dict[str, Any],
    *,
    pick_default: str,
    pick_by_slot: Dict[int, str],
    provider: str,
) -> List[Tuple[int, Dict[str, Any], Path]]:
    """Return list of (slotIndex, job, path) sorted by phaseOrder then slotIndex.

    pick_default:
      - a / b: 各スロット1枚
      - ab / both: 各スロットの A→B（目視最大10枚向け）
    """
    jobs = list(meta.get("jobs") or [])
    by_slot: Dict[int, Dict[str, Dict[str, Any]]] = {}
    for j in jobs:
        si = int(j.get("slotIndex") or 0)
        if si < 1:
            continue
        by_slot.setdefault(si, {})[_proposal_key(j)] = j

    def sort_key(si: int) -> Tuple[int, int]:
        any_j = next(iter(by_slot[si].values()))
        return (int(any_j.get("phaseOrder") or 99), si)

    mode = str(pick_default or "a").strip().lower()
    both = mode in ("ab", "both", "a+b", "a,b")
    selected: List[Tuple[int, Dict[str, Any], Path]] = []
    for si in sorted(by_slot.keys(), key=sort_key):
        if both:
            wants = ["a", "b"]
        else:
            want = pick_by_slot.get(si, mode)
            want = "b" if str(want).lower() in ("b", "2", "abb") else "a"
            wants = [want]
        for want in wants:
            job = by_slot[si].get(want)
            if not job and not both:
                job = by_slot[si].get("a") or by_slot[si].get("b")
            if not job:
                if not both:
                    LOG.warning("slot S%02d: no job", si)
                continue
            path = _job_image_path(job, provider)
            if not path:
                LOG.warning("slot S%02d: no image file (pick=%s)", si, want)
                continue
            selected.append((si, job, path))
    return selected


def _normalize_pick_mode(pick: str) -> str:
    p = str(pick or "ab").strip().lower()
    if p in ("ab", "both", "a+b", "a,b"):
        return "ab"
    if p in ("b", "2", "abb"):
        return "b"
    return "a"


def _load_master_rows(master_csv: Optional[Path] = None) -> List[List[str]]:
    try:
        from sheets_master import fetch_master_rows

        rows, _info = fetch_master_rows()
        if rows:
            return rows
    except Exception as e:
        LOG.warning("sheets master 取得失敗 → CSVへ: %s", e)
    path = Path(master_csv) if master_csv else DEFAULT_MASTER_CSV
    if not path.is_file():
        raise ValueError(f"マスタ行を取得できません（Sheets失敗・CSV無し: {path}）")
    from master_sets import _read_table

    rows, _h, _i = _read_table(path)
    return rows


def resolve_review_product_key(
    *,
    jan: str = "",
    child_sku: str = "",
    product_key: str = "",
    master_csv: Optional[Path] = None,
    compose_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    目視用の品番キー1件（全子展開しない）。
    優先: product_key → child_sku → JAN(8桁+) → 親SKU → 代表1子 → composeフォルダ名。
    """
    pk = str(product_key or "").strip()
    if pk:
        return {"key": pk, "via": "product_key", "jan": str(jan or "").strip()}

    csku = str(child_sku or "").strip()
    if csku:
        return {"key": csku, "via": "child_sku", "jan": str(jan or "").strip()}

    j = str(jan or "").strip()
    if len(j) >= 8 and j.isdigit():
        return {"key": j, "via": "jan", "jan": j}

    parents: List[str] = []
    try:
        rows = _load_master_rows(master_csv)
        from master_sets import load_set_children_for_parent, parents_for_jan_from_rows

        if j:
            parents = parents_for_jan_from_rows(rows, j)
        path = Path(master_csv) if master_csv else DEFAULT_MASTER_CSV
        if parents:
            # JANが短い／非数字でも親が取れたら親を優先
            return {
                "key": parents[0],
                "via": "parent_sku",
                "jan": j,
                "parents": list(parents),
            }
        if j and path.is_file():
            # 親無し時のみ代表子は不可
            pass
    except Exception as e:
        LOG.warning("review key master soft-fail: %s", e)

    # 親リストがある場合の代表1子（上で親returnしなかった分岐用）
    if parents:
        try:
            path = Path(master_csv) if master_csv else DEFAULT_MASTER_CSV
            from master_sets import load_set_children_for_parent

            best_sku = None
            best_n = 10**9
            for p in parents:
                kids, _ = load_set_children_for_parent(path, p)
                for k in kids:
                    n = int(k.set_count) if getattr(k, "set_count", None) else 999
                    if k.child_sku and n < best_n:
                        best_n = n
                        best_sku = k.child_sku
            if best_sku:
                return {
                    "key": best_sku,
                    "via": "rep_child",
                    "jan": j,
                    "parents": list(parents),
                }
        except Exception as e:
            LOG.warning("review key rep_child soft-fail: %s", e)

    if compose_dir:
        token = Path(compose_dir).name.split("_")[0].strip()
        if token:
            return {"key": token, "via": "compose_dir", "jan": j}

    raise ValueError(
        "目視用品番キーを解決できません（--product-key / --child-sku / --jan / compose_dir）"
    )


def export_subs(
    *,
    child_sku: str,
    src_dir: Path,
    out_dir: Path,
    dry_run: bool = False,
    max_subs: int = 10,
) -> List[dict]:
    src_dir = Path(src_dir)
    out_dir = Path(out_dir)
    if not src_dir.is_dir():
        raise SystemExit(f"src-dir not found: {src_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    files = list_source_images(src_dir)
    if not files:
        raise SystemExit(f"no images in {src_dir}")

    report = []
    for i, src in enumerate(files[: max(1, int(max_subs))], start=1):
        dest_name = sub_filename(child_sku, i)
        dest = out_dir / dest_name
        entry = {
            "src": str(src),
            "dest": str(dest),
            "subIndex": i,
            "childSku": child_sku,
        }
        if dry_run:
            LOG.info("DRY %s -> %s", src.name, dest_name)
        else:
            shutil.copy2(src, dest)
            LOG.info("copied %s -> %s", src.name, dest_name)
        report.append(entry)
    if len(files) > max_subs:
        LOG.warning(
            "truncated %s images; only first %s exported as sub1..sub%s",
            len(files),
            max_subs,
            max_subs,
        )
    return report


def export_from_compose(
    *,
    compose_dir: Path,
    child_sku: str,
    out_dir: Path,
    pick: str = "a",
    pick_manifest: Optional[Path] = None,
    provider: str = "openai",
    dry_run: bool = False,
    max_subs: int = 10,
) -> List[dict]:
    compose_dir = Path(compose_dir)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    meta = _load_run_meta(compose_dir)
    # ensure outRoot for path fallback
    if not meta.get("outRoot"):
        meta["outRoot"] = str(compose_dir)
    for j in meta.get("jobs") or []:
        j.setdefault("outRoot", meta["outRoot"])

    pick_mode = _normalize_pick_mode(pick)
    pick_by_slot = _load_pick_manifest(pick_manifest)
    # manifest があるときはスロット別上書きのため a/b モード扱い（ab は全スロット両出し）
    selected = select_jobs_from_meta(
        meta,
        pick_default=pick_mode,
        pick_by_slot=pick_by_slot if pick_mode != "ab" else {},
        provider=provider,
    )
    if not selected:
        raise SystemExit(f"no exportable slots in {compose_dir}")

    cap = max(1, min(int(max_subs), 10))
    report = []
    for i, (slot, job, src) in enumerate(selected[:cap], start=1):
        dest_name = sub_filename(
            child_sku,
            i,
            pattern=job.get("themeSlug") or job.get("themeName") or "",
        )
        dest = out_dir / dest_name
        entry = {
            "src": str(src),
            "dest": str(dest),
            "subIndex": i,
            "slotIndex": slot,
            "phaseOrder": job.get("phaseOrder"),
            "themeId": job.get("themeId"),
            "themeSlug": job.get("themeSlug"),
            "proposal": _proposal_key(job),
            "childSku": child_sku,
            "jan": meta.get("jan"),
        }
        if dry_run:
            LOG.info(
                "DRY S%02d P%s %s -> %s",
                slot,
                job.get("phaseOrder"),
                src.name,
                dest_name,
            )
        else:
            shutil.copy2(src, dest)
            LOG.info(
                "copied S%02d P%s %s -> %s",
                slot,
                job.get("phaseOrder"),
                src.name,
                dest_name,
            )
        report.append(entry)

    meta_out = out_dir / f"_export_meta_{child_sku}.json"
    if not dry_run:
        meta_out.write_text(
            json.dumps(
                {
                    "composeDir": str(compose_dir),
                    "childSku": child_sku,
                    "pickDefault": pick_mode,
                    "provider": provider,
                    "exports": report,
                    "mode": "single_key",
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
    return report


def export_for_human_review(
    *,
    compose_dir: Path,
    out_dir: Path,
    jan: str = "",
    child_sku: str = "",
    product_key: str = "",
    master_csv: Optional[Path] = None,
    pick: str = "ab",
    pick_manifest: Optional[Path] = None,
    provider: str = "openai",
    dry_run: bool = False,
    max_subs: int = REVIEW_MAX_SUBS,
) -> Tuple[List[dict], Dict[str, Any]]:
    """
    目視用: 品番キー1件 × 最大10枚（既定 pick=ab = A/B両出し）。
    全子SKUへの複製はしない。
    """
    compose_dir = Path(compose_dir)
    key_info = resolve_review_product_key(
        jan=jan,
        child_sku=child_sku,
        product_key=product_key,
        master_csv=master_csv,
        compose_dir=compose_dir,
    )
    key = str(key_info["key"])
    LOG.info(
        "human-review export key=%s via=%s jan=%s max=%s pick=%s",
        key,
        key_info.get("via"),
        key_info.get("jan"),
        max_subs,
        pick,
    )
    rows = export_from_compose(
        compose_dir=compose_dir,
        child_sku=key,
        out_dir=out_dir,
        pick=pick,
        pick_manifest=pick_manifest,
        provider=provider,
        dry_run=dry_run,
        max_subs=max_subs,
    )
    return rows, key_info


def resolve_child_sku_arg(
    *,
    child_sku: str,
    parent_sku: str,
    set_count: int,
    master_csv: Optional[Path],
) -> str:
    if child_sku:
        return str(child_sku).strip()
    if not parent_sku or not set_count:
        raise SystemExit("need --child-sku or (--parent-sku and --set-count) or --jan")
    if not master_csv:
        raise SystemExit("--master-csv required with --parent-sku/--set-count")
    from master_sets import load_set_children_for_parent, resolve_child_by_set_count

    children, _ = load_set_children_for_parent(Path(master_csv), parent_sku)
    return resolve_child_by_set_count(children, int(set_count)).child_sku


def resolve_child_skus_for_jan(
    jan: str,
    *,
    master_csv: Optional[Path] = None,
    child_sku: str = "",
) -> List[str]:
    """
    本番複製用: 明示子SKUがあれば1件。無ければ JAN→出品CK子のみ。
    出品CK0件時の全セット子フォールバックは禁止。
    """
    if child_sku and str(child_sku).strip():
        return [str(child_sku).strip()]

    rows = _load_master_rows(master_csv)
    from master_sets import resolve_checked_children_for_jan

    kids = resolve_checked_children_for_jan(jan, rows=rows)
    skus = [k.child_sku for k in kids]
    if skus:
        return skus

    raise ValueError(
        f"JAN={jan}: 出品CK子が0件。全セット子フォールバックは禁止。"
        "目視は品番キー1件（export_for_human_review）。"
        "本番複製はマスタで出品CKを付けてから --to-checked-children。"
    )


def export_compose_to_children(
    *,
    compose_dir: Path,
    child_skus: List[str],
    out_dir: Path,
    pick: str = "a",
    pick_manifest: Optional[Path] = None,
    provider: str = "openai",
    dry_run: bool = False,
    max_subs: int = 10,
) -> List[dict]:
    """同一 compose を複数子SKUへ `{子SKU}_subN.jpg` として複製 export。"""
    all_rows: List[dict] = []
    for sku in child_skus:
        rows = export_from_compose(
            compose_dir=compose_dir,
            child_sku=sku,
            out_dir=out_dir,
            pick=pick,
            pick_manifest=pick_manifest,
            provider=provider,
            dry_run=dry_run,
            max_subs=max_subs,
        )
        all_rows.extend(rows)
    return all_rows


def write_export_plan(
    compose_dir: Path,
    *,
    out_dir: Path,
    child_skus: List[str],
    pick: str,
    provider: str,
    jan: str = "",
    product_key: str = "",
    mode: str = "review",
    key_via: str = "",
) -> Path:
    from work_paths import meta_dir_for

    plan = {
        "composeDir": str(compose_dir),
        "outDir": str(out_dir),
        "childSkus": list(child_skus),
        "pick": pick,
        "provider": provider,
        "jan": jan,
        "productKey": product_key or (child_skus[0] if child_skus else ""),
        "mode": mode,
        "keyVia": key_via,
        "maxSubs": REVIEW_MAX_SUBS,
    }
    path = meta_dir_for(Path(compose_dir)) / "export_plan.json"
    path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def export_from_plan(
    compose_dir: Path,
    *,
    plan_path: Optional[Path] = None,
    dry_run: bool = False,
) -> List[dict]:
    from work_paths import meta_dir_for

    root = Path(compose_dir)
    path = Path(plan_path) if plan_path else (meta_dir_for(root) / "export_plan.json")
    if not path.is_file():
        raise SystemExit(f"export_plan.json がありません: {path}")
    plan = json.loads(path.read_text(encoding="utf-8"))
    return export_compose_to_children(
        compose_dir=root,
        child_skus=list(plan.get("childSkus") or []),
        out_dir=Path(plan.get("outDir") or DEFAULT_RAKUTEN_UPLOAD_DIR),
        pick=str(plan.get("pick") or "a"),
        provider=str(plan.get("provider") or "openai"),
        dry_run=dry_run,
    )


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(
        description="Export sub images for human review (product key × max 10) or checked children"
    )
    ap.add_argument("--child-sku", default="", help="品番キー固定／単一子SKU")
    ap.add_argument("--product-key", default="", help="目視用品番キー明示")
    ap.add_argument("--jan", default="", help="JAN（目視キー優先／本番は出品CK子）")
    ap.add_argument("--parent-sku", default="", help="子SKU省略時: 親SKU")
    ap.add_argument("--set-count", type=int, default=0, help="子SKU省略時: セット数N")
    ap.add_argument("--master-csv", type=Path, default=None)
    ap.add_argument("--src-dir", type=Path, default=None, help="従来: 画像ディレクトリ")
    ap.add_argument(
        "--from-compose-dir",
        type=Path,
        default=None,
        help="compose 出力ルート（run_meta.json）",
    )
    ap.add_argument(
        "--from-plan",
        action="store_true",
        help="compose の _meta/export_plan.json から再export",
    )
    ap.add_argument(
        "--to-checked-children",
        action="store_true",
        help="本番用: 出品CK子へ複製（0件時はエラー。全セット子FO禁止）",
    )
    ap.add_argument(
        "--pick",
        default="ab",
        help="a|b|ab。目視既定=ab（最大10枚）。本番は a 推奨",
    )
    ap.add_argument(
        "--pick-manifest",
        type=Path,
        default=None,
        help='JSON {"1":"a","2":"b"} または {"slots":[{"slotIndex":1,"pick":"a"}]}',
    )
    ap.add_argument(
        "--provider",
        default="openai",
        choices=("openai", "gemini", "fal"),
        help="優先プロバイダ（無ければフォールバック）",
    )
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=DEFAULT_RAKUTEN_UPLOAD_DIR,
        help="既定: 02.楽天アップロード画像保存場所（人間目視の正）",
    )
    ap.add_argument("--max-subs", type=int, default=10)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    out_dir = Path(args.out_dir) if args.out_dir else DEFAULT_RAKUTEN_UPLOAD_DIR

    if args.from_compose_dir and args.from_plan:
        rows = export_from_plan(Path(args.from_compose_dir), dry_run=args.dry_run)
        print(f"exported={len(rows)} out={out_dir}")
        return

    if args.from_compose_dir and args.to_checked_children:
        jan = str(args.jan or "").strip()
        if not jan:
            meta_p = Path(args.from_compose_dir) / "_meta" / "run_meta.json"
            if meta_p.is_file():
                jan = str(json.loads(meta_p.read_text(encoding="utf-8")).get("jan") or "")
        skus = resolve_child_skus_for_jan(
            jan,
            master_csv=args.master_csv,
            child_sku=args.child_sku,
        )
        pick = "a" if str(args.pick).lower() in ("ab", "both", "a+b") else args.pick
        rows = export_compose_to_children(
            compose_dir=Path(args.from_compose_dir),
            child_skus=skus,
            out_dir=out_dir,
            pick=pick,
            pick_manifest=args.pick_manifest,
            provider=args.provider,
            dry_run=args.dry_run,
            max_subs=min(int(args.max_subs), 10),
        )
        write_export_plan(
            Path(args.from_compose_dir),
            out_dir=out_dir,
            child_skus=skus,
            pick=str(pick),
            provider=str(args.provider),
            jan=jan,
            mode="checked_children",
        )
        print(f"exported={len(rows)} children={skus} out={out_dir}")
        return

    if args.from_compose_dir:
        # 既定: 目視（品番キー1件 × 最大10）
        jan = str(args.jan or "").strip()
        if not jan:
            meta_p = Path(args.from_compose_dir) / "_meta" / "run_meta.json"
            if meta_p.is_file():
                jan = str(json.loads(meta_p.read_text(encoding="utf-8")).get("jan") or "")
        if args.parent_sku and args.set_count and not args.child_sku and not args.product_key:
            child = resolve_child_sku_arg(
                child_sku="",
                parent_sku=args.parent_sku,
                set_count=int(args.set_count or 0),
                master_csv=args.master_csv,
            )
            rows = export_from_compose(
                compose_dir=args.from_compose_dir,
                child_sku=child,
                out_dir=out_dir,
                pick=args.pick,
                pick_manifest=args.pick_manifest,
                provider=args.provider,
                dry_run=args.dry_run,
                max_subs=min(int(args.max_subs), 10),
            )
            write_export_plan(
                Path(args.from_compose_dir),
                out_dir=out_dir,
                child_skus=[child],
                pick=str(args.pick),
                provider=str(args.provider),
                jan=jan,
                mode="single",
            )
            print(f"exported={len(rows)} child={child} out={out_dir}")
            return

        rows, key_info = export_for_human_review(
            compose_dir=Path(args.from_compose_dir),
            out_dir=out_dir,
            jan=jan,
            child_sku=str(args.child_sku or ""),
            product_key=str(args.product_key or ""),
            master_csv=args.master_csv,
            pick=str(args.pick or "ab"),
            pick_manifest=args.pick_manifest,
            provider=args.provider,
            dry_run=args.dry_run,
            max_subs=min(int(args.max_subs), 10),
        )
        key = str(key_info.get("key") or "")
        write_export_plan(
            Path(args.from_compose_dir),
            out_dir=out_dir,
            child_skus=[key] if key else [],
            pick=str(args.pick or "ab"),
            provider=str(args.provider),
            jan=jan,
            product_key=key,
            mode="review",
            key_via=str(key_info.get("via") or ""),
        )
        print(
            f"exported={len(rows)} reviewKey={key} via={key_info.get('via')} out={out_dir}"
        )
        return

    child = resolve_child_sku_arg(
        child_sku=args.child_sku,
        parent_sku=args.parent_sku,
        set_count=int(args.set_count or 0),
        master_csv=args.master_csv,
    )
    if not args.src_dir:
        raise SystemExit("need --from-compose-dir or --src-dir")
    rows = export_subs(
        child_sku=child,
        src_dir=args.src_dir,
        out_dir=out_dir,
        dry_run=args.dry_run,
        max_subs=args.max_subs,
    )
    print(f"exported={len(rows)} child={child} out={out_dir}")


if __name__ == "__main__":
    main()
