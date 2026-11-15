# -*- coding: utf-8 -*-
"""B-④ 既存ラン再利用の自動判定（人が毎回選ばない）。

正本: docs/org/D_MENU_B5_RAKUTEN_MAIN_BASE_REQUIREMENTS.md §3.6.2
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from work_paths import DEFAULT_WORK_ROOT


def _compose_root(work_root: Optional[Path] = None) -> Path:
    from b5_main_base_review_launch import compose_root

    return compose_root(Path(work_root or DEFAULT_WORK_ROOT))


def _load_checked(path: Path) -> Dict[str, Any]:
    raw = path.read_text(encoding="utf-8-sig")
    obj = json.loads(raw)
    if not isinstance(obj, dict):
        raise SystemExit("checked JSON は object である必要がある")
    return obj


def _latest_run_for_parent(
    root: Path, *, jan: str, parent_sku: str, short_name: str
) -> Optional[Path]:
    """親に一致する最新 compose ラン（比較に画像があるもの）。"""
    from b5_main_base_review_launch import list_latest_info

    runs = list_latest_info(root) or []
    jan_s = str(jan or "").strip()
    sku_s = str(parent_sku or "").strip()
    short_s = str(short_name or "").strip()
    for row in runs:
        if not isinstance(row, dict):
            continue
        if sku_s and str(row.get("parentSku") or "").strip() == sku_s:
            d = Path(str(row.get("composeDir") or ""))
            if d.is_dir():
                return d
        if jan_s and str(row.get("jan") or "").strip() == jan_s:
            if short_s and str(row.get("shortName") or "").strip() not in ("", short_s):
                continue
            d = Path(str(row.get("composeDir") or ""))
            if d.is_dir():
                return d
        if short_s and str(row.get("shortName") or "").strip() == short_s:
            d = Path(str(row.get("composeDir") or ""))
            if d.is_dir():
                return d
    return None


def _compare_stems(compose_dir: Path) -> List[str]:
    cmp_dir = Path(compose_dir) / "比較"
    if not cmp_dir.is_dir():
        return []
    out: List[str] = []
    for p in list(cmp_dir.glob("*.jpg")) + list(cmp_dir.glob("*.png")):
        if p.name.startswith("__regen_staging_"):
            continue
        out.append(p.stem)
    return out


def _catch_present(catch: str, stems: Sequence[str]) -> bool:
    c = str(catch or "").strip()
    if not c:
        return False
    for stem in stems:
        if c in str(stem):
            return True
    return False


def _meta_text_color(compose_dir: Path) -> str:
    prop = Path(compose_dir) / "b5_proposals.json"
    if not prop.is_file():
        return ""
    try:
        obj = json.loads(prop.read_text(encoding="utf-8-sig"))
    except Exception:
        return ""
    if not isinstance(obj, dict):
        return ""
    return str(obj.get("textColor") or obj.get("text_color") or "").strip()


def _meta_package_truth(compose_dir: Path) -> str:
    prop = Path(compose_dir) / "b5_proposals.json"
    if not prop.is_file():
        return ""
    try:
        obj = json.loads(prop.read_text(encoding="utf-8-sig"))
    except Exception:
        return ""
    if not isinstance(obj, dict):
        return ""
    return str(obj.get("packageTruth") or obj.get("package_truth") or "").strip()


def decide(
    checked: Dict[str, Any],
    *,
    work_root: Optional[Path] = None,
    force_full: bool = False,
) -> Dict[str, Any]:
    """mode: full | reuse | missing_only"""
    jan = str(checked.get("jan") or "").strip()
    parent_sku = str(checked.get("parentSku") or checked.get("parent_sku") or "").strip()
    short = str(checked.get("shortName") or checked.get("short_name") or "").strip()
    text_color = str(checked.get("textColor") or checked.get("text_color") or "").strip()
    truth = str(checked.get("packageTruth") or checked.get("package_truth") or "").strip()
    proposals = [
        p for p in (checked.get("proposals") or []) if isinstance(p, dict)
    ]
    root = _compose_root(work_root)
    run = _latest_run_for_parent(root, jan=jan, parent_sku=parent_sku, short_name=short)
    out: Dict[str, Any] = {
        "ok": True,
        "jan": jan,
        "parentSku": parent_sku,
        "shortName": short,
        "composeDir": str(run) if run else "",
        "mode": "full",
        "reason": "",
        "missing": [],
        "present": [],
        "forceFull": bool(force_full),
    }
    if force_full:
        out["mode"] = "full"
        out["reason"] = "人が全部作り直しを指定"
        return out
    if run is None:
        out["mode"] = "full"
        out["reason"] = "比較ランが無い"
        return out
    stems = _compare_stems(run)
    if not stems:
        out["mode"] = "full"
        out["reason"] = "比較に画像が無い"
        return out
    old_color = _meta_text_color(run)
    old_truth = _meta_package_truth(run)
    if text_color and old_color and text_color != old_color:
        out["mode"] = "full"
        out["reason"] = f"文字色が違う（旧={old_color} 新={text_color}）"
        return out
    if truth and old_truth:
        try:
            same = Path(truth).resolve() == Path(old_truth).resolve()
        except Exception:
            same = str(truth).replace("\\", "/") == str(old_truth).replace("\\", "/")
        if not same:
            out["mode"] = "full"
            out["reason"] = "正本袋（packageTruth）が違う"
            return out
    missing: List[Dict[str, Any]] = []
    present: List[Dict[str, Any]] = []
    for p in proposals:
        catch = str(p.get("catchBurn") or p.get("catch") or "").strip()
        tid = str(p.get("themeId") or "").strip()
        row = {"themeId": tid, "catch": catch}
        if _catch_present(catch, stems):
            present.append(row)
        else:
            missing.append(row)
    out["missing"] = missing
    out["present"] = present
    if not proposals:
        out["mode"] = "reuse"
        out["reason"] = "案0件（レビューのみ）"
        return out
    if not missing:
        out["mode"] = "reuse"
        out["reason"] = "今回のキャッチが比較にすべてある"
        return out
    if len(missing) == len(proposals):
        out["mode"] = "full"
        out["reason"] = "チェック案のキャッチがすべて未作成"
        return out
    out["mode"] = "missing_only"
    out["reason"] = f"足りないキャッチ {len(missing)} 件（既存へ足す）"
    return out


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="B-④ 再利用モード自動判定")
    ap.add_argument("--from-json", type=Path, required=True)
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument(
        "--full",
        action="store_true",
        help="人が全部作り直しを指定したとき",
    )
    args = ap.parse_args(list(argv) if argv is not None else None)
    checked = _load_checked(Path(args.from_json))
    out = decide(checked, work_root=args.work_root, force_full=bool(args.full))
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
