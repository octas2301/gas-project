# -*- coding: utf-8 -*-
"""
テーマ別新規＋競合本線の混在マニフェスト（段階A/B: 選定・export・再生成）。

例:
  python sub_image_mix_manifest.py --jan 4906283045119 \\
    --theme-new-dir _out/theme_new_run_4906283045119 \\
    --compose-dir "G:/.../sub_image_ai_compose/4906283045119_20260830_101355"
"""
from __future__ import annotations

import argparse
import json
import logging
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from sub_image_lp_themes import THEME_BY_ID
from work_paths import meta_dir_for

LOG = logging.getLogger("set_main_image.mix_manifest")
SCRIPT_DIR = Path(__file__).resolve().parent

PROVIDER_DIRS = {
    "openai": "03_openai",
    "gemini": "02_gemini",
    "fal": "04_fal",
}


def unique_mix_manifest_path(
    jan: str,
    theme_new_dir: Optional[Path] = None,
    compose_dir: Optional[Path] = None,
    *,
    out_root: Optional[Path] = None,
) -> Path:
    """同じJANでも theme-new／compose が違えば別ファイル（同時レビューの上書き防止）。"""
    root = Path(out_root) if out_root else (SCRIPT_DIR / "_out")
    tag = ""
    if theme_new_dir and Path(theme_new_dir).is_dir():
        tag = Path(theme_new_dir).resolve().name
    elif compose_dir and Path(compose_dir).is_dir():
        tag = Path(compose_dir).resolve().name
    safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in tag)[:80]
    if safe:
        return root / f"mix_manifest_{jan}_{safe}.json"
    return root / f"mix_manifest_{jan}.json"
_R_SUFFIX = re.compile(r"__r(\d+)$", re.I)
_THEME_STEM = re.compile(r"^T(\d+)(?:__r(\d+))?$", re.I)


def _theme_meta(tid: int) -> Dict[str, Any]:
    m = THEME_BY_ID.get(int(tid)) or {}
    return {
        "themeId": int(tid),
        "themeName": str(m.get("name") or f"T{int(tid):02d}"),
        "themeSlug": str(m.get("nameSlug") or f"t{int(tid):02d}"),
        "phaseOrder": int(m.get("phaseOrder") or 0),
    }


def _parse_theme_file(name: str) -> Optional[Dict[str, Any]]:
    stem = Path(name).stem
    m = _THEME_STEM.match(stem)
    if not m:
        return None
    gen = int(m.group(2)) if m.group(2) else 1
    return {
        "themeId": int(m.group(1)),
        "stem": stem,
        "gen": gen,
        "baseStem": f"T{int(m.group(1)):02d}",
    }


def cards_from_theme_new(
    theme_new_dir: Path, *, jan: str
) -> tuple[List[Dict[str, Any]], str]:
    root = Path(theme_new_dir)
    cards: List[Dict[str, Any]] = []
    msg_path = root / "messages.json"
    by_tid: Dict[int, Dict[str, Any]] = {}
    product_name = ""
    if msg_path.is_file():
        msg = json.loads(msg_path.read_text(encoding="utf-8"))
        product_name = str(msg.get("productName") or "")
        for r in msg.get("rows") or []:
            try:
                tid = int(r.get("themeId") or 0)
            except (TypeError, ValueError):
                continue
            if tid:
                by_tid[tid] = r
    for p in sorted(root.glob("T*.jpg")):
        parsed = _parse_theme_file(p.name)
        if not parsed:
            continue
        tid = int(parsed["themeId"])
        tm = _theme_meta(tid)
        row = by_tid.get(tid) or {}
        gen = int(parsed["gen"])
        gen_label = "" if gen <= 1 else f" / gen{gen}"
        cards.append(
            {
                "id": f"theme_new_{parsed['stem']}",
                "source": "theme_new",
                "jan": jan,
                "stem": parsed["stem"],
                "baseStem": parsed["baseStem"],
                "gen": gen,
                "themeId": tid,
                "themeName": tm["themeName"],
                "themeSlug": tm["themeSlug"],
                "phaseOrder": tm["phaseOrder"],
                "catch": str(row.get("catch") or ""),
                "path": str(p.resolve()),
                "rel": p.name,
                "composeDir": "",
                "proposal": "a",
                "genLabel": gen_label,
            }
        )
    return cards, product_name


def cards_from_competitor(
    compose_dir: Path,
    *,
    jan: str,
    provider: str = "openai",
) -> tuple[List[Dict[str, Any]], str]:
    root = Path(compose_dir)
    sub = PROVIDER_DIRS.get(provider) or PROVIDER_DIRS["openai"]
    folder = root / sub
    jobs_by_stem: Dict[str, Any] = {}
    meta_path = meta_dir_for(root) / "run_meta.json"
    product_name = ""
    if meta_path.is_file():
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        product_name = str(meta.get("productName") or "")
        for j in meta.get("jobs") or []:
            jobs_by_stem[str(j.get("stem") or "")] = j
    cards: List[Dict[str, Any]] = []
    if not folder.is_dir():
        return cards, product_name
    for p in sorted(folder.glob("*.jpg")):
        if p.name.startswith("__regen_staging_"):
            continue
        stem = p.stem
        base = _R_SUFFIX.sub("", stem)
        m = _R_SUFFIX.search(stem)
        gen = int(m.group(1)) if m else 1
        job = jobs_by_stem.get(stem) or jobs_by_stem.get(base) or {}
        tid = int(job.get("themeId") or 0)
        tm = (
            _theme_meta(tid)
            if tid
            else {
                "themeId": 0,
                "themeName": "",
                "themeSlug": "sub",
                "phaseOrder": 0,
            }
        )
        gen_label = "" if gen <= 1 else f" / gen{gen}"
        cards.append(
            {
                "id": f"competitor_{stem}",
                "source": "competitor",
                "jan": jan,
                "stem": stem,
                "baseStem": base,
                "gen": gen,
                "themeId": tid or tm["themeId"],
                "themeName": str(job.get("themeName") or tm["themeName"]),
                "themeSlug": str(job.get("themeSlug") or tm["themeSlug"]),
                "phaseOrder": int(job.get("phaseOrder") or tm["phaseOrder"] or 0),
                "catch": "",
                "path": str(p.resolve()),
                "rel": f"{sub}/{p.name}".replace("\\", "/"),
                "composeDir": str(root.resolve()),
                "proposal": str(job.get("proposal") or "a"),
                "genLabel": gen_label,
            }
        )
    return cards, product_name


def build_mix_manifest(
    *,
    jan: str,
    theme_new_dir: Optional[Path] = None,
    compose_dir: Optional[Path] = None,
    provider: str = "openai",
    product_name: str = "",
) -> Dict[str, Any]:
    cards: List[Dict[str, Any]] = []
    pname = product_name
    if theme_new_dir and Path(theme_new_dir).is_dir():
        tn_cards, tn_name = cards_from_theme_new(Path(theme_new_dir), jan=jan)
        cards.extend(tn_cards)
        pname = pname or tn_name
    if compose_dir and Path(compose_dir).is_dir():
        cp_cards, cp_name = cards_from_competitor(
            Path(compose_dir), jan=jan, provider=provider
        )
        cards.extend(cp_cards)
        pname = pname or cp_name
    # 表示順: themeId → theme_new 先 → competitor → 世代新しい順（同テーマ内）
    cards.sort(
        key=lambda c: (
            int(c.get("phaseOrder") or 99),
            int(c.get("themeId") or 99),
            0 if c.get("source") == "theme_new" else 1,
            -int(c.get("gen") or 1),
            str(c.get("id") or ""),
        )
    )
    food_canon: Dict[str, Any] = {}
    if theme_new_dir and Path(theme_new_dir).is_dir():
        try:
            from food_unit_canon import ensure_for_theme_new_dir

            food_canon = ensure_for_theme_new_dir(Path(theme_new_dir), force=False)
        except Exception as e:
            LOG.warning("food canon attach failed: %s", e)
            food_canon = {"ok": False, "reason": str(e)}
    parent_sku = ""
    try:
        from b5_skip_generate_finish import parent_sku_from_path_hint

        for key in ("truth", "path", "chosenName"):
            parent_sku = parent_sku_from_path_hint(str((food_canon or {}).get(key) or ""))
            if parent_sku:
                break
    except Exception as e:
        LOG.warning("parentSku from foodCanon failed: %s", e)
    return {
        "version": 2,
        "jan": jan,
        "parentSku": parent_sku,
        "productName": pname,
        "builtAt": datetime.now(timezone.utc).isoformat(),
        "themeNewDir": str(Path(theme_new_dir).resolve()) if theme_new_dir else "",
        "composeDir": str(Path(compose_dir).resolve()) if compose_dir else "",
        "provider": provider,
        "foodCanon": food_canon,
        "cards": cards,
        "note": "段階B: 独自・競合とも再生成可。食材は FOOD_UNIT_CANON で横断統一。出口は親1セット（子SKU名）。",
    }


def write_mix_manifest(manifest: Dict[str, Any], out_path: Path) -> Path:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    LOG.info("wrote mix manifest n=%s -> %s", len(manifest.get("cards") or []), out_path)
    return out_path


def load_mix_manifest(path: Path) -> Dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="混在マニフェスト作成")
    ap.add_argument("--jan", required=True)
    ap.add_argument("--theme-new-dir", type=Path, default=None)
    ap.add_argument("--compose-dir", type=Path, default=None)
    ap.add_argument("--provider", default="openai")
    ap.add_argument(
        "--out",
        type=Path,
        default=None,
        help="省略時は theme-new／compose フォルダ名付き（同じJANの上書き防止）",
    )
    args = ap.parse_args(argv)
    if not args.theme_new_dir and not args.compose_dir:
        raise SystemExit("--theme-new-dir か --compose-dir が必要です")
    man = build_mix_manifest(
        jan=str(args.jan),
        theme_new_dir=args.theme_new_dir,
        compose_dir=args.compose_dir,
        provider=str(args.provider),
    )
    out = args.out or unique_mix_manifest_path(
        str(args.jan), args.theme_new_dir, args.compose_dir
    )
    write_mix_manifest(man, out)
    print(json.dumps({"ok": True, "n": len(man["cards"]), "out": str(out)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
