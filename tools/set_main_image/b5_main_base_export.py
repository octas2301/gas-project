# -*- coding: utf-8 -*-
"""
B-⑤ P3 — 採用1枚を 02.楽天ベース へコピー（比較フォルダは残す）。

  python b5_main_base_export.py --compose-dir "...\\かつお_20260829_154935" --pick-file "かつお_T4_忙しい日の一品に.jpg"
  python b5_main_base_export.py --compose-dir "..." --theme-id T4

出口名:
  単一: {親SKU}_単体.png
  P3b:  --adopt-key T4 → {親SKU}_単体__T4.png（複数回呼べば複数枚）
JPG は PNG に変換して書く。楽天UL（子SKU_rakuten.jpg）は書かない（C量産）。
マップ JSON はブラウザ review_loop の一括 export が正（本CLIは単枚）。
"""
from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from PIL import Image

from b5_main_base_names import export02_filename
from b5_set_base_map import adopt_key_from_filename
from work_paths import DEFAULT_WORK_ROOT, SUB_RAKUTEN

LOG = logging.getLogger("set_main_image.b5_main_base_export")


def _load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def _resolve_pick(
    compose_dir: Path,
    *,
    pick_file: str = "",
    theme_id: str = "",
) -> Path:
    pick_dir = compose_dir / "比較"
    if not pick_dir.is_dir():
        raise SystemExit(f"比較フォルダが無い: {pick_dir}")

    if str(pick_file or "").strip():
        p = pick_dir / Path(pick_file).name
        if not p.is_file():
            # ルート直下にもあった旧配置
            alt = compose_dir / Path(pick_file).name
            if alt.is_file():
                return alt
            raise SystemExit(f"pick が無い: {p}")
        return p

    tid = str(theme_id or "").strip()
    if not tid:
        raise SystemExit("--pick-file か --theme-id が必要です")

    # summary / ファイル名に themeId が含まれるもの
    cands = sorted(pick_dir.glob(f"*_{tid}_*.jpg")) + sorted(pick_dir.glob(f"*_{tid}_*.png"))
    if not cands:
        cands = sorted(pick_dir.glob(f"{tid}.jpg")) + sorted(pick_dir.glob(f"{tid}.png"))
    if len(cands) == 1:
        return cands[0]
    if not cands:
        raise SystemExit(f"themeId={tid} の画像が比較に無い")
    raise SystemExit(
        f"themeId={tid} が複数あります。--pick-file で指定: "
        + ", ".join(p.name for p in cands)
    )


def export_adopted(
    *,
    compose_dir: Path,
    pick_path: Path,
    work_root: Optional[Path] = None,
    parent_sku: str = "",
    adopt_key: str = "",
    dry_run: bool = False,
) -> Dict[str, Any]:
    compose_dir = Path(compose_dir)
    pick_path = Path(pick_path)
    if not pick_path.is_file():
        raise SystemExit(f"pick が無い: {pick_path}")

    proposals_path = compose_dir / "b5_proposals.json"
    summary_path = compose_dir / "summary.json"
    meta: Dict[str, Any] = {}
    if proposals_path.is_file():
        meta.update(_load_json(proposals_path))
    if summary_path.is_file():
        meta.update(_load_json(summary_path))

    sku = str(parent_sku or meta.get("parentSku") or meta.get("parent_sku") or "").strip()
    if not sku:
        raise SystemExit("parentSku が JSON にも --parent-sku にも無い")

    akey = str(adopt_key or "").strip() or adopt_key_from_filename(pick_path.name)
    work = Path(work_root or DEFAULT_WORK_ROOT)
    dest_dir = work / SUB_RAKUTEN
    dest_dir.mkdir(parents=True, exist_ok=True)
    compose_name = compose_dir.name
    dest = dest_dir / export02_filename(sku, adopt_key=akey, collide=False)
    collide = dest.is_file()
    if collide:
        dest = dest_dir / export02_filename(
            sku, adopt_key=akey, compose_dir=compose_name, collide=True
        )

    report = {
        "ok": True,
        "parentSku": sku,
        "jan": meta.get("jan") or "",
        "shortName": meta.get("shortName") or "",
        "composeDir": str(compose_dir),
        "pickFile": str(pick_path),
        "dest": str(dest),
        "adoptKey": akey,
        "collided": collide,
        "compareKept": True,
        "dryRun": dry_run,
        "note": "比較は残す。楽天ULは書かない（C量産）。マップは review_loop 一括が正。",
    }

    if dry_run:
        LOG.info("dry-run would copy %s -> %s", pick_path, dest)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return report

    # JPG→PNG（要件の _単体.png）
    im = Image.open(pick_path).convert("RGBA")
    im.save(dest, format="PNG")
    LOG.info("exported %s -> %s (compare kept)", pick_path.name, dest)

    # 採用記録
    adopt = {
        "parentSku": sku,
        "pickFile": pick_path.name,
        "dest": str(dest),
        "composeDir": compose_name,
    }
    (compose_dir / "adopted.json").write_text(
        json.dumps(adopt, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return report


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="B-⑤ 採用1枚 → 02.楽天ベース（比較は残す）")
    ap.add_argument("--compose-dir", type=Path, required=True, help="compose ランのルート")
    ap.add_argument("--pick-file", default="", help="比較内のファイル名")
    ap.add_argument("--theme-id", default="", help="pick-file 省略時。一意ならそれで選ぶ")
    ap.add_argument("--parent-sku", default="", help="JSON に無いとき上書き")
    ap.add_argument("--adopt-key", default="", help="例 T4 / T4__r2。省略時はファイル名から推定")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args(argv)

    compose_dir = Path(args.compose_dir)
    if not compose_dir.is_dir():
        raise SystemExit(f"--compose-dir が無い: {compose_dir}")
    pick = _resolve_pick(
        compose_dir, pick_file=str(args.pick_file or ""), theme_id=str(args.theme_id or "")
    )
    export_adopted(
        compose_dir=compose_dir,
        pick_path=pick,
        work_root=args.work_root,
        parent_sku=str(args.parent_sku or ""),
        adopt_key=str(args.adopt_key or ""),
        dry_run=bool(args.dry_run),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
