# -*- coding: utf-8 -*-
"""
出力フォルダ直下を「画像のみ」に整える。

非画像（json / txt / logic 等）は _meta/ へ移動（削除しない）。
desktop.ini とサブフォルダは対象外。

  python tidy_images_only_folder.py
  python tidy_images_only_folder.py --folder "G:\\...\\07.白抜きの置き場（人間が入れる）"
  python tidy_images_only_folder.py --also-test-out
"""
from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from work_paths import (
    DEFAULT_AMAZON_07_OUT,
    SUB_META,
    SUB_TEST_OUT,
    default_work_root,
)

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}
SKIP_NAMES = {"desktop.ini", "thumbs.db"}


def tidy_folder_images_only(folder: Path) -> Dict[str, Any]:
    folder = Path(folder)
    report: Dict[str, Any] = {
        "folder": str(folder),
        "at": datetime.now().isoformat(timespec="seconds"),
        "moved": [],
        "skipped": [],
        "errors": [],
    }
    if not folder.is_dir():
        report["errors"].append(f"not_a_dir:{folder}")
        return report

    meta = folder / SUB_META
    meta.mkdir(parents=True, exist_ok=True)

    for p in sorted(folder.iterdir()):
        if p.is_dir():
            continue
        name_l = p.name.lower()
        if name_l in SKIP_NAMES:
            report["skipped"].append(p.name)
            continue
        if p.suffix.lower() in IMAGE_EXTS:
            continue
        dest = meta / p.name
        if dest.exists():
            stem, suf = p.stem, p.suffix
            dest = meta / f"{stem}__moved_{datetime.now().strftime('%Y%m%d_%H%M%S')}{suf}"
        try:
            shutil.move(str(p), str(dest))
            report["moved"].append({"from": p.name, "to": str(dest)})
        except OSError as e:
            report["errors"].append(f"{p.name}:{e}")
    return report


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Move non-image files to _meta/")
    ap.add_argument(
        "--folder",
        type=Path,
        default=None,
        help="対象フォルダ（省略時は 07.白抜きの置き場）",
    )
    ap.add_argument(
        "--also-test-out",
        action="store_true",
        help="05…/00.テスト出力 も同様に整理",
    )
    args = ap.parse_args(argv)

    targets: List[Path] = []
    targets.append(args.folder or DEFAULT_AMAZON_07_OUT)
    if args.also_test_out:
        targets.append(default_work_root() / SUB_TEST_OUT)

    reports = []
    for t in targets:
        r = tidy_folder_images_only(t)
        reports.append(r)
        print(
            json.dumps(
                {
                    "folder": r["folder"],
                    "moved": len(r["moved"]),
                    "errors": len(r["errors"]),
                },
                ensure_ascii=False,
            )
        )

    # レポート自体も 07/_meta へ
    out_meta = (args.folder or DEFAULT_AMAZON_07_OUT) / SUB_META
    out_meta.mkdir(parents=True, exist_ok=True)
    rep_path = out_meta / f"TIDY_IMAGES_ONLY_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    rep_path.write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"report -> {rep_path}")
    return 0 if all(not r.get("errors") for r in reports) else 1


if __name__ == "__main__":
    raise SystemExit(main())
