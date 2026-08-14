# -*- coding: utf-8 -*-
"""
01.amazon白抜きベース の透過マット2点チェック。

  ① 透過PNGか
  ② Canva背景リムーバー実施の見込みか

  python check_amazon_base_matte.py
  python check_amazon_base_matte.py --require-ok
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import List

from transparent_bg import inspect_alpha
from work_paths import SUB_AMAZON, default_work_root


def main(argv: List[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Amazon base matte 2-point check")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument(
        "--require-ok",
        action="store_true",
        help="1枚でも matteOk=False なら exit 2",
    )
    args = ap.parse_args(argv)

    folder = (args.work_root or default_work_root()) / SUB_AMAZON
    if not folder.is_dir():
        print(f"フォルダがありません: {folder}")
        return 1

    results = []
    ok_n = 0
    ng_n = 0
    for p in sorted(folder.iterdir()):
        if not p.is_file():
            continue
        if p.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
            continue
        info = inspect_alpha(p, announce=True)
        results.append(info)
        if info.get("matteOk"):
            ok_n += 1
        else:
            ng_n += 1

    summary = {
        "folder": str(folder),
        "ok": ok_n,
        "ng": ng_n,
        "results": results,
    }
    print(json.dumps({"ok": ok_n, "ng": ng_n, "folder": str(folder)}, ensure_ascii=False, indent=2))
    # サマリはテスト出力の _meta へ（07直下を汚さない）
    out_dir = (args.work_root or default_work_root()) / "00.テスト出力" / "_meta"
    out_dir.mkdir(parents=True, exist_ok=True)
    sp = out_dir / "MATTE_CHECK_latest.json"
    sp.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"summary -> {sp}")

    if not results:
        print("画像がありません（01直下）")
        return 1
    if args.require_ok and ng_n > 0:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
