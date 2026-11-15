# -*- coding: utf-8 -*-
"""おだし系（まぐろ／ぶり）本文3枚を assets/yahoo_lib_src へコピーし、lib_ready を生成。

正本（かつお本採用と同系統）:
  G:/マイドライブ/05.画像生成（セットMAIN）/00.テスト出力/sub_image_ai_compose/
    {JAN}_{run}/03_openai/S*_T03|T04|T11_*.jpg

※ theme_new_full_* は不採用（古い生成パイプライン）。使わない。

  python tools/yahoo_page_mock/prep_odashi_lib_bodies.py
  → assets/yahoo_lib_src/{maguro|buri}_body_*.jpg
  → tools/yahoo_page_mock/out/lib_ready/（圧縮JPEG）

続けて: python tools/yahoo_page_mock/upload_lib_images_api.py
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT_SRC = ROOT / "assets" / "yahoo_lib_src"
COMPOSE = Path(
    r"G:/マイドライブ/05.画像生成（セットMAIN）/00.テスト出力/sub_image_ai_compose"
)

# かつお HUMAN_IMAGE_TEST と同テーマ（T03 シズル / T04 味 / T11 アレンジ）
SPECS = [
    {
        "slug": "maguro",
        "run": "4906283045317_20260830_160211",
        "files": {
            "body_t03_sizzle.jpg": "S02_T03_shizuru_sizzle_ABa.jpg",
            "body_t04_aji.jpg": "S03_T04_aji_kitai_ABa.jpg",
            "body_t11_arrange.jpg": "S04_T11_arrange_recipe_ABa.jpg",
        },
    },
    {
        "slug": "buri",
        "run": "4906283047410_20260831_123235",
        "files": {
            "body_t03_sizzle.jpg": "S01_T03_shizuru_sizzle_ABa.jpg",
            "body_t04_aji.jpg": "S02_T04_aji_kitai_ABa.jpg",
            "body_t11_arrange.jpg": "S03_T11_arrange_recipe_ABa.jpg",
        },
    },
]


def main() -> int:
    OUT_SRC.mkdir(parents=True, exist_ok=True)
    for spec in SPECS:
        folder = COMPOSE / spec["run"] / "03_openai"
        if not folder.is_dir():
            print("MISSING_DIR", folder, file=sys.stderr)
            return 1
        for dest_suffix, src_name in spec["files"].items():
            src = folder / src_name
            if not src.is_file():
                print("MISSING", src, file=sys.stderr)
                return 1
            dest = OUT_SRC / f"{spec['slug']}_{dest_suffix}"
            shutil.copy2(src, dest)
            print("copied", dest.name, "<-", spec["run"], src_name, dest.stat().st_size)
    prep = ROOT / "tools" / "yahoo_page_mock" / "prep_lib_images.py"
    r = subprocess.run(
        [
            sys.executable,
            str(prep),
            "--src",
            str(OUT_SRC),
            "--out",
            str(ROOT / "tools" / "yahoo_page_mock" / "out" / "lib_ready"),
        ],
        cwd=str(ROOT),
    )
    return int(r.returncode)


if __name__ == "__main__":
    raise SystemExit(main())
