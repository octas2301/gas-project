# -*- coding: utf-8 -*-
"""
fal.ai 単体スモーク（課金少なめ）

1. テキストのみ: flux/schnell × 1
2. 参照編集: flux-kontext/dev × 1（--ref 必須）

キー: FAL_KEY または secrets/fal_api_key.txt
"""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

from fal_image import (
    DEFAULT_EDIT_ENDPOINT,
    DEFAULT_TXT_ENDPOINT,
    generate_text_to_image,
    generate_with_references,
)
from work_paths import default_work_root, meta_dir_for

LOG = logging.getLogger("set_main_image.fal_poc")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="fal.ai 画像スモーク")
    ap.add_argument("--prompt", default="Amazon food product secondary image, beige background, clean layout, Japanese grocery")
    ap.add_argument("--ref", type=Path, default=None, help="参照画像（Kontext編集）")
    ap.add_argument("--endpoint", default=None, help="上書き endpoint")
    ap.add_argument("--txt-only", action="store_true", help="schnell のみ（参照無視）")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_dir = default_work_root() / "00.テスト出力" / "sub_image_fal_poc" / run_id
    out_dir.mkdir(parents=True, exist_ok=True)
    meta_dir_for(out_dir)

    if args.txt_only or not args.ref:
        data, ep = generate_text_to_image(
            prompt=args.prompt,
            endpoint=args.endpoint or DEFAULT_TXT_ENDPOINT,
        )
        mode = "txt2img"
    else:
        data, ep = generate_with_references(
            prompt=args.prompt,
            image_paths=[args.ref],
            endpoint=args.endpoint or DEFAULT_EDIT_ENDPOINT,
        )
        mode = "edit"

    out = out_dir / f"{mode}.jpg"
    out.write_bytes(data)
    (out_dir / "meta.txt").write_text(
        f"mode={mode}\nendpoint={ep}\nprompt={args.prompt}\nref={args.ref}\n",
        encoding="utf-8",
    )
    LOG.info("wrote %s endpoint=%s", out, ep)
    print(str(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
