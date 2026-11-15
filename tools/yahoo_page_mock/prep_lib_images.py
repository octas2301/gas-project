# -*- coding: utf-8 -*-
"""Yahoo追加画像向け: リサイズ＋JPEG圧縮（Pillow）。

使い方:
  python tools/yahoo_page_mock/prep_lib_images.py
  python tools/yahoo_page_mock/prep_lib_images.py --src assets/yahoo_lib_src --out tools/yahoo_page_mock/out/lib_ready

既定:
  - 長辺 max 1200px
  - JPEG quality 82（目標を超える場合は下げて再試行）
  - 1枚 max 280KB
出力先のファイルを、ストアクリエイターPro「追加画像」へ手動UPする。
"""
from __future__ import annotations

import argparse
import io
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
DEFAULT_SRC = ROOT / "assets" / "yahoo_lib_src"
DEFAULT_FALLBACK = ROOT / "assets" / "yahoo_coupons"
DEFAULT_OUT = Path(__file__).resolve().parent / "out" / "lib_ready"
EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif"}


def compress_one(
    src: Path,
    dest: Path,
    *,
    max_side: int = 1200,
    quality: int = 82,
    max_kb: int = 280,
) -> dict:
    im = Image.open(src)
    if im.mode in ("RGBA", "P"):
        bg = Image.new("RGB", im.size, (255, 255, 255))
        rgba = im.convert("RGBA")
        bg.paste(rgba, mask=rgba.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")

    w, h = im.size
    scale = min(1.0, float(max_side) / max(w, h))
    if scale < 1.0:
        im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.Resampling.LANCZOS)

    q = quality
    data = b""
    while q >= 45:
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=q, optimize=True, progressive=True)
        data = buf.getvalue()
        if len(data) <= max_kb * 1024:
            break
        q -= 7

    dest.parent.mkdir(parents=True, exist_ok=True)
    # 追加画像は拡張子必須・半角ファイル名推奨
    out_name = dest.stem.lower().replace(" ", "_") + ".jpg"
    out_path = dest.parent / out_name
    out_path.write_bytes(data)
    return {
        "src": str(src),
        "out": str(out_path),
        "px": f"{im.size[0]}x{im.size[1]}",
        "kb": round(len(data) / 1024, 1),
        "quality": q,
    }


def collect_inputs(src_dir: Path) -> list[Path]:
    if not src_dir.is_dir():
        return []
    files = []
    for p in sorted(src_dir.rglob("*")):
        if p.is_file() and p.suffix.lower() in EXTS and not p.name.startswith("."):
            files.append(p)
    return files


def main() -> None:
    ap = argparse.ArgumentParser(description="Prepare Yahoo additional (lib) images")
    ap.add_argument("--src", type=Path, default=DEFAULT_SRC)
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT)
    ap.add_argument("--max-side", type=int, default=1200)
    ap.add_argument("--quality", type=int, default=82)
    ap.add_argument("--max-kb", type=int, default=280)
    ap.add_argument(
        "--also-coupons-cta",
        action="store_true",
        help="assets/yahoo_coupons の cta_gourmet_sets/bestsellers も処理",
    )
    args = ap.parse_args()

    inputs = collect_inputs(args.src)
    if args.also_coupons_cta or not inputs:
        for name in ("cta_gourmet_sets.jpg", "cta_gourmet_bestsellers.jpg"):
            p = DEFAULT_FALLBACK / name
            if p.is_file():
                inputs.append(p)

    if not inputs:
        print("NO_INPUT")
        print(f"Place source images under: {args.src}")
        print("Or pass --also-coupons-cta")
        return

    print(f"src_count\t{len(inputs)}")
    total = 0.0
    for src in inputs:
        # yahoo_lib_src 以下は相対パスをファイル名に畳む
        if args.src in src.parents or src.parent == args.src:
            rel = src.relative_to(args.src) if args.src in src.parents or src.parent == args.src else src.name
            try:
                rel = src.relative_to(args.src)
                dest = args.out / rel
            except ValueError:
                dest = args.out / src.name
        else:
            dest = args.out / src.name
        info = compress_one(
            src,
            dest,
            max_side=args.max_side,
            quality=args.quality,
            max_kb=args.max_kb,
        )
        total += info["kb"]
        print(f"OK\t{info['kb']}KB\tq{info['quality']}\t{info['px']}\t{Path(info['out']).name}")
    print(f"TOTAL_KB\t{round(total, 1)}")
    print(f"OUT_DIR\t{args.out}")


if __name__ == "__main__":
    main()
