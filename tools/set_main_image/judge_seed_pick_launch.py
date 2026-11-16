# -*- coding: utf-8 -*-
"""判定種（judgeSeed）ラジオ選択の起動。

例:
  python judge_seed_pick_launch.py --seed path\\to\\seed.jpg
  python judge_seed_pick_launch.py --seed path\\to\\seed.jpg --port 8850 --fresh-port
  python judge_seed_pick_launch.py --seed path\\to\\seed.jpg --auto-if-single
  python judge_seed_pick_launch.py --seed path\\to\\seed.jpg --build-only

正本: docs/org/PLATING_AND_PIECE_GENERIC_LOGIC.md §3.3
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from judge_seed_pick import (
    DEFAULT_PORT,
    _write_index,
    build_candidates,
    find_free_port,
    resolve_judge_seed,
)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="judgeSeed 候補→ラジオ選択")
    ap.add_argument("--seed", required=True, type=Path, help="種画像（複数片でも可）")
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="作業フォルダ（省略時 seed隣の judge_seed_pick_タイムスタンプ）",
    )
    ap.add_argument("--port", type=int, default=DEFAULT_PORT)
    ap.add_argument("--fresh-port", action="store_true", help="空きポートを探す")
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument(
        "--build-only",
        action="store_true",
        help="候補とindexだけ作って終了（サーバなし）",
    )
    ap.add_argument(
        "--auto-if-single",
        action="store_true",
        help="候補が実質1つならラジオを飛ばして自動確定",
    )
    ap.add_argument("--timeout-sec", type=float, default=None)
    args = ap.parse_args(argv)

    seed = args.seed.resolve()
    if not seed.is_file():
        print(f"seed missing: {seed}", file=sys.stderr)
        return 2

    if args.out_dir:
        out = args.out_dir.resolve()
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = seed.parent / f"judge_seed_pick_{stamp}"

    port = find_free_port(args.port) if args.fresh_port else int(args.port)

    if args.build_only:
        meta = build_candidates(seed, out)
        _write_index(out, meta)
        print(f"JUDGE_SEED_PICK_DIR {out}")
        print(f"CANDIDATES {len(meta.get('candidates') or [])}")
        return 0

    jp = resolve_judge_seed(
        seed,
        out,
        port=port,
        open_browser=not args.no_browser,
        auto_if_single=bool(args.auto_if_single),
        timeout_sec=args.timeout_sec,
    )
    print(f"JUDGE_SEED {jp}")
    meta = json.loads((out / "candidates.json").read_text(encoding="utf-8"))
    print(f"SELECTED {meta.get('selectedId')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
