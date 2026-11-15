# -*- coding: utf-8 -*-
"""
混在レビュー一括起動（独自制作＋競合）。

例:
  python sub_image_mix_review_launch.py --jan 4906283045119 --jan 4906283045317 --no-wait
"""
from __future__ import annotations

import argparse
import logging
import subprocess
import sys
import time
from pathlib import Path
from typing import List, Optional, Tuple

from sub_image_mix_manifest import (
    build_mix_manifest,
    unique_mix_manifest_path,
    write_mix_manifest,
)
from sub_image_review_launch import compose_root, latest_compose_dir
from work_paths import DEFAULT_WORK_ROOT

LOG = logging.getLogger("set_main_image.mix_review_launch")
SCRIPT_DIR = Path(__file__).resolve().parent
MIX_LOOP = SCRIPT_DIR / "sub_image_mix_review_loop.py"


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )


def resolve_theme_new_dir(jan: str, explicit: Optional[Path] = None) -> Optional[Path]:
    if explicit and Path(explicit).is_dir():
        return Path(explicit)
    cand = SCRIPT_DIR / "_out" / f"theme_new_run_{jan}"
    if cand.is_dir() and any(cand.glob("T*.jpg")):
        return cand
    # 新しい overview_copy_* も候補
    out = SCRIPT_DIR / "_out"
    if not out.is_dir():
        return None
    cands = [
        p
        for p in out.iterdir()
        if p.is_dir() and jan in p.name and any(p.glob("T*.jpg"))
    ]
    if not cands:
        return None
    return max(cands, key=lambda p: p.name)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="混在レビュー一括起動")
    ap.add_argument("--jan", action="append", default=[], required=False)
    ap.add_argument("--work-root", type=Path, default=DEFAULT_WORK_ROOT)
    ap.add_argument("--compose-root", type=Path, default=None)
    ap.add_argument("--theme-new-dir", action="append", default=[], help="JAN=path 形式可。または順に --jan と対応")
    ap.add_argument("--base-port", type=int, default=8780)
    ap.add_argument("--provider", default="openai")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--no-wait", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    jans = [str(j).strip() for j in (args.jan or []) if str(j).strip()]
    if not jans:
        raise SystemExit("--jan を1つ以上指定してください")

    # theme-new-dir: "JAN=path" or positional match
    tn_map = {}
    tn_list = []
    for raw in args.theme_new_dir or []:
        s = str(raw).strip()
        if "=" in s:
            j, p = s.split("=", 1)
            tn_map[j.strip()] = Path(p.strip())
        else:
            tn_list.append(Path(s))

    root = Path(args.compose_root) if args.compose_root else compose_root(Path(args.work_root))
    targets: List[Tuple[str, Path, Optional[Path], Optional[Path]]] = []
    for i, jan in enumerate(jans):
        tn = tn_map.get(jan) or (tn_list[i] if i < len(tn_list) else None)
        tn = resolve_theme_new_dir(jan, tn)
        cp = latest_compose_dir(root, jan) if root.is_dir() else None
        if not tn and not cp:
            LOG.warning("JAN=%s: theme-new も compose も無し", jan)
            continue
        # 同じJANの別商品名ベースを同時レビューしても上書きしない（フォルダ名をキーに含める）
        man_path = unique_mix_manifest_path(jan, tn, cp)
        targets.append((jan, man_path, tn, cp))

    if not targets:
        raise SystemExit("起動対象がありません")

    print("--- mix review launch ---")
    procs = []
    for i, (jan, man_path, tn, cp) in enumerate(targets):
        port = int(args.base_port) + i
        man = build_mix_manifest(
            jan=jan,
            theme_new_dir=tn,
            compose_dir=cp,
            provider=str(args.provider),
        )
        write_mix_manifest(man, man_path)
        n_tn = sum(1 for c in man["cards"] if c["source"] == "theme_new")
        n_cp = sum(1 for c in man["cards"] if c["source"] == "competitor")
        print(
            f"  [{i+1}] JAN={jan} port={port} 独自={n_tn} 競合={n_cp}\n"
            f"       theme-new={tn}\n"
            f"       compose={cp}\n"
            f"       http://127.0.0.1:{port}/"
        )
        if args.dry_run:
            continue
        cmd = [
            sys.executable,
            str(MIX_LOOP),
            "--manifest",
            str(man_path),
            "--port",
            str(port),
        ]
        if args.no_browser:
            cmd.append("--no-browser")
        kwargs = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP  # type: ignore[attr-defined]
        procs.append(subprocess.Popen(cmd, cwd=str(SCRIPT_DIR), **kwargs))
        time.sleep(0.6)

    if args.dry_run:
        return 0
    print(f"started {len(procs)} mix review(s). 各タブで採用→出口へ")
    if args.no_wait:
        return 0
    try:
        for p in procs:
            p.wait()
    except KeyboardInterrupt:
        for p in procs:
            p.terminate()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
