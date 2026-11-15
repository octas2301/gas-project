# -*- coding: utf-8 -*-
"""
サブ画像 review_loop 一括起動。

マスタ出品CKの親 JAN → 各JANの最新 compose-dir → ポート連番でタブを開く。

例:
  python sub_image_review_launch.py --from-sheets
  python sub_image_review_launch.py --jan 4906283045119 --jan 4906283045317
  python sub_image_review_launch.py --from-sheets --dry-run
"""
from __future__ import annotations

import argparse
import logging
import subprocess
import sys
import time
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from master_sets import load_checked_parent_jans_from_rows
from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT

LOG = logging.getLogger("set_main_image.sub_image_review_launch")

SCRIPT_DIR = Path(__file__).resolve().parent
REVIEW_LOOP = SCRIPT_DIR / "sub_image_review_loop.py"
COMPOSE_REL = Path(SUB_TEST_OUT) / "sub_image_ai_compose"


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )


def compose_root(work_root: Path) -> Path:
    return Path(work_root) / COMPOSE_REL


def _provider_image_count(compose_dir: Path) -> int:
    """openai/gemini/fal 出力枚数（ステージング除外）。多いほど再利用向き。"""
    n = 0
    for sub in ("03_openai", "02_gemini", "04_fal"):
        folder = Path(compose_dir) / sub
        if not folder.is_dir():
            continue
        for p in folder.glob("*.jpg"):
            if p.name.startswith("__regen_staging_"):
                continue
            n += 1
    return n


def latest_compose_dir(root: Path, jan: str) -> Optional[Path]:
    """JAN配下の compose-dir。優先: フォルダ名（日時）が新しい → 同名なら枚数が多い。"""
    want = (jan or "").strip()
    if not want or not root.is_dir():
        return None
    prefix = f"{want}_"
    cands = [p for p in root.iterdir() if p.is_dir() and p.name.startswith(prefix)]
    if not cands:
        return None
    return max(
        cands,
        key=lambda p: (p.name, _provider_image_count(p)),
    )


def resolve_jans_from_sheets(
    *,
    config: Optional[Path] = None,
) -> List[Tuple[str, str]]:
    from sheets_master import fetch_master_rows

    rows, info = fetch_master_rows(config_path=config)
    LOG.info(
        "sheets master %s rows=%s",
        info.get("masterSheet"),
        info.get("rowCount"),
    )
    return load_checked_parent_jans_from_rows(rows)


def resolve_jans_from_csv(master_csv: Path) -> List[Tuple[str, str]]:
    from master_sets import _read_table

    rows, _h, _i = _read_table(Path(master_csv))
    return load_checked_parent_jans_from_rows(rows)


def build_targets(
    jan_parent: Sequence[Tuple[str, str]],
    root: Path,
) -> Tuple[List[Tuple[str, str, Path]], List[str]]:
    """(jan, parent_sku, compose_dir) と欠落JAN一覧。"""
    ok: List[Tuple[str, str, Path]] = []
    missing: List[str] = []
    for jan, parent in jan_parent:
        d = latest_compose_dir(root, jan)
        if d is None:
            missing.append(jan)
            LOG.warning("compose なし JAN=%s parent=%s root=%s", jan, parent, root)
            continue
        ok.append((jan, parent, d))
        LOG.info("JAN=%s parent=%s → %s", jan, parent, d.name)
    return ok, missing


def launch_one(
    compose_dir: Path,
    port: int,
    provider: str,
    *,
    no_browser: bool,
) -> subprocess.Popen:
    cmd = [
        sys.executable,
        str(REVIEW_LOOP),
        "--compose-dir",
        str(compose_dir),
        "--port",
        str(port),
        "--provider",
        provider,
    ]
    if no_browser:
        cmd.append("--no-browser")
    LOG.info("spawn: %s", " ".join(cmd))
    kwargs = {}
    if sys.platform == "win32":
        # 親ターミナルと分離（Ctrl+C で一括停止は --wait 側）
        kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP  # type: ignore[attr-defined]
    return subprocess.Popen(cmd, cwd=str(SCRIPT_DIR), **kwargs)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="出品CKレ点親の最新 compose で review_loop をポート連番起動"
    )
    src = ap.add_mutually_exclusive_group()
    src.add_argument(
        "--from-sheets",
        action="store_true",
        help="マスタ（Sheets）の出品CKレ点親 JAN を使う",
    )
    src.add_argument(
        "--master-csv",
        type=Path,
        help="ローカルマスタCSVから出品CKレ点親 JAN",
    )
    ap.add_argument(
        "--jan",
        action="append",
        default=[],
        help="JAN を明示（複数可）。Sheets/CSV と併用時は明示分を優先して先頭に",
    )
    ap.add_argument(
        "--work-root",
        type=Path,
        default=DEFAULT_WORK_ROOT,
        help="05.画像生成（セットMAIN）",
    )
    ap.add_argument(
        "--compose-root",
        type=Path,
        default=None,
        help="省略時は work-root/00.テスト出力/sub_image_ai_compose",
    )
    ap.add_argument("--base-port", type=int, default=8765)
    ap.add_argument("--provider", default="openai", choices=("openai", "gemini", "fal"))
    ap.add_argument("--config", type=Path, default=None, help="C1 config（--from-sheets）")
    ap.add_argument("--dry-run", action="store_true", help="起動せず一覧のみ")
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument(
        "--no-wait",
        action="store_true",
        help="起動後にすぐ終了（子プロセスは残す）",
    )
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    if not REVIEW_LOOP.is_file():
        raise SystemExit(f"review_loop がありません: {REVIEW_LOOP}")

    pairs: List[Tuple[str, str]] = []
    if args.from_sheets:
        pairs = resolve_jans_from_sheets(config=args.config)
    elif args.master_csv:
        pairs = resolve_jans_from_csv(Path(args.master_csv))

    explicit = [str(j).strip() for j in (args.jan or []) if str(j).strip()]
    if explicit:
        # 明示 JAN を先頭（parent 不明は空）
        seen = set()
        merged: List[Tuple[str, str]] = []
        for j in explicit:
            if j in seen:
                continue
            seen.add(j)
            parent = next((p for jj, p in pairs if jj == j), "")
            merged.append((j, parent or "(explicit)"))
        for j, p in pairs:
            if j in seen:
                continue
            seen.add(j)
            merged.append((j, p))
        pairs = merged

    if not pairs:
        raise SystemExit(
            "対象JANがありません。--from-sheets / --master-csv / --jan のいずれかを指定してください。"
        )

    root = Path(args.compose_root) if args.compose_root else compose_root(Path(args.work_root))
    if not root.is_dir():
        raise SystemExit(f"compose-root がありません: {root}")

    targets, missing = build_targets(pairs, root)
    if missing:
        LOG.warning("compose 欠落 %d 件: %s", len(missing), ",".join(missing))
    if not targets:
        raise SystemExit("起動できる compose-dir が1件もありません")

    print("--- review launch plan ---")
    for i, (jan, parent, d) in enumerate(targets):
        port = int(args.base_port) + i
        print(f"  [{i + 1}] JAN={jan} parent={parent} port={port}")
        print(f"       {d}")
        print(f"       http://127.0.0.1:{port}/")
    if missing:
        print(f"  (skip no-compose) {', '.join(missing)}")

    if args.dry_run:
        print("dry-run: not started")
        return 0

    procs: List[Tuple[subprocess.Popen, int, str]] = []
    for i, (jan, _parent, d) in enumerate(targets):
        port = int(args.base_port) + i
        # ブラウザ同時オープンを少しずらす
        if i and not args.no_browser:
            time.sleep(0.6)
        proc = launch_one(d, port, str(args.provider), no_browser=bool(args.no_browser))
        procs.append((proc, port, jan))

    print(f"started {len(procs)} review_loop(s). 各タブでレビュー →「修正なしで終了」")
    if args.no_wait:
        return 0

    print("Ctrl+C で全プロセスを停止します")
    try:
        while True:
            alive = [(p, port, jan) for p, port, jan in procs if p.poll() is None]
            if not alive:
                break
            time.sleep(0.8)
    except KeyboardInterrupt:
        print("stopping…")
        for p, port, jan in procs:
            if p.poll() is None:
                LOG.info("terminate JAN=%s port=%s pid=%s", jan, port, p.pid)
                p.terminate()
        for p, _port, _jan in procs:
            try:
                p.wait(timeout=5)
            except Exception:
                p.kill()
    return 0


if __name__ == "__main__":
    sys.exit(main())
