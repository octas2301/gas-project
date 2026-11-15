# -*- coding: utf-8 -*-
"""
B-⑤ review_loop 一括起動（B-④ sub_image_review_launch と同型）。

  python b5_main_base_review_launch.py --fresh --prefix ishi-4906283045119-oya3
  python b5_main_base_review_launch.py --fresh --from-root-latest --no-wait
  python b5_main_base_review_launch.py --compose-dir "...\\かつお_…" --compose-dir "...\\まぐろ_…"
  python b5_main_base_review_launch.py --list-latest   # 既存ラン一覧（再利用判断用）

--fresh: 先に当該ポートの旧 review_loop を止めてから、今の .py で起動（古い画面の再利用禁止）。
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import List, Optional, Sequence

from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT

LOG = logging.getLogger("set_main_image.b5_main_base_review_launch")
SCRIPT_DIR = Path(__file__).resolve().parent
REVIEW_LOOP = SCRIPT_DIR / "b5_main_base_review_loop.py"
COMPOSE_REL = Path(SUB_TEST_OUT) / "b5_main_base_compose"
_TS = re.compile(r"^(.+)_(\d{8})_(\d{6})$")
_JAN_DIR = re.compile(r"^\d{13}_\d{8}_\d{6}$")


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )


def compose_root(work_root: Path) -> Path:
    return Path(work_root) / COMPOSE_REL


def _compare_count(compose_dir: Path) -> int:
    pick = Path(compose_dir) / "比較"
    if not pick.is_dir():
        return 0
    n = 0
    for p in list(pick.glob("*.jpg")) + list(pick.glob("*.png")):
        if p.name.startswith("__regen_staging_"):
            continue
        n += 1
    return n


def latest_per_short_name(root: Path, *, include_jan_dirs: bool = False) -> List[Path]:
    """省略名ごとに最良の compose dir。

    優先: 比較枚数が多い方 → 同時刻なら新しいタイムスタンプ。
    （再生成の副作用でできた1枚フォルダより、元の4枚フォルダを選ぶ）
    """
    if not root.is_dir():
        return []
    best: dict[str, Path] = {}
    for p in root.iterdir():
        if not p.is_dir() or p.name.startswith("_"):
            continue
        if not (p / "比較").is_dir():
            continue
        if not include_jan_dirs and _JAN_DIR.match(p.name):
            continue
        m = _TS.match(p.name)
        key = m.group(1) if m else p.name
        ts = f"{m.group(2)}_{m.group(3)}" if m else p.name
        n = _compare_count(p)
        prev = best.get(key)
        if prev is None:
            best[key] = p
            continue
        pm = _TS.match(prev.name)
        prev_ts = f"{pm.group(2)}_{pm.group(3)}" if pm else prev.name
        prev_n = _compare_count(prev)
        # 新しいタイムスタンプ優先（再生成で枚数が増えた旧フォルダに負けない）
        if ts > prev_ts or (ts == prev_ts and n > prev_n):
            best[key] = p
    return [best[k] for k in sorted(best.keys())]


def list_latest_info(root: Path) -> List[dict]:
    out = []
    for d in latest_per_short_name(root):
        pick = d / "比較"
        n = len(list(pick.glob("*.jpg"))) + len(list(pick.glob("*.png"))) if pick.is_dir() else 0
        meta = {}
        bp = d / "b5_proposals.json"
        if bp.is_file():
            try:
                meta = json.loads(bp.read_text(encoding="utf-8-sig"))
            except Exception:
                meta = {}
        themes = [
            str(x.get("themeId"))
            for x in (meta.get("proposals") or [])
            if isinstance(x, dict) and x.get("themeId")
        ]
        m = _TS.match(d.name)
        out.append(
            {
                "shortName": meta.get("shortName") or (m.group(1) if m else d.name),
                "parentSku": meta.get("parentSku") or "",
                "jan": meta.get("jan") or "",
                "composeDir": str(d),
                "name": d.name,
                "compareN": n,
                "themeIds": themes,
                "reuseHint": (
                    f'既存を使う: python b5_main_base_review_launch.py --compose-dir "{d}"'
                ),
            }
        )
    return out


def launch_one(compose_dir: Path, port: int) -> subprocess.Popen:
    cmd = [
        sys.executable,
        str(REVIEW_LOOP),
        "--compose-dir",
        str(compose_dir),
        "--port",
        str(port),
    ]
    LOG.info("launch port=%s %s", port, compose_dir.name)
    return subprocess.Popen(cmd, cwd=str(SCRIPT_DIR))


def latest_matching_prefix(root: Path, prefix: str) -> Optional[Path]:
    """親SKU／省略名プレフィックスに一致する最新タイムスタンプの compose-dir。"""
    key = str(prefix or "").strip()
    if not key or not root.is_dir():
        return None
    cands = [
        p
        for p in root.iterdir()
        if p.is_dir() and not p.name.startswith("_") and p.name.startswith(key) and (p / "比較").is_dir()
    ]
    if not cands:
        return None
    cands.sort(key=lambda p: p.name, reverse=True)
    return cands[0]


def kill_review_on_ports(ports: Sequence[int]) -> List[int]:
    """当該ポートの Listen と b5_main_base_review_loop を止める。戻値=止めた PID。"""
    killed: List[int] = []
    if os.name != "nt":
        for port in ports:
            try:
                out = subprocess.check_output(
                    ["bash", "-lc", f"lsof -ti tcp:{int(port)} || true"],
                    text=True,
                )
                for line in out.splitlines():
                    pid = int(line.strip())
                    os.kill(pid, 9)
                    killed.append(pid)
            except Exception as e:
                LOG.warning("kill port %s failed: %s", port, e)
        return killed

    # Windows: ポート Listen の PID + review_loop コマンドライン
    ps = r"""
$ports = @(%PORTS%);
$pids = New-Object 'System.Collections.Generic.HashSet[int]';
foreach ($port in $ports) {
  Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
    ForEach-Object { [void]$pids.Add([int]$_.OwningProcess) }
}
Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -and ($_.CommandLine -like '*b5_main_base_review_loop*') } |
  ForEach-Object { [void]$pids.Add([int]$_.ProcessId) }
$pids | ForEach-Object { $_ }
"""
    port_csv = ",".join(str(int(p)) for p in ports)
    script = ps.replace("%PORTS%", port_csv)
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", script],
            text=True,
            stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        LOG.warning("kill list failed: %s", e)
        return killed
    for line in out.splitlines():
        line = line.strip()
        if not line.isdigit():
            continue
        pid = int(line)
        if pid <= 0:
            continue
        try:
            subprocess.check_call(
                ["powershell", "-NoProfile", "-Command", f"Stop-Process -Id {pid} -Force -ErrorAction SilentlyContinue"],
                stderr=subprocess.DEVNULL,
            )
            killed.append(pid)
            LOG.info("killed pid=%s", pid)
        except Exception:
            pass
    if killed:
        time.sleep(1.0)
    return killed


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="B-⑤ review_loop 一括起動")
    ap.add_argument("--compose-dir", type=Path, action="append", default=[])
    ap.add_argument(
        "--from-root-latest",
        action="store_true",
        help="b5_main_base_compose 配下で省略名ごとの最新ランを開く",
    )
    ap.add_argument(
        "--prefix",
        action="append",
        default=[],
        help="親SKU等の接頭辞で最新 compose-dir を1つ選ぶ（複数可）。例: --prefix ishi-…-oya3",
    )
    ap.add_argument(
        "--list-latest",
        action="store_true",
        help="最新ラン一覧を出して終了（既存再利用の判断用。compose はしない）",
    )
    ap.add_argument(
        "--include-jan-dirs",
        action="store_true",
        help="JAN_時刻 名の旧ランも --from-root-latest に含める",
    )
    ap.add_argument(
        "--fresh",
        action="store_true",
        help="起動前に base-port 帯の旧 review を止める（コード変更後は必須）",
    )
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--base-port", type=int, default=8865)
    ap.add_argument("--no-wait", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(list(argv) if argv is not None else None)
    _setup_log(args.verbose)

    root = compose_root(Path(args.work_root or DEFAULT_WORK_ROOT))
    if args.list_latest:
        info = list_latest_info(root)
        print(json.dumps({"ok": True, "root": str(root), "runs": info}, ensure_ascii=False, indent=2))
        print(
            "\n【再利用】Agent は b5_reuse_decide.py で自動判定（三択を人に聞かない）。"
            "reuse=composeせず --fresh --compose-dir／missing_only=足りないキャッチだけ足す／"
            "full=通常 compose。全部作り直しは人の明示か --full。"
            "開き直しは --fresh（旧プロセス禁止）。"
        )
        return 0

    dirs: List[Path] = [Path(p) for p in (args.compose_dir or [])]
    for pref in args.prefix or []:
        hit = latest_matching_prefix(root, str(pref))
        if hit is None:
            LOG.warning("prefix miss %s", pref)
            continue
        LOG.info("prefix %s -> %s", pref, hit.name)
        dirs.append(hit)
    if args.from_root_latest:
        dirs.extend(latest_per_short_name(root, include_jan_dirs=bool(args.include_jan_dirs)))
    seen = set()
    uniq: List[Path] = []
    for d in dirs:
        key = str(d.resolve()) if d.exists() else str(d)
        if key in seen:
            continue
        seen.add(key)
        uniq.append(d)

    if not uniq:
        raise SystemExit(
            "起動する compose-dir がありません（--compose-dir / --prefix / --from-root-latest）"
        )

    ports = [int(args.base_port) + i for i in range(len(uniq))]
    if args.fresh:
        killed = kill_review_on_ports(ports)
        LOG.info("fresh kill n=%s ports=%s", len(killed), ports)

    procs = []
    for i, d in enumerate(uniq):
        if not d.is_dir():
            LOG.warning("skip missing %s", d)
            continue
        port = int(args.base_port) + i
        procs.append(launch_one(d, port))
        print(f"URL http://127.0.0.1:{port}/  dir={d.name}")
        time.sleep(0.5)

    if not procs:
        raise SystemExit("1つも起動できませんでした")

    print(f"started {len(procs)} review UI(s). base-port={args.base_port} fresh={bool(args.fresh)}")
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
