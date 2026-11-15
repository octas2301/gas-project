# -*- coding: utf-8 -*-
"""
食材見本ゲート一括起動。

例:
  python sub_image_food_canon_launch.py --jan 4906283045119 --jan 4906283045317 --force-pick --no-wait
"""
from __future__ import annotations

import argparse
import logging
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import List, Optional, Sequence

from sub_image_food_canon_gate import SCRIPT_DIR, prepare_gate
from work_paths import DEFAULT_WORK_ROOT

LOG = logging.getLogger("set_main_image.food_canon_launch")
GATE = SCRIPT_DIR / "sub_image_food_canon_gate.py"


def kill_food_canon_on_ports(ports: Sequence[int]) -> List[int]:
    """当該ポート Listen と food_canon_gate を止める（二重 Listen で別商品が開くのを防ぐ）。"""
    killed: List[int] = []
    ports = [int(p) for p in ports if int(p) > 0]
    if not ports:
        return killed
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
    # 対象ポートの Listen だけ止める（他親のゲートまで落とさない）
    ps = r"""
$ports = @(%PORTS%);
$pids = New-Object 'System.Collections.Generic.HashSet[int]';
foreach ($port in $ports) {
  Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
    ForEach-Object { [void]$pids.Add([int]$_.OwningProcess) }
}
$pids | ForEach-Object { $_ }
"""
    script = ps.replace("%PORTS%", ",".join(str(p) for p in ports))
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
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    f"Stop-Process -Id {pid} -Force -ErrorAction SilentlyContinue",
                ],
                stderr=subprocess.DEVNULL,
            )
            killed.append(pid)
            LOG.info("killed pid=%s", pid)
        except Exception:
            pass
    if killed:
        time.sleep(1.0)
        LOG.info("fresh kill n=%s ports=%s", len(killed), list(ports))
    return killed


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="食材見本ゲート一括起動")
    ap.add_argument("--jan", action="append", default=[])
    ap.add_argument("--name-base", action="append", default=[], dest="name_base")
    ap.add_argument("--parent-sku", action="append", default=[], dest="parent_sku")
    ap.add_argument("--part-name", action="append", default=[], dest="part_name")
    ap.add_argument("--part-sku", action="append", default=[], dest="part_sku")
    ap.add_argument("--part-jan", action="append", default=[], dest="part_jan")
    ap.add_argument("--part-bag", action="append", default=[], dest="part_bag")
    ap.add_argument("--work-root", type=Path, default=DEFAULT_WORK_ROOT)
    ap.add_argument("--base-port", type=int, default=8790)
    ap.add_argument("--refresh-candidates", action="store_true")
    ap.add_argument("--force-pick", action="store_true")
    ap.add_argument(
        "--context",
        default="",
        help="main = B-④ MAIN用（下ボタン2つ: 変更なしで閉じる / 選び直したらここで保存）",
    )
    ap.add_argument(
        "--no-fresh",
        action="store_true",
        help="起動前に対象ポートの旧ゲートを止めない（既定は止める）",
    )
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--no-wait", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    jans = [str(j).strip() for j in (args.jan or []) if str(j).strip()]
    part_names = [str(x).strip() for x in (args.part_name or []) if str(x).strip()]
    # 親SKU指定かつ part-name 無し → マスタセット構成から組み合わせへ自動切替
    if len(part_names) < 2 and args.parent_sku:
        try:
            from food_unit_canon import combo_parts_from_sheet

            sku0 = str((args.parent_sku[0] if args.parent_sku else "") or "").strip()
            if sku0:
                bom = combo_parts_from_sheet(parent_sku=sku0, work_root=Path(args.work_root))
                if len(bom) >= 2:
                    part_names = [str(p.get("nameBase") or p.get("jan") or "").strip() for p in bom]
                    part_names = [x for x in part_names if x]
                    args.part_sku = [str(p.get("parentSku") or "").strip() for p in bom]
                    args.part_jan = [str(p.get("jan") or "").strip() for p in bom]
                    args.part_bag = [str(p.get("packageTruth") or "").strip() for p in bom]
                    if not args.name_base:
                        args.name_base = [""]
                    LOG.info(
                        "launch mix from sheet BOM parent=%s parts=%s",
                        sku0,
                        part_names,
                    )
        except Exception as e:
            LOG.warning("sheet BOM auto-mix fail: %s", e)
    if len(part_names) >= 2:
        print("--- food canon mix confirm ---")
        port = int(args.base_port)
        nb = (args.name_base[0] if args.name_base else "") or ""
        sku = (args.parent_sku[0] if args.parent_sku else "") or "mix"
        print(f"  mix nameBase={nb} parentSku={sku} port={port} http://127.0.0.1:{port}/")
        if args.dry_run:
            return 0
        if not args.no_fresh:
            kill_food_canon_on_ports([port])
        cmd = [
            sys.executable,
            str(GATE),
            "--port",
            str(port),
            "--parent-sku",
            str(sku).strip(),
        ]
        if str(nb).strip():
            cmd.extend(["--name-base", str(nb).strip()])
        for i, pn in enumerate(part_names):
            cmd.extend(["--part-name", pn])
            if i < len(args.part_sku) and str(args.part_sku[i] or "").strip():
                cmd.extend(["--part-sku", str(args.part_sku[i]).strip()])
            if i < len(args.part_jan) and str(args.part_jan[i] or "").strip():
                cmd.extend(["--part-jan", str(args.part_jan[i]).strip()])
            if i < len(args.part_bag) and str(args.part_bag[i] or "").strip():
                cmd.extend(["--part-bag", str(args.part_bag[i]).strip()])
        if args.no_browser:
            cmd.append("--no-browser")
        if str(args.context or "").strip():
            cmd.extend(["--context", str(args.context).strip()])
        kwargs = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP  # type: ignore[attr-defined]
        proc = subprocess.Popen(cmd, cwd=str(SCRIPT_DIR), **kwargs)
        print("started mix confirm. MAIN用: 変更なしで閉じる → B-④「次へ」")
        if args.no_wait:
            return 0
        try:
            proc.wait()
        except KeyboardInterrupt:
            proc.terminate()
        return 0
    if not jans:
        raise SystemExit("--jan を1つ以上（組み合わせは --part-name を2つ以上）")

    print("--- food canon gate launch ---")
    ports = [int(args.base_port) + i for i in range(len(jans))]
    if not args.dry_run and not args.no_fresh:
        kill_food_canon_on_ports(ports)
    procs = []
    for i, jan in enumerate(jans):
        port = ports[i]
        tn = SCRIPT_DIR / "_out" / f"theme_new_run_{jan}"
        prepare_gate(
            jan=jan,
            theme_new_dir=tn,
            work_root=Path(args.work_root),
            refresh_candidates=bool(args.refresh_candidates),
        )
        if args.force_pick:
            from sub_image_food_canon_gate import _load_session, _save_session

            s = _load_session(tn)
            s["phase"] = "pick"
            if str(args.context or "").strip():
                s["context"] = str(args.context).strip()
            nb0 = (args.name_base[i] if i < len(args.name_base) else "") or ""
            sku0 = (args.parent_sku[i] if i < len(args.parent_sku) else "") or ""
            if str(nb0).strip():
                s["nameBase"] = str(nb0).strip()
            if str(sku0).strip():
                s["parentSku"] = str(sku0).strip()
            _save_session(tn, s)
        print(f"  [{i+1}] JAN={jan} port={port} http://127.0.0.1:{port}/")
        if args.dry_run:
            continue
        cmd = [
            sys.executable,
            str(GATE),
            "--jan",
            jan,
            "--theme-new-dir",
            str(tn),
            "--port",
            str(port),
        ]
        nb = (args.name_base[i] if i < len(args.name_base) else "") or ""
        sku = (args.parent_sku[i] if i < len(args.parent_sku) else "") or ""
        if str(nb).strip():
            cmd.extend(["--name-base", str(nb).strip()])
        if str(sku).strip():
            cmd.extend(["--parent-sku", str(sku).strip()])
        if args.no_browser:
            cmd.append("--no-browser")
        if args.force_pick:
            cmd.append("--force-pick")
        if str(args.context or "").strip():
            cmd.extend(["--context", str(args.context).strip()])
        if args.refresh_candidates:
            cmd.append("--refresh-candidates")
        kwargs = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP  # type: ignore[attr-defined]
        procs.append(subprocess.Popen(cmd, cwd=str(SCRIPT_DIR), **kwargs))
        time.sleep(0.5)

    if args.dry_run:
        return 0
    print(f"started {len(procs)} gate(s).", end="")
    if str(args.context or "").strip().lower() == "main":
        print(" MAIN用: 変更なしで閉じる / 選び直したらここで保存 → B-④「次へ」")
    else:
        print(" 見本レ点→実行→同画面レビュー")
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
