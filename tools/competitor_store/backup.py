# -*- coding: utf-8 -*-
"""Keep the latest local backup only. Does not write listing master. Drive copy is dry-run unless --apply."""
from __future__ import annotations

import argparse
import shutil
from datetime import datetime, timezone
from pathlib import Path

from store import LocalStore

PREFIX = "退避_"
DRIVE_PREFIX = "競合ストア退避_"


def drive_trash_plan(files: list[dict], live_id: str, new_id: str, protect_ids: list[str] | None = None) -> list[str]:
    """files: {id, name}. Never trash the live workbook or protected ids."""
    protect = {str(live_id or ""), str(new_id or "")}
    for p in protect_ids or []:
        if p:
            protect.add(str(p))
    trash = []
    for f in files:
        fid = str(f.get("id") or "")
        name = str(f.get("name") or "")
        if not fid or fid in protect:
            continue
        if name.startswith(DRIVE_PREFIX):
            trash.append(fid)
    return trash


def backup_local(root: Path | None = None, apply: bool = False) -> dict:
    store = LocalStore(root)
    store.ensure_schema()
    base = store.root
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    dest = base / "backups" / (PREFIX + stamp)
    existing = sorted((base / "backups").glob(PREFIX + "*")) if (base / "backups").exists() else []
    to_delete = existing[:]
    report = {"dest": str(dest), "would_delete": [str(p) for p in to_delete], "apply": apply}
    if not apply:
        return report
    dest.mkdir(parents=True, exist_ok=True)
    for p in base.glob("*.csv"):
        shutil.copy2(p, dest / p.name)
    for p in to_delete:
        if p.is_dir() and p.resolve() != dest.resolve():
            shutil.rmtree(p)
    return report


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    print(backup_local(apply=args.apply))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
