# -*- coding: utf-8 -*-
"""ローカル Drive Desktop パス解決（B-T0 と同型）。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Optional

HERE = Path(__file__).resolve().parent


def load_config(path: Optional[Path] = None) -> Dict[str, Any]:
    cfg_path = path or (HERE / "config.local.json")
    if not cfg_path.is_file():
        cfg_path = HERE / "config.example.json"
    return json.loads(cfg_path.read_text(encoding="utf-8"))


def folder_path(cfg: Dict[str, Any], which: str) -> Path:
    root = Path(str(cfg.get("local_root") or "")).expanduser()
    key = {"01": "local_01", "02": "local_02", "03": "local_03"}[which]
    name = str(cfg.get(key) or "")
    p = root / name
    if not p.is_dir():
        alt = Path("G:/") / "マイドライブ" / "06.販促（タイムセール・クーポンなど）" / name
        if alt.is_dir():
            return alt
    return p


def latest_xlsx(folder: Path) -> Optional[Path]:
    if not folder.is_dir():
        return None
    files = [
        p
        for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$")
    ]
    if not files:
        return None
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0]
