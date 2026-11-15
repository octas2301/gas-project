# -*- coding: utf-8 -*-
"""共通スプシ登録表。IDは全プロジェクト固定。正本は docs/org/COMMON_SS_VIEW_REQUIREMENTS.md"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REGISTRY_PATH = ROOT / "registry.json"
OUT_DIR = ROOT / "out"


def load_registry() -> dict:
    return json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))


def save_registry(reg: dict) -> None:
    REGISTRY_PATH.write_text(
        json.dumps(reg, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def file_keys(reg: dict | None = None) -> dict[str, str]:
    """省略名/別名 → 省略名。"""
    reg = reg or load_registry()
    out = {}
    for key, rec in reg["files"].items():
        out[key] = key
        out[str(rec.get("id") or "")] = key
        for a in rec.get("aliases") or []:
            out[str(a)] = key
    return out


def resolve_file(name: str, reg: dict | None = None) -> tuple[str, str]:
    """(省略名, spreadsheetId)。"""
    reg = reg or load_registry()
    key = file_keys(reg).get(name) or file_keys(reg).get(name.strip())
    if not key or key not in reg["files"]:
        raise KeyError("未登録のファイル: %s" % name)
    return key, str(reg["files"][key]["id"])


def file_id(name: str, reg: dict | None = None) -> str:
    return resolve_file(name, reg)[1]


def resolve_menu(name: str, reg: dict | None = None) -> str:
    reg = reg or load_registry()
    if name in reg["menus"]:
        return name
    alias = (reg.get("menu_aliases") or {}).get(name)
    if alias:
        return alias
    raise KeyError("未登録のメニュー: %s" % name)


def resolve_intent(name: str, reg: dict | None = None) -> str:
    reg = reg or load_registry()
    if name in reg["intents"]:
        return name
    alias = (reg.get("intent_aliases") or {}).get(name)
    if alias:
        return alias
    return "judge"


def files_for_menu(menu: str, reg: dict | None = None) -> list[str]:
    reg = reg or load_registry()
    menu = resolve_menu(menu, reg)
    return list(reg["menus"][menu])


def tabs_for_intent(intent: str, file_key: str, reg: dict | None = None) -> list[str]:
    reg = reg or load_registry()
    intent = resolve_intent(intent, reg)
    by_file = (reg.get("intents") or {}).get(intent) or {}
    return list(by_file.get(file_key) or [])
