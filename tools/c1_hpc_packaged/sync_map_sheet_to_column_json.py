#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
▼設定(Amazonマッピング) MAP → *_column_map.json（実行用）。

正本: docs/org/LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md
  正本=sheet／派生=MD／実行=JSON（sheetから生成）
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_RW = SCRIPT_DIR / "secrets" / "token_sheets_rw.json"
SHEET_TITLE = "▼設定(Amazonマッピング)"

PROFILES = {
    "grocery": {
        "json": "food_fish_grocery_column_map.json",
        "sheet_product_types": {"FOOD", "GROCERY"},
    },
    "seasoning": {
        "json": "food_seasoning_column_map.json",
        "sheet_product_types": {"FOOD", "SEASONING"},
    },
}

# sheet transform（日本語）→ c1_quantity_policy
TRANSFORM_POLICY = {
    "セット数を数値化": ("parse_set_count", True),
    "PARSE_SET_COUNT": ("parse_set_count", True),
    "ユニット数＝セット数": ("use_set_count_for_unit_count", True),
    "USE_SET_COUNT": ("use_set_count_for_unit_count", True),
    "サイズ名から重量を取る": ("parse_weight_from_size", True),
    "PARSE_WEIGHT_FROM_SIZE": ("parse_weight_from_size", True),
}

INHERIT_CHILD_ONLY = {"子のみ", "CHILD_ONLY"}


def load_cfg() -> dict:
    for name in ("config.local.json", "config.json"):
        p = SCRIPT_DIR / name
        if p.is_file():
            return json.loads(p.read_text(encoding="utf-8"))
    raise FileNotFoundError("config.local.json がありません")


def get_creds() -> Credentials:
    cred_path = SCRIPT_DIR / "secrets" / "credentials.json"
    creds = None
    if TOKEN_RW.is_file():
        creds = Credentials.from_authorized_user_file(str(TOKEN_RW), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_RW.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_RW.write_text(creds.to_json(), encoding="utf-8")
    return creds


def _cell(row: List[Any], idx: Dict[str, int], key: str) -> str:
    i = idx.get(key)
    if i is None or i >= len(row):
        return ""
    return str(row[i] if row[i] is not None else "").strip()


def _truthy(s: str) -> bool:
    return s.upper() in ("TRUE", "YES", "1", "有効")


def _split_aliases(raw: str) -> List[str]:
    parts = [p.strip() for p in (raw or "").replace("\n", "|").split("|")]
    return [p for p in parts if p]


def fetch_map_rows(svc: Any, spreadsheet_id: str) -> Tuple[Dict[str, int], List[Dict[str, str]]]:
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range="'%s'!A1:Q400" % SHEET_TITLE)
        .execute()
        .get("values")
        or []
    )
    header_i = None
    for i, r in enumerate(vals):
        if r and str(r[0]).strip() == "productCategory" and len(r) > 2 and str(r[2]).strip() == "attrKey":
            header_i = i
            break
    if header_i is None:
        raise RuntimeError("MAP ヘッダ行（productCategory/attrKey）が見つかりません")

    headers = [str(h).strip() for h in vals[header_i]]
    idx = {h: i for i, h in enumerate(headers)}
    rows: List[Dict[str, str]] = []
    # skip JP label row + description row after English header
    for r in vals[header_i + 1 :]:
        if not r or len(r) < 3:
            continue
        ak = _cell(r, idx, "attrKey")
        if not ak or ak in ("内部キー", "コード・jsonが使う安定名"):
            continue
        # skip pure Japanese header duplicates
        if ak.startswith("SCの") or "安定名" in ak:
            continue
        pt = _cell(r, idx, "productType")
        if not pt or pt in ("商品タイプ（技術値）",):
            continue
        rows.append(
            {
                "productType": pt,
                "attrKey": ak,
                "scHeaderJa": _cell(r, idx, "scHeaderJa"),
                "scHeaderAlias": _cell(r, idx, "scHeaderAlias"),
                "required": _cell(r, idx, "required"),
                "masterColPrimary": _cell(r, idx, "masterColPrimary"),
                "masterColFallback": _cell(r, idx, "masterColFallback"),
                "valueSource": _cell(r, idx, "valueSource"),
                "inherit": _cell(r, idx, "inherit"),
                "transform": _cell(r, idx, "transform"),
                "defaultValue": _cell(r, idx, "defaultValue"),
                "doNotUse": _cell(r, idx, "doNotUse"),
                "sourceNote": _cell(r, idx, "sourceNote"),
                "notes": _cell(r, idx, "notes"),
                "enabled": _cell(r, idx, "enabled") or "TRUE",
            }
        )
    return idx, rows


def apply_rows_to_colmap(
    colmap: dict,
    rows: List[Dict[str, str]],
    sheet_pts: set,
) -> Dict[str, Any]:
    """Merge sheet MAP into column_map. Returns diff summary."""
    diff: Dict[str, Any] = {
        "defaults": {},
        "master_columns": {},
        "policy": {},
        "aliases_ja": {},
        "skipped": [],
    }
    defaults = dict(colmap.get("defaults") or {})
    master_columns = dict(colmap.get("master_columns") or {})
    aliases = dict(colmap.get("xlsm_header_aliases") or {})
    policy = dict(colmap.get("c1_quantity_policy") or {})

    # quantity policy defaults for grocery-style transforms
    policy.setdefault("use_set_count_for_number_items", False)
    policy.setdefault("parse_set_count", False)
    policy.setdefault("use_set_count_for_unit_count", False)
    policy.setdefault("parse_weight_from_size", False)

    for row in rows:
        if row["productType"] not in sheet_pts:
            continue
        if not _truthy(row["enabled"]) and row["enabled"]:
            # REF rows still useful for doNotUse / set_count
            if not row["attrKey"].endswith("_ref"):
                continue

        ak = row["attrKey"]
        if ak.endswith("_ref"):
            # refs: only drive policy / set_count master
            if ak == "set_count_ref" or "セット数" in row["transform"]:
                policy["parse_set_count"] = True
                policy["use_set_count_for_number_items"] = True
                masters = _split_aliases(row["masterColPrimary"]) + _split_aliases(
                    row["masterColFallback"]
                )
                if masters:
                    old = master_columns.get("set_count") or []
                    merged = list(dict.fromkeys(masters + list(old)))
                    if merged != old:
                        diff["master_columns"]["set_count"] = {"from": old, "to": merged}
                    master_columns["set_count"] = merged
            continue

        # transforms → policy
        tr = row["transform"]
        for key, (flag, val) in TRANSFORM_POLICY.items():
            if key in tr:
                if policy.get(flag) != val:
                    diff["policy"][flag] = {"from": policy.get(flag), "to": val}
                policy[flag] = val
                if flag == "parse_set_count":
                    policy["use_set_count_for_number_items"] = True

        # omit heuristics from notes / empty default + OPT
        notes = row["notes"] + " " + row["sourceNote"]
        if ak == "color" and (
            "未出力" in notes or (not row["defaultValue"] and row["required"] in ("OPT", "任意"))
        ):
            if not policy.get("omit_color"):
                diff["policy"]["omit_color"] = {"from": False, "to": True}
            policy["omit_color"] = True
            defaults.pop("color", None)
        if ak == "item_form" and ("未出力" in notes or "黒セル" in notes):
            policy["omit_item_form"] = True
            defaults.pop("item_form", None)
            diff["policy"]["omit_item_form"] = True
        if ak == "temperature_rating" and (
            "未出力" in notes or "黒セル" in notes or "保存方法長文" in notes
        ):
            policy["omit_temperature_rating"] = True
            diff["policy"]["omit_temperature_rating"] = True

        # defaults from FIXED / defaultValue
        vs = row["valueSource"]
        if row["defaultValue"] and (
            vs in ("固定", "FIXED") or tr in ("固定値を使う", "FIXED") or row["defaultValue"]
        ):
            if vs in ("固定", "FIXED") or tr in ("固定値を使う", "FIXED") or ak in (
                "var_theme",
                "action",
                "brand",
                "id_type",
                "unit_uom",
                "item_weight_unit",
                "condition",
                "fulfillment",
                "shipping",
                "import_type",
                "exclusive",
                "heat",
                "hazmat",
                "liquid",
            ):
                if defaults.get(ak) != row["defaultValue"] and ak in (
                    list(defaults.keys())
                    + [
                        "var_theme",
                        "action",
                        "brand",
                        "id_type",
                        "unit_uom",
                        "item_weight_unit",
                        "condition",
                        "shipping",
                        "import_type",
                        "exclusive",
                        "heat",
                        "hazmat",
                        "liquid",
                        "product_type",
                    ]
                ):
                    if ak in defaults or ak in (
                        "var_theme",
                        "unit_uom",
                        "item_weight_unit",
                        "action",
                        "brand",
                        "id_type",
                    ):
                        if defaults.get(ak) != row["defaultValue"]:
                            diff["defaults"][ak] = {
                                "from": defaults.get(ak),
                                "to": row["defaultValue"],
                            }
                        defaults[ak] = row["defaultValue"]

        # master_columns (skip GENERATED-only / FIXED-only without master)
        if vs not in ("GENERATED", "固定", "FIXED") or row["masterColPrimary"]:
            masters: List[str] = []
            masters.extend(_split_aliases(row["masterColPrimary"]))
            masters.extend(_split_aliases(row["masterColFallback"]))
            banned = set(_split_aliases(row["doNotUse"]))
            masters = [m for m in masters if m and m not in banned]
            # strip banned from existing too when doNotUse present
            if ak in master_columns or masters:
                old = list(master_columns.get(ak) or [])
                if banned:
                    old = [m for m in old if m not in banned]
                merged = list(dict.fromkeys(masters + old)) if masters else old
                # for CHILD_ONLY qty fields, prefer sheet masters only
                if row["inherit"] in INHERIT_CHILD_ONLY and masters:
                    merged = list(dict.fromkeys(masters))
                if merged != (master_columns.get(ak) or []):
                    diff["master_columns"][ak] = {
                        "from": master_columns.get(ak),
                        "to": merged,
                    }
                if merged:
                    master_columns[ak] = merged

        # xlsm_header_aliases: SC日本語見出しのみ（マスタ列名・保存方法等は入れない）
        if ak in aliases:
            ja = row["scHeaderJa"]
            cur = list(aliases[ak])
            added = []
            if (
                ja
                and ja not in cur
                and "#" not in ja
                and "[" not in ja
                and not ja.startswith("▼マスタ")
                and "保存方法" not in ja
                and "一人分" not in ja
            ):
                # scHeaderAlias の英テクニカル名は既に aliases にある想定。JPは scHeaderJa のみ追加
                cur.append(ja)
                added.append(ja)
            if added:
                diff["aliases_ja"][ak] = added
                aliases[ak] = cur

        # size: prefer size attribute over package_size_name when notes say so
        if ak == "size" and ("AT" in notes or "size属性" in notes or "package_size_nameではない" in notes):
            aliases["size"] = [
                "size[marketplace_id=A1VC38T7YXB528][language_tag=ja_JP]#1.value",
                "サイズ",
            ]
            if (colmap.get("cols") or {}).get("size") != 46:
                colmap.setdefault("cols", {})["size"] = 46
                diff["aliases_ja"]["size_col"] = 46

    # unit_uom_default from defaults
    if defaults.get("unit_uom"):
        policy["unit_uom_default"] = defaults["unit_uom"]

    if policy.get("parse_set_count") and policy.get("use_set_count_for_unit_count"):
        policy["use_set_count_for_number_items"] = True

    policy["notes"] = (
        "synced from ▼設定(Amazonマッピング) MAP. "
        + (policy.get("notes") or "")
    )[:300]

    colmap["defaults"] = defaults
    colmap["master_columns"] = master_columns
    colmap["xlsm_header_aliases"] = aliases
    if any(policy.get(k) for k in ("parse_set_count", "parse_weight_from_size", "omit_color")):
        colmap["c1_quantity_policy"] = policy

    colmap["map_sheet_sync"] = {
        "syncedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "sheetTitle": SHEET_TITLE,
        "source": "MAP",
        "rowCountApplied": sum(1 for r in rows if r["productType"] in sheet_pts),
        "diffKeys": {
            "defaults": list(diff["defaults"].keys()),
            "master_columns": list(diff["master_columns"].keys()),
            "policy": list(diff["policy"].keys()),
            "aliases_ja": list(diff["aliases_ja"].keys()),
        },
    }
    return diff


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="MAP sheet → column_map.json")
    ap.add_argument("--profile", choices=sorted(PROFILES.keys()), required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--config", default="")
    args = ap.parse_args(argv)

    sys.stdout.reconfigure(encoding="utf-8")
    cfg = load_cfg() if not args.config else json.loads(Path(args.config).read_text(encoding="utf-8"))
    prof = PROFILES[args.profile]
    json_path = SCRIPT_DIR / prof["json"]
    if not json_path.is_file():
        print("JSONなし:", json_path, file=sys.stderr)
        return 2

    colmap = json.loads(json_path.read_text(encoding="utf-8"))
    creds = get_creds()
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    _, rows = fetch_map_rows(svc, cfg["spreadsheet_id"])
    print("MAP rows fetched:", len(rows))

    before = json.dumps(colmap, ensure_ascii=False, sort_keys=True)
    diff = apply_rows_to_colmap(colmap, rows, prof["sheet_product_types"])
    after = json.dumps(colmap, ensure_ascii=False, sort_keys=True)

    print("diff.defaults", json.dumps(diff["defaults"], ensure_ascii=False))
    print("diff.master_columns", json.dumps(diff["master_columns"], ensure_ascii=False))
    print("diff.policy", json.dumps(diff["policy"], ensure_ascii=False))
    print("diff.aliases_ja", json.dumps(diff["aliases_ja"], ensure_ascii=False))
    print("c1_quantity_policy", json.dumps(colmap.get("c1_quantity_policy"), ensure_ascii=False, indent=2))

    if args.dry_run:
        print("DRY_RUN: no write. changed=", before != after)
        return 0

    bak = json_path.with_suffix(json_path.suffix + ".bak_map_sync")
    shutil.copy2(json_path, bak)
    json_path.write_text(
        json.dumps(colmap, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print("wrote", json_path)
    print("backup", bak)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
