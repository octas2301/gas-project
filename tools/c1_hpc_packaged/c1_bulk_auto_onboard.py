#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
新しい純正 xlsm（09/06）から指紋＋列マップを棚へ載せる。

E③ の 09 取込は Browse ルーティングだけ。PACKAGED 充填には PT＋指紋＋列マップが要る。
09 に置いて E③ したあと、B-T1 がここを呼ぶとアップロード用 xlsm まで作れる。

HERB.xlsm は本線化しない（B-T2 ロック）。
"""

from __future__ import annotations

import json
import logging
import re
import shutil
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import openpyxl
except ImportError:
    raise SystemExit("openpyxl が必要です")

from c1_bulk_name_map import resolve_cols_by_aliases
from c1_bulk_template_required import apply_template_required
from c1_packaged import compute_header_fingerprint
from c1_shelf_browse_extract import (  # noqa: E402
    DEFAULT_06,
    DEFAULT_DRIVE04,
    SCRIPT_DIR,
    extract_browse_rows,
    merge_catalog,
    sync_registry_browse_index,
)

LOG = logging.getLogger("c1_bulk_auto_onboard")
DEFAULT_09 = DEFAULT_DRIVE04 / "09.SC純正バルクxlsm保存（人間がDLして保存→Agent確認後06へ）"
SKIP_STEMS = {"HERB"}
# 単体 HERB.xlsm は載せない。PT=HERB は複合（SEASONING 棚）で本線化する。
COMPOSITE_ALIASES = {"HERB": "SEASONING"}
ALIAS_SOURCE = SCRIPT_DIR / "food_fish_grocery_column_map.json"
FP_ROWS = [3, 4, 5]


def _utc() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _save_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _pt_norm(pt: str) -> str:
    return re.sub(r"[^A-Z0-9_]", "_", str(pt or "").strip().upper())


def find_xlsm_for_pt(pt: str, dir09: Path, dir06: Path) -> Optional[Path]:
    stem = _pt_norm(pt)
    if stem in SKIP_STEMS:
        return None
    want = (stem + ".xlsm").upper()
    for folder in (dir06, dir09):
        if not folder.is_dir():
            continue
        for p in folder.glob("*.xlsm"):
            if p.name.startswith("~$"):
                continue
            if p.name.upper() == want:
                return p
    return None


def ensure_in_06(src: Path, dir06: Path) -> Path:
    dir06.mkdir(parents=True, exist_ok=True)
    dest = dir06 / src.name
    if dest.is_file():
        return dest
    shutil.copy2(src, dest)
    LOG.info("copied_to_06 %s", dest.name)
    return dest


def _pick_template(wb: Any) -> Any:
    for name in ("テンプレート", "Template", "template"):
        if name in wb.sheetnames:
            return wb[name]
    raise RuntimeError("テンプレートシートがありません: %s" % (wb.sheetnames,))


def build_column_map(xlsm: Path, pt: str, sha: str, max_col: int) -> Dict[str, Any]:
    grocery = _load_json(ALIAS_SOURCE)
    aliases = dict(grocery.get("xlsm_header_aliases") or {})
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    try:
        ws = _pick_template(wb)
        resolved, gaps = resolve_cols_by_aliases(
            ws,
            aliases,
            header_rows=list(grocery.get("header_alias_rows") or [4, 5]),
            max_col=max_col,
            legacy_cols={},
        )
    finally:
        wb.close()
    miss_req = [g for g in gaps if g.get("required") and g.get("status") == "MISS"]
    if miss_req:
        keys = ",".join(str(g.get("logicalKey")) for g in miss_req)
        raise RuntimeError("必須列が項目名で解決できません: %s / %s" % (xlsm.name, keys))
    drop = [g["logicalKey"] for g in gaps if g.get("status") == "MISS"]
    for k in drop:
        aliases.pop(k, None)
        resolved.pop(k, None)
    m = deepcopy(grocery)
    m["version"] = "C1-AUTO-1"
    m["profile"] = pt
    m["output_label"] = pt
    m["fingerprint_max_col"] = int(max_col)
    m["clear_max_col"] = int(max_col)
    m["cols"] = {k: int(v) for k, v in resolved.items()}
    m["xlsm_header_aliases"] = aliases
    m["defaults"] = dict(m.get("defaults") or {})
    m["defaults"]["product_type"] = pt
    m["p4b_c1_product_types"] = [pt]
    m["p4b_require_master_pt_browse"] = True
    policy = dict(m.get("c1_quantity_policy") or {})
    policy["honor_omit_for_black_cells"] = False
    policy["omit_item_form"] = False
    policy["omit_temperature_rating"] = False
    policy["notes"] = "auto-onboard. aliases=name-map. 必須はデータ定義から。GROCERY黒セルomitはコピーしない。"
    m["c1_quantity_policy"] = policy
    m["notes"] = "B-T2 auto-onboard %s %s sha=%s" % (_utc(), xlsm.name, sha)
    m.pop("map_sheet_sync", None)
    return apply_template_required(xlsm, m, resolved)


def ensure_composite_alias(registry: dict, pt: str) -> Dict[str, Any]:
    src_pt = COMPOSITE_ALIASES.get(pt)
    out: Dict[str, Any] = {"ok": False, "productType": pt}
    if not src_pt:
        out["reason"] = "not_alias"
        return out
    src = None
    existing = None
    for e in registry.get("entries") or []:
        ept = str(e.get("productType") or "").strip().upper()
        if ept == pt:
            existing = e
        if ept == src_pt:
            src = e
    if existing:
        sha = str(existing.get("fingerprintSha256") or "").strip()
        cmap = str(existing.get("columnMapPath") or "").strip()
        if sha and cmap and (SCRIPT_DIR / cmap).is_file():
            out["ok"] = True
            out["reason"] = "already_on_shelf"
            out["registryEntry"] = existing
            return out
    if not src or not str(src.get("fingerprintSha256") or "").strip():
        out["reason"] = "composite_missing"
        out["message"] = (
            "PT=%s は複合 %s で本線化するが、棚に %s がありません。"
            % (pt, src_pt, src_pt)
        )
        return out
    entry = dict(src)
    entry["productType"] = pt
    entry["notes"] = "composite alias of %s. %s.xlsm は使わない。" % (src_pt, pt)
    upsert_entry(
        registry,
        pt,
        str(src.get("fingerprintSha256") or ""),
        int(src.get("fingerprintMaxCol") or 321),
        str(src.get("preferredFileName") or ""),
        str(src.get("columnMapPath") or ""),
        str(src.get("fingerprintBaselinePath") or ""),
    )
    # upsert_entry overwrites notes; keep alias note
    for e in registry.get("entries") or []:
        if str(e.get("productType") or "").strip().upper() == pt:
            e["notes"] = entry["notes"]
            entry = e
            break
    out.update(
        {
            "ok": True,
            "reason": "composite_alias",
            "registryEntry": entry,
            "aliasOf": src_pt,
        }
    )
    return out


def upsert_entry(registry: dict, pt: str, sha: str, max_col: int, file_name: str, map_name: str, fp_name: str) -> dict:
    entry = {
        "productType": pt,
        "fingerprintSha256": sha,
        "fingerprintRows": list(FP_ROWS),
        "fingerprintMaxCol": int(max_col),
        "preferredFileName": file_name,
        "templateKind": "single",
        "columnMapPath": map_name,
        "fingerprintBaselinePath": fp_name,
        "notes": "auto-onboard %s" % _utc(),
    }
    entries = list(registry.get("entries") or [])
    out = []
    replaced = False
    for e in entries:
        if str(e.get("productType") or "").strip().upper() == pt:
            out.append(entry)
            replaced = True
        else:
            out.append(e)
    if not replaced:
        out.append(entry)
    registry["entries"] = out
    return entry


def sync_browse(xlsm: Path, pt: str, sha: str, map_name: str, registry_path: Path, catalog_path: Path, sync_drive: bool) -> int:
    rows = extract_browse_rows(xlsm, [pt], pt, sha, map_name)
    existing: Dict[str, Any] = {"rows": []}
    if catalog_path.is_file():
        existing = _load_json(catalog_path)
    catalog = merge_catalog(existing, rows, xlsm.name)
    _save_json(catalog_path, catalog)
    sync_registry_browse_index(registry_path, catalog)
    if sync_drive and DEFAULT_DRIVE04.is_dir():
        shutil.copy2(catalog_path, DEFAULT_DRIVE04 / "shelf_browse_catalog.json")
        shutil.copy2(registry_path, DEFAULT_DRIVE04 / "shelf_registry.json")
        LOG.info("synced Drive04 registry+catalog")
    return len(rows)


def auto_onboard_pt(
    product_type: str,
    *,
    dir09: Path = DEFAULT_09,
    dir06: Path = DEFAULT_06,
    registry_path: Path = SCRIPT_DIR / "shelf_registry.json",
    sync_drive: bool = True,
) -> Dict[str, Any]:
    pt = _pt_norm(product_type)
    out: Dict[str, Any] = {"ok": False, "productType": pt}
    if not pt:
        out["reason"] = "pt_empty"
        return out
    registry = _load_json(registry_path) if registry_path.is_file() else {"entries": []}
    if pt in COMPOSITE_ALIASES:
        aliased = ensure_composite_alias(registry, pt)
        if aliased.get("reason") == "composite_alias":
            _save_json(registry_path, registry)
            if sync_drive and DEFAULT_DRIVE04.is_dir():
                shutil.copy2(registry_path, DEFAULT_DRIVE04 / "shelf_registry.json")
        return aliased

    for e in registry.get("entries") or []:
        if str(e.get("productType") or "").strip().upper() != pt:
            continue
        sha = str(e.get("fingerprintSha256") or "").strip()
        cmap = str(e.get("columnMapPath") or "").strip()
        if sha and cmap and (SCRIPT_DIR / cmap).is_file():
            out["ok"] = True
            out["reason"] = "already_on_shelf"
            out["registryEntry"] = e
            return out

    found = find_xlsm_for_pt(pt, dir09, dir06)
    if not found:
        out["reason"] = "xlsm_missing"
        out["message"] = (
            "09/06 に %s.xlsm がありません。SC から純正を DL して 09 へ保存し、E③ のあと再実行してください。"
            % pt
        )
        return out

    xlsm06 = ensure_in_06(found, dir06)
    wb = openpyxl.load_workbook(xlsm06, read_only=True, data_only=True)
    try:
        ws = _pick_template(wb)
        max_col = int(ws.max_column or 300)
        sha = compute_header_fingerprint(ws, FP_ROWS, max_col)
    finally:
        wb.close()

    map_name = ("%s_column_map.json" % pt.lower())
    fp_name = "fingerprints/%s_header_r3_r5.json" % pt.lower()
    colmap = build_column_map(xlsm06, pt, sha, max_col)
    _save_json(SCRIPT_DIR / map_name, colmap)
    _save_json(
        SCRIPT_DIR / fp_name,
        {
            "templatePath": str(xlsm06),
            "preferredFileName": xlsm06.name,
            "sheetName": "テンプレート",
            "rows": list(FP_ROWS),
            "maxCol": max_col,
            "sha256": sha,
            "productTypes": [pt],
            "notes": "auto-onboard %s" % _utc(),
        },
    )
    entry = upsert_entry(registry, pt, sha, max_col, xlsm06.name, map_name, fp_name)
    _save_json(registry_path, registry)
    browse_n = sync_browse(
        xlsm06, pt, sha, map_name, registry_path, SCRIPT_DIR / "shelf_browse_catalog.json", sync_drive
    )
    LOG.info("onboarded pt=%s sha=%s map=%s browse=%d", pt, sha, map_name, browse_n)
    out.update(
        {
            "ok": True,
            "reason": "onboarded",
            "templatePath": str(xlsm06),
            "sha256": sha,
            "columnMapPath": map_name,
            "browseRows": browse_n,
            "registryEntry": entry,
        }
    )
    return out


def main(argv: Optional[List[str]] = None) -> int:
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    ap = argparse.ArgumentParser(description="09/06 の純正 xlsm から指紋＋列マップを棚へ載せる")
    ap.add_argument("--product-type", required=True)
    ap.add_argument("--registry", default=str(SCRIPT_DIR / "shelf_registry.json"))
    ap.add_argument("--no-sync-drive04", action="store_true")
    args = ap.parse_args(argv)
    result = auto_onboard_pt(
        args.product_type,
        registry_path=Path(args.registry),
        sync_drive=not args.no_sync_drive04,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    raise SystemExit(main())
