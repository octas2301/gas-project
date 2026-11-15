#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
純正バルクの「データ定義」から、そのファイルの必須／条件付き必須を読む。

汎用ルール（PT固定しない）:
- 「必須」だけ常時埋める対象。
- 「条件付きで必須」は常時必須にしない。マスタ／コピー／安全既定があるときだけ書く。
- データ定義の「例」は値にしない（50・75・レッド等）。
- 寸法・重量・物流・FBA階層は必須扱いしない（マスタにあれば書く）。
- GROCERY の黒セル omit は缶飯だけ（honor_omit_for_black_cells）。
- ブラウズの書き方はそのファイルのプルダウン実値（数字だけ／経路＋番号）。缶飯の書き方を全件に流さない。
- 親行にはサイズ・在庫・親番号を書かない。専用列はマスタ生値で上書きしない。
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    import openpyxl
except ImportError:
    raise SystemExit("openpyxl が必要です")

LOG = logging.getLogger("c1_bulk_template_required")
WRITE_ALWAYS = {"必須"}
WRITE_COND = {"条件付きで必須"}
WRITE_REQ = WRITE_ALWAYS | WRITE_COND
# 条件付きでも「必須扱いしない」。マスタ値があるときだけ通す。
SKIP_REQUIRE_PREFIXES = {
    "item_package_dimensions",
    "item_package_weight",
    "item_dimensions",
    "item_weight",
    "master_pack_layers_per_pallet_quantity",
    "master_packs_per_layer_quantity",
    "package_contains_sku",
    "fc_shelf_life",
    "liquid_volume",
    "liquid_contents_description",
    "liquid_packaging_type",
    "is_liquid_double_sealed",
    "ghs",
    "hazmat",
    "safety_data_sheet_url",
    "melting_temperature",
    "produce_sold_as_format",
    "produce_classification",
    "package_size_name",
    "set_name",
    "each_unit_count",
    "container",
}
ATTR_TO_LOGICAL = {
    "special_ingredients": "special_ingredients",
    "temperature_rating": "temperature_rating",
    "item_form": "item_form",
    "ingredients": "ingredients",
    "color": "color",
    "flavor": "flavor",
    "is_heat_sensitive": "heat",
    "is_exclusive_product": "exclusive",
    "distribution_designation": "import_type",
    "number_of_items": "number_items",
    "part_number": "mfr_part",
    "unit_count": "unit_count",
    "condition_type": "condition",
    "list_price": "list_price",
    "is_expiration_dated_product": "expiration_dated",
    "product_expiration_type": "expiration_type",
    "fc_shelf_life": "shelf_life",
    "contains_liquid_contents": "liquid",
    "item_weight": "item_weight",
    "size": "size",
}
LOCAL_TO_LOGICAL = {
    "特記すべき原材料": "special_ingredients",
    "温度の定格": "temperature_rating",
    "商品形態": "item_form",
    "商品の形式": "item_form",
    "原料": "ingredients",
    "色": "color",
    "味": "flavor",
}
COPY_FROM = {
    "special_ingredients": "ingredients",
}
SAFE_DEFAULTS = {
    "temperature_rating": "常温：室温",
    "heat": "いいえ",
    "exclusive": "いいえ",
    "import_type": "正規品",
    "liquid": "いいえ",
    "condition": "新品",
}
# 専用の書き手がいる列。マスタ生値を passthrough で被せない。
PASSTHROUGH_NEVER = {
    "browse",
    "sku",
    "product_type",
    "parentage",
    "parent_sku",
    "size",
    "inventory",
}
BROWSE_FMT_DIGIT = "digit"
BROWSE_FMT_PATH_ID = "path_id"
BROWSE_FMT_PATH = "path"


def _attr_prefix(field: str) -> str:
    s = str(field or "").strip()
    if not s:
        return ""
    return s.split("[", 1)[0].split("#", 1)[0].strip()


def _pick_template(wb: Any) -> Any:
    for name in ("テンプレート", "Template", "template"):
        if name in wb.sheetnames:
            return wb[name]
    raise RuntimeError("テンプレートシートがありません")


def load_data_definition(xlsm: Path) -> List[dict]:
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    try:
        if "データ定義" not in wb.sheetnames:
            LOG.warning("データ定義シートなし: %s", xlsm.name)
            return []
        ws = wb["データ定義"]
        out: List[dict] = []
        for row in ws.iter_rows(min_row=3, values_only=True):
            field = str(row[1] or "").strip() if row and len(row) > 1 else ""
            local = str(row[2] or "").strip() if row and len(row) > 2 else ""
            example = str(row[4] or "").strip() if row and len(row) > 4 else ""
            req = str(row[5] or "").strip() if row and len(row) > 5 else ""
            if not field or req == "必須ですか?":
                continue
            out.append(
                {
                    "field": field,
                    "local": local,
                    "example": example,
                    "requirement": req,
                    "write": req in WRITE_REQ,
                    "prefix": _attr_prefix(field),
                }
            )
        return out
    finally:
        wb.close()


def _header_index(ws: Any, max_col: int) -> List[Tuple[int, str, str]]:
    """(col, row4 local, row5 attr)"""
    rows: List[Tuple[int, str, str]] = []
    for c in range(1, max_col + 1):
        local = str(ws.cell(4, c).value or "").strip()
        attr = str(ws.cell(5, c).value or "").strip()
        if local or attr:
            rows.append((c, local, attr))
    return rows


def _first_slot_col(headers: List[Tuple[int, str, str]], field: str, local: str) -> Optional[int]:
    field = str(field or "").strip()
    hits: List[int] = []
    for c, loc, attr in headers:
        if field and attr == field:
            hits.append(c)
        elif field and attr.startswith(field):
            hits.append(c)
        elif local and loc == local and "#2." not in attr and "#3." not in attr:
            hits.append(c)
    if not hits:
        return None
    return min(hits)


def _logical_key(item: dict, colmap: dict) -> str:
    prefix = item.get("prefix") or ""
    local = item.get("local") or ""
    field = item.get("field") or ""
    if "単位" in local and prefix == "unit_count":
        return "unit_uom"
    if prefix == "fc_shelf_life" and (".unit" in field or "残存" in local):
        return "shelf_life_unit"
    if prefix == "item_weight" and "単位" in local:
        return "item_weight_unit"
    if prefix in ATTR_TO_LOGICAL:
        return ATTR_TO_LOGICAL[prefix]
    if local in LOCAL_TO_LOGICAL:
        return LOCAL_TO_LOGICAL[local]
    aliases = colmap.get("xlsm_header_aliases") or {}
    for key, names in aliases.items():
        for n in names or []:
            if n and (n == field or str(n) == local):
                return str(key)
    return ""


def _enum_default(example: str) -> str:
    # データ定義の「例」は値に使わない（呼び出し側で不使用）。残置は互換。
    return ""


def _classify_browse_sample(value: str) -> str:
    s = str(value or "").strip()
    if not s:
        return ""
    if s.isdigit():
        return BROWSE_FMT_DIGIT
    if ">" in s and "(" in s:
        return BROWSE_FMT_PATH_ID
    if ">" in s:
        return BROWSE_FMT_PATH
    if "(" in s:
        return BROWSE_FMT_PATH_ID
    return ""


def _iter_defined_names(wb: Any):
    dns = wb.defined_names
    if hasattr(dns, "items"):
        return list(dns.items())
    return [(dn.name, dn) for dn in getattr(dns, "definedName", [])]


def detect_browse_xlsm_format(xlsm: Path) -> dict:
    """その純正ファイルのプルダウン実値から、ブラウズ欄の書き方を決める。PT固定しない。"""
    from openpyxl.utils import range_boundaries

    by_pt: Dict[str, str] = {}
    wb = openpyxl.load_workbook(xlsm, read_only=False, data_only=True)
    try:
        suffix = "recommended_browse_nodesmarketplace_idA1VC38T7YXB5281.value"
        for name, dn in _iter_defined_names(wb):
            n = str(name or "")
            if not n.endswith(suffix):
                continue
            pt = n[: -len(suffix)]
            if not pt or pt.startswith("form_") or pt.startswith("only_") or pt.startswith("req"):
                continue
            dests = list(dn.destinations)
            if not dests:
                continue
            sheet, coord = dests[0]
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            min_col, min_row, max_col, max_row = range_boundaries(coord)
            counts: Dict[str, int] = {}
            for r in range(min_row, min(max_row, min_row + 80) + 1):
                kind = _classify_browse_sample(ws.cell(r, min_col).value)
                if kind:
                    counts[kind] = counts.get(kind, 0) + 1
            if not counts:
                continue
            fmt = max(counts, key=counts.get)
            by_pt[pt] = fmt
            LOG.info(
                "browseFormat pt=%s fmt=%s counts=%s xlsm=%s",
                pt,
                fmt,
                counts,
                xlsm.name,
            )
    finally:
        wb.close()
    default = ""
    if by_pt:
        tally: Dict[str, int] = {}
        for fmt in by_pt.values():
            tally[fmt] = tally.get(fmt, 0) + 1
        default = max(tally, key=tally.get)
    return {
        "browse_xlsm_format": default or BROWSE_FMT_PATH_ID,
        "browse_xlsm_format_by_pt": by_pt,
    }


def apply_template_required(xlsm: Path, colmap: dict, resolved_cols: Dict[str, int]) -> dict:
    """データ定義を colmap に載せる。例では埋めない。寸法・重量は必須扱いしない。"""
    defs = load_data_definition(xlsm)
    write_items = [d for d in defs if d.get("write")]
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    try:
        ws = _pick_template(wb)
        max_col = int(colmap.get("fingerprint_max_col") or colmap.get("clear_max_col") or ws.max_column or 300)
        headers = _header_index(ws, max_col)
    finally:
        wb.close()

    cols = dict(resolved_cols or colmap.get("cols") or {})
    passthrough = list(colmap.get("passthrough_fields") or [])
    copy_from = dict(colmap.get("template_copy_from") or {})
    copy_from.update(COPY_FROM)
    orig_defaults = dict(colmap.get("defaults") or {})
    defaults = dict(orig_defaults)
    required_keys: List[str] = []
    plan: List[dict] = []

    for item in write_items:
        prefix = str(item.get("prefix") or "")
        skip_require = prefix in SKIP_REQUIRE_PREFIXES
        col = _first_slot_col(headers, item["field"], item["local"])
        if not col:
            continue
        key = _logical_key(item, colmap)
        if not key:
            continue
        default = str(orig_defaults.get(key) or "") or ("" if skip_require else str(SAFE_DEFAULTS.get(key) or ""))
        src = COPY_FROM.get(key)
        known = (
            key in cols
            or key in (colmap.get("xlsm_header_aliases") or {})
            or bool(src)
            or bool(default)
        )
        if skip_require:
            # 必須リストに載せない。既存 passthrough（マスタ値）は残す。例・安全既定は付けない。
            continue
        if not known:
            continue
        if key not in cols:
            cols[key] = int(col)
        always = item.get("requirement") in WRITE_ALWAYS
        if always and key not in required_keys:
            required_keys.append(key)
        if key not in PASSTHROUGH_NEVER and key not in passthrough:
            passthrough.append(key)
        if default and not defaults.get(key):
            defaults[key] = default
        plan.append(
            {
                "logicalKey": key,
                "col": int(col),
                "local": item.get("local"),
                "requirement": item.get("requirement"),
                "copyFrom": src or "",
                "hasDefault": bool(defaults.get(key)),
                "skipRequire": False,
            }
        )

    policy = dict(colmap.get("c1_quantity_policy") or {})
    honor_omit = bool(policy.get("honor_omit_for_black_cells"))
    if not honor_omit:
        planned_keys = {p.get("logicalKey") for p in plan}
        if "temperature_rating" in planned_keys:
            policy["omit_temperature_rating"] = False
        if "item_form" in planned_keys:
            policy["omit_item_form"] = False
    # 色はマスタがあるときだけ。既定レッドは禁止のまま。
    passthrough = [k for k in passthrough if k not in PASSTHROUGH_NEVER]
    browse_fmt = detect_browse_xlsm_format(xlsm)
    out = dict(colmap)
    out["cols"] = cols
    out["passthrough_fields"] = passthrough
    out["defaults"] = defaults
    out["c1_quantity_policy"] = policy
    out["template_required_keys"] = required_keys
    out["template_copy_from"] = copy_from
    out["template_required_plan"] = plan
    out["template_skip_require_prefixes"] = sorted(SKIP_REQUIRE_PREFIXES)
    out["browse_xlsm_format"] = browse_fmt.get("browse_xlsm_format") or BROWSE_FMT_PATH_ID
    out["browse_xlsm_format_by_pt"] = browse_fmt.get("browse_xlsm_format_by_pt") or {}
    LOG.info(
        "template_required xlsm=%s writeDefs=%d planned=%d alwaysKeys=%s browseFmt=%s",
        xlsm.name,
        len(write_items),
        len(plan),
        ",".join(required_keys[:20]),
        out["browse_xlsm_format"],
    )
    return out
