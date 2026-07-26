#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
C1 / C1-1b: HPC 純正 xlsm → PACKAGED（ローカル本線）

- GENERATED + マスタCSV併読（列名＝マスタ見出し）
- 必須マスタ欠落・URL空 → 親SKU一式除外
- 商品タックスコードはマスタ値（固定しない）
- 指紋不一致 → 本番停止／DRY_RUN警告
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    import openpyxl
except ImportError:
    print("openpyxl が必要です: pip install -r requirements.txt", file=sys.stderr)
    sys.exit(2)

SCRIPT_DIR = Path(__file__).resolve().parent
LOG = logging.getLogger("c1_packaged")


def _utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")


def _load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _save_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")


def _cell_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def resolve_path(p: str, base: Path) -> Path:
    path = Path(p)
    if not path.is_absolute():
        path = (base / path).resolve()
    return path


def compute_header_fingerprint(ws, rows: List[int], max_col: int = 300) -> str:
    parts: List[str] = []
    for r in rows:
        for c in range(1, max_col + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            parts.append("%d:%d:%s" % (r, c, _cell_str(v)))
    raw = "\n".join(parts).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


_REQUIRED_GENERATED_HEADERS = (
    "parentSku",
    "sellerSku",
    "priceAmazon",
    "shippingTemplate",
    "variationRole",
)
_ROLE_TOKENS = frozenset({"parent", "child", "親", "子供", "子"})


def _normalize_list_price(raw: str) -> str:
    """税込み参考価格用。数字のみ残す。範囲・調査メモは無効。"""
    s = _cell_str(raw)
    if not s:
        return ""
    s = re.sub(r"[円￥¥,\s]", "", s)
    if re.fullmatch(r"\d+(\.\d+)?", s):
        return s
    return ""


def load_generated_csv(path: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        raw_rows = list(csv.reader(f))
    if not raw_rows:
        raise ValueError("GENERATED CSV が空です: %s" % path)
    headers = [_cell_str(h) for h in raw_rows[0]]
    if not any(headers):
        raise ValueError("GENERATED CSV にヘッダーがありません: %s" % path)
    missing = [h for h in _REQUIRED_GENERATED_HEADERS if h not in headers]
    if missing:
        raise ValueError("GENERATED CSV 必須列不足: %s" % ", ".join(missing))
    n = len(headers)
    rows: List[Dict[str, str]] = []
    for i, raw in enumerate(raw_rows[1:], start=2):
        if not raw or not any(_cell_str(c) for c in raw):
            continue
        if len(raw) != n:
            raise ValueError(
                "GENERATED CSV 列数不一致: line=%d expected=%d actual=%d "
                "(カンマ不足・余分で shippingTemplate がずれます)" % (i, n, len(raw))
            )
        row = {headers[j]: _cell_str(raw[j] if j < len(raw) else "") for j in range(n)}
        ship = _cell_str(row.get("shippingTemplate"))
        role = _cell_str(row.get("variationRole")).lower()
        if ship.lower() in _ROLE_TOKENS or ship in _ROLE_TOKENS:
            raise ValueError(
                "GENERATED CSV shippingTemplate に親子ロールが入っています (line=%d value=%r)。"
                "列ずれの可能性。csv.writer で書き直してください。" % (i, ship)
            )
        if role and role not in ("parent", "child"):
            LOG.warning("GENERATED variationRole 非標準 line=%d value=%r", i, role)
        rows.append(row)
    return headers, rows


def _norm_header(name: Any) -> str:
    """Sheets書き出しで改行入り見出しになることがあるため正規化。"""
    return _cell_str(name).replace("\n", " ").replace("\r", " ")


def _pick_master_field(row: Dict[str, str], aliases: List[str]) -> str:
    for name in aliases:
        if name in row and _cell_str(row.get(name)):
            return _cell_str(row.get(name))
    # strip / 改行除去後の一致
    norm_map = {_norm_header(k): v for k, v in row.items()}
    for name in aliases:
        key = _norm_header(name)
        if key in norm_map and _cell_str(norm_map.get(key)):
            return _cell_str(norm_map.get(key))
    return ""


def _find_master_header_row(rows: List[List[str]], master_columns: dict) -> int:
    """Google Sheets全件CSVは先頭に注記行があり、列名行が1行目とは限らない。"""
    must = set()
    for key in ("parent_sku", "child_sku"):
        for a in master_columns.get(key) or []:
            must.add(_norm_header(a))
    if not must:
        must = {"親SKU", "子SKU"}
    for i, row in enumerate(rows[:40]):
        cells = {_norm_header(c) for c in row}
        if must.issubset(cells):
            return i
    raise ValueError(
        "マスタCSVに列名行（親SKU/子SKU）が見つかりません。Sheets書き出しか列名を確認してください。"
    )


def load_master_csv(path: Path, master_columns: dict) -> Dict[str, Dict[str, str]]:
    """子SKU優先・親SKUでも引ける index: sku -> normalized fields。"""
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        all_rows = list(csv.reader(f))
    if not all_rows:
        raise ValueError("マスタCSVが空です: %s" % path)
    header_idx = _find_master_header_row(all_rows, master_columns)
    raw_headers = all_rows[header_idx]
    # 重複列名は後ろを優先（マスタ後半の本番列が前半の旧列を上書き）
    headers: List[str] = []
    seen: Dict[str, int] = {}
    for h in raw_headers:
        key = _norm_header(h) or ("__empty_%d" % len(headers))
        if key in seen:
            headers[seen[key]] = "__dup_%d_%s" % (seen[key], key)
        seen[key] = len(headers)
        headers.append(key)

    index: Dict[str, Dict[str, str]] = {}
    data_rows = 0
    for raw in all_rows[header_idx + 1 :]:
        if not raw or not any(_cell_str(c) for c in raw):
            continue
        row = {
            headers[i]: _cell_str(raw[i] if i < len(raw) else "")
            for i in range(len(headers))
        }
        fields: Dict[str, str] = {}
        for key, aliases in master_columns.items():
            fields[key] = _pick_master_field(row, aliases)
        child = fields.get("child_sku") or ""
        parent = fields.get("parent_sku") or ""
        if not child and not parent:
            continue
        data_rows += 1
        if child:
            index[child] = fields
        if parent and parent not in index:
            index[parent] = fields
        if parent and not child:
            index[parent] = fields
    LOG.info(
        "マスタCSV読込: %s header_row=%d data_rows=%d rows_index=%d",
        path,
        header_idx + 1,
        data_rows,
        len(index),
    )
    return index


def group_by_parent(rows: List[Dict[str, str]]) -> Dict[str, Dict[str, Any]]:
    groups: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        parent = row.get("parentSku") or ""
        if not parent:
            continue
        g = groups.setdefault(parent, {"parent": None, "children": []})
        role = (row.get("variationRole") or "").lower()
        if role == "parent" or (not row.get("childSku") and row.get("sellerSku") == parent):
            g["parent"] = row
        else:
            g["children"].append(row)
    return groups


def sub_batch_id_from_generated(path: Path, rows: List[Dict[str, str]]) -> str:
    name = path.name
    m = re.match(r"^(.+)_GENERATED\.csv$", name, re.I)
    if m:
        return m.group(1)
    if rows and rows[0].get("track"):
        return "%s_%s" % (rows[0].get("track"), _utc_stamp())
    return "C1_%s" % _utc_stamp()


def _yes_no_jp(raw: str, default: str = "いいえ") -> str:
    s = _cell_str(raw).lower()
    if not s:
        return default
    if s in ("はい", "yes", "true", "1", "y"):
        return "はい"
    if s in ("いいえ", "no", "false", "0", "n"):
        return "いいえ"
    if "はい" in _cell_str(raw):
        return "はい"
    if "いいえ" in _cell_str(raw):
        return "いいえ"
    return default


def split_keywords(text: str, n: int = 5) -> List[str]:
    parts = re.split(r"[\s　,、/|]+", _cell_str(text))
    parts = [p for p in parts if p]
    out = parts[:n]
    while len(out) < n:
        out.append("")
    return out


def master_for_sku(master_index: Dict[str, Dict[str, str]], sku: str, parent_sku: str) -> Dict[str, str]:
    """子行に空欄が多いSheetsマスタ向け: 親の値を継承し、子の非空だけ上書き。
    定価は親の非数字（調査メモ等）を子へ継承しない。
    """
    parent_fields = master_index.get(parent_sku) or {} if parent_sku else {}
    child_fields = master_index.get(sku) or {} if sku else {}
    if not parent_fields and not child_fields:
        return {}
    if sku and sku == parent_sku:
        return dict(parent_fields or child_fields)
    merged = dict(parent_fields)
    for k, v in child_fields.items():
        if _cell_str(v):
            merged[k] = v
    if not _cell_str(child_fields.get("list_price")):
        parent_lp = _cell_str(parent_fields.get("list_price"))
        merged["list_price"] = parent_lp if _normalize_list_price(parent_lp) else ""
    return merged


def resolve_url(
    row: Dict[str, str],
    url_override: Dict[str, str],
    master: Dict[str, str],
) -> str:
    """Amazon MAIN URL のみ。GENERATED / CDN へのフォールバック禁止（社長決定）。"""
    sku = row.get("sellerSku") or row.get("childSku") or row.get("parentSku") or ""
    if sku and sku in url_override:
        return _cell_str(url_override[sku])
    return _cell_str(master.get("amazon_main_url"))


def resolve_size(row: Dict[str, str], size_map: Dict[str, str], master: Dict[str, str]) -> str:
    sku = row.get("sellerSku") or row.get("childSku") or ""
    if sku and sku in size_map:
        return _cell_str(size_map[sku])
    if _cell_str(master.get("variation_value")):
        return _cell_str(master.get("variation_value"))
    for key in ("variationValue", "size", "バリエーション値"):
        if _cell_str(row.get(key)):
            return _cell_str(row.get(key))
    return _cell_str(row.get("setCount"))


def resolve_list_price(
    sku: str,
    parent_sku: str,
    master: Dict[str, str],
    list_price_override: Dict[str, str],
    master_index: Optional[Dict[str, Dict[str, str]]] = None,
    is_parent: bool = False,
    from_children: Optional[List[str]] = None,
) -> str:
    """優先: override → 当SKU行の定価 →（親行なら）子の定価 → マージ結果の数字のみ。"""
    for key in (sku, parent_sku):
        if key and key in list_price_override:
            norm = _normalize_list_price(list_price_override[key])
            if norm:
                return norm
            LOG.warning("list_price_override が数字ではない sku=%s value=%r", key, list_price_override[key])

    if master_index and sku and sku in master_index:
        own = _normalize_list_price(master_index[sku].get("list_price") or "")
        if own:
            return own

    if is_parent:
        for lp in from_children or []:
            norm = _normalize_list_price(lp)
            if norm:
                return norm
        if master_index and parent_sku:
            for fields in master_index.values():
                if fields.get("parent_sku") != parent_sku or not fields.get("child_sku"):
                    continue
                norm = _normalize_list_price(fields.get("list_price") or "")
                if norm:
                    return norm

    return _normalize_list_price(master.get("list_price") or "")


def resolve_shipping(gen_row: Dict[str, str], defaults: dict) -> str:
    ship = _cell_str(gen_row.get("shippingTemplate"))
    if not ship or ship.lower() in _ROLE_TOKENS or ship in _ROLE_TOKENS:
        fallback = defaults.get("shipping", "送料無料パターン")
        if ship:
            LOG.warning("shippingTemplate 不正 %r → 既定 %r", ship, fallback)
        return fallback
    return ship


def build_row_attrs(
    gen_row: Dict[str, str],
    master: Dict[str, str],
    defaults: dict,
    url_override: Dict[str, str],
    size_map: Dict[str, str],
    list_price_override: Dict[str, str],
    is_parent: bool,
    parent_sku: str,
    master_index: Optional[Dict[str, Dict[str, str]]] = None,
    list_price_from_children: Optional[List[str]] = None,
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    sku = _cell_str(gen_row.get("sellerSku") or gen_row.get("childSku") or parent_sku)
    url = resolve_url(gen_row, url_override, master)
    if not url:
        return None, "Amazon MAIN URL 空（フォールバック禁止）"

    title = _cell_str(master.get("product_name_amazon"))
    if not title or "手入力" in title:
        title = _cell_str(gen_row.get("productName"))
    if not title or "手入力" in title:
        return None, "商品名不正/プレースホルダ"

    price = _cell_str(gen_row.get("priceAmazon"))
    if not price:
        return None, "priceAmazon 空"

    bullet1 = _cell_str(master.get("bullet1"))
    if not bullet1:
        return None, "マスタ 商品説明の箇条書き① 空"

    tax = _cell_str(master.get("tax_code"))
    if not tax:
        return None, "マスタ 商品タックスコード 空"

    origin = _cell_str(master.get("origin"))
    if not origin:
        return None, "マスタ 原産国 空"

    unit_count = _cell_str(master.get("unit_count"))
    unit_uom = _cell_str(master.get("unit_uom"))
    if not unit_count or not unit_uom:
        return None, "マスタ ユニット数/単位 空"

    mfr_name = _cell_str(master.get("mfr_name"))
    if not mfr_name:
        return None, "マスタ メーカー名 空"

    size = "" if is_parent else resolve_size(gen_row, size_map, master)
    if not is_parent and not size:
        return None, "サイズ/バリエーション値 空"

    mfr_part = _cell_str(gen_row.get("manufacturerPart")) or ("" if is_parent else sku)
    bullets = [
        bullet1,
        _cell_str(master.get("bullet2")) or bullet1,
        _cell_str(master.get("bullet3")) or bullet1,
        _cell_str(master.get("bullet4")) or bullet1,
        _cell_str(master.get("bullet5")) or bullet1,
    ]
    kws = split_keywords(master.get("keywords") or "")

    heat = _yes_no_jp(master.get("heat") or "", defaults.get("heat", "いいえ"))
    liquid = _yes_no_jp(master.get("liquid") or "", defaults.get("liquid", "いいえ"))
    color = _cell_str(master.get("color")) or defaults.get("color", "その他")
    list_price = resolve_list_price(
        sku,
        parent_sku,
        master,
        list_price_override,
        master_index=master_index,
        is_parent=is_parent,
        from_children=list_price_from_children,
    )
    if not list_price:
        return None, "税込み参考価格（定価）が数字でない。子SKU行の定価か list_price_override_map を数字に"
    ingredients = _cell_str(master.get("ingredients"))

    return {
        "sku": sku,
        "url": url,
        "title": title,
        "price": price,
        "size": size or None,
        "mfr": mfr_part or None,
        "mfr_name": mfr_name,
        "highlight": bullet1,
        "desc": bullet1,
        "specs": bullets,
        "keywords": kws,
        "color": color,
        "heat": heat,
        "liquid": liquid,
        "origin": origin,
        "tax_code": tax,
        "unit_count": unit_count,
        "unit_uom": unit_uom,
        "list_price": list_price,
        "ingredients": ingredients or None,
        "shipping": resolve_shipping(gen_row, defaults),
        "inventory": _cell_str(gen_row.get("inventory") or "0"),
    }, None


def evaluate_parent(
    parent_sku: str,
    bundle: Dict[str, Any],
    url_override: Dict[str, str],
    size_map: Dict[str, str],
    list_price_override: Dict[str, str],
    master_index: Dict[str, Dict[str, str]],
    defaults: dict,
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    parent_row = bundle.get("parent")
    children = bundle.get("children") or []
    if not children:
        return None, "子SKUなし"
    if not parent_row:
        return None, "親行なし（GENERATED）"

    # 子を先に構築し、親の参考価格は子SKU行の定価を使う
    built_children = []
    for ch in children:
        sku = _cell_str(ch.get("sellerSku") or ch.get("childSku"))
        if not sku:
            return None, "子 sellerSku 空"
        ch_master = master_for_sku(master_index, sku, parent_sku)
        attrs, cerr = build_row_attrs(
            ch,
            ch_master,
            defaults,
            url_override,
            size_map,
            list_price_override,
            False,
            parent_sku,
            master_index=master_index,
        )
        if attrs is None:
            return None, "子(%s):%s" % (sku, cerr or "unknown")
        built_children.append(attrs)

    child_prices = [c.get("list_price") or "" for c in built_children]
    parent_master = master_for_sku(master_index, parent_sku, parent_sku)
    parent_attrs, err = build_row_attrs(
        parent_row,
        parent_master,
        defaults,
        url_override,
        size_map,
        list_price_override,
        True,
        parent_sku,
        master_index=master_index,
        list_price_from_children=child_prices,
    )
    if parent_attrs is None:
        return None, "親:" + (err or "unknown")

    for attrs in built_children:
        if not attrs["price"]:
            attrs["price"] = parent_attrs["price"]

    return {
        "parent_sku": parent_sku,
        "parent": parent_attrs,
        "children": built_children,
    }, None


def write_family_rows(ws, start_row: int, family: Dict[str, Any], colmap: dict, defaults: dict) -> int:
    cols = colmap["cols"]

    def setc(r: int, key: str, value: Any) -> None:
        if key not in cols:
            return
        if value is None or value == "":
            return
        ws.cell(r, cols[key]).value = value

    def write_common(r: int, attrs: Dict[str, Any], parentage: str, parent_sku_val: Any) -> None:
        setc(r, "sku", attrs["sku"])
        setc(r, "product_type", defaults["product_type"])
        setc(r, "action", defaults["action"])
        setc(r, "parentage", parentage)
        setc(r, "parent_sku", parent_sku_val)
        setc(r, "var_theme", defaults["var_theme"])
        setc(r, "title", attrs["title"])
        setc(r, "highlight", attrs["highlight"])
        setc(r, "brand", defaults["brand"])
        setc(r, "id_type", defaults["id_type"])
        setc(r, "browse", defaults["browse"])
        setc(r, "mfr_name", attrs["mfr_name"])
        setc(r, "main_image_url", attrs["url"])
        setc(r, "desc", attrs["desc"])
        specs = attrs.get("specs") or []
        for i, key in enumerate(["spec1", "spec2", "spec3", "spec4", "spec5"]):
            if i < len(specs):
                setc(r, key, specs[i])
        kws = attrs.get("keywords") or []
        for i, key in enumerate(["keyword1", "keyword2", "keyword3", "keyword4", "keyword5"]):
            if i < len(kws) and kws[i]:
                setc(r, key, kws[i])
        setc(r, "color", attrs.get("color"))
        setc(r, "size", attrs.get("size"))
        setc(r, "mfr_part", attrs.get("mfr"))
        setc(r, "import_type", defaults["import_type"])
        setc(r, "exclusive", defaults["exclusive"])
        setc(r, "heat", attrs.get("heat") or defaults["heat"])
        setc(r, "ingredients", attrs.get("ingredients"))
        setc(r, "unit_count", attrs.get("unit_count"))
        setc(r, "unit_uom", attrs.get("unit_uom"))
        setc(r, "condition", defaults["condition"])
        setc(r, "list_price", attrs.get("list_price"))
        setc(r, "tax_code", attrs.get("tax_code"))
        setc(r, "fulfillment", defaults["fulfillment"])
        inv = attrs.get("inventory") or "0"
        setc(r, "inventory", int(inv) if str(inv).isdigit() else inv)
        price = attrs["price"]
        setc(r, "price", float(price) if _is_number(price) else price)
        setc(r, "shipping", attrs.get("shipping") or defaults.get("shipping"))
        setc(r, "origin", attrs.get("origin"))
        setc(r, "battery_needed", defaults["battery_needed"])
        setc(r, "battery_included", defaults["battery_included"])
        setc(r, "hazmat", defaults["hazmat"])
        setc(r, "liquid", attrs.get("liquid") or defaults["liquid"])

    r = start_row
    write_common(r, family["parent"], "親", None)
    r += 1
    for ch in family["children"]:
        write_common(r, ch, "子供", family["parent_sku"])
        r += 1
    return r


def _is_number(s: str) -> bool:
    try:
        float(s)
        return True
    except (TypeError, ValueError):
        return False


def clear_data_rows(ws, start_row: int, max_col: int = 228) -> None:
    if ws.max_row < start_row:
        return
    for r in range(start_row, ws.max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None


def build_mapping_report(families: List[Dict[str, Any]], colmap: dict) -> List[dict]:
    cols = colmap["cols"]
    report = []
    for fam in families:
        p = fam["parent"]
        report.append(
            {
                "parentSku": fam["parent_sku"],
                "childCount": len(fam["children"]),
                "taxCode": p.get("tax_code"),
                "mfrName": p.get("mfr_name"),
                "mainImageCol": cols.get("main_image_url"),
                "children": [
                    {
                        "sku": c["sku"],
                        "size": c.get("size"),
                        "urlPresent": bool(c.get("url")),
                        "taxCode": c.get("tax_code"),
                    }
                    for c in fam["children"]
                ],
            }
        )
    return report


def run(config_path: Path, mode: str) -> int:
    cfg = _load_json(config_path)
    base = config_path.parent

    mode = (mode or cfg.get("mode") or "dry_run").strip().lower()
    if mode not in ("dry_run", "prod"):
        raise SystemExit("mode は dry_run または prod")

    template_path = resolve_path(cfg["template_path"], base)
    generated_csv = resolve_path(cfg["generated_csv"], base)
    output_dir = resolve_path(cfg["output_dir"], base)
    log_dir = resolve_path(cfg.get("log_dir") or str(output_dir), base)
    fp_path = resolve_path(cfg.get("fingerprint_path") or "fingerprints/hpc_header_r3_r5.json", base)
    cm_raw = cfg.get("column_map_path") or "hpc_column_map.json"
    map_path = Path(cm_raw)
    if not map_path.is_absolute():
        cand = (base / map_path).resolve()
        map_path = cand if cand.is_file() else (SCRIPT_DIR / map_path).resolve()

    colmap = _load_json(map_path)
    defaults = dict(colmap.get("defaults") or {})
    defaults.update(cfg.get("defaults") or {})
    master_columns = colmap.get("master_columns") or {}

    master_index: Dict[str, Dict[str, str]] = {}
    master_csv_path = cfg.get("master_csv") or ""
    if master_csv_path:
        mp = resolve_path(master_csv_path, base)
        if not mp.is_file():
            raise FileNotFoundError("マスタCSVがありません: %s" % mp)
        master_index = load_master_csv(mp, master_columns)
    else:
        LOG.warning("master_csv 未設定 — C1-1b必須列は欠落し親除外の可能性大")

    url_override = {str(k): str(v) for k, v in (cfg.get("url_override_map") or {}).items()}
    size_map = {str(k): str(v) for k, v in (cfg.get("size_map") or {}).items()}
    list_price_override = {
        str(k): str(v) for k, v in (cfg.get("list_price_override_map") or {}).items()
    }
    parent_filter = set(cfg.get("parent_sku_filter") or [])

    output_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    if not template_path.is_file():
        raise FileNotFoundError("テンプレがありません: %s" % template_path)
    if not generated_csv.is_file():
        raise FileNotFoundError("GENERATED CSVがありません: %s" % generated_csv)

    headers, gen_rows = load_generated_csv(generated_csv)
    sub_batch_id = cfg.get("sub_batch_id") or sub_batch_id_from_generated(generated_csv, gen_rows)
    run_id = "C1_%s_%s" % (sub_batch_id, _utc_stamp())

    groups = group_by_parent(gen_rows)
    if parent_filter:
        groups = {k: v for k, v in groups.items() if k in parent_filter}

    accepted: List[Dict[str, Any]] = []
    excluded: List[Dict[str, str]] = []
    for parent_sku, bundle in sorted(groups.items()):
        fam, reason = evaluate_parent(
            parent_sku,
            bundle,
            url_override,
            size_map,
            list_price_override,
            master_index,
            defaults,
        )
        if fam is None:
            excluded.append({"parentSku": parent_sku, "reason": reason or "unknown"})
            LOG.warning("EXCLUDE parent=%s reason=%s", parent_sku, reason)
        else:
            accepted.append(fam)

    wb_probe = openpyxl.load_workbook(template_path, keep_vba=True, data_only=False)
    sheet_name = colmap.get("sheet_name") or "テンプレート"
    if sheet_name not in wb_probe.sheetnames:
        raise RuntimeError("シートがありません: %s / %s" % (sheet_name, wb_probe.sheetnames))
    ws_probe = wb_probe[sheet_name]
    fp_rows = colmap.get("header_fingerprint_rows") or [3, 4, 5]
    current_fp = compute_header_fingerprint(ws_probe, fp_rows)
    wb_probe.close()

    fp_status = "missing_baseline"
    fp_ok = True
    if fp_path.is_file():
        baseline = _load_json(fp_path)
        expected = baseline.get("sha256") or ""
        if expected and expected != current_fp:
            fp_ok = False
            fp_status = "mismatch"
            LOG.error("テンプレ指紋不一致 expected=%s current=%s", expected, current_fp)
        else:
            fp_status = "match"
            LOG.info("テンプレ指紋一致 sha256=%s", current_fp)
    else:
        LOG.warning("指紋ベースライン無し: %s", fp_path)

    report = {
        "runId": run_id,
        "mode": mode,
        "version": "C1-1b",
        "subBatchId": sub_batch_id,
        "templatePath": str(template_path),
        "generatedCsv": str(generated_csv),
        "masterCsv": master_csv_path or None,
        "generatedHeaders": headers,
        "fingerprint": {
            "status": fp_status,
            "sha256": current_fp,
            "baselinePath": str(fp_path),
            "rows": fp_rows,
        },
        "acceptedParents": [f["parent_sku"] for f in accepted],
        "excludedParents": excluded,
        "mapping": build_mapping_report(accepted, colmap),
        "counts": {
            "acceptedParents": len(accepted),
            "excludedParents": len(excluded),
            "acceptedChildren": sum(len(f["children"]) for f in accepted),
        },
        "sanctuary": {
            "masterWrite": False,
            "templateOverwrite": False,
            "outputOverwrite": False,
            "urlFallback": False,
            "taxCodeFixed": False,
        },
    }

    report_path = log_dir / ("%s_C1_REPORT.json" % run_id)
    _save_json(report_path, report)
    LOG.info("レポート: %s", report_path)

    if mode == "prod" and not fp_ok:
        LOG.error("本番停止: テンプレ指紋不一致")
        return 3
    if mode == "prod" and fp_status == "missing_baseline":
        LOG.error("本番停止: 指紋ベースラインなし")
        return 3
    if mode == "prod" and not accepted:
        LOG.error("本番停止: 出力対象の親が0件")
        return 4

    write_xlsm = mode == "prod" or bool(cfg.get("write_dryrun_xlsm"))
    if not write_xlsm:
        LOG.info("DRY_RUN: xlsm 非作成")
        return 0 if accepted or excluded else 1

    if mode == "dry_run" and not fp_ok:
        LOG.warning("DRY_RUN: 指紋不一致だが継続")

    if not accepted:
        LOG.warning("accepted=0 — xlsm 非作成")
        return 1

    suffix = "_DRYRUN" if mode == "dry_run" else ""
    parent_tag = accepted[0]["parent_sku"] if len(accepted) == 1 else ("n%d" % len(accepted))
    out_name = "%s_PACKAGED_HPC_%s%s.xlsm" % (sub_batch_id, parent_tag, suffix)
    out_path = output_dir / out_name
    if out_path.exists():
        out_path = output_dir / ("%s_PACKAGED_HPC_%s_%s%s.xlsm" % (sub_batch_id, parent_tag, _utc_stamp(), suffix))

    work = output_dir / ("_work_%s.xlsm" % run_id)
    shutil.copy2(template_path, work)
    LOG.info("テンプレコピー: %s -> %s", template_path, work)

    try:
        wb = openpyxl.load_workbook(work, keep_vba=True, data_only=False)
        ws = wb[sheet_name]
        data_start = int(colmap.get("data_start_row") or 7)
        sample_row = int(colmap.get("sample_row") or 6)
        clear_data_rows(ws, data_start)
        row_cursor = data_start
        for fam in accepted:
            row_cursor = write_family_rows(ws, row_cursor, fam, colmap, defaults)
        wb.save(work)
        wb.close()
        shutil.move(str(work), str(out_path))
        LOG.info("出力: %s", out_path)
        report["outputPath"] = str(out_path)
        report["rowsWritten"] = row_cursor - data_start
        report["sampleRowPreserved"] = sample_row
        _save_json(report_path, report)
    except Exception:
        if work.exists():
            fail_keep = log_dir / ("%s_PARTIAL.xlsm" % run_id)
            try:
                shutil.move(str(work), str(fail_keep))
                LOG.error("途中失敗: 部分ファイルを残しました %s", fail_keep)
                report["partialPath"] = str(fail_keep)
                _save_json(report_path, report)
            except Exception:
                LOG.exception("部分ファイルの退避にも失敗")
        raise

    return 0


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="C1 HPC PACKAGED builder (C1-1b)")
    parser.add_argument("--config", required=True, help="config.json パス")
    parser.add_argument("--mode", choices=["dry_run", "prod"], default=None)
    parser.add_argument("--write-fingerprint", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    config_path = Path(args.config).resolve()
    cfg = _load_json(config_path)
    base = config_path.parent

    if args.write_fingerprint:
        template_path = resolve_path(cfg["template_path"], base)
        map_path = resolve_path(cfg.get("column_map_path") or str(SCRIPT_DIR / "hpc_column_map.json"), base)
        if not map_path.exists():
            map_path = SCRIPT_DIR / "hpc_column_map.json"
        colmap = _load_json(map_path)
        fp_path = resolve_path(cfg.get("fingerprint_path") or "fingerprints/hpc_header_r3_r5.json", base)
        wb = openpyxl.load_workbook(template_path, keep_vba=True, data_only=False)
        ws = wb[colmap.get("sheet_name") or "テンプレート"]
        rows = colmap.get("header_fingerprint_rows") or [3, 4, 5]
        sha = compute_header_fingerprint(ws, rows)
        wb.close()
        _save_json(
            fp_path,
            {
                "sha256": sha,
                "rows": rows,
                "templatePath": str(template_path),
                "recordedAt": datetime.now(timezone.utc).isoformat(),
                "productType": "HEALTH_PERSONAL_CARE",
            },
        )
        LOG.info("指紋を保存: %s sha256=%s", fp_path, sha)
        return 0

    return run(config_path, args.mode or cfg.get("mode") or "dry_run")


if __name__ == "__main__":
    sys.exit(main())
