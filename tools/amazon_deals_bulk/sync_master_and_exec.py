# -*- coding: utf-8 -*-
"""
②推奨xlsx → マスタ名簿の補完 ＋ タイムセール施策行を同期。

- 対象は「タイムセール_マスタ」の有効SKUのみ
- レーンB: 名付き（horizon内）＋②SKU行のスケジュール列ドロップダウン（日付付き全候補）
- レーンA: **運用停止**（生成しない・既存削除）
- 公式とカスタムが重なるとき: カスタム側を短縮しメール通知
- 施策の並び: 開始日の降順 → 商品名の昇順
"""
from __future__ import annotations

import argparse
import hashlib
import json
import logging
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
LOG = logging.getLogger("amazon_deals_bulk.sync")

from paths import folder_path, latest_xlsx, load_config  # noqa: E402
from qty_logic import compute_q_deal, deal_day_count  # noqa: E402
from schedule_class import (  # noqa: E402
    fill_a_windows,
    format_ymd,
    is_a_equivalent_schedule,
    is_official_b_schedule,
    parse_ymd,
    pick_named_within_horizon,
    pick_schedules_sku_local,
    ranges_overlap,
    shrink_range_avoiding_blockers,
)
from points_logic import point_fields_from_row  # noqa: E402
from price_recovery_logic import recovery_fields_from_row  # noqa: E402
from sheet_schema import (  # noqa: E402
    ANALYSIS_SHEET,
    LANE_A,
    LANE_B,
    LEGACY_EXEC,
    LEGACY_LANE_A,
    MASTER_HEADER_ALIASES,
    MASTER_HEADERS,
    MASTER_HEADERS_DROPPED,
    MASTER_SHEET,
    SALE_HEADERS,
    SALE_SHEET,
    a_track_fields_from_row,
    normalize_master_header_name,
)
from sheets_io import (  # noqa: E402
    apply_master_display_formulas,
    apply_master_header_group_colors,
    apply_master_human_input_yellow,
    apply_master_point_unit_formats,
    apply_master_recovery_validations,
    ensure_headers_append,
    ensure_sheet,
    read_sheet_rows,
    sheets_service,
    write_headers_and_rows,
)
from template_parse import (  # noqa: E402
    collect_dated_dropdown_schedules,
    collect_schedule_catalog,
    collect_schedules,
    find_template_sheet,
    load_column_map,
    merge_schedules,
    read_template_rows,
)
from v30_source import resolve_v30_map  # noqa: E402

PRODUCT_SS = "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"
PRODUCT_SHEET = "▼商品マスタ(人間作業用)"


def prefer_row(cands: List[Dict[str, Any]]) -> Dict[str, Any]:
    for r in cands:
        if "おすすめ" in str(r.get("deal_type") or ""):
            return r
    return cands[0]


def _sale_id(lane: str, sku: str, key: str) -> str:
    return hashlib.sha1(f"{lane}|{sku}|{key}".encode("utf-8")).hexdigest()[:12]


def _truthy(v: Any) -> bool:
    return str(v or "").strip().upper() in ("TRUE", "はい", "YES", "Y", "1", "○")


def normalize_image_url(v: Any) -> str:
    """セル値から http(s) URL を取り出す（IMAGE式からも可）。"""
    s = str(v or "").strip()
    if not s:
        return ""
    if s.upper().startswith("=IMAGE("):
        m = re.search(r'=IMAGE\(\s*"([^"]+)"', s, flags=re.IGNORECASE)
        if m:
            s = m.group(1).strip()
    if s.startswith("http://") or s.startswith("https://"):
        return s
    return ""


def image_formula(url: Any) -> str:
    """施策シート用 IMAGE 関数（ファイルはURLのみで軽い）。"""
    u = normalize_image_url(url)
    if not u:
        return ""
    return f'=IMAGE("{u.replace(chr(34), chr(34) + chr(34))}")'


def master_values_row(d: Dict[str, Any]) -> List[Any]:
    return [d.get(h, "") for h in MASTER_HEADERS]


def lookup_cost_by_asin(svc) -> Dict[str, Dict[str, str]]:
    N = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=PRODUCT_SS, range=f"'{PRODUCT_SHEET}'!N8:N8000")
        .execute()
        .get("values")
        or []
    )
    U = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=PRODUCT_SS, range=f"'{PRODUCT_SHEET}'!U8:U8000")
        .execute()
        .get("values")
        or []
    )
    AK = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=PRODUCT_SS, range=f"'{PRODUCT_SHEET}'!AK8:AK8000")
        .execute()
        .get("values")
        or []
    )

    def g(arr, i):
        return str(arr[i][0]).strip() if i < len(arr) and arr[i] else ""

    out: Dict[str, Dict[str, str]] = {}
    for i in range(len(N)):
        asin = g(N, i).upper()
        if asin.startswith("B0"):
            out[asin] = {"原価U": g(U, i), "商品マスタSKU": g(AK, i)}
    return out


def index_template(rows: List[Dict[str, Any]]):
    by_sku: Dict[str, List[Dict[str, Any]]] = {}
    by_asin: Dict[str, List[Dict[str, Any]]] = {}
    for r in rows:
        sku = str(r.get("sku") or "").strip()
        asin = str(r.get("deal_asin") or "").strip().upper()
        if sku:
            by_sku.setdefault(sku, []).append(r)
        if asin:
            by_asin.setdefault(asin, []).append(r)
    return by_sku, by_asin


def collect_roster_from_master(svc, sid: str) -> List[Dict[str, Any]]:
    """タイムセール_マスタの有効行のみ（対象名簿の正本）。"""
    _h, rows = read_sheet_rows(svc, sid, MASTER_SHEET)
    if _h == ["案内"] or not _h:
        return []
    out: List[Dict[str, Any]] = []
    seen = set()
    for r in rows:
        sku = str(r.get("SKU") or "").strip()
        asin = str(r.get("ASIN") or "").strip().upper()
        key = sku or asin
        if not key or key in seen:
            continue
        enabled = r.get("有効")
        if str(enabled or "").strip() == "":
            enabled = "TRUE"
        if not _truthy(enabled):
            continue
        seen.add(key)
        out.append(dict(r, SKU=sku, ASIN=asin, 有効=enabled))
    return out


def bootstrap_master_from_sales_if_empty(svc, sid: str) -> None:
    """マスタが空／案内のみ／旧ヘッダなら、名簿を起こすまたは列移行。"""
    ensure_sheet(svc, sid, MASTER_SHEET)
    mh, mrows = read_sheet_rows(svc, sid, MASTER_SHEET)
    if mh == MASTER_HEADERS and mrows:
        return
    if mh and mh != ["案内"] and mrows:
        # 旧ヘッダ → 新ヘッダ（画像URL含む）へ移行
        migrated = []
        for r in mrows:
            sku = str(r.get("SKU") or "").strip()
            asin = str(r.get("ASIN") or "").strip().upper()
            if not sku and not asin:
                continue
            if sku.startswith("B0") and not asin:
                asin, sku = sku, ""
            migrated.append(
                master_values_row(
                    {
                        "SKU": sku,
                        "ASIN": asin,
                        "親ASIN": str(r.get("親ASIN") or ""),
                        "商品名": str(r.get("商品名") or ""),
                        "画像URL": normalize_image_url(r.get("画像URL") or r.get("画像")),
                        "marketplace": str(r.get("marketplace") or "JP") or "JP",
                        "通貨": str(r.get("通貨") or "JPY") or "JPY",
                        "有効": str(r.get("有効") or "TRUE") or "TRUE",
                        "出品者価格_SC": str(r.get("出品者価格_SC") or ""),
                        "タイムセール価格_SC": str(r.get("タイムセール価格_SC") or ""),
                        "販売商品数_SC": str(r.get("販売商品数_SC") or ""),
                        "V30": str(r.get("V30") or ""),
                        "Q_fba": str(r.get("Q_fba") or ""),
                        "原価U": str(r.get("原価U") or ""),
                        **point_fields_from_row(r),
                        **recovery_fields_from_row(r),
                        **a_track_fields_from_row(r),
                        "メモ": str(r.get("メモ") or ""),
                    }
                )
            )
        if migrated:
            write_headers_and_rows(svc, sid, MASTER_SHEET, MASTER_HEADERS, migrated, clear=True)
            return

    # 施策から起こす
    _sh, sales = read_sheet_rows(svc, sid, SALE_SHEET)
    seen = set()
    rows = []
    for r in sales:
        sku = str(r.get("SKU") or "").strip()
        asin = str(r.get("ASIN") or "").strip().upper()
        key = sku or asin
        if not key or key in seen:
            continue
        if str(r.get("sale_id") or "") == "" and str(r.get("レーン") or "") == "" and not sku and not asin:
            continue
        seen.add(key)
        rows.append(
            master_values_row(
                {
                    "SKU": sku,
                    "ASIN": asin,
                    "親ASIN": str(r.get("親ASIN") or ""),
                    "商品名": str(r.get("商品名") or ""),
                    "画像URL": normalize_image_url(r.get("画像URL") or r.get("画像")),
                    "marketplace": "JP",
                    "通貨": "JPY",
                    "有効": "TRUE",
                    "出品者価格_SC": str(r.get("出品者価格_SC") or ""),
                    "タイムセール価格_SC": str(r.get("タイムセール価格_SC") or ""),
                    "販売商品数_SC": str(r.get("販売商品数_SC") or ""),
                    "原価U": str(r.get("原価U") or ""),
                    "メモ": "施策シートから名簿復元",
                }
            )
        )
    write_headers_and_rows(svc, sid, MASTER_SHEET, MASTER_HEADERS, rows, clear=True)
    LOG.info("マスタを施策から復元: %s件", len(rows))


def empty_row() -> Dict[str, Any]:
    return {h: "" for h in SALE_HEADERS}


def ensure_unified_schema(svc, sid: str) -> None:
    ensure_sheet(svc, sid, MASTER_SHEET)
    ensure_sheet(svc, sid, SALE_SHEET)
    ensure_sheet(svc, sid, ANALYSIS_SHEET)
    bootstrap_master_from_sales_if_empty(svc, sid)

    ah, _ = read_sheet_rows(svc, sid, ANALYSIS_SHEET)
    if not ah or ah == ["案内"]:
        write_headers_and_rows(
            svc,
            sid,
            ANALYSIS_SHEET,
            ["説明"],
            [
                ["ソース: タイムセール（施策）／対象はタイムセール_マスタ"],
                ["B=名付き公式のみバルク／A=月・カスタム相当はAPI自動化"],
                ["並び: 開始日の降順 → 商品名の昇順"],
                ["画像: マスタ=画像URL／施策=商品名直後のIMAGE関数"],
            ],
            clear=True,
        )
    for title, msg in (
        (LEGACY_EXEC, "廃止: 施策は「タイムセール」。対象名簿は「タイムセール_マスタ」。"),
        (LEGACY_LANE_A, "廃止: レーンAも「タイムセール」に統合。"),
    ):
        try:
            ensure_sheet(svc, sid, title)
            write_headers_and_rows(svc, sid, title, ["案内"], [[msg]], clear=True)
        except Exception as e:
            LOG.warning("legacy note skip %s: %s", title, e)

    # 列順・別名移行 → ヘッダ色・人入力黄・ポイント表示・戻しプルダウン
    try:
        ensure_headers_append(svc, sid, SALE_SHEET, list(SALE_HEADERS))
        realigned = realign_master_to_schema(svc, sid)
        if not realigned:
            ensure_headers_append(svc, sid, MASTER_SHEET, list(MASTER_HEADERS))
        n = apply_master_header_group_colors(svc, sid, MASTER_SHEET)
        y = apply_master_human_input_yellow(svc, sid, MASTER_SHEET)
        f = apply_master_display_formulas(svc, sid, MASTER_SHEET)
        apply_master_point_unit_formats(svc, sid, MASTER_SHEET)
        apply_master_recovery_validations(svc, sid, MASTER_SHEET)
        LOG.info(
            "master realign=%s header_colors=%s human_yellow_cols=%s display_formula_cells=%s validations=ok",
            realigned,
            n,
            y,
            f,
        )
    except Exception as e:
        LOG.warning("master header format skip: %s", e)


def _normalize_master_row_keys(r: Dict[str, Any]) -> Dict[str, Any]:
    """旧列名を現行名に寄せる（最終売価円→目標売価円 等）。"""
    out: Dict[str, Any] = {}
    for k, v in r.items():
        nk = normalize_master_header_name(str(k or "").strip())
        if not nk:
            continue
        cur = out.get(nk)
        if cur is not None and str(cur).strip() != "" and str(v or "").strip() == "":
            continue
        if (
            nk in out
            and str(out.get(nk) or "").strip()
            and str(k or "").strip() in MASTER_HEADER_ALIASES
        ):
            continue
        out[nk] = v
    return out


def realign_master_to_schema(svc, sid: str) -> bool:
    """
    マスタ列を MASTER_HEADERS 順に揃える（別名移行・欠落列追加・余分列は末尾保全）。
    並び／別名が既に一致なら False。書き換えたら True。
    """
    mh, mrows = read_sheet_rows(svc, sid, MASTER_SHEET)
    if not mh or mh == ["案内"]:
        return False
    extras: List[str] = []
    for h in mh:
        raw = str(h or "").strip()
        if not raw:
            continue
        if raw in MASTER_HEADER_ALIASES:
            continue
        if raw in MASTER_HEADERS_DROPPED:
            continue
        nk = normalize_master_header_name(raw)
        if nk in MASTER_HEADERS_DROPPED:
            continue
        if nk not in MASTER_HEADERS and raw not in extras:
            extras.append(raw)
    want = list(MASTER_HEADERS) + extras
    if mh == want and all(
        normalize_master_header_name(h) == h for h in mh if str(h or "").strip()
    ):
        return False
    if mh == MASTER_HEADERS and not extras:
        return False

    values: List[List[Any]] = []
    for r in mrows:
        nr = _normalize_master_row_keys(r)
        nr.update(point_fields_from_row(nr))
        nr.update(recovery_fields_from_row(nr))
        nr.update(a_track_fields_from_row(nr))
        row = [nr.get(h, "") if nr.get(h) is not None else "" for h in MASTER_HEADERS]
        for h in extras:
            val = nr.get(h, r.get(h, ""))
            row.append(val if val is not None else "")
        if any(str(c).strip() for c in row):
            values.append(row)

    write_headers_and_rows(svc, sid, MASTER_SHEET, want, values, clear=True)
    LOG.info(
        "master realigned: cols=%s rows=%s extras=%s (from %s)",
        len(want),
        len(values),
        extras,
        mh,
    )
    return True


def fill_master_basics(
    roster: List[Dict[str, Any]],
    by_sku,
    by_asin,
    costs: Dict[str, Dict[str, str]],
    v30_map: Dict[str, float],
) -> List[List[Any]]:
    """マスタ行を②＋原価＋画像URL＋V30/Q_fbaで埋めた二次元配列を返す。"""
    out = []
    for r in roster:
        sku = str(r.get("SKU") or "").strip()
        asin = str(r.get("ASIN") or "").strip().upper()
        cands = (by_sku.get(sku) if sku else None) or by_asin.get(asin) or []
        enabled = r.get("有効") or "TRUE"
        keep_url = normalize_image_url(r.get("画像URL") or r.get("画像"))
        keep_v30 = r.get("V30")
        try:
            keep_v30_f = float(str(keep_v30).replace(",", "")) if str(keep_v30 or "").strip() else None
        except ValueError:
            keep_v30_f = None
        if cands:
            hit = prefer_row(cands)
            if not sku:
                sku = str(hit.get("sku") or "").strip()
            if not asin:
                asin = str(hit.get("deal_asin") or "").strip().upper()
            cost = costs.get(asin, {}).get("原価U") or str(r.get("原価U") or "")
            memo = str(r.get("メモ") or "")
            msku = costs.get(asin, {}).get("商品マスタSKU") or ""
            if msku and "商品マスタSKU" not in memo:
                memo = (memo + " " if memo else "") + f"商品マスタSKU={msku}"
            img = normalize_image_url(hit.get("image_url")) or keep_url
            q_fba = hit.get("seller_quantity")
            if q_fba is None or str(q_fba).strip() == "":
                q_fba = r.get("Q_fba") or ""
            v30 = v30_map.get(asin)
            if v30 is None:
                v30 = keep_v30_f
            pts = point_fields_from_row(r)
            rec = recovery_fields_from_row(r)
            out.append(
                master_values_row(
                    {
                        "SKU": sku,
                        "ASIN": asin,
                        "親ASIN": str(hit.get("parent_asin") or r.get("親ASIN") or ""),
                        "商品名": str(hit.get("product_name") or r.get("商品名") or ""),
                        "画像URL": img,
                        "marketplace": "JP",
                        "通貨": "JPY",
                        "有効": enabled,
                        "出品者価格_SC": hit.get("seller_price")
                        if hit.get("seller_price") is not None
                        else r.get("出品者価格_SC") or "",
                        "タイムセール価格_SC": hit.get("deal_price")
                        if hit.get("deal_price") is not None
                        else r.get("タイムセール価格_SC") or "",
                        "販売商品数_SC": hit.get("committed_units")
                        if hit.get("committed_units") is not None
                        else r.get("販売商品数_SC") or "",
                        "V30": v30 if v30 is not None else "",
                        "Q_fba": q_fba if q_fba is not None else "",
                        "原価U": cost,
                        **pts,
                        **rec,
                        **a_track_fields_from_row(r),
                        "メモ": memo,
                    }
                )
            )
        else:
            cost = costs.get(asin, {}).get("原価U") or str(r.get("原価U") or "")
            v30 = v30_map.get(asin)
            if v30 is None:
                v30 = keep_v30_f
            out.append(
                master_values_row(
                    {
                        "SKU": sku,
                        "ASIN": asin,
                        "親ASIN": str(r.get("親ASIN") or ""),
                        "商品名": str(r.get("商品名") or ""),
                        "画像URL": keep_url,
                        "marketplace": str(r.get("marketplace") or "JP") or "JP",
                        "通貨": str(r.get("通貨") or "JPY") or "JPY",
                        "有効": enabled,
                        "出品者価格_SC": str(r.get("出品者価格_SC") or ""),
                        "タイムセール価格_SC": str(r.get("タイムセール価格_SC") or ""),
                        "販売商品数_SC": str(r.get("販売商品数_SC") or ""),
                        "V30": v30 if v30 is not None else "",
                        "Q_fba": str(r.get("Q_fba") or ""),
                        "原価U": cost,
                        **point_fields_from_row(r),
                        **recovery_fields_from_row(r),
                        **a_track_fields_from_row(r),
                        "メモ": str(r.get("メモ") or "②テンプレに無し"),
                    }
                )
            )
    return out


def load_opt_in_keys_from_audits(cfg: dict) -> set:
    """過去の build_submit 監査から opt_in 済み SKU||スケジュール を拾う（再UL防止・台帳復元用）。"""
    keys = set()
    paths = list((HERE / "_work").glob("DEALS_*_audit.json"))
    try:
        folder03 = folder_path(cfg, "03")
        paths.extend(folder03.glob("*_audit.json"))
    except Exception:
        pass
    for p in paths:
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        for a in data.get("audits") or []:
            if a.get("action") != "opt_in":
                continue
            sku = str(a.get("sku") or "").strip()
            sch = str(a.get("schedule") or "").strip()
            if sku and sch:
                keys.add(f"{sku}||{sch}")
    return keys


def sync(cfg: dict, *, source: Optional[Path], write: bool) -> int:
    folder02 = folder_path(cfg, "02")
    src = source or latest_xlsx(folder02)
    if not src or not src.is_file():
        LOG.error("②にxlsxがありません: %s", folder02)
        return 1

    sid = str(cfg.get("ads_spreadsheet_id") or "").strip()
    svc = sheets_service(write=True)
    ensure_unified_schema(svc, sid)

    roster = collect_roster_from_master(svc, sid)
    if not roster:
        LOG.error(
            "タイムセール_マスタに有効なSKU/ASINがありません。名簿に追加してから再実行してください。"
        )
        return 1

    t_rows, colmap, wb = read_template_rows(src)
    by_sku, by_asin = index_template(t_rows)
    cmap = load_column_map()
    try:
        template_ws = find_template_sheet(
            wb, list(cmap.get("template_sheet_name_candidates") or [])
        )
    except Exception:
        template_ws = None
    schedule_col = int(colmap.get("schedule") or 0)
    # カタログは監視・名付き日付補完用。新規登録は②SKUのセル＋ドロップダウン日付付きが正
    all_sched = merge_schedules(
        collect_schedules(t_rows),
        collect_schedule_catalog(wb),
    )
    costs = lookup_cost_by_asin(svc)
    asin_list = [str(r.get("ASIN") or "").strip().upper() for r in roster]
    v30_map = resolve_v30_map(asin_list, master_rows=roster, use_spapi=True)
    LOG.info(
        "master_roster=%s template=%s catalog_schedules=%s v30=%s",
        len(roster),
        len(t_rows),
        len(all_sched),
        len(v30_map),
    )

    master_values = fill_master_basics(roster, by_sku, by_asin, costs, v30_map)

    _h, existing = read_sheet_rows(svc, sid, SALE_SHEET)
    # §2.1.1: B公式は台帳。予定・登録済を消さない（提出対象外≠行削除）
    # 2026-08-11: レーンAは運用しない → 既存A行は保持せず削除
    B_LEDGER_STATES = {
        "予定",
        "要確認",
        "数量改定済",
        "延期",
        "UL済",
        "アップロード済",
        "実施中",
        "終了",
        "失敗",
        "見送り",
    }
    # マスタ画像URL索引（B行の画像空欄を埋める）
    master_img: Dict[str, str] = {}
    for mv in master_values:
        msku, masin = str(mv[0] or "").strip(), str(mv[1] or "").strip().upper()
        url = normalize_image_url(mv[4] if len(mv) > 4 else "")
        if msku and url:
            master_img[msku] = url
        if masin and url:
            master_img[masin] = url

    kept: List[Dict[str, Any]] = []
    dropped_a = 0
    for r in existing:
        lane = str(r.get("レーン") or "").strip()
        st = str(r.get("状態") or "").strip()
        if lane == LANE_A or lane.startswith("A_"):
            dropped_a += 1
            continue
        if lane == LANE_B:
            if st in B_LEDGER_STATES or st == "":
                r = dict(r)
                if not st:
                    r["状態"] = "予定"
                # 画像が空ならマスタ／テンプレから埋める
                if not str(r.get("画像") or "").strip():
                    sku_k = str(r.get("SKU") or "").strip()
                    asin_k = str(r.get("ASIN") or "").strip().upper()
                    url = master_img.get(sku_k) or master_img.get(asin_k) or ""
                    if not url:
                        cands0 = (by_sku.get(sku_k) if sku_k else None) or by_asin.get(asin_k) or []
                        if cands0:
                            url = normalize_image_url(prefer_row(cands0).get("image_url"))
                    if url:
                        r["画像"] = image_formula(url)
                kept.append(r)
    if dropped_a:
        LOG.info("レーンA行を削除（運用停止）: %s", dropped_a)

    def _sku_sched_key(r: Dict[str, Any]) -> str:
        return f"{str(r.get('SKU') or '').strip()}||{str(r.get('スケジュール') or '').strip()}"

    kept_keys = {_sku_sched_key(r) for r in kept if _sku_sched_key(r) != "||"}
    submitted_keys = load_opt_in_keys_from_audits(cfg)
    LOG.info("ledger kept=%s prior_opt_in_keys=%s", len(kept), len(submitted_keys))

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_id = datetime.now().strftime("SYNC_%Y%m%d_%H%M%S")
    new_rows: List[Dict[str, Any]] = []
    clip_alerts: List[Dict[str, Any]] = []
    # limit_bulk_custom: 0以下は実質無制限（ドロップダウン日付付きを全て拾う）
    lim_custom = int(cfg.get("limit_bulk_custom") if cfg.get("limit_bulk_custom") is not None else 20)
    if lim_custom <= 0:
        lim_custom = 999

    for item in roster:
        sku = str(item.get("SKU") or "").strip()
        asin = str(item.get("ASIN") or "").strip().upper()
        # 埋め後の値を優先
        for mv in master_values:
            if (sku and mv[0] == sku) or (asin and mv[1] == asin):
                sku, asin = mv[0], mv[1]
                break
        cands = (by_sku.get(sku) if sku else None) or by_asin.get(asin) or []
        if not cands:
            row = empty_row()
            keep_url = normalize_image_url(item.get("画像URL") or item.get("画像"))
            for mv in master_values:
                if (sku and mv[0] == sku) or (asin and mv[1] == asin):
                    # MASTER: SKU, ASIN, 親, 商品名, 画像URL
                    keep_url = normalize_image_url(mv[4]) or keep_url
                    break
            row.update(
                {
                    "sale_id": _sale_id("ROSTER", sku or asin, "missing"),
                    "レーン": LANE_A,
                    "SKU": sku,
                    "ASIN": asin,
                    "商品名": str(item.get("商品名") or ""),
                    "画像": image_formula(keep_url),
                    "marketplace": "JP",
                    "通貨": "JPY",
                    "有効": "TRUE",
                    "状態": "要確認",
                    "提出対象": "いいえ",
                    "更新日時": now,
                    "runId": run_id,
                    "メッセージ": "②テンプレに無し（マスタ登録済）",
                    "原価U": costs.get(asin, {}).get("原価U", ""),
                }
            )
            new_rows.append(row)
            continue

        hit = prefer_row(cands)
        if not sku:
            sku = str(hit.get("sku") or "").strip()
        if not asin:
            asin = str(hit.get("deal_asin") or "").strip().upper()
        cost = costs.get(asin, {}).get("原価U", "")
        parent = str(hit.get("parent_asin") or "")
        pname = str(hit.get("product_name") or "")
        img = image_formula(
            normalize_image_url(hit.get("image_url"))
            or normalize_image_url(item.get("画像URL") or item.get("画像"))
        )
        seller_p = hit.get("seller_price")
        deal_p = hit.get("deal_price")
        qty_sc = hit.get("committed_units")
        q_fba = hit.get("seller_quantity")
        v30_val = v30_map.get(asin)
        for mv in master_values:
            if (sku and mv[0] == sku) or (asin and mv[1] == asin):
                if v30_val is None and str(mv[11] or "").strip() != "":
                    try:
                        v30_val = float(str(mv[11]).replace(",", ""))
                    except ValueError:
                        pass
                if (q_fba is None or str(q_fba).strip() == "") and str(mv[12] or "").strip() != "":
                    q_fba = mv[12]
                break

        local_cells = collect_schedules(cands)
        local_dd: List[Dict[str, Any]] = []
        if template_ws is not None and schedule_col:
            local_dd = collect_dated_dropdown_schedules(
                wb,
                template_ws,
                cands,
                schedule_col=schedule_col,
                osusume_only=True,
            )
        local = merge_schedules(local_cells, local_dd)
        catalog_named = pick_named_within_horizon(
            all_sched,
            b_horizon_days=int(cfg.get("b_horizon_days") or 90),
            limit_b=int(cfg.get("limit_b") or 2),
        )
        skip_names = {
            str(r.get("スケジュール") or "").strip()
            for r in kept
            if str(r.get("SKU") or "").strip() == sku
            and str(r.get("状態") or "").strip() in {"UL済", "実施中", "終了"}
        }
        existing_names = {
            str(r.get("スケジュール") or "").strip()
            for r in kept
            if str(r.get("SKU") or "").strip() == sku
        }
        b_picks, bulk_customs, _a_unused = pick_schedules_sku_local(
            local,
            limit_b=int(cfg.get("limit_b") or 2),
            limit_a=0,  # レーンA運用停止
            limit_bulk_custom=lim_custom,
            ab_gap_days=int(cfg.get("ab_gap_days") or 0),
            a_max_days=int(cfg.get("a_max_days") or 14),
            b_horizon_days=int(cfg.get("b_horizon_days") or 90),
            early_fee_deadlines=cfg.get("early_fee_deadlines"),
            skip_schedule_names=skip_names | existing_names,
        )
        # ドロップダウン／セルに出ている名付きのみ新規追加（カタログだけは追加しない）
        for p in catalog_named:
            name = str(p.get("schedule") or "").strip()
            if not name or name in skip_names or name in existing_names:
                continue
            if any(str(x.get("schedule") or "") == name for x in b_picks):
                continue
            dd_names = {str(x.get("schedule") or "").strip() for x in local_dd}
            cell_names = {str(x.get("schedule") or "").strip() for x in local_cells}
            if name in dd_names or name in cell_names:
                b_picks.append(p)
        # 名付き確定後にカスタムを公式優先で再クリップ
        blockers = []
        for b in b_picks:
            s, e = parse_ymd(b.get("start")), parse_ymd(b.get("end"))
            if s and e:
                blockers.append((s, e))
        # 台帳上の既存名付きもブロッカーに含める
        for r in kept:
            if str(r.get("SKU") or "").strip() != sku:
                continue
            if not is_official_b_schedule(str(r.get("スケジュール") or "")):
                continue
            s, e = parse_ymd(r.get("開始日")), parse_ymd(r.get("終了日"))
            if s and e:
                blockers.append((s, e))
        reclipped = []
        for c in bulk_customs:
            s, e = parse_ymd(c.get("start")), parse_ymd(c.get("end"))
            if not s or not e:
                continue
            sh = shrink_range_avoiding_blockers(s, e, blockers)
            if not sh:
                clip_alerts.append(
                    {
                        "SKU": sku,
                        "ASIN": asin,
                        "スケジュール": c.get("schedule"),
                        "旧開始": s.isoformat(),
                        "旧終了": e.isoformat(),
                        "新開始": "",
                        "新終了": "",
                        "状態": "見送り",
                        "理由": "名付き公式と全期間重複のため独自セールを見送り",
                    }
                )
                continue
            cs, ce = sh
            entry = dict(c)
            entry["start"], entry["end"] = cs.isoformat(), ce.isoformat()
            if (cs, ce) != (s, e):
                entry["clipped"] = True
                clip_alerts.append(
                    {
                        "SKU": sku,
                        "ASIN": asin,
                        "スケジュール": c.get("schedule"),
                        "旧開始": s.isoformat(),
                        "旧終了": e.isoformat(),
                        "新開始": cs.isoformat(),
                        "新終了": ce.isoformat(),
                        "状態": "予定",
                        "理由": "名付き公式優先で独自セール期間を短縮（新規）",
                    }
                )
            reclipped.append(entry)
            blockers.append((cs, ce))
        bulk_customs = reclipped
        a_picks: List[Dict[str, Any]] = []  # レーンA運用停止（合成しない）
        LOG.info(
            "SKU=%s cell=%s dropdown=%s catalog_near=%s B=%s bulk_custom=%s A=disabled",
            sku,
            [x.get("schedule") for x in local_cells],
            [x.get("schedule") for x in local_dd],
            [x.get("schedule") for x in catalog_named],
            [x.get("schedule") for x in b_picks],
            [x.get("schedule") for x in bulk_customs],
        )

        b_rows_meta = []

        def _price_from_pick(p: dict):
            dp = p.get("deal_price")
            if dp is not None and str(dp).strip() != "":
                return dp
            return deal_p

        def _qty_sc_from_pick(p: dict):
            q = p.get("committed_units")
            if q is not None and str(q).strip() != "":
                return q
            return qty_sc

        def _append_b_row(p: dict, *, msg_prefix: str, registered: bool = False):
            name = str(p.get("schedule") or "").strip()
            key = f"{sku}||{name}"
            if key in kept_keys or not name:
                return
            has_dates = bool(p.get("start") and p.get("end"))
            d_days = deal_day_count(p.get("start"), p.get("end"))
            row_deal = _price_from_pick(p)
            row_qty_sc = _qty_sc_from_pick(p)
            qd = compute_q_deal(
                v30=v30_val,
                d_days=d_days,
                schedule=name,
                q_fba=q_fba,
            )
            msg = f"{msg_prefix} | qty: {qd['note']}"
            if registered:
                st = "UL済"
                submit = "いいえ"
                msg = f"登録済・台帳保全 | {msg}"
            else:
                st = "予定" if has_dates else "要確認"
                submit = "いいえ" if qd.get("deferred") else "はい"
                if qd.get("need_v30"):
                    st = "要確認"
                    msg += " | V30未取得"
                if qd.get("deferred"):
                    st = "延期"
            row = empty_row()
            row.update(
                {
                    "sale_id": _sale_id(LANE_B, sku, name),
                    "レーン": LANE_B,
                    "SKU": sku,
                    "ASIN": asin,
                    "親ASIN": parent,
                    "商品名": pname,
                    "画像": img,
                    "marketplace": "JP",
                    "通貨": "JPY",
                    "有効": "FALSE" if (qd.get("deferred") and not registered) else "TRUE",
                    "種別": "おすすめタイムセール",
                    "スケジュール": name,
                    "開始日": p.get("start") or "",
                    "終了日": p.get("end") or "",
                    "出品者価格_SC": seller_p if seller_p is not None else "",
                    "タイムセール価格_SC": row_deal if row_deal is not None else "",
                    "タイムセール価格_確定": row_deal if row_deal is not None else "",
                    "販売商品数_SC": row_qty_sc if row_qty_sc is not None else "",
                    "V30": qd["V30"] if qd["V30"] is not None else "",
                    "販売商品数_確定": qd["Q_deal"],
                    "原価U": cost,
                    "提出対象": submit,
                    "状態": st,
                    "更新日時": now,
                    "runId": run_id,
                    "メッセージ": msg,
                }
            )
            new_rows.append(row)
            b_rows_meta.append(row)
            kept_keys.add(key)

        for p in b_picks:
            name = str(p.get("schedule") or "").strip()
            in_dd = any(str(x.get("schedule") or "") == name for x in local_dd)
            in_cell = any(str(x.get("schedule") or "") == name for x in local_cells)
            key = f"{sku}||{name}"
            registered = key in submitted_keys or name in skip_names
            if in_dd:
                msg = "名付き公式・②ドロップダウン"
            elif in_cell:
                msg = "名付き公式・②セル"
            else:
                msg = "名付き公式"
            _append_b_row(p, msg_prefix=msg, registered=registered)

        for p in bulk_customs:
            src = str(p.get("source") or "")
            msg = (
                "②ドロップダウン日付付き（カスタム/月）"
                if src == "dropdown"
                else "②セル／バルク（カスタム/月）"
            )
            if p.get("clipped"):
                msg += " | 公式優先で期間短縮"
            _append_b_row(p, msg_prefix=msg, registered=False)

        # 台帳の既存独自セールを公式ブロッカーで再調整
        named_blockers: List[Tuple[date, date]] = []
        for r in kept:
            if str(r.get("SKU") or "").strip() != sku:
                continue
            if not is_official_b_schedule(str(r.get("スケジュール") or "")):
                continue
            s, e = parse_ymd(r.get("開始日")), parse_ymd(r.get("終了日"))
            if s and e:
                named_blockers.append((s, e))
        for p in b_picks:
            s, e = parse_ymd(p.get("start")), parse_ymd(p.get("end"))
            if s and e:
                named_blockers.append((s, e))
        for r in kept:
            if str(r.get("SKU") or "").strip() != sku:
                continue
            sch = str(r.get("スケジュール") or "").strip()
            if not sch or is_official_b_schedule(sch):
                continue
            if not is_a_equivalent_schedule(sch) and not str(sch).startswith("カスタム"):
                # 月／カスタム以外は触らない
                if not (sch.startswith("月") or "カスタム" in sch):
                    continue
            s0, e0 = parse_ymd(r.get("開始日")), parse_ymd(r.get("終了日"))
            if not s0 or not e0:
                continue
            if not any(ranges_overlap(s0, e0, b0, b1) for b0, b1 in named_blockers):
                continue
            sh = shrink_range_avoiding_blockers(s0, e0, named_blockers)
            old_s, old_e = s0.isoformat(), e0.isoformat()
            if not sh:
                r["状態"] = "見送り"
                r["提出対象"] = "いいえ"
                r["有効"] = "FALSE"
                r["メッセージ"] = "公式と全期間重複→独自セール見送り（要SC確認）"
                r["更新日時"] = now
                r["runId"] = run_id
                clip_alerts.append(
                    {
                        "SKU": sku,
                        "ASIN": asin,
                        "スケジュール": sch,
                        "旧開始": old_s,
                        "旧終了": old_e,
                        "新開始": "",
                        "新終了": "",
                        "状態": "見送り",
                        "理由": "名付き公式と全期間重複のため台帳の独自セールを見送り",
                    }
                )
                continue
            cs, ce = sh
            if (cs, ce) == (s0, e0):
                continue
            r["開始日"] = cs.isoformat()
            r["終了日"] = ce.isoformat()
            r["メッセージ"] = "公式優先で独自セール期間を短縮（台帳再調整・要SC確認）"
            r["更新日時"] = now
            r["runId"] = run_id
            # UL済でも日付変更を提出対象に戻す（人の再UL判断用）
            if str(r.get("状態") or "").strip() in {"UL済", "アップロード済", "予定", "要確認"}:
                r["提出対象"] = "はい"
            clip_alerts.append(
                {
                    "SKU": sku,
                    "ASIN": asin,
                    "スケジュール": sch,
                    "旧開始": old_s,
                    "旧終了": old_e,
                    "新開始": cs.isoformat(),
                    "新終了": ce.isoformat(),
                    "状態": str(r.get("状態") or ""),
                    "理由": "名付き公式優先で台帳の独自セール期間を短縮",
                }
            )

        for p in a_picks:
            akey = f"{sku}||{str(p.get('schedule') or '').strip()}"
            if akey in kept_keys:
                continue
            try:
                normal = float(seller_p) if seller_p is not None else 0.0
            except (TypeError, ValueError):
                normal = 0.0
            try:
                cval = float(str(cost).replace(",", "")) if cost else 0.0
            except ValueError:
                cval = 0.0
            sale = round(normal * 0.92) if normal else ""
            d_days = deal_day_count(p.get("start"), p.get("end"))
            qd = compute_q_deal(
                v30=v30_val,
                d_days=d_days or 14,
                schedule=str(p.get("schedule") or ""),
                q_fba=q_fba,
            )
            q = qd["Q_deal"]
            profit = int((float(sale) - cval) * q) if sale and cval else ""
            row = empty_row()
            row.update(
                {
                    "sale_id": _sale_id(LANE_A, sku, p["schedule"] or f"{p.get('start')}-{p.get('end')}"),
                    "レーン": LANE_A,
                    "SKU": sku,
                    "ASIN": asin,
                    "親ASIN": parent,
                    "商品名": pname,
                    "画像": img,
                    "marketplace": "JP",
                    "通貨": "JPY",
                    "有効": "TRUE",
                    "承認済": "FALSE",
                    "スケジュール": p["schedule"],
                    "開始日": p.get("start") or "",
                    "終了日": p.get("end") or "",
                    "出品者価格_SC": seller_p if seller_p is not None else "",
                    "タイムセール価格_SC": deal_p if deal_p is not None else "",
                    "通常価格": normal if normal else "",
                    "セール価格": sale,
                    "販売商品数_SC": qty_sc if qty_sc is not None else "",
                    "V30": qd["V30"] if qd["V30"] is not None else "",
                    "販売商品数_確定": q,
                    "目標販売": q,
                    "想定利益": profit,
                    "原価U": cost,
                    "提出対象": "いいえ",
                    "状態": "下書き",
                    "更新日時": now,
                    "runId": run_id,
                    "メッセージ": (
                        f"B固定後の空きA（最大14日・gap={p.get('ab_gap_days', cfg.get('ab_gap_days', 0))}）"
                        f" | qty: {qd['note']}"
                    ),
                    "メモ": "P1aは承認後API。A↔B間隔は§10.8で検証。Bと重なれば停止",
                }
            )
            a0, a1 = parse_ymd(row["開始日"]), parse_ymd(row["終了日"])
            for br in b_rows_meta:
                if ranges_overlap(
                    a0, a1, parse_ymd(br.get("開始日")), parse_ymd(br.get("終了日"))
                ):
                    row["有効"] = "FALSE"
                    row["状態"] = "停止"
                    row["メッセージ"] = "名付き公式Bと期間重なり→A停止（B優先）"
                    break
            # 台帳上の既存Bとも重なりチェック
            for br in kept:
                if str(br.get("SKU") or "").strip() != sku:
                    continue
                if str(br.get("レーン") or "") != LANE_B:
                    continue
                if ranges_overlap(
                    a0, a1, parse_ymd(br.get("開始日")), parse_ymd(br.get("終了日"))
                ):
                    row["有効"] = "FALSE"
                    row["状態"] = "停止"
                    row["メッセージ"] = "台帳Bと期間重なり→A停止（B優先）"
                    break
            new_rows.append(row)
            kept_keys.add(akey)

    def sale_sort_key(r: Dict[str, Any]):
        # 開始日降順（日付なしは末尾）→ 商品名昇順
        d = parse_ymd(r.get("開始日"))
        date_rank = -(d.toordinal()) if d else float("inf")
        return (date_rank, str(r.get("商品名") or ""), str(r.get("sale_id") or ""))

    def normalize_row_dates_(r: Dict[str, Any]) -> Dict[str, Any]:
        out = dict(r)
        for col in ("開始日", "終了日"):
            ymd = format_ymd(out.get(col))
            if ymd:
                out[col] = ymd
        return out

    final = [
        normalize_row_dates_(r)
        for r in sorted(kept + new_rows, key=sale_sort_key)
    ]

    # 不変条件（並び・Aなし・画像・日付）— 壊したらここで落とす
    try:
        from invariants import check_sale_sheet_invariants

        inv_errs = check_sale_sheet_invariants(final)
        for msg in inv_errs:
            LOG.error("INVARIANT: %s", msg)
        if inv_errs and write:
            LOG.error("不変条件違反のためシート書き込みを中止（%s件）", len(inv_errs))
            return 1
    except Exception as e:
        LOG.warning("invariants check skipped: %s", e)

    LOG.info(
        "roster=%s sales_new=%s kept=%s B=%s A=%s",
        len(roster),
        len(new_rows),
        len(kept),
        sum(1 for r in new_rows if r.get("レーン") == LANE_B),
        sum(1 for r in new_rows if r.get("レーン") == LANE_A),
    )

    if not write:
        LOG.info("dry-run。本番は --write")
        return 0

    write_headers_and_rows(svc, sid, MASTER_SHEET, MASTER_HEADERS, master_values, clear=True)
    values = [[r.get(h, "") for h in SALE_HEADERS] for r in final]
    write_headers_and_rows(svc, sid, SALE_SHEET, SALE_HEADERS, values, clear=True)
    LOG.info(
        "更新完了 master=%s sales=%s runId=%s",
        MASTER_SHEET,
        SALE_SHEET,
        run_id,
    )
    if clip_alerts:
        try:
            from notify_mail import build_custom_clip_alert_body, send_notify

            subj = "[amazonタイムセール] 独自セール日時調整（公式と衝突）"
            body = build_custom_clip_alert_body(clip_alerts)
            if send_notify(cfg, subj, body):
                LOG.info("衝突調整メール送信OK（%s件）", len(clip_alerts))
            else:
                LOG.warning("衝突調整メール送信失敗。件数=%s", len(clip_alerts))
                LOG.info("ALERT_BODY:\n%s", body[:2000])
        except Exception as e:
            LOG.warning("衝突調整メール例外: %s", e)
    return 0


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="マスタ有効SKU→タイムセール施策同期（A/B分類）")
    ap.add_argument("--config", type=Path, default=None)
    ap.add_argument("--source", type=Path, default=None)
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--schema-only", action="store_true")
    args = ap.parse_args(argv)
    local = HERE / "config.local.json"
    if not local.is_file():
        local.write_text(
            (HERE / "config.example.json").read_text(encoding="utf-8"), encoding="utf-8"
        )
    cfg = load_config(args.config or local)
    if args.schema_only:
        svc = sheets_service(write=True)
        ensure_unified_schema(svc, str(cfg.get("ads_spreadsheet_id") or "").strip())
        LOG.info("schema-only OK")
        return 0
    return sync(cfg, source=args.source, write=bool(args.write))


if __name__ == "__main__":
    raise SystemExit(main())
