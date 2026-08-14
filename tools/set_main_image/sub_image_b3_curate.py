# -*- coding: utf-8 -*-
"""
サブ画像用: B-③シート右列にレ点・自動判定を書き、採用のみログを同期する。

方針:
  - 作業面: 「競合画像取得（必要時B-③実行）」の右列（商品名〜サブ採用CK）
  - 採用＝レ点（チェックボックス）
  - 永続化: 「サブ画像採用ログ」に採用行のみ
  - 別シート「サブ画像競合候補（人間確認）」は削除
  - B-③再実行時のレ点復元は GAS 側（ログ＋破棄前レ点マージ）

マスタ／R2／出品データは書かない。
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from b3_comp_catalog import (
    B3_SHEET_TITLE,
    B3ImageRow,
    cache_b3_images,
    fetch_b3_rows,
    parse_b3_image_rows,
    prefer_order,
)
from sheets_master import fetch_master_rows
from sheets_rw import (
    apply_checkbox_column,
    build_sheets_rw,
    delete_sheet_by_title,
    ensure_sheet,
    freeze_header_and_autosize,
    list_sheet_titles,
    read_sheet_values,
    spreadsheet_id,
    write_sheet_values,
)
from work_paths import default_work_root, meta_dir_for

LOG = logging.getLogger("set_main_image.sub_image_b3_curate")

# 旧PoCシート（削除対象）
LEGACY_CURATE_SHEET_TITLE = "サブ画像競合候補（人間確認）"
ADOPT_LOG_SHEET_TITLE = "サブ画像採用ログ"

B3_BASE_HEADERS = [
    "マスタ行",
    "JAN",
    "A.セット商品数",
    "区分",
    "ASINまたは商品URL",
    "画像番号",
    "画像URL",
    "プレビュー",
]
B3_EXTRA_HEADERS = [
    "商品名",
    "自動判定",
    "自動理由",
    "意図ラベル",
    "OCR要約",
    "参照themeId",
    "参照phaseOrder",
    "参照cluster",
    "サブ採用CK",
    "メモ",
]
B3_FULL_HEADERS = B3_BASE_HEADERS + B3_EXTRA_HEADERS
SUB_ADOPT_CK = "サブ採用CK"
# 0-based index of サブ採用CK in full row
COL_SUB_ADOPT_0 = B3_FULL_HEADERS.index(SUB_ADOPT_CK)

ADOPT_LOG_HEADERS = [
    "JAN",
    "商品名",
    "区分",
    "ASINまたは商品URL",
    "画像番号",
    "画像URL",
    "メモ",
    "更新日時",
]

KNOWN_NAMES = {
    "4538872180149": "吉野家 唐辛子 30g",
    "4538872281013": "吉野家 缶飯 牛丼",
    "4538872285127": "吉野家 缶飯 焼鳥丼",
}


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _truthy(v: Any) -> bool:
    if v is True:
        return True
    s = str(v or "").strip().upper()
    return s in ("TRUE", "1", "YES", "Y", "ON")


def _normalize_b3_row(old_headers: Sequence[str], row: Sequence[Any]) -> List[Any]:
    """旧ヘッダー行を現行 B3_FULL_HEADERS にマッピング（列追加に耐える）。"""
    idx = {str(h).strip(): i for i, h in enumerate(old_headers) if str(h).strip()}
    out: List[Any] = []
    for h in B3_FULL_HEADERS:
        if h in idx and idx[h] < len(row):
            out.append(row[idx[h]])
        else:
            out.append("")
    return out


def resolve_product_names(jans: List[str]) -> Dict[str, str]:
    out = {j: KNOWN_NAMES.get(j, "") for j in jans}
    try:
        rows, _info = fetch_master_rows()
    except Exception as e:
        LOG.warning("master name lookup skipped: %s", e)
        return out
    from master_sets import _col, _find_header_row

    header_i, idx = _find_header_row(rows)
    i_jan = _col(idx, ("JANコード", "JAN", "jan"))
    i_name = _col(idx, ("商品名", "▼マスタ(商品名)", "品名"))
    if i_jan is None:
        return out
    want = set(jans)
    for row in rows[header_i + 1 :]:
        if i_jan >= len(row):
            continue
        jan = str(row[i_jan] or "").strip()
        if jan not in want:
            continue
        name = ""
        if i_name is not None and i_name < len(row):
            name = str(row[i_name] or "").strip()
        if name:
            out[jan] = name
    return out


def load_adopt_keys_from_log(svc, sid: str) -> Dict[Tuple[str, str], bool]:
    titles = list_sheet_titles(svc, sid)
    if ADOPT_LOG_SHEET_TITLE not in titles:
        return {}
    rows = read_sheet_values(svc, sid, ADOPT_LOG_SHEET_TITLE)
    if len(rows) < 2:
        return {}
    header = [str(c).strip() for c in rows[0]]
    try:
        i_jan = header.index("JAN")
        i_url = header.index("画像URL")
    except ValueError:
        return {}
    out: Dict[Tuple[str, str], bool] = {}
    for r in rows[1:]:
        if max(i_jan, i_url) >= len(r):
            continue
        jan = str(r[i_jan] or "").strip()
        url = str(r[i_url] or "").strip()
        if jan and url.startswith("http"):
            out[(jan, url)] = True
    LOG.info("adopt log keys=%d", len(out))
    return out


def load_checks_from_b3(svc, sid: str) -> Dict[Tuple[str, str], bool]:
    rows = read_sheet_values(svc, sid, B3_SHEET_TITLE)
    if len(rows) < 2:
        return {}
    header = [str(c).strip() for c in rows[0]]
    try:
        i_jan = header.index("JAN")
        i_url = header.index("画像URL")
        i_ck = header.index(SUB_ADOPT_CK)
    except ValueError:
        return {}
    out: Dict[Tuple[str, str], bool] = {}
    for r in rows[1:]:
        if max(i_jan, i_url, i_ck) >= len(r):
            continue
        if not _truthy(r[i_ck]):
            continue
        jan = str(r[i_jan] or "").strip()
        url = str(r[i_url] or "").strip()
        if jan and url.startswith("http"):
            out[(jan, url)] = True
    LOG.info("b3 existing checks=%d", len(out))
    return out


def load_checks_from_legacy_curate(svc, sid: str) -> Dict[Tuple[str, str], bool]:
    titles = list_sheet_titles(svc, sid)
    if LEGACY_CURATE_SHEET_TITLE not in titles:
        return {}
    rows = read_sheet_values(svc, sid, LEGACY_CURATE_SHEET_TITLE)
    if len(rows) < 2:
        return {}
    header = [str(c).strip() for c in rows[0]]
    try:
        i_jan = header.index("JAN")
        i_url = header.index("画像URL")
        i_ck = header.index("採用CK") if "採用CK" in header else header.index(SUB_ADOPT_CK)
    except ValueError:
        return {}
    out: Dict[Tuple[str, str], bool] = {}
    for r in rows[1:]:
        if max(i_jan, i_url, i_ck) >= len(r):
            continue
        if not _truthy(r[i_ck]):
            continue
        jan = str(r[i_jan] or "").strip()
        url = str(r[i_url] or "").strip()
        if jan and url.startswith("http"):
            out[(jan, url)] = True
    LOG.info("legacy curate checks migrated=%d", len(out))
    return out


def structural_verdict(row: B3ImageRow) -> Tuple[str, str, bool]:
    if row.image_index <= 1:
        return (
            "drop_main",
            "画像番号=1（メイン相当の可能性）。サブ候補から除外デフォルト",
            False,
        )
    return ("review", "サブ候補（意図分類待ち／人手確認）", False)


def write_adopt_log_only(
    svc,
    sid: str,
    adopted_rows: Sequence[Sequence[Any]],
) -> None:
    """adopted_rows はログ用データ行（ヘッダーなし）。"""
    now = datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M:%S")
    body: List[List[Any]] = [list(ADOPT_LOG_HEADERS)]
    for r in adopted_rows:
        body.append(list(r) + ([now] if len(r) < len(ADOPT_LOG_HEADERS) else []))
    # normalize: expect 7 fields before 更新日時
    fixed: List[List[Any]] = [list(ADOPT_LOG_HEADERS)]
    for r in adopted_rows:
        rr = list(r)
        while len(rr) < 7:
            rr.append("")
        fixed.append(rr[:7] + [now])
    sheet_id = ensure_sheet(svc, sid, ADOPT_LOG_SHEET_TITLE)
    write_sheet_values(svc, sid, ADOPT_LOG_SHEET_TITLE, fixed)
    freeze_header_and_autosize(svc, sid, sheet_id, len(ADOPT_LOG_HEADERS))
    LOG.info("adopt log rows=%d", len(fixed) - 1)


def sync_adopt_log_from_b3(svc=None, sid: str = "") -> int:
    """B-③のサブ採用CKレ点のみ → 採用ログへ書き換え。"""
    svc = svc or build_sheets_rw()
    sid = sid or spreadsheet_id()
    rows = read_sheet_values(svc, sid, B3_SHEET_TITLE)
    if len(rows) < 2:
        write_adopt_log_only(svc, sid, [])
        return 0
    header = [str(c).strip() for c in rows[0]]
    need = ["JAN", "商品名", "区分", "ASINまたは商品URL", "画像番号", "画像URL", SUB_ADOPT_CK]
    for n in need:
        if n not in header:
            raise RuntimeError(f"B-③に列がありません: {n}（先に curate を実行）")
    i_memo = header.index("メモ") if "メモ" in header else -1
    adopted: List[List[Any]] = []
    for r in rows[1:]:
        if not _truthy(r[header.index(SUB_ADOPT_CK)]):
            continue
        url = str(r[header.index("画像URL")] or "").strip()
        if not url.startswith("http"):
            continue
        adopted.append(
            [
                str(r[header.index("JAN")] or "").strip(),
                str(r[header.index("商品名")] or "").strip(),
                str(r[header.index("区分")] or "").strip(),
                str(r[header.index("ASINまたは商品URL")] or "").strip(),
                r[header.index("画像番号")],
                url,
                str(r[i_memo] or "").strip() if i_memo >= 0 else "",
            ]
        )
    write_adopt_log_only(svc, sid, adopted)
    return len(adopted)


def curate_onto_b3(
    *,
    jans: Optional[List[str]],
    work_root: Path,
    classify: bool,
    max_classify_per_jan: int,
    preserve_checks: bool,
    delete_legacy: bool,
    sync_log: bool,
) -> Path:
    svc = build_sheets_rw()
    sid = spreadsheet_id()

    b3_raw = fetch_b3_rows()
    # 全JAN or 指定
    all_parsed = parse_b3_image_rows(b3_raw, jan="")
    if jans:
        want = set(j.strip() for j in jans if j.strip())
        all_parsed = [r for r in all_parsed if r.jan in want]
    jan_list = sorted({r.jan for r in all_parsed})
    if not jan_list:
        raise SystemExit("B-③に対象JANの画像URLがありません")

    names = resolve_product_names(jan_list)
    LOG.info("JANs=%s", jan_list)

    adopt_keys: Dict[Tuple[str, str], bool] = {}
    if preserve_checks:
        adopt_keys.update(load_adopt_keys_from_log(svc, sid))
        adopt_keys.update(load_checks_from_b3(svc, sid))
        adopt_keys.update(load_checks_from_legacy_curate(svc, sid))

    client = None
    if classify:
        from gemini_image import load_api_key, make_client

        client = make_client(load_api_key())

    cache_root = work_root / "00.テスト出力" / "sub_image_b3_curate" / "cache"
    out_meta = work_root / "00.テスト出力" / "sub_image_b3_curate"
    out_meta.mkdir(parents=True, exist_ok=True)

    # JANごとに分類（混在防止）
    by_jan: Dict[str, List[B3ImageRow]] = {}
    for r in all_parsed:
        by_jan.setdefault(r.jan, []).append(r)

    classified_urls: Dict[str, Dict[str, Any]] = {}
    theme_urls: Dict[str, Dict[str, Any]] = {}
    if classify and client:
        from sub_image_intent import classify_competitor_image
        from sub_image_lp_themes import classify_image_themes

        for jan in jan_list:
            rows = prefer_order(by_jan.get(jan) or [])
            targets = [r for r in rows if r.image_index >= 2][: max(0, max_classify_per_jan)]
            cached = cache_b3_images(targets, cache_root / jan, limit=len(targets))
            for item in cached:
                try:
                    classified_urls[item["url"]] = classify_competitor_image(
                        Path(item["path"]), client=client
                    )
                except Exception as e:
                    classified_urls[item["url"]] = {
                        "intentLabel": "other_unknown",
                        "decision": "review",
                        "ocrTextPreview": "",
                        "reasonJa": str(e),
                    }
                try:
                    theme_urls[item["url"]] = classify_image_themes(
                        Path(item["path"]), client=client
                    )
                except Exception as e:
                    theme_urls[item["url"]] = {
                        "primaryThemeId": 0,
                        "phaseOrder": 0,
                        "reasonJa": str(e),
                        "reject": True,
                    }

    # B-③の元順序を保ちつつ右列を付与（JANフィルタ時は該当行のみ出力）
    # シート全体を書き直すと他JANが消えるため、jans指定時は「既存B-③を読み、対象JAN行だけ置換」する
    existing = read_sheet_values(svc, sid, B3_SHEET_TITLE)
    existing_by_key: Dict[Tuple[str, str], List[str]] = {}
    if existing and len(existing) > 1:
        eh = [str(c).strip() for c in existing[0]]
        try:
            ej = eh.index("JAN")
            eu = eh.index("画像URL")
            for r in existing[1:]:
                if max(ej, eu) >= len(r):
                    continue
                existing_by_key[(str(r[ej]).strip(), str(r[eu]).strip())] = list(r)
        except ValueError:
            pass

    sheet_rows: List[List[Any]] = [list(B3_FULL_HEADERS)]
    stats: Dict[str, Any] = {}

    # 出力対象: 指定がなければB-③全件、あれば指定JANのみだが他JAN行は既存から残す
    if jans:
        keep_other = []
        if existing and len(existing) > 1:
            eh = [str(c).strip() for c in existing[0]]
            try:
                ej = eh.index("JAN")
                for r in existing[1:]:
                    jan = str(r[ej] if ej < len(r) else "").strip()
                    if jan and jan not in want:
                        keep_other.append(_normalize_b3_row(eh, r))
            except ValueError:
                pass
        for rr in keep_other:
            sheet_rows.append(rr)

    for jan in jan_list:
        product_name = names.get(jan) or KNOWN_NAMES.get(jan) or jan
        rows = prefer_order(by_jan.get(jan) or [])
        jan_stats = {"total": 0, "keep": 0, "drop_main": 0, "drop_unrelated": 0, "review": 0}
        for r in rows:
            verdict, reason, default_accept = structural_verdict(r)
            intent = ""
            ocr = ""
            c = classified_urls.get(r.url)
            if c:
                intent = str(c.get("intentLabel") or "")
                ocr = str(c.get("ocrTextPreview") or "")[:200]
                decision = str(c.get("decision") or "review")
                if decision == "reject":
                    verdict = "drop_unrelated"
                    reason = f"意図除外: {intent} / {c.get('reasonJa') or ''}"
                    default_accept = False
                elif decision == "use" and verdict != "drop_main":
                    verdict = "keep"
                    reason = f"商品直結: {intent}"
                    default_accept = True
                elif decision == "review" and verdict != "drop_main":
                    verdict = "review"
                    reason = f"要確認: {intent}"
                    default_accept = False

            key = (jan, r.url)
            if preserve_checks and key in adopt_keys:
                accepted = True
            else:
                accepted = bool(default_accept)

            ref_tid = ""
            ref_phase = ""
            ref_cluster = ""
            tmeta = theme_urls.get(r.url) if classify else None
            if tmeta and not tmeta.get("reject"):
                tid = int(tmeta.get("primaryThemeId") or 0)
                if tid >= 1:
                    ref_tid = tid
                    ref_phase = int(tmeta.get("phaseOrder") or 0) or ""
                    cluster = ""
                    try:
                        from sub_image_lp_themes import THEME_BY_ID

                        cluster = str((THEME_BY_ID.get(tid) or {}).get("contentCluster") or "")
                    except Exception:
                        cluster = ""
                    ref_cluster = cluster

            preview = f'=IMAGE("{r.url}",2)' if r.url.startswith("http") else ""
            sheet_rows.append(
                [
                    r.master_row,
                    jan,
                    r.set_qty,
                    r.kind,
                    r.listing_key,
                    r.image_index if r.image_index else "",
                    r.url,
                    preview,
                    product_name,
                    verdict,
                    reason,
                    intent,
                    ocr,
                    ref_tid,
                    ref_phase,
                    ref_cluster,
                    True if accepted else False,
                    "",
                ]
            )
            jan_stats["total"] += 1
            jan_stats[verdict] = jan_stats.get(verdict, 0) + 1
        stats[jan] = {"name": product_name, **jan_stats}
        LOG.info("JAN=%s stats=%s", jan, jan_stats)

    sheet_id = ensure_sheet(svc, sid, B3_SHEET_TITLE)
    write_sheet_values(svc, sid, B3_SHEET_TITLE, sheet_rows)
    if len(sheet_rows) > 1:
        apply_checkbox_column(
            svc,
            sid,
            sheet_id,
            start_row_0=1,
            end_row_0=len(sheet_rows),
            col_0=COL_SUB_ADOPT_0,
        )
    freeze_header_and_autosize(svc, sid, sheet_id, len(B3_FULL_HEADERS))

    if sync_log:
        n = sync_adopt_log_from_b3(svc, sid)
        LOG.info("synced adopt log n=%d", n)

    if delete_legacy:
        delete_sheet_by_title(svc, sid, LEGACY_CURATE_SHEET_TITLE)

    meta_path = meta_dir_for(out_meta) / "curate_meta.json"
    meta = {
        "spreadsheetId": sid,
        "b3Sheet": B3_SHEET_TITLE,
        "adoptLogSheet": ADOPT_LOG_SHEET_TITLE,
        "rowCount": len(sheet_rows) - 1,
        "refThemeFilled": sum(
            1
            for r in sheet_rows[1:]
            if len(r) > B3_FULL_HEADERS.index("参照themeId")
            and str(r[B3_FULL_HEADERS.index("参照themeId")] or "").strip()
        ),
        "jans": stats,
        "legacyDeleted": delete_legacy,
    }
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"SHEET={B3_SHEET_TITLE}")
    print(f"ADOPT_LOG={ADOPT_LOG_SHEET_TITLE}")
    print(f"URL=https://docs.google.com/spreadsheets/d/{sid}/edit")
    print(f"META={meta_path}")
    return meta_path


def read_accepted_from_b3(
    *,
    jans: Optional[List[str]] = None,
) -> Dict[str, List[Dict[str, str]]]:
    """B-③のサブ採用CKレ点のみ、JAN別に返す。"""
    svc = build_sheets_rw()
    sid = spreadsheet_id()
    rows = read_sheet_values(svc, sid, B3_SHEET_TITLE)
    if len(rows) < 2:
        return {}
    header = [str(c).strip() for c in rows[0]]
    for n in ("JAN", "画像URL", SUB_ADOPT_CK, "区分", "ASINまたは商品URL", "画像番号"):
        if n not in header:
            raise RuntimeError(f"B-③に列がありません: {n}")
    i_name = header.index("商品名") if "商品名" in header else -1
    want = set(j.strip() for j in (jans or []) if j.strip()) or None
    out: Dict[str, List[Dict[str, str]]] = {}
    for r in rows[1:]:
        if not _truthy(r[header.index(SUB_ADOPT_CK)]):
            continue
        jan = str(r[header.index("JAN")] or "").strip()
        url = str(r[header.index("画像URL")] or "").strip()
        if not jan or not url.startswith("http"):
            continue
        if want is not None and jan not in want:
            continue
        out.setdefault(jan, []).append(
            {
                "jan": jan,
                "productName": str(r[i_name] or "").strip() if i_name >= 0 else "",
                "url": url,
                "kind": str(r[header.index("区分")] or "").strip(),
                "listingKey": str(r[header.index("ASINまたは商品URL")] or "").strip(),
                "imageIndex": str(r[header.index("画像番号")] or "").strip(),
            }
        )
    return out


# 後方互換エイリアス
def read_accepted_from_curate_sheet(
    *,
    jans: Optional[List[str]] = None,
    sheet_title: str = "",
) -> Dict[str, List[Dict[str, str]]]:
    _ = sheet_title
    return read_accepted_from_b3(jans=jans)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="B-③右列にサブ採用レ点・自動判定を書く／採用ログ同期"
    )
    ap.add_argument("--jan", action="append", default=None, help="対象JAN（省略時はB-③全JAN）")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--classify", action="store_true", help="Vision意図分類")
    ap.add_argument("--max-classify-per-jan", type=int, default=8)
    ap.add_argument(
        "--no-preserve-checks",
        action="store_true",
        help="既存レ点・採用ログを無視して自動デフォルトのみ",
    )
    ap.add_argument(
        "--keep-legacy-sheet",
        action="store_true",
        help="旧『サブ画像競合候補（人間確認）』を削除しない",
    )
    ap.add_argument(
        "--sync-adopt-log-only",
        action="store_true",
        help="B-③のレ点→採用ログだけ同期（自動判定の再計算なし）",
    )
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)

    if args.sync_adopt_log_only:
        n = sync_adopt_log_from_b3()
        print(f"ADOPT_LOG_SYNCED={n}")
        return 0

    curate_onto_b3(
        jans=args.jan,
        work_root=args.work_root or default_work_root(),
        classify=bool(args.classify),
        max_classify_per_jan=max(0, int(args.max_classify_per_jan)),
        preserve_checks=not bool(args.no_preserve_checks),
        delete_legacy=not bool(args.keep_legacy_sheet),
        sync_log=True,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
