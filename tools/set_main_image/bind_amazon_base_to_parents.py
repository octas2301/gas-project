# -*- coding: utf-8 -*-
"""
01.amazon白抜きベース の未紐付けPNGを、出品CKレ点親の「楽天メイン画像1」と
Vision照合して親SKUへ自動割当し、`{親SKU}_単体.png` へリネームする。

人手確認なし（閾値未満・取り合い負けは SKIP）。

  python bind_amazon_base_to_parents.py --from-sheets
  python bind_amazon_base_to_parents.py --from-sheets --dry-run
  python bind_amazon_base_to_parents.py --from-sheets --require-bound
"""
from __future__ import annotations

import argparse
import hashlib
import json
import logging
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from competitor_fact import (
    DEFAULT_COLOR_MIN,
    DEFAULT_OVERALL_MIN,
    DEFAULT_SHAPE_MIN,
    FactMatchScore,
    download_image,
    match_product_images,
)
from master_sets import (
    CHILD_HINTS,
    PARENT_HINTS,
    _col,
    _find_header_row,
    _norm,
    load_checked_set_children_from_rows,
)
from sheets_master import fetch_master_rows
from work_paths import SUB_AMAZON, TAG_UNIT, default_work_root

LOG = logging.getLogger("set_main_image.bind_base")

RAKUTEN_MAIN_HINTS = (
    "楽天メイン画像1",
    "楽天メイン画像",
    "メイン画像1",
)
IMAGE_FORMULA_RE = re.compile(
    r'(?i)=\s*IMAGE\s*\(\s*"([^"]+)"|=\s*IMAGE\s*\(\s*\'([^\']+)\''
)
DRIVE_ID_RE = re.compile(
    r"(?:/d/|id=|open\?id=)([a-zA-Z0-9_-]{20,})"
)


@dataclass
class ParentRef:
    parent_sku: str
    rakuten_main_url: str = ""
    product_name: str = ""


@dataclass
class BindResult:
    status: str  # renamed | already_bound | skip | error
    base_file: str = ""
    parent_sku: str = ""
    new_name: str = ""
    old_name: str = ""
    reason: str = ""
    overall: Optional[int] = None
    shape: Optional[int] = None
    color: Optional[int] = None
    rakuten_main_url: str = ""
    scores: List[Dict[str, Any]] = field(default_factory=list)


def _safe_stem(s: str) -> str:
    out: List[str] = []
    for ch in (s or "").strip():
        if ch.isalnum() or ch in ("-", "_", "."):
            out.append(ch)
        elif ch in (" ", "\t"):
            out.append("_")
        else:
            out.append("_")
    return "".join(out)[:120] or "sku"


def extract_image_url(cell: str) -> str:
    """セル値から http(s) URL または R-Cabinet 相対パスを取り出す。"""
    s = (cell or "").strip()
    if not s:
        return ""
    m = IMAGE_FORMULA_RE.search(s)
    if m:
        return (m.group(1) or m.group(2) or "").strip()
    if s.startswith("http://") or s.startswith("https://"):
        return re.split(r"[\s,\"']", s, maxsplit=1)[0]
    m2 = re.search(r"https?://[^\s,\"']+", s)
    if m2:
        return m2.group(0)
    # 楽天キャビネット相対: /12644827/imgrc….jpg
    if s.startswith("/") and ("/" in s[1:]) and re.search(r"\.(jpe?g|png|gif|webp)$", s, re.I):
        return s
    if re.match(r"^\d{5,}/\S+\.(jpe?g|png|gif|webp)$", s, re.I):
        return "/" + s
    return ""


def load_rakuten_shop_id(
    *,
    config_path: Optional[Path] = None,
    spreadsheet_id: str = "",
) -> str:
    """▼設定(マッピング) の『ショップＩＤ楽天』を読む（例: octas）。"""
    from googleapiclient.discovery import build

    from sheets_master import _get_credentials, load_sheets_settings

    settings = load_sheets_settings(config_path, spreadsheet_id=spreadsheet_id)
    creds = _get_credentials(settings["credentials_path"], settings["token_path"])
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    data = (
        sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=settings["spreadsheet_id"],
            range="'▼設定(マッピング)'!A1:B40",
            majorDimension="ROWS",
            valueRenderOption="FORMATTED_VALUE",
        )
        .execute()
    )
    for row in data.get("values") or []:
        if not row:
            continue
        key_raw = str(row[0]).strip()
        key = (
            key_raw.replace(" ", "")
            .replace("　", "")
            .replace("Ｉ", "I")
            .replace("Ｄ", "D")
            .replace("id", "ID")
        )
        val = str(row[1]).strip() if len(row) > 1 else ""
        if not val:
            continue
        if "ショップID楽天" in key or (key.startswith("ショップ") and "楽天" in key and "ID" in key):
            return val
    return ""


def resolve_rakuten_image_url(ref: str, *, shop_id: str) -> str:
    """
    相対パスを https://image.rakuten.co.jp/{shop}/cabinet{path} へ。
    既に http(s) ならそのまま。
    """
    s = (ref or "").strip()
    if not s:
        return ""
    if s.startswith("http://") or s.startswith("https://"):
        return s
    if not s.startswith("/"):
        s = "/" + s
    shop = (shop_id or "").strip()
    if not shop:
        LOG.warning("rakuten shop_id empty — cannot expand cabinet path %s", s)
        return ""
    return f"https://image.rakuten.co.jp/{shop}/cabinet{s}"

def _drive_file_id(url: str) -> str:
    m = DRIVE_ID_RE.search(url or "")
    return m.group(1) if m else ""


def fetch_url_to_file(
    url: str,
    dest: Path,
    *,
    credentials_path: Optional[Path] = None,
    token_path: Optional[Path] = None,
) -> bool:
    """http(s) または Google Drive ファイルを取得。"""
    dest = Path(dest)
    if dest.is_file() and dest.stat().st_size > 500:
        return True
    drive_id = _drive_file_id(url)
    if drive_id and credentials_path and token_path:
        try:
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaIoBaseDownload
            import io

            from sheets_master import _get_credentials

            creds = _get_credentials(Path(credentials_path), Path(token_path))
            drive = build("drive", "v3", credentials=creds, cache_discovery=False)
            req = drive.files().get_media(fileId=drive_id)
            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            data = buf.getvalue()
            if len(data) < 500:
                LOG.warning("Drive download too small id=%s", drive_id)
                return False
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(data)
            return True
        except Exception as e:
            LOG.warning("Drive download failed id=%s: %s", drive_id, e)
            # 公開リンク風にフォールバック
            url = f"https://drive.google.com/uc?export=download&id={drive_id}"
    return download_image(url, dest)


def list_base_pngs(folder: Path) -> List[Path]:
    out: List[Path] = []
    if not folder.is_dir():
        return out
    for p in sorted(folder.iterdir(), key=lambda x: x.name.lower()):
        if not p.is_file():
            continue
        if p.suffix.lower() != ".png":
            continue
        out.append(p)
    return out


def already_bound_to_parent(path: Path, parent_sku: str) -> bool:
    name = path.name.lower()
    stem = path.stem.lower()
    want = (parent_sku or "").strip().lower()
    if want and want in name:
        return True
    for m in re.findall(r"\d{8,}", want):
        if m and m in name:
            return True
    # JAN断片が stem にある場合も「紐付け済」扱い
    return False


def any_parent_token_in_name(path: Path, parents: List[ParentRef]) -> Optional[str]:
    for pr in parents:
        if already_bound_to_parent(path, pr.parent_sku):
            return pr.parent_sku
    return None


def _merge_formula_image_urls(
    rows_unformatted: List[List[str]], rows_formula: List[List[str]]
) -> List[List[str]]:
    """UNFORMATTED行に、FORMULA行から抽出した画像URLを埋め戻す。"""
    if not rows_unformatted or not rows_formula:
        return rows_unformatted
    header_i, idx = _find_header_row(rows_unformatted)
    i_main = _col(idx, RAKUTEN_MAIN_HINTS)
    if i_main is None:
        return rows_unformatted
    out = [list(r) for r in rows_unformatted]
    for ri in range(header_i + 1, len(out)):
        if ri >= len(rows_formula):
            break
        fr = rows_formula[ri]
        if i_main >= len(out[ri]):
            out[ri].extend([""] * (i_main - len(out[ri]) + 1))
        cur = extract_image_url(out[ri][i_main] if i_main < len(out[ri]) else "")
        if cur:
            continue
        if i_main < len(fr):
            url = extract_image_url(fr[i_main])
            if url:
                out[ri][i_main] = url
    return out


def enrich_master_rows_rakuten_main_urls(
    rows: List[List[str]],
    *,
    config_path: Optional[Path] = None,
    spreadsheet_id: str = "",
    master_sheet: str = "",
    shop_id: str = "",
) -> List[List[str]]:
    """
    楽天メイン参照の補完。
    - 相対パスは shop_id 解決は load 時に行うため、ここでは FORMULA 補完のみ。
    """
    sid = shop_id or load_rakuten_shop_id(
        config_path=config_path, spreadsheet_id=spreadsheet_id
    )
    parents_probe = load_checked_parents_with_rakuten_main(rows, shop_id=sid)
    if not parents_probe or any(p.rakuten_main_url for p in parents_probe):
        return rows
    LOG.info("楽天メイン参照が空のため FORMULA 再取得を試行")
    rows_f, _info_f = fetch_master_rows(
        config_path=config_path,
        spreadsheet_id=spreadsheet_id,
        master_sheet=master_sheet,
        value_render_option="FORMULA",
    )
    return _merge_formula_image_urls(rows, rows_f)

def load_checked_parents_with_rakuten_main(
    rows: List[List[str]],
    *,
    shop_id: str = "",
) -> List[ParentRef]:
    children, _food = load_checked_set_children_from_rows(rows)
    parent_skus = sorted({c.parent_sku for c in children})
    header_i, idx = _find_header_row(rows)
    i_p = _col(idx, PARENT_HINTS)
    i_c = _col(idx, CHILD_HINTS)
    i_main = _col(idx, RAKUTEN_MAIN_HINTS)
    i_name = _col(idx, ("商品名", "▼マスタ(商品名)", "品名"))
    i_amz = _col(idx, ("出品者SKUのメイン画像URL", "Amazon MAIN URL"))
    if i_p is None:
        raise ValueError("親SKU列がありません")
    if i_main is None:
        raise ValueError("楽天メイン画像1列がありません")

    out: List[ParentRef] = []
    for psku in parent_skus:
        want = _norm(psku)
        parent_row = None
        any_row = None
        for row in rows[header_i + 1 :]:
            if i_p >= len(row) or _norm(row[i_p]) != want:
                continue
            any_row = row
            child = _norm(row[i_c]) if i_c is not None and i_c < len(row) else ""
            if not child:
                parent_row = row
                break
        row = parent_row or any_row
        url = ""
        pname = ""
        if row is not None:
            if i_main < len(row):
                url = resolve_rakuten_image_url(
                    extract_image_url(row[i_main]), shop_id=shop_id
                )
            if i_name is not None and i_name < len(row):
                pname = _norm(row[i_name])
            if not url and i_amz is not None and i_amz < len(row):
                url = extract_image_url(row[i_amz])
        # 親行にメインが空なら同親の子行から探す
        if not url:
            for row in rows[header_i + 1 :]:
                if i_p >= len(row) or _norm(row[i_p]) != want:
                    continue
                if i_main < len(row):
                    url = resolve_rakuten_image_url(
                        extract_image_url(row[i_main]), shop_id=shop_id
                    )
                    if url:
                        if i_name is not None and i_name < len(row) and not pname:
                            pname = _norm(row[i_name])
                        break
                if not url and i_amz is not None and i_amz < len(row):
                    url = extract_image_url(row[i_amz])
                    if url:
                        break
        out.append(
            ParentRef(parent_sku=psku, rakuten_main_url=url, product_name=pname)
        )
    return out

def passes_gate(
    score: FactMatchScore,
    *,
    overall_min: int,
    shape_min: int,
    color_min: int,
) -> bool:
    return (
        int(score.overall) >= int(overall_min)
        and int(score.shape) >= int(shape_min)
        and int(score.color) >= int(color_min)
    )


def assign_greedy(
    scores: List[Tuple[Path, str, FactMatchScore]],
    *,
    overall_min: int,
    shape_min: int,
    color_min: int,
) -> Dict[Path, Tuple[str, FactMatchScore]]:
    """
    overall 降順で貪欲割当。1 PNG ↔ 1 親。
    """
    ranked = sorted(
        scores,
        key=lambda t: (t[2].overall, t[2].shape, t[2].color),
        reverse=True,
    )
    used_png: Set[Path] = set()
    used_parent: Set[str] = set()
    assigned: Dict[Path, Tuple[str, FactMatchScore]] = {}
    for png, psku, sc in ranked:
        if png in used_png or psku in used_parent:
            continue
        if not passes_gate(
            sc, overall_min=overall_min, shape_min=shape_min, color_min=color_min
        ):
            continue
        assigned[png] = (psku, sc)
        used_png.add(png)
        used_parent.add(psku)
    return assigned


def target_name_for_parent(parent_sku: str, folder: Path) -> Path:
    base = f"{_safe_stem(parent_sku)}_{TAG_UNIT}.png"
    dest = folder / base
    if not dest.exists():
        return dest
    for i in range(2, 30):
        alt = folder / f"{_safe_stem(parent_sku)}_{TAG_UNIT}_{i}.png"
        if not alt.exists():
            return alt
    return folder / f"{_safe_stem(parent_sku)}_{TAG_UNIT}_{datetime.now().strftime('%H%M%S')}.png"


def run_bind(
    *,
    work_root: Path,
    rows: List[List[str]],
    dry_run: bool,
    overall_min: int,
    shape_min: int,
    color_min: int,
    credentials_path: Optional[Path],
    token_path: Optional[Path],
    shop_id: str = "",
) -> Dict[str, Any]:
    folder = work_root / SUB_AMAZON
    cache = folder / "_bind_cache"
    cache.mkdir(parents=True, exist_ok=True)

    sid = (shop_id or "").strip()
    parents = load_checked_parents_with_rakuten_main(rows, shop_id=sid)
    LOG.info("rakuten shop_id=%s parents_with_url=%s", sid, sum(1 for p in parents if p.rakuten_main_url))
    pngs = list_base_pngs(folder)
    results: List[BindResult] = []

    # 既にSKU入りのものはスキップ
    unbound: List[Path] = []
    for p in pngs:
        hit = any_parent_token_in_name(p, parents)
        if hit:
            results.append(
                BindResult(
                    status="already_bound",
                    base_file=p.name,
                    parent_sku=hit,
                    old_name=p.name,
                    reason="filename_already_has_parent_token",
                )
            )
        else:
            unbound.append(p)

    LOG.info(
        "bind parents=%s unbound_png=%s already=%s",
        len(parents),
        len(unbound),
        len(pngs) - len(unbound),
    )

    # 楽天メイン取得
    parent_img: Dict[str, Path] = {}
    for pr in parents:
        if not pr.rakuten_main_url:
            LOG.warning("no rakuten main url parent=%s", pr.parent_sku)
            continue
        h = hashlib.sha1(pr.rakuten_main_url.encode("utf-8")).hexdigest()[:16]
        dest = cache / f"{_safe_stem(pr.parent_sku)}_{h}.img"
        ok = fetch_url_to_file(
            pr.rakuten_main_url,
            dest,
            credentials_path=credentials_path,
            token_path=token_path,
        )
        if ok:
            # 拡張子推定
            try:
                from PIL import Image

                im = Image.open(dest)
                ext = ".jpg" if (im.format or "").upper() in ("JPEG", "JPG") else ".png"
                named = dest.with_suffix(ext)
                if named != dest:
                    if named.exists():
                        try:
                            dest.unlink()
                        except OSError:
                            pass
                    else:
                        dest.rename(named)
                    dest = named
            except Exception:
                pass
            parent_img[pr.parent_sku] = dest
        else:
            LOG.warning(
                "failed to fetch rakuten main parent=%s url=%s",
                pr.parent_sku,
                pr.rakuten_main_url[:80],
            )

    pair_scores: List[Tuple[Path, str, FactMatchScore]] = []
    score_log: Dict[str, List[Dict[str, Any]]] = {}

    for png in unbound:
        score_log[png.name] = []
        for pr in parents:
            ref = parent_img.get(pr.parent_sku)
            if ref is None:
                score_log[png.name].append(
                    {
                        "parentSku": pr.parent_sku,
                        "status": "no_ref_image",
                    }
                )
                continue
            LOG.info("vision match base=%s vs parent=%s", png.name, pr.parent_sku)
            sc = match_product_images(png, ref)
            if sc is None:
                score_log[png.name].append(
                    {
                        "parentSku": pr.parent_sku,
                        "status": "vision_failed",
                    }
                )
                continue
            pair_scores.append((png, pr.parent_sku, sc))
            score_log[png.name].append(
                {
                    "parentSku": pr.parent_sku,
                    "status": "scored",
                    "overall": sc.overall,
                    "shape": sc.shape,
                    "color": sc.color,
                    "text": sc.text,
                    "package": sc.package,
                    "capacity": sc.capacity,
                    "gateOk": passes_gate(
                        sc,
                        overall_min=overall_min,
                        shape_min=shape_min,
                        color_min=color_min,
                    ),
                }
            )

    assigned = assign_greedy(
        pair_scores,
        overall_min=overall_min,
        shape_min=shape_min,
        color_min=color_min,
    )
    parent_url = {p.parent_sku: p.rakuten_main_url for p in parents}

    for png in unbound:
        if png in assigned:
            psku, sc = assigned[png]
            dest = target_name_for_parent(psku, folder)
            br = BindResult(
                status="renamed" if not dry_run else "would_rename",
                base_file=png.name,
                parent_sku=psku,
                new_name=dest.name,
                old_name=png.name,
                reason="vision_match",
                overall=sc.overall,
                shape=sc.shape,
                color=sc.color,
                rakuten_main_url=parent_url.get(psku, ""),
                scores=score_log.get(png.name) or [],
            )
            if not dry_run:
                try:
                    png.rename(dest)
                    LOG.info("renamed %s -> %s (overall=%s)", png.name, dest.name, sc.overall)
                except OSError as e:
                    br.status = "error"
                    br.reason = f"rename_failed:{e}"
            results.append(br)
        else:
            # ベストスコアを理由に残す
            best = None
            for cand in score_log.get(png.name) or []:
                if cand.get("status") != "scored":
                    continue
                if best is None or int(cand.get("overall") or 0) > int(
                    best.get("overall") or 0
                ):
                    best = cand
            reason = "below_threshold_or_conflict"
            if best is None:
                reason = "no_scores"
            elif not best.get("gateOk"):
                reason = (
                    f"below_threshold best={best.get('parentSku')} "
                    f"overall={best.get('overall')}"
                )
            else:
                reason = (
                    f"lost_assignment best={best.get('parentSku')} "
                    f"overall={best.get('overall')}"
                )
            results.append(
                BindResult(
                    status="skip",
                    base_file=png.name,
                    old_name=png.name,
                    reason=reason,
                    overall=int(best["overall"]) if best and best.get("overall") is not None else None,
                    scores=score_log.get(png.name) or [],
                )
            )

    # レ点親で画像が取れなかったもの
    for pr in parents:
        if pr.parent_sku in parent_img:
            continue
        if any(r.parent_sku == pr.parent_sku and r.status in ("renamed", "would_rename", "already_bound") for r in results):
            continue
        results.append(
            BindResult(
                status="skip",
                parent_sku=pr.parent_sku,
                reason="no_rakuten_main_image",
                rakuten_main_url=pr.rakuten_main_url,
            )
        )

    renamed_n = sum(1 for r in results if r.status in ("renamed", "would_rename"))
    already_n = sum(1 for r in results if r.status == "already_bound")
    skip_n = sum(1 for r in results if r.status == "skip")
    err_n = sum(1 for r in results if r.status == "error")

    summary = {
        "at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "dryRun": dry_run,
        "folder": str(folder),
        "overallMin": overall_min,
        "shapeMin": shape_min,
        "colorMin": color_min,
        "parents": [asdict(p) for p in parents],
        "counts": {
            "renamed": renamed_n,
            "alreadyBound": already_n,
            "skip": skip_n,
            "error": err_n,
            "unboundIn": len(unbound),
        },
        "results": [asdict(r) for r in results],
        "noteJa": "白抜きPNGを楽天メイン画像1とVision照合し親SKUへ自動リネーム。人手確認なし。",
    }
    return summary


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    ap = argparse.ArgumentParser(
        description="Bind 01 amazon base PNGs to checked parents via Rakuten MAIN Vision"
    )
    ap.add_argument("--from-sheets", action="store_true", required=True)
    ap.add_argument("--c1-config", type=Path, default=None)
    ap.add_argument("--spreadsheet-id", default="")
    ap.add_argument("--master-sheet", default="")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument(
        "--require-bound",
        action="store_true",
        help="1枚でも unbound のまま／error なら exit 2",
    )
    ap.add_argument("--overall-min", type=int, default=DEFAULT_OVERALL_MIN)
    ap.add_argument("--shape-min", type=int, default=DEFAULT_SHAPE_MIN)
    ap.add_argument("--color-min", type=int, default=DEFAULT_COLOR_MIN)
    args = ap.parse_args(argv)

    work_root = Path(args.work_root) if args.work_root else default_work_root()
    rows, info = fetch_master_rows(
        config_path=args.c1_config,
        spreadsheet_id=args.spreadsheet_id,
        master_sheet=args.master_sheet,
    )
    from sheets_master import load_sheets_settings

    settings = load_sheets_settings(
        args.c1_config,
        spreadsheet_id=args.spreadsheet_id,
        master_sheet=args.master_sheet,
    )

    shop_id = load_rakuten_shop_id(
        config_path=args.c1_config, spreadsheet_id=args.spreadsheet_id
    )
    LOG.info("loaded rakuten shop_id=%r", shop_id)
    rows = enrich_master_rows_rakuten_main_urls(
        rows,
        config_path=args.c1_config,
        spreadsheet_id=args.spreadsheet_id,
        master_sheet=args.master_sheet,
        shop_id=shop_id,
    )

    summary = run_bind(
        work_root=work_root,
        rows=rows,
        dry_run=bool(args.dry_run),
        overall_min=int(args.overall_min),
        shape_min=int(args.shape_min),
        color_min=int(args.color_min),
        credentials_path=settings["credentials_path"],
        token_path=settings["token_path"],
        shop_id=shop_id,
    )
    summary["rakutenShopId"] = shop_id
    summary["sheets"] = {
        "spreadsheetId": info.get("spreadsheet_id") or settings["spreadsheet_id"],
        "masterSheet": info.get("master_sheet") or settings["master_sheet"],
    }

    meta_dir = work_root / "00.テスト出力" / "_meta"
    meta_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = meta_dir / f"BASE_BIND_{stamp}.json"
    latest = meta_dir / "BASE_BIND_latest.json"
    text = json.dumps(summary, ensure_ascii=False, indent=2)
    out_path.write_text(text, encoding="utf-8")
    latest.write_text(text, encoding="utf-8")

    print(
        json.dumps(
            {
                "summary": str(out_path),
                "counts": summary["counts"],
                "dryRun": summary["dryRun"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    for r in summary["results"]:
        print(
            f"  {r.get('status')}: {r.get('old_name') or r.get('base_file') or '-'} "
            f"-> {r.get('new_name') or '-'} parent={r.get('parent_sku') or '-'} "
            f"overall={r.get('overall')} reason={r.get('reason')}"
        )

    if args.require_bound:
        # 未割当PNGが残っていれば NG（already_bound / renamed 以外の base png）
        folder = work_root / SUB_AMAZON
        left = []
        parents = [p["parentSku"] for p in summary.get("parents") or []]
        parent_refs = [
            ParentRef(parent_sku=p) for p in parents
        ]
        for p in list_base_pngs(folder):
            if any_parent_token_in_name(p, parent_refs) is None:
                left.append(p.name)
        if left or summary["counts"]["error"] > 0:
            LOG.error("require-bound failed unbound=%s errors=%s", left, summary["counts"]["error"])
            return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
