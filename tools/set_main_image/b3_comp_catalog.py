# -*- coding: utf-8 -*-
"""
B-③「競合画像取得（必要時B-③実行）」シートの読取・URLキャッシュ。

API再取得はしない（シート上の画像URLのみ消費）。
"""
from __future__ import annotations

import hashlib
import logging
import re
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from sheets_master import _get_credentials, load_sheets_settings

LOG = logging.getLogger("set_main_image.b3_comp_catalog")

B3_SHEET_TITLE = "競合画像取得（必要時B-③実行）"
HEADERS = (
    "マスタ行",
    "JAN",
    "A.セット商品数",
    "区分",
    "ASINまたは商品URL",
    "画像番号",
    "画像URL",
    "プレビュー",
)

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)


@dataclass
class B3ImageRow:
    master_row: str
    jan: str
    set_qty: str
    kind: str
    listing_key: str
    image_index: int
    url: str

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def fetch_b3_rows(
    *,
    config_path: Optional[Path] = None,
    spreadsheet_id: str = "",
    sheet_title: str = B3_SHEET_TITLE,
) -> List[List[str]]:
    from googleapiclient.discovery import build

    settings = load_sheets_settings(
        config_path,
        spreadsheet_id=spreadsheet_id,
        master_sheet=sheet_title,  # reuse loader fields; title overridden below
    )
    creds = _get_credentials(settings["credentials_path"], settings["token_path"])
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    sid = settings["spreadsheet_id"]
    title = sheet_title
    rng = "'%s'" % title.replace("'", "''")
    data = (
        sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=sid,
            range=rng,
            majorDimension="ROWS",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
    )
    raw = data.get("values") or []
    if not raw:
        raise RuntimeError(f"B-③シートが空です: {title}")
    width = max(len(r) for r in raw)
    out: List[List[str]] = []
    for r in raw:
        padded = list(r) + [""] * (width - len(r))
        out.append(["" if c is None else str(c) for c in padded])
    LOG.info("B-③ loaded sheet=%r rows=%d", title, len(out))
    return out


def parse_b3_image_rows(rows: Sequence[Sequence[str]], *, jan: str = "") -> List[B3ImageRow]:
    if not rows:
        return []
    header = [str(c).strip() for c in rows[0]]
    # 列位置（ヘッダー名優先、なければ既定順）
    def col(name: str, default: int) -> int:
        try:
            return header.index(name)
        except ValueError:
            return default

    i_jan = col("JAN", 1)
    i_master = col("マスタ行", 0)
    i_set = col("A.セット商品数", 2)
    i_kind = col("区分", 3)
    i_key = col("ASINまたは商品URL", 4)
    i_idx = col("画像番号", 5)
    i_url = col("画像URL", 6)

    want = (jan or "").strip()
    out: List[B3ImageRow] = []
    for r in rows[1:]:
        if i_url >= len(r):
            continue
        url = str(r[i_url] or "").strip()
        if not url.startswith("http"):
            continue
        row_jan = str(r[i_jan] if i_jan < len(r) else "").strip()
        if want and row_jan != want:
            continue
        idx_raw = str(r[i_idx] if i_idx < len(r) else "").strip()
        try:
            image_index = int(re.sub(r"[^0-9]", "", idx_raw) or "0")
        except ValueError:
            image_index = 0
        out.append(
            B3ImageRow(
                master_row=str(r[i_master] if i_master < len(r) else "").strip(),
                jan=row_jan,
                set_qty=str(r[i_set] if i_set < len(r) else "").strip(),
                kind=str(r[i_kind] if i_kind < len(r) else "").strip(),
                listing_key=str(r[i_key] if i_key < len(r) else "").strip(),
                image_index=image_index,
                url=url,
            )
        )
    return out


def prefer_order(rows: List[B3ImageRow]) -> List[B3ImageRow]:
    """説明系サブ優先: Amazonかつ画像番号>=2 → 他モール画像番号>=2 → 残り。"""

    def key(r: B3ImageRow) -> tuple:
        is_amz = 0 if "Amazon" in r.kind else 1
        is_sub = 0 if r.image_index >= 2 else 1
        return (is_sub, is_amz, r.image_index, r.kind, r.url)

    return sorted(rows, key=key)


def url_cache_name(url: str) -> str:
    h = hashlib.sha1(url.encode("utf-8")).hexdigest()[:16]
    ext = ".jpg"
    low = url.lower()
    if ".png" in low:
        ext = ".png"
    elif ".webp" in low:
        ext = ".webp"
    return f"{h}{ext}"


def download_url(url: str, dest: Path, *, timeout: int = 40) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.is_file() and dest.stat().st_size > 100:
        return dest
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "image/*,*/*"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"download HTTP {e.code}: {url[:120]}") from e
    except Exception as e:
        raise RuntimeError(f"download failed: {url[:120]} ({e})") from e
    if not data or len(data) < 50:
        raise RuntimeError(f"download empty: {url[:120]}")
    dest.write_bytes(data)
    return dest


def cache_b3_images(
    rows: List[B3ImageRow],
    cache_dir: Path,
    *,
    limit: int = 12,
) -> List[Dict[str, Any]]:
    """選定済み行をDL。戻りは meta dict リスト（path付き）。"""
    cache_dir.mkdir(parents=True, exist_ok=True)
    out: List[Dict[str, Any]] = []
    for r in prefer_order(rows):
        if len(out) >= limit:
            break
        name = url_cache_name(r.url)
        path = cache_dir / name
        try:
            download_url(r.url, path)
        except Exception as e:
            LOG.warning("skip download: %s", e)
            continue
        meta = r.to_dict()
        meta["path"] = str(path)
        meta["cacheName"] = name
        out.append(meta)
    return out
