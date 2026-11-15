# -*- coding: utf-8 -*-
"""
見本ゲート向け・人間手動投入（コピペ優先／URL・ファイル予備）。

- 常時任意・最大 MANUAL_MAX 枚
- B-③競合ゼロ時も同じ保存先（案内は UI 側）
"""
from __future__ import annotations

import base64
import hashlib
import json
import logging
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from b3_comp_catalog import download_url

LOG = logging.getLogger("set_main_image.food_canon_manual")

MANUAL_MAX = 6
MANUAL_PREFIX = "MANUAL_"
MANUAL_META = "manual_ingest.json"


def _refs(out_dir: Path) -> Path:
    r = Path(out_dir) / "01_refs"
    r.mkdir(parents=True, exist_ok=True)
    return r


def list_manual_paths(out_dir: Path) -> List[Path]:
    refs = _refs(out_dir)
    out: List[Path] = []
    for i in range(1, MANUAL_MAX + 1):
        for ext in (".jpg", ".jpeg", ".png", ".webp"):
            p = refs / f"{MANUAL_PREFIX}{i:02d}{ext}"
            if p.is_file():
                out.append(p)
                break
    return out


def load_manual_meta(out_dir: Path) -> Dict[str, Any]:
    p = Path(out_dir) / MANUAL_META
    if not p.is_file():
        return {"items": [], "max": MANUAL_MAX}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {"items": [], "max": MANUAL_MAX}


def _save_manual_meta(out_dir: Path, meta: Dict[str, Any]) -> None:
    meta = {**meta, "max": MANUAL_MAX, "updatedAt": datetime.now(timezone.utc).isoformat()}
    (Path(out_dir) / MANUAL_META).write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _next_slot(out_dir: Path) -> int:
    used = set()
    for p in list_manual_paths(out_dir):
        m = re.search(r"MANUAL_(\d+)", p.name, re.I)
        if m:
            used.add(int(m.group(1)))
    for i in range(1, MANUAL_MAX + 1):
        if i not in used:
            return i
    raise RuntimeError(f"手動画像は最大{MANUAL_MAX}枚です。不要なものを削除してください。")


def _bytes_to_jpeg(data: bytes, dest: Path) -> Path:
    dest = dest.with_suffix(".jpg")
    try:
        from io import BytesIO

        from PIL import Image

        im = Image.open(BytesIO(data))
        im = im.convert("RGB")
        # 長辺上限
        w, h = im.size
        m = max(w, h)
        if m > 2048:
            scale = 2048 / m
            im = im.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
        dest.parent.mkdir(parents=True, exist_ok=True)
        im.save(dest, quality=92)
        return dest
    except Exception:
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(data)
        return dest


def save_manual_bytes(
    out_dir: Path,
    data: bytes,
    *,
    source: str = "paste",
    note: str = "",
) -> Dict[str, Any]:
    if not data or len(data) < 50:
        raise ValueError("画像データが空です")
    slot = _next_slot(out_dir)
    refs = _refs(out_dir)
    dest = refs / f"{MANUAL_PREFIX}{slot:02d}.jpg"
    path = _bytes_to_jpeg(data, dest)
    meta = load_manual_meta(out_dir)
    items = list(meta.get("items") or [])
    items.append(
        {
            "slot": slot,
            "name": path.name,
            "path": str(path.resolve()),
            "source": source,
            "note": note[:200],
            "at": datetime.now(timezone.utc).isoformat(),
        }
    )
    meta["items"] = items
    _save_manual_meta(out_dir, meta)
    LOG.info("manual save slot=%s source=%s -> %s", slot, source, path.name)
    return {"ok": True, "slot": slot, "name": path.name, "path": str(path)}


def save_manual_data_url(out_dir: Path, data_url: str) -> Dict[str, Any]:
    """data:image/png;base64,... """
    s = (data_url or "").strip()
    m = re.match(r"^data:image/([\w+.-]+);base64,(.+)$", s, re.I | re.S)
    if not m:
        raise ValueError("data URL 形式ではありません")
    raw = base64.b64decode(m.group(2))
    return save_manual_bytes(out_dir, raw, source="paste")


def save_manual_url(out_dir: Path, url: str) -> Dict[str, Any]:
    url = (url or "").strip()
    if not url.startswith("http"):
        raise ValueError("http(s) の画像URLを指定してください")
    refs = _refs(out_dir)
    slot = _next_slot(out_dir)
    tmp = refs / f"_tmp_manual_{hashlib.sha1(url.encode()).hexdigest()[:10]}.bin"
    download_url(url, tmp)
    data = tmp.read_bytes()
    tmp.unlink(missing_ok=True)
    return save_manual_bytes(out_dir, data, source="url", note=url[:180])


def delete_manual_slot(out_dir: Path, slot: int) -> Dict[str, Any]:
    refs = _refs(out_dir)
    removed = []
    for ext in (".jpg", ".jpeg", ".png", ".webp"):
        p = refs / f"{MANUAL_PREFIX}{int(slot):02d}{ext}"
        if p.is_file():
            p.unlink()
            removed.append(p.name)
    meta = load_manual_meta(out_dir)
    meta["items"] = [x for x in (meta.get("items") or []) if int(x.get("slot") or 0) != int(slot)]
    _save_manual_meta(out_dir, meta)
    return {"ok": True, "removed": removed}


def merge_manual_into_candidates(
    candidates: List[Dict[str, Any]], out_dir: Path
) -> List[Dict[str, Any]]:
    """手動画像を候補先頭付近に足す（重複名は除外）。"""
    names = {str(c.get("name") or "") for c in candidates}
    out = list(candidates)
    for p in list_manual_paths(out_dir):
        if p.name in names:
            continue
        out.insert(
            0,
            {
                "index": -1,
                "path": str(p.resolve()),
                "name": p.name,
                "score": 100,
                "reasonJa": "人間が手動投入（コピペ／URL／ファイル）",
                "suggestedPrimary": len(out) == 0,
                "source": "manual",
            },
        )
        names.add(p.name)
    # index 振り直し
    for i, c in enumerate(out):
        c["index"] = i
    return out


def count_b3_style_comps(out_dir: Path) -> int:
    """MANUAL / PACKAGE / FOOD_UNIT 以外の 01_refs 枚数。"""
    refs = _refs(out_dir)
    n = 0
    for p in refs.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in (".jpg", ".jpeg", ".png", ".webp"):
            continue
        u = p.name.upper()
        if u.startswith(MANUAL_PREFIX):
            continue
        if u.startswith("PACKAGE_TRUTH") or u.startswith("FOOD_UNIT"):
            continue
        n += 1
    return n
