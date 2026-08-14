# -*- coding: utf-8 -*-
"""
PACKAGE_TRUTH（商品パッケージ改変不可正本）の解決と縦横比計測。

正本は B-④確認済みの単体／N=1相当画像（缶・瓶・箱・パウチ可）。
N≥2 のセットMAINコラージュは使わない。ファイル名にSKU必須ではない（JAN紐付けで足りる）。
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PIL import Image

from work_paths import SUB_AMAZON, TAG_UNIT, default_work_root

LOG = logging.getLogger("set_main_image.package_truth")

IMG_EXTS = {".jpg", ".jpeg", ".png", ".webp"}
# セットMAINっぽい命名を避ける（単体タグがあれば優先して通す）
_SET_MAIN_HINT = re.compile(
    r"(セットmain|set[_-]?main|_n[2-9]|_×\d|x\d+個|コラージュ)",
    re.IGNORECASE,
)


def parse_truth_specs(specs: Sequence[str]) -> Dict[str, Path]:
    """
    CLI `--package-truth` 用。
    - `JAN=path`
    - `path` のみ（単一JAN実行時に後から紐付け）
    """
    out: Dict[str, Path] = {}
    for raw in specs or []:
        s = str(raw or "").strip().strip('"')
        if not s:
            continue
        if "=" in s:
            jan, p = s.split("=", 1)
            jan = jan.strip()
            path = Path(p.strip().strip('"'))
            if jan:
                out[jan] = path
        else:
            out[""] = Path(s)
    return out


def _list_images(folder: Path, *, recursive_processed: bool = False) -> List[Path]:
    if not folder.is_dir():
        return []
    files = [
        p
        for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in IMG_EXTS
    ]
    if recursive_processed:
        proc = folder / "処理済み"
        if proc.is_dir():
            files.extend(
                [
                    p
                    for p in proc.iterdir()
                    if p.is_file() and p.suffix.lower() in IMG_EXTS
                ]
            )
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)


def _looks_like_set_main(path: Path) -> bool:
    name = path.stem
    if TAG_UNIT in name:
        return False
    return bool(_SET_MAIN_HINT.search(name))


def _score_candidate(path: Path, jan: str) -> Tuple[int, float]:
    """高いほど良い。(優先度, mtime)。"""
    name = path.name
    score = 0
    if jan and jan in name:
        score += 100
    if TAG_UNIT in path.stem:
        score += 50
    if _looks_like_set_main(path):
        score -= 200
    return score, path.stat().st_mtime


def resolve_package_truth(
    jan: str,
    *,
    work_root: Optional[Path] = None,
    explicit: Optional[Path] = None,
    truth_dir: Optional[Path] = None,
    truth_map: Optional[Dict[str, Path]] = None,
) -> Optional[Path]:
    """
    正本パスを解決。無ければ None（呼び出し側で JAN スキップ）。
    探索順: 明示 map → explicit → truth_dir → 01.amazon白抜きベース。
    """
    j = str(jan or "").strip()
    root = Path(work_root) if work_root else default_work_root()

    if truth_map:
        if j in truth_map and truth_map[j].is_file():
            return truth_map[j]
        anon = truth_map.get("")
        if anon and anon.is_file() and len(truth_map) == 1:
            return anon

    if explicit and Path(explicit).is_file():
        return Path(explicit)

    cands: List[Path] = []
    if truth_dir and Path(truth_dir).is_dir():
        cands.extend(_list_images(Path(truth_dir), recursive_processed=True))
    amazon_dir = root / SUB_AMAZON
    cands.extend(_list_images(amazon_dir, recursive_processed=True))

    # 重複除去
    seen = set()
    uniq: List[Path] = []
    for p in cands:
        key = str(p.resolve())
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p)

    ranked = sorted(uniq, key=lambda p: _score_candidate(p, j), reverse=True)
    for p in ranked:
        sc, _ = _score_candidate(p, j)
        if j:
            # 誤商品防止: JAN指定時はファイル名にJANが必須（単体タグだけでは不可）
            if j not in p.name:
                continue
        elif sc < 50:
            continue
        if _looks_like_set_main(p) and TAG_UNIT not in p.stem:
            LOG.warning("PACKAGE_TRUTH候補をセットMAIN疑いのため除外: %s", p.name)
            continue
        LOG.info("PACKAGE_TRUTH resolved jan=%s path=%s score=%s", j, p, sc)
        return p
    return None


def measure_package_aspect(path: Path) -> Dict[str, Any]:
    """
    正本の不透明／非白領域 AABB から縦横比 R=W/H を計測。
    角度はそのままでよく、アスペクトだけロックする用途。
    """
    p = Path(path)
    im = Image.open(p)
    w, h = im.size
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        rgba = im.convert("RGBA")
        alpha = rgba.split()[-1]
        mask = alpha.point(lambda a: 255 if a >= 12 else 0)
        method = "alpha"
    else:
        from PIL import ImageChops

        rgb = im.convert("RGB")
        # 近白以外を商品シルエットとみなす（差分→輝度閾値）
        white = Image.new("RGB", (w, h), (255, 255, 255))
        diff = ImageChops.difference(rgb, white).convert("L")
        mask = diff.point(lambda v: 255 if v >= 10 else 0)
        method = "non_near_white"

    bbox = mask.getbbox()
    if not bbox:
        ratio = float(w) / float(h) if h else 1.0
        return {
            "path": str(p),
            "imageW": w,
            "imageH": h,
            "bbox": None,
            "aspectWH": round(ratio, 4),
            "method": method + "_fallback_full",
        }
    x0, y0, x1, y1 = bbox
    bw = max(1, x1 - x0)
    bh = max(1, y1 - y0)
    ratio = float(bw) / float(bh)
    return {
        "path": str(p),
        "imageW": w,
        "imageH": h,
        "bbox": [int(x0), int(y0), int(x1), int(y1)],
        "bboxW": int(bw),
        "bboxH": int(bh),
        "aspectWH": round(ratio, 4),
        "method": method,
    }


def format_aspect_lock_ja(aspect: Dict[str, Any]) -> str:
    r = aspect.get("aspectWH")
    if r is None:
        return ""
    return (
        f"【PACKAGE_ASPECT_LOCK】正本AABBの縦横比 R=W/H={r} "
        f"（計測={aspect.get('method')}）。"
        "パッケージの縦横比を変えないこと。潰し・引き延ばし禁止。"
        "軽い角度は可。アスペクトが崩れる変形は禁止。"
    )
