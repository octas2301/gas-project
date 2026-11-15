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


def amazon_base_dir(work_root: Optional[Path] = None) -> Path:
    root = Path(work_root) if work_root else default_work_root()
    return root / SUB_AMAZON


def rel_under_amazon_base(path: Path, *, work_root: Optional[Path] = None) -> str:
    """01.amazon白抜きベース からの相対（処理済み\\name.png）。GAS紐付けと同じ表記。"""
    amazon = amazon_base_dir(work_root)
    p = Path(path)
    try:
        rel = p.resolve().relative_to(amazon.resolve())
        return str(rel).replace("/", "\\")
    except Exception:
        name = p.name
        if (amazon / "処理済み" / name).is_file():
            return f"処理済み\\{name}"
        return name


def _other_parent_bag_blocks(
    *,
    parent_sku: str = "",
    work_root: Optional[Path] = None,
) -> Dict[str, set]:
    """他親に紐付いた袋を落とすためのキー。

    - rels: packageTruthRel 全文＋ファイル名だけ
    - skus: 他親の parentSku
    - jans: 他親の JAN（ファイル名に含まれる単体・ミックスも落とす）
    """
    sku = str(parent_sku or "").strip()
    out: Dict[str, set] = {"rels": set(), "skus": set(), "jans": set()}
    try:
        from b5_parent_shared import load_parent_shared
    except Exception:
        return out
    by = load_parent_shared(work_root).get("byParentSku") or {}
    if not isinstance(by, dict):
        return out
    for other, row in by.items():
        if not isinstance(row, dict):
            continue
        other_s = str(other or "").strip()
        if not other_s or (sku and other_s == sku):
            continue
        if other_s.lower() == "test":
            continue
        out["skus"].add(other_s.lower())
        rel = str(row.get("packageTruthRel") or "").strip().replace("/", "\\")
        if rel:
            out["rels"].add(rel.lower())
            out["rels"].add(Path(rel).name.lower())
        jan = "".join(ch for ch in str(row.get("jan") or "") if ch.isdigit())
        if len(jan) >= 8:
            out["jans"].add(jan)
    return out


def _rels_linked_by_other_parents(
    *,
    parent_sku: str = "",
    work_root: Optional[Path] = None,
) -> set:
    """互換: 他親の packageTruthRel 集合。"""
    return set(
        _other_parent_bag_blocks(parent_sku=parent_sku, work_root=work_root).get("rels")
        or []
    )


def list_package_truth_candidates(
    *,
    jan: str = "",
    parent_sku: str = "",
    name_base: str = "",
    current: Optional[Path] = None,
    work_root: Optional[Path] = None,
    limit: int = 16,
    hide_other_parents: bool = True,
) -> List[Dict[str, Any]]:
    """
    袋正本の候補（サムネ選択用）。
    出どころは GAS ① と同じ 01.amazon白抜きベース（処理済み含む）。
    hide_other_parents=True のとき、他親が紐付け済みの袋は出さない（いま選んでいる袋は残す）。
    判定は相対パス完全一致に加え、同名ファイル・他親SKU・他親JANがファイル名に含まれるもの。
    """
    j = str(jan or "").strip()
    sku = str(parent_sku or "").strip()
    nb = str(name_base or "").strip()
    amazon = amazon_base_dir(work_root)
    files = _list_images(amazon, recursive_processed=True)
    cur = Path(current) if current and Path(current).is_file() else None
    cur_key = str(cur.resolve()) if cur is not None else ""
    blocks = (
        _other_parent_bag_blocks(parent_sku=sku, work_root=work_root)
        if hide_other_parents
        else {"rels": set(), "skus": set(), "jans": set()}
    )
    # いまの親の JAN は落とさない（同JANの別ファイル候補を残す）
    self_jan = "".join(ch for ch in j if ch.isdigit())
    if self_jan:
        blocks["jans"].discard(self_jan)
    # いまの親SKUは落とさない
    if sku:
        blocks["skus"].discard(sku.lower())

    ranked: List[Tuple[Tuple[int, float], Path]] = []
    for p in files:
        if _looks_like_set_main(p) and TAG_UNIT not in p.stem:
            continue
        if nb:
            # チーズ等の明らかな混入だけ落とす（food_unit_canon と同趣旨）
            if "チーズ" in p.name and "チーズ" not in nb:
                continue
        sc, mt = _score_candidate(p, j)
        if sku and p.stem.lower() == f"{sku}_単体".lower():
            sc += 80
        if cur_key and str(p.resolve()) == cur_key:
            sc += 40
        ranked.append(((sc, mt), p))
    ranked.sort(key=lambda x: x[0], reverse=True)

    def _row(p: Path, sc: int) -> Dict[str, Any]:
        key = str(p.resolve())
        return {
            "rel": rel_under_amazon_base(p, work_root=work_root),
            "path": str(p),
            "name": p.name,
            "score": int(sc),
            "selected": bool(cur_key and key == cur_key),
        }

    def _blocked(p: Path) -> bool:
        if cur_key and str(p.resolve()) == cur_key:
            return False
        rel = rel_under_amazon_base(p, work_root=work_root).replace("/", "\\")
        rel_l = rel.lower()
        name_l = p.name.lower()
        stem_l = p.stem.lower()
        if rel_l in blocks["rels"] or name_l in blocks["rels"]:
            return True
        for other_sku in blocks["skus"]:
            if other_sku and other_sku in stem_l:
                return True
        for other_jan in blocks["jans"]:
            if other_jan and other_jan in name_l:
                # いまの親SKUが同じJAN枝を含むファイルは残す
                if sku and sku.lower() in stem_l:
                    continue
                return True
        return False

    def _related(p: Path) -> bool:
        if cur_key and str(p.resolve()) == cur_key:
            return True
        if j and j in p.name:
            return True
        if sku and sku.lower() in p.stem.lower():
            return True
        return False

    # 親SKU／JANがあるときは関連ファイルを先に（誤選防止）。足りなければ全体へ広げる。
    related = [
        ((sc, mt), p) for (sc, mt), p in ranked if _related(p) and not _blocked(p)
    ]
    pool = related if related else [
        ((sc, mt), p) for (sc, mt), p in ranked if not _blocked(p)
    ]

    out: List[Dict[str, Any]] = []
    seen = set()
    for (sc, _mt), p in pool:
        key = str(p.resolve())
        if key in seen or _blocked(p):
            continue
        seen.add(key)
        out.append(_row(p, sc))
        if len(out) >= max(1, int(limit)):
            break
    if len(out) < 2 and related and len(related) < len(ranked):
        for (sc, _mt), p in ranked:
            key = str(p.resolve())
            if key in seen or _blocked(p):
                continue
            seen.add(key)
            out.append(_row(p, sc))
            if len(out) >= max(1, int(limit)):
                break
    if cur is not None and cur_key and not any(c.get("selected") for c in out):
        out.insert(0, _row(cur, 999))
        out = out[: max(1, int(limit))]
    return out


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
