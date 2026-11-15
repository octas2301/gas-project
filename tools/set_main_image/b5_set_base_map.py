# -*- coding: utf-8 -*-
"""B-⑤ P3b: {親SKU}_set_base_map.json の読書き・検証・02解決。

文字色（ALIGN §3）: 親SKU単位の共有ストアを正本とする。
解決順: 共有 `_b5_text_color_by_parent.json` → set_base_map → フォールバック。
マップの textColor は共有のミラー（書込時に同期／読取時に共有へ昇格可）。
"""
from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

LOG = logging.getLogger("set_main_image.b5_set_base_map")

MAP_SUFFIX = "_set_base_map.json"
SHARED_TEXT_COLOR_NAME = "_b5_text_color_by_parent.json"


def map_path(rakuten_dir: Path, parent_sku: str) -> Path:
    return Path(rakuten_dir) / f"{str(parent_sku).strip()}{MAP_SUFFIX}"


def load_map(rakuten_dir: Path, parent_sku: str) -> Optional[Dict[str, Any]]:
    p = map_path(rakuten_dir, parent_sku)
    if p.is_file():
        return json.loads(p.read_text(encoding="utf-8-sig"))
    # C量産後は 02/処理済み へ移るため、共有昇格・再解決用に読む
    done = Path(rakuten_dir) / "処理済み" / f"{str(parent_sku).strip()}{MAP_SUFFIX}"
    if done.is_file():
        try:
            return json.loads(done.read_text(encoding="utf-8-sig"))
        except Exception:
            return None
    return None


def seed_shared_text_colors_from_maps(
    work_root: Optional[Path] = None,
    *,
    include_processed: bool = True,
) -> Dict[str, str]:
    """02直下＋処理済みのマップ textColor を共有へマージ（既存運用の救済）。"""
    from work_paths import SUB_RAKUTEN, default_work_root

    root = Path(work_root) if work_root else default_work_root()
    rak = root / SUB_RAKUTEN
    merged = load_shared_text_colors(root)
    dirs = [rak]
    if include_processed:
        dirs.append(rak / "処理済み")
    for d in dirs:
        if not d.is_dir():
            continue
        for p in d.glob(f"*{MAP_SUFFIX}"):
            try:
                data = json.loads(p.read_text(encoding="utf-8-sig"))
            except Exception:
                continue
            if not isinstance(data, dict):
                continue
            sku = str(data.get("parentSku") or p.name.replace(MAP_SUFFIX, "")).strip()
            if not sku or data.get("textColor") in (None, ""):
                continue
            if sku in merged:
                continue
            merged[sku] = normalize_text_color(data.get("textColor"))
    if merged:
        path = shared_text_color_path(root)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
        LOG.info("seed shared textColor n=%s path=%s", len(merged), path)
    return merged


def adopt_key_from_filename(name: str) -> str:
    """絵ごとのキー。テーマ番号が同じでも見出しが違えば別。再作成は末尾 __rN。"""
    stem = Path(name).stem
    rm = re.search(r"__r(\d+)$", stem)
    regen = f"__r{rm.group(1)}" if rm else ""
    body = stem[: rm.start()] if rm else stem
    m = re.search(r"_([NT]\d+[a-z]?)(?:_|$)", body)
    tid = m.group(1) if m else "X"
    tail = body[m.end() :].strip("_") if m else ""
    if tail:
        return f"{tid}__{tail}{regen}"
    return f"{tid}{regen}"


def picture_family_key(adopt_key: str) -> str:
    """再作成前後を同じ絵と見る。見出しが違えば別。"""
    return re.sub(r"__r\d+$", "", str(adopt_key or "").strip())


def next_regen_filename(name: str) -> str:
    """上書きせず世代を増やす: foo.jpg → foo__r2.jpg → foo__r3.jpg"""
    p = Path(name)
    stem, suf = p.stem, p.suffix or ".jpg"
    m = re.search(r"__r(\d+)$", stem)
    if m:
        n = int(m.group(1)) + 1
        base = stem[: m.start()]
        return f"{base}__r{n}{suf}"
    return f"{stem}__r2{suf}"


def validate_map(
    *,
    by_child: Dict[str, str],
    adopted: Dict[str, str],
    required_children: Sequence[str],
) -> List[str]:
    """エラーメッセージ一覧。空ならOK。"""
    errs: List[str] = []
    req = [str(c).strip() for c in required_children if str(c).strip()]
    if not adopted:
        errs.append("採用ファイルが0件です")
    seen_child: Dict[str, str] = {}
    for ch, key in (by_child or {}).items():
        c = str(ch).strip()
        k = str(key).strip()
        if not c or not k:
            continue
        if c in seen_child and seen_child[c] != k:
            errs.append(f"子SKU二重割当: {c}")
        seen_child[c] = k
        if k not in adopted:
            errs.append(f"adoptKey={k} が adopted に無い（子 {c}）")
    missing = [c for c in req if c not in seen_child]
    if missing:
        errs.append("レ点子の割当漏れ: " + ", ".join(missing[:12]) + ("…" if len(missing) > 12 else ""))
    extra = [c for c in seen_child if c not in req]
    if extra:
        errs.append("レ点に無い子がマップにある: " + ", ".join(extra[:8]))
    cap = len(req)
    if cap and len(adopted) > cap:
        errs.append(
            f"採用版が多すぎます（{len(adopted)}）。最大はレ点子の数（{cap}）"
        )
    return errs


def normalize_text_color(value: Any, default: str = "1") -> str:
    """C／マップ用。'1'..'5' 以外は default。"""
    s = str(value or "").strip()
    if s in ("1", "2", "3", "4", "5"):
        return s
    return default if default in ("1", "2", "3", "4", "5") else "1"


def text_color_from_map(mapping: Optional[Dict[str, Any]], default: str = "1") -> str:
    if not mapping:
        return normalize_text_color(default)
    return normalize_text_color(mapping.get("textColor"), default=default)


def shared_text_color_path(work_root: Optional[Path] = None) -> Path:
    from work_paths import default_work_root

    root = Path(work_root) if work_root else default_work_root()
    return root / SHARED_TEXT_COLOR_NAME


def load_shared_text_colors(work_root: Optional[Path] = None) -> Dict[str, str]:
    p = shared_text_color_path(work_root)
    if not p.is_file():
        return {}
    try:
        obj = json.loads(p.read_text(encoding="utf-8-sig"))
    except Exception as e:
        LOG.warning("shared textColor load fail: %s", e)
        return {}
    if not isinstance(obj, dict):
        return {}
    out: Dict[str, str] = {}
    for k, v in obj.items():
        sku = str(k or "").strip()
        tc = normalize_text_color(v, default="")
        if sku and tc:
            out[sku] = tc
    return out


def save_shared_text_color(
    parent_sku: str,
    text_color: Any,
    *,
    work_root: Optional[Path] = None,
) -> Path:
    """親SKUの文字色を共有ストアへマージ保存。"""
    sku = str(parent_sku or "").strip()
    tc = normalize_text_color(text_color)
    if not sku:
        raise ValueError("parent_sku required")
    root = Path(work_root) if work_root else None
    if root is None:
        from work_paths import default_work_root

        root = default_work_root()
    path = shared_text_color_path(root)
    data = load_shared_text_colors(root)
    data[sku] = tc
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info("shared textColor saved parent=%s tc=%s path=%s", sku, tc, path)
    return path


def resolve_text_color_for_parent(
    parent_sku: str,
    *,
    work_root: Optional[Path] = None,
    rakuten_dir: Optional[Path] = None,
    default: str = "1",
    promote_map_to_shared: bool = True,
) -> Tuple[str, str]:
    """
    共有 → マップ → default。
    戻り値: (textColor, source) source は shared|map|default|none
    マップにあり共有に無いとき promote_map_to_shared なら共有へ昇格。
    """
    from work_paths import SUB_RAKUTEN, default_work_root

    sku = str(parent_sku or "").strip()
    root = Path(work_root) if work_root else default_work_root()
    rak = Path(rakuten_dir) if rakuten_dir else (root / SUB_RAKUTEN)
    if not sku:
        return normalize_text_color(default), "default"
    shared = load_shared_text_colors(root)
    if sku in shared:
        return shared[sku], "shared"
    m = load_map(rak, sku)
    if m and m.get("textColor") not in (None, ""):
        tc = text_color_from_map(m, default=default)
        if promote_map_to_shared:
            try:
                save_shared_text_color(sku, tc, work_root=root)
            except Exception as e:
                LOG.warning("promote map→shared fail: %s", e)
        return tc, "map"
    return normalize_text_color(default), "default"


def resolve_text_color_for_jan(
    jan: str,
    *,
    work_root: Optional[Path] = None,
    parent_sku: Optional[str] = None,
    default: str = "1",
) -> str:
    """
    文字色解決（ALIGN §3）: 共有 → マップ → default。
    親SKU省略時は ishi-{JAN}-oya と *{JAN}*_set_base_map.json を試す。
    """
    from work_paths import SUB_RAKUTEN, default_work_root

    j = str(jan or "").strip()
    root = Path(work_root) if work_root else default_work_root()
    rak = root / SUB_RAKUTEN
    tried: List[str] = []
    skus: List[str] = []
    ps = str(parent_sku or "").strip()
    if ps:
        skus.append(ps)
    if j:
        skus.append(f"ishi-{j}-oya")
    for sku in skus:
        if not sku or sku in tried:
            continue
        tried.append(sku)
        tc, src = resolve_text_color_for_parent(
            sku,
            work_root=root,
            rakuten_dir=rak,
            default=default,
            promote_map_to_shared=True,
        )
        if src in ("shared", "map"):
            return tc
    if rak.is_dir() and j:
        for p in sorted(rak.glob(f"*{j}*{MAP_SUFFIX}")):
            try:
                m = json.loads(p.read_text(encoding="utf-8-sig"))
                if m.get("textColor") not in (None, ""):
                    tc = text_color_from_map(m, default=default)
                    parent = str(m.get("parentSku") or p.name.replace(MAP_SUFFIX, "")).strip()
                    if parent:
                        try:
                            save_shared_text_color(parent, tc, work_root=root)
                        except Exception:
                            pass
                    return tc
            except Exception:
                continue
    return normalize_text_color(default)


def parse_text_color_for_specs(specs: Optional[Sequence[str]]) -> Dict[str, str]:
    """CLI `--text-color-for JAN=5` → {JAN: '5'}。"""
    out: Dict[str, str] = {}
    for raw in specs or []:
        s = str(raw or "").strip()
        if "=" not in s:
            continue
        j, v = s.split("=", 1)
        j = j.strip()
        tc = normalize_text_color(v.strip(), default="")
        if j and tc:
            out[j] = tc
    return out


def write_map(
    rakuten_dir: Path,
    *,
    parent_sku: str,
    by_child: Dict[str, str],
    adopted: Dict[str, str],
    extra: Optional[Dict[str, Any]] = None,
    text_color: Optional[str] = None,
    work_root: Optional[Path] = None,
    sync_shared: bool = True,
) -> Path:
    sku = str(parent_sku).strip()
    payload: Dict[str, Any] = {
        "parentSku": sku,
        "version": 1,
        "byChildSku": dict(by_child),
        "adopted": dict(adopted),
    }
    if extra:
        payload.update(extra)
    if text_color is not None:
        payload["textColor"] = normalize_text_color(text_color)
    elif "textColor" in payload:
        payload["textColor"] = normalize_text_color(payload.get("textColor"))
    dest = map_path(rakuten_dir, sku)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    if sync_shared and payload.get("textColor"):
        try:
            from work_paths import default_work_root

            rak = Path(rakuten_dir)
            if work_root is not None:
                root = Path(work_root)
            elif rak.name == "02.楽天ベース":
                root = rak.parent
            else:
                root = default_work_root()
            save_shared_text_color(sku, payload["textColor"], work_root=root)
        except Exception as e:
            LOG.warning("write_map sync shared fail: %s", e)
    return dest


def resolve_base_for_child(
    rakuten_dir: Path,
    mapping: Dict[str, Any],
    child_sku: str,
) -> Optional[Path]:
    """マップから子の土台 Path。無ければ None（呼び出し側でフォールバック）。"""
    by = mapping.get("byChildSku") or {}
    adopted = mapping.get("adopted") or {}
    key = by.get(str(child_sku).strip())
    if not key:
        return None
    fname = adopted.get(str(key).strip())
    if not fname:
        return None
    p = Path(rakuten_dir) / Path(str(fname)).name
    return p if p.is_file() else None


def list_map_related_files(rakuten_dir: Path, parent_sku: str) -> List[Path]:
    """処理済み移動用: マップ＋adopted 参照ファイル＋互換 _単体.png。"""
    out: List[Path] = []
    m = load_map(rakuten_dir, parent_sku)
    mp = map_path(rakuten_dir, parent_sku)
    if mp.is_file():
        out.append(mp)
    if m:
        for fname in (m.get("adopted") or {}).values():
            p = Path(rakuten_dir) / Path(str(fname)).name
            if p.is_file():
                out.append(p)
    legacy = Path(rakuten_dir) / f"{str(parent_sku).strip()}_単体.png"
    if legacy.is_file() and legacy not in out:
        out.append(legacy)
    # テーマ別でマップ外に残っているものも親プレフィックスで拾う
    pref = f"{str(parent_sku).strip()}_単体"
    for p in Path(rakuten_dir).glob(f"{pref}*.png"):
        if p not in out:
            out.append(p)
    return out
