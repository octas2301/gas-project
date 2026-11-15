# -*- coding: utf-8 -*-
"""食材見本ゲート — 一片見本レーン（個包装と同型の有無・貼付・レ点）。

案件名を焼かない。保存先は theme_new_run 配下のみ。
切戻し: 環境変数 B5_PIECE_GATE_UI=0 でUI非表示（ゲート本体は既存のまま）。
"""
from __future__ import annotations

import base64
import json
import os
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
from urllib.request import urlopen

from b5_piece_truth_rules import write_piece_set_source

HUMAN_DIR_REL = Path("01_refs") / "PIECE_TRUTH_HUMAN"
LIBRARY_DIR_REL = Path("01_refs") / "PIECE_TRUTH_LIBRARY"
SET_DIR_REL = Path("food_piece_set5_virt") / "piece_set"
SET_OFF_REL = Path("food_piece_set5_virt") / "piece_set_off"
MAX_PIECE = 5
MAX_SEED = 8


def piece_gate_ui_enabled() -> bool:
    raw = str(os.environ.get("B5_PIECE_GATE_UI") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def human_truth_dir(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / HUMAN_DIR_REL


def library_dir(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / LIBRARY_DIR_REL


def list_human_piece_files(theme_new_dir: Path) -> List[Path]:
    d = human_truth_dir(theme_new_dir)
    if not d.is_dir():
        return []
    files: List[Path] = []
    for pat in ("piece_truth_0*.png", "piece_truth_0*.jpg", "piece_truth_0*.jpeg", "piece_truth_0*.webp"):
        files.extend(d.glob(pat))
    return sorted({p.resolve() for p in files if p.is_file()})


def list_library_piece_files(theme_new_dir: Path) -> List[Path]:
    d = library_dir(theme_new_dir)
    if not d.is_dir():
        return []
    files: List[Path] = []
    for pat in (
        "piece_lib_*.png",
        "piece_lib_*.jpg",
        "piece_lib_*.jpeg",
        "piece_lib_*.webp",
        "piece_truth_0*.png",
        "piece_truth_0*.jpg",
        "piece_truth_0*.jpeg",
        "piece_truth_0*.webp",
    ):
        files.extend(d.glob(pat))
    # 作り直し退避（piece_lib_arch_*）は一覧に出さない（ゴミが増えて上限に見える）
    out = []
    for p in files:
        if not p.is_file():
            continue
        if re.match(r"^piece_lib_arch_", p.stem, re.I):
            continue
        out.append(p)
    return sorted({p.resolve() for p in out}, key=lambda p: p.name)


def _piece_slot_key(path: Path) -> str:
    stem = path.stem
    m = re.match(r"^piece_(?:truth|lib)_(\d+)$", stem, re.I)
    if m:
        return f"slot:{int(m.group(1)):02d}"
    return f"name:{path.name}"


def list_piece_display_files(theme_new_dir: Path) -> List[Path]:
    """ゲート③表示用。ライブラリ＋HUMAN。

    同じスロットは HUMAN 優先（二重表示しない）。
    ソフト削除で HUMAN から外れたライブラリ分は残す。
    作り直し退避 arch は出さない。
    """
    by_key: Dict[str, Path] = {}
    for p in list_library_piece_files(theme_new_dir):
        by_key[_piece_slot_key(p)] = p
    for p in list_human_piece_files(theme_new_dir):
        by_key[_piece_slot_key(p)] = p
    return sorted(by_key.values(), key=lambda p: (_piece_slot_key(p), p.name))


def list_piece_set_files(theme_new_dir: Path, *, max_n: int = 0) -> List[Path]:
    d = Path(theme_new_dir) / SET_DIR_REL
    if not d.is_dir():
        return []
    files = sorted(d.glob("piece_set_0*.jpg"))
    n = int(max_n or 0)
    if n > 0:
        files = files[:n]
    return files


def discover_human_seed_paths(theme_new_dir: Path) -> List[Path]:
    """人間が渡した生成用の元見本（種）。PIECE_SEED_HUMAN*（複数可・採用はラジオ1枚）。"""
    refs = Path(theme_new_dir) / "01_refs"
    if not refs.is_dir():
        return []
    numbered = sorted(
        p for p in refs.glob("PIECE_SEED_HUMAN_0*.*") if p.is_file()
    )
    if numbered:
        return [p.resolve() for p in numbered]
    out: List[Path] = []
    for pat in ("PIECE_SEED_HUMAN.*", "piece_seed_0*.*"):
        out.extend(p for p in refs.glob(pat) if p.is_file())
    seen = set()
    uniq: List[Path] = []
    for p in sorted(out, key=lambda x: x.name):
        key = str(p.resolve())
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p.resolve())
    return uniq


def _next_seed_slot(theme_new_dir: Path) -> int:
    refs = Path(theme_new_dir) / "01_refs"
    used = set()
    for p in refs.glob("PIECE_SEED_HUMAN_0*.*"):
        try:
            used.add(int(p.stem.split("_")[-1]))
        except Exception:
            pass
    for i in range(1, MAX_SEED + 1):
        if i not in used:
            return i
    raise ValueError(f"元見本は最大{MAX_SEED}枚です。不要な種を消してから追加してください。")


def activate_piece_seed(theme_new_dir: Path, name: str) -> Dict[str, Any]:
    """ラジオで選んだ1枚を piece_set_source の正にする。"""
    theme_new_dir = Path(theme_new_dir)
    name = str(name or "").strip()
    if not name:
        raise ValueError("元見本名が空です")
    refs = theme_new_dir / "01_refs"
    src = refs / name
    if not src.is_file():
        for p in discover_human_seed_paths(theme_new_dir):
            if p.name == name:
                src = p
                break
    if not src.is_file():
        raise ValueError(f"元見本が見つかりません: {name}")
    write_piece_set_source(
        theme_new_dir,
        kind="human",
        paths=[src],
        extra={
            "note": "active seed selected in gate",
            "seedOnly": True,
            "activeSeed": src.name,
            "builtAt": datetime.now(timezone.utc).isoformat(),
        },
    )
    return {"ok": True, "name": src.name, "path": str(src)}


def install_human_seed_bytes(
    theme_new_dir: Path,
    raw: bytes,
    *,
    suffix: str = ".png",
    slot: int = 0,
    replace_all: bool = False,
    activate: bool = True,
) -> Dict[str, Any]:
    """元見本を登録。既定は追記（複数登録可）。採用は1枚（activate）。"""
    theme_new_dir = Path(theme_new_dir)
    refs = theme_new_dir / "01_refs"
    refs.mkdir(parents=True, exist_ok=True)
    suf = suffix if str(suffix).startswith(".") else f".{suffix}"
    if suf.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
        suf = ".png"
    if replace_all:
        for old in list(refs.glob("PIECE_SEED_HUMAN*")) + list(refs.glob("piece_seed_0*")):
            try:
                if old.is_file():
                    old.unlink()
            except Exception:
                pass
        slot = 1
    if int(slot or 0) < 1:
        slot = _next_seed_slot(theme_new_dir)
    dest = refs / f"PIECE_SEED_HUMAN_{int(slot):02d}{suf}"
    dest.write_bytes(raw)
    if activate:
        alias = refs / f"PIECE_SEED_HUMAN{suf}"
        shutil.copy2(dest, alias)
        write_piece_set_source(
            theme_new_dir,
            kind="human",
            paths=[dest],
            extra={
                "note": "human seed installed (gate/chat)",
                "seedOnly": True,
                "activeSeed": dest.name,
                "builtAt": datetime.now(timezone.utc).isoformat(),
            },
        )
    return {"ok": True, "name": dest.name, "path": str(dest), "slot": int(slot)}


def restore_piece_library_from_prev(
    theme_new_dir: Path,
    *,
    also_human: bool = True,
) -> Dict[str, Any]:
    """_prev_shape / 現行 piece_set / FOOD_UNIT_PIECE_SET からライブラリへ復元。"""
    theme_new_dir = Path(theme_new_dir)
    set_dir = theme_new_dir / SET_DIR_REL
    srcs: List[Path] = []
    if set_dir.is_dir():
        prev = sorted(set_dir.glob("_prev_shape_0*.jpg"))
        if len(prev) >= 3:
            srcs = prev[:5]
        if not srcs:
            srcs = sorted(set_dir.glob("piece_set_0*.jpg"))[:5]
    refs = theme_new_dir / "01_refs"
    if len(srcs) < 5 and refs.is_dir():
        for name in (
            "FOOD_UNIT_PIECE_SET_01.jpg",
            "FOOD_UNIT_PIECE_SET_02.jpg",
            "FOOD_UNIT_PIECE_SET_03.jpg",
            "FOOD_UNIT_PIECE_SET_04.jpg",
            "FOOD_UNIT_PIECE_SET_05.jpg",
        ):
            p = refs / name
            if p.is_file() and p not in srcs:
                srcs.append(p)
            if len(srcs) >= 5:
                break
    if not srcs:
        return {"ok": False, "message": "復元元がありません", "n": 0, "names": []}
    lib = library_dir(theme_new_dir)
    lib.mkdir(parents=True, exist_ok=True)
    written: List[str] = []
    for i, src in enumerate(srcs[:5], start=1):
        dest = lib / f"piece_lib_{i:02d}.jpg"
        shutil.copy2(src, dest)
        written.append(dest.name)
        if also_human:
            truth = human_truth_dir(theme_new_dir)
            truth.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, truth / f"piece_truth_{i:02d}.jpg")
    return {"ok": True, "n": len(written), "names": written, "from": [p.name for p in srcs[:5]]}


def save_piece_seed_data_url(theme_new_dir: Path, data_url: str) -> Dict[str, Any]:
    s = (data_url or "").strip()
    m = re.match(r"^data:image/([\w+.-]+);base64,(.+)$", s, re.I | re.S)
    if not m:
        raise ValueError("data URL 形式ではありません")
    kind = str(m.group(1) or "png").lower().replace("jpeg", "jpg")
    if "png" in kind:
        suf = ".png"
    elif "webp" in kind:
        suf = ".webp"
    else:
        suf = ".jpg"
    raw = base64.b64decode(m.group(2))
    # 追記登録（既存種は残し、ラジオで1枚採用）
    return install_human_seed_bytes(theme_new_dir, raw, suffix=suf, replace_all=False, activate=True)


def save_piece_seed_url(theme_new_dir: Path, url: str) -> Dict[str, Any]:
    url = (url or "").strip()
    if not url.startswith("http"):
        raise ValueError("http(s) の画像URLを指定してください")
    with urlopen(url, timeout=60) as resp:  # noqa: S310
        raw = resp.read()
        ctype = str(resp.headers.get("Content-Type") or "").lower()
    suf = ".png" if "png" in ctype else (".webp" if "webp" in ctype else ".jpg")
    return install_human_seed_bytes(theme_new_dir, raw, suffix=suf, replace_all=False, activate=True)


def list_piece_seed_display(
    theme_new_dir: Path,
) -> Dict[str, Any]:
    """一片生成の元見本（種）をゲート表示用に返す。

    優先: PIECE_SEED_HUMAN → piece_set_source.json の paths → resolve_preferred_seed_kind。
    rel は theme_new からの相対（/file/ 用）。無いとき empty。
    """
    theme_new_dir = Path(theme_new_dir)
    kind = ""
    paths: List[Path] = []
    human_seeds = discover_human_seed_paths(theme_new_dir)
    if human_seeds:
        kind = "human"
        paths = human_seeds
    try:
        from b5_piece_truth_rules import (
            SOURCE_NAME,
            resolve_preferred_seed_kind,
        )

        if not paths:
            src_json = theme_new_dir / "food_piece_set5_virt" / SOURCE_NAME
            if src_json.is_file():
                try:
                    obj = json.loads(src_json.read_text(encoding="utf-8"))
                except Exception:
                    obj = {}
                if isinstance(obj, dict):
                    kind = str(obj.get("kind") or "").strip().lower()
                    for raw in obj.get("paths") or []:
                        p = Path(str(raw))
                        if p.is_file():
                            paths.append(p.resolve())
        if not paths:
            kind2, paths2 = resolve_preferred_seed_kind(theme_new_dir)
            kind = kind or kind2
            paths = list(paths2)
    except Exception:
        if not paths:
            paths = []
    root = theme_new_dir.resolve()
    items: List[Dict[str, str]] = []
    for p in paths:
        try:
            rel = str(p.resolve().relative_to(root)).replace("\\", "/")
        except Exception:
            # theme 外はコピーせず名前だけ
            rel = ""
        items.append(
            {
                "name": p.name,
                "rel": rel,
                "path": str(p),
            }
        )
    kind_lab = {
        "human": "人間実物",
        "wrap": "個包装（WRAP）",
        "virt": "仮想フォールバック（CANON等）",
        "none": "なし",
    }.get(kind, kind or "不明")
    return {"ok": True, "kind": kind, "kindLabel": kind_lab, "items": items}


def restore_human_truth_from_piece_set(
    theme_new_dir: Path,
    *,
    max_n: int = MAX_PIECE,
    overwrite: bool = False,
) -> Dict[str, Any]:
    """既存 piece_set をゲート表示用の PIECE_TRUTH_HUMAN へ戻す。

    消えたのではなく HUMAN 未作成のとき用。元の piece_set は残す。
    kind を human に書き換えない（virt/wrap の出自は source のまま）。
    """
    theme_new_dir = Path(theme_new_dir)
    existing = list_human_piece_files(theme_new_dir)
    if existing and not overwrite:
        return {
            "ok": True,
            "skipped": True,
            "reason": "human_exists",
            "n": len(existing),
            "names": [p.name for p in existing],
        }
    srcs = list_piece_set_files(theme_new_dir, max_n=int(max_n or MAX_PIECE))
    if not srcs:
        return {"ok": False, "message": "piece_set がありません", "n": 0, "names": []}
    truth_dir = human_truth_dir(theme_new_dir)
    if overwrite and truth_dir.is_dir():
        for old in list_human_piece_files(theme_new_dir):
            try:
                old.unlink()
            except Exception:
                pass
    truth_dir.mkdir(parents=True, exist_ok=True)
    written: List[Path] = []
    for i, src in enumerate(srcs, start=1):
        dest = truth_dir / f"piece_truth_{i:02d}.jpg"
        shutil.copy2(src, dest)
        written.append(dest)
    # canons の表示フラグだけ立てる（source kind は触らない）
    try:
        from food_unit_canon import load_food_canons

        meta = load_food_canons(theme_new_dir) if theme_new_dir.is_dir() else {}
        if not isinstance(meta, dict):
            meta = {}
        meta["hasPieceSet"] = True
        meta["pieceHas"] = True
        meta["pieceNames"] = [p.name for p in written]
        meta["pieceSetN"] = len(srcs)
        meta["pieceRestoredFromSetAt"] = datetime.now(timezone.utc).isoformat()
        _write_canons_meta(theme_new_dir, meta)
    except Exception:
        pass
    return {
        "ok": True,
        "skipped": False,
        "n": len(written),
        "names": [p.name for p in written],
        "from": [p.name for p in srcs],
    }


def detect_piece_has(theme_new_dir: Path, *, draft: Optional[Dict[str, Any]] = None) -> bool:
    if isinstance(draft, dict) and "pieceHas" in draft:
        return bool(draft.get("pieceHas"))
    if list_human_piece_files(theme_new_dir):
        return True
    if list_piece_set_files(theme_new_dir, max_n=1):
        return True
    try:
        from food_unit_canon import load_food_canons

        meta = load_food_canons(theme_new_dir)
        if meta.get("hasPieceSet") is True:
            return True
        if meta.get("pieceHas") is True:
            return True
    except Exception:
        pass
    return False


def _next_truth_path(theme_new_dir: Path, suffix: str = ".png") -> Path:
    d = human_truth_dir(theme_new_dir)
    d.mkdir(parents=True, exist_ok=True)
    for i in range(1, MAX_PIECE + 1):
        p = d / f"piece_truth_{i:02d}{suffix}"
        if not p.is_file():
            return p
    raise ValueError(f"一片見本は最大{MAX_PIECE}枚です。不要なものを削除してから追加してください。")


def save_piece_bytes(theme_new_dir: Path, raw: bytes, *, suffix: str = ".png") -> Dict[str, Any]:
    if len(list_human_piece_files(theme_new_dir)) >= MAX_PIECE:
        raise ValueError(f"一片見本は最大{MAX_PIECE}枚です")
    dest = _next_truth_path(theme_new_dir, suffix=suffix if suffix.startswith(".") else f".{suffix}")
    dest.write_bytes(raw)
    # ライブラリにも控え（削除しても過去分が残る）
    lib = library_dir(theme_new_dir)
    lib.mkdir(parents=True, exist_ok=True)
    slot = int(dest.stem.split("_")[-1])
    lib_dest = lib / f"piece_lib_{slot:02d}{dest.suffix.lower()}"
    shutil.copy2(dest, lib_dest)
    return {"ok": True, "name": dest.name, "path": str(dest), "slot": slot}


def save_piece_data_url(theme_new_dir: Path, data_url: str) -> Dict[str, Any]:
    s = (data_url or "").strip()
    m = re.match(r"^data:image/([\w+.-]+);base64,(.+)$", s, re.I | re.S)
    if not m:
        raise ValueError("data URL 形式ではありません")
    kind = str(m.group(1) or "png").lower().replace("jpeg", "jpg")
    if "png" in kind:
        suf = ".png"
    elif "webp" in kind:
        suf = ".webp"
    else:
        suf = ".jpg"
    raw = base64.b64decode(m.group(2))
    return save_piece_bytes(theme_new_dir, raw, suffix=suf)


def save_piece_url(theme_new_dir: Path, url: str) -> Dict[str, Any]:
    url = (url or "").strip()
    if not url.startswith("http"):
        raise ValueError("http(s) の画像URLを指定してください")
    with urlopen(url, timeout=60) as resp:  # noqa: S310 — 人間が指定した画像URLのみ
        raw = resp.read()
        ctype = str(resp.headers.get("Content-Type") or "").lower()
    suf = ".png" if "png" in ctype else (".webp" if "webp" in ctype else ".jpg")
    return save_piece_bytes(theme_new_dir, raw, suffix=suf)


def delete_piece_slot(theme_new_dir: Path, slot: int) -> Dict[str, Any]:
    """ソフト削除: HUMAN から外すがライブラリへ退避（過去分は残す）。"""
    d = human_truth_dir(theme_new_dir)
    lib = library_dir(theme_new_dir)
    lib.mkdir(parents=True, exist_ok=True)
    removed: List[str] = []
    archived: List[str] = []
    for ext in (".png", ".jpg", ".jpeg", ".webp"):
        p = d / f"piece_truth_{int(slot):02d}{ext}"
        if not p.is_file():
            continue
        dest = lib / f"piece_lib_{int(slot):02d}{ext}"
        # 既存ライブラリを潰さない
        if dest.is_file():
            stamp = datetime.now().strftime("%H%M%S")
            dest = lib / f"piece_lib_{int(slot):02d}_{stamp}{ext}"
        shutil.copy2(p, dest)
        archived.append(dest.name)
        p.unlink()
        removed.append(p.name)
    return {"ok": True, "removed": removed, "archived": archived, "soft": True}


def purge_piece_name(theme_new_dir: Path, name: str) -> Dict[str, Any]:
    """ダメな一片を完全削除（復元しない）。HUMAN／LIBRARY／同番 piece_set。"""
    theme_new_dir = Path(theme_new_dir)
    name = str(name or "").strip()
    if not name or "/" in name or "\\" in name or ".." in name:
        raise ValueError("削除名が不正です")
    removed: List[str] = []
    bases = (
        human_truth_dir(theme_new_dir),
        library_dir(theme_new_dir),
        theme_new_dir / SET_DIR_REL,
    )
    for base in bases:
        if not base.is_dir():
            continue
        p = base / name
        if p.is_file():
            p.unlink()
            removed.append(str(p.relative_to(theme_new_dir)).replace("\\", "/"))
            continue
        # stem 一致（拡張子違い）
        stem = Path(name).stem
        for hit in base.glob(stem + ".*"):
            if hit.is_file():
                hit.unlink()
                removed.append(str(hit.relative_to(theme_new_dir)).replace("\\", "/"))
    # HUMAN 削除時は同スロットの piece_set も落とす
    m = re.match(r"^piece_(?:truth|lib)_(\d+)$", Path(name).stem, re.I)
    if m:
        slot = int(m.group(1))
        set_dir = theme_new_dir / SET_DIR_REL
        if set_dir.is_dir():
            for hit in set_dir.glob(f"piece_set_{slot:02d}.*"):
                if hit.is_file():
                    hit.unlink()
                    removed.append(str(hit.relative_to(theme_new_dir)).replace("\\", "/"))
    if not removed:
        raise ValueError(f"削除対象が見つかりません: {name}")
    return {"ok": True, "removed": removed, "hard": True}


def purge_piece_archives(theme_new_dir: Path) -> Dict[str, Any]:
    """作り直し退避（piece_lib_arch_*）をまとめて完全削除。"""
    lib = library_dir(theme_new_dir)
    removed: List[str] = []
    if lib.is_dir():
        for p in sorted(lib.glob("piece_lib_arch_*.*")):
            if p.is_file():
                p.unlink()
                removed.append(p.name)
    return {"ok": True, "n": len(removed), "removed": removed}


def _archive_piece_set(theme_new_dir: Path) -> None:
    """なし選択時: 現行 piece_set を piece_set_off へ退避（復元可）。"""
    set_dir = Path(theme_new_dir) / SET_DIR_REL
    if not set_dir.is_dir():
        return
    files = list(set_dir.glob("piece_set_0*"))
    if not files:
        return
    off = Path(theme_new_dir) / SET_OFF_REL
    off.mkdir(parents=True, exist_ok=True)
    for p in files:
        dest = off / p.name
        if dest.is_file():
            dest.unlink()
        shutil.move(str(p), str(dest))


def resolve_source_image(theme_new_dir: Path, name: str) -> Optional[Path]:
    """候補名・手動名・一片名・ライブラリ名から実ファイルを探す。"""
    name = str(name or "").strip()
    if not name:
        return None
    refs = Path(theme_new_dir) / "01_refs"
    for base in (
        refs,
        human_truth_dir(theme_new_dir),
        library_dir(theme_new_dir),
        Path(theme_new_dir) / SET_DIR_REL,
        Path(theme_new_dir) / "manual",
        Path(theme_new_dir),
    ):
        if not base.is_dir():
            continue
        p = base / name
        if p.is_file():
            return p
    # piece_truth_0N / piece_lib_0N → piece_set_0N
    stem = Path(name).stem
    m = re.match(r"^piece_(?:truth|lib)_(\d+)$", stem, re.I)
    if m:
        idx = int(m.group(1))
        for cand in (
            human_truth_dir(theme_new_dir) / f"piece_truth_{idx:02d}.jpg",
            library_dir(theme_new_dir) / f"piece_lib_{idx:02d}.jpg",
            Path(theme_new_dir) / SET_DIR_REL / f"piece_set_{idx:02d}.jpg",
            Path(theme_new_dir) / SET_DIR_REL / f"_prev_shape_{idx:02d}.jpg",
        ):
            if cand.is_file():
                return cand
    # stem match
    for base in (
        refs,
        human_truth_dir(theme_new_dir),
        library_dir(theme_new_dir),
        Path(theme_new_dir) / SET_DIR_REL,
    ):
        if not base.is_dir():
            continue
        for p in base.glob(stem + ".*"):
            if p.is_file():
                return p
    return None


def apply_piece_selection(
    theme_new_dir: Path,
    *,
    piece_has: bool,
    piece_names: Sequence[str],
) -> Dict[str, Any]:
    """確定時: レ点の画像を人間正本＋piece_set に同期。なしなら hasPieceSet=false のみ。"""
    theme_new_dir = Path(theme_new_dir)
    from food_unit_canon import load_food_canons

    meta = load_food_canons(theme_new_dir) if theme_new_dir.is_dir() else {}
    if not isinstance(meta, dict):
        meta = {}

    if not piece_has:
        _archive_piece_set(theme_new_dir)
        meta["hasPieceSet"] = False
        meta["pieceHas"] = False
        meta["pieceNames"] = []
        meta["pieceSetN"] = 0
        _write_canons_meta(theme_new_dir, meta)
        return {"ok": True, "hasPieceSet": False, "n": 0, "names": []}

    names = [str(x).strip() for x in (piece_names or []) if str(x).strip()]
    # 名前が空なら既存人間正本を使う
    src_paths: List[Path] = []
    if names:
        for n in names[:MAX_PIECE]:
            p = resolve_source_image(theme_new_dir, n)
            if not p:
                raise ValueError(f"一片見本画像が見つかりません: {n}")
            src_paths.append(p)
    else:
        src_paths = list_human_piece_files(theme_new_dir)[:MAX_PIECE]
        if not src_paths:
            src_paths = list_piece_set_files(theme_new_dir, max_n=MAX_PIECE)
    if not src_paths:
        raise ValueError(
            "一片見本ありなのに画像がありません。貼付するか候補にレ点を付けてください。"
        )

    # 先に中身をメモリへ（同一 HUMAN 内の選び直しで削除→copy が WinError 2 になるのを防ぐ）
    payloads: List[Tuple[str, bytes]] = []
    for src in src_paths:
        ext = src.suffix.lower() if src.suffix else ".jpg"
        if ext not in (".png", ".jpg", ".jpeg", ".webp"):
            ext = ".jpg"
        payloads.append((ext, Path(src).read_bytes()))

    truth_dir = human_truth_dir(theme_new_dir)
    truth_dir.mkdir(parents=True, exist_ok=True)
    for old in list_human_piece_files(theme_new_dir):
        try:
            old.unlink()
        except Exception:
            pass
    written_truth: List[Path] = []
    for i, (ext, raw) in enumerate(payloads, start=1):
        dest = truth_dir / f"piece_truth_{i:02d}{ext}"
        dest.write_bytes(raw)
        written_truth.append(dest)

    # piece_set へ jpg 同期（compose が読む正）
    set_dir = theme_new_dir / SET_DIR_REL
    set_dir.mkdir(parents=True, exist_ok=True)
    for old in set_dir.glob("piece_set_0*"):
        try:
            old.unlink()
        except Exception:
            pass
    written_set: List[Path] = []
    try:
        from PIL import Image
    except Exception:
        Image = None  # type: ignore
    for i, src in enumerate(written_truth, start=1):
        dest = set_dir / f"piece_set_{i:02d}.jpg"
        if Image is not None:
            im = Image.open(src).convert("RGB")
            im.save(dest, quality=92)
        else:
            shutil.copy2(src, dest)
        written_set.append(dest)

    write_piece_set_source(
        theme_new_dir,
        kind="human",
        paths=written_truth,
        extra={
            "n": len(written_truth),
            "pieceNames": [p.name for p in written_truth],
            "builtAt": datetime.now(timezone.utc).isoformat(),
            "fromGate": True,
        },
    )

    meta["hasPieceSet"] = True
    meta["pieceHas"] = True
    meta["pieceNames"] = [p.name for p in written_truth]
    meta["pieceSetN"] = len(written_set)
    meta["pieceConfirmedAt"] = datetime.now(timezone.utc).isoformat()
    _write_canons_meta(theme_new_dir, meta)

    return {
        "ok": True,
        "hasPieceSet": True,
        "n": len(written_set),
        "names": [p.name for p in written_truth],
        "setNames": [p.name for p in written_set],
    }


def _write_canons_meta(theme_new_dir: Path, meta: Dict[str, Any]) -> None:
    """food_canons.json をマージ更新（無いキーは壊さない）。"""
    from food_unit_canon import FOOD_CANONS_NAME, load_food_canons

    root = Path(theme_new_dir)
    cur = load_food_canons(root) if root.is_dir() else {}
    if not isinstance(cur, dict):
        cur = {}
    cur.update(meta)
    dest = root / FOOD_CANONS_NAME
    # 既存が 01_refs 側ならそこへ
    alt = root / "01_refs" / FOOD_CANONS_NAME
    if alt.is_file() and not dest.is_file():
        dest = alt
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")


def remake_one_piece(
    theme_new_dir: Path,
    *,
    name: str,
    note: str = "",
    seed_name: str = "",
    quality: str = "low",
    lock_current: bool = True,
) -> Dict[str, Any]:
    """一片カード1枚だけ作り直す（他枚は触らない）。

    lock_current=True: 現行写真を1枚目で色・外形ロック。要望は差分のみ。
    lock_current=False: 種ベースで作り直し（色・形の変更も可）。
    種は seed_name または activeSeed / PIECE_SEED_HUMAN。
    """
    from openai_image import generate_with_references, resolve_openai_image_model
    from b5_piece_truth_rules import (
        build_shape_policy_prompt_block,
        infer_shape_family_from_texts,
        resolve_preferred_seed_kind,
        resolve_shape_donor_paths,
        shape_donor_attach_prompt_block,
    )

    theme_new_dir = Path(theme_new_dir)
    name = str(name or "").strip()
    if not name:
        raise ValueError("対象の一片名が空です")
    note = " ".join(str(note or "").split())[:400]
    lock_current = bool(lock_current)

    # 種
    if seed_name:
        activate_piece_seed(theme_new_dir, seed_name)
    kind, seeds = resolve_preferred_seed_kind(theme_new_dir)
    if kind == "none" or not seeds:
        raise ValueError("元見本（種）がありません。②で種を選んでから作り直してください。")

    # 置き先スロット（piece_truth_03 / piece_lib_prev_02 → 03/02）
    m = re.search(r"(\d+)$", Path(name).stem)
    slot = int(m.group(1)) if m else 0
    if slot < 1 or slot > MAX_PIECE:
        # ライブラリ変則名 → HUMAN の空きへ
        try:
            dest_truth = _next_truth_path(theme_new_dir, ".jpg")
            slot = int(dest_truth.stem.split("_")[-1])
        except Exception as ex:
            raise ValueError(f"スロットを決められません: {name}") from ex
    else:
        dest_truth = human_truth_dir(theme_new_dir) / f"piece_truth_{slot:02d}.jpg"

    # 旧ファイルをライブラリへ退避（ロック時は退避コピーを1枚目参照に使う）
    old = resolve_source_image(theme_new_dir, name)
    lib = library_dir(theme_new_dir)
    lib.mkdir(parents=True, exist_ok=True)
    archived = ""
    lock_ref: Optional[Path] = None
    if old and old.is_file():
        stamp = datetime.now().strftime("%H%M%S")
        archived_path = lib / f"piece_lib_arch_{slot:02d}_{stamp}{old.suffix.lower() or '.jpg'}"
        shutil.copy2(old, archived_path)
        archived = archived_path.name
        lock_ref = archived_path
    if lock_current and (lock_ref is None or not lock_ref.is_file()):
        raise ValueError(
            "現行写真ロックには、このカードの現行画像が必要です。"
            "ロックを外すか、先に一片画像を置いてください。"
        )

    family = infer_shape_family_from_texts([note, kind] + [p.name for p in seeds])
    policy = build_shape_policy_prompt_block(family)
    pose_i = min(max(slot - 1, 0), 4)
    pose = (
        "向き01: やや斜め上。厚みと表面の質感がよく見える。",
        "向き02: 真上寄り。上面と境目の様子が見える。",
        "向き03: 別角度の斜め。同じ1片の見た目を保て。",
        "向き04: やや低角度。厚みの断面感が見える。",
        "向き05: 別の斜め。シルエットを別物に変えるな。",
    )[pose_i]

    image_paths: List[Path] = []
    if lock_current and lock_ref is not None:
        image_paths.append(lock_ref)
        # 種は参考（1枚目ロックと同一バイトなら二重添付しない）
        for sp in seeds:
            try:
                if sp.resolve() == lock_ref.resolve():
                    continue
                if sp.is_file() and sp.read_bytes() == lock_ref.read_bytes():
                    continue
            except Exception:
                pass
            image_paths.append(sp)
            break
        prompt_lines = [
            "【個品・一片見本・現行写真ロック・1枚作り直し】",
            "添付1枚目＝このカードの現行写真。色・外形・向き・大きさ・血合い・厚みの正。",
            "全面描き直し禁止。別物の削り節・別シルエットに寄せるな。",
            "【差分のみ】人間要望があるときだけ、指定された点（縁のほつれ等）を最小変更。",
            "要望に無い色・形・向きの変更は禁止。",
            "【必須】画面に出す個品はちょうど1片だけ。複数片・山盛り・ばら撒き禁止。",
            "写真の四角コピペ貼付禁止。実写として描け。",
            "手・指・皿・袋・個包装フィルム・文字は描くな。背景は白〜生成りのみ。",
            "向き・構図は1枚目と同じ。角度を変えるな。",
            policy.strip(),
        ]
        if len(image_paths) > 1:
            prompt_lines.insert(
                2,
                f"添付2枚目以降は種（優先={kind}）の参考のみ。1枚目と食い違うな。",
            )
    else:
        image_paths = list(seeds)
        prompt_lines = [
            "【個品・一片見本・実物由来・1枚作り直し】",
            f"添付の種写真（優先={kind}）が開封後1片の正。色・縁・質感・外形系統を合わせろ。",
            "【必須】画面に出す個品はちょうど1片だけ。複数片・山盛り・ばら撒き禁止。",
            "写真の四角コピペ貼付禁止。実写として描け。",
            "手・指・皿・袋・個包装フィルム・文字は描くな。背景は白〜生成りのみ。",
            policy.strip(),
            pose,
        ]
    if note:
        if lock_current:
            prompt_lines.append(f"【人間要望・差分のみ】{note}")
        else:
            prompt_lines.append(f"【人間要望・この1枚】{note}")
    elif lock_current:
        prompt_lines.append(
            "【人間要望】なし。現行写真とほぼ同一でよい（不要な変更をするな）。"
        )

    # T4: 同ファミリー形状ドナーを自動添付（色は触らず形だけ借りる）
    shape_donor_n = 0
    try:
        from b5_piece_deepen import surface_shape_split_enabled

        if surface_shape_split_enabled():
            dmeta = resolve_shape_donor_paths(theme_new_dir, max_n=3)
            donors = [Path(p) for p in (dmeta.get("paths") or []) if Path(p).is_file()]
            if donors:
                image_paths.extend(donors)
                shape_donor_n = len(donors)
                sb = shape_donor_attach_prompt_block(max_n=shape_donor_n)
                if sb:
                    prompt_lines.append(sb.strip())
    except Exception:
        shape_donor_n = 0

    prompt = "\n".join(prompt_lines)

    model = resolve_openai_image_model(None)
    raw, used = generate_with_references(
        prompt=prompt,
        image_paths=image_paths,
        model=model,
        size="1024x1024",
        quality=str(quality or "low"),
    )

    # piece_set の同番号も更新（compose が読む正）
    set_dir = theme_new_dir / SET_DIR_REL
    set_dir.mkdir(parents=True, exist_ok=True)
    dest_set = set_dir / f"piece_set_{slot:02d}.jpg"
    dest_set.write_bytes(raw)
    dest_set.with_suffix(".txt").write_text(prompt, encoding="utf-8")

    human_truth_dir(theme_new_dir).mkdir(parents=True, exist_ok=True)
    # jpg で統一
    try:
        from PIL import Image
        import io

        im = Image.open(io.BytesIO(raw)).convert("RGB")
        im.save(dest_truth, quality=92)
    except Exception:
        dest_truth.write_bytes(raw)

    lib_dest = lib / f"piece_lib_{slot:02d}.jpg"
    shutil.copy2(dest_truth, lib_dest)

    mode_lab = "現行ロック" if lock_current else "種ベース"
    return {
        "ok": True,
        "name": dest_truth.name,
        "slot": slot,
        "path": str(dest_truth),
        "setName": dest_set.name,
        "seed": seeds[0].name,
        "model": used,
        "archived": archived,
        "note": note,
        "lockCurrent": lock_current,
        "refCount": len(image_paths),
        "shapeDonorN": shape_donor_n,
        "message": (
            f"1枚だけ作り直しました（{dest_truth.name}／{mode_lab}）。"
            "他の枚は変えていません。"
            + (f" 形状ドナー{shape_donor_n}枚同送。" if shape_donor_n else "")
        ),
    }
