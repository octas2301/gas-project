# -*- coding: utf-8 -*-
"""
B-④ ①′ 個品見本ゲート — 人間確定状態を Drive JSON に書き、GAS が読めるようにする。

ローカル正: tools/set_main_image/_out/theme_new_run_{JAN}/food_canons.json
Drive 正:  05.画像生成（セットMAIN）/00.テスト出力/_b5_food_canon_status.json

例:
  python b5_food_canon_status.py --jan 4906283045119 --export-drive
  python b5_food_canon_status.py --export-drive --all-known
  python b5_food_canon_status.py --jan 4906283045119
"""
from __future__ import annotations

import argparse
import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from food_unit_canon import (
    FOOD_CANONS_NAME,
    existing_wrap_path,
    load_food_canons,
    theme_new_dir_for_jan,
)
from work_paths import DEFAULT_WORK_ROOT

LOG = logging.getLogger("set_main_image.b5_food_canon_status")

SCRIPT_DIR = Path(__file__).resolve().parent
STATUS_NAME = "_b5_food_canon_status.json"


def drive_status_path(work_root: Optional[Path] = None) -> Path:
    root = Path(work_root) if work_root is not None else DEFAULT_WORK_ROOT
    return root / "00.テスト出力" / STATUS_NAME


def _read_json_obj(path: Path) -> Dict[str, Any]:
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _binding_from_theme_new(
    tn: Path,
    *,
    parent_sku: str = "",
    name_base: str = "",
) -> Tuple[str, str]:
    """親SKU／商品名ベース。food_canons → messages → gate_session → 引数。"""
    meta = load_food_canons(tn) if tn.is_dir() else {}
    msg = _read_json_obj(tn / "messages.json") if tn.is_dir() else {}
    gate = _read_json_obj(tn / "gate_session.json") if tn.is_dir() else {}
    sku = (
        str(parent_sku or "").strip()
        or str(meta.get("parentSku") or "").strip()
        or str(msg.get("parentSku") or "").strip()
        or str(gate.get("parentSku") or "").strip()
    )
    nb = (
        str(name_base or "").strip()
        or str(meta.get("nameBase") or "").strip()
        or str(msg.get("nameBase") or "").strip()
        or str(msg.get("productName") or "").strip()
        or str(gate.get("nameBase") or "").strip()
        or str(gate.get("productName") or "").strip()
    )
    return sku, nb


def _row_has_wrap(row: Optional[Dict[str, Any]]) -> bool:
    if not isinstance(row, dict):
        return False
    if row.get("hasWrap") is True:
        return True
    if row.get("hasWrap") is False:
        return False
    wn = str(row.get("wrapName") or row.get("wrapPath") or "").strip()
    return bool(wn) and wn not in ("__none__", "NONE", "none")


def resolve_has_wrap_from_status_doc(
    doc: Dict[str, Any],
    *,
    parent_sku: str = "",
    name_base: str = "",
    jan: str = "",
) -> bool:
    """判定順: 親SKU → 商品名ベース → JAN。欠けると片方だけになるのを防ぐ。"""
    by_sku = (doc or {}).get("byParentSku") or {}
    by_nb = (doc or {}).get("byNameBase") or {}
    by_jan = (doc or {}).get("byJan") or {}
    sku = str(parent_sku or "").strip()
    nb = str(name_base or "").strip()
    j = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
    row: Optional[Dict[str, Any]] = None
    if sku and isinstance(by_sku.get(sku), dict):
        row = by_sku[sku]
    elif nb and isinstance(by_nb.get(nb), dict):
        row = by_nb[nb]
    elif j and isinstance(by_jan.get(j), dict):
        row = by_jan[j]
    return _row_has_wrap(row)


def status_for_jan(
    jan: str,
    *,
    script_dir: Optional[Path] = None,
    parent_sku: str = "",
    name_base: str = "",
) -> Dict[str, Any]:
    """1 JAN の food_canons.json を読み、GAS 向けサマリを返す。hasWrap は常に bool。"""
    jan_d = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
    tn = theme_new_dir_for_jan(jan_d, script_dir=script_dir)
    meta = load_food_canons(tn) if tn.is_dir() else {}
    confirmed = str(meta.get("confirmedBy") or "")
    package_only = bool(meta.get("packageOnly"))
    identity_status = str(meta.get("identityStatus") or "")
    ok = confirmed == "human"
    label = "未確定"
    if ok and package_only:
        label = "確定（正本のみ）"
    elif ok:
        label = "確定（個品あり）"
    elif confirmed == "machine":
        label = "機械のみ（要人間確定）"
    sku, nb = _binding_from_theme_new(
        tn, parent_sku=parent_sku, name_base=name_base
    )
    wrap_p = existing_wrap_path(tn) if tn.is_dir() else None
    wrap_name_meta = str(meta.get("wrapName") or "").strip()
    wrap_status = str(meta.get("wrapStatus") or "").strip().lower()
    has_wrap_file = bool(wrap_p and Path(wrap_p).is_file())
    has_wrap_meta = wrap_status == "present" or (
        bool(wrap_name_meta) and wrap_name_meta not in ("__none__", "NONE", "none")
    )
    has_wrap = bool(has_wrap_file or has_wrap_meta) and not package_only
    wrap_name_out = ""
    if has_wrap:
        if wrap_p and Path(wrap_p).is_file():
            wrap_name_out = Path(wrap_p).name
        elif wrap_name_meta:
            wrap_name_out = wrap_name_meta
    hand_hold_name = ""
    has_hand_hold = False
    if tn.is_dir() and not package_only:
        try:
            from b4_wrap_hand_hold import selected_path as _wrap_hand_selected

            hh = _wrap_hand_selected(tn)
            if hh is not None and Path(hh).is_file():
                has_hand_hold = True
                hand_hold_name = Path(hh).name
        except Exception:
            has_hand_hold = False
    has_piece_set = False
    piece_names_out: List[str] = []
    piece_n = 0
    if tn.is_dir() and not package_only:
        if meta.get("hasPieceSet") is True or meta.get("pieceHas") is True:
            has_piece_set = True
        try:
            from b5_piece_gate import list_human_piece_files, list_piece_set_files

            human_n = list_human_piece_files(tn)
            set_n = list_piece_set_files(tn, max_n=5)
            if human_n or set_n:
                if meta.get("hasPieceSet") is not False:
                    has_piece_set = True
            piece_names_out = [p.name for p in (human_n or set_n)]
            piece_n = len(set_n) or len(human_n)
        except Exception:
            piece_names_out = [str(x) for x in (meta.get("pieceNames") or []) if str(x).strip()]
            piece_n = int(meta.get("pieceSetN") or len(piece_names_out) or 0)
            if meta.get("hasPieceSet") is True:
                has_piece_set = True
        if meta.get("hasPieceSet") is False:
            has_piece_set = False
            piece_n = 0
    return {
        "jan": jan_d,
        "parentSku": sku,
        "nameBase": nb,
        "themeNewDir": str(tn) if tn.is_dir() else "",
        "confirmedBy": confirmed,
        "ok": ok,
        "label": label,
        "packageOnly": package_only,
        "hasWrap": bool(has_wrap),
        "wrapName": wrap_name_out,
        "hasWrapHandHold": bool(has_hand_hold),
        "wrapHandHoldName": hand_hold_name,
        "hasPieceSet": bool(has_piece_set),
        "pieceNames": list(piece_names_out),
        "pieceSetN": int(piece_n or 0),
        "identityStatus": identity_status,
        "identityName": meta.get("identityName") or meta.get("primaryName"),
        "sizzleName": meta.get("sizzleName"),
        "secondaryNames": list(meta.get("secondaryNames") or []),
        "builtAt": meta.get("builtAt"),
        "checkedAt": datetime.now(timezone.utc).isoformat(),
    }


def load_drive_status(work_root: Optional[Path] = None) -> Dict[str, Any]:
    empty = {
        "schema": "b5_food_canon_status_v1",
        "updatedAt": "",
        "byJan": {},
        "byParentSku": {},
        "byNameBase": {},
    }
    p = drive_status_path(work_root)
    if not p.is_file():
        return dict(empty)
    try:
        data = json.loads(p.read_text(encoding="utf-8-sig"))
        if isinstance(data, dict) and isinstance(data.get("byJan"), dict):
            data.setdefault("byParentSku", {})
            data.setdefault("byNameBase", {})
            return data
    except Exception as e:
        LOG.warning("load_drive_status fail %s", e)
    return dict(empty)


def export_status_to_drive(
    jans: Sequence[str],
    *,
    work_root: Optional[Path] = None,
    script_dir: Optional[Path] = None,
    merge: bool = True,
    binding_by_jan: Optional[Dict[str, Dict[str, str]]] = None,
) -> Path:
    """JAN 群の状態を Drive JSON にマージ書込。全件に hasWrap(bool) を必ず書く。"""
    dest = drive_status_path(work_root)
    dest.parent.mkdir(parents=True, exist_ok=True)
    doc = load_drive_status(work_root) if merge else {
        "schema": "b5_food_canon_status_v1",
        "byJan": {},
        "byParentSku": {},
        "byNameBase": {},
    }
    by_jan: Dict[str, Any] = dict(doc.get("byJan") or {})
    by_sku: Dict[str, Any] = dict(doc.get("byParentSku") or {})
    by_nb: Dict[str, Any] = dict(doc.get("byNameBase") or {})
    hints = binding_by_jan or {}
    for jan in jans:
        j = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
        if not j:
            continue
        hint = hints.get(j) or hints.get(str(jan)) or {}
        rec = status_for_jan(
            j,
            script_dir=script_dir,
            parent_sku=str(hint.get("parentSku") or ""),
            name_base=str(hint.get("nameBase") or ""),
        )
        # 欠けるとテーマ+1が片方だけになるため、必ず bool で書く
        rec["hasWrap"] = bool(rec.get("hasWrap"))
        rec["hasWrapHandHold"] = bool(rec.get("hasWrapHandHold"))
        rec["hasPieceSet"] = bool(rec.get("hasPieceSet"))
        by_jan[j] = rec
        sku = str(rec.get("parentSku") or "").strip()
        if sku:
            by_sku[sku] = rec
        nb = str(rec.get("nameBase") or "").strip()
        if nb:
            by_nb[nb] = rec
    doc["schema"] = "b5_food_canon_status_v1"
    doc["updatedAt"] = datetime.now(timezone.utc).isoformat()
    doc["byJan"] = by_jan
    doc["byParentSku"] = by_sku
    doc["byNameBase"] = by_nb
    dest.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info("exported %s jans=%s -> %s", len(jans), list(jans), dest)
    return dest


def list_known_jans(script_dir: Optional[Path] = None) -> List[str]:
    """メンテ専用。B-④ は --jan 指定のみ（全フォルダなめ禁止）。"""
    root = Path(script_dir) if script_dir is not None else SCRIPT_DIR
    out_dir = root / "_out"
    if not out_dir.is_dir():
        return []
    jans: List[str] = []
    for p in sorted(out_dir.glob("theme_new_run_*")):
        if not p.is_dir():
            continue
        suffix = p.name.replace("theme_new_run_", "", 1)
        if suffix and (p / FOOD_CANONS_NAME).is_file():
            jans.append(suffix)
    if len(jans) >= 80:
        LOG.warning("theme_new_run_* = %s （--all-known は遅い。B-④は --jan だけ使う）", len(jans))
    return jans


def export_after_human_confirm(
    jan: str,
    *,
    work_root: Optional[Path] = None,
    parent_sku: str = "",
    name_base: str = "",
) -> None:
    """ゲート確定直後に呼ぶ（1 JAN だけマージ更新）。親SKU／商品名ベースも書く。"""
    try:
        j = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
        hints: Dict[str, Dict[str, str]] = {}
        if j and (parent_sku or name_base):
            hints[j] = {
                "parentSku": str(parent_sku or "").strip(),
                "nameBase": str(name_base or "").strip(),
            }
        export_status_to_drive(
            [jan], work_root=work_root, merge=True, binding_by_jan=hints or None
        )
    except Exception as e:
        LOG.warning("export_after_human_confirm jan=%s %s", jan, e)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="B-④ 個品見本確定状態（Drive JSON）")
    ap.add_argument("--jan", action="append", default=[])
    ap.add_argument("--export-drive", action="store_true")
    ap.add_argument("--all-known", action="store_true", help="メンテ専用。_out/theme_new_run_* を全部（B-④禁止）")
    ap.add_argument("--work-root", type=Path, default=DEFAULT_WORK_ROOT)
    ap.add_argument("--no-merge", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    jans = [str(j).strip() for j in (args.jan or []) if str(j).strip()]
    if args.all_known:
        jans = list_known_jans()
        if len(jans) >= 80:
            LOG.warning("all-known n=%s は遅い。常用は --jan", len(jans))
    if not jans:
        raise SystemExit("--jan または --all-known が必要です")
    if args.export_drive:
        p = export_status_to_drive(
            jans, work_root=args.work_root, merge=not args.no_merge
        )
        print(str(p.resolve()))
        return 0
    for j in jans:
        print(json.dumps(status_for_jan(j), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
