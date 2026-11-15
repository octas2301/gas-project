# -*- coding: utf-8 -*-
"""MAIN／サブ compose が読む個包装見本（FOOD_UNIT_WRAP）の正。

方針:
- 寄せ集め（販促コラージュ）や小さすぎる見本は compose に使わない
- ゲートで人が選んだ「1個画」（wrap_single）を FOOD_UNIT_WRAP にする
- 元の貼付見本は FOOD_UNIT_WRAP_SEED に残し、手持ち／単体の作り直し種にする
切戻し: 呼び出し側で使わなければ旧どおり wrap カードのまま
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image

WRAP_STEM = "FOOD_UNIT_WRAP"
WRAP_SEED_STEM = "FOOD_UNIT_WRAP_SEED"
MIN_COMPOSE_EDGE = 512


def _refs(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / "01_refs"


def _find_stem(refs: Path, stem: str) -> Optional[Path]:
    for ext in (".jpg", ".jpeg", ".png", ".webp", ".JPG", ".PNG"):
        p = refs / f"{stem}{ext}"
        if p.is_file():
            return p
    return None


def existing_compose_wrap_path(theme_new_dir: Path) -> Optional[Path]:
    return _find_stem(_refs(theme_new_dir), WRAP_STEM)


def existing_wrap_seed_path(theme_new_dir: Path) -> Optional[Path]:
    return _find_stem(_refs(theme_new_dir), WRAP_SEED_STEM)


def resolve_wrap_generation_ref(theme_new_dir: Path) -> Optional[Path]:
    """手持ち／単体の生成・作り直しの種。SEED優先、無ければ WRAP。"""
    seed = existing_wrap_seed_path(theme_new_dir)
    if seed is not None:
        return seed
    return existing_compose_wrap_path(theme_new_dir)


def assess_wrap_compose_fitness(path: Path) -> Dict[str, Any]:
    """
    compose 用として足りるか。
    - low_res: 短辺 < 512
    - collage_suspect: 単体ファイル名以外で、販促寄せの疑い（複数被写体っぽい巨大余白は見ない。運用で単体必須にする）
    """
    path = Path(path)
    reasons: List[str] = []
    w = h = 0
    if not path.is_file():
        return {
            "ok": False,
            "reasons": ["missing"],
            "width": 0,
            "height": 0,
            "minEdge": 0,
        }
    try:
        with Image.open(path) as im:
            w, h = im.size
    except Exception:
        return {
            "ok": False,
            "reasons": ["unreadable"],
            "width": 0,
            "height": 0,
            "minEdge": 0,
        }
    edge = min(int(w), int(h))
    if edge < MIN_COMPOSE_EDGE:
        reasons.append("low_res")
    name = path.name
    stem = Path(name).stem
    # 検収単体 or 昇格後の FOOD_UNIT_WRAP は1個画。SEED・貼付名は断定しない
    checked = stem.startswith("wrap_single_") or stem == WRAP_STEM
    if not checked:
        reasons.append("not_checked_single")
    # not_checked_single 単独では落とさない（単体が無い旧運用）。必須化は resolve 側。
    ok = "low_res" not in reasons
    return {
        "ok": ok,
        "reasons": reasons,
        "width": int(w),
        "height": int(h),
        "minEdge": edge,
    }


def _resolve_src(theme_new_dir: Path, name: str) -> Optional[Path]:
    name = str(name or "").strip()
    if not name:
        return None
    refs = _refs(theme_new_dir)
    for p in (
        refs / name,
        Path(theme_new_dir) / "wrap_hand_hold" / name,
        Path(theme_new_dir) / name,
    ):
        if p.is_file():
            return p
    return None


def resolve_compose_wrap_choice(
    theme_new_dir: Path,
    *,
    wrap_name: str = "",
    wrap_single_name: str = "",
    wrap_has: bool = False,
    singles_available: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    確定時: compose に書くファイル名と種を決める。
    singles があるのに単体未選択 → NG（寄せ集めを捨てる）。
    """
    theme_new_dir = Path(theme_new_dir)
    if not wrap_has:
        return {
            "ok": True,
            "composeName": "",
            "seedName": "",
            "message": "",
            "assess": {},
        }

    wrap_name = str(wrap_name or "").strip()
    wrap_single_name = str(wrap_single_name or "").strip()
    if singles_available is None:
        try:
            from b4_wrap_hand_hold import list_single_files

            singles_available = [p.name for p in list_single_files(theme_new_dir)]
        except Exception:
            singles_available = []

    singles_available = [str(x) for x in (singles_available or []) if str(x).strip()]

    if singles_available and not wrap_single_name:
        if wrap_name.startswith("wrap_single_") and wrap_name in singles_available:
            wrap_single_name = wrap_name
        else:
            return {
                "ok": False,
                "composeName": "",
                "seedName": wrap_name,
                "message": (
                    "単体タイプで採用を選んでください。"
                    "MAINの個包装見本は、寄せ集めや小さすぎる写真ではなく、"
                    "検収した1個画（単体）を使います。"
                ),
                "assess": {},
            }

    compose_name = wrap_single_name or wrap_name
    if not compose_name:
        return {
            "ok": False,
            "composeName": "",
            "seedName": "",
            "message": "個包装ありのとき、見本または単体の採用が必要です。",
            "assess": {},
        }

    src = _resolve_src(theme_new_dir, compose_name)
    if src is None:
        return {
            "ok": False,
            "composeName": compose_name,
            "seedName": wrap_name,
            "message": f"個包装の1個画が見つかりません: {compose_name}",
            "assess": {},
        }

    assess = assess_wrap_compose_fitness(src)
    if not assess.get("ok"):
        why = ",".join(assess.get("reasons") or [])
        return {
            "ok": False,
            "composeName": compose_name,
            "seedName": wrap_name,
            "message": (
                f"個包装見本がMAIN用として不足です（{why}）。"
                "単体タイプの大きい1個画を選ぶか、はっきり写った写真を登録してください。"
            ),
            "assess": assess,
        }

    seed_name = wrap_name if (wrap_name and wrap_name != compose_name) else ""
    return {
        "ok": True,
        "composeName": compose_name,
        "seedName": seed_name,
        "message": "",
        "assess": assess,
        "composeSrc": str(src),
    }


def materialize_wrap_seed(
    theme_new_dir: Path,
    seed_name: str = "",
    *,
    also_bootstrap_wrap: bool = True,
) -> Optional[Path]:
    """種を FOOD_UNIT_WRAP_SEED に書く。WRAPが無ければ初回だけ同じ内容で bootstrap。"""
    theme_new_dir = Path(theme_new_dir)
    refs = _refs(theme_new_dir)
    refs.mkdir(parents=True, exist_ok=True)

    src: Optional[Path] = None
    if seed_name:
        src = _resolve_src(theme_new_dir, seed_name)
    if src is None:
        src = existing_compose_wrap_path(theme_new_dir)
    if src is None:
        return None

    ext = src.suffix.lower() or ".jpg"
    # 旧SEED掃除
    for p in refs.glob(f"{WRAP_SEED_STEM}.*"):
        if p.resolve() != src.resolve():
            p.unlink(missing_ok=True)
    dest = refs / f"{WRAP_SEED_STEM}{ext}"
    if src.resolve() != dest.resolve():
        shutil.copy2(src, dest)

    if also_bootstrap_wrap and existing_compose_wrap_path(theme_new_dir) is None:
        wdest = refs / f"{WRAP_STEM}{ext}"
        shutil.copy2(src, wdest)
    return dest


def promote_checked_single_to_compose_wrap(
    theme_new_dir: Path,
    compose_name: str,
    *,
    seed_name: str = "",
) -> Dict[str, Any]:
    """
    検収1個画 → FOOD_UNIT_WRAP。
    元見本があれば FOOD_UNIT_WRAP_SEED へ退避（既にあれば維持優先で seed_name から更新）。
    """
    theme_new_dir = Path(theme_new_dir)
    refs = _refs(theme_new_dir)
    refs.mkdir(parents=True, exist_ok=True)

    # いまの WRAP を種へ（compose 上書き前）
    cur = existing_compose_wrap_path(theme_new_dir)
    if seed_name:
        materialize_wrap_seed(theme_new_dir, seed_name, also_bootstrap_wrap=False)
    elif cur is not None:
        materialize_wrap_seed(theme_new_dir, cur.name, also_bootstrap_wrap=False)

    src = _resolve_src(theme_new_dir, compose_name)
    if src is None:
        raise FileNotFoundError(compose_name)

    for p in refs.glob(f"{WRAP_STEM}.*"):
        p.unlink(missing_ok=True)
    ext = src.suffix.lower() or ".jpg"
    dest = refs / f"{WRAP_STEM}{ext}"
    shutil.copy2(src, dest)
    # refs に単体名も残す（apply_human_selection の名前解決用）
    named = refs / compose_name
    if src.resolve() != named.resolve():
        shutil.copy2(src, named)

    return {
        "ok": True,
        "composePath": str(dest.resolve()),
        "composeName": compose_name,
        "seedPath": str(existing_wrap_seed_path(theme_new_dir) or ""),
    }
