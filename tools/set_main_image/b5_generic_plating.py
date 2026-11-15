# -*- coding: utf-8 -*-
"""一片見本・盛り（MAIN/SUB）共通レイヤ（C1〜C7 / T4/T8核）。

正本: docs/org/PLATING_AND_PIECE_GENERIC_LOGIC.md
汎用性の定義: Agent Store GENERICITY_DEFINITION.md（型 vs 一件焼き）

切戻し: B5_GENERIC_PLATING_CORE=0
個別適用プロファイル: B5_PLATING_PROFILE=generic|food_flake（既定 food_flake＝削り系）
"""
from __future__ import annotations

import os
import re
from pathlib import Path
from typing import List, Optional, Sequence

# --- C2: 危険文（縁の規則アーティファクトを増やす言い回し）---
_DANGER_EDGE_PATTERNS = [
    r"おつまみのさきいかのように[^\n]*",
    r"細い長い繊維が縁から飛び出し[^\n]*",
    r"ちりぢり・糸状ほつれが不規則[^\n]*",
    r"縁と切れ端は糸状の魚肉繊維がほつれ・毛羽立ち・裂けていること[^\n]*",
    r"／人間修正:[^\n]*",
    r"細かなギザつきや凹凸[^\n]*",
    r"細かなギザギザやほつれ[^\n]*",
    r"不揃いなギザつき[^\n]*",
    r"全体的に細かな凹凸やギザつき[^\n]*",
]


def generic_plating_core_enabled() -> bool:
    raw = str(os.environ.get("B5_GENERIC_PLATING_CORE") or "1").strip().lower()
    return raw not in ("0", "false", "off", "no")


def plating_profile() -> str:
    raw = str(os.environ.get("B5_PLATING_PROFILE") or "food_flake").strip().lower()
    return raw or "food_flake"


def sanitize_danger_prompt(text: str) -> str:
    """C2: NGを増やす指示文を落とす（SKU非焼付）。"""
    raw = str(text or "")
    if not raw:
        return ""
    for pat in _DANGER_EDGE_PATTERNS:
        raw = re.sub(pat, "", raw)
    raw = re.sub(
        r"(-\s*縁:\s*)([^\n]+)",
        r"\1滑らか〜弱い不規則の丸み。等間隔ギザ・鋸歯・規則的とげ列は禁止。",
        raw,
    )
    return re.sub(r"\n{3,}", "\n\n", raw).strip()


def plating_core_prompt_block(*, caller: str = "plating") -> str:
    """盛り（MAIN/SUB）および一片に同送する短い固定核。

    caller: piece | plating | main | sub（ログ用。文面は共通＋プロファイル）
    """
    _ = caller
    prof = plating_profile()
    # 汎用核（どの商品でも同じ型）
    core = (
        "【共通・盛り／一片核】\n"
        "参照の役割を混ぜるな（面の正・形の正・盛り配置・袋は別）。\n"
        "規則的なアーティファクト（等間隔の歯・くし歯・作った感じの縁）は不合格。\n"
        "見たい属性は面の中に残せ。縁の飾りで代用するな。\n"
    )
    # 個別適用（プロファイル）
    if prof == "food_flake":
        core += (
            "【個別適用・削り／薄片】\n"
            "身の平らな面に長い木目筋を残せ。つるつる無地は不合格。\n"
            "縁は滑らか〜弱い不規則の丸み／薄いささくれ。鋸歯・ピンキング禁止。\n"
            "しなやかなねじれ・丸まりを一部に。硬い板の平行並べ禁止。\n"
        )
    else:
        core += (
            "【個別適用・generic】\n"
            "その商品で必須の質感属性を面に残せ。規則縁で代用するな。\n"
        )
    return core.strip()


def ensure_surface_face_crop(
    surface_src: Path,
    dest: Path,
    *,
    box_frac: tuple[float, float, float, float] = (0.22, 0.22, 0.78, 0.78),
) -> Optional[Path]:
    """C4: 縁をほぼ切った面中央クロップ。"""
    try:
        from PIL import Image
    except ImportError:
        return None
    src = Path(surface_src)
    if not src.is_file():
        return None
    im = Image.open(src).convert("RGB")
    w, h = im.size
    l, t, r, b = box_frac
    box = (int(w * l), int(h * t), int(w * r), int(h * b))
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    im.crop(box).save(dest, quality=95)
    return dest if dest.is_file() else None


def prefer_face_crop_surfaces(
    theme_dir: Path,
    surface_paths: Sequence[Path],
    *,
    max_n: int = 3,
    work_dir: Optional[Path] = None,
) -> List[Path]:
    """SURFACE先頭に面中央クロップを足す（既存パスは維持）。"""
    paths = [Path(p) for p in surface_paths if Path(p).is_file()]
    if not paths or not generic_plating_core_enabled():
        return paths[: int(max_n or 3)]
    work = Path(work_dir) if work_dir else (Path(theme_dir) / "food_piece_set5_virt" / "deepen_refs")
    # 優先種: phone → 先頭SURFACE
    phone = Path(theme_dir) / "food_piece_set5_virt/deepen_refs/ref_surface_phone_IMG_8567.jpg"
    src = phone if phone.is_file() else paths[0]
    crop = ensure_surface_face_crop(src, work / "_generic_face_crop.jpg")
    out: List[Path] = []
    if crop and crop.is_file():
        out.append(crop)
    for p in paths:
        if p.resolve() not in {x.resolve() for x in out}:
            out.append(p)
        if len(out) >= int(max_n or 3):
            break
    return out[: int(max_n or 3)]


def apply_core_to_prompt_text(text: str, *, caller: str = "plating") -> str:
    """危険文掃除＋核ブロック追記。"""
    if not generic_plating_core_enabled():
        return str(text or "")
    base = sanitize_danger_prompt(text)
    core = plating_core_prompt_block(caller=caller)
    if core and core not in base:
        base = (base + "\n\n" + core).strip() if base else core
    return base


def recommend_quality(logic: str) -> str:
    """C7: 判定できる最安段。high は返さない（既定オフ）。"""
    key = str(logic or "").strip().lower()
    medium_keys = (
        "fiber",
        "suji",
        "筋",
        "micro",
        "chiai",
        "血合",
        "pale",
        "淡色",
        "adopt",
        "合否",
        "plating_final",
        "盛り最終",
    )
    if any(k in key for k in medium_keys):
        return "medium"
    return "low"
