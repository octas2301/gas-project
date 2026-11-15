# -*- coding: utf-8 -*-
"""MAIN場面の個包装を見本に寄せる後処理（仮想PASS案）。

- フィルム分岐は機械（wrap look の透明度・材質）。人間は検収のみ
- 透明: 切抜き＋暗い台載せ見本＋乳白色禁止の差し替え
- 不透明: 生見本＋色・印字を守る差し替え（乳白色禁止は付けない）
- 左下は個包装金丸聖域。差し替え指示でも左下禁止
切戻し: B4_WRAP_SCENE_REFINE=0
"""
from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import numpy as np
from PIL import Image

LOG = logging.getLogger("set_main_image.wrap_scene_refine")


def wrap_scene_refine_enabled() -> bool:
    raw = str(os.environ.get("B4_WRAP_SCENE_REFINE") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def classify_wrap_film_branch(fields: Optional[Dict[str, Any]]) -> str:
    """transparent | opaque | unknown。案件名を焼かない。"""
    f = fields if isinstance(fields, dict) else {}
    blob = " ".join(
        str(f.get(k) or "") for k in ("透明度", "材質", "縁", "密着", "形")
    )
    opaque_hits = (
        "不透明",
        "乳白",
        "アルミ",
        "紙袋",
        "印刷で中",
        "中身が見えない",
        "中が見えない",
        "白い袋",
        "マットで中",
    )
    # 「不透明」は「透明」を含むので、裸の「透明」一致は不透明時に使わない
    clear_hits_safe = (
        "透け",
        "クリア",
        "中身がはっきり",
        "中がはっきり",
        "全体が透明",
        "光沢フィルム",
    )
    op = any(h in blob for h in opaque_hits)
    cl = any(h in blob for h in clear_hits_safe)
    if "透明" in blob and "不透明" not in blob:
        cl = True
    if "透けない" in blob:
        cl = False
        op = True
    if op and not cl:
        return "opaque"
    if cl and not op:
        return "transparent"
    if op:
        return "opaque"
    if cl:
        return "transparent"
    mat = str(f.get("材質") or "")
    if "フィルム" in mat and "不透明" not in mat:
        return "transparent"
    return "unknown"


def film_branch_from_looks(looks: Sequence[Dict[str, Any]]) -> str:
    """複数見本は opaque が1つでもあれば opaque。透明のみなら transparent。"""
    branches = []
    for rec in looks or []:
        if not isinstance(rec, dict):
            continue
        b = str(rec.get("filmBranch") or "").strip()
        if not b:
            b = classify_wrap_film_branch(
                rec.get("fields") if isinstance(rec.get("fields"), dict) else {}
            )
        branches.append(b)
    if not branches:
        return "unknown"
    if any(b == "opaque" for b in branches):
        return "opaque"
    if any(b == "transparent" for b in branches):
        return "transparent"
    return "unknown"


def prep_wrap_cutout_on_scene_dark(
    wrap_path: Path,
    scene_path: Path,
    out_path: Path,
) -> Path:
    """透明分岐用: 明るい背景を落とし、場面の暗い台色の上に載せる。"""
    wrap_path = Path(wrap_path)
    scene_path = Path(scene_path)
    out_path = Path(out_path)
    wim = Image.open(wrap_path).convert("RGBA")
    arr = np.array(wim)
    corners = np.stack(
        [
            arr[0:15, 0:15, :3].reshape(-1, 3).mean(0),
            arr[0:15, -15:, :3].reshape(-1, 3).mean(0),
            arr[-15:, 0:15, :3].reshape(-1, 3).mean(0),
            arr[-15:, -15:, :3].reshape(-1, 3).mean(0),
        ],
        axis=0,
    )
    bg = corners.mean(axis=0)
    dist = np.linalg.norm(arr[:, :, :3].astype(np.float32) - bg, axis=2)
    mx = arr[:, :, :3].max(axis=2).astype(np.float32)
    mn = arr[:, :, :3].min(axis=2).astype(np.float32)
    sat = (mx - mn) / (mx + 1e-6)
    alpha = np.where((dist < 42) | ((mx > 180) & (sat < 0.2)), 0, 255).astype(np.uint8)
    blueish = (arr[:, :, 2] > arr[:, :, 0] + 25) & (arr[:, :, 2] > 80)
    alpha[blueish] = 255
    arr[:, :, 3] = alpha
    cut = Image.fromarray(arr, "RGBA")

    mim = Image.open(scene_path).convert("RGB")
    mw, mh = mim.size
    samples = []
    for fx, fy in ((0.08, 0.88), (0.12, 0.92), (0.05, 0.78), (0.2, 0.9)):
        samples.append(
            np.array(mim.getpixel((int(mw * fx), int(mh * fy))), dtype=np.float32)
        )
    dark = tuple(int(x) for x in np.median(np.stack(samples, axis=0), axis=0))
    base = Image.new("RGB", cut.size, dark)
    base.paste(cut, (0, 0), cut)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    base.save(out_path, quality=92)
    return out_path


def build_wrap_replace_prompt(
    *,
    branch: str,
    look_lines: str = "",
) -> str:
    place = (
        "個包装は皿の横か後ろのみ。画面左下（個包装金丸の聖域）に個包装を置くな。"
        "カメラ最近傍・下端手前は開封後の個品。個包装を手前に出すな。"
    )
    base = (
        "【厳守】2枚目が個包装の正。1枚目の開封前個包装だけを差し替え。\n"
        "大袋・開封後・文字・金メダル・Octasは変えるな。個包装は1〜2個。\n"
        f"{place}\n"
        "2枚目の机・植物などの背景はコピーするな（個包装の物体だけ）。\n"
    )
    if look_lines:
        base += f"機械読み: {look_lines}\n"
    if branch == "opaque":
        base += (
            "フィルム／袋は見本の不透明さ・色・印字・絵どおり。"
            "勝手に透明化・乳白色化・別柄にしない。密着・空気は見本どおり。\n"
        )
    else:
        # transparent / unknown → 透明側の安全策（仮想PASS）
        base += (
            "フィルムは無色透明の硬い光沢（ガラス〜硬いPET）。\n"
            "乳白色・曇り白・不透明白のフィルム禁止。白く見えるのはハイライト反射だけ。\n"
            "2枚目で縁が白く見えるのは背景が透けているだけ。フィルム色ではない。\n"
            "完成画の台が暗ければ、縁のシール部も台の色が透けて見えること。\n"
            "密着・空気なし・あけ口は2枚目どおり。\n"
        )
    return base


def resolve_wrap_ref_for_refine(
    *,
    wrap_path: Path,
    scene_path: Path,
    branch: str,
    work_dir: Path,
) -> Path:
    wrap_path = Path(wrap_path)
    if branch == "opaque":
        return wrap_path
    # transparent / unknown
    dest = Path(work_dir) / f"wrap_ref_{branch}_{wrap_path.stem}.jpg"
    try:
        return prep_wrap_cutout_on_scene_dark(wrap_path, scene_path, dest)
    except Exception as e:
        LOG.warning("wrap cutout prep failed: %s", e)
        return wrap_path


def refine_scene_wrap_to_file(
    scene_path: Path,
    wrap_path: Path,
    out_path: Path,
    *,
    looks: Optional[Sequence[Dict[str, Any]]] = None,
    work_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    場面の個包装を見本に寄せて out へ書く。
    戻り: ok / skipped / branch / model / note
    """
    if not wrap_scene_refine_enabled():
        return {"ok": True, "skipped": True, "reason": "B4_WRAP_SCENE_REFINE=0"}
    scene_path = Path(scene_path)
    wrap_path = Path(wrap_path)
    out_path = Path(out_path)
    if not scene_path.is_file() or not wrap_path.is_file():
        return {"ok": False, "skipped": True, "reason": "missing_file"}

    branch = film_branch_from_looks(looks or [])
    if branch == "unknown":
        # 単一見本の fields から再判定
        if looks:
            fields = {}
            for rec in looks:
                if isinstance(rec, dict) and isinstance(rec.get("fields"), dict):
                    fields = rec["fields"]
                    break
            branch = classify_wrap_film_branch(fields)

    look_lines = ""
    for rec in looks or []:
        if isinstance(rec, dict) and rec.get("lines"):
            look_lines = str(rec.get("lines"))
            break

    work = Path(work_dir or (out_path.parent / "_wrap_refine"))
    work.mkdir(parents=True, exist_ok=True)
    ref = resolve_wrap_ref_for_refine(
        wrap_path=wrap_path,
        scene_path=scene_path,
        branch=branch if branch in ("transparent", "opaque") else "transparent",
        work_dir=work,
    )
    prompt = build_wrap_replace_prompt(
        branch=branch if branch in ("transparent", "opaque") else "transparent",
        look_lines=look_lines,
    )
    from openai_image import generate_with_references

    raw, model = generate_with_references(
        prompt=prompt,
        image_paths=[scene_path, ref],
        size="1024x1024",
        quality="high",
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(raw)
    meta = {
        "ok": True,
        "skipped": False,
        "branch": branch,
        "model": model,
        "wrapRef": str(ref),
        "out": str(out_path),
        "promptHead": prompt[:180],
    }
    (work / "last_refine.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    LOG.info(
        "wrap scene refine branch=%s model=%s out=%s",
        branch,
        model,
        out_path.name,
    )
    return meta
