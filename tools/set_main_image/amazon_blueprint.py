# -*- coding: utf-8 -*-
"""
Amazon 見本（03）からセット数に合う LAYOUT_BLUEPRINT を選ぶ。

選定は「見本スキャン特徴（patternHint / leftShare / inkRatio）」に基づく。
ファイル名にセット数が無いため、N→密度・構図パターンの対応表でスコアする。
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from PIL import Image

from work_paths import SUB_AMAZON_REF, default_work_root

LOG = logging.getLogger("set_main_image.amazon_blueprint")

_SCAN_REPORT = Path(__file__).resolve().parent / "sample_scan_report.json"


@dataclass
class BlueprintDecision:
    set_count: int
    path: Path
    file_name: str
    pattern_hint: str
    band: str
    preferred_pattern: str
    target_ink: float
    ink_ratio: float
    left_share: float
    score: float
    reason_ja: str
    alternatives: List[Dict[str, Any]]


def _band_for(n: int) -> str:
    if n <= 2:
        return "few"  # 少数・ヒーロー大きめ
    if n == 3:
        return "triangleish"  # 3個感（ピラミッド寄り）
    if n <= 6:
        return "medium"  # 横／浅めスタック
    if n <= 15:
        return "dense_stack"  # 左スタック厚め
    return "many_cluster"  # 多数・中央寄り可


def _preferred_pattern(band: str) -> str:
    # layout_rules: 「左側に数量スタック、右下に大きめヒーロー」が正。
    # 多数でも centered 山積みは「裏にもある」誤認を招きやすい → スタック型を優先。
    del band
    return "hero_right_stack_left"


def _target_ink(n: int) -> float:
    """セット数が多いほど見本のインク占有（密度）を高く取る。"""
    n2 = max(2, min(int(n), 30))
    return round(0.64 + (n2 - 2) / 28.0 * 0.22, 3)


def _ink_ratio(path: Path) -> float:
    im = Image.open(path).convert("RGB").resize((240, 240))
    px = im.load()
    ink = 0
    for y in range(240):
        for x in range(240):
            r, g, b = px[x, y]
            if r + g + b < 720:
                ink += 1
    return round(ink / (240 * 240), 3)


def _load_scan_items() -> List[Dict[str, Any]]:
    if not _SCAN_REPORT.is_file():
        return []
    data = json.loads(_SCAN_REPORT.read_text(encoding="utf-8"))
    return list(((data.get("amazonRefs") or {}).get("items")) or [])


def _catalog(work_root: Path) -> List[Dict[str, Any]]:
    folder = work_root / SUB_AMAZON_REF
    if not folder.is_dir():
        raise FileNotFoundError(f"見本フォルダがありません: {folder}")
    scan = {it["file"]: it for it in _load_scan_items() if it.get("file")}
    out: List[Dict[str, Any]] = []
    for p in sorted(folder.iterdir(), key=lambda x: x.name):
        if p.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
            continue
        meta = dict(scan.get(p.name) or {})
        meta["path"] = p
        meta["file"] = p.name
        if "inkRatio" not in meta:
            meta["inkRatio"] = _ink_ratio(p)
        if "patternHint" not in meta:
            meta["patternHint"] = "unknown"
        if "leftShare" not in meta:
            meta["leftShare"] = 0.5
        out.append(meta)
    if not out:
        raise FileNotFoundError(f"見本画像がありません: {folder}")
    return out


def select_amazon_blueprint(
    set_count: int,
    work_root: Optional[Path] = None,
    *,
    prefer_unused: Optional[Set[str]] = None,
    explicit: Optional[Path] = None,
) -> BlueprintDecision:
    """
    set_count に合う LAYOUT_BLUEPRINT を1枚選ぶ。
    prefer_unused: バッチ内で既に使ったファイル名を避けたいときに渡す。
    """
    root = work_root or default_work_root()
    if explicit and explicit.is_file():
        ink = _ink_ratio(explicit)
        return BlueprintDecision(
            set_count=int(set_count),
            path=explicit,
            file_name=explicit.name,
            pattern_hint="explicit",
            band=_band_for(int(set_count)),
            preferred_pattern=_preferred_pattern(_band_for(int(set_count))),
            target_ink=_target_ink(int(set_count)),
            ink_ratio=ink,
            left_share=0.0,
            score=999.0,
            reason_ja=f"CLI/--reference で明示指定: {explicit.name}",
            alternatives=[],
        )

    n = int(set_count)
    band = _band_for(n)
    pref = _preferred_pattern(band)
    target = _target_ink(n)
    catalog = _catalog(root)
    used = prefer_unused or set()

    scored: List[Tuple[float, Dict[str, Any], str]] = []
    for it in catalog:
        pattern = str(it.get("patternHint") or "unknown")
        ink = float(it.get("inkRatio") or 0)
        left = float(it.get("leftShare") or 0.5)
        score = 0.0
        bits: List[str] = []

        if pattern == pref:
            score += 100
            bits.append(f"希望パターン一致({pref})")
        elif pref == "hero_right_stack_left" and pattern == "hero_left_stack_right":
            score += 55
            bits.append("左右反転型（許容・次点）")
        elif pattern == "centered_cluster":
            # 山積み暗示になりやすいので減点（明示指定時以外）
            score += 5
            bits.append("中央クラスター（個数誤認リスク・低優先）")
        else:
            score += 10
            bits.append(f"パターン不一致({pattern})")

        # 密度: セット数に応じた ink 目標との近さ
        dens = max(0.0, 40.0 - abs(ink - target) * 220.0)
        score += dens
        bits.append(f"密度ink={ink} 目標={target} 近さ点={dens:.1f}")

        # 左スタック厚み（hero_right_stack_left 系）
        if pref == "hero_right_stack_left":
            # Nが大きいほど leftShare が高い見本を優遇
            want_left = 0.50 + min(0.08, (n - 2) * 0.003)
            left_pts = max(0.0, 25.0 - abs(left - want_left) * 120.0)
            score += left_pts
            bits.append(f"leftShare={left} 目標≈{want_left:.3f} 点={left_pts:.1f}")

        if it["file"] in used:
            score -= 40
            bits.append("バッチ内再利用ペナルティ")

        reason = " / ".join(bits)
        scored.append((score, it, reason))

    scored.sort(key=lambda x: (-x[0], x[1]["file"]))
    best_score, best, best_bits = scored[0]
    alts = []
    for sc, it, bits in scored[1:4]:
        alts.append(
            {
                "file": it["file"],
                "score": round(sc, 2),
                "patternHint": it.get("patternHint"),
                "inkRatio": it.get("inkRatio"),
                "leftShare": it.get("leftShare"),
                "scoreDetail": bits,
            }
        )

    reason_ja = (
        f"N={n} → 帯={band}。希望構図={pref}（layout_rules: 左スタック＋右ヒーロー）。"
        f"見本ファイル名にセット数が無いため、スキャン特徴（patternHint/leftShare）と"
        f"実測ink密度でスコア選定。採用={best['file']}（score={best_score:.1f}）。"
        f"内訳: {best_bits}"
    )
    LOG.info(
        "blueprint N=%s file=%s pattern=%s score=%.1f | %s",
        n,
        best["file"],
        best.get("patternHint"),
        best_score,
        reason_ja,
    )
    return BlueprintDecision(
        set_count=n,
        path=best["path"],
        file_name=best["file"],
        pattern_hint=str(best.get("patternHint")),
        band=band,
        preferred_pattern=pref,
        target_ink=target,
        ink_ratio=float(best.get("inkRatio") or 0),
        left_share=float(best.get("leftShare") or 0),
        score=round(best_score, 2),
        reason_ja=reason_ja,
        alternatives=alts,
    )


def decision_to_dict(d: BlueprintDecision) -> Dict[str, Any]:
    out = asdict(d)
    out["path"] = str(d.path)
    return out
