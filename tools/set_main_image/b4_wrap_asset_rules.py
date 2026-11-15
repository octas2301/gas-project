# -*- coding: utf-8 -*-
"""P3-W 個包装アピール素材（見本ゲート分岐・仮想／本線共通の純関数）。

- WRAP確定後だけ候補生成レーンに入る
- 背景: 白〜生成り＋植物緑のぼけ可。木目必須にしない。グレー一色禁止
- 手持ち実寸: 実物は約2.5–3.5cm（親指の先〜第一関節程度）。巨大化禁止
- 採用1枚を MAIN左下金丸／サブ個包装訴求／場面WRAP参照へ共有
切戻し想定:
- 手持ち素材レーン: B4_WRAP_ASSET_LANE=0
- サブ個包装テーマ+1: B4_WRAP_APPEAL_THEME=0
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

# LP 1..20 の外。材料門では出さず、WRAP ありのときだけ機械追加。
THEME_ID_WRAP_APPEAL = 21
THEME_NAME_WRAP_APPEAL = "個包装アピール"
MAX_THEMES_BASE = 10
MAX_THEMES_WITH_WRAP_EXTRA = 11  # 通常10＋個包装訴求1枠


def wrap_asset_lane_enabled() -> bool:
    raw = str(os.environ.get("B4_WRAP_ASSET_LANE") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def wrap_appeal_theme_enabled() -> bool:
    raw = str(os.environ.get("B4_WRAP_APPEAL_THEME") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def max_themes_allowed(*, has_wrap_appeal: bool = False) -> int:
    if has_wrap_appeal:
        return MAX_THEMES_WITH_WRAP_EXTRA
    return MAX_THEMES_BASE


def mechanical_wrap_appeal_proposal() -> Dict[str, Any]:
    """人手の型・文案発明なし。固定の機械文（検収で直してよい）。"""
    return {
        "themeId": THEME_ID_WRAP_APPEAL,
        "name": THEME_NAME_WRAP_APPEAL,
        "phase": "プロダクト理解",
        "phaseOrder": 2,
        "source": "wrap_asset",
        "kind": "wrap_appeal",
        "evidence": "FOOD_UNIT_WRAP",
        "appealAxis": "個包装の便利さ",
        "overview": (
            "個包装の真空パックを主役に、食べたい分だけ取り出せる便利さと持ち運びやすさを伝える。"
            "手元や卓上で個包装がはっきり見える。明るい白〜生成り。植物の緑はぼけて少し。"
            "大袋カタログ並べは主役にしない。"
        ),
        "catch": "個包装で、好きな分だけ",
        "subs": ["真空パックで鮮度をキープ", "持ち運びにも便利"],
        "needPackage": False,
        "checkedDefault": True,
        "selected": True,
        "wrapAppeal": True,
    }


def proposals_have_wrap_appeal(proposals: Sequence[Dict[str, Any]]) -> bool:
    for pr in proposals or []:
        if not isinstance(pr, dict):
            continue
        try:
            if int(pr.get("themeId") or 0) == THEME_ID_WRAP_APPEAL:
                return True
        except (TypeError, ValueError):
            pass
        if pr.get("wrapAppeal") is True:
            return True
    return False


def merge_wrap_appeal_theme(
    proposals: Sequence[Dict[str, Any]],
    *,
    has_wrap: bool,
) -> List[Dict[str, Any]]:
    """
    WRAP あり → themeId=21 を末尾に1枠追加（既にあれば触らない）。
    通常案は最大10のまま。合計は最大11。
    """
    out: List[Dict[str, Any]] = [dict(p) for p in (proposals or []) if isinstance(p, dict)]
    if not wrap_appeal_theme_enabled() or not has_wrap:
        return out
    if proposals_have_wrap_appeal(out):
        return out
    # 通常枠を10までに整え、+1で個包装を足す
    base = [p for p in out if int(p.get("themeId") or 0) != THEME_ID_WRAP_APPEAL]
    if len(base) > MAX_THEMES_BASE:
        base = base[:MAX_THEMES_BASE]
    base.append(mechanical_wrap_appeal_proposal())
    return base


def should_branch_after_wrap_confirm(*, has_wrap: bool, wrap_path: Optional[Path] = None) -> bool:
    """見本ゲートで WRAP を選んだときだけ素材レーンへ分岐。"""
    if not wrap_asset_lane_enabled():
        return False
    if not has_wrap:
        return False
    if wrap_path is not None and not Path(wrap_path).is_file():
        return False
    return True


def hand_hold_bg_contract_errors(prompt: str) -> List[str]:
    """手持ち生成指示の背景契約。"""
    p = str(prompt or "")
    errs: List[str] = []
    # 「グレー一色禁止」は正。背景をグレーに指定しているときだけ落とす。
    bans_gray = ("グレー一色禁止" in p) or ("not flat gray" in p.lower()) or ("NOT gray" in p)
    asks_gray = (
        ("背景はグレー" in p)
        or ("背景をグレー" in p)
        or ("flat gray background" in p.lower())
        or (("グレー一色" in p) and ("禁止" not in p))
    )
    if asks_gray and not bans_gray:
        errs.append("bg:flat_gray_ban_self")
    # 必須語（白系）
    white_ok = any(x in p for x in ("白", "生成り", "off-white", "soft white", "warm off-white"))
    if not white_ok:
        errs.append("bg:white_required")
    # 木目を正にしない（禁止というより「必須にしない」＝木目だけの指定はNG）
    if ("木目" in p or "wood grain" in p.lower()) and not white_ok:
        errs.append("bg:wood_without_white")
    # 植物緑は推奨（仮想では有りを検査）
    plant_ok = any(x in p for x in ("植物", "緑", "foliage", "leaves", "plant"))
    if not plant_ok:
        errs.append("bg:plant_recommended")
    return errs


def hand_hold_scale_contract_errors(prompt: str) -> List[str]:
    """実寸・手対比の契約。"""
    p = str(prompt or "")
    errs: List[str] = []
    need = (
        ("2.5", "scale:cm_hint"),
        ("親指", "scale:thumb_ja"),
    )
    # English fallbacks acceptable
    if "2.5" not in p and "3cm" not in p.lower() and "3.5" not in p:
        errs.append("scale:cm_hint")
    if "親指" not in p and "thumb" not in p.lower():
        errs.append("scale:thumb")
    if "巨大" not in p and "NOT a huge" not in p and "not oversized" not in p.lower() and "小さく" not in p:
        errs.append("scale:anti_giant")
    return errs


def build_hand_hold_prompt_ja(
    *,
    circular: bool = True,
    include_plant: bool = True,
) -> str:
    """手持ち生成の定型（案件名を焼かない）。

    見本選び用の写真。金丸はめ込みは MAIN 後載せ（b4_wrap_bl_gold_stamp）側の仕事。
    ここでは金枠・丸切り前提の構図指示を入れない。
    """
    lines = [
        "【個包装手持ち・実寸必須】",
        "個包装は実物どおり一口サイズ（およそ幅2.5〜3.5cm）。大人の手では親指の先〜第一関節くらい。",
        "指に対して小さく見えよ。巨大化禁止。",
        "個包装の見た目の正は添付 WRAP 見本のみ。スケール参考写真の食品の形・色はコピーするな（手の大きさ・背景の雰囲気だけ借りる）。",
        "フィルムは透明。あけ口以外の印字・ロゴ・青帯を足すな。空気だまりを見本に無く作るな。大袋は出すな。",
        "背景は白〜生成り（どこにでもなじむ）。グレー一色禁止。木目は必須にしない。",
        "金枠・金リング・丸い金色の縁・メダル枠は描くな。",
    ]
    if include_plant:
        lines.append("背景の端に、ぼけた植物の緑が少し映るとよい（主役は個包装）。")
    if circular:
        # 旧互換の引数。見本選び用なので丸切り・金枠差し込みは指示しない
        pass
    lines.append("指とフィルムが溶けて一体になるな。写真実写。正方形の実写写真。")
    return "\n".join(lines)


def candidate_manifest(
    *,
    jan: str,
    wrap_path: str,
    candidates: List[str],
    selected: str = "",
) -> Dict[str, Any]:
    """採用前の候補一覧。selected 空＝未採用。"""
    return {
        "kind": "wrap_hand_hold_candidates",
        "jan": str(jan or ""),
        "wrapPath": str(wrap_path or ""),
        "candidates": [str(x) for x in candidates],
        "selected": str(selected or ""),
        "bg": "white_offwhite_plant_bokeh",
        "scale": "thumb_2_5_to_3_5cm",
    }


def apply_selection(manifest: Dict[str, Any], selected_path: str) -> Dict[str, Any]:
    out = dict(manifest or {})
    sel = str(selected_path or "").strip()
    cands = [str(x) for x in (out.get("candidates") or [])]
    if sel and sel not in cands:
        raise ValueError("selected not in candidates")
    out["selected"] = sel
    return out


def resolve_wrap_asset_for_consumers(manifest: Dict[str, Any]) -> Dict[str, str]:
    """
    採用後の届け先キー。
    - MAIN左下金丸／サブ個包装訴求 → 手持ち selected
    - 場面WRAP参照 → 検収1個画 composeWrap（無ければ手持ちにフォールバック）
    """
    man = manifest or {}
    sel = str(man.get("selected") or "").strip()
    compose = str(man.get("composeWrap") or man.get("selectedSingle") or "").strip()
    if not sel and not compose:
        return {"mainBlGold": "", "subWrapAppeal": "", "sceneWrapRef": ""}
    return {
        "mainBlGold": sel,
        "subWrapAppeal": sel,
        "sceneWrapRef": compose or sel,
    }


def sub_wrap_appeal_via_theme_pipeline() -> bool:
    """②: サブ個包装訴求は専用寂しレイアウト固定ではなく、テーマ案パイプラインに載せる。"""
    return True


def is_wrap_appeal_theme_id(theme_id: Any) -> bool:
    try:
        return int(theme_id or 0) == THEME_ID_WRAP_APPEAL
    except (TypeError, ValueError):
        return False
