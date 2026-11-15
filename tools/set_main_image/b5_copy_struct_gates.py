# -*- coding: utf-8 -*-
"""B-⑤／B-④ 画像前・文案構造ゲート（要件 §0.1 S1〜S14）。

画像API・GASは呼ばない。fixture の提案JSON／プロンプト断片を機械判定する。
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Tuple

# 全角換算（要件: 全角=1、半角=0.5、切り上げに近い簡易）
def zenkaku_len(s: str) -> float:
    n = 0.0
    for ch in str(s or ""):
        o = ord(ch)
        n += 0.5 if o < 128 else 1.0
    return n


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s or ""))


# 不自然な末尾（短縮・キャッチ）
_BAD_TAIL = re.compile(r"[てにをはがのとでもへや・／、。]$")

# 旧プロンプト禁止（S5）
_OLD_PERSONA_BAN = "材料に無い相手（ペルソナ属性）を作って用事にしない"


def gate_s1_mix_both(
    proposals: List[Dict[str, Any]],
    *,
    is_mix: bool,
    mix_marks: List[str],
) -> Tuple[bool, str]:
    """S1: ミックスなら両方の特徴が読める。"""
    if not is_mix:
        return True, "skip_not_mix"
    marks = [m for m in mix_marks if str(m).strip()]
    if len(marks) < 2:
        return False, "mix_marks_need_2"
    for i, p in enumerate(proposals):
        blob = _norm(
            str(p.get("catch") or "")
            + str(p.get("catchBurn") or "")
            + str(p.get("salesTalk") or "")
        )
        hits = [m for m in marks if _norm(m) and _norm(m) in blob]
        if len(hits) < 2:
            return False, f"proposal[{i}] only_hits={hits}"
    return True, "ok"


def gate_s2_catch_sales_roles(proposals: List[Dict[str, Any]]) -> Tuple[bool, str]:
    """S2: 全案に catch と salesTalk。複製禁止。字数帯。"""
    for i, p in enumerate(proposals):
        c = str(p.get("catch") or "").strip()
        s = str(p.get("salesTalk") or "").strip()
        if not c:
            return False, f"proposal[{i}] empty_catch"
        if not s:
            return False, f"proposal[{i}] empty_salesTalk"
        if _norm(c) == _norm(s):
            return False, f"proposal[{i}] catch_eq_sales"
        cl, sl = zenkaku_len(c), zenkaku_len(s)
        if cl < 8 or cl > 40:
            return False, f"proposal[{i}] catch_len={cl}"
        # 画面用は最大20。12超なら salesBurn（画像用）必須
        if sl < 4 or sl > 20:
            return False, f"proposal[{i}] sales_len={sl}"
        if sl > 12:
            sb = str(p.get("salesBurn") or "").strip()
            if not sb or zenkaku_len(sb) < 4 or zenkaku_len(sb) > 12:
                return False, f"proposal[{i}] long_sales_needs_salesBurn"
    return True, "ok"


def gate_s3_persona(proposals: List[Dict[str, Any]], *, ui_can_open_detail: bool) -> Tuple[bool, str]:
    """S3: 短文＋詳細。UIで詳細を開けること。"""
    if not ui_can_open_detail:
        return False, "ui_cannot_open_detail"
    for i, p in enumerate(proposals):
        short = str(p.get("personaShort") or "").strip()
        detail = str(p.get("personaDetail") or "").strip()
        if not short:
            return False, f"proposal[{i}] no_personaShort"
        if not detail:
            return False, f"proposal[{i}] no_personaDetail"
        # 画像プロンプトに詳細全文を焼かない（フィールド burnPersona 禁止）
        if p.get("burnPersona") or p.get("personaInImage"):
            return False, f"proposal[{i}] persona_marked_for_image"
    return True, "ok"


def gate_s4_catch_burn(proposals: List[Dict[str, Any]]) -> Tuple[bool, str]:
    """S4: catchBurn 6〜18・末尾違和感なし。"""
    for i, p in enumerate(proposals):
        b = str(p.get("catchBurn") or "").strip()
        if not b:
            return False, f"proposal[{i}] empty_catchBurn"
        ln = zenkaku_len(b)
        if ln < 6 or ln > 18:
            return False, f"proposal[{i}] catchBurn_len={ln}"
        # 【】内だけでも可。全体末尾チェック
        body = b
        m = re.search(r"【([^】]+)】(.*)$", b)
        if m:
            body = (m.group(1) + m.group(2)).strip()
        if _BAD_TAIL.search(body):
            return False, f"proposal[{i}] bad_tail={body[-3:]}"
        # 明らかな途中切れ（ひらがな1字で終わる「て」等は上で検知）
        if re.search(r"[一-龥][ぁ-ん]$", body) and len(body) >= 2:
            # 語尾の助詞・接続は上で禁止。漢字＋ひらがな1字（例: 丼）はここでは落とさない
            pass
    return True, "ok"


def gate_s5_persona_in_gen(
    *,
    polish_prompt: str,
    persona_short: str,
) -> Tuple[bool, str]:
    """S5: 整文プロンプトにペルソナがあり、旧禁止文が無い。"""
    prompt = str(polish_prompt or "")
    if _OLD_PERSONA_BAN in prompt:
        return False, "old_persona_ban_line_present"
    ps = str(persona_short or "").strip()
    if not ps:
        return False, "no_persona_short"
    if ps not in prompt and "ペルソナ" not in prompt:
        return False, "persona_not_in_prompt"
    if ps not in prompt:
        # 「ペルソナ:」ラベルだけでも可とするが、短文も載ることが望ましい
        if "ペルソナ:" not in prompt and "ペルソナ：" not in prompt:
            return False, "persona_label_missing"
    return True, "ok"


def gate_s6_remake_checked_only(
    *,
    before: List[Dict[str, Any]],
    after: List[Dict[str, Any]],
    checked_keys: List[str],
) -> Tuple[bool, str]:
    """S6: チェックした key 以外の catch が変わっていない。"""
    if not checked_keys:
        return False, "no_checked_keys"
    before_map = {_proposal_key(p): p for p in before}
    after_map = {_proposal_key(p): p for p in after}
    for k, bp in before_map.items():
        if k in checked_keys:
            continue
        ap = after_map.get(k)
        if not ap:
            return False, f"unchecked_missing_after key={k}"
        if str(bp.get("catch") or "") != str(ap.get("catch") or ""):
            return False, f"unchecked_catch_changed key={k}"
        if str(bp.get("salesTalk") or "") != str(ap.get("salesTalk") or ""):
            return False, f"unchecked_sales_changed key={k}"
        if str(bp.get("catchBurn") or "") != str(ap.get("catchBurn") or ""):
            return False, f"unchecked_burn_changed key={k}"
        if str(bp.get("salesBurn") or "") != str(ap.get("salesBurn") or ""):
            return False, f"unchecked_salesBurn_changed key={k}"
    for k in checked_keys:
        if k not in after_map:
            return False, f"checked_missing_after key={k}"
    return True, "ok"


def _proposal_key(p: Dict[str, Any]) -> str:
    return f"{p.get('themeId') or ''}:{p.get('classId') or ''}:{p.get('sellRank') or ''}"


def gate_s7_scene_hint(
    proposals: List[Dict[str, Any]],
    *,
    ui_shows_scene_editable: bool,
) -> Tuple[bool, str]:
    """S7: sceneHint があり、画面で表示・編集できる。"""
    if not ui_shows_scene_editable:
        return False, "ui_scene_not_editable"
    for i, p in enumerate(proposals):
        sc = str(p.get("sceneHint") or "").strip()
        if not sc:
            return False, f"proposal[{i}] empty_sceneHint"
        # 1〜3文のゆるい目安（句点または長さ）
        if zenkaku_len(sc) < 6:
            return False, f"proposal[{i}] scene_too_short"
    return True, "ok"


def naive_slice_burn(catch: str, max_n: float = 18.0) -> str:
    """字数切りだけの短縮（S8の不合格参照）。"""
    t = str(catch or "").strip()
    out = []
    n = 0.0
    for ch in t:
        add = 0.5 if ord(ch) < 128 else 1.0
        if n + add > max_n:
            break
        out.append(ch)
        n += add
    return "".join(out)


def gate_s8_semantic_burn(proposals: List[Dict[str, Any]]) -> Tuple[bool, str]:
    """S8: 長い catch に対し、単なる末尾スライスと同一なら不合格。"""
    for i, p in enumerate(proposals):
        c = str(p.get("catch") or "").strip()
        b = str(p.get("catchBurn") or "").strip()
        if not b:
            return False, f"proposal[{i}] empty_catchBurn"
        if zenkaku_len(c) <= 18:
            # 短い catch はそのまま短縮でも可
            continue
        naive = naive_slice_burn(c, 18.0)
        if _norm(b) == _norm(naive) and _BAD_TAIL.search(b.rstrip("】")):
            return False, f"proposal[{i}] burn_is_naive_slice_with_bad_tail"
        if _norm(b) == _norm(naive):
            # スライス同一かつ catch が明らかに長い → 意味単位とは言えない
            return False, f"proposal[{i}] burn_equals_naive_slice"
    return True, "ok"


def gate_s9_distinct_token_policy(
    *,
    is_mix: bool,
    sanitize_allows_single_distinct_for_mix: bool,
) -> Tuple[bool, str]:
    """S9: ミックスでは区別語1語残しポリシーをオフ（魚種に限らない）。"""
    if is_mix and sanitize_allows_single_distinct_for_mix:
        return False, "mix_still_allows_single_distinct"
    return True, "ok"


# 旧名互換
gate_s9_fish_policy = gate_s9_distinct_token_policy

_TASTE_FIRST = re.compile(
    r"^(おいしい|美味|香ばしい|うまい|サクサク|もちもち|濃厚|旨み|あま[いく]|甘[いく])"
)
_SCENE_SALES = ("お得", "最強", "今すぐ", "クリック", "送料無料", "必須アイテム")
_SCENE_PERSONA = ("主婦", "会社員", "歳の", "ママ", "ペルソナ", "ターゲット")
_PHOTO_CUE = ("袋", "皿", "手前", "背景", "光", "盛り", "テーブル", "見える", "置く", "並ぶ")
_DEFAULT_GATES = [
    "S1",
    "S2",
    "S3",
    "S4",
    "S5",
    "S6",
    "S7",
    "S8",
    "S9",
    "S10",
    "S11",
    "S12",
    "S13",
    "S14",
]


def _bracket_heading(s: str) -> str:
    m = re.search(r"【([^】]+)】", str(s or ""))
    return (m.group(1) or "").strip() if m else ""


def _errand_stems(heading: str) -> List[str]:
    h = str(heading or "").strip()
    if not h:
        return []
    stems = [h]
    for suf in ("ために", "への", "への一品", "に", "へ", "を", "の"):
        if h.endswith(suf) and len(h) > len(suf) + 1:
            stems.append(h[: -len(suf)])
            break
    # 短い重複を除く
    out: List[str] = []
    for t in stems:
        t = t.strip()
        if len(t) >= 2 and t not in out:
            out.append(t)
    return out


def gate_s10_same_errand(proposals: List[Dict[str, Any]]) -> Tuple[bool, str]:
    """S10: 1案でペルソナ・catch・burn・scene が同じ用事。"""
    for i, p in enumerate(proposals):
        tokens = [str(t).strip() for t in list(p.get("errandTokens") or []) if str(t).strip()]
        if not tokens:
            tokens = _errand_stems(_bracket_heading(str(p.get("catch") or "")))
        if not tokens:
            return False, f"proposal[{i}] no_errand_token"
        persona = _norm(str(p.get("personaShort") or "") + str(p.get("personaDetail") or ""))
        if not any(_norm(t) in persona for t in tokens):
            return False, f"proposal[{i}] persona_errand_mismatch"
        burn_h = _bracket_heading(str(p.get("catchBurn") or ""))
        catch_h = _bracket_heading(str(p.get("catch") or ""))
        if catch_h and burn_h and _norm(burn_h) != _norm(catch_h):
            # 短縮で助詞落ちは stems で許容
            if not (
                any(_norm(t) in _norm(burn_h) for t in tokens)
                or any(_norm(t) in _norm(catch_h) for t in _errand_stems(burn_h))
            ):
                return False, f"proposal[{i}] catchBurn_heading_mismatch"
        scene = str(p.get("sceneHint") or "")
        scene_h = _bracket_heading(scene)
        if scene_h and catch_h and _norm(scene_h) != _norm(catch_h):
            if not any(_norm(t) in _norm(scene_h) for t in tokens):
                return False, f"proposal[{i}] scene_other_errand"
        # 場面に別用件の明示トークン（fixture が衝突を置く）
        conflict = [str(t).strip() for t in list(p.get("sceneConflictTokens") or []) if str(t).strip()]
        if conflict and any(_norm(t) in _norm(scene) for t in conflict):
            return False, f"proposal[{i}] scene_conflict_errand"
        # salesTalk が別用件見出しを持つときだけ落とす
        sales = str(p.get("salesTalk") or "")
        sales_h = _bracket_heading(sales)
        if sales_h and catch_h and _norm(sales_h) != _norm(catch_h):
            if not any(_norm(t) in _norm(sales_h) for t in tokens):
                return False, f"proposal[{i}] sales_other_errand"
    return True, "ok"


def gate_s11_errand_before_taste(proposals: List[Dict[str, Any]]) -> Tuple[bool, str]:
    """S11: 商品名明示以外は用事の結果が先。味だけ先頭／見出しだけは不合格。"""
    for i, p in enumerate(proposals):
        tid = str(p.get("themeId") or "")
        if tid == "N1" or tid.startswith("N1"):
            continue
        catch = str(p.get("catch") or "").strip()
        h = _bracket_heading(catch)
        body = catch
        m = re.search(r"【[^】]+】(.*)$", catch)
        if m:
            body = (m.group(1) or "").strip()
        if h and zenkaku_len(body) < 4:
            return False, f"proposal[{i}] heading_only"
        head_check = h or body
        if _TASTE_FIRST.search(head_check):
            return False, f"proposal[{i}] taste_first_heading"
        if body and _TASTE_FIRST.search(body) and not h:
            return False, f"proposal[{i}] taste_first_body"
        # 見出しが味語のみ
        if h and _TASTE_FIRST.search(h) and zenkaku_len(re.sub(_TASTE_FIRST, "", h)) < 2:
            return False, f"proposal[{i}] taste_only_heading"
    return True, "ok"


def gate_s12_scene_photo_not_sales(
    proposals: List[Dict[str, Any]],
    *,
    is_food: bool,
) -> Tuple[bool, str]:
    """S12: 場面は写真指示。売文・ペルソナ属性禁止。食品は料理キーばらけ。"""
    dish_keys: List[str] = []
    for i, p in enumerate(proposals):
        sc = str(p.get("sceneHint") or "").strip()
        if not sc:
            return False, f"proposal[{i}] empty_sceneHint"
        for w in _SCENE_SALES:
            if w in sc:
                return False, f"proposal[{i}] scene_sales_word={w}"
        for w in _SCENE_PERSONA:
            if w in sc:
                return False, f"proposal[{i}] scene_persona_word={w}"
        has_photo = any(c in sc for c in _PHOTO_CUE)
        if not has_photo and zenkaku_len(sc) < 12:
            return False, f"proposal[{i}] scene_dish_name_only"
        dk = str(p.get("dishKey") or "").strip()
        if dk:
            dish_keys.append(dk)
    if is_food and len(proposals) >= 2 and dish_keys:
        if len(set(dish_keys)) < 2:
            return False, "food_dish_keys_not_varied"
    return True, "ok"


def gate_s13_no_entrance_gate(case: Dict[str, Any]) -> Tuple[bool, str]:
    """S13: テーマ入口ゲートで候補を落とさない。事実の門は別（捏造フラグ）。"""
    if bool(case.get("entrance_gate_applied")):
        return False, "entrance_gate_dropped_candidates"
    # 材料に無い事実をコピーに載せたまま＝事実の門が効いていない（参考検知）
    if bool(case.get("invented_fact_in_copy")) and not bool(case.get("fact_gate_blocked")):
        return False, "invented_fact_not_blocked"
    return True, "ok"


def gate_s14_image_large_and_small(case: Dict[str, Any]) -> Tuple[bool, str]:
    """S14: 画像の字は大＝catchBurn＋小＝salesBurn（無ければ salesTalk）。文字20%に両方。"""
    fields = case.get("image_text_fields")
    if fields is None:
        mode = str(case.get("image_text_mode") or "large_and_small")
        if mode in ("burn_only", "large_only", "catch_only"):
            return False, "sales_not_in_image_text"
        fields = ["catchBurn", "salesBurn"]
    norm_fields = [str(f).strip() for f in list(fields)]
    if "catchBurn" not in norm_fields:
        return False, "catchBurn_not_in_image_text"
    has_small = "salesBurn" in norm_fields or "salesTalk" in norm_fields
    if not has_small:
        return False, "sales_not_in_image_text"
    # 案に salesBurn があれば長さ門（長すぎる salesTalk の逃げ道）
    for i, p in enumerate(list(case.get("proposals") or [])):
        sb = str(p.get("salesBurn") or "").strip()
        st = str(p.get("salesTalk") or "").strip()
        use = sb or st
        if not use:
            return False, f"proposal[{i}] empty_sales_for_image"
        if sb and (zenkaku_len(sb) < 4 or zenkaku_len(sb) > 12):
            return False, f"proposal[{i}] salesBurn_len={zenkaku_len(sb)}"
        if not sb and zenkaku_len(st) > 12:
            return False, f"proposal[{i}] salesTalk_too_long_without_salesBurn"
    if "text20_counts_sales" in case and not bool(case.get("text20_counts_sales")):
        return False, "text20_ignores_sales"
    return True, "ok"


def run_all_gates(case: Dict[str, Any]) -> Dict[str, Any]:
    """1ケース分。case は fixture 辞書。"""
    props = list(case.get("proposals") or [])
    results = {}
    results["S1"] = gate_s1_mix_both(
        props,
        is_mix=bool(case.get("is_mix")),
        mix_marks=list(case.get("mix_marks") or []),
    )
    results["S2"] = gate_s2_catch_sales_roles(props)
    results["S3"] = gate_s3_persona(
        props, ui_can_open_detail=bool(case.get("ui_can_open_detail"))
    )
    results["S4"] = gate_s4_catch_burn(props)
    results["S5"] = gate_s5_persona_in_gen(
        polish_prompt=str(case.get("polish_prompt") or ""),
        persona_short=str(case.get("persona_short") or (props[0].get("personaShort") if props else "")),
    )
    results["S6"] = gate_s6_remake_checked_only(
        before=list(case.get("remake_before") or props),
        after=list(case.get("remake_after") or props),
        checked_keys=list(case.get("checked_keys") or []),
    )
    results["S7"] = gate_s7_scene_hint(
        props, ui_shows_scene_editable=bool(case.get("ui_shows_scene_editable"))
    )
    results["S8"] = gate_s8_semantic_burn(props)
    results["S9"] = gate_s9_distinct_token_policy(
        is_mix=bool(case.get("is_mix")),
        sanitize_allows_single_distinct_for_mix=bool(
            case.get("sanitize_allows_single_distinct_for_mix")
            if "sanitize_allows_single_distinct_for_mix" in case
            else case.get("sanitize_allows_single_fish_for_mix")
        ),
    )
    results["S10"] = gate_s10_same_errand(props)
    results["S11"] = gate_s11_errand_before_taste(props)
    results["S12"] = gate_s12_scene_photo_not_sales(
        props, is_food=bool(case.get("is_food", True))
    )
    results["S13"] = gate_s13_no_entrance_gate(case)
    results["S14"] = gate_s14_image_large_and_small(case)
    # 適用ゲートの選択（単品は S1/S9 の mix 側を緩和済み）
    apply = list(case.get("apply_gates") or list(_DEFAULT_GATES))
    fails = []
    for gid in apply:
        ok, reason = results[gid]
        if not ok:
            fails.append(f"{gid}:{reason}")
    return {
        "id": case.get("id"),
        "expect_pass": bool(case.get("expect_pass", True)),
        "pass": len(fails) == 0,
        "fails": fails,
        "details": {k: {"ok": v[0], "reason": v[1]} for k, v in results.items()},
    }
