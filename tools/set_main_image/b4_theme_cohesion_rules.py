# -*- coding: utf-8 -*-
"""B-⑤ テーマ案のつなぎ・訴求軸・買う人（P2 第2スライス）。GASと同趣旨。画像・Geminiなし。

正本: docs/org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md §4.5.9
対象: T6 / T7 / T13
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Sequence, Set

# T13: 対面手渡し劇場（ネット注文者が買う前提とずれる）
FACE_HANDOFF_WORDS = (
    "手渡し",
    "手渡す",
    "直接渡",
    "玄関で渡",
    "相手の手に",
    "渡す瞬間",
    "渡しながら",
)

_STOP = {
    "する",
    "して",
    "した",
    "いる",
    "ある",
    "なる",
    "れる",
    "られる",
    "こと",
    "もの",
    "ため",
    "よう",
    "これ",
    "それ",
    "ここ",
    "とき",
    "とき",
    "など",
    "また",
    "から",
    "まで",
    "より",
    "への",
    "との",
    "では",
    "には",
    "のは",
    "のが",
    "です",
    "ます",
    "たい",
    "ない",
}


def normalize_axis(s: str) -> str:
    t = str(s or "")
    t = re.sub(r"\s+", "", t)
    t = t.replace("　", "")
    return t.strip().lower()


def content_tokens(text: str, *, min_len: int = 2) -> List[str]:
    """ひらがな・カタカナ・漢字・英数の連続をトークン化。短すぎ・ストップは除く。"""
    raw = re.findall(r"[一-龥ぁ-んァ-ヶーa-zA-Z0-9]{%d,}" % min_len, str(text or ""))
    out: List[str] = []
    seen: Set[str] = set()
    for w in raw:
        if w in _STOP:
            continue
        if w in seen:
            continue
        seen.add(w)
        out.append(w)
    return out


def axis_tokens(axis: str) -> List[str]:
    """軸を内容語に分解（助詞で分割。一件専用の語は持たない）。"""
    base = content_tokens(axis, min_len=2)
    out: List[str] = []
    seen: Set[str] = set()
    for t in base:
        parts = re.split(r"[のをおにへとはもで・×x/／]", t)
        for p in parts + [t]:
            p = str(p or "").strip()
            if len(p) < 2 or p in _STOP or p in seen:
                continue
            seen.add(p)
            out.append(p)
    return out


def field_blob(pr: Dict[str, Any]) -> Dict[str, str]:
    subs = pr.get("subs") or []
    first_sub = str(subs[0] or "") if subs else ""
    return {
        "overview": str(pr.get("overview") or ""),
        "catch": str(pr.get("catch") or ""),
        "sub0": first_sub,
        "axis": str(pr.get("appealAxis") or pr.get("errand") or ""),
        "dish": str(pr.get("dishName") or ""),
    }


def cohesion_hit_count(pr: Dict[str, Any]) -> int:
    """
    T6: 軸の内容語が overview / catch / sub0 の何欄に載るか。
    軸が空なら 0。
    """
    fields = field_blob(pr)
    toks = axis_tokens(fields["axis"])
    if not toks:
        return 0
    hit = 0
    for key in ("overview", "catch", "sub0"):
        blob = fields[key]
        if any(t in blob for t in toks):
            hit += 1
    return hit


def cohesion_ok(pr: Dict[str, Any]) -> bool:
    """T6合格: 軸あり、かつ2欄以上に軸語。"""
    fields = field_blob(pr)
    if not normalize_axis(fields["axis"]):
        return False
    return cohesion_hit_count(pr) >= 2


def dish_scene_ok(pr: Dict[str, Any]) -> bool:
    """
    食べ方があるとき、overviewかcatchに食べる／食卓系が1回。
    dishNameが空なら常にOK。
    """
    dish = str(pr.get("dishName") or "").strip()
    if not dish:
        return True
    blob = str(pr.get("overview") or "") + "\n" + str(pr.get("catch") or "")
    cues = ("食", "皿", "ご飯", "丼", "麺", "鍋", "おかず", "つまみ", "晩酌", "朝食", "お弁当")
    if dish and dish in blob:
        return True
    return any(c in blob for c in cues)


def find_face_handoff(text: str) -> List[str]:
    t = str(text or "")
    return [w for w in FACE_HANDOFF_WORDS if w and w in t]


def gate_proposal_cohesion(pr: Dict[str, Any]) -> List[str]:
    """1案。T6/T13。"""
    errs: List[str] = []
    fields = field_blob(pr)
    if not normalize_axis(fields["axis"]):
        errs.append("T6:no_axis")
    elif not cohesion_ok(pr):
        errs.append("T6:cohesion")
    if not dish_scene_ok(pr):
        errs.append("T6:dish_scene")
    surface = fields["catch"] + "\n" + fields["sub0"]
    hand = find_face_handoff(surface)
    if hand:
        errs.append("T13:" + ",".join(hand))
    return errs


def duplicate_axis_pairs(proposals: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """T7: 正規化軸の重複ペア。"""
    indexed: List[tuple] = []
    for i, pr in enumerate(proposals or []):
        ax = normalize_axis(str((pr or {}).get("appealAxis") or (pr or {}).get("errand") or ""))
        if ax:
            indexed.append((i, ax, (pr or {}).get("themeId")))
    pairs = []
    seen = set()
    for a in range(len(indexed)):
        for b in range(a + 1, len(indexed)):
            if indexed[a][1] != indexed[b][1]:
                continue
            key = (indexed[a][1], min(indexed[a][0], indexed[b][0]), max(indexed[a][0], indexed[b][0]))
            if key in seen:
                continue
            seen.add(key)
            pairs.append(
                {
                    "a": indexed[a][0],
                    "b": indexed[b][0],
                    "axis": indexed[a][1],
                    "themeA": indexed[a][2],
                    "themeB": indexed[b][2],
                }
            )
    return pairs


def gate_proposals_cohesion(proposals: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    per: List[Dict[str, Any]] = []
    fail: Set[int] = set()
    for i, pr in enumerate(proposals or []):
        errs = gate_proposal_cohesion(pr or {})
        per.append({"index": i, "themeId": (pr or {}).get("themeId"), "errors": errs})
        if errs:
            fail.add(i)
    dups = duplicate_axis_pairs(proposals)
    for d in dups:
        fail.add(d["a"])
        fail.add(d["b"])
        for row in per:
            if row["index"] in (d["a"], d["b"]):
                row["errors"] = list(row["errors"]) + ["T7:dup_axis"]
    fail_indexes = sorted(fail)
    return {
        "ok": len(fail_indexes) == 0,
        "failIndexes": fail_indexes,
        "perProposal": per,
        "dupAxes": dups,
    }


def remake_should_change_axis(note: str, old_axis: str) -> bool:
    """再作成要望に軸変更のヒントが無くても、門メモ用。今は note 非空なら True（必須は既存）。"""
    return bool(str(note or "").strip()) and bool(str(old_axis or "").strip())
