# -*- coding: utf-8 -*-
"""B-⑤ テーマ案の文・門・ミックス契約（P2 第1スライス）。GASと同趣旨。画像・Geminiなし。

正本: docs/org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md §4.5.8
対象: T2 / T3 / T4 / T5 / T8 / T14 / T15（第1スライス）
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence, Tuple

# T8: 自動リトライ上限（人手のレ点外し再作成は別枠）
AUTO_RETRY_MAX = 3

# T2: 提案で落とす語（材料にあっても見出し・補足に載せない／材料に無い煽り）
BANNED_CLAIM_WORDS = (
    "効能",
    "健康",
    "栄養",
    "成長",
    "糖質",
    "ダイエット",
    "血圧",
    "免疫",
    "高タンパク",
    "ヘルシー",
    "お得",
    "最安",
    "激安",
    "コスパ最強",
    "最強",
    "受賞",
    "金賞",
    "No.1",
    "ナンバーワン",
    "累計",
)

# T15: catch／subs に手順・体験の書き方を載せない
PROCEDURE_WORDS = (
    "開けて",
    "のせて",
    "並べて",
    "混ぜて",
    "手順",
    "まず",
    "次に",
    "してから",
)


def proposal_text_blob(pr: Optional[Dict[str, Any]]) -> str:
    if not isinstance(pr, dict):
        return ""
    parts = [str(pr.get("overview") or ""), str(pr.get("catch") or "")]
    for s in pr.get("subs") or []:
        parts.append(str(s or ""))
    return "\n".join(parts)


def find_banned_claims(text: str) -> List[str]:
    t = str(text or "")
    hit = []
    for w in BANNED_CLAIM_WORDS:
        if w and w in t:
            hit.append(w)
    return hit


def find_procedure_words(text: str) -> List[str]:
    t = str(text or "")
    hit = []
    for w in PROCEDURE_WORDS:
        if w and w in t:
            hit.append(w)
    return hit


def has_bracket_markers(text: str) -> bool:
    t = str(text or "")
    return ("【" in t) or ("】" in t)


def catch_equals_any_sub(pr: Dict[str, Any]) -> bool:
    catch = str(pr.get("catch") or "").strip()
    if not catch:
        return False
    for s in pr.get("subs") or []:
        if str(s or "").strip() == catch:
            return True
    return False


def _norm_token(s: str) -> str:
    return "".join(str(s or "").split())


def distinctive_chunks(text: str, *, min_len: int = 4) -> List[str]:
    """短い連続かな漢字カタカナ英数をチャンク化（使い回し検出用）。"""
    t = _norm_token(text)
    if len(t) < min_len:
        return [t] if t else []
    out: List[str] = []
    # スライド窓で代表チャンク（過検出を抑えるため先頭寄り＋全体）
    if len(t) <= 12:
        out.append(t)
    else:
        out.append(t[:8])
        out.append(t[-8:])
        mid = len(t) // 2
        out.append(t[mid - 4 : mid + 4])
    # 重複除去
    seen = set()
    uniq = []
    for c in out:
        if len(c) >= min_len and c not in seen:
            seen.add(c)
            uniq.append(c)
    return uniq


def reused_chunks_across_themes(
    proposals: Sequence[Dict[str, Any]],
) -> List[Tuple[int, int, str]]:
    """
    T3: 同親の他テーマ catch／subs で同じ語塊を使い回していないか。
    catch と各 sub を別文としてチャンク化（結合正規化で別案化しない）。
    戻り: (先index, 後index, chunk)
    """
    indexed: List[Tuple[int, str]] = []
    for i, pr in enumerate(proposals or []):
        lines = [str(pr.get("catch") or "")] + [str(s or "") for s in (pr.get("subs") or [])]
        for line in lines:
            line = str(line or "").strip()
            if not line:
                continue
            for ch in distinctive_chunks(line):
                indexed.append((i, ch))
    hits: List[Tuple[int, int, str]] = []
    for a in range(len(indexed)):
        ia, ca = indexed[a]
        for b in range(a + 1, len(indexed)):
            ib, cb = indexed[b]
            if ia == ib:
                continue
            if ca == cb:
                hits.append((ia, ib, ca))
    # 重複 (i,j,chunk) を1つに
    seen = set()
    uniq = []
    for h in hits:
        key = (min(h[0], h[1]), max(h[0], h[1]), h[2])
        if key not in seen:
            seen.add(key)
            uniq.append((key[0], key[1], key[2]))
    return uniq


# 構成区別語の表記ゆれ（商品固定にしない。ひらがな・カタカナ・常用漢字）
_MARK_ALIAS_GROUPS = (
    ("かつお", "カツオ", "鰹"),
    ("まぐろ", "マグロ", "鮪"),
    ("ぶり", "ブリ", "鰤"),
    ("さば", "サバ", "鯖"),
    ("いわし", "イワシ", "鰯"),
    ("チーズ", "ちーず"),
)


def mark_match_needles(mark: str) -> List[str]:
    """1つの区別語に対し、文中ヒットとみなす表記ゆれ一覧。"""
    m = str(mark or "").strip()
    if not m:
        return []
    out = [m]
    for group in _MARK_ALIAS_GROUPS:
        if m in group:
            for g in group:
                if g not in out:
                    out.append(g)
            break
    return out


def mark_appears_in_text(text: str, mark: str) -> bool:
    t = str(text or "")
    for needle in mark_match_needles(mark):
        if needle and needle in t:
            return True
    return False


def mix_marks_mentioned(text: str, marks: Sequence[str]) -> List[str]:
    return [m for m in marks if m and mark_appears_in_text(text, str(m))]


def is_mix_one_sided(text: str, marks: Sequence[str]) -> bool:
    """
    T4: ミックスで構成区別語が2つ以上あるとき、文面が片方の語だけに寄っていないか。
    表記ゆれ（かつお／鰹 等）は同一視。どちらも無し／両方あり → False。片方だけ → True。
    """
    marks = [str(m).strip() for m in (marks or []) if str(m).strip()]
    if len(marks) < 2:
        return False
    hit = mix_marks_mentioned(text, marks)
    return len(hit) == 1


def gate_proposal(
    pr: Dict[str, Any],
    *,
    mix_marks: Optional[Sequence[str]] = None,
    check_overview: bool = True,
) -> List[str]:
    """1案の門。空なら合格。"""
    errs: List[str] = []
    catch = str(pr.get("catch") or "")
    subs_blob = "\n".join(str(s or "") for s in (pr.get("subs") or []))
    ov = str(pr.get("overview") or "") if check_overview else ""
    surface = "\n".join([catch, subs_blob])  # T5/T15 は catch/subs 中心
    all_text = "\n".join([ov, surface])

    banned = find_banned_claims(all_text)
    if banned:
        errs.append("T2:" + ",".join(banned))

    if has_bracket_markers(surface):
        errs.append("T5:bracket")
    if catch_equals_any_sub(pr):
        errs.append("T5:catch=subs")

    proc = find_procedure_words(surface)
    if proc:
        errs.append("T15:" + ",".join(proc))

    # T4: 概要も含めて片寄り判定（キャッチだけ片方・概要で両方、を過剰落とししない）
    if mix_marks and is_mix_one_sided(all_text, mix_marks):
        errs.append("T4:one_sided")

    return errs


def gate_proposals(
    proposals: Sequence[Dict[str, Any]],
    *,
    mix_marks: Optional[Sequence[str]] = None,
) -> Dict[str, Any]:
    """親単位の門。T3は横断。"""
    per: List[Dict[str, Any]] = []
    fail_indexes: List[int] = []
    for i, pr in enumerate(proposals or []):
        errs = gate_proposal(pr, mix_marks=mix_marks)
        row = {"index": i, "themeId": (pr or {}).get("themeId"), "errors": errs}
        per.append(row)
        if errs:
            fail_indexes.append(i)

    reuse = reused_chunks_across_themes(proposals)
    t3_indexes = sorted({i for pair in reuse for i in (pair[0], pair[1])})
    for ix in t3_indexes:
        if ix not in fail_indexes:
            fail_indexes.append(ix)
        # annotate
        for row in per:
            if row["index"] == ix:
                chunks = [c for a, b, c in reuse if a == ix or b == ix]
                if chunks:
                    row["errors"] = list(row["errors"]) + ["T3:" + chunks[0]]

    fail_indexes = sorted(set(fail_indexes))
    return {
        "ok": len(fail_indexes) == 0,
        "failIndexes": fail_indexes,
        "perProposal": per,
        "reusePairs": [{"a": a, "b": b, "chunk": c} for a, b, c in reuse],
    }


def should_auto_retry(attempt: int) -> bool:
    """T8: attempt は 0始まり。3回まで（0,1,2）。"""
    return 0 <= int(attempt) < AUTO_RETRY_MAX


def next_retry_indexes(
    gate_result: Dict[str, Any],
    *,
    attempt: int,
) -> List[int]:
    """自動リトライ対象。上限超過なら空。"""
    if not should_auto_retry(attempt):
        return []
    return list(gate_result.get("failIndexes") or [])
