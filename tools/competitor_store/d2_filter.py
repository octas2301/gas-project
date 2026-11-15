# -*- coding: utf-8 -*-
"""D2 機械絞込ロジック（GAS prD2Classify_ と同型）。"""
from __future__ import annotations

import re
from typing import Any

from maker_alias import JITTAI_CHECK, JITTAI_EXCLUDE, JITTAI_MAKER, JITTAI_UNCHECKED

FURUSATO = re.compile(r"ふるさと|返礼品")
SELECT = re.compile(
    r"よりどり|種類が選べ|選べるセット|から選択|種類から選|お好みで選|味が選べ"
)
USED = re.compile(r"中古品|中古")
BOOK_CATS = ("本", "洋書", "ミュージック", "DVD", "Kindleストア", "クラシック", "デジタルミュージック")
OVERSIZE_TIER = re.compile(r"^大型|^特大")
WEIGHT_MAX_DEFAULT_G = 5000.0


def is_book(cat: str) -> bool:
    c = str(cat or "").strip()
    if not c:
        return False
    return any(c == b or c.startswith(b) for b in BOOK_CATS)


def parse_weight_g(raw: Any) -> float:
    s = str(raw or "").strip().replace(",", "")
    if not s:
        return float("nan")
    m = re.match(r"^([0-9]+(?:\.[0-9]+)?)\s*(kg|g)?$", s, re.I)
    if m:
        v = float(m.group(1))
        if v <= 0:
            return float("nan")
        if str(m.group(2) or "").lower() == "kg":
            return v * 1000.0
        return v
    s2 = re.sub(r"g$", "", s, flags=re.I).strip()
    try:
        v = float(s2)
        return v if v > 0 else float("nan")
    except ValueError:
        return float("nan")


def is_oversize_tier(tier: str) -> bool:
    return bool(OVERSIZE_TIER.match(str(tier or "").strip()))


def profile_weight_max_g(prof: dict | None) -> float:
    if not prof:
        return WEIGHT_MAX_DEFAULT_G
    try:
        n = float(prof.get("重量上限g") or 0)
    except (TypeError, ValueError):
        n = 0.0
    return n if n > 0 else WEIGHT_MAX_DEFAULT_G


def profile_oversize_drop(prof: dict | None) -> bool:
    if not prof:
        return True
    v = str(prof.get("大型") or "落とす").strip()
    return v not in ("落とさない", "いいえ", "0", "OFF")


def profile_for_category(cat: str, profs: list[dict]) -> dict:
    blob = " ".join(p.strip() for p in re.split(r"[›>]", str(cat or "")) if p.strip())
    common: dict | None = None
    for m in profs:
        c = str(m.get("カテゴリ") or "")
        if c == "共通":
            common = m
            continue
        needles = str(m.get("マッチ語") or "").split("|")
        for n in needles:
            if n and n in blob:
                return m
        if c and c in blob:
            return m
        if c == "食品" and blob.startswith("食品"):
            return m
    if common:
        return common
    return profs[0] if profs else {}


def classify(
    asin: str,
    maker: str,
    title: str,
    cat: str,
    cand: dict | None,
    jitta: str,
    *,
    qs_tier: str = "",
    weight: str = "",
    prof: dict | None = None,
) -> tuple[str, str]:
    if jitta == JITTAI_EXCLUDE:
        return "見送り", "非仕入先"
    if jitta == JITTAI_UNCHECKED:
        return "見送り", "未調査"
    if jitta == JITTAI_CHECK:
        return "見送り", "要確認"
    if not cand:
        return "見送り", "①候補なし"
    gate = str(cand.get("門結果") or "").strip()
    if gate == "落ち":
        return "見送り", str(cand.get("門理由") or "門落ち")[:80]
    if gate != "通過":
        return "見送り", "門未通過"
    t = str(title or "")
    if FURUSATO.search(t):
        return "見送り", "ふるさと"
    if SELECT.search(t):
        return "見送り", "選択式"
    if USED.search(t):
        return "見送り", "中古"
    if is_book(cat):
        return "見送り", "本系カテゴリ"
    if profile_oversize_drop(prof) and is_oversize_tier(qs_tier):
        return "見送り", "大型サイズ"
    wg = parse_weight_g(weight)
    wmax = profile_weight_max_g(prof)
    if wg == wg and wg >= wmax:  # nan check
        return "見送り", "重量>=" + str(int(wmax) if wmax == int(wmax) else wmax) + "g"
    return "見積候補", ""
