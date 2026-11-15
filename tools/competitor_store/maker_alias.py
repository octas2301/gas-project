# -*- coding: utf-8 -*-
"""メーカー別名・実態ラベル。サマリ寄せ用。"""
from __future__ import annotations

# 実態（確定名）
JITTAI_MAKER = "製造メーカー"
JITTAI_BRAND = "入口ブランド"
JITTAI_CHECK = "要確認"
JITTAI_UNCHECKED = "未調査"
JITTAI_EXCLUDE = "非仕入先"
JITTAI_VALUES = (JITTAI_MAKER, JITTAI_BRAND, JITTAI_CHECK, JITTAI_EXCLUDE)
JITTAI_SUMMARY_VALUES = JITTAI_VALUES + (JITTAI_UNCHECKED,)

# 入口語（ブランド等） → 親メーカー名
MAKER_ALIASES = {
    "AGF": "味の素",
    "エージーエフ": "味の素",
    "AGF(エージーエフ)": "味の素",
    "agf": "味の素",
    "いしもと": "いし本食品工業",
    "いし本食品工業": "いし本食品工業",
    "S&B": "エスビー食品",
    "エスビー食品": "エスビー食品",
    "ヱスビー": "エスビー食品",
    "Nagatanien": "永谷園",
    "Sato": "サトウ食品",
}

# 親メーカー向けの入口語ヒット種（手動／小テスト）
ENTRY_HIT_SEED = {
    "味の素": "AGF/エージーエフ page0=20満杯(2026-08-17)",
    "エバラ": "エバラ page0=20満杯(2026-08-17)",
    "エースコック": "エースコック page0=20満杯(2026-08-17)",
    "エスビー食品": "S&B/エスビー食品 page0=20満杯(2026-08-17)",
    "いし本食品工業": "いしもと/いし本食品工業 page0=20満杯(2026-08-17)",
}

# 倉庫ブランド名 → page0結果（dry_maker_brand_page0 が追記）
BRAND_PAGE0_HITS: dict[str, str] = {
    "AGF(エージーエフ)": "page0=20満杯(2026-08-17)",
    "Granja San Francis": "page0=7(2026-08-17)",
    "Nagatanien": "page0=20満杯(2026-08-17)",
    "Sato": "page0=20満杯(2026-08-17)",
    "Tiptree": "page0=20満杯(2026-08-17)",
    "はっとり海苔": "page0=20満杯(2026-08-17)",
    "わが街とくさんネット": "page0=20満杯(2026-08-17)",
    "エバラ": "page0=20満杯(2026-08-17)",
    "オキハム": "page0=20満杯(2026-08-17)",
    "フロム四十五（from45）": "page0=20(2026-08-17)",
    "ムソー": "page0=20満杯(2026-08-17)",
    "リフココ": "page0=20満杯(2026-08-17)",
    "リフ工房": "page0=20満杯(2026-08-17)",
    "五木食品": "page0=20満杯(2026-08-17)",
    "天然生活": "page0=20満杯(2026-08-17)",
    "尾西食品": "page0=20満杯(2026-08-17)",
    "扶桑社": "page0=20満杯(2026-08-17)",
    "石原水産": "page0=20満杯(2026-08-17)",
    "黒糖本舗垣乃花": "page0=20満杯(2026-08-17)",
    "ＫＥＹＲＯＯＭ": "page0=20満杯(2026-08-17)",
}

# 実態シード（最小）。非仕入先のみ。他は AI または要確認。
MAKER_JITTAI: dict[str, str] = {
    "わが街とくさんネット": JITTAI_EXCLUDE,
    "扶桑社": JITTAI_EXCLUDE,
    "天然生活": JITTAI_EXCLUDE,
    "ノーブランド品": JITTAI_EXCLUDE,
    "Generic": JITTAI_EXCLUDE,
    "未確定": JITTAI_EXCLUDE,
    "リフココ": JITTAI_EXCLUDE,
    "リフ工房": JITTAI_EXCLUDE,
}


def resolve_alias(name: str) -> str:
    t = (name or "").strip()
    if not t:
        return ""
    if t in MAKER_ALIASES:
        return MAKER_ALIASES[t]
    low = t.lower()
    for k, v in MAKER_ALIASES.items():
        if k.lower() == low:
            return v
    return t


def normalize_maker_key(name: str) -> str:
    """法人格ゆれを除いた照合キー（GAS cjNormalizeMakerKey_ / prNormalizeMakerKey_ と同期）。"""
    import re

    s = (name or "").strip()
    if not s:
        return ""
    s = re.sub(
        r"株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|一般財団法人",
        "",
        s,
    )
    s = re.sub(r"[（(]株[）)]", "", s)
    s = re.sub(r"[（(]有[）)]", "", s)
    s = s.replace("㈱", "").replace("㈲", "")
    s = re.sub(r"\s+", "", s)
    return s


def entry_preferred(maker: str) -> str:
    """入口推奨。実態と連動。"""
    j = jitta_for_maker(maker)
    if j == JITTAI_BRAND:
        return "ブランド"
    if j == JITTAI_MAKER:
        return "メーカー"
    if j == JITTAI_EXCLUDE:
        return "非仕入先"
    if j == JITTAI_CHECK:
        return "要確認"
    if j == JITTAI_UNCHECKED:
        return "未調査"
    return "未設定"


def jitta_for_maker(maker: str, override: str = "") -> str:
    """実態。MAKER_JITTAI（コード）優先。未登録のみシート override。"""
    m = (maker or "").strip()
    if m in MAKER_JITTAI:
        return MAKER_JITTAI[m]
    o = (override or "").strip()
    if o in JITTAI_SUMMARY_VALUES:
        return o
    for k, v in MAKER_ALIASES.items():
        if v == m and k != m and k.lower() != m.lower():
            return JITTAI_BRAND
    return JITTAI_UNCHECKED


def amazon_search_url(maker: str) -> str:
    """Amazon.co.jp 検索（メーカー名入力済み）。実態ck 列用。"""
    from urllib.parse import quote

    m = (maker or "").strip()
    if not m:
        return ""
    return "https://www.amazon.co.jp/s?k=" + quote(m)


def hit_text_for_maker(maker: str) -> str:
    m = (maker or "").strip()
    parts = []
    if m in ENTRY_HIT_SEED:
        parts.append(ENTRY_HIT_SEED[m])
    for brand, text in sorted(BRAND_PAGE0_HITS.items()):
        if resolve_alias(brand) == m or brand == m:
            parts.append("%s/%s %s" % (brand, m, text))
    seen = set()
    uniq = []
    for p in parts:
        if p not in seen:
            seen.add(p)
            uniq.append(p)
    return "; ".join(uniq) if uniq else "未計測"
