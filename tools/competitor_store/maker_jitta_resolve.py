# -*- coding: utf-8 -*-
"""メーカー実態リゾルバ（dry/本番共通ロジック）。

正本: docs/org/B_MAKER_JITTA_CLASSIFICATION.md
"""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone

from maker_alias import (
    JITTAI_BRAND,
    JITTAI_CHECK,
    JITTAI_EXCLUDE,
    JITTAI_MAKER,
    JITTAI_UNCHECKED,
    JITTAI_VALUES,
    entry_preferred,
    jitta_for_maker,
)
from schema import AMAZON_SELLER_CAT_COLS

GYOSHU_UNKNOWN = "不明"
GYOSHU_INFERRED_SUFFIX = "（推定）"

# fixture = AI/Web 調査結果（本番では Gemini/OpenAI が同形 JSON を返す想定）
AI_FIXTURES: dict[str, dict] = {
    "永谷園": {
        "jitta_candidate": JITTAI_MAKER,
        "entry_preferred": "メーカー",
        "entry_needle": "",
        "confidence": "high",
        "source": "official",
        "evidence_urls": [
            "https://www.nagatanien.co.jp/company/profile.html",
            "https://www.nagatanien-hd.co.jp/",
        ],
        "gyoshu": "即席食品・中食（食品製造）",
        "gyoshu_source": "official",
        "amazon_root_categories": "食品・飲料・お酒",
        "amazon_browse_shelves": "お茶漬け・ふりかけ|即席みそ汁|惣菜・食材・調味料",
        "reason": "会社概要に製造販売と明記。棚名も永谷園。",
    },
    "味の素": {
        "jitta_candidate": JITTAI_BRAND,
        "entry_preferred": "ブランド",
        "entry_needle": "AGF",
        "confidence": "high",
        "source": "official",
        "evidence_urls": ["https://www.ajinomoto.co.jp/"],
        "gyoshu": "調味料・加工食品",
        "gyoshu_source": "official",
        "amazon_root_categories": "食品・飲料・お酒",
        "amazon_browse_shelves": "調味料|コーヒー・紅茶|冷凍食品|スープ・みそ汁",
        "reason": "親は味の素。棚では AGF・ほんだし等の別ブランド名。",
    },
    "あさくま": {
        "jitta_candidate": JITTAI_CHECK,
        "entry_preferred": "要確認",
        "entry_needle": "",
        "confidence": "medium",
        "source": "official",
        "evidence_urls": ["https://www.asakuma.co.jp/company.html"],
        "gyoshu": "外食・飲食",
        "gyoshu_source": "official",
        "amazon_root_categories": "食品・飲料・お酒",
        "amazon_browse_shelves": "調味料・ソース",
        "reason": "本業はレストラン。ソース販売はあるが卸可否は要確認。",
    },
    "ＫＥＹＲＯＯＭ": {
        "jitta_candidate": JITTAI_CHECK,
        "entry_preferred": "要確認",
        "entry_needle": "",
        "confidence": "low",
        "source": "gemini",
        "evidence_urls": [],
        "gyoshu": "",
        "gyoshu_source": "none",
        "amazon_root_categories": "",
        "amazon_browse_shelves": "",
        "reason": "東証・日経・公式HPから業種を特定できず。要確認。",
    },
}

MAKER_MASTER_FIELDS = (
    "実態候補",
    "実態信頼度",
    "実態採取元",
    "実態根拠URL",
    "実態取得日時",
    "業種",
    "Amazonルートカテゴリ",
    "AmazonBrowse棚",
    "入口語候補",
    "判定メモ",
)

SUMMARY_MIRROR_FIELDS = (
    "実態採取元",
    "実態信頼度",
    "実態根拠URL",
    "業種",
    "Amazonルートカテゴリ",
    "AmazonBrowse棚",
)

AMAZON_ROOT_SET = frozenset(AMAZON_SELLER_CAT_COLS)


def _now_utc() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def lookup_ai_fixture(maker: str) -> dict | None:
    key = (maker or "").strip()
    if not key:
        return None
    raw = AI_FIXTURES.get(key)
    return deepcopy(raw) if raw else None


def evidence_urls_text(urls: list[str] | str) -> str:
    if isinstance(urls, str):
        return urls.strip()
    return "|".join(u.strip() for u in urls if str(u).strip())


def split_pipe_values(text: str) -> list[str]:
    return [p.strip() for p in str(text or "").split("|") if p.strip()]


def join_pipe_values(values: list[str]) -> str:
    seen: set[str] = set()
    out: list[str] = []
    for v in values:
        t = (v or "").strip()
        if not t or t in seen:
            continue
        seen.add(t)
        out.append(t)
    return "|".join(out)


def normalize_gyoshu(resolved: dict) -> str:
    """業種の正規化。ソース無しは空（シート上は人が後から埋める）。"""
    raw = str(resolved.get("gyoshu") or "").strip()
    src = str(resolved.get("gyoshu_source") or "").strip()
    if raw:
        if src in ("gemini", "openai", "product_inferred") and GYOSHU_INFERRED_SUFFIX not in raw:
            return raw + GYOSHU_INFERRED_SUFFIX
        return raw
    if src == "none":
        return GYOSHU_UNKNOWN
    return ""


def normalize_amazon_root_categories(text: str) -> str:
    """ルートはセラー台帳と同じ語彙のみ。未知語は除外。"""
    kept = [v for v in split_pipe_values(text) if v in AMAZON_ROOT_SET]
    return join_pipe_values(kept)


def resolved_to_master_patch(resolved: dict, *, fetched_at: str = "") -> dict:
    """マスタ行へ書くパッチ。"""
    at = fetched_at or _now_utc()
    return {
        "実態候補": resolved.get("jitta_candidate", ""),
        "実態信頼度": resolved.get("confidence", ""),
        "実態採取元": resolved.get("source", ""),
        "実態根拠URL": evidence_urls_text(resolved.get("evidence_urls") or []),
        "実態取得日時": at,
        "業種": normalize_gyoshu(resolved),
        "Amazonルートカテゴリ": normalize_amazon_root_categories(
            resolved.get("amazon_root_categories", "")
        ),
        "AmazonBrowse棚": join_pipe_values(split_pipe_values(resolved.get("amazon_browse_shelves", ""))),
        "入口語候補": resolved.get("entry_needle", ""),
        "判定メモ": resolved.get("reason", ""),
    }


def mirror_master_to_summary(master_row: dict) -> dict:
    """サマリ末尾ミラー列。"""
    return {k: master_row.get(k, "") for k in SUMMARY_MIRROR_FIELDS}


def should_auto_confirm_jitta(summary_row: dict, resolved: dict) -> bool:
    """high かつ要確認以外。非仕入先のみ自動確定しない。"""
    human = str(summary_row.get("実態") or "").strip()
    if human == JITTAI_EXCLUDE:
        return False
    cand = str(resolved.get("jitta_candidate") or "").strip()
    conf = str(resolved.get("confidence") or "").strip()
    return conf == "high" and cand in JITTAI_VALUES and cand != JITTAI_CHECK


def _promote_summary_jitta_after_ai(summary_row: dict, resolved: dict) -> str:
    """AI調査後にサマリ.実態へ候補を反映（未調査・レガシー要確認のみ）。"""
    human = str(summary_row.get("実態") or "").strip()
    cand = str(resolved.get("jitta_candidate") or "").strip()
    conf = str(resolved.get("confidence") or "").strip()
    if not cand or not conf:
        return human or jitta_for_maker(str(summary_row.get("メーカー") or ""))
    if human in ("", JITTAI_UNCHECKED) or human == JITTAI_CHECK:
        return cand
    return human


def summary_ai_confirmed_(summary_row: dict) -> bool:
    return str(summary_row.get("実態信頼度") or "").strip() == "high"


def resolve_maker_jitta(
    maker: str,
    *,
    summary_row: dict | None = None,
    master_row: dict | None = None,
    use_fixture: bool = True,
) -> dict:
    """1メーカー分のリゾルバ結果（副作用なし）。"""
    name = (maker or "").strip()
    summary_row = summary_row or {}
    master_row = master_row or {}

    human_jitta = str(summary_row.get("実態") or "").strip()
    if summary_ai_confirmed_(summary_row):
        return {
            "maker": name,
            "skipped": True,
            "skip_reason": "ai_confirmed",
            "jitta": human_jitta,
            "entry_preferred": entry_preferred(name),
        }

    if master_row and str(master_row.get("実態再採取しない") or "").upper() in ("TRUE", "1"):
        patch = {k: master_row.get(k, "") for k in MAKER_MASTER_FIELDS if k in master_row}
        mirror = mirror_master_to_summary(master_row)
        jitta = str(master_row.get("実態候補") or jitta_for_maker(name))
        return {
            "maker": name,
            "skipped": True,
            "skip_reason": "cache_no_resample",
            "master_patch": patch,
            "summary_mirror": mirror,
            "jitta": jitta,
            "entry_preferred": entry_preferred(name),
            "auto_confirm": False,
        }

    resolved = lookup_ai_fixture(name) if use_fixture else None
    if not resolved:
        return {
            "maker": name,
            "skipped": False,
            "resolved": None,
            "jitta": jitta_for_maker(name),
            "entry_preferred": entry_preferred(name),
            "auto_confirm": False,
        }

    master_patch = resolved_to_master_patch(resolved)
    summary_mirror = mirror_master_to_summary(master_patch)
    auto = should_auto_confirm_jitta(summary_row, resolved)
    jitta = resolved["jitta_candidate"] if auto else _promote_summary_jitta_after_ai(summary_row, resolved)

    return {
        "maker": name,
        "skipped": False,
        "resolved": resolved,
        "master_patch": master_patch,
        "summary_mirror": summary_mirror,
        "jitta": jitta,
        "entry_preferred": resolved.get("entry_preferred") or entry_preferred(name),
        "auto_confirm": auto,
    }


# 後方互換（dry テスト等）
split_amazon_categories = split_pipe_values
