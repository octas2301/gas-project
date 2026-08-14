# -*- coding: utf-8 -*-
"""
食品EC LP向けサブ画像テーマ（①〜⑳）と、競合画像への紐付け。

方針（2026-08-09）:
- 分類はページ（サブ画像1枚）全体のみ（パーツ単位は将来拡張メモ）
- themeId を AI 判定 → phaseOrder はマスタから付与
- スロット: 心理プロセス順 5（条件で6）× 各 A/B
- フォールバック: 競合ユニークphase → 競合phase被り → 想像（最大2・事実数字禁止）
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from sub_image_intent import TEXT_MODELS, _extract_json_obj

LOG = logging.getLogger("set_main_image.lp_themes")

# フェーズ欠落時の想像デフォルト themeId（型のみ・数字事実はプロンプト禁止）
INVENTED_DEFAULT_BY_PHASE: Dict[int, int] = {
    1: 2,   # ベネフィット
    2: 6,   # 素材・産地（一般表現のみ）
    3: 12,  # 利用シーン
    4: 17,  # 安心安全（一般表現のみ）
    5: 20,  # セット／バリエーション（価格禁止）
}

MAX_INVENTED_SLOTS = 2

# id は 1..20（xvi → 16）
LP_THEMES: List[Dict[str, Any]] = [
    {
        "id": 1,
        "phaseOrder": 1,
        "phaseKey": "hook",
        "phase": "フック",
        "name": "ターゲットの悩みへの共感",
        "nameSlug": "nayami_kyokan",
        "contentCluster": "emotion",
        "hint": "潜在的な不満・痛みの代弁。長文コピーは避け短く。",
        "signal": "悩み・あるある・痛みを代弁するコピーが主役",
    },
    {
        "id": 2,
        "phaseOrder": 1,
        "phaseKey": "hook",
        "phase": "フック",
        "name": "理想の未来（ベネフィット）の提示",
        "nameSlug": "benefit_mirai",
        "contentCluster": "emotion",
        "hint": "購入後の最高体験を短く提示。",
        "signal": "買った後の理想体験・時短ベネフィットが主役",
    },
    {
        "id": 3,
        "phaseOrder": 1,
        "phaseKey": "hook",
        "phase": "フック",
        "name": "視覚的なシズル感のアピール",
        "nameSlug": "shizuru_sizzle",
        "contentCluster": "sizzle",
        "hint": "湯気・照り・食感など「美味しそう」を主役に。文字は最小。",
        "signal": "湯気・照り・料理接写が全面。説明カードが弱い",
    },
    {
        "id": 4,
        "phaseOrder": 2,
        "phaseKey": "product_proof",
        "phase": "プロダクト理解",
        "name": "味の正しいアピール（期待値調整）",
        "nameSlug": "aji_kitai",
        "contentCluster": "spec_proof",
        "hint": "味の方向性を正直・簡潔に。",
        "signal": "甘さ控えめ等、味の方向性・期待値調整コピー",
    },
    {
        "id": 5,
        "phaseOrder": 2,
        "phaseKey": "product_proof",
        "phase": "プロダクト理解",
        "name": "圧倒的なこだわり・開発秘話",
        "nameSlug": "kodawari_story",
        "contentCluster": "story_people",
        "hint": "ストーリーは要点のみ。文字過多禁止。",
        "signal": "試作・開発秘話・こだわりストーリー",
    },
    {
        "id": 6,
        "phaseOrder": 2,
        "phaseKey": "product_proof",
        "phase": "プロダクト理解",
        "name": "素材・産地の証明",
        "nameSlug": "sozai_sanchi",
        "contentCluster": "spec_proof",
        "hint": "素材スペックをカード数枚で。",
        "signal": "産地・等級・原料スペックカード",
    },
    {
        "id": 7,
        "phaseOrder": 2,
        "phaseKey": "product_proof",
        "phase": "プロダクト理解",
        "name": "製法・技術の独自性",
        "nameSlug": "seiho_gijutsu",
        "contentCluster": "spec_proof",
        "hint": "製法の強みを図解寄りで。",
        "signal": "製法名・工程図・独自技術",
    },
    {
        "id": 8,
        "phaseOrder": 2,
        "phaseKey": "product_proof",
        "phase": "プロダクト理解",
        "name": "生産者の顔・想い",
        "nameSlug": "seisansha_omoi",
        "contentCluster": "story_people",
        "hint": "人の顔・メッセージは短文。",
        "signal": "生産者・職人の顔写真とメッセージ",
    },
    {
        "id": 9,
        "phaseOrder": 3,
        "phaseKey": "lifestyle",
        "phase": "ライフスタイル",
        "name": "健康的価値・栄養素のアピール",
        "nameSlug": "eiyo_kenko",
        "contentCluster": "health_life",
        "hint": "栄養訴求。栄養成分表は密でも可。",
        "signal": "カロリー・ビタミン・栄養成分表が主役",
    },
    {
        "id": 10,
        "phaseOrder": 3,
        "phaseKey": "lifestyle",
        "phase": "ライフスタイル",
        "name": "一番美味しい食べ方・作り方",
        "nameSlug": "tabekata_howto",
        "contentCluster": "how_to",
        "hint": "手順は3〜5ステップ程度。",
        "signal": "手順番号・作り方図解",
    },
    {
        "id": 11,
        "phaseOrder": 3,
        "phaseKey": "lifestyle",
        "phase": "ライフスタイル",
        "name": "飽きさせないアレンジレシピ集",
        "nameSlug": "arrange_recipe",
        "contentCluster": "how_to",
        "hint": "アレンジは最大3案。各案は短タイトル＋1行。",
        "signal": "複数レシピ・別用途アレンジ",
    },
    {
        "id": 12,
        "phaseOrder": 3,
        "phaseKey": "lifestyle",
        "phase": "ライフスタイル",
        "name": "日常への組み込み提案（利用シーン）",
        "nameSlug": "nichijo_scene",
        "contentCluster": "health_life",
        "hint": "朝・夜などシーンをビジュアル中心で。",
        "signal": "朝・夜・デスク等の利用シーン",
    },
    {
        "id": 13,
        "phaseOrder": 3,
        "phaseKey": "lifestyle",
        "phase": "ライフスタイル",
        "name": "誰と楽しむか（ギフト・シェア）",
        "nameSlug": "gift_share",
        "contentCluster": "emotion",
        "hint": "ギフト／シェア用途を簡潔に。",
        "signal": "贈る・シェア・ご褒美用途",
    },
    {
        "id": 14,
        "phaseOrder": 4,
        "phaseKey": "trust",
        "phase": "信頼獲得",
        "name": "実績アピール（ランキング・販売数）",
        "nameSlug": "jisseki_rank",
        "contentCluster": "social_proof",
        "hint": "実績バッジ中心。数字は競合にあるもののみ。",
        "signal": "ランキング・累計販売など画像内の実績数字",
    },
    {
        "id": 15,
        "phaseOrder": 4,
        "phaseKey": "trust",
        "phase": "信頼獲得",
        "name": "お客様のリアルな声（レビュー・UGC）",
        "nameSlug": "review_ugc",
        "contentCluster": "social_proof",
        "hint": "口コミは2〜3件まで。各1〜2行。",
        "signal": "星・吹き出し・お客様の声",
    },
    {
        "id": 16,
        "phaseOrder": 4,
        "phaseKey": "trust",
        "phase": "信頼獲得",
        "name": "メディア掲載・プロの推薦",
        "nameSlug": "media_suisen",
        "contentCluster": "social_proof",
        "hint": "第三者権威はロゴ／短文のみ。",
        "signal": "雑誌・TV・推薦者など第三者権威",
    },
    {
        "id": 17,
        "phaseOrder": 4,
        "phaseKey": "trust",
        "phase": "信頼獲得",
        "name": "安全・安心への取り組み",
        "nameSlug": "anzen_anshin",
        "contentCluster": "spec_proof",
        "hint": "無添加・検査・工場など安心要素。",
        "signal": "無添加・検査・工場認証などの安心表示",
    },
    {
        "id": 18,
        "phaseOrder": 4,
        "phaseKey": "trust",
        "phase": "信頼獲得",
        "name": "よくある質問（Q&A）",
        "nameSlug": "faq_qa",
        "contentCluster": "how_to",
        "hint": "Q&Aは最大3組。各回答は短文。",
        "signal": "Q.&A / よくある質問形式",
    },
    {
        "id": 19,
        "phaseOrder": 5,
        "phaseKey": "closing",
        "phase": "クロージング",
        "name": "価格の妥当性・コスパの提示",
        "nameSlug": "price_cospa",
        "contentCluster": "close",
        "hint": "価格・コスパは競合にある数字のみ。煽り禁止。",
        "signal": "1食あたり・価格比較など画像内の価格根拠",
    },
    {
        "id": 20,
        "phaseOrder": 5,
        "phaseKey": "closing",
        "phase": "クロージング",
        "name": "セット内容・バリエーションと購入導線",
        "nameSlug": "set_variation",
        "contentCluster": "close",
        "hint": "セット／バリエーション案内。カートUI・架空価格は描かない。",
        "signal": "セット内容・種類バリエーション一覧（カートUI除く）",
    },
]

THEME_BY_ID = {int(t["id"]): t for t in LP_THEMES}


def theme_catalog_text() -> str:
    lines = []
    for t in LP_THEMES:
        lines.append(
            f"{t['id']:02d}. [P{t['phaseOrder']}:{t['phase']}] {t['name']} "
            f"— 判断目安: {t['signal']}／制作: {t['hint']}"
        )
    return "\n".join(lines)


def _slug(s: str) -> str:
    s = re.sub(r"[^\w\-]+", "_", (s or "").strip())
    return (s[:40] or "theme").strip("_")


def make_stem(*, slot: int, theme_id: int, proposal: str) -> str:
    meta = THEME_BY_ID[int(theme_id)]
    ab = "a" if str(proposal).lower() in ("a", "1", "01") else "b"
    return f"S{slot:02d}_T{theme_id:02d}_{_slug(meta['nameSlug'])}_AB{ab}"


def classify_image_themes(
    path: Path,
    *,
    client=None,
    model_id: Optional[str] = None,
) -> Dict[str, Any]:
    """
    1枚の競合サブ画像（ページ全体）をテーマ①〜⑳に紐付け。
    phaseOrder はマスタから付与。テーマは最大2（primary必須）。
    """
    from gemini_image import load_api_key, make_client, mime_for
    import base64

    if client is None:
        client = make_client(load_api_key())

    raw = path.read_bytes()
    b64 = base64.b64encode(raw).decode("ascii")
    mime = mime_for(path)

    prompt = f"""あなたは食品ECの商品ページ（LP）サブ画像のテーマ分類器です。
対象は画像全体（ページ単位）。パーツ単位の分割はしない。
文字は読み取り専用（書き直し案は出さない）。画像に無い事実・数字テーマを付けない。

テーマ一覧:
{theme_catalog_text()}

判断の優先順位:
1) 読める文字・数字・バッジ
2) レイアウト型（Q&A、手順番号、レビューカード、価格表など）
3) 被写体（シズル、人物、利用シーン）
4) 迷ったら面積・視線誘導の主役

タイブレーク:
- シズル大＋隅に産地 → primary=03、secondary=06可
- 栄養表が主役 → 09（安心表示が副なら17をsecondary可）
- レビュー見出し主 → 15／順位バッジ主 → 14
- 手順番号主 → 10

ルール:
- primaryThemeId: 1〜20を1つ（除外時0）
- secondaryThemeId: 無いなら null。あるなら primary 以外の1つだけ
- テーマ総数は最大2（第3は禁止）
- 配送・店舗連絡・クーポンのみ → primaryThemeId=0, secondaryThemeId=null, reject=true
- 該当が弱い場合でも最も近い1つは入れる（confidenceを下げる）

JSONのみ:
{{
  "primaryThemeId": 6,
  "secondaryThemeId": 3,
  "confidence": 0.82,
  "evidenceType": "text",
  "ocrTextPreview": "読めた文字の要約（300文字以内）",
  "reasonJa": "紐付け理由（短文）",
  "reject": false
}}
"""

    models = [model_id] if model_id else list(TEXT_MODELS)
    last_err: Optional[Exception] = None
    for mid in models:
        if not mid:
            continue
        try:
            resp = client.models.generate_content(
                model=mid,
                contents=[
                    {
                        "role": "user",
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": mime, "data": b64}},
                        ],
                    }
                ],
            )
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                try:
                    text = resp.candidates[0].content.parts[0].text
                except Exception:
                    text = ""
            obj = _extract_json_obj(text)
            if not obj:
                raise RuntimeError(f"JSON parse failed: {text[:200]}")
            if obj.get("reject") is True:
                return {
                    "path": str(path),
                    "primaryThemeId": 0,
                    "secondaryThemeId": None,
                    "themeIds": [],
                    "phaseOrder": 0,
                    "confidence": float(obj.get("confidence") or 0.0),
                    "ocrTextPreview": str(obj.get("ocrTextPreview") or "")[:500],
                    "reasonJa": str(obj.get("reasonJa") or "reject"),
                    "evidenceType": str(obj.get("evidenceType") or ""),
                    "model": mid,
                    "reject": True,
                }
            primary = int(obj.get("primaryThemeId") or 0)
            sec_raw = obj.get("secondaryThemeId", None)
            secondary: Optional[int] = None
            if sec_raw is not None and str(sec_raw).strip().lower() not in ("", "null", "none"):
                try:
                    sn = int(sec_raw)
                    if 1 <= sn <= 20 and sn != primary:
                        secondary = sn
                except (TypeError, ValueError):
                    secondary = None
            # 後方互換: themeIds があれば使う（最大2）
            ids: List[int] = []
            if 1 <= primary <= 20:
                ids.append(primary)
            if secondary:
                ids.append(secondary)
            for x in obj.get("themeIds") or []:
                try:
                    n = int(x)
                except (TypeError, ValueError):
                    continue
                if 1 <= n <= 20 and n not in ids:
                    ids.append(n)
                if len(ids) >= 2:
                    break
            if primary and 1 <= primary <= 20 and primary not in ids:
                ids = [primary] + ids
            ids = ids[:2]
            if ids and primary not in ids:
                primary = ids[0]
            secondary = ids[1] if len(ids) > 1 else None
            conf = float(obj.get("confidence") or 0.5)
            phase_order = int(THEME_BY_ID[primary]["phaseOrder"]) if primary in THEME_BY_ID else 0
            result = {
                "path": str(path),
                "primaryThemeId": primary if 1 <= primary <= 20 else 0,
                "secondaryThemeId": secondary,
                "themeIds": ids,
                "phaseOrder": phase_order,
                "confidence": conf,
                "ocrTextPreview": str(obj.get("ocrTextPreview") or "")[:500],
                "reasonJa": str(obj.get("reasonJa") or ""),
                "evidenceType": str(obj.get("evidenceType") or ""),
                "model": mid,
                "reject": False,
            }
            LOG.info(
                "theme %s primary=%s secondary=%s phaseOrder=%s",
                path.name,
                result["primaryThemeId"],
                result["secondaryThemeId"],
                result["phaseOrder"],
            )
            return result
        except Exception as e:
            last_err = e
            LOG.warning("theme classify model=%s failed: %s", mid, e)
    return {
        "path": str(path),
        "primaryThemeId": 0,
        "secondaryThemeId": None,
        "themeIds": [],
        "phaseOrder": 0,
        "confidence": 0.0,
        "ocrTextPreview": "",
        "reasonJa": f"分類失敗: {last_err}",
        "model": None,
        "error": str(last_err),
        "reject": True,
    }


def aggregate_theme_hits(
    classifications: List[Dict[str, Any]],
) -> Dict[int, Dict[str, Any]]:
    """themeId -> {count, score, paths[], ocrSnippets[], phaseOrder, contentCluster}"""
    out: Dict[int, Dict[str, Any]] = {}
    for c in classifications:
        if c.get("reject"):
            continue
        path = c.get("path") or ""
        ocr = str(c.get("ocrTextPreview") or "")
        conf = float(c.get("confidence") or 0.5)
        primary = int(c.get("primaryThemeId") or 0)
        for tid in c.get("themeIds") or []:
            tid = int(tid)
            if tid < 1 or tid > 20:
                continue
            meta = THEME_BY_ID[tid]
            bucket = out.setdefault(
                tid,
                {
                    "themeId": tid,
                    "count": 0,
                    "score": 0.0,
                    "paths": [],
                    "ocrSnippets": [],
                    "phaseOrder": int(meta["phaseOrder"]),
                    "contentCluster": meta["contentCluster"],
                },
            )
            weight = conf * (1.25 if tid == primary else 1.0)
            bucket["count"] += 1
            bucket["score"] += weight
            if path and path not in bucket["paths"]:
                bucket["paths"].append(path)
            if ocr and ocr not in bucket["ocrSnippets"]:
                bucket["ocrSnippets"].append(ocr[:300])
    return out


def _cluster(tid: int) -> str:
    return str(THEME_BY_ID[int(tid)]["contentCluster"])


def _phase(tid: int) -> int:
    return int(THEME_BY_ID[int(tid)]["phaseOrder"])


def seo_keywords_from_product_name(product_name: str) -> List[str]:
    """
    マスタ商品名から SEO ヒント語を粗く切り出す（2文字以上）。
    スロット選定の軽微ブーストとプロンプト注記に使う。
    """
    raw = str(product_name or "").strip()
    if not raw:
        return []
    parts = re.split(r"[\s　/／|｜・,，.。\-ー（）()【】\[\]「」『』]+", raw)
    out: List[str] = []
    seen = set()
    for p in parts:
        t = p.strip()
        if len(t) < 2:
            continue
        if t.isdigit():
            continue
        if t in seen:
            continue
        seen.add(t)
        out.append(t)
    return out[:12]


def _theme_seo_text(tid: int) -> str:
    meta = THEME_BY_ID.get(int(tid)) or {}
    return " ".join(
        [
            str(meta.get("name") or ""),
            str(meta.get("hint") or ""),
            str(meta.get("signal") or ""),
            str(meta.get("phase") or ""),
            str(meta.get("contentCluster") or ""),
        ]
    )


def seo_boost_for_theme(tid: int, keywords: Sequence[str]) -> float:
    if not keywords:
        return 0.0
    hay = _theme_seo_text(tid)
    boost = 0.0
    for k in keywords:
        if k and k in hay:
            boost += 1.5
    return boost


def format_seo_hints_ja(product_name: str, keywords: Sequence[str]) -> str:
    if not keywords:
        return "（商品名からのSEO語なし）"
    joined = "／".join(keywords[:8])
    return (
        f"商品名「{product_name}」由来のSEO語: {joined}。"
        "パネル見出し・吹き出しに自然に1〜2語まで織り込んでよい（誇大・事実創作は禁止）。"
    )


def _pick_best(
    cands: Sequence[Dict[str, Any]],
    *,
    used_ids: Sequence[int],
    prev_cluster: Optional[str],
    seo_keywords: Optional[Sequence[str]] = None,
) -> Optional[Dict[str, Any]]:
    kws = list(seo_keywords or [])

    def _key(b: Dict[str, Any]) -> Tuple[float, float, int, int]:
        tid = int(b["themeId"])
        return (
            float(b["score"]) + seo_boost_for_theme(tid, kws),
            float(b["count"]),
            -tid,
            tid,
        )

    ranked = sorted(cands, key=_key, reverse=True)
    for b in ranked:
        tid = int(b["themeId"])
        if tid in used_ids:
            continue
        if prev_cluster and _cluster(tid) == prev_cluster:
            continue
        return b
    # クラスタ制約だけ緩和
    for b in ranked:
        tid = int(b["themeId"])
        if tid not in used_ids:
            return b
    return None


def select_lp_slots(
    hits: Dict[int, Dict[str, Any]],
    *,
    target_slots: int = 5,
    proposals_per_slot: int = 2,
    allow_invented: bool = True,
    max_invented: int = MAX_INVENTED_SLOTS,
    product_name: str = "",
) -> List[Dict[str, Any]]:
    """
    心理プロセス順スロットを埋め、各スロット proposals_per_slot 案のジョブを返す。
    source: competitor | competitor_dup_phase | invented
    product_name があれば SEO 語で軽微ブースト。
    """
    target_slots = max(5, min(6, int(target_slots)))
    proposals_per_slot = max(1, min(2, int(proposals_per_slot)))
    seo_kws = seo_keywords_from_product_name(product_name)
    by_phase: Dict[int, List[Dict[str, Any]]] = {i: [] for i in range(1, 6)}
    for b in hits.values():
        po = int(b.get("phaseOrder") or _phase(b["themeId"]))
        if 1 <= po <= 5:
            by_phase[po].append(b)

    slots: List[Dict[str, Any]] = []
    used_ids: List[int] = []
    invented_n = 0
    prev_cluster: Optional[str] = None

    # Pass1: 各 phaseOrder から競合ユニーク
    for po in range(1, 6):
        if len(slots) >= target_slots:
            break
        pick = _pick_best(
            by_phase.get(po) or [],
            used_ids=used_ids,
            prev_cluster=prev_cluster,
            seo_keywords=seo_kws,
        )
        if not pick:
            continue
        tid = int(pick["themeId"])
        slots.append(
            {
                "themeId": tid,
                "source": "competitor",
                "phaseOrder": po,
                "hit": pick,
            }
        )
        used_ids.append(tid)
        prev_cluster = _cluster(tid)

    # Pass2: 競合フェーズ被りで不足を補う
    all_hits = sorted(
        hits.values(),
        key=lambda b: (
            float(b["score"]) + seo_boost_for_theme(int(b["themeId"]), seo_kws),
            b["count"],
            -int(b["themeId"]),
        ),
        reverse=True,
    )
    while len(slots) < target_slots:
        pick = _pick_best(
            all_hits,
            used_ids=used_ids,
            prev_cluster=prev_cluster,
            seo_keywords=seo_kws,
        )
        if not pick:
            break
        tid = int(pick["themeId"])
        slots.append(
            {
                "themeId": tid,
                "source": "competitor_dup_phase",
                "phaseOrder": _phase(tid),
                "hit": pick,
            }
        )
        used_ids.append(tid)
        prev_cluster = _cluster(tid)

    # Pass3: 想像（フェーズ欠を優先・最大 max_invented）
    covered_phases = {int(s["phaseOrder"]) for s in slots}
    missing_phases = [p for p in range(1, 6) if p not in covered_phases]
    guard = 0
    while (
        allow_invented
        and len(slots) < target_slots
        and invented_n < max_invented
        and guard < 12
    ):
        guard += 1
        po = missing_phases.pop(0) if missing_phases else ((len(slots) % 5) + 1)
        tid = int(INVENTED_DEFAULT_BY_PHASE.get(po) or 12)
        alts = [
            int(t["id"])
            for t in LP_THEMES
            if int(t["phaseOrder"]) == po and int(t["id"]) not in used_ids
        ]
        if tid not in alts:
            tid = alts[0] if alts else 0
        if not tid:
            continue
        if prev_cluster and _cluster(tid) == prev_cluster:
            non_contig = [a for a in alts if _cluster(a) != prev_cluster]
            if non_contig:
                tid = non_contig[0]
        slots.append(
            {
                "themeId": tid,
                "source": "invented",
                "phaseOrder": _phase(tid),
                "hit": {
                    "themeId": tid,
                    "count": 0,
                    "score": 0.0,
                    "paths": [],
                    "ocrSnippets": [],
                    "phaseOrder": _phase(tid),
                    "contentCluster": _cluster(tid),
                },
            }
        )
        used_ids.append(tid)
        prev_cluster = _cluster(tid)
        invented_n += 1
        covered_phases.add(_phase(tid))

    # phaseOrder 順に並べ替え（被り追加分も心理順に近づける）
    slots.sort(key=lambda s: (int(s["phaseOrder"]), -float(s["hit"].get("score") or 0)))

    jobs: List[Dict[str, Any]] = []
    for slot_i, s in enumerate(slots[:target_slots], start=1):
        tid = int(s["themeId"])
        meta = THEME_BY_ID[tid]
        hit = s["hit"]
        for pi in range(proposals_per_slot):
            proposal = "a" if pi == 0 else "b"
            stem = make_stem(slot=slot_i, theme_id=tid, proposal=proposal)
            jobs.append(
                {
                    "jobIndex": len(jobs) + 1,
                    "slotIndex": slot_i,
                    "stem": stem,
                    "themeId": tid,
                    "secondaryThemeId": None,
                    "variant": pi + 1,
                    "proposal": proposal,
                    "phase": meta["phase"],
                    "phaseOrder": int(meta["phaseOrder"]),
                    "phaseKey": meta["phaseKey"],
                    "themeName": meta["name"],
                    "themeSlug": meta["nameSlug"],
                    "themeHint": meta["hint"],
                    "contentCluster": meta["contentCluster"],
                    "source": s["source"],
                    "refPaths": list(hit.get("paths") or []),
                    "ocrSnippets": list(hit.get("ocrSnippets") or []),
                    "score": float(hit.get("score") or 0.0),
                    "count": int(hit.get("count") or 0),
                    "seoKeywords": list(seo_kws),
                }
            )
    return jobs


def select_themes_for_jobs(
    hits: Dict[int, Dict[str, Any]],
    *,
    max_themes: int = 5,
    total_jobs: int = 10,
    allow_invented: bool = True,
    product_name: str = "",
) -> List[Dict[str, Any]]:
    """
    互換API。total_jobs=10→5スロット×2案、12→6スロット×2案。
    max_themes は target_slots の目安（5 or 6）。
    """
    if total_jobs >= 12:
        target_slots = 6
    elif max_themes >= 6:
        target_slots = 6
    else:
        target_slots = 5
    return select_lp_slots(
        hits,
        target_slots=target_slots,
        proposals_per_slot=2,
        allow_invented=allow_invented,
        product_name=product_name,
    )


def jobs_summary_ja(jobs: List[Dict[str, Any]]) -> str:
    lines = ["# 選択スロット（心理プロセス順・ページ分類）", ""]
    seen_slots = set()
    for j in jobs:
        si = j.get("slotIndex") or j.get("variant")
        if si in seen_slots:
            continue
        seen_slots.add(si)
        src = j.get("source") or "competitor"
        lines.append(
            f"- S{int(j.get('slotIndex') or 0):02d} P{j.get('phaseOrder', '?')} "
            f"T{j['themeId']:02d} [{j['phase']}] {j['themeName']} "
            f"（source={src} / 競合hit={j.get('count', 0)} / score={j.get('score', 0):.2f}）"
        )
    lines.append("")
    lines.append(f"# 生成ジョブ（{len(jobs)}枚 = スロット×A/B）")
    for j in jobs:
        lines.append(
            f"- {j['stem']}: {j['themeName']} proposal={j.get('proposal', j.get('variant'))} "
            f"source={j.get('source')}"
        )
    return "\n".join(lines)
