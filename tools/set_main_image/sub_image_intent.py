# -*- coding: utf-8 -*-
"""
サブ画像用: 競合画像の文字列読取（選定専用）＋意図分類。

- 読む: 意図・用途判定のため（必須）
- 画像上の文字は改変・再描画しない（合成側の禁止）
"""
from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

LOG = logging.getLogger("set_main_image.sub_image_intent")

# 採用 / 除外
INTENT_USE = {
    "product_benefit",
    "spec_volume",
    "brand_intro",
    "howto_storage",
    "comparison",
}
INTENT_REJECT = {
    "shipping",
    "store_hours",
    "contact_support",
    "promo_only",
    "ui_policy",
}

TEXT_MODELS = (
    "gemini-flash-latest",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
)


def _extract_json_obj(text: str) -> Optional[dict]:
    if not text:
        return None
    t = text.strip()
    if t.startswith("```"):
        t = re.sub(r"^```(?:json)?\s*", "", t)
        t = re.sub(r"\s*```$", "", t)
    try:
        obj = json.loads(t)
        if isinstance(obj, dict):
            return obj
    except json.JSONDecodeError:
        pass
    m = re.search(r"\{[\s\S]*\}", t)
    if not m:
        return None
    try:
        obj = json.loads(m.group(0))
        return obj if isinstance(obj, dict) else None
    except json.JSONDecodeError:
        return None


def classify_competitor_image(
    path: Path,
    *,
    client=None,
    model_id: Optional[str] = None,
) -> Dict[str, Any]:
    """
    1枚について OCR相当の読取＋意図ラベル。
    戻り: intentLabel, decision(use|reject|review), ocrTextPreview, reasonJa, hasProductVisible
    """
    from gemini_image import load_api_key, make_client, mime_for
    import base64

    if client is None:
        client = make_client(load_api_key())

    raw = path.read_bytes()
    b64 = base64.b64encode(raw).decode("ascii")
    mime = mime_for(path)

    prompt = """あなたは Amazon / 楽天 の商品サブ画像の分類器です。
画像内の日本語・英語の文字を読み取り、商品説明・アピール用途かを判定してください。
文字は読み取り専用です（書き直し案は出さない）。

intentLabel は次のいずれか1つ:
- product_benefit … 特徴・メリット・ポイント訴求
- spec_volume … 内容量・スペック・セット数・成分など
- brand_intro … ブランド／シリーズ紹介（商品実物なしでも可）
- howto_storage … 使い方・保存・調理など
- comparison … 比較・選び方
- shipping … 配送・送料・お届け（商品と直接関係薄い）
- store_hours … 店舗・営業時間
- contact_support … 問い合わせ・電話・サポート
- promo_only … クーポン・セール煽りのみ
- ui_policy … 返品手続などモール/店舗ポリシー
- other_unknown … 判断不能

decision:
- use … 上記の商品直結系
- reject … shipping/store_hours/contact_support/promo_only/ui_policy
- review … other_unknown や迷う場合

hasProductVisible: 商品実物がはっきり写っているか true/false
layoutHint: product_present | text_explainer | mixed

JSONのみ:
{
  "intentLabel": "...",
  "decision": "use|reject|review",
  "ocrTextPreview": "読めた文字の要約（200文字以内）",
  "reasonJa": "判定理由（短文）",
  "hasProductVisible": true,
  "layoutHint": "text_explainer"
}
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
            label = str(obj.get("intentLabel") or "other_unknown").strip()
            decision = str(obj.get("decision") or "").strip().lower()
            if not decision:
                if label in INTENT_REJECT:
                    decision = "reject"
                elif label in INTENT_USE:
                    decision = "use"
                else:
                    decision = "review"
            # 安全側: 除外ラベルは必ず reject
            if label in INTENT_REJECT:
                decision = "reject"
            result = {
                "intentLabel": label,
                "decision": decision,
                "ocrTextPreview": str(obj.get("ocrTextPreview") or "")[:400],
                "reasonJa": str(obj.get("reasonJa") or ""),
                "hasProductVisible": bool(obj.get("hasProductVisible")),
                "layoutHint": str(obj.get("layoutHint") or ""),
                "model": mid,
                "path": str(path),
            }
            LOG.info(
                "intent %s decision=%s label=%s",
                path.name,
                decision,
                label,
            )
            return result
        except Exception as e:
            last_err = e
            LOG.warning("classify model=%s failed: %s", mid, e)
    return {
        "intentLabel": "other_unknown",
        "decision": "review",
        "ocrTextPreview": "",
        "reasonJa": f"分類失敗: {last_err}",
        "hasProductVisible": False,
        "layoutHint": "",
        "model": None,
        "path": str(path),
        "error": str(last_err),
    }


def classify_many(
    paths: List[Path],
    *,
    client=None,
) -> List[Dict[str, Any]]:
    from gemini_image import load_api_key, make_client

    if client is None:
        client = make_client(load_api_key())
    return [classify_competitor_image(p, client=client) for p in paths]
