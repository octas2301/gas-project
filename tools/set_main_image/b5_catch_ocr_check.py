# -*- coding: utf-8 -*-
"""後載せキャッチ／セールスの文字検査。マス数ではなく中身。"""
from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, Optional

from gemini_image import load_api_key, make_client, mime_for
from sub_image_intent import TEXT_MODELS

LOG = logging.getLogger("set_main_image.b5_catch_ocr")


def _norm(s: str) -> str:
    t = str(s or "")
    t = re.sub(r"[\s　、。・，．!！?？【】()（）「」『』]", "", t)
    return t


def text_covers(expected: str, ocr: str) -> bool:
    """指定文が読取に含まれる、または読取が指定の延長（て／で等）なら通す。"""
    e = _norm(expected)
    o = _norm(ocr)
    if not e:
        return True
    if not o:
        return False
    if e in o or o in e:
        return True
    if o.startswith(e) or e.startswith(o):
        return True
    return False


def _parse_obj(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        return None
    try:
        obj = json.loads(m.group(0))
    except Exception:
        return None
    return obj if isinstance(obj, dict) else None


def check_catch_text(
    img_path: Path,
    *,
    catch_head: str,
    sales_talk: str = "",
) -> Dict[str, Any]:
    """絵の大きい字・小さい字を読む。読めないときは検査スキップ（落とさない）。"""
    p = Path(img_path)
    out: Dict[str, Any] = {
        "ok": True,
        "skipped": False,
        "ocrCatch": "",
        "ocrSales": "",
        "catchOk": True,
        "salesOk": True,
        "retryNote": "",
        "error": "",
    }
    if not p.is_file():
        out["skipped"] = True
        out["error"] = "no_file"
        return out
    prompt = (
        "楽天メイン画像です。後載せの大きい見出しと、あれば小さいセールストークだけ読んでください。\n"
        "袋の印刷は無視。推測で足すな。\n"
        'JSONのみ: {"catch":"読めた見出し","sales":"読めたセールス。無ければ空"}'
    )
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        out["skipped"] = True
        out["error"] = str(e)
        return out
    import base64

    mime = mime_for(p)
    data = base64.b64encode(p.read_bytes()).decode("ascii")
    last = ""
    obj = None
    for mid in TEXT_MODELS:
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
                            {"inline_data": {"mime_type": mime, "data": data}},
                        ],
                    }
                ],
            )
            text = getattr(resp, "text", None) or ""
            obj = _parse_obj(text)
            if obj:
                out["model"] = mid
                break
            last = (text or "")[:160]
        except Exception as e:
            last = str(e)
            LOG.warning("catch ocr fail model=%s: %s", mid, e)
    if obj is None:
        out["skipped"] = True
        out["error"] = last or "parse_failed"
        return out
    ocr_c = str(obj.get("catch") or "").strip()
    ocr_s = str(obj.get("sales") or "").strip()
    out["ocrCatch"] = ocr_c
    out["ocrSales"] = ocr_s
    catch_ok = text_covers(catch_head, ocr_c)
    sales_ok = True
    if str(sales_talk or "").strip():
        sales_ok = text_covers(sales_talk, ocr_s) or text_covers(sales_talk, ocr_c)
    out["catchOk"] = catch_ok
    out["salesOk"] = sales_ok
    out["ok"] = catch_ok and sales_ok
    if not out["ok"]:
        out["retryNote"] = (
            "前回の後載せ文字が指定と違う。"
            f"キャッチ（大）は「{catch_head}」のまま。"
            + (f"セールストーク（小）は「{sales_talk}」のまま。" if sales_talk else "")
            + "言い換え・語の追加禁止。"
        )
    LOG.info(
        "catch ocr ok=%s catch=%s/%s sales=%s/%s",
        out["ok"],
        ocr_c,
        catch_head,
        ocr_s,
        sales_talk,
    )
    return out
