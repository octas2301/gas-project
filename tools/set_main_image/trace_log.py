# -*- coding: utf-8 -*-
"""PoC実行トレース（見本がどう渡されたかを後から追える）。"""
from __future__ import annotations

import hashlib
import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

LOG = logging.getLogger("set_main_image.trace")


def file_fingerprint(path: Path) -> Dict[str, Any]:
    data = path.read_bytes()
    h = hashlib.sha256(data).hexdigest()[:16]
    try:
        from PIL import Image
        from io import BytesIO

        im = Image.open(BytesIO(data))
        wh = {"width": im.size[0], "height": im.size[1], "mode": im.mode}
    except Exception:
        wh = {}
    return {
        "path": str(path),
        "name": path.name,
        "bytes": len(data),
        "sha256_16": h,
        **wh,
    }


def write_run_trace(
    out_json: Path,
    *,
    mall: str,
    engine: str,
    mode: str,
    model_id: Optional[str],
    set_count: int,
    unit: str,
    prompt: str,
    inputs: List[Dict[str, Any]],
    api_path: Optional[str],
    notes: List[str],
    extra: Optional[Dict[str, Any]] = None,
) -> Path:
    """
    mode 例:
      - amazon_ai_fullgen … AIが全面生成（見本は参照入力）
      - rakuten_layer_only … ベース画素をコピーし数字レイヤのみ重ね
    """
    payload: Dict[str, Any] = {
        "runAt": datetime.now(timezone.utc).isoformat(),
        "mall": mall,
        "engine": engine,
        "mode": mode,
        "modelId": model_id,
        "setCount": set_count,
        "unit": unit,
        "apiPath": api_path,
        "inputs": inputs,
        "prompt": prompt,
        "promptChars": len(prompt or ""),
        "notes": notes,
        "extra": extra or {},
    }
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info(
        "trace mall=%s mode=%s model=%s inputs=%s -> %s",
        mall,
        mode,
        model_id,
        [i.get("role") for i in inputs],
        out_json.name,
    )
    return out_json
