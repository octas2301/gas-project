# -*- coding: utf-8 -*-
"""一片・詳細見本からの代表面切り出し（F型）。

汎用: JAN／商品名を焼かない。structure_ref（詳細画像）だけを正とする。
色変換・外形ロックは別パス。ここでは「詳細の微細描写を壊さず1片化する」だけ。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Optional, Tuple

# F型で勝ちに近づいた指示（色は維持可・描き直し禁止）
DETAIL_EXTRACT_PROMPT = """【詳細見本→1片・最小加工】
添付は詳細見本（表面の微細描写の正）。
白〜生成り背景に、描写がいちばん見える代表の1面／1片だけを切り出せ。
添付にある線・粒・層・繊維などの微細描写をそのまま残せ。
新しく模様を描き足すな。なめらか化・櫛目化・網目化・裂き糸化で作り替えるな。
色は添付のままでよい（別色に変えなくてよい）。
盛り全体・複数個・缶全体・袋・文字・手は描くな。
写真の四角コピペ枠は出すな。実写として自然に切り出せ。"""


def build_detail_extract_prompt(*, extra: str = "") -> str:
    extra = (extra or "").strip()
    if not extra:
        return DETAIL_EXTRACT_PROMPT
    return DETAIL_EXTRACT_PROMPT + "\n" + extra


def extract_detail_face_piece(
    *,
    structure_ref: Path,
    out_path: Path,
    model: Optional[str] = None,
    size: str = "1024x1024",
    quality: str = "high",
    extra_prompt: str = "",
) -> Dict[str, Any]:
    """詳細見本1枚 → 代表面1片JPEG。戻りに model / prompt / bytes メタ。"""
    from openai_image import generate_with_references, resolve_openai_image_model

    structure_ref = Path(structure_ref)
    out_path = Path(out_path)
    if not structure_ref.is_file():
        raise FileNotFoundError(f"structure_ref missing: {structure_ref}")

    prompt = build_detail_extract_prompt(extra=extra_prompt)
    used_model = resolve_openai_image_model(model)
    raw, used = generate_with_references(
        prompt=prompt,
        image_paths=[structure_ref],
        model=used_model,
        size=size,
        quality=quality,
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(raw)
    meta_path = out_path.with_suffix(out_path.suffix + ".txt")
    meta_path.write_text(prompt, encoding="utf-8")
    return {
        "ok": True,
        "out": str(out_path),
        "model": used,
        "nRef": 1,
        "structureRef": str(structure_ref),
        "prompt": prompt,
        "bytes": len(raw),
    }
