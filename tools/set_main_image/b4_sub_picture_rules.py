# -*- coding: utf-8 -*-
"""B-⑤ サブ絵合否（P3 第1スライス）。MAIN袋・WRAP契約をサブ経路へ寄せる。画像生成なしの純関数。

正本: docs/org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md §4.5.10
対象: G7 / G6寄せ / G1寄せ
切戻し: 環境変数 B4_SUB_PICTURE_GATE=0
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence


def picture_gate_enabled() -> bool:
    raw = str(os.environ.get("B4_SUB_PICTURE_GATE") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def scrub_scene_for_sub(text: str) -> str:
    """旧WRAP混ぜ文を落とす（MAIN scrub と同趣旨）。"""
    from b5_virt_abc import scrub_old_wrap_mix_ja

    return scrub_old_wrap_mix_ja(str(text or ""))


def strip_invent_wrap_ban(text: str, *, has_wrap: bool) -> str:
    """WRAP見本があるとき『個包装を描くな』系を除去。"""
    t = str(text or "")
    if not has_wrap:
        return t
    bans = (
        "個包装を新たに描くな",
        "個包装を描くな",
        "包装を新たに描くな",
        "小袋を描くな",
    )
    out = t
    for b in bans:
        out = out.replace(b, "")
    return out


def prepare_overview_for_prompt(overview: str, *, has_wrap: bool) -> str:
    t = scrub_scene_for_sub(overview)
    return strip_invent_wrap_ban(t, has_wrap=has_wrap).strip()


def wrap_gate_failed(chk: Optional[Dict[str, Any]]) -> bool:
    """
    WRAP合否。skipped は落とさない。
    frontOk 欠落は check_drawn_wraps 側で false 扱い済み。
    """
    if not isinstance(chk, dict):
        return False
    if chk.get("skipped"):
        return False
    return not bool(chk.get("ok"))


def prompt_wrap_contract_errors(prompt: str, *, has_wrap: bool) -> List[str]:
    """指示文に必須契約が載っているか（仮想・配線検収用）。"""
    p = str(prompt or "")
    errs: List[str] = []
    if not has_wrap:
        return errs
    need = (
        ("横か後ろ", "G7:place"),
        ("カメラ最近傍", "G7:front"),
        ("小皿に混ぜるな", "G7:no_mix"),
        ("1〜2個", "G7:count"),
        ("大袋", "G7:sep_bag"),
        ("個包装", "G7:wrap"),
        ("左下", "G7:bl_keepout"),
    )
    for needle, code in need:
        if needle not in p:
            errs.append(code)
    if "少し混ぜ" in p:
        errs.append("G7:old_mix")
    return errs


def prompt_wrap_look_ok(prompt: str, *, has_wrap: bool, look_snippet: str = "") -> bool:
    if not has_wrap:
        return True
    p = str(prompt or "")
    if "個包装見本の見え方" not in p and not look_snippet:
        return True  # look が空のときはブロック無しでも可（読み失敗）
    if look_snippet and look_snippet not in p:
        return False
    return "個包装見本の見え方" in p or not look_snippet


def bag_aspect_rows_from_paths(
    bags: Sequence[Any],
) -> List[Dict[str, Any]]:
    """袋パスから縦÷横行を作る。測れない袋はスキップ。商品名は入れない。"""
    from b5_bag_aspect import content_metrics

    rows: List[Dict[str, Any]] = []
    i = 0
    for raw in bags or []:
        p = Path(str(raw or ""))
        if not p.is_file():
            continue
        m = content_metrics(p)
        if not m:
            continue
        i += 1
        rows.append(
            {
                "role": f"IMAGE_PACKAGE_PART_{i}",
                "ratio": m.get("ratio"),
                "contentH": m.get("contentH"),
                "source": "file",
                "path": str(p),
            }
        )
    return rows


def bag_prompt_and_verdict(
    bags: Sequence[Any],
    *,
    out_dir: Optional[Path] = None,
    theme_id: str = "",
) -> Dict[str, Any]:
    """
    G1: 袋2枚以上で指示ブロック＋任意で袋判定.txt。
    1枚以下は空。商品名は書かない。
    """
    from b5_bag_aspect import aspect_prompt_block

    rows = bag_aspect_rows_from_paths(bags)
    out: Dict[str, Any] = {"rows": rows, "block": "", "verdictPath": ""}
    if len(rows) < 2:
        return out
    out["block"] = aspect_prompt_block(rows)
    if out_dir is not None:
        dest = Path(out_dir)
        dest.mkdir(parents=True, exist_ok=True)
        lines = ["袋判定（サブ・実測）", f"theme={theme_id or '-'}"]
        heights = [float(r.get("contentH") or 0) for r in rows]
        base = max(heights) if heights else 0.0
        for r in rows:
            h = float(r.get("contentH") or 0)
            rel = (h / base) if base > 0 else 0.0
            lines.append(
                f"- {r.get('role')}: 縦÷横={r.get('ratio')} contentH={r.get('contentH')} 対比={rel:.2f}"
            )
        path = dest / "袋判定.txt"
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        out["verdictPath"] = str(path)
    return out


def build_sub_food_front(
    *,
    has_identity: bool,
    has_sizzle: bool = False,
    has_secondary: bool = False,
    has_wrap: bool = False,
    has_piece_set: bool = False,
    food_mode: str = "identity",
    mix_parts: Optional[Sequence[str]] = None,
    mix_part_bags: bool = False,
    food_look_block: str = "",
    wrap_look_block: str = "",
    bag_aspect_block: str = "",
    mix_block: str = "",
) -> str:
    """サブ用・食材正面ブロック（MAIN food_unit_prompt_block_main を流用）。"""
    from food_unit_canon import food_unit_prompt_block_main

    if not has_identity and not has_wrap and not has_piece_set:
        extra = str(mix_block or "").strip()
        return extra

    front = food_unit_prompt_block_main(
        has_identity=bool(has_identity),
        has_sizzle=bool(has_sizzle),
        has_secondary=bool(has_secondary),
        has_wrap=bool(has_wrap),
        has_piece_set=bool(has_piece_set),
        food_mode=str(food_mode or "identity"),
        mix_parts=list(mix_parts or []),
        mix_part_bags=bool(mix_part_bags),
    ).rstrip()
    look = str(food_look_block or "").strip()
    if look:
        front = front + "\n\n" + look
    wlook = str(wrap_look_block or "").strip()
    if has_wrap and wlook:
        front = front + "\n\n" + wlook
    bag = str(bag_aspect_block or "").strip()
    if bag:
        front = front + "\n\n" + bag
    extra = str(mix_block or "").strip()
    if extra:
        front = extra + "\n\n" + front
    return front
