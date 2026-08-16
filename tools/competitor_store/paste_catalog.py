# -*- coding: utf-8 -*-
"""P1: Catalog 階段で貼付ブロックの空ASINだけ埋める（仮想テスト用）。HTML禁止。"""
from __future__ import annotations

import re
from typing import Callable

STAGES_MAIN = ("A_id", "B_kw_jan", "D_maker_name", "E_core")
STAGE_DIAG = "C_ai_name"


def norm_text(s: str) -> str:
    return re.sub(r"[\s\u3000]+", "", str(s or "")).lower()


def title_has_maker(title: str, maker: str) -> bool:
    m = str(maker or "").strip()
    if not m:
        return True
    return norm_text(m) in norm_text(title)


def short_core(name: str, maker: str) -> str:
    s = str(name or "")
    mk = str(maker or "").strip()
    if mk:
        s = s.replace(mk, " ")
    words = [w for w in re.split(r"[\s\u3000]+", s) if len(w) >= 2]
    if words:
        return " ".join(words[:2])
    return re.sub(r"\s+", "", s)[:12]


def is_circle(eval_cell) -> bool:
    return str(eval_cell or "").strip() == "◎"


def fill_empty_asins(block: dict, search: Callable[[str, dict], list[dict]], max_fill: int = 15) -> dict:
    """block: jan, maker, name, rows[{asin,eval,title}]. search(stage, block)->[{asin,title}]."""
    rows = [dict(r) for r in (block.get("rows") or [])]
    maker = block.get("maker") or ""
    used = {str(r.get("asin") or "").strip().upper() for r in rows if str(r.get("asin") or "").strip()}
    slots = [
        r
        for r in rows
        if not str(r.get("asin") or "").strip() and not is_circle(r.get("eval"))
    ]
    if rows and not slots:
        return {
            "rows": rows,
            "stage": "skip_no_empty",
            "filled": 0,
            "skipped_circle": sum(1 for r in rows if is_circle(r.get("eval"))),
        }
    stage_used = ""
    hits: list[dict] = []
    for stage in STAGES_MAIN:
        raw = search(stage, block) or []
        kept = []
        for h in raw:
            asin = str(h.get("asin") or "").strip().upper()
            title = str(h.get("title") or "")
            if not re.match(r"^[A-Z0-9]{10}$", asin):
                continue
            if not title_has_maker(title, maker):
                continue
            if asin in used:
                continue
            kept.append({"asin": asin, "title": title, "stage": stage})
        if kept:
            hits = kept
            stage_used = stage
            break

    filled = 0
    for h in hits:
        if filled >= max_fill:
            break
        placed = False
        for r in rows:
            if str(r.get("asin") or "").strip():
                continue
            if is_circle(r.get("eval")):
                continue
            r["asin"] = h["asin"]
            if not str(r.get("title") or "").strip():
                r["title"] = h["title"]
            used.add(h["asin"])
            filled += 1
            placed = True
            break
        if not placed and not rows:
            rows.append({"asin": h["asin"], "title": h["title"], "eval": ""})
            used.add(h["asin"])
            filled += 1

    return {
        "rows": rows,
        "stage": stage_used,
        "filled": filled,
        "skipped_circle": sum(1 for r in (block.get("rows") or []) if is_circle(r.get("eval"))),
    }
