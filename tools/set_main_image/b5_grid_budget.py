# -*- coding: utf-8 -*-
"""B-⑤ MAIN — 10×10 マス予算（おしゃれ優先・段階閾値）。

Pillow 後載せ文字は使わない。AI 生成＋条件指定プロンプト＋事後 Gemini gate。
"""
from __future__ import annotations

import os
from typing import Any, Dict, List, Optional, Tuple

from b5_virt_abc import (
    P3_CATCH_ROI_COLS,
    P3_CATCH_ROI_ROWS,
    P3_CATCH_TARGET_CELLS,
    P3_TEXT_MAX_CELLS,
    P3_TEXT_RETRY_MAX_CELLS,
    _p3_expected_zone_cells,
    _p3_text_zones,
)

# 21–24 = 20に収まらなかったときのバッファ（retry しない）
P3_TEXT_BUFFER_MAX_CELLS = 24

TIER_OK = "ok"  # ≤20
TIER_BUFFER = "buffer"  # 21–24
TIER_RETRY = "retry"  # ≥25


def retry_threshold_cells() -> int:
    """再生成を開始する hitN（既定 25 = 25以上で retry）。"""
    raw = (os.environ.get("B5_GRID_RETRY_AT") or "").strip()
    if raw.isdigit():
        return max(1, int(raw))
    return int(P3_TEXT_RETRY_MAX_CELLS)


def tier_policy() -> str:
    """tiered（既定）| strict20（副作用時の切戻し: 21超で retry 扱い）。"""
    return (os.environ.get("B5_GRID_TIER_POLICY") or "tiered").strip().lower()


def classify_grid_tier(hit_n: int) -> str:
    n = int(hit_n or 0)
    if tier_policy() == "strict20":
        if n <= P3_TEXT_MAX_CELLS:
            return TIER_OK
        return TIER_RETRY
    if n <= P3_TEXT_MAX_CELLS:
        return TIER_OK
    if n <= P3_TEXT_BUFFER_MAX_CELLS:
        return TIER_BUFFER
    return TIER_RETRY


def grid_needs_retry(hit_n: int) -> bool:
    return classify_grid_tier(hit_n) == TIER_RETRY


def grid_rank_penalty(hit_n: int) -> int:
    """try 選定 rank 用。バッファ帯はペナルティなし。"""
    return 500 if grid_needs_retry(hit_n) else 0


def compute_grid_budget(
    catch_head: str,
    sales_talk: str,
) -> Dict[str, Any]:
    catch_head = (catch_head or "").strip()
    sales_talk = (sales_talk or "").strip()
    catch_cols, catch_rows, sales_cols, sales_rows = _p3_text_zones(
        catch_head, sales_talk
    )
    expected = _p3_expected_zone_cells(
        catch_cols, catch_rows, sales_cols, sales_rows
    )
    return {
        "targetCells": P3_CATCH_TARGET_CELLS,
        "acceptMaxCells": P3_TEXT_MAX_CELLS,
        "bufferMaxCells": P3_TEXT_BUFFER_MAX_CELLS,
        "retryAtCells": retry_threshold_cells(),
        "catchCols": catch_cols,
        "catchRows": list(catch_rows),
        "salesCols": sales_cols,
        "salesRows": list(sales_rows),
        "expectedZoneN": len(expected),
        "catchHead": catch_head,
        "salesTalk": sales_talk,
    }


def build_soft_budget_prompt_block(
    catch_head: str,
    sales_talk: str,
) -> str:
    """AI 向け条件指定（努力目標）。Pillow なし。"""
    if (os.environ.get("B5_GRID_SOFT_BUDGET") or "1").strip().lower() in (
        "0",
        "false",
        "no",
    ):
        return ""
    b = compute_grid_budget(catch_head, sales_talk)
    catch_rows_s = ",".join(str(y) for y in b["catchRows"]) or "0–2"
    sales_rows_s = ",".join(str(y) for y in b["salesRows"]) or "（なし）"
    sales_line = (
        f"セールストーク小: 行{b['salesRows']}・列0–{max(0, b['salesCols'] - 1)}。"
        if b["salesTalk"]
        else "セールストーク: 無し可。"
    )
    return f"""【10×10マス・努力目標（おしゃれさ優先）】
- **目指す**: 後載せ文字＋装飾が **{b['acceptMaxCells']}マス以内**（理想{b['targetCells']}。金丸4マスは後工程）。
- **21–24マス**はやむを得ないバッファ。**25マス以上は避ける**。
- ROI: 左上 {P3_CATCH_ROI_ROWS}行×{P3_CATCH_ROI_COLS}列。キャッチ大: 行{catch_rows_s}・列0–{max(0, b['catchCols'] - 1)}。{sales_line}
- **IMAGE_STYLE_TYPO のおしゃれさ・目を引く力を最優先**。コンパクトに収めつつ地味にしない。
- 白矩形・掛け軸・画面横断帯・多段囲みは禁止。上部左1ブロック。
"""


def build_vertical_range_prompt_block(pattern: Dict[str, Any]) -> str:
    """縦書き採用パターン（範囲のみ）を AI プロンプトに変換。"""
    cells = pattern.get("cells") or []
    cell_n = int(pattern.get("cellN") or len(cells))
    if not cells:
        return ""
    cols = max(int(c) for c, _ in cells) + 1
    rows = max(int(r) for _, r in cells) + 1
    cell_lines = ", ".join(f"({c},{r})" for c, r in sorted(cells, key=lambda x: (x[1], x[0])))
    return f"""【縦書き・配置範囲のみ（vertical-rl）】
- 日本語縦組み（**和風トーン商品向けの書字方向**。縦書き＝おしゃれの同義ではない）。
- **キャッチとセールスの位置・大小分割は指定しない**（文案量でAIが決める）。
- 文字＋装飾の努力目標: **{cell_n}マス**（理想16–20）。左上起点・**左3列(col0–2)** 付近。
- 目安ボックス: 約{cols}列×{rows}行。使用セル: {cell_lines}
- この範囲から大きくはみ出さない。範囲内の文言配分は固定しない。
- **IMAGE_STYLE_TYPO のおしゃれさ最優先**。白矩形・掛け軸・太帯禁止。
"""



def build_horizontal_split_prompt_block(pattern: Dict[str, Any]) -> str:
    """横書き採用パターン（キャッチ/セールス分割）。"""
    cells = pattern.get("cells") or []
    cell_n = int(pattern.get("cellN") or len(cells))
    if not cells:
        return ""
    zones = pattern.get("zones") or []
    catch_z = next((z for z in zones if z.get("role") == "catch"), None)
    sales_z = next((z for z in zones if z.get("role") == "sales"), None)
    catch_line = (
        f"キャッチ目安: col {catch_z['c0']}–{catch_z['c0'] + catch_z['w'] - 1}, "
        f"row {catch_z['r0']}–{catch_z['r0'] + catch_z['h'] - 1} "
        f"({pattern.get('catchN')}マス)"
        if catch_z
        else "キャッチ目安: （zones参照）"
    )
    sales_line = (
        f"セールス目安: col {sales_z['c0']}–{sales_z['c0'] + sales_z['w'] - 1}, "
        f"row {sales_z['r0']}–{sales_z['r0'] + sales_z['h'] - 1} "
        f"({pattern.get('salesN')}マス)。**キャッチ > セールス**。"
        if sales_z
        else "セールストーク: 無し可。"
    )
    cell_lines = ", ".join(f"({c},{r})" for c, r in sorted(cells, key=lambda x: (x[1], x[0])))
    return f"""【横書き・採用レイアウト目安（horizontal-tb）】
- id={pattern.get('id')} · 努力目標 **{cell_n}マス**（16–20）。
- 横書きでも縦書きと同等以上におしゃれに仕上げよ（縦書き＝おしゃれ、ではない）。
- {catch_line}
- {sales_line}
- 使用セル: {cell_lines}
- 枠は目安。行数は上限。見出しが1行で収まるなら1行にせよ。行を埋めるために改行するな。
- キャッチは、行を増やす前に字を小さくして横幅へ収めよ。助詞（が、の、に、で、を、へ、と、は、も）の直後で改行するな。
- おしゃれさ最優先。白矩形・掛け軸・太帯禁止。
"""



def build_adopted_zone_prompt_block(
    adopted: Optional[Dict[str, Any]] = None,
    *,
    theme_index: int = 0,
    n_themes: int = 1,
    materials: Optional[Dict[str, Any]] = None,
) -> Tuple[str, Dict[str, Any]]:
    """採用 JSON から1テーマ分のプロンプト追記。戻り: (block, meta)。"""
    from b5_catch_zone_adopted import (
        infer_wafu_score,
        pattern_cells,
        pick_adopted_for_theme_balanced,
        plan_writing_modes,
    )

    data = adopted if adopted is not None else None
    p = pick_adopted_for_theme_balanced(
        theme_index,
        n_themes,
        materials,
        data,
    )
    if not p:
        return "", {}
    modes = plan_writing_modes(n_themes, infer_wafu_score(materials))
    cells = sorted(pattern_cells(p), key=lambda x: (x[1], x[0]))
    meta = {
        "id": p.get("id"),
        "zoneMode": p.get("zoneMode"),
        "writingMode": p.get("writingMode"),
        "cellN": p.get("cellN"),
        "cells": [list(c) for c in cells],
        "themeIndex": theme_index,
        "plannedMode": modes[theme_index % len(modes)] if modes else "",
        "wafuScore": round(infer_wafu_score(materials), 3),
        "source": "compose",
    }
    if p.get("zoneMode") == "range-only" or p.get("writingMode") == "vertical-rl":
        block = build_vertical_range_prompt_block(p)
    else:
        block = build_horizontal_split_prompt_block(p)
    return block, meta


def build_compose_grid_prompt_block(
    catch_head: str,
    sales_talk: str,
    *,
    theme_index: int = 0,
    n_themes: int = 1,
    materials: Optional[Dict[str, Any]] = None,
) -> Tuple[str, Dict[str, Any]]:
    """compose 用マス指示。採用パターンあり時は旧横 ROI を出さない（Phase 3）。"""
    adopted_block, meta = build_adopted_zone_prompt_block(
        theme_index=theme_index,
        n_themes=n_themes,
        materials=materials,
    )
    if adopted_block:
        return adopted_block, meta
    return build_soft_budget_prompt_block(catch_head, sales_talk), {}


def build_retry_note_for_hit(
    hit_n: int,
    *,
    checker: bool,
    hero_half: int,
) -> Tuple[str, str]:
    """(retry_note, reason_tag)"""
    tier = classify_grid_tier(hit_n)
    if tier == TIER_RETRY:
        note = (
            f"前回は後載せ文字が{hit_n}マス（努力目標{P3_TEXT_MAX_CELLS}・"
            f"21–24はバッファ・{retry_threshold_cells()}以上は避ける）。"
        )
        note += (
            "おしゃれさは維持しつつ、帯・囲みを細く。文言は切るな。"
            "キャッチのインパクトは落とさない。"
            "白矩形・掛け軸・太帯・多段囲み禁止。上部左1ブロック。"
            "右上金丸聖域・右下Octas聖域に文字を入れるな。"
            "IMAGE_PACKAGE_TRUTHの袋は必ず残せ。"
        )
        if checker:
            note += " 市松模様を描くな。"
        return note, "retry_25plus"
    if tier == TIER_BUFFER:
        return (
            f"前回は{hit_n}マス（バッファ帯）。**20以内が理想**だがおしゃれさ優先。"
            "装飾を1段だけコンパクトに。キャッチ文言は維持。"
            "IMAGE_PACKAGE_TRUTHの袋は必ず残せ。",
            "buffer_soft",
        )
    return "", "ok"
