# -*- coding: utf-8 -*-
"""B-⑤ キャッチ置き場パターン生成（16〜20マス・金丸除く）。

ルール:
- キャッチ(赤)とセールス(青)は隣接必須（両方ある場合）
- キャッチのマス数 > セールスのマス数（両方ある場合）
- 縦書き（vertical-rl）: **範囲のみ**（赤/青の役割分割は指定しない）
- 縦型: 起点は左上(origin)のみ・左3列(col0–2)内
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, FrozenSet, List, Optional, Set, Tuple

Cell = Tuple[int, int]  # (col, row)

MIN_CELLS = 16
MAX_CELLS = 20
USABLE_COLS = 7
USABLE_ROWS = 4
# 横書きは ROI（緑枠 4×7）。縦書きは金丸列を除く 7×10 まで可（緑枠外も可）
CANVAS_TEXT_COLS = 7
CANVAS_ROWS = 10
TATE_ANCHOR = "origin"
TATE_MAX_COLS = 3
GOLD_COLS = {7, 8, 9}
OCTAS_ROWS = {7, 8, 9}


@dataclass(frozen=True)
class ZoneRect:
    c0: int
    r0: int
    w: int
    h: int
    role: str  # catch | sales | deco | range

    def cells(self) -> Set[Cell]:
        return {
            (self.c0 + dc, self.r0 + dr)
            for dr in range(self.h)
            for dc in range(self.w)
        }

    def label(self) -> str:
        return f"col {self.c0}–{self.c0 + self.w - 1}, row {self.r0}–{self.r0 + self.h - 1}"


@dataclass
class CatchPattern:
    id: str
    anchor: str
    anchor_label: str
    orientation: str
    writing_mode: str
    zones: List[ZoneRect]
    cell_n: int
    cells: List[List[int]]
    catch_n: int
    sales_n: int
    zone_mode: str = "split"  # split | range-only
    note: str = ""

    def to_dict(self) -> Dict:
        d: Dict = {
            "id": self.id,
            "anchor": self.anchor,
            "anchorLabel": self.anchor_label,
            "orientation": self.orientation,
            "writingMode": self.writing_mode,
            "zoneMode": self.zone_mode,
            "cellN": self.cell_n,
            "cells": self.cells,
            "zones": [
                {
                    "role": z.role,
                    "c0": z.c0,
                    "r0": z.r0,
                    "w": z.w,
                    "h": z.h,
                    "label": z.label(),
                }
                for z in self.zones
            ],
            "note": self.note,
        }
        if self.zone_mode == "range-only":
            d["catchN"] = None
            d["salesN"] = None
        else:
            d["catchN"] = self.catch_n
            d["salesN"] = self.sales_n
        return d


ANCHORS: Dict[str, Tuple[int, int, str]] = {
    "origin": (0, 0, "原点 (0,0)"),
    "inset_rc": (1, 1, "行列1空け (row1,col1起点)"),
    "inset_r": (0, 1, "上1行空け (row1起点)"),
    "inset_c": (1, 0, "左1列空け (col1起点)"),
}


def _in_bounds(c: Cell, *, canvas: bool = False) -> bool:
    col, row = c
    max_c = CANVAS_TEXT_COLS if canvas else USABLE_COLS
    max_r = CANVAS_ROWS if canvas else USABLE_ROWS
    return 0 <= col < max_c and 0 <= row < max_r


def _valid_cells(cells: Set[Cell], *, canvas: bool = False) -> bool:
    n = len(cells)
    if n < MIN_CELLS or n > MAX_CELLS:
        return False
    return all(_in_bounds(c, canvas=canvas) for c in cells)


def _merge_zones(zones: List[ZoneRect], *, canvas: bool = False) -> Optional[Set[Cell]]:
    merged: Set[Cell] = set()
    for z in zones:
        merged |= z.cells()
    if not _valid_cells(merged, canvas=canvas):
        return None
    return merged


def _zones_share_edge(a: ZoneRect, b: ZoneRect) -> bool:
    ca, cb = a.cells(), b.cells()
    for c, r in ca:
        for dc, dr in ((0, 1), (0, -1), (1, 0), (-1, 0)):
            if (c + dc, r + dr) in cb:
                return True
    return False


def _pair_ok(catch: ZoneRect, sales: ZoneRect) -> bool:
    if not _zones_share_edge(catch, sales):
        return False
    return len(catch.cells()) > len(sales.cells())


def _orient_and_mode(zones: List[ZoneRect], *, tate: bool) -> Tuple[str, str]:
    if tate:
        return "vertical", "vertical-rl"
    catch = next(z for z in zones if z.role == "catch")
    if catch.w >= catch.h * 2:
        return "horizontal", "horizontal-tb"
    if len(zones) > 1:
        return "hybrid", "horizontal-tb"
    return "horizontal", "horizontal-tb"


def _mk_pattern(
    anchor_key: str,
    zones: List[ZoneRect],
    *,
    suffix: str = "",
    note: str = "",
    tate: bool = False,
    canvas_bounds: bool = False,
) -> Optional[CatchPattern]:
    catch_z = [z for z in zones if z.role == "catch"]
    sales_z = [z for z in zones if z.role == "sales"]
    if len(catch_z) != 1:
        return None
    catch = catch_z[0]
    if sales_z:
        if len(sales_z) != 1:
            return None
        if not _pair_ok(catch, sales_z[0]):
            return None

    col_off, row_off, anchor_label = ANCHORS[anchor_key]
    shifted = [
        ZoneRect(
            c0=z.c0 + col_off,
            r0=z.r0 + row_off,
            w=z.w,
            h=z.h,
            role=z.role,
        )
        for z in zones
    ]
    merged = _merge_zones(shifted, canvas=canvas_bounds)
    if merged is None:
        return None

    orient, wmode = _orient_and_mode(shifted, tate=tate)
    cz = shifted[0]
    pid = f"{anchor_key}-{orient[:1]}"
    if tate:
        pid += "tate"
    pid += f"-{cz.w}x{cz.h}"
    if len(shifted) > 1:
        sz = next(z for z in shifted if z.role == "sales")
        pid += f"+s{sz.w}x{sz.h}"
    if suffix:
        pid += f"-{suffix}"

    catch_n = len(cz.cells())
    sales_n = len(next(z for z in shifted if z.role == "sales").cells()) if sales_z else 0

    return CatchPattern(
        id=pid,
        anchor=anchor_key,
        anchor_label=anchor_label,
        orientation=orient,
        writing_mode=wmode,
        zones=shifted,
        cell_n=len(merged),
        cells=sorted([list(c) for c in merged]),
        catch_n=catch_n,
        sales_n=sales_n,
        note=note,
    )


def _cells_in_tate_cols(cells: Set[Cell]) -> bool:
    return all(0 <= c < TATE_MAX_COLS for c, _ in cells)


def _bbox_zone(cells: Set[Cell], role: str = "range") -> ZoneRect:
    min_c = min(c for c, _ in cells)
    max_c = max(c for c, _ in cells)
    min_r = min(r for _, r in cells)
    max_r = max(r for _, r in cells)
    return ZoneRect(
        min_c,
        min_r,
        max_c - min_c + 1,
        max_r - min_r + 1,
        role,
    )


def _mk_range_pattern(
    anchor_key: str,
    cells: Set[Cell],
    *,
    suffix: str = "",
    note: str = "",
) -> Optional[CatchPattern]:
    """縦書き: セル集合のみ（キャッチ/セールス分割なし）。"""
    if not _cells_in_tate_cols(cells):
        return None
    if not _valid_cells(cells, canvas=True):
        return None

    _, _, anchor_label = ANCHORS[anchor_key]
    zone = _bbox_zone(cells, "range")
    min_c, min_r = zone.c0, zone.r0
    max_c = zone.c0 + zone.w - 1
    max_r = zone.r0 + zone.h - 1
    cell_n = len(cells)
    pid = f"{anchor_key}-vtate-range-n{cell_n}-{min_c}{min_r}{max_c}{max_r}"
    if suffix:
        pid += f"-{suffix}"

    extends = max_r + 1 > USABLE_ROWS or max_c + 1 > USABLE_COLS
    roi_tag = "（緑枠外含む）" if extends else ""
    full_note = f"{note} · 範囲のみ（キャッチ/セールス分割なし）{roi_tag}".strip()

    return CatchPattern(
        id=pid,
        anchor=anchor_key,
        anchor_label=anchor_label,
        orientation="vertical",
        writing_mode="vertical-rl",
        zones=[zone],
        cell_n=cell_n,
        cells=sorted([list(c) for c in cells]),
        catch_n=0,
        sales_n=0,
        zone_mode="range-only",
        note=full_note,
    )


def _gen_horizontal_yoko(anchor_key: str) -> List[CatchPattern]:
    """横書き: キャッチ上 + セールス下（接続）。"""
    out: List[CatchPattern] = []
    col_off, row_off, _ = ANCHORS[anchor_key]
    max_w = USABLE_COLS - col_off
    max_h = USABLE_ROWS - row_off

    def _try(zones: List[ZoneRect], *, note: str, suffix: str = "") -> None:
        p = _mk_pattern(anchor_key, zones, note=note, suffix=suffix, tate=False)
        if p:
            out.append(p)

    for ch in range(1, max_h + 1):
        for cw in range(3, max_w + 1):
            catch = ZoneRect(0, 0, cw, ch, "catch")
            cn = cw * ch
            if MIN_CELLS <= cn <= MAX_CELLS:
                _try([catch], note=f"横書き・キャッチのみ {ch}×{cw}={cn}マス")

            sr = ch
            if sr >= max_h:
                continue
            for sh in (1, 2):
                if sr + sh > max_h:
                    continue
                for sw in range(2, max_w + 1):
                    sales = ZoneRect(0, sr, sw, sh, "sales")
                    if not _pair_ok(catch, sales):
                        continue
                    merged = _merge_zones([catch, sales])
                    if merged is None:
                        continue
                    nn = len(merged)
                    if MIN_CELLS <= nn <= MAX_CELLS:
                        _try(
                            [catch, sales],
                            note=(
                                f"横書き キャッチ{ch}×{cw}({cn})＋"
                                f"sales{sh}×{sw} 接続 · 計{nn}マス"
                            ),
                        )
    return out


def _zones_in_tate_cols(zones: List[ZoneRect]) -> bool:
    """縦型は左3列(col 0–2)内に収める。"""
    for z in zones:
        if z.c0 < 0 or z.c0 + z.w > TATE_MAX_COLS:
            return False
    return True


def _gen_tate_gaki(anchor_key: str) -> List[CatchPattern]:
    """縦書き: 左上起点・左3列・範囲（セル集合）のみ提案。"""
    if anchor_key != TATE_ANCHOR:
        return []

    out: List[CatchPattern] = []
    col_off, row_off, _ = ANCHORS[anchor_key]
    max_w = TATE_MAX_COLS - col_off
    max_h = CANVAS_ROWS - row_off

    def _shift(cells: Set[Cell]) -> Set[Cell]:
        return {(c + col_off, r + row_off) for c, r in cells}

    def _sales_adjacent(catch: ZoneRect) -> List[ZoneRect]:
        opts: List[ZoneRect] = []
        c0, r0, cw, ch = catch.c0, catch.r0, catch.w, catch.h
        for sw in range(1, max_w - c0 + 1):
            for sh in range(1, max_h - r0 - ch + 1):
                s = ZoneRect(c0, r0 + ch, sw, sh, "sales")
                if _pair_ok(catch, s):
                    opts.append(s)
        for sw in range(1, max_w - c0 - cw + 1):
            for sh in range(1, max_h - r0 + 1):
                s = ZoneRect(c0 + cw, r0, sw, sh, "sales")
                if _pair_ok(catch, s):
                    opts.append(s)
        for sw in range(1, max_w - c0 + 1):
            for sh in range(1, r0 + 1):
                s = ZoneRect(c0, r0 - sh, sw, sh, "sales")
                if _pair_ok(catch, s):
                    opts.append(s)
        for sw in range(1, c0 + 1):
            for sh in range(1, max_h - r0 + 1):
                s = ZoneRect(c0 - sw, r0, sw, sh, "sales")
                if _pair_ok(catch, s):
                    opts.append(s)
        return opts

    seen_local: Set[FrozenSet[Cell]] = set()

    def _emit_cells(local_cells: Set[Cell], *, note: str, suffix: str) -> None:
        shifted = _shift(local_cells)
        key = frozenset(shifted)
        if key in seen_local:
            return
        seen_local.add(key)
        p = _mk_range_pattern(anchor_key, shifted, note=note, suffix=suffix)
        if p:
            out.append(p)

    for cw in range(1, max_w + 1):
        for ch in range(1, max_h + 1):
            catch = ZoneRect(0, 0, cw, ch, "catch")
            cn = cw * ch
            shape = f"左{max_w}列中{cw}列×{ch}行"

            if MIN_CELLS <= cn <= MAX_CELLS:
                _emit_cells(
                    catch.cells(),
                    note=f"縦書き範囲 {shape} = {cn}マス",
                    suffix=f"{cw}x{ch}",
                )

            if cn >= MAX_CELLS:
                continue

            for sales in _sales_adjacent(catch):
                merged = catch.cells() | sales.cells()
                if not (MIN_CELLS <= len(merged) <= MAX_CELLS):
                    continue
                _emit_cells(
                    merged,
                    note=f"縦書き範囲 {shape}＋接続 = {len(merged)}マス",
                    suffix=f"{cw}x{ch}+m",
                )

    return out


def generate_all_patterns() -> List[CatchPattern]:
    all_p: List[CatchPattern] = []
    seen_cells: Set[Tuple[FrozenSet[Cell], str]] = set()
    seen_ids: Set[str] = set()
    for anchor in ANCHORS:
        batch = _gen_horizontal_yoko(anchor) + _gen_tate_gaki(anchor)
        for p in batch:
            cell_key = frozenset((c[0], c[1]) for c in p.cells)
            dedup_key = (cell_key, p.writing_mode)
            if dedup_key in seen_cells:
                continue
            seen_cells.add(dedup_key)
            base_id = p.id
            n = 1
            while p.id in seen_ids:
                n += 1
                p.id = f"{base_id}-{n}"
            seen_ids.add(p.id)
            all_p.append(p)
    all_p.sort(
        key=lambda x: (
            x.writing_mode,
            x.anchor,
            x.cell_n,
            x.id,
        )
    )
    return [p for p in all_p if MIN_CELLS <= p.cell_n <= MAX_CELLS]


def patterns_by_anchor(patterns: List[CatchPattern]) -> Dict[str, List[CatchPattern]]:
    out: Dict[str, List[CatchPattern]] = {k: [] for k in ANCHORS}
    for p in patterns:
        out.setdefault(p.anchor, []).append(p)
    return out
