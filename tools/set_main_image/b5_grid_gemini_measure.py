# -*- coding: utf-8 -*-
"""100マスガイド — Gemini Vision 計測＋オーバレイ／try 比較 HTML。"""
from __future__ import annotations

import base64
import html
import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

from PIL import Image, ImageDraw

from b5_virt_abc import (
    CANVAS,
    GRID_MEASURE_ALGO,
    P3_CATCH_ROI_COLS,
    P3_CATCH_ROI_ROWS,
    P3_CATCH_TARGET_CELLS,
    P3_TEXT_MAX_CELLS,
    _p3_expected_zone_cells,
    _p3_text_check,
    _p3_text_zones,
    measure_grid20_p3,
)

try:
    from b5_grid_budget import classify_grid_tier
except ImportError:
    def classify_grid_tier(hit_n: int) -> str:  # type: ignore[misc]
        n = int(hit_n or 0)
        if n <= P3_TEXT_MAX_CELLS:
            return "ok"
        if n <= 24:
            return "buffer"
        return "retry"
from gemini_image import load_api_key as load_gemini_key
from gemini_image import make_client as make_gemini_client
from gemini_image import mime_for

LOG = logging.getLogger("set_main_image.grid_gemini")

DEFAULT_GEMINI_GRID_MODEL = "gemini-3-flash-preview"
GEMINI_GRID_FALLBACKS = (
    "gemini-3-flash-preview",
    "gemini-flash-latest",
    "gemini-2.5-flash",
    "gemini-3.6-flash",
)
JPY_PER_USD = 150.0
PRICE_USD_PER_M = {
    "gemini-3-flash-preview": {"in": 0.50, "out": 3.00},
    "gemini-2.5-flash": {"in": 0.30, "out": 2.50},
    "gemini-3.6-flash": {"in": 1.50, "out": 7.50},
}


def _parse_cell_list(raw: Any) -> List[Tuple[int, int]]:
    out: List[Tuple[int, int]] = []
    if not isinstance(raw, list):
        return out
    for item in raw:
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            try:
                cx, cy = int(item[0]), int(item[1])
            except (TypeError, ValueError):
                continue
            if 0 <= cx <= 9 and 0 <= cy <= 9:
                out.append((cx, cy))
    return out


def _filter_roi(cells: Sequence[Tuple[int, int]]) -> Set[Tuple[int, int]]:
    return {
        (cx, cy)
        for cx, cy in cells
        if cx < P3_CATCH_ROI_COLS and cy < P3_CATCH_ROI_ROWS
    }


def _parse_ai_grid_json(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    t = re.sub(r"```\w*\n?", "", text).strip()
    start, end = t.find("{"), t.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        obj = json.loads(t[start : end + 1])
    except json.JSONDecodeError:
        return None
    if not isinstance(obj, dict):
        return None
    cells = _filter_roi(_parse_cell_list(obj.get("cells")))
    catch = _filter_roi(_parse_cell_list(obj.get("catchCells")))
    sales = _filter_roi(_parse_cell_list(obj.get("salesCells")))
    if not cells and (catch or sales):
        cells = catch | sales
    if not catch and not sales and cells:
        catch = cells
    return {
        "cells": sorted(cells),
        "catchCells": sorted(catch),
        "salesCells": sorted(sales),
    }


def build_grid_prompt(*, catch_head: str, sales_talk: str) -> str:
    catch_head = (catch_head or "").strip()
    sales_talk = (sales_talk or "").strip()
    return (
        "この画像は 1200×1200px の楽天 MAIN 画像です。10×10 の格子（列0–9左→右、行0–9上→下）に分割してください。\n"
        f"【判定対象 ROI】左上 {P3_CATCH_ROI_ROWS}行×{P3_CATCH_ROI_COLS}列（行0–{P3_CATCH_ROI_ROWS - 1}、列0–{P3_CATCH_ROI_COLS - 1}）のみ。\n"
        "【数える】後載せキャッチ見出し、セールストーク、それに付随する吹き出し・帯・装飾パネル（文字に連結した部分）。\n"
        "【数えない】金丸バッジ（右上）、ランキング表示、商品パッケージ印刷、木目背景、食品本体。\n"
        f"キャッチ見出し（参考）: 「{catch_head or '（なし）'}」\n"
        f"セールストーク（参考）: 「{sales_talk or '（なし）'}」\n"
        "各セルは [列, 行] の整数配列で返してください。\n"
        "回答は JSON のみ:\n"
        '{"cells":[[col,row],...],"catchCells":[[col,row],...],"salesCells":[[col,row],...]}'
    )


def _measured_from_cells(
    *,
    parsed: Dict[str, Any],
    catch_head: str,
    sales_talk: str,
    algo: str,
    img_name: str,
    meta_extra: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    hit_catch = {tuple(p) for p in parsed.get("catchCells") or []}
    hit_sales = {tuple(p) for p in parsed.get("salesCells") or []}
    hit = {tuple(p) for p in parsed.get("cells") or []} or (hit_catch | hit_sales)
    catch_cols, catch_rows, sales_cols, sales_rows = _p3_text_zones(
        catch_head, sales_talk
    )
    expected = _p3_expected_zone_cells(
        catch_cols, catch_rows, sales_cols, sales_rows
    )
    hit_n = len(hit)
    text_check = _p3_text_check(
        hit_catch,
        hit_sales,
        catch_head=catch_head,
        sales_talk=sales_talk,
        catch_cols=catch_cols,
        catch_rows=catch_rows,
        sales_cols=sales_cols,
        sales_rows=sales_rows,
    )
    out: Dict[str, Any] = {
        "file": img_name,
        "algo": algo,
        "hitN": hit_n,
        "over16": max(0, hit_n - P3_CATCH_TARGET_CELLS),
        "over20": max(0, hit_n - P3_TEXT_MAX_CELLS),
        "ok20": hit_n <= P3_TEXT_MAX_CELLS,
        "ok16": hit_n <= P3_CATCH_TARGET_CELLS,
        "cells": [list(p) for p in sorted(hit)],
        "catchCells": [list(p) for p in sorted(hit_catch)],
        "salesCells": [list(p) for p in sorted(hit_sales)],
        "expectedZone": [list(p) for p in sorted(expected)],
        "textCheck": text_check,
    }
    if meta_extra:
        out["apiMeta"] = meta_extra
    return out


def gemini_model_chain(preferred: str) -> List[str]:
    out: List[str] = []
    for m in (preferred,) + GEMINI_GRID_FALLBACKS:
        if m and m not in out:
            out.append(m)
    return out


def _estimate_cost_jpy(model: str, in_tok: int, out_tok: int) -> float:
    p = PRICE_USD_PER_M.get(model, PRICE_USD_PER_M["gemini-3-flash-preview"])
    usd = (in_tok * p["in"] + out_tok * p["out"]) / 1_000_000.0
    return round(usd * JPY_PER_USD, 3)


class GeminiGridMeasurer:
    """compose ループ向け。クライアントを使い回す。"""

    def __init__(self, *, model: str = DEFAULT_GEMINI_GRID_MODEL) -> None:
        self.model = model
        self._client: Any = None
        self._disabled = False
        self.total_cost_jpy = 0.0
        self.call_count = 0

    def _client_or_none(self) -> Any:
        if self._disabled:
            return None
        if self._client is not None:
            return self._client
        try:
            self._client = make_gemini_client(load_gemini_key())
            return self._client
        except SystemExit as e:
            LOG.warning("Gemini grid measurer disabled: %s", e)
            self._disabled = True
            return None

    def measure(
        self,
        img_path: Path,
        *,
        catch_head: str,
        sales_talk: str,
    ) -> Tuple[Optional[Dict[str, Any]], Dict[str, Any]]:
        meta: Dict[str, Any] = {"provider": "gemini", "model": self.model}
        client = self._client_or_none()
        if client is None:
            meta["error"] = "gemini unavailable"
            return None, meta
        prompt = build_grid_prompt(catch_head=catch_head, sales_talk=sales_talk)
        b64 = base64.b64encode(img_path.read_bytes()).decode("ascii")
        mime = mime_for(img_path)
        resp = None
        used_model = self.model
        last_err: Optional[Exception] = None
        for mid in gemini_model_chain(self.model):
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
                used_model = mid
                break
            except Exception as e:
                last_err = e
                LOG.warning("gemini grid model=%s failed: %s", mid, e)
        if resp is None:
            meta["error"] = str(last_err) if last_err else "all models failed"
            return None, meta
        meta["model"] = used_model
        text = getattr(resp, "text", None) or ""
        if not text and getattr(resp, "candidates", None):
            try:
                text = resp.candidates[0].content.parts[0].text
            except Exception:
                text = ""
        usage = getattr(resp, "usage_metadata", None)
        in_tok = int(getattr(usage, "prompt_token_count", 0) or 1430)
        out_tok = int(getattr(usage, "candidates_token_count", 0) or 120)
        meta["inputTokens"] = in_tok
        meta["outputTokens"] = out_tok
        meta["costJpy"] = _estimate_cost_jpy(used_model, in_tok, out_tok)
        self.total_cost_jpy += float(meta["costJpy"])
        self.call_count += 1
        parsed = _parse_ai_grid_json(text)
        if not parsed:
            meta["error"] = "JSON parse failed"
            meta["rawPreview"] = (text or "")[:300]
            return None, meta
        measured = _measured_from_cells(
            parsed=parsed,
            catch_head=catch_head,
            sales_talk=sales_talk,
            algo=f"gemini:{used_model}",
            img_name=img_path.name,
            meta_extra=meta,
        )
        return measured, meta


def measure_grid_dual(
    img_path: Path,
    *,
    catch_head: str,
    sales_talk: str,
    gemini: Optional[GeminiGridMeasurer] = None,
    writing_mode: str = "",
    guide_pattern: Optional[Dict[str, Any]] = None,
    guide_cells: Optional[Iterable[Tuple[int, int]]] = None,
) -> Dict[str, Any]:
    """local catch_v8b ＋ Gemini（参考）。ゲート既定=local（レビュー UI と同基準）。"""
    local = measure_grid20_p3(
        img_path,
        catch_head=catch_head,
        sales_talk=sales_talk,
        writing_mode=writing_mode,
        guide_cells=guide_cells,
        guide_pattern=guide_pattern,
    )
    gemini_m: Optional[Dict[str, Any]] = None
    gemini_meta: Dict[str, Any] = {}
    if gemini is not None:
        gemini_m, gemini_meta = gemini.measure(
            img_path, catch_head=catch_head, sales_talk=sales_talk
        )
    gate_policy = (os.environ.get("B5_GRID_GATE") or "local").strip().lower()
    if gate_policy in ("gemini", "gemini_first") and gemini_m is not None:
        gate_hit = int(gemini_m.get("hitN") or 0)
        gate_src = "gemini"
        gate_algo = gemini_m.get("algo")
    else:
        gate_hit = int(local.get("hitN") or 0)
        gate_src = "local"
        gate_algo = local.get("algo") or GRID_MEASURE_ALGO
    return {
        "local": local,
        "gemini": gemini_m,
        "geminiMeta": gemini_meta,
        "gateHitN": gate_hit,
        "gateSource": gate_src,
        "gateAlgo": gate_algo,
        "localHitN": int(local.get("hitN") or 0),
        "geminiHitN": int(gemini_m.get("hitN") or 0) if gemini_m else None,
    }


def grid_over_cap(measured: Dict[str, Any], *, max_cells: int = P3_TEXT_MAX_CELLS) -> bool:
    return int(measured.get("hitN") or 0) > max_cells


def render_grid_overlay(img_path: Path, measured: Dict[str, Any], dest: Path) -> None:
    hit = {tuple(p) for p in measured.get("cells") or []}
    expected = {tuple(p) for p in measured.get("expectedZone") or []}
    im = Image.open(img_path).convert("RGB").resize(
        (CANVAS, CANVAS), Image.Resampling.LANCZOS
    )
    w, h = im.size
    cw, ch = w / 10.0, h / 10.0
    overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    for cy in range(10):
        for cx in range(10):
            x0, y0 = int(cx * cw), int(cy * ch)
            x1, y1 = int((cx + 1) * cw) - 1, int((cy + 1) * ch) - 1
            in_roi = cx < P3_CATCH_ROI_COLS and cy < P3_CATCH_ROI_ROWS
            if (cx, cy) in expected and (cx, cy) not in hit:
                d.rectangle([x0, y0, x1, y1], fill=(255, 160, 0, 38))
            if (cx, cy) in hit:
                d.rectangle([x0, y0, x1, y1], fill=(255, 40, 40, 110))
            outline = (0, 160, 80, 200) if in_roi else (180, 180, 180, 80)
            d.rectangle([x0, y0, x1, y1], outline=outline, width=2 if in_roi else 1)
    mixed = Image.alpha_composite(im.convert("RGBA"), overlay).convert("RGB")
    dest.parent.mkdir(parents=True, exist_ok=True)
    mixed.save(dest, quality=90)


def _badge_class(hit_n: int) -> str:
    if hit_n <= P3_CATCH_TARGET_CELLS:
        return "ok"
    if hit_n <= P3_TEXT_MAX_CELLS:
        return "warn"
    return "bad"


def archive_try_with_overlays(
    *,
    try_path: Path,
    tries_dir: Path,
    theme_id: str,
    attempt: int,
    catch_head: str,
    sales_talk: str,
    dual: Dict[str, Any],
    picked: bool = False,
) -> Dict[str, Any]:
    """try 画像と local/gemini オーバレイを _tries/{theme}/ に保存（上書きしない）。"""
    sub = tries_dir / _safe_theme(theme_id)
    sub.mkdir(parents=True, exist_ok=True)
    stem = f"try{attempt}"
    raw_dest = sub / f"{stem}.jpg"
    if not raw_dest.is_file():
        import shutil

        shutil.copy2(try_path, raw_dest)
    local_ov = sub / f"{stem}_local.jpg"
    gemini_ov = sub / f"{stem}_gemini.jpg"
    render_grid_overlay(try_path, dual["local"], local_ov)
    gemini_m = dual.get("gemini")
    if gemini_m:
        render_grid_overlay(try_path, gemini_m, gemini_ov)
    return {
        "attempt": attempt,
        "picked": picked,
        "raw": raw_dest.name,
        "localOverlay": local_ov.name,
        "geminiOverlay": gemini_ov.name if gemini_m else "",
        "localHitN": dual.get("localHitN"),
        "geminiHitN": dual.get("geminiHitN"),
        "gateHitN": dual.get("gateHitN"),
        "gateSource": dual.get("gateSource"),
        "gateAlgo": dual.get("gateAlgo"),
        "gridTier": classify_grid_tier(int(dual.get("gateHitN") or 0)),
        "geminiCostJpy": (dual.get("geminiMeta") or {}).get("costJpy"),
        "dir": sub.name,
    }


def _safe_theme(theme_id: str) -> str:
    return re.sub(r"[^\w.\-]+", "_", (theme_id or "X").strip())[:32] or "X"


def build_compose_tries_html(
    *,
    out_root: Path,
    theme_rows: List[Dict[str, Any]],
    title: str,
    gemini_cost_jpy: float = 0.0,
) -> Path:
    """compose 完了後の try 比較 HTML（_tries/index.html）。"""
    tries_dir = out_root / "_tries"
    tries_dir.mkdir(exist_ok=True)
    sections: List[str] = []
    for row in theme_rows:
        tid = str(row.get("themeId") or "?")
        tries = row.get("tryArchive") or []
        if not tries:
            continue
        picked = str(row.get("pickedTry") or "")
        cards: List[str] = []
        sub_name = _safe_theme(tid)
        for tr in tries:
            att = tr.get("attempt")
            is_pick = bool(tr.get("picked")) or picked.endswith(f"try{att}.jpg")
            pick_cls = " picked" if is_pick else ""
            lh = tr.get("localHitN")
            gh = tr.get("geminiHitN")
            gh_s = str(gh) if gh is not None else "—"
            gate = tr.get("gateHitN")
            tier = tr.get("gridTier") or ""
            src = tr.get("gateSource") or ""
            base = f"{sub_name}/"
            raw = tr.get("raw") or ""
            loc = tr.get("localOverlay") or ""
            gem = tr.get("geminiOverlay") or ""
            cards.append(
                f"""
                <div class="try-card{pick_cls}">
                  <div class="try-head">try{att}{' ★採用' if is_pick else ''}</div>
                  <div class="stats">gate={gate} ({html.escape(str(src))}) tier={html.escape(str(tier))} /
                    local={lh} / gemini={gh_s}</div>
                  <div class="cols">
                    <figure><figcaption>原画</figcaption>
                      <img src="{html.escape(base + raw)}" alt="try{att}"/></figure>
                    <figure><figcaption>catch_v7</figcaption>
                      <img src="{html.escape(base + loc)}" alt="local"/></figure>
                    <figure><figcaption>Gemini</figcaption>
                      {'<img src="' + html.escape(base + gem) + '" alt="gemini"/>' if gem else '<div class="fail">—</div>'}
                    </figure>
                  </div>
                </div>"""
            )
        sections.append(
            f"<section><h2>{html.escape(tid)} — {html.escape(str(row.get('themeName') or ''))}</h2>"
            f"<p class='catch'>{html.escape(str(row.get('catch') or '')[:120])}</p>"
            f"{''.join(cards)}</section>"
        )
    body = "".join(sections) or "<p>try なし</p>"
    doc = f"""<!DOCTYPE html>
<html lang="ja"><head>
<meta charset="utf-8"/>
<title>{html.escape(title)}</title>
<style>
body{{font-family:"Yu Gothic UI",Meiryo,sans-serif;margin:16px;background:#f5f5f5}}
h1{{font-size:20px}} h2{{font-size:16px;margin-top:24px;border-bottom:1px solid #ccc}}
.note{{font-size:13px;color:#555}}
.try-card{{background:#fff;border:1px solid #ddd;margin:12px 0;padding:10px;border-radius:6px}}
.try-card.picked{{border:2px solid #1976d2;box-shadow:0 0 0 2px rgba(25,118,210,.15)}}
.try-head{{font-weight:bold;font-size:14px}}
.stats{{font-size:12px;color:#666;margin:4px 0 8px}}
.cols{{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}}
figure{{margin:0}} figcaption{{font-size:11px;color:#888}}
img{{max-width:100%;height:auto;border:1px solid #ccc}}
.fail{{padding:40px;text-align:center;color:#999;border:1px dashed #ccc}}
.catch{{font-size:12px;color:#444}}
</style></head><body>
<h1>{html.escape(title)}</h1>
<p class="note">compose try 比較。緑=ROI / 橙=期待帯 / 赤=検出。
tier: ok=≤20 / buffer=21–24（retryしない） / retry=≥25で再生成。
Gemini API 概算合計≈{gemini_cost_jpy:.2f}円</p>
{body}
</body></html>"""
    index = tries_dir / "index.html"
    index.write_text(doc, encoding="utf-8")
    return index
