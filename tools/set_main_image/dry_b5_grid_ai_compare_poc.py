# -*- coding: utf-8 -*-
"""
100マスガイド A/B 比較 PoC（別ページ HTML）

catch_v7（ローカル） vs Gemini 2.5 Flash vs GPT-4o 高解像度を並べて比較する。

例:
  python dry_b5_grid_ai_compare_poc.py ^
    --compose-dir "G:\\...\\b5_main_base_compose\\まぐろ_20260901_202115" ^
    --themes T2,T3,T4,N1,T1 ^
    --open

  python dry_b5_grid_ai_compare_poc.py --images path\\to\\a.jpg path\\to\\b.jpg --open
"""
from __future__ import annotations

import argparse
import base64
import html
import json
import logging
import re
import webbrowser
from datetime import datetime, timezone
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Thread
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

from PIL import Image, ImageDraw

from b5_main_base_review_loop import list_compare_rows
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
from gemini_image import load_api_key as load_gemini_key
from gemini_image import make_client as make_gemini_client
from gemini_image import mime_for
from openai_image import load_openai_api_key

LOG = logging.getLogger("set_main_image.grid_ai_compare")

DEFAULT_GEMINI_MODEL = "gemini-3-flash-preview"
DEFAULT_OPENAI_MODEL = "gpt-4o"
DEFAULT_THEMES = ("T2", "T3", "T4", "N1", "T1")
JPY_PER_USD = 150.0
GEMINI_FALLBACKS = (
    "gemini-3-flash-preview",
    "gemini-flash-latest",
    "gemini-2.5-flash",
    "gemini-3.6-flash",
)

# 概算 $/1M tokens（2026-08 時点）
PRICE_USD = {
    "gemini-3-flash-preview": {"in": 0.50, "out": 3.00},
    "gemini-2.5-flash": {"in": 0.30, "out": 2.50},
    "gemini-3.6-flash": {"in": 1.50, "out": 7.50},
    "gpt-4o": {"in": 2.50, "out": 10.00},
}
# GPT-4o high detail 1200px 級 ≈ 765 image tokens（概算）
OPENAI_IMAGE_TOKENS_HIGH = 765


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _theme_from_name(name: str) -> str:
    m = re.search(r"_([NT]\d)_", name)
    return m.group(1) if m else ""


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
    start = t.find("{")
    end = t.rfind("}")
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


def _build_prompt(*, catch_head: str, sales_talk: str) -> str:
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
        "over24": max(0, hit_n - P3_TEXT_MAX_CELLS),
        "ok20": hit_n <= P3_TEXT_MAX_CELLS,
        "ok16": hit_n <= P3_CATCH_TARGET_CELLS,
        "ok25": hit_n <= 25,
        "cells": [list(p) for p in sorted(hit)],
        "catchCells": [list(p) for p in sorted(hit_catch)],
        "salesCells": [list(p) for p in sorted(hit_sales)],
        "expectedZone": [list(p) for p in sorted(expected)],
        "textCheck": text_check,
    }
    if meta_extra:
        out["apiMeta"] = meta_extra
    return out


def _estimate_cost_usd(
    *,
    provider: str,
    model: str,
    input_tokens: int,
    output_tokens: int,
) -> float:
    if provider == "openai":
        p = PRICE_USD["gpt-4o"]
        return (input_tokens * p["in"] + output_tokens * p["out"]) / 1_000_000.0
    p = PRICE_USD.get(model, PRICE_USD["gemini-2.5-flash"])
    return (input_tokens * p["in"] + output_tokens * p["out"]) / 1_000_000.0


def _gemini_model_chain(preferred: str) -> List[str]:
    out: List[str] = []
    for m in (preferred,) + GEMINI_FALLBACKS:
        if m and m not in out:
            out.append(m)
    return out


def _call_gemini(
    img_path: Path,
    *,
    catch_head: str,
    sales_talk: str,
    model: str,
) -> Tuple[Optional[Dict[str, Any]], Dict[str, Any]]:
    meta: Dict[str, Any] = {"provider": "gemini", "model": model}
    try:
        client = make_gemini_client(load_gemini_key())
    except SystemExit as e:
        meta["error"] = str(e)
        return None, meta
    prompt = _build_prompt(catch_head=catch_head, sales_talk=sales_talk)
    b64 = base64.b64encode(img_path.read_bytes()).decode("ascii")
    mime = mime_for(img_path)
    resp = None
    used_model = model
    last_err: Optional[Exception] = None
    for mid in _gemini_model_chain(model):
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
            LOG.warning("gemini model=%s failed: %s", mid, e)
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
    meta["rawPreview"] = (text or "")[:400]
    usage = getattr(resp, "usage_metadata", None)
    in_tok = int(getattr(usage, "prompt_token_count", 0) or 1430)
    out_tok = int(getattr(usage, "candidates_token_count", 0) or 120)
    meta["inputTokens"] = in_tok
    meta["outputTokens"] = out_tok
    meta["costUsd"] = _estimate_cost_usd(
        provider="gemini",
        model=used_model,
        input_tokens=in_tok,
        output_tokens=out_tok,
    )
    meta["costJpy"] = round(meta["costUsd"] * JPY_PER_USD, 3)
    parsed = _parse_ai_grid_json(text)
    if not parsed:
        meta["error"] = "JSON parse failed"
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


def _call_openai(
    img_path: Path,
    *,
    catch_head: str,
    sales_talk: str,
    model: str,
) -> Tuple[Optional[Dict[str, Any]], Dict[str, Any]]:
    meta: Dict[str, Any] = {"provider": "openai", "model": model, "detail": "high"}
    try:
        from openai import OpenAI
    except ImportError as e:
        meta["error"] = f"openai not installed: {e}"
        return None, meta
    try:
        api_key = load_openai_api_key()
    except SystemExit as e:
        meta["error"] = str(e)
        return None, meta
    prompt = _build_prompt(catch_head=catch_head, sales_talk=sales_talk)
    b64 = base64.b64encode(img_path.read_bytes()).decode("ascii")
    mime = mime_for(img_path)
    client = OpenAI(api_key=api_key)
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime};base64,{b64}",
                                "detail": "high",
                            },
                        },
                    ],
                }
            ],
            max_tokens=512,
        )
    except Exception as e:
        meta["error"] = str(e)
        LOG.warning("openai failed: %s", e)
        return None, meta
    text = ""
    try:
        text = resp.choices[0].message.content or ""
    except Exception:
        pass
    meta["rawPreview"] = (text or "")[:400]
    usage = getattr(resp, "usage", None)
    in_tok = int(getattr(usage, "prompt_tokens", 0) or 0)
    out_tok = int(getattr(usage, "completion_tokens", 0) or 120)
    if in_tok <= 0:
        in_tok = len(prompt) // 4 + OPENAI_IMAGE_TOKENS_HIGH
    meta["inputTokens"] = in_tok
    meta["outputTokens"] = out_tok
    meta["costUsd"] = _estimate_cost_usd(
        provider="openai", model=model, input_tokens=in_tok, output_tokens=out_tok
    )
    meta["costJpy"] = round(meta["costUsd"] * JPY_PER_USD, 3)
    parsed = _parse_ai_grid_json(text)
    if not parsed:
        meta["error"] = "JSON parse failed"
        return None, meta
    measured = _measured_from_cells(
        parsed=parsed,
        catch_head=catch_head,
        sales_talk=sales_talk,
        algo=f"openai:{model}:high",
        img_name=img_path.name,
        meta_extra=meta,
    )
    return measured, meta


def _jaccard(a: Set[Tuple[int, int]], b: Set[Tuple[int, int]]) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _badge_class(hit_n: int) -> str:
    if hit_n <= P3_CATCH_TARGET_CELLS:
        return "ok"
    if hit_n <= P3_TEXT_MAX_CELLS:
        return "warn"
    return "bad"


def render_overlay(img_path: Path, measured: Dict[str, Any], dest: Path) -> None:
    """レビューUIに近い色: 緑=ROI、橙=期待帯、赤=検出。"""
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
            if in_roi:
                d.rectangle(
                    [x0, y0, x1, y1],
                    outline=(0, 160, 80, 200),
                    width=2 if cx == 0 or cy == 0 else 1,
                )
            else:
                d.rectangle([x0, y0, x1, y1], outline=(180, 180, 180, 80), width=1)
    mixed = Image.alpha_composite(im.convert("RGBA"), overlay).convert("RGB")
    dest.parent.mkdir(parents=True, exist_ok=True)
    mixed.save(dest, quality=90)


def _pick_one_per_theme(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """同一 themeId は再生成版(__rN)より通常版を優先。"""
    best: Dict[str, Tuple[Tuple[int, str], Dict[str, Any]]] = {}
    for r in rows:
        tid = (r.get("themeId") or "?").upper()
        name = str(r.get("name") or "")
        rank = (1 if re.search(r"__r\d+", name) else 0, name)
        prev = best.get(tid)
        if prev is None or rank < prev[0]:
            best[tid] = (rank, r)
    return [pair[1] for _, pair in sorted(best.items(), key=lambda x: x[0])]


def _select_rows(
    compose_dir: Optional[Path],
    images: Sequence[Path],
    themes: Sequence[str],
) -> List[Dict[str, Any]]:
    if compose_dir:
        rows = list_compare_rows(compose_dir)
        if themes:
            want = {t.strip().upper() for t in themes if t.strip()}
            order = {t: i for i, t in enumerate(themes)}
            rows = [r for r in rows if (r.get("themeId") or "").upper() in want]
            rows = _pick_one_per_theme(rows)
            rows.sort(
                key=lambda r: order.get((r.get("themeId") or "").upper(), 99)
            )
        out: List[Dict[str, Any]] = []
        for r in rows:
            img = (compose_dir / str(r.get("rel") or "").replace("/", "\\")).resolve()
            if img.is_file():
                out.append({**r, "path": img})
        return out
    out = []
    for p in images:
        p = Path(p).resolve()
        if not p.is_file():
            LOG.warning("skip missing: %s", p)
            continue
        out.append(
            {
                "name": p.name,
                "rel": p.name,
                "themeId": _theme_from_name(p.name),
                "catchHead": "",
                "salesTalk": "",
                "path": p,
            }
        )
    return out


def build_html(
    *,
    title: str,
    rows: List[Dict[str, Any]],
    providers: List[Tuple[str, str]],
    summary: Dict[str, Any],
) -> str:
    provider_headers = "".join(f"<th>{html.escape(ja)}</th>" for _, ja in providers)
    body_rows = []
    for row in rows:
        cells = []
        for pid, _ in providers:
            pr = (row.get("providers") or {}).get(pid) or {}
            overlay = pr.get("overlayRel") or ""
            hit_n = pr.get("hitN")
            algo = pr.get("algo") or pid
            err = pr.get("error") or ""
            badge = _badge_class(int(hit_n or 99)) if hit_n is not None else "bad"
            if overlay:
                img_html = (
                    f'<a href="{html.escape(overlay)}" target="_blank">'
                    f'<img src="{html.escape(overlay)}" alt="{html.escape(pid)}"/></a>'
                )
                stat = (
                    f'<div class="badge {badge}">hitN={hit_n}</div>'
                    f'<div class="algo">{html.escape(str(algo))}</div>'
                )
                if pr.get("costJpy") is not None:
                    stat += f'<div class="cost">≈{pr["costJpy"]}円</div>'
            else:
                img_html = f'<div class="fail">{html.escape(err or "FAIL")}</div>'
                stat = ""
            cells.append(f"<td>{img_html}{stat}</td>")
        j = row.get("jaccard") or {}
        j_line = " / ".join(
            f"{k}={v:.0%}" for k, v in sorted(j.items()) if isinstance(v, (int, float))
        )
        body_rows.append(
            f"<tr><td class='meta'><b>{html.escape(str(row.get('themeId') or '?'))}</b>"
            f"<br/>{html.escape(str(row.get('name') or ''))}"
            f"<br/><small>{html.escape(str(row.get('catchHead') or ''))}"
            f" / {html.escape(str(row.get('salesTalk') or ''))}</small>"
            f"<br/><small>{html.escape(j_line)}</small></td>{''.join(cells)}</tr>"
        )
    cost = summary.get("costJpy") or {}
    cost_line = ", ".join(f"{k}≈{v}円" for k, v in cost.items())
    return f"""<!DOCTYPE html>
<html lang="ja"><head>
<meta charset="utf-8"/>
<title>{html.escape(title)}</title>
<style>
body{{font-family:"Yu Gothic UI",Meiryo,sans-serif;margin:16px;background:#f5f5f5;color:#222}}
h1{{font-size:20px;margin:0 0 8px}}
.note{{font-size:13px;color:#555;margin-bottom:12px}}
table{{border-collapse:collapse;width:100%;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
th,td{{border:1px solid #ddd;padding:8px;vertical-align:top}}
th{{background:#263238;color:#fff;font-size:13px}}
td img{{max-width:320px;height:auto;display:block;border:1px solid #ccc}}
.meta{{font-size:12px;min-width:140px}}
.badge{{display:inline-block;margin-top:6px;padding:2px 8px;border-radius:4px;font-size:12px;font-weight:bold}}
.badge.ok{{background:#e8f5e9;color:#063}}
.badge.warn{{background:#fff8e1;color:#a60}}
.badge.bad{{background:#fdecea;color:#b00020}}
.algo,.cost{{font-size:11px;color:#666;margin-top:2px}}
.fail{{width:280px;height:200px;display:flex;align-items:center;justify-content:center;background:#fafafa;border:2px dashed #c62828;color:#b00020;font-size:13px;padding:8px}}
.legend{{font-size:12px;margin:12px 0;padding:8px 10px;background:#fff;border:1px solid #ddd}}
</style></head><body>
<h1>{html.escape(title)}</h1>
<p class="note">緑枠=ROI(4×7) / 橙=文字長期待帯 / 赤=検出セル。合計 API 概算: {html.escape(cost_line or '—')}（1USD={JPY_PER_USD:.0f}円）</p>
<div class="legend">合格≤{P3_CATCH_TARGET_CELLS} / 注意17–{P3_TEXT_MAX_CELLS} / 不合格≥25 または textNg</div>
<table>
<thead><tr><th>画像</th>{provider_headers}</tr></thead>
<tbody>{''.join(body_rows)}</tbody>
</table>
</body></html>"""


def _serve_dir(directory: Path, port: int) -> None:
    directory = directory.resolve()

    class Handler(SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=str(directory), **kwargs)

    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    LOG.info("HTTP server http://127.0.0.1:%s/", port)
    httpd.serve_forever()


def run_compare(args: argparse.Namespace) -> Path:
    compose_dir = Path(args.compose_dir).resolve() if args.compose_dir else None
    images = [Path(p) for p in (args.images or [])]
    themes = [t.strip() for t in (args.themes or "").split(",") if t.strip()]
    if not themes and compose_dir:
        themes = list(DEFAULT_THEMES)

    rows_in = _select_rows(compose_dir, images, themes)
    if not rows_in:
        raise SystemExit("比較対象画像がありません（--compose-dir または --images を確認）")

    if compose_dir:
        out_root = compose_dir / "_grid_ai_compare" / datetime.now().strftime("%Y%m%d_%H%M%S")
    else:
        out_root = Path(__file__).resolve().parent / "_out" / "_grid_ai_compare" / datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
    out_root.mkdir(parents=True, exist_ok=True)

    providers: List[Tuple[str, str]] = [
        ("local", f"catch_v7（{GRID_MEASURE_ALGO}）"),
    ]
    if not args.skip_gemini:
        providers.append(("gemini", f"Gemini {args.gemini_model}"))
    if not args.skip_openai:
        providers.append(("openai", f"GPT-4o high"))

    result_rows: List[Dict[str, Any]] = []
    cost_total: Dict[str, float] = {"gemini": 0.0, "openai": 0.0}

    for row in rows_in:
        img_path: Path = row["path"]
        catch_head = str(row.get("catchHead") or "")
        sales_talk = str(row.get("salesTalk") or "")
        stem = Path(row.get("name") or img_path.name).stem
        prov_out: Dict[str, Any] = {}

        # local
        local = measure_grid20_p3(
            img_path, catch_head=catch_head, sales_talk=sales_talk
        )
        local_overlay = out_root / f"{stem}_local.jpg"
        render_overlay(img_path, local, local_overlay)
        prov_out["local"] = {
            "algo": local.get("algo"),
            "hitN": local.get("hitN"),
            "overlayRel": local_overlay.name,
            "cells": local.get("cells"),
            "textOk": (local.get("textCheck") or {}).get("textOk"),
        }

        if not args.skip_gemini:
            g_meas, g_meta = _call_gemini(
                img_path,
                catch_head=catch_head,
                sales_talk=sales_talk,
                model=args.gemini_model,
            )
            if g_meas:
                g_overlay = out_root / f"{stem}_gemini.jpg"
                render_overlay(img_path, g_meas, g_overlay)
                prov_out["gemini"] = {
                    "algo": g_meas.get("algo"),
                    "hitN": g_meas.get("hitN"),
                    "overlayRel": g_overlay.name,
                    "cells": g_meas.get("cells"),
                    "textOk": (g_meas.get("textCheck") or {}).get("textOk"),
                    "costJpy": g_meta.get("costJpy"),
                }
                cost_total["gemini"] += float(g_meta.get("costJpy") or 0)
            else:
                prov_out["gemini"] = {"error": g_meta.get("error") or "failed"}

        if not args.skip_openai:
            o_meas, o_meta = _call_openai(
                img_path,
                catch_head=catch_head,
                sales_talk=sales_talk,
                model=args.openai_model,
            )
            if o_meas:
                o_overlay = out_root / f"{stem}_openai.jpg"
                render_overlay(img_path, o_meas, o_overlay)
                prov_out["openai"] = {
                    "algo": o_meas.get("algo"),
                    "hitN": o_meas.get("hitN"),
                    "overlayRel": o_overlay.name,
                    "cells": o_meas.get("cells"),
                    "textOk": (o_meas.get("textCheck") or {}).get("textOk"),
                    "costJpy": o_meta.get("costJpy"),
                }
                cost_total["openai"] += float(o_meta.get("costJpy") or 0)
            else:
                prov_out["openai"] = {"error": o_meta.get("error") or "failed"}

        sets: Dict[str, Set[Tuple[int, int]]] = {}
        for pid, pr in prov_out.items():
            if pr.get("cells") is not None:
                sets[pid] = {tuple(c) for c in pr["cells"]}
        jaccard: Dict[str, float] = {}
        if "local" in sets and "gemini" in sets:
            jaccard["local∩gemini"] = _jaccard(sets["local"], sets["gemini"])
        if "local" in sets and "openai" in sets:
            jaccard["local∩openai"] = _jaccard(sets["local"], sets["openai"])
        if "gemini" in sets and "openai" in sets:
            jaccard["gemini∩openai"] = _jaccard(sets["gemini"], sets["openai"])

        result_rows.append({**row, "providers": prov_out, "jaccard": jaccard})
        LOG.info(
            "done %s local=%s gemini=%s openai=%s",
            stem,
            prov_out.get("local", {}).get("hitN"),
            prov_out.get("gemini", {}).get("hitN"),
            prov_out.get("openai", {}).get("hitN"),
        )

    json_rows = []
    for row in result_rows:
        jr = {k: v for k, v in row.items() if k != "path"}
        json_rows.append(jr)
    summary = {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "composeDir": str(compose_dir) if compose_dir else None,
        "themes": themes,
        "costJpy": {k: round(v, 3) for k, v in cost_total.items() if v > 0},
        "rows": json_rows,
    }
    summary_path = out_root / "summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    title = f"100マスガイド A/B — {compose_dir.name if compose_dir else 'images'}"
    html_text = build_html(
        title=title,
        rows=result_rows,
        providers=providers,
        summary=summary,
    )
    index_path = out_root / "index.html"
    index_path.write_text(html_text, encoding="utf-8")
    LOG.info("wrote %s", index_path)
    return out_root


def main() -> None:
    ap = argparse.ArgumentParser(description="100マスガイド catch_v7 vs AI 比較 PoC")
    ap.add_argument("--compose-dir", help="b5_main_base_compose フォルダ")
    ap.add_argument("--images", nargs="*", help="単体画像パス（compose-dir なし時）")
    ap.add_argument(
        "--themes",
        default=",".join(DEFAULT_THEMES),
        help=f"compose-dir 時のテーマ filter（既定: {','.join(DEFAULT_THEMES)}）",
    )
    ap.add_argument("--gemini-model", default=DEFAULT_GEMINI_MODEL)
    ap.add_argument("--openai-model", default=DEFAULT_OPENAI_MODEL)
    ap.add_argument("--skip-gemini", action="store_true")
    ap.add_argument("--skip-openai", action="store_true")
    ap.add_argument("--open", action="store_true", help="生成後ブラウザで開く")
    ap.add_argument(
        "--serve",
        type=int,
        default=0,
        metavar="PORT",
        help="簡易HTTPサーバで配信（例: 8870）。--open と併用可",
    )
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()
    _setup_log(args.verbose)
    if not args.compose_dir and not args.images:
        ap.error("--compose-dir または --images が必要です")
    out_root = run_compare(args)
    url = f"http://127.0.0.1:{args.serve}/" if args.serve else out_root / "index.html"
    if args.serve:
        if args.open:
            port = args.serve

            def _open_delayed() -> None:
                import time

                time.sleep(0.4)
                webbrowser.open(f"http://127.0.0.1:{port}/")

            Thread(target=_open_delayed, daemon=True).start()
        _serve_dir(out_root, args.serve)
    elif args.open:
        webbrowser.open((out_root / "index.html").as_uri())


if __name__ == "__main__":
    main()
