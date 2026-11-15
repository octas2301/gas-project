# -*- coding: utf-8 -*-
"""B-⑤ キャッチ置き場パターン — Web ピッカー（レ点採用）。

  python b5_catch_zone_picker.py
  python b5_catch_zone_picker.py --serve --port 8870
"""
from __future__ import annotations

import argparse
import json
import logging
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import List

from b5_catch_zone_patterns import (
    ANCHORS,
    GOLD_COLS,
    MAX_CELLS,
    MIN_CELLS,
    OCTAS_ROWS,
    USABLE_COLS,
    USABLE_ROWS,
    CANVAS_ROWS,
    TATE_MAX_COLS,
    CatchPattern,
    generate_all_patterns,
)
from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT

LOG = logging.getLogger("set_main_image.b5_catch_zone_picker")

DEFAULT_OUT = (
    Path(DEFAULT_WORK_ROOT)
    / SUB_TEST_OUT
    / "b5_main_base_compose"
    / "_catch_zone_picker"
)


def _html(patterns: List[CatchPattern], out_dir: Path) -> str:
    data = {
        "minCells": MIN_CELLS,
        "maxCells": MAX_CELLS,
        "usableCols": USABLE_COLS,
        "usableRows": USABLE_ROWS,
        "canvasRows": CANVAS_ROWS,
        "tateMaxCols": TATE_MAX_COLS,
        "tateAnchor": "origin",
        "goldCols": sorted(GOLD_COLS),
        "octasRows": sorted(OCTAS_ROWS),
        "anchors": {k: v[2] for k, v in ANCHORS.items()},
        "patterns": [p.to_dict() for p in patterns],
    }
    payload = json.dumps(data, ensure_ascii=False)
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>B-⑤ キャッチ置き場パターン採用</title>
<style>
:root {{
  --catch: rgba(200,40,40,.55);
  --sales: rgba(40,90,200,.5);
  --deco: rgba(140,60,180,.45);
  --gold: rgba(255,200,0,.35);
  --octas: rgba(255,140,0,.25);
  --margin: rgba(180,180,180,.25);
  --roi: rgba(0,200,80,.12);
  --range: rgba(120,80,20,.45);
  --grid: rgba(0,120,60,.35);
}}
* {{ box-sizing: border-box; }}
body {{ font-family: "Yu Gothic UI", Meiryo, sans-serif; margin: 0; background: #f0ede8; color: #222; }}
header {{ background: #fff; border-bottom: 1px solid #ddd; padding: 12px 16px; position: sticky; top: 0; z-index: 10; }}
h1 {{ font-size: 18px; margin: 0 0 6px; }}
.sub {{ font-size: 13px; color: #555; line-height: 1.5; }}
.toolbar {{ display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin-top: 10px; }}
.toolbar label {{ font-size: 13px; }}
select, button {{ font-size: 13px; padding: 6px 10px; }}
button.primary {{ background: #1565c0; color: #fff; border: none; border-radius: 4px; cursor: pointer; }}
button.primary:hover {{ background: #0d47a1; }}
#status {{ font-size: 13px; color: #333; margin-left: 8px; }}
main {{ padding: 16px; max-width: 1400px; margin: 0 auto; }}
.legend {{ display: flex; flex-wrap: wrap; gap: 12px; font-size: 12px; margin-bottom: 16px; }}
.legend span {{ display: inline-flex; align-items: center; gap: 4px; }}
.swatch {{ width: 14px; height: 14px; border: 1px solid #999; display: inline-block; }}
.grid-cards {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 12px; }}
.card {{
  background: #fff; border: 2px solid #ddd; border-radius: 8px; padding: 10px;
  cursor: pointer; transition: border-color .15s, box-shadow .15s;
}}
.card:hover {{ border-color: #90caf9; }}
.card.selected {{ border-color: #1565c0; box-shadow: 0 0 0 2px rgba(21,101,192,.2); }}
.card-head {{ display: flex; align-items: flex-start; gap: 8px; margin-bottom: 8px; }}
.card-head input {{ margin-top: 3px; width: 18px; height: 18px; cursor: pointer; }}
.card-title {{ font-weight: bold; font-size: 13px; line-height: 1.3; }}
.card-meta {{ font-size: 11px; color: #666; margin-top: 2px; }}
.mini-grid {{
  display: grid; grid-template-columns: repeat(10, 1fr); grid-template-rows: repeat(10, 1fr);
  aspect-ratio: 1; border: 1px solid #bbb; background: #faf8f4;
}}
.cell {{ border: 1px solid var(--grid); min-height: 0; }}
.cell.gold {{ background: var(--gold); }}
.cell.octas {{ background: var(--octas); }}
.cell.roi {{ background: var(--roi); }}
.cell.tate-band {{ background: rgba(0,200,80,.06); }}
.cell.margin {{ background: repeating-linear-gradient(45deg, var(--margin), var(--margin) 2px, transparent 2px, transparent 6px); }}
.cell.catch {{ background: var(--catch); }}
.cell.sales {{ background: var(--sales); }}
.cell.deco {{ background: var(--deco); }}
.cell.range {{ background: var(--range); }}
.card-note {{ font-size: 11px; color: #777; margin-top: 6px; min-height: 2.4em; }}
footer {{
  position: sticky; bottom: 0; background: #fff; border-top: 1px solid #ddd;
  padding: 12px 16px; display: flex; flex-wrap: wrap; gap: 10px; align-items: center;
}}
#selectedList {{ font-size: 12px; color: #444; flex: 1; min-width: 200px; max-height: 4em; overflow: auto; }}
</style>
</head>
<body>
<header>
  <h1>B-⑤ キャッチ置き場パターン（16〜20マス・金丸除く）</h1>
  <div class="sub">
    採用候補は <b>16〜20マス</b> に限定（金丸セット数4マスを除くAI側の目安）。
    <b>レ点（チェック）</b>で採用パターンを選び、JSON を保存してください。
    1マス = 120×120px（1200px キャンバス）。<br/>
    <b>ルール:</b>
    ① 赤（キャッチ）と青（セールス）は<b>隣接必須</b>（両方ある場合）。
    ② キャッチのマス数 &gt; セールスのマス数。枠は<b>目安のみ</b>（実デザインを枠内に収める義務は設けない）。
    ③ <b>縦書き</b>は<b>左上起点</b>・<b>左3列</b>・<b>茶色=範囲のみ</b>（赤青の役割分割は指定しない）。横書きは赤/青で分割。
  </div>
  <div class="toolbar">
    <label>起点フィルタ
      <select id="filterAnchor">
        <option value="">すべて</option>
      </select>
    </label>
    <label>書字方向
      <select id="filterWritingMode">
        <option value="">すべて</option>
        <option value="horizontal-tb">横書き</option>
        <option value="vertical-rl">縦書き</option>
      </select>
    </label>
    <label>向き（レイアウト）
      <select id="filterOrient">
        <option value="">すべて</option>
        <option value="horizontal">横組み</option>
        <option value="vertical">縦組み</option>
        <option value="hybrid">ハイブリッド</option>
      </select>
    </label>
    <label>緑枠
      <select id="filterExtendsRoi">
        <option value="">すべて</option>
        <option value="beyond">緑枠外含むのみ</option>
        <option value="within">緑枠内のみ</option>
      </select>
    </label>
    <label>マス数
      <select id="filterCellN">
        <option value="">16〜20 すべて</option>
        <option value="16">16マス</option>
        <option value="17">17マス</option>
        <option value="18">18マス</option>
        <option value="19">19マス</option>
        <option value="20">20マス</option>
      </select>
    </label>
    <button type="button" id="btnSelectVisible">表示中を全選択</button>
    <button type="button" id="btnClear">全解除</button>
    <span id="status"></span>
  </div>
</header>
<main>
  <div class="legend">
    <span><i class="swatch" style="background:var(--catch)"></i>横書き・キャッチ</span>
    <span><i class="swatch" style="background:var(--sales)"></i>横書き・セールス</span>
    <span><i class="swatch" style="background:var(--range)"></i>縦書き・範囲のみ</span>
    <span><i class="swatch" style="background:var(--gold)"></i>金丸聖域</span>
    <span><i class="swatch" style="background:var(--margin)"></i>起点マージン（空け）</span>
    <span><i class="swatch" style="background:var(--roi)"></i>横書き目安（緑枠4×7）</span>
    <span><i class="swatch" style="background:rgba(0,200,80,.06);border:1px solid #6a6"></i>縦型帯（左3列）</span>
  </div>
  <div id="cards" class="grid-cards"></div>
</main>
<footer>
  <strong>採用 <span id="selCount">0</span> 件</strong>
  <div id="selectedList"></div>
  <button type="button" class="primary" id="btnDownload">採用JSONをダウンロード</button>
  <button type="button" class="primary" id="btnSaveDrive">採用JSONをDriveフォルダに保存</button>
  <button type="button" id="btnCopy">JSONをコピー</button>
</footer>
<script>
const DATA = {payload};
const STORAGE_KEY = "b5_catch_zone_picker_v6_range_cellkey";
const selected = new Set(JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]"));

function cellKey(p) {{
  return (p.cells || []).map(c => c[0] + "," + c[1]).sort().join(";");
}}

function patternByCellKey() {{
  const m = new Map();
  for (const p of DATA.patterns) m.set(cellKey(p), p);
  return m;
}}

function migrateSelection() {{
  const byId = new Map(DATA.patterns.map(p => [p.id, p]));
  const byCell = patternByCellKey();
  const next = new Set();
  for (const token of selected) {{
    if (byCell.has(token)) next.add(token);
    else if (byId.has(token)) next.add(cellKey(byId.get(token)));
  }}
  selected.clear();
  for (const k of next) selected.add(k);
  localStorage.setItem(STORAGE_KEY, JSON.stringify([...selected]));
}}
migrateSelection();

function anchorMarginCells(anchor) {{
  const m = new Set();
  if (anchor === "inset_rc" || anchor === "inset_r") {{
    for (let c = 0; c < 10; c++) m.add(c + ",0");
  }}
  if (anchor === "inset_rc" || anchor === "inset_c") {{
    for (let r = 0; r < 10; r++) m.add("0," + r);
  }}
  return m;
}}

function zoneMap(p) {{
  const map = new Map();
  if (p.zoneMode === "range-only") {{
    for (const c of (p.cells || [])) map.set(c[0] + "," + c[1], "range");
    return map;
  }}
  for (const z of p.zones) {{
    for (let dr = 0; dr < z.h; dr++) {{
      for (let dc = 0; dc < z.w; dc++) {{
        map.set((z.c0 + dc) + "," + (z.r0 + dr), z.role);
      }}
    }}
  }}
  return map;
}}

function renderMiniGrid(p) {{
  const zm = zoneMap(p);
  const margin = anchorMarginCells(p.anchor);
  let html = '<div class="mini-grid">';
  for (let r = 0; r < 10; r++) {{
    for (let c = 0; c < 10; c++) {{
      const key = c + "," + r;
      let cls = "cell";
      if (DATA.goldCols.includes(c) && r <= 2) cls += " gold";
      else if (DATA.octasRows.includes(r) && c >= 7) cls += " octas";
      else if (margin.has(key)) cls += " margin";
      else if (c < DATA.tateMaxCols && r < DATA.canvasRows && !zm.has(key)) cls += " tate-band";
      else if (c < DATA.usableCols && r < DATA.usableRows && !zm.has(key)) cls += " roi";
      const role = zm.get(key);
      if (role === "catch") cls += " catch";
      else if (role === "sales") cls += " sales";
      else if (role === "deco") cls += " deco";
      else if (role === "range") cls += " range";
      html += '<div class="' + cls + '"></div>';
    }}
  }}
  html += '</div>';
  return html;
}}

function zoneSummary(p) {{
  if (p.zoneMode === "range-only") {{
    const z = p.zones[0];
    return "range " + z.label + " · " + p.cellN + "マス";
  }}
  return p.zones.map(z => z.role + " " + z.label).join(" / ");
}}

function extendsRoi(p) {{
  for (const z of p.zones) {{
    if (z.r0 + z.h > DATA.usableRows || z.c0 + z.w > DATA.usableCols) return true;
  }}
  return false;
}}

function filteredPatterns() {{
  const a = document.getElementById("filterAnchor").value;
  const wm = document.getElementById("filterWritingMode").value;
  const o = document.getElementById("filterOrient").value;
  const ext = document.getElementById("filterExtendsRoi").value;
  const exactN = document.getElementById("filterCellN").value;
  return DATA.patterns.filter(p =>
    p.cellN >= DATA.minCells &&
    p.cellN <= DATA.maxCells &&
    (!a || p.anchor === a) &&
    (!wm || p.writingMode === wm) &&
    (!o || p.orientation === o) &&
    (!ext || (ext === "beyond" ? extendsRoi(p) : !extendsRoi(p))) &&
    (!exactN || p.cellN === Number(exactN))
  );
}}

function renderCards() {{
  const list = filteredPatterns();
  const root = document.getElementById("cards");
  root.innerHTML = "";
  for (const p of list) {{
    const ck = cellKey(p);
    const card = document.createElement("div");
    card.className = "card" + (selected.has(ck) ? " selected" : "");
    card.dataset.cellKey = ck;
    const zones = zoneSummary(p);
    const wmLabel = p.writingMode === "vertical-rl" ? "縦書き" : "横書き";
    const ratio = p.zoneMode === "range-only"
      ? (" · 範囲" + p.cellN + "マス")
      : (p.salesN > 0 ? (" · 赤" + p.catchN + " &gt; 青" + p.salesN) : (" · 赤" + p.catchN));
    const beyond = extendsRoi(p) ? " · 緑枠外" : "";
    card.innerHTML =
      '<div class="card-head">' +
        '<input type="checkbox" class="pick"' + (selected.has(ck) ? " checked" : "") + '/>' +
        '<div><div class="card-title">' + p.id + '</div>' +
        '<div class="card-meta">' + p.anchorLabel + ' · ' + wmLabel + ' · <b>' + p.cellN + 'マス</b>' + ratio + beyond + '</div></div>' +
      '</div>' +
      renderMiniGrid(p) +
      '<div class="card-note">' + zones + (p.note ? '<br/>' + p.note : '') + '</div>';
    const toggle = () => {{
      if (selected.has(ck)) selected.delete(ck); else selected.add(ck);
      localStorage.setItem(STORAGE_KEY, JSON.stringify([...selected]));
      renderCards();
      updateFooter();
    }};
    card.querySelector(".pick").addEventListener("click", e => {{ e.stopPropagation(); toggle(); }});
    card.addEventListener("click", toggle);
    root.appendChild(card);
  }}
  document.getElementById("status").textContent = "表示 " + list.length + " / 全 " + DATA.patterns.length;
}}

function updateFooter() {{
  document.getElementById("selCount").textContent = String(selected.size);
  const keys = [...selected];
  document.getElementById("selectedList").textContent = keys.length ? keys.join(" | ") : "（未選択）";
}}

function exportJson() {{
  const byCell = patternByCellKey();
  const picked = [...selected].map(k => byCell.get(k)).filter(Boolean);
  return {{
    version: 2,
    note: "縦書きは zoneMode=range-only（範囲のみ。キャッチ/セールス分割は指定しない）",
    pickedAt: new Date().toISOString(),
    minCells: DATA.minCells,
    maxCells: DATA.maxCells,
    patterns: picked,
  }};
}}

async function saveToDrive() {{
  const body = exportJson();
  if (!body.patterns.length) {{ alert("採用が0件です"); return; }}
  try {{
    const res = await fetch("/api/save-adopted", {{
      method: "POST",
      headers: {{ "Content-Type": "application/json" }},
      body: JSON.stringify(body),
    }});
    if (!res.ok) throw new Error(await res.text());
    const j = await res.json();
    alert("保存しました: " + j.path);
  }} catch (e) {{
    alert("Drive保存は --serve 起動時のみ。ダウンロードして _catch_zone_picker/ に置いてください。");
  }}
}}

document.getElementById("filterAnchor").addEventListener("change", renderCards);
document.getElementById("filterWritingMode").addEventListener("change", renderCards);
document.getElementById("filterExtendsRoi").addEventListener("change", renderCards);
document.getElementById("filterOrient").addEventListener("change", renderCards);
document.getElementById("filterCellN").addEventListener("change", renderCards);
document.getElementById("btnClear").addEventListener("click", () => {{
  selected.clear();
  localStorage.setItem(STORAGE_KEY, "[]");
  renderCards(); updateFooter();
}});
document.getElementById("btnSelectVisible").addEventListener("click", () => {{
  for (const p of filteredPatterns()) selected.add(cellKey(p));
  localStorage.setItem(STORAGE_KEY, JSON.stringify([...selected]));
  renderCards(); updateFooter();
}});
document.getElementById("btnSaveDrive").addEventListener("click", saveToDrive);
document.getElementById("btnDownload").addEventListener("click", () => {{
  const blob = new Blob([JSON.stringify(exportJson(), null, 2)], {{type: "application/json"}});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "b5_catch_zone_adopted.json";
  a.click();
}});
document.getElementById("btnCopy").addEventListener("click", async () => {{
  await navigator.clipboard.writeText(JSON.stringify(exportJson(), null, 2));
  alert("採用JSONをコピーしました");
}});

for (const [k, label] of Object.entries(DATA.anchors)) {{
  const opt = document.createElement("option");
  opt.value = k;
  opt.textContent = label;
  document.getElementById("filterAnchor").appendChild(opt);
}}
renderCards();
updateFooter();
</script>
</body>
</html>"""


def write_picker(out_dir: Path) -> Path:
    patterns = generate_all_patterns()
    out_dir.mkdir(parents=True, exist_ok=True)
    html_path = out_dir / "index.html"
    html_path.write_text(_html(patterns, out_dir), encoding="utf-8")
    meta = {
        "minCells": MIN_CELLS,
        "maxCells": MAX_CELLS,
        "patternCount": len(patterns),
        "byAnchor": {k: len([p for p in patterns if p.anchor == k]) for k in ANCHORS},
        "byExtendsRoi": {
            "beyond": len(
                [
                    p
                    for p in patterns
                    if any(
                        z.r0 + z.h > USABLE_ROWS or z.c0 + z.w > USABLE_COLS
                        for z in p.zones
                    )
                ]
            ),
        },
        "byWritingMode": {
            wm: len([p for p in patterns if p.writing_mode == wm])
            for wm in ("horizontal-tb", "vertical-rl")
        },
        "byCellN": {
            str(n): len([p for p in patterns if p.cell_n == n])
            for n in range(MIN_CELLS, MAX_CELLS + 1)
        },
    }
    (out_dir / "patterns_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_dir / "patterns_all.json").write_text(
        json.dumps([p.to_dict() for p in patterns], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    LOG.info("patterns=%s out=%s", len(patterns), html_path)
    return html_path


def serve(out_dir: Path, port: int) -> None:
    from b5_catch_zone_adopted import save_adopted
    from urllib.parse import urlparse

    out_dir = Path(out_dir)

    class Handler(SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=str(out_dir), **kwargs)

        def do_POST(self) -> None:
            if urlparse(self.path).path != "/api/save-adopted":
                self.send_error(404)
                return
            try:
                length = int(self.headers.get("Content-Length", 0))
                raw = self.rfile.read(length).decode("utf-8")
                data = json.loads(raw)
                if not isinstance(data, dict):
                    raise ValueError("body must be object")
                saved = save_adopted(data, out_dir)
                payload = json.dumps(
                    {"ok": True, "path": str(saved)},
                    ensure_ascii=False,
                ).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            except Exception as exc:
                LOG.exception("save-adopted")
                msg = json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False)
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(msg.encode("utf-8"))

        def log_message(self, fmt: str, *args) -> None:
            LOG.info("%s - %s", self.address_string(), fmt % args)

    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    LOG.info("serve http://127.0.0.1:%s/", port)
    httpd.serve_forever()


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    ap = argparse.ArgumentParser(description="キャッチ置き場パターン Web ピッカー")
    ap.add_argument("--out-dir", type=Path, default=DEFAULT_OUT)
    ap.add_argument("--serve", action="store_true", help="生成後に HTTP サーバ起動")
    ap.add_argument("--port", type=int, default=8870)
    ap.add_argument("--open", action="store_true", help="既定ブラウザで開く")
    ap.add_argument(
        "--import-adopted",
        type=Path,
        help="ダウンロード済み b5_catch_zone_adopted.json を Drive ピッカーフォルダへコピー",
    )
    args = ap.parse_args()
    out_dir = Path(args.out_dir)
    if args.import_adopted:
        from b5_catch_zone_adopted import save_adopted

        src = Path(args.import_adopted)
        data = json.loads(src.read_text(encoding="utf-8"))
        saved = save_adopted(data, out_dir)
        print(saved)
        return 0
    html_path = write_picker(out_dir)
    print(html_path)
    if args.open:
        import os

        os.startfile(str(html_path))  # noqa: S606 — Windows
    if args.serve:
        serve(out_dir, int(args.port))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
