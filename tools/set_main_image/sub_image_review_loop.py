# -*- coding: utf-8 -*-
"""
サブ画像 人間レビュー／再生成ループ（Vision自動QAの代替）。

- compose 出力を一覧し、チェック＋要望コメントを付けたものだけ再生成
- チェックがゼロになるまでループ（OK）
- 使い方:
  python sub_image_review_loop.py --compose-dir "...\\sub_image_ai_compose\\JAN_runId"
  → ブラウザでチェック → 「キュー保存」→ 「再生成実行」
"""
from __future__ import annotations

import argparse
import json
import logging
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urlparse

from work_paths import meta_dir_for

LOG = logging.getLogger("set_main_image.sub_image_review_loop")

PROVIDER_DIRS = {
    "openai": "03_openai",
    "gemini": "02_gemini",
    "fal": "04_fal",
}


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def list_outputs(compose_dir: Path, provider: str) -> List[Dict[str, Any]]:
    root = Path(compose_dir)
    sub = PROVIDER_DIRS.get(provider) or PROVIDER_DIRS["openai"]
    folder = root / sub
    meta_path = meta_dir_for(root) / "run_meta.json"
    jobs_by_stem: Dict[str, Any] = {}
    if meta_path.is_file():
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        for j in meta.get("jobs") or []:
            jobs_by_stem[str(j.get("stem"))] = j

    rows: List[Dict[str, Any]] = []
    if not folder.is_dir():
        return rows
    for p in sorted(folder.glob("*.jpg")):
        stem = p.stem
        job = jobs_by_stem.get(stem) or {}
        rel = f"{sub}/{p.name}".replace("\\", "/")
        rows.append(
            {
                "stem": stem,
                "provider": provider,
                "path": str(p),
                "rel": rel,
                "themeId": job.get("themeId"),
                "themeName": job.get("themeName"),
                "proposal": job.get("proposal"),
                "slotIndex": job.get("slotIndex"),
                "phaseOrder": job.get("phaseOrder"),
            }
        )
    return rows


def build_html(compose_dir: Path, provider: str, rows: List[Dict[str, Any]]) -> str:
    from review_feedback_templates import TEMPLATES, VERSION as TPL_VER

    cards = []
    for r in rows:
        cards.append(
            f"""
            <div class="card" data-stem="{r['stem']}" data-provider="{r['provider']}">
              <label class="chk"><input type="checkbox" class="need-regen"/> 再生成する</label>
              <div class="meta">S{int(r.get('slotIndex') or 0):02d}
                T{int(r.get('themeId') or 0):02d} {r.get('themeName') or ''}
                / AB{r.get('proposal') or '?'} / {r['stem']}</div>
              <img src="/file/{r['rel']}" alt="{r['stem']}"/>
              <div class="tpl-btns"></div>
              <textarea class="comment" placeholder="要望（下のテンプレボタンで挿入可）"></textarea>
            </div>
            """
        )
    cards_html = "\n".join(cards) or "<p>画像がありません。</p>"
    tpl_json = json.dumps(TEMPLATES, ensure_ascii=False)
    tpl_list_html = "".join(
        f'<button type="button" class="tpl-global" data-idx="{i}">{t["label"]}</button> '
        for i, t in enumerate(TEMPLATES)
    )
    return f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>サブ画像レビュー — {Path(compose_dir).name}</title>
<style>
body{{font-family:sans-serif;margin:16px;background:#f6f3ee;color:#222}}
h1{{font-size:18px;margin:0 0 8px}}
.toolbar{{position:sticky;top:0;background:#f6f3ee;padding:8px 0 12px;z-index:2;border-bottom:1px solid #ccc}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;margin-top:12px}}
.card{{background:#fff;border:1px solid #ddd;padding:8px;border-radius:6px}}
.card img{{width:100%;height:auto;display:block;border:1px solid #eee}}
.meta{{font-size:12px;color:#555;margin:4px 0 6px}}
textarea{{width:100%;min-height:56px;font-size:12px;box-sizing:border-box}}
button{{margin-right:8px;padding:8px 12px;cursor:pointer}}
.tpl-global,.tpl-local{{font-size:11px;padding:4px 8px;margin:2px 4px 2px 0}}
#status{{margin-left:8px;font-size:13px;color:#333}}
.hint{{font-size:12px;color:#666;margin:4px 0 0}}
.tpl-bar{{font-size:12px;margin-top:8px;padding:8px;background:#fff;border:1px solid #ddd}}
</style></head><body>
<div class="toolbar">
  <h1>サブ画像レビュー（チェック＝再生成）</h1>
  <div class="hint">dir: {compose_dir} ／ provider: {provider}<br/>
  人間の最終目視は楽天アップロードフォルダ（再生成後も自動で再export）。このUIはNG修正用。<br/>
  チェック＋要望→「キュー保存」→「再生成実行」。OKならチェック0のまま「完了」。<br/>
  コメントテンプレ v{TPL_VER}（偽物感・PACKAGE_LOCK等）</div>
  <div class="tpl-bar"><b>テンプレを最後にフォーカスした要望欄へ挿入:</b><br/>{tpl_list_html}</div>
  <p style="margin:10px 0 0">
    <button onclick="saveQueue()">キュー保存</button>
    <button onclick="runRegen()">再生成実行</button>
    <button onclick="markDone()">チェック0で完了</button>
    <button onclick="location.reload()">再読込</button>
    <span id="status"></span>
  </p>
</div>
<div class="grid">{cards_html}</div>
<script>
const TPL = {tpl_json};
let lastTa = null;
document.querySelectorAll('textarea.comment').forEach(ta=>{{
  ta.addEventListener('focus',()=>{{ lastTa=ta; }});
}});
document.querySelectorAll('.card').forEach(card=>{{
  const bar=card.querySelector('.tpl-btns');
  TPL.forEach((t,i)=>{{
    const b=document.createElement('button');
    b.type='button'; b.className='tpl-local'; b.textContent=t.label;
    b.onclick=()=>{{
      const ta=card.querySelector('.comment');
      ta.value=(ta.value?ta.value.trim()+'\\n':'')+t.text;
      card.querySelector('.need-regen').checked=true;
      lastTa=ta;
    }};
    bar.appendChild(b);
  }});
}});
document.querySelectorAll('.tpl-global').forEach(btn=>{{
  btn.onclick=()=>{{
    const t=TPL[Number(btn.dataset.idx)];
    if(!t) return;
    const ta=lastTa||document.querySelector('textarea.comment');
    if(!ta) return;
    ta.value=(ta.value?ta.value.trim()+'\\n':'')+t.text;
    const card=ta.closest('.card');
    if(card) card.querySelector('.need-regen').checked=true;
    ta.focus();
  }};
}});
async function collectItems(){{
  const cards=[...document.querySelectorAll('.card')];
  return cards.map(c=>({{
    stem:c.dataset.stem,
    provider:c.dataset.provider,
    checked:!!c.querySelector('.need-regen').checked,
    comment:(c.querySelector('.comment').value||'').trim()
  }})).filter(x=>x.checked);
}}
async function saveQueue(){{
  const items=await collectItems();
  const res=await fetch('/api/save-queue',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{items}})}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
}}
async function runRegen(){{
  await saveQueue();
  document.getElementById('status').textContent='再生成中…';
  const res=await fetch('/api/regen',{{method:'POST'}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
  if(j.ok) setTimeout(()=>location.reload(), 800);
}}
async function markDone(){{
  const items=await collectItems();
  if(items.length){{
    document.getElementById('status').textContent='まだチェックが残っています（'+items.length+'）';
    return;
  }}
  const res=await fetch('/api/done',{{method:'POST'}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||'完了';
}}
</script>
</body></html>
"""


def make_handler(compose_dir: Path, provider: str, script_dir: Path):
    root = Path(compose_dir).resolve()
    queue_path = meta_dir_for(root) / "regen_queue.json"
    done_path = meta_dir_for(root) / "review_done.json"

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt: str, *args) -> None:  # noqa: A003
            LOG.info("%s - %s", self.address_string(), fmt % args)

        def _send(self, code: int, body: bytes, content_type: str) -> None:
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            if parsed.path in ("/", "/index.html"):
                rows = list_outputs(root, provider)
                html = build_html(root, provider, rows)
                self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
                return
            if parsed.path.startswith("/file/"):
                rel = parsed.path[len("/file/") :]
                target = (root / rel).resolve()
                if not str(target).startswith(str(root)) or not target.is_file():
                    self._send(404, b"not found", "text/plain")
                    return
                data = target.read_bytes()
                ctype = "image/jpeg" if target.suffix.lower() in (".jpg", ".jpeg") else "application/octet-stream"
                self._send(200, data, ctype)
                return
            self._send(404, b"not found", "text/plain")

        def do_POST(self) -> None:  # noqa: N802
            length = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(length) if length else b"{}"
            parsed = urlparse(self.path)
            try:
                payload = json.loads(raw.decode("utf-8") or "{}")
            except Exception:
                payload = {}

            if parsed.path == "/api/save-queue":
                items = payload.get("items") or []
                queue = {
                    "composeDir": str(root),
                    "provider": provider,
                    "savedAt": datetime.now(timezone.utc).isoformat(),
                    "items": items,
                }
                queue_path.parent.mkdir(parents=True, exist_ok=True)
                queue_path.write_text(json.dumps(queue, ensure_ascii=False, indent=2), encoding="utf-8")
                msg = f"キュー保存: {len(items)} 件 → {queue_path.name}"
                self._send(200, json.dumps({"ok": True, "message": msg}, ensure_ascii=False).encode("utf-8"), "application/json")
                return

            if parsed.path == "/api/regen":
                if not queue_path.is_file():
                    self._send(
                        400,
                        json.dumps({"ok": False, "message": "キューがありません。先に保存してください。"}, ensure_ascii=False).encode("utf-8"),
                        "application/json",
                    )
                    return
                q = json.loads(queue_path.read_text(encoding="utf-8"))
                n = len([x for x in (q.get("items") or []) if x.get("checked")])
                if n == 0:
                    self._send(
                        200,
                        json.dumps({"ok": True, "message": "チェック0件。再生成不要（OK）。"}, ensure_ascii=False).encode("utf-8"),
                        "application/json",
                    )
                    return
                cmd = [
                    sys.executable,
                    str(script_dir / "sub_image_ai_compose_poc.py"),
                    "--regen-from",
                    str(root),
                    "--regen-queue",
                    str(queue_path),
                    "--providers",
                    provider,
                ]
                LOG.info("spawn regen: %s", " ".join(cmd))
                try:
                    proc = subprocess.run(cmd, cwd=str(script_dir), capture_output=True, text=True, timeout=3600)
                    ok = proc.returncode == 0
                    msg = (
                        f"再生成完了 ({n}件)" if ok else f"再生成失敗 code={proc.returncode}"
                    )
                    if proc.stderr:
                        msg += " / " + proc.stderr.strip()[-300:]
                    self._send(
                        200 if ok else 500,
                        json.dumps({"ok": ok, "message": msg}, ensure_ascii=False).encode("utf-8"),
                        "application/json",
                    )
                except Exception as e:
                    self._send(
                        500,
                        json.dumps({"ok": False, "message": str(e)}, ensure_ascii=False).encode("utf-8"),
                        "application/json",
                    )
                return

            if parsed.path == "/api/done":
                done_path.write_text(
                    json.dumps(
                        {
                            "ok": True,
                            "at": datetime.now(timezone.utc).isoformat(),
                            "composeDir": str(root),
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
                self._send(
                    200,
                    json.dumps({"ok": True, "message": "レビュー完了（チェック0）。exportへ進んでください。"}, ensure_ascii=False).encode("utf-8"),
                    "application/json",
                )
                return

            self._send(404, b"{}", "application/json")

    return Handler


def serve(compose_dir: Path, provider: str, port: int, open_browser: bool) -> int:
    script_dir = Path(__file__).resolve().parent
    handler = make_handler(compose_dir, provider, script_dir)
    httpd = ThreadingHTTPServer(("127.0.0.1", port), handler)
    url = f"http://127.0.0.1:{port}/"
    LOG.info("review UI: %s", url)
    print(url)
    if open_browser:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        LOG.info("stopped")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="サブ画像レビュー／再生成ループ UI")
    ap.add_argument("--compose-dir", type=Path, required=True)
    ap.add_argument("--provider", default="openai", choices=("openai", "gemini", "fal"))
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)
    compose = Path(args.compose_dir)
    if not compose.is_dir():
        raise SystemExit(f"compose-dir がありません: {compose}")
    return serve(compose, str(args.provider), int(args.port), open_browser=not args.no_browser)


if __name__ == "__main__":
    sys.exit(main())
