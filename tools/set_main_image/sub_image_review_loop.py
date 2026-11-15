# -*- coding: utf-8 -*-
"""
サブ画像 人間レビュー／再生成ループ（Vision自動QAの代替）。

P-align-2a: lightbox／採用／フォルダ開く／再生成世代保持（__rN）に寄せる。

  python sub_image_review_loop.py --compose-dir "...\\sub_image_ai_compose\\JAN_runId"
"""
from __future__ import annotations

import argparse
import html
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import unquote, urlparse

from work_paths import DEFAULT_RAKUTEN_UPLOAD_DIR, meta_dir_for

LOG = logging.getLogger("set_main_image.sub_image_review_loop")

PROVIDER_DIRS = {
    "openai": "03_openai",
    "gemini": "02_gemini",
    "fal": "04_fal",
}
_STAGING_PREFIX = "__regen_staging_"
_R_SUFFIX = re.compile(r"__r(\d+)$", re.I)


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _base_stem(stem: str) -> str:
    return _R_SUFFIX.sub("", stem)


def _purge_staging(folder: Path) -> None:
    if not folder.is_dir():
        return
    for p in list(folder.glob(f"{_STAGING_PREFIX}*")):
        try:
            p.unlink()
        except Exception:
            LOG.warning("staging delete fail: %s", p)


def _message_by_theme(root: Path) -> Dict[int, Dict[str, Any]]:
    path = root / "messages.json"
    out: Dict[int, Dict[str, Any]] = {}
    if not path.is_file():
        return out
    try:
        msg = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return out
    for row in msg.get("rows") or []:
        try:
            tid = int(row.get("themeId") or 0)
        except (TypeError, ValueError):
            continue
        if tid:
            out[tid] = row
    return out


def list_outputs(compose_dir: Path, provider: str) -> List[Dict[str, Any]]:
    root = Path(compose_dir)
    sub = PROVIDER_DIRS.get(provider) or PROVIDER_DIRS["openai"]
    folder = root / sub
    _purge_staging(folder)
    messages = _message_by_theme(root)
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
        if p.name.startswith(_STAGING_PREFIX):
            continue
        stem = p.stem
        job = jobs_by_stem.get(stem) or jobs_by_stem.get(_base_stem(stem)) or {}
        theme_m = re.match(r"T(\d+)", stem, re.I)
        try:
            tid = int(theme_m.group(1)) if theme_m else int(job.get("themeId") or 0)
        except (TypeError, ValueError):
            tid = 0
        msg = messages.get(tid) or {}
        rel = f"{sub}/{p.name}".replace("\\", "/")
        gen = ""
        m = _R_SUFFIX.search(stem)
        if m:
            gen = f"再生成r{m.group(1)}"
        rows.append(
            {
                "stem": stem,
                "baseStem": _base_stem(stem),
                "provider": provider,
                "path": str(p),
                "rel": rel,
                "themeId": tid or job.get("themeId"),
                "themeName": job.get("themeName") or msg.get("name") or "",
                "catch": msg.get("catch") or "",
                "proposal": job.get("proposal"),
                "slotIndex": job.get("slotIndex"),
                "phaseOrder": job.get("phaseOrder"),
                "genLabel": gen,
            }
        )
    return rows


def prefer_adopted_jobs(meta: Dict[str, Any], adopted_stems: List[str]) -> Dict[str, Any]:
    """採用した stem のジョブだけ残す。選んでいないテーマは確認フォルダへ出さない。"""
    want = {str(s).strip() for s in adopted_stems if str(s).strip()}
    if not want:
        return meta
    bases = {_base_stem(s) for s in want}
    jobs = list(meta.get("jobs") or [])
    exact = []
    related = []
    for j in jobs:
        stem = str(j.get("stem") or "")
        if stem in want:
            exact.append(j)
        elif stem in bases or _base_stem(stem) in bases:
            related.append(j)
    kept = related + exact
    meta = dict(meta)
    meta["jobs"] = kept
    meta["adoptedStems"] = sorted(want)
    meta["adoptedAt"] = datetime.now(timezone.utc).isoformat()
    LOG.info("adopt filter jobs=%s kept=%s", len(jobs), len(kept))
    return meta


def _export_theme_new_adopted(root: Path, stems: List[str]) -> str:
    """採用したテーマ別の絵だけを、確認フォルダへ出す。"""
    msgj = json.loads((root / "messages.json").read_text(encoding="utf-8"))
    jan = str(msgj.get("jan") or "").strip()
    if not jan:
        raise RuntimeError("商品番号がありません")
    from rakuten_image_names import sanitize_pattern_slug, sub_filename
    from sub_image_mix_manifest import cards_from_theme_new

    cards, _name = cards_from_theme_new(root, jan=jan)
    want = {s for s in stems}
    best: Dict[int, Dict[str, Any]] = {}
    for c in cards:
        if str(c.get("stem") or "") not in want:
            continue
        tid = int(c.get("themeId") or 0)
        if not tid:
            continue
        prev = best.get(tid)
        if not prev or int(c.get("gen") or 1) >= int(prev.get("gen") or 1):
            best[tid] = c
    ordered = sorted(best.values(), key=lambda x: (int(x.get("phaseOrder") or 0), int(x.get("themeId") or 0)))
    if not ordered:
        raise RuntimeError("選んだ絵がフォルダにありません")
    out_dir = Path(DEFAULT_RAKUTEN_UPLOAD_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    from b5_skip_generate_finish import (
        child_skus_for_image_bind,
        represent_child_skus_for_layout_export,
        retire_product_subs,
    )

    # Cレイアウト: 親あたり1子SKU名だけ（全子複製すると未整理・候補へ溢れる）
    skus = represent_child_skus_for_layout_export(jan, out_dir=out_dir)
    if not skus:
        raise RuntimeError("子SKUがありません。マスタのこの商品に子SKUが無いため、確認フォルダへ出せません")
    moved = retire_product_subs(out_dir, jan)
    for old in child_skus_for_image_bind(jan, out_dir):
        moved.extend(retire_product_subs(out_dir, old))
    names: List[str] = []
    for i, c in enumerate(ordered[:10], start=1):
        src = Path(str(c.get("path") or ""))
        if not src.is_file():
            continue
        slug = sanitize_pattern_slug(c.get("themeSlug") or c.get("themeName") or "")
        for sku in skus:
            dest = out_dir / sub_filename(sku, i, pattern=slug)
            shutil.copy2(src, dest)
            names.append(dest.name)
            LOG.info("adopt copy jan=%s sku=%s theme=%s dest=%s", jan, sku, c.get("themeId"), dest.name)
    LOG.info(
        "adopt export jan=%s themes=%s skus=%s files=%s retired=%s",
        jan,
        len(ordered),
        len(skus),
        len(names),
        len(moved),
    )
    return (
        f"採用 {len(ordered)} テーマを、親あたり1子SKU（{len(skus)}件）名で確認フォルダへ出しました。"
        f"前のサブ {len(moved)} 枚は外しました"
    )


def build_html(compose_dir: Path, provider: str, rows: List[Dict[str, Any]]) -> str:
    from review_feedback_templates import TEMPLATES, VERSION as TPL_VER

    cards = []
    for r in rows:
        gen = f" <b style=\"color:#0653c7\">{html.escape(str(r.get('genLabel') or ''))}</b>" if r.get("genLabel") else ""
        catch = str(r.get("catch") or "").strip()
        theme = str(r.get("themeName") or "").strip()
        if catch:
            title = f"<b>{html.escape(catch)}</b><br/>{html.escape(theme)}"
        else:
            title = (
                f"S{int(r.get('slotIndex') or 0):02d} "
                f"T{int(r.get('themeId') or 0):02d} {html.escape(theme)} "
                f"/ AB{html.escape(str(r.get('proposal') or '?'))} / {html.escape(str(r['stem']))}"
            )
        cards.append(
            f"""
            <div class="card" data-stem="{html.escape(str(r['stem']), quote=True)}" data-provider="{html.escape(str(r['provider']), quote=True)}">
              <label class="chk"><input type="checkbox" class="need-regen"/> 再生成する</label>
              <label class="chk adopt"><input type="checkbox" class="adopt-chk"/> 採用する（出口へ）</label>
              <div class="meta">{title}{gen}</div>
              <img class="thumb" src="/file/{r['rel']}" alt="{html.escape(catch or str(r['stem']))}" title="クリックで拡大"/>
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
.card img.thumb{{width:100%;height:auto;display:block;border:1px solid #eee;cursor:zoom-in}}
.meta{{font-size:12px;color:#555;margin:4px 0 6px}}
textarea{{width:100%;min-height:56px;font-size:12px;box-sizing:border-box}}
button{{margin-right:8px;padding:8px 12px;cursor:pointer}}
.tpl-global,.tpl-local{{font-size:11px;padding:4px 8px;margin:2px 4px 2px 0}}
#status{{margin-left:8px;font-size:13px;color:#333}}
.hint{{font-size:12px;color:#666;margin:4px 0 0}}
.tpl-bar{{font-size:12px;margin-top:8px;padding:8px;background:#fff;border:1px solid #ddd}}
label.chk{{display:block;font-size:13px;margin:2px 0}}
label.adopt{{font-weight:bold;color:#063}}
#lightbox{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:99;align-items:center;justify-content:center;cursor:zoom-out}}
#lightbox.on{{display:flex}}
#lightbox img{{max-width:96vw;max-height:96vh;box-shadow:0 0 24px #000}}
</style></head><body>
<div class="toolbar">
  <h1>サブ画像の確認</h1>
  <div id="nextStep" style="margin:10px 0 0;padding:10px 12px;border:2px solid #0653c7;border-radius:6px;background:#eef5ff;font-size:15px;line-height:1.55;color:#063;">
    <b>いまやること</b><br/>
    1. 絵を見る。直したい絵だけ「再生成する」にチェック。直してほしいことがあれば、その下の欄に書く。<br/>
    2. ①「レ点再生成実行」を押す。終わったら「再読込」。同じテーマの新しい絵が並ぶ。<br/>
    3. 残す絵だけ「採用する」にチェック。同じテーマが2枚あるときは、残す1枚だけ。<br/>
    4. ②「採用分を確認フォルダへ」を押し、「確認フォルダを開く」で出すファイルを見る。
  </div>
  <p style="margin:10px 0 0">
    <b>①</b>
    <button onclick="runRegen()">レ点再生成実行</button>
    <b style="margin-left:10px">②</b>
    <button onclick="exportAdopted()" style="font-weight:bold;font-size:15px;">採用分を確認フォルダへ</button>
    <button onclick="openFolder()">確認フォルダを開く</button>
    <button onclick="location.reload()">再読込</button>
    <span id="status"></span>
  </p>
  <details style="margin:8px 0 0;font-size:13px;color:#444;">
    <summary style="cursor:pointer;">文言の定型と、その他のボタン</summary>
    <div class="tpl-bar"><b>定型（最後に書いた要望欄へ入る）:</b><br/>{tpl_list_html}</div>
    <p style="margin:8px 0 0;">
      <button onclick="saveQueue()">修正指示を一時保存</button>
      <button onclick="markDone()">修正なしで終了</button>
    </p>
    <p class="hint">作業フォルダ: {html.escape(str(compose_dir))}</p>
  </details>
</div>
<div class="grid">{cards_html}</div>
<div id="lightbox"><img id="lbImg" alt=""/></div>
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
  card.querySelector('img.thumb').onclick=(e)=>{{
    e.stopPropagation();
    document.getElementById('lbImg').src=e.target.src;
    document.getElementById('lightbox').classList.add('on');
  }};
}});
document.getElementById('lightbox').onclick=()=>document.getElementById('lightbox').classList.remove('on');
document.addEventListener('keydown',(e)=>{{ if(e.key==='Escape') document.getElementById('lightbox').classList.remove('on'); }});
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
async function collectRegen(){{
  return [...document.querySelectorAll('.card')].map(c=>({{
    stem:c.dataset.stem,
    provider:c.dataset.provider,
    checked:!!c.querySelector('.need-regen').checked,
    comment:(c.querySelector('.comment').value||'').trim()
  }})).filter(x=>x.checked);
}}
function collectAdopted(){{
  return [...document.querySelectorAll('.card')].filter(c=>c.querySelector('.adopt-chk').checked)
    .map(c=>c.dataset.stem);
}}
async function saveQueue(){{
  const items=await collectRegen();
  const res=await fetch('/api/save-queue',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{items}})}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
}}
async function runRegen(){{
  await saveQueue();
  document.getElementById('status').textContent='再生成中（旧画像は残します）…';
  const res=await fetch('/api/regen',{{method:'POST'}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
  if(j.ok) setTimeout(()=>location.reload(), 800);
}}
async function exportAdopted(){{
  const stems=collectAdopted();
  if(!stems.length){{
    alert('採用する画像にチェックしてください');
    document.getElementById('status').textContent='採用0件';
    return;
  }}
  document.getElementById('status').textContent='採用分を再export中…';
  const res=await fetch('/api/export-adopted',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{stems}})}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
  if(!j.ok && j.message) alert(j.message);
}}
async function openFolder(){{
  const res=await fetch('/api/open-folder',{{method:'POST'}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
}}
async function markDone(){{
  const items=await collectRegen();
  if(items.length){{
    document.getElementById('status').textContent='まだ再生成レ点が残っています（'+items.length+'）。外すか実行してください';
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
    meta_path = meta_dir_for(root) / "run_meta.json"

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt: str, *args) -> None:  # noqa: A003
            LOG.info("%s - %s", self.address_string(), fmt % args)

        def _send(self, code: int, body: bytes, content_type: str) -> None:
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            self.send_header("Cache-Control", "no-store")
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
                rel = unquote(parsed.path[len("/file/") :]).replace("/", "\\")
                target = (root / rel).resolve()
                root_s = str(root)
                if not (
                    str(target) == root_s
                    or str(target).startswith(root_s + "\\")
                    or str(target).startswith(root_s + "/")
                ):
                    self._send(404, b"not found", "text/plain")
                    return
                if not target.is_file():
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
                msg = f"修正指示を一時保存: {len(items)} 件 → {queue_path.name}"
                self._send(200, json.dumps({"ok": True, "message": msg}, ensure_ascii=False).encode("utf-8"), "application/json")
                return

            if parsed.path == "/api/open-folder":
                folder = Path(DEFAULT_RAKUTEN_UPLOAD_DIR)
                folder.mkdir(parents=True, exist_ok=True)
                try:
                    subprocess.Popen(["explorer", str(folder)])
                    msg = f"目視フォルダを開きました: {folder}"
                    ok = True
                except Exception as e:
                    msg = f"フォルダを開けません: {e}"
                    ok = False
                self._send(
                    200 if ok else 500,
                    json.dumps({"ok": ok, "message": msg}, ensure_ascii=False).encode("utf-8"),
                    "application/json",
                )
                return

            if parsed.path == "/api/export-adopted":
                stems = [str(s).strip() for s in (payload.get("stems") or []) if str(s).strip()]
                if not stems:
                    self._send(
                        400,
                        json.dumps({"ok": False, "message": "残す絵が選ばれていません。「採用する」にチェックしてください。"}, ensure_ascii=False).encode("utf-8"),
                        "application/json",
                    )
                    return
                theme_new = (root / "messages.json").is_file() and not meta_path.is_file()
                if theme_new:
                    try:
                        msg = _export_theme_new_adopted(root, stems)
                        self._send(
                            200,
                            json.dumps({"ok": True, "message": msg}, ensure_ascii=False).encode("utf-8"),
                            "application/json",
                        )
                    except Exception as e:
                        LOG.exception("export theme_new")
                        self._send(
                            500,
                            json.dumps({"ok": False, "message": str(e)}, ensure_ascii=False).encode("utf-8"),
                            "application/json",
                        )
                    return
                if not meta_path.is_file():
                    self._send(
                        400,
                        json.dumps({"ok": False, "message": "run_meta.json がありません"}, ensure_ascii=False).encode("utf-8"),
                        "application/json",
                    )
                    return
                try:
                    original_meta = meta_path.read_text(encoding="utf-8")
                    try:
                        meta = json.loads(original_meta)
                        meta = prefer_adopted_jobs(meta, stems)
                        if not (meta.get("jobs") or []):
                            raise RuntimeError("選んだ絵がジョブに見つかりません")
                        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
                        (meta_dir_for(root) / "adopted_review.json").write_text(
                            json.dumps(
                                {"stems": stems, "at": meta.get("adoptedAt")},
                                ensure_ascii=False,
                                indent=2,
                            ),
                            encoding="utf-8",
                        )
                        from export_sub_images_for_rakuten_matrix import export_from_plan
                        from b5_skip_generate_finish import child_skus_for_image_bind, retire_product_subs

                        plan_p = meta_dir_for(root) / "export_plan.json"
                        if not plan_p.is_file():
                            raise RuntimeError("export_plan.json がありません（初回 compose の計画を確認）")
                        plan = json.loads(plan_p.read_text(encoding="utf-8"))
                        out_dir = Path(str(plan.get("outDir") or DEFAULT_RAKUTEN_UPLOAD_DIR))
                        jan_key = str(plan.get("jan") or meta.get("jan") or "").strip()
                        skus = child_skus_for_image_bind(jan_key, out_dir) if jan_key else []
                        if not skus:
                            raise RuntimeError("子SKUがありません。マスタのこの商品に子SKUが無いため、確認フォルダへ出せません")
                        plan["childSkus"] = skus
                        plan_p.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
                        keys = list(skus)
                        if jan_key:
                            keys.append(jan_key)
                        moved_n = 0
                        for key in keys:
                            moved_n += len(retire_product_subs(out_dir, key))
                        rows = export_from_plan(root)
                        msg = f"採用 {len(rows)} 枚を確認フォルダへ出しました。前のサブ {moved_n} 枚は外しました"
                        self._send(
                            200,
                            json.dumps({"ok": True, "message": msg}, ensure_ascii=False).encode("utf-8"),
                            "application/json",
                        )
                    finally:
                        meta_path.write_text(original_meta, encoding="utf-8")
                except Exception as e:
                    LOG.exception("export-adopted")
                    self._send(
                        500,
                        json.dumps({"ok": False, "message": str(e)}, ensure_ascii=False).encode("utf-8"),
                        "application/json",
                    )
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
                        json.dumps({"ok": True, "message": "レ点0件。再生成不要（OK）。"}, ensure_ascii=False).encode("utf-8"),
                        "application/json",
                    )
                    return
                theme_new = (root / "messages.json").is_file() and not meta_path.is_file()
                if theme_new:
                    cmd = [
                        sys.executable,
                        str(script_dir / "dry_sub_overview_copy_test.py"),
                        "--regen-from",
                        str(root),
                        "--regen-queue",
                        str(queue_path),
                        "--quality",
                        "medium",
                        "--base-color",
                        "beige",
                    ]
                    try:
                        jan = str(
                            json.loads((root / "messages.json").read_text(encoding="utf-8")).get("jan") or ""
                        )
                    except (OSError, json.JSONDecodeError, TypeError):
                        jan = ""
                    if jan:
                        cmd.extend(["--jan", jan])
                else:
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
                    env = dict(os.environ)
                    env["PYTHONIOENCODING"] = "utf-8"
                    proc = subprocess.run(
                        cmd,
                        cwd=str(script_dir),
                        capture_output=True,
                        text=True,
                        timeout=3600,
                        env=env,
                    )
                    ok = proc.returncode == 0
                    if ok and theme_new:
                        shown = root / "03_openai"
                        shown.mkdir(exist_ok=True)
                        for pic in root.glob("T*.jpg"):
                            shutil.copy2(pic, shown / pic.name)
                    msg = (
                        f"再生成完了（旧残・__rN追加） ({n}件)" if ok else f"再生成失敗 code={proc.returncode}"
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
                    json.dumps(
                        {
                            "ok": True,
                            "message": (
                                "この親はレビュー完了。タブを閉じ、次の親JANの review_loop を開いてください。"
                                "（採用分の再exportが未なら「採用分を再export」→目視フォルダ）"
                            ),
                        },
                        ensure_ascii=False,
                    ).encode("utf-8"),
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
