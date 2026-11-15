# -*- coding: utf-8 -*-
"""
テーマ別新規＋競合本線の混在レビュー（段階B）。

- 1画面で両方を見て「採用する」→ 楽天目視フォルダへ連番 export（最大10）
- 再生成: 独自（dry_sub --regen-from）＋競合（compose --regen-from）。旧残し・__rN

例:
  python sub_image_mix_review_loop.py --manifest _out/mix_manifest_4906283045119_overview_copy_….json --port 8780
"""
from __future__ import annotations

import argparse
import json
import logging
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

from rakuten_image_names import sanitize_pattern_slug, sub_filename
from sub_image_mix_manifest import load_mix_manifest
from work_paths import DEFAULT_RAKUTEN_UPLOAD_DIR, meta_dir_for

LOG = logging.getLogger("set_main_image.mix_review")


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def build_html(manifest: Dict[str, Any]) -> str:
    from review_feedback_templates import TEMPLATES, VERSION as TPL_VER

    cards = []
    for c in manifest.get("cards") or []:
        src = str(c.get("source") or "")
        src_label = "独自制作" if src == "theme_new" else "競合本線"
        src_color = "#0a5" if src == "theme_new" else "#05a"
        catch = str(c.get("catch") or "")
        catch_html = f"<p class='catch'>{catch}</p>" if catch else ""
        gen_label = str(c.get("genLabel") or "")
        cards.append(
            f"""
            <div class="card" data-id="{c.get('id')}" data-source="{src}" data-stem="{c.get('stem') or ''}" data-base-stem="{c.get('baseStem') or c.get('stem') or ''}">
              <div class="src" style="color:{src_color};font-weight:bold">{src_label}</div>
              <label class="chk"><input type="checkbox" class="need-regen"/> 再生成する</label>
              <label class="chk adopt"><input type="checkbox" class="adopt-chk"/> 採用する（出口へ）</label>
              <div class="meta">T{int(c.get('themeId') or 0):02d} {c.get('themeName') or ''}{gen_label}
                / {c.get('id')}</div>
              {catch_html}
              <img class="thumb" src="/img/{c.get('id')}" alt="{c.get('id')}" title="クリックで拡大"/>
              <div class="tpl-btns"></div>
              <textarea class="comment" placeholder="要望（テンプレ可・独自／競合とも再生成に反映）"></textarea>
            </div>
            """
        )
    cards_html = "\n".join(cards) or "<p>画像がありません。マニフェストを確認してください。</p>"
    tpl_json = json.dumps(TEMPLATES, ensure_ascii=False)
    tpl_list_html = "".join(
        f'<button type="button" class="tpl-global" data-idx="{i}">{t["label"]}</button> '
        for i, t in enumerate(TEMPLATES)
    )
    jan = manifest.get("jan") or ""
    product_name = str(manifest.get("productName") or "")
    theme_new_dir = str(manifest.get("themeNewDir") or "")
    theme_new_label = Path(theme_new_dir).name if theme_new_dir else ""
    n_tn = sum(1 for c in (manifest.get("cards") or []) if c.get("source") == "theme_new")
    n_cp = sum(1 for c in (manifest.get("cards") or []) if c.get("source") == "competitor")
    fc = manifest.get("foodCanon") or {}
    fc_src = str(fc.get("source") or "")
    fc_score = fc.get("score")
    fc_reason = str(fc.get("reasonJa") or fc.get("reason") or "")
    fc_name = str(fc.get("chosenName") or "")
    fc_ok = bool(fc.get("ok") or (fc.get("path") and fc_src))
    if fc_src in ("competitor_human", "manual", "human"):
        src_ja = "人間選定（IDENTITY／手動）"
    elif fc_src == "competitor" and fc_ok:
        src_ja = "競合画像を見本に選定"
    elif fc_src in ("package_truth", "package_truth_copy", "package_truth_human"):
        src_ja = "正本パッケージを見本に"
    elif fc_src == "package_only_fallback":
        src_ja = "個品なし（正本のみ）"
    elif fc_ok and (fc.get("path") or fc_name):
        src_ja = "見本あり"
    else:
        src_ja = "未選定または失敗"
    cand_html = []
    for cand in fc.get("candidates") or []:
        picked = bool(cand.get("picked"))
        border = "3px solid #c00" if picked else "1px solid #ccc"
        label = "★見本" if picked else f"候補{cand.get('index')}"
        name = str(cand.get("name") or "")
        cand_html.append(
            f"""<div class="cand" style="border:{border}">
              <div class="cand-lab">{label}</div>
              <img src="/file/01_refs/{name}" alt="{name}" title="{name}" class="cand-img"/>
              <div class="cand-name">{name}</div>
            </div>"""
        )
    cand_block = "".join(cand_html) or "<p class='hint'>候補画像なし</p>"
    score_bit = f"／ score={fc_score}" if fc_score is not None else ""
    name_bit = f"／ {fc_name}" if fc_name else ""
    canon_block = f"""
    <div class="canon-panel">
      <h2>食材見本 FOOD_UNIT_CANON</h2>
      <p class="hint">{src_ja}{score_bit}{name_bit}</p>
      <p class="hint">{fc_reason}</p>
      <div class="canon-row">
        <div class="canon-main">
          <div class="cand-lab">採用見本</div>
          <img src="/canon" alt="FOOD_UNIT_CANON" class="canon-img" title="クリックで拡大"/>
        </div>
        <div class="cand-strip">{cand_block}</div>
      </div>
    </div>
    """
    return f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>混在レビュー — {jan}</title>
<style>
body{{font-family:sans-serif;margin:16px;background:#f6f3ee;color:#222}}
h1{{font-size:18px;margin:0 0 8px}}
h2{{font-size:15px;margin:0 0 6px}}
.toolbar{{position:sticky;top:0;background:#f6f3ee;padding:8px 0 12px;z-index:2;border-bottom:1px solid #ccc}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;margin-top:12px}}
.card{{background:#fff;border:1px solid #ddd;padding:8px;border-radius:6px}}
.card img.thumb{{width:100%;height:auto;display:block;border:1px solid #eee;cursor:zoom-in}}
.meta,.catch{{font-size:12px;color:#555;margin:4px 0 6px}}
.catch{{color:#333}}
.src{{font-size:12px;margin-bottom:4px}}
textarea{{width:100%;min-height:56px;font-size:12px;box-sizing:border-box}}
button{{margin-right:8px;padding:8px 12px;cursor:pointer}}
.tpl-global,.tpl-local{{font-size:11px;padding:4px 8px;margin:2px 4px 2px 0}}
#status{{margin-left:8px;font-size:13px;color:#333}}
.hint{{font-size:12px;color:#666;margin:4px 0 0}}
.tpl-bar{{font-size:12px;margin-top:8px;padding:8px;background:#fff;border:1px solid #ddd}}
label.chk{{display:block;font-size:13px;margin:2px 0}}
label.adopt{{font-weight:bold;color:#063}}
.canon-panel{{margin-top:12px;padding:12px;background:#fff;border:2px solid #c45;border-radius:8px}}
.canon-row{{display:flex;flex-wrap:wrap;gap:16px;align-items:flex-start;margin-top:8px}}
.canon-main{{flex:0 0 220px}}
.canon-img{{width:220px;height:auto;border:2px solid #c45;cursor:zoom-in;display:block}}
.cand-strip{{display:flex;flex-wrap:wrap;gap:10px;flex:1}}
.cand{{width:140px;background:#fafafa;padding:6px;border-radius:4px}}
.cand-img{{width:100%;min-height:96px;height:auto;display:block;cursor:zoom-in;background:#f0ece6}}
.cand-lab{{font-size:11px;font-weight:bold;color:#c45;margin-bottom:4px}}
.cand-name{{font-size:10px;color:#666;word-break:break-all;margin-top:4px}}
#lightbox{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:99;align-items:center;justify-content:center;cursor:zoom-out}}
#lightbox.on{{display:flex}}
#lightbox img{{max-width:96vw;max-height:96vh;box-shadow:0 0 24px #000}}
#regenBanner{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:200;align-items:center;justify-content:center}}
#regenBanner.on{{display:flex}}
#regenBanner .box{{background:#b00020;color:#fff;font-size:32px;font-weight:bold;padding:28px 40px;border-radius:8px;text-align:center;line-height:1.45;max-width:90vw}}
#regenBanner.done .box{{background:#1565c0}}
#regenBanner .box span{{display:block;margin-top:10px;font-size:16px;font-weight:normal}}
</style></head><body>
<div id="regenBanner"><div class="box">再作成しています<span>終わるまでこの画面を閉じないでください</span></div></div>
<div class="toolbar">
  <h1>混在レビュー（独自制作＋競合本線）— JAN {jan}</h1>
  <div class="hint"><b style="font-size:14px;color:#222;">{product_name or '（商品名なし）'}</b><br/>
  作業フォルダ: {theme_new_label or '（theme-newなし）'}<br/>
  独自 {n_tn} 枚 ／ 競合 {n_cp} 枚 ／ 出口は採用分のみ（最大10）。同テーマはどちらか1枚を採用推奨。<br/>
  テンプレ v{TPL_VER} ／ 段階B（独自・競合とも再生成・__rN）。食材見た目は FOOD_UNIT_CANON で横断統一。<br/>
  画像クリック＝拡大。同じJANの別商品タブと取り違えないよう、上の商品名と作業フォルダを確認。</div>
  <div id="nextStep" style="margin:10px 0 0;padding:10px 12px;border:2px solid #0653c7;border-radius:6px;background:#eef5ff;font-size:15px;line-height:1.45;font-weight:bold;color:#063;">いまやること: 準備中…</div>
  <div class="tpl-bar"><b>テンプレ挿入:</b><br/>{tpl_list_html}</div>
  <p style="margin:10px 0 0">
    <b>①</b>
    <button onclick="saveQueue()">修正指示を一時保存</button>
    <button onclick="runRegen()">再生成（独自＋競合）</button>
    <b style="margin-left:8px;">②</b>
    <button onclick="exportAdopted()" style="font-weight:bold;font-size:15px;">採用分を出口へ</button>
    <button onclick="openFolder()">目視フォルダを開く</button>
    <button onclick="markDone()">終了</button>
    <button onclick="location.reload()">再読込</button>
    <span id="status"></span>
  </p>
</div>
{canon_block}
<div class="grid">{cards_html}</div>
<div id="lightbox"><img id="lbImg" alt=""/></div>
<script>
const TPL = {tpl_json};
let lastTa = null;
function openLb(src){{
  document.getElementById('lbImg').src=src;
  document.getElementById('lightbox').classList.add('on');
}}
function refreshNextStep(){{
  const el=document.getElementById('nextStep');
  if(!el) return;
  const nAdopt=[...document.querySelectorAll('.adopt-chk')].filter(x=>x.checked).length;
  const nRegen=[...document.querySelectorAll('.need-regen')].filter(x=>x.checked).length;
  if(nRegen>0 && nAdopt===0){{
    el.style.borderColor='#0653c7';
    el.style.background='#eef5ff';
    el.style.color='#063';
    el.innerHTML='いまやること:<br/>'
      +'① 要望を書いた絵の「再生成する」が付いている（'+nRegen+'枚）→「再生成（独自＋競合）」を押す<br/>'
      +'② できた絵を見て、残すものに「採用する」';
    return;
  }}
  if(nAdopt===0){{
    el.style.borderColor='#0653c7';
    el.style.background='#eef5ff';
    el.style.color='#063';
    el.innerHTML='いまやること:<br/>'
      +'① おかしい絵 → そのカードに要望を書く →「再生成する」にレ点 →「再生成（独自＋競合）」<br/>'
      +'② 残す絵に「採用する」を付ける（出口は採用分だけ・最大10。同テーマは独自か競合のどちらか1枚）';
    return;
  }}
  el.style.borderColor='#0a7';
  el.style.background='#f3fff6';
  el.style.color='#063';
  el.innerHTML='いまやること:<br/>'
    +'③ 「採用分を出口へ」を押す（いま '+nAdopt+' 枚）→ 親1セット（子SKU名）で楽天フォルダへ<br/>'
    +'④ 「目視フォルダを開く」で確認。問題なければ「終了」。Cは親サブ→全子同一';
}}
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
      const rg=card.querySelector('.need-regen');
      if(rg && !rg.disabled) rg.checked=true;
      lastTa=ta;
      refreshNextStep();
    }};
    bar.appendChild(b);
  }});
  card.querySelector('img.thumb').onclick=(e)=>{{
    e.stopPropagation();
    openLb(e.target.src);
  }};
  card.querySelectorAll('.need-regen,.adopt-chk').forEach(ck=>{{
    ck.addEventListener('change', refreshNextStep);
  }});
}});
document.querySelectorAll('.canon-img,.cand-img').forEach(img=>{{
  img.onclick=(e)=>{{ e.stopPropagation(); openLb(e.target.src); }};
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
    if(card){{ const rg=card.querySelector('.need-regen'); if(rg && !rg.disabled) rg.checked=true; }}
    ta.focus();
    refreshNextStep();
  }};
}});
refreshNextStep();
async function collectRegen(){{
  return [...document.querySelectorAll('.card')].map(c=>({{
    id:c.dataset.id,
    source:c.dataset.source,
    stem:c.dataset.stem,
    baseStem:c.dataset.baseStem||c.dataset.stem,
    checked:!!(c.querySelector('.need-regen')&&c.querySelector('.need-regen').checked),
    comment:(c.querySelector('.comment').value||'').trim()
  }})).filter(x=>x.checked);
}}
function collectAdopted(){{
  return [...document.querySelectorAll('.card')].filter(c=>c.querySelector('.adopt-chk').checked)
    .map(c=>c.dataset.id);
}}
async function saveQueue(){{
  const items=await collectRegen();
  const res=await fetch('/api/save-queue',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{items}})}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
}}
async function runRegen(){{
  const items=await collectRegen();
  if(!items.length){{
    alert('再生成する画像にチェックしてください');
    return;
  }}
  await saveQueue();
  const banner=document.getElementById('regenBanner');
  if(banner){{
    banner.classList.remove('done');
    banner.classList.add('on');
    const box=banner.querySelector('.box');
    if(box) box.innerHTML='再作成しています<span>終わるまでこの画面を閉じないでください（'+items.length+'枚）</span>';
  }}
  document.getElementById('status').textContent='再作成しています';
  try{{
    const res=await fetch('/api/regen',{{method:'POST'}});
    const j=await res.json();
    document.getElementById('status').textContent=j.message||JSON.stringify(j);
    if(j.ok){{
      if(banner){{
        banner.classList.add('done');
        const box=banner.querySelector('.box');
        if(box) box.innerHTML='再作成が終わりました<span>画面を読み直しています</span>';
      }}
      setTimeout(()=>location.reload(), 1200);
    }} else if(banner) banner.classList.remove('on');
    if(!j.ok && j.message) alert(j.message);
  }}catch(e){{
    if(banner) banner.classList.remove('on');
    document.getElementById('status').textContent='再作成に失敗しました';
    alert('再作成に失敗しました');
  }}
}}
async function exportAdopted(){{
  const ids=collectAdopted();
  if(!ids.length){{
    alert('採用する画像にチェックしてください');
    return;
  }}
  document.getElementById('status').textContent='採用分を出口へ…';
  const res=await fetch('/api/export-adopted',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{ids}})}});
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
  const res=await fetch('/api/done',{{method:'POST'}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||'完了';
}}
</script>
</body></html>
"""


def _parent_sku_from_mix_manifest(manifest: Dict[str, Any]) -> str:
    explicit = str(manifest.get("parentSku") or "").strip()
    if explicit:
        return explicit
    from b5_skip_generate_finish import parent_sku_from_path_hint

    fc = manifest.get("foodCanon") or {}
    for key in ("truth", "path"):
        p = parent_sku_from_path_hint(str(fc.get(key) or ""))
        if p:
            return p
    return parent_sku_from_path_hint(str(fc.get("chosenName") or ""))


def export_adopted_cards(
    manifest: Dict[str, Any],
    card_ids: List[str],
    *,
    out_dir: Path,
    max_subs: int = 10,
) -> List[Dict[str, Any]]:
    """
    採用分を02直下へ。Cレイアウト用に親あたり子SKU名1セットだけ出す。
    JAN名出口・全子複製はしない（枠外候補の元になる）。
    """
    from b5_skip_generate_finish import (
        child_skus_for_image_bind,
        represent_child_skus_for_layout_export,
        retire_product_subs,
    )

    by_id = {str(c.get("id")): c for c in (manifest.get("cards") or [])}
    jan = str(manifest.get("jan") or "").strip()
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    parent_sku = _parent_sku_from_mix_manifest(manifest)
    skus = represent_child_skus_for_layout_export(
        jan, parent_sku=parent_sku, out_dir=out_dir
    )
    if not skus:
        raise RuntimeError(
            "出口用の子SKUが解決できません。マスタの出品CKまたは "
            f"_rakuten.jpg（親={parent_sku or '不明'} JAN={jan or '不明'}）を確認してください"
        )
    key = skus[0]
    # 旧JAN名・当該親の子サブ余りを外してから1セットだけ置く（別親は触らない）
    allow = {key}
    if parent_sku:
        from master_sets import load_checked_set_children_from_rows
        from sheets_master import fetch_master_rows

        rows, _ = fetch_master_rows()
        kids, _ = load_checked_set_children_from_rows(rows, parent_sku=parent_sku)
        allow = {k.child_sku for k in kids} or {key}
    elif jan:
        allow = set(child_skus_for_image_bind(jan, out_dir)) | {key}
    retire_product_subs(out_dir, jan)
    for old in sorted(allow):
        retire_product_subs(out_dir, old)

    cap = max(1, min(int(max_subs), 10))
    report: List[Dict[str, Any]] = []
    n = 0
    for cid in card_ids:
        if n >= cap:
            break
        c = by_id.get(str(cid))
        if not c:
            continue
        src = Path(str(c.get("path") or ""))
        if not src.is_file():
            LOG.warning("missing src %s", src)
            continue
        n += 1
        slug = sanitize_pattern_slug(c.get("themeSlug") or c.get("themeName") or "")
        dest_name = sub_filename(key, n, pattern=slug)
        dest = out_dir / dest_name
        if dest.exists():
            stem = dest.stem
            dest = out_dir / f"{stem}__mix{datetime.now().strftime('%H%M%S')}{dest.suffix}"
        shutil.copy2(src, dest)
        report.append(
            {
                "id": cid,
                "source": c.get("source"),
                "themeId": c.get("themeId"),
                "src": str(src),
                "dest": str(dest),
                "subIndex": n,
                "exportSku": key,
                "parentSku": parent_sku,
            }
        )
        LOG.info("export %s -> %s parent=%s", cid, dest.name, parent_sku)
    return report


def make_handler(manifest_path: Path, script_dir: Path):
    manifest_path = Path(manifest_path).resolve()
    work_dir = manifest_path.parent
    queue_path = work_dir / f"{manifest_path.stem}_regen_queue.json"
    done_path = work_dir / f"{manifest_path.stem}_done.json"

    def _load() -> Dict[str, Any]:
        return load_mix_manifest(manifest_path)

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
            man = _load()
            if parsed.path in ("/", "/index.html"):
                html = build_html(man)
                self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
                return
            if parsed.path.startswith("/img/"):
                cid = unquote(parsed.path[len("/img/") :])
                by_id = {str(c.get("id")): c for c in (man.get("cards") or [])}
                card = by_id.get(cid)
                if not card:
                    self._send(404, b"not found", "text/plain")
                    return
                target = Path(str(card.get("path") or ""))
                if not target.is_file():
                    self._send(404, b"not found", "text/plain")
                    return
                data = target.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(data)
                return
            if parsed.path == "/canon":
                fc = man.get("foodCanon") or {}
                target = Path(str(fc.get("path") or ""))
                if not target.is_file():
                    # themeNewDir から再解決
                    tn = Path(str(man.get("themeNewDir") or ""))
                    if tn.is_dir():
                        from food_unit_canon import existing_canon_path

                        ex = existing_canon_path(tn)
                        if ex:
                            target = ex
                if not target.is_file():
                    self._send(404, b"canon not found", "text/plain")
                    return
                data = target.read_bytes()
                ctype = "image/png" if target.suffix.lower() == ".png" else "image/jpeg"
                self._send(200, data, ctype)
                return
            if parsed.path.startswith("/file/"):
                rel = unquote(parsed.path[len("/file/") :]).replace("\\", "/").lstrip("/")
                tn = Path(str(man.get("themeNewDir") or ""))
                if not tn.is_dir():
                    self._send(404, b"not found", "text/plain")
                    return
                target = (tn / rel).resolve()
                root_s = str(tn.resolve())
                if not (
                    str(target) == root_s
                    or str(target).startswith(root_s + "\\")
                    or str(target).startswith(root_s + "/")
                ):
                    self._send(404, b"not found", "text/plain")
                    return
                if not target.is_file():
                    self._send(404, b"ref not found", "text/plain")
                    return
                data = target.read_bytes()
                ctype = "image/png" if target.suffix.lower() == ".png" else "image/jpeg"
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
            man = _load()

            if parsed.path == "/api/save-queue":
                items = payload.get("items") or []
                queue = {
                    "manifest": str(manifest_path),
                    "savedAt": datetime.now(timezone.utc).isoformat(),
                    "items": items,
                }
                queue_path.write_text(
                    json.dumps(queue, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                msg = f"修正指示を一時保存: {len(items)} 件"
                self._send(
                    200,
                    json.dumps({"ok": True, "message": msg}, ensure_ascii=False).encode("utf-8"),
                    "application/json",
                )
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
                ids = [str(x).strip() for x in (payload.get("ids") or []) if str(x).strip()]
                if not ids:
                    self._send(
                        400,
                        json.dumps(
                            {"ok": False, "message": "採用 id が空です"}, ensure_ascii=False
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return
                try:
                    rows = export_adopted_cards(
                        man, ids, out_dir=Path(DEFAULT_RAKUTEN_UPLOAD_DIR), max_subs=10
                    )
                    (work_dir / f"{manifest_path.stem}_adopted.json").write_text(
                        json.dumps(
                            {"ids": ids, "exports": rows, "at": datetime.now(timezone.utc).isoformat()},
                            ensure_ascii=False,
                            indent=2,
                        ),
                        encoding="utf-8",
                    )
                    sku = (rows[0].get("exportSku") if rows else "") or ""
                    msg = (
                        f"採用 {len(rows)} 枚を出口へ（親1セット・{sku or '子SKU'}）: "
                        f"{DEFAULT_RAKUTEN_UPLOAD_DIR}"
                    )
                    self._send(
                        200,
                        json.dumps({"ok": True, "message": msg, "n": len(rows)}, ensure_ascii=False).encode(
                            "utf-8"
                        ),
                        "application/json",
                    )
                except Exception as e:
                    LOG.exception("export-adopted")
                    self._send(
                        500,
                        json.dumps({"ok": False, "message": str(e)}, ensure_ascii=False).encode(
                            "utf-8"
                        ),
                        "application/json",
                    )
                return

            if parsed.path == "/api/regen":
                if not queue_path.is_file():
                    self._send(
                        400,
                        json.dumps(
                            {"ok": False, "message": "キューがありません。先に保存してください。"},
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return
                q = json.loads(queue_path.read_text(encoding="utf-8"))
                items = [x for x in (q.get("items") or []) if x.get("checked")]
                comp_items = [x for x in items if x.get("source") == "competitor"]
                tn_items = [x for x in items if x.get("source") == "theme_new"]
                if not items:
                    self._send(
                        200,
                        json.dumps(
                            {"ok": True, "message": "再生成レ点が0件。"},
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return
                msgs: List[str] = []
                ok_all = True

                if tn_items:
                    theme_new_dir = str(man.get("themeNewDir") or "")
                    if not theme_new_dir or not Path(theme_new_dir).is_dir():
                        ok_all = False
                        msgs.append("独自: themeNewDir がありません")
                    else:
                        tn_q = {
                            "themeNewDir": theme_new_dir,
                            "savedAt": datetime.now(timezone.utc).isoformat(),
                            "items": [
                                {
                                    "stem": x.get("baseStem") or x.get("stem"),
                                    "themeId": None,
                                    "checked": True,
                                    "comment": x.get("comment") or "",
                                }
                                for x in tn_items
                            ],
                        }
                        tn_q_path = Path(theme_new_dir) / "regen_queue.json"
                        tn_q_path.write_text(
                            json.dumps(tn_q, ensure_ascii=False, indent=2), encoding="utf-8"
                        )
                        cmd = [
                            sys.executable,
                            str(script_dir / "dry_sub_overview_copy_test.py"),
                            "--regen-from",
                            theme_new_dir,
                            "--regen-queue",
                            str(tn_q_path),
                        ]
                        LOG.info("spawn theme_new regen: %s", " ".join(cmd))
                        try:
                            proc = subprocess.run(
                                cmd,
                                cwd=str(script_dir),
                                capture_output=True,
                                text=True,
                                timeout=3600,
                            )
                            if proc.returncode == 0:
                                msgs.append(f"独自再生成完了 {len(tn_items)} 件")
                            else:
                                ok_all = False
                                err = (proc.stderr or proc.stdout or "").strip()[-300:]
                                msgs.append(f"独自再生成失敗 code={proc.returncode} {err}")
                        except Exception as e:
                            ok_all = False
                            LOG.exception("theme_new regen")
                            msgs.append(f"独自再生成例外: {e}")

                if comp_items:
                    compose_dir = str(man.get("composeDir") or "")
                    if not compose_dir or not Path(compose_dir).is_dir():
                        ok_all = False
                        msgs.append("競合: composeDir がありません")
                    else:
                        regen_q = {
                            "composeDir": compose_dir,
                            "provider": man.get("provider") or "openai",
                            "savedAt": datetime.now(timezone.utc).isoformat(),
                            "items": [
                                {
                                    "stem": x.get("baseStem") or x.get("stem"),
                                    "provider": man.get("provider") or "openai",
                                    "checked": True,
                                    "comment": x.get("comment") or "",
                                }
                                for x in comp_items
                            ],
                        }
                        rq_path = meta_dir_for(Path(compose_dir)) / "regen_queue.json"
                        rq_path.parent.mkdir(parents=True, exist_ok=True)
                        rq_path.write_text(
                            json.dumps(regen_q, ensure_ascii=False, indent=2), encoding="utf-8"
                        )
                        cmd = [
                            sys.executable,
                            str(script_dir / "sub_image_ai_compose_poc.py"),
                            "--regen-from",
                            compose_dir,
                            "--regen-queue",
                            str(rq_path),
                            "--providers",
                            str(man.get("provider") or "openai"),
                        ]
                        LOG.info("spawn competitor regen: %s", " ".join(cmd))
                        try:
                            proc = subprocess.run(
                                cmd,
                                cwd=str(script_dir),
                                capture_output=True,
                                text=True,
                                timeout=3600,
                            )
                            if proc.returncode == 0:
                                msgs.append(f"競合再生成完了 {len(comp_items)} 件")
                            else:
                                ok_all = False
                                err = (proc.stderr or proc.stdout or "").strip()[-300:]
                                msgs.append(f"競合再生成失敗 code={proc.returncode} {err}")
                        except Exception as e:
                            ok_all = False
                            LOG.exception("competitor regen")
                            msgs.append(f"競合再生成例外: {e}")

                if ok_all and (man.get("themeNewDir") or man.get("composeDir")):
                    from sub_image_mix_manifest import build_mix_manifest, write_mix_manifest

                    man2 = build_mix_manifest(
                        jan=str(man.get("jan") or ""),
                        theme_new_dir=Path(man["themeNewDir"]) if man.get("themeNewDir") else None,
                        compose_dir=Path(man["composeDir"]) if man.get("composeDir") else None,
                        provider=str(man.get("provider") or "openai"),
                        product_name=str(man.get("productName") or ""),
                    )
                    write_mix_manifest(man2, manifest_path)

                msg = " ／ ".join(msgs) if msgs else "再生成なし"
                self._send(
                    200 if ok_all else 500,
                    json.dumps({"ok": ok_all, "message": msg}, ensure_ascii=False).encode(
                        "utf-8"
                    ),
                    "application/json",
                )
                return

            if parsed.path == "/api/done":
                done_path.write_text(
                    json.dumps(
                        {"at": datetime.now(timezone.utc).isoformat(), "jan": man.get("jan")},
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
                self._send(
                    200,
                    json.dumps(
                        {"ok": True, "message": "この親の混在レビュー完了。タブを閉じて次へ。"},
                        ensure_ascii=False,
                    ).encode("utf-8"),
                    "application/json",
                )
                return

            self._send(404, b"not found", "text/plain")

    return Handler


def serve(manifest_path: Path, port: int, open_browser: bool) -> int:
    script_dir = Path(__file__).resolve().parent
    handler = make_handler(manifest_path, script_dir)
    httpd = ThreadingHTTPServer(("127.0.0.1", port), handler)
    url = f"http://127.0.0.1:{port}/"
    LOG.info("mix review UI: %s", url)
    print(url)
    if open_browser:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        LOG.info("stopped")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="独自制作＋競合 混在レビュー（段階B）")
    ap.add_argument("--manifest", type=Path, required=True)
    ap.add_argument("--port", type=int, default=8780)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)
    man_path = Path(args.manifest)
    if not man_path.is_file():
        raise SystemExit(f"manifest がありません: {man_path}")
    return serve(man_path, int(args.port), open_browser=not args.no_browser)


if __name__ == "__main__":
    raise SystemExit(main())
