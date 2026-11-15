# -*- coding: utf-8 -*-
"""
食材見本ゲート → 生成 → 同画面レビュー（B45 §7.12）。

1. 機械が B-③ 候補を提示＋常時任意コピペ（最大 MANUAL_MAX＝6）
2. 競合ゼロ時は必須案内。個品なしは正本のみフォールバック可
3. 人間が主1＋副0〜2（または正本のみ）→ 実行 → 同画面レビュー

例:
  python sub_image_food_canon_gate.py --jan 4906283045119 --port 8790
"""
from __future__ import annotations

import argparse
import html as html_lib
import json
import logging
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

from dry_sub_overview_copy_test import _download_comp_refs, _find_truth
from dry_sub_overview_openai import _find_truth as _find_truth_alt
from food_unit_canon import (
    MAX_SECONDARY,
    FOOD_CANONS_NAME,
    apply_human_selection,
    existing_canon_path,
    existing_sizzle_path,
    existing_wrap_path,
    load_candidates_meta,
    load_canon_meta,
    load_food_canons,
    load_part_bundle,
    resolve_package_bag_for_name,
    resolve_comps_for_dir,
    secondary_paths,
    write_candidates_shortlist,
    CANDIDATES_NAME,
)
from food_canon_manual_ingest import (
    MANUAL_MAX,
    count_b3_style_comps,
    delete_manual_slot,
    list_manual_paths,
    load_manual_meta,
    merge_manual_into_candidates,
    save_manual_data_url,
    save_manual_url,
)
from sub_image_mix_manifest import (
    build_mix_manifest,
    unique_mix_manifest_path,
    write_mix_manifest,
)
from sub_image_mix_review_loop import build_html as build_mix_html
from sub_image_review_launch import compose_root, latest_compose_dir
from package_truth import list_package_truth_candidates, rel_under_amazon_base
from work_paths import DEFAULT_WORK_ROOT

LOG = logging.getLogger("set_main_image.food_canon_gate")
SCRIPT_DIR = Path(__file__).resolve().parent
SESSION_NAME = "gate_session.json"
BAG_CAND_LIMIT = 16


def _current_bag_path(session: Dict[str, Any]) -> Optional[Path]:
    for key in ("bag", "packageTruth", "truth"):
        p = Path(str(session.get(key) or "").strip())
        if p.is_file():
            return p
    try:
        hit = resolve_package_bag_for_name(
            name_base=str(session.get("nameBase") or session.get("productName") or "").strip(),
            parent_sku=str(session.get("parentSku") or "").strip(),
            jan_for_folder=str(session.get("jan") or ""),
            bag_path=str(session.get("bag") or session.get("packageTruth") or ""),
        )
        if hit is not None and hit.is_file():
            return hit
    except Exception as e:
        LOG.warning("current bag resolve: %s", e)
    return None


def _bag_candidates_for_session(session: Dict[str, Any]) -> List[Dict[str, Any]]:
    cur = _current_bag_path(session)
    cands = list_package_truth_candidates(
        jan=str(session.get("jan") or ""),
        parent_sku=str(session.get("parentSku") or ""),
        name_base=str(session.get("nameBase") or session.get("productName") or ""),
        current=cur,
        limit=BAG_CAND_LIMIT,
    )
    session["bagCandidates"] = [
        {"rel": str(c.get("rel") or ""), "path": str(c.get("path") or "")} for c in cands
    ]
    return cands


def _bag_pick_section_html(session: Dict[str, Any]) -> str:
    cands = _bag_candidates_for_session(session)
    cards: List[str] = []
    for i, c in enumerate(cands):
        rel = str(c.get("rel") or "")
        selected = bool(c.get("selected"))
        checked = "checked" if selected else ""
        cls = "bag-card selected" if selected else "bag-card"
        cards.append(
            f"""
            <label class="{cls}">
              <img src="/bag-cand/{i}" alt="{html_lib.escape(rel)}" class="bag-img"/>
              <input type="radio" name="bagRel" value="{html_lib.escape(rel)}" {checked}/>
              <span class="cand-meta"><code>{html_lib.escape(rel)}</code></span>
            </label>
            """
        )
    grid = "\n".join(cards) or "<p class='warn'>01.amazon白抜きベース に候補がありません。</p>"
    return f"""
<section class="bag-pick">
  <h2 style="font-size:15px;margin:0 0 6px">袋の正本（画像で選ぶ）</h2>
  <p class="hint">候補はスプシ①と同じ <code>01.amazon白抜きベース</code>（処理済み含む）です。
  <b>他の親に紐付け済みの袋は出しません</b>（いまこの親で選んでいる袋は残します）。
  画像を選んで「袋をこの画像にする」→ 親共有へ保存します。</p>
  <div class="grid bag-grid">{grid}</div>
  <p style="margin-top:8px">
    <button type="button" id="btnSetBag" style="font-weight:bold">袋をこの画像にする</button>
    <span id="bagStatus"></span>
  </p>
</section>
"""


def _bag_pick_js() -> str:
    return """
document.querySelectorAll('.bag-img').forEach(img=>{
  img.onclick=(e)=>{
    e.preventDefault();
    const lb=document.getElementById('lbImg');
    const box=document.getElementById('lightbox');
    if(lb&&box){ lb.src=img.src; box.classList.add('on'); }
  };
});
const btnBag=document.getElementById('btnSetBag');
if(btnBag) btnBag.onclick=async()=>{
  const rel=(document.querySelector('input[name=bagRel]:checked')||{}).value||'';
  const st=document.getElementById('bagStatus');
  if(!rel){ alert('袋画像を1つ選んでください'); return; }
  if(st) st.textContent='袋を保存中…';
  const res=await fetch('/api/set-bag',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rel})});
  const j=await res.json();
  if(st) st.textContent=j.message||'';
  if(j.ok) location.reload();
};
"""


def apply_bag_pick(
    theme_new_dir: Path,
    session: Dict[str, Any],
    *,
    rel: str = "",
) -> Dict[str, Any]:
    """袋正本を画像選択で確定。01_refs／session／_b5_parent_shared を同じ相対パスで更新。"""
    want = str(rel or "").strip().replace("/", "\\")
    if not want:
        raise ValueError("袋の相対パスが空です")
    cands = list(session.get("bagCandidates") or []) or _bag_candidates_for_session(session)
    path: Optional[Path] = None
    for c in cands:
        if str(c.get("rel") or "").replace("/", "\\") == want:
            path = Path(str(c.get("path") or ""))
            break
    if path is None or not path.is_file():
        # セッション外の相対も amazon 配下だけ許可
        from food_unit_canon import _resolve_bag_rel

        path = _resolve_bag_rel(want)
    if path is None or not path.is_file():
        raise ValueError(f"袋画像が見つかりません: {want}")
    resolved_rel = rel_under_amazon_base(path)
    refs = Path(theme_new_dir) / "01_refs"
    refs.mkdir(parents=True, exist_ok=True)
    for old in refs.glob("PACKAGE_TRUTH.*"):
        try:
            old.unlink()
        except Exception:
            pass
    pkg = refs / f"PACKAGE_TRUTH{path.suffix.lower() or '.png'}"
    try:
        from PIL import Image

        Image.open(path).convert("RGB").save(pkg.with_suffix(".jpg"), quality=95)
        pkg = pkg.with_suffix(".jpg")
    except Exception:
        shutil.copy2(path, pkg)
    session["truth"] = str(path)
    session["packageTruth"] = str(path)
    session["bag"] = str(path)
    session["packageTruthRel"] = resolved_rel
    parent_sku = str(session.get("parentSku") or "").strip()
    if parent_sku:
        try:
            from b5_parent_shared import upsert_parent_shared

            upsert_parent_shared(
                parent_sku=parent_sku,
                jan=str(session.get("jan") or ""),
                name_base=str(session.get("nameBase") or session.get("productName") or ""),
                package_truth_rel=resolved_rel,
            )
        except Exception as e:
            LOG.warning("parent_shared bag upsert: %s", e)
    try:
        meta = load_food_canons(theme_new_dir)
        if isinstance(meta, dict) and meta:
            meta["truth"] = str(path)
            meta["packageTruthRel"] = resolved_rel
            meta["packageTruthUpdatedAt"] = datetime.now(timezone.utc).isoformat()
            (Path(theme_new_dir) / FOOD_CANONS_NAME).write_text(
                json.dumps(meta, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
    except Exception as e:
        LOG.warning("food_canons truth update: %s", e)
    _bag_candidates_for_session(session)
    _save_session(theme_new_dir, session)
    LOG.info("bag pick saved rel=%s path=%s parent=%s", resolved_rel, path, parent_sku)
    return {
        "ok": True,
        "rel": resolved_rel,
        "path": str(path),
        "message": f"袋の正本を保存しました: {resolved_rel}",
    }


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _truth_for(jan: str, theme_new_dir: Path) -> Path:
    """正本。summary の誤結線（他SKU）を信じず、まず JAN ヒント解決を優先。"""
    # 1) messages の商品名からヒント
    hint = ""
    msg = theme_new_dir / "messages.json"
    if msg.is_file():
        try:
            m = json.loads(msg.read_text(encoding="utf-8-sig"))
            pname = str(m.get("productName") or "")
            for key in ("かつお", "まぐろ", "ぶり"):
                if key in pname:
                    hint = key
                    break
        except Exception:
            pass
    # 2) JAN→ヒント表付き _find_truth
    try:
        found = Path(_find_truth(jan, hint))
        if found.is_file():
            return found
    except Exception:
        pass
    try:
        found = Path(_find_truth_alt(jan, hint))
        if found.is_file():
            return found
    except Exception:
        pass
    # 3) summary は最終手段（かつヒントと矛盾したら捨てる）
    summary = theme_new_dir / "summary.json"
    if summary.is_file():
        try:
            s = json.loads(summary.read_text(encoding="utf-8-sig"))
            t = Path(str(s.get("truth") or ""))
            if t.is_file():
                if hint and hint not in t.name:
                    LOG.warning(
                        "summary truth を無視（ヒント不一致） jan=%s hint=%s truth=%s",
                        jan,
                        hint,
                        t.name,
                    )
                else:
                    return t
        except Exception:
            pass
    raise SystemExit(f"正本がありません JAN={jan} hint={hint}")


def _ensure_messages(jan: str, theme_new_dir: Path) -> Path:
    msg = theme_new_dir / "messages.json"
    if msg.is_file():
        return msg
    src = SCRIPT_DIR / "_out" / f"b5_theme_new_messages_{jan}.json"
    if not src.is_file():
        raise SystemExit(f"messages がありません: {msg} / {src}")
    theme_new_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, msg)
    return msg


def prepare_gate(
    *,
    jan: str,
    theme_new_dir: Path,
    work_root: Path,
    refresh_candidates: bool = False,
    max_comps: int = 10,
) -> Dict[str, Any]:
    theme_new_dir = Path(theme_new_dir)
    theme_new_dir.mkdir(parents=True, exist_ok=True)
    refs = theme_new_dir / "01_refs"
    refs.mkdir(parents=True, exist_ok=True)
    truth = _truth_for(jan, theme_new_dir)
    if not truth.is_file():
        raise SystemExit(f"正本がありません JAN={jan}")
    # 正本コピー
    pkg = refs / f"PACKAGE_TRUTH{truth.suffix.lower() or '.png'}"
    if not pkg.is_file():
        try:
            from PIL import Image

            Image.open(truth).convert("RGB").save(pkg.with_suffix(".jpg"), quality=95)
            pkg = pkg.with_suffix(".jpg")
        except Exception:
            shutil.copy2(truth, pkg)
    comps = resolve_comps_for_dir(theme_new_dir)
    # 手動以外を優先DL
    b3_before = count_b3_style_comps(theme_new_dir)
    if b3_before < 6:
        _download_comp_refs(jan, refs, max_n=max_comps)
    comps_all = resolve_comps_for_dir(theme_new_dir)
    comps_rank = [
        p
        for p in comps_all
        if not p.name.upper().startswith("MANUAL_")
    ]
    cand = write_candidates_shortlist(
        out_dir=theme_new_dir,
        truth=truth,
        comps=comps_rank,
        top_k=6,
        force=refresh_candidates,
    )
    merged = merge_manual_into_candidates(
        list(cand.get("candidates") or []), theme_new_dir
    )
    cand = {**cand, "candidates": merged}
    (theme_new_dir / CANDIDATES_NAME).write_text(
        json.dumps(cand, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    msg = _ensure_messages(jan, theme_new_dir)
    product_name = ""
    try:
        product_name = str(json.loads(msg.read_text(encoding="utf-8")).get("productName") or "")
    except Exception:
        pass
    root = compose_root(Path(work_root))
    compose = latest_compose_dir(root, jan) if root.is_dir() else None
    b3_n = count_b3_style_comps(theme_new_dir)
    manual_n = len(list_manual_paths(theme_new_dir))
    session = {
        "jan": jan,
        "productName": product_name,
        "phase": "pick",
        "themeNewDir": str(theme_new_dir.resolve()),
        "composeDir": str(compose.resolve()) if compose else "",
        "truth": str(truth),
        "builtAt": datetime.now(timezone.utc).isoformat(),
        "gen": {"state": "idle", "message": "", "ok": None},
        "competitorCount": b3_n,
        "manualCount": manual_n,
        "needManualHint": b3_n == 0,
    }
    (theme_new_dir / SESSION_NAME).write_text(
        json.dumps(session, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return {**session, "candidates": merged}


def _load_session(theme_new_dir: Path) -> Dict[str, Any]:
    p = Path(theme_new_dir) / SESSION_NAME
    if p.is_file():
        return json.loads(p.read_text(encoding="utf-8"))
    return {}


def _save_session(theme_new_dir: Path, session: Dict[str, Any]) -> None:
    (Path(theme_new_dir) / SESSION_NAME).write_text(
        json.dumps(session, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _sync_manual_into_candidates(theme_new_dir: Path) -> Dict[str, Any]:
    """手動追加／削除後に候補 JSON を差し替える（B-③順位は維持）。"""
    cand = load_candidates_meta(theme_new_dir)
    existing = list(cand.get("candidates") or [])
    on_disk = {p.name for p in list_manual_paths(theme_new_dir)}
    kept = []
    for c in existing:
        name = str(c.get("name") or "")
        if name.upper().startswith("MANUAL_"):
            if name in on_disk:
                kept.append(c)
            continue
        kept.append(c)
    merged = merge_manual_into_candidates(kept, theme_new_dir)
    out = {**cand, "candidates": merged, "manualSyncedAt": datetime.now(timezone.utc).isoformat()}
    (Path(theme_new_dir) / CANDIDATES_NAME).write_text(
        json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return out


def _refresh_session_counts(theme_new_dir: Path, session: Dict[str, Any]) -> Dict[str, Any]:
    b3_n = count_b3_style_comps(theme_new_dir)
    manual_n = len(list_manual_paths(theme_new_dir))
    session = {
        **session,
        "competitorCount": b3_n,
        "manualCount": manual_n,
        "needManualHint": b3_n == 0,
    }
    _save_session(theme_new_dir, session)
    return session


def _theme_ids_from_messages(msg_path: Path) -> str:
    data = json.loads(Path(msg_path).read_text(encoding="utf-8"))
    ids = [str(int(r["themeId"])) for r in (data.get("rows") or []) if r.get("themeId")]
    return ",".join(ids)


def run_theme_new_generate(theme_new_dir: Path, jan: str) -> Dict[str, Any]:
    msg = _ensure_messages(jan, theme_new_dir)
    data = json.loads(msg.read_text(encoding="utf-8"))
    text_color = str(data.get("textColor") or "1")
    image_themes = _theme_ids_from_messages(msg)
    cmd = [
        sys.executable,
        str(SCRIPT_DIR / "dry_sub_overview_copy_test.py"),
        "--jan",
        jan,
        "--out-dir",
        str(theme_new_dir),
        "--messages-json",
        str(msg),
        "--reuse-comps-dir",
        str(Path(theme_new_dir) / "01_refs"),
        "--image-themes",
        image_themes,
        "--text-color",
        text_color,
        "--max-comps",
        "8",
        "--quality",
        "medium",
        "--auto-export",
    ]
    LOG.info("generate: %s", " ".join(cmd))
    proc = subprocess.run(
        cmd, cwd=str(SCRIPT_DIR), capture_output=True, text=True, timeout=7200
    )
    return {
        "ok": proc.returncode == 0,
        "code": proc.returncode,
        "stdout": (proc.stdout or "")[-800:],
        "stderr": (proc.stderr or "")[-800:],
    }


def _prior_pick_state(theme_new_dir: Path) -> Dict[str, Any]:
    """人間確定済みなら、前回の IDENTITY／SIZZLE／WRAP／副を返す（再オープン時の復元用）。"""
    meta = load_food_canons(theme_new_dir)
    if str(meta.get("confirmedBy") or "") != "human":
        return {"hasPrior": False}
    identity = str(meta.get("identityName") or meta.get("primaryName") or "").strip()
    chosen = str(meta.get("chosenName") or "").strip()
    # 一片由来の盛り差し替えは IDENTITY主のレ点復元に使わない（元画像を優先）
    if _is_remake_identity_name(identity) and chosen and not _is_remake_identity_name(chosen):
        identity = chosen
    sizzle = str(meta.get("sizzleName") or "").strip()
    wrap = str(meta.get("wrapName") or "").strip()
    secondary = [str(x).strip() for x in (meta.get("secondaryNames") or []) if str(x).strip()]
    package_only = bool(meta.get("packageOnly")) or identity in (
        "__package_only__",
        "PACKAGE_ONLY",
    )
    if package_only:
        identity = "__package_only__"
        sizzle = ""
        wrap = ""
        secondary = []
    canon = existing_canon_path(theme_new_dir)
    sizzle_path = existing_sizzle_path(theme_new_dir)
    wrap_path = existing_wrap_path(theme_new_dir)
    secs = secondary_paths(theme_new_dir)
    return {
        "hasPrior": True,
        "identity": identity,
        "sizzle": sizzle,
        "wrap": wrap,
        "secondary": secondary,
        "packageOnly": package_only,
        "builtAt": str(meta.get("builtAt") or ""),
        "hasCanonFile": bool(canon and Path(canon).is_file()),
        "hasSizzleFile": bool(sizzle_path and Path(sizzle_path).is_file()),
        "hasWrapFile": bool(wrap_path and Path(wrap_path).is_file()),
        "secondaryCount": len(secs),
    }


def _sort_candidates_for_prior(
    candidates: List[Dict[str, Any]], prior: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """前回選んだ画像を先頭に（確認しやすくする）。"""
    if not prior.get("hasPrior"):
        return list(candidates)
    identity = str(prior.get("identity") or "")
    sizzle = str(prior.get("sizzle") or "")
    wrap = str(prior.get("wrap") or "")
    secondary = set(prior.get("secondary") or [])

    def rank(c: Dict[str, Any]) -> tuple:
        name = str(c.get("name") or "")
        if name == identity:
            return (0, name)
        if name == sizzle:
            return (1, name)
        if name == wrap:
            return (2, name)
        if name in secondary:
            return (2, name)
        if str(c.get("source") or "") == "manual" or name.upper().startswith("MANUAL_"):
            return (3, name)
        return (4, name)

    return sorted(list(candidates), key=rank)


def _mix_parts_from_session(session: Dict[str, Any]) -> List[Dict[str, Any]]:
    raw = session.get("parts") or []
    out: List[Dict[str, Any]] = []
    for p in raw:
        if not isinstance(p, dict):
            continue
        nb = str(p.get("nameBase") or "").strip()
        if not nb:
            continue
        out.append(
            load_part_bundle(
                name_base=nb,
                jan_for_folder=str(p.get("jan") or ""),
                parent_sku=str(p.get("parentSku") or ""),
                bag_path=str(p.get("bag") or p.get("packageTruth") or ""),
            )
        )
    return out


def _mix_parent_bag(session: Dict[str, Any]) -> Optional[Path]:
    return resolve_package_bag_for_name(
        name_base=str(session.get("nameBase") or session.get("productName") or "").strip(),
        parent_sku=str(session.get("parentSku") or "").strip(),
        jan_for_folder=str(session.get("jan") or ""),
        bag_path=str(session.get("bag") or session.get("packageTruth") or ""),
    )


def build_mix_confirm_html(session: Dict[str, Any]) -> str:
    display_name = str(session.get("nameBase") or session.get("productName") or "組み合わせ").strip()
    parent_sku = str(session.get("parentSku") or "").strip()
    mix_jan = str(session.get("jan") or "").strip()
    size_banner = master_size_banner(mix_jan, parent_sku)
    mix_bag = _mix_parent_bag(session)
    parts = _mix_parts_from_session(session)
    blocks = []
    if mix_bag is not None:
        blocks.append(
            """
            <div class="part ok">
              <h2>組み合わせの袋</h2>
              <p class="hint">このセットそのものの袋の正本です。</p>
              <div class="prior-thumbs"><div><img src="/mix/bag" alt="組み合わせ袋"/><div>袋の正本</div></div></div>
            </div>
            """
        )
    else:
        blocks.append(
            "<p class='warn'>組み合わせの袋がまだありません。下は構成の単品です。</p>"
        )
    all_ok = True
    any_no_wrap = False
    for i, p in enumerate(parts):
        ok = bool(p.get("ok"))
        has_wrap = bool(p.get("wrap"))
        if not ok:
            all_ok = False
        if ok and not has_wrap:
            any_no_wrap = True
        thumbs = []
        if p.get("bag"):
            thumbs.append(f'<div><img src="/part/{i}/bag" alt="袋"/><div>袋の正本</div></div>')
        if p.get("identity"):
            thumbs.append(f'<div><img src="/part/{i}/canon" alt="中身"/><div>中身の正</div></div>')
        if p.get("sizzle"):
            thumbs.append(f'<div><img src="/part/{i}/sizzle" alt="食欲"/><div>食欲</div></div>')
        if has_wrap:
            thumbs.append(f'<div><img src="/part/{i}/wrap" alt="個包装"/><div>個包装</div></div>')
        for j, _s in enumerate(p.get("secondary") or []):
            thumbs.append(
                f'<div><img src="/part/{i}/sec/{j}" alt="副"/><div>副 {j+1}</div></div>'
            )
        miss = ""
        if not ok:
            miss = (
                "<p class='warn'>まだ決まっていません。"
                "下の「この単品を選び直す」で見本画面を開き、確定してください。</p>"
            )
        elif not has_wrap:
            miss = (
                "<p class='warn'>個包装（WRAP）が未登録です。"
                "紐付けるなら「この単品を選び直す」→ WRAP を選んで保存。</p>"
            )
        blocks.append(
            f"""
            <div class="part {'ok' if ok else 'miss'}">
              <h2>構成{i+1}　{p.get('nameBase') or ''}</h2>
              <p class="hint">{p.get('label') or ''}{' ／ 親SKU ' + p.get('parentSku') if p.get('parentSku') else ''}
              {' ／ 個包装あり' if has_wrap else ' ／ 個包装なし'}</p>
              {miss}
              <div class="prior-thumbs">{''.join(thumbs) or '<p class="hint">写真なし</p>'}</div>
              <p style="margin-top:8px">
                <button type="button" class="btnOpenPart" data-part="{i}">
                  この単品を選び直す（個包装・中身）
                </button>
              </p>
            </div>
            """
        )
    btn = ""
    if str(session.get("context") or "").strip().lower() == "main":
        wrap_note = (
            '<p class="warn">個包装が無い構成があります。MAINで個包装を出すなら、上のボタンで単品を選び直してください。</p>'
            if any_no_wrap
            else ""
        )
        if all_ok:
            btn = (
                wrap_note
                + '<p class="hint" style="background:#eef6ff;border:1px solid #8af;padding:10px;border-radius:8px;">'
                "中身の確定は揃っています。確認だけなら "
                "<b>変更なしで閉じる</b> → B-④「次へ」。"
                "個包装を足す／変えるときは各構成の「この単品を選び直す」。"
                "</p>"
                '<p><button type="button" id="btnMainClose">変更なしで閉じる</button> '
                '<span id="status"></span></p>'
            )
        else:
            btn = (
                '<p class="warn">欠けている構成があります。各構成の「この単品を選び直す」で揃えてから、この画面を再読込してください。</p>'
                '<p><button type="button" id="btnMainClose">閉じる</button> <span id="status"></span></p>'
            )
    return f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>食材見本（組み合わせ）— {display_name}</title>
<style>
body{{font-family:sans-serif;margin:16px;background:#f6f3ee;color:#222}}
h1{{font-size:18px}} h2{{font-size:16px;margin:0 0 6px}}
.hint{{font-size:13px;color:#555;line-height:1.45}}
.warn{{background:#fff3cd;border:2px solid #c90;padding:12px;border-radius:8px;margin:12px 0;font-size:13px}}
.part{{background:#fff;border:2px solid #ccc;border-radius:8px;padding:12px;margin:12px 0}}
.part.ok{{border-color:#2e7d32}}
.part.miss{{border-color:#c90}}
.prior-thumbs{{display:flex;flex-wrap:wrap;gap:12px;margin-top:8px}}
.prior-thumbs img{{height:120px;border:1px solid #ccc;border-radius:4px;cursor:zoom-in}}
.prior-thumbs div{{font-size:11px;color:#333;text-align:center}}
.bag-pick{{background:#fff;border:2px solid #1565c0;border-radius:8px;padding:12px;margin:12px 0}}
.bag-grid .bag-card{{display:block;background:#fafafa;border:2px solid #ddd;border-radius:6px;padding:8px;cursor:pointer}}
.bag-grid .bag-card.selected{{border-color:#1565c0;box-shadow:0 0 0 2px rgba(21,101,192,.2)}}
.bag-grid .bag-img{{width:100%;height:140px;object-fit:contain;background:#f0ece6;border:1px solid #eee;display:block}}
.bag-grid .cand-meta{{font-size:11px;color:#555;margin-top:6px;word-break:break-all}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin:8px 0}}
button{{padding:10px 14px;margin-right:8px;cursor:pointer}}
#lightbox{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:9;align-items:center;justify-content:center}}
#lightbox.on{{display:flex}}
#lightbox img{{max-width:96vw;max-height:96vh}}
</style></head><body>
<h1>食材見本（組み合わせ）— {display_name}</h1>
{size_banner}
<p class="hint">親SKU {parent_sku or '（なし）'}<br/>
構成ごとの中身は、単品で確定した写真です。最上段が組み合わせの袋、その下が各単品の袋と中身です。袋だけはこの画面で画像を見ながら選び直せます。</p>
{_bag_pick_section_html(session)}
{''.join(blocks) or '<p class="warn">構成がありません</p>'}
{btn}
<div id="lightbox"><img id="lbImg"/></div>
<script>
document.querySelectorAll('.prior-thumbs img').forEach(img=>{{
  img.onclick=()=>{{ document.getElementById('lbImg').src=img.src; document.getElementById('lightbox').classList.add('on'); }};
}});
document.getElementById('lightbox').onclick=()=>document.getElementById('lightbox').classList.remove('on');
const closeBtn=document.getElementById('btnMainClose');
if(closeBtn) closeBtn.onclick=()=>{{
  document.getElementById('status').textContent='B-④ の「次へ」へ進んでください';
  try {{ window.close(); }} catch(e) {{}}
}};
document.querySelectorAll('.btnOpenPart').forEach(btn=>{{
  btn.onclick=async()=>{{
    const i=Number(btn.getAttribute('data-part')||'-1');
    const st=document.getElementById('status');
    if(st) st.textContent='単品の見本画面を開いています…';
    const res=await fetch('/api/open-part-gate',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{partIndex:i}})}});
    const j=await res.json();
    if(st) st.textContent=j.message||'';
    if(j.ok && j.url) window.open(j.url,'_blank');
    else if(j.message) alert(j.message);
  }};
}});
{_bag_pick_js()}
</script>
</body></html>
"""


def master_size_banner(jan: str, parent_sku: str = "") -> str:
    """見本確認の前に、マスタのサイズ列を出す。面の縦横が無ければその旨。"""
    key = "%s|%s" % (jan, parent_sku)
    cached = getattr(master_size_banner, "_cache", {})
    if key in cached:
        return cached[key]
    text = _master_size_text(str(jan or ""), str(parent_sku or ""))
    if not text:
        banner = '<p class="hint">マスタのサイズは、この画面では読めませんでした。</p>'
    else:
        banner = '<p class="warn">マスタのサイズ（確認用）<br/>%s</p>' % html_lib.escape(text).replace(
            "\n", "<br/>"
        )
    cached[key] = banner
    master_size_banner._cache = cached
    return banner


def _master_size_text(jan: str, parent_sku: str) -> str:
    digits = "".join(ch for ch in str(jan or "") if ch.isdigit())
    if not digits:
        return ""
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "common_ss_view"))
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "competitor_store"))
        from client import sheets_service
        from registry import load_registry

        reg = load_registry()
        sid = str(reg["files"]["自動出品260221_セラセン出力"]["id"])
        svc = sheets_service(write=False, interactive=False)
        if not svc:
            return "マスタを読めませんでした。"
        title = "▼商品マスタ(人間作業用)"
        hdr_rng = "'%s'!8:8" % title.replace("'", "''")
        hdr = (
            svc.spreadsheets().values().get(spreadsheetId=sid, range=hdr_rng).execute().get("values")
            or [[]]
        )[0]
        want = {
            "JANコード": None,
            "商品名ベース": None,
            "親SKU": None,
            "子SKU": None,
            "(商品本体サイズ)": None,
            "セット構成_幅_cm": None,
        }
        for i, name in enumerate(hdr):
            if name in want and want[name] is None:
                want[name] = i
        if want["JANコード"] is None:
            return "マスタに JANコード 列がありません。"

        def letter(n: int) -> str:
            s = ""
            while n:
                n, r = divmod(n - 1, 26)
                s = chr(65 + r) + s
            return s

        ranges = []
        order = []
        for name, idx in want.items():
            if idx is None:
                continue
            col = letter(idx + 1)
            ranges.append("'%s'!%s:%s" % (title.replace("'", "''"), col, col))
            order.append(name)
        got = (
            svc.spreadsheets()
            .values()
            .batchGet(spreadsheetId=sid, ranges=ranges)
            .execute()
            .get("valueRanges")
            or []
        )
        cols = {}
        for name, block in zip(order, got):
            cols[name] = block.get("values") or []
        jans = cols.get("JANコード") or []
        lines = []
        for ri, cell in enumerate(jans):
            val = cell[0] if cell else ""
            if digits not in str(val):
                continue
            child = ""
            ch = cols.get("子SKU") or []
            if ri < len(ch) and ch[ri]:
                child = str(ch[ri][0]).strip()
            if child:
                continue
            sku = ""
            sk = cols.get("親SKU") or []
            if ri < len(sk) and sk[ri]:
                sku = str(sk[ri][0]).strip()
            if parent_sku and sku and sku != parent_sku:
                continue

            def cell_at(col_name: str) -> str:
                col = cols.get(col_name) or []
                if ri < len(col) and col[ri]:
                    return str(col[ri][0]).strip()
                return ""

            name = cell_at("商品名ベース") or sku or digits
            body = cell_at("(商品本体サイズ)") or "（空）"
            width = cell_at("セット構成_幅_cm") or "（空）"
            lines.append(
                "%s: 商品本体サイズ＝%s。セット構成の幅cm＝%s。面の縦と横が両方ある列ではないので、食器との比には使っていません。"
                % (name, body, width)
            )
            if parent_sku and sku == parent_sku:
                break
        if not lines:
            return "このJANの親行に、商品本体サイズとセット構成の幅は見つかりませんでした。"
        return "\n".join(lines)
    except Exception as e:
        LOG.warning("master size read failed: %s", e)
        return "マスタのサイズを読めませんでした。"


def build_wrap_hand_hold_html(session: Dict[str, Any]) -> str:
    """P3-W第2: WRAP確定後の手持ち候補→1枚採用。"""
    from b4_wrap_hand_hold import (
        list_candidate_files,
        load_manifest,
        selected_path,
    )

    jan = str(session.get("jan") or "")
    display_name = str(session.get("nameBase") or session.get("productName") or jan).strip() or jan
    parent_sku = str(session.get("parentSku") or "").strip()
    is_main = str(session.get("context") or "").strip().lower() == "main"
    theme_new_dir = Path(str(session.get("themeNewDir") or ""))
    man = load_manifest(theme_new_dir) if theme_new_dir.is_dir() else {}
    cands = [p.name for p in list_candidate_files(theme_new_dir)] if theme_new_dir.is_dir() else []
    if not cands:
        cands = [str(x) for x in (man.get("candidates") or []) if str(x).strip()]
    selected = str(man.get("selected") or "")
    sel_p = selected_path(theme_new_dir) if theme_new_dir.is_dir() else None
    cards: List[str] = []
    for name in cands:
        checked = "checked" if name == selected else ""
        cls = "hh-card selected" if name == selected else "hh-card"
        cards.append(
            f"""
            <label class="{cls}">
              <img src="/wrap-hand-hold/{html_lib.escape(name)}" alt="{html_lib.escape(name)}" class="hh-img"/>
              <input type="radio" name="hh" value="{html_lib.escape(name)}" {checked}/>
              <div class="cand-meta"><code>{html_lib.escape(name)}</code></div>
            </label>
            """
        )
    cards_html = "\n".join(cards) or "<p class='hint'>まだ候補がありません。「候補を生成」を押してください。</p>"
    selected_note = (
        f"<p class='info'>採用済み: <code>{html_lib.escape(Path(sel_p).name if sel_p else selected)}</code></p>"
        if (sel_p or selected)
        else ""
    )
    after_hint = (
        "採用したらこのタブを閉じて B-④ の「次へ」。"
        if is_main
        else "採用後は下のボタンでサブ画像生成またはレビューへ進めます。"
    )
    continue_btn = ""
    if not is_main:
        continue_btn = """
  <button type="button" id="btnContinueGen" style="font-weight:bold">採用して画像生成へ</button>
  <button type="button" id="btnContinueReview">採用してレビューへ（生成しない）</button>
"""
    return f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>個包装手持ち — {html_lib.escape(display_name)}</title>
<style>
body{{font-family:sans-serif;margin:16px;background:#f6f3ee;color:#222}}
h1{{font-size:18px}}
.hint{{font-size:13px;color:#555;line-height:1.45}}
.info{{background:#eef6ff;border:1px solid #8af;padding:10px;border-radius:8px;margin:12px 0;font-size:13px}}
.err{{display:none;background:#fdecea;border:2px solid #c62828;color:#b71c1c;padding:12px;border-radius:8px;margin:12px 0;font-size:14px;font-weight:bold}}
.err.on{{display:block}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;margin:12px 0}}
.hh-card{{display:block;background:#fff;border:2px solid #ddd;border-radius:8px;padding:8px;cursor:pointer}}
.hh-card.selected{{border-color:#00695c;box-shadow:0 0 0 2px rgba(0,105,92,.2)}}
.hh-img{{width:100%;height:200px;object-fit:contain;background:#f0ece6;border:1px solid #eee;display:block}}
.cand-meta{{font-size:11px;color:#555;margin-top:6px}}
button{{padding:10px 14px;margin:4px 8px 4px 0;cursor:pointer}}
#status{{margin-left:8px}}
</style></head><body>
<h1>個包装の手持ち素材 — {html_lib.escape(display_name)}</h1>
<p class="hint">親SKU {html_lib.escape(parent_sku or '（なし）')} ／ JAN {html_lib.escape(jan)}<br/>
WRAP確定のあと、手持ち候補から<strong>1枚だけ</strong>採用します（型・文案の発明は不要）。<br/>
背景は白〜生成り＋植物ぼけ。個包装は親指サイズ（巨大化禁止）。<br/>
{html_lib.escape(after_hint)}</p>
{selected_note}
<div id="errBox" class="err" role="alert"></div>
<div class="grid">{cards_html}</div>
<p>
  <button type="button" id="btnGen">候補を生成（未作成時）</button>
  <button type="button" id="btnForce">作り直す</button>
  <button type="button" id="btnSelect" style="font-weight:bold">この1枚を採用</button>
  <button type="button" id="btnSkip">手持ちなしで進む</button>
  {continue_btn}
  <a href="/pick">← 見本の選び直し</a>
  <span id="status"></span>
</p>
<script>
function showErr(m){{const e=document.getElementById('errBox');e.textContent=m||'';e.classList.toggle('on',!!m);}}
function clearErr(){{showErr('');}}
function picked(){{
  const el=document.querySelector('input[name=hh]:checked');
  return el?el.value:'';
}}
document.getElementById('btnGen').onclick=async()=>{{
  clearErr();
  document.getElementById('status').textContent='生成中（数分かかることがあります）…';
  const res=await fetch('/api/wrap-hand-hold/generate',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{force:false}})}});
  const j=await res.json();
  if(!j.ok){{showErr(j.message||'生成失敗');document.getElementById('status').textContent='';return;}}
  document.getElementById('status').textContent=j.message||'完了';
  location.reload();
}};
document.getElementById('btnForce').onclick=async()=>{{
  clearErr();
  if(!confirm('既存候補を捨てて作り直します。よろしいですか？')) return;
  document.getElementById('status').textContent='再生成中…';
  const res=await fetch('/api/wrap-hand-hold/generate',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{force:true}})}});
  const j=await res.json();
  if(!j.ok){{showErr(j.message||'生成失敗');document.getElementById('status').textContent='';return;}}
  location.reload();
}};
document.getElementById('btnSelect').onclick=async()=>{{
  clearErr();
  const name=picked();
  if(!name){{showErr('候補を1枚選んでください');return;}}
  document.getElementById('status').textContent='採用保存中…';
  const res=await fetch('/api/wrap-hand-hold/select',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{name}})}});
  const j=await res.json();
  if(!j.ok){{showErr(j.message||'採用失敗');document.getElementById('status').textContent='';return;}}
  document.getElementById('status').textContent=j.message||'採用しました';
  {'try { window.close(); } catch(e) {}' if is_main else 'setTimeout(()=>location.reload(), 400);'}
}};
document.getElementById('btnSkip').onclick=async()=>{{
  clearErr();
  const res=await fetch('/api/wrap-hand-hold/skip',{{method:'POST'}});
  const j=await res.json();
  if(!j.ok){{showErr(j.message||'スキップ失敗');return;}}
  document.getElementById('status').textContent='手持ちなしで進めます';
  {'try { window.close(); } catch(e) {}' if is_main else ''}
}};
{'' if is_main else '''
document.getElementById('btnContinueGen').onclick=async()=>{
  clearErr();
  let name=picked();
  if(name){
    let res=await fetch('/api/wrap-hand-hold/select',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})});
    let j=await res.json();
    if(!j.ok){ showErr(j.message||'採用失敗'); return; }
  }
  document.getElementById('status').textContent='画像生成中…';
  let res=await fetch('/api/generate',{method:'POST'});
  let j=await res.json();
  document.getElementById('status').textContent=j.message||'';
  if(j.ok) setTimeout(()=>location.reload(), 600);
};
document.getElementById('btnContinueReview').onclick=async()=>{
  clearErr();
  let name=picked();
  if(name){
    let res=await fetch('/api/wrap-hand-hold/select',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})});
    let j=await res.json();
    if(!j.ok){ showErr(j.message||'採用失敗'); return; }
  }
  let res=await fetch('/api/to-review',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({requireImages:true})});
  let j=await res.json();
  if(!j.ok){ showErr(j.message||'レビューへ進めません'); return; }
  location.reload();
};
'''}
// 候補ゼロなら自動生成を1回試す
(async()=>{{
  if({('true' if not cands else 'false')}){{
    document.getElementById('status').textContent='候補を自動生成中…';
    const res=await fetch('/api/wrap-hand-hold/generate',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{force:false}})}});
    const j=await res.json();
    if(j.ok) location.reload();
    else {{ showErr(j.message||'自動生成に失敗'); document.getElementById('status').textContent=''; }}
  }}
}})();
</script>
</body></html>
"""


def _wrap_inline_section_html(
    session: Dict[str, Any],
    *,
    theme_new_dir: Path,
    prior_wrap: str,
    draft_hand: str = "",
    draft_single: str = "",
    draft_remake_note: str = "",
) -> str:
    """個品見本画面内の個包装ブロック（あり選択で開く）。"""
    from b4_wrap_asset_rules import wrap_asset_lane_enabled
    from b4_wrap_hand_hold import (
        list_candidate_files,
        list_single_files,
        load_manifest,
        out_dir_for,
        recommend_wrap_assets,
    )

    if not wrap_asset_lane_enabled():
        return ""
    man = load_manifest(theme_new_dir) if theme_new_dir.is_dir() else {}
    hands = [p.name for p in list_candidate_files(theme_new_dir)] if theme_new_dir.is_dir() else []
    singles = [p.name for p in list_single_files(theme_new_dir)] if theme_new_dir.is_dir() else []
    rec = recommend_wrap_assets(theme_new_dir) if theme_new_dir.is_dir() else {}
    rec_hand = str(rec.get("handHold") or man.get("recommendedHand") or (hands[0] if hands else ""))
    # 保存済み単体（selectedSingle／composeWrap／canons）を先頭より優先。でないと常に左端に戻る
    canon_wrap = ""
    try:
        from food_unit_canon import load_food_canons

        cm = load_food_canons(theme_new_dir) if theme_new_dir.is_dir() else {}
        canon_wrap = str(
            (cm.get("wrap") or {}).get("primaryName") or cm.get("wrapName") or ""
        ).strip()
    except Exception:
        canon_wrap = ""
    rec_single = str(
        rec.get("wrapSingle")
        or man.get("selectedSingle")
        or man.get("composeWrap")
        or (canon_wrap if canon_wrap.startswith("wrap_single_") else "")
        or man.get("recommendedSingle")
        or (singles[0] if singles else "")
    )
    # 作り直し前の下書き採用を優先（リロードでおすすめ先頭に戻さない）
    if draft_hand and draft_hand in hands:
        rec_hand = draft_hand
    if draft_single and draft_single in singles:
        rec_single = draft_single
    wrap_on = bool(prior_wrap) or bool(hands) or bool(singles) or bool(
        existing_wrap_path(theme_new_dir) if theme_new_dir.is_dir() else None
    )
    on_checked = "checked" if wrap_on else ""
    off_checked = "" if wrap_on else "checked"
    disp = "block" if wrap_on else "none"

    _hh_root = out_dir_for(theme_new_dir) if theme_new_dir.is_dir() else None

    def _cards(
        names: List[str],
        radio_name: str,
        remake_class: str,
        recommended: str,
        *,
        show_ring_preview: bool = False,
    ) -> str:
        out: List[str] = []
        for name in names:
            chk = "checked" if name == recommended else ""
            sug = " <b style='color:#0a7'>おすすめ</b>" if name == recommended else ""
            esc = html_lib.escape(name)
            ring_block = ""
            if show_ring_preview and str(name).startswith("hand_hold_"):
                bust = "0"
                if _hh_root is not None:
                    fp = _hh_root / name
                    if fp.is_file():
                        bust = str(int(fp.stat().st_mtime))
                ring_block = f"""
              <div class="hh-ring-lab">金枠イメージ（採用判断用・課金なし）</div>
              <img src="/wrap-hand-hold-ring/{esc}?v={html_lib.escape(bust)}" class="hh-img hh-ring-preview" alt="金枠プレビュー"/>
"""
            out.append(
                f"""
            <div class="hh-card{' selected' if name == recommended else ''}">
              <img src="/wrap-hand-hold/{esc}" class="hh-img" alt=""/>
              {ring_block}
              <label><input type="radio" name="{radio_name}" value="{esc}" {chk}/> 採用{sug}</label>
              <label><input type="checkbox" class="{remake_class}" value="{esc}"/> 作り直し対象</label>
              <div class="cand-meta"><code>{esc}</code></div>
            </div>"""
            )
        return "\n".join(out) or "<p class='hint'>まだありません。「候補を生成」で初回作成してください。</p>"

    return f"""
<section class="wrap-gate" id="wrapGate">
  <h2 style="font-size:15px;margin:16px 0 8px">① 個包装はありますか</h2>
  <p>
    <label class="wrap-lab"><input type="radio" name="wrapHas" value="0" id="wrapHasOff" {off_checked}/> 個包装なし</label>
    <label class="wrap-lab"><input type="radio" name="wrapHas" value="1" id="wrapHasOn" {on_checked}/> 個包装あり</label>
  </p>
  <div id="wrapPanel" style="display:{disp}">
    <div class="prior" id="wrapGuide">
      <b>個包装ありのときの流れ</b><br/>
      ② 下で個包装の見本をコピペ／URL登録（または候補カードで WRAP を選ぶ）<br/>
      ③ 「候補を生成」＝初回（または不足分）。近いものだけ「作り直し対象」にレ点→各タイプの作り直し<br/>
      ④ 手持ち＝金枠用／単体＝MAINが読む個包装の正（寄せ集めは使わない）<br/>
      ⑤ IDENTITY など個品見本も選ぶ<br/>
      ⑥ いちばん下の保存で、個品見本と個包装見本をまとめて確定
    </div>
    <h3 style="font-size:14px">② 個包装見本の登録</h3>
    <p class="hint">個包装の写真をここに Ctrl+V。URL・ファイルでも可。登録後に候補カードへ載り、WRAP として使えます。</p>
    <div id="wrapPasteBox" class="paste-box" tabindex="0">個包装見本：ここをクリックして Ctrl+V</div>
    <div>
      <input type="url" id="wrapUrlIn" placeholder="https://… 個包装画像の直URL"/>
      <button type="button" id="btnWrapUrl">URLから個包装を追加</button>
      <input type="file" id="wrapFileIn" accept="image/*"/>
      <button type="button" id="btnWrapFile">ファイルから追加</button>
    </div>
    <h3 style="font-size:14px;margin-top:14px">③④ 手持ち／単体の生成と採用</h3>
    <p class="hint">
      <b>候補を生成</b>＝まだ無いとき（または不足）にまとめて作る。既存があれば再利用（課金なし）。<br/>
      <b>作り直し</b>＝イメージが近いものだけ「作り直し対象」にレ点 → 手持ち／単体を別々に作り直し（レ点分だけ課金）。
    </p>
    <p>
      <button type="button" id="btnWrapGen">候補を生成（初回・不足分）</button>
      <span id="wrapGenStatus"></span>
    </p>
    <label style="display:block;font-size:13px;margin:8px 0 4px"><b>作り直しの要望</b>（大きさ・質感など。手持ち／単体の両方に使います）</label>
    <textarea id="wrapRemakeNote" rows="2" maxlength="400" placeholder="例: 実物より大きく。指に対して小さく見えないように。フィルムの質感をもう少し厚く。" style="width:100%;max-width:640px;box-sizing:border-box;padding:8px;font-size:13px">{html_lib.escape(draft_remake_note)}</textarea>
    <h4 style="font-size:13px">手持ちタイプ（採用＝ラジオ／作り直し＝レ点）</h4>
    <p class="hint">上＝元写真／下＝MAIN左下金枠にはめたイメージ（Pillow合成・課金なし）。金枠イメージを見て採用を決めてください。</p>
    <div class="grid" id="wrapHandGrid">{_cards(hands, "handHoldPick", "remakeHand", rec_hand, show_ring_preview=True)}</div>
    <p><button type="button" id="btnRemakeHand">手持ちのレ点だけ作り直す</button></p>
    <h4 style="font-size:13px">単体タイプ（採用＝ラジオ／作り直し＝レ点）※MAINの個包装見本はここ</h4>
    <p class="hint">保存すると、選んだ単体が FOOD_UNIT_WRAP になります（寄せ集め・小さすぎる貼付見本は使いません。元見本は種として残します）。</p>
    <div class="grid" id="wrapSingleGrid">{_cards(singles, "wrapSinglePick", "remakeSingle", rec_single)}</div>
    <p><button type="button" id="btnRemakeSingle">単体のレ点だけ作り直す</button></p>
  </div>
</section>
"""


def _piece_set_lang_path(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / "food_piece_set5_virt" / "piece_set_lang.json"


def _identity_plating_lang_path(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / "food_piece_set5_virt" / "identity_plating_lang.json"


def _identity_plating_instruction_path(theme_new_dir: Path) -> Path:
    """②生成時の指示（生成後盛り言語化と同一欄）。"""
    return Path(theme_new_dir) / "food_piece_set5_virt" / "identity_plating_from_pieces.txt"


_REMAKE_IDENTITY_NAMES = frozenset(
    {
        "PIECE_PLATING_HUMAN.jpg",
        "PIECE_PLATING_HUMAN.jpeg",
        "PIECE_PLATING_HUMAN.png",
        "PIECE_PLATING_HUMAN.webp",
        "FOOD_UNIT_CANON.jpg",
        "FOOD_UNIT_CANON.jpeg",
        "FOOD_UNIT_CANON.png",
        "FOOD_UNIT_CANON.webp",
    }
)


def _is_remake_identity_name(name: str) -> bool:
    n = Path(str(name or "").strip()).name
    return n.upper() in {x.upper() for x in _REMAKE_IDENTITY_NAMES}


def _ref_rel_if_exists(theme_new_dir: Path, name: str) -> str:
    name = Path(str(name or "").strip()).name
    if not name:
        return ""
    p = Path(theme_new_dir) / "01_refs" / name
    if p.is_file():
        return f"01_refs/{name}"
    return ""


def _load_identity_plating_obj(theme_new_dir: Path) -> Dict[str, Any]:
    p = _identity_plating_lang_path(theme_new_dir)
    if not p.is_file():
        return {}
    try:
        obj = json.loads(p.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}
    return obj if isinstance(obj, dict) else {}


def resolve_identity_plating_panels(
    theme_new_dir: Path, meta: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """①元IDENTITY画像＋言語化／②生成後盛り＋生成指示 を分ける（SKU非焼付）。"""
    theme_new_dir = Path(theme_new_dir)
    meta = dict(meta or load_food_canons(theme_new_dir) or {})
    ident = meta.get("identity") if isinstance(meta.get("identity"), dict) else {}
    chosen = str(
        meta.get("chosenName")
        or ident.get("primaryName")
        or ""
    ).strip()
    identity_name = str(
        meta.get("identityName") or meta.get("primaryName") or ""
    ).strip()
    note = str(meta.get("identityNote") or "").strip().lower()
    remake_note = "remake" in note or "piece-derived" in note

    source_name = ""
    if chosen and not _is_remake_identity_name(chosen):
        source_name = Path(chosen).name
    elif identity_name and not _is_remake_identity_name(identity_name):
        source_name = Path(identity_name).name
    else:
        for c in meta.get("candidates") or []:
            if not isinstance(c, dict):
                continue
            if str(c.get("role") or "").strip().lower() != "identity":
                continue
            n = Path(str(c.get("name") or "").strip()).name
            if n and not _is_remake_identity_name(n):
                source_name = n
                break

    gen_name = ""
    refs = theme_new_dir / "01_refs"
    for pat in ("PIECE_PLATING_HUMAN.*",):
        for p in sorted(refs.glob(pat)):
            if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp"):
                gen_name = p.name
                break
        if gen_name:
            break
    if not gen_name:
        for stem in ("FOOD_UNIT_CANON.jpg", "FOOD_UNIT_CANON.jpeg", "FOOD_UNIT_CANON.png"):
            if (refs / stem).is_file():
                # 元と同一バイトなら「生成後」として出さない
                same = False
                if source_name and (refs / source_name).is_file():
                    try:
                        same = (refs / source_name).read_bytes() == (refs / stem).read_bytes()
                    except Exception:
                        same = False
                if not same or remake_note or _is_remake_identity_name(identity_name):
                    if not same:
                        gen_name = stem
                break

    source_rel = _ref_rel_if_exists(theme_new_dir, source_name)
    gen_rel = _ref_rel_if_exists(theme_new_dir, gen_name)

    obj = _load_identity_plating_obj(theme_new_dir)
    prompt = str(obj.get("promptBlock") or "").strip()
    source_text = str(obj.get("sourcePromptBlock") or "").strip()
    gen_text = str(obj.get("generatedPromptBlock") or "").strip()
    lang_path_name = Path(str(obj.get("path") or "")).name
    lang_is_remake = _is_remake_identity_name(lang_path_name) or (
        bool(gen_name) and lang_path_name == gen_name
    )
    lang_is_source = bool(source_name) and lang_path_name == source_name

    if not source_text and prompt and lang_is_source:
        source_text = prompt
    if not gen_text and prompt and lang_is_remake:
        gen_text = prompt
    if not gen_text:
        ip = _identity_plating_instruction_path(theme_new_dir)
        if ip.is_file():
            try:
                gen_text = ip.read_text(encoding="utf-8-sig").strip()
            except Exception:
                gen_text = ""
    # 旧データ: 言語化が生成後由来なのに promptBlock だけある → ①は空、②へ
    if not source_text and prompt and not lang_is_source and lang_is_remake:
        source_text = ""
        if not gen_text:
            gen_text = prompt

    return {
        "sourceName": source_name,
        "sourceRel": source_rel,
        "sourceText": source_text,
        "generatedName": gen_name,
        "generatedRel": gen_rel,
        "generatedText": gen_text,
        "promptBlock": prompt,
        "langPathName": lang_path_name,
        "remakeNote": remake_note,
    }


def _load_piece_lang_text_map(theme_new_dir: Path) -> Dict[str, str]:
    """piece_set_0N.jpg → promptBlock。番号キーでも引ける。"""
    out: Dict[str, str] = {}
    p = _piece_set_lang_path(theme_new_dir)
    if not p.is_file():
        return out
    try:
        obj = json.loads(p.read_text(encoding="utf-8-sig"))
    except Exception:
        return out
    for it in obj.get("pieces") or []:
        if not isinstance(it, dict):
            continue
        name = str(it.get("name") or Path(str(it.get("path") or "")).name).strip()
        block = str(it.get("promptBlock") or "").strip()
        if not name or not block:
            continue
        out[name] = block
        m = re.search(r"(\d+)", Path(name).stem)
        if m:
            out[f"#{int(m.group(1))}"] = block
    return out


def _lang_text_for_piece_name(theme_new_dir: Path, name: str) -> str:
    mp = _load_piece_lang_text_map(theme_new_dir)
    if name in mp:
        return mp[name]
    m = re.search(r"(\d+)", Path(str(name or "")).stem)
    if not m:
        return ""
    n = int(m.group(1))
    return mp.get(f"piece_set_{n:02d}.jpg") or mp.get(f"#{n}") or ""


def _rebuild_piece_set_lang_pack(obj: Dict[str, Any]) -> str:
    pieces = [x for x in (obj.get("pieces") or []) if isinstance(x, dict)]
    lines = [
        "【一片セット全枚・機械言語化】場面の各片は、対応する番号の文章＋写真に合わせろ。",
        "血合いがある片にだけ血合いを出せ。無い片へ縁の濃褐色帯を付け足すな。",
        "「縁は濃褐色」を全片の共通ルールにするな（血合いの誤用禁止）。",
        "人間が修正するときは、画像とこの言語化の両方を直すこと。",
    ]
    for it in pieces:
        block = str(it.get("promptBlock") or "").strip()
        if block:
            lines.append(block)
    return "\n".join(lines)


def save_piece_lang_text(theme_new_dir: Path, *, name: str, text: str) -> Dict[str, Any]:
    """一片言語化を人間が直して保存（SKU非焼付・汎用）。"""
    name = str(name or "").strip()
    text = str(text or "").strip()
    if not name:
        raise ValueError("一片名が空です")
    if not text:
        raise ValueError("言語化が空です")
    path = _piece_set_lang_path(theme_new_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    obj: Dict[str, Any] = {}
    if path.is_file():
        try:
            obj = json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            obj = {}
    if not isinstance(obj, dict):
        obj = {}
    pieces = [x for x in (obj.get("pieces") or []) if isinstance(x, dict)]
    # piece_truth_01 → piece_set_01.jpg に寄せる
    target = name
    m = re.search(r"(\d+)", Path(name).stem)
    if m and not name.lower().startswith("piece_set_"):
        target = f"piece_set_{int(m.group(1)):02d}.jpg"
    hit = None
    for it in pieces:
        n0 = str(it.get("name") or Path(str(it.get("path") or "")).name).strip()
        if n0 == target or n0 == name:
            hit = it
            break
    if hit is None:
        hit = {"ok": True, "name": target, "fields": {}, "humanEdited": True}
        pieces.append(hit)
    hit["name"] = str(hit.get("name") or target)
    hit["promptBlock"] = text
    hit["humanEdited"] = True
    hit["humanEditedAt"] = datetime.now(timezone.utc).isoformat()
    obj["pieces"] = pieces
    obj["promptBlock"] = _rebuild_piece_set_lang_pack(obj)
    obj["humanEdited"] = True
    obj["humanEditedAt"] = datetime.now(timezone.utc).isoformat()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if path.is_file():
        shutil.copy2(path, path.with_name(f"piece_set_lang_prev_{stamp}.json"))
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"ok": True, "name": hit["name"], "path": str(path)}


def load_identity_plating_text(theme_new_dir: Path, *, role: str = "source") -> str:
    """role=source → ①元IDENTITY言語化（本線正）。role=generated → ②生成指示／生成後言語化。"""
    panels = resolve_identity_plating_panels(theme_new_dir)
    role_l = str(role or "source").strip().lower()
    if role_l in ("generated", "gen", "instruction", "result"):
        return str(panels.get("generatedText") or "").strip()
    # 本線互換: sourcePromptBlock が無ければ旧 promptBlock
    src = str(panels.get("sourceText") or "").strip()
    if src:
        return src
    return str(panels.get("promptBlock") or "").strip()


def save_identity_plating_text(
    theme_new_dir: Path, *, text: str, role: str = "source"
) -> Dict[str, Any]:
    text = str(text or "").strip()
    if not text:
        raise ValueError("IDENTITY盛り言語化が空です")
    role_l = str(role or "source").strip().lower()
    if role_l not in ("source", "generated", "gen", "instruction", "result"):
        raise ValueError("role は source か generated です")
    is_gen = role_l in ("generated", "gen", "instruction", "result")
    panels = resolve_identity_plating_panels(theme_new_dir)
    path = _identity_plating_lang_path(theme_new_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    obj = _load_identity_plating_obj(theme_new_dir)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if path.is_file():
        shutil.copy2(path, path.with_name(f"identity_plating_lang_prev_{stamp}.json"))
    now = datetime.now(timezone.utc).isoformat()
    if is_gen:
        obj["generatedPromptBlock"] = text
        gen_name = str(panels.get("generatedName") or "").strip()
        if gen_name:
            obj["generatedPath"] = f"01_refs/{gen_name}"
        obj["generatedHumanEdited"] = True
        obj["generatedHumanEditedAt"] = now
        # 生成指示ファイルにも同文を書く（②＝生成時の指示）
        ip = _identity_plating_instruction_path(theme_new_dir)
        if ip.is_file():
            shutil.copy2(ip, ip.with_name(f"identity_plating_from_pieces_prev_{stamp}.txt"))
        ip.write_text(text + "\n", encoding="utf-8")
        msg = "生成後盛り言語化（生成指示）を保存しました"
    else:
        obj["sourcePromptBlock"] = text
        # 本線同送の正は①
        obj["promptBlock"] = text
        src_name = str(panels.get("sourceName") or "").strip()
        if src_name:
            obj["path"] = str(theme_new_dir / "01_refs" / src_name)
            obj["sourceName"] = src_name
        obj["ok"] = True
        obj["humanEdited"] = True
        obj["humanEditedAt"] = now
        msg = "IDENTITY元画像の盛り言語化を保存しました"
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"ok": True, "path": str(path), "role": "generated" if is_gen else "source", "message": msg}


def _piece_inline_section_html(
    session: Dict[str, Any],
    *,
    theme_new_dir: Path,
    draft_piece_names: Optional[List[str]] = None,
) -> str:
    """個品見本画面内の一片見本ブロック（個包装と同型・有無＋貼付＋レ点）。"""
    try:
        from b5_piece_gate import (
            MAX_PIECE,
            detect_piece_has,
            list_human_piece_files,
            list_piece_display_files,
            list_piece_set_files,
            piece_gate_ui_enabled,
        )
    except Exception:
        return ""
    if not piece_gate_ui_enabled():
        return ""
    draft = session.get("draftPicks") if isinstance(session.get("draftPicks"), dict) else {}
    piece_on = detect_piece_has(theme_new_dir, draft=draft if draft else None)
    on_checked = "checked" if piece_on else ""
    off_checked = "" if piece_on else "checked"
    disp = "block" if piece_on else "none"
    checked_names = {
        str(x).strip()
        for x in (draft_piece_names or draft.get("pieceNames") or [])
        if str(x).strip()
    }
    display_files = (
        list_piece_display_files(theme_new_dir) if theme_new_dir.is_dir() else []
    )
    human_files = list_human_piece_files(theme_new_dir) if theme_new_dir.is_dir() else []
    set_files = (
        list_piece_set_files(theme_new_dir, max_n=MAX_PIECE)
        if theme_new_dir.is_dir()
        else []
    )

    def _piece_name_aliases(name: str) -> set[str]:
        n = str(name or "").strip()
        out = {n} if n else set()
        m = re.search(r"(\d+)", Path(n).stem)
        if not m:
            return out
        i = int(m.group(1))
        out.add(f"piece_set_{i:02d}.jpg")
        out.add(f"piece_truth_{i:02d}.jpg")
        out.add(f"piece_lib_{i:02d}.jpg")
        return out

    # draft の piece_truth_* も piece_set_* レ点と同一視
    checked_expanded: set[str] = set()
    for n in checked_names:
        checked_expanded |= _piece_name_aliases(n)
    checked_names = checked_expanded
    # 既定レ点: draft → piece_set 先頭3 → HUMAN 先頭3
    if not checked_names and set_files:
        checked_names = {p.name for p in set_files[:3]}
    if not checked_names and human_files:
        for p in human_files[:3]:
            checked_names |= _piece_name_aliases(p.name)
    if not checked_names and display_files:
        for p in display_files[:3]:
            checked_names |= _piece_name_aliases(p.name)

    def _is_checked(name: str) -> bool:
        return bool(_piece_name_aliases(name) & checked_names)

    def _piece_card_html(
        *,
        name: str,
        rel: str,
        where: str,
        checked: bool,
        slot_n: int = 0,
        lang_text: str = "",
        lang_key: str = "",
        can_remove: bool = False,
    ) -> str:
        esc = html_lib.escape(name)
        rel_esc = html_lib.escape(rel)
        chk = "checked" if checked else ""
        del_btn = ""
        if can_remove and slot_n > 0:
            del_btn = (
                f'<button type="button" class="btnPieceDel" data-slot="{slot_n}">'
                f"一覧から外す</button>"
            )
        # ダメなものは完全削除（ライブラリも現行も）
        purge_btn = (
            f'<button type="button" class="btnPiecePurge" data-name="{esc}">'
            f"この画像を削除</button>"
        )
        lang_key_use = (lang_key or name).strip()
        lang_esc_key = html_lib.escape(lang_key_use)
        lang_side = ""
        if lang_text or lang_key_use:
            lang_val = html_lib.escape(lang_text)
            lang_side = f"""
              <div class="piece-card-lang">
                <label class="piece-lang-lab"><b>一片言語化</b>（生成に同送・編集可）</label>
                <textarea class="pieceLangText" data-lang-name="{lang_esc_key}" rows="9"
                  placeholder="形状・縁・ねじれ・ほつれ等。空なら未作成。"
                >{lang_val}</textarea>
                <p class="piece-lang-actions">
                  <button type="button" class="btnPieceLangSave" data-lang-name="{lang_esc_key}">言語化を保存</button>
                  <span class="pieceLangStatus" data-lang-name="{lang_esc_key}"></span>
                </p>
              </div>"""
        card_cls = "piece-card hh-card"
        if checked:
            card_cls += " selected"
        if lang_side:
            card_cls += " has-lang"
        return f"""
            <div class="{card_cls}" data-piece-name="{esc}">
              <div class="piece-card-media">
                <img src="/file/{rel_esc}" class="hh-img piece-img" alt="" title="クリックで拡大"/>
                <label class="piece-pick-lab"><input type="checkbox" class="piecePick" value="{esc}" {chk}/> セットに含める</label>
                <div class="cand-meta"><code>{esc}</code> · {where}</div>
                <label class="piece-note-lab"><b>この1枚への要望</b></label>
                <textarea class="pieceNote" data-name="{esc}" rows="2" maxlength="400"
                  placeholder="ロック時は差分のみ。例: 縁のほつれだけ少し増やす"></textarea>
                <label class="piece-lock-lab">
                  <input type="checkbox" class="pieceLockCurrent" data-name="{esc}" checked/>
                  この写真で色・形をロック
                </label>
                <p class="piece-lock-hint">オフ＝種ベースで別形・色も可</p>
                <p class="piece-card-actions">
                  <button type="button" class="btnPieceRemake" data-name="{esc}">この1枚だけ作り直す</button>
                  {del_btn}
                  {purge_btn}
                </p>
                <div class="pieceRemakeStatus" data-name="{esc}"></div>
              </div>
              {lang_side}
            </div>"""

    cards: List[str] = []
    shown_slots: set[int] = set()
    shown_names: set[str] = set()

    def _slot_of(name: str) -> int:
        m = re.search(r"(\d+)", Path(str(name or "")).stem)
        return int(m.group(1)) if m else 0

    # 正本表示は piece_set 一本（HUMAN と二重に出さない）
    for p in set_files:
        slot_n = _slot_of(p.name)
        cards.append(
            _piece_card_html(
                name=p.name,
                rel=f"food_piece_set5_virt/piece_set/{p.name}",
                where="一片セット（本線）",
                checked=_is_checked(p.name),
                slot_n=slot_n,
                lang_text=_lang_text_for_piece_name(theme_new_dir, p.name),
                lang_key=p.name,
                can_remove=slot_n > 0,
            )
        )
        shown_names.add(p.name)
        if slot_n:
            shown_slots.add(slot_n)
    # piece_set が無いときだけ HUMAN／ライブラリを出す（初期投入用）
    if not set_files:
        for p in display_files:
            if "PIECE_TRUTH_HUMAN" in str(p).replace("\\", "/"):
                rel = f"01_refs/PIECE_TRUTH_HUMAN/{p.name}"
                where = "貼付・作業中"
                slot_n = _slot_of(p.name)
                can_remove = slot_n > 0
            else:
                rel = f"01_refs/PIECE_TRUTH_LIBRARY/{p.name}"
                where = "過去ライブラリ"
                slot_n = 0
                can_remove = False
            lang_key = ""
            if slot_n:
                lang_key = f"piece_set_{slot_n:02d}.jpg"
            cards.append(
                _piece_card_html(
                    name=p.name,
                    rel=rel,
                    where=where,
                    checked=_is_checked(p.name),
                    slot_n=slot_n,
                    lang_text=_lang_text_for_piece_name(theme_new_dir, p.name),
                    lang_key=lang_key or p.name,
                    can_remove=can_remove,
                )
            )
            shown_names.add(p.name)
    else:
        # セット未反映の新規貼付だけ追加表示（同じスロットは出さない）
        for p in human_files:
            slot_n = _slot_of(p.name)
            if slot_n and slot_n in shown_slots:
                continue
            if p.name in shown_names:
                continue
            cards.append(
                _piece_card_html(
                    name=p.name,
                    rel=f"01_refs/PIECE_TRUTH_HUMAN/{p.name}",
                    where="未反映の貼付",
                    checked=_is_checked(p.name),
                    slot_n=slot_n,
                    lang_text=_lang_text_for_piece_name(theme_new_dir, p.name),
                    lang_key=(
                        f"piece_set_{slot_n:02d}.jpg" if slot_n else p.name
                    ),
                    can_remove=slot_n > 0,
                )
            )
            shown_names.add(p.name)
            if slot_n:
                shown_slots.add(slot_n)
    grid = "\n".join(cards) or (
        "<p class='hint'>まだありません。下の「セット用」枠で1片写真を貼るか、元見本から生成したセットを使います。"
        "レ点した枚だけ画像生成に使います。</p>"
    )
    # 生成の元見本（種）— 精度の正は1枚（ラジオ）
    seed_cards: List[str] = []
    seed_kind_lab = ""
    try:
        from b5_piece_gate import list_piece_seed_display

        seed = list_piece_seed_display(theme_new_dir) if theme_new_dir.is_dir() else {}
        seed_kind_lab = html_lib.escape(str(seed.get("kindLabel") or seed.get("kind") or ""))
        items = list(seed.get("items") or [])
        active = ""
        if isinstance(draft, dict):
            active = str(draft.get("pieceSeedName") or "").strip()
        # draft が空でも piece_set_source.activeSeed を優先（左端固定にしない）
        if not active and theme_new_dir.is_dir():
            try:
                src_json = (
                    Path(theme_new_dir) / "food_piece_set5_virt" / "piece_set_source.json"
                )
                if src_json.is_file():
                    sobj = json.loads(src_json.read_text(encoding="utf-8-sig"))
                    if isinstance(sobj, dict):
                        active = str(sobj.get("activeSeed") or "").strip()
            except Exception:
                pass
        if not active and items:
            active = str(items[0].get("name") or "")
        for it in items:
            rel = str(it.get("rel") or "").strip()
            name = str(it.get("name") or "")
            name_esc = html_lib.escape(name)
            chk = "checked" if name == active else ""
            if rel:
                rel_esc = html_lib.escape(rel)
                img = f'<img src="/file/{rel_esc}" class="hh-img piece-img" alt=""/>'
            else:
                img = "<div class='hint'>（theme外のためサムネ不可）</div>"
            seed_cards.append(
                f"""
            <div class="hh-card{' selected' if chk else ''}">
              {img}
              <label class="piece-lab"><input type="radio" name="pieceSeedPick" value="{name_esc}" {chk}/> この1枚を元見本にする</label>
              <div class="cand-meta"><code>{name_esc}</code></div>
            </div>"""
            )
    except Exception:
        seed_cards = []
        seed_kind_lab = ""
    seed_grid = "\n".join(seed_cards) or (
        "<p class='hint'>まだありません。下の赤枠に実物写真を Ctrl+V（<b>1枚だけ</b>。新しい貼付で差し替え）。</p>"
    )
    deepen_top = ""
    deepen_seed = ""
    deepen_set = ""
    try:
        from b5_piece_deepen import panels_for_theme

        _dp = panels_for_theme(theme_new_dir)
        deepen_top = str(_dp.get("top") or "")
        deepen_seed = str(_dp.get("seed") or "")
        deepen_set = str(_dp.get("set") or "")
    except Exception:
        pass
    return f"""
<section class="wrap-gate piece-gate" id="pieceGate">
  <h2 style="font-size:15px;margin:16px 0 8px">① 一片見本はありますか</h2>
  {deepen_top}
  <p>
    <label class="piece-lab"><input type="radio" name="pieceHas" value="0" id="pieceHasOff" {off_checked}/> 一片見本なし</label>
    <label class="piece-lab"><input type="radio" name="pieceHas" value="1" id="pieceHasOn" {on_checked}/> 一片見本あり</label>
  </p>
  <div id="piecePanel" style="display:{disp}">
    <div class="prior" id="pieceGuide">
      <b>おすすめ（精度優先）</b><br/>
      ② 元見本は<b>登録は複数可・採用はラジオ1枚</b><br/>
      ③ 一片セットは<strong>1枚につき1カード</strong>（MAIN/SUBが読む正＝ここ）。二重表示しません。<br/>
      ※ 「一覧から外す」＝いったん外す。「この画像を削除」＝完全削除。<br/>
      作り直しの古い退避分は一覧に出しません。レ点最大{MAX_PIECE}（おすすめ3）
    </div>
    {deepen_seed}
    <h3 style="font-size:14px">② 生成の元見本（種・採用は1枚）</h3>
    <p class="hint">種別: <b>{seed_kind_lab or "未設定"}</b>。実物を貼ると登録に追加。
    <b>使う1枚だけラジオ</b>で選ぶ（選んだ時点で保存されます。クリックで拡大可）。
    作り直し時はこの種の色・表面に引っ張られやすいです。</p>
    <div id="pieceSeedPasteBox" class="paste-box" tabindex="0">元見本（種）：ここをクリックして Ctrl+V（登録追加）</div>
    <div>
      <input type="url" id="pieceSeedUrlIn" placeholder="https://… 元見本の直URL"/>
      <button type="button" id="btnPieceSeedUrl">URLから元見本を追加</button>
      <input type="file" id="pieceSeedFileIn" accept="image/*"/>
      <button type="button" id="btnPieceSeedFile">ファイルから追加</button>
    </div>
    <div class="grid" id="pieceSeedGrid" style="margin-top:8px">{seed_grid}</div>
    {deepen_set}
    <h3 style="font-size:14px;margin-top:14px">③ 一片セット（レ点したものだけ生成反映・おすすめ3／最大{MAX_PIECE}）</h3>
    <p class="hint">ここが<strong>唯一の正</strong>です。直すならこのカードを直す。言語化も同じカード。<br/>
    <b>作り直しは「この1枚だけ作り直す」</b>（他枚は変わりません）。ロックONなら色・形固定で差分のみ。<br/>
    下の大きな「見本確定して画像生成」は出品用MAIN/サブ用です。</p>
    <div id="piecePasteBox" class="paste-box" tabindex="0">セット用の1片：ここをクリックして Ctrl+V</div>
    <div>
      <input type="url" id="pieceUrlIn" placeholder="https://… 1片画像の直URL"/>
      <button type="button" id="btnPieceUrl">URLからセットへ追加</button>
      <input type="file" id="pieceFileIn" accept="image/*"/>
      <button type="button" id="btnPieceFile">ファイルからセットへ追加</button>
    </div>
    <div class="piece-row-grid" id="pieceHumanGrid">{grid}</div>
  </div>
</section>
"""


def build_pick_html(session: Dict[str, Any], candidates: List[Dict[str, Any]]) -> str:
    jan = session.get("jan") or ""
    need_hint = bool(session.get("needManualHint"))
    b3_n = int(session.get("competitorCount") or 0)
    manual_n = int(session.get("manualCount") or 0)
    theme_new_dir = Path(str(session.get("themeNewDir") or ""))
    prior = _prior_pick_state(theme_new_dir) if theme_new_dir.is_dir() else {"hasPrior": False}
    candidates = _sort_candidates_for_prior(candidates, prior)
    prior_identity = str(prior.get("identity") or "") if prior.get("hasPrior") else ""
    prior_sizzle = str(prior.get("sizzle") or "") if prior.get("hasPrior") else ""
    prior_wrap = str(prior.get("wrap") or "") if prior.get("hasPrior") else ""
    prefer_wrap = str(session.get("preferredWrapName") or "").strip()
    # 優先: 画面下書き > 直後の個包装コピペ > 前回確定（food_canons）
    # 旧: prefer は prior が空のときだけ → 右側を選び直しても左（前回）に戻る
    if prefer_wrap:
        prior_wrap = prefer_wrap
    draft = session.get("draftPicks") if isinstance(session.get("draftPicks"), dict) else {}
    draft_hand = str(draft.get("handHoldName") or "").strip()
    draft_single = str(draft.get("wrapSingleName") or "").strip()
    draft_remake_note = str(draft.get("remakeNote") or "").strip()[:400]
    draft_piece_names: List[str] = []
    if draft:
        if str(draft.get("identityName") or draft.get("primary") or "").strip():
            prior_identity = str(draft.get("identityName") or draft.get("primary") or "").strip()
        if "sizzleName" in draft:
            prior_sizzle = str(draft.get("sizzleName") or "").strip()
        if "wrapName" in draft:
            prior_wrap = str(draft.get("wrapName") or "").strip()
        if isinstance(draft.get("secondary"), list):
            prior_sec_draft = {str(x) for x in draft.get("secondary") if str(x).strip()}
        else:
            prior_sec_draft = None
        if isinstance(draft.get("pieceNames"), list):
            draft_piece_names = [
                str(x).strip() for x in draft.get("pieceNames") if str(x).strip()
            ]
    else:
        prior_sec_draft = None
    prior_sec = (
        prior_sec_draft
        if prior_sec_draft is not None
        else (set(prior.get("secondary") or []) if prior.get("hasPrior") else set())
    )
    prior_piece = set(draft_piece_names)
    if not prior_piece and theme_new_dir.is_dir():
        try:
            from b5_piece_gate import list_human_piece_files

            prior_piece = {p.name for p in list_human_piece_files(theme_new_dir)}
        except Exception:
            prior_piece = set()
        try:
            cm_piece = load_food_canons(theme_new_dir)
            for n in cm_piece.get("pieceNames") or []:
                if str(n).strip():
                    prior_piece.add(str(n).strip())
        except Exception:
            pass
    use_prior = bool(prior.get("hasPrior")) or bool(prefer_wrap) or bool(draft)
    is_main = str(session.get("context") or "").strip().lower() == "main"
    display_name = str(session.get("nameBase") or session.get("productName") or jan).strip() or jan
    parent_sku = str(session.get("parentSku") or "").strip()
    folder_line = (
        (f"親SKU {parent_sku} ／ " if parent_sku else "")
        + f"見本フォルダ JAN {jan}"
    )
    size_banner = master_size_banner(str(jan), parent_sku)

    if use_prior:
        prior_banner = f"""
        <div class="prior">
          <b>前回の選定（人間確定）を復元しています</b>
          <span style="color:#555;font-size:12px;"> — {prior.get("builtAt") or ""}</span><br/>
          IDENTITY主: <code>{prior_identity or "（なし）"}</code>
          ／ SIZZLE: <code>{prior_sizzle or "なし"}</code>
          ／ WRAP: <code>{prior_wrap or "なし"}</code>
          ／ 副: <code>{", ".join(sorted(prior_sec)) or "なし"}</code>
          ／ 一片: <code>{", ".join(sorted(prior_piece)) or "なし"}</code><br/>
          <span class="hint">下のカードでレ点が入っているものが前回の選択です。変えるときだけ選び直して確定。</span>
          <div class="prior-thumbs">
            {'<div><img src="/canon" alt="IDENTITY"/><div>IDENTITY確定画像</div></div>' if prior.get("hasCanonFile") and not prior.get("packageOnly") else ""}
            {'<div><img src="/sizzle" alt="SIZZLE"/><div>SIZZLE確定画像</div></div>' if prior.get("hasSizzleFile") else ""}
            {'<div><img src="/wrap" alt="WRAP"/><div>WRAP確定画像</div></div>' if prior.get("hasWrapFile") else ""}
          </div>
        </div>
        """
    else:
        prior_banner = ""

    # ①元IDENTITY言語化／②生成後盛り＝生成指示 — 二段表示
    panels = (
        resolve_identity_plating_panels(theme_new_dir)
        if theme_new_dir.is_dir()
        else {}
    )
    src_rel = str(panels.get("sourceRel") or "")
    gen_rel = str(panels.get("generatedRel") or "")
    src_name = str(panels.get("sourceName") or "")
    gen_name = str(panels.get("generatedName") or "")
    src_esc = html_lib.escape(str(panels.get("sourceText") or ""))
    gen_esc = html_lib.escape(str(panels.get("generatedText") or ""))
    src_img = (
        f'<div style="flex:0 0 auto">'
        f'<img src="/file/{html_lib.escape(src_rel)}" class="hh-img cand-img" alt="①元IDENTITY" '
        f'title="クリックで拡大" style="max-width:300px;border:2px solid #0a7;display:block"/>'
        f'<div class="cand-meta" style="margin-top:4px"><b>①盛りの元画像（IDENTITY正）</b><br/>'
        f'<code>{html_lib.escape(src_rel)}</code></div></div>'
        if src_rel
        else "<p class='hint'>①元IDENTITY画像がありません。候補で IDENTITY主を選んで確定してください。</p>"
    )
    gen_img = (
        f'<div style="flex:0 0 auto">'
        f'<img src="/file/{html_lib.escape(gen_rel)}" class="hh-img cand-img" alt="②生成後盛り" '
        f'title="クリックで拡大" style="max-width:300px;border:2px solid #c60;display:block"/>'
        f'<div class="cand-meta" style="margin-top:4px"><b>②生成後の盛り画像</b><br/>'
        f'<code>{html_lib.escape(gen_rel)}</code></div></div>'
        if gen_rel
        else "<p class='hint'>②生成後の盛り画像はまだありません（一片から作った盛り差し替え等）。</p>"
    )
    src_warn = ""
    if not str(panels.get("sourceText") or "").strip() and str(
        panels.get("generatedText") or ""
    ).strip():
        src_warn = (
            "<p class='hint' style='color:#a40'>いまの言語化は生成後画像由来の可能性が高いです。"
            "①は元画像を言語化して保存してください（本線の正は①）。</p>"
        )
    deepen_plate = ""
    try:
        from b5_piece_deepen import panels_for_theme as _deepen_panels_plate

        deepen_plate = str(_deepen_panels_plate(theme_new_dir).get("plate") or "")
    except Exception:
        deepen_plate = ""
    identity_plating_section = f"""
<section class="wrap-gate" id="identityPlatingLangGate" style="margin:16px 0;padding:12px;border:2px solid #2a6;border-radius:8px;background:#f3faf3">
  {deepen_plate}
  <h2 style="font-size:15px;margin:0 0 8px">IDENTITY盛り・二段言語化（①正本／②生成指示）</h2>
  <p class="hint" style="margin-top:0">
    <b>①</b>＝盛りの<strong>元画像</strong>の言語化（配置の正・MAIN/SUB同送）。<br/>
    <b>②</b>＝<strong>生成後</strong>の盛り画像の言語化＝<strong>生成時の指示</strong>（代用品・検収用。①の代わりにしない）。<br/>
    個品の色・縁・ほつれは一片側。ここは器・量・重なり・ばらけ・ねじれ散りだけ。
  </p>
  {src_warn}
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:16px;align-items:start">
    <div style="padding:10px;border:1px solid #0a7;border-radius:8px;background:#fff">
      <h3 style="font-size:13px;margin:0 0 8px;color:#0a7">① IDENTITYの言語化（元画像）</h3>
      <div style="display:flex;flex-wrap:wrap;gap:12px;align-items:flex-start">
        {src_img}
        <div style="flex:1;min-width:220px">
          <label style="font-size:12px;font-weight:bold">元画像の盛り言語化（本線の正・編集可）</label>
          <textarea id="identityPlatingLangTextSource" rows="12"
            placeholder="元画像の器・量感・重なり・配置・山の高さ・ばらけ幅・枚数目安・禁じ方"
            style="width:100%;box-sizing:border-box;font-size:12px;padding:8px;line-height:1.4;margin-top:4px">{src_esc}</textarea>
          <p style="margin:6px 0 0">
            <button type="button" id="btnIdentityPlatingLangSaveSource" style="font-weight:bold">①を保存（本線正）</button>
            <button type="button" id="btnIdentityPlatingLangReadSource" data-role="source" data-name="{html_lib.escape(src_name)}">元画像を機械言語化</button>
            <span id="identityPlatingLangStatusSource" style="font-size:11px;color:#555;margin-left:8px"></span>
          </p>
        </div>
      </div>
    </div>
    <div style="padding:10px;border:1px solid #c60;border-radius:8px;background:#fff">
      <h3 style="font-size:13px;margin:0 0 8px;color:#c60">② 生成後の盛り言語化（＝生成時の指示）</h3>
      <div style="display:flex;flex-wrap:wrap;gap:12px;align-items:flex-start">
        {gen_img}
        <div style="flex:1;min-width:220px">
          <label style="font-size:12px;font-weight:bold">生成指示／生成後の言語化（編集可）</label>
          <textarea id="identityPlatingLangTextGenerated" rows="12"
            placeholder="生成時に渡す指示。生成後画像の配置確認にも使う"
            style="width:100%;box-sizing:border-box;font-size:12px;padding:8px;line-height:1.4;margin-top:4px">{gen_esc}</textarea>
          <p style="margin:6px 0 0">
            <button type="button" id="btnIdentityPlatingLangSaveGenerated" style="font-weight:bold">②を保存（生成指示）</button>
            <button type="button" id="btnIdentityPlatingLangReadGenerated" data-role="generated" data-name="{html_lib.escape(gen_name)}">生成後画像を機械言語化</button>
            <span id="identityPlatingLangStatusGenerated" style="font-size:11px;color:#555;margin-left:8px"></span>
          </p>
        </div>
      </div>
    </div>
  </div>
</section>
"""

    cards = []
    for i, c in enumerate(candidates):
        name = str(c.get("name") or "")
        score = c.get("score")
        reason = str(c.get("reasonJa") or "")
        src = str(c.get("source") or "")
        src_lab = "【手動】" if src == "manual" or name.upper().startswith("MANUAL_") else ""
        sug = "（機械の第一候補→IDENTITY向き）" if c.get("suggestedPrimary") else ""
        if use_prior:
            id_checked = "checked" if name == prior_identity else ""
            sz_checked = "checked" if name and name == prior_sizzle else ""
            wrap_checked = "checked" if name and name == prior_wrap else ""
            sec_checked = "checked" if name in prior_sec else ""
            roles = []
            if name == prior_identity:
                roles.append("前回IDENTITY")
            if name == prior_sizzle:
                roles.append("前回SIZZLE")
            if name == prior_wrap:
                roles.append("前回WRAP")
            if name in prior_sec:
                roles.append("前回副")
            role_lab = (" <b style='color:#0a7'>" + "・".join(roles) + "</b>") if roles else ""
            card_cls = "cand-card selected" if roles else "cand-card"
        else:
            id_checked = "checked" if c.get("suggestedPrimary") else ""
            sz_checked = ""
            wrap_checked = ""
            sec_checked = ""
            role_lab = ""
            card_cls = "cand-card"
        cards.append(
            f"""
            <div class="{card_cls}">
              <img src="/file/01_refs/{name}" alt="{name}" class="cand-img"/>
              <div class="cand-meta">{src_lab}#{i} score={score} {sug}{role_lab}<br/>{reason}<br/><code>{name}</code></div>
              <label class="id-lab"><input type="radio" name="identity" value="{name}" {id_checked}/> IDENTITY主（開封後・色の正）</label>
              <label class="sz-lab"><input type="radio" name="sizzle" value="{name}" {sz_checked}/> SIZZLE主（炙り・とろけ等・任意）</label>
              <label class="wrap-lab"><input type="radio" name="wrap" value="{name}" {wrap_checked}/> WRAP（個包装・任意）</label>
              <label><input type="checkbox" class="sec" value="{name}" {sec_checked}/> IDENTITY副（盛り・形・最大{MAX_SECONDARY}）</label>
            </div>
            """
        )
    cards_html = "\n".join(cards) or "<p>候補がありません（下のコピペで追加するか、正本のみで進めてください）</p>"
    if need_hint:
        zero_banner = f"""
        <div class="warn">
          <b>出品ファイルに競合候補がありません（B-③採用0）。</b><br/>
          Webで<strong>個品（開封・中身・盛り）</strong>の商品画像を探し、下の枠に<strong>Ctrl+V でコピペ</strong>してください（最大{MANUAL_MAX}）。
          袋だけの写真は避ける。<br/>
          見つからなければ「競合・個品なしで生成（正本のみ）」で進めます（個品見本は使いません）。
        </div>
        """
    else:
        zero_banner = f"""
        <div class="info">競合候補 {b3_n} 枚 ／ 手動 {manual_n} 枚。必要なら下で画像を<strong>追加コピペ</strong>できます（任意・最大{MANUAL_MAX}）。</div>
        """
    truth_checked = "checked" if use_prior and prior_identity in ("__truth__", "PACKAGE_TRUTH") else ""
    pkg_checked = "checked" if use_prior and prior.get("packageOnly") else ""
    sizzle_none_checked = (
        "checked"
        if (not use_prior) or (use_prior and not prior_sizzle)
        else ""
    )
    wrap_none_checked = (
        "checked"
        if (not use_prior) or (use_prior and not prior_wrap)
        else ""
    )
    wrap_section = _wrap_inline_section_html(
        session,
        theme_new_dir=theme_new_dir,
        prior_wrap=prior_wrap,
        draft_hand=draft_hand,
        draft_single=draft_single,
        draft_remake_note=draft_remake_note,
    )
    piece_section = _piece_inline_section_html(
        session,
        theme_new_dir=theme_new_dir,
        draft_piece_names=draft_piece_names or list(prior_piece),
    )
    if is_main:
        action_block = """
<p class="hint" style="background:#eef6ff;border:1px solid #8af;padding:10px;border-radius:8px;">
  <b>B-④ MAIN用</b> — この画面からセールストークには進みません。<br/>
  確認だけなら <b>「変更なしで閉じる」</b> → B-④ の <b>「次へ」</b>。<br/>
  レ点を変えたら <b>「選び直したらここで保存」</b> → タブを閉じて B-④ の <b>「次へ」</b>。
</p>
<p>
  <button type="button" id="btnMainClose">変更なしで閉じる</button>
  <button type="button" id="btnMainSave" style="font-weight:bold">選び直したらここで保存</button>
  <span id="status"></span>
</p>
"""
        sub_buttons_js = """
document.getElementById('btnMainClose').onclick=()=>{
  document.getElementById('status').textContent='B-④ の「次へ」へ進んでください';
  try { window.close(); } catch(e) {}
};
document.getElementById('btnMainSave').onclick=async()=>{
  clearErr();
  const p=picks();
  const verr=validatePicks(p);
  if(verr){ showErr(verr); document.getElementById('status').textContent=''; return; }
  document.getElementById('status').textContent='保存中…';
  const res=await fetch('/api/confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});
  const j=await res.json();
  if(!j.ok){ showErr(j.message||'保存に失敗しました'); document.getElementById('status').textContent=''; return; }
  document.getElementById('status').textContent='保存しました。このタブを閉じて B-④ の「次へ」へ';
  try { window.close(); } catch(e) {}
};
"""
    else:
        action_block = """
<p class="hint" style="background:#fff8e1;border:1px solid #e6b800;padding:10px;border-radius:8px;">
  <b>注意</b> — 下のボタンは <b>IDENTITY確定→出品用MAIN/サブ画像</b> 用です。<br/>
  一片見本の作り直しは、上の各カードの <b>「この1枚だけ作り直す」</b> を使ってください（他の一片は一緒に作り直されません）。
</p>
<p>
  <button id="btnRun" style="font-weight:bold">見本確定して画像生成</button>
  <button id="btnConfirmSkip">見本確定のみ（画像の生成不要）</button>
  <button id="btnRefresh">候補を再取得</button>
  <button id="btnReview">既存画像のレビューへ（生成しない）</button>
  <span id="status"></span>
</p>
<p class="hint">個包装ありのときも、この画面で見本登録・生成・採用まで済ませてから保存します。</p>
"""
        sub_buttons_js = """
document.getElementById('btnRun').onclick=async()=>{
  clearErr();
  const p=picks();
  const verr=validatePicks(p);
  if(verr){ showErr(verr); document.getElementById('status').textContent=''; return; }
  if(p.packageOnly && NEED_HINT){
    if(!confirm('競合・個品見本なしで正本のみ生成します。よろしいですか？')) return;
  }
  document.getElementById('status').textContent='見本確定中…';
  let res=await fetch('/api/confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});
  let j=await res.json();
  if(!j.ok){ showErr(j.message||'確定に失敗しました'); document.getElementById('status').textContent=''; return; }
  document.getElementById('status').textContent='画像生成中（数分かかることがあります）…';
  res=await fetch('/api/generate',{method:'POST'});
  j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
  if(j.ok) setTimeout(()=>location.reload(), 600);
};
document.getElementById('btnConfirmSkip').onclick=async()=>{
  clearErr();
  const p=picks();
  const verr=validatePicks(p);
  if(verr){ showErr(verr); document.getElementById('status').textContent=''; return; }
  document.getElementById('status').textContent='見本確定（生成スキップ）…';
  let res=await fetch('/api/confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});
  let j=await res.json();
  if(!j.ok){ showErr(j.message||'確定に失敗しました'); document.getElementById('status').textContent=''; return; }
  res=await fetch('/api/to-review',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({requireImages:true})});
  j=await res.json();
  if(!j.ok){ showErr(j.message||'レビューへ進めません（既存画像を確認）'); document.getElementById('status').textContent=''; return; }
  document.getElementById('status').textContent=j.message||'レビューへ';
  setTimeout(()=>location.reload(), 400);
};
document.getElementById('btnRefresh').onclick=async()=>{
  document.getElementById('status').textContent='候補再取得中…';
  const res=await fetch('/api/refresh-candidates',{method:'POST'});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||'';
  if(j.ok) location.reload();
};
document.getElementById('btnReview').onclick=async()=>{
  const res=await fetch('/api/to-review',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({requireImages:true})});
  const j=await res.json();
  if(!j.ok){ showErr(j.message||'既存画像がありません'); document.getElementById('status').textContent=''; return; }
  location.reload();
};
"""
    try:
        from b5_piece_deepen import deepen_css as _deepen_css
        from b5_piece_deepen import deepen_js as _deepen_js

        deepen_css_block = _deepen_css()
        deepen_js_block = _deepen_js()
    except Exception:
        deepen_css_block = ""
        deepen_js_block = ""
    return f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>食材見本ゲート — {display_name}</title>
<style>
body{{font-family:sans-serif;margin:16px;background:#f6f3ee;color:#222}}
{deepen_css_block}

h1{{font-size:18px}}
.hint{{font-size:13px;color:#555;line-height:1.45}}
.warn{{background:#fff3cd;border:2px solid #c90;padding:12px;border-radius:8px;margin:12px 0;font-size:13px;line-height:1.5}}
.prior{{background:#e8f5e9;border:2px solid #2e7d32;padding:12px;border-radius:8px;margin:12px 0;font-size:13px;line-height:1.5}}
.prior-thumbs{{display:flex;flex-wrap:wrap;gap:12px;margin-top:8px}}
.prior-thumbs img{{height:96px;border:2px solid #2e7d32;border-radius:4px;cursor:zoom-in}}
.prior-thumbs div{{font-size:11px;color:#333;text-align:center}}
.err{{display:none;background:#fdecea;border:2px solid #c62828;color:#b71c1c;padding:12px;border-radius:8px;margin:12px 0;font-size:14px;font-weight:bold;line-height:1.45}}
.err.on{{display:block}}
.info{{background:#eef6ff;border:1px solid #8af;padding:10px;border-radius:8px;margin:12px 0;font-size:13px}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:12px;margin:12px 0}}
.cand-card{{background:#fff;border:1px solid #ddd;padding:8px;border-radius:6px}}
.cand-card.selected{{border:2px solid #2e7d32;box-shadow:0 0 0 2px rgba(46,125,50,.15)}}
.cand-img{{width:100%;min-height:120px;height:auto;display:block;cursor:zoom-in;border:1px solid #eee;background:#f0ece6}}
.cand-meta{{font-size:11px;color:#555;margin:6px 0;word-break:break-all}}
label{{display:block;font-size:12px;margin:3px 0}}
.id-lab{{color:#0d47a1;font-weight:600}}
.sz-lab{{color:#6a1b9a}}
.wrap-lab{{color:#00695c}}
.piece-lab{{color:#b71c1c;font-weight:600}}
button{{padding:10px 14px;margin-right:8px;cursor:pointer}}
#status{{margin-left:8px}}
.truth{{max-width:180px;border:2px solid #888}}
.paste-box{{border:2px dashed #c45;background:#fff;min-height:72px;padding:12px;border-radius:8px;margin:8px 0;outline:none}}
.paste-box:focus{{border-color:#a00;background:#fff8f5}}
.manual-row{{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px}}
.manual-row img{{height:72px;border:1px solid #ccc}}
#lightbox{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:9;align-items:center;justify-content:center}}
#lightbox.on{{display:flex}}
#lightbox img{{max-width:96vw;max-height:96vh}}
input[type=url],input[type=file]{{width:min(480px,100%);margin:4px 0}}
.legend{{font-size:12px;background:#fff;border:1px solid #ccc;padding:8px;border-radius:6px}}
.bag-pick{{background:#fff;border:2px solid #1565c0;border-radius:8px;padding:12px;margin:12px 0}}
.bag-grid .bag-card{{display:block;background:#fafafa;border:2px solid #ddd;border-radius:6px;padding:8px;cursor:pointer}}
.bag-grid .bag-card.selected{{border-color:#1565c0;box-shadow:0 0 0 2px rgba(21,101,192,.2)}}
.bag-grid .bag-img{{width:100%;height:140px;object-fit:contain;background:#f0ece6;border:1px solid #eee;display:block}}
.bag-grid .cand-meta{{font-size:11px;color:#555;margin-top:6px;word-break:break-all}}
.wrap-gate{{margin:16px 0}}
.hh-card{{display:block;background:#fff;border:2px solid #ddd;border-radius:8px;padding:8px;cursor:pointer}}
.hh-card.selected{{border-color:#00695c;box-shadow:0 0 0 2px rgba(0,105,92,.2)}}
.hh-img{{width:100%;height:160px;object-fit:contain;background:#f0ece6;border:1px solid #eee;display:block;cursor:zoom-in}}
.hh-ring-lab{{font-size:11px;color:#666;margin:6px 0 2px}}
.hh-ring-preview{{height:140px;background:#f5f2ec;border:1px solid #c9a227}}
/* 一片＋言語化: 左メディア／右言語化。狭い縦長カードをやめる */
.piece-row-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(520px,1fr));gap:10px;margin:8px 0}}
.piece-card{{display:block;cursor:default;padding:10px}}
.piece-card.has-lang{{display:grid;grid-template-columns:168px minmax(0,1fr);gap:10px;align-items:stretch}}
.piece-card-media .piece-img{{height:132px;width:100%}}
.piece-pick-lab,.piece-note-lab{{margin:4px 0 2px;font-size:11px}}
.piece-lock-lab{{margin:6px 0 0;font-size:11px;font-weight:600;color:#0d47a1}}
.piece-lock-hint{{margin:2px 0 0;font-size:10px;color:#666}}
.piece-card-media .pieceNote{{width:100%;box-sizing:border-box;font-size:11px;padding:4px;line-height:1.3;resize:vertical}}
.piece-card-actions{{margin:4px 0 0;display:flex;flex-wrap:wrap;gap:4px}}
.piece-card-actions button{{padding:6px 8px;margin:0;font-size:11px}}
.pieceRemakeStatus{{font-size:11px;color:#555;min-height:1em;margin-top:2px}}
.piece-card-lang{{display:flex;flex-direction:column;min-width:0;min-height:0}}
.piece-lang-lab{{font-size:11px;color:#333;margin:0 0 4px}}
.piece-card-lang .pieceLangText{{flex:1;width:100%;min-height:148px;box-sizing:border-box;font-size:12px;padding:6px 8px;line-height:1.4;resize:vertical;font-family:ui-monospace,Consolas,monospace}}
.piece-lang-actions{{margin:4px 0 0}}
.piece-lang-actions button{{padding:6px 10px;margin:0;font-size:12px}}
.pieceLangStatus{{font-size:11px;color:#555;margin-left:6px}}
@media (max-width:640px){{
  .piece-row-grid{{grid-template-columns:1fr}}
  .piece-card.has-lang{{grid-template-columns:1fr}}
}}
</style></head><body>
<h1>食材見本ゲート — {display_name}</h1>
{size_banner}
<p class="hint">{folder_line}<br/>
<strong>IDENTITY</strong>（必須）＝開封後の中身。色・厚み・1片の「何の商品か」。<br/>
<strong>SIZZLE</strong>（任意）＝炙り・とろけ等の食欲状態。<br/>
<strong>WRAP</strong>（任意）＝個包装の見た目。無いときは画像に個包装を描かない。<br/>
<strong>一片見本</strong>（任意）＝開封後1片の実物正。IDENTITY（盛り）とは別。無いときは一片ロックを使わない。<br/>
同一写真を IDENTITY と SIZZLE の両方にしてもよい。副は IDENTITY 用・最大{MAX_SECONDARY}枚。</p>
<div class="legend">T03→SIZZLE優先 ／ T06→IDENTITY ／ T1・T2・料理系→両方 ／ WRAPは個包装を見せるときだけ ／ 一片見本は薄片ロック用 ／ needPackage→袋＋IDENTITY参照</div>
{prior_banner}
{identity_plating_section}
{zero_banner}
<div id="errBox" class="err" role="alert"></div>
{_bag_pick_section_html(session)}
<div class="grid">{cards_html}</div>
{wrap_section}
{piece_section}
<section>
  <h2 style="font-size:15px">手動で画像を追加（IDENTITY／SIZZLE用・最大{MANUAL_MAX}）</h2>
  <p class="hint">ここは<strong>候補カード（IDENTITY等）用</strong>です。一片の元見本・セットは上の「一片見本」ブロック内へ貼ってください。予備: URL・ファイル。</p>
  <div id="pasteBox" class="paste-box" tabindex="0">IDENTITY等：ここをクリックしてから Ctrl+V</div>
  <div>
    <input type="url" id="urlIn" placeholder="https://… 画像の直URL（予備）"/>
    <button type="button" id="btnUrl">URLから追加</button>
  </div>
  <div>
    <input type="file" id="fileIn" accept="image/*"/>
    <button type="button" id="btnFile">ファイルから追加</button>
  </div>
  <div class="manual-row" id="manualPrev"></div>
</section>
<p>
  <label class="id-lab"><input type="radio" name="identity" value="__truth__" id="radTruth" {truth_checked}/> IDENTITY＝正本パッケージのみ</label>
  <label class="id-lab"><input type="radio" name="identity" value="__package_only__" id="radPkgOnly" {pkg_checked}/> 競合・個品なし（正本のみ／IDENTITY・SIZZLE・WRAPとも使わない）</label>
  <label class="sz-lab"><input type="radio" name="sizzle" value="__none__" {sizzle_none_checked}/> SIZZLEなし（任意・既定）</label>
  <label class="wrap-lab"><input type="radio" name="wrap" value="__none__" {wrap_none_checked}/> WRAPなし（任意・既定）</label>
</p>
{action_block}
<div id="lightbox"><img id="lbImg"/></div>
<script>
const NEED_HINT = {str(need_hint).lower()};
const ERR_NO_IDENTITY = 'IDENTITY（主見本）が未選択です。副見本や SIZZLE／WRAP だけでは確定できません。開封後の個品を IDENTITY 主にするか、「正本のみ／個品なし」を選んでください。';
function showErr(msg){{
  const el=document.getElementById('errBox');
  el.textContent=msg;
  el.classList.add('on');
  el.scrollIntoView({{behavior:'smooth',block:'center'}});
}}
function clearErr(){{ document.getElementById('errBox').classList.remove('on'); }}
document.getElementById('lightbox').onclick=()=>document.getElementById('lightbox').classList.remove('on');
function picks(){{
  const identity=(document.querySelector('input[name=identity]:checked')||{{}}).value||'';
  let sizzle=(document.querySelector('input[name=sizzle]:checked')||{{}}).value||'__none__';
  if(sizzle==='__none__') sizzle='';
  let wrap=(document.querySelector('input[name=wrap]:checked')||{{}}).value||'__none__';
  if(wrap==='__none__') wrap='';
  const wrapHas=!!(document.querySelector('input[name=wrapHas]:checked')&&document.querySelector('input[name=wrapHas]:checked').value==='1');
  const handHold=(document.querySelector('input[name=handHoldPick]:checked')||{{}}).value||'';
  const wrapSingle=(document.querySelector('input[name=wrapSinglePick]:checked')||{{}}).value||'';
  // 単体おすすめを WRAP 未選択時の見本に使える
  if(wrapHas && !wrap && wrapSingle) wrap=wrapSingle;
  if(!wrapHas) wrap='';
  const pieceHasEl=document.querySelector('input[name=pieceHas]:checked');
  const pieceHas=!!(pieceHasEl&&pieceHasEl.value==='1');
  const pieceNames=[...document.querySelectorAll('#pieceHumanGrid .piecePick:checked')].map(x=>x.value)
    .filter((v,i,a)=>v&&a.indexOf(v)===i).slice(0,5);
  const pieceSeed=(document.querySelector('input[name=pieceSeedPick]:checked')||{{}}).value||'';
  const secondary=[...document.querySelectorAll('.sec:checked')].map(x=>x.value)
    .filter(v=>v&&v!==identity).slice(0,{MAX_SECONDARY});
  return {{
    primary: identity,
    identityName: identity,
    sizzleName: sizzle,
    wrapName: wrap,
    wrapHas: wrapHas,
    handHoldName: wrapHas ? handHold : '',
    wrapSingleName: wrapHas ? wrapSingle : '',
    pieceHas: pieceHas,
    pieceNames: pieceHas ? pieceNames : [],
    pieceSeedName: pieceHas ? pieceSeed : '',
    secondary,
    packageOnly: identity==='__package_only__'
  }};
}}
function validatePicks(p){{
  const secN=[...document.querySelectorAll('.sec:checked')].length;
  const hasSz=[...document.querySelectorAll('input[name=sizzle]:checked')].some(x=>x.value&&x.value!=='__none__');
  const hasWrap=[...document.querySelectorAll('input[name=wrap]:checked')].some(x=>x.value&&x.value!=='__none__');
  if(!p.identityName){{
    if(secN>0 || hasSz || hasWrap || p.wrapHas || p.pieceHas) return ERR_NO_IDENTITY;
    return 'IDENTITY（主見本）を1つ選んでください（または正本のみ／個品なし）。';
  }}
  if(p.wrapHas){{
    if(!p.wrapName && !p.wrapSingleName){{
      return '個包装ありです。見本をコピペするか、単体候補／WRAPレ点を選んでください。';
    }}
  }}
  if(p.pieceHas){{
    const n=(p.pieceNames||[]).length;
    const humanN=document.querySelectorAll('#pieceHumanGrid .piecePick').length;
    if(n<1 && humanN<1){{
      return '一片見本ありです。貼付するか、「③ セットに含める」にレ点を付けてください。';
    }}
  }}
  return '';
}}
function syncWrapPanel(){{
  const on=document.getElementById('wrapHasOn');
  const panel=document.getElementById('wrapPanel');
  if(!panel) return;
  panel.style.display=(on&&on.checked)?'block':'none';
}}
document.querySelectorAll('input[name=wrapHas]').forEach(el=>el.addEventListener('change', syncWrapPanel));
syncWrapPanel();
function syncPiecePanel(){{
  const on=document.getElementById('pieceHasOn');
  const panel=document.getElementById('piecePanel');
  if(!panel) return;
  panel.style.display=(on&&on.checked)?'block':'none';
}}
document.querySelectorAll('input[name=pieceHas]').forEach(el=>el.addEventListener('change', syncPiecePanel));
syncPiecePanel();
function remakeNoteText(){{
  const el=document.getElementById('wrapRemakeNote');
  return el ? String(el.value||'').trim().slice(0,400) : '';
}}
let _draftSaveTimer=null;
async function saveDraftPicks(){{
  const p=picks();
  p.remakeNote=remakeNoteText();
  try {{
    await fetch('/api/draft-picks',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify(p)}});
  }} catch(e) {{}}
}}
function scheduleDraftPicks(){{
  if(_draftSaveTimer) clearTimeout(_draftSaveTimer);
  _draftSaveTimer=setTimeout(()=>{{ saveDraftPicks(); }}, 150);
}}
document.querySelectorAll(
  'input[name=identity],input[name=sizzle],input[name=wrap],input[name=wrapHas],input[name=pieceHas],input[name=handHoldPick],input[name=wrapSinglePick],input.sec,input.piecePick'
).forEach(el=>el.addEventListener('change', scheduleDraftPicks));
// 元見本は遅延なしで即保存（リロードで左端に戻るのを防ぐ）
document.querySelectorAll('input[name=pieceSeedPick]').forEach(el=>{{
  el.addEventListener('change', async()=>{{
    document.getElementById('status').textContent='元見本を保存中…';
    await saveDraftPicks();
    document.getElementById('status').textContent='元見本を切り替えました';
  }});
}});
const remakeNoteEl=document.getElementById('wrapRemakeNote');
if(remakeNoteEl){{
  remakeNoteEl.addEventListener('change', scheduleDraftPicks);
  remakeNoteEl.addEventListener('blur', scheduleDraftPicks);
}}
async function wrapPasteSave(dataUrl){{
  await saveDraftPicks();
  document.getElementById('status').textContent='個包装見本を保存中…';
  const res=await fetch('/api/manual-paste',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{dataUrl, preferWrap:true}})}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
  if(j.ok) location.reload();
}}
const wrapPaste=document.getElementById('wrapPasteBox');
if(wrapPaste){{
  wrapPaste.addEventListener('paste', async (e)=>{{
    const items=e.clipboardData&&e.clipboardData.items;
    if(!items) return;
    for(const it of items){{
      if(!String(it.type||'').startsWith('image/')) continue;
      e.preventDefault();
      const file=it.getAsFile();
      if(!file) continue;
      const reader=new FileReader();
      reader.onload=async()=>wrapPasteSave(reader.result);
      reader.readAsDataURL(file);
      break;
    }}
  }});
}}
const btnWrapUrl=document.getElementById('btnWrapUrl');
if(btnWrapUrl){{
  btnWrapUrl.onclick=async()=>{{
    const url=(document.getElementById('wrapUrlIn').value||'').trim();
    if(!url){{ alert('URLを入力'); return; }}
    await saveDraftPicks();
    document.getElementById('status').textContent='個包装URL取得中…';
    const res=await fetch('/api/manual-url',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{url, preferWrap:true}})}});
    const j=await res.json();
    document.getElementById('status').textContent=j.message||'';
    if(j.ok) location.reload();
  }};
}}
const btnWrapFile=document.getElementById('btnWrapFile');
if(btnWrapFile){{
  btnWrapFile.onclick=async()=>{{
    const f=document.getElementById('wrapFileIn').files[0];
    if(!f){{ alert('ファイルを選択'); return; }}
    const reader=new FileReader();
    reader.onload=async()=>wrapPasteSave(reader.result);
    reader.readAsDataURL(f);
  }};
}}
async function piecePasteSave(dataUrl){{
  await saveDraftPicks();
  document.getElementById('status').textContent='セット用の1片を保存中…';
  const res=await fetch('/api/piece-paste',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{dataUrl}})}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
  if(j.ok) location.reload();
}}
async function pieceSeedPasteSave(dataUrl){{
  await saveDraftPicks();
  document.getElementById('status').textContent='元見本（種）を差し替え中…';
  const res=await fetch('/api/piece-seed-paste',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{dataUrl}})}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
  if(j.ok) location.reload();
}}
function bindPasteBox(el, saver){{
  if(!el) return;
  el.addEventListener('paste', async (e)=>{{
    const items=e.clipboardData&&e.clipboardData.items;
    if(!items) return;
    for(const it of items){{
      if(!String(it.type||'').startsWith('image/')) continue;
      e.preventDefault();
      const file=it.getAsFile();
      if(!file) continue;
      const reader=new FileReader();
      reader.onload=async()=>saver(reader.result);
      reader.readAsDataURL(file);
      break;
    }}
  }});
}}
bindPasteBox(document.getElementById('piecePasteBox'), piecePasteSave);
bindPasteBox(document.getElementById('pieceSeedPasteBox'), pieceSeedPasteSave);
const btnPieceUrl=document.getElementById('btnPieceUrl');
if(btnPieceUrl){{
  btnPieceUrl.onclick=async()=>{{
    const url=(document.getElementById('pieceUrlIn').value||'').trim();
    if(!url){{ alert('URLを入力'); return; }}
    await saveDraftPicks();
    document.getElementById('status').textContent='セット用URL取得中…';
    const res=await fetch('/api/piece-url',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{url}})}});
    const j=await res.json();
    document.getElementById('status').textContent=j.message||'';
    if(j.ok) location.reload();
  }};
}}
const btnPieceFile=document.getElementById('btnPieceFile');
if(btnPieceFile){{
  btnPieceFile.onclick=async()=>{{
    const f=document.getElementById('pieceFileIn').files[0];
    if(!f){{ alert('ファイルを選択'); return; }}
    const reader=new FileReader();
    reader.onload=async()=>piecePasteSave(reader.result);
    reader.readAsDataURL(f);
  }};
}}
const btnPieceSeedUrl=document.getElementById('btnPieceSeedUrl');
if(btnPieceSeedUrl){{
  btnPieceSeedUrl.onclick=async()=>{{
    const url=(document.getElementById('pieceSeedUrlIn').value||'').trim();
    if(!url){{ alert('URLを入力'); return; }}
    await saveDraftPicks();
    document.getElementById('status').textContent='元見本URL取得中…';
    const res=await fetch('/api/piece-seed-url',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{url}})}});
    const j=await res.json();
    document.getElementById('status').textContent=j.message||'';
    if(j.ok) location.reload();
  }};
}}
const btnPieceSeedFile=document.getElementById('btnPieceSeedFile');
if(btnPieceSeedFile){{
  btnPieceSeedFile.onclick=async()=>{{
    const f=document.getElementById('pieceSeedFileIn').files[0];
    if(!f){{ alert('ファイルを選択'); return; }}
    const reader=new FileReader();
    reader.onload=async()=>pieceSeedPasteSave(reader.result);
    reader.readAsDataURL(f);
  }};
}}
document.querySelectorAll('.btnPieceDel').forEach(btn=>{{
  btn.onclick=async()=>{{
    const slot=Number(btn.getAttribute('data-slot')||0);
    if(!slot) return;
    if(!confirm('セット用の1片 slot='+slot+' を一覧から外しますか？（ライブラリには残ります）')) return;
    await saveDraftPicks();
    const res=await fetch('/api/piece-delete',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{slot}})}});
    const j=await res.json();
    document.getElementById('status').textContent=j.message||'';
    if(j.ok) location.reload();
  }};
}});
document.querySelectorAll('.btnPiecePurge').forEach(btn=>{{
  btn.onclick=async()=>{{
    const name=btn.getAttribute('data-name')||'';
    if(!name) return;
    if(!confirm('「'+name+'」を完全削除します。復元できません。よろしいですか？')) return;
    await saveDraftPicks();
    const res=await fetch('/api/piece-purge',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{name}})}});
    const j=await res.json();
    document.getElementById('status').textContent=j.message||'';
    if(j.ok) location.reload();
    else showErr(j.message||'削除失敗');
  }};
}});
document.querySelectorAll('.btnPieceRemake').forEach(btn=>{{
  btn.onclick=async()=>{{
    const name=btn.getAttribute('data-name')||'';
    if(!name) return;
    const noteEl=document.querySelector('textarea.pieceNote[data-name="'+CSS.escape(name)+'"]');
    const lockEl=document.querySelector('input.pieceLockCurrent[data-name="'+CSS.escape(name)+'"]');
    const statusEl=document.querySelector('.pieceRemakeStatus[data-name="'+CSS.escape(name)+'"]');
    const note=(noteEl&&noteEl.value||'').trim();
    const lockCurrent=!(lockEl&&lockEl.checked===false);
    const seed=(document.querySelector('input[name=pieceSeedPick]:checked')||{{}}).value||'';
    const modeLab=lockCurrent?'現行ロック（色・形固定）':'種ベース（色・形も可変）';
    if(!confirm('「'+name+'」だけ作り直します（'+modeLab+'）。他の一片は変わりません。よろしいですか？')) return;
    await saveDraftPicks();
    btn.disabled=true;
    if(statusEl) statusEl.textContent='この1枚だけ生成中（'+modeLab+'）…';
    document.getElementById('status').textContent='一片 '+name+' だけ作り直し中…';
    try{{
      const res=await fetch('/api/piece-remake',{{method:'POST',headers:{{'Content-Type':'application/json'}},
        body:JSON.stringify({{name, note, seedName:seed, lockCurrent}})}});
      const j=await res.json();
      if(!j.ok){{
        if(statusEl) statusEl.textContent='';
        showErr(j.message||'作り直し失敗');
        document.getElementById('status').textContent='';
        btn.disabled=false;
        return;
      }}
      if(statusEl) statusEl.textContent=j.message||'完了';
      document.getElementById('status').textContent=j.message||'完了';
      location.reload();
    }}catch(e){{
      if(statusEl) statusEl.textContent='';
      showErr(String(e));
      btn.disabled=false;
    }}
  }};
}});
document.querySelectorAll('.btnPieceLangSave').forEach(btn=>{{
  btn.onclick=async()=>{{
    const name=btn.getAttribute('data-lang-name')||'';
    if(!name) return;
    const ta=document.querySelector('textarea.pieceLangText[data-lang-name="'+CSS.escape(name)+'"]');
    const st=document.querySelector('.pieceLangStatus[data-lang-name="'+CSS.escape(name)+'"]');
    const text=(ta&&ta.value||'').trim();
    if(!text){{ showErr('言語化が空です'); return; }}
    btn.disabled=true;
    if(st) st.textContent='保存中…';
    try{{
      const res=await fetch('/api/piece-lang-save',{{method:'POST',headers:{{'Content-Type':'application/json'}},
        body:JSON.stringify({{name, text}})}});
      const j=await res.json();
      if(!j.ok){{ showErr(j.message||'言語化の保存に失敗'); if(st) st.textContent=''; btn.disabled=false; return; }}
      if(st) st.textContent='保存しました';
      document.getElementById('status').textContent=j.message||'一片言語化を保存しました';
    }}catch(e){{ showErr(String(e)); if(st) st.textContent=''; }}
    btn.disabled=false;
  }};
}});
async function saveIdentityPlatingLang(role){{
  const isGen=role==='generated';
  const ta=document.getElementById(isGen?'identityPlatingLangTextGenerated':'identityPlatingLangTextSource');
  const st=document.getElementById(isGen?'identityPlatingLangStatusGenerated':'identityPlatingLangStatusSource');
  const btn=document.getElementById(isGen?'btnIdentityPlatingLangSaveGenerated':'btnIdentityPlatingLangSaveSource');
  const text=(ta&&ta.value||'').trim();
  if(!text){{ showErr((isGen?'②生成後盛り言語化':'①IDENTITY元画像言語化')+'が空です'); return; }}
  if(btn) btn.disabled=true;
  if(st) st.textContent='保存中…';
  try{{
    const res=await fetch('/api/identity-plating-lang-save',{{method:'POST',headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{text, role: isGen?'generated':'source'}})}});
    const j=await res.json();
    if(!j.ok){{ showErr(j.message||'保存失敗'); if(st) st.textContent=''; if(btn) btn.disabled=false; return; }}
    if(st) st.textContent='保存しました';
    document.getElementById('status').textContent=j.message||'保存しました';
  }}catch(e){{ showErr(String(e)); if(st) st.textContent=''; }}
  if(btn) btn.disabled=false;
}}
async function readIdentityPlatingLang(role, name){{
  const isGen=role==='generated';
  const ta=document.getElementById(isGen?'identityPlatingLangTextGenerated':'identityPlatingLangTextSource');
  const st=document.getElementById(isGen?'identityPlatingLangStatusGenerated':'identityPlatingLangStatusSource');
  if(!name){{ showErr('対象画像がありません'); return; }}
  if(st) st.textContent='機械言語化中…';
  try{{
    const res=await fetch('/api/identity-plating-lang-read',{{method:'POST',headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{role: isGen?'generated':'source', name}})}});
    const j=await res.json();
    if(!j.ok){{ showErr(j.message||'機械言語化に失敗'); if(st) st.textContent=''; return; }}
    if(ta) ta.value=j.text||'';
    if(st) st.textContent='欄に反映しました（未保存）。内容を確認して保存してください';
    document.getElementById('status').textContent=j.message||'機械言語化完了（未保存）';
  }}catch(e){{ showErr(String(e)); if(st) st.textContent=''; }}
}}
const btnIdPlateSaveSrc=document.getElementById('btnIdentityPlatingLangSaveSource');
if(btnIdPlateSaveSrc) btnIdPlateSaveSrc.onclick=()=>saveIdentityPlatingLang('source');
const btnIdPlateSaveGen=document.getElementById('btnIdentityPlatingLangSaveGenerated');
if(btnIdPlateSaveGen) btnIdPlateSaveGen.onclick=()=>saveIdentityPlatingLang('generated');
const btnIdPlateReadSrc=document.getElementById('btnIdentityPlatingLangReadSource');
if(btnIdPlateReadSrc) btnIdPlateReadSrc.onclick=()=>readIdentityPlatingLang('source', btnIdPlateReadSrc.getAttribute('data-name')||'');
const btnIdPlateReadGen=document.getElementById('btnIdentityPlatingLangReadGenerated');
if(btnIdPlateReadGen) btnIdPlateReadGen.onclick=()=>readIdentityPlatingLang('generated', btnIdPlateReadGen.getAttribute('data-name')||'');
async function runWrapGen(){{
  clearErr();
  const p=picks();
  const seed=p.wrapName||p.wrapSingleName||'';
  await saveDraftPicks();
  document.getElementById('wrapGenStatus').textContent='生成中（未作成分）…';
  const res=await fetch('/api/wrap-hand-hold/generate',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{force:false, wrapSeedName:seed, bundle:true}})}});
  const j=await res.json();
  if(!j.ok){{ showErr(j.message||'生成失敗'); document.getElementById('wrapGenStatus').textContent=''; return; }}
  document.getElementById('wrapGenStatus').textContent=j.message||'完了';
  location.reload();
}}
async function runWrapRemake(kind){{
  clearErr();
  const cls=kind==='hand'?'remakeHand':'remakeSingle';
  const names=[...document.querySelectorAll('input.'+cls+':checked')].map(x=>x.value);
  if(!names.length){{ showErr((kind==='hand'?'手持ち':'単体')+'の「作り直し対象」にレ点を付けてください'); return; }}
  if(!confirm((kind==='hand'?'手持ち':'単体')+'のレ点 '+names.length+' 枚だけ作り直します。よろしいですか？')) return;
  const p=picks();
  const seed=p.wrapName||p.wrapSingleName||'';
  const userNote=remakeNoteText();
  await saveDraftPicks();
  document.getElementById('wrapGenStatus').textContent='レ点分だけ作り直し中…';
  const res=await fetch('/api/wrap-hand-hold/remake',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{kind, names, wrapSeedName:seed, userNote}})}});
  const j=await res.json();
  if(!j.ok){{ showErr(j.message||'作り直し失敗'); document.getElementById('wrapGenStatus').textContent=''; return; }}
  document.getElementById('wrapGenStatus').textContent=j.message||'完了';
  location.reload();
}}
const btnWrapGen=document.getElementById('btnWrapGen');
if(btnWrapGen) btnWrapGen.onclick=()=>runWrapGen();
const btnRemakeHand=document.getElementById('btnRemakeHand');
if(btnRemakeHand) btnRemakeHand.onclick=()=>runWrapRemake('hand');
const btnRemakeSingle=document.getElementById('btnRemakeSingle');
if(btnRemakeSingle) btnRemakeSingle.onclick=()=>runWrapRemake('single');
function bindLightbox(root){{
  (root||document).querySelectorAll('.cand-img,.truth,.prior-thumbs img,.hh-img,.piece-img,#manualPrev img').forEach(img=>{{
    if(img.dataset.lbBound) return;
    img.dataset.lbBound='1';
    img.style.cursor='zoom-in';
    img.onclick=()=>{{ document.getElementById('lbImg').src=img.src; document.getElementById('lightbox').classList.add('on'); }};
  }});
}}
async function refreshManual(){{
  const res=await fetch('/api/manual-list');
  const j=await res.json();
  const box=document.getElementById('manualPrev');
  box.innerHTML='';
  (j.items||[]).forEach(it=>{{
    const w=document.createElement('div');
    w.innerHTML=`<img src="/file/01_refs/${{it.name}}" title="${{it.name}}（クリックで拡大）" class="manual-thumb"/><br/><button type="button" data-slot="${{it.slot}}">削除</button>`;
    w.querySelector('button').onclick=async()=>{{
      await fetch('/api/manual-delete',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{slot:it.slot}})}});
      location.reload();
    }};
    box.appendChild(w);
  }});
  bindLightbox(box);
}}
refreshManual();
bindLightbox(document);
const pasteBox=document.getElementById('pasteBox');
pasteBox.addEventListener('paste', async (e)=>{{
  const items=e.clipboardData&&e.clipboardData.items;
  if(!items) return;
  for(const it of items){{
    if(!String(it.type||'').startsWith('image/')) continue;
    e.preventDefault();
    const file=it.getAsFile();
    if(!file) continue;
    const reader=new FileReader();
    reader.onload=async()=>{{
      document.getElementById('status').textContent='コピペ画像を保存中…';
      const res=await fetch('/api/manual-paste',{{method:'POST',headers:{{'Content-Type':'application/json'}},
        body:JSON.stringify({{dataUrl:reader.result}})}});
      const j=await res.json();
      document.getElementById('status').textContent=j.message||JSON.stringify(j);
      if(j.ok) location.reload();
    }};
    reader.readAsDataURL(file);
    break;
  }}
}});
document.getElementById('btnUrl').onclick=async()=>{{
  const url=document.getElementById('urlIn').value.trim();
  if(!url){{ alert('URLを入力'); return; }}
  document.getElementById('status').textContent='URL取得中…';
  const res=await fetch('/api/manual-url',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{url}})}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||'';
  if(j.ok) location.reload();
}};
document.getElementById('btnFile').onclick=async()=>{{
  const f=document.getElementById('fileIn').files[0];
  if(!f){{ alert('ファイルを選択'); return; }}
  const reader=new FileReader();
  reader.onload=async()=>{{
    const res=await fetch('/api/manual-paste',{{method:'POST',headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{dataUrl:reader.result}})}});
    const j=await res.json();
    document.getElementById('status').textContent=j.message||'';
    if(j.ok) location.reload();
  }};
  reader.readAsDataURL(f);
}};
{sub_buttons_js}
{_bag_pick_js()}
{deepen_js_block}
</script>
</body></html>
"""


def make_handler(theme_new_dir: Path, script_dir: Path, *, listen_port: int = 8790):
    theme_new_dir = Path(theme_new_dir).resolve()
    listen_port = int(listen_port or 8790)

    def _ensure_mix_manifest(session: Dict[str, Any]) -> Path:
        jan = str(session.get("jan") or "")
        man = build_mix_manifest(
            jan=jan,
            theme_new_dir=theme_new_dir,
            # 見本ゲート成果のみ。旧競合 compose は混ぜない
            compose_dir=None,
            provider="openai",
            product_name=str(session.get("productName") or ""),
        )
        path = unique_mix_manifest_path(jan, theme_new_dir, None, out_root=script_dir / "_out")
        write_mix_manifest(man, path)
        return path

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt: str, *args) -> None:  # noqa: A003
            LOG.info("%s - %s", self.address_string(), fmt % args)

        def _send(self, code: int, body: bytes, content_type: str) -> None:
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _json(self, code: int, obj: Dict[str, Any]) -> None:
            self._send(
                code,
                json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                "application/json",
            )

        def do_GET(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            session = _load_session(theme_new_dir)
            if parsed.path.startswith("/bag-cand/"):
                try:
                    idx = int(parsed.path.split("/")[-1])
                except Exception:
                    self._send(404, b"bad index", "text/plain")
                    return
                cands = list(session.get("bagCandidates") or [])
                if not cands:
                    cands = [
                        {"rel": c.get("rel"), "path": c.get("path")}
                        for c in _bag_candidates_for_session(session)
                    ]
                    _save_session(theme_new_dir, session)
                if idx < 0 or idx >= len(cands):
                    self._send(404, b"no cand", "text/plain")
                    return
                tpath = Path(str(cands[idx].get("path") or ""))
                if not tpath.is_file():
                    self._send(404, b"missing", "text/plain")
                    return
                ctype = "image/png" if tpath.suffix.lower() == ".png" else "image/jpeg"
                self._send(200, tpath.read_bytes(), ctype)
                return
            if str(session.get("mode") or "") == "mix":
                if parsed.path in ("/", "/index.html"):
                    html = build_mix_confirm_html(session)
                    _save_session(theme_new_dir, session)
                    self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
                    return
                if parsed.path == "/mix/bag":
                    bag = _mix_parent_bag(session)
                    tpath = Path(str(bag or ""))
                    if not tpath.is_file():
                        self._send(404, b"missing", "text/plain")
                        return
                    ctype = "image/png" if tpath.suffix.lower() == ".png" else "image/jpeg"
                    self._send(200, tpath.read_bytes(), ctype)
                    return
                if parsed.path.startswith("/part/"):
                    bits = [x for x in parsed.path.split("/") if x]
                    # part / {i} / bag|canon|sizzle|wrap|sec / {j}
                    try:
                        idx = int(bits[1])
                    except Exception:
                        self._send(404, b"bad part", "text/plain")
                        return
                    parts = _mix_parts_from_session(session)
                    if idx < 0 or idx >= len(parts):
                        self._send(404, b"no part", "text/plain")
                        return
                    p = parts[idx]
                    kind = bits[2] if len(bits) > 2 else ""
                    target = None
                    if kind == "bag":
                        target = p.get("bag")
                    elif kind == "canon":
                        target = p.get("identity")
                    elif kind == "sizzle":
                        target = p.get("sizzle")
                    elif kind == "wrap":
                        target = p.get("wrap")
                    elif kind == "sec":
                        try:
                            sj = int(bits[3])
                        except Exception:
                            sj = -1
                        secs = list(p.get("secondary") or [])
                        target = secs[sj] if 0 <= sj < len(secs) else None
                    tpath = Path(str(target or ""))
                    if not tpath.is_file():
                        self._send(404, b"missing", "text/plain")
                        return
                    ctype = "image/png" if tpath.suffix.lower() == ".png" else "image/jpeg"
                    self._send(200, tpath.read_bytes(), ctype)
                    return
                self._send(404, b"not found", "text/plain")
                return
            if parsed.path == "/pick":
                session["phase"] = "pick"
                _save_session(theme_new_dir, session)
                self.send_response(302)
                self.send_header("Location", "/")
                self.end_headers()
                return
            if parsed.path in ("/", "/index.html"):
                phase = session.get("phase") or "pick"
                if phase == "review":
                    man_path = _ensure_mix_manifest(session)
                    man = json.loads(man_path.read_text(encoding="utf-8"))
                    html = build_mix_html(man)
                    gen_msg = (session.get("gen") or {}).get("message", "")
                    inject = (
                        "<p class='hint'><a href='/pick'>← 食材見本の選び直し</a> ／ "
                        f"phase=review ／ {gen_msg}</p>"
                    )
                    html = html.replace(
                        '<div class="toolbar">', '<div class="toolbar">' + inject, 1
                    )
                    self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
                    return
                if phase == "wrap_hand_hold":
                    html = build_wrap_hand_hold_html(session)
                    self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
                    return
                cand = load_candidates_meta(theme_new_dir).get("candidates") or []
                # 手動コピペ画像を候補に必ず載せる（再オープン時の見落とし防止）
                try:
                    cand = _sync_manual_into_candidates(theme_new_dir).get("candidates") or cand
                except Exception as ex:
                    LOG.warning("manual sync on pick: %s", ex)
                session = _refresh_session_counts(theme_new_dir, session)
                html = build_pick_html(session, cand)
                _save_session(theme_new_dir, session)
                self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
                return
            if parsed.path == "/truth":
                t = Path(str(session.get("truth") or ""))
                if not t.is_file():
                    self._send(404, b"no truth", "text/plain")
                    return
                ctype = "image/png" if t.suffix.lower() == ".png" else "image/jpeg"
                self._send(200, t.read_bytes(), ctype)
                return
            if parsed.path == "/sizzle":
                ex = existing_sizzle_path(theme_new_dir)
                if not ex:
                    self._send(404, b"no sizzle", "text/plain")
                    return
                ctype = "image/png" if Path(ex).suffix.lower() == ".png" else "image/jpeg"
                self._send(200, Path(ex).read_bytes(), ctype)
                return
            if parsed.path == "/wrap":
                ex = existing_wrap_path(theme_new_dir)
                if not ex:
                    self._send(404, b"no wrap", "text/plain")
                    return
                ctype = "image/png" if Path(ex).suffix.lower() == ".png" else "image/jpeg"
                self._send(200, Path(ex).read_bytes(), ctype)
                return
            if parsed.path.startswith("/wrap-hand-hold/"):
                from b4_wrap_hand_hold import out_dir_for

                name = unquote(parsed.path[len("/wrap-hand-hold/") :]).replace("\\", "/").lstrip("/")
                if not name or "/" in name or ".." in name:
                    self._send(404, b"bad name", "text/plain")
                    return
                target = (out_dir_for(theme_new_dir) / name).resolve()
                root_s = str(out_dir_for(theme_new_dir).resolve())
                if not str(target).startswith(root_s):
                    self._send(404, b"not found", "text/plain")
                    return
                if not target.is_file():
                    self._send(404, b"missing", "text/plain")
                    return
                ctype = "image/png" if target.suffix.lower() == ".png" else "image/jpeg"
                self._send(200, target.read_bytes(), ctype)
                return
            if parsed.path.startswith("/wrap-hand-hold-ring/"):
                from b4_wrap_bl_gold_stamp import render_hand_hold_ring_zoom_jpeg
                from b4_wrap_hand_hold import out_dir_for

                name = unquote(parsed.path[len("/wrap-hand-hold-ring/") :]).replace("\\", "/").lstrip("/")
                if not name or "/" in name or ".." in name or not name.startswith("hand_hold_"):
                    self._send(404, b"bad name", "text/plain")
                    return
                target = (out_dir_for(theme_new_dir) / name).resolve()
                root_s = str(out_dir_for(theme_new_dir).resolve())
                if not str(target).startswith(root_s) or not target.is_file():
                    self._send(404, b"missing", "text/plain")
                    return
                try:
                    jpeg = render_hand_hold_ring_zoom_jpeg(target)
                except Exception as e:
                    LOG.exception("wrap-hand-hold-ring")
                    self._send(500, str(e).encode("utf-8", errors="replace"), "text/plain")
                    return
                self._send(200, jpeg, "image/jpeg")
                return
            if parsed.path.startswith("/file/"):
                rel = unquote(parsed.path[len("/file/") :]).replace("\\", "/").lstrip("/")
                target = (theme_new_dir / rel).resolve()
                root_s = str(theme_new_dir.resolve())
                if not (
                    str(target) == root_s
                    or str(target).startswith(root_s + "\\")
                    or str(target).startswith(root_s + "/")
                ):
                    self._send(404, b"not found", "text/plain")
                    return
                if not target.is_file():
                    self._send(404, b"missing", "text/plain")
                    return
                suf = target.suffix.lower()
                ctype = "image/png" if suf == ".png" else "image/jpeg"
                self._send(200, target.read_bytes(), ctype)
                return
            if parsed.path == "/canon":
                ex = existing_canon_path(theme_new_dir)
                if not ex:
                    self._send(404, b"no canon", "text/plain")
                    return
                self._send(200, ex.read_bytes(), "image/jpeg")
                return
            if parsed.path.startswith("/img/"):
                cid = unquote(parsed.path[len("/img/") :])
                man_path = _ensure_mix_manifest(session)
                man = json.loads(man_path.read_text(encoding="utf-8"))
                by_id = {str(c.get("id")): c for c in (man.get("cards") or [])}
                card = by_id.get(cid)
                if not card:
                    self._send(404, b"not found", "text/plain")
                    return
                target = Path(str(card.get("path") or ""))
                if not target.is_file():
                    self._send(404, b"not found", "text/plain")
                    return
                self._send(200, target.read_bytes(), "image/jpeg")
                return
            if parsed.path == "/api/manual-list":
                meta = load_manual_meta(theme_new_dir)
                items = []
                for p in list_manual_paths(theme_new_dir):
                    m = None
                    for it in meta.get("items") or []:
                        if str(it.get("name") or "") == p.name:
                            m = it
                            break
                    slot = int((m or {}).get("slot") or 0)
                    if not slot:
                        try:
                            slot = int(p.stem.split("_")[-1])
                        except Exception:
                            slot = 0
                    items.append({"name": p.name, "slot": slot, "source": (m or {}).get("source")})
                self._json(200, {"ok": True, "items": items, "max": MANUAL_MAX})
                return
            self._send(404, b"not found", "text/plain")

        def do_POST(self) -> None:  # noqa: N802
            length = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(length) if length else b"{}"
            try:
                payload = json.loads(raw.decode("utf-8") or "{}")
            except Exception:
                payload = {}
            parsed = urlparse(self.path)
            session = _load_session(theme_new_dir)
            jan = str(session.get("jan") or "")

            if parsed.path == "/api/set-bag":
                try:
                    result = apply_bag_pick(
                        theme_new_dir,
                        session,
                        rel=str(payload.get("rel") or ""),
                    )
                    self._json(200, result)
                except ValueError as e:
                    self._json(400, {"ok": False, "message": str(e)})
                except Exception as e:
                    LOG.exception("set-bag")
                    self._json(500, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/open-part-gate":
                try:
                    idx = int(payload.get("partIndex"))
                except Exception:
                    self._json(400, {"ok": False, "message": "partIndex が不正です"})
                    return
                parts = list(session.get("parts") or [])
                if idx < 0 or idx >= len(parts) or not isinstance(parts[idx], dict):
                    self._json(400, {"ok": False, "message": "構成がありません"})
                    return
                part = parts[idx]
                jan = "".join(
                    ch for ch in str(part.get("jan") or "") if ch.isdigit()
                ) or str(part.get("jan") or "").strip()
                if not jan:
                    self._json(400, {"ok": False, "message": "構成に JAN がありません"})
                    return
                part_port = int(listen_port) + 10 + int(idx)
                tn = SCRIPT_DIR / "_out" / f"theme_new_run_{jan}"
                try:
                    from sub_image_food_canon_launch import kill_food_canon_on_ports

                    kill_food_canon_on_ports([part_port])
                except Exception as e:
                    LOG.warning("open-part kill: %s", e)
                cmd = [
                    sys.executable,
                    str(SCRIPT_DIR / "sub_image_food_canon_gate.py"),
                    "--jan",
                    jan,
                    "--theme-new-dir",
                    str(tn),
                    "--port",
                    str(part_port),
                    "--force-pick",
                    "--context",
                    "main",
                ]
                nb = str(part.get("nameBase") or "").strip()
                psku = str(part.get("parentSku") or "").strip()
                if nb:
                    cmd.extend(["--name-base", nb])
                if psku:
                    cmd.extend(["--parent-sku", psku])
                kwargs: Dict[str, Any] = {}
                if sys.platform == "win32":
                    kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP  # type: ignore[attr-defined]
                subprocess.Popen(cmd, cwd=str(SCRIPT_DIR), **kwargs)
                url = f"http://127.0.0.1:{part_port}/"
                self._json(
                    200,
                    {
                        "ok": True,
                        "url": url,
                        "message": f"単品見本を開きました（{nb or jan}）。WRAP を選んで「選び直したらここで保存」→ この組み合わせ画面を再読込。",
                    },
                )
                return

            if parsed.path == "/api/manual-paste":
                try:
                    result = save_manual_data_url(
                        theme_new_dir, str(payload.get("dataUrl") or "")
                    )
                    _sync_manual_into_candidates(theme_new_dir)
                    session = _refresh_session_counts(theme_new_dir, session)
                    if bool(payload.get("preferWrap")):
                        wname = str(result.get("name") or "")
                        session["preferredWrapName"] = wname
                        draft = dict(session.get("draftPicks") or {})
                        draft["wrapName"] = wname
                        draft["wrapHas"] = True
                        session["draftPicks"] = draft
                        _save_session(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"手動画像を追加しました（{result.get('name')}）",
                            **result,
                            "manualCount": session.get("manualCount"),
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/manual-url":
                try:
                    result = save_manual_url(theme_new_dir, str(payload.get("url") or ""))
                    _sync_manual_into_candidates(theme_new_dir)
                    session = _refresh_session_counts(theme_new_dir, session)
                    if bool(payload.get("preferWrap")):
                        wname = str(result.get("name") or "")
                        session["preferredWrapName"] = wname
                        draft = dict(session.get("draftPicks") or {})
                        draft["wrapName"] = wname
                        draft["wrapHas"] = True
                        session["draftPicks"] = draft
                        _save_session(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"URLから追加しました（{result.get('name')}）",
                            **result,
                            "manualCount": session.get("manualCount"),
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/manual-delete":
                try:
                    slot = int(payload.get("slot") or 0)
                    if slot < 1:
                        self._json(400, {"ok": False, "message": "slot が不正です"})
                        return
                    result = delete_manual_slot(theme_new_dir, slot)
                    _sync_manual_into_candidates(theme_new_dir)
                    session = _refresh_session_counts(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"手動画像 slot={slot} を削除",
                            **result,
                            "manualCount": session.get("manualCount"),
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-paste":
                try:
                    from b5_piece_gate import save_piece_data_url

                    result = save_piece_data_url(
                        theme_new_dir, str(payload.get("dataUrl") or "")
                    )
                    draft = dict(session.get("draftPicks") or {})
                    draft["pieceHas"] = True
                    names = [
                        str(x)
                        for x in (draft.get("pieceNames") or [])
                        if str(x).strip()
                    ]
                    nm = str(result.get("name") or "")
                    if nm and nm not in names:
                        names.append(nm)
                    draft["pieceNames"] = names[:5]
                    session["draftPicks"] = draft
                    _save_session(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"セット用の1片を追加しました（{nm}）",
                            **result,
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-seed-paste":
                try:
                    from b5_piece_gate import save_piece_seed_data_url

                    result = save_piece_seed_data_url(
                        theme_new_dir, str(payload.get("dataUrl") or "")
                    )
                    draft = dict(session.get("draftPicks") or {})
                    draft["pieceHas"] = True
                    draft["pieceSeedName"] = str(result.get("name") or "")
                    session["draftPicks"] = draft
                    _save_session(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"元見本（種）を1枚に差し替えました（{result.get('name')}）",
                            **result,
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-url":
                try:
                    from b5_piece_gate import save_piece_url

                    result = save_piece_url(theme_new_dir, str(payload.get("url") or ""))
                    draft = dict(session.get("draftPicks") or {})
                    draft["pieceHas"] = True
                    names = [
                        str(x)
                        for x in (draft.get("pieceNames") or [])
                        if str(x).strip()
                    ]
                    nm = str(result.get("name") or "")
                    if nm and nm not in names:
                        names.append(nm)
                    draft["pieceNames"] = names[:5]
                    session["draftPicks"] = draft
                    _save_session(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"URLからセット用1片を追加しました（{nm}）",
                            **result,
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-seed-url":
                try:
                    from b5_piece_gate import save_piece_seed_url

                    result = save_piece_seed_url(
                        theme_new_dir, str(payload.get("url") or "")
                    )
                    draft = dict(session.get("draftPicks") or {})
                    draft["pieceHas"] = True
                    draft["pieceSeedName"] = str(result.get("name") or "")
                    session["draftPicks"] = draft
                    _save_session(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"URLから元見本を差し替えました（{result.get('name')}）",
                            **result,
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-delete":
                try:
                    from b5_piece_gate import delete_piece_slot

                    slot = int(payload.get("slot") or 0)
                    if slot < 1:
                        self._json(400, {"ok": False, "message": "slot が不正です"})
                        return
                    result = delete_piece_slot(theme_new_dir, slot)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"一片見本 slot={slot} を一覧から外しました",
                            **result,
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-purge":
                try:
                    from b5_piece_gate import purge_piece_name

                    name = str(payload.get("name") or "").strip()
                    if not name:
                        self._json(400, {"ok": False, "message": "削除名が空です"})
                        return
                    result = purge_piece_name(theme_new_dir, name)
                    draft = session.get("draftPicks")
                    if isinstance(draft, dict):
                        names = [
                            str(x)
                            for x in (draft.get("pieceNames") or [])
                            if str(x).strip() and str(x).strip() != name
                        ]
                        draft["pieceNames"] = names[:5]
                        _save_session(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"「{name}」を完全削除しました",
                            **result,
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-lang-save":
                try:
                    name = str(payload.get("name") or "").strip()
                    text = str(payload.get("text") or "").strip()
                    result = save_piece_lang_text(
                        theme_new_dir, name=name, text=text
                    )
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"一片言語化を保存しました（{result.get('name')}）",
                            **result,
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/identity-plating-lang-save":
                try:
                    text = str(payload.get("text") or "").strip()
                    role = str(payload.get("role") or "source").strip() or "source"
                    result = save_identity_plating_text(
                        theme_new_dir, text=text, role=role
                    )
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": str(
                                result.get("message")
                                or "IDENTITY盛り言語化を保存しました"
                            ),
                            **result,
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/identity-plating-lang-read":
                try:
                    from _tmp_food_piece_plate_identity_lang_virt import (
                        read_identity_plating_lang,
                    )

                    role = str(payload.get("role") or "source").strip().lower() or "source"
                    name = Path(str(payload.get("name") or "").strip()).name
                    panels = resolve_identity_plating_panels(theme_new_dir)
                    if not name:
                        if role in ("generated", "gen", "result"):
                            name = str(panels.get("generatedName") or "")
                        else:
                            name = str(panels.get("sourceName") or "")
                    if not name:
                        raise ValueError("対象画像名がありません")
                    img = theme_new_dir / "01_refs" / name
                    if not img.is_file():
                        raise ValueError(f"画像がありません: {name}")
                    lang = read_identity_plating_lang(img)
                    if not lang.get("ok"):
                        raise ValueError(
                            str(lang.get("error") or "機械言語化に失敗しました")
                        )
                    text = str(lang.get("promptBlock") or "").strip()
                    if not text:
                        raise ValueError("言語化結果が空です")
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"{name} を機械言語化しました（未保存）",
                            "text": text,
                            "name": name,
                            "role": (
                                "generated"
                                if role in ("generated", "gen", "result")
                                else "source"
                            ),
                            "fields": lang.get("fields") or {},
                        },
                    )
                except Exception as e:
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-deepen":
                try:
                    from b5_piece_deepen import (
                        adopt_after_to_piece_set,
                        deepen_ui_enabled,
                        save_deepen,
                    )

                    if not deepen_ui_enabled():
                        self._json(
                            400,
                            {
                                "ok": False,
                                "message": "深堀UIはオフです（B5_PIECE_DEEPEN_UI=0）",
                            },
                        )
                        return
                    if payload.get("clear"):
                        obj = save_deepen(theme_new_dir, clear=True)
                        self._json(
                            200,
                            {
                                "ok": True,
                                "message": "深堀を閉じました",
                                "deepen": obj,
                            },
                        )
                        return
                    if payload.get("adopt"):
                        meta = adopt_after_to_piece_set(theme_new_dir)
                        draft = session.get("draftPicks")
                        if not isinstance(draft, dict):
                            draft = {}
                            session["draftPicks"] = draft
                        draft["pieceHas"] = True
                        draft["pieceNames"] = list(meta.get("adopted") or [])[:5]
                        _save_session(theme_new_dir, session)
                        self._json(
                            200,
                            {
                                "ok": True,
                                "message": (
                                    f"下段を一片セットへ採用しました（{meta.get('n') or 0}枚）。"
                                ),
                                "adopt": meta,
                            },
                        )
                        return
                    obj = save_deepen(
                        theme_new_dir,
                        tag=str(payload.get("tag") or ""),
                        note=str(payload.get("note") or ""),
                        active=(
                            bool(payload.get("active"))
                            if "active" in payload
                            else None
                        ),
                    )
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": "深堀を更新しました（参考・前回を自動構成）",
                            "deepen": obj,
                        },
                    )
                except Exception as e:
                    LOG.exception("piece-deepen failed")
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/piece-remake":
                # 一片カード1枚だけ作り直し（他枚・MAIN生成は触らない）
                try:
                    from b5_piece_gate import remake_one_piece

                    name = str(payload.get("name") or "").strip()
                    note = str(payload.get("note") or "").strip()[:400]
                    seed_name = str(
                        payload.get("seedName") or payload.get("pieceSeedName") or ""
                    ).strip()
                    lock_raw = payload.get("lockCurrent", payload.get("lock_current", True))
                    if isinstance(lock_raw, bool):
                        lock_current = lock_raw
                    else:
                        lock_current = str(lock_raw or "").strip().lower() not in (
                            "0",
                            "false",
                            "no",
                            "off",
                        )
                    if not name:
                        self._json(400, {"ok": False, "message": "対象の一片名が空です"})
                        return
                    result = remake_one_piece(
                        theme_new_dir,
                        name=name,
                        note=note,
                        seed_name=seed_name,
                        lock_current=lock_current,
                    )
                    # 深堀中: 結果は下段へ。本線 piece_set は前回生成のまま残す
                    deepen_rec: Dict[str, Any] = {}
                    try:
                        from b5_piece_deepen import load_deepen, record_remake_to_after

                        if load_deepen(theme_new_dir).get("active"):
                            deepen_rec = record_remake_to_after(
                                theme_new_dir,
                                set_name=str(result.get("setName") or ""),
                                truth_path=str(result.get("path") or ""),
                                slot=int(result.get("slot") or 0),
                            )
                    except Exception as e:
                        LOG.warning("deepen after record failed: %s", e)
                        deepen_rec = {"ok": False, "error": str(e)}
                    # 作り直した新名を draft のレ点に足す（旧名・新名どちらも残す）
                    draft = session.get("draftPicks")
                    if not isinstance(draft, dict):
                        draft = {}
                        session["draftPicks"] = draft
                    names = [
                        str(x)
                        for x in (draft.get("pieceNames") or [])
                        if str(x).strip()
                    ]
                    new_name = str(result.get("name") or "").strip()
                    if new_name and new_name not in names:
                        names.append(new_name)
                    draft["pieceNames"] = names[:5]
                    if seed_name:
                        draft["pieceSeedName"] = seed_name
                    _save_session(theme_new_dir, session)
                    msg = (
                        f"「{name}」だけ作り直しました"
                        + (f"（→{new_name}）" if new_name and new_name != name else "")
                        + "。他の一片は変更していません。"
                    )
                    if deepen_rec.get("ok"):
                        msg += (
                            " 深堀中のため結果は下段へ載せ、本線の一片セットは前回のまま残しました。"
                        )
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": msg,
                            "deepenAfter": deepen_rec,
                            **result,
                        },
                    )
                except Exception as e:
                    LOG.exception("piece-remake failed")
                    self._json(400, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/draft-picks":
                # 作り直し／生成の reload 前に画面の選択を一時保存（確定ではない）
                draft = {
                    "primary": str(payload.get("primary") or payload.get("identityName") or "").strip(),
                    "identityName": str(
                        payload.get("identityName") or payload.get("primary") or ""
                    ).strip(),
                    "sizzleName": str(payload.get("sizzleName") or "").strip(),
                    "wrapName": str(payload.get("wrapName") or "").strip(),
                    "wrapHas": bool(payload.get("wrapHas")),
                    "handHoldName": str(payload.get("handHoldName") or "").strip(),
                    "wrapSingleName": str(payload.get("wrapSingleName") or "").strip(),
                    "pieceHas": bool(payload.get("pieceHas")),
                    "pieceNames": [
                        str(x)
                        for x in (payload.get("pieceNames") or [])
                        if str(x).strip()
                    ][:5],
                    "pieceSeedName": str(payload.get("pieceSeedName") or "").strip(),
                    "remakeNote": str(payload.get("remakeNote") or "").strip()[:400],
                    "secondary": [
                        str(x) for x in (payload.get("secondary") or []) if str(x).strip()
                    ],
                    "packageOnly": bool(payload.get("packageOnly")),
                }
                session["draftPicks"] = draft
                if draft.get("wrapName"):
                    session["preferredWrapName"] = str(draft["wrapName"])
                # 元見本ラジオは下書きだけでなく activeSeed にも即反映（リロードで左端へ戻さない）
                seed_saved = str(draft.get("pieceSeedName") or "").strip()
                if draft.get("pieceHas") and seed_saved:
                    try:
                        from b5_piece_gate import activate_piece_seed

                        activate_piece_seed(theme_new_dir, seed_saved)
                    except Exception as e:
                        LOG.warning("draft activate_piece_seed failed: %s", e)
                _save_session(theme_new_dir, session)
                self._json(200, {"ok": True, "message": "下書き選択を保存"})
                return

            if parsed.path == "/api/confirm":
                primary = str(
                    payload.get("identityName")
                    or payload.get("primary")
                    or ""
                ).strip()
                sizzle_name = str(payload.get("sizzleName") or "").strip()
                wrap_name = str(payload.get("wrapName") or "").strip()
                wrap_has = bool(payload.get("wrapHas"))
                hand_hold_name = str(payload.get("handHoldName") or "").strip()
                wrap_single_name = str(payload.get("wrapSingleName") or "").strip()
                piece_has = bool(payload.get("pieceHas"))
                piece_names = [
                    str(x) for x in (payload.get("pieceNames") or []) if str(x).strip()
                ][:5]
                piece_seed_name = str(payload.get("pieceSeedName") or "").strip()
                secondary = [
                    str(x) for x in (payload.get("secondary") or []) if str(x).strip()
                ]
                package_only = bool(payload.get("packageOnly")) or primary in (
                    "__package_only__",
                    "PACKAGE_ONLY",
                )
                if not wrap_has or package_only:
                    wrap_name = ""
                    hand_hold_name = ""
                    wrap_single_name = ""
                if package_only:
                    piece_has = False
                    piece_names = []
                if not primary:
                    self._json(
                        400,
                        {
                            "ok": False,
                            "message": (
                                "IDENTITY（主見本）が未選択です。"
                                "副見本や SIZZLE／WRAP だけでは確定できません。"
                            ),
                        },
                    )
                    return
                # 一片あり時: IDENTITY＝一片同一バイトは不合格（盛り見本を別途）
                if piece_has and (not package_only):
                    try:
                        from b5_piece_truth_rules import identity_piece_collision_errors

                        id_errs = identity_piece_collision_errors(
                            theme_new_dir, identity_name=primary, max_n=5
                        )
                        if id_errs:
                            self._json(
                                400,
                                {
                                    "ok": False,
                                    "message": id_errs[0],
                                },
                            )
                            return
                    except Exception as ex:
                        LOG.warning("identity-piece collision check skipped: %s", ex)
                # MAINの個包装正＝検収済み1個画。寄せ集め／低解像は捨てる
                if wrap_has and (not package_only):
                    try:
                        from b4_wrap_compose_ref import (
                            promote_checked_single_to_compose_wrap,
                            resolve_compose_wrap_choice,
                        )

                        compose_prep = resolve_compose_wrap_choice(
                            theme_new_dir,
                            wrap_name=wrap_name,
                            wrap_single_name=wrap_single_name,
                            wrap_has=True,
                        )
                        if not compose_prep.get("ok"):
                            self._json(
                                400,
                                {
                                    "ok": False,
                                    "message": str(
                                        compose_prep.get("message")
                                        or "個包装の1個画を選んでください"
                                    ),
                                },
                            )
                            return
                        wrap_name = str(compose_prep.get("composeName") or "")
                        if wrap_name.startswith("wrap_single_"):
                            wrap_single_name = wrap_name
                        if wrap_name:
                            promote_checked_single_to_compose_wrap(
                                theme_new_dir,
                                wrap_name,
                                seed_name=str(compose_prep.get("seedName") or ""),
                            )
                    except Exception as ex:
                        LOG.exception("compose wrap prepare")
                        self._json(
                            400,
                            {
                                "ok": False,
                                "message": f"個包装見本の確定に失敗: {ex}",
                            },
                        )
                        return
                try:
                    meta = apply_human_selection(
                        out_dir=theme_new_dir,
                        truth=Path(str(session.get("truth") or "")),
                        primary_name=primary,
                        identity_name=primary,
                        sizzle_name=sizzle_name,
                        wrap_name=wrap_name,
                        secondary_names=[] if package_only else secondary,
                        package_only=package_only,
                        name_base=str(
                            session.get("nameBase") or session.get("productName") or ""
                        ),
                        parent_sku=str(session.get("parentSku") or ""),
                    )
                    piece_meta: Dict[str, Any] = {"ok": True, "hasPieceSet": False, "n": 0}
                    try:
                        from b5_piece_gate import activate_piece_seed, apply_piece_selection

                        if piece_has and not package_only and piece_seed_name:
                            activate_piece_seed(theme_new_dir, piece_seed_name)
                        piece_meta = apply_piece_selection(
                            theme_new_dir,
                            piece_has=bool(piece_has) and not package_only,
                            piece_names=piece_names,
                        )
                        meta["hasPieceSet"] = bool(piece_meta.get("hasPieceSet"))
                        meta["pieceHas"] = bool(piece_meta.get("hasPieceSet"))
                        meta["pieceNames"] = list(piece_meta.get("names") or [])
                        meta["pieceSeedName"] = piece_seed_name or None
                    except Exception as ex:
                        LOG.exception("piece selection on confirm")
                        self._json(
                            400,
                            {
                                "ok": False,
                                "message": f"一片見本の確定に失敗: {ex}",
                            },
                        )
                        return
                    try:
                        from b5_food_canon_status import export_after_human_confirm

                        export_after_human_confirm(
                            str(session.get("jan") or ""),
                            parent_sku=str(session.get("parentSku") or ""),
                            name_base=str(
                                session.get("nameBase")
                                or session.get("productName")
                                or ""
                            ),
                        )
                    except Exception as ex:
                        LOG.warning("food canon status export: %s", ex)
                    # 手持ち採用は同一画面で完結（別phaseへ飛ばさない）
                    if hand_hold_name:
                        try:
                            from b4_wrap_hand_hold import (
                                load_manifest,
                                save_manifest,
                                select_hand_hold,
                            )

                            select_hand_hold(
                                theme_new_dir,
                                hand_hold_name,
                                jan=str(session.get("jan") or ""),
                            )
                            if wrap_single_name:
                                man_hh = load_manifest(theme_new_dir)
                                man_hh["selectedSingle"] = wrap_single_name
                                man_hh["composeWrap"] = wrap_single_name
                                man_hh["recommendedSingle"] = wrap_single_name
                                from b4_wrap_asset_rules import (
                                    resolve_wrap_asset_for_consumers,
                                )

                                man_hh["routes"] = resolve_wrap_asset_for_consumers(
                                    man_hh
                                )
                                save_manifest(theme_new_dir, man_hh)
                            session["wrapHandHold"] = {
                                "state": "selected",
                                "selected": hand_hold_name,
                                "selectedSingle": wrap_single_name or "",
                            }
                        except Exception as ex:
                            LOG.warning("hand-hold select on confirm: %s", ex)
                    elif wrap_single_name:
                        try:
                            from b4_wrap_hand_hold import load_manifest, save_manifest
                            from b4_wrap_asset_rules import (
                                resolve_wrap_asset_for_consumers,
                            )

                            man_hh = load_manifest(theme_new_dir) or {}
                            man_hh["selectedSingle"] = wrap_single_name
                            man_hh["composeWrap"] = wrap_single_name
                            man_hh["recommendedSingle"] = wrap_single_name
                            man_hh["routes"] = resolve_wrap_asset_for_consumers(man_hh)
                            save_manifest(theme_new_dir, man_hh)
                        except Exception as ex:
                            LOG.warning("composeWrap on confirm: %s", ex)
                    session["phase"] = "pick"
                    session.pop("draftPicks", None)
                    session.pop("preferredWrapName", None)
                    piece_msg = (
                        f" 一片={piece_meta.get('n') or 0}枚"
                        if piece_meta.get("hasPieceSet")
                        else " 一片=なし"
                    )
                    session["gen"] = {
                        "state": "confirmed",
                        "message": (
                            "見本確定（正本のみ・個品なし）"
                            if package_only
                            else (
                                f"見本確定 IDENTITY={meta.get('identityName')}"
                                f" SIZZLE={meta.get('sizzleName') or 'なし'}"
                                f" WRAP={meta.get('wrapName') or 'なし'}"
                                + (
                                    f" 手持ち={hand_hold_name}"
                                    if hand_hold_name
                                    else ""
                                )
                                + piece_msg
                            )
                        ),
                        "ok": True,
                        "packageOnly": package_only,
                        "nextPhase": "",
                    }
                    _save_session(theme_new_dir, session)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": session["gen"]["message"],
                            "nextPhase": "",
                            "meta": {
                                "primary": meta.get("primaryName"),
                                "identity": meta.get("identityName"),
                                "sizzle": meta.get("sizzleName"),
                                "wrap": meta.get("wrapName"),
                                "secondary": meta.get("secondaryNames"),
                                "identityStatus": meta.get("identityStatus"),
                                "sizzleStatus": meta.get("sizzleStatus"),
                                "wrapStatus": meta.get("wrapStatus"),
                                "packageOnly": meta.get("packageOnly"),
                                "hasPieceSet": bool(piece_meta.get("hasPieceSet")),
                                "pieceNames": list(piece_meta.get("names") or []),
                            },
                        },
                    )
                except ValueError as e:
                    self._json(400, {"ok": False, "message": str(e)})
                except Exception as e:
                    LOG.exception("confirm")
                    self._json(500, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/generate":
                session["phase"] = "generating"
                session["gen"] = {"state": "running", "message": "生成中", "ok": None}
                _save_session(theme_new_dir, session)
                try:
                    result = run_theme_new_generate(theme_new_dir, jan)
                    session = _load_session(theme_new_dir)
                    if result.get("ok"):
                        session["phase"] = "review"
                        session["gen"] = {
                            "state": "done",
                            "message": "生成完了",
                            "ok": True,
                        }
                    else:
                        session["phase"] = "pick"
                        session["gen"] = {
                            "state": "failed",
                            "message": f"生成失敗 code={result.get('code')}",
                            "ok": False,
                            "detail": (result.get("stderr") or "")[-400:],
                        }
                    _save_session(theme_new_dir, session)
                except Exception as e:
                    LOG.exception("generate")
                    session["phase"] = "pick"
                    session["gen"] = {
                        "state": "failed",
                        "message": str(e),
                        "ok": False,
                    }
                    _save_session(theme_new_dir, session)
                session = _load_session(theme_new_dir)
                ok = bool((session.get("gen") or {}).get("ok"))
                self._json(
                    200 if ok else 500,
                    {
                        "ok": ok,
                        "message": (session.get("gen") or {}).get("message") or "",
                        "phase": session.get("phase"),
                    },
                )
                return

            if parsed.path == "/api/refresh-candidates":
                try:
                    prepare_gate(
                        jan=jan,
                        theme_new_dir=theme_new_dir,
                        work_root=DEFAULT_WORK_ROOT,
                        refresh_candidates=True,
                    )
                    s = _load_session(theme_new_dir)
                    s["phase"] = "pick"
                    _save_session(theme_new_dir, s)
                    self._json(200, {"ok": True, "message": "候補を更新しました"})
                except Exception as e:
                    self._json(500, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/to-review":
                require_images = True
                try:
                    raw = self.rfile.read(int(self.headers.get("Content-Length") or 0))
                    if raw:
                        body = json.loads(raw.decode("utf-8"))
                        if "requireImages" in body:
                            require_images = bool(body.get("requireImages"))
                except Exception:
                    pass
                n_img = sum(1 for p in theme_new_dir.glob("T*.jpg") if p.is_file())
                if require_images and n_img <= 0:
                    self._json(
                        400,
                        {
                            "ok": False,
                            "message": (
                                "既存のテーマ画像（T*.jpg）がありません。"
                                "初回は「見本確定して画像生成」を使うか、先に dry_sub で生成してください。"
                            ),
                        },
                    )
                    return
                session["phase"] = "review"
                session["gen"] = {
                    "state": "skipped",
                    "message": f"画像の生成不要（既存 {n_img} 枚でレビュー）",
                    "ok": True,
                    "skipGenerate": True,
                }
                _save_session(theme_new_dir, session)
                self._json(
                    200,
                    {
                        "ok": True,
                        "message": f"レビューへ（生成スキップ・既存 {n_img} 枚）",
                        "imageN": n_img,
                    },
                )
                return

            if parsed.path == "/api/wrap-hand-hold/generate":
                force = bool(payload.get("force"))
                bundle = bool(payload.get("bundle"))
                wrap_seed = str(
                    payload.get("wrapSeedName") or payload.get("wrapSeed") or ""
                ).strip()
                try:
                    from b4_wrap_hand_hold import (
                        generate_hand_hold_candidates,
                        generate_wrap_asset_bundle,
                    )

                    if bundle:
                        result = generate_wrap_asset_bundle(
                            theme_new_dir,
                            jan=str(session.get("jan") or ""),
                            wrap_seed_name=wrap_seed,
                            force=force,
                        )
                        n_h = len((result.get("hand") or {}).get("candidates") or [])
                        n_s = len((result.get("single") or {}).get("candidates") or [])
                        session["wrapHandHold"] = {
                            "state": "candidates",
                            "nHand": n_h,
                            "nSingle": n_s,
                            "recommend": result.get("recommend") or {},
                        }
                        _save_session(theme_new_dir, session)
                        self._json(
                            200,
                            {
                                "ok": True,
                                "message": f"手持ち{n_h}／単体{n_s}（再利用含む）",
                                "result": result,
                            },
                        )
                    else:
                        result = generate_hand_hold_candidates(
                            theme_new_dir,
                            jan=str(session.get("jan") or ""),
                            force=force,
                            wrap_seed_name=wrap_seed,
                        )
                        session["wrapHandHold"] = {
                            "state": "candidates",
                            "n": len(result.get("candidates") or []),
                            "reused": bool(result.get("reused")),
                        }
                        _save_session(theme_new_dir, session)
                        self._json(
                            200,
                            {
                                "ok": True,
                                "message": (
                                    "既存候補を表示"
                                    if result.get("reused")
                                    else f"候補 {len(result.get('candidates') or [])} 枚を生成"
                                ),
                                "candidates": result.get("candidates") or [],
                                "reused": bool(result.get("reused")),
                            },
                        )
                except Exception as e:
                    LOG.exception("wrap-hand-hold generate")
                    self._json(500, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/wrap-hand-hold/remake":
                try:
                    from b4_wrap_hand_hold import remake_named_wrap_assets

                    result = remake_named_wrap_assets(
                        theme_new_dir,
                        names=list(payload.get("names") or []),
                        kind=str(payload.get("kind") or ""),
                        jan=str(session.get("jan") or ""),
                        wrap_seed_name=str(
                            payload.get("wrapSeedName") or payload.get("wrapSeed") or ""
                        ).strip(),
                        user_note=str(
                            payload.get("userNote") or payload.get("remakeNote") or ""
                        ).strip(),
                    )
                    n = len(result.get("done") or [])
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"{result.get('kind')} を {n} 枚だけ作り直しました",
                            **result,
                        },
                    )
                except ValueError as e:
                    self._json(400, {"ok": False, "message": str(e)})
                except Exception as e:
                    LOG.exception("wrap-hand-hold remake")
                    self._json(500, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/wrap-hand-hold/select":
                try:
                    raw = self.rfile.read(int(self.headers.get("Content-Length") or 0))
                    body = json.loads(raw.decode("utf-8")) if raw else {}
                    name = str(body.get("name") or "").strip()
                    from b4_wrap_hand_hold import select_hand_hold
                    from b5_food_canon_status import export_after_human_confirm

                    result = select_hand_hold(
                        theme_new_dir,
                        name,
                        jan=str(session.get("jan") or ""),
                    )
                    session["wrapHandHold"] = {
                        "state": "selected",
                        "selected": result.get("selected"),
                        "path": result.get("selectedPath"),
                    }
                    if str(session.get("context") or "").strip().lower() == "main":
                        session["phase"] = "pick"
                    _save_session(theme_new_dir, session)
                    try:
                        export_after_human_confirm(
                            str(session.get("jan") or ""),
                            parent_sku=str(session.get("parentSku") or ""),
                            name_base=str(
                                session.get("nameBase")
                                or session.get("productName")
                                or ""
                            ),
                        )
                    except Exception as ex:
                        LOG.warning("status export after hand-hold: %s", ex)
                    self._json(
                        200,
                        {
                            "ok": True,
                            "message": f"採用しました: {result.get('selected')}",
                            "selected": result.get("selected"),
                            "routes": result.get("routes") or {},
                        },
                    )
                except ValueError as e:
                    self._json(400, {"ok": False, "message": str(e)})
                except Exception as e:
                    LOG.exception("wrap-hand-hold select")
                    self._json(500, {"ok": False, "message": str(e)})
                return

            if parsed.path == "/api/wrap-hand-hold/skip":
                try:
                    from b4_wrap_hand_hold import clear_selection

                    clear_selection(theme_new_dir)
                    session["wrapHandHold"] = {"state": "skipped"}
                    if str(session.get("context") or "").strip().lower() == "main":
                        session["phase"] = "pick"
                    _save_session(theme_new_dir, session)
                    self._json(200, {"ok": True, "message": "手持ちなしで進めます"})
                except Exception as e:
                    LOG.exception("wrap-hand-hold skip")
                    self._json(500, {"ok": False, "message": str(e)})
                return

            self._send(404, b"not found", "text/plain")

    return Handler


def serve(theme_new_dir: Path, port: int, open_browser: bool) -> int:
    port_i = int(port)
    handler = make_handler(theme_new_dir, SCRIPT_DIR, listen_port=port_i)

    class _NoReuseHTTPServer(ThreadingHTTPServer):
        """SO_REUSEADDR を切る。起動前掃除と合わせ、同じポートの二重 Listen を減らす。"""

        allow_reuse_address = False

    try:
        httpd = _NoReuseHTTPServer(("127.0.0.1", port_i), handler)
    except OSError as e:
        raise SystemExit(
            f"ポート {port_i} を使えない（旧ゲートが残っている可能性）: {e}\n"
            "起動は sub_image_food_canon_launch.py 経由（対象ポートを先に掃除）。"
        ) from e
    url = f"http://127.0.0.1:{port_i}/"
    LOG.info("food canon gate: %s", url)
    print(url)
    if open_browser:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        LOG.info("stopped")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="食材見本ゲート→生成→レビュー")
    ap.add_argument("--jan", default="")
    ap.add_argument("--name-base", default="", dest="name_base")
    ap.add_argument("--parent-sku", default="", dest="parent_sku")
    ap.add_argument("--part-name", action="append", default=[], dest="part_name")
    ap.add_argument("--part-sku", action="append", default=[], dest="part_sku")
    ap.add_argument("--part-jan", action="append", default=[], dest="part_jan")
    ap.add_argument("--part-bag", action="append", default=[], dest="part_bag")
    ap.add_argument("--theme-new-dir", type=Path, default=None)
    ap.add_argument("--work-root", type=Path, default=DEFAULT_WORK_ROOT)
    ap.add_argument("--port", type=int, default=8790)
    ap.add_argument("--refresh-candidates", action="store_true")
    ap.add_argument("--force-pick", action="store_true", help="既存があっても見本選択から")
    ap.add_argument(
        "--context",
        default="",
        help="main = B-④ MAIN用2ボタンUI",
    )
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)
    part_names = [str(x).strip() for x in (args.part_name or []) if str(x).strip()]
    if len(part_names) >= 2:
        sku = str(args.parent_sku or "").strip() or "mix"
        tn = (
            Path(args.theme_new_dir)
            if args.theme_new_dir
            else (SCRIPT_DIR / "_out" / f"mix_run_{sku}")
        )
        tn.mkdir(parents=True, exist_ok=True)
        parts = []
        for i, nb in enumerate(part_names):
            parts.append(
                {
                    "nameBase": nb,
                    "parentSku": (args.part_sku[i] if i < len(args.part_sku) else "") or "",
                    "jan": (args.part_jan[i] if i < len(args.part_jan) else "") or "",
                    "bag": (args.part_bag[i] if i < len(args.part_bag) else "") or "",
                }
            )
        session = {
            "mode": "mix",
            "jan": str(args.jan or ""),
            "nameBase": str(args.name_base or "").strip(),
            "parentSku": str(args.parent_sku or "").strip(),
            "parts": parts,
            "phase": "pick",
            "context": str(args.context or "").strip(),
            "themeNewDir": str(tn.resolve()),
        }
        _save_session(tn, session)
        return serve(tn, int(args.port), open_browser=not args.no_browser)
    if not str(args.jan or "").strip():
        raise SystemExit("単品は --jan、組み合わせは --part-name を2つ以上")
    tn = (
        Path(args.theme_new_dir)
        if args.theme_new_dir
        else (SCRIPT_DIR / "_out" / f"theme_new_run_{args.jan}")
    )
    session = prepare_gate(
        jan=str(args.jan),
        theme_new_dir=tn,
        work_root=Path(args.work_root),
        refresh_candidates=bool(args.refresh_candidates),
    )
    if args.force_pick:
        slim = {k: session[k] for k in session if k != "candidates"}
        slim["phase"] = "pick"
        if str(args.context or "").strip():
            slim["context"] = str(args.context).strip()
        if str(args.name_base or "").strip():
            slim["nameBase"] = str(args.name_base).strip()
        if str(args.parent_sku or "").strip():
            slim["parentSku"] = str(args.parent_sku).strip()
        _save_session(tn, slim)
    else:
        s = _load_session(tn)
        if str(args.context or "").strip():
            s["context"] = str(args.context).strip()
        if str(args.name_base or "").strip():
            s["nameBase"] = str(args.name_base).strip()
        if str(args.parent_sku or "").strip():
            s["parentSku"] = str(args.parent_sku).strip()
        if s:
            _save_session(tn, s)
    return serve(tn, int(args.port), open_browser=not args.no_browser)


if __name__ == "__main__":
    raise SystemExit(main())
