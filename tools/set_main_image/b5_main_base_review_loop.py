# -*- coding: utf-8 -*-
"""
B-⑤ 楽天MAIN基礎 — レビュー／再生成／複数採用＋レ点子割当／02（P3b）。

  python b5_main_base_review_loop.py --compose-dir "...\\b5_main_base_compose\\かつお_…"
"""
from __future__ import annotations

import argparse
import html
import json
import logging
import re
import shutil
import subprocess
import threading
import webbrowser
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import unquote, urlparse

from b5_main_base_auto_assign import build_auto_assign_plan
from b5_checked_children import load_checked_children_for_parent
from b5_main_base_names import (
    catch_head_ja,
    export02_filename,
    merge_catch_for_regen,
    pick_filename,
    split_catch_for_ui,
)
from b5_review_feedback_templates import TEMPLATES, VERSION as TPL_VER
from b5_set_base_map import (
    adopt_key_from_filename,
    next_regen_filename,
    picture_family_key,
    validate_map,
    write_map,
)
from work_paths import DEFAULT_WORK_ROOT, SUB_RAKUTEN

LOG = logging.getLogger("set_main_image.b5_main_base_review_loop")


def _setup_log(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _load_proposals(compose_dir: Path) -> Dict[str, Any]:
    p = Path(compose_dir) / "b5_proposals.json"
    if not p.is_file():
        return {}
    return json.loads(p.read_text(encoding="utf-8-sig"))


def _theme_from_name(name: str) -> str:
    m = re.search(r"_([NT]\d+[a-z]?)_", name)
    return m.group(1) if m else ""


def _pair_label(adopt_key: str, *, has_sibling: bool) -> str:
    """同じ見出しが複数あるときだけ、前／再作成と出す。"""
    m = re.search(r"__r(\d+)$", str(adopt_key or ""))
    if not m:
        return "前" if has_sibling else ""
    n = int(m.group(1))
    if n <= 2:
        return "再作成"
    return f"再作成{n - 1}回目"


_STAGING_PREFIX = "__regen_staging_"


def _purge_regen_staging(compose_dir: Path) -> int:
    """再生成用ステージングを比較から削除。表示二重化の原因になるため残さない。"""
    pick = Path(compose_dir) / "比較"
    n = 0
    if not pick.is_dir():
        return 0
    for p in list(pick.glob(f"{_STAGING_PREFIX}*")):
        try:
            p.unlink()
            n += 1
        except Exception:
            LOG.warning("staging delete fail: %s", p)
    return n


def _proposal_for_file(
    proposals: List[Any],
    *,
    akey: str,
    tid: str,
) -> Dict[str, Any]:
    """ファイル名の見出しで案を取る。同じテーマ番号の最後の案で上書きしない。"""
    rows = [p for p in proposals if isinstance(p, dict)]
    family = picture_family_key(akey)
    tail = family.split("__", 1)[1].strip() if "__" in family else ""

    def _heads(p: Dict[str, Any]) -> set:
        out = set()
        for raw in (p.get("catch"), p.get("catchBurn")):
            s = str(raw or "").strip()
            if not s:
                continue
            out.add(s)
            head = catch_head_ja(s)
            if head:
                out.add(head)
        return out

    if tail:
        exact = [p for p in rows if tail in _heads(p)]
        loose = [
            p
            for p in rows
            if any(h.startswith(tail) or tail.startswith(h) for h in _heads(p) if h)
        ]
        pool = exact or loose
        if tid:
            themed = [p for p in pool if str(p.get("themeId") or "") == tid]
            if themed:
                return themed[0]
        if pool:
            return pool[0]
        # ファイル見出しがあるのに案が無いとき、同テーマ先頭へ落とさない（別カードの文言が混ざる）
        return {}
    for p in rows:
        if str(p.get("themeId") or "") == tid:
            return p
    return {}


def _file_headline_from_akey(akey: str) -> str:
    """adoptKey の見出し部分（再作成 __rN を除く）。"""
    fam = picture_family_key(akey)
    if "__" not in fam:
        return ""
    return fam.split("__", 1)[1].strip()


def _ui_catch_sales_for_card(
    pr: Dict[str, Any],
    *,
    file_headline: str,
) -> tuple[str, str, str, str]:
    """カード欄用: 焼成文言を優先。catch 全文の機械分割は焼成が無いときだけ。

    Returns: (catch_full, catch_head, sales_talk, catch_extra)
    """
    catch_full = str(pr.get("catch") or "").strip()
    burn = str(pr.get("catchBurn") or "").strip()
    sales = str(pr.get("salesBurn") or pr.get("salesTalk") or "").strip()
    extra = ""
    if burn:
        # ファイル名見出しと焼成が食い違う古い残骸は、ファイル名をタイトル／欄の正にする
        if file_headline and file_headline != burn and file_headline not in burn and burn not in file_headline:
            head = file_headline
        else:
            head = burn
        if not sales:
            _, sales_split, extra = split_catch_for_ui(catch_full)
            sales = sales_split
        return catch_full or burn, head, sales, extra
    head, sales_split, extra = split_catch_for_ui(catch_full)
    if file_headline and not head:
        head = file_headline
    return catch_full, head, (sales or sales_split), extra


def _load_regen_queue_items(compose_dir: Path) -> List[Dict[str, Any]]:
    qpath = Path(compose_dir) / "_meta" / "regen_queue.json"
    if not qpath.is_file():
        return []
    try:
        data = json.loads(qpath.read_text(encoding="utf-8"))
    except Exception:
        return []
    items = data.get("items") if isinstance(data, dict) else None
    return [x for x in (items or []) if isinstance(x, dict)]


def _pending_queue_for_row(
    row: Dict[str, Any],
    by_name: Dict[str, Dict[str, Any]],
    by_akey: Dict[str, Dict[str, Any]],
) -> Dict[str, str]:
    """前回レ点入力のうち、いまの焼成欄と違うものだけ（自動上書きはしない）。"""
    it = by_name.get(str(row.get("name") or "")) or by_akey.get(
        str(row.get("adoptKey") or "")
    )
    if not it:
        return {}
    ch = str(it.get("catchHead") or "").strip()
    st = str(it.get("salesTalk") or "").strip()
    cur_ch = str(row.get("catchHead") or "").strip()
    cur_st = str(row.get("salesTalk") or "").strip()
    out: Dict[str, str] = {}
    if ch and ch != cur_ch:
        out["catchHead"] = ch
    if st and st != cur_st:
        out["salesTalk"] = st
    return out


def _attach_pending_queue_meta(compose_dir: Path, rows: List[Dict[str, Any]]) -> None:
    by_name: Dict[str, Dict[str, Any]] = {}
    by_akey: Dict[str, Dict[str, Any]] = {}
    for it in _load_regen_queue_items(compose_dir):
        name = str(it.get("name") or "").strip()
        akey = str(it.get("adoptKey") or "").strip()
        if name:
            by_name[name] = it
        if akey:
            by_akey[akey] = it
    for r in rows:
        pend = _pending_queue_for_row(r, by_name, by_akey)
        if pend:
            r["pendingQueue"] = pend


def list_compare_rows(compose_dir: Path) -> List[Dict[str, Any]]:
    pick = Path(compose_dir) / "比較"
    rows: List[Dict[str, Any]] = []
    if not pick.is_dir():
        return rows
    _purge_regen_staging(compose_dir)
    meta = _load_proposals(compose_dir)
    proposals = [x for x in (meta.get("proposals") or []) if isinstance(x, dict)]
    files = sorted(pick.glob("*.jpg")) + sorted(pick.glob("*.png"))
    # 履歴フォルダは一覧に出さない。再生成ステージングも出さない
    for p in files:
        if p.name.startswith(_STAGING_PREFIX):
            continue
        tid = _theme_from_name(p.name)
        akey = adopt_key_from_filename(p.name)
        pr = _proposal_for_file(proposals, akey=akey, tid=tid)
        file_head = _file_headline_from_akey(akey)
        catch_full, catch_head, sales_talk, catch_extra = _ui_catch_sales_for_card(
            pr, file_headline=file_head
        )
        # 欄の正＝絵に焼いてある文言。queue を自動で上書きしない（ズレの原因になる）
        if not catch_head and file_head:
            catch_head = file_head
        rows.append(
            {
                "name": p.name,
                "rel": f"比較/{p.name}".replace("\\", "/"),
                "themeId": tid or pr.get("themeId") or "",
                "adoptKey": akey,
                "themeName": pr.get("themeName") or "",
                "catch": catch_full,
                "catchHead": catch_head,
                "salesTalk": sales_talk,
                "catchExtra": catch_extra,
            }
        )
    _attach_pending_queue_meta(compose_dir, rows)
    return rows


def _patch_proposal_card(
    compose_dir: Path,
    *,
    theme_id: str,
    akey: str,
    catch: str = "",
    catch_burn: str = "",
    sales_talk: str = "",
) -> None:
    """再生成後に、ファイル見出しで特定した案だけ同期（同じ themeId の先頭上書き禁止）。"""
    p = Path(compose_dir) / "b5_proposals.json"
    if not p.is_file():
        return
    try:
        meta = json.loads(p.read_text(encoding="utf-8-sig"))
        props = [x for x in (meta.get("proposals") or []) if isinstance(x, dict)]
        row = _proposal_for_file(props, akey=akey, tid=theme_id)
        if not row:
            LOG.warning("proposal patch skip: no row theme=%s akey=%s", theme_id, akey)
            return
        if catch.strip():
            row["catch"] = catch.strip()
        if catch_burn.strip():
            row["catchBurn"] = catch_burn.strip()
        if sales_talk.strip():
            row["salesTalk"] = sales_talk.strip()
            row["salesBurn"] = sales_talk.strip()
        p.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        LOG.warning("proposal patch failed theme=%s akey=%s", theme_id, akey)


def _regen_dest_name(
    *,
    compose_dir: Path,
    old_name: str,
    short_name: str,
    theme_id: str,
    catch_head: str,
) -> str:
    """キャッチ変更時は新見出しのファイル名。同家系は旧名から __rN。別見出しで衝突なら __rN。"""
    pick = Path(compose_dir) / "比較"
    head = str(catch_head or "").strip()
    if not head:
        head = _file_headline_from_akey(adopt_key_from_filename(old_name))
    if not head:
        return next_regen_filename(old_name)
    ideal = pick_filename(short_name, theme_id, head)
    old_fam = picture_family_key(adopt_key_from_filename(old_name))
    new_fam = picture_family_key(adopt_key_from_filename(ideal))
    if old_fam == new_fam:
        return next_regen_filename(old_name)
    if not (pick / ideal).is_file():
        return ideal
    return next_regen_filename(ideal)


def _catch_texts_align(expected: str, actual: str) -> bool:
    """指定キャッチと OCR／焼成が見た目上一致するか。"""
    a = re.sub(r"\s+", "", str(expected or "").strip())
    b = re.sub(r"\s+", "", str(actual or "").strip())
    if not a or not b:
        return False
    if a == b or a in b or b in a:
        return True
    ha = catch_head_ja(a)
    hb = catch_head_ja(b)
    return bool(ha and hb and (ha == hb or ha in hb or hb in ha))


def _regen_name_base(meta: Dict[str, Any]) -> str:
    """ミックス部品解決用の親商品名。シート材料より親SKU共有の nameBase を優先。"""
    nb = str(meta.get("nameBase") or meta.get("name_base") or "").strip()
    if nb:
        return nb
    sku = str(meta.get("parentSku") or "").strip()
    if sku:
        try:
            from b5_parent_shared import load_parent_shared

            row = (load_parent_shared(None).get("byParentSku") or {}).get(sku) or {}
            nb = str(row.get("nameBase") or "").strip()
            if nb:
                return nb
        except Exception as e:
            LOG.warning("regen nameBase from parent_shared: %s", e)
    mats = meta.get("materials") if isinstance(meta.get("materials"), dict) else {}
    return str(mats.get("nameBase") or mats.get("name") or "").strip()


def _result_picked_try(row: Dict[str, Any]) -> Dict[str, Any]:
    tries = [t for t in (row.get("tries") or []) if isinstance(t, dict)]
    if not tries:
        return {}
    picked = str(row.get("pickedTry") or "").strip()
    if picked:
        for t in tries:
            if str(t.get("file") or "") == picked:
                return t
    # rank 最小が採用（compose 側と同じ）
    return min(tries, key=lambda t: int(t.get("rank") or 10**9))


def _assert_regen_catch_matches(
    *,
    expected_head: str,
    result_row: Dict[str, Any],
    old_name: str,
    dest_name: str,
) -> None:
    """指定キャッチと焼成／保存名が見出しと食い違うときは __r2 に載せない。"""
    expected = str(expected_head or "").strip()
    if not expected:
        return
    catch_image = str(result_row.get("catchImage") or "").strip()
    if catch_image and not _catch_texts_align(expected, catch_image):
        raise RuntimeError(
            f"生成指定キャッチ不一致のため保存しない: 欄={expected} 生成={catch_image} "
            f"（旧={old_name}）"
        )
    picked = _result_picked_try(result_row)
    ocr = str(picked.get("ocrCatch") or "").strip()
    ocr_ok = picked.get("catchOcrOk")
    ocr_skipped = bool(picked.get("catchOcrSkipped"))
    if ocr and ocr_ok is False and not ocr_skipped:
        raise RuntimeError(
            f"焼成キャッチ不一致のため保存しない: 指定={expected} OCR={ocr} "
            f"（旧={old_name} → 予定={dest_name}）"
        )
    if ocr and not ocr_skipped and not _catch_texts_align(expected, ocr):
        raise RuntimeError(
            f"焼成キャッチ不一致のため保存しない: 指定={expected} OCR={ocr} "
            f"（旧={old_name} → 予定={dest_name}）"
        )
    dest_head = _file_headline_from_akey(adopt_key_from_filename(dest_name))
    if dest_head and not _catch_texts_align(expected, dest_head):
        raise RuntimeError(
            f"保存名の見出しが指定キャッチと不一致: 指定={expected} 名={dest_name}"
        )


def build_html(
    compose_dir: Path,
    rows: List[Dict[str, Any]],
    meta: Dict[str, Any],
    children: List[Dict[str, Any]],
    children_err: str,
) -> str:
    short = str(meta.get("shortName") or Path(compose_dir).name)
    sku = str(meta.get("parentSku") or "")
    jan = str(meta.get("jan") or "")
    auto_plan = build_auto_assign_plan(
        rows, meta, children, compose_dir=Path(compose_dir)
    )
    auto_json = json.dumps(auto_plan, ensure_ascii=False)
    child_json = json.dumps(children, ensure_ascii=False)
    family_n: Dict[str, int] = {}
    for r in rows:
        fam = picture_family_key(str(r.get("adoptKey") or r.get("themeId") or ""))
        family_n[fam] = family_n.get(fam, 0) + 1
    cards = []
    for r in rows:
        akey = r.get("adoptKey") or r.get("themeId") or "?"
        ch_val = html.escape(str(r.get("catchHead") or ""), quote=True)
        sl_val = html.escape(str(r.get("salesTalk") or ""), quote=True)
        catch_show = str(r.get("catchHead") or "").strip() or str(r.get("themeName") or akey)
        fam = picture_family_key(str(akey))
        ver = _pair_label(str(akey), has_sibling=family_n.get(fam, 0) > 1)
        title = catch_show if not ver else f"{catch_show}（{ver}）"
        pend = r.get("pendingQueue") if isinstance(r.get("pendingQueue"), dict) else {}
        pend_ch = html.escape(str(pend.get("catchHead") or ""), quote=True)
        pend_sl = html.escape(str(pend.get("salesTalk") or ""), quote=True)
        pend_btn = ""
        if pend_ch or pend_sl:
            pend_btn = (
                f'<button type="button" class="pend-restore" '
                f'data-pend-catch="{pend_ch}" data-pend-sales="{pend_sl}">'
                f"前回の未反映入力を欄へ</button>"
            )
        cards.append(
            f"""
            <div class="card" data-name="{html.escape(str(r['name']), quote=True)}" data-theme="{html.escape(str(r.get('themeId') or ''), quote=True)}" data-adopt-key="{html.escape(str(akey), quote=True)}" data-regen-mode="full">
              <label class="chk"><input type="checkbox" class="need-regen"/> 再生成する</label>
              <label class="chk adopt"><input type="checkbox" class="adopt-chk"/> 採用する（02へ）</label>
              <div class="meta"><b>{html.escape(title)}</b><br/><span class="file">{html.escape(str(r['name']))}</span></div>
              <div class="img-wrap" data-rel="{html.escape(str(r['rel']), quote=True)}">
                <img class="thumb" src="/file/{r['rel']}" alt="{html.escape(str(r['name']), quote=True)}" title="クリックで拡大"/>
                <div class="grid-overlay" aria-hidden="true"><div class="grid-cells"></div></div>
              </div>
              <div class="grid-badge" data-grid-badge></div>
              <div class="child-box"><b>レ点子SKU割当</b><div class="child-list"></div></div>
              <div class="tpl-btns"></div>
              <div class="text-edit">
                <label>キャッチ（大）</label>
                <input class="catch-head" type="text" value="{ch_val}" title="絵の大きい字。再生成時に焼成キャッチへ反映"/>
                <label>セールストーク（小）</label>
                <input class="sales-talk" type="text" value="{sl_val}" title="絵の小さい字。再生成時に反映"/>
                <label>個品微調整（任意・厚み等）</label>
                <input class="food-tune" type="text" placeholder="例: 厚みを1段薄く／短冊はこの厚さ維持"/>
                {pend_btn}
              </div>
              <label>構図・演出の要望</label>
              <textarea class="comment" placeholder="定型可。文字だけ差替なら「文字だけ差替」ボタン＋上のキャッチ欄を編集"></textarea>
            </div>
            """
        )
    cards_html = "\n".join(cards) or "<p>比較フォルダに画像がありません。</p>"
    tpl_json = json.dumps(TEMPLATES, ensure_ascii=False)
    tpl_list = "".join(
        f'<button type="button" class="tpl-global" data-idx="{i}">{t["label"]}</button> '
        for i, t in enumerate(TEMPLATES)
    )
    err_banner = (
        f'<p class="warn">レ点子読取: {children_err}</p>' if children_err else ""
    )
    tries_index = Path(compose_dir) / "_tries" / "index.html"
    tries_link = (
        f'<p style="margin:6px 0 0;font-size:13px">'
        f'<a href="/file/_tries/index.html" target="_blank">compose try 比較（Gemini gate／catch_v8b）</a>'
        f"</p>"
        if tries_index.is_file()
        else ""
    )
    return f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>B-⑤レビュー — {Path(compose_dir).name}</title>
<style>
body{{font-family:sans-serif;margin:16px;background:#f6f3ee;color:#222}}
h1{{font-size:18px;margin:0 0 8px}}
.toolbar{{position:sticky;top:0;background:#f6f3ee;padding:8px 0 12px;z-index:2;border-bottom:1px solid #ccc}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:12px;margin-top:12px}}
.card{{background:#fff;border:1px solid #ddd;padding:8px;border-radius:6px}}
.card.grid-over16{{border:2px solid #e6a700;box-shadow:0 0 0 2px rgba(230,167,0,.15)}}
.card.grid-over24{{border:2px solid #c62828;box-shadow:0 0 0 2px rgba(198,40,40,.15)}}
.card.grid-text-warn{{outline:2px dashed #e65100;outline-offset:2px}}
#gridOverList{{font-size:13px;margin:8px 0 0;padding:8px 10px;background:#fff;border:1px solid #ddd;display:none}}
#gridOverList.has-over{{display:block;border-color:#e6a700;background:#fffdf5}}
#gridOverList li.over24{{color:#b00020;font-weight:bold}}
.img-wrap{{position:relative;width:100%}}
.img-wrap img.thumb{{width:100%;height:auto;display:block;border:1px solid #eee;cursor:zoom-in}}
.grid-overlay{{display:none;position:absolute;inset:0;pointer-events:none;z-index:3}}
.grid-overlay .grid-cells{{display:grid;grid-template-columns:repeat(10,1fr);grid-template-rows:repeat(10,1fr);width:100%;height:100%;position:absolute;inset:0}}
.grid-overlay span.grid-cell{{box-sizing:border-box;min-width:0;min-height:0;border:1px solid rgba(255,255,255,.78);background:rgba(255,255,255,.06)}}
.grid-overlay span.roi-focus{{border:2px solid rgba(0,230,120,.96);background:rgba(0,190,95,.2);box-shadow:0 0 0 1px rgba(255,255,255,.45)}}
.grid-overlay span.exp-zone{{border:1px dashed rgba(255,200,120,.85);background:rgba(255,160,0,.12)}}
.grid-overlay span.hit-catch{{background:rgba(210,25,35,.58);border-color:rgba(255,255,255,.78);border-width:1px;border-style:solid}}
.grid-overlay span.hit-sales{{background:rgba(110,45,175,.52);border-color:rgba(255,255,255,.78);border-width:1px;border-style:solid}}
.grid-overlay span.hit-both{{background:linear-gradient(135deg,rgba(210,25,35,.55) 0%,rgba(110,45,175,.55) 100%);border-color:rgba(255,255,255,.78);border-width:1px;border-style:solid}}
.grid-overlay span.roi-focus.hit-catch,.grid-overlay span.roi-focus.hit-sales,.grid-overlay span.roi-focus.hit-both{{border:2px solid rgba(0,230,120,.96);box-shadow:inset 0 0 0 2px rgba(0,230,120,.88)}}
body.grid-on .grid-overlay{{display:block}}
body.grid-on .grid-badge{{display:inline-block!important}}
.grid-badge{{font-size:11px;margin-top:4px;padding:2px 6px;border-radius:4px;display:inline-block}}
.grid-badge.ok{{background:#e8f5e9;color:#063}}
.grid-badge.warn{{background:#fff8e1;color:#a60}}
.grid-badge.bad{{background:#fdecea;color:#b00020}}
.meta{{font-size:12px;color:#555;margin:4px 0 6px}}
.meta b{{font-size:18px;color:#222}}
.meta .file{{font-size:11px;color:#888}}
.text-edit{{font-size:12px;margin:8px 0;padding:8px;background:#f7f9fc;border:1px solid #dde}}
.text-edit label{{display:block;margin:4px 0 2px;color:#444;font-size:11px}}
.text-edit input{{width:100%;box-sizing:border-box;font-size:12px;margin-bottom:6px}}
.text-edit button.pend-restore{{font-size:11px;padding:4px 8px;margin:2px 0 6px;background:#fff8e1;border:1px solid #e6a700}}
textarea{{width:100%;min-height:48px;font-size:12px;box-sizing:border-box}}
button{{margin-right:8px;padding:8px 12px;cursor:pointer}}
.tpl-global,.tpl-local{{font-size:11px;padding:4px 8px;margin:2px 4px 2px 0}}
#status{{margin-left:8px;font-size:13px;color:#333}}
#regenBanner{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:200;align-items:center;justify-content:center}}
#regenBanner.on{{display:flex}}
#regenBanner .box{{background:#b00020;color:#fff;font-size:32px;font-weight:bold;padding:28px 40px;border-radius:8px;text-align:center;line-height:1.45;max-width:90vw}}
#regenBanner.done .box{{background:#1565c0}}
#regenBanner .box span{{display:block;margin-top:10px;font-size:16px;font-weight:normal}}
.hint{{font-size:12px;color:#666;margin:4px 0 0}}
.tpl-bar{{font-size:12px;margin-top:8px;padding:8px;background:#fff;border:1px solid #ddd}}
label.chk{{display:block;font-size:13px;margin:2px 0}}
label.adopt{{font-weight:bold;color:#063}}
.child-box{{font-size:12px;margin:6px 0;padding:6px;background:#faf8f4;border:1px solid #eee}}
.child-list label{{display:block;margin:2px 0}}
.child-list label.child-mine{{color:#0653c7;font-weight:bold}}
.child-list label.child-taken{{color:#b0b0b0;opacity:0.55}}
.child-list label.child-free{{color:#222}}
#childRemain{{font-size:13px;margin:8px 0 0;padding:6px 8px;background:#fff;border:1px solid #ddd}}
#childRemain.ok{{border-color:#6a9;color:#063}}
#childRemain.ng{{border-color:#c66;background:#fff8f8}}
.warn{{color:#b00020;font-size:13px}}
#lightbox{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:99;align-items:center;justify-content:center;cursor:zoom-out}}
#lightbox.on{{display:flex}}
#lightbox .lb-wrap{{position:relative;max-width:96vw;max-height:96vh}}
#lightbox img{{max-width:96vw;max-height:96vh;box-shadow:0 0 24px #000;display:block}}
#lightbox .grid-overlay{{background-size:10% 10%}}
#gridHint{{font-size:12px;color:#555;margin-left:4px}}
</style></head><body>
<div id="regenBanner"><div class="box">再作成しています<span>終わるまでこの画面を閉じないでください</span></div></div>
<div class="toolbar">
  <h1>B-⑤ MAINレビュー（写真／再生成世代保持／複数採用＋レ点子→02）</h1>
  <div class="hint">dir: {compose_dir}<br/>
  {short} ／ JAN {jan} ／ 親SKU {sku}<br/>
  画像クリック＝拡大。再生成は上書きせず新ファイル追加。同じ見出しは残す1枚だけ「採用する」。子SKUはその1枚に付く。<br/>
  割当済み＝<span style="color:#0653c7;font-weight:bold">青字</span>／他画像で使用中＝薄字／未割当＝通常色。未割当があると02へ進めません。<br/>
  テンプレ v{TPL_VER}</div>
  {tries_link}
  {err_banner}
  <div class="tpl-bar"><b>定型（フォーカス中の要望欄へ）:</b><br/>{tpl_list}</div>
  <div id="childRemain" class="ng">レ点子の割当状況を集計中…</div>
  <div id="nextStep" style="margin:10px 0 0;padding:10px 12px;border:2px solid #0653c7;border-radius:6px;background:#eef5ff;font-size:15px;line-height:1.45;font-weight:bold;color:#063;">いまやること: 準備中…</div>
  <p style="margin:10px 0 0">
    <b>①</b>
    <button onclick="export02()" id="btnExport02" style="font-weight:bold;font-size:15px;">採用分を02へコピー</button>
    <span id="step2Label" style="margin-left:10px;color:#666;">→ ①のあと ②</span>
    <button onclick="showCSetMainPrompt()" id="btnCSetMain" disabled title="①のあとに使います">続けてC楽天量産の指示（今の親）</button>
    <span id="status" style="margin-left:8px;"></span>
  </p>
  <details style="margin:8px 0 0;font-size:13px;color:#444;">
    <summary style="cursor:pointer;">途中の作業（再作成・確認）</summary>
    <p style="margin:8px 0 0;">
      <button onclick="saveQueue()">修正指示を一時保存</button>
      <button onclick="runRegen()">レ点再生成実行</button>
      <button onclick="openFolder()">比較フォルダを開く</button>
      <button onclick="location.reload()">再読込</button>
      <button type="button" id="btnGrid" onclick="toggleGrid()">100マスガイド</button>
      <span id="gridHint">（OFF＝ガイド非表示）</span>
    </p>
  </details>
  <div id="cSetMainBox" style="display:none;margin-top:10px;padding:10px;border:1px solid #0a7;border-radius:6px;background:#f3fff6;">
    <div style="font-weight:bold;margin-bottom:6px;">② 次工程: C セットMAIN量産（楽天）— 今の親だけ</div>
    <p style="font-size:13px;margin:0 0 8px;color:#063;">②は「指示を全選択」→ Cursor の新しいチャットに貼る。終わったら③＝スプシの B-⑤ サブ画像（指示をまた Cursor に貼る）。</p>
    <p id="cSetMainHint" style="font-size:12px;color:#333;margin:0 0 8px;"></p>
    <textarea id="cSetMainTa" style="width:100%;height:200px;font-size:12px;font-family:Consolas,monospace;"></textarea>
    <p style="margin:8px 0 0;">
      <b>②</b>
      <button type="button" onclick="copyCSetMainPrompt()" style="font-weight:bold;font-size:15px;">指示を全選択（コピー用）</button>
    </p>
    <details style="margin:10px 0 0;font-size:12px;color:#555;">
      <summary style="cursor:pointer;">イレギュラー時だけ（普段は開かない）</summary>
      <p style="margin:8px 0 0;">
        <button type="button" id="btnCSetMainRun" onclick="runCSetMain()">この親だけCを実行</button>
        <button type="button" onclick="showCSetMainPrompt(true)">全親一括の指示に切替</button>
        <span> Agent を使わずこのPCで走らせる／全親まとめて指示を出すとき用。数分かかることがあります。</span>
      </p>
      <pre id="cSetMainRunLog" style="display:none;margin:8px 0 0;max-height:180px;overflow:auto;font-size:11px;background:#fff;border:1px solid #ccc;padding:8px;"></pre>
    </details>
  </div>
  <div id="gridOverList"></div>
</div>
<div class="grid">{cards_html}</div>
<div id="lightbox"><div class="lb-wrap" data-rel=""><img id="lbImg" alt=""/><div class="grid-overlay" aria-hidden="true"><div class="grid-cells"></div></div></div></div>
<script>
const TPL = {tpl_json};
const CHILDREN = {child_json};
const AUTO_PLAN = {auto_json};
let lastTa = null;
let gridOn = false;
let export02Done = false;
const gridCache = {{}};
const CHILD_ASSIGN_STATE_KEY = 'b5_main_child_assign|v5|' + {json.dumps(str(compose_dir))};
const GRID_TARGET=16;
const GRID_HARD=24;
const ROI_ROWS=4;
const ROI_COLS=7;
const GRID_UI_VER='p3';
function cellKey(c){{ return c[0]+','+c[1]; }}
function cellSet(arr){{ return new Set((arr||[]).map(cellKey)); }}
function guideOverlapPct(guide, detected){{
  if(!guide||!guide.size) return null;
  let n=0;
  guide.forEach(k=>{{ if(detected.has(k)) n++; }});
  return Math.round(100*n/guide.size);
}}
function badgeClass(n){{
  if(n<=GRID_TARGET) return 'ok';
  if(n<=GRID_HARD) return 'warn';
  return 'bad';
}}
function badgeText(data){{
  const n=Number(data.hitN||0);
  const cn=(data.catchCells||[]).length;
  const sn=(data.salesCells||[]).length;
  const o16=Number(data.over16!=null?data.over16:Math.max(0,n-GRID_TARGET));
  const o24=Number(data.over24!=null?data.over24:Math.max(0,n-GRID_HARD));
  let t;
  if(n<=GRID_TARGET) t='文字マス '+n+'（キャッチ'+cn+'+セール'+sn+'・目標16以内 ✓）';
  else if(n<=GRID_HARD) t='文字マス '+n+'（キャッチ'+cn+'+セール'+sn+'・16超 +'+o16+'）';
  else t='文字マス '+n+'（キャッチ'+cn+'+セール'+sn+'・24超 +'+o24+' 要修正）';
  const tc=data.textCheck;
  if(tc&&!tc.textOk){{
    const miss=[];
    if(tc.catchHead&&!tc.catchOk) miss.push('キャッチ未検出');
    if(tc.salesTalk&&!tc.salesOk) miss.push('セール未検出');
    if(miss.length) t+=' ⚠'+miss.join('／');
  }}
  return t;
}}
function applyGridCardStyle(card, data){{
  if(!card) return;
  card.classList.remove('grid-over16','grid-over24','grid-text-warn');
  const n=Number(data.hitN||0);
  if(n>GRID_HARD) card.classList.add('grid-over24');
  else if(n>GRID_TARGET) card.classList.add('grid-over16');
  if(data.textCheck&&!data.textCheck.textOk) card.classList.add('grid-text-warn');
}}
function paintGridCells(wrap, data){{
  const box=wrap.querySelector('.grid-cells');
  if(!box) return;
  box.innerHTML='';
  const catchSet=cellSet(data.catchCells);
  const salesSet=cellSet(data.salesCells);
  const detected=new Set([...catchSet,...salesSet]);
  const exp=cellSet(data.expectedZone);
  const guide=cellSet(data.adoptedGuide?.cells);
  const useGuide=guide.size>0;
  for(let y=0;y<10;y++)for(let x=0;x<10;x++){{
    const s=document.createElement('span');
    const k=x+','+y;
    s.classList.add('grid-cell');
    if(useGuide){{
      if(guide.has(k)) s.classList.add('roi-focus');
    }}else if(y<ROI_ROWS&&x<ROI_COLS) s.classList.add('roi-focus');
    const isCatch=catchSet.has(k);
    const isSales=salesSet.has(k);
    if(isCatch&&isSales) s.classList.add('hit-both');
    else if(isCatch) s.classList.add('hit-catch');
    else if(isSales) s.classList.add('hit-sales');
    else if(!useGuide&&exp.has(k)) s.classList.add('exp-zone');
    box.appendChild(s);
  }}
  const card=wrap.closest('.card');
  const badge=wrap.parentElement.querySelector('[data-grid-badge]');
  if(badge){{
    const n=Number(data.hitN||0);
    badge.className='grid-badge '+badgeClass(n);
    let t=badgeText(data);
    const g=data.adoptedGuide;
    const ov=guideOverlapPct(guide, detected);
    if(g&&g.id){{
      t+=' ／ 緑='+g.id+(g.source==='compose'?'（compose）':'（最近傍）');
      if(ov!=null) t+=' 緑一致'+ov+'%';
    }}
    badge.textContent=t;
  }}
  applyGridCardStyle(card, data);
}}
function updateGridOverList(rows){{
  const el=document.getElementById('gridOverList');
  if(!el) return;
  if(!gridOn){{
    el.style.display='none';
    el.innerHTML='';
    el.classList.remove('has-over');
    return;
  }}
  const over=rows.filter(r=>r.n>GRID_TARGET).sort((a,b)=>b.n-a.n);
  if(!over.length){{
    el.style.display='block';
    el.classList.remove('has-over');
    el.innerHTML='<b>キャッチ文字マス:</b> 全画像が目標16以内 ✓';
    return;
  }}
  el.style.display='block';
  el.classList.add('has-over');
  const items=over.map(r=>{{
    const cls=r.n>GRID_HARD?' class=\"over24\"':'';
    const tag=r.n>GRID_HARD?('24超 +'+r.o24):('16超 +'+r.o16);
    return '<li'+cls+'><b>'+r.label+'</b> … '+r.n+'マス（'+tag+'）</li>';
  }}).join('');
  el.innerHTML='<b>要確認（キャッチ＋セールストーク）:</b><ul style=\"margin:6px 0 0 18px;padding:0\">'+items+'</ul>';
}}
async function ensureGridHits(wrap){{
  const rel=wrap.dataset.rel;
  if(!rel) return null;
  const card=wrap.closest('.card');
  const catchHead=(card?.querySelector('.catch-head')?.value||'').trim();
  const salesTalk=(card?.querySelector('.sales-talk')?.value||'').trim();
  const cacheKey=rel+'|'+catchHead+'|'+salesTalk;
  if(gridCache[cacheKey]) return gridCache[cacheKey];
  const qs=new URLSearchParams({{rel,catchHead,salesTalk}});
  const res=await fetch('/api/grid-hit?'+qs.toString());
  let j={{ok:false, message:'HTTP '+res.status}};
  try{{ j=await res.json(); }}catch(e){{ j.message='JSON parse error'; }}
  if(!res.ok||!j.ok){{
    j.message=j.message||('HTTP '+res.status);
    return j;
  }}
  gridCache[cacheKey]=j;
  return j;
}}
async function refreshAllGrids(){{
  if(!gridOn){{
    updateGridOverList([]);
    document.querySelectorAll('[data-grid-badge]').forEach(b=>{{ b.style.display='none'; }});
    document.querySelectorAll('.grid-cells').forEach(b=>{{ b.innerHTML=''; }});
    document.querySelectorAll('.card').forEach(c=>c.classList.remove('grid-over16','grid-over24'));
    return;
  }}
  const overRows=[];
  for(const wrap of document.querySelectorAll('.card .img-wrap')){{
    const card=wrap.closest('.card');
    const badge=wrap.parentElement.querySelector('[data-grid-badge]');
    if(badge){{
      badge.style.display='inline-block';
      badge.className='grid-badge warn';
      badge.textContent='計測中…';
    }}
    const j=await ensureGridHits(wrap);
    if(j&&j.ok){{
      paintGridCells(wrap,j);
      const n=Number(j.hitN||0);
      const label=(card&&(card.dataset.adoptKey||card.dataset.theme))||'';
      const name=(card&&card.dataset.name)||'';
      overRows.push({{
        n, o16:Number(j.over16||Math.max(0,n-GRID_TARGET)), o24:Number(j.over24||Math.max(0,n-GRID_HARD)),
        label: label+(name?(' '+name):'')
      }});
    }}else{{
      if(badge){{
        badge.className='grid-badge bad';
        badge.textContent='計測失敗: '+(j&&j.message?j.message:'不明');
      }}
      applyGridCardStyle(card,{{hitN:99}});
    }}
  }}
  updateGridOverList(overRows);
  const lbWrap=document.querySelector('#lightbox .lb-wrap');
  if(lbWrap&&lbWrap.dataset.rel){{
    const j=await ensureGridHits(lbWrap);
    if(j&&j.ok) paintGridCells(lbWrap,j);
  }}
}}
function toggleGrid(){{
  gridOn = !gridOn;
  document.body.classList.toggle('grid-on', gridOn);
  const btn = document.getElementById('btnGrid');
  const hint = document.getElementById('gridHint');
  if(btn) btn.textContent = gridOn ? '100マスガイド OFF' : '100マスガイド';
  if(hint) hint.textContent = gridOn
    ? '（枠=白+緑指示｜塗り=ROI内の実文字｜'+GRID_UI_VER+'）'
    : '（OFF＝ガイド非表示）';
  refreshAllGrids();
}}
if(new URLSearchParams(location.search).get('grid')==='1'){{
  gridOn=true;
  document.body.classList.add('grid-on');
  const btn=document.getElementById('btnGrid');
  const hint=document.getElementById('gridHint');
  if(btn) btn.textContent='100マスガイド OFF';
  if(hint) hint.textContent='（枠=白+緑指示｜塗り=ROI内の実文字｜'+GRID_UI_VER+'）';
  refreshAllGrids();
}}
function applyTemplate(card, t){{
  const ta=card.querySelector('.comment');
  ta.value=(ta.value?ta.value.trim()+'\\n':'')+t.text;
  if(t.regenMode) card.dataset.regenMode=t.regenMode;
  card.querySelector('.need-regen').checked=true;
  lastTa=ta;
  ta.focus();
}}
document.querySelectorAll('textarea.comment').forEach(ta=>{{
  ta.addEventListener('focus',()=>{{ lastTa=ta; }});
}});
function pictureFamilyKey(akey){{
  return String(akey||'').replace(/__r\\d+$/,'');
}}
/** 再作成前後だけ: このカードに子が無いとき、兄弟カードから1子を移す。通常のレ点では呼ばない。 */
function claimFamilyIfEmpty(card){{
  const fam=pictureFamilyKey(card.dataset.adoptKey||'');
  if(!fam) return;
  const mine=[...card.querySelectorAll('.child-ck:checked')];
  if(mine.length) return;
  const moved=[];
  document.querySelectorAll('.card').forEach(other=>{{
    if(other===card) return;
    const ok=other.dataset.adoptKey||'';
    if(pictureFamilyKey(ok)!==fam) return;
    if(ok===(card.dataset.adoptKey||'')) return;
    other.querySelectorAll('.child-ck:checked').forEach(ck=>{{
      moved.push(ck.dataset.sku);
      ck.checked=false;
    }});
    const ad=other.querySelector('.adopt-chk');
    if(ad) ad.checked=false;
  }});
  if(!moved.length) return;
  const ck=card.querySelector('.child-ck[data-sku="'+moved[0]+'"]');
  if(ck) ck.checked=true;
}}
/** このカードで sku を唯一の子にし、他カードの同じ sku は外す（1子=1画像）。 */
function assignChildToCard(card, sku){{
  card.querySelectorAll('.child-ck').forEach(o=>{{
    o.checked=(o.dataset.sku===sku);
  }});
  document.querySelectorAll('.card').forEach(other=>{{
    if(other===card) return;
    other.querySelectorAll('.child-ck').forEach(o=>{{
      if(o.dataset.sku===sku && o.checked){{
        o.checked=false;
        const any=[...other.querySelectorAll('.child-ck')].some(x=>x.checked);
        const ad=other.querySelector('.adopt-chk');
        if(ad&&!any) ad.checked=false;
      }}
    }});
  }});
  const ad=card.querySelector('.adopt-chk');
  if(ad) ad.checked=true;
}}
document.querySelectorAll('.card').forEach(card=>{{
  const bar=card.querySelector('.tpl-btns');
  TPL.forEach((t)=>{{
    const b=document.createElement('button');
    b.type='button'; b.className='tpl-local'; b.textContent=t.label;
    b.onclick=()=>applyTemplate(card, t);
    bar.appendChild(b);
  }});
  const box=card.querySelector('.child-list');
  CHILDREN.forEach(ch=>{{
    const lab=document.createElement('label');
    const ck=document.createElement('input');
    ck.type='checkbox';
    ck.className='child-ck';
    ck.dataset.sku=ch.childSku;
    lab.appendChild(ck);
    lab.appendChild(document.createTextNode(' '+ch.label));
    box.appendChild(lab);
  }});
  if(!CHILDREN.length){{
    box.innerHTML='<span class="warn">レ点子がありません（マスタ読取を確認）</span>';
  }}
  card.querySelectorAll('.child-ck').forEach(ck=>{{
    ck.addEventListener('change',()=>{{
      if(ck.checked){{
        // この画像1子＋全体ユニーク。他画像からの奪取は同じSKUだけ（並びの上下は触らない）
        assignChildToCard(card, ck.dataset.sku);
      }} else {{
        const any=[...card.querySelectorAll('.child-ck')].some(o=>o.checked);
        const ad=card.querySelector('.adopt-chk');
        if(ad&&!any) ad.checked=false;
      }}
      refreshChildStyles();
      persistChildAssignments();
    }});
  }});
  card.querySelector('.adopt-chk').addEventListener('change',()=>{{
    const ad=card.querySelector('.adopt-chk');
    if(ad&&ad.checked) claimFamilyIfEmpty(card);
    if(ad&&!ad.checked){{
      card.querySelectorAll('.child-ck').forEach(ck=>{{ ck.checked=false; }});
    }}
    refreshChildStyles();
    persistChildAssignments();
  }});
  card.querySelectorAll('.catch-head,.sales-talk,.food-tune').forEach(inp=>{{
    inp.addEventListener('input',()=>{{ card.querySelector('.need-regen').checked=true; }});
  }});
  const pendBtn=card.querySelector('.pend-restore');
  if(pendBtn){{
    pendBtn.onclick=()=>{{
      const ch=pendBtn.getAttribute('data-pend-catch')||'';
      const sl=pendBtn.getAttribute('data-pend-sales')||'';
      const chi=card.querySelector('.catch-head');
      const sli=card.querySelector('.sales-talk');
      if(chi&&ch) chi.value=ch;
      if(sli&&sl) sli.value=sl;
      card.querySelector('.need-regen').checked=true;
      pendBtn.remove();
    }};
  }}
  card.querySelector('img.thumb').onclick=(e)=>{{
    e.stopPropagation();
    const src=e.target.src;
    document.getElementById('lbImg').src=src;
    const lbWrap=document.querySelector('#lightbox .lb-wrap');
    const cardWrap=card.querySelector('.img-wrap');
    if(lbWrap&&cardWrap) lbWrap.dataset.rel=cardWrap.dataset.rel||'';
    document.getElementById('lightbox').classList.add('on');
    if(gridOn) refreshAllGrids();
  }};
}});
function takenMap(){{
  const taken={{}};
  document.querySelectorAll('.card').forEach(c=>{{
    const akey=c.dataset.adoptKey||c.dataset.theme;
    c.querySelectorAll('.child-ck:checked').forEach(ck=>{{ taken[ck.dataset.sku]=akey; }});
  }});
  return taken;
}}
function missingChildren(taken){{
  return CHILDREN.filter(ch=>!taken[ch.childSku]);
}}
function refreshChildStyles(){{
  const taken=takenMap();
  document.querySelectorAll('.child-list label').forEach(lab=>{{
    const ck=lab.querySelector('.child-ck');
    if(!ck) return;
    const sku=ck.dataset.sku;
    lab.classList.remove('child-mine','child-taken','child-free');
    if(ck.checked) lab.classList.add('child-mine');
    else if(taken[sku]) lab.classList.add('child-taken');
    else lab.classList.add('child-free');
  }});
  const el=document.getElementById('childRemain');
  if(!el) return;
  if(!CHILDREN.length){{
    el.className='ng';
    el.textContent='レ点子が0件です（マスタの出品CKを確認）';
    return;
  }}
  const miss=missingChildren(taken);
  if(!miss.length){{
    el.className='ok';
    el.textContent=export02Done
      ? 'レ点子: 全 '+CHILDREN.length+' セット割当済み／02コピー済み'
      : 'レ点子: 全 '+CHILDREN.length+' セット割当済み → 次は①';
  }} else {{
    el.className='ng';
    el.innerHTML='未割当 <b style="color:#b00020">'+miss.length+'</b> / '+CHILDREN.length
      +'　→　'+miss.map(m=>m.label).join('、');
  }}
  updateNextStep();
}}
function updateNextStep(){{
  const el=document.getElementById('nextStep');
  const btnC=document.getElementById('btnCSetMain');
  const step2=document.getElementById('step2Label');
  if(!el) return;
  const miss=CHILDREN.length ? missingChildren(takenMap()) : [];
  if(!CHILDREN.length){{
    el.style.borderColor='#c66';
    el.style.background='#fff8f8';
    el.style.color='#b00020';
    el.textContent='いまやること: レ点子がありません（マスタの出品CKを確認）';
    if(btnC) btnC.disabled=true;
    return;
  }}
  if(miss.length){{
    el.style.borderColor='#c66';
    el.style.background='#fff8f8';
    el.style.color='#b00020';
    el.textContent='いまやること: 残す絵に「採用する」と子SKUを付ける（未割当 '+miss.length+'）';
    if(btnC) btnC.disabled=true;
    if(step2) step2.textContent='→ ①のあと ②';
    return;
  }}
  if(!export02Done){{
    el.style.borderColor='#0653c7';
    el.style.background='#eef5ff';
    el.style.color='#063';
    el.innerHTML='いまやること:<br/>'
      +'・絵がおかしければ → そのカードの「構図・演出の要望」に書く → 「再生成する」にレ点 → 下の「途中の作業」から「レ点再生成実行」<br/>'
      +'・残す絵が揃ったら → ①「採用分を02へコピー」';
    if(btnC) btnC.disabled=true;
    if(step2) step2.textContent='→ 絵が揃ってから ① → ②';
    return;
  }}
  el.style.borderColor='#0a7';
  el.style.background='#f3fff6';
  el.style.color='#063';
  el.innerHTML='いまやること:<br/>'
    +'② 上の「指示を全選択」→ <b>Cursor の新しいチャットに貼る</b> → Agent が C（金丸セットMAIN）を実行<br/>'
    +'③ Cが終わったら → スプシメニュー <b>B-⑤ サブ画像</b> を開く → 出てきた指示を Cursor の新しいチャットに貼る<br/>'
    +'<span style="font-weight:normal">（B-⑤済みなら次は B-⑥ Amazon MAIN。下の実行ボタンはイレギュラー時だけ）</span>';
  if(btnC) btnC.disabled=false;
  if(step2) step2.textContent='→ ②のあと ③';
}}
function applyAutoAssign(){{
  const plan=AUTO_PLAN&&AUTO_PLAN.assignments;
  if(!plan||!plan.length) return;
  const byKey={{}};
  plan.forEach(p=>{{ byKey[p.adoptKey||p.themeId]=p; }});
  document.querySelectorAll('.card').forEach(card=>{{
    const akey=card.dataset.adoptKey||card.dataset.theme;
    const p=byKey[akey];
    if(!p) return;
    const adopt=card.querySelector('.adopt-chk');
    if(adopt) adopt.checked=!!p.adopt;
    card.querySelectorAll('.child-ck').forEach(ck=>{{ ck.checked=false; }});
    const skus=Array.isArray(p.childSkus)
      ? p.childSkus
      : (p.childSku ? [p.childSku] : []);
    const one=skus[0];
    if(one){{
      const ck=card.querySelector('.child-ck[data-sku="'+one+'"]');
      if(ck) ck.checked=true;
    }}
  }});
  refreshChildStyles();
  const el=document.getElementById('status');
  if(el){{
    const n=AUTO_PLAN.assignedChildN!=null?AUTO_PLAN.assignedChildN:(AUTO_PLAN.childOrder||[]).length;
    const total=AUTO_PLAN.childN!=null?AUTO_PLAN.childN:(AUTO_PLAN.childOrder||[]).length;
    el.textContent='画像1枚につき子SKUは1つ（自動割当 '+n+'/'+total+'）。見出しが違えばテーマ番号が同じでも別枠。';
  }}
}}
function makeChildAssignState(){{
  const state={{}};
  document.querySelectorAll('.card').forEach(card=>{{
    const akey=card.dataset.adoptKey||card.dataset.theme;
    const skus=[...card.querySelectorAll('.child-ck:checked')].map(ck=>ck.dataset.sku).slice(0,1);
    state[akey]=skus;
  }});
  return state;
}}
function persistChildAssignments(){{
  try{{
    const state=makeChildAssignState();
    localStorage.setItem(CHILD_ASSIGN_STATE_KEY, JSON.stringify(state));
  }}catch(e){{}}
}}
function restoreChildAssignments(){{
  try{{
    const raw=localStorage.getItem(CHILD_ASSIGN_STATE_KEY);
    if(!raw) return;
    const state=JSON.parse(raw||'{{}}');
    // 全クリア
    document.querySelectorAll('.card').forEach(card=>{{
      const adopt=card.querySelector('.adopt-chk');
      if(adopt) adopt.checked=false;
      card.querySelectorAll('.child-ck').forEach(ck=>{{ ck.checked=false; }});
    }});
    // 重複（念のため）を解消: sku→akey は最初に現れたカードを採用
    const taken={{}};
    document.querySelectorAll('.card').forEach(card=>{{
      const akey=card.dataset.adoptKey||card.dataset.theme;
      const skus=Array.isArray(state[akey]) ? state[akey] : [];
      skus.forEach(sku=>{{
        if(!taken[sku]) taken[sku]=akey;
      }});
    }});
    document.querySelectorAll('.card').forEach(card=>{{
      const akey=card.dataset.adoptKey||card.dataset.theme;
      const skus=Array.isArray(state[akey]) ? state[akey] : [];
      let any=false;
      const adopt=card.querySelector('.adopt-chk');
      skus.forEach(sku=>{{
        if(taken[sku]!==akey) return;
        const ck=card.querySelector('.child-ck[data-sku="'+sku+'"]');
        if(ck){{ ck.checked=true; any=true; }}
      }});
      if(adopt) adopt.checked=any;
    }});
    refreshChildStyles();
  }}catch(e){{}}
}}
applyAutoAssign();
restoreChildAssignments();
document.getElementById('lightbox').onclick=()=>document.getElementById('lightbox').classList.remove('on');
document.addEventListener('keydown',(e)=>{{ if(e.key==='Escape') document.getElementById('lightbox').classList.remove('on'); }});
document.querySelectorAll('.tpl-global').forEach(btn=>{{
  btn.onclick=()=>{{
    const t=TPL[Number(btn.dataset.idx)];
    if(!t) return;
    const ta=lastTa||document.querySelector('textarea.comment');
    if(!ta) return;
    const card=ta.closest('.card');
    if(card) applyTemplate(card, t);
    else {{
      ta.value=(ta.value?ta.value.trim()+'\\n':'')+t.text;
      ta.focus();
    }}
  }};
}});
async function collectRegen(){{
  return [...document.querySelectorAll('.card')].map(c=>({{
    name:c.dataset.name,
    themeId:c.dataset.theme,
    adoptKey:c.dataset.adoptKey,
    checked:!!c.querySelector('.need-regen').checked,
    comment:(c.querySelector('.comment').value||'').trim(),
    catchHead:(c.querySelector('.catch-head')?.value||'').trim(),
    salesTalk:(c.querySelector('.sales-talk')?.value||'').trim(),
    foodTuneNote:(c.querySelector('.food-tune')?.value||'').trim(),
    regenMode:c.dataset.regenMode||'full'
  }})).filter(x=>x.checked);
}}
function collectAdopt(){{
  const adopted=[];
  const byChild={{}};
  const used={{}};
  for(const c of document.querySelectorAll('.card')){{
    const childCks=[...c.querySelectorAll('.child-ck:checked')];
    // 子SKUが付いていれば採用扱い（採用チェック漏れ防止）
    if(childCks.length) c.querySelector('.adopt-chk').checked=true;
    if(!c.querySelector('.adopt-chk').checked) continue;
    const akey=c.dataset.adoptKey||c.dataset.theme;
    const name=c.dataset.name;
    adopted.push({{adoptKey:akey, fileName:name, themeId:c.dataset.theme}});
    for(const ck of childCks){{
      const sku=ck.dataset.sku;
      if(used[sku] && used[sku]!==akey){{
        return {{error:'子SKU二重割当: '+sku+'（'+used[sku]+' と '+akey+'）'}};
      }}
      used[sku]=akey;
      byChild[sku]=akey;
    }}
  }}
  const miss=missingChildren(byChild);
  if(miss.length){{
    return {{
      error:'未割当のレ点子があります（'+miss.length+'件）。すべて割当してから02へ進んでください。\\n\\n'
        +miss.map(m=>'・'+m.label).join('\\n')
    }};
  }}
  return {{adopted, byChildSku:byChild}};
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
  const banner=document.getElementById('regenBanner');
  if(banner) banner.classList.add('on');
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
async function export02(){{
  refreshChildStyles();
  const pack=collectAdopt();
  if(pack.error){{
    alert(pack.error);
    document.getElementById('status').textContent=pack.error.split('\\n')[0];
    return;
  }}
  if(!pack.adopted.length){{
    const msg='採用する画像にチェックするか、子SKUを割当してください';
    alert(msg);
    document.getElementById('status').textContent=msg;
    return;
  }}
  document.getElementById('status').textContent='02へコピー中…';
  const res=await fetch('/api/export-02',{{method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify(pack)}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
  if(!j.ok && j.message) alert(j.message);
  if(j.ok){{
    export02Done=true;
    updateNextStep();
    await showCSetMainPrompt();
  }}
}}
async function showCSetMainPrompt(allParents){{
  if(!export02Done && !allParents){{
    alert('先に①「採用分を02へコピー」を押してください');
    updateNextStep();
    return;
  }}
  const scope=allParents?'all':'focus';
  document.getElementById('status').textContent='C楽天量産の指示を生成中（'+scope+'）…';
  const res=await fetch('/api/c-set-main-prompt?scope='+encodeURIComponent(scope));
  const j=await res.json();
  const box=document.getElementById('cSetMainBox');
  const ta=document.getElementById('cSetMainTa');
  const hint=document.getElementById('cSetMainHint');
  const runBtn=document.getElementById('btnCSetMainRun');
  if(!j.ok){{
    box.style.display='block';
    hint.textContent=j.message||'指示の生成に失敗しました';
    ta.value='';
    document.getElementById('status').textContent=j.message||'C指示失敗';
    return;
  }}
  box.style.display='block';
  const scopeLabel=(j.scope==='focus')?'今の親だけ':'全親一括';
  hint.textContent='['+scopeLabel+'] '+(j.summaryLines||[]).slice(0,4).join(' ／ ')
    + ' ／ フォールバック色='+(j.fallbackTextColor||'')+' '+(j.fallbackNameJa||'');
  ta.value=j.prompt||'';
  if(runBtn){{
    runBtn.style.display=(j.scope==='focus')?'inline-block':'none';
    runBtn.disabled=false;
  }}
  document.getElementById('status').textContent=
    (j.scope==='focus')
      ? '②の指示を表示。ふだんは「指示を全選択」→ Cursor へ'
      : 'C指示（全親一括）を表示。ふだんはコピーして Agent へ';
  updateNextStep();
  box.scrollIntoView({{behavior:'smooth',block:'nearest'}});
}}
function copyCSetMainPrompt(){{
  const ta=document.getElementById('cSetMainTa');
  ta.focus();
  ta.select();
  try{{
    document.execCommand('copy');
    document.getElementById('status').textContent='指示をコピーしました。Cursor の新しいチャットに貼ってください';
  }}catch(e){{
    document.getElementById('status').textContent='全文を選択しました。Ctrl+C でコピーしてください';
  }}
}}
async function runCSetMain(){{
  const runBtn=document.getElementById('btnCSetMainRun');
  const logEl=document.getElementById('cSetMainRunLog');
  if(runBtn) runBtn.disabled=true;
  document.getElementById('status').textContent='この親だけC量産を実行中…（数分かかることがあります）';
  if(logEl){{ logEl.style.display='block'; logEl.textContent='実行中…'; }}
  try{{
    const res=await fetch('/api/c-set-main-run',{{method:'POST',headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{scope:'focus'}})}});
    const j=await res.json();
    const lines=(j.logTail||[]).join('\\n');
    if(logEl){{
      logEl.textContent=(j.message||'')+'\\n'+(j.refreshNote||'')+(lines?'\\n---\\n'+lines:'');
    }}
    document.getElementById('status').textContent=j.message||JSON.stringify(j);
    if(!j.ok && j.message) alert(j.message);
  }}catch(e){{
    document.getElementById('status').textContent='C実行失敗: '+e;
    if(logEl) logEl.textContent=String(e);
    alert('C実行失敗: '+e);
  }}finally{{
    if(runBtn) runBtn.disabled=false;
  }}
}}
async function openFolder(){{
  const res=await fetch('/api/open-folder',{{method:'POST'}});
  const j=await res.json();
  document.getElementById('status').textContent=j.message||JSON.stringify(j);
}}
</script>
</body></html>
"""


def _regen_one(
    *,
    compose_dir: Path,
    item: Dict[str, Any],
    meta: Dict[str, Any],
) -> str:
    """1テーマ再 compose。旧ファイルは残し、新名で比較へ追加。"""
    from b5_main_base_compose import compose_rakuten_main_base

    old_name = str(item.get("name") or "").strip()
    if old_name.startswith(_STAGING_PREFIX):
        old_name = old_name[len(_STAGING_PREFIX) :]
    theme_id = str(item.get("themeId") or _theme_from_name(old_name))
    akey_old = adopt_key_from_filename(old_name) if old_name else str(item.get("adoptKey") or "")
    file_head = _file_headline_from_akey(akey_old)
    props = [p for p in (meta.get("proposals") or []) if isinstance(p, dict)]
    # 見出し（ファイル名）で案を固定。同じ themeId の先頭案へ落とさない
    src = _proposal_for_file(props, akey=akey_old, tid=theme_id)
    if not src:
        raise RuntimeError(
            f"proposals に見出し一致の案が無い（themeId={theme_id} akey={akey_old}）。"
            "同テーマ先頭へのフォールバックは禁止。"
        )
    hit = dict(src)
    # 場面は見出しで特定した案のものだけ（別T4の場面を混ぜない）
    scene_locked = str(hit.get("sceneHint") or "")
    orig_catch = str(hit.get("catch") or "")
    catch_head = str(item.get("catchHead") or "").strip()
    sales_talk = str(item.get("salesTalk") or "").strip()
    if not catch_head:
        catch_head = (
            str(hit.get("catchBurn") or "").strip()
            or file_head
            or catch_head_ja(orig_catch)
        )
    if not sales_talk:
        sales_talk = str(hit.get("salesBurn") or hit.get("salesTalk") or "").strip()
    if catch_head or sales_talk:
        hit["catch"] = merge_catch_for_regen(
            catch_head,
            sales_talk,
            original_catch=orig_catch,
        )
        if catch_head:
            hit["catchBurn"] = catch_head
        if sales_talk:
            hit["salesTalk"] = sales_talk
            hit["salesBurn"] = sales_talk
    hit["sceneHint"] = scene_locked
    comment = str(item.get("comment") or "").strip()
    if comment:
        hit["sceneHint"] = (str(hit.get("sceneHint") or "") + "／修正: " + comment[:300]).strip("／")
    regen_mode = str(item.get("regenMode") or "full").strip()
    food_tune = str(item.get("foodTuneNote") or "").strip()
    truth = Path(str(meta.get("packageTruth") or ""))
    if not truth.is_file():
        raise RuntimeError(f"正本なし: {truth}")

    short_name = str(meta.get("shortName") or "item")
    name_base = _regen_name_base(meta)
    combo_parts = [
        p
        for p in (meta.get("comboParts") or meta.get("combo_parts") or [])
        if isinstance(p, dict) and str(p.get("nameBase") or p.get("name") or "").strip()
    ]
    if not old_name:
        old_name = pick_filename(short_name, theme_id, catch_head or str(hit.get("catch") or ""))
    old_path = Path(compose_dir) / "比較" / old_name
    if regen_mode in ("text_only", "food_tune") and old_path.is_file():
        hit["_regenTextOnly"] = True
        hit["_regenReference"] = str(old_path.resolve())
    if food_tune:
        hit["_foodTuneNote"] = food_tune[:200]
    # compose より前に旧版を退避（同名上書き事故を防ぐ）
    hist = Path(compose_dir) / "比較" / "履歴"
    hist.mkdir(parents=True, exist_ok=True)
    if old_path.is_file():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            shutil.copy2(old_path, hist / f"{stamp}_{old_name}")
        except Exception:
            pass
    _purge_regen_staging(compose_dir)

    LOG.info(
        "regen lock akey=%s catch=%s nameBase=%s comboParts=%s",
        akey_old,
        catch_head,
        name_base[:40] if name_base else "",
        len(combo_parts),
    )
    summary = compose_rakuten_main_base(
        jan=str(meta.get("jan") or ""),
        parent_sku=str(meta.get("parentSku") or ""),
        package_truth=truth,
        proposals=[hit],
        text_color=str(meta.get("textColor") or "1"),
        short_name=short_name,
        dry_run=False,
        out_dir=Path(compose_dir),  # 同一フォルダ。比較の本尊名は上書きしない
        name_base=name_base,
        combo_parts=combo_parts,
    )
    if not summary.get("ok"):
        raise RuntimeError("compose 失敗")
    new_pick = None
    result_row: Dict[str, Any] = {}
    for r in summary.get("results") or []:
        if r.get("ok") and r.get("pickFile"):
            new_pick = Path(str(r["pickFile"]))
            result_row = r if isinstance(r, dict) else {}
            break
    if new_pick is None or not new_pick.is_file():
        raise RuntimeError("再生成出力が無い")

    dest_name = _regen_dest_name(
        compose_dir=compose_dir,
        old_name=old_name,
        short_name=short_name,
        theme_id=theme_id,
        catch_head=catch_head or str(hit.get("catchBurn") or ""),
    )
    try:
        _assert_regen_catch_matches(
            expected_head=catch_head or str(hit.get("catchBurn") or ""),
            result_row=result_row,
            old_name=old_name,
            dest_name=dest_name,
        )
    except RuntimeError:
        _purge_regen_staging(compose_dir)
        if new_pick.is_file() and new_pick.name.startswith(_STAGING_PREFIX):
            try:
                new_pick.unlink()
            except Exception:
                pass
        raise
    dest = Path(compose_dir) / "比較" / dest_name
    guard = 0
    while dest.is_file() and guard < 30:
        dest_name = next_regen_filename(dest_name)
        dest = Path(compose_dir) / "比較" / dest_name
        guard += 1
    if dest.is_file():
        raise RuntimeError("再生成ファイル名を確保できない")
    shutil.copy2(new_pick, dest)
    _patch_proposal_card(
        compose_dir,
        theme_id=theme_id,
        akey=akey_old,
        catch=str(hit.get("catch") or ""),
        catch_burn=str(hit.get("catchBurn") or catch_head or ""),
        sales_talk=str(hit.get("salesTalk") or sales_talk or ""),
    )
    # ステージングは必ず消す（本尊・__rN だけ残す。残ると同一絵が二重表示になる）
    _purge_regen_staging(compose_dir)
    if new_pick.is_file() and new_pick.name.startswith(_STAGING_PREFIX):
        try:
            new_pick.unlink()
        except Exception:
            pass
    kept = "残す" if old_path.is_file() else "旧は履歴のみ"
    return f"{theme_id}: 旧={old_name}（{kept}） / 新={dest.name}"


def make_handler(compose_dir: Path, script_dir: Path):
    root = Path(compose_dir).resolve()
    meta_dir = root / "_meta"
    meta_dir.mkdir(exist_ok=True)
    queue_path = meta_dir / "regen_queue.json"

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
                meta = _load_proposals(root)
                rows = list_compare_rows(root)
                sku = str(meta.get("parentSku") or "")
                ch_res = load_checked_children_for_parent(sku) if sku else {
                    "ok": False,
                    "children": [],
                    "error": "parentSku 無し",
                }
                html = build_html(
                    root,
                    rows,
                    meta,
                    ch_res.get("children") or [],
                    "" if ch_res.get("ok") else str(ch_res.get("error") or "レ点子0件"),
                )
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
                ctype = "image/jpeg" if target.suffix.lower() in (".jpg", ".jpeg") else "image/png"
                self._send(200, data, ctype)
                return
            if parsed.path == "/api/grid-hit":
                from urllib.parse import parse_qs

                qs = parse_qs(parsed.query or "")
                rel = unquote((qs.get("rel") or [""])[0]).replace("/", "\\")
                catch_head = unquote((qs.get("catchHead") or [""])[0]).strip()
                sales_talk = unquote((qs.get("salesTalk") or [""])[0]).strip()
                target = (root / rel).resolve()
                root_s = str(root)
                if not (
                    str(target).startswith(root_s + "\\")
                    or str(target).startswith(root_s + "/")
                ) or not target.is_file():
                    self._send(
                        404,
                        json.dumps({"ok": False, "message": "not found"}, ensure_ascii=False).encode(
                            "utf-8"
                        ),
                        "application/json",
                    )
                    return
                if not catch_head and not sales_talk:
                    for row in list_compare_rows(root):
                        row_rel = str(row.get("rel") or "").replace("/", "\\")
                        if row_rel == rel.replace("/", "\\"):
                            catch_head = str(row.get("catchHead") or "").strip()
                            sales_talk = str(row.get("salesTalk") or "").strip()
                            break
                cache_dir = meta_dir / "grid_cache"
                cache_dir.mkdir(exist_ok=True)
                cache_key = target.name.replace("/", "_") + ".catch_v8b.json"
                cache_path = cache_dir / cache_key
                theme_id = _theme_from_name(target.name)
                summary = None
                summary_path = root / "summary.json"
                if summary_path.is_file():
                    try:
                        summary = json.loads(summary_path.read_text(encoding="utf-8"))
                    except (OSError, json.JSONDecodeError):
                        summary = None
                from b5_catch_zone_adopted import (
                    adopted_guide_for_theme,
                    pattern_by_id,
                )

                expected_pre: set[tuple[int, int]] = set()
                measured = None
                if cache_path.is_file() and cache_path.stat().st_mtime >= target.stat().st_mtime:
                    cached = json.loads(cache_path.read_text(encoding="utf-8"))
                    if cached.get("algo") == "catch_v8b":
                        measured = cached
                        expected_pre = {
                            tuple(p)
                            for p in (measured.get("expectedZone") or [])
                            if isinstance(p, (list, tuple)) and len(p) >= 2
                        }
                guide_pre = adopted_guide_for_theme(
                    theme_id=theme_id,
                    summary=summary,
                    expected_cells=expected_pre or None,
                )
                writing_mode = ""
                guide_pattern = None
                guide_cells_arg: Optional[set[tuple[int, int]]] = None
                if guide_pre:
                    writing_mode = str(guide_pre.get("writingMode") or "")
                    pid = str(guide_pre.get("id") or "").strip()
                    if pid:
                        guide_pattern = pattern_by_id(pid)
                    guide_cells_arg = {
                        tuple(p)
                        for p in (guide_pre.get("cells") or [])
                        if isinstance(p, (list, tuple)) and len(p) >= 2
                    }
                if measured is None:
                    from b5_virt_abc import measure_grid20_p3

                    measured = measure_grid20_p3(
                        target,
                        catch_head=catch_head,
                        sales_talk=sales_talk,
                        writing_mode=writing_mode,
                        guide_cells=guide_cells_arg,
                        guide_pattern=guide_pattern,
                    )
                    cache_path.write_text(
                        json.dumps(measured, ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
                else:
                    from b5_virt_abc import _p3_text_check, _p3_text_zones

                    hit_catch = {tuple(p) for p in (measured.get("catchCells") or [])}
                    hit_sales = {tuple(p) for p in (measured.get("salesCells") or [])}
                    catch_cols, catch_rows, sales_cols, sales_rows = _p3_text_zones(
                        catch_head, sales_talk
                    )
                    measured = {
                        **measured,
                        "textCheck": _p3_text_check(
                            hit_catch,
                            hit_sales,
                            catch_head=catch_head,
                            sales_talk=sales_talk,
                            catch_cols=catch_cols,
                            catch_rows=catch_rows,
                            sales_cols=sales_cols,
                            sales_rows=sales_rows,
                            writing_mode=writing_mode,
                        ),
                    }
                expected = {
                    tuple(p)
                    for p in (measured.get("expectedZone") or [])
                    if isinstance(p, (list, tuple)) and len(p) >= 2
                }
                guide = guide_pre or adopted_guide_for_theme(
                    theme_id=theme_id,
                    summary=summary,
                    expected_cells=expected,
                )
                if guide:
                    measured["adoptedGuide"] = guide
                self._send(
                    200,
                    json.dumps({"ok": True, **measured}, ensure_ascii=False).encode("utf-8"),
                    "application/json",
                )
                return
            if parsed.path == "/api/c-set-main-prompt":
                from urllib.parse import parse_qs

                from c_set_main_rakuten_prompt import build_c_set_main_rakuten_prompt

                meta = _load_proposals(root)
                sku = str(meta.get("parentSku") or "").strip()
                qs = parse_qs(parsed.query or "")
                scope = str((qs.get("scope") or ["focus"])[0] or "focus").strip().lower()
                try:
                    out = build_c_set_main_rakuten_prompt(
                        focus_parent_sku=sku,
                        fallback_text_color=str(meta.get("textColor") or "1"),
                        scope=scope,
                    )
                except Exception as e:
                    self._send(
                        500,
                        json.dumps(
                            {"ok": False, "message": f"C指示生成失敗: {e}"},
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return
                self._send(
                    200,
                    json.dumps(out, ensure_ascii=False).encode("utf-8"),
                    "application/json",
                )
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

            if parsed.path == "/api/c-set-main-run":
                from c_set_main_rakuten_prompt import run_c_set_main_rakuten

                meta = _load_proposals(root)
                sku = str(meta.get("parentSku") or "").strip()
                scope = str(payload.get("scope") or "focus").strip().lower()
                try:
                    out = run_c_set_main_rakuten(
                        focus_parent_sku=sku,
                        fallback_text_color=str(meta.get("textColor") or "1"),
                        scope=scope,
                        refresh_master_from_sheets=True,
                    )
                except Exception as e:
                    LOG.exception("c-set-main-run fail")
                    self._send(
                        500,
                        json.dumps(
                            {"ok": False, "message": f"C実行失敗: {e}"},
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return
                self._send(
                    200 if out.get("ok") else 500,
                    json.dumps(out, ensure_ascii=False).encode("utf-8"),
                    "application/json",
                )
                return

            if parsed.path == "/api/save-queue":
                items = payload.get("items") or []
                queue_path.write_text(
                    json.dumps(
                        {
                            "composeDir": str(root),
                            "savedAt": datetime.now(timezone.utc).isoformat(),
                            "items": items,
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
                self._send(
                    200,
                    json.dumps(
                        {"ok": True, "message": f"修正指示を一時保存: {len(items)} 件"},
                        ensure_ascii=False,
                    ).encode("utf-8"),
                    "application/json",
                )
                return

            if parsed.path == "/api/open-folder":
                folder = root / "比較"
                folder.mkdir(parents=True, exist_ok=True)
                try:
                    subprocess.Popen(["explorer", str(folder)])
                    msg = f"比較フォルダを開きました: {folder}"
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
                if not items:
                    self._send(
                        200,
                        json.dumps(
                            {"ok": True, "message": "レ点0件。再生成不要。"},
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return
                meta = _load_proposals(root)
                msgs: List[str] = []
                errors: List[str] = []
                for it in items:
                    try:
                        msgs.append(_regen_one(compose_dir=root, item=it, meta=meta))
                    except Exception as e:
                        LOG.exception("regen fail %s", it)
                        errors.append(f"{it.get('name')}: {e}")
                ok = not errors
                msg = "再生成完了（旧残）: " + ", ".join(msgs) if msgs else "再生成なし"
                if errors:
                    msg += " / 失敗: " + "; ".join(errors)[:400]
                self._send(
                    200 if ok else 500,
                    json.dumps({"ok": ok, "message": msg}, ensure_ascii=False).encode("utf-8"),
                    "application/json",
                )
                return

            if parsed.path == "/api/export-02":
                meta = _load_proposals(root)
                sku = str(meta.get("parentSku") or "").strip()
                if not sku:
                    self._send(
                        400,
                        json.dumps(
                            {"ok": False, "message": "b5_proposals.json に parentSku が無い"},
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return
                adopted_list = payload.get("adopted") or []
                by_child = payload.get("byChildSku") or {}
                if not adopted_list:
                    self._send(
                        400,
                        json.dumps(
                            {"ok": False, "message": "採用が0件"},
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return

                ch_res = load_checked_children_for_parent(sku)
                required = [c["childSku"] for c in (ch_res.get("children") or [])]
                if not required:
                    self._send(
                        400,
                        json.dumps(
                            {
                                "ok": False,
                                "message": "出品CKレ点子が読めません: "
                                + str(ch_res.get("error") or "0件"),
                            },
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return
                adopted_files: Dict[str, str] = {}
                for a in adopted_list:
                    akey = str(a.get("adoptKey") or "").strip()
                    fname = str(a.get("fileName") or "").strip()
                    if not akey or not fname:
                        continue
                    src = root / "比較" / Path(fname).name
                    if not src.is_file():
                        self._send(
                            400,
                            json.dumps(
                                {"ok": False, "message": f"比較に無い: {fname}"},
                                ensure_ascii=False,
                            ).encode("utf-8"),
                            "application/json",
                        )
                        return
                    adopted_files[akey] = fname

                errs = validate_map(
                    by_child={str(k): str(v) for k, v in by_child.items()},
                    adopted={k: export02_filename(sku, adopt_key=k) for k in adopted_files},
                    required_children=required,
                )
                if errs:
                    self._send(
                        400,
                        json.dumps(
                            {"ok": False, "message": "／".join(errs)},
                            ensure_ascii=False,
                        ).encode("utf-8"),
                        "application/json",
                    )
                    return

                dest_dir = Path(DEFAULT_WORK_ROOT) / SUB_RAKUTEN
                dest_dir.mkdir(parents=True, exist_ok=True)
                from PIL import Image

                written = []
                adopted_dest: Dict[str, str] = {}
                for akey, fname in adopted_files.items():
                    src = root / "比較" / Path(fname).name
                    dest = dest_dir / export02_filename(sku, adopt_key=akey, collide=False)
                    if dest.is_file():
                        dest = dest_dir / export02_filename(
                            sku, adopt_key=akey, compose_dir=root.name, collide=True
                        )
                    Image.open(src).convert("RGBA").save(dest, format="PNG")
                    adopted_dest[akey] = dest.name
                    written.append(dest.name)

                # 互換: 採用が1枚だけなら _単体.png も書く
                if len(adopted_files) == 1:
                    only_key = next(iter(adopted_files))
                    src = root / "比較" / Path(adopted_files[only_key]).name
                    legacy = dest_dir / export02_filename(sku, collide=False)
                    if legacy.is_file():
                        legacy = dest_dir / export02_filename(
                            sku, compose_dir=root.name, collide=True
                        )
                    Image.open(src).convert("RGBA").save(legacy, format="PNG")
                    written.append(legacy.name)

                map_file = write_map(
                    dest_dir,
                    parent_sku=sku,
                    by_child={str(k): str(v) for k, v in by_child.items()},
                    adopted=adopted_dest,
                    text_color=str(meta.get("textColor") or "1"),
                    work_root=Path(DEFAULT_WORK_ROOT),
                    extra={
                        "composeDir": root.name,
                        "sourceFiles": adopted_files,
                        "at": datetime.now(timezone.utc).isoformat(),
                    },
                )
                (root / "adopted.json").write_text(
                    json.dumps(
                        {
                            "parentSku": sku,
                            "byChildSku": by_child,
                            "adopted": adopted_dest,
                            "mapPath": str(map_file),
                            "at": datetime.now(timezone.utc).isoformat(),
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
                msg = (
                    f"02へコピー完了（比較は残す）: {', '.join(written)} ＋ {map_file.name}"
                    " → 下の「C楽天量産の指示」へ続けられます"
                )
                self._send(
                    200,
                    json.dumps(
                        {
                            "ok": True,
                            "message": msg,
                            "parentSku": sku,
                            "mapFile": map_file.name,
                        },
                        ensure_ascii=False,
                    ).encode("utf-8"),
                    "application/json",
                )
                return

            self._send(404, b"{}", "application/json")

    return Handler


def serve(compose_dir: Path, port: int, open_browser: bool) -> int:
    script_dir = Path(__file__).resolve().parent
    handler = make_handler(compose_dir, script_dir)
    httpd = ThreadingHTTPServer(("127.0.0.1", port), handler)
    url = f"http://127.0.0.1:{port}/"
    LOG.info("B-⑤ review UI: %s", url)
    print(url)
    if open_browser:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        LOG.info("stopped")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="B-⑤ MAINレビュー／再生成／複数採用→02")
    ap.add_argument("--compose-dir", type=Path, required=True)
    ap.add_argument("--port", type=int, default=8865)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    _setup_log(args.verbose)
    compose = Path(args.compose_dir)
    if not compose.is_dir():
        raise SystemExit(f"compose-dir がありません: {compose}")
    return serve(compose, int(args.port), open_browser=not args.no_browser)


if __name__ == "__main__":
    raise SystemExit(main())
