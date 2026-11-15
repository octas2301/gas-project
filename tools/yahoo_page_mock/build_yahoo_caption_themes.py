# -*- coding: utf-8 -*-
"""P6c: マスタ「楽天サブ画像1〜10」の入っている枚だけを pool にし、AIが売れそう順の上位3を付ける。

  python tools/yahoo_page_mock/build_yahoo_caption_themes.py
  python tools/yahoo_page_mock/build_yahoo_caption_themes.py --jan 4906283045317 --prep-lib

人手で型・文案を増やさない。正本 docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md §0.1
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[2]
SET_MAIN = ROOT / "tools" / "set_main_image"
sys.path.insert(0, str(SET_MAIN))
sys.path.insert(0, str(ROOT / "tools" / "common_ss_view"))
sys.path.insert(0, str(ROOT / "tools" / "competitor_store"))

from sub_image_intent import TEXT_MODELS, _extract_json_obj  # noqa: E402
from sub_image_lp_themes import THEME_BY_ID  # noqa: E402

OUT_DIR = ROOT / "tools" / "yahoo_page_mock" / "out" / "caption_themes"
JS_FILE = ROOT / "YahooCaptionRich.js"
COMPOSE_ROOT = Path(
    r"G:/マイドライブ/05.画像生成（セットMAIN）/00.テスト出力/sub_image_ai_compose"
)
EXPORT_META_DIR = Path(
    r"G:/マイドライブ/03.楽天・Yahoo!商品登録（CSV一括UL）/02.楽天アップロード画像保存場所"
)
UPLOADED_DIR = EXPORT_META_DIR / "アップロード済み画像"
LISTING_SSID = "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"
MASTER_TAB = "▼商品マスタ(人間作業用)"
_THEME_NUM = re.compile(r"(?:^|[_\-])T(\d+)", re.I)

# 紹介の種に使える安全欄（列があれば使う。カテゴリ専用にしない）
_SAFE_FACT_HEADERS = (
    "商品名ベース",
    "商品名",
    "メーカー名",
    "メーカー名ベース",
    "特徴キーワード",
    "▼マスタ(特徴キーワード)",
    "原料",
    "原材料",
    "原材料(食品)",
)
_FORBIDDEN_FACT_HEADERS = (
    "商品説明文AIから取得",
    "商品説明",
    "▼マスタ(商品説明)",
)

_MASTER_CACHE: Optional[Tuple[List[str], List[List[Any]], int]] = None

DEFAULT_JOBS = [
    # 同一JANでも親が違う商品は別ジョブ（パックキー＝parentSku）
    {
        "jan": "4906283045119",
        "parentSku": "ishi-4906283045119-oya3",
        "slug": "5119_oya3",
        "preferredSku": "ishi-beruodashtpmv-97s12-2",
        "out_stem": "4906283045119_oya3",
        "compose_run": "4906283045119_20260831_122756",
    },
    {
        "jan": "4906283045119",
        "parentSku": "ishi-4906283045119-oya2",
        "slug": "5119_oya2",
        "preferredSku": "ishi-beruodashncdo-64s18-2",
        "out_stem": "4906283045119_oya2",
        "compose_run": "4906283045119_20260831_122756",
    },
    {
        "jan": "4906283045317",
        "parentSku": "ishi-4906283045317-oya2",
        "slug": "maguro",
        "preferredSku": "ishi-4906283045317-32s14-2",
        "out_stem": "4906283045317",
        "compose_run": "4906283045317_20260830_160211",
    },
    {
        "jan": "4906283047410",
        "slug": "buri",
        "preferredSku": "ishi-4906283047410-32s14-2",
        "out_stem": "4906283047410",
        "compose_run": "4906283047410_20260831_123235",
    },
]


def _theme_meta(tid: int) -> Dict[str, Any]:
    m = THEME_BY_ID.get(int(tid)) or {}
    return {
        "themeId": int(tid),
        "themeName": str(m.get("name") or f"T{int(tid):02d}"),
        "themeSlug": str(m.get("nameSlug") or f"t{int(tid):02d}"),
    }


def _load_rows(messages_path: Path) -> tuple[str, List[Dict[str, Any]]]:
    msg = json.loads(messages_path.read_text(encoding="utf-8"))
    name = str(msg.get("productName") or "")
    rows = []
    for r in msg.get("rows") or []:
        try:
            tid = int(r.get("themeId") or 0)
        except (TypeError, ValueError):
            continue
        if not tid:
            continue
        catch = str(r.get("catch") or "").strip()
        overview = str(r.get("overview") or "").strip()
        if not catch and not overview:
            continue
        subs = [str(x).strip() for x in (r.get("subs") or []) if str(x).strip()]
        rows.append(
            {
                "themeId": tid,
                "catch": catch,
                "overview": overview,
                "subs": subs,
                "salesTalk": str(r.get("salesTalk") or "").strip(),
            }
        )
    return name, rows


def _lib_file(slug: str, tid: int, theme_slug: str, sub_index: int) -> str:
    safe = re.sub(r"[^a-z0-9_]+", "", str(theme_slug or f"t{tid:02d}").lower())
    return f"{slug}_s{int(sub_index):02d}_t{tid:02d}_{safe}.jpg"


def _file_md5(path: Path) -> str:
    h = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _ahash(path: Path, size: int = 8) -> str:
    from PIL import Image

    im = Image.open(path).convert("L").resize((size, size))
    pix = list(im.getdata())
    avg = (sum(pix) / len(pix)) if pix else 0
    return "".join("1" if x > avg else "0" for x in pix)


def _tid_from_name(name: str) -> int:
    m = _THEME_NUM.search(Path(name).name)
    return int(m.group(1)) if m else 0


def _parent_tail(parent_sku: str) -> str:
    """ishi-…-oya3 → oya3。空なら oya。"""
    p = str(parent_sku or "").strip()
    if not p:
        return "oya"
    m = re.search(r"(oya\d*)$", p, re.I)
    return m.group(1).lower() if m else "oya"


def _oya_src(jan: str, sub_index: int, parent_sku: str = "") -> Optional[Path]:
    """アップロード済みの親サブ。親違い（oya2/oya3）を先に試し、旧 oya_sub にフォールバック。"""
    n = int(sub_index)
    tails = []
    t = _parent_tail(parent_sku)
    if t:
        tails.append(t)
    if "oya" not in tails:
        tails.append("oya")
    for tail in tails:
        p = UPLOADED_DIR / f"ishi-{jan}-{tail}_sub{n}.jpg"
        if p.is_file():
            return p
    return None


def _theme_file_index(jan: str) -> Tuple[Dict[str, Tuple[int, Path]], Dict[str, Tuple[int, Path]]]:
    """md5 / ahash(距離0用) → (themeId, path)。似た画の誤当てはしない。"""
    files: List[Path] = []
    run = SET_MAIN / "_out" / f"theme_new_run_{jan}"
    if run.is_dir():
        files.extend(run.glob("T*.jpg"))
        files.extend(run.glob("T*.png"))
    out_root = SET_MAIN / "_out"
    if out_root.is_dir():
        for folder in out_root.glob(f"overview_copy_{jan}_*"):
            if folder.is_dir():
                files.extend(folder.glob("T*.jpg"))
                files.extend(folder.glob("T*.png"))
    for folder in COMPOSE_ROOT.glob(str(jan) + "_*"):
        d = folder / "03_openai"
        if d.is_dir():
            files.extend(d.glob("*.jpg"))
    by_md5: Dict[str, Tuple[int, Path]] = {}
    by_ahash: Dict[str, Tuple[int, Path]] = {}
    for p in files:
        tid = _tid_from_name(p.name)
        if not tid:
            continue
        try:
            by_md5[_file_md5(p)] = (tid, p)
            by_ahash[_ahash(p)] = (tid, p)
        except Exception:
            continue
    return by_md5, by_ahash


def _load_master_sub_slots(jan: str, preferred_sku: str) -> tuple[str, List[Dict[str, Any]], str]:
    """マスタ 楽天サブ画像1〜10 の入っている枠。"""
    from client import sheets_service  # noqa: E402

    svc = sheets_service(write=False, interactive=False)
    if not svc:
        print("[YAHOO_CAPTION_P6C] skip jan=" + jan + " reason=no_sheets", file=sys.stderr)
        return "", [], "no_sheets"
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=LISTING_SSID, range="'" + MASTER_TAB.replace("'", "''") + "'")
        .execute()
        .get("values")
        or []
    )
    hdr = None
    hdr_i = 0
    for i, r in enumerate(vals[:20]):
        if any(str(c).startswith("楽天サブ画像") for c in r):
            hdr = r
            hdr_i = i
            break
    if not hdr:
        print("[YAHOO_CAPTION_P6C] skip jan=" + jan + " reason=no_master_header", file=sys.stderr)
        return "", [], "no_master_header"
    try:
        sku_i = hdr.index("子SKU")
        jan_i = next(i for i, h in enumerate(hdr) if str(h) in ("JANコード", "JAN"))
    except (ValueError, StopIteration):
        print("[YAHOO_CAPTION_P6C] skip jan=" + jan + " reason=no_sku_jan_col", file=sys.stderr)
        return "", [], "no_sku_jan_col"
    sub_cols = [(i, int(str(h).replace("楽天サブ画像", ""))) for i, h in enumerate(hdr) if str(h).startswith("楽天サブ画像")]
    sub_cols = [(i, n) for i, n in sub_cols if 1 <= n <= 10]
    candidates: List[Tuple[str, List[Dict[str, Any]]]] = []
    for r in vals[hdr_i + 1 :]:
        if jan_i >= len(r) or str(r[jan_i]).strip() != jan:
            continue
        sku = str(r[sku_i]).strip() if sku_i < len(r) else ""
        slots = []
        for ci, n in sub_cols:
            url = str(r[ci]).strip() if ci < len(r) else ""
            if url:
                slots.append({"subIndex": n, "url": url, "sku": sku})
        if slots:
            candidates.append((sku, slots))
    if not candidates:
        print("[YAHOO_CAPTION_P6C] skip jan=" + jan + " reason=no_master_subs", file=sys.stderr)
        return "", [], "no_master_subs"
    want = str(preferred_sku or "").strip()
    picked = next((c for c in candidates if c[0] == want), None)
    if picked is None:
        picked = max(candidates, key=lambda c: len(c[1]))
    sku, slots = picked
    print(
        "[YAHOO_CAPTION_P6C] master jan=" + jan
        + " sku=" + sku
        + " n=" + str(len(slots))
        + " slots=" + ",".join(str(s["subIndex"]) for s in slots),
        file=sys.stderr,
    )
    return sku, slots, ""


def _load_master_table() -> Tuple[List[str], List[List[Any]], int]:
    """マスタ全表（1実行内キャッシュ）。"""
    global _MASTER_CACHE
    if _MASTER_CACHE is not None:
        return _MASTER_CACHE
    from client import sheets_service  # noqa: E402

    svc = sheets_service(write=False, interactive=False)
    if not svc:
        raise RuntimeError("no_sheets")
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=LISTING_SSID, range="'" + MASTER_TAB.replace("'", "''") + "'")
        .execute()
        .get("values")
        or []
    )
    hdr = None
    hdr_i = 0
    for i, r in enumerate(vals[:30]):
        cells = [str(c) for c in (r or [])]
        if "親SKU" in cells and ("商品名ベース" in cells or "商品名" in cells):
            hdr = cells
            hdr_i = i
            break
        if any(c.startswith("楽天サブ画像") for c in cells) and "親SKU" in cells:
            hdr = cells
            hdr_i = i
            break
    if not hdr:
        raise RuntimeError("no_master_header")
    _MASTER_CACHE = (hdr, vals, hdr_i)
    return _MASTER_CACHE


def _col_map(hdr: List[str]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for i, h in enumerate(hdr):
        key = str(h or "").strip()
        if key and key not in out:
            out[key] = i
    return out


def _cell(row: List[Any], col_map: Dict[str, int], name: str) -> str:
    i = col_map.get(name)
    if i is None or i >= len(row):
        return ""
    s = str(row[i] or "").strip()
    # 括弧まわりの余分な空白を詰める（汎用）
    s = re.sub(r"[（(]\s+", "（", s)
    s = re.sub(r"\s+[）)]", "）", s)
    return s.strip()


def _find_row_by_parent_sku(
    vals: List[List[Any]], hdr_i: int, col_map: Dict[str, int], parent_sku: str
) -> Optional[List[Any]]:
    want = str(parent_sku or "").strip()
    if not want:
        return None
    pi = col_map.get("親SKU")
    if pi is None:
        return None
    for r in vals[hdr_i + 1 :]:
        if pi < len(r) and str(r[pi] or "").strip() == want:
            return r
    return None


def _combo_component_skus(row: List[Any], col_map: Dict[str, int]) -> List[str]:
    """セット構成1〜5の親SKU（空は除く）。商品名の「・」では分けない。"""
    out: List[str] = []
    seen = set()
    for n in range(1, 6):
        sku = _cell(row, col_map, "セット構成%d_親SKU" % n)
        if not sku or sku in seen:
            continue
        seen.add(sku)
        out.append(sku)
    return out


def _safe_fact_dict(row: List[Any], col_map: Dict[str, int]) -> Dict[str, str]:
    """1行から安全欄だけ。禁止列・空・内容量数字は入れない。"""
    forbidden = set(_FORBIDDEN_FACT_HEADERS)
    out: Dict[str, str] = {}
    for name in _SAFE_FACT_HEADERS:
        if name in forbidden:
            continue
        v = _cell(row, col_map, name)
        if not v:
            continue
        if "内容量" in name or "重量" in name:
            continue
        v = _strip_net_qty(v)
        v = re.sub(r"\d+\s*袋", "", v)
        v = re.sub(r"\s{2,}", " ", v).strip(" ／、,")
        if not v:
            continue
        out[name] = v
    return out


def _attach_composition_facts(doc: Dict[str, Any]) -> Dict[str, Any]:
    """
    構成品の商品名→各マスタ行の安全事実をまとめ、doc に載せる。
    構成親SKUが2種以上ならミックス。カテゴリ専用分岐はしない。
    """
    parent_sku = str(doc.get("parentSku") or "").strip()
    empty = {
        "isMix": False,
        "productName": str(doc.get("productName") or "").strip(),
        "componentNames": [],
        "components": [],
        "factsBlob": "",
        "error": "",
    }
    if not parent_sku:
        doc["compositionFacts"] = empty
        return doc
    try:
        hdr, vals, hdr_i = _load_master_table()
    except Exception as e:
        empty["error"] = str(e)
        doc["compositionFacts"] = empty
        print(
            "[YAHOO_CAPTION_P6C] composition skip parent=" + parent_sku + " err=" + str(e),
            file=sys.stderr,
        )
        return doc
    col_map = _col_map(hdr)
    parent_row = _find_row_by_parent_sku(vals, hdr_i, col_map, parent_sku)
    if parent_row is None:
        empty["error"] = "parent_row_missing"
        doc["compositionFacts"] = empty
        return doc
    parent_name = (
        _cell(parent_row, col_map, "商品名ベース")
        or _cell(parent_row, col_map, "商品名")
        or str(doc.get("productName") or "").strip()
    )
    parent_name = _strip_net_qty(parent_name)
    component_skus = _combo_component_skus(parent_row, col_map)
    components: List[Dict[str, Any]] = []
    names: List[str] = []
    seen_names = set()
    for sku in component_skus:
        crow = _find_row_by_parent_sku(vals, hdr_i, col_map, sku)
        if crow is None:
            components.append({"parentSku": sku, "missing": True})
            continue
        facts = _safe_fact_dict(crow, col_map)
        cname = facts.get("商品名ベース") or facts.get("商品名") or sku
        cname = _strip_net_qty(cname)
        facts["商品名ベース"] = cname
        components.append({"parentSku": sku, "missing": False, "facts": facts})
        if cname and cname not in seen_names:
            seen_names.add(cname)
            names.append(cname)
    is_mix = len(names) >= 2 or len(component_skus) >= 2
    # 紐づく構成が1種類しか取れない親は単品扱い（ALIGN §7.16 と同趣旨）
    if len(names) <= 1 and len([c for c in components if not c.get("missing")]) <= 1:
        is_mix = False
    blob_parts: List[str] = []
    for c in components:
        if c.get("missing"):
            continue
        f = c.get("facts") or {}
        bits = [str(f.get("商品名ベース") or "").strip()]
        for k, v in f.items():
            if k in ("商品名ベース", "商品名"):
                continue
            # 原料全文は【商品について】側。紹介の種に載せると二重表示になる
            if "原料" in str(k):
                continue
            vv = str(v or "").strip()
            if vv and vv not in bits:
                bits.append(vv)
        line = "／".join([b for b in bits if b])
        if line:
            blob_parts.append(line)
    facts_blob = " ｜ ".join(blob_parts)
    if is_mix and names:
        product_name = "・".join(names)
    else:
        product_name = parent_name or (names[0] if names else "")
    comp = {
        "isMix": is_mix,
        "productName": product_name,
        "parentName": parent_name,
        "componentNames": names,
        "componentSkus": component_skus,
        "components": components,
        "factsBlob": facts_blob,
        "error": "",
    }
    doc["compositionFacts"] = comp
    if product_name:
        doc["productName"] = product_name
    print(
        "[YAHOO_CAPTION_P6C] composition parent="
        + parent_sku
        + " mix="
        + str(is_mix)
        + " n="
        + str(len(names))
        + " names="
        + ",".join(names[:5]),
        file=sys.stderr,
    )
    return doc


def _resolve_theme_by_catch(src: Path, rows: List[Dict[str, Any]]) -> int:
    """画像内キャッチを messages の catch と突合。文案は作らない。"""
    if not src.is_file() or not rows:
        return 0
    from PIL import Image

    from gemini_image import load_api_key, make_client

    lines = []
    for r in rows:
        lines.append(
            "- themeId=%s catch=%s" % (r["themeId"], r.get("catch") or "")
        )
    prompt = (
        "この商品サブ画像の主キャッチ（画像内の一番大きい見出し）に一致する themeId を1つ選べ。\n"
        "候補の catch と文字が同じ／ほぼ同じものを選ぶ。新しい文は作らない。\n"
        "JSONのみ: {\"themeId\":1}\n\n"
        + "\n".join(lines)
    )
    img = Image.open(src)
    client = make_client(load_api_key())
    for mid in TEXT_MODELS:
        try:
            resp = client.models.generate_content(model=mid, contents=[prompt, img])
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                text = resp.candidates[0].content.parts[0].text
            obj = _extract_json_obj(text) or {}
            tid = int(obj.get("themeId") or 0)
            if tid in {int(r["themeId"]) for r in rows}:
                return tid
        except Exception:
            continue
    return 0


def build_pool(job: Dict[str, Any]) -> Dict[str, Any]:
    jan = job["jan"]
    slug = job["slug"]
    parent_sku = str(job.get("parentSku") or "").strip()
    msg_path = SET_MAIN / "_out" / f"theme_new_run_{jan}" / "messages.json"
    product_name = ""
    rows: List[Dict[str, Any]] = []
    if msg_path.is_file():
        product_name, rows = _load_rows(msg_path)
    else:
        print("[YAHOO_CAPTION_P6C] no_messages jan=" + jan + " path=" + str(msg_path), file=sys.stderr)
    by_tid = {int(r["themeId"]): r for r in rows}

    sku, slots, skip_reason = _load_master_sub_slots(jan, str(job.get("preferredSku") or ""))
    compose_dir = COMPOSE_ROOT / job.get("compose_run", "") / "03_openai"
    by_md5, by_ahash = _theme_file_index(jan) if slots else ({}, {})

    pool = []
    for sl in slots:
        n = int(sl["subIndex"])
        src = _oya_src(jan, n, parent_sku)
        if src is None:
            print(
                "[YAHOO_CAPTION_P6C] skip_no_oya jan=" + jan + " sub=" + str(n),
                file=sys.stderr,
            )
            continue
        tid = 0
        how = ""
        try:
            hx = _file_md5(src)
            if hx in by_md5:
                tid, _p = by_md5[hx]
                how = "md5"
            else:
                ah = _ahash(src)
                if ah in by_ahash:
                    tid, _p = by_ahash[ah]
                    how = "ahash0"
        except Exception:
            tid = 0
        if not tid:
            tid = _resolve_theme_by_catch(src, rows)
            if tid:
                how = "catch"
        if not tid:
            print(
                "[YAHOO_CAPTION_P6C] skip_no_theme jan=" + jan + " sub=" + str(n),
                file=sys.stderr,
            )
            continue
        row = by_tid.get(tid)
        if not row:
            print(
                "[YAHOO_CAPTION_P6C] skip_no_copy jan=" + jan
                + " sub=" + str(n)
                + " themeId=" + str(tid),
                file=sys.stderr,
            )
            continue
        meta = _theme_meta(tid)
        rec = dict(row)
        rec.update(meta)
        rec["srcImage"] = str(src.resolve())
        rec["libFile"] = _lib_file(slug, tid, meta["themeSlug"], n)
        rec["subIndex"] = n
        rec["masterUrl"] = sl.get("url") or ""
        rec["masterSku"] = sku
        rec["themeFrom"] = how
        rec["adopted"] = True
        pool.append(rec)
        print(
            "[YAHOO_CAPTION_P6C] slot jan=" + jan
            + " sub=" + str(n)
            + " themeId=" + str(tid)
            + " via=" + how,
            file=sys.stderr,
        )

    if not pool and not skip_reason:
        skip_reason = "master_subs_unresolved"

    doc = {
        "jan": jan,
        "parentSku": parent_sku,
        "slug": slug,
        "productName": product_name,
        "messagesPath": str(msg_path) if msg_path.is_file() else "",
        "masterSku": sku,
        "masterSlots": [int(s["subIndex"]) for s in slots],
        "adoptedThemeIds": [int(x["themeId"]) for x in pool],
        "poolSkipReason": skip_reason,
        "composeDir": str(compose_dir) if compose_dir.is_dir() else "",
        "pool": pool,
        "ranked": [],
        "rankModel": "",
        "rankError": "",
    }
    return _attach_composition_facts(doc)


def rank_top3(doc: Dict[str, Any]) -> Dict[str, Any]:
    pool = list(doc.get("pool") or [])
    if len(pool) <= 3:
        doc["ranked"] = pool
        doc["rankModel"] = "pool_le3"
        return doc
    from gemini_image import load_api_key, make_client

    lines = []
    for r in pool:
        lines.append(
            f"- subIndex={r.get('subIndex')} themeId={r['themeId']} name={r.get('themeName')} "
            f"catch={r.get('catch')} overview={r.get('overview')}"
        )
    prompt = (
        "Yahooショッピングの商品説明（caption）に載せる。\n"
        "候補はマスタ「楽天サブ画像1〜10」にURLがある当選サブだけ。落選は入っていない。\n"
        "購入につながりやすい順に subIndex を最大3つ。同じ悩みの連打は避ける。\n"
        "虚偽の効能・賞は下げる。JSONのみ: {\"subIndexes\":[1,2,3]}\n\n"
        + "\n".join(lines)
    )
    client = make_client(load_api_key())
    last = ""
    by_sub = {int(r.get("subIndex") or 0): r for r in pool if r.get("subIndex")}
    by_id = {int(r["themeId"]): r for r in pool}
    for mid in TEXT_MODELS:
        try:
            resp = client.models.generate_content(model=mid, contents=prompt)
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                text = resp.candidates[0].content.parts[0].text
            obj = _extract_json_obj(text) or {}
            subs = [int(x) for x in (obj.get("subIndexes") or [])]
            ranked = [by_sub[i] for i in subs if i in by_sub][:3]
            if not ranked:
                ids = [int(x) for x in (obj.get("themeIds") or [])]
                ranked = [by_id[i] for i in ids if i in by_id][:3]
            if not ranked:
                raise RuntimeError("empty ranked")
            doc["ranked"] = ranked
            doc["rankModel"] = mid
            doc["rankError"] = ""
            return doc
        except Exception as e:
            last = str(e)
            continue
    doc["ranked"] = []
    doc["rankError"] = last or "rank_failed"
    return doc


# Yahoo 文字数（JS String.length と同じ。caption 5000 / スマホ用説明 10000）
_CAPTION_LIMIT = 5000
_SP_LIMIT = 10000
_LEN_MARGIN = 80
_REPLACE_RE = re.compile(r"<b>忙しい日のもう一品に。</b>[\s\S]*?(?=<a name=\"sets\"></a>)")
_LIB_BASE = "https://shopping.c.yimg.jp/lib/octas/"


def _decode_js_str(raw: str) -> str:
    return raw.replace(r"\"", '"').replace(r"\n", "\n").replace(r"\t", "\t").replace(r"\\", "\\")


def _template_html(field: str) -> str:
    text = JS_FILE.read_text(encoding="utf-8")
    m = re.search(rf"\b{re.escape(field)}:\s*\"((?:\\.|[^\"\\])*)\"", text)
    if not m:
        return ""
    return _decode_js_str(m.group(1))


def _rest_and_old_block(field: str) -> Tuple[int, int]:
    html = _template_html(field)
    m = _REPLACE_RE.search(html)
    old = m.group(0) if m else ""
    return len(html) - len(old), len(old)


def _esc_html(s: str) -> str:
    return (
        str(s or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _build_theme_block(ranked: List[Dict[str, Any]], img_w: str, texts: List[Dict[str, str]]) -> str:
    parts: List[str] = []
    for i, r in enumerate(ranked[:3]):
        t = texts[i] if i < len(texts) else {}
        catch = _esc_html(r.get("catch") or "")
        ov = _esc_html(t.get("overview") or "")
        sub = _esc_html(t.get("sub") or "")
        fn = str(r.get("libFile") or "").strip()
        if not fn:
            continue
        url = fn if re.match(r"^https?://", fn, re.I) else _LIB_BASE + fn
        parts.append("<b>" + catch + "</b><br>")
        if sub:
            parts.append(sub + "<br>")
        if ov:
            parts.append("<br>" + ov + "<br>")
        parts.append('<br><img src="' + url + '" width="' + img_w + '" alt="' + catch + '"><br><br>')
    return "\n".join(parts)


def _block_budget(field: str, limit: int) -> Tuple[int, int, int]:
    rest, old_block = _rest_and_old_block(field)
    target_total = limit - _LEN_MARGIN
    budget = max(0, target_total - rest)
    # スマホは本番で P5 が通った長さを超えない（it-01118 再発防止）
    if field == "sp_additional" and old_block:
        budget = min(budget, old_block)
    return rest, old_block, budget


def _ensure_orig_copy(r: Dict[str, Any]) -> None:
    if "overviewOrig" not in r:
        r["overviewOrig"] = r.get("overview") or ""
    if "subsOrig" not in r:
        r["subsOrig"] = list(r.get("subs") or [])


def _mechanical_texts(
    ranked: List[Dict[str, Any]],
    img_w: str,
    max_block: int,
    start: List[Dict[str, str]],
) -> List[Dict[str, str]]:
    texts = [dict(x) for x in start]
    while len(texts) < min(3, len(ranked)):
        texts.append({"overview": "", "sub": ""})

    def blen() -> int:
        return len(_build_theme_block(ranked, img_w, texts))

    if blen() <= max_block:
        return texts
    for i in range(len(texts) - 1, -1, -1):
        texts[i]["overview"] = ""
        if blen() <= max_block:
            return texts
    for i in range(len(texts) - 1, -1, -1):
        texts[i]["sub"] = ""
        if blen() <= max_block:
            return texts
    return texts


def _ai_fit_texts(ranked: List[Dict[str, Any]], text_budget: int, surface: str) -> Optional[List[Dict[str, str]]]:
    if text_budget <= 0:
        return [{"overview": "", "sub": ""} for _ in ranked[:3]]
    from gemini_image import load_api_key, make_client

    lines = []
    for r in ranked[:3]:
        _ensure_orig_copy(r)
        lines.append(
            "- themeId=%s catch=%s overview=%s subs=%s"
            % (
                r.get("themeId"),
                r.get("catch") or "",
                r.get("overviewOrig") or "",
                " / ".join(r.get("subsOrig") or []),
            )
        )
    prompt = (
        "Yahooショッピングの商品説明に載せる既存文を、指定の合計文字数以内に短くする。\n"
        "新しい効能・賞・型名・文案は作らない。既存の意味だけ残す。catch は変えない。\n"
        "上の theme ほど文字を残してよい。空でもよい。\n"
        "表面=%s。overview+subs の合計は %d 文字以内。\n"
        "JSONのみ: {\"items\":[{\"themeId\":1,\"overview\":\"...\",\"subs\":[\"...\"]}]}\n\n"
        % (surface, text_budget)
        + "\n".join(lines)
    )
    client = make_client(load_api_key())
    by_id = {int(r["themeId"]): r for r in ranked[:3]}
    for mid in TEXT_MODELS:
        try:
            resp = client.models.generate_content(model=mid, contents=prompt)
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                text = resp.candidates[0].content.parts[0].text
            obj = _extract_json_obj(text) or {}
            items = obj.get("items") or []
            out: List[Dict[str, str]] = []
            used = 0
            for r in ranked[:3]:
                hit = next((x for x in items if int(x.get("themeId") or 0) == int(r["themeId"])), None)
                if not hit:
                    hit = {}
                ov = str(hit.get("overview") or "").strip()
                subs = [str(x).strip() for x in (hit.get("subs") or []) if str(x).strip()]
                used += len(ov) + sum(len(x) for x in subs[:1])
                out.append({"overview": ov, "sub": (subs[0] if subs else "")})
            if used > text_budget + 8:
                continue
            if len(out) == len(ranked[:3]) and all(int(r["themeId"]) in by_id for r in ranked[:3]):
                return out
        except Exception:
            continue
    return None


_ZW_DIGIT = str.maketrans("０１２３４５６７８９", "0123456789")
_NET_QTY_RE = re.compile(
    r"(?:の)?\s*"
    r"[0-9]+(?:\.[0-9]+)?\s*"
    r"(?:kg|ｍｇ|mg|ml|ｍｌ|mL|cc|ｇ|g|リットル|ミリリットル|キログラム|グラム|l|L)"
    r"\s*(?:入り)?"
    r"の?",
    re.I,
)
_NET_PER_BAG_RE = re.compile(
    r"1\s*袋\s*(?:あたり|当たり)?\s*[/／]?\s*[0-9]+(?:\.[0-9]+)?\s*"
    r"(?:kg|mg|ml|mL|cc|ｇ|g|リットル|グラム|l|L)?",
    re.I,
)


def _strip_net_qty(text: str, extras: Optional[List[str]] = None) -> str:
    """紹介から内容量の数字だけ消す。大容量・たっぷり・お徳用は残す。"""
    s = str(text or "")
    if not s:
        return ""
    for ex in extras or []:
        t = str(ex or "").strip()
        if len(t) < 2:
            continue
        s = re.sub(re.escape(t), "", s, flags=re.I)
        s = re.sub(re.escape(t.translate(_ZW_DIGIT).replace("ｇ", "g").replace("Ｇ", "g")), "", s, flags=re.I)
    s2 = s.translate(_ZW_DIGIT).replace("ｇ", "g").replace("Ｇ", "g")
    s2 = _NET_PER_BAG_RE.sub("", s2)
    s2 = _NET_QTY_RE.sub("", s2)
    s2 = re.sub(r"\s{2,}", " ", s2)
    s2 = re.sub(r"[、，]{2,}", "、", s2)
    s2 = re.sub(r"のの", "の", s2)
    s2 = re.sub(r"。+", "。", s2)
    return s2.strip()


def _net_qty_extras(doc: Dict[str, Any], facts: Optional[Dict[str, str]] = None) -> List[str]:
    extras = []
    for k in ("masterNetWeight", "netWeight"):
        v = str((facts or {}).get(k) or doc.get(k) or "").strip()
        if v:
            extras.append(v)
    return extras


def _safe_facts(doc: Dict[str, Any]) -> Dict[str, str]:
    """紹介の種。内容量・商品説明文AIは入れない。構成マスタの安全欄のみ。"""
    if str(doc.get("parentSku") or "").strip():
        # 毎回マスタから取り直す（fit-only でも古くならない）
        _attach_composition_facts(doc)
    comp = doc.get("compositionFacts") or {}
    name = str(comp.get("productName") or doc.get("productName") or "").strip()
    name = _strip_net_qty(name)
    is_mix = bool(comp.get("isMix"))
    names = [str(x).strip() for x in (comp.get("componentNames") or []) if str(x).strip()]
    blob = str(comp.get("factsBlob") or "").strip()
    label = ""
    if not is_mix:
        m = re.search(r"[（(]([^）)]+)[）)]", name)
        label = (m.group(1) if m else "").strip()
    set_seed = ""
    if is_mix and names:
        set_seed = (
            "構成"
            + str(len(names))
            + "種（"
            + "・".join(names)
            + "）。複数種を一度に楽しめるセットである"
        )
    return {
        "productName": name,
        "isMix": "はい" if is_mix else "いいえ",
        "componentNames": "、".join(names),
        "componentCount": str(len(names)) if is_mix and names else "",
        "setAppealSeed": set_seed,
        "factsBlob": blob,
        "labelJa": label,
    }


def _seed_subs(r: Dict[str, Any], doc: Dict[str, Any]) -> List[str]:
    if r.get("subsOrig"):
        return [str(x).strip() for x in r.get("subsOrig") or [] if str(x).strip()]
    sub_i = int(r.get("subIndex") or 0)
    tid = int(r.get("themeId") or 0)
    for p in doc.get("pool") or []:
        if sub_i and int(p.get("subIndex") or 0) == sub_i:
            return [str(x).strip() for x in (p.get("subs") or []) if str(x).strip()]
        if tid and int(p.get("themeId") or 0) == tid and not sub_i:
            return [str(x).strip() for x in (p.get("subs") or []) if str(x).strip()]
    return [str(x).strip() for x in (r.get("subs") or []) if str(x).strip()]


_SET_APPEAL_SENT_RE = re.compile(
    r"[^。]*?(?:種を一度に楽しめるセット|種セットにより|種セットで|種セットです|"
    r"複数種を一度に|食べ比べできる|食べ比べ)[^。]*。?"
)


def _strip_set_appeal_sentences(text: str) -> str:
    s = str(text or "").strip()
    if not s:
        return ""
    out = _SET_APPEAL_SENT_RE.sub("", s)
    out = re.sub(r"。{2,}", "。", out)
    out = re.sub(r"\s{2,}", " ", out).strip(" 。")
    return (out + "。") if out and not out.endswith("。") else out


def _mechanical_sell_overview(
    r: Dict[str, Any],
    facts: Dict[str, str],
    subs: List[str],
    *,
    with_set_appeal: bool = False,
) -> str:
    bits = [str(r.get("salesTalk") or "").strip()] + subs[:2]
    bits = [b for b in bits if b]
    if with_set_appeal and facts.get("isMix") == "はい":
        seed = str(facts.get("setAppealSeed") or "").strip()
        if seed:
            bits.insert(0, seed.split("。")[0].strip())
    extra = str(facts.get("factsBlob") or facts.get("productName") or facts.get("labelJa") or "").strip()
    if extra:
        bits.append(extra.split("｜")[0].strip() if "｜" in extra else extra)
    return _strip_net_qty("。".join(bits[:3]) + ("。" if bits else ""))


def _postprocess_mix_set_appeal_(ranked: List[Dict[str, Any]], facts: Dict[str, str]) -> None:
    """セット訴求は上位1枠だけ残す。2枠目以降から種数アピール文を落とす。"""
    if facts.get("isMix") != "はい" or not ranked:
        return
    for i, r in enumerate(ranked[:3]):
        if i == 0:
            continue
        ov = str(r.get("overview") or "").strip()
        cleaned = _strip_set_appeal_sentences(ov)
        if not cleaned or len(cleaned) < 20:
            cleaned = _mechanical_sell_overview(
                r, facts, list(r.get("subsOrig") or r.get("subs") or []), with_set_appeal=False
            )
        r["overview"] = cleaned
        r["overviewOrig"] = cleaned


def rewrite_overviews_from_sales(doc: Dict[str, Any]) -> Dict[str, Any]:
    """紹介＝キャッチ＋セールストーク／subs＋安全事実。写真用overviewは種にしない。"""
    ranked = list(doc.get("ranked") or [])
    if not ranked:
        return doc
    facts = _safe_facts(doc)
    extras = _net_qty_extras(doc, facts)
    lines = []
    first_tid = int(ranked[0].get("themeId") or 0) if ranked else 0
    for r in ranked[:3]:
        subs = _seed_subs(r, doc)
        r["subsOrig"] = subs
        lines.append(
            "- themeId=%s catch=%s salesTalk=%s subs=%s"
            % (
                r.get("themeId"),
                r.get("catch") or "",
                r.get("salesTalk") or "",
                " / ".join(subs),
            )
        )
    mix_note = ""
    if facts.get("isMix") == "はい":
        mix_note = (
            "ミックス商品である。構成品名が複数あるときは片方だけを主役にしない。\n"
            "セット訴求（setAppealSeed／食べ比べ・○種セット）は themeId=%s の overview にだけ1文入れる。"
            "2番目以降の overview にはセット訴求・種数アピールを書かない。各 catch の買う理由だけ書く。\n"
            "商品名のフル繰り返しは避ける（括弧内の短い呼称でよい）。"
            "日本語として自然な文にする（誤変換・意味不通の複合語は出さない）。\n"
            "構成の安全事実にある語だけ使う。原料の成分表は書かない（【商品について】側）。"
            "カテゴリ専用の型名は作らない。\n"
            % first_tid
        )
    prompt = (
        "Yahooショッピング商品説明の紹介文を、次の種だけから書く。\n"
        "買う理由の文にする。写真の撮り方（クローズアップ・構図・演出）は書かない。\n"
        "新しい効能・賞・型名・事実は作らない。種に無いことは足さない。\n"
        "内容量の数字（g / ml / 個 など単位を問わず）は書かない。大容量・たっぷり・お徳用は種にあれば残してよい。\n"
        "catch と themeId は変えない。overview は2文・80〜140字。\n"
        + mix_note
        + "JSONのみ: {\"items\":[{\"themeId\":1,\"overview\":\"...\"}]}\n\n"
        "事実（安全欄のみ。内容量と商品説明文AIから取得は使わない）: "
        "商品名=%s ミックス=%s 構成種数=%s 構成品名=%s setAppealSeed=%s "
        "構成の安全事実=%s 括弧ラベル=%s\n\n"
        % (
            facts.get("productName") or "",
            facts.get("isMix") or "いいえ",
            facts.get("componentCount") or "",
            facts.get("componentNames") or "",
            facts.get("setAppealSeed") or "",
            facts.get("factsBlob") or "",
            facts.get("labelJa") or "",
        )
        + "\n".join(lines)
    )
    wrote = 0
    last = ""
    try:
        from gemini_image import load_api_key, make_client

        client = make_client(load_api_key())
        for mid in TEXT_MODELS:
            try:
                resp = client.models.generate_content(model=mid, contents=prompt)
                text = getattr(resp, "text", None) or ""
                if not text and getattr(resp, "candidates", None):
                    text = resp.candidates[0].content.parts[0].text
                obj = _extract_json_obj(text) or {}
                items = obj.get("items") or []
                for r in ranked[:3]:
                    hit = next((x for x in items if int(x.get("themeId") or 0) == int(r["themeId"])), None)
                    ov = _strip_net_qty(str((hit or {}).get("overview") or "").strip(), extras)
                    if not ov:
                        continue
                    r["overview"] = ov
                    r["overviewOrig"] = ov
                    wrote += 1
                if wrote:
                    doc["rewriteModel"] = mid
                    break
            except Exception as e:
                last = str(e)
                continue
    except Exception as e:
        last = str(e)
    if not wrote:
        for i, r in enumerate(ranked[:3]):
            ov = _mechanical_sell_overview(
                r, facts, _seed_subs(r, doc), with_set_appeal=(i == 0)
            )
            r["overview"] = ov
            r["overviewOrig"] = ov
        doc["rewriteModel"] = "mechanical_sales"
        doc["rewriteError"] = last
    else:
        for i, r in enumerate(ranked[:3]):
            if not str(r.get("overviewOrig") or "").strip():
                ov = _mechanical_sell_overview(
                    r, facts, _seed_subs(r, doc), with_set_appeal=(i == 0)
                )
                r["overview"] = ov
                r["overviewOrig"] = ov
        doc["rewriteError"] = ""
    _postprocess_mix_set_appeal_(ranked, facts)
    for r in ranked[:3]:
        r["overview"] = _strip_net_qty(r.get("overview") or "", extras)
        r["overviewOrig"] = _strip_net_qty(r.get("overviewOrig") or r.get("overview") or "", extras)
        if r.get("overviewSp") is not None:
            r["overviewSp"] = _strip_net_qty(r.get("overviewSp") or "", extras)
    doc["ranked"] = ranked
    doc["rewriteFacts"] = facts
    print(
        "[YAHOO_CAPTION_P6C] rewrite jan=" + str(doc.get("jan") or "")
        + " model=" + str(doc.get("rewriteModel") or "-")
        + " n=" + str(sum(1 for r in ranked[:3] if r.get("overviewOrig"))),
        file=sys.stderr,
    )
    return doc


def fit_ranked_for_limits(doc: Dict[str, Any]) -> Dict[str, Any]:
    """PC(caption) とスマホ(sp_additional) を同じ仕様で枠内に収める。画・themeId は変えない。"""
    ranked = list(doc.get("ranked") or [])
    if not ranked:
        return doc
    for r in ranked:
        _ensure_orig_copy(r)
    rest_pc, old_pc, budget_pc = _block_budget("caption", _CAPTION_LIMIT)
    rest_sp, old_sp, budget_sp = _block_budget("sp_additional", _SP_LIMIT)
    empty = [{"overview": "", "sub": ""} for _ in ranked[:3]]
    chrome_pc = len(_build_theme_block(ranked, "600", empty))
    chrome_sp = len(_build_theme_block(ranked, "100%", empty))
    text_pc = max(0, budget_pc - chrome_pc)
    text_sp = max(0, budget_sp - chrome_sp)
    extras = _net_qty_extras(doc)

    start_full = []
    for r in ranked[:3]:
        subs = list(r.get("subsOrig") or r.get("subs") or [])
        start_full.append(
            {
                "overview": r.get("overviewOrig") or r.get("overview") or "",
                "sub": (subs[0] if subs else r.get("salesTalk") or ""),
            }
        )

    pc_ai = _ai_fit_texts(ranked, text_pc, "caption") if text_pc < sum(
        len(x["overview"]) + len(x["sub"]) for x in start_full
    ) else start_full
    if pc_ai is None:
        pc_ai = start_full
    pc_txt = _mechanical_texts(ranked, "600", budget_pc, pc_ai)

    need_sp_ai = text_sp < sum(len(x["overview"]) + len(x["sub"]) for x in start_full)
    sp_ai = _ai_fit_texts(ranked, text_sp, "sp_additional") if need_sp_ai else start_full
    if sp_ai is None:
        sp_ai = [{"overview": "", "sub": x["sub"]} for x in start_full] if text_sp < 40 else start_full
    sp_txt = _mechanical_texts(ranked, "100%", budget_sp, sp_ai)

    for i, r in enumerate(ranked[:3]):
        r["overview"] = _strip_net_qty(pc_txt[i]["overview"], extras)
        r["subs"] = [pc_txt[i]["sub"]] if pc_txt[i]["sub"] else []
        sp_ov = _strip_net_qty(sp_txt[i]["overview"], extras)
        if not str(sp_ov or "").strip():
            sp_ov = r["overview"]
        r["overviewSp"] = sp_ov
        r["subsSp"] = [sp_txt[i]["sub"]] if sp_txt[i]["sub"] else []
    doc["ranked"] = ranked
    pc_block = len(_build_theme_block(ranked, "600", pc_txt))
    sp_block = len(_build_theme_block(ranked, "100%", sp_txt))
    doc["fitNote"] = (
        "caption rest=%d old=%d block=%d total=%d/%d; "
        "sp rest=%d old=%d block=%d total=%d/%d"
        % (
            rest_pc, old_pc, pc_block, rest_pc + pc_block, _CAPTION_LIMIT,
            rest_sp, old_sp, sp_block, rest_sp + sp_block, _SP_LIMIT,
        )
    )
    print("[YAHOO_CAPTION_P6C] fit jan=" + str(doc.get("jan") or "") + " " + doc["fitNote"], file=sys.stderr)
    return doc


def _compact_ranked(ranked: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = []
    for r in ranked:
        rec = {
            "themeId": int(r["themeId"]),
            "catch": r.get("catch") or "",
            "overview": r.get("overview") or "",
            "subs": (r.get("subs") or [])[:2],
            "salesTalk": r.get("salesTalk") or "",
            "libFile": r.get("libFile") or "",
        }
        if "overviewSp" in r:
            rec["overviewSp"] = r.get("overviewSp") or ""
        if "subsSp" in r:
            rec["subsSp"] = (r.get("subsSp") or [])[:2]
        out.append(rec)
    return out


def write_js_pack(docs: List[Dict[str, Any]]) -> None:
    """親SKUキーを優先。同一JANに親が複数あるとき JANキーは書かない（誤当て防止）。"""
    pack: Dict[str, Any] = {}
    by_key: Dict[str, Dict[str, Any]] = {}
    for p in sorted(OUT_DIR.glob("*.json")):
        d = json.loads(p.read_text(encoding="utf-8"))
        key = str(d.get("parentSku") or d.get("out_stem") or d.get("jan") or p.stem).strip()
        if key:
            by_key[key] = d
    for d in docs:
        key = str(d.get("parentSku") or d.get("jan") or "").strip()
        if key:
            by_key[key] = d
    jan_counts: Dict[str, int] = {}
    for d in by_key.values():
        j = str(d.get("jan") or "").strip()
        if j:
            jan_counts[j] = jan_counts.get(j, 0) + 1
    for d in by_key.values():
        ranked = _compact_ranked(d.get("ranked") or [])
        if not ranked:
            continue
        rec = {
            "slug": d.get("slug") or "",
            "parentSku": d.get("parentSku") or "",
            "jan": d.get("jan") or "",
            "ranked": ranked,
        }
        parent = str(d.get("parentSku") or "").strip()
        jan = str(d.get("jan") or "").strip()
        if parent:
            pack[parent] = rec
        if rec["slug"]:
            pack[str(rec["slug"])] = rec
        # 親が1つだけのJANだけ互換キーを残す
        if jan and jan_counts.get(jan, 0) == 1:
            pack[jan] = rec
    blob = "var YAHOO_CAPTION_P6C_PACK_ = " + json.dumps(pack, ensure_ascii=False, indent=2) + ";"
    text = JS_FILE.read_text(encoding="utf-8")
    begin = "// P6C_PACK_BEGIN"
    end = "// P6C_PACK_END"
    if begin not in text or end not in text:
        raise SystemExit("YahooCaptionRich.js に P6C_PACK マーカーがありません")
    pre, rest = text.split(begin, 1)
    _, post = rest.split(end, 1)
    JS_FILE.write_text(pre + begin + "\n" + blob + "\n" + end + post, encoding="utf-8", newline="\n")


def prep_lib(docs: List[Dict[str, Any]]) -> None:
    src_dir = ROOT / "assets" / "yahoo_lib_src"
    src_dir.mkdir(parents=True, exist_ok=True)
    n = 0
    for d in docs:
        for r in d.get("ranked") or []:
            src = Path(r.get("srcImage") or "")
            dest = src_dir / str(r.get("libFile") or "")
            if not src.is_file() or not dest.name:
                continue
            shutil.copy2(src, dest)
            n += 1
            print("copied", dest.name)
    if not n:
        print("prep-lib: コピー0")
        return
    prep = ROOT / "tools" / "yahoo_page_mock" / "prep_lib_images.py"
    subprocess.run(
        [
            sys.executable,
            str(prep),
            "--src",
            str(src_dir),
            "--out",
            str(ROOT / "tools" / "yahoo_page_mock" / "out" / "lib_ready"),
        ],
        cwd=str(ROOT),
        check=False,
    )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--jan", action="append", default=[])
    ap.add_argument("--skip-rank", action="store_true")
    ap.add_argument("--fit-only", action="store_true", help="既存JSONの上位3だけ文字数合わせ（pool再取得しない）")
    ap.add_argument("--prep-lib", action="store_true")
    ap.add_argument("--no-js", action="store_true")
    args = ap.parse_args()
    jobs = DEFAULT_JOBS
    if args.jan:
        want = set(args.jan)
        jobs = [j for j in DEFAULT_JOBS if j["jan"] in want]
        if not jobs:
            raise SystemExit("未知のJAN")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if args.fit_only:
        docs = []
        want_jan = set(j["jan"] for j in jobs)
        for p in sorted(OUT_DIR.glob("*.json")):
            d = json.loads(p.read_text(encoding="utf-8"))
            if str(d.get("jan") or p.stem) not in want_jan and not any(
                str(d.get("jan") or "").startswith(j) or p.stem.startswith(j) for j in want_jan
            ):
                continue
            if str(d.get("jan") or "") not in want_jan and p.stem.split("_")[0] not in want_jan:
                continue
            d = rewrite_overviews_from_sales(d)
            d = fit_ranked_for_limits(d)
            p.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
            print("fit", p.name, d.get("fitNote") or "")
            docs.append(d)
        if not args.no_js:
            write_js_pack(docs)
            print("updated YahooCaptionRich.js P6C pack")
        return 0
    docs = []
    for job in jobs:
        doc = build_pool(job)
        if args.skip_rank:
            doc["ranked"] = (doc.get("pool") or [])[:3]
            doc["rankModel"] = "skip_rank_first3"
        else:
            try:
                doc = rank_top3(doc)
            except SystemExit as e:
                doc["rankError"] = str(e)
                print("RANK_SKIP", job["jan"], e, file=sys.stderr)
        stem = str(job.get("out_stem") or job["jan"]).strip()
        outp = OUT_DIR / f"{stem}.json"
        slim = dict(doc)
        slim["out_stem"] = stem
        # src path は残す。巨大化しない
        outp.write_text(json.dumps(slim, ensure_ascii=False, indent=2), encoding="utf-8")
        print(
            "wrote",
            outp.name,
            "parent",
            job.get("parentSku") or "-",
            "pool",
            len(doc.get("pool") or []),
            "ranked",
            len(doc.get("ranked") or []),
            "model",
            doc.get("rankModel") or "-",
            "err",
            (doc.get("rankError") or "")[:80],
        )
        if doc.get("ranked"):
            doc = rewrite_overviews_from_sales(doc)
            doc = fit_ranked_for_limits(doc)
            doc["out_stem"] = stem
            outp.write_text(json.dumps(dict(doc), ensure_ascii=False, indent=2), encoding="utf-8")
        docs.append(doc)
    # 旧・JAN単体JSONは親分割後に誤ってパックへ載るので除く
    for job in jobs:
        if job.get("out_stem") and job.get("out_stem") != job["jan"]:
            legacy = OUT_DIR / f"{job['jan']}.json"
            if legacy.is_file():
                legacy.unlink()
                print("removed legacy", legacy.name)
    if not args.no_js:
        write_js_pack(docs)
        print("updated YahooCaptionRich.js P6C pack")
    if args.prep_lib:
        prep_lib(docs)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
