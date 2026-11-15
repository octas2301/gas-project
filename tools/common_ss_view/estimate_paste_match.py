# -*- coding: utf-8 -*-
"""貼付sheetと品番リストの接合（JAN優先・商品名部分一致は要確認）。

正: docs/org/B_ESTIMATE_INSPECTION_REQUIREMENTS.md
- JAN一致 → 完全一致（アラートなし）
- 商品名部分一致 → 完全一致ではない（人間へアラート）
- 部分一致が複数 → 転記せず ambiguous（アラート）
"""
from __future__ import annotations

import re
import unicodedata
from typing import Any

MATCH_JAN = "jan"
MATCH_NAME = "name_partial"
MATCH_AMBIG = "ambiguous"
MATCH_NONE = "none"

MATCH_LABEL = {
    MATCH_JAN: "接合=JAN一致（完全一致）",
    MATCH_NAME: "接合=商品名部分一致（完全一致ではない・要確認）",
    MATCH_AMBIG: "接合=商品名が複数候補（未転記・要確認）",
    MATCH_NONE: "接合=なし",
}


def jan_norm(v) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    if s.replace(".", "", 1).isdigit():
        try:
            s = str(int(float(s)))
        except Exception:
            pass
    if s.isdigit() and len(s) < 13:
        s = s.zfill(13)
    return s


def num(v):
    try:
        if v is None or v == "":
            return None
        x = float(str(v).replace(",", "").replace("円", "").strip())
        return x if x == x else None
    except Exception:
        return None


def norm_name(s: str) -> str:
    """比較用。空白・記号を落とし小文字化。"""
    t = unicodedata.normalize("NFKC", str(s or "")).lower()
    t = t.replace("\u3000", "").replace(" ", "")
    t = re.sub(r"[【】\[\]（）()「」『』・/／\-_=+*~|]", "", t)
    return t


def paste_rows_from_values(rows: list[list]) -> list[dict[str, Any]]:
    """貼付 A3:G の values（先頭行ヘッダ）→ 卸値あり行のリスト。"""
    out = []
    if not rows:
        return out
    for r in rows[1:]:
        while len(r) < 7:
            r.append("")
        j = jan_norm(r[2])
        p = num(r[5])
        if len(j) < 8 or not p or p <= 0:
            continue
        out.append(
            {
                "jan": j,
                "ex": float(p),
                "cond": str(r[6] or "").strip(),
                "maker": str(r[0] or "").strip(),
                "name": str(r[3] or "").strip(),
            }
        )
    return out


def paste_by_jan(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    return {r["jan"]: r for r in rows}


def _maker_ok(paste_maker: str, list_maker: str, makers_filter: list[str] | None) -> bool:
    pm = str(paste_maker or "").strip()
    lm = str(list_maker or "").strip()
    if makers_filter:
        if lm not in makers_filter:
            return False
        # 貼付メーカーがフィルタ内、またはリストメーカーと一致
        if pm in makers_filter or pm == lm:
            return True
        # 「オリエンタル」⊂「オリエンタル洋行」は不可（逆のみゆるす）
        return any(pm == m or lm == m for m in makers_filter)
    return (not pm) or pm == lm


def _looks_bulk(text: str) -> bool:
    t = str(text or "")
    if "業務用" in t or "お徳用" in t:
        return True
    if re.search(r"(?<!\d)(1kg|１kg|1000g)", t, re.I):
        return True
    return False


def _name_candidates(
    paste_rows: list[dict[str, Any]],
    list_maker: str,
    list_name: str,
    list_title: str,
    makers_filter: list[str] | None,
) -> list[dict[str, Any]]:
    blob = norm_name(f"{list_name} {list_title}")
    if len(blob) < 6:
        return []
    list_bulk = _looks_bulk(list_name) or _looks_bulk(list_title)
    hits = []
    for pr in paste_rows:
        if not _maker_ok(pr["maker"], list_maker, makers_filter):
            continue
        pn = norm_name(pr["name"])
        # 短い貼付名は誤爆しやすい（例: 共通ブランド断片）
        if len(pn) < 6:
            continue
        # 一方向のみ: 貼付の商品名が Amazon 側テキストに含まれる
        if pn not in blob:
            continue
        # 業務用リストへ小売貼付を流さない
        if list_bulk and not _looks_bulk(pr["name"]):
            continue
        hits.append(pr)
    # 長い貼付名を優先（より具体）
    hits.sort(key=lambda x: len(norm_name(x["name"])), reverse=True)
    if not hits:
        return []
    best_len = len(norm_name(hits[0]["name"]))
    top = [h for h in hits if len(norm_name(h["name"])) == best_len]
    uniq = {}
    for h in top:
        uniq[h["jan"]] = h
    return list(uniq.values())


def resolve_paste_match(
    *,
    paste_rows: list[dict[str, Any]],
    paste_map: dict[str, dict[str, Any]],
    jan: str,
    maker: str,
    name: str,
    title: str = "",
    makers_filter: list[str] | None = None,
    allow_name_partial: bool = False,
) -> dict[str, Any]:
    """1行分の接合結果。

    戻り: kind / jan / ex / paste_name / paste_maker / label / alert
    """
    j = jan_norm(jan)
    if j and j in paste_map:
        pr = paste_map[j]
        return {
            "kind": MATCH_JAN,
            "jan": j,
            "ex": pr["ex"],
            "paste_name": pr.get("name") or "",
            "paste_maker": pr.get("maker") or "",
            "cond": pr.get("cond") or "",
            "label": MATCH_LABEL[MATCH_JAN],
            "alert": False,
        }
    if not allow_name_partial:
        return {
            "kind": MATCH_NONE,
            "jan": j,
            "ex": None,
            "paste_name": "",
            "paste_maker": "",
            "cond": "",
            "label": MATCH_LABEL[MATCH_NONE],
            "alert": False,
        }
    cands = _name_candidates(paste_rows, maker, name, title, makers_filter)
    if not cands:
        return {
            "kind": MATCH_NONE,
            "jan": j,
            "ex": None,
            "paste_name": "",
            "paste_maker": "",
            "cond": "",
            "label": MATCH_LABEL[MATCH_NONE],
            "alert": False,
        }
    if len(cands) > 1:
        names = "／".join(c.get("name") or c["jan"] for c in cands[:3])
        return {
            "kind": MATCH_AMBIG,
            "jan": j,
            "ex": None,
            "paste_name": names,
            "paste_maker": cands[0].get("maker") or "",
            "cond": "",
            "label": MATCH_LABEL[MATCH_AMBIG],
            "alert": True,
        }
    pr = cands[0]
    return {
        "kind": MATCH_NAME,
        "jan": pr["jan"],
        "ex": pr["ex"],
        "paste_name": pr.get("name") or "",
        "paste_maker": pr.get("maker") or "",
        "cond": pr.get("cond") or "",
        "label": MATCH_LABEL[MATCH_NAME],
        "alert": True,
    }


def best_match_for_rows(
    rows: list[dict],
    paste_rows: list[dict[str, Any]],
    paste_map: dict[str, dict[str, Any]],
    makers_filter: list[str] | None = None,
    allow_name_partial: bool = True,
) -> dict[str, Any]:
    """商品キー単位（複数ASIN行）の代表接合。JAN一致を最優先。"""
    # 1) いずれかの JAN 一致
    for r in rows:
        m = resolve_paste_match(
            paste_rows=paste_rows,
            paste_map=paste_map,
            jan=r.get("jan") or "",
            maker=r.get("maker") or "",
            name=r.get("name") or "",
            title=r.get("title") or "",
            makers_filter=makers_filter,
            allow_name_partial=False,
        )
        if m["kind"] == MATCH_JAN:
            return m
    # 2) 名前（代表行）
    if not rows:
        return {
            "kind": MATCH_NONE,
            "jan": "",
            "ex": None,
            "paste_name": "",
            "paste_maker": "",
            "cond": "",
            "label": MATCH_LABEL[MATCH_NONE],
            "alert": False,
        }
    r0 = rows[0]
    return resolve_paste_match(
        paste_rows=paste_rows,
        paste_map=paste_map,
        jan=r0.get("jan") or "",
        maker=r0.get("maker") or "",
        name=r0.get("name") or "",
        title=r0.get("title") or "",
        makers_filter=makers_filter,
        allow_name_partial=allow_name_partial,
    )
