# -*- coding: utf-8 -*-
"""P2: 貼付ブロックの非候補化＋並べ替え。行削除しない。◎は必ず候補。"""
from __future__ import annotations

import re
from statistics import median

from apply_to_master import normalize_fullwidth_digits, parse_set_count_from_title
from paste_amazon import bag_for_amazon_circle

_FURUSATO = re.compile(r"ふるさと納税|返礼品")
_SELECT = re.compile(r"よりどり|種類が選べ|選べるセット|から選択")
_USED = re.compile(r"中古品|(?<![新])中古")
_EACH_BAG = re.compile(r"各\s*\d+\s*袋")
_UNIT_RATIO = 2.0


def eval_score(ev) -> int | None:
    s = str(ev or "").strip()
    if s == "◎":
        return 100
    m = re.match(r"^(\d+)%?$", s)
    if m:
        return int(m.group(1))
    return None


def has_keepa_attr(row: dict) -> bool:
    if eval_score(row.get("eval")) is not None:
        return True
    if str(row.get("price") or "").strip():
        return True
    if str(row.get("set_count_cell") or "").strip():
        return True
    return False


def classify_row(row: dict, maker: str) -> tuple[str, str]:
    """評価◎は消さない。機械◎でもふるさと・ミックス等は非候補。"""
    title = normalize_fullwidth_digits(row.get("title") or "")
    if str(row.get("asin") or "").strip() and not has_keepa_attr(row):
        return "未属性", "Keepa未取得"
    if _FURUSATO.search(title):
        return "非候補", "ふるさと"
    if _SELECT.search(title):
        return "非候補", "選択式"
    if _USED.search(title):
        return "非候補", "中古"
    if _EACH_BAG.search(title) and parse_set_count_from_title(title)[0] is None:
        return "非候補", "各N袋"
    mk = re.sub(r"[\s\u3000]+", "", str(maker or "")).lower()
    if mk and mk not in re.sub(r"[\s\u3000]+", "", title).lower():
        return "非候補", "メーカー無し"
    return "候補", ""


def _num(v) -> float | None:
    try:
        n = float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None
    return n if n == n and n > 0 else None


def apply_unit_outliers(rows: list[dict]) -> None:
    groups: dict[int, list[dict]] = {}
    for r in rows:
        if r.get("p2") != "候補":
            continue
        bag = bag_for_amazon_circle(str(r.get("title") or ""), r.get("set_count_cell"))
        price = _num(r.get("price"))
        if not bag or not price:
            continue
        r["_bag"] = bag
        r["_unit"] = price / bag
        groups.setdefault(bag, []).append(r)
    for bag, grp in groups.items():
        if len(grp) < 2:
            continue
        for r in grp:
            others = [x["_unit"] for x in grp if x is not r]
            if not others:
                continue
            med = median(others)
            if not med:
                continue
            if r["_unit"] >= med * _UNIT_RATIO:
                r["p2"] = "非候補"
                r["p2_reason"] = "単価2倍"


def sort_block(rows: list[dict], maker: str) -> list[dict]:
    """ASIN無しは末尾。上段候補（袋昇順・同袋は評点降順）→未属性→非候補。"""
    work = []
    blanks = []
    for r in rows:
        rec = dict(r)
        if not str(rec.get("asin") or "").strip():
            blanks.append(rec)
            continue
        kind, reason = classify_row(rec, maker)
        rec["p2"] = kind
        rec["p2_reason"] = reason
        rec["_bag"] = bag_for_amazon_circle(str(rec.get("title") or ""), rec.get("set_count_cell"))
        rec["_score"] = eval_score(rec.get("eval"))
        work.append(rec)
    apply_unit_outliers(work)

    cand = [r for r in work if r.get("p2") == "候補"]
    unk = [r for r in work if r.get("p2") == "未属性"]
    bad = [r for r in work if r.get("p2") == "非候補"]

    def cand_key(r):
        bag = r.get("_bag")
        score = r.get("_score")
        bag_ord = bag if bag else 10**9
        score_ord = -(score if score is not None else -1)
        return (0 if bag else 1, bag_ord, score_ord)

    cand.sort(key=cand_key)
    out = cand + unk + bad + blanks
    for r in out:
        r.pop("_bag", None)
        r.pop("_unit", None)
        r.pop("_score", None)
    return out
