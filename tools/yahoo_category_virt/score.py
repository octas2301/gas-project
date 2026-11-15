# -*- coding: utf-8 -*-
"""Yahooカテゴリ選定（仮想）— 点数制門 → 投票 → ゆるいシリーズFB → AI/Driveガード。

本番 GAS の amazonAiPickYahooCategoryPopularWeightOnly_ と揃える。
- 同商品: 票+3
- 同シリーズ(メーカー一致): 票+2
- 同シリーズ(ゆるい・メーカー空のみ可): 票+1（ヒット少で本戦投入／全滅時の第1FB）
- メーカー衝突はゆるいシリーズにも入れない
- 表記ゆれ: おだし⇔お出し 等
- 食品／ペット等のドメイン辞書は使わない
"""
from __future__ import annotations

import re
from collections import defaultdict
from typing import Any

VOTE_SAME = 3
VOTE_SERIES_MAKER = 2
VOTE_SERIES_LOOSE = 1
SERIES_CORE_MIN_LEN = 4
FEW_HITS_THRESHOLD = 5  # 以下なら本戦でもゆるいシリーズを加点

_NOISE = {
    "セット",
    "まとめ買い",
    "送料無料",
    "おつまみ",
    "ギフト",
    "プレゼント",
    "人気",
    "おすすめ",
}


def _norm(s: str) -> str:
    t = str(s or "").strip().lower()
    t = t.replace("（", "(").replace("）", ")")
    t = re.sub(r"[\s\u3000]+", "", t)
    return t


def _variant_fold(s: str) -> str:
    """最小の表記ゆれ（おだし⇔お出し 等）。"""
    t = _norm(s)
    t = t.replace("お出し", "おだし").replace("出汁", "だし")
    return t


def _clean_token(t: str) -> str:
    t = str(t or "").strip().lower()
    t = re.sub(r"^[（(\[【『「]+|[）)\]】』」・.。,、]+$", "", t)
    t = t.strip("（）()[]【】『』「」・")
    return _variant_fold(t)


def _tokens(s: str) -> set[str]:
    raw = str(s or "")
    out: set[str] = set()
    for p in re.split(r"[\s\u3000/／|｜]+", raw):
        p = _clean_token(p)
        if len(p) >= 2:
            out.add(p)
    for m in re.finditer(r"[一-龥ぁ-んァ-ンa-zA-Z0-9]{2,}", raw):
        p = _clean_token(m.group(0))
        if len(p) >= 2:
            out.add(p)
    return out


def hit_title(h: dict) -> str:
    return str(h.get("title") or h.get("name") or h.get("商品名") or "")


def hit_maker(h: dict) -> str:
    return str(h.get("maker") or h.get("brand") or h.get("メーカー") or "")


def hit_cat_id(h: dict) -> str:
    cid = h.get("genreCategoryId")
    if cid is None:
        cid = h.get("categoryId")
    return str(cid or "").strip()


def hit_path(h: dict) -> str:
    p = h.get("path") or h.get("genreCategoryPath") or ""
    if p:
        return str(p)
    names = []
    parents = h.get("parentGenreCategories") or []
    try:
        parents = sorted(parents, key=lambda x: int((x or {}).get("depth") or 0))
    except Exception:
        pass
    for x in parents:
        n = str((x or {}).get("name") or "").strip()
        if n:
            names.append(n)
    leaf = str(h.get("genreCategoryName") or "").strip()
    if leaf:
        names.append(leaf)
    return ":".join(names)


def _maker_ok(maker_name: str, hit_maker_s: str) -> bool:
    mn = _norm(maker_name)
    hm = _norm(hit_maker_s)
    if not mn or len(mn) < 2 or not hm:
        return False
    return mn in hm or hm in mn


def _maker_conflict(maker_name: str, hit_maker_s: str) -> bool:
    mn = _norm(maker_name)
    hm = _norm(hit_maker_s)
    if not mn or len(mn) < 2 or not hm or len(hm) < 2:
        return False
    return not (mn in hm or hm in mn)


def _excluded_title(title: str) -> bool:
    return bool(
        re.search(
            r"ふるさと納税|返礼品|よりどり|種類が選べ|選べるセット|から選択|中古",
            title,
        )
    )


def series_cores(product_name: str) -> list[str]:
    pt = _tokens(product_name)
    cores = [t for t in pt if len(t) >= SERIES_CORE_MIN_LEN and t not in _NOISE]
    cores.sort(key=lambda x: (-len(x), x))
    return cores


def _title_has_core(title: str, cores: list[str]) -> bool:
    if not cores:
        return False
    title_f = _variant_fold(title)
    ht = _tokens(title)
    for c in cores:
        cf = _variant_fold(c)
        if cf in ht or cf in title_f:
            return True
    return False


def _meaningful_overlap(product_name: str, title: str) -> set[str]:
    pt = _tokens(product_name)
    ht = _tokens(title)
    if not pt or not ht:
        return set()
    inter = pt & ht
    return {t for t in inter if t not in _NOISE and len(t) >= 2}


def is_same_product(h: dict, product_name: str, maker_name: str) -> bool:
    title = hit_title(h)
    if not title or _excluded_title(title):
        return False
    if _maker_conflict(maker_name, hit_maker(h)):
        return False
    meaningful = _meaningful_overlap(product_name, title)
    maker_ok = _maker_ok(maker_name, hit_maker(h))
    if maker_ok and len(meaningful) >= 2:
        return True
    if len(meaningful) >= 3:
        return True
    return False


def is_same_series_maker(h: dict, product_name: str, maker_name: str) -> bool:
    """メーカー一致＋シリーズ核（厚い票）。"""
    title = hit_title(h)
    if not title or _excluded_title(title):
        return False
    if not _maker_ok(maker_name, hit_maker(h)):
        return False
    return _title_has_core(title, series_cores(product_name))


def is_same_series_loose(h: dict, product_name: str, maker_name: str) -> bool:
    """ゆるいシリーズ: 核あり。メーカー衝突は不可。メーカー空は可。"""
    title = hit_title(h)
    if not title or _excluded_title(title):
        return False
    if _maker_conflict(maker_name, hit_maker(h)):
        return False
    return _title_has_core(title, series_cores(product_name))


def classify_hit(
    h: dict,
    product_name: str,
    maker_name: str,
    *,
    allow_loose: bool,
) -> tuple[str, int] | None:
    """('same'|'series_maker'|'series_loose', vote) or None"""
    if is_same_product(h, product_name, maker_name):
        return ("same", VOTE_SAME)
    if is_same_series_maker(h, product_name, maker_name):
        return ("series_maker", VOTE_SERIES_MAKER)
    if allow_loose and is_same_series_loose(h, product_name, maker_name):
        return ("series_loose", VOTE_SERIES_LOOSE)
    return None


def _vote_pool(
    capped: list[dict],
    product_name: str,
    maker_name: str,
    *,
    allow_loose: bool,
) -> tuple[list[dict], int, int, int]:
    pool: list[dict] = []
    n_same = n_sm = n_sl = 0
    for h in capped:
        kind = classify_hit(h, product_name, maker_name, allow_loose=allow_loose)
        if not kind:
            continue
        k, v = kind
        h = dict(h)
        h["_kind"] = k
        h["_vote"] = v
        pool.append(h)
        if k == "same":
            n_same += 1
        elif k == "series_maker":
            n_sm += 1
        else:
            n_sl += 1
    return pool, n_same, n_sm, n_sl


def _pick_from_pool(
    pool: list[dict],
    log: list[str],
    *,
    n_same: int,
    reason: str,
) -> dict[str, Any]:
    votes: dict[str, int] = defaultdict(int)
    weight_by_cat: dict[str, int] = defaultdict(int)
    samples: dict[str, dict] = {}
    for h in pool:
        cid = hit_cat_id(h)
        if not re.match(r"^\d+$", cid):
            continue
        votes[cid] += int(h.get("_vote") or 0)
        weight_by_cat[cid] += int(h.get("_w") or 0)
        prev = samples.get(cid)
        if not prev or (h.get("_kind") == "same") or (
            prev.get("_kind") != "same" and h.get("_kind") == "series_maker"
        ):
            samples[cid] = h
    if not votes:
        log.append("BRANCH_VOTE none")
        return {
            "categoryId": "",
            "path": "",
            "reason": "pool_but_no_genre",
            "weight": 0,
            "needsReview": True,
            "vote": {},
            "sameCount": n_same,
            "seriesCount": 0,
            "seriesLooseCount": 0,
            "log": log,
        }
    max_v = max(votes.values())
    top_cats = [c for c, v in votes.items() if v == max_v]
    log.append(
        "BRANCH_VOTE max=%d cats=%s detail=%s"
        % (max_v, ",".join(top_cats), dict(votes))
    )
    best = max(top_cats, key=lambda c: (weight_by_cat.get(c, 0), -int(c)))
    hit = samples.get(best) or {}
    path = hit_path(hit)
    log.append(
        "BRANCH_WEIGHT_IN_VOTE id=%s w=%s path=%s"
        % (best, weight_by_cat.get(best, 0), path[:60])
    )
    return {
        "categoryId": best,
        "path": path,
        "reason": reason,
        "weight": weight_by_cat.get(best, 0),
        "needsReview": n_same < 1 and max_v < VOTE_SERIES_MAKER,
        "vote": dict(votes),
        "sameCount": n_same,
        "seriesCount": sum(1 for h in pool if h.get("_kind") == "series_maker"),
        "seriesLooseCount": sum(1 for h in pool if h.get("_kind") == "series_loose"),
        "log": log,
    }


def fallback_path_related(product_name: str, path: str) -> bool:
    """商品語（len>=3）がpathに1つでもあれば関連あり。無ければガードNG。"""
    pt = {t for t in _tokens(product_name) if len(t) >= 3 and t not in _NOISE}
    if not pt:
        return True
    path_f = _variant_fold(path or "")
    if not path_f:
        return False
    for t in pt:
        if _variant_fold(t) in path_f:
            return True
    return False


def pick_yahoo_category(
    product_name: str,
    maker_name: str,
    hits: list[dict],
    *,
    top_n: int = 15,
    fallback: dict | None = None,
) -> dict[str, Any]:
    fallback = fallback or {}
    log: list[str] = []
    capped = list(hits or [])[: max(1, int(top_n))]

    for i, h in enumerate(capped):
        h = dict(h)
        h["_rank"] = i
        h["_w"] = top_n - i
        capped[i] = h

    cores = series_cores(product_name)
    log.append("BRANCH_SERIES_CORE cores=%s" % (",".join(cores) or "(none)"))

    # ヒットタイトル先頭（比較・実ラン差分用）
    for i, h in enumerate(capped[:5]):
        log.append(
            "BRANCH_HIT#%d maker=%s title=%s cat=%s"
            % (
                i + 1,
                (hit_maker(h) or "(empty)")[:20],
                hit_title(h)[:40],
                hit_cat_id(h) or "-",
            )
        )

    few = len(capped) <= FEW_HITS_THRESHOLD
    pool, n_same, n_sm, n_sl = _vote_pool(
        capped, product_name, maker_name, allow_loose=few
    )
    log.append(
        "BRANCH_SAME n_same=%d n_series_maker=%d n_series_loose=%d pool=%d/%d few=%s"
        % (n_same, n_sm, n_sl, len(pool), len(capped), few)
    )

    if pool:
        return _pick_from_pool(
            pool,
            log,
            n_same=n_same,
            reason="same_series_weighted_vote_then_weight",
        )

    # 第1FB: ゆるいシリーズ（本戦で未投入のとき＝ヒット多）。メーカー空可・衝突不可
    pool2, n_same2, n_sm2, n_sl2 = _vote_pool(
        capped, product_name, maker_name, allow_loose=True
    )
    log.append(
        "BRANCH_FALLBACK_SERIES_LOOSE n_same=%d n_series_maker=%d n_series_loose=%d pool=%d"
        % (n_same2, n_sm2, n_sl2, len(pool2))
    )
    if pool2:
        return _pick_from_pool(
            pool2,
            log,
            n_same=n_same2,
            reason="fallback_series_loose_vote",
        )

    # 第2/3FB: AI/Drive 等（呼び出し側が渡す）＋関連ガード
    fb_id = str(
        fallback.get("yahooCategoryId")
        or fallback.get("rakutenMappedYahooId")
        or fallback.get("amazonMappedYahooId")
        or ""
    ).strip()
    fb_path = str(fallback.get("path") or fallback.get("yahooPath") or "")
    fb_src = str(fallback.get("source") or "fallback")
    if fb_id and re.match(r"^\d+$", fb_id):
        if not fallback_path_related(product_name, fb_path):
            log.append(
                "BRANCH_FALLBACK_GUARD reject src=%s id=%s path=%s"
                % (fb_src, fb_id, fb_path[:40])
            )
            return {
                "categoryId": "",
                "path": "",
                "reason": "no_same_or_series|fallback_guard|" + fb_src,
                "weight": 0,
                "needsReview": True,
                "vote": {},
                "sameCount": 0,
                "seriesCount": 0,
                "seriesLooseCount": 0,
                "log": log,
            }
        log.append("BRANCH_FALLBACK src=%s id=%s" % (fb_src, fb_id))
        return {
            "categoryId": fb_id,
            "path": fb_path,
            "reason": "no_same_or_series|" + fb_src,
            "weight": 0,
            "needsReview": True,
            "vote": {},
            "sameCount": 0,
            "seriesCount": 0,
            "seriesLooseCount": 0,
            "log": log,
        }

    log.append("BRANCH_FALLBACK src=%s id=(empty)" % fb_src)
    return {
        "categoryId": "",
        "path": "",
        "reason": "no_same_or_series|no_fallback",
        "weight": 0,
        "needsReview": True,
        "vote": {},
        "sameCount": 0,
        "seriesCount": 0,
        "seriesLooseCount": 0,
        "log": log,
    }


def legacy_weight_only(hits: list[dict], top_n: int = 15) -> dict[str, Any]:
    weights: dict[str, int] = defaultdict(int)
    samples: dict[str, dict] = {}
    for i, h in enumerate((hits or [])[:top_n]):
        cid = hit_cat_id(h)
        if not re.match(r"^\d+$", cid):
            continue
        weights[cid] += top_n - i
        if cid not in samples:
            samples[cid] = h
    if not weights:
        return {"categoryId": "", "reason": "legacy_none"}
    best = max(weights.keys(), key=lambda c: (weights[c], -int(c)))
    return {
        "categoryId": best,
        "path": hit_path(samples[best]),
        "reason": "legacy_popular_weight_only",
        "weight": weights[best],
    }
