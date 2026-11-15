# -*- coding: utf-8 -*-
"""未広告ASIN: 名前／親対策のローカル再現（スプシdump上で複数案を実行）"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

OUT = Path(__file__).resolve().parent / "common_ss_view" / "out"
NAME_MAX = 128


def find_tsv(name: str) -> Path:
    hits = list(OUT.rglob(name))
    if not hits:
        raise SystemExit(f"missing {name}")
    return hits[0]


def load(path: Path):
    lines = path.read_text(encoding="utf-8").splitlines()
    h = lines[0].split("\t")
    idx = {n: i for i, n in enumerate(h)}
    rows = []
    for line in lines[1:]:
        cells = line.split("\t")
        rows.append({n: (cells[i] if i < len(cells) else "") for n, i in idx.items()})
    return rows


def norm(s: str) -> str:
    return re.sub(r"\s+", "", (s or "").replace("　", ""))


def sanitize(s: str) -> str:
    t = re.sub(r"\s+", " ", (s or "").strip())
    t = re.sub(r'[\\/:*?"<>|]', "", t)
    return t[:NAME_MAX]


def clip(s: str) -> str:
    return sanitize(s)[:NAME_MAX]


def is_ainori(note: str) -> bool:
    return "相乗り" in (note or "")


def sim_names(rows, mode: str):
    """mode: base_all | ainori_title | title_all"""
    # parent -> first stem choice in sheet order
    used = set()
    parent_name = {}
    parent_collision = {}
    parent_ainori = {}
    for r in rows:
        p = r.get("親ASIN") or ""
        if not p or p in parent_name:
            # still track ainori if any child is
            if is_ainori(r.get("備考", "")):
                parent_ainori[p] = True
            continue
        parent_ainori[p] = is_ainori(r.get("備考", ""))
        base = r.get("商品名ベース") or ""
        title = r.get("商品名") or ""
        if mode == "base_all":
            stem = clip(base or title)
        elif mode == "ainori_title":
            if parent_ainori[p]:
                stem = clip(title or base)
            else:
                stem = clip(base or title)
        else:  # title_all
            stem = clip(title or base)
        n = norm(stem)
        coll = n in used
        if not coll:
            used.add(n)
            used.add(norm("Manual_" + stem))
        parent_name[p] = stem
        parent_collision[p] = coll

    # re-pass: first row of parent may not be ainori if sorted differently —
    # rebuild by grouping: if ANY child ainori, parent uses ainori naming
    if mode == "ainori_title":
        used = set()
        parent_name = {}
        parent_collision = {}
        by_p = defaultdict(list)
        for r in rows:
            by_p[r.get("親ASIN") or ""].append(r)
        for p, kids in by_p.items():
            if not p:
                continue
            ain = any(is_ainori(k.get("備考", "")) for k in kids)
            # use first kid's texts
            k0 = kids[0]
            base = k0.get("商品名ベース") or ""
            title = k0.get("商品名") or ""
            stem = clip(title or base) if ain else clip(base or title)
            n = norm(stem)
            coll = n in used
            if not coll:
                used.add(n)
                used.add(norm("Manual_" + stem))
            parent_name[p] = stem
            parent_collision[p] = coll
            parent_ainori[p] = ain

    yoshi = []
    for p, name in parent_name.items():
        kids = [r for r in rows if r.get("親ASIN") == p]
        if not kids:
            continue
        campish = name + (kids[0].get("商品名ベース") or "")
        if "吉野家" in campish and "缶飯" in campish:
            yoshi.append(
                {
                    "parent": p,
                    "name": name,
                    "collision": parent_collision[p],
                    "ainori": parent_ainori.get(p, False),
                    "would_check": (not parent_collision[p])
                    and (not parent_ainori.get(p, False))
                    and kids[0].get("広告状態") == "なし",
                }
            )
    coll_n = sum(1 for v in parent_collision.values() if v)
    return {
        "mode": mode,
        "collision_parents": coll_n,
        "yoshinoya": yoshi,
        "yoshi_collision": sum(1 for y in yoshi if y["collision"]),
        "yoshi_own_checkable": sum(1 for y in yoshi if y["would_check"]),
    }


def sim_parents(cur, prev, mode: str):
    """mode: none | cache | sibling | cache_then_sibling"""
    pc = {r["ASIN"]: r["親ASIN"] for r in prev if r.get("ASIN")}
    # siblings that currently have Catalog parent
    catalog_kids_by_parent = defaultdict(set)
    for r in cur:
        if "親=Catalog" in (r.get("備考") or "") and r.get("親ASIN") != r.get("ASIN"):
            catalog_kids_by_parent[r["親ASIN"]].add(r["ASIN"])

    fixed = []
    still = []
    for r in cur:
        a = r.get("ASIN") or ""
        p = r.get("親ASIN") or ""
        if not a or p != a:
            continue
        if "親ASINフォールバック=子" not in (r.get("備考") or ""):
            continue
        new_p = ""
        how = ""
        if mode in ("cache", "cache_then_sibling"):
            pp = pc.get(a) or ""
            if pp and pp != a:
                new_p, how = pp, "cache"
        if not new_p and mode in ("sibling", "cache_then_sibling"):
            # if prev parent still has catalog siblings, use that parent
            pp = pc.get(a) or ""
            if pp and pp in catalog_kids_by_parent and catalog_kids_by_parent[pp]:
                new_p, how = pp, "sibling"
        if new_p:
            fixed.append((a, new_p, how))
        else:
            still.append(a)
    return {"mode": mode, "fixed": fixed, "still_self": still}


if __name__ == "__main__":
    cur_path = [p for p in OUT.rglob("未広告ASIN一覧.tsv") if "前回" not in p.name][0]
    prev_path = list(OUT.rglob("未広告ASIN一覧_前回.tsv"))[0]
    cur = load(cur_path)
    prev = load(prev_path)
    print("CUR", cur_path.name, "rows", len(cur))
    print("PREV", prev_path.name, "rows", len(prev))
    print("\n== NAME strategies ==")
    for mode in ("base_all", "ainori_title", "title_all"):
        r = sim_names(cur, mode)
        print(
            f"[{mode}] collision_parents={r['collision_parents']} "
            f"yoshi_collision={r['yoshi_collision']} yoshi_own_checkable={r['yoshi_own_checkable']}"
        )
        for y in r["yoshinoya"]:
            if any(k in y["name"] for k in ("豚丼", "焼鳥", "焼塩", "焼鶏")):
                print(
                    f"  {y['parent']} ainori={y['ainori']} coll={y['collision']} "
                    f"checkable={y['would_check']} | {y['name'][:70]}"
                )
    print("\n== PARENT strategies (self-fallback rows) ==")
    for mode in ("none", "cache", "sibling", "cache_then_sibling"):
        r = sim_parents(cur, prev, mode)
        print(f"[{mode}] fixed={len(r['fixed'])} still_self={len(r['still_self'])}")
        for a, p, how in r["fixed"][:12]:
            print(f"  {a} -> {p} ({how})")
        if len(r["fixed"]) > 12:
            print(f"  ... +{len(r['fixed']) - 12}")
