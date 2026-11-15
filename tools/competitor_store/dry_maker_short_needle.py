# -*- coding: utf-8 -*-
"""ブランド型メーカーの短語 needle 追加リスク評価（dry。シート非書）。

Catalog page0 を brand entry で取得し、needle セットごとの
  title_hit / 真陽性 / 偽陽性（別ブランド混入）
を比較する。
"""
from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SMOKE = ROOT.parent / "spapi_smoke"
PATH3 = ROOT.parent / "purchase_research_path3"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(SMOKE))
sys.path.insert(0, str(PATH3))

from client import sheets_service  # noqa: E402
from dry_maker_sum import as_dicts, read_all  # noqa: E402
from maker_alias import (  # noqa: E402
    JITTAI_BRAND,
    MAKER_ALIASES,
    jitta_for_maker,
)
from schema import SHEET_KEEPA_FULL  # noqa: E402
from search_catalog_keywords import asins_from_body, search_catalog_items  # noqa: E402
from spapi_smoke import _load_config, request_lwa_access_token, resolve_cred  # noqa: E402

RESEARCH_SS = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
T_CAND = "①候補"

# 入口語ヒット列の汚れ再現（現行 GAS prSplitNeedles_ 相当）
PAGE0_SUFFIX_RE = re.compile(r"\s*page0=.*$", re.I)

BRAND_CASES = [
    {
        "parent": "サトウ食品",
        "entry": "Sato",
        "brand_words": ["Sato", "サトウ食品", "サトウの"],
        "short": "サトウ",
        "hit_polluted": "Sato page0=20満杯(2026-08-17)",
    },
    {
        "parent": "味の素",
        "entry": "AGF",
        "brand_words": ["AGF", "エージーエフ", "味の素"],
        "short": "味の素",
        "hit_polluted": "AGF page0=20満杯(2026-08-17)",
    },
    {
        "parent": "エスビー食品",
        "entry": "S&B",
        "brand_words": ["S&B", "エスビー", "エスビー食品"],
        "short": "エスビー",
        "hit_polluted": "S&B/エスビー食品 page0=20満杯(2026-08-17)",
    },
    {
        "parent": "永谷園",
        "entry": "Nagatanien",
        "brand_words": ["永谷園", "Nagatanien"],
        "short": "永谷",
        "hit_polluted": "Nagatanien page0=20満杯(2026-08-17)",
    },
    {
        "parent": "いし本食品工業",
        "entry": "いしもと",
        "brand_words": ["いしもと", "いし本食品工業"],
        "short": "いし本",
        "hit_polluted": "いしもと/いし本食品工業 page0=20満杯(2026-08-17)",
    },
]


def split_needles(parent: str, entry_hit: str) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()

    def push(x: str) -> None:
        t = str(x or "").strip()
        if not t or t in seen:
            return
        seen.add(t)
        out.append(t)

    push(parent)
    for part in re.split(r"[|,、／/]", str(entry_hit or "")):
        push(part)
    return out


def short_from_parent(parent: str) -> str:
    t = str(parent or "").strip()
    t = re.sub(r"(食品|工業|株式会社|有限会社|ホールディングス).*$", "", t)
    return t.strip()


def is_true_positive(blob: str, parent: str, brand_words: list[str]) -> bool:
    b = blob or ""
    if parent and parent in b:
        return True
    for w in brand_words:
        if w and w in b:
            return True
    return False


def eval_needles(
    rows: list[dict],
    needles: list[str],
    parent: str,
    brand_words: list[str],
) -> dict:
    hits: list[dict] = []
    for r in rows:
        blob = str(r.get("title") or "") + " " + str(r.get("brand") or "")
        if any(n and n in blob for n in needles):
            hits.append(r)
    tp = [r for r in hits if is_true_positive(
        str(r.get("title") or "") + " " + str(r.get("brand") or ""),
        parent,
        brand_words,
    )]
    fp = [r for r in hits if r not in tp]
    return {
        "hit": len(hits),
        "tp": len(tp),
        "fp": len(fp),
        "fp_samples": fp[:3],
    }


def warehouse_sato_collision(full: list[dict]) -> None:
    """倉庫ブランドに サトウ 系が他社混入していないか（再GETなし）。"""
    makers: dict[str, set[str]] = {}
    for r in full:
        brand = str(r.get("ブランド") or "").strip()
        title = str(r.get("商品名") or "").strip()
        blob = title + " " + brand
        if "サトウ" not in blob:
            continue
        parent = str(r.get("製造者") or r.get("メーカー") or "").strip() or "(空)"
        makers.setdefault(parent, set()).add(brand or "(空brand)")
    print("=== 倉庫: タイトル/ブランドに「サトウ」含有（製造者別）===")
    for parent, brands in sorted(makers.items(), key=lambda x: -len(x[1])):
        print("  製造者=%s brands=%s" % (parent, sorted(brands)[:5]))
    print("  計 %d 製造者" % len(makers))


def cand_sato_collision(cand: list[dict]) -> None:
    rows = [r for r in cand if "サトウ" in str(r.get("商品名") or "")]
    parents = sorted({str(r.get("メーカー") or "") for r in rows})
    print("=== ①候補: 商品名に「サトウ」含有 ===")
    print("  行数=%d メーカー=%s" % (len(rows), parents))


def spapi_access():
    cfg = _load_config(SMOKE / "config.local.json")
    access = str(
        request_lwa_access_token(
            resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID"),
            resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET"),
            resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN"),
        ).get("access_token")
        or ""
    )
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    mp = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    return access, endpoint, mp, "OctasShortNeedle/1.0"


def catalog_page0(entry: str, access, endpoint, mp, ua) -> tuple[list[dict], int, bool]:
    resp = search_catalog_items(endpoint, access, mp, ua, entry, page_size=20, page_token="")
    body = resp.json() if resp.content else {}
    rows = asins_from_body(body)
    next_tok = bool((body.get("pagination") or {}).get("nextToken"))
    return rows, resp.status_code, next_tok


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--maker", default="", help="1社だけ（例: サトウ食品）")
    ap.add_argument("--skip-api", action="store_true", help="倉庫/①候補のみ")
    args = ap.parse_args()

    svc = sheets_service(write=False, interactive=True)
    if not svc:
        print("no_creds")
        return 1
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    warehouse_sato_collision(full)
    cand_sato_collision(cand)

    cases = BRAND_CASES
    if args.maker:
        cases = [c for c in BRAND_CASES if c["parent"] == args.maker]
        if not cases:
            print("unknown maker: %s" % args.maker)
            return 2

    if args.skip_api:
        print("skip_api_ok")
        return 0

    access, endpoint, mp, ua = spapi_access()
    print("\n=== Catalog page0 × needle セット比較（dry）===")
    for case in cases:
        parent = case["parent"]
        entry = case["entry"]
        short = case.get("short") or short_from_parent(parent)
        polluted = case["hit_polluted"]
        needles_bad = split_needles(parent, polluted)
        if jitta_for_maker(parent) != JITTAI_BRAND:
            print("SKIP %s 実態=%s" % (parent, jitta_for_maker(parent)))
            continue

        rows, code, next_tok = catalog_page0(entry, access, endpoint, mp, ua)
        print(
            "\n--- %s entry=%s http=%s page0=%d next=%s ---"
            % (parent, entry, code, len(rows), "Y" if next_tok else "N")
        )
        strategies = {
            "A_現行汚れ(entry=needles[1])": (
                [needles_bad[1]] if len(needles_bad) > 1 else [entry],
                needles_bad,
            ),
            "B_正entry+親名only": ([entry], [parent]),
            "C_正entry+親+brand": ([entry], [parent] + case["brand_words"]),
            "D_正entry+親+brand+短語": ([entry], [parent] + case["brand_words"] + [short]),
            "E_正entry+短語only": ([entry], [short]),
        }
        for label, (_e, needles) in strategies.items():
            uniq = []
            seen = set()
            for n in needles:
                n = PAGE0_SUFFIX_RE.sub("", str(n).strip())
                if n and n not in seen:
                    seen.add(n)
                    uniq.append(n)
            ev = eval_needles(rows, uniq, parent, case["brand_words"])
            risk = (ev["fp"] / ev["hit"] * 100) if ev["hit"] else 0.0
            print(
                "  %s needles=%s → hit=%d tp=%d fp=%d fp率=%.0f%%"
                % (label, uniq, ev["hit"], ev["tp"], ev["fp"], risk)
            )
            for fp in ev["fp_samples"]:
                t = str(fp.get("title") or "")[:70]
                b = str(fp.get("brand") or "")
                print("    FP例 asin=%s brand=%s | %s" % (fp.get("asin"), b, t))
        time.sleep(1.2)

    print("\ndry_short_needle_ok (no sheet write)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
