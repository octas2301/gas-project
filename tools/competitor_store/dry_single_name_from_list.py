# -*- coding: utf-8 -*-
"""仕入検討品番リスト実データ × 単品名パターン仮想。

既定: シート読取のみ（非書）。AI は --ai で Gemini→OpenAI。
存在確認は Catalog 本線ではなく、タイトル内包の簡易判定（仮想）。

  python tools/competitor_store/dry_single_name_from_list.py
  python tools/competitor_store/dry_single_name_from_list.py --n 20 --ai
  python tools/competitor_store/dry_single_name_from_list.py --n 20 --ai --quote-only
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "spapi_smoke"))

from apply_keepa_full import RESEARCH_SS, as_dicts, read_all, sheets_service  # noqa: E402
from dry_t1_list_paste import DATA_START, HEADER_ROW, LIST_TITLE, ORIG_SS  # noqa: E402
from single_name_classify import classify_single_name  # noqa: E402

OUT_DIR = ROOT / "out" / "single_name_live"
SET_STRIP_RE = re.compile(
    r"\d+\s*個\s*(セット|入)|"
    r"\d+\s*袋\s*(セット|入)|"
    r"[×xX]\s*\d+\s*(袋|個|本)|"
    r"\d+\s*食(?:入|セット)?|"
    r"まとめ|お買い得|送料無料|【[^】]*】",
    re.I,
)


def _col(row: dict, *names: str) -> str:
    for n in names:
        if n in row and str(row.get(n) or "").strip():
            return str(row.get(n) or "").strip()
    # fuzzy: header may have newlines
    for k, v in row.items():
        kk = str(k or "").replace("\n", " ").strip()
        for n in names:
            if kk == n or kk.startswith(n):
                return str(v or "").strip()
    return ""


def load_list_rows(svc) -> list[dict]:
    raw = read_all(svc, RESEARCH_SS, LIST_TITLE)
    if not raw or len(raw) < HEADER_ROW:
        return []
    headers = [str(x).replace("\n", " ").strip() for x in raw[HEADER_ROW - 1]]
    out: list[dict] = []
    for r in raw[DATA_START - 1 :]:
        while len(r) < len(headers):
            r.append("")
        d = {headers[i]: r[i] for i in range(len(headers))}
        asin = str(d.get("ASIN") or "").strip().upper()
        if len(asin) == 10:
            out.append(d)
    return out


def strip_hint(title: str) -> str:
    s = SET_STRIP_RE.sub(" ", str(title or ""))
    s = re.sub(r"\s+", " ", s).strip()
    return s


def sale_form_code(raw: str) -> str:
    s = str(raw or "")
    if "③" in s or "別品番" in s and "セット" in s:
        if "④" in s or ("まとめ" in s and "別" in s):
            return "④"
        return "③別品番とセット売り"
    if "④" in s:
        return "④"
    if "②" in s or "まとめ" in s:
        return "②"
    if "①" in s or "単品" in s:
        return "①"
    return "①"


def parse_qty(raw: str) -> int:
    try:
        n = int(float(str(raw or "").replace(",", "")))
        return n if n >= 1 else 1
    except ValueError:
        return 1


def parse_jans(raw: str) -> list[str]:
    s = str(raw or "").strip()
    if not s:
        return []
    parts = re.split(r"[,|｜\s]+", s)
    out = []
    for p in parts:
        p = re.sub(r"\D", "", p)
        if len(p) >= 8 and p not in out:
            out.append(p)
    return out


def title_match_level(draft: str, title: str) -> str:
    d = re.sub(r"\s+", "", draft or "")
    t = re.sub(r"\s+", "", title or "")
    if not d or not t:
        return "low"
    if d in t or t in d:
        return "high"
    # token overlap (2+ chars)
    toks = [d[i : i + 2] for i in range(len(d) - 1)]
    if not toks:
        return "low"
    hit = sum(1 for x in toks if x in t)
    return "high" if hit / len(toks) >= 0.45 else "low"


def soft_existence(draft: str, title: str, maker: str) -> str:
    """Catalog の代わり。仮名がタイトルに十分近ければ ok。"""
    if not draft:
        return "ng"
    if title_match_level(draft, title) == "high":
        return "ok"
    # maker + draft both appear loosely
    if maker and maker.replace(" ", "") in (title or "").replace(" ", "") and len(draft) >= 2:
        if any(draft[i : i + 2] in (title or "") for i in range(max(0, len(draft) - 1))):
            return "ok"
    return "ng"


def ai_draft_name(title: str, maker: str, qty: int, sale: str) -> tuple[str, str]:
    """Returns (name, confidence high|low). mix は low + 空名。"""
    from dry_single_name_compare import (  # type: ignore
        PROMPT_ROW_AI,
        force_mix,
        gemini_text,
        openai_text,
        parse_kind_name,
    )

    if force_mix(title):
        return "", "low"
    prompt = PROMPT_ROW_AI.format(maker=maker, sale=sale, qty=qty, title=title)
    text = openai_text(prompt, 400) or gemini_text(prompt, 1024)
    kind, name = parse_kind_name(text or "")
    if kind == "mix":
        return "", "low"
    if kind == "single" and name:
        return name, "high"
    return strip_hint(title), "low"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=20, help="件数上限")
    ap.add_argument("--ai", action="store_true", help="Gemini/OpenAI で単品名案")
    ap.add_argument("--quote-only", action="store_true", help="見積候補のみ")
    ap.add_argument("--sleep", type=float, default=0.8, help="AI呼び出し間隔秒")
    args = ap.parse_args()

    if RESEARCH_SS == ORIG_SS:
        print("FAIL refuse original SS")
        return 2

    svc = sheets_service(write=False)
    rows = load_list_rows(svc)
    if not rows:
        print("FAIL no ASIN rows on", LIST_TITLE)
        return 1

    picked: list[dict] = []
    for r in rows:
        if args.quote_only:
            q = _col(r, "見積候補")
            if q != "見積候補":
                continue
        picked.append(r)
        if len(picked) >= max(1, args.n):
            break

    print("list_asin_total", len(rows), "picked", len(picked), "ai", bool(args.ai), "quote_only", args.quote_only)
    if not picked:
        print("FAIL no rows picked")
        return 1

    results = []
    for i, r in enumerate(picked, 1):
        asin = _col(r, "ASIN").upper()
        title = _col(r, "商品名")
        maker = _col(r, "メーカー名", "メーカー")
        jan_raw = _col(r, "JANコード")
        sale = sale_form_code(_col(r, "販売形態"))
        qty = parse_qty(_col(r, "同商品のセット数量"))
        quote = _col(r, "見積候補")

        if args.ai:
            draft, conf = ai_draft_name(title, maker, qty, sale)
            time.sleep(max(0.0, args.sleep))
        else:
            draft, conf = strip_hint(title), "high"

        keepa_jans = parse_jans(jan_raw)
        # 仮想: 検索JANは Keepa と同型（Catalog未接続）
        search_jans = list(keepa_jans)
        tmatch = title_match_level(draft, title)
        existence = soft_existence(draft, title, maker)
        # 詰め合わせは存在OKでも P-MIX
        inp = {
            "draft_name": draft,
            "sale_form": sale,
            "set_qty": qty,
            "existence": existence,
            "ai_confidence": conf,
            "title_match": tmatch,
            "row_asin": asin,
            "hit_asin": asin if existence == "ok" else "",
            "keepa_jans": keepa_jans,
            "search_jans": search_jans,
            "keepa_title": title,
        }
        cls = classify_single_name(inp)
        rec = {
            "i": i,
            "asin": asin,
            "maker": maker,
            "title": title[:80],
            "draft": draft,
            "ai_conf": conf,
            "sale": sale,
            "qty": qty,
            "quote": quote,
            "existence_soft": existence,
            "title_match": tmatch,
            "pattern": cls["pattern"],
            "single_trust": cls["single_trust"],
            "jan_trust": cls["jan_trust"],
            "for_estimate": cls["for_estimate"],
            "need_human": cls["need_human"],
            "corpus_assist_ok": cls["corpus_assist_ok"],
            "write_single_name": cls["write_single_name"],
        }
        results.append(rec)
        print(
            f"{i:02d} {cls['pattern']:5} est={cls['for_estimate']} human={cls['need_human']} "
            f"conf={conf} | {maker[:8]} | {draft[:28]} | {asin}"
        )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    out_json = OUT_DIR / f"live_{stamp}.json"
    out_json.write_text(json.dumps(results, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    from collections import Counter

    pat = Counter(r["pattern"] for r in results)
    est_n = sum(1 for r in results if r["for_estimate"])
    hum_n = sum(1 for r in results if r["need_human"])
    print("pattern_counts", dict(pat.most_common()))
    print("for_estimate", est_n, "/", len(results), "need_human", hum_n)
    print("wrote", out_json)
    print("NOTE existence=soft(title)。Catalog本線ではない。AI未使用時はストリップ仮名。")
    print("DRY", "PASS" if len(results) >= 1 else "FAIL")
    return 0 if results else 1


if __name__ == "__main__":
    raise SystemExit(main())
