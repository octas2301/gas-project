# -*- coding: utf-8 -*-
"""Yahoo itemSearch クエリパターン別ヒット率調査（読取のみ・マスタ非書込）。

  python tools/yahoo_category_virt/hit_rate_yahoo_queries.py
  python tools/yahoo_category_virt/hit_rate_yahoo_queries.py --products products.sample.json
  python tools/yahoo_category_virt/hit_rate_yahoo_queries.py --dry-run

パターン定義: HIT_RATE_PATTERNS.md
認証: config.local.json または purchase_research_path3/config.local.json
     または環境変数 YAHOO_SHOPPING_CLIENT_ID
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import statistics
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "out"
PATH3_CFG = ROOT.parent / "purchase_research_path3" / "config.local.json"
API = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch"
SLEEP_SEC = 1.25
RESULTS = 20
SORT = "-score"
MAX_429_RETRY = 2

PET_RE = re.compile(
    r"カルカン|キャット|ドッグ|ペット|猫|犬|ニャン|わんこ|キャットフード|ドッグフード",
    re.I,
)


def load_cfg() -> dict[str, Any]:
    cfg: dict[str, Any] = {}
    for p in (ROOT / "config.local.json", PATH3_CFG):
        if p.is_file():
            cfg.update(json.loads(p.read_text(encoding="utf-8")))
            break
    v = os.environ.get("YAHOO_SHOPPING_CLIENT_ID", "").strip()
    if v:
        cfg["YAHOO_SHOPPING_CLIENT_ID"] = v
    return cfg


def simplify_name(name: str) -> str:
    s = str(name or "")
    s = re.sub(r"[（）()]", " ", s)
    s = re.sub(r"[・·／/]", " ", s)
    s = re.sub(r"[\s\u3000]+", " ", s)
    return s.strip()


def to_halfwidth(s: str) -> str:
    out = []
    for ch in s:
        o = ord(ch)
        if o == 0x3000:
            out.append(" ")
        elif 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        else:
            out.append(ch)
    return re.sub(r"\s+", " ", "".join(out)).strip()


def guess_series_core(name: str, override: str) -> str:
    if override and str(override).strip():
        return str(override).strip()
    n = str(name or "")
    for cand in ("食べるおだし", "食べるお出し"):
        if cand in n:
            return cand
    # 括弧より前の4文字以上
    head = re.split(r"[（(]", n, maxsplit=1)[0].strip()
    head = re.sub(r"[\s\u3000]+", "", head)
    if len(head) >= 4:
        return head[:12]
    return simplify_name(n)[:12]


def guess_flavor(name: str) -> str:
    m = re.search(r"[（(]\s*([^）)]+)", str(name or ""))
    if not m:
        return ""
    inner = simplify_name(m.group(1))
    parts = [p for p in inner.split() if p]
    return parts[0] if parts else ""


def build_patterns(product: dict[str, Any]) -> list[dict[str, Any]]:
    maker = str(product.get("maker") or "").strip()
    name = str(product.get("name") or "").strip()
    jan = str(product.get("jan") or "").strip()
    main_kw = str(product.get("mainKw") or "").strip()
    kind_n = int(product.get("kindN") or 0)
    series = guess_series_core(name, str(product.get("seriesCore") or ""))
    simple = simplify_name(name)
    flavor = guess_flavor(name)
    kind_tag = ("%d種" % kind_n) if kind_n >= 2 else ""

    specs: list[dict[str, Any]] = []

    def add(pid: str, label: str, mode: str, text: str, gas_like: bool) -> None:
        t = str(text or "").strip()
        if mode == "jan":
            if not t or not re.fullmatch(r"\d{8,}", t):
                return
        else:
            if len(t) < 2:
                return
        specs.append(
            {
                "patternId": pid,
                "label": label,
                "mode": mode,
                "query": t,
                "gasLike": gas_like,
            }
        )

    add("P01_JAN", "JAN(jan_code)", "jan", jan, True)
    add("P02_MAKER_NAME", "メーカー+商品名", "query", ("%s %s" % (maker, name)).strip(), True)
    add("P03_NAME", "商品名のみ", "query", name, True)
    add("P04_SIMPLE", "簡易名", "query", simple, True)
    add("P05_MAKER_SIMPLE", "メーカー+簡易名", "query", ("%s %s" % (maker, simple)).strip(), True)
    add("P06_SERIES", "シリーズ核のみ", "query", series, False)
    add("P07_MAKER_SERIES", "メーカー+シリーズ", "query", ("%s %s" % (maker, series)).strip(), False)
    if kind_tag:
        add("P08_KIND", "簡易+N種", "query", ("%s %s" % (simple, kind_tag)).strip(), True)
        add(
            "P09_MAKER_KIND",
            "メーカー+簡易+N種",
            "query",
            ("%s %s %s" % (maker, simple, kind_tag)).strip(),
            True,
        )
    if main_kw:
        add("P10_MAIN_KW", "メインKW", "query", main_kw, True)
    add(
        "P11_NORM_MS",
        "半角正規化メーカー+簡易",
        "query",
        to_halfwidth(("%s %s" % (maker, simple)).strip()),
        False,
    )
    if flavor:
        add(
            "P12_MAKER_SERIES_FLAVOR",
            "メーカー+シリーズ+味",
            "query",
            ("%s %s %s" % (maker, series, flavor)).strip(),
            False,
        )
    return specs


def http_get(url: str, timeout: int = 60) -> tuple[int, str]:
    req = urllib.request.Request(url, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return int(resp.status), resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace") if e.fp else ""
        return int(e.code), body
    except Exception as e:
        return 0, str(e)


def yahoo_search(client_id: str, mode: str, value: str) -> dict[str, Any]:
    qs: dict[str, str] = {
        "appid": client_id,
        "results": str(RESULTS),
        "sort": SORT,
    }
    if mode == "jan":
        qs["jan_code"] = value
    else:
        qs["query"] = value
    url = API + "?" + urllib.parse.urlencode(qs)
    out: dict[str, Any] = {
        "http": 0,
        "error": "",
        "total": 0,
        "returned": 0,
        "requestQuery": "",
        "titles": [],
        "cats": [],
    }
    code, body = 0, ""
    for attempt in range(MAX_429_RETRY + 1):
        code, body = http_get(url)
        if code != 429:
            break
        time.sleep(2.0 * (attempt + 1))
    out["http"] = code
    if code != 200:
        out["error"] = "HTTP %s %s" % (code, body[:160].replace("\n", " "))
        return out
    try:
        data = json.loads(body)
    except json.JSONDecodeError:
        out["error"] = "JSON error"
        return out
    out["total"] = int(data.get("totalResultsAvailable") or 0)
    hits = data.get("hits") or []
    out["returned"] = len(hits)
    req = data.get("request") or {}
    out["requestQuery"] = str(req.get("query") or req.get("jan_code") or "")
    for h in hits[:RESULTS]:
        out["titles"].append(str((h or {}).get("name") or ""))
        gc = (h or {}).get("genreCategory") or {}
        out["cats"].append(str(gc.get("id") or ""))
    return out


def score_titles(
    titles: list[str], maker: str, series: str
) -> tuple[int, int, int]:
    top = titles[:5]
    maker_n = 0
    series_n = 0
    pet_n = 0
    mk = maker.strip()
    sc = series.strip()
    for t in top:
        if mk and mk in t:
            maker_n += 1
        if sc and sc in t:
            series_n += 1
        if PET_RE.search(t):
            pet_n += 1
    return maker_n, series_n, pet_n


def load_products(path: Path) -> list[dict[str, Any]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise SystemExit("products JSON must be a list")
    return data


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    keys = list(rows[0].keys())
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=keys, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def summarize(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for r in rows:
        if r.get("error"):
            continue
        by[str(r["patternId"])].append(r)
    out = []
    for pid, xs in sorted(by.items()):
        bins = [int(x["hit_bin"]) for x in xs]
        totals = [int(x["total"]) for x in xs]
        maker_avg = statistics.mean(int(x["maker_top5"]) for x in xs) if xs else 0
        pet_avg = statistics.mean(int(x["pet_top5"]) for x in xs) if xs else 0
        out.append(
            {
                "patternId": pid,
                "label": xs[0]["label"] if xs else "",
                "gasLike": xs[0]["gasLike"] if xs else "",
                "n": len(xs),
                "hit_rate": round(sum(bins) / len(bins), 3) if bins else 0,
                "total_median": int(statistics.median(totals)) if totals else 0,
                "total_mean": int(statistics.mean(totals)) if totals else 0,
                "maker_top5_mean": round(maker_avg, 2),
                "pet_top5_mean": round(pet_avg, 2),
            }
        )
    out.sort(key=lambda r: (-float(r["hit_rate"]), -int(r["total_median"])))
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Yahoo query pattern hit-rate bakeoff")
    ap.add_argument(
        "--products",
        default=str(ROOT / "products.sample.json"),
        help="商品 JSON（配列）",
    )
    ap.add_argument("--sleep", type=float, default=SLEEP_SEC)
    ap.add_argument("--dry-run", action="store_true", help="APIを呼ばずクエリ一覧のみ")
    ap.add_argument("--out-dir", default=str(OUT_DIR))
    args = ap.parse_args()

    products = load_products(Path(args.products))
    cfg = load_cfg()
    client_id = str(cfg.get("YAHOO_SHOPPING_CLIENT_ID") or "").strip()
    if not args.dry_run and not client_id:
        raise SystemExit(
            "YAHOO_SHOPPING_CLIENT_ID がありません。"
            " tools/yahoo_category_virt/config.local.json を作成するか、"
            " purchase_research_path3/config.local.json / 環境変数を設定してください。"
        )

    stamp = datetime.now(timezone.utc).astimezone().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    detail_rows: list[dict[str, Any]] = []
    for prod in products:
        pid = str(prod.get("id") or "unknown")
        maker = str(prod.get("maker") or "")
        series = guess_series_core(str(prod.get("name") or ""), str(prod.get("seriesCore") or ""))
        patterns = build_patterns(prod)
        print("=== %s (%d patterns) ===" % (pid, len(patterns)))
        for spec in patterns:
            row: dict[str, Any] = {
                "productId": pid,
                "sheetRow": prod.get("sheetRow") or "",
                "patternId": spec["patternId"],
                "label": spec["label"],
                "gasLike": "Y" if spec["gasLike"] else "N",
                "mode": spec["mode"],
                "query": spec["query"],
                "http": "",
                "error": "",
                "total": 0,
                "hit_bin": 0,
                "returned": 0,
                "maker_top5": 0,
                "series_top5": 0,
                "pet_top5": 0,
                "top1_title": "",
                "top1_cat": "",
                "requestQuery": "",
            }
            if args.dry_run:
                print("  %s\t%s" % (spec["patternId"], spec["query"][:70]))
                detail_rows.append(row)
                continue

            res = yahoo_search(client_id, spec["mode"], spec["query"])
            time.sleep(max(0.0, float(args.sleep)))
            row["http"] = res["http"]
            row["error"] = res["error"]
            row["total"] = res["total"]
            row["hit_bin"] = 1 if res["total"] >= 1 else 0
            row["returned"] = res["returned"]
            row["requestQuery"] = res["requestQuery"]
            m5, s5, p5 = score_titles(res["titles"], maker, series)
            row["maker_top5"] = m5
            row["series_top5"] = s5
            row["pet_top5"] = p5
            row["top1_title"] = (res["titles"][0] if res["titles"] else "")[:60]
            row["top1_cat"] = res["cats"][0] if res["cats"] else ""
            print(
                "  %s total=%s hit=%s maker5=%s pet5=%s q=%s"
                % (
                    spec["patternId"],
                    row["total"],
                    row["hit_bin"],
                    row["maker_top5"],
                    row["pet_top5"],
                    spec["query"][:40],
                )
            )
            if res["error"]:
                print("    ERR %s" % res["error"][:100])
            detail_rows.append(row)

    detail_path = out_dir / ("hit_rate_detail_%s.csv" % stamp)
    summary_path = out_dir / ("hit_rate_summary_%s.csv" % stamp)
    write_csv(detail_path, detail_rows)
    summary = summarize(detail_rows) if not args.dry_run else []
    write_csv(summary_path, summary)

    md_path = out_dir / ("hit_rate_report_%s.md" % stamp)
    lines = [
        "# Yahoo クエリヒット率レポート",
        "",
        "- 実行: %s" % stamp,
        "- 商品数: %d" % len(products),
        "- sort: %s / results: %d" % (SORT, RESULTS),
        "- dry_run: %s" % args.dry_run,
        "",
        "## パターン別サマリ（ヒット率降順）",
        "",
    ]
    if summary:
        lines.append(
            "| patternId | hit_rate | total_median | maker_top5_mean | pet_top5_mean | gasLike | label |"
        )
        lines.append("|---|---:|---:|---:|---:|---|---|")
        for s in summary:
            lines.append(
                "| %s | %.0f%% | %s | %s | %s | %s | %s |"
                % (
                    s["patternId"],
                    100 * float(s["hit_rate"]),
                    s["total_median"],
                    s["maker_top5_mean"],
                    s["pet_top5_mean"],
                    s["gasLike"],
                    s["label"],
                )
            )
    else:
        lines.append("（dry-run のため API 集計なし）")
    lines.append("")
    lines.append("- detail: `%s`" % detail_path.as_posix())
    lines.append("- summary: `%s`" % summary_path.as_posix())
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print("")
    print("wrote %s" % detail_path)
    print("wrote %s" % summary_path)
    print("wrote %s" % md_path)
    if summary:
        print("")
        print("TOP patterns by hit_rate:")
        for s in summary[:8]:
            print(
                "  %s  hit_rate=%.0f%%  median_total=%s  pet5=%.2f"
                % (s["patternId"], 100 * float(s["hit_rate"]), s["total_median"], s["pet_top5_mean"])
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
