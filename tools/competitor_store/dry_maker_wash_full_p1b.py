# -*- coding: utf-8 -*-
"""P1b ドライ: >500 時のタイトル→カテゴリ推定（Gemini→OpenAI）＋分割語列挙。

既定: Catalog列挙でサンプルタイトル取得→AI推定のみ（シート非書）。
--enum-words: 推定語で Catalog 再検索（件数確認、シート非書）。
--apply-split: 調査複製 `①洗い分割語` へ追記（慎重）。
Keepa／門／最終調査日は書かない。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SMOKE = ROOT.parent / "spapi_smoke"
PATH3 = ROOT.parent / "purchase_research_path3"
SECRETS = ROOT.parent / "set_main_image" / "secrets"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(SMOKE))
sys.path.insert(0, str(PATH3))

from apply_keepa_full import (  # noqa: E402
    RESEARCH_SS,
    as_dicts,
    read_all,
    sheets_service,
    utc_now,
)
from search_catalog_keywords import asins_from_body, search_catalog_items  # noqa: E402
from spapi_smoke import _load_config, request_lwa_access_token, resolve_cred  # noqa: E402

SPLIT_SHEET = "①洗い分割語"
SPLIT_HEADERS = ["メーカー", "分割語", "状態", "推定元", "サンプル数", "列挙ASIN数", "更新"]
SPLIT_THRESHOLD = 500
SAMPLE_MAX = 200
WORD_MAX = 15
DEFAULT_MAKER = "尾西食品"
GEMINI_MODEL = "gemini-3.6-flash"
OPENAI_MODEL = "gpt-4o-mini"


def title_hits(rows: list[dict], needles: list[str]) -> list[dict]:
    out = []
    for r in rows:
        blob = str(r.get("title") or "") + " " + str(r.get("brand") or "")
        if any(n and n in blob for n in needles):
            out.append(r)
    return out


def load_key(*names: str) -> str:
    for n in names:
        v = (os.environ.get(n) or "").strip()
        if v:
            return v
    for n in names:
        for p in (SECRETS / f"{n}.txt", SECRETS / (n.lower() + ".txt")):
            if p.is_file():
                t = p.read_text(encoding="utf-8").strip()
                if t and not t.startswith("#"):
                    return t
    return ""


def catalog_page(access: str, endpoint: str, mp: str, ua: str, keywords: str, token: str = ""):
    resp = search_catalog_items(endpoint, access, mp, ua, keywords, page_size=20, page_token=token or "")
    body = resp.json() if resp.content else {}
    rows = asins_from_body(body)
    nxt = str((body.get("pagination") or {}).get("nextToken") or "").strip()
    return resp.status_code, rows, nxt


def enumerate_with_titles(
    access: str,
    endpoint: str,
    mp: str,
    ua: str,
    keywords: str,
    needles: list[str],
    max_pages: int = 200,
    stop_at: int = 0,
) -> tuple[list[str], list[str], int, bool]:
    seen: set[str] = set()
    ordered: list[str] = []
    titles: list[str] = []
    token = ""
    pages = 0
    while pages < max_pages:
        if stop_at > 0 and len(ordered) >= stop_at:
            return ordered, titles, pages, False
        code, rows, nxt = catalog_page(access, endpoint, mp, ua, keywords, token)
        if code != 200:
            raise RuntimeError("Catalog HTTP %s keywords=%s" % (code, keywords))
        pages += 1
        for r in title_hits(rows, needles):
            tit = str(r.get("title") or "").strip()
            if tit and len(titles) < SAMPLE_MAX:
                titles.append(tit)
            a = str(r.get("asin") or "").strip().upper()
            if a and a not in seen:
                seen.add(a)
                ordered.append(a)
        token = nxt
        if not token:
            return ordered, titles, pages, True
        time.sleep(0.25)
    return ordered, titles, pages, False


def build_prompt(parent: str, titles: list[str]) -> str:
    sample = titles[:SAMPLE_MAX]
    return (
        "あなたは日本の食品ECの商品分類アシスタントです。\n"
        "メーカー「%s」の Amazon 商品名サンプルから、Catalog 再検索用の短いカテゴリ語を最大%d個挙げてください。\n"
        "条件: 日本語の短い名詞句のみ。メーカー名自体は入れない。冷凍・冷蔵・生鮮は入れない。\n"
        "出力は JSON 配列のみ（例: [\"みそ汁\",\"おかゆ\",\"カレー\"]）。説明文は不要。\n\n"
        "商品名サンプル:\n%s"
    ) % (parent, WORD_MAX, "\n".join(sample))


def parse_words(text: str) -> list[str]:
    raw = (text or "").strip()
    if not raw:
        return []
    m = re.search(r"\[[\s\S]*\]", raw)
    if m:
        try:
            arr = json.loads(m.group(0))
            if isinstance(arr, list):
                out = []
                seen = set()
                for x in arr:
                    w = str(x or "").strip()
                    if w and w not in seen and len(w) <= 20:
                        seen.add(w)
                        out.append(w)
                return out[:WORD_MAX]
        except json.JSONDecodeError:
            pass
    parts = re.split(r"[\n|｜,、/／]+", raw)
    out: list[str] = []
    seen: set[str] = set()
    for p in parts:
        w = re.sub(r"^[\d.\-\*\s]+", "", p)
        w = w.replace('"', "").replace("「", "").replace("」", "").strip()
        if not w or len(w) > 20 or w in seen:
            continue
        if re.match(r"^(json|list|カテゴリ|以下)", w, re.I):
            continue
        seen.add(w)
        out.append(w)
    return out[:WORD_MAX]


def gemini_text(prompt: str) -> str | None:
    key = load_key("GEMINI_API_KEY", "GOOGLE_API_KEY")
    if not key:
        print("WARN no GEMINI_API_KEY")
        return None
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        + GEMINI_MODEL
        + ":generateContent?key="
        + key
    )
    body = json.dumps(
        {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.2, "maxOutputTokens": 2048},
        }
    ).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        print("WARN Gemini HTTP", e.code, e.read()[:200])
        return None
    except Exception as e:
        print("WARN Gemini", e)
        return None
    try:
        return str(data["candidates"][0]["content"]["parts"][0]["text"] or "").strip() or None
    except (KeyError, IndexError, TypeError):
        return None


def openai_text(prompt: str) -> str | None:
    key = load_key("OPENAI_API_KEY")
    if not key:
        print("WARN no OPENAI_API_KEY")
        return None
    body = json.dumps(
        {
            "model": OPENAI_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.2,
            "max_completion_tokens": 600,
        }
    ).encode("utf-8")
    req = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=body,
        headers={"Content-Type": "application/json", "Authorization": "Bearer " + key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        print("WARN OpenAI HTTP", e.code, e.read()[:200])
        return None
    except Exception as e:
        print("WARN OpenAI", e)
        return None
    if data.get("error"):
        print("WARN OpenAI error", data["error"])
        return None
    try:
        return str(data["choices"][0]["message"]["content"] or "").strip() or None
    except (KeyError, IndexError, TypeError):
        return None


def infer_words(parent: str, titles: list[str]) -> tuple[list[str], str]:
    prompt = build_prompt(parent, titles)
    g = gemini_text(prompt)
    words = parse_words(g or "")
    if words:
        return words, "gemini"
    print("Gemini empty/fail → OpenAI")
    o = openai_text(prompt)
    words = parse_words(o or "")
    if words:
        return words, "openai"
    return [], ""


def ensure_split_headers(svc, sid: str) -> None:
    meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
    titles = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if SPLIT_SHEET not in titles:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [{"addSheet": {"properties": {"title": SPLIT_SHEET}}}]},
        ).execute()
    got = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=sid, range="'" + SPLIT_SHEET + "'!1:1")
        .execute()
        .get("values")
        or []
    )
    if not got:
        svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range="'" + SPLIT_SHEET + "'!A1",
            valueInputOption="RAW",
            body={"values": [SPLIT_HEADERS]},
        ).execute()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--maker", default=DEFAULT_MAKER)
    ap.add_argument("--entry", default="")
    ap.add_argument("--enum-words", action="store_true", help="推定語でCatalog再検索（件数のみ）")
    ap.add_argument("--max-words", type=int, default=5, help="--enum-words 時の最大語数")
    ap.add_argument("--apply-split", action="store_true", help="①洗い分割語へ追記")
    args = ap.parse_args()
    parent = str(args.maker).strip()
    entry = str(args.entry or parent).strip()
    needles = [parent]

    cfg = _load_config(SMOKE / "config.local.json")
    access = str(
        request_lwa_access_token(
            resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID"),
            resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET"),
            resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN"),
        ).get("access_token")
        or ""
    )
    if not access:
        print("FAIL no access_token")
        return 1
    endpoint = str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com")
    mp = str(cfg.get("marketplace_id") or "A1VC38T7YXB528")
    ua = str(cfg.get("user_agent") or "PurchaseResearchDry/1.0")

    print("enum_main maker=%s entry=%s stop_at=%d" % (parent, entry, SPLIT_THRESHOLD + 1))
    asins, titles, pages, done = enumerate_with_titles(
        access,
        endpoint,
        mp,
        ua,
        entry,
        needles,
        max_pages=200,
        stop_at=SPLIT_THRESHOLD + 1,
    )
    print(
        "enum_main uniq=%d pages=%d titles=%d done=%s"
        % (len(asins), pages, len(titles), done)
    )
    if len(asins) <= SPLIT_THRESHOLD:
        print("P1b_DRY SKIP under_threshold (use p1 for ≤500)")
        return 0
    if not titles:
        print("P1b_DRY FAIL no_titles")
        return 2

    words, source = infer_words(parent, titles)
    print("infer source=%s n=%d words=%s" % (source, len(words), words))
    if not words:
        print("P1b_DRY FAIL infer_empty")
        return 3

    if args.enum_words:
        seen = set(asins)
        for w in words[: max(1, args.max_words)]:
            q = parent + " " + w
            g2, _t2, p2, d2 = enumerate_with_titles(
                access, endpoint, mp, ua, q, needles, max_pages=40, stop_at=0
            )
            add = 0
            for a in g2:
                if a not in seen:
                    seen.add(a)
                    add += 1
            print("word=%s hit=%d new=%d pages=%d done=%s" % (w, len(g2), add, p2, d2))
        print("after_split unique_est=%d" % len(seen))

    if not args.apply_split:
        print("dry_ok (no sheet write)")
        print("P1b_DRY PASS")
        return 0

    svc = sheets_service(write=True, interactive=True)
    if not svc:
        print("no_creds for --apply-split")
        return 1
    ensure_split_headers(svc, RESEARCH_SS)
    now = utc_now()
    rows = [[parent, w, "未", source, len(titles), "", now] for w in words]
    svc.spreadsheets().values().append(
        spreadsheetId=RESEARCH_SS,
        range="'" + SPLIT_SHEET + "'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()
    print("apply_split_ok n=%d" % len(rows))
    print("P1b_DRY PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
