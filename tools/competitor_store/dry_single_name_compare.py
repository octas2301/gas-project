# -*- coding: utf-8 -*-
"""単品商品名: 現状 vs 行AI（S相当） vs 辞書寄せ の仮想比較。

品番リストは非書。調査複製へ比較タブのみ作成／上書き。
既定対象: 見積候補のみ（--all-list でリスト全ASIN）。

  python tools/competitor_store/dry_single_name_compare.py
  python tools/competitor_store/dry_single_name_compare.py --no-row-ai
  python tools/competitor_store/dry_single_name_compare.py --all-list
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "spapi_smoke"))
sys.path.insert(0, str(ROOT.parent / "common_ss_view"))

from apply_keepa_full import RESEARCH_SS, sheets_service  # noqa: E402
from dry_maker_wash_full_p1b import load_key  # noqa: E402
from dry_single_name_from_list import (  # noqa: E402
    _col,
    load_list_rows,
    parse_jans,
    parse_qty,
    sale_form_code,
    strip_hint,
)
from dry_t1_list_paste import ORIG_SS  # noqa: E402

OUT_DIR = ROOT / "out" / "single_name_compare"
COMPARE_SHEET = "単品名比較_仮想"
GEMINI_MODEL = "gemini-3.6-flash"
OPENAI_MODEL = "gpt-4o-mini"
SAMPLE_TITLES = 45
SLEEP_SEC = 0.7

# --- 行AI: 金標準34件・2回目 PASS（PROMPT_ROW_AI.winner.txt と同文）---
_WINNER_PROMPT = ROOT / "out" / "single_name_compare" / "PROMPT_ROW_AI.winner.txt"
PROMPT_ROW_AI = _WINNER_PROMPT.read_text(encoding="utf-8") if _WINNER_PROMPT.is_file() else (
    "メーカー: {maker}\n販売形態: {sale}\nセット数量: {qty}\nタイトル: {title}\n"
    'JSONのみ: {{"kind":"single","name":"品目 内容量"}} 異種は kind=mix name空\n'
)

PROMPT_DICT_AI = (
    "あなたは食品の仕入れ見積用の商品名整理係です。\n"
    "同一メーカーのAmazonタイトル群から、問屋・メーカー向けの短い単品名の辞書をJSONで返してください。\n"
    '形式: {{"names":["チーズかつお 1kg","チーズかつお 220g"],"notes":"短い説明"}}\n'
    "正の形: 「味・品目＋内容量」。シリーズ飾り・販促語（お魚屋さんがつくった…、徳用、大容量、業務用、おつまみ等）は落とす。\n"
    "内容量が違うものは別エントリ。メーカー名は names に入れない。\n"
    "既存の短い呼び名候補があれば、その表記に揃える。\n"
    "メーカー: {maker}\n"
    "既存候補:\n{seeds}\n"
    "タイトル一覧:\n{titles}\n"
)

DECOR_RE = re.compile(
    r"お魚屋さんがつくった大人のおつまみ|"
    r"お魚屋さんがつくった|"
    r"大人のおつまみ|"
    r"お徳用|徳用|業務用|大容量|大容量サイズ|おつまみ|珍味|"
    r"個包装|角切|送料無料|まとめ買い|お買い得|製造|"
    r"【[^】]*】",
    re.I,
)
SIZE_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*(kg|g|ml|mL|l|L|リットル|食|袋|本|個)",
    re.I,
)

HEADERS = [
    "ASIN",
    "メーカー名",
    "見積候補",
    "商品名",
    "現状_単品商品名",
    "行AI_正式商品名",
    "行AI_種別",
    "行AI_型番",
    "行AI_照合メモ",
    "新案_辞書寄せ",
    "現状=行AI",
    "現状=辞書",
    "行AI=辞書",
    "寄せ方法",
    "辞書件数",
    "JANコード",
    "販売形態",
    "同商品のセット数量",
    "備考",
]


def gemini_text(prompt: str, max_tokens: int = 4096) -> str | None:
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
            "generationConfig": {"temperature": 0.1, "maxOutputTokens": max_tokens},
        }
    ).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        parts = ((data.get("candidates") or [{}])[0].get("content") or {}).get("parts") or []
        texts = [str(p.get("text") or "") for p in parts if p.get("text")]
        out = "\n".join(texts).strip()
        if out:
            return out
        # thinking モデルは parts が空で MAX_TOKENS になりうる
        fr = str(((data.get("candidates") or [{}])[0].get("finishReason") or ""))
        if fr:
            print("WARN Gemini empty finishReason=%s" % fr)
        return None
    except Exception as e:
        print("WARN Gemini", e)
        return None


def openai_text(prompt: str, max_tokens: int = 2500) -> str | None:
    key = load_key("OPENAI_API_KEY")
    if not key:
        print("WARN no OPENAI_API_KEY")
        return None
    body = json.dumps(
        {
            "model": OPENAI_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1,
            "max_completion_tokens": max_tokens,
        }
    ).encode("utf-8")
    req = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=body,
        headers={"Content-Type": "application/json", "Authorization": "Bearer " + key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return str(data["choices"][0]["message"]["content"] or "").strip() or None
    except urllib.error.HTTPError as e:
        print("WARN OpenAI HTTP", e.code, e.read()[:200])
        return None
    except Exception as e:
        print("WARN OpenAI", e)
        return None


def gemini_dict(prompt: str) -> str | None:
    return gemini_text(prompt, 4096)


def openai_dict(prompt: str) -> str | None:
    return openai_text(prompt, 2500)


MIX_TITLE_RE = re.compile(r"食べ比べ|各\d+\s*袋|詰め合わせ|アソート|よりどり")
TWO_SKU_RE = re.compile(r"チーズかつお.{0,40}まぐろ|まぐろチーズ.{0,40}チーズかつお|かつお.{0,20}まぐろ.{0,20}ぶり")


def force_mix(title: str) -> bool:
    t = str(title or "")
    if MIX_TITLE_RE.search(t) and TWO_SKU_RE.search(t):
        return True
    if t.count("＋") >= 1 and MIX_TITLE_RE.search(t):
        return True
    return False


def parse_kind_name(text: str) -> tuple[str, str]:
    if not text:
        return "", ""
    cleaned = re.sub(r"```json|```", "", text)
    m = re.search(r"\{[\s\S]*\}", cleaned)
    if not m:
        return "", ""
    try:
        obj = json.loads(m.group(0))
    except json.JSONDecodeError:
        return "", ""
    kind = str(obj.get("kind") or "").strip().lower()
    name = str(obj.get("name") or "").strip()
    if kind not in ("single", "mix"):
        kind = "mix" if not name else "single"
    return kind, name


def row_ai_draft(title: str, maker: str, qty: int, sale: str) -> tuple[str, str, str, str]:
    """合格プロンプト。Returns (name, kind, '', memo). mix は要確認。"""
    if force_mix(title):
        return "要確認（異種セット）", "mix", "", "タイトルが食べ比べ/異種"
    prompt = PROMPT_ROW_AI.format(maker=maker, sale=sale, qty=qty, title=title)
    text = openai_text(prompt, 400) or gemini_text(prompt, 1024)
    kind, name = parse_kind_name(text or "")
    if kind == "mix":
        return "要確認（異種セット）", "mix", "", "AI kind=mix"
    if kind == "single" and name:
        return name, "single", "", ""
    return strip_hint(title), "low", "", "parse失敗"


def extract_sizes(text: str) -> list[str]:
    out = []
    for m in SIZE_RE.finditer(text or ""):
        num, unit = m.group(1), m.group(2)
        u = unit.lower()
        if u in ("l", "リットル"):
            u = "L"
        elif u == "ml":
            u = "ml"
        elif u == "kg":
            u = "kg"
        elif u == "g":
            u = "g"
        key = ("%s%s" % (num, u)).replace(" ", "")
        if key not in out:
            out.append(key)
    return out


def heuristic_canon(title: str, maker: str) -> str:
    s = strip_hint(title or "")
    if maker:
        s = re.sub(re.escape(maker), " ", s, count=1)
    s = DECOR_RE.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip()
    sizes = extract_sizes(s)
    # コア: サイズ直前までの短い塊を優先
    core = s
    if sizes:
        # 最後のサイズ周辺を残す
        m = list(SIZE_RE.finditer(s))
        if m:
            last = m[-1]
            before = s[: last.start()].strip()
            # 末尾の短い語群
            toks = [t for t in before.split() if t]
            keep = toks[-4:] if len(toks) > 4 else toks
            core = (" ".join(keep) + " " + last.group(0)).strip()
    core = DECOR_RE.sub(" ", core)
    core = re.sub(r"\s+", " ", core).strip()
    return core[:80]


def parse_names_json(text: str) -> list[str]:
    if not text:
        return []
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        # 行リスト fallback
        names = []
        for line in text.splitlines():
            line = line.strip().lstrip("-*・0123456789.）) ")
            if 2 <= len(line) <= 60 and "タイトル" not in line:
                names.append(line)
        return names[:80]
    try:
        obj = json.loads(m.group(0))
    except json.JSONDecodeError:
        return []
    raw = obj.get("names") or obj.get("辞書") or []
    out = []
    for x in raw:
        n = str(x or "").strip()
        if n and n not in out:
            out.append(n)
    return out


def build_dict_for_maker(
    maker: str,
    titles: list[str],
    existing_shorts: list[str],
    use_ai: bool,
) -> tuple[list[str], str]:
    """Returns (names, source)."""
    # seed: 短い「味+内容量」のみ。シリーズ飾り入りの現状名は入れない（比較が現状に引きずられるため）
    seeds = []
    for s in existing_shorts:
        s = str(s or "").strip()
        if not s:
            continue
        if DECOR_RE.search(s):
            continue
        if "徳用" in s or "業務用" in s or "大容量" in s:
            continue
        if extract_sizes(s) and len(s) <= 28 and s not in seeds:
            seeds.append(s)
    heur = []
    for t in titles:
        h = heuristic_canon(t, maker)
        if h and h not in heur:
            heur.append(h)
    if not use_ai:
        merged = list(seeds)
        for h in heur:
            if h not in merged:
                merged.append(h)
        return merged[:80], "heuristic"

    sample = titles[:SAMPLE_TITLES]
    prompt = PROMPT_DICT_AI.format(
        maker=maker,
        seeds="\n".join("- " + x for x in seeds[:30]) or "(なし)",
        titles="\n".join("%d. %s" % (i + 1, t[:120]) for i, t in enumerate(sample)),
    )
    text = gemini_dict(prompt) or openai_dict(prompt)
    names = parse_names_json(text or "")
    if not names:
        merged = list(seeds)
        for h in heur:
            if h not in merged:
                merged.append(h)
        return merged[:80], "ai_fail_heuristic"
    # AI案も飾りを落とす／長すぎは heur 優先で混ぜる
    cleaned = []
    for n in names:
        n2 = DECOR_RE.sub(" ", n)
        n2 = re.sub(r"\s+", " ", n2).strip()
        if not n2:
            continue
        if "徳用" in n2 or "業務用" in n2 or "大容量" in n2:
            n2 = re.sub(r"徳用|業務用|大容量サイズ|大容量", " ", n2)
            n2 = re.sub(r"\s+", " ", n2).strip()
        if n2 and n2 not in cleaned:
            cleaned.append(n2)
    names = cleaned
    # merge seeds first
    out = list(seeds)
    for n in names:
        if n not in out:
            out.append(n)
    for h in heur:
        if h not in out and len(h) <= 28:
            out.append(h)
    return out[:100], "ai"


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s or "")).lower()


def score_name(name: str, title: str) -> float:
    nn = _norm(name)
    tt = _norm(title)
    if not nn or not tt:
        return 0.0
    score = 0.0
    if nn in tt:
        score += 5.0
    # bigram overlap
    if len(nn) >= 2:
        toks = [nn[i : i + 2] for i in range(len(nn) - 1)]
        hit = sum(1 for x in toks if x in tt)
        score += 3.0 * (hit / len(toks))
    ns = extract_sizes(name)
    ts = extract_sizes(title)
    if ns and ts:
        if any(x in ts for x in ns):
            score += 2.0
        else:
            score -= 1.5
    # 長い飾り入り名は減点（辞書側が短いほど良い）
    if len(name) > 28:
        score -= 0.5
    return score


def match_row(title: str, maker: str, names: list[str], jan: str, jan_map: dict[str, str]) -> tuple[str, str, str]:
    """Returns (new_name, method, note)."""
    j = str(jan or "").strip()
    if j and j in jan_map and jan_map[j]:
        return jan_map[j], "jan_unify", "JAN一致で寄せ"
    best = ""
    best_s = 0.0
    for n in names:
        s = score_name(n, title)
        if s > best_s:
            best_s = s
            best = n
    if best and best_s >= 3.0:
        return best, "dict_match", "score=%.2f" % best_s
    h = heuristic_canon(title, maker)
    if best and best_s >= 2.0:
        return best, "dict_weak", "score=%.2f fallback_heur=%s" % (best_s, h[:40])
    return h, "heuristic", "dict未ヒット"


def ensure_compare_sheet(svc, sid: str, title: str) -> None:
    meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
    titles = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if title not in titles:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [{"addSheet": {"properties": {"title": title}}}]},
        ).execute()
        return
    # clear existing
    svc.spreadsheets().values().clear(
        spreadsheetId=sid,
        range="'%s'" % title.replace("'", "''"),
    ).execute()


def write_sheet(svc, sid: str, title: str, rows: list[list[str]]) -> None:
    ensure_compare_sheet(svc, sid, title)
    # Sheets API: max ~セル数。まとめて update
    chunk = 500
    for i in range(0, len(rows), chunk):
        part = rows[i : i + chunk]
        start_row = i + 1
        rng = "'%s'!A%d" % (title.replace("'", "''"), start_row)
        svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range=rng,
            valueInputOption="RAW",
            body={"values": part},
        ).execute()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-ai", action="store_true", help="辞書AIなし（ヒューリスティックのみ）")
    ap.add_argument("--no-row-ai", action="store_true", help="行AI列を出さない")
    ap.add_argument("--no-sheet", action="store_true", help="スプシ非書（TSVのみ）")
    ap.add_argument("--all-list", action="store_true", help="リスト全ASIN（既定は見積候補のみ）")
    ap.add_argument("--quote-only", action="store_true", help="互換: 見積候補のみ（既定と同じ）")
    ap.add_argument("--sheet", default=COMPARE_SHEET)
    ap.add_argument("--sleep", type=float, default=SLEEP_SEC)
    args = ap.parse_args()
    quote_only = not args.all_list

    if RESEARCH_SS == ORIG_SS:
        print("FAIL refuse original SS")
        return 2

    svc = sheets_service(write=not args.no_sheet, interactive=True)
    rows = load_list_rows(svc)
    if not rows:
        print("FAIL no rows")
        return 1

    picked = []
    for r in rows:
        if quote_only and _col(r, "見積候補") != "見積候補":
            continue
        picked.append(r)

    by_maker: dict[str, list[dict]] = defaultdict(list)
    for r in picked:
        m = _col(r, "メーカー名", "メーカー") or "(空)"
        by_maker[m].append(r)

    print(
        "picked=%d makers=%d dict_ai=%s row_ai=%s quote_only=%s"
        % (len(picked), len(by_maker), not args.no_ai, not args.no_row_ai, quote_only)
    )

    maker_dicts: dict[str, tuple[list[str], str]] = {}
    for mi, (maker, rs) in enumerate(sorted(by_maker.items(), key=lambda x: -len(x[1])), 1):
        titles = [_col(r, "商品名") for r in rs]
        shorts = [_col(r, "単品商品名") for r in rs if _col(r, "単品商品名")]
        print("[%d/%d] dict maker=%s rows=%d shorts=%d" % (mi, len(by_maker), maker, len(rs), len(shorts)))
        names, src = build_dict_for_maker(maker, titles, shorts, use_ai=not args.no_ai)
        maker_dicts[maker] = (names, src)
        print("  -> names=%d src=%s sample=%s" % (len(names), src, (names[:3] if names else [])))
        if not args.no_ai:
            time.sleep(max(0.0, args.sleep))

    jan_map: dict[str, str] = {}
    for r in picked:
        jans = parse_jans(_col(r, "JANコード"))
        cur = _col(r, "単品商品名")
        if (
            cur
            and extract_sizes(cur)
            and len(cur) <= 28
            and not DECOR_RE.search(cur)
            and "徳用" not in cur
            and "業務用" not in cur
        ):
            for j in jans:
                if j not in jan_map or len(cur) < len(jan_map[j]):
                    jan_map[j] = cur
    for maker, rs in by_maker.items():
        names, _ = maker_dicts[maker]
        for r in rs:
            jans = parse_jans(_col(r, "JANコード"))
            title = _col(r, "商品名")
            best = ""
            best_s = 0.0
            for n in names:
                s = score_name(n, title)
                if s > best_s:
                    best_s, best = s, n
            if best and best_s >= 3.0:
                for j in jans:
                    if j not in jan_map or len(best) < len(jan_map[j]):
                        jan_map[j] = best

    out_rows: list[list[str]] = []
    empty_cur = 0
    same_cur_row = same_cur_dict = same_row_dict = 0
    n_cur = n_row = 0
    for i, r in enumerate(picked, 1):
        maker = _col(r, "メーカー名", "メーカー") or "(空)"
        title = _col(r, "商品名")
        cur = _col(r, "単品商品名")
        jan = _col(r, "JANコード")
        sale = sale_form_code(_col(r, "販売形態"))
        qty = parse_qty(_col(r, "同商品のセット数量"))
        names, src = maker_dicts.get(maker, ([], "none"))
        new_name, method, note = match_row(title, maker, names, jan, jan_map)

        if args.no_row_ai:
            row_name, row_brand, row_sku, row_memo = "", "", "", ""
        else:
            print("row_ai %d/%d %s" % (i, len(picked), _col(r, "ASIN").upper()))
            row_name, row_brand, row_sku, row_memo = row_ai_draft(title, maker, qty, sale)
            time.sleep(max(0.0, args.sleep))

        if not cur:
            empty_cur += 1
            a_cur_row = a_cur_dict = ""
        else:
            n_cur += 1
            a_cur_row = "Y" if row_name and _norm(cur) == _norm(row_name) else ("N" if row_name else "")
            a_cur_dict = "Y" if _norm(cur) == _norm(new_name) else "N"
            if a_cur_row == "Y":
                same_cur_row += 1
            if a_cur_dict == "Y":
                same_cur_dict += 1
        if row_name:
            n_row += 1
            a_row_dict = "Y" if _norm(row_name) == _norm(new_name) else "N"
            if a_row_dict == "Y":
                same_row_dict += 1
        else:
            a_row_dict = ""

        out_rows.append(
            [
                _col(r, "ASIN").upper(),
                maker,
                _col(r, "見積候補"),
                title,
                cur,
                row_name,
                row_brand,
                row_sku,
                row_memo,
                new_name,
                a_cur_row,
                a_cur_dict,
                a_row_dict,
                method + "/" + src,
                str(len(names)),
                jan,
                sale,
                str(qty),
                note,
            ]
        )

    def rate(a: int, b: int) -> float:
        return (a / b * 100.0) if b else 0.0

    summary = [
        ["比較日時", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["対象", "仕入検討品番リスト・見積候補のみ" if quote_only else "仕入検討品番リスト・全ASIN行"],
        ["行数", str(len(picked))],
        ["メーカー数", str(len(by_maker))],
        ["現状あり", str(n_cur)],
        ["現状空", str(empty_cur)],
        ["現状=行AI", "%d / %d (%.1f%%)" % (same_cur_row, n_cur, rate(same_cur_row, n_cur))],
        ["現状=辞書", "%d / %d (%.1f%%)" % (same_cur_dict, n_cur, rate(same_cur_dict, n_cur))],
        ["行AI=辞書", "%d / %d (%.1f%%)" % (same_row_dict, n_row, rate(same_row_dict, n_row))],
        ["行AIプロンプト", "金標準PASS（最少名・揃え・異種mix）PROMPT_ROW_AI.winner.txt"],
        ["PROMPT版", "quote_min_v2"],
        [],
        ["PROMPT_ROW_AI", PROMPT_ROW_AI.replace("\n", " / ")],
        ["PROMPT_DICT_AI", PROMPT_DICT_AI.replace("\n", " / ")],
        [],
    ]

    sheet_body = summary + [HEADERS] + out_rows

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    tsv = OUT_DIR / ("compare_%s.tsv" % stamp)
    with tsv.open("w", encoding="utf-8-sig", newline="") as f:
        for line in sheet_body:
            f.write("\t".join(str(x).replace("\t", " ") for x in line) + "\n")
    meta = {
        "picked": len(picked),
        "makers": len(by_maker),
        "n_cur": n_cur,
        "same_cur_row": same_cur_row,
        "same_cur_dict": same_cur_dict,
        "same_row_dict": same_row_dict,
        "n_row": n_row,
        "quote_only": quote_only,
        "tsv": str(tsv),
        "sheet": None if args.no_sheet else args.sheet,
        "ss": RESEARCH_SS,
    }
    (OUT_DIR / ("compare_%s.json" % stamp)).write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("meta", json.dumps(meta, ensure_ascii=False))

    if not args.no_sheet:
        write_sheet(svc, RESEARCH_SS, args.sheet, sheet_body)
        print("sheet_written", args.sheet, "rows", len(out_rows))
        print("https://docs.google.com/spreadsheets/d/%s/edit" % RESEARCH_SS)

    print("PASS compare done")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
