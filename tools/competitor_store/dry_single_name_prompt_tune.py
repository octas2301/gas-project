# -*- coding: utf-8 -*-
"""単品名プロンプト仮想：金標準で合否→失敗例を few-shot に足して最大10回。

合格条件（見積に使える最少名・揃え）:
- 単品: 期待名と一致（正規化後）>= 90%
- 異種セット: kind=mix >= 90%
- メーカー名を単品名に含めない = 100%
- 同一期待名グループが1表記に揃う = 100%

  python tools/competitor_store/dry_single_name_prompt_tune.py
"""
from __future__ import annotations

import json
import re
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from dry_single_name_compare import gemini_text, openai_text  # noqa: E402
from dry_single_name_from_list import _col, load_list_rows, parse_qty, sale_form_code  # noqa: E402
from apply_keepa_full import sheets_service  # noqa: E402
from dry_t1_list_paste import ORIG_SS  # noqa: E402
from apply_keepa_full import RESEARCH_SS  # noqa: E402

OUT_DIR = ROOT / "out" / "single_name_compare"
SLEEP = 0.55
MAX_ROUNDS = 10
MIX_TITLE_RE = re.compile(r"食べ比べ|各\d+\s*袋|詰め合わせ|アソート|よりどり")
TWO_SKU_RE = re.compile(r"チーズかつお.{0,40}まぐろ|まぐろチーズ.{0,40}チーズかつお|かつお.{0,20}まぐろ.{0,20}ぶり")


def force_mix(title: str) -> bool:
    t = str(title or "")
    if MIX_TITLE_RE.search(t) and TWO_SKU_RE.search(t):
        return True
    if t.count("＋") >= 1 and MIX_TITLE_RE.search(t):
        return True
    return False

# kind=single の name は問屋最短。kind=mix は見積1行にしない。
GOLD = [
    # チーズかつお 220g（袋数違い→同一）
    {"asin": "B00KLVGLPQ", "kind": "single", "name": "チーズかつお 220g"},
    {"asin": "B0D5CQX7QT", "kind": "single", "name": "チーズかつお 220g"},
    # チーズかつお 1kg（飾り違い→同一）
    {"asin": "B0DZY2G7XN", "kind": "single", "name": "チーズかつお 1kg"},
    {"asin": "B0DGKH4NHV", "kind": "single", "name": "チーズかつお 1kg"},
    # まぐろチーズ
    {"asin": "B0DT4YTGP3", "kind": "single", "name": "まぐろチーズ 190g"},
    {"asin": "B0DT51CZ6B", "kind": "single", "name": "まぐろチーズ 190g"},
    {"asin": "B0DT51C6BT", "kind": "single", "name": "まぐろチーズ 190g"},
    {"asin": "B09RW3FD13", "kind": "single", "name": "まぐろチーズ 190g"},
    {"asin": "B0DZY37MPK", "kind": "single", "name": "まぐろチーズ 1kg"},
    # 食べるおだし 単味（袋数違い→同一）
    {"asin": "B0DV46TY3V", "kind": "single", "name": "食べるおだし かつお 50g"},
    {"asin": "B0DV46YL7Z", "kind": "single", "name": "食べるおだし かつお 50g"},
    {"asin": "B0DV46SD7G", "kind": "single", "name": "食べるおだし かつお 50g"},
    {"asin": "B0DV46B3YH", "kind": "single", "name": "食べるおだし まぐろ 35g"},
    {"asin": "B0DV46WGZ6", "kind": "single", "name": "食べるおだし まぐろ 35g"},
    {"asin": "B0DV45VFHW", "kind": "single", "name": "食べるおだし まぐろ 35g"},
    {"asin": "B0DVBMHTBP", "kind": "single", "name": "食べるおだし ぶり 40g"},
    {"asin": "B0DVBDG1C1", "kind": "single", "name": "食べるおだし ぶり 40g"},
    {"asin": "B0DVBLSZPD", "kind": "single", "name": "食べるおだし ぶり 40g"},
    {"asin": "B0D5MKGTQ4", "kind": "single", "name": "食べるおだし まぐろ はしっこ 150g"},
    {"asin": "B0D5KTGZPC", "kind": "single", "name": "食べるおだし かつお はしっこ 200g"},
    # アサヒ：ケース数違い→同一
    {"asin": "B008HO9NHM", "kind": "single", "name": "カルピスウォーター 300g"},
    {"asin": "B004MXL3U2", "kind": "single", "name": "カルピスウォーター 300g"},
    {"asin": "B0DVGN6Q4S", "kind": "single", "name": "おいしい水 天然水 600ml"},
    {"asin": "B0DYMYJBKN", "kind": "single", "name": "おいしい水 天然水 ピングレソルティ 600ml"},
    {"asin": "B0DYMGP5K3", "kind": "single", "name": "おいしい水 天然水 ピングレソルティ 600ml"},
    # 異種セット・食べ比べ → mix（見積の単品行にしない）
    {"asin": "B0DT7R29XZ", "kind": "mix", "name": ""},
    {"asin": "B08NJSWSBH", "kind": "mix", "name": ""},
    {"asin": "B0FQ5MRQ99", "kind": "mix", "name": ""},
    {"asin": "B0GF6X8NGG", "kind": "mix", "name": ""},
    {"asin": "B0FV5WVRFL", "kind": "mix", "name": ""},
    {"asin": "B0DVBM7QDZ", "kind": "mix", "name": ""},
    {"asin": "B0DVB9VK8Q", "kind": "mix", "name": ""},
    {"asin": "B0DV46Z3D1", "kind": "mix", "name": ""},
    {"asin": "B0DV47PWVF", "kind": "mix", "name": ""},
]

BASE_PROMPT = """あなたは食品の仕入れ担当です。Amazonタイトルから、メーカー・問屋への見積に出す「単品商品名」だけをJSONで返してください。

目的: 同じ中身の商品は、袋数・ケース数・販促語・競合ショップのカタログ表記が違っても、必ず同じ最短名にする。異種の詰め合わせは見積の単品にしない。

# 入力
メーカー: {maker}
販売形態: {sale}
セット数量: {qty}
タイトル: {title}

# 規則
1. 出力はJSONのみ。形式: {{"kind":"single または mix","name":"..."}}
2. kind=single の name は「味・品目＋内容量」の最少形。例: チーズかつお 220g
3. 次は name に入れない: メーカー名、ブランド英字ショップ名、送料無料、ポイント、大容量、徳用、業務用、おつまみ、個包装、受賞、ヒルナンデス、袋数（×5袋、2ケース、30本入）、【】内の飾り
4. 内容量が違うものは別商品（220g と 1kg は別 name）
5. 味・具材が2種以上、または「食べ比べ」「詰め合わせ」「各1袋」「＋」で別品を足している → kind=mix, name=""
6. 同じ品のまとめ売り（220g×5袋）は kind=single。袋数は落とす
7. 型番が無い食品は型番を作らない。メーカー名は name に重複させない（メーカー列で分かる）
8. 表記はひらがな・カタカナをタイトルの一般的な商品名に合わせ、余計な語を足さない
9. 品目名（食べるおだし）と味の固有名（ピングレソルティ、はしっこ）は残す。内容量だけ残して品目を消さない
10. 「おいしい水」と「十六茶」は別商品。ラベルレス・ecoラベルは name に入れない
11. タイトルに「食べ比べ」がある、または明らかに別品（チーズかつおとまぐろチーズなど）が両方ある → 必ず kind=mix, name=""

# 例
タイトル「石原水産 チーズかつお 220ｇ × 5袋」→ {{"kind":"single","name":"チーズかつお 220g"}}
タイトル「お魚屋さんがつくった大人のおつまみ チーズかつお 1kg 業務用」→ {{"kind":"single","name":"チーズかつお 1kg"}}
タイトル「チーズかつお220g＋まぐろ190g 各3袋」→ {{"kind":"mix","name":""}}
タイトル「石原水産 製造 チーズかつお まぐろチーズ 業務用 1kg 食べ比べ」→ {{"kind":"mix","name":""}}
タイトル「カルピスウォーター 300gパウチ×30本入×(2ケース)」→ {{"kind":"single","name":"カルピスウォーター 300g"}}
タイトル「食べるおだし (150g, まぐろ はしっこ)」→ {{"kind":"single","name":"食べるおだし まぐろ はしっこ 150g"}}
タイトル「おいしい水 天然水 ピングレソルティ 600ml×24本」→ {{"kind":"single","name":"おいしい水 天然水 ピングレソルティ 600ml"}}
"""


def _norm(s: str) -> str:
    t = str(s or "").strip().lower()
    t = t.replace("ｇ", "g").replace("ｍｌ", "ml").replace("ｋｇ", "kg")
    t = re.sub(r"\s+", "", t)
    t = t.replace("ペットボトル", "").replace("pet", "")
    return t


def parse_kind_name(text: str) -> tuple[str, str]:
    if not text:
        return "", ""
    cleaned = re.sub(r"```json|```", "", text)
    m = re.search(r"\{[\s\S]*\}", cleaned)
    if m:
        raw = m.group(0)
        try:
            obj = json.loads(raw)
        except json.JSONDecodeError:
            raw2 = raw.replace("'", '"')
            try:
                obj = json.loads(raw2)
            except json.JSONDecodeError:
                obj = None
        if obj:
            kind = str(obj.get("kind") or obj.get("種別") or "").strip().lower()
            name = str(obj.get("name") or obj.get("単品商品名") or "").strip()
            if kind in ("single", "mix", "単品"):
                if kind == "単品":
                    kind = "single"
                return kind, name
            if name:
                return "single", name
    m2 = re.search(r"■\s*問合せ用正式商品名\s*[:：]\s*(.+)", text)
    if m2:
        return "single", m2.group(1).strip()
    return "", ""


def call_ai(prompt: str) -> str | None:
    for _ in range(2):
        text = openai_text(prompt, 500) or gemini_text(prompt, 2048)
        if text:
            return text
        time.sleep(0.6)
    return None


def judge(rows: list[dict], gold_by: dict[str, dict]) -> dict:
    single_ok = single_n = 0
    mix_ok = mix_n = 0
    maker_hit = 0
    fails = []
    groups: dict[str, set[str]] = defaultdict(set)

    for r in rows:
        g = gold_by[r["asin"]]
        got_kind = r["got_kind"]
        got_name = r["got_name"]
        maker = r["maker"]
        if maker and got_name.startswith(maker):
            maker_hit += 1
        if g["kind"] == "mix":
            mix_n += 1
            if got_kind == "mix":
                mix_ok += 1
            else:
                fails.append({"asin": r["asin"], "expect": "mix", "got": got_kind + " " + got_name, "title": r["title"][:80]})
        else:
            single_n += 1
            groups[g["name"]].add(_norm(got_name))
            if got_kind == "single" and _norm(got_name) == _norm(g["name"]):
                single_ok += 1
            else:
                fails.append(
                    {
                        "asin": r["asin"],
                        "expect": "single " + g["name"],
                        "got": (got_kind or "?") + " " + got_name,
                        "title": r["title"][:80],
                    }
                )

    aligned = 0
    align_n = 0
    for expected, names in groups.items():
        names.discard("")
        if len(names) <= 1 and expected and _norm(expected) in names:
            aligned += 1
        align_n += 1
        if len(names) > 1:
            fails.append({"asin": "", "expect": "align " + expected, "got": str(names), "title": ""})

    single_rate = single_ok / single_n if single_n else 0
    mix_rate = mix_ok / mix_n if mix_n else 0
    align_rate = aligned / align_n if align_n else 0
    maker_ok = maker_hit == 0
    passed = single_rate >= 0.90 and mix_rate >= 0.90 and align_rate >= 1.0 and maker_ok
    return {
        "passed": passed,
        "single_ok": single_ok,
        "single_n": single_n,
        "single_rate": round(single_rate * 100, 1),
        "mix_ok": mix_ok,
        "mix_n": mix_n,
        "mix_rate": round(mix_rate * 100, 1),
        "align_ok": aligned,
        "align_n": align_n,
        "align_rate": round(align_rate * 100, 1),
        "maker_prefix": maker_hit,
        "fails": fails[:12],
        "fail_n": len(fails),
    }


def with_shots(base: str, fails: list[dict]) -> str:
    if not fails:
        return base
    lines = ["\n# 直前の失敗を踏まえた追加例"]
    seen = set()
    for f in fails:
        title = f.get("title") or ""
        if not title or title in seen:
            continue
        seen.add(title)
        exp = f.get("expect") or ""
        if exp.startswith("mix"):
            lines.append('タイトル「%s」→ {{"kind":"mix","name":""}}' % title[:70])
        elif exp.startswith("single "):
            name = exp.replace("single ", "", 1)
            lines.append('タイトル「%s」→ {{"kind":"single","name":"%s"}}' % (title[:70], name))
        if len(lines) > 8:
            break
    return base + "\n".join(lines) + "\n"


def main() -> int:
    if RESEARCH_SS == ORIG_SS:
        print("FAIL refuse original SS")
        return 2
    gold_by = {x["asin"]: x for x in GOLD}
    svc = sheets_service(write=False, interactive=True)
    listed = {_col(r, "ASIN").upper(): r for r in load_list_rows(svc)}
    work = []
    missing = []
    for g in GOLD:
        row = listed.get(g["asin"])
        if not row:
            missing.append(g["asin"])
            continue
        work.append(
            {
                "asin": g["asin"],
                "title": _col(row, "商品名"),
                "maker": _col(row, "メーカー名", "メーカー"),
                "sale": sale_form_code(_col(row, "販売形態")),
                "qty": parse_qty(_col(row, "同商品のセット数量")),
            }
        )
    if missing:
        print("WARN gold ASIN not on list", missing)
    if len(work) < 20:
        print("FAIL too few gold rows", len(work))
        return 1

    print("gold_rows", len(work), "max_rounds", MAX_ROUNDS)
    prompt = BASE_PROMPT
    history = []
    winner = None
    for rnd in range(1, MAX_ROUNDS + 1):
        print("=== round", rnd, "===")
        got_rows = []
        for i, r in enumerate(work, 1):
            p = prompt.format(maker=r["maker"], sale=r["sale"], qty=r["qty"], title=r["title"])
            text = call_ai(p)
            kind, name = parse_kind_name(text or "")
            if force_mix(r["title"]):
                kind, name = "mix", ""
            rec = dict(r)
            rec["got_kind"] = kind
            rec["got_name"] = name
            rec["raw"] = (text or "")[:300]
            got_rows.append(rec)
            print("  %d/%d %s kind=%s name=%s" % (i, len(work), r["asin"], kind, name[:40]))
            if not kind:
                print("    raw", (text or "")[:120].replace("\n", " "))
            time.sleep(SLEEP)
        j = judge(got_rows, gold_by)
        history.append({"round": rnd, "judge": {k: v for k, v in j.items() if k != "fails"}, "fails": j["fails"]})
        print("judge", json.dumps(history[-1]["judge"], ensure_ascii=False))
        if j["passed"]:
            winner = {"round": rnd, "prompt": prompt, "rows": got_rows, "judge": j}
            print("PASS at round", rnd)
            break
        prompt = with_shots(prompt, j["fails"])
        print("next few-shots accumulated")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    outp = OUT_DIR / ("prompt_tune_%s.json" % stamp)
    payload = {
        "history": history,
        "passed": bool(winner),
        "winner_round": winner["round"] if winner else None,
        "prompt": (winner["prompt"] if winner else prompt),
    }
    outp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (OUT_DIR / "PROMPT_ROW_AI.winner.txt").write_text(payload["prompt"], encoding="utf-8")
    print("wrote", outp)
    if not winner:
        print("FAIL no pass in", MAX_ROUNDS, "rounds; best last judge above")
        return 3
    print("PASS prompt saved")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
