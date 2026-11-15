# -*- coding: utf-8 -*-
"""P-align-2b 続き: サブテーマ門の仮想。--gemini で概要2文（画像なし）。"""
from __future__ import annotations

import argparse
import json
import logging
import re
import time
from typing import Any, Dict, List, Optional, Tuple

from b5_virt_abc import LOCKED, _compact_facts_from_sheets
from sub_image_intent import TEXT_MODELS, _extract_json_obj
from sub_image_lp_themes import LP_THEMES

LOG = logging.getLogger("set_main_image.dry_sub_theme_gate")
DEFAULT_JAN = "4906283045119"
MAGURO_JAN = "4906283045317"
TARGET_MIN = 80
TARGET_MAX = 140
SOFT_MAX = 160

# 前回ロックした門。数字・効能の発明はしない。1・2は毎回。8・18は根拠があるときだけ。
GATES: Dict[int, Dict[str, Any]] = {
    1: {"mode": "always_imagine", "needles": ("面倒", "悩み", "忙しい", "手間", "とらず", "だしをとらず")},
    2: {"mode": "always", "needles": ("時短", "すぐ", "さっと", "だしをとらず", "一品", "手軽", "そのまま食べる")},
    3: {"mode": "food", "needles": ("食感", "旨", "うま", "湯気", "照り", "だし")},
    4: {"needles": ("味", "旨", "うま", "しっとり", "甘", "塩", "濃い", "味付き")},
    5: {"needles": ("開発", "秘話", "試作")},
    6: {"needles": ("国産", "産地", "原料", "かつお節", "本枯", "枕崎", "鹿児島")},
    7: {"needles": ("製法", "独自技術", "工程")},
    8: {"mode": "skip_unless", "needles": ("生産者", "職人", "漁師", "顔写真")},
    9: {"needles": ("栄養成分", "kcal")},
    10: {"needles": ("お茶漬け", "振る", "かける", "食べ方", "作り方", "ご飯", "のせる", "トッピング", "冷奴")},
    11: {"needles": ("おつまみ", "アレンジ", "レシピ", "両用")},
    12: {"needles": ("朝", "夜", "晩酌", "お茶漬け", "デスク", "食卓")},
    13: {"needles": ("ギフト", "贈", "シェア", "ご褒美")},
    # 14: 針ヒットだけでは不可。judge内で数字実績／否定文を再判定する
    14: {"needles": ("受賞", "金賞", "累計", "突破", "ランキング", "億粒", "万袋", "万枚"), "mode": "need_proof"},
    15: {"needles": ("レビュー", "口コミ", "お客様の声", "星")},
    16: {"needles": ("メディア", "雑誌", "テレビ", "TV", "推薦")},
    17: {"needles": ("無添加", "検査", "工場認証", "安心安全")},
    18: {"mode": "skip_unless", "needles": ("よくある", "Q&A", "アレルギー", "保存方法")},
    19: {"needles": ("円", "コスパ", "1食あたり", "価格")},
    20: {"needles": ("セット内容", "バリエーション", "種類一覧", "○袋セット")},
}


def _blob(materials: Dict[str, Any]) -> str:
    parts = [
        materials.get("name") or "",
        materials.get("maker") or "",
        materials.get("feats") or "",
        materials.get("kw") or "",
        materials.get("yahoo") or "",
        materials.get("rakuten") or "",
        materials.get("desc") or "",
        " ".join(materials.get("bullets") or []),
    ]
    return "\n".join(str(p) for p in parts if p)


def _hits(blob: str, needles: Tuple[str, ...]) -> List[str]:
    found = []
    for n in needles:
        if n and n in blob:
            found.append(n)
    return found


def fallback_katsuo_materials() -> Dict[str, Any]:
    # 「累計・受賞は…無い」等の否定文を入れない（T14針の誤ヒット防止）
    return {
        "name": "食べるおだし",
        "maker": "",
        "feats": "食べるおだし。だしをとらずにそのまま食べる。お茶漬け・おつまみ。さっと一振り。やわらかい食感・だしの旨み。",
        "kw": "お茶漬け おつまみ",
        "yahoo": "だしをとらずにそのまま食べる",
        "rakuten": "だしをとらずさっと一振り。お茶漬け・おつまみ。",
        "bullets": ["だしをとらずにそのまま食べる", "お茶漬け・おつまみ"],
        "desc": "食べるおだし。だしをとらずにそのまま食べる。お茶漬け・おつまみ。さっと一振り。やわらかい食感・だしの旨み。",
    }


def _has_rank_proof(blob: str) -> bool:
    """T14用。否定文中の累計/受賞は無視。数字実績か明確な受賞表現があるときだけ。"""
    if re.search(r"(累計|受賞|ランキング|販売).{0,24}(無い|ない|書くな|禁止)", blob):
        return False
    if re.search(
        r"(累計\s*\d|\d+\s*(万|億)|ランキング\s*\d|第\s*\d|万袋|万枚|億粒|金賞|受賞)",
        blob,
    ):
        return True
    return False


def judge(materials: Dict[str, Any]) -> List[Dict[str, Any]]:
    blob = _blob(materials)
    foodish = any(x in blob for x in ("おだし", "かつお", "まぐろ", "食品", "食べ", "だし"))
    rows = []
    for t in LP_THEMES:
        tid = int(t["id"])
        g = GATES[tid]
        mode = str(g.get("mode") or "need")
        needles = tuple(g.get("needles") or ())
        ev = _hits(blob, needles)
        if tid == 5 and "こだわり" in blob and "製法" in blob:
            if "こだわり+製法" not in ev:
                ev = list(ev) + ["こだわり+製法"]
        if tid == 9:
            if re.search(r"\d+\s*(kcal|キロカロリー|カロリー)", blob, re.I):
                if "数字カロリー" not in ev:
                    ev = list(ev) + ["数字カロリー"]
            elif "低カロリー" in blob or "高タンパク" in blob or "低糖質" in blob:
                ev = []
        propose = False
        kind = "skip"
        if mode == "always_imagine":
            propose = True
            kind = "fact" if ev else "imagine"
        elif mode == "always":
            propose = True
            kind = "fact" if ev else "imagine"
        elif mode == "food":
            propose = bool(foodish or ev)
            kind = "fact" if ev else ("food_ok" if propose else "skip")
        elif mode == "skip_unless":
            propose = bool(ev)
            kind = "fact" if ev else "skip_default"
        elif mode == "need_proof":
            # T14: 針＋実績証明。材料に数字/受賞が無いなら出さない
            propose = bool(ev) and _has_rank_proof(blob)
            kind = "fact" if propose else "skip"
        else:
            propose = bool(ev)
            kind = "fact" if ev else "skip"
        rows.append(
            {
                "themeId": tid,
                "phase": t["phase"],
                "name": t["name"],
                "propose": propose,
                "kind": kind,
                "phaseOrder": int(t["phaseOrder"]),
                "evidence": ev,
            }
        )
    return cap_rows(rows)


MUST_THEME_IDS = (1, 2)
EXTRA_CAP = 8
TOTAL_CAP = 10


def cap_rows(
    rows: List[Dict[str, Any]], b3_counts: Optional[Dict[int, int]] = None
) -> List[Dict[str, Any]]:
    """1・2は必須。ほか最大8。合計上限10。"""
    b3_counts = b3_counts or {}
    by_id = {int(r["themeId"]): r for r in rows}
    for tid in MUST_THEME_IDS:
        r = by_id.get(tid)
        if not r:
            continue
        r["propose"] = True
        if r.get("kind") in ("skip", "skip_default"):
            r["kind"] = "imagine"
    extras = [r for r in rows if r.get("propose") and int(r["themeId"]) not in MUST_THEME_IDS]
    extras.sort(
        key=lambda r: (
            -int(b3_counts.get(int(r["themeId"]), 0)),
            -len(r.get("evidence") or []),
            int(r.get("phaseOrder") or 9),
            int(r["themeId"]),
        )
    )
    picked: List[Dict[str, Any]] = []
    seen_po = set()
    for r in extras:
        po = int(r.get("phaseOrder") or 0)
        if po in seen_po:
            continue
        picked.append(r)
        seen_po.add(po)
        if len(picked) >= EXTRA_CAP:
            break
    if len(picked) < EXTRA_CAP:
        for r in extras:
            if r in picked:
                continue
            picked.append(r)
            if len(picked) >= EXTRA_CAP:
                break
    keep = {int(r["themeId"]) for r in picked}
    keep.update(MUST_THEME_IDS)
    for r in rows:
        if int(r["themeId"]) in MUST_THEME_IDS:
            continue
        if int(r["themeId"]) not in keep:
            r["propose"] = False
            if r.get("kind") not in ("skip", "skip_default"):
                r["kind"] = "capped"
    return rows


PROMPT_V1 = (
    "食品ECのサブ画像用。出品材料だけを根拠に、指定テーマごとに「どんな画像を作るか」の概要を書く。\n"
    "各テーマは本文2文。全体でだいたい80〜140字（160字超禁止）。\n"
    "1文目: 画面に何が写るか（構図・被写体）。2文目: 載せてよい文字・事実（材料の語だけ）。\n"
    "材料に無い賞・累計・効能・価格・認証は書くな。kind=imagine のときだけ【想像】を文頭に付け、一般的な食卓の不便／楽になる程度まで。\n"
    "数字の発明禁止。MAINと同文にしない。サブは購入動線用。\n"
    'JSONのみ: {"items":[{"themeId":1,"overview":"…"}]}\n'
)

# 変更要件＋テーマ一致強制。GAS未移植の仮想検証用。
PROMPT_V2 = (
    "食品ECのサブ画像用。購入動線のための「シーン・訴求」概要を書く。商品カタログ写真にしない。\n"
    "各テーマは本文2文。全体でだいたい80〜140字（160字超禁止）。\n"
    "1文目: そのテーマ固有のシーンと主役。2文目: そのテーマ固有の刺す訴求（材料の語だけ）。\n"
    "【最重要・テーマ一致】themeIdごとに核が違う。他テーマに見える概要は禁止。"
    "全テーマで「手元で食べるおだしを振る／盛る」の言い換えに収束させるな。構図を使い回すな。\n"
    "核の定義:\n"
    "T1悩み=不便・焦り・罪悪感が主役（訴求文字側）。人物を出すなら明るい安堵顔（苦悩顔禁止）。解決の手軽さは薄く（T2に取られるな）。\n"
    "T2ベネフィット=使ったあとの良い感情・理想の状態（笑顔・余裕・満足）。手順説明にするな。\n"
    "T3シズル=照り・断面・しっとり質感の寄り。説明は最小。"
    "常温珍味本体からの湯気禁止（温めて食べる前提に見せない）。湯気は熱いご飯／お茶漬け等があるときだけ。\n"
    "T4味の期待値調整=どんな味か分かる（味付き・だし感・塩気・しっとり等。材料語のみ）。使用シーンだけは不可。\n"
    "T6素材産地=原料・産地が見て分かる。職人ドラマ（生産者の顔）にするな。\n"
    "T10食べ方=手順が分かる（2〜3ステップ以上の工程感）。\n"
    "T11アレンジ集=同じ画で2〜3用法が分かる。1用法だけは不可。\n"
    "T12日常シーン=時間帯・場面（朝・夜・デスク等）。T2の感情未来と差を付けろ。\n"
    "T14実績=材料にある数字・ランキングのみ。無いならこのテーマは書くな（品質自慢で埋めるな）。\n"
    "【パッケージ】非必須。味・産地・表記など必要なテーマだけ。全テーマで「袋は残す」禁止。\n"
    "【T1・T2】人物（手元・横顔・後ろ姿可）推奨。必須ではない。偽UGC顔禁止。\n"
    "材料に無い賞・累計・効能・価格・認証・数字は書くな。禁忌リストを概要本文に書くな。\n"
    "kind=imagine のときだけ文頭【想像】。MAINと同文・同構図にしない。\n"
    'JSONのみ: {"items":[{"themeId":1,"overview":"…"}]}\n'
)

# テーマごとの必須核（プロンプトに再掲）
THEME_CORE: Dict[int, str] = {
    1: "悩み共感。不便が主役。解決は薄く",
    2: "理想の未来。購入後の良い感情",
    3: "シズル。照り・断面。本体湯気禁止",
    4: "味の期待値。味の輪郭。使用シーンだけ不可",
    6: "素材・産地が視認できる",
    10: "食べ方の手順（工程感）",
    11: "アレンジ2〜3用法",
    12: "日常の時間帯・場面",
    14: "実績数字のみ。無ければ書くな",
}

BAD_PHRASES_V2 = (
    "袋は残す",
    "袋は端に",
    "数字・効能の発明はしない",
    "数字の発明禁止",
)


def _build_gemini_prompt(
    materials: Dict[str, Any], hits: List[Dict[str, Any]], prompt_ver: str
) -> str:
    themes = []
    for r in hits:
        tid = int(r["themeId"])
        item = {
            "themeId": tid,
            "name": r["name"],
            "kind": r["kind"],
            "evidence": r["evidence"],
        }
        if prompt_ver == "v2":
            item["mustCore"] = THEME_CORE.get(tid, "テーマ名に厳密一致")
        themes.append(item)
    head = PROMPT_V2 if prompt_ver == "v2" else PROMPT_V1
    return (
        head
        + "\n商品材料:\n"
        + _blob(materials)[:2200]
        + "\n\n対象テーマ（mustCoreを必ず満たせ）:\n"
        + json.dumps(themes, ensure_ascii=False)
    )


def _score_overview_v2(overview: str, theme_id: int) -> Dict[str, Any]:
    """仮想採点（粗い）。人が最終判断する。"""
    ov = str(overview or "")
    flags = []
    for p in BAD_PHRASES_V2:
        if p in ov:
            flags.append("bad:" + p)
    if "袋" in ov and "パッケージ" not in ov and theme_id in (1, 2, 3, 12):
        flags.append("warn:袋語（T1/2/3/12ではパッケージ無し推奨）")
    if theme_id in (1, 2):
        if not any(x in ov for x in ("人", "手元", "横顔", "後ろ", "帰宅", "食卓に座る", "キッチン")):
            flags.append("warn:T1/T2人物シーン弱め")
    if len(ov) < TARGET_MIN:
        flags.append("short")
    if len(ov) > SOFT_MAX:
        flags.append("too_long")
    return {"flags": flags, "okish": not any(f.startswith("bad:") for f in flags)}


def _gemini_overviews(
    materials: Dict[str, Any], hits: List[Dict[str, Any]], prompt_ver: str = "v1"
) -> Dict[str, Any]:
    from gemini_image import load_api_key, make_client

    prompt = _build_gemini_prompt(materials, hits, prompt_ver)
    client = make_client(load_api_key())
    last_err = ""
    t0 = time.time()
    for mid in TEXT_MODELS:
        try:
            resp = client.models.generate_content(model=mid, contents=prompt)
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                try:
                    text = resp.candidates[0].content.parts[0].text
                except Exception:
                    text = ""
            obj = _extract_json_obj(text)
            if not obj:
                raise RuntimeError("JSON parse failed: " + text[:180])
            items = obj.get("items") if isinstance(obj, dict) else obj
            if not isinstance(items, list):
                raise RuntimeError("items missing")
            by_id = {}
            for it in items:
                try:
                    tid = int(it.get("themeId"))
                except (TypeError, ValueError, AttributeError):
                    continue
                ov = str((it or {}).get("overview") or "").strip()
                by_id[tid] = ov
            elapsed_ms = int((time.time() - t0) * 1000)
            return {"ok": True, "model": mid, "ms": elapsed_ms, "byId": by_id, "error": ""}
        except Exception as e:
            last_err = str(e)
            LOG.warning("gemini fail model=%s: %s", mid, e)
    return {"ok": False, "model": "", "ms": int((time.time() - t0) * 1000), "byId": {}, "error": last_err}


def _len_flag(n: int) -> str:
    if n < TARGET_MIN:
        return "short"
    if n > SOFT_MAX:
        return "too_long"
    if n > TARGET_MAX:
        return "longish"
    return "ok"


def run_gate(
    jan: str,
    *,
    name_base: str = "",
    parent_sku: str = "",
    gemini: bool = False,
    prompt_ver: str = "v1",
    theme_ids: str = "",
    out_path: str = "",
) -> Dict[str, Any]:
    """出品材料を商品名ベースで読み、門を通した概要を返す。画像は作らない。"""
    jan = "".join(ch for ch in str(jan) if ch.isdigit())
    name_base = str(name_base or "").strip()
    parent_sku = str(parent_sku or "").strip()
    packed: Dict[str, Any] = {}
    materials = {}
    source = "sheets"
    last_sheet_err = ""
    for attempt in range(1, 4):
        packed = _compact_facts_from_sheets(
            jan, name_base=name_base, parent_sku=parent_sku
        )
        materials = packed.get("materials") or {}
        last_sheet_err = str(packed.get("error") or "")
        if _blob(materials):
            source = "sheets" if attempt == 1 else f"sheets_retry{attempt}"
            break
        LOG.warning("sheets materials empty attempt=%s err=%s", attempt, last_sheet_err[:120])
        time.sleep(2 * attempt)
    if not _blob(materials):
        if name_base or parent_sku:
            source = "missing_name_base:" + (last_sheet_err or "empty")
        else:
            materials = fallback_katsuo_materials()
            source = "fallback_virt:" + (last_sheet_err or "empty")
    rows = judge(materials) if source and not str(source).startswith("missing_name_base") else []
    forced: List[int] = []
    if str(theme_ids or "").strip():
        for part in str(theme_ids).split(","):
            part = part.strip()
            if part.isdigit():
                forced.append(int(part))
    if forced:
        want = set(forced)
        for r in rows:
            r["propose"] = int(r["themeId"]) in want
            if r["propose"] and r.get("kind") in ("skip", "skip_default", "capped"):
                r["kind"] = "imagine" if int(r["themeId"]) in MUST_THEME_IDS else "fact"
        hit = [r for r in rows if r["propose"]]
        # 指定順を保つ
        by = {int(r["themeId"]): r for r in hit}
        hit = [by[i] for i in forced if i in by]
    else:
        hit = [r for r in rows if r["propose"]]
    skip = [r for r in rows if not r["propose"]]
    gem = {"ok": False, "byId": {}, "ms": 0, "model": "", "error": "skipped"}
    if gemini and hit:
        gem = _gemini_overviews(materials, hit, prompt_ver=prompt_ver)
        for r in hit:
            ov = gem.get("byId", {}).get(r["themeId"], "")
            n = len(ov)
            r["overview"] = ov
            r["overviewChars"] = n
            r["lenFlag"] = _len_flag(n) if ov else "empty"
            if prompt_ver == "v2":
                r["v2Score"] = _score_overview_v2(ov, int(r["themeId"]))
    char_list = [int(r.get("overviewChars") or 0) for r in hit if r.get("overview")]
    out = {
        "jan": jan,
        "promptVer": prompt_ver,
        "source": source,
        "sheetError": packed.get("error") or "",
        "materials": materials,
        "proposeIds": [r["themeId"] for r in hit],
        "skipIds": [r["themeId"] for r in skip],
        "proposeCount": len(hit),
        "tooMany": len(hit) > TOTAL_CAP,
        "forcedThemeIds": forced,
        "gemini": {
            "ok": gem.get("ok"),
            "model": gem.get("model"),
            "ms": gem.get("ms"),
            "error": gem.get("error") or "",
        },
        "length": {
            "target": f"{TARGET_MIN}-{TARGET_MAX}",
            "softMax": SOFT_MAX,
            "n": len(char_list),
            "min": min(char_list) if char_list else 0,
            "max": max(char_list) if char_list else 0,
            "avg": int(sum(char_list) / len(char_list)) if char_list else 0,
            "tooLongN": sum(1 for r in hit if r.get("lenFlag") == "too_long"),
            "longishN": sum(1 for r in hit if r.get("lenFlag") == "longish"),
            "shortN": sum(1 for r in hit if r.get("lenFlag") == "short"),
        },
        "rows": rows,
        "hitOverviews": [
            {
                "themeId": r["themeId"],
                "name": r["name"],
                "kind": r["kind"],
                "overview": r.get("overview") or "",
                "chars": r.get("overviewChars") or 0,
                "lenFlag": r.get("lenFlag") or "",
                "v2Score": r.get("v2Score"),
            }
            for r in hit
        ],
        "nameBase": name_base,
        "parentSku": parent_sku,
        "note": (
            "門＋任意Gemini。prompt-v2=変更要件仮想。"
            "1・2毎回。8・18は根拠があるときだけ。"
            "nameBase / parentSku があるときはその親行の材料だけ。"
        ),
    }
    out_path = str(out_path or "").strip()
    if not out_path:
        from pathlib import Path

        out_dir = Path(__file__).resolve().parent / "_out"
        out_dir.mkdir(parents=True, exist_ok=True)
        suffix = prompt_ver + ("_nb" if (name_base or parent_sku) else "")
        out_path = str(out_dir / f"dry_sub_theme_gate_{jan}_{suffix}.json")
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        out["savedTo"] = out_path
        LOG.info("saved %s", out_path)
    except OSError as e:
        LOG.warning("save fail: %s", e)
    return out


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="サブテーマ門の仮想。Gemini概要は --gemini")
    ap.add_argument("--jan", default=MAGURO_JAN, help="既定=まぐろ 4906283045317")
    ap.add_argument("--name-base", default="", help="商品名ベース。同じJANの単品行と混ぜない")
    ap.add_argument("--parent-sku", default="", help="親SKU。商品名ベースと併用可")
    ap.add_argument("--gemini", action="store_true", help="提案テーマだけ Gemini 2文（1回）")
    ap.add_argument(
        "--prompt-v2",
        action="store_true",
        help="変更要件プロンプト（シーン主役・パッケージ非必須・T1/T2人物推奨）",
    )
    ap.add_argument(
        "--theme-ids",
        default="",
        help="門を無視して指定IDだけ概要化（例: 1,2,3,4,6,10）。空=門の提案",
    )
    ap.add_argument(
        "--out",
        default="",
        help="結果JSONの保存先（空= tools/set_main_image/_out/dry_sub_theme_gate_{jan}_{ver}.json）",
    )
    args = ap.parse_args()
    prompt_ver = "v2" if args.prompt_v2 else "v1"
    out = run_gate(
        jan=str(args.jan),
        name_base=str(args.name_base or ""),
        parent_sku=str(args.parent_sku or ""),
        gemini=bool(args.gemini),
        prompt_ver=prompt_ver,
        theme_ids=str(args.theme_ids or ""),
        out_path=str(args.out or ""),
    )
    gem = out.get("gemini") or {}
    print(json.dumps(out, ensure_ascii=False, indent=2))
    print(
        "PROPOSE",
        out.get("proposeIds"),
        "n=",
        out.get("proposeCount"),
        "prompt=",
        prompt_ver,
        "tooMany=",
        out.get("tooMany"),
        "len",
        out.get("length"),
        "gemini_ms=",
        gem.get("ms"),
        "source=",
        out.get("source"),
        "saved=",
        out.get("savedTo"),
    )
    if prompt_ver == "v2":
        print("--- V2 OVERVIEWS ---")
        for r in out.get("hitOverviews") or []:
            sc = r.get("v2Score") or {}
            print(
                f"T{r['themeId']} [{r.get('lenFlag')}] flags={sc.get('flags')}\n  {r.get('overview') or '(empty)'}\n"
            )
    return 0 if (not args.gemini or gem.get("ok")) else 2


if __name__ == "__main__":
    raise SystemExit(main())

