# -*- coding: utf-8 -*-
"""B-④ 案画面相当の仮想（画像なし・マスタ出品CKレ点親）。

使い方:
  cd tools/set_main_image
  python dry_b5_propose_virt.py
  python dry_b5_propose_virt.py --parent-sku ishi-4906283045119-oya2
  python dry_b5_propose_virt.py --no-gemini   # ルール案のみ

出すもの（案ごと）:
  personaShort / personaDetail
  catch / salesTalk（画面＝焼成。キャッチは6〜12の1本。セールスは希望4〜10、不成文のみ〜12）
  catchBurn / salesBurn（同文。別短縮しない）
  dishName / dishSourceUrl（食品。気分・ストック用事は料理なし可）
  appealScore・試行回数（T3/T4。作り直し予算はGAS同様・実行あたり合計最大3）

合格の意味:
  レ点親の材料から件数＝レ点子SKU（2未満なら4・上限12）の案表が出ること。
  生セールス完結優先・セールス【】禁止・モール販売URL除外（S15〜S18）。
  構造ゲート（dry_b5_copy_struct_gates）とは別レーン。GASダイアログと完全一致は保証しない。
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from b5_copy_struct_gates import naive_slice_burn, zenkaku_len
from b5_virt_abc import _compact_facts_from_sheets, _jan_digits
from master_sets import _find_header_row
from sheets_master import fetch_master_rows

LOG = logging.getLogger("dry_b5_propose_virt")
ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "_out_b5_propose_virt"
OFFICIAL_URL_CACHE_PATH = ROOT / "b5_official_url_cache.json"
OFFICIAL_URL_TTL_SEC = 30 * 24 * 3600  # 30日。過ぎたら再検証／再検索

# 画面＝焼成。キャッチは6〜12。超えたら1回書き直し、まだ超えたら焼かない（途中切れ切り詰めは禁止）
CATCH_PREF_MIN = 6.0
CATCH_PREF_MAX = 12.0
CATCH_FALLBACK_MAX = CATCH_PREF_MAX
SALES_PREF_MIN = 4.0
SALES_PREF_MAX = 10.0
SALES_FALLBACK_MAX = 12.0
# 互換エイリアス（旧名）
CATCH_MIN = CATCH_PREF_MIN
CATCH_MAX = CATCH_FALLBACK_MAX
BURN_MIN = CATCH_PREF_MIN
BURN_MAX = CATCH_FALLBACK_MAX
SALES_TALK_MAX = int(SALES_FALLBACK_MAX)
SALES_BURN_MIN = int(SALES_PREF_MIN)
SALES_BURN_MAX = int(SALES_FALLBACK_MAX)
LADDER_REWRITE_MAX = 3  # GAS B5_LADDER_REWRITE_MAX_（作り直し回数。初回は別）
PROPOSE_MAX = 12

PURPOSE = [
    "時短",
    "すぐ",
    "さっと",
    "手軽",
    "一品",
    "晩酌",
    "おつまみ",
    "お茶漬け",
    "お茶請け",
    "おやつ",
    "忙しい",
    "食卓",
    "酒",
    "ビール",
    "振る",
    "かける",
    "早変わり",
    "専門店",
    "だしをとらず",
    "両用",
]
TASTE = [
    "おいしい",
    "美味し",
    "旨み",
    "旨味",
    "ジュワッ",
    "じゅわ",
    "コク",
    "香ば",
    "甘み",
    "塩気",
    "風味",
    "味わい",
    "うまい",
    "美味",
]
BAN = [
    "開ける",
    "のせる",
    "並べる",
    "混ぜる",
    "包む",
    "加熱",
    "レンジで",
    "手順",
    "体験談",
    "健康",
    "お得",
    "最強",
    "ヘルシー",
    "ダイエット",
    "血圧",
    "免疫",
]


def zenkaku_slice(s: str, max_n: float) -> str:
    out = []
    n = 0.0
    for ch in str(s or ""):
        add = 0.5 if 0x20 <= ord(ch) <= 0x7E else 1.0
        if n + add > max_n:
            break
        out.append(ch)
        n += add
    return "".join(out)


def looks_incomplete_(s: str) -> bool:
    """語・文の途中切れ（助詞止め・未完了の語幹など）。"""
    t = str(s or "").strip()
    if not t:
        return True
    if re.search(r"[・／、。（「『]$", t):
        return True
    # 助詞・接続で終わる
    if re.search(r"[てにをはがのとでもへや]$", t):
        return True
    # 「魚のお」「一品を彩」「酒が進」「晩酌が整」
    if re.search(r"(のお|をご|がお)$", t):
        return True
    # 送り仮名欠けやすい語幹（厚削／仕立／味付／海の恵）
    if re.search(r"[削燻分旨溢逸整進彩添恵立仕付]$", t) and not re.search(
        r"(厚削り|使い分け|進む|整う|彩る|溢れる|旨味|旨み|恵み|仕立て|味付け)$", t
    ):
        return True
    # 「ぶりの深い」「かつおとぶ」「深いコ」「しっとり食」など未完
    if re.search(
        r"(の深い|ぶり味|かつお味|まぐろ味|だしが|味が|深いコ|しっとり食|とぶ|まぐろとぶ)$", t
    ):
        return True
    if re.search(r"[食]$", t) and not re.search(r"(食感|食卓|食事|食品|美食)$", t):
        return True
    # 「そのまま出」「すぐ使」など動詞の途中
    if re.search(r"(そのまま出|すぐ使|手が込|味が染)$", t):
        return True
    if re.search(r"[出]$", t) and not re.search(r"(出荷|出品|出来|出発)$", t):
        return True
    # 接続で文が途中（「だしなら」）
    if re.search(r"(なら|ので|ため)$", t):
        return True
    # 文の途中で次語が切れた（「おやつに。だ」「出せる、かつ」）
    if re.search(r"[。、][ぁ-んァ-ヶ一-龥]{1,2}$", t):
        return True
    # 魚種名の途中（「かつ」「まぐ」）
    if re.search(r"(かつ|まぐ|ぶりと|おと)$", t):
        return True
    # ひらがな1字だけ（「お」等）で終わるのは途切れ扱い（終止形は除外）
    if re.search(r"[ぁ-ん]$", t) and not re.search(
        r"(る|た|だ|い|く|す|む|う|つ|ぬ|ぶ|ぐ|ん|め|よ|ね|な|か|ば|ら|り|れ|ろ|っ|ー)$", t
    ):
        return True
    return False


def zenkaku_slice_complete_(s: str, max_n: float, min_n: float = 0) -> str:
    """字数上限で切ったあと、途切れなら区切りまで戻す。無理なら空。"""
    raw = str(s or "").strip()
    if not raw:
        return ""
    if zenkaku_len(raw) <= max_n and not looks_incomplete_(raw):
        return raw
    cut = zenkaku_slice(raw, max_n)
    # 読点・句点・】の直後まで戻す
    while cut and looks_incomplete_(cut):
        cut = cut[:-1]
    # さらに助詞だらけなら、最後の区切りへ
    if looks_incomplete_(cut) or (min_n and zenkaku_len(cut) < min_n):
        for sep in ("。", "！", "？", "】", "、", "　"):
            i = cut.rfind(sep)
            if i >= 3:
                cand = cut[: i + (0 if sep == "、" else 1)].rstrip("、")
                if sep == "】":
                    cand = cut[: i + 1]
                    # 】のあとが必要ならこの候補は短すぎるので次へ
                    if zenkaku_len(cand) < (min_n or 0):
                        continue
                if not looks_incomplete_(cand) and zenkaku_len(cand) >= (min_n or 0):
                    return cand
        return ""
    if min_n and zenkaku_len(cut) < min_n:
        return ""
    return cut


# 買わせ文に継ぎ足し禁止（材料にあってもセールスへ載せない）
_BAN_IN_COPY = (
    "ヘルシー",
    "健康",
    "ダイエット",
    "糖質オフ",
    "カロリーオフ",
    "お得",
    "最強",
)


def persona_pool(mix: bool) -> List[Dict[str, str]]:
    """短文＋詳細スロット（いつどこ／誰／困り／用事／根拠／料理）。誰は年齢帯・同居・役割のいずれか。"""
    pool = [
        {
            "id": "hurry",
            "short": "忙しい日に一品足したい人",
            "errand": "平日のもう一品",
            "detail": (
                "いつどこ: 平日の夜、帰宅後の食卓\n"
                "誰: 30代・共働き夫婦・夕食の支度担当\n"
                "困り: 時間がなく品数が足りない\n"
                "得たい結果（用事）: 袋から出してすぐ、もう一品を足せる\n"
                "材料の根拠: 味付き・そのまま食べられる類の記述\n"
                "料理の方向: 皿に盛った一品／ご飯の横"
            ),
        },
        {
            "id": "drink",
            "short": "晩酌を少し華やかにしたい人",
            "errand": "晩酌のつまみ",
            "detail": (
                "いつどこ: 仕事帰りの夜、家の食卓\n"
                "誰: 40代・一人暮らしまたは夫婦・晩酌をする側\n"
                "困り: お酒に合うつまみがパッと出ない\n"
                "得たい結果（用事）: すぐ出せて席が少し華やぐ\n"
                "材料の根拠: おつまみ・珍味・晩酌系の用途\n"
                "料理の方向: 小皿に盛ったおつまみ"
            ),
        },
        {
            "id": "family",
            "short": "家族の食卓を広げたい人",
            "errand": "家族で分けて使う",
            "detail": (
                "いつどこ: 週末の食卓\n"
                "誰: 30〜40代・子どもあり世帯・食事を用意する親\n"
                "困り: 大人と子どもの好みが分かれる\n"
                "得たい結果（用事）: 同じセットから用途を分けて出せる\n"
                "材料の根拠: 2種構成・用途が複数ある記述\n"
                "料理の方向: 取り分け皿が並ぶ食卓"
            ),
        },
        {
            "id": "tea",
            "short": "お茶請け・おやつを探している人",
            "errand": "午後のお茶請け",
            "detail": (
                "いつどこ: 午後の休憩、リビング\n"
                "誰: 50代・夫婦二人暮らし・お茶を淹れる側\n"
                "困り: 甘いもの以外の軽いお茶請けが欲しい\n"
                "得たい結果（用事）: 軽くつまめてお茶が進む\n"
                "材料の根拠: おやつ・お茶請けの用途\n"
                "料理の方向: 湯飲みの横に小皿"
            ),
        },
        {
            "id": "gift",
            "short": "ちょっとした手土産を考える人",
            "errand": "手土産・差し入れ",
            "detail": (
                "いつどこ: 出向く前の準備、相手宅や職場へ\n"
                "誰: 40代・取引先や知人への挨拶役\n"
                "困り: 重くない・袋がきちんとして見えるものが欲しい\n"
                "得たい結果（用事）: 渡したときに品が良く見える\n"
                "材料の根拠: 袋・セット感が伝わる商品\n"
                "料理の方向: 袋が正面で読める置き方（料理主役にしない）"
            ),
        },
        {
            "id": "cook",
            "short": "料理のアクセントに使いたい人",
            "errand": "丼やおかずのアクセント",
            "detail": (
                "いつどこ: 平日の夕飯づくり、キッチン\n"
                "誰: 30代・家族または自炊・料理担当\n"
                "困り: いつもの丼・おかずが単調\n"
                "得たい結果（用事）: のせる／足すだけで味が変わる\n"
                "材料の根拠: トッピング・ふりかけ・具材寄りの使い方\n"
                "料理の方向: 丼や温かいおかずの上"
            ),
        },
    ]
    if mix:
        pool.append(
            {
                "id": "mix_choice",
                "short": "その日の気分で選べるセットを探す人",
                "errand": "2種から気分で選ぶ",
                "detail": (
                    "いつどこ: 買い物後のストック、家庭の棚\n"
                    "誰: 30〜40代・世帯の買い出し担当\n"
                    "困り: 味が一つだと飽きる／買い足しに出たくない\n"
                    "得たい結果（用事）: 2種入っていてその日の気分で選べる\n"
                    "材料の根拠: ミックス構成（2種以上）\n"
                    "料理の方向: 2袋が並び、片方を開けた小皿"
                ),
            }
        )
    return pool


def assign_personas(n: int, mix: bool) -> List[Dict[str, str]]:
    pool = persona_pool(mix)
    out: List[Dict[str, str]] = []
    used: Dict[str, bool] = {}
    for i in range(n):
        pick = None
        for p in pool:
            if not used.get(p["id"]):
                pick = p
                break
        if not pick:
            pick = pool[i % len(pool)]
        used[pick["id"]] = True
        out.append(dict(pick))
    return out


def _materials_blob(materials: Dict[str, Any]) -> str:
    return "\n".join(
        [
            str(materials.get("name") or ""),
            str(materials.get("maker") or ""),
            str(materials.get("feats") or ""),
            str(materials.get("kw") or ""),
            str(materials.get("yahoo") or ""),
            str(materials.get("rakuten") or ""),
            " ".join(materials.get("bullets") or []),
            str(materials.get("desc") or ""),
        ]
    )


def _slot_grounded(text: str, blob: str) -> str:
    t = re.sub(r"\s+", " ", str(text or "")).strip()
    if not t:
        return ""
    for part in re.split(r"[・、,，/／\s　]+", t):
        if len(part) >= 2 and part in blob:
            return t
    return ""


def _compact(s: str) -> str:
    return re.sub(r"[\s　・、,，/／（）()「」]+", "", str(s or ""))


def _errand_is_dish_name_only(text: str) -> bool:
    """使い方が料理名そのものだけなら落とす。文の途中の料理名は残す。"""
    t = re.sub(r"[\s　・、,，/／。．.「」『』（）()【】\[\]！!？?～〜]+", "", str(text or ""))
    if not t:
        return False
    names = sorted((a for group in _DISH_SEED_ALIASES for a in group if a), key=len, reverse=True)
    left = t
    hit = False
    for _ in range(12):
        if not left:
            break
        found = next((name for name in names if name in left), "")
        if not found:
            break
        hit = True
        left = left.replace(found, "")
    return hit and not left


def _errand_is_identity(text: str, materials: Dict[str, Any]) -> bool:
    t = _compact(text)
    if len(t) < 2:
        return True
    name = _compact(materials.get("name"))
    maker = _compact(materials.get("maker"))
    if maker and (t == maker or (maker in t and len(t.replace(maker, "")) < 2)):
        return True
    if name and (t == name or t in name or (name in t and len(t.replace(name, "")) < 2)):
        return True
    return False


def _use_facts(materials: Dict[str, Any]) -> List[str]:
    src = "\n".join(
        [
            str(materials.get("kw") or ""),
            str(materials.get("feats") or ""),
            "\n".join(materials.get("bullets") or []),
            str(materials.get("yahoo") or ""),
            str(materials.get("rakuten") or ""),
            str(materials.get("desc") or ""),
        ]
    )
    out: List[str] = []
    seen: Dict[str, bool] = {}
    for part in re.split(r"[\n。．・、,，/／|｜\s　]+", src):
        p = str(part or "").strip()
        if zenkaku_len(p) < 2 or zenkaku_len(p) > 18:
            continue
        if _errand_is_identity(p, materials):
            continue
        key = _compact(p)
        if not key or seen.get(key):
            continue
        seen[key] = True
        out.append(p)
    uses = [p for p in out if any(mark in p for mark in PURPOSE)]
    rest = [p for p in out if p not in uses]
    return (uses + rest)[:12]


def _fact_overlaps_used(fact: str, used: Dict[str, bool]) -> bool:
    c = _compact(fact)
    if not c:
        return True
    for key in used:
        u = _compact(key)
        if not u:
            continue
        if c == u:
            return True
        short, long = (c, u) if len(c) <= len(u) else (u, c)
        if len(short) >= 3 and short in long:
            return True
    return False


def _complete_persona_row(row: Dict[str, str], used: Dict[str, bool]) -> Dict[str, str]:
    fact = str(row.get("errand") or "").strip()

    def take(key: str, val: str) -> str:
        t = str(val or "").strip()
        slot = f"{key}:{t}"
        if not t or used.get(slot):
            return ""
        used[slot] = True
        return t

    who = str(row.get("who") or "")
    when = str(row.get("whenWhere") or "")
    trouble = str(row.get("trouble") or "")
    dish_dir = str(row.get("dishDir") or "")
    who_ok = who.count("・") >= 2 and zenkaku_len(who) >= 8 and fact not in who
    row["who"] = take("who", who) if who_ok else ""
    row["whenWhere"] = take("when", when) if zenkaku_len(when) >= 4 and when not in (fact, f"{fact}のとき") else ""
    row["trouble"] = take("trouble", trouble) if zenkaku_len(trouble) >= 6 and trouble not in (fact, f"{fact}が足りない") else ""
    row["evidence"] = take("ev", str(row.get("evidence") or "")) or take("ev", fact)
    row["dishDir"] = take("dish", dish_dir) if zenkaku_len(dish_dir) >= 4 and dish_dir not in (fact, f"{fact}に添える") else ""
    row["errand"] = fact
    return row


def _persona_detail(row: Dict[str, str]) -> str:
    return (
        f"いつどこ: {row.get('whenWhere') or ''}\n"
        f"誰: {row.get('who') or ''}\n"
        f"起点: {row.get('anchor') or ''}\n"
        f"困り: {row.get('trouble') or ''}\n"
        f"足場: {row.get('foothold') or ''}\n"
        f"状況: {row.get('situation') or ''}\n"
        f"得たい結果（用事）: {row.get('errand') or ''}\n"
        f"材料の根拠: {row.get('evidence') or ''}\n"
        f"料理の方向: {row.get('dishDir') or ''}"
    )


_EFFICACY_RE = re.compile(r"健康|栄養|成長|糖質|ダイエット|血圧|免疫|罪悪感|高タンパク|高たんぱく|ヘルシー")
_HANDOFF_RE = re.compile(r"インターホン|手渡|玄関|抱えて届")
_GIFT_FOOT = ("届ける", "渡す", "贈る", "手土産", "差し入れ", "贈り物", "贈答")


def _req_log(req: str, i: Any, state: str, reason: str = "", detail: str = "") -> None:
    LOG.info(
        "[b5Req_] req=%s i=%s state=%s%s%s",
        req,
        i,
        state,
        f" reason={reason}" if reason else "",
        f" detail={str(detail).replace(chr(10), ' ')[:140]}" if detail else "",
    )


def _lead_situations(materials: Dict[str, Any], proposals: List[Dict[str, Any]], client: Any) -> None:
    """足場は材料にそのままある使い方。状況は別段で、その場の物まで導く。"""
    if not proposals or client is None:
        return
    blob = materials_prompt_block(materials)
    candidates: List[str] = []
    seen: Dict[str, bool] = {}

    def add(name: str, source: str) -> None:
        s = str(name or "").strip()
        if not s or seen.get(s):
            return
        if s not in blob:
            _req_log("足場", "-", "fail", "not_in_materials", f"{source}:{s}")
            return
        if _EFFICACY_RE.fullmatch(s):
            _req_log("足場", "-", "fail", "efficacy", s)
            return
        seen[s] = True
        candidates.append(s)
        _req_log("足場", "-", "ok", source, s)

    for word in _GIFT_FOOT:
        if word in blob:
            add(word, "gift")
    listed = _gemini_array(
        "材料にそのまま書いてある使い方だけを短い語で返す。JSON配列だけ。\n"
        '[{"phrase":""}]\n'
        "材料に無い語は出さない。体への効能は出さない。\n材料:\n"
        + blob,
        client,
    )
    for row in listed:
        add(str((row or {}).get("phrase") or ""), "list")
    _req_log("別段", "-", "ok", f"foothold_n={len(candidates)}", "")
    assigned: Dict[int, str] = {}
    if candidates:
        lines = "\n".join(
            f"{i + 1}. 誰:{p.get('appealWho') or ''} 困り:{p.get('appealTrouble') or ''} 得たい結果:{errand_of(p)}"
            for i, p in enumerate(proposals)
        )
        picks = _gemini_array(
            "各案に、候補から使い方を1つ選ぶ。JSON配列だけ。\n"
            '[{"i":1,"foothold":""}]\n'
            "footholdは候補の語そのもの。合うものが無ければ空。候補に無い語は書かない。案ごとに別の語。\n"
            f"候補:{'、'.join(candidates)}\n案:\n{lines}\n",
            client,
        )
        for row in picks:
            try:
                assigned[int(row.get("i")) - 1] = str(row.get("foothold") or "").strip()
            except (TypeError, ValueError, AttributeError):
                continue
    used: Dict[str, bool] = {}
    sit_lines = []
    for i, p in enumerate(proposals):
        fh = assigned.get(i) or ""
        if fh not in candidates or used.get(fh):
            _req_log("足場", i, "fail", "not_in_candidates" if fh else "empty", fh or "-")
            fh = ""
        else:
            used[fh] = True
            _req_log("足場", i, "ok", "assigned", fh)
        p["situation"] = ""
        if not str(p.get("appealWho") or "").strip() or not errand_of(p):
            if fh:
                used.pop(fh, None)
            p["foothold"] = ""
            _req_log("買う人", i, "fail", "empty" if not str(p.get("appealWho") or "").strip() else "no_errand", "状況は書かない")
            continue
        p["foothold"] = fh
        if fh:
            sit_lines.append(
                f"{i + 1}. 足場:{fh} 誰:{p.get('appealWho') or ''} 困り:{p.get('appealTrouble') or ''} 得たい結果:{errand_of(p)}"
            )
    if not sit_lines:
        return
    wrote = _gemini_array(
        "各案の状況を、足場が決まったあとで書く。JSON配列だけ。\n"
        '[{"i":1,"situation":"","linkWhy":""}]\n'
        "linkWhyは、足場のどの語を、その人のどの場に結び付けたか。\n"
        "足場の語を文に残す。主語は、ネットの店で注文するその人。その人の場が見えるところまで延ばす。その場に普通にある物は、材料に名詞が無くても書いてよい。"
        "書いてよい物は、その足場の場の物だけ。別の場の物は書かない。玄関で渡す、インターホンを押す、箱を抱えて届ける、は書かない。"
        "体の効能をうたう語は書かない。健康、栄養、成長、糖質、ダイエット、血圧、免疫、罪悪感、高タンパク、ヘルシーは例で、同じ種類も書かない。\n"
        + "\n".join(sit_lines)
        + "\n",
        client,
    )
    by_i = {}
    for row in wrote:
        try:
            by_i[int(row.get("i")) - 1] = {
                "situation": str(row.get("situation") or "").strip(),
                "linkWhy": str(row.get("linkWhy") or "").replace("\n", " ")[:180],
            }
        except (TypeError, ValueError, AttributeError):
            continue
    for i, p in enumerate(proposals):
        fh = str(p.get("foothold") or "")
        if not fh:
            continue
        slot = by_i.get(i) or {}
        text = str(slot.get("situation") or "") if isinstance(slot, dict) else str(slot or "")
        if isinstance(slot, dict) and slot.get("linkWhy"):
            p["situationLinkWhy"] = slot.get("linkWhy")
        hit = _EFFICACY_RE.search(text)
        hand = _HANDOFF_RE.search(text)
        if hit or hand or fh not in text or not text:
            why = (
                f"efficacy:{hit.group(0)}"
                if hit
                else f"handoff:{hand.group(0)}"
                if hand
                else "foothold_missing"
            )
            _req_log("状況", i, "fail", why, text or "-")
            p["situation"] = ""
        else:
            p["situation"] = text
            _req_log("状況", i, "ok", "", text)
            _req_log("場の物", i, "note", "", text)
        p["personaDetail"] = "訴求方針: " + str(p.get("appealPolicy") or "") + "\n" + _persona_detail(
            {
                "who": p.get("appealWho") or "",
                "trouble": p.get("appealTrouble") or "",
                "foothold": fh,
                "situation": p.get("situation") or "",
                "errand": errand_of(p),
                "whyAppeal": p.get("whyAppeal") or "",
            }
        )


def _persona_map(parsed: List[Dict[str, Any]], expect: List[int]) -> Dict[int, Dict[str, Any]]:
    return {expect[pos]: src0 for pos, src0 in enumerate(parsed) if pos < len(expect)}


def fill_personas_from_materials(
    materials: Dict[str, Any],
    proposals: List[Dict[str, Any]],
    *,
    client: Any,
    use_gemini: bool,
) -> None:
    """買う人を先に書く。使い方はそのあと。固定の一覧は使わない。"""
    if not proposals:
        return
    parsed: List[Dict[str, Any]] = []
    if use_gemini and client is not None:
        lines = "\n".join(
            f"{i + 1}. {p.get('themeId') or ''} {p.get('themeName') or ''}"
            for i, p in enumerate(proposals)
        )
        prompt = (
            "提案画面のペルソナ。次のJSON配列だけ返す。件数はちょうど"
            f"{len(proposals)}。説明は付けない。\n"
            '[{"anchor":"","whenWhere":"","who":"","evidence":"","trouble":"","errand":"","whyAppeal":"","dishDir":"","useClass":"table"}]\n'
            "案ごとに、材料にある売り方か使い道を1つ選び、anchorにその語を書く。その語を起点に、ネットで注文する人・困り・得たい結果・理由・区分を一つの話にする。anchorは材料にある語。材料に無い語は書かない。\n"
            "買う人を先に書く。使い方はそのあと。飛ばさない。状況はこの返答に書かない。\n"
            "1. whoは、この商品をネットの店で注文する人。年齢の幅・世帯・役目を「・」で3つつなぐ。役目は、画面で注文する役目。食べているだけの人、届いた箱を受け取るだけの人、店の人は書かない。材料に年齢が無くても、商品の売り方から想像して書く。原料名・商品名・料理名は入れない。用意した人物の一覧は使わない。案ごとに別人。\n"
            "2. whenWhereは、その人が注文を決めるとき。食べている最中は書かない。玄関で渡す、インターホンを押す、箱を抱えて届ける、は書かない。\n"
            "3. errandは得たい結果。その困りが終わる結果にする。材料・販売文の使い方は、その困りに合うときだけ選ぶ。足りない案だけ、その困りから結果を短い文で書く。原料名・料理名だけは禁止。他の案と同じ文は禁止。体の効能をうたう語は書かない。健康、栄養、成長、糖質、ダイエット、血圧、免疫、罪悪感、高タンパク、ヘルシーは例で、同じ種類も書かない。\n"
            "4. evidenceは、材料から選んだときだけその語。困りから書いたときは空。dishDirは食べ方の方向だけ。完成料理名は書かない。\n"
            "5. 他の案と同じ who・whenWhere・trouble は禁止。\n"
            "6. whyAppealは、その得たい結果が困りを終わらせる説明だけ。1文。\n"
            "7. 困り・得たい結果・訴求理由は、欄へ書く前に一つの流れで考える。順は、困り、それが終わる結果、その説明。考えた流れの文は返さない。根拠の次に、この3欄を続ける。料理の方向と区分は、その後ろ。troubleは、その人が注文する理由。原料名は禁止。得たい結果は1つ。二つを同時に書かない。贈り物の人に、毎日の食事の結果を書かない。\n"
            "8. useClassは gift / stock / table のどれか。贈答用にネットで注文する人は gift。買い置きや気分で注文する人は stock。自分や家族の食卓のために注文する人は table。渡す動作では分けない。\n"
            "材料:\n"
            f"{materials_prompt_block(materials)}\n"
            "案:\n"
            f"{lines}\n"
        )
        raw = gemini_text(prompt, client, thinking="minimal")
        parsed = _persona_parsed(raw)
    by_i = _persona_map(parsed, list(range(len(proposals))))
    used: Dict[str, bool] = {}
    used_slots: Dict[str, bool] = {}
    blob = materials_prompt_block(materials)
    for i, p in enumerate(proposals):
        src = by_i.get(i) or {}
        p["personaWhoBefore"] = str(src.get("who") or "").strip()
        p["personaErrBefore"] = str(src.get("errand") or "").strip()
        errand = str(src.get("errand") or "").strip()
        if (
            _errand_is_identity(errand, materials)
            or _errand_is_dish_name_only(errand)
            or _fact_overlaps_used(errand, used)
        ):
            errand = ""
        if errand:
            used[errand] = True
        row = {
            "whenWhere": str(src.get("whenWhere") or "").strip(),
            "who": str(src.get("who") or "").strip(),
            "trouble": str(src.get("trouble") or "").strip(),
            "foothold": "",
            "situation": "",
            "errand": errand,
            "evidence": str(src.get("evidence") or "").strip(),
            "dishDir": str(src.get("dishDir") or "").strip(),
        }
        _complete_persona_row(row, used_slots)
        policy = f"誰に:{row.get('who') or ''}／何を:{row.get('errand') or ''}"
        p["personaId"] = f"m{i}"
        p["personaShort"] = str(row.get("who") or "")
        p["personaDetail"] = "訴求方針: " + policy + "\n" + _persona_detail(row)
        p["errand"] = errand
        p["personaErrand"] = errand
        p["appealWho"] = row.get("who") or ""
        p["appealWhat"] = row.get("errand") or ""
        p["appealPolicy"] = policy
        p["appealTrouble"] = str(row.get("trouble") or "")
        p["situation"] = ""
        p["foothold"] = ""
        p["whyAppeal"] = str(src.get("whyAppeal") or "").strip()
        anchor = str(src.get("anchor") or "").strip()
        p["anchor"] = anchor if anchor and anchor in blob else ""
        if anchor and not p["anchor"]:
            _req_log("起点", i, "fail", "not_in_materials", anchor)
        p["useClass"] = str(src.get("useClass") or "").strip().lower()
        if p["useClass"] not in ("gift", "stock", "table"):
            p["useClass"] = ""
        row["anchor"] = p["anchor"]
        p["personaDetail"] = "訴求方針: " + policy + "\n" + _persona_detail(row)
    missing = [i for i, p in enumerate(proposals) if not p.get("appealWho") or not errand_of(p) or not p.get("useClass")]
    if missing and use_gemini and client is not None:
        lines2 = "\n".join(
            f"{i + 1}. {proposals[i].get('themeId') or ''} {proposals[i].get('themeName') or ''}" for i in missing
        )
        prompt2 = prompt.replace(f"件数はちょうど{len(proposals)}", f"件数はちょうど{len(missing)}", 1)
        prompt2 = prompt2.split("案:\n", 1)[0] + "案:\n" + lines2 + "\n"
        raw2 = gemini_text(prompt2, client, thinking="minimal")
        parsed2 = _persona_parsed(raw2)
        by2 = _persona_map(parsed2, missing)
        LOG.info("persona_keep ask2 back=%s/%s", len(parsed2), len(missing))
        for i in missing:
            src = by2.get(i) or {}
            proposals[i]["personaWhoBefore"] = str(src.get("who") or "").strip()
            proposals[i]["personaErrBefore"] = str(src.get("errand") or "").strip()
            errand = str(src.get("errand") or "").strip()
            if _errand_is_identity(errand, materials) or _errand_is_dish_name_only(errand):
                errand = ""
            row = {
                "whenWhere": str(src.get("whenWhere") or "").strip(),
                "who": str(src.get("who") or "").strip(),
                "trouble": str(src.get("trouble") or "").strip(),
                "foothold": "",
                "situation": "",
                "errand": errand,
                "evidence": str(src.get("evidence") or "").strip(),
                "dishDir": str(src.get("dishDir") or "").strip(),
            }
            _complete_persona_row(row, {})
            anchor = str(src.get("anchor") or "").strip()
            anchor = anchor if anchor and anchor in blob else ""
            row["anchor"] = anchor
            p = proposals[i]
            policy = f"誰に:{row.get('who') or ''}／何を:{row.get('errand') or ''}"
            p["personaShort"] = str(row.get("who") or "")
            p["personaDetail"] = "訴求方針: " + policy + "\n" + _persona_detail(row)
            p["errand"] = row.get("errand") or ""
            p["personaErrand"] = p["errand"]
            p["appealWho"] = row.get("who") or ""
            p["appealWhat"] = p["errand"]
            p["appealPolicy"] = policy
            p["appealTrouble"] = str(row.get("trouble") or "")
            p["whyAppeal"] = str(src.get("whyAppeal") or "").strip()
            p["anchor"] = anchor
            p["useClass"] = str(src.get("useClass") or "").strip().lower()
            if p["useClass"] not in ("gift", "stock", "table"):
                p["useClass"] = ""
    for i, p in enumerate(proposals):
        if not p.get("appealWho") or not errand_of(p):
            _req_log("買う人", i, "fail", "empty" if not p.get("appealWho") else "no_errand", "困りと得たい結果と理由を空にする")
            p["appealTrouble"] = ""
            p["errand"] = ""
            p["personaErrand"] = ""
            p["appealWhat"] = ""
            p["whyAppeal"] = ""
    LOG.info("persona_keep ask1 back=%s/%s", len(parsed), len(proposals))
    for i, p in enumerate(proposals):
        LOG.info(
            "persona_keep i=%s whoBefore=%s errBefore=%s who=%s class=%s",
            i,
            str(p.get("personaWhoBefore") or "")[:80],
            str(p.get("personaErrBefore") or "")[:80],
            str(p.get("appealWho") or "")[:40],
            p.get("useClass") or "-",
        )
    if use_gemini and client is not None:
        _lead_situations(materials, proposals, client)
        _one_errand_align(materials, proposals, client)
        _rewrite_errand_claims(materials, proposals, client)


def strip_ban_copy(s: str) -> str:
    t = str(s or "")
    for b in _BAN_IN_COPY:
        t = t.replace(b, "")
    t = re.sub(r"[・／/]{2,}", "・", t)
    t = re.sub(r"^[・／、。．\s]+|[・／、。．\s]+$", "", t)
    return t.strip()


def errand_of(p: Dict[str, Any]) -> str:
    return str(p.get("errand") or p.get("personaErrand") or "").strip()


def build_sales_burn_paired(sales_talk: str, errand: str, catch_burn: str) -> str:
    """セールス画像用。字数切りではなく意味単位。キャッチ画像用と役割が被らない。"""
    t = strip_ban_copy(re.sub(r"\s+", "", str(sales_talk or "")))
    if not t:
        t = "すぐ出せる一品" if "一品" in errand else (errand[:8] if errand else "選ぶ理由が分かる")
    # キャッチ画像用の見出し語だけをセールスにしない
    m = re.search(r"【([^】]+)】", str(catch_burn or ""))
    if m and t == m.group(1):
        t = "袋からすぐ" if "すぐ" not in t else "食卓に足せる"
    out = build_sales_burn(t)
    out = strip_ban_copy(out)
    if not (SALES_BURN_MIN <= zenkaku_len(out) <= SALES_BURN_MAX):
        out = zenkaku_slice(out, SALES_BURN_MAX)
        out = re.sub(r"[てにをはがのとでもへや・／、。]$", "", out)
    return out


# 料理名として弾く（個品見本・撮影指示・抽象）
_DISH_BAN_RE = re.compile(
    r"(小皿|切り身盛り|厚削り盛り|削り節盛り|袋が正面|置き方|おともに$|一品$|おつまみ$|"
    r"取り分け|湯飲みの横|2袋|気分で選ぶ|ご飯の横$|おかずの上$)"
)

# 場面の型（案間で使い回さない）
_SCENE_LAYOUTS = (
    "右下に個品見本の小皿、左に添える料理",
    "中央に添える料理、手前右に個品見本の小皿",
    "奥に添える料理、手前に個品見本の小皿",
    "木のトレー上。個品見本と添える料理を左右に",
    "食卓の俯瞰。個品見本の小皿と添える料理が並ぶ",
    "和食器の一角。個品見本は小皿、添える料理は丼または茶碗",
    "晩酌の卓。個品見本の小皿と添える料理、酒器は小さく",
)


_SCENE_DRAW_RULES = (
    "商品の見本を前に置く。包材は見本の後ろ。袋を並べて主役にしない。"
    "組み合わせは袋と個品の並び順をそろえる。顔と手渡す手元は出さない。"
    "描く物はその案の状況の場だけ。選んだ食べ方と別の食べ物は足さない。"
    "状況の文は画像の文字にしない。"
)


def _append_scene_draw_rules(scene: str) -> str:
    s = str(scene or "").strip()
    if "見本を前" in s:
        return s
    if not s:
        return _SCENE_DRAW_RULES
    return s.rstrip("。") + "。" + _SCENE_DRAW_RULES


def _strip_catch_brackets(text: str) -> str:
    t = re.sub(r"【([^】]*)】", r"\1", str(text or ""))
    return re.sub(r"[【】]", "", t).strip()


def _strip_sales_brackets(text: str) -> str:
    return re.sub(r"【[^】]*】", "", str(text or "")).strip()


def scene_from_persona(p: Dict[str, Any], materials: Dict[str, Any]) -> str:
    """フォールバック。状況の場に商品を置く。手渡す手元は出さない。"""
    dish = str(p.get("dishName") or "").strip()
    if dish and _DISH_BAN_RE.search(dish):
        dish = ""
    situation = str(p.get("situation") or "").strip()
    head = f"場は{situation}。" if situation else "商品そのもの。"
    dish_bit = f"添える料理は{dish}。" if dish else ""
    box = "包んだ箱を商品の横に置く。" if proposal_is_gift_(p) else ""
    return _append_scene_draw_rules(head + dish_bit + box)


def product_match_tokens_(materials: Dict[str, Any]) -> List[str]:
    """URLページに載っているべき商品の手がかり（短い語）。"""
    name = str(materials.get("name") or "").strip()
    toks: List[str] = []
    # 括弧前のシリーズ名
    head = re.split(r"[（(]", name, 1)[0].strip()
    if head and len(head) >= 3:
        toks.append(head)
    if "食べるおだし" in name:
        toks.append("食べるおだし")
    if "石原" in name or "ishi" in str(materials.get("parentSku") or "").lower():
        toks.append("石原")
    for m in list(materials.get("mixMarks") or [])[:2]:
        if m and m not in toks:
            toks.append(str(m))
    # 重複除去・空除去
    out: List[str] = []
    for t in toks:
        t = str(t).strip()
        if t and t not in out:
            out.append(t)
    return out


def is_weak_recipe_source_url_(url: str) -> bool:
    """モール販売ページ等。食べ方の根拠URLにしない（公式以外）。"""
    u = str(url or "").strip().lower()
    if not u:
        return False
    bad = (
        "store.shopping.yahoo.co.jp",
        "shopping.yahoo.co.jp",
        "paypayfleamarket.yahoo",
        "item.rakuten.co.jp",
        "www.amazon.co.jp",
        "amazon.co.jp/",
        "auctions.yahoo.co.jp",
        "page.auctions.yahoo",
    )
    return any(b in u for b in bad)


def is_media_article_url_(url: str) -> bool:
    """雑誌・メディア記事（メーカー公式ではない）。公式URL解決では後回し。"""
    u = str(url or "").strip().lower()
    if not u:
        return False
    hints = (
        "hpplus.jp",
        "baila.",
        "note.com",
        "ameblo.jp",
        "getnavi.jp",
        "rocketnews",
        "youtube.",
        "instagram.",
        "twitter.com",
        "x.com/",
        "facebook.com",
    )
    return any(h in u for h in hints)


def dish_optional_for_errand_(errand: str) -> bool:
    """ストックか、届ける用事。添える料理なし。"""
    return stock_errand_(errand) or gift_errand_(errand)


def stock_errand_(errand: str) -> bool:
    return bool(re.search(r"ストック|気分で選|常備|棚に|買い置き", str(errand or "")))


def gift_errand_(errand: str) -> bool:
    return bool(re.search(r"届ける|渡す|贈る|喜ばれる|手土産|差し入れ|贈り物|贈答", str(errand or "")))


def person_blob_(proposal: Dict[str, Any]) -> str:
    proposal = proposal or {}
    return "\n".join(
        str(proposal.get(k) or "")
        for k in ("appealWho", "personaShort", "appealTrouble", "errand", "appealPolicy", "whyAppeal")
    )


def use_class_of(proposal: Dict[str, Any]) -> str:
    s = str((proposal or {}).get("useClass") or "").strip().lower()
    return s if s in ("gift", "stock", "table") else ""


def proposal_is_gift_(proposal: Dict[str, Any]) -> bool:
    c = use_class_of(proposal)
    if c == "gift":
        return True
    if c in ("stock", "table"):
        return False
    return gift_errand_(person_blob_(proposal) + "\n" + errand_of(proposal))


def meal_errand_(errand: str) -> bool:
    return bool(re.search(r"晩酌|食事|一皿|食卓|朝食|昼食|夕食|夜食|軽食|つまみ|肴", str(errand or "")))


def proposal_is_stock_(proposal: Dict[str, Any]) -> bool:
    c = use_class_of(proposal)
    if c == "stock":
        err = errand_of(proposal)
        if stock_errand_(err):
            return True
        if meal_errand_(err):
            LOG.info("use_class_ stock ignored meal")
            return False
        return True
    if c in ("gift", "table"):
        return False
    return stock_errand_(person_blob_(proposal) + "\n" + errand_of(proposal))


def dish_optional_for_proposal_(proposal: Dict[str, Any]) -> bool:
    return proposal_is_gift_(proposal) or proposal_is_stock_(proposal)


def _gemini_array(prompt: str, client: Any) -> List[Dict[str, Any]]:
    raw = gemini_text(prompt, client)
    m = re.search(r"\[[\s\S]*\]", raw or "")
    if not m:
        return []
    try:
        got = json.loads(m.group(0))
    except json.JSONDecodeError:
        return []
    return [x for x in got if isinstance(x, dict)] if isinstance(got, list) else []


def _gemini_obj(prompt: str, client: Any) -> Optional[Dict[str, Any]]:
    raw = gemini_text(prompt, client)
    m = re.search(r"\{[\s\S]*\}", raw or "")
    if not m:
        return None
    try:
        got = json.loads(m.group(0))
    except json.JSONDecodeError:
        return None
    return got if isinstance(got, dict) else None


def remake_align_ok_(parent: str, ingredient: str, method: str, dish_name: str, client: Any) -> bool:
    if not parent or not ingredient or not method or not dish_name or client is None:
        return False
    got = _gemini_obj(
        "親の料理と、作り直した料理が同じ主材料かを返す。JSONだけ。\n"
        '{"sameIngredient":true,"nameMatches":true}\n'
        "sameIngredientは、書いた主材料が親の料理の主材料と同じときだけ true。\n"
        "nameMatchesは、料理名がその主材料と、変えた作り方の組み合わせのときだけ true。用途の言い換えだけは false。\n"
        f"親:{parent}\n主材料:{ingredient}\n作り方:{method}\n料理名:{dish_name}\n",
        client,
    ) or {}
    ok = got.get("sameIngredient") is True and got.get("nameMatches") is True
    LOG.info("remake_align_ok_ parent=%s ingredient=%s dish=%s ok=%s", parent, ingredient, dish_name, ok)
    return ok


def scene_uses_ingredient_(ingredient: str, scene: str, client: Any) -> bool:
    if not ingredient:
        return True
    if client is None:
        return False
    got = _gemini_obj(
        "場面が、指定の主材料を描いているかを返す。JSONだけ。\n"
        '{"ok":true}\n'
        "別の食材なら false。主材料が場面に無いときも false。\n"
        f"主材料:{ingredient}\n場面:{scene or ''}\n",
        client,
    ) or {}
    ok = got.get("ok") is True
    LOG.info("scene_uses_ingredient_ ingredient=%s ok=%s", ingredient, ok)
    return ok


def _write_appeal(p: Dict[str, Any], trouble: str, errand: str, why: str) -> None:
    p["appealTrouble"] = str(trouble or "")
    p["errand"] = str(errand or "")
    p["personaErrand"] = p["errand"]
    p["appealWhat"] = p["errand"]
    p["whyAppeal"] = str(why or "")
    p["appealPolicy"] = f"誰に:{p.get('appealWho') or ''}／何を:{p['errand']}"
    p["personaDetail"] = "訴求方針: " + p["appealPolicy"] + "\n" + _persona_detail(
        {
            "who": p.get("appealWho") or "",
            "anchor": p.get("anchor") or "",
            "trouble": p.get("appealTrouble") or "",
            "situation": p.get("situation") or "",
            "errand": p.get("errand") or "",
            "whyAppeal": p.get("whyAppeal") or "",
        }
    )


def _rewrite_errand_claims(materials: Dict[str, Any], proposals: List[Dict[str, Any]], client: Any) -> None:
    if not proposals or client is None:
        return
    idx = []
    lines = []
    for i, p in enumerate(proposals):
        err = errand_of(p)
        hit = _EFFICACY_RE.search(err or "")
        if not hit or not p.get("appealWho"):
            continue
        idx.append(i)
        lines.append(
            f"{i + 1}. 誰:{p.get('appealWho') or ''} 困り:{p.get('appealTrouble') or ''} 得たい結果:{err}"
        )
        _req_log("得たい結果", i, "fail", "efficacy:" + hit.group(0), err)
    if not idx:
        return
    fixed = _gemini_array(
        "次の案だけ、得たい結果を書き直す。JSON配列だけ。\n"
        '[{"i":1,"errand":""}]\n'
        "困りが終わる結果は1つ。体の効能をうたう語は書かない。健康、栄養、成長、糖質、ダイエット、血圧、免疫、罪悪感、高タンパク、ヘルシーは例で、同じ種類も書かない。状況は返さない。\n"
        f"材料:\n{materials_prompt_block(materials)}\n案:\n" + "\n".join(lines) + "\n",
        client,
    )
    by_i = {}
    for row in fixed:
        try:
            by_i[int(row.get("i")) - 1] = row
        except (TypeError, ValueError):
            continue
    for i in idx:
        row = by_i.get(i) or {}
        nxt = str(row.get("errand") or "").strip()
        if nxt:
            proposals[i]["errand"] = nxt
            proposals[i]["personaErrand"] = nxt
            proposals[i]["appealWhat"] = nxt
            who = proposals[i].get("appealWho") or ""
            proposals[i]["appealPolicy"] = f"誰に:{who}／何を:{nxt}"
        hit2 = _EFFICACY_RE.search(errand_of(proposals[i]) or "")
        if hit2:
            _req_log("得たい結果", i, "note", "keep:" + hit2.group(0), errand_of(proposals[i]))
        else:
            _req_log("得たい結果", i, "ok", "retry", errand_of(proposals[i]))


def _one_errand_align(materials: Dict[str, Any], proposals: List[Dict[str, Any]], client: Any) -> None:
    if not proposals or client is None:
        return
    kept = []
    for i, p in enumerate(proposals):
        if not p.get("appealWho") or not errand_of(p):
            _req_log("買う人", i, "fail", "empty" if not p.get("appealWho") else "no_errand", "書き直しに入れない")
            continue
        kept.append(i)
    if not kept:
        return
    lines = "\n".join(
        f"{i + 1}. 起点:{proposals[i].get('anchor') or ''} 誰:{proposals[i].get('appealWho') or ''} 困り:{proposals[i].get('appealTrouble') or ''} 得たい結果:{errand_of(proposals[i])} 訴求理由:{proposals[i].get('whyAppeal') or ''}"
        for i in kept
    )
    got = _gemini_array(
        "各案の得たい結果の数と、3欄が同じ話かを返す。JSON配列だけ。\n"
        '[{"i":1,"n":1,"same":true}]\n'
        "nは得たい結果の数。sameは困り・得たい結果・訴求理由が同じ話なら true。\n"
        f"案:\n{lines}\n",
        client,
    )
    if not got:
        return
    bad = set()
    for row in got:
        try:
            idx = int(row.get("i")) - 1
        except (TypeError, ValueError):
            continue
        same = row.get("same") is True
        try:
            n = int(row.get("n"))
        except (TypeError, ValueError):
            n = 0
        if not (n == 1 and same):
            bad.add(idx)
    if not bad:
        return
    fix_lines = "\n".join(
        f"{i + 1}. 起点:{proposals[i].get('anchor') or ''} 誰:{proposals[i].get('appealWho') or ''} 困り:{proposals[i].get('appealTrouble') or ''} 得たい結果:{errand_of(proposals[i])} 訴求理由:{proposals[i].get('whyAppeal') or ''}"
        for i in sorted(bad)
        if 0 <= i < len(proposals) and proposals[i].get("appealWho") and errand_of(proposals[i])
    )
    fixed = _gemini_array(
        "次の案だけ、困り・得たい結果・訴求理由を同じ話の1つに書き直す。JSON配列だけ。\n"
        '[{"i":1,"trouble":"","errand":"","whyAppeal":""}]\n'
        "得たい結果は1つ。状況は返さない。変えない。起点と誰の、ネットで注文する話にする。材料に無い商品の話は書かない。\n"
        f"材料:\n{materials_prompt_block(materials)}\n案:\n{fix_lines}\n",
        client,
    )
    by_i = {}
    for row in fixed:
        try:
            by_i[int(row.get("i")) - 1] = row
        except (TypeError, ValueError):
            continue
    for i in sorted(bad):
        if i not in by_i or not (0 <= i < len(proposals)):
            continue
        row = by_i[i]
        _write_appeal(proposals[i], str(row.get("trouble") or ""), str(row.get("errand") or ""), str(row.get("whyAppeal") or ""))
    verify = "\n".join(
        f"{i + 1}. 困り:{proposals[i].get('appealTrouble') or ''} 得たい結果:{errand_of(proposals[i])} 訴求理由:{proposals[i].get('whyAppeal') or ''}"
        for i in sorted(bad)
        if 0 <= i < len(proposals)
    )
    again = _gemini_array(
        "各案の得たい結果の数と、3欄が同じ話かを返す。JSON配列だけ。\n"
        '[{"i":1,"n":1,"same":true}]\n'
        "nは得たい結果の数。sameは困り・得たい結果・訴求理由が同じ話なら true。\n"
        f"案:\n{verify}\n",
        client,
    )
    still = set()
    for row in again:
        try:
            idx = int(row.get("i")) - 1
            n = int(row.get("n"))
        except (TypeError, ValueError):
            continue
        if not (n == 1 and row.get("same") is True):
            still.add(idx)
    for i in sorted(bad):
        if not (0 <= i < len(proposals)):
            continue
        if i not in by_i or not again or i in still or not errand_of(proposals[i]):
            if errand_of(proposals[i]):
                LOG.info("one_errand_ i=%s keep reason=two", i)
                proposals[i]["personaDrop"] = (
                    (str(proposals[i].get("personaDrop") or "") + "+two").lstrip("+")
                    if proposals[i].get("personaDrop")
                    else "two"
                )
                _req_log("買う人", i, "note", "two", "揃わないが残す")
            else:
                LOG.info("one_errand_ i=%s empty reason=two", i)
                _write_appeal(proposals[i], "", "", "")
        else:
            LOG.info("one_errand_ i=%s n=1 after rewrite", i)


def url_page_mentions_product_(url: str, tokens: List[str], *, timeout: float = 8.0) -> bool:
    """ページ本文に商品手がかりが1つ以上あるか（別商品URLを落とす）。"""
    if not url or not tokens:
        return False
    if is_weak_recipe_source_url_(url):
        return False
    try:
        import urllib.request

        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; B5Virt/1.0)"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read(400_000)
        html = raw.decode("utf-8", "replace")
        # script大半は無視してテキスト寄りに
        html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
        html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
        # 必須: シリーズ名があればそれを優先。無ければトークンいずれか
        strong = [t for t in tokens if t in ("食べるおだし",) or "おだし" in t]
        if strong:
            return any(t in html for t in strong)
        return any(t in html for t in tokens)
    except Exception as e:
        LOG.warning("url_page_mentions_product_ fail url=%s err=%s", url[:80], e)
        return False


# 材料／公式ページから拾う食べ方の種（一件専用にしない。汎用語）
# (正規名, ページ上の別名…)
_DISH_SEED_ALIASES: Tuple[Tuple[str, ...], ...] = (
    ("茶漬け", "お茶漬け"),
    ("冷奴", "冷や奴", "冷ややっこ"),
    ("うどん", "おうどん"),
    ("そば", "おそば"),
    ("おにぎり",),
    ("サラダ",),
    ("チャーハン", "炒飯"),
    ("味噌汁", "お味噌汁"),
    ("卵かけご飯", "卵かけ", "TKG"),
    ("パスタ",),
    ("丼",),
    ("おやつ",),
)

_DISH_SEED_WORDS = tuple(a[0] for a in _DISH_SEED_ALIASES)

# 用事↔料理のゆるい適合（一件の献立表ではない。汎用ヒント）
_ERRAND_DISH_GROUPS: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
    (r"お茶請け|お茶うけ|おやつ", ("おやつ", "茶漬け", "お茶漬け")),
    (r"晩酌|つまみ|酒|ビール", ("冷奴",)),
    (r"一品|平日|時短|食卓", ("卵かけご飯", "うどん", "味噌汁", "おにぎり", "丼", "パスタ")),
    (r"家族|分けて", ("卵かけご飯", "うどん", "味噌汁", "パスタ", "丼")),
    (r"アクセント|丼|おかず", ("丼", "うどん", "サラダ", "パスタ", "卵かけご飯")),
    (r"手土産|差し入れ", ()),  # 料理なし優先（袋主役）
)


def normalize_dish_name_(raw: str) -> str:
    t = str(raw or "").strip()
    if not t:
        return ""
    for aliases in _DISH_SEED_ALIASES:
        canon = aliases[0]
        for a in aliases:
            if a == t or a in t:
                return canon
    return t


def dishes_from_text_blob_(
    blob: str,
    *,
    evidence: str,
    source_tier: str,
    source_url: str = "",
    errand_fit: str = "",
) -> List[Dict[str, str]]:
    """本文に出る食べ方種だけ（APIなし・一件専用語を焼かない）。"""
    out: List[Dict[str, str]] = []
    seen = set()
    text = str(blob or "")
    for aliases in _DISH_SEED_ALIASES:
        canon = aliases[0]
        if canon in seen or _DISH_BAN_RE.search(canon):
            continue
        if any(a and a in text for a in aliases):
            seen.add(canon)
            out.append(
                {
                    "dishName": canon,
                    "evidence": evidence,
                    "errandFit": errand_fit,
                    "sourceUrl": source_url,
                    "sourceTier": source_tier,
                }
            )
    return out


def dishes_from_materials_(materials: Dict[str, Any]) -> List[Dict[str, str]]:
    """材料・説明文にある食べ方だけ（APIなし）。"""
    return dishes_from_text_blob_(
        mat_blob(materials),
        evidence="材料・商品文に食べ方の記述あり",
        source_tier="materials",
    )


def fetch_page_(url: str, *, timeout: float = 8.0) -> Tuple[str, str]:
    """HTML本文と最終URL。失敗時は ('', '')。"""
    if not url or is_weak_recipe_source_url_(url):
        return "", ""
    try:
        import urllib.request

        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; B5Virt/1.0)"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            final = str(resp.geturl() or url)
            raw = resp.read(400_000)
        html = raw.decode("utf-8", "replace")
        html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
        html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
        text = re.sub(r"<[^>]+>", " ", html)
        return final, re.sub(r"\s+", " ", text)
    except Exception as e:
        LOG.warning("fetch_page_text_ fail url=%s err=%s", url[:80], e)
        return "", ""


def fetch_page_text_(url: str, *, timeout: float = 8.0) -> str:
    """HTMLからざっくり本文テキスト。失敗時空。"""
    _final, text = fetch_page_(url, timeout=timeout)
    return text


def urls_in_materials_(materials: Dict[str, Any]) -> List[str]:
    """材料文中の https（モール除外）。"""
    blob = mat_blob(materials)
    found = re.findall(r"https?://[^\s<>\"')\]]+", blob, flags=re.I)
    out: List[str] = []
    for u in found:
        u = u.rstrip(".,;）)")
        if is_weak_recipe_source_url_(u):
            continue
        if u not in out:
            out.append(u)
    return out


def url_path_fallbacks_(url: str) -> List[str]:
    """404対策の汎用パス揺れ（product/products・末尾スラッシュ）。一件専用スラッグは作らない。"""
    u = str(url or "").strip()
    if not u:
        return []
    cands = [u]
    if "/product/" in u and "/products/" not in u:
        cands.append(u.replace("/product/", "/products/", 1))
    if "/products/" in u:
        cands.append(u.replace("/products/", "/product/", 1))
    if u.endswith("/"):
        cands.append(u.rstrip("/"))
    else:
        cands.append(u + "/")
    out: List[str] = []
    for c in cands:
        if c and c not in out:
            out.append(c)
    return out


def first_live_product_url_(
    candidates: List[str],
    tokens: List[str],
    *,
    materials: Optional[Dict[str, Any]] = None,
) -> str:
    """候補URLを順に開き、商品手がかりがある本文のものを返す。"""
    tried = set()
    ranked: List[Tuple[int, str, str]] = []  # score, url, text
    for raw in candidates:
        if not raw or is_weak_recipe_source_url_(raw):
            continue
        for u in url_path_fallbacks_(raw):
            if u in tried:
                continue
            tried.add(u)
            final, text = fetch_page_(u)
            if not text:
                continue
            use = final or u
            if is_weak_recipe_source_url_(use):
                continue
            strong = [t for t in tokens if t in ("食べるおだし",) or "おだし" in t]
            ok = any(t in text for t in strong) if strong else any(t in text for t in tokens)
            if not ok:
                continue
            score = 0
            if "/product" in use.lower():
                score += 3
            if is_media_article_url_(use):
                score -= 5
            if "grounding-api-redirect" in use or "vertexaisearch.cloud.google.com" in use:
                score -= 8
            ranked.append((score, use, text))
    if not ranked:
        return ""
    ranked.sort(key=lambda x: (-x[0], x[1]))
    best_score, best_url, best_text = ranked[0]
    # メディアしか無いときは採用するが、ログで分かるようにする
    if best_score < 0:
        LOG.info("official_url_weak_fallback url=%s score=%s", best_url[:100], best_score)
    if materials is not None:
        materials["_officialPageText"] = best_text
    return best_url


def _official_cache_key_(materials: Dict[str, Any]) -> str:
    jan = re.sub(r"\D", "", str(materials.get("jan") or materials.get("parentSku") or ""))
    if len(jan) >= 8:
        return "jan:" + jan[-13:]
    name = re.sub(r"\s+", "", str(materials.get("name") or ""))[:40]
    return "name:" + name if name else ""


def load_official_url_cache_() -> Dict[str, Any]:
    if not OFFICIAL_URL_CACHE_PATH.is_file():
        return {"version": 1, "byKey": {}}
    try:
        data = json.loads(OFFICIAL_URL_CACHE_PATH.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {"version": 1, "byKey": {}}
        data.setdefault("byKey", {})
        return data
    except Exception as e:
        LOG.warning("load_official_url_cache_: %s", e)
        return {"version": 1, "byKey": {}}


def save_official_url_cache_(data: Dict[str, Any]) -> None:
    try:
        OFFICIAL_URL_CACHE_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        LOG.warning("save_official_url_cache_: %s", e)


def cache_get_official_url_(
    materials: Dict[str, Any],
    tokens: List[str],
    *,
    force_refresh: bool = False,
) -> str:
    """キャッシュ命中かつ生きていれば返す。メディア・404・TTL切れは破棄。"""
    if force_refresh:
        return ""
    key = _official_cache_key_(materials)
    if not key:
        return ""
    pack = load_official_url_cache_()
    row = (pack.get("byKey") or {}).get(key) or {}
    url = str(row.get("url") or "").strip()
    saved = float(row.get("savedAt") or 0)
    if not url:
        return ""
    if is_weak_recipe_source_url_(url) or is_media_article_url_(url):
        LOG.info("official_cache_drop weak/media key=%s", key)
        pack["byKey"].pop(key, None)
        save_official_url_cache_(pack)
        return ""
    if saved and (time.time() - saved) > OFFICIAL_URL_TTL_SEC:
        LOG.info("official_cache_ttl_expired key=%s", key)
        # TTL切れは破棄せず再検証だけ（生きていれば延長）
    hit = first_live_product_url_([url], tokens, materials=materials)
    if not hit:
        LOG.info("official_cache_drop dead key=%s url=%s", key, url[:80])
        pack["byKey"].pop(key, None)
        save_official_url_cache_(pack)
        return ""
    # 延長保存
    pack["byKey"][key] = {
        "url": hit,
        "savedAt": time.time(),
        "source": str(row.get("source") or "cache"),
        "name": str(materials.get("name") or "")[:80],
    }
    save_official_url_cache_(pack)
    LOG.info("official_cache_hit key=%s url=%s", key, hit[:100])
    return hit


def cache_put_official_url_(
    materials: Dict[str, Any],
    url: str,
    *,
    source: str,
) -> None:
    """商品名ゲート通過済みの公式URLだけ保存。メディアは保存しない。"""
    key = _official_cache_key_(materials)
    u = str(url or "").strip()
    if not key or not u:
        return
    if is_weak_recipe_source_url_(u) or is_media_article_url_(u):
        LOG.info("official_cache_skip_save media/weak url=%s", u[:80])
        return
    if "grounding-api-redirect" in u or "vertexaisearch" in u:
        return
    pack = load_official_url_cache_()
    pack.setdefault("byKey", {})[key] = {
        "url": u,
        "savedAt": time.time(),
        "source": source,
        "name": str(materials.get("name") or "")[:80],
    }
    save_official_url_cache_(pack)
    LOG.info("official_cache_save key=%s source=%s url=%s", key, source, u[:100])


def resolve_official_product_page_(
    materials: Dict[str, Any],
    *,
    client: Any = None,
    force_refresh: bool = False,
) -> str:
    """公式っぽい商品ページを汎用解決（一件のURLをコードに焼かない）。キャッシュ優先。"""
    cached = str(materials.get("_officialUrl") or "").strip()
    if cached and not force_refresh:
        return cached
    tokens = product_match_tokens_(materials)
    # 0) ディスクキャッシュ（TTL・生存確認・メディア除外）
    hit = cache_get_official_url_(materials, tokens, force_refresh=force_refresh)
    if hit:
        materials["_officialUrl"] = hit
        materials["_officialUrlSource"] = "cache"
        return hit
    # 1) 材料に載っている自社URL
    hit = first_live_product_url_(urls_in_materials_(materials), tokens, materials=materials)
    if hit:
        materials["_officialUrl"] = hit
        materials["_officialUrlSource"] = "materials"
        cache_put_official_url_(materials, hit, source="materials")
        return hit
    # 2) Gemini検索（最大2回。キャッシュがあれば次回は検索しない）
    if client is not None:
        name = str(materials.get("name") or "").strip()
        maker = str(materials.get("maker") or "").strip()
        base_prompt = (
            "ウェブ検索して、次の商品のメーカー公式（またはブランド公式）の商品紹介ページの"
            "https URLを返す。モール店（Yahoo/楽天/Amazon）は禁止。"
            "必ず実在しブラウザで開けるURLにすること。\n"
            f"メーカー:{maker or '（材料から不明）'}\n"
            f"商品名:{name}\n"
            'JSONのみ: {"url":"https://…"}\n'
        )
        try:
            from google.genai import types as genai_types

            tool = genai_types.Tool(google_search=genai_types.GoogleSearch())
            failed_note = ""
            seen_domains: List[str] = []
            max_attempts = 2
            for attempt in range(max_attempts):
                if attempt == 1 and seen_domains:
                    prompt = (
                        f"ウェブ検索: site:{seen_domains[0]} 「{name}」の商品紹介ページ。\n"
                        "モール禁止。実在するhttpsだけ。雑誌記事よりメーカー公式を優先。\n"
                        'JSONのみ: {"url":"https://…"}\n'
                    )
                else:
                    prompt = base_prompt + failed_note
                resp = None
                text = ""
                for mid in _GEMINI_MODELS:
                    try:
                        resp = client.models.generate_content(
                            model=mid,
                            contents=prompt,
                            config=genai_types.GenerateContentConfig(
                                tools=[tool], temperature=0.1
                            ),
                        )
                        text = str(getattr(resp, "text", None) or "").strip()
                        if text or resp is not None:
                            break
                    except Exception as e:
                        LOG.warning("resolve_official_product_page_ model=%s: %s", mid, e)
                cands: List[str] = []
                m = re.search(r"\{[\s\S]*\}", text or "")
                if m:
                    try:
                        data = json.loads(m.group(0))
                        cands.append(str(data.get("url") or "").strip())
                    except Exception:
                        pass
                for u in re.findall(r"https?://[^\s<>\"')\]]+", text or ""):
                    cands.append(u.rstrip(".,;）)"))
                if resp is not None:
                    cands.extend(extract_grounding_urls_(resp))
                for c in cands:
                    dm = re.match(r"https?://([^/]+)/", str(c or ""), re.I)
                    if not dm:
                        continue
                    host = dm.group(1).lower()
                    if any(x in host for x in ("google.", "gstatic.", "yahoo.", "rakuten.", "amazon.")):
                        continue
                    if host not in seen_domains:
                        seen_domains.append(host)
                hit = first_live_product_url_(cands, tokens, materials=materials)
                if hit:
                    materials["_officialUrl"] = hit
                    materials["_officialUrlSource"] = "search"
                    cache_put_official_url_(materials, hit, source="search")
                    LOG.info("official_url_resolved url=%s attempt=%s", hit[:120], attempt + 1)
                    return hit
                failed_note = (
                    "\n前回の候補は404または商品名なしだった。"
                    "別の公式商品ページURLを探し直せ。\n"
                )
                LOG.info(
                    "official_url_miss attempt=%s cands=%s",
                    attempt + 1,
                    [c[:80] for c in cands[:6]],
                )
        except Exception as e:
            LOG.warning("resolve_official_product_page_: %s", e)
    materials["_officialUrl"] = ""
    return ""


def dishes_from_official_page_(
    official_url: str,
    materials: Dict[str, Any],
) -> List[Dict[str, str]]:
    """公式ページ本文に出る食べ方（ページキャッシュ付き）。"""
    if not official_url:
        return []
    cached = materials.get("_officialPageText")
    if cached is None:
        cached = fetch_page_text_(official_url)
        materials["_officialPageText"] = cached
    text = str(cached or "")
    if not text:
        return []
    return dishes_from_text_blob_(
        text,
        evidence="公式商品ページ本文に食べ方の記述あり",
        source_tier="official_page",
        source_url=official_url,
    )


def official_page_mentions_dish_(materials: Dict[str, Any], dish: str) -> bool:
    text = str(materials.get("_officialPageText") or "")
    if not text or not dish:
        return False
    d = normalize_dish_name_(dish)
    for aliases in _DISH_SEED_ALIASES:
        if aliases[0] == d:
            return any(a and a in text for a in aliases)
    return d in text


def dish_allowed_(materials: Dict[str, Any], dish: str, fits: Dict[str, bool]) -> bool:
    """商品名に料理名があるときは落とさない。ほかは、その案の判定だけ。"""
    dn = normalize_dish_name_(dish)
    name = str((materials or {}).get("name") or "")
    if dn and name and dn in name:
        return True
    return bool((fits or {}).get(dn))


def judge_dish_fits_(
    materials: Dict[str, Any],
    proposal: Dict[str, Any],
    pool: List[Dict[str, str]],
    used: List[str],
    client: Any,
) -> Dict[str, Any]:
    """案につき1回。合う候補が無ければ、同じ応答で作り直す。"""
    fits: Dict[str, bool] = {}
    rest: Dict[str, bool] = {}
    if client is None or not pool or dish_optional_for_proposal_(proposal):
        if dish_optional_for_proposal_(proposal):
            LOG.info("judge_dish_fits_ skip=no_dish")
        return {"fits": fits, "rest": rest, "derived": None}
    lines = "\n".join(f"{i}. {r.get('dishName') or ''}" for i, r in enumerate(pool[:12], 1))
    prompt = (
        "添える料理を1つ決める。次のJSONだけ返す。説明は付けない。\n"
        '{"place":"fit","dishName":"","parent":"","mainIngredient":"","methodChange":""}\n'
        "placeは fit / remake / empty だけ。\n"
        "fit＝その人の得たい結果が、その料理の一枚で見える。dishNameは候補の名前そのもの。1つ。\n"
        "隣に置けるだけでは fit にしない。候補の番号、一覧の先頭、前の案の次の名前では決めない。\n"
        "remake＝候補のどれも一枚で見せられないときだけ。parentは候補の名前。dishNameは新しい名前。"
        "mainIngredientは親と同じ主材料。methodChangeは変えた作り方か和洋中。どちらかが空なら place を empty にする。用途の言い換えだけは作らない。\n"
        "empty＝料理なし。\n"
        "使用済みの名前は dishName にしない。商品名にその料理名が含まれるときは、その料理を fit にする。\n"
        f"商品名:{materials.get('name') or ''}\n"
        f"訴求:{proposal.get('appealPolicy') or ''}\n"
        f"用事:{errand_of(proposal) or ''}\n"
        f"キャッチ:{proposal.get('catch') or ''}\n"
        f"セールス:{proposal.get('salesTalk') or ''}\n"
        f"使用済み:{'、'.join(used) if used else 'なし'}\n"
        f"候補:\n{lines}\n"
    )

    def pool_row(name: str) -> Optional[Dict[str, str]]:
        dn = normalize_dish_name_(name)
        for r in pool:
            if normalize_dish_name_(str(r.get("dishName") or "")) == dn:
                return r
        return None

    place = "empty"
    dish_name = ""
    parent_name = ""
    ingredient = ""
    method = ""
    derived = None
    try:
        text = gemini_text(prompt, client)
        m = re.search(r"\{[\s\S]*\}", text or "")
        if m:
            got = json.loads(m.group(0))
            if isinstance(got, dict):
                place = str(got.get("place") or "").lower()
                dish_name = str(got.get("dishName") or "").strip()
                parent_name = normalize_dish_name_(str(got.get("parent") or ""))
                ingredient = str(got.get("mainIngredient") or "").strip()
                method = str(got.get("methodChange") or "").strip()
    except Exception as e:
        LOG.warning("judge_dish_fits_: %s", e)
        place = "empty"
    named = ""
    for r in pool[:12]:
        pn = normalize_dish_name_(str(r.get("dishName") or ""))
        if pn and pn in str(materials.get("name") or "") and pn not in used:
            named = pn
            break
    if named:
        place = "fit"
        dish_name = named
        parent_name = ""
    canon = normalize_dish_name_(dish_name)
    exact = bool(canon and dish_name == canon and pool_row(dish_name))
    parent = pool_row(parent_name)
    if place == "fit" and exact and canon not in used:
        fits[canon] = True
        dish_name = str(pool_row(dish_name).get("dishName") or canon)
    elif (
        place == "remake"
        and parent
        and dish_name
        and dish_name != str(parent.get("dishName") or "")
        and not exact
        and dish_name not in used
        and ingredient
        and method
        and remake_align_ok_(str(parent.get("dishName") or ""), ingredient, method, dish_name, client)
    ):
        derived = {
            "dishName": dish_name,
            "sourceUrl": str(parent.get("sourceUrl") or ""),
            "sourceTier": str(parent.get("sourceTier") or "pool"),
            "evidence": str(parent.get("evidence") or ""),
            "parent": str(parent.get("dishName") or ""),
            "mainIngredient": ingredient,
            "methodChange": method,
        }
        place = "remake"
    elif place == "remake" and exact and canon not in used:
        fits[canon] = True
        place = "fit"
        dish_name = str(pool_row(dish_name).get("dishName") or canon)
        derived = None
    else:
        place = "empty"
        dish_name = ""
        derived = None
    LOG.info(
        "judge_dish_fits_ place=%s dish=%s%s",
        place,
        dish_name or "-",
        f" parent={parent_name}" if place == "remake" else "",
    )
    return {"fits": fits, "rest": rest, "derived": derived}


def _locked_from_judged(judged: Dict[str, Any], pool: List[Dict[str, str]]) -> Dict[str, str]:
    derived = (judged or {}).get("derived") or None
    if isinstance(derived, dict) and derived.get("dishName"):
        return derived
    fits = (judged or {}).get("fits") or {}
    name = next((k for k, v in fits.items() if v), "")
    if not name:
        return {}
    row = next((r for r in pool if normalize_dish_name_(str(r.get("dishName") or "")) == normalize_dish_name_(name)), None)
    if not row:
        return {"dishName": name, "sourceUrl": "", "sourceTier": "", "evidence": "", "parent": ""}
    return {
        "dishName": str(row.get("dishName") or name),
        "sourceUrl": str(row.get("sourceUrl") or ""),
        "sourceTier": str(row.get("sourceTier") or ""),
        "evidence": str(row.get("evidence") or ""),
        "parent": "",
    }


def remake_dish_once_(
    materials: Dict[str, Any],
    proposal: Dict[str, Any],
    pool: List[Dict[str, str]],
    used: List[str],
    client: Any,
) -> Dict[str, str]:
    """食べ方案で料理が決まらないときだけ1回。主材料と作り方が両方なければ空。"""
    if client is None or not pool or dish_optional_for_proposal_(proposal):
        return {}
    lines = "\n".join(f"{i}. {r.get('dishName') or ''}" for i, r in enumerate(pool[:12], 1))
    prompt = (
        "候補では得たい結果の一枚にならない。親を1つ選び、同じ主材料のまま作り方か和洋中だけ変えた料理を1つ。JSONだけ返す。\n"
        '{"parent":"","dishName":"","mainIngredient":"","methodChange":""}\n'
        "parentは候補の名前。mainIngredientは親と同じ主材料。methodChangeは変えた作り方か和洋中。どちらか書けなければ全部空。用途の言い換えだけは作らない。\n"
        f"訴求:{proposal.get('appealPolicy') or ''}\n"
        f"用事:{errand_of(proposal) or ''}\n"
        f"使用済み:{'、'.join(used) if used else 'なし'}\n"
        f"候補:\n{lines}\n"
    )
    try:
        text = gemini_text(prompt, client)
        m = re.search(r"\{[\s\S]*\}", text or "")
        if not m:
            return {}
        got = json.loads(m.group(0))
        parent_name = normalize_dish_name_(str(got.get("parent") or ""))
        dish_name = str(got.get("dishName") or "").strip()
        ingredient = str(got.get("mainIngredient") or "").strip()
        method = str(got.get("methodChange") or "").strip()
        parent = next((r for r in pool if normalize_dish_name_(str(r.get("dishName") or "")) == parent_name), None)
        canon = normalize_dish_name_(dish_name)
        exact = bool(canon and dish_name == canon)
        if (
            not parent
            or not dish_name
            or not ingredient
            or not method
            or exact
            or dish_name in used
            or not remake_align_ok_(str(parent.get("dishName") or ""), ingredient, method, dish_name, client)
        ):
            LOG.info("remake_dish_once_ empty parent=%s", parent_name or "-")
            return {}
        LOG.info("remake_dish_once_ parent=%s dish=%s ingredient=%s method=%s", parent.get("dishName"), dish_name, ingredient, method)
        return {
            "dishName": dish_name,
            "sourceUrl": str(parent.get("sourceUrl") or ""),
            "sourceTier": str(parent.get("sourceTier") or "pool"),
            "evidence": str(parent.get("evidence") or ""),
            "parent": str(parent.get("dishName") or ""),
            "mainIngredient": ingredient,
            "methodChange": method,
        }
    except Exception as e:
        LOG.warning("remake_dish_once_: %s", e)
        return {}


def derive_dish_from_hit_(
    materials: Dict[str, Any],
    proposal: Dict[str, Any],
    pool: List[Dict[str, str]],
    used: List[str],
    *,
    client: Any,
) -> Dict[str, str]:
    """fitが無いときだけ。親は候補。同じ主材料で作り方か和洋中だけ変える。"""
    if client is None or not pool or gift_errand_(errand_of(proposal)) or stock_errand_(errand_of(proposal)):
        return {}
    lines = "\n".join(f"{i}. {r.get('dishName') or ''}" for i, r in enumerate(pool[:12], 1))
    prompt = (
        "ヒットした料理を親にして、同じ主材料のまま、作り方か和洋中だけを変えた料理名を1つ。JSONだけ返す。\n"
        '{"parent":"","dishName":""}\n'
        "親は候補の名前そのもの。材料がつながらない料理は作らず、parentとdishNameを空にする。\n"
        f"訴求:{proposal.get('appealPolicy') or ''}\n"
        f"用事:{errand_of(proposal) or ''}\n"
        f"候補:\n{lines}\n"
    )
    try:
        text = gemini_text(prompt, client)
        m = re.search(r"\{[\s\S]*\}", text or "")
        if not m:
            return {}
        got = json.loads(m.group(0))
        parent_name = normalize_dish_name_(str(got.get("parent") or ""))
        dish_name = str(got.get("dishName") or "").strip()
        parent = next(
            (r for r in pool if normalize_dish_name_(str(r.get("dishName") or "")) == parent_name),
            None,
        )
        if (
            not parent
            or not dish_name
            or normalize_dish_name_(dish_name) == parent_name
            or dish_name in used
        ):
            return {}
        return {
            "dishName": dish_name,
            "sourceUrl": str(parent.get("sourceUrl") or ""),
            "evidence": str(parent.get("evidence") or ""),
            "parent": str(parent.get("dishName") or ""),
        }
    except Exception as e:
        LOG.warning("derive_dish_from_hit_: %s", e)
        return {}


def errand_dish_affinity_(errand: str, dish: str, appeal: str = "") -> int:
    """料理名の一覧では決めない。合否はその案の判定。"""
    return 0


def _gemini_search_json_array_(
    prompt: str,
    *,
    client: Any,
) -> Tuple[List[Dict[str, Any]], float]:
    """Google検索付きで JSON 配列を1回取る。"""
    t0 = time.time()
    if client is None:
        return [], 0.0
    try:
        from google.genai import types as genai_types

        tool = genai_types.Tool(google_search=genai_types.GoogleSearch())
        text = ""
        for mid in _GEMINI_MODELS:
            try:
                resp = client.models.generate_content(
                    model=mid,
                    contents=prompt,
                    config=genai_types.GenerateContentConfig(tools=[tool], temperature=0.2),
                )
                text = str(getattr(resp, "text", None) or "").strip()
                if text:
                    break
            except Exception as e:
                LOG.warning("_gemini_search_json_array_ model=%s: %s", mid, e)
                continue
        m = re.search(r"\[[\s\S]*\]", text or "")
        if not m:
            return [], time.time() - t0
        data = json.loads(m.group(0))
        rows = [r for r in data if isinstance(r, dict)] if isinstance(data, list) else []
        return rows, time.time() - t0
    except Exception as e:
        LOG.warning("_gemini_search_json_array_: %s", e)
        return [], time.time() - t0


def gemini_recipe_pool_(
    materials: Dict[str, Any],
    *,
    client: Any,
    need: int = 10,
    use_competitor_hints: bool = False,
    force_refresh_official: bool = False,
) -> List[Dict[str, str]]:
    """公式ページ本文→材料（ページにもあるもの）→Google。モール除外。競合既定オフ。"""
    t0 = time.time()
    name = str(materials.get("name") or "").strip()
    maker = str(materials.get("maker") or "").strip()
    marks = list(materials.get("mixMarks") or [])
    both = "・".join(marks[:2]) if len(marks) >= 2 else ""
    tokens = product_match_tokens_(materials)
    timing: Dict[str, float] = {
        "official": 0.0,
        "materials": 0.0,
        "google": 0.0,
        "competitor": 0.0,
    }

    to0 = time.time()
    official = resolve_official_product_page_(
        materials, client=client, force_refresh=force_refresh_official
    )
    timing["official"] = time.time() - to0
    if official and materials.get("_officialPageText") is None:
        materials["_officialPageText"] = fetch_page_text_(official)

    out: List[Dict[str, str]] = []
    seen = set()
    dropped = 0

    def accept(row: Dict[str, str], *, require_url: bool) -> None:
        nonlocal dropped
        dn = normalize_dish_name_(str(row.get("dishName") or "").strip())
        url = str(row.get("sourceUrl") or "").strip()
        if not dn or _DISH_BAN_RE.search(dn) or dn in seen:
            return
        if require_url:
            if is_weak_recipe_source_url_(url):
                LOG.info("recipe_url_reject_mall dish=%s url=%s", dn, url[:100])
                dropped += 1
                return
            if not re.match(r"^https?://", url, re.I):
                dropped += 1
                return
            # 公式ページ由来は本文抽出済みなので商品名再フェッチを省略可
            if row.get("sourceTier") != "official_page":
                if not url_page_mentions_product_(url, tokens):
                    alt = first_live_product_url_([url], tokens, materials=materials)
                    if alt:
                        url = alt
                    else:
                        LOG.info("recipe_url_reject dish=%s url=%s", dn, url[:100])
                        dropped += 1
                        return
        elif url and (is_weak_recipe_source_url_(url) or not url_page_mentions_product_(url, tokens)):
            url = ""
        seen.add(dn)
        out.append(
            {
                "dishName": dn,
                "evidence": str(row.get("evidence") or "").strip(),
                "errandFit": str(row.get("errandFit") or "").strip(),
                "sourceUrl": url,
                "sourceTier": str(row.get("sourceTier") or ""),
            }
        )

    # --- Tier0: マスタの食べ方（URLなし可）---
    tm0 = time.time()
    for row in dishes_from_materials_(materials):
        row["evidence"] = "マスタの文に食べ方の記述あり"
        accept(row, require_url=False)
    timing["materials"] = time.time() - tm0

    # --- Tier1: 公式ページ。同じ料理にはURLを足す ---
    for row in dishes_from_official_page_(official, materials):
        dn = row["dishName"]
        hit = next((r for r in out if r["dishName"] == dn), None)
        if hit is not None:
            if official and not hit.get("sourceUrl"):
                hit["sourceUrl"] = official
                hit["evidence"] = "マスタと公式ページの両方に食べ方の記述あり"
            continue
        accept(row, require_url=True)

    # --- Tier2: 公式／商品名での Google 1回 ---
    if len(out) < need and client is not None:
        prompt = (
            f"次をウェブ検索する（検索は1回分の回答で足りるように）。\n"
            f"1) メーカー公式・ブランド公式の「{name}」レシピ／食べ方"
            f"{('（メーカー:' + maker + '）') if maker else ''}\n"
            f"2) 商品名「{name}」が本文にあるレシピ・記事\n"
            f"Yahoo店・楽天店・Amazonなどモール販売ページはsourceUrlに使わない。\n"
            f"別メーカーの類似品は禁止。空想禁止。ミックス:{both or 'なし'}\n"
            f"JSON配列のみ（最大{need}件）:\n"
            '[{"dishName":"完成料理名","evidence":"要約","errandFit":"用事",'
            '"sourceUrl":"https://…"}]\n'
        )
        rows, sec = _gemini_search_json_array_(prompt, client=client)
        timing["google"] = sec
        for r in rows:
            accept(
                {
                    "dishName": str(r.get("dishName") or ""),
                    "evidence": str(r.get("evidence") or ""),
                    "errandFit": str(r.get("errandFit") or ""),
                    "sourceUrl": str(r.get("sourceUrl") or r.get("url") or ""),
                    "sourceTier": "google_official",
                },
                require_url=True,
            )
            if len(out) >= need:
                break

    # --- Tier3: 競合ヒント（テスト用）---
    if use_competitor_hints and len(out) < need and client is not None:
        prompt = (
            f"競合の同カテゴリ食品の売れ筋ページや口コミを検索し、"
            f"よく載っている完成料理名だけを集める（最大6件）。\n"
            f"当社商品「{name}」のURLは出さない。sourceUrlは空文字。\n"
            f"JSON配列のみ:\n"
            '[{"dishName":"完成料理名","evidence":"競合でよく見る食べ方","errandFit":"用事",'
            '"sourceUrl":""}]\n'
        )
        rows, sec = _gemini_search_json_array_(prompt, client=client)
        timing["competitor"] = sec
        for r in rows:
            dn = normalize_dish_name_(str(r.get("dishName") or "").strip())
            if not dn or dn in seen:
                continue
            accept(
                {
                    "dishName": dn,
                    "evidence": "競合ヒント（別商品）。当社URLは公式ページのみ",
                    "errandFit": str(r.get("errandFit") or ""),
                    "sourceUrl": official if official else "",
                    "sourceTier": "competitor_hint",
                },
                require_url=bool(official),
            )
            if len(out) >= need:
                break

    with_url = [r for r in out if re.match(r"^https?://", r.get("sourceUrl") or "", re.I)]
    LOG.info(
        "recipe_pool size=%s with_url=%s dropped=%s official=%s "
        "sec_total=%.1f sec_off=%.1f sec_mat=%.1f sec_google=%.1f sec_comp=%.1f "
        "tiers=%s names=%s",
        len(out),
        len(with_url),
        dropped,
        (official or "")[:80],
        time.time() - t0,
        timing["official"],
        timing["materials"],
        timing["google"],
        timing["competitor"],
        [r.get("sourceTier") for r in with_url],
        [r["dishName"] for r in with_url],
    )
    materials["_recipeTiming"] = timing
    return out[:need]


def extract_grounding_urls_(resp: Any) -> List[str]:
    """Gemini 検索グラウンディングから https URL を拾う。"""
    urls: List[str] = []
    try:
        cands = getattr(resp, "candidates", None) or []
        for c in cands:
            gm = getattr(c, "grounding_metadata", None) or getattr(c, "groundingMetadata", None)
            if not gm:
                continue
            chunks = getattr(gm, "grounding_chunks", None) or getattr(gm, "groundingChunks", None) or []
            for ch in chunks:
                web = getattr(ch, "web", None)
                if not web:
                    continue
                u = str(getattr(web, "uri", None) or getattr(web, "url", None) or "").strip()
                if re.match(r"^https?://", u, re.I) and u not in urls:
                    urls.append(u)
    except Exception as e:
        LOG.warning("extract_grounding_urls_: %s", e)
    return urls


def resolve_dish_source_urls_batch_(
    product_name: str,
    dishes: List[str],
    *,
    client: Any,
) -> Dict[str, str]:
    """複数料理のソースURLを1回の検索で取る。"""
    if client is None or not dishes:
        return {}
    listing = "\n".join(f"- {d}" for d in dishes)
    prompt = (
        f"商品「{product_name}」について、次の各料理の使い方が書かれているウェブページを検索する。\n"
        f"{listing}\n"
        "各料理ごとに https URL を1つ。商品販売ページだけで食べ方記述が無いものは不可。\n"
        "空想URL禁止。JSONのみ:\n"
        '{"items":[{"dishName":"…","sourceUrl":"https://…"}]}\n'
    )
    try:
        from google.genai import types as genai_types

        tool = genai_types.Tool(google_search=genai_types.GoogleSearch())
        for mid in _GEMINI_MODELS:
            try:
                resp = client.models.generate_content(
                    model=mid,
                    contents=prompt,
                    config=genai_types.GenerateContentConfig(tools=[tool], temperature=0.2),
                )
                text = str(getattr(resp, "text", None) or "").strip()
                m = re.search(r"\{[\s\S]*\}", text)
                if not m:
                    continue
                data = json.loads(m.group(0))
                items = data.get("items") if isinstance(data, dict) else None
                out: Dict[str, str] = {}
                for row in items or []:
                    if not isinstance(row, dict):
                        continue
                    dn = str(row.get("dishName") or "").strip()
                    u = str(row.get("sourceUrl") or "").strip()
                    if dn and re.match(r"^https?://", u, re.I):
                        out[dn] = u
                if out:
                    return out
            except Exception as e:
                LOG.warning("resolve_dish_source_urls_batch_ fail model=%s: %s", mid, e)
                continue
    except Exception as e:
        LOG.warning("resolve_dish_source_urls_batch_: %s", e)
    return {}


def resolve_dish_source_url_(
    product_name: str,
    dish_name: str,
    *,
    client: Any,
    prefer: Optional[List[str]] = None,
) -> str:
    """その料理名＋商品の組み合わせで実在ページURLを1つ取る。"""
    if client is None or not dish_name:
        return ""
    prompt = (
        f"ウェブ検索して、商品「{product_name}」を「{dish_name}」に使う／載せる／トッピングする、"
        f"と書かれているレシピ・記事・口コミ・公式の食べ方ページのURLを1つだけ返す。\n"
        f"出力はJSONのみ: {{\"sourceUrl\":\"https://...\",\"title\":\"ページ題名\"}}\n"
        f"見つからなければ {{\"sourceUrl\":\"\"}}\n"
        f"商品の販売ページだけ（食べ方・レシピ記述が無い）は不可。"
        f"空想のURLは禁止。\n"
    )
    try:
        from google.genai import types as genai_types

        tool = genai_types.Tool(google_search=genai_types.GoogleSearch())
        for mid in _GEMINI_MODELS:
            try:
                resp = client.models.generate_content(
                    model=mid,
                    contents=prompt,
                    config=genai_types.GenerateContentConfig(tools=[tool], temperature=0.2),
                )
                text = str(getattr(resp, "text", None) or "").strip()
                urls = extract_grounding_urls_(resp)
                m = re.search(r"\{[\s\S]*\}", text)
                url = ""
                if m:
                    data = json.loads(m.group(0))
                    url = str(data.get("sourceUrl") or "").strip()
                if not re.match(r"^https?://", url, re.I):
                    # グラウンディング先頭を採用（prefer より今回検索分を優先）
                    if urls:
                        url = urls[0]
                    elif prefer:
                        url = prefer[0]
                if re.match(r"^https?://", url, re.I):
                    return url
            except Exception as e:
                LOG.warning("resolve_dish_source_url_ fail model=%s: %s", mid, e)
                continue
    except Exception as e:
        LOG.warning("resolve_dish_source_url_: %s", e)
    return ""


def _dish_already_used(name: str, used: List[str]) -> bool:
    key = normalize_dish_name_(name)
    if not key:
        return False
    return any(normalize_dish_name_(u) == key for u in used)


def _copy_mentions_dish(text: str, dish: str) -> bool:
    dn = normalize_dish_name_(dish)
    if not text or not dn:
        return False
    for aliases in _DISH_SEED_ALIASES:
        if aliases[0] != dn:
            continue
        return any(a and a in text for a in aliases)
    return dn in text


def pick_recipe_for_errand_(
    pool: List[Dict[str, str]],
    errand: str,
    used: List[str],
    copy_text: str = "",
    materials: Optional[Dict[str, Any]] = None,
    fits: Optional[Dict[str, bool]] = None,
) -> Optional[Dict[str, str]]:
    """用事に合う実在レシピを1つ。使った料理は選ばない。候補が無ければ空。"""
    if not pool:
        return None
    err = str(errand or "")
    if dish_optional_for_errand_(err):
        return None

    def score_row(row: Dict[str, str]) -> int:
        dn = str(row.get("dishName") or "")
        if fits is not None:
            return 1 if dish_allowed_(materials or {}, dn, fits) else -1
        return 0

    # 1) 未使用かつ適合あり
    unused: List[Tuple[int, Dict[str, str]]] = []
    for row in pool:
        src_row = str(row.get("sourceUrl") or "")
        master_dish = row.get("sourceTier") == "materials"
        if not master_dish and not re.match(r"^https?://", src_row, re.I):
            continue
        if re.match(r"^https?://", src_row, re.I) and is_weak_recipe_source_url_(src_row):
            continue
        sc = score_row(row)
        if sc <= 0:
            continue
        if _dish_already_used(str(row["dishName"]), used):
            continue
        unused.append((sc, row))
    unused.sort(key=lambda x: (-x[0], x[1]["dishName"]))
    if unused:
        return unused[0][1]
    return None


def pick_rest_dish_(
    pool: List[Dict[str, str]],
    used: List[str],
    rest: Dict[str, bool],
) -> Optional[Dict[str, str]]:
    """空の案へ回す。矛盾しないと判定された未使用候補だけ。"""
    if not pool or not rest:
        return None
    for row in pool:
        dn = normalize_dish_name_(str(row.get("dishName") or ""))
        if not dn or not rest.get(dn) or _dish_already_used(dn, used):
            continue
        src_row = str(row.get("sourceUrl") or "")
        master_dish = row.get("sourceTier") == "materials"
        if not master_dish and not re.match(r"^https?://", src_row, re.I):
            continue
        if re.match(r"^https?://", src_row, re.I) and is_weak_recipe_source_url_(src_row):
            continue
        return row
    return None


def piece_marks_(materials: Dict[str, Any]) -> List[str]:
    """個品の素材。商品名の括弧内。売り文句は外す。"""
    name = str((materials or {}).get("name") or "")
    m = re.search(r"[（(]([^）)]+)[）)]", name)
    raw = m.group(1) if m else ""
    drop = ["高タンパク", "低糖質", "しっとり", "ソフト削り", "ソフト", "ヘルシー", "食感", "低カロリー", "無添加", "使用", "国産"]
    out: List[str] = []
    seen = set()
    for bit in re.split(r"[・／/、,，\s　]+", raw):
        rest = str(bit or "").strip()
        for w in drop:
            rest = rest.replace(w, "")
        rest = re.sub(r"^[なの・]+|[なの・]+$", "", rest)
        if len(rest) < 2 or rest in seen:
            continue
        seen.add(rest)
        out.append(rest)
    LOG.info("piece_ marks=%s", "、".join(out) or "-")
    return out


def dish_candidates_for_request_(
    materials: Dict[str, Any],
    proposal: Dict[str, Any],
    pool: List[Dict[str, str]],
    client: Any,
) -> List[Dict[str, str]]:
    """マスタか公式ページにある食べ方だけ。新しい料理名は作らない。"""
    out: List[Dict[str, str]] = []
    for row in pool or []:
        name = str(row.get("dishName") or "").strip()
        if not name or _DISH_BAN_RE.search(name):
            continue
        out.append({"kind": "pool", "dishName": name, "sourceUrl": str(row.get("sourceUrl") or ""), "usedPieces": ""})
        LOG.info("dish_cand_ kind=pool dish=%s", name)
    return out


def pick_dish_for_situation_(
    proposal: Dict[str, Any],
    candidates: List[Dict[str, str]],
    used: List[str],
    client: Any,
) -> Dict[str, str]:
    """状況に合う既存の食べ方を1つ。候補に無い名前は残さない。"""
    situation = str(proposal.get("situation") or "").strip()
    if not situation:
        LOG.info("dish_pick_ empty reason=no_situation")
        return {"dishName": "", "sourceUrl": ""}
    by_name: Dict[str, Dict[str, str]] = {}
    for row in candidates or []:
        name = str(row.get("dishName") or "").strip()
        if not name or name in used or name in by_name:
            continue
        by_name[name] = row
    if not by_name:
        LOG.info("dish_pick_ empty reason=none")
        return {"dishName": "", "sourceUrl": ""}
    if client is None:
        LOG.info("dish_pick_ empty reason=no_api")
        return {"dishName": "", "sourceUrl": ""}
    got = _gemini_obj(
        "状況に合う食べ方を1つ選ぶ。JSONだけ。\n"
        '{"dishName":""}\n'
        "dishNameは候補の名前そのもの。合うものが無ければ空。新しい名前は作らない。\n"
        f"状況:{situation}\n"
        f"得たい結果:{errand_of(proposal) or ''}\n"
        f"候補:{'、'.join(by_name)}\n",
        client,
    ) or {}
    chosen = str(got.get("dishName") or "").strip()
    if chosen not in by_name:
        LOG.info("dish_pick_ empty reason=no_fit raw=%s", chosen or "-")
        return {"dishName": "", "sourceUrl": ""}
    LOG.info("dish_pick_ dish=%s", chosen)
    return {"dishName": chosen, "sourceUrl": str(by_name[chosen].get("sourceUrl") or "")}


def gemini_burns_scene_dish_(
    materials: Dict[str, Any],
    p: Dict[str, Any],
    *,
    client: Any,
    recipe_pool: Optional[List[Dict[str, str]]] = None,
    used_dishes: Optional[List[str]] = None,
    used_layouts: Optional[List[str]] = None,
) -> Optional[Dict[str, str]]:
    """場面／添える料理（catch/salesは呼び出し側で確定済み。ここでは場面用）。"""
    if client is None:
        return None
    marks = list(materials.get("mixMarks") or [])
    both = "・".join(marks[:2]) if len(marks) >= 2 else ""
    err = errand_of(p) or "すぐ使える"
    used_dishes = list(used_dishes or [])
    used_layouts = list(used_layouts or [])
    p["dishCandidates"] = dish_candidates_for_request_(materials, p, list(recipe_pool or []), client)
    picked = pick_dish_for_situation_(p, p["dishCandidates"], used_dishes, client)
    p["_lockedDish"] = picked
    gift = proposal_is_gift_(p)
    layout_candidates = [x for x in _SCENE_LAYOUTS if x not in used_layouts] or list(_SCENE_LAYOUTS)
    layout = layout_candidates[0]
    pool_lines = ""
    prefer = str(picked.get("dishName") or "")
    if prefer:
        dish_rule = f"- dishNameは「{prefer}」。別の料理名に変えない。新しい名前は作らない。\n"
    else:
        dish_rule = "- dishNameは空。新しい料理名は作らない。\n"
    situation = str(p.get("situation") or "").strip()
    scene_dish = (
        "状況の場が絵で特定できるところまで。商品の見本を前に置く。包材は見本の後ろ。"
        "袋を並べて主役にしない。組み合わせは袋と個品の並び順をそろえる。"
        "顔と手渡す手元は出さない。描く物はその案の状況の場だけ。選んだ食べ方と別の食べ物は足さない。"
        "状況の文は画像の文字にしない。保管を棚の一種類にしない。買い置きを理由に棚にしない。"
        + (f"場は{situation}。" if situation else "状況が空なら商品そのもの。")
        + ("包んだ箱を商品の横に置く。手渡す手元は使わない。" if gift else "")
        + (f"添える料理は{prefer}。" if prefer else "料理は描かない。")
    )
    prompt = (
        "食品ECのメイン画像用。次のJSONだけ返す（説明文なし）。\n"
        '{"catchBurn":"","salesBurn":"","sceneHint":"","dishName":"","dishEvidence":"","linkWhy":""}\n'
        "linkWhyは、状況のどの場を、絵のどの物に結び付けたか。\n"
        f"商品名:{materials.get('name') or ''}\n"
        f"用事:{err}\n"
        f"提案キャッチ:{p.get('catch') or ''}\n"
        f"提案セールス:{p.get('salesTalk') or ''}\n"
        f"ペルソナ短文:{p.get('personaShort') or ''}\n"
        f"ミックス魚種:{both or 'なし'}\n"
        f"推奨レイアウト（必ず使う）:{layout}\n"
        f"推奨料理名（優先）:{prefer or '空'}\n"
        f"状況:{situation or '空'}\n"
        f"{('主材料は「' + str((picked or {}).get('mainIngredient') or '') + '」。場面はこの主材料だけ。別の食材に変えない。') if (picked or {}).get('mainIngredient') else ''}\n"
        f"使用済み料理名（選ばない。残りの候補が無ければ料理名は空）:"
        f"{'／'.join(used_dishes) if used_dishes else 'なし'}\n"
        f"{pool_lines}"
        "用語の定義（重要）:\n"
        f"{dish_rule}"
        "- 小皿に盛った商品の身・厚削り＝【個品見本】であり料理名ではない。\n"
        "ルール:\n"
        f"- catchBurn: 全角{int(BURN_MIN)}〜{int(BURN_MAX)}字。【】1つ＋本文。"
        "語の途中で切らない。完結。数字・袋数・円なし。\n"
        f"- salesBurn: 全角{SALES_BURN_MIN}〜{SALES_BURN_MAX}字。【】は付けない。"
        "catchBurnと意味が被らない補足。語の途中で切らない。完結。\n"
        "- dishEvidence: 候補の evidence を短く写す（無い・料理なしのとき空で可）。\n"
        "- sceneHint: 写真指示 2文まで。売文禁止。"
        f"①推奨レイアウトを具体化 {scene_dish} "
        "③「小皿は個品見本（料理ではない）」④「包材は盛り付けの後ろ」。\n"
        "「袋を手前」「手前に料理だけ」は禁止。案ごとに背景・器・配置を変える。\n"
        "ミックスなら個品見本で両魚種が分かる。\n"
        "- 「袋を描くな」は書かない。\n"
    )
    try:
        text = gemini_text(prompt, client)
        m = re.search(r"\{[\s\S]*\}", text)
        if not m:
            return None
        data = json.loads(m.group(0))
        if prefer:
            dish = prefer
            src = str((picked or {}).get("sourceUrl") or "").strip()
        else:
            dish = ""
            src = ""
        out = {
            "catchBurn": strip_ban_copy(str(data.get("catchBurn") or "")),
            "salesBurn": re.sub(r"【[^】]*】", "", strip_ban_copy(str(data.get("salesBurn") or ""))).strip(),
            "sceneHint": str(data.get("sceneHint") or "").strip(),
            "sceneLinkWhy": str(data.get("linkWhy") or "").replace("\n", " ")[:180],
            "dishName": dish,
            "dishEvidence": (
                ""
                if not dish
                else str(data.get("dishEvidence") or (picked or {}).get("evidence") or "").strip()
            ),
            "dishSourceUrl": src if dish else "",
            "dishSource": str((picked or {}).get("sourceTier") or ""),
            "sceneLayout": layout,
        }
        if not out["catchBurn"] or not out["salesBurn"] or not out["sceneHint"]:
            return None
        if "袋を手前" in out["sceneHint"]:
            out["sceneHint"] = out["sceneHint"].replace("袋を手前に", "包材は盛り付けの後ろ").replace(
                "袋を手前", "包材は盛り付けの後ろ"
            )
        if "包材" not in out["sceneHint"]:
            out["sceneHint"] = out["sceneHint"].rstrip("。") + "。包材は後ろ。"
        if gift and "箱" not in out["sceneHint"]:
            out["sceneHint"] += "包んだ箱を商品の横に置く。"
        out["sceneHint"] = out["sceneHint"].replace("手渡す手元", "包んだ箱")
        out["sceneHint"] = _append_scene_draw_rules(out["sceneHint"])
        if dish and dish not in out["sceneHint"]:
            out["sceneHint"] = f"添える料理は{dish}。" + out["sceneHint"]
        ing = str((picked or {}).get("mainIngredient") or "")
        if dish and ing and not scene_uses_ingredient_(ing, out["sceneHint"], client):
            LOG.info("remake_align_ scene empty ingredient=%s", ing)
            dish = ""
            out["dishName"] = ""
            out["dishEvidence"] = ""
            out["dishSourceUrl"] = ""
            bare = dict(p)
            bare["dishName"] = ""
            bare["dish"] = ""
            out["sceneHint"] = scene_from_persona(bare, materials)
        # 字数オーバー・途切れはスライスで無理に残さない。1回だけ完結短縮を依頼
        cb = out["catchBurn"]
        sb = out["salesBurn"]
        if (
            zenkaku_len(cb) > BURN_MAX
            or zenkaku_len(sb) > SALES_BURN_MAX
            or looks_incomplete_(cb)
            or looks_incomplete_(sb)
        ):
            fix = gemini_text(
                "次を完結した短い日本語に直す。JSONのみ。"
                f'{{"catchBurn":"全角{int(BURN_MIN)}〜{int(BURN_MAX)}字","salesBurn":"全角{SALES_BURN_MIN}〜{SALES_BURN_MAX}字"}}\n'
                "語の途中で切らない。助詞・「厚削」「かつ」「だ」で終わらせない。\n"
                f"catchBurn元:{cb}\nsalesBurn元:{sb}\n",
                client,
            )
            m2 = re.search(r"\{[\s\S]*\}", fix or "")
            if m2:
                d2 = json.loads(m2.group(0))
                cb = strip_ban_copy(str(d2.get("catchBurn") or cb))
                sb = strip_ban_copy(str(d2.get("salesBurn") or sb))
        if (
            not cb
            or not sb
            or looks_incomplete_(cb)
            or looks_incomplete_(sb)
            or zenkaku_len(cb) > BURN_MAX
            or zenkaku_len(sb) > SALES_BURN_MAX
            or zenkaku_len(cb) < BURN_MIN
            or zenkaku_len(sb) < SALES_BURN_MIN
        ):
            return None
        out["catchBurn"] = cb
        out["salesBurn"] = sb
        return out
    except Exception as e:
        print(f"[warn] gemini_burns_scene_dish_: {e}", file=sys.stderr)
        return None


def is_true(v: Any) -> bool:
    return str(v or "").strip().upper() in ("TRUE", "1", "✓", "✔", "ON")


def list_checked_parents() -> List[Dict[str, Any]]:
    rows, _info = fetch_master_rows()
    hi, idx = _find_header_row(rows)
    ck_i = idx.get("出品CK")
    parent_i = idx.get("親SKU")
    child_i = idx.get("子SKU")
    jan_i = idx.get("JANコード") or idx.get("JAN")
    name_i = idx.get("商品名ベース")
    cat_i = next((idx[k] for k in idx if "カテゴリー" in k and "食品" in k), idx.get("カテゴリー"))
    combo_i = idx.get("セット構成1_親SKU")
    if ck_i is None or parent_i is None:
        raise SystemExit("マスタに出品CKまたは親SKUがありません")

    by_parent: Dict[str, Dict[str, Any]] = {}
    for r in rows[hi + 1 :]:
        if ck_i >= len(r) or not is_true(r[ck_i]):
            continue
        parent = str(r[parent_i] or "").strip() if parent_i < len(r) else ""
        child = str(r[child_i] or "").strip() if child_i is not None and child_i < len(r) else ""
        if not parent:
            continue
        slot = by_parent.setdefault(
            parent,
            {
                "parentSku": parent,
                "jan": "",
                "nameBase": "",
                "category": "",
                "comboSku1": "",
                "childSkus": [],
                "parentRowCk": False,
            },
        )
        if not child:
            slot["parentRowCk"] = True
            if jan_i is not None and jan_i < len(r):
                slot["jan"] = _jan_digits(r[jan_i]) or slot["jan"]
            if name_i is not None and name_i < len(r) and str(r[name_i]).strip():
                slot["nameBase"] = str(r[name_i]).strip()
            if cat_i is not None and cat_i < len(r):
                slot["category"] = str(r[cat_i] or "").strip()
            if combo_i is not None and combo_i < len(r):
                slot["comboSku1"] = str(r[combo_i] or "").strip()
        else:
            if child not in slot["childSkus"]:
                slot["childSkus"].append(child)
            if not slot["nameBase"] and name_i is not None and name_i < len(r) and str(r[name_i]).strip():
                slot["nameBase"] = str(r[name_i]).strip()
            if not slot["jan"]:
                # 親SKUから JAN を拾う（ishi-4906…-oya2）
                m = re.search(r"(\d{8,14})", parent)
                if m:
                    slot["jan"] = m.group(1)

    out = []
    for p in by_parent.values():
        if not p["childSkus"] and not p["parentRowCk"]:
            continue
        if not p["jan"]:
            m = re.search(r"(\d{8,14})", p["parentSku"])
            if m:
                p["jan"] = m.group(1)
        p["childCount"] = len(p["childSkus"])
        p["isMix"] = bool(str(p.get("comboSku1") or "").strip())
        out.append(p)
    return out


def mix_marks_from_name(name: str) -> List[str]:
    t = str(name or "")
    m = re.search(r"[（(]([^）)]+)[）)]", t)
    if not m:
        return []
    bits = re.split(r"[・／/,，、\s]+", m.group(1))
    out = []
    seen = set()
    for b in bits:
        b = b.strip()
        if not b or b in seen:
            continue
        if zenkaku_len(b) < 1 or zenkaku_len(b) > 12:
            continue
        seen.add(b)
        out.append(b)
    return out


def mat_blob(materials: Dict[str, Any]) -> str:
    return "\n".join(
        [
            str(materials.get("name") or ""),
            str(materials.get("feats") or ""),
            str(materials.get("kw") or ""),
            str(materials.get("yahoo") or ""),
            str(materials.get("rakuten") or ""),
            str(materials.get("desc") or ""),
            " ".join(str(x) for x in (materials.get("bullets") or [])),
        ]
    )


def has_any(text: str, marks: List[str]) -> bool:
    t = str(text or "")
    return any(m in t for m in marks)


def material_stages(blob: str) -> Dict[str, Any]:
    purpose_hits = [m for m in PURPOSE if m in blob]
    taste_hits = [m for m in TASTE if m in blob]
    # unique preserve
    def uniq(xs: List[str]) -> List[str]:
        s, o = set(), []
        for x in xs:
            if x not in s:
                s.add(x)
                o.append(x)
        return o

    purpose_hits = uniq(purpose_hits)
    taste_hits = uniq(taste_hits)
    return {
        "purposeCount": len(purpose_hits),
        "tasteOnly": len(purpose_hits) == 0 and len(taste_hits) > 0,
        "purposeHits": purpose_hits,
        "tasteHits": taste_hits,
    }


def short_catch(text: str, fallback: str = "") -> str:
    raw = " ".join(str(text or "").split()).strip() or str(fallback or "")
    m = re.search(r"【([^】]{2,40})】", raw)
    if m:
        t = "【" + m.group(1).strip() + "】"
        after = raw[m.end() :].replace(" ", "")
        after = re.sub(r"^[。．、，・／]+", "", after)
        if after:
            t += after
    else:
        t = re.sub(r"\s+", "", raw)
    return t


def ladder_build(materials: Dict[str, Any], theme_id: str, fallback_ev: str) -> Dict[str, str]:
    blob = mat_blob(materials)
    stages = material_stages(blob)
    purpose = stages["purposeHits"]
    taste = stages["tasteHits"]
    marks = list(materials.get("mixMarks") or [])
    # テーマの証拠語を見出し優先（全案が同じ晩酌に潰れない）
    head = str(fallback_ev or "").strip()
    if not head or head not in blob:
        head = purpose[0] if purpose else str(fallback_ev or ("使い方" if theme_id == "T3" else "成果"))
    body_bits = []
    for p in purpose:
        if p != head and p not in body_bits:
            body_bits.append(p)
        if len(body_bits) >= 1:
            break
    if not body_bits and purpose:
        body_bits.append(purpose[0])
    for m in marks[:2]:
        if m not in body_bits:
            body_bits.append(m)
    if taste and taste[0] not in body_bits:
        body_bits.append(taste[0])
    body = "・".join(body_bits) if body_bits else head
    catch = short_catch("【" + re.sub(r"^【|】$", "", head) + "】" + body, head + body)
    sales = re.sub(r"[・／]+", "、", body) or "すぐ使える"
    return {
        "catch": catch,
        "salesTalk": sales,
        "evidence": "／".join(([head] + body_bits)[:4]) or fallback_ev,
    }


def machine_gate(catch: str, materials: Dict[str, Any], theme_id: str) -> Tuple[bool, str]:
    t = str(catch or "")
    if zenkaku_len(t) > CATCH_FALLBACK_MAX:
        return False, f"{int(CATCH_PREF_MAX)}字超"
    if theme_id not in ("T3", "T4"):
        return True, ""
    if not t:
        return False, "空"
    for b in BAN:
        if b in t:
            return False, "禁止語:" + b
    if re.search(r"【\s*(簡単|すぐ本格|本格|すぐ)\s*】", t):
        return False, "採点語を見出し定型にしている"
    return True, ""


def appeal_score(catch: str, materials: Dict[str, Any], theme_id: str) -> Tuple[int, str]:
    if theme_id not in ("T3", "T4"):
        return 0, ""
    t = str(catch or "")
    blob = mat_blob(materials)
    stages = material_stages(blob)
    ok, reason = machine_gate(t, materials, theme_id)
    if not ok:
        return 1, reason or "ゲート不合"
    has_p = has_any(t, PURPOSE)
    has_t = has_any(t, TASTE)
    purpose_first = False
    hm = re.search(r"【([^】]+)】", t)
    if hm and has_any(hm.group(1), PURPOSE):
        purpose_first = True
    if not purpose_first and has_p:
        p_idx = min((t.find(m) for m in PURPOSE if m in t), default=9999)
        t_idx = min((t.find(m) for m in TASTE if m in t), default=9999)
        purpose_first = p_idx < t_idx
    if not purpose_first:
        if has_t:
            return 3, "味が先、または用事が弱い"
        return 1, "意味が分からない"
    if stages["purposeCount"] >= 2:
        return 5, "目的が先。味以外の段が2つ以上材料にあり、味は最後"
    if stages["purposeCount"] >= 1:
        return 4, "目的が先。味以外の段が1つある"
    return 3, "用事語はあるが材料の段が薄い"


def build_catch_burn_rule(catch: str) -> str:
    t = str(catch or "").strip()
    if not t:
        return ""
    m = re.search(r"【([^】]{2,18})】", t)
    if m and BURN_MIN <= zenkaku_len(m.group(1)) <= BURN_MAX:
        return "【" + m.group(1) + "】"
    if m:
        head = m.group(1)
        if zenkaku_len(head) > BURN_MAX:
            head = zenkaku_slice(head, BURN_MAX)
            head = re.sub(r"[てにをはがのと・／]$", "", head)
            return "【" + head + "】"
        after = t[t.index("】") + 1 :].replace(" ", "")
        if after:
            room = BURN_MAX - zenkaku_len(head) - 1
            if room >= 2:
                piece = re.sub(r"[てにをはがのと・／]$", "", zenkaku_slice(after, room))
                if piece:
                    return "【" + head + "】" + piece
        return "【" + head + "】"
    body = re.sub(r"[てにをはがのと・／]$", "", zenkaku_slice(t, BURN_MAX))
    return body if zenkaku_len(body) >= BURN_MIN else t


def burn_looks_ok(burn: str) -> bool:
    b = str(burn or "").strip()
    if not b:
        return False
    ln = zenkaku_len(b)
    if ln < BURN_MIN or ln > BURN_MAX:
        return False
    if looks_incomplete_(b):
        return False
    # 【見出し】だけの短縮は不合格（】の後ろに本文が必要）
    m = re.match(r"^【([^】]+)】(.*)$", b)
    if m and zenkaku_len(m.group(2) or "") < 4:
        return False
    if re.match(r"^【[^】]+】$", b):
        return False
    return True


def copy_length_ok_(text: str, *, pref_min: float, pref_max: float, fallback_max: float) -> Tuple[bool, str]:
    """短い枠優先。不成文でなければフォールバック上限まで可。"""
    t = str(text or "").strip()
    if not t or looks_incomplete_(t):
        return False, "incomplete"
    ln = zenkaku_len(t)
    if ln < pref_min:
        return False, "short"
    if ln <= pref_max:
        return True, "pref"
    if ln <= fallback_max:
        return True, "fallback"
    return False, "long"


def ensure_sales(p: Dict[str, Any], materials: Dict[str, Any]) -> str:
    s = str(p.get("salesTalk") or "").strip()
    raw = str(p.get("salesTalkRaw") or "").strip()
    s = re.sub(r"【[^】]*】", "", s).strip()
    raw = re.sub(r"【[^】]*】", "", raw).strip()

    def whole(t: str) -> str:
        t = re.sub(r"\s+", "", str(t or "").strip())
        if not t or looks_incomplete_(t):
            return ""
        return t

    keep = whole(s) or whole(raw)
    if keep:
        return keep
    return re.sub(r"\s+", "", s or raw or "")


def build_sales_burn(sales_talk: str) -> str:
    t = re.sub(r"\s+", "", str(sales_talk or "").strip())
    if not t:
        return ""
    if SALES_BURN_MIN <= zenkaku_len(t) <= SALES_BURN_MAX and not re.search(
        r"[てにをはがのとでもへや・／、。]$", t
    ):
        return t
    cut = zenkaku_slice(t, SALES_BURN_MAX)
    cut = re.sub(r"[てにをはがのとでもへや・／、。]$", "", cut)
    return cut if zenkaku_len(cut) >= SALES_BURN_MIN else t


def ensure_mix_both(p: Dict[str, Any], marks: List[str]) -> Dict[str, Any]:
    """GAS b5EnsureMixBothOnProposal_ と同じ。セールス全文を魚種名に置換しない。"""
    if len(marks) < 2:
        return p
    need = marks[:2]
    blob = str(p.get("catch") or "") + str(p.get("catchBurn") or "") + str(p.get("salesTalk") or "")
    miss = [m for m in need if m not in blob]
    if not miss:
        return p
    add = "・".join(miss)
    c = str(p.get("catch") or "")
    if add not in c:
        # 字数を壊さないよう追記のみ（全文作り直ししない）
        p["catch"] = (c + ("・" if c else "") + add) if c else f"【セット】{add}"
    s = str(p.get("salesTalk") or "")
    # GAS同様: 短いセールスに足りない区別語を1語足すだけ。置換禁止
    if s and miss[0] not in s and zenkaku_len(s) < 10:
        p["salesTalk"] = s + "・" + miss[0]
    return p


_GEMINI_MODELS = (
    "gemini-3-flash-preview",  # GAS MODEL_GEMINI と揃える
    "gemini-3.8-flash",
)


def _complete_json_objects(raw: str) -> List[Dict[str, Any]]:
    """終わりの括弧が無くても、閉じて完結している人だけ残す。途中の最後は捨てる。"""
    s = str(raw or "")
    start = s.find("[")
    if start < 0:
        return []
    out: List[Dict[str, Any]] = []
    i = start + 1
    n = len(s)
    while i < n:
        while i < n and s[i] in " \n\r\t,":
            i += 1
        if i >= n or s[i] == "]":
            break
        if s[i] != "{":
            break
        begin = i
        depth = 0
        in_str = False
        esc = False
        closed = False
        while i < n:
            c = s[i]
            if in_str:
                if esc:
                    esc = False
                elif c == "\\":
                    esc = True
                elif c == '"':
                    in_str = False
                i += 1
                continue
            if c == '"':
                in_str = True
                i += 1
                continue
            if c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    i += 1
                    closed = True
                    try:
                        obj = json.loads(s[begin:i])
                        if isinstance(obj, dict):
                            out.append(obj)
                    except json.JSONDecodeError:
                        pass
                    break
            i += 1
        if not closed or in_str:
            break
    return out


def _persona_parsed(raw: str) -> List[Dict[str, Any]]:
    m = re.search(r"\[[\s\S]*\]", raw or "")
    if m:
        try:
            got = json.loads(m.group(0))
            if isinstance(got, list):
                return [x for x in got if isinstance(x, dict)]
        except json.JSONDecodeError:
            pass
    return _complete_json_objects(raw)


def _joined_gemini_text(resp: Any) -> str:
    cands = getattr(resp, "candidates", None) or []
    if cands:
        content = getattr(cands[0], "content", None)
        parts = getattr(content, "parts", None) or []
        bits = []
        for part in parts:
            if getattr(part, "thought", False) is True:
                continue
            t = getattr(part, "text", None) or ""
            if t:
                bits.append(str(t))
        if bits:
            return "".join(bits).strip()
    return str(getattr(resp, "text", None) or "").strip()


def gemini_text(prompt: str, client: Any, model: str = "", thinking: str = "") -> str:
    models = (model,) if model else _GEMINI_MODELS
    last_err = ""
    config = None
    if thinking:
        try:
            from google.genai import types as genai_types

            config = genai_types.GenerateContentConfig(
                thinking_config=genai_types.ThinkingConfig(thinking_level=thinking)
            )
        except Exception as e:
            LOG.warning("thinking config skip: %s", e)
    for mid in models:
        if not mid:
            continue
        try:
            kwargs = {"model": mid, "contents": prompt}
            if config is not None:
                kwargs["config"] = config
            resp = client.models.generate_content(**kwargs)
            text = _joined_gemini_text(resp)
            if text:
                return text
        except Exception as e:
            last_err = str(e)
            LOG.warning("gemini_text fail model=%s: %s", mid, e)
            continue
    if last_err:
        LOG.warning("gemini_text all models failed last=%s", last_err)
    return ""


def _parse_self_appeal_(raw: str) -> Tuple[int, str]:
    t = str(raw or "")
    i0 = t.find("{")
    i1 = t.rfind("}")
    if i0 < 0 or i1 <= i0:
        return 0, ""
    try:
        obj = json.loads(t[i0 : i1 + 1])
    except Exception:
        return 0, ""
    try:
        n = int(obj.get("appealScore") or 0)
    except (TypeError, ValueError):
        return 0, ""
    if n < 1 or n > 5:
        return 0, ""
    return n, str(obj.get("appealReason") or "")[:40]


def parse_catch_sales_json(raw: str) -> Tuple[str, str]:
    """GAS b5ParseOneGeminiCatchSales_ 相当。"""
    t = str(raw or "").strip()
    if not t:
        return "", "", ""
    if "```" in t:
        a = t.find("```")
        b = t.find("```", a + 3)
        block = t[a + 3 : b] if b > a else t[a + 3 :]
        t = re.sub(r"^json\s*", "", block.strip(), flags=re.I)
    i0 = t.find("{")
    i1 = t.rfind("}")
    if i0 >= 0 and i1 > i0:
        try:
            obj = json.loads(t[i0 : i1 + 1])
            return (
                _strip_catch_brackets(str(obj.get("catch") or "")),
                _strip_sales_brackets(str(obj.get("salesTalk") or "")),
                str(obj.get("lenWhy") or "").replace("\n", " ")[:180],
            )
        except Exception:
            pass
    m = re.search(r"\{[^{}]*\"catch\"[^{}]*\}", t, re.S)
    if m:
        try:
            obj = json.loads(m.group(0))
            return (
                _strip_catch_brackets(str(obj.get("catch") or "")),
                _strip_sales_brackets(str(obj.get("salesTalk") or "")),
                str(obj.get("lenWhy") or "").replace("\n", " ")[:180],
            )
        except Exception:
            pass
    return t.splitlines()[0].strip(), "", ""


def materials_prompt_block(materials: Dict[str, Any]) -> str:
    lines = [
        f"商品名: {materials.get('name') or ''}",
        f"メーカー: {materials.get('maker') or ''}",
        f"特徴: {materials.get('feats') or ''}",
        f"用途KW: {materials.get('kw') or ''}",
        f"楽天キャッチ: {materials.get('rakuten') or ''}",
        f"Yahooキャッチ: {materials.get('yahoo') or ''}",
        f"箇条書き: {'／'.join(str(x) for x in (materials.get('bullets') or []))}",
        f"説明: {materials.get('desc') or ''}",
        f"構成区別語: {'・'.join(materials.get('mixMarks') or [])}",
    ]
    return "\n".join(lines)


def polish_one(
    materials: Dict[str, Any],
    p: Dict[str, Any],
    avoid: List[str],
    *,
    client: Any,
    use_gemini: bool,
    used_angles: Optional[List[str]] = None,
    force_body: bool = False,
    len_retry: str = "",
) -> Dict[str, Any]:
    """整文プロンプトは GAS b5GeminiCatchOneTheme_ に揃える（短縮しない）。"""
    if not use_gemini or client is None:
        return p
    used_angles = used_angles or []
    tid = str(p.get("themeId") or "")
    cid = str(p.get("classId") or "")
    errand = errand_of(p)
    avoid_all = list(avoid) + list(used_angles)
    ban_extra = "・".join(_BAN_IN_COPY)
    prompt = (
        f"楽天MAINのキャッチを、このテーマだけ、全角{int(CATCH_PREF_MIN)}〜{int(CATCH_PREF_MAX)}字で1つ作ってください（他テーマは考えない）。セールスは別の補足です。\n"
        f"テーマ: {tid} / {cid} / {p.get('themeName') or ''}\n"
        f"根拠: {p.get('evidence') or ''}\n"
        f"角度: {p.get('sceneHint') or ''}\n"
        f"商品名: {materials.get('name') or ''}\n"
        f"この案の状況: {p.get('situation') or ''}\n"
        f"この案の得たい結果: {errand}\n"
    )
    if avoid_all:
        prompt += "前の案と同じ方向・同じ文言の並びは禁止: " + json.dumps(avoid_all[:12], ensure_ascii=False) + "\n"
    prompt += (
        "この順で考える。前の段を飛ばさない。\n"
        "1. 状況、困り、得たい結果、訴求理由を読む。得たい結果は状況の場から。四つは同じ話にする。\n"
        "2. その人に向けるキャッチとセールスを書く。見出しは状況が見えること。材料に無い場は書かない。商品名とメーカー名は見出しにしない。体の効能をうたう語は書かない。健康、栄養、成長、糖質、ダイエット、血圧、免疫、罪悪感、高タンパク、ヘルシーは例で、同じ種類も書かない。お得、最強、開ける、のせる、並べる、は見出しに出さない。体の効能でない材料の語は、その人の訴求なら使ってよい。\n"
        "3. 材料は、その訴求を支えるときだけ使う。材料の文の言い換えから始めない。\n"
        "一枚で、その人の得たい結果が見える。文として完結しなくても、イメージが止まればよい。\n"
        f"訴求方針: {p.get('appealPolicy') or ''}\n"
    )
    prompt += "条件:\n"
    prompt += (
        f"- 【】は付けない。キャッチは全角{int(CATCH_PREF_MIN)}〜{int(CATCH_PREF_MAX)}字の1本。補足はセールスに書く。\n"
    )
    prompt += (
        f"- catchは全角{int(CATCH_PREF_MIN)}〜{int(CATCH_PREF_MAX)}字。この範囲を超えない。"
        "語の途中で切らない。画面に出す文＝画像に焼く文（別短縮しない）\n"
        "- 完成料理名は書かない。添える料理は別欄\n"
    )
    prompt += f"ペルソナ短文: {p.get('personaShort') or ''}\n"
    prompt += f"ペルソナ詳細:\n{p.get('personaDetail') or ''}\n"
    prompt += "材料（支えるときだけ）:\n" + materials_prompt_block(materials) + "\n"
    if tid == "N1":
        prompt += "- N1: 商品名フル＋特徴（味・製法）。名前だけ／メーカーだけは不可。特徴も同じ文に書く\n"
    else:
        prompt += (
            "- T1〜T4: 商品名フル・内容量・シリーズ定番名の繰り返し禁止。"
            "構成の語は、その訴求を支えるときだけ使う。\n"
        )
    if tid in ("T3", "T4"):
        prompt += (
            f"- 【目的の階段】キャッチはペルソナの得たい結果が分かる全角{int(CATCH_PREF_MIN)}〜{int(CATCH_PREF_MAX)}字。味だけで終わらせない\n"
            "- 「簡単」「すぐ本格」「本格」を見出しの定型にしない（採点語のみ）\n"
            "- 手順の動作（開ける・のせる・並べる等）・体験談禁止。"
            "体の効能をうたう語は書かない。健康、栄養、成長、糖質、ダイエット、血圧、免疫、罪悪感、高タンパク、ヘルシーは例で、同じ種類も書かない。お得、最強の継ぎ足し禁止。\n"
        )
    prompt += (
        "- 困りと得たい結果から、その人が買おうと思える文を書く。材料の文の言い換えから始めない\n"
        f"- salesTalk（全角{int(SALES_PREF_MIN)}〜{int(SALES_PREF_MAX)}字希望。"
        f"成立しないときだけ最大{int(SALES_FALLBACK_MAX)}字）。"
        "止まったあとに読む決断の補足。catch と同じ文の複製は禁止。"
        "【】見出しは付けない（キャッチ側の役割）。別欄への短縮はしない\n"
        "- salesTalk は魚種名・構成名だけの列挙にしない（それはキャッチ側で両方示す）\n"
        f"- 禁語（本文に出さない）: {ban_extra}\n"
        "- ペルソナの用事と catch／salesTalk／場面は同じ用事でつなぐ。用事が先（味だけ先頭禁止）\n"
        "- 材料は訴求を支えるときだけ使う。材料に無い効能・累計・受賞・No.1は書かない。書いたあとに材料の語を足さない\n"
        "- lenWhyは、残した語、捨てた語、全角の字数。\n"
    )
    if len_retry == "long":
        prompt += f"- 前回は長すぎた。全角{int(CATCH_PREF_MIN)}〜{int(CATCH_PREF_MAX)}字に収める。\n"
    elif len_retry == "short":
        prompt += f"- 前回は短すぎた。全角{int(CATCH_PREF_MIN)}〜{int(CATCH_PREF_MAX)}字にする。\n"
    if tid in ("T3", "T4"):
        prompt += (
            "- appealScoreはこの文だけを1〜5で自己採点。"
            "5=用事が先で意味がすぐ分かる。4=用事が先で味以外が1段。"
            "3=おいしくなるで止まっている。2=味の擬音だけ。1=名前の羅列か意味が分からない。"
            "材料の語数で点数を上げない。appealReasonは20字以内\n"
        )
    prompt += (
        f'- 出力は JSON のみ: {{"themeId":"{tid}","classId":"{cid}",'
        '"catch":"……","salesTalk":"……","lenWhy":"……","appealScore":5,"appealReason":"……"}'
    )
    if len_retry:
        LOG.info("catch_len_ retry=%s theme=%s", len_retry, tid)
    raw = gemini_text(prompt, client)
    p["geminiRawSnippet"] = zenkaku_slice(re.sub(r"\s+", " ", raw), 80)
    catch, sales, len_why = parse_catch_sales_json(raw)
    if catch and zenkaku_len(catch) < CATCH_PREF_MIN and not len_retry:
        return polish_one(
            materials,
            p,
            avoid,
            client=client,
            use_gemini=use_gemini,
            used_angles=used_angles,
            len_retry="short",
        )
    if catch and zenkaku_len(catch) >= CATCH_PREF_MIN:
        catch = strip_ban_copy(catch)
        p["catch"] = catch
        if len_why:
            p["catchLenWhy"] = len_why
        p["catchSource"] = "gemini"
        ap_n, ap_reason = _parse_self_appeal_(raw)
        if ap_n:
            p["appealScore"] = ap_n
            p["appealReason"] = ap_reason
            p["appealFromGemini"] = True
        if sales:
            sales = strip_ban_copy(sales)
            p["salesTalkRaw"] = sales
            p["salesTalk"] = sales
    return p


def build_catch_burn_smart(
    catch: str,
    materials: Dict[str, Any],
    p: Dict[str, Any],
    *,
    client: Any,
    use_gemini: bool,
) -> str:
    fallback = build_catch_burn_rule(catch)
    t = str(catch or "").strip()
    if not t:
        return ""
    if zenkaku_len(t) <= BURN_MAX and burn_looks_ok(t):
        return t
    naive = naive_slice_burn(t, BURN_MAX)
    if not use_gemini or client is None:
        return fallback if burn_looks_ok(fallback) else (naive if burn_looks_ok(naive) else fallback)
    errand = errand_of(p)
    prompt = (
        "EC商品画像の大きい字用に短縮してください。日本語。意味単位。語の途中で切らない。\n"
        f"全角{int(BURN_MIN)}〜{int(BURN_MAX)}字。【】は付けない。助詞や「・」で終わらせない。\n"
        f"用事（見出しに残す）: {errand}\n"
        f"ペルソナ: {p.get('personaShort') or ''}\n"
        f"元文: {t}\n"
        "味だけ見出しにしない。返答は短縮文だけ。JSON不要。"
    )
    raw = gemini_text(prompt, client)
    raw = re.sub(r"[\r\n]+", " ", raw)
    raw = re.sub(r'^["「]|["」]$', "", raw).strip()
    if burn_looks_ok(raw):
        return raw
    if burn_looks_ok(fallback) and fallback != naive:
        return fallback
    return fallback if burn_looks_ok(fallback) else zenkaku_slice(t, int(CATCH_PREF_MAX))


def propose_rule(materials: Dict[str, Any]) -> List[Dict[str, Any]]:
    blob = mat_blob(materials)
    name = str(materials.get("name") or "商品")
    feats = str(materials.get("feats") or "").replace("／", "・")
    feat_bits = [x.strip() for x in re.split(r"[・]", feats) if x.strip()][:2]
    n1 = "【" + zenkaku_slice(name, 18) + "】" + (feat_bits[0] if feat_bits else "")
    raw: List[Dict[str, Any]] = [
        {
            "themeId": "N1",
            "themeName": "商品名明示",
            "catch": short_catch(n1, name),
            "salesTalk": zenkaku_slice(feat_bits[0] if feat_bits else "特徴が分かる", 10),
            "sceneHint": "袋＋商品名が大きく読める。特徴（味・製法）も短く明示。",
            "evidence": name + ("／" + feats if feats else ""),
            "classId": "name",
            "sellScore": 10,
        }
    ]
    t1 = ["受賞", "金賞", "累計", "突破", "億粒", "万枚", "万袋"]
    t2 = ["国産", "産地", "原料", "かつお節", "本枯", "枕崎", "鹿児島"]
    t3 = ["お茶漬け", "おつまみ", "振る", "かける", "晩酌", "おやつ", "お茶請け"]
    if has_any(blob, t1):
        ev = next(m for m in t1 if m in blob)
        raw.append(
            {
                "themeId": "T1",
                "themeName": "実績・受賞",
                "catch": short_catch("【" + ev + "】" + ev, ev),
                "salesTalk": zenkaku_slice(ev, 10),
                "sceneHint": "袋＋シズル。トロフィーは作らない。",
                "evidence": ev,
                "classId": "proof",
                "sellScore": 70,
            }
        )
    if has_any(blob, t2):
        ev = next(m for m in t2 if m in blob)
        raw.append(
            {
                "themeId": "T2",
                "themeName": "産地・素材",
                "catch": short_catch("【" + ev + "】素材の信頼", ev),
                "salesTalk": zenkaku_slice(ev, 10),
                "sceneHint": "原料の静物。袋は残す。",
                "evidence": ev,
                "classId": "origin",
                "sellScore": 30,
            }
        )
    stages = material_stages(blob)
    if has_any(blob, t3) or stages["purposeCount"] > 0:
        ev = next((m for m in t3 if m in blob), stages["purposeHits"][0] if stages["purposeHits"] else "使い方")
        ladder = ladder_build(materials, "T3", ev)
        sc, reason = appeal_score(ladder["catch"], materials, "T3")
        raw.append(
            {
                "themeId": "T3",
                "themeName": "使い方・シーン",
                "catch": ladder["catch"],
                "salesTalk": ladder["salesTalk"],
                "sceneHint": f"キャッチにある食べ方（{ev}）。用事が先・味は最後。",
                "evidence": ladder["evidence"] or ev,
                "classId": "usage",
                "sellScore": 50,
                "appealScore": sc,
                "appealReason": reason,
            }
        )
    sell_classes = [
        ("time_save", 100, ["時短", "すぐ", "さっと", "だしをとらず", "一品", "手軽"], "時短の食卓"),
        ("drink_snack", 80, ["晩酌", "おつまみ", "酒", "ビール", "お酒"], "晩酌が華やぐ"),
        ("tea_snack", 60, ["お茶請け", "おやつ", "お茶うけ"], "お茶請け・おやつ"),
        ("table_expand", 40, ["食卓が広が", "両用", "お茶漬け", "おつまみ"], "食卓が広がる"),
    ]
    if not stages["tasteOnly"]:
        for cid, w, needles, label in sell_classes:
            if cid == "table_expand":
                if not (
                    ("お茶漬け" in blob and "おつまみ" in blob)
                    or "食卓が広が" in blob
                    or "両用" in blob
                ):
                    continue
            if not has_any(blob, needles):
                continue
            ev = next((n for n in needles if n in blob), label)
            ladder = ladder_build(materials, "T4", ev)
            sc, reason = appeal_score(ladder["catch"], materials, "T4")
            raw.append(
                {
                    "themeId": "T4",
                    "themeName": "成果・" + label,
                    "catch": ladder["catch"],
                    "salesTalk": ladder["salesTalk"],
                    "sceneHint": f"買うと良い未来（{label}）。用事が先・味は最後。",
                    "evidence": ladder["evidence"] or ev,
                    "classId": cid,
                    "sellScore": w,
                    "appealScore": sc,
                    "appealReason": reason,
                }
            )

    raw.sort(key=lambda x: int(x.get("sellScore") or 0), reverse=True)
    child = int(materials.get("childCount") or 0)
    target = child if child >= 2 else 4
    target = min(max(target, 1), PROPOSE_MAX)
    personas = assign_personas(target, bool(materials.get("isMix")))
    out: List[Dict[str, Any]] = []
    seen_class: set[str] = set()
    seen_theme: set[str] = set()

    def take(row: Dict[str, Any], force_class: bool = False) -> None:
        key = f"{row['themeId']}:{row['classId']}"
        if key in seen_class:
            return
        if not force_class and row["themeId"] in seen_theme:
            return
        seen_class.add(key)
        if not force_class:
            seen_theme.add(row["themeId"])
        r = dict(row)
        r["catchRule"] = r.get("catch")
        r["sellRank"] = len(out) + 1
        per = personas[len(out)] if len(out) < len(personas) else personas[0]
        r["personaId"] = per["id"]
        r["personaShort"] = per["short"]
        r["personaDetail"] = per["detail"]
        r["errand"] = per.get("errand") or ""
        r["personaErrand"] = r["errand"]
        r["attempts"] = 1
        r["ladderRewrites"] = 0
        out.append(r)

    for row in raw:
        take(row, force_class=False)
        if len(out) >= target:
            break
    if len(out) < target:
        for row in raw:
            take(row, force_class=True)
            if len(out) >= target:
                break
    # 足りなければ用途バリエーションで埋める（同じ材料の言い換え）
    while len(out) < target:
        base = dict(out[-1] if out else raw[0])
        per = personas[len(out)] if len(out) < len(personas) else personas[len(out) % len(personas)]
        ladder = ladder_build(materials, "T4", base.get("evidence") or "一品")
        fill = {
            "themeId": "T4",
            "themeName": "成果・補充",
            "catch": ladder["catch"],
            "salesTalk": ladder["salesTalk"],
            "sceneHint": "袋が読める。用事が先。",
            "evidence": ladder["evidence"],
            "classId": f"fill_{len(out)}",
            "sellScore": 20,
            "personaId": per["id"],
            "personaShort": per["short"],
            "personaDetail": per["detail"],
            "errand": per.get("errand") or "",
            "personaErrand": per.get("errand") or "",
            "sellRank": len(out) + 1,
            "catchRule": ladder["catch"],
            "attempts": 1,
            "ladderRewrites": 0,
        }
        out.append(fill)

    for p in out:
        p["salesTalk"] = strip_ban_copy(ensure_sales(p, materials))
        p["sceneHint"] = scene_from_persona(p, materials)
    return out


def _catch_head(text: str) -> str:
    m = re.search(r"【([^】]{1,40})】", str(text or ""))
    return m.group(1).strip() if m else str(text or "").strip()


def _headline_local(p: Dict[str, Any], materials: Dict[str, Any]) -> str:
    t = str(p.get("catch") or "")
    if not t:
        return "空"
    if zenkaku_len(t) > CATCH_FALLBACK_MAX:
        return f"{int(CATCH_PREF_MAX)}字超"
    if str(p.get("themeId") or "") in ("T3", "T4"):
        ok, why = machine_gate(t, materials, str(p.get("themeId") or ""))
        if not ok:
            return why or "ゲート不合"
    return ""


def _headline_story(p: Dict[str, Any], materials: Dict[str, Any], client: Any) -> str:
    got = _gemini_array(
        "見出しが得たい結果と同じ話か、材料に無い効能・累計・受賞・No.1を足しているかを返す。JSON配列だけ。\n"
        '[{"i":1,"same":true,"extra":false}]\n'
        "sameが false なら見出しが得たい結果と違う。extraが true なら材料に無い効能を足している。\n"
        f"材料:\n{materials_prompt_block(materials)}\n"
        f"案:\n1. 得たい結果:{errand_of(p)} 見出し:{_catch_head(str(p.get('catch') or ''))}\n",
        client,
    )
    if not got:
        return ""
    row = got[0]
    if row.get("same") is not True:
        return "見出しが得たい結果と違う"
    if row.get("extra") is True:
        return "材料に無い効能"
    return ""


def _headline_one(
    materials: Dict[str, Any],
    p: Dict[str, Any],
    *,
    client: Any,
    use_gemini: bool,
) -> None:
    reason = _headline_local(p, materials) or _headline_story(p, materials, client)
    if not reason:
        return
    LOG.info("headline_ rewrite reason=%s", reason)
    p2 = polish_one(
        materials,
        dict(p),
        [str(p.get("catch") or "")],
        client=client,
        use_gemini=use_gemini,
        len_retry="long" if reason == f"{int(CATCH_PREF_MAX)}字超" else "",
    )
    got = str(p2.get("catch") or "")
    if zenkaku_len(got) >= CATCH_PREF_MIN:
        p["catch"] = got
        if p2.get("salesTalk"):
            p["salesTalk"] = strip_ban_copy(str(p2.get("salesTalk")))
        p["catchSource"] = "gemini_headline"
    why = _headline_local(p, materials) or _headline_story(p, materials, client)
    if why:
        p["burnSkip"] = True
        p["burnSkipReason"] = why
        LOG.info("burn_skip_ reason=%s", why)
    else:
        LOG.info("headline_ keep")


def polish_batch(
    materials: Dict[str, Any],
    proposals: List[Dict[str, Any]],
    *,
    client: Any,
    use_gemini: bool,
    force_refresh_official: bool = False,
) -> List[Dict[str, Any]]:
    ladder_rewrites = 0
    used_angles: List[str] = []
    used_dishes: List[str] = []
    used_layouts: List[str] = []
    recipe_pool: List[Dict[str, str]] = []
    fill_personas_from_materials(
        materials, proposals, client=client, use_gemini=use_gemini
    )
    for p in proposals:
        p["sceneHint"] = scene_from_persona(p, materials)
    if use_gemini and client is not None:
        recipe_pool = gemini_recipe_pool_(
            materials,
            client=client,
            need=max(10, len(proposals) + 3),
            force_refresh_official=force_refresh_official,
        )
        time.sleep(0.2)
    for i, p in enumerate(proposals):
        avoid = [str(q.get("catch") or "") for j, q in enumerate(proposals) if j != i and q.get("catch")]
        before = str(p.get("catch") or "")
        p = polish_one(
            materials,
            p,
            avoid,
            client=client,
            use_gemini=use_gemini,
            used_angles=used_angles,
        )
        p["attempts"] = int(p.get("attempts") or 1)
        if (p.get("themeId") in ("T3", "T4")) and ladder_rewrites < LADDER_REWRITE_MAX:
            if p.get("appealFromGemini"):
                sc0, reason0 = int(p.get("appealScore") or 1), str(p.get("appealReason") or "")
            else:
                sc0, reason0 = appeal_score(str(p.get("catch") or ""), materials, str(p.get("themeId")))
            ok0, gate_reason = machine_gate(str(p.get("catch") or ""), materials, str(p.get("themeId")))
            if not ok0:
                sc0, reason0 = 1, gate_reason
            p["appealScore"] = sc0
            p["appealReason"] = reason0
            # 共有の作り直しは使わない。見出しは案ごとに1回、_headline_one で見る。
            if False and (not ok0 or sc0 < 4) and ladder_rewrites < LADDER_REWRITE_MAX:
                ladder_rewrites += 1
                p["ladderRewrites"] = int(p.get("ladderRewrites") or 0) + 1
                p["attempts"] = int(p.get("attempts") or 1) + 1
                time.sleep(0.2)
                p2 = polish_one(
                    materials,
                    dict(p),
                    avoid + [str(p.get("catch") or ""), before],
                    client=client,
                    use_gemini=use_gemini,
                    used_angles=used_angles,
                )
                got = str(p2.get("catch") or "")
                if got:
                    okL, gateL = machine_gate(got, materials, str(p.get("themeId")))
                    if p2.get("appealFromGemini"):
                        scL, reasonL = int(p2.get("appealScore") or 1), str(p2.get("appealReason") or "")
                    else:
                        scL, reasonL = appeal_score(got, materials, str(p.get("themeId")))
                    if not okL:
                        scL, reasonL = 1, gateL
                    if okL and scL >= sc0:
                        p["catch"] = got
                        if p2.get("salesTalk"):
                            p["salesTalk"] = strip_ban_copy(str(p2.get("salesTalk")))
                        p["catchSource"] = "gemini_ladder"
                        p["appealScore"] = scL
                        p["appealReason"] = reasonL
                        p["appealFromGemini"] = bool(p2.get("appealFromGemini"))
        p["salesTalk"] = strip_ban_copy(ensure_sales(p, materials))
        p["sceneHint"] = scene_from_persona(p, materials)
        marks = list(materials.get("mixMarks") or [])
        # 魚種名だけのセールスは不合格扱い→Gemini生 or 用事補足へ戻す
        st = str(p.get("salesTalk") or "").strip()
        marks_join = "・".join(marks[:2]) if len(marks) >= 2 else ""
        if marks_join and (
            st == marks_join or re.sub(r"[・／/\s]", "", st) == re.sub(r"[・／/\s]", "", marks_join)
        ):
            raw_s = strip_ban_copy(str(p.get("salesTalkRaw") or ""))
            if raw_s and raw_s != marks_join:
                p["salesTalk"] = raw_s
            else:
                err = errand_of(p) or "すぐ使える"
                p["salesTalk"] = f"{err}がすぐ分かる"
            p["salesTalkFixed"] = "marks_only_rejected"
        if p.get("themeId") in ("T3", "T4") and not p.get("appealFromGemini"):
            sc, reason = appeal_score(str(p.get("catch") or ""), materials, str(p.get("themeId")))
            ok_end, why_end = machine_gate(str(p.get("catch") or ""), materials, str(p.get("themeId")))
            if not ok_end:
                sc, reason = 1, why_end
            p["appealScore"] = sc
            p["appealReason"] = reason
        elif p.get("themeId") in ("T3", "T4"):
            ok_end, why_end = machine_gate(str(p.get("catch") or ""), materials, str(p.get("themeId")))
            if not ok_end:
                p["appealScore"] = 1
                p["appealReason"] = why_end
        # 画面＝焼成（転換なし）。場面／添える料理だけ別生成
        p["catch"] = _strip_catch_brackets(str(p.get("catch") or ""))
        ok_c, len_c = copy_length_ok_(
            str(p.get("catch") or ""),
            pref_min=CATCH_PREF_MIN,
            pref_max=CATCH_PREF_MAX,
            fallback_max=CATCH_FALLBACK_MAX,
        )
        ok_s, len_s = copy_length_ok_(
            str(p.get("salesTalk") or ""),
            pref_min=SALES_PREF_MIN,
            pref_max=SALES_PREF_MAX,
            fallback_max=SALES_FALLBACK_MAX,
        )
        p["catchLenMode"] = len_c
        p["salesLenMode"] = len_s
        # 同一文を画像用に使う（書いたあとに短くしない）。見出しが落ちた案は画面に残し、焼成には使わない
        if use_gemini and client is not None:
            _headline_one(materials, p, client=client, use_gemini=use_gemini)
        if p.get("burnSkip"):
            p["catchBurn"] = ""
            LOG.info("burn_skip_ reason=%s", p.get("burnSkipReason") or "")
        else:
            p["catchBurn"] = str(p.get("catch") or "")
        p["salesBurn"] = str(p.get("salesTalk") or "")
        p["burnSource"] = "identity"
        p["salesTalkImage"] = p["salesBurn"]
        scene_pack = None
        if use_gemini and client is not None:
            scene_pack = gemini_burns_scene_dish_(
                materials,
                p,
                client=client,
                recipe_pool=recipe_pool,
                used_dishes=used_dishes,
                used_layouts=used_layouts,
            )
            time.sleep(0.15)
        if scene_pack:
            # キャッチ／セールスは上で確定済み。場面と料理だけ採用
            p["dishName"] = scene_pack.get("dishName") or ""
            p["dishEvidence"] = scene_pack.get("dishEvidence") or ""
            p["dishSourceUrl"] = scene_pack.get("dishSourceUrl") or ""
            p["sceneHint"] = scene_pack.get("sceneHint") or p.get("sceneHint")
            if scene_pack.get("sceneLinkWhy"):
                p["sceneLinkWhy"] = scene_pack.get("sceneLinkWhy")
            p["sceneLayout"] = scene_pack.get("sceneLayout") or ""
            if p["dishName"]:
                used_dishes.append(str(p["dishName"]))
            if p.get("sceneLayout"):
                used_layouts.append(str(p["sceneLayout"]))
        else:
            locked = p.get("_lockedDish") or {}
            if locked.get("dishName"):
                p["dishName"] = locked["dishName"]
                p["dishEvidence"] = locked.get("evidence") or ""
                p["dishSourceUrl"] = locked.get("sourceUrl") or ""
                used_dishes.append(locked["dishName"])
            else:
                p["dishName"] = ""
                p["dishEvidence"] = ""
                p["dishSourceUrl"] = ""
            layout_candidates = [x for x in _SCENE_LAYOUTS if x not in used_layouts] or list(_SCENE_LAYOUTS)
            p["sceneLayout"] = layout_candidates[0]
            used_layouts.append(p["sceneLayout"])
            p["sceneHint"] = scene_from_persona(p, materials)
        if p.get("appealPolicy"):
            used_angles.append(str(p.get("appealPolicy")))
        angle = re.search(r"【([^】]+)】", str(p.get("catch") or ""))
        if angle:
            used_angles.append(angle.group(1))
        used_angles.append(str(p.get("salesTalk") or ""))
        proposals[i] = p
        if use_gemini:
            time.sleep(0.15)
    for p in proposals:
        if p.get("dishName") or dish_optional_for_errand_(errand_of(p)):
            continue
        LOG.info("leftover empty")
    return proposals


def load_materials_for_parent(parent: Dict[str, Any]) -> Dict[str, Any]:
    """親SKU行の商品名ベースを正。JAN検索の単品名で上書きしない。"""
    jan = str(parent.get("jan") or "")
    name_base = str(parent.get("nameBase") or "").strip()
    pack = _compact_facts_from_sheets(jan) if jan else {"materials": {}, "error": "JAN無し"}
    materials = dict(pack.get("materials") or {})
    # ミックス親の表示名はマスタ親行を優先（JAN一致が単品行のことがある）
    if name_base:
        materials["name"] = name_base
    elif not materials.get("name"):
        materials["name"] = ""
    materials["childCount"] = int(parent.get("childCount") or 0)
    materials["isMix"] = bool(parent.get("isMix"))
    materials["category"] = parent.get("category") or materials.get("category") or ""
    materials["parentSku"] = parent.get("parentSku") or ""
    materials["jan"] = jan or str(materials.get("jan") or "")
    materials["comboSku1"] = parent.get("comboSku1") or ""
    marks = mix_marks_from_name(name_base or str(materials.get("name") or ""))
    if len(marks) < 2:
        # feats からも補完（・区切り）
        extra = mix_marks_from_name("（" + str(materials.get("feats") or "") + "）")
        for m in extra:
            if m not in marks:
                marks.append(m)
    materials["mixMarks"] = marks
    materials["_sheetsError"] = pack.get("error") or ""
    materials["_aiNameWas"] = str((pack.get("materials") or {}).get("name") or "")
    return materials


def format_human(parent: Dict[str, Any], materials: Dict[str, Any], proposals: List[Dict[str, Any]]) -> str:
    lines = []
    lines.append(f"## {parent.get('parentSku')} / JAN {parent.get('jan')}")
    lines.append(f"商品名: {materials.get('name')}")
    lines.append(
        f"子SKUレ点={materials.get('childCount')} 案数={len(proposals)} "
        f"mix={materials.get('isMix')} marks={materials.get('mixMarks')}"
    )
    lines.append("")
    for p in proposals:
        lines.append(f"### 案{p.get('sellRank')} {p.get('themeId')} {p.get('themeName')} "
                     f"(訴求度={p.get('appealScore') or '-'} 試行={p.get('attempts')})")
        lines.append(f"- 用事: {p.get('errand') or p.get('personaErrand') or ''}")
        lines.append(f"- ペルソナ短文: {p.get('personaShort')}")
        lines.append(f"- ペルソナ詳細:\n{p.get('personaDetail')}")
        lines.append(
            f"- キャッチ（画面＝焼成）: {p.get('catch')} "
            f"(字数={p.get('catchLenMode') or '-'})"
        )
        lines.append(
            f"- セールス（画面＝焼成）: {p.get('salesTalk')} "
            f"(字数={p.get('salesLenMode') or '-'})"
        )
        if p.get("salesTalkRaw") and p.get("salesTalkRaw") != p.get("salesTalk"):
            lines.append(f"- Gemini生セールス: {p.get('salesTalkRaw')}")
        if p.get("salesTalkFixed"):
            lines.append(f"- セールス修正: {p.get('salesTalkFixed')}")
        lines.append(f"- 添える料理: {p.get('dishName') or '（なし）'}")
        if p.get("dishEvidence"):
            lines.append(f"- 料理根拠: {p.get('dishEvidence')}")
        url = str(p.get("dishSourceUrl") or "").strip()
        if url:
            lines.append(f"- 料理ソース: [{p.get('dishName') or '確認'}]({url})")
        elif p.get("dishName"):
            lines.append("- 料理ソース: （URLなし・不合格）")
        else:
            lines.append("- 料理ソース: （料理なし）")
        lines.append(f"- 場面レイアウト: {p.get('sceneLayout') or '-'}")
        lines.append(f"- 場面: {p.get('sceneHint')}")
        lines.append(f"- catchSource: {p.get('catchSource') or 'rule'}")
        if p.get("appealReason"):
            lines.append(f"- 訴求度理由: {p.get('appealReason')}")
        lines.append("")
    return "\n".join(lines)


def run_one(
    parent: Dict[str, Any],
    *,
    use_gemini: bool,
    client: Any,
    force_refresh_official: bool = False,
) -> Dict[str, Any]:
    t0 = time.time()
    materials = load_materials_for_parent(parent)
    proposals = propose_rule(materials)
    proposals = polish_batch(
        materials,
        proposals,
        client=client,
        use_gemini=use_gemini,
        force_refresh_official=force_refresh_official,
    )
    elapsed = time.time() - t0
    LOG.info(
        "run_one_done parent=%s sec=%.1f dishes_with_url=%s officialSrc=%s",
        parent.get("parentSku"),
        elapsed,
        sum(1 for p in proposals if re.match(r"^https?://", str(p.get("dishSourceUrl") or ""), re.I)),
        materials.get("_officialUrlSource") or "-",
    )
    return {
        "parent": parent,
        "materials": {
            k: materials[k]
            for k in (
                "name",
                "maker",
                "feats",
                "kw",
                "childCount",
                "isMix",
                "category",
                "mixMarks",
                "parentSku",
                "jan",
                "_sheetsError",
                "_officialUrl",
                "_officialUrlSource",
            )
            if k in materials
        },
        "proposals": proposals,
        "human": format_human(parent, materials, proposals),
        "elapsedSec": round(elapsed, 1),
    }


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="B-④案画面相当仮想（マスタ出品CK）")
    ap.add_argument("--parent-sku", default="", help="1親だけ。空なら出品CK親すべて")
    ap.add_argument("--no-gemini", action="store_true", help="ルール案のみ")
    ap.add_argument("--out", default="", help="出力JSONパス")
    ap.add_argument(
        "--recipe-timing",
        action="store_true",
        help="料理ソース取得だけ計測（案整文なし）。競合あり／なしを比較",
    )
    ap.add_argument(
        "--refresh-official-url",
        action="store_true",
        help="公式URLキャッシュを使わず再検索（404・移転時）",
    )
    args = ap.parse_args(argv)

    parents = list_checked_parents()
    if args.parent_sku:
        want = str(args.parent_sku).strip()
        parents = [p for p in parents if p["parentSku"] == want]
        if not parents:
            raise SystemExit(f"出品CKに親SKUがありません: {want}")
    if not parents:
        raise SystemExit("出品CKの親が0件です")

    use_gemini = not args.no_gemini
    client = None
    if use_gemini:
        try:
            from gemini_image import make_client

            client = make_client()
        except Exception as e:
            LOG.warning("Gemini無効化: %s", e)
            use_gemini = False

    if args.recipe_timing:
        parent = parents[0]
        materials = load_materials_for_parent(parent)
        rows = []
        for label, use_comp in (("no_competitor", False), ("with_competitor", True)):
            t0 = time.time()
            # 競合ありは need を大きめにして Tier3 が必ず走るよう計測
            pool = gemini_recipe_pool_(
                materials,
                client=client,
                need=6 if not use_comp else 12,
                use_competitor_hints=use_comp,
                force_refresh_official=bool(args.refresh_official_url),
            )
            sec = time.time() - t0
            timing = materials.get("_recipeTiming") or {}
            rows.append(
                {
                    "mode": label,
                    "sec": round(sec, 1),
                    "count": len(pool),
                    "timing": timing,
                    "dishes": [
                        {
                            "dishName": p["dishName"],
                            "tier": p.get("sourceTier"),
                            "url": p.get("sourceUrl"),
                        }
                        for p in pool
                    ],
                }
            )
            print(
                f"[recipe_timing] {label} sec={sec:.1f} count={len(pool)} "
                f"mat={timing.get('materials', 0):.1f} "
                f"google={timing.get('google', 0):.1f} "
                f"comp={timing.get('competitor', 0):.1f}"
            )
            for p in pool:
                print(f"  - {p.get('sourceTier')}: {p['dishName']} | {p.get('sourceUrl')}")
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = OUT_DIR / f"recipe_timing_{ts}.json"
        out_path.write_text(
            json.dumps(
                {
                    "ok": True,
                    "parentSku": parent.get("parentSku"),
                    "name": materials.get("name"),
                    "baseline_note": "前回全体360〜494秒のうち検索は約28秒。本計測はソース取得のみ。",
                    "modes": rows,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(json.dumps({"ok": True, "out": str(out_path)}, ensure_ascii=False))
        return 0

    results = []
    for parent in parents:
        LOG.info("run parent=%s children=%s", parent["parentSku"], parent["childCount"])
        results.append(
            run_one(
                parent,
                use_gemini=use_gemini,
                client=client,
                force_refresh_official=bool(args.refresh_official_url),
            )
        )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = Path(args.out) if args.out else OUT_DIR / f"propose_virt_{ts}.json"
    payload = {
        "ok": True,
        "useGemini": use_gemini,
        "ladderRewriteBudget": LADDER_REWRITE_MAX,
        "note": "整文プロンプトはGAS b5GeminiCatchOneTheme_揃え。ensure_mixはGAS同様（セールス置換禁止）。salesBurn別欄。",
        "results": [
            {
                "parent": r["parent"],
                "materials": r["materials"],
                "proposals": r["proposals"],
            }
            for r in results
        ],
    }
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    md_path = out_path.with_suffix(".md")
    md_path.write_text("\n\n".join(r["human"] for r in results), encoding="utf-8")

    print(md_path.read_text(encoding="utf-8"))
    print(json.dumps({"ok": True, "out": str(out_path), "md": str(md_path), "parents": len(results)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
