# -*- coding: utf-8 -*-
"""B-⑤ テーマ案 文・門・ミックス P2第1スライス 仮想（Gemini・Driveなし）。

  cd tools/set_main_image
  python dry_b4_theme_text_virt.py

正本: B45 §4.5.8 / b4_theme_text_rules.py
"""
from __future__ import annotations

import sys
from typing import Callable, List, Tuple

from b4_theme_text_rules import (
    AUTO_RETRY_MAX,
    catch_equals_any_sub,
    find_banned_claims,
    find_procedure_words,
    gate_proposal,
    gate_proposals,
    has_bracket_markers,
    is_mix_one_sided,
    next_retry_indexes,
    reused_chunks_across_themes,
    should_auto_retry,
)


def _ok(name: str, cond: bool, detail: str = "") -> Tuple[str, bool, str]:
    return (name, cond, detail)


def test_t2_bans_claim() -> Tuple[str, bool, str]:
    hit = find_banned_claims("毎日の健康を支える最強のだし")
    return _ok("T2 効能・最強を検出", "健康" in hit and "最強" in hit, str(hit))


def test_t2_clean_ok() -> Tuple[str, bool, str]:
    hit = find_banned_claims("だしをとらずに一品完成")
    return _ok("T2 普通の訴求は通る", hit == [], str(hit))


def test_t5_bracket() -> Tuple[str, bool, str]:
    return _ok("T5 【】検出", has_bracket_markers("【時短】すぐ食べ"), True)


def test_t5_catch_eq_sub() -> Tuple[str, bool, str]:
    pr = {"catch": "同じ文です", "subs": ["同じ文です"]}
    return _ok("T5 catch=subs", catch_equals_any_sub(pr), "")


def test_t15_procedure() -> Tuple[str, bool, str]:
    hit = find_procedure_words("開けてすぐのせて食べる")
    return _ok("T15 手順語", "開けて" in hit and "のせて" in hit, str(hit))


def test_t4_one_sided() -> Tuple[str, bool, str]:
    marks = ["チーズ", "まぐろ"]
    bad = is_mix_one_sided("濃厚チーズの旨み", marks)
    good_both = not is_mix_one_sided("チーズとまぐろの食べ比べ", marks)
    good_none = not is_mix_one_sided("だしをとらず一品", marks)
    good_alias = not is_mix_one_sided("チーズと鮪の食べ比べ", marks)  # まぐろ＝鮪
    return _ok(
        "T4 片寄りのみ不合格",
        bad and good_both and good_none and good_alias,
        f"bad={bad} both={good_both} none={good_none} alias={good_alias}",
    )


def test_t4_alias_and_overview_rescue() -> Tuple[str, bool, str]:
    """キャッチがチーズ寄りでも概要に鰹があれば通る（表記ゆれ＋全体判定）。"""
    pr = {
        "themeId": 3,
        "overview": "かつおの断面にチーズが絡む接写。",
        "catch": "しっとり食感とチーズが絡み合う",
        "subs": ["焼津の老舗が引き出す極上の旨み"],
    }
    marks = ["チーズ", "かつお"]
    errs = gate_proposal(pr, mix_marks=marks)
    # 鰹表記でも通る
    pr2 = dict(pr)
    pr2["overview"] = "鰹の断面にチーズが絡む接写。"
    errs2 = gate_proposal(pr2, mix_marks=marks)
    ok = ("T4:one_sided" not in errs) and ("T4:one_sided" not in errs2)
    return _ok("T4 概要＋表記ゆれで救済", ok, f"errs={errs} errs2={errs2}")


def test_t3_reuse() -> Tuple[str, bool, str]:
    props = [
        {"themeId": 1, "catch": "だしをとらず一品", "subs": ["忙しい朝に"]},
        {"themeId": 2, "catch": "だしをとらず一品", "subs": ["夜にも"]},
    ]
    pairs = reused_chunks_across_themes(props)
    return _ok("T3 同文チャンク検出", len(pairs) >= 1, str(pairs))


def test_gate_mix_fail() -> Tuple[str, bool, str]:
    props = [
        {
            "themeId": 3,
            "overview": "袋が並ぶ構図。",
            "catch": "濃厚チーズだけ",
            "subs": ["お土産にも"],
            "needPackage": False,
        },
        {
            "themeId": 4,
            "overview": "別シーン。",
            "catch": "まぐろの旨み",
            "subs": ["食卓に"],
            "needPackage": False,
        },
    ]
    g = gate_proposals(props, mix_marks=["チーズ", "まぐろ"])
    # 各案が片方だけ → T4
    return _ok("T4 門で片寄り落とし", (not g["ok"]) and 0 in g["failIndexes"], str(g["failIndexes"]))


def test_gate_clean_mix() -> Tuple[str, bool, str]:
    props = [
        {
            "themeId": 1,
            "overview": "食卓のシーン。",
            "catch": "チーズもまぐろも楽しむ",
            "subs": ["食べ比べセット"],
            "needPackage": False,
        },
        {
            "themeId": 2,
            "overview": "お弁当のシーン。",
            "catch": "だしをとらず一品",
            "subs": ["忙しい朝に"],
            "needPackage": False,
        },
    ]
    g = gate_proposals(props, mix_marks=["チーズ", "まぐろ"])
    return _ok("両方・非片寄りは通る", g["ok"], str(g))


def test_t8_retry_cap() -> Tuple[str, bool, str]:
    g = {"failIndexes": [1, 2]}
    a0 = next_retry_indexes(g, attempt=0)
    a2 = next_retry_indexes(g, attempt=2)
    a3 = next_retry_indexes(g, attempt=3)
    ok = (
        a0 == [1, 2]
        and a2 == [1, 2]
        and a3 == []
        and should_auto_retry(0)
        and not should_auto_retry(AUTO_RETRY_MAX)
    )
    return _ok("T8 自動リトライ上限3", ok, f"a0={a0} a2={a2} a3={a3}")


def test_single_gate_t2() -> Tuple[str, bool, str]:
    errs = gate_proposal(
        {"catch": "免疫を支える一品", "subs": [], "overview": "x"},
        mix_marks=[],
    )
    return _ok("単案T2", any(e.startswith("T2:") for e in errs), str(errs))


TESTS: List[Callable[[], Tuple[str, bool, str]]] = [
    test_t2_bans_claim,
    test_t2_clean_ok,
    test_t5_bracket,
    test_t5_catch_eq_sub,
    test_t15_procedure,
    test_t4_one_sided,
    test_t4_alias_and_overview_rescue,
    test_t3_reuse,
    test_gate_mix_fail,
    test_gate_clean_mix,
    test_t8_retry_cap,
    test_single_gate_t2,
]


def main() -> int:
    failed = 0
    for fn in TESTS:
        name, cond, detail = fn()
        mark = "PASS" if cond else "FAIL"
        if not cond:
            failed += 1
        print(f"[{mark}] {name}" + (f" | {detail}" if detail else ""))
    print(f"\n{len(TESTS) - failed}/{len(TESTS)} PASS")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
