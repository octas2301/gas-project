# -*- coding: utf-8 -*-
"""B-⑤ テーマ案 つなぎ・軸・買う人 P2第2スライス 仮想（Gemini・Driveなし）。

  cd tools/set_main_image
  python dry_b4_theme_cohesion_virt.py

正本: B45 §4.5.9 / b4_theme_cohesion_rules.py
"""
from __future__ import annotations

import sys
from typing import Callable, List, Tuple

from b4_theme_cohesion_rules import (
    cohesion_ok,
    dish_scene_ok,
    duplicate_axis_pairs,
    find_face_handoff,
    gate_proposal_cohesion,
    gate_proposals_cohesion,
    normalize_axis,
)


def _ok(name: str, cond: bool, detail: str = "") -> Tuple[str, bool, str]:
    return (name, cond, detail)


def test_normalize_axis() -> Tuple[str, bool, str]:
    a = normalize_axis("夜の 自分ねぎらい")
    b = normalize_axis("夜の自分ねぎらい")
    return _ok("軸正規化で空白無視", a == b and a != "", a)


def test_t6_cohesion_pass() -> Tuple[str, bool, str]:
    pr = {
        "themeId": 1,
        "appealAxis": "夜の自分ねぎらい",
        "overview": "夜の食卓で自分ねぎらいの一杯。",
        "catch": "頑張った夜の自分ねぎらい",
        "subs": ["ほっと一息つく時間"],
    }
    return _ok("T6 2欄以上でつなぎ", cohesion_ok(pr), str(gate_proposal_cohesion(pr)))


def test_t6_cohesion_fail() -> Tuple[str, bool, str]:
    pr = {
        "themeId": 2,
        "appealAxis": "夜の自分ねぎらい",
        "overview": "焼津の港の風景。",
        "catch": "老舗の伝統製法",
        "subs": ["職人が大釜で炊く"],
    }
    return _ok("T6 つながらない", not cohesion_ok(pr), str(gate_proposal_cohesion(pr)))


def test_t6_no_axis() -> Tuple[str, bool, str]:
    errs = gate_proposal_cohesion({"themeId": 3, "catch": "x", "overview": "y", "subs": []})
    return _ok("T6 軸なし不合格", "T6:no_axis" in errs, str(errs))


def test_t6_dish_scene() -> Tuple[str, bool, str]:
    bad = {
        "themeId": 10,
        "appealAxis": "冷奴アレンジ",
        "dishName": "冷奴",
        "overview": "袋が並ぶ棚の写真。",
        "catch": "冷奴アレンジの新定番",
        "subs": ["冷奴に合う"],
    }
    # catchに冷奴あり → OK
    ok1 = dish_scene_ok(bad)
    bad2 = dict(bad)
    bad2["catch"] = "新定番の一品"
    bad2["subs"] = ["忙しい朝に"]
    bad2["overview"] = "袋が並ぶ棚。"
    ok2 = dish_scene_ok(bad2)
    return _ok("T6 食べ方は場面必須", ok1 and (not ok2), f"ok1={ok1} ok2={ok2}")


def test_t7_dup_axis() -> Tuple[str, bool, str]:
    props = [
        {"themeId": 1, "appealAxis": "夜の自分ねぎらい", "overview": "夜の自分ねぎらい", "catch": "夜の自分ねぎらい", "subs": ["ねぎらい"]},
        {"themeId": 2, "appealAxis": "夜の 自分ねぎらい", "overview": "夜の自分ねぎらい食事", "catch": "ねぎらいの夜", "subs": ["ほっと"]},
    ]
    dups = duplicate_axis_pairs(props)
    g = gate_proposals_cohesion(props)
    return _ok("T7 軸被り検出", len(dups) >= 1 and (not g["ok"]), str(dups))


def test_t7_unique_ok() -> Tuple[str, bool, str]:
    props = [
        {
            "themeId": 1,
            "appealAxis": "夜の自分ねぎらい",
            "overview": "夜の自分ねぎらいの食卓。",
            "catch": "夜の自分ねぎらい",
            "subs": ["ほっと一息"],
        },
        {
            "themeId": 2,
            "appealAxis": "手土産の安心",
            "overview": "手土産の安心が伝わる包み。",
            "catch": "手土産の安心を添えて",
            "subs": ["相手を選ばない味"],
        },
    ]
    g = gate_proposals_cohesion(props)
    return _ok("T7 軸が違えば通る", g["ok"], str(g))


def test_t13_face_handoff() -> Tuple[str, bool, str]:
    hit = find_face_handoff("手渡しで喜ばれる一品")
    errs = gate_proposal_cohesion(
        {
            "themeId": 13,
            "appealAxis": "手土産の安心",
            "overview": "手土産の安心を届ける。",
            "catch": "手渡しで喜ばれる一品",
            "subs": ["手土産の安心"],
        }
    )
    return _ok("T13 手渡し劇場を落とす", bool(hit) and any(e.startswith("T13:") for e in errs), str(errs))


TESTS: List[Callable[[], Tuple[str, bool, str]]] = [
    test_normalize_axis,
    test_t6_cohesion_pass,
    test_t6_cohesion_fail,
    test_t6_no_axis,
    test_t6_dish_scene,
    test_t7_dup_axis,
    test_t7_unique_ok,
    test_t13_face_handoff,
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
