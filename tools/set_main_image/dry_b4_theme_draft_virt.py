# -*- coding: utf-8 -*-
"""B-⑤ テーマ案下書き P1 仮想テスト（Gemini・Driveなし）。

  cd tools/set_main_image
  python dry_b4_theme_draft_virt.py

正本: B45 §4.5.7 / b4_theme_draft_rules.py
"""
from __future__ import annotations

import sys
from typing import Callable, List, Tuple

from b4_theme_draft_rules import (
    MAX_THEMES,
    apply_remake,
    clip_theme_count,
    decide_open_mode,
    draft_file_stem,
    draft_matches_binding,
    duplicate_catch_indexes,
    is_draft_reusable,
    is_proposal_full,
    merge_by_theme_id,
    normalize_selected,
    reject_partial_overlay,
    remake_indexes_unchecked,
)


def _ok(name: str, cond: bool, detail: str = "") -> Tuple[str, bool, str]:
    return (name, cond, detail)


def test_stem_prefers_parent_sku() -> Tuple[str, bool, str]:
    s = draft_file_stem("ishi-4906283045515-oya5", "4906283045515", "チーズ")
    return _ok("U1 stem=親SKU優先", s == "ishi-4906283045515-oya5", s)


def test_stem_fallback_jan_name() -> Tuple[str, bool, str]:
    s = draft_file_stem("", "4906283045515", "チーズかつお 220g")
    return _ok("U1 stem=JAN+nameBase", s.startswith("4906283045515__"), s)


def test_thin_catch_only_not_full() -> Tuple[str, bool, str]:
    thin = {"themeId": 1, "catch": "短いフック"}
    return _ok("U5 薄いcatchのみはフルでない", not is_proposal_full(thin), str(thin))


def test_full_row_ok() -> Tuple[str, bool, str]:
    full = {
        "themeId": 1,
        "overview": "構図の概要文です。",
        "catch": "悩みに寄り添う一皿",
        "subs": ["補足A"],
        "needPackage": False,
    }
    return _ok("U5 丸ごと行はフル", is_proposal_full(full), "")


def test_reject_partial_overlay() -> Tuple[str, bool, str]:
    bad = reject_partial_overlay([{"themeId": 1, "catch": "のみ"}])
    good = not reject_partial_overlay(
        [
            {
                "themeId": 1,
                "overview": "概要",
                "catch": "キャッチ",
                "subs": [],
                "needPackage": False,
            }
        ]
    )
    return _ok("U5 部分overlay拒否", bad and good, f"bad={bad} good={good}")


def test_reuse_when_draft() -> Tuple[str, bool, str]:
    draft = {
        "schema": "b4_theme_draft_v1",
        "proposals": [
            {
                "themeId": 1,
                "overview": "概要1",
                "catch": "キャッチ1",
                "subs": [],
                "needPackage": False,
            }
        ],
    }
    mode = decide_open_mode(draft)
    return _ok("U1 下書きあればreuse", mode == "reuse" and is_draft_reusable(draft), mode)


def test_generate_when_empty() -> Tuple[str, bool, str]:
    mode = decide_open_mode(None)
    return _ok("U1 無ければgenerate", mode == "generate", mode)


def test_force_full() -> Tuple[str, bool, str]:
    draft = {
        "proposals": [
            {
                "themeId": 1,
                "overview": "概要",
                "catch": "キャッチ",
                "subs": [],
                "needPackage": False,
            }
        ]
    }
    mode = decide_open_mode(draft, force_full=True)
    return _ok("U1 force_fullはgenerate", mode == "generate", mode)


def test_remake_only_unchecked() -> Tuple[str, bool, str]:
    props = [
        {
            "themeId": 1,
            "overview": "A",
            "catch": "残す",
            "subs": [],
            "needPackage": False,
            "selected": True,
        },
        {
            "themeId": 2,
            "overview": "B",
            "catch": "作り直し前",
            "subs": [],
            "needPackage": False,
            "selected": False,
        },
        {
            "themeId": 3,
            "overview": "C",
            "catch": "残す2",
            "subs": [],
            "needPackage": False,
            "selected": True,
        },
    ]
    idxs = remake_indexes_unchecked(props)
    remade = apply_remake(
        props,
        {
            1: {
                "themeId": 2,
                "overview": "新概要",
                "catch": "新キャッチ",
                "subs": ["新補足"],
                "needPackage": True,
            }
        },
    )
    ok = (
        idxs == [1]
        and remade[0]["catch"] == "残す"
        and remade[1]["catch"] == "新キャッチ"
        and remade[1].get("selected") is True
        and remade[2]["catch"] == "残す2"
    )
    return _ok("U2/U4 外した案だけ差し替え＋レ点ON", ok, str(idxs))


def test_normalize_default_on() -> Tuple[str, bool, str]:
    rows = normalize_selected(
        [{"themeId": 1, "overview": "o", "catch": "c", "subs": [], "needPackage": False}]
    )
    return _ok("U4 初期selected欠落はON", rows[0].get("selected") is True, str(rows[0]))


def test_duplicate_catch_later() -> Tuple[str, bool, str]:
    props = [
        {"themeId": 1, "catch": "同じ文", "overview": "a", "subs": [], "needPackage": False},
        {"themeId": 2, "catch": "別", "overview": "b", "subs": [], "needPackage": False},
        {"themeId": 3, "catch": "同じ文", "overview": "c", "subs": [], "needPackage": False},
    ]
    dups = duplicate_catch_indexes(props)
    return _ok("U10 後案の同文catch", dups == [2], str(dups))


def test_merge_theme_id() -> Tuple[str, bool, str]:
    a = [
        {"themeId": 1, "catch": "旧1", "overview": "o", "subs": [], "needPackage": False},
        {"themeId": 2, "catch": "旧2", "overview": "o", "subs": [], "needPackage": False},
    ]
    b = [{"themeId": 2, "catch": "新2", "overview": "o2", "subs": [], "needPackage": True}]
    m = merge_by_theme_id(a, b)
    return _ok(
        "U11 themeIdが正（同ID上書き）",
        len(m) == 2 and m[1]["catch"] == "新2",
        str(m),
    )


def test_clip_not_child_count() -> Tuple[str, bool, str]:
    props = [
        {
            "themeId": i,
            "catch": f"c{i}",
            "overview": "o",
            "subs": [],
            "needPackage": False,
        }
        for i in range(1, 15)
    ]
    clipped = clip_theme_count(props, MAX_THEMES)
    # 子レ点3でも10まで残せる（連動しない）
    return _ok(
        "U9 テーマ上限10・子レ点非連動",
        len(clipped) == 10 and MAX_THEMES == 10,
        str(len(clipped)),
    )


def test_back_keeps_draft_semantics() -> Tuple[str, bool, str]:
    """戻る＝破棄しない。保存済みdraftがあれば再びreuse。"""
    draft = {
        "parentSku": "ishi-4906283045515-oya5",
        "jan": "4906283045515",
        "nameBase": "チーズかつお 220g 食べるおだし（ かつお ） 50ｇ",
        "proposals": [
            {
                "themeId": 1,
                "overview": "保持",
                "catch": "保持キャッチ",
                "subs": [],
                "needPackage": False,
            }
        ],
    }
    mode = decide_open_mode(
        draft,
        parent_sku="ishi-4906283045515-oya5",
        jan="4906283045515",
        name_base="チーズかつお 220g 食べるおだし（ かつお ） 50ｇ",
    )
    return _ok("U6 戻る後もdraftでreuse可", mode == "reuse", mode)


def test_reject_wrong_parent_sku() -> Tuple[str, bool, str]:
    draft = {
        "parentSku": "ishi-4906283045119-oya",
        "jan": "4906283045119",
        "nameBase": "食べるおだし（ かつお ） 50ｇ",
        "proposals": [
            {
                "themeId": 1,
                "overview": "おだし概要",
                "catch": "おだしの旨み",
                "subs": [],
                "needPackage": False,
            }
        ],
    }
    ok_match = draft_matches_binding(
        draft,
        parent_sku="ishi-4906283045515-oya5",
        jan="4906283045515",
        name_base="チーズかつお 220g 食べるおだし（ かつお ） 50ｇ",
    )
    mode = decide_open_mode(
        draft,
        parent_sku="ishi-4906283045515-oya5",
        jan="4906283045515",
        name_base="チーズかつお 220g 食べるおだし（ かつお ） 50ｇ",
    )
    return _ok(
        "親SKU不一致の下書きはreuseしない",
        (not ok_match) and mode == "generate",
        f"match={ok_match} mode={mode}",
    )


def test_reject_wrong_name_base() -> Tuple[str, bool, str]:
    draft = {
        "parentSku": "ishi-4906283045515-oya5",
        "jan": "4906283045515",
        "nameBase": "チーズかつお 220g",
        "proposals": [
            {
                "themeId": 1,
                "overview": "単品概要",
                "catch": "チーズかつお",
                "subs": [],
                "needPackage": False,
            }
        ],
    }
    mode = decide_open_mode(
        draft,
        parent_sku="ishi-4906283045515-oya5",
        jan="4906283045515",
        name_base="チーズかつお 220g 食べるおだし（ かつお ） 50ｇ",
    )
    return _ok("商品名ベース不一致はreuseしない", mode == "generate", mode)


def main() -> int:
    tests: List[Callable[[], Tuple[str, bool, str]]] = [
        test_stem_prefers_parent_sku,
        test_stem_fallback_jan_name,
        test_thin_catch_only_not_full,
        test_full_row_ok,
        test_reject_partial_overlay,
        test_reuse_when_draft,
        test_generate_when_empty,
        test_force_full,
        test_remake_only_unchecked,
        test_normalize_default_on,
        test_duplicate_catch_later,
        test_merge_theme_id,
        test_clip_not_child_count,
        test_back_keeps_draft_semantics,
        test_reject_wrong_parent_sku,
        test_reject_wrong_name_base,
    ]
    failed = 0
    for fn in tests:
        name, cond, detail = fn()
        mark = "PASS" if cond else "FAIL"
        if not cond:
            failed += 1
        print(f"[{mark}] {name}" + (f" | {detail}" if detail else ""))
    print(f"failed_tests={failed} total={len(tests)}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
