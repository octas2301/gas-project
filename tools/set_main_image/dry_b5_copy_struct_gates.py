# -*- coding: utf-8 -*-
"""要件 §0.1 S1〜S14 仮想テスト（画像・GASなし）。

使い方:
  cd tools/set_main_image
  python dry_b5_copy_struct_gates.py

合格の意味:
  - expect_pass=True の fixture → 全ゲート PASS
  - expect_pass=False の fixture → ゲートが欠陥を検知（1つ以上 FAIL）

fixture の商品例は検証用。合格文・ゲートは構成／区別語の汎用ルール（特定ブランド専用にしない）。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List

from b5_copy_struct_gates import naive_slice_burn, run_all_gates, zenkaku_len

ROOT = Path(__file__).resolve().parent


def _base_good_proposal(**over: Any) -> Dict[str, Any]:
    """既定例は非魚ミックス（チーズ×梅）。魚の例は別fixture。"""
    p = {
        "themeId": "T3",
        "themeName": "使い方",
        "classId": "use_0",
        "sellRank": 1,
        "catch": "【お茶請けに】チーズも梅も選べる",
        "catchBurn": "【お茶請けに】チーズ・梅",
        "salesTalk": "その日の気分で選べる組み合わせ",
        "salesBurn": "その日の気分で",
        "personaShort": "お茶請け・おやつを探している人",
        "personaDetail": "午後や休憩。お茶請けに軽く食べられるものを探す。",
        "sceneHint": "袋を手前に。小皿にチーズと梅が並ぶ。",
        "dishKey": "cheese_ume_plate",
        "errandTokens": ["お茶請け"],
    }
    p.update(over)
    return p


def _good_case_common() -> Dict[str, Any]:
    return {
        "ui_can_open_detail": True,
        "ui_shows_scene_editable": True,
        "sanitize_allows_single_distinct_for_mix": False,
        "entrance_gate_applied": False,
        "invented_fact_in_copy": False,
        "image_text_fields": ["catchBurn", "salesBurn"],
        "text20_counts_sales": True,
        "is_food": True,
    }


def fixtures() -> List[Dict[str, Any]]:
    good = _base_good_proposal()
    long_catch = "【忙しい日に】チーズも梅もすぐ出せて食卓が整う午後の一品"
    assert zenkaku_len(long_catch) > 18

    return [
        {
            "id": "GOOD_mix_cheese_ume",
            "expect_pass": True,
            "note": "非魚ミックスの検証例",
            "is_mix": True,
            "mix_marks": ["チーズ", "梅"],
            **_good_case_common(),
            "persona_short": good["personaShort"],
            "polish_prompt": (
                "ペルソナ: お茶請け・おやつを探している人\n"
                "詳細: 午後や休憩。お茶請けに軽く食べられるものを探す。\n"
                "ペルソナの用事に沿い材料だけ言い換える。\n"
                '出力JSON: {"catch":"...","salesTalk":"..."}'
            ),
            "proposals": [
                good,
                _base_good_proposal(
                    themeId="T4",
                    classId="ben_0",
                    sellRank=2,
                    catch="【休憩に】チーズ・梅の組み合わせ",
                    catchBurn="【休憩に】チーズ梅",
                    salesTalk="迷わず一袋",
                    salesBurn="迷わず一袋",
                    personaShort="休憩に一品足したい人",
                    personaDetail="時間がない平日。休憩に手間を増やさず食卓を整えたい。",
                    sceneHint="両方の袋が見える。盛り付けは少量。",
                    dishKey="both_bags_rest",
                    errandTokens=["休憩"],
                ),
            ],
            "checked_keys": ["T3:use_0:1"],
            "remake_before": [
                good,
                _base_good_proposal(
                    themeId="T4",
                    classId="ben_0",
                    sellRank=2,
                    catch="【旧】保持される案",
                    catchBurn="【旧】保持文",
                    errandTokens=["旧"],
                    personaShort="旧の用事の人",
                    personaDetail="旧の用事で保持確認。",
                    dishKey="keep_a",
                ),
            ],
            "remake_after": [
                _base_good_proposal(catch="【新】レ点だけ更新", catchBurn="【新】レ点更新"),
                _base_good_proposal(
                    themeId="T4",
                    classId="ben_0",
                    sellRank=2,
                    catch="【旧】保持される案",
                    catchBurn="【旧】保持文",
                    errandTokens=["旧"],
                    personaShort="旧の用事の人",
                    personaDetail="旧の用事で保持確認。",
                    dishKey="keep_a",
                ),
            ],
        },
        {
            "id": "GOOD_mix_fish_sample",
            "expect_pass": True,
            "note": "魚ミックスも同じ門で通る検証例（専用ルールではない）",
            "is_mix": True,
            "mix_marks": ["ぶり", "かつお"],
            **_good_case_common(),
            "persona_short": "晩酌を少し華やかにしたい人",
            "polish_prompt": (
                "ペルソナ: 晩酌を少し華やかにしたい人\n"
                "詳細: 仕事帰り。晩酌のお酒のつまみをすぐ出したい。\n"
                "ペルソナの用事に沿い材料だけ言い換える。"
            ),
            "proposals": [
                _base_good_proposal(
                    catch="【晩酌に】ぶりとかつおが選べる",
                    catchBurn="【晩酌に】ぶり・かつお",
                    salesTalk="その日の気分で",
                    salesBurn="その日の気分で",
                    personaShort="晩酌を少し華やかにしたい人",
                    personaDetail="仕事帰り。晩酌のお酒のつまみをすぐ出したい。",
                    sceneHint="袋を手前に。皿にぶりとかつおが並ぶ。",
                    dishKey="buri_katsuo",
                    errandTokens=["晩酌"],
                )
            ],
            "checked_keys": ["T3:use_0:1"],
            "remake_before": [
                _base_good_proposal(
                    catch="【晩酌に】ぶりとかつおが選べる",
                    catchBurn="【晩酌に】ぶり・かつお",
                    errandTokens=["晩酌"],
                    personaShort="晩酌を少し華やかにしたい人",
                    personaDetail="仕事帰り。晩酌のお酒のつまみをすぐ出したい。",
                )
            ],
            "remake_after": [
                _base_good_proposal(
                    catch="【晩酌に】ぶりとかつおが選べる",
                    catchBurn="【晩酌に】ぶり・かつお",
                    errandTokens=["晩酌"],
                    personaShort="晩酌を少し華やかにしたい人",
                    personaDetail="仕事帰り。晩酌のお酒のつまみをすぐ出したい。",
                )
            ],
        },
        {
            "id": "BAD_S1_single_part",
            "expect_pass": False,
            "apply_gates": ["S1"],
            "is_mix": True,
            "mix_marks": ["チーズ", "梅"],
            "proposals": [
                _base_good_proposal(
                    catch="【お茶請けに】チーズだけ",
                    catchBurn="【お茶請けに】チーズ",
                    salesTalk="香ばしい",
                )
            ],
        },
        {
            "id": "BAD_S2_no_sales",
            "expect_pass": False,
            "apply_gates": ["S2"],
            "is_mix": False,
            "proposals": [_base_good_proposal(salesTalk="")],
        },
        {
            "id": "BAD_S3_no_detail_ui",
            "expect_pass": False,
            "apply_gates": ["S3"],
            "ui_can_open_detail": False,
            "proposals": [_base_good_proposal()],
        },
        {
            "id": "BAD_S4_tail",
            "expect_pass": False,
            "apply_gates": ["S4"],
            "proposals": [_base_good_proposal(catchBurn="【お茶請けに】チズて")],
        },
        {
            "id": "BAD_S5_old_prompt",
            "expect_pass": False,
            "apply_gates": ["S5"],
            "persona_short": "お茶請け・おやつを探している人",
            "polish_prompt": "材料に無い相手（ペルソナ属性）を作って用事にしない\n出力JSON",
            "proposals": [_base_good_proposal()],
        },
        {
            "id": "BAD_S6_remake_all",
            "expect_pass": False,
            "apply_gates": ["S6"],
            "checked_keys": ["T3:use_0:1"],
            "remake_before": [
                _base_good_proposal(),
                _base_good_proposal(
                    themeId="T4", classId="ben_0", sellRank=2, catch="【旧】保持すべき"
                ),
            ],
            "remake_after": [
                _base_good_proposal(catch="【新】T3"),
                _base_good_proposal(
                    themeId="T4", classId="ben_0", sellRank=2, catch="【新】T4まで変わった"
                ),
            ],
            "proposals": [_base_good_proposal()],
        },
        {
            "id": "BAD_S7_no_scene_ui",
            "expect_pass": False,
            "apply_gates": ["S7"],
            "ui_shows_scene_editable": False,
            "proposals": [_base_good_proposal()],
        },
        {
            "id": "BAD_S8_naive_slice",
            "expect_pass": False,
            "apply_gates": ["S8"],
            "proposals": [
                _base_good_proposal(
                    catch=long_catch,
                    catchBurn=naive_slice_burn(long_catch, 18.0),
                )
            ],
        },
        {
            "id": "BAD_S9_mix_distinct_policy",
            "expect_pass": False,
            "apply_gates": ["S9"],
            "is_mix": True,
            "sanitize_allows_single_distinct_for_mix": True,
            "proposals": [_base_good_proposal()],
        },
        {
            "id": "BAD_S10_errand_split",
            "expect_pass": False,
            "apply_gates": ["S10"],
            "proposals": [
                _base_good_proposal(
                    catch="【お茶請けに】チーズも梅も選べる",
                    catchBurn="【お茶請けに】チーズ・梅",
                    personaShort="晩酌を華やかにしたい人",
                    personaDetail="仕事帰り。お酒のつまみをすぐ出したい。",
                    errandTokens=["お茶請け"],
                    sceneConflictTokens=["晩酌"],
                    sceneHint="袋を手前に。晩酌のつまみが並ぶ。",
                )
            ],
        },
        {
            "id": "BAD_S11_taste_first",
            "expect_pass": False,
            "apply_gates": ["S11"],
            "proposals": [
                _base_good_proposal(
                    themeId="T3",
                    catch="【おいしい】チーズと梅",
                    catchBurn="【おいしい】チーズ梅",
                )
            ],
        },
        {
            "id": "BAD_S11_heading_only",
            "expect_pass": False,
            "apply_gates": ["S11"],
            "proposals": [
                _base_good_proposal(
                    themeId="T4",
                    catch="【お茶請けに】",
                    catchBurn="【お茶請けに】短",
                )
            ],
        },
        {
            "id": "BAD_S12_scene_sales",
            "expect_pass": False,
            "apply_gates": ["S12"],
            "is_food": True,
            "proposals": [
                _base_good_proposal(sceneHint="今すぐクリック。お得な最強おやつ。")
            ],
        },
        {
            "id": "BAD_S12_dish_repeat",
            "expect_pass": False,
            "apply_gates": ["S12"],
            "is_food": True,
            "proposals": [
                _base_good_proposal(dishKey="same_dish", sceneHint="袋を手前に。皿に盛る。"),
                _base_good_proposal(
                    themeId="T4",
                    classId="ben_0",
                    sellRank=2,
                    dishKey="same_dish",
                    sceneHint="袋が見える。同じ盛り付け。",
                    errandTokens=["休憩"],
                    catch="【休憩に】チーズ・梅",
                    catchBurn="【休憩に】チーズ梅",
                    personaShort="休憩に一品足したい人",
                    personaDetail="休憩に手間を増やしたくない。",
                ),
            ],
        },
        {
            "id": "BAD_S13_entrance_gate",
            "expect_pass": False,
            "apply_gates": ["S13"],
            "entrance_gate_applied": True,
            "proposals": [_base_good_proposal()],
        },
        {
            "id": "BAD_S14_burn_only",
            "expect_pass": False,
            "apply_gates": ["S14"],
            "image_text_mode": "burn_only",
            "image_text_fields": ["catchBurn"],
            "text20_counts_sales": False,
            "proposals": [_base_good_proposal()],
        },
        {
            "id": "BAD_S14_long_sales_no_burn",
            "expect_pass": False,
            "apply_gates": ["S14"],
            "image_text_fields": ["catchBurn", "salesTalk"],
            "text20_counts_sales": True,
            "proposals": [
                _base_good_proposal(
                    salesTalk="その日の気分でチーズも梅も迷わず選べる長い文",
                    salesBurn="",
                )
            ],
        },
        {
            "id": "GOOD_S14_long_sales_with_burn",
            "expect_pass": True,
            "apply_gates": ["S14"],
            "image_text_fields": ["catchBurn", "salesBurn"],
            "text20_counts_sales": True,
            "proposals": [
                _base_good_proposal(
                    salesTalk="その日の気分でチーズも梅も迷わず選べる長い文",
                    salesBurn="その日の気分で",
                )
            ],
        },
        {
            "id": "SNAP_current_gas_like",
            "expect_pass": False,
            "note": "現行実装に近い欠陥束（特定商品専用ではない）。",
            "is_mix": True,
            "mix_marks": ["チーズ", "梅"],
            "ui_can_open_detail": False,
            "ui_shows_scene_editable": False,
            "sanitize_allows_single_distinct_for_mix": True,
            "entrance_gate_applied": True,
            "image_text_mode": "burn_only",
            "persona_short": "お茶請け・おやつを探している人",
            "polish_prompt": "材料に無い相手（ペルソナ属性）を作って用事にしない",
            "proposals": [
                {
                    "themeId": "N1",
                    "classId": "name",
                    "sellRank": 1,
                    "catch": "【セット】チーズ",
                    "catchBurn": naive_slice_burn(long_catch, 18.0),
                    "salesTalk": "",
                    "personaShort": "お茶請け・おやつを探している人",
                    "personaDetail": "",
                    "sceneHint": "袋が大きく読める。",
                    "errandTokens": ["お茶請け"],
                }
            ],
            "checked_keys": ["N1:name:1"],
            "remake_before": [
                {
                    "themeId": "N1",
                    "classId": "name",
                    "sellRank": 1,
                    "catch": "A",
                    "salesTalk": "x",
                    "catchBurn": "【A】短い",
                },
                {
                    "themeId": "T3",
                    "classId": "u",
                    "sellRank": 2,
                    "catch": "B",
                    "salesTalk": "y",
                    "catchBurn": "【B】短い",
                },
            ],
            "remake_after": [
                {
                    "themeId": "N1",
                    "classId": "name",
                    "sellRank": 1,
                    "catch": "A2",
                    "salesTalk": "x2",
                    "catchBurn": "【A2】",
                },
                {
                    "themeId": "T3",
                    "classId": "u",
                    "sellRank": 2,
                    "catch": "B2",
                    "salesTalk": "y2",
                    "catchBurn": "【B2】",
                },
            ],
        },
    ]


def main() -> int:
    rows = []
    failed_tests = 0
    for case in fixtures():
        r = run_all_gates(case)
        gate_pass = bool(r["pass"])
        expect = bool(case.get("expect_pass", True))
        test_ok = gate_pass if expect else (not gate_pass)
        if not test_ok:
            failed_tests += 1
        rows.append(
            {
                "id": case["id"],
                "expect_pass": expect,
                "gates_pass": gate_pass,
                "test_ok": test_ok,
                "fails": r["fails"],
            }
        )
        mark = "OK" if test_ok else "NG"
        print(
            f"[{mark}] {case['id']} expect_pass={expect} gates_pass={gate_pass} fails={r['fails']}"
        )

    out = ROOT / "_out_b5_copy_struct_gates.json"
    out.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"wrote {out}")
    print(f"summary failed_tests={failed_tests}/{len(rows)}")
    return 1 if failed_tests else 0


if __name__ == "__main__":
    sys.exit(main())
