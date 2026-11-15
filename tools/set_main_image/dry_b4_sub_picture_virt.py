# -*- coding: utf-8 -*-
"""B-⑤ サブ絵合否 P3第1 仮想（画像生成・Geminiなし）。

  cd tools/set_main_image
  python dry_b4_sub_picture_virt.py

正本: B45 §4.5.10 / b4_sub_picture_rules.py
"""
from __future__ import annotations

import sys
from typing import Callable, List, Tuple

from b4_sub_picture_rules import (
    build_sub_food_front,
    picture_gate_enabled,
    prepare_overview_for_prompt,
    prompt_wrap_contract_errors,
    prompt_wrap_look_ok,
    scrub_scene_for_sub,
    strip_invent_wrap_ban,
    wrap_gate_failed,
)


def _ok(name: str, cond: bool, detail: str = "") -> Tuple[str, bool, str]:
    return (name, cond, detail)


def test_gate_default_on() -> Tuple[str, bool, str]:
    import os

    prev = os.environ.pop("B4_SUB_PICTURE_GATE", None)
    try:
        on = picture_gate_enabled()
    finally:
        if prev is not None:
            os.environ["B4_SUB_PICTURE_GATE"] = prev
    return _ok("切戻し既定ON", on, f"on={on}")


def test_gate_off() -> Tuple[str, bool, str]:
    import os

    prev = os.environ.get("B4_SUB_PICTURE_GATE")
    os.environ["B4_SUB_PICTURE_GATE"] = "0"
    try:
        off = not picture_gate_enabled()
    finally:
        if prev is None:
            os.environ.pop("B4_SUB_PICTURE_GATE", None)
        else:
            os.environ["B4_SUB_PICTURE_GATE"] = prev
    return _ok("切戻し B4_SUB_PICTURE_GATE=0", off, "")


def test_scrub_old_mix() -> Tuple[str, bool, str]:
    old = "小皿に少し混ぜて出せ。山にしない。"
    scrubbed = scrub_scene_for_sub(old)
    return _ok("G7 scrub 混ぜ文除去", "少し混ぜ" not in scrubbed, scrubbed[:80])


def test_strip_ban_when_wrap() -> Tuple[str, bool, str]:
    t = "個品を置け。個包装を新たに描くな。お茶。"
    a = strip_invent_wrap_ban(t, has_wrap=True)
    b = strip_invent_wrap_ban(t, has_wrap=False)
    return _ok(
        "WRAPありなら描くな除去",
        ("個包装を新たに描くな" not in a) and ("個包装を新たに描くな" in b),
        f"a={a!r}",
    )


def test_prepare_overview() -> Tuple[str, bool, str]:
    ov = "小皿に少し混ぜて出せ。個包装を新たに描くな。"
    out = prepare_overview_for_prompt(ov, has_wrap=True)
    ok = ("少し混ぜ" not in out) and ("個包装を新たに描くな" not in out)
    return _ok("概要 scrub+ban", ok, out)


def test_wrap_front_contract() -> Tuple[str, bool, str]:
    front = build_sub_food_front(
        has_identity=True,
        has_wrap=True,
        food_mode="identity",
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    errs = prompt_wrap_contract_errors(front, has_wrap=True)
    product_leak = "チーズかつお" in front or "まぐろチーズ" in front
    return _ok(
        "G7 WRAP指示契約",
        not errs and not product_leak,
        f"errs={errs} leak={product_leak}",
    )


def test_no_wrap_allows_ban() -> Tuple[str, bool, str]:
    front = build_sub_food_front(has_identity=True, has_wrap=False)
    ok = "個包装" in front and ("描くな" in front or "創作" in front)
    return _ok("WRAP無しは創作禁止", ok, front[200:320] if len(front) > 200 else front)


def test_wrap_look_injected() -> Tuple[str, bool, str]:
    from b5_food_look_read import wrap_look_block_ja

    fake = [
        {
            "label": "構成A・個包装",
            "lines": "密着: 中身に密着／空気: 見えない／材質: 硬いフィルム",
            "fields": {"密着": "中身に密着", "空気": "見えない", "材質": "硬いフィルム"},
        }
    ]
    block = wrap_look_block_ja(fake)
    front = build_sub_food_front(
        has_identity=True,
        has_wrap=True,
        wrap_look_block=block,
        mix_parts=["商品A", "商品B"],
    )
    ok = prompt_wrap_look_ok(front, has_wrap=True, look_snippet="硬いフィルム")
    return _ok("G6 look注入", ok and "チーズ" not in front, f"ok={ok}")


def test_wrap_gate_front_missing() -> Tuple[str, bool, str]:
    # skipped は落とさない
    skip = wrap_gate_failed({"skipped": True, "ok": False})
    # ok false → fail
    bad = wrap_gate_failed({"skipped": False, "ok": False, "frontOk": False})
    good = wrap_gate_failed({"skipped": False, "ok": True, "frontOk": True})
    return _ok(
        "G7 frontOk合否",
        (not skip) and bad and (not good),
        f"skip={skip} bad={bad} good={good}",
    )


def test_bag_needs_two() -> Tuple[str, bool, str]:
    from b4_sub_picture_rules import bag_prompt_and_verdict

    one = bag_prompt_and_verdict([])
    return _ok("G1 袋不足は空", one.get("block") == "" and one.get("rows") == [], str(one))


def test_dry_sub_prompt_wires_wrap() -> Tuple[str, bool, str]:
    from dry_sub_overview_copy_test import _build_image_prompt

    prompt = _build_image_prompt(
        product_name="材料テスト",
        theme_id=3,
        theme_name="使い方",
        overview="小皿に少し混ぜて出せ。個包装を新たに描くな。",
        catch="短いキャッチ",
        subs=["補足"],
        n_comp=0,
        has_food_unit_canon=True,
        has_food_unit_wrap=True,
        wrap_look_block="【個包装見本の見え方・機械読み・正】\n硬いフィルム",
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    errs = prompt_wrap_contract_errors(prompt, has_wrap=True)
    scrub_ok = "少し混ぜ" not in prompt and "個包装を新たに描くな" not in prompt
    look_ok = "硬いフィルム" in prompt
    return _ok(
        "dry_sub 配線: WRAP契約+scrub",
        not errs and scrub_ok and look_ok,
        f"errs={errs} scrub={scrub_ok} look={look_ok}",
    )


def main() -> int:
    tests: List[Callable[[], Tuple[str, bool, str]]] = [
        test_gate_default_on,
        test_gate_off,
        test_scrub_old_mix,
        test_strip_ban_when_wrap,
        test_prepare_overview,
        test_wrap_front_contract,
        test_no_wrap_allows_ban,
        test_wrap_look_injected,
        test_wrap_gate_front_missing,
        test_bag_needs_two,
        test_dry_sub_prompt_wires_wrap,
    ]
    failed = 0
    print("=== dry_b4_sub_picture_virt ===")
    for fn in tests:
        name, cond, detail = fn()
        mark = "PASS" if cond else "FAIL"
        if not cond:
            failed += 1
        print(f"[{mark}] {name}")
        if detail:
            print(f"       {detail}")
    print(f"total={len(tests)} failed={failed}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
