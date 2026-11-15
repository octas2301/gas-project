# -*- coding: utf-8 -*-
"""P3-W 個包装素材レーン 仮想（画像生成なし）。

  cd tools/set_main_image
  python dry_b4_wrap_asset_virt.py

正本方針: 見本ゲートで WRAP 確定 → 手持ち＋単体 → 手持ち＝金枠／単体＝FOOD_UNIT_WRAP（compose）。
寄せ集め・低解像は compose に使わない。種は FOOD_UNIT_WRAP_SEED。
P1〜P3（案UI・文門・G6/G7）は触らない。本仮想は分岐と契約のみ。
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Callable, List, Tuple

from b4_wrap_asset_rules import (
    THEME_ID_WRAP_APPEAL,
    apply_selection,
    build_hand_hold_prompt_ja,
    candidate_manifest,
    hand_hold_bg_contract_errors,
    hand_hold_scale_contract_errors,
    max_themes_allowed,
    merge_wrap_appeal_theme,
    proposals_have_wrap_appeal,
    resolve_wrap_asset_for_consumers,
    should_branch_after_wrap_confirm,
    sub_wrap_appeal_via_theme_pipeline,
    wrap_appeal_theme_enabled,
    wrap_asset_lane_enabled,
)


def _ok(name: str, cond: bool, detail: str = "") -> Tuple[str, bool, str]:
    return (name, cond, detail)


def test_lane_default_on() -> Tuple[str, bool, str]:
    import os

    prev = os.environ.pop("B4_WRAP_ASSET_LANE", None)
    try:
        on = wrap_asset_lane_enabled()
    finally:
        if prev is not None:
            os.environ["B4_WRAP_ASSET_LANE"] = prev
    return _ok("レーン既定ON", on, f"on={on}")


def test_lane_off() -> Tuple[str, bool, str]:
    import os

    prev = os.environ.get("B4_WRAP_ASSET_LANE")
    os.environ["B4_WRAP_ASSET_LANE"] = "0"
    try:
        off = not wrap_asset_lane_enabled()
        no_branch = not should_branch_after_wrap_confirm(has_wrap=True)
    finally:
        if prev is None:
            os.environ.pop("B4_WRAP_ASSET_LANE", None)
        else:
            os.environ["B4_WRAP_ASSET_LANE"] = prev
    return _ok("切戻し B4_WRAP_ASSET_LANE=0", off and no_branch, "")


def test_branch_only_with_wrap() -> Tuple[str, bool, str]:
    a = should_branch_after_wrap_confirm(has_wrap=True)
    b = should_branch_after_wrap_confirm(has_wrap=False)
    return _ok("WRAPありだけ分岐", a and (not b), f"a={a} b={b}")


def test_branch_missing_file() -> Tuple[str, bool, str]:
    missing = Path("__no_such_wrap__.jpg")
    ok = not should_branch_after_wrap_confirm(has_wrap=True, wrap_path=missing)
    return _ok("WRAPパス無しは分岐しない", ok, "")


def test_prompt_bg_and_scale() -> Tuple[str, bool, str]:
    p = build_hand_hold_prompt_ja(circular=True, include_plant=True)
    bg = hand_hold_bg_contract_errors(p)
    sc = hand_hold_scale_contract_errors(p)
    return _ok(
        "手持ち指示に白系・植物・実寸",
        (not bg) and (not sc),
        f"bg={bg} scale={sc} prompt={p[:60]}",
    )


def test_prompt_rejects_gray_only() -> Tuple[str, bool, str]:
    bad = "背景はグレー一色。親指サイズ。"
    errs = hand_hold_bg_contract_errors(bad)
    return _ok("グレー一色を検出", "bg:white_required" in errs or "bg:flat_gray_ban_self" in errs, str(errs))


def test_manifest_select_and_resolve() -> Tuple[str, bool, str]:
    m = candidate_manifest(
        jan="4906283045614",
        wrap_path="FOOD_UNIT_WRAP.jpg",
        candidates=["a.jpg", "b.jpg", "c.jpg"],
    )
    m2 = apply_selection(m, "b.jpg")
    routes = resolve_wrap_asset_for_consumers(m2)
    # composeWrap 無し時は場面も手持ちへフォールバック
    ok = (
        m2.get("selected") == "b.jpg"
        and routes["mainBlGold"] == "b.jpg"
        and routes["subWrapAppeal"] == "b.jpg"
        and routes["sceneWrapRef"] == "b.jpg"
    )
    m2["composeWrap"] = "wrap_single_01.jpg"
    routes2 = resolve_wrap_asset_for_consumers(m2)
    ok2 = routes2["mainBlGold"] == "b.jpg" and routes2["sceneWrapRef"] == "wrap_single_01.jpg"
    return _ok("採用→金枠=手持ち／場面=単体(あれば)", ok and ok2, str(routes2))


def test_select_must_be_candidate() -> Tuple[str, bool, str]:
    m = candidate_manifest(jan="x", wrap_path="w", candidates=["a.jpg"])
    try:
        apply_selection(m, "z.jpg")
        ok = False
        detail = "no raise"
    except ValueError:
        ok = True
        detail = "ValueError"
    return _ok("候補外採用は拒否", ok, detail)


def test_sub_via_theme_pipeline() -> Tuple[str, bool, str]:
    return _ok(
        "②サブ個包装訴求はテーマ案パイプラインへ",
        sub_wrap_appeal_via_theme_pipeline(),
        "専用寂し固定レイアウトにしない",
    )


def test_merge_adds_theme_21() -> Tuple[str, bool, str]:
    base = [{"themeId": i, "catch": f"c{i}", "overview": "o"} for i in range(1, 11)]
    out = merge_wrap_appeal_theme(base, has_wrap=True)
    ok = (
        len(out) == 11
        and proposals_have_wrap_appeal(out)
        and int(out[-1]["themeId"]) == THEME_ID_WRAP_APPEAL
        and out[-1].get("needPackage") is False
        and max_themes_allowed(has_wrap_appeal=True) == 11
    )
    return _ok("WRAPあり→+1枠(themeId=21)合計11", ok, f"n={len(out)}")


def test_merge_no_wrap() -> Tuple[str, bool, str]:
    base = [{"themeId": 1, "catch": "a", "overview": "o"}]
    out = merge_wrap_appeal_theme(base, has_wrap=False)
    return _ok("WRAPなし→追加しない", len(out) == 1 and not proposals_have_wrap_appeal(out), f"n={len(out)}")


def test_merge_idempotent() -> Tuple[str, bool, str]:
    base = [{"themeId": 1, "catch": "a", "overview": "o"}]
    once = merge_wrap_appeal_theme(base, has_wrap=True)
    twice = merge_wrap_appeal_theme(once, has_wrap=True)
    n21 = sum(1 for p in twice if int(p.get("themeId") or 0) == THEME_ID_WRAP_APPEAL)
    return _ok("二重追加しない", len(twice) == 2 and n21 == 1, f"n={len(twice)} n21={n21}")


def test_appeal_theme_off() -> Tuple[str, bool, str]:
    import os

    prev = os.environ.get("B4_WRAP_APPEAL_THEME")
    os.environ["B4_WRAP_APPEAL_THEME"] = "0"
    try:
        off = not wrap_appeal_theme_enabled()
        out = merge_wrap_appeal_theme(
            [{"themeId": 1, "catch": "a", "overview": "o"}], has_wrap=True
        )
        ok = off and not proposals_have_wrap_appeal(out)
    finally:
        if prev is None:
            os.environ.pop("B4_WRAP_APPEAL_THEME", None)
        else:
            os.environ["B4_WRAP_APPEAL_THEME"] = prev
    return _ok("切戻し B4_WRAP_APPEAL_THEME=0", ok, "")


def test_has_wrap_lookup_order() -> Tuple[str, bool, str]:
    from b5_food_canon_status import resolve_has_wrap_from_status_doc

    doc = {
        "byParentSku": {
            "ishi-oya5": {"hasWrap": True, "jan": "5515", "nameBase": "かつお"},
        },
        "byNameBase": {
            "かつお": {"hasWrap": False, "jan": "5515"},
            "まぐろ": {"hasWrap": True, "jan": "5614"},
        },
        "byJan": {
            "5515": {"hasWrap": False},
            "5614": {"hasWrap": False},
        },
    }
    by_sku = resolve_has_wrap_from_status_doc(
        doc, parent_sku="ishi-oya5", name_base="かつお", jan="5515"
    )
    by_nb = resolve_has_wrap_from_status_doc(
        doc, parent_sku="", name_base="まぐろ", jan="5614"
    )
    by_jan_only = resolve_has_wrap_from_status_doc(
        doc, parent_sku="", name_base="", jan="5614"
    )
    # 親SKUが優先され、商品名ベースの false に負けない
    # まぐろは親SKU無しでも商品名ベースで true（JANの false に負けない）
    # JANだけは false
    ok = by_sku and by_nb and (not by_jan_only)
    return _ok(
        "hasWrap判定は親SKU→商品名ベース→JAN",
        ok,
        f"sku={by_sku} nb={by_nb} jan={by_jan_only}",
    )


def test_compose_wrap_requires_single_when_present() -> Tuple[str, bool, str]:
    import tempfile
    from pathlib import Path

    from b4_wrap_compose_ref import resolve_compose_wrap_choice
    from PIL import Image

    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        refs = root / "01_refs"
        refs.mkdir(parents=True)
        # 販促っぽい貼付（大きいが単体ではない）
        Image.new("RGB", (900, 900), (200, 180, 160)).save(refs / "MANUAL_04.jpg", quality=85)
        wh = root / "wrap_hand_hold"
        wh.mkdir(parents=True)
        Image.new("RGB", (1024, 1024), (240, 240, 240)).save(wh / "wrap_single_01.jpg", quality=85)
        bad = resolve_compose_wrap_choice(
            root,
            wrap_name="MANUAL_04.jpg",
            wrap_single_name="",
            wrap_has=True,
            singles_available=["wrap_single_01.jpg"],
        )
        good = resolve_compose_wrap_choice(
            root,
            wrap_name="MANUAL_04.jpg",
            wrap_single_name="wrap_single_01.jpg",
            wrap_has=True,
            singles_available=["wrap_single_01.jpg"],
        )
        ok = (not bad.get("ok")) and bool(good.get("ok")) and good.get("composeName") == "wrap_single_01.jpg"
        return _ok("単体ありなら単体必須・compose=単体", ok, f"bad={bad.get('ok')} good={good}")


def test_compose_wrap_rejects_low_res() -> Tuple[str, bool, str]:
    import tempfile
    from pathlib import Path

    from b4_wrap_compose_ref import resolve_compose_wrap_choice
    from PIL import Image

    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        refs = root / "01_refs"
        refs.mkdir(parents=True)
        Image.new("RGB", (238, 217), (100, 80, 60)).save(refs / "tiny.jpg", quality=85)
        r = resolve_compose_wrap_choice(
            root,
            wrap_name="tiny.jpg",
            wrap_single_name="",
            wrap_has=True,
            singles_available=[],
        )
        return _ok("低解像はcompose拒否", not r.get("ok"), str(r.get("message") or "")[:80])


def test_hand_hold_select_and_routes() -> Tuple[str, bool, str]:
    import shutil
    import tempfile
    from pathlib import Path

    from b4_wrap_hand_hold import (
        generate_hand_hold_candidates,
        out_dir_for,
        select_hand_hold,
        selected_path,
    )

    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        refs = root / "01_refs"
        refs.mkdir(parents=True)
        wrap = refs / "FOOD_UNIT_WRAP.jpg"
        wrap.write_bytes(b"fake-wrap")
        od = out_dir_for(root)
        od.mkdir(parents=True)
        for i in range(1, 3):
            (od / f"hand_hold_{i:02d}.jpg").write_bytes(f"cand{i}".encode())
        # 既存候補があれば OpenAI なしで reuse
        r = generate_hand_hold_candidates(root, jan="x", force=False)
        ok_reuse = bool(r.get("reused")) and len(r.get("candidates") or []) == 2
        sel = select_hand_hold(root, "hand_hold_02.jpg", jan="x")
        sp = selected_path(root)
        ok_sel = (
            sel.get("selected") == "hand_hold_02.jpg"
            and sp is not None
            and sp.is_file()
            and sel.get("routes", {}).get("mainBlGold") == "hand_hold_02.jpg"
        )
        return _ok(
            "手持ち候補reuse→採用→届け先",
            ok_reuse and ok_sel,
            f"reuse={ok_reuse} sel={ok_sel} routes={sel.get('routes')}",
        )


def main() -> int:
    tests: List[Callable[[], Tuple[str, bool, str]]] = [
        test_lane_default_on,
        test_lane_off,
        test_branch_only_with_wrap,
        test_branch_missing_file,
        test_prompt_bg_and_scale,
        test_prompt_rejects_gray_only,
        test_manifest_select_and_resolve,
        test_select_must_be_candidate,
        test_sub_via_theme_pipeline,
        test_merge_adds_theme_21,
        test_merge_no_wrap,
        test_merge_idempotent,
        test_appeal_theme_off,
        test_has_wrap_lookup_order,
        test_compose_wrap_requires_single_when_present,
        test_compose_wrap_rejects_low_res,
        test_hand_hold_select_and_routes,
    ]
    failed = 0
    for fn in tests:
        name, ok, detail = fn()
        mark = "PASS" if ok else "FAIL"
        print(f"{mark}\t{name}\t{detail}")
        if not ok:
            failed += 1
    print(f"\n{len(tests) - failed}/{len(tests)} PASS")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
