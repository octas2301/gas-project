# -*- coding: utf-8 -*-
"""一片・実物由来ゲート 仮想（画像生成なし・JAN非依存フィクスチャ）。

  cd tools/set_main_image
  python dry_b5_piece_truth_virt.py
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Callable, List, Tuple

from b5_piece_truth_rules import (
    build_shape_policy_prompt_block,
    cube_forbidden_hits,
    evaluate_piece_truth_gate,
    find_identity_same_as_piece,
    infer_shape_family_from_texts,
    piece_truth_gate_mode,
    resolve_preferred_seed_kind,
    shape_variety_policy,
    write_piece_set_source,
)


def _ok(name: str, cond: bool, detail: str = "") -> Tuple[str, bool, str]:
    return (name, cond, detail)


def test_mode_default_warn() -> Tuple[str, bool, str]:
    prev = os.environ.pop("B5_PIECE_TRUTH_GATE", None)
    try:
        mode = piece_truth_gate_mode()
    finally:
        if prev is not None:
            os.environ["B5_PIECE_TRUTH_GATE"] = prev
    return _ok("既定warn", mode == "warn", f"mode={mode}")


def test_mode_off_and_strict() -> Tuple[str, bool, str]:
    prev = os.environ.get("B5_PIECE_TRUTH_GATE")
    try:
        os.environ["B5_PIECE_TRUTH_GATE"] = "0"
        off = piece_truth_gate_mode() == "off"
        os.environ["B5_PIECE_TRUTH_GATE"] = "strict"
        st = piece_truth_gate_mode() == "strict"
    finally:
        if prev is None:
            os.environ.pop("B5_PIECE_TRUTH_GATE", None)
        else:
            os.environ["B5_PIECE_TRUTH_GATE"] = prev
    return _ok("切戻しoff／strict", off and st, "")


def test_family_from_shape_text() -> Tuple[str, bool, str]:
    cube = infer_shape_family_from_texts(["角丸の直方体キューブ。市松2×2。"])
    flake = infer_shape_family_from_texts(["薄く削った帯状の片。緩いねじれ。"])
    other = infer_shape_family_from_texts(["丸い粒。"])
    return _ok(
        "ファミリー推定",
        cube == "cube" and flake == "flake" and other == "other",
        f"cube={cube} flake={flake} other={other}",
    )


def test_cube_forbid_tokens() -> Tuple[str, bool, str]:
    hits = cube_forbidden_hits("角丸キューブだが扇形にくさび状")
    neg = cube_forbidden_hits(
        "角丸の直方体キューブ。完全な幾何四角・扇・三角禁止。向きだけ違う。"
        "直方体ブロック同型。三角・扇・くさびシルエット禁止。"
    )
    pol = shape_variety_policy("cube")
    ok = (
        ("扇形" in hits or "くさび" in hits)
        and (not neg)
        and (not pol["allowSilhouetteScatter"])
    )
    return _ok("cube禁止語（否定は除外）", ok, f"hits={hits} neg={neg}")


def test_flake_allows_scatter() -> Tuple[str, bool, str]:
    pol = shape_variety_policy("flake")
    return _ok("flakeは外形差可", bool(pol["allowSilhouetteScatter"]), "")


def test_policy_prompt_no_sku() -> Tuple[str, bool, str]:
    block = build_shape_policy_prompt_block("cube")
    banned = ("かつお", "まぐろ", "ぶり", "チーズ", "490628")
    hit = [b for b in banned if b in block]
    return _ok("方針文にSKU無し", not hit, f"hit={hit}")


def _touch_jpg(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    # 最小JPEGでなく空でもゲートは存在だけ見る
    path.write_bytes(b"\xff\xd8\xff\xd9")


def test_preferred_human_over_wrap() -> Tuple[str, bool, str]:
    td = Path(tempfile.mkdtemp(prefix="piece_truth_"))
    try:
        _touch_jpg(td / "01_refs" / "FOOD_UNIT_WRAP.jpg")
        _touch_jpg(td / "01_refs" / "PIECE_TRUTH_HUMAN" / "piece_truth_01.png")
        kind, paths = resolve_preferred_seed_kind(td)
        ok = kind == "human" and len(paths) == 1
        return _ok("種優先 human>wrap", ok, f"kind={kind}")
    finally:
        shutil.rmtree(td, ignore_errors=True)


def test_gate_fail_cube_fan() -> Tuple[str, bool, str]:
    td = Path(tempfile.mkdtemp(prefix="piece_truth_"))
    try:
        _touch_jpg(td / "01_refs" / "PIECE_TRUTH_HUMAN" / "piece_truth_01.png")
        _touch_jpg(td / "food_piece_set5_virt" / "piece_set" / "piece_set_01.jpg")
        _touch_jpg(td / "food_piece_set5_virt" / "piece_set" / "piece_set_02.jpg")
        _touch_jpg(td / "food_piece_set5_virt" / "piece_set" / "piece_set_03.jpg")
        write_piece_set_source(td, kind="human", paths=[td / "01_refs" / "PIECE_TRUTH_HUMAN" / "piece_truth_01.png"])
        lang = {
            "pieces": [
                {
                    "name": "piece_set_01.jpg",
                    "fields": {"形状": "角丸の直方体キューブ"},
                    "promptBlock": "【一片セット piece_set_01.jpg】",
                },
                {
                    "name": "piece_set_02.jpg",
                    "fields": {"形状": "直方体キューブ"},
                    "promptBlock": "ok",
                },
                {
                    "name": "piece_set_03.jpg",
                    "fields": {"形状": "扇形のくさびブロック"},
                    "promptBlock": "ng",
                },
            ]
        }
        (td / "food_piece_set5_virt" / "piece_set_lang.json").write_text(
            json.dumps(lang, ensure_ascii=False), encoding="utf-8"
        )
        g = evaluate_piece_truth_gate(td, max_n=3)
        ok = (not g["ok"]) and g["family"] == "cube" and any("cube外形暴走" in e for e in g["errors"])
        return _ok("cube扇はFAIL", ok, f"errors={g['errors']}")
    finally:
        shutil.rmtree(td, ignore_errors=True)


def test_gate_pass_cube_human() -> Tuple[str, bool, str]:
    td = Path(tempfile.mkdtemp(prefix="piece_truth_"))
    try:
        truth = td / "01_refs" / "PIECE_TRUTH_HUMAN" / "piece_truth_01.png"
        _touch_jpg(truth)
        for i in range(1, 4):
            _touch_jpg(td / "food_piece_set5_virt" / "piece_set" / f"piece_set_0{i}.jpg")
        write_piece_set_source(td, kind="human", paths=[truth])
        lang = {
            "pieces": [
                {
                    "name": f"piece_set_0{i}.jpg",
                    "fields": {"形状": "角丸の直方体キューブ。向きだけ違う。"},
                    "promptBlock": f"piece_set_0{i}",
                }
                for i in range(1, 4)
            ]
        }
        (td / "food_piece_set5_virt" / "piece_set_lang.json").write_text(
            json.dumps(lang, ensure_ascii=False), encoding="utf-8"
        )
        g = evaluate_piece_truth_gate(td, max_n=3)
        return _ok("cube+human PASS", bool(g["ok"]), f"g={g}")
    finally:
        shutil.rmtree(td, ignore_errors=True)


def test_gate_fail_human_missing_source() -> Tuple[str, bool, str]:
    td = Path(tempfile.mkdtemp(prefix="piece_truth_"))
    try:
        _touch_jpg(td / "01_refs" / "PIECE_TRUTH_HUMAN" / "piece_truth_01.png")
        _touch_jpg(td / "food_piece_set5_virt" / "piece_set" / "piece_set_01.jpg")
        g = evaluate_piece_truth_gate(td, max_n=1)
        ok = (not g["ok"]) and any("piece_set_source" in e for e in g["errors"])
        return _ok("人間正本あり・出自無しはFAIL", ok, f"errors={g['errors']}")
    finally:
        shutil.rmtree(td, ignore_errors=True)


def test_gate_fail_identity_same_as_piece() -> Tuple[str, bool, str]:
    td = Path(tempfile.mkdtemp(prefix="piece_truth_"))
    try:
        truth = td / "01_refs" / "PIECE_TRUTH_HUMAN" / "piece_truth_01.png"
        _touch_jpg(truth)
        piece = td / "food_piece_set5_virt" / "piece_set" / "piece_set_01.jpg"
        _touch_jpg(piece)
        piece.write_bytes(b"\xff\xd8SAME_AS_IDENT\xff\xd9")
        canon = td / "01_refs" / "FOOD_UNIT_CANON.jpg"
        canon.parent.mkdir(parents=True, exist_ok=True)
        canon.write_bytes(piece.read_bytes())
        write_piece_set_source(td, kind="human", paths=[truth])
        lang = {
            "pieces": [
                {
                    "name": "piece_set_01.jpg",
                    "fields": {"形状": "薄い帯状の片"},
                    "promptBlock": "p1",
                }
            ]
        }
        (td / "food_piece_set5_virt" / "piece_set_lang.json").write_text(
            json.dumps(lang, ensure_ascii=False), encoding="utf-8"
        )
        g = evaluate_piece_truth_gate(td, max_n=1)
        hit = find_identity_same_as_piece(td, max_n=1)
        ok = (
            (not g["ok"])
            and (not hit["ok"])
            and any("一片と同一" in e for e in g["errors"])
        )
        return _ok("IDENTITY＝一片はFAIL", ok, f"errors={g['errors']}")
    finally:
        shutil.rmtree(td, ignore_errors=True)


def test_gate_pass_identity_plating_differs() -> Tuple[str, bool, str]:
    td = Path(tempfile.mkdtemp(prefix="piece_truth_"))
    try:
        truth = td / "01_refs" / "PIECE_TRUTH_HUMAN" / "piece_truth_01.png"
        _touch_jpg(truth)
        piece = td / "food_piece_set5_virt" / "piece_set" / "piece_set_01.jpg"
        _touch_jpg(piece)
        piece.write_bytes(b"\xff\xd8PIECE_ONLY\xff\xd9")
        plating = td / "01_refs" / "FOOD_UNIT_CANON.jpg"
        plating.parent.mkdir(parents=True, exist_ok=True)
        plating.write_bytes(b"\xff\xd8PLATING_DIFF\xff\xd9")
        write_piece_set_source(td, kind="human", paths=[truth])
        lang = {
            "pieces": [
                {
                    "name": "piece_set_01.jpg",
                    "fields": {"形状": "薄い帯状の片"},
                    "promptBlock": "p1",
                }
            ]
        }
        (td / "food_piece_set5_virt" / "piece_set_lang.json").write_text(
            json.dumps(lang, ensure_ascii=False), encoding="utf-8"
        )
        g = evaluate_piece_truth_gate(td, max_n=1)
        return _ok("盛りIDENTITY別物はPASS", bool(g["ok"]), f"errors={g['errors']}")
    finally:
        shutil.rmtree(td, ignore_errors=True)


TESTS: List[Callable[[], Tuple[str, bool, str]]] = [
    test_mode_default_warn,
    test_mode_off_and_strict,
    test_family_from_shape_text,
    test_cube_forbid_tokens,
    test_flake_allows_scatter,
    test_policy_prompt_no_sku,
    test_preferred_human_over_wrap,
    test_gate_fail_cube_fan,
    test_gate_pass_cube_human,
    test_gate_fail_human_missing_source,
    test_gate_fail_identity_same_as_piece,
    test_gate_pass_identity_plating_differs,
]


def main() -> int:
    failed = 0
    for fn in TESTS:
        name, ok, detail = fn()
        mark = "PASS" if ok else "FAIL"
        print(f"[{mark}] {name}" + (f" · {detail}" if detail else ""))
        if not ok:
            failed += 1
    total = len(TESTS)
    print(f"\n{total - failed}/{total} PASS")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
