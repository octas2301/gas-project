# -*- coding: utf-8 -*-
"""MAIN袋・WRAP・採用キーの仮想テスト（画像生成なし）。

使い方:
  cd tools/set_main_image
  python dry_b5_bag_wrap_virt.py

合格の意味:
  - 汎用契約（案件名を焼かない／見本依存／実測比）をコードが満たすこと
  - 提案中の強化文言が指示に載っていること
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Callable, List, Tuple

ROOT = Path(__file__).resolve().parent


def _ok(name: str, cond: bool, detail: str = "") -> Tuple[str, bool, str]:
    return (name, cond, detail)


def test_keep_try_prefers_aspect_ok_among_ocr_pass() -> Tuple[str, bool, str]:
    from b5_main_base_compose import _keep_try_key

    better = _keep_try_key(
        ocr_fail=False, aspect_fail=False, aspect_rel=0.0, rank=10, attempt=2
    )
    worse = _keep_try_key(
        ocr_fail=False, aspect_fail=True, aspect_rel=0.09, rank=5, attempt=1
    )
    return _ok(
        "keep_try: OCR合格どうしは袋OKを先に採用",
        better < worse,
        f"better={better} worse={worse}",
    )


def test_keep_try_food_beats_ocr() -> Tuple[str, bool, str]:
    from b5_main_base_compose import _keep_try_key

    food_ok_ocr_ng = _keep_try_key(
        ocr_fail=True,
        aspect_fail=False,
        aspect_rel=0.0,
        rank=20,
        attempt=2,
        food_fail=False,
    )
    food_ng_ocr_ok = _keep_try_key(
        ocr_fail=False,
        aspect_fail=False,
        aspect_rel=0.0,
        rank=1,
        attempt=1,
        food_fail=True,
    )
    return _ok(
        "keep_try: 個片NGより個片OKを優先（OCR合格でも逆転しない）",
        food_ok_ocr_ng < food_ng_ocr_ok,
        f"food_ok_ocr_ng={food_ok_ocr_ng} food_ng_ocr_ok={food_ng_ocr_ok}",
    )


def test_keep_try_ocr_among_food_ok() -> Tuple[str, bool, str]:
    from b5_main_base_compose import _keep_try_key

    ocr_ok = _keep_try_key(
        ocr_fail=False, aspect_fail=True, aspect_rel=0.2, rank=20, attempt=2, food_fail=False
    )
    ocr_ng = _keep_try_key(
        ocr_fail=True, aspect_fail=False, aspect_rel=0.0, rank=1, attempt=3, food_fail=False
    )
    return _ok(
        "keep_try: 個片OKどうしはOCR合格を優先",
        ocr_ok < ocr_ng,
        f"ocr_ok={ocr_ok} ocr_ng={ocr_ng}",
    )


def test_bag_retry_prefix_keeps_numbers() -> Tuple[str, bool, str]:
    from b5_bag_aspect import _retry_note

    note = _retry_note(
        [
            {
                "role": "PART_1",
                "kind": "aspect",
                "ratio": 1.40,
                "drawnRatio": 1.10,
            }
        ]
    )
    ok = (
        note.startswith("袋の縦横・大きさだけ直せ。他は触るな。")
        and "1.40" in note
        and "1.10" in note
    )
    return _ok("袋やり直し: 短い前置き＋数字メモ", ok, note[:120])


def test_wrap_place_behind_not_front() -> Tuple[str, bool, str]:
    from food_unit_canon import food_unit_prompt_block_main

    block = food_unit_prompt_block_main(
        has_identity=True,
        has_wrap=True,
        food_mode="identity",
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    ok = (
        ("横か後ろ" in block)
        and ("個品より前" in block)
        and ("空気だまり" in block)
        and ("小皿に混ぜるな" in block)
        and ("カメラ最近傍" in block)
        and ("少し混ぜ" not in block)
        and ("チーズ" not in block)
    )
    return _ok(
        "WRAP置き: 横か後ろのみ・混ぜ禁止・最近傍は個品（汎用）",
        ok,
        f"behind={'横か後ろ' in block} no_mix={'小皿に混ぜるな' in block} near={'カメラ最近傍' in block}",
    )


def test_wrap_count_one_to_two() -> Tuple[str, bool, str]:
    from food_unit_canon import food_unit_prompt_block_main

    block = food_unit_prompt_block_main(
        has_identity=True,
        has_wrap=True,
        food_mode="identity",
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    ok = (
        ("1〜2個" in block)
        and ("1〜数個" not in block)
        and ("IDENTITY" in block or "主見本" in block)
    )
    return _ok("WRAP量: 1〜2個・個品主役", ok, f"has12={'1〜2個' in block}")


def test_wrap_prompt_separates_package() -> Tuple[str, bool, str]:
    from food_unit_canon import food_unit_prompt_block_main

    block = food_unit_prompt_block_main(
        has_identity=True,
        has_sizzle=False,
        has_secondary=False,
        has_wrap=True,
        food_mode="identity",
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    need = [
        "WRAP",
        "大袋",
        "PACKAGE",
        "個包装",
    ]
    missing = [k for k in need if k not in block]
    return _ok(
        "WRAP指示: 大袋と個包装を分離",
        not missing,
        f"missing={missing}",
    )


def test_wrap_look_block_injected_not_product_fixed() -> Tuple[str, bool, str]:
    """個包装は機械読みを指示に載せる。案件固定の真空禁止文ではない。"""
    from b5_food_look_read import wrap_look_block_ja
    from b5_virt_abc import _prompt_p3_free

    fake = [
        {
            "label": "構成A・個包装",
            "lines": "密着: 中身に密着／空気: 見えない／材質: 硬いフィルム／縁: ノッチあり／透明度: 中身が見える／形: 正方形に近い",
            "fields": {
                "密着": "中身に密着",
                "空気": "見えない",
                "材質": "硬いフィルム",
                "縁": "ノッチあり",
                "透明度": "中身が見える",
                "形": "正方形に近い",
            },
        }
    ]
    block = wrap_look_block_ja(fake)
    prompt = _prompt_p3_free(
        "材料テスト",
        theme_id="T3",
        name_ja="使い方",
        scene="食卓。",
        has_food_unit_canon=True,
        has_food_unit_wrap=True,
        wrap_look_block=block,
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    ok = (
        "個包装見本の見え方" in prompt
        and "硬いフィルム" in prompt
        and "空気: 見えない" in prompt
        and "チーズかつお" not in prompt
    )
    return _ok(
        "WRAP look: 機械読みが指示に載り案件名は焼かない",
        ok,
        f"has_block={'個包装見本の見え方' in prompt}",
    )


def test_wrap_lock_mentions_air_and_material_from_sample() -> Tuple[str, bool, str]:
    from food_unit_canon import food_unit_prompt_block_main

    block = food_unit_prompt_block_main(
        has_identity=True,
        has_wrap=True,
        food_mode="identity",
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    ok = ("空気" in block) and ("材質" in block or "硬さ" in block or "やわらか" in block)
    return _ok(
        "WRAPロック: 空気・材質は見本／機械読みに従うと書く",
        ok,
        f"air={'空気' in block} material={'材質' in block or '硬さ' in block}",
    )


def test_edge_prompt_forbids_cutout_halo() -> Tuple[str, bool, str]:
    from b5_virt_abc import _prompt_p3_free

    prompt = _prompt_p3_free(
        "材料テスト",
        theme_id="T3",
        name_ja="使い方",
        scene="食卓の俯瞰。",
        has_food_unit_canon=True,
        has_food_unit_wrap=True,
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    ok = (
        ("切り抜き" in prompt or "白フチ" in prompt or "白い" in prompt)
        and ("縁" in prompt)
        and ("浮" in prompt or "柔らか" in prompt)
    )
    return _ok(
        "縁指示: 切り抜き白フチ／浮き禁止が含まれる",
        ok,
        f"len={len(prompt)} cut={'切り抜き' in prompt} white={'白フチ' in prompt or '白い' in prompt}",
    )


def test_scene_strips_wrap_ban_when_wrap() -> Tuple[str, bool, str]:
    from b5_virt_abc import _prompt_p3_free

    prompt = _prompt_p3_free(
        "材料テスト",
        theme_id="T3",
        name_ja="使い方",
        scene="小皿の個品は開封後。個包装を新たに描くな。お茶を置く。",
        has_food_unit_canon=True,
        has_food_unit_wrap=True,
    )
    return _ok(
        "場面: WRAPありなら『個包装を描くな』を除去",
        "個包装を新たに描くな" not in prompt,
        "still_banned" if "個包装を新たに描くな" in prompt else "stripped",
    )


def test_scene_scrubs_old_wrap_mix() -> Tuple[str, bool, str]:
    from b5_virt_abc import _prompt_p3_free, scrub_old_wrap_mix_ja

    old = (
        "手前の小皿に盛る。個包装の見本があるときだけ皿の横・周りに1〜2個置くか"
        "小皿に少し混ぜて出せ。山にしない。"
    )
    scrubbed = scrub_old_wrap_mix_ja(old)
    prompt = _prompt_p3_free(
        "材料テスト",
        theme_id="T4",
        name_ja="成果",
        scene=old,
        has_food_unit_canon=True,
        has_food_unit_wrap=True,
    )
    ok = (
        ("少し混ぜ" not in scrubbed)
        and ("少し混ぜ" not in prompt)
        and ("カメラ最近傍" in prompt)
        and ("小皿に混ぜるな" in prompt)
    )
    return _ok(
        "場面: 旧『混ぜて出せ』を除去し新置きを載せる",
        ok,
        f"scrub_mix={'少し混ぜ' in scrubbed} prompt_mix={'少し混ぜ' in prompt} near={'カメラ最近傍' in prompt}",
    )


def test_height_ratio_from_scale_rows_not_hardcoded() -> Tuple[str, bool, str]:
    from b5_bag_aspect import aspect_prompt_block

    rows = [
        {
            "role": "IMAGE_PACKAGE_PART_1",
            "ratio": 1.53,
            "contentH": 906.0,
            "source": "scale",
        },
        {
            "role": "IMAGE_PACKAGE_PART_2",
            "ratio": 1.22,
            "contentH": 721.0,
            "source": "scale",
        },
    ]
    rel = 721.0 / 906.0
    block = aspect_prompt_block(rows)
    has_rel = f"{rel:.2f}" in block
    has_scale = "SCALE" in block or "並び正本" in block
    return _ok(
        "高さ比: SCALE実測が指示に載り案件定数を焼かない",
        has_rel and has_scale and "0.95" not in block,
        f"rel={rel:.3f} scale_label={has_scale}",
    )


def test_window_contents_locked_in_prompt() -> Tuple[str, bool, str]:
    from b5_virt_abc import _prompt_p3_free

    prompt = _prompt_p3_free(
        "材料テスト",
        theme_id="T3",
        name_ja="使い方",
        scene="食卓。",
        has_food_unit_canon=True,
        mix_parts=["商品A", "商品B"],
        mix_part_bags=True,
    )
    ok = "窓の中" in prompt and "正本" in prompt
    return _ok("窓内中身: 正本どおりの指示がある", ok, f"window={'窓の中' in prompt}")


def test_no_product_name_baked_in_keep() -> Tuple[str, bool, str]:
    import inspect
    from b5_main_base_compose import _keep_try_key

    src = inspect.getsource(_keep_try_key)
    banned = ["チーズ", "かつお", "おだし", "490628", "0.95"]
    hit = [b for b in banned if b in src]
    return _ok("keep_try: 案件定数を焼いていない", not hit, f"hit={hit}")


def main() -> int:
    tests: List[Callable[[], Tuple[str, bool, str]]] = [
        test_keep_try_prefers_aspect_ok_among_ocr_pass,
        test_keep_try_food_beats_ocr,
        test_keep_try_ocr_among_food_ok,
        test_bag_retry_prefix_keeps_numbers,
        test_wrap_place_behind_not_front,
        test_wrap_count_one_to_two,
        test_wrap_prompt_separates_package,
        test_wrap_look_block_injected_not_product_fixed,
        test_wrap_lock_mentions_air_and_material_from_sample,
        test_edge_prompt_forbids_cutout_halo,
        test_scene_strips_wrap_ban_when_wrap,
        test_scene_scrubs_old_wrap_mix,
        test_height_ratio_from_scale_rows_not_hardcoded,
        test_window_contents_locked_in_prompt,
        test_no_product_name_baked_in_keep,
    ]
    rows: List[Tuple[str, bool, str]] = []
    for fn in tests:
        try:
            rows.append(fn())
        except Exception as e:
            rows.append((fn.__name__, False, f"EXC {e}"))
    failed = [r for r in rows if not r[1]]
    print("=== dry_b5_bag_wrap_virt ===")
    for name, ok, detail in rows:
        mark = "PASS" if ok else "FAIL"
        print(f"[{mark}] {name}")
        if detail:
            print(f"       {detail}")
    print(f"total={len(rows)} failed={len(failed)}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
