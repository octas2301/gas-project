# -*- coding: utf-8 -*-
"""
B-⑤ P2 — compose --role rakuten_main_base

人がチェックした案（最大はレ点子SKU数。案の上限12）だけ OpenAI で画像を作る。
正本袋必須。金丸・Octas は Pillow。02 へは書かない（P3）。

  python b5_main_base_compose.py --role rakuten_main_base --from-json checked.json
  python b5_main_base_compose.py --role rakuten_main_base --from-json checked.json --dry-run
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from PIL import Image, ImageFilter

from b5_grid_budget import (
    build_compose_grid_prompt_block,
    build_retry_note_for_hit,
    classify_grid_tier,
    compute_grid_budget,
    grid_needs_retry,
    grid_rank_penalty,
    retry_threshold_cells,
)
from b5_catch_zone_adopted import pattern_by_id
from b5_grid_gemini_measure import (
    DEFAULT_GEMINI_GRID_MODEL,
    GeminiGridMeasurer,
    archive_try_with_overlays,
    build_compose_tries_html,
    measure_grid_dual,
)
from b5_virt_abc import (
    CANVAS,
    OPENAI_QUALITY_DEFAULT,
    P3_FREE_MAX_TRIES,
    P3_TEXT_MAX_CELLS,
    _checkerboard_suspected,
    _compact_facts_from_sheets,
    _finish_p3_overlays,
    _gen_ai,
    _gold_params,
    _load_gold_disk,
    _oshare_typo_paths,
    _prompt_p3_free,
    _write_grid20_p3,
    _write_style_sheet,
    reflect_checked_scene_hints,
    sets_from_proposals,
)
from b5_virt_ai_headline import _to_1200
from b5_main_base_names import (
    catch_head_ja,
    catch_sales_pair_ja,
    compose_dir_name,
    derive_short_name,
    pick_filename,
    sanitize_fs_name,
)
from openai_image import resolve_openai_image_model
from b5_layer_stamp import (
    BAG_BOX,
    FOOD_BOX,
    N1_BAG_BOX,
    N1_FOOD_BOX,
    collect_stamp_assets,
    dump_stamp_meta,
    stamp_layers_enabled,
    stamp_truth_layers,
    write_zone_guide,
)
from b5_bag_color import color_prompt_block, measure_role_colors
from b5_bag_aspect import (
    aspect_prompt_block,
    aspect_worst_rel,
    check_drawn_bags,
    check_face_vs_tableware,
    face_cm_prompt_block,
    format_bag_verdict_text,
    measure_mix_aspects_from_scale,
    measure_role_aspects,
    write_bag_verdict_bundle,
)
from b5_catch_ocr_check import check_catch_text
from b5_food_look_read import (
    check_drawn_foods,
    check_drawn_wraps,
    look_block_ja,
    read_looks_for_paths,
    read_wrap_looks_for_paths,
    wrap_look_block_ja,
)
from b5_photo_pass import (
    MODE_JA,
    apply_photo_pass,
    resolve_photo_modes,
    write_compare_html,
)
from b5_piece_truth_rules import (
    build_shape_policy_prompt_block,
    evaluate_piece_truth_gate,
    piece_truth_gate_mode,
)
from food_unit_canon import (
    combo_parts_for_mix_name,
    food_scale_lock_prompt_block,
    load_part_bundle,
    identity_plating_prompt_blocks_for_parts,
    piece_lang_prompt_blocks_for_parts,
    piece_set_lock_prompt_block,
    resolve_main_food_refs,
    resolve_package_bag_for_name,
    resolve_piece_set_paths,
    theme_new_dir_for_jan,
)
from text_color import resolve_text_color
from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT, resolve_octas

LOG = logging.getLogger("set_main_image.b5_main_base_compose")

ROLE = "rakuten_main_base"
MAX_PROPOSALS = 12


def _attach_piece_set_refs(
    img_paths: List[Path],
    img_roles: List[str],
    *,
    jans: List[str],
    max_n: int,
) -> Dict[str, Any]:
    """一片セット／S1時は SURFACE＋SHAPE を同送。max_n<=0 なら何もしない。"""
    flags: Dict[str, Any] = {
        "pieceSetMax": int(max_n or 0),
        "pieceSetN": 0,
        "surfaceN": 0,
        "shapeN": 0,
        "s1": False,
        "jans": [],
    }
    if int(max_n or 0) <= 0:
        return flags
    seen = {str(p.resolve()) for p in img_paths if p.is_file()}
    # S1: 自SKU一片本体を付けず SURFACE→SHAPE
    try:
        from b5_piece_deepen import s1_omit_own_piece_enabled, surface_shape_split_enabled
        from b5_piece_truth_rules import (
            resolve_shape_donor_paths,
            resolve_surface_ref_paths,
        )

        if surface_shape_split_enabled() and s1_omit_own_piece_enabled() and jans:
            jan0 = str(jans[0])
            tn = theme_new_dir_for_jan(jan0)
            smeta = resolve_surface_ref_paths(tn, max_n=max_n)
            surfs = [Path(p) for p in (smeta.get("paths") or []) if Path(p).is_file()]
            dmeta = resolve_shape_donor_paths(tn, max_n=max_n)
            donors = [Path(p) for p in (dmeta.get("paths") or []) if Path(p).is_file()]
            if surfs or donors:
                flags["s1"] = True
                flags["jans"].append(jan0)
                for i, p in enumerate(surfs, 1):
                    rp = p.resolve()
                    if str(rp) in seen:
                        continue
                    img_paths.append(p)
                    img_roles.append(f"IMAGE_FOOD_UNIT_SURFACE_{i}")
                    seen.add(str(rp))
                    flags["surfaceN"] += 1
                for i, p in enumerate(donors, 1):
                    rp = p.resolve()
                    if str(rp) in seen:
                        continue
                    img_paths.append(p)
                    img_roles.append(f"IMAGE_FOOD_UNIT_SHAPE_{i}")
                    seen.add(str(rp))
                    flags["shapeN"] += 1
                LOG.info(
                    "attach S1 surfaceN=%s shapeN=%s (no own piece)",
                    flags["surfaceN"],
                    flags["shapeN"],
                )
                return flags
    except Exception as e:
        LOG.warning("attach S1 skip, fallback piece_set: %s", e)
    n_added = 0
    for jan_d in jans:
        paths = resolve_piece_set_paths(jan_d, max_n=max_n)
        if not paths:
            continue
        flags["jans"].append(jan_d)
        for p in paths:
            rp = p.resolve()
            if str(rp) in seen:
                continue
            n_added += 1
            img_paths.append(p)
            img_roles.append(f"IMAGE_FOOD_UNIT_PIECE_{n_added}")
            seen.add(str(rp))
    flags["pieceSetN"] = n_added
    return flags


def _attach_food_refs(
    img_paths: List[Path],
    img_roles: List[str],
    refs: Dict[str, Any],
) -> Dict[str, Any]:
    """IDENTITY／SIZZLE／WRAP／SECONDARY を OpenAI 参照に足す。"""
    flags = {
        "hasIdentity": False,
        "hasSizzle": False,
        "hasWrap": False,
        "hasSecondary": False,
        "foodMode": str(refs.get("mode") or "package"),
        "reason": str(refs.get("reason") or ""),
        "confirmedBy": str(refs.get("confirmedBy") or ""),
        "themeNewDir": str(refs.get("themeNewDir") or ""),
    }
    ident = refs.get("identity")
    if ident and Path(ident).is_file():
        p = Path(ident)
        if p.resolve() != img_paths[0].resolve():
            img_paths.append(p)
            img_roles.append("IMAGE_FOOD_UNIT_CANON")
            flags["hasIdentity"] = True
    sizzle = refs.get("sizzle")
    if sizzle and Path(sizzle).is_file():
        img_paths.append(Path(sizzle))
        img_roles.append("IMAGE_FOOD_UNIT_SIZZLE")
        flags["hasSizzle"] = True
    wrap = refs.get("wrap")
    if wrap and Path(wrap).is_file():
        img_paths.append(Path(wrap))
        img_roles.append("IMAGE_FOOD_UNIT_WRAP")
        flags["hasWrap"] = True
    for i, sec in enumerate(list(refs.get("secondary") or []), start=1):
        if sec and Path(sec).is_file():
            img_paths.append(Path(sec))
            img_roles.append(f"IMAGE_FOOD_UNIT_SECONDARY_{i}")
            flags["hasSecondary"] = True
    return flags


def _attach_mix_parts(
    img_paths: List[Path],
    img_roles: List[str],
    combo_parts: List[Dict[str, Any]],
    *,
    include_secondary: bool = True,
) -> Dict[str, Any]:
    """構成ごとの袋＋中身。判定は商品名ベース。"""
    flags = {
        "mix": True,
        "partNames": [],
        "hasIdentity": False,
        "hasSizzle": False,
        "hasWrap": False,
        "hasSecondary": False,
        "hasPartBags": False,
        "foodMode": "identity",
        "reason": "mix_parts",
        "confirmedBy": "human",
    }
    seen = {str(p.resolve()) for p in img_paths if p.is_file()}
    for i, raw in enumerate(combo_parts, start=1):
        if not isinstance(raw, dict):
            continue
        nb = str(raw.get("nameBase") or raw.get("name") or "").strip()
        jan_p = "".join(ch for ch in str(raw.get("jan") or "") if ch.isdigit())
        if not nb and not jan_p:
            continue
        bundle = load_part_bundle(
            name_base=nb or jan_p,
            jan_for_folder=jan_p,
            parent_sku=str(raw.get("parentSku") or ""),
            bag_path=str(raw.get("packageTruth") or raw.get("bag") or ""),
        )
        flags["partNames"].append(nb or jan_p)
        bag = bundle.get("bag")
        if bag and Path(bag).is_file():
            rp = Path(bag).resolve()
            if str(rp) not in seen:
                img_paths.append(Path(bag))
                img_roles.append(f"IMAGE_PACKAGE_PART_{i}")
                seen.add(str(rp))
                flags["hasPartBags"] = True
        ident = bundle.get("identity")
        if ident and Path(ident).is_file():
            img_paths.append(Path(ident))
            img_roles.append(f"IMAGE_FOOD_UNIT_IDENTITY_{i}")
            flags["hasIdentity"] = True
        sizzle = bundle.get("sizzle")
        if sizzle and Path(sizzle).is_file():
            img_paths.append(Path(sizzle))
            img_roles.append(f"IMAGE_FOOD_UNIT_SIZZLE_{i}")
            flags["hasSizzle"] = True
        wrap = bundle.get("wrap")
        if wrap and Path(wrap).is_file():
            img_paths.append(Path(wrap))
            img_roles.append(f"IMAGE_FOOD_UNIT_WRAP_{i}")
            flags["hasWrap"] = True
        if not include_secondary:
            continue
        for j, sec in enumerate(list(bundle.get("secondary") or []), start=1):
            if sec and Path(sec).is_file():
                img_paths.append(Path(sec))
                img_roles.append(f"IMAGE_FOOD_UNIT_SECONDARY_{i}_{j}")
                flags["hasSecondary"] = True
    return flags


def _look_items_from_mix(combo_parts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    seen = set()
    for raw in combo_parts:
        if not isinstance(raw, dict):
            continue
        nb = str(raw.get("nameBase") or raw.get("name") or "").strip()
        jan_p = "".join(ch for ch in str(raw.get("jan") or "") if ch.isdigit())
        if not nb and not jan_p:
            continue
        bundle = load_part_bundle(
            name_base=nb or jan_p,
            jan_for_folder=jan_p,
            parent_sku=str(raw.get("parentSku") or ""),
            bag_path=str(raw.get("packageTruth") or raw.get("bag") or ""),
        )
        ident = bundle.get("identity")
        if ident and Path(ident).is_file():
            key = str(Path(ident).resolve())
            if key in seen:
                continue
            seen.add(key)
            items.append({"label": nb or jan_p, "path": Path(ident)})
    return items


def _wrap_items_from_mix(combo_parts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    seen = set()
    for raw in combo_parts:
        if not isinstance(raw, dict):
            continue
        nb = str(raw.get("nameBase") or raw.get("name") or "").strip()
        jan_p = "".join(ch for ch in str(raw.get("jan") or "") if ch.isdigit())
        if not nb and not jan_p:
            continue
        bundle = load_part_bundle(
            name_base=nb or jan_p,
            jan_for_folder=jan_p,
            parent_sku=str(raw.get("parentSku") or ""),
            bag_path=str(raw.get("packageTruth") or raw.get("bag") or ""),
        )
        wrap = bundle.get("wrap")
        if wrap and Path(wrap).is_file():
            key = str(Path(wrap).resolve())
            if key in seen:
                continue
            seen.add(key)
            items.append({"label": (nb or jan_p) + "・個包装", "path": Path(wrap)})
    return items


def _look_items_from_refs(food_by_theme: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    seen = set()
    for tid, refs in food_by_theme.items():
        ident = (refs or {}).get("identity")
        if ident is True or not ident:
            continue
        p = Path(str(ident))
        if not p.is_file():
            continue
        key = str(p.resolve())
        if key in seen:
            continue
        seen.add(key)
        items.append({"label": str(tid or "見本"), "path": p})
    return items


def _wrap_items_from_refs(food_by_theme: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    seen = set()
    for tid, refs in food_by_theme.items():
        wrap = (refs or {}).get("wrap")
        if wrap is True or not wrap:
            continue
        p = Path(str(wrap))
        if not p.is_file():
            continue
        key = str(p.resolve())
        if key in seen:
            continue
        seen.add(key)
        items.append({"label": str(tid or "見本") + "・個包装", "path": p})
    return items


def _safe_name(s: str) -> str:
    return sanitize_fs_name(s, max_len=48)


def _load_payload(path: Path) -> Dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(data, dict):
        raise SystemExit("--from-json はオブジェクトである必要があります")
    return data


def _normalize_proposals(
    raw: Any,
    *,
    materials: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    if not isinstance(raw, list) or not raw:
        raise SystemExit("proposals が空です（チェックした案が必要）")
    raw_list = list(raw)
    out: List[Dict[str, Any]] = []
    for i, row in enumerate(raw_list[:MAX_PROPOSALS]):
        if not isinstance(row, dict):
            continue
        catch = str(row.get("catch") or "").strip()
        theme_id = str(row.get("themeId") or f"T{i + 1}").strip()
        if not catch:
            LOG.warning("skip empty catch themeId=%s", theme_id)
            continue
        catch_burn = str(row.get("catchBurn") or "").strip()
        sales_talk = str(row.get("salesTalk") or "").strip()
        sales_burn = str(row.get("salesBurn") or "").strip() or sales_talk
        burn_skip = bool(row.get("burnSkip"))
        out.append(
            {
                "themeId": theme_id,
                "themeName": str(row.get("themeName") or theme_id),
                "catch": catch,
                "catchBurn": "" if burn_skip else (catch_burn or catch),
                "burnSkip": burn_skip,
                "burnSkipReason": str(row.get("burnSkipReason") or ""),
                "salesTalk": sales_talk,
                "salesBurn": sales_burn,
                "evidence": str(row.get("evidence") or ""),
                "sceneHint": str(row.get("sceneHint") or ""),
                "sellRank": row.get("sellRank") or (i + 1),
                "sellScore": row.get("sellScore"),
                # ミックス: 料理の上に載せる構成（落とすと他構成が料理に混ざる）
                "dishName": str(row.get("dishName") or row.get("dish") or ""),
                "dishPartName": str(row.get("dishPartName") or ""),
                "dishPartSku": str(row.get("dishPartSku") or ""),
                "dishPartJan": str(row.get("dishPartJan") or ""),
                # レビュー／再生成メタ（落とすと個品微調整・構図維持が効かない）
                "_foodTuneNote": str(row.get("_foodTuneNote") or "").strip(),
                "_regenTextOnly": bool(row.get("_regenTextOnly")),
                "_regenReference": str(row.get("_regenReference") or "").strip(),
            }
        )
    if not out:
        raise SystemExit("有効な proposals（catch あり）がありません")
    return out[:MAX_PROPOSALS]


_SIZE_ARMS = {
    "scene_only": (
        "【サイズ試験・この1文だけ】\n"
        "変えてよいのは場面だけ。包材の形・縦横・印字は、添付の見本画像のまま。"
        "包材を描き直して別の形にするな。上の「場面の中に描け」より、この1文を優先。\n"
    ),
    "front": (
        "【サイズ試験・この1文だけ】\n"
        "包材の面はカメラと平行。カメラは目の高さ。上からも横からも傾けない。"
        "中望遠。近い面だけが大きく見える斜めは禁止。"
        "50mmや浅いボケの指定より、この正面を優先。\n"
    ),
    "ref_object": (
        "【サイズ試験・この1文だけ】\n"
        "包材と同じ面に、大きさの分かる器を1つだけ置く。"
        "包材とその器は同じ距離で、両方ピントを合わせる。"
        "大きさをごまかす斜めは禁止。手前の盛り付けは比に使わない。\n"
    ),
    "part_first": "",
    "ratio_card": (
        "【サイズ試験・この1つだけ】\n"
        "IMAGE_ASPECT_CARD は、包材の面の縦横の形だけ。完成画の面の形を、このカードに合わせろ。"
        "絵柄は IMAGE_PACKAGE_PART のまま。カード自体は完成画に置くな。\n"
    ),
    "paste_only": (
        "【サイズ試験・この1文だけ】\n"
        "包材は描くな。食品と場面だけ。包材を置く場所は、盛り付けの後ろに空けておく。\n"
    ),
    "paste_shadow": (
        "【サイズ試験・この1文だけ】\n"
        "包材は描くな。食品と場面だけ。包材を置く場所は、盛り付けの後ろに空けておく。\n"
    ),
    "paste_relight": (
        "【サイズ試験・この1文だけ】\n"
        "包材は描くな。食品と場面だけ。包材を置く場所は、盛り付けの後ろに空けておく。\n"
    ),
}


def _move_parts_first(
    paths: List[Path], roles: List[str]
) -> tuple:
    parts = [
        (p, r)
        for p, r in zip(paths, roles)
        if str(r).startswith("IMAGE_PACKAGE_PART_")
    ]
    rest = [
        (p, r)
        for p, r in zip(paths, roles)
        if not str(r).startswith("IMAGE_PACKAGE_PART_")
    ]
    merged = parts + rest
    return [p for p, _ in merged], [r for _, r in merged]


def _write_aspect_cards(meta: Path, rows: List[Dict[str, Any]]) -> List[tuple]:
    out = []
    for row in rows:
        role = str(row.get("role") or "")
        if not role.startswith("IMAGE_PACKAGE_PART_"):
            continue
        ratio = float(row.get("ratio") or 0)
        if ratio <= 0:
            continue
        width = 240
        height = max(8, int(round(width * ratio)))
        card = Image.new("RGB", (width + 48, height + 48), (236, 236, 236))
        slab = Image.new("RGB", (width, height), (24, 24, 24))
        card.paste(slab, (24, 24))
        path = meta / f"{role}_aspect_card.jpg"
        card.save(path, quality=90)
        out.append((path, role.replace("IMAGE_PACKAGE_PART_", "IMAGE_ASPECT_CARD_")))
    return out


def _paste_parts_with_shadow(
    canvas: Image.Image, rows: List[Dict[str, Any]], *, shadow: bool
) -> Image.Image:
    parts = [r for r in rows if str(r.get("role") or "").startswith("IMAGE_PACKAGE_PART_")]
    if not parts:
        return canvas
    base = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    base.paste(canvas.convert("RGBA"))
    width, height = base.size
    base_h = float(parts[0].get("contentH") or 0) or 1.0
    target_h = int(height * 0.46)
    pieces = []
    for row in parts:
        src = Image.open(row["path"]).convert("RGBA")
        mask = src.split()[-1].point(lambda px: 255 if px > 16 else 0)
        box = mask.getbbox()
        if not box:
            continue
        src = src.crop(box)
        rel = float(row.get("contentH") or base_h) / base_h
        piece_h = max(8, int(round(target_h * rel)))
        piece_w = max(8, int(round(piece_h * src.size[0] / float(src.size[1]))))
        pieces.append(src.resize((piece_w, piece_h), Image.Resampling.LANCZOS))
    if not pieces:
        return canvas
    gap = int(width * 0.02)
    total_w = sum(p.size[0] for p in pieces) + gap * (len(pieces) - 1)
    x = max(0, (width - total_w) // 2)
    bottom = int(height * 0.62)
    for src in pieces:
        piece_w, piece_h = src.size
        y = max(0, bottom - piece_h)
        alpha = src.split()[-1]
        if shadow:
            shade = Image.new("RGBA", src.size, (0, 0, 0, 0))
            shade.putalpha(alpha.point(lambda px: int(px * 0.45)))
            shade = shade.filter(ImageFilter.GaussianBlur(8))
            layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
            layer.paste(shade, (x, min(height - 4, y + 10)), shade)
            base = Image.alpha_composite(base, layer)
        base.paste(src, (x, y), src)
        x += piece_w + gap
    return base.convert("RGB")


def _keep_try_key(
    *,
    ocr_fail: bool,
    aspect_fail: bool,
    aspect_rel: Optional[float],
    rank: int,
    attempt: int,
    food_fail: bool = False,
    face_fail: bool = False,
    wrap_fail: bool = False,
) -> tuple:
    """個片（IDENTITY）不合格を最優先で落とす。次にキャッチOCR。同群内は袋・顔・WRAP。"""
    food_group = 1 if food_fail else 0
    ocr_group = 1 if ocr_fail else 0
    soft_fail = int(bool(aspect_fail)) + int(bool(face_fail)) + int(bool(wrap_fail))
    if aspect_rel is None:
        aspect_group = 2
        err = 0.0
    elif aspect_fail:
        aspect_group = 1
        err = float(aspect_rel)
    else:
        aspect_group = 0
        err = 0.0
    return (food_group, ocr_group, soft_fail, aspect_group, err, int(rank), int(attempt))


def _size_arm_block(arm: str) -> str:
    key = str(arm or "").strip()
    if not key:
        return ""
    if key not in _SIZE_ARMS:
        raise SystemExit(f"--size-arm は {', '.join(_SIZE_ARMS)} のいずれか（指定={key!r}）")
    return _SIZE_ARMS[key]


def compose_rakuten_main_base(
    *,
    jan: str,
    parent_sku: str,
    package_truth: Path,
    proposals: List[Dict[str, Any]],
    text_color: str = "1",
    openai_quality: str = OPENAI_QUALITY_DEFAULT,
    work_root: Optional[Path] = None,
    dry_run: bool = False,
    short_name: str = "",
    out_dir: Optional[Path] = None,
    skip_grid_gemini: bool = False,
    gemini_grid_model: str = DEFAULT_GEMINI_GRID_MODEL,
    combo_parts: Optional[List[Dict[str, Any]]] = None,
    name_base: str = "",
    stamp_layers: Optional[bool] = None,
    photo_pass: Optional[str] = None,
    catch_ocr: bool = True,
    package_face_h_cm: Optional[float] = None,
    package_face_w_cm: Optional[float] = None,
    size_arm: str = "",
    include_food_secondary: bool = True,
    photo_arm: str = "",
    mix_explicit: Optional[bool] = None,
    food_explicit: Optional[bool] = None,
    piece_set_max: int = 0,
) -> Dict[str, Any]:
    """チェック済み案だけ画像生成。02 未書込。

    out_dir 指定時（レビュー再生成）: 新規タイムスタンプフォルダを作らず、その比較へ出力。
    piece_set_max: 一片セット添付枚数（0=オフ／3=先頭3／5以上=あるだけ）。既定オフ。
    """
    jan_d = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
    if not jan_d:
        raise SystemExit("jan が空")
    truth = Path(package_truth)
    def _mix_part_ok(p: Any) -> bool:
        if not isinstance(p, dict):
            return False
        return bool(
            str(p.get("nameBase") or p.get("name") or "").strip()
            or str(p.get("jan") or "").strip()
            or str(p.get("parentSku") or "").strip()
        )

    mix_list = [p for p in (combo_parts or []) if _mix_part_ok(p)]
    if mix_explicit is False:
        mix_list = []
        combo_parts = []
        LOG.info("mix from sheet=false; BOM skipped")
    elif mix_explicit is True:
        LOG.info(
            "mix from sheet=true parts=%s; auto BOM skipped",
            len(mix_list),
        )
    elif len(mix_list) < 2:
        filled = combo_parts_for_mix_name(
            name_base=str(name_base or "").strip(),
            parent_sku=parent_sku,
            work_root=work_root,
        )
        if len(filled) >= 2:
            combo_parts = filled
            mix_list = filled
            LOG.info(
                "combo parts from master BOM n=%s names=%s",
                len(filled),
                [p.get("nameBase") or p.get("jan") for p in filled],
            )
    if food_explicit is False:
        include_food_secondary = False
    if len(mix_list) >= 2:
        mix_bag = resolve_package_bag_for_name(
            name_base=str(name_base or "").strip(),
            parent_sku=parent_sku,
            jan_for_folder=jan_d,
            bag_path=str(truth) if truth.is_file() else "",
        )
        if mix_bag is not None:
            truth = mix_bag
            LOG.info("mix parent bag nameBase=%s path=%s", name_base, truth)
    if not truth.is_file():
        raise SystemExit(f"正本なし: {truth}")
    LOG.info("include_food_secondary=%s photo_arm=%s", include_food_secondary, photo_arm or "-")

    mat = _compact_facts_from_sheets(jan_d)
    facts_line = str(mat.get("factsLine") or "")
    materials = mat.get("materials") or {}

    props = _normalize_proposals(proposals, materials=materials)
    props = reflect_checked_scene_hints(
        props, materials, mix=len(mix_list) >= 2, food=food_explicit
    )
    for p in props:
        LOG.info(
            "scene theme=%s hint=%s",
            p.get("themeId"),
            str(p.get("sceneHint") or "")[:180],
        )
    color = resolve_text_color(text_color)
    work = Path(work_root or DEFAULT_WORK_ROOT)

    short = derive_short_name(
        package_truth=truth,
        explicit=short_name,
        materials_name=str((materials or {}).get("name") or ""),
    )
    reuse_dir = Path(out_dir) if out_dir is not None else None
    if reuse_dir is not None:
        out_root = reuse_dir
        out_root.mkdir(parents=True, exist_ok=True)
        dir_name = out_root.name
    else:
        dir_name = compose_dir_name(short)
        out_root = work / SUB_TEST_OUT / "b5_main_base_compose" / dir_name
        out_root.mkdir(parents=True, exist_ok=True)
    meta = out_root / "_meta"
    meta.mkdir(exist_ok=True)
    pick_dir = out_root / "比較"
    pick_dir.mkdir(exist_ok=True)
    grid_dir = out_root / "_grid20"
    grid_dir.mkdir(exist_ok=True)
    tries_dir = out_root / "_tries"
    tries_dir.mkdir(exist_ok=True)

    use_grid_gemini = not skip_grid_gemini and not (
        os.environ.get("B5_GRID_GEMINI", "1").strip().lower() in ("0", "false", "no")
    )
    grid_measurer: Optional[GeminiGridMeasurer] = None
    if use_grid_gemini:
        grid_measurer = GeminiGridMeasurer(model=gemini_grid_model)
        LOG.info("grid gate=local catch_v8b (gemini=参考 B5_GRID_GATE=gemini で切替)")
    else:
        LOG.info("grid gate=local catch_v8b only (B5_GRID_GEMINI=0)")

    shutil.copy2(truth, meta / f"package_truth{truth.suffix.lower() or '.png'}")

    run_sets = sets_from_proposals(props)
    props_by_tid = {
        str(p.get("themeId") or ""): p for p in props if isinstance(p, dict)
    }

    # 個品見本: 組み合わせは構成の商品名ベース。単品は従来の theme_new_run
    mix_list = [p for p in (combo_parts or []) if isinstance(p, dict) and str(p.get("nameBase") or p.get("name") or "").strip()]
    food_by_theme: Dict[str, Dict[str, Any]] = {}
    if len(mix_list) >= 2:
        mix_flags = _attach_mix_parts(
            [], [], mix_list, include_secondary=include_food_secondary
        )
        for spec in run_sets:
            tid = str(spec.get("themeId") or "")
            food_by_theme[tid] = {
                "mode": "identity",
                "confirmedBy": "human",
                "reason": "mix_parts",
                "identity": True if mix_flags.get("hasIdentity") else None,
                "sizzle": True if mix_flags.get("hasSizzle") else None,
                "wrap": True if mix_flags.get("hasWrap") else None,
                "secondary": [1] if mix_flags.get("hasSecondary") else [],
                "mixParts": mix_list,
                "hasPartBags": bool(mix_flags.get("hasPartBags")),
                "partNames": list(mix_flags.get("partNames") or []),
            }
    else:
        for spec in run_sets:
            tid = str(spec.get("themeId") or "")
            food_by_theme[tid] = resolve_main_food_refs(jan_d, tid, require_human=True)
    food_snapshot = {
        tid: {
            "confirmedBy": r.get("confirmedBy"),
            "reason": r.get("reason"),
            "mode": r.get("mode"),
            "themeNewDir": r.get("themeNewDir"),
            "identity": str(r["identity"]) if r.get("identity") else "",
            "sizzle": str(r["sizzle"]) if r.get("sizzle") else "",
            "wrap": str(r["wrap"]) if r.get("wrap") else "",
            "secondaryN": len(r.get("secondary") or []),
        }
        for tid, r in food_by_theme.items()
    }
    (meta / "food_refs.json").write_text(
        json.dumps(food_snapshot, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    LOG.info("food refs jan=%s %s", jan_d, food_snapshot)
    look_items = (
        _look_items_from_mix(mix_list)
        if len(mix_list) >= 2
        else _look_items_from_refs(food_by_theme)
    )
    wrap_items = (
        _wrap_items_from_mix(mix_list)
        if len(mix_list) >= 2
        else _wrap_items_from_refs(food_by_theme)
    )
    food_looks = read_looks_for_paths(look_items) if look_items else []
    food_look_block = look_block_ja(food_looks)
    # 一片作成時の機械言語化を同送（写真lookだけでは縁・厚み・筋が落ちる）
    piece_lang_parts = mix_list if len(mix_list) >= 2 else []
    if not piece_lang_parts and jan_d:
        piece_lang_parts = [{"jan": jan_d, "nameBase": str(name_base or jan_d)}]
    piece_set_n = int(piece_set_max or 0)
    # 添付枚数と言語化番号を一致（§4.5.10bis-T）。max_n=0 は一片オフ＝従来全文。
    piece_lang_block = piece_lang_prompt_blocks_for_parts(
        piece_lang_parts, max_n=piece_set_n
    )
    plating_block = identity_plating_prompt_blocks_for_parts(piece_lang_parts)
    piece_lock_block = piece_set_lock_prompt_block(max_n=piece_set_n)
    # F2: 実物由来＋外形ファミリー方針（SKU名を焼かない）
    piece_truth_gates: List[Dict[str, Any]] = []
    shape_policy_block = ""
    gate_mode = piece_truth_gate_mode()
    if piece_set_n > 0 and gate_mode != "off":
        seen_gate_jan: set[str] = set()
        for raw in piece_lang_parts:
            if not isinstance(raw, dict):
                continue
            jd = "".join(ch for ch in str(raw.get("jan") or "") if ch.isdigit())
            if not jd or jd in seen_gate_jan:
                continue
            seen_gate_jan.add(jd)
            theme_dir = theme_new_dir_for_jan(jd)
            gate = evaluate_piece_truth_gate(theme_dir, max_n=piece_set_n)
            gate["jan"] = jd
            piece_truth_gates.append(gate)
            if gate.get("policyPrompt") and not shape_policy_block:
                shape_policy_block = build_shape_policy_prompt_block(
                    str(gate.get("family") or "other")
                )
            if gate.get("ok"):
                LOG.info(
                    "piece_truth_gate PASS jan=%s family=%s kind=%s",
                    jd,
                    gate.get("family"),
                    gate.get("preferredKind"),
                )
            else:
                LOG.warning(
                    "piece_truth_gate FAIL jan=%s errors=%s",
                    jd,
                    gate.get("errors"),
                )
                if gate_mode == "strict":
                    raise RuntimeError(
                        "piece_truth_gate strict FAIL jan=%s errors=%s"
                        % (jd, gate.get("errors"))
                    )
        (meta / "piece_truth_gate.json").write_text(
            json.dumps(
                {"mode": gate_mode, "gates": piece_truth_gates},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
    has_wrap_any = bool(wrap_items) or any(
        bool((food_by_theme.get(t) or {}).get("wrap")) for t in food_by_theme
    )
    piece_size_hint = ""
    try:
        from b5_slot_lang import piece_size_hint_from_jan

        if jan_d:
            piece_size_hint = piece_size_hint_from_jan(jan_d, max_n=max(piece_set_n, 3))
    except Exception:
        piece_size_hint = ""
    scale_lock = food_scale_lock_prompt_block(
        has_wrap=has_wrap_any, piece_size_hint=piece_size_hint
    )
    # S1: SURFACE/SHAPE 言語化＋合成の正（自SKU一片本体は後段で添付しない）
    s1_food_extra = ""
    try:
        from b5_piece_deepen import s1_omit_own_piece_enabled, surface_shape_split_enabled
        from b5_piece_truth_rules import (
            resolve_shape_donor_paths,
            resolve_surface_ref_paths,
            shape_donor_attach_prompt_block,
            surface_attach_prompt_block,
        )
        from b5_slot_lang import (
            build_shape_lang_prompt_from_donors,
            build_surface_lang_prompt_block,
            s1_synthesis_prompt_block,
        )

        if (
            piece_set_n > 0
            and jan_d
            and surface_shape_split_enabled()
            and s1_omit_own_piece_enabled()
        ):
            tn = theme_new_dir_for_jan(jan_d)
            smeta = resolve_surface_ref_paths(tn, max_n=piece_set_n)
            surfs = [Path(p) for p in (smeta.get("paths") or []) if Path(p).is_file()]
            try:
                from b5_generic_plating import (
                    generic_plating_core_enabled,
                    prefer_face_crop_surfaces,
                )

                if generic_plating_core_enabled() and surfs:
                    surfs = prefer_face_crop_surfaces(tn, surfs, max_n=piece_set_n)
            except Exception as ge:
                LOG.warning("compose generic face crop skip: %s", ge)
            dmeta = resolve_shape_donor_paths(tn, max_n=piece_set_n)
            donors = [Path(p) for p in (dmeta.get("paths") or []) if Path(p).is_file()]
            if surfs or donors:
                s1_food_extra = "\n".join(
                    x
                    for x in [
                        s1_synthesis_prompt_block(),
                        surface_attach_prompt_block(max_n=len(surfs)) if surfs else "",
                        build_surface_lang_prompt_block(tn, surfs, ensure=True)
                        if surfs
                        else "",
                        shape_donor_attach_prompt_block(max_n=len(donors))
                        if donors
                        else "",
                        build_shape_lang_prompt_from_donors(
                            [Path(x) for x in (dmeta.get("donorDirs") or [])],
                            max_n=len(donors),
                        )
                        if donors
                        else "",
                        "【MAIN・S1】自SKU一片本体は添付しない。"
                        "色・筋＝SURFACE／形＝SHAPE／盛り＝IDENTITY。",
                    ]
                    if x
                )
                # 一片本体ロック／一片言語化はS1では載せない（汚染防止）
                piece_lock_block = ""
                piece_lang_block = ""
                LOG.info(
                    "compose S1 surfaceN=%s shapeN=%s",
                    len(surfs),
                    len(donors),
                )
    except Exception as e:
        LOG.warning("compose S1 slot lang skip: %s", e)
    if (
        piece_lang_block
        or scale_lock
        or piece_lock_block
        or plating_block
        or shape_policy_block
        or s1_food_extra
    ):
        food_look_block = (
            str(food_look_block or "")
            + ("\n" if food_look_block else "")
            + s1_food_extra
            + piece_lock_block
            + shape_policy_block
            + plating_block
            + piece_lang_block
            + scale_lock
        )
    try:
        from b5_generic_plating import apply_core_to_prompt_text, generic_plating_core_enabled

        if generic_plating_core_enabled() and food_look_block:
            food_look_block = apply_core_to_prompt_text(
                str(food_look_block), caller="main"
            )
            LOG.info("compose generic plating core applied")
    except Exception as ge:
        LOG.warning("compose generic plating core skip: %s", ge)
    (meta / "food_look.json").write_text(
        json.dumps(food_looks, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (meta / "piece_lang_block.txt").write_text(
        piece_lang_block or "(none)", encoding="utf-8"
    )
    (meta / "piece_set_lock.txt").write_text(
        piece_lock_block or "(none)", encoding="utf-8"
    )
    (meta / "identity_plating_block.txt").write_text(
        plating_block or "(none)", encoding="utf-8"
    )
    LOG.info(
        "piece_lang chars=%s plating chars=%s piece_set_max=%s scale_lock=%s truth_gates=%s",
        len(piece_lang_block),
        len(plating_block),
        piece_set_n,
        bool(scale_lock),
        len(piece_truth_gates),
    )
    wrap_looks = read_wrap_looks_for_paths(wrap_items) if wrap_items else []
    wrap_look_block = wrap_look_block_ja(wrap_looks)
    (meta / "wrap_look.json").write_text(
        json.dumps(wrap_looks, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    photo_modes = resolve_photo_modes(
        photo_pass if photo_pass is not None else os.environ.get("B5_PHOTO_PASS")
    )
    LOG.info(
        "photo_pass modes=%s catch_ocr=%s lookN=%s wrapLookN=%s",
        photo_modes,
        catch_ocr,
        len(food_looks),
        len(wrap_looks),
    )
    use_stamp = stamp_layers_enabled(stamp_layers)
    first_food = next(iter(food_by_theme.values()), {}) if food_by_theme else {}
    ident0 = first_food.get("identity")
    stamp_plan = collect_stamp_assets(
        parent_bag=truth,
        mix_parts=mix_list if len(mix_list) >= 2 else [],
        identity=Path(ident0) if ident0 and ident0 is not True and Path(str(ident0)).is_file() else None,
    )
    (meta / "stamp_plan.json").write_text(
        json.dumps(
            {
                "enabled": use_stamp,
                "bagMode": stamp_plan.get("bagMode"),
                "bags": [str(p) for p in stamp_plan.get("bags") or []],
                "foods": [str(p) for p in stamp_plan.get("foods") or []],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    LOG.info("stamp_layers=%s plan=%s", use_stamp, stamp_plan.get("bagMode"))

    proposals_path = out_root / "b5_proposals.json"
    mix_meta_parts = [
        {
            "parentSku": str(p.get("parentSku") or ""),
            "nameBase": str(p.get("nameBase") or p.get("name") or ""),
            "jan": str(p.get("jan") or ""),
            "packageTruth": str(p.get("packageTruth") or p.get("bag") or ""),
        }
        for p in mix_list
        if isinstance(p, dict) and str(p.get("nameBase") or p.get("name") or "").strip()
    ]
    name_base_persist = str(name_base or "").strip()
    if not name_base_persist and parent_sku:
        try:
            from b5_parent_shared import load_parent_shared

            row = (load_parent_shared(work).get("byParentSku") or {}).get(
                str(parent_sku).strip()
            ) or {}
            name_base_persist = str(row.get("nameBase") or "").strip()
        except Exception:
            pass
    # 再生成で既存フォルダへ書くときは proposals を上書きしない（元の4案を残す）
    if reuse_dir is None or not proposals_path.is_file():
        payload = {
            "role": ROLE,
            "jan": jan_d,
            "parentSku": parent_sku or "",
            "shortName": short,
            "composeDir": dir_name,
            "textColor": str(color.get("id") or text_color),
            "packageTruth": str(truth),
            "nameBase": name_base_persist,
            "comboParts": mix_meta_parts,
            "foodRefs": food_snapshot,
            "error": mat.get("error") or "",
            "materials": materials,
            "proposals": props,
            "export02": False,
            "note": "P2 compose。個品見本は theme_new_run（§7.14）。【02】は P3 export。",
        }
        proposals_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    else:
        # 再利用時も nameBase／部品だけ補完（再作成が単体 theme_new_run に落ちないように）
        try:
            existing = json.loads(proposals_path.read_text(encoding="utf-8-sig"))
            if not isinstance(existing, dict):
                existing = {}
            changed = False
            if name_base_persist and not str(existing.get("nameBase") or "").strip():
                existing["nameBase"] = name_base_persist
                changed = True
            if mix_meta_parts and not (
                isinstance(existing.get("comboParts"), list)
                and existing.get("comboParts")
            ):
                existing["comboParts"] = mix_meta_parts
                changed = True
            if changed:
                proposals_path.write_text(
                    json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                LOG.info(
                    "proposals meta mix upsert nameBase=%s parts=%s",
                    name_base_persist[:40],
                    len(mix_meta_parts),
                )
        except Exception as e:
            LOG.warning("proposals mix meta upsert fail: %s", e)

    if dry_run:
        for spec in run_sets:
            tid = str(spec.get("themeId") or "")
            refs = food_by_theme.get(tid) or {}
            dry_head = str(spec.get("catchBurn") or "").strip()
            dry_sales = str(spec.get("salesBurn") or spec.get("salesTalk") or "").strip()
            if spec.get("burnSkip"):
                dry_head = ""
            elif not dry_head:
                dry_head = str(spec.get("catch") or "").strip()
            prompt = _prompt_p3_free(
                facts_line,
                theme_id=tid,
                name_ja=str(spec.get("nameJa") or ""),
                scene=str(spec.get("scene") or ""),
                catch_head=dry_head,
                sales_given=dry_sales,
                text_color=str(color.get("id") or text_color),
                has_food_unit_canon=bool(refs.get("identity")),
                has_food_unit_sizzle=bool(refs.get("sizzle")),
                has_food_unit_secondary=bool(refs.get("secondary")),
                has_food_unit_wrap=bool(refs.get("wrap")),
                has_food_unit_piece_set=bool(piece_set_n > 0),
                food_mode=str(refs.get("mode") or "package"),
                layout_reference=bool((props_by_tid.get(tid) or {}).get("_regenTextOnly")),
                food_tune_note=str(
                    (props_by_tid.get(tid) or {}).get("_foodTuneNote") or ""
                ),
                mix_parts=list(refs.get("partNames") or []),
                mix_part_bags=bool(refs.get("hasPartBags")),
                stamp_layers=use_stamp,
                food_look_block=food_look_block,
                wrap_look_block=wrap_look_block,
                size_arm_block=_size_arm_block(size_arm),
                photo_arm=photo_arm,
                dish_part_name=str(
                    (props_by_tid.get(tid) or {}).get("dishPartName")
                    or spec.get("dishPartName")
                    or ""
                ),
            )
            tid_safe = _safe_name(str(spec.get("themeId") or spec["id"]))
            (meta / f"{tid_safe}_prompt.txt").write_text(prompt, encoding="utf-8")
        summary = {
            "ok": True,
            "dryRun": True,
            "role": ROLE,
            "jan": jan_d,
            "parentSku": parent_sku or "",
            "shortName": short,
            "composeDir": dir_name,
            "outRoot": str(out_root),
            "pickDir": str(pick_dir),
            "n": len(run_sets),
            "foodRefs": food_snapshot,
            "stampLayers": use_stamp,
            "export02": False,
        }
        (out_root / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        LOG.info("dry-run out=%s n=%s", out_root, len(run_sets))
        return summary

    octas_path = resolve_octas(work)
    if octas_path is None or not octas_path.is_file():
        raise SystemExit("99.octas期限管理シール素材 なし")
    octas_src = Image.open(octas_path).convert("RGBA")
    gold = _load_gold_disk(work, _gold_params()[2])
    gold.save(meta / "gold_disk.png")

    typo_refs = _oshare_typo_paths(6)
    typo_sheet: Optional[Path] = None
    if typo_refs:
        typo_sheet = _write_style_sheet(typo_refs, meta / "style_typo_sheet.jpg")
    else:
        LOG.warning("IMAGE_STYLE_TYPO なしで続行")

    openai_model = resolve_openai_image_model()
    LOG.info(
        "compose role=%s jan=%s short=%s n=%s model=%s quality=%s out=%s",
        ROLE,
        jan_d,
        short,
        len(run_sets),
        openai_model,
        openai_quality,
        out_root,
    )

    results: List[Dict[str, Any]] = []
    used_sales: set = set()
    face_block = face_cm_prompt_block(package_face_h_cm, package_face_w_cm)
    for theme_i, spec in enumerate(run_sets):
        sid = str(spec["id"])
        theme_id = str(spec.get("themeId") or sid)
        catch = str(spec.get("catch") or "")
        # 正: 大＝確認済み catchBurn、小＝salesBurn（無ければ salesTalk）。無いときだけ catch から分割
        burn = str(spec.get("catchBurn") or "").strip()
        sales_burn = str(spec.get("salesBurn") or spec.get("salesTalk") or "").strip()
        if burn and sales_burn:
            catch_image, sales_talk = burn, sales_burn
            used_sales.add(sales_talk)
        else:
            catch_image, sales_talk = catch_sales_pair_ja(catch, used=used_sales)
            if burn:
                catch_image = burn
            if sales_burn:
                sales_talk = sales_burn
        if spec.get("burnSkip"):
            catch_image = ""
            LOG.info("burn skip theme=%s reason=%s", theme_id, spec.get("burnSkipReason") or "")
        tid_safe = _safe_name(theme_id)
        pick_name = pick_filename(short, theme_id, catch)
        sdir = out_root / tid_safe
        sdir.mkdir(exist_ok=True)
        row: Dict[str, Any] = {
            "id": sid,
            "themeId": theme_id,
            "themeName": spec.get("nameJa"),
            "catch": catch,
            "catchImage": catch_image,
            "salesTalk": sales_talk,
            "pickName": pick_name,
        }
        try:
            img_paths: List[Path] = []
            img_roles: List[str] = []
            food_refs = food_by_theme.get(theme_id) or {}
            mix_list_theme = food_refs.get("mixParts") or []
            if use_stamp:
                zone_path = write_zone_guide(
                    meta / f"{tid_safe}_layout_zones.jpg",
                    bag_box=N1_BAG_BOX if theme_id == "N1" else BAG_BOX,
                    food_box=N1_FOOD_BOX if theme_id == "N1" else FOOD_BOX,
                )
                img_paths.append(zone_path)
                img_roles.append("IMAGE_LAYOUT_ZONES")
                if typo_sheet is not None:
                    img_paths.append(typo_sheet)
                    img_roles.append("IMAGE_STYLE_TYPO")
                if isinstance(mix_list_theme, list) and len(mix_list_theme) >= 2:
                    food_flags = {
                        "mix": True,
                        "partNames": list(food_refs.get("partNames") or []),
                        "hasIdentity": bool(stamp_plan.get("foods")),
                        "hasSizzle": False,
                        "hasWrap": False,
                        "hasSecondary": False,
                        "hasPartBags": bool(food_refs.get("hasPartBags")),
                        "foodMode": "identity",
                        "stampLayers": True,
                    }
                else:
                    food_flags = {
                        "hasIdentity": bool(stamp_plan.get("foods")),
                        "hasSizzle": False,
                        "hasWrap": False,
                        "hasSecondary": False,
                        "foodMode": str(food_refs.get("mode") or "package"),
                        "stampLayers": True,
                    }
            else:
                img_paths = [truth]
                img_roles = [
                    "IMAGE_PACKAGE_SCALE"
                    if isinstance(mix_list_theme, list) and len(mix_list_theme) >= 2
                    else "IMAGE_PACKAGE_TRUTH"
                ]
                if typo_sheet is not None:
                    img_paths.append(typo_sheet)
                    img_roles.append("IMAGE_STYLE_TYPO")
                if isinstance(mix_list_theme, list) and len(mix_list_theme) >= 2:
                    food_flags = _attach_mix_parts(
                        img_paths,
                        img_roles,
                        mix_list_theme,
                        include_secondary=include_food_secondary,
                    )
                else:
                    food_flags = _attach_food_refs(img_paths, img_roles, food_refs)
                # 一片セット同送（G2／③）。stamp 時も参照に残す（描き直しロック文と併用）
                piece_jans: List[str] = []
                if isinstance(mix_list_theme, list) and len(mix_list_theme) >= 2:
                    for raw in mix_list_theme:
                        if not isinstance(raw, dict):
                            continue
                        jd = "".join(
                            ch for ch in str(raw.get("jan") or "") if ch.isdigit()
                        )
                        if jd and jd not in piece_jans:
                            piece_jans.append(jd)
                if not piece_jans and jan_d:
                    piece_jans = [jan_d]
                piece_flags = _attach_piece_set_refs(
                    img_paths,
                    img_roles,
                    jans=piece_jans,
                    max_n=piece_set_n,
                )
                if isinstance(food_flags, dict):
                    food_flags["pieceSet"] = piece_flags
            row["food"] = food_flags
            row["stampLayers"] = use_stamp
            orig_prop = props_by_tid.get(theme_id) or {}
            regen_ref = str(orig_prop.get("_regenReference") or "").strip()
            if regen_ref and Path(regen_ref).is_file():
                img_paths.append(Path(regen_ref))
                img_roles.append("IMAGE_LAYOUT_REFERENCE")
                row["layoutReference"] = Path(regen_ref).name
            if size_arm == "part_first":
                img_paths, img_roles = _move_parts_first(img_paths, img_roles)
                LOG.info("part_first roles=%s", img_roles)
            if not include_food_secondary:
                LOG.info("no food secondary roles=%s", img_roles)

            if reuse_dir is not None:
                try_theme_key = (
                    f"{theme_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                )
            else:
                try_theme_key = theme_id
            aspect_rows: List[Dict[str, Any]] = []
            if not use_stamp:
                scale_idx = next(
                    (
                        i
                        for i, r in enumerate(img_roles)
                        if str(r) == "IMAGE_PACKAGE_SCALE"
                    ),
                    -1,
                )
                part_pairs = [
                    (str(r), Path(p))
                    for p, r in zip(img_paths, img_roles)
                    if str(r).startswith("IMAGE_PACKAGE_PART_")
                ]
                if scale_idx >= 0 and len(part_pairs) >= 2:
                    aspect_rows = measure_mix_aspects_from_scale(
                        Path(img_paths[scale_idx]),
                        [r for r, _ in part_pairs],
                        out_dir=meta,
                    )
                    if aspect_rows:
                        (meta / "bag_aspect_scale.json").write_text(
                            json.dumps(aspect_rows, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                if not aspect_rows:
                    aspect_rows = measure_role_aspects(img_paths, img_roles)
            color_rows = [] if use_stamp else measure_role_colors(img_paths, img_roles)
            if color_rows:
                (meta / "bag_color.json").write_text(
                    json.dumps(color_rows, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
            if size_arm == "ratio_card" and aspect_rows:
                for card_path, card_role in _write_aspect_cards(meta, aspect_rows):
                    img_paths.append(card_path)
                    img_roles.append(card_role)
                LOG.info("ratio_card roles=%s", [r for r in img_roles if "ASPECT_CARD" in r])
            aspect_block = (
                aspect_prompt_block(aspect_rows)
                + color_prompt_block(color_rows)
                + face_block
            )
            grid_budget = compute_grid_budget(catch_image, sales_talk)
            budget_prompt, adopted_meta = build_compose_grid_prompt_block(
                catch_image,
                sales_talk,
                theme_index=theme_i,
                n_themes=len(run_sets),
                materials=materials,
            )
            if adopted_meta:
                row["adoptedZone"] = adopted_meta
            guide_pattern = None
            if adopted_meta.get("id"):
                guide_pattern = pattern_by_id(str(adopted_meta["id"]))
            writing_mode = str(adopted_meta.get("writingMode") or "")
            guide_cells = adopted_meta.get("cells") if adopted_meta else None
            retry_note = ""
            tries: List[Dict[str, Any]] = []
            try_archive: List[Dict[str, Any]] = []
            best: Optional[Dict[str, Any]] = None
            p3_path = sdir / pick_name
            for attempt in range(1, P3_FREE_MAX_TRIES + 1):
                p3_prompt = _prompt_p3_free(
                    facts_line,
                    theme_id=theme_id,
                    name_ja=str(spec.get("nameJa") or sid),
                    scene=str(spec.get("scene") or ""),
                    retry_note=retry_note,
                    catch_head=catch_image,
                    sales_given=sales_talk,
                    text_color=str(color.get("id") or text_color),
                    has_food_unit_canon=bool(food_flags.get("hasIdentity")),
                    has_food_unit_sizzle=bool(food_flags.get("hasSizzle")),
                    has_food_unit_secondary=bool(food_flags.get("hasSecondary")),
                    has_food_unit_wrap=bool(food_flags.get("hasWrap")),
                    has_food_unit_piece_set=bool(piece_set_n > 0),
                    food_mode=str(food_flags.get("foodMode") or "package"),
                    layout_reference=bool(orig_prop.get("_regenTextOnly")),
                    food_tune_note=str(orig_prop.get("_foodTuneNote") or ""),
                    grid_budget_block=budget_prompt,
                    mix_parts=list(food_flags.get("partNames") or []),
                    mix_part_bags=bool(food_flags.get("hasPartBags")),
                    stamp_layers=use_stamp,
                    food_look_block=food_look_block,
                    wrap_look_block=wrap_look_block,
                    bag_aspect_block=aspect_block,
                    size_arm_block=_size_arm_block(size_arm),
                    photo_arm=photo_arm,
                    dish_part_name=str(
                        orig_prop.get("dishPartName") or spec.get("dishPartName") or ""
                    ),
                )
                (meta / f"{tid_safe}_try{attempt}_prompt.txt").write_text(
                    p3_prompt, encoding="utf-8"
                )
                raw3, g3 = _gen_ai(
                    "openai",
                    prompt=p3_prompt,
                    image_paths=img_paths,
                    image_roles=img_roles,
                    openai_model=openai_model,
                    openai_quality=openai_quality,
                )
                cand_im = _to_1200(raw3)
                if size_arm in ("paste_only", "paste_shadow", "paste_relight") and aspect_rows:
                    cand_im = _paste_parts_with_shadow(
                        cand_im, aspect_rows, shadow=(size_arm == "paste_shadow")
                    )
                if size_arm == "paste_relight" and aspect_rows:
                    relight_in = meta / f"{tid_safe}_try{attempt}_relight_in.jpg"
                    cand_im.save(relight_in, quality=92)
                    raw_r, _g_r = _gen_ai(
                        "openai",
                        prompt=(
                            "包材の明るさと影の方向だけ、周りの食品と場面に合わせろ。"
                            "形・縦横・印字・位置は変えるな。食品と背景は維持。"
                        ),
                        image_paths=[relight_in],
                        image_roles=["IMAGE_SCENE"],
                        openai_model=openai_model,
                        openai_quality=openai_quality,
                    )
                    cand_im = _paste_parts_with_shadow(
                        _to_1200(raw_r), aspect_rows, shadow=False
                    )
                if use_stamp:
                    cand_im, stamp_meta = stamp_truth_layers(
                        cand_im,
                        bag_paths=list(stamp_plan.get("bags") or []),
                        food_paths=list(stamp_plan.get("foods") or []),
                        theme_id=theme_id,
                        bag_mode=str(stamp_plan.get("bagMode") or "single"),
                    )
                    dump_stamp_meta(
                        meta / f"{tid_safe}_try{attempt}_stamp.json", stamp_meta
                    )
                try_path = sdir / f"{tid_safe}_try{attempt}.jpg"
                cand_im.convert("RGB").save(try_path, quality=92)
                refine_meta: Dict[str, Any] = {"ok": True, "skipped": True}
                if wrap_items and not use_stamp:
                    try:
                        from b4_wrap_scene_refine import (
                            refine_scene_wrap_to_file,
                            wrap_scene_refine_enabled,
                        )

                        if wrap_scene_refine_enabled():
                            w0 = Path(str(wrap_items[0].get("path") or ""))
                            if w0.is_file():
                                rpath = meta / f"{tid_safe}_try{attempt}_wrap_refine.jpg"
                                refine_meta = refine_scene_wrap_to_file(
                                    try_path,
                                    w0,
                                    rpath,
                                    looks=wrap_looks,
                                    work_dir=meta
                                    / f"{tid_safe}_try{attempt}_wrap_refine_work",
                                )
                                if (
                                    refine_meta.get("ok")
                                    and not refine_meta.get("skipped")
                                    and rpath.is_file()
                                ):
                                    shutil.copy2(rpath, try_path)
                                    cand_im = Image.open(try_path).convert("RGB")
                                (meta / f"{tid_safe}_try{attempt}_wrap_refine.json").write_text(
                                    json.dumps(refine_meta, ensure_ascii=False, indent=2),
                                    encoding="utf-8",
                                )
                    except Exception as e:
                        LOG.warning("wrap scene refine skip: %s", e)
                        refine_meta = {"ok": False, "skipped": True, "error": str(e)[:200]}
                dual = measure_grid_dual(
                    try_path,
                    catch_head=catch_image,
                    sales_talk=sales_talk,
                    gemini=grid_measurer,
                    writing_mode=writing_mode,
                    guide_pattern=guide_pattern,
                    guide_cells=guide_cells,
                )
                grid_try = dual["local"]
                _write_grid20_p3(
                    try_path, grid_dir / f"{tid_safe}_try{attempt}_grid20.jpg"
                )
                hit_n = int(dual.get("gateHitN") or 0)
                local_hit = int(dual.get("localHitN") or 0)
                gemini_hit = dual.get("geminiHitN")
                gate_src = str(dual.get("gateSource") or "local")
                tier = classify_grid_tier(hit_n)
                needs_retry = grid_needs_retry(hit_n)
                no_catch = hit_n < 1
                checker = _checkerboard_suspected(cand_im)
                ocr = {"ok": True, "skipped": True}
                if catch_ocr:
                    ocr = check_catch_text(
                        try_path, catch_head=catch_image, sales_talk=sales_talk
                    )
                    (meta / f"{tid_safe}_try{attempt}_catch_ocr.json").write_text(
                        json.dumps(ocr, ensure_ascii=False, indent=2), encoding="utf-8"
                    )
                ocr_fail = bool(catch_ocr and not ocr.get("skipped") and not ocr.get("ok"))
                aspect = {"ok": True, "skipped": True}
                if aspect_rows and not use_stamp:
                    aspect = check_drawn_bags(try_path, aspect_rows)
                    (meta / f"{tid_safe}_try{attempt}_bag_aspect.json").write_text(
                        json.dumps(aspect, ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
                    (meta / f"{tid_safe}_try{attempt}_袋判定.txt").write_text(
                        format_bag_verdict_text(
                            aspect,
                            theme_id=str(theme_id),
                            try_name=f"try{attempt}",
                        ),
                        encoding="utf-8",
                    )
                aspect_fail = bool(
                    aspect_rows and not aspect.get("skipped") and not aspect.get("ok")
                )
                aspect_rel = aspect_worst_rel(aspect) if aspect_rows else None
                face = {"ok": True, "skipped": True}
                if not use_stamp:
                    face = check_face_vs_tableware(
                        try_path,
                        height_cm=package_face_h_cm,
                        width_cm=package_face_w_cm,
                    )
                    (meta / f"{tid_safe}_try{attempt}_face_scale.json").write_text(
                        json.dumps(face, ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
                face_fail = bool(not face.get("skipped") and not face.get("ok"))
                food_piece = {"ok": True, "skipped": True}
                if look_items and not use_stamp:
                    food_piece = check_drawn_foods(try_path, look_items)
                    (meta / f"{tid_safe}_try{attempt}_food_piece.json").write_text(
                        json.dumps(food_piece, ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
                food_fail = bool(
                    look_items and not food_piece.get("skipped") and not food_piece.get("ok")
                )
                wrap_chk = {"ok": True, "skipped": True}
                if wrap_items and not use_stamp:
                    wrap_chk = check_drawn_wraps(try_path, wrap_items, wrap_looks)
                    (meta / f"{tid_safe}_try{attempt}_wrap_place.json").write_text(
                        json.dumps(wrap_chk, ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
                wrap_fail = bool(
                    wrap_items and not wrap_chk.get("skipped") and not wrap_chk.get("ok")
                )
                rank = (
                    (1 if checker else 0) * 1000
                    + (1 if ocr_fail else 0) * 400
                    + (1 if aspect_fail else 0) * 350
                    + (1 if face_fail else 0) * 340
                    + (1 if wrap_fail else 0) * 330
                    + (1 if food_fail else 0) * 320
                    + grid_rank_penalty(hit_n)
                    + (1 if no_catch else 0) * 200
                    + hit_n
                )
                tries.append(
                    {
                        "attempt": attempt,
                        "hitN": hit_n,
                        "gridTier": tier,
                        "localHitN": local_hit,
                        "geminiHitN": gemini_hit,
                        "gateSource": gate_src,
                        "gateAlgo": dual.get("gateAlgo"),
                        "ok20": hit_n <= P3_TEXT_MAX_CELLS,
                        "buffer21_24": P3_TEXT_MAX_CELLS < hit_n <= 24,
                        "needsRetry": needs_retry,
                        "noCatch": no_catch,
                        "checker": checker,
                        "catchOcrOk": ocr.get("ok"),
                        "catchOcrSkipped": ocr.get("skipped"),
                        "ocrCatch": ocr.get("ocrCatch"),
                        "bagAspectOk": aspect.get("ok"),
                        "bagAspectSkipped": aspect.get("skipped"),
                        "bagAspectRel": aspect_rel,
                        "faceScaleOk": face.get("ok"),
                        "faceScaleSkipped": face.get("skipped"),
                        "foodPieceOk": food_piece.get("ok"),
                        "foodPieceSkipped": food_piece.get("skipped"),
                        "wrapPlaceOk": wrap_chk.get("ok"),
                        "wrapPlaceSkipped": wrap_chk.get("skipped"),
                        "wrapAirOk": wrap_chk.get("airOk"),
                        "wrapFrontOk": wrap_chk.get("frontOk"),
                        "rank": rank,
                        "api": g3.get("apiPath"),
                        "file": try_path.name,
                    }
                )
                try_archive.append(
                    archive_try_with_overlays(
                        try_path=try_path,
                        tries_dir=tries_dir,
                        theme_id=try_theme_key,
                        attempt=attempt,
                        catch_head=catch_image,
                        sales_talk=sales_talk,
                        dual=dual,
                        picked=False,
                    )
                )
                LOG.info(
                    "%s try=%s gate=%s(%s) tier=%s local=%s gemini=%s "
                    "needsRetry=%s noCatch=%s checker=%s ocrFail=%s aspectFail=%s aspectRel=%s faceFail=%s wrapFail=%s foodFail=%s",
                    theme_id,
                    attempt,
                    hit_n,
                    gate_src,
                    tier,
                    local_hit,
                    gemini_hit,
                    needs_retry,
                    no_catch,
                    checker,
                    ocr_fail,
                    aspect_fail,
                    aspect_rel,
                    face_fail,
                    wrap_fail,
                    food_fail,
                )
                keep_key = _keep_try_key(
                    ocr_fail=ocr_fail,
                    aspect_fail=aspect_fail,
                    aspect_rel=aspect_rel,
                    rank=rank,
                    attempt=attempt,
                    food_fail=food_fail,
                    face_fail=face_fail,
                    wrap_fail=wrap_fail,
                )
                if best is None or keep_key < best["keepKey"]:
                    best = {
                        "im": cand_im,
                        "grid": grid_try,
                        "dual": dual,
                        "api": g3.get("apiPath"),
                        "model": g3.get("model"),
                        "hitN": hit_n,
                        "localHitN": local_hit,
                        "geminiHitN": gemini_hit,
                        "gateSource": gate_src,
                        "rank": rank,
                        "keepKey": keep_key,
                        "file": try_path.name,
                        "bagAspect": aspect,
                        "bagAspectFail": aspect_fail,
                        "attempt": attempt,
                    }
                if (
                    not needs_retry
                    and not checker
                    and not no_catch
                    and not ocr_fail
                    and not aspect_fail
                    and not face_fail
                    and not wrap_fail
                    and not food_fail
                ):
                    break
                hero_half = max(1, int(_gold_params()[2]) // 2)
                retry_bits: List[str] = []
                if aspect_fail and aspect.get("retryNote"):
                    retry_bits.append(str(aspect.get("retryNote")))
                if face_fail and face.get("retryNote"):
                    retry_bits.append(str(face.get("retryNote")))
                if wrap_fail and wrap_chk.get("retryNote"):
                    retry_bits.append(str(wrap_chk.get("retryNote")))
                if food_fail and food_piece.get("retryNote"):
                    retry_bits.append(str(food_piece.get("retryNote")))
                if ocr_fail and ocr.get("retryNote"):
                    retry_bits.append(str(ocr.get("retryNote")))
                elif no_catch:
                    retry_bits.append(
                        "前回は後載せ文字が無かった。指定のキャッチ（大）を一字も落とさず載せろ。"
                        "セールストーク（小）があれば、それも一字も落とすな。"
                        "0本は不可。上部・左寄り。"
                        f"キャッチのキーワード1語以上は文字高さ約{hero_half}px以上。"
                        "右上金丸聖域・右下Octas聖域に文字を入れるな。"
                        + (
                            "袋は描くな。後載せする。"
                            if use_stamp
                            else "IMAGE_PACKAGE_TRUTHの袋は必ず残せ。"
                        )
                    )
                elif needs_retry or checker:
                    grid_note, _ = build_retry_note_for_hit(
                        hit_n, checker=checker, hero_half=hero_half
                    )
                    retry_bits.append(grid_note)
                if retry_bits:
                    retry_note = "\n".join(retry_bits)

            if best is None:
                raise RuntimeError("generate empty")

            picked_file = str(best.get("file") or "")
            for ta in try_archive:
                ta["picked"] = picked_file.endswith(f"try{ta.get('attempt')}.jpg")

            pre_gold = best["im"].convert("RGBA")
            pre_path = meta / f"{tid_safe}_pre_gold.jpg"
            pre_gold.convert("RGB").save(pre_path, quality=92)
            pass_rows: List[Dict[str, Any]] = []
            first_pick = ""
            first_file = ""
            g3s = {}
            for mi, pmode in enumerate(photo_modes):
                applied = apply_photo_pass(
                    pre_gold,
                    mode=pmode,
                    package_truth=truth,
                    openai_quality=openai_quality,
                    work_tmp=meta,
                )
                p3 = _finish_p3_overlays(applied["im"], gold, octas_src)
                if p3.size != (CANVAS, CANVAS):
                    p3 = p3.resize((CANVAS, CANVAS), Image.Resampling.LANCZOS)
                stem = Path(pick_name).stem
                ext = Path(pick_name).suffix or ".jpg"
                if len(photo_modes) > 1:
                    out_name = f"{stem}_{MODE_JA.get(pmode, pmode)}{ext}"
                else:
                    out_name = pick_name
                if reuse_dir is not None and mi == 0:
                    out_name = f"__regen_staging_{out_name}"
                dest = (sdir / out_name) if mi == 0 else (pick_dir / out_name)
                if mi == 0:
                    p3.convert("RGB").save(p3_path, quality=92)
                    g3s = _write_grid20_p3(p3_path, grid_dir / f"{tid_safe}_grid20.jpg")
                    final_pick = pick_dir / out_name
                    shutil.copy2(p3_path, final_pick)
                    first_pick = str(final_pick)
                    first_file = out_name
                else:
                    p3.convert("RGB").save(pick_dir / out_name, quality=92)
                    final_pick = pick_dir / out_name
                pass_rows.append(
                    {
                        "mode": pmode,
                        "label": MODE_JA.get(pmode, pmode),
                        "file": out_name,
                        "rel": out_name,
                        "skipped": applied.get("skipped"),
                        "error": applied.get("error") or "",
                    }
                )
                LOG.info("photo pass %s %s -> %s", theme_id, pmode, final_pick)
            if len(photo_modes) > 1:
                write_compare_html(
                    pick_dir,
                    rows=pass_rows,
                    title=f"写真感 3枚比較 — {short} {theme_id}",
                )
            row.update(
                {
                    "ok": True,
                    "file": first_file or pick_name,
                    "pickFile": first_pick,
                    "pickName": pick_name,
                    "grid20": g3s,
                    "gridBudget": grid_budget,
                    "api": best.get("api"),
                    "model": best.get("model"),
                    "tries": tries,
                    "tryArchive": try_archive,
                    "pickedTry": best.get("file"),
                    "gateSource": best.get("gateSource"),
                    "localHitN": best.get("localHitN"),
                    "geminiHitN": best.get("geminiHitN"),
                    "gridTier": classify_grid_tier(int(best.get("hitN") or 0)),
                    "gold": "pillow_stamp",
                    "octas": "pillow_br_1p3cell",
                    "photoPass": pass_rows,
                    "foodLookN": len(food_looks),
                    "bagAspect": best.get("bagAspect"),
                    "bagAspectOk": (
                        None
                        if not isinstance(best.get("bagAspect"), dict)
                        else (
                            None
                            if best["bagAspect"].get("skipped")
                            else bool(best["bagAspect"].get("ok"))
                        )
                    ),
                    "bagAspectTry": best.get("attempt"),
                }
            )
            LOG.info("ok %s -> %s", theme_id, first_pick)
        except Exception as e:
            LOG.exception("fail theme=%s", theme_id)
            row.update({"ok": False, "error": str(e)})
        results.append(row)

    gemini_grid_cost = float(grid_measurer.total_cost_jpy) if grid_measurer else 0.0
    tries_html: Optional[str] = None
    ok_rows = [r for r in results if r.get("ok")]
    if ok_rows:
        html_path = build_compose_tries_html(
            out_root=out_root,
            theme_rows=ok_rows,
            title=f"compose try 比較 — {short}",
            gemini_cost_jpy=gemini_grid_cost,
        )
        tries_html = str(html_path)
        LOG.info("tries compare html=%s gemini_cost≈%.2fJPY", html_path, gemini_grid_cost)

    bag_verdict_rows = []
    for r in results:
        if not r.get("ok"):
            continue
        asp = r.get("bagAspect")
        if not isinstance(asp, dict):
            continue
        bag_verdict_rows.append(
            {
                "themeId": r.get("themeId") or r.get("id") or "",
                "tryName": f"try{r.get('bagAspectTry') or '?'}→採用",
                "aspect": asp,
            }
        )
    bag_verdict_txt = ""
    bag_verdict_json = ""
    if bag_verdict_rows:
        vtxt = pick_dir / "袋判定.txt"
        write_bag_verdict_bundle(vtxt, bag_verdict_rows)
        bag_verdict_txt = str(vtxt)
        meta_root = out_root / "_meta"
        meta_root.mkdir(parents=True, exist_ok=True)
        vjson = meta_root / "袋判定.json"
        vjson.write_text(
            json.dumps(bag_verdict_rows, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        bag_verdict_json = str(vjson)
        LOG.info("bag verdict written %s", vtxt)
        try:
            print(vtxt.read_text(encoding="utf-8"))
        except Exception:
            pass

    ok_n = sum(1 for r in results if r.get("ok"))
    summary = {
        "ok": ok_n > 0,
        "dryRun": False,
        "role": ROLE,
        "jan": jan_d,
        "parentSku": parent_sku or "",
        "shortName": short,
        "composeDir": dir_name,
        "textColor": str(color.get("id") or text_color),
        "packageTruth": str(truth),
        "openaiModel": openai_model,
        "openaiQuality": openai_quality,
        "gridGate": (os.environ.get("B5_GRID_GATE") or "local").strip().lower(),
        "gridTierPolicy": (
            (os.environ.get("B5_GRID_TIER_POLICY") or "tiered").strip()
        ),
        "gridRetryAt": retry_threshold_cells(),
        "geminiGridModel": gemini_grid_model if use_grid_gemini else "",
        "geminiGridCostJpy": round(gemini_grid_cost, 3),
        "triesCompareHtml": tries_html or "",
        "bagVerdictTxt": bag_verdict_txt,
        "bagVerdictJson": bag_verdict_json,
        "export02": False,
        "outRoot": str(out_root),
        "pickDir": str(pick_dir),
        "okN": ok_n,
        "failN": len(results) - ok_n,
        "foodRefs": food_snapshot,
        "photoPassModes": photo_modes,
        "sizeArm": size_arm,
        "foodLookN": len(food_looks),
        "stampLayers": use_stamp,
        "stampPlan": {
            "bagMode": stamp_plan.get("bagMode"),
            "bagN": len(stamp_plan.get("bags") or []),
            "foodN": len(stamp_plan.get("foods") or []),
        },
        "results": results,
        "factsError": mat.get("error") or "",
        "note": "P2。比較は商品名省略＋themeId＋キャッチ。02 は b5_main_base_export.py。",
    }
    (out_root / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return summary


def _num_or_none(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        n = float(v)
    except (TypeError, ValueError):
        return None
    return n if n > 0 else None


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="B-⑤ compose --role rakuten_main_base")
    ap.add_argument(
        "--role",
        default=ROLE,
        help=f"既定={ROLE}（他 role は拒否）",
    )
    ap.add_argument("--from-json", type=Path, required=True, help="人が選んだ案の JSON")
    ap.add_argument("--work-root", type=Path, default=None)
    ap.add_argument(
        "--openai-quality",
        default=OPENAI_QUALITY_DEFAULT,
        choices=("low", "medium", "high"),
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="APIを呼ばず prompt と JSON だけ書く",
    )
    ap.add_argument(
        "--only-theme",
        action="append",
        default=[],
        help="再生成する themeId のみ（複数可）。例: --only-theme T4",
    )
    ap.add_argument(
        "--skip-grid-gemini",
        action="store_true",
        help="20マス判定を catch_v7 のみにする（Gemini API 不使用）",
    )
    ap.add_argument(
        "--gemini-grid-model",
        default=DEFAULT_GEMINI_GRID_MODEL,
        help=f"20マス Gemini 判定モデル（既定 {DEFAULT_GEMINI_GRID_MODEL}）",
    )
    ap.add_argument(
        "--no-stamp-layers",
        action="store_true",
        help="袋・個品をAI描き直しに戻す（切戻し。既定は正本画素の後載せ）",
    )
    ap.add_argument(
        "--photo-pass",
        default="none",
        choices=("none", "package", "whole", "compare"),
        help="金丸前の写真感。compare=なし/袋だけ/全体",
    )
    ap.add_argument(
        "--no-catch-ocr",
        action="store_true",
        help="見出し文字検査をしない",
    )
    ap.add_argument(
        "--size-arm",
        default="",
        choices=(
            "",
            "scene_only",
            "front",
            "ref_object",
            "part_first",
            "ratio_card",
            "paste_shadow",
            "paste_only",
            "paste_relight",
        ),
        help="サイズ試験の追加文を1つだけ足す。空なら足さない",
    )
    ap.add_argument(
        "--no-food-secondary",
        action="store_true",
        help="個品の副見本画像を渡さない。主見本は渡す",
    )
    ap.add_argument(
        "--piece-set-max",
        type=int,
        default=0,
        help="一片セット添付枚数。0=オフ（既定）／3=先頭3／5=あるだけ。IDENTITYは残す",
    )
    ap.add_argument(
        "--photo-arm",
        default="",
        choices=("", "drop_dof", "mm85"),
        help="写真の指定を1つだけ変える。空なら今の50mm",
    )
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    if str(args.role or "").strip() != ROLE:
        raise SystemExit(f"--role は {ROLE} のみ（指定={args.role!r}）")

    payload = _load_payload(Path(args.from_json))
    jan = str(payload.get("jan") or "").strip()
    parent_sku = str(payload.get("parentSku") or payload.get("parent_sku") or "").strip()
    truth_s = str(payload.get("packageTruth") or payload.get("package_truth") or "").strip()
    if not truth_s:
        raise SystemExit("JSON に packageTruth が必要です")
    text_color = str(payload.get("textColor") or payload.get("text_color") or "1")
    proposals = payload.get("proposals")
    if args.only_theme:
        want = {str(t).strip() for t in args.only_theme if str(t).strip()}
        if isinstance(proposals, list):
            proposals = [
                p
                for p in proposals
                if isinstance(p, dict) and str(p.get("themeId") or "").strip() in want
            ]
        LOG.info("only-theme filter=%s n=%s", sorted(want), len(proposals or []))
    short_name = str(payload.get("shortName") or payload.get("short_name") or "").strip()

    summary = compose_rakuten_main_base(
        jan=jan,
        parent_sku=parent_sku,
        package_truth=Path(truth_s),
        proposals=list(proposals or []),
        text_color=text_color,
        openai_quality=str(args.openai_quality),
        work_root=args.work_root,
        dry_run=bool(args.dry_run),
        short_name=short_name,
        skip_grid_gemini=bool(args.skip_grid_gemini),
        gemini_grid_model=str(args.gemini_grid_model),
        combo_parts=list(payload.get("comboParts") or payload.get("combo_parts") or []),
        name_base=str(payload.get("nameBase") or payload.get("name_base") or "").strip(),
        stamp_layers=False if args.no_stamp_layers else None,
        photo_pass=str(args.photo_pass),
        catch_ocr=not bool(args.no_catch_ocr),
        package_face_h_cm=_num_or_none(payload.get("packageFaceHcm")),
        package_face_w_cm=_num_or_none(payload.get("packageFaceWcm")),
        size_arm=str(args.size_arm or ""),
        include_food_secondary=not bool(args.no_food_secondary),
        photo_arm=str(args.photo_arm or ""),
        mix_explicit=bool(payload.get("mix")) if "mix" in payload else None,
        food_explicit=bool(payload.get("food")) if "food" in payload else None,
        piece_set_max=int(args.piece_set_max or 0),
    )
    if args.dry_run:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary.get("ok") or summary.get("dryRun") else 1


if __name__ == "__main__":
    raise SystemExit(main())
