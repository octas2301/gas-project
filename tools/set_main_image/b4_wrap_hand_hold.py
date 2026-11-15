# -*- coding: utf-8 -*-
"""P3-W 第2: WRAP確定後の手持ち候補生成・採用。

見本ゲートから呼ぶ。採用1枚を MAIN左下／サブ個包装訴求／場面参照へ共有する正。
切戻し: B4_WRAP_ASSET_LANE=0（b4_wrap_asset_rules）。
"""
from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from b4_wrap_asset_rules import (
    apply_selection,
    build_hand_hold_prompt_ja,
    candidate_manifest,
    resolve_wrap_asset_for_consumers,
    should_branch_after_wrap_confirm,
    wrap_asset_lane_enabled,
)
from food_unit_canon import existing_wrap_path

LOG = logging.getLogger("set_main_image.b4_wrap_hand_hold")

OUT_DIR_NAME = "wrap_hand_hold"
MANIFEST_NAME = "wrap_hand_hold_manifest.json"
SELECTED_NAME = "WRAP_HAND_HOLD.jpg"
DEFAULT_N = 5
DEFAULT_QUALITY = "low"
DEFAULT_SIZE = "1024x1024"

# 構図のばらけ（案件名を焼かない）。実寸・背景は共通プロンプト側。
_VARIANT_LINES_JA = (
    "構図A: 親指と人差し指で個包装を1つつまむ。寄り。",
    "構図B: 指先に個包装を載せる。寄り。",
    "構図C: 丸枠向きの寄り。個包装が枠の中央付近。",
    "構図D: 両手の指先でそっと差し出す。寄り。",
    "構図E: さらにマクロ寄りのつまみ。指先と個包装だけ。",
)


def out_dir_for(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / OUT_DIR_NAME


def manifest_path(theme_new_dir: Path) -> Path:
    return out_dir_for(theme_new_dir) / MANIFEST_NAME


def selected_path(theme_new_dir: Path) -> Optional[Path]:
    """採用済みがあればそのパス。"""
    p = out_dir_for(theme_new_dir) / SELECTED_NAME
    if p.is_file():
        return p
    refs = Path(theme_new_dir) / "01_refs" / SELECTED_NAME
    if refs.is_file():
        return refs
    return None


def load_manifest(theme_new_dir: Path) -> Dict[str, Any]:
    p = manifest_path(theme_new_dir)
    if not p.is_file():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8-sig"))
        return data if isinstance(data, dict) else {}
    except Exception as e:
        LOG.warning("load_manifest fail %s", e)
        return {}


def save_manifest(theme_new_dir: Path, doc: Dict[str, Any]) -> Path:
    od = out_dir_for(theme_new_dir)
    od.mkdir(parents=True, exist_ok=True)
    p = manifest_path(theme_new_dir)
    p.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    return p


def needs_hand_hold_lane(theme_new_dir: Path, *, has_wrap: bool) -> bool:
    if not wrap_asset_lane_enabled():
        return False
    wrap_p = existing_wrap_path(theme_new_dir) if has_wrap else None
    return should_branch_after_wrap_confirm(has_wrap=has_wrap, wrap_path=wrap_p)


def list_candidate_files(theme_new_dir: Path) -> List[Path]:
    od = out_dir_for(theme_new_dir)
    if not od.is_dir():
        return []
    return sorted(od.glob("hand_hold_*.jpg"))


def list_single_files(theme_new_dir: Path) -> List[Path]:
    od = out_dir_for(theme_new_dir)
    if not od.is_dir():
        return []
    return sorted(od.glob("wrap_single_*.jpg"))


_SINGLE_LINES_JA = (
    "構図S1: 個包装だけ1つ。正面やや斜め。手なし。",
    "構図S2: 個包装だけ1つ。真上寄り。手なし。",
    "構図S3: 個包装だけ1つ。横向き。手なし。",
)


def _prompt_for_index(i: int) -> str:
    base = build_hand_hold_prompt_ja(circular=True, include_plant=True)
    var = _VARIANT_LINES_JA[i % len(_VARIANT_LINES_JA)]
    return base + "\n" + var


def _prompt_for_single(i: int) -> str:
    lines = [
        "【個包装単体見本】",
        "個包装の真空パックを1つだけ。手は出さない。大袋は出さない。",
        "見た目の正は添付 WRAP 見本のみ。フィルムは透明。あけ口以外の印字・ロゴを足すな。",
        "空気だまりを見本に無く作るな。",
        "背景は白〜生成り。グレー一色禁止。木目は必須にしない。ぼけた植物の緑は少し可。",
        "MAIN／サブの個包装見本として使える実写写真。",
        _SINGLE_LINES_JA[i % len(_SINGLE_LINES_JA)],
    ]
    return "\n".join(lines)


def ensure_wrap_stem_from_name(theme_new_dir: Path, source_name: str) -> Path:
    """候補／手動画像を生成用の種（FOOD_UNIT_WRAP_SEED）へ。WRAPが無ければ初回bootstrap。"""
    from b4_wrap_compose_ref import materialize_wrap_seed

    theme_new_dir = Path(theme_new_dir)
    name = str(source_name or "").strip()
    if not name:
        raise ValueError("個包装見本のファイル名が空です")
    dest = materialize_wrap_seed(theme_new_dir, name, also_bootstrap_wrap=True)
    if dest is None or not Path(dest).is_file():
        raise ValueError(f"個包装見本が見つかりません: {name}")
    return Path(dest)


def recommend_wrap_assets(theme_new_dir: Path) -> Dict[str, str]:
    """既定レ点。保存済み採用を最優先（先頭ファイルに戻さない）。"""
    hands = [p.name for p in list_candidate_files(theme_new_dir)]
    singles = [p.name for p in list_single_files(theme_new_dir)]
    man = load_manifest(theme_new_dir)
    hand = str(man.get("selected") or "") or (hands[0] if hands else "")
    single = (
        str(man.get("selectedSingle") or "").strip()
        or str(man.get("composeWrap") or "").strip()
        or str(man.get("recommendedSingle") or "").strip()
        or (singles[0] if singles else "")
    )
    if single and singles and single not in singles:
        single = singles[0] if singles else ""
    return {"handHold": hand, "wrapSingle": single}


def generate_hand_hold_candidates(
    theme_new_dir: Path,
    *,
    jan: str = "",
    n: int = DEFAULT_N,
    quality: str = DEFAULT_QUALITY,
    size: str = DEFAULT_SIZE,
    force: bool = False,
    wrap_seed_name: str = "",
) -> Dict[str, Any]:
    """
    WRAP見本を参照して手持ち候補を n 枚生成。
    force=False かつ候補が既にあれば再生成しない（課金抑制）。
    """
    theme_new_dir = Path(theme_new_dir)
    if str(wrap_seed_name or "").strip():
        ensure_wrap_stem_from_name(theme_new_dir, wrap_seed_name)
    from b4_wrap_compose_ref import resolve_wrap_generation_ref

    wrap_p = resolve_wrap_generation_ref(theme_new_dir) or existing_wrap_path(theme_new_dir)
    if wrap_p is None or not Path(wrap_p).is_file():
        raise ValueError(
            "個包装見本がありません。コピペ／URLで登録するか WRAP を選んでから生成してください。"
        )
    if not wrap_asset_lane_enabled():
        raise RuntimeError("B4_WRAP_ASSET_LANE=0 のため手持ち生成は停止")

    od = out_dir_for(theme_new_dir)
    od.mkdir(parents=True, exist_ok=True)
    # 試作フォルダ wrap_hand_hold_test があれば本線へ一度だけ寄せる
    test_od = theme_new_dir / "wrap_hand_hold_test"
    if test_od.is_dir() and not list_candidate_files(theme_new_dir):
        import shutil

        for p in sorted(test_od.glob("hand_hold_*.jpg")):
            dest = od / p.name
            if not dest.is_file():
                shutil.copy2(p, dest)
                LOG.info("migrated test hand_hold -> %s", dest.name)
    existing = list_candidate_files(theme_new_dir)
    if existing and not force:
        rels = [p.name for p in existing]
        man = candidate_manifest(
            jan=jan, wrap_path=str(wrap_p), candidates=rels, selected=""
        )
        prev = load_manifest(theme_new_dir)
        if prev.get("selected"):
            man["selected"] = str(prev.get("selected") or "")
        man["reused"] = True
        man["updatedAt"] = datetime.now(timezone.utc).isoformat()
        save_manifest(theme_new_dir, man)
        return {"ok": True, "reused": True, "candidates": rels, "manifest": man}

    from openai_image import generate_with_references, resolve_openai_image_model

    model = resolve_openai_image_model(None)
    rows: List[Dict[str, Any]] = []
    rels: List[str] = []
    n = max(1, min(int(n or DEFAULT_N), 8))
    for i in range(1, n + 1):
        prompt = _prompt_for_index(i - 1)
        raw, used_model = generate_with_references(
            prompt=prompt,
            image_paths=[Path(wrap_p)],
            model=model,
            size=size,
            quality=quality,
        )
        name = f"hand_hold_{i:02d}.jpg"
        dest = od / name
        dest.write_bytes(raw)
        (od / f"hand_hold_{i:02d}.txt").write_text(prompt, encoding="utf-8")
        rels.append(name)
        rows.append(
            {
                "i": i,
                "path": str(dest),
                "name": name,
                "bytes": len(raw),
                "model": used_model,
            }
        )
        LOG.info("hand_hold generated %s bytes=%s", name, len(raw))

    man = candidate_manifest(
        jan=jan, wrap_path=str(wrap_p), candidates=rels, selected=""
    )
    man["reused"] = False
    man["quality"] = quality
    man["size"] = size
    man["rows"] = rows
    man["updatedAt"] = datetime.now(timezone.utc).isoformat()
    save_manifest(theme_new_dir, man)
    summary = {
        "jan": jan,
        "wrapSample": Path(wrap_p).name,
        "n": n,
        "quality": quality,
        "engine": "openai",
        "rows": rows,
    }
    (od / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return {"ok": True, "reused": False, "candidates": rels, "manifest": man}


def generate_wrap_single_candidates(
    theme_new_dir: Path,
    *,
    jan: str = "",
    n: int = 3,
    quality: str = DEFAULT_QUALITY,
    size: str = DEFAULT_SIZE,
    force: bool = False,
    wrap_seed_name: str = "",
) -> Dict[str, Any]:
    """手なしの個包装単体候補（MAIN／サブ見本用）。"""
    theme_new_dir = Path(theme_new_dir)
    if str(wrap_seed_name or "").strip():
        ensure_wrap_stem_from_name(theme_new_dir, wrap_seed_name)
    from b4_wrap_compose_ref import resolve_wrap_generation_ref

    wrap_p = resolve_wrap_generation_ref(theme_new_dir) or existing_wrap_path(theme_new_dir)
    if wrap_p is None or not Path(wrap_p).is_file():
        raise ValueError(
            "個包装見本がありません。コピペ／URLで登録するか WRAP を選んでから生成してください。"
        )
    if not wrap_asset_lane_enabled():
        raise RuntimeError("B4_WRAP_ASSET_LANE=0 のため停止")

    od = out_dir_for(theme_new_dir)
    od.mkdir(parents=True, exist_ok=True)
    existing = list_single_files(theme_new_dir)
    if existing and not force:
        rels = [p.name for p in existing]
        return {"ok": True, "reused": True, "candidates": rels}

    from openai_image import generate_with_references, resolve_openai_image_model

    model = resolve_openai_image_model(None)
    rels: List[str] = []
    n = max(1, min(int(n or 3), 6))
    for i in range(1, n + 1):
        prompt = _prompt_for_single(i - 1)
        raw, used_model = generate_with_references(
            prompt=prompt,
            image_paths=[Path(wrap_p)],
            model=model,
            size=size,
            quality=quality,
        )
        name = f"wrap_single_{i:02d}.jpg"
        (od / name).write_bytes(raw)
        (od / f"wrap_single_{i:02d}.txt").write_text(prompt, encoding="utf-8")
        rels.append(name)
        LOG.info("wrap_single generated %s model=%s", name, used_model)
    man = load_manifest(theme_new_dir)
    man["singleCandidates"] = rels
    man["recommendedSingle"] = rels[0] if rels else ""
    man["updatedAt"] = datetime.now(timezone.utc).isoformat()
    save_manifest(theme_new_dir, man)
    return {"ok": True, "reused": False, "candidates": rels}


def generate_wrap_asset_bundle(
    theme_new_dir: Path,
    *,
    jan: str = "",
    wrap_seed_name: str = "",
    force: bool = False,
    n_hand: int = DEFAULT_N,
    n_single: int = 3,
) -> Dict[str, Any]:
    """手持ち＋単体をまとめて生成（見本ゲート同一画面用）。"""
    hand = generate_hand_hold_candidates(
        theme_new_dir,
        jan=jan,
        n=n_hand,
        force=force,
        wrap_seed_name=wrap_seed_name,
    )
    single = generate_wrap_single_candidates(
        theme_new_dir,
        jan=jan,
        n=n_single,
        force=force,
        wrap_seed_name=wrap_seed_name,
    )
    rec = recommend_wrap_assets(theme_new_dir)
    man = load_manifest(theme_new_dir)
    man["recommendedHand"] = rec.get("handHold") or ""
    man["recommendedSingle"] = rec.get("wrapSingle") or ""
    save_manifest(theme_new_dir, man)
    return {
        "ok": True,
        "hand": hand,
        "single": single,
        "recommend": rec,
    }


def remake_named_wrap_assets(
    theme_new_dir: Path,
    *,
    names: Sequence[str],
    kind: str,
    jan: str = "",
    wrap_seed_name: str = "",
    user_note: str = "",
    quality: str = DEFAULT_QUALITY,
    size: str = DEFAULT_SIZE,
) -> Dict[str, Any]:
    """
    レ点した候補だけ作り直し（課金抑制）。
    kind: hand | single
    同名ファイルを上書き。無い名前はスキップ。
    user_note: 人間の大きさ・質感などの要望（指示にそのまま足す）。
    """
    theme_new_dir = Path(theme_new_dir)
    kind_s = str(kind or "").strip().lower()
    if kind_s not in ("hand", "single"):
        raise ValueError("kind は hand か single")
    want = [str(x).strip() for x in (names or []) if str(x).strip()]
    if not want:
        raise ValueError("作り直す候補にレ点がありません")
    if str(wrap_seed_name or "").strip():
        ensure_wrap_stem_from_name(theme_new_dir, wrap_seed_name)
    from b4_wrap_compose_ref import resolve_wrap_generation_ref

    wrap_p = resolve_wrap_generation_ref(theme_new_dir) or existing_wrap_path(theme_new_dir)
    if wrap_p is None or not Path(wrap_p).is_file():
        raise ValueError("個包装見本がありません")
    if not wrap_asset_lane_enabled():
        raise RuntimeError("B4_WRAP_ASSET_LANE=0 のため停止")

    from openai_image import generate_with_references, resolve_openai_image_model

    note = " ".join(str(user_note or "").split())[:400]
    od = out_dir_for(theme_new_dir)
    od.mkdir(parents=True, exist_ok=True)
    model = resolve_openai_image_model(None)
    done: List[str] = []
    skipped: List[str] = []
    man0 = load_manifest(theme_new_dir)
    remake_counts: Dict[str, int] = dict(man0.get("remakeCount") or {})
    for name in want:
        dest = od / name
        if not dest.is_file():
            skipped.append(name)
            continue
        # 名前からスロット番号を取る（無ければ 0）
        idx = 0
        try:
            stem = Path(name).stem
            idx = max(0, int(stem.split("_")[-1]) - 1)
        except Exception:
            idx = 0
        round_n = int(remake_counts.get(name) or 0) + 1
        # 参照は個包装見本のみ（いまの hand_hold を種にしない＝似た構図の固定を避ける）
        # 指示はスロットをずらし、前回と同構図を明示禁止
        if kind_s == "hand":
            if not name.startswith("hand_hold_"):
                skipped.append(name)
                continue
            prompt = _prompt_for_index((idx + round_n * 2) % len(_VARIANT_LINES_JA))
            prompt += (
                f"\n【作り直し {round_n} 回目】"
                "前回の手持ち写真と同じ構図・同じ指の形・同じ寄りにしない。"
                "角度・指の向き・植物の位置をはっきり変える。"
                "個包装の見た目だけ添付見本に合わせる（手のポーズは変える）。"
                "個包装は実物サイズに見える大きさ。指に対して小さくしすぎない。"
            )
        else:
            if not name.startswith("wrap_single_"):
                skipped.append(name)
                continue
            prompt = _prompt_for_single((idx + round_n) % len(_SINGLE_LINES_JA))
            prompt += (
                f"\n【作り直し {round_n} 回目】"
                "前回の単体写真と同じ角度・同じ置き方にしない。"
                "向き・寄りを変える。個包装の見た目だけ添付見本に合わせる。"
                "個包装は実物サイズに見える大きさ。小さくしすぎない。"
            )
        if note:
            prompt += f"\n【人間の要望（優先）】{note}"
        raw, used_model = generate_with_references(
            prompt=prompt,
            image_paths=[Path(wrap_p)],
            model=model,
            size=size,
            quality=quality,
        )
        dest.write_bytes(raw)
        dest.with_suffix(".txt").write_text(prompt, encoding="utf-8")
        remake_counts[name] = round_n
        done.append(name)
        LOG.info(
            "remake %s kind=%s round=%s model=%s refs=%s note=%s",
            name,
            kind_s,
            round_n,
            used_model,
            Path(wrap_p).name,
            note[:80] if note else "",
        )
    if not done:
        raise ValueError("作り直せた候補がありません（名前やファイルを確認）")
    man = load_manifest(theme_new_dir)
    man["remakeCount"] = remake_counts
    man["lastRemake"] = {
        "kind": kind_s,
        "done": done,
        "skipped": skipped,
        "at": datetime.now(timezone.utc).isoformat(),
        "jan": jan,
        "seed": Path(wrap_p).name,
        "userNote": note,
        "note": "ref=FOOD_UNIT_WRAP_SEED(or WRAP) only; not previous candidate image",
    }
    save_manifest(theme_new_dir, man)
    return {"ok": True, "kind": kind_s, "done": done, "skipped": skipped, "userNote": note}


def select_hand_hold(
    theme_new_dir: Path,
    selected_name: str,
    *,
    jan: str = "",
) -> Dict[str, Any]:
    """候補から1枚採用。WRAP_HAND_HOLD.jpg にコピー。"""
    import shutil

    theme_new_dir = Path(theme_new_dir)
    name = str(selected_name or "").strip()
    if not name:
        raise ValueError("採用する候補名が空です")
    src = out_dir_for(theme_new_dir) / name
    if not src.is_file():
        raise ValueError(f"候補がありません: {name}")

    man = load_manifest(theme_new_dir) or candidate_manifest(
        jan=jan,
        wrap_path=str(existing_wrap_path(theme_new_dir) or ""),
        candidates=[p.name for p in list_candidate_files(theme_new_dir)],
    )
    man = apply_selection(man, name)
    # 場面の個包装正は検収単体（composeWrap）。無ければ FOOD_UNIT_WRAP 名
    if not str(man.get("composeWrap") or "").strip():
        from b4_wrap_compose_ref import existing_compose_wrap_path

        cw = existing_compose_wrap_path(theme_new_dir)
        if cw is not None:
            man["composeWrap"] = cw.name
    routes = resolve_wrap_asset_for_consumers(man)
    od = out_dir_for(theme_new_dir)
    dest = od / SELECTED_NAME
    shutil.copy2(src, dest)
    refs = theme_new_dir / "01_refs"
    refs.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, refs / SELECTED_NAME)
    man["selectedPath"] = str(dest)
    man["routes"] = routes
    man["selectedAt"] = datetime.now(timezone.utc).isoformat()
    save_manifest(theme_new_dir, man)
    return {
        "ok": True,
        "selected": name,
        "selectedPath": str(dest),
        "routes": routes,
        "manifest": man,
    }


def clear_selection(theme_new_dir: Path) -> None:
    """スキップ時。採用ファイルを消さず selected だけ空にする。"""
    man = load_manifest(theme_new_dir)
    if not man:
        return
    man["selected"] = ""
    man["routes"] = resolve_wrap_asset_for_consumers(man)
    man["skippedAt"] = datetime.now(timezone.utc).isoformat()
    save_manifest(theme_new_dir, man)
