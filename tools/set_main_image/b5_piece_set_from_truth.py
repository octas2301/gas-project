# -*- coding: utf-8 -*-
"""一片セットを実物由来の種から作る（汎用・JAN指定のみ）。

種の優先は b5_piece_truth_rules.resolve_preferred_seed_kind:
  人間実物(PIECE_TRUTH_HUMAN) → WRAP → CANON仮想

案件名・魚種・SKU固有の質感文をコードに焼かない。
人間メモは --note で渡す（言語化へ反映）。

  cd tools/set_main_image
  python b5_piece_set_from_truth.py --jan 4906283045614 --n 3 --note "表面はマット。つるつる禁止"
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from openai_image import generate_with_references, resolve_openai_image_model

from b5_piece_truth_rules import (
    build_shape_policy_prompt_block,
    evaluate_piece_truth_gate,
    infer_shape_family_from_texts,
    resolve_preferred_seed_kind,
    write_piece_set_source,
)
from _tmp_food_piece_set5_lang_scenes import (
    build_set_lang_pack,
    read_set_piece_lang,
)

SCRIPT = Path(__file__).resolve().parent
DEFAULT_N = 3
QUALITY = "low"
SIZE = "1024x1024"

_POSE = (
    "向き01: やや斜め上。厚みと表面の質感がよく見える。",
    "向き02: 真上寄り。上面と境目の様子が見える。",
    "向き03: 別角度の斜め。同じ1片の見た目を保て。",
    "向き04: やや低角度。厚みの断面感が見える。",
    "向き05: 別の斜め。シルエットを別物に変えるな。",
)


def _theme(jan: str) -> Path:
    jan_d = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
    return SCRIPT / "_out" / f"theme_new_run_{jan_d}"


def main(argv: List[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="一片セットを実物由来種から生成")
    ap.add_argument("--jan", required=True, help="数字JAN（theme_new_run_{jan}）")
    ap.add_argument("--n", type=int, default=DEFAULT_N, help="枚数（既定3・最大5）")
    ap.add_argument(
        "--note",
        default="",
        help="人間検収1行。質感・禁止事項。コードに案件文を焼かない代わりにここへ",
    )
    ap.add_argument(
        "--family",
        default="",
        choices=["", "cube", "flake", "other"],
        help="外形ファミリー上書き。空なら種パス名＋noteから推定",
    )
    ap.add_argument("--quality", default=QUALITY)
    ap.add_argument("--dry-seed", action="store_true", help="種解決だけして終了")
    args = ap.parse_args(argv)

    n = max(1, min(5, int(args.n or DEFAULT_N)))
    theme = _theme(args.jan)
    if not theme.is_dir():
        raise SystemExit(f"theme_new 無し: {theme}")

    kind, seeds = resolve_preferred_seed_kind(theme)
    if kind == "none" or not seeds:
        raise SystemExit(
            "種が無い。01_refs/PIECE_TRUTH_HUMAN/piece_truth_0* か "
            "FOOD_UNIT_WRAP か FOOD_UNIT_CANON を置け。"
        )
    print(f"seed kind={kind} n={len(seeds)} paths={[p.name for p in seeds[:5]]}")
    if args.dry_seed:
        return 0

    note = " ".join(str(args.note or "").split())[:240]
    family = str(args.family or "").strip()
    if not family:
        family = infer_shape_family_from_texts(
            [note, kind] + [p.name for p in seeds]
        )
    policy = build_shape_policy_prompt_block(family)

    out = theme / "food_piece_set5_virt"
    set_dir = out / "piece_set"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if set_dir.is_dir() and any(set_dir.glob("piece_set_0*.jpg")):
        backup = out / f"piece_set_prev_{stamp}"
        if backup.exists():
            shutil.rmtree(backup)
        shutil.copytree(set_dir, backup)
        print("backup ->", backup)
    set_dir.mkdir(parents=True, exist_ok=True)
    for p in set_dir.glob("piece_set_0*"):
        p.unlink()

    model = resolve_openai_image_model(None)
    set_files: List[Path] = []
    for i in range(1, n + 1):
        dest = set_dir / f"piece_set_{i:02d}.jpg"
        pose = _POSE[min(i - 1, len(_POSE) - 1)]
        prompt_lines = [
            "【個品・一片見本・実物由来】",
            f"添付の種写真（優先={kind}）が開封後1片の正。色・縁・質感・外形系統を合わせろ。",
            "写真の四角コピペ貼付禁止。実写として描け。",
            "手・指・皿・袋・個包装フィルム・文字は描くな。背景は白〜生成りのみ。",
            policy.strip(),
            pose,
        ]
        if note:
            prompt_lines.append(f"【人間検収】{note}")
        prompt = "\n".join(prompt_lines)
        raw, used = generate_with_references(
            prompt=prompt,
            image_paths=list(seeds),
            model=model,
            size=SIZE,
            quality=str(args.quality or QUALITY),
        )
        dest.write_bytes(raw)
        dest.with_suffix(".txt").write_text(prompt, encoding="utf-8")
        set_files.append(dest)
        print(f"set {dest.name} model={used}")

    write_piece_set_source(
        theme,
        kind=kind,
        paths=seeds,
        extra={
            "jan": "".join(ch for ch in str(args.jan) if ch.isdigit()),
            "n": n,
            "family": family,
            "humanNote": note,
            "builtAt": datetime.now(timezone.utc).isoformat(),
        },
    )

    pieces: List[Dict[str, Any]] = []
    for p in set_files:
        print(f"lang {p.name} ...")
        one = read_set_piece_lang(p)
        if not one.get("ok"):
            raise SystemExit(f"言語化失敗: {p.name} {one.get('error')}")
        if note:
            one["humanNote"] = note
            # 形状欄へ禁止語つきメモを直書きしない（ゲート誤検知の原因）。
            # 表面の質感メモだけ短く足す。
            if isinstance(one.get("fields"), dict):
                fields = dict(one["fields"])
                surface_bit = note
                for tok in ("三角", "扇", "くさび", "楔", "禁止"):
                    surface_bit = surface_bit.replace(tok, "")
                surface_bit = " ".join(surface_bit.split())[:80]
                if surface_bit:
                    cur = str(fields.get("表面") or "")
                    fields["表面"] = (cur + f"／人間修正:{surface_bit}")[:140]
                    one["fields"] = fields
        pieces.append(one)

    block = build_set_lang_pack(pieces)
    pack = {
        "ok": True,
        "jan": "".join(ch for ch in str(args.jan) if ch.isdigit()),
        "rule": f"実物由来 kind={kind} family={family}。向き差のみ。扇くさび禁止はcube方針に従う。",
        "humanNote": note,
        "n": n,
        "family": family,
        "pieces": pieces,
        "promptBlock": block,
        "builtAt": datetime.now(timezone.utc).isoformat(),
        "truthRefs": [str(p) for p in seeds],
    }
    (out / "piece_set_lang.json").write_text(
        json.dumps(pack, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out / "piece_set_lang_block.txt").write_text(block, encoding="utf-8")

    gate = evaluate_piece_truth_gate(theme, max_n=n)
    (out / "piece_truth_gate_after_build.json").write_text(
        json.dumps(gate, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("gate", "PASS" if gate.get("ok") else "FAIL", gate.get("errors") or gate.get("warnings"))
    if not gate.get("ok"):
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
