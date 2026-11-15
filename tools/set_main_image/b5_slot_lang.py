# -*- coding: utf-8 -*-
"""スロット別・絵＋言語化セット（SURFACE／SHAPE／WRAP相対）。

正:
- 添付するスロットには、その役割の言語化をセットで同送する
- SURFACE: 筋・色・血合い色・表面感のみ（外形は書かない）
- SHAPE: 細長さ・反り・ねじれ・薄さのみ（色・筋は書かない）
- WRAP: 既存 wrap_look ＋開封後1片との相対サイズ文

切戻し: B5_SLOT_LANG=0（役割文のみ。Vision面言語化を走らせない）
"""
from __future__ import annotations

import base64
import json
import logging
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

LOG = logging.getLogger("set_main_image.b5_slot_lang")

SURFACE_LANG_NAME = "surface_lang.json"
SURFACE_KEYS = ("色", "筋", "血合い色", "表面感", "寄り指示")
SHAPE_KEYS = ("形状", "厚み硬さ")  # ドナーから形・薄さだけ


def slot_lang_enabled() -> bool:
    raw = str(os.environ.get("B5_SLOT_LANG") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def s1_synthesis_prompt_block() -> str:
    """仮想C型・合成の正（SKU非焼付）。S1で必ず同送。"""
    return (
        "【合成の正・S1・必須】\n"
        "各片の表面＝SURFACE（琥珀色系の面・白い締まり線と細い身の筋・細い血合い帯の色）。\n"
        "各片の形＝SHAPE（細長い削り片。反り・丸まり・ねじれのバリエ。薄さ）。\n"
        "盛り＝IDENTITY（小皿に適量、重なりとばらけ）。\n"
        "【禁止】SURFACEの厚い台形・板状シルエットを皿に並べるな。"
        "SHAPEの色・筋に寄せるな。全部の片を同じ曲がり・同じシルエットにするな。"
        "縁毛羽で面の筋の代用をするな。白背景の単一片標本禁止。\n"
        "【見え方】少し寄って面の筋が読める。形は細長く反ったまま。板状の厚い切片に戻すな。"
        "細長い薄い削り片。一部はねじれ＋丸まりが重なる。単純な一方向の反りだけに揃えるな。\n"
        "【食品面の光・必須】削り片・食品面はしっとりとし、光は鈍く返す。"
        "乾いた板・カサつき禁止。SURFACE言語化に「乾き」とあっても、しっとり・鈍い戻りを優先。"
        "鏡面ハイライト・片目だけのテカリ・オイルの照り返し禁止。"
        "テーマ文の「適度な照り」「スペキュラ」「環境反射」は缶・瓶・皿・ガラスのみ。"
        "削り片に当てるな。\n"
        "【筋の描き方・必須】面とほぼ同じ高さの細い繊維・木目。"
        "溝・段差・彫り込み・深い影の線として描くな。のっぺり無地も禁止。\n"
    )


def _surface_lang_path(theme_dir: Path) -> Path:
    return Path(theme_dir) / "food_piece_set5_virt" / SURFACE_LANG_NAME


def _normalize_surface_fields(fields: Dict[str, str]) -> Dict[str, str]:
    """食品面の湿り・筋の高さ。Visionの乾き／溝表現を上書き。"""
    out = {k: str(fields.get(k) or "").strip()[:120] for k in SURFACE_KEYS}
    feel = out.get("表面感") or ""
    if (not feel) or ("乾" in feel) or ("カサ" in feel) or ("粉" in feel):
        out["表面感"] = "しっとり。光は鈍く返す。乾いた板・カサつき・鏡面テカリ禁止"
    suji = out.get("筋") or ""
    if suji and (("溝" in suji) or ("彫" in suji) or ("段差" in suji)):
        out["筋"] = "面とほぼ同じ高さの細い木目状繊維。溝・段差・彫り込みではない"
    elif suji and ("溝・彫り込みにするな" not in suji) and ("同じ高さ" not in suji):
        out["筋"] = (suji.rstrip("。") + "（面と同じ高さの細い繊維。溝・彫り込みにするな）")[:120]
    return out


def _prompt_block_surface(name: str, fields: Dict[str, str]) -> str:
    fields = _normalize_surface_fields(fields)
    lines = [f"【SURFACE {name}・機械言語化・面の正】"]
    for k in SURFACE_KEYS:
        v = str(fields.get(k) or "").strip()
        if v:
            lines.append(f"- {k}: {v}")
    lines.append(
        "外形・反り・ねじれは書くな／使うな。"
        "湿り・筋の高さは上記の正規化を、色の見えは写真を優先。"
    )
    return "\n".join(lines)


def read_surface_lang_one(path: Path) -> Dict[str, Any]:
    """面アップ1枚から構造化言語化（外形は書かせない）。"""
    from b5_food_look_read import TEXT_MODELS, _parse_obj, load_api_key, make_client, mime_for

    p = Path(path)
    out: Dict[str, Any] = {
        "ok": False,
        "path": str(p),
        "name": p.name,
        "fields": {},
        "promptBlock": "",
        "error": "",
    }
    if not p.is_file():
        out["error"] = "no_file"
        return out
    prompt = (
        "食品の「面アップ」見本です。面の質感だけ書いてください。\n"
        "外形・反り・丸まり・ねじれ・細長さ・全体シルエットは書くな。\n"
        "商品名・推測・数字の厚みは書くな。見えない項目は空文字。\n"
        "色: 身の色相（琥珀色・淡橙など）。\n"
        "筋: 面に見える繊維の筋・締まり線の見え方（木目状・密度・向き）。\n"
        "血合い色: 細い血合い帯の色があれば。無ければ空。\n"
        "表面感: しっとり／乾き。光は「鈍い戻り」か「鏡面テカリ」かを必ず区別。"
        "湿って見える面はしっとり＋鈍い戻り光と書け。強い鏡面ハイライトと書くな。\n"
        "寄り指示: 「面の筋が読める寄り」に使える短い一言。\n"
        'JSONのみ: {"色":"","筋":"","血合い色":"","表面感":"","寄り指示":""}\n'
        "各値は全角60字以内。"
    )
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        out["error"] = str(e)
        return out
    data = base64.b64encode(p.read_bytes()).decode("ascii")
    last = ""
    for mid in TEXT_MODELS:
        if not mid:
            continue
        try:
            resp = client.models.generate_content(
                model=mid,
                contents=[
                    {
                        "role": "user",
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": mime_for(p), "data": data}},
                        ],
                    }
                ],
            )
            obj = _parse_obj(getattr(resp, "text", None) or "")
            if not obj:
                last = (getattr(resp, "text", None) or "")[:160]
                continue
            fields = _normalize_surface_fields(
                {k: str(obj.get(k) or "").strip()[:120] for k in SURFACE_KEYS}
            )
            bits = [f"{k}: {v}" for k, v in fields.items() if v]
            out["ok"] = bool(bits)
            out["fields"] = fields
            out["model"] = mid
            out["promptBlock"] = _prompt_block_surface(p.name, fields)
            return out
        except Exception as e:
            last = str(e)
    out["error"] = last or "parse_failed"
    return out


def load_surface_lang_pack(theme_dir: Path) -> Dict[str, Any]:
    p = _surface_lang_path(theme_dir)
    if not p.is_file():
        return {}
    try:
        obj = json.loads(p.read_text(encoding="utf-8-sig"))
        return obj if isinstance(obj, dict) else {}
    except Exception:
        return {}


def ensure_surface_lang(
    theme_dir: Path,
    surface_paths: Sequence[Path],
    *,
    force: bool = False,
) -> Dict[str, Any]:
    """SURFACE各枚を言語化して保存。既存があり force=False なら再利用。"""
    theme_dir = Path(theme_dir)
    paths = [Path(p) for p in surface_paths if Path(p).is_file()]
    empty = {
        "ok": False,
        "items": [],
        "promptBlock": "",
        "n": 0,
        "reason": "no_paths",
    }
    if not paths:
        return empty
    if not slot_lang_enabled():
        empty["reason"] = "B5_SLOT_LANG=0"
        return empty
    prev = load_surface_lang_pack(theme_dir)
    prev_by = {}
    for it in prev.get("items") or []:
        if isinstance(it, dict) and it.get("name"):
            prev_by[str(it["name"])] = it
    items: List[Dict[str, Any]] = []
    for p in paths:
        cached = prev_by.get(p.name)
        if (
            not force
            and isinstance(cached, dict)
            and cached.get("ok")
            and str(cached.get("path") or "") == str(p.resolve())
        ):
            items.append(cached)
            continue
        rec = read_surface_lang_one(p)
        rec["path"] = str(p.resolve())
        items.append(rec)
        LOG.info(
            "surface_lang %s ok=%s err=%s",
            p.name,
            rec.get("ok"),
            (rec.get("error") or "")[:80],
        )
    blocks = [
        str(it.get("promptBlock") or "").strip()
        for it in items
        if it.get("ok") and str(it.get("promptBlock") or "").strip()
    ]
    pack = {
        "ok": bool(blocks),
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "n": len(items),
        "items": items,
        "promptBlock": "\n\n".join(blocks),
    }
    out_p = _surface_lang_path(theme_dir)
    out_p.parent.mkdir(parents=True, exist_ok=True)
    out_p.write_text(json.dumps(pack, ensure_ascii=False, indent=2), encoding="utf-8")
    return pack


def build_surface_lang_prompt_block(
    theme_dir: Path,
    surface_paths: Sequence[Path],
    *,
    ensure: bool = True,
) -> str:
    if not surface_paths:
        return ""
    pack = load_surface_lang_pack(theme_dir)
    need = ensure and (
        not pack.get("ok")
        or len(pack.get("items") or []) < len([p for p in surface_paths if Path(p).is_file()])
    )
    if need:
        pack = ensure_surface_lang(theme_dir, surface_paths)
    body = str(pack.get("promptBlock") or "").strip()
    if not body:
        return ""
    head = (
        "【SURFACE・機械言語化】面の色・筋・表面感は次の文＋SURFACE写真が正。"
        "外形はSURFACE言語化から取るな。"
    )
    return head + "\n" + body


def build_shape_lang_prompt_from_donors(
    donor_theme_dirs: Sequence[Path],
    *,
    max_n: int = 3,
) -> str:
    """ドナー theme の piece_set_lang から形・薄さだけ抜く。"""
    from b5_piece_truth_rules import load_piece_set_lang

    n = int(max_n or 0)
    chunks = [
        "【SHAPE・機械言語化・形のみ】細長さ・反り・丸まり・ねじれ・薄さだけ。"
        "色・筋・血合い色・縁の質感はSHAPE言語化から取るな。",
    ]
    taken = 0
    for td in donor_theme_dirs:
        if taken >= n > 0:
            break
        lang = load_piece_set_lang(Path(td))
        pieces = lang.get("pieces") if isinstance(lang, dict) else None
        if not isinstance(pieces, list):
            # alternate shapes
            items = lang.get("items") if isinstance(lang, dict) else None
            pieces = items if isinstance(items, list) else []
        for it in pieces:
            if taken >= n > 0:
                break
            if not isinstance(it, dict):
                continue
            fields = it.get("fields") if isinstance(it.get("fields"), dict) else {}
            name = str(it.get("name") or f"shape_{taken + 1}.jpg")
            lines = [f"【SHAPE {name}・形のみ】"]
            n_f = 0
            for k in SHAPE_KEYS:
                v = str(fields.get(k) or "").strip()
                if not v:
                    continue
                lines.append(f"- {k}: {v}")
                n_f += 1
            # promptBlock から形状行だけ拾うフォールバック
            if n_f <= 0:
                pb = str(it.get("promptBlock") or "")
                for line in pb.splitlines():
                    s = line.strip()
                    if s.startswith("- 形状:") or s.startswith("- 厚み硬さ:"):
                        lines.append(s)
                        n_f += 1
            if n_f <= 0:
                continue
            lines.append("色・筋はSURFACEが正。")
            chunks.append("\n".join(lines))
            taken += 1
    if taken <= 0:
        return ""
    return "\n\n".join(chunks)


def wrap_piece_relative_size_block(
    *,
    piece_size_hint: str = "",
    has_wrap: bool = False,
) -> str:
    """開封後1片と個包装の相対サイズ。wrap_look と併用。"""
    if not has_wrap:
        return ""
    hint = str(piece_size_hint or "").strip()
    lines = [
        "【個包装×開封後1片・相対サイズ・必須】",
        "個包装の中に見える中身／個包装の見かけの大きさは、開封後の1片と同じ実寸感。",
        "個包装だけ巨大／極小／別物サイズにするな。",
        "個包装は脇役（皿の横か後ろ・1〜2個）。開封後の盛りより手前に出すな。",
    ]
    if hint:
        lines.append(f"開封後1片の大きさの目安（言語化）: {hint}")
    return "\n".join(lines) + "\n"


def piece_size_hint_from_jan(jan: str, *, max_n: int = 3) -> str:
    """一片言語化の厚み硬さ／形状から短い目安を取る。"""
    try:
        from food_unit_canon import load_piece_lang_prompt_block, theme_new_dir_for_jan
        from b5_piece_truth_rules import load_piece_set_lang
    except Exception:
        return ""
    td = theme_new_dir_for_jan(str(jan or ""))
    lang = load_piece_set_lang(td)
    pieces = lang.get("pieces") if isinstance(lang, dict) else None
    if not isinstance(pieces, list):
        pieces = lang.get("items") if isinstance(lang, dict) else []
    hints: List[str] = []
    for it in list(pieces or [])[: int(max_n or 3)]:
        if not isinstance(it, dict):
            continue
        fields = it.get("fields") if isinstance(it.get("fields"), dict) else {}
        th = str(fields.get("厚み硬さ") or "").strip()
        sh = str(fields.get("形状") or "").strip()
        bit = "／".join(x for x in (th, sh[:40] if sh else "") if x)
        if bit:
            hints.append(bit)
    if hints:
        return "；".join(hints[:3])
    # fallback: raw prompt block snippet
    pb = load_piece_lang_prompt_block(str(jan), max_n=int(max_n or 3))
    for line in pb.splitlines():
        if "厚み硬さ:" in line:
            return line.split("厚み硬さ:", 1)[-1].strip()[:80]
    return ""


def identity_plating_s1_footer() -> str:
    return (
        "配置の正は上記のみ。個品の色・縁・筋・厚み・外形は SURFACE＋SHAPE が正。"
        "一片本体を添付していないときは一片セット言語化を個品の正にするな。"
    )
