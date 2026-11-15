# -*- coding: utf-8 -*-
"""
SKU横断の「1枚あたりの食材見た目」正本（FOOD_UNIT_CANON）。

運用（2026-08-30）:
1. 機械が競合から候補 5〜6 枚をスコア付きで提示
2. 人間が主1＋副0〜2をレ点確定
3. その後に画像生成（ゲートUI）

- 主見本: 色・厚み・1片サイズの正
- 副見本: 盛り付け（置き方・重なり）の参考のみ。1片の形・大きさ・厚み・色には使わない
- 不適なら PACKAGE_TRUTH のみ
"""
from __future__ import annotations

import base64
import json
import logging
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

LOG = logging.getLogger("set_main_image.food_unit_canon")

CANON_STEM = "FOOD_UNIT_CANON"
SECONDARY_STEM = "FOOD_UNIT_SECONDARY"
SIZZLE_STEM = "FOOD_UNIT_SIZZLE"
WRAP_STEM = "FOOD_UNIT_WRAP"
META_NAME = "food_unit_canon.json"
FOOD_CANONS_NAME = "food_canons.json"
CANDIDATES_NAME = "food_unit_candidates.json"
DEFAULT_TOP_K = 6
MIN_SHOW = 5
MAX_SECONDARY = 2

# §7.12.6 テーマ mode（needPackage=true のときは package 優先）
_THEME_FOOD_MODE = {
    1: "blend",
    2: "blend",
    3: "sizzle",
    4: "identity",
    5: "identity",
    6: "identity",
    10: "blend",
    11: "blend",
    12: "blend",
    13: "package",
}


def theme_food_mode(theme_id: int, *, need_package: bool = False) -> str:
    if need_package:
        return "package"
    return _THEME_FOOD_MODE.get(int(theme_id), "identity")

FOOD_UNIT_LOCK_JA = """
【食材ユニット正本（FOOD_UNIT／全テーマ共通・必須）】
目的: テーマが違っても、見本と同じ1片に見えるようにする。

■ 正本の優先順
1) IMAGE_FOOD_UNIT_CANON（主見本）→ 開封後の中身の色・厚み・1片サイズの唯一の正
2) IMAGE_FOOD_UNIT_WRAP（個包装・任意）→ **個包装だけ**の見た目の正。密着・空気・材質（硬い／やわらかい）・縁は見本と機械読みどおり。あるときは皿の横か後ろに1〜2個（小皿に混ぜるな）。カメラ最近傍は開封後の個品。個包装はそれより手前に出すな。無いときは個包装を描くな。WRAP写真に大袋・背景・手が写っていてもそれらは使うな（大袋の正は PACKAGE だけ）
3) IMAGE_FOOD_UNIT_SECONDARY_*（副見本・任意）→ 盛り付け（置き方・重なり）の参考だけ。1片の形・大きさ・厚み・色・質感には使うな。主見本と同じ盛り付けのコピーにはしない
4) 無いとき → IMAGE_PACKAGE_TRUTH の窓・抜き絵・中身カット
5) IMAGE_COMP_* はレイアウト・カード・フレームのみ。食材の色・厚み・1片サイズの正本にしない

■ 揃えること（枚ごとの見た目）
- 色相・照り・しっとり感の系統（主見本／正本窓）
- 1枚（1片）の大きさ感・厚み・形
- 競合の照明差・加工差で色が違っても、色は主見本に寄せる（明るい競合の色に引っ張られるな）

■ 禁止
- テーマごとに別物の食材に見える盛り
- 見本に無い形・厚み・別色の塗り替え
- 副見本やCOMPの色を主見本より優先すること
"""


def existing_canon_path(root: Path) -> Optional[Path]:
    root = Path(root)
    for folder in (root, root / "01_refs"):
        if not folder.is_dir():
            continue
        for ext in (".jpg", ".jpeg", ".png", ".webp"):
            p = folder / f"{CANON_STEM}{ext}"
            if p.is_file():
                return p
    return None


def existing_sizzle_path(root: Path) -> Optional[Path]:
    root = Path(root)
    for folder in (root, root / "01_refs"):
        if not folder.is_dir():
            continue
        for ext in (".jpg", ".jpeg", ".png", ".webp"):
            p = folder / f"{SIZZLE_STEM}{ext}"
            if p.is_file():
                return p
    return None


def existing_wrap_path(root: Path) -> Optional[Path]:
    """個包装見本（任意）。IDENTITY（開封後）とは別。"""
    root = Path(root)
    for folder in (root, root / "01_refs"):
        if not folder.is_dir():
            continue
        for ext in (".jpg", ".jpeg", ".png", ".webp"):
            p = folder / f"{WRAP_STEM}{ext}"
            if p.is_file():
                return p
    return None


def load_food_canons(root: Path) -> Dict[str, Any]:
    for name in (FOOD_CANONS_NAME, META_NAME):
        p = Path(root) / name
        if not p.is_file():
            continue
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
    return {}


def resolve_food_refs_for_theme(
    out_dir: Path,
    theme_id: int,
    *,
    need_package: bool = False,
) -> Dict[str, Any]:
    """
    テーマ mode に応じて載せる食材見本パスを返す。
    keys: mode, identity, sizzle, wrap, secondary, useIdentityAsCanon, useSizzle, useWrap
    """
    mode = theme_food_mode(theme_id, need_package=need_package)
    ident = existing_canon_path(out_dir)
    sizzle = existing_sizzle_path(out_dir)
    wrap = existing_wrap_path(out_dir)
    secs = secondary_paths(out_dir)
    meta = load_food_canons(out_dir)
    package_only = bool(meta.get("packageOnly"))
    if package_only or meta.get("identityStatus") == "absent":
        return {
            "mode": "package",
            "identity": None,
            "sizzle": None,
            "wrap": None,
            "secondary": [],
            "useIdentityAsCanon": False,
            "useSizzle": False,
            "useWrap": False,
            "packageOnly": True,
        }
    use_sizzle = False
    use_ident = True
    if mode == "sizzle":
        use_sizzle = bool(sizzle)
        use_ident = not use_sizzle  # シズル無ければ IDENTITY
    elif mode == "blend":
        use_sizzle = bool(sizzle)
        use_ident = True
    elif mode == "package":
        use_ident = bool(ident)  # 袋テーマでも色参照用に IDENTITY 可
        use_sizzle = False
    else:  # identity
        use_sizzle = False
        use_ident = True
    # WRAP はテーマ mode に依存せず、人間確定があるときだけ載せる
    use_wrap = bool(wrap)
    return {
        "mode": mode,
        "identity": ident if use_ident else None,
        "sizzle": sizzle if use_sizzle else None,
        "wrap": wrap if use_wrap else None,
        "secondary": secs if mode in ("identity", "blend") else [],
        "useIdentityAsCanon": bool(use_ident and ident),
        "useSizzle": bool(use_sizzle and sizzle),
        "useWrap": bool(use_wrap and wrap),
        "packageOnly": False,
    }


def theme_new_dir_for_jan(jan: str, *, script_dir: Optional[Path] = None) -> Path:
    """サブ見本ゲートの正: tools/set_main_image/_out/theme_new_run_{jan}"""
    root = Path(script_dir) if script_dir is not None else Path(__file__).resolve().parent
    jan_d = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
    return root / "_out" / f"theme_new_run_{jan_d}"


# 一片作成時の機械言語化（MAINへ同送）。見つかった最初の promptBlock を正とする。
# セット全枚言語化を最優先（§4.5.10bis）。
_PIECE_LANG_REL = (
    "food_piece_set5_virt/piece_set_lang.json",
    "food_piece_lang.json",
    "food_piece_set5_virt/ref_fuzz_lang.json",
    "food_piece_lang_virt/piece_lang.json",
)


_PIECE_LANG_FIELD_KEYS = ("形状", "縁", "色透過", "表面", "厚み硬さ", "血合い")


def _piece_prompt_from_fields(name: str, fields: Dict[str, Any]) -> str:
    """一片の structured fields から同送文を組む（汚染された promptBlock より優先）。"""
    if not isinstance(fields, dict) or not fields:
        return ""
    lines = [f"【一片セット {name}・機械言語化・正】"]
    n = 0
    for key in _PIECE_LANG_FIELD_KEYS:
        val = str(fields.get(key) or "").strip()
        if not val:
            continue
        lines.append(f"- {key}: {val}")
        n += 1
    if n <= 0:
        return ""
    lines.append("文章と写真が矛盾したら写真を優先。")
    return "\n".join(lines)


def build_piece_set_lang_prompt_from_pieces(
    pieces: Sequence[Dict[str, Any]], *, max_n: int = 0
) -> str:
    """pieces[] から同送用 promptBlock を組む。max_n>0 なら先頭 max_n 件だけ。

    fields があれば fields から組む（後付けで promptBlock に混入した個品外文を避ける）。
    """
    rows: List[Dict[str, Any]] = [p for p in (pieces or []) if isinstance(p, dict)]
    n = int(max_n or 0)
    if n > 0:
        rows = rows[:n]
    if not rows:
        return ""
    chunks = [
        "【一片セット・機械言語化】場面の各片は、対応する番号の文章＋写真に合わせろ。",
        f"今回の正は添付番号 1〜{len(rows)}（＝piece_set 先頭{len(rows)}枚）だけ。"
        "添付していない番号の言語化は使うな。",
        "血合いがある片にだけ血合いを出せ。無い片へ縁の濃褐色帯を付け足すな。",
        "「縁は濃褐色」を全片の共通ルールにするな（血合いの誤用禁止）。",
        "人間が修正するときは、画像とこの言語化の両方を直すこと。",
    ]
    for i, p in enumerate(rows, 1):
        name = str(p.get("name") or f"piece_set_{i:02d}.jpg").strip()
        fields = p.get("fields") if isinstance(p.get("fields"), dict) else {}
        from_fields = _piece_prompt_from_fields(name, fields)
        if from_fields:
            chunks.append(from_fields)
            continue
        block = str(p.get("promptBlock") or "").strip()
        if block:
            chunks.append(block)
    return "\n".join(chunks)


def load_piece_lang_prompt_block(
    jan: str,
    *,
    max_n: int = 0,
    script_dir: Optional[Path] = None,
) -> str:
    """theme_new_run_{jan} 配下の一片言語化。

    max_n>0 かつ pieces[] があるときは先頭 max_n 件だけ再組立（写真枚数と一致）。
    max_n<=0 は従来どおり保存済み promptBlock 全文（一片オフ時の後方互換）。
    """
    tn = theme_new_dir_for_jan(jan, script_dir=script_dir)
    if not tn.is_dir():
        return ""
    n = int(max_n or 0)
    for rel in _PIECE_LANG_REL:
        p = tn / rel
        if not p.is_file():
            continue
        try:
            obj = json.loads(p.read_text(encoding="utf-8-sig"))
        except Exception:
            continue
        if not isinstance(obj, dict):
            continue
        pieces = obj.get("pieces")
        if n > 0 and isinstance(pieces, list) and pieces:
            rebuilt = build_piece_set_lang_prompt_from_pieces(pieces, max_n=n)
            if rebuilt:
                try:
                    from b5_generic_plating import (
                        generic_plating_core_enabled,
                        sanitize_danger_prompt,
                    )

                    if generic_plating_core_enabled():
                        return sanitize_danger_prompt(rebuilt)
                except Exception:
                    pass
                return rebuilt
        block = str(obj.get("promptBlock") or "").strip()
        if block:
            try:
                from b5_generic_plating import (
                    generic_plating_core_enabled,
                    sanitize_danger_prompt,
                )

                if generic_plating_core_enabled():
                    return sanitize_danger_prompt(block)
            except Exception:
                pass
            return block
    return ""


def resolve_piece_set_paths(
    jan: str,
    *,
    max_n: int = 0,
    script_dir: Optional[Path] = None,
) -> List[Path]:
    """一片セット jpg。max_n<=0 なら空。ファイル名順の先頭 max_n 枚。

    ゲートで「一片見本なし」確定（hasPieceSet=false）のときは空を返す。
    """
    if int(max_n or 0) <= 0:
        return []
    tn = theme_new_dir_for_jan(jan, script_dir=script_dir)
    try:
        meta = load_food_canons(tn) if tn.is_dir() else {}
        if isinstance(meta, dict) and meta.get("hasPieceSet") is False:
            return []
    except Exception:
        pass
    set_dir = tn / "food_piece_set5_virt" / "piece_set"
    if not set_dir.is_dir():
        return []
    files = sorted(set_dir.glob("piece_set_0*.jpg"))
    return files[: int(max_n)]


def piece_set_lock_prompt_block(*, max_n: int) -> str:
    """一片セット写真を添付するときの描き直し禁止（G2系）。"""
    n = int(max_n or 0)
    if n <= 0:
        return ""
    base = (
        "【一片セット・ロック・最優先】\n"
        f"添付の IMAGE_FOOD_UNIT_PIECE_1〜{n} が開封後個品の正（色・縁・筋・厚み・外形・血合い）。\n"
        "場面に出す各片はこれらの見た目を保て。別物に描き直すな。\n"
        f"言語化も 1〜{n} 番だけが正。添付していない番号の文は無視せよ。\n"
        "【正の一本化】次を無効化せよ:"
        "「CANONが唯一の正」「厚み・色は主見本IDENTITYだけ」「個品の色はIDENTITY優先」。\n"
        "IMAGE_FOOD_UNIT_CANON／IDENTITY／SECONDARY は盛り配置・量感の参考のみ"
        "（個品の色・縁・血合いの正にしない）。\n"
        "袋窓の中身は PACKAGE 正本のみ。一片セットで窓の中身を描き変えるな。\n"
        "血合いがある片にだけ血合い筋を出せ。無い片へ縁の濃褐色帯を付け足すな。"
        "「縁は濃褐色」を全片の共通ルールにするな。\n"
        "料理上の具も同じ一片ルール。汁・熱で色を濃くするな。厚みを変えるな。\n"
        "切り抜き貼付の指示ではない。セット写真を見本にした実写生成をせよ。\n"
        "同一シルエットのコピー増殖禁止。1本の長大帯禁止。実寸を揃えろ。\n"
        "【見せ方】白背景の単一片標本は禁止（特にSUB）。"
        "白い小皿に適量盛り、顧客が買いたくなる見せ方にせよ。\n"
    )
    try:
        from b5_piece_deepen import surface_shape_role_prompt_block

        extra = surface_shape_role_prompt_block()
    except Exception:
        extra = ""
    return base + (extra if extra else "")


# IDENTITY盛り＝配置・量感だけ（個品の縁・表面・筋は載せない）
_IDENTITY_PLATING_FIELD_KEYS = (
    "器",
    "量感",
    "重なり",
    "配置",
    "山の高さ",
    "ばらけ幅",
    "枚数目安",
    "背景光",
    "禁じ方",
)
# 配置ブロックの見出し。これ以外の【…】は個品混入として落とす候補
_IDENTITY_PLATING_SECTION_OK = (
    "IDENTITY盛り",
    "配置の正",
    "盛り・配置",
    "盛り配置",
)
# 明らかに個品の縁・表面をIDENTITYに足した見出し
_IDENTITY_PIECE_STATE_SECTION_MARKERS = (
    "盛り上の個品",
    "個品の状態",
    "個品・状態",
    "縁の質感",
    "表面の質感",
)


def identity_plating_fields_only_enabled() -> bool:
    """切戻し: B5_IDENTITY_PLATING_FIELDS_ONLY=0"""
    import os

    raw = str(os.environ.get("B5_IDENTITY_PLATING_FIELDS_ONLY") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def build_identity_plating_prompt_from_fields(fields: Dict[str, Any]) -> str:
    """配置fieldsだけからIDENTITY同送文を組む（SKU非焼付）。"""
    if not isinstance(fields, dict) or not fields:
        return ""
    lines = ["【IDENTITY盛り・機械言語化・配置の正】", ""]
    n_field = 0
    for key in _IDENTITY_PLATING_FIELD_KEYS:
        val = str(fields.get(key) or "").strip()
        if not val:
            continue
        lines.append(f"- {key}: {val}")
        n_field += 1
    if n_field <= 0:
        return ""
    lines.append("")
    lines.append(
        "配置の正は上記のみ。個品の色・縁・筋・厚み・外形は、"
        "添付している SURFACE／SHAPE／一片セット＋各スロットの言語化が正。"
        "添付していないスロットの文を個品の正にするな。"
        "IDENTITYに個品の縁・表面文を足すな。"
    )
    return "\n".join(lines).strip()


def sanitize_identity_plating_text(text: str) -> str:
    """IDENTITY文から個品の縁・表面ブロックを落とす（語のブラックリストではない）。"""
    raw = str(text or "").strip()
    if not raw:
        return ""
    # 【見出し】単位で分割
    parts = re.split(r"(?=【)", raw)
    kept: List[str] = []
    for part in parts:
        chunk = str(part or "").strip()
        if not chunk:
            continue
        head_m = re.match(r"【([^】]+)】", chunk)
        if head_m:
            head = head_m.group(1)
            if any(m in head for m in _IDENTITY_PIECE_STATE_SECTION_MARKERS):
                continue
            # 配置系見出し／fields箇条書きだけの塊は残す
            if any(ok in head for ok in _IDENTITY_PLATING_SECTION_OK):
                kept.append(chunk)
                continue
            # 未知見出しは、配置キーを含むなら残し、個品語だらけなら落とす
            if any(f"- {k}:" in chunk or f"{k}:" in chunk for k in _IDENTITY_PLATING_FIELD_KEYS):
                # 個品状態の長文が後ろに付いていたら、配置行だけ残す
                lines_keep = []
                for line in chunk.splitlines():
                    s = line.strip()
                    if not s:
                        lines_keep.append(line)
                        continue
                    if s.startswith("【"):
                        lines_keep.append(line)
                        continue
                    if any(s.startswith(f"- {k}:") or s.startswith(f"{k}:") for k in _IDENTITY_PLATING_FIELD_KEYS):
                        lines_keep.append(line)
                        continue
                    # 配置の説明文（禁じ方の補足など）は短文のみ残す
                    if s.startswith("配置の正") or s.startswith("禁止"):
                        lines_keep.append(line)
                kept.append("\n".join(lines_keep).strip())
                continue
            # 配置キーも無く個品見出しでもない → 落とす（安全側）
            continue
        # 見出し無しの前置文: 配置キー行だけ拾う
        if any(f"- {k}:" in chunk for k in _IDENTITY_PLATING_FIELD_KEYS):
            lines_keep = [
                line
                for line in chunk.splitlines()
                if any(
                    line.strip().startswith(f"- {k}:") or line.strip().startswith(f"{k}:")
                    for k in _IDENTITY_PLATING_FIELD_KEYS
                )
                or line.strip().startswith("【")
            ]
            if lines_keep:
                kept.append("\n".join(lines_keep).strip())
    out = "\n\n".join(k for k in kept if k).strip()
    return out


def load_identity_plating_lang_block(
    jan: str, *, script_dir: Optional[Path] = None
) -> str:
    """IDENTITY盛りの機械言語化（配置のみ）。①fields優先。混入した個品文は除去。"""
    tn = theme_new_dir_for_jan(jan, script_dir=script_dir)
    p = tn / "food_piece_set5_virt" / "identity_plating_lang.json"
    if not p.is_file():
        return ""
    try:
        obj = json.loads(p.read_text(encoding="utf-8-sig"))
    except Exception:
        return ""
    if not isinstance(obj, dict):
        return ""
    fields = obj.get("fields") if isinstance(obj.get("fields"), dict) else {}
    if identity_plating_fields_only_enabled():
        from_fields = build_identity_plating_prompt_from_fields(fields)
        if from_fields:
            return from_fields
        src = str(obj.get("sourcePromptBlock") or "").strip()
        if src:
            cleaned = sanitize_identity_plating_text(src)
            if cleaned:
                return cleaned
        pb = str(obj.get("promptBlock") or "").strip()
        if pb:
            cleaned = sanitize_identity_plating_text(pb)
            if cleaned:
                return cleaned
        return ""
    # 切戻し: 旧どおり source 優先
    src = str(obj.get("sourcePromptBlock") or "").strip()
    if src:
        return src
    return str(obj.get("promptBlock") or "").strip()


def identity_plating_lock_prompt_block() -> str:
    """IDENTITY盛り写真＋言語化を配置の正にする汎用ロック（SKU非焼付）。"""
    return (
        "【IDENTITY盛り・ロック・配置の正】\n"
        "添付の IMAGE_FOOD_UNIT_CANON／IDENTITY 写真と IDENTITY盛り言語化が、"
        "器・量感・重なり・交差・ばらけ幅・山の高さの正。\n"
        "言語化が粗い／空の項目は写真を見ろ。写真と違う盛り型に描き直すな。\n"
        "禁止の例: 一列標本並べ、円形の均一放射だけ、過剰な高い山、空疎なバラまき、"
        "見本と違う密度。\n"
        "料理・シズル場面でも、白い小皿そのものの複製は必須でないが、"
        "具の重なり・密度・ばらけ・交差の型は IDENTITY盛りを移植せよ。\n"
        "個品の色・縁・厚み・外形・血合いは一片セット写真＋一片言語化が正"
        "（IDENTITY盛りを個品の正に戻すな）。\n"
    )


def identity_plating_prompt_blocks_for_parts(
    combo_parts: Sequence[Dict[str, Any]],
    *,
    script_dir: Optional[Path] = None,
) -> str:
    """構成ごとの IDENTITY盛り言語化。"""
    chunks: List[str] = []
    seen: set[str] = set()
    for raw in combo_parts or []:
        if not isinstance(raw, dict):
            continue
        jan_d = "".join(ch for ch in str(raw.get("jan") or "") if ch.isdigit())
        if not jan_d or jan_d in seen:
            continue
        block = load_identity_plating_lang_block(jan_d, script_dir=script_dir)
        if not block:
            continue
        seen.add(jan_d)
        lab = str(raw.get("nameBase") or raw.get("name") or jan_d).strip() or jan_d
        chunks.append(f"【{lab}】\n{block}")
    if not chunks:
        return ""
    return (
        identity_plating_lock_prompt_block()
        + "【IDENTITY盛り・必須】器・量・重なり・配置・光だけ。"
        "個品の見た目は一片セット写真＋一片言語化が正。\n"
        + "\n\n".join(chunks)
        + "\n"
    )


def piece_lang_prompt_blocks_for_parts(
    combo_parts: Sequence[Dict[str, Any]],
    *,
    max_n: int = 0,
    script_dir: Optional[Path] = None,
) -> str:
    """構成ごとの一片作成時言語化をMAINプロンプト用に連結。

    max_n>0 のとき各JANの言語化を添付枚数に揃える（§4.5.10bis-T）。
    """
    chunks: List[str] = []
    seen_jan: set[str] = set()
    n = int(max_n or 0)
    for raw in combo_parts or []:
        if not isinstance(raw, dict):
            continue
        jan_d = "".join(ch for ch in str(raw.get("jan") or "") if ch.isdigit())
        if not jan_d or jan_d in seen_jan:
            continue
        block = load_piece_lang_prompt_block(
            jan_d, max_n=n, script_dir=script_dir
        )
        if not block:
            continue
        seen_jan.add(jan_d)
        lab = str(raw.get("nameBase") or raw.get("name") or jan_d).strip() or jan_d
        chunks.append(
            f"【{lab}・一片作成時の機械言語化・正】\n"
            "小皿の開封後も、料理上の具も、下記の片ごと言語化＋添付 IMAGE_FOOD_UNIT_PIECE に合わせろ。\n"
            "IDENTITY／SECONDARY は盛り配置のみ（色・縁・血合いの正にしない）。\n"
            "汁・熱で色を濃くするな。厚みを変えるな。"
            "血合いが無い片へ縁の濃褐色帯を足すな。"
            "血合いがある片は身の縁の血合い筋のみ（焦げ枠・外皮ライン禁止）。\n"
            f"{block}"
        )
    if not chunks:
        return ""
    head = (
        "【一片言語化・必須】見本写真だけの推測で厚み・縁・筋・血合いを変えるな。"
        "下の作成時言語化が正。写真と矛盾したら写真優先。"
        "片ごとのフィールドを要約して落とすな。\n"
    )
    if n > 0:
        head += f"添付番号は 1〜{n} まで。それ以外の片言語は使うな。\n"
    return head + "\n\n".join(chunks) + "\n"


def food_scale_lock_prompt_block(
    *, has_wrap: bool = False, piece_size_hint: str = ""
) -> str:
    """開封後と個包装・料理上の大きさ揃え。"""
    lines = [
        "【大きさ・必須】",
        "開封後の1片／キューブと、料理の上に載せる具は同じ実寸感。卵かけ上だけ細煙・別サイズにするな。",
    ]
    if has_wrap:
        try:
            from b5_slot_lang import wrap_piece_relative_size_block

            rel = wrap_piece_relative_size_block(
                has_wrap=True, piece_size_hint=str(piece_size_hint or "")
            )
            if rel:
                lines.append(rel.rstrip())
            else:
                lines.append(
                    "個包装の中身も開封後と同じ大きさ。個包装だけ巨大／極小にしない。"
                    "個包装は脇役（1〜2個・皿の横か後ろ）。開封後より手前に出すな。"
                )
        except Exception:
            lines.append(
                "個包装の中身も開封後と同じ大きさ。個包装だけ巨大／極小にしない。"
                "個包装は脇役（1〜2個・皿の横か後ろ）。開封後より手前に出すな。"
            )
    return "\n".join(lines) + "\n"


def _named_ref_file(root: Path, name: str) -> Optional[Path]:
    name = str(name or "").strip()
    if not name:
        return None
    for folder in (Path(root) / "01_refs", Path(root)):
        p = folder / name
        if p.is_file():
            return p
    return None


def _clear_food_stems(refs: Path) -> None:
    if not refs.is_dir():
        return
    for p in list(refs.iterdir()):
        if not p.is_file():
            continue
        stem = p.stem.upper()
        if stem in (CANON_STEM, SIZZLE_STEM) or stem.startswith(SECONDARY_STEM + "_"):
            p.unlink()


def prepare_sub_food_canon(
    jan: str,
    out_dir: Path,
    *,
    truth: Optional[Path] = None,
    script_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """サブ生成へ、人間が確定した個品だけを写す。機械の選び直しはしない。"""
    tn = theme_new_dir_for_jan(jan, script_dir=script_dir)
    meta = load_food_canons(tn) if tn.is_dir() else {}
    confirmed = str(meta.get("confirmedBy") or "")
    base: Dict[str, Any] = {
        "ok": False,
        "blocked": True,
        "packageOnly": False,
        "confirmedBy": confirmed,
        "source": "",
        "path": "",
        "identity": "",
        "secondary": [],
        "sizzle": "",
        "chosenName": "",
        "themeNewDir": str(tn),
        "reasonJa": "",
    }
    if confirmed != "human":
        base["reasonJa"] = (
            "個品見本が未確定です。見本ゲートで主見本を確定してから生成してください。"
            "機械が選んだ競合写真では生成しません。"
        )
        LOG.info("sub food blocked jan=%s confirmedBy=%s", jan, confirmed or "empty")
        return base

    package_only = bool(meta.get("packageOnly")) or meta.get("identityStatus") == "absent"
    out_dir = Path(out_dir)
    refs = out_dir / "01_refs"
    refs.mkdir(parents=True, exist_ok=True)
    same_dir = out_dir.resolve() == tn.resolve()

    if package_only:
        if not same_dir:
            _clear_food_stems(refs)
            slim = {
                "builtAt": datetime.now(timezone.utc).isoformat(),
                "truth": str(truth or meta.get("truth") or ""),
                "confirmedBy": "human",
                "packageOnly": True,
                "identityStatus": "absent",
                "sizzleStatus": "absent",
                "source": "human_package_only",
                "themeNewDir": str(tn),
            }
            for name in (FOOD_CANONS_NAME, META_NAME):
                (out_dir / name).write_text(
                    json.dumps(slim, ensure_ascii=False, indent=2), encoding="utf-8"
                )
        base.update(
            {
                "ok": True,
                "blocked": False,
                "packageOnly": True,
                "source": "human_package_only",
                "reasonJa": "人間確定は正本のみ。個品プレートは付けない。",
            }
        )
        LOG.info("sub food package-only jan=%s", jan)
        return base

    identity_name = str(meta.get("identityName") or meta.get("primaryName") or "").strip()
    src = _named_ref_file(tn, identity_name)
    if src is None:
        base["reasonJa"] = f"人間確定の主見本ファイルがありません: {identity_name or '(名前なし)'}"
        LOG.info("sub food missing identity jan=%s name=%s", jan, identity_name)
        return base

    if not same_dir:
        _clear_food_stems(refs)
    dest = refs / f"{CANON_STEM}{src.suffix.lower() or '.jpg'}"
    if dest.resolve() != src.resolve():
        shutil.copy2(src, dest)

    secondary_out: List[str] = []
    for i, name in enumerate(list(meta.get("secondaryNames") or [])[:MAX_SECONDARY], 1):
        sp = _named_ref_file(tn, str(name))
        if sp is None:
            continue
        sd = refs / f"{SECONDARY_STEM}_{i}{sp.suffix.lower() or '.jpg'}"
        if sd.resolve() != sp.resolve():
            shutil.copy2(sp, sd)
        secondary_out.append(str(sd))

    sizzle_out = ""
    sizzle_name = str(meta.get("sizzleName") or "").strip()
    if sizzle_name:
        sz = _named_ref_file(tn, sizzle_name)
        if sz is not None:
            sdest = refs / f"{SIZZLE_STEM}{sz.suffix.lower() or '.jpg'}"
            if sdest.resolve() != sz.resolve():
                shutil.copy2(sz, sdest)
            sizzle_out = str(sdest)

    if not same_dir:
        slim = {
            "builtAt": datetime.now(timezone.utc).isoformat(),
            "truth": str(truth or meta.get("truth") or ""),
            "confirmedBy": "human",
            "identityName": identity_name,
            "primaryName": identity_name,
            "sizzleName": sizzle_name or None,
            "secondaryNames": list(meta.get("secondaryNames") or [])[:MAX_SECONDARY],
            "identityStatus": "present",
            "sizzleStatus": "present" if sizzle_out else "absent",
            "packageOnly": False,
            "source": "human_confirmed",
            "path": str(dest),
            "chosenName": identity_name,
            "themeNewDir": str(tn),
        }
        for name in (FOOD_CANONS_NAME, META_NAME):
            (out_dir / name).write_text(
                json.dumps(slim, ensure_ascii=False, indent=2), encoding="utf-8"
            )

    base.update(
        {
            "ok": True,
            "blocked": False,
            "source": "human_confirmed",
            "path": str(dest),
            "identity": str(dest),
            "secondary": secondary_out,
            "sizzle": sizzle_out,
            "chosenName": identity_name,
            "reasonJa": "人間確定の主見本を使用。機械の選び直しはしない。",
        }
    )
    LOG.info(
        "sub food human jan=%s identity=%s secondary=%s",
        jan,
        identity_name,
        len(secondary_out),
    )
    return base


def normalize_name_base_key(s: str) -> str:
    return " ".join(str(s or "").split())


def _script_dir(script_dir: Optional[Path] = None) -> Path:
    return Path(script_dir) if script_dir is not None else Path(__file__).resolve().parent


def _bag_name_ok(path: Path, name_base: str) -> bool:
    """ファイル名が商品名ベースと矛盾しない（チーズ混入などを落とす）。"""
    nb = normalize_name_base_key(name_base)
    fn = path.name
    if "チーズ" in fn and "チーズ" not in nb:
        return False
    return True


def _folder_name_base(root: Path) -> str:
    for fname in ("gate_session.json", FOOD_CANONS_NAME, META_NAME):
        p = Path(root) / fname
        if not p.is_file():
            continue
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(data, dict):
            continue
        nb = str(data.get("nameBase") or data.get("productName") or "").strip()
        if nb:
            return nb
    return ""


def _amazon_bag_dirs(work_root: Optional[Path] = None) -> List[Path]:
    from work_paths import DEFAULT_WORK_ROOT, SUB_AMAZON, SUB_RAKUTEN_DONE

    root = Path(work_root) if work_root is not None else DEFAULT_WORK_ROOT
    am = root / SUB_AMAZON
    return [am, am / SUB_RAKUTEN_DONE]


def _resolve_bag_rel(rel: str, *, work_root: Optional[Path] = None) -> Optional[Path]:
    name = Path(str(rel or "").replace("/", "\\").strip()).name
    if not name:
        return None
    for folder in _amazon_bag_dirs(work_root):
        if not folder.is_dir():
            continue
        p = folder / name
        if p.is_file():
            return p
    return None


def mix_side_tokens(name_base: str) -> List[str]:
    """互換用。正のミックス紐付けは combo_parts_from_sheet（セット構成マスタ）。"""
    s = str(name_base or "")
    if "・" not in s:
        return []
    m = re.search(r"（([^）]+)）", s)
    inner = m.group(1) if m else s
    parts = [p.strip() for p in inner.split("・")]
    parts = [p for p in parts if len(p) >= 2]
    return parts if len(parts) >= 2 else []


def _master_parent_index(
    rows: List[List[Any]],
) -> tuple:
    """(header_i, idx, parent_col, child_col, name_col, jan_col)。失敗時は ValueError。"""
    from master_sets import CHILD_HINTS, PARENT_HINTS, _col, _find_header_row

    hi, idx = _find_header_row(rows)
    ip = _col(idx, PARENT_HINTS)
    ic = _col(idx, CHILD_HINTS)
    if ip is None or ic is None:
        raise ValueError("親SKU/子SKU列が解決できません")
    name_i = idx.get("商品名ベース")
    jan_i = idx.get("JANコード") if "JANコード" in idx else idx.get("JAN")
    return hi, idx, ip, ic, name_i, jan_i


def _lookup_parent_row_meta(
    rows: List[List[Any]],
    *,
    parent_sku: str,
) -> Dict[str, Any]:
    """単品親SKU → nameBase / jan（親行のみ）。"""
    from master_sets import _norm

    want = _norm(parent_sku)
    if not want:
        return {}
    hi, _idx, ip, ic, name_i, jan_i = _master_parent_index(rows)
    for row in rows[hi + 1 :]:
        if max(ip, ic) >= len(row):
            continue
        if _norm(row[ic]):
            continue
        if _norm(row[ip]) != want:
            continue
        nb = _norm(row[name_i]) if name_i is not None and name_i < len(row) else ""
        jan_raw = row[jan_i] if jan_i is not None and jan_i < len(row) else ""
        jan = "".join(ch for ch in str(jan_raw or "") if ch.isdigit())
        return {"parentSku": want, "nameBase": nb, "jan": jan}
    return {"parentSku": want, "nameBase": "", "jan": ""}


def combo_component_skus_from_sheet(
    parent_sku: str,
    *,
    rows: Optional[List[List[Any]]] = None,
) -> List[Dict[str, Any]]:
    """マスタ セット構成1〜5 の親SKU（数量は任意・無くてもSKUを残す）。"""
    from master_sets import _leading_count, _norm

    want = _norm(parent_sku)
    if not want:
        return []
    if rows is None:
        try:
            from sheets_master import fetch_master_rows

            rows, _info = fetch_master_rows()
        except Exception as e:
            LOG.warning("combo sheet fetch fail: %s", e)
            return []
    try:
        hi, idx, ip, ic, _name_i, _jan_i = _master_parent_index(rows)
    except Exception as e:
        LOG.warning("combo sheet header fail: %s", e)
        return []
    hit = None
    for row in rows[hi + 1 :]:
        if max(ip, ic) >= len(row):
            continue
        if _norm(row[ic]):
            continue
        if _norm(row[ip]) != want:
            continue
        hit = row
        break
    if hit is None:
        LOG.info("combo sheet parent miss sku=%s", want)
        return []
    out: List[Dict[str, Any]] = []
    for n in range(1, 6):
        sku_i = idx.get(f"セット構成{n}_親SKU")
        qty_i = idx.get(f"セット構成{n}_数量")
        if sku_i is None or sku_i >= len(hit):
            continue
        sku = _norm(hit[sku_i])
        if not sku:
            continue
        raw_qty = _norm(hit[qty_i]) if qty_i is not None and qty_i < len(hit) else ""
        q = _leading_count(raw_qty)
        out.append({"parentSku": sku, "qty": int(q) if q else 0, "slot": n})
    return out


def combo_parts_from_sheet(
    *,
    parent_sku: str,
    work_root: Optional[Path] = None,
    rows: Optional[List[List[Any]]] = None,
) -> List[Dict[str, Any]]:
    """ミックス個品の正: マスタセット構成1〜5 → 各単品の nameBase/JAN/袋。

    2種未満なら空（単品扱い）。商品名の中黒では分けない。
    """
    sku_me = str(parent_sku or "").strip()
    if not sku_me:
        return []
    if rows is None:
        try:
            from sheets_master import fetch_master_rows

            rows, _info = fetch_master_rows()
        except Exception as e:
            LOG.warning("combo_parts_from_sheet fetch fail: %s", e)
            return []
    comps = combo_component_skus_from_sheet(sku_me, rows=rows)
    if len(comps) < 2:
        LOG.info(
            "combo_parts_from_sheet single-or-empty parent=%s n=%s",
            sku_me,
            len(comps),
        )
        return []
    try:
        from b5_parent_shared import load_parent_shared
    except Exception as e:
        LOG.warning("combo_parts_from_sheet shared: %s", e)
        shared_by: Dict[str, Any] = {}
    else:
        shared_by = load_parent_shared(work_root).get("byParentSku") or {}

    parts: List[Dict[str, Any]] = []
    for comp in comps:
        csku = str(comp.get("parentSku") or "").strip()
        meta = _lookup_parent_row_meta(rows, parent_sku=csku)
        nb = str(meta.get("nameBase") or "").strip()
        jan = str(meta.get("jan") or "").strip()
        shared = shared_by.get(csku) if isinstance(shared_by.get(csku), dict) else {}
        if not nb and isinstance(shared, dict):
            nb = str(shared.get("nameBase") or "").strip()
        if not jan and isinstance(shared, dict):
            jan = "".join(ch for ch in str(shared.get("jan") or "") if ch.isdigit())
        rel = ""
        if isinstance(shared, dict):
            rel = str(shared.get("packageTruthRel") or "").strip()
        bag = _resolve_bag_rel(rel, work_root=work_root) if rel else None
        if bag is None:
            bag = _find_sku_bag_file(csku, work_root=work_root)
        parts.append(
            {
                "parentSku": csku,
                "nameBase": nb,
                "jan": jan,
                "qty": int(comp.get("qty") or 0),
                "packageTruth": str(bag) if bag is not None else "",
            }
        )
    if any(not str(p.get("nameBase") or "").strip() for p in parts):
        LOG.warning(
            "combo_parts_from_sheet missing nameBase parent=%s parts=%s",
            sku_me,
            [(p.get("parentSku"), p.get("nameBase"), p.get("jan")) for p in parts],
        )
    # nameBase が空でも SKU+JAN があれば compose 側でフォルダ解決できるよう残す
    ok_parts = [
        p
        for p in parts
        if str(p.get("parentSku") or "").strip()
        and (str(p.get("nameBase") or "").strip() or str(p.get("jan") or "").strip())
    ]
    if len(ok_parts) < 2:
        LOG.info(
            "combo_parts_from_sheet unresolved parent=%s ok=%s/%s",
            sku_me,
            len(ok_parts),
            len(parts),
        )
        return []
    LOG.info(
        "combo_parts_from_sheet parent=%s n=%s names=%s",
        sku_me,
        len(ok_parts),
        [p.get("nameBase") or p.get("jan") for p in ok_parts],
    )
    return ok_parts


def combo_parts_for_mix_name(
    *,
    name_base: str = "",
    parent_sku: str = "",
    work_root: Optional[Path] = None,
) -> List[Dict[str, Any]]:
    """互換入口。正はマスタセット構成（combo_parts_from_sheet）。名前中黒は使わない。"""
    _ = name_base  # 呼び出し互換のみ
    sku = str(parent_sku or "").strip()
    if not sku:
        LOG.warning("combo_parts_for_mix_name: parent_sku 空のため構成を組めない")
        return []
    return combo_parts_from_sheet(parent_sku=sku, work_root=work_root)


def _is_exact_sku_unit_file(path: Path, parent_sku: str) -> bool:
    """ファイル名が {親SKU}_単体 と完全一致。接頭辞（oya ⊂ oya2）は落とす。"""
    sku = str(parent_sku or "").strip()
    if not sku:
        return False
    return path.stem.lower() == f"{sku}_単体".lower()


def _sku_unit_core_key(parent_sku: str) -> str:
    """親SKUから JAN＋枝（例: 4906283045515-oya）。ベンダー接頭辞の違いを吸収する。"""
    s = str(parent_sku or "").strip()
    m = re.search(r"(\d{10,14}(?:-oya\d*)?)$", s, re.I)
    return (m.group(1) if m else "").lower()


def _is_sku_unit_file(path: Path, parent_sku: str) -> bool:
    """{親SKU}_単体、または同一 JAN＋oya枝でベンダーだけ違う *_単体。"""
    if _is_exact_sku_unit_file(path, parent_sku):
        return True
    core = _sku_unit_core_key(parent_sku)
    if not core:
        return False
    stem = path.stem.lower()
    if stem == f"{core}_単体":
        return True
    return stem.endswith(f"-{core}_単体")


def _find_sku_bag_file(parent_sku: str, *, work_root: Optional[Path] = None) -> Optional[Path]:
    sku = str(parent_sku or "").strip()
    if not sku:
        return None
    exact: List[Path] = []
    core_hits: List[Path] = []
    for folder in _amazon_bag_dirs(work_root):
        if not folder.is_dir():
            continue
        for p in folder.iterdir():
            if not p.is_file() or p.suffix.lower() not in {".png", ".jpg", ".jpeg", ".webp"}:
                continue
            if _is_exact_sku_unit_file(p, sku):
                exact.append(p)
            elif _is_sku_unit_file(p, sku):
                core_hits.append(p)
    hits = exact or core_hits
    if not hits:
        LOG.info("sku bag exact miss parentSku=%s", sku)
        return None
    picked = sorted(hits, key=lambda p: p.stat().st_mtime, reverse=True)[0]
    how = "exact" if exact else "core"
    LOG.info("sku bag %s parentSku=%s file=%s", how, sku, picked.name)
    return picked


def _scan_out_bag_by_name_base(name_base: str, *, script_dir: Optional[Path] = None) -> Optional[Path]:
    want = normalize_name_base_key(name_base)
    if not want:
        return None
    out = _script_dir(script_dir) / "_out"
    if not out.is_dir():
        return None
    dirs = list(out.glob("mix_run_*")) + list(out.glob("theme_new_run_*"))
    for d in dirs:
        if not d.is_dir():
            continue
        if normalize_name_base_key(_folder_name_base(d)) != want:
            continue
        for cand in (
            d / "01_refs" / "PACKAGE_TRUTH.jpg",
            d / "01_refs" / "PACKAGE_TRUTH.png",
            d / "PACKAGE_TRUTH.jpg",
            d / "PACKAGE_TRUTH.png",
        ):
            if cand.is_file():
                return cand
    return None


def resolve_package_bag_for_name(
    *,
    name_base: str,
    parent_sku: str = "",
    jan_for_folder: str = "",
    bag_path: str = "",
    work_root: Optional[Path] = None,
    script_dir: Optional[Path] = None,
) -> Optional[Path]:
    """袋の正本。判定は商品名ベース／親SKU。JANフォルダの PACKAGE_TRUTH は名前が一致するときだけ。"""
    nb = str(name_base or "").strip()
    explicit = Path(str(bag_path or "").strip()) if str(bag_path or "").strip() else None
    if explicit is not None and explicit.is_file() and _bag_name_ok(explicit, nb):
        return explicit

    sku_file = _find_sku_bag_file(parent_sku, work_root=work_root)
    if sku_file is not None and _bag_name_ok(sku_file, nb):
        return sku_file

    if parent_sku:
        try:
            from b5_parent_shared import load_parent_shared, package_truth_rel_for_parent

            rel = package_truth_rel_for_parent(parent_sku, work_root=work_root)
            hit = _resolve_bag_rel(rel, work_root=work_root) if rel else None
            if hit is not None and _bag_name_ok(hit, nb):
                return hit
            by = (load_parent_shared(work_root).get("byParentSku") or {})
            want = normalize_name_base_key(nb)
            if want:
                for row in by.values():
                    if not isinstance(row, dict):
                        continue
                    if normalize_name_base_key(str(row.get("nameBase") or "")) != want:
                        continue
                    hit2 = _resolve_bag_rel(str(row.get("packageTruthRel") or ""), work_root=work_root)
                    if hit2 is not None and _bag_name_ok(hit2, nb):
                        return hit2
        except Exception as e:
            LOG.warning("parent_shared bag: %s", e)

    scanned = _scan_out_bag_by_name_base(nb, script_dir=script_dir)
    if scanned is not None and _bag_name_ok(scanned, nb):
        return scanned

    jan_d = "".join(ch for ch in str(jan_for_folder) if ch.isdigit()) or str(jan_for_folder).strip()
    if jan_d:
        tn = theme_new_dir_for_jan(jan_d, script_dir=script_dir)
        if tn.is_dir() and normalize_name_base_key(_folder_name_base(tn)) == normalize_name_base_key(nb):
            for cand in (tn / "01_refs" / "PACKAGE_TRUTH.jpg", tn / "01_refs" / "PACKAGE_TRUTH.png"):
                if cand.is_file() and _bag_name_ok(cand, nb):
                    return cand
            meta = load_food_canons(tn)
            truth = Path(str(meta.get("truth") or ""))
            if truth.is_file() and _bag_name_ok(truth, nb):
                return truth
    return None


def _main_theme_num(theme_id: str) -> int:
    """MAIN themeId（T4／N1）→ §7.12 mode 用の番号。N1 は identity（4）。"""
    t = str(theme_id or "").strip().upper()
    if t.startswith("N"):
        return 4
    m = re.match(r"^T?(\d+)$", t)
    if m:
        return int(m.group(1))
    return 4


def resolve_main_food_refs(
    jan: str,
    theme_id: str,
    *,
    require_human: bool = True,
    script_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """MAIN 用: サブと同じ theme_new_run_{jan} の個品見本を読む（§7.14）。

    人間確定（confirmedBy=human）を正とする。未確定／無しは packageOnly（正本のみ）。
    """
    tn = theme_new_dir_for_jan(jan, script_dir=script_dir)
    meta = load_food_canons(tn) if tn.is_dir() else {}
    confirmed = str(meta.get("confirmedBy") or "")
    package_only = bool(meta.get("packageOnly")) or meta.get("identityStatus") == "absent"
    base: Dict[str, Any] = {
        "themeNewDir": str(tn) if tn.is_dir() else "",
        "confirmedBy": confirmed,
        "ok": False,
        "reason": "",
    }
    if not tn.is_dir():
        base.update(
            {
                "mode": "package",
                "identity": None,
                "sizzle": None,
                "wrap": None,
                "secondary": [],
                "packageOnly": True,
                "reason": "theme_new_run なし",
            }
        )
        return base
    if require_human and confirmed != "human" and not package_only:
        base.update(
            {
                "mode": "package",
                "identity": None,
                "sizzle": None,
                "wrap": None,
                "secondary": [],
                "packageOnly": True,
                "reason": f"未確定(confirmedBy={confirmed or 'empty'})",
            }
        )
        return base
    tid = _main_theme_num(theme_id)
    refs = resolve_food_refs_for_theme(tn, tid, need_package=False)
    # MAIN: 中身の色・厚みの正は常に IDENTITY（あるとき）。SIZZLE は mode に応じ追加
    ident = existing_canon_path(tn)
    if ident and ident.is_file() and not refs.get("packageOnly"):
        refs["identity"] = ident
    # WRAP はテーマ mode に依存せず、ファイルがあれば載せる
    wrap = existing_wrap_path(tn)
    if wrap and wrap.is_file() and not refs.get("packageOnly"):
        refs["wrap"] = wrap
        refs["useWrap"] = True
    refs.update(base)
    refs["ok"] = bool(
        refs.get("identity")
        or refs.get("sizzle")
        or refs.get("wrap")
        or refs.get("packageOnly")
        or package_only
    )
    if refs.get("identity") or refs.get("sizzle") or refs.get("wrap"):
        refs["reason"] = "human_canon" if confirmed == "human" else (confirmed or "canon")
    elif refs.get("packageOnly"):
        refs["reason"] = "package_only"
    return refs


def load_part_bundle(
    *,
    name_base: str,
    jan_for_folder: str = "",
    parent_sku: str = "",
    bag_path: str = "",
    script_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """構成1件。判定は商品名ベース。jan は見本フォルダの場所だけ。"""
    nb = str(name_base or "").strip()
    jan_d = "".join(ch for ch in str(jan_for_folder) if ch.isdigit()) or str(jan_for_folder).strip()
    tn = theme_new_dir_for_jan(jan_d, script_dir=script_dir) if jan_d else None
    bag = resolve_package_bag_for_name(
        name_base=nb,
        parent_sku=parent_sku,
        jan_for_folder=jan_d,
        bag_path=bag_path,
        script_dir=script_dir,
    )
    ident = None
    sizzle = None
    wrap = None
    secondary: List[Path] = []
    confirmed = ""
    package_only = False
    if tn is not None and tn.is_dir():
        meta = load_food_canons(tn)
        confirmed = str(meta.get("confirmedBy") or "")
        package_only = bool(meta.get("packageOnly"))
        ident = existing_canon_path(tn)
        sizzle = existing_sizzle_path(tn)
        wrap = existing_wrap_path(tn)
        secondary = secondary_paths(tn)
    ok = confirmed == "human" and (package_only or bool(ident and Path(ident).is_file()))
    return {
        "nameBase": nb,
        "parentSku": str(parent_sku or "").strip(),
        "jan": jan_d,
        "themeNewDir": str(tn) if tn and tn.is_dir() else "",
        "confirmedBy": confirmed,
        "ok": ok,
        "packageOnly": package_only,
        "identity": ident if ident and Path(ident).is_file() else None,
        "sizzle": sizzle if sizzle and Path(sizzle).is_file() else None,
        "wrap": wrap if wrap and Path(wrap).is_file() else None,
        "secondary": [p for p in secondary if p.is_file()],
        "bag": bag if bag is not None and bag.is_file() else None,
        "label": (
            "確定（正本のみ）"
            if ok and package_only
            else ("確定（個品あり）" if ok else "未確定")
        ),
    }


def food_unit_prompt_block_main(
    *,
    has_identity: bool,
    has_sizzle: bool = False,
    has_secondary: bool = False,
    has_wrap: bool = False,
    has_piece_set: bool = False,
    food_mode: str = "identity",
    mix_parts: Optional[Sequence[str]] = None,
    mix_part_bags: bool = False,
    stamp_layers: bool = False,
) -> str:
    """MAIN プロンプト用。stamp_layers 時は個品を描かせない（画素後載せ）。

    has_piece_set=True のとき個品の正は PIECE＋言語化、IDENTITY／CANON は盛りのみ
    （§4.5.10bis-T。旧「IDENTITY唯一」文と衝突させない）。
    """
    names = [str(x).strip() for x in (mix_parts or []) if str(x).strip()]
    if stamp_layers:
        lines = [
            "【層・必須】袋と個品は後載せ。AIは描き直すな。",
            "- 皿の上の商品（個品）を描くな。空皿、またはごく薄い空皿だけ。",
            "- 見本に無い形・厚みを皿にも背景にも足すな。",
            "- 袋を描くな。左下に袋用の空きを残せ（IMAGE_LAYOUT_ZONES）。",
        ]
        if len(names) >= 2:
            lines.append("- 構成は " + "、".join(names) + "。皿の上で混ぜて一色にしない（中身は後載せ）。")
        return "\n".join(lines) + "\n"
    if has_piece_set:
        if has_wrap:
            wrap_rule = (
                "開封後の中身の色・縁・厚み・外形は IMAGE_FOOD_UNIT_PIECE＋一片言語化が正。"
                "IDENTITY／CANON は盛り配置・量感のみ。小皿の主役は開封後。"
                "添える料理の上に載せる具も一片セットの見た目そのもの。一般的なトッピングや別形に寄せるな。"
                "個包装の見本（IMAGE_FOOD_UNIT_WRAP／WRAP_i）があるときは個包装を出せ。"
                "主役は開封後の個品。個包装は脇役。"
                "置き方は皿の横か後ろに1〜2個のみ（3個以上・山は禁止。小皿の開封後に混ぜるな）。"
                "カメラ最近傍（手前レイヤ）は開封後の個品だけ。個包装はそれより手前に出すな。"
                "個包装の密着・空気の有無・材質の硬さやわらかさ・縁・形は WRAP 見本と機械読みに忠実。"
                "見本に無い空気だまりを作るな。見本が空気なしなら空気だまり禁止。"
                "見本と違う別素材の小袋を創作するな。"
                "WRAP写真に写る大袋・背景・手元は無視。大袋の色・印字・形は IMAGE_PACKAGE_PART／TRUTH だけ。"
                "WRAPの見た目で大袋を描き変えるな。WRAPが無い構成の個包装は描くな。"
            )
        else:
            wrap_rule = (
                "小皿に描く個品は一片セット写真＋言語化どおり開封後の中身だけ。"
                "個包装・小袋・包み紙を新たに描くな。"
                "包装の見本が無いときは包装を描くな。"
                "正本袋の窓以外の包装見た目を創作するな。"
            )
        pixel_ban = f"""【個品の使い方・必須／一片セットあり】
見本写真を四角く貼るな。場面の中に描け。
厚み・色・質感・1片の大きさ・縁・血合いは IMAGE_FOOD_UNIT_PIECE（一片セット）＋一片言語化が正。
IDENTITY／CANON は盛り合わせ・重なり・量感の【配置だけ】（個品の色・縁・厚み・血合い・1片サイズの正にしない）。
IDENTITY盛り写真＋IDENTITY盛り言語化の重なり・密度・ばらけ・山の高さを保て。
言語化が粗い項目は IDENTITY 写真を見ろ。一列標本・円形均一・過剰山・空疎は禁止。
料理場面では小皿複製は必須でないが、具の重なり・密度・ばらけの型は IDENTITY盛りを移植せよ。
キャッチやセールストークの語で厚みや色や大きさを変えるな。
小皿の盛りも、料理の上に載せる具も、トッピングも、一片セットと同じ1片サイズ・同じ見た目にせよ（拡大・縮小・別形のトッピング禁止）。
場面の暖色・夜景で下層を黒く落とすな（影の向きだけ場面に合わせる）。
副見本（SECONDARY）は盛り付け（置き方・重なり）の参考だけ。1片の形・大きさ・厚み・色には使うな。
見本に無い形・厚みを、皿にも背景にも描くな。
原料の雰囲気は背景に限り可。個品の皿に混ぜるな。見本に無い形は原料イメージにも出すな。
白背景の単一片標本は禁止。IDENTITYを一片標本のまま場面の正にするな。
{wrap_rule}
"""
        lock_head = (
            "【食材ユニット正本（一片セットあり・必須）】\n"
            "目的: テーマが違っても、一片セットと同じ1片に見えるようにする。\n"
            "場面は配置・量感・料理の脇役だけ変え、個品は一片ロックを崩すな。\n"
            "\n"
            "■ 正本の優先順\n"
            "1) IMAGE_FOOD_UNIT_PIECE_* ＋一片言語化 → 開封後の色・縁・筋・厚み・外形・血合い・1片サイズの正\n"
            "2) IMAGE_FOOD_UNIT_CANON／IDENTITY → 盛り配置・量感のみ（個品の正にしない）\n"
            "3) IMAGE_FOOD_UNIT_WRAP（個包装・任意）→ 個包装だけの見た目の正\n"
            "4) IMAGE_FOOD_UNIT_SECONDARY_* → 盛り付け参考のみ\n"
            "5) 無いとき → IMAGE_PACKAGE_TRUTH の窓・抜き絵\n"
            "6) IMAGE_COMP_* はレイアウトのみ。食材の正本にしない\n"
        )
    else:
        if has_wrap:
            wrap_rule = (
                "開封後の中身は IDENTITY（IMAGE_FOOD_UNIT_CANON／IDENTITY_i）が正。小皿の主役は開封後。"
                "添える料理の上に載せる具も IDENTITY の1片そのもの。一般的なトッピングや別形に寄せるな。"
                "個包装の見本（IMAGE_FOOD_UNIT_WRAP／WRAP_i）があるときは個包装を出せ。"
                "主役は開封後の個品。個包装は脇役。"
                "置き方は皿の横か後ろに1〜2個のみ（3個以上・山は禁止。小皿の開封後に混ぜるな）。"
                "カメラ最近傍（手前レイヤ）は開封後の個品だけ。個包装はそれより手前に出すな。"
                "個包装の密着・空気の有無・材質の硬さやわらかさ・縁・形は WRAP 見本と機械読みに忠実。"
                "見本に無い空気だまりを作るな。見本が空気なしなら空気だまり禁止。"
                "見本と違う別素材の小袋を創作するな。"
                "WRAP写真に写る大袋・背景・手元は無視。大袋の色・印字・形は IMAGE_PACKAGE_PART／TRUTH だけ。"
                "WRAPの見た目で大袋を描き変えるな。WRAPが無い構成の個包装は描くな。"
            )
        else:
            wrap_rule = (
                "小皿に描く個品は見本どおり開封後の中身だけ。"
                "個包装・小袋・包み紙を新たに描くな。"
                "包装の見本が無いときは包装を描くな。"
                "正本袋の窓以外の包装見た目を創作するな。"
            )
        pixel_ban = f"""【個品の使い方・必須】
見本写真を四角く貼るな。場面の中に描け。
厚み・色・質感・1片の大きさは主見本（IDENTITY）だけ。キャッチやセールストークの語で厚みや色や大きさを変えるな。
添える料理・トッピングに載せる具も主見本の1片そのもの。一般的な別食材の形に置き換えるな。
色と明るさも主見本を優先。場面の暖色・夜景で下層を黒く落とすな（影の向きだけ場面に合わせる）。
副見本（SECONDARY）は盛り付け（置き方・重なり）の参考だけ。1片の形・大きさ・厚み・色には使うな。副が無ければ盛り付けも主見本。
見本に無い形・厚みを、皿にも背景にも描くな。
原料の雰囲気は背景に限り可。個品の皿に混ぜるな。見本に無い形は原料イメージにも出すな。
{wrap_rule}
"""
        lock_head = FOOD_UNIT_LOCK_JA.strip()
    sku_re = re.compile(r"(?i)(ishi-|sanky-|sku|oya\d+|親SKU|子SKU|[A-Za-z]+-\d{10,})")
    names = [n for n in names if n and not sku_re.search(n)]
    if len(names) >= 2:
        identity_truth = (
            "IMAGE_FOOD_UNIT_IDENTITY_i は構成 i の盛り配置・量感の参考。"
            "個品の色・縁・厚みは IMAGE_FOOD_UNIT_PIECE（＋言語化）が正。"
            "混ぜて一色にしない。貼るな。セールストークに合わせるな。"
            if has_piece_set
            else "IMAGE_FOOD_UNIT_IDENTITY_i は構成 i の厚み・色・質感・1片の大きさの正。混ぜて一色にしない。貼るな。セールストークに合わせるな。"
        )
        lines = [
            lock_head,
            "",
            pixel_ban.strip(),
            "",
            "【組み合わせ・複数個品】",
            f"- 構成は商品名ベースで {len(names)} 種（この順が正）: " + "、".join(names),
            f"- {identity_truth}",
            (
                "- 種ごとに分けて置け。一山に混ぜるな。各構成の一片セット色を維持。"
                if has_piece_set
                else "- 種ごとに分けて置け。一山に混ぜるな。各 identity の色を維持。"
            ),
            "- IMAGE_PACKAGE_SCALE は大きさの見本だけ。完成画にその写真を置くな。",
            "- IMAGE_PACKAGE_PART_i を場面の中に描け。各袋の縦横は、ファイル実測の縦÷横。置いた面と同じ距離のピントに乗せろ。",
            "- 袋どうしの大きさの対比だけ SCALE と同じ。奥行きは盛り付けの後ろ。上から貼るな。",
            "【袋と個品の並び・必須】",
            "- 袋の主軸（通常は左→右）は PART_1 → PART_2 → … の順。構成名の順と一致させろ。入れ替えるな。",
            "- 個品を小皿・別皿に分けるときも、同じ主軸・同じ順で IDENTITY_1 → _2 → … を置け。袋の並びと個品の並びを連動させろ（例: 左のかつお袋の手前／同列にかつお個品）。",
            "- 「種類が分かればよい」だけで順を崩すな。袋だけ正しい順で個品が違う順は不合格。",
        ]
        if has_wrap:
            lines.append(
                "- WRAPがある構成は個包装を1〜2個出す（皿の横か後ろのみ。小皿に混ぜるな。カメラ最近傍は開封後。個品より前に出すな）。山にしない。主役は開封後の IDENTITY。正は IMAGE_FOOD_UNIT_WRAP_i。無い構成は包装を描くな。"
            )
            lines.append(
                "- WRAP画像の大袋・背景は使うな。大袋は PART_i のみ。個包装だけ WRAP に合わせろ。"
            )
        else:
            lines.append(
                "- 小皿は開封後の中身だけ。個包装を創作するな。包装見本が無いときは包装を描くな。"
            )
        if mix_part_bags:
            lines.append(
                "- 描く袋は IMAGE_PACKAGE_PART_i。取り違えるな。見本の袋を後ろに残すな。"
            )
        if has_secondary:
            lines.append(
                "- IMAGE_FOOD_UNIT_SECONDARY_* は盛り付け（置き方・重なり）の参考だけ。"
                + (
                    "1片の形・大きさ・厚み・色・質感は一片セット。"
                    if has_piece_set
                    else "1片の形・大きさ・厚み・色・質感は IDENTITY。"
                )
                + "主見本と同じ盛り付けのコピーにはしない。"
            )
        if has_sizzle:
            lines.append(
                "- IMAGE_FOOD_UNIT_SIZZLE_i は食欲状態。形・厚みの正にしない。"
            )
        lines.append("- 競合レイアウト・吹き出し・人物のコピーは禁止（個品の見た目のみ借りる）。")
        return "\n".join(lines) + "\n"
    if not has_identity and not has_sizzle and not has_piece_set:
        return """【個品見本・無し】
IMAGE_FOOD_UNIT_* は添付なし。中身は IMAGE_PACKAGE_TRUTH の窓・抜き絵のみを参考にせよ。
見本に無い形・厚み・別色の塗り替え・別商品化は禁止。
"""
    lines = [
        lock_head,
        "",
        pixel_ban.strip(),
        "",
        "【MAIN・個品見本の使い方】",
        f"- 本テーマ foodMode={food_mode}。構図・キャッチは MAIN ルール。個品の見た目だけ見本に合わせる。",
    ]
    if has_piece_set:
        lines.append(
            "- IMAGE_FOOD_UNIT_PIECE_*: 開封後の厚み・色・質感・縁・血合いの正。写真を貼るな。セールストークに合わせるな。"
        )
        if has_identity:
            lines.append(
                "- IMAGE_FOOD_UNIT_CANON: 盛り配置・量感のみ。個品の色・厚みの正にしない。"
            )
    elif has_identity:
        lines.append(
            "- IMAGE_FOOD_UNIT_CANON: 開封後の厚み・色・質感・1片の大きさの正。写真を貼るな。セールストークに合わせるな。"
        )
    if has_wrap:
        lines.append(
            "- IMAGE_FOOD_UNIT_WRAP: 個包装の正。1〜2個出す（皿の横か後ろのみ。小皿に混ぜるな。カメラ最近傍は開封後）。山にしない。主役は開封後。密着・空気・材質・縁は見本と機械読みどおり。無い柄を創作するな。写真内の大袋・背景・手は無視（大袋は PACKAGE のみ）。"
        )
    if has_sizzle:
        lines.append(
            "- IMAGE_FOOD_UNIT_SIZZLE: 食欲状態の参考可。厚み・色・質感の正にしない。"
        )
    if has_secondary:
        lines.append(
            "- IMAGE_FOOD_UNIT_SECONDARY_*: 盛り付け（置き方・重なり）の参考だけ。"
            + (
                "1片の形・大きさ・厚み・色・質感は一片セット。"
                if has_piece_set
                else "1片の形・大きさ・厚み・色・質感は IDENTITY。"
            )
            + "主見本と同じ盛り付けのコピーにはしない。"
        )
    lines.append("- 競合レイアウト・吹き出し・人物のコピーは禁止（個品の見た目のみ借りる）。")
    return "\n".join(lines) + "\n"


def secondary_paths(root: Path) -> List[Path]:
    refs = Path(root) / "01_refs"
    if not refs.is_dir():
        return []
    out: List[Path] = []
    for i in range(1, MAX_SECONDARY + 1):
        for ext in (".jpg", ".jpeg", ".png", ".webp"):
            p = refs / f"{SECONDARY_STEM}_{i}{ext}"
            if p.is_file():
                out.append(p)
                break
    return out


def load_canon_meta(root: Path) -> Dict[str, Any]:
    p = Path(root) / META_NAME
    if not p.is_file():
        p = Path(root) / "01_refs" / META_NAME
    if not p.is_file():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def load_candidates_meta(root: Path) -> Dict[str, Any]:
    p = Path(root) / CANDIDATES_NAME
    if not p.is_file():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


# --- 構図類似の除外（同一写真・文字パネル違い・ほぼ同一レイアウト） ---
AHASH_SIZE = 16
# フル画像: ほぼ同一（#4/#5）
FULL_HAMMING_MAX = 12
# 領域同士: 同構図
REGION_HAMMING_MAX = 10
# 片方がフル、もう片方が左／中央に同じ写真（縦書きパネル付き）
CROSS_HAMMING_MAX = 12


def _hamming(a: int, b: int) -> int:
    return (a ^ b).bit_count()


def _ahash_from_image(im: Any, size: int = AHASH_SIZE) -> int:
    from PIL import Image as _Image

    g = im.convert("L").resize((size, size), _Image.Resampling.LANCZOS)
    pixels = list(g.getdata())
    avg = sum(pixels) / max(1, len(pixels))
    bits = 0
    for i, px in enumerate(pixels):
        if px >= avg:
            bits |= 1 << i
    return bits


def composition_signatures(path: Path) -> Optional[Dict[str, int]]:
    """フル／左半分／中央／上半分の aHash。読めない画像は None。"""
    try:
        from PIL import Image

        im = Image.open(path).convert("RGB")
        w, h = im.size
        if w < 8 or h < 8:
            return None
        left = im.crop((0, 0, max(1, int(w * 0.55)), h))
        center = im.crop(
            (int(w * 0.12), int(h * 0.12), max(1, int(w * 0.88)), max(1, int(h * 0.88)))
        )
        top = im.crop((0, 0, w, max(1, int(h * 0.55))))
        return {
            "full": _ahash_from_image(im),
            "left": _ahash_from_image(left),
            "center": _ahash_from_image(center),
            "top": _ahash_from_image(top),
        }
    except Exception as e:
        LOG.debug("composition_signatures fail %s: %s", path, e)
        return None


def compositions_similar(
    a: Dict[str, int], b: Dict[str, int]
) -> bool:
    if _hamming(a["full"], b["full"]) <= FULL_HAMMING_MAX:
        return True
    for key in ("left", "center", "top"):
        if _hamming(a[key], b[key]) <= REGION_HAMMING_MAX:
            return True
    # 片方が「写真＋文字」、もう片方が同じ写真の寄り
    if _hamming(a["left"], b["full"]) <= CROSS_HAMMING_MAX:
        return True
    if _hamming(b["left"], a["full"]) <= CROSS_HAMMING_MAX:
        return True
    if _hamming(a["center"], b["full"]) <= CROSS_HAMMING_MAX:
        return True
    if _hamming(b["center"], a["full"]) <= CROSS_HAMMING_MAX:
        return True
    return False


def _prefer_cleaner_comp(a: Path, b: Path, sig_a: Dict[str, int], sig_b: Dict[str, int]) -> Path:
    """
    同構図なら「食材寄り」を残す。
    left≈full の方が文字パネル少なめ、とみなす。
    """
    fill_a = _hamming(sig_a["left"], sig_a["full"])
    fill_b = _hamming(sig_b["left"], sig_b["full"])
    if fill_a != fill_b:
        return a if fill_a < fill_b else b
    # 縦長・横長で極端に文字が多い方を捨てる → 面積が小さい方を優先しない、正方形に近い方
    try:
        from PIL import Image

        wa, ha = Image.open(a).size
        wb, hb = Image.open(b).size
        ra = max(wa, ha) / max(1, min(wa, ha))
        rb = max(wb, hb) / max(1, min(wb, hb))
        if abs(ra - rb) > 0.15:
            return a if ra < rb else b
    except Exception:
        pass
    return a


def dedupe_comps_by_composition(
    comps: Sequence[Path],
) -> Tuple[List[Path], Dict[str, Any]]:
    """
    同構図・ほぼ同一画像を1枚に潰す。戻り: (残すパス, メタ)。
    """
    kept: List[Path] = []
    kept_sigs: List[Dict[str, int]] = []
    dropped: List[Dict[str, str]] = []
    for p in comps:
        p = Path(p)
        if not p.is_file():
            continue
        sig = composition_signatures(p)
        if sig is None:
            kept.append(p)
            continue
        dup_at = None
        for i, ks in enumerate(kept_sigs):
            if compositions_similar(sig, ks):
                dup_at = i
                break
        if dup_at is None:
            kept.append(p)
            kept_sigs.append(sig)
            continue
        winner = _prefer_cleaner_comp(kept[dup_at], p, kept_sigs[dup_at], sig)
        loser = p if winner == kept[dup_at] else kept[dup_at]
        if winner != kept[dup_at]:
            kept[dup_at] = winner
            kept_sigs[dup_at] = composition_signatures(winner) or sig
        dropped.append(
            {
                "dropped": loser.name,
                "kept": winner.name,
                "reason": "composition_similar",
            }
        )
        LOG.info("dedupe composition drop=%s keep=%s", loser.name, winner.name)
    return kept, {"dropped": dropped, "nIn": len(list(comps)), "nOut": len(kept)}


def dedupe_ranked_rows(
    rows: Sequence[Dict[str, Any]], *, max_keep: int
) -> List[Dict[str, Any]]:
    """スコア順の行から同構図を落とし、最大 max_keep 枚。"""
    ordered = sorted(rows, key=lambda r: int(r.get("score") or 0), reverse=True)
    out: List[Dict[str, Any]] = []
    out_sigs: List[Dict[str, int]] = []
    for r in ordered:
        p = Path(str(r.get("path") or ""))
        if not p.is_file():
            continue
        sig = composition_signatures(p)
        if sig is not None and any(compositions_similar(sig, ks) for ks in out_sigs):
            LOG.info("dedupe ranked skip %s", p.name)
            continue
        out.append(r)
        if sig is not None:
            out_sigs.append(sig)
        if len(out) >= max_keep:
            break
    for i, r in enumerate(out):
        r["suggestedPrimary"] = i == 0
        r["index"] = i
    return out


def _mime(p: Path) -> str:
    s = p.suffix.lower()
    if s == ".png":
        return "image/png"
    if s == ".webp":
        return "image/webp"
    return "image/jpeg"


def _b64_part(p: Path) -> Dict[str, Any]:
    return {
        "inline_data": {
            "mime_type": _mime(p),
            "data": base64.b64encode(p.read_bytes()).decode("ascii"),
        }
    }


def _parse_ranked(text: str, n: int) -> Optional[List[Dict[str, Any]]]:
    if not text:
        return None
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        return None
    try:
        obj = json.loads(m.group(0))
    except Exception:
        return None
    if not isinstance(obj, dict):
        return None
    ranked = obj.get("ranked")
    if not isinstance(ranked, list):
        # 旧形式 bestIndex 互換
        try:
            bi = int(obj.get("bestIndex"))
            sc = int(obj.get("score") or 0)
        except (TypeError, ValueError):
            return None
        if bi < 0:
            return []
        return [{"index": bi, "score": sc, "reasonJa": str(obj.get("reasonJa") or "")[:240]}]
    out: List[Dict[str, Any]] = []
    seen = set()
    for it in ranked:
        if not isinstance(it, dict):
            continue
        try:
            idx = int(it.get("index"))
            sc = int(it.get("score") or 0)
        except (TypeError, ValueError):
            continue
        if idx < 0 or idx >= n or idx in seen:
            continue
        seen.add(idx)
        out.append(
            {
                "index": idx,
                "score": sc,
                "reasonJa": str(it.get("reasonJa") or "")[:240],
            }
        )
    return out


def rank_comps_vs_truth(
    truth: Path,
    comps: Sequence[Path],
    *,
    top_k: int = DEFAULT_TOP_K,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    正本中身に近い順に競合を並べ、上位 top_k 件を返す。
    各要素: {index, path, name, score, reasonJa}
    """
    comps = [Path(p) for p in comps if Path(p).is_file()]
    truth = Path(truth)
    if not truth.is_file() or not comps:
        return [], {"ok": False, "reason": "no_truth_or_comps"}

    from gemini_image import load_api_key, make_client
    from sub_image_intent import TEXT_MODELS

    try:
        client = make_client(load_api_key())
    except SystemExit:
        LOG.warning("Gemini key missing — food canon rank fallback")
        return [], {"ok": False, "reason": "no_gemini_key"}

    k = max(1, min(int(top_k), len(comps)))
    prompt = (
        "食品EC用です。1枚目は正本パッケージ（IMAGE_PACKAGE_TRUTH）。"
        "続く画像は競合サブです。\n"
        "正本の「中身・窓・抜き絵」の食材（色・厚み・1片の形）に近い順に、"
        f"最大{k}枚までランキングしてください。\n"
        "袋デザイン一致は不要。食材見た目一致を最優先。\n"
        "【構図の重複禁止・必須】ほぼ同一の写真／同じレイアウトテンプレ"
        "（例: 同じ盛り＋袋インセット＋赤丸バッジ、同じボウル寄りを文字パネル付きで使い回し）は"
        "ranked に1つだけ。似た盛り・似た寄りが複数あるときは最も見本向きの1枚だけ残し、"
        "残りは落とす。盛り方・背景・カード型が違うものを優先して多様に選ぶ。\n"
        "別物・食材なし・説明カードのみは ranked に入れない。\n"
        f"競合 index は 0..{len(comps)-1}。\n"
        'JSONのみ: {"ranked":[{"index":0,"score":0-100,"reasonJa":"短い理由"}]}\n'
        "score は正本中身との一致度。"
    )
    parts: List[Any] = [{"text": prompt}, _b64_part(truth)]
    for i, c in enumerate(comps):
        parts.append({"text": f"COMP_INDEX={i}"})
        parts.append(_b64_part(c))

    last_err = ""
    for mid in TEXT_MODELS:
        if not mid:
            continue
        try:
            resp = client.models.generate_content(
                model=mid,
                contents=[{"role": "user", "parts": parts}],
            )
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                try:
                    text = resp.candidates[0].content.parts[0].text
                except Exception:
                    text = ""
            ranked = _parse_ranked(text, len(comps))
            if ranked is None:
                raise RuntimeError("parse failed: " + text[:200])
            # score 降順で整え、閾値未満は落とす（ただし最低枚数を確保）
            ranked = sorted(ranked, key=lambda x: int(x.get("score") or 0), reverse=True)
            good = [r for r in ranked if int(r.get("score") or 0) >= 40]
            if len(good) < MIN_SHOW:
                # 足りなければ残りをスコア順で補充
                rest = [r for r in ranked if r not in good]
                good = good + rest
            good = good[:k]
            # それでも足りなければ未ランクを末尾補充（人間が落とせるように）
            used = {int(r["index"]) for r in good}
            if len(good) < MIN_SHOW:
                for i, p in enumerate(comps):
                    if i in used:
                        continue
                    good.append(
                        {
                            "index": i,
                            "score": 30,
                            "reasonJa": "機械ランク外の補充候補（要人間判断）",
                        }
                    )
                    used.add(i)
                    if len(good) >= MIN_SHOW:
                        break
            rows: List[Dict[str, Any]] = []
            for r in good[:k]:
                idx = int(r["index"])
                p = comps[idx]
                rows.append(
                    {
                        "index": idx,
                        "path": str(p.resolve()),
                        "name": p.name,
                        "score": int(r.get("score") or 0),
                        "reasonJa": str(r.get("reasonJa") or ""),
                        "suggestedPrimary": len(rows) == 0,
                    }
                )
            meta = {
                "ok": bool(rows),
                "model": mid,
                "nComp": len(comps),
                "topK": k,
                "nShown": len(rows),
            }
            LOG.info("food canon rank n=%s model=%s", len(rows), mid)
            return rows, meta
        except Exception as e:
            last_err = str(e)
            LOG.warning("food canon rank model=%s failed: %s", mid, e)
    return [], {"ok": False, "reason": "all_models_failed", "error": last_err}


def pick_best_comp_vs_truth(
    truth: Path,
    comps: Sequence[Path],
) -> Tuple[Optional[Path], Dict[str, Any]]:
    """後方互換: ランク1位を返す。"""
    rows, meta = rank_comps_vs_truth(truth, comps, top_k=1)
    if not rows:
        return None, meta
    top = rows[0]
    if int(top.get("score") or 0) < 45:
        meta = {**meta, "ok": False, "reason": "no_suitable_comp", "score": top.get("score")}
        return None, meta
    return Path(top["path"]), {
        **meta,
        "ok": True,
        "bestIndex": top["index"],
        "score": top["score"],
        "reasonJa": top.get("reasonJa") or "",
        "chosen": top["path"],
    }


def resolve_comps_for_dir(out_dir: Path) -> List[Path]:
    refs = Path(out_dir) / "01_refs"
    if not refs.is_dir():
        return []
    out: List[Path] = []
    for p in sorted(refs.iterdir()):
        if not p.is_file():
            continue
        if p.suffix.lower() not in (".jpg", ".jpeg", ".png", ".webp"):
            continue
        n = p.name.upper()
        if (
            n.startswith(CANON_STEM)
            or n.startswith(SECONDARY_STEM)
            or n.startswith(SIZZLE_STEM)
            or n.startswith("PACKAGE_TRUTH")
        ):
            continue
        out.append(p)
    return out


def write_candidates_shortlist(
    *,
    out_dir: Path,
    truth: Path,
    comps: Sequence[Path],
    top_k: int = DEFAULT_TOP_K,
    force: bool = False,
) -> Dict[str, Any]:
    """ゲート用: 候補5〜6枚を food_unit_candidates.json に書く。同構図は除外。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    cand_path = out_dir / CANDIDATES_NAME
    if not force and cand_path.is_file():
        old = load_candidates_meta(out_dir)
        if old.get("candidates"):
            LOG.info("reuse candidates %s", cand_path)
            return old
    unique_comps, dedupe_meta = dedupe_comps_by_composition(list(comps))
    # ランクは少し多めに取り、後段で同構図を落として top_k に揃える
    rank_n = min(len(unique_comps), max(top_k + 4, top_k))
    rows, rank_meta = rank_comps_vs_truth(Path(truth), unique_comps, top_k=rank_n)
    # 機械失敗時は先頭から機械なし候補
    if not rows:
        for i, p in enumerate(list(unique_comps)[: max(MIN_SHOW, top_k)]):
            rows.append(
                {
                    "index": i,
                    "path": str(Path(p).resolve()),
                    "name": Path(p).name,
                    "score": 0,
                    "reasonJa": "機械ランク失敗のため順不同提示",
                    "suggestedPrimary": i == 0,
                }
            )
    rows = dedupe_ranked_rows(rows, max_keep=top_k)
    # 不足時は未採用の unique から補充（同構図チェック付き）
    if len(rows) < MIN_SHOW:
        used = {str(r.get("name") or "") for r in rows}
        for p in unique_comps:
            if p.name in used:
                continue
            probe = {
                "index": len(rows),
                "path": str(p.resolve()),
                "name": p.name,
                "score": 25,
                "reasonJa": "同構図除外後の補充候補",
                "suggestedPrimary": False,
            }
            merged = dedupe_ranked_rows(list(rows) + [probe], max_keep=top_k)
            if len(merged) > len(rows):
                rows = merged
                used.add(p.name)
            if len(rows) >= MIN_SHOW:
                break
    payload = {
        "builtAt": datetime.now(timezone.utc).isoformat(),
        "truth": str(Path(truth)),
        "status": "awaiting_human",
        "candidates": rows,
        "rankMeta": rank_meta,
        "dedupeMeta": dedupe_meta,
        "rules": {
            "primaryRequired": True,
            "secondaryMax": MAX_SECONDARY,
            "primaryRole": "色・厚み・1片サイズの正",
            "secondaryRole": "盛り・形のみ（色の正にしない）",
            "compositionDedupe": True,
        },
    }
    cand_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def apply_human_selection(
    *,
    out_dir: Path,
    truth: Path,
    primary_name: str = "",
    secondary_names: Optional[Sequence[str]] = None,
    comps: Optional[Sequence[Path]] = None,
    package_only: bool = False,
    identity_name: Optional[str] = None,
    sizzle_name: Optional[str] = None,
    wrap_name: Optional[str] = None,
    name_base: str = "",
    parent_sku: str = "",
) -> Dict[str, Any]:
    """
    人間レ点を確定し FOOD_UNIT_CANON / SIZZLE / WRAP / SECONDARY_* と food_canons.json を書く。
    identity（主）必須。副だけ・未選択は ValueError。
    sizzle／wrap は任意（空／__none__＝無し）。
    parent_sku／name_base は Drive 状態のキー（JAN単独判定を避ける）。
    """
    out_dir = Path(out_dir)
    refs = out_dir / "01_refs"
    refs.mkdir(parents=True, exist_ok=True)
    comps = list(comps) if comps is not None else resolve_comps_for_dir(out_dir)
    by_name = {p.name: Path(p) for p in comps if Path(p).is_file()}
    # 旧コピー掃除
    for p in refs.glob(f"{CANON_STEM}.*"):
        p.unlink(missing_ok=True)
    for p in refs.glob(f"{SECONDARY_STEM}_*"):
        p.unlink(missing_ok=True)
    for p in refs.glob(f"{SIZZLE_STEM}.*"):
        p.unlink(missing_ok=True)
    for p in refs.glob(f"{WRAP_STEM}.*"):
        p.unlink(missing_ok=True)

    identity = str(identity_name if identity_name is not None else primary_name or "").strip()
    sizzle = str(sizzle_name or "").strip()
    if sizzle in ("__none__", "NONE", "none"):
        sizzle = ""
    wrap = str(wrap_name or "").strip()
    if wrap in ("__none__", "NONE", "none"):
        wrap = ""

    secondary_names = [str(x).strip() for x in (secondary_names or []) if str(x).strip()]
    secondary_names = secondary_names[:MAX_SECONDARY]

    package_only = bool(package_only) or identity in (
        "__package_only__",
        "PACKAGE_ONLY",
    )
    if package_only:
        identity = "__package_only__"
        secondary_names = []
        sizzle = ""
        wrap = ""

    if not identity:
        raise ValueError(
            "IDENTITY（主見本）が未選択です。副見本や SIZZLE／WRAP だけでは確定できません。"
            "個品寄りの1枚を IDENTITY 主にするか、正本のみ／個品なしを選んでください。"
        )

    if identity in secondary_names:
        secondary_names = [x for x in secondary_names if x != identity]

    meta: Dict[str, Any] = {
        "builtAt": datetime.now(timezone.utc).isoformat(),
        "truth": str(Path(truth)),
        "confirmedBy": "human",
        "primaryName": identity,
        "identityName": identity,
        "sizzleName": sizzle or None,
        "wrapName": wrap or None,
        "secondaryNames": list(secondary_names),
        "identityStatus": "absent" if package_only else "present",
        "sizzleStatus": (
            "absent"
            if package_only or not sizzle
            else "present"
        ),
        "wrapStatus": (
            "absent"
            if package_only or not wrap
            else "present"
        ),
        "unitPlatingForbidden": bool(package_only),
        "packageOnly": bool(package_only),
        "nameBase": str(name_base or "").strip(),
        "parentSku": str(parent_sku or "").strip(),
    }
    cand_meta = load_candidates_meta(out_dir)
    meta["candidates"] = cand_meta.get("candidates") or [
        {"index": i, "name": p.name, "path": str(p), "picked": False}
        for i, p in enumerate(comps)
    ]

    def _copy_named(src_key: str, stem: str) -> Path:
        if src_key in ("__truth__", "PACKAGE_TRUTH", "__package_only__", ""):
            t = Path(truth)
            if not t.is_file():
                raise ValueError("正本がありません")
            dest = refs / f"{stem}{t.suffix.lower() or '.png'}"
            shutil.copy2(t, dest)
            return dest
        src = by_name.get(src_key)
        if not src or not src.is_file():
            # MANUAL 等が comps に無い場合 refs 直読み
            alt = refs / src_key
            if alt.is_file():
                src = alt
            else:
                raise ValueError(f"見本画像が見つかりません: {src_key}")
        dest = refs / f"{stem}{src.suffix.lower() or '.jpg'}"
        shutil.copy2(src, dest)
        return dest

    dest = _copy_named(identity, CANON_STEM)
    meta["source"] = (
        "package_only_fallback"
        if package_only
        else (
            "package_truth_human"
            if identity in ("__truth__", "PACKAGE_TRUTH")
            else "competitor_human"
        )
    )
    meta["path"] = str(dest.resolve())
    meta["chosenName"] = (
        Path(truth).name
        if identity in ("__truth__", "PACKAGE_TRUTH", "__package_only__")
        else identity
    )
    meta["ok"] = True

    sizzle_path_written: Optional[str] = None
    if sizzle:
        sdest = _copy_named(sizzle, SIZZLE_STEM)
        sizzle_path_written = str(sdest.resolve())
        meta["sizzlePath"] = sizzle_path_written

    wrap_path_written: Optional[str] = None
    if wrap:
        wdest = _copy_named(wrap, WRAP_STEM)
        wrap_path_written = str(wdest.resolve())
        meta["wrapPath"] = wrap_path_written

    for c in meta["candidates"]:
        n = c.get("name")
        roles = []
        if identity not in ("__truth__", "__package_only__", "") and n == identity:
            roles.append("identity")
        if sizzle and n == sizzle:
            roles.append("sizzle")
        if wrap and n == wrap:
            roles.append("wrap")
        if n in secondary_names:
            roles.append("secondary")
        c["picked"] = bool(roles)
        c["role"] = "+".join(roles) if roles else "dropped"

    sec_written: List[str] = []
    for i, name in enumerate(secondary_names, start=1):
        src = by_name.get(name) or (refs / name)
        if not Path(src).is_file():
            LOG.warning("secondary missing %s", name)
            continue
        dest_s = refs / f"{SECONDARY_STEM}_{i}{Path(src).suffix.lower() or '.jpg'}"
        shutil.copy2(src, dest_s)
        sec_written.append(str(dest_s.resolve()))
    meta["secondaryPaths"] = sec_written

    food_canons = {
        **meta,
        "schema": "food_canons_v1",
        "identity": {
            "status": meta["identityStatus"],
            "primaryName": identity,
            "path": meta.get("path"),
            "secondaryNames": list(secondary_names),
            "secondaryPaths": sec_written,
        },
        "sizzle": {
            "status": meta["sizzleStatus"],
            "primaryName": sizzle or None,
            "path": sizzle_path_written,
        },
        "wrap": {
            "status": meta["wrapStatus"],
            "primaryName": wrap or None,
            "path": wrap_path_written,
        },
        "themeModeMap": dict(_THEME_FOOD_MODE),
    }
    (out_dir / FOOD_CANONS_NAME).write_text(
        json.dumps(food_canons, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_dir / META_NAME).write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    if cand_meta:
        cand_meta = {
            **cand_meta,
            "status": "confirmed",
            "primaryName": identity,
            "identityName": identity,
            "sizzleName": sizzle or None,
            "wrapName": wrap or None,
            "secondaryNames": list(secondary_names),
            "confirmedAt": meta["builtAt"],
        }
        (out_dir / CANDIDATES_NAME).write_text(
            json.dumps(cand_meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    LOG.info(
        "human food canons identity=%s sizzle=%s wrap=%s secondary=%s",
        identity,
        sizzle or "(none)",
        wrap or "(none)",
        secondary_names,
    )
    return meta


def ensure_food_unit_canon(
    *,
    truth: Path,
    comps: Sequence[Path],
    out_dir: Path,
    force: bool = False,
) -> Tuple[Optional[Path], Dict[str, Any]]:
    """
    人間確定済みがあればそれを優先。無ければ機械1位で CANON を作る（後方互換）。
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    refs = out_dir / "01_refs"
    refs.mkdir(parents=True, exist_ok=True)

    if not force:
        ex = existing_canon_path(out_dir)
        meta = load_canon_meta(out_dir)
        if ex and meta.get("confirmedBy") == "human":
            meta = {**meta, "reused": True, "path": str(ex)}
            return ex, meta
        if ex and not force:
            if not meta.get("candidates"):
                meta["candidates"] = [
                    {
                        "index": i,
                        "name": p.name,
                        "path": str(p),
                        "picked": p.name == meta.get("chosenName"),
                    }
                    for i, p in enumerate(resolve_comps_for_dir(out_dir))
                ]
            meta = {**meta, "reused": True, "path": str(ex)}
            LOG.info("food canon reuse %s", ex)
            return ex, meta

    chosen, pick_meta = pick_best_comp_vs_truth(Path(truth), comps)
    meta: Dict[str, Any] = {
        "builtAt": datetime.now(timezone.utc).isoformat(),
        "truth": str(Path(truth)),
        "source": "competitor" if chosen else "package_truth",
        "confirmedBy": "machine",
        **pick_meta,
    }
    if chosen and Path(chosen).is_file():
        dest = refs / f"{CANON_STEM}{Path(chosen).suffix.lower() or '.jpg'}"
        if dest.resolve() != Path(chosen).resolve():
            shutil.copy2(chosen, dest)
        meta["path"] = str(dest)
        meta["chosenName"] = Path(chosen).name
        meta["chosenRel"] = f"01_refs/{Path(chosen).name}"
        meta["candidates"] = [
            {
                "index": i,
                "name": Path(p).name,
                "path": str(Path(p)),
                "picked": Path(p).resolve() == Path(chosen).resolve(),
            }
            for i, p in enumerate(comps)
            if Path(p).is_file()
        ]
        (out_dir / META_NAME).write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return dest, meta

    t = Path(truth)
    if t.is_file():
        dest = refs / f"{CANON_STEM}{t.suffix.lower() or '.png'}"
        if dest.resolve() != t.resolve():
            shutil.copy2(t, dest)
        meta["path"] = str(dest)
        meta["chosenName"] = t.name
        meta["candidates"] = [
            {"index": i, "name": Path(p).name, "path": str(Path(p)), "picked": False}
            for i, p in enumerate(comps)
            if Path(p).is_file()
        ]
        meta["source"] = "package_truth_copy"
        (out_dir / META_NAME).write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return dest, meta
    return None, meta


def ensure_for_theme_new_dir(theme_new_dir: Path, *, force: bool = False) -> Dict[str, Any]:
    root = Path(theme_new_dir)
    summary_path = root / "summary.json"
    truth: Optional[Path] = None
    if summary_path.is_file():
        try:
            s = json.loads(summary_path.read_text(encoding="utf-8"))
            t = Path(str(s.get("truth") or ""))
            if t.is_file():
                truth = t
        except Exception:
            pass
    if not truth:
        pkg = list((root / "01_refs").glob("PACKAGE_TRUTH.*")) if (root / "01_refs").is_dir() else []
        truth = pkg[0] if pkg else None
    comps = resolve_comps_for_dir(root)
    if not truth or not truth.is_file():
        return {"ok": False, "reason": "no_truth", "path": "", "candidates": []}
    path, meta = ensure_food_unit_canon(
        truth=truth, comps=comps, out_dir=root, force=force
    )
    meta = {**meta, "path": str(path or meta.get("path") or "")}
    return meta


def food_unit_prompt_block(*, has_canon_image: bool, has_secondary: bool = False) -> str:
    extra = ""
    if has_canon_image:
        extra += (
            "\n- IMAGE_FOOD_UNIT_CANON（主見本）がある。色・厚み・1片サイズはこれに合わせよ。\n"
        )
    if has_secondary:
        extra += (
            "- IMAGE_FOOD_UNIT_SECONDARY_* は盛り付け（置き方・重なり）の参考だけ。1片の形・大きさ・厚み・色には使うな。"
            "色相は主見本を優先（副の照明差に引っ張られるな）。\n"
        )
    return FOOD_UNIT_LOCK_JA.strip() + extra
