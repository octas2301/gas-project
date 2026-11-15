# -*- coding: utf-8 -*-
"""一片セット・実物由来ゲート（§4.5.10bis-T F2）。

案件名・JAN・魚種をロジックに焼かない。入力は theme_new_run 配下のファイルと
piece_set_lang の機械言語化だけ。

正:
- 種の優先: 人間実物(PIECE_TRUTH_HUMAN) → WRAP →（最終手段）仮想
- 外形ファミリーが cube のとき、扇・くさび・三角などのシルエット暴走禁止
  （向き・角度の差は可。外形テンプレで別物にしない）
- 添付 max_n と言語化番号の一致は F1（ここでは補助チェック）
- IDENTITY／CANON が一片セット（または HUMAN 一片）と同一バイトなら不合格
  （盛り見本を別途用意。案件名・魚種は焼かない）

切戻し: 環境変数 B5_PIECE_TRUTH_GATE=0|off（検査しない）
厳格: B5_PIECE_TRUTH_GATE=strict（compose で不合格なら中断）
既定: warn（meta に書き、ログ警告のみ）
"""
from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

SOURCE_NAME = "piece_set_source.json"
LANG_REL = Path("food_piece_set5_virt") / "piece_set_lang.json"
SET_DIR_REL = Path("food_piece_set5_virt") / "piece_set"
HUMAN_DIR_REL = Path("01_refs") / "PIECE_TRUTH_HUMAN"
FOOD_CANONS_NAME = "food_canons.json"
CANON_STEM = "FOOD_UNIT_CANON"

# 言語化「形状」からファミリーを決める語（案件固定なし）
_CUBE_MARKERS = (
    "キューブ",
    "市松",
    "直方体",
    "ダイス",
    "ブロック",
    "角型",
    "正方形",
    "立方",
)
_FLAKE_MARKERS = (
    "削り",
    "薄片",
    "フレーク",
    "帯状",
    "リボン",
    "笹",
    "片状",
    "薄めの",
    "細長",
    "ねじれ",
    "反り",
    "丸まり",
    "カール",
    "舟形",
    "削り片",
)

# cube で禁止するシルエット語（向きの差は別。外形バラつきテンプレ禁止）
_CUBE_FORBIDDEN_SHAPE = (
    "扇形",
    "扇状",
    "くさび",
    "楔",
    "三角形",
    "三角に",
    "三角の",
    "三角ブ",
    "ウェッジ",
    "wedge",
    "fan shape",
)

_FAMILY_CUBE = "cube"
_FAMILY_FLAKE = "flake"
_FAMILY_OTHER = "other"


def piece_truth_gate_mode() -> str:
    """未設定=warn / 0|off=off / strict|1|on=strict。"""
    if "B5_PIECE_TRUTH_GATE" not in os.environ:
        return "warn"
    raw = str(os.environ.get("B5_PIECE_TRUTH_GATE") or "").strip().lower()
    if raw in ("", "warn", "warning", "log"):
        return "warn"
    if raw in ("0", "false", "no", "off"):
        return "off"
    if raw in ("strict", "error", "fail", "1", "on", "true", "yes"):
        return "strict"
    return "warn"


def discover_human_truth_paths(theme_dir: Path) -> List[Path]:
    d = Path(theme_dir) / HUMAN_DIR_REL
    if not d.is_dir():
        return []
    files: List[Path] = []
    for pat in ("piece_truth_0*.png", "piece_truth_0*.jpg", "piece_truth_0*.jpeg", "piece_truth_0*.webp"):
        files.extend(d.glob(pat))
    return sorted({p.resolve() for p in files if p.is_file()})


def discover_wrap_paths(theme_dir: Path) -> List[Path]:
    refs = Path(theme_dir) / "01_refs"
    if not refs.is_dir():
        return []
    out: List[Path] = []
    for stem in ("FOOD_UNIT_WRAP", "FOOD_UNIT_WRAP_SEED"):
        for ext in (".jpg", ".jpeg", ".png", ".webp"):
            p = refs / f"{stem}{ext}"
            if p.is_file():
                out.append(p.resolve())
                break
    # 追加の WRAP_01 など
    out.extend(sorted(p.resolve() for p in refs.glob("FOOD_UNIT_WRAP_*.*") if p.is_file()))
    # 重複除去・順序維持
    seen = set()
    uniq: List[Path] = []
    for p in out:
        s = str(p)
        if s in seen:
            continue
        seen.add(s)
        uniq.append(p)
    return uniq


def discover_piece_set_paths(theme_dir: Path, *, max_n: int = 0) -> List[Path]:
    set_dir = Path(theme_dir) / SET_DIR_REL
    if not set_dir.is_dir():
        return []
    files = sorted(set_dir.glob("piece_set_0*.jpg"))
    n = int(max_n or 0)
    if n > 0:
        files = files[:n]
    return files


def load_piece_set_lang(theme_dir: Path) -> Dict[str, Any]:
    p = Path(theme_dir) / LANG_REL
    if not p.is_file():
        return {}
    try:
        obj = json.loads(p.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}
    return obj if isinstance(obj, dict) else {}


def load_piece_set_source(theme_dir: Path) -> Dict[str, Any]:
    p = Path(theme_dir) / "food_piece_set5_virt" / SOURCE_NAME
    if not p.is_file():
        return {}
    try:
        obj = json.loads(p.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}
    return obj if isinstance(obj, dict) else {}


def write_piece_set_source(
    theme_dir: Path,
    *,
    kind: str,
    paths: Sequence[Path],
    extra: Optional[Dict[str, Any]] = None,
) -> Path:
    """生成側が種の出自を残す（人間／WRAP／仮想）。

    既存の activeSeed（ゲートで選んだ元見本）は、extra で明示しない限り消さない。
    """
    out_dir = Path(theme_dir) / "food_piece_set5_virt"
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / SOURCE_NAME
    prev_active = ""
    if dest.is_file():
        try:
            prev = json.loads(dest.read_text(encoding="utf-8-sig"))
            if isinstance(prev, dict):
                prev_active = str(prev.get("activeSeed") or "").strip()
        except Exception:
            prev_active = ""
    payload: Dict[str, Any] = {
        "ok": True,
        "kind": str(kind or "").strip().lower() or "virt",
        "paths": [str(Path(p)) for p in paths],
        "rule": "human > wrap > virt",
    }
    if extra:
        payload.update(extra)
    if not str(payload.get("activeSeed") or "").strip():
        if prev_active:
            payload["activeSeed"] = prev_active
        elif len(list(paths)) == 1:
            payload["activeSeed"] = Path(list(paths)[0]).name
    dest.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return dest


def discover_piece_seed_human_paths(theme_dir: Path) -> List[Path]:
    """生成用の元見本（種）。PIECE_SEED_HUMAN_0* を優先。"""
    refs = Path(theme_dir) / "01_refs"
    if not refs.is_dir():
        return []
    numbered = sorted(p for p in refs.glob("PIECE_SEED_HUMAN_0*.*") if p.is_file())
    if numbered:
        return [p.resolve() for p in numbered]
    out: List[Path] = []
    for pat in ("PIECE_SEED_HUMAN.*", "piece_seed_0*.*"):
        out.extend(p for p in refs.glob(pat) if p.is_file())
    return [p.resolve() for p in sorted(out, key=lambda x: x.name)]


def resolve_preferred_seed_kind(theme_dir: Path) -> Tuple[str, List[Path]]:
    """利用可能な最優先種。戻り値 kind は human|wrap|virt|none。

    優先: PIECE_SEED_HUMAN（生成の種）→ PIECE_TRUTH_HUMAN → WRAP → CANON仮想。
    piece_set_source.activeSeed があればその1枚を採用。
    """
    seeds = discover_piece_seed_human_paths(theme_dir)
    if seeds:
        active_name = ""
        try:
            src_json = Path(theme_dir) / "food_piece_set5_virt" / SOURCE_NAME
            if src_json.is_file():
                obj = json.loads(src_json.read_text(encoding="utf-8"))
                if isinstance(obj, dict):
                    active_name = str(obj.get("activeSeed") or "").strip()
        except Exception:
            active_name = ""
        if active_name:
            for p in seeds:
                if p.name == active_name:
                    return "human", [p]
        return "human", [seeds[0]]
    human = discover_human_truth_paths(theme_dir)
    if human:
        return "human", human
    wrap = discover_wrap_paths(theme_dir)
    if wrap:
        return "wrap", wrap
    # 仮想フォールバック種: CANON
    refs = Path(theme_dir) / "01_refs"
    for name in ("FOOD_UNIT_CANON.jpg", "FOOD_UNIT_CANON.png", "MANUAL_02.jpg", "MANUAL_01.jpg"):
        p = refs / name
        if p.is_file():
            return "virt", [p.resolve()]
    return "none", []


def infer_shape_family_from_texts(texts: Sequence[str]) -> str:
    """形状文の多数決で cube / flake / other。"""
    cube_n = 0
    flake_n = 0
    for raw in texts or []:
        t = str(raw or "")
        if any(m in t for m in _CUBE_MARKERS):
            cube_n += 1
        if any(m in t for m in _FLAKE_MARKERS):
            flake_n += 1
    if cube_n == 0 and flake_n == 0:
        return _FAMILY_OTHER
    if cube_n >= flake_n and cube_n > 0:
        return _FAMILY_CUBE
    if flake_n > 0:
        return _FAMILY_FLAKE
    return _FAMILY_OTHER


def infer_shape_family_from_lang(lang: Dict[str, Any]) -> str:
    texts: List[str] = []
    for p in lang.get("pieces") or []:
        if not isinstance(p, dict):
            continue
        fields = p.get("fields") if isinstance(p.get("fields"), dict) else {}
        texts.append(str(fields.get("形状") or ""))
        texts.append(str(p.get("promptBlock") or ""))
    texts.append(str(lang.get("promptBlock") or ""))
    texts.append(str(lang.get("rule") or ""))
    return infer_shape_family_from_texts(texts)


_SHAPE_TOK = r"(?:扇(?:形|状)?|くさび|楔|三角(?:形)?|ウェッジ|wedge)"
_NEGATION_SHAPE_RE = re.compile(
    r"(?:完全な幾何四角[・･]?)?"
    + _SHAPE_TOK
    + r"(?:[・･／/]"
    + _SHAPE_TOK
    + r"|[^。．\n禁止]{0,12})*"
    + r"禁止",
    re.IGNORECASE,
)


def _strip_shape_negation_clauses(text: str) -> str:
    """『扇・三角禁止』『三角シルエット禁止』等の否定を除き、実使用の形語だけ残す。"""
    t = str(text or "")
    t = _NEGATION_SHAPE_RE.sub(" ", t)
    t = re.sub(
        _SHAPE_TOK + r"[^。．\n]{0,16}禁止",
        " ",
        t,
        flags=re.IGNORECASE,
    )
    t = re.sub(
        r"禁止[^。．\n]{0,16}" + _SHAPE_TOK,
        " ",
        t,
        flags=re.IGNORECASE,
    )
    return t


def cube_forbidden_hits(text: str) -> List[str]:
    t = _strip_shape_negation_clauses(text)
    hits = [tok for tok in _CUBE_FORBIDDEN_SHAPE if tok.lower() in t.lower()]
    # 短い「三角」はキューブ文脈で、否定除去後に残っていたときだけ
    if "三角" in t and any(m in t for m in _CUBE_MARKERS):
        if "三角" not in hits and "三角形" not in hits:
            hits.append("三角")
    return hits


def shape_variety_policy(family: str) -> Dict[str, Any]:
    """生成プロンプト用の汎用方針（SKU非依存）。"""
    fam = str(family or _FAMILY_OTHER)
    if fam == _FAMILY_CUBE:
        return {
            "family": fam,
            "allowSilhouetteScatter": False,
            "forbidTokens": list(_CUBE_FORBIDDEN_SHAPE),
            "promptJa": (
                "【外形ファミリー・キューブ】各片は同じ直方体／キューブ系統を保て。"
                "変えてよいのは向き・角度・わずかな歪みだけ。"
                "扇・くさび・三角シルエット・別外形テンプレは禁止。"
            ),
        }
    if fam == _FAMILY_FLAKE:
        return {
            "family": fam,
            "allowSilhouetteScatter": True,
            "forbidTokens": [],
            "promptJa": (
                "【外形ファミリー・削り／薄片】片ごとの外形差は可。"
                "ただし色・縁・筋・厚み系統はセットで揃えろ。"
                "血合いの有無は片の言語化に従い、全片共通の濃縁帯にするな。"
            ),
        }
    return {
        "family": fam,
        "allowSilhouetteScatter": True,
        "forbidTokens": list(_CUBE_FORBIDDEN_SHAPE),
        "promptJa": (
            "【外形】種写真の系統を保て。扇・くさびへ勝手に寄せるな。"
            "向きの差は可。別商品の形に描き直すな。"
        ),
    }


def build_shape_policy_prompt_block(family: str) -> str:
    return str(shape_variety_policy(family).get("promptJa") or "").strip() + "\n"


def theme_shape_family(theme_dir: Path) -> str:
    """theme_new_run の piece_set_lang から外形ファミリーを返す。"""
    lang = load_piece_set_lang(theme_dir)
    if lang:
        return infer_shape_family_from_lang(lang)
    return _FAMILY_OTHER


def _theme_jan_digits(theme_dir: Path) -> str:
    name = Path(theme_dir).name
    m = re.search(r"(\d{8,})", name)
    return m.group(1) if m else ""


def list_theme_new_run_dirs(*, script_dir: Optional[Path] = None) -> List[Path]:
    root = Path(script_dir) if script_dir is not None else Path(__file__).resolve().parent
    out = root / "_out"
    if not out.is_dir():
        return []
    return sorted(p for p in out.glob("theme_new_run_*") if p.is_dir())


def shape_donor_attach_prompt_block(*, max_n: int) -> str:
    """IMAGE_FOOD_UNIT_SHAPE_* 同送時の役割文（SKU非焼付）。"""
    n = int(max_n or 0)
    if n <= 0:
        return ""
    return (
        "【SHAPE添付・形状のみ・必須】\n"
        f"添付の IMAGE_FOOD_UNIT_SHAPE_1〜{n} は同ファミリー別SKUの一片。"
        "細長さ・反り・丸まり・ねじれ・薄さの形だけ借りる。\n"
        "SHAPEの色・筋・血合い色・縁の質感を自SKUの正にするな。"
        "表面・縁の正は自SKUの SURFACE（あれば）。"
        "一片本体を添付していないときはSHAPEの外形だけ借り、縁毛羽を面の代用にするな。\n"
    )


def surface_attach_prompt_block(*, max_n: int) -> str:
    """IMAGE_FOOD_UNIT_SURFACE_* 同送時の役割文（SKU非焼付）。"""
    n = int(max_n or 0)
    if n <= 0:
        return ""
    return (
        "【SURFACE添付・表面のみ・必須】\n"
        f"添付の IMAGE_FOOD_UNIT_SURFACE_1〜{n} は面のアップ／面クロップ。"
        "筋・色・血合い色・表面感だけ借りる。\n"
        "外形・反り・丸まり・ねじれ・細長さはSURFACEから取るな。"
        "縁の毛羽で面の筋の代用をするな。\n"
        "面とほぼ同じ高さの細い木目状繊維を各片で読ませる。"
        "溝・段差・彫り込み・のっぺり無地禁止。\n"
        "食品面はしっとり・鈍い戻り（言語化の「乾き」より優先）。"
        "鏡面ハイライト・片目テカリ禁止（缶・瓶・皿のスペキュラと混ぜるな）。\n"
    )


def face_crop_for_surface(src: Path, dest: Path) -> Path:
    """外形の影響を弱めるため中央寄りを切り出す（表面スロット用・SKU非依存）。"""
    from PIL import Image

    src = Path(src)
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    im = Image.open(src).convert("RGB")
    w, h = im.size
    box = (int(w * 0.18), int(h * 0.18), int(w * 0.82), int(h * 0.82))
    crop = im.crop(box)
    side = max(crop.size)
    if side < 768:
        s = 768 / float(side)
        crop = crop.resize(
            (max(1, int(crop.size[0] * s)), max(1, int(crop.size[1] * s))),
            Image.Resampling.LANCZOS,
        )
    crop.save(dest, quality=95)
    return dest


def resolve_surface_ref_paths(
    theme_dir: Path,
    *,
    max_n: int = 3,
    allow_piece_face_crop: Optional[bool] = None,
) -> Dict[str, Any]:
    """表面スロット用画像。専用面アップを採用。足りなければ枚数を減らす（既定）。

    優先: deepen_refs/ref_surface* → 01_refs の面系（SEED／CLOSE／SURFACE）→
    （任意・B5_SURFACE_FACE_CROP=1）piece_set の面クロップ。案件名・魚種は焼かない。
    """
    if allow_piece_face_crop is None:
        try:
            from b5_piece_deepen import surface_face_crop_fill_enabled

            allow_piece_face_crop = bool(surface_face_crop_fill_enabled())
        except Exception:
            allow_piece_face_crop = False
    theme_dir = Path(theme_dir).resolve()
    n = int(max_n or 0)
    empty: Dict[str, Any] = {
        "ok": False,
        "paths": [],
        "sources": [],
        "reason": "",
        "usedFaceCrop": False,
    }
    if n <= 0:
        empty["reason"] = "max_n<=0"
        return empty

    # (priority昇順, -bytes, name) で並べ、大きい面アップを先に
    cands: List[Tuple[int, int, str, Path, str]] = []

    def _add(pri: int, p: Path, src: str) -> None:
        if not p.is_file():
            return
        if p.suffix.lower() not in (".jpg", ".jpeg", ".png", ".webp"):
            return
        name_u = p.name.upper()
        if "CHIAI" in name_u or "血合" in p.name:
            return
        # スマホマクロ／試験産物（reinhard等）は本線同送外
        try:
            from b5_generic_plating import is_excluded_surface_attach_name

            if is_excluded_surface_attach_name(p.name):
                return
        except Exception:
            name_l = p.name.lower()
            if (
                "phone" in name_l
                or "ref_surface_cam" in name_l
                or "reinhard" in name_l
            ):
                return
        try:
            sz = int(p.stat().st_size)
        except Exception:
            sz = 0
        # 小さすぎるサムネは面の正にしない
        if sz < 80_000 and "SURFACE" not in name_u and "CLOSE" not in name_u:
            return
        cands.append((pri, -sz, p.name.lower(), p.resolve(), src))

    # 1) 深堀参考の表面
    deep = theme_dir / "food_piece_set5_virt" / "deepen_refs"
    if deep.is_dir():
        for p in deep.glob("ref_surface*"):
            _add(0, p, "deepen_ref_surface")
    # 2) activeSeed（ゲートで選んだ面種）
    refs = theme_dir / "01_refs"
    active = ""
    try:
        src_json = theme_dir / "food_piece_set5_virt" / SOURCE_NAME
        if src_json.is_file():
            sobj = json.loads(src_json.read_text(encoding="utf-8-sig"))
            if isinstance(sobj, dict):
                active = str(sobj.get("activeSeed") or "").strip()
    except Exception:
        active = ""
    if active and refs.is_dir():
        _add(1, refs / active, "active_seed")
    # 3) 01_refs の面アップのみ（他SKU試験の種一覧で枠を埋めない）
    #    種フォールバックは activeSeed の1枚だけ（ゲート選定）。
    if refs.is_dir():
        patterns = (
            "*SURFACE*CLOSE*.*",
            "*surface*close*.*",
            "PIECE_SURFACE*.*",
            "ref_surface*.*",
        )
        for pat in patterns:
            for p in refs.glob(pat):
                name_u = p.name.upper()
                pri = 2 if ("SURFACE" in name_u or "CLOSE" in name_u) else 3
                _add(pri, p, "refs_surface_or_seed")
        # activeSeed 以外の PIECE_SEED_HUMAN_* は SURFACE に使わない
        if active:
            _add(3, refs / active, "active_seed_surface_fallback")

    paths: List[Path] = []
    sources: List[str] = []
    seen: set[str] = set()

    def _take(p: Path, src: str) -> bool:
        if len(paths) >= n:
            return False
        key = str(p.resolve())
        if key in seen:
            return False
        try:
            digest = _file_sha256(p)
        except Exception:
            digest = key
        if digest in seen:
            return False
        seen.add(key)
        seen.add(digest)
        paths.append(p.resolve())
        sources.append(src)
        return True

    # 専用面（pri0-2）を先に
    cands.sort(key=lambda x: (x[0], x[1], x[2]))
    for pri, _neg_sz, _name, p, src in cands:
        if pri >= 3:
            continue
        _take(p, src)

    # 足りなければ一片の面クロップ（小さいSEEDより面情報が取りやすい）
    used_crop = False
    if len(paths) < n and allow_piece_face_crop:
        pieces = discover_piece_set_paths(theme_dir, max_n=n)
        crop_dir = theme_dir / "food_piece_set5_virt" / "surface_face_crops"
        for i, src_p in enumerate(pieces, 1):
            if len(paths) >= n:
                break
            dest = crop_dir / f"surface_face_{i:02d}.jpg"
            try:
                face_crop_for_surface(src_p, dest)
            except Exception:
                continue
            if _take(dest, "piece_set_face_crop"):
                used_crop = True

    # それでも足りなければ種SEED
    for pri, _neg_sz, _name, p, src in cands:
        if pri < 3:
            continue
        _take(p, src)

    if not paths:
        empty["reason"] = (
            "表面見本が無い"
            if not allow_piece_face_crop
            else "表面見本も一片面クロップも無い"
        )
        return empty
    return {
        "ok": True,
        "paths": paths[:n],
        "sources": sources[:n],
        "n": len(paths[:n]),
        "usedFaceCrop": used_crop,
        "reason": "",
    }


def resolve_shape_donor_paths(
    theme_dir: Path,
    *,
    max_n: int = 3,
    script_dir: Optional[Path] = None,
    family: str = "",
) -> Dict[str, Any]:
    """同ファミリー別 theme の piece_set を形状ドナーとして返す。

    案件名・魚種を焼かない。family は言語化から機械判定（未指定なら自 theme）。
    hasPieceSet=false の theme は除外。切戻しは呼び出し側の B5_SURFACE_SHAPE_SPLIT。
    """
    theme_dir = Path(theme_dir).resolve()
    n = int(max_n or 0)
    empty: Dict[str, Any] = {
        "ok": False,
        "family": "",
        "paths": [],
        "donorJans": [],
        "donorDirs": [],
        "reason": "",
    }
    if n <= 0:
        empty["reason"] = "max_n<=0"
        return empty
    fam = str(family or "").strip().lower() or theme_shape_family(theme_dir)
    empty["family"] = fam
    self_jan = _theme_jan_digits(theme_dir)
    candidates: List[Tuple[int, Path, List[Path]]] = []
    for other in list_theme_new_run_dirs(script_dir=script_dir):
        other_r = other.resolve()
        if other_r == theme_dir:
            continue
        other_jan = _theme_jan_digits(other_r)
        if self_jan and other_jan and other_jan == self_jan:
            continue
        # ゲートで一片なし確定なら除外
        canons = other_r / FOOD_CANONS_NAME
        if canons.is_file():
            try:
                meta = json.loads(canons.read_text(encoding="utf-8-sig"))
                if isinstance(meta, dict) and meta.get("hasPieceSet") is False:
                    continue
            except Exception:
                pass
        ofam = theme_shape_family(other_r)
        if ofam != fam:
            continue
        files = discover_piece_set_paths(other_r, max_n=0)
        if not files:
            continue
        candidates.append((len(files), other_r, files))
    if not candidates:
        empty["reason"] = f"同ファミリー({fam})の別SKU一片が無い"
        return empty
    # 枚数が多い theme を先に（形バリエが取りやすい）
    candidates.sort(key=lambda x: (-x[0], x[1].name))
    paths: List[Path] = []
    donor_jans: List[str] = []
    donor_dirs: List[str] = []
    seen: set[str] = set()
    for _cnt, odir, files in candidates:
        if len(paths) >= n:
            break
        took = False
        for p in files:
            if len(paths) >= n:
                break
            key = str(p.resolve())
            if key in seen:
                continue
            seen.add(key)
            paths.append(p.resolve())
            took = True
        if took:
            donor_dirs.append(str(odir))
            dj = _theme_jan_digits(odir)
            if dj and dj not in donor_jans:
                donor_jans.append(dj)
    if not paths:
        empty["reason"] = "候補はあるがファイルを取れない"
        return empty
    return {
        "ok": True,
        "family": fam,
        "paths": paths,
        "donorJans": donor_jans,
        "donorDirs": donor_dirs,
        "n": len(paths),
        "reason": "",
    }


def _file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _resolve_named_ref(theme_dir: Path, name: str) -> Optional[Path]:
    name = str(name or "").strip()
    if not name or "/" in name or "\\" in name or ".." in name:
        return None
    theme_dir = Path(theme_dir)
    for folder in (theme_dir / "01_refs", theme_dir):
        p = folder / name
        if p.is_file():
            return p.resolve()
    return None


def resolve_identity_candidate_paths(
    theme_dir: Path,
    *,
    identity_name: str = "",
) -> List[Path]:
    """IDENTITY／CANON 候補（案件名非依存）。明示名 → food_canons → FOOD_UNIT_CANON。"""
    theme_dir = Path(theme_dir)
    out: List[Path] = []
    seen = set()

    def _add(p: Optional[Path]) -> None:
        if p is None or not p.is_file():
            return
        key = str(p.resolve())
        if key in seen:
            return
        seen.add(key)
        out.append(p.resolve())

    _add(_resolve_named_ref(theme_dir, identity_name))
    for meta_name in (FOOD_CANONS_NAME, Path("01_refs") / FOOD_CANONS_NAME):
        meta_path = theme_dir / meta_name
        if not meta_path.is_file():
            continue
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8-sig"))
        except Exception:
            continue
        if not isinstance(meta, dict):
            continue
        for key in ("identityName", "primaryName", "primary"):
            _add(_resolve_named_ref(theme_dir, str(meta.get(key) or "")))
    for folder in (theme_dir / "01_refs", theme_dir):
        if not folder.is_dir():
            continue
        for ext in (".jpg", ".jpeg", ".png", ".webp"):
            _add(folder / f"{CANON_STEM}{ext}")
    return out


def find_identity_same_as_piece(
    theme_dir: Path,
    *,
    identity_name: str = "",
    max_n: int = 0,
) -> Dict[str, Any]:
    """IDENTITY／CANON が一片と同一バイトなら collisions を返す（不合格材料）。"""
    theme_dir = Path(theme_dir)
    idents = resolve_identity_candidate_paths(theme_dir, identity_name=identity_name)
    pieces = list(discover_piece_set_paths(theme_dir, max_n=max_n))
    human = theme_dir / HUMAN_DIR_REL
    if human.is_dir():
        pieces.extend(
            sorted(
                p
                for p in human.glob("piece_truth_0*.*")
                if p.is_file() and p.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp"}
            )
        )
    piece_digests: Dict[str, Path] = {}
    for p in pieces:
        try:
            piece_digests[_file_sha256(p)] = p
        except Exception:
            continue
    collisions: List[Dict[str, str]] = []
    for ident in idents:
        try:
            dig = _file_sha256(ident)
        except Exception:
            continue
        hit = piece_digests.get(dig)
        if hit is None:
            continue
        collisions.append(
            {
                "identity": ident.name,
                "piece": hit.name,
                "sha256": dig[:16],
            }
        )
    return {
        "ok": not collisions,
        "identityN": len(idents),
        "pieceN": len(piece_digests),
        "collisions": collisions,
    }


def identity_piece_collision_errors(
    theme_dir: Path,
    *,
    identity_name: str = "",
    max_n: int = 0,
) -> List[str]:
    hit = find_identity_same_as_piece(
        theme_dir, identity_name=identity_name, max_n=max_n
    )
    errors: List[str] = []
    for row in hit.get("collisions") or []:
        if not isinstance(row, dict):
            continue
        errors.append(
            "IDENTITY／CANON が一片と同一です"
            f"（{row.get('identity')}＝{row.get('piece')}）。"
            "盛り配置・量感の見本を別途用意するか、IDENTITYを盛り写真に差し替えよ。"
            "一片は個品の正、IDENTITYは盛り見本。"
        )
    return errors


def evaluate_piece_truth_gate(
    theme_dir: Path,
    *,
    max_n: int = 0,
    source: Optional[Dict[str, Any]] = None,
    lang: Optional[Dict[str, Any]] = None,
    identity_name: str = "",
) -> Dict[str, Any]:
    """一片セットの実物由来＋外形暴走ゲート。

    ok=False のとき errors に理由。warnings は情報。
    """
    theme_dir = Path(theme_dir)
    errors: List[str] = []
    warnings: List[str] = []
    preferred_kind, preferred_paths = resolve_preferred_seed_kind(theme_dir)
    src = dict(source) if isinstance(source, dict) else load_piece_set_source(theme_dir)
    lang_obj = dict(lang) if isinstance(lang, dict) else load_piece_set_lang(theme_dir)
    set_files = discover_piece_set_paths(theme_dir, max_n=max_n)
    family = infer_shape_family_from_lang(lang_obj) if lang_obj else _FAMILY_OTHER
    policy = shape_variety_policy(family)

    recorded = str(src.get("kind") or "").strip().lower()
    if preferred_kind == "human":
        if recorded and recorded != "human":
            errors.append(
                f"人間実物があるのに piece_set_source.kind={recorded}。"
                "種は human を使え。"
            )
        if not recorded:
            # 旧セット: 人間正本があるのに出自未記録
            errors.append(
                "人間実物(PIECE_TRUTH_HUMAN)があるが piece_set_source.json が無い。"
                "実物から作り直し、kind=human を書け。"
            )
    elif preferred_kind == "wrap":
        if recorded in ("virt", "virtual", "canon"):
            warnings.append(
                "WRAPがあるのに種が仮想扱い。可能なら WRAP 中身を種にせよ。"
            )
        if not recorded:
            warnings.append(
                "WRAPがあるが piece_set_source.json が無い。出自を wrap で残せよ。"
            )
    elif preferred_kind == "virt":
        if not recorded:
            warnings.append("種は仮想フォールバック。piece_set_source に kind=virt を残せ。")

    if not set_files:
        errors.append("piece_set 画像が無い。")

    # 言語化と枚数
    pieces = lang_obj.get("pieces") if isinstance(lang_obj.get("pieces"), list) else []
    n = int(max_n or 0)
    if n > 0 and pieces and len(pieces) < n and len(set_files) >= n:
        warnings.append(
            f"言語化 pieces={len(pieces)} < max_n={n}。足りない番号は言語化せよ。"
        )
    if n > 0 and pieces and len(pieces) > n:
        # F1 で切り詰めるが、ファイル上は超過しうる
        warnings.append(
            f"言語化 pieces={len(pieces)} > max_n={n}。同送は先頭{n}だけ（F1）。"
        )

    if family == _FAMILY_CUBE and pieces:
        limit = n if n > 0 else len(pieces)
        for i, row in enumerate(pieces[:limit]):
            if not isinstance(row, dict):
                continue
            fields = row.get("fields") if isinstance(row.get("fields"), dict) else {}
            blob = " ".join(
                [
                    str(fields.get("形状") or ""),
                    str(row.get("promptBlock") or ""),
                    str(row.get("name") or ""),
                ]
            )
            hits = cube_forbidden_hits(blob)
            if hits:
                name = str(row.get("name") or f"piece[{i+1}]")
                errors.append(
                    f"cube外形暴走: {name} に禁止語 {hits}。"
                    "同型キューブ＋向き差のみに直せよ。"
                )

    # IDENTITY＝一片（同一バイト）は盛り見本欠落 → 不合格
    id_coll = find_identity_same_as_piece(
        theme_dir, identity_name=identity_name, max_n=max_n
    )
    for row in id_coll.get("collisions") or []:
        if not isinstance(row, dict):
            continue
        errors.append(
            "IDENTITY／CANON が一片と同一です"
            f"（{row.get('identity')}＝{row.get('piece')}）。"
            "盛り配置・量感の見本を別途用意するか、IDENTITYを盛り写真に差し替えよ。"
            "一片は個品の正、IDENTITYは盛り見本。"
        )

    ok = not errors
    return {
        "ok": ok,
        "themeDir": str(theme_dir),
        "preferredKind": preferred_kind,
        "preferredN": len(preferred_paths),
        "recordedKind": recorded or "",
        "family": family,
        "allowSilhouetteScatter": bool(policy.get("allowSilhouetteScatter")),
        "pieceSetN": len(set_files),
        "langPieceN": len(pieces),
        "maxN": n,
        "identityPieceCollisions": list(id_coll.get("collisions") or []),
        "errors": errors,
        "warnings": warnings,
        "policyPrompt": str(policy.get("promptJa") or ""),
    }


def evaluate_piece_truth_gate_for_jan(
    jan: str,
    *,
    script_dir: Optional[Path] = None,
    max_n: int = 0,
) -> Dict[str, Any]:
    root = Path(script_dir) if script_dir is not None else Path(__file__).resolve().parent
    jan_d = "".join(ch for ch in str(jan) if ch.isdigit()) or str(jan).strip()
    theme = root / "_out" / f"theme_new_run_{jan_d}"
    return evaluate_piece_truth_gate(theme, max_n=max_n)
