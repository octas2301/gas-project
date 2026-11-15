# -*- coding: utf-8 -*-
"""一片／SUB 深堀パネル（§4.5.10bis-T4）。

不合格タグ surface／shape／plate／other に応じ、作業付近に大きめ手順を出す。
案件名・魚種を焼かない。切戻し: B5_PIECE_DEEPEN_UI=0
"""
from __future__ import annotations

import html as html_lib
import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

DEEPEN_REL = Path("food_piece_set5_virt") / "piece_deepen.json"
REFS_REL = Path("food_piece_set5_virt") / "deepen_refs"
PREV_REL = Path("food_piece_set5_virt") / "deepen_prev_pieces"
AFTER_REL = Path("food_piece_set5_virt") / "deepen_after_pieces"
SET_REL = Path("food_piece_set5_virt") / "piece_set"
VALID_TAGS = ("surface", "shape", "plate", "other")

# 段ごとの手順（SKU非依存）
_STEPS: Dict[str, Dict[str, Any]] = {
    "surface": {
        "stage": "S2",
        "title": "深堀・表面（筋・色）",
        "failOne": "表面の筋・色・血合い色が足りない／弱い",
        "anchor": "seed",  # ②元見本の近く
        "steps": [
            "① 下の「② 生成の元見本」に、面の筋がはっきり写る写真を貼る（または既存種をラジオで選ぶ）",
            "② 必要なら「この1枚だけ作り直す」で面だけ直す（形のねじれは触らない）",
            "③ 面の合否だけ見る。合格なら下の「深堀を閉じる」→ 機械が浅い生成へ戻る",
        ],
        "touch": "触ってよい: 元見本・表面の作り直し・表面の要望1行",
        "dont": "触るな: 形状ドナー差し替え、盛りIDENTITYの配置、袋",
        "resume": "閉じたあと: S1（浅い一片／SUB）を再実行して再検収",
    },
    "shape": {
        "stage": "S3",
        "title": "深堀・形状（反り・ねじれ・細長さ）",
        "failOne": "形・反り・丸まり・ねじれ・細長さが足りない／板状のまま",
        "anchor": "set",  # ③一片セットの近く
        "steps": [
            "① 参考段に同ファミリーの形状ドナーが自動で載る（無ければ他themeの一片セットを用意）",
            "② 下の「③ 一片セット」で、形だけ直したいカードを選ぶ",
            "③ 「この1枚への要望」に形状だけ書く（例: もっと細長く・ねじれを足す）。色・筋の話は書かない",
            "④ 「この1枚だけ作り直す」。形状ドナー＋要望で形を寄せ、表面の正は変えない",
            "⑤ 形の合否だけ見る。合格なら「深堀を閉じる」",
        ],
        "touch": "触ってよい: 一片カードの形状要望・1枚作り直し・形状言語化",
        "dont": "触るな: 元見本の差し替えで色を変えること、盛りIDENTITY、ドナーの色を正にすること",
        "resume": "閉じたあと: S1 を再実行して再検収",
    },
    "plate": {
        "stage": "S4",
        "title": "深堀・盛り（IDENTITY）",
        "failOne": "盛り・量・ばらけ・重なりが足りない／IDENTITY欠落",
        "anchor": "plate",  # IDENTITY盛り欄の近く
        "steps": [
            "① 下の IDENTITY盛り欄で、元画像（①）が正しい盛り見本か確認する",
            "② 無ければ候補から IDENTITY主を選び直す／盛り見本を差し替える",
            "③ ①の言語化を直す（配置・量感のみ。個品の色・筋は書かない）",
            "④ 盛りだけ合否。合格なら「深堀を閉じる」",
        ],
        "touch": "触ってよい: IDENTITY主・盛り言語化①",
        "dont": "触るな: 一片セットの色・筋、形状だけの作り直しを盛り不合格で全部やる",
        "resume": "閉じたあと: S1 を再実行して再検収",
    },
    "other": {
        "stage": "S?",
        "title": "深堀・その他",
        "failOne": "surface／shape／plate に当てはまらない不合格",
        "anchor": "top",
        "steps": [
            "① 下の要望欄に、何がダメかを1行書く",
            "② 可能なら surface／shape／plate のどれかに近いタグへ切り替える",
            "③ 切り替えられないときだけ、担当にこの画面のまま相談",
        ],
        "touch": "触ってよい: 要望1行・タグの切り替え",
        "dont": "触るな: 関係ないレーンを全部作り直す",
        "resume": "タグが決まったら、その深堀の手順に従う",
    },
}


def deepen_ui_enabled() -> bool:
    raw = str(os.environ.get("B5_PIECE_DEEPEN_UI") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def surface_shape_split_enabled() -> bool:
    raw = str(os.environ.get("B5_SURFACE_SHAPE_SPLIT") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def s1_omit_own_piece_enabled() -> bool:
    """S1 SUB: 自SKU一片本体を添付しない（仮想C／B45 S1）。

    切戻し: B5_S1_OMIT_OWN_PIECE=0（一片本体を再び同送）
    """
    if not surface_shape_split_enabled():
        return False
    raw = str(os.environ.get("B5_S1_OMIT_OWN_PIECE") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def surface_face_crop_fill_enabled() -> bool:
    """SURFACE枚数不足時に一片 face_crop で埋めるか。既定OFF。

    切戻し: B5_SURFACE_FACE_CROP=1
    """
    raw = str(os.environ.get("B5_SURFACE_FACE_CROP") or "0").strip().lower()
    return raw in ("1", "true", "yes", "on")


def deepen_path(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / DEEPEN_REL


def load_deepen(theme_new_dir: Path) -> Dict[str, Any]:
    p = deepen_path(theme_new_dir)
    if not p.is_file():
        return {"active": False, "tag": "", "note": "", "updatedAt": ""}
    try:
        obj = json.loads(p.read_text(encoding="utf-8-sig"))
        return obj if isinstance(obj, dict) else {"active": False, "tag": ""}
    except Exception:
        return {"active": False, "tag": ""}


def refs_dir(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / REFS_REL


def prev_dir(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / PREV_REL


def after_dir(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / AFTER_REL


def set_dir(theme_new_dir: Path) -> Path:
    return Path(theme_new_dir) / SET_REL


def _copy_img(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)


def snapshot_prev_from_piece_set(theme_new_dir: Path, *, force: bool = False) -> List[str]:
    """現行 piece_set を前回段へ退避（残す）。forceで上書き。"""
    src_d = set_dir(theme_new_dir)
    dst_d = prev_dir(theme_new_dir)
    dst_d.mkdir(parents=True, exist_ok=True)
    existing = list(dst_d.glob("prev_*.jpg"))
    if existing and not force:
        return [p.name for p in sorted(existing)]
    for old in existing:
        try:
            old.unlink()
        except Exception:
            pass
    names: List[str] = []
    files = sorted(src_d.glob("piece_set_0*.jpg")) if src_d.is_dir() else []
    for i, src in enumerate(files, 1):
        dest = dst_d / f"prev_{i:02d}.jpg"
        _copy_img(src, dest)
        names.append(dest.name)
    return names


def _resolve_active_seed(theme_new_dir: Path) -> Optional[Path]:
    try:
        from b5_piece_gate import list_piece_seed_display

        seed = list_piece_seed_display(theme_new_dir)
        items = list(seed.get("items") or [])
        active = ""
        src_json = Path(theme_new_dir) / "food_piece_set5_virt" / "piece_set_source.json"
        if src_json.is_file():
            try:
                sobj = json.loads(src_json.read_text(encoding="utf-8-sig"))
                if isinstance(sobj, dict):
                    active = str(sobj.get("activeSeed") or "").strip()
            except Exception:
                pass
        if not active and items:
            active = str(items[0].get("name") or "")
        for it in items:
            if str(it.get("name") or "") == active:
                rel = str(it.get("rel") or "").strip()
                if rel:
                    p = Path(theme_new_dir) / rel.replace("/", os.sep)
                    if p.is_file():
                        return p
        # fallback: 01_refs の種
        refs = Path(theme_new_dir) / "01_refs"
        for pat in ("PIECE_SEED_HUMAN_0*.*", "PIECE_SEED_HUMAN.*"):
            hits = sorted(refs.glob(pat)) if refs.is_dir() else []
            if hits:
                return hits[0]
    except Exception:
        pass
    return None


def _find_chiai_ref(theme_new_dir: Path) -> Optional[Path]:
    """血合い拡大見本。専用ファイルがあれば優先。"""
    refs = Path(theme_new_dir) / "01_refs"
    cands: List[Path] = []
    if refs.is_dir():
        for pat in (
            "*CHIAI*",
            "*chiai*",
            "*血合*",
            "PIECE_CHIAI*",
            "ref_chiai*",
        ):
            cands.extend(refs.glob(pat))
    rd = refs_dir(theme_new_dir)
    if rd.is_dir():
        cands.extend(rd.glob("ref_chiai*"))
    for p in cands:
        if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp"):
            return p
    return None


def ensure_shape_donors(theme_new_dir: Path, *, max_n: int = 3) -> Dict[str, Any]:
    """同ファミリー別SKUの一片を deepen_refs へ自動コピー（形状ドナー）。"""
    out: Dict[str, Any] = {
        "ok": False,
        "names": [],
        "family": "",
        "donorJans": [],
        "reason": "",
    }
    if not surface_shape_split_enabled():
        out["reason"] = "B5_SURFACE_SHAPE_SPLIT=0"
        return out
    try:
        from b5_piece_truth_rules import resolve_shape_donor_paths
    except Exception as e:
        out["reason"] = f"import: {e}"
        return out
    meta = resolve_shape_donor_paths(Path(theme_new_dir), max_n=int(max_n or 3))
    out["family"] = str(meta.get("family") or "")
    out["donorJans"] = list(meta.get("donorJans") or [])
    if not meta.get("ok"):
        out["reason"] = str(meta.get("reason") or "donorなし")
        return out
    rd = refs_dir(theme_new_dir)
    rd.mkdir(parents=True, exist_ok=True)
    # 旧 shape 参考を掃除してから載せ替え
    for old in rd.glob("ref_shape_*"):
        try:
            old.unlink()
        except Exception:
            pass
    names: List[str] = []
    for i, src in enumerate(list(meta.get("paths") or []), 1):
        src_p = Path(src)
        if not src_p.is_file():
            continue
        dest = rd / f"ref_shape_{i:02d}.jpg"
        _copy_img(src_p, dest)
        names.append(dest.name)
    out["ok"] = bool(names)
    out["names"] = names
    out["n"] = len(names)
    if not names:
        out["reason"] = "コピー先に書けない"
    return out


def ensure_refs(theme_new_dir: Path, *, tag: str = "") -> Dict[str, Any]:
    """生成参考を残す。表面=種、血合い=専用見本、形状=同ファミリー別SKU。"""
    rd = refs_dir(theme_new_dir)
    rd.mkdir(parents=True, exist_ok=True)
    out: Dict[str, Any] = {}
    seed = _resolve_active_seed(theme_new_dir)
    if seed and seed.is_file():
        dest = rd / "ref_surface_close.jpg"
        _copy_img(seed, dest)
        out["surface"] = dest.name
    chiai = _find_chiai_ref(theme_new_dir)
    if chiai and chiai.is_file():
        dest = rd / "ref_chiai_close.jpg"
        # 既に同じ内容ならそのまま
        _copy_img(chiai, dest)
        out["chiai"] = dest.name
    # 形状ドナーは split ON のとき常に（surface深堀でも形を残す）
    shape_meta = ensure_shape_donors(theme_new_dir, max_n=3)
    out["shape"] = shape_meta
    out["tag"] = str(tag or "")
    return out


def ensure_deepen_stack(
    theme_new_dir: Path,
    *,
    tag: str,
    refresh_prev: bool = False,
    clear_after: bool = False,
) -> Dict[str, Any]:
    """深堀開始時: 参考＋前回を自動構成。下段は作り直しでのみ増える。"""
    theme_new_dir = Path(theme_new_dir)
    tag_n = str(tag or "").strip().lower()
    prev_names = snapshot_prev_from_piece_set(theme_new_dir, force=refresh_prev)
    ref_meta = ensure_refs(theme_new_dir, tag=tag_n)
    ad = after_dir(theme_new_dir)
    ad.mkdir(parents=True, exist_ok=True)
    if clear_after:
        for p in ad.glob("after_*.jpg"):
            try:
                p.unlink()
            except Exception:
                pass
    return {
        "prev": prev_names,
        "refs": ref_meta,
        "afterCount": len(list(ad.glob("after_*.jpg"))),
        "tag": tag_n,
    }


def append_after_from_path(
    theme_new_dir: Path,
    src: Path,
    *,
    slot: int = 0,
) -> Path:
    """作り直し結果を下段へ。piece_set は触らない（呼び出し側で戻す）。"""
    src = Path(src)
    if not src.is_file():
        raise ValueError(f"下段へ載せる画像がありません: {src}")
    ad = after_dir(theme_new_dir)
    ad.mkdir(parents=True, exist_ok=True)
    n = int(slot or 0)
    if n < 1:
        existing = sorted(ad.glob("after_*.jpg"))
        n = len(existing) + 1
    dest = ad / f"after_{n:02d}.jpg"
    _copy_img(src, dest)
    return dest


def restore_piece_set_slot_from_prev(theme_new_dir: Path, slot: int) -> bool:
    """深堀中の作り直しで上書きされた piece_set を前回段から戻す。"""
    slot = int(slot or 0)
    if slot < 1:
        return False
    prev = prev_dir(theme_new_dir) / f"prev_{slot:02d}.jpg"
    if not prev.is_file():
        # prev_01 連番と piece_set 番号がずれても、同番号優先
        alt = list(prev_dir(theme_new_dir).glob("prev_*.jpg"))
        if len(alt) >= slot:
            prev = sorted(alt)[slot - 1]
        else:
            return False
    dest = set_dir(theme_new_dir) / f"piece_set_{slot:02d}.jpg"
    dest.parent.mkdir(parents=True, exist_ok=True)
    _copy_img(prev, dest)
    try:
        from b5_piece_gate import human_truth_dir

        ht = human_truth_dir(theme_new_dir) / f"piece_truth_{slot:02d}.jpg"
        ht.parent.mkdir(parents=True, exist_ok=True)
        _copy_img(prev, ht)
    except Exception:
        pass
    return True


def record_remake_to_after(
    theme_new_dir: Path,
    *,
    set_name: str = "",
    truth_path: str = "",
    slot: int = 0,
) -> Dict[str, Any]:
    """深堀中の remake 結果を下段へ。本線 piece_set は前回へ戻す。"""
    st = load_deepen(theme_new_dir)
    if not st.get("active"):
        return {"skipped": True, "reason": "deepen inactive"}
    slot_n = int(slot or 0)
    src: Optional[Path] = None
    if truth_path and Path(truth_path).is_file():
        src = Path(truth_path)
    elif set_name:
        p = set_dir(theme_new_dir) / set_name
        if p.is_file():
            src = p
    if src is None and slot_n > 0:
        p = set_dir(theme_new_dir) / f"piece_set_{slot_n:02d}.jpg"
        if p.is_file():
            src = p
    if src is None:
        return {"ok": False, "error": "remake結果ファイルが無い"}
    # 先に下段へコピー（戻す前）
    after_path = append_after_from_path(theme_new_dir, src, slot=slot_n)
    restored = restore_piece_set_slot_from_prev(theme_new_dir, slot_n)
    return {
        "ok": True,
        "after": after_path.name,
        "restoredPrev": restored,
        "slot": slot_n,
    }


def adopt_after_to_piece_set(theme_new_dir: Path) -> Dict[str, Any]:
    """下段の深堀結果を piece_set 本線へ採用（差分）。"""
    theme_new_dir = Path(theme_new_dir)
    ad = after_dir(theme_new_dir)
    files = sorted(ad.glob("after_*.jpg"))
    if not files:
        raise ValueError("下段（深堀後）に採用できる画像がありません")
    sd = set_dir(theme_new_dir)
    sd.mkdir(parents=True, exist_ok=True)
    adopted: List[str] = []
    for i, src in enumerate(files, 1):
        dest = sd / f"piece_set_{i:02d}.jpg"
        _copy_img(src, dest)
        adopted.append(dest.name)
        try:
            from b5_piece_gate import human_truth_dir, library_dir

            ht = human_truth_dir(theme_new_dir) / f"piece_truth_{i:02d}.jpg"
            ht.parent.mkdir(parents=True, exist_ok=True)
            _copy_img(src, ht)
            lib = library_dir(theme_new_dir)
            lib.mkdir(parents=True, exist_ok=True)
            _copy_img(src, lib / f"piece_lib_{i:02d}.jpg")
        except Exception:
            pass
    # スナップショットを更新（採用後が新しい「前回」）
    snapshot_prev_from_piece_set(theme_new_dir, force=True)
    meta = {
        "ok": True,
        "adopted": adopted,
        "n": len(adopted),
        "adoptedAt": datetime.now(timezone.utc).isoformat(),
    }
    st = load_deepen(theme_new_dir)
    st["lastAdopt"] = meta
    deepen_path(theme_new_dir).write_text(
        json.dumps(st, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return meta


def save_deepen(
    theme_new_dir: Path,
    *,
    tag: str = "",
    note: str = "",
    active: Optional[bool] = None,
    clear: bool = False,
) -> Dict[str, Any]:
    theme_new_dir = Path(theme_new_dir)
    p = deepen_path(theme_new_dir)
    p.parent.mkdir(parents=True, exist_ok=True)
    if clear:
        obj = {
            "active": False,
            "tag": "",
            "note": "",
            "updatedAt": datetime.now(timezone.utc).isoformat(),
            "clearedAt": datetime.now(timezone.utc).isoformat(),
        }
        p.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
        return obj
    tag_n = str(tag or "").strip().lower()
    if tag_n and tag_n not in VALID_TAGS:
        raise ValueError(f"不明なタグ: {tag_n}（surface/shape/plate/other）")
    prev = load_deepen(theme_new_dir)
    use_tag = tag_n or str(prev.get("tag") or "").strip().lower()
    if active is None:
        active_v = bool(use_tag)
    else:
        active_v = bool(active)
    if active_v and not use_tag:
        raise ValueError("深堀を開くにはタグが必要です")
    tag_changed = bool(active_v and use_tag and use_tag != str(prev.get("tag") or ""))
    newly_opened = bool(active_v and not prev.get("active"))
    stack_meta: Dict[str, Any] = {}
    if active_v:
        stack_meta = ensure_deepen_stack(
            theme_new_dir,
            tag=use_tag,
            refresh_prev=newly_opened or tag_changed,
            clear_after=tag_changed,
        )
    obj = {
        "active": active_v,
        "tag": use_tag if active_v else "",
        "note": str(note if note is not None else prev.get("note") or "").strip()[:400],
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "stage": str((_STEPS.get(use_tag) or {}).get("stage") or ""),
        "stack": stack_meta,
    }
    p.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    return obj


def surface_shape_role_prompt_block() -> str:
    """SUB／場面向け。表面と形状の役割分離（T4）。フラグOFFなら空。"""
    if not surface_shape_split_enabled():
        return ""
    return (
        "【表面／形状・役割分離・必須】\n"
        "SURFACE添付（面クロップ／詳細個品の面）＝筋・色・血合い色・表面感だけ。"
        "外形・反り・丸まり・ねじれ・細長さはSURFACEから取るな。\n"
        "SHAPE添付（同ファミリー別SKU一片）＝細長さ・反り・丸まり・ねじれ・薄さの形だけ。"
        "SHAPEの色・筋・縁の質感を自SKUの正にするな。\n"
        "SURFACE添付（面アップ／面クロップ）があるときは筋・色はSURFACEが正。\n"
        "PLATE／IDENTITY＝盛り配置・量感・ばらけだけ。個品の色・筋の正にしない。\n"
        "板状台形の一片シルエットをそのまま皿に並べるな。"
        "細長い削り片の形バリエを出せ。全部同じ曲がりにするな。\n"
    )


def _panel_html(
    *,
    tag: str,
    note: str = "",
    show_picker: bool = False,
) -> str:
    info = _STEPS.get(tag) or _STEPS["other"]
    steps = "".join(f"<li>{html_lib.escape(s)}</li>" for s in info["steps"])
    note_esc = html_lib.escape(note or "")
    picker = ""
    if show_picker:
        btns = []
        labs = {
            "surface": "表面がダメ",
            "shape": "形がダメ",
            "plate": "盛りがダメ",
            "other": "その他",
        }
        for t in VALID_TAGS:
            btns.append(
                f'<button type="button" class="btnDeepenTag" data-tag="{t}">'
                f"{html_lib.escape(labs[t])}</button>"
            )
        picker = (
            '<p class="deepen-pick"><b>不合格の面を選ぶ</b>（選ぶとこの付近に手順が出ます）<br/>'
            + " ".join(btns)
            + "</p>"
        )
    close_btn = (
        '<button type="button" class="btnDeepenClear">深堀を閉じる（浅い段へ戻る）</button>'
        if tag
        else ""
    )
    return f"""
<div class="deepen-panel" data-deepen-tag="{html_lib.escape(tag or '')}" data-deepen-anchor="{html_lib.escape(str(info.get('anchor') or 'top'))}">
  <div class="deepen-title">{html_lib.escape(str(info['title']))}</div>
  <p class="deepen-fail"><b>いま不合格だった面:</b> {html_lib.escape(str(info['failOne']))}</p>
  {picker}
  <ol class="deepen-steps">{steps}</ol>
  <p class="deepen-touch"><b>{html_lib.escape(str(info['touch']))}</b></p>
  <p class="deepen-dont"><b>{html_lib.escape(str(info['dont']))}</b></p>
  <p class="deepen-resume"><b>終わったら:</b> {html_lib.escape(str(info['resume']))}</p>
  <label class="deepen-note-lab">修正要望（1行・任意）</label>
  <textarea class="deepenNote" rows="2" maxlength="400" placeholder="例: 面の白い締まり線が消えている">{note_esc}</textarea>
  <p class="deepen-actions">
    <button type="button" class="btnDeepenSaveNote">要望を保存</button>
    {close_btn}
    <span class="deepenStatus"></span>
  </p>
</div>
"""


def deepen_css() -> str:
    return """
.deepen-panel{margin:12px 0;padding:14px 16px;border:3px solid #c45;border-radius:10px;background:#fff5f5;font-size:14px;line-height:1.45}
.deepen-panel .deepen-title{font-size:18px;font-weight:700;color:#a22;margin:0 0 8px}
.deepen-panel .deepen-fail{margin:0 0 8px}
.deepen-panel .deepen-steps{margin:8px 0 8px 1.2em;padding:0}
.deepen-panel .deepen-steps li{margin:4px 0}
.deepen-panel .deepen-touch,.deepen-panel .deepen-dont,.deepen-panel .deepen-resume{margin:6px 0}
.deepen-panel .deepen-dont{color:#840}
.deepen-panel textarea.deepenNote{width:100%;max-width:640px;box-sizing:border-box;padding:8px;font-size:13px}
.deepen-panel .btnDeepenTag,.deepen-panel .btnDeepenClear,.deepen-panel .btnDeepenSaveNote{margin:4px 6px 4px 0;padding:8px 12px;font-size:14px;cursor:pointer}
.deepen-panel .btnDeepenTag{background:#fff;border:2px solid #c45;border-radius:6px;font-weight:600}
.deepen-idle{margin:12px 0;padding:12px;border:2px dashed #c45;border-radius:8px;background:#fffaf8}
.deepen-idle .deepen-title{font-size:16px;font-weight:700;color:#a22;margin:0 0 6px}
.deepen-stack{margin:16px 0;padding:14px;border:3px solid #1565c0;border-radius:10px;background:#f3f8ff}
.deepen-stack-h{font-size:17px;margin:0 0 8px;color:#0d47a1}
.deepen-stack-h3{font-size:14px;margin:14px 0 6px;color:#333}
.deepen-stack-h3.deepen-stack-after{color:#a22;border-top:2px dashed #c45;padding-top:12px}
.deepen-stack-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px}
.deepen-stack-after-grid{border:2px solid #c45;border-radius:8px;padding:10px;background:#fff5f5}
.deepen-stack-card{background:#fff;border:1px solid #bbb;border-radius:6px;padding:6px}
.deepen-stack-card img{width:100%;height:140px;object-fit:contain;background:#f0ece6;border:1px solid #eee;display:block;cursor:zoom-in}
"""


def deepen_js() -> str:
    return """
async function deepenPost(body){
  const res=await fetch('/api/piece-deepen',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  const j=await res.json();
  return j;
}
function deepenBind(){
  document.querySelectorAll('.btnDeepenTag').forEach(btn=>{
    btn.onclick=async()=>{
      const tag=btn.getAttribute('data-tag')||'';
      const noteEl=btn.closest('.deepen-panel,.deepen-idle')?.querySelector('.deepenNote');
      const note=noteEl?noteEl.value:'';
      const st=btn.closest('.deepen-panel,.deepen-idle')?.querySelector('.deepenStatus');
      if(st) st.textContent='保存中…';
      const j=await deepenPost({tag, note, active:true});
      if(!j.ok){ if(st) st.textContent=j.message||'失敗'; alert(j.message||'深堀を開けません'); return; }
      location.reload();
    };
  });
  document.querySelectorAll('.btnDeepenClear').forEach(btn=>{
    btn.onclick=async()=>{
      if(!confirm('深堀を閉じます。浅い段の再検収に戻ります。よろしいですか？')) return;
      const st=btn.closest('.deepen-panel')?.querySelector('.deepenStatus');
      if(st) st.textContent='閉じる…';
      const j=await deepenPost({clear:true});
      if(!j.ok){ if(st) st.textContent=j.message||'失敗'; return; }
      location.reload();
    };
  });
  document.querySelectorAll('.btnDeepenSaveNote').forEach(btn=>{
    btn.onclick=async()=>{
      const box=btn.closest('.deepen-panel,.deepen-idle');
      const note=box?.querySelector('.deepenNote')?.value||'';
      const tag=box?.getAttribute('data-deepen-tag')||'';
      const st=box?.querySelector('.deepenStatus');
      if(st) st.textContent='保存中…';
      const j=await deepenPost({tag, note, active: !!tag});
      if(st) st.textContent=j.ok?'要望を保存しました':(j.message||'失敗');
    };
  });
  document.querySelectorAll('.btnDeepenAdopt').forEach(btn=>{
    btn.onclick=async()=>{
      if(!confirm('下段の深堀結果を一片セット本線へ採用します。よろしいですか？')) return;
      const st=btn.closest('.deepen-stack,.deepen-adopt')?.querySelector('.deepenStatus')
        || document.querySelector('.deepen-stack .deepenStatus');
      if(st) st.textContent='採用中…';
      const j=await deepenPost({adopt:true});
      if(!j.ok){ if(st) st.textContent=j.message||'失敗'; alert(j.message||'採用に失敗'); return; }
      if(st) st.textContent=j.message||'採用しました';
      location.reload();
    };
  });
}
if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', deepenBind);
else deepenBind();
"""


def idle_picker_html() -> str:
    """深堀未開始時。一片ゲート上部に小さく置き、タグ選択で本パネルへ。"""
    btns = []
    labs = {
        "surface": "表面がダメ → 深堀",
        "shape": "形がダメ → 深堀",
        "plate": "盛りがダメ → 深堀",
        "other": "その他 → 深堀",
    }
    for t in VALID_TAGS:
        btns.append(
            f'<button type="button" class="btnDeepenTag" data-tag="{t}">'
            f"{html_lib.escape(labs[t])}</button>"
        )
    return f"""
<div class="deepen-idle" id="deepenIdle">
  <div class="deepen-title">不合格のとき（深堀）</div>
  <p>基本は自動の浅い段です。ダメな面だけ選ぶと、<b>その作業の近く</b>に大きめの次手順が出ます。</p>
  <p class="deepen-pick">{" ".join(btns)}</p>
  <span class="deepenStatus"></span>
</div>
"""


def _list_imgs(folder: Path, patterns: tuple[str, ...]) -> List[Path]:
    if not folder.is_dir():
        return []
    out: List[Path] = []
    for pat in patterns:
        out.extend(folder.glob(pat))
    return sorted({p.resolve() for p in out if p.is_file()}, key=lambda p: p.name)


def deepen_stack_html(theme_new_dir: Path) -> str:
    """参考見本（残す）→ 前回生成（残す）→ 深堀後（下段）。フォルダがあれば出す。"""
    theme_new_dir = Path(theme_new_dir)
    base = theme_new_dir / "food_piece_set5_virt"
    refs = _list_imgs(base / "deepen_refs", ("ref_*.jpg", "ref_*.png", "*.jpg", "*.png"))
    prevs = _list_imgs(base / "deepen_prev_pieces", ("prev_*.jpg", "*.jpg"))
    afters = _list_imgs(base / "deepen_after_pieces", ("after_*.jpg", "*.jpg"))
    if not (refs or prevs or afters):
        return ""

    def cards(paths: List[Path], rel_prefix: str, label_prefix: str) -> str:
        bits: List[str] = []
        for p in paths:
            rel = f"{rel_prefix}/{p.name}".replace("\\", "/")
            lab = html_lib.escape(f"{label_prefix} {p.stem}")
            bits.append(
                f'<div class="deepen-stack-card">'
                f'<img src="/file/{html_lib.escape(rel)}" class="hh-img piece-img" alt="" title="クリックで拡大"/>'
                f'<div class="cand-meta"><b>{lab}</b><br/><code>{html_lib.escape(p.name)}</code></div>'
                f"</div>"
            )
        return "\n".join(bits) or "<p class='hint'>（なし）</p>"

    ref_cards = cards(refs, "food_piece_set5_virt/deepen_refs", "参考")
    prev_cards = cards(prevs, "food_piece_set5_virt/deepen_prev_pieces", "前回生成")
    after_cards = cards(afters, "food_piece_set5_virt/deepen_after_pieces", "深堀後")
    return f"""
<section class="deepen-stack" id="deepenStack">
  <h2 class="deepen-stack-h">深堀の見え方（参考は残す → 前回生成は残す → 下段が深堀結果）</h2>
  <p class="hint">生成の参考にした拡大見本は消さない。前回の一片生成も残し、表面などを変えた結果だけ下の段に出す。</p>
  <h3 class="deepen-stack-h3">A. 生成の参考（残す）— 表面／血合い／形状ドナー（同ファミリー別SKU）</h3>
  <div class="deepen-stack-grid">{ref_cards}</div>
  <h3 class="deepen-stack-h3">B. 前回つかった一片の生成画像（残す）</h3>
  <div class="deepen-stack-grid">{prev_cards}</div>
  <h3 class="deepen-stack-h3 deepen-stack-after">C. 下段・深堀後（例: 表面を変えたもの）</h3>
  <div class="deepen-stack-grid deepen-stack-after-grid">{after_cards}</div>
  <p class="deepen-adopt" style="margin-top:10px">
    <button type="button" class="btnDeepenAdopt" style="font-weight:bold;padding:10px 14px;font-size:14px">
      下段を一片セットへ採用（本線へ）
    </button>
    <span class="deepenStatus"></span>
  </p>
  <p class="hint">採用すると C が ③一片セットの正になります。B（前回）は消さず残ります。全部作り直しはしません。</p>
</section>
"""


def panels_for_theme(theme_new_dir: Path) -> Dict[str, str]:
    """anchor → HTML。active でなければ idle のみ top。スタックがあれば常に top へ載せる。"""
    if not deepen_ui_enabled():
        return {"top": "", "seed": "", "set": "", "plate": ""}
    st = load_deepen(theme_new_dir)
    active = bool(st.get("active"))
    tag = str(st.get("tag") or "").strip().lower()
    note = str(st.get("note") or "")
    out = {"top": "", "seed": "", "set": "", "plate": ""}
    stack = deepen_stack_html(theme_new_dir)
    if not active or tag not in VALID_TAGS:
        out["top"] = (stack + idle_picker_html()) if stack else idle_picker_html()
        return out
    info = _STEPS[tag]
    anchor = str(info.get("anchor") or "top")
    panel = _panel_html(tag=tag, note=note, show_picker=True)
    # スタックは常に一片ゲート直下（参考→前回→下段）
    head = stack
    if anchor != "top":
        head += (
            f'<p class="hint" id="deepenJumpHint" style="border-left:4px solid #c45;padding:8px 12px;background:#fff5f5">'
            f'<b>深堀中（{html_lib.escape(str(info["title"]))}）</b> — '
            f'手順は下の作業欄に出しています。'
            f'<a href="#deepenWorkAnchor">作業欄へ</a></p>'
        )
        out[anchor] = f'<div id="deepenWorkAnchor">{panel}</div>'
        out["top"] = head
    else:
        out["top"] = head + panel
    return out
