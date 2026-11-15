# -*- coding: utf-8 -*-
"""個品見本の見え方を機械が読む（人間は型を足さない）。"""
from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from gemini_image import load_api_key, make_client, mime_for
from sub_image_intent import TEXT_MODELS

LOG = logging.getLogger("set_main_image.b5_food_look")

_KEYS = ("色", "透け", "縁", "重なり", "曲がり", "光沢")
# 個包装見本（WRAP）。案件固定語は置かない。見えていることだけ。
_WRAP_KEYS = ("密着", "空気", "材質", "縁", "透明度", "形")


def _parse_obj(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        return None
    try:
        obj = json.loads(m.group(0))
    except Exception:
        return None
    return obj if isinstance(obj, dict) else None


def read_food_look(path: Path, *, label: str = "") -> Dict[str, Any]:
    """見本1枚から、見えていることだけ返す。数字の厚みは作らない。"""
    p = Path(path)
    out: Dict[str, Any] = {
        "label": str(label or p.stem),
        "path": str(p),
        "ok": False,
        "lines": "",
        "fields": {},
        "error": "",
    }
    if not p.is_file():
        out["error"] = "no_file"
        return out
    prompt = (
        "食品の個品見本写真です。写っている見た目だけを短く書いてください。\n"
        "推測・商品名・厚みの数字・キャッチの語は書くな。\n"
        "見えない項目は空文字。\n"
        'JSONのみ: {"色":"","透け":"","縁":"","重なり":"","曲がり":"","光沢":""}\n'
        "各値は全角20字以内。"
    )
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        out["error"] = str(e)
        return out
    mime = mime_for(p)
    b64 = p.read_bytes()
    import base64

    data = base64.b64encode(b64).decode("ascii")
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
                            {"inline_data": {"mime_type": mime, "data": data}},
                        ],
                    }
                ],
            )
            text = getattr(resp, "text", None) or ""
            obj = _parse_obj(text)
            if not obj:
                last = (text or "")[:160]
                continue
            fields = {k: str(obj.get(k) or "").strip()[:40] for k in _KEYS}
            bits = [f"{k}: {v}" for k, v in fields.items() if v]
            out["ok"] = bool(bits)
            out["fields"] = fields
            out["lines"] = "／".join(bits)
            out["model"] = mid
            return out
        except Exception as e:
            last = str(e)
            LOG.warning("food look fail model=%s: %s", mid, e)
    out["error"] = last or "parse_failed"
    return out


def check_drawn_foods(
    scene_path: Path, items: Sequence[Dict[str, Any]]
) -> Dict[str, Any]:
    """描いた1片が主見本の厚み・色・大きさと違うときだけ落とす。盛り付けの違いは落とさない。"""
    import base64

    out: Dict[str, Any] = {
        "ok": True,
        "skipped": False,
        "items": [],
        "retryNote": "",
        "error": "",
    }
    scene = Path(scene_path)
    usable = []
    for it in items:
        if not isinstance(it, dict):
            continue
        p = Path(str(it.get("path") or ""))
        if p.is_file():
            usable.append({"label": str(it.get("label") or p.stem), "path": p})
    if not scene.is_file() or not usable:
        out["skipped"] = True
        out["error"] = "no_file"
        return out
    prompt = (
        "完成画の皿の上の個品を、各主見本と比べてください。\n"
        "見るのは1片の厚み・色・質感・1片の大きさ。盛り付け（皿かご飯か、重ね方）が違うのは不合格にしない。\n"
        "主見本の写真を貼ったように見えるのも不合格。\n"
        "JSONのみ: "
        '{"items":[{"label":"見本名","samePiece":true,"note":""}]}\n'
        "samePiece が false のとき note に、厚み・色・大きさのどれが違いかを短く。"
    )
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        out["skipped"] = True
        out["error"] = str(e)
        return out
    parts: List[Dict[str, Any]] = [{"text": prompt}, {"text": "完成画"}]
    parts.append(
        {
            "inline_data": {
                "mime_type": mime_for(scene),
                "data": base64.b64encode(scene.read_bytes()).decode("ascii"),
            }
        }
    )
    for it in usable:
        parts.append({"text": "主見本 " + it["label"]})
        parts.append(
            {
                "inline_data": {
                    "mime_type": mime_for(it["path"]),
                    "data": base64.b64encode(it["path"].read_bytes()).decode("ascii"),
                }
            }
        )
    obj = None
    last = ""
    for mid in TEXT_MODELS:
        if not mid:
            continue
        try:
            resp = client.models.generate_content(
                model=mid,
                contents=[{"role": "user", "parts": parts}],
            )
            text = getattr(resp, "text", None) or ""
            obj = _parse_obj(text)
            if obj:
                out["model"] = mid
                break
            last = (text or "")[:160]
        except Exception as e:
            last = str(e)
            LOG.warning("food piece fail model=%s: %s", mid, e)
    if obj is None:
        out["skipped"] = True
        out["error"] = last or "parse_failed"
        return out
    notes = []
    raw_items = obj.get("items") if isinstance(obj.get("items"), list) else []
    for item in raw_items:
        if not isinstance(item, dict):
            continue
        same = item.get("samePiece")
        if isinstance(same, str):
            same = same.strip().lower() in ("1", "true", "yes", "はい")
        rec = {
            "label": str(item.get("label") or ""),
            "samePiece": bool(same),
            "note": str(item.get("note") or "")[:80],
        }
        out["items"].append(rec)
        if not rec["samePiece"]:
            notes.append(
                f"{rec['label']} の1片が主見本と違う（{rec['note'] or '厚み・色・大きさ'}）。"
                "厚み・色・質感・1片の大きさは主見本に戻せ。盛り付けは変えてよい。"
            )
    if not out["items"]:
        out["skipped"] = True
        out["error"] = "no_items"
        return out
    out["ok"] = not notes
    out["retryNote"] = "\n".join(notes)
    LOG.info("food piece ok=%s n=%s", out["ok"], len(out["items"]))
    return out


def look_block_ja(reads: Sequence[Dict[str, Any]]) -> str:
    rows = [r for r in reads if isinstance(r, dict) and str(r.get("lines") or "").strip()]
    if not rows:
        return ""
    lines = ["【見本の見え方・機械読み・正】"]
    lines.append("添付の個品見本写真が正。次は写真から読んだ見え方。厚みの数字ではない。")
    lines.append(
        "小皿の開封後も、添える料理の上の具も、この見本の色・質感・厚み・1片の形に合わせろ。"
        "一般トッピングや別形に寄せるな。"
    )
    lines.append("文章が空の項目は、見本の写真を見ろ。空だから無視するな。")
    for r in rows:
        lab = str(r.get("label") or "").strip() or "見本"
        lines.append(f"- {lab}: {r.get('lines')}")
        fields = r.get("fields") if isinstance(r.get("fields"), dict) else {}
        empty = [k for k in _KEYS if not str(fields.get(k) or "").strip()]
        if empty:
            lines.append(
                "  空: " + "、".join(empty) + "。この項目は文章が無い。見本の写真を見ろ。"
            )
    return "\n".join(lines) + "\n"


def read_wrap_look(path: Path, *, label: str = "") -> Dict[str, Any]:
    """個包装見本1枚から、個包装だけの見え方を返す。大袋・背景・手は書かない。"""
    p = Path(path)
    out: Dict[str, Any] = {
        "label": str(label or p.stem),
        "path": str(p),
        "ok": False,
        "lines": "",
        "fields": {},
        "error": "",
        "kind": "wrap",
    }
    if not p.is_file():
        out["error"] = "no_file"
        return out
    prompt = (
        "個包装（1個ずつ入った小包装）の見本写真です。\n"
        "個包装そのものの見た目だけを短く書いてください。\n"
        "大袋・背景・皿・手・文字・ロゴは無視して書くな。\n"
        "推測・商品名・キャッチの語は書くな。見えない項目は空文字。\n"
        "密着: 中身にフィルムが密着しているか、ゆとりがあるか。\n"
        "空気: 袋内に空気だまりが見えるか、見えないか。\n"
        "材質: 硬いフィルムに見えるか、やわらかい袋に見えるか（見えている質感だけ）。\n"
        "縁: シール縁・切れ目・ノッチなど、見える縁の特徴。\n"
        "透明度: 中身がどの程度透けて見えるか。\n"
        "形: 個包装の外形（正方形に近い／枕型 等、見えていることだけ）。\n"
        'JSONのみ: {"密着":"","空気":"","材質":"","縁":"","透明度":"","形":""}\n'
        "各値は全角20字以内。"
    )
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        out["error"] = str(e)
        return out
    import base64

    mime = mime_for(p)
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
                            {"inline_data": {"mime_type": mime, "data": data}},
                        ],
                    }
                ],
            )
            text = getattr(resp, "text", None) or ""
            obj = _parse_obj(text)
            if not obj:
                last = (text or "")[:160]
                continue
            fields = {k: str(obj.get(k) or "").strip()[:40] for k in _WRAP_KEYS}
            bits = [f"{k}: {v}" for k, v in fields.items() if v]
            out["ok"] = bool(bits)
            out["fields"] = fields
            out["lines"] = "／".join(bits)
            out["model"] = mid
            try:
                from b4_wrap_scene_refine import classify_wrap_film_branch

                out["filmBranch"] = classify_wrap_film_branch(fields)
            except Exception:
                out["filmBranch"] = "unknown"
            return out
        except Exception as e:
            last = str(e)
            LOG.warning("wrap look fail model=%s: %s", mid, e)
    out["error"] = last or "parse_failed"
    return out


def check_drawn_wraps(
    scene_path: Path,
    items: Sequence[Dict[str, Any]],
    looks: Optional[Sequence[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """個包装の空気・置き位置を見本と照らす。無い／読めないときは落とさない。"""
    import base64

    out: Dict[str, Any] = {
        "ok": True,
        "skipped": False,
        "airOk": True,
        "frontOk": True,
        "countOk": True,
        "blCornerOk": True,
        "filmClearOk": True,
        "needFilmClear": False,
        "wrapCount": None,
        "retryNote": "",
        "error": "",
    }
    scene = Path(scene_path)
    usable = []
    for it in items:
        if not isinstance(it, dict):
            continue
        p = Path(str(it.get("path") or ""))
        if p.is_file():
            usable.append({"label": str(it.get("label") or p.stem), "path": p})
    if not scene.is_file() or not usable:
        out["skipped"] = True
        out["error"] = "no_file"
        return out
    look_by_lab: Dict[str, Dict[str, Any]] = {}
    for rec in looks or []:
        if not isinstance(rec, dict):
            continue
        lab = str(rec.get("label") or "").strip()
        fields = rec.get("fields") if isinstance(rec.get("fields"), dict) else {}
        if lab:
            look_by_lab[lab] = fields
    air_hints = []
    need_film_clear = False
    for it in usable:
        fields = look_by_lab.get(it["label"]) or {}
        air = str(fields.get("空気") or "").strip()
        if air:
            air_hints.append(f"{it['label']}: 見本の空気＝{air}")
    for rec in looks or []:
        if not isinstance(rec, dict):
            continue
        br = str(rec.get("filmBranch") or "").strip()
        if not br:
            try:
                from b4_wrap_scene_refine import classify_wrap_film_branch

                fields = rec.get("fields") if isinstance(rec.get("fields"), dict) else {}
                br = classify_wrap_film_branch(fields)
            except Exception:
                br = "unknown"
        if br == "transparent":
            need_film_clear = True
            break
    out["needFilmClear"] = need_film_clear
    air_line = "／".join(air_hints) if air_hints else "（見本の空気欄は空。WRAP写真の見た目に合わせろ）"
    clear_line = (
        "6) filmClearOk: 見本が透明フィルムのとき必須。"
        "完成画の個包装フィルムが乳白色・曇り白・不透明白なら false。"
        "無色透明の硬い光沢なら true。白いのはハイライトだけ可。\n"
        if need_film_clear
        else "6) filmClearOk: 今回は true でよい（透明フィルム必須ではない）。\n"
    )
    prompt = (
        "完成画の個包装（開封前の小袋）を、各WRAP見本と比べてください。\n"
        "大袋・開封後の個品は個包装ではない。\n"
        "判定:\n"
        "1) airOk: 見本の空気の有無と完成画が一致するか。"
        "見本が空気なし・見えない・密着なら、完成画に空気だまりがあれば false。\n"
        "2) frontOk: カメラ最近傍（画面下端・手前）が開封後の個品なら true。"
        "個包装が小皿や開封後より手前／下端寄り／前レイヤなら false。"
        "皿の横・後ろだけなら true。迷ったら false。\n"
        "3) wrapCount: 見える個包装の個数（おおよそ）。\n"
        "4) countOk: wrapCount が 1 または 2 なら true。0や3以上は false。"
        "見本があるのに0も false。\n"
        "5) blCornerOk: 画面左下（幅の左約25%×下約25%、個包装金丸の聖域）に"
        "開封前の個包装が主役のように置かれていたら false。"
        "左下は金丸後載せの聖域。個包装は皿の横か後ろのみ。迷ったら false。\n"
        f"{clear_line}"
        f"見本の空気メモ: {air_line}\n"
        "JSONのみ: "
        '{"airOk":true,"frontOk":false,"countOk":true,"blCornerOk":true,'
        '"filmClearOk":true,"wrapCount":1,"note":""}\n'
        "不合格のとき note に短く理由。frontOk／blCornerOk は疑わしいとき false。"
    )
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        out["skipped"] = True
        out["error"] = str(e)
        return out
    parts: List[Dict[str, Any]] = [{"text": prompt}, {"text": "完成画"}]
    parts.append(
        {
            "inline_data": {
                "mime_type": mime_for(scene),
                "data": base64.b64encode(scene.read_bytes()).decode("ascii"),
            }
        }
    )
    for it in usable:
        parts.append({"text": "WRAP見本 " + it["label"]})
        parts.append(
            {
                "inline_data": {
                    "mime_type": mime_for(it["path"]),
                    "data": base64.b64encode(it["path"].read_bytes()).decode("ascii"),
                }
            }
        )
    obj = None
    last = ""
    for mid in TEXT_MODELS:
        if not mid:
            continue
        try:
            resp = client.models.generate_content(
                model=mid,
                contents=[{"role": "user", "parts": parts}],
            )
            text = getattr(resp, "text", None) or ""
            obj = _parse_obj(text)
            if obj:
                out["model"] = mid
                break
            last = (text or "")[:160]
        except Exception as e:
            last = str(e)
            LOG.warning("wrap check fail model=%s: %s", mid, e)
    if obj is None:
        out["skipped"] = True
        out["error"] = last or "parse_failed"
        return out

    def _b(key: str, default: bool = True) -> bool:
        v = obj.get(key)
        if isinstance(v, str):
            return v.strip().lower() in ("1", "true", "yes", "はい")
        if v is None:
            return default
        return bool(v)

    air_ok = _b("airOk", True)
    # 手前・左下は取りこぼしが多いので欠落時は不合格側へ倒す
    front_ok = _b("frontOk", False)
    count_ok = _b("countOk", True)
    bl_corner_ok = _b("blCornerOk", False)
    film_clear_ok = _b("filmClearOk", not need_film_clear)
    try:
        wrap_n = int(obj.get("wrapCount"))
    except (TypeError, ValueError):
        wrap_n = None
    note = str(obj.get("note") or "").strip()[:120]
    out["airOk"] = air_ok
    out["frontOk"] = front_ok
    out["countOk"] = count_ok
    out["blCornerOk"] = bl_corner_ok
    out["filmClearOk"] = film_clear_ok
    out["wrapCount"] = wrap_n
    notes: List[str] = []
    if not air_ok:
        notes.append(
            "個包装の空気が見本と違う。"
            "見本に無い空気だまりを作るな。密着・空気の有無は WRAP 見本どおり。"
            + (f"（{note}）" if note else "")
        )
    if not front_ok:
        notes.append(
            "個包装が開封後の個品より手前に出ている。"
            "個包装は皿の横か後ろのみ（小皿に混ぜるな）。カメラ最近傍は開封後の個品。個品より前に出すな。"
            + (f"（{note}）" if note and air_ok else "")
        )
    if not count_ok:
        notes.append(
            f"個包装の数が見本運用と違う（見える数={wrap_n}）。1〜2個にせよ。山にするな。"
        )
    if not bl_corner_ok:
        notes.append(
            "個包装が画面左下（個包装金丸の聖域）にある。"
            "左下に開封前の個包装を置くな。皿の横か後ろのみ。"
            + (f"（{note}）" if note and air_ok and front_ok else "")
        )
    if need_film_clear and not film_clear_ok:
        notes.append(
            "個包装フィルムが乳白色・曇り白。"
            "無色透明の硬い光沢にせよ（ハイライトの白だけ可）。"
            + (f"（{note}）" if note and air_ok and front_ok and bl_corner_ok else "")
        )
    out["ok"] = not notes
    out["retryNote"] = "\n".join(notes)
    LOG.info(
        "wrap check ok=%s air=%s front=%s count=%s bl=%s clear=%s needClear=%s n=%s",
        out["ok"],
        air_ok,
        front_ok,
        count_ok,
        bl_corner_ok,
        film_clear_ok,
        need_film_clear,
        wrap_n,
    )
    return out


def wrap_look_block_ja(reads: Sequence[Dict[str, Any]]) -> str:
    rows = [r for r in reads if isinstance(r, dict) and str(r.get("lines") or "").strip()]
    if not rows:
        return ""
    lines = [
        "【個包装見本の見え方・機械読み・正】",
        "添付の IMAGE_FOOD_UNIT_WRAP が正。次は写真から読んだ個包装だけの見え方。",
        "この読みに合わせて個包装を描け。空気の有無・材質の硬さ／やわらかさは読みどおり。",
        "置き位置: 皿の横か後ろのみ（小皿に混ぜるな）。カメラ最近傍は開封後の個品。個包装はそれより手前に出すな。見本に無い空気だまり禁止。",
        "画面左下（個包装金丸聖域）に開封前の個包装を置くな。左下は背景の続きのみ。",
        "大袋の色や形には使うな。文章が空の項目は、WRAP写真の個包装を見ろ。",
    ]
    for r in rows:
        lab = str(r.get("label") or "").strip() or "個包装"
        br = str(r.get("filmBranch") or "").strip()
        head = f"- {lab}: {r.get('lines')}"
        if br:
            head += f"／分岐={br}"
        lines.append(head)
        if br == "transparent":
            lines.append(
                "  透明分岐: フィルムは無色透明の硬い光沢。乳白色・曇り白・不透明白禁止。"
            )
        fields = r.get("fields") if isinstance(r.get("fields"), dict) else {}
        empty = [k for k in _WRAP_KEYS if not str(fields.get(k) or "").strip()]
        if empty:
            lines.append(
                "  空: " + "、".join(empty) + "。この項目は文章が無い。WRAP写真を見ろ。"
            )
    return "\n".join(lines) + "\n"


def read_looks_for_paths(
    items: Sequence[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        p = Path(str(it.get("path") or ""))
        lab = str(it.get("label") or "")
        if not p.is_file():
            continue
        rec = read_food_look(p, label=lab)
        LOG.info("food look %s ok=%s lines=%s", lab, rec.get("ok"), rec.get("lines"))
        out.append(rec)
    return out


def read_wrap_looks_for_paths(
    items: Sequence[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        p = Path(str(it.get("path") or ""))
        lab = str(it.get("label") or "")
        if not p.is_file():
            continue
        rec = read_wrap_look(p, label=lab)
        LOG.info("wrap look %s ok=%s lines=%s", lab, rec.get("ok"), rec.get("lines"))
        out.append(rec)
    return out
