# -*- coding: utf-8 -*-
"""袋ファイルの縦÷横を測り、描いた袋が外れていたら差を書いて描き直す。商品名は入れない。"""
from __future__ import annotations

import base64
import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from PIL import Image

from gemini_image import load_api_key, make_client, mime_for
from sub_image_intent import TEXT_MODELS

LOG = logging.getLogger("set_main_image.b5_bag_aspect")

_ROLE_OK = re.compile(r"^IMAGE_PACKAGE_(TRUTH|PART_\d+)$")


_TOL_MAX = 0.03
_BOWL_CM = 12.0
_PLATE_CM = 15.0


def aspect_tolerance() -> float:
    """包材の縦横・高さの許容。上限は3%。それより大きい指定は3%に下げる。"""
    raw = str(os.environ.get("B5_BAG_ASPECT_TOL") or "").strip()
    if not raw:
        return _TOL_MAX
    try:
        v = float(raw)
    except ValueError:
        return _TOL_MAX
    if v <= 0:
        return _TOL_MAX
    return min(v, _TOL_MAX)


def face_cm_prompt_block(height_cm: Optional[float], width_cm: Optional[float]) -> str:
    """面の縦横が両方あるときだけ、食器との比を書く。数字は入力。無いときは空。"""
    try:
        h = float(height_cm or 0)
        w = float(width_cm or 0)
    except (TypeError, ValueError):
        return ""
    if h <= 0 or w <= 0:
        return ""
    bowl = h / _BOWL_CM
    plate = h / _PLATE_CM
    return (
        "【包材の実寸と、同じ面の器】\n"
        f"この包材の面は縦{h:g}cm、横{w:g}cm。この数字は入力。無い商品には書かない。\n"
        "比べるのは、包材と同じ面（カメラから同じ距離）にある器だけ。手前の盛り付けは比に使うな。\n"
        f"同じ面にご飯茶碗があれば、口径は約{_BOWL_CM:g}cm。包材の高さはその約{bowl:.2f}倍。\n"
        f"同じ面に小皿があれば、口径は約{_PLATE_CM:g}cm。包材の高さはその約{plate:.2f}倍。\n"
        "同じ面に器が無ければ、この比は守らなくてよい。実寸の数字は画像に描くな。\n"
    )


def content_metrics(path: Path) -> Optional[Dict[str, float]]:
    """余白を除いた包材本体の縦と、縦÷横。透過が無ければ測らない。"""
    p = Path(path)
    if not p.is_file():
        return None
    im = Image.open(p).convert("RGBA")
    alpha = im.getchannel("A")
    transparent = sum(alpha.histogram()[:17])
    if transparent < im.size[0] * im.size[1] * 0.02:
        return None
    bb = alpha.point(lambda px: 255 if px > 16 else 0).getbbox()
    if not bb:
        return None
    w = bb[2] - bb[0]
    h = bb[3] - bb[1]
    if w < 8 or h < 8:
        return None
    return {"contentH": float(h), "ratio": round(h / float(w), 2)}


def content_hw_ratio(path: Path) -> Optional[float]:
    met = content_metrics(path)
    if not met:
        return None
    return float(met["ratio"])


def measure_role_aspects(
    image_paths: Sequence[Path], image_roles: Sequence[str]
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for path, role in zip(image_paths, image_roles):
        if not _ROLE_OK.match(str(role or "")):
            continue
        met = content_metrics(Path(path))
        if met is None:
            LOG.info("bag aspect skip role=%s (no alpha bbox)", role)
            continue
        rows.append(
            {
                "role": str(role),
                "path": str(path),
                "ratio": met["ratio"],
                "contentH": met["contentH"],
            }
        )
        LOG.info("bag aspect role=%s h/w=%s contentH=%s", role, met["ratio"], met["contentH"])
    return rows


def aspect_prompt_block(rows: Sequence[Dict[str, Any]]) -> str:
    if not rows:
        return ""
    from_scale = any(str(r.get("source") or "") == "scale" for r in rows)
    if from_scale:
        lines = [
            "【包材の縦横・並び正本（SCALE）実測】",
            "数字は、並び正本から切った各袋の縦÷横。単体ファイルの大きさは使うな。袋・缶・瓶・箱のどれでも同じ。商品名は関係ない。",
            "輪郭は並び正本の縁のまま。波・しわ・シール縁を、直線の四角に伸ばすな。",
        ]
    else:
        lines = [
            "【包材の縦横・ファイル実測】",
            "数字は、そのファイルの包材本体（余白を除く）の縦÷横。袋・缶・瓶・箱のどれでも同じ。商品名は関係ない。",
            "輪郭はファイルの縁のまま。波・しわ・シール縁を、直線の四角に伸ばすな。",
        ]
    for row in rows:
        lines.append(
            f"- {row['role']} は縦が横の {float(row['ratio']):.2f} 倍。"
            "これより縦に伸ばすな。横に潰すな。"
        )
    if len(rows) >= 2 and all(float(r.get("contentH") or 0) > 0 for r in rows):
        base = rows[0]
        if from_scale:
            lines.append(
                "包材どうしの高さは、並び正本上の見た目の比。片方だけ大きくするな。"
            )
        else:
            lines.append("包材どうしの高さは、ファイル本体の縦の比。片方だけ大きくするな。")
        for row in rows[1:]:
            rel = float(row["contentH"]) / float(base["contentH"])
            lines.append(
                f"- {row['role']} の高さは {base['role']} の {rel:.2f} 倍。"
            )
    return "\n".join(lines) + "\n"


def measure_mix_aspects_from_scale(
    scale_path: Path,
    part_roles: Sequence[str],
    *,
    out_dir: Optional[Path] = None,
) -> List[Dict[str, Any]]:
    """並び正本（SCALE）から各袋を切って縦横・高さを測る。単体ファイルは測らない。"""
    scale = Path(scale_path)
    roles = [str(r) for r in part_roles if _ROLE_OK.match(str(r or "")) and "PART_" in str(r)]
    if not scale.is_file() or len(roles) < 2:
        return []
    try:
        im = Image.open(scale).convert("RGBA")
    except Exception as e:
        LOG.warning("scale open fail: %s", e)
        return []
    w0, h0 = im.size
    n = len(roles)
    prompt = (
        "この写真には商品の袋が複数あります。左から右の順に各袋の外接矩形を返してください。\n"
        "重なっていても、袋ごとの外枠。座標は画像左上原点のピクセル。JSONのみ:\n"
        '{"bags":[{"i":1,"x0":0,"y0":0,"x1":10,"y1":10}]}\n'
        f"袋の数はちょうど {n}。画像サイズは横{w0}縦{h0}。"
    )
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        LOG.warning("scale measure no client: %s", e)
        return []
    data = base64.b64encode(scale.read_bytes()).decode("ascii")
    mime = mime_for(scale)
    obj = None
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
            if obj:
                LOG.info("scale bags model=%s", mid)
                break
        except Exception as e:
            LOG.warning("scale bags fail model=%s: %s", mid, e)
    bags = obj.get("bags") if isinstance(obj, dict) else None
    if not isinstance(bags, list) or len(bags) < n:
        LOG.warning("scale bags parse fail n=%s got=%s", n, bags)
        return []
    ordered = sorted(
        [b for b in bags if isinstance(b, dict)],
        key=lambda b: (int(b.get("i") or 0) or 999, float(b.get("x0") or 0)),
    )[:n]
    if len(ordered) < n:
        return []
    out_dir_p = Path(out_dir) if out_dir else None
    if out_dir_p is not None:
        out_dir_p.mkdir(parents=True, exist_ok=True)
    rows: List[Dict[str, Any]] = []
    for i, role in enumerate(roles):
        b = ordered[i]
        try:
            x0 = max(0, int(float(b.get("x0"))))
            y0 = max(0, int(float(b.get("y0"))))
            x1 = min(w0, int(float(b.get("x1"))))
            y1 = min(h0, int(float(b.get("y1"))))
        except (TypeError, ValueError):
            return []
        if x1 - x0 < 16 or y1 - y0 < 16:
            return []
        crop = im.crop((x0, y0, x1, y1))
        cw, ch = crop.size
        ratio = round(ch / float(cw), 2)
        crop_path = scale
        if out_dir_p is not None:
            crop_path = out_dir_p / f"scale_crop_{role}.png"
            crop.save(crop_path)
        rows.append(
            {
                "role": role,
                "path": str(crop_path),
                "ratio": ratio,
                "contentH": float(ch),
                "source": "scale",
                "bbox": [x0, y0, x1, y1],
            }
        )
        LOG.info(
            "bag aspect scale role=%s h/w=%s contentH=%s bbox=%s",
            role,
            ratio,
            ch,
            [x0, y0, x1, y1],
        )
    return rows


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


def _retry_note(fails: Sequence[Dict[str, Any]]) -> str:
    bits = []
    for row in fails:
        role = str(row.get("role") or "")
        file_r = float(row.get("ratio") or 0)
        drawn = row.get("drawnRatio")
        try:
            drawn_f = float(drawn)
        except (TypeError, ValueError):
            drawn_f = 0.0
        if str(row.get("kind") or "") == "missing":
            bits.append(
                f"{role} が判定できなかった（完成画で袋が見えない／測れない）。"
                "列挙した袋はすべて描き、縦横と高さが測れる大きさにせよ。印字と場面は維持。"
            )
            continue
        if str(row.get("kind") or "") == "height":
            bits.append(
                f"{role} の高さが {row.get('baseRole')} と合っていない。"
                f"ファイルでは {role} は {row.get('baseRole')} の {float(row.get('expect') or 0):.2f} 倍。"
                f"描いた絵では {float(row.get('got') or 0):.2f} 倍。"
                "高さの対比をファイルに戻せ。縁のうねりは残す。印字と場面は維持。"
            )
            continue
        if str(row.get("kind") or "") == "printHue":
            bits.append(
                f"{role} の印刷面の色相がファイルと違う。"
                "黒・無彩色・別色に落とすな。ファイルの色相と彩度に戻せ。"
                "影の向きだけ場面に合わせる。縦横と印字は維持。"
            )
            continue
        if str(row.get("kind") or "") == "printLight":
            bits.append(
                f"{role} の印刷面が正本より明らかに暗い、またはハイライトの色みが消えている。"
                "ベタ面の明度帯とハイライトの青み（または正本の色み）を戻せ。"
                "黒一色に潰すな。影の向きだけ場面に合わせる。縦横と印字は維持。"
            )
            continue
        if drawn_f > file_r:
            how = "縦に伸びている"
        else:
            how = "横に潰れている"
        bits.append(
            f"{role} の包材が{how}。"
            f"ファイルは縦が横の {file_r:.2f} 倍。"
            f"描いた包材は約 {drawn_f:.2f} 倍。"
            "ファイルの比に戻せ。縁のうねりは残す。印字と場面は維持。"
        )
    if not bits:
        return ""
    return "袋の縦横・大きさだけ直せ。他は触るな。\n" + "\n".join(bits)


def check_drawn_bags(scene_path: Path, rows: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    """完成画の袋の面を、ファイルの縦÷横と比べる。読めなければ落とさない。"""
    out: Dict[str, Any] = {
        "ok": True,
        "skipped": False,
        "items": [],
        "retryNote": "",
        "error": "",
    }
    scene = Path(scene_path)
    usable = [r for r in rows if Path(str(r.get("path") or "")).is_file()]
    if not scene.is_file() or not usable:
        out["skipped"] = True
        out["error"] = "no_file"
        return out
    tol = aspect_tolerance()
    lines = [
        "完成画と、各包材ファイルを比べてください。袋・缶・瓶・箱のどれでも同じです。",
        "drawnRatio は包材の面の縦÷横。外接矩形ではなく、上辺と側辺。傾きで外接が縦長でも、面がファイルと同じならその値。",
        "heightShare は、その包材の面の高さが完成画の高さに占める割合（0から1）。外接ではなく面の高さ。",
        "printHueOk は印刷ベタ面の色相がファイルと同じ系統か。"
        "例: ファイルが濃紺なのに完成が黒・グレー・無彩色なら false。ロゴの細部は見ない。",
        "printLightOk は印刷ベタの明度帯。"
        "正本ベタより明らかに暗い、またはハイライトの色み（青み等）が消えて黒一色なら false。"
        "影の向きだけ場面に合わせるのは可。明度を大きく落とすのは不可。",
        "JSONのみ:",
        '{"items":[{"role":"ROLE","drawnRatio":1.45,"heightShare":0.40,"printHueOk":true,"printLightOk":true}]}',
        "下に列挙した role はすべて items に入れる（完成画に見えない袋も、見えないならその旨ではなく role を省略せず、見える袋だけでも列挙漏れは不可。列挙分は必ず返す）。",
        "role は次のとおり。",
    ]
    for row in usable:
        lines.append(f"- {row['role']} ファイル比 {float(row['ratio']):.2f}")
    prompt = "\n".join(lines)
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        out["skipped"] = True
        out["error"] = str(e)
        return out
    parts: List[Dict[str, Any]] = [{"text": prompt}]
    parts.append({"text": "完成画"})
    parts.append(
        {
            "inline_data": {
                "mime_type": mime_for(scene),
                "data": base64.b64encode(scene.read_bytes()).decode("ascii"),
            }
        }
    )
    for row in usable:
        p = Path(str(row["path"]))
        parts.append({"text": str(row["role"])})
        parts.append(
            {
                "inline_data": {
                    "mime_type": mime_for(p),
                    "data": base64.b64encode(p.read_bytes()).decode("ascii"),
                }
            }
        )
    last = ""
    obj = None
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
            LOG.warning("bag aspect fail model=%s: %s", mid, e)
    if obj is None:
        out["skipped"] = True
        out["error"] = last or "parse_failed"
        return out
    by_role = {str(r["role"]): r for r in usable}
    items_in = obj.get("items") if isinstance(obj.get("items"), list) else []
    fails: List[Dict[str, Any]] = []
    seen = 0
    got_roles = set()
    for item in items_in:
        if not isinstance(item, dict):
            continue
        role = str(item.get("role") or "")
        src = by_role.get(role)
        if src is None:
            continue
        try:
            drawn = float(item.get("drawnRatio"))
        except (TypeError, ValueError):
            continue
        if drawn <= 0.2 or drawn >= 5:
            continue
        seen += 1
        got_roles.add(role)
        file_r = float(src["ratio"])
        rel = abs(drawn / file_r - 1.0)
        ok = rel <= tol
        try:
            share = float(item.get("heightShare"))
        except (TypeError, ValueError):
            share = 0.0
        if share < 0.05 or share > 0.95:
            share = 0.0
        hue_raw = item.get("printHueOk")
        if isinstance(hue_raw, bool):
            hue_ok = hue_raw
        elif hue_raw is None:
            hue_ok = True  # 古い応答互換。欠落は落とさない
        else:
            hue_ok = str(hue_raw).strip().lower() in ("1", "true", "yes", "ok")
        light_raw = item.get("printLightOk")
        if isinstance(light_raw, bool):
            light_ok = light_raw
        elif light_raw is None:
            light_ok = True  # 欠落は落とさない（モデル未対応時）
        else:
            light_ok = str(light_raw).strip().lower() in ("1", "true", "yes", "ok")
        rec = {
            "role": role,
            "ratio": file_r,
            "drawnRatio": round(drawn, 2),
            "rel": round(rel, 3),
            "heightShare": round(share, 3) if share else None,
            "printHueOk": hue_ok,
            "printLightOk": light_ok,
            "ok": ok and hue_ok and light_ok,
        }
        out["items"].append(rec)
        if not ok:
            fails.append(rec)
        elif not hue_ok:
            fails.append({"role": role, "kind": "printHue", "ok": False})
        elif not light_ok:
            fails.append({"role": role, "kind": "printLight", "ok": False})
    # 列挙した袋はすべて測る（片方だけだと対比が走らない）
    for row in usable:
        role = str(row["role"])
        if role in got_roles:
            continue
        miss = {"role": role, "kind": "missing", "ok": False}
        out["items"].append(miss)
        fails.append(miss)
    out["items"].sort(key=lambda r: str((r or {}).get("role") or ""))
    if len([x for x in out["items"] if x.get("heightShare")]) >= 2:
        # PART_1 があれば基準。無ければ先頭の測れた袋。
        measured = [x for x in out["items"] if x.get("heightShare") and "kind" not in x]
        base = None
        for cand in measured:
            if str(cand.get("role") or "").endswith("_1") or str(cand.get("role")) == "IMAGE_PACKAGE_TRUTH":
                base = cand
                break
        if base is None and measured:
            base = measured[0]
        if base is not None:
            base_src = by_role.get(str(base.get("role")))
            base_h = float((base_src or {}).get("contentH") or 0)
            base_s = float(base.get("heightShare") or 0)
            for other in measured:
                if other is base:
                    continue
                src = by_role.get(str(other.get("role")))
                oh = float((src or {}).get("contentH") or 0)
                os_ = float(other.get("heightShare") or 0)
                if base_h < 8 or oh < 8 or base_s <= 0 or os_ <= 0:
                    continue
                expect = oh / base_h
                got = os_ / base_s
                hrel = abs(got / expect - 1.0)
                other["heightRel"] = round(hrel, 3)
                other["heightExpect"] = round(expect, 2)
                other["heightGot"] = round(got, 2)
                other["heightBaseRole"] = base.get("role")
                if hrel > tol:
                    other["ok"] = False
                    fails.append(
                        {
                            "role": other.get("role"),
                            "kind": "height",
                            "baseRole": base.get("role"),
                            "expect": round(expect, 2),
                            "got": round(got, 2),
                        }
                    )
    if seen == 0:
        out["skipped"] = True
        out["error"] = "no_items"
        return out
    # 2袋以上あるのに対比できなければ不合格（対比結果を必ず返す）
    need_pair = len(usable) >= 2
    paired = any(
        isinstance(x, dict) and x.get("heightRel") is not None for x in out["items"]
    )
    if need_pair and not paired and not any(str(f.get("kind")) == "missing" for f in fails):
        fails.append(
            {
                "role": "PAIR",
                "kind": "height",
                "baseRole": str((out["items"][0] or {}).get("role") or ""),
                "expect": 0,
                "got": 0,
                "note": "heightShare不足で対比不能",
            }
        )
        out["pairSkipped"] = True
    out["ok"] = not fails
    out["retryNote"] = _retry_note(fails)
    LOG.info(
        "bag aspect check ok=%s n=%s fail=%s pair=%s",
        out["ok"],
        seen,
        ",".join(str(f.get("role")) for f in fails),
        paired,
    )
    return out


def format_bag_verdict_text(aspect: Dict[str, Any], *, theme_id: str = "", try_name: str = "") -> str:
    """人が読める袋判定（縦横＋対比）。"""
    lines: List[str] = []
    head = "袋判定"
    if theme_id:
        head += f" {theme_id}"
    if try_name:
        head += f" ({try_name})"
    lines.append(head)
    if not isinstance(aspect, dict):
        lines.append("結果: 不明")
        return "\n".join(lines) + "\n"
    if aspect.get("skipped"):
        lines.append(f"結果: スキップ（{aspect.get('error') or 'n/a'}）")
        return "\n".join(lines) + "\n"
    lines.append("結果: " + ("OK" if aspect.get("ok") else "NG"))
    for item in aspect.get("items") or []:
        if not isinstance(item, dict):
            continue
        role = str(item.get("role") or "")
        if str(item.get("kind") or "") == "missing":
            lines.append(f"- {role}: 測れず（欠落） NG")
            continue
        rel_v = item.get("rel")
        bit = (
            f"- {role}: 縦横 ファイル{item.get('ratio')} / 描画{item.get('drawnRatio')}"
            f" 差{rel_v}"
        )
        if item.get("heightRel") is not None:
            h_ok = float(item.get("heightRel") or 0) <= aspect_tolerance()
            bit += (
                f" → 縦横{'OK' if (rel_v is not None and float(rel_v) <= aspect_tolerance()) else 'NG'}"
                f" ／対比 vs {item.get('heightBaseRole') or '?'}"
                f" 期待{item.get('heightExpect')} 実測{item.get('heightGot')}"
                f" 差{item.get('heightRel')} → {'OK' if h_ok else 'NG'}"
            )
        else:
            bit += " → " + ("OK" if item.get("ok") else "NG")
        lines.append(bit)
    note = str(aspect.get("retryNote") or "").strip()
    if note:
        lines.append("やり直し指示:")
        lines.append(note)
    return "\n".join(lines) + "\n"


def write_bag_verdict_bundle(
    path: Path,
    rows: Sequence[Dict[str, Any]],
) -> Path:
    """比較フォルダ等へ袋判定まとめを書く。rows: themeId/try/aspect/ok。"""
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    blocks = ["【袋判定まとめ】縦横比＋袋どうしの大きさ対比", ""]
    for row in rows:
        asp = row.get("aspect") if isinstance(row, dict) else None
        blocks.append(
            format_bag_verdict_text(
                asp if isinstance(asp, dict) else {},
                theme_id=str((row or {}).get("themeId") or ""),
                try_name=str((row or {}).get("tryName") or ""),
            ).rstrip()
        )
        blocks.append("")
    p.write_text("\n".join(blocks).rstrip() + "\n", encoding="utf-8")
    return p


def aspect_worst_rel(aspect: Dict[str, Any]) -> Optional[float]:
    """測れた面と高さの差のうち、大きい方。読めなければ None。"""
    if not isinstance(aspect, dict) or aspect.get("skipped"):
        return None
    worst: Optional[float] = None
    for item in aspect.get("items") or []:
        if not isinstance(item, dict):
            continue
        for key in ("rel", "heightRel"):
            raw = item.get(key)
            if raw is None or raw == "":
                continue
            try:
                val = float(raw)
            except (TypeError, ValueError):
                continue
            if worst is None or val > worst:
                worst = val
    return worst


def check_face_vs_tableware(
    scene_path: Path, *, height_cm: Optional[float], width_cm: Optional[float]
) -> Dict[str, Any]:
    """面の縦cmがあるとき、包材と同じ面の器だけと比べる。手前の器や、同じ面に器が無ければ落とさない。"""
    out: Dict[str, Any] = {
        "ok": True,
        "skipped": False,
        "retryNote": "",
        "error": "",
    }
    try:
        h = float(height_cm or 0)
        w = float(width_cm or 0)
    except (TypeError, ValueError):
        h, w = 0.0, 0.0
    scene = Path(scene_path)
    if h <= 0 or w <= 0 or not scene.is_file():
        out["skipped"] = True
        out["error"] = "no_face_cm"
        return out
    tol = aspect_tolerance()
    prompt = (
        "完成画です。次を測ってください。無いものは 0。\n"
        "packageHeightShare: 包材の面の高さが、画像の高さに占める割合。\n"
        "samePlaneBowlShare: 包材とカメラから同じ距離にあるご飯茶碗の口径が、画像の高さに占める割合。手前の茶碗、または無いときは 0。\n"
        "samePlanePlateShare: 包材とカメラから同じ距離にある小皿の口径が、画像の高さに占める割合。手前の皿、または無いときは 0。\n"
        "JSONのみ: "
        '{"packageHeightShare":0.4,"samePlaneBowlShare":0.2,"samePlanePlateShare":0}'
    )
    try:
        client = make_client(load_api_key())
    except SystemExit as e:
        out["skipped"] = True
        out["error"] = str(e)
        return out
    parts = [
        {"text": prompt},
        {"text": "完成画"},
        {
            "inline_data": {
                "mime_type": mime_for(scene),
                "data": base64.b64encode(scene.read_bytes()).decode("ascii"),
            }
        },
    ]
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
            LOG.warning("face scale fail model=%s: %s", mid, e)
    if obj is None:
        out["skipped"] = True
        out["error"] = last or "parse_failed"
        return out

    def _share(key: str) -> float:
        try:
            v = float(obj.get(key) or 0)
        except (TypeError, ValueError):
            return 0.0
        if v < 0.05 or v > 0.95:
            return 0.0
        return v

    pkg = _share("packageHeightShare")
    bowl = _share("samePlaneBowlShare")
    plate = _share("samePlanePlateShare")
    out["packageHeightShare"] = round(pkg, 3) if pkg else None
    out["samePlaneBowlShare"] = round(bowl, 3) if bowl else None
    out["samePlanePlateShare"] = round(plate, 3) if plate else None
    if pkg <= 0 or (bowl <= 0 and plate <= 0):
        out["skipped"] = True
        out["error"] = "no_same_plane"
        return out
    if bowl > 0:
        expect = h / _BOWL_CM
        got = pkg / bowl
        label = "ご飯茶碗"
    else:
        expect = h / _PLATE_CM
        got = pkg / plate
        label = "小皿"
    rel = abs(got / expect - 1.0)
    out["expect"] = round(expect, 2)
    out["got"] = round(got, 2)
    out["ref"] = label
    fails = []
    if rel > tol:
        fails.append(
            f"包材の高さは、同じ面の{label}の口径の {expect:.2f} 倍。"
            f"描いた絵は {got:.2f} 倍。同じ面の器との比に戻せ。手前の皿は動かすな。縁のうねりと印字は維持。"
        )
    out["ok"] = not fails
    out["retryNote"] = "\n".join(fails)
    LOG.info("face scale ok=%s ref=%s got=%s", out["ok"], label, out.get("got"))
    return out
