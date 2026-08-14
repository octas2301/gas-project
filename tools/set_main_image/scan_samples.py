# -*- coding: utf-8 -*-
"""見本スキャン改: 右上ROIで金丸＋内部の暗色文字ボックスを推定。"""
from __future__ import annotations

import json
import statistics
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, ImageFilter

ROOT = Path(r"G:/マイドライブ/05.画像生成（セットMAIN）")
OUT = Path(__file__).resolve().parent / "sample_scan_report.json"
RULES_OUT = Path(__file__).resolve().parent / "layout_rules.json"


def _is_gold(r: int, g: int, b: int) -> bool:
    if r < 170 or g < 110:
        return False
    if b > 170:
        return False
    if r + g < 300:
        return False
    if abs(r - g) > 100:
        return False
    if min(r, g) > 240 and b > 200:
        return False
    return True


def _is_ink(r: int, g: int, b: int) -> bool:
    """金丸内の数字・単位（赤〜黒）。"""
    # dark
    if r + g + b < 220:
        return True
    # red digit
    if r > 100 and r > g * 1.4 and r > b * 1.4 and g < 120 and b < 120:
        return True
    return False


def find_gold_and_ink(im: Image.Image) -> Optional[Dict[str, Any]]:
    rgb = im.convert("RGB")
    w, h = rgb.size
    # 右上ROI
    x0s, y0s = int(w * 0.62), 0
    x1s, y1s = w, int(h * 0.42)
    px = rgb.load()

    gx, gy = [], []
    for y in range(y0s, y1s):
        for x in range(x0s, x1s):
            if _is_gold(*px[x, y]):
                gx.append(x)
                gy.append(y)
    if len(gx) < 200:
        return None
    gbox = (min(gx), min(gy), max(gx), max(gy))
    # 金丸が巨大すぎる場合は中央付近で円近似
    gw, gh = gbox[2] - gbox[0], gbox[3] - gbox[1]
    if gw > w * 0.45 or gh > h * 0.45:
        # 右上コーナーにクリップ
        gbox = (
            max(gbox[0], int(w * 0.70)),
            max(gbox[1], 0),
            min(gbox[2], w - 1),
            min(gbox[3], int(h * 0.32)),
        )
        gw, gh = gbox[2] - gbox[0], gbox[3] - gbox[1]

    ix, iy = [], []
    # 金丸内を少し内側
    pad = int(min(gw, gh) * 0.12)
    for y in range(gbox[1] + pad, gbox[3] - pad):
        for x in range(gbox[0] + pad, gbox[2] - pad):
            if _is_ink(*px[x, y]):
                ix.append(x)
                iy.append(y)
    if len(ix) < 30:
        # fallback: inner 55% of gold as digit area
        dig = {
            "x": gbox[0] + int(gw * 0.18),
            "y": gbox[1] + int(gh * 0.12),
            "w": int(gw * 0.55),
            "h": int(gh * 0.55),
        }
        return {"goldBBox": list(gbox), "digitBox": dig, "inkPixels": 0, "digitLenGuess": None}

    ib = (min(ix), min(iy), max(ix), max(iy))
    iw, ih = ib[2] - ib[0], ib[3] - ib[1]
    # 幅で1桁/2桁推定（セット文字を含むので少し広め）
    aspect = iw / max(ih, 1)
    # 数字本体は上部寄りのことが多い → 上60%で再計測
    top_cut = ib[1] + int(ih * 0.62)
    ix2 = [x for x, y in zip(ix, iy) if y <= top_cut]
    iy2 = [y for y in iy if y <= top_cut]
    if len(ix2) >= 20:
        ib2 = (min(ix2), min(iy2), max(ix2), max(iy2))
        iw2, ih2 = ib2[2] - ib2[0], ib2[3] - ib2[1]
        aspect2 = iw2 / max(ih2, 1)
        dig = {"x": ib2[0], "y": ib2[1], "w": iw2, "h": ih2}
        # 2桁は横に広がる
        digit_len = 2 if aspect2 >= 1.15 or iw2 >= gh * 0.42 else 1
    else:
        dig = {"x": ib[0], "y": ib[1], "w": iw, "h": int(ih * 0.55)}
        digit_len = 2 if aspect >= 1.2 else 1

    return {
        "goldBBox": list(gbox),
        "digitBox": dig,
        "inkPixels": len(ix),
        "digitLenGuess": digit_len,
    }


def scan_rakuten(folder: Path) -> Dict[str, Any]:
    rows = []
    for p in sorted(folder.iterdir(), key=lambda x: x.name):
        if p.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
            continue
        im = Image.open(p)
        hit = find_gold_and_ink(im)
        row: Dict[str, Any] = {"file": p.name, "size": list(im.size)}
        if hit:
            row.update(hit)
        rows.append(row)

    boxes1 = [r["digitBox"] for r in rows if r.get("digitLenGuess") == 1 and r.get("digitBox")]
    boxes2 = [r["digitBox"] for r in rows if r.get("digitLenGuess") == 2 and r.get("digitBox")]
    golds = [r["goldBBox"] for r in rows if r.get("goldBBox")]

    def med_box(lst: List[Dict[str, int]]) -> Optional[Dict[str, int]]:
        if not lst:
            return None
        return {
            "x": int(statistics.median([b["x"] for b in lst])),
            "y": int(statistics.median([b["y"] for b in lst])),
            "w": int(statistics.median([b["w"] for b in lst])),
            "h": int(statistics.median([b["h"] for b in lst])),
            "nSamples": len(lst),
        }

    # gold circle center median
    gold_rule = None
    if golds:
        gold_rule = {
            "x0": int(statistics.median([g[0] for g in golds])),
            "y0": int(statistics.median([g[1] for g in golds])),
            "x1": int(statistics.median([g[2] for g in golds])),
            "y1": int(statistics.median([g[3] for g in golds])),
            "nSamples": len(golds),
        }

    return {
        "count": len(rows),
        "withGold": len(golds),
        "goldBBoxMedian": gold_rule,
        "digitBox1": med_box(boxes1),
        "digitBox2": med_box(boxes2),
        "items": rows,
    }


def scan_amazon_layouts(folder: Path) -> Dict[str, Any]:
    items = []
    for p in sorted(folder.iterdir(), key=lambda x: x.name):
        if p.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
            continue
        im = Image.open(p).convert("RGB")
        # downscale
        small = im.resize((240, 240))
        px = small.load()
        xs, ys = [], []
        for y in range(240):
            for x in range(240):
                r, g, b = px[x, y]
                if r + g + b < 700:
                    xs.append(x)
                    ys.append(y)
        if len(xs) < 50:
            continue
        # cluster left vs right mass
        left = sum(1 for x in xs if x < 120)
        right = len(xs) - left
        top = sum(1 for y in ys if y < 120)
        bot = len(xs) - top
        items.append(
            {
                "file": p.name,
                "cx": round(statistics.mean(xs) / 240, 3),
                "cy": round(statistics.mean(ys) / 240, 3),
                "leftShare": round(left / len(xs), 3),
                "rightShare": round(right / len(xs), 3),
                "topShare": round(top / len(xs), 3),
                "bottomShare": round(bot / len(xs), 3),
                "patternHint": (
                    "hero_right_stack_left"
                    if left > right * 1.05 and bot > top * 0.9
                    else (
                        "hero_left_stack_right"
                        if right > left * 1.05
                        else "centered_cluster"
                    )
                ),
            }
        )
    patterns: Dict[str, int] = {}
    for it in items:
        patterns[it["patternHint"]] = patterns.get(it["patternHint"], 0) + 1
    return {"count": len(items), "patternCounts": patterns, "items": items}


def build_rules(rakuten: Dict[str, Any], amazon: Dict[str, Any]) -> Dict[str, Any]:
    # Prefer measured boxes; fallback sensible defaults for 1200 canvas
    d1 = rakuten.get("digitBox1") or {"x": 980, "y": 55, "w": 95, "h": 120}
    d2 = rakuten.get("digitBox2") or {"x": 955, "y": 60, "w": 130, "h": 110}
    # 2桁は少しだけ小さく（高さ比）
    if d1.get("h") and d2.get("h"):
        # ensure 2-digit height slightly smaller than 1-digit if samples say otherwise flip
        pass

    gold = rakuten.get("goldBBoxMedian") or {
        "x0": 900,
        "y0": 20,
        "x1": 1180,
        "y1": 300,
    }

    return {
        "canvas": 1200,
        "amazon": {
            "background": "transparent_or_pure_white",
            "preprocess": "remove_near_white_bg_to_alpha_before_compose",
            "layoutFromSamples": {
                "dominantPatterns": amazon.get("patternCounts"),
                "rules": [
                    "白（または透過）背景のみ。影は最小。",
                    "見本の大小比・重なり・スタック方向を転写する（グリッド禁止）。",
                    "よくある型: 左側に数量スタック、右下に大きめヒーロー（開封・実物寄り）。",
                    "商品は01素材のみ。見本の商品デザインはコピーしない。",
                    "食品は Octas をヒーロー右下付近に小さく。",
                    "特別セット（箸上げ・特殊ポージング）は都度応相談。",
                ],
            },
            "gemini": {
                "recommended": True,
                "note": "PoCで品質改善を確認。プロンプトで LAYOUT_BLUEPRINT 拘束。",
            },
            "openai": {
                "recommended": False,
                "note": "ラベル文字崩れがあり本線候補外（比較用残置可）。",
            },
        },
        "rakuten": {
            "mode": "layer_only_base_locked",
            "goldCircleApprox1200": gold,
            "digitBoxByDigitLen1200": {
                "1": {k: d1[k] for k in ("x", "y", "w", "h") if k in d1},
                "2": {k: d2[k] for k in ("x", "y", "w", "h") if k in d2},
            },
            "sizeRule": {
                "1digit": "金丸内でやや大きく。高さは digitBox.h の約 0.85〜0.92",
                "2digit": "1桁よりわずかに小さく。高さは digitBox.h の約 0.72〜0.82（横幅優先で収める）",
            },
            "textStructure": {
                "primary": "{set_count}{unit}",
                "secondaryFixed": "セット",
                "secondaryNote": "見本では数字＋単位の下に『セット』が固定。ベースに既に含まれる場合は重ねない。",
                "specialSets": "例外レイアウトは都度応相談",
            },
            "fonts": {
                "canvaPreferred": ["筑紫明朝H", "源柔ゴシック"],
                "localFontIds": {
                    "badge_number_mincho": "tsukushi_mincho_like",
                    "badge_unit_mincho": "tsukushi_mincho_like",
                    "headline_gothic": "genjyuu_gothic_like",
                },
                "geminiApi": {
                    "note": "画像生成APIは任意TTFを埋め込めない。レイヤー合成はローカルFontを使用。",
                    "if_generating_digit_only": "serif heavy / Mincho-like Japanese numeral",
                },
                "openaiApi": {
                    "note": "images.edit も同様にAPIフォント指定不可。レイヤー合成を本線とする。",
                    "if_generating_digit_only": "Tsukushi Mincho / heavy serif Japanese digits",
                },
            },
            "fillColorDefault": [120, 20, 20, 255],
        },
        "scanMeta": {
            "rakutenSamples": rakuten.get("count"),
            "rakutenWithGold": rakuten.get("withGold"),
            "amazonSamples": amazon.get("count"),
        },
    }


def main() -> None:
    rakuten = scan_rakuten(ROOT / "04.楽天見本")
    amazon = scan_amazon_layouts(ROOT / "03.amazon見本")
    report = {"rakutenRefs": rakuten, "amazonRefs": amazon}
    OUT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    rules = build_rules(rakuten, amazon)
    # 人手補正: 見本目視（まぐろチーズ系）の金丸は右上小さめ。検出が大きく外れた場合の安全側
    g = rules["rakuten"]["goldCircleApprox1200"]
    if g and (g.get("x1", 0) - g.get("x0", 0)) > 400:
        rules["rakuten"]["goldCircleApprox1200"] = {
            "x0": 920,
            "y0": 25,
            "x1": 1175,
            "y1": 280,
            "nSamples": g.get("nSamples"),
            "note": "auto-median too wide; clamped to typical top-right badge",
        }
        rules["rakuten"]["digitBoxByDigitLen1200"] = {
            "1": {"x": 955, "y": 55, "w": 100, "h": 125},
            "2": {"x": 940, "y": 62, "w": 145, "h": 108},
        }
        rules["rakuten"]["digitBoxSource"] = "clamped_from_scan_plus_eyeball"
    else:
        rules["rakuten"]["digitBoxSource"] = "scan_median"

    RULES_OUT.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
    print("gold", rules["rakuten"]["goldCircleApprox1200"])
    print("digit1", rules["rakuten"]["digitBoxByDigitLen1200"]["1"])
    print("digit2", rules["rakuten"]["digitBoxByDigitLen1200"]["2"])
    print("amazon patterns", amazon.get("patternCounts"))
    print("wrote", OUT.name, RULES_OUT.name)


if __name__ == "__main__":
    main()
