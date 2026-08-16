# -*- coding: utf-8 -*-
"""出品 00_設定マスタ の FBA first-fit／自己発送 first-fit（①用。コード.js と同判定）。"""
from __future__ import annotations

import math
import re
from typing import Any


def _num(v: Any) -> float:
    if v is None or v == "":
        return float("nan")
    try:
        return float(str(v).replace(",", "").replace("円", "").strip())
    except (TypeError, ValueError):
        return float("nan")


def parse_fba_remark(remark: str) -> dict:
    out = {
        "mode": "none",
        "maxL": float("nan"),
        "maxW": float("nan"),
        "maxH": float("nan"),
        "maxSum": float("nan"),
        "maxWeightG": float("nan"),
    }
    s = re.sub(r"\s+", "", str(remark or "").replace(",", ""))
    if not s:
        return out
    w = re.search(r"/\s*([0-9]+(?:\.[0-9]+)?)\s*(kg|g)", s, re.I) or re.search(
        r"([0-9]+(?:\.[0-9]+)?)\s*(kg|g)\s*以下", s, re.I
    ) or re.search(r"([0-9]+(?:\.[0-9]+)?)\s*(kg|g)", s, re.I)
    if w:
        wv = float(w.group(1))
        wu = w.group(2).lower()
        out["maxWeightG"] = wv * 1000 if wu == "kg" else wv
    box = re.search(
        r"([0-9]+(?:\.[0-9]+)?)\s*[x×]\s*([0-9]+(?:\.[0-9]+)?)\s*[x×]\s*([0-9]+(?:\.[0-9]+)?)\s*cm",
        s,
        re.I,
    )
    if box:
        dims = sorted([float(box.group(1)), float(box.group(2)), float(box.group(3))], reverse=True)
        out["mode"] = "box"
        out["maxL"], out["maxW"], out["maxH"] = dims
        return out
    sm = re.search(r"(?:三辺合計|合計)?\s*([0-9]+(?:\.[0-9]+)?)\s*cm", s, re.I)
    if sm:
        out["mode"] = "sum"
        out["maxSum"] = float(sm.group(1))
        return out
    return out


def parse_fba_table(rows: list[list[Any]]) -> list[dict]:
    out = []
    ord_ = 0
    for r in rows:
        a = str(r[0] if r else "").strip()
        if a != "FBA手数料":
            continue
        name = str(r[1] if len(r) > 1 else "").strip()
        fee = _num(r[3] if len(r) > 3 else "")
        remark = str(r[5] if len(r) > 5 else "").strip()
        if not name or not math.isfinite(fee) or fee < 0:
            continue
        p = parse_fba_remark(remark)
        p.update({"name": name, "fee": fee, "remark": remark, "rowOrder": ord_})
        ord_ += 1
        out.append(p)
    return out


def pick_fba_tier(table: list[dict], l_cm: float, w_cm: float, h_cm: float, weight_g: float) -> dict:
    sides = [x for x in (l_cm, w_cm, h_cm) if math.isfinite(x) and x > 0]
    if len(sides) < 3:
        return {"tier": "", "fee": "", "reason": "dims_incomplete", "source": "none", "sumCm": ""}
    sides.sort(reverse=True)
    longest, median, shortest = sides
    sum_cm = longest + median + shortest
    g = weight_g if math.isfinite(weight_g) else float("nan")
    in_std = longest <= 45 and median <= 35 and shortest <= 20 and (not math.isfinite(g) or g <= 9000)
    if not table:
        return {"tier": "", "fee": "", "reason": "settings_empty", "source": "none", "sumCm": str(round(sum_cm, 1))}
    for t in table:
        if t["mode"] == "none":
            continue
        is_large = bool(re.match(r"^大型|^特大", t["name"]))
        is_small = bool(re.match(r"^小型|^標準", t["name"]))
        if is_small and not in_std:
            continue
        if is_large and in_std:
            continue
        if math.isfinite(t["maxWeightG"]) and math.isfinite(g) and g > t["maxWeightG"]:
            continue
        if math.isfinite(t["maxWeightG"]) and not math.isfinite(g):
            continue
        if t["mode"] == "box":
            if longest <= t["maxL"] and median <= t["maxW"] and shortest <= t["maxH"]:
                return {
                    "tier": t["name"],
                    "fee": str(int(t["fee"]) if t["fee"] == int(t["fee"]) else t["fee"]),
                    "reason": "settings_box " + t["remark"],
                    "source": "settings",
                    "sumCm": str(round(sum_cm, 1)),
                }
            continue
        if t["mode"] == "sum" and math.isfinite(t["maxSum"]) and sum_cm <= t["maxSum"]:
            return {
                "tier": t["name"],
                "fee": str(int(t["fee"]) if t["fee"] == int(t["fee"]) else t["fee"]),
                "reason": "settings_sum " + t["remark"],
                "source": "settings",
                "sumCm": str(round(sum_cm, 1)),
            }
    return {"tier": "", "fee": "", "reason": "no_settings_fit", "source": "none", "sumCm": str(round(sum_cm, 1))}


def is_compact_name(size: str) -> bool:
    return bool(re.search(r"コンパクト|compact", str(size or ""), re.I))


def infer_edge_max(size: str) -> float:
    s = str(size or "")
    if re.search(r"ネコ|Nekopos|nekopos|ねこぽ", s, re.I):
        return 60.0
    m = re.search(r"(\d{2,3})", s)
    return float(m.group(1)) if m else float("nan")


def parse_ship_table(rows: list[list[Any]]) -> list[dict]:
    compact_a, compact_b = 50.0, 58.8
    out = []
    ord_ = 0
    for r in rows:
        a = str(r[0] if r else "").strip()
        if a not in ("自己発送", "自己配送"):
            continue
        b = str(r[1] if len(r) > 1 else "").strip()
        value1 = _num(r[2] if len(r) > 2 else "")
        price = _num(r[3] if len(r) > 3 else "")
        if not b or not math.isfinite(price) or price < 0:
            continue
        edge_kind = 0
        edge_max = float("nan")
        edge_alt = float("nan")
        if is_compact_name(b):
            edge_kind = 2
            edge_max = compact_a
            edge_alt = compact_b
        elif math.isfinite(value1) and 10 <= value1 <= 500:
            edge_kind = 1
            edge_max = value1
        else:
            inf = infer_edge_max(b)
            if math.isfinite(inf):
                edge_kind = 1
                edge_max = inf
        out.append(
            {
                "size": b,
                "price": price,
                "value1": value1,
                "edgeKind": edge_kind,
                "edgeMax": edge_max,
                "edgeAlt": edge_alt,
                "_rowOrder": ord_,
            }
        )
        ord_ += 1
    out.sort(
        key=lambda x: (
            x["value1"] if math.isfinite(x["value1"]) else float("inf"),
            x["price"],
            x["_rowOrder"],
        )
    )
    for i, x in enumerate(out):
        x["rankIndex"] = i
    return out


def ship_accepts(row: dict, packed: float) -> bool:
    if not math.isfinite(packed) or packed <= 0:
        return True
    if row["edgeKind"] == 2:
        a, b = row["edgeMax"], row["edgeAlt"]
        return (math.isfinite(a) and packed <= a) or (math.isfinite(b) and packed <= b)
    if row["edgeKind"] == 1:
        return math.isfinite(row["edgeMax"]) and packed <= row["edgeMax"]
    return True


def pick_self_ship(table: list[dict], l_cm: float, w_cm: float, h_cm: float) -> dict:
    sides = [x for x in (l_cm, w_cm, h_cm) if math.isfinite(x) and x > 0]
    if len(sides) < 3 or not table:
        return {"size": "", "price": "", "reason": "no_dims_or_table", "sumCm": ""}
    packed = sum(sides)
    for row in table:
        if ship_accepts(row, packed):
            p = row["price"]
            return {
                "size": row["size"],
                "price": str(int(p) if p == int(p) else p),
                "reason": "first_fit",
                "sumCm": str(round(packed, 1)),
            }
    return {"size": "", "price": "", "reason": "no_fit", "sumCm": str(round(packed, 1))}


def keepa_pack_cm_g(product: dict) -> tuple[float, float, float, float]:
    """Keepa package* は mm / g。"""

    def mm_to_cm(v: Any) -> float:
        n = _num(v)
        if not math.isfinite(n) or n <= 0:
            return float("nan")
        return n / 10.0

    l = mm_to_cm(product.get("packageLength"))
    w = mm_to_cm(product.get("packageWidth"))
    h = mm_to_cm(product.get("packageHeight"))
    g = _num(product.get("packageWeight"))
    return l, w, h, g
