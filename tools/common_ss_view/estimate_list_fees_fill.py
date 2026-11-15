# -*- coding: utf-8 -*-
"""品番リストの空手数料を埋める（推測なし）。

正: FLOW §5.3.2
- 料率は出品 00_設定マスタ
- カテゴリ空 → 販手を書かない
- 寸法空 → 保管を書かない（Keepa梱包があれば事実として使う）
- 既存の数値は上書きしない

使い方:
  python tools/common_ss_view/estimate_list_fees_fill.py --dry
  python tools/common_ss_view/estimate_list_fees_fill.py
  python tools/common_ss_view/estimate_list_fees_fill.py --asins B005FOBQOA,B0DV45VFHW
"""
from __future__ import annotations

import argparse
import math
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

JST = timezone(timedelta(hours=9))

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "competitor_store"))
from client import sheets_service
from listing_fees import parse_fba_table, pick_fba_tier
from schema import MASTER_SS_ID

RES = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
COMP = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
LIST = "仕入検討品番リスト"
SET_TAB = "00_設定マスタ"


def col_letter(n1: int) -> str:
    n = n1
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def num(v):
    try:
        if v is None or v == "":
            return None
        x = float(str(v).replace(",", "").replace("円", "").replace("cm", "").strip())
        return x if x == x else None
    except Exception:
        return None


def parse_size_cm(text: str) -> tuple[float, float, float] | None:
    s = str(text or "").replace("×", "x").replace("＊", "x")
    m = re.search(
        r"(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)",
        s,
    )
    if not m:
        return None
    dims = sorted([float(m.group(1)), float(m.group(2)), float(m.group(3))], reverse=True)
    if any(d <= 0 for d in dims):
        return None
    return dims[0], dims[1], dims[2]


def parse_weight_g(text: str) -> float | None:
    s = str(text or "").strip().lower().replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*kg", s)
    if m:
        return float(m.group(1)) * 1000.0
    m = re.search(r"(\d+(?:\.\d+)?)\s*g", s)
    if m:
        return float(m.group(1))
    v = num(s)
    return v


def load_settings(svc) -> dict:
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=MASTER_SS_ID, range=f"'{SET_TAB}'!A1:F120")
        .execute()
        .get("values")
        or []
    )
    sales = []  # {family, rate, lim}
    storage = {}  # key -> rate
    for r in vals:
        while len(r) < 6:
            r.append("")
        kind = str(r[0]).strip()
        name = str(r[1]).strip()
        rate = num(r[3])
        lim = num(r[4])
        if kind == "販売手数料" and rate is not None:
            fam = name.split(":")[0].strip() if ":" in name else name
            sales.append({"name": name, "family": fam, "rate": rate, "lim": lim})
        if kind == "在庫保管手数料" and rate is not None:
            storage[name] = rate
    fba_tbl = parse_fba_table(vals)
    return {"sales": sales, "storage": storage, "fba": fba_tbl}


def pick_sales_rate(sales: list[dict], category: str, sell: float) -> float | None:
    cat = str(category or "").strip()
    if not cat:
        return None
    # 家族名の優先（長い一致）
    families = []
    for s in sales:
        fam = s["family"]
        if fam in cat or (fam == "食品" and "食品" in cat):
            families.append(s)
    if not families:
        # 食品リサーチ本線: 食品・飲料・お酒
        if "食品" in cat or "飲料" in cat or "お酒" in cat:
            families = [s for s in sales if s["family"] == "食品"]
    if not families:
        return None
    # 帯: lim ありを価格で選ぶ。lim無しは単一料率
    banded = [s for s in families if s["lim"] is not None]
    flat = [s for s in families if s["lim"] is None]
    if banded:
        # 上限昇順で「以下」に入る最初、無ければ最大帯
        ordered = sorted(banded, key=lambda x: x["lim"])
        for s in ordered:
            # 1501円以上は lim=1501 で「以上」扱い
            if "以上" in s["name"]:
                if sell >= s["lim"]:
                    return s["rate"]
            else:
                if sell <= s["lim"]:
                    return s["rate"]
        return ordered[-1]["rate"]
    if flat:
        return flat[0]["rate"]
    return None


def is_standard_pack(l: float, w: float, h: float, g: float | None) -> bool:
    sides = sorted([l, w, h], reverse=True)
    ok = sides[0] <= 45 and sides[1] <= 35 and sides[2] <= 20
    if g is not None and g > 9000:
        return False
    return ok


def storage_fee_yen(
    storage: dict, l: float, w: float, h: float, g: float | None, month: int
) -> int | None:
    vol = l * w * h
    if vol <= 0:
        return None
    season = "1-9月" if 1 <= month <= 9 else "10-12月"
    std = is_standard_pack(l, w, h, g)
    key = f"{'標準' if std else '大型'}_{season}"
    rate = storage.get(key)
    if rate is None:
        return None
    return int(round(vol / 1000.0 * rate))


def referral_fee_yen(rate: float, sell: float) -> int:
    return int(round(sell * rate))


def load_keepa_dims(svc) -> dict[str, tuple[float, float, float, float | None]]:
    """ASIN → (L,W,H,g) from Keepaフル index+cols."""
    kh = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMP, range="'Keepaフル'!1:1")
        .execute()
        .get("values")
        or [[]]
    )[0]
    kidx = {str(h).strip(): i for i, h in enumerate(kh)}
    need = ["ASIN", "梱包_L_cm", "梱包_W_cm", "梱包_H_cm", "梱包_重量_g"]
    for h in need:
        if h not in kidx:
            print("Keepa missing", h)
            return {}
    amap = {}
    for r in (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMP, range="'Keepaフル索引'!A:B")
        .execute()
        .get("values")
        or []
    )[1:]:
        if len(r) >= 2:
            try:
                amap[str(r[0]).strip().upper()] = int(float(r[1]))
            except Exception:
                pass
    # too many to pull all; return loader function via cache on demand
    return {"_amap": amap, "_kidx": kidx, "_svc": svc}  # type: ignore


def keepa_dims_one(cache: dict, asin: str) -> tuple[float, float, float, float | None] | None:
    amap = cache.get("_amap") or {}
    kidx = cache.get("_kidx") or {}
    svc = cache.get("_svc")
    grow = amap.get(asin)
    if not grow or not svc:
        return None
    letter = col_letter(max(kidx.values()) + 1)
    vals = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=COMP,
            range=f"'Keepaフル'!A{grow}:{letter}{grow}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or [[]]
    )[0]
    def g(h):
        i = kidx[h]
        return vals[i] if i < len(vals) else None

    if str(g("ASIN") or "").strip().upper() != asin:
        return None
    l, w, h = num(g("梱包_L_cm")), num(g("梱包_W_cm")), num(g("梱包_H_cm"))
    wt = num(g("梱包_重量_g"))
    if l and w and h and l > 0 and w > 0 and h > 0:
        return float(l), float(w), float(h), (wt if wt and wt > 0 else None)
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    ap.add_argument("--asins", default="", help="カンマ区切り。空＝全行の空欄のみ")
    ap.add_argument(
        "--keepa-dims",
        action="store_true",
        help="リストにサイズが無いとき Keepa梱包寸法を使う（遅い）",
    )
    args = ap.parse_args()
    asin_filter = {a.strip().upper() for a in args.asins.split(",") if a.strip()}

    svc = sheets_service(write=not args.dry, interactive=False)
    settings = load_settings(svc)
    month = datetime.now(JST).month
    print("month", month, "sales_bands", len(settings["sales"]), "storage", settings["storage"])

    hdr = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{LIST}'!4:4")
        .execute()["values"][0]
    )
    idx = {str(h).strip(): i for i, h in enumerate(hdr)}
    for need in (
        "ASIN",
        "カテゴリー",
        "サイズ",
        "重量",
        "最安値",
        "FBA手数料",
        "販売手数料（セット）",
        "在庫保管手数料（セット）",
    ):
        if need not in idx:
            raise SystemExit(f"missing {need}")

    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{LIST}'!A5:CI",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    kcache = load_keepa_dims(svc) if args.keepa_dims else {"_amap": {}, "_kidx": {}, "_svc": None}

    bv_updates = []  # (row, value)
    bw_updates = []
    fba_updates = []
    stats = {
        "ref_fill": 0,
        "st_fill": 0,
        "fba_fill": 0,
        "ref_skip_cat": 0,
        "st_skip_dim": 0,
        "keepa_dim_used": 0,
    }

    for i, r in enumerate(rows):
        while len(r) < len(hdr):
            r.append("")
        asin = str(r[idx["ASIN"]] or "").strip().upper()
        if not re.match(r"^B[0-9A-Z]{9}$", asin):
            continue
        if asin_filter and asin not in asin_filter:
            continue
        row = i + 5
        sell = num(r[idx["最安値"]])
        cat = str(r[idx["カテゴリー"]] or "")
        ref_now = num(r[idx["販売手数料（セット）"]])
        st_now = num(r[idx["在庫保管手数料（セット）"]])
        fba_now = num(r[idx["FBA手数料"]])

        # dims: list first, else Keepa
        dims = parse_size_cm(str(r[idx["サイズ"]] or ""))
        wt = parse_weight_g(str(r[idx["重量"]] or ""))
        # Keepa寸法は --keepa-dims のときだけ（リストにサイズが無い行）
        if args.keepa_dims and dims is None and (st_now is None or fba_now is None):
            kd = keepa_dims_one(kcache, asin)
            if kd:
                dims = (kd[0], kd[1], kd[2])
                if wt is None:
                    wt = kd[3]
                stats["keepa_dim_used"] += 1

        # 販手
        if ref_now is None and sell and sell > 0:
            rate = pick_sales_rate(settings["sales"], cat, sell)
            if rate is None:
                stats["ref_skip_cat"] += 1
            else:
                yen = referral_fee_yen(rate, sell)
                bv_updates.append((row, yen))
                stats["ref_fill"] += 1
                r[idx["販売手数料（セット）"]] = yen

        # 保管
        if st_now is None:
            if not dims:
                stats["st_skip_dim"] += 1
            else:
                yen = storage_fee_yen(settings["storage"], dims[0], dims[1], dims[2], wt, month)
                if yen is not None:
                    bw_updates.append((row, yen))
                    stats["st_fill"] += 1
                    r[idx["在庫保管手数料（セット）"]] = yen

        # FBA（寸法＋重量が揃うときだけ）
        if fba_now is None and dims and wt is not None and wt > 0:
            picked = pick_fba_tier(settings["fba"], dims[0], dims[1], dims[2], wt)
            fee = num(picked.get("fee"))
            if fee is not None:
                fba_updates.append((row, int(round(fee))))
                stats["fba_fill"] += 1

    print("stats", stats)
    print("updates BV", len(bv_updates), "BW", len(bw_updates), "FBA", len(fba_updates))
    for label, ups in (("BV", bv_updates[:5]), ("BW", bw_updates[:5]), ("FBA", fba_updates[:5])):
        print("sample", label, ups)

    if args.dry:
        print("DRY ok")
        return

    def write_col(col_name: str, updates: list[tuple[int, float]]):
        if not updates:
            return
        letter = col_letter(idx[col_name] + 1)
        # batch by contiguous chunks is hard; write as data list of single cells in batches
        data = [
            {"range": f"'{LIST}'!{letter}{row}", "values": [[val]]}
            for row, val in updates
        ]
        for j in range(0, len(data), 100):
            svc.spreadsheets().values().batchUpdate(
                spreadsheetId=RES,
                body={"valueInputOption": "USER_ENTERED", "data": data[j : j + 100]},
            ).execute()
            print("WROTE", col_name, j, "/", len(data))

    write_col("販売手数料（セット）", bv_updates)
    write_col("在庫保管手数料（セット）", bw_updates)
    write_col("FBA手数料", fba_updates)
    print("DONE", stats)


if __name__ == "__main__":
    main()
