# -*- coding: utf-8 -*-
"""見積検品へ Keepaフルのメイン画像＋希望／提示価格＋値下げ交渉列を回収。

正:
- 画像は競合DB Keepaフルの「画像」列のみ（サブ画像禁止）
- 希望価格＝上限仕入（品番リスト 目標仕入額税抜）
- 提示価格＝問屋向け二段（セットで利益率≥30%なら上乗せ。200円下限は切らない）
- 値下げ交渉は人が書く＋機械印

使い方:
  python tools/common_ss_view/estimate_inspection_main_images.py --dry
  python tools/common_ss_view/estimate_inspection_main_images.py
"""
from __future__ import annotations

import math
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "competitor_store"))
from client import sheets_service  # noqa: E402

RES = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
COMP = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
SHEET = "見積検品"

# F判断 G画像 H問屋見積 I希望 J提示 K差分 L交渉（2026-11-15）
COL_JUDGE = 6
COL_IMG = 7
COL_QUOTE = 8
COL_HOPE = 9
COL_OFFER = 10
COL_DIFF = 11
COL_NEGO = 12

MARGIN_CAP = 0.30  # これ以上は安すぎ
MARGIN_AIM = 0.299  # 上乗せ後の着地（30%未満）
PROFIT_FLOOR = 200.0
PROFIT_PCT = 0.15

KEY_QUERIES = [
    ("食べるおだし かつお", ["食べるおだし", "かつお"], ["まぐろ", "ぶり", "チーズ", "はしっこ"], ["50g"]),
    ("食べるおだし まぐろ", ["食べるおだし", "まぐろ"], ["かつお", "ぶり", "チーズ", "はしっこ"], ["35g"]),
    ("食べるおだし ぶり", ["食べるおだし", "ぶり"], ["かつお", "まぐろ", "チーズ", "はしっこ"], ["40g"]),
    ("チーズかつお", ["チーズかつお"], ["業務用", "お徳用", "1kg", "kg", "各", "セット", "×", "ｘ"], ["220g"]),
    ("まぐろチーズ", ["まぐろチーズ"], ["業務用", "お徳用", "1kg", "kg", "各", "かつお", "セット", "｜"], ["190g"]),
]


def hyperlink_url(cell: str) -> str:
    s = str(cell or "").strip()
    if s.startswith("http"):
        return s.split("|")[0].strip()
    m = re.search(r'HYPERLINK\s*\(\s*"([^"]+)"', s, re.I)
    return m.group(1) if m else ""


def need_profit_set(sell: float) -> float:
    return max(float(math.floor(sell * PROFIT_PCT)), PROFIT_FLOOR)


def offer_from_hope(
    hope_ex: float,
    sell: float,
    fba: float,
    bags: int,
    tax: float,
) -> tuple[float, str]:
    """希望(上限)から提示を組む。戻り値=(提示税抜1個, 印)."""
    if not (hope_ex == hope_ex) or hope_ex <= 0:
        return hope_ex, ""
    if not (sell > 0) or not (bags >= 1) or not (tax >= 0):
        return round(hope_ex, 1), ""
    if not (fba == fba) or fba < 0:
        return round(hope_ex, 1), "FBA欠落のため提示=希望"

    def profit(cost_ex: float) -> float:
        return sell - fba - cost_ex * (1.0 + tax) * bags

    def margin(cost_ex: float) -> float:
        return profit(cost_ex) / sell

    # 単品は30%ルールなし
    if bags <= 1:
        return round(hope_ex, 1), ""

    need = need_profit_set(sell)
    m0 = margin(hope_ex)
    if m0 < MARGIN_CAP:
        return round(hope_ex, 1), ""

    # 利益率≥30%: 提示を上げて30%未満へ。ただし need は切らない
    target_profit = sell * MARGIN_AIM
    if target_profit < need:
        return (
            round(hope_ex, 1),
            "200円下限のため利益率30%超。上乗せ不可",
        )
    # cost_ex = (sell - fba - target_profit) / bags / (1+tax)
    offer = (sell - fba - target_profit) / bags / (1.0 + tax)
    # 希望より安くはしない（上限側に寄せる上乗せのみ）
    if offer < hope_ex:
        offer = hope_ex
    # 上げ過ぎて need 割れしないか再確認
    if profit(offer) + 1e-9 < need:
        offer = hope_ex
        return round(offer, 1), "200円下限のため利益率30%超。上乗せ不可"
    return round(offer, 1), "利益率30%未満になるよう提示を上乗せ"


def score_row(
    name: str, want: list[str], ban: list[str], prefer: list[str], maker_ok: bool
) -> int:
    n = str(name or "")
    if not all(w in n for w in want):
        return -1
    if any(b in n for b in ban):
        return -1
    sc = 100
    if maker_ok:
        sc += 50
    if "各" in n:
        sc -= 90
    m_sum = re.search(r"合計\s*(\d+)\s*袋", n)
    if m_sum:
        sc -= 20 + int(m_sum.group(1))
    elif "セット" in n:
        sc -= 40
    if "1袋" in n and "合計" not in n and "各" not in n:
        sc += 40
    for p in prefer:
        if p and p in n:
            sc += 60
            break
    if "業務用" in n or "お徳用" in n or re.search(r"1kg|１kg|1000g", n):
        sc -= 100
    return sc


def load_asin_index(svc) -> dict[str, int]:
    rows = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMP, range="'Keepaフル索引'!A:B")
        .execute()
        .get("values")
        or []
    )
    out: dict[str, int] = {}
    for r in rows[1:]:
        if len(r) < 2:
            continue
        a = str(r[0]).strip().upper()
        try:
            out[a] = int(float(r[1]))
        except Exception:
            continue
    return out


def col_letter(n1: int) -> str:
    """1-based column number → A1 letter."""
    n = n1
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def keepa_main_url(svc, asin: str, amap: dict[str, int], img_col0: int) -> str:
    row = amap.get(asin)
    if not row:
        return ""
    letter = col_letter(img_col0 + 1)
    rng = f"'Keepaフル'!{letter}{row}"
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMP, range=rng, valueRenderOption="FORMULA")
        .execute()
        .get("values")
        or []
    )
    cell = vals[0][0] if vals and vals[0] else ""
    return hyperlink_url(str(cell))


def main() -> int:
    dry = "--dry" in sys.argv
    svc = sheets_service(write=not dry, interactive=False)

    insp = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{SHEET}'!A1:T200")
        .execute()
        .get("values")
        or []
    )
    if not insp:
        print("見積検品 empty")
        return 1

    # Keepaフル header
    kh = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=COMP, range="'Keepaフル'!1:1")
        .execute()
        .get("values")
        or [[]]
    )[0]
    kidx = {str(h).strip(): i for i, h in enumerate(kh)}
    if "画像" not in kidx or "ASIN" not in kidx:
        print("Keepaフル missing 画像/ASIN")
        return 1
    amap = load_asin_index(svc)
    print("keepa_index", len(amap), "img_col", kidx["画像"])

    # 品番リスト
    hdr = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range="'仕入検討品番リスト'!4:4")
        .execute()
        .get("values")
        or [[]]
    )[0]
    headers = [str(c).strip() for c in hdr]
    idx = {h: i for i, h in enumerate(headers)}
    for need in (
        "ASIN",
        "メーカー名",
        "商品名",
        "単品商品名",
        "目標仕入額(税抜き）",
        "最安値",
        "FBA手数料",
        "税率",
        "同商品のセット数量",
    ):
        if need not in idx:
            print("missing", need)
            return 1

    rows_f = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range="'仕入検討品番リスト'!A5:ZZ",
            valueRenderOption="FORMULA",
        )
        .execute()
        .get("values")
        or []
    )
    rows_v = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range="'仕入検討品番リスト'!A5:ZZ",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    while len(rows_v) < len(rows_f):
        rows_v.append([])

    def fnum(x):
        try:
            return float(x)
        except Exception:
            return float("nan")

    catalog = []
    for i, row in enumerate(rows_f):
        while len(row) < len(headers):
            row.append("")
        rv = rows_v[i] if i < len(rows_v) else []
        while len(rv) < len(headers):
            rv.append("")
        asin = str(row[idx["ASIN"]]).strip().upper()
        if not re.match(r"^B[0-9A-Z]{9}$", asin):
            continue
        hope_n = fnum(rv[idx["目標仕入額(税抜き）"]])
        bags = fnum(rv[idx["同商品のセット数量"]])
        if not (bags > 0):
            bags = float("nan")
        tax = fnum(rv[idx["税率"]])
        if not (tax == tax) or tax <= 0:
            tax = 0.08
        catalog.append(
            {
                "asin": asin,
                "maker": str(row[idx["メーカー名"]]),
                "name": str(row[idx["商品名"]]),
                "single": str(row[idx["単品商品名"]]).strip(),
                "hope": hope_n if hope_n == hope_n else None,
                "sell": fnum(rv[idx["最安値"]]),
                "fba": fnum(rv[idx["FBA手数料"]]),
                "tax": tax,
                "bags": int(bags) if bags == bags else None,
            }
        )
    print("catalog", len(catalog))

    def pick(maker: str, key: str) -> dict | None:
        q = ban = prefer = None
        for prefix, words, bans, prefs in KEY_QUERIES:
            if key.startswith(prefix) or prefix in key:
                q, ban, prefer = words, bans, prefs
                break
        if q:
            best = None
            best_sc = -1
            for c in catalog:
                maker_ok = bool(maker) and (
                    maker in c["maker"] or ("石原" in maker and "石原" in c["maker"])
                )
                if "食べるおだし" in key and not maker_ok:
                    continue
                sc = score_row(c["name"], q, ban, prefer, maker_ok)
                if sc > best_sc:
                    best_sc = sc
                    best = c
            return best

        # 汎用: 商品キー「名前／仕入…」→ メーカー＋単品商品名一致
        base = str(key or "").split("／")[0].strip()
        if not base or not maker:
            return None
        cands = [
            c
            for c in catalog
            if maker in c["maker"] and c["single"] and c["single"] == base
        ]
        if not cands:
            # タイトルに単品名が全部入る行
            cands = [
                c
                for c in catalog
                if maker in c["maker"] and base and base in c["name"]
            ]
        if not cands:
            return None
        # 希望が取れる行を優先。同じなら袋数少ない方
        def rank_c(c):
            hope_ok = 0 if c["hope"] is not None and c["hope"] == c["hope"] else 1
            bags = c["bags"] if c["bags"] else 9999
            return (hope_ok, bags, c["asin"])

        return sorted(cands, key=rank_c)[0]

    img_vals = []
    hope_vals = []
    offer_vals = []
    nego_vals = []
    report = []
    for r_i, row in enumerate(insp[1:], start=2):
        while len(row) < 3:
            row.append("")
        maker = str(row[1] if len(row) > 1 else "")
        key = str(row[2] if len(row) > 2 else "")
        hit = pick(maker, key)
        if not hit:
            report.append((r_i, key, "NO_HIT", "", "", ""))
            img_vals.append([""])
            hope_vals.append([""])
            offer_vals.append([""])
            nego_vals.append([""])
            continue
        main_url = keepa_main_url(svc, hit["asin"], amap, kidx["画像"])
        if not main_url:
            img_vals.append([""])
        else:
            u = main_url.replace('"', "").replace("+", "%2B")
            img_vals.append([f'=IMAGE("{u}",1)'])
        hope = hit["hope"]
        hope_r = None
        offer = ""
        mark = ""
        if hope is not None and hope == hope:
            hope_r = round(float(hope), 1)
            hope_vals.append([hope_r])
            bags = hit["bags"] or 1
            offer, mark = offer_from_hope(
                hope_r, hit["sell"], hit["fba"], int(bags), float(hit["tax"])
            )
            offer_vals.append([offer if offer == offer else ""])
        else:
            hope_vals.append([""])
            offer_vals.append([""])
        # 値下げ交渉: 人の既存メモを残し、機械印を先頭に
        existing = ""
        hdr0 = insp[0] if insp else []
        if len(row) >= COL_NEGO and len(hdr0) >= COL_NEGO and "値下げ" in str(
            hdr0[COL_NEGO - 1]
        ):
            existing = str(row[COL_NEGO - 1] or "").strip()
        elif len(row) >= COL_OFFER and len(hdr0) >= COL_OFFER and "値下げ" in str(
            hdr0[COL_OFFER - 1]
        ):
            existing = str(row[COL_OFFER - 1] or "").strip()
        # 機械印だけの古いセルは上書きしてよい
        if existing.startswith("200円下限") or existing.startswith("利益率30%"):
            existing = ""
        if mark:
            existing = (mark + (" / " + existing if existing else "")).strip()
        nego_vals.append([existing])
        report.append(
            (
                r_i,
                key,
                hit["asin"],
                f"bags={hit['bags']}",
                f"hope={hope_r}",
                f"offer={offer_vals[-1][0]}",
                mark or "-",
            )
        )

    print("--- picks ---")
    for t in report:
        print(t)

    if dry:
        print("DRY ok", len(img_vals))
        return 0

    headers_out = [
        [
            "メイン画像",
            "問屋見積価格（税抜）",
            "希望価格（税抜）",
            "提示価格（税抜）",
            "差分（希望−見積）",
            "値下げ交渉",
        ]
    ]
    svc.spreadsheets().values().update(
        spreadsheetId=RES,
        range=f"'{SHEET}'!G1:L1",
        valueInputOption="USER_ENTERED",
        body={"values": headers_out},
    ).execute()
    n = len(img_vals)
    if n:
        end = 1 + n
        svc.spreadsheets().values().update(
            spreadsheetId=RES,
            range=f"'{SHEET}'!G2:G{end}",
            valueInputOption="USER_ENTERED",
            body={"values": img_vals},
        ).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=RES,
            range=f"'{SHEET}'!I2:I{end}",
            valueInputOption="USER_ENTERED",
            body={"values": hope_vals},
        ).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=RES,
            range=f"'{SHEET}'!J2:J{end}",
            valueInputOption="USER_ENTERED",
            body={"values": offer_vals},
        ).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=RES,
            range=f"'{SHEET}'!L2:L{end}",
            valueInputOption="USER_ENTERED",
            body={"values": nego_vals},
        ).execute()

    meta = (
        svc.spreadsheets()
        .get(
            spreadsheetId=RES,
            ranges=[f"'{SHEET}'"],
            fields="sheets(properties(sheetId,title))",
        )
        .execute()
    )
    sid = meta["sheets"][0]["properties"]["sheetId"]
    reqs = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sid,
                    "dimension": "ROWS",
                    "startIndex": 1,
                    "endIndex": 1 + n,
                },
                "properties": {"pixelSize": 90},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sid,
                    "dimension": "COLUMNS",
                    "startIndex": COL_IMG - 1,
                    "endIndex": COL_IMG,
                },
                "properties": {"pixelSize": 100},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sid,
                    "dimension": "COLUMNS",
                    "startIndex": COL_QUOTE - 1,
                    "endIndex": COL_NEGO,
                },
                "properties": {"pixelSize": 120},
                "fields": "pixelSize",
            }
        },
    ]
    svc.spreadsheets().batchUpdate(spreadsheetId=RES, body={"requests": reqs}).execute()
    print("WROTE F/G/H/I rows", n)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
