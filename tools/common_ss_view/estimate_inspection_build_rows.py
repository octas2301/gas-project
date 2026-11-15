# -*- coding: utf-8 -*-
"""見積検品へ、仕入付き見積候補を1商品1行で回収（汎用）。

正: docs/org/B_ESTIMATE_INSPECTION_REQUIREMENTS.md
- 本線 FBA。利益線はセット利益 ≥100円
- 同じ大カテ順位は販売数1つ。実月販が無ければ順位推計
- 業務用・お徳用・kg・500g超を、より小さい単品キーへ流用しない
- 条件増分（石原特売）は他メーカーでは出さない（差は0）
- 既存の石原行は触らない。同じ商品キーは最新1行に置換

使い方:
  python tools/common_ss_view/estimate_inspection_build_rows.py --dry
  python tools/common_ss_view/estimate_inspection_build_rows.py
"""
from __future__ import annotations

import argparse
import math
import re
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "competitor_store"))
from client import sheets_service  # noqa: E402

RES = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
SANK = "17nXSRjFp8xy-3EaBmp-LIu_tMc9ZOgfkhJoZjO3FLIc"
LIST = "仕入検討品番リスト"
SHEET = "見積検品"
PROFIT_LINE = 100
RANK_CAP = 60000
COL_L = 14  # 対象セット（1-based）。F右に見積・差分挿入後は N列=14


def utf16_len(s: str) -> int:
    """Sheets textFormatRuns の startIndex は UTF-16 コード単位。"""
    return len(s.encode("utf-16-le")) // 2


def num(v):
    try:
        if v is None or v == "":
            return None
        x = float(str(v).replace(",", "").replace("円", "").strip())
        return x if x == x else None
    except Exception:
        return None


def jan_norm(v) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    if s.replace(".", "", 1).isdigit():
        try:
            s = str(int(float(s)))
        except Exception:
            pass
    if s.isdigit() and len(s) < 13:
        s = s.zfill(13)
    return s


def col_letter(n1: int) -> str:
    n = n1
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def yen(n: float | int | None) -> str:
    if n is None or n != n:
        return "-"
    v = int(round(float(n)))
    if v < 0:
        return f"▲{abs(v):,}"
    return f"{v:,}"


def est_sold(rank: int) -> int:
    # 2.2 × 8,900 × 順位^−0.75 。1未満は1
    x = 2.2 * 8900.0 * (float(rank) ** -0.75)
    return max(1, int(math.floor(x + 0.5)))


def parse_grams(text: str) -> int | None:
    m = re.search(r"(\d+)\s*g", str(text or ""), re.I)
    return int(m.group(1)) if m else None


def is_bulk_title(title: str) -> bool:
    t = str(title or "")
    if "業務用" in t or "お徳用" in t:
        return True
    if re.search(r"(?<!\d)(1kg|１kg|1000g)", t, re.I):
        return True
    g = parse_grams(t)
    return bool(g and g >= 500)


def key_grams(name: str) -> int | None:
    return parse_grams(name)


def exclude_bulk_for_key(key_name: str, title: str) -> bool:
    """小さい単品キーに、より大きい規格（500g超・kg）のカタログを流用しない。

    タイトルに「業務用」とあっても、内容量がキーと同じなら残す（ケース売りの印）。
    """
    kg = key_grams(key_name)
    tg = parse_grams(title)
    if kg is not None and kg < 1000:
        if re.search(r"(?<!\d)(1kg|１kg|1000g)", title, re.I):
            return True
    if kg is not None and tg is not None and kg < 500 and tg >= 500:
        return True
    return False


def load_paste_meta(svc) -> dict[str, dict]:
    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=SANK,
            range="'貼付sheet'!A3:G",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    out = {}
    for r in rows[1:]:
        while len(r) < 7:
            r.append("")
        j = jan_norm(r[2])
        p = num(r[5])
        if len(j) < 8 or not p:
            continue
        out[j] = {
            "ex": p,
            "cond": str(r[6] or "").strip(),
            "maker": str(r[0] or "").strip(),
            "name": str(r[3] or "").strip(),
        }
    return out


def load_list(svc):
    hdr = [
        str(c).strip()
        for c in (
            (
                svc.spreadsheets()
                .values()
                .get(spreadsheetId=RES, range=f"'{LIST}'!4:4")
                .execute()
                .get("values")
                or [[]]
            )[0]
        )
    ]
    idx = {h: i for i, h in enumerate(hdr)}
    need = [
        "メーカー名",
        "ASIN",
        "商品名",
        "JANコード",
        "単品商品名",
        "同商品のセット数量",
        "過去1か月販売数（表示あれば）",
        "最安値",
        "順位",
        "1点仕入れ額（税込み）",
        "見積候補",
        "セット当たり利益",
        "利益の状態",
        "販売手数料（セット）",
        "在庫保管手数料（セット）",
        "FBA手数料",
    ]
    for n in need:
        if n not in idx:
            raise SystemExit(f"missing {n}")
    last = max(idx[n] for n in need)
    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{LIST}'!A5:{col_letter(last + 1)}",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    out = []
    for r in rows:
        while len(r) <= last:
            r.append("")
        if str(r[idx["見積候補"]]).strip() != "見積候補":
            continue
        name = str(r[idx["単品商品名"]]).strip()
        cost = num(r[idx["1点仕入れ額（税込み）"]])
        j = jan_norm(r[idx["JANコード"]])
        out.append(
            {
                "maker": str(r[idx["メーカー名"]]).strip(),
                "asin": str(r[idx["ASIN"]]).strip().upper(),
                "title": str(r[idx["商品名"]]).strip(),
                "jan": j,
                "name": name,
                "yokakunin": (not name) or ("要確認" in name),
                "q": int(num(r[idx["同商品のセット数量"]]) or 0),
                "sold": num(r[idx["過去1か月販売数（表示あれば）"]]),
                "sell": num(r[idx["最安値"]]),
                "rank": int(num(r[idx["順位"]]) or 0),
                "cost": cost if cost and cost > 0 else None,
                "set_profit": num(r[idx["セット当たり利益"]]),
                "state": str(r[idx["利益の状態"]]).strip(),
                "ref": num(r[idx["販売手数料（セット）"]]),
                "storage": num(r[idx["在庫保管手数料（セット）"]]),
                "fba": num(r[idx["FBA手数料"]]),
            }
        )
    return out


def title_matches_product(title: str, product_name: str) -> bool:
    """単品名の主要語がタイトルに入っているか（MIX・要確認行の兄弟拾い）。"""
    t = str(title or "")
    n = str(product_name or "")
    if not t or not n:
        return False
    tl = t.lower().replace(" ", "")
    g = parse_grams(n)
    if g is not None:
        if f"{g}g" not in tl and f"{g}ｇ" not in t:
            return False
    tokens = [
        w
        for w in re.split(r"[\s　/／]+", n)
        if len(w) >= 2 and not re.fullmatch(r"\d+g", w, re.I)
    ]
    if not tokens:
        return norm_name(n) in norm_name(t)
    return all(tok in t for tok in tokens)


def bag_label(title: str, q: int) -> str:
    if "各" in title or ("＋" in title and ("まぐろ" in title or "ぶり" in title)):
        return f"MIX{q}袋"
    return f"単品{q}袋"


def collect_rows_for_product(catalog: list[dict], maker: str, name: str, cost: float) -> list[dict]:
    """同一単品＋タイトル兄弟（要確認MIX含む）。"""
    out = []
    seen = set()
    nm = norm_name(maker)
    for r in catalog:
        if not (nm in norm_name(r["maker"]) or norm_name(r["maker"]) in nm):
            continue
        ok = False
        if not r.get("yokakunin") and norm_name(r["name"]) == norm_name(name):
            ok = True
        elif not r.get("yokakunin") and (
            norm_name(name) in norm_name(r["name"]) or norm_name(r["name"]) in norm_name(name)
        ):
            ok = True
        elif title_matches_product(r["title"], name):
            ok = True
        if not ok:
            continue
        if exclude_bulk_for_key(name, r["title"]):
            continue
        asin = r["asin"]
        if not asin or asin in seen:
            continue
        seen.add(asin)
        rr = dict(r)
        rr["cost"] = cost
        rr["name"] = name
        out.append(rr)
    return out


def set_profit_of(row: dict) -> tuple[float | None, str]:
    """戻り: (利益, メモ). 手数料不明は None。"""
    if row["state"] == "手数料不明":
        return None, "手数料不明"
    sp = row["set_profit"]
    if sp is not None:
        return float(sp), row["state"] or ""
    # 欠けるときは計算しない（正本どおり空は空）
    return None, "利益空"


def build_product(rows: list[dict], paste: dict[str, dict]) -> dict:
    maker = rows[0]["maker"]
    name = rows[0]["name"]
    cost = rows[0]["cost"]
    sets = []
    for r in rows:
        if exclude_bulk_for_key(name, r["title"]):
            continue
        if r["rank"] and r["rank"] > RANK_CAP:
            continue
        if not r["rank"]:
            continue
        if r["q"] < 1 or not r["sell"]:
            continue
        profit, note = set_profit_of(r)
        sets.append({**r, "profit": profit, "note": note})

    # 同じ袋数は売値が一番安いカタログ
    by_q: dict[int, list] = defaultdict(list)
    for s in sets:
        by_q[s["q"]].append(s)
    picked = []
    for q, lst in by_q.items():
        lst_sorted = sorted(lst, key=lambda x: (x["sell"] is None, x["sell"] or 9e18))
        picked.append(lst_sorted[0])
    picked.sort(key=lambda x: x["q"])

    # 順位ごとの実月販（利益の有無を問わず最大）
    sold_by_rank: dict[int, float] = {}
    for s in picked:
        if s["sold"] and s["sold"] > 0:
            sold_by_rank[s["rank"]] = max(sold_by_rank.get(s["rank"], 0), s["sold"])

    profitable = [s for s in picked if s["profit"] is not None and s["profit"] >= PROFIT_LINE]
    fee_unk = sum(1 for s in picked if s["note"] == "手数料不明")
    # 本線: 利益100以上のうち売値最安
    main = None
    if profitable:
        main = sorted(profitable, key=lambda x: x["sell"])[0]
        rank = main["rank"]
        if rank in sold_by_rank:
            sales = int(sold_by_rank[rank])
            sales_src = "実月販"
        else:
            sales = est_sold(rank)
            sales_src = "順位推計"
        units = sales * main["q"]  # 単品はセット袋数全部
        month_profit = int(round(sales * main["profit"]))
        decision = "可"
    else:
        sales = 0
        sales_src = ""
        units = 0
        month_profit = 0
        # 利益はあるが月販も順位もない → 月販不明（今の行は順位あり前提）
        if any(s["profit"] is not None and s["profit"] >= PROFIT_LINE for s in sets):
            decision = "月販不明"
        elif fee_unk > 0:
            # 利益判定が未完 → 見送りにせず人間へアラート
            decision = "手数料不明"
        else:
            decision = "見送り"

    # 対象セット文言（各行末尾にASIN。行全体をリンク）
    target_lines = []
    target_asins = []
    for s in picked:
        kind = bag_label(s.get("title") or "", s["q"])
        sell_s = yen(s["sell"])
        if s["profit"] is None:
            p_s = "-"
        else:
            p_s = yen(s["profit"])
        month = "月販-"
        if main and s["asin"] == main["asin"]:
            month = f"月販{sales}"
        elif main and s["rank"] == main["rank"]:
            month = "月販-（同順位）"
        elif s["sold"] and s["sold"] > 0 and (not main or s["rank"] != main["rank"]):
            month = f"月販{int(s['sold'])}"
        asin = s["asin"]
        target_lines.append(
            f"{kind}（{sell_s}円）／{s['rank']:,}位／FBA利益{p_s}／{month}／{asin}"
        )
        target_asins.append(asin)

    # 推奨
    rec_lines = []
    if main:
        rec_lines.append("【本線】")
        lab = bag_label(main.get("title") or "", main["q"])
        rec_lines.append(
            f"・{lab}（{yen(main['sell'])}円）"
            f"（利益{yen(main['profit'])}／月販{sales}（{sales_src}）"
            f"／入数{main['q']}／点数{units}／月利{yen(month_profit)}／経路:FBA）"
            f"／ASIN {main['asin']}"
        )
        refs = [
            s
            for s in profitable
            if s["asin"] != main["asin"]
        ]
        if refs:
            rec_lines.append("【参考】")
            for s in refs:
                tag = "同順位のため本線に含む" if main and s["rank"] == main["rank"] else "FBA"
                lab = bag_label(s.get("title") or "", s["q"])
                rec_lines.append(
                    f"・{lab}（{yen(s['sell'])}円）"
                    f"（利益{yen(s['profit'])}／{tag}）／ASIN {s['asin']}"
                )
    else:
        rec_lines.append("【本線】なし（利益100円以上のセットなし）")
        if picked:
            rec_lines.append("【参考】")
            for s in picked:
                note = s["note"] or ""
                lab = bag_label(s.get("title") or "", s["q"])
                rec_lines.append(
                    f"・{lab}（{yen(s['sell'])}円）"
                    f"（利益{yen(s['profit']) if s['profit'] is not None else '-'}／{note}）"
                    f"／ASIN {s['asin']}"
                )

    # 分析（機械・所持データのみ）。長い説明はここに残す
    loss = sum(1 for s in picked if s["profit"] is not None and s["profit"] < PROFIT_LINE)
    ok_n = len(profitable)
    analysis = [
        f"結論: {decision}",
        f"カタログ: {len(picked)}（利益≥100: {ok_n}／不足: {loss}／手数料不明: {fee_unk}）",
    ]
    if main:
        analysis.append(
            f"本線: 単品{main['q']}袋 {main['asin']}／{main['rank']:,}位／月販{sales}（{sales_src}）"
        )
    else:
        analysis.append("本線なし")

    # 可否右隣の端的な判断（人がぱっと見る用）
    judgment = short_judgment_(decision, main, picked, sales, sales_src)

    # 条件（貼付の文だけ。増分は0）
    jans = [r["jan"] for r in rows if r["jan"]]
    cond_txt = ""
    for j in jans:
        if j in paste and paste[j]["cond"]:
            cond_txt = f"見積条件: {paste[j]['cond']}"
            break
    if not cond_txt:
        cond_txt = "見積条件: （貼付の条件文なし）"

    kind = "単品"
    if len(picked) > 1:
        kind = f"単品＋拡張{len(picked)-1}"

    memo = "同じ順位は販売数1つ。貼付単価→仕入済み行の回収。条件増分は石原特売以外は0"

    return {
        "maker": maker,
        "name": name,
        "cost": int(round(cost)) if abs(cost - round(cost)) < 1e-6 else cost,
        "key": f"{name}／仕入{int(round(cost)) if abs(cost - round(cost)) < 1e-6 else cost}",
        "kind": kind,
        "decision": decision,
        "judgment": judgment,
        "analysis": "\n".join(analysis),
        "rec": "\n".join(rec_lines),
        "targets": "\n".join(target_lines) if target_lines else "（対象セットなし）",
        "sales_block": (
            f"Amazon\n単品{main['q']}袋: {sales}／点数{units}\n"
            if main
            else "Amazon\n本線なし\n"
        ),
        "units": units if decision == "可" else 0,
        "month_profit": month_profit if decision == "可" else 0,
        "cond": cond_txt,
        "cond_diff": (
            "利益が出るセット: 増分なし（他メーカー特売未適用）\n"
            f"仕入目安: {units}袋→{units}袋（+0袋）\n"
            f"見込利益: {yen(month_profit)}円→{yen(month_profit)}円（+0円）"
        ),
        "units_cond": 0,
        "profit_cond": 0,
        "memo": memo,
        "gray": decision == "見送り",
        "alert_fee": decision == "手数料不明",
        "asins": [s["asin"] for s in picked],
        "target_asins": target_asins,
    }


def short_judgment_(
    decision: str,
    main: dict | None,
    picked: list[dict],
    sales: int,
    sales_src: str,
) -> str:
    """仕入可否の右隣用。セットと利益を1行で。"""
    if decision == "可" and main:
        lab = bag_label(main.get("title") or "", main["q"])
        return (
            f"{lab}・売{yen(main['sell'])}円・利益{yen(main['profit'])}円"
            f"／月販{sales} → 可"
        )
    if decision == "月販不明":
        if main:
            lab = bag_label(main.get("title") or "", main["q"])
            return f"{lab}・利益{yen(main['profit'])}円だが月販・順位なし → 月販不明"
        return "利益は出るが月販・順位なし → 月販不明"
    if decision == "手数料不明":
        fu = next((s for s in picked if s.get("note") == "手数料不明"), None)
        if fu:
            lab = bag_label(fu.get("title") or "", fu["q"])
            return f"{lab}・売{yen(fu['sell'])}円・手数料欠落 → 判定不能（要確認）"
        return "手数料欠落 → 判定不能（要確認）"
    # 見送り
    known = [s for s in picked if s.get("profit") is not None]
    if known:
        s = sorted(known, key=lambda x: x["profit"], reverse=True)[0]
        lab = bag_label(s.get("title") or "", s["q"])
        return f"{lab}・利益{yen(s['profit'])}円（100円未満）→ 見送り"
    if picked:
        s = picked[0]
        lab = bag_label(s.get("title") or "", s["q"])
        return f"{lab}・利益判定できず → 見送り"
    return "利益100円以上のセットなし → 見送り"


def target_rich_runs(text: str, asins: list[str]) -> list[dict]:
    """各行全体を、その行の ASIN へのリンクにする（正本どおり）。"""
    if not text or text.startswith("（") or not asins:
        return []
    lines = text.split("\n")
    runs: list[dict] = []
    cursor = 0
    for i, line in enumerate(lines):
        asin = asins[i] if i < len(asins) else ""
        if asin:
            runs.append(
                {
                    "startIndex": cursor,
                    "format": {
                        "link": {"uri": f"https://www.amazon.co.jp/dp/{asin}"},
                    },
                }
            )
        cursor += utf16_len(line)
        if i < len(lines) - 1:
            cursor += utf16_len("\n")
    return runs


def apply_l_links_batch(svc, sheet_id: int, items: list[dict]):
    """items: {row, text, asins, gray}。L列だけ更新。

    リンクは textFormatRuns のみ先に書き、書式（折返し・背景）は次リクエストで付ける。
    同時指定すると runs が落ちることがある。
    """
    link_reqs = []
    fmt_reqs = []
    for it in items:
        text = it["text"]
        runs = target_rich_runs(text, it.get("asins") or [])
        row_1based = it["row"]
        rng = {
            "sheetId": sheet_id,
            "startRowIndex": row_1based - 1,
            "endRowIndex": row_1based,
            "startColumnIndex": COL_L - 1,
            "endColumnIndex": COL_L,
        }
        cell = {"userEnteredValue": {"stringValue": text}}
        fields = "userEnteredValue"
        if runs:
            cell["textFormatRuns"] = runs
            fields += ",textFormatRuns"
        link_reqs.append(
            {
                "updateCells": {
                    "range": rng,
                    "rows": [{"values": [cell]}],
                    "fields": fields,
                }
            }
        )
        fmt_reqs.append(
            {
                "repeatCell": {
                    "range": rng,
                    "cell": {
                        "userEnteredFormat": {
                            "wrapStrategy": "WRAP",
                            "verticalAlignment": "TOP",
                            "backgroundColor": (
                                {"red": 0.85, "green": 0.85, "blue": 0.85}
                                if it.get("gray")
                                else {"red": 1, "green": 1, "blue": 1}
                            ),
                        }
                    },
                    "fields": "userEnteredFormat(wrapStrategy,verticalAlignment,backgroundColor)",
                }
            }
        )
    # 書式→リンクの順（リンクを後勝ち）
    for i in range(0, len(fmt_reqs), 40):
        svc.spreadsheets().batchUpdate(
            spreadsheetId=RES, body={"requests": fmt_reqs[i : i + 40]}
        ).execute()
    for i in range(0, len(link_reqs), 40):
        svc.spreadsheets().batchUpdate(
            spreadsheetId=RES, body={"requests": link_reqs[i : i + 40]}
        ).execute()


def norm_name(s: str) -> str:
    return re.sub(r"\s+", "", str(s or ""))


def parse_key(key: str) -> tuple[str, float | None]:
    base = str(key or "").split("／")[0].strip()
    m = re.search(r"仕入\s*([0-9.]+)", str(key or ""))
    cost = float(m.group(1)) if m else None
    return base, cost


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    ap.add_argument(
        "--include-ishihara",
        action="store_true",
        help="石原も再生成（既定は石原以外）",
    )
    ap.add_argument(
        "--jans",
        default="",
        help="カンマ区切りJANに限定（空＝貼付にある仕入行のうち石原以外）",
    )
    ap.add_argument(
        "--relink",
        action="store_true",
        help="行の追加なし。既存の見積検品L列（対象セット）にASINリンクだけ付け直す",
    )
    args = ap.parse_args()

    # 第1弾（2026-11-14 貼付→仕入）の12 JAN
    DEFAULT_JANS = {
        "4964937022633",
        "4964937022640",
        "4964937022664",
        "4904001022596",
        "4901276130018",
        "4522968301785",
        "4901159004818",
        "4901159047532",
        "4589850821806",
        "4582152498264",
        "4582152498622",
        "4902388003207",
    }
    jan_filter = {jan_norm(x) for x in args.jans.split(",") if x.strip()} or set(
        DEFAULT_JANS
    )

    svc = sheets_service(write=not args.dry, interactive=False)
    paste = load_paste_meta(svc)
    catalog = load_list(svc)

    if args.relink:
        relink_existing_(svc, catalog, paste, dry=args.dry)
        return
    seeds = []  # (maker, name, cost)
    seen_seed = set()
    for r in catalog:
        if r.get("yokakunin"):
            continue
        if not r.get("cost"):
            continue
        if not args.include_ishihara and "石原" in r["maker"]:
            continue
        if r["jan"] not in jan_filter:
            continue
        k = (r["maker"], r["name"], r["cost"])
        if k in seen_seed:
            continue
        seen_seed.add(k)
        seeds.append(k)

    products = []
    for maker, name, cost in sorted(seeds, key=lambda x: (x[0], x[1])):
        rows = collect_rows_for_product(catalog, maker, name, cost)
        if not rows:
            continue
        products.append(build_product(rows, paste))

    print(f"products={len(products)}")
    for p in products:
        print(
            f"  {p['decision']:4} {p['maker']} | {p['key']} | "
            f"袋={p['units']} 月利={p['month_profit']} asins={p['asins']}"
        )

    # 既存シート
    existing = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{SHEET}'!A1:T")
        .execute()
        .get("values")
        or []
    )
    if not existing:
        raise SystemExit("見積検品 missing")
    header = existing[0]
    old_rows = existing[1:]
    keep = []
    new_keys = {p["key"] for p in products}
    for r in old_rows:
        while len(r) < 3:
            r.append("")
        key = str(r[2])
        maker = str(r[1])
        if key in new_keys:
            continue  # replace
        if (not args.include_ishihara) and "石原" in maker:
            keep.append(r)
            continue
        if key not in new_keys:
            keep.append(r)

    today = date.today().isoformat()
    new_rows = []
    for p in products:
        new_rows.append(
            [
                today,
                p["maker"],
                p["key"],
                p["kind"],
                p["decision"],
                "",  # F 画像 → 別スクリプト
                "",  # G 希望
                "",  # H 提示
                "",  # I 交渉
                p["analysis"],
                p["rec"],
                p["targets"],
                p["sales_block"],
                p["units"],
                p["month_profit"],
                p["cond"],
                p["cond_diff"],
                p["units_cond"],
                p["profit_cond"],
                p["memo"],
            ]
        )

    out = [header] + keep + new_rows
    print(f"write rows: keep={len(keep)} new={len(new_rows)} total_data={len(keep)+len(new_rows)}")

    if args.dry:
        print("DRY ok")
        for p in products[:3]:
            print("--- sample", p["key"], "---")
            print(p["analysis"])
            print(p["rec"][:400])
        return

    # clear + write
    # ensure enough rows
    end_row = 1 + len(out) - 1
    svc.spreadsheets().values().clear(
        spreadsheetId=RES,
        range=f"'{SHEET}'!A2:T",
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=RES,
        range=f"'{SHEET}'!A1:T{end_row}",
        valueInputOption="USER_ENTERED",
        body={"values": out},
    ).execute()

    # gray 見送り
    meta = (
        svc.spreadsheets()
        .get(
            spreadsheetId=RES,
            ranges=[f"'{SHEET}'"],
            fields="sheets(properties(sheetId))",
        )
        .execute()
    )
    sid = meta["sheets"][0]["properties"]["sheetId"]
    reqs = []
    # reset bg then gray
    data_end = len(out)
    reqs.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sid,
                    "startRowIndex": 1,
                    "endRowIndex": data_end,
                    "startColumnIndex": 0,
                    "endColumnIndex": 20,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 1, "green": 1, "blue": 1}
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        }
    )
    for i, row in enumerate(out[1:], start=1):
        if len(row) > 4 and row[4] == "見送り":
            reqs.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sid,
                            "startRowIndex": i,
                            "endRowIndex": i + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 20,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 0.85,
                                    "green": 0.85,
                                    "blue": 0.85,
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor",
                    }
                }
            )
    if reqs:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=RES, body={"requests": reqs}
        ).execute()

    # L列 ASINリンク（textFormatRuns）
    link_items = []
    for i, p in enumerate(new_rows):
        sheet_row = 1 + len(keep) + i + 1  # header + keep + 1-based
        prod = products[i]
        link_items.append(
            {
                "row": sheet_row,
                "text": prod["targets"],
                "asins": prod.get("target_asins") or prod.get("asins") or [],
                "gray": prod.get("gray"),
            }
        )
    if link_items:
        apply_l_links_batch(svc, sid, link_items)
        print("LINKS L", len(link_items))
    print("WROTE 見積検品", len(new_rows), "new")


def relink_existing_(svc, catalog, paste, dry: bool = False):
    """既存行の対象セット列だけ、兄弟ASIN込みで付け直す＋リンク。"""
    existing = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{SHEET}'!A1:Z")
        .execute()
        .get("values")
        or []
    )
    if len(existing) < 2:
        print("no rows")
        return
    hdr = [str(c).strip() for c in existing[0]]
    # 対象セット列を見出しから探す（列挿入後も追従）
    global COL_L
    try:
        COL_L = next(i for i, h in enumerate(hdr) if "対象セット" in h) + 1
    except StopIteration:
        pass
    print("target col", COL_L, hdr[COL_L - 1] if COL_L <= len(hdr) else "?")

    items = []
    for i, row in enumerate(existing[1:], start=2):
        while len(row) < 5:
            row.append("")
        maker = str(row[1]).strip()
        key = str(row[2]).strip()
        base, cost = parse_key(key)
        if not base or not cost:
            print("skip parse", i, key)
            continue
        rows = collect_rows_for_product(catalog, maker, base, cost)
        if not rows:
            print("skip no catalog", i, key)
            continue
        prod = build_product(rows, paste)
        items.append(
            {
                "row": i,
                "text": prod["targets"],
                "asins": prod.get("target_asins") or [],
                "gray": str(row[4]).strip() == "見送り" or prod.get("gray"),
                "key": key,
            }
        )
        print(
            f"  r{i} {key[:28]} lines={len(prod.get('target_asins') or [])} "
            f"asins={prod.get('target_asins')}"
        )

    print(f"relink targets={len(items)}")
    if dry:
        for it in items[:2]:
            print(it["text"][:300])
            print("runs", target_rich_runs(it["text"], it["asins"])[:3])
        print("DRY ok")
        return

    meta = (
        svc.spreadsheets()
        .get(
            spreadsheetId=RES,
            ranges=[f"'{SHEET}'"],
            fields="sheets(properties(sheetId))",
        )
        .execute()
    )
    sid = meta["sheets"][0]["properties"]["sheetId"]
    apply_l_links_batch(svc, sid, items)
    print("RELINKED L", len(items))


if __name__ == "__main__":
    main()

