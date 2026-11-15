# -*- coding: utf-8 -*-
"""見積検品へ、品番リストから1商品1行で回収（汎用）。

正: docs/org/B_ESTIMATE_INSPECTION_REQUIREMENTS.md
- 既定: 見積候補かつ仕入あり
- `--makers` + `--include-skip` + `--allow-no-cost` で門落ち見送り・仕入なしも載せる
- 本線 FBA。利益線はセット利益 ≥100円
- 同じ大カテ順位は販売数1つ。実月販が無ければ順位推計
- 業務用・お徳用・kg・500g超を、より小さい単品キーへ流用しない
- 条件増分（石原特売）は他メーカーでは出さない（差は0）
- 既存の石原行は触らない。同じ商品キーは最新1行に置換

使い方:
  python tools/common_ss_view/estimate_inspection_build_rows.py --dry
  python tools/common_ss_view/estimate_inspection_build_rows.py
  python tools/common_ss_view/estimate_inspection_build_rows.py --makers オリエンタル,株式会社オリエンタル,オリエンタル洋行 --include-skip --allow-no-cost
"""
from __future__ import annotations

import argparse
import math
import re
import sys
import unicodedata
from collections import defaultdict
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "competitor_store"))
from client import sheets_service  # noqa: E402
import estimate_paste_match as pm  # noqa: E402

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


def load_paste_rows(svc) -> list[dict]:
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
    return pm.paste_rows_from_values(rows)


def load_paste_meta(svc) -> dict[str, dict]:
    """互換: JAN → 貼付メタ。"""
    return pm.paste_by_jan(load_paste_rows(svc))


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
        flag = str(r[idx["見積候補"]]).strip()
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
                "flag": flag,
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


def listing_core_name_(title: str) -> str:
    """Amazon出品タイトルからセット枚数・装飾を除き、単品キー候補にする。"""
    t = unicodedata.normalize("NFKC", str(title or ""))
    t = re.sub(r"【[^】]*】", " ", t)
    t = re.sub(r"★[^★]*★", " ", t)
    t = re.sub(r"（[^）]*セット[^）]*）", " ", t)
    t = re.sub(r"\([^)]*セット[^)]*\)", " ", t)
    t = re.sub(r"\d+\s*箱セット", " ", t)
    t = re.sub(r"[×xX]\s*\d+\s*(箱|袋|個|本|入)?", " ", t)
    t = re.sub(r"まとめ買い", " ", t)
    t = re.sub(r"お試しセット", " ", t)
    t = re.sub(r"ゆうパケット.*$", " ", t)
    t = re.sub(r"\[.*?\]", " ", t)
    # 出品文言のゆれ（同じ単品なのに行が割れるのを防ぐ）
    t = re.sub(r"チャツネ\s*付?", " ", t)
    t = re.sub(r"粉末タイプ", " ", t)
    t = re.sub(r"レトロカレー", " ", t)
    t = re.sub(r"マースル[ウウー－ー]+", " ", t)
    t = re.sub(r"ル[ウウー－ー]+", " ", t)
    for prefix in ("オリエンタル洋行", "株式会社オリエンタル", "オリエンタル"):
        if t.startswith(prefix):
            t = t[len(prefix) :].strip()
            break
    t = re.sub(r"\s+", " ", t).strip(" 　/-")
    return t[:48] if t else ""


def collect_rows_for_product(
    catalog: list[dict],
    maker: str,
    name: str,
    cost: float,
    makers_scope: list[str] | None = None,
    paste_jan: str | None = None,
) -> list[dict]:
    """同一単品＋タイトル兄弟（要確認MIX含む）。

    paste_jan あり（貼付で特定した単品）: 原則そのJANの行だけ。
    「チャツネ付」など付属語で別単品へ巻き込まない。
    """
    out = []
    seen = set()
    scope = [m for m in (makers_scope or []) if m]
    nm = norm_name(maker)
    pn = norm_name(name)
    pj = jan_norm(paste_jan) if paste_jan else ""
    for r in catalog:
        rm = str(r.get("maker") or "").strip()
        if scope:
            if rm not in scope:
                continue
        else:
            if not (nm in norm_name(rm) or norm_name(rm) in nm):
                continue
        rj = jan_norm(r.get("jan") or "")
        ok = False
        if pj:
            # 貼付特定商品: JAN一致のみ（タイトル部分一致は付属チャツネ等で誤結合する）
            ok = bool(rj and rj == pj)
        else:
            if not r.get("yokakunin") and norm_name(r["name"]) == pn:
                ok = True
            elif not r.get("yokakunin") and (
                pn in norm_name(r["name"]) or norm_name(r["name"]) in pn
            ):
                ok = True
            elif title_matches_product(r["title"], name):
                ok = True
            elif pn and pn in norm_name(r.get("title") or ""):
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


def resolve_product_seed_(
    r: dict,
    paste_rows: list[dict],
    paste_map: dict[str, dict],
    makers_filter: list[str] | None,
    allow_name_partial: bool,
) -> dict:
    """1カタログ行 → 1商品キー（ASINタイトルではなく単品／貼付名）。"""
    name = str(r.get("name") or "").strip()
    title = str(r.get("title") or "").strip()
    match = pm.resolve_paste_match(
        paste_rows=paste_rows,
        paste_map=paste_map,
        jan=r.get("jan") or "",
        maker=r.get("maker") or "",
        name=name if name and "要確認" not in name else (title[:40] if title else ""),
        title=title,
        makers_filter=makers_filter,
        allow_name_partial=allow_name_partial,
    )
    cost = r.get("cost") if r.get("cost") else None
    if match["kind"] in (pm.MATCH_JAN, pm.MATCH_NAME) and match.get("paste_name"):
        prod = str(match["paste_name"]).strip()
        maker = str(match.get("paste_maker") or r.get("maker") or "").strip()
        # 同じ貼付JANは名前ゆれ（500ｇ付き等）でも1商品
        pj = str(match.get("jan") or "").strip()
        group = f"jan:{pj}" if pj else norm_name(prod)
        return {
            "maker": maker,
            "name": prod,
            "cost": cost,
            "match_kind": match["kind"],
            "group": group,
        }
    if name and not r.get("yokakunin") and "要確認" not in name:
        return {
            "maker": str(r.get("maker") or "").strip(),
            "name": name,
            "cost": cost,
            "match_kind": match["kind"],
            "group": norm_name(name),
        }
    core = listing_core_name_(title)
    if not core:
        core = title[:40] or str(r.get("asin") or "")
    return {
        "maker": str(r.get("maker") or "").strip(),
        "name": core,
        "cost": cost,
        "match_kind": match["kind"],
        "group": norm_name(core),
    }


def set_profit_of(row: dict) -> tuple[float | None, str]:
    """戻り: (利益, メモ). 手数料不明／見積なしは None。"""
    if not row.get("cost"):
        return None, "見積なし"
    if row["state"] == "手数料不明":
        return None, "手数料不明"
    sp = row["set_profit"]
    if sp is not None:
        return float(sp), row["state"] or ""
    # 欠けるときは計算しない（正本どおり空は空）
    return None, "利益空"


def build_product(
    rows: list[dict],
    paste: dict[str, dict],
    *,
    paste_rows: list[dict] | None = None,
    makers_filter: list[str] | None = None,
    allow_name_partial: bool = True,
) -> dict:
    maker = rows[0]["maker"]
    name = rows[0]["name"]
    cost = rows[0]["cost"]
    prow = paste_rows if paste_rows is not None else list(paste.values())
    match = pm.best_match_for_rows(
        rows,
        prow,
        paste,
        makers_filter=makers_filter,
        allow_name_partial=allow_name_partial,
    )
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
        elif any(s.get("note") == "見積なし" for s in picked) or not cost:
            decision = "見積なし"
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
    cond_txt = ""
    if match.get("cond"):
        cond_txt = f"見積条件: {match['cond']}"
    else:
        jans = [r["jan"] for r in rows if r["jan"]]
        for j in jans:
            if j in paste and paste[j].get("cond"):
                cond_txt = f"見積条件: {paste[j]['cond']}"
                break
    if not cond_txt:
        cond_txt = "見積条件: （貼付の条件文なし）"

    kind = "単品"
    if len(picked) > 1:
        kind = f"単品＋拡張{len(picked)-1}"

    memo = (
        "同じ順位は販売数1つ。貼付単価→仕入済み行の回収。条件増分は石原特売以外は0｜"
        + match["label"]
    )
    if match["kind"] == pm.MATCH_NAME and match.get("paste_name"):
        memo += f"（貼付名:{match['paste_name']}）"

    if cost is None:
        cost_disp = None
        key = f"{name}／仕入なし"
    else:
        cost_disp = int(round(cost)) if abs(cost - round(cost)) < 1e-6 else cost
        key = f"{name}／仕入{cost_disp}"
    if decision == "見積なし" and not judgment:
        judgment = "問屋見積なし → 依頼候補"
    # 一致度は判断の先頭（人が見落とさない）
    if match["kind"] != pm.MATCH_NONE:
        judgment = match["label"] + "｜" + (judgment or "")
    elif decision == "見積なし":
        judgment = match["label"] + "｜" + (judgment or "")
    analysis.insert(1, match["label"])
    return {
        "maker": maker,
        "name": name,
        "cost": cost_disp if cost_disp is not None else "",
        "key": key,
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
        "gray": decision in ("見送り", "見積なし"),
        "alert_fee": decision == "手数料不明",
        "alert_match": bool(match.get("alert")),
        "match_kind": match["kind"],
        "match_label": match["label"],
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
    if decision == "見積なし":
        if picked:
            s = picked[0]
            lab = bag_label(s.get("title") or "", s["q"])
            return f"{lab}・問屋見積なし → 依頼候補"
        return "問屋見積なし → 依頼候補"
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
    raw = str(key or "")
    if "仕入なし" in raw:
        return base, None
    m = re.search(r"仕入\s*([0-9.]+)", raw)
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
    ap.add_argument(
        "--makers",
        default="",
        help="カンマ区切りメーカー名（部分一致）。指定時はJAN既定フィルタを使わない",
    )
    ap.add_argument(
        "--include-skip",
        action="store_true",
        help="見送り行も種にする（門落ち込み検品用）",
    )
    ap.add_argument(
        "--allow-no-cost",
        action="store_true",
        help="仕入なし行も種にする（見積なし判定）",
    )
    ap.add_argument(
        "--no-name-partial",
        action="store_true",
        help="商品名部分一致を使わない（既定は使う）",
    )
    args = ap.parse_args()
    allow_name_partial = not bool(args.no_name_partial)

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
    maker_filter = [m.strip() for m in args.makers.split(",") if m.strip()]
    jan_filter = {jan_norm(x) for x in args.jans.split(",") if x.strip()}
    if not jan_filter and not maker_filter:
        jan_filter = set(DEFAULT_JANS)

    svc = sheets_service(write=not args.dry, interactive=False)
    paste_rows = load_paste_rows(svc)
    paste = pm.paste_by_jan(paste_rows)
    catalog = load_list(svc)

    if args.relink:
        relink_existing_(svc, catalog, paste, dry=args.dry)
        return
    # 種は「1商品」。ASINタイトルごとではなく貼付名／単品名でまとめる
    seed_map: dict[str, dict] = {}
    for r in catalog:
        if maker_filter:
            # 完全一致（「オリエンタル」がエンタープライズ等を飲み込まない）
            if r["maker"] not in maker_filter:
                continue
        flag = str(r.get("flag") or "").strip()
        if flag == "見積候補":
            pass
        elif args.include_skip and flag == "見送り":
            pass
        else:
            continue
        name0 = str(r.get("name") or "").strip()
        if r.get("yokakunin"):
            if maker_filter and args.allow_no_cost:
                if not name0 and not str(r.get("title") or "").strip() and not r.get("asin"):
                    continue
            else:
                continue
        if not r.get("cost") and not args.allow_no_cost:
            continue
        if not args.include_ishihara and "石原" in r["maker"] and not maker_filter:
            continue
        if jan_filter and r["jan"] not in jan_filter:
            continue
        seed = resolve_product_seed_(
            r,
            paste_rows,
            paste,
            maker_filter or None,
            allow_name_partial,
        )
        if not seed.get("name"):
            continue
        # メーカー指定時はメーカー表記ゆれを同一グループに（洋行／株式会社も同じ単品）
        g_maker = "|".join(maker_filter) if maker_filter else seed["maker"]
        gk = f"{g_maker}::{seed['group']}"
        seed["members"] = [r]
        prev = seed_map.get(gk)
        if not prev:
            seed_map[gk] = seed
            continue
        prev["members"] = prev.get("members") or []
        prev["members"].extend(seed["members"])
        # 仕入ありを優先。同じなら JAN一致を優先
        rank = {"jan": 2, "name_partial": 1, "ambiguous": 0, "none": 0}
        better_cost = seed.get("cost") and not prev.get("cost")
        better_match = rank.get(seed.get("match_kind") or "none", 0) > rank.get(
            prev.get("match_kind") or "none", 0
        )
        if better_cost or (seed.get("cost") and better_match) or (
            not prev.get("cost") and better_match
        ):
            if prev.get("cost") and not seed.get("cost"):
                seed["cost"] = prev["cost"]
            seed["members"] = prev["members"]
            seed_map[gk] = seed
        elif prev.get("cost") is None and seed.get("cost"):
            prev["cost"] = seed["cost"]

    # 兄弟ASINを拾ってから、同じASINを持つ種を1商品へ
    def _seed_rank_(s: dict) -> tuple:
        rk = {"jan": 2, "name_partial": 1, "ambiguous": 0, "none": 0}
        return (
            1 if s.get("cost") else 0,
            rk.get(s.get("match_kind") or "none", 0),
            1 if str(s.get("group") or "").startswith("jan:") else 0,
        )

    def _merge_seed_(keep: str, drop: str) -> None:
        if keep == drop or keep not in seed_map or drop not in seed_map:
            return
        km, dm = seed_map[keep], seed_map[drop]
        km["members"] = (km.get("members") or []) + (dm.get("members") or [])
        if dm.get("cost") and not km.get("cost"):
            km["cost"] = dm["cost"]
        if _seed_rank_(dm) > _seed_rank_(km):
            km["name"] = dm["name"]
            km["maker"] = dm["maker"]
            km["match_kind"] = dm["match_kind"]
            km["group"] = dm["group"]
            if dm.get("cost"):
                km["cost"] = dm["cost"]
        del seed_map[drop]

    for gk, s in list(seed_map.items()):
        paste_jan = None
        if str(s.get("group") or "").startswith("jan:"):
            paste_jan = str(s.get("group") or "").split(":", 1)[-1]
        rows = collect_rows_for_product(
            catalog,
            s["maker"],
            s["name"],
            s.get("cost"),
            makers_scope=maker_filter or None,
            paste_jan=paste_jan,
        )
        if rows:
            s["members"] = rows
        else:
            # 核名が変わり兄弟検索に落ちたとき、種になった行だけ
            fixed = []
            seen_a = set()
            for r in s.get("members") or []:
                a = r.get("asin")
                if not a or a in seen_a:
                    continue
                seen_a.add(a)
                rr = dict(r)
                rr["name"] = s["name"]
                rr["cost"] = s.get("cost")
                fixed.append(rr)
            s["members"] = fixed

    def _prod_stem_(n: str) -> str:
        """グラム・（大）（小）だけ除く。レトルト／辛口は別単品として残す。"""
        t = norm_name(n)
        t = re.sub(r"\d+g", "", t)
        t = re.sub(r"[（(]?[大小][）)]?$", "", t)
        t = re.sub(r"[大小]$", "", t)
        return t

    # タイトル核の種を、同じ貼付単品（jan:）へ吸収（セット違い出品のまとめ）
    for gk, s in list(seed_map.items()):
        g = str(s.get("group") or "")
        if g.startswith("jan:") or gk not in seed_map:
            continue
        stem = _prod_stem_(s.get("name") or "")
        if len(stem) < 4:
            continue
        s_bulk = is_bulk_title(s.get("name") or "")
        best = None
        best_gk = None
        for jgk, js in seed_map.items():
            jg = str(js.get("group") or "")
            if not jg.startswith("jan:"):
                continue
            if is_bulk_title(js.get("name") or "") != s_bulk:
                continue
            jstem = _prod_stem_(js.get("name") or "")
            if len(jstem) < 4:
                continue
            # 片方が他方を包含し、かつ長さ差が小さい（別商品の誤吸収を防ぐ）
            if stem == jstem or (
                (stem in jstem or jstem in stem)
                and abs(len(stem) - len(jstem)) <= 4
            ):
                if best is None or _seed_rank_(js) > _seed_rank_(best):
                    best = js
                    best_gk = jgk
        if best is not None and best_gk:
            _merge_seed_(best_gk, gk)

    changed = True
    while changed:
        changed = False
        asin_owner: dict[str, str] = {}
        for gk, s in list(seed_map.items()):
            for r in s.get("members") or []:
                a = str(r.get("asin") or "").strip().upper()
                if not a:
                    continue
                other = asin_owner.get(a)
                if not other:
                    asin_owner[a] = gk
                    continue
                if other == gk or other not in seed_map or gk not in seed_map:
                    continue
                g1 = str(seed_map[other].get("group") or "")
                g2 = str(seed_map[gk].get("group") or "")
                # 別の貼付JAN同士はASINが被ってもマージしない
                if g1.startswith("jan:") and g2.startswith("jan:") and g1 != g2:
                    continue
                keep, drop = other, gk
                if _seed_rank_(seed_map[gk]) > _seed_rank_(seed_map[other]):
                    keep, drop = gk, other
                _merge_seed_(keep, drop)
                changed = True
                break
            if changed:
                break

    products = []
    seen_prod_key = set()
    for s in sorted(seed_map.values(), key=lambda x: (x["maker"], x["name"])):
        rows = s.get("members") or []
        if not rows:
            continue
        maker = s["maker"]
        name = s["name"]
        cost = s.get("cost")
        makers_in = [str(x.get("maker") or "") for x in rows]
        if makers_in:
            maker = max(set(makers_in), key=makers_in.count)
        for x in rows:
            x["name"] = name
            x["cost"] = cost
            x["maker"] = maker
        prod = build_product(
            rows,
            paste,
            paste_rows=paste_rows,
            makers_filter=maker_filter or None,
            allow_name_partial=allow_name_partial,
        )
        if prod["key"] in seen_prod_key:
            continue
        seen_prod_key.add(prod["key"])
        products.append(prod)

    print(f"products={len(products)}")
    for p in products:
        print(
            f"  {p['decision']:4} [{p.get('match_kind')}] {p['maker']} | {p['key']} | "
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
        if maker_filter and maker in maker_filter:
            continue  # メーカー指定の再生成で当該メーカーの旧行は落とす
        if (not args.include_ishihara) and "石原" in maker:
            keep.append(r)
            continue
        if key not in new_keys:
            keep.append(r)

    today = date.today().isoformat()
    hidx = {str(h).strip(): i for i, h in enumerate(header)}

    def cell_row_(p: dict) -> list:
        # 現行ヘッダに合わせて埋める（判断＝接合一致度の先頭表示）
        n = max(len(header), 20)
        row = [""] * n
        def put(col_name: str, val):
            i = hidx.get(col_name)
            if i is not None and i < n:
                row[i] = val

        put("提案日", today)
        put("メーカー", p["maker"])
        put("商品キー", p["key"])
        put("種別", p["kind"])
        put("仕入可否", p["decision"])
        put("判断", p.get("judgment") or "")
        put("分析結果", p["analysis"])
        put("推奨", p["rec"])
        put("対象セット", p["targets"])
        # 旧名ゆれ
        if "対象セット（順位・月販）" in hidx:
            put("対象セット（順位・月販）", p["targets"])
        put("月販想定", p["sales_block"])
        put("仕入目安袋", p["units"])
        put("月次利益", p["month_profit"])
        put("条件", p["cond"])
        put("条件時セット（利益差）", p["cond_diff"])
        put("仕入目安袋_条件", p["units_cond"])
        put("月次利益_条件", p["profit_cond"])
        put("判断メモ", p["memo"])
        # ヘッダが古い20列固定のとき用フォールバック
        if "仕入可否" not in hidx:
            return [
                today,
                p["maker"],
                p["key"],
                p["kind"],
                p["decision"],
                p.get("judgment") or "",
                "",
                "",
                "",
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
        return row

    new_rows = [cell_row_(p) for p in products]

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
    alert_bg = {"red": 1.0, "green": 0.95, "blue": 0.6}
    # new_rows は keep の後ろ。products[i] と new_rows[i] が対応
    prod_by_key = {p["key"]: p for p in products}
    for i, row in enumerate(out[1:], start=1):
        dec = row[4] if len(row) > 4 else ""
        if dec in ("見送り", "見積なし"):
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
        key = str(row[2]) if len(row) > 2 else ""
        p = prod_by_key.get(key)
        if p and (p.get("alert_match") or p.get("alert_fee")):
            # 可否・判断（E/F）をアラート色。列ずれ時は E のみ
            for c0 in (4, 5):
                reqs.append(
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sid,
                                "startRowIndex": i,
                                "endRowIndex": i + 1,
                                "startColumnIndex": c0,
                                "endColumnIndex": c0 + 1,
                            },
                            "cell": {
                                "userEnteredFormat": {"backgroundColor": alert_bg}
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

