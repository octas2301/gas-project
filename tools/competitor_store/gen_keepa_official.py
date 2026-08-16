# -*- coding: utf-8 -*-
"""docs/org/competitor_fields/keepa_official.csv を Keepa product 辞書から生成。"""
from __future__ import annotations

import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "purchase_research_path3"))
from keepa_field_coverage import API_ONLY, MAP  # noqa: E402

OUT = Path(__file__).resolve().parents[2] / "docs" / "org" / "competitor_fields" / "keepa_official.csv"

HEADERS = [
    "api",
    "scope",
    "field",
    "name_ja",
    "agent_pick",
    "priority",
    "effect_setcount",
    "effect_price",
    "effect_human",
    "flow_tags",
    "note",
    "in_logical_crosswalk",
    "keepa_full_column",
]

# 列化する Keepa キー（◎＋K）。csv[] は絶対に列にしない。
FULL_COL = {
    "asin": "ASIN",
    "ean / eanList": "商品コード: EAN",
    "title": "商品名",
    "imagesCSV / image": "画像",
    "manufacturer": "製造者",
    "brand": "ブランド",
    "parentAsin / variations / variationCSV": "親ASIN",
    "https://www.amazon.co.jp/dp/{asin}": "URL: Amazon",
    "https://keepa.com/#!product/5-{asin}": "URL: Keepa",
    "stats.current[0] / csv[0]": "Amazon: 現在価格",
    "stats.avg30[18]": "Buy Box: 30 日平均",
    "stats.avg90[18]": "Buy Box: 90 日平均",
    "stats.avg90[0]": "Amazon: 90 日平均",
    "stats.avg90[1]": "新品: 90 日平均",
    "stats.avg90[4]": "参考価格: 90 日平均",
    "stats.avg90[3]": "売れ筋ランキング: 90 日平均",
    "monthlySold": "月間売上",
    "csv[16] 直近 /10 または rating": "レビュー: 評価",
    "csv[17] / reviewCount": "レビュー: 評価件数",
    "releaseDate": "発売日",
    "categoryTree[0]": "カテゴリ: ルート",
    "categoryTree 結合": "カテゴリ: ツリー",
    "numberOfItems": "アイテム数",
    "packageQuantity": "パッケージ数量",
    "packageLength /10": "梱包_L_cm",
    "packageWidth /10": "梱包_W_cm",
    "packageHeight /10": "梱包_H_cm",
    "packageWeight": "梱包_重量_g",
    "fbaFees": "FBA手数料",
    "offers[].sellerName (isBuyBox)": "BuyBoxセラー",
    "offers[].isFBA (isBuyBox)": "BuyBox_FBA",
}

LOGICAL = {
    "asin": "Y",
    "ean / eanList": "Y",
    "title": "Y",
    "imagesCSV / image": "Y",
    "manufacturer": "Y",
    "brand": "Y",
    "parentAsin / variations / variationCSV": "Y",
    "numberOfItems": "Y",
    "packageLength /10": "Y",
    "packageWidth /10": "Y",
    "packageHeight /10": "Y",
    "packageWeight": "Y",
    "fbaFees": "Y",
    "monthlySold": "Y",
}


def _is_csv_history(path: str, name_ja: str) -> bool:
    if "csv[*]" in path or "全履歴" in name_ja or "全履歴" in path:
        return True
    if "履歴" in path and "csv[" in path:
        return True
    if path.strip().startswith("csv[") and "直近" not in path and "review" not in path.lower():
        return True
    return False


def pick_for(name_ja: str, api: str, path: str) -> tuple[str, str, str]:
    """agent_pick, priority, keepa_full_column Y/N."""
    if _is_csv_history(path, name_ja):
        return "×今回使わない", "P3", "N"
    if api == "derive" and "履歴" in path:
        return "×今回使わない", "P3", "N"
    if path in FULL_COL or name_ja in (
        "ASIN",
        "商品名",
        "商品コード: EAN",
        "画像",
        "製造者",
        "ブランド",
    ):
        if name_ja in ("ASIN",) or path == "asin":
            return "K Keepa専用", "P0", "Y"
        if "package" in path.lower() or name_ja.startswith("パッケージ"):
            return "K Keepa専用", "P0", "Y"
        if "売れ筋" in name_ja or path.startswith("stats.avg90[3]"):
            return "K Keepa専用", "P1", "Y"
        if path == "fbaFees":
            return "K Keepa専用", "P1", "Y"
        if "numberOfItems" in path or name_ja == "アイテム数":
            return "○補助", "P1", "Y"
        if name_ja.startswith("カテゴリ"):
            return "△後続", "P2", "Y"
        if name_ja == "ブランド":
            return "△後続", "P2", "Y"
        if name_ja == "発売日":
            return "△後続", "P2", "Y"
        if "親" in name_ja or "parentAsin" in path:
            return "○補助", "P1", "Y"
        return "◎使えそう", "P0", "Y"
    if api == "hard":
        return "△後続", "P2", "N"
    if api == "offers":
        if path in FULL_COL:
            return "◎使えそう", "P0", "Y"
        return "○補助", "P1", "N"
    if api == "derive":
        return "△後続", "P2", "N"
    return "○補助", "P1", "N"


def main() -> None:
    rows = []
    seen = set()

    def add(scope, field, name_ja, api, note):
        key = (field, name_ja)
        if key in seen:
            return
        seen.add(key)
        pick, pri, full = pick_for(name_ja, api, field)
        if _is_csv_history(field, name_ja):
            pick, pri, full = "×今回使わない", "P3", "N"
            note = (note + " csv[]は倉庫非保存。").strip()
        rows.append(
            [
                "Keepa product",
                scope,
                field,
                name_ja,
                pick,
                pri,
                "高" if "アイテム" in name_ja else ("中" if pick.startswith("◎") else "無"),
                "高" if "価格" in name_ja or "Buy Box" in name_ja else "低",
                "中",
                "A_Keepa,FBA" if pick.startswith("K") else "A_Keepa",
                note,
                LOGICAL.get(field, "N"),
                full,
            ]
        )

    for name_ja, api, path, note in MAP:
        add("product", path, name_ja, api, note)
    for name, note in API_ONLY:
        add("api_only", name, name, "yes", note)

    add("warehouse", "fetchedAt", "取得日時", "yes", "Keepaフルメタ")
    add("warehouse", "stats.current[18]", "Buy Box: 現在価格", "yes", "④のAmazon価格。csv非使用")
    add("warehouse", "stats.avg30[0]", "Amazon: 30 日平均", "yes", "")
    add("warehouse", "priceFingerprint", "価格指紋", "yes", "変わったときだけ追記判定")
    add("warehouse", "rawJson", "生JSON", "yes", "csv[]を落とした product")

    for i, rec in enumerate(rows):
        if rec[3] in ("取得日時", "価格指紋", "生JSON", "Buy Box: 現在価格"):
            rec[4] = "◎使えそう"
            rec[5] = "P0"
            rec[12] = "Y"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADERS)
        w.writerows(rows)
    print("wrote", OUT, "n=", len(rows))


if __name__ == "__main__":
    main()
