#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ユーザーKeepaエクスポート列 vs product API の過不足表。API再取得しない。"""

from __future__ import annotations

import csv
from pathlib import Path

# api: yes=そのまま / derive=計算可 / offers=offers=20 / build=URL組み立て / hard=公式でも薄い / no=取れない
MAP = [
    ("画像", "yes", "imagesCSV / image", "トークンからURL組み立て"),
    ("商品名", "yes", "title", ""),
    ("商品ハイライト", "yes", "features", "配列。空のことが多い"),
    ("売れ筋ランキング: 90 日平均", "yes", "stats.avg90[3]", "stats指定が必要"),
    ("売れ筋ランキング: 365 日平均", "yes", "stats.avg365[3]", "stats=365"),
    ("売れ筋ランキング: 90日間の下落 %", "derive", "avg90 と現在の比", "CSVの既算値。APIは自分で割る"),
    ("売れ筋ランキング: 過去30日間の減少", "derive", "csv[3] 履歴", "履歴から差分。statsだけだと無し"),
    ("売れ筋ランキング: 過去90日間の減少", "derive", "csv[3] 履歴", ""),
    ("売れ筋ランキング: 過去180日間の減少", "derive", "csv[3] 履歴", ""),
    ("売れ筋ランキング: 過去365日間の減少", "derive", "csv[3] 履歴", ""),
    ("月間売上トレンド: 先月の購入", "yes", "monthlySold", "無いASINもある"),
    ("レビュー: 評価", "yes", "csv[16] 直近 /10 または rating", "43→4.3"),
    ("レビュー: 評価件数", "yes", "csv[17] / reviewCount", ""),
    ("Buy Box: 30 日平均", "yes", "stats.avg30[18]", "BUY_BOX_SHIPPING"),
    ("Buy Box: 90 日平均", "yes", "stats.avg90[18]", ""),
    ("Buy Box: 90日間の下落 %", "derive", "stats平均の比", ""),
    ("Buy Box: Buy Box セラー", "offers", "offers[].sellerName (isBuyBox)", "offers=20 が必要。名前はCSV専用に近い"),
    ("Buy Box: FBAです", "offers", "offers[].isFBA (isBuyBox)", ""),
    ("Amazon: 現在価格", "yes", "stats.current[0] / csv[0]", "空＝本体なし"),
    ("Amazon: 90 日平均", "yes", "stats.avg90[0]", ""),
    ("Amazon: 365 日平均", "yes", "stats.avg365[0]", ""),
    ("Amazon: 90日間の下落 %", "derive", "平均の比", ""),
    ("Amazon: 在庫", "hard", "amazonオファー stock", "2品番CSVも空。APIも薄い"),
    ("Amazon: 90日間在庫切れ", "yes", "stats.outOfStockPercentage90[0]", "100 と 100% の表記差"),
    ("Amazon: 在庫切れカウント 30 日間", "derive", "csv[0] の -1 連続", "既算値はCSV寄り"),
    ("Amazon: 在庫切れカウント 90 日間", "derive", "csv[0] 履歴", "門の「6ヶ月ほぼ不在」は履歴で可"),
    ("新品: 90 日平均", "yes", "stats.avg90[1]", "円の整数"),
    ("新品: 365 日平均", "yes", "stats.avg365[1]", ""),
    ("新品: 90日間の下落 %", "derive", "平均の比", ""),
    ("新しい、第三者FBA: 90 日平均", "yes", "stats.avg90[10]", "NEW_FBA"),
    ("新しい、第三者FBA: 365 日平均", "yes", "stats.avg365[10]", ""),
    ("新しい、第三者FBA: 90日間の下落 %", "derive", "平均の比", ""),
    ("最安の FBA セラー", "offers", "offers isFBA 最安の sellerName", "名前はoffers"),
    ("新品 第三者 FBM: 90 日平均", "yes", "stats.avg90[7]", "NEW_FBM_SHIPPING 送料込"),
    ("新品 第三者 FBM: 365 日平均", "yes", "stats.avg365[7]", ""),
    ("新品 第三者 FBM: 90日間の下落 %", "derive", "平均の比", ""),
    ("最安の FBM セラー", "offers", "offers isFBA=false 最安", ""),
    ("参考価格: 90 日平均", "yes", "stats.avg90[4]", ""),
    ("参考価格: 365 日平均", "yes", "stats.avg365[4]", ""),
    ("参考価格: 90日間の下落 %", "derive", "平均の比", ""),
    ("新品アイテム数: 90 日平均", "yes", "stats.avg90[11]", "COUNT_NEW"),
    ("新品アイテム数: 365 日平均", "yes", "stats.avg365[11]", ""),
    ("新品アイテム数: 90日間の下落 %", "derive", "平均の比", ""),
    ("URL: Amazon", "build", "https://www.amazon.co.jp/dp/{asin}", ""),
    ("URL: Keepa", "build", "https://keepa.com/#!product/5-{asin}", ""),
    ("カテゴリ: ルート", "yes", "categoryTree[0]", ""),
    ("カテゴリ: サブ", "yes", "categoryTree 末尾", ""),
    ("カテゴリ: ツリー", "yes", "categoryTree 結合", ""),
    ("ASIN", "yes", "asin", ""),
    ("商品コード: EAN", "yes", "ean / eanList", "前回フラット化漏れ。フィールドはある"),
    ("製造者", "yes", "manufacturer", ""),
    ("ブランド", "yes", "brand", ""),
    ("単位の詳細: 単位の価値", "yes", "unitCount 等", "ASINによる"),
    ("アイテム数", "yes", "numberOfItems", ""),
    ("発売日", "yes", "releaseDate", "Keepa分"),
    ("パッケージ: 長さ (cm)", "yes", "packageLength /10", "APIはmm"),
    ("パッケージ: 幅 (cm)", "yes", "packageWidth /10", ""),
    ("パッケージ: 高さ (cm)", "yes", "packageHeight /10", ""),
    ("パッケージ: 重さ (g)", "yes", "packageWeight", ""),
    ("パッケージ: 数量", "yes", "packageQuantity", ""),
    ("商品: 長さ (cm)", "yes", "itemLength /10", ""),
    ("商品: 幅 (cm)", "yes", "itemWidth /10", ""),
    ("商品: 高さ (cm)", "yes", "itemHeight /10", ""),
    ("商品: 重さ (g)", "yes", "itemWeight", ""),
    ("Buy Box: 定期おトク便", "hard", "sns / csv SNS系", "コード.jsもAPI要確認"),
    ("ワンタイムクーポン: 定期おトク便 %", "hard", "coupon", "クーポンオブジェクト。常時あるわけではない"),
    ("ビジネス割引: パーセンテージ", "hard", "businessDiscount", "無いことが多い"),
    ("Buy Box: % Amazon 30 日", "yes", "stats.buyBoxUsed/amazon 系 or csv", "statsに比率あり"),
    ("Buy Box: % Amazon 90 日", "yes", "stats", ""),
    ("Buy Box: % Amazon 180 日", "yes", "stats", "stats期間を長く"),
    ("Buy Box: % Amazon 365 日", "yes", "stats", ""),
    ("Buy Box: % トップセラー 30 日", "yes", "stats", ""),
    ("Buy Box: % トップセラー 90 日", "yes", "stats", ""),
    ("Buy Box: % トップセラー 180 日", "yes", "stats", ""),
    ("Buy Box: % トップセラー 365 日", "yes", "stats", ""),
    ("Buy Box: 勝者数 30 日", "yes", "stats", ""),
    ("Buy Box: 勝者数 90 日", "yes", "stats", ""),
    ("Buy Box: 勝者数 180 日", "yes", "stats", ""),
    ("Buy Box: 勝者数 365 日", "yes", "stats", ""),
    ("Buy Box: 標準偏差 30 日", "yes", "stats", ""),
    ("Buy Box: 標準偏差 90 日", "yes", "stats", ""),
    ("Buy Box: 標準偏差 365 日", "yes", "stats", ""),
    ("Buy Box: 変動性 30 日", "yes", "stats", "flipability"),
    ("Buy Box: 変動性 90 日", "yes", "stats", ""),
    ("Buy Box: 変動性 365 日", "yes", "stats", ""),
]

API_ONLY = [
    ("availabilityAmazon", "本体が今いるか。-1=なし。門の本体いま"),
    ("csv[*] 全履歴", "6ヶ月本体不在はここから。エクスポートの「カウント」より生データ"),
    ("offers[]", "いまの出品者・BuyBox・FBA。offers=20 でトークン増"),
    ("parentAsin / variations / variationCSV", "バリエーション。エクスポートに無い"),
    ("salesRankReference / salesRanks", "カテゴリ別順位"),
    ("fbaFees", "FBA手数料目安"),
    ("listedSince / lastUpdate / trackingSince", "追跡開始・更新"),
    ("eanList / upcList", "JAN複数"),
    ("partNumber", "品番。出品②で使用"),
    ("parentTitle 等", "親ASIN情報"),
]


def main() -> None:
    src = Path(__file__).resolve().parent / "ishihara_keepa_2.csv"
    dest = Path(__file__).resolve().parent / "石原水産_Keepa項目過不足.csv"
    with src.open(encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        cols = list(r.fieldnames or [])
        rows = list(r)
    by = {m[0]: m for m in MAP}
    with dest.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["side", "column", "api", "api_path", "export_filled_2asin", "note"])
        for c in cols:
            filled = sum(1 for rec in rows if str(rec.get(c) or "").strip() not in ("", "-"))
            if c in by:
                _, api, path, note = by[c]
                w.writerow(["export", c, api, path, "%s/2" % filled, note])
            else:
                w.writerow(["export", c, "unmapped", "", "%s/2" % filled, "対応表に無い列名"])
        for name, note in API_ONLY:
            w.writerow(["api_only", name, "yes", name, "", note])
    print("cols", len(cols), "mapped", sum(1 for c in cols if c in by), "out", dest)


if __name__ == "__main__":
    main()
