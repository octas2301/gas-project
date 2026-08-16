# -*- coding: utf-8 -*-
"""競合ストアのシート名・ヘッダー（日本語）。Keepaスナップショット列は出品キャッシュと揃える。"""

SHEET_FIELD_MAP = "項目マップ"
SHEET_HITS = "モールヒット"
SHEET_KEEPA = "Keepaスナップショット"
SHEET_KEEPA_FULL = "Keepaフル"
SHEET_META = "運用メタ"
SHEET_MAKER = "メーカーマスタ"
SHEET_SELLER = "セラー"
SHEETS = (
    SHEET_FIELD_MAP,
    SHEET_HITS,
    SHEET_KEEPA,
    SHEET_KEEPA_FULL,
    SHEET_META,
    SHEET_MAKER,
    SHEET_SELLER,
)

MAKER_HEADERS = [
    "メーカー",
    "第2クエリ語",
    "採取元",
    "公式URL",
    "取得日時",
    "再採取しない",
    "メモ",
]

SELLER_HEADERS = [
    "sellerId",
    "店名",
    "対象カテゴリ",
    "巡回日",
    "asinList件数",
    "抽出メーカー",
    "ピック",
    "採取元",
    "メモ",
    "カテゴリ構成",
    "食品比率",
    "ストアASIN数",
    "卸仮説",
    "メインカテゴリ",
]

# Amazon ルート相当（人が渡した一覧。最新公式と差があれば列名だけ直す）
AMAZON_SELLER_CAT_COLS = [
    "洋書",
    "本",
    "ミュージック",
    "ファッション",
    "DVD",
    "クラシック",
    "ホーム&キッチン",
    "PCソフト",
    "TVゲーム",
    "文房具・オフィス用品",
    "家電&カメラ",
    "ドラッグストア",
    "ビューティー",
    "DIY・工具・ガーデン",
    "おもちゃ",
    "産業・研究開発用品",
    "スポーツ&アウトドア",
    "パソコン・周辺機器",
    "車＆バイク",
    "ホビー",
    "ベビー&マタニティ",
    "ペット用品",
    "楽器",
    "大型家電",
    "食品・飲料・お酒",
    "デジタルミュージック",
    "Kindleストア",
    "不明",
]

SELLER_HEADERS = SELLER_HEADERS + AMAZON_SELLER_CAT_COLS

PURPOSE_RESEARCH = "リサーチ"
PURPOSE_SCHEDULED = "定時"
PURPOSE_LISTING = "出品"

FIELD_MAP_HEADERS = [
    "マスタ候補列名",
    "論理名",
    "Amazon",
    "楽天",
    "Yahoo",
    "適用開始日",
    "変換メモ",
    "仕分け",
    "優先度",
    "取得先",
    "ソースAPI",
    "フィールド",
]

KEEP_PICKS = ("◎使えそう", "○補助", "△後続", "K Keepa専用")

HITS_HEADERS = [
    "取得日時",
    "目的",
    "モール",
    "検索JAN",
    "YahooヒットJAN",
    "ASIN",
    "店名",
    "店舗コード",
    "店舗URL",
    "店商品コード",
    "モール商品コード",
    "商品名",
    "表示価格",
    "税抜価格",
    "楽天税フラグ",
    "Yahoo税込",
    "楽天価格min購入可",
    "セール価格",
    "送料フラグ",
    "送料条件名",
    "楽天ポイント％",
    "楽天還元円",
    "楽天ポイント開始",
    "楽天ポイント終了",
    "Yahooポイント数",
    "Yahooポイント倍率",
    "在庫",
    "レビュー件数",
    "画像有無",
    "画像URL",
    "画像URL小",
    "画像ID",
    "商品URL",
    "説明文",
    "ブランド名",
    "ヒット順位",
    "クエリ",
    "マップ版",
    "生JSON",
    "競合確定価格",
]

KEEPA_SNAP_HEADERS = [
    "取得日時",
    "画像",
    "商品名",
    "ASIN",
    "商品コード: EAN",
    "製造者",
    "ブランド",
    "アイテム数",
    "発売日",
    "URL: Amazon",
    "URL: Keepa",
    "カテゴリ: ルート",
    "カテゴリ: ツリー",
    "売れ筋ランキング: 90 日平均",
    "レビュー: 評価",
    "レビュー: 評価件数",
    "Buy Box: 現在価格",
    "Buy Box: 30 日平均",
    "Buy Box: 90 日平均",
    "Amazon: 現在価格",
    "Amazon: 30 日平均",
    "Amazon: 90 日平均",
    "新品: 90 日平均",
    "参考価格: 90 日平均",
    "セット数",
    "タイトル由来セット数",
    "セット数理由",
    "梱包_L_cm",
    "梱包_W_cm",
    "梱包_H_cm",
    "梱包_重量_g",
    "梱包確認済",
]

# 倉庫。◎＋K＋メタ＋生JSON。csv[] は列にも生JSONにも置かない（書込時に落とす）
KEEPA_FULL_HEADERS = [
    "取得日時",
    "目的",
    "ASIN",
    "商品コード: EAN",
    "商品名",
    "画像",
    "サブ画像",
    "製造者",
    "ブランド",
    "親ASIN",
    "URL: Amazon",
    "URL: Keepa",
    "Buy Box: 現在価格",
    "Buy Box: 30 日平均",
    "Buy Box: 90 日平均",
    "Amazon: 現在価格",
    "Amazon: 30 日平均",
    "Amazon: 90 日平均",
    "新品: 90 日平均",
    "参考価格: 90 日平均",
    "売れ筋ランキング: 90 日平均",
    "月間売上",
    "出品者数",
    "Amazon直販",
    "新品: 現在価格",
    "売れ筋ランキング: 現在",
    "Amazon: 180日在庫切れ%",
    "レビュー: 評価",
    "レビュー: 評価件数",
    "発売日",
    "カテゴリ: ルート",
    "カテゴリ: ツリー",
    "アイテム数",
    "パッケージ数量",
    "梱包_L_cm",
    "梱包_W_cm",
    "梱包_H_cm",
    "梱包_重量_g",
    "FBA手数料",
    "BuyBoxセラー",
    "BuyBox_FBA",
    "出品FBAティア",
    "出品FBA手数料",
    "自己発送サイズ",
    "自己発送送料",
    "梱包3辺合計_cm",
    "価格指紋",
    "生JSON",
]

# マスタ Keepa キャッシュ → 専用スナップショット
KEEPA_HEADER_ALIASES = {
    "setCount": "セット数",
    "setCountFromTitle": "タイトル由来セット数",
    "setCountReason": "セット数理由",
    "梱包_checked": "梱包確認済",
}

META_HEADERS = ["キー", "値"]

MASTER_SS_ID = "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"
MASTER_KEEPA_SHEET = "Keepa取得_キャッシュ"
