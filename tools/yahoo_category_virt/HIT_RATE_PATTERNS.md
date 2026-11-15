# Yahoo itemSearch ヒット率 — クエリパターン表

**目的**: 画面ではヒットするのに API が 0 件になる問題を、**検索語パターン別に** `totalResultsAvailable` で比較する。  
**対象 API**: ShoppingWebService V3 `itemSearch`（`sort=-score`）。読取のみ。  
**実行**: `python tools/yahoo_category_virt/hit_rate_yahoo_queries.py`

---

## パターン一覧（固定 ID）

商品入力: `maker` / `name`（商品名ベース） / `jan` / 任意 `mainKw` / 任意 `kindN`（コンボ種数） / 任意 `seriesCore`（省略時は自動推定）

| ID | ラベル | 生成ルール | 現行GASに近い |
|----|--------|------------|----------------|
| P01_JAN | JAN（jan_code） | `jan_code=` | ○ 単品先頭 |
| P02_MAKER_NAME | メーカー＋商品名 | `{maker} {name}` | ○ |
| P03_NAME | 商品名のみ | `{name}` | ○ |
| P04_SIMPLE | 簡易名 | 括弧・中点除去＋空白正規化 | ○ |
| P05_MAKER_SIMPLE | メーカー＋簡易名 | `{maker} {simple}` | ○ |
| P06_SERIES | シリーズ核のみ | `seriesCore`（例: 食べるおだし） | 提案 |
| P07_MAKER_SERIES | メーカー＋シリーズ | `{maker} {seriesCore}` | 提案 |
| P08_KIND | 簡易＋N種 | `{simple} {kindN}種` | ○ コンボ先頭付近 |
| P09_MAKER_KIND | メーカー＋簡易＋N種 | `{maker} {simple} {kindN}種` | ○ コンボ先頭 |
| P10_MAIN_KW | メインKW | `{mainKw}` | ○ あれば |
| P11_NORM_MS | 半角正規化メーカー簡易 | P05 の全角英数・全角空白を半角化 | 提案 |
| P12_MAKER_SERIES_FLAVOR | メーカー＋シリーズ＋味 | `{maker} {seriesCore} {flavor}`（括弧内先頭トークン） | 提案 |

`kindN` が無い／1 以下のとき P08/P09 はスキップ。  
`jan` が無いとき P01 はスキップ。  
`mainKw` が無いとき P10 はスキップ。

---

## 指標

| 列 | 意味 |
|----|------|
| total | `totalResultsAvailable` |
| hit_bin | total≥1 なら 1 |
| returned | 返却 `hits` 件数 |
| maker_top5 | 先頭5件タイトルにメーカー文字列を含む件数 |
| series_top5 | 先頭5件にシリーズ核を含む件数 |
| pet_top5 | 先頭5件にペットノイズ（カルカン等）を含む件数 |
| top1_title | 1件目タイトル（先頭60字） |

集計: パターン ID ごとの **hit_bin 平均（ヒット率）** / total 中央値 / maker_top5 平均。

---

## サンプル商品

`products.sample.json`（行20／47／110相当）。

---

## 認証

1. `tools/yahoo_category_virt/config.local.json`（`config.example.json` をコピー）  
2. 無ければ `tools/purchase_research_path3/config.local.json`  
3. 環境変数 `YAHOO_SHOPPING_CLIENT_ID`

レート: 既定 **1.25 秒／クエリ**（Yahoo 目安 1q/s）。429 時は最大2回リトライ。
