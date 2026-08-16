# 競合フィールド正本（楽天・Yahoo・Keepa）

**文書種別**: 正本（フィールド辞書＋参照ルール）  
**最終更新**: 2026-08-15  
**状態**: 運用中（都度改修）。コード実装は別承認。  
**保存・段階導入**: [COMPETITOR_STORE_REQUIREMENTS.md](COMPETITOR_STORE_REQUIREMENTS.md)  
**作業コピー（色付きExcel）**: `docs/org/competitor_fields/competitor_fields_workbook.xlsx`  
**再生成**: `python tools/purchase_research_path3/_gen_crosswalk_xlsx.py`

Git で差分を追う正は **CSV**（下）。Excel は閲覧用。

| ファイル | 中身 |
|----------|------|
| [competitor_fields/logical_crosswalk.csv](competitor_fields/logical_crosswalk.csv) | 論理列（同じ列にできるか）＋印・優先度・効果・フロータグ |
| [competitor_fields/rakuten_official.csv](competitor_fields/rakuten_official.csv) | 楽天 Ichiba Item Search **出力の全項目** |
| [competitor_fields/yahoo_official.csv](competitor_fields/yahoo_official.csv) | Yahoo 商品検索 v3 **hits 配下の全項目** |
| [competitor_fields/keepa_official.csv](competitor_fields/keepa_official.csv) | Keepa **product** 辞書。`csv[]` は×／列化しない。`Keepaフル` は◎＋K＋生JSON |

公式出典: [楽天 Ichiba Search](https://webservice.rakuten.co.jp/documentation/ichiba-item-search) / [Yahoo itemSearch v3](https://developer.yahoo.co.jp/webapi/shopping/v3/itemsearch.html) / Keepa product（材料: [石原水産_Keepa項目過不足.csv](../../tools/purchase_research_path3/石原水産_Keepa項目過不足.csv)）

---

## 0. 共通ルール（どのPJでも）

1. **競合の価格・袋数・実質を設計するときは、先に本正本を開く。** コードやベイクオフCSVだけを正にしない。
2. **1件目の価格を競合価格にしない。** 袋数クラスタ・除外（ふるさと納税・選択式・中古）のあと。
3. **カテゴリID・ブランドID・ASIN をモール横断で1列にマージしない。** JAN/EAN だけが JOIN キー（楽天ヒット行に JAN 欄は無い → 検索に使った JAN で紐づける）。
4. **送料の円は3モールとも検索APIではほぼ取れない。** フラグ／`shipping.code` のみ。**楽天ポイントは pointRate＝％（ポイント数ではない）。** Yahoo は単位未確定（旧 `point.amount` は 0 固定、現行は `lyLimitedBonus*` の可能性）。同一列の生値を足さない。
5. **Keepa `numberOfItems` を袋数の正にしない。** 内容量（4P）になりやすい。楽天・Yahooタイトルの袋数＋単価を優先。
6. **マスタの評価列・人間◎は機械で消さない**（Amazon貼付要件と同じ思想）。
7. フィールドをコードで新規に使う／捨てる／意味が変わったら **CSV を先に直し、本ファイル §改修履歴に1行**。実装承認は別。

印: `◎使えそう`＝raw保存して後続で再利用 / `○補助` / `△後続`（カテゴリ・CPO） / `×使わない` / `K` Keepa専用。  
優先度 `P0`〜`P3`。効果 `effect_setcount` / `effect_price` / `effect_human`。フローは `flow_tags` と Y/N 列。

---

## 1. 場面 → 開く正本（組み立てルール）

競合情報を触る作業は、**本正本＋下表のフロー正本**の両方を読む。フィールドの意味は本正本、処理順・書込列・聖域はフロー正本。

| 場面（PJ） | フロー正本（処理・列・禁止） | 本正本の見方 |
|------------|------------------------------|--------------|
| **A Keepa・セット数** | [Keepa分析_モール横断前倒し](../Keepa分析_モール横断前倒し_要件定義.md) | `flow_A_Keepa=Y` かつ P0。袋数＝タイトル＋単価アンカー |
| **B Step2 モール横断** | [RAKUTEN_YAHOO_COMPETITIVE_PRICE](../RAKUTEN_YAHOO_COMPETITIVE_PRICE_REQUIREMENTS.md) | `flow_B_Step2`。キャッシュ再利用 |
| **B Step7.6 カテゴリ** | [YAHOO_CATEGORY_BRAND_STAGE](../YAHOO_CATEGORY_BRAND_STAGE.md)／楽天 Nav | `flow_B_Step76`。IDはマージしない。売れ筋は別クエリ |
| **⑤ CPO・価格対抗** | [CPO_PRICING](../CPO_PRICING.md)／[PRICING_V1](../PRICING_V1_REQUIREMENTS.md) | `flow_CPO`＋実質価格。1件目禁止 |
| **定時・日次の競合チェック** | [RUNBOOK_DAY_WEEK_MONTH](../RUNBOOK_DAY_WEEK_MONTH.md) | 同じ検索rawを使い回す。再検索は TTL 切れのみ |
| **領域1 仕入れリサーチ** | [DOMAIN1](../DOMAIN1_RESEARCH_PURCHASING.md)／[NOW](B_PURCHASE_RESEARCH_NOW.md)／[TASK_STRUCTURE](B_PURCHASE_RESEARCH_TASK_STRUCTURE.md) | JAN JOIN。KeepaはKeepaフル。洗いCatalogはモールヒットにしない |
| **Amazon 競合貼付** | [B_AMAZON_COMPETITOR_PASTE](B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md) | 出品②のその商品検索。Catalog階段。Keepaはスナップショット（薄い）＋`Keepaフル`。◎は人間 |
| **ベイクオフ／検索診断** | シート「競合検索ベイクオフ」／`RakutenYahooBakeoff.js` | A〜Eは診断。本線に5クエリ全部をA前で回さない。Amazon A–E も診断専用 |

サブGAS（領域1複製スプシ等）でも **フィールド意味はこの正本**。clasp の scriptId は [gas-multi-project](../../.cursor/rules/gas-multi-project.mdc) どおり混ぜない。

---

## 2. 都度改修の手順

1. 公式ドキュメントで項目追加・0固定・改名を確認する。  
2. `competitor_fields/*.csv` または `_gen_crosswalk_xlsx.py` のリストを直す（公式全項目はスクリプト側がマスターの場合、CSVは再生成で上書き）。  
3. `python tools/purchase_research_path3/_gen_crosswalk_xlsx.py`  
4. 本ファイルの **§改修履歴** と [CHANGE_LEDGER](../CHANGE_LEDGER.md) に1行。  
5. コードが薄いパースのままなら、CSVの◎は「raw保存すれば使える」であり、実装済みではない。

**今のコードがパースしているもの（2026-08-15c）**  
モールヒットは **◎＋分析用○を列化**（中古・キャッチは生JSON）。△と×は列にしない。  
**Keepaスナップショット**＝運用の薄い列（マスタキャッシュと同系統）。辞書どおりの倉庫は **`Keepaフル`＋生JSON**（Z併存）。`csv[]` は保存しない。  
ポイントは **楽天％／楽天還元円／Yahoo数／Yahoo倍率** に分離。

---

## 3. 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-08-16 | B Step2.1＝12-⑭。ストアON時は横断シート直書きをスキップ。15〜23はB外。 |
| 2026-08-16 | F2c: サブ画像を画像の右隣へ。K8: BB現在/30/90・レビュー・BuyBox_FBAは今JSON空。取得/代替は未実施メモ |
| 2026-08-16 | F2b: 画像=メインHYPERLINK。サブ画像=2枚目以降を `|` で結合（splitしやすい）。旧画像一覧は列名変更 |
| 2026-08-15 | F2 画像・画像一覧をimages[]から列化。BB価格・レビューはJSON空 |
| 2026-08-15 | F1: カテゴリ・梱包・Keepa FBA手数料・BBセラーをフル列化。BB現在価格はJSON空（offers/csvなし） |
| 2026-08-15 | L1: 出品FBAティア／手数料・自己発送は設定マスタfirst-fit。Keepa FBA手数料列は別 |
| 2026-08-15 | F0列化: Amazon直販・新品現在・現行順位・180日在庫切れ%。BB不使用 |
| 2026-08-15 | 課題K1–K7（offers本線禁止・BB不要・標2bは出品FBA流用）。列追加なし |
| 2026-08-15 | 直販＝availabilityAmazon（0/-1）。メーカー出品照合はoffers[]将来。L1は梱包first-fit。列追加なし |
| 2026-08-15 | 出品者数＝Keepa COUNT_NEW（current[11]）。クイックショップ一致。offersユニークは使わない。列化済 |
| 2026-08-15 | リサーチKeepaの正はKeepaフル。①候補は門の転記。フィールド変更なし |
| 2026-08-15 | 領域1進捗の正は [B_PURCHASE_RESEARCH_NOW.md](B_PURCHASE_RESEARCH_NOW.md)。フィールド変更なし |
| 2026-08-15 | P0: `keepa_official.csv`。Keepaフル列＝◎＋K＋生JSON。csv[]非保存 |
| 2026-08-15 | 12-⑭：専用ヒットをセット数クラスタしてマスタ競合列へ。1件目禁止。 |
| 2026-08-15 | A+C: 計N袋優先・各N袋除外。単価2倍外れは12-⑭で捨て。課題 [COMPETITOR_SETCOUNT_ACCURACY.md](COMPETITOR_SETCOUNT_ACCURACY.md)。 |
| 2026-08-15 | 自店 Octas をモールヒット・Step2から除外（店名・shopId。商品名では落とさない）。 |
| 2026-08-15 | モールヒット「ポイント」＝楽天％。Yahoo単位は未確定。列は辞書全項目より意図的に少ない（論理キー＋生JSON）。 |
