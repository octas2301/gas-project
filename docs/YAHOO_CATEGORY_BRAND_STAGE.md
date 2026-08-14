# Yahooカテゴリ／ブランド — Stage 要件（都度API・マスタ書込）

**文書種別**: requirements（実装前の正）  
**最終更新**: 2026-08-12  
**状態**: **実装済**（SHP 正本検証＋**7.5 price_aware／7.6 popular_only**）  
**承認パッケージ**: [org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md](org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md)  
**手順**: [org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md](org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md)  
**親**: メニュー8（[org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)）／楽天同型 [RAKUTEN_NAV_GENRE_STAGE3.md](RAKUTEN_NAV_GENRE_STAGE3.md)
**E確認**: [org/D_MENU_E_GENRE_YAHOO_REQUIREMENTS_CONFIRM.md](org/D_MENU_E_GENRE_YAHOO_REQUIREMENTS_CONFIRM.md)

---

## 0. 方針（確定・社長確認済）

| 項目 | 内容 |
|------|------|
| **プロダクトカテゴリの正本** | Circus **`getShopCategoryList`**（SHP／ストアと同じ系統）。書込前に必ず検証 |
| 候補ソース | Yahoo!ショッピング商品検索 v3（売れ筋＋最安）／競合ゼロ時は AI列・Drive CSV（**候補のみ**） |
| AIの★推奨Yahooカテゴリ／ブランド | **正本にしない**（候補・フォールバックのみ） |
| 書込先（親行） | `YahooカテゴリID` / `(Yahooカテゴリ名)` / `Yahooブランドコード` |
| **YahooカテゴリID の選定** | **モード分岐**（下表）。いずれも書込前に **SHP検証OKのみ** |
| `(Yahooカテゴリ名)` | SHP の `PathName` を優先（`＞`/`>` → `:`）。無ければ候補の階層連結 |
| ショップ運用 | プロダクト階層をショップ path に写す前提 |
| 入口 | **Z 7.5**＝`price_aware`／**Z 7.6**＝`popular_only`／**B統合は 7.6**（7.5は差し替え） |
| 聖域 | `generateRakutenCSV`／**Yahoo.js の editItem／画像／在庫送信は非改変**（認証再利用と SHP 読取ヘルパのみ追加可）／B統合境界非改変 |
| トグル | `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED` 未設定＝**ON**（緊急停止のみ false） |
| 認証（SHP） | `▼設定(Yahooマッピング)` の Client／Secret／Refresh／**seller_id(B15)**（出品と同じ） |

---

## 0.1 モード（2026-08-12 確定）

| モード | 入口 | カテゴリ | ブランド |
|--------|------|----------|----------|
| `price_aware` | Z **7.5** | 売れ筋重み＋自社 `Yahoo!価格設定` 最安プローブ（従来） | 市場投票／AI・Drive FB → SHP |
| `popular_only` | Z **7.6**／**B Step7.6** | 売れ筋重み最大のみ（価格・最安プローブなし） | **メーカー名 SHP 優先** → 失敗時 **`38074`（ブランド登録なし）**。AI/Drive ブランド候補は使わない |

定数: `YAHOO_CAT_MODE_PRICE_AWARE_` / `YAHOO_CAT_MODE_POPULAR_ONLY_` / `YAHOO_BRAND_CODE_NO_BRAND_='38074'`。

---

## 1. なぜ都度APIか／正本がSHPか

- Shopping itemSearch の `genreCategory.id` は検索用ジャンルで、**SCのプロダクトカテゴリと一致しないことがある**（実例: `45944` がSCに無く `43065` が正）。
- Drive CSV も更新遅れがあり得る。
- よって **候補出しは市場／AI可**、**マスタへ書く ID／path は `getShopCategoryList` で存在確認できたものだけ**とする。

参照:

- [SHPカテゴリ検索API](https://developer.yahoo.co.jp/webapi/shopping/getShopCategoryList.html)
- [商品検索（v3）](https://developer.yahoo.co.jp/webapi/shopping/v3/itemsearch.html)（候補用）

**API制約（候補側）**: V3 に sold 順は無い → `-review_count` 近似。

---

## 2. 処理フロー（v1.2）

```text
メニュー8（レ点親）・YahooトグルON
  → 【候補】JAN/商品名 → itemSearch sort=-review_count
       price_aware: 重み上位Kを最安プローブ＋自社価格比較
       popular_only: 重み最大のみ（プローブなし）
       price_aware かつ失敗時 → AI列 / Drive CSV（§2.7）
       popular_only かつ失敗時 → カテゴリは要確認（ブランドは38074可）
  → 【正本検証】getShopCategoryList（Yahoo ID連携）… §2.8
  → ブランド:
       price_aware: 市場投票／§2.7 → getShopBrandList
       popular_only: メーカー名 name 検索 → 無ければ 38074
  → 親行へ書込
```

### 2.1〜2.3（候補ルール）

- `price_aware`: 売れ筋＋自社最安／ブランド投票／`:` 連結は候補用。書込 path は SHP PathName 優先  
- `popular_only`: `amazonAiPickYahooCategoryPopularWeightOnly_`（TOP_N 重み最大）

### 2.8 SHP 正本検証（必須）

| 結果 | 挙動 |
|------|------|
| `CategoryCode` が候補IDと一致 | 書込。path は PathName を `:` 化。reason に `shp_validated` |
| ID不一致だが名前検索で1件以上 | **SC側のコードに差し替え**て書込。reason に `shp_resolved`。要確認可 |
| 0件／認証失敗／APIエラー | **カテゴリは書かない**＋要確認（誤IDを残さない） |

レート: 目安 1クエリ/秒。親あたり検証1〜2回＋sleep。

`downloadShopCategories` 全件DLは本 Stage 対象外（将来キャッシュ用）。

### 2.3 ブランド（候補→SHP正本）

| 段階 | 規則 |
|------|------|
| 候補（`price_aware`） | 市場ヒットの `brand.id` 投票、または §2.7 AI/Drive |
| 候補（`popular_only`） | **市場投票・AI/Drive 不使用**。メーカー名のみ |
| **正本** | Circus **`getShopBrandList`** |
| 検証（`price_aware`） | 候補コード `type=code` → 無ければメーカー `type=name` |
| 検証（`popular_only`） | メーカー `type=name` を先に照合 → 失敗／認証失敗時 **`38074`**（`no_brand_fallback`） |
| 失敗（`price_aware`） | ブランド非書込＋要確認 |

### 2.9 SHPブランド正本（必須）

参照: [SHPブランドコード検索API](https://developer.yahoo.co.jp/webapi/shopping/getShopBrandList.html)

認証はカテゴリと同じ（Yahooマッピング B12–B15）。

### 2.7 競合ゼロ時フォールバック（AI推定→API連携マスタ）

市場 itemSearch が **hits空**（または CLIENT_ID 未設定）のとき:

1. **AI推奨列**（`▼マスタ(★推奨YahooカテゴリID)` / `▼マスタ(Yahooカテゴリ名)` / `▼マスタ(★推奨Yahooブランドコード)`）の先頭候補を採用  
2. それでも無いとき **Drive CSV**（`YAHOO_CATEGORY_FOLDER_ID` / `YAHOO_BRAND_FOLDER_ID`＝API連携マスタ）を `searchCandidatesWithScore` で商品名ベース＋メーカー検索し、スコア1位の ID／名称を採用  
3. 名称は `＞`/`>` を `:` に正規化して `(Yahooカテゴリ名)` へ  
4. 理由コード: `ai_recommend_col` / `drive_master_csv`。**要確認フラグを付ける**（市場根拠が無いため）  
5. 1も2も失敗 → 従来どおり要確認・非書込  
6. **フォールバックで得た ID も §2.8 SHP 検証を通す**（通らなければ非書込）

通常パス（競合あり）では AI を正本にしない。フォールバックは候補のみ。書込正本は常に SHP。

### 2.4 AI候補（通常パス）

- 通常の売れ筋＋最安パスでは AI推奨・Drive CSV を **正本にしない**  
- 競合ゼロ時のみ §2.7（その後も §2.8）

### 2.5 既存値との関係

| 現状 | 挙動（v1案） |
|------|----------------|
| ID・名が選定結果と同一 | skip（書込統計のみ） |
| 既に値あり・選定結果が異なる | **上書き**（メニュー8のジャンルと同型。ログに before→after・選定理由） |
| 選定不採用／SHP検証失敗 | 既存値を触らない（誤IDで上書きしない） |

### 2.6 ログ（選定理由の必須）

親1件ごとに Logger へ要約（シークレット無し）:

- クエリ種別・hits件数・候補ID・minPrice・自社価格  
- SHP結果: `shp_validated` / `shp_resolved` / `shp_reject` / `shp_auth_fail`  
- 採用IDと理由コード

---

## 3. トグル・認証

| Key | 既定 | 意味 |
|-----|------|------|
| `AMAZON_AI_AUTO_ADOPT_ENABLED` | false | メニュー8本体（既存） |
| `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED` | **未設定＝ON** | Yahooカテゴリ／ブランド。false なら当該3列を触らない |
| `YAHOO_SHOPPING_CLIENT_ID` | （既存） | 商品検索 v3（候補用） |
| `AMAZON_AI_ADOPT_YAHOO_CAT_CANDIDATE_K` | **3**（任意） | 最安プローブ候補数 |
| Yahooマッピング B12–B15 | （既存） | SHP用 Client／Secret／Refresh／seller_id |

---

## 4. 制約

- 楽天CSV生成・Yahoo.js の editItem／画像／在庫送信・B統合境界を変えない（SHP読取ヘルパ追加は可）  
- マスタ埋めのみ／価格セルは読取のみ  
- シークレットをログに出さない  
- レ点親のみ・MAX_PARENTS  
- UrlFetch: 候補側＋**SHP検証1〜2回／親**  
- ショップ path 未作成は出品時エラーになり得る（本 Stage 対象外）

---

## 5. 検収

- [ ] Yahooトグル false で当該3列が不変  
- [ ] SCに無いID（例: 45944）は書かれない／名前解決で正しいコードに差し替わる  
- [ ] ログに `shp_validated` または `shp_resolved` / `shp_reject`  
- [ ] SHP認証失敗時はカテゴリ非書込＋要確認  
- [ ] editItem 本体が従来どおり  
- [ ] Property を false に戻せる  

---

## 6. Stage 境界

| Stage | 内容 | 状態 |
|-------|------|------|
| **0** | 要件・承認 | 完了 |
| **1** | メニュー8候補書込 | 実装済 |
| **1b** | **SHP getShopCategoryList 正本検証** | **実装済** |
| **1c** | **SHP getShopBrandList 正本検証・メーカー名取得** | **実装済** |
| 2 | 競合ゼロ AI/Drive 候補 | 実装済（検証必須） |
| 3 | downloadShopCategories キャッシュ | **未着手・必要時のみ**（§6.1） |
| 4 | shop-categories path **新規作成** | **未着手・必要時のみ**（§6.1）。現状はマスタへ path 文字列書込のみ |

### 6.1 必要になったら実装するメモ（コードなし）

いまは実装しない。症状が出たら別承認で着手。

| 候補 | いつ必要か | 実装のヒント |
|------|------------|--------------|
| **Stage3 `downloadShopCategories` キャッシュ** | メニュー8で親件数が増え、都度 `getShopCategoryList` がレート制限・遅延になるとき | 全件DL→Drive/シートに CategoryCode／PathName を保持し、書込前検証をローカル照合＋不足時のみAPI。聖域: editItem 非改変 |
| **Stage4 ショップ path 自動作成** | マスタの `(Yahooカテゴリ名)` がストアサイトマップに無く、出品時に path エラーが頻発するとき | Yahoo API `category.shop-categories.insert`（要別承認）。既存同名は作成スキップ。メニュー8は「無いとき作成→書込」か出品前チェックのどちらにするか要決定 |

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-12 | **7.5維持／7.6 popular_only＋38074**。B統合は Step7.6 に差し替え。 |
| 2026-07-26 | Stage3/4 は**必要時メモのみ**（実装しない）。 |
| 2026-07-26 | **ブランド正本=getShopBrandList**。コード検証＋メーカー名でAPI取得。 |
| 2026-07-26 | **正本=SHP getShopCategoryList**。候補は市場/AI、書込前検証必須。 |
| 2026-07-26 | **競合ゼロフォールバック**: AI推奨列→Drive YahooマスタCSV。 |
| 2026-07-26 | **実装**: メニュー8へ接続。 |
| 2026-07-26 | 初版〜売れ筋＋自社最安。 |
