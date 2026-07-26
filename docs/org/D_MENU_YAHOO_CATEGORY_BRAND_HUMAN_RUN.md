# Yahooカテゴリ／ブランド Stage — 人間手順

**正本**: [../YAHOO_CATEGORY_BRAND_STAGE.md](../YAHOO_CATEGORY_BRAND_STAGE.md)  
**承認**: [LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md](LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md)  
**状態**: **実装済**（要 clasp push）

## Properties

| Key | 値 |
|-----|-----|
| `AMAZON_AI_AUTO_ADOPT_ENABLED` | `true`（実行後 false） |
| `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED` | Yahooもやるとき `true`（実行後 false） |
| `YAHOO_SHOPPING_CLIENT_ID` | 商品検索 v3（候補用・既存） |
| Yahooマッピング B12–B15 | **SHP検証必須**（Client／Secret／Refresh／seller_id） |

任意（同時実行可）:

| Key | 値 |
|-----|-----|
| `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED` | 楽天ジャンルもやるとき `true` |

## 手順

1. `clasp push`（`コード.js` と **`Yahoo.js`**）
2. 上記 Property ＋ Yahooマッピング認証が有効なこと
3. レ点親を1件に絞る
4. Z → **7.5**
5. 確認:
   - `YahooカテゴリID` が **SCのプロダクトカテゴリ検索に存在する**こと
   - ログに `shp_validated` / `shp_resolved` / `shp_brand_validated` / `shp_brand_from_maker`
   - ブランドも SC／API 上の有効コードであること
   - ショップ path 未作成は出品時エラーになり得る（**自動作成は未実装**。既存サイトマップに同名がある場合だけ通る）
6. 全トグル **false**

## スモーク注意

- 競合0 → AI/Driveは**候補のみ**。書込前に必ず SHP 検証
- SHP認証失敗時はカテゴリ／ブランド非書込＋要確認
- ショップ path の **新規作成 API は未実装**（マスタへ path 文字列を書くだけ）
