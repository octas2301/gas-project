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
5. 確認（**7.6 推奨・Bと同系**）:
   - ログ `BRANCH_SAME` / `BRANCH_VOTE` / `BRANCH_WEIGHT_IN_VOTE`
   - `YahooカテゴリID` が SC に存在（`shp_validated`）
6. 仮想だけ先に見る場合（B不要）:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/yahoo_category_virt/run_virt.py
```

7. 全トグル **false**

## スモーク注意

- 競合0／同商品0 → AI/Driveフォールバック（要確認付き）。書込前に必ず SHP 検証
- SHP認証失敗時はカテゴリ／ブランド非書込＋要確認
- ショップ path の **新規作成 API は未実装**（マスタへ path 文字列を書くだけ）
- 単独はレ点 **1親** に絞って Z **7.6**（B全体より短い）
