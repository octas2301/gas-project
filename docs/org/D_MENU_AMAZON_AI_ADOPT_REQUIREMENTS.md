# Amazon 向け KW選択・重複削除（メニュー8）— 要件定義

**最終更新**: 2026-07-26  
**状態**: **v1.8 実装済**（Yahooカテゴリ／ブランド都度API・要 clasp push）

---

## 目的

KW横断・最終名文字数・バリエーションテーマに加え、**楽天ジャンルID／名称を都度APIで選定**する（AI推奨ジャンルは使わない）。  
**v1.8**: YahooカテゴリID・(Yahooカテゴリ名)・Yahooブランドコードも都度APIで選定。

---

## 入口

Z → **7.5** / `AMAZON_AI_AUTO_ADOPT_ENABLED`  
楽天ジャンル追加: `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED`（既定 **false**）  
Yahooカテゴリ／ブランド: `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED`（既定 **false**）

---

## 処理（v1.8）

### A〜B. KW・テーマ（v1.5〜1.6）

（従来どおり）

### C. 楽天ジャンル（Stage3・都度API）

正本: [RAKUTEN_NAV_GENRE_STAGE3.md](../RAKUTEN_NAV_GENRE_STAGE3.md)

1. トグルON時のみ  
2. クエリ: JAN → 商品名ベース  
3. Ichiba Item Search で `genreId` 投票（過半数 or ≥3票）  
4. NavigationAPI で名称確定 → `楽天ジャンルID` / `楽天ジャンルID名`  
5. AIの★推奨楽天ジャンルは読まない  
6. `generateRakutenCSV` 非改変  

### D. Yahooカテゴリ／ブランド（Stage・都度API）

正本: [YAHOO_CATEGORY_BRAND_STAGE.md](../YAHOO_CATEGORY_BRAND_STAGE.md)  
承認: [LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md](LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md)

1. トグルON時のみ  
2. クエリ: JAN → 商品名ベース  
3. **カテゴリID**: 候補（売れ筋＋最安／競合0はAI・Drive）→ **`getShopCategoryList` 検証後のみ書込**  
4. 無効IDは却下。名前一致なら SC の CategoryCode に差し替え（`shp_resolved`）  
5. path は SHP PathName を `:` 正規化して優先  
6. ブランド投票（不足時フォールバック）／価格列は読取のみ  
7. Yahoo.js **editItem非改変**（SHP読取ヘルパのみ追加）  

---

## 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-26 | **v1.8 実装**: Yahooカテゴリ／ブランド都度API（売れ筋＋自社最安）。要 clasp push |
| 2026-07-26 | **v1.8 要件改定**: Yahooカテゴリ＝売れ筋競合＋自社最安優先（単純多数決廃止） |
| 2026-07-26 | **v1.8 要件起草**: Yahooカテゴリ／ブランド都度API（実装は承認後） |
| 2026-07-26 | **v1.7**: 楽天ジャンル都度API（Ichiba→Nav）をメニュー8へ |
| 2026-07-26 | v1.6: バリエーションテーマ |
| 2026-07-26 | v1.5: 下限・容量・特徴 |
