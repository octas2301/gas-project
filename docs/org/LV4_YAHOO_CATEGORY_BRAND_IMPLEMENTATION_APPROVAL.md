# 実装承認パッケージ — Yahooカテゴリ／ブランド Stage（都度API）

**日付**: 2026-07-26  
**状態**: **承認済・実装済**（B: getShopCategoryList 正本検証を含む）  
**正本**: [../YAHOO_CATEGORY_BRAND_STAGE.md](../YAHOO_CATEGORY_BRAND_STAGE.md)  
**手順**: [D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md](D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md)

---

## 変更予定ファイル

| ファイル | 内容 |
|----------|------|
| `コード.js` | メニュー8に Yahooカテゴリ／ブランド都度APIフェーズ追加。itemSearch で genreCategory／brand 投票→階層名 `:` 連結→マスタ3列書込。共通ヘルパ |
| `docs/YAHOO_CATEGORY_BRAND_STAGE.md` | 要件（作成済） |
| `docs/org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md` | 人間手順（作成済） |
| `docs/org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md` | v1.8 予定として追記 |
| `docs/org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md` | Yahooトグル追記 |
| `docs/CHANGE_LEDGER.md` / `CURRENT_PHASE.md` / `AGENT_HANDOVER.md` | 記録 |

**新規**: 本承認パッケージ／Stage 要件／HUMAN_RUN（上記）

**触らない**: `generateRakutenCSV`／**Yahoo.js**／B統合境界／AIカテゴリ生成プロンプト（正本にしないだけ）／shop-categories 作成API

---

## 変更概要

1. AIの★推奨Yahooカテゴリ／ブランドは使わない  
2. 都度 **Yahoo!ショッピング商品検索 v3**（既存 `YAHOO_SHOPPING_CLIENT_ID`）  
3. **YahooカテゴリID**: 売れ筋近似（`sort=-review_count`）の競合を参照し、候補カテゴリごとに最安をプローブ。**自社 `Yahoo!価格設定` が競合最安より安くなるカテゴリを優先**（単純多数決ではない）  
4. `parentGenreCategories`（depth順）＋葉名を **`:` 連結** → `(Yahooカテゴリ名)`  
5. `brand.id` を投票 → `Yahooブランドコード`  
6. 専用トグル `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED`（既定 false）  
7. **価格列は読取のみ**（書き換えない）  

---

## 想定リスク

| リスク | 緩和 |
|--------|------|
| 「売れ筋」が review_count 近似（sold非対応） | 要件・ログに明記。将来API拡張時に差し替え可 |
| 誤カテゴリ／誤ブランド | 候補K制限・要確認・ログに理由 |
| 自社価格未設定で最安判定不可 | popular_only フォールバック＋ログ |
| Shopping ID とストア product_category 不一致 | スモーク1件で突合 |
| ショップに同名 path 未作成 | マスタ埋めのみ。出品は別 |
| UrlFetch増（売れ筋＋Kプローブ） | MAX_PARENTS・K既定3・トグル |
| EC重要変更（マスタ一括書込） | **本承認後のみ実装** |

---

## 戻し方

- Property `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED=false`（本体トグルも false）  
- `git revert`  

---

## 承認文（コピー用）

> **Yahooカテゴリ／ブランド Stage（都度API・メニュー8）を承認**

承認後に Agent が `コード.js` 実装へ進みます。
