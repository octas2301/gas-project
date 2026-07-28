# SP-API Listings 複数行バッチ v1.1 — 実装承認パッケージ

**日付**: 2026-07-28  
**状態**: **承認済・実装済・実機合格**（2026-07-28・ride01 prod／SC反映）  
**前提**: v1 1SKU prod 合格（発汗）／[D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md](D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)  
**親承認**: [LV4_SPAPI_LISTINGS_WRITE_IMPLEMENTATION_APPROVAL.md](LV4_SPAPI_LISTINGS_WRITE_IMPLEMENTATION_APPROVAL.md)

---

## 1. 目的

ローカル CSV で **複数 SKU** の既存カタログ相乗り（LISTING_OFFER_ONLY）を dry_run→prod。  
**GAS／スプシ直結なし。**

---

## 2. 変更ファイル

| 種別 | パス |
|------|------|
| 改修 | `tools/spapi_listings_write/spapi_listings_write.py` |
| 新規 | `tools/spapi_listings_write/items.example.csv` |
| 改修 | `config.example.json`／`.gitignore`（`items.csv` 除外） |
| 新規 | 本ファイル |
| 更新 | HUMAN_RUN／CURRENT_PHASE／HANDOVER／CHANGE_LEDGER |

**やらない**: GAS・スプシ直結・全件ループ・新規カタログ作成

---

## 3. 仕様要約

- `items_csv`（`sku,asin,price,quantity[,note]`）優先。無ければ従来1件  
- `max_items` 既定 5（1〜50）  
- 行単位処理・失敗しても続行・サマリ ok/fail  
- prod は `allow_prod=true` 必須  
- 試験: 安眠相乗り `B00A0J0D30`／アルコール相乗り `B0091G3AHY`、**価格1000・在庫0**  
- 相乗り用 seller SKU は C1 新規カタログSKUと衝突しない `…-ride01`（変更可）

---

## 4. リスク

| リスク | 緩和 |
|--------|------|
| 複数誤送信 | max_items／allow_prod／在庫0 |
| 既存SKUとASIN衝突 | ride01 新規SKU／dry_run |
| 秘密漏洩 | items.csv・config.local は gitignore |

---

## 5. 社長承認欄

- [x] **承認する**（2026-07-28・複数行バッチ v1.1／相乗りASIN・価格1000・在庫0）  
- [ ] 却下／条件付き
