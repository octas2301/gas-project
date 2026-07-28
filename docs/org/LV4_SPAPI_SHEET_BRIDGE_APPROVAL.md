# SP-API スプシ橋渡し v1.2 — 実装承認パッケージ

**日付**: 2026-07-28  
**状態**: **承認済・実装済・実機合格**（2026-07-28）  
**前提**: Listings v1／v1.1 ローカル合格  
**手順**: [D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md](D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)

---

## 1. 目的

マスタ選択行から **SP-API用 items CSV** を Drive に出し、既存ローカル `spapi_listings_write` で dry_run→prod。  
**GAS から SP-API は呼ばない。全件ループ・在庫>0無人は対象外。**

> **追記（v1.2c）**: 21-⑧ の対象は **選択行 → 子SKU＋出品CK（レ点）のみ** に変更。選択行モードは廃止。[CHECKBOX_EXPORT](LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md)

---

## 2. 変更ファイル

| 種別 | パス |
|------|------|
| 新規 | `AmazonSpapiExport.js` |
| 改修 | `コード.js`（メニュー 21-⑧ のみ） |
| 新規 | 本ファイル／HUMAN_RUN |
| 更新 | CURRENT_PHASE／HANDOVER／CHANGE_LEDGER／WRITE HUMAN_RUN |

**やらない**: UrlFetch Listings／LWAをGASへ／楽天・Yahoo改変／全件

---

## 3. 仕様

- Property `APPROVAL_AMAZON_SPAPI_EXPORT_ENABLED` 既定 **false**
- マスタ選択行 → `sku,asin,price,quantity,note` CSV → Drive
- ASIN: `ASINコード` → `競合店ASINコード` → URL（Lv4ヘルパ流用）
- 価格: `販売価格amazon`（行→親フォールバック）
- 在庫: 既定 **0強制**（`FORCE_QTY_0` 既定 true）
- `max_items` 既定 5（1〜50）

---

## 4. リスク

| リスク | 緩和 |
|--------|------|
| 誤行出力 | 選択行のみ・asin/価格必須・件数上限・トグル |
| 在庫誤爆 | FORCE_QTY_0 |
| EC直書込 | **本v1.2ではGASから書かない** |

---

## 5. 合格条件

- [x] 21-⑧ で CSV が Drive に出る  
- [x] ローカル dry_run／prod → SC 反映（ride01）  
- [x] Property 作業後 false  

## 6. 社長承認欄

- [x] **承認する**（2026-07-28・CSVエクスポートのみ・Drive経由・ローカル継続）  
- [ ] 却下／条件付き

**実機**: 2026-07-28 橋渡し→SC確認済（安眠／アルコール ride01・1000/0）。
