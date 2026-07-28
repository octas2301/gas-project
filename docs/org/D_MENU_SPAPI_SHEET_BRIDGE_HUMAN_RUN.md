# SP-API スプシ橋渡し v1.2（人間手順）

**状態**: **実機合格**（2026-07-28）／21-⑧→Drive→ローカル dry_run/prod→SC反映確認済  
**承認**: [LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md](LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md)  
**書込本体**: [D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md](D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)（ローカル）

---

## 0. できること／できないこと

| できる | できない（別承認） |
|--------|-------------------|
| マスタ選択行 → Drive に items CSV | GAS から SP-API PUT |
| ローカル dry_run／prod | 全件ループ・在庫>0無人 |

---

## 1. Script Properties

| キー | 既定 | 内容 |
|------|------|------|
| `APPROVAL_AMAZON_SPAPI_EXPORT_ENABLED` | **false** | `true` でメニュー有効 |
| `APPROVAL_AMAZON_SPAPI_EXPORT_MAX_ITEMS` | `5` | 超過は拒否 |
| `APPROVAL_AMAZON_SPAPI_EXPORT_FORCE_QTY_0` | **true** | 在庫0強制 |
| `APPROVAL_AMAZON_SPAPI_EXPORT_FOLDER_ID` | 空 | 空なら Lv4 GENERATED フォルダ流用 |

---

## 2. 手順

1. `clasp push`（`AmazonSpapiExport.js`＋`コード.js`）  
2. Property `APPROVAL_AMAZON_SPAPI_EXPORT_ENABLED=true`  
3. マスタで **相乗り用の子行**（ASIN・販売価格amazon あり）を選択  
4. **Z → 21 → 21-⑧ SP-API用items CSV出力**  
5. Drive の `SPAPI_EXPORT_*_SPAPI_ITEMS.csv` をダウンロード  
6. `tools/spapi_listings_write/items.csv` に上書き配置  
7. `python spapi_listings_write.py --mode dry_run` → OKなら prod（`allow_prod`）  
8. 終わったら Property を **false**、`allow_prod` も **false**

---

## 3. 列の取り方

| CSV列 | マスタ |
|-------|--------|
| sku | 子SKU（空なら親SKU） |
| asin | ASINコード → 競合店ASIN → URL |
| price | 販売価格amazon（行→親） |
| quantity | 既定0（FORCE） |

---

## 4. 合格目安

- [x] 21-⑧ で CSV が Drive に出る  
- [x] ローカル dry_run が読める  
- [x] ローカル prod → SC で価格／在庫確認（試験SKU）  
- [x] Property 既定／作業後 false  

### 4.1 実機記録（2026-07-28）

| 項目 | 結果 |
|------|------|
| 21-⑧ | Drive `SPAPI_EXPORT_*_SPAPI_ITEMS.csv` 出力 OK |
| アルコール | `lifec-4560151300139-ride01`／`B0091G3AHY`／1000／0・SC反映 |
| 安眠 | `lifec-4560151300924-ride01`／`B00A0J0D30`／1000／0・SC反映 |
| 発汗（先行 v1） | `…48s11`／`B07YND44VN`・出品中確認 |
| 注意 | 競合ASINが制限付きノーブランドだと code 5886。相乗り可ASINを正にする |
| 注意 | Excel 保存の Shift-JIS CSV は UTF-8 で再保存（または v1.2a） |

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-28 | **実機合格**記録。ride01 アルコール／安眠・1000/0。 |
| 2026-07-28 | v1.2 初版。選択行→Drive CSV。 |
