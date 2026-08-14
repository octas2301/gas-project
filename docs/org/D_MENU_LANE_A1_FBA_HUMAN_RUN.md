# レーンA1 — 相乗り FBA（人間手順）

**状態**: **検収OK**（2026-08-01。FBA dry_run VALID → prod ACCEPTED。属性実装込み）  
**承認**: [LV4_LANE_A1_FBA_OFFER_APPROVAL.md](LV4_LANE_A1_FBA_OFFER_APPROVAL.md)／[属性](LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)  
**親手順**: [D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md) §1  
**二列**: [LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md](LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)（**検収OK**）

---

## 0. 目的

D の **相乗りFBA** を、自己発と同型で dry_run → prod まで記録する。

---

## 0b. 実機対比（属性実装前・2026-08-01）

同一ASIN `B01N5A6ESU`。

| | 自己発（MFN） | FBA（属性前） |
|--|--------------|---------------|
| SKU | `…19as13` | `…19af13` |
| dry_run | **VALID** `…105918_199878` | **90220** `…105527_f64afc` |

**結論**: FBAでは電池・危険物必須 → [属性実装](LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md) で解消。

---

## 0c. 合格記録（属性実装後・2026-08-01）

| 段階 | 結果 | runId | SKU／ASIN |
|------|------|-------|-----------|
| **A1-a dry_run／FBA** | **VALID issues=0** | `SPAPI_PUT_OFFER_CK_DRY_20260801_111613_41ce9e` | `sanky-B01N5A6ESU-19af13`／`B01N5A6ESU` |
| **A1-b prod／FBA** | **ACCEPTED issues=0** | `SPAPI_PUT_OFFER_CK_PROD_20260801_111845_6bd20f` | 同上 |

- ログ: `FBA compliance attrs ON`  
- 在庫: ZERO。GET 404 → 新規オファー作成  
- 当時は **NF暫定保存**（二列前）。二列後は `_FBA` 列へ移設  
- 作業後: PUT トグル **false** に戻す  

---

## 1. 事前（二列実装後の再現）

| 項目 | 内容 |
|------|------|
| clasp | `AmazonSpapiPut.js`／`コード.js` 反映済 |
| 列 | ヘッダに **`Amazon相乗りSKU_FBA`** を追加 |
| 移行 | NF の `…af…` は人手で `_FBA` へ移し、NFは自己発用に空or `…as…` |
| N列 | 既存カタログASIN（試験は `B01N5A6ESU`） |
| Property | `FBA_COMPLIANCE_ATTRS` 未設定＝true |

---

## 2〜3. 手順要約

dry_run（FBA）→ VALID → **`Amazon相乗りSKU_FBA` のみ更新**（NF不変）→ `ALLOW_PROD` → prod（FBA）→ トグル false。  
失敗時は【おすすめ】表示 → 再 dry_run（prod直禁止）。

A1合格記録は §0c（二列前の暫定NF保存）。

---

## 4. 合格目安

- [x] 自己発 dry_run 対比（七味）  
- [x] FBA dry_run VALID（属性後）  
- [x] FBA prod 1SKU ACCEPTED  
- [x] （運用）作業後トグル false  
- [x] （二列後）FBA dry_run で `_FBA` のみ更新… `…114820_d6ed67`（自己発 dry `…114254_8fa79e`）  
- [x] （二列後）各系統 prod… 自己発 `…115554_4ed30e`／FBA `…115648_eb2511`  

---

## 5. やってはいけない

- コンプライアンス変更後の dry_run スキップ  
- `ALLOW_MASTER_QTY=true` での A1  
- FBA実行で自己発列を手で消す（コードは触らないが、人手移設ミスに注意）  

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草〜対比・属性実装。 |
| 2026-08-01 | **検収OK**: dry_run `…111613_41ce9e`／prod `…111845_6bd20f`。 |
| 2026-08-01 | デュアル Phase1実装後の保存先＝`Amazon相乗りSKU_FBA` に改定。 |
| 2026-08-01 | **二列 dry_run OK**: FBA `…114820_d6ed67`／自己発 `…114254_8fa79e`。 |
| 2026-08-01 | **二列 prod OK**: 自己発 `…4ed30e`／FBA `…eb2511`。 |
