# デュアル Phase2 — 両系統1実行（人間手順）

**状態**: **検収OK**（2026-08-01）— 両系統 dry／prod  
**承認**: [LV4_DUAL_OFFER_PHASE2_APPROVAL.md](LV4_DUAL_OFFER_PHASE2_APPROVAL.md)  
**親**: [D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)

---

## 0. 目的

D で相乗り自己発と FBA を同時チェックし、1実行で最大2PUTする。

---

## 1. 事前

| Property | 値 |
|----------|-----|
| `ENABLED` | true |
| `ALLOW_PROD` | dry後に true |
| `MAX_ITEMS` | 1 推奨 |
| 列 | NF＋`Amazon相乗りSKU_FBA` あり。両系統に SKU が入っているか（prod時） |

作業後は PUT トグルを **false** に戻す。

---

## 2. 手順

1. レ点1行  
2. D → Amazonのみ → 相乗り → **自己発と FBA の両方にチェック**  
3. dry_run → 確認に「2系統」と保存列 → OK  
4. ログ: 自己発 PUT のあと FBA PUT。各 `SAVE_OFFER_SKU` が正しい header  
5. prod 同様  
6. トグル false  

片方だけチェック＝Phase1どおり。

---

## 3. 合格記録

| 段階 | 結果 | runId |
|------|------|-------|
| 両方 dry | **OK** | 自己発: `SPAPI_PUT_OFFER_CK_DRY_20260801_141229_f372b8`（VALID／NF／`…33as19`）／FBA: `…141243_40d85e`（VALID／`_FBA`／`…33af19`） |
| 両方 prod | **OK** | 自己発: `SPAPI_PUT_OFFER_CK_PROD_20260801_141420_8fcbc2`（ACCEPTED）／FBA: `…141432_cc7c72`（ACCEPTED・compliance ON） |

**対象**: 行494／ASIN `B0D9VK8YPS`／親 `sanky-4538872081514-oya`／`inventoryMode=ZERO`  
**順序**: 自己発→FBA（`throwOnFail=0`）。最終 `offerRunId=mfn:…\|fba:…`

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。 |
| 2026-08-01 | コード実装済。実機手順は §2。 |
| 2026-08-01 | 初回 dry で `amazonSpapiPutOfferSellerSkuHeader_ is not defined` → 定義復元。要 `clasp push` 再実行。 |
| 2026-08-01 | **検収OK**: 両方 dry `…f372b8`／`…40d85e`・両方 prod `…8fcbc2`／`…cc7c72`。 |
