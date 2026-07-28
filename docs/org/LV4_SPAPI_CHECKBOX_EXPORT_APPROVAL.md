# SP-API v1.2c — 21-⑧ 子SKUレ点のみ（選択行廃止）

**日付**: 2026-07-28  
**状態**: **承認済・実装済・実機合格**（2026-07-28／29）  
**前提**: 橋渡し v1.2／Drive取得 v1.3  
**手順**: [D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md](D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)

---

## 1. 目的

21-⑧ の対象を **マスタの出品CK付き子SKU行のみ** に変更する。  
**行選択モードは完全廃止。親レ点のみでは出さない**（Yahoo 出品と同じ「子レ点」原則）。

---

## 2. 変更ファイル

| 種別 | パス |
|------|------|
| 改修 | `AmazonSpapiExport.js`（`menuAmazonSpapiExportItemsCsv`） |
| 改修 | `コード.js`（21-⑧ メニュー文言） |
| 新規 | 本ファイル |
| 更新 | HUMAN_RUN／CURRENT_PHASE／HANDOVER／CHANGE_LEDGER／SHEET_BRIDGE 注記 |

**やらない**: GAS からの SP-API 直呼び／親レ点で全子展開／選択行互換フラグ

---

## 3. 仕様

- マスタ全行走査 → **子SKUあり かつ 出品CK**（boolean / `TRUE` 両対応）のみ CSV 化
- 親行（子SKU空）にレ点があっても **出力しない**（件数はログに `parentCkOnlySkipped`）
- `max_items` / `FORCE_QTY_0` / Property トグルは従来どおり
- 21-⑨（承認①済）は変更なし

---

## 4. リスク

| リスク | 緩和 |
|--------|------|
| 意図しない多数レ点 | `max_items` 超過で拒否・作業後 Property false |
| 親だけレ点で「出ない」混乱 | ダイアログで親レ点のみ件数を明示 |
| EC直書込 | GAS は CSV のみ（従来どおり） |

---

## 5. 合格条件

- [x] clasp push 後、子レ点のみで Drive CSV が出る（`SPAPI_EXPORT_20260728_235455_52100a`）  
- [x] 親レ点のみでは出さない（**社長完了扱い** 2026-07-29。21-⑧専用0件ダイアログの単独試験は省略）  
- [x] `--fetch-drive --mode dry_run`（VALID）→ **prod ACCEPTED**  

**実機**（2026-07-28）: SKU `lifec-4560151300924-48s11`／ASIN `B00A0J0D30`／レポート `SPAPI_LISTINGS_WRITE_20260728_145928`。Property／`allow_prod` 作業後 false。

## 6. 社長承認欄

- [x] **承認する**（2026-07-28・選択行完全廃止・子SKUレ点のみ）  
- [ ] 却下／条件付き
