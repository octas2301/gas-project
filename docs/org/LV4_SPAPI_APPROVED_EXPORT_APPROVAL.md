# SP-API v1.2b 承認①済一括CSV — 実装承認パッケージ

**日付**: 2026-07-28  
**状態**: **承認済・実装済・実機合格**（2026-07-29・21-⑨→`--fetch-drive` dry_run／prod）  
**前提**: 橋渡し v1.2／Drive取得 v1.3  
**手順**: [D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md](D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)

---

## 1. 目的

選択行ではなく、**最新承認①済 Amazon（子SKU）** をまとめて Drive CSV 化（メニュー **21-⑨**）。

---

## 2. 変更ファイル

| 種別 | パス |
|------|------|
| 改修 | `AmazonSpapiExport.js`（`menuAmazonSpapiExportApprovedItemsCsv`） |
| 改修 | `コード.js`（21-⑨） |
| 新規 | 本ファイル |
| 更新 | HUMAN_RUN／CURRENT_PHASE 等 |

**やらない**: SP-API直呼び、全件マスタ、親行のみの強制出品

---

## 3. 仕様

- Property トグルは 21-⑧ と同じ `APPROVAL_AMAZON_SPAPI_EXPORT_ENABLED`
- 対象: 最新 APPROVED バッチの mall=amazon・子SKUあり・ASIN・販売価格amazon
- max_items／FORCE_QTY_0 共通
- 親行のみ（子SKU空）はスキップ

---

## 4. 社長承認欄

- [x] **承認する**（おすすめ順・v1.2b／2026-07-28）  
- [ ] 却下／条件付き

## 5. 実機合格（2026-07-29）

- [x] 21-⑨ → `SPAPI_EXPORT_APPR_20260729_001815_127e8a`（batch `A1_20260727_224939_b7a053`・親 `…oya` スキップ）  
- [x] dry_run VALID → prod ACCEPTED（`lifec-4560151300832-48s11`／`B07YND44VN`・レポート `…152611`）  
- [x] Property／`allow_prod` 作業後 false  
