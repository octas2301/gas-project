# Amazon カテゴリ／PT・競合参照 — 読取 PoC 調査承認（P4a）

**日付**: 2026-08-01  
**状態**: **実機合格**（読取のみ。マスタ書込なし。三点不要）  
**手順**: [LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md](LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md)  
**ロードマップ**: [AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md)（読取済。本線書込＝[P4b](LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)）  
**関連**: SP-API Product Type Definitions／Catalog Items／メニューA Keepa（[RESEARCH_AND_ESTIMATE.md](../RESEARCH_AND_ESTIMATE.md)）／既存スモーク [D_MENU_SPAPI_SMOKE_HUMAN_RUN.md](D_MENU_SPAPI_SMOKE_HUMAN_RUN.md)  
**三者レビュー**: **不要**（読取・件数制限・本番D非接続）

---

## 1. 目的

現場ではカテゴリ／PT決定が作業先頭に来る。公式 API と Keepa で **読取だけ**確かめ、楽天ジャンル連携なしで Amazon 側を賄えるかを結論づける。

---

## 2. PoC 範囲（含む）

| # | 検証 | 成功の定義 | 実装経路 |
|---|------|------------|----------|
| 1 | SP-API `searchDefinitionsProductTypes`（keywords／itemName） | JPマーケットで候補PTが返る | `tools/spapi_smoke` `--poc-category` |
| 2 | SP-API `getDefinitionsProductType`（必要なら） | 定義／推奨 browse 等を取得できる | 同上（先頭候補または config） |
| 3 | Catalog Items で既存競合ASINの属性参照 | 分類に使えるフィールドが取れる | 同上（includedData 拡張） |
| 4 | Keepa（メニューA延長）で競合ASINのカテゴリ系 | 使えるカテゴリ情報がある／ないを記録。**token制限・少件数** | **人間・メニューA**（GAS改変なし） |
| 5 | 純正 `.xlsm` テンプレの自動DL | **できる／できない**を結論 | **調査結論**: SP-API Definitions＝JSON。xlsm自動DLは不可想定→レポート明記 |

**制約**: マスタ・承認ログ・SCへの書込禁止。全件ループ禁止。Keepaは目安数ASIN以内。

---

## 3. 含まない（P4b以降）

- マスタ「amazon カテゴリー」等への自動書込  
- C1列マップへの本線接続  
- Dメニューからの本番カテゴリ決定  
- 楽天ジャンル Stage3 の削除（Amazon用に必須でなくなれば依存を外すだけ）  

---

## 4. 成果物

- `tools/spapi_smoke/out/SPAPI_CATEGORY_PT_*_REPORT.json`  
- 本ファイルまたは HUMAN_RUN への実機結果追記  
- 次ゲート案: P4b に進む／Definitionsのみ採用／Keepaは補完のみ、等  

---

## 5. 社長確認

- [x] 調査承認（本 PoC 着手可）… 2026-08-01「おすすめの順で進めて」  
- [x] Keepa token 消費に同意（少件数・メニューA手動）  
- [x] 実装: 既存 `tools/spapi_smoke` 拡張＋HUMAN_RUN（新規GASメニューなし・マスタ非書込）  

---

## 6. 実装メモ（読取のみ）

| ファイル | 内容 |
|----------|------|
| `tools/spapi_smoke/spapi_smoke.py` | `--poc-category`: search／getDefinitions／Catalog拡張 |
| `tools/spapi_smoke/config.example.json` | keywords／item_name／product_type 例 |
| docs | 本承認・HUMAN_RUN・PHASE／LEDGER |

**リスク**: LWA／ロール不足で 403。Keepa token。レポートへの秘密混入（トークンは書かない）。  
**戻し**: `--poc-category` を使わなければ従来スモークのみ。git revert。

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。ロードマップ P4a。 |
| 2026-08-01 | 調査承認。smoke拡張＋HUMAN_RUN。実機は人間実行。 |
| 2026-08-01 | **実機合格**（七味→HERB/SEASONING、Catalog分類可、xlsm自動DL不可）。次＝P1。 |
| 2026-08-01 | 次開発誘導: **P4b** — [LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md](LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)。 |
