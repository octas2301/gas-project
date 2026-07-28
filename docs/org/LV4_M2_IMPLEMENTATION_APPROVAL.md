# Amazon M2（TRACK=A）— 承認パッケージ

**日付**: 2026-07-27  
**状態**: **実装承認済（社長決定反映）・v1実装済・実機未**  
**要件正本**: [LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) §8・§12.4  
**ギャップ**: [LV4_M2_TRACK_A_GAP_ANALYSIS.md](LV4_M2_TRACK_A_GAP_ANALYSIS.md)  
**HUMAN_RUN**: [D_MENU_M2_HUMAN_RUN.md](D_MENU_M2_HUMAN_RUN.md)  
**Facade**: U6  

---

## 0. 社長決定（2026-07-27）

| # | 問い | 決定 |
|---|------|------|
| 1 | PACKAGED | **案L**（Listing Loader／在庫ファイル系 CSV） |
| 2 | ブランドゲート | **人間の目**＋`APPROVAL_AMAZON_LV4_BRAND_GATE_MODE=manual_ok` |
| 3 | 試験SKU | **発汗チェッカー** `lifec-4560151300832-oya`／子`…-48s11`／ASIN`B07YND44VN`（競合列） |

---

## 1. 変更ファイル一覧（v1）

| 種別 | パス | 内容 |
|------|------|------|
| 改修 | `AmazonApprovalExport.js` | A用 ASIN（`ASINコード`／競合店ASIN／URL）・`SKIPPED_BRAND_GATE` |
| 新規 | `tools/m2_offer_packaged/*` | 案L CSV 変換 |
| 更新 | docs（本ファイル／HUMAN_RUN／ギャップ／CURRENT_PHASE等） | 状態 |
| **触らない** | 楽天CSV／Yahoo出品API／C1 HPC／マスタ在庫書込 | 聖域 |

---

## 2. 概要

TRACK=`A` GENERATED（offer）→ `m2_offer_packaged.py` → SC手UP。C1は使わない。

---

## 3. 想定リスク

| リスク | 緩和 |
|--------|------|
| TRACK=B取り違え | Property 明示 |
| ブランド制限 | manual_ok 必須 |
| 在庫>0スキップ | HUMAN_RUNで試験時0 |
| 公式テンプレ列差 | `column_map.json` で調整 |
| 既存出品衝突 | 小バッチ・目視 |

---

## 4. 復元

- `APPROVAL_AMAZON_LV4_ENABLED=false`  
- `APPROVAL_AMAZON_LV4_BRAND_GATE_MODE` 削除  
- `git revert`／`tools/m2_offer_packaged` 削除  

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-27 | 決定反映＋v1実装。試験=発汗。 |
| 2026-07-27 | 下書き初版。 |
