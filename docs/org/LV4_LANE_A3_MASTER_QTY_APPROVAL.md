# レーンA3 — 自己発マスタ在庫>0（承認パッケージ）

**日付**: 2026-08-01  
**状態**: **検収OK**（2026-08-01）。三点スキップ。原則コードなし（P0経路）  
**親**: [LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md](LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md)／[LV4_LANE_A1_FBA_OFFER_APPROVAL.md](LV4_LANE_A1_FBA_OFFER_APPROVAL.md) §8  
**前提**: A1・デュアル Phase1・A2 **検収OK**  
**手順**: [D_MENU_LANE_A3_HUMAN_RUN.md](D_MENU_LANE_A3_HUMAN_RUN.md)  
**三者レビュー**: **スキップ**

---

## 1. 目的

P0 実装済の **「マスタ在庫で出す」**（承認②相当）を、レーンAの相乗り **自己発**で dry_run → prod（qty>0）まで実機検収する。

| 段階 | 内容 | 成功定義 |
|------|------|----------|
| **A3-a** | 自己発＋MASTER＋dry_run（1SKU・qty>0） | VALID。`inventoryMode=MASTER`／`forceQty0=false` |
| **A3-b** | 同条件 prod（1SKU） | ACCEPTED |
| **A3-c** | トグル戻し | `ALLOW_MASTER_QTY`／`ALLOW_PROD`／`ENABLED` → **false** |

**含まない**: FBAのqty検収、無人全件、FLOOR(仕入÷セット)、マスタ列書込、楽天／Yahoo／B統合。

---

## 2. 社長確定方針（2026-08-01）

| # | 論点 | 決定 |
|---|------|------|
| 1 | 対象 | **自己発（MFN）のみ** |
| 2 | 件数 | **1SKU** |
| 3 | 数量 | マスタ「在庫数」**生値** |
| 4 | コード | **原則なし** |
| 5 | マスタ列 | **書込禁止** |
| 6 | 三点 | **スキップ** |

---

## 3. 変更ファイル

| 種別 | パス | 内容 |
|------|------|------|
| 新規 | 本ファイル／HUMAN_RUN | 正本・手順 |
| 更新 | PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | 到達点 |

---

## 4. 想定リスクと緩和

| リスク | 緩和 |
|--------|------|
| qty>0 実売・過剰 | 1SKU・確認内訳・専用トグル |
| トグル忘れ | A3-c＋CURRENT_PHASE |
| FBAでqty期待 | A3は自己発のみ |

---

## 5. 検収

- [x] 方針・docs起草… **2026-08-01**  
- [x] A3-a… **`SPAPI_PUT_OFFER_CK_DRY_20260801_121701_49a49e`** VALID（MASTER／`forceQty0=false`／`…19as13`）  
- [x] A3-b… **`SPAPI_PUT_OFFER_CK_PROD_20260801_121813_f677a3`** ACCEPTED  
- [x] docs 記録  
- [ ] A3-c トグル false… **人間確認**（作業後必須）  
- [ ] マスタ「在庫数」不変… **人間目視**（実行前後メモ）  

### 5.1 実機ログ要約

| mode | runId | 要点 |
|------|-------|------|
| dry_run | `…121701_49a49e` | `inventoryMode=MASTER`／items=1／VALID |
| prod | `…121813_f677a3` | 同上／ACCEPTED／行503 |

---

## 6. 社長確認

- [x] §2… **2026-08-01**  
- [x] dry／prod 実機… **2026-08-01**  
- [ ] A3-c（トグル）… 作業後に実施  

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草・承認反映。実機待ち。 |
| 2026-08-01 | **dry／prod 検収OK**（`…49a49e`／`…f677a3`）。トグル戻しは人間。 |
