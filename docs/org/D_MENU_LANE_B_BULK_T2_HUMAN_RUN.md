# レーンB — B-T2（新PT／単体・別複合）人間手順

**状態**: **方針ロック済**（2026-08-01）／**実需1系統（GROCERY）棚登録 2026-08-02**  
**承認**: [LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md](LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md)  
**親**: [B-T1 HUMAN_RUN](D_MENU_LANE_B_BULK_T1_HUMAN_RUN.md)／[B-T0](D_MENU_LANE_B_BULK_T0_HUMAN_RUN.md)  
**実行**: **ローカル Python は Agent モード**。SC DL／UL は人間。

---

## 0. 要約（ロック定義）

| 項目 | 内容 |
|------|------|
| 複合 | **複数あり得る**。いまの SEASONING 用複合は棚の1つ |
| DL不要 | 棚に **同型（PT＋指紋）** があれば不要 |
| DL必要 | 棚に無い系統／単体が要る、または **Amazon純正更新** |
| browse | 人間の判断材料。自動OKの正ではない |
| HERB.xlsm | 自動本線化しない |
| 実装 | 実需PTが決まってから |
| 七味 | **出品中**（本手順外） |

### 0b. 実需記録 — GROCERY（缶飯・2026-08-02）

| 項目 | 値 |
|------|-----|
| 純正 | `FOOD_FISH_GROCERY.xlsm`（候補 PT: **FOOD / FISH / GROCERY**。**MEATなし**） |
| Catalog | 参考ASINが `MEAT` を返した → **エイリアス MEAT→GROCERY**（`AmazonCategoryPt.js`／Dゲート） |
| 指紋 sha | `74ccdcf96c22879dc80cbe87e8b41aa615e923529f151f091552ccbe3cefb010`（行3–5・maxCol310） |
| 棚 | `shelf_registry` version **B-T1-2**（GROCERY／FOOD／FISH 同一指紋＋SEASONING） |
| C1マップ | `food_fish_grocery_column_map.json` |
| Drive | 09→指紋後 **06** 配置。registry は Drive 同期必須 |

---

## 1. 人間チェックリスト（実需が出たら）

1. P4b等で PT／browse を確定  
2. Agent に棚引き（`--product-type`）  
3. `DL_NOT_NEEDED` → **B-T1 へ**（B-T2不要）  
4. `DL_REQUIRED` → SCで **必要なら別系統の複合**（または単体）を DL → **09**  
5. B-T0 →（必要ならエイリアス）→ 06／registry → B-T1 dry_run／prod → SC UP  

---

## 2. Agent（実装後・仮）

```text
cd tools\c1_hpc_packaged
python c1_bulk_shelf_lookup.py --product-type （PT）
```

---

## 3. 合格記録

| 段階 | 結果 | メモ |
|------|------|------|
| 方針ロック | **OK** | 2026-08-01（複合複数・PT＋指紋明示） |
| 三点スキップ | **OK** | 2026-08-01 |
| 実需PT決定 | **OK** | GROCERY（缶飯・2026-08-02） |
| 実装承認 | **OK（三点スキップ継続）** | 棚＋エイリアス＋C1マップ |
| 1系統 棚登録 | **OK** | B-T0 `…032818`／registry B-T1-2 |
| D〜PACKAGED dry_run | □ | `clasp push` 後 |

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。 |
| 2026-08-01 | **方針ロック反映**（複合複数・実装は実需後）。 |
| 2026-08-02 | **GROCERY実需**: FOOD_FISH_GROCERY 指紋・棚・MEAT→GROCERY・C1マップ。 |
