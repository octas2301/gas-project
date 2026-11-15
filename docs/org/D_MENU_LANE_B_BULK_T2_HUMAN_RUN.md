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
| HERB.xlsm | 単体は自動本線化しない。PT=`HERB` は複合で本線 |
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

### 0c. 実需記録 — PACKAGED_SOUP_AND_STEW（食べるおだし・2026-09-18）

| 項目 | 値 |
|------|-----|
| 純正 | `PACKAGED_SOUP_AND_STEW.xlsm`（単体） |
| 指紋 sha | `c6372856adbb32e99e644a336e1ae873731845e2d56316af0bbd2f5d8899ab5a`（行3–5・maxCol293） |
| C1マップ | `packaged_soup_and_stew_column_map.json`（項目名はGROCERY流用＋**データ定義で必須特定**。単位既定=袋） |
| 実機 prod | `CK_a7a907ed2f61_B1_PACKAGED_PACKAGED_SOUP_AND_STEW_n2_20260918_023529.xlsm`（かつお＋まぐろ。梱包寸法は例を書かない） |

E③ は GENERATED まで。PACKAGED は Cursor 手渡し→B-T1。**09 に `{PT}.xlsm` があれば** B-T1 が指紋＋列マップを自動登録して 03 に出す。必須列は **そのバルクの「データ定義」**。「必須」だけ常時埋める。「条件付きで必須」はマスタ／コピー／安全既定があるときだけ。寸法・重量・物流・FBA階層は必須から除外。データ定義の例は値にしない。GROCERYの黒セル省略は缶飯専用。ブラウズの書き方は **そのファイルのプルダウン実値**（缶飯の経路表記を全件に流さない）。親行にはサイズ・在庫・親番号を書かない。

---

## 1. 人間チェックリスト（実需が出たら）

1. P4b等で PT／browse を確定  
2. Agent に棚引き（`--product-type`）  
3. `DL_NOT_NEEDED` → **B-T1 へ**（B-T2不要）  
4. `DL_REQUIRED` → SCで **必要なら別系統の複合**（または単体）を DL → **09** → E③  
5. B-T1 prod（`--product-type ALL`）。09 の xlsm から指紋＋列マップを自動登録 → 03 PACKAGED → 人間が SC UP  
6. 必須列 MISS／09 に xlsm 無しのときだけ停止（人手で列マップを発明しない）  

---

## 2. Agent（実装後・仮）

```text
cd tools\c1_hpc_packaged
python c1_bulk_fill_by_name.py --mode prod --product-type ALL ^
  --generated-csv "（GENERATED.csv）" ^
  --master-csv "（マスタCSV）" ^
  --output-dir "G:/マイドライブ/04.amazonカタログ作成（CSV一括UL）/03.SCへ上げる完成xlsm（人間が置く／DL）" ^
  --parent-sku-filter （親SKU）
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
| 2026-09-18 | **ブラウズ＝そのファイルのプルダウン**。親にサイズ・在庫・親番号を書かない。マスタ生値で専用列を上書きしない。 |
| 2026-09-18 | **HERB**: 単体 xlsm は非登録。PT=HERB は複合で PACKAGED。 |
| 2026-09-18 | **バルクごと必須埋め**: データ定義の必須／条件付き必須を読み、値ソースがある列だけ書く。GROCERY黒セルomitは缶飯専用。 |
| 2026-09-18 | **auto-onboard**: 09+E③のあと B-T1 が指紋＋項目名で新PTを本線化。必須はデータ定義。 |
| 2026-09-18 | **PACKAGED_SOUP_AND_STEW 実需**: 指紋・棚・C1マップ。かつお＋まぐろ prod。 |
| 2026-08-02 | **GROCERY実需**: FOOD_FISH_GROCERY 指紋・棚・MEAT→GROCERY・C1マップ。 |
