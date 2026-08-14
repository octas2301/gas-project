# レーンB — B-T1 prod 第2段 人間手順

**状態**: **実装済／スモーク合格**（2026-08-01）  
**承認**: [LV4_LANE_B_BULK_TEMPLATE_T1_PROD_APPROVAL.md](LV4_LANE_B_BULK_TEMPLATE_T1_PROD_APPROVAL.md)  
**親手順**: [D_MENU_LANE_B_BULK_T1_HUMAN_RUN.md](D_MENU_LANE_B_BULK_T1_HUMAN_RUN.md)（dry_run）  
**実行**: **ローカル Python は Agent モード**。SC UL は人間。

---

## 0. 要約

| 項目 | 内容 |
|------|------|
| コマンド | `c1_bulk_fill_by_name.py --mode prod` |
| 出力 | 03 に **`_DRYRUN` なし** PACKAGED |
| 必須 | 棚 `DL_NOT_NEEDED`＋指紋一致＋必須列解決 |
| やらない | SC自動UL・06破壊・03上書き |

---

## 1. 推奨順序

1. （任意）`--mode dry_run` で確認  
2. Agent に **prod**（親SKUフィルタ推奨）  
3. 03／05 を確認  
4. 人間が SC 事前チェック → UL  

---

## 2. Agent コマンド

```text
cd tools\c1_hpc_packaged
python c1_bulk_fill_by_name.py --mode prod --product-type SEASONING ^
  --generated-csv "（GENERATED.csv）" ^
  --master-csv "（マスタCSV）" ^
  --output-dir "G:/マイドライブ/04.amazonカタログ作成（CSV一括UL）/03.SCへ上げる完成xlsm（人間が置く／DL）" ^
  --report-dir "G:/マイドライブ/04.amazonカタログ作成（CSV一括UL）/05.SC処理結果・ログ退避（人間）" ^
  --parent-sku-filter （親SKU）
```

---

## 3. 合格記録

| 段階 | 結果 | メモ |
|------|------|------|
| 方針ロック | **OK** | 2026-08-01 |
| 三点スキップ | **OK** | 2026-08-01 |
| 実装承認 | **OK** | 2026-08-01 |
| prod 実機 | **OK** | `B_T1_FILL_PROD_SEASONING_20260801_085431` → `…_oya_20260801_085442.xlsm` |
| 棚なし拒否 | **OK** | HERB exit 2 |
| 指紋不一致拒否 | **OK** | exit 3／`…085444` |

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。 |
| 2026-08-01 | **実装＋スモーク**。 |
