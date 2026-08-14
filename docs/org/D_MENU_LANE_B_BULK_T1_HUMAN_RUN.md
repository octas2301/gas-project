# レーンB — 純正バルク B-T1（項目名マッピング）人間手順

**状態**: **実装済／スモーク合格**（2026-08-01）— dry_run のみ  
**承認**: [LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md](LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md)  
**前提**: [B-T0 HUMAN_RUN](D_MENU_LANE_B_BULK_T0_HUMAN_RUN.md)／[P4b](LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)  
**実行**: **ローカル Python は Agent モード**

---

## 0. 方針（要約）

| 項目 | 内容 |
|------|------|
| テンプレ | **複合**＋棚（06） |
| PT | **`SEASONING`**（当面） |
| 列解決 | **項目名**（`xlsm_header_aliases`） |
| 人間入力 | **マスタのみ** |
| 再DL | 同型（PT＋指紋）が棚にあれば **不要** |
| `HERB.xlsm` | 本線外 |
| prod | **実装済** — [PROD承認](LV4_LANE_B_BULK_TEMPLATE_T1_PROD_APPROVAL.md)／[PROD HUMAN_RUN](D_MENU_LANE_B_BULK_T1_PROD_HUMAN_RUN.md) |

---

## 1. 定常フロー

```text
① スプシでカテゴリ／PTを決める（P4b等）
② Agent: 棚引き
     ・あり → 「SCダウンロード不要」→ 06から充填へ
     ・なし → 「SCでバルクをDLし 09 へ保存」→ 人間が09へ → B-T0 → 06登録後に充填
③ dry_run 充填（項目名マップ）→ 03／05
④（[prod第2段](LV4_LANE_B_BULK_TEMPLATE_T1_PROD_APPROVAL.md)承認後）prod → 人間が SC UP
```

---

## 2. Agent コマンド

### 2.1 棚引きのみ

```text
cd tools\c1_hpc_packaged
python c1_bulk_shelf_lookup.py --product-type SEASONING
```

- exit 0 = `DL_NOT_NEEDED`  
- exit 2 = `DL_REQUIRED`  
- レポート: **05** `B_T1_SHELF_*`

### 2.2 棚引き＋項目名 dry_run 充填

```text
cd tools\c1_hpc_packaged
python c1_bulk_fill_by_name.py --product-type SEASONING ^
  --generated-csv "（GENERATED.csv のパス）" ^
  --master-csv "（マスタCSVのパス）" ^
  --output-dir "G:/マイドライブ/04.amazonカタログ作成（CSV一括UL）/03.SCへ上げる完成xlsm（人間が置く／DL）" ^
  --report-dir "G:/マイドライブ/04.amazonカタログ作成（CSV一括UL）/05.SC処理結果・ログ退避（人間）"
```

---

## 3. 人間チェックリスト

| 状況 | やること |
|------|----------|
| 「DL不要」 | 09 に落とさない。03 DRYRUN／05 不足・ギャップを確認 |
| 「09へDL要求」 | SC→**09** → Agent に B-T0 → 06登録後に再依頼 |
| 充填後 | **05** の `*_NAME_MAP_GAPS.json`／C1_REPORT を確認 |

---

## 4. 合格記録

| 段階 | 結果 | メモ |
|------|------|------|
| 方針ロック（棚引き含む） | **OK** | 2026-08-01 |
| 三点スキップ | **OK** | 2026-08-01 |
| 実装承認（dry_run） | **OK** | 2026-08-01 |
| 棚あり＝DL不要 | **OK** | `B_T1_SHELF_SEASONING_20260801_084530` |
| 棚なし＝DL要求 | **OK** | HERB → `registry_miss`／`B_T1_SHELF_HERB_…` |
| 項目名解決 | **OK** | 66/66 HIT・legacy差0 |
| dry_run 充填 | **OK** | `B_T1_FILL_SEASONING_20260801_084627` → `…_DRYRUN.xlsm` |

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。棚引き追記。 |
| 2026-08-01 | **実装＋スモーク**（棚／項目名／dry_run）。 |
