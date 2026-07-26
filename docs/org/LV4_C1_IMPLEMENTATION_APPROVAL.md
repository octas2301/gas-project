# Amazon C1 PACKAGED — 承認パッケージ

**日付**: 2026-07-26  
**要件正本**: [D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)  
**多数決**: [D_MENU_C1_THREE_REVIEW_MAJORITY.md](D_MENU_C1_THREE_REVIEW_MAJORITY.md)  
**HUMAN_RUN**: [D_MENU_C1_HUMAN_RUN.md](D_MENU_C1_HUMAN_RUN.md)  
**親**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)／[LV4_R2_IMAGE_PIPELINE_POC.md](LV4_R2_IMAGE_PIPELINE_POC.md) §5  
**状態**: **C1実装承認済＋C1-1b実装承認済（2026-07-26）**。次＝未送信SKUでSC  

依存: U2／U3／U4 実機合格。画像は U4 `Amazon MAIN URL`。必須属性は **マスタCSV**。

---

## 1. 変更ファイル一覧（C1-1b）

| 種別 | パス | 内容 |
|------|------|------|
| 更新 | `tools/c1_hpc_packaged/c1_packaged.py` | master_csv併読・必須列・タックスはマスタ |
| 更新 | `tools/c1_hpc_packaged/hpc_column_map.json` | C1-1b列＋master_columns |
| 新規 | `tools/c1_hpc_packaged/testdata/sample_MASTER.csv` | スモーク用マスタ |
| 更新 | docs（HUMAN_RUN／要件／CURRENT_PHASE等） | 状態 |
| **触らない** | GAS／楽天／Yahoo／マスタ本体 | 聖域（CSVは読取のみ） |

---

## 2. 概要

- 本格 PACKAGED（HPC・新規行一式）・**ローカル Python**  
- URL空＝スキップ＋親一式除外（フォールバック禁止）  
- テンプレ指紋不一致＝本番停止  
- 06読取コピーのみ／03新規のみ／マスタ非書込  

---

## 3. 想定リスク

| リスク | 緩和 |
|--------|------|
| テンプレ更新で列ずれ | 指紋＋本番停止 |
| openpyxlでVBA破損 | `keep_vba=True`・HUMAN_RUNでSC確認 |
| size表記不足 | `size_map`／GENERATED拡張 |
| 06破壊 | コピー後編集・原本パスは読取のみ |
| 聖域 | GAS非改変・マスタ非書込 |

---

## 4. ゲート順

1. ~~方針ロック~~ **済**  
2. ~~要件起草~~ **済**  
3. ~~三点レビュー~~ **済**  
4. ~~多数決反映~~ **済**  
5. ~~実装承認~~ **済（2026-07-26「C1実装を承認」）**  
6. **HUMAN_RUN**（次）  

---

## 5. 社長向け一言

> C1実装承認済。ローカル `tools/c1_hpc_packaged` で DRY_RUN→本番→03へ。次はHUMAN_RUN実機。

---

## 6. 実装承認前チェックリスト（完了記録）

### 6.1 要件・ゲート（docs）

- [x] 三点レビュー完了・多数決メモ作成  
- [x] 社長決定3項を要件に反映  
- [x] 変更ファイル一覧再提示＋**実装承認**（本§1）  
- [x] HUMAN_RUN ファイル名確定: `D_MENU_C1_HUMAN_RUN.md`  

### 6.2 技術スパイク

- [x] 言語: **Python + openpyxl**（`keep_vba=True`）  
- [x] 行5非書込・行6サンプル維持・行7〜データ（コード仕様）  
- [x] 指紋範囲: **行3–5** → `fingerprints/hpc_header_r3_r5.json`  
- [x] Drive到達: **同期／手動コピー**（APIなし）  

### 6.3 入出力契約

- [x] GENERATED: `*_GENERATED.csv`（subBatchIdはファイル名）  
- [x] 対象: CSV親一式＋任意 `parent_sku_filter`  
- [x] トグル: `--mode dry_run|prod`  
- [x] 置き場: Git=`tools/…`（claspignore）／成果xlsm=Downloads配下  
- [x] パスは `config.local.json`（Git外推奨）  

### 6.4 聖域・安全

- [x] 06読取コピーのみ・03上書き禁止・マスタ非書込  
- [x] 楽天／Yahoo／B統合非干渉（GAS未変更）  
- [x] 親一式除外＋URLフォールバック禁止  
- [x] 指紋不一致時は本番ファイル非作成  

### 6.5 承認

**承認文言**: 「C1実装を承認」（2026-07-26）

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-26 | **C1-1b実装承認**・master_csv必須列。次＝未送信SKUでSC。 |
| 2026-07-26 | 実装承認・§1確定・§6完了・ツール実装。 |
| 2026-07-26 | 三点反映後に §6 追加。 |
| 2026-07-26 | 初版（三点前）。 |
