# C1 — PACKAGED HPC（人間手順）

**状態**: **C1-1b SC合格**（2026-07-26・relax）。21-③は手置き経路のため代替記録  
**承認**: C1実装承認／**C1-1b実装承認**  
**正本**: [D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)  
**列マップ**: [D_MENU_C1_MASTER_HPC_COLUMN_MAP.md](D_MENU_C1_MASTER_HPC_COLUMN_MAP.md)  
**ツール**: `tools/c1_hpc_packaged/c1_packaged.py`  
**入力取得（案A）**: `c1_fetch_inputs.py`／[D_MENU_C1_FETCH_INPUTS_REQUIREMENTS.md](D_MENU_C1_FETCH_INPUTS_REQUIREMENTS.md)

---

## 0. 固定値（C1-1b）

| 項目 | 固定 |
|------|------|
| 言語 | Python 3 + openpyxl（`keep_vba=True`） |
| 入力 | GENERATED CSV ＋ **`master_csv`（マスタ列名ヘッダー）** |
| タックス | マスタ `商品タックスコード`（固定禁止） |
| URL空／必須マスタ欠落 | 親SKU一式除外 |
| トグル | `--mode dry_run` / `--mode prod` |

---

## 1. 準備

### 1a. 手ダウンロード（従来）

1. 純正06 `HEALTH_PERSONAL_CARE.xlsm` をローカルへ  
2. `*_GENERATED.csv`  
3. **マスタ書き出しCSV**（A: 全マスタCSV可）→ `master_csv`  
   - Sheets「▼商品マスタ」の **ファイル→ダウンロード→CSV**。Excelで開いて上書きしない  
   - 先頭に注記行があっても可（ツールが `親SKU`/`子SKU` 行を検出）  
   - **親行にも `Amazon MAIN URL` が必要**（C1）。**U4／E-2 後は子→親へ空欄自動コピー**（2026-07-27〜）  
4. `pip install -r requirements.txt`  
5. `config.local.json` でパス設定（`parent_sku_filter` で対象親を絞る）  

### 1b. 入力自動取得（案A・推奨）

1. 純正06テンプレは従来どおりローカルへ  
2. GCP で OAuth「デスクトップ」クライアント JSON → `tools/c1_hpc_packaged/secrets/credentials.json`  
3. `config.local.json` に `spreadsheet_id` と `fetch`（`config.example.json` 参照）  
4.

```text
cd tools\c1_hpc_packaged
pip install -r requirements.txt
python c1_fetch_inputs.py --config config.local.json --latest
```

subBatchId 指定例:

```text
python c1_fetch_inputs.py --config config.local.json --sub-batch A1_20260726_225610_4f0558_B2
```

5. 続けて §3 DRY_RUN → §4 prod  
6. `secrets/`・`token.json` は Git 禁止（読取スコープのみ）

スモーク例: `testdata/sample_MASTER.csv`／実データ例: `…/input/master_export.csv`

---

## 2. 指紋

```text
cd tools\c1_hpc_packaged
python c1_packaged.py --config config.local.json --write-fingerprint
```

（純正06に対して）

---

## 3. DRY_RUN

```text
python c1_packaged.py --config config.local.json --mode dry_run
```

- [ ] `acceptedParents` あり  
- [ ] mapping の `taxCode` がマスタ値  
- [ ] xlsm にメーカー名・説明／仕様・輸入・感熱・原産国・危険物・液体・タックス等  

---

## 4. 本番 → Drive 03

```text
python c1_packaged.py --config config.local.json --mode prod
```

---

## 5. SC（**未送信SKU推奨**）

既存掲載への再UPは避ける。事前チェック通過後に送信→21-③（または E-5）。

**安眠（2026-07-27）**: prod 済 → SC 送信済（Batch `182816020660`・ファイル名接頭辞は `relax` だが中身は0924）。外出先確認は [D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md](D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)。E-5 ID=`A1_20260726_225610_4f0558_B2`。

---

## 7. 実機合格記録

| 項目 | 結果 |
|------|------|
| 指紋 | **OK** |
| DRY_RUN（骨格） | **OK** |
| DRY_RUN（C1-1b） | **OK**（サンプルmaster） |
| DRY_RUN（全マスタCSV A） | **OK**（2026-07-26・`master_export.csv`・tax=`A_GEN_STANDARD`・親URLは確認用override） |
| 定価の都度修正 | **子SKU行の `定価、市場価格`（数字）を優先**。親行は子の値を使用。足りなければ `list_price_override_map`。フルCSV再DLは骨格変更時のみ |
| SC | **合格**（2026-07-26）`relax_PACKAGED_HPC_lifec-4560151300405-oya_20260726_042714` → Drive05 `…-processing-summary.xlsm`。処理SKU=2／成功=2／失敗=0／警告=0／エラー総数=0。親=`lifec-4560151300405-oya`・子=`…-16s184` |
| 21-③ | **スキップ（代替記録）**。C1手置き `relax_GENERATED` のため `▼Lv4実行ログ(Amazon)` に対象 GENERATED 行なし（`subBatchId=relax` は拒否）。**SC合格＋Drive05レポート**をもって完了記録の代替とする。本線は次回から 21-①正式 GENERATED → 同一 subBatchId で 21-③ |

### 定価・市場価格調査の役割分担（2026-07-26確定）

| 列 | 役割 |
|----|------|
| **`定価、市場価格`** | **数値／計算式専用**（C1の税込み参考価格・人間の定価）。調査文言を入れない |
| **`▼マスタ(市場価格調査)`** | AI／リサーチの調査メモ・文言の置き場 |

AI同期は **`定価、市場価格` へ転記しない**（式の上書き防止）。

### 定価を計算式にしたときの注意（他メニュー）

| メニュー／処理 | 影響 |
|----------------|------|
| AI同期・市場価格調査 | 調査は **▼マスタ(市場価格調査) のみ**（上記確定後） |
| 楽天／Yahoo 出品 | `定価、市場価格` を直接は使わない（影響小） |
| C1 | 子SKU行の数字を優先。調査メモ文言は無視 |

---

## 8. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-27 | **安眠** C1 prod→SC送信（キュー）。外出先用 [ANMIN_REMOTE_CHECKLIST](D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)。 |
| 2026-07-27 | **案A fetch**（`c1_fetch_inputs.py`）。GENERATED＋マスタCSV自動取得。 |
| 2026-07-26 | 21-③スキップ代替記録（C1手置き・ログにGENERATEDなし）。SC合格でマイルストーン区切り可。 |
| 2026-07-26 | **SC合格**（relax・SKU2/2・エラー0）。次＝21-③。 |
| 2026-07-26 | 定価＝数値／式専用・調査は▼マスタ(市場価格調査)のみ（確定）。 |
| 2026-07-26 | Sheets全件CSV対応（ヘッダー検出・親継承）。A経路DRY_RUN確認。 |
| 2026-07-26 | C1-1b（master_csv・必須列）。 |
| 2026-07-26 | 骨格パイプライン合格・SC不足記録。 |
| 2026-07-26 | 初版。 |
