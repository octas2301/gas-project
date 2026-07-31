# C1 — PACKAGED HPC／SEASONING（人間手順）

**状態**: **HPC C1-1b SC合格／SEASONING 七味実データPACKAGED再送済**（SC停止中・再送結果待ち）
**承認**: C1実装承認／**C1-1b実装承認**  
**正本**: [D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)  
**列マップ**: [D_MENU_C1_MASTER_HPC_COLUMN_MAP.md](D_MENU_C1_MASTER_HPC_COLUMN_MAP.md)／[SEASONING](D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md)  
**ツール**: `tools/c1_hpc_packaged/c1_packaged.py`  
**入力取得（案A）**: `c1_fetch_inputs.py`／[D_MENU_C1_FETCH_INPUTS_REQUIREMENTS.md](D_MENU_C1_FETCH_INPUTS_REQUIREMENTS.md)  
**安眠外出先メモ**: [D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md](D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)

---

## 0b. SEASONING（七味）実機メモ（2026-07-31）

| 項目 | 値 |
|------|-----|
| subBatch | `CK_daba393f8055_B2` |
| 最新PACKAGED | `…_20260731_032459.xlsm`（Drive 03） |
| 初回サマリ | 処理2／成功0／その他エラー2。**99016**（KW5枠）＋**100521**（審査） |
| 再送 | KW1枠修正版をUP済。結果待ち |
| SC在庫 | 親`B0HC9S8PRN`／子`B0HC9RRCBP` は**停止中** |
| 必須既定 | 粉末／グラム／KW1枠。詳細は [列マップ](D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md) |

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
   - **`generated_csv`**: `{subBatchId}_GENERATED.csv` 形式（**`relax_GENERATED` 固定は使わない**）  
   - **`sub_batch_id`** または `c1_packaged.py --sub-batch …` で置換  

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

`{subBatchId}` 利用時:

```text
python c1_packaged.py --config config.local.json --mode dry_run --sub-batch A1_20260726_225610_4f0558_B2
```

（`config.sub_batch_id` があれば `--sub-batch` 省略可）

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

**安眠（2026-07-27）**: **SC合格＋E-5済**。  
- 送信: `relax_PACKAGED_HPC_lifec-4560151300924-oya.xlsm`（Batch `182816020660`・接頭辞 `relax` は設定残り）  
- サマリ: Downloads `relax_PACKAGED_HPC_lifec-4560151300924-oya-processing-summary.xlsm`  
- 処理SKU=2／成功=2／失敗=0／警告=0／エラー総数=0  
- 親=`lifec-4560151300924-oya`・子=`lifec-4560151300924-19s124`（テンプレ行とも「成功」）  
- E-5／ログ: `A1_20260726_225610_4f0558_B2`（本線正式 GENERATED）

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
| SC（安眠・本線） | **合格**（2026-07-27）`relax_PACKAGED_HPC_lifec-4560151300924-oya` → Downloads `…-processing-summary.xlsm`。処理SKU=2／成功=2／失敗=0／警告=0／エラー総数=0。親=`…0924-oya`・子=`…19s124`。**E-5済** `A1_20260726_225610_4f0558_B2` |
| 21-③ | **スキップ（代替記録）**（0405手置きのみ）。安眠は **E-5** で同一 subBatchId 記録済。 |

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

## 8. SEASONING（吉野家 七味唐辛子）

列マップ: [D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md](D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md)

1. Dで吉野家の新規のみを実行し、`*_GENERATED.csv` を作る
2. 最新マスタCSVを取得する
3. `config.local.json` の3項目を次へ変更する

```json
{
  "template_path": "C:/Users/takuy/Downloads/FOOD_HERB_SEASONING_FISH_VEGETABLE.xlsm",
  "column_map_path": "food_seasoning_column_map.json",
  "fingerprint_path": "fingerprints/food_seasoning_header_r3_r5.json"
}
```

4. `--mode dry_run`。`profile=SEASONING`、指紋match、acceptedParentsありを確認
5. 出力行で PT=`SEASONING`、ブラウズ=`唐辛子 (2430212051)`、在庫=0、タックスがマスタ値であることを確認
6. ブランド=`ノーブランド品`／GTIN免除証跡／メーカー／原料を人間確認
7. 合格後だけ `--mode prod` → SC事前チェック

現在は手順4の機械スモークまで合格。最新マスタ読取では対象親
`sanky-4538872180149-oya`／レ点子`…B01N5A6ESU-19s13`を確認済み。
実データDRY_RUN前の残りは次の3点。

- D新規のGENERATED作成
- 親子の `Amazon MAIN URL`（現状空）
- 子の数値 `定価、市場価格` または `list_price_override_map`（現状、親は調査文言・子は空）

原料・メーカー・タックス・原産国・バリエーション値はマスタに存在する。

---

## 9. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-30 | **C1-1c SEASONING**: 新321列テンプレの指紋・行8開始・唐辛子ノードで機械DRY_RUN合格。 |
| 2026-07-27 | **C1-clean**: `generated_csv`=`{subBatchId}_GENERATED.csv`／`sub_batch_id`・`--sub-batch`。`relax` 固定廃止。 |
| 2026-07-27 | **安眠 SC合格＋E-5**: 処理2／成功2／エラー0。サマリ=Downloads。subBatchId=`A1_20260726_225610_4f0558_B2`。 |
| 2026-07-27 | **安眠** C1 prod→SC送信（キュー）。外出先用 [ANMIN_REMOTE_CHECKLIST](D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)。 |
| 2026-07-27 | **案A fetch**（`c1_fetch_inputs.py`）。GENERATED＋マスタCSV自動取得。 |
| 2026-07-26 | 21-③スキップ代替記録（C1手置き・ログにGENERATEDなし）。SC合格でマイルストーン区切り可。 |
| 2026-07-26 | **SC合格**（relax・SKU2/2・エラー0）。次＝21-③。 |
| 2026-07-26 | 定価＝数値／式専用・調査は▼マスタ(市場価格調査)のみ（確定）。 |
| 2026-07-26 | Sheets全件CSV対応（ヘッダー検出・親継承）。A経路DRY_RUN確認。 |
| 2026-07-26 | C1-1b（master_csv・必須列）。 |
| 2026-07-26 | 骨格パイプライン合格・SC不足記録。 |
| 2026-07-26 | 初版。 |
