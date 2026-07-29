# SP-API スプシ橋渡し（人間手順）

**状態**: **v1.2／v1.2b／v1.2c／v1.3 実機合格**／**v1.2a 実装済**（2026-07-28／29）  
**承認**: [SHEET_BRIDGE](LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md)／[DRIVE_FETCH](LV4_SPAPI_DRIVE_FETCH_APPROVAL.md)／[APPROVED_EXPORT](LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md)／[CHECKBOX_EXPORT v1.2c](LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md)  
**書込本体**: [D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md](D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)

---

## 0. できること／できないこと

| できる | できない（別承認 v1.4） |
|--------|------------------------|
| 21-⑧ **子SKU＋出品CK（レ点）** → Drive CSV | GAS から SP-API PUT（[v1.4 API合格](LV4_SPAPI_GAS_PUT_APPROVAL.md)・SC最終更新は反映待ち） |
| 21-⑨ 承認①済 Amazon → Drive CSV | 親レ点のみで全子出品／行選択 |
| ローカル `--fetch-drive` → dry_run／prod | 全件ループ・在庫>0無人 |

**出品＝レ点（子）**。行選択は **完全廃止**。親行だけのレ点では出さない。

---

## 1. Script Properties

| キー | 既定 | 内容 |
|------|------|------|
| `APPROVAL_AMAZON_SPAPI_EXPORT_ENABLED` | **false** | `true` で 21-⑧／21-⑨ 有効 |
| `APPROVAL_AMAZON_SPAPI_EXPORT_MAX_ITEMS` | `5` | 超過は拒否 |
| `APPROVAL_AMAZON_SPAPI_EXPORT_FORCE_QTY_0` | **true** | 在庫0強制 |
| `APPROVAL_AMAZON_SPAPI_EXPORT_FOLDER_ID` | 空 | 空なら Lv4 GENERATED フォルダ流用 |

---

## 2. 手順（推奨・v1.3＋v1.2c）

1. `clasp push`（`AmazonSpapiExport.js`＋`コード.js`）  
2. Property `APPROVAL_AMAZON_SPAPI_EXPORT_ENABLED=true`  
3. **どちらか**  
   - **21-⑧**: マスタで **出品したい子行に出品CK** → CSV 出力（行選択不要）  
   - **21-⑨**: 最新承認①済 Amazon（子SKU＋ASIN）を一括 CSV  
4. ローカル:

```text
cd tools\spapi_listings_write
python -m pip install -r requirements.txt
```

`config.local.json` の `drive` を設定（初回のみ）:

- `credentials_path`: C1 の `../c1_hpc_packaged/secrets/credentials.json` 流用可  
- `folder_id`: 21-⑧／⑨ の出力フォルダ ID（空なら Drive 全体で `SPAPI_ITEMS` 最新を検索）  
- 初回 OAuth でブラウザ認可 → `secrets/token_drive.json` 生成  

```text
python spapi_listings_write.py --fetch-drive --mode dry_run
```

5. OKなら `allow_prod=true` → `--fetch-drive --mode prod`（または取得済み items.csv で prod）  
6. Property／`allow_prod` を **false** に戻す  

単独取得のみ:

```text
python spapi_fetch_drive_csv.py
```

---

## 2.1 旧手順（手動 DL・互換）

Drive CSV を手で `items.csv` に置き → `python spapi_listings_write.py --mode dry_run`

---

## 3. 列の取り方

| CSV列 | マスタ |
|-------|--------|
| sku | 子SKU（空なら親SKU）※21-⑧は子必須 |
| asin | ASINコード → 競合店ASIN → URL |
| price | 販売価格amazon（行→親） |
| quantity | 既定0（FORCE） |

21-⑧／21-⑨ とも **親行のみ（子SKU空）はスキップ**。ASIN／価格が無い行もスキップ。

---

## 4. 合格目安

### v1.2（済・当時は選択行）

- [x] 21-⑧ → Drive CSV  
- [x] ローカル dry_run／prod → SC（ride01）  
- [x] Property／allow_prod false  

### v1.2c（実機合格）

- [x] 子レ点のみで CSV 出る  
- [x] 親レ点のみでは出さない（社長完了扱い・2026-07-29。親行スキップは 21-⑨ でも `…oya` スキップを確認）  
- [x] 行選択なしで動作  

### v1.3／v1.2b（実機合格※v1.2a除く）

- [x] `--fetch-drive` で最新 CSV 取得 → dry_run／prod  
- [x] 21-⑨ で承認①済から CSV → dry_run／prod  
- [ ] cp932 CSV でも dry_run 可能（v1.2a・専用実機未）  

### 4.1 実機記録（2026-07-28・v1.2）

| 項目 | 結果 |
|------|------|
| アルコール | `…0139-ride01`／`B0091G3AHY`／1000／0 |
| 安眠 | `…0924-ride01`／`B00A0J0D30`／1000／0 |
| 発汗 | `…48s11`／`B07YND44VN` |
| 注意 | 制限付きノーブランド ASIN は code 5886 |

### 4.2 実機記録（2026-07-28・v1.2c＋v1.3）

| 項目 | 結果 |
|------|------|
| Drive CSV | `SPAPI_EXPORT_20260728_235455_52100a_SPAPI_ITEMS.csv`（子レ点1行） |
| SKU／ASIN | `lifec-4560151300924-48s11`／`B00A0J0D30` |
| dry_run | VALID・ok=1 |
| prod | ACCEPTED・ok=1（レポート `…145928`） |
| 経路 | 21-⑧子レ点 → `--fetch-drive` → prod |

### 4.3 実機記録（2026-07-29・v1.2b）

| 項目 | 結果 |
|------|------|
| Drive CSV | `SPAPI_EXPORT_APPR_20260729_001815_127e8a_SPAPI_ITEMS.csv` |
| batch | `A1_20260727_224939_b7a053`（件数1／スキップ1＝親 `…0832-oya`） |
| SKU／ASIN | `lifec-4560151300832-48s11`／`B07YND44VN` |
| dry_run | VALID・ok=1（GET 200） |
| prod | ACCEPTED・ok=1（レポート `…152611`） |
| 経路 | 21-⑨承認①済 → `--fetch-drive` → prod |

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-29 | **v1.2b 実機合格**＋親レ点0件を完了扱い。21-⑨ APPR→prod（`…0832-48s11`）。 |
| 2026-07-28 | **v1.2c／v1.3 実機合格**: 子レ点→Drive→fetch-drive dry_run／prod（安眠 `…48s11`）。 |
| 2026-07-28 | **v1.2c**: 21-⑧＝子SKUレ点のみ。選択行完全廃止。親レ点のみ出さない。 |
| 2026-07-28 | 21-⑧/⑨: スキップ理由をダイアログに列名付きで表示。 |
| 2026-07-28 | **v1.3** Drive自動取得／**v1.2b** 21-⑨／**v1.2a** 文字コード。 |
| 2026-07-28 | **実機合格**記録。ride01 アルコール／安眠・1000/0。 |
| 2026-07-28 | v1.2 初版。選択行→Drive CSV（のち v1.2c で廃止）。 |
