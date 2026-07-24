# D×Amazon U3（薄いファサード）— 人間手順

**状態**: **U3 v1 実機合格**（`clasp push` 済・2026-07-25）  
**正本**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md) §5・§6.1・§6.1.1  
**承認**: 2026-07-25「U3 v1 承認」（amazon／full_amazon・21-①呼出のみ・トリガーに Amazon 非搭載）

**実機証跡（スモーク）**:
- `runId=LV4_20260725_072635_948892`
- course=`amazon` → Lv4確認 → 実行結果（`idempotentBlocked=1`・親0＝既存バッチ冪等で想定内）→ Da完了ダイアログ
- 事後: `APPROVAL_AMAZON_LV4_ENABLED=false`

---

## 1. 事前（Script Properties）

| Key | 値 |
|-----|-----|
| `APPROVAL_AMAZON_LV4_ENABLED` | `true`（実行時のみ。終わったら **false**） |
| `APPROVAL_AMAZON_LV4_TRACK` | `B`（M1）など |
| `AMAZON_DRIVE_IMAGE_FOLDER_ID` | 空なら既定の Drive `02` |
| `AMAZON_DRIVE_R2_POC_SKU` | **任意**。設定時はその `{SKU}.MAIN.jpg` が無いと **停止** |

オフ推奨（触らない）: `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`

---

## 2. clasp push（ローカル人間）

変更ファイル例:

- `コード.js`（Dモーダル・`runBatchExportAmazonFacade`）
- `AmazonApprovalExport.js`（`menuApprovalAmazonLv4Run` 戻り値）

```text
clasp push
```

Cloud Agent からは push しない。

---

## 3. 実行

1. スプレッドシートメニュー **D. 画像＋出品全モール一括アップロード（選択可）**
2. コース:
   - **Amazonのみ（Da・21-①呼出）** … 即時。裏は `menuApprovalAmazonLv4Run`
   - **フル → Amazon Da** … 同期で楽天→Yahooのあと Amazon Da（**Amazonはトリガーに載せない**）
3. Lv4 確認ダイアログ → 実行結果 → **Da 完了ダイアログ**（PACKAGED／手ZIP／SCの2手／21-③はZ／Property off）

楽天のみ・Yahooのみ・フル（Amazonなし）は従来どおり **約1分後トリガー**。

---

## 4. Da 後の人間作業

1. PACKAGED（Drive `03`・C0）  
2. 手 ZIP（Drive `04`）  
3. SCの2手（xlsm UP ＋ ZIP UP）  
4. Z → **21-③**  
5. `APPROVAL_AMAZON_LV4_ENABLED=false`

---

## 5. 戻し方

- Property を false  
- `git revert`（当該コミット）  
- 必要なら再 `clasp push`
