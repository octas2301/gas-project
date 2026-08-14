# D×Amazon U2（C 画像）— 人間手順

**状態**: **実機合格**（2026-07-25）  
**正本**: [D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md](D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md)  
**承認**: U2 実装 v1（案α・マスタ永続・`02` コピー）／C を子レ点対象に変更

---

## 0. 実機合格記録（2026-07-25）

| 項目 | 結果 |
|------|------|
| C-Amazon② | OK（`07` 候補読込） |
| C-Amazon③ | OK（更新行=2・列追加=`Amazon画像モード`,`Amazon MAIN 参照`,`Amazon PT 参照`） |
| マスタ | `REUSE_RAKUTEN` ＋ MAIN Drive ID 保存確認 |
| C-Amazon④ | OK（`MAIN成功=1` / `PT成功=0` / `失敗=0`）。`REUSE` のため PT=0 は正常 |
| Drive `02` | `lifec-4560151300139-16s100.MAIN.jpg` 確認 |
| （任意）D → Amazonのみ | 実行到達。`runId=LV4_20260725_094425_914290`・`idempotentBlocked=1`・親成功=0＝**既存バッチ冪等で想定内**（U3と同型）。Da 完了ダイアログ表示 |

**Property 例**: `AMAZON_IMAGE_CANDIDATE_FOLDER_ID=1IRWfwLtZOoacDmIUWWbuLBC9xcMbjQjf`（Drive `07`）

**運用後**: `AMAZON_IMAGE_U2_ENABLED=false`（および使用していれば `APPROVAL_AMAZON_LV4_ENABLED=false`）

---

## 1. Script Properties

| Key | 値 |
|-----|-----|
| `AMAZON_IMAGE_U2_ENABLED` | `true`（実行時のみ。終わったら **false**） |
| `AMAZON_IMAGE_CANDIDATE_FOLDER_ID` | Drive `07.Amazon白抜き候補（入力）` の ID（例: `1IRWfwLtZOoacDmIUWWbuLBC9xcMbjQjf`）。楽天ソースと分離 |
| `AMAZON_DRIVE_IMAGE_FOLDER_ID` | Drive `02`（空なら既定 ID・出口のみ） |
| `AMAZON_IMAGE_CANDIDATE_ARCHIVE_ENABLED` | 未設定＝**有効**。④で02コピー成功した候補だけ `07/アップロード済み画像` へ退避。`false` でOFF |

オフ推奨: `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`

---

## 2. clasp push

```text
clasp push
```

含む: `AmazonImageMatrixExport.js` / `コード.js`（メニュー・C フック・**子レ点選定**）

---

## 3. C の出品対象（選定規則）

| マスタのレ点 | C に出る行 |
|--------------|------------|
| **子に1つ以上レ点** | その親行＋**レ点付き子だけ**（兄弟セットは出ない） |
| **子レ点なし・親のみレ点** | 従来どおり親＋**全子**（楽天運用互換） |

Amazon 試験は **対象セットの子だけにレ点**を付ける。

---

## 4. 操作順

**本線（推奨）**: [C. 画像コース HUMAN_RUN](D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)（C-1 Amazon → ドラッグ → C-2）

### 分割（Z → C-Amazon 互換）

1. **対象の子SKU行に出品CK**（推奨）。親のみレ点は全子が出るので注意  
2. 白抜きを Drive **`07`**（候補）へ置く。**`02` は出口のみ（日常は手で触らない。手置きは復旧例外）**  
3. Z → C-Amazon → 旧マトリクス生成 または **C-1**  
4. ①〜④を必要に応じて分割実行（本線は Cコース。④成功分は自動退避）  
5. （任意）D → Amazonのみ  

### 再生成後
C をやり直すと sheet は消える → マスタが正。再度 C または C-Amazon① で **復元**。

### 「画像が見つかりません」／旧商品が残る場合
- `AMAZON_IMAGE_U2_ENABLED=true` を確認して C を再実行する。U2無効時は、楽天従来契約どおり画像0件で停止する
- C-Amazon②で「マッチングsheetが現在のレ点対象と一致しません」が出た場合、③・④へ進まず C を再実行する
- 旧商品のsheetへ新商品の候補を並べた状態では、誤保存防止のため **③・④を実行しない**

### 非退行確認
- 楽天青枠（列6–25）からの R-Cabinet アップが従来どおり動くこと  
- Amazon 枠は列76以降であること  
- 親のみレ点の従来 C 運用が壊れていないこと  

---

## 5. 終わったら

`AMAZON_IMAGE_U2_ENABLED=false`

---

## 6. 戻し方

- Property false  
- `git revert`／新規 js 削除＋メニュー削除  
- C 選定のみ戻す場合: `generateAiImageMatrix` の商品構造ループを親レ点＋全子に戻す  

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-25 | 実機合格記録。 |
| 2026-08-01 | P1: 02は日常手操作禁止・復旧例外を明記。 |
