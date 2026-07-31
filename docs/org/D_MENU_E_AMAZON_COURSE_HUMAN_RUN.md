# E. Amazon出品コース（一時）— 人間手順

**状態**: **実装済**（要 clasp push）  
**性質**: 一時ファサード。裏は既存 C-Amazon／18／21／U4 のみ。将来は **D** に吸収予定。  
**本線**: Amazon の起点は **D**（新規カタログ／既存相乗りは D ラジオ。[D入口承認](LV4_SPAPI_D_ENTRY_APPROVAL.md)）。E はテスト・段階実行用。  
**分割逃げ道**: トップ C-Amazon①〜④／Z→18・21（削除しない）

---

## 0. 実行順（人が押すのはこれだけ）

```text
E-0 前提チェック
 → C（SKU入れ替え時）→ E-1 →【人間】MAINドラッグ
 → E-2 ③→④→U4（URL）
 → E-3 承認候補  →【人間】Web承認①
 → E-4 GENERATED →【人間】C1→SC
 → E-5 UP成功記録（subBatchId）
```

**ダイアログ方針（2026-07-26／E-3例外 2026-07-27）**: E-1・E-2・E-4 は **成功時に OK 確認を出さない**（toast ＋ Logger）。**E-3 のみ例外**: 完了時に **承認Web URL付きダイアログ**を出し、Web承認①へ誘導（18-②相当）。**エラーで止まったとき**は各ステップで原因の alert。E-0（前提一覧）と E-5（subBatchId入力）と案内メニューは従来どおり。C-Amazon／Z-21 単独実行の確認ダイアログは変更なし。

**親 MAIN URL（2026-07-27）**: E-2／U4 後、子の `Amazon MAIN URL` を親行（空欄時）へ自動コピー。Lv4 Build も親空なら子URLへフォールバック（手コピペ不要）。

---

## 1. Script Properties（区間ごと）

| 区間 | Key | 値 |
|------|-----|-----|
| E-1 / E-2 | `AMAZON_IMAGE_U2_ENABLED` | `true`（終わったら false） |
| E-2 | `AMAZON_U4_URL_EMBED_ENABLED` | `true`（終わったら false） |
| E-3 | `APPROVAL_QUEUE_V1_ENABLED` | `true` |
| E-4 | `APPROVAL_AMAZON_LV4_ENABLED` | `true`（終わったら false） |
| E-4 | `APPROVAL_AMAZON_LV4_TRACK` | `B`（M1） |

U2 候補: `AMAZON_IMAGE_CANDIDATE_FOLDER_ID`（Drive 07）  
U4: `R2_*`（T2/U4 と同じ）

---

## 2. clasp push

```text
clasp push
```

含む: `コード.js`（`createEAmazonCourseMenu`／E-0〜E-5）

シートを開き直すか、メニュー再読込で **E. Amazon出品コース（一時）** が出る。

---

## 3. 安眠など新規SKUの最短

1. 子に出品CK／白抜きを Drive 07  
2. U2・U4・キュー・Lv4 の Property を必要区間だけ true  
3. **E-0 → E-1 →（ドラッグ）→ E-2 → E-3 →（承認）→ E-4 → C1/SC → E-5**  
4. 全トグル false  

承認①済み・MAIN URL だけ足りない場合: **E-2 → E-4** で可。

---

## 4. 戻し方

- Property false  
- `git revert`（当該コミット）／メニュー E 削除  
- C-Amazon／Z-18/21 は残る  

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-29 | **本線＝D** を明記（新規／既存相乗りは D ラジオ。E はテスト用）。E-0 案内1行。 |
| 2026-07-27 | **E-3**: 完了時に承認Web URL付きダイアログ（18-②不要）。要 clasp push。 |
| 2026-07-27 | **安眠クローズ**: SC合格（2/2・エラー0）＋E-5 `A1_20260726_225610_4f0558_B2`。サマリ=Downloads。 |
| 2026-07-27 | **安眠**: E-4済→C1/SC送信済。次=E-5 `A1_20260726_225610_4f0558_B2`。[REMOTE](D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)。 |
| 2026-07-27 | **親 MAIN URL**: U4/E-2 後に子→親へ空欄コピー。Lv4 Build は子フォールバック。 |
| 2026-07-26 | E-1〜E-4 **silent**（成功は toast・エラー時のみ alert）。U4/`menuApprovalAmazonLv4Run({silent:true})` |
| 2026-07-26 | 初版。コース＋ゲートの一時メニュー E。 |
