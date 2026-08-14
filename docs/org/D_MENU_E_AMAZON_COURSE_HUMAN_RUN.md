# E. Amazon出品コース（一時）— 人間手順

**性質**: 互換ファサード。画像は **C. 画像コース** が本線（E-1/E-2＝C-1/C-2互換）。裏は既存 C-Amazon／18／21／U4。  
**本線**: Amazon 出品の起点は **D**。画像は **Cコース**。[C承認](LV4_C_COURSE_CONSOLIDATION_APPROVAL.md)／[C HUMAN_RUN](D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)  
**分割逃げ道**: Z→C-Amazon①〜④／Z→18・21（削除しない）

---

## 0. 実行順（人が押すのはこれだけ）

```text
【画像】C-0 → C-1 Amazon →【人間】MAINドラッグ → C-2
【出品】E-0 → E-3 →【人間】Web承認① → E-4 → C1/SC → E-5
（E-1/E-2 は旧名の互換。使わず C-1/C-2 でよい）
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
3. **画像は C-1 Amazon →（ドラッグ）→ C-2**。その後 **E-3 →（承認）→ E-4 → C1/SC → E-5**  
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
| 2026-08-01 | **Cコース本線化**: 画像は C-1/C-2。E-1/E-2は互換。 |
| 2026-07-27 | **E-3**: 完了時に承認Web URL付きダイアログ（18-②不要）。要 clasp push。 |
| 2026-07-27 | **安眠クローズ**: SC合格（2/2・エラー0）＋E-5 `A1_20260726_225610_4f0558_B2`。サマリ=Downloads。 |
| 2026-07-27 | **安眠**: E-4済→C1/SC送信済。次=E-5 `A1_20260726_225610_4f0558_B2`。[REMOTE](D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)。 |
| 2026-07-27 | **親 MAIN URL**: U4/E-2 後に子→親へ空欄コピー。Lv4 Build は子フォールバック。 |
| 2026-07-26 | E-1〜E-4 **silent**（成功は toast・エラー時のみ alert）。U4/`menuApprovalAmazonLv4Run({silent:true})` |
| 2026-07-26 | 初版。コース＋ゲートの一時メニュー E。 |
