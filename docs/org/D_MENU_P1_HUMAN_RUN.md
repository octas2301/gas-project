# Amazon P1 — 07以降 File最少（人間手順）

**状態**: **検収OK**（2026-08-01。Cモーダル実機＋他項目認識確認）  
**承認**: [LV4_P1_FILE_MIN_APPROVAL.md](LV4_P1_FILE_MIN_APPROVAL.md)  
**親**: [D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md](D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)／[D_MENU_U2_HUMAN_RUN.md](D_MENU_U2_HUMAN_RUN.md)／[D_MENU_U4_HUMAN_RUN.md](D_MENU_U4_HUMAN_RUN.md)

---

## 0. 日常フロー（P1）

```text
（常設）AMAZON_IMAGE_CANDIDATE_FOLDER_ID = Drive07
 → C「Amazon新規：準備」
 →【人間・File】白抜きを 07 へ置く
 →【人間・sheet】MAINドラッグ（ONLYならPT）
 → C「Amazon新規：ドラッグ後」（③→④→U4。Drive File はGASのみ）
 → D で出品
 →【人間・ZIP】SC Upload Images（手ZIP。U5は別）
```

**やってはいけない（日常）**

- Drive `02` へ手で MAIN/PT を置く・リネームする  
- `07/アップロード済み` へ手で移す  
- U4／D の前に `02` を開いて確認する  

失敗時の第一案内は **C-1→ドラッグ→C-2 再実行**（ダイアログ／ログに明記）。

**復旧例外**: Z → C-Amazon①〜④。`02` 手置きは例外・復旧のみ。

---

## 1. Script Properties

| Key | 扱い |
|-----|------|
| `AMAZON_IMAGE_CANDIDATE_FOLDER_ID` | 常設必須（07） |
| `AMAZON_DRIVE_IMAGE_FOLDER_ID` | 空なら `02` 既定 |
| `AMAZON_IMAGE_CANDIDATE_ARCHIVE_ENABLED` | 未設定＝有効（成功分のみ退避） |
| U2/U4 | C本線は一時ON→復元 |

新規 Property なし。

---

## 2. clasp push

```text
clasp push
```

含む: `コード.js`／`AmazonImageMatrixExport.js`／`AmazonDriveImageExport.js`

---

## 3. 検収

- [x] 日常手順に `02` 手操作が無い  
- [x] C-0／Cモーダル／C-2完了 toast に File最少・02触らない旨がある（**Cモーダル実機**: 「02は触らない」＋P1注記）  
- [x] ④失敗／U4 MAINなし／Dゲート失敗の案内が「C-2再実行」で、Drive手直しを第一にしない（認識）  
- [x] ZIP自動・MAINドラッグ自動化が混入していない（認識）  
- [x] 楽天C非退行（文言のみ変更・認識）  

---

## 4. 戻し方

- **Git**: 当該差分 revert  
- 挙動トグルなし（案内文言のみ）  

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 下書き。 |
| 2026-08-01 | 実装（案内・ゲート文言）。HUMAN_RUN更新。 |
| 2026-08-01 | **検収OK**（Cモーダル実機＋他認識）。 |
