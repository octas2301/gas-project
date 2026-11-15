# C／D／E メニュー分割 — 要件（ロック）

**文書種別**: requirements  
**状態**: 2026-09-05 チャットロック・実装  
**関連**: [B45_MAIN_SUB_ALIGN_REQUIREMENTS.md](B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)／[D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md](D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)／[D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md](D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)

---

## 1. トップ入口

| 記号 | 名称 | 役割 |
|------|------|------|
| **C** | 楽天　画像マッチング | 02仕分けまで。アップしない |
| **D** | 楽天・Yahoo 出品アップロード | R-Cabinet＋楽天／Yahoo出品。**Amazonなし** |
| **E** | Amazon 画像マッチング＋出品アップロード | 07マッチング→R2→Amazon出品 |
| **Z** | 分割・互換 | 旧D補助段階・承認①温存・C-Amazon①〜④ |

**廃止（トップ）**: `D補助. Amazon段階実行（旧E）`／C内の Ama①②／Dの「Amazonのみ」「フル→Amazon」

---

## 2. 項目

### C

- `02楽天画像フォルダ→マッチングシート作成`（本線）
- `金丸数字やり直し（一括）※常用はB-④`（まれ）

### D

- フル（楽天→Yahoo）／楽天のみ／Yahooのみ（従来どおり）
- 画像スキップ／ベストエフォート（スキップは出品CKの楽天メイン画像1そろい時に初期ON。再送は人間がレ点を外す）

### E（下層なし・3項目）

1. `E① 07→マッチング／MAINをsheetで並び替え`（前段生成は **B-⑥**。楽天未整理フォルダは読まない）
2. `E② 並び替え結果をマスタsheetへ反映→R2へUP保存`（目視後。前提: Dで楽天サブURL。U4は未充足全件・22分で切って機械再開）
3. `E③ Amazon出品ファイル作成（新規／相乗り）`（旧DのAmazonコース）

---

## 3. 日常順

```text
C → D（画像オン）→ E① → 目視 → E② → E出品
```

楽天・Yahooのみ: `C → D`

---

## 4. 実装メモ

- E①（`menuCImageCourseC1Amazon`）は楽天未整理フォルダを読まない（`AMAZON_E1_KEEP_RAKUTEN_ENABLED` 未設定=true）
  - **レ点一致**: シートを消さない。楽天欄はそのまま。Amazon 枠＋07候補だけ
  - **レ点不一致**: 枠を今のレ点で作り直し。親子SKUが重なる行だけ楽天欄を引き継ぐ。外れた行は捨てる。新規行の楽天欄は空（C が必要）
- E②/U4: 未充足子SKUを全件キュー。期限は E②開始＋22分。残りは `runAmazonU4ResumeFromTrigger`（人が同じメニューを件数分割で押さない）
  - **シートなし**: 空枠＋07のみ
  - 切戻し: Property `AMAZON_E1_KEEP_RAKUTEN_ENABLED=false`（旧 `generateAiImageMatrix`）
- `showCImageCourseModal`／`showBatchExportModal`／`showEAmazonExportModal`
- Amazon facade: `runBatchExportAmazonFacade`（入口のみ E へ）
- Z: `createEAmazonCourseMenu`＝「E段階・互換（旧D補助）」

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-10-19 | トップのC／D／Eが古い画像コースに戻っていたため、9月18日の入口へ再掲。要 clasp |
| 2026-09-16 | E②/U4: 未充足全件をキュー。切れはE②開始＋22分→トリガー再開。人が同じメニューを件数分割で押さない。要 clasp。 |
| 2026-09-16 | E①: 楽天欄保持＋レ点差分は重なりSKUだけ引継ぎ。未整理フォルダ非読。切戻し `AMAZON_E1_KEEP_RAKUTEN_ENABLED=false`。要 clasp。 |
| 2026-09-16 | E下層に **E①／E②／E③** 接頭辞（文言本体は変更なし）。要 clasp。 |
| 2026-09-05 | 初版ロック・実装 |
