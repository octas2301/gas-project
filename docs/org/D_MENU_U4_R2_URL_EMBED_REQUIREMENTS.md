# U4 — R2 URL → GENERATED／PACKAGED 埋め要件

**文書種別**: 要件定義（**U4 v1 実機合格**）  
**最終更新**: 2026-07-26  
**状態**: **実機合格**（[HUMAN_RUN §0](D_MENU_U4_HUMAN_RUN.md)）  
**親**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md) §6.2 Db・§9 U4  
**設計参照**: [LV4_R2_IMAGE_PIPELINE_POC.md](LV4_R2_IMAGE_PIPELINE_POC.md) T4／[LV4_T2_HUMAN_RUN.md](LV4_T2_HUMAN_RUN.md)  
**依存クローズ**: U2 実機合格・T2 PoC・**T2再検証合格（2026-07-26・URL単独・18320なし）**

---

## 1. ゴール一文

Drive `02` の `{sellerSku}.MAIN.jpg`（必要なら PT）を **R2 に上げ、公開 https URL を GENERATED（およびマスタ参照）へ自動で書き**、人間が xlsm に URL を手貼りする手間を減らす。  
**純正 `.xlsm` の VBA／行マップ直編集は U4 ではやらない**（それは C1・別ゲート）。

---

## 2. 背景（なぜ今）

| 事実 | 含意 |
|------|------|
| T2（21-⑥）で Drive→R2 Put＋URL 取得は済 | 部品はある |
| 2026-07-26 再検証: `80s10` で **URL単独・ZIPなし・18320なし・店頭OK** | URL本線が成立し得る → **T3急がない** |
| 現状ボトルネック | 人間が R2 URL を PACKAGED xlsm の「メイン画像のURL」へ手貼り |
| GENERATED | 現状は主に **楽天メイン画像** URL を参照（`AmazonApprovalExport`） |

---

## 3. スコープ

### 3.1 作るもの（U4 v1）

1. **対象SKUの MAIN（＋ONLY時 PT）** を Drive `02` から R2 へ Put（T2ロジック再利用・複数SKU可だが **トグル＋件数上限**）  
2. 公開 URL を次へ書く（優先順）:  
   - **A. マスタ列**（新規 or 既存拡張）例: `Amazon MAIN URL`／`Amazon PT URL`（実装承認で正式名固定）  
   - **B. 直後の GENERATED 再生成時**に、Amazon URL があれば **楽天メインより優先**して `mainImageUrl` 等へ載せる  
3. メニュー（案）: **Z → 21-⑦**「Drive02→R2→マスタURL」（名称は実装承認で固定）  
4. 調査ログ: `runId`／sku／http／url／書込行  
5. HUMAN_RUN 1枚  

### 3.2 作らないもの（禁止・別ゲート）

| 対象 | 理由 |
|------|------|
| 純正 `.xlsm` バイナリ編集（C1） | テンプレ陳腐化リスク。POC T5＝当面スキップ |
| T3 ZIP 自動（U5） | T2再検証後も **急がない**。別承認 |
| ε 自動 MAIN 紐付け | バックログ |
| 楽天 `03`／`generateRakutenCSV`／Yahoo.js | 聖域 |
| 全カタログ無制限ループ | 件数上限＋トグル |
| SC API（U7） | 将来必須・別 |

### 3.3 人間に残る作業（U4後も）

- C で MAIN 紐付け → ④で `02` 出力（既存 U2）  
- PACKAGED（C0: GENERATED／マスタURLを見て xlsm へ流し込み。手貼りは減る）  
- SC スプシ UP（ZIPは **任意**。URLで足りるケースは ZIP省略可）  
- 21-③  

---

## 4. 入出力契約

| 項目 | 契約 |
|------|------|
| 入力 | Drive `02` の `{sellerSku}.MAIN.jpg`（U2④成功分）。PTは `AMAZON_ONLY` 時のみ |
| R2 キー | ファイル名と同じ（T2どおり） |
| URL | `R2_PUBLIC_BASE` + `/` + key |
| マスタ書込 | **URL文字列のみ**（在庫・JAN・価格は書かない） |
| GENERATED | 次回 21-①／Da で Amazon URL 優先。既存 GENERATED ファイルの事後パッチは **任意・v1.1** |
| 失敗 | 当該SKUスキップ＋ログ。他SKU継続可（ベストエフォートは Property） |

---

## 5. トグル・安全

| Property | 既定 | 意味 |
|----------|------|------|
| `AMAZON_U4_URL_EMBED_ENABLED` | **false** | メニュー実行ガード |
| `AMAZON_DRIVE_R2_UPLOAD_ENABLED` | false | R2 Put 時に true（既存）。成功後 false 推奨／将来は成功時自動OFF可 |
| `AMAZON_U4_MAX_SKUS` | 例 20 | 1実行の上限 |
| `AMAZON_DRIVE_IMAGE_FOLDER_ID` | 02 | 既存 |

シークレットは既存 `R2_*` のみ。コード・コミット禁止。

---

## 6. 検収（HUMAN_RUN）

- [ ] トグル false で何もしない  
- [ ] 1子SKU: `02` MAIN → R2 200 → マスタに URL  
- [ ] 続く GENERATED（または明示メニュー）で **その URL が画像列に載る**  
- [ ] 楽天CSV／Yahoo 経路非改変  
- [ ] （任意）xlsm へ流し込み→SC スプシのみで **18320なし**（ZIPなし再確認・別SKUでも可）  

---

## 7. 実装チケット

| ID | 内容 | 状態 |
|----|------|------|
| **U4-0** | 本要件＋承認パッケージ | **クローズ** |
| **U4-1** | R2複数SKU Put（T2再利用）＋マスタURL列 | **v1 実装**（21-⑦） |
| **U4-2** | GENERATED ビルドで Amazon URL 優先 | **v1 実装** |
| **U4-HR** | HUMAN_RUN | **実機合格** → [D_MENU_U4_HUMAN_RUN.md](D_MENU_U4_HUMAN_RUN.md) §0 |
| **U4-ε** | 既存 GENERATED CSV の事後パッチ | バックログ |

---

## 8. 他ゲートとの関係

| ゲート | U4後の位置づけ |
|--------|----------------|
| **T3／U5** | 必須ではない。URL失敗時のバックアップ |
| **C1 xlsm自動** | 別。U4は URL 供給まで |
| **U7 Dc** | 将来。U4は過渡の手間削減 |
| **ε** | 紐付け自動化。U4とは独立 |

---

## 9. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-26 | **実機合格**記録（21-⑦・マスタ URL・D冪等想定内）。 |
| 2026-07-26 | **U4 v1 実装**（21-⑦・GENERATED優先・HUMAN_RUN）。承認「U4 v1 承認」。 |
| 2026-07-26 | 初版起草。T2再検証合格を前提に URL本線・xlsm直編集除外。 |
