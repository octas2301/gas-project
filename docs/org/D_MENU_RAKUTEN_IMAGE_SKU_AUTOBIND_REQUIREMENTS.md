# 楽天画像 SKU 自動紐付け — 要件

**文書種別**: 要件定義  
**最終更新**: 2026-08-09  
**状態**: **実装済（コード）**／実機検収は HUMAN_RUN  
**承認**: [LV4_RAKUTEN_IMAGE_SKU_AUTOBIND_APPROVAL.md](LV4_RAKUTEN_IMAGE_SKU_AUTOBIND_APPROVAL.md)  
**親**: [D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md](D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)／Amazon 対称 [D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md](D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md)（U2-ε）

---

## 1. ゴール一文

数量セット違いの各子SKUについて、**ファイル名に子SKUを含む画像**を楽天マッチング sheet（`★画像AIマッチング(操作用)`）の当該行へ自動投入し、既存 `executeRenameAndUploadFromMatrix` で `子コード.jpg`／`親_subN.jpg` として R-Cabinet へ載せる。セット数はマスタ列で解決する。

---

## 2. 方針（固定）

| 項目 | 方針 |
|------|------|
| 紐付けキー | **子SKU（ファイル名一致・最長優先）** |
| セット数 | マスタの総個数／セット数列で子SKUを選んでから `{子SKU}_….jpg` を書く |
| Vision で N 推定 | **本フェーズ対象外**（Amazon ε「セット個数マッチ」と同バックログ） |
| 永続の正 | **マスタ列**（アップ後 URL）。sheet は作業面 |
| アップ時リネーム | 既存経路を維持（作り直さない） |

```text
マスタ（子SKU＋セット数）
  → 生成ツールが {子SKU}_rakuten.jpg / {子SKU}_{pattern}_subN.jpg
  → Drive 楽天ソース
  → generateAiImageMatrix（ファイル名自動 → 余りは従来 AI）
  → 目視 → リネーム＆R-Cabinet
```

---

## 3. 作るもの

| # | 要件 |
|---|------|
| R1 | **MAIN自動**: ファイル名に**子SKU**を含むとき、当該子行の **楽天メイン画像1** が空なら自動投入。既存非上書き。親SKUのみ一致・余り順置きは MAIN 自動しない |
| R2 | **トグル**: `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED`（未設定=true）。`false` でファイル名自動オフ |
| R3 | **命名**: MAIN＝`{子SKU}_rakuten.jpg`（既存 compose）。子SKUはマスタ「セット数＝N」行から決定 |
| R4 | **サブ**: `{子SKU}_{pattern}_subN.jpg`（N=1..10、pattern=themeSlug）。空の **楽天サブ画像N** へ。`_subN`／`_pN` を解釈。pattern無しの `{子SKU}_subN.jpg` も互換 |
| R5 | マスタ永続・再生成契約を壊さない |
| R6 | ログ: `runId`／MAIN・サブ自動件数／既存スキップ／未使用件数 |
| R7 | HUMAN_RUN・検収チェックリスト |

---

## 4. 作らないもの

- 画像からセット数 N を OCR/Vision 推定して SKU を決めること（本チケットの本線にしない）
- `generateRakutenCSV`・バリエーション構造変更
- Yahoo 出品 API・Amazon U2 の破壊
- Drive `03\02` の破壊的整理
- アップ時リネーム規約の再設計

---

## 5. 受け入れ条件

- [ ] セット数違いの子が2つ以上ある JAN で、各 `{子SKU}_rakuten.jpg` が正しい行の楽天メイン1に入り、他行に入らない
- [ ] 既存メインがある行は上書きされない
- [ ] Property `false` でファイル名自動投入が止まる
- [ ] `{子SKU}_subN.jpg` が空のサブNへ入る
- [ ] マトリクスアップ後、マスタの楽天メイン／サブ URL が従来どおり埋まる
- [ ] 楽天 CSV・Amazon MAIN 自動（既存 ε）の回帰が壊れない

---

## 6. リスク・衝突ルール

| 状況 | 扱い |
|------|------|
| 短い SKU 断片 | **最長一致優先** |
| 親SKUのみ一致 | **候補のみ**（MAIN 自動しない） |
| 同一子に MAIN 複数 | **メイン1へ1枚**・余りは候補 |
| ファイル名パスと旧 Vision | **ファイル名を先**。使用済みは AI 対象外 |

---

## 7. バックログ

- Amazon U2-ε 残り＝**画像からセット個数マッチ** → 完成後に楽天と共通化検討
- サブの phaseOrder 連番と `_subN` の運用自動化の更なる統合（エクスポート補助は本フェーズで提供）

---

## 8. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-09 | 初版。U2-ε 転用・ファイル名＋マスタセット数。Vision N は対象外。 |
