# Lv3 Yahoo 日中掲載（既存経路オーケストレーション）— 要件定義

**文書種別**: 要件定義＋実装済み＋**人間検収済み（2026-07-20）**  
**最終更新**: 2026-07-20  
**親**: [LEVELLED_IMPLEMENTATION_PLAN.md](LEVELLED_IMPLEMENTATION_PLAN.md) ・ [AI_APPROVAL_MATRIX.md](AI_APPROVAL_MATRIX.md) ・ [LV1_APPROVAL_QUEUE_REQUIREMENTS.md](LV1_APPROVAL_QUEUE_REQUIREMENTS.md) ・ [LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md) ・ [PHASE0_THREE_REVIEW_MAJORITY.md](PHASE0_THREE_REVIEW_MAJORITY.md) §4.2・§4.3  
**コード**: `YahooApprovalExport.js`（本体）／`ApprovalQueue.js` に Yahoo読取ヘルパ／`コード.js` メニュー20のみ／**`Yahoo.js` の出品API本体は非改変**（`runYahooExport` 呼出のみ）  
**ゴール一文**: 承認①済みの Yahoo 行（**子SKU**）だけを、既存 `runYahooExport`（画像API・editItem・setStock）経路で **呼び出して** 在庫0/1掲載する。**Yahoo出品ロジック本体は改変しない。**  
**手動逃げ道**: 既存「🛒 Yahoo!出品 → ▶ 出品実行」および一括出品内の Yahoo は **非改変・継続利用可**。

---

## 1. スコープ

### 1.1 作るもの

| # | 成果物 |
|---|--------|
| 1 | 承認キュー（`▼承認キュー(出品①)`）から **mall=`yahoo` かつ lineStatus=`APPROVED`** を読む入口 |
| 2 | 実行サブバッチ分割（**主: 実働約25分**／**副: ユニーク画像 ≤ 50**＝楽天運用揃え。§4） |
| 3 | マスタ上の出品対象合わせ（承認済み **子SKU行のみ** レ点ON。親だけレ点で全子出品は禁止） |
| 4 | 在庫0（既定）／在庫1（バッチ `inventoryMode=ONE` のみ）の掲載前マスタ在庫の扱い（書込範囲を明示） |
| 5 | `runId` / `batchId` / `subBatchId` / SKU / `state` の調査ログ＋失敗メール |
| 6 | まず **手動キック**（メニュー20）。日中自動トリガーは後続で可 |

### 1.2 作らないもの（禁止）

- `Yahoo.js` 内の `YahooApiClient` / `YahooImageUploader` / `YahooDataBuilder` の仕様変更（パラメータ名・seller_id 位置・文字数制約等）  
- `generateRakutenCSV` および楽天 Lv2 経路の改変  
- Amazon の実行  
- 承認②（補充）・販売中SKUへの無人上書き（U1）  
- B統合 `B_INTEGRATED_STEP_FUNCTIONS` の順序・境界変更  
- **親だけレ点で全子出品**  
- `clasp push` 自動化  

### 1.3 聖域の守り方

```text
[Lv3 ラッパー]
  → 承認済み Yahoo 子SKUに対象を絞る（一時レ点）
  → 在庫0/1をマスタの在庫列へ必要最小だけ合わせる（要実装設計・承認）
  → runYahooExport(ssOverride) を呼ぶ
  → 出品CKをスナップショットへ復元
```

既存の手動 Yahoo 出品・一括出品は **触らない**。Lv3は承認キュー起点の別エントリとする。

---

## 2. 前提（着手条件）

| # | 条件 |
|---|------|
| 1 | **Lv1 人間検収完了**（2026-07-17） |
| 2 | **Lv2 人間検収完了**（2026-07-20。楽天オーケストレーションの型が確定） |
| 3 | 手動「▶ 出品実行」（`runYahooExport`）が現状どおり動くこと |
| 4 | 本要件の社長確認（§9）→ **実装承認済（2026-07-20）**。人間検収は §10 |

---

## 3. 入力・対象SKU

### 3.1 承認キュー

- シート: `▼承認キュー(出品①)`（Lv1）  
- 対象: ヘッダ `status=APPROVED` かつ明細 `mall=yahoo` かつ `lineStatus=APPROVED`  
- `REJECTED` / `CANCELLED` / `ORPHANED` は実行しない  
- 実行直前再チェック: マスタに該当子SKUが無い → ORPHAN扱い・スキップ（方式A）  
- 販売中かつ在庫>0 → **原則スキップ**（Lv1の `MAY_SKIP_IN_STOCK` プレビューと整合。補充は②）

### 3.2 Yahoo の行単位（子SKUのみ）

- Lv1抽出どおり **子行の出品CK**（親だけレ点は候補に含めない）。  
- `runYahooExport` はマスタの **出品CK付き行**から `YahooDataBuilder` が商品を組むため、ラッパーは次を採用:

| 案 | 内容 | 備考 |
|----|------|------|
| **A（採用）** | 実行直前にスナップショット → 承認済み **子SKU行のみ** レ点ON／他はOFF → `runYahooExport` → 復元 | Lv2案Aと同型。親行は原則OFF（子のみが対象） |
| B | 人間レ点 ⊆ 承認済みのみ実行 | 二重管理。Lv3では不採用 |

**復元失敗時はメール必須・処理中断。**

### 3.3 在庫0/1

- バッチヘッダ `inventoryMode`: `ZERO`（既定）/ `ONE`  
- 掲載前に、対象子（および `runYahooExport` / setStock が参照する在庫列）を **0または1** に合わせる処理が必要なら、**対象SKUの在庫列のみ**・実行ログ必須・PropertyトグルでOFF可  
- 販売可能数への引き上げは **しない**（承認②）

### 3.4 新規 vs 既存更新

- 当面 **既存更新中心**で閉じる（多数決メモ・マトリクス U6）。  
- 新規出品（`submitItem` 不可・`it-07004` 手動反映）の日中バッチ分離は **後送り**。実装時は既存更新で失敗／新規扱いになった行をログ＋スキップ／メールし、自動リトライで新規を押し切らない。

---

## 4. 実行分割（Yahoo適用）

詳細の正（共通）: [AI_APPROVAL_MATRIX.md](AI_APPROVAL_MATRIX.md) §5 ・多数決メモ §4.3。  
**楽天の「ユニーク50」は楽天画像アップ上限向け。Yahoo公式の1バッチ枚数上限ではない。**

### 4.1 公式制約（参考・分割の主因にしない）

| 公式 | 内容 |
|------|------|
| 画像アップロード | **10,000枚／1時間**（全経路）。超過時 429／503 等。([uploadItemImage](https://developer.yahoo.co.jp/webapi/shopping/uploadItemImage.html)) |
| 1ファイル | 例: 2MB以下、gif/jpg 等 |
| 1商品の画像ひも付け | 最大21枚程度（アップロードCSV仕様） |

通常の日中件数では 10,000／時にはまず当たらない。Lv3の分割は **GAS時間**と **楽天との運用揃え**を優先する。

### 4.2 Lv3の分割規則（採用）

| 優先 | 規則 | 内容 |
|------|------|------|
| **主** | 時間 | 1実行あたり実働 **約25分**で中断→Script Properties に再開位置→トリガー再開（課金30分前提） |
| **副** | 画像（運用揃え） | サブバッチ内の **ユニーク画像数 ≤ 50**。共有画像は1枚。識別子は §6（Lv2と同じ） |
| — | 品番 | 固定N品番は必須としない。主・副に収まるだけ子SKUを積む |
| — | 12:00 | 未完了は翌朝続き（明示取消まで有効） |
| — | 冪等 | 同一 `batchId`+子SKU で成功済みは再出品しない。失敗側のみ最大2回 |

**なぜ副にユニーク50を残すか（社長方針 2026-07-20）**: Yahoo公式のバッチ上限ではないが、**楽天と出品分量を揃えると運用が安定する**ため。将来ゆるめる場合は Property で閾値変更（§7）。

`subBatchId` 例: `{batchId}_Y{n}`（n=1,2,…）

---

## 5. 状態・ログ

### 5.1 実行状態

例: `PENDING_RUN` / `RUNNING` / `DONE` / `FAILED` / `SKIPPED_IN_STOCK` / `SKIPPED_ORPHAN` / `SKIPPED_NEW_ITEM` / `RETRYING`

### 5.2 Logger 必須

- `runId` / `batchId` / `subBatchId` / `functionName` / `state`  
- 対象子SKU件数・ユニーク画像概算・スキップ理由  
- **シークレット・トークン全文は出さない**

### 5.3 メール

- サブバッチ失敗・レ点復元失敗・25分切断後の再開失敗 → **メール必須**（件名に「Lv3 Yahoo」等で楽天と区別）

---

## 6. ユニーク画像の識別子

Lv2と同一（運用揃え）:

| 候補 | 採用 |
|------|------|
| **Drive ファイルID** | 優先 |
| 正規化URL（クエリ除去） | URLのみのときのフォールバック |

---

## 7. エントリポイント（実装）

| 種別 | 内容 |
|------|------|
| メニュー | Z → **20. 承認①済→Yahoo日中掲載(Lv3)** → 20-①実行／20-②状態クリア |
| Property | `APPROVAL_YAHOO_LV3_ENABLED`（既定 `false`） |
| 任意 | `APPROVAL_YAHOO_LV3_APPLY_STOCK`（既定 true）／`APPROVAL_YAHOO_LV3_SKIP_EXPORT`（ドライラン）／`APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_LIMIT`（既定 `50`。運用揃えの副制約。公式上限ではない） |
| トリガー | 実働25分超過時に自動再開（`runApprovalYahooLv3FromTrigger`） |

既存「▶ 出品実行」は残す（聖域・手動逃げ道）。

---

## 8. 検収条件（Lv3完了）

1. Property OFF ではメニューが動かない（または即return）  
2. テスト少数子SKUで、承認①→Lv3実行→`runYahooExport` 経路まで到達  
3. `Yahoo.js` の出品API／Builder 本体の意図しない改変が無い（ラッパー新規＋呼出＋メニュー最小）  
4. 25分レジューム・ユニーク50（副）がログで追える（件数少なら分割ロジックの単体検証可）  
5. 販売中在庫>0・ORPHAN・REJECTED・（方針どおり）新規押し切り が実行されない  
6. レ点スナップショット案Aで、実行後にレ点が意図どおり復元される（実行中は承認済み子のみON）  
7. 必須3点セット（本docs・調査ログ・Propertyトグル／revert）済み  

### 8.1 人間検収記録（2026-07-20）

| 項目 | 結果 |
|------|------|
| 実施内容 | Property ON → ドライラン（`SKIP_EXPORT=true`）→ 本番20-①。在庫列同名二重の切り分け後に再実行 |
| 例 runId（ドライラン） | `LV3_20260720_111629_027775`（`candidates=7`・`SKIP_EXPORT`・サブバッチ1/1） |
| 例 runId（本番） | `LV3_20260720_112238_777998`（`childrenDone=7`・`skipped=0`） |
| batchId | `A1_20260720_083227_1f0b30` |
| テスト子SKU | `lifec-4560151300139-*`（7子） |
| 在庫列 | マスタに「在庫数」が2列あったため先勝ちで左列（旧AM）を誤読。左列ヘッダーを **「在庫数計算」** に変更し、出品用は **HX「在庫数」**（`col在庫数=231`）に一本化 |
| Yahoo反映 | 本番で未反映（在庫0）がストアクリエイターProに積まれた。公開反映はせず、在庫復元後に手動「▶ 出品実行」で上書き（20-①は販売中スキップのため復元用途に使わない） |
| 聖域 | `Yahoo.js` 出品本体は非改変（呼出のみ） |
| Property | 検収後 `APPROVAL_YAHOO_LV3_ENABLED=false` に戻すこと |
| 判定 | **Lv3 完了**。次は Lv4 Amazon（または次優先） |

### 8.2 運用メモ（検収で確定）

- スキップ／出品判定の「在庫数」は **出品用1列のみ**（計算用列は別名にする）。  
- 在庫0テスト後の本番在庫復元は **手動 Yahoo 出品**。Lv3（在庫0原則＋販売中スキップ）では戻せない。  
- ストアクリエイターProの未反映は一覧から消せず、**反映しない**か **正しい内容で上書きしてから反映**。

---

## 9. 実装時の承認パッケージ

**2026-07-20 社長明示承認「要件OK・実装して」により実装実施。**

- 変更ファイル: `YahooApprovalExport.js`（新規）、`ApprovalQueue.js`（`approvalQueueGetLatestApprovedYahoo_`）、`コード.js`（メニュー20）、本docs  
- 概要: 承認済み Yahoo 子のみ・`runYahooExport` 呼出・分割（主25分／副ユニーク50）・在庫0/1  
- リスク: レ点一時変更の復元漏れ／在庫列の誤更新／親レ点誤用／API書込の誤対象  

**復元**: Property `APPROVAL_YAHOO_LV3_ENABLED=false`／`git revert`／新規js削除＋メニュー20削除。

---

## 10. 人間向け検証手順（実装後・自宅）

1. Lv1・Lv2検収済みであること  
2. `git pull` → `clasp push`（**YahooApprovalExport.js が含まれること**）  
3. Property: `APPROVAL_YAHOO_LV3_ENABLED=true`（初回は任意で `APPROVAL_YAHOO_LV3_SKIP_EXPORT=true`）  
4. 承認①に Yahoo 子（レ点由来）があること。販売中在庫>0はスキップ想定  
5. 20-① → ログに `childrenOn=`・`uniqueImages=`／復元ログ  
6. 実行後、出品CKが実行前に戻っていること  
7. （任意）YahooストアクリエイターProで対象SKUの在庫0/1・更新有無を目視  
8. 終わったら Property を false  

---

## 11. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-20 | **人間検収完了**（§8.1）。runId例 `LV3_20260720_112238_777998`。在庫数列の二重ヘッダー運用メモ（§8.2）。 |
| 2026-07-20 | スキップ切り分けログ追加（sheet名・stockRaw・ORPHAN/IN_STOCK）。 |
| 2026-07-20 | **実装**: `YahooApprovalExport.js`・メニュー20・Yahoo読取ヘルパ。`runYahooExport` 呼出のみ。人間検収待ち。 |
| 2026-07-20 | 初版ドラフト。`runYahooExport` 呼出のみ。分割は主25分・副ユニーク50（楽天運用揃え。Yahoo公式は10,000枚/時）。コード未実装。 |
