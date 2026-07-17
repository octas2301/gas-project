# Lv1 承認キュー骨格 — 要件定義

**文書種別**: 要件定義（実装は別承認後）  
**最終更新**: 2026-07-17  
**親**: [LEVELLED_IMPLEMENTATION_PLAN.md](LEVELLED_IMPLEMENTATION_PLAN.md) ・ [AI_APPROVAL_MATRIX.md](AI_APPROVAL_MATRIX.md) ・ [PHASE0_THREE_REVIEW_MAJORITY.md](PHASE0_THREE_REVIEW_MAJORITY.md)  
**ゴール一文**: 朝モバイルで「今日出してよいSKU」を承認／取消でき、結果がシートに残る。**モールへの出品・在庫API・楽天CSVは一切呼ばない。**

---

## 1. スコープ

### 1.1 作るもの

| # | 成果物 |
|---|--------|
| 1 | 候補抽出（マスタのレ点 → 承認候補行） |
| 2 | 承認①バッチの永続化（専用シート） |
| 3 | モバイル向け承認UI（既定: GAS Web・Googleログイン） |
| 4 | 承認／一部否認／バッチ取消 |
| 5 | 実行はしないが、Lv2以降が読める **承認済みスナップショット** |

### 1.2 作らないもの（Lv1禁止）

- 楽天 `generateRakutenCSV`／FTP／反映確認の起動  
- Yahoo `editItem` / `setStock` / 画像API／`submitItem`  
- Amazon バルク  
- マスタの価格・タイトル・ジャンル等の一括更新  
- 承認②（補充）  
- 販売中SKUへの無人上書き（当面手動・U1）  
- `clasp push` 自動化  
- B統合 Step 境界の変更  

### 1.3 役割の整理（レ点／承認／スキップ）

| 概念 | Lv1での扱い |
|------|-------------|
| **レ点（出品CK）** | マスタ上の **候補選択**。抽出の入力 |
| **承認①** | レ点由来リストへの **朝のハンコ**（本Lvの本体） |
| **スキップ（在庫>0）** | **出品実行時（Lv2以降）**のブレーキ。Lv1では「プレビュー警告」まで可（書込・APIはしない） |

楽天: 親レ点運用あり（従来）。Yahoo: **子SKUレ点のみ**が候補（親だけレ点で全子は禁止）。Lv1の抽出はモール別に行を分けて保持する（§4）。

---

## 2. シート設計

### 2.1 新規シート名（正）

| シート名 | 用途 |
|----------|------|
| `▼承認キュー(出品①)` | バッチヘッダ＋明細（またはヘッダ／明細を2シートに分けても可） |

推奨は **1シート＋種別列**（実装単純）。行数が増えたら後で分割。

代替（分割案）:

- `▼承認バッチ(出品①)` … 1バッチ1行  
- `▼承認明細(出品①)` … 1SKU（または1対象行）1行  

### 2.2 バッチヘッダ列

| 列キー | 必須 | 内容 |
|--------|------|------|
| batchId | ○ | §3 の規則 |
| createdAt | ○ | 作成日時（JST） |
| createdBy | ○ | 実行者メール or `menu` / `trigger` |
| status | ○ | `DRAFT` / `PENDING_APPROVAL` / `APPROVED` / `CANCELLED` / `EXPIRED`（Lv1は EXPIRED 未使用可） |
| approvedAt | | 承認日時 |
| approvedBy | | 承認者（当面社長アカウント） |
| cancelledAt | | 取消日時 |
| inventoryMode | ○ | `ZERO`（既定）/ `ONE`（在庫1フォールバック。画面で明示選択した場合のみ） |
| note | | 人間用メモ |
| sourceSummary | | 例: `rakutenParents=3;yahooChildren=5` |
| schemaVersion | ○ | 固定 `lv1-1` |

### 2.3 明細列

| 列キー | 必須 | 内容 |
|--------|------|------|
| batchId | ○ | ヘッダと同一 |
| lineId | ○ | バッチ内一意（連番または uuid 短） |
| mall | ○ | `rakuten` / `yahoo` / `amazon`（Lv1は amazon 行を空でも可。出さなくてよい） |
| masterRow | ○ | マスタの行番号（1始まり） |
| parentSku | ○ | |
| childSku | | Yahoo行は必須。楽天親行は空可 |
| productName | | 表示用（マスタからコピー） |
| checkboxCol | ○ | 判定に使った列名（例: 出品CK） |
| lineStatus | ○ | `CANDIDATE` / `APPROVED` / `REJECTED` |
| previewFlag | | 実行時スキップ見込みのヒント。`MAY_SKIP_IN_STOCK` 等（任意） |
| rejectReason | | 否認時 |

**マスタ本体への書き戻しはしない**（レ点を外す／付けるのは人間。Lv1はキュー側のみ更新）。

---

## 3. batchId 規則

```text
A1_{yyyyMMdd}_{HHmmss}_{6桁乱数}
例: A1_20260717_090415_a3f91c
```

- タイムゾーン: `Asia/Tokyo`  
- 同一秒の衝突は乱数で回避  
- ログ・UI・明細はすべて同一 `batchId` で紐づける  

---

## 4. 候補抽出規則

### 4.1 共通

1. 対象シート: `▼商品マスタ(人間作業用)`（名称は既存正に合わせる）  
2. チェックボックス判定: boolean `true` と文字列 `"TRUE"` の両対応（既存 Yahoo ヘルパー相当）  
3. **全件を無条件承認しない**。抽出 → `PENDING_APPROVAL` → 人が承認  

### 4.2 楽天候補（mall=`rakuten`）

- **親行**（子SKU空かつ出品CK）にレ点がある行を候補とする（既存楽天運用に合わせる）  
- 詳細な親／子判定は既存マスタ前提（[HANDOVER.md](../../HANDOVER.md) §5.2.1）に従う  
- Lv1では「行番号＋親SKU＋商品名」が分かれば足りる  

### 4.3 Yahoo候補（mall=`yahoo`）

- **子SKU行の出品CKのみ**  
- 親だけレ点の行は **Yahoo候補に入れない**（禁止仕様の踏襲）  

### 4.4 Amazon

- Lv1では抽出 **任意**（未実装で空でも Lv1 完了可）  
- 列仕様が固まるまでスキップしてよい  

### 4.5 プレビュー警告（任意・推奨）

抽出時にマスタの「在庫数」相当が **>0** と読める行には `previewFlag=MAY_SKIP_IN_STOCK` を付け、UIで「実行時スキップ見込み」と表示する。  
**Lv1では在庫を変更しない・APIしない。**

---

## 5. 承認UI（GAS Web）

### 5.1 認証・権限

| 項目 | 要件 |
|------|------|
| ログイン | Google アカウント（GAS Web 既定） |
| 許可アカウント | Script Properties の許可リスト（当面社長メール1件以上） |
| 未許可 | 表示のみ拒否／操作不可 |
| LINE | Lv1では **実装しない**（代替は後続） |

Property 例（実装時）:

- `APPROVAL_UI_ALLOWED_EMAILS` … カンマ区切り  
- `APPROVAL_QUEUE_SHEET_NAME` … 既定 `▼承認キュー(出品①)`  

### 5.2 画面操作（最小）

1. 最新の `PENDING_APPROVAL` バッチを表示（なければ「候補を作成」）  
2. **候補を作成**: マスタから抽出して新 batchId で保存。status=`PENDING_APPROVAL`  
3. 明細一覧（mall／SKU／名前／警告フラグ）  
4. **一括承認**: 全 `CANDIDATE` → `APPROVED`、ヘッダ `APPROVED`  
5. **行ごと否認**: `REJECTED`（理由任意）  
6. **バッチ取消**: ヘッダ `CANCELLED`（明示取消まで有効の反転操作）  
7. inventoryMode: 既定 `ZERO`。`ONE` にする場合は確認ダイアログ必須  

### 5.3 メニュー（スプレッドシート側・任意）

- Z 配下または「承認」メニューに  
  - 「承認候補を作成（出品①・書込なし）」  
  - 「承認WebのURLを表示」  
- どちらも **EC API を呼ばない**  

---

## 6. 状態遷移

```text
(なし) --作成--> PENDING_APPROVAL
PENDING_APPROVAL --承認--> APPROVED
PENDING_APPROVAL --取消--> CANCELLED
APPROVED --取消--> CANCELLED
REJECTED 行は APPROVED バッチ内に混在可（実行時は APPROVED 行のみ）
```

- 有効期限方針 **(C) 明示取消まで有効**（マトリクスどおり）  
- Lv1では日次自動 EXPIRE は **必須ではない**（後で追加可）  

---

## 7. ログ

`Logger.log` に最低限:

- `runId` / `batchId` / `functionName` / `state`（PENDING/RUNNING/DONE/FAILED）  
- 抽出件数（mall別）  
- 承認者メール（シークレット以外）  
- **ESA・APIキーは出さない**  

---

## 8. 検収条件（Lv1完了）

すべて満たすこと:

1. 許可アカウント以外は承認操作ができない  
2. レ点から候補が作れ、シートに `batchId` 付きで残る  
3. Yahoo明細に「親のみレ点」由来の行が含まれない  
4. 承認・行否認・バッチ取消がUIからでき、シートの status と一致する  
5. 一連の操作のあと、商品マスタの値・楽天／Yahoo／Amazon 側に **意図しない変更がない**（EC API・CSV未呼び出し）  
6. `generateRakutenCSV` および Yahoo 出品関数が当該機能から呼ばれない（静的／実行ログで確認）  

---

## 9. 復元・トグル

| 手段 | 内容 |
|------|------|
| Property | 承認Webデプロイの無効化／許可メールを空に近い状態へ |
| シート | `▼承認キュー(出品①)` を削除またはリネームしてよい（本番出品データではない） |
| Git | 実装コミットを `git revert` |

推奨 Property: `APPROVAL_QUEUE_V1_ENABLED`（既定 `false`。true のときだけ候補作成メニューが動く）

---

## 10. 実装時の承認パッケージ（次チケット）

実装に入る前に提示するもの:

- 変更ファイル一覧（想定: `コード.js` または新規 `ApprovalQueue.js`、`appsscript.json` の webapp 必要なら、本 docs）  
- 概要: シート＋Webのみ。EC書込なし  
- リスク: 誤って出品関数を呼ぶ／マスタを更新する → コードレビューで呼出禁止を確認  

---

## 11. 人間向け検証手順（実装後）

1. `git pull` → `clasp push`（Agentはしない）  
2. Property: 許可メール・`APPROVAL_QUEUE_V1_ENABLED=true`  
3. マスタにテスト用レ点（Yahooは子、楽天は親）  
4. 候補作成 → Webで承認 → シート確認  
5. 取消 → status=`CANCELLED`  
6. マスタとECが無変更であることを確認  
7. 終わったら Property を false に戻してよい  

---

## 12. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-17 | 初版。Lv0承認後の Lv1 要件。レ点／承認／スキップ分離。EC書込なし。 |
