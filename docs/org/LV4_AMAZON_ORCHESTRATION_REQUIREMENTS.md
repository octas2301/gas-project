# Lv4 Amazon バルク掲載（2トラック）— 要件定義

**文書種別**: 要件定義ドラフト（**コード未実装**・実装は別承認）  
**最終更新**: 2026-07-20  
**親**: [LEVELLED_IMPLEMENTATION_PLAN.md](LEVELLED_IMPLEMENTATION_PLAN.md) ・ [AI_APPROVAL_MATRIX.md](AI_APPROVAL_MATRIX.md) ・ [LV1_APPROVAL_QUEUE_REQUIREMENTS.md](LV1_APPROVAL_QUEUE_REQUIREMENTS.md) ・ [LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md) ・ [LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md](LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md) ・ [PHASE0_THREE_REVIEW_MAJORITY.md](PHASE0_THREE_REVIEW_MAJORITY.md) §4.2・§4.3  
**三点レビュー**: [LV4_THREE_REVIEW_MAJORITY.md](LV4_THREE_REVIEW_MAJORITY.md)（2026-07-20・条件付き→社長確定反映済）  
**列メモ（参考）**: [AMAZON_REQUIREMENTS.md](../AMAZON_REQUIREMENTS.md)（列メモは維持。**トラック別識別子・在庫書込方針は本ドキュメントが正**）  
**コード（実装時想定）**: GAS=`AmazonApprovalExport.js`（新規）／`ApprovalQueue.js` Amazon読取／`コード.js` メニュー21相当のみ／**楽天CSV・`Yahoo.js` 非改変**。**純正 `.xlsm` 埋めは Cursor外（Claude Code／Codex／Python 等）**（§1.4）。  
**ゴール一文**: 承認①済みの Amazon **親SKU（＋同一親の承認済み子）** から、GASで埋め用データを `GENERATED` → ローカルで純正テンプレ `.xlsm` を `PACKAGED` → 人間が SC 手動UP → `UPLOADED_OK`。M1はノーブランド新規の **バリエーション＋画像**を主対象。  
**手動逃げ道**: Seller Central 上の手動カタログ登録・既存の手作業バルクは **継続利用可**。  
**社長Q&A（第2回三点レビュー後）**: §17。

---

## 0. マイルストーン優先（確定 2026-07-20）

| 順 | トラック | 内容 | 検収の核 |
|----|----------|------|----------|
| **M1** | **B: ノーブランド新規カタログ** | **親＋子バリエーション**＋画像。純正 `.xlsm`（PACKAGED） | `GENERATED`→`PACKAGED`→SC手動UP→`UPLOADED_OK` |
| **M2** | **A: 既存カタログ出品** | 既存ASINオファー。**単品が多い**（まれにバリエーション） | 同上 |

**M1先行の位置づけ**: 憲章・Lvプランの「在庫0/1バルク最小版」のうち、現場で最も時間がかかるのは **ノーブランド新規のバリエーション＋画像**であるため M1=B を先に通す。既存ASINへのオファー（M2=A）は憲章の3ヶ月像にも含まれるが **M1検収後**に回す。M2未着手の間は「既存カタログの日中無人掲載」は手作業継続。

実装承認・コード着手は **本要件の社長確認（§11）＋多数決採用項反映済**のうえ、**対象カテゴリの GTIN免除承認**を必須ゲートとする（[LV4_THREE_REVIEW_MAJORITY.md](LV4_THREE_REVIEW_MAJORITY.md)）。

---

## 1. スコープ

### 1.1 作るもの

| # | 成果物 |
|---|--------|
| 1 | 承認キューから **mall=`amazon` かつ lineStatus=`APPROVED`** を読む入口（**親SKU単位**・§3.1） |
| 2 | **ルート判定**（§3.2）: A=既存ASIN出品 / B=ノーブランド新規カタログ |
| 3 | **B（M1）**: バリエーション親＋承認済み子の埋め用データ（画像URL・価格・在庫0/1値・自己配送最低限） |
| 4 | **A（M2）**: 既存ASINオファー用の埋め用データ（単品が多い） |
| 5 | Drive 等への **埋め用データ保存**（GAS）→ **`GENERATED`** |
| 6 | **純正 `.xlsm` への流し込み**（ローカル AI／スクリプト・§1.4）→ **`PACKAGED`** |
| 7 | **Lv4専用実行ログ／状態シート**＋メニュー「アップロード成功を記録」→ **`UPLOADED_OK`** |
| 8 | 実行分割・調査用ログ＋失敗メール |
| 9 | 手動キック（メニュー21想定） |

### 1.2 作らないもの（禁止）

- `generateRakutenCSV` および楽天 Lv2 経路の改変  
- `Yahoo.js` 出品API本体および Lv3 経路の改変  
- B統合 `B_INTEGRATED_STEP_FUNCTIONS` の順序・境界変更  
- **Seller Central への自動アップロード**（初版。SP-API Feeds / Listings Items は後段）  
- **マスタ「在庫数」（出品用）への書込**（在庫0/1はバルク内の値のみ。楽天/Yahoo共有列への副作用防止）  
- **FBA 納品・切替の自動化**（後続要件。§9）  
- Keepa / Category Listings Report の **常時フルシード**（§4 限定運用）  
- 承認②（補充）・販売中SKUへの無人上書き（U1）  
- 広告（スポンサード等）  
- `clasp push` 自動化  
- **GAS から純正 `.xlsm` を直接編集**（初版の正は §1.4）  
- **マスタ JAN の消去・空化**（Bでもマスタ JAN は残す。バルク GTIN 列だけ空）  

### 1.3 聖域の守り方

```text
[GAS Lv4]
  → 承認済み Amazon 親SKU（＋同一親の承認済み子）を読む
  → ルート判定（A / B / 保留）
  → マスタ読取のみ（在庫はスキップ判定のみ・書込禁止。JANは消さない）
  → 埋め用データ＋メタを Drive へ → GENERATED

[ローカル AI / スクリプト（Claude Code・Codex・Python 等）]
  → 人手DL済みの純正 .xlsm をコピー
  → GENERATED データを流し込み → PACKAGED

[人間]
  → Seller Central 手動アップロード
  → 処理レポート確認 → メニューで UPLOADED_OK
```

既存の手動 Amazon 出品は **触らない**。出品CKスナップショットは **不要が原則**。**Cursor 内に閉じず、ローカル他ツールとの組み合わせを正とする**（D-1）。

### 1.4 成果物パイプライン（D-1・確定）

| 段階 | 担当 | 成果物 | 状態 |
|------|------|--------|------|
| 1 | GAS | 埋め用データ（列マッピング済み表／CSV等）＋ `subBatchId` メタ | `GENERATED` |
| 2 | ローカル AI／スクリプト | Seller Central **純正テンプレ `.xlsm`** に埋め込んだファイル | `PACKAGED` |
| 3 | 人間 | SC 手動UP＋処理レポート成功 | `UPLOADED_OK` |

- テンプレ形式の正: **純正 `.xlsm`（等）を埋めて出力**（Q1=A）。  
- GASは段階1まで。段階2は Excel を扱えるローカル環境。Cloud Agent 単独は想定しない。  
- `PACKAGED` は状態シートにファイル名・パス／ハッシュ・実施手段を追記してよい。  

---

## 2. 前提（着手条件）

| # | 条件 |
|---|------|
| 1 | **Lv1〜Lv3 人間検収完了**（2026-07-17〜20） |
| 2 | 現状どおり、マスタの最低限情報で手作業のノーブランド新規登録ができること（運用前提） |
| 3 | 本要件の社長確認（**§11**）および [LV4_THREE_REVIEW_MAJORITY.md](LV4_THREE_REVIEW_MAJORITY.md) 反映済 → **実装承認**後にコード着手 |
| 4 | **（必須ゲート）** 対象カテゴリで **GTIN免除（ブランド名「ノーブランド品」）**が承認済みで、**Lv4状態シートに証跡が記録済み**であること（Q13）。未承認・未記録カテゴリは B を実行しない |
| 5 | M1試験用 **Product Type / バリエーションテーマ**は検収前に1カテゴリ固定（§11。着手ゲートではない） |
| 6 | **Lv1 が `mall=amazon` の親行＋子行を同一承認①バッチに出せること**（Q7=A）。**Lv4実装承認と同一チケット必須**（候補0のまま M1 着手禁止）。現状は親→rakuten／子→yahoo の排他明細のみで amazon 行未生成のため、**加算生成**で埋める |

---

## 3. 入力・対象SKU・ルート判定

### 3.1 承認キュー・候補抽出（親SKU単位・確定）

- シート: `▼承認キュー(出品①)`（Lv1）  
- 対象: ヘッダ `status=APPROVED` かつ明細 `mall=amazon` かつ `lineStatus=APPROVED`  
- **Lv1への載せ方（Q7=A・確定）**: 楽天／Yahoo と **同じ流れ**で、**親行＋各子行を明細化し個別承認**。原則 **同一承認①バッチに3モール同時**に載せる（Q9=A）。  
  - 切り分け: 明細の `mall` 列＋ Lv2／Lv3／Lv4 は **別メニュー・別 `runId`／状態シート**のため一括承認でも原因追及可能。  
  - フォールバック（運用で切り分け不能と判明した場合のみ）: 楽天→Yahoo のあと Amazon だけ別バッチ（Q9=B）  
- **抽出単位（実行時）**: **親SKU**を単位にまとめ、同一親の **承認済み子SKU** をバリエーション候補とする（楽天と同型）  
  - 親だけ承認・子が0件 → スキップ（不完全生成禁止）  
  - **子SKU空の親のみ（単品行）**: 運用上ほぼ存在しない。M1（B）では **対象外**（`SKIPPED_INCOMPLETE_VARIATION`）。単品中心は **A（既存カタログ）** 側の話  
  - 子の一部のみ承認 → **承認済み子のみ**をバリエーションに含める（未承認子は載せない）。親行必須属性が欠ける場合は親ごとスキップ  
- `REJECTED` / `CANCELLED` / `ORPHANED` は実行しない  
- 実行直前再チェック: マスタに該当親／子が無い → ORPHAN扱い・スキップ（方式A）  
- 販売中かつ在庫>0 → **原則スキップ**（マスタ出品用「在庫数」の **読取のみ**で判定。補充は②）  

※ **着手依存**: `ApprovalQueue` に amazon（親＋子）抽出を追加する。空のままでは候補0。

### 3.1.1 実行直前再検証（生成前）

不足時は **ファイルを生成しない**（ログ＋スキップ）:

| 検証 | 内容 |
|------|------|
| 承認 | 取消・REJECTED になっていない |
| 親子 | 親SKU存在・承認済み子が1件以上（B） |
| 価格 | `販売価格amazon` が有効 |
| 識別子 | A: ASIN必須／B: GTIN列空＋免除カテゴリOK・ブランド=ノーブランド品。**出品者SKU=子SKU**。**メーカー型番=`メーカー品番`（空→子SKU）** |
| 画像 | メイン画像URLが空でない（到達確認の厳格度は§11） |
| 自己配送 | 配送テンプレート名＝`送料無料パターン`（変更可）。リードタイム列は初版なし |
| 販売中 | 在庫>0はスキップ |

### 3.2 ルート判定（A / B / 保留）

各 **親SKU**（＋承認済み子）について。**Property `APPROVAL_AMAZON_LV4_TRACK` が優先**（Q8）:

```text
IF TRACK 未設定（空・不明値）
  → 保留（実行しない）。ログ＋メール。ASIN有無で A/B を推定しない
ELSE IF TRACK = B（M1検収の既定）
  → 強制 B（ノーブランド新規）。マスタに ASIN があっても無視／オファーしない
ELSE IF TRACK = A
  → A のみ（既存ASINオファー）。B 条件でも新規カタログにしない
ELSE IF TRACK = BOTH
  IF 既存ASINに出品可能（マスタASIN一致・出品制限なし・ブランドゲートなし）
    → A: 既存カタログ出品（M2）
  ELSE IF ノーブランド新規（GTIN免除済み・ブランド=ノーブランド品）
    → B: 新規カタログ作成（M1）
  ELSE
    → 保留（SKIPPED_BRAND_GATE / SKIPPED_GTIN_EXEMPTION / SKIPPED_NEED_HUMAN）
       ログ＋メール。自動で押し切らない
```

M1検収は **`TRACK=B` を明示設定**（未設定は実行しない）。ASIN付き親もノーブランド新規として進める。
### 3.3 行単位（Amazon）

| トラック | 単位 | 備考 |
|----------|------|------|
| **B（M1）** | **親＋承認済み子（バリエーション）** | 単品親（子SKU空）は対象外。不完全な親子は生成禁止 |
| **A（M2）** | **SKU＋対象ASIN** | **単品・バリエーションなしが多い**（まれにバリエーション） |

### 3.4 在庫0/1（マスタ書込禁止・確定）

- 承認キュー **ヘッダ** `inventoryMode`: `ZERO`（既定）/ `ONE`  
- **バルクの在庫列**にはヘッダ値をそのまま出す（Q11=A）: `ZERO`→**0**／`ONE`→**1**。未選択も ZERO 扱い  
- **マスタの出品用「在庫数」へは書き込まない**（共有列による楽天/Yahoo副作用防止）  
- スキップ判定のためマスタ在庫は **読取のみ**（計算列「在庫数計算」と出品用「在庫数」の区別は Lv3 どおり）  
- **販売可能数への引き上げはしない**（承認②）  
- **販売中SKUの無人上書き（U1）は当面手動・後送り**（Q11b=A）。既に Amazon で売れている（マスタ在庫>0 等でスキップ対象）SKUへの更新バルクは Lv4 では出さない。`inventoryMode` は **新規／再掲載用バルクに載せる在庫値**の話に限る  

### 3.5 データ正本（マスタ）

**常時の入力源は Keepa ではなく `▼商品マスタ(人間作業用)`**（現状の手作業出品と同じ最低限）。

| 用途 | マスタ列（例） | 備考 |
|------|----------------|------|
| 商品名（新規） | `オリジナルカタログ商品名` | Bトラック正本 |
| 価格 | `販売価格amazon` | CPO V2 済み想定 |
| 在庫（読取） | `在庫数`（出品用） | **書込禁止**。スキップ判定のみ |
| 在庫（出力） | （バルク列） | `inventoryMode` に従い 0または1 |
| 画像 | 楽天メイン／サブ画像URL 等 | URL方式（§11で検収前確認） |
| カテゴリ | `amazon カテゴリー` | Product Type前段 |
| 識別（A） | ASIN必須。JAN/EANは既存カタログ出品に使う（マスタに残す） | 欠けたら生成停止 |
| 識別（B） | **マスタ JAN は残す**（卸コード等・他モール用）。**Amazon バルクの GTIN／商品コード列だけ空** | 楽天／Yahoo と同じ商品でも Amazon ノーブランドは JAN なし登録。**Lv4はマスタ JAN を消さない**。GTIN免除必須。[AMAZON_REQUIREMENTS.md](../AMAZON_REQUIREMENTS.md) の「JAN必須」は **Bのバルク出力では本正本が上書き** |
| セット | `A.セット商品数` | バリエーション属性の材料 |
| ブランド（B） | 固定値 `ノーブランド品` | Keepaのbrandは使わない |
| 出品者SKU | **子SKU**（Q10b） | 親行はテンプレ仕様に従い親SKUまたは空 |
| メーカー型番（B） | マスタ **`メーカー品番`**（なければ `型番`）。**空なら子SKUにフォールバック**（Q10b=B） | 取得ロジック（15-⑤／⑥）は見直してよい。空のまま出さない |

列メモの詳細は [AMAZON_REQUIREMENTS.md](../AMAZON_REQUIREMENTS.md)。Product Type 必須列対応表はテンプレ固定後に本ドキュメントへ追記（§11）。

### 3.6 配送（初版・Q14確定）

- **既定: 自己配送（MFN）**（マスタに自己配送／FBAの選択がある場合はそれに従い、初版の主対象は自己配送）  
- **配送テンプレート名**: 固定 **`送料無料パターン`**（Script Property または設定マスタで変更可。Seller Central の登録名と完全一致必須）  
- **リードタイム（出荷作業日数）**: マスタに列なし。手作業も配送テンプレート側の設定に任せているため、**初版バルクでは出力しない**。純正テンプレ／処理レポートで必須エラーになった場合のみ、Property の固定日数を後付けで追加  
- FBA関連列はテンプレ上許される範囲で空欄または未送信  
- **FBA納品・切替は後続要件**（§9）

---

## 4. 限定ツール運用（通常は使わない）

| ツール | 通常 | 使うとき | 使い方 |
|--------|------|----------|--------|
| **Keepa** | 使わない | アップロードエラーの切り分け／参照ASINの属性確認 | 確認のみ。brand・title・features・画像の丸写し禁止。サイズ/重量は意図的に未登録（必須エラー時のみマスタから埋める） |
| **Category Listings Report** | 使わない | 既存SKUの編集・エラー時のテンプレ正当性確認 | 新規Bのたたき台にしない |
| **推奨値（Valid Values）シート** | 使わない | Browse node・バリエーションテーマ等の許容値確認／エラー時 | 辞書用途。常時全件照合はしない |

---

## 5. 実行分割

詳細の正（共通）: [AI_APPROVAL_MATRIX.md](AI_APPROVAL_MATRIX.md) §5。

| 優先 | 規則 | 内容 |
|------|------|------|
| **主** | 時間 | 1実行あたり実働 **約25分**で中断→Script Properties に再開位置→トリガー再開 |
| **副** | SKU件数 | 仮置き（親N件）。楽天ユニーク画像50は楽天制約。AmazonはPropertyで仮置きし実測調整 |
| — | 12:00 | 未完了は翌朝続き（明示取消まで有効） |
| — | 冪等 | 同一 `batchId`+親SKU(+track) で **`GENERATED` 済みの成功分は原則再生成しない**。`UPLOAD_FAILED`（または SC 失敗）後は **親単位で再入場し、新 `subBatchId` で GENERATED**（旧行は追記履歴として残す）。**監査ログ・状態シート行は追記のみ（消さない）**。`UPLOADED_OK` 済みは再記録不要 |

`subBatchId` 例: `{batchId}_A{n}`／`{batchId}_B{n}`

---

## 6. 状態・ログ

### 6.1 実行状態（例）

| 状態 | 意味 |
|------|------|
| `PENDING_RUN` / `RUNNING` / `FAILED` / `RETRYING` | 生成処理の進行 |
| `GENERATED` | GAS が埋め用データを Drive 等へ保存成功 |
| `PACKAGED` | ローカルで純正 `.xlsm` への埋め完了（任意記録可・検収では推奨） |
| `UPLOADED_OK` | 人間が処理レポート成功を確認し、メニューで記録した |
| `UPLOAD_FAILED` | 人間が失敗を記録。**親単位で再入場し、新 `subBatchId` で再 GENERATED 可**。ログは追記 |
| `SKIPPED_*` | `SKIPPED_IN_STOCK` / `SKIPPED_ORPHAN` / `SKIPPED_BRAND_GATE` / `SKIPPED_GTIN_EXEMPTION` / `SKIPPED_NEED_HUMAN` / `SKIPPED_INCOMPLETE_VARIATION` |

**検収・冪等で「完了」と言わないこと**: `GENERATED`／`PACKAGED` のみでは Seller Central 反映完了ではない。掲載完了の主張は `UPLOADED_OK` 後（かつ§11のU5未検証時は断言しない）。

### 6.1.1 失敗後の再生成（Q5・Q12・確定）

- SC 処理レポート失敗／`UPLOAD_FAILED` のあと → **親単位で再入場し、新 `subBatchId` で GENERATED**（旧 `UPLOAD_FAILED` 行は残す。同一ID上書きではない）  
- **部分成功時のやり直し単位は親SKU一式**（Q12=A）: 成功した子だけ残して行単位で切らない。当該親のバリエーション全体をまとめて再 PACKAGED／再UP  
- 状態の付け方（推奨）: 親単位で成功ならその親を `UPLOADED_OK`、失敗親は `UPLOAD_FAILED`。M1初回は `PARENTS_PER_SUB=1` 推奨（親＝サブバッチ）  
- **21-③ は最新 GENERATED の `subBatchId` を記録**（旧番号を誤入力しない）  
- **調査用ログ・状態シートの履歴行は残す**（再生成回数・理由・時刻を追記）  
- Drive 上の旧ファイルは版履歴または別名退避を推奨（完全物理削除はしない）

### 6.2 Lv4専用状態シート（案B・確定）

- シート名: **`▼Lv4実行ログ(Amazon)`**（確定）  
- 粒度: **`subBatchId` 単位**（親SKU一覧・track・ファイル名・Drive URL・`GENERATED_at`・`PACKAGED_at`・`UPLOADED_OK_at`・メモ）  
- メニュー: **21-③ アップロード成功を記録**／**21-④ 失敗記録**  
- **GTIN免除証跡（Q13）**: 同シートの `recordType=EXEMPTION` 行に **カテゴリ実値または `*`**／`ノーブランド品`／**承認日**／**証跡URL** をすべて記入。マスタの amazon カテゴリが空の親は B 不可（fail-closed）。テンプレ行の `（カテゴリ or *）` は無効。  
- **冪等**: 同一 `batchId`+親+track で最新が `GENERATED`/`UPLOADED_OK`/`PACKAGED` の親は再生成しない。`UPLOAD_FAILED` が最新の親のみ再生成可。`DRY_RUN`（SKIP_EXPORT）と `SKIP` 行はブロック対象外。状態は追記のみ。`subBatchId` は単調増加。  
- **ブランド（B）**: 免除証跡の brand は **`ノーブランド品` 完全一致**（新規カタログ限定）。  
- 目的: 原因追及・調査・着手ゲートの検収。Logger と二重でもよいが、シートは人間が追える正本とする  

### 6.3 Logger 必須

- `runId` / `batchId` / `subBatchId` / `functionName` / `state` / `track`（A|B）  
- 対象親件数・子件数・生成ファイル名・スキップ理由  
- UPLOADED_OK 記録時も `subBatchId`・操作者想定・時刻をログ  
- **シークレット・トークン全文は出さない**

### 6.4 メール

- サブバッチ失敗・ファイル生成失敗・25分切断後の再開失敗 → **メール必須**（件名に「Lv4 Amazon」）

---

## 7. Bトラック詳細（M1・ノーブランド新規）

### 7.1 目的

手作業で特に時間がかかっている **バリエーション設定**と **画像設定**をバルクで行う。

### 7.2 必須成果物（バルク内）

| 項目 | 要件 |
|------|------|
| Product Type | 試験カテゴリで1つ固定（§11）。テンプレート列構造に合わせる |
| ブランド | `ノーブランド品`（GTIN免除と一致） |
| GTIN / JAN | **マスタ JAN は変更しない**。**バルクの GTIN／商品コード列は空**（免除済み前提） |
| 親SKU / 子SKU | マスタどおり。親子関係・バリエーションテーマを埋める |
| 出品者SKU | **子SKU** |
| メーカー型番 | **`メーカー品番`**（空なら **子SKU**） |
| バリエーションテーマ | 許容値を確認したうえで固定。実装検収前に1つ決定 |
| 画像 | メイン＋サブを **URL列**で埋める |
| 価格・在庫 | `販売価格amazon`／バルク内在庫0または1（**マスタ非書込**） |
| 自己配送 | **配送テンプレート**＝`送料無料パターン`（Property／設定マスタで変更可）。**リードタイム列は初版出力しない** |

### 7.3 意図的に入れないもの（初版）

- パッケージ寸法・重量（必須エラー時のみマスタから埋める）  
- FBA関連  
- Keepa由来のブランド・文言・画像の転載  

### 7.4 エラー時ループ

1. Seller Central の処理レポートでエラー分類  
2. 必須属性欠落 → マスタ列を補完して再生成（`GENERATED` 冪等の例外として失敗側のみ）  
3. 許容値不一致 → 推奨値シートで確認  
4. 属性の意味不明 → Keepaで参照ASINを確認（転載しない）  
5. 小バッチ再UP → 成功後に `UPLOADED_OK` 記録  

---

## 8. Aトラック詳細（M2・既存カタログ）

### 8.1 目的

既存ASINに対し、在庫0/1値・価格・SKUでオファーを載せる（Offer Only / Inventory Loader 系）。

### 8.2 必須成果物

| 項目 | 要件 |
|------|------|
| 対象ASIN | マスタ `ASINコード` 等（欠けたら生成停止） |
| 出品者SKU | 子SKU（または商品管理番号ルール） |
| 価格・在庫 | 価格はマスタ／在庫はバルク内0または1 |
| コンディション | 新品等の固定値 |
| 自己配送 | **配送テンプレート**＝`送料無料パターン`（変更可）。リードタイム初版なし |

### 8.3 関門

- **ブランド認証**・出品制限 → `SKIPPED_BRAND_GATE`  
- Category Listings Report は既存SKU更新用途に限定  

---

## 9. 後送り（本Lv4の範囲外）

| 項目 | 扱い |
|------|------|
| FBA納品・自己配送→FBA切替 | 後続要件 |
| SP-API（Listings Items / JSON_LISTINGS_FEED / VALIDATION_PREVIEW） | 手動UPが安定してから |
| 広告 | [AMAZON_REQUIREMENTS.md](../AMAZON_REQUIREMENTS.md) §4どおり別要件 |
| 販売中SKUの無人上書き | U1・当面手動 |
| 全カテゴリ横断のテンプレ自動選択 | M1は1 Product Type固定から |
| マスタ在庫への0/1書込 | **採用しない**（社長確定） |

---

## 10. エントリポイント（実装時想定）

| 種別 | 内容 |
|------|------|
| メニュー | Z → **21. 承認①済→Amazonバルク(Lv4)** → 21-① GENERATED／21-②状態クリア／21-③ UPLOADED_OK／21-④ UPLOAD_FAILED |
| Property | `APPROVAL_AMAZON_LV4_ENABLED`（既定 `false`） |
| 任意 | `APPROVAL_AMAZON_LV4_SKIP_EXPORT`／`APPROVAL_AMAZON_LV4_TRACK`（**M1は明示 `B`。未設定＝実行しない**）／`APPROVAL_AMAZON_LV4_PARENTS_PER_SUB`／**`APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE`**（既定 `送料無料パターン`）／`APPROVAL_AMAZON_LV4_FOLDER_ID`／（エラー時のみ）リードタイム日数。**在庫マスタ書込トグルは設けない** |
| トリガー | 実働25分超過時に自動再開 |
| 出力 | Drive フォルダ `Lv4_Amazon_GENERATED`（または FOLDER_ID）に CSV＋meta JSON。**アップロードは人間** |
| 状態シート | `▼Lv4実行ログ(Amazon)` に `GENERATED` / `UPLOADED_OK` / `EXEMPTION` |

---

## 11. 実装前／検収前に閉じる未決（社長確認チェックリスト）

| # | 未決 | 仮運用（閉じるまで） | 着手ゲートか |
|---|------|----------------------|--------------|
| 1 | **GTIN免除**（カテゴリ×ノーブランド品）＋**状態シートへ証跡記録**（Q13） | 未記録・未承認カテゴリはB保留 | **必須（着手）** |
| 2 | 画像はテンプレ上 **URL列で通るか** | 検収前に手動1件で確認 | 検収前 |
| 3 | M1試験用 **Product Type / バリエーションテーマ** | 手作業成功テンプレを正とする | 検収前 |
| 4 | Browse node | 試験テンプレの値を固定 | 検収前 |
| 5 | 純正 `.xlsm` の版・必須列対応表（GENERATED列↔テンプレ列） | 手作業成功テンプレを正とし、対応表を本docsへ追記してからM1検収 | 検収前 |
| 6 | Amazon在庫0/1の実機見え方（U5） | 1SKU確認前に「掲載完了」と断言しない | 検収前 |
| 7 | A/B同時実行 | M1は `TRACK=B` のみ（ASINあっても強制B・Q8） | — |
| 8 | PACKAGED 工程の標準ツール | ローカルで再現可能な手順書を検収前に1つ固定 | 検収前 |
| 9 | リードタイム列がテンプレ必須になるか | 初版は未出力。必須エラー時のみ Property 固定日数を追加 | エラー時 |
| 10 | 出品CKスナップショット | **不要が原則** | — |

---

## 12. 検収条件

### 12.1 M1（Bトラック）— GENERATED（GAS）

1. Property OFF ではメニューが動かない（または即return）  
2. 承認①→埋め用データが Drive に出力され、状態が `GENERATED`  
3. マスタ「在庫数」・**マスタ JAN が書き換わっていない**  
4. Bデータでバルク用 GTIN 列が空になることがメタ／サンプルで確認できる  
5. 楽天・Yahoo 聖域を壊していない  
6. 必須3点セット済み  

### 12.2 M1 — PACKAGED（ローカル）

1. 純正 `.xlsm` にバリエーション＋画像等が埋められ、SC が受け付ける形になっている  
2. 状態シートに `PACKAGED`（または同等のファイル記録）があること（推奨）  

### 12.3 M1 — UPLOADED_OK

1. Seller Central 手動UP→処理レポート成功  
2. メニューで `UPLOADED_OK`  
3. （§11-6）在庫0の見え方は未検証なら「掲載完了」と断言しない  

### 12.4 M2（Aトラック）完了（M1後）

1. 既存ASINオファー用データが `GENERATED`→`PACKAGED` できる  
2. ブランドゲートはスキップ  
3. `UPLOADED_OK` まで同様  

### 12.5 共通完了

- 検収終了後 **`APPROVAL_AMAZON_LV4_ENABLED=false` に戻す**

 ### 12.6 人間検収記録

（実装済・検収待ち 2026-07-20。`clasp push` 後に追記）

**実装ファイル**: `AmazonApprovalExport.js` / `ApprovalQueue.js`（amazon加算）/ `コード.js`（メニュー21）

---

## 13. 実装時の承認パッケージ（コード着手時に提示）

- 変更ファイル: `AmazonApprovalExport.js`（新規）、`ApprovalQueue.js`（**amazon 親＋子明細の加算生成・必須**）、`コード.js`（メニュー21相当）、本docs・状態シート  
- 概要: Lv1 amazon抽出＋親単位読取・ルート判定（TRACK未設定は非実行）・B優先バルク生成・マスタ在庫非書込・`GENERATED`／`UPLOADED_OK`・調査ログ  
- リスク: 誤トラックでGTINやブランドを載せる／URL画像失敗／テンプレ列ずれ／amazon候補がキューに載らない  
- **同一チケット必須**: `ApprovalQueue` の `mall=amazon` 親＋子抽出が無い状態では Lv4 実行コードだけをマージしない  

**復元**: Property `APPROVAL_AMAZON_LV4_ENABLED=false`／`git revert`／新規js削除＋メニュー削除。

---

## 14. 人間向け検証手順（実装後・概要）

1. Lv1〜3検収済みであること（GTIN免除証跡は本 GENERATED 前に状態シートへ）  
2. `clasp push`  
3. Property: `APPROVAL_AMAZON_LV4_ENABLED=true`、`APPROVAL_AMAZON_LV4_TRACK=B`、**初回は `APPROVAL_AMAZON_LV4_SKIP_EXPORT=true`（ドライラン）**。任意で `PARENTS_PER_SUB=1`  
4. `▼Lv4実行ログ(Amazon)` の EXEMPTION に brand=`ノーブランド品`・承認日・証跡URL・カテゴリ（または `*`）を記入  
5. 18-①で候補作成→承認①に amazon 親＋子 APPROVED。21-②で STATE 空を確認  
6. 21-① → 状態が `DRY_RUN`。マスタ在庫／JAN不変。楽天/Yahoo候補件数の非回帰を1回確認可  
7. 問題なければ `SKIP_EXPORT=false` で再実行 → `GENERATED`（Drive CSV）  
8. ローカルで純正 `.xlsm` を PACKAGED  
9. Seller Central 手動UP → **21-③に最新 GENERATED の `subBatchId`（例 `{batchId}_B1`）を入力**  
10. Property を `ENABLED=false` に戻す  

---

## 15. 調査ソース（設計根拠）


| テーマ | URL |
|--------|-----|
| Listings Items API | https://developer-docs.amazon.com/sp-api/docs/listings-items-api |
| Listings APIs FAQ | https://developer-docs.amazon.com/sp-api/docs/listings-apis-faq |
| Product Type Definitions | https://developer-docs.amazon.com/sp-api/docs/product-type-definitions-api |
| Inventory file templates | https://sellercentral.amazon.com/help/hub/reference/G1641 |
| 製品コードがない商品（日本） | https://sellercentral-japan.amazon.com/gp/help/200426310 |
| Keepa API | https://keepa.com/#!api |

---

## 16. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-21 | **合意反映**: Q5/§5/§6.1.1 を「失敗後は新subBatchId」に統一。§14をドライラン手順に更新。clasp push へ。 |
| 2026-07-20 | **再レビュー採用修正**: DRY_RUN専用status／冪等latest汚染除去（SKIP recordType）／subBatchId単調増加／ブランド=ノーブランド品完全一致／archive失敗は停止。 |
| 2026-07-20 | **実装レビュー修正**: レジュームindex0／価格画像SKIPPED記録／GTIN空カテゴリfail-closed・完全一致／GENERATED冪等／追記専用ステータス／clear禁止。 |
| 2026-07-20 | **実装着手（承認済）**: `AmazonApprovalExport.js` 新規／`ApprovalQueue` amazon加算／メニュー21。コード実装。 |
| 2026-07-20 | **第3回三点レビュー後**: TRACK未設定＝実行しない／Lv1 amazon抽出は同一実装チケット必須／列メモとの相互注記。 |
| 2026-07-20 | **Q11–Q14反映**: inventoryMode同期・販売中上書きは手動維持・部分失敗は親SKU単位・GTIN証跡は状態シート・配送テンプレ=送料無料パターン・リードタイム初版なし。 |
| 2026-07-20 | **Q7–Q10b反映**: Lv1は親＋子個別承認・3モール同一バッチ（切り分けはmall+runId）。TRACK=BはASIN無視で強制B。出品者SKU=子SKU／メーカー型番=メーカー品番（空→子SKU）。 |
| 2026-07-20 | **社長Q&A反映（§17）**: 純正xlsm・D-1（GAS+ローカルPACKAGED）・BはマスタJAN残しGTIN列のみ空・M1バリエーションのみ・失敗後同一subBatchId上書き＋ログ追記。 |
| 2026-07-20 | **三点レビュー反映**（[LV4_THREE_REVIEW_MAJORITY.md](LV4_THREE_REVIEW_MAJORITY.md)）。在庫マスタ書込禁止・DONE分離・親SKU抽出・UPLOADED_OK専用シート・GTIN着手ゲート。 |
| 2026-07-20 | **初版ドラフト**。M1=B／M2=A。Keepa・レポート・推奨値は限定運用。自己配送既定・FBA後送り。コード未実装。 |

---

## 17. 社長Q&A確定事項（第2回三点レビュー後・2026-07-20）

| ID | 決定 |
|----|------|
| Q1 | ファイル形式 = **Seller Central 純正テンプレ（`.xlsm` 等）を埋めて出力** |
| Q2→Q4 | Bの「JANなし」= **Amazon バルクの GTIN 列のみ空**。**マスタ JAN は残す**（楽天／Yahoo 同一商品用）。既存カタログ（A）は JAN 登録する |
| Q3 | M1（B）= **バリエーションのみ**。子SKU空の親のみは運用上ほぼ無く対象外 |
| Q5 | 失敗後 = **親単位で再入場し、新 `subBatchId` で GENERATED**（旧行は追記保持）。**調査ログは追記で残す** |
| Q6' | **D-1**: 初版から **GAS=`GENERATED` ＋ ローカルAI/スクリプト=`PACKAGED(.xlsm)`**。Cursor外組み合わせを正とする |
| Q7 | **A**: 親行＋各子行を明細化し個別承認。楽天／Yahoo と同じ承認①フロー |
| Q8 | **A**: `TRACK=B` 時は強制 B。マスタ ASIN は無視／オファーしない |
| Q9 | **A**: 楽天／Yahoo と **同一承認①バッチ**に amazon も載せる（3モール一括承認）。明細 `mall`＋モール別 `runId` で切り分け。切り分け不能時のみ別バッチ（B）をフォールバック |
| Q10 | メーカー型番 ← マスタ **`メーカー品番`**（取得ロジック見直し可） |
| Q10b | **B**: 出品者SKU＝**子SKU**／メーカー型番＝**メーカー品番**（空なら **子SKU** フォールバック） |
| Q11 | **A**: 承認①ヘッダ `inventoryMode` をバルク在庫に反映（ZERO→0／ONE→1） |
| Q11b | **A**: 在庫値の話のみ。販売中SKUの無人上書き（U1）は当面手動。Lv4はスキップ維持 |
| Q12 | **A**: 部分失敗時のやり直し単位＝**親SKU一式**（バリエーションまとめて） |
| Q13 | **A**: GTIN免除証跡は **Lv4状態シート**にカテゴリ／ノーブランド品／承認日／URL or ケースIDを記録してから B 実行可 |
| Q14 | MFN必須＝配送テンプレート **`送料無料パターン`**（変更可）。**リードタイムは初版出力しない**（必須エラー時のみ Property 固定日数を後付け） |
| Q15 | **TRACK 未設定＝実行しない**（明示 `A`/`B`/`BOTH` のみ）。誤A防止 |
| Q16 | **Lv1 amazon 親＋子抽出は Lv4 実装承認と同一チケット必須**（候補0で着手禁止） |

**未回答（次ラウンド）**: 方針Q&Aは閉じた。検収前§11（Product Type固定・画像URL・PACKAGED手順書・U5）と列対応表が残作業。

