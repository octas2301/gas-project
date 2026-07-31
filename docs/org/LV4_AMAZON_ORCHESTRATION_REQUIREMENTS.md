# Lv4 Amazon バルク掲載（2トラック）— 要件定義

**文書種別**: 要件定義ドラフト（**コード未実装**・実装は別承認）  
**最終更新**: 2026-07-30
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

**当面のD手動特則（2026-07-30）**: AIが対象を自動決定しない期間は、人間が付けた子SKU `出品CK` を承認①相当として、DからB（新規）／A（既存相乗り）を実行できる。ApprovalQueue経路は削除せず、AIレ点・無人実行・トリガー接続前に再接続する。詳細は [LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)。

### 3.1.1 実行直前再検証（生成前）

不足時は **ファイルを生成しない**（ログ＋スキップ）:

| 検証 | 内容 |
|------|------|
| 承認 | 取消・REJECTED になっていない |
| 親子 | 親SKU存在・承認済み子が1件以上（B） |
| 価格 | `販売価格amazon` が有効 |
| 識別子 | A: N列`ASINコード`必須・**出品者SKU=`Amazon相乗りSKU`**／B: GTIN列空＋免除カテゴリOK・ブランド=ノーブランド品・**出品者SKU=子SKU**。**メーカー型番=`メーカー品番`（空→子SKU）** |
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
- **Dレ点新規（`source=child_ck`）は在庫>0でもスキップしない**（2026-07-30 承認）。別カタログ（ノーブランドセット）を作るため。バルクの在庫列は `inventoryMode` 準拠で 0／1、マスタ在庫は非改変。承認①経路は従来どおり `SKIPPED_IN_STOCK`。戻し: Property `APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK=false`  
- **販売可能数への引き上げはしない**（承認②）  
- **販売中SKUの無人上書き（U1）は当面手動・後送り**（Q11b=A）。既に Amazon で売れている（マスタ在庫>0 等でスキップ対象）SKUへの更新バルクは Lv4 では出さない。`inventoryMode` は **新規／再掲載用バルクに載せる在庫値**の話に限る  

### 3.5 データ正本（マスタ）

**常時の入力源は Keepa ではなく `▼商品マスタ(人間作業用)`**（現状の手作業出品と同じ最低限）。

| 用途 | マスタ列（例） | 備考 |
|------|----------------|------|
| 商品名（新規） | `オリジナルカタログ商品名` | Bトラック正本 |
| 価格 | `販売価格amazon` | 親が空なら **承認済み子の同列**へフォールバック（Logger `[Lv4Price]`）。CPO V2 済み想定 |
| 在庫（読取） | `在庫数`（出品用） | **書込禁止**。スキップ判定のみ |
| 在庫（出力） | （バルク列） | `inventoryMode` に従い 0または1 |
| 画像 | 楽天メイン／サブ画像URL 等 | URL方式（§11で検収前確認） |
| カテゴリ | `amazon カテゴリー`（なければ `amazonカテゴリー`、さらに **`カテゴリー`（T列・注記付き見出し可）**） | Product Type前段。解決元は Logger `[Lv4Cat]` |
| 識別（A） | ASIN必須。JAN/EANは既存カタログ出品に使う（マスタに残す） | 欠けたら生成停止 |
| 識別（B） | **マスタ JAN は残す**（卸コード等・他モール用）。**Amazon バルクの GTIN／商品コード列だけ空** | 楽天／Yahoo と同じ商品でも Amazon ノーブランドは JAN なし登録。**Lv4はマスタ JAN を消さない**。GTIN免除必須。[AMAZON_REQUIREMENTS.md](../AMAZON_REQUIREMENTS.md) の「JAN必須」は **Bのバルク出力では本正本が上書き** |
| セット | `A.セット商品数` | バリエーション属性の材料 |
| ブランド（B） | 固定値 `ノーブランド品` | Keepaのbrandは使わない |
| 出品者SKU（B/M1・新規） | **子SKU**（Q10b） | 親行はテンプレ仕様に従い親SKUまたは空 |
| 出品者SKU（A/M2・既存相乗り） | **`Amazon相乗りSKU`（NF列）** | 子SKU中央の識別値をN列ASINへ置換し、`s1/f1` を `as1/af1` へ変換。dry_run VALID/issues=0後にGAS保存、prodは保存値を再利用 |
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
| **Category Listings Report** | 使わない（常時フルシード禁止） | 既存SKUの編集・エラー時のテンプレ正当性確認。**成功後の受理値辞書化は §11.5.4（後続・未着手）** | 新規Bのたたき台にしない |
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
| `NEEDS_VARIATION_LINK` | SC上で子（または親）は掲載済みだが **バリエーション家族が未完成**。**同一SKUのまま**修正登録（§6.1.2） |
| `CORRECTIVE_UPDATE` | 修正用 PACKAGED をSCへUPし、処理レポートで親子付け直しを確認した（`UPLOADED_OK` 前の中間可） |
| `SKIPPED_*` | `SKIPPED_IN_STOCK` / `SKIPPED_ORPHAN` / `SKIPPED_BRAND_GATE` / `SKIPPED_GTIN_EXEMPTION` / `SKIPPED_NEED_HUMAN` / `SKIPPED_INCOMPLETE_VARIATION` |

**検収・冪等で「完了」と言わないこと**: `GENERATED`／`PACKAGED` のみでは Seller Central 反映完了ではない。掲載完了の主張は `UPLOADED_OK` 後（かつ§11のU5未検証時は断言しない）。**SKU成功＝バリエーション完成ではない**（親子未リンクでも 7/7 成功になり得る）。

### 6.1.1 失敗後の再生成（Q5・Q12・確定）

- SC 処理レポート失敗／`UPLOAD_FAILED` のあと → **親単位で再入場し、新 `subBatchId` で GENERATED**（旧 `UPLOAD_FAILED` 行は残す。同一ID上書きではない）  
- **部分成功時のやり直し単位は親SKU一式**（Q12=A）: 成功した子だけ残して行単位で切らない。当該親のバリエーション全体をまとめて再 PACKAGED／再UP  
- **ただし SC に既に載っている出品者SKUは原則変更しない**（§6.1.2）。新規SKUでの作り直しは最終手段  
- 状態の付け方（推奨）: 親単位で成功ならその親を `UPLOADED_OK`、失敗親は `UPLOAD_FAILED`。親子未リンクのみなら `NEEDS_VARIATION_LINK`。M1初回は `PARENTS_PER_SUB=1` 推奨（親＝サブバッチ）  
- **21-③ は最新 GENERATED の `subBatchId` を記録**（旧番号を誤入力しない）  
- **調査用ログ・状態シートの履歴行は残す**（再生成回数・理由・時刻を追記）  
- Drive 上の旧ファイルは版履歴または別名退避を推奨（完全物理削除はしない）

### 6.1.2 修正登録（SKU維持・確定・2026-07-22）

Amazon SC では「SKU成功」でも **親がフィード対象外・親子未リンク** になり、子が単品バラ出品になることがある。その都度 SKU を変えて新規出品するのは非効率なため、次を **正**とする。

| 原則 | 内容 |
|------|------|
| **SKU不変** | 既に SC に存在する **子SKU／親SKU は残す**。付け直し・属性修正は **同一出品者SKU** で行う |
| **優先手段** | 親追加＋既存子の親SKU紐づけ（出品情報アクション＝**作成または置換／部分更新**）。削除→新SKU再登録は最終手段 |
| **やり直し単位** | 親SKU一式（Q12）。成功済み子だけ切り捨てて別SKU化しない |
| **GENERATED** | マスタ内容が変わらない修正登録は **再GENERATED不要**。ローカル再 PACKAGED（`packagePurpose=CORRECTIVE_LINK` 等）で可。内容変更時のみ新 `subBatchId` で GENERATED |
| **完了条件** | 処理レポート成功 **かつ** SC上でバリエーション家族が目視確認できたときだけ `UPLOADED_OK` |

**典型フロー（親子未リンク）**:

```
SC: 子のみ単品成功（親SKU未作成／未処理）
  → 状態 NEEDS_VARIATION_LINK（21-⑤）
  → 同一SKUで修正用 PACKAGED（親＋子・行構造は§11.5.3）
  → SC 手動UP（作成または置換）
  → 親子表示を目視確認 → 21-③ UPLOADED_OK
```

**禁止に近い運用**: バリエーション失敗のたびに子SKU接尾辞を変えて新規カタログを量産すること（在庫・広告・注文履歴が分断される）。

### 6.2 Lv4専用状態シート（案B・確定）

- シート名: **`▼Lv4実行ログ(Amazon)`**（確定）  
- 粒度: **`subBatchId` 単位**（親SKU一覧・track・ファイル名・Drive URL・`GENERATED_at`・`PACKAGED_at`・`UPLOADED_OK_at`・メモ）  
- メニュー: **21-③ アップロード成功を記録**／**21-④ 失敗記録**／**21-⑤ 修正登録（親子付け直し）を記録**（`NEEDS_VARIATION_LINK`／`CORRECTIVE_UPDATE`・§6.1.2。実装後続可・手順は先に正本化）  
- **GTIN免除証跡（Q13）**: 同シートの `recordType=EXEMPTION` 行に **カテゴリ実値または `*`**／`ノーブランド品`／**承認日**／**証跡URL** をすべて記入。マスタの amazon カテゴリが空の親は B 不可（fail-closed）。テンプレ行の `（カテゴリ or *）` は無効。
- **記録手段（2026-07-31 実装）**: メニュー **21-⑭ GTIN免除証跡を記録**。レ点新規の親からカテゴリ実値を検出し、人間が証跡を入力・確認して追記する（手入力も可）。有効な証跡があるカテゴリは追記しない。`*` は Property `APPROVAL_AMAZON_LV4_EXEMPTION_ALL_CATEGORIES=true` のときのみで、追加警告あり。**カテゴリ別記録を原則**とする。  
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
| 対象ASIN | マスタN列 `ASINコード` のみ（子→同一親の親行）。競合店ASIN／URLは使わない。欠けたら停止 |
| 出品者SKU | `Amazon相乗りSKU`。未作成時は子SKUから相乗り専用規則で生成し、dry_run VALID/issues=0後に保存 |
| 価格・在庫 | 価格はマスタ／在庫はバルク内0または1 |
| コンディション | 新品等の固定値 |
| 自己配送 | **配送テンプレート**＝`送料無料パターン`（変更可）。リードタイム初版なし |

### 8.3 関門

- **ブランド認証**・出品制限 → `SKIPPED_BRAND_GATE`  
- Category Listings Report は既存SKU更新用途に限定  
- 新ASIN型 `Amazon相乗りSKU` と旧JAN型子SKUは別sellerSku。各経路で同一sellerSkuが既存なら更新、無ければ新規登録する。別sellerSku同士は相互に上書きしない
- 保存済み `Amazon相乗りSKU` 内のASINと今回確定ASINが不一致なら自動上書きせず停止

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

**本線UX**: A〜D の **D モール選択に Amazon**（U3 v1: `amazon`／`full_amazon`・薄いファサード）は [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)／[D_MENU_U3_HUMAN_RUN.md](D_MENU_U3_HUMAN_RUN.md)。エンジン仕様の正本は本ファイルのまま。分割・復旧は下表の Z→21。

| 種別 | 内容 |
|------|------|
| メニュー | Z → **21. 承認①済→Amazonバルク(Lv4)** → 21-① GENERATED／21-②状態クリア／21-③ UPLOADED_OK／21-④ UPLOAD_FAILED／**21-⑤ 修正登録記録**（`NEEDS_VARIATION_LINK`／`CORRECTIVE_UPDATE`・SKU維持・§6.1.2）／**21-⑥ Drive→R2 PoC** |
| Property | `APPROVAL_AMAZON_LV4_ENABLED`（既定 `false`） |
| 任意 | `APPROVAL_AMAZON_LV4_SKIP_EXPORT`／`APPROVAL_AMAZON_LV4_TRACK`（**M1は明示 `B`。未設定＝実行しない**）／`APPROVAL_AMAZON_LV4_PARENTS_PER_SUB`／**`APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE`**（既定 `送料無料パターン`）／`APPROVAL_AMAZON_LV4_FOLDER_ID`／（エラー時のみ）リードタイム日数。**在庫マスタ書込トグルは設けない** |
| トリガー | 実働25分超過時に自動再開 |
| 出力 | Drive フォルダ `Lv4_Amazon_GENERATED`（または FOLDER_ID）に CSV＋meta JSON。**アップロードは人間** |
| 状態シート | `▼Lv4実行ログ(Amazon)` に `GENERATED` / `UPLOADED_OK` / `EXEMPTION` |

---

## 11. 実装前／検収前に閉じる未決（社長確認チェックリスト）

### 11.0 M1 HPC（HEALTH PERSONAL CARE）クローズ（2026-07-23・社長確定）

**一言**: HPC（アルコールパッチ）について **§11-1〜8・10 および U5 を閉じる**。正本は `…titlefix` 成功＋U5（在庫0）。**画像は ZIP 優先**。FOOD／他 Product Type／M2／21-⑤実装は **別ゲート**。

| 項目 | 値 |
|------|-----|
| 試験 | subBatchId `A1_20260721_083100_06b90a_B2`／親SKU `lifec-4560151300139-oya`／親ASIN `B0H9WZV641` |
| 成功成果物 | `…_PACKAGED_corrective_titlefix.xlsm`（processing-summary **8/8**）→ 親子目視 → **21-③ `UPLOADED_OK`** → **`APPROVAL_AMAZON_LV4_ENABLED=false`** |
| U5 | 2026-07-23 … SC在庫で子が **在庫切れ・FBM在庫=0**（例 `…-40s10`→`B0H9T9HPM1`）。**本試験は掲載完了と断言してよい** |
| 成功スナップショット | `Downloads/Lv4_Amazon_PACKAGED/accepted_values_db/success/…titlefix.json` |

| # | 項目 | HPC結論 | 状態 |
|---|------|---------|------|
| 1 | GTIN免除＋証跡 | EXEMPTION 記録済みで B 実行・成功。未記録カテゴリは引き続き B 保留 | **HPC分クローズ**（他PTは都度） |
| 2 | 画像 | R2 https だけでは **18320** になり得た。**本運用の正＝SC Upload Images ZIP（`{SKU}.MAIN.jpg`）**。R2はステージング／予備。取り込み後は Amazon 側画像が正 | **HPC分クローズ** |
| 3 | Product Type／テーマ | PT=`HEALTH_PERSONAL_CARE`／テーマ=`サイズ` | **クローズ** |
| 4 | Browse node | `ドラッグストア > 衛生用品・ヘルスケア > 検査キット (4520899051)` | **クローズ** |
| 5 | 列対応 | §11.5 ＋成功スナップショットを HPC 正本。マスタ必須候補方針は維持 | **HPC確定**（他PTは別表） |
| 6 | U5 在庫0/1 | 在庫0バルク → **在庫切れ・数量0** 確認済 | **クローズ** |
| 7 | A/B同時 | M1は `TRACK=B` のみ | **維持（クローズ扱い）** |
| 8 | PACKAGED標準 | 行5属性厳守／行6サンプル維持／行7〜実データ（7=親）／同一SKU修正（§6.1.2）／ハイライト空／電池フラグ＝いいえ×2／titlefix系。**手順書**: [LV4_HPC_M1_PACKAGED_RUNBOOK.md](LV4_HPC_M1_PACKAGED_RUNBOOK.md) | **方針クローズ** |
| 9 | リードタイム | HPC成功経路では未出力のまま可。必須エラー時のみ Property | **HPC仮クローズ**（未発火） |
| 10 | 出品CKスナップショット | 不要が原則 | **クローズ** |

**別ゲート（閉じない）**: FOOD／他 Product Type の §11 再確認／**M2（Track A）**／**21-⑤ GAS実装**／GENERATED へのテーマ・ハイライト等の列拡張（手埋め成功済み・自動化は別承認）。

**HPCから固定する運用ルール**

1. 「商品のハイライト」（`title_differentiation`）＝**空**。箇条書きは仕様／説明へ  
2. メーカー／タイトルに登録ブランド名を載せない（Track B・ノーブランド品）  
3. SKU成功≠バリエーション完成。未リンク時は **同一SKU修正登録**（新SKU禁止）

---

| # | 未決 | 仮運用／クローズ後 | 着手ゲートか |
|---|------|-------------------|--------------|
| 1 | **GTIN免除**（カテゴリ×ノーブランド品）＋**状態シートへ証跡記録**（Q13） | **HPCクローズ**（§11.0）。他カテゴリは未記録なら B 保留 | **必須（着手）** |
| 2 | 画像はテンプレ上 **URL列で通るか** | **HPCクローズ**: **ZIP優先**（§11.0）。URL単独は 18320 になり得る。他PTは初回再確認可 | 検収前→**HPC済** |
| 3 | M1試験用 **Product Type / バリエーションテーマ** | **HPCクローズ**: `HEALTH_PERSONAL_CARE`／`サイズ`（§11.0） | 検収前→**HPC済** |
| 4 | Browse node | **HPCクローズ**: `…検査キット (4520899051)`（§11.0） | 検収前→**HPC済** |
| 5 | 純正 `.xlsm` の版・必須列対応表（マスタ↔テンプレ↔GENERATED） | **HPC確定**: §11.5＋titlefix成功スナップショット。他PTは別表 | 検収前→**HPC済** |
| 6 | Amazon在庫0/1の実機見え方（U5） | **HPCクローズ（2026-07-23）**: 在庫0＝在庫切れ確認済。他PTは初回のみ再確認 | 検収前→**HPC済** |
| 7 | A/B同時実行 | M1は `TRACK=B` のみ（ASINあっても強制B・Q8） | — |
| 8 | PACKAGED 工程の標準ツール | **HPC方針クローズ**（§11.0）。手順書: [LV4_HPC_M1_PACKAGED_RUNBOOK.md](LV4_HPC_M1_PACKAGED_RUNBOOK.md) | 検収前→**HPC方針済** |
| 9 | リードタイム列がテンプレ必須になるか | 初版は未出力。必須エラー時のみ Property 固定日数を追加（HPCでは未発火） | エラー時 |
| 10 | 出品CKスナップショット | **不要が原則**（クローズ） | — |

### 11.5 列対応表ドラフト（M1・HEALTH PERSONAL CARE・2026-07-21）

**試験 Product Type（短期固定）**: `HEALTH PERSONAL CARE`（ドラッグストア／検査キット系で DL した「登録されていない商品を登録」純正テンプレ）。  
**正本シート**: テンプレ内 **`テンプレート`**（日本語ラベルはおおむね4行目）。  
**マスタ正本**: `▼商品マスタ(人間作業用)`。Amazon 手作業出品でよく必須だった列が揃っているため、**当面はこれらを PACKAGED／GENERATED の必須登録候補とする**（成功率優先。PT固有で不要と判明した列だけ後から任意化）。

**作業順（社長方針 2026-07-21）**:

1. **先に** マスタ列 ↔ 純正 Excel（テンプレ列）のマッピングを固める（本節）  
2. そのうえで GENERATED 列を拡張し、ローカル流し込みで手コピーを減らす  
3. カテゴリー別出品レポートは **新規Bのたたき台にしない**（§4）。成功後の辞書化は本節末の後続アイデア

**資料**: 共有スプレッドシート `amazonバルクファイル資料`（GENERATED＋HPCテンプレ取込）。カテゴリー別出品レポート `.xlsm` 3本は列参考用（既存出品データ入り・新規流し込み先ではない）。

#### 11.5.1 方針: マスタ手作業列＝必須登録候補

手作業で Amazon 登録するときにマスタへ入れている項目は、SC 側で必須になりやすい実績がある。M1 の PACKAGED では次を原則とする。

- マスタに値がある列 → テンプレ対応列へ **必ず載せる**（空のまま送らない）  
- マスタが空で SC 必須エラー → マスタ補完してから再 PACKAGED／必要なら再 GENERATED  
- GENERATED にまだ無い列は、初回はテンプレへ手埋め可。中長期は GENERATED 拡張（要実装承認）

#### 11.5.2 マスタ ↔ テンプレ ↔ GENERATED

凡例: GENERATED＝現行 `AmazonApprovalExport` 出力列。空＝未出力。

**A. 対応済み／すぐ写せる**

| マスタ（手作業・必須候補） | テンプレ列（HPC） | GENERATED | メモ |
|----------------------------|-------------------|-----------|------|
| 親SKU | 親SKU | `parentSku` | 親行・子の親参照 |
| 子SKU | SKU | `sellerSku` / `childSku` | 子＝子SKU。親行は親SKU |
| バリエーションテーマ | バリエーション テーマ | （空） | **要追加**。推奨値と完全一致 |
| バリエーション値 | テーマ応じた列（色/サイズ/個数等） | `setCount` | セット数中心。テーマと列の固定が必要 |
| 商品名amazon／オリジナルカタログ商品名 | 商品名 | `productName` | プレースホルダ不可 |
| （固定）ノーブランド品 | ブランド名 | `brand` | 完全一致 |
| 販売価格amazon | 価格（JPオファー系） | `priceAmazon` | テンプレ表記に合わせる |
| （バルク在庫0/1） | 在庫数 (JP) | `inventory` | **マスタ在庫へは書かない** |
| メーカー型番／メーカー品番 | メーカー型番 | `manufacturerPart` | 空→子SKU |
| amazon カテゴリー | 商品タイプ材料／ノード材料 | `amazonCategory` | PTは本節固定値をテンプレへ |
| （配送）送料無料パターン | 配送テンプレ名系 | `shippingTemplate` | SC登録名と一致 |
| 楽天メイン／サブ画像URL | メイン画像のURL／その他の画像のURL | `mainImageUrl` / `subImageUrls` | **https フルURL** |
| （B: GTIN空） | 商品ID／商品IDの種類 | `gtin` 空 | 免除前提。マスタJANは消さない |
| （親子） | 親子レベル | `variationRole` | parent / child |

**B. マスタ手作業で使うが GENERATED 未出力（必須候補・手埋め or 拡張待ち）**

| マスタ（手作業・必須候補） | テンプレ列（HPC・対応候補） | 優先 |
|----------------------------|------------------------------|------|
| 商品説明の箇条書き①〜⑤ | 商品のハイライト／商品の仕様 | 高 |
| 検索キーワード | 検索用キーワード | 高 |
| メーカー名 | メーカー名 | 高 |
| 特記すべき原材料／原料 | 原料 | 中（食品寄り。HPCにも列あり） |
| 商品は感熱性ですか？ | 温度管理が必要な商品ですか？ | 中 |
| 一人分の数量／一人分の数量単位 | 該当属性があれば | 低〜中 |
| ユニット数／商品のユニット数の単位 | ユニット数／単位 | 中 |
| 定価、市場価格 | 税込みの参考価格 | 低 |
| 商品タックスコード | 商品タックスコード | 中 |
| 消費税 | 税・タックスとセット | 低 |
| 原産国/地域 | 原産国系があれば | 低 |
| 危険物規制の種類 | 危険物／GHS系 | 中（該当時必須化しやすい） |
| 液体物含有 | 液体・容量系があれば | 低 |
| amazon送料 | 個別送料列より配送テンプレに寄せる | — |

**C. テンプレ側の固定／手入力（マスタに無い）**

| テンプレ列 | 初回の値 |
|------------|----------|
| 商品タイプ | `HEALTH PERSONAL CARE` |
| 出品情報アクション | 新規作成（SC表記に合わせる） |
| 推奨されるブラウズノード | 手作業成功値を1つ固定（§11-4） |
| 商品ID／商品IDの種類 | 空（B・GTIN免除） |

#### 11.5.3 現行 B2（`..._B2`）で PACKAGED 前に直すこと

1. `productName` が `JAN重複時に手入力` のまま → マスタの商品名を直す  
2. 画像がパス断片 → フルURL化（**診断用PACKAGEDは画像空可**。Firebase等は後続）  
3. バリエーションテーマが GENERATED に無い → テンプレへ手作業で1テーマ固定  

**行構造（重要・2026-07-22 処理サマリーで確定）**:

| 行 | 扱い |
|----|------|
| **5** | 属性IDマップ。**上書き禁止**（ここにデータを書くと列解釈が壊れ大量エラー） |
| **6** | Amazon サンプル行。**実SKUを置かない**（ここに親を置くと処理対象外になり、processing-summary ではサンプル `ABC123` に戻る事例あり） |
| **7〜** | **実データ**。修正登録・新規とも **7行目＝親、8行目以降＝子** |

誤って「6行目＝実データで置換」とする旧理解は **廃案**（B2 `from_master` で親が未処理・子7件のみ成功し単品バラになった）。

**マスタ直結PACKAGED（2026-07-22）**: `..._PACKAGED_from_master.xlsm`（6行目親＝失敗パターン・参考のみ）。  
**修正登録PACKAGED（同日・推奨）**: `..._PACKAGED_corrective_link.xlsm` — **同一SKU**・6行目サンプル維持・**7行目親＋8行目以降子**・アクション＝作成または置換。目的は既存子への親紐づけ（§6.1.2）。商品名←`商品名amazon`、サイズ←`バリエーション値`（HPCテーマは`サイズ`）。画像は相対パスのため空。**HPC親必須**: `電池/バッテリーが必要な商品ですか？`＝いいえ、あわせて `この商品に電池/バッテリーは含まれていますか？`＝いいえ（パッチテスト等・非電池）。

**画像（HPC確定・2026-07-23）**: マスタ相対パスでは SC が **18320（メイン画像不足）** になる。**本運用の正＝SC Upload Images ZIP（`{SKU}.MAIN.jpg`）**。R2（https直URL）はステージング／予備。取り込み成功後は Amazon 側画像が正（外部ホスト消滅で掲載済みが消える前提にしない）。Imgur・Drive直URL・GitHubは応急のみ。**画像パイプライン設計（2026-07-24）**: [LV4_R2_IMAGE_PIPELINE_POC.md](LV4_R2_IMAGE_PIPELINE_POC.md) — **Drive起点GAS案が正**（`04.amazonカタログ作成`／楽天02と同型）。**Amazonのみ**。xlsm自動埋めは提案のみ（テンプレ変化の精度懸念→当面Cursor／手作業、採否は後確定）。GAS実装は別承認。

#### 11.5.4 成功値データベース（確定方針・2026-07-22）

- **動機**: Product Typeごとに必須列・許容値が違う。エラー時に「前回成功した値」を参照して修正を速くする。  
- **正本の置き場（当面）**: ローカル `Downloads/Lv4_Amazon_PACKAGED/accepted_values_db/`（GAS外。Drive同期可）。将来はスプレッドシート／Script Properties 化を検討。  
- **何を貯めるか**:
  1. **テンプレ種子**: 純正 `.xlsm` の `推奨値` から抽出した Product Type × 列名 × 許容値（FOOD/HPC 等）  
  2. **成功スナップショット**: SC 処理が成功（かつ目視OK）した PACKAGED の **SKU×列×実値**（processing-summary の成功行も可）  
  3. **エラー事例**: エラーコード × フィールド × 失敗値 × 修正後成功値（任意）  
- **使わないこと**: 成功DBを **新規Bのたたき台の唯一ソース**にしない（§4）。エラー時の照合・PACKAGED修正の参考にする。  
- **タイミング**: HPC の `UPLOADED_OK`（親子目視後）で1件目を入れる。FOOD 試験成功時も同様に追記。  
- **Category Listings Report**: 成功後の補強ソースとして可（常時フルシード禁止は維持）。

#### 11.6 FOOD（食品）テンプレ調査（2026-07-22）

**ファイル**: `Downloads/Lv4_Amazon_PACKAGED/FOOD.xlsm`（純正・feedType=256・新規登録用）。  
**商品タイプ**: `FOOD`（推奨値シートも FOOD 固定）。HPC（ドラッグストア／HEALTH_PERSONAL_CARE）とは **別テンプレ・別列・別許容値**。

| 観点 | FOOD | HPC（参考） |
|------|------|-------------|
| ラベル列数 | 約296 | 約228 |
| バリエーションテーマ | **英語ENUM**（例 `FLAVOR/SIZE`・`COLOR/FLAVOR/PACKAGE_SIZE_NAME` 等 **54種**） | 日本語（例 `サイズ`） |
| 食品寄り列 | 栄養・炭水化物／エネルギー単位・温度の定格（常温／冷蔵／冷凍）等が多い | 電池フラグ等が目立つ |
| GTIN免除 | 推奨値に **GTIN免除** あり | 同様 |
| ブラウズ例 | 食品・飲料・お酒系ノード | ドラッグストア／検査キット系 |

**行構造（FOOD・重要・HPCと違う）**:

| 行 | 扱い |
|----|------|
| **5** | 属性IDマップ。**上書き禁止** |
| **6** | サンプル（`ABC123`）。**実SKUを置かない**（HPCと同じ教訓） |
| **7** | プリファレンスプロファイル注記行。**削除しない**（テンプレ指示） |
| **8〜** | **実データ**（このファイルでは8行目に `FOOD`／`ノーブランド品`／メーカー等がプリセット済み） |

**バルク作成は可能か**: **可能**。ただし次を守る。

1. HPC用の列マッピング／テーマ「サイズ」を **流用しない**（FOODの推奨値ENUMを使う）  
2. 進行中の HPC corrective と **SKU・ファイルを混ぜない**  
3. メイン画像は R2 直URL（または ZIP）必須想定  
4. マスタの食品テスト行を確認し、FOOD向け必須列（温度帯・栄養等が赤必須なら）を埋めてからSCへ  
5. 成功したら §11.5.4 の DB にスナップショット追記  

**いまの優先**: HPC 修正登録（画像付き corrective）の SC 結果待ちが先。FOOD は **並行調査・試験PACKAGED可**、本番UPは HPC 一段落後推奨。

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
8. ローカルで純正 `.xlsm` を PACKAGED（**実データは7行目起点・§11.5.3**）  
9. Seller Central 手動UP → 処理レポート成功 **かつ** バリエーション家族を目視確認 → **21-③**（`subBatchId` 例 `{batchId}_B2`）  
9b. 子のみ単品成功・親子未リンク → **21-⑤** `NEEDS_VARIATION_LINK` → 同一SKUで修正PACKAGED再UP → 目視後 21-③（§6.1.2。SKU変更による新規出品はしない）  
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
| 2026-07-27 | **M2キックオフ**: [LV4_M2_TRACK_A_GAP_ANALYSIS.md](LV4_M2_TRACK_A_GAP_ANALYSIS.md)／承認・HUMAN_RUN下書き。GENERATED(A)骨格あり・PACKAGED未。実装は承認後。 |
| 2026-07-24 | §10 に本線D参照（[D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)）と21-⑥を追記。 |
| 2026-07-23 | **§11.0 HPCクローズ**: §11-1〜8・10およびU5を閉じる。正本=`…titlefix`＋U5（在庫0）。画像＝ZIP優先。FOOD／他PT／M2／21-⑤は別ゲート。 |
| 2026-07-22 | **FOOD調査＋成功値DB方針**: `FOOD.xlsm` は新規用として利用可（行7=プリファレンス注記・実データ8行〜・テーマ英語ENUM）。§11.5.4を確定方針化。種子JSONは `Lv4_Amazon_PACKAGED/accepted_values_db/`。 |
| 2026-07-22 | **画像中期方針**: 原本自社＋R2ステージング＋Amazon取り込み後が正。correctiveは18320で画像必須と判明（§11.5.3追記）。 |
| 2026-07-22 | **§6.1.2 修正登録（SKU維持）**: 親子未リンク時は同一SKUで親追加＋子紐づけ。状態 `NEEDS_VARIATION_LINK`／`CORRECTIVE_UPDATE`。メニュー21-⑤。行構造は6行目サンプル維持・実データ7行目起点（§11.5.3）。 |
| 2026-07-21 | **§11.5**: HPC列対応ドラフト。マスタAmazon手作業列＝必須登録候補。出品成功後のカテゴリー別レポート辞書化は後続（§11.5.4）。 |
| 2026-07-21 | 価格: 親`販売価格amazon`空なら子へフォールバック。Logger `[Lv4Price]`。 |
| 2026-07-21 | カテゴリ解決: T列`カテゴリー`フォールバック＋見出しゆれ（改行※注記）対応。Logger `[Lv4Cat]`。 |
| 2026-07-20 | **再レビュー採用修正**: DRY_RUN専用status／冪等latest汚染除去（SKIP recordType）／subBatchId単調増加／ブランド=ノーブランド品完全一致／archive失敗は停止。 |
| 2026-07-20 | **実装レビュー修正**: レジュームindex0／価格画像SKIPPED記録／GTIN空カテゴリfail-closed・完全一致／GENERATED冪等／追記専用ステータス／clear禁止。 |
| 2026-07-20 | **実装着手（承認済）**: `AmazonApprovalExport.js` 新規／`ApprovalQueue` amazon加算／メニュー21。コード実装。 |
| 2026-07-20 | **第3回三点レビュー後**: TRACK未設定＝実行しない／Lv1 amazon抽出は同一実装チケット必須／列メモとの相互注記。 |
| 2026-07-20 | **Q11–Q14反映**: inventoryMode同期・販売中上書きは手動維持・部分失敗は親SKU単位・GTIN証跡は状態シート・配送テンプレ=送料無料パターン・リードタイム初版なし。 |
| 2026-07-30 | **D手動レ点特則・sellerSku例外**: 当面は人間の子SKUレ点を承認①相当。B/M1は子SKU、A/M2は`Amazon相乗りSKU`。同一sellerSkuは更新、未登録は新規登録。 |
| 2026-07-20 | **Q7–Q10b反映**: Lv1は親＋子個別承認・3モール同一バッチ（切り分けはmall+runId）。TRACK=BはASIN無視で強制B。B/M1の出品者SKU=子SKU／メーカー型番=メーカー品番（空→子SKU）。 |
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
| Q10b | **B/M1の正**: 出品者SKU＝**子SKU**／メーカー型番＝**メーカー品番**（空なら **子SKU** フォールバック）。A/M2は2026-07-30承認の相乗り専用例外 `Amazon相乗りSKU` |
| Q11 | **A**: 承認①ヘッダ `inventoryMode` をバルク在庫に反映（ZERO→0／ONE→1） |
| Q11b | **A**: 在庫値の話のみ。販売中SKUの無人上書き（U1）は当面手動。Lv4はスキップ維持 |
| Q12 | **A**: 部分失敗時のやり直し単位＝**親SKU一式**（バリエーションまとめて） |
| Q13 | **A**: GTIN免除証跡は **Lv4状態シート**にカテゴリ／ノーブランド品／承認日／URL or ケースIDを記録してから B 実行可 |
| Q14 | MFN必須＝配送テンプレート **`送料無料パターン`**（変更可）。**リードタイムは初版出力しない**（必須エラー時のみ Property 固定日数を後付け） |
| Q15 | **TRACK 未設定＝実行しない**（明示 `A`/`B`/`BOTH` のみ）。誤A防止 |
| Q16 | **Lv1 amazon 親＋子抽出は Lv4 実装承認と同一チケット必須**（候補0で着手禁止） |

**未回答（次ラウンド）**: 方針Q&Aは閉じた。検収前§11（Product Type固定・画像URL・PACKAGED手順書・U5）。列対応表は **§11.5 ドラフト**（マスタ必須候補方針・HPC）。手作業成功で確定が残作業。

