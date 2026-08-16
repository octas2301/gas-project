# セッション引継ぎ：リサーチ・見積もり系（2026-08-14）

> **2026-08-15**: 領域1①のタスク範囲・開発予定は [org/B_PURCHASE_RESEARCH_TASK_STRUCTURE.md](org/B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)（次は W1。貼付は W4）。本書は **②出品用**。

> **2026-08-14 追記**: 領域1の本線は **①仕入れ検討**（[DOMAIN1_RESEARCH_PURCHASING.md](DOMAIN1_RESEARCH_PURCHASING.md)）。本書は **②出品用**（B統合 Step1〜4）のキャッチアップ用。CPO 異常 setCount 等は他PJ。

**対象 Agent**: リサーチ（出品用）・セット構成・モール横断・物流試算・CPO／価格 の開発を続ける人  
**元チャット**: [リサーチ見積もり引継ぎ](e232e570-3355-4089-b116-4478e668f43f)  
**日付**: 2026-08-14  
**この文書の役割**: 長い会話の結論だけを、次 Agent が実装を再開できる粒度で固定する。正本の詳細は下表の docs を読む。

---

## 0. 新 Agent が最初にやること（5分）

1. このファイルを最後まで読む。  
2. `docs/CURRENT_PHASE.md` を読む（**いまの全体フォーカスは B統合ハード死対策**。本セッションはそれと並行の「リサーチ・見積もり」スレッド）。  
3. 担当箇所だけ深掘りする（全部読まなくてよい）:

| 担当 | 必読 |
|------|------|
| セット構成・横断・画像ガード | `docs/RESEARCH_AND_ESTIMATE.md` §8.8.22〜§8.8.24 |
| Amazon CPO / 楽天Yahoo CPO | `docs/PRICING_CPO_V2_REQUIREMENTS.md`／`docs/PRICING_CPO_RY_V2_REQUIREMENTS.md`／`docs/CPO_PRICING.md`／`docs/CPO_PROMPT_V2.md` |
| 物流・再③ | `docs/PRICING_V1_REQUIREMENTS.md`／`docs/org/B_SPAPI_RESEARCH_P0_P2_REQUIREMENTS.md` |
| 統合実行の順 | `コード.js` の `B_INTEGRATED_STEP_FUNCTIONS`（**コードが正**） |
| 楽天聖域 | `HANDOVER.md` §0（`generateRakutenCSV` 等は触らない） |

**実装前**: EC一括更新・新規ファイル・複数ファイル変更はユーザー承認が必要（ユーザー規則）。

---

## 1. いまの処理順（コード正）

`B.統合実行`（`コード.js` `B_INTEGRATED_STEP_FUNCTIONS`）:

| Step | 関数 | 役割 |
|------|------|------|
| 1 | `menuSetCompositionProposal` | セット構成。テンプレコピー Q〜HJ。完了後に Amazon カテゴリ自動入力 |
| 2 | `menuResearchBatchCrossMallAndPropose` | 楽天/Yahoo 不足セット抽出。画像ガードあり |
| 3 | `menuCPOProposePricesRound1` | **Amazon CPO V2**（②.5前。送料はプロンプトに含めない） |
| 3.1 | `runEstimateLogisticsCostStep` | 想定物流費。**3D内寸フィットが既定** |
| 3.2 | `menuRakutenYahooCpoProposePricesV2ForBIntegrated_` | 楽天・Yahoo CPO V2 |
| 4 | `menuRunRound3PriceAdjustIntegrated_` | 再③（全モール）。利益200円／競合行は100円まで |
| 5 | `generateListingDataComparison` | AI出品取得 |
| 6 | `syncAiDataToMaster` | AI→マスタ同期 |
| 6.5 | `menuFillAsinNForBIntegratedStep_` | N列 ASIN 自動 |
| 6.6 | `menuFetchMakerModelForBIntegratedStep_` | メーカー品番 |
| 7 | `menuProductNameAndDropdownForBIntegratedStep_` | 商品名案・プルダウン |
| 7.6 | `menuAmazonAiAdopt76ForBIntegratedStep_` | KW・ジャンル・Yahoo売れ筋 |
| 8 | `menuSetParentRowHeightForBIntegratedStep_` | 親行高さ 60px |

**切替**: Script Properties `CPO_ENGINE` = `v2`（未設定も V2） / `legacy`。

**再開注意**: `B_INTEGRATED_RUN_STATE` の `nextStepIndex` は配列位置依存。ステップ構成を変えたあとは **途中再開せず最初から**（または Property 削除）。

---

## 2. このチャットで確定したこと（仕様）

### 2.1 セット構成（Step1）

- 業務用セットは **12ヶ月（360日）**。Gemini `daily_packets` 採用。  
- 競合セットは 12ヶ月上限でフィルタ後、最大 **15種**。  
- 3か月フィルタだけで種が 2 以下になるとき **`[1,2,3]` を必ず含める**（`ensureMinThreeVariationSetCounts123_`）。12ヶ月／業務用だけでは発動しない。  
- 親行水色。根拠文は **子SKU先頭行の `amazon価格戦略`**（セット構成提案）。  
- テンプレコピーは **Q〜HJ**。`amazon価格戦略` と賞味期限列は除外。  
- **JANコードは I 列**（G はメーカー名）。式・COUNTIFS で G を使うと壊れる。

### 2.2 モール横断・画像ガード（Step2）

- 参照画像は **`AI情報取得data` の JAN 一致行の `参考情報(画像URL)` のみ**。マスタ先頭行へフォールバックしない。  
- 空／評価不能 → **その JAN の統合セット数は採用しない**（行追加しない＝安全側）。  
- ◎相当: `IMAGE_MATCH_MIN_FOR_CIRCLE` = **70**。Keepa 後処理 `shape≥55 & color≥55` で overall フロア（`applyKeepaStyleImageMatchOverallFloor_`）。  
- ユーザー確定: JAN で楽天候補を取った**直後**に画像一致。不一致なら以降しない。一致したら価格まで書く。  
- 採用候補 0 件でも、競合価格を書けなかった行には **`要確認楽天` に理由**を残す方向（「行追加したか」ではなく「書けなかったか」）。

### 2.3 物流（Step3.1）— **後続セッションで上書き済み**

このチャット初期の「`maxLogistics` 以下で最安（ネコポス偏重）」は **廃止方向**。

| 時期 | ルール |
|------|--------|
| 本チャット前半で合意 | `total <= maxLogistics` のうち **diff 最小（= total 最大）**＝推定上限に近い配送 |
| **現行（2026-08-13 正）** | **サイズ昇順 first-fit（3D内寸）**。箱代最安ではない。寸法あり不適合は **空欄**（ネコポス埋め禁止）。利益不可時のネコポス強制も廃止 |

次 Agent は **3Dフィットを正**とし、closest-diff は補助（`B_LOGISTICS_PICK_CLOSEST_DIFF`）と読む。  
競合の `postageFlag` から配送区分を直結する設計は **未採用**（Step3.1 の入力は競合価格→`maxLogistics`）。

### 2.4 Amazon カテゴリ

- Step1 末尾の `menuFillAmazonCategoryByGemini` はある。  
- **抜けセット数行挿入の直後には無かった** → Property `B_CROSSMALL_FILL_AMAZON_CATEGORY_AFTER_INSERT_ENABLED` で挿入後に空欄親だけ埋める方針で実装済み（切戻し可）。

### 2.5 CPO V2（実装済み）

- V2 は **売値たたき台＋単価グラデーション**。利益200／F/P／マージン上限の後処理は **しない**。  
- 競合行は **round(競合)−1**。欠損は線形補間 or 1%。単調補正は非競合を動かす。  
- 旧 CPO `runCpoProposePricesWithRound_` は残置。  
- 既知バグ（残）: JSON に `setCount=29800` 等が混入し、正規セット（例 JAN `4580152230235` の 60）へ価格が乗らない。**ホワイトリスト／異常値除外は未完了**。

### 2.6 F/P 負担感ガード・AIメモ（要件確定・V2 では既定オフ）

ユーザー確定（本チャット）:

- ガードは利益条件と **別行**: \(P \ge F_{\text{円}} / r_{\max}\)。  
- 価格帯 **1000円刻み、1〜50,000円、50,001円以上は共通1行**。  
- サイズ: 〜5,000円は S/M、**5,001円〜に L**。  
- F の正は **`00_設定マスタ` 自己配送**（ネット仮値はシードのみ。実数は運用で上書き）。  
- マスタ位置: **同一シート 行120列A〜**（表B→表A）。未シード時 GAS が作成。  
- **V2 では F/P で売値を上げない**（`PRICING_CPO_V2_REQUIREMENTS.md` §3.2）。legacy の round2 で `CPO_FP_GUARD_ENABLE_ROUND2` 既定 true。

**AI判断の書き込み（確定）**:

| 内容 | 列 | いつ書く |
|------|----|----------|
| 戦略マトリクス等 | 親 `amazon価格戦略` | 従来どおり（JSON除く本文） |
| 長文ナラティブ「必勝セット数…」 | 親 `販売価格amazon` | 要望あり。**数値列汚染リスクあり**（親行のみなら比較的安全） |
| 難しい判断の短文 | 子 `楽天価格戦略`／`Yahoo!価格戦略` | **`needs_ai` のときだけ**。通常は触らない。結論1行＋根拠2〜3行、上限約400字 |

実装: `maybeWriteCpoAiNotes_` は **needsFlags A〜D のときだけ** 子の楽天・Yahoo 戦略列へ短文。`CPO_AI_STRATEGY_WRITE` 未設定=ON。  
**未実装ギャップ**: 親 `販売価格amazon` へのナラティブ移設は、V2 要件 §5 でも「未実装」。親数値列を壊すので実装時は要確認。

### 2.7 利益ルール（マスタ列ヘッダが正）

| モール | 利益額列 | 利益率列 |
|--------|----------|----------|
| Amazon | `ama利益額200円以上` | `ama利益率最低8%、15％以上` |
| 楽天 | `楽天利益額200円以上` | `楽天利益率最低8%、15％以上` |
| Yahoo! | `Yahoo利益額200円以上` | `Yahoo利益率最低8%、15％以上` |

- 原則 **税引後利益マイナス禁止・200円**。競合がある行のみ **100円まで**（再③）。  
- 15% は目安。V2 は利益レールを掛けない → **再③（Step4）が主戦場**。  
- T-8.8.24-2（利益列と CPO の警告連携）は **未完了**。

### 2.8 Step7 バリエーション

- ◎行商品名から `1個当たり内容量`（`B_STEP7_CONTENT_FROM_ASIN_CIRCLE_ENABLED`）。  
- 切戻し: `B_STEP7_WRITE_VARIATION_ENABLED`（variation フェーズ全体）。

---

## 3. 実装済み vs 次タスク

### 実装済み（再実装しない）

- CPO V2 / 楽天Yahoo CPO V2 と B 統合への組み込み  
- F/P マスタシード（行120〜）と legacy 側ガード  
- `maybeWriteCpoAiNotes_`（難しい判断のみ短文）  
- 画像ガード（AI情報取得data JAN、Keepa フロア、採用しない安全側）  
- `[1,2,3]` 下限  
- ◎内容量 → Step7  
- 3D箱フィット／ネコポス強制廃止（**後続セッション**）  
- 挿入後 Amazon カテゴリ埋め（Property）  
- N列 ASIN・メーカー品番・メニュー8 7.6（後続セッション）

### 次にやる（リサーチ・見積もり）

優先は上から。

| 優先 | 内容 | 場所 | 注意 |
|------|------|------|------|
| **1** | CPO JSON の異常 `setCount` 除外・マスタセット数ホワイトリスト・欠落フォールバック | `runCpoProposePricesV2_` パース周り | 実例 JAN `4580152230235` set=60 |
| **2** | T-8.8.24-2: 利益列ヘッダと再③／警告の整合 | Step4 中心（V2 には利益レールを戻さない） | 全角 `％` に注意 |
| **3** | 親 `販売価格amazon` への長文ナラティブ | 要設計。別メモ列の方が安全 | ユーザー要望 vs 数値汚染 |
| **4** | F/P 表A の `r_max` を実出荷で校正 | `00_設定マスタ` 行120〜 | コードの仮値のまま運用しない |
| **5** | テンプレコピー範囲外 WARNING（JAN/出品CK） | `menuSetCompositionProposal` | 列追加で壊れやすい |
| **6** | `競合価格のみ修正を反映` メニュー | 未着手 | JAN＋`A.セット商品数` で行特定 |
| **7** | 楽天・Yahoo 競合価格の API＋送料スクレイピング本線 | 後回し可 | 横断は一部実装済 |

**やらないこと**

- 楽天 CSV 生成ロジックの改修（聖域）。  
- V2 に利益レール／F/P 売値引き上げを戻す（legacy か Step4 でやる）。  
- ネコポス最安固定への回帰。  
- 仕入れ検討用リサーチ（①）の実装（要件のみ）。

---

## 4. Script Properties（この領域）

未設定はコード既定。緊急停止は `false`。

| Key | 既定 | 意味 |
|-----|------|------|
| `CPO_ENGINE` | v2 | `legacy` で旧 CPO |
| `CPO_FP_GUARD_ENABLE_ROUND1` | false | legacy round1 の F/P |
| `CPO_FP_GUARD_ENABLE_ROUND2` | true | legacy round2 の F/P |
| `CPO_AI_STRATEGY_WRITE` | true | 難しい判断メモ書込 |
| `CPO_AI_NOTE_MAX_CHARS` | 400 | メモ上限 |
| `PRICING_ROUND3_ENABLED` | true | 再③ |
| `B_LOGISTICS_USE_3D_FIT` | true | 3D箱 |
| `B_LOGISTICS_USE_PHYS_RANK` | false | 旧 3辺和ランク |
| `B_LOGISTICS_PICK_CLOSEST_DIFF` | true | 候補が複数のとき差最小 |
| `B_CROSSMALL_IMAGE_MATCH_GUARD_ENABLED` | true | 画像ガード |
| `B_CROSSMALL_IMAGE_MATCH_GUARD_EVAL_LIMIT` | 8 | 評価上限 |
| `B_CROSSMALL_FILL_AMAZON_CATEGORY_AFTER_INSERT_ENABLED` | （実装時 ON） | 挿入後カテゴリ |
| `B_SET_COMPOSITION_MIN_123_ENABLED` | true | [1,2,3] |
| `B_STEP7_CONTENT_FROM_ASIN_CIRCLE_ENABLED` | true | ◎内容量 |
| `B_STEP7_WRITE_VARIATION_ENABLED` | true | variation 書込 |

物流の旧 `B_LOGISTICS_NEKOPOS_ONLY_MIN_UNPROFIT_SET` は **未参照**。

---

## 5. 列・シート（間違えやすいもの）

| 名前 | 場所 |
|------|------|
| JANコード | マスタ **I列**（Gはメーカー名） |
| 親SKU / 子SKU | HY / HV（テンプレ1行目=親、2行目=子） |
| AK | セット識別子。`sanky-{JAN}-{UとX由来}`。重複時 `-2`。**HY の末尾置換の元** |
| 出品CK | チェックボックス。統合の対象はレ点ブロック |
| `00_設定マスタ` 行120〜 | F/P 表B（区分×F）・表A（帯×r_max） |
| `AI情報取得data` | 1行1JAN。`参考情報(画像URL)` が横断の唯一の参照画像 |

バイト制限（楽天 item_url）は **出品系**。本セッション後半で別件として出た。詳細は §7。

---

## 6. 調査の入口（コード）

| 知りたいこと | 関数・記号 |
|--------------|------------|
| 統合順 | `B_INTEGRATED_STEP_FUNCTIONS` |
| セット構成 | `menuSetCompositionProposal` / `ensureMinThreeVariationSetCounts123_` |
| 横断・画像 | `menuTestCrossMallSetCountJudge` / `getRefImageUrlFromAiDataByJan_` / `IMAGE_MATCH_*` |
| 物流 | `runEstimateLogisticsCost` / 3D は `B_LOGISTICS_USE_3D_FIT` |
| Amazon CPO V2 | `runCpoProposePricesV2_` / `getCPOPromptTemplateV2` / `applyCpoV2PricePostProcess_` |
| 楽天Yahoo CPO V2 | `runRakutenYahooCpoProposePricesV2_` 系 |
| 再③ | `runRound3PriceAdjustAllMalls_` |
| F/P シード | `CPO_FP_*` / 行120 |
| AIメモ | `maybeWriteCpoAiNotes_` |
| Step7 内容量 | `getPerSetContentFromAsinCircleTitles_` |

ログ: `Bタイムオーバー3〜5回目.txt`。プレフィックス例: `[セット構成提案]` `[モール横断セット数]` `[CPO V2]` `[pricingV1]` `[logistics]` `[round3]`。

---

## 7. 同じチャット後半の別件（出品・親SKU）※本線ではない

リサーチ・見積もりの続きではない。楽天 **32バイト（item_url）** の調査。

- 商品: JAN **4538872281013**（吉野家 缶飯）。親 `sanky-4538872281013-0s1-oya`（27B）＋接尾辞 `-48s130`（7B）= **34B 超過**。  
- HY 旧式 `=REGEXREPLACE(AK, "[^-]+$", "oya")` は **末尾1セグメントだけ**置換。AK が `…-0s1-2` だと **`-0s1` が残る**。  
- **AK列は直さない**（`-0s1` と重複連番はセット識別用）。直すのは **HY**。  
- 推奨 HY（1行目からコピー可）: `"sanky-" & I1 & "-" & suffix`（`oya` / `oya2` / `oya3`）。`-0s1` を消す。  
- 同JANで既存 `-oya` と被らせないなら `-oya2`。24+7=**31B**。  
- 式の JAN 参照は **I列**。G はメーカー名。  
- 未着手: 569行への式適用、1行目テンプレ揃え、楽天のみ D 再実行（既存出品は別 item_url になる）。

---

## 8. 新チャットに貼る文（コピペ）

```
docs/SESSION_HANDOVER_RESEARCH_ESTIMATE_20260814.md を読んで、リサーチ・見積もり系の開発を引き継いでください。
まず「いまの処理順」と「実装済み vs 次タスク」を確認し、勝手に楽天CSV・V2への利益レール復活・ネコポス最安回帰はしないでください。
次タスクの第1候補は CPO V2 の異常 setCount 除外です。実装に入る前に変更ファイル・概要・リスクを出して承認を取ってください。
```
