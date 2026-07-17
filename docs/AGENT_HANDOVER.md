# エージェント向け 引き継ぎ指示

**対象**: 本プロジェクト（gas-project）に関与する**すべてのエージェント**。  
**目的**: プロジェクト全体の要件・前提を共通認識とし、フローごとの実装要件定義と双方向のインプットを徹底する。

---

## 1. 方針

- **プロジェクト全体の要件定義・前提条件**は、**ここ（gas-project）の docs およびプロジェクト内資料**で確認する。すべてのエージェントが同じ前提で開発を進める。
- **他エージェント**は、**各フローごとに区切って**実装の要件定義を詰める（例: Phase 3 マスタ参照列整備、Amazon 出品、リサーチ自動化、見積もり効率化 等）。
- **他エージェント**も、**ここで詰めた要件定義・前提条件を必ず知ったうえで**作業する。そのため **gas-project に保存された資料を全エージェントにインプット**し、共通認識の下で開発を進める。
- **他エージェントが新たに設定した要件定義・前提条件**は、**必ずここ（gas-project）にインプットする**。該当ドキュメントの追記・更新、またはプロジェクト全体を担当するエージェントへの報告を行い、docs に反映させる。

---

## 1.5 全体理解と「いまのフェーズ」

- **プロジェクト全体の地図**は **§2 の必読一覧**と [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) が正とする。
- **現在どの開発にフォーカスするか**は [CURRENT_PHASE.md](CURRENT_PHASE.md) に集約する。
- **推奨インプット順**: ① [CURRENT_PHASE.md](CURRENT_PHASE.md) で **いまのフェーズ・次タスク**を把握 → ② **§2** の一覧で **プロジェクト全体**をインプットする。
- [CURRENT_PHASE.md](CURRENT_PHASE.md) は **フェーズ切替・セッション終了時**に更新する（誰が直したか・日付を残すとよい）。

---

## 2. 全エージェントが必ずインプットする資料（gas-project）

作業を開始する前に、**[CURRENT_PHASE.md](CURRENT_PHASE.md)** を読み、続けて以下を**すべてインプット**し、共通認識を持ってから開発・要件定義に進むこと。

### 2.1 docs 内ドキュメント（優先順）

| 順 | ファイル | 内容 |
|----|----------|------|
| 0 | [CURRENT_PHASE.md](CURRENT_PHASE.md) | **全体の位置づけと現在の開発フォーカス**（§2 に入る前に読む） |
| 0.5 | [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)・[org/AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md)・[org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)・[org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)・[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)・[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md) | **一人社長＋AI部門**。Lv0済。Lv1要件済・次は実装。モール順は楽天→Yahoo→Amazon |
| 1 | [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) | フロー間の接続・必須要件・クリティカルパス・自動化優先順位・次のアクション・§8 以降の進め方 |
| 2 | [REQUIREMENTS.md](REQUIREMENTS.md) | 6領域の要件定義・タスク・AI効率化・優先度。出庫・障がい者施設等の必須要件 |
| 3 | [MASTER_LINKAGE_TASKS.md](MASTER_LINKAGE_TASKS.md) | 既存マスタ連携の Phase 1〜2 実施結果。価格列（販売価格amazon／楽天価格設定／Yahoo!価格設定）、在庫、同期対象列、確定値の運用、今後の方針 |
| 4 | [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md) | 目的・対象モール・6領域・前提・参考資料・成果物一覧 |
| 5 | [RUNBOOK_DAY_WEEK_MONTH.md](RUNBOOK_DAY_WEEK_MONTH.md) | 日次・週次・隔週・月次の定型タスク（Runbook） |
| 6 | [AMAZON_REQUIREMENTS.md](AMAZON_REQUIREMENTS.md) | Amazon 出品の要件整理（フロー・共通必須・Amazon 固有項目。詳細は後で詰める） |
| 7 | [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) | リサーチ・見積もりの整理（出品時 vs 仕入れ時の項目・取得方法の選択肢・見積もり） |
| 8 | [ROADMAP.md](ROADMAP.md) | フェーズ案・依存関係・既存 gas-project の位置づけ |
| 9 | [PRICING_V1_REQUIREMENTS.md](PRICING_V1_REQUIREMENTS.md) | **価格・送料・再③ v1**、Script Properties による切戻し、ログ（`価格送料ロジックログ` 任意） |
| 10 | [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md) | **商品情報まわりの Gemini / OpenAI の役割**、11-③ と B Step7 の差、429 時の挙動の正 |

### 2.2 プロジェクト直下

| ファイル | 内容 |
|----------|------|
| [HANDOVER.md](../HANDOVER.md) | 楽天・Yahoo 出品の開発・実装の詳細（コード参照・シート名・列名等） |

### 2.3 参考資料（フォルダ・CSV）

| 対象 | 内容 |
|------|------|
| `参考資料：出品自動化スプシデータ` | マスタシート等をCSVで保存。Amazon・楽天・Yahoo! に必要な項目を揃えた現状の列を確認する。必要に応じて「必要sheet）.csv」「必要sheetCSV形式違い）.csv」等を参照する。 |
| **参考資料：商品リサーチ** | 現状のリサーチファイル（Amazon調査 by octas のCSV）、**バイヤーガイド2500720.pdf**、**ゆーへい問屋スクール記事購入.pdf** を格納。リサーチの外注運用・取得項目の前提。開発要件はゆーへい問屋スクール記事をベースに組み立てる。詳細は [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) §6・§7。 |
| 参考資料：メーカー・問屋仕入れ | 仕入れ・見積の前提として参照。 |

---

## 3. 役割と作業の進め方

### 3.1 プロジェクト全体の要件を扱うエージェント

- **[CURRENT_PHASE.md](CURRENT_PHASE.md)** と上記 **§2 の資料をすべてインプット**したうえで、**全体の要件定義・前提条件の確認・更新**を行う。
- 他エージェントからインプットされた**新たな要件・前提**を、適切な docs に**追記・更新**する。

### 3.2 各フローごとに実装の要件定義を詰めるエージェント

- **作業開始前**に、**[CURRENT_PHASE.md](CURRENT_PHASE.md)** と上記 **§2 の資料をすべてインプット**し、ここで詰めた要件・前提を**必ず把握**する。
- 担当フロー（例: Phase 3、Amazon 出品、リサーチ、見積もり）について**実装の要件定義**を詰める。
- **新たに設定した要件定義・前提条件**は、**必ず gas-project にインプット**する。  
  - 該当する doc（例: AMAZON_REQUIREMENTS.md、RESEARCH_AND_ESTIMATE.md、MASTER_LINKAGE_TASKS.md）に追記・更新する。  
  - または、プロジェクト全体を担当するエージェントに「〇〇を要件として追加した。docs の △△ に反映してほしい」と伝え、反映させる。

---

## 4. 双方向インプットの必須ルール

- **ここ → 他エージェント**: gas-project の docs と参考資料が**プロジェクト全体の要件・前提の正**である。他エージェントは必ずこれを読んでから作業する。
- **他エージェント → ここ**: 他エージェントが**新規に設定した**要件定義・前提条件は、**必ず gas-project のドキュメントに反映**する。反映しないまま別エージェントが動くと共通認識が崩れるため、**必須**とする。

### 4.5 ユーザーへの確認事項の出し方

- ユーザーが**質問文をコピペせずに答えられる**よう、**記号選択または短文で答えられる形式**で確認事項を出すこと。
- 例: 「1-1: [ ]契約済み [ ]これから契約 [ ]当面CSV・手動のみ」のように番号＋選択肢を並べ、ユーザーは「1-1 契約済み」のように返すだけでよい形にする。詳細は [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) **§8.5.2 回答用テンプレートの例**を参照。

---

## 5. 引き継ぎ時の一文（他エージェント起動時）

新しくエージェントを起動するときは、次の旨を伝えるとよい。

- 「gas-project の **docs/AGENT_HANDOVER.md** に従い、まず **docs/CURRENT_PHASE.md** で **いまのフェーズ**を確認し、続けて **§2 の資料をすべてインプット**したうえで共通認識で作業する。担当は［Phase 3 / Amazon 出品 / リサーチ 等］の実装要件定義。新たに決めた要件・前提は **docs に追記する**か、プロジェクト全体担当エージェントにインプットする。**実装承認後は §9（docs更新・調査ログ・復元手段）を個別指示なしで満たす。**」

---

## 6. 更新履歴

- 2026-07-17: **Lv1要件**: [org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)（承認キュー・EC書込なし）。次は実装承認。
- 2026-07-17: **Lv0最終承認**＋モール順 **楽天→Yahoo→Amazon**。レ点＝候補／スキップ＝販売中在庫>0の出品①除外／上書きは当面手動(U1)。[PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)・[LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)。
- 2026-07-17: **Lv別実装プラン叩き台**: [org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)。
- 2026-07-17: **楽天ジャンル Nav Stage1 実装**: [RAKUTEN_NAV_GENRE_STAGE1.md](RAKUTEN_NAV_GENRE_STAGE1.md)。`menuDiagnoseRakutenNavigationGenreStage1Write`（17-⑥/99-⑩）。書込は `▼診断(楽天ジャンルNav)` のみ。`RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED` 既定 false。Stage0 は [RAKUTEN_NAV_GENRE_DIAG.md](RAKUTEN_NAV_GENRE_DIAG.md)。
- 2026-07-15: **AI組織 Phase0 多数決反映**: [org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)・[org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)、憲章／マトリクス採用項、`.cursor/rules/three-review-runbook.mdc`。3者は親1＋並列サブ3が基本。**実装コードなし**。
- 2026-07-15: **AI組織 Phase0**: [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)・[org/AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md) を追加。朝承認→日中在庫0/1出品→12時前完了→夜確認。§2.1 に順0.5、CURRENT_PHASE を同期。**実装コードは含めない**（次は3者検証）。
- 2026-07-15: 憲章の **部署名を確定**（戦略企画室＝略称AI部長、販売部、サプロジ部、CS部、商品部、マーケ部、情シス部）。組織詳細は後続で組み直し可。
- 2026-07-15: 憲章 §5 に **組織図・部門役割（担当／成果物／社長提出条件）・AI社員＝ジョブ＋ルール＋docs** を追記。
- 2026-02: 初版作成。プロジェクト全体の要件はここで確認、フローごとに実装要件を詰める方針、全 agent のインプット資料一覧、双方向インプットの必須ルールを記載。
- 2026-02: **Phase 3 / Amazon 出品 / リサーチ** の実装要件定義を詰め、[MASTER_LINKAGE_TASKS.md](MASTER_LINKAGE_TASKS.md)・[AMAZON_REQUIREMENTS.md](AMAZON_REQUIREMENTS.md)・[RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) に追記。変更サマリは §7 に記載。
- 2026-02: **商品リサーチ・競合価格の前提と今後のAI方針**を [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) に追記（競合価格の書き込み先はマスタの競合価格 Amazon/楽天/Yahoo!、市場価格は参考用、既存「2.全データ一括生成」は残さず新フローに置換、AIで埋めて人間は修正のみ）。
- 2026-02: **商品リサーチ（競合価格）の実装**: コード.js にメニュー「商品リサーチ: 選択行に競合価格を入力」と `writeCompetitivePricesToMaster` を追加。手動入力でマスタの競合価格 Amazon/楽天/Yahoo! に書き込み。API連携時は同関数を呼ぶ形で拡張可能。
- 2026-02: **出品時優先・仕入れ時は後回し**: RESEARCH_AND_ESTIMATE §1.3 に「実装範囲と進め方」を追記。A（出品時分だけ実装）vs B（リサーチ全体を先に設計）の選択肢と判断用の質問を記載。**ユーザー判断待ち**。
- 2026-02: **B採用・スクレイピングで高度自動化**: ユーザー判断で**B（リサーチ全体を先に設計、実装は出品時から）**に決定。出品時で足りる情報は競合価格＋セット数で十分。**販売価格は送料加味の取得が必要**・**セット数はJANでヒットしない競合が多く画像抽出は難易度高**のため、**スクレイピングを採用**して高度な自動化に挑戦。§1.3 に判断結果と設計上の注意、§1.4 にリサーチ全体設計（出品時・仕入れ時・取得方法）、§3.1 にスクレイピング採用方針を追記。
- 2026-02: **商品リサーチ（AI提案）の実装**: メニュー「商品リサーチ: 選択行に価格・セット数提案を反映」と `proposePriceAndSetFromCompetitive` を追加。競合価格・卸値(税抜)から販売価格amazon／楽天価格設定／Yahoo!価格設定を提案してマスタに書き込む。卸値5%以上確保。セット数は呼び出し時オプションで書き込み可能。RESEARCH_AND_ESTIMATE §5 に実装済みを追記。
- 2026-02: **リサーチすり合わせ（B・C・D）確定**: 競合価格＝送料込み最安値、送料マスタ（00_設定マスタコピー・Vlookup）、出品検討手順、外注簡素化、Keepa・各モールAPI希望、閾値。RESEARCH_AND_ESTIMATE §6.4・§7・§2.5 に反映。
- 2026-02: **取得手段の確認**: Keepa API（送料込み最安値＝NEW_FBM_SHIPPING 取得可能、19eur/月、1ASIN=1トークン）、楽天・Yahoo! API（価格取得可、送料は要仕様確認）。RESEARCH_AND_ESTIMATE §8.4 に追記。
- 2026-02: **リサーチCSV取り込み・Keepa API 実装**: メニュー「リサーチCSVから競合価格Amazonをマスタに反映」、「選択行のASINでKeepaから競合価格を取得」。RESEARCH_AND_ESTIMATE §5・§8 に追記。
- 2026-02: **商品リサーチの2種類に分離**: **① 仕入れ検討用**（要件定義のみ・実装は後）と **② 出品用**（要件定義＋実装済み）をプログラム上で明確に分離。メニューは「商品リサーチ」→「① 仕入れ検討用」／「② 出品用」で選択。①は準備中案内、②に既存4機能を配置。RESEARCH_AND_ESTIMATE §1.4.0・§1.4.2 を追加・更新。
- 2026-02: **§8.5 確認事項の回答結果を反映**: Keepa契約済み・APIキー所持、楽天・Yahoo!はすぐ実装・送料込み希望、閾値フィルタ不要・全種個数セット取得、セット数はAI提案・スクレイピングメインでKeepa/APIは補完。RESEARCH_AND_ESTIMATE §8.2・§8.3・§8.5 を更新。**次回以降の確認は記号選択・短文で答えられる形式**（§8.5.2）で出す旨を §8.5 と AGENT_HANDOVER §4.5 に記載。
- 2026-02: **続きの確認（5・6・7）を反映**: 実装順は Amazon → 楽天 → Yahoo!（いずれも今から構築するスコープ、Amazon も済ではない）。楽天・Yahoo! API キーは両方所持。送料込みは 7-B（最初から送料込みを狙い、送料取得用スクレイピングも一緒に設計）。RESEARCH_AND_ESTIMATE §8.2・§8.3・§8.5.1 を更新。
- 2026-02: **確認（8・9・10・11）を反映**: Amazon（Keepa）は 8-1 C で設計からやり直し。楽天・Yahoo! の競合特定は JAN 検索＋キーワード検索（9-1）。API キーは楽天・Yahoo も Script Properties（10-1 はい）。送料スクレイピング対象は Amazon・楽天・Yahoo! の3モール（11-1）。Amazon は Keepa で確実に取れれば不要。RESEARCH_AND_ESTIMATE §8.2・§8.3・§8.5.1 を更新。
- 2026-02: **確認（12・13・14）を反映**: マスタ列は JANコード・商品名ベース（12-1・12-2）。着手は Amazon（Keepa 設計やり直し）から（13-1 はい）。送料スクレイピングは GAS ではなく Python 希望（14-1）。有料は量・金額で比較提示→ §8.6 に実行環境の選択肢と料金比較を追加。RESEARCH_AND_ESTIMATE §8.2・§8.3・§8.5.1・§8.6 を更新。
- 2026-02-22: **画像マッチング（Gemini Vision）デバッグ**: `getImageMatchScoreByGemini` が null を返す問題を調査。`fetchImageAsBase64` に詳細ログ8箇所を追加、User-Agent をフルブラウザ互換に変更、画像評価フロア補正、商品名クリーニングプロンプト改善、Keepa取得_ログに画像5軸スコア列追加。**画像取得失敗の原因特定は未完了**（clasp push → 実行 → ログ確認が必要）。HANDOVER.md §8.3.1・RESEARCH_AND_ESTIMATE.md「画像取得失敗の調査状況」に記載。
- 2026-02: **リサーチ運用（N-2・◎対象）確定**: N-2（一連実行）採用、対象は▼商品マスタ(人間作業用)の◎行、実行順は Keepa取り込み→人間が評価・セット数確認→その後のフロー。セット数提案はマスタ固定列、人間はセット数正しさと対象品合致のみ確認、セット数ごとの最低価格はGASで判断。親行・子行だけ再作成オプションを用意。FLOW_AND_PRIORITY.md §8・RESEARCH_AND_ESTIMATE.md §8.8.16 に反映。
- 2026-03: **④価格設定1（リサーチ・見積もり）の確定実装**: セット構成提案（業務用12ヶ月・賞味期限列のテンプレートコピー除外・賞味期限をAI情報から入力日起算で計算・親行水色・Amazonカテゴリー自動実行・オリジナルカタログ・購入日）、CPO（親行にJSON除く本文・Markdown表フォールバック・送料未入力時ポップアップ確認・原因追跡用ログ）、賞味期限列のテンプレートコピー除外を [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md)・[MASTER_LINKAGE_TASKS.md](MASTER_LINKAGE_TASKS.md)・[CPO_PRICING.md](CPO_PRICING.md) に反映。§8 に引き継ぎ要約を追加。
- 2026-03: **Yahoo!・楽天 競合価格取得**: メニュー「選択行のJANでYahoo!から競合価格を取得」「選択行のJANで楽天から競合価格を取得」と `menuFetchCompetitivePriceFromYahoo`・`menuFetchCompetitivePriceFromRakuten` を追加。Yahoo! ショッピング API / 楽天市場 API で JAN 検索し、取得価格を競合価格 Yahoo!・競合価格楽天に書き込む。楽天は 403 対応中でもコードは用意済み。RESEARCH_AND_ESTIMATE.md §5 に実装済みを追記。
- 2026-03: **⑤ 楽天・Yahoo! Gemini 価格提案を実装**: レ点（出品CK）が付いた行を JAN 単位で処理。行選択は不要。1 JAN あたり楽天 1 回・Yahoo 1 回 Gemini を呼び出し、計算式アウトライン（競合あり=競合-1、競合なし=前後単価流用、卸値×1.05 は使わない、単価・総額逆転禁止）をプロンプトで渡す。返答後にコードで競合-1 上書き・単価逆転防止・総額逆転防止を適用。親行に楽天価格戦略・Yahoo!価格戦略（戦略テキストのみ）、子行に楽天価格設定・Yahoo!価格設定を書き込む。③は Amazon のみのまま。
- 2026-03: **RAKUTEN_YAHOO_COMPETITIVE_PRICE_REQUIREMENTS.md** に⑤の書き込み先（楽天価格戦略・Yahoo!価格戦略・楽天価格設定・Yahoo!価格設定）を追記。
- 2026-03: **統合実行の運用標準を追加**: `B_INTEGRATED_STEP_FUNCTIONS` を単独メニュー境界の正とし、障害時の切り分けを高速化するため `runId/stepIndex/stepName/functionName/state` のログ標準、Step別の一次確認ログ、差分復元運用（小さなコミット＋変更台帳）を要件化。
- 2026-03: **§9 実装承認の既定条件**を追加。承認には自動的に「①要件docs追記 ②調査用ログ ③Git＋Script Properties での復元」が含まれる旨を明文化（個別指示不要）。
- 2026-03: **RESEARCH_AND_ESTIMATE §8.8.24（T-1/4/5/7）を `コード.js` に反映**: モール横断の参照画像を `AI情報取得data` の JAN 行のみに変更、画像 overall フロアを Keepa と共通化、セット構成のユニーク種≤2 時に `[1,2,3]` を子行直前で含有、Step7 `variation` で ◎ 行の Keepa 商品名から内容量を優先。詳細・Script Properties は RESEARCH_AND_ESTIMATE.md **§8.8.24.3**。
- 2026-03: **Step2.5 / ③⑤ 価格ロジックの確定**: 想定物流費を **3辺和+コンパクトOR+ネコポス推定**に基づき **`lastChosenRank+1` を廃止**。CPO は **モール別・全子行競合ゼロ**時に上限 **20%** 目標、`CPO_ALL_ZERO_COMP_TARGET_MARGIN` 等の Script Properties。詳細は本書 **§8 引き継ぎ要約**。
- 2026-03-19: **PRICING_V1（価格・送料・再③）**: [PRICING_V1_REQUIREMENTS.md](PRICING_V1_REQUIREMENTS.md) を正とする。寸法ランクは既定オフ（`B_LOGISTICS_USE_PHYS_RANK`）、送料はスコアタイブレーク＋利益確保不可時は**最小セットのみネコポス**、B統合は **2.6 再③** で **3b(CPO2)を置換**、B-②は **再③を末尾に追加**。再③無効は `PRICING_ROUND3_ENABLED=false`。
- 2026-03-22: **B 統合の順序**：`3a→2.5→4（楽天Yahoo）→2.6（再③）→5…`。途中再開は `B_INTEGRATED_RUN_STATE` が旧 index の可能性あり → **最初からやり直し**または state 削除を推奨（[PRICING_V1_REQUIREMENTS.md](PRICING_V1_REQUIREMENTS.md) §7）。
- 2026-03-19: **Amazon CPO V2**: [PRICING_CPO_V2_REQUIREMENTS.md](PRICING_CPO_V2_REQUIREMENTS.md)・[CPO_PROMPT_V2.md](CPO_PROMPT_V2.md)。`runCpoProposePricesV2_` は利益レール・F/P ガードを行わない。既定は V2（`CPO_ENGINE` 未設定 or `v2`）。`legacy` で旧 `runCpoProposePricesWithRound_`。メニュー「3-V2. Amazon販売価格をAI提案(V2・単体)」で常に V2。③の利益200/100・送料先調整は未実装（次フェーズ）。
- 2026-03-19: **楽天・Yahoo! CPO V2（第一段階）**: [PRICING_CPO_RY_V2_REQUIREMENTS.md](PRICING_CPO_RY_V2_REQUIREMENTS.md)・[CPO_PROMPT_V2_RY.md](CPO_PROMPT_V2_RY.md)。`runRakutenYahooCpoProposePricesV2_` は Amazon V2 と同じ `applyCpoV2PricePostProcess_`。**旧⑤**（`4. 楽天Yahoo!販売価格をAI提案`）は維持。メニュー「**4-V2. 楽天Yahoo!販売価格をAI提案(V2・単体)**」→ `menuRakutenYahooCpoProposePricesV2Standalone`。プロンプトは `getRakutenYahooCPOPromptTemplateV2` / `buildRakutenYahooCPOPromptForJANV2`。
- 2026-03-22: **CPO RY V2**: 当該モールで**競合が1件もない**ときは **`販売価格amazon` を子行ごと同期**（Gemini スキップ）。欠損行があれば Logger のうえ **Gemini フォールバック**（§3.6）。
- 2026-03-22: **Z メニュー・B 統合**: 主フロー **3／3.1／3.2／4**（Amazon V2 → 物流費 → 楽天Yahoo V2 → 再③）。旧 CPO は **「旧CPO（legacy・任意）」**サブメニュー。B Step **3.2** は `menuRakutenYahooCpoProposePricesV2ForBIntegrated_`。B-② は楽天Yahoo **V2** → Amazon **②.5後 V2** → 再③。
- 2026-03-23: **B／Z 番号**: **3／3.1／3.2／4** 表記に統一（`コード.js` `B_INTEGRATED_STEP_FUNCTIONS`・`createZSplitMenu`）。
- 2026-03-19: **Gemini / OpenAI（商品情報）**: [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md) を追加。キーワード・商品名案の本線は **OpenAI**、バリエーション単位は **Gemini→OpenAI 補完**、**11-③** と **B Step7** の処理差・**429 時**の挙動を仕様として明文化。[商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md) に同内容を追記。
- 2026-03-19: **全体＋現在フェーズの引き継ぎ**: [CURRENT_PHASE.md](CURRENT_PHASE.md) を新設。§1.5・§2 先頭（順0）・§3・§5 を更新。**いまの優先**を CURRENT_PHASE に集約し、§2 でプロジェクト全体をインプットする流れに統一。
- 2026-03-22: **§9.1** にユーザー指定の「実装時必須セット（要件docs・調査ログ・復元）」を追記。**`runProductNameProposalsForRows`**: OpenAI 失敗時もマスタ商品名があればバリエーション継続（既定）。`PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL`。詳細は [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md)。
- 2026-03-22: **◎ 複数行の perSetContent 総合判断**（`pickPerSetContentFromCircleTitles_`、g/ml 優先）。`CIRCLE_COMBINED_PER_SET_CONTENT`。AI_ROUTING・CHANGE_LEDGER 更新。

---

## 7. 実装要件定義の更新サマリ（Phase 3 / Amazon / リサーチ）

全体担当エージェントおよび他エージェントが把握しやすいよう、2026-02 に追記した要件・前提の要約を記載する。

| ドキュメント | 追記した内容 |
|--------------|----------------|
| **MASTER_LINKAGE_TASKS.md** | Phase 3 の**判断基準・実施条件・前提**。3.1 の判断基準（「楽天価格設定」「Yahoo!価格設定」参照が未実施なら要変更）。3.2 の実施条件（列名分散 or Amazon 用列追加時に Config 等を検討、楽天CSV出力は変更しない）。3.3 の実施条件（商品情報取得機能追加時に ProductInfoAcquisition.js）。前提（Phase 2 確定：価格列B採用、調査項目CSV は参考資料に配置、**同期対象列は変更しない**。**2026-05**: `syncAiDataToMaster` の突合に **`商品名ベース`** を追加。詳細は TITLE_WORKAREA_DROPDOWN_AND_MALL_NAMES_REQUIREMENTS.md）。 |
| **AMAZON_REQUIREMENTS.md** | **マスタ/参考資料CSV列との対応表**（§3.1）。参考資料「必要sheetCSV形式違い」行4の列名を基準に、商品管理番号・販売価格・在庫数・商品コード・コンディション・対象ASIN・リードタイム・画像URL・フルフィルメント等の対応と**必須/任意**を整理。§3.2 で必須/任意の一覧を明文化。§5 に**前提・制約**（既存連携は変更しない、Amazon は追加のみ、広告は対象外）を追記。 |
| **RESEARCH_AND_ESTIMATE.md** | **人間とAIのフロー・役割分担**（§1.2）：出品までの7ステップと担当・スプシのシート・列の対応。**実装範囲と進め方**（§1.3）：**B採用**。**リサーチ全体設計**（§1.4）：出品時・仕入れ時・取得方法。**取得方法**（§3.1）：スクレイピング採用。**実装済み**：競合価格入力（メニュー・writeCompetitivePricesToMaster）、**AI提案**（メニュー「選択行に価格・セット数提案を反映」・proposePriceAndSetFromCompetitive）。§5 に上記実装済みと次にやること（スクレイピング・リスク評価）を記載。**N の運用詳細**（§8.8.16）：N-2・◎対象・実行順（Keepa⇒人間修正⇒その後のフロー）・セット数提案はマスタ固定列・人間の確認範囲・セット数ごと最低価格はGAS・親行・子行だけ再作成オプション（2026-02 確定）。 |
| **RESEARCH_AND_ESTIMATE.md（2026-03）** | §5 に**セット構成提案（Gemini）**・**CPO（Gemini価格提案）**の実装済みを追記。§8.8.22 を業務用12ヶ月・12ヶ月上限フィルター・テンプレートから賞味期限列除外・賞味期限計算・親行水色・Amazonカテゴリー自動実行・オリジナルカタログ・購入日に更新。次フェーズに「CPO（Gemini）は実装済み」を注記。 |
| **MASTER_LINKAGE_TASKS.md（2026-03）** | セット構成提案の実装確定仕様で業務用12ヶ月・12ヶ月上限フィルター・テンプレートからamazon価格戦略列・賞味期限列除外・賞味期限（AI情報から入力日起算・「賞味期限（ある場合のみ）…」列優先）・オリジナルカタログ・購入日・親行水色・セット構成提案完了後のAmazonカテゴリー自動入力・セット数最大15種類を追記。 |
| **CPO_PRICING.md（2026-03）** | §6 実装メモを追加：親行にはJSONブロックを除いた本文のみ書き込む、パースは先頭```jsonブロック優先＋Markdown表フォールバック、子SKUのサイズ＆自己発/FBA・梱包箱指定が空白なら実行前ポップアップ、JANごとにマスタ/パースのセット数・未反映子行をLoggerで出力。 |

上記の追記を踏まえ、Phase 3 のコード対応・Amazon 出品実装・リサーチ実装を行う際は、各ドキュメントの該当セクションを参照すること。

---

## 8. 引き継ぎ要約（統合実行の安定化）

次の Agent が「実装済み・未解決・次アクション」を即判断できるように、統合実行を中心に整理する。

**【目的】**  
- 単独メニュー群を連結し、`B.統合実行` を「1つの実行入口」として安定動作させる。  
- 全ステップで要件漏れなく処理し、出品に必要な情報（SKU/JAN/価格/物流費/商品情報）を欠落なく出力する。  
- 30分制限下でも途中保存・再開により、最終的に全処理完了できることを必須要件とする。

**【現在の実装状況】**  
- 時間制限対策（`Date.now()`監視、途中保存、再開）が実装済み。Step5/Step7 の長時間処理はチェックポイント再開対応済み。  
- **B 統合 × Step7（2026-03 追補）**: `B_INTEGRATED_SAFE_BUDGET_MS` を **21 分**に短縮、`B_INTEGRATED_MAX_TRIGGER_RUNS` **12**。統合の Step7 は `menuProductNameAndDropdownForBIntegratedStep_` が `runStep7TitleJob(..., { deferStep7Trigger: true, bIntegratedRunId })` を呼び、**Step7 専用トリガーを二重に立てず**未完了時は `TIME_SLICE` で **統合トリガーが同じ Step7 を再開**。`menuRunBIntegrated` で **続きから**選ぶときは保存済み **`runId` を引き継ぎ**（Step7 状態の `bIntegratedRunId` と整合）。Z メニュー「7」は未完了時 **続き／最初から** を `ui.alert` で選択。  
- `LAST_PROGRESS` 等の Script Properties で進捗追跡を実装済み。  
- 物流費AI試算は「利益確保不可」時も最小物流費で仮入力し、理由を別列で管理する仕様へ変更済み。  
- **Step2.5 追伸（実装）**: 00_設定マスタ「自己発送/自己配送」行。**B=サイズ名**、**C=並び兼3辺和上限(cm)**（10〜500 を数値として解釈できるとき制約として利用）、**D=送料**。ランクは **C の value1 昇順**（同順位は価格・元行順）。**物理下限は重量を使わない**。`AI情報取得data` の AI 梱包寸法が取れているとき（`okDims`）のみ、**3辺和 = 幅+奥行+高さ×セット数** で `physMinRank` を決定；寸法未取得時は `physMin` なし（従来どおり `maxLogistics` のみ）。**`lastChosenRank+1` は廃止**（セット数が増えても同一送料帯の使い回し可）。**送料単調**（前より送料が下がらないよう前段に揃える）は維持。宅急便コンパクトは **B列名**で判定し、Script Properties **`B_LOGISTICS_COMPACT_EDGE_SUM_A` / `B_LOGISTICS_COMPACT_EDGE_SUM_B`**（既定 **50 / 58.8**）の **OR** で収まるか判定。ネコポス等は **`B_LOGISTICS_NEKOPOS_MAX_EDGE_SUM`**（既定 **60**）または B 列名からの推定。AI梱包参照は `B_LOGISTICS_USE_AI_PACK_DIMS`（既定 true）。  
- **Step3/⑤ 追伸（実装）**: CPO **③ Amazon**・**⑤ 楽天/Yahoo** は JAN×**当該モール**で **「全子行が競合ゼロ」** のとき利益率上限を **目標 20% 寄り**（`CPO_ALL_ZERO_COMP_TARGET_MARGIN`、既定 **0.20**）。**競合混在 JAN**では **競合ありの子行**は **8〜15%**（`CPO_COMP_ROW_MAX_MARGIN`、既定 **0.15**、競合−1キャップは従来どおり）、**競合なしの子行**は **8〜20%** で上限まで高め（`CPO_MIXED_NO_COMP_MAX_MARGIN`、既定 **0.20**）。**単価逆転防止・総額単調**は既存の2パスを**最優先**で維持。**手数料+販促+送料+梱包+原価** の利益式は⑤も同様。利益 200 円未満は薄赤の **目安** のみ。⑤の原価列は **`セット卸値（税込み）` 優先**。  
- **既知の限界**: ⑤の手数料は列名候補フォールバックのため、**誤った汎用「手数料率」に吸い寄せられた場合は利益表示がずれる**（マスタ列の整理を推奨）。  
- 楽天/Yahoo価格提案は価格列を文字列で潰さず、数値書き込み＋不可フラグ列で管理する仕様へ変更済み。  
- セット数上限は3ヶ月（90日）絶対上限へ変更済み。  
- 子SKU背景の強制クリアを削除し、子行の元背景を保持する挙動へ復帰済み（3回目挙動）。  
- `Bタイムオーバー5回目` では、Step7手前の保存（`nextStepIndex=6`）後に再開し、最終的に全8ステップ完了を確認済み。

**【関係ファイル】**  
- コード: `コード.js`（`menuRunBIntegrated`、`menuSetCompositionProposal`、`menuInsertMissingSetCountRows`、`runEstimateLogisticsCost`、`menuCPOProposePrices`、`menuProposePriceAndSetToSelection`、`runStep7TitleJob` ほか）  
- ログ: `Bタイムオーバー3回目.txt`、`Bタイムオーバー4回目.txt`、`Bタイムオーバー5回目.txt`、`B実行タイムオーバー履歴.txt`  
- 要件: `RESEARCH_AND_ESTIMATE.md`、`MASTER_LINKAGE_TASKS.md`、`CPO_PRICING.md`、`FLOW_AND_PRIORITY.md`

**【データ仕様（SKU/JAN/CSV/API）】**  
- 対象シートは `▼商品マスタ(人間作業用)`、`AI情報取得data`、`00_設定マスタ`。  
- 親行判定は「子SKU空行＋出品CK」の組み合わせを基準に処理。  
- 親SKU/子SKU/AK等の式は `JANコード` 参照が前提で、JAN欠落時は式結果が空になり得る。  
- `menuSetCompositionProposal` は AIデータ由来JANをマスタへ書き込むため、AI側JAN欠落はマスタ欠落へ波及する。  
- `menuInsertMissingSetCountRows` は挿入行にJAN/セット数/レ点を上書き設定する。  
- CPO価格反映はJAN単位で、パース済みセット数配列とマスタセット数配列の一致が前提。

**【未解決の課題】**  
- CPOパース結果に異常値（例: `setCount=29800`）が混入し、一部セット数へ価格反映不能になるケースが残存。  
- 列範囲警告（JAN列/出品CK列がテンプレコピー範囲外）の恒久対策が未完了。  
- 子行背景クリアを外したため、テンプレート由来で親色が子へ波及する再発可能性がある（仕様トレードオフとして許容中）。  
- 「統合実行1回で必ず全完了」はデータ量とGemini応答品質に依存し、再開運用前提が残る。

**【次に行う具体的作業】**  
- CPO入力検証を強化（許可セット数ホワイトリスト、異常値除外、欠落時フォールバック）し、価格未反映を減らす。  
- テンプレコピー範囲の設計を見直し、列追加時でもJAN/出品CK依存が壊れないようにする。  
- 統合実行の受け入れテストを固定化（開始→保存→再開→完了、JAN欠落有無、親子色表示、価格反映可否）。  
- 要件ドキュメントに「実装済み」「既知課題」「再開運用前提」を明文化し、次Agentの実装判断を統一する。

**【統合実行の境界定義（コードを正とする）】**  
- `B.統合実行` の Step 境界は `コード.js` の `B_INTEGRATED_STEP_FUNCTIONS` を正とする。  
- 現在の境界（2026-03 更新、`B_INTEGRATED_STEP_FUNCTIONS` と一致）:  
  - Step **1**: `menuSetCompositionProposal`  
  - Step **2**: `menuResearchBatchCrossMallAndPropose`  
  - Step **3**: `menuCPOProposePricesRound1`（Amazon CPO V2・②.5前。F/Pガード既定オフ。送料スナップショット保存）  
  - Step **3.1**: `runEstimateLogisticsCostStep`（想定物流費AI試算）  
  - Step **3.2**: `menuRakutenYahooCpoProposePricesV2ForBIntegrated_`（楽天Yahoo CPO V2）  
  - Step **4**: `menuRunRound3PriceAdjustIntegrated_`（再③・全モール）  
  - Step **5**: `generateListingDataComparison`  
  - Step **6**: `syncAiDataToMaster`  
  - Step **7**（統合）: `menuProductNameAndDropdownForBIntegratedStep_`（チェックポイント維持・未完了時 `TIME_SLICE` で統合トリガー再開）  
  - Step **7**（Z メニュー単体）: `menuProductNameAndDropdownForCheckedParentRows`（続き／最初からダイアログ）  
- **送料F/Pマスタ**: `00_設定マスタ` 同一シート **行120列A〜**。未シード時に表B（SS/S/M/L×関東自己配送F）・表A（帯1〜50000＋50001〜・`r_max`・`暫定区分_R1`）を自動作成。Script Properties: `CPO_FP_GUARD_ENABLE_ROUND1`（既定false）, `CPO_FP_GUARD_ENABLE_ROUND2`（既定true）, `CPO_AI_STRATEGY_WRITE`, `CPO_AI_NOTE_MAX_CHARS`, `CPO_FP_REVIEW_SHIPPING_DELTA_YEN`。  
- **単独メニュー** Z「3. Amazon…(V2・単体)」は `menuCPOProposePricesV2Standalone`。**旧CPO** サブメニューに `menuCPOProposePrices`（legacy 2ラウンド）等。**B-②** は `menuCPOProposePricesRound2Only`（②.5後・V2 時は送料参照付き V2）。

**【障害調査を速くするログ標準】**  
- 1回の統合実行を束ねる `runId` を必須にする。  
- すべての主要ログに `runId` / `stepIndex` / `stepName` / `functionName` / `state` を含める。  
- `state` は `PENDING` / `RUNNING` / `DONE` / `FAILED` / `RETRYING` を標準とする。  
- 調査順は「`runId`特定 → `afterStep` がないStepを特定 → 該当単独メニューの詳細ログを確認」で固定する。

**【Step別 一次確認ログ（運用固定）】**  
- Step1: `[セット構成提案]`（必要に応じて `[カテゴリ自動入力]` / `[抜けセット数行]`）  
- Step2: `[モール横断セット数]`（必要に応じて `[楽天セット別]` / `[Yahoo!セット別]`）  
- Step3.1: 物流費AI試算のStep専用プレフィックス（実装時に統一）  
- Step3.2: 楽天Yahoo CPO V2 のStep専用プレフィックス（実装時に統一）  
- Step4（再③）: 全モール価格調整のStep専用プレフィックス  
- Step3（Amazon CPO）: `[CPO]` + `round=` / V2 ログ（`パース結果のセット数` / `反映できませんでした` / F/Pガードログを確認）  
- Step5: AI出品取得のStep専用プレフィックス（実装時に統一）  
- Step6: 同期処理のStep専用プレフィックス（実装時に統一）  
- Step7: `[Step7]`（必要に応じて `[キーワードプルダウン]`）

**【差分復元（前に戻す）運用】**  
- MDのみでの復元運用は禁止。必ず Git の差分を正とする。  
- 方針は「1タスク1ブランチ」「小さなコミット」「コミット単位で revert 可能」を必須とする。  
- 補助として変更台帳（例: `docs/CHANGE_LEDGER.md`）を使う場合は、コード全文ではなく「対象関数・目的・戻し方（commit hash）」のみ記録する。  

---

## 9. 実装承認の既定条件（毎回・個別指示不要）

ユーザーが実装内容について **「承認します」** としたときは、**次の3点が承認に自動的に含まれる**ものとする。エージェントは、都度「追記して」「ログを」「戻せるように」と指示されなくても、実装とセットで実施する。

| # | 内容 | 実施の目安 |
|---|------|------------|
| **① 要件定義（docs）** | 確定した仕様・運用ルール・列名・タスクIDは、**該当する docs に追記または更新**する。 | 例: [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md)、[REQUIREMENTS.md](REQUIREMENTS.md)、担当領域の要件MD、および本ファイル §6 更新履歴への1行サマリ。**フェーズや優先タスクが変わったら [CURRENT_PHASE.md](CURRENT_PHASE.md) も更新する。** |
| **② 調査用ログ** | 不具合時に**原因箇所まで辿れる**ログを残す。 | 統合実行・単独メニューとも **§8「障害調査を速くするログ標準」**（`runId` / `stepIndex` / `stepName` / `functionName` / `state` 等）に沿う。**新規・変更した分岐**には、判断理由・入力サイズ・エラー要約を `Logger.log` する。 |
| **③ 復元（ロールバック）手段** | **大きな不具合時に以前の挙動へ戻せる**ようにする。 | **Git**: 小さなコミット・revert 可能な単位（§8［差分復元］）。**補助**: 新規ロジックは可能な限り **Script Properties のブール／数値**で on/off・閾値切替（既存の `getBoolScriptProperty_` 等と同パターン）。キー名と既定値は docs かコメントに残す。 |

**ユーザー側の運用**: 承認時に上記3点を毎回列挙しなくてよい。例外（例: 緊急1行パッチのみで docs 更新不要）がある場合は、そのときだけ承認文で明示する。

### 9.1 コード実装時の必須セット（ユーザー指定・§9 と一体）

実装を行う際は、**毎回**次の **3 点をセット**で実施する（**§9 の①②③と同義**だが、ユーザー文言で明示する）。

1. **要件定義書への加筆・修正** — 確定仕様を該当 docs に反映する（§9①）。
2. **不具合調査用ログ** — 新規・変更分岐で原因追跡できる `Logger.log`（§9②・§8 ログ標準）。
3. **復元手段** — 大きな不具合時に元へ戻せるよう **Git**（小さなコミット）および必要なら **Script Properties** のトグル（§9③）。**`docs/CHANGE_LEDGER.md`** に対象関数・目的・戻し方（commit hash）を記録してよい。

---
