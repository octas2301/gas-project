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
| 0.5 | [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)・[org/AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md)・[org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)・[org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)・[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)・[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)・[org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)・[org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md](org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md)・[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)（**§11.0**）・[org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md](org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md)（**U3実機合格**）・[org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md](org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md)（**U2実機合格**）・[org/D_MENU_U2_HUMAN_RUN.md](org/D_MENU_U2_HUMAN_RUN.md)・[org/D_MENU_U2_THREE_REVIEW_MAJORITY.md](org/D_MENU_U2_THREE_REVIEW_MAJORITY.md)・[org/D_MENU_U3_HUMAN_RUN.md](org/D_MENU_U3_HUMAN_RUN.md)・[org/D_MENU_AMAZON_FACADE_THREE_REVIEW_MAJORITY.md](org/D_MENU_AMAZON_FACADE_THREE_REVIEW_MAJORITY.md)・[org/LV4_R2_IMAGE_PIPELINE_POC.md](org/LV4_R2_IMAGE_PIPELINE_POC.md)・[org/LV4_T2_HUMAN_RUN.md](org/LV4_T2_HUMAN_RUN.md)・[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md)・[org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)（**メニュー8実装済・実機未**）・[org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)・[org/LV4_AMAZON_AI_ADOPT_IMPLEMENTATION_APPROVAL.md](org/LV4_AMAZON_AI_ADOPT_IMPLEMENTATION_APPROVAL.md) | **Lv4**: U2/U3実機合格。メニュー8実装済。次はT3／ε等。[CURRENT_PHASE.md](CURRENT_PHASE.md) §0 |
| 0.6 | [org/LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](org/LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md) | **Amazonレ点本線＋Amazon相乗りSKU**。相乗り自己発＋新規同時 実機合格。七味SEASONINGはSC停止中確認／サブZIP待ち |
| 1 | [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) | フロー間の接続・必須要件・クリティカルパス・自動化優先順位・次のアクション・§8 以降の進め方 |
| 2 | [REQUIREMENTS.md](REQUIREMENTS.md) | 6領域の要件定義・タスク・AI効率化・優先度。出庫・障がい者施設等の必須要件 |
| 3 | [MASTER_LINKAGE_TASKS.md](MASTER_LINKAGE_TASKS.md) | 既存マスタ連携の Phase 1〜2 実施結果。価格列（販売価格amazon／楽天価格設定／Yahoo!価格設定）、在庫、同期対象列、確定値の運用、今後の方針 |
| 4 | [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md) | 目的・対象モール・6領域・前提・参考資料・成果物一覧 |
| 5 | [RUNBOOK_DAY_WEEK_MONTH.md](RUNBOOK_DAY_WEEK_MONTH.md) | 日次・週次・隔週・月次の定型タスク（Runbook） |
| 6 | [AMAZON_REQUIREMENTS.md](AMAZON_REQUIREMENTS.md) | Amazon 出品の列メモ（参考）。**Lv4正本**は [org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) |
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

- 2026-08-10: **Amazon本番常時ONセット**: `PUT_ENABLED`／`ALLOW_PROD`／`LV4_ENABLED`／`SC_SUMMARY` の未設定既定=true（明示falseで緊急停止）。`MASTER_QTY`／`P4B`／`U2`／`U4`は常時OFF。A1〜B2・A3合格。要 clasp push＋旧falseキー削除or true。[CURRENT_PHASE](CURRENT_PHASE.md)／[D_ENTRY](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-08-10: **サブ画像ベース色 PoC 3択**（beige/warm_white/soft_gray）: B-④ドロップダウン → compose `--base-color`（別名 `--tonmana`）。`tonmana_palette.py`／GAS `b4TonmanaPalette_` 同期。背景・カードのみ（PACKAGE_LOCK）。要 clasp push。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)。
- 2026-08-10: **B-④正本人間紐付け／写真実写ルール／サブファイル名にpattern**: Driveの`01.amazon白抜きベース`画像をJANと対応付け（`--package-truth`）。`photo_realism_rules.py`をcomposeが毎回読込。exportは`{子SKU}_{themeSlug}_subN.jpg`。要 clasp push／任意 Property `SET_MAIN_AMAZON_BASE_FOLDER_ID`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像1話完結 auto-export**: compose後に `02.楽天アップロード画像保存場所` へ `{子SKU}_subN` 自動配置（JAN→出品CK子）。人間目視はこの1フォルダのみ。再生成後も再export。要 clasp push。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像 PACKAGE_TRUTH／LOCK強化**: 正本必須（無ければJANスキップ）・正本OCR禁止・アスペクトR・文字量上限・商品名SEO・`sub_image_review_loop.py`（チェック＋要望→再生成）。B-④で正本確認。Vision自動QA・ベース色メニューは対象外。要 clasp push。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)。
- 2026-08-09: **B-④ サブ画像 Python／Cursor 指示**: サブ採用CK（なければ採用ログ）の JAN を集め、compose 本線コマンドをダイアログ表示（GAS非合成）。要 clasp push。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像コース（楽天出口）本線**: B-③参照列（themeId/phaseOrder/cluster）／compose→`{子SKU}_subN` export（`--from-compose-dir`/`--pick`）／マトリクス自動→R-Cabinet。人手＝採用CK＋短い目視。Amazonは U4 REUSE。Yahoo・Vision N・パーツ分類はバックログ。要 clasp push。[承認](org/LV4_SUB_IMAGE_RAKUTEN_COURSE_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)。
- 2026-08-09: **楽天セット数→SKU紐付け転用**: ファイル名に子SKU→楽天メイン1自動／`_subN`→サブ。セット数はマスタで子SKU決定（Vision Nはバックログ）。Property `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED`。要 clasp push。[承認](org/LV4_RAKUTEN_IMAGE_SKU_AUTOBIND_APPROVAL.md)／[要件](org/D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md)／[HUMAN_RUN](org/D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像 AI合成 v3**: 心理プロセス順スロット5×A/B／ページ全体テーマ分類／競合→phase被り→想像FO／`gpt-image-2` medium本線／パッケージ厳禁。ファイル名にテーマslug。パーツ単位分類は将来拡張（HUMAN_RUNメモ）。例 `--providers openai --openai-quality medium --jan 4538872281013`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像 hybrid段階1**: Gemini本線＋falはT03シズルのみ（文字/ロゴ禁止）。`--providers gemini,fal --route auto`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **fal.ai 画像PoC**: `fal_image.py`／`sub_image_fal_poc.py`。既定 `flux-kontext/dev`（編集）＋`flux/schnell`（最安）。compose は `--providers fal`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像 AI合成 v2**: 競合実在LPテーマのみ×各2枚／ベージュ固定／文字量抑制（成分表例外）／AI≤50%。日本語指示ボード同梱。`sub_image_ai_compose_poc.py`＋`sub_image_lp_themes.py`＋`sub_image_instruction_report.py`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像 AI合成（Gemini+OpenAI）**: Pillow精度不足のため切替。`sub_image_ai_compose_poc.py`。`gemini-3.1-flash-image` + `gpt-image-2`。競合パーツ≥約80%／新規≤約20%。例 `sub_image_ai_compose/4538872180149_20260808_192500` ほか缶飯2JAN。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **B-③右列サブ採用レ点＋採用のみログ**: `サブ採用CK`（レ点）／`サブ画像採用ログ`。再実行時GASが復元。旧「サブ画像競合候補」削除。要 clasp push。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)／[B3 docs](B3_COMPETITOR_IMAGE_AND_B_MENU_CONFIRMATIONS.md)。
- 2026-08-09: **サブ画像 JAN仕分けシート**: `サブ画像競合候補（人間確認）` にB-③をJAN紐付け。メイン/無関係を自動除外デフォルト→人が採用CK。合成は `--from-curate-sheet`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像パーツPoC（B-③連動）**: 文字読取で意図選定／配送等除外／パーツ自動提案＋安全BG合成。商品任意。例 `sub_image_parts_poc/4538872180149_*`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像PoC 10種**: 実写背景Pillow（P01–04）／競合掛け合わせ（P05–06）／写真BG誘導シーンAI（P07–10）。MAIN白抜きなし。例 `sub_image_poc/20260808_174349`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-12: **販促 docs表記揃え**: P1c＝実装済（prod検収待ち）。G8＝終着空でも apply 可。`セール前復元済`＝減衰中%へ。[要件](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)／[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-12: **販促 2層減衰**: カレンダー`減衰中%`＋B中店頭1%。restore=減衰中%（終着へ一気に戻さない）。`減衰開始日`。fetchは現在%のみ。clasp push／`--schema-only` 済。[要件§10.13/10.14](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)／[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-12: **販促 P1c-0数式修正**: 次回減衰後%の f-string バグ（`{c_pct}` 残存→#ERROR!）。`--schema-only`＋減衰中%=22。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-12: **販促メニュー分割**: `⏱ Amazonタイムセール・ポイント減衰設定`（本線1〜6／99）。広告運用から分離。減衰送信は日次＋99-⑦。要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-12: **販促 P1c実装**: `taper_send.py`（dry_run／prodフィード＋結果メール）＋列（減衰中%／次回後%／実行依頼／最終実行）＋GAS依頼メニュー。リカバリは手動。C／D後続。[要件§10.14](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)／[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-12: **販促 P1c要件**: ポイント現スナップ＋減衰進捗。細承認なし／結果メール。起動E。リカバリURL・C・Dは後続。[要件§10.14](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-12: **減衰列名**: `減衰期間`／`減衰段%`／`減衰間隔` 等。`販促売価円`削除。`--schema-only`＋clasp push。[要件§2.2/§10.12](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-12: **マスタ列並べ替え**: 目的グループ／人入力先頭／行2〜黄。`最終売価円`→`目標売価円`。GAS「マスタ列並べ替え」または `sync --schema-only`。[要件§2.2](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-12: **販促実質戻し改訂**: 人必須=`目標売価円`＋`販促ポイント%`。円・実質は表示。売価段階上げ廃止。GAS/Python提案をポイント減衰式に。要 clasp push。[要件§10.12](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-12: **販促メモ**: ポイントは過去価格を汚しにくい（記事＋二次支持・公式原文はSC確認待ち）。Dealは概ね過去30日最低／参考価格は90日系。「8週間」は公式固定でなく我慢目安。[MEMO](org/MEMO_POINTS_VS_REFERENCE_PRICE.md)。
- 2026-08-11: **販促価格戻しAPI／G10**: `price_recovery_send`（dry_run既定・our_price PATCH）。B中スキップ／期間中適用済は中止。GAS Cursor指示要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)／[要件§10.12](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促Points G5**: GAS「ポイント apply／restore（Cursor指示）」。SP-APIはCLIのまま。要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促Points 運用予行**: Smile T-1 下書き2件OK。リマインドを日程基準に修正（差分ゼロでも催促）。bのセール前%を1%退避。次は clasp push／8/27送信。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促Points G9**: listings `pointsNumber`÷price で%本線（b/s 実機1%）。`points_fetch --write` で現在%/円同期＋空のセール前のみ埋める。残り G5任意／G10。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促Points G8**: applyはセール前%必須（backup失敗も中止）。`--allow-missing-before`で解除可。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促Points G4**: `points_send`既定を施策B連動（開始接近／終了直後）。`--all-master`で解除。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促Points G2/G3**: `mail_points_remind`（apply T-1／restore 終了+1）。GASメニュー要 clasp push。次はG4。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促Points G7**: `ポイント状態`語彙固定（退避済／期間中適用済／セール前復元済／フィード*）。次はG3/G2。[要件§10.10](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促Points G1**: 1SKU本番往復検収（2%→1% feed DONE×2）。次はG7状態語彙。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: 提出xlsxを成功UL(`…114817_submitv2`)準拠に変更＝**compact＋日付直書き＋保護OFF**。全行残しはVLOOKUP壊れでSC処理失敗。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: 提出xlsxは **開始/終了の VLOOKUP を残し**、参加中＋スケジュールのみ書く（日付直書きはSC失敗の原因）。in-place・`--only-schedule` で8/9月分離。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: 提出xlsxを **in-place（行残し・参加いいえ）** に変更。8月/9月は `--only-schedule` で別ファイル。②古いDLは `使用済み/`。手削除はドロップダウン破壊の原因。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: 提出xlsxの開始/終了をデータ定義どおり `YYYY-MM-DD` 文字列で出力（Sheetsシリアル混入修正）。メニュー相当 sync→build で `DEALS_20260811_201459_all_submit.xlsx` 再作成。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: §10.13 ポイント×価格戻し連携メモ。Phase0残ギャップ（本番未検収・日程トリガー・リマインド等）をHUMAN_RUNに列挙。[要件§10.13](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)／[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: ②スケジュール列ドロップダウンの日付付き候補を全て台帳登録→提出。公式衝突時は独自セールを短縮＋メール。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: §9.7 実送信経路整備（Gmail API／GAS下書きシート＋メニュー／WebApp）。`contact@octas2301.com` へ T-14（Smile×2, `--tol 3`）送信成功。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: HUMAN_RUNをA未運用に揃え。`mail_qty_confirm` はUL済も対象＋`--tol`。T-14下書き生成済。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)／§9.7。
- 2026-08-11: **販促不変条件§0.1**: 並びI1等を sync 前検証。日付シリアルで並び崩壊を修正（`parse_ymd`＋`invariants.py`）。[要件§0.1](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促**: A未運用（行削除）／作成=UL済・再作成可／台帳は終了まで／IMAGE FORMULA保全。要件§2.1.0。[要件](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促台帳保全**: syncがBの`予定`等を消さない（Smile消失再発防止）。horizon内カタログ名付きは復元、遠方BFは自動追加しない。要確認は要件§2.1.1。[要件](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促P1b**: ②SKU行正本（カタログBF除外）／カスタム提出／UL手順はCursorダイアログ下部／再buildメニュー。③例 `DEALS_20260811_181628_all_submit.xlsx`（8/12–13×2）。要 clasp push（広告運用）。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促P1bメニュー**: ②最新確認ダイアログ→Cursor→③保存→「提出後のSCアップロード手順」。公式Saleは出たら即登録。要 clasp push（`TimeSaleP1bMenus.js`）。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促A実施フラグ**: マスタ／施策に `A実施`・ログ参照。`lane_a_send` 書込＋`--backfill-a-logs`。P1bは `--sku` で1SKU提出可。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促マスタヘッダ色分け**: 基本灰／SC青／ポイント紫／**価格戻し琥珀**／メモ黄。GAS「⏱ マスタヘッダ色付け」＋sync再適用。要 clasp push（広告運用GAS）。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促§10.12**: マスタ価格戻し列＋GAS都度提案メニュー（採用即書き込み）。TS札はBのみ。要 clasp push（広告運用GAS）。[要件§10.12](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促**: `points_fetch`（セール前API退避）／`mail_qty_confirm`（§9.7・スプシ行リンク）／`detect_new_schedules`（月次新Sale差分）。A延長ヘルパなし。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促Points列意味**: `期間中%`／`セール前%`（復元用）／`出品者現在%`。`--mode restore`。店頭合計％は退避しない。既存2SKUセール前=20（記憶）。[要件§10.10](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促Points Phase0**: `タイムセール_マスタ`にポイント列。`points_send.py`（dry_run既定／Pointsフィード）。Phase1全商品は§10.10.2メモのみ。[要件§10.10](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)／[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促**: 数量確認メールT-21/T-14（§9.7）。BFはASIN候補出現後のみ。A本番1SKU VALIDATION済→prod HTTP200。
- 2026-08-11: **販促**: 数量改定はSC画面編集が正（再バルクUL不可寄り）。`build --event bf|smile` targets-first。A dry_run HTTP200。
- 2026-08-11: **販促続き**: `ops_status`／`revise_b_qty`（§10.9.1数量改定）／`lane_a_patch`（A JSON・APIは承認後）。
- 2026-08-11: **販促提出xlsx**: データ行縦結合解除＋対象SKUのみ。§10.9.1 公式Saleは早期登録→接近時数量改定。
- 2026-08-11: **販促提出xlsx=対象SKUのみ**／§10.9早期申請割引500円（感謝祭8/5済・BF/CMは9/30まで）。[G202111590](https://sellercentral.amazon.co.jp/help/hub/reference/G202111590)。
- 2026-08-11: **販促§10.6 B固定→A合成**: 最大14日・`ab_gap_days`／`b_horizon_days`。A↔B間隔は§10.8実機検証。`plan_ab_windows.py`。
- 2026-08-11: **販促§10.7クールダウン**: おすすめ再実施14日（JP）。月(A)とSmile(B)は別系統で可。B同士は選定時に14日適用。[要件§10.7](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促§10.2数量**: `qty_logic`＋V30（SP-API Sales/CSV/マスタ）。`販売商品数_確定`へ反映。おすすめのみ提出。[要件§10](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促スケジュール選定**: ValidationDataSheetから日付付きB（Smile/BF）／Aは非重なり月枠。P1b提出xlsx再生成。[要件§2](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促画像列**: マスタ`画像URL`／施策`画像`（IMAGE関数・商品名直後）。syncで②から取得。[要件§2](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促マスタ復活**: `タイムセール_マスタ`＝対象名簿（有効のみsync）。施策は`タイムセール`。初回空なら施策から名簿復元。[要件§2](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促1枚化**: `タイムセール`＋レーン。名付き公式のみBバルク／月・カスタム=A。新着上。sync＋P1b更新。[要件§2](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促デモ埋め**: 原価U＝商品マスタU(セット卸値)ASIN突合。期間値下げ・実行を想定データで充填。[seed_demo_fill.py](../tools/amazon_deals_bulk/seed_demo_fill.py)。
- 2026-08-11: **販促シート改訂実装**: マスタ薄型SKU先頭／実行(B)／期間値下げ(A)。`sync_master_and_exec.py`＋P1bは実行シート参照。③に提出xlsx生成済。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)。
- 2026-08-11: **販促P1b着手**: Drive①②③分割。提出xlsx=Python（`tools/amazon_deals_bulk`）。GASはCursor指示。[HUMAN_RUN](org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md)／[要件](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促P0着手**: 公式終了→A自動復帰／D出品SKU重なりはAスキップ＋メール／次々回不足→2週DL。広告GASに `bootstrapTimeSaleSheetsP0`。[要件](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-11: **販促三点必須修正反映**: レーンA承認制・SKU・状態表・売切95%・次々回14日方針。マトリクスにA/B行。[要件](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)／[多数決](org/D_MENU_AMAZON_DEALS_BULK_THREE_REVIEW_MAJORITY.md)。
- 2026-08-11: **販促三点レビュー**: 3者とも条件付きYES。[多数決](org/D_MENU_AMAZON_DEALS_BULK_THREE_REVIEW_MAJORITY.md)。§14埋込済。必須修正の正本反映は社長確認後。
- 2026-08-11: **販促要件確定**: おすすめ本線・FBA/数量/価格§10・①本線②保留。P0可／P1は不足解消後。三点は推奨1回。[要件](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)。
- 2026-08-09: **サブ画像PoC軸分離**: A=`COMP_MICRO`（競合SKU微改変）／B=`OWN_REMAKE`（自社作り変え）＋実写寄り指示。例 `sub_image_poc/20260808_173136`。
- 2026-08-09: **サブ画像PoCに微改変（COMP_LIGHT）追加**: 競合フレーム維持＋軽置換。再出力例 `sub_image_poc/20260808_172326`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **サブ画像PoC（競合改変＋自社フル）**: `sub_image_poc.py`／注釈ボード。テスト出力のみ。例 `00.テスト出力/sub_image_poc/20260808_171428`。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)。
- 2026-08-09: **B-③同一画像→高画素優先**: AmazonメディアID／URL正規化で同一判定。SL²・楽天`_ex`で大きい方を残す。Keepaは同一IDを1枚化。要 clasp push。[B3 docs](B3_COMPETITOR_IMAGE_AND_B_MENU_CONFIRMATIONS.md)。
- 2026-08-09: **B-③案A**: モール横断の統合セット数「不明」／画像ガード行でも URL 付き楽天・Yahoo を画像取得対象に（`allowUnknownIntegratedSetForImages`）。要 clasp push。[B3 docs](B3_COMPETITOR_IMAGE_AND_B_MENU_CONFIRMATIONS.md)。
- 2026-08-09: **B-③ Amazon ASIN解決強化**: Keepa取得_ログへAI由来JAN追記。解決順＝ログ◎→貼付◎→マスタ親子（各段全ASIN・画像0なら次段）。Amazon0件でも楽天・Yahoo実行。[B3 docs](B3_COMPETITOR_IMAGE_AND_B_MENU_CONFIRMATIONS.md)。要 clasp push。
- 2026-08-09: **セットMAIN横長タイプ合格固定＋C③**: `landscape_layout.py`／`--aspect landscape`。N=1中央／N2縦二段／N3–4階段／N≥5上hero+グリッド。楽天N=1もC指示反映。要 clasp push。[RULES §1.2](org/SET_MAIN_LAYOUT_RULES.md)。
- 2026-08-08: **セットMAIN横長草案＋注釈テスト**: N2=下hero左辺／上unit右辺。N≥5=半端行をhero直下で全幅配分。テスト `00…/landscape_rule_annot/`（N=2/3/4/6/10/12/20）。[RULES §1.2](org/SET_MAIN_LAYOUT_RULES.md)。
- 2026-08-08: **セットMAIN横長ルール草案**: 見本154–158。N2縦積み／N3–4階段／N≥5上hero+下グリッド。四隅四辺・素材等は縦長踏襲。実装前。[RULES §1.2](org/SET_MAIN_LAYOUT_RULES.md)。
- 2026-08-08: **PUT max_items 既定10**: Property未設定時の相乗り／PUT上限を5→10。ハード上限50は維持。`APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` は省略可。[HUMAN_RUN](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)。
- 2026-08-08: **セットMAIN 01白抜き自動紐付け**: レ点親の楽天メイン画像1とVision照合→`{親SKU}_単体.png`。`bind_amazon_base_to_parents.py`／batch `--auto-bind-bases`。[HUMAN_RUN](org/D_MENU_SET_MAIN_IMAGE_PHASE_A_HUMAN_RUN.md)。
- 2026-08-08: **セットMAIN N=3細長合格確定**: Q＝ProductQuad上辺中点→円弧中点Mの後、緑法線上+35px（`unit0NudgeUpAlongNormalPx`）。缶・唐辛子検収。Cメニュー②→07→**D本線**。[RULES §1.1](org/SET_MAIN_LAYOUT_RULES.md)。
- 2026-08-08: **セットMAIN 縦長タイプ合格固定**: N=2/3/4/≥5 `locked_pass`。Cメニュー②「縦長タイプ（合格固定）」。`--aspect portrait`。unit3 BR=枠下接触。出力07→D本線。[RULES §1.1](org/SET_MAIN_LAYOUT_RULES.md)／[HUMAN_RUN](org/D_MENU_SET_MAIN_IMAGE_PHASE_A_HUMAN_RUN.md)。
- 2026-08-05: **セットMAIN 縦長 N=3合格固定**: 頂円弧・下はみ出し36・unit1下55・傾き−14/18/38。**N=4**は同ルール拡張（−14/18/38/48）。[RULES §1.1](org/SET_MAIN_LAYOUT_RULES.md)。
- 2026-08-05: **セットMAIN 縦長パターン**: `--aspect portrait`（斜めファン・傾き15°・N≤4・キャンバス正方形）。見本=縦型レイアウト基本。`portrait_layout.py`。[RULES](org/SET_MAIN_LAYOUT_RULES.md)／[HUMAN_RUN](org/D_MENU_SET_MAIN_IMAGE_PHASE_A_HUMAN_RUN.md)。GAS Cメニュー portrait=実装済。
- 2026-08-08: **P4b-d**: 誤競合ASIN対策。複数ASIN一致チェック＋browse多数決。競合無しはJAN→SHELF/楽天/Yahoo/名称の重み付き投票。Browse未確定ならPT非書込。D一括失敗はalert。[承認](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md) §2。
- 2026-08-08: **U2-ε MAIN自動**: C①で `{子SKU}_amazon.jpg` 一致分をマッチングsheet Amazon MAIN（BX）へ自動投入。目視→C②。Property `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED`（未設定=true）。[C HUMAN_RUN](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)。
- 2026-08-08: **U4途中切れ対策**: MAIN/PT URL充足スキップ・1SKU例外継続・Put/GETリトライ・約4.5分スライス→トリガー再開。C②再実行で残りだけ上がる。[HUMAN_RUN](org/D_MENU_U4_HUMAN_RUN.md) §1b。
- 2026-08-05: **C1 定価フォールバック**: `resolve_list_price` はマスタ `定価、市場価格` 優先。**空のときのみ**販売価格amazon／GENERATED `priceAmazon`。[HUMAN_RUN](org/D_MENU_C1_HUMAN_RUN.md)。缶飯 B3 再PACKAGED=`…_20260804_231959.xlsm`（list≠price 確認済）。SC ULは人間。
- 2026-08-05: **D条件付き自動再GENERATED（案B）**: 冪等0件時に PT/Browse 指紋差分があれば自動解除→再実行。Property `APPROVAL_AMAZON_LV4_PARENT_PT_BROWSE_FP`。[REMAKE](org/LV4_D_REMAKE_MENU_APPROVAL.md)／[D_ENTRY](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-08-05: **D新規ゲート Browse必須**: P4bはPT空またはBrowse空。Browse空／Node未登録はGENERATED前にハード停止。[承認](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md)。
- 2026-08-05: **Lv4 Track B 親 inventory=0固定**: MASTERでも親マスタ在庫は非読取（子のみ厳密）。親 `#DIV/0!` で送信停止しない。[D_ENTRY](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)／`AmazonApprovalExport.js`。
- 2026-08-02: **D新規ゲート＋Cursor手渡し実装**: 新規ON→PT空ならP4b→Drive棚なし停止。完了ダイアログにCursor全文コピペ。Property `AMAZON_SHELF_REGISTRY_FILE_ID`。[承認](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md)。
- 2026-08-02: **相乗り通常運用=prod直**: `Amazon相乗りSKU` 空でもprodで生成→PUT→成功後保存。as/af正規化もprod可。dry_runは上級・任意。[D_ENTRY](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-08-02: **C表示名＋兄弟PT補完**: C本線を「Ama新カタログ①／②」へ明確化。U4末尾で今回対象かつ `Amazon PT URL` 空欄の子SKUだけ、同親の最上位非空兄弟から補完。既存値・対象外は非更新。[C HUMAN_RUN](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)。
- 2026-08-01: **Amazon候補を1子SKU＝1枚**: Cの候補表示を全行タイルから子SKU行1枚ずつへ。ファイル名に子SKU／親SKUを含めばその行、残りは順に配置。戻しは `AMAZON_IMAGE_CANDIDATE_TILE_ALL_ENABLED=true`。[C HUMAN_RUN](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)。
- 2026-08-01: **相乗り ASIN空 soft skip**: 相乗りのみでも N列空は行スキップ・全体継続（有効0件のみ停止）。`コード.js`。[承認§2.1](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／[D_ENTRY](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-08-01: **D新規ゲート＋Cursor手渡し方針ロック**: 新規ON→PT空ならP4b→棚なし停止＋DL指示。完了後はCursor全文コピペ（GAS→Python起動なし）。docsのみ・実装別承認。[承認](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md)。
- 2026-08-01: **七味再送サマリ確定**: `…032459-processing-summary`＝処理2／その他エラー2／**100521のみ**（**99016解消**）。[台帳](org/LANE_B_SC_ERROR_LEDGER.md)。
- 2026-08-01: **レーンBエラー／成功台帳**: 初版＋七味99016/100521・HPC成功シード。運用=[HUMAN_RUN](org/D_MENU_LANE_B_LEDGER_HUMAN_RUN.md)／[台帳](org/LANE_B_SC_ERROR_LEDGER.md)。
- 2026-08-01: **A3検収OK**: dry `SPAPI_PUT_OFFER_CK_DRY_20260801_121701_49a49e` VALID／prod `…121813_f677a3` ACCEPTED（MASTER・自己発・`…19as13`）。トグル false 必須。[HUMAN_RUN §5](org/D_MENU_LANE_A3_HUMAN_RUN.md)。
- 2026-08-01: **A3承認包・HUMAN_RUN起草**: 自己発マスタ在庫>0（P0経路検収）。三点スキップ・原則コードなし。次＝実機。[承認](org/LV4_LANE_A3_MASTER_QTY_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_LANE_A3_HUMAN_RUN.md)。
- 2026-08-01: **B-T2方針ロック／三点スキップ**: 複合は複数・SEASONING複合は唯一でない・DL不要＝PT＋指紋・Amazon更新時は再DL・**実装は実需PT後**。[承認](org/LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md)。
- 2026-08-01: **B-T2方針ドラフト**: 新PTは実需起動・複合優先・HERB単体非強制・レジストリはPT＋指紋。コードなし。[承認](org/LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_LANE_B_BULK_T2_HUMAN_RUN.md)。
- 2026-08-01: **D在庫補強**: 新規カタログでもMASTER選択時は開始前にqty内訳確認（P0ギャップ埋め）。UIで新規GENERATED＋相乗り共通を明示。[P0](org/LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md)／[D_ENTRY](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-08-01: **B-T1 prod実装＋スモーク**: `--mode prod`。棚なし拒否／指紋不一致拒否／本番xlsm `…_oya_20260801_085442`。[承認](org/LV4_LANE_B_BULK_TEMPLATE_T1_PROD_APPROVAL.md)。
- 2026-08-01: **B-T1 prod第2段ドラフト**: dry_run同経路の本番xlsm。棚`DL_NOT_NEEDED`必須・指紋一致必須・SC非自動。[承認](org/LV4_LANE_B_BULK_TEMPLATE_T1_PROD_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_LANE_B_BULK_T1_PROD_HUMAN_RUN.md)。次＝§4ロック。
- 2026-08-01: **B-T1実装＋スモーク**: 棚引き（SEASONING=DL不要／HERB=要DL）＋項目名66/66＋dry_run FILL `B_T1_FILL_SEASONING_20260801_084627`。[承認](org/LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_LANE_B_BULK_T1_HUMAN_RUN.md)。
- 2026-08-01: **B-T1ドラフト追記**: テンプレ棚引き（PT＋指紋）。あり＝SC DL不要／なし＝09へ要求。定常フローを承認§5・HUMAN_RUNへ。[承認](org/LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md)。
- 2026-08-02: **MAP正本=sheet／派生=MD／実行=JSON**: PACKAGED前に `sync_map_sheet_to_column_json.py`。[承認](org/LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md)／[LEDGER](org/MAP_SC_ERROR_LEDGER.md)。
- 2026-08-02: **SHELF Browse 網羅カタログ**: 06純正`データを閲覧する`全行→`shelf_browse_catalog.json`／MAP右SHELF。P4b/Dは browseNodeId で採用PT。缶飯=GROCERY。MEATエイリアスはフォールバック。[承認](org/LV4_SHELF_BROWSE_CATALOG_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md)。失敗後再GENERATEDはD内レ点（案A・**実装済＋clasp push済**）[承認](org/LV4_D_REMAKE_MENU_APPROVAL.md)。
- 2026-08-02: **セットMAIN文字色5種＋Cメニュー**: えんじ/青/黒/緑/茶（`--text-color`）。見本`94/01_〜05_`。C「セットMAIN量産（Cursor指示）」出力先=楽天ULフォルダ。[HUMAN_RUN](org/D_MENU_SET_MAIN_IMAGE_PHASE_A_HUMAN_RUN.md)／[LAYOUT](org/SET_MAIN_LAYOUT_RULES.md)。
- 2026-08-02: **セットMAIN AI PoC**: Nano Banana（通常 Flash Image 最新・PRO禁止）。見本03/04。[承認](org/LV4_SET_MAIN_IMAGE_PHASE_AI_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_SET_MAIN_IMAGE_PHASE_AI_HUMAN_RUN.md)。Pillow Phase Aは凍結。
- 2026-08-02: **セットMAIN画像 Phase A 実装**: `tools/set_main_image`（Pillow）。Drive07＋C。子SKU命名。マスタ一括。Amazonコラージュ→楽天／Yahoo数字レイヤ。[承認](org/LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_SET_MAIN_IMAGE_PHASE_A_HUMAN_RUN.md)。
- 2026-08-02: **GROCERY本線（缶飯）**: `FOOD_FISH_GROCERY.xlsm` 指紋登録。棚 B-T1-2（GROCERY/FOOD/FISH）。Catalog `MEAT`→`GROCERY` エイリアス（P4b＋Dゲート）。C1マップ `food_fish_grocery_column_map.json`。次＝`clasp push`→D→PACKAGED。[T2](org/D_MENU_LANE_B_BULK_T2_HUMAN_RUN.md)。
- 2026-08-02: **C1 FOOD改定**: 許可PTに`FOOD`追加。マスタ PT/browse 必須（空・唐辛子既定埋込禁止）。ハイライト=楽天キャッチ→Yahoo→箇条書き①（タイトル≤75）。メーカー型番=HJ優先。次＝マスタにFOOD+browse投入→PACKAGED再作成。[列マップ](org/D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md)。
- 2026-08-01: **B-T1方針ドラフト**: 複合＋SEASONING。項目名マッピング。マスタのみ書込。HERB単体非本線。[承認](org/LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md)。
- 2026-08-01: **P4b-b合格**: マスタHERBのまま C1 SEASONING dry_run（`C1_CK_daba393f8055_B2_20260801_073624`）→ xlsm PT=`SEASONING`／browse=defaults。ローカルPythonはAgent実行と明記。[HUMAN_RUN](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)。
- 2026-08-01: **P4b-a合格→P4b-bへ**: runId`…161210_546de6`／rival_asin／HERB。[HUMAN_RUN](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)。
- 2026-08-01: **P4bマスタ競合版実装**: `AmazonCategoryPt.js` を競合店ASIN→URL→自ASIN→Definitions＋WARNに置換。多数決削除。[承認](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)。
- 2026-08-01: **P4b §2マスタ競合改定**: 正本＝`競合店ASINコード`→`競合AmazonページURL`→自`ASINコード`→Definitions。多数決廃止。正しさチェックはWARNのみ。[承認](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)。
- 2026-08-01: **P4b §2多数決改定＋実装**: 評価◎・売れ筋上位最大5で Catalog 多数決。同票は売上1位。SEASONING優先廃止。◎0は自ASIN／Definitions。（§2は後にマスタ競合へ再改定）[承認](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)。
- 2026-08-01: **Zメニュー連番サブ化**: 直下を「1〜10／11〜17／18〜21／C・E互換／99」に折りたたみ。21は本線・SP-API・SC/P4b。項目番号・関数名は不変。`clasp push`後にシート再読込で反映。
- 2026-08-01: **P4b実装**（三点スキップ）: `AmazonCategoryPt.js`・21-⑱・C1マスタPT/browse優先（SEASONING/HPC）。Keepa新規取得なし。次＝`clasp push`→実機。[承認](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)。
- 2026-08-01: **P4b §2方針ロック**: Catalog→既存Keepa browse→Definitions→人間。提案PT＝SEASONING/HERB/HPC、C1本線＝SEASONING/HPC。Keepa新規取得なし。次＝実装承認→コード。[承認](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)。
- 2026-08-01: **P4b承認包起草**: SP-API PT／browse → マスタ空セル書込・C1接続。手数料`amazon カテゴリー`非対象。コードなし。次＝§2方針ロック→実装承認。[承認](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)。
- 2026-08-01: **デュアル Phase2 検収OK**: dry 自己発 `SPAPI_PUT_OFFER_CK_DRY_20260801_141229_f372b8`／FBA `…141243_40d85e`（VALID）・prod `…141420_8fcbc2`／`…141432_cc7c72`（ACCEPTED・行494／`B0D9VK8YPS`）。次＝レーンB台帳運用・トグル戻し。[承認§5](org/LV4_DUAL_OFFER_PHASE2_APPROVAL.md)／[HUMAN_RUN §3](org/D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md)。
- 2026-08-01: **デュアル Phase2 コード実装**: Dチェック複数選択・自己発→FBA順次PUT・部分成功可。三点スキップ。次＝`clasp push`→両系統 dry／prod。[承認](org/LV4_DUAL_OFFER_PHASE2_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md)。
- 2026-08-01: **A2検収OK**: キャンセル`cancelled_by_user`（12:07）／E0 `menuAmazonCourseE0Precheck` checked=1。次＝A3。[HUMAN_RUN §4](org/D_MENU_LANE_A2_HUMAN_RUN.md)。
- 2026-08-01: **A2承認包・HUMAN_RUN起草**: 確認キャンセル／トグル戻し／逃げ道疎通。三点スキップ・原則コードなし。次＝実機。[承認](org/LV4_LANE_A2_OPS_HARDENING_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_LANE_A2_HUMAN_RUN.md)。
- 2026-08-01: **デュアル Phase1検収OK**: prod 自己発 `SPAPI_PUT_OFFER_CK_PROD_20260801_115554_4ed30e`／FBA `…115648_eb2511` ACCEPTED（dry `…8fa79e`／`…d6ed67`）。次＝A2／A3。[承認§6.1](org/LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)。
- 2026-08-01: **デュアル Phase1 dry_run実機OK**: 自己発 `SPAPI_PUT_OFFER_CK_DRY_20260801_114254_8fa79e`（NF／`…as13`）／FBA `…114820_d6ed67`（`_FBA`／`…af13`）。prod未提示。次＝prod任意 or A2／A3。[承認§6.1](org/LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)。
- 2026-08-01: **デュアル Phase1実装**: 自己発=`Amazon相乗りSKU`／FBA=`Amazon相乗りSKU_FBA`（読取・保存は系統のみ・他列不変）。三点スキップ。次＝マスタ列追加→clasp push→両系統 dry_run 検収。[承認](org/LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)。
- 2026-08-01: **レーンA1検収OK**: FBA dry_run `SPAPI_PUT_OFFER_CK_DRY_20260801_111613_41ce9e` VALID／prod `SPAPI_PUT_OFFER_CK_PROD_20260801_111845_6bd20f` ACCEPTED（`sanky-B01N5A6ESU-19af13`・compliance attrs）。次＝A2／デュアル Phase1／A3。[HUMAN_RUN §0c](org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md)。
- 2026-08-01: **FBA属性実装**: FBA時 `batteries_required`/`batteries_included`=false・`supplier_declared_dg_hz_regulation`=not_applicable。失敗時おすすめをダイアログ表示。Property `APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_ATTRS`（既定true）。次＝clasp push→七味FBA再検収。[承認](org/LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md)。
- 2026-08-01: **FBA属性方針承認**: §3折衷（dry_run必須→失敗時おすすめ→再dry_run→prod）。三点スキップ。次＝実装承認。[承認](org/LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)。
- 2026-08-01: **FBA属性承認§3を折衷に改定**: 既定dry_run→失敗時issues＋おすすめ提示→人間確認→再dry_run必須→prod。初手prod／コンプライアンスのdry_runスキップは不採用。[属性承認](org/LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)。
- 2026-08-01: **A1対比確定＋FBA属性承認ドラフト**: 七味`B01N5A6ESU`で自己発dry_run VALID（`…105918_199878`）／FBA 90220（`…105527_f64afc`）。次＝FBA時に電池・危険物属性追加。[HUMAN_RUN §0b](org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md)／[属性承認](org/LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)。
- 2026-08-01: **デュアルオファー Phase1方針ロック**: 新列`Amazon相乗りSKU_FBA`、既存NF＝自己発専用、Dは1回1系統のまま列分離、初回はDのみ。実装は別承認。[承認](org/LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)。A1暫定はNF空＋七味（`B01N5A6ESU`）。
- 2026-08-01: **レーンA1方針承認**: 相乗りFBA dry_run→**1SKU prod**まで。三点スキップ。原則コードなし。次＝[HUMAN_RUN](org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md)実機。[承認](org/LV4_LANE_A1_FBA_OFFER_APPROVAL.md)。
- 2026-08-01: **レーンA1承認包ドラフト**: 相乗りFBA dry_run→少件prod（原則コードなし・在庫>0はA3）。[承認](org/LV4_LANE_A1_FBA_OFFER_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md)。次＝社長方針承認。
- 2026-08-01: **方針ロック（A/B/C）**: **A**既存カタログ＝SP-API本開発／**B**新規＝xlsm＋SC人手で取り貯め／**C**新規JSON＝ゲート後（1〜2PT・同等dry_run等）。P2のxlsm①②追従・今の新規JSON全面はしない。[承認§7.4](org/LV4_P2_DC_123_INVESTIGATION_APPROVAL.md)／[ROADMAP §1.1](org/AMAZON_DEV_ROADMAP_P0_P4.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-08-01: **P2調査完了**: PACKAGED xlsmのAPI UP／SC processing-summary自動DLは不可→人手①②＋21-⑮温存。③中身パースは任意・低優先。JSON Listings寄せはレーンC。[承認§7](org/LV4_P2_DC_123_INVESTIGATION_APPROVAL.md)。
- 2026-08-01: **P2調査承認包起草**: Dc①UP／②時間DL／③取込強化。④⑤・非公式自動化・本番自動UPコードは非同梱。過渡21-⑮温存。[承認](org/LV4_P2_DC_123_INVESTIGATION_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_P2_DC_HUMAN_RUN.md)。
- 2026-08-01: **P1検収OK＋吉野家承認済**: Cモーダル実機＋他認識。Amazonメールで吉野家出品承認（反映〜24h・停止時は保留在庫解消）。次＝七味ライブ確認＋P2調査承認。[HUMAN_RUN](org/D_MENU_P1_HUMAN_RUN.md)。
- 2026-08-01: **P1実装**: C/D/U4/④の案内を File最少化（02手置き＝復旧例外のみ）。要 clasp push。[HUMAN_RUN](org/D_MENU_P1_HUMAN_RUN.md)／[承認](org/LV4_P1_FILE_MIN_APPROVAL.md)。
- 2026-08-01: **P1方針承認**: §2既定（07＋sheetドラッグ＋手ZIP残／02手File廃止）。三点スキップ。次＝実装承認（案内・ゲート文言中心）。[承認](org/LV4_P1_FILE_MIN_APPROVAL.md)。
- 2026-08-01: **P1承認包起草**: 07以降 File最少（`02`手操作廃止・退避/U4はGAS。MAINドラッグ自動化・U5非同梱）。三点今回スキップ案。[承認](org/LV4_P1_FILE_MIN_APPROVAL.md)／[HUMAN_RUN下書き](org/D_MENU_P1_HUMAN_RUN.md)。
- 2026-08-01: **P4a読取PoC 実機合格**: `keywords=七味`→HERB/SEASONING（getはSEASONING優先）、Catalog分類可、xlsm自動DL不可。次＝P1承認。[HUMAN_RUN](org/LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md)。
- 2026-08-01: **P4a読取PoC実装**: `tools/spapi_smoke` に `--poc-category`（search／getDefinitions／Catalog分類）。GAS非改変・マスタ非書込。次＝人間実機。[HUMAN_RUN](org/LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md)／[承認](org/LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md)。
- 2026-08-01: **CをD型選択モーダル化**: トップCは1項目。Amazon準備／ドラッグ後／楽天／前提チェックを選択。C-0〜C-2はZ互換に残置。[HUMAN_RUN](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)。
- 2026-08-01: **C本線 Property一時ON**: C-1/C-2（E互換含む）で U2/U4 を実行中だけ true→finally復元。常設は `AMAZON_IMAGE_CANDIDATE_FOLDER_ID` のみ。C-2開始前確認＋ScriptLock。[HUMAN_RUN](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)。
- 2026-08-01: **Cコース実装**: トップC.画像コース（C-0/C-1/C-2）。①〜④はZ互換。E-1/E-2→C寄せ。[承認](org/LV4_C_COURSE_CONSOLIDATION_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)。
- 2026-08-01: **Cコース統合 方針ロック**: [LV4_C_COURSE_CONSOLIDATION_APPROVAL](org/LV4_C_COURSE_CONSOLIDATION_APPROVAL.md)。トップ1本・U4はAmazon新規のC-2のみ・E-1/E-2一本化・三点スキップ・P1/ドラッグ自動化は後。
- 2026-08-01: **P0実装**: D在庫0\|マスタ（ALLOW_MASTER_QTY）・prod既定・E→D補助/Z互換。[P0承認](org/LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md)／[D_ENTRY HUMAN_RUN](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。在庫0＋dry_run／MASTER_QTYゲート確認OK。
- 2026-08-01: **P0三点多数決**: [MAJORITY](org/LV4_D_P0_THREE_REVIEW_MAJORITY.md)。社長(B)マスタqty＝承認②＋親正本改定。1SKU必須免除。
- 2026-08-01: **Amazonロードマップ確定**: [AMAZON_DEV_ROADMAP_P0_P4](org/AMAZON_DEV_ROADMAP_P0_P4.md)。次＝P0 HUMAN_RUN＋並列P4a調査。[P0承認](org/LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md)／[P4a](org/LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md)。
- 2026-07-31: **docs同期**: CURRENT_PHASE§0をYahoo再認証済・SC監視自動ON/OFF・許可待ちに更新。Property作業後チェックリストを正本化。
- 2026-07-31: **SCサマリ監視 自動ON/OFF**: GENERATED成功で待ちリスト＋トリガー自動設置。待ちが全て UPLOADED_OK／FAILED で削除。上限72h。21-⑯⑰は手動残置。[HUMAN_RUN §1e](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-07-31: **Yahoo再認証半自動**: メニューで認可code貼付→B14＋**C14取得日**。残り≤7日はB17へメール（任意日次トリガー）。`YAHOO_OAUTH_REDIRECT_URI` 必須。[HUMAN_RUN](org/YAHOO_OAUTH_REAUTH_HUMAN_RUN.md)。出品本体は非改変。
- 2026-07-31: **docs整理（外出先）**: CURRENT_PHASE§0を七味到達点に更新。SCサマリ手順§1e短文化。在庫停止中メモ（親`B0HC9S8PRN`／子`B0HC9RRCBP`／相乗り`…19as13`）。コミットは人間指示待ち。
- 2026-07-31: **U4楽天サブ→R2**: PT参照空かつ非AMAZON_ONLYなら `楽天サブ画像1〜8` を取得→R2→`Amazon PT URL`。D自動U4も検知。七味でPT成功=3（`U4_20260731_124440_329c1a`）。手ZIPは帰宅後。[U4 HUMAN_RUN](org/D_MENU_U4_HUMAN_RUN.md)。
- 2026-07-31: **SEASONING SC必須知見**: item_form=粉末／item_weight_unit=グラム／検索KWは1枠（99016対策）。PACKAGED再送 `…_20260731_032459.xlsm`。初回サマリは99016＋100521。[列マップ](org/D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md)。
- 2026-07-31: **新規＋相乗りprod 同時 実機合格**: 同一レ点行から新規 `CK_daba393f8055_B2_GENERATED.csv`（rows=2）＋相乗り `sanky-B01N5A6ESU-19as13` `ACCEPTED issues=0`（qty=0）。**新規が冪等除外で0件だとファサードが例外で止まり相乗りへ到達しない**点に注意（作り直しは 21-④ で `UPLOAD_FAILED` 後に再実行）。フルコースは楽天OK・**Yahooはrefresh token期限切れで失敗**（要再認証）。[HUMAN_RUN §1b](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-07-31: **SC処理サマリ自動記録（21-⑮⑯⑰）**: SC UP後のサマリを監視フォルダへ置くだけで `UPLOADED_OK` を自動記録（メニュー操作ゼロ）。**判定はファイル名のみ**（C1出力が `{subBatchId}_PACKAGED_…` のため subBatchId を機械抽出）。`_NG`／`UPLOAD_FAILED` を含む名前は `UPLOAD_FAILED`。`GENERATED` 行が無い subBatchId は記録せずファイルも残す（fail-closed）。既定 false・要 `APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_ID`。push済。[HUMAN_RUN §1e](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-07-31: **21-⑭ 証跡おすすめ**: 同カテゴリのマスタ`ASINコード`から `過去成功ASIN＋dp URL` を提示。OKでそのまま記録、キャンセルで手入力。SC申請URL／ケースIDは自動不可。
- 2026-07-31: **GTIN免除証跡の記録メニュー（21-⑭）**: 新規カタログはカテゴリごとに `EXEMPTION` 行が必要（HPCの証跡は食品に効かない）。レ点新規の親からカテゴリ実値を検出し、証跡入力＋確認ダイアログを経て `▼Lv4実行ログ(Amazon)` へ追記する。有効な証跡があるカテゴリは追記しない。全カテゴリ `*` は `APPROVAL_AMAZON_LV4_EXEMPTION_ALL_CATEGORIES=true` のときのみ（追加警告あり、原則カテゴリ別）。[HUMAN_RUN §1d](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-07-30: **Dレ点新規の在庫ゲート緩和**: `source=child_ck` は在庫>0でも `SKIPPED_IN_STOCK` にせず GENERATED を作る（同じレ点行から作るのは既存単品とは別のノーブランドセットカタログ）。マスタ在庫は読取のみ・非改変、バルク在庫は `inventoryMode` 準拠。承認①経路は従来どおり除外。戻し＝`APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK=false`。[HUMAN_RUN §1c](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-07-30: **Amazon画像Cの旧sheet誤紐付け対策**: U2有効時は楽天未整理画像0件でも現在レ点からCを再生成。C-Amazon②は現在レ点とsheet親子SKUの完全一致を確認し、不一致なら候補配置前に安全停止。[HUMAN_RUN](org/D_MENU_U2_HUMAN_RUN.md)。
- 2026-07-30: **C1-1c SEASONING**: 新規 `FOOD_HERB_SEASONING_FISH_VEGETABLE.xlsm` を調査。吉野家七味はPT=`SEASONING`、ブラウズ=`唐辛子 (2430212051)`。321列PT別マップ／指紋、行7注記保持・行8開始を実装。機械DRY_RUN＋HPC回帰合格。[列マップ](org/D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md)。
- 2026-07-30: **同一レ点行の新規＋相乗り同時出品**: Dで両方選択時、N列ASIN有無で排他分割せず同じ子行を両方へ出す。新規=`子SKU`、相乗り=`Amazon相乗りSKU`（N列ASIN必須）。[HUMAN_RUN](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-07-30: **D相乗り自己発 実機合格**: dry_run OK=1（`SPAPI_PUT_OFFER_CK_DRY_20260730_081735_99e`）／prod OK=1（`SPAPI_PUT_OFFER_CK_PROD_20260730_082325_9f2287`）。SKU=`sanky-B084RJSH7W-48as12`。次＝コード開発なしでFBA・新規・同時・フルの実機確認。[HUMAN_RUN](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。
- 2026-07-30: **D相乗り運用修正**: 相乗り自己発/FBAをDで選択（X列は新規SKU式用のまま）。中央が既にASINの子SKUはそのまま`Amazon相乗りSKU`に使う。マスタ在庫>0でも送信quantityは常に0（非0出品はしない）。
- 2026-07-30: **Amazonレ点本線＋Amazon相乗りSKU実装**: `コード.js` Dを新規／相乗り複数選択へ。`AmazonApprovalExport.js` は人間レ点の新規Bを承認①互換明細でGENERATED。`AmazonSpapiPut.js` は相乗り行をN列`ASINコード`のみでPUTし、dry_run VALID/issues=0後にNF列`Amazon相乗りSKU`保存、prodは保存値を再利用。O列／競合URLは出品先に使わない。承認①／21-⑫⑬は温存。
- 2026-07-30: **Amazonレ点本線＋Amazon相乗りSKU承認包**: 3者独立レビューは全員条件付き、2/3以上を社長承認して憲章・マトリクス・Lv4正本・商品マスタ要件へ反映。Dは新規＋相乗り同時可、**フル＋相乗りprodも開始前確認付きで許可**。新旧sellerSkuとも同一SKUは更新、未登録は新規登録。[承認](org/LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)。
- 2026-07-29: **A（D入口版）実装**: `コード.js` — D ラジオに `amazon_offer_dry`／`amazon_offer_prod` を追加。`runBatchExportAmazonOfferFacade_` が 21-⑫⑬ を呼ぶ（`02` MAIN ゲートなし）。E-0 に本線＝D の1行。要 clasp push→[HUMAN_RUN](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。[承認](org/LV4_SPAPI_D_ENTRY_APPROVAL.md)。
- 2026-07-29: **A（D入口版）承認**: 社長「承認」。**実装可**（範囲は承認 §2・§3。`コード.js` の `showBatchExportModal` ラジオ拡張＋`runBatchExportAmazonFacade` 分岐で既存 21-⑫⑬ を呼ぶ薄いファサード。フル＋相乗り prod 禁止・相乗りでは Drive `02` MAIN ゲート不適用・E/Z 残す）。[承認](org/LV4_SPAPI_D_ENTRY_APPROVAL.md)。
- 2026-07-29: **A（D入口版）承認起草**: 本線 D を Amazon 起点に。ラジオ＝新規カタログ／既存相乗り dry_run／既存相乗り prod／フル→新規。**D から prod 可**（ALLOW_PROD＋確認必須）。フル＋既存相乗り prod は禁止。E/Z はテスト期間残可。**コードなし・社長承認待ち**。三者レビュー不要。[承認](org/LV4_SPAPI_D_ENTRY_APPROVAL.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-29: **SP-API v1.4 第2段 実機合格（API）**: 21-⑫ VALID（`…231605_1b24e2`）／21-⑬ **ACCEPTED**（`…232041_3d83f3`）。batch `A1_20260727_224939_b7a053`・`lifec-4560151300832-48s11`／`B07YND44VN`・在庫0・親行スキップ1。prod は確認ダイアログで **OK を押さないと `cancelled_by_user` で PUT されない**（1回目は取消）。SC最終更新の目視は反映待ち。[HUMAN_RUN §3.3](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)。
- 2026-07-29: **SP-API v1.4 第2段 実装**: `AmazonSpapiPut.js` に承認①済収集＋**21-⑫ dry_run／21-⑬ prod**。Property は第1段と共用。[第2段承認](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)。
- 2026-07-29: **SP-API v1.4 第2段 承認**: 社長「承認する」。**実装可**（範囲は承認 §2・§3＝承認①済抽出は 21-⑨ 流用・子SKU必須・LISTING_OFFER_ONLY・Property は第1段と共用）。[第2段承認](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)。
- 2026-07-29: **SP-API v1.4 第2段 承認起草**: 承認①済 Amazon（21-⑨相当）を GAS 直 PUT する案。三者レビュー不要。[第2段承認](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-29: **SP-API v1.4 実機合格（API）**: 21-⑩ VALID／21-⑪ ACCEPTED（`lifec-4560151300924-48s11`／`B00A0J0D30`）。`https:\` ENDPOINT正規化・hostヘッダ削除。SC最終更新の目視は反映待ち。**clasp push 済（`-f`）**。[HUMAN_RUN](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-29: **SP-API v1.4 実装**: `AmazonSpapiPut.js`＋21-⑩ dry_run／21-⑪ prod（子SKUレ点）。Property `APPROVAL_AMAZON_SPAPI_PUT_*`／LWA。[承認](org/LV4_SPAPI_GAS_PUT_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)。
- 2026-07-29: **SP-API v1.4 承認**: 社長「承認する」。起草→実装可。[承認](org/LV4_SPAPI_GAS_PUT_APPROVAL.md)。
- 2026-07-29: **SP-API v1.4 承認起草**: GAS UrlFetch Listings（第1段＝子レ点・dry_run／prod 分離トグル）。三者レビュー不要。[承認](org/LV4_SPAPI_GAS_PUT_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-29: **SP-API v1.2b 実機合格**: 21-⑨→Drive APPR→`--fetch-drive` dry_run／prod。`lifec-4560151300832-48s11`／`B07YND44VN` ACCEPTED。親`…oya`スキップ確認。親レ点出さないは社長完了扱い。[APPROVED承認](org/LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md)／[BRIDGE HUMAN_RUN](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-29: **SP-API v1.2c／v1.3 実機合格**: 子レ点→Drive→`--fetch-drive` dry_run／prod。SKU`lifec-4560151300924-48s11`／`B00A0J0D30` ACCEPTED。[CHECKBOX承認](org/LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md)／[BRIDGE HUMAN_RUN](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-28: **SP-API v1.2c**: 21-⑧＝マスタ**子SKU＋出品CK**のみ→Drive CSV。選択行完全廃止。親レ点のみ出さない。[CHECKBOX承認](org/LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md)／[BRIDGE HUMAN_RUN](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)。
- 2026-07-28: **SP-API v1.3／v1.2b／v1.2a**: Drive自動取得（`--fetch-drive`）・21-⑨承認①済一括CSV・UTF-8/cp932。GAS直呼びなし。[BRIDGE HUMAN_RUN](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)／[DRIVE承認](org/LV4_SPAPI_DRIVE_FETCH_APPROVAL.md)／[APPROVED承認](org/LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md)。要 clasp push。
- 2026-07-28: **SP-API スプシ橋渡し v1.2 実機合格**: 21-⑧→Drive CSV→ローカル dry_run/prod→SC確認。安眠`…0924-ride01`／アルコール`…0139-ride01`・1000/0。Property・allow_prod false戻し済（当時は選択行。のち v1.2c でレ点化）。[HUMAN_RUN](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)／[承認](org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-28: **SP-API スプシ橋渡し v1.2**: `AmazonSpapiExport.js`＋メニュー21-⑧（当初選択行→Drive CSV）。GASからSP-API直呼びなし。[HUMAN_RUN](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)／[承認](org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md)。
- 2026-07-28: **SP-API Listings v1.1 複数行バッチ**: `items_csv`／`max_items`／行単位。試験=安眠`B00A0J0D30`・アルコール`B0091G3AHY`・ride01・1000/0。GAS直結なし。prod実機合格。[HUMAN_RUN](org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)／[BATCH承認](org/LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md)。
- 2026-07-28: **SP-API Listings書込 v1**: `tools/spapi_listings_write`（putListingsItem LISTING_OFFER_ONLY／dry_run=VALIDATION_PREVIEW／prodは allow_prod）。発汗1SKU prod合格。[HUMAN_RUN](org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)／[承認](org/LV4_SPAPI_LISTINGS_WRITE_IMPLEMENTATION_APPROVAL.md)。
- 2026-07-28: **SP-API読取スモーク合格**: LWA＋Catalog／`B07YND44VN`＝発汗チェッカー1500。レポート `SPAPI_SMOKE_20260728_115740`。書込は[承認包](org/LV4_SPAPI_LISTINGS_WRITE_IMPLEMENTATION_APPROVAL.md)。[HUMAN_RUN](org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md)。
- 2026-07-28: **SP-API読取スモーク**: `tools/spapi_smoke`（LWA＋Catalog GET）。秘密は config.local.json。GAS書込は別承認。[HUMAN_RUN](org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md)。
- 2026-07-28: **M2公式Loader自動埋め**: `tools/m2_offer_packaged/m2_listing_loader_fill.py`＋`listing_loader_map.json`。dry_run/prod。SC自動UPなし。[HUMAN_RUN](org/D_MENU_M2_HUMAN_RUN.md)。
- 2026-07-28: **M2実機合格＋SP-API認証完了**: 発汗 `0832-48s11`／`B07YND44VN`。公式 ListingLoader 埋め→SC成功1／E-5=`A1_20260727_224939_b7a053_A2`。簡易CSV単独UPは不可と確定。SP-API本番アプリ自己承認＋認証3点手元保存（Git禁止）。次＝Loader自動埋め（要承認）。[HUMAN_RUN](org/D_MENU_M2_HUMAN_RUN.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-27: **M2 v1実装**: 案L `tools/m2_offer_packaged`／競合ASIN解決／`BRAND_GATE_MODE=manual_ok`。試験=発汗 `0832-48s11`/`B07YND44VN`。要 clasp push。[HUMAN_RUN](org/D_MENU_M2_HUMAN_RUN.md)／[承認](org/LV4_M2_IMPLEMENTATION_APPROVAL.md)。
- 2026-07-27: **M2（TRACK=A）キックオフ docs**: ギャップ洗い／HUMAN_RUN下書き／承認パッケージ。コード未着手。次＝社長§4回答＋実装承認。[GAP](org/LV4_M2_TRACK_A_GAP_ANALYSIS.md)／[承認](org/LV4_M2_IMPLEMENTATION_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_M2_HUMAN_RUN.md)。
- 2026-07-27: **C1-clean**: `generated_csv`=`{subBatchId}_GENERATED.csv`＋`sub_batch_id`／`--sub-batch`。`relax` 固定廃止。dry_run出力=`A1_…_B2_PACKAGED_…`。[HUMAN_RUN](org/D_MENU_C1_HUMAN_RUN.md)。
- 2026-07-27: **安眠 SC合格＋E-5完了**: 処理2／成功2／エラー0。Downloads `…0924-oya-processing-summary.xlsm`。E-5=`A1_20260726_225610_4f0558_B2`。[HUMAN_RUN](org/D_MENU_C1_HUMAN_RUN.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-27: **安眠 SC送信済（外出先確認待ち）**: Batch `182816020660`／E-5=`A1_20260726_225610_4f0558_B2`。[REMOTE_CHECKLIST](org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)／[CURRENT_PHASE §0](CURRENT_PHASE.md)。
- 2026-07-27: **C1入力案A（fetch）**: `c1_fetch_inputs.py` で GENERATED＋マスタCSVを自動取得（OAuth読取のみ）。[要件](org/D_MENU_C1_FETCH_INPUTS_REQUIREMENTS.md)／[HUMAN_RUN](org/D_MENU_C1_HUMAN_RUN.md) §1b。
- 2026-07-27: **親 Amazon MAIN URL**: U4/E-2 後に子→親（空欄）自動コピー。Lv4 Build は子フォールバック。要 clasp push。[E HUMAN_RUN](org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)。
- 2026-07-26: **メニューE Amazon出品コース（一時）**: E-0〜E-5 薄いファサード（C/18/21/U4呼出）。要 clasp push。[HUMAN_RUN](org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)。
- 2026-07-27: **E-3完了ダイアログ**: 承認Web URL付き（`approvalQueueBuildWebUrl_`）。18-②必須ではなくなる。要 clasp push。[E HUMAN_RUN](org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)。
- 2026-07-27: **メニュー8 v1.10**: 共有4列（メイン/CTR/特徴/用途）の**類似横断dedupe**（末尾品・用／5文字以上の包含）。要 clasp push。[要件](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)。
- 2026-07-27: **syncAiDataToMaster キャッチ1行化**: 楽天／Yahooキャッチの作業エリアへは `getFirstDropdownOptionLine_` で先頭1案のみ（改行全文転記禁止・赤セル防止）。要 clasp push。[商品マスタ作業エリア要件](商品マスタ_人間作業エリアとマスタエリア_要件定義.md)。
- 2026-07-27: **メニュー8 v1.9**: KW区切り＝半角スペースのみ（`,`/`、`禁止→正規化＋要確認）／メインKW＝1語（大中カテゴリSEO）／最終名上限超過時は式の**後ろ列・後ろワード**から厳格trim。要 clasp push。[要件](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)／[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **Yahooブランド正本=getShopBrandList**: コード検証＋メーカー名でAPI取得。要 clasp push。[STAGE](YAHOO_CATEGORY_BRAND_STAGE.md) §2.9。
- 2026-07-26: **Yahooカテゴリ正本=SHP getShopCategoryList**: 書込前検証。無効ID却下／名前でSCコードへ解決。要 clasp push（`Yahoo.js`含む）。[STAGE](YAHOO_CATEGORY_BRAND_STAGE.md) §2.8。
- 2026-07-26: **Yahoo競合ゼロフォールバック**: AI推奨列→DriveマスタCSV。要確認付き。要 clasp push。[STAGE](YAHOO_CATEGORY_BRAND_STAGE.md) §2.7。
- 2026-07-26: **Yahooカテゴリ／ブランド Stage 実装**: メニュー8へ。売れ筋(review_count近似)＋自社最安。トグル `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED`。要 clasp push。[STAGE](YAHOO_CATEGORY_BRAND_STAGE.md)／[HUMAN_RUN](org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md)。
- 2026-07-26: **Yahooカテゴリ選定改定**: 売れ筋（`-review_count`近似）＋自社`Yahoo!価格設定`が最安になりやすいカテゴリ優先。単純多数決廃止。[STAGE](YAHOO_CATEGORY_BRAND_STAGE.md)。
- 2026-07-26: **Yahooカテゴリ／ブランド Stage 要件起草**: 都度API（itemSearch投票）・階層名`:`連結・ショップ＝プロダクト写し。実装は承認後。[STAGE](YAHOO_CATEGORY_BRAND_STAGE.md)／[承認](org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md)。
- 2026-07-26: **楽天ジャンル Stage3**: 都度API（Ichiba投票→Nav）をメニュー8へ。AI推奨不使用。トグル `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED`。要 clasp push。[STAGE3](RAKUTEN_NAV_GENRE_STAGE3.md)／[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **メニュー8 v1.6**: バリエーションテーマ選定をメニュー8へ（マスタは日本語。HPC=`サイズ`／食品=`パッケージサイズ`。Amazon裏マップで FOOD→`PACKAGE_SIZE_NAME`）。要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **メニュー8 v1.5**: 下限=上限−5・半角0.5換算・容量語1つ・特徴/用途は商品名酷似なら空可。要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **メニュー8 v1.4**: 楽天・Yahooも横断dedupe。最終名上限 Amz75/楽天120/Yahoo75（検索KWから弱語削除）。要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **メニュー8 v1.3**: 最終商品名amazon式の左優先・完全一致横断dedupe（部分一致なし）。要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **メニュー8 v1.2**: 再生成・カテゴリ自動・バリ再推論・商品名案を除外。KW9選択＋dedupeのみ。要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **メニュー8緊急修正**: `syncAiDataToMaster` を除外（レ点外・列全体破壊の主因）。要 clasp push。[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **メニュー8（Amazon AI一括採用）v1実装**: 承認済。Z→**7.5**（番号8は楽天CSV）。空欄のみ・M-A・要確認列ごと・商品名dedupeメモのみ。次＝clasp push→[HUMAN_RUN](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)。
- 2026-07-26: **メニュー8（Amazon AI一括採用）要件＋承認パッケージ**: M-A／空欄のみ／generate_then_adopt／要確認列ごと（濃い赤白字＋メモ）／商品名dedupeはメモのみ／価格・ブランド後回し。次＝**実装承認**。[要件](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)／[承認](org/LV4_AMAZON_AI_ADOPT_IMPLEMENTATION_APPROVAL.md)。
- 2026-07-26: **C1 SC合格＋21-③代替記録**: relax（手置きGENERATEDのためログに行なし→21-③スキップ）。SC＋Drive05で完了代替。次＝本線は21-①正式ID／New Agent可。[HUMAN_RUN](org/D_MENU_C1_HUMAN_RUN.md)。
- 2026-07-26: **C1 SC合格**: relax（親`lifec-4560151300405-oya`／子`…-16s184`）SKU2成功・エラー0。Drive05 processing-summary。次＝**21-③**。[HUMAN_RUN](org/D_MENU_C1_HUMAN_RUN.md)。
- 2026-07-26: **定価列方針確定**: 調査文言は `▼マスタ(市場価格調査)` のみ。`定価、市場価格` は数値／計算式専用（転記廃止）。`コード.js`＋要件docs。要 clasp push。
- 2026-07-26: **C1-1b実装**: `master_csv`併読・必須列埋め・タックスはマスタ。[HUMAN_RUN](org/D_MENU_C1_HUMAN_RUN.md)。次＝未送信SKUでSC。
- 2026-07-26: **C1列マップ下書き**: [D_MENU_C1_MASTER_HPC_COLUMN_MAP.md](org/D_MENU_C1_MASTER_HPC_COLUMN_MAP.md)。生成OK・SCは必須列不足。
- 2026-07-26: **C1実装承認**: `tools/c1_hpc_packaged`（Python/openpyxl）／[HUMAN_RUN](org/D_MENU_C1_HUMAN_RUN.md)。GAS非改変。
- 2026-07-26: **C1三点反映**: [MAJORITY](org/D_MENU_C1_THREE_REVIEW_MAJORITY.md)／要件更新／[承認包 §6](org/LV4_C1_IMPLEMENTATION_APPROVAL.md)。URL空スキップ・親一式除外・指紋v1本番停止。
- 2026-07-26: **C1要件起草**: [D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](org/D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)／[LV4_C1_IMPLEMENTATION_APPROVAL.md](org/LV4_C1_IMPLEMENTATION_APPROVAL.md)。本線＝ローカル／Cursor。
- 2026-07-26: **U4 実機合格**: runId `U4_20260726_090920_1366af`・マスタに Amazon MAIN URL。Dは冪等想定内。[HUMAN_RUN §0](org/D_MENU_U4_HUMAN_RUN.md)。
- 2026-07-26: **U4 v1 実装**: 21-⑦ Drive02→R2→マスタURL＋GENERATED優先。[HUMAN_RUN](org/D_MENU_U4_HUMAN_RUN.md)。次=clasp push。
- 2026-07-26: **U4要件起草**: [D_MENU_U4_R2_URL_EMBED_REQUIREMENTS.md](org/D_MENU_U4_R2_URL_EMBED_REQUIREMENTS.md)／[LV4_U4_IMPLEMENTATION_APPROVAL.md](org/LV4_U4_IMPLEMENTATION_APPROVAL.md)。実装は社長承認後。
- 2026-07-26: **T2再検証合格**: `lifec-…-80s10`／URL単独・ZIPなし・18320なし・店頭OK。T3必須ではない。[LV4_T2_HUMAN_RUN.md](org/LV4_T2_HUMAN_RUN.md)。
- 2026-07-25: **U2 実機合格**: ②③④・`02` MAIN1件。Dは `idempotentBlocked=1`（`LV4_20260725_094425_914290`）想定内。[HUMAN_RUN §0](org/D_MENU_U2_HUMAN_RUN.md)。Propertyはfalseへ。
- 2026-07-25: **U2 v1＋C子レ点**: `AmazonImageMatrixExport.js`＋C-Amazon①〜④。Cは子レ点優先（親のみは全子）。候補=`07` Property済。
- 2026-07-25: **U2三点＋社長回答反映**: ONLY=sheet、マスタ永続／復元、候補=Amazon用フォルダ。[org/D_MENU_U2_THREE_REVIEW_MAJORITY.md](org/D_MENU_U2_THREE_REVIEW_MAJORITY.md)。次＝コミット指示→実装承認。
- 2026-07-25: **U2方針確定**: 本線案α（MAIN＝マッチングsheet・子SKU行）、`02`＝出口、ε＝バックログ。[org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md](org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md)。次＝実装別承認。
- 2026-07-25: **U3実機合格**（runId `LV4_20260725_072635_948892`）＋**U2要件起草** [org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md](org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md)。T3／full_amazonはゲート維持。
- 2026-07-25: **D×Amazon U3 v1**: Dに `amazon`/`full_amazon`。裏は `menuApprovalAmazonLv4Run`＋Daダイアログ。トリガー非搭載。[org/D_MENU_U3_HUMAN_RUN.md](org/D_MENU_U3_HUMAN_RUN.md)。次は clasp push。
- 2026-07-25: **D×Amazon U0クローズ**: 3者多数決反映＋社長回答（手ZIP正／T3実装待ち／将来API必須）。[org/D_MENU_AMAZON_FACADE_THREE_REVIEW_MAJORITY.md](org/D_MENU_AMAZON_FACADE_THREE_REVIEW_MAJORITY.md)。次はU3別承認。
- 2026-07-24: **D×Amazon要件U0**: [org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md](org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md)（本線A〜D・Da・C案α）。コードなし。
- 2026-07-24: **T2 PoC成功**: runId `R2T2_20260724_221107_7f9cf7`・SKU `lifec-4560151300139-oya`・公開URLで画像表示。HMAC Byte[]修正済。トグルfalseへ。**T3は保留**。[CURRENT_PHASE.md](CURRENT_PHASE.md) §0。
- 2026-07-24: **T2 clasp push済**: 8 files（`AmazonDriveImageExport.js`含む）。次は Property＋21-⑥→URL200→トグルoff。[CURRENT_PHASE.md](CURRENT_PHASE.md) §0／[org/LV4_T2_HUMAN_RUN.md](org/LV4_T2_HUMAN_RUN.md)。
- 2026-07-24: **帰宅引き継ぎ**: CURRENT_PHASE §0 を自宅PC続行用に全面更新。次は T2 `clasp push`＋21-⑥。[CURRENT_PHASE.md](CURRENT_PHASE.md) §0／[org/LV4_T2_HUMAN_RUN.md](org/LV4_T2_HUMAN_RUN.md)。
- 2026-07-24: **T2実装**: `AmazonDriveImageExport.js`＋メニュー21-⑥（Drive MAIN1枚→R2）。手順 [org/LV4_T2_HUMAN_RUN.md](org/LV4_T2_HUMAN_RUN.md)。人間が clasp push。T3は別承認。
- 2026-07-24: **Nav Stage1実機PASS**: 2/2・runId `navS1_20260724_203812_a61c6b17`。Propertyはfalseへ戻す。次はAI枠確認→T2承認任意。
- 2026-07-24: **並行 Nav→AI→T2**: Agent静的確認（Navコード／429継続既定）。人間実行メモ [RAKUTEN_NAV_STAGE1_HUMAN_RUN.md](RAKUTEN_NAV_STAGE1_HUMAN_RUN.md)。T2承認書 [org/LV4_T2_IMPLEMENTATION_APPROVAL.md](org/LV4_T2_IMPLEMENTATION_APPROVAL.md)（未実装）。
- 2026-07-24: **FOOD v5成功**: 親＋子6バリエーション反映・在庫切れ。成功DBを Drive `05` に記録。CURRENT_PHASE §0。
- 2026-07-24: **並行docs A→D**: HPC手順1枚 [org/LV4_HPC_M1_PACKAGED_RUNBOOK.md](org/LV4_HPC_M1_PACKAGED_RUNBOOK.md)／画像T2チケット（POC §7.1）／AI §5.1／Nav §9.1。コードなし。FOODは100521待機。
- 2026-07-24: **FOOD v5 corrective準備**: SC確認で親＋`30s175`のみ登録。Drive `03/…_v5_corrective.xlsm`＋`04/…_MISSING5_MAIN_for_SC.zip`。人間がxlsm→ZIPを1回。CURRENT_PHASE §0。
- 2026-07-24: **外出リモート引き継ぎ**: CURRENT_PHASE §0更新。画像はDrive起点＋サブ二モード（REUSE_RAKUTEN/AMAZON_ONLY）設計済・コード未実装。FOOD再UP禁止。GAS/R2実装は別承認。
- 2026-07-24: **Drive起点GAS案を正に**: [org/LV4_R2_IMAGE_PIPELINE_POC.md](org/LV4_R2_IMAGE_PIPELINE_POC.md) 更新。`04.amazonカタログ作成` フォルダ構成。xlsm自動埋めは提案のみ（テンプレ変化で精度懸念→当面Cursor）。GAS実装は別承認。
- 2026-07-24: **R2×Amazon画像PoC設計**: [org/LV4_R2_IMAGE_PIPELINE_POC.md](org/LV4_R2_IMAGE_PIPELINE_POC.md)（**Amazonのみ**／楽天・Yahoo対象外）。実装は別承認。
- 2026-07-23: **帰宅引き継ぎ**: HPC `…_suburl_test.xlsm` UP結果待ち（子`…-40s10`にPT01〜05 URL）。FOOD v4ログ再解析（100521/100730/13013・再UP禁止・v5はZIP＋親価格）。[CURRENT_PHASE.md](CURRENT_PHASE.md) §0。
- 2026-07-23: **§11.0 HPCクローズ**: §11-1〜8・10＋U5。正本 titlefix＋在庫0。画像ZIP優先。FOOD／他PT／M2／21-⑤は別ゲート。[LV4要件](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) §11.0／[CURRENT_PHASE.md](CURRENT_PHASE.md)。
- 2026-07-23: **外出・リモート引き継ぎ**: HPC titlefix SC 8/8 → 21-③ `UPLOADED_OK` 済・`ENABLED=false`。次は FOOD v4（100521/100730/13013待機）。成功値DB追記済（Downloads）。[CURRENT_PHASE.md](CURRENT_PHASE.md) §0。
- 2026-07-22: **帰宅引き継ぎ**: R2公開済（`pub-d974…r2.dev`）・8枚`{SKU}.MAIN.jpg`はHTTP200。次はcorrectiveへURL埋め→SC再UP→親子目視→21-③（`…_B2`）。[CURRENT_PHASE.md](CURRENT_PHASE.md) §0。
- 2026-07-22: **外出記録**: corrective_link SCは18320/8007。中期は原本自社＋R2＋Amazon取り込み後が正（R2準備はその後完了）。
- 2026-07-22: **Lv4 §6.1.2**: 修正登録（SKU維持）・親子未リンク時は親追加＋既存子紐づけ。メニュー21-⑤要件化。実データは7行目起点。
- 2026-07-21: **Lv4 §11.5**: HEALTH PERSONAL CARE 列対応ドラフト。マスタAmazon手作業列＝必須登録候補（マスタ↔Excel先）。成功後カテゴリー別レポート辞書化は後続。
- 2026-07-21: **自宅IDE引き継ぎ**: Lv4 DRY_RUN＋本GENERATED成功（runId `LV4_20260721_212702_217672`／subBatchId `A1_20260721_083100_06b90a_B2`）。カテゴリ・子価格フォールバックは clasp push済。次はPACKAGED→SC→21-③。[CURRENT_PHASE.md](CURRENT_PHASE.md) **§0**。未コミットあり・commitは指示まで禁止。
- 2026-07-21: **セッション終了（深夜）**: Lv4 clasp push成功。当時の次はドライラン（その後完了）。
- 2026-07-20: **Lv4 再レビュー採用修正**: DRY_RUN専用／冪等latest汚染除去／subBatchId単調増加／ブランド=ノーブランド品完全一致。次は修正後3者レビュー→clasp push。
- 2026-07-20: **Lv4 実装レビュー修正**: レジューム二重スキップ解消・価格/画像はSKIPPED記録・GTIN空カテゴリ不合格・GENERATED冪等・ステータス追記専用。次は clasp push→ドライラン検収。
- 2026-07-20: **Lv4 実装**: `AmazonApprovalExport.js` 新規、`ApprovalQueue.js` に amazon 親+子加算、`コード.js` メニュー21。マスタ在庫/JAN非書込・TRACK未設定は非実行・GTINは状態シート EXEMPTION。次は clasp push→検収。
- 2026-07-20: **Lv4 Q15/Q16＋列メモ**: TRACK未設定＝実行しない／amazon親＋子抽出は同一実装チケット必須／`AMAZON_REQUIREMENTS.md` に JAN・リードタイム・FBA の Lv4上書き注記。次は実装承認。
- 2026-07-20: **Lv4 Q11–Q14反映**: inventoryMode同期・U1手動維持・部分失敗は親SKU単位・GTIN証跡は状態シート・配送=`送料無料パターン`・リードタイム初版なし。方針Q&A一通り閉じ。次は三点レビュー再実施。
- 2026-07-20: **Lv4 Q7–Q10b反映**: 親＋子個別承認・3モール同一バッチ（Q9=A）・TRACK=BはASIN無視・出品者SKU=子SKU／メーカー型番=メーカー品番（空→子SKU）。
- 2026-07-20: **Lv4社長Q&A反映**: 純正xlsm・D-1（GAS GENERATED＋ローカル PACKAGED）・BはマスタJAN残しGTIN列のみ空・M1バリエーションのみ・失敗後同一subBatchId上書き＋ログ追記。[LV4要件](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) §17。
- 2026-07-20: **Lv4三点レビュー反映**: [org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md)＋要件更新。在庫マスタ書込禁止・`GENERATED`/`UPLOADED_OK`・親SKU抽出・GTIN着手ゲート。コード未実装・次は実装承認。
- 2026-07-20: **Lv4要件ドラフト**: [org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)。M1=Bノーブランド新規（バリエーション＋画像）／M2=A既存ASIN。手動UP・自己配送既定。コード未実装・実装は別承認。
- 2026-07-20: **Lv3人間検収完了**（本番 runId `LV3_20260720_112238_777998` / batch `A1_20260720_083227_1f0b30`・`childrenDone=7`）。在庫列は出品用「在庫数」に一本化（計算列は「在庫数計算」）。CURRENT_PHASE のフォーカスを Lv4 Amazon へ。
- 2026-07-20: **Lv3実装**: `YahooApprovalExport.js`・Zメニュー20。案A（子のみレ点）・主25分／副ユニーク50・`runYahooExport` 呼出のみ。Property `APPROVAL_YAHOO_LV3_ENABLED` 既定 false。手動「▶ 出品実行」非改変。
- 2026-07-20: **Lv3要件ドラフト**: [org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md](org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md)。`runYahooExport` 呼出のみ。分割は主25分・副ユニーク50（楽天運用揃え。Yahoo公式は10,000枚/時）。コード未実装・実装は別承認。
- 2026-07-20: **Lv2人間検収完了**（runId `LV2_20260720_100744_244553` / batch `A1_20260720_083227_1f0b30`・`childrenOn=7`・MAKE/FTP成功）。Property `APPROVAL_RAKUTEN_LV2_ENABLED=false`。CURRENT_PHASE のフォーカスを Lv3 Yahoo へ。
- 2026-07-17: **Lv1人間検収完了**（batch `A1_20260717_224813_ad7e65` → APPROVED）。
- 2026-07-20: **Lv2案A修正**: 一時レ点を「承認済み親＋紐づく子」に変更（通常出品と同じ）。親だけONだとバリエーションがシングルSKU扱いになり32バイト超過で失敗していた。`generateRakutenCSV` 非改変。
- 2026-07-19: **Lv2実装**: `RakutenApprovalExport.js`・Zメニュー19。案Aレ点・ユニーク画像50・25分レジューム。`generateRakutenCSV` 呼出のみ。Property `APPROVAL_RAKUTEN_LV2_ENABLED` 既定 false。手動「楽天CSV出力」非改変。
- 2026-07-17: **Lv2要件ドラフト**: [org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)（聖域呼出・ユニーク50／25分）。
- 2026-07-17: **運用方針docs**: ハイブリッド承認・残リストA（ORPHAN非表示）・日中分割仮既定（GAS実働25分／ユニーク画像50枚）。[PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md) §4.2・§4.3・[AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md) §5。
- 2026-07-17: **Lv1実装**: `ApprovalQueue.js`・Zメニュー18・Webは `?action=approval_queue`（Yahoo doGet分岐）。`APPROVAL_QUEUE_V1_ENABLED` 既定 false。EC書込なし。
- 2026-07-17: **Lv1要件**: [org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)。
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
