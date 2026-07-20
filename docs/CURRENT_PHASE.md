# プロジェクト全体の位置づけと現在の開発フォーカス

**最終更新**: 2026-07-21（深夜・セッション終了）  
**読み方**: 次の Agent は `docs/AGENT_HANDOVER.md` の **§1.5・§2** に従い、**本ファイルを最初に読み**、続けて §2 の必読一覧でプロジェクト全体をインプットする。

---

## 0. セッション終了メモ（2026-07-21 00:25 頃・自宅IDE）

**外出中の続き**: 社長は出勤・**agentWindows 等でリモート操作**予定。この節を最初に読むこと。

### いまの到達点
| 項目 | 状態 |
|------|------|
| Lv4 コード | 実装済＋再レビュー採用修正済（DRY_RUN／冪等／subBatchId単調／ノーブランド品厳密） |
| docs | Q5 を「失敗後は**新** subBatchId」に統一。§14=ドライラン手順 |
| 三点レビュー | 合意協議まで実施。多数派: **コード必須なしで clasp push→ドライラン可** |
| `clasp push` | **成功（2026-07-21）** 7ファイル（`AmazonApprovalExport.js` 含む） |
| 人間ドライラン | **未実施**（今夜はここまで） |

### 明日いちばん最初にやること（ドライラン）
1. スプレッドシートを開く（GAS は push 済み想定）
2. Script Properties:
   - `APPROVAL_AMAZON_LV4_ENABLED=true`
   - `APPROVAL_AMAZON_LV4_TRACK=B`
   - `APPROVAL_AMAZON_LV4_SKIP_EXPORT=true`（必須・初回）
3. メニュー 21-① を一度実行 → シート `▼Lv4実行ログ(Amazon)` 作成
4. EXEMPTION 行: brand=`ノーブランド品`・承認日・証跡URL・カテゴリ（または `*`）
5. 18-① 候補作成 → amazon 親＋子を承認①
6. 21-① → 状態 `DRY_RUN`。マスタ「在庫数」「JAN」不変を確認
7. （任意）Lv2/Lv3 候補件数が従来どおりか目視
8. 作業後は `APPROVAL_AMAZON_LV4_ENABLED=false`（本 GENERATED は別日で `SKIP_EXPORT=false`）

### 触ってよい／ダメ
- **可**: Property・メニュー18/21・状態シート EXEMPTION・ログ確認
- **不可（ドライラン中）**: SC 自動UP想定なし／マスタ在庫・JAN書込なし／楽天CSV・Yahoo.js 改変なし
- **本 GENERATED・PACKAGED・SC UP**: ドライラン合格後の別ゲート（§11 未決あり）

### 正本
- 要件: [org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)
- コード: `AmazonApprovalExport.js` / `ApprovalQueue.js` / `コード.js`（メニュー21）

### 未コミットの可能性
ローカルに Lv4 関連の未コミット変更が残っている可能性あり。リモートで触る前に `git status` を確認。**コミットは社長指示があるまでしない。**

---

## 1. プロジェクト全体（30秒）

| 項目 | 内容 |
|------|------|
| **目的** | 楽天・Yahoo・Amazon 向けに、リサーチ〜商品情報・価格・セット構成〜出品に必要なデータを **スプレッドシート＋GAS** で整備・自動化する。一人社長＋AI部門モデルは [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)。 |
| **クリティカルパス** | [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) — リサーチ・見積もり → 出品情報 → 各モール出品。 |
| **実装の中心** | `コード.js`、主シートは `▼商品マスタ(人間作業用)`、`AI情報取得data`、`00_設定マスタ`。 |
| **6領域・成果物** | [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md)。 |

---

## 2. 現在のフェーズ（いま優先している開発）

- **フォーカス領域**: **Lv4 Amazon バルク（実装済・人間検収待ち）**。  
- **Lv4（Amazonバルク）**: **2026-07-20 実装**（`AmazonApprovalExport.js`／`ApprovalQueue` amazon加算／メニュー21）。  
  - Property: `APPROVAL_AMAZON_LV4_ENABLED`（既定false）・`TRACK=B` 明示必須・GTIN証跡は `▼Lv4実行ログ(Amazon)` の EXEMPTION 行。  
  - 成果: Drive `GENERATED` CSV → ローカル PACKAGED → SC手動UP → 21-③。  
  - **次**: **clasp push 成功（2026-07-21）** → ドライラン検収（SKIP_EXPORT=true・TRACK=B・EXEMPTION記入）。  
- **Lv3（Yahooオーケストレーション）**: **2026-07-20 人間検収完了**（[org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md](org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md)）。  
  - ドライラン `LV3_20260720_111629_027775` → 本番 `LV3_20260720_112238_777998`（`childrenDone=7`）。  
  - 在庫列は出品用「在庫数」（HX）に一本化（計算列は「在庫数計算」）。  
  - Property `APPROVAL_YAHOO_LV3_ENABLED=false` に戻すこと。  
- **Lv2（楽天オーケストレーション）**: **2026-07-20 人間検収完了**。  
- **Lv1（承認キュー）**: **2026-07-17 人間検収完了**。  

- **このフェーズの完了条件（目安）**  
  - Lv4 **ドライラン検収** →（合格後）本 GENERATED／PACKAGED／SC。着手ゲートは対象カテゴリの **GTIN免除＋状態シート証跡**。  

- **並行・継続（後回し可）**: OpenAI 429 / 11-③ vs B Step7。楽天ジャンル Nav Stage1 人間検証。  

- **スコープ外（次モール着手前）**  
  - 承認②、販売中SKU無人上書き、`clasp push` 自動化。  
  - `generateRakutenCSV` 本体・`Yahoo.js` 出品API本体の改変。

---

## 3. 直近で確定した仕様・前提（docs 反映済み）

- **AI組織**: **ハイブリッド**（承認①掲載は軽い／本命は承認②在庫反映）→ 日中無人出品（**ユニーク画像50枚＋実働25分**で自動分割・原則12:00前）→ 夜確認。残リストは明示取消まで残し、マスタ無しは画面非表示（履歴残す）。在庫0原則。販売中(在庫>0)は原則スキップ。CS／問屋は送信禁止・下書き可。詳細は org 憲章・マトリクス・多数決メモ §4.2・§4.3。  
- **キーワード案・商品名案の本線は OpenAI**。Gemini 版の商品名案は **コードに存在するが本流メニュー未配線**。  
- **バリエーション単位**は **Gemini → OpenAI 補完**（◎ASIN 経路も同様）。**内容量パース**は Gemini 中心（ChatGPT 自動フォールバックなし）。  
- **11-③（`runProductNameProposalsForRows`）** は、**OpenAI 失敗時もマスタ商品名があればバリエーションに届く**（既定）。旧挙動は `PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL=false`。**B Step7** は **variation まで進み得る**（商品名が空でも単位・内容量のみ更新されうる）。  
- 詳細は [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md)、[商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md)。

---

## 4. 次にやること（優先順）

1. **Lv4 Amazon ドライラン検収**（push済）。Property ON・TRACK=B・SKIP_EXPORT=true・EXEMPTION → 21-①。詳細は **§0**。  
2. ドライラン合格後: `SKIP_EXPORT=false` で GENERATED → PACKAGED → SC → 21-③（最新 subBatchId）。完了後 ENABLED=false。  
3. （任意）Yahoo未反映の在庫が正しいことを確認してからストア反映。  
4. （並行可）OpenAI 429 / `insufficient_quota` の解消と、11-③ と B Step7 のログ比較。  
5. （並行・できたとき）楽天ジャンル Nav Stage1 人間検証（毎日必須ではない）。

---

## 5. 深掘りリンク（領域別）

| テーマ | ドキュメント |
|--------|----------------|
| **AI組織・承認** | [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)、[org/AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md)、[org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)、[org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)、[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)、[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)、[org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)、[org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md](org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md)、[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)、[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md) |
| 楽天ジャンル Nav（並行） | [RAKUTEN_NAV_GENRE_DIAG.md](RAKUTEN_NAV_GENRE_DIAG.md)（Stage0）、[RAKUTEN_NAV_GENRE_STAGE1.md](RAKUTEN_NAV_GENRE_STAGE1.md)（Stage1要件） |
| Gemini / OpenAI・11-③ vs B Step7・429 | [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md) |
| 商品マスタ人間作業エリア | [商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md) |
| 価格・送料・再③・B 統合 | [PRICING_V1_REQUIREMENTS.md](PRICING_V1_REQUIREMENTS.md)、AGENT_HANDOVER **§8** |
| リサーチ・セット構成 | [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) |
| エージェント共通・必読一覧 | [AGENT_HANDOVER.md](AGENT_HANDOVER.md) **§2** |

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-21 | **セッション終了**: clasp push成功。§0に外出用引き継ぎ。次はドライラン検収。 |
| 2026-07-21 | **Lv4**: docsを新subBatchIdに統一。clasp pushは invalid_rapt で未完了 → `clasp login` 後に再実行。 |
| 2026-07-20 | **Lv4 実装**: AmazonApprovalExport／ApprovalQueue amazon加算／メニュー21。次は人間検収。 |
| 2026-07-20 | **Lv4 第3回三点＋Q15/Q16**: TRACK未設定＝非実行／amazon抽出同一チケット／列メモ注記。次は実装承認。 |
| 2026-07-20 | **Lv4 Q11–Q14反映**（inventoryMode・親単位再試行・GTIN証跡シート・送料無料パターン）。方針Q&Aは一通り閉じた。次は三点レビュー再実施→実装承認。 |
| 2026-07-20 | **Lv4 Q7–Q10b反映**（3モール同一承認①・TRACK=B強制・出品者SKU=子SKU／メーカー品番）。次は残未決→実装承認。 |
| 2026-07-20 | **Lv4社長Q&A反映**（純正xlsm・D-1 PACKAGED・BはマスタJAN残し・再生成上書き＋ログ追記）。次は残Q（Lv1抽出等）→実装承認。 |
| 2026-07-20 | **Lv4三点レビュー反映**（[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md)）。在庫書込禁止・DONE分離・親SKU抽出・GTIN着手ゲート。次は実装承認。 |
| 2026-07-20 | **Lv4要件ドラフト**（[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)）。M1=Bノーブランド新規／M2=A既存。 |
| 2026-07-20 | **Lv3人間検収完了**（本番 runId `LV3_20260720_112238_777998`）。フォーカスを Lv4 Amazon へ。在庫列「在庫数／在庫数計算」分離を§8.2相当で記録。 |
| 2026-07-20 | **Lv2案A修正**: 一時レ点を親＋子に変更（バリエーション商品のシングルSKU誤認を解消）。CSV本体非改変。次は再検収（clasp push → 19-①）。 |
| 2026-07-19 | **Lv2実装**（`RakutenApprovalExport.js`・メニュー19・案A）。次は人間検収。手動楽天CSVは非改変。 |
| 2026-07-17 | **Lv1人間検収完了**（batch `A1_20260717_224813_ad7e65` → APPROVED）。フォーカスを Lv2要件確認／実装承認へ。 |
| 2026-07-17 | **Lv2要件ドラフト**（[org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)）。実装は Lv1検収後。 |
| 2026-07-17 | ハイブリッド・残リストA・実行分割仮既定（25分／ユニーク50）を§3に反映。 |
| 2026-07-17 | **Lv1実装**（ApprovalQueue.js・メニュー18・Web approval_queue）。次は人間検収。 |
| 2026-07-17 | **Lv1要件**追加（[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)）。 |
| 2026-07-17 | **Lv0最終承認**反映。モール順楽天先。フォーカスを Lv1 へ。上書きは当面手動(U1)。 |
| 2026-07-17 | **Lv別実装プラン叩き台**追加（[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)、コードなし）。 |
| 2026-07-17 | 並行: 楽天ジャンル Nav Stage1 実装（専用診断シート追記・Property 既定オフ）。 |
| 2026-07-15 | **AI組織 Phase0**: 3者多数決反映・RUNBOOK（親1＋並列3）・多数決メモ追加。実装は次フェーズ。 |
| 2026-07-15 | **AI組織 Phase0**: org 憲章・承認マトリクスをフォーカスに。実装は次フェーズ。 |
| 2026-03-19 | 初版。全体＋現在フェーズの引き継ぎを本ファイルに集約。 |
| 2026-03-22 | OpenAI 失敗時の 11-③ バリエーション継続（既定）を §3 に反映。 |
