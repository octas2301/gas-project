# プロジェクト全体の位置づけと現在の開発フォーカス

**最終更新**: 2026-07-25（U2 三点＋社長回答反映。コミット待ち）  
**読み方**: 次の Agent は `docs/AGENT_HANDOVER.md` の **§1.5・§2** に従い、**本ファイルを最初に読み**、続けて §2 の必読一覧でプロジェクト全体をインプットする。

---

## 0. セッション引き継ぎ（2026-07-25）

**場所**: **自宅PC（ローカル）**。この節を最初に読むこと。  
**コミット**: **指示まで待つ**（三点反映 docs 未コミットの可能性）。  
**clasp push**: T2・U3 済。

### 自宅PC起動時に貼る一文
> `docs/CURRENT_PHASE.md` §0 を最初に読め。**U2 三点＋社長回答反映済**（マスタ永続／sheet復元、ONLY=sheet、候補=Amazon用フォルダ）。次は **コミット指示** → **U2実装前承認**。T3／ε／U7は各ゲート。

### いまの到達点（確定）
| 項目 | 状態 |
|------|------|
| Lv4 HPC／FOOD／T2 | **完了** |
| D×Amazon 要件 U0 | **クローズ** |
| **U3 D UI** | **v1 実機合格** |
| **U2 C×Amazon** | **三点＋社長回答反映済**。GAS未。次＝実装承認 |
| T3／ε／Dc | 各ゲート／バックログ |

### Amazon画像方針（要約）
| 項目 | 方針 |
|------|------|
| MAIN紐付け | マッチング sheet・子SKU・人間 |
| 永続 | **マスタ列**。再生成後はマスタ→sheet 復元 |
| 候補 | **Amazon 用フォルダ**（楽天と分離） |
| 出口 | Drive `02` へ**コピー** |
| ONLY PT | sheet 本線（`02`手置き＝例外） |
| ε | バックログ |

### 次にやること（優先順）
1. **コミット**（指示時）— MAJORITY＋U2要件＋POC 等  
2. **U2 実装前承認**（変更ファイル一覧／概要／リスク＋HUMAN_RUN）  
3. 実装 → clasp push → 実機  
4. T3／ε／U7 は各ゲート  

### IDs（常用）
```
U3 smoke runId: LV4_20260725_072635_948892
T2 runId: R2T2_20260724_221107_7f9cf7
subBatchId HPC: A1_20260721_083100_06b90a_B2
Drive 02 ID: 1T6_E6T-qd9whSF8Re8lyRVB2n-P4BM84
R2 public: https://pub-d974bd81c7d84f9bbc65f8479d3f85d4.r2.dev
```

### 正本・手順リンク
- [org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md](org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md)  
- [org/D_MENU_U2_THREE_REVIEW_MAJORITY.md](org/D_MENU_U2_THREE_REVIEW_MAJORITY.md)  
- [org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md](org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md)  
- [org/D_MENU_U3_HUMAN_RUN.md](org/D_MENU_U3_HUMAN_RUN.md)  
- [org/D_MENU_AMAZON_FACADE_THREE_REVIEW_MAJORITY.md](org/D_MENU_AMAZON_FACADE_THREE_REVIEW_MAJORITY.md)  
- [org/LV4_T2_HUMAN_RUN.md](org/LV4_T2_HUMAN_RUN.md)  
- [org/LV4_R2_IMAGE_PIPELINE_POC.md](org/LV4_R2_IMAGE_PIPELINE_POC.md)  

---

## 1. プロジェクト全体（30秒）

| 項目 | 内容 |
|------|------|
| **目的** | 楽天・Yahoo・Amazon 向けに、リサーチ〜商品情報・価格・セット構成〜出品に必要なデータを **スプレッドシート＋GAS** で整備・自動化する。一人社長＋AI部門モデルは [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)。 |
| **クリティカルパス** | [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) — リサーチ・見積もり → 出品情報 → 各モール出品。 |
| **実装の中心** | コード.js、主シートは ▼商品マスタ(人間作業用)、AI情報取得data、00_設定マスタ。 |
| **6領域・成果物** | [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md)。 |

---

## 2. 現在のフェーズ（いま優先している開発）

- **フォーカス領域**: **Lv4 — U2三点反映済／次はコミット→実装承認**（§0）。  
- **Lv4（Amazonバルク）**:  
  - HPC／FOOD／T2: 完了。  
  - 画像: マスタ永続／sheet復元、候補=Amazon用フォルダ、`02`=コピー出口、ONLY=sheet。ε＝バックログ。  
  - 本線UX: U3 実機合格。Cは [org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md](org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md)。  
- **楽天Nav Stage1**: 実機PASS。Propertyはfalse。  
- **Lv3（Yahoo）**: 2026-07-20 人間検収完了。  
- **Lv2（楽天）**: 2026-07-20 人間検収完了。  
- **Lv1（承認キュー）**: 2026-07-17 人間検収完了。  

- **このフェーズの完了条件（目安）**  
  - HPC／FOOD／T2／U0／U3／**U2-0 三点反映**: **達成**。  
  - 次: コミット → U2 実装承認。  

- **並行・継続（後回し可）**: T3実装承認／Dc API／M2／21-⑤／xlsm自動C1。  

- **スコープ外（次モール着手前）**  
  - 承認②、販売中SKU無人上書き、clasp push 自動化。  
  - generateRakutenCSV 本体・Yahoo.js 出品API本体の改変。

---

## 3. 直近で確定した仕様・前提（docs 反映済み）

- **AI組織**: **ハイブリッド**（承認①掲載は軽い／本命は承認②在庫反映）→ 日中無人出品（**ユニーク画像50枚＋実働25分**で自動分割・原則12:00前）→ 夜確認。残リストは明示取消まで残し、マスタ無しは画面非表示（履歴残す）。在庫0原則。販売中(在庫>0)は原則スキップ。CS／問屋は送信禁止・下書き可。詳細は org 憲章・マトリクス・多数決メモ §4.2・§4.3。  
- **キーワード案・商品名案の本線は OpenAI**。Gemini 版の商品名案は **コードに存在するが本流メニュー未配線**。  
- **バリエーション単位**は **Gemini → OpenAI 補完**（◎ASIN 経路も同様）。**内容量パース**は Gemini 中心（ChatGPT 自動フォールバックなし）。  
- **11-③（runProductNameProposalsForRows）** は、**OpenAI 失敗時もマスタ商品名があればバリエーションに届く**（既定）。旧挙動は PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL=false。**B Step7** は **variation まで進み得る**（商品名が空でも単位・内容量のみ更新されうる）。  
- 詳細は [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md)、[商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md)。

---

## 4. 次にやること（優先順）

1. **コミット**（指示時）。  
2. **U2 実装前承認** → 実装。  
3. T3・ε・U7・M2 は各ゲート。  
4. Property トグルは実行時のみ true。  

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
| 2026-07-25 | **U2三点＋社長回答反映**（MAJORITY新規）。コミット待ち。 |
| 2026-07-25 | **U2方針確定**: 案α本線・MAIN=sheet／02=出口・εバックログ。 |
| 2026-07-25 | **U3実機合格**＋**U2要件起草**。runId `LV4_20260725_072635_948892`。 |
| 2026-07-25 | **D×Amazon U3 v1**: D `amazon`/`full_amazon`・即時ファサード。clasp push待ち。 |
| 2026-07-25 | **D×Amazon U0クローズ**: 3者反映＋社長回答。手ZIP正／T3実装待ち／将来API必須。 |
| 2026-07-24 | **D×Amazon要件U0**: [org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md](org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md)。T3保留。次=社長承認→3者。 |
| 2026-07-24 | **T2 PoC成功**: runId `R2T2_20260724_221107_7f9cf7`・URL画像表示。トグルfalseへ。次=T3要承認。 |
| 2026-07-24 | **T2 clasp push済**: 8 files（`AmazonDriveImageExport.js`含む）。次=Property＋21-⑥→URL200→トグルoff。 |
| 2026-07-24 | **帰宅引き継ぎ**: §0全面更新。FOOD成功・Nav PASS・T2実装済・次=自宅 clasp push＋21-⑥。 |
| 2026-07-23 | **帰宅引き継ぎ**: HPC suburl試験UP中・FOOD v4再解析（再UP禁止）。§0全面更新。 |
| 2026-07-23 | **§11.0 HPCクローズ**: 1〜8・10＋U5。画像ZIP優先。正本 titlefix。FOOD／M2／21-⑤は別ゲート。 |
| 2026-07-22 | **帰宅**: R2＋8枚MAINは200確認。§0をURL埋め→SC再UP向けに更新。 |
| 2026-07-22 | **自宅IDE引き継ぎ**: DRY_RUN＋本GENERATED成功（..._B2）。当時はPACKAGED→SC。 |
| 2026-07-21 | **セッション終了（深夜）**: clasp push成功。当時は次＝ドライラン（完了済み）。 |
| 2026-07-21 | **Lv4**: docsを新subBatchIdに統一。clasp pushは invalid_rapt で未完了 → clasp login 後に再実行。 |
| 2026-07-20 | **Lv4 実装**: AmazonApprovalExport／ApprovalQueue amazon加算／メニュー21。次は人間検収。 |
| 2026-07-20 | **Lv4 第3回三点＋Q15/Q16**: TRACK未設定＝非実行／amazon抽出同一チケット／列メモ注記。次は実装承認。 |
| 2026-07-20 | **Lv4 Q11–Q14反映**（inventoryMode・親単位再試行・GTIN証跡シート・送料無料パターン）。方針Q&Aは一通り閉じた。次は三点レビュー再実施→実装承認。 |
| 2026-07-20 | **Lv4 Q7–Q10b反映**（3モール同一承認①・TRACK=B強制・出品者SKU=子SKU／メーカー品番）。次は残未決→実装承認。 |
| 2026-07-20 | **Lv4社長Q&A反映**（純正xlsm・D-1 PACKAGED・BはマスタJAN残し・再生成上書き＋ログ追記）。次は残Q（Lv1抽出等）→実装承認。 |
| 2026-07-20 | **Lv4三点レビュー反映**（[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md)）。在庫書込禁止・DONE分離・親SKU抽出・GTIN着手ゲート。次は実装承認。 |
| 2026-07-20 | **Lv4要件ドラフト**（[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)）。M1=Bノーブランド新規／M2=A既存。 |
| 2026-07-20 | **Lv3人間検収完了**（本番 runId LV3_20260720_112238_777998）。フォーカスを Lv4 Amazon へ。在庫列「在庫数／在庫数計算」分離を§8.2相当で記録。 |
| 2026-07-20 | **Lv2案A修正**: 一時レ点を親＋子に変更（バリエーション商品のシングルSKU誤認を解消）。CSV本体非改変。次は再検収（clasp push → 19-①）。 |
| 2026-07-19 | **Lv2実装**（RakutenApprovalExport.js・メニュー19・案A）。次は人間検収。手動楽天CSVは非改変。 |
| 2026-07-17 | **Lv1人間検収完了**（batch A1_20260717_224813_ad7e65 → APPROVED）。フォーカスを Lv2要件確認／実装承認へ。 |
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
