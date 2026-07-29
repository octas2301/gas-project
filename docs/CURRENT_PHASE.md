# プロジェクト全体の位置づけと現在の開発フォーカス

**最終更新**: 2026-07-29（SP-API **v1.4 実機合格（API）**／**第2段 承認済・実装待ち**）  
**読み方**: 次の Agent は `docs/AGENT_HANDOVER.md` の **§1.5・§2** に従い、**本ファイルを最初に読み**、続けて §2 の必読一覧でプロジェクト全体をインプットする。

---

## 0. セッション引き継ぎ（2026-07-29）

**場所**: 自宅PC。  
**コミット**: 本セッションで v1.4 合格＋ENDPOINT正規化＋**第2段承認起草**。  
**clasp**: ENDPOINT正規化は **push 済（`-f`）**。第2段はコード未着手のため push 不要。  
**Property**: PUT／ALLOW_PROD／EXPORT とも作業後 **false**。`SPAPI_ENDPOINT` は空推奨。

### 起動時に貼る一文
> `docs/CURRENT_PHASE.md` §0 を読め。SP-API **v1.4 第1段は実機合格（API）**（21-⑩ VALID／21-⑪ ACCEPTED・`…48s11`）で push 済。次＝**第2段（承認①済→GAS直PUT）が 2026-07-29 承認済・実装待ち**。範囲は §2・§3 に限定し、Property は第1段と共用。[第2段承認](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)／[GAS PUT HUMAN_RUN](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)。

### いまの到達点
| 項目 | 状態 |
|------|------|
| 安眠 E〜E-5／C1 | **済** |
| **メニュー8** | v1.10＋キャッチ1行化（push済み認識） |
| **M2** | **実機合格** |
| **SP-API読取／書込 v1・v1.1** | **実機合格** |
| **SP-API橋渡し v1.2〜v1.3／v1.2b／v1.2c** | **実機合格** |
| **SP-API v1.4 第1段** | **実機合格（API）・push済**。SC最終更新目視は反映待ち。[承認](org/LV4_SPAPI_GAS_PUT_APPROVAL.md)／[HUMAN_RUN](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md) |
| **SP-API v1.4 第2段** | **承認済・実装待ち**（2026-07-29）。[承認](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md) |
| 試験SKU | `…0924-48s11`／`B00A0J0D30`（v1.4）ほか発汗・ride01 |

### 次にやること（優先順）
1. ~~v1.4 承認・実装・API実機~~／~~clasp push -f（ENDPOINT正規化）~~／~~第2段 社長承認~~  
2. **Agent**: v1.4 第2段を実装（承認①済→GAS直PUT・[承認](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md) §2・§3 の範囲厳守）→ clasp push → dry_run  
3. **人間**: SC最終更新が prod 以降に進んだら HUMAN_RUN の待ちチェックを閉じる  
4. Frontier／Sandbox は触らない  

### M2 正本リンク
- [org/D_MENU_M2_HUMAN_RUN.md](org/D_MENU_M2_HUMAN_RUN.md)  
- [org/LV4_M2_IMPLEMENTATION_APPROVAL.md](org/LV4_M2_IMPLEMENTATION_APPROVAL.md)  
- [org/LV4_M2_TRACK_A_GAP_ANALYSIS.md](org/LV4_M2_TRACK_A_GAP_ANALYSIS.md)  
- `tools/m2_offer_packaged/`  
- [org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md](org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md)／`tools/spapi_smoke/`  
- [org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md](org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)／`tools/spapi_listings_write/`  
- [org/LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md](org/LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md)（v1.1）  
- [org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)／[org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md](org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md)  
- [org/LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md](org/LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md)（v1.2c）  
- [org/LV4_SPAPI_DRIVE_FETCH_APPROVAL.md](org/LV4_SPAPI_DRIVE_FETCH_APPROVAL.md)／[org/LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md](org/LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md)  
- [org/LV4_SPAPI_GAS_PUT_APPROVAL.md](org/LV4_SPAPI_GAS_PUT_APPROVAL.md)／[org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)（v1.4 第1段・API実機合格／SC最終更新は反映待ち）  
- [org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)（v1.4 **第2段・承認済／実装待ち**）  

### 外出先チェックリスト
[org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md](org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md) — 安眠完了済み（履歴）  

### 正本・手順リンク
- [org/LV4_M2_IMPLEMENTATION_APPROVAL.md](org/LV4_M2_IMPLEMENTATION_APPROVAL.md)（**M2 v1・実機合格**）  
- [org/LV4_M2_TRACK_A_GAP_ANALYSIS.md](org/LV4_M2_TRACK_A_GAP_ANALYSIS.md)／[org/D_MENU_M2_HUMAN_RUN.md](org/D_MENU_M2_HUMAN_RUN.md)（**公式Loader正**）  
- [org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md](org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)（安眠完了・履歴）  
- [org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md](org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)  
- [org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)／[org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)  
- [YAHOO_CATEGORY_BRAND_STAGE.md](YAHOO_CATEGORY_BRAND_STAGE.md)／[org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md](org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md)／[org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md](org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md)  
- [org/D_MENU_C1_HUMAN_RUN.md](org/D_MENU_C1_HUMAN_RUN.md)  
- [org/D_MENU_C1_MASTER_HPC_COLUMN_MAP.md](org/D_MENU_C1_MASTER_HPC_COLUMN_MAP.md)  
- [org/D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](org/D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)  

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

- **フォーカス領域**: **Lv4 — U2/U3 実機合格。次は T3／ε 等ゲート**（§0）。  
- **Lv4（Amazonバルク）**:  
  - HPC／FOOD／T2／U3／**U2**: 完了。  
  - U2: `AmazonImageMatrixExport.js`＋C-Amazon①〜④＋C子レ点。  
  - 本線UX記録: [org/D_MENU_U2_HUMAN_RUN.md](org/D_MENU_U2_HUMAN_RUN.md) §0。  
- **楽天Nav Stage1**: 実機PASS。Propertyはfalse。  
- **Lv3（Yahoo）**: 2026-07-20 人間検収完了。  
- **Lv2（楽天）**: 2026-07-20 人間検収完了。  
- **Lv1（承認キュー）**: 2026-07-17 人間検収完了。  

- **このフェーズの完了条件（目安）**  
  - U2: **実機合格達成**。  

- **並行・継続（後回し可）**: T3実装承認／Dc API／M2／21-⑤／ε（C1は本線・§0優先）。  

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

1. Property `AMAZON_IMAGE_U2_ENABLED=false`（未戻しなら）。  
2. T3／ε／U7 は各ゲート承認後。  
3. remote `git push` は指示時。  

---

## 5. 深掘りリンク（領域別）

| テーマ | ドキュメント |
|--------|----------------|
| **AI組織・承認** | [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)、[org/AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md)、[org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)、[org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)、[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)、[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)、[org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)、[org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md](org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md)、[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)、[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md) |
| 楽天ジャンル Nav（並行） | [RAKUTEN_NAV_GENRE_DIAG.md](RAKUTEN_NAV_GENRE_DIAG.md)（Stage0）、[RAKUTEN_NAV_GENRE_STAGE1.md](RAKUTEN_NAV_GENRE_STAGE1.md)（Stage1）、[RAKUTEN_NAV_GENRE_STAGE3.md](RAKUTEN_NAV_GENRE_STAGE3.md)（**Stage3都度API・実装済・要push**） |
| Gemini / OpenAI・11-③ vs B Step7・429 | [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md) |
| 商品マスタ人間作業エリア | [商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md) |
| 価格・送料・再③・B 統合 | [PRICING_V1_REQUIREMENTS.md](PRICING_V1_REQUIREMENTS.md)、AGENT_HANDOVER **§8** |
| リサーチ・セット構成 | [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) |
| エージェント共通・必読一覧 | [AGENT_HANDOVER.md](AGENT_HANDOVER.md) **§2** |

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-29 | **SP-API v1.4 第2段 承認**（社長「承認する」）。実装可・範囲は承認 §2・§3。 |
| 2026-07-29 | **SP-API v1.4 第2段 承認起草**（承認①済→GAS直PUT・コードなし）。三者レビュー不要。 |
| 2026-07-29 | **SP-API v1.4 実機合格（API）**＋ENDPOINT正規化。SC最終更新は反映待ち。 |
| 2026-07-29 | **SP-API v1.4 実装**: 21-⑩⑪／`AmazonSpapiPut.js`。要 clasp push→実機。 |
| 2026-07-29 | **SP-API v1.4 承認起草**（GAS直呼び・コードなし）。三者レビュー不要。 |
| 2026-07-29 | **SP-API v1.2b 実機合格**＋親レ点出さないを完了扱い。 |
| 2026-07-29 | **SP-API v1.2c／v1.3 実機合格**: 子レ点→Drive→fetch-drive prod（`…48s11`）。 |
| 2026-07-28 | **SP-API v1.2c**: 21-⑧＝子SKUレ点のみ（選択行廃止）。 |
| 2026-07-26 | **メニュー8緊急修正**: sync除外。要 clasp push。 |
| 2026-07-26 | **メニュー8 v1実装**: Z→7.5・空欄のみ採用。次＝clasp push→HUMAN_RUN。 |
| 2026-07-26 | **メニュー8要件＋承認パッケージ**: Amazon AI生成＆一括採用（M-A・空欄のみ・要確認列ごと）。実装承認待ち。 |
| 2026-07-26 | **C1-1b実装**（master_csv・必須列・タックスはマスタ）。次＝未送信SKUでSC。 |
| 2026-07-26 | **C1列マップ下書き**（成功PACKAGED差分・SC必須列）。次＝C1-1b。 |
| 2026-07-26 | **C1実装承認**（`tools/c1_hpc_packaged`・HUMAN_RUN手順）。次＝実機。 |
| 2026-07-26 | **C1三点反映**（MAJORITY・URL空スキップ・親一式除外・指紋v1本番停止）。次＝実装承認。 |
| 2026-07-26 | **C1要件起草**（ローカル本線・HPC・DRY_RUN・03新規）。次＝3者。 |
| 2026-07-26 | **U4 実機合格**: 21-⑦ `U4_20260726_090920_1366af`・マスタ URL 確認。D冪等は想定内。 |
| 2026-07-26 | **U4 v1 実装**: 21-⑦・GENERATED Amazon URL優先。次=clasp push＋HUMAN_RUN。 |
| 2026-07-26 | **U4要件＋承認パッケージ起草**。次＝社長実装承認。 |
| 2026-07-26 | **T2再検証合格**: `80s10` URL単独・18320なし・店頭OK。T3急がない。候補退避 clasp push済。 |
| 2026-07-25 | **U2 実機合格**: ②③④・`02` MAIN1件。Dは冪等ブロック想定内（`LV4_20260725_094425_914290`）。 |
| 2026-07-25 | **C子レ点選定**＋候補`07` Property済。次=clasp push→HUMAN_RUN。 |
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
