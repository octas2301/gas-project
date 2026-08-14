# B Step1 — ヘッダー直後への上挿入（末尾追記廃止）

**文書種別**: 要件定義（正本）  
**日付**: 2026-08-14  
**状態**: **実装済み（ローカル）** — 人間: `clasp push` ＋ Script Property `B_STEP1_TOP_INSERT_ENABLED=true` ＋ HUMAN_RUN  
**親**: [CURRENT_PHASE.md](../CURRENT_PHASE.md)／セット構成提案（`コード.js`）  
**三点レビュー**: [B_STEP1_TOP_INSERT_THREE_REVIEW_MAJORITY.md](B_STEP1_TOP_INSERT_THREE_REVIEW_MAJORITY.md)  
**手順**: [B_STEP1_TOP_INSERT_HUMAN_RUN.md](B_STEP1_TOP_INSERT_HUMAN_RUN.md)  
**ゴール一文**: 新規親子をヘッダー直下に挿入し、**既存出品SKU／親SKUは不変**のまま、新規だけ重複時連番する。末尾追記は Property OFF で残置。

---

## 0. 社長ロック（2026-08-13〜14）

| ID | 決定 |
|----|------|
| P1 | 挿入位置＝**本番データ先頭の直前**（見本・バッファの下。最新がデータ最上段） |
| P2 | **末尾追記は廃止**（既定は Property OFF＝旧挙動。ON で上挿入） |
| P3 | 行確保（式・CF・入力規則）→商品データ上書きを **同一改修**に含める |
| **1A** | **既存 AK／IB は絶対不変**（変えるのはNG） |
| **2A** | 同JAN再B時、連番・枝番は**新品だけ**（既存はNG） |
| **3A** | 式＋上挿入＋戻しを同一単位。Propertyで戻せる。**実行前に対象ファイルをコピー退避** |
| T1 | **1行目（親テンプレ）・2行目（子テンプレ）の式更新は必須** |
| T2 | **既存データ行の AK／IB は値固定**（式を残して全域化しない） |

### 0.1 ローカル復元点（2026-08-14・実装前）

| 項目 | 値 |
|------|-----|
| フォルダ | `_local_backup/pre_B_STEP1_TOP_INSERT_20260814_001210/`（`.gitignore`） |
| 退避 | `コード.js`／`Yahoo.js`／`appsscript.json`／`.clasp.json`／`.claspignore` |
| Git HEAD（参考） | `2f9ef6c` |
| 手順 | 同フォルダ内 `RESTORE.md` |
| スプシ | Drive コピー済み前提。値固定・テンプレは本番マスタへ適用済み（2026-08-14） |

### 0.1b 実機レイアウト確認（2026-08-14・Sheets API）

シート: `▼商品マスタ(人間作業用)`（SS `1LIWp0…`）

| 行 | 内容 |
|----|------|
| **8** | **ヘッダー**（`ASINコード` あり。コードの `ANCHOR_HEADER_NAME` 検出行） |
| **9** | 細いバッファ行（ほぼ空・高さ約21px） |
| **10** | **見本行**（A列に `見本コピーしない⇒`） |
| **11〜** | **本番データ先頭**（A=`1` / B=`No.1` …） |

現行実装（改修後）: `masterFindTopInsertStartRow1Based_` → **見本の次＝本番データ先頭（現状11行目）の直前**に挿入。  
バッファ9・見本10は維持。`insertRowsBefore` は **25行チャンク**（Sheets切断対策）。

### 0.2 実装結果（2026-08-14）

| # | 内容 | 結果 |
|---|------|------|
| F3 | 既存 AK/IB 値固定 | AK 式1383件・IB 式643件 → 表示値で固定（`tools/_tmp_freeze_ak_ib_and_templates.py`） |
| F1/F2 | テンプレ1–2 | AK=上下COUNTIF合算（自己除外・循環なし）。IB親=`COUNTIFS(I$9:I,HY$9:HY,"")`／子=`=IB1` |
| C1/C3 | Step1 上挿入 | `insertRowsBefore(header+1, totalNewRows)` → AI順書込。Property ON 時のみ |
| C2 | スラッグ | `masterHasSameJanParentExists_`（全体検索） |
| C7 | Property | `B_STEP1_TOP_INSERT_ENABLED`（**未設定=false＝旧末尾追記**） |
| C8 | フィルタ | 有効時は中断 |
| — | 15-㉒ | データ全行（ヘッダー直下〜最終行）に相対化 |

**確定テンプレ式（要約・下カウント統一）**
- **AK（1・2行同じ形）**: `count = COUNTIF(INDIRECT("AK"&(ROW()+1)&":AK"), base_id&"*")` のみ（下方向）。0なら base、否则 `base-(count+1)`。上下合算は廃止（循環→`#REF!` 対策）
- **IB親**: `below = COUNTIFS(下方向 I / HY="")` → `seq=below+1` → oya / oyaN。**子IB**は `=IB親` のまま
- 同一実行内で枝番の付き方が上下逆でも、**ユニークなら可**（社長確認済み）

**調査ログ**: `[STEP1_TOP][runId]` … START / topInsert+writeStartRow / plan / insertDone / block / DONE|FAIL

### 0.2 実装結果（2026-08-14）

| # | 対象 | 改修 |
|---|------|------|
| C1 | B Step1 | ✅ データ先頭（見本の次）へ分割 insert（Property ON） |
| C2 | カタログスラッグ | ✅ マスタ全体 |
| C3 | 複数AI行 | ✅ 一括 insert 後 AI 順 |
| C7 | 復元 | ✅ Property + `_local_backup` |
| C8 | フィルタ | ✅ 中断 |

### 本改修から外す

| # | 内容 |
|---|------|
| C4 | Yahoo.js「下の親優先」→ 改変しない |
| C5 | 抜けセット数（15-⑩）→ 現状維持 |
| C6 | `RESEARCH_AND_ESTIMATE.md` §20-1 → 非改変。`メニューとフロー_サマリー.md` §6.3 は更新済み |

---

## 2〜4. （方針・検収は HUMAN_RUN / 多数決を正）

実装チェックリスト:
- [x] F3 値固定
- [x] F1/F2 テンプレ
- [x] C1/C2/C3/C7/C8
- [x] 15-㉒ 相対化
- [x] Step1: 依頼日(D) をテンプレ1/2からコピー（`STEP1_PRE_Q_FORMULA_COLUMNS_`）

## 5. 聖域・非改変

- `generateRakutenCSV`／Yahoo API 本体／Yahoo 親マージ／`B_INTEGRATED_STEP_FUNCTIONS` 順

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-14 | 実装（スプシ値固定＋テンプレ＋GAS上挿入）。状態を実装済み（ローカル）へ |
| 2026-08-14 | 1A/2A/3A・テンプレ1–2・既存値固定をロック |
| 2026-08-13 | 初版 |
