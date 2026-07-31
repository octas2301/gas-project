# C1 — PACKAGED（純正 xlsm）ほぼ自動 要件

**文書種別**: 要件定義（**実装承認済・HUMAN_RUN待ち**）  
**最終更新**: 2026-07-26  
**状態**: **HPC実機合格／C1-1c SEASONING実装・機械DRY_RUN合格**。次＝吉野家実データDRY_RUN
**親**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md) §6.1 Da／§6.2 Db（xlsm自動＝C1）  
**設計参照**: [LV4_R2_IMAGE_PIPELINE_POC.md](LV4_R2_IMAGE_PIPELINE_POC.md) §5  
**多数決**: [D_MENU_C1_THREE_REVIEW_MAJORITY.md](D_MENU_C1_THREE_REVIEW_MAJORITY.md)  
**HUMAN_RUN**: [D_MENU_C1_HUMAN_RUN.md](D_MENU_C1_HUMAN_RUN.md)  
**列マップ下書き**: [D_MENU_C1_MASTER_HPC_COLUMN_MAP.md](D_MENU_C1_MASTER_HPC_COLUMN_MAP.md)  
**行構造の正**: [LV4_HPC_M1_PACKAGED_RUNBOOK.md](LV4_HPC_M1_PACKAGED_RUNBOOK.md)／LV4要件 §11.5.3  
**依存クローズ**: U2／U3／U4 実機合格・T2再検証合格。画像URLは U4 マスタ列を前提。

**状態メモ（2026-07-26）**: 指紋→DRY_RUN→prodのファイル生成はスモーク成功。SC事前チェックは必須列不足で失敗（想定内）。次＝列マップ確定→**C1-1b**。  

---

## 1. ゴール一文

Drive **06 の純正 HPC `.xlsm`** を読み、GENERATED＋マスタ（**`Amazon MAIN URL` 優先**）から **新規データ行一式**を埋め、Drive **03 に新規 PACKAGED ファイル**として保存する。  
**v1 実装本線はローカル／Cursor**（GASは準備・指示・ログまで）。GAS による純正 xlsm 直編集は v1 対象外。

---

## 2. 方針ロック（2026-07-26・社長／三点反映）

| 項目 | 確定 |
|------|------|
| ゴール | 本格 PACKAGED（新規行一式のほぼ自動） |
| 実装本線 | **B: ローカル／Cursor** |
| GAS | CSV／マスタURL準備・配置案内・ログ。**xlsm直編集はしない**（スパイク成功後のオプション） |
| 画像URL | マスタ **`Amazon MAIN URL`（U4）優先**。**空ならスキップ＋ログのみ**（GENERATED／楽天CDNへフォールバックしない）。URL欠落も親一式除外 |
| 原本 | Drive **06 読取のみ**（破壊禁止） |
| 出力 | Drive **03 に新規コピー**（既存 PACKAGED 上書き禁止） |
| Product Type | v1=`HEALTH_PERSONAL_CARE`、C1-1c=`SEASONING`（七味唐辛子）。PT別マップ・指紋を分離 |
| 失敗単位 | **必須欠落・URL欠落・親子不整合は親SKU一式を出力から除外**。他親は続行可。途中例外はファイル残す＋ログ |
| DRY_RUN | **必須**（本番前1回以上）。本番トグル分離 |
| テンプレ検知 | **v1必須**: ヘッダー／属性行指紋不一致 → DRY_RUN警告＋**本番停止**。accepted_values_db は **C2** |
| 優先 | C1を次の本線（T3急がない） |
| レビュー | ~~要件確定前に3者~~ **済**（[MAJORITY](D_MENU_C1_THREE_REVIEW_MAJORITY.md)） |

---

## 3. スコープ

### 3.1 作るもの（C1 v1）

1. **入力**  
   - 06: `HEALTH_PERSONAL_CARE.xlsm`（純正・読取）  
   - Drive `01` 等の最新 GENERATED（CSV／メタ）。特定規則は実装承認で固定  
   - マスタ: **`Amazon MAIN URL`（必須）**、SKU・価格等（GENERATEDと整合）。`Amazon PT URL` は任意（v1必須ではない）  
   - 対象選定: subBatchId／親SKUリスト／承認①連携のいずれか（実装承認で入力契約固定）  
2. **処理（ローカル／Cursor）**  
   - 06をコピー → HPC行構造に従い親子一式を書く（§4）  
   - メイン画像URL列に U4 `Amazon MAIN URL` を書く  
   - ヘッダー指紋チェック（§4）  
   - DRY_RUN: セル対応表・件数・欠落必須列・**除外親SKU一覧**をレポート（ファイル非作成または `_DRYRUN` 名）  
3. **出力**  
   - 03 へ `{runIdまたはsubBatchId}_PACKAGED_….xlsm` **新規**（親SKUを含む一意名推奨）  
4. **ログ**  
   - runId／sku一覧／書いた行／欠落／除外親／警告を Drive `05` またはリポジトリ外作業ログへ  
   - GAS状態シートへの `PACKAGED` 書込は当面必須としない（05ログ＋人間が21-③確認）  
5. **（任意・薄い）GAS**  
   - GENERATED置場・マスタURL確認ダイアログ／HUMAN_RUNメニュー案内のみ。xlsmバイナリは触らない  

### 3.2 作らないもの

| 対象 | 理由 |
|------|------|
| 06 原本の上書き | 聖域 |
| 03 既存 PACKAGED の上書き | 原因追及・復元のため禁止 |
| **マスタへの書込**（在庫・JAN・価格・URL） | 聖域。C1は読取のみ |
| FOOD／他 PT | `SEASONING`のみC1-1cで追加。その他PTは別ゲート |
| GAS での `.xlsm` 直編集 | v1外（POCどおり） |
| T3 ZIP・ε・U7・楽天CSV・Yahoo.js | 別ゲート |
| accepted_values_db 本格連携 | **C2** |
| 日中トリガー無人 PACKAGED | 後 |

### 3.3 人間に残る作業

- C／U2／U4／Da（GENERATED）まで既存どおり  
- C1 DRY_RUN結果の確認 → 本番実行  
- SC「商品スプレッドシート」UP（ZIPは任意・URL本線可）  
- 21-③  

---

## 4. 入出力契約

| 項目 | 契約 |
|------|------|
| テンプレ | PT別純正のみ。HPC／SEASONINGのマップ・指紋を混用しない。行5属性マップ・VBAを壊さない |
| 行構造（正） | HPC: 行7から実データ。SEASONING: **行7注記保持・行8から実データ**。[SEASONING列マップ](D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md) |
| 親子最小 | **1親＋N子**を一式で出力。子の必須欠落・URL欠落があれば**親一式を書かない** |
| URL | `Amazon MAIN URL` 優先。**空＝スキップ＋ログ。フォールバック禁止** |
| テンプレ指紋 | ヘッダー／属性行（行3〜5相当）のハッシュを成功時指紋と比較。不一致 → DRY_RUN警告・**本番停止** |
| 命名 | 03: 衝突しない一意名（日時／runId／subBatchId／親SKU） |
| DRY_RUN | 本番トグルと分離。DRY_RUNなしの本番禁止。成果物に対応表・欠落・除外親を含む |
| Drive | 06読取→編集コピー→03新規。到達手段（同期／API／手動）は実装承認で固定 |

---

## 5. 失敗・安全

1. **DRY_RUN必須**（本番前に1回以上）  
2. 本番は **03新規のみ**（06非破壊）  
3. 途中失敗: **そこまで書いたファイルは残す**＋ログ（人間が続き／破棄判断）  
4. **必須欠落・URL欠落・親子不整合**: **当該親SKU一式を出力から除外**し、他親は続行可  
5. **テンプレ指紋不一致**: 本番は実行しない（ファイルを作って列ずれを埋めない）  
6. 18320なしは T2再検証1SKU実績のみ。U5相当まで断言しない  

---

## 6. 三点レビュー

- 実施済: [D_MENU_C1_THREE_REVIEW_MAJORITY.md](D_MENU_C1_THREE_REVIEW_MAJORITY.md)  
- 総合: **条件付き YES** → 本ファイルへ反映済  

---

## 7. 検収（HUMAN_RUN骨格）

手順（詳細は `D_MENU_C1_HUMAN_RUN.md` を実装承認後に作成）:

1. DRY_RUNで欠落・対応表・除外親が読める  
2. テンプレ指紋一致を確認（不一致なら本番中止）  
3. 本番で03に新規 xlsm、06不変  
4. マスタ非書込・楽天／Yahoo非改変  
5. HPC 1親＋子で SC スプシUPが致命エラーなし（URL本線。18320は参考観測）  
6. 21-③  

チェックリスト（案）:

- [ ] DRY_RUNレポート可読  
- [ ] 指紋一致（または意図的な指紋更新手順を踏んだ）  
- [ ] 03新規・06不変  
- [ ] 親一式除外ログが欠落ケースで出る  
- [ ] SC致命エラーなし  
- [ ] 楽天／Yahoo／マスタ非改変  

---

## 8. 実装チケット

| ID | 内容 | 状態 |
|----|------|------|
| **C1-0** | 本要件＋方針ロック | **済** |
| **C1-3R** | 三点レビュー＋多数決反映 | **済** |
| **C1-1** | ローカル PACKAGED 生成（DRY_RUN＋本番＋指紋） | **済**（骨格） |
| **C1-1b** | 成功相当必須列（マスタ→HPC） | **実装済**（`master_csv` 併読・タックスはマスタ） |
| **C1-HR** | HUMAN_RUN | 生成OK・**SC再検は未送信SKUで** |
| **C1-GAS** | GAS薄い案内のみ | 任意・未 |
| **C1-1c** | FOOD系複合テンプレの `SEASONING`（七味唐辛子） | **実装済・機械DRY_RUN合格／実データ待ち** |
| **C1-ε** | FOOD／他PT・GAS直編集スパイク | バックログ |
| **C2** | accepted_values_db／深い純正差分 | バックログ |

---

## 9. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-30 | **C1-1c SEASONING**: 321列複合テンプレ、PT別マップ／指紋、行7注記保持、行8開始。機械DRY_RUN＋HPC回帰合格。 |
| 2026-07-26 | **C1-1b実装**: マスタCSV併読・必須列埋め・タックスはマスタ。次＝未送信SKUでSC。 |
| 2026-07-26 | 列マップ下書き追加。SCスモークは必須列不足で不合格（想定内）。次＝C1-1b。 |
| 2026-07-26 | 実装承認＋`tools/c1_hpc_packaged`。次＝HUMAN_RUN実機。 |
| 2026-07-26 | 三点＋社長決定反映。URL空=スキップ／親一式除外／指紋v1本番停止。次＝実装承認。 |
| 2026-07-26 | 初版。方針ロック反映。本線＝ローカル／Cursor。次＝3者。 |
