# D×Amazon P0 — E吸収＋在庫選択＋prod既定（承認パッケージ）

**日付**: 2026-08-01  
**状態**: **実装済**（方針ロック反映・`clasp push` 後に HUMAN_RUN で検収）  
**ロードマップ**: [AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md)  
**多数決**: [LV4_D_P0_THREE_REVIEW_MAJORITY.md](LV4_D_P0_THREE_REVIEW_MAJORITY.md)  
**親**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)／[LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)／[AI_APPROVAL_MATRIX.md](AI_APPROVAL_MATRIX.md)  
**三者レビュー**: **実施済・条件付き**（在庫マスタ送信は承認②相当として親正本同時改定）  
**RUNBOOK**: [THREE_REVIEW_RUNBOOK.md](THREE_REVIEW_RUNBOOK.md)

---

## 1. 目的

人間クリックを減らしつつ、レ点本線の安全弁を維持する。

1. **E. Amazon出品コースを D に吸収**（段階実行は D 内。Z・旧メニューは逃げ道として残置）  
2. **送信在庫**: D で選択。**既定＝0（承認①相当）**／任意＝**マスタ在庫に基づく数量（承認②相当）**  
3. **dry_run**: **既定＝prod**。dry_run は折りたたみ（上級）。Z-21 の dry_run は残してよい  

マスタ出品用「在庫数」列への **書込は禁止のまま**（楽天／Yahoo共有列の副作用防止）。送る数量の話に限る。

---

## 2. スコープ

### 2.1 含む（三点の対象・採用条件反映）

| 項目 | 仕様 |
|------|------|
| 在庫UI | `在庫0で出す`（既定・承認①）／`マスタ在庫で出す`（承認②相当）。送信qty＝マスタ**「在庫数」生値** |
| マスタ在庫選択時 | 開始前確認に **件数・SKU例・送信qty（合計および内訳）** を必須。空・負・非数は**送信停止** |
| 承認ゲート | マスタqty経路は **専用 Script Property**（例: `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY`、**既定 false**）が true のときのみ。既存 `ALLOW_PROD` 等と併用 |
| 相乗り PUT・自己発 | 既定0時は quantity=0。マスタ選択＋トグルON時のみ quantity＝マスタ「在庫数」生値。**FBAは quantity 非送信** |
| 新規 GENERATED／バルク | 既定 ZERO。マスタ選択時はバルク在庫列へ「在庫数」生値（マスタ非書込）。**開始前確認に新規qty内訳必須**（2026-08-01補強） |
| FORCE_QTY_0 | マスタqty経路が有効な実行でのみ緩和。それ以外は既存 FORCE_QTY_0 を優先 |
| prod既定 | D で prod を既定。dry_run は折りたたみ。**フル→Amazon も含める** |
| prodゲート（削らない） | **ALLOW_PROD＋開始前確認**必須。**コードから ALLOW_PROD を自動ONしない**。相乗りは **dry_run VALID先行／`Amazon相乗りSKU`空ならprod停止** を存続 |
| E吸収 | E-0〜E-5 相当を D から呼べる。トップの E は**互換サブへ移動**（消去しない） |
| Z残置 | 21・C-Amazon・復旧用は削除しない |
| 初回1SKU必須 | **実装ゲートにしない**（社長 2026-08-01。テスト即時確認前提） |

### 2.2 含まない

- P4a／P4b、P1、P2／P3／Dcループ  
- 承認マトリクスのリトライ回数変更（**最大2回のまま**）  
- 楽天聖域・Yahoo出品本体の改変  
- B統合 Step 境界の改変・GENERATED本体のD内再実装（薄いファサード維持）  

### 2.3 親正本の同時改定（(B)必須）

| 文書 | 改定内容 |
|------|----------|
| レ点本線承認包 | 「quantity常に0」に **承認②＝マスタqty経路の例外**を追記 |
| 承認マトリクス | Amazon D のマスタqty送信＝**承認②（補充）**行を明記 |
| 本包 | 本節 |

---

## 3. リスクと緩和

| リスク | 緩和 |
|--------|------|
| レ点ミス＋qty>0で実売 | 既定0。マスタ経路は専用トグル既定false＋確認ダイアログ。社長はテスト即時確認で1SKU必須免除 |
| prod誤反映 | ALLOW_PROD＋開始前確認。自動ON禁止。dry_run先行存続 |
| 承認②迂回の見た目 | マスタqty＝承認②とマトリクス／本包に明記 |
| マスタ列書込 | 明示禁止 |
| FBA誤qty | FBAはquantity非送信 |

---

## 4. 実装見積り（実装承認後）

| ファイル（見込み） | 概要 |
|--------------------|------|
| `コード.js` | Dダイアログ（在庫・prod/dry_run）、E吸収、Property一時ONは**ガード系以外** |
| `AmazonApprovalExport.js`／`AmazonSpapiPut.js` | qty分岐・FORCE_QTY_0優先・FBA非qty |
| docs | HUMAN_RUN／本承認状態／CHANGE_LEDGER |

着手前に **変更予定ファイル一覧／概要／リスク** を再提示し社長**実装承認**を得る。

---

## 5. 検収（実装後）

- [ ] 既定在庫0で従来動作  
- [ ] マスタqtyはトグルON時のみ送信され、マスタ列不変、FBAはqtyなし  
- [ ] prod既定でも確認キャンセルで停止。ALLOW_PROD自動ONなし  
- [ ] dry_run VALID／相乗りSKU空でprod停止  
- [ ] E互換またはDから同等到達、Z残置  
- [ ] ログ標準・Property戻し  

---

## 6. 残確認（社長・MAJORITY §4）

**2026-08-01 確定済**（追加確認なし・方針ロック）:

1. 送信qty＝マスタ **`在庫数` 生値**（`FLOOR(仕入÷セット)` は本P0非適用）  
2. Eトップ＝**互換サブへ移動**  
3. フル→Amazon＝prod既定に**含める**  
4. qty不正＝**送信停止**  

次ゲート: **HUMAN_RUN 検収**（既定0／マスタqtyトグル／prod確認／E互換）。

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。 |
| 2026-08-01 | 三点条件付き。社長(B)・1SKU必須免除。MAJORITY反映・承認②整合。 |
| 2026-08-01 | **補強**: 新規のみMASTER時も開始前にqty内訳確認（相乗りと同水準）。D UI文言で新規＋相乗り共通を明示。 |
| 2026-08-01 | 追加社長回答: 在庫数生値／E互換サブ／フル含める／qty不正は停止。方針ロック。 |
| 2026-08-01 | **実装**: D在庫UI・prod既定・ALLOW_MASTER_QTY・E→D補助/Z互換。 |
