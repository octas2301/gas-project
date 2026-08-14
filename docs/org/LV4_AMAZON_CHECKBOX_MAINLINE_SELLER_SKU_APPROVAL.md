# Amazon レ点本線＋Amazon相乗りSKU列 — 運用切替承認パッケージ

**日付**: 2026-07-30  
**状態**: **3者レビュー・実装承認済／相乗り自己発 dry_run・prod 実機合格**（2026-07-30）。FBA・新規・同時・フルは実機確認待ち  
**目的**: 人間が全出品対象へ付ける `出品CK` を当面の承認意思として D から実行し、承認①は将来の AI 無人実行用ガードとして温存する  
**親**: [AI_ORG_CHARTER.md](AI_ORG_CHARTER.md)／[AI_APPROVAL_MATRIX.md](AI_APPROVAL_MATRIX.md)  
**関連**: [LV4_SPAPI_D_ENTRY_APPROVAL.md](LV4_SPAPI_D_ENTRY_APPROVAL.md)／[LV4_SPAPI_GAS_PUT_APPROVAL.md](LV4_SPAPI_GAS_PUT_APPROVAL.md)／[LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md](LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)  
**3者レビュー**: [THREE_REVIEW_RUNBOOK.md](THREE_REVIEW_RUNBOOK.md) に従い実施。[多数決メモ](LV4_AMAZON_CHECKBOX_MAINLINE_THREE_REVIEW_MAJORITY.md) は社長承認・正本反映済み  

---

## 0. 3者レビュー結果（2026-07-30）

- 3者とも **条件付き**。聖域は全員5/5
- 2/3以上の採用事項は社長承認済み
- 本書に、D旧ラジオからの置換、レ点互換、prodゲート、列追加影響、相乗り専用sellerSku例外、正式ヘッダ確定ゲートを反映
- 社長回答: **フル＋既存相乗りprodを許可**。開始前にprod確認を完了する
- 社長回答: 新ASIN型／旧JAN型とも、Amazon上に同一sellerSkuがあれば **PUT・登録で更新**、無ければ新規登録する。sellerSkuが異なる2経路は別SKUとして扱う

---

## 1. 背景と決定した考え方

承認①の本来目的は、**AI社員が候補を選び、そのまま勝手に出品することを防ぐガード**である。

当面は AI が出品対象を自動決定せず、**人間が出品する子SKUすべてに `出品CK`（レ点）を付ける**。このため Web 承認①を重ねると同じ人間判断が二重になり、運用負荷だけが増える。

当面の扱いを次のように変更する。

| 項目 | 当面 | 将来 AI 無人化時 |
|------|------|------------------|
| 出品対象の承認意思 | **人間の子SKUレ点** | 承認①（Web／バッチ）を再接続 |
| D 本線 | レ点行を実行 | 承認①済みバッチを実行 |
| 承認①メニュー・データ | **削除せず温存** | 再利用 |
| 承認②（在庫反映） | 本件対象外・別ゲート | 維持 |

本件は承認制の廃止ではなく、**人間レ点を当面の軽い承認①相当として扱う運用切替**である。

---

## 2. D を唯一の本線起点にする

### 2.1 UI

Amazon 出品方式はラジオではなく **複数選択可能なチェックボックス**とする。

```text
【コース】
○ フル（楽天 → Yahoo!）
○ 楽天のみ
○ Yahoo!のみ
○ Amazonのみ
○ フル → Amazon

「Amazon」を含む場合:

【Amazon出品方式】（両方同時選択可）
□ 新規カタログ
   レ点 → GENERATED → PACKAGED → SC

□ 既存カタログに相乗り
   レ点 → SP-API

「既存カタログに相乗り」の場合:
○ 相乗り自己発
○ 相乗りFBA
○ dry_run
○ prod
```

本承認後は、このチェックボックスUIが [LV4_SPAPI_D_ENTRY_APPROVAL.md](LV4_SPAPI_D_ENTRY_APPROVAL.md) の相互排他ラジオに代わる **D本線の正**となる。E／Z-21 の復旧用入口は残す。

`フル → Amazon` でも新規・相乗りを両方選択でき、相乗りprodも許可する。ただし **楽天／Yahooの開始前**に、Amazon prodの全トグル・対象件数・SKU例・送信在庫（既定0。マスタqty経路ON時はそのqty）を表示し、人間が確認ダイアログでOKする。取消時はフル全体を開始しない。

E／Z-21 はテスト・復旧用に残す。本線定着後の E 縮小は別承認。

### 2.2 レ点行の新規／相乗り振り分け

相乗りの発送区分（自己発／FBA）は **Dのラジオ**で選ぶ。X列は新規SKU式用のため、相乗り時に変更しない。

| 経路 | 対象 |
|------|------|
| 相乗りのみ | レ点付き子SKUすべて（X不問）。N列ASIN必須（後段） |
| 新規のみ | レ点付き子SKUすべて |
| 両方 | **同じレ点子SKUを新規と相乗りの両方へ出す**。新規=`子SKU`／相乗り=`Amazon相乗りSKU`。相乗りだけN列ASIN必須 |

`出品CK` は boolean `true` または文字列 `"TRUE"`。子SKU空の親レ点のみは除外する。実行前に件数・SKU例を表示する。

---

## 3. 新規カタログ（レ点本線）

- 対象: `出品CK=true` かつ子SKUあり（X不問。N列ASINがあっても新規から除外しない）
- 承認①キューは必須にしない
- **マスタ在庫>0でもスキップしない**（2026-07-30 承認）。同じレ点行から作るのは既存単品とは別のノーブランドセットカタログのため。マスタ在庫は読取のみ・非改変で、GENERATED の在庫列は `inventoryMode` に従い 0／1。Property `APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK=false` で旧 `SKIPPED_IN_STOCK` に戻す
- 承認①経路（21-①等）は従来どおり在庫>0を `SKIPPED_IN_STOCK` で除外。販売中上書き・補充は本件で行わない
- GENERATED → PACKAGED → SC の既存 Da 契約を維持
- Drive `02` MAIN ゲートは従来どおり必須
- 商品は期限管理シール等を含むセット販売であり、既存単品カタログ相乗りとは別商品
- `子SKU` は現行 JAN 入りSKUを維持（X列は新規SKU式用）
- 新規カタログ用のマッピング・バリエーション構造は本件で変更しない
- 相乗りと同時選択時も、新規は`子SKU`・相乗りは`Amazon相乗りSKU`で識別子を分離する

承認①済み新規経路（21-①等）は削除せず、Z／将来AI用に残す。

---

## 4. 既存カタログ相乗り（レ点本線）

### 4.1 対象

- `出品CK=true` かつ子SKUあり（X列の相乗り値は不要。発送区分はD選択）
- SP-API `LISTING_OFFER_ONLY`
- 自己発送は `fulfillment_channel_code=DEFAULT`＋**送信quantityは原則0**（承認①＝掲載）
- **例外（P0・承認②相当・2026-08-01）**: D で「マスタ在庫で出す」を選び、専用トグル（例: `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY`）が true のときのみ、自己発へマスタ由来 quantity を送ってよい。詳細は [LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md](LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md)／[MAJORITY](LV4_D_P0_THREE_REVIEW_MAJORITY.md)。マスタ列への書込は禁止のまま
- FBAは `fulfillment_channel_code=AMAZON_JP`。FBA在庫はAmazon管理のためquantityを送らない（P0でも変更しない）
- マスタ在庫>0でもスキップしない（承認①経路ではAmazonへは0。非0は上記承認②例外のみ）
- dry_run VALID 後にのみ prod 可
- 承認①済み経路（21-⑫⑬）は残すが、当面の D 本線では使わない

### 4.2 ASINの正

1. 子行 `ASINコード`（N列。人間が確定した相乗り先）
2. 子で空なら同一親の親行 `ASINコード`
3. 無ければ停止

`競合店ASINコード`（O列）および競合URL／DB列はリサーチ参考値であり、出品先ASINには使用しない。

JAN → Catalog Items API 自動検索は本件に含めない。

### 4.3 新規列

`▼商品マスタ(人間作業用)` の人間作業エリア（`子SKU`／`対象ASIN` 近傍）に次を追加する。

| ヘッダ | 入力者 | 用途 |
|--------|--------|------|
| `Amazon相乗りSKU`（NF列） | GAS | **自己発（MFN）専用** sellerSku |
| `Amazon相乗りSKU_FBA` | GAS | **FBA専用** sellerSku（[デュアル Phase1](LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)・2026-08-01） |

- 物理列番号は固定しない。ヘッダ名で解決
- 列が無ければ **当該系統**の相乗りを拒否し、追加手順を案内
- `子SKU` は変更しない（楽天／Yahoo／新規Amazonへの波及防止）
- 人間手入力を正にしない
- Dは1回1系統のまま。自己発実行はNFのみ／FBA実行は `_FBA` のみ更新（他系統列は触らない）
- 列追加前後に `[セット構成提案][列範囲チェック]` を確認し、親SKU／子SKU等がコピー範囲内であることを実機検収する

### 4.4 Amazon相乗りSKU生成

例:

```text
子SKU（新規側）:        lifec-4560151300832-48s11
ASIN:                  B07YND44VN
Amazon相乗りSKU:         lifec-B07YND44VN-48as11
```

規則:

1. 子SKU中央がすでに今回のASINなら **置換せずそのまま使う**
2. そうでなければ現行子SKU内のJAN等を **完全一致トークン**としてASINへ置換
3. 置換時のみ、D選択に従い `f1 → af1`／`s1 → as1`（既に `as`/`af` があれば維持）
4. 重複時末尾 `-2` 等は維持
5. 置換対象トークンが0件または複数件なら停止（中央ASIN済みは除く）
6. 生成後SKUは半角・Amazon許容長を検証
7. 既存 `Amazon相乗りSKU` があれば再生成せず再利用
8. 既存値内のASINと今回確定ASINが不一致なら停止（自動上書き禁止）

`LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md` の「出品者SKU＝子SKU」は **新規カタログ（B/M1）** の正とし、既存相乗り（A/M2）だけ `Amazon相乗りSKU` を使用する例外とする。

Amazon API／登録処理は sellerSku 単位の upsert とする。

- 旧JAN型子SKU: 新規カタログ経路。同一sellerSkuがAmazonにあれば更新、無ければ新規登録
- 新ASIN型Amazon相乗りSKU: 既存相乗り経路。同一sellerSkuがAmazonにあれば更新、無ければ新規登録
- 旧JAN型と新ASIN型は別sellerSkuのため、相互に上書きされない
- Amazon上に同一sellerSkuが存在することだけを理由に `-2` 等の別SKUを生成しない

子SKU式からN列参照は人間が削除済み。相乗りSKU生成時は子SKU内の `JANコード` または `オリジナルカタログ商品名` の完全一致トークンをASINへ置換する。

### 4.5 書き込みタイミング

- dry_run（上級・任意）が **status=VALID かつ issues=0** のSKUだけ **当該系統の列**へ保存
- **通常運用はprod直**。当該系統列が空なら **prodでSKUを生成してPUT**し、成功後に列へ保存（dry_run必須にしない）
- 保存済み値が正しい as/af なら再利用し、再生成しない
- 列の `s` 残存は prod／dry_run とも as/af へ正規化し、成功後に保存
- prod ACCEPTED は runId／SKU／ASIN とともにログへ記録
- prod失敗時に未保存なら列は空のまま。再実行で同じ規則で再生成可

これにより、通常運用では dry_run なしで新規＋相乗り同時出品できる。系統別の仮／確定列は追加しない（Phase1は2列で完結）。

---

## 5. 同時選択時の実行契約

1. Dで選んだ経路へ独立に載せる（**両方選択時は同じレ点行を新規と相乗りの両方へ**）
2. 新規件数・相乗り件数・SKU例を確認表示（識別子: 新規=`子SKU`／相乗り=`Amazon相乗りSKU`）
3. 相乗りはN列ASIN必須。両方選択時にASIN無し行は相乗りだけスキップし、新規は続行可
4. prod選択時は主トグル＋`ALLOW_PROD`＋確認ダイアログを、フルの他モール開始前に完了
5. 新規: GENERATED（`子SKU`）
6. 相乗り: 選択モード（dry_run／prod、`Amazon相乗りSKU`）
7. 結果を経路別 runId／成功／失敗で表示

部分成功はあり得る。自動ロールバックしない。

- 新規成功・相乗り失敗: 新規成果物を消さず、相乗りだけ再試行
- 相乗り成功・新規失敗: PUTを取り消さず、新規だけ再試行
- 成功SKUは同じ実行で再処理しない
- 後日の意図した再実行は同じsellerSkuを更新し、Amazon上に別sellerSkuを増やさない
- 新規は既存の状態管理、相乗りは `Amazon相乗りSKU`＋dry_run runId／ASIN／状態ログを冪等判定に使う

---

## 6. ガード・復元

### 6.1 Property

| キー | 既定 | 用途 |
|------|------|------|
| `APPROVAL_AMAZON_LV4_ENABLED` | false | 新規カタログGENERATED |
| `APPROVAL_AMAZON_SPAPI_PUT_ENABLED` | false | 相乗りSP-API |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD` | false | prod許可 |
| `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` | 10（未設定時） | 上限（1〜50） |
| `APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_0` | true | 在庫0 |

### 6.2 承認①を再接続する条件

次のいずれかを導入する前に承認①を再接続し、実機検収する。

- AI が `出品CK` を自動で付ける
- AI が D を無人実行する
- 日中トリガーへ Amazon 出品を載せる
- 人間のレ点確認なしに候補を出品対象へ昇格する

承認①関連のコード・メニュー・ApprovalQueue・batchId・docs は削除しない。

### 6.3 復元

1. `APPROVAL_AMAZON_LV4_ENABLED=false` と `APPROVAL_AMAZON_SPAPI_PUT_ENABLED=false`
2. Z の承認①候補作成 → Web承認 → 21-①／21-⑫⑬へ戻す
3. `Amazon相乗りSKU` 列は履歴として残し、削除・一括クリアしない
4. Git revert

---

## 7. 変更予定ファイル（実装は承認・3者反映後）

| 種別 | パス | 内容 |
|------|------|------|
| 改修 | `コード.js` | D階層UI、レ点新規の薄い入口、同時選択・振り分け・確認 |
| 改修 | `AmazonSpapiPut.js` | レ点相乗り、Amazon相乗りSKU生成・VALID後書込・prod再利用 |
| 改修 | `AmazonSpapiExport.js` | 必要なら item生成時のsellerSku差替えを共通化 |
| 更新 | `AI_ORG_CHARTER.md`／`AI_APPROVAL_MATRIX.md` | 当面レ点＝人間承認、AI無人化時は承認①再接続 |
| 更新 | `LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md` | B＝子SKU、A＝Amazon相乗りSKUの例外・upsert契約 |
| 更新 | 商品マスタ要件 | `Amazon相乗りSKU` 列 |
| 更新 | D／SP-API HUMAN_RUN、CURRENT_PHASE、HANDOVER、CHANGE_LEDGER | 手順・復元 |

**禁止**

- 楽天聖域、`Yahoo.js`、B統合Step境界の変更
- `子SKU` の上書き
- `Amazon相乗りSKU` の一括上書き／一括クリア
- ASIN不一致時の自動上書き
- 承認①関連機能の削除
- 承認②・在庫補充の同時実装
- コード実装前の社長承認

---

## 8. リスク

| リスク | 緩和 |
|--------|------|
| 将来AIがレ点を付けたのに承認①なしで出品 | AIレ点・無人化前の承認①再接続条件を正本化 |
| 新規セットと既存相乗りの誤分類 | 出荷方式値を正。ASIN有無で分類しない。未知値は全停止 |
| 販売中SKUを在庫0へ上書き | マスタ在庫>0は `SKIPPED_IN_STOCK`。補充・販売中更新は別承認 |
| dry_runとprodでsellerSkuが変わる | VALID後保存、prodは保存値必須 |
| ASIN変更で別SKUが増殖 | 保存済SKUとASIN不一致は停止・人間判断 |
| マスタ列追加による既存ロジック波及 | ヘッダ解決・`子SKU`非改変・対象列だけ単セル/列書込 |
| 列挿入でテンプレコピー範囲がずれる | `[セット構成提案][列範囲チェック]` と重要列数式を追加前後で実機確認 |
| 同時実行の部分成功 | 経路別runId／成功分を消さない／失敗側のみ再試行 |
| 承認マトリクスとの矛盾 | 3者レビュー＋社長承認後に正本へ反映 |

---

## 9. 合格条件

- [ ] 新規／相乗りの既存Property falseで各経路が拒否され、承認①経路は動く
- [ ] 子SKUレ点のみ対象、親レ点のみは除外
- [ ] `出品CK` の boolean `true`／文字列 `"TRUE"` 両対応
- [ ] 在庫>0でも相乗りは停止せず送信在庫=0（マスタ在庫は書き換えない）
- [ ] D選択で新規／相乗りへ独立に載せる（両方時は同行列を両方へ）。X列は新規SKU式用であり経路分割に使わない
- [ ] 新規と相乗りを同時選択でき、確認画面に件数・SKU例（両方時は同時出品の注記）
- [ ] 新規は既存 Da の成果物・`02`ゲートを維持
- [x] 相乗りdry_run VALID/issues=0後だけ `Amazon相乗りSKU` を保存（自己発1SKU・任意経路）
- [x] prodは保存値でACCEPTED（自己発1SKU）
- [ ] **prod直**: SKU列空でも生成→PUT→成功後保存（2026-08-02実装。実機待ち）
- [x] prodは主トグル＋ALLOW_PROD＋確認OKが必須。フル時も他モール開始前に確認（開始前確認UIは実装済・フル経路の実機は未）
- [x] 同一sellerSkuは更新、未登録sellerSkuは新規登録。旧JAN型と新ASIN型を混同しない（自己発1SKUで確認）
- [ ] ASIN不一致で停止し自動上書きしない
- [ ] 承認①メニュー・21-①／21-⑫⑬が従来どおり
- [ ] 列追加前後の列範囲チェックで親SKU／子SKU等が範囲内
- [x] 出荷方式ヘッダ・許容値とtarget_val元ヘッダ・優先順をdocsに確定
- [ ] トグルをfalseへ戻す
- [x] docs／ログ／revert手順（本合格記録）

---

## 10. 3者レビューで特に確認する事項

1. 人間レ点を当面の承認①相当とすることが憲章の「不可逆操作は人が最終ゲート」と矛盾しないか
2. AI無人化時の承認①再接続条件が十分か
3. `Amazon相乗りSKU` の生成・保存・ASIN不一致停止で識別子増殖を防げるか
4. 新規／相乗り同時選択の部分成功契約が安全か
5. 実装前未決の2ヘッダ（出荷方式・target_val元）を未決のままコード着手してよいか（既定: 不可）

---

## 11. 社長承認欄

- [x] **3者多数決案を承認し、正本へ反映してよい**（2026-07-30）
- [x] **実装を承認する**（2026-07-30。列名=`Amazon相乗りSKU`、NF列）
- [ ] 条件付き／却下（条件: ）

**コード着手条件**: 充足済み。コード実装済、次は clasp push → HUMAN_RUN。
