# 仕入れ検討① — 人間作業（複製書き・専用GAS）

**日付**: 2026-08-15  
**正**: [B_PURCHASE_RESEARCH_NOW.md](B_PURCHASE_RESEARCH_NOW.md)／[V1](B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md)／[TASK_STRUCTURE](B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)

出品用 `.clasp.json` は **触らない**。`clasp push` は親フォルダで走らせない。

---

## 0. いまの成果（ローカル・貼付用）

`tools/purchase_research_path3/out/ishihara_gate_paste.csv`  
石原143件。門: 通過65／落ち78（価格30・順位40・冷凍8）。閾値は仮（2000円・15万）。

---

## 1. 書き分け

| 置き場 | タブ | 中身 |
|--------|------|------|
| 競合DB `1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs` | `Keepaフル` | Keepa属性の正（W1から①が書く。90日。csv[]なし） |
| 競合DB 同上 | `Keepaフル索引` | ASIN→行の自動索引（書込時にASIN列だけで再構築。人手メンテなし） |
| 競合DB 同上 | `メーカーマスタ` | 第2クエリ語。語があれば再採取しない |
| 競合DB 同上 | `メーカーサマリ` | メーカー1行。実態4分類。要確認は列 **実態ck**（Amazon検索URL）で人確認 |
| 調査複製 `1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE` | `①候補` ほか | 門の転記。Catalog生・モールヒットは置かない |

調査複製は `contact@octas2301.com` に編集者共有が必要。

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\purchase_research_path3
python apply_research_ss.py
```

---

## 2. 専用 GAS（clone 済）

scriptId: `1oRfQzsECsGbuft8S7ksuxfKwAgq-AgFdSse41lE2lJVP-rPb4qx4yNRr`  
ローカル: `purchase_research/`（親 `.clasp.json` は出品のまま）

**push は人間が PowerShell で行う。Agent は clasp push しない。指示するときは必ず:**

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp pull
clasp logs
clasp push
```

**中身**: 外注マクロは残す。①は `PurchaseResearch.js`（**レ点の店／メーカー**・モリタ `/query`・門・メーカー洗い既定cap20・D1/D2）。

**初回 push 前に** GAS エディタ → プロジェクトの設定 → Script Properties:

Keepa 契約は **1本**。出品と同じキー文字列を、①プロジェクトの Property にも入れる（プロジェクトが違うので自動では入らない）。トークン残は出品 A と①で共用（**日次枠ではない**。**正は応答の refillRate**。実測 2026-08-20 は **5/分・バケツ約300**）。洗い完走トリガー待ちは **次スライス分**（`min(残り, cap)×単価` − `max(0,tokensLeft)`）の回復分＋1分（上限=300回復＋1）。**残り全体×単価は使わない**（特大で常に61分になるため。2026-08-29）。単価は **取得プロファイル**（メニュー別ではない）。`offers=20` 付き `/product` は既定 **6**（実測≈5.5–6。Property `KEEPA_PRODUCT_TOKEN_COST`）。スライス応答があれば **consumed/件数** を優先。メーカー洗いは出品と同じ SP-API LWA（読取のみ）。洗いスライス既定 **30件**（残トークンが足りれば同実行内で継続。Property `PR_MAKER_WASH_CAP`）。

| Key | 値 |
|-----|-----|
| `KEEPA_API_KEY` | 既存の Keepa キー（出品と同じ） |
| `SPAPI_LWA_CLIENT_ID` | 出品／smoke と同じ（秘密。Git禁止） |
| `SPAPI_LWA_CLIENT_SECRET` | 同上 |
| `SPAPI_REFRESH_TOKEN` | 同上 |
| `SPAPI_ENDPOINT` | 省略可（既定 FE） |
| `SPAPI_MARKETPLACE_ID` | 省略可（既定 JP） |
| `PR_MAKER_WASH_ENABLED` | 省略=ON。切戻しは `false` |
| `PR_MAKER_WASH_CAP` | 省略=**20**。上限50 |
| `KEEPA_REFILL_RATE` | 省略=**5**（実測フォールバック）。**応答 `refillRate` 優先**。トリガー／待ち分数の推定用 |
| `KEEPA_TOKEN_BUCKET_MAX` | 省略=**300**。洗い完走トリガー待ちの上限（MAX回復＋1分） |
| `PR_KEEPA_FULL_INDEX_ENABLED` | 省略=ON。Keepaフル索引（ASIN→行）自動。D1は必要ASINだけ読む。切戻し `false`＝旧全件読込 |
| `KEEPA_PRODUCT_TOKEN_COST` | 省略=**6**（`product_offers20`＝stats90+offers20）。実測≈5.5–6。洗い待ち・afford。応答の consumed/件数があればそちら優先 |
| `PR_WASH_QUEUE_COUNT_MAP` | 省略=ON。洗いキュー件数を「メーカー→件数」表で1回読む。切戻し `false`＝社ごとに全行読む旧動作（選定が数分〜十数分に戻る） |
| `PR_CYCLE_POOL_REUSE` | 省略=ON。選定プール（小中/不明/大/特大）を1実行1回だけ数える。切戻し `false`＝呼ばれるたび数え直す旧動作 |
| `PR_WASH_QUEUE_SNAPSHOT` | 省略=ON。洗い対象メーカーのキュー行を1実行1回だけ読む（列挙の追記・重複整理・未済取得・Keepa済・残数を共有）。切戻し `false`＝呼ばれるたび全行読む旧動作（1スライス20分超に戻る） |
| `PR_WASH_ARM_FULL_CAP` | 省略=ON。**実残トークンが満タンなら待ちは2分まで**（overlap 再予約など）。ckpt は同一実行の安全ネットなので対象外。切戻し `false` |
| `PR_KEEPA_FULL_BLANK_SCAN_BATCH` | 省略=ON。Keepaフルの空行掃除を「ASIN列1回読み→候補行だけ全列」に。切戻し `false`＝最終行から1行ずつ読む旧動作（**Keepa取得のあるスライスが毎回+22分**になる） |
| `PR_KEEPA_TWO_STAGE_GATE` | 省略=ON。洗いの Keepa GET を2段に。**1段目 offers なし＝1トークン/件**で門を判定し、**通過だけ offers=20**（実測 通過3%→1件あたり約5→約1.2トークン）。切戻し `false`＝従来の一発 offers=20 |
| `KEEPA_TWO_STAGE_TOKEN_COST` | 省略=**2**（`product_two_stage` の安全側既定）。2段ゲート時の afford・待ち計算の単価。応答の consumed/件数があればそちら優先 |
| `KEEPA_STATS90_TOKEN_COST` | 省略=**1**（offers なし `/product` の単価） |
| `PR_D1_ENRICH_BLANK_BB` | 省略=ON。**門通過なのに Keepaフルの `BuyBoxセラー`・`BuyBox_FBA` が両方空**の行だけ offers=20 で1回太らせる（2段ゲートで薄い行が残るため）。切戻し `false`＝「行があれば補強しない」旧動作 |
| `PR_D1_ENRICH_CAP` | 省略=**20**。上の補強 GET の1回上限（BuyBox が元々無い商品での空回り抑制） |
| `PR_LIST_DIFF_USE_FORMULA` | 省略=ON。品番リストの差分判定で**数式セルは数式として比べる**。切戻し `false`＝表示結果で比べる旧動作（**画像URLがある行を洗いスライスごとに削除→行5再挿入**し、人の手入力列が消える） |
| `PR_KEEPA_FULL_WRITE_ENABLED` | 省略=ON。切戻しは `false` |
| `PR_REGATE_FROM_FULL_ENABLED` | 省略=ON。99「Keepaフル再門」。切戻し `false` |
| `PR_REGATE_ONLY_BLANK` | 省略=true。再門は門結果空白のみ。`false`＝全行を新門で上書き |
| `PR_LIST_APPEND_ENABLED` | 省略=ON。D1切戻しは `false` |
| `PR_CYCLE_MID_D1_ENABLED` | 省略=ON。洗い途中の品番リスト転記切戻しは `false` |
| `PR_CYCLE_LIST_CHECK_ENABLED` | 省略=ON。週次転記チェック（列・D1回収・メール・トリガー）切戻しは `false` |
| `PR_CYCLE_LIST_CHECK_DAYS` | 省略=**7**。週次の `取得日時` 窓（日） |
| `PR_CYCLE_AUTO_CK_DAILY_CAP` | 省略=**8**。自動レ点の1日上限 |
| `PR_ALIGHT_DAILY_TARGET` | 省略=**5**。1日の店フル完走目標 |
| `PR_ALIGHT_DAILY_MAX` | 省略=**10**。1日の店フル上限 |
| `PR_ALIGHT_NO_MAKER_ROUNDS_MAX` | 省略=**10**。新メーカー0のときの同日追加ラウンド上限 |
| `PR_ALIGHT_SHOP_CAP` | 省略=**1**（1実行1店）。旧「浅い3店」は廃止方向 |
| `PR_ALIGHT_PRODUCT_CAP` | 省略=**20**。**1スライス**の件数（店全体ではない） |
| `PR_LIST_APPEND_CAP` | 省略=100。1回の追記上限（最大500） |
| `PR_LIST_FILTER_ENABLED` | 省略=ON。D2切戻しは `false` |
| `PR_LIST_FILTER_CAP` | 省略=**1000**。**未判定 ASIN** の1回上限（最大2000） |
| `PR_D2_ONLY_BLANK` | 省略=ON。見積候補／見送り済み行は D2 で触らない。全行再判定は `false` |
| `PR_CYCLE_D2_OVERFLOW` | **自動**。D2 上限超で C4 停止した parent（\| 区切り）。メニュー28または D2 解消でクリア |
| `PR_CK_SELLER_ENABLED` | 省略=ON。レ点の店切戻しは `false` |
| `PR_CK_MAKER_ENABLED` | 省略=ON。レ点のメーカー切戻しは `false` |
| `PR_CK_BATCH_CAP` | 省略=3。1回に進める店／メーカー数（最大10） |
| `PR_AUTO_MAKER_ENABLED` | 省略=ON。「未調査メーカー自動」切戻しは `false` |
| `PR_AUTO_MAKER_CAP` | 省略=1。1回の自動選定件数（最大5） |
| `PR_AUTO_MAKER_REWALK_DAYS` | 省略=90。最終調査日がこれより新しければ自動対象外 |
| `PR_WASH_QUEUE_RETENTION_DAYS` | 省略=**90**（3か月）。Keepa済キューの期限切れ削除。`0`=オフ |
| `PR_SELLER_REWALK_DAYS` | 省略=180（**仮6か月**。店の取扱見直し。Keepa 90日とは別） |
| `PR_SELLER_REWALK_OVERWRITE` | 省略=ON。再調査で機械欄を上書き。切戻しは `false`（空のときだけ書く） |
| `PR_SELLER_NAME_GET_ENABLED` | 省略=ON。店名の Keepa GET 切戻しは `false` |
| `PR_SELLER_NAME_GET_CAP` | 省略=40。1実行の店名GET上限（最大80） |
| `PR_SELLER_ACCUM_ENABLED` | 省略=ON。C0 セラー貯め切戻しは `false` |
| `PR_SELLER_ACCUM_LOOP` | 省略=ON。429 時に同一メニュー内で待機再試行（100件 or 未済0まで） |
| `PR_SELLER_ACCUM_429_WAIT_MS` | 429 再試行前待機（省略=90000） |
| `PR_SELLER_ACCUM_429_RETRY_MAX` | 1メニューあたり429再試行上限（省略=5） |
| `PR_SELLER_ACCUM_CAP` | 省略=**100**（最大100）。1回の新規店上限 |
| `PR_SELLER_ACCUM_SLEEP_MS` | 省略=**1100**。getItemOffers 間隔（ms） |
| `PR_SELLER_ACCUM_AFTER_ASIN` | **自動**。429／25分切れ時の再開 ASIN（成功完走で削除） |
| `PR_SELLER_ACCUM_BATCH` | 省略=OFF。`true` で getItemOffersBatch（20 ASIN/回） |
| `PR_ALIGHT_ENABLED` | 省略=ON。C1+C2 切戻しは `false` |
| `PR_ALIGHT_SHOP_CAP` | 省略=**1**（最大10）。**1実行あたりの店数**（1店フルが正） |
| `PR_ALIGHT_PRODUCT_CAP` | 省略=**20**（最大30）。**1スライス**の商品数（店全体ではない） |
| `PR_JITTA_RESOLVE_ENABLED` | 省略=ON。実態AI全体切戻しは `false` |
| `PR_JITTA_RESOLVE_CAP` | 省略=**100**（最大100）。手動メニュー用 |
| `PR_JITTA_HB_LANE_ENABLED` | 省略=ON。心拍で実態AI（店と分離）。切戻し `false`＝旧 A-light 末尾AI |
| `PR_JITTA_HB_CAP` | 省略=**8**（最大20）。心拍1回の実態AI件数 |
| `PR_JITTA_BACKLOG_GATE` | 省略=**30**。未調査がこれ以上なら自動レ点しない |
| `PR_JITTA_AFTER_ALIGHT` | 省略=レーンONならOFF。明示 `true` で A-light 末尾もAI |
| `PR_JITTA_OWNER` | 省略=`hb`（①心拍）。競合DBスクリプト運用時は **`comp`**（①の実態OFF）。切戻しは削除または `hb`。[競合実態AI](B_COMPETITOR_JITTA_REQUIREMENTS.md)。**法人格ゆれ**: 同一キーへミラー／heal（`heal=`）。要 clasp（競合＋①） |

コンテナバインドのため SS ID Property は不要（`getActiveSpreadsheet`）。

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

**メーカー洗い完走:** 正本 [FULL](B_MAKER_WASH_FULL_REQUIREMENTS.md)。メニュー「列挙ドライ」「続き」「ループ」「**トリガー開始／停止**」「**洗いキュー 期限切れ削除**」。**トリガー開始＝予約のみ**（本処理は時間トリガー）。**arm 時は Prop を先に ON→トークン回復見積で予約**（A→D 含む）。①ログに `トリガー予約`／detail `trigger_armed afterMin=`。単一実行ロックあり。Keepaフルは **upsert（1ASIN1行）**。キュー Keepa済は既定 **90日**で削除可（`PR_WASH_QUEUE_RETENTION_DAYS`、0=オフ。完走時も掃除）。**完走時にメーカーサマリの調査CKを外す**（トリガー完走含む）。切戻しは「トリガー停止」または Prop `PR_WASH_FULL_TRIGGER_ENABLED=false`。①に `GEMINI_API_KEY`。`appsscript.json` に `script.scriptapp` 必須。push 後は必要なら再承認。

**門・プロファイル:** `①プロファイル` に **出品者数下限／上限**（初期 **1／20**）、**順位下限**（初期 **1**＝欠測落ち）、**価格上限**（初期 **100000**・カテゴリ別変更可）。列は左寄せ（価格上下限隣接→順位段階＋下限隣接→出品者数）。門で範囲外なら落ち。Amazon本体は **Amazon直販**（`availabilityAmazon`）。洗いの Keepa は **offers=20** で Keepaフルへ書く。D1はフルに行があれば再GETしない（欠落ASINのみ取得）。

**店規模（CYCLE §4.1bis・GAS実装済・要 clasp）:** 最左に **`調査カテゴリー`**（チェック）＋ **`AmazonルートID`**。正の初期行は **`食品・飲料・お酒`**（競合DBと同名）＋ ID **`57239051`**（＝同ルート）。カテゴリ名を「食品」だけにしない。誤行（カテゴリ＝ID／別名「食品」）は削除。飲料・お酒の絞りはルート変更せず門・後工程。

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

`python tools/competitor_store/dry_maker_wash_full_p1.py --maker 石原水産`

**D1→D2（推奨）:** メニュー「D1→D2 門通過→品番リスト→見積候補」。D1のあとD2まで連続。単独の D1／D2 は切り分け・再実行用。

**Keepaフル再門:** 99「Keepaフル再門（①候補・GETなし）」。競合DB `Keepaフル` の生JSONだけで門を再評価（`/product` 消費なし）。既定は **門結果空白行のみ**。全行を新門で上書きする場合は Property `PR_REGATE_ONLY_BLANK=false`。

**D1 品番リスト（GAS）:** メニュー「D1 門通過→品番リスト」。**新規**は行5から挿入。**既存 ASIN** は Keepa 機械列に差分があれば上書きして行5へ移動。差分なしは触らない。利益の式は上書きしない。差分判定は**数式セルを数式として比べる**（`PR_LIST_DIFF_USE_FORMULA`。画像URLの `=HYPERLINK` が毎回差分になるのを防ぐ）。

**D2 機械絞込（GAS）:** メニュー「D2 機械絞込（見積候補）」。**未判定 ASIN のみ**処理（`PR_D2_ONLY_BLANK` 既定 ON）。空行・判定済み行は cap に数えない。`見積候補`／`見送り理由`。合格: `leftover=0`（未判定が cap 超で残った件数）。

**D2 上限超（C4 停止）:** C4 内 D2 で leftover>0 → `d2_overflow` で心拍へ・**メール（本文に復帰手順）**・`PR_CYCLE_D2_OVERFLOW` 登録。C4 recover／新規 C4 予約は止める。**復帰**: 99→6「D2だけ」で leftover=0 → 99→28「D2 backlog解消→C4再開」（または 4「完走連鎖」）。D2 解消時に needs が1社なら C4 自動再開。

### 2.1 A–D を人間が回す手順（データが溜まるか見る）

**使うブック**: 調査複製（原本は開かない・書かない）。  
**メニュー**: スプレッドシート上部 **「自動リサーチ」**（表は A / B / C / 99 のみ）  
**先に push**（このフォルダだけ）:

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

毎回、タブ `①ログ` の最新行を見る。失敗なら止めてログを残す。成功時は画面右下の **toast**（数秒で消える）。詳細は①ログ。失敗だけダイアログ。

**実行時間**: メニュー実行は出品Bと同じく **課金30分前提・25分で止めて leftover**。画面上の帯「起動時間の最大値を超えました」＋詳細／表示しないは、**シートを開いたときの onOpen（常に30秒）**が多い。①ログに新しい FAILED が無ければメニューは動いていない。**表示しない**で閉じ、メニュー「自動リサーチ」があれば作業してよい。品番リストの式が多いと帯が出やすい。Keepa の店GETは **レ点の店だけ**（まとめでは呼ばない）。

**レ点**: 競合DBのタブ **セラー**。A列 `調査CK`。行が青になってから実行。完了するとレ点は外れる。**同じ店は1行**（余分な行は消す）。店名・内訳・食品比率は今回の結果で上書きする。メモは追記。店名を人が直した行はメモに「手修正」と書く。店全体の商品数は相手が返さないと空のまま。

**①候補**: 同じ ASIN は1行。別の店で当たったら `sellerId`（参考）と `店名` を `|` で別セルへ足す。店名は台帳にあれば写す。**ストアASIN数はレ点の店**で `/seller storefront=0` から書く。まとめメニューは Keepa しない。

**日常（人が覚えること）**: 普段は **A-1 だけ**。止まったら **A-2**。店／メーカーだけ手でやるなら **B**。見積は後で **C**。部品・停止・試験は **99**。

| 順 | 人が押す | 見る場所 | 溜まったと言える目安 |
|----|----------|----------|----------------------|
| **A-1** | **A. 全自動 → A-1. 開始／再開（セラー・メーカー・品番リサーチ全自動）** | ①ログ `A-1`／`a1_park`／`heartbeat`／`C4` | 実行中を一時停止→心拍が先頭から再選定。キュー残し。ロック中は約1分後予約 |
| **A-2** | **A-2. 全診断（時計・沈黙・再武装）** | ①ログ `A-2 全診断` | 時計・inflight・沈黙。沈黙なら再武装 |
| **B-1** | **B. 個別調査 → 1. レ点の店だけ調べる** | セラー・メーカーサマリ・①ログ | 旧 C1+C2（A-light＋ヒット）。門・①候補なし |
| **B-2** | **2. レ点のメーカーだけ洗う** | メーカーサマリ・洗いキュー | レ点メーカーのみ |
| **B-3** | **3. 実態だけAI判定（未調査）** | メーカーマスタ・実態列 | 未キャッシュのみ |
| **C** | **C. 見積・発注 →（準備中）** | — | E〜J 未。toast のみ |
| **99** | **99. 運用・切り分け**（停止・診断単体・D単独・C0/C3/C4・メンテ）＋**テスト** | ①ログ | 日常は開かない。旧 C0〜C4・S・モリタ等 |

内部（メニューに出さない／99 のみ）: C0 セラー貯め・C3+T・C4・S 単品名は **A-1 の心拍**が回す。切り分け時だけ 99。

AGF・モリタ固定は **99→テスト**。店の個別はレ点。1回あたり cap=3（Property `PR_CK_BATCH_CAP`）。再巡回の6か月は仮。

E（見積の可／不可・提案文）は **次スライス半自動**（CYCLE §13）。今の心拍は D2＋S まで。メニューはまだ出さない。メモは FLOW §5.3.1。

**M3 人間（W3 GAS）:** 99「Keepaフル再門（①候補・GETなし）」で門結果空白を埋める。全行を新門で上書きは `PR_REGATE_ONLY_BLANK=false`。モリタ門は店 `/query` 分のみ消費。

---

## 10. CYCLE 日常本線（C5 確定・2026-08-24）

メーカーレーンの**日常操作は C3+T 1本**。C3 単体は1スライス確認用。

### 10.1 流れ

```text
C3+T（1回）
  → 選定（洗い途中優先）→ 洗い1スライス（armTrigger）→ 途中D1（門通過分）
  → 未完なら P3 トリガー自動予約（**継続する**。C3単体は予約しない）
full_trigger（無人・25分×N ＋ トークン待ち）
  → Keepa→門→①候補 → **途中D1**（wrote>0・未完時）
  → 完走時 C4 自動（**約1分後の別実行**）: D1→D2→心拍（次社は心拍選定。旧 `C4_arm` 直次社は使わない）
full_trigger / heartbeat
  → 次社の洗い（同上）
```

**停滞対策（2026-08-25）**: 1スライス成功でも未完なら再予約。`lock_busy` 時も再予約。発火時 `RUNNING/fire` を①ログへ。PROP=OFFでも続き対象があれば `trigger_rescue_on`。

**強制終了耐性（2026-08-25）**:
- 実行予算 **20分**（プラットフォーム約30分殺しの前に正規終了）
- ループ開始時＋**各 Keepa スライス完了時**にチェックポイント再予約（`*_ckpt` / `洗いスライス`）
  - **効率ループ ON（2026-09-03・既定）**: 実行中 ckpt は**張らない**。再開は `finishIncomplete` の Keepa `waitMin` のみ。満タンなら `afterMin≤2`（①ログ `eff_loop_cap=1`）
  - 切戻し `PR_WASH_EFF_LOOP_ENABLED=false` → 旧 ckpt0（GAS残＋2≈21分）
- `lock_busy` の待ちは **ロック残り秒＋1分**（2分連打しない）
- **C4 完走連鎖は洗い完走と別 GAS 実行**（`C4トリガー予約` → 約1分後 `C4 trigger` fire）。長い洗い完走直後の 30 分壁で D2 前切断しない（2026-08-27）
- **発火ライフサイクル**（2026-08-27）: `inflight` を expect より先に立てる。①ログ `full_trigger phase`（after_clear / pick_begin / pick_end）。二重発火は `overlap_skip`。例外は `outer FAILED`＋`watchdog_rearm`。30分殺し後は inflight が残り、次発火または診断で `inflight_stale`
- **心拍全体の生存**（2026-08-27／28）: 洗い発火だけでなく C4・心拍分岐・pause も対象。15分ごとの `heartbeat liveness`。沈黙・**予定超過（Keepa最大回復＋5分）**なら `hard_reset`→`liveness_reset`。
  - **実行中印だけの腐りは 35分基準**（2026-09-08・`PR_CYCLE_INFLIGHT_STALE_FAST` 省略=ON）: GAS は約30分で殺すので、`inflight` が35分を超えたら生きている可能性はない。トークン回復基準の66分を待つと**沈黙死を拾うまで70分空く**（実測 09/08 00:02 発火→`phase=pick_begin` のまま沈黙→01:12 `hard_reset`＝Keepa作業ゼロが179分）。予約トリガーが残っていると `reason=wash_trigger` で生存扱いになり、この inflight 判定だけが唯一の検知口。**予定時刻超過（`expect_overdue`）は従来どおり66分**。`liveness_reset` に `deadPhase=` `overdueMin=` が付く。切戻し `PR_CYCLE_INFLIGHT_STALE_FAST=false`心拍の終わりは `heartbeat end`（washN/c4N/inflight）。C4 の D失敗は停止せず心拍へ。**A-1 だけで再開**（腐った実行印・予約も消す）
- **洗い→リスト転記は別実行**（維持）: 完走後 C4 は約1分後の別 GAS。同一実行に混ぜない
- **Keepa 残量**: 予約前 `/token`。`calcLeft` 欠落／不足でも `liveLeft` で次スライス可なら `wait_shortened`（長待ち短縮）。**`*_ckpt` / `overlap_wait` は短縮しない**（2026-09-01。`PR_WASH_CKPT_NO_SHORTEN` 未設定=ON）。効率ループ ON 時は実行中 ckpt 自体を出さない
- **overlap**（2026-09-01／**整流 2026-09-03**）: age&lt;8分は `overlap_skip`＋短再試行（効率ONは **2分**）。**age≥8分は即 takeover**（`inflight_overlap_takeover`）して本処理。切戻し `PR_WASH_EFF_LOOP_ENABLED=false` で旧 stale 寄せ長待ち
- **WAIT_TOKEN 前 inflight clear**（2026-09-03）: `finishIncomplete` 予約直前に clear（次発火の overlap 連鎖防止）
- **店レーン譲り**（2026-09-03）: 洗い続き無し／大後回しで心拍へ入る直前に inflight clear＋`shop_lane_yield`（約3分）＋busy。A-light 1店完了で busy 解除＋`hb_after_alight` 1分。busy 中の洗い発火は `shop_hb_busy_skip`。切戻し `PR_SHOP_LANE_CLEAR_INFLIGHT=false`
- **trigger_missed 抑制**（2026-09-01）: 非 stale inflight 生存中は expect 超過でも `trigger_missed` にしない（誤判定防止）

**トリガー切り分け（①ログ）**:
| 行の目印 | 意味 |
|----------|------|
| `Keepa token` / `liveLeft=` / `wait_shortened` | 予約直前の `/token`（消費0）。計算残欠落でも実残があれば待ち短縮 |
| `トリガー予約` / `expectedAt=` | 予約成功・発火予定（日本時間）。`liveLeft=` 併記 |
| `洗いスライス` / `checkpoint_armed` | スライス完了＋安全ネット予約 |
| `full_trigger … RUNNING` / `fire` | **発火した**（lagMin=遅れ分） |
| `洗いトリガー解除` / `trigger_cleared` | 誰が消したか（complete / menu_stop / no_target 等） |
| `洗いトリガー診断` / `trigger_missed` | 予定時刻＋3分を過ぎても未発火 |
| `C4トリガー予約` / `c4_trigger_armed` | 完走後 C4 を約1分後に別実行予約 |
| `C4 trigger` / `fire` | C4 連鎖が独立実行で発火 |
| `full_trigger phase` | 発火後の段階（after_clear / pick_begin / pick_end）。**fire の次がこれ無しなら選定前で死んだ** |
| `full_trigger overlap` / `overlap_skip` | 先実行が inflight 中の二重発火。age&lt;8は本処理しない＋短再試行。**age≥8は `inflight_overlap_takeover` で本処理**（2026-09-03） |
| `outer FAILED` / `watchdog_rearm` | 選定等の例外。約2分後に再武装 |
| `inflight_stale` | 実行中印が35分超。先実行が catch 未到達で死んだ |
| `inflight_overlap_takeover` | 効率ループ: overlap で age≥8 の即奪取 |
| `eff_loop_cap=1` | 満タンなのに長待ちを2分に抑えた |
| `shop_lane_yield` / `shop_hb_busy_skip` | 店レーン譲り（inflight 捨て＋短予約）。busy 中は洗い本処理しない |
| `効率合否` / `PASS`/`WARN`/`FAIL` | **効率ループ合否**（2026-09-03）。`event=` slice/arm/overlap/stale · `cause=` · `fixHint=` · `gapMin=`/`afterMin=`/`liveLeft=`。比較・対策立案用 |
| `heartbeat liveness` / `liveness_reset` / `hard_reset` | 15分番犬。沈黙または予定超過（最大回復＋5分）でリセット再武装。A-1 も同じ |
| `heartbeat end` | 心拍1回の終わり。`washN` `c4N` `inflight` |
| `title0_skip` / `title0_retry` | タイトル含有0が3回→保留＋心拍。`PR_WASH_TITLE0_SKIP_SET` 永久除外。未満は再試行 |
| `done_no_candidates` / `false_complete_rewash` / `empty_research` | 偽完了→再洗い（自動1回）／空調査は保留。手動は **99 → 偽完了を列挙中に戻す** |
| `shop_after_sm` / `xl_week_cap` / `brand_hint` | 中小完走→店挟み（案B）。大特大週間N=2。title0→ブランド行追加 |
| `C4_d_fail_heartbeat` / `hb_no_maker` | D失敗・メーカー空でも次時計 |
| `hint=prop_on_but_no_trigger` 等 | 未発火時の推定 |
| メニュー **洗いトリガー診断** | 99→2。いつでも現状を①ログへ |
| メニュー **心拍生存診断** | 99→3 |
| メニュー **A-1** | 実行中を `a1_park`＋abort世代→孤児ロック破棄→**店日次カウンタ（auto_ck／A-light／ラウンド）リセット**→心拍／C4。**フロー先頭から再選定**（キュー残し）。週間大特大枠は残す。ロック詰まっても1分再試行。**小中=0 なら店ラウンド≤10を尽くすまで大・特大を選ばない**（CYCLE §4.4） |
| メニュー **A-2 全診断** | 時計＋沈黙。再武装あり（日常の切り分け） |

**品番リスト**: Keepa スライスごとに門通過分を途中転記（§6.4bis・①ログ `途中D1`）。完走後の C4 は D1→D2→心拍。切戻し `PR_CYCLE_MID_D1_ENABLED=false`。

**週次転記チェック**（§6.4ter）: ①候補 `転記チェック`（未／済／漏れ）。日曜トリガー（A-1／再開で武装）。漏れは D1 してメール。手動は 99「転記チェック（週次・今すぐ）」。切戻し `PR_CYCLE_LIST_CHECK_ENABLED=false`。

**Catalog先行**（2026-08-29）: サマリ `Catalog見積` 未了の洗えるメーカーを Keepa 新規より先に列挙。洗い本体も `catalog_first_self`／`catalog_first_other` でゲート。Catalog 429/503 は社内メール（回復目安分・クールダウン60分）。保留（empty_research／title0 等）は選定プール外。

**大→特大・週間枠・店挟み**（2026-08-30）: 洗える大が残る間は特大の続き／キューを後回し（`defer_xl_large_remain`）。大特大の続きでも週間 N を消費／満杯なら後回し。**小中=0 のとき**は店ラウンド（レ点→A-light、必要時のみ C0、最大10）を尽くすまで大・特大を選定・続き開始しない（pick→defer 禁止）。小中完走後は店1回。大特大完走後は門通過→セラー追加→セラー調査。枠・大残り／店ラウンド中は Keepa中進捗を消さない（`keep_progress`）。

**取得日時**: Keepaフル等は **日本時間**（`…+09:00`）。①ログ `at` も東京。
### 10.2 実機合格の目安（アサヒ飲料・2026-08-24）

| 段 | runId / 時刻 | 合格ログ・状態 |
|----|--------------|----------------|
| C3+T | `pr_20260824_013324` | `trigger_armed parent=アサヒ飲料` |
| 完走 | 同 runId 1:42 | `LOOP COMPLETE`・`left=0`・サマリ `洗い状態=完了` |
| C4 | 同 runId 1:43 | `d_ok arm=永谷園`・D2 `leftover=0` |
| D1 | 1:42 | `pass=37 append=37`（品番リストへ反映） |

合否確認:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/common_ss_view/dump.py dump --menu 仕入れ検討① --intent maker
```

**効率ループ合否（2026-09-03）**: ①ログに `効率合否` 行（`PASS`/`WARN`/`FAIL`）。`cause`・`fixHint` で原因比較。集計:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/common_ss_view/dump.py dump --menu 仕入れ検討① --intent keepa --no-limit
python tools/common_ss_view/eff_judge_report.py --hours 24
```

**P-eff-1/2 反映後の確認**（clasp push 後・数時間）:
- `full_trigger_ckpt0` の長待ち（afterMin≈20）が**出ない**
- 満タン予約に `eff_loop_cap=1` または `afterMin=2`
- overlap は age≥8 で `inflight_overlap_takeover`（stale 37分待ちに行かない）
- `効率合否` の FAIL 原因 `ckpt_full_bucket_long_wait` が減る

| event | PASS 目安 | FAIL の典型 cause → fixHint |
|-------|-----------|------------------------------|
| slice | gapMin≤25 | gap_hole → ckpt_overlap_stale |
| arm | 満タンで afterMin短。**ckpt は afterMin≤予算+2+3（既定25分）なら PASS**（`ckpt_safety_net`） | arm_full_bucket_long_wait → deficit0_afterMin_le_2 |

**ckpt の合否は予算基準（2026-09-07・`PR_EFF_CKPT_BUDGET_RULE` 省略=ON）**: `*_ckpt` は `残り予算+2分`で貼るのが正（同一実行と重ならないための安全ネット。満タン2分上限も ckpt を除外している）。よって 15〜22分は正常で、旧判定の `ckpt_long_net` WARN と `ckpt_full_bucket_long_wait` FAIL は**誤検知**だった（2026-09-07 実測7件すべて。実害＝空き時間はゼロ）。新判定は `ckptMax=25` を超え**かつ満タン**のときだけ `ckpt_long_despite_tokens` WARN（待ち計算式の異常）。トークン枯れで長いのは PASS。切戻し `PR_EFF_CKPT_BUDGET_RULE=false` で旧固定15分判定。
| overlap | age&lt;8分 | overlap_inflight_stuck → immediate_takeover |
| stale | （出たら常に FAIL） | inflight_stale_takeover → clear_inflight_before_arm |
### 10.3 Property（切戻し）

| Property | 既定 | 意味 |
|----------|------|------|
| `PR_CYCLE_C4_ENABLED` | ON | 完走後 C4 |
| `PR_WASH_FULL_TRIGGER_ENABLED` | OFF→C3+TでON | P3／心拍 |
| `PR_MAKER_WASH_FULL_ENABLED` | ON | 洗い完走本線 |
| `PR_CYCLE_BAND_SELECT` | ON（実装後） | OFF で旧 ≤500→ヒット降順に戻す |
| `PR_CYCLE_HEARTBEAT_PAUSED` | — | 他メニューで一時停止 |
| `PR_CYCLE_DAILY_MAIL_ENABLED` | ON（実装後） | 日次生存メール切戻し |
| `PR_EFF_JUDGE_ENABLED` | ON | ①ログ「効率合否」。切戻し `false` |
| `PR_EFF_CKPT_BUDGET_RULE` | ON | **ckpt 合否を予算基準に**（≤予算+2+3=25分は PASS）。切戻し `false`（旧 固定15分で WARN） |
| `PR_CYCLE_INFLIGHT_STALE_FAST` | ON | **番犬の実行中印を35分基準に**（30分壁準拠。沈黙死の検知を70分→40分に）。切戻し `false`（旧66分） |
| `PR_MAKER_BAND_MAP` | ON | **規模帯を「メーカー→帯」表1回読みで引く**（空振り選定 2〜11.9分→数秒）。切戻し `false`（1社ずつ3回読み） |
| `PR_WASH_EFF_LOOP_ENABLED` | ON | **P-eff-1/2**: ckpt実行中廃止・満タン≤2・inflight clear・overlap即takeover。切戻し `false`（旧ckpt0） |
| `PR_SHOP_LANE_CLEAR_INFLIGHT` | ON | **店レーン譲り**: yield で inflight clear＋busy。切戻し `false` |
| `PR_JITTA_HB_LANE_ENABLED` | ON | **実態AI分離**: 心拍 `jitta_resolve`。切戻し `false`（旧 A-light 末尾） |
| `PR_JITTA_OWNER` | hb | **`comp`**=競合DB実態スクリプトに委譲。切戻し `hb` |

停止: **99 → 1. 自動を止める**。すぐ戻す: **A-1**（ロック空なら即心拍／忙しければ約1分後）。切り分け: **A-2**。C4 だけ止める: `PR_CYCLE_C4_ENABLED=false`。

### 10.4 CYCLE チェックリスト（運用）

- [x] メーカー **1社完走→次社**（アサヒ→永谷園。2026-08-24）
- [x] **小→中**優先。小中=0 なら店ラウンド≤10のあと大・特大（正本 §4.4／§7。GAS実装済・要 clasp）
- [x] 他メニューで停止 → 手動再開 or **01:00**
- [x] **日次生存メール**（停止中も送る・本文に完走／門通過）
- [ ] 人は **要確認** ＋ E〜J（次）

### 10.5 規模帯・心拍（2026-08-26・心拍完成形）

正本 [CYCLE §6.3〜§7・§13](B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)。

仮想: `python tools/competitor_store/dry_cycle_c3_select.py`（11/11）／ `python tools/competitor_store/dry_cycle_heartbeat.py`（19/19）

GAS: `prRunWashFullFromTrigger` が続き無し時に `prDispatchHeartbeat_`（列挙／小中洗い／**実態AI**／店フル／自動レ点／C0／大）。**列挙（enumOnly）後は `Catalog見積`/`規模帯` を残し Keepa中にしない → 心拍再arm**。小中が残る間／**規模不明が残る間**は大・特大のキュー続きを後回し。**小中=0かつ未調査ありは店より先に実態AI**（§4.5bis）。店は **1店フル**（特大店≥501はフルしない）。日次レ点既定8・店目標5・上限10・新メーカー0で最大10ラウンド。未調査≥30で自動レ点抑制。切戻し `PR_CYCLE_HEARTBEAT_ENABLED=false`／実態AIレーン `PR_JITTA_HB_LANE_ENABLED=false`。

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

push 後は必要なら再承認（`gmail.send`）。再開: **A-1**。沈黙の切り分けは **A-2**（または 99 の診断単体）。完全停止は `PR_CYCLE_HEARTBEAT_ENABLED=false`（洗いトリガー停止だけだと、心拍ONなら番犬が再武装する）。

仮想（心拍が止まらない全体）: `python tools/competitor_store/dry_cycle_heartbeat_liveness.py`（16/16）

---

## 3. やらない

- 出品 `コード.js` への①追加
- 親フォルダでの clasp push
- 原本調査 SS への書き
- 公式サイトの商品照合（ASIN突合）。カテゴリ語の採取だけは可（NOW）
- 問屋・メーカーへの **自動送信**（E〜J でも聖域）
