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
| `PR_KEEPA_FULL_ASIN_MAP` | 省略=ON。**案A**: Keepaフルのマップ／flat は要求 ASIN だけ読む（索引→ずれ時は ASIN列走査）。全件 `getValues` 禁止。切戻し `false`＝旧全件一括（サイズ上限再発あり） |
| `PR_KEEPA_FULL_ASIN_READ_CHUNK` | 省略=**10**。案Aの連続行読取チャンク（最大50） |
| `KEEPA_PRODUCT_TOKEN_COST` | 省略=**6**（`product_offers20`＝stats90+offers20）。実測≈5.5–6。洗い待ち・afford。応答の consumed/件数があればそちら優先 |
| `PR_WASH_QUEUE_COUNT_MAP` | 省略=ON。洗いキュー件数を「メーカー→件数」表で1回読む。切戻し `false`＝社ごとに全行読む旧動作（選定が数分〜十数分に戻る） |
| `PR_CYCLE_POOL_REUSE` | 省略=ON。選定プール（小中/不明/大/特大）を1実行1回だけ数える。切戻し `false`＝呼ばれるたび数え直す旧動作 |
| `PR_WASH_QUEUE_SNAPSHOT` | 省略=ON。洗い対象メーカーのキュー行を1実行1回だけ読む（列挙の追記・重複整理・未済取得・Keepa済・残数を共有）。切戻し `false`＝呼ばれるたび全行読む旧動作（1スライス20分超に戻る） |
| `PR_WASH_ARM_FULL_CAP` | 省略=ON。**実残トークンが満タンなら待ちは2分まで**（overlap 再予約など）。ckpt は同一実行の安全ネットなので対象外。切戻し `false` |
| `PR_KEEPA_FULL_BLANK_SCAN_BATCH` | 省略=ON。Keepaフルの空行掃除を「ASIN列1回読み→候補行だけ全列」に。切戻し `false`＝最終行から1行ずつ読む旧動作（**Keepa取得のあるスライスが毎回+22分**になる） |
| `PR_KEEPA_FULL_INSERT_CHUNK` | 省略=**5**。Keepaフル新規行の `setValues` 分割件数（生JSONが重いため）。**0＝旧一括**。2026-10-17「セルの範囲を小さく」対策 |
| `PR_KEEPA_FULL_INDEX_CHUNK` | 省略=**2000**。Keepaフル索引の書込分割行数。**0＝旧一括** |
| `PR_KEEPA_FULL_BLANK_READ_CHUNK` | 省略=**40**。空行掃除で全列を読む最大連続行。**0＝連続塊ごと一括** |
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
| `PR_CYCLE_AUTO_CK_DAILY_CAP` | 省略=**0**（1日上限なし）。旧動作は `8` |
| `PR_CYCLE_AUTO_CK_PER_WAKE` | 省略=**1**。1回に付ける店数 |
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
| `PR_C0_SKIP_HTTP` | 省略=ON。新品オファー200以外は済にして次へ。①ログ `C0 skip_http`。切戻し `false` |
| （心拍）洗える小中ありの C0 | **しない**。レ点／名前取りが無ければ残洗い。①ログ `restock_skip_c0`／`restock_wash` |
| `PR_WASH_QUEUE_XL_PENDING_MAKERS_MAX` | 省略=**10**。大・特大の未済をキューに残す社数。`0`でオフ。小中は対象外 |
| `PR_ALIGHT_ENABLED` | 省略=ON。C1+C2 切戻しは `false` |
| `PR_ALIGHT_SHOP_CAP` | 省略=**1**（最大10）。**1実行あたりの店数**（1店フルが正） |
| `PR_ALIGHT_PRODUCT_CAP` | 省略=**20**（最大30）。**1スライス**の商品数（店全体ではない） |
| `PR_JITTA_RESOLVE_ENABLED` | 省略=ON。実態AI全体切戻しは `false` |
| `PR_JITTA_RESOLVE_CAP` | 省略=**100**（最大100）。手動メニュー用 |
| `PR_JITTA_HB_LANE_ENABLED` | 省略=ON。心拍で実態AI（店と分離）。切戻し `false`＝旧 A-light 末尾AI |
| `PR_JITTA_HB_CAP` | 省略=**8**（最大20）。心拍1回の実態AI件数 |
| `PR_JITTA_BACKLOG_GATE` | 省略=**30**。実態AI前の未調査がこれ以上なら自動レ点しない。要確認＋信頼度ありは数えない |
| `PR_JITTA_AFTER_ALIGHT` | 省略=レーンONならOFF。明示 `true` で A-light 末尾もAI |
| `PR_JITTA_OWNER` | 省略=`hb`（①心拍）。競合DBスクリプト運用時は **`comp`**（①の実態OFF）。切戻しは削除または `hb`。[競合実態AI](B_COMPETITOR_JITTA_REQUIREMENTS.md)。**法人格ゆれ**: 同一キーへミラー／heal（`heal=`）。要 clasp（競合＋①） |
| `PR_HIT_RECOUNT_OWNER` | 省略=`comp`（競合GASがセラーヒット／小規模出現）。切戻し **`hb`**=①が名前取りのあと数える |

コンテナバインドのため SS ID Property は不要（`getActiveSpreadsheet`）。

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

当たり直しの競合時計（①と混ぜない）:

```powershell
cd C:\Users\takuy\Desktop\gas-project\competitor_jitta
clasp push
```

競合メニュー **実態AI（競合）→ 5. ヒット時計 開始**（1回）。

**メーカー洗い完走:** 正本 [FULL](B_MAKER_WASH_FULL_REQUIREMENTS.md)。メニュー「列挙ドライ」「続き」「ループ」「**トリガー開始／停止**」「**洗いキュー 期限切れ削除**」。**トリガー開始＝予約のみ**（本処理は時間トリガー）。**arm 時は Prop を先に ON→トークン回復見積で予約**（A→D 含む）。①ログに `トリガー予約`／detail `trigger_armed afterMin=`。単一実行ロックあり。Keepaフルは **upsert（1ASIN1行）**。読取は **案A（`PR_KEEPA_FULL_ASIN_MAP`）**＝要求ASINだけ（全件一括は切戻し専用）。キュー Keepa済は既定 **90日**で削除可（`PR_WASH_QUEUE_RETENTION_DAYS`、0=オフ。完走時も掃除）。**完走時にメーカーサマリの調査CKを外す**（トリガー完走含む）。切戻しは「トリガー停止」または Prop `PR_WASH_FULL_TRIGGER_ENABLED=false`。①に `GEMINI_API_KEY`。`appsscript.json` に `script.scriptapp` 必須。push 後は必要なら再承認。

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
| **B-2** | **2. レ点のメーカーだけ洗う** | メーカーサマリ・洗いキュー | レ点メーカーのみ（実態未調査は洗わない） |
| **B-3** | **3. 実態だけAI判定（未調査）** | メーカーマスタ・実態列 | 未キャッシュのみ（レ点優先ではない） |
| **B-4** | **4. レ点メーカーを実態→洗い（自動一時停止）** | ①ログ `B4`／メーカーサマリ | **推奨**: 自動を一時停止→レ点だけ実態→洗い→飲み干しで自動再開。切れ目は B4 専用トリガー。切戻し `PR_B4_ENABLED=false`。要① clasp |
| **C** | **C. 見積・発注 →（準備中）** | — | E〜J 未。toast のみ |
| **99** | **99. 運用・切り分け**（停止・診断単体・D単独・C0/C3/C4・メンテ）＋**テスト** | ①ログ | 日常は開かない。旧 C0〜C4・S・モリタ等 |

内部（メニューに出さない／99 のみ）: C0 セラー貯め・C3+T・C4・S 単品名は **A-1 の心拍**が回す。切り分け時だけ 99。

AGF・モリタ固定は **99→テスト**。店の個別はレ点。1回あたり cap=3（Property `PR_CK_BATCH_CAP`）。再巡回の6か月は仮。

E〜H は **要件改訂**（FLOW §5.3.2）。見積候補だけ出す。見送りと、単品商品名の要確認は出さない。届いた見積はメールの添付から取る。送信は人。メール取得のコードの前は3者。メニュー C はまだ準備中。

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
- **店レーン譲り**（2026-09-03／20）: 洗い続き無し／大後回しで心拍へ入る直前に inflight clear＋`shop_lane_yield`。**名前取り続きでは busy を残さない**（次の起床が分岐に入る）。実行中の二重発火だけ busy で弾く。切戻しは Git revert。`PR_SHOP_LANE_CLEAR_INFLIGHT=false` は譲り全体の旧動作
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
| `enum_batch n=8/8 … stop=` | **まとめ列挙の結果**（2026-09-08）。`stop=sm_found=（社名）`＝小中が出て打ち切り／`stop=budget`＝予算残5分で中断／`stop=list_end`＝対象切れかcap到達。`skip=` は例外で飛ばした社数。帯の内訳（`小=1 特大=4` 等）も並ぶ |
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
| `PR_CYCLE_ENUM_BATCH_N` | **8** | **1実行でまとめて規模帯を確定する社数**（未判定244社の消化 約28時間→約7時間）。小中1社／予算残5分で打ち切り。切戻し `1`（心拍1回=1社） |
| `PR_CYCLE_ENUM_BATCH_STOP_QUEUE` | ON | **まとめ列挙の打ち切りを「キューが実際に増えた社」に**（中身ゼロの小で止まらない。実測 31回中29回が1社目で打ち切られ 1.5社/回だった）。ログ `stop=queue_found=`。切戻し `false`（旧 `stop=sm_found=`） |
| `PR_MAKER_NEEDLE_LEGAL` | ON | **社名の法人格ゆれを Catalog 照合語に足す**（`大平水産(株)`→`大平水産`／`大平水産株式会社`）。0件時は法人格なしで1回だけ再検索。ログ `needle_variants` / `legal_retry`。切戻し `false`（旧 表示名だけ） |
| `PR_CYCLE_TITLE0_POOL_SKIP` | ON | **タイトル含有0を1回出した社を選定プールから外す**（中身ゼロの社が unk に居座り特大の洗いを止めるのを防ぐ）。ログ `pool_title0_skip`。切戻し `false`（旧 3回連続で保留） |
| `PR_KEEPA_MISS_DONE` | ON | **Keepa が返さない ASIN を見切って Keepa済にする**（同じ ASIN を4分ごとに取り直す無限ループを止める。実測1社で1時間20分・27回）。ログ `keepa_miss_done`・①ログ `miss=`。切戻し `false` |
| `PR_KEEPA_MISS_DONE_AFTER` | 2 | 上記の見切り回数。1回目は据え置き・2回目で落とす。取れたら回数はリセット |
| `PR_CAND_PASS_BY_QUEUE` | ON | **門通過を自社キューの ASIN でも数える**（社名が包含関係の社に合格を奪われて「通過0」に見えるのを防ぐ。415社中53組が該当）。ログ `cand_pass_by_queue`。切戻し `false`（旧 メーカー列一致のみ） |
| `PR_FALSE_COMPLETE_PENDING_ONLY` | ON | **偽完了救済で Keepa中 に戻すのは未処理>0のときだけ**（旧は Keepa済も含む全件を見て、未処理0の社を Keepa中 に戻していた）。全件洗い済みで候補0は保留。ログ `empty_research queue=… all_done`。切戻し `false` |
| `PR_WASH_PHANTOM_FIX` | ON | **「Keepa中なのに未処理0」の社を連続2回で自動解消**（洗えない続き相手が心拍の列挙レーンを永久に塞ぐのを止める。実測25時間・549回）。ログ `phantom_keepa` / `phantom_keepa_resolved`。切戻し `false` |
| `PR_EFF_STALL_GUARD` | ON | **切片が60分出ていなければ 効率合否 を FAIL**（空転中も PASS が出続けていた穴を塞ぐ）。ログ `cause=no_slice_stall`・`stallMin=`。切戻し `false` |
| `PR_WASH_AUTO_RECOVER` | ON | **洗い自動リカバリー**（実態 skip 後の次予約／ピック除外／stall 番犬・FLOW §5.6／CYCLE §6.11）。ログ `auto_recover *`。切戻し `false` |
| `PR_CYCLE_ENUM_STOCK_N` | 3 | **洗える社を3社ためてから洗いへ譲る**（社の切替22分の主因・§10.3quin）。`1` で1社ずつ譲る旧動作 |
| `PR_WASH_QUEUE_PENDING_MAP` | ON | **未処理件数を1回読みの表に**（選定の社ごと全読み26,238行を廃止・§10.3quin）。切戻し `false` |
| `PR_CATALOG_RANK_PROBE` | ON | **列挙で Amazon の順位有無を記録**（門の落ちの33.8%は判定材料なし・§10.3sex）。`false` で `salesRanks` を取らない旧動作 |
| `PR_CATALOG_RANK_CUT` | **`true`** | 順位なしをキューに入れない。2026-09-13 検証済（予測と実際のズレ 中央値3pt・**門通過457件のうち順位なしは0件**）。切戻しは未設定に戻す |
| `PR_CYCLE_SHOP_AFTER_SM_ONCE` | ON | **完走後の店レーンを1回だけにする**（1社が大・特大102社を止める空転・§10.3sept）。切戻し `false` |
| `PR_SHOP_BEFORE_LARGE_CACHE` | ON | **店ラウンド判定の引数なし呼びを実行内1回にする**（pick が9分になり A-light が実行枠切れで死んでいた・§10.3dec）。切戻し `false`＝毎回外部シートを読む旧動作 |
| `PR_ALIGHT_SHOP_ROUNDS_UNTIL_CAP` | ON | **店ラウンドを日次上限まで回す**（旧は新メーカーが1社出たら店を閉じて大へ行っていた・§10.3undec）。切戻し `false`＝新メーカー数で閉じる旧動作 |
| `PR_SHOP_WHILE_TOKEN` | ON | **小中が0のとき liveLeft>予約 なら日次10店を超えて店を続ける**。切戻し `false`＝日次10で打ち切る旧動作 |
| `PR_CYCLE_ONE_SHOP_YIELD` | ON | **洗いキュー未済を続きにしない**。店を洗いより先にはしない。切戻し `false`＝キュー続きの旧動作 |
| `PR_SHOP_TOKEN_RESERVE` | 60 | 余力なしの上限（これ以下はKeepa作業しない） |
| `PR_TOKEN_AMPLE` | 200 | 余力ありの下限。中は61–199。店フルはありのときだけ |
| `PR_SHOP_TOKEN_GATE` | ON | **余力なしでは名取りを始めない**。200以上だけ店1軒。完了店の抽出メーカーをサマリに足す。切戻し `false` |
| `PR_SELLER_BAND_S_MAX` | 300 | **店の小**（SKU見積）。メーカー洗いの小≤100とは別。切戻し `100` |
| `PR_SELLER_BAND_M_MAX` | 600 | **店の中**。切戻し `300` |
| `PR_SELLER_BAND_L_MAX` | 900 | **店の大**。切戻し `500` |
| `PR_QUERY_TOTAL_HOLD` | 未設定＝大+1（901） | **これ以上は特大・規模測定のみ（名取りしない）**。切戻し `501` |
| `PR_SURVEY_ROOT_GATE` | ON | **門・D1・D2で調査カテゴリON以外を落とす**。加えてサマリ `Amazonルートカテゴリ` がONと無交差なら洗い選定・続き・本体も止め調査対象外。切戻し `false` |
| `PR_B4_ENABLED` | ON | **B4 レ点メーカー実態→洗い**（自動一時停止→専用トリガーで飲み干し→自動再開）。切戻し `false` |
| `PR_ALIGHT_XL_EXTRACT` | 未設定＝OFF | **特大の浅い名取り**。`true` で旧 page0＋最大20件 |
| `PR_SM_RESTOCK_BELOW` | 3 | **洗える小中がこれ未満なら店で補充**（残り2）。`1` で0になってから店へ行く旧動作 |
| `PR_JITTA_RETRY_NON_HIGH` | ON | **信頼度 medium／low の 要確認・未調査 も実態AIの対象に**（28社が行き止まりで9日間 jitta が動いていなかった・§10.3duodec）。切戻し `false`＝旧の「信頼度が空だけ」判定 |
| `PR_JITTA_RETRY_EXHAUST` | ON | **再調査しても 要確認 に戻る社を1回で打ち切る**（マスタ `実態再採取しない` を立てる。毎日 Gemini を呼び直さない）。429・タイムアウトでは立てない。切戻し `false`＝打ち切らない |
| `PR_JITTA_HB_DAILY_MAX` | 8 | **安全弁**。実態AIレーンに入る回数を1日8回まで（× `PR_JITTA_HB_CAP` 8社＝64社ぶん）。本体の空転止めは上の打ち切りなので、これはフラグ書き込みが失敗したときの保険 |
| `PR_WASH_ST_SPLIT` | ON（未設定=ON） | **洗い状態の終わりを3区分**（完了／調査対象外／要修復。保留廃止・§10.3tredec）。切戻し `false`＝完了＋通過0を保留に戻す旧動作 |
| `PR_WASH_ST_SPLIT_MIGRATED` | 空→`Y` | 保留行の1回移行が済んだ印。消すと再移行する。人がシートを手で直さない |
| `PR_WASH_QUEUE_PURGE_DONE` | ON | **完了／保留の社の Keepa済行を完走時に削除**（キュー75,931行の35%が残り続ける・§10.3non）。切戻し `false`＝日数条件だけの旧動作 |
| `PR_WASH_QUEUE_PURGE_BLOCK_MAX` | 2000 | これ以下はブロック削除、超えたら書き戻し方式。初回26,027行は書き戻しになる |
| `PR_CYCLE_DAILY_MAIL_TO` | **設定済**（例: `contact@octas2301.com`・2026-09-13 確認） | 日次生存・Catalog上限・週次転記の宛先。4か所とも先にこの Property を読む。**空のときだけ** `Session.getActiveUser()` に落ちる。`oauthScopes` に `userinfo.email` が無いので空だと例外（実測 2026-09-10 20:32／09-11 04:10 等）。届かないときは `①ログ` の `日次メール`／`週次転記メール`／`Catalog上限メール` が `DONE sent=` か `FAILED` かで見る |
| `PR_WASH_ST_SYNC` | ON | **書いた洗い状態を実行内にも反映**（無駄な `catalog_first_other` 譲りを消す・§10.3quat）。切戻し `false` |
| `PR_HB_YIELD_ARM_NO_SHORTEN` | ON | **心拍譲りの予約を5分・短縮対象外**（心拍実測4分。`shop_busy` 空振りを消す・§10.3quat）。切戻し `false`（3分＋短縮あり） |
| `PR_WASH_EFF_LOOP_ENABLED` | **OFF** | **P-eff-1/2**: ckpt実行中廃止・満タン≤2・inflight clear・**overlap即takeover(8分)**。2026-09-07 に ckpt を戻すため `false` 済み。**副作用で即takeoverも消えている**（CYCLE §6.4sept）。満タン上限は `PR_WASH_ARM_FULL_CAP` で切り離し済み |
| `PR_SHOP_LANE_CLEAR_INFLIGHT` | ON | **店レーン譲り**: yield で inflight clear＋busy。切戻し `false` |
| `PR_JITTA_HB_LANE_ENABLED` | ON | **実態AI分離**: 心拍 `jitta_resolve`。切戻し `false`（旧 A-light 末尾） |
| `PR_JITTA_OWNER` | hb | **`comp`**=競合DB実態スクリプトに委譲。切戻し `hb` |
| `PR_HIT_RECOUNT_OWNER` | comp | **省略も委譲**。①が数える切戻しは **`hb`** |

停止: **99 → 1. 自動を止める**。すぐ戻す: **A-1**（ロック空なら即心拍／忙しければ約1分後）。切り分け: **A-2**。C4 だけ止める: `PR_CYCLE_C4_ENABLED=false`。

#### 10.3bis 4対策の同時検証（2026-09-09）

①②③④は**別々のトグルと別々のログ目印**を持つので、同時に入れても後から切り分けできる。疑わしい1つだけ `false` にすれば、その対策だけ旧動作に戻る。

| # | 対策 | Property | ログの目印 | 合否に見る数字 |
|---|---|---|---|---|
| ① | 法人格ゆれを照合語に足す | `PR_MAKER_NEEDLE_LEGAL` | `needle_variants` / `legal_retry` | `title0_retry` の件数（0 に近づく） |
| ② | 空振り社を選定から外す | `PR_CYCLE_TITLE0_POOL_SKIP` | `pool_title0_skip n=` | 心拍 pool の `unk=` |
| ③ | 取れない ASIN を見切る | `PR_KEEPA_MISS_DONE` | `keepa_miss_done` ／ ①ログ `miss=` | 効率合否 `cause=wrote0` の FAIL 件数（0） |
| ④ | まとめ列挙の打ち切り変更 | `PR_CYCLE_ENUM_BATCH_STOP_QUEUE` | `stop=queue_found=` | `enum_batch n=` の平均（1.5→8 に近づく） |

**押さえどころ**: 大・特大は `smN>0` **または** `unkN>0` の間ずっと後回し（`prWashContinueGuard_`）。つまり **`unk=` が 0 になるまで伊藤園・ロッテ等は洗われない**。①②④はすべて `unk=` を速く減らすための対策で、`unk=` の減り方（社/時）が最上位の指標。

**push 後に人間が1回だけやること**: すでに空振り印が付いた社に①を効かせるため、Script Properties の `PR_WASH_TITLE0_FAIL_JSON` と `PR_WASH_TITLE0_SKIP_SET` を削除する（消さないと、その社は②で選定から外れたまま①の再検索を受けられない）。

#### 10.3ter 25時間空転の3対策（2026-09-10）

2026-09-09 04:01 から 25時間・549回、洗いが1件も進まなかった。原因は CYCLE §6.4quint の3点。**A・B・C も別トグル・別ログなので同時に検証できる**。

| # | 対策 | Property | ログの目印 | 合否に見る数字 |
|---|---|---|---|---|
| A | 通過を自社キューでも数える | `PR_CAND_PASS_BY_QUEUE` | `cand_pass_by_queue parent=… stolen=…` | `偽完了救済 n=`（誤発動が消える） |
| B | 救済は未処理>0のときだけ | `PR_FALSE_COMPLETE_PENDING_ONLY` | `empty_research queue=… all_done` | `洗い状態=Keepa中` で未処理0の社数（0） |
| C | 幽霊の続き相手を自動解消 | `PR_WASH_PHANTOM_FIX` | `phantom_keepa` / `phantom_keepa_resolved` | 同じ社の `catalog_first_other` 連続回数 |
| C' | 空転を合否で拾う | `PR_EFF_STALL_GUARD` | `効率合否 FAIL no_slice_stall` | `stallMin=`（60未満に保つ） |

**push 直後に起きること（手作業は不要）**: 中村食品産業は Keepa中・未処理0 なので、洗いが2回回ると C が働いて**完了へ自動で書き戻る**（A により候補1件が見えるので保留ではなく完了）。幽霊が消えるとペンギン食堂の列挙が始まる。**シートの `洗い状態` を人が書き換えないこと**（書き換えると C が働いたのか確認できなくなる）。

**確認**: push から約10分後に ①ログ で `phantom_keepa_resolved parent=中村食品産業` を探す。出ていれば C は合格。以後 `効率合否` の `stallMin=` が 60 未満に戻る。

**結果（2026-09-10 07:00〜18:25・11.4時間）**: A=合格（`cand=1` で完了へ）／C=合格（06:41:11 自動解消）／C'=合格（空転中に `no_slice_stall` FAIL 4件、復帰後 PASS）／B=判定保留（発動機会なし）。**1,478件を洗い・門通過61件・23社完走・品番リスト新規59件・1.23トークン/件**。再発ゼロ。

#### 10.3quat 段取りの空白の2対策（2026-09-10）

上の3対策で空転は止まったが、**129件/時＝Keepa理論上限244件/時の53%**しか出ていない。15分超の空白が19回・486分（窓の71%）。律速はトークンではなく段取り。正本は CYCLE §6.4sex。

| # | 対策 | Property | ログの目印 | 合否に見る数字 |
|---|---|---|---|---|
| 効率1 | 書いた洗い状態を手元にも反映 | `PR_WASH_ST_SYNC` | `catalog_first_other`（**減る**） | `full_trigger 洗い完走` の `catalog_first_other` 回数（25回→0に近づく） |
| 効率2 | 心拍譲りの予約を短縮しない | `PR_HB_YIELD_ARM_NO_SHORTEN` | `shop_lane_yield afterMin=5`／`wait_shortened` が付かない | `full_trigger shop_busy` の回数（36回→0に近づく） |

**押さえどころ**: 効率2 は「心拍が4分かかるのに次の洗いを2分後に予約していた」ので、`shop_lane_yield` の `afterMin=5` と、そのログに `wait_shortened=1` が**付かないこと**を見る。効率1 は「1社洗うたびに1回必ず出ていた `catalog_first_other`」が消えるかを見る（他社のカタログ見積が本当に未了なケースは残ってよい）。

**push 直後に起きること**: どちらも予約タイミングと分岐条件だけの変更で、シート・キュー・品番リストには触らない。既存の予約が消化されるまで最大5分は旧挙動が混じる。

**まだ直していない損失**（CYCLE §6.4sex・§6.4sept）:

- **相手の選定に134分（20%）**: 候補1社ごとに①洗いキュー（26,238行）を丸ごと読み直す。次段で「1回読んで全社ぶん数える」に置き換える
- **沈黙死の引き継ぎが35分待ち**: 8分で即 takeover する分岐はあるが `PR_WASH_EFF_LOOP_ENABLED=false` に束ねられて消えている。切り離すには新 Property が必要
- **Keepa API の一時障害**（17:52〜18:19）: `token_probe ok=0` のとき `liveLeft=0` なのに `効率合否` が `arm_ok PASS` を出す。判定が甘い

#### 10.3quin 社の切替22分を縮める2対策（2026-09-11）

正本 CYCLE §6.4oct。**§10.3quat で「心拍の自己衝突 約108分」としたのは誤りだった**ので訂正する。実測すると心拍が終わってから洗いが動き出すまでは押す前1.5分／押した後1.7分で**不変**（どちらも窓の6%）。心拍完了時に `hb_after_enum` が必ず1分予約を貼り直すため。効率2で減ったのは空振り実行36回→12回だけ。

本体は **1社完走→次の1件を洗い始めるまで中央値22.4分・窓の75%**。原因は「すぐ洗える相手の在庫がゼロ」で、列挙が1社見つかった時点で譲るため、洗い終わるたびに列挙からやり直している。

| # | 対策 | Property | ログの目印 | 合否に見る数字 |
|---|---|---|---|---|
| 在庫 | 洗える社が3社たまるまで列挙を続ける | `PR_CYCLE_ENUM_STOCK_N`（既定3・`1`で旧） | `enum_batch_stock`（GAS実行ログ）／`stop=queue_found=… stock=3/3` | `enum_batch n=` の平均（1.57→3に近づく） |
| 選定 | 未処理件数を1回読みの表にする | `PR_WASH_QUEUE_PENDING_MAP` | `queue_pending_map rows=… elapsedMs=`（GAS実行ログ） | `phase=pick_end elapsedMs=`（106秒→20秒台） |

**押さえどころ**: 最上位の指標は **LOOP COMPLETE から次の 洗いスライス までの分数**（22.4分 → 約4分が目標）。切片の刻み（同一社内 2.1分）は Keepa のトークン律速なので変わらない。`enum_batch_stock` と `queue_pending_map` は `Logger.log` なので①ログには出ない。シートで見るのは `enum_batch n=` と `phase=pick_end elapsedMs=`。

**心拍譲りは5分のまま**。7分にしても空転は縮まず、心拍が落ちたときの復帰だけ2分遅れる。

**押した直後に起きること**: 心拍1回の所要が約6分→約10分に伸びる（3社ぶん列挙するため）。そのぶん心拍の回数が約1/3に減るので、Catalog の総消費と `unk=` の減り方は悪化しない。心拍が長くなると `full_trigger shop_busy` が一時的に増えるが、§10.3quat のとおり実害はない。

#### 10.3sex 洗う前に「順位なし」を数える（2026-09-11・記録だけ）

正本 CYCLE §6.5。処理量はすでに **Keepa のトークン上限**（5/分＝300/時。236.8件/時 × 1.29トークン/件 ＝ 305トークン/時）に張り付いている。**回す速さの改善はもう効かない。** 次の余地は「何を洗うか」だけ。

門の落ち 17,209件のうち `stats空` 4,769（27.7%）＋`順位欠測` 1,487（8.6%）＝**6,256件（33.8%）は、1件1トークン払って「分からなかった」だけ**。

SP-API で400件を突き合わせた（`python tools/purchase_research_path3/_tmp_salesranks_check.py`・読取のみ・書込なし）。

| 門の結果 | Amazon 順位なし |
|---|---|
| 通過 (n=100) | **1%** |
| stats空 (n=100) | **99%** |
| 順位欠測 (n=100) | **96%** |

**順位の「値」では切れない。** Amazon はサブカテゴリ順位、Keepa は食品全体の順位でスケールが別物（`Keepa=順位>60000` の商品が `Amazon=22位 食品・飲料・お酒 | 干菓子`）。使えるのは「順位が1つも付いていない」だけ。

| # | 対策 | Property | ログの目印 | 合否に見る数字 |
|---|---|---|---|---|
| 記録 | 列挙で `salesRanks` も取り、順位なしを数える | `PR_CATALOG_RANK_PROBE`（未設定=ON・`false`で旧） | `列挙完了 … noRank=41/120(34%)`（①ログ）／`enum main … noRank=`（GAS実行ログ） | `noRank` の割合が **30〜35%** に乗るか |
| 切り | 順位なしをキューに入れない | `PR_CATALOG_RANK_CUT`（**未設定=OFF**。`true` で切る） | `noRank=… cut=ON` | 門通過率 7.0%→10.6%／門通過 14.1→約21件/時 |

**手順（この順を飛ばさない）**

1. ~~`PR_CATALOG_RANK_PROBE` は未設定のままでよい（既定ON）。`clasp push` して1日回す~~ **完了（2026-09-12〜13・25.5時間）**
2. ~~①ログの `列挙完了` 行で `noRank=` の割合を見る~~ **完了。54社・43.1%。社ごとの予測と実際のズレは中央値3pt（CYCLE §6.5.5）**
3. ~~乗らない・極端なら切らない~~ **判定: 切ってよい。門通過457件のうち Amazon 順位なしは0件＝失う通過なし**
4. Script Properties に `PR_CATALOG_RANK_CUT` = `true` を追加 ← **いまここ**
5. 切った後の合否は **門通過/時**（17.9→約29）で見る。件/時は下がって当然（洗う母数が38%減るため）

**注意**

- **順位の値で落とす実装を足さないこと**（CYCLE §6.5.3）。Amazon 順位と Keepa 順位の閾値を同一視すると通過の93%を捨てる
- `PR_CATALOG_RANK_PROBE=false` にすると `rankN` は **-1（判定不能）** になり、CUT も自動で効かなくなる
- 既にキューにある26,238行には効かない。既存キューへの一括適用は全件ループなので**別途承認**
- `Catalog見積` の数字が34%小さくなる。過去の完了社と単純比較しない
- 新発売で順位が付いていない商品は取りこぼす（実測1%。黒糖本舗垣乃花の3袋セットがこれ）。捨てた ASIN の再評価は未設計

#### 10.3sept 1社が102社を止める空転を直す（2026-09-13）

正本 CYCLE §6.6。**2026-09-13 03:00〜04:20 は洗い0件・1.5時間**。Keepa は満タン（`liveLeft=300`）でトークン待ちではない。約200件ぶんの取り逃し。

原因は「Keepa完走直後に店レーンを**1回**挟む」（§6.3bis 案B）が、店の未調査が残る間ずっと居座っていたこと。心拍が `wash_small` に届かず、小・中1社が残ったまま `defer_xl_sm_remain` で大26社・特大76社も止まった。

| # | 対策 | Property | ログの目印 | 合否に見る数字 |
|---|---|---|---|---|
| 1回だけ | 店レーンを1つ返したら印を消す | `PR_CYCLE_SHOP_AFTER_SM_ONCE`（未設定=ON・`false`で旧） | `shop_after_sm_once`（GAS実行ログ）／detail の ` once=Y` | `効率合否 FAIL cause=no_slice_stall` が **0回** になるか |

node で実コードを検証: 6/6 PASS（`node tools/common_ss_view/_tmp_shop_once_check.js`）。

**押さえどころ**: 小はほぼ終わり（96完了・71保留・24空欄）で、残りは大26社・特大76社。**この空転はこれから常態化するので、押さないと夜間が止まる。**

**見るもの**: `効率合否` の FAIL と、そのとき同じ行に出る `sm=` / `unk=`。`sm=1 unk=0` で `liveLeft` が高いまま切片が出ていなければ、まだ同じ形が残っている。

**残るリスク**: `sm>0` の1社が別の理由（列挙失敗・幽霊）で進めないと、印を消しても同じ停止になる。そのときは `no_slice_stall` が再び出るので報告する。

#### 10.3oct 未対応で残っている2件（2026-09-13）

| 件 | 実測 | 原因 | 状態 |
|---|---|---|---|
| Keepa 429 が3回 | 09/12 09:33・11:43・11:49。`tokLeft=-2 / -1 / 0` | トークン残の見積りがずれて残り以上を要求。`tokLeft` 中央値が ③106→④**17** まで下がり常に上限ギリギリ | **未対応**。3回とも自動回復。頻度が上がったら見積り式を見る |
| メール権限エラーが2回 | 09/12 21:35（`hb_enum 洗い完走` を落とした）・09/13 09:40（`週次転記チェック`） | `Session.getActiveUser()` 例外。コードは4か所とも先に `PR_CYCLE_DAILY_MAIL_TO` を読む（「読まない経路」ではない）。**当時 Property が空だったか、空扱いだった**ときだけ右側に落ちる。`oauthScopes` に `userinfo.email` が無く例外。2026-09-13 時点では Property は設定済。残るのは `MailApp` 用 `script.send_mail` が明示スコープに無い点（`gmail.send` のみ） | **一部解消**（宛先は設定済）。スコープ追加は未対応。21:35 は列挙処理そのものが FAILED |

#### 10.3dec 店の件数読みで A-light が殺される（2026-09-13）

正本 [CYCLE §6.6bis](B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)。

| 実測 | 値 |
|---|---|
| 洗い 11:19〜13:08 | **0件**（トークン `liveLeft=300` 満タン・`liveConsumed=0`） |
| pick（やることなし） | 96〜129秒 → **470〜627秒** |
| 境目 | 小中が0になった 11:52（`defer_xl_sm_remain` → `defer_shop_before_large`） |
| A-light 完走 | 11:42 が最後。12:09・12:51 は開始のみで**完了ログなし**（実行枠切れ。エラーは残らない） |
| `shops_ck_on` | 12:09〜13:08 の**59分間 21のまま** |
| `効率合否` | `FAIL cause=no_slice_stall` ×5 連続 |

原因は洗いキューの行数ではなく、**キュー未済106社ぶん繰り返される外部シート読み**（`セラー` 全読み×2＋`①候補` 26,009行）。pick＋pool で25〜30分を使い切るため、21分かかる A-light が毎回途中で死ぬ。

**押したあとに見るところ**

- `①ログ` の `deferred_skip` に **`shopCacheHit=`** が出る。社数ぶん（100前後）なら1回読みで効いている
- `pick_end elapsedMs=` が **十数秒**に戻っているか
- `heartbeat alight` の **DONE が出るか**（12:09 以降は出ていなかった）
- `shops_ck_on=` が1回の実行ごとに1つ減るか

**注意**: 拾う店が特大ばかりだと `status=hold band=特大` で `done=0` のまま1軒ずつしか減らない（実測2軒とも 8,835 SKU / 99,200 SKU）。21軒を抜けるのに21回の実行が必要。

#### 10.3undec 大・特大の発火条件が逆だった（2026-09-13）

正本 [CYCLE §6.8](B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)。

**想定していた順番**（人）: 小中を3件ストック → 小中が見つからなくなる → セラーリサーチで新メーカー → 洗いで小中大特大へ仕分け → 小中あり なら小中へ。**大・特大はこれが尽きたときだけ**。

**実際に起きたこと**: 新メーカーが**見つかった**ら大へ行っていた。

| 実測（2026-09-13） | 値 |
|---|---|
| 10:44 A-light | `new_makers=2 daily_done=6 rounds=6` |
| 11:20 選定 | `cont=キーコーヒー via=queue allow_xl sm=0 unk=0 large=28 band=大` |
| そのとき未調査の店 | **21軒**（まだ店ラウンドが残っている） |
| キーコーヒー | Catalog見積 306件 ≒ 487トークン ≒ 回復 **1.6時間** |
| 増えた2社 | **株式会社マッセル／株式会社K&K**。実態=未調査・規模帯=空・見積=空＝**洗えない** |
| 11:42 A-light | `new_makers=0` → 12:01 からは正しく `defer_shop_before_large` |

判定は `prAlightShouldShopBeforeLarge_()` の `lastNew === 0 && rounds < roundMax`。`done=6`（目標5超）・`lastNew=2` で最後の `return false` に落ちていた。「見つかった数」を「洗える材料が増えた」と同一視していたのが誤り。

**押したあとに見るところ**

- `①ログ` の `deferred_skip` に **`shop_rounds_until_cap`** が出る
- `heartbeat alight` の `daily_done=` が **10/10 まで伸びる**（途中で大に切り替わらない）
- `shops_ck_on=` が減り続けるか
- 10ラウンド使い切って小中が0なら、従来どおり大・特大に入る（週2社）。**これは正**

**空転しない理由**: `done >= maxD`（既定10）と `rounds >= roundMax`（既定10）はどちらも日次カウンタ。日付が変わるまで必ず止まる。node 実コード検証 **10/10 PASS**。

#### 10.3duodec 新メーカーが「要確認」で行き止まる（2026-09-13）

正本 [CYCLE §6.9](B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)。

`action=jitta_resolve` の最終実行は **2026-09-04 7:47**。9日間動いていない。理由は「順番で負けた」のではなく **`jittaPend = 0` と数えられている**から。

| 実態 | 信頼度 | 社数 | 状態 |
|---|---|---|---|
| 要確認 | high | 9 | AI確定。**人が判断するしかない**（対象外のまま） |
| 要確認 | medium | 10 | 行き止まり → 直した |
| 要確認 | low | 16 | 行き止まり → 直した |
| 未調査 | medium | 2 | 行き止まり（マッセル／K&K）→ 直した |

`prMakerNeedsJittaResolve_` は「信頼度 `high` 以外は再調査が必要」と見るのに、`prMakerSummaryRowUninvestigated_` は「信頼度が空でなければ対象外」と見る。両方を AND で見るため **28社が誰にも引き取られない**。

**AI もトークンも要らずに洗える4社**（マスタに答えがあるのにサマリに映っていなかった）:

| メーカー | サマリ.実態 | マスタ.実態候補 |
|---|---|---|
| 株式会社マッセル | 未調査 | **製造メーカー**(medium) |
| 株式会社K&K | 未調査 | **製造メーカー**(medium) |
| 高橋商店 | 要確認 | **製造メーカー**(medium) |
| 笑喜亭　タカヒラ | 要確認 | **製造メーカー**(medium) |

**押したあとに見るところ**

- `①ログ` に **`heartbeat jitta`** が出る（9/4 以降出ていなかった）。detail は `resolve=… pending=… round=1/8`
- メーカーサマリで上の4社の `実態` が **製造メーカー** に変わるか。変われば `規模帯` 付け（列挙）に進む
- 解決しなかった社は GAS 実行ログに `jitta_retry_exhausted maker=… jitta=要確認` が出て、マスタの `実態再採取しない` に レ点 が付く。**二度と AI を呼ばない**
- 初回だけ Gemini 呼び出しが最大21件（マスタに行が無い社）。以後は0件
- `heartbeat pool` の `sm=` が0から増えるか

**人がやること**: 要確認 high の9社は自動では動かないので、洗いたければメーカーサマリの `実態` を手で `製造メーカー` か `非仕入先` に直す。実メーカーは 鈴商（輸入食品菓子卸）・モンテ物産（酒類食品輸入卸）・日仏貿易（食品飲料輸入商社）・フジフードサービス・豊産業。非食品は アスク=PC周辺機器・HITOSIZUKU=化粧品・大倉商亊=アパレル副資材。

**打ち切った社を復活させたいとき**: メーカーマスタの `実態再採取しない` のレ点を外す。

#### 10.3tredec 洗い状態の終わりを3区分にする（2026-09-13・実装済）

正本 [CYCLE §6.10](B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)。**実装済**。人がサマリの `洗い状態` を手で直さない。初回心拍で移行する。

| 洗い状態 | 自動 | 90日後 |
|---|---|---|
| 完了（通過0も含む） | 選ばない | また選ぶ |
| 調査対象外 | 選ばない | 戻さない |
| 要修復 | 選ばない | 戻さない |

実装後に見ること:

- ①ログ `wash_st_split n=完了/対象外/要修復`（初回だけ。76／title0／7 が目安）
- エバラ食品などが `完了`＋`done_no_pass` になっているか（保留のままなら移行前）
- 五洋食品産業などが `要修復`＋`cand_missing` か
- `heartbeat pool` の `sm=` が移行直後に76増えていないか（`最終調査日` を今日にしたか）
- 偽完了救済の自動が0件か

切戻し: `PR_WASH_ST_SPLIT=false`。`PR_WASH_ST_SPLIT_MIGRATED` を消すと再移行する。

メーカーサマリの **大／特大** 行は薄いグレー（条件付き書式。心拍または規模帯書込のあと）。小・中は色なし。

#### 10.3non 洗いキュー7.5万行とセル上限（2026-09-13）

正本 [CYCLE §6.7](B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)。

**先に効くのはセル上限。** `自動リサーチ260814_アマゾン調査` は **7,003,555セル／上限1,000万＝70.0%** だった。上限に当たるとスプシ全体が書き込み不能になり①が全停止する。

| 手 | 状態 | 効果 |
|---|---|---|
| A. 機械タブの空列削除（①洗いキュー26→7列 ほか5タブ） | **実施済 2026-09-13** | 7,003,555 → **4,590,446セル（70.0%→45.9%）**。行は1行も消していない |
| B. 完了／調査対象外／要修復の社の Keepa済行を自動削除 | 未実装（要件 §6.7.4／§6.10） | 75,931 → 49,904行。スライスの全行走査が34%減 |

**なぜ90日削除では1行も減らないか**: 列挙日時が 8/22〜9/13 しかなく、90日を超えた行が0件。`洗いキュー期限切れ削除` のログは 8/22 以降ずっと `deleted=0`。

**Bの手順（実装後）**

1. 心拍を止める: Property `PR_CYCLE_HEARTBEAT_ENABLED=false`
2. 走っているスライスが終わるのを `①ログ` で確認する
3. メニュー **99運用 → 14. 洗いキューの期限切れ削除** を1回
4. `①ログ` の `洗いキュー期限切れ削除` が `deleted=26000前後 method=rewrite` になっているか見る
5. 心拍を戻す: `PR_CYCLE_HEARTBEAT_ENABLED` を削除（または `true`）

**初回の一括だけ人がやる理由**: このコードには `LockService` が1か所も無い。2万行の書き戻し中に別のトリガーが追記すると行が消える。2回目以降は完走1社ぶん（≈400行）なので、完走処理と同じ実行内で自動で消える。

**禁止**

- 大・特大の**未済48,167行を消す**こと（これが洗う材料そのもの。消したら列挙からやり直し）
- `洗い状態` を人が書き換えて掃除対象を作ること
- 人が使うタブ（見積もり依頼PV・③PV・②仕入れリスト・調査済みメーカー・マスタ等）の列削除。数式や予定列が隠れうる

### 10.4 CYCLE チェックリスト（運用）

- [x] メーカー **1社完走→次社**（アサヒ→永谷園。2026-08-24）
- [x] **小→中**優先。小中=0 なら店ラウンド≤10のあと大・特大（正本 §4.4／§7。GAS実装済・要 clasp）
- [x] 他メニューで停止 → 手動再開 or **01:00**
- [x] **日次生存メール**（停止中も送る・本文に完走／門通過）
- [ ] 人は **要確認** ＋ E〜J（次）

### 10.5 規模帯・心拍（2026-08-26・心拍完成形）

正本 [CYCLE §6.3〜§7・§13](B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)。

仮想: `python tools/competitor_store/dry_cycle_c3_select.py`（11/11）／ `python tools/competitor_store/dry_cycle_heartbeat.py`（39/39）

GAS: `prRunWashFullFromTrigger` が続き無し時に `prDispatchHeartbeat_`（列挙／小中洗い／**実態AI**／店フル／自動レ点／C0／大）。**列挙（enumOnly）後は `Catalog見積`/`規模帯` を残し Keepa中にしない → 心拍再arm**。小中が残る間／**規模不明が残る間**は大・特大のキュー続きを後回し。**小中=0かつ未調査ありは店より先に実態AI**（§4.5bis）。店は **1店フル**（特大店≥501はフルしない）。日次レ点既定8・店目標5・上限10・新メーカー0で最大10ラウンド。自動レ点は小→未知→中。実態AI前の未調査≥30で自動レ点抑制（要確認山では止めない）。切戻し `PR_CYCLE_HEARTBEAT_ENABLED=false`／実態AIレーン `PR_JITTA_HB_LANE_ENABLED=false`。

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
