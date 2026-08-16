# 仕入れ検討① — いまここ

**日付**: 2026-08-16  
**完成形フロー・状況報告の正**: [B_PURCHASE_RESEARCH_FLOW.md](B_PURCHASE_RESEARCH_FLOW.md) §3（A–J / A′）。契約・A/D/E は §5。3者は §6  
**正**: [DOMAIN1_RESEARCH_PURCHASING.md](../DOMAIN1_RESEARCH_PURCHASING.md) §3.5・**§3.6**・§6  
**第1版要件**: [B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md](B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md)  
**タスク範囲・開発予定の正**: [B_PURCHASE_RESEARCH_TASK_STRUCTURE.md](B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)（ツリー本体はそこ。本ファイルは同文＋数字。ツリーの A/B/C は旧領域ラベル）  
**他PJ共有**: ツリーは TASK_STRUCTURE。貼付・競合ストア・出品Bは TASK_STRUCTURE の「他PJ」行を守る。

---

## いまここ

**進め方（2026-08-15 再ロック）**: Keepa の正は競合DB `Keepaフル`（90日再利用）。`①候補` は門の作業面（転記）。Catalog 生はモールヒットに入れない。セラー台帳は競合DB `セラー`（第1版はモリタ1行）。Keepaフルからセラーは貯めない。

### 門検収（C1–C8）— 済

作業面に門が残ること。石原＋五木1語。**済。** これは倉庫コンセプトの一部であり、全部ではない。

| ID | タスク | 状態 |
|----|--------|------|
| C1–C8 | Catalog→Keepa門→①候補（石原＋五木通過5） | **済** |

### 倉庫本線（W）— いまの次

コンセプトどおり Keepa を競合DBに載せ、90日内は取り直さない。

| ID | タスク | 状態 |
|----|--------|------|
| W1 | ①の Keepa `history=0&stats=90` 応答を `Keepaフル` へ（purpose=research。csv[]落とす。価格指紋） | **済**（スライス5） |
| W2 | 既存①候補 ASIN を Keepaフルへバックフィル（90日内なら再GETしない） | **済** 181/181 |
| W3 | 以降の門は Keepaフル優先。ミスしたら GET。①候補は転記 | **済**（転記＋メニュー 20:20 pass22/drop8） |
| W4 | 貼付Aは Keepaフルを90日内読む（メーカー洗いはしない） | **dry合格** paste60 skip39 need21。本線は出品 Property |

**共有用・進捗つき構造（2026-08-15）** 正は [TASK_STRUCTURE](B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)。凡例: 済／今／次／後／外／**止**

```text
領域1
├─ A 問屋                         外
├─ B 商品リサーチ
│  ├─ ① 仕入れ検討 ★リサーチ見積もり agent
│  │  ├─ 進め方
│  │  │    Keepaの正=競合DB Keepaフル  ロック（90日。出品と共用）
│  │  │    ①候補=門の作業面（転記）   ロック（倉庫ではない）
│  │  ├─ 実行基盤
│  │  │    調査複製スプシ              済（①候補 181行。競合タブを増やさない）
│  │  │    専用 GAS purchase_research/ 済（出品 clasp 禁止。Catalog は未搭載）
│  │  ├─ 門検収 C1–C8                 済
│  │  ├─ 倉庫本線
│  │  │    W1 Keepaフルへ①書き       スライス5合格
│  │  │    W2 ①候補ASINバックフィル  済（181/181。skip5＋GET176）
│  │  │    W3 門はフル優先→①候補転記 済（メニュー20:20）
│  │  │    W4 貼付90日後読み          出品PJ（dry: paste60 skip39 need21）
│  │  ├─ 前段①–⑧
│  │  │    ① セラー貯め→ストア正     済（タブ「セラー」モリタ1行。フルから貯めない）
│  │  │    ② ピック                   済（食品root・モリタ。S4で構成比）
│  │  │    ③ 巡回                     済（第1版1セラー。query total=30。productなし）
│  │  │    ④ /query GET               済（モリタ30）
│  │  │    ⑤⑥⑦ 分類・食品・メーカー  済・仮想（モリタ30）
│  │  │    ⑧ 週次メーカー決定         済（今週=永谷園。Catalog GETなし）
│  │  ├─ 品番リスト転記
│  │  │    O1–O3 出品者数 COUNT_NEW  済（QS一致。空61＝今オファーなし）
│  │  │    F0 直販・新品現在・現行順位 済（GETなし。いる34/いない186）
│  │  │    F1 Keepa列 flatten       済（cat217/梱包157/FBA164。BB価格はJSON空）
│  │  │    F2 画像／サブ画像          済（Eメイン・Fサブ。K8は空メモ）
│  │  │    T1–T2 品番リスト          済（通過82 JOIN。画像HYPERLINK。式32非書）
│  │  │    課題 K1–K8               止／流用／不要（下表）
│  │  └─ 見積書ルート                 後
│  └─ ② 出品用                       外（貼付 Catalog 階段と混ぜない）
└─ C 見積                             外
```

### 品番リスト×Keepa（2026-08-15 夜・ロック）

人間の **仕入検討品番リスト**（複製スプシ、ヘッダ行4、利益式あり）は `①候補` ではない。転記可能なものは Keepaフルの JSON／列から。Keepaに無いものは転記しない。

| 項目 | 正・メモ |
|------|----------|
| 出品者数 | `stats.current[11]` **COUNT_NEW**。クイックショップ一致。Keepa `offers[]` のユニーク sellerId は過大。SP-API `getItemOffers` TotalOfferCount は COUNT_NEW と一致（O2: 200×2）。本線で `offers=20`／ItemOffers は使わない |
| 空61 | tot=0・c11=-1・新品現在価格なし。**0を書かない**。うち門通過11は avg90 のみ（カタログ／旧CSVの死に体） |
| Amazon直販 | `availabilityAmazon` → 列 **Amazon直販**（いる/いない）。220全埋め。GETなし |
| 新品: 現在価格 | `current[1]`。159埋め／61空。最安の正。BBは使わない |
| 売れ筋ランキング: 現在 | `current[3]`。169埋め |
| Amazon: 180日在庫切れ% | `outOfStockPercentage180[0]`。220全埋め。csv[]不要 |
| メーカーが出品者 | 課題 **K1**。無理にやらない |
| FBA出品者（人数） | 課題 **K2**。無理にやらない |
| Buy Box 店名／現在BB価格 | 課題 **K3**。品番リストでは不要。列は空（**K8**） |
| 標2b／自己発送概算 | 課題 **K4・K5**。出品ロジック流用（L1） |
| おまけ・出品エラー・閲覧数・売れる確率 | 課題 **K6**。不要 |
| 過去1か月販売数 | 課題 **K7**。`monthlySold` 欠測多い（20/220）。GET増では埋まらない |
| 利益・仕入・有名／調査者 | Keepa外（人間・卸・式） |

### 課題（Keepaで埋まらない／今はやらない）— 2026-08-15 判断

品番リスト突合の残り。**追加 GET の本線には載せない。** 再開は各行の条件のみ。

| ID | 項目 | 判断 | 再開条件・メモ |
|----|------|------|----------------|
| K1 | メーカーが出品者 | **無理にやらない** | 契約で token が増えたら通過行だけ `offers=20`（sellerId／店名 vs ブランド）。目安 ≈7 token/ASIN |
| K2 | FBA出品者（人数） | **無理にやらない** | `offerCountFBA` は全件 -2。人数も `offers[]` が要る。K1 と同じ再開 |
| K3 | Buy Box 店名／現在BB価格 | **品番リストでは不要** | `current[18]` 空。最安は新品 `current[1]`。店名は offers。**列は残す。空のまま（K8）** |
| K4 | クイックショップ「標2b」 | **出品ファイルのロジックを流用** | APIに文字列は無い。梱包mm→出品 `00_設定マスタ` FBA手数料 first-fit（L1） |
| K5 | 自己配送 送料概算 | **出品ファイルのロジックを流用** | Keepa外。出品の送料マスタ／梱包VLOOKUPと同系。FBAと混ぜない（L1の別枝） |
| K6 | おまけ疑い・出品画面エラー・閲覧数・売れる確率 | **不要** | Keepa product に無い。門にも入れない（Q5/Q8） |
| K7 | 過去1か月販売数 | **欠測のまま** | `monthlySold` は Keepa が返さないことが多い。パラメータ追加では埋まらない |
| K8 | Buy Box現在／30／90、レビュー評価・件数、BuyBox_FBA | **今のJSONでは取れない。いずれ取得方法か代替を検討** | 本線 GET には載せない。候補は下表 |

**K8（今の倉庫JSONで空）— 2026-08-16 メモ。実装しない。**

Keepaフル列はある。本線は `history=0`・`stats=90`・offersなし・`csv[]` 非保存のため、flattenしても空。

| 列 | 今空の理由 | いずれ取るなら（未実施） | 代替（未実施） |
|----|------------|--------------------------|----------------|
| Buy Box: 現在／30日／90日 | `current[18]`・`avg30/avg90[18]` が倉庫で空。価格時系列は `csv[]` 側 | 一時 GET で stats だけ列化（csvは捨てる）。token増 | 人間の Keepa CSV（石原突合では BB平均は csv付きAPIと一致）。出品 Keepaスナップショットの同名列 JOIN |
| レビュー: 評価／件数 | `current[16]`/`[17]` 空。`rating`/`reviewCount` も欠 | 同上。または SP-API カタログの星・件数 | Keepa画面／スナップショットのレビュー列。Quick Shop 目視 |
| BuyBox_FBA | `stats.buyBoxIsFBA` が無い | `offers=20` の BB 行 `isFBA`（K1と同コスト目安） | Keepa CSV「FBAです」。出品スナップショットに無ければ目視 |

やらないこと（いま）: 倉庫に `csv[]` を残す。全件 offers GET。品番リスト判定に BB を必須にしない。

**他PJが守ること（2026-08-15）**

| 相手 | 守ること | ①側のいま |
|------|----------|-----------|
| 出品・貼付 | メーカー洗い全件 Catalog を貼付階段に使わない。**Keepaは Keepaフルを90日内読む**（①が書いた行） | 門の作業面は `①候補` |
| 競合ストア | Catalog 生を `モールヒット` に入れない。**Keepaフルへ①が書く（W1）**。メーカーマスタ・**セラー**は競合DB | 台帳モリタ1行済 |
| 出品 clasp | ルート `.clasp.json` を①に差し替えない | ①は `purchase_research/` のみ |

**ドライラン（2026-08-15・こちら実施）**

| ID | 結果 | 判定 |
|----|------|------|
| T-query-page | **GET selection=** total=**30** n=30。POSTボディ形式は不合格のまま | **GET合格** |
| T-gate-csv | 石原DL143。通過65／落ち78（価格30・順位40・冷凍8） | 仮想合格。①候補へ書込済 |
| T-itsuki-class | keywords200: 麺165。brand500: 麺367 | 分類手順は可。第2クエリ「麺」では500打ち切りは解けない |
| T-itsuki-label | brand500内: ラーメン139・麺以外117・うどん86・そば71・その他麺41・スパ19・チャンポン16・そうめん11 | 単ラベルでも最大139。500打ち切りは「ラーメン」1語では解けない見込み |
| T-itsuki-q2 | 公式6語×8頁はすべて HTTP200 n=160（満杯）。タイトル含有: ラーメン146／うどん129／そば123／そうめん142／スパ145／ちゃんぽん109 | 第2クエリは可。8頁では打ち切り判定不可 |
| T-itsuki-q3 | `五木食品 ラーメン 乾麺` 16頁309（終了）。`うどん 乾麺` 13頁248 | 打ち切りは解けた。乾麺はページ分割用 |
| T-itsuki-champon-e2e | ちゃんぽん Catalog 13頁242終了。タイトル含有158／語39。門 通過5／落ち19／stats空15／消費39 | **五木試す合格** |
| T-c8-itsuki-cand | runId `pr_20260815_itsuki_c8`。五木通過5。new=5。①候補181。門作業面 | **門検収合格**。Keepaフルは W1 |
| T-w2-backfill | runId `pr_20260815_w2`。Keepaフル181。skip5＋GET176 verify 176/176 | **書込合格** |
| T-w3-gate | フル再門 match=181。転記26。メニュー20:20 pass22/drop8 | **転記＋メニュー合格** |
| T-s2-seller | 競合DB「セラー」1行 モリタ。storefront=1なし。フルから貯めない | **書込合格** |
| T-s8-week | 今週=永谷園。石原・Generic除外。サトウは次点。Catalog GETなし | **決定合格** |
| T-s8-cat | 永谷園1語 Catalog 16頁320打ち切り。みそ汁終了319・あさげ終了283。Keepa非GET | **打ち切り観測合格。一括属性は不合格（過大）** |
| T-naga20 | あさげタイトル先頭20。Keepa19＋skip1。門 pass4/drop16(価格)。①候補+19。ヒット127 | **キャップ洗い合格** |
| T-sato20 | 切り餅タイトル先頭20。Keepa20。門 pass7/drop13。①候補+20。ヒット127 | **キャップ洗い合格** |
| T-o1-offers | Keepaフル220。offers[]=0。totalOfferCount 220（うち0が61）。COUNT_NEW 159でtotalと一致。FBA/FBMは全部-2 | **GETなし計測合格。人数の正には未。O2で突合** |
| T-o2-offers | 通過5。Keepa offers消費36。SP 200×2は TotalOfferCount=COUNT_NEW。429×3。Keepaユニークセラー≠COUNT_NEW | **人数の正=COUNT_NEW/totalOfferCount。offersリストは人数に使わない** |
| T-o3-col | Keepaフル「出品者数」159埋め／61空。GETなし。人間確認でQS一致 | **列化合格** |
| T-o4-live | 159は新品現在価格あり＝今オファーあり。61は tot=0・新品価格なし。うち門通過11はavg90のみ | **空61は今出品なしと整合。GETなし** |
| T-l1-dry | 設定マスタ FBA21／自己発14。梱包あり157。first-fit 157/157。GETなし。出品コード非改 | **dry合格** |
| T-t1-dry | 通過82・既存0・式32列。原本ガード。GETなし | **dry合格** |
| T-f2-img | images[] 全220。画像1枚目220・一覧220（複数183）。GETなし | **列化合格** |
| T-f2b-img | 画像=HYPERLINKメイン220。画像一覧→サブ画像（\|区切り・2枚目以降183）。GETなし | **列化合格** |
| T-f2c-move | サブ画像を画像の右隣（E→F）へ移動。GETなし | **列位置合格** |
| T-t2-join | 品番リスト125中 JOIN82。画像URL=HYPERLINK82。式32非書。miss43はGETせず | **列化合格** |
| T-s3-dry | セラー1行 AYC4Z8PML8T30。ヘッダ揃い。GETなし | **計画合格** |
| T-s3-query | GET /query page0。total=30 n=30 consumed=11。①候補ヒット30。productなし。台帳巡回日・件数更新 | **合格** |
| T-s4d-ver | セラー28列合計100。食品46.9。GETなし | **再確認合格** |
| T-l1-clean | 旧構成%4列削除。28列・食品46.9維持 | **合格** |
| T-l2-miss | 品番miss42をKeepaフルへ append42 consumed42。品番非書 | **合格** |
| T-l3-join | リスト125 miss0。式32非書 wrote321 | **合格** |
| T-l4-seller | 次セラー AN1VRQENFRJN5 storefront=0 cats10 食品0%。asinListなし | **合格** |
| T-l5-query | 同セラー /query page0 total=97843 n=100。productなし | **合格**（page進めず） |
| T-l6-naga | 永谷園みそ汁 Catalog1頁20。門 pass1/drop18。①候補+16。ヒット不変 | **合格** |
| T-p1-ocha | 永谷園お茶づけ Catalog1頁。タイトル含有12。Keepa12 consumed12 left268。門 pass3/drop9。①候補+12。ヒット425不変 | **フェーズ1合格** |
| T-p2-seller | モリタ30のBuyBox。既存2店以外0。次店なし | **フェーズ2停止** |
| T-p3-query | 次店query | **未（2で停止）** |
| T-p4-list | 通過4を品番リスト追記。JOIN miss0 cells0。式33非書。原本非書 | **フェーズ4合格** |
| T-p2b-asin | B09LD4YFF5 倉庫あり BuyBox空 offers0。出品一覧GETせず | **①補充不可** |
| T-p2c-seller | A3L1BD8LA6USKU storefront=0 食品61.5 cats10 n台帳3 | **合格** |
| T-p3c-q | A3P240C6ET053M query page0 total=163 n=100 in_cand=4。1000以下 | **合格** |
| T-p1-pg1 | 同店 page1 n=63 total=163。全件ダンプなし | **フェーズ1合格** |
| T-p2-k20 | page1 miss63のうちKeepa20 rest43。製造者は天然生活7件ほか | **フェーズ2合格（残りあり）** |
| T-p4-osu | 永谷園お吸いもの Catalog1頁 タイトル13 Keepa13 pass1/drop12 ①+13 リスト+1 | **フェーズ4・5合格** |
| T-p2-skip | A3L1 メモ「洗わない」。出品一覧GETなし | **確定** |
| T-p1f | 永谷園ふりかけ Catalog1頁 タイトル17 Keepa17 pass1/drop16 ①+17 ヒット不変 リスト+1 | **合格** |
| T-f0-dry | JSON flatten 220。いる34／いない186／新品159／順位169／oos180=220。GETなし | **dry合格** |
| T-f0-col | Keepaフル列4追加。直販220／oos180=220。①ログ f0col。原本非書 | **列化合格** |
| T-w4-dry | 貼付ASIN60 vs Keepaフル181。skip_fresh=39 need_get=21。GETしない・専用非書 | **計画合格**（本線は出品 12-㉓） |
| T-ishihara-nav | 公式 `/products/`: 食べるおだし・チーズかつお/まぐろチーズ・刺身・佃煮（＋生ハム） | 第2語に刺身/たたきは入れない（冷凍落ち） |
| T-morita-gate-ss | runId `pr_20260815_172558` ①ログ DONE pass=22 drop=8。①候補のうち当該30件 unique30。価格<2000が8 | **GAS合格**（Pythonドライラン一致） |
| T-query-gate-join | モリタ30 × 石原CSV: join 5／miss 25。hit内 通過2・価格落ち3 | セラー取扱の大半はメーカーCSVに無い。門は product 2段が本線 |
| T-seller-accum | BBセラーから22 ID（モリタ含む） | 貯めの材料は既存リサーチから取れる |
| T-ishihara-q2-official | 公式5語すべて HTTP200・非capped。件数 95/121/102/115/43。タイトル「石原水産」ユニーク **140**。Keepa CSV143 と join 108／Catalogのみ32／CSVのみ35 | **Catalogドライラン合格**。生はモールヒットに書かない |
| T-ishihara-catalog-only-gate | Catalogのみ32。`history=0&stats=90`。合計 **通過16／落ち16／消費32**。stats空0 | **Keepa門合格** |
| T-m5-cand-upsert | runId `pr_20260815_catalog_m5`。発見経路 `catalog_title`。①候補 upsert new=8 upd=8。verify 16/16。行数176。モールヒット非書込 | **書込合格** |

**止めていた理由（解消済）:** POST ボディ形式はフィルタ無効。**GET で解消。** GAS は GET のみ。

---

| 項目 | 状態 |
|------|------|
| 詳細要件 | 組み直し済（セラー起点・Catalog発見・Keepa属性・門はスプシ閾値） |
| 実行基盤 | **ロック**。① GAS=`purchase_research/`。Catalog 洗いは Python のまま |
| `/query` の正 | **ロック**（④はこの形。`sellerIds` は可変。下記） |
| 属性 PoC | Keepa product API＝人のDLと値一致。門は `stats=90` |
| 発見 PoC | タイトル含有が正。石原公式5語で Catalog のみ32→門16通過 |
| セラー | ①–⑧同じ第1版。取扱は `/query`。台帳シートは競合PJの次 |
| 専用スプシ | 複製済。①候補 **176行**（M5 verify 16/16）。メーカーマスタは競合DB |
| clasp | 出品用を触らない。① push は `cd purchase_research` |

**自動化用複製**（原本ではない）  
https://docs.google.com/spreadsheets/d/1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE/edit?gid=685859874#gid=685859874  
ID: `1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE`  
Property 候補: `PURCHASE_RESEARCH_SS_ID`（コンテナバインドでは不要）。Keepa キーは **契約1本**。① GAS の Script Properties に **同じキー**を入れる（プロジェクト別なのでコピーが必要）。トークンは出品と共用。  
**競合DB**: 調査複製に競合シートを増やさない。共有は `COMPETITOR_SS_ID`=`1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs`（[COMPETITOR_STORE_REQUIREMENTS.md](COMPETITOR_STORE_REQUIREMENTS.md)）。`purpose=research` のみ追記。出品 clasp 非差し替え。

---

## 最終形（短い）

```text
セラーリスト（sellerId・対象カテゴリ）
  → 定期巡回 → 取扱抽出 → 食品に絞る → メーカー名＋商品名一覧
  → 週1サマリーで洗うメーカーを決める
  → Catalog キーワード＋タイトル含有 → Keepa 属性（正は Keepaフル）→ 門
  → 作業面 ①候補（転記）／倉庫 競合DB Keepaフル
```

別ルート（残す）: 問屋見積／提案 → ドキュメント化 → JAN またはメーカー＋名 → Keepa 品番検索 → 同じ門。

発見はメーカー洗いが SP-API（タイトル含有）。セラーの取扱は Keepa 出品者検索＝`/query`。`offers[]` は使わない。ペヤング型・公式サイトは第1版に載せない。

---

## ④の正（製品ファインダー `/query`）

人が画面の「API クエリを表示する」で確定（2026-08-15）。**この形が取扱抽出の正。** `sellerIds` だけ巡回対象に差し替える。モリタストアは試行例であり必須入力ではない。

**④ /query HTTP（ロック・2026-08-15 再測）**: 画面と同じ **GET**  
`/query?domain=5&selection=` + JSONを URL エンコード。  
**POST で `{"selection": "..."}` はフィルタ無効**（total≈2.8億・デフォルト50件）。使うな。

再測: totalResults=**30**、asinList=**30**、ISBN 0。画面「30件」と一致。先頭に `B0DV46Z3D1` 等（石原リサーチCSVにもある）。consumed=11。キーは出さない。
- `offers[]`・`/seller?storefront=1` の全ASINは使わない
- 画面結果の例: 30 / 合計 30（モリタ・食品ルート）

```json
{
    "rootCategory": ["57239051"],
    "sellerIds": ["AYC4Z8PML8T30"],
    "productType": ["0"],
    "sort": [["current_SALES", "asc"], ["monthlySold", "desc"]],
    "page": 0,
    "perPage": 100
}
```

| キー | 固定／可変 |
|------|------------|
| `rootCategory` `57239051` | 食品ルート。おかし等はプロファイルで別 ID を検討 |
| `sellerIds` | **可変**（リストの各セラー） |
| `productType` `0` | 通常商品 |
| `sort` | 順位昇順 → 月販降順 |
| `page` / `perPage` | ページング。100件超は page を進める |

キーは URL に載せない（Script Properties）。

---

## ステップ（経路3先行は終了）

### 出品② Amazon貼付（2026-08-15 ロック・前提）

正本: [B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md](B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md)。**このチャットは実装しない。**

| 前提 | ①側の含意 |
|------|-----------|
| メーカー洗いは貼付の対象外 | ①の Catalog 全件洗いと、貼付の Catalog 階段（A′JAN→B′→D′メーカー＋名→E′核。C′は診断）を**混ぜない** |
| 領域1成果の後読みは任意のトークン節約 | 貼付の本線クエリではない。①は貼付用階段を書かない |
| Z併存 | `Keepaスナップショット`は薄い運用のまま。倉庫は `Keepaフル`＋生JSON。**①が書く（W1）**。スナップショットはフル化しない。`csv[]` 非保存 |
| Catalog 生は `モールヒット` に混ぜない | ①の Catalog 行は調査複製 `①候補`。競合DBへは `目的=リサーチ` と `メーカーマスタ` のみ |
| brandNames 単独は正でない | ①と同じ |
| 評価◎は人間 | ①の門は通過／落ち。◎を付けない |
| 出品 clasp・今値 GET・A–E 全本 | ①外 |

Keepa 属性の再利用: ①が書いた `Keepaフル` を90日内は貼付が読む（W4）。スキーマは P0 済。①書きは **W1**。

### 済

- 専用スプシ複製。出品 clasp 非差し替え
- Catalog 発見の正（タイトル含有）
- Keepa product 値一致
- 実行基盤ロック（調査複製専用 GAS）
- ④の正＝`/query` 形（`sellerIds` 可変）

## セラー台帳（他チャットへ）

正はストア新シート（hits ではない）。①は `目的=リサーチ` で追記。巡回キューだけ調査複製。列案: sellerId, 名, カテゴリ比, 取扱数, 最終発見日, 発見元。シート実装は競合PJ。

### 次（塊）

1. **W1**（承認後）: ① Keepa `history=0&stats=90` を競合DB `Keepaフル` へ（purpose=research。csv[]落とす）。コード未着手。
2. W2 既存①候補のバックフィル（90日内は再GETしない）→ W3 門はフル優先・①候補は転記。
3. セラー①–⑧は後。出品 clasp 禁止。Catalog 生はモールヒット禁止。

### 巨大メーカーのカテゴリ語（2026-08-15）

500件サンプルの細ラベルは **Amazon上位の偏り**であり取扱カテゴリの正ではない。

| 優先 | 手段 | コスト | 使い方 |
|------|------|--------|--------|
| 1 | 公式サイトの商品一覧ナビ | ほぼ0 | 五木: `itsukifoods.jp/product_list.html` → ラーメン／うどん／そば／そうめん／スパゲティ／ちゃんぽん（焼そば含む）。第2クエリの語 |
| 2 | Gemini `google_search` **メーカー1回** | 安い | 公式が無い／ナビが取れないとき。JSONで第2クエリ用の短い語だけ。90日キャッシュ。ASINごと呼ばない |
| 3 | OpenAI Web | 高い・出品429あり | Gemini失敗時のみ。本線にしない |

禁止: Amazon HTMLスクレイプ。モール横断のカテゴリIDマージ。公式サイトとASINの商品照合（第1版外）。LLMの語は **Catalog キーワードの接尾**であり Amazon ノードIDではない。

### 実装設計草案（コードなし・2026-08-15）

| 塊 | 中身 | 制約 |
|----|------|------|
| G0 プロジェクト | 別 clasp／別 scriptId。出品 `コード.js` に①を足さない。複製 SS `1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE` のみ | clasp 差し替え禁止 |
| G1 取扱抽出 | NOW の `/query` JSON。`sellerIds` だけ差替。page 進め。食品は `rootCategory` | `offers[]`／storefront 全ASINなし |
| G2 戻り確認 | **GET合格** 30件。POST `{"selection":}` は使うな | GASは GET＋URLエンコード |
| G3 メーカー洗い | 既存 `search_catalog_keywords.py` の正（キーワード＋タイトル含有）を GAS へ移植するとき同じ規則 | brandNames 比較用は本番の正にしない |
| G4 属性＋門 | Keepa product（属性 PoC 合格済み）。門の閾値は複製スプシのプロファイル行。コードにマジックナンバーを書かない | トークン待ち再開は① GAS |
| G5 週次 | トリガーは承認後。第1版はメニュー手動で①–⑧を1セラー | 自動全セラー巡回は後続 |
| G6 競合 | `COMPETITOR_SS_ID` へ `目的=リサーチ` のみ。複製に競合タブを作らない | ENABLED 相当は①側で別途 |

GAS ファイル案（未作成）: `PurchaseResearch.js` ＋ 薄い `appsscript.json`。Python path3 は検収比較用に残す。

### `/query` 実測（2026-08-15・モリタ `AYC4Z8PML8T30`）

**正は GET** `/query?key=…&domain=5&selection=` + JSON URL エンコード。  
**POST** `{"selection": "<json>"}` は HTTP 200 でもフィルタ無効（total≈2.8億・既定50件）。使うな。

| 項目 | 結果（GET） |
|------|------|
| HTTP | 200 |
| 本体 | **`asinList` のみ**（30 件＝画面）。`products` 配列は無い |
| トップキー | asinList, totalResults, tokensLeft, tokensConsumed, processingTimeInMs |
| title / current_SALES / monthlySold / price | **乗らない** |
| 含意 | 門・メーカー名は **product API 2段目**。`offers[]` は使わない |
| 注意 | ISBN 形が混ざる場合あり。食品絞りは2段目のカテゴリ／タイトルで再確認 |

キーは出さない。詳細 JSON は gitignore の `tools/purchase_research_path3/out/`。

### 0. 専用スプシ

**済**。GAS は openById 想定。書きは承認後。

---

## PoC の回し方（ローカル・検証用）

```text
cd tools/purchase_research_path3
python search_catalog_keywords.py --keywords "メーカー名" --max-pages 10
python keepa_csv_vs_api.py ishihara_keepa_2.csv
```

brandNames は検証比較用。本番の正はキーワード＋タイトル含有。出品②・原本は触らない。

---

## 実測メモ: 石原水産（2026-08-14・いまのテストメーカー）

五木は件数過多で差が見えにくかったため、メーカーを石原水産に変更。

| クエリ | ページ | 件数 | 名前に「石原水産」 |
|--------|--------|------|-------------------|
| keywords のみ | 18（最後まで） | 342 | 182／342。ノイズ多（ノーブランド、石原バイオサイエンス等） |
| keywords + brandNames=`石原水産` | 2（最後まで） | **31** | 31／31 |

Keepa `KeepaExport-2026-08-14石原水産リサーチ用.csv`（コピー `ishihara_keepa_research.csv`）:

- 143行・ASIN一意・**全行タイトルに「石原水産」**
- Catalog タイトル含有143件と **ASIN完全一致（B−A=0、A−B=0）**
- ブランド内訳も同じ（ノーブランド67／石原水産28／わが街とくさんネット19 等）
- 別目的の2行ファイルとは別物。こちらはキーワード網と同じ集合

含意: この取り方では経路3の「Keepaに無いASINを足す」は **0件**。価値は Keepa 側の **順位・Amazon価格・在庫**（Catalog summaries に無い門の材料）。抜粋: `石原水産_Keepaリサーチ用_抜粋.csv`。差分サマリ: `ishihara_diff_summary.csv`。

Catalog 結果（utf-8-sig）:

- `tools/purchase_research_path3/石原水産_SPAPI_キーワード_タイトル含有.csv`（**143件**。タイトルに「石原水産」。経路3の網は当面こちら）
- `石原水産_SPAPI_キーワード.csv`（342件・絞る前）
- `石原水産_SPAPI_ブランド指定.csv`（31件。brand完全一致。セット店が落ちる）
- 同内容英名: `ishihara_brand_all.csv`／`ishihara_keywords_all.csv`

読み取り（**第1版の正**）: キーワード検索のあと **タイトルにメーカー名が含まれる** で絞る。brandNames だけだと寄せ集めセットが落ちる。

---

## 実測メモ: 五木食品（2026-08-14・参考・件数過多）

手段: SP-API Catalog Items（読取）。HTMLスクレイプなし。専用スプシ・出品②非書込。レポート本体は `tools/purchase_research_path3/out/`（gitignore）。

| クエリ | ページ | 件数 | タイトル/ブランドに「五木食品」 |
|--------|--------|------|--------------------------------|
| keywords のみ | 10（上限打ち） | 200 | 189／200 |
| keywords + brandNames=`五木食品` | 15 | 300 | 300／300 |
| 同上 | 25（上限打ち） | **500** | 500／500 |

読み取り:

- **キーワードだけだとノイズが入る**（養命酒の五養粥、からだシフト、リンガーハット、マルちゃん、OTOKI 等。11件／200）。メーカー救いには `brandNames` が必要。
- **brandNames でも 500 でまだ次ページがある。** 五木はセット・パックが多く、経路3を「Amazon検索の全件」にすると 1行1ASIN が膨らむ。門（価格・順位）の前に件数爆発する。
- Catalog summaries には **価格・順位・本体在席が無い。** 属性は Keepa API（②シートは触らない）。
- **Keepa 集合 A との差分（2026-08-14）**: A=495（`itsuki_keepa.csv` の ASIN 列）、B=500（ブランド指定 Catalog）。**共通 495 / B−A 5 / A−B 0**。Keepa 495件は Catalog 500件に全部含まれる。CSV: `itsuki_diff_summary.csv`／`itsuki_diff_B_minus_A.csv`／`itsuki_diff_A_minus_B.csv`／`itsuki_diff_both.csv`。
- 両方とも約500で頭打ちの可能性が高い。B−A=5 は「Keepa未登録の救い」か「ソートの端」か未確定。真の漏れはこの件数より大きい可能性。
- CSV（utf-8-sig）: 全件 `itsuki_keywords_all.csv`（200）／`itsuki_brand_all.csv`（500・打ち切り）。先頭50は `itsuki_keywords_50.csv`／`itsuki_brand_50.csv`。置き場 `tools/purchase_research_path3/`。

方針: 五木は第1版で**試す**。全件500打ち切りのまま門に流さず、キーワードにカテゴリ／アイテム（例: `五木食品　麺`）を足す。それでも溢れたら人または後続。

---

## 物販Next PDF「商品リサーチ方法」との整合（2026-08-14）

原本: Downloads `商品リサーチ方法.pdf`（コンサル生専用・8頁。リポジトリには置かない）。領域1の運用正は DOMAIN1 の調査スプシ／スライド。本PDFはユーザーが言う **リサーチの元資料**。

PDFの骨格:

- 利益の一番大事な作業は **商品・メーカーリサーチ**（その後が問屋・見積・利益計算・発注）。
- 種類は3つ: ①セラーリサーチ ②オークファン（Keepa製品ファインダーで代替可）③ギフトショー（出展メーカー名をExcelに上げて1社ずつ）。
- オークファンの効率機能に **キーワード／ブランド／メーカー名** 検索がある。例: メーカー名欄に「まるか食品」。
- 結果はランキング順。Amazon本体「なし」を選べ、とある。
- 目安: おおむね10万位以内（カテゴリにより5万）、本体の有無、メーカー出品、出品者数は順位との相対、利益500円・粗利8%は **見積前の荒い計算**（ワカルンダ）。食品の掛率目安60–70%。
- **取りこぼし警告（最終頁）**: Amazonでメーカー名検索しても、ヒット商品が少ないことがある。例: 「まるか食品」では弱く、売れているのは「ペヤング」。メーカー公式サイトの品揃えを見て漏れを防ぐ。

今回の一本化案（キーワード＝メーカー名 → タイトル含有 → Keepaで属性）との対応:

| PDF | 今回 | 整合 |
|-----|------|------|
| 単位はメーカー1社 | メーカー1件 | 合う |
| ギフトショー→名前リスト→1社ずつ検索 | 入力がメーカー名 | 合う（ショー自体は領域1のAで後回し） |
| Keepa／オークファンで発見＋順位 | Keepaは属性（順位・価格） | 役割は残る。発見の主役をキーワードに移す案 |
| メーカー名欄（構造化フィールド） | Catalog の brandNames／Keepa製造者 | これはノーブランドが漏れやすい。PDFの「メーカー名検索」に近いが、石原で弱かった側 |
| Amazonキーワード（人が検索画面） | Catalog keywords＋タイトル含有 | **人がやっていた救いと同じ種類**。石原143一致 |
| 公式サイトで品揃え確認 | 未実装 | PDFはここを取りこぼし防止の本命にしている。キーワード一本化では **商品名にメーカー名が無いヒット商品（ペヤング型）が落ちる** |
| セラーリサーチ | **最終形の起点。第1版に乗せる**（sellerId） | PDFの①と一致。メーカー洗いはその下流 |
| 利益500円・8%・掛率 | 第1版の門では粗利を使わない（見積後） | 資料の「荒い事前計算」とは時期が違う。矛盾ではなく門の置き場が後ろ |

結論: PDFの3種類のうち、**①セラーが最終形の入口**、メーカー洗いはその後。発見の正はキーワード＋タイトル含有（石原型）。公式サイトとペヤング型は第1版に載せない（PDF警告は後続）。

### 第1版に載せない（底）

**ペヤング型・公式サイト照合。** タイトル含有では漏れる。別名辞書などは後の開発。セラーリサーチは載せない対象ではない。

---

## Keepa API vs CSV（2品番・石原水産の誤ダウンロード）

対象: Downloads `KeepaExport-2026-08-14石原水産.csv`（2行。リサーチ用143件とは別）。ASIN `B0FQB34W4B`／`B08W56WVHD`。  
スクリプト: `tools/purchase_research_path3/keepa_csv_vs_api.py`（既存 GAS と同じ `product?domain=5&stats=365&offers=20`）。②のシートは触らない。

**2026-08-14 再実行（列の有無ではなく値）**: `product?domain=5&stats=365&offers=20`。`tokensLeft=290`。  
詳細（85列×2ASIN）: `tools/purchase_research_path3/石原水産_Keepa値一致一覧.csv`  
列サマリ: `tools/purchase_research_path3/石原水産_Keepa値一致_列サマリ.csv`  
変換: 日本価格は円のまま、評価は÷10、寸法はmm→cm、EANは `eanList[0]`。

件数（170セル＝85列×2）: 一致83／ほぼ一致（順位時刻差）4／ほぼ一致（丸め）3／一致（セラーID）4／両方空29／CSVあり・API空40／CSV空・APIあり2／不一致5。

### 一致しているもの（単位変換後。門に使う列はここ）

- 商品名、ブランド、製造者、ASIN、EAN（eanList）
- カテゴリ ルート／サブ／ツリー、Amazon URL、Keepa URL
- レビュー評価（4.3）、評価件数
- Buy Box 30/90日平均（円）、90日下落%、FBAです（yes/no）、セラーID（CSVの店名付き表記の末尾ID）
- Amazon 90日在庫切れ（100%）、在庫切れカウント30/90
- Amazon本体価格・平均・下落%: **両方空**（本体なし。値として一致）
- 新品 90/365平均・90日下落%、第三者FBA 90/365平均
- 新品アイテム数 90日平均・90日下落%
- パッケージ 縦横高さ(cm)・重さ(g)、商品の重さ(g)
- 単位の価値、順位の減少回数 30/90/180日（`salesRankDrops*`）
- 商品ハイライト・ビジネス割引・Amazon在庫数・Buy Box標準偏差30日: **両方空**

### ほぼ一致（使える。時刻・丸め）

- 順位90/365日平均（例 12286↔12294、20857↔20850）
- 新品アイテム数365日平均（1↔0、7↔6）
- FBM 365日平均（617↔616、片方ASIN）

### 不一致（値が違う）

| 列 | 内容 |
|----|------|
| 画像 | URL文字列は一致しない（`+` のエンコード・枚数）。APIは `imagesCSV` から組み立てる。門ではASIN識別に使えば足りる |
| 売れ筋ランキング: 90日間の下落 % | CSV -9%/46% vs こちらが (avg90-avg30)/avg30 で計算した値は別物。画面の既算％はAPIにそのまま無い |
| 売れ筋ランキング: 過去365日間の減少 | 1ASINのみ 292↔291（1差。他方は一致） |

### CSVにあって今回のAPIフラットに出していないもの（画面の既算・期間シェア）

Buy Box の Amazon%、トップセラー%、勝者数、標準偏差90/365、変動性、定期おトク便、クーポン%。  
参考価格は1ASIN一致、もう1ASINはCSVが0でAPI空（-1扱い）。  
発売日はCSV空、APIは Keepa分（人がDLした列は空）。

**結論（値）**: 門に使う識別・順位平均・価格平均・レビュー・OOS・寸法重量・EANは、変換すれば人とDLと同じ値（順位は取得時刻で数〜十程度の差）。画面専用の下落％（順位）とBuy Box期間シェアはAPI生フィールドではない。

## Keepa API 全項目 vs DL85列（再検証・1ASIN実レスポンス）

対象ASIN `B08W56WVHD`。`product?domain=5&stats=365&offers=20`。キー一覧249行: `石原水産_KeepaAPI全項目.csv`。85列突合: `石原水産_KeepaDL85_vs_API全項目.csv`。`tokensLeft=295`。

| | 件数 |
|--|--|
| DL列 | 85（絞ったエクスポート） |
| このASINのAPIにキーとして見えた | 84／85（「画像」は `images[]` あり。判定漏れ） |
| リサーチ可能（yes/derive/build/offers） | **81** |
| 任意（薄くても門に不要） | 4（Amazon在庫数、定期おトク便、クーポン%、ビジネス割引） |

**ダウンロード項目があればリサーチ可能、でよい。** 必須85のうち門・候補行に使うものは API で揃う。画面CSV専用なのはセラー名（offersで代替）と既算の下落%（自分で割る）程度。APIの方が多い（履歴・offers・availabilityAmazon・fbaFees・eanList）。






