# Yahoo 商品説明テンプレ＋LOCK — 要件

**文書種別**: requirements  
**対象**: Yahoo! 出品・商品説明（caption）／フリースペース1・2／スマホ用フリースペース  
**状態**: **確定（案A'・2026-09-14）** — 手貼り廃止。GAS組立HTMLを editItem に**毎回同送**（省略禁止）。**完成形ロック 2026-09-15**（§0.1。P5は第1スライス＝仮本線）  
**関連**: `YahooCaptionRich.js`・`Yahoo.js`・`tools/yahoo_page_mock/`・`paste/HUMAN_RUN_YAHOO_CAPTION_RICH_API.md`・[B45 §7.2](B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)

---

## 0. 目的

検証で固めた HTML を、**マスタに巨大HTMLを直書きせず**、テンプレID＋短い変数から組み、**GASが editItem でYahooへ載せる**。

**必須方針**
- **自動出品コンセプト**（中身の生成も自動化へ。商品ファミリーごとにHTMLを人手量産しない）
- **届け口はGAS同送**（案A'）。生成結果を editItem に載せる
- **シリーズ同一HTML**（同一親の袋数SKUは同一生成結果）
- **人間のPro手貼りは本線にしない**
- **人手で型名・文案・第2テンプレを増やさない**（社長に型を選ばせない）
- **editItem は省略＝デフォルト上書き**（Yahoo公式）。装飾を守るには「送らない」ではなく**生成HTMLを同送**する

**層の分離（運用破綻防止）**

| 層 | 意味 | 人間の役割 |
|----|------|------------|
| 届け口 | 出品／APIで caption・FS を同送 | ほぼ無し（Property・検収） |
| 中身の生成 | 採用テーマの catch／overview ＋器。P6cで上位3枚 | **手で型・文案を増やすのは本線外** |
| 見本骨格 | かつおで固めたレイアウト | 初号の参考。永久に魚種ごと複製しない。**P5固定文は仮本線** |

---

## 0.1 完成形ロック（2026-09-15）

P5の `katsuo_odashi` 固定文＋T03/T04/T11 ファイル名差し替えは **第1スライス（仮本線）**。完成形ではない。実装は別承認。

### 正（完成形）

| 項目 | 内容 |
|------|------|
| 本文の正の入力 | マスタ **`楽天サブ画像1〜10`** にURLがある枚（当選）。themeId＋その画像＋**catch**＋**salesTalk／subs**＋**overview**。[ALIGN §7.2](B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 見出し | **catch**（画像内の主キャッチ） |
| 補足 | **salesTalk**（MAIN小）または **subs** |
| 紹介 | **overview**。captionでは **catch＋salesTalk／subs＋安全なマスタ事実**から機械が買う理由の文に組む。**ミックス**はセット構成の構成品商品名が2種以上のとき、各構成のマスタ安全欄をまとめて種にする（カテゴリ専用にしない。商品名の「・」だけでは分けない）。ミックス時は種に **構成種数＋構成品名** を載せ、**セット訴求は上位3のうち先頭1枠の overview にだけ1文**（2枠目以降は繰り返さない。固定文リスト禁止。単一品種の袋数セットは付けない）。紹介の種に**原料全文は載せない**。**内容量の数字は紹介に使わない**。写真用の構図overviewは種にしない。`商品説明文AIから取得` は使わない。overviewSp 空なら overview を使う |
| 【商品について】枠 | **廃止（案C・2026-10-27）**。名称・保存の枠は出さない。「こんな方に」はペルソナが説明パイプラインに未接続のため載せない（B見送り） |
| 原料 | マスタ列 **`原料`**。**食品だけ**。**`caption` のみ**（`sp_additional` には出さない＝二重表示防止）。複数成分は `/`・`／`・`【ラベル】` で改行＋1行空け。重複ラベル除去。GAS: `yahooCaptionRichApplyProductAboutC_` |
| 注記 | **残す**（`※原材料・アレルゲン…`／クーポン注記） |
| 内容量 | **出さない**（ミス源のため排除。紹介・多言語の Net weight 行も同方針） |
| 画と文 | **同一 themeId**。文だけ／画だけを別テーマから当てない |
| 載せる枚数 | **3**。AIが**当選サブの中だけ**を「売れそう」順に並べ **上位3件**。当選が3未満ならある分だけ（人が埋めない。落選・かつお固定画で埋めない） |
| 画像URL | Yahooホスト／lib のみ（楽天CDN直載せ禁止） |
| 人間 | 検収・Property・再実行のみ |

### 禁止

- 人手で型名・文案・第2テンプレを書く／「この型を1つ決めて」と社長に聞く
- `商品説明文AIから取得` を caption 本文の種にする（[§6.1 AI-DESC-FISH-MIX](#61-既知課題メモai情報取得改善候補)）
- 固定 T03/T04/T11 を完成形の枠契約にする
- caption 用にセールストークを人が新規作成する（サブ／MAIN作成時のものを使う）
- 「機械が選んだ分だけ載せる」（枚数不定）。正は **売れそう順の上位3**
- サブ落選・マスタ空欄を pool に入れて順位付けする（messages 全件・`_export_meta` 単独は禁止）
- 【商品について】や紹介・多言語へ **内容量の数字**を載せる（ミス防止。確定 2026-10-27）
- 非食品へ原料行を出す／列名を `主原料` に固定してマスタ列 `原料` を見ない

### 実装メモ（P6c-1・2026-09-16）

- 出口: `python tools/yahoo_page_mock/build_yahoo_caption_themes.py` → `tools/yahoo_page_mock/out/caption_themes/{jan}.json`
- **pool の正** = 出品マスタ `楽天サブ画像1〜10` の入っている枠。画像は `アップロード済み画像/ishi-{JAN}-oya_subN.jpg`（楽天CDN直載せ禁止）
- catch／overview は同一 themeId の messages（ファイル完全一致、または画像内キャッチ突合）。似た画の誤当てはしない
- マスタ空／oya無しはスキップ（固定 T03/T04/T11 で埋めない）
- AIが当選 pool を売れそう順に並べ **上位3**。3未満ならある分
- GAS: Property `YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED`（**未設定=false**。ONのときだけ caption／sp の固定文ブロックを上位3に置換）
- 画像URLは Yahoo lib のみ。新ファイル名 `{slug}_t{id}_{themeSlug}.jpg`（固定 T03/T04/T11 契約にしない）
- **紹介の組立**: 同一 themeId の catch＋salesTalk／subs＋安全事実。内容量の数字は種にしない・機械で削る（g/ml等。大容量は残す）。写真用overviewは種にしない。`商品説明文AIから取得` は禁止
- **ミックス分岐（汎用）**: 親の `セット構成1〜5_親SKU` → 各構成の商品名ベース。2種以上ならミックス。構成マスタの安全欄（商品名ベース・メーカー・特徴・**原料** など列があるもの）だけまとめ紹介の種へ。魚種などカテゴリ専用ロジックは置かない。`build_yahoo_caption_themes.py` の `compositionFacts`
- **ミックス・セット訴求（2026-10-27 確定・同日修正）**: 種に構成種数＋構成品名。**セット訴求は上位1枠のみ**（2枠目以降は機械で種数アピール文を除去）。固定決め文句リスト禁止。単一品種袋数セットは対象外。紹介種から原料全文は除外
- **原料（2026-10-27 確定・案C）**: 列はマスタ **`原料`**。食品のみ。**caption のみ**（spには出さない）。GAS: `yahooCaptionRichApplyProductAboutC_`
- **原料・複数成分の改行（2026-10-27）**: `/`・`／`・`【ラベル】` で2つ以上なら `<br><br>`。重複ラベル除去
- **【商品について】廃止（案C・2026-10-27）**: 枠・名称・保存を除去。※注記は残す。ペルソナ未接続のため「こんな方に」は載せない
- **内容量排除（2026-10-27 確定）**: 紹介・多言語 Net weight・旧商品について枠から内容量行／数字を出さない。GAS: `yahooCaptionRichStripContentQtyLines_`
- **文字数合わせ（PC・スマホ同じ仕様）**: 組んだ overview／subs を AI＋機械で枠内に収める。画・themeId は変えない。caption 5000／スマホ用説明 10000。PC は `overview`、スマホは `overviewSp`。GAS はまだ超える欄だけ紹介→補足の順で落とし、それでも超える欄は P5 固定文のまま送る
- 手順: [HUMAN_RUN_YAHOO_CAPTION_P6C_THEMES.md](../../tools/yahoo_page_mock/paste/HUMAN_RUN_YAHOO_CAPTION_P6C_THEMES.md)

---

## 1. 欄分担

| Pro欄 | APIキー | 内容 | 誰が書く |
|-------|---------|------|----------|
| 商品情報 | `explanation` | プレーン短文 | 通常出品（JOIN可） |
| ひと言コメント | `abstract` | 短文（JOIN可） | 通常出品（装飾HTMLではない） |
| 商品説明（PC） | `caption` | CTA・画像・セット | **テンプレ注入** |
| フリースペース1/2 | `additional1` / `additional2` | 多言語 | **テンプレ注入** |
| スマホFS | `sp_additional` | 日本語＋多言語 | **テンプレ注入** |

---

## 2. 仕組み（案A'）

| 層 | 内容 |
|----|------|
| A | 骨格正本: リポ＋`YahooCaptionRich.js`（初号ID `katsuo_odashi`＝食べるおだし系の見本） |
| B | 親行: テンプレID（任意）・画像ファイル名（任意）。未設定時はルールで骨格解決（かつおJAN等） |
| B' | **P5**: 同一骨格へ **マスタ由来の変数を自動埋込**（魚種名・内容量・セットリンク等） |
| B'' | **P6a**: 本文画像URLの**自動解決・差し込み**（列空でも）。UP自動化は P6b |
| C | メニュー「Yahoo説明をAPI反映」…初回／再投入用 |
| D | **通常「出品実行」**: Property ON かつテンプレ解決可 → caption/FS に生成HTMLを**必ず同送**（JOINは載せない） |

Property `YAHOO_CAPTION_RICH_LOCK_ENABLED=true` の意味（更新）:  
**「装飾を省略する」ではない。** 「テンプレがある商品は生成HTMLを同送する」扉。

テンプレ未解決の出品CKは、従来どおりマッピング送信（JOIN可）。

```
Property ON + テンプレ解決
  → 通常出品の editItem に caption/additional1/2/sp_additional = 生成HTML
Property OFF
  → 従来マッピングのみ
```

---

## 3. マスタ列

| 列 | 必須 |
|----|------|
| `▼マスタ(Yahoo説明テンプレID)` | 任意（空でもルール解決可。手で魚種ごとIDを増やさない） |
| `▼マスタ(Yahoo説明画像1..3)` | 任意（ファイル名。最優先） |
| `楽天サブ画像1..3`（既存） | P6a: 説明画像列が空のときの本文画像URLソース |
| 既存マスタ（名称・JAN・子SKU等） | P5の変数ソース（新規巨大HTML列は作らない） |
| `▼マスタ(Yahoo説明LOCK)` | 不要 |

---

## 4. ログ

- `[YAHOO_CAPTION_RICH] vars … fish=… net=… sets=… imgSrc=…`
- `[YAHOO_CAPTION_RICH] inject item_code=… tmpl=… keys=caption,additional1,…`
- `[YAHOO_CAPTION_RICH] no_template …`（Property ONだが未解決）
- 投入メニュー: `[YAHOO_CAPTION_RICH] editItem …`

---

## 5. フェーズ

| 順 | 内容 | 状態 |
|----|------|------|
| P0–P2 | 検証・列確定 | 完了 |
| P4 | 省略LOCK（誤り） | **廃止**（Yahoo仕様と矛盾） |
| P4b | API投入メニュー | 完了 |
| **P4c** | **案A': 通常出品でテンプレ同送** | **検収済（かつお1袋）** |
| **P3** | かつお全袋へ同一HTML（API反映） | **検収済**（`ycr_20260914_230551` 成功:10。人間Pro確認OK） |
| **P5** | **自動生成 第1スライス**: 共通骨格＋マスタ変数の自動埋込 | **検収済**（`ycr_20260915_021123` API 12/0。商品情報の魚種混在は AI-DESC-FISH-MIX） |
| **P6a** | **本文画像の自動解決・差し込み**（Yahoo安全URL／列／かつおlib。楽天CDN直載せは欠画で廃止） | **欠画FIX検収済** |
| **P6b** | 魚種専用画像を追加画像libへUP→自動差し込み | **lib API UP済・説明再反映待ち** |
| **P6c** | **完成形**: **サブ当選**からAIが売れそう順に上位3・同一themeIdで見出し／紹介／画像を組む。人手型増殖なし。PC・スマホは既存文だけ文字数合わせ | **P6c-1 当選限定＋文字数合わせ** |

---

## 6. 検収

**P4c（済・かつお1袋）**
- [x] 出品実行後も CTA／本採用サブが残る
- [x] ログ `inject … keys=caption,additional1,additional2,sp_additional`

**P3（シリーズ）**
- [x] 全袋SKUへ API反映（`ycr_20260914_230551` 成功:10／失敗:0）
- [x] 人間Pro確認で装飾OK（同一テンプレ）
- [ ] （任意）公開CDNで3袋・10袋の反映遅れがないか最終目視

**P5（自動生成 第1スライス）**
- [x] 手製第2テンプレなしで共通骨格＋変数（API `tmpl=katsuo_odashi`）
- [x] 魚種・内容量・セットリンクの自動埋込経路（画像一致は対象外）
- [x] まぐろ／ぶり1袋へ inject（`…5317-32s14-2`／`…7410-32s14-2`）＋かつお全袋再送 OK（`ycr_20260915_021123`）
- [x] 人間が魚種ごとにHTMLを手増殖しない
- [ ] （任意）Proで caption 側の魚種文言の最終目視メモ

**P6a（本文画像自動差し込み）**
- [x] API経路で画像URL解決ログあり（`imgSrc=`）
- [x] **楽天CDN直載せは欠画のため廃止**（skip→かつおlib）。Yahooホストのみ可
- [x] かつおlib復旧後、まぐろで欠画解消（`ycr_20260915_023659`＋人間Pro確認 2026-09-15）
- [x] Property `YAHOO_CAPTION_RICH_P6_IMG_ENABLED=false` で旧挙動に戻せる

**P6b（魚種lib）**
- [x] 命名 `{maguro|buri}_body_t03_sizzle|t04_aji|t11_arrange.jpg`
- [x] 素材prep: `prep_odashi_lib_bodies.py` → `out/lib_ready/`
- [x] GAS: lib存在チェック後に `imgSrc=fish_lib_{slug}` で差し込み（未UPはかつおlib）
- [x] **API UP試験OK**（`upload_lib_images_api.py` 2026-09-15: 6/0・公開URL 200）
- [ ] まぐろ／ぶり再API反映で魚種画像表示（欠画なし）
- [x] Pro手UPは本線にしない（公式 `uploadLibImage`）

手順: [HUMAN_RUN_YAHOO_CAPTION_P6B_LIB.md](../../tools/yahoo_page_mock/paste/HUMAN_RUN_YAHOO_CAPTION_P6B_LIB.md)

**関連・別課題（P5不合格ではない）**
- Yahoo **商品情報**（`explanation`）に「まぐろとぶり」など複合魚種が載る問題は、**AI情報取得→マスタ商品説明**側。テンプレ注入の対象外。改善は別タイミング（下記メモ）。

**P5実装メモ（2026-09-15）**
- 骨格IDは引き続き `katsuo_odashi`（別名 `odashi_gourmet`）。手増殖IDは作らない
- 解決: 親テンプレ列 or おだしJAN（5119/5317/7410）or 商品名 `食べるおだし（…）`
- 変数: 括弧内魚種・`1袋/○g`・兄弟子SKUのセット表（2袋以上）・画像列（任意）
- チーズかつお混在名は自動解決対象外
- 手順: [HUMAN_RUN_YAHOO_CAPTION_P5_VARS.md](../../tools/yahoo_page_mock/paste/HUMAN_RUN_YAHOO_CAPTION_P5_VARS.md)

**P6a 優先順位（画像）** ※第1スライスの現行。完成形は §0.1（売れそう順の上位3＋同一themeId）
1. `▼マスタ(Yahoo説明画像1..3)`（**libファイル名**、または **yimg.jp 等Yahooホストの絶対URL**）
2. **P6b** 魚種専用 lib（`maguro_body_*` / `buri_body_*`）※Pro追加画像にUP済みのときだけ
3. かつお本採用 lib 既定3枚
4. （不可）楽天サブ直載せ … 欠画

**P6a失敗メモ**: 「楽天サブをそのままcaptionへ」は**不合格**。表示できるのはYahoo側ホストの画像のみ。

**P5でやらない（明示）**
- まぐろ／ぶり専用の手製第2テンプレ量産

**P6でやらない（P6a時点）**
- AIで本文画像を新規生成（→ P6b/c）
- Pro「追加画像」への自動UP（→ P6b）

---

## 6.1 既知課題メモ（AI情報取得・改善候補）

| ID | 内容 | 状態 |
|----|------|------|
| AI-DESC-FISH-MIX | 単品SKUなのに `商品説明文AIから取得`／`▼マスタ(商品説明)` が「まぐろとぶり」等の**複合魚種**になる。Yahoo商品情報（explanation）へJOINされ誤表示。2026-09-15 まぐろ単品で実測 | **コード済**（2026-09-17）。`syncAiDataToMaster` は一意ASIN→一意商品名ベース＝AI商品名→一意JAN（メーカー両方非空なら一致必須。ブランドはJOINキーにしない）。生成プロンプトは1SKU固定。未設定=ON。caption 本文の種にはしない（§0.1）。要出品 clasp＋おだしAI再生成→同期→説明API |

---

## 7. 復元

- Property `false` → 従来マッピング
- テンプレ再生成: 骨格HTMLのみ `_gen_yahoo_caption_rich_js.py`（スニペット出力。本ファイル丸ごと上書き禁止）
- P5切戻し: Git revert（`YahooCaptionRich.js`／`Yahoo.js`）
- P6a切戻し: Property `YAHOO_CAPTION_RICH_P6_IMG_ENABLED=false`
- P6c切戻し: Property `YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED=false`（未設定＝OFF）
- AI-DESC-FISH-MIX切戻し: Property `AI_DESC_FISH_MIX_JOIN_ENABLED=false`（JAN後勝ち・検索結果最優先）／Git revert（`コード.js`）

---

## 8. 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-09-14 | 案A・手貼り廃止・投入メニュー |
| 2026-09-14 | **案A'**: 公式「省略＝デフォルト上書き」確定。通常出品はテンプレHTML同送に変更 |
| 2026-09-15 | **P5再定義**: 「手製第2テンプレ」廃止 → **共通骨格＋変数自動埋込（自動生成第1スライス）**。最終は自動出品。P6=AI／型選定 |
| 2026-09-15 | **P5実装**: `YahooCaptionRich.js` 変数埋込＋`Yahoo.js` 兄弟全件。まぐろ系はテンプレ列空でも解決。検収は HUMAN_RUN P5 |
| 2026-09-15 | **P5検収済**（API）。商品情報の魚種混在は **AI-DESC-FISH-MIX** としてAI情報取得バックログへ切離し。次は P6（画像自動） |
| 2026-09-15 | **P6a実装**: 本文画像＝説明画像列 → 楽天サブURL → かつおlib既定。Property `YAHOO_CAPTION_RICH_P6_IMG_ENABLED`。lib自動UPはP6b |
| 2026-09-15 | **P6a欠画FIX**: 楽天CDN直載せをやめ、Yahooホスト／lib既定のみ。まぐろ再反映で欠画解消→P6bへ |
| 2026-09-15 | **P6b**: まぐろ／ぶり本文3枚を lib 命名＋prep。GASはUP検知後に自動差し込み。Pro手UPは HUMAN_RUN P6b |
| 2026-09-15 | **P6b API UP試験**: `uploadLibImage` で maguro/buri 6枚 OK・公開200。`Yahoo.js` に `uploadLib` 追加。Pro手UP不要 |
| 2026-09-15 | **P6c-1**: 採用テーマJSON出口＋AI上位3＋GAS組込（`YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED` 未設定=OFF）。HUMAN_RUN P6c |
| 2026-09-16 | **P6c pool＝B-⑤当選のみ**: `_export_meta` の exports。落選は順位付けしない |
| 2026-10-26 | **P6c 同一JAN複数親**: パックキー＝親SKU。`oya2`／`oya3` のアップロード済み `…-oyaN_subN.jpg` を参照。旧 `…-oya_sub` 単体JANパックは廃止 |
| 2026-10-26 | **P6c ミックス分岐（汎用）**: 構成品商品名→各マスタ安全欄をまとめて紹介の種。カテゴリ専用にしない |
| 2026-10-27 | **原料＝列`原料`・食品のみ／内容量は説明から排除**: §0.1 確定 |
| 2026-10-27 | **実装**: `YahooCaptionRich.js`（内容量行除去・原料は食品のみ）。`build_yahoo_caption_themes.py` 安全欄に `原料`。要出品 clasp＋説明API |
| 2026-10-27 | **ミックス訴求＋原料改行**: 紹介に構成種数・品名由来のセット買理由。原料の複数成分は改行＋1行空け。要パック再生成＋出品 clasp＋説明API |
| 2026-10-27 | **セット訴求は先頭1枠／原料ラベル重複除去／紹介種から原料除外／overviewSp空はoverview**: 要 `--fit-only`＋出品 clasp＋説明API |
| 2026-10-27 | **案C**: 【商品について】枠廃止・注記残す・原料はcaptionのみ。要出品 clasp＋説明API |
| 2026-10-27 | **セットリンク親SKU固定**: 兄弟は親SKU。袋／個パース。骨格の単品JANリンクは置換必須（空でも消す）。要出品 clasp＋説明API |
| 2026-09-16 | **P6c pool＝マスタ楽天サブ画像1〜10**: 入っている枠が当選。`_export_meta` 単独は使わない |
| 2026-09-16 | **P6c 文字数合わせ**: PC・スマホ同じ仕様。既存文だけ短縮。caption 5000／スマホ 10000。人手文案なし |
| 2026-09-16 | **P6c 紹介組立**: catch＋salesTalk／subs＋安全事実。写真用overviewと商品説明文AIは種にしない |
| 2026-09-16 | **P6c 紹介から内容量数字を除外**: 単位不問。大容量は残す。機械削り |
| 2026-09-17 | **AI-DESC-FISH-MIX**: 一意キーJOIN＋説明プロンプトSKU固定。Property `AI_DESC_FISH_MIX_JOIN_ENABLED` 未設定=ON。caption 種は変えない |
