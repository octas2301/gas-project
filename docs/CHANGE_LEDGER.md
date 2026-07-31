# 変更台帳（復元用メモ）

コード全文ではなく **対象・目的・戻し方** のみ記録する（[AGENT_HANDOVER.md](AGENT_HANDOVER.md) §8・§9.1）。

| 日付 | 対象 | 目的 | 戻し方 |
|------|------|------|--------|
| 2026-07-31 | `c1_packaged.py`／`food_seasoning_column_map.json` | **SEASONING検索KWは1枠**（SC 99016: generic_keywordは最大1回。空白結合。keyword2〜5列マップ削除） | **Git**: 対象差分をrevert。旧5枠分割に戻る |
| 2026-07-31 | `AmazonDriveImageExport.js`／`コード.js` autoU4／U4 HUMAN_RUN | **楽天サブ→R2流用**（PT参照空かつ非AMAZON_ONLYなら `楽天サブ画像1〜8` を取得→R2→`Amazon PT URL`。D自動U4も楽天サブ不足を検知。MAINファイル無し時は既存MAIN URL維持可） | **Git**: 対象差分をrevert。autoU4判定は旧「ptRef無しならスキップ」に戻る |
| 2026-07-31 | `AmazonDriveImageExport.js`／`コード.js`（D新規）／`c1_packaged.py`／SEASONING列マップ／U4・C1 docs | **サブ画像＋商品の形式＋U4自動化**（U4はREUSE_RAKUTENでもPTをR2へ→`Amazon PT URL`（親へも伝播）、D新規はGENERATED前にU4自動実行、C1はPT URLを24〜31列へ、`item_form` 既定=粉末） | **Property**: `AMAZON_U4_AUTO_IN_D_ENABLED=false` でD自動U4のみ停止（21-⑦手動は従来どおり）。**Git**: 対象差分をrevert。列マップは `item_form`／`other_image1..8` の行を削れば旧挙動 |
| 2026-07-31 | `AmazonApprovalExport.js`／`コード.js` 21-⑮⑯⑰／D_ENTRY HUMAN_RUN §1e | **SC処理サマリ自動記録**（監視フォルダ＋時間主導トリガー→ファイル名から subBatchId→`UPLOADED_OK`／`UPLOAD_FAILED`→`_処理済`退避） | **Property**: `APPROVAL_AMAZON_LV4_SC_SUMMARY_ENABLED=false` で即停止。**21-⑰** でトリガー削除。**Git**: 当該差分revert＋メニュー3行削除。誤記録行はシート上で削除 |
| 2026-07-31 | `AmazonApprovalExport.js`／D_ENTRY HUMAN_RUN | **21-⑭ 証跡おすすめ**（同カテゴリマスタASIN→過去成功文面。OKで記録／Cancelで手入力） | **Git**: 当該差分revert |
| 2026-07-31 | `AmazonApprovalExport.js`／`コード.js` 21-⑭／Lv4正本／D_ENTRY HUMAN_RUN | **GTIN免除証跡の記録メニュー**（レ点カテゴリ検出→人間確認→EXEMPTION追記。既存証跡は追記せず、`*` はProperty＋警告） | **Git**: 対象差分をrevert＋メニュー行削除。誤記録した EXEMPTION 行はシート上で削除（追記のみのため他行に影響なし） |
| 2026-07-30 | `AmazonApprovalExport.js`／レ点本線承認包／Lv4正本／D_ENTRY HUMAN_RUN | **Dレ点新規は在庫>0でもGENERATED**（別カタログのため。承認①経路は従来どおりスキップ／マスタ在庫非改変） | **Property**: `APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK=false`。**Git**: 対象差分をrevert |
| 2026-07-30 | `コード.js`／`AmazonImageMatrixExport.js`／U2 docs | **Amazon画像の旧sheet誤紐付け防止**（U2時は楽天画像0件でもC再生成、②で現在レ点／sheet完全一致ゲート） | **Git**: 対象差分をrevert。`AMAZON_IMAGE_U2_ENABLED=false` |
| 2026-07-30 | `c1_packaged.py`／SEASONING列マップ・指紋／C1 docs | **C1-1c SEASONING**（新321列テンプレ、唐辛子ノード、行7保持・行8開始、PT別出力） | **Git**: 対象差分をrevert。HPCマップは維持 |
| 2026-07-30 | `AmazonApprovalExport.js`／`コード.js`／D_ENTRY docs | **同一レ点行を新規＋相乗りへ同時出品**（N列分割を廃止。識別子は子SKU／Amazon相乗りSKUで分離） | **Git**: revert。Property false |
| 2026-07-30 | D_ENTRY HUMAN_RUN／PHASE／HANDOVER／承認包 | **D相乗り自己発 dry_run／prod 実機合格記録**（`…48as12`／B084RJSH7W） | **Git**: revert docs。Propertyは作業後false |
| 2026-07-30 | `コード.js`／`AmazonApprovalExport.js`／`AmazonSpapiPut.js`／HUMAN_RUN | **D相乗り修正**（Dで自己発/FBA選択、X非依存、ASIN済みSKUそのまま、在庫>0でも送信0） | **Git**: 対象差分をrevert。Propertyをfalseへ |
| 2026-07-30 | `コード.js`／`AmazonApprovalExport.js`／`AmazonSpapiPut.js`／レ点本線docs | **Dレ点本線＋Amazon相乗りSKU実装**（新規／相乗り同時、N列ASINのみ、NF列VALID後保存、prod再利用） | **Git**: 対象差分をrevert。`APPROVAL_AMAZON_LV4_ENABLED=false`／`APPROVAL_AMAZON_SPAPI_PUT_ENABLED=false`。承認①メニューへ戻す |
| 2026-07-30 | レ点本線承認包／憲章／承認マトリクス／Lv4正本／商品マスタ要件／PHASE／HANDOVER | **3者多数決反映**（人間レ点＝当面承認①相当、A/M2専用Amazon相乗りSKU、フル＋相乗りprod許可）。コードなし | **Git**: revert docs。実装前のためProperty／GAS影響なし |
| 2026-07-29 | `コード.js` Dラジオ／offer facade／D_ENTRY HUMAN_RUN／PHASE | **A（D入口版）実装**（既存相乗りを D から呼出） | **Git**: revert。要 clasp push |
| 2026-07-29 | `LV4_SPAPI_D_ENTRY_APPROVAL` §6／PHASE／HANDOVER | **A（D入口版）社長承認**（実装可・コードなし） | **Git**: revert docs |
| 2026-07-29 | `LV4_SPAPI_D_ENTRY_APPROVAL`（新規）／PHASE／HANDOVER | **A（D入口版）承認起草**（Dラジオで新規／既存相乗り・コードなし） | **Git**: revert docs |
| 2026-07-29 | GAS PUT HUMAN_RUN／第2段承認／PHASE／HANDOVER | **v1.4 第2段 実機合格記録**（21-⑫ VALID／21-⑬ ACCEPTED） | **Git**: revert docs |
| 2026-07-29 | `AmazonSpapiPut.js`／`コード.js` 21-⑫⑬／GAS PUT docs | **v1.4 第2段実装**（承認①済→GAS直PUT） | **Git**: revert。Property false。clasp push 済 |
| 2026-07-29 | `LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL` §6／GAS PUT承認・HUMAN_RUN／PHASE／HANDOVER | **v1.4 第2段 社長承認**（実装可・コードなし） | **Git**: revert docs |
| 2026-07-29 | `LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL`（新規）／GAS PUT承認／PHASE | **v1.4 第2段 承認起草**（承認①済→GAS直PUT・コードなし） | **Git**: revert docs |
| 2026-07-29 | `AmazonSpapiPut.js`／GAS PUT docs／PHASE | **v1.4 API実機合格記録＋ENDPOINT `https:\` 正規化** | **Git**: revert。clasp push 済 |
| 2026-07-29 | `AmazonSpapiPut.js`／`コード.js` 21-⑩⑪／GAS PUT docs | **v1.4 GAS Listings直呼び実装** | **Git**: revert。Property false。要 clasp push |
| 2026-07-29 | `LV4_SPAPI_GAS_PUT_APPROVAL`／GAS PUT HUMAN_RUN／PHASE | **v1.4 承認起草**（コードなし） | **Git**: revert docs |
| 2026-07-29 | APPROVED／BRIDGE／CHECKBOX／CURRENT_PHASE／HANDOVER | **v1.2b実機合格＋親レ点完了扱い** | **Git**: revert docs |
| 2026-07-29 | CHECKBOX／BRIDGE／CURRENT_PHASE／HANDOVER | **v1.2c／v1.3 実機合格記録**（`…48s11` prod） | **Git**: revert docs |
| 2026-07-28 | `AmazonSpapiExport.js`／`コード.js` 21-⑧／CHECKBOX承認 | **v1.2c: 子SKUレ点のみ→Drive CSV（選択行廃止）** | **Git**: revert。Property false。要 clasp push |
| 2026-07-28 | `spapi_fetch_drive_csv.py`／listings_write／21-⑨／DRIVE・APPROVED承認 | **v1.3 Drive取得＋v1.2b承認一括＋v1.2a文字コード** | **Git**: revert。Property false。要 clasp push |
| 2026-07-28 | CURRENT_PHASE／HANDOVER／BRIDGE・WRITE HUMAN_RUN／承認包 | **橋渡し v1.2 実機合格を正本反映**（ride01 SC確認・Property false） | **Git**: revert docs |
| 2026-07-28 | `AmazonSpapiExport.js`／`コード.js` 21-⑧／BRIDGE docs | **スプシ→Drive CSV**（SP-API直呼びなし） | **Git**: revert。Property false。要 clasp push |
| 2026-07-28 | `spapi_listings_write` v1.1／BATCH承認／HUMAN_RUN | **複数行CSV**（max_items・安眠/アルコール相乗り1000/0） | **Git**: revert。items.csv・config.local は追跡しない |
| 2026-07-28 | `tools/spapi_listings_write/*`／WRITE HUMAN_RUN／承認包 | **Listings書込 v1**（1SKU offer／dry_run・prod） | **Git**: revert。config.local は追跡しない |
| 2026-07-28 | SPAPI SMOKE／CURRENT_PHASE／書込承認包 | **読取スモーク合格反映**＋書込は承認待ち docs | **Git**: revert docs |
| 2026-07-28 | `tools/spapi_smoke/*`／D_MENU_SPAPI_SMOKE_HUMAN_RUN | **SP-API読取スモーク**（LWA＋Catalog）。秘密は local | **Git**: revert。config.local は追跡しない |
| 2026-07-28 | `tools/m2_offer_packaged/m2_listing_loader_fill.py` 等／M2 docs | **公式 ListingLoader 自動埋め**（人手DL＋fill） | **Git**: revert 当該ファイル |
| 2026-07-28 | M2／CURRENT_PHASE／HANDOVER／GAP docs | **M2実機合格＋SP-API認証完了を正本反映**。公式LoaderがSC正 | **Git**: revert docs |
| 2026-07-27 | `コード.js`（E-3）／`ApprovalQueue.js`／E HUMAN_RUN | **E-3完了時**: 承認Web URL付きダイアログ（`approvalQueueBuildWebUrl_`） | **Git**: revert。要 clasp push |
| 2026-07-27 | `コード.js`（メニュー8）／AI_ADOPT docs | **v1.10**: 共有4列の類似横断dedupe（品/用・包含） | **Git**: revert。要 clasp push |
| 2026-07-27 | `コード.js`（`syncAiDataToMaster`）／商品マスタ作業エリア要件 | **キャッチ1行化**: 楽天／Yahooキャッチ作業エリアへは先頭1案のみ（全文転記禁止） | **Git**: revert。要 clasp push |
| 2026-07-27 | `コード.js`（メニュー8）／AI_ADOPT docs | **v1.9**: 区切り＝半角スペースのみ／メインKW1語／最終名は式の後ろ列・後ろワードから厳格trim | **Git**: revert。要 clasp push |
| 2026-07-27 | `AmazonApprovalExport.js`／`tools/m2_offer_packaged`／M2 docs | **M2 v1**: 案L CSV・競合ASIN・manual_ok。試験=発汗 | **Git**: revert。Property false／BRAND_GATE 削除 |
| 2026-07-27 | `docs/org/LV4_M2_*`／`D_MENU_M2_HUMAN_RUN`／CURRENT_PHASE／Facade U6 | **M2キックオフ**: ギャップ＋承認下書き（コードなし） | **Git**: revert docs |
| 2026-07-27 | `c1_packaged.py`／`config.example.json`／C1 HUMAN_RUN | **C1-clean**: `{subBatchId}`＋`--sub-batch`／`relax`固定廃止 | **Git**: revert `c1_packaged.py`。local は gitignore |
| 2026-07-27 | `tools/c1_hpc_packaged/c1_fetch_inputs.py`／要件／HUMAN_RUN | **C1入力案A**: Drive GENERATED＋マスタCSV自動取得（OAuth読取） | **Git**: revert。secrets は追跡しない |
| 2026-07-27 | `AmazonDriveImageExport.js`／`AmazonApprovalExport.js`／`コード.js`／E・C1 HUMAN_RUN | **親MAIN URL**: U4子→親コピー＋Lv4 Build子フォールバック | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`／`AmazonDriveImageExport.js`／`AmazonApprovalExport.js`／E HUMAN_RUN | **Eコース silent**: 成功時OKなし・エラー時のみ停止。U4/Lv4 silent | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`／`D_MENU_E_AMAZON_COURSE_HUMAN_RUN` | **メニューE**: Amazonコース（E-0〜E-5）薄いファサード | **Git**: revert。要 clasp push |
| 2026-07-26 | `Yahoo.js`／`コード.js`／STAGE | **ブランド正本=getShopBrandList**（検証＋メーカー名API取得） | **Git**: revert。要 clasp push |
| 2026-07-26 | `Yahoo.js`（SHP読取）／`コード.js`／YAHOO_CATEGORY_BRAND_STAGE | **正本=getShopCategoryList**: 書込前検証・無効ID却下／名前解決 | **Git**: revert。要 clasp push（Yahoo.js含む） |
| 2026-07-26 | `コード.js`（Yahooフォールバック）／YAHOO_CATEGORY_BRAND_STAGE | **競合ゼロ時**: AI推奨列→DriveマスタCSV選定 | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／YAHOO_CATEGORY_BRAND_STAGE／AI_ADOPT docs | **v1.8**: Yahooカテゴリ／ブランド都度API（売れ筋＋自社最安）をメニュー8へ | **Property**: `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED=false`。**Git**: revert。要 clasp push |
| 2026-07-26 | `YAHOO_CATEGORY_BRAND_STAGE`／承認包／AI_ADOPT | **Yahooカテゴリ選定改定**: 売れ筋＋自社最安優先（docsのみ） | **Git**: revert |
| 2026-07-26 | `YAHOO_CATEGORY_BRAND_STAGE`／承認包／HUMAN_RUN／AI_ADOPT docs／CURRENT_PHASE | **Yahooカテゴリ／ブランド Stage 要件起草**（都度API・`:`階層名・実装未） | **Git**: revert（コード差分なし） |
| 2026-07-26 | `コード.js`（メニュー8）／RAKUTEN_NAV_GENRE_STAGE3／AI_ADOPT docs | **Stage3**: 楽天ジャンル都度API（Ichiba投票→Nav）をメニュー8へ。AI推奨不使用 | **Property**: `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED=false`。**Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT docs | **v1.6**: バリエーションテーマ選定（HPCサイズ／食品パッケージサイズ＋Amazon裏マップ）をメニュー8へ | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT docs | **v1.5**: 下限=上限−5・半角0.5・容量1語・特徴/用途不適なら空 | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT 要件・HUMAN_RUN | **v1.4**: 楽天・Yahoo横断dedupe＋最終名上限（75/120/75）。検索KWから弱語削除 | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT 要件・HUMAN_RUN | **v1.3**: 最終商品名amazon式の左優先・完全一致横断dedupe（部分一致なし） | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT 要件・HUMAN_RUN | **v1.2**: 再生成停止。KW9選択＋dedupeのみ | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT 要件・HUMAN_RUN | **緊急**: syncAiDataToMaster をメニュー8から除外（レ点外破壊防止） | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー7.5・AI一括採用）／要件・HUMAN_RUN・承認包 | Amazon AI生成→空欄のみ採用→要確認。トグル既定false | **Property** OFF。**Git**: revert |
| 2026-07-26 | `コード.js`（MASTER_TO_WORKAREA／AI_COLUMN_MAP）／商品マスタ要件／RESEARCH／C1 HUMAN_RUN | 市場価格調査→定価転記を廃止。定価＝数値／式専用 | **Git**: revert。**clasp push** 反映後は再pushで戻す |
| 2026-07-26 | `tools/c1_hpc_packaged/**`（C1-1b）／HUMAN_RUN／要件 | マスタCSV併読・必須列。Propertyなし | **Git**: revert |
| 2026-07-26 | `D_MENU_C1_MASTER_HPC_COLUMN_MAP`／C1要件・HUMAN_RUN・CURRENT_PHASE | マスタ→HPC付け合わせ下書き。コードなし | **Git**: revert |
| 2026-07-26 | `tools/c1_hpc_packaged/**`／C1 HUMAN_RUN／承認包／`.claspignore` | C1実装承認＋ローカルPACKAGEDツール | **Git**: revert。成果xlsmはDownloads削除 |
| 2026-07-26 | `D_MENU_C1_THREE_REVIEW_MAJORITY`／C1要件／`LV4_C1_IMPLEMENTATION_APPROVAL` §6／CURRENT_PHASE 等 | C1三点＋社長決定反映。コードなし。次＝実装承認 | **Git**: revert |
| 2026-07-26 | `D_MENU_C1_…REQUIREMENTS`／`LV4_C1_IMPLEMENTATION_APPROVAL`／POC§5／CURRENT_PHASE 等 | C1方針ロック＋要件起草。コードなし。次＝3者 | **Git**: revert |
| 2026-07-26 | U4 HUMAN_RUN／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | **U4実機合格**記録（21-⑦・マスタURL） | **Git**: revert |
| 2026-07-26 | `AmazonDriveImageExport.js`／`AmazonApprovalExport.js`／`コード.js`／U4 HUMAN_RUN 等 | **U4 v1**: 21-⑦ R2→マスタURL＋GENERATED優先 | Property OFF。**Git**: revert |
| 2026-07-26 | `D_MENU_U4_…REQUIREMENTS`／`LV4_U4_IMPLEMENTATION_APPROVAL`／CURRENT_PHASE 等 | U4要件＋承認パッケージ起草。コードなし | **Git**: revert |
| 2026-07-26 | LV4_T2_HUMAN_RUN／CURRENT_PHASE／AGENT_HANDOVER | T2再検証合格（80s10・URL単独・18320なし）記録 | **Git**: revert |
| 2026-07-26 | `AmazonImageMatrixExport.js`／U2 HUMAN_RUN | ④成功MAIN/PTのみ `07/アップロード済み画像` へ退避 | Property `AMAZON_IMAGE_CANDIDATE_ARCHIVE_ENABLED=false` または git revert |
| 2026-07-25 | U2 HUMAN_RUN／CURRENT_PHASE／U2要件／AGENT_HANDOVER／CHANGE_LEDGER | **U2実機合格**記録（②③④・冪等D）。コード差分は同コミットの実装分 | **Git**: revert |
| 2026-07-25 | `コード.js`（`generateAiImageMatrix` 選定）／U2 HUMAN_RUN／CURRENT_PHASE | C: 子レ点優先（親のみレ点は全子のまま）。候補=`07` Property済 | **Git**: revert 該当差分 |
| 2026-07-25 | `AmazonImageMatrixExport.js`（新規）／`コード.js`（C-Amazonメニュー・フック）／U2 HUMAN_RUN／CURRENT_PHASE | U2 v1: sheet MAIN/PT・マスタ永続・02コピー。トグル既定false | **Property** OFF。**Git**: revert＋新規js削除 |
| 2026-07-25 | `D_MENU_U2_THREE_REVIEW_MAJORITY.md`（新規）／U2要件／POC／D要件／CURRENT_PHASE 等 | 三点採用＋社長回答反映。コードなし | **Git**: `git revert` |
| 2026-07-25 | `D_MENU_U2_…REQUIREMENTS.md`／D要件§7／CURRENT_PHASE | U2方針: 案α本線・MAIN=sheet紐付け・02=出口・εバックログ。コードなし | **Git**: `git revert` |
| 2026-07-25 | U3実機合格記録／`D_MENU_U2_…REQUIREMENTS.md`（新規）／CURRENT_PHASE 等 | U3クローズ＋U2起草。コード追加なし（本行） | **Git**: `git revert` |
| 2026-07-25 | `コード.js`（D amazon／full_amazon・`runBatchExportAmazonFacade`）／`AmazonApprovalExport.js`（21-①戻り値）／`D_MENU_U3_HUMAN_RUN.md`／要件・CURRENT_PHASE／AGENT_HANDOVER | U3 v1: 薄いファサード。トリガーにAmazon非搭載 | **Git**: `git revert`。Property OFF |
| 2026-07-24 | `D_MENU_AMAZON_FACADE_REQUIREMENTS.md`（新規）／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4・POCリンク | D本線Amazonファサード要件U0。T3保留。コードなし | **Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER（T2 PoC成功記録） | runId `R2T2_20260724_221107_7f9cf7`・URL画像表示。コード追加なし | **Property** `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`。**Git**: `git revert` |
| 2026-07-24 | `AmazonDriveImageExport.js` HMAC Byte[]揃え／clasp push | SigV4で (String,Byte[]) 例外を修正。T2のみ | **Git**: `git revert`／再push |
| 2026-07-24 | CURRENT_PHASE §0（clasp push済反映） | T2: `clasp push` 8 files 成功。コード差分なし（反映のみ） | 戻し不要（GASは前回リビジョンへ） |
| 2026-07-24 | CURRENT_PHASE §0 帰宅引き継ぎ／AGENT_HANDOVER | 自宅PC続行用に本日分を記録。コード追加なし（T2は既存） | **Git**: `git revert` |
| 2026-07-24 | `AmazonDriveImageExport.js`（新規）／`コード.js`（21-⑥）／LV4_T2_HUMAN_RUN／POC／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | T2: Drive→R2 MAIN1枚PoC。トグル既定false | **Property** `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`。**Git**: `git revert`／新規js削除＋メニュー削除 |
| 2026-07-24 | `RAKUTEN_NAV_STAGE1_HUMAN_RUN.md`／AI_ROUTING §5.1静的／`LV4_T2_IMPLEMENTATION_APPROVAL.md`／CURRENT_PHASE／AGENT_HANDOVER | Nav・AI静的＋T2承認書。GAS未実装 | **Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE／AGENT_HANDOVER／Drive `05` success JSON／FOOD_在庫確認_STATUS | FOOD v5成功（6バリエーション）。コードなし | **Git**: `git revert`／Drive success JSON削除可 |
| 2026-07-24 | Drive `03/…_v5_corrective.xlsm`・`04/…_MISSING5_MAIN_for_SC.zip`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | FOOD欠け5子corrective（親部分更新・30s175除外・袋数タイトル）。GASなし | **Drive**: v5ファイル削除。**Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE §0／LV4_R2_IMAGE_PIPELINE_POC（§2.1二モード）／AGENT_HANDOVER／CHANGE_LEDGER | 外出リモート引き継ぎ＋サブREUSE/ONLY方針。コードなし | **Git**: `git revert` |
| 2026-07-24 | Drive `04` 配下01〜06作成／POC §2 ID表／CURRENT_PHASE／`DRIVE_04_FOLDER_IDS.md`(Downloads) | フォルダ実体＋Folder ID控え。GAS・Property未設定 | **Drive**: 空フォルダ削除可。**Git**: `git revert` |
| 2026-07-24 | `LV4_R2_IMAGE_PIPELINE_POC.md`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4要件 | Drive起点GAS案を正に。xlsm自動は提案のみ。コードなし | **Git**: `git revert` |
| 2026-07-24 | `docs/org/LV4_R2_IMAGE_PIPELINE_POC.md`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4要件リンク | R2×Amazon画像PoC設計（楽天・Yahoo対象外）。コードなし | **Git**: `git revert`／PoC md削除 |
| 2026-07-23 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER | 帰宅引き継ぎ: HPC suburl試験・FOOD再UP禁止。コードなし | **Git**: `git revert` |
| 2026-07-23 | LV4要件 §11.0・§11表・§11.5.3画像／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | HPC §11-1〜8・10＋U5クローズ。画像ZIP優先。FOOD／M2／21-⑤は別ゲート。コードなし | **Git**: `git revert` |
| 2026-07-23 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER／accepted_values_db(Downloads) | 外出リモート引き継ぎ: HPC 21-③済・FOOD待機。コードなし | **Git**: `git revert`／DBはDownloads削除可 |
| 2026-07-22 | LV4要件 §11.5.4・§11.6／CURRENT_PHASE／`accepted_values_db/` | FOODテンプレ確認＋成功値DB方針。種子JSONはDownloads配下 | **Git**: `git revert`／DBフォルダ削除可 |
| 2026-07-22 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER | 外出引き継ぎ: corrective画像18320・中期R2方針。コードなし | **Git**: `git revert` |
| 2026-07-22 | LV4要件 §6.1.2・§10・§11.5.3／CURRENT_PHASE / AGENT_HANDOVER / CHANGE_LEDGER | 修正登録（SKU維持）・21-⑤・実データ7行目起点。GASメニュー実装は後続。ローカル corrective_link.xlsm | **Git**: `git revert`／PACKAGEDファイル削除 |
| 2026-07-21 | LV4要件 §11.5・§4／CURRENT_PHASE / AGENT_HANDOVER | HPC列対応ドラフト＋マスタ必須候補方針＋成功値辞書化メモ（後続）。コードなし | **Git**: `git revert` |
| 2026-07-21 | LV4要件 §5・§6.1.1・§14・§17 Q5／CURRENT_PHASE / CHANGE_LEDGER | subBatchId方針を新番号に統一＋push手順。コードは既存実装どおり | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js | 再レビュー採用: DRY_RUN／冪等汚染／subBatchId単調／ブランド厳密／archive失敗停止 | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js | 実装レビュー採用修正（レジューム・SKIPPED記録・GTIN・冪等・追記ログ）。戻し: git revert | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js（新規）／ApprovalQueue.js／コード.js（メニュー21）／LV4要件・CURRENT_PHASE / AGENT_HANDOVER | Lv4実装。ENABLED=false既定。戻し: Property OFF＋git revert＋新規js削除 | **Git**: `git revert` / Property `APPROVAL_AMAZON_LV4_ENABLED=false` |
| 2026-07-20 | LV4要件 §3.2・§2・§13・§17／AMAZON_REQUIREMENTS §3／多数決 §4.2／CURRENT_PHASE / AGENT_HANDOVER | Q15/Q16＋列メモ注記。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §3.4–3.6・§6・§17／多数決／CURRENT_PHASE / AGENT_HANDOVER | Q11–Q14反映。方針Q&A一通り閉じ。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §3.1–3.2・§3.5・§17／多数決 §4.1／CURRENT_PHASE / AGENT_HANDOVER | Q7–Q10b: 3モール同一承認・TRACK=B強制・SKU/メーカー品番。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §1.4・§17／多数決メモ §4.1／CURRENT_PHASE / AMAZON_REQUIREMENTS | 社長Q&A: D-1 PACKAGED・JAN出力規則・バリエーションのみ・再生成＋ログ。コードなし | **Git**: `git revert` |
| 2026-07-20 | `docs/org/LV4_THREE_REVIEW_MAJORITY.md`・`LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN / AMAZON_REQUIREMENTS | Lv4三点レビュー多数決保存＋社長確定反映（在庫書込禁止・DONE分離・親SKU・GTINゲート）。コードなし | **Git**: `git revert` |
| 2026-07-20 | `docs/org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN / AMAZON_REQUIREMENTS | Lv4要件ドラフト（M1=B新規・M2=A既存・手動UP・限定ツール）。コードなし | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-20 | CURRENT_PHASE / LV3要件 §8.1・§8.2 / AGENT_HANDOVER / LEVELLED_PLAN | Lv3人間検収完了を共有。フォーカスを Lv4 Amazon へ。在庫列二重ヘッダー運用メモ | **Git**: `git revert` |
| 2026-07-20 | `YahooApprovalExport.js`（resolveログ） | Lv3スキップ切り分け: sheet名・在庫生値・ORPHAN/IN_STOCK を Logger 出力 | **Git**: `git revert`／当該 Logger 行を削除 |
| 2026-07-20 | `YahooApprovalExport.js`（新規）・`ApprovalQueue.js`（Yahoo読取）・`コード.js`（メニュー20）・Lv3 docs | Lv3: 承認①済Yahoo子を案A＋主25分／副ユニーク50で `runYahooExport` 呼出。Yahoo本体非改変 | **Property**: `APPROVAL_YAHOO_LV3_ENABLED=false`。**Git**: `git revert`／新規js削除＋メニュー20削除 |
| 2026-07-20 | `docs/org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN | Lv3要件ドラフト（runYahooExport呼出・主25分／副ユニーク50運用揃え）。コードなし | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-20 | CURRENT_PHASE / LV2要件 §8.1・§10.1 / AGENT_HANDOVER / LEVELLED_PLAN | Lv2人間検収完了を共有。フォーカスを Lv3 Yahoo へ。目視手順追記 | **Git**: `git revert` |
| 2026-07-20 | `RakutenApprovalExport.js`（案Aレ点）・Lv2 docs | 案Aを「承認済み親＋紐づく子も一時レ点」に変更。CSV本体非改変。バリエーション親のシングルSKU誤認を解消 | **Git**: `git revert`。関数 `rakutenApprovalLv2ApplyPlanACheckboxes_` を親だけONに戻す |
| 2026-07-21 | `AmazonApprovalExport.js`（価格解決）・LV4 docs | 親`販売価格amazon`空→承認済み子へフォールバック。`[Lv4Price]`ログ | **Git**: `git revert` |
| 2026-07-21 | `AmazonApprovalExport.js`（カテゴリ解決）・LV4 docs | T列`カテゴリー`フォールバック＋見出しゆれ対応。`[Lv4Cat]`ログ | **Git**: `git revert`／当該関数を旧2行lookupに戻す |
| 2026-07-17 | CURRENT_PHASE / LV1要件 §8.1 / AGENT_HANDOVER / LEVELLED_PLAN / LV2前提 | Lv1人間検収完了をプロジェクト全体へ共有。フォーカスを Lv2 へ | **Git**: `git revert` |
| 2026-07-17 | `ApprovalQueue.js`・`Yahoo.js`(doGet分岐のみ)・`コード.js`(メニュー18)・`appsscript.json` | Lv1承認キュー実装。シート追記＋Web承認。出品API/CSV非呼出。Property既定false | **Property**: `APPROVAL_QUEUE_V1_ENABLED=false`。シート削除可。**Git**: `git revert` |
| 2026-07-17 | `docs/org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md` 等 | Lv1承認キュー要件（シート列・batchId・レ点抽出・Web・EC書込なし） | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-17 | org 多数決・Lvプラン・マトリクス・CURRENT_PHASE 等 | Lv0最終承認・モール順楽天先・レ点/スキップ/U1手動上書きの文書化（コードなし） | **Git**: `git revert` |
| 2026-07-17 | `docs/org/LEVELLED_IMPLEMENTATION_PLAN.md`、CURRENT_PHASE / AGENT_HANDOVER 追記 | AI組織の Lv0〜5＋並行P の実装順叩き台（コードなし） | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-17 | `コード.js`（`menuDiagnoseRakutenNavigationGenreStage1Write` 等）、`docs/RAKUTEN_NAV_GENRE_STAGE1.md` | Stage1: 専用シート `▼診断(楽天ジャンルNav)` への追記のみ（マスタ/CSV非書込・楽天CSV非改変）。Property 既定 false | **Property**: `RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED=false`。シート削除可。**Git**: `git revert`。メニュー17-⑥/99-⑩と Stage1 関数を削除 |
| 2026-07-17 | `docs/RAKUTEN_NAV_GENRE_STAGE1.md`（要件） | Stage0完了後の Stage1：テストシート追記隔離の要件正本化 | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-16 | `parseRakutenNavigationGenreResponse_` | NavigationAPI 2.0 の `nameJa` / `nameJaPath` を解釈して namePath を出す | **Git**: `git revert` |
| 2026-07-16 | `diagnoseRakutenNavigationGenreGetOne_` の URL | NavigationAPI 2.0 を `/es/2.0/navigation/genres/{id}` に修正（GF0002 404 対策） | **Git**: `git revert` または URL 配列を旧 query 形式に戻す |
| 2026-07-16 | `コード.js`（`menuDiagnoseRakutenNavigationGenreGet` 等）、`docs/RAKUTEN_NAV_GENRE_DIAG.md` | 楽天ジャンルID NavigationAPI 読取疎通診断（マスタ/CSV書込なし・CSV経路非改変） | **Property**: `RAKUTEN_NAV_GENRE_DIAG_ENABLED=false`。**Git**: `git revert`。診断関数・メニュー17-⑤/99-⑨を削除 |
| 2026-07-15 | `docs/org/PHASE0_THREE_REVIEW_MAJORITY.md`・`THREE_REVIEW_RUNBOOK.md`、憲章・マトリクス更新、`.cursor/rules/three-review-runbook.mdc`、CURRENT_PHASE / AGENT_HANDOVER | 3者多数決の記録・採用項反映・親1＋並列サブ3運用の正本化（実装コードなし） | **Git**: `git revert`。ルール削除は `three-review-runbook.mdc` を削除 |
| 2026-07-15 | `docs/org/AI_ORG_CHARTER.md`・`docs/org/AI_APPROVAL_MATRIX.md`、および `CURRENT_PHASE.md` / `AGENT_HANDOVER.md` の参照追記 | 一人社長＋AI部門の組織憲章・承認マトリクスを docs 正本化（実装なし。Phase0→3者検証前提） | **ファイル削除**: `docs/org/` の当該2ファイルを削除し、CURRENT_PHASE / AGENT_HANDOVER / 本台帳の該当行を戻す。**Git**: `git revert` |
| 2026-07-11 | `Yahoo.js`（`findYahooMasterHeaderRowIndex_`・`getYahooMasterHeaderContext_`・`_loadMasterData`・`updateMasterYahooId`・`updateMasterDeleteFlag`・`showDeleteSelectionDialog`・`listDeletableItems`） | マスタヘッダー行の `data[7]` 固定を廃止し、出品 Builder と同じ動的検出に統一（`docs/REVIEW_Yahoo_js.md` 指摘#1） | **ファイル**: `Yahoo.js.bak_before_header_unify_20260711` または `Yahoo.js.local_backup` を `Yahoo.js` に上書きコピー。**Git**: 本変更コミット後は `git revert` |
| 2026-05-09 | `コード.js`（`keepaProductPrimaryImageToken_`・`keepaAmazonMediaUrlFromImageToken_`・`getKeepaProductImageUrl`・`getKeepaProductImageUrlsAll`・`getKeepaProductImageUrlsMaxForCompetitorTest_`・`buildKeepaTableRow`） | Keepa API が返す **images 配列（l / m）** を解釈する。従来の image / imagesCSV のみでは現行レスポンスで画像URLが空になる問題を修正 | **Git**: `git revert`。**GAS**: 同ブロックを旧実装に戻す |
| 2026-05-06 | `コード.js`・`.claspignore` | R-Cabinet 診断ログ（パイプラインタグ・送信 bytes・JPEG 比・cap 付きサムネ）、Capacity 時に長辺 cap を段階的に下げる再試行、**`clasp push` で本番反映**。`.claspignore` に `**/*.html` 等で誤 push 防止 | **Git**: `git revert`。**GAS**: script.google.com で誤追加の `.html` を削除（以前の誤 push 分）。再試行を単発に戻す場合は `executeRenameAndUploadFromMatrixProgrammatic_` 内の `capList` ループを旧 1 回 `tryFetch` に戻す |
| 2026-05-06 | `RAKUTEN_CABINET_MAX_IMAGE_BYTES`・`convertBlobToJpegForRakuten_`・`prepareBlobForRakutenCabinetUpload_`・Drive サムネ系（`コード.js`） | R-Cabinet `Capacity` 回避のため目安を **1,900,000 バイト**に締め、サムネ／オリジナルとも **送信直前に JPEG 化**し **JPEG 後バイト**で上限判定する一本化 | **Git**: `git revert`。**手動**: 定数を `2*1024*1024` に戻し `convertBlobToJpegForRakuten_` を削除、サムネ戻り値の `convertBlobToJpegForRakuten_` 呼び出しと `prepare` の二段ロジックを旧実装に戻す（判断は `docs/BATCH_EXPORT_IMAGE_GATE_REQUIREMENTS.md`） |
| 2026-05-06 | `generateRakutenCSV`（`コード.js`）・`getRakutenSpecHtmlCell_` | PC/スマホ説明文の「商品情報」表を、▼マスタ列なしの明示フォールバック（メーカー名・マスタシリーズ／特記すべき原材料・原材料／括弧付き賞味・保存＋従来名）で取得 | **Git**: `git revert`。スペック4行を旧 `getMVal(ブランド/シリーズ/原材料/賞味/保存)` に戻し `getRakutenSpecHtmlCell_` 関数削除 |
| 2026-05-02 | `generateRakutenCSV`（`コード.js`）・ヘルパー `isMasterCellEmptyForRakutenAttr_` / `isRakutenAttrValueEmptyForCleaning_` | マルチSKUで親の `楽天セット数`（0/空）が `virtualRow`・`parentDataValues` 経由で子の属性値を潰し、かつ `attr_value` の `!val` で項目・単位だけ消える問題を防止 | **Git**: 該当コミットを `git revert`。**手動**: `virtualRow` のコピー・SKU属性継承・`attr_value` クリーニング3点を元に戻す（判断経緯は `docs/RAKUTEN_CSV_ATTR_AND_SETCOUNT_REQUIREMENTS.md`） |
| 2026-03-22 | `runProductNameProposalsForRows`（`コード.js`） | OpenAI 失敗時もマスタ商品名があればバリエーション（単位・内容量）を実行。Script Property `PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL`（既定 `true`）で旧挙動に戻せる | **Git**: 本変更のコミットを `git revert <commit>`。**運用**: Script Properties に `PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL` = `false` で旧挙動（OpenAI 失敗行はバリエーションもスキップ） |
| 2026-03-22 | `inferVariationFromAsinCircleForJan_`・`pickPerSetContentFromCircleTitles_` 等（`コード.js`） | 同一 JAN ブロック内の **全 ◎ 行**から 1セットあたり内容量を総合判断（g/ml 優先）。`CIRCLE_COMBINED_PER_SET_CONTENT`（既定 `true`）で無効化すると旧「先頭ヒット打切り」 | **Git**: `git revert <commit>`。**運用**: `CIRCLE_COMBINED_PER_SET_CONTENT` = `false` |

**コミットハッシュ**: ローカルで `git log -1 --oneline` を実行したうえで、上表に追記すること。
