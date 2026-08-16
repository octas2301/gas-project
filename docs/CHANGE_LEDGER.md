# 変更台帳（復元用メモ）

コード全文ではなく **対象・目的・戻し方** のみ記録する（[AGENT_HANDOVER.md](AGENT_HANDOVER.md) §8・§9.1）。

| 2026-08-16 | フェーズ1–5 page0 Keepa40 そうざい フロム四十五 | ①候補+21／Keepaフル+45／リスト無増 | Git revert |
| 2026-08-16 | リサーチ完成形フロー A–J / A′ を正本化 | FLOW.md 削除／DOMAIN1・V1 のポインタ戻し | docs Git revert |
| 2026-08-16 | L1–L6 一括（構成%削除・miss Keepa・JOIN・次セラー・みそ汁20） | 各塊の戻しは NOW | Git revert |

| 2026-08-16 | 処理時間 ▼実行時間（B Step＋Step5 APIのみ）。全関数ばら撒き禁止 | `OP_TIMING_ENABLED=false`。ルール `op-timing-scope.mdc`。要 clasp |
| 2026-08-16 | Step5 OpenAIはGemini失敗時のみ。AI梱包停止 | `B_STEP5_SKIP_OPENAI_IF_GEMINI_OK=false`／`B_STEP5_SKIP_PACK_AI=false`。要 clasp |
| 2026-08-16 | 行高さ B Step8 廃止（CLIP/setRowHeight削除） | Git revert。要 clasp |
| 2026-08-16 | A.準備 商品名SEOスペース（AIシートのみ） | メニュー3/4削除／Git revert。要 clasp |
| 2026-08-16 | B Step2.1＝12-⑭。ストアON時は横断シート直書きスキップ | `B_COMPETITOR_STORE_APPLY_ENABLED=false`／Git revert。要 clasp |
| 2026-08-16 | S4b セラー メイン＋構成％列（カテゴリ名） | 列削除 | Git revert |
| 2026-08-16 | S4 セラー構成比 storefront=0 | 4列クリア | Git revert |
| 2026-08-16 | S3 モリタ /query page0。台帳件数30 | 巡回日・件数を戻す | Git revert |
| 2026-08-16 | T2 品番リスト JOIN82 画像HYPERLINK | 当該セルをURLに戻す | Git revert |
| 2026-08-16 | A画像Gemini OFF検証は後続。今は実装なし | Property付けない | docs revert |
| 2026-08-16 | K8 BB/レビュー/BuyBox_FBA は今JSON空。代替はメモのみ | 実装なし | docs revert |
| 2026-08-16 | F2b 画像HYPERLINK／サブ画像\|（画像一覧リネーム） | 列クリア or ヘッダ戻し | Git revert |
| 2026-08-15 | F2 画像・画像一覧（生JSON images[]） | 列クリア | Git revert |
| 2026-08-15 | F1 Keepaフル カテゴリ・梱包・FBA手数料等を生JSON展開 | 当該列クリア | Git revert |
| 2026-08-15 | T1 品番リスト通過82追記（複製・式非上書き） | 末尾82行削除 | Git revert |
| 2026-08-15 | L1 出品FBA／自己発送 first-fit をKeepaフルへ | 5列削除 | Git revert |
| 2026-08-15 | F0 flatten 直販・新品現在・現行順位・oos180 | 列削除 | Git revert |
| 2026-08-15 | 品番リスト課題K1–K7。flattenはGETなし。L1は出品流用 | 実装なし | docs revert |
| 2026-08-15 | A.準備メニュー（Catalog空欄。12-⑮⑯同一） | メニューから項目削除 | Git revert。要 clasp |
| 2026-08-15 | P5 A末尾でP2並べ替え | `AMAZON_PASTE_P5_RANK_AFTER_A_ENABLED=false` | Git revert。要 clasp |
| 2026-08-15 | 出品者数＝COUNT_NEW。Keepaフル列化159。61空 | 列削除 | Git revert |
| 2026-08-15 | P1計画 dry（空ASIN0・Catalog GETなし） | 貼付非書 | スクリプト削除／docs revert |
| 2026-08-15 | P5計画 dry（P2並べ替え読取、A未結線） | 貼付非書。◎維持 | スクリプト削除／docs revert |
| 2026-08-15 | WRITE計画 dry（貼付 vs Keepaフル、非書） | Property OFF。raw-only フック | スクリプト削除／docs revert |
| 2026-08-15 | 競合DB「セラー」モリタ1行。フルから貯めない | タブ削除／行削除 | Git revert |
| 2026-08-15 | W4 hydrateは門stats空ならGET（T28） | 日付だけ新鮮な空JSONを載せない | Git revert。要 clasp |
| 2026-08-15 | ①タスク構造化メモをツリー正（NOW同文）に修正し他PJ共有 | 洗いと貼付を混ぜない。W4は出品 | docs Git revert |
| 2026-08-15 | ①タスク構造化メモを他PJ共有 | 洗いと貼付を混ぜない。W4は出品 | docs Git revert |
| 2026-08-15 | Keepa倉庫=Keepaフル。①候補は転記。W1が次 | Catalogはモールヒットに書かない | docs Git revert |
| 2026-08-15 | 領域1①進捗を NOW／DOMAIN1 §1.1／PHASE §0re に共有 | 他PJが洗いと貼付を混ぜない | docs Git revert |
| 2026-08-15 | P3 2種以上はマスタ非載 | ◎でもミックスを単品JANに載せない | Git revert。要 clasp |
| 2026-08-15 | 袋数 N種×各M／P2機械◎でも非候補／P3タイトル優先 | 人が直す前の機械精度。Keepaセルよりタイトル。A非改変 | Git revert。要 clasp（コード.js＋AmazonCompetitorPaste.js） |
| 2026-08-15 | P2 12-⑰⑱／paste_rank.py | 貼付非候補化。A非改変。評価◎維持 | Git revert。要 clasp。先に⑰ |
| 2026-08-15 | `purchase_research/` clone | ① GAS を Cursor で見えるようにした。出品 clasp 非変更 | フォルダ削除／出品 clasp は触らない |
| 2026-08-15 | `parseSetCountFromItemNameWithSource`／12-⑭外れ単価／ACCURACY.md | 計N袋優先。各N袋誤クラスタ防止。単価2倍は12-⑭で除外。課題を文書化 | Git revert。要 clasp push |
| 2026-08-15 | `コード.js` 12-⑭／`apply_to_master.py` | 専用ヒットを JAN＋セット数でマスタ競合列へ | Prop `COMPETITOR_MASTER_APPLY_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-15 | `コード.js` flush 指紋スキップ | 変化なしモールヒット非書込 | Git revert。要 clasp push |
| 2026-08-15 | FIELDS／STORE／NOW | 楽天ポイント＝％。ヒット列は意図的に絞る。領域1実装設計草案 | docs Git revert |
| 2026-08-15 | `docs/org/COMPETITOR_STORE_DEV_MAP.md` | 開発構造の一時メモ。完了後削除 | ファイル削除 |
| 2026-08-15 | `docs/org/COMPETITOR_STORE_REQUIREMENTS.md`／`tools/competitor_store/`／`コード.js` dual-write | 競合専用スプシ段階1。マスタKeepaは残す。ENABLED未設定=OFF | Property OFF／Git revert。要 clasp push（GAS分） |
| 2026-08-15 | `docs/org/COMPETITOR_FIELDS.md`／`competitor_fields/*.csv`／`.cursor/rules/competitor-fields.mdc` | 競合項目の正本化とPJ横断の参照ルール。コードなし | docs Git revert |
| 2026-08-15 | `docs/org/B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md` ほか | Amazon貼付自動化の制約（◎人間／候補列／キャッシュ＋領域1再利用）。コードなし | docs Git revert |
| 2026-08-14 | `コード.js` Keepaキャッシュ読取 | setCount空で旧形式誤判定しブランド等がログから消えるのを修正。ヘッダー名読取＋空上書き禁止 | Git revert。要 clasp push |
| 2026-08-14 | `コード.js` Keepa | 梱包欠落キャッシュはAPI再取得。`梱包_checked` で空返し再取得防止 | Git revert／`KEEPA_CACHE_REFETCH_WHEN_PKG_EMPTY=false`。要 clasp push |
| 2026-08-14 | `コード.js` Keepaキャッシュ | 梱包4列をキャッシュ保存＋ログ追記コピー。既存ログ非改変 | Git revert。要 clasp push |
| 2026-08-14 | `コード.js`／FOLLOWUP／PHASE | 親行CLIP後勝ち63px／選定不可はCT〜CV式維持／B 6.05改行／Yahoo query=JAN／楽天名フォールバック＋袋パース | Git revert。要 clasp push。OFF: `B_PARENT_ROW_HEIGHT_ENABLED` |
| 2026-08-15 | NOW ④の正 | `/query` 形をロック。sellerIds 可変 | docs Git revert |
| 2026-08-14 | `コード.js` `getKeepaCachedResults` | キャッシュのブランド／製造者をログ追記へコピー。既存ログ非改変 | Git revert。要 clasp push |
| 2026-08-14 | Keepa API 2品番 | CSVエクスポートと product API 比較。tokensLeft=278。キー非コミット | スクリプト削除／docs revert |
| 2026-08-14 | `docs/org/B_PURCHASE_RESEARCH_NOW.md`／`tools/purchase_research_path3/` | 領域1①を**経路3先行**に差し替え。差分PoC＋SP-APIキーワード読取骨格。GAS第1版なし | docs/tools Git revert |
| 2026-08-14 | `コード.js`／`B_RUN_20260814_FOLLOWUP`／PRICING_V1 | **B完走フォロー**: 寸法なし空欄・CPO非塗色・Step5当B clear・BにFBA P1b(6.55)・KW seen1本・親行CLIP60。⑥Nはコード非改修 | Git revert。要 clasp push。OFF: `B_FBA_P1B_WRITE_ENABLED`／`B_PARENT_ROW_HEIGHT_ENABLED` |
| 2026-08-14 | `docs/org/B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md` | 領域1①詳細稿。Keepa件数キャップ削除・上限300トークン | docs Git revert |
| 2026-08-14 | `docs/DOMAIN1_RESEARCH_PURCHASING.md` ほか | 領域1正本（①第1・C下書き・調査URL） | docs Git revert |
| 2026-08-14 | `points_logic`／`points_send` | restore の `--today` をカレンダー%に渡す | Git revert |
| 2026-08-14 | `コード.js` | **AI列名引き**: ASIN貼り付け式・参考画像・卸値。`原価税込`別名 | Git revert。要 clasp push |
| 2026-08-14 | DEALS HUMAN_RUN／PHASE | **taper 1段目 prod**: b=`185913020679`／s=`185917020679` DONE。運用マニュアル化を後続タスクにメモ | docs Git revert |
| 2026-08-14 | `コード.js`／`B_HARD_DEATH_SCOPE_*`／PHASE／WATCHDOG | **Bハード死対策**: 入口3・Step1切断・insertedProductsマージ・集合フィルタ・getUi安全・番犬非設置 | Prop トグルOFF／Git revert。要 clasp push。`B_WATCHDOG_ENABLED=false` は残してよい |
| 2026-08-14 | `コード.js` | Step1で **依頼日(D)** を1行目→親／2行目→子へ式コピー（Qより左の個別対応） | 当該関数削除／revert |
| 2026-08-14 | `_local_backup/pre_B_STEP1_TOP_INSERT_20260814_001210/`（gitignore）／CHANGE_LEDGER | **復元点**: 上挿入実装前の `コード.js`／Yahoo.js／appsscript／clasp をローカル退避。HEAD=`2f9ef6c` | フォルダから Copy-Item（RESTORE.md） |
| 2026-08-14 | `docs/org/B_STEP1_TOP_INSERT_*`（多数決／要件／HUMAN_RUN）／PHASE／HANDOVER | **上挿入要件ロック**: 1A/2A/3A・既存値固定・テンプレ1–2。コード未 | docs Git revert |
| 2026-08-13 | `docs/org/B_STEP1_TOP_INSERT_REQUIREMENTS.md`／PHASE／HANDOVER | **上挿入改修メモ**初版（式方針は08-14で更新） | docs Git revert |
| 2026-08-13 | `コード.js` masterColorApplyLocked*／15-㉑㉒／色 docs §11–13／スプシ11–775 | **⑧f 列ロック確定**。親赤白薄赤。Step1自動塗 | Git revert。スプシ色は再塗or手戻し。要 clasp push |
| 2026-08-13 | `コード.js` MASTER_COLOR／CPO・Step1・カテゴリ等／色 docs | **⑧c 役割色実装**。オレンジ後回し。条件付き書式非変更 | Git revert。要 clasp push |
| 2026-08-13 | `docs/org/B_MASTER_CELL_COLOR_RULES_*`／PHASE／HANDOVER | **⑧a マスタ色ルール正本**（灰・オレンジ含む）。コード一斉塗なし | docs Git revert |
| 2026-08-13 | `コード.js` B Step8／WATCHDOG §7.1／PHASE／HANDOVER | **⑪親行高さ60px**。⑨改修不要・⑩運用検証メモ | Prop `B_PARENT_ROW_HEIGHT_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` Step1／17-⑦／MASTER_LINKAGE／RESEARCH／PHASE | **⑦ DF割引式**: 先頭子より下へ DF2 `ROUND` コピー。DDは対象外 | Prop `B_DF_STRATEGY_FORMULA_COPY_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑳／承認包／B_SPAPI docs | **⑤競合空のみ自動**: U3高信頼→競合列。U7本線当面不要を明記 | Prop `B_COMP_ASIN_AUTOFILL_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑲／B_SPAPI docs／⑤承認包 | **U7診断**: Catalog型番 vs Keepa。マスタ非書込。⑤競合列は承認包のみ | Prop `B_U7_PART_DIAG_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | CURRENT_PHASE／DEALS HUMAN_RUN／HANDOVER | **8/14前 taper dry_run**: 当日0・Cinderellas2SKU予定。FBA完了→減衰へフォーカス切替 | docs Git revert |
| 2026-08-13 | B_SPAPI要件／HUMAN_RUN／PHASE／HANDOVER | **FBAゲート暫定OK**: 必須=HTTP+梱包。型番/JAN→U7。U7未着手 | docs Git revert |
| 2026-08-13 | B_SPAPI HUMAN_RUN／PHASE／HANDOVER | **P1b実機** `P1b_20260813_162652_6111aa`: ティア/手数料 11/11書込。U7は型番JAN方針待ち | docs Git revert。マスタ値は手戻し |
| 2026-08-13 | マスタ出品CK／B_SPAPI HUMAN_RUN／PHASE | **P1a `595a71f2`集計**: HTTP20/20・梱包11/20。生姜湯2ASINレ点OFF→梱包あり6ASINのみON（P1b用） | レ点は手で戻す。docsは Git revert |
| 2026-08-13 | `コード.js` B統合 | **U6**: Step**6.6** メーカー品番（U5 quiet）。`B_INTEGRATED_STEP_FUNCTIONS` | Prop `B_MAKER_MODEL_FETCH_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` U5 | **U5自社品番**: `INT-`＋8桁hex（計12）。旧20字から短縮 | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` U5 | **U5候補ガード**: JAN同一却下／純数字は4桁未満のみ却下（16100可） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑤⑥ | **U5メーカー**: `メーカー名ベース`のみ（`メーカー名`不使用） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑥ | **U5結果シート**: setValues行数不一致修正（getRange第3=行数） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑤⑥ | **U5書込先**: **`メーカー品番下書き`のみ**（メーカー型番等NG） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑤⑥ | **U5列名**: `メーカー品番`無し→`メーカー品番下書き`／`メーカー型番`フォールバック | 次行で下書きのみに訂正 |
| 2026-08-13 | `コード.js` 15-⑤⑥／B_SPAPI docs | **U5 メーカー品番**: Keepa→Serp→INT-自社（黄）。型番不明廃止。空のみ | Prop `B_MAKER_MODEL_FETCH_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | マスタ列／`コード.js` 15-⑱／B_SPAPI docs | **FBA P1b**: `FBAティア`／`FBA手数料_円` 空のみ書込（自己発列非改変） | Prop `B_FBA_P1B_WRITE_ENABLED=false`／Git revert。列は手削除可 |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4f**: brandのみ照合・セット親×Catalog単品ヒント（0s1） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／`AmazonCategoryPt.js`／B_SPAPI docs | **⑤ U3.4e**: brand/製造者のみ照合・タイトル除外・unknown非推奨 | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4d**: U3メーカー正＝`メーカー名ベース`のみ | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4c**: maker一般照合（カナ↔ローマ字＋Catalog brand。別名表なし） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4b**: brand yes Catalog優先（0s1）／conflictはmakerInTitle／兄弟は他親のみ | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4a**: Catalogタイトル bi-gram≥0.80＋LCS。P4b関数は不変更 | Prop `B_KEEPA_ASIN_VOTE_U3_TITLE_SCORE`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.3**: 同一JAN同セットの兄弟競合票＋Catalogタイトル不一致抑制 | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.2**: ブランド一致・Catalogヒント・競合404・set_mismatch を推奨に反映 | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.1b**: 親`A.セット商品数`空→同一親SKUの子から継承（1優先／min） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／`AmazonCategoryPt.js`／B_SPAPI docs | **⑤ U3.1**: 親セット数ゲート・setGuess/ブランド一致列・Catalog票は単品親のみ | Prop `B_KEEPA_ASIN_VOTE_U3_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑯／B_SPAPI docs | **⑤ U3**: 競合ASIN投票診断（◎／N列／競合店／Catalog JAN→`競合ASIN投票診断_U3`）。マスタ非書込 | Prop `B_KEEPA_ASIN_VOTE_U3_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 12-⑧／B_SPAPI docs | **⑤ U2**: Keepa取得_ログ→archive（90日超 or JAN最新3実行以外）。マスタ非書込 | Prop `B_KEEPA_LOG_ARCHIVE_ENABLED=false`／archiveから手戻し／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` Keepaログ／B_SPAPI docs | **⑤ U1**: Keepa取得_ログに partNumber／ブランド／製造者／梱包寸法・重量（追加APIなし） | Prop `B_KEEPA_LOG_EXT_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `00_設定マスタ` FBA手数料／`update_fba_fee_table_00.py`／承認包 | **FBA手数料を公式現行（1,000円超）へ更新**。FBA末尾に行挿入。販売手数料等は非改変 | スプシ版履歴／承認包旧値 |
| 2026-08-13 | `コード.js` FBA P1a | **設定マスタFBA手数料連動**: 仮ティア＝区分名／手数料円。F備考をbox・三辺和でパース | 要 clasp push／Git revert |
| 2026-08-13 | `コード.js` `fbaP1aLengthToCm_` | **FBA P1a単位バグ修正**: `centimeters` が meters 誤判定→×100。cmを先判定・mは厳密一致 | 要 clasp push／Git revert |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **FBA P1a**: Z 15-⑰ Catalog寸法→`FBAティア診断_P1a`・仮ティア。マスタサイズ列は書かない | Prop `B_FBA_P1A_DIAG_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 3Dフィット | **rigid優先**（softは rigid 全滅後）。Compact stagger 当面維持 | 要 clasp push |
| 2026-08-13 | `コード.js` 3Dフィット | **ずらしバグ修正**（半辺で誤OK）。`[3D][try]` で内寸・soft・bestCap・ok をログ | 要 clasp push |
| 2026-08-13 | `コード.js` Step3.1 | **AI梱包寸法**: 商品名＋卸値(税込)優先／JAN+卸値／JANフォールバック。同JAN複数行で誤寸法を防止 | 要 clasp push |
| 2026-08-13 | `コード.js` Step3.1 | **サイズ昇順 first-fit**。送料=自己発送／箱代=資材D／合計=利益のみ。寸法あり不適合は空（ネコポス埋め禁止） | 要 clasp push。Git revert |
| 2026-08-13 | `コード.js` Step3.1 | **ネコポス利益上書き廃止**。3D時はサイズのみで箱選定。利益不足はフラグのみ | 要 clasp push。Git revert |
| 2026-08-13 | `コード.js` Step3.1／設定マスタ exclude／B_SPAPI docs | **④自己発3D**: 内寸フィット・exclude除外。`B_LOGISTICS_USE_3D_FIT` | Prop `false`／Git revert。要 clasp push |
| 2026-08-13 | `00_設定マスタ` E〜H／`write_box_inner_dims.py`／B_SPAPI_RESEARCH docs | **④⑤⑥要件＋ダンボール内寸列** | 列E〜H手削除／docs revert |
| 2026-08-12 | `コード.js` Step5/6／SERIES_INFER docs | **シリーズ不明禁止＋商品名推定（方針A）** | Prop `SERIES_INFER_FROM_NAME_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-12 | AI_ADOPT 要件／HUMAN_RUN | **商品名ベース可変不可**を要件明文化（読取のみ） | docs revert |
| 2026-08-12 | `コード.js` メニュー8／Yahoo Stage／E・YAHOO docs | **7.5維持／7.6 popular_only＋38074。B Step7.6差し替え** | Git revert。要 clasp push |
| 2026-08-12 | `コード.js` Step5／メニュー8／E docs | **E本線化**: ジャンル／Yahoo未設定=ON＋Step5ジャンルAI・Drive候補スキップ | 各Prop `false`／`B_STEP5_SKIP_GENRE_YAHOO_AI_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-12 | `コード.js` Step6.5／15-⑮／B_ASIN_N docs | **⑥ N列ASIN自動**: ◎×ブランド＝メーカー→空Nのみ＋黄セル要確認 | **Property** `B_ASIN_N_AUTO_FILL_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-12 | `コード.js` メニュー8／AI_ADOPT docs | **v1.13**: FO二重計上解消・商品名非削り・表示文字数trim | Git revert。要 clasp push |
| 2026-08-12 | `コード.js` メニュー8／AI_ADOPT docs | **v1.12**: 短すぎ対策（FO非削り・下限ブレーキ・特徴用途を空にしない） | Git revert。要 clasp push |
| 2026-08-12 | `コード.js` メニュー8／AI_ADOPT docs／B Step7.5 | **v1.11**: =LEN()厳格trim・FO先頭・B統合載せ | **Property** `B_INTEGRATED_MENU8_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-12 | `コード.js` 15-⑬⑭／要件 | **P2-A**: 月次メーカー辞書（手修正維持）＋保存方法改行→読点 | Git revert。要 clasp push |
| 2026-08-12 | `コード.js`／設定マスタA181／AI G列／要件 | **B P1**: シリーズ揃え・メーカー辞書・カタログスラッグ13 | Property各`false`／Git revert。要 clasp push |
| 2026-08-12 | `コード.js`／`B_WATCHDOG_*`／要件 | **B番犬P0**: ハード死後自動再開・Step5行CP・プルダウンTIME_SLICE・サマリ・進捗メニュー | **Property** `B_WATCHDOG_ENABLED=false` または **Git** revert。要 clasp push |
| 2026-08-12 | DEALS／HUMAN_RUN／HANDOVER | **表記揃え**: P1c実装済。G8＝終着空でもapply可。復元済＝減衰中%へ | **Git**: revert 当該docs |
| 2026-08-12 | `price_recovery_logic`／`taper_send`／`points_*`／GAS／DEALS§10.13-14 | **2層**: カレンダー減衰中%＋B中1%オーバーレイ。restore=減衰中%。fetchは現在%のみ。列`減衰開始日` | **Git**: revert。`--schema-only`＋clasp push |
| 2026-08-12 | `sheets_io` 次回減衰後%数式 | **f-string 二重引用で `{c_pct}` が残るバグ修正** | **Git**: revert。`--schema-only` 再実行 |
| 2026-08-12 | 広告運用GAS `onOpen`／HUMAN_RUN／DEALS | **メニュー分割**: タイムセール本線＋99。広告運用は分析のみ | **Git**: revert。要 clasp push |
| 2026-08-12 | `taper_send.py`／sheet_schema／GAS／DEALS§10.14 | **P1c**: taper＋メール＋実行依頼E。リカバリ手動 | **Git**: revert。要 clasp push／`--schema-only`。日次タスクは人手 |
| 2026-08-12 | DEALS§10.14／HUMAN_RUN／PHASE／HANDOVER | **P1c要件**: 現スナップ＋進捗。E先行。C／D／URL後続 | **Git**: 当該docs revert |
| 2026-08-12 | `sheet_schema` 減衰*改名／販促売価削除／GAS／DEALS | **戻し→減衰列名。減衰段%明示。販促売価円削除** | **Git**: revert。`--schema-only`＋clasp push |
| 2026-08-12 | `sheets_io` display formulas／GAS／DEALS§2.2 | **販促ポイント円・実質価格円をシート数式化**（販促売価円は非推奨で空） | **Git**: revert。`--schema-only` 再実行可 |
| 2026-08-12 | `sheet_schema`／sync realign／GAS Realign／DEALS§2.2 | **マスタ列：目的グループ＋人入力先頭＋黄セル** | **Git**: revert。要 clasp push／`--schema-only` またはメニュー並べ替え |
| 2026-08-12 | `price_recovery_*`／`sheet_schema`／GAS PriceRecovery／DEALS§10.12 | **実質戻し**: 販促ポイント%＋円表示。売価段上げ廃止 | **Git**: revert。要 clasp push（広告）・`sync --schema-only` |
| 2026-08-12 | `docs/org/MEMO_POINTS_VS_REFERENCE_PRICE.md` | **ポイント×過去価格メモ**: 記事要約・公式原文未特定・30日/90日/8週整理 | **ファイル削除**または revert |
| 2026-08-11 | `price_recovery_*`／GAS QtyMail／DEALS§10.12 | **価格戻しAPI＋G10**: our_price段階上げ・B/restoreガード | **Git**: revert。要 clasp push（広告） |
| 2026-08-11 | `TimeSaleQtyMail.js`／`コード.js`／DEALS docs | **Points G5**: apply／restore Cursor指示メニュー | **Git**: revert。要 clasp push（広告） |
| 2026-08-11 | `mail_points_remind`／HUMAN_RUN | **Points運用予行**: 日程基準リマインド＋Smile下書き確認 | **Git**: revert |
| 2026-08-11 | `points_fetch`／`test_points_fetch`／要件§10.10 | **Points G9**: listings pointsNumber÷price で%本線 | **Git**: revert |
| 2026-08-11 | `points_send`／`points_fetch`／要件§10.10 | **Points G8**: セール前%空・backup失敗でapply中止 | **Git**: revert。`--allow-missing-before` |
| 2026-08-11 | `points_logic`／`points_send`／要件§10.10.1c | **Points G4**: 施策連動SKUフィルタ（既定ON） | **Git**: revert。`--all-master`で旧挙動 |
| 2026-08-11 | `mail_points_remind`／GAS QtyMail／要件§10.10.1b | **Points G2/G3**: apply/restore リマインド | **Git**: revert。要 clasp push（広告） |
| 2026-08-11 | `points_logic`／send／fetch／要件§10.10 | **Points G7**: 状態語彙固定＋sheet更新揃え | **Git**: revert |
| 2026-08-11 | `points_send`／HUMAN_RUN | **Points G1**: 1SKU prod 2%→1%（feed DONE）。読取GETは未 | **Git**: revert。%はマスタで戻し可 |
| 2026-08-11 | `build_submit_xlsx`／HUMAN_RUN | **成功submitv2準拠**: compact+日付直書き+保護OFF（全行残しはSC処理失敗） | **Git**: revert |
| 2026-08-11 | `build_submit_xlsx`／HUMAN_RUN | **開始/終了はVLOOKUP残す**（参加中＋スケジュールのみ書込。日付直書きはSC失敗） | **Git**: revert |
| 2026-08-11 | `build_submit_xlsx`／HUMAN_RUN | **提出xlsx日付=YYYY-MM-DD文字**（データ定義準拠・シリアル混入修正） | **Git**: revert |
| 2026-08-11 | sync／template_parse／notify_mail／要件§10.9 | **②ドロップダウン日付付き全登録**＋公式衝突時独自短縮メール | **Git**: revert |
| 2026-08-11 | `mail_qty_confirm`／GAS `TimeSaleQtyMail`／HUMAN_RUN§9.7 | **数量確認メール実送信**（Gmail API＋下書きシート＋メニュー）→ contact@octas2301.com | **Git**: revert／token_gmail_send は secrets 外 |
| 2026-08-11 | HUMAN_RUN／`mail_qty_confirm`／DEALS§9.7 | **A未運用表記揃え**／数量確認はUL済対象＋`--tol` | **Git**: revert |
| 2026-08-11 | `parse_ymd`／`invariants.py`／sync／DEALS§0.1 | **並び修復＋不変条件自動検証**（抜け落ち再発防止） | **Git**: revert |
| 2026-08-11 | sync／build／sheets_io／DEALS§2.1.0 | **A未運用・作成=UL済・画像保全・台帳終了まで** | **Git**: revert |
| 2026-08-11 | sync／schedule_class／DEALS§2.1.1 | **B台帳保全**（予定消禁止）／Smile復元・BF非自動 | **Git**: revert |
| 2026-08-11 | `schedule_class`／sync／build／`TimeSaleP1bMenus.js`／DEALS docs | **P1b: ②SKU行正本・カスタム提出・UL固定文・再build** | **Git**: revert。要 clasp push（広告運用） |
| 2026-08-11 | `TimeSaleP1bMenus.js`／コード.jsメニュー／DEALS§10.9.1・P1b HUMAN | **P1b: ②確認→Cursor→③→UL手順メニュー**。公式即登録 | **Git**: revert。要 clasp push（広告運用） |
| 2026-08-11 | `a_track.py`／`lane_a_send`／sheet_schema／build `--sku`／DEALS docs | **A実施フラグ＋ログ参照**／1SKU提出 | **Git**: revert。列は手で残して可 |
| 2026-08-11 | `sheet_schema`／`sheets_io`／`TimeSalePriceRecovery.js`／DEALS§2.2 | **マスタヘッダ色分け**（戻し=琥珀） | **Git**: revert。シート背景は手でクリア可。要 clasp push（広告運用） |
| 2026-08-11 | `TimeSalePriceRecovery.js`／`price_recovery_logic`／sheet_schema／DEALS§10.12 | **戻し列＋GAS都度提案（即書き込み）** | **Git**: revert。要 clasp push（広告運用） |
| 2026-08-11 | `points_fetch`／`mail_qty_confirm`／`detect_new_schedules`／ops／DEALS§9.7・10.10-11 | **ポイントAPI退避・数量確認メール(スプシリンク)・新Sale差分** | **Git**: revert |
| 2026-08-11 | `points_*`／`sheet_schema`／DEALS§10.10 | **ポイント列意味**: 期間中／セール前退避／restoreモード | **Git**: revert。シート列は手で戻し可 |
| 2026-08-11 | `points_logic`／`points_send`／`sheet_schema`／sync／DEALS要件§10.10 | **Points Phase0**: マスタ列＋差分TSV／フィード。Phase1はメモのみ | **Git**: revert。マスタ列は手で残して可 |
| 2026-08-11 | `ops_status`／`revise_b_qty`／`lane_a_patch` | 運用ダッシュ・数量改定・AパッチJSON | **Git**: revert |
| 2026-08-11 | `tools/amazon_deals_bulk/*`／広告GAS／DEALS要件§2 | **シート改訂**: マスタ薄型・実行・A別タブ。sync＋P1b | **Git**: revert。シートは手で戻し可 |
| 2026-08-11 | `tools/amazon_deals_bulk/*`／広告GAS Cursor指示／DEALS要件・P1b HUMAN_RUN | **販促P1b**: ②人DL→Python提出xlsx→③。GASは指示のみ | **Git**: revert。③のxlsxは手削除 |
| 2026-08-11 | `広告運用GAS/コード.js`（clone push済） | **販促P0**: `bootstrapTimeSaleSheetsP0`＋メニュー。マスタ／ログ／分析タブ作成 | **Git**: revert 該当差分。シートは手削除可。要 clone 経由 clasp push |
| 2026-08-10 | `AmazonSpapiPut.js`／`AmazonApprovalExport.js`／`コード.js`／CURRENT_PHASE／D_ENTRY | **本番常時ONセット**: PUT／ALLOW_PROD／LV4／SCサマリ未設定=true。MASTER_QTY等はOFF。要 clasp push＋旧falseキー削除 | **Git**: revert。緊急停止は明示 false |
| 2026-08-10 | `tonmana_palette.py`／compose／`コード.js` B-④／SUB docs | **ベース色 PoC 3択**（beige/warm_white/soft_gray）B-④＋`--base-color`。背景のみ。要 clasp push | **Git**: revert |
| 2026-08-10 | `コード.js` B-④／`review_feedback_templates.py`／review_loop | **目視チェック必須パネル**＋再生成コメントテンプレ（偽物感等）。要 clasp push | **Git**: revert |
| 2026-08-10 | `export_sub_images_for_rakuten_matrix.py`／compose | **目視=品番キー×最大10（pick=ab）**。全セット子FO禁止。`--to-checked-children` は出品CKのみ | **Git**: revert |
| 2026-08-10 | `photo_realism_rules.py`／compose fal | **CAMERA_LOOK 案A**（Canon R5 50mm f/1.8・creamy bokeh・輪郭ソフトネス）。目視10枚/全子FOは後続 | **Git**: revert |
| 2026-08-10 | `sub_image_ai_compose_poc.py` | **auto-export前に run_meta を書く**（順序バグで SystemExit→バッチ停止）。SystemExit を捕捉して次JAN継続 | **Git**: revert |
| 2026-08-10 | `コード.js` B-④ | **`finishB4AfterTruthBind` 公開名化**（末尾`_`削除）。`google.script.run` private 固着修正。要 clasp push | **Git**: revert |
| 2026-08-10 | `コード.js` B-④ | **診断ログ diag=v1**（ms刻み・`finishB4`・Drive resolve via・クライアント時刻）。要 clasp push → ログ貼付で原因切り分け | **Git**: revert |
| 2026-08-10 | `コード.js` B-④ | 紐付け後の指示を**同一ダイアログ内表示**（`google.script.run`→`showModalDialog`禁止でハング修正） | **Git**: revert。要 clasp push |
| 2026-08-10 | `コード.js`／`photo_realism_rules.py`／`rakuten_image_names.py`／export／compose／SUB docs | **B-④ JAN↔正本人間紐付け**／写真実写ルール進化ファイル／`{sku}_{pattern}_subN` | **Git**: revert。要 clasp push。`SET_MAIN_AMAZON_BASE_FOLDER_ID` 任意 |
| 2026-08-09 | `work_paths.py`／`master_sets.py`／`export_sub_images_for_rakuten_matrix.py`／`sub_image_ai_compose_poc.py`／`コード.js`／SUB docs | **1話完結 auto-export**: compose後に楽天フォルダへ直出し。人間目視フォルダは1つ。JAN→出品CK子SKU解決 | **Git**: revert。`--no-auto-export`。要 clasp push |
| 2026-08-09 | `package_truth.py`／`sub_image_ai_compose_poc.py`／`sub_image_lp_themes.py`／`sub_image_review_loop.py`／`コード.js`／SUB docs | **PACKAGE_TRUTH必須・LOCK強化・文字量上限・SEO・人間レビュー再生成**。B-④正本確認。Vision QA／ベース色は対象外 | **Git**: revert。要 clasp push |
| 2026-08-09 | `コード.js`／SUB HUMAN_RUN／B3 docs | **B-④**: サブ採用CK JAN→Python／Cursor指示ダイアログ（非合成） | **Git**: revert。要 clasp push（メニュー反映） |
| 2026-08-09 | `sub_image_b3_curate.py`／`export_sub_images_for_rakuten_matrix.py`／`コード.js`／サブ画像楽天 docs | **サブ画像楽天本線**: B-③参照列＋compose→subN export＋E2E手順。AmazonはU4 REUSE | **Git**: revert。要 clasp push。`RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false` |
| 2026-08-09 | `コード.js`／`master_sets.py`／`rakuten_image_names.py`／`export_sub_images_for_rakuten_matrix.py`／楽天SKU docs | **楽天セット数→SKU紐付け転用**: ファイル名子SKU→楽天メイン1／`_subN`→サブ。マスタでN解決。Vision Nは対象外 | **Property** `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false`。**Git**: revert。要 clasp push |
| 2026-08-09 | `sub_image_lp_themes.py`／compose／instruction_report／HUMAN_RUN | AI合成v3: 心理順スロット5×A/B・page分類・想像FO・OpenAI medium本線・pkgロック。パーツ分類は将来メモ | **Git**: 当該ファイル revert |
| 2026-08-09 | `sub_image_fal_edit_compare_poc.py`／HUMAN_RUN | fal競合改変4本比較（背景/光/湯気/文字色・商品色ロック） | **Git**: 当該ファイル revert |
| 2026-08-09 | `sub_image_ai_compose_poc.py`／HUMAN_RUN | 段階1: hybrid振り分け（T03→fal文字禁止／他→Gemini） | **Git**: 当該ファイル revert |
| 2026-08-09 | `fal_image.py`／`sub_image_fal_poc.py`／compose／HUMAN_RUN | fal.ai（FLUX Kontext/Schnell）格安画像PoC | **Git**: 当該ファイル revert。キーは secrets |
| 2026-08-09 | `sub_image_ai_compose_poc.py`／`sub_image_lp_themes.py`／`sub_image_instruction_report.py`／HUMAN_RUN | AI合成v2: 競合実在テーマ／ベージュ／AI≤50%／日本語指示ボード | **Git**: 当該ファイル revert |
| 2026-08-09 | `sub_image_ai_compose_poc.py`／`openai_image.py`／HUMAN_RUN | サブ画像をGemini+OpenAI最新画像モデルで再合成（パーツ流用≥80%／新規≤20%） | **Git**: 当該ファイル revert |
| 2026-08-09 | `コード.js` B-③／`sub_image_b3_curate.py`／`sheets_rw.py`／docs | B-③右列サブ採用レ点＋採用のみログ＋再実行復元。旧確認シート削除 | **Git**: revert。要 clasp push。ログ／旧シートは手削除可 |
| 2026-08-09 | `sub_image_b3_curate.py`／`sheets_rw.py`／`sub_image_part_poc.py`／HUMAN_RUN | JAN別仕分け＋シート「サブ画像競合候補（人間確認）」。メイン/無関係除外デフォルト | **Git**: revert。シートは手動削除可 |
| 2026-08-09 | `b3_comp_catalog.py`／`sub_image_intent.py`／`sub_image_parts.py`／`sub_image_part_poc.py`／HUMAN_RUN | B-③連動: 文字読取で選定・パーツ提案・安全BG合成（文字改変なし）。商品任意 | **Git**: 当該ファイル revert |
| 2026-08-09 | `sub_image_poc.py`／`sub_image_prompts.py`／`_sub_bg_photos/`／SUB IMAGE HUMAN_RUN | サブ画像10種（実写背景Pillow＋競合掛け合わせ＋シーンAI）。MAIN相当なし | **Git**: 当該ファイル revert。背景JPGは `_sub_bg_photos` 削除可 |
| 2026-08-09 | `sub_image_poc.py`／`sub_image_prompts.py`／SUB IMAGE HUMAN_RUN | サブ画像PoC（競合改変＋自社フル・注釈ボード・テスト出力のみ） | **Git**: 当該ファイル削除／revert |
| 2026-08-09 | `コード.js`（画像キー／画素スコア／Keepa列挙／B-③ push）／B3 docs | B-③同一画像は高画素を残す（AmazonメディアID＋SL/_ex） | **Git**: revert／clasp push |
| 2026-08-09 | `コード.js`（parseCrossMall／pick／B-③）／B3 docs | **案A**: 統合セット数不明でも楽天Yahoo競合画像を取得 | **Git**: revert／clasp push |
| 2026-08-09 | `コード.js`（KeepaログJAN／B-③ Amazon段0〜2）／B3 docs／RESEARCH | B-③: ログ◎→貼付◎→マスタ親子。0件次段。Amazon0でも楽天Yahoo。ログJAN=AI | **Git**: revert／clasp push |
| 2026-08-09 | `landscape_layout.py`（新）／`amazon_paste.py`／`amazon_paste_batch.py`／`layout_rules.json`／`コード.js`／docs | 横長合格固定をC③実装。楽天N=1をC指示に反映 | **Git**: 当該ファイル revert／clasp push でGAS反映 |
| 2026-08-09 | `rakuten_badge.py`／`compose_set_main.py`／RULES／HUMAN_RUN／landscape annot | 楽天N=1生成許可。横長N1中央+Octas。N2+合格記録 | **Git**: 当該ファイル revert |
| 2026-08-09 | `SET_MAIN_LAYOUT_RULES.md`／`_export_landscape_rule_annots.py` | 横長: N3/4階段横進み復元＋Octas。N≥5案A（右辺接触廃止） | **Git**: 当該docs/script revert |
| 2026-08-09 | `SET_MAIN_LAYOUT_RULES.md` §1.2／`_export_landscape_rule_annots.py` | 横長: N3/4枠ピン・N≥5 hero三辺＋半端行中央寄せ。注釈再出力 | **Git**: 当該docs/script revert |
| 2026-08-08 | `SET_MAIN_LAYOUT_RULES.md` §1.2／HANDOVER | **横長セットMAINルール草案**（見本154–158・実装前） | **Git**: 当該docs revert |
| 2026-08-08 | `AmazonSpapiPut.js`／`コード.js`／PUT docs | **PUT max_items 既定 5→10**（Property未設定時。ハード上限50は維持） | **Git**: revert。運用で下げるなら `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` を再設定 |
| 2026-08-08 | `AmazonCategoryPt.js`／`コード.js`／P4b承認・HUMAN_RUN | **P4b-d**: ASIN検証・複数多数決・JAN/SHELF重み投票・Browse無しPT禁止＋Dエラーalert | **Git**: revert。`APPROVAL_AMAZON_P4B_PT_MAX_ASINS` 削除可 |
| 2026-08-08 | `AmazonImageMatrixExport.js`／`コード.js`／C・U2 docs | **U2-ε MAIN自動**: 子SKU名一致→BX自動投入（既存非上書き・親一致/余りは対象外） | **Property** `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED=false`。**Git**: revert |
| 2026-08-08 | `AmazonDriveImageExport.js`／`コード.js`／U4 HUMAN_RUN | **U4途中切れ対策**: URL充足スキップ・1SKU例外継続・Putリトライ・時間スライス＋トリガー再開 | **Property** `AMAZON_U4_FORCE_REUPLOAD`／`AMAZON_U4_SLICE_MS` 削除可。**Git**: revert。残留トリガーは `runAmazonU4ResumeFromTrigger` を削除 |
| 2026-08-08 | `bind_amazon_base_to_parents.py`／`amazon_paste_batch`／SET_MAIN HUMAN_RUN | **01白抜き↔楽天メインVision自動紐付け**（`{親SKU}_単体.png`・複数商品量産） | **Git**: revert。リネームは `BASE_BIND_*.json` の old_name で手戻し |
| 2026-08-08 | `layout_rules`／`portrait_layout`／`amazon_paste`／SET_MAIN RULES・HUMAN_RUN | **N=3細長合格確定**: Q(四隅上辺中点)→M＋緑法線上+35px。C②→07→D | **Git**: revert。ノブ `unit0NudgeUpAlongNormalPx=0` でも旧挙動に近い |
| 2026-08-08 | `portrait_layout`／RULES | **N=3細長**: Q＝四隅上辺(TL–TR)中点→円弧中点M（AABB不可） | **Git**: revert |
| 2026-08-08 | `portrait_layout`／RULES | **N=3細長**: Q＝オレンジAABB上辺中点→円弧中点M | **Git**: revert |
| 2026-08-08 | `portrait_layout`／RULES | **N=3細長**: Q＝オレンジ外枠(ProductQuad)上辺中点→円弧中点M | **Git**: revert |
| 2026-08-08 | `portrait_layout`／RULES | **N=3細長**: unit1上辺中点Q＝円弧中点M（法線∩弧） | **Git**: revert |
| 2026-08-08 | `portrait_layout`／`layout_rules`／RULES | **N=3細長**: unit1上辺左右中点＝円弧中点の法線上 | **Git**: revert |
| 2026-08-08 | `portrait_layout`／`layout_rules`／RULES | **N=3細長**: hero枠内＋unit0 AABB中心中点（見た目バランス） | **Git**: revert |
| 2026-08-08 | `portrait_layout`／`layout_rules`／`octas_prep`／`amazon_paste`／RULES | **N=3細長**: 中角8.8°（頂X中点）＋composeでOctas未傾き時に8°適用 | **Git**: revert（Pythonのみ） |
| 2026-08-08 | `portrait_layout.py`／`layout_rules.json`／SET_MAIN RULES | **N=3細長**: hero AABB下辺＝枠下辺（扇剛体平行移動／`pinHeroFrameBottom`） | **Git**: revert（Pythonのみ・clasp不要） |
| 2026-08-08 | `portrait_layout.py`／`layout_rules.json`／SET_MAIN RULES | **N=3分岐**: 細長→top_arc_legacy＋高さ≤92%（N=4同型）／幅広→upright現行 | **Git**: revert（Pythonのみ・clasp不要） |
| 2026-08-08 | `portrait_layout.py`／`layout_rules.json`／SET_MAIN RULES | **N=4分岐見直し**: 細長(H/W≥1.75・缶含む)→**n4_legacy頂円弧**／幅広→upright_quad現行 | **Git**: revert（Pythonのみ・clasp不要） |
| 2026-08-08 | `amazon_paste.py`／`layout_rules.json`／SET_MAIN RULES | **N≥5 Octas**: unit非被覆＋下枠接触＋**hero隠し≤5%で縮小**／upright_quad素材1枚化 | **Git**: revert（Pythonのみ・clasp不要） |
| 2026-08-08 | `layout_rules.json`／`portrait_layout.py`／`コード.js`／`work_paths.py`／SET_MAIN RULES・HUMAN_RUN・LV4／HANDOVER | **縦長タイプ合格固定**（N=2〜≥5・Cメニュー②登録・unit3 BR枠下） | **Git**: revert。GASは clasp push 前ならローカル戻し |
| 2026-08-05 | `portrait_layout.py`／`layout_rules.json`／SET_MAIN RULES・HUMAN_RUN | **縦長N=3合格固定（頂円弧）＋N=4拡張初版** | **Git**: revert。`portraitTilt.n3/n4` を戻す |
| 2026-08-05 | `portrait_layout.py`／`amazon_paste*.py`／`layout_rules.json`／`コード.js`／SET_MAIN docs | **縦長パターン（斜めファン）** | **Git**: revert。GASは clasp push 前ならローカル戻し |
| 2026-08-05 | `c1_packaged.py`／C1 HUMAN_RUN／HANDOVER | **定価空時のみ販売価格フォールバック**（B3はマスタ定価で再PACKAGED） | **Git**: revert `resolve_list_price`。03の旧xlsm（timestampなし）は販売価格誤反映のため使わない |
| 2026-08-02 | `layout_rules`／`rakuten_layer` 3段組版／`master_sets` バリエーション単位／RULES・HUMAN | **楽天Python本線**（金丸数値組版＋単位ヘッダ連携） | **Git**: revert |
| 2026-08-02 | `tools/set_main_image/ai_*`／`gemini_image`／`model_policy`／AI承認・HUMAN_RUN／PHASE／LEDGER | **セットMAIN AI PoC**（Nano Banana＝通常Flash最新・PRO禁止・見本03/04） | **Git**: revert。secrets/はコミットしない |
| 2026-08-02 | `LV4_SET_MAIN_IMAGE_PHASE_AI_APPROVAL`／PHASE／LEDGER | **セットMAINを見本参照AI生成へ振替提案**（Pillow本線凍結） | **Git**: 当該docs revert |
| 2026-08-02 | `tools/set_main_image/*`／`D_MENU_SET_MAIN_IMAGE_*`／承認／PHASE／HANDOVER／LEDGER | **セットMAIN Phase A 実装**（Amazonコラージュ＋楽天数字レイヤ・07出力） | **Git**: revert。07の生成jpgは人手削除 |
| 2026-08-02 | `LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL`／PHASE／LEDGER | **セットMAIN Phase A 方針ロック**（数字レイヤ・07/C・一括・品質でリサーチ切替可） | **Git**: 当該docs revert |
| 2026-08-02 | `LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL`／PHASE／LEDGER | **セットMAIN Phase A 承認ドラフト**（Pillow合成・生成AI不使用・B/C後回し） | **Git**: 当該docs revert |
| 2026-08-02 | 台帳／PHASE／ROADMAP／REMAKE／HANDOVER／LEDGER | **七味出品中**（48h待ち削除）＋D内レ点 clasp push済反映 | **Git**: 当該docs revert |
| 2026-08-02 | `コード.js`／`AmazonApprovalExport.js`／`LV4_D_REMAKE_*`／Facade／PHASE／HANDOVER／LEDGER | **D内レ点実装**（失敗後再GENERATED・案A・確認後に最新openのみ UPLOAD_FAILED） | **Git**: revert。Z 21-④は従来どおり |
| 2026-08-02 | `LV4_D_REMAKE_MENU_APPROVAL`／MAP HUMAN_RUN／Facade§6／SHELF・PHASE／HANDOVER／LEDGER | **D内レ点方針ロック（案A）**＋MAP sync HUMAN_RUN §0b 固め。コードなし | **Git**: 当該docs revert |
| 2026-08-02 | `c1_packaged.py`／`コード.js`（D手渡し人間手順）／MAP・LANE_B台帳／SHELF HUMAN_RUN／LEDGER | **browse=Node IDのみ**（90194/90225対策）＋人間手順を成功/失敗で分離（失敗=05・08禁止） | **Git**: revert |
| 2026-08-02 | `c1_packaged.py`／`コード.js`／MAP・LANE台帳／LEDGER | **browse=BrowsePath（プルダウン表記）**＋人間手順成功/失敗分離。数値ID単独禁止 | **Git**: revert |
| 2026-08-02 | `LV4_MAP_SHEET_JSON_SYNC_*`／`sync_map_sheet_to_column_json.py`／`push_map_attr_patches_to_sheet.py`／`append_map_sheet_error.py`／`MAP_SC_ERROR_LEDGER.md`／PHASE／HANDOVER／LEDGER | **正本=sheet／派生=MD／実行=JSON（sheet→生成）** | **Git**: revert。JSON は `.bak_map_sync` から戻す |
| 2026-08-02 | `c1_packaged.py`／`food_fish_grocery_column_map.json`／MAP生成／LEDGER | **缶飯数量ポリシー**（入数/ユニット=セット缶数、重量=サイズg、AT=size、色・形態・温度未出力） | **Git**: revert。03の該当PACKAGEDを旧版に戻す |
| 2026-08-02 | `food_fish_grocery_column_map.json`／`create_amazon_mapping_sheet.py`／C1 FOOD列マップ／SHELF承認／LEDGER | **GROCERYテーマ=`サイズ`**（純正プルダウンのみ。`SET_NAME`廃止） | **Git**: revert。旧xlsmは03のタイムスタンプ付きを使わない |
| 2026-08-02 | `LV4_SHELF_BROWSE_CATALOG_*`／`c1_shelf_browse_extract.py`／`shelf_browse_catalog.json`／sync MAP／`AmazonCategoryPt.js`／`コード.js`／T2・P4b・PHASE／HANDOVER／LEDGER | **SHELF Browse網羅＋Nodeルーティング**（缶飯GROCERY・MEATエイリアスはフォールバック） | **Git**: revert。Drive catalog／registry 旧版 |
| 2026-08-02 | `shelf_registry.json`／`food_fish_grocery_column_map.json`／指紋／`AmazonCategoryPt.js`／`コード.js`／PHASE／HANDOVER／LEDGER | **GROCERY本線**（FOOD_FISH_GROCERY・Catalog MEAT→GROCERYエイリアス・Dゲート書換） | **Git**: revert。Drive `shelf_registry.json` を旧版に戻す |
| 2026-08-02 | `c1_packaged.py`／`food_seasoning_column_map.json`／`create_amazon_mapping_sheet.py`／C1 FOOD列マップ／P4b承認／PHASE／HANDOVER／LEDGER | **C1 FOOD: PT/browse必須・既定禁止・FOOD許可・ハイライトB（楽天→Yahoo→箇条書き①）・HJ型番** | **Git**: revert 当該差分 |
| 2026-08-02 | `コード.js`／`AmazonCategoryPt.js`／D新規ゲート承認・HUMAN_RUN／Facade／P4b／PHASE／HANDOVER／LEDGER | **D新規ゲート＋Cursor手渡し実装**（Drive棚 File ID） | **Git**: revert。Property `AMAZON_SHELF_REGISTRY_FILE_ID` 削除可 |
| 2026-08-02 | `AmazonSpapiPut.js`／`コード.js`／D_ENTRY／デュアル承認／レ点本線承認／HANDOVER／LEDGER | **相乗り通常運用=prod直**（SKU空でも生成・成功後保存。dry_run任意） | **Git**: revert 当該差分 |
| 2026-08-02 | `コード.js`／`AmazonDriveImageExport.js`／C画像コース HUMAN_RUN／U2要件／HANDOVER／LEDGER | **C表示名をAma新カタログ①／②へ明確化＋U4対象SKU限定の兄弟PT URL補完** | **Git**: revert 当該差分。補完済みURLは必要な行だけ人手で空欄へ戻す |
| 2026-08-01 | `AmazonImageMatrixExport.js`／C画像コース HUMAN_RUN／U2要件／PHASE／HANDOVER／LEDGER | **Amazon候補を1子SKU＝1枚に**（全行タイル廃止・名前一致優先） | **Property**: `AMAZON_IMAGE_CANDIDATE_TILE_ALL_ENABLED=true` で旧表示。**Git**: revert |
| 2026-08-01 | `コード.js`（`runBatchExportAmazonFacade`）／D新規ゲート承認§2.1／D_ENTRY／HUMAN_RUN／PHASE／HANDOVER／LEDGER | **相乗りASIN空 soft skip**（相乗りのみでも行スキップ・全体継続） | **Git**: revert 当該差分 |
| 2026-08-01 | `LV4_D_NEW_PT_SHELF_GATE_…`／`D_MENU_D_NEW_PT_SHELF_GATE_…`／Facade§6／P4b-c／PHASE／HANDOVER／LEDGER | **D新規ゲート＋Cursor手渡し方針ロック**（docsのみ・実装未） | **Git**: 当該docs revert |
| 2026-08-01 | B-T2承認／HUMAN_RUN／T0・T1誘導／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T2§4ロック**（複合複数明示・三点スキップ・実装は実需後） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_B_BULK_TEMPLATE_T2_…`／`D_MENU_LANE_B_BULK_T2_…`／T0・T1誘導／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T2方針ドラフト**（コードなし・複合優先・実需起動） | **Git**: 当該docs revert |
| 2026-08-01 | `コード.js`／`AmazonApprovalExport.js`／P0承認／D_ENTRY／PHASE／HANDOVER／LEDGER | **D送信在庫: 新規もMASTER時qty内訳確認**（相乗りと同水準・UI明示） | **Git**: revert 当該差分 |
| 2026-08-01 | 台帳／C1 §0b／D_ENTRY §1f／PHASE／HANDOVER／LEDGER | **七味再送サマリ確定**（100521のみ・99016解消） | **Git**: 当該docs revert |
| 2026-08-01 | `コード.js`（`createZSplitMenu`） | **Zメニュー連番サブ化**（1〜10／11〜17／18〜21・21内分割。関数名・番号不変） | **Git**: revert |
| 2026-08-01 | `c1_bulk_fill_by_name.py`／T1_PROD承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T1 prod実装**（`--mode prod`・三点スキップ・スモークOK） | **Git**: revert 当該差分 |
| 2026-08-01 | `LV4_LANE_B_BULK_TEMPLATE_T1_PROD_…`／`D_MENU_LANE_B_BULK_T1_PROD_…`／親T1／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T1 prod第2段方針ドラフト**（コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `c1_bulk_shelf_lookup.py`／`c1_bulk_fill_by_name.py`／`c1_bulk_name_map.py`／`shelf_registry.json`／`food_seasoning_column_map.json`／B-T1 docs | **B-T1実装**（棚引き＋項目名 dry_run・三点スキップ・スモークOK） | **Git**: revert 当該ツール＋docs。prod未 |
| 2026-08-01 | B-T1承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER | **B-T1に棚引き・DL要否案内を追記** | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_B_BULK_TEMPLATE_T1_…`／`D_MENU_LANE_B_BULK_T1_…`／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T1方針ドラフト**（項目名マップ・複合SEASONING・コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `c1_bulk_fingerprint.py`／B-T0承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T0実装**（09指紋→05・三点スキップ・スモークmatch） | **Git**: revert ツール＋docs |
| 2026-08-01 | B-T0承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T0入口を09に確定** | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_B_BULK_TEMPLATE_T0_…`／`D_MENU_LANE_B_BULK_T0_…`／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T0承認ドラフト**（指紋差分・`04`集約・コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | P4b／C1 HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b-b合格**（HERB→C1 PT=SEASONING）＋ローカルPython=Agent実行明記 | **Git**: 当該docs revert |
| 2026-08-01 | P4b承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b-a合格・P4b-b=HERBフォールバック手順** | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonCategoryPt.js`／P4b承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4bマスタ競合版実装**（多数決廃止・WARNのみ） | **Git**: revert。Property OFF |
| 2026-08-01 | P4b承認§2／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b§2マスタ競合改定**（多数決廃止・WARNのみ・コード未） | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonCategoryPt.js`／P4b承認§2／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b多数決改定**（◎上位5・同票は売上1位・SEASONING優先廃止） | **Git**: revert。Property OFF |
| 2026-08-01 | `AmazonCategoryPt.js`／`AmazonSpapiPut.js`／`コード.js`／`c1_packaged.py`／column_map／P4b承認／HUMAN_RUN／PHASE等 | **P4b実装**（21-⑱・空セル書込・C1マスタ優先・三点スキップ） | **Git**: revert。Property OFF |
| 2026-08-01 | P4b承認§2／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b方針ロック**（Keepa既存browse・提案HERB可・C1本線SEASONING/HPC） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_AMAZON_CATEGORY_PT_P4B_…`／`D_MENU_P4B_…`／PHASE／HANDOVER／LEDGER／ROADMAP／P4a | **P4b承認包起草**（方針待ち・コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | Phase2承認§5／HUMAN_RUN§3／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **デュアル Phase2検収OK**（dry `…f372b8`／`…40d85e`・prod `…8fcbc2`／`…cc7c72`） | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonSpapiPut.js` | **`amazonSpapiPutOfferSellerSkuHeader_` 復元**（Phase2編集で消失→`ReferenceError`。系統→保存列名） | **Git**: revert |
| 2026-08-01 | `コード.js`／`AmazonSpapiPut.js`／`LV4_DUAL_OFFER_PHASE2_…`／`D_MENU_DUAL_OFFER_PHASE2_…`／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **デュアル Phase2実装**（チェックUI・順次PUT・部分成功） | **Git**: revert。UIはラジオに戻る |
| 2026-08-01 | `LV4_LANE_B_…`／`LANE_B_SC_ERROR_LEDGER`／`D_MENU_LANE_B_LEDGER_…`／PHASE／HANDOVER／LEDGER／ROADMAP／P2・C1・D_ENTRY | **レーンB台帳初版＋シード** | **Git**: 当該docs revert |
| 2026-08-01 | A3承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **A3 dry／prod検収OK**（`…49a49e`／`…f677a3`） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_A3_…`／`D_MENU_LANE_A3_HUMAN_RUN`／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **A3承認包・HUMAN_RUN起草**（実機待ち・原則コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | A2承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **A2検収OK記録** | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_A2_…`／`D_MENU_LANE_A2_HUMAN_RUN`／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **A2承認包・HUMAN_RUN起草**（原則コードなし・実機待ち） | **Git**: 当該docs revert |
| 2026-08-01 | デュアル承認§6／D_ENTRY§1b2／A1 HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **デュアル prod検収OK**（`…4ed30e`／`…eb2511`） | **Git**: 当該docs revert |
| 2026-08-01 | デュアル承認§6／D_ENTRY§1b2／A1 HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **デュアル dry_run実機OK記録**（`…8fa79e`／`…d6ed67`。prod未） | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonSpapiPut.js`／`コード.js`／`AmazonApprovalExport.js`／デュアル承認／D_ENTRY・A1 HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **デュアル Phase1実装**（系統別SKU列・三点スキップ） | **Git**: revert。マスタの `_FBA` 列は人手削除可 |
| 2026-08-01 | A1 HUMAN_RUN／属性・A1承認／PHASE／ROADMAP／HANDOVER／D_ENTRY／LEDGER | **A1検収OK記録**（FBA dry_run／prod runId） | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonSpapiPut.js`／`コード.js`／属性承認／A1 HUMAN_RUN／PHASE／HANDOVER／LEDGER | **FBA compliance属性実装**（電池・危険物既定＋失敗時おすすめUI） | **Git**: revert。Property `…FBA_COMPLIANCE_ATTRS=false` |
| 2026-08-01 | `LV4_A1_FBA_COMPLIANCE_ATTRS_…`／PHASE／ROADMAP／HANDOVER／LEDGER | **FBA属性方針承認**（§3折衷・三点スキップ。実装待ち） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL`／A1 HUMAN_RUN／PHASE／ROADMAP／HANDOVER／LEDGER | **A1対比記録＋FBA属性承認ドラフト**（コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_DUAL_OFFER_MFN_FBA_APPROVAL`／A1 HUMAN_RUN／PHASE／ROADMAP／HANDOVER／LEDGER | **デュアル Phase1方針ロック**＋A1暫定NF空七味。コードなし | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_A1_…`／HUMAN_RUN／PHASE／ROADMAP／HANDOVER／LEDGER | **A1方針承認**（prodまで・三点スキップ。実機待ち） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_A1_FBA_…APPROVAL`／`D_MENU_LANE_A1_FBA_HUMAN_RUN`／PHASE／ROADMAP／HANDOVER／LEDGER | **レーンA1承認包ドラフト**（相乗りFBA検収。コード原則なし） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_P2_…APPROVAL` §7.4／PHASE／ROADMAP／HANDOVER／P2 HUMAN_RUN／LEDGER | **方針ロック**: 既存API本線(A)／新規xlsm取り貯め(B)／新規JSONゲート後(C)。コードなし | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_P2_DC_123_…APPROVAL`／P2 HUMAN_RUN／PHASE／ROADMAP／HANDOVER／LEDGER | **P2調査承認＋結論**（①②API不可／③温存。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_P2_DC_123_INVESTIGATION_APPROVAL`／`D_MENU_P2_DC_HUMAN_RUN`／PHASE／ROADMAP／HANDOVER／LEDGER | **P2調査承認包起草**（Dc①②③。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | PHASE／ROADMAP／HANDOVER／LEDGER／P1承認・HUMAN_RUN | **P1検収OK＋吉野家承認済**記録（コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `コード.js`／`AmazonImageMatrixExport.js`／`AmazonDriveImageExport.js`／P1・C・U2 docs／PHASE／LEDGER | **P1実装**: 02手File禁止の案内・Dゲート／U4／④メッセージ | **Git**: revert。要 clasp push |
| 2026-08-01 | `LV4_P1_FILE_MIN_APPROVAL`／PHASE／ROADMAP／HANDOVER／LEDGER | **P1方針承認**（§2既定・三点スキップ。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_P1_FILE_MIN_APPROVAL`／`D_MENU_P1_HUMAN_RUN`／PHASE／ROADMAP／HANDOVER／LEDGER | **P1承認包起草**（07以降 File最少。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `tools/spapi_smoke/spapi_smoke.py`／config.example／P4a承認・HUMAN_RUN／PHASE／ROADMAP／HANDOVER | **P4a読取PoC**: `--poc-category`（Definitions search/get＋Catalog分類。マスタ非書込） | **Git**: revert。フラグ無しなら従来スモークのみ |
| 2026-08-01 | `コード.js`／C承認／C HUMAN_RUN／PHASE／HANDOVER／LEDGER | **CをD型選択モーダル1本へ**（C-0〜C-2はZ互換） | **Git**: revert。トップを `createCImageCourseMenu` サブメニューへ戻す |
| 2026-08-01 | `コード.js`／C承認／C HUMAN_RUN／PHASE／HANDOVER／LEDGER | **C本線 Property一時ON→復元**（常設は07フォルダID。C-2確認＋ScriptLock） | **Git**: revert。復元失敗時は U2/U4 を手動false/削除 |
| 2026-08-01 | `コード.js`／`LV4_C_COURSE_…APPROVAL`／`D_MENU_C_IMAGE_COURSE_HUMAN_RUN`／PHASE／HANDOVER／LEDGER | **Cコース実装**（C-0/1/2・①〜④をZ・E寄せ） | **Git**: revert。メニューは旧C/C-Amazon行に戻す |
| 2026-08-01 | `LV4_C_COURSE_CONSOLIDATION_APPROVAL`／PHASE／HANDOVER／ROADMAP／ファサード§7 | **Cコース統合 方針ロック**（IA・三点スキップ・P1非同梱。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `コード.js`／`AmazonSpapiPut.js`／`AmazonApprovalExport.js`／P0承認／D_ENTRY HUMAN_RUN／PHASE／LEDGER | **P0実装**: D在庫0\|マスタ・prod既定・ALLOW_MASTER_QTY・E→D補助/Z互換 | **Property** MASTER_QTY/ALLOW_PROD false。**Git**: revert |
| 2026-08-01 | `LV4_D_P0_THREE_REVIEW_MAJORITY`／P0承認／CHECKBOX／マトリクス／Lv4§3.4／PHASE | **P0三点(B)反映**（マスタqty＝承認②。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `AMAZON_DEV_ROADMAP_P0_P4`／`LV4_D_P0_…APPROVAL`／`LV4_AMAZON_CATEGORY_PT_POC_APPROVAL`／PHASE／ファサード／HANDOVER | **Amazon P0〜P4／Dcロードマップ確定**（P4a並列調査・P0三点対象。コードなし） | **Git**: 当該docs revert |
| 2026-07-31 | `AmazonApprovalExport.js`／D_ENTRY HUMAN_RUN §1e | **SCサマリ監視の自動ON/OFF**（GENERATED成功→待ちリスト＋トリガー設置。終端ステータスで削除。上限72h。Property `…_SC_SUMMARY_WAIT`） | **21-⑰**でトリガー＋待ちクリア。**Property** ENABLED=false。**Git**: 対象差分revert |
| 2026-07-31 | `Yahoo.js`／`コード.js`／`YAHOO_OAUTH_REAUTH_HUMAN_RUN.md` | **Yahoo再認証半自動**（認可code貼付→B14更新＋**C14取得日**。残り≤7日でB17へメール。任意日次トリガー。出品本体非改変） | **メニュー**でトリガー削除。**Property** `YAHOO_OAUTH_REDIRECT_URI` 削除可。**Git**: 対象差分revert＋メニュー行削除 |
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
