# 変更台帳（復元用メモ）

コード全文ではなく **対象・目的・戻し方** のみ記録する（[AGENT_HANDOVER.md](AGENT_HANDOVER.md) §8・§9.1）。

| 2026-09-08 | `PurchaseResearch.js`（`prCycleRunEnumBatch_`／`prCycleCatalogPendingList_`／`prCycleEnumBatchN_`／`prCycleEnumBatchHasBudget_`／`prDispatchHeartbeat_` の `enum_unknown` 分岐）／CYCLE §6.3・§7.1 | **規模帯の確定（Catalog列挙）を1実行でまとめる**: 心拍1回=1社→最大8社（`PR_CYCLE_ENUM_BATCH_N`）。列挙そのものは実測16〜48秒なのに、1社ごとに 洗い→店→心拍 の受け渡しをやり直すため**1社あたり7分**かかっていた（選定1.6分＋yield2分＋店busy2分＋往復1分）。未判定244社の消化が約28時間→**約7時間**。打ち切りは 小中が1社出た（`stop=sm_found`）／予算残5分（`stop=budget`）／対象切れ。1社の例外は飛ばして継続（title0 は保留処理のみ）。`prCycleFirstCatalogPending_` は一覧版の cap=1 に寄せて条件の二重管理を排除。模擬6ケースで打ち切り確認。要 clasp（①） | Prop `PR_CYCLE_ENUM_BATCH_N=1`（心拍1回=1社の旧動作）。Git revert。要 clasp |
| 2026-09-08 | `PurchaseResearch.js`（`prMakerBandFromSum_`／`prMakerBandRowMap_`／`prMakerBandMapClear_`／`prWriteMakerSumBand_`）／FULL §7bis 10 | **規模帯引きの全社ループ内シート読みを撤去**: `prMakerBandFromSum_` を「メーカー→{帯,進捗,見積}」表1回読みに。旧は1社ごとに `prFindMakerSumRow_`（サマリ458行全読み）＋ヘッダ＋当該行の3回読みで、選定はキューに出る社ごとにこれを呼んでいた（模擬120社で **480回→2回**・判定結果は一致）。**目的=洗える社が無いときの選定 2〜11.9分→数秒**（1社見つかる選定は19秒。空振り24回・夜間10.3hの61%が Keepa作業ゼロだった真因）。規模帯書き換え時は `prMakerBandMapClear_()`。要 clasp（①） | Prop `PR_MAKER_BAND_MAP=false`（1社ずつ読み直す旧動作）。Git revert。要 clasp |
| 2026-09-08 | `PurchaseResearch.js`（`prCycleAliveReason_`／`prCycleInflightStaleLimitMs_`／`prCycleInflightStaleFastEnabled_`／`prCycleEnsureLiveness_` のログ）／HUMAN_RUN 強制終了耐性 | **番犬の実行中印を35分基準に**: `inflight` の腐り判定だけトークン回復基準（66分）から30分壁基準（35分＝洗い本体の `PR_WASH_INFLIGHT_STALE_MS_` と同値）へ。GAS は約30分で殺すので35分超の実行中印は生存しえない。**沈黙死の検知が70分→40分**（15分周期のため実測ケースで 01:12→00:42）。予約が残ると `reason=wash_trigger` で生存扱いになるため、この判定が唯一の検知口だった（実測 09/08 00:02 発火→`pick_begin` 沈黙→179分 Keepa作業ゼロ・唯一の FAIL `gap_hole`）。`expect_overdue` は不変。`liveness_reset` に `deadPhase=`／`overdueMin=` を追加。要 clasp（①） | Prop `PR_CYCLE_INFLIGHT_STALE_FAST=false`（旧66分）。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prEffJudgeAfterArm_`／`prEffCkptBudgetRuleEnabled_`／`prEffCkptNormalAfterMin_`）／HUMAN_RUN §10.2・FULL §10 | **ckpt の効率合否を予算基準に**: `*_ckpt` は「残り予算+2分」で貼るのが正（`prCheckpointWashFullTrigger_`。同一実行と重ならない安全ネットなので短縮しない。満タン2分上限も ckpt 除外）。判定側だけ固定15分の古い前提で、正常な afterMin 15〜22分を WARN `ckpt_long_net`／FAIL `ckpt_full_bucket_long_wait` にしていた（**実害の空き時間はゼロ＝誤検知でログが埋まるだけ**。実測7件すべて PASS へ）。予算+2+3（既定25分）以下は `ckpt_safety_net` PASS、超過**かつ満タン**だけ `ckpt_long_despite_tokens` WARN。ckpt 以外の予約・overlap の判定は不変。要 clasp（①） | Prop `PR_EFF_CKPT_BUDGET_RULE=false`（旧 固定15分）。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prListExistingAsinRows_`／`prListDiffUseFormulaEnabled_`）／FLOW §5.2 | **品番リストの毎回書き直しを停止**: 差分判定で数式セルを数式として比べる。`getValues` は `=HYPERLINK(...)` を「メイン」で返すため、画像URLがある行が**常に差分あり**になり、洗いスライスごとに削除→行5再挿入していた（実測 `refresh=4 skip_same=1` を80秒おき。**人が手入力した列が毎回消えていた**）。要 clasp（①） | Prop `PR_LIST_DIFF_USE_FORMULA=false`。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prKeepaWashGateProducts_`／`prKeepaProductStats90Fetch_`／`prKeepaHasOffersInfo_`／`prFullNeedsD1Enrich_`／`prEnrichKeepaForPassAsins_`／`prWashFullKeepaSlice_`）／FULL §5.1・5.2 | **Keepa GET を2段ゲートに**: 門は `stats` だけで判定できる（出品者数は `stats.current[11]`。offers は門に使っていない）ので、1段目は offers なし＝1トークン/件。通過（実測3%）だけ offers=20 で取り直す。**目的=1件あたり約5→約1.2トークン**。落ちも薄い行で Keepaフルに残す（GETなし再門の資産）。通過で2段目が429なら Keepa済にせず次スライスへ。D1補強に「通過なのに BuyBox 2列とも空」を追加（実測 通過782件中5件・20件/回上限）。要 clasp（①） | Prop `PR_KEEPA_TWO_STAGE_GATE=false`（従来の一発 offers=20）／`PR_D1_ENRICH_BLANK_BB=false`。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prCleanupKeepaFullBlankRows_`／`prKeepaFullAsinRowIndex_`）／FULL §7bis 8・9 | **Keepaフルの1行ずつ読みを撤去**: 空行掃除は ASIN 列1回読み→候補行だけ全列。索引づくりも ASIN 列だけ。**目的=Keepa取得のあるスライス22〜29分→数分**（実データ8,386行に対し毎回8,385回の個別読み取り。1ASINでも22.9分）。要 clasp（①） | Prop `PR_KEEPA_FULL_BLANK_SCAN_BATCH=false`。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prWashQueueSnapshot_`／`prWashQueueAppendAsins_`／`prWashQueueDedupeAsins_`／`prWashQueuePendingKeepa_`／`prWashQueueMarkKeepaDoneByAsins_`／`prScheduleWashFullTrigger_`）／FULL §7bis | **洗い本体の全走査撤去＋満タン2分上限の分離**: 1社の処理中はキューを1回だけ読む（列挙の追記は継ぎ足し）。Keepa済は連続行まとめ書き。満タン時の待ち2分上限を `PR_WASH_ARM_FULL_CAP` として効率ループから独立（ckpt は対象外）。**目的=1スライス22〜29分→数分・空白待ちの解消**。要 clasp（①） | Prop `PR_WASH_QUEUE_SNAPSHOT=false`／`PR_WASH_ARM_FULL_CAP=false`。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prCyclePoolCounts_`／`prWashQueueUniqCount_`／`prMakerCatalogEstimate_`／`prFinalizeEnumOnlySize_`／`prRunMakerWashFullBody_`）／FULL §7bis | **選定の全走査撤去**: キュー件数は「メーカー→件数」表を1回作って共有。選定プールは1実行1回。**目的=選定24分→数分**（pick 12.8分＋pick後11.4分で予算20分を食い潰し Keepa 到達0だった）。要 clasp（①） | Prop `PR_WASH_QUEUE_COUNT_MAP=false`／`PR_CYCLE_POOL_REUSE=false`。Git revert。要 clasp |
| 2026-09-06 | `competitor_jitta/JittaResolve.js`／`PurchaseResearch.js`／`maker_alias.py`／B_COMPETITOR_JITTA | **法人格ゆれ Phase A**: 同一キーへ実態ミラー・マスタ heal・二重AI防止。行削除なし。要 clasp（競合＋①） | Git revert。要 clasp |
| 2026-09-07 | `appsscript.json` | **`script.scriptapp` 復元**（D一括の `newTrigger` 権限。作業ツリーから欠落→clasp push でGAS側も欠落）。要 clasp＋**再承認** | スコープ削除で戻る。要 clasp |
| 2026-09-07 | `コード.js`／CDE要件 | **C／D／E／Bメニュー復元**（truncate後Git復元で旧IAに戻った分）。bak先頭＋B5完結関数。要 clasp・シート再読込 | bak差戻し／Git revert。要 clasp |
| 2026-09-06 | `コード.js`／マトリクス表示 | **Driveプレビューを thumbnail `sz=w1000`**（`rakutenDrivePreviewImageFormula_`）。実ファイル解像度・Dアップは不変。要 clasp | ヘルパーを `uc?export=view` に戻す／Git revert。要 clasp |
| 2026-09-06 | `コード.js`／SUB_PACK_ASC要件 | **サブ昇順詰め**: N昇順→同Nはファイル名昇順で列1〜10。番号なしは余り枠／候補。`RAKUTEN_MATRIX_SUB_PACK_ASC_ENABLED`。要 clasp | Prop false。Git revert。要 clasp |
| 2026-09-05 | `コード.js`／WORK_MERGE_SCOPE要件 | **C①02+一時合算**＋**②レ点外JANスキップ**（先頭行青枠投入廃止）。Prop `RAKUTEN_MATRIX_MERGE_WORK_FOLDER`／`RAKUTEN_MATRIX_SCOPE_FILTER_ENABLED`。要 clasp | Prop false。Git revert。要 clasp |
| 2026-09-05 | `コード.js`／SKIP_EXISTING要件 | **C取得上限100**＋**マスタ既存URL枠はアップ省略**（`RAKUTEN_UPLOAD_SKIP_EXISTING_MASTER_URL`）。要 clasp | Prop false。PROCESS_LIMIT戻し。Git revert。要 clasp |
| 2026-09-05 | `コード.js`／MATRIX_LAYOUT要件／LEDGER | **楽天マトリクス L1〜L3**: メイン2〜10他セット昇順・サブ全子共有・親=最少セット。JANサブ→親。要 clasp。戻し `RAKUTEN_MATRIX_LAYOUT_EXPAND_ENABLED=false` | Prop false。Git revert。要 clasp |
| 2026-09-05 | `コード.js`／CDE要件／CURRENT／ALIGN／HANDOVER | **C／D／E分割**: C=楽天マッチング。D=楽天Yahooのみ。E=Amazon画像＋出品。D補助廃止→Z。要 clasp | Git revert。要 clasp |
| 2026-09-05 | `コード.js`／B6要件／ALIGN §1.2／SET_MAIN HUMAN／CURRENT／HANDOVER | **B-⑥ Amazon MAIN**: メニュー本線。CからAmazon量産削除・補助化。要 clasp | Git revert。要 clasp |
| 2026-09-04 | `コード.js`／ALIGN §1.1／B3確認／CURRENT／HANDOVER | **B→B-③ 案1**: メニュー増なし。起動確認＋`B_INTEGRATED_INCLUDE_B3_IMAGES`。全パック後に約1分で非対話B-③（Amazon自動）。要 clasp | Prop false。Git revert。要 clasp |
| 2026-09-04 | `コード.js`／`b5_parent_shared.py`／ALIGN | **正本・個品・文字色の完全連携**: B-④で正本初期選択＋個品ラベル。親SKU共有＋Drive `_b5_parent_shared.json`。要 clasp | Git revert。要 clasp |
| 2026-09-04 | ALIGN §1／§3／§1.1／B3確認／SET_MAIN | **文字色＝親共有**（MAIN／サブ順不同）。**B一括＋B-③＝オプション**（未設定OFF・実装後段） | docs のみ。実装は別タスク |
| 2026-09-04 | `c_set_main_rakuten_prompt`／`b5_main_base_review_loop`／ALIGN／SET_MAIN HUMAN_RUN | **弱い統合・親絞り**: 02後C指示＝今の親だけ（`--parent-sku`）。任意実行ボタン。`byChildSku`集計修正。全親は scope=all | Git revert |
| 2026-09-05 | `competitor_jitta/`／`PurchaseResearch.js`／CYCLE／HUMAN_RUN | **競合DB実態AI第一段**: トリガー可・① `PR_JITTA_OWNER=comp` で二重防止。切戻し OWNER=hb／競合トリガー停止。要 clasp（競合フォルダ＋①） | Prop／トリガー。Git revert |
| 2026-09-04 | `PurchaseResearch.js`／`cycle_heartbeat`／CYCLE§4.5bis／HUMAN_RUN | **実態AI分離**: 心拍 `jitta_resolve`（cap8）・A-light末尾デフォルトOFF・未調査≥30で自動レ点抑制。仮想 T-CHB-24〜27（27/27）。切戻し `PR_JITTA_HB_LANE_ENABLED=false`。要 clasp | Git revert。Prop false。要 clasp |
| 2026-09-03 | `PurchaseResearch.js`／watchdog dry／HUMAN_RUN | **店レーン譲り**: 洗い→心拍で inflight clear＋3分安全網＋shop busy。alight 完了で busy 解除＋1分。切戻し `PR_SHOP_LANE_CLEAR_INFLIGHT=false`。仮想 T-CTW-23/24。要 clasp | Git revert。Prop false。要 clasp |
| 2026-09-03 | `PurchaseResearch.js`／watchdog dry／HUMAN_RUN | **P-eff-1/2 効率ループ整流**: 実行中 ckpt 廃止・満タン `afterMin≤2`（`eff_loop_cap`）・WAIT_TOKEN 前 inflight clear・overlap age≥8 即 takeover／短再試行2分。切戻し `PR_WASH_EFF_LOOP_ENABLED=false`。要 clasp | Git revert。Prop false。要 clasp |
| 2026-09-03 | `PurchaseResearch.js`／`eff_judge_report.py`／HUMAN_RUN | **効率合否ログ**: ①ログ `効率合否`（slice/arm/overlap/stale・cause・fixHint）。切戻し `PR_EFF_JUDGE_ENABLED=false`。要 clasp | Git revert。要 clasp |
| 2026-09-02 | `PurchaseResearch.js`／`maker_alias.py`／`cycle_c3_select.py`／`d2_filter.py` | **実態ラベル分離**: 新規行は `未調査`（AI前）。`要確認` は AI後の人判定待ち。`prCollectJittaResolveTargets_` は信頼度空のみ。移行メニュー 99→16。要 clasp | Git revert。要 clasp |
| 2026-09-01 | `PurchaseResearch.js`／`cycle_trigger_watchdog.py`／fixtures T-CTW-13〜18 | **洗い ckpt 自己衝突対策**: ckpt/overlap で `wait_shortened` 無効・overlap 必ず再予約（stale 寄せ）・inflight 中は `trigger_missed` 抑制。切戻し `PR_WASH_CKPT_NO_SHORTEN=false`／`PR_WASH_OVERLAP_STALE_AWARE=false`。仮想18/18。要 clasp＋実機1回 | Git revert。要 clasp |
| 2026-09-01 | `b5_main_base_review_loop`／`b5_virt_abc` measure_grid20 | **100マスガイド**: 文字/装飾マスを薄赤、金丸予約4マス黄。16/24/25 目安 | Git revert |
| 2026-09-01 | `b5_virt_abc`／`b5_main_base_compose`／templates | **20マス厳守**: キャッチサイズ縮小・目標16・retry上限25 | Git revert |
| 2026-09-01 | `sub_image_food_canon_gate`／`sub_image_food_canon_launch`／`コード.js` beginB5AfterTruthBind | **MAIN用見本ゲート2ボタン**（変更なしで閉じる／選び直したらここで保存）。`--context main`。要 clasp | Git revert |
| 2026-09-01 | `b5_main_base_names`／compose `--only-theme`／ALIGN | **salesTalk途切れ修正**（硬切り禁止・だけで後半優先） | Git revert |
| 2026-09-01 | `コード.js`／ALIGN §7.14 | **B-④指示に見本ゲート必須**（finishB5AfterTruthBind）。要 clasp | Git revert。要 clasp（出品） |
| 2026-09-01 | `food_unit_canon`／`b5_main_base_compose`／`b5_virt_abc`／ALIGN §7.14 | **P-align-3**: MAIN が theme_new_run 人間確定 CANON を参照 | Git revert |
| 2026-09-01 | ALIGN §7.14／§1／§5／MAIN要件／CURRENT_PHASE | **個品見本 MAIN↔サブ共有**要件ロック。構図は別。P-align-3 実装未 | docs のみ。実装は別タスク |
| 2026-09-01 | `b5_main_base_names`／`b5_virt_abc`／compose／ALIGN §7.6ter | **MAIN catch大＋sales小**: 階層プロンプト。20マス厳守 | Git revert |
| 2026-08-31 | `b5_virt_abc`／`b5_main_base_compose`／ALIGN §7.6ter | **MAIN Phase1**: scene/catch分離・catchImage短文化・STYLE_TYPO色競合解消 | Git revert |
| 2026-08-31 | `b5_virt_abc`／`b5_main_base_compose`／ALIGN §7.6bis | **MAIN textColor**: `_prompt_p3_free` に accent_text_block 配線 | Git revert |
| 2026-08-31 | `コード.js`／`b5_set_base_map`／compose／dry_sub／ALIGN §7.6bis | **textColor 親ごと**: Drive/マップ自動＋`--text-color-for`。ハードコード1廃止 | Git revert。要 clasp（出品） |
| 2026-08-31 | `PurchaseResearch.js`／CYCLE§6.4quat／HUMAN_RUN | **D2 overflow**: onlyBlank・cap1000・停止メール（復帰手順）・overflow gate・メニュー28・D2後自動C4 | Git revert。要 clasp |
| 2026-08-31 | `mix_review`／`_find_truth`／gate `_truth_for` | **「未選定または失敗」誤表示修正**（人間選定を正しく表示）。正本解決のまぐろ既定を廃止・JANヒント優先 | Git revert |
| 2026-08-31 | `コード.js`／ALIGN §7.13 | **生成不要を正本紐付けへ**: ON時テーマ案Geminiスキップ→指示直行。要 clasp | Git revert。要 clasp（出品） |
| 2026-08-30 | `コード.js`／`b5_skip_generate_finish`／gate／dry_sub／compose／ALIGN §7.13 | **画像の生成不要**: B-⑤ CK→skip指示・既存 export＋レビュー完走。要 clasp | Git revert。要 clasp（出品） |
| 2026-08-30 | `food_unit_canon`／`food_canon_gate`／`dry_sub`／ALIGN §7.12.6 | **IDENTITY／SIZZLE UI**＋副のみエラー＋テーマ mode 参照 | Git revert |
| 2026-08-30 | `コード.js`／ALIGN §7.12.3／RAKUTEN_COURSE HUMAN_RUN | **B-⑤指示先頭＝見本ゲート必須**（確定前に dry_sub/compose 禁止） | Git revert。要 clasp（出品） |
| 2026-08-30 | ALIGN §7.12／`food_canon_manual_ingest`／`food_unit_canon`／gate | **§7.12**: コピペ＋ゼロ案内＋同構図デデュプ（aHash／ランク多様） | Git revert |
| 2026-08-30 | `food_unit_canon`／`sub_image_food_canon_gate*`／ALIGN §7.11 | **見本ゲート**: 候補5〜6・主1＋副0〜2→生成→同画面レビュー | Git revert |
| 2026-08-30 | `food_unit_canon`／`sub_image_mix_*`／`dry_sub_overview_copy_test`／`compose_poc`／ALIGN §7.10–7.11 | **混在段階B**: 独自再生成＋FOOD_UNIT_CANON機械統一 | Git revert |
| 2026-08-31 | `PurchaseResearch.js`／HUMAN_RUN | **A-1**: 店日次カウンタリセット（auto_ck／alight done／rounds／lastNew）。週間大特大枠は残す | Git revert。要 clasp |
| 2026-08-31 | `PurchaseResearch.js` | **自動レ点**: 調査済み（巡回日／alight+shop_full）を除外。誤レ点は A-light で CK 外し | Git revert。要 clasp |
| 2026-08-31 | `PurchaseResearch.js`／`cycle_heartbeat`／CYCLE§4.4 | **根因修正**: 自動レ点を店ラウンド尽きた判定より先に。順=auto_ck→alight。lastNew未設定=-1即終了を修正。T-CHB-23 | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§4.4 | **full_trigger**: 店ラウンド中は `prWashContinueDeferInfo_` で大続きを選定除外（`defer_shop_before_large`）。target=大→shop_first を禁止 | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `PurchaseResearch.js`／`cycle_heartbeat`／CYCLE§4.4 | **実装**: 大を選ばず `prCycleShopBeforeLargeAction_`。大特大完走後は `shop_after_large`（C0→auto_ck→alight）。仮想 T-CHB-14/22 | Git revert。要 clasp＋A-1 |
| 2026-08-30 | CYCLE§4.4・§6.3bis・§7／HUMAN_RUN | **小中=0時は店ラウンド≤10尽くすまで大・特大を選定しない**（pick→defer禁止）。大特大完走後は門通過→セラー→調査 | docs revert |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§7.5／HUMAN_RUN | **A-1後のCatalog遅延対策**: abort世代・孤児ロック破棄・心拍 other_busy は1分再試行。旧Keepaは打ち切り | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§7.1・§7.5／HUMAN_RUN | **A-1=実行中一時停止＋先頭再選定**: `a1_park`・キュー残し・続き選定スキップ・C4 recover は D1未のみ。番犬 hard_reset は park しない | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `sub_image_mix_*`／`sub_image_review_launch`／ALIGN §7.10 | **混在レビュー段階A**: 独自＋競合1画面・採用export。compose最新は日時優先 | Git revert |
| 2026-08-30 | `sub_image_lp_themes`／`sub_image_ai_compose_poc`／`theme_compose_rules`／`コード.js`／ALIGN §7.9 | **競合本線**: チェック∩写真・1テーマ1枚・想像OFF。袋はパーツ外。スケール／T03／T13／needPackage／核語色 | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `review_feedback_templates`／`コード.js`／`dry_sub_overview_copy_test`／ALIGN §7.6bis | **サイズ定型＋強調色**: textColorマップ固定・微差可。テンプレ2本追加 | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `コード.js`／`photo_realism_rules`／`dry_sub_overview_copy_test`／ALIGN §7 | **P-align-2d テーマ別新規本線**: 概要＋訴求＋needPackage。人物詳細・物背景実写。袋は必要時のみ。切戻し `B5_THEME_NEW_PATH=false` | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `PurchaseResearch.js`／①プロファイル実機 | **食品・飲料・お酒**=調査ON+`57239051`。別名「食品」行削除。seed判定をルート名に合わせる | Git revert。要 clasp |
| 2026-08-30 | `PurchaseResearch.js`／schema／CYCLE§4.1bis | **店規模実装**: ①プロファイル調査カテゴリー＋ルートID。カテゴリ別query見積。A-light OR歩き。セラー内訳列。構成比は規模に使わない | Git revert。要 clasp |
| 2026-08-30 | CYCLE§4.1bis／V1／NOW／FLOW／HUMAN_RUN | **店規模分担**: 調査カテゴリー＝①プロファイル。SKU見積＝カテゴリ別query。構成比は優先のみ。501=単独≥501。在庫サンプル中止。内訳はレ点2以上 | docs revert。実装は別タスク |
| 2026-08-30 | PurchaseResearch.js／CYCLE§7.5／HUMAN_RUN | **A-1店レーン優先**: sm=0で大/特大Keepa続きを保留し自動レ点→A-lightへ。full_trigger shop_first。進捗leftは残す | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `コード.js`／`dry_sub_theme_gate`／ALIGN | **P-align-2b 本線**: 材料門＋Gemini 2文。1・2必須＋ほか最大8＝10。切戻し `B5_PROPOSE_GEMINI_ENABLED=false` | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `コード.js`／`sub_image_lp_themes`／`sub_image_ai_compose_poc`／ALIGN | **P-align-2b A4**: B-⑤テーマ概要チェック→`--theme-ids-for`。切戻しは指示からフラグを外す（自動スロット） | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§4／HUMAN_RUN／`dry_cycle_alight_hit` | **店A-light 1店フル**: 小中大はqueryページ＋productスライス続き。特大≥501は浅いhold。日次レ点8・店目標5・上限10・新メーカー0で追加round≤10。心拍 `alight_continue`／`shop_before_large` | Git revert。要 clasp＋止めてから A-1 |
| 2026-08-30 | `sub_image_review_loop`／`sub_image_ai_compose_poc`／`sub_image_review_launch`／ALIGN | **P-align-2a サブUI寄せ**: lightbox・`__rN`・採用再export・目視フォルダ・枚数優先 | Git revert |
| 2026-08-30 | `コード.js`／`compose_set_main`／`b5_set_base_map`／review_loop／ALIGN要件 | **表示B-④＝MAIN／B-⑤＝サブ**。マップ`textColor`→C親ごと。既存3親追記 | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `b5_main_base_review_*.py`／map／`compose_set_main`／`コード.js`／B5要件 | **B-⑤ P3b 実装**: 複数土台＋レ点子マップ＋既存再利用＋拡大＋再生成世代 | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | B5要件／HUMAN_RUN／CURRENT_PHASE | **B-⑤ P3b 要件ロック**（実装前） | docs revert |
| 2026-08-29 | `b5_main_base_review_*.py`／`コード.js`／B5要件 | **B-⑤ P3 ブラウザ本線**: 写真レビュー・レ点再生成・ラジオ→02。GAS採用UI外し | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `コード.js`／B5要件 | **B-⑤ P3（旧）**: スプシ採用ボタン＋Drive makeCopy（UI外し・API残置可） | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `b5_main_base_names.py`／export／compose／`コード.js` | **B-⑤ 比較名=商品名省略＋theme＋キャッチ**。02へコピー CLI。既存ランリネーム | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `b5_main_base_compose.py`／`コード.js`／B5要件 | **B-⑤ P2**: compose `--role rakuten_main_base`。チェック案JSON→OpenAI＋Pillow金丸/Octas。02未書込。実画像は明示指示時のみ | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§6.3bis・§7／HUMAN_RUN | **大先・週間枠を続きにも適用・大特大完走後も店レーン**。`defer_xl_large_remain`／`defer_xl_week_cap`。枠・大残りは進捗維持 | Git revert。要 clasp＋止めてから A-1 |
| 2026-08-29 | `PurchaseResearch.js`／CYCLE§6.3〜§7／HUMAN_RUN | **C4→心拍一本化**・**Catalog見積先行**（Keepa新規ゲート・上限メール）・選定から**保留除外**。旧次社 `PR_CYCLE_C4_ARM_MAKER` は成功連鎖でセットしない | Git revert。要 clasp＋A-1 |
| 2026-08-29 | `コード.js`／B5要件 | **B-⑤ M1**: ダイアログで材料→最大4案（門＋Gemini整文）→複数チェック→指示。切戻し `B5_PROPOSE_GEMINI_ENABLED=false` | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `コード.js`／`b5_main_base_propose.py`／B5要件 | **B-⑤ P1**: メニュー＋正本紐付け＋Cursor指示。Property `B5_RAKUTEN_MAIN_BASE_ENABLED=false` で非表示。切戻し同キー false | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `PurchaseResearch.js`／CYCLE§6.4ter／HUMAN_RUN | **週次転記チェック**: ①候補`転記チェック`・日曜トリガー・漏れ時D1・社内メール。切戻し `PR_CYCLE_LIST_CHECK_ENABLED=false` | Git revert。要 clasp＋A-1（トリガー武装） |
| 2026-08-29 | `PurchaseResearch.js`／FULL／HUMAN_RUN | **洗い再開待ち=次cap分**（`prKeepaWashResumeWaitMin_`）。負 tokLeft=0。特大の常時61分を解消 | Git revert。要 clasp |
| 2026-08-29 | `PurchaseResearch.js`／CYCLE§6.4bis／HUMAN_RUN | **途中D1**: Keepaスライス(`wrote>0`・未完)ごとに門通過→品番リスト。C4 D1_DONEなし。切戻し `PR_CYCLE_MID_D1_ENABLED=false` | Git revert。要 clasp＋A-1 |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | キャッチ6〜18字・途中終了禁止。袋名→トーク名は後でメモ | Git revert |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | **abc16**: キャッチ【】優先・語途中禁止。Octas聖域。上部左寄り。再生成 | Git revert |
| 2026-08-29 | `b5_virt_abc.py`／HUMAN_RUN | **abc15**: 材料提案→最大4案×OpenAI画像。Gemini画像拒否。`--legacy-sets`で旧固定 | Git revert |
| 2026-08-29 | B5要件／HUMAN_RUN | **案F旧文掃除**: 字数事前ゲート／食感マクロ／全案修正を削除。Gemini画像生成しない | docs revert |
| 2026-08-29 | B5要件／HUMAN_RUN | **売れる順=1商品ずつ**: カテゴリ差し替え表なし。点数は試行 | docs revert |
| 2026-08-29 | B5要件／HUMAN_RUN | **案F本線**: 複数チェック→画像複数→1枚採用。根拠薄い案は出さない。T4成果・売れる順。旧案A字数事前修正廃止 | docs revert |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | **v13**: T3×3＋T4×2。③Pillow Octas右下1.3マス。`compare/*_all_P3b.jpg` | ファイル削除／docs revert |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | **画像本線=OpenAI**（仮想v12人間ロック）。`--engines` 既定を openai。Gemini画像は比較用 | 既定を gemini,openai に戻す／docs revert |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | **v12**: ③b/cに主婦向け IMAGE_STYLE_TYPO（99.おしゃれ文字装飾見本）。T3新規30は既存・退避除外 | ファイル削除／docs revert |
| 2026-08-29 | `PurchaseResearch.js`／CYCLE§6.3bis〜§6.6／`cycle_heartbeat` | **案B店挟み**・**大特大週間N=2**・**title0→ブランド追加**・empty_research／偽完了。仮想 T-CHB-20/21。Keepa cap既定30 | Git revert。要 clasp＋A-1 |
| 2026-08-29 | `PurchaseResearch.js` `prCandGatePassCountForMaker_` | **偽完了走査の欠落関数修正**: 存在しない `prCandRowsForListAppend_` → D1同型の `prLoadCandPassByAsin_` | Git revert。要 clasp＋A-1 |
| 2026-08-28 | `PurchaseResearch.js`／CYCLE§7.5／仮想10件 | **フェーズ3**: `PR_WASH_TITLE0_SKIP_SET` 永久skip・選定除外・C4 needs=候補ありのみ・偽完了(`done_no_candidates`/`c4_no_candidates rewash`) | Git revert。要 clasp＋A-1 |
| 2026-08-28 | `PurchaseResearch.js`／CYCLE§7.5／HUMAN_RUN | **title0_skip**（3回→保留＋心拍）・**C4 needs/D1 recover**（完走済み複数の D1 未を A-1 で回収） | Git revert。要 clasp＋A-1 |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v11**: ③b/cスマホ主役サイズ（金丸半分）＋おしゃれ・可読性。Gemini/OpenAI再比較 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v10**: ③に写真実写＋キャッチ必須1本。Gemini/OpenAI再比較 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v9**: Gemini/OpenAI同じ指示比較。右上空け→Pillow金丸。OpenAI=gpt-image-2 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v8**: ③b/c スプシ材料・IMAGE_GOLDをAI配置（後載せなし）・文字20%超は再作成 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v7**: ③b/c 虚偽禁止・金丸固定・白背景不可・文字20%・材料内売り | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v6**: ③a添付構図固定。③b/cは短い売れる指示のみ | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／layout_rules／B5要件 | **v5**: ②見出しは帯ルマで黒/白。③人間ベストの対吹き出し＋三本線。P1流用 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／layout_rules／B5要件 | **v4**: ①シズル+袋左下傾き+Octas。②③空金丸。③止まる仕掛け | ファイル削除／docs revert |
| 2026-08-27 | `b5_virt_ai_headline.py`／`layout_rules.json`／B5要件 | **v6**: 見本番地をAIへ。下段白。下線半透明重ね。UDモトヤマルベリ | Git revert |
| 2026-08-27 | `b5_virt_ai_headline.py`／`work_paths.py`／B5要件 | **v5**: 98筆下線参照。20マスに下線含む。金丸解決が下線PNGを拾わない | Git revert |
| 2026-08-27 | `b5_virt_ai_headline.py`／`layout_rules.json`／B5要件 | **見出しAI v4**: 金丸をAI参照渡し。100マス上2行。字数12-18/4-8。アンダーライン可 | ファイル削除／docs revert |
| 2026-08-27 | `b5_virt_ai_headline.py`／`layout_rules.json`／B5要件 | **見出しAI v2**: 金丸消し+98後貼り。文字≤20%。色=C `textColors`。02未書込 | ファイル削除／docs revert |
| 2026-08-27 | `b5_virt_ai_headline.py` | **見出しAI仮想** Gemini 4枚。02未書込。Pillowなし | ファイル削除 |
| 2026-08-26 | B5要件 | **案Aロック**。テーマ対最大4→人1枚。Pillowアピール。選定UI。§3.4–3.7 | docs revert |
| 2026-08-26 | B5要件 §3.4 | **画像もテーマ連動複数提案**（メモ→案Aでロック） | docs revert |
| 2026-08-26 | B5要件 | **Gemini4案＋テーマ**。材料＝キャッチ／箇条書き／AI情報取得data。人修正 | docs revert |
| 2026-08-26 | B5要件 | **ハイブリッド承認**。文言材料＝正本＋商品名ベース。AI発明なし | docs revert |
| 2026-08-26 | B5要件 §3.3 | **人フロー**: 文言事前＋画像は review（ハイブリッド） | docs revert |
| 2026-08-26 | `layout_rules.json`／B5要件 | **アピール2段**: name/talk。上段≤18・下段≤10。`appealHeadline1200` | Git revert |
| 2026-08-26 | `layout_rules.json`／B5要件 | **見本金丸はみ出し**: cx=1102, cy=103, d=343。数字番地は非変更 | Git revert |
| 2026-08-26 | `b5_virt_katsuo.py`／B5 HUMAN_RUN | **B-⑤仮想V1–V3**（たべるおだしかつお・空金丸→数字載せ） | ファイル削除 |
| 2026-08-26 | `D_MENU_B5_RAKUTEN_MAIN_BASE_*` | **B-⑤ P0**: 楽天MAIN基礎（右上金丸・数字は後続）。見本 `04.楽天見本` | docs revert |
| 2026-08-26 | `PurchaseResearch.js`／CYCLE§7／HUMAN_RUN | **後回し原因ログ**: sanitize/pool/defer_xl を①ログへ。不明残でも大特大後回し | Git revert。要 clasp |
| 2026-08-26 | `PurchaseResearch.js`／CYCLE§7／HUMAN_RUN | **enumOnly後**: 規模帯記録・Keepa中解除・心拍再arm。小中残存時は大特大キュー続きを後回し | Git revert。要 clasp |
| 2026-08-26 | `PurchaseResearch.js`／CYCLE§6.4／HUMAN_RUN | **C4 pick=0→心拍arm**。メニュー「心拍を今すぐ再開」追加 | Git revert。要 clasp |
| 2026-08-26 | `cycle_heartbeat`／fixtures／`PurchaseResearch.js`／CYCLE§7／HUMAN_RUN | **心拍完成形**: 仮想19/19。トリガー→列挙/A-light/自動CK/C0/大。切戻し `PR_CYCLE_HEARTBEAT_ENABLED=false` | Git revert。要 clasp＋gmail再承認 |
| 2026-08-25 | `tools/common_ss_view/registry.json`／COMMON_SS_VIEW／`common-ss-view.mdc` | **見積発注_三協乾物** 登録（メニュー見積・発注）。戻し: Git revert |
| 2026-08-25 | `PurchaseResearch.js` | **洗い強制終了耐性**: 予算20分・スライスごとckpt再予約・①ログ洗いスライス・lock待ち=残り秒 | Git revert。要 clasp |
| 2026-08-25 | `PurchaseResearch.js` | **洗いトリガー診断ログ**（expectedAt／fire／cleared／trigger_missed／診断メニュー） | Git revert。要 clasp |
| 2026-08-25 | `PurchaseResearch.js` | **取得日時JST**＋**洗いトリガー停滞対策**（未完再武装・lock_busy再予約・rescue） | Git revert。要 clasp |
| 2026-08-26 | `export_sub_images_for_rakuten_matrix.py` | **目視export上書き禁止**（同名は `__composeDir` 別名） | Git revert |
| 2026-08-26 | `sub_image_review_launch.py`／master_sets／B-④案内 | **レ点親→最新compose→ポート連番で複数タブ** | ファイル削除／Git revert |
| 2026-08-26 | B-④／review_loop／`コード.js` | 完了文言: 次の親reviewへ・手export不要。GAS案内も同文言 | Git revert。要 clasp（GAS） |
| 2026-08-26 | B-④／review_loop | 文言: JAN＋商品名・全選択（コピー）・修正なしで終了等 | Git revert。要 clasp（GAS側） |
| 2026-08-25 | `コード.js`／B3 docs | **B-③プレビュー正方形**: 行高＝列幅100px。autoResizeはプレビュー除外 | Git revert。要 clasp |
| 2026-08-25 | `コード.js`／B3 docs | **B-③**: 親のみ・1親ずつ・途中保存＋トリガー再開（`B3_IMAGE_TIME_SLICE_ENABLED`） | Prop=false／Git revert。要 clasp |
| 2026-08-25 | `PurchaseResearch.js` | **S**: 要確認文言（異種セット等）＋赤白／黄セル | Git revert。要 clasp |
| 2026-08-25 | `PurchaseResearch.js` | **S 単品名**: 合格プロンプト（最少名・異種mix・同一JAN寄せ） | Git revert。要 clasp |
| 2026-08-24 | `PurchaseResearch.js` | **S 単品商品名AI充填**（見積候補・既存非上書き・簡易存在）。`PR_SINGLE_NAME_AI_ENABLED` | Git revert。要 clasp |
| 2026-08-24 | `dry_single_name_from_list.py` | **実リスト20件＋AI** 単品名仮想（見積候補・非書・存在は簡易） | ファイル削除 |
| 2026-08-24 | `single_name_classify.py`／fixtures | **単品名パターン仮想 20/20 PASS**（A–F・コーパス限定） | ファイル削除 |
| 2026-08-24 | `docs/org/B_PURCHASE_RESEARCH_SINGLE_NAME_REQUIREMENTS.md` | **単品名・JAN 要件定義**（存在=Catalog・F不一致・コーパスはAI迷い限定） | docs revert |
| 2026-08-24 | `docs/org/B_PURCHASE_RESEARCH_SINGLE_NAME_REQUIREMENTS.md` | **単品名・JAN**: 存在＝Catalog正本・公式不採用・JAN複数パターン再整理 | docs revert |
| 2026-08-24 | `docs/org/B_PURCHASE_RESEARCH_SINGLE_NAME_REQUIREMENTS.md` | **単品名・JAN充填** 要件メモ案（AI先→JAN照合。E前） | docs revert |
| 2026-08-24 | `PurchaseResearch.js`／`d2_filter.py` | **D2**: 大型・特大ティア／重量上限g（①プロファイル・欠測は見送りにしない） | Git revert。要 clasp |
| 2026-08-28 | `PurchaseResearch.js`／liveness／HUMAN_RUN | **予定超過リセット**（Keepa最大回復+5分）・A-1単独復旧・token待ち短縮・完走→C4別実行維持 | Git revert。要 clasp |
| 2026-08-28 | `PurchaseResearch.js`／HUMAN_RUN | **洗いトリガー予約前 Keepa `/token` ログ**（liveLeft vs calcLeft）。待ち妥当性の切り分け | Git revert。要 clasp |
| 2026-08-27 | `PurchaseResearch.js`／HUMAN_RUN | **メニュー整理**: ルート「自動リサーチ」A全自動/B個別/C見積枠/99。A-1即心拍・A-2全診断 | Git revert。要 clasp |
| 2026-08-27 | `cycle_heartbeat_liveness.py`／`PurchaseResearch.js` | **心拍全体の沈黙**: 生存番犬15分・D失敗→心拍・no_maker再武装。仮想15/15 | Git revert。要 clasp |
| 2026-08-27 | `cycle_trigger_watchdog.py`／`PurchaseResearch.js` | **洗いトリガー沈黙**: inflight・phaseログ・二重発火skip・outer catch再武装。仮想12/12 | Git revert。要 clasp |
| 2026-08-24 | `docs/org/B_PURCHASE_RESEARCH_*` | **C5** CYCLE docs 整備（C3+T 本線・実機合格・HUMAN_RUN §10） | docs revert |
| 2026-08-22 | `コード.js`／STAGE | **JANカテゴリ**: 口コミ1位優先。レビュー差&lt;max(10,rc2×0.25)なら多数決。要 clasp | Git revert |
| 2026-08-22 | `コード.js`／STAGE | **Yahoo検索sort分離**: JAN=`-review_count`（口コミ）、文言=`-score`（おすすめ）。要 clasp | Git revert |
| 2026-08-23 | `PurchaseResearch.js` | **C3+T**（C3+完走トリガー1メニュー）・`prArmWashFullTriggerForContinue_` | Git revert。要 clasp |
| 2026-08-23 | `cycle_c4_chain.py`／`PurchaseResearch.js` | **C4完走連鎖**（完走→D1→D2→次社 P3 arm）。`PR_CYCLE_C4_ENABLED` | Git revert。要 clasp |
| 2026-08-23 | `cycle_c3_select.py`／`PurchaseResearch.js` | **C3自動選定**・実態プルダウン・マスタ4社クリアメニュー | Git revert。要 clasp |
| 2026-08-23 | `maker_jitta_resolve.py`／`schema.py` | **Amazonルート＋Browse棚の2列**・業種フォールバック（不明）・dry 6PASS | Git revert |
| 2026-08-23 | `docs/org/B_MAKER_JITTA_CLASSIFICATION.md`／`schema.py` | **実態仕分け正本**（平易版・マスタ/サマリミラー列・AI判定・永谷園/味の素/あさくま事例） | docs revert |
| 2026-08-22 | `PurchaseResearch.js`／FLOW／HUMAN_RUN | **C1+C2 GAS**: A-light（query+product cap・門/①候補/Keepaフルなし）→ セラーヒット全店再計算。`PR_ALIGHT_*` Property | `PR_ALIGHT_ENABLED=false`／Git revert。要 clasp |
| 2026-08-22 | `tools/competitor_store/dry_cycle_seller_accum.py` | **C0 fixture 4件**（基本/cap/429/除外）。GAS 前の正本 dry | ファイル削除 |
| 2026-08-22 | `PurchaseResearch.js`／FLOW §5.4／HUMAN_RUN | **C0 セラー貯め GAS**: 門通過→BB列+SP-API新品→`セラー`新規100/回。調査CK OFF。429/時間切れ再開 Property | Prop切戻し／Git revert。要 clasp |
| 2026-09-01 | `sub_image_food_canon_gate.py` | **見本ゲート**: 人間確定済みを再オープン時にレ点復元＋緑枠／前回バナー／手動画像を候補へ再同期 | Git revert |
| 2026-09-01 | `コード.js` | **B-④ ①′**: 見本画面表示と同時に `buildB5ProposalsForBindings` 裏実行（品番順・見本と同時進行） | Git revert。要 clasp |
| 2026-09-01 | `コード.js`／`b5_food_canon_status.py`／ALIGN §7.12.3・§7.15 | **B-④ ①′ 個品見本ゲート**: ダイアログ差し込み・Drive 状態 JSON・未確定ブロック。§7.15 ミックス5品番ドラフト | Git revert。要 clasp |
| 2026-08-22 | `tools/competitor_store/dry_cycle_alight_hit.py` | C1+C2 仮想 fixture 7件（A-light+ヒット再計算） | ファイル削除 |
| 2026-08-22 | `docs/org/B_PURCHASE_RESEARCH_CYCLE_DEV_MEMO.md` | C0–C5 仮想テスト方針 | docs revert |
| 2026-08-22 | `docs/org/B_PURCHASE_RESEARCH_CYCLE_*`／FLOW | **3レーン運用サイクル**要件＋3者多数決（条件付きYES）。A-light・セラーヒット再計算・完走連鎖・⑧週待ち廃止 | docs revert。実装 C0–C5 は別承認 |
| 2026-08-22 | `PurchaseResearch.js` | **C0 `prAppendSellerAccumRows_` setValues 修正**（37行/826行不一致。getRange 第3引数=行数） | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／V1／HUMAN_RUN | **門**: ①プロファイルに順位下限（初期1・欠測落ち）＋価格上限（初期10万・カテゴリ別可）。列自動確保 | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／STAGE | **Yahoo/Ichibaクエリ順**: JAN常に先頭（ミックスも）。FBはヒット率順（メーカー名→名→KW→N種）。要 clasp | Git revert |
| 2026-08-22 | `yahoo_category_virt/hit_rate_*` | **Yahooクエリヒット率調査**: パターン表P01–P12＋API bakeoff（読取のみ）。sample=行20/47/110相当 | ツール削除で戻す。clasp不要 |
| 2026-08-22 | `PurchaseResearch.js`／HUMAN_RUN／schema | **Keepaフル索引B1**: 書込時ASIN列だけ見て`Keepaフル索引`自動再構築。D1は必要ASINだけ読込。Prop `PR_KEEPA_FULL_INDEX_ENABLED`。99に再構築メニュー | Prop=false／Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／HUMAN_RUN | **Keepa単価**: プロファイル`product_offers20`既定6（旧2）。再開待ちは実測consumed/件数優先。残tokで次スライス継続（sliceWait≥5即終了廃止） | Prop上書き／Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js` | **店ゲート Keepa消費を①ログへ**: tokensConsumed列＋detailに qTok/pTok/consumed（offers=20の店調査が洗いと別枠で見える） | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js` | **洗いトリガー予約修正**: arm時はProp先ON→トークン見積で予約。①ログに`トリガー予約`/`trigger_armed afterMin=`。A→D未武装時も再予約 | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／schema／FULL | **実態ck**: 実態=要確認のみ Amazon検索URL（列名`実態ck`・末尾追加。M列実態は不変）。メニュー99「実態ck付与」 | Git revert。要 clasp |
| 2026-08-22 | FLOW §7.1 | **全店自動の最終形=案B**: 日次自動選定＋調査CK自動オン＋消化無人。日次件数上限で抑止。実装はモリタ後 | docs revert |
| 2026-08-22 | `コード.js`／STAGE | **Yahoo検索**: 打ち切り廃止・最大6qマージ・上位20・sort=-score（おすすめ順）。ログ BRANCH2_MERGE | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／`yahoo_category_virt`／STAGE | **Yahooカテゴリ改修**: 点数制・ゆるいシリーズ第1FB・表記ゆれ・Driveガード。仮想9PASS（BASELINE比） | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／`tools/yahoo_category_virt`／STAGE | **Yahooカテゴリ**: 同シリーズ層（+1）＋同商品（+2）。メーカー門。仮想6PASS | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／`tools/yahoo_category_virt`／STAGE | **Yahooカテゴリ**: 同商品フィルタ→投票→票内重み。食品3a廃止。仮想4PASS（旧ロジックは41392等にFAIL） | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／マスタGB-GD数式／STAGE | **診断ログ**: SKU監査／ミックスCOMPARE+Gemini／Yahoo BRANCH1-5+食品path優先／DUP_URL。GB「・」偽重複を数式修正 | Git revert＋数式戻し。要 clasp |
| 2026-08-22 | `docs/org/B_RERUN_CLEAR_AND_JUDGE_HUMAN_RUN.md` | **B再実行**: 消去列チェックリスト＋再B合否（G1–G7・ログ早見） | docs revert |
| 2026-08-23 | `CURRENT_PHASE` §0ab／`B_RERUN_CLEAR_AND_JUDGE` §4.3 | **G7保留**: ミックス楽天7207・同一URL複数親＝楽天に正セット無し時の汎用ヒット。非クリティカル・コード改修は問題拡大時まで | docsのみ |
| 2026-08-22 | `コード.js`／CURRENT／YAHOO_STAGE | **A→B検収FIX**: Yahoo選定ログ＋SHP数字query回避／一人分先頭数値／Step1色＝子SKU判定／ミックスWRITEログ | Git revert。要 clasp（出品） |
| 2026-08-22 | `PurchaseResearch.js` | **実態シード**（maker_alias同期）＋サマリ空欄へ実態付与。未登録=要確認。A→Dは製造/ブランド優先・Generic除外 | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／FLOW／HUMAN_RUN | **本線メニュー**: 1.店1件通し／0.試験A→D(メーカー最大1)／5.D1→D2。旧は99サブ。全店自動はFLOW§7.1メモのみ | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／FULL／HUMAN_RUN | **洗いGET offers=20**＋**D1案A**（フル行あればenrichしない）。単価既定2。薄い行は指紋同でもoffers上書き | Prop `KEEPA_PRODUCT_TOKEN_COST`／Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／FLOW／HUMAN_RUN | **洗い完走で調査CK解除**（トリガー完走含む）。**D1→D2 統合メニュー**（単独D1/D2は残置） | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL／HUMAN_RUN | **洗いキュー期限切れ削除**: Keepa済＋列挙日時>既定90日。メニュー＋完走時。Prop `PR_WASH_QUEUE_RETENTION_DAYS`（0=オフ） | Prop=0／Git revert。要 clasp |
| 2026-08-21 | `コード.js`／COMBO／SETCOUNT | **ミックスセット数**: A列＝何セットか（各N→N）。計N袋不使用。Vision＋参考画像ゲート・上限20。単品パース分離 | Git revert／Limit Property。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／プロファイル／FULL | 洗い完走ロック・トリガー予約のみ・Keepa upsert・門出品者数1–20・D1 offers補強 | トリガー停止／Prop。要 clasp |
| 2026-08-21 | `PurchaseResearch.js` | `PR_MAKER_WASH_FULL_PROP_` 欠落でトリガー ReferenceError → 定数追加 | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL／HUMAN_RUN | 洗い完走トリガー待ち: 残り×単価（上限300）÷rate＋1。旧30分キャップ撤廃 | Prop / Git revert。要 clasp |
| 2026-08-21 | `purchase_research/appsscript.json` | P3用に `script.scriptapp` スコープ追加（`getProjectTriggers` 権限エラー対策） | スコープ行を削除して clasp push。要**再承認** |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **P3 洗い完走トリガー**（開始／停止／完走解除）。Keepa 429→RATE_LIMIT 温和停止 | Prop `PR_WASH_FULL_TRIGGER_ENABLED=false`＋トリガー停止メニュー。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **洗い完走ループ**メニュー。1実行内スライス反復（トリガーなし・25分上限） | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **続きがKeepaに進むよう修正**: needEnum見直し・分割済の再推定禁止・キュー充分時は分割を済扱い | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **洗いキュー**: ASIN重複削除＋Keepa未済ユニーク＋同ASIN一括済。行増はユニーク増加（重複ではない） | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **分割語 nextToken 再開**。続きはサマリ→分割語→キューの順で対象選定 | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js` | **P1bバグ**: `prWriteSplitWords_` の getRange 第3を終端行扱い→「15行 vs 16行」で FAILED。numRows に修正 | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／dry_p1b／FULL | **洗い完走 P1b**: 分割B（タイトル→Gemini→OpenAI）・`①洗い分割語`・完走消込。尾西ドライ PASS | `PR_MAKER_WASH_FULL_ENABLED=false`。要 clasp＋①に GEMINI_API_KEY |
| 2026-08-21 | `コード.js` | **画像袋数 JSON修復**: フェンス除去・途中切れフィールド救出・MIME json・maxTokens512 | Git revert。要 clasp |
| 2026-08-21 | `コード.js` | **画像袋数ログ**: 不採用時も conf/set/texts/reason/img を Logger 出力 | Git revert。要 clasp |
| 2026-08-21 | `コード.js`／COMBO／SETCOUNT | **画像袋数改**: 抜き出しJSON・単位一般・highのみ採用。単品12-⑭もタイトル不能＋実行上限5 | Git revert／Property=false。要 clasp |
| 2026-08-21 | `コード.js`／COMBO | **ミックス袋数 Vision**: タイトル不能時のみ画像AI（親上限5・Yahoo優先）。`MIX_COMPETITOR_SET_COUNT_FROM_IMAGE` 未設定=ON | Git revert／Property=false。要 clasp |
| 2026-08-21 | `コード.js`／COMBO | **ミックス競合修正**: 除外後0件は次クエリ。N種/アソート標題のみ採用。子の楽天Yahooを空にしてから書込（単品N袋誤載せ・足し算残骸対策） | Git revert。要 clasp |
| 2026-08-21 | docs FULL／V1／FLOW | **洗い分割B**: タイトル→カテゴリ推定。Gemini主・OpenAI失敗時のみ。分割語シート＋消込 | docs revert |
| 2026-08-20 | `PurchaseResearch.js`／dry_p1／FULL | **洗い完走 P1+P2コード**。P1ドライ: 石原PASS／尾西 need_q2 でP2ドライ未着手 | `PR_MAKER_WASH_FULL_ENABLED=false`。要 clasp |
| 2026-08-20 | docs（FULL／FLOW／V1） | **メーカー洗い完走（案A+C）** 要件ロック。`①洗いキュー`・第2語分割・完走時のみ最終調査日。実装は別承認 | docs revert |
| 2026-08-20 | `コード.js`／COMBO／STORE | **ミックス楽天Yahoo**: N種モール検索→袋数一致の価格＋URL。0件は空（足し算なし）。R12足し算はAmazonのみ。選択・よりどり除外強化 | Git revert。要 clasp |
| 2026-08-20 | Keepaフル／ヒット／snapshot／①／Paste／Python | **先頭挿入**（最新=上）。セラー末尾のまま。メーカーサマリ新規のみ先頭。`dry_maker_sum --apply` 全消しは非本線 | Git revert。要 clasp（出品＋①） |
| 2026-08-20 | docs／`コード.js`／`PurchaseResearch.js` | **Keepa**: 応答優先。実測 refillRate=5 をフォールバック。洗い cap20。待ち=ceil(消費/rate) | docs／Property／Git revert。要 clasp |
| 2026-08-20 | docs／`コード.js`／`PurchaseResearch.js` | **Keepa Starter=20/分表記**（後に実測5へ訂正） | docs／Property／Git revert。要 clasp（出品＋①） |
| 2026-08-20 | docs／`コード.js` コメント | **Keepa トークン表現**: 日次枠ではない。毎分 `refillRate`・未使用60分失効（旧5/分表記はStarter20へ更新） | docs／Git revert |
| 2026-08-20 | `コード.js` | **競合 JAN≤5**: 1ページ目生ヒット5件以下ならメーカー+商品名をマージ（13-②/Yahooセット別/B Step2）。14-②は除外 | Git revert。要 clasp |
| 2026-08-20 | `コード.js` | **Yahoo 競合**: 楽天同順フォールバック（jan_code→query=JAN→メーカー+商品名→商品名）。14-②/セット別/B Step2/定時。Amazon◎は不変更 | Git revert。要 clasp |
| 2026-08-20 | `コード.js`／`bakeoff_rakuten_yahoo.py` | **楽天 Ichiba**: API版 `20220601`→`20260701`（旧版は Configuration not found）。Property `RAKUTEN_ICHIBA_ITEM_SEARCH_API_VERSION` で上書き可 | Property=`20220601`／Git revert。要 clasp |
| 2026-08-20 | `コード.js` | **楽天ジャンル（メニュー8）**: 1ロジック×多段フォールバック（セット親はN種＋JAN最後／単品はJAN優先）。投票フィルタ汎用化。Yahoo 7.6 共用 | Git revert。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **未調査メーカー自動**: レ点不要・90日/空を1件選定→洗い→D1（`auto_maker_catalog`） | `PR_AUTO_MAKER_ENABLED=false`。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **D1 Q/P列**: 出品 ASIN貼り付けのセット数量 regex を移植＋alt（種×袋・5食）。P=販売形態ラベル | Git revert。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **D1 行2式**: skipMap 列の式が prListWriteRecs_ で空消しされていたバグ修正 | Git revert。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **D1 品番リスト**: 画像URL=Keepaサブ画像からHYPERLINK、税率=0.08、挿入時に2行目の式コピー | Git revert。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **D1 FBA列**: L1 first-fit（設定マスタ）＋Keepa FBA手数料フォールバック。空出品FBA手数料バグ修正 | Git revert。要 clasp |
| 2026-08-19 | `purchase_research/PurchaseResearch.js` | ブランド型メーカー洗い: 入口語の `page0=` 除去・Catalog entry 正規化・`PR_BRAND_NEEDLES_` 明示 needle | Git revert。要 clasp |
| 2026-08-19 | `tools/competitor_store/maker_alias.py` | 入口語ヒットを `brand/親 page0=…` 形式に（GAS split 整合） | Git revert |
| 2026-08-19 | `tools/common_ss_view/` | 共通スプシ参照の registry + dump。競合ツールの ID は registry 経由 | Git revert `tools/common_ss_view/` |
| 2026-08-19 | `コード.js`／COMBO／B_SPAPI | **ミックス親 FBA**: Catalog梱包空ならセット構成3辺で P1b 書込 | `B_FBA_P1B_COMBO_PACK_FALLBACK_ENABLED=false`／Git revert。要 clasp |
| 2026-08-19 | `コード.js`／COMBO／B_SPAPI | **ミックス親 FBAティア**: 1袋子ASINでCatalog。親競合ASIN非書。CQ自己発は子のまま | `B_FBA_COMBO_CHILD_ASIN_ENABLED=false`／Git revert。要 clasp |
| 2026-08-19 | メニュー成功は toast（alert は失敗のみ）。OK 待ちで別SS未反映を解消 | Git revert。要 clasp |
| 2026-08-19 | セラー再調査は機械欄上書き・同一IDは1行。切戻し `PR_SELLER_REWALK_OVERWRITE=false` | Git revert。要 clasp |
| 2026-08-19 | 同一sellerIdのA①空行へ店名行の構成比をコピー（採取元は残す） | Git revert。要 clasp |
| 2026-08-19 | レ点なしアラートに読込行・TRUE件数。同じsellerIdは店名あり行から補完 | Git revert。要 clasp |
| 2026-08-19 | 店GETはレ点の店のみ。onOpenはメニューだけ。CK枠の毎回再挿入をやめる | Git revert。要 clasp |
| 2026-08-18 | ① GAS メニューは25分で leftover（課金30分）。onOpen は30秒 | Git revert。要 clasp |
| 2026-08-18 | 店名欠けるIDは /seller storefront=0 で取り台帳へ | `PR_SELLER_NAME_GET_ENABLED=false`。要 clasp |
| 2026-08-18 | ①候補に店名列（sellerIdは参考・`|` 連結） | Git revert。要 clasp |
| 2026-08-18 | ①候補 1ASIN1行・sellerId は `|` 連結 | Git revert。要 clasp |
| 2026-08-18 | 調査CK にチェック枠＋レ点行を青字 | 条件付き書式削除／validation解除 |
| 2026-08-18 | ① GAS レ点の店／メーカー（調査CK・1000超は次へ・再巡回仮6か月） | `PR_CK_SELLER_ENABLED=false`／`PR_CK_MAKER_ENABLED=false`。要 clasp |
| 2026-08-18 | D1 新規を行5挿入（既存下へ） | Git revert。要 clasp |
| 2026-08-18 | E 実装保留メモ FLOW §5.3.1 | docs revert |
| 2026-08-18 | ① GAS D2 機械絞込 `prMenuListFilterD2` | `PR_LIST_FILTER_ENABLED=false`／列クリア。要 clasp |
| 2026-08-18 | ① GAS D1 品番リスト追記（実機PASS） | `PR_LIST_APPEND_ENABLED=false`／追記行削除。要 clasp |
| 2026-08-17 | ① GAS Keepaフル追記（GET分） | `PR_KEEPA_FULL_WRITE_ENABLED=false`／末尾行削除。要 clasp |
| 2026-08-17 | ① GAS メーカー洗い cap5（Catalog→門） | `PR_MAKER_WASH_ENABLED=false`／Git。要 clasp |
| 2026-08-17 | 接続スイート T2–T4（AGF/リフココ/BBQ） | ①候補末尾5（T2）／実行ログ | Git／行削除 |
| 2026-08-17 | 接続T1 尾西食品 cap5 | ①候補末尾5／フル+4 | Git／行削除 |

| 2026-08-16 | メーカーサマリ2層（A/B） | タブ削除／clear | Git revert |

| 2026-08-18 | R12 ミックス寸法: 長辺・短辺は最大、最短のみ数量倍 | Git revert。要 clasp |
| 2026-08-18 | R12 ミックスKeepa◎袋はAmazonその値段／物流は親1袋3辺 | Git revert。要 clasp |
| 2026-08-18 | R12 ミックスは12-⑭非書／1袋欠けは別モール1袋借用 | Git revert。要 clasp |
| 2026-08-18 | R12 競合子行セット1／Bキュー分離／構成ゲート | Git revert。要 clasp |
| 2026-08-18 | R12 大きさは単品親の競合Keepa。同シリーズは1袋のみ流用 | Git revert。要 clasp |
| 2026-08-17 | R12 構成照合は商品名（セット税込＝和・照合キーにしない） | Git revert。要 clasp |
| 2026-08-16 | R12 組み合わせ（色・Keepa＝AI行・構成記入・足し算） | Git revert。要 clasp |
| 2026-08-16 | Peacefulショップ 一連・リスト+8 | 8行目クリア／リスト末尾8 | セル・行戻し |
| 2026-08-16 | 門に使う順位段階を文字列プルダウン | 数字1に戻す／検証解除 | Git＋シート。要 clasp |
| 2026-08-16 | ストアASIN数は応答がある店だけ | docs戻し | Git revert |
| 2026-08-16 | Camel Trading 台帳列＋リスト+8 | 7行目クリア／リスト末尾8 | セル・行戻し |
| 2026-08-16 | 台帳→構成比 Camel Trading 既存行 | 店名・食品比率等クリア | セル戻し |
| 2026-08-16 | リフココ Keepa+20・リスト+14。いしとき query1 | リスト末尾14／台帳 | 行削除 |
| 2026-08-16 | 輪1周 A18COIYGTMT3AC Keepa20・リスト+11 | リスト末尾11行／台帳巡回日 | 行削除／巡回日クリア |
| 2026-08-16 | A① 2秒再開+7・まとめ+34・429 | セラー末尾41行付近 | Git revert |
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
