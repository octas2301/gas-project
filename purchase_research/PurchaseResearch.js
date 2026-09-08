/**
 * 仕入れ検討①。出品 コード.js には置かない。
 * /query は GET + selection= のみ。POST {"selection":} は使うな。
 */
var PR_LOG_SHEET_ = '①ログ';
var PR_CAND_SHEET_ = '①候補';
var PR_PROF_SHEET_ = '①プロファイル';
var PR_FOOD_ROOT_ = '57239051';
var PR_SELLER_MORITA_ = 'AYC4Z8PML8T30';
var PR_KEEPA_QUERY_ = 'https://api.keepa.com/query';
var PR_KEEPA_PRODUCT_ = 'https://api.keepa.com/product';
var PR_KEEPA_SELLER_ = 'https://api.keepa.com/seller';
/** Keepa 残量確認（消費0）。予約前の診断用。 */
var PR_KEEPA_TOKEN_ = 'https://api.keepa.com/token';
var PR_SELLER_NAME_GET_ENABLED_PROP_ = 'PR_SELLER_NAME_GET_ENABLED';
var PR_SELLER_NAME_GET_CAP_PROP_ = 'PR_SELLER_NAME_GET_CAP';
var PR_COMPETITOR_SS_ = '1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs';
/** 出品マスタ（00_設定マスタ FBA手数料表。L1 と同じ first-fit） */
var PR_LISTING_SS_ = '1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28';
var PR_SETTINGS_MASTER_SHEET_ = '00_設定マスタ';
var PR_FBA_TABLE_CACHE_ = null;
var PR_LIST_SHEET_ = '仕入検討品番リスト';
var PR_LIST_HEADER_ROW_ = 4;
var PR_LIST_DATA_START_ = 5;
/** 品番リストの計算式テンプレ行（D1 挿入時に PASTE_FORMULA の正本） */
var PR_LIST_FORMULA_TEMPLATE_ROW_ = 2;
var PR_ORIG_SS_ = '1Uk-DEfoSTVqb5N8XsfSqciaZcIjzbbPpGvg4hQYgDl4';
var PR_LIST_APPEND_ENABLED_PROP_ = 'PR_LIST_APPEND_ENABLED';
/** 品番リストの差分判定で数式セルを数式として比べる。未設定=ON。false で旧動作。 */
var PR_LIST_DIFF_USE_FORMULA_PROP_ = 'PR_LIST_DIFF_USE_FORMULA';
/** Keepaフル再門メニュー。未設定=ON。false で切戻し。 */
var PR_REGATE_FROM_FULL_ENABLED_PROP_ = 'PR_REGATE_FROM_FULL_ENABLED';
/** true=門結果空白行のみ。false=全行を新門で上書き（通過→落ちも可）。 */
var PR_REGATE_ONLY_BLANK_PROP_ = 'PR_REGATE_ONLY_BLANK';
var PR_LIST_APPEND_CAP_PROP_ = 'PR_LIST_APPEND_CAP';
var PR_LIST_FILTER_ENABLED_PROP_ = 'PR_LIST_FILTER_ENABLED';
var PR_LIST_FILTER_CAP_PROP_ = 'PR_LIST_FILTER_CAP';
/** D2 1回あたりの未判定 ASIN 上限（Property 未設定時）。 */
var PR_LIST_FILTER_CAP_DEFAULT_ = 1000;
/** true=見積候補／見送り済み行は D2 で触らない。false=全 ASIN 再判定。 */
var PR_D2_ONLY_BLANK_PROP_ = 'PR_D2_ONLY_BLANK';
/** D2 上限超で C4 を止めた parent（| 区切り）。解消まで C4 recover しない。 */
var PR_CYCLE_D2_OVERFLOW_PROP_ = 'PR_CYCLE_D2_OVERFLOW';
var PR_LIST_COL_QUOTE_ = '見積候補';
var PR_LIST_COL_SKIP_REASON_ = '見送り理由';
/** 単品商品名 AI 充填（見積前 S）。未設定=ON。 */
var PR_SINGLE_NAME_ENABLED_PROP_ = 'PR_SINGLE_NAME_AI_ENABLED';
var PR_SINGLE_NAME_CAP_PROP_ = 'PR_SINGLE_NAME_CAP';
/** 1回あたりの AI 件数。未設定=100（続きは再実行／D2後自動）。 */
var PR_SINGLE_NAME_CAP_DEFAULT_ = 100;
/** D2 完了後に単品名充填を続ける。未設定=ON。 */
var PR_SINGLE_NAME_AFTER_D2_PROP_ = 'PR_SINGLE_NAME_AFTER_D2';
var PR_LIST_COL_SINGLE_NAME_ = '単品商品名';
var PR_LIST_COL_SINGLE_TRUST_ = '単品名信頼';
var PR_LIST_COL_SINGLE_PATTERN_ = '単品名パターン';
var PR_KEEPA_FULL_SHEET_ = 'Keepaフル';
/** Keepaフルの ASIN→行 自動索引（人手メンテなし。書込時にASIN列だけ見て再構築）。 */
var PR_KEEPA_FULL_INDEX_SHEET_ = 'Keepaフル索引';
var PR_KEEPA_FULL_INDEX_HEADERS_ = ['ASIN', '行'];
var PR_KEEPA_FULL_INDEX_PROP_ = 'PR_KEEPA_FULL_INDEX_ENABLED';
/** Keepaフルの空行掃除を「ASIN列1回読み→候補だけ全列」にする（2026-09-07）。false で旧1行ずつ。 */
var PR_KEEPA_FULL_BLANK_SCAN_BATCH_PROP_ = 'PR_KEEPA_FULL_BLANK_SCAN_BATCH';
var PR_MAKER_SUM_SHEET_ = 'メーカーサマリ';
var PR_SELLER_SHEET_ = 'セラー';
var PR_CK_HEADER_ = '調査CK';
/** total≥501 は特大店（フルしない）。旧1000超holdを前倒し。 */
var PR_QUERY_TOTAL_HOLD_ = 501;
var PR_CK_SELLER_ENABLED_PROP_ = 'PR_CK_SELLER_ENABLED';
var PR_CK_MAKER_ENABLED_PROP_ = 'PR_CK_MAKER_ENABLED';
var PR_CK_BATCH_CAP_PROP_ = 'PR_CK_BATCH_CAP';
var PR_AUTO_MAKER_ENABLED_PROP_ = 'PR_AUTO_MAKER_ENABLED';
var PR_AUTO_MAKER_CAP_PROP_ = 'PR_AUTO_MAKER_CAP';
var PR_AUTO_MAKER_REWALK_DAYS_PROP_ = 'PR_AUTO_MAKER_REWALK_DAYS';
var PR_SELLER_REWALK_DAYS_PROP_ = 'PR_SELLER_REWALK_DAYS';
var PR_SELLER_REWALK_OVERWRITE_PROP_ = 'PR_SELLER_REWALK_OVERWRITE';
/** C0 セラー貯め（FLOW §5.4）。門通過 ASIN の BB 列＋SP-API 新品オファー。Keepa GET なし。 */
var PR_SELLER_ACCUM_ENABLED_PROP_ = 'PR_SELLER_ACCUM_ENABLED';
var PR_SELLER_ACCUM_CAP_PROP_ = 'PR_SELLER_ACCUM_CAP';
var PR_SELLER_ACCUM_SLEEP_MS_PROP_ = 'PR_SELLER_ACCUM_SLEEP_MS';
var PR_SELLER_ACCUM_AFTER_ASIN_PROP_ = 'PR_SELLER_ACCUM_AFTER_ASIN';
var PR_SELLER_ACCUM_BATCH_PROP_ = 'PR_SELLER_ACCUM_BATCH';
var PR_SELLER_ACCUM_LOOP_PROP_ = 'PR_SELLER_ACCUM_LOOP';
var PR_SELLER_ACCUM_429_WAIT_MS_PROP_ = 'PR_SELLER_ACCUM_429_WAIT_MS';
var PR_SELLER_ACCUM_429_RETRY_MAX_PROP_ = 'PR_SELLER_ACCUM_429_RETRY_MAX';
var PR_SELLER_ACCUM_CAP_DEFAULT_ = 100;
var PR_SELLER_ACCUM_CAP_MAX_ = 100;
var PR_SELLER_ACCUM_SLEEP_DEFAULT_MS_ = 1100;
var PR_SELLER_ACCUM_429_WAIT_DEFAULT_MS_ = 90000;
var PR_SELLER_ACCUM_429_RETRY_DEFAULT_ = 5;
var PR_SELLER_ACCUM_SKIP_IDS_ = { ATVPDKIKX0DER: true, A1VC38T7YXB528: true };
/** C1 A-light + C2 セラーヒット（CYCLE §4–§5）。門・①候補・Keepaフル upsert なし。 */
var PR_ALIGHT_ENABLED_PROP_ = 'PR_ALIGHT_ENABLED';
var PR_ALIGHT_SHOP_CAP_PROP_ = 'PR_ALIGHT_SHOP_CAP';
var PR_ALIGHT_PRODUCT_CAP_PROP_ = 'PR_ALIGHT_PRODUCT_CAP';
var PR_MAKER_SUM_HIT_COL_ = 'セラーヒット';
/** 1実行あたり店数。未設定=1（1店フル）。 */
var PR_ALIGHT_SHOP_CAP_DEFAULT_ = 1;
var PR_ALIGHT_PRODUCT_CAP_DEFAULT_ = 20;
var PR_ALIGHT_PRODUCT_CAP_MAX_ = 30;
var PR_ALIGHT_DAILY_TARGET_PROP_ = 'PR_ALIGHT_DAILY_TARGET';
var PR_ALIGHT_DAILY_MAX_PROP_ = 'PR_ALIGHT_DAILY_MAX';
var PR_ALIGHT_NO_MAKER_ROUNDS_MAX_PROP_ = 'PR_ALIGHT_NO_MAKER_ROUNDS_MAX';
var PR_ALIGHT_DAILY_TARGET_DEFAULT_ = 5;
var PR_ALIGHT_DAILY_MAX_DEFAULT_ = 10;
var PR_ALIGHT_NO_MAKER_ROUNDS_DEFAULT_ = 10;
var PR_ALIGHT_DAILY_DONE_DAY_PROP_ = 'PR_ALIGHT_DAILY_DONE_DAY';
var PR_ALIGHT_DAILY_DONE_COUNT_PROP_ = 'PR_ALIGHT_DAILY_DONE_COUNT';
var PR_ALIGHT_ROUND_DAY_PROP_ = 'PR_ALIGHT_ROUND_DAY';
var PR_ALIGHT_ROUND_COUNT_PROP_ = 'PR_ALIGHT_ROUND_COUNT';
var PR_ALIGHT_LAST_NEW_PROP_ = 'PR_ALIGHT_LAST_NEW_MAKERS';
var PR_ALIGHT_CONT_PROP_ = 'PR_ALIGHT_CONT_JSON';
var PR_SELLER_BAND_S_MAX_ = 100;
var PR_SELLER_BAND_M_MAX_ = 300;
var PR_SELLER_BAND_L_MAX_ = 500;
var PR_SELLER_BAND_COL_ = '規模帯';
var PR_SELLER_SKU_COL_ = '在庫SKU見積';
var PR_SELLER_SKU_BREAKDOWN_COL_ = 'SKU見積内訳';
var PR_SELLER_BAND_BREAKDOWN_COL_ = '規模帯内訳';
/** ①プロファイル: 店 /query の調査カテゴリ（CYCLE §4.1bis）。 */
var PR_PROF_SURVEY_CK_HEADER_ = '調査カテゴリー';
var PR_PROF_ROOT_ID_HEADER_ = 'AmazonルートID';
var PR_FRESH_MS_ = 90 * 24 * 60 * 60 * 1000;
/** ①メニューは課金枠30分前提。強制終了前に再予約できるよう **20分** で止めてログを残す。onOpen（シートを開く）は常に30秒で別。 */
var PR_GAS_BUDGET_MS_ = 20 * 60 * 1000;
var PR_GAS_START_MS_ = 0;
var PR_KEEPA_FULL_MAP_CACHE_ = null;
var PR_KEEPA_FULL_INDEX_CACHE_ = null;
var PR_PROF_CACHE_ = null;
var PR_CK_READY_ = {};
var PR_CAND_MEM_ = null;
/**
 * 選定プール件数（小中・不明・大・特大）の実行内キャッシュ。
 * 2026-09-07: 1実行で prCyclePoolCounts_ が10か所以上から呼ばれ、そのたびに
 * メーカーサマリ全社ぶんの規模数え直し（＝洗いキュー全読み×社数）が走っていた。
 * 実測: 選定 pick 12.8分＋pick後 11.4分＝24分 > 予算20分 → Keepa に到達せず空振り。
 * 切戻し: Property PR_CYCLE_POOL_REUSE=false（毎回数え直す旧動作）。
 */
var PR_CYCLE_POOL_CACHE_ = null;
var PR_CYCLE_POOL_REUSE_PROP_ = 'PR_CYCLE_POOL_REUSE';
/**
 * 洗いキューの「メーカー→ユニークASIN数」実行内キャッシュ。
 * 旧: prWashQueueAsinSet_(parent) がメーカー1社ごとに全行（約26,000行）を読み直していた。
 * 切戻し: Property PR_WASH_QUEUE_COUNT_MAP=false（1社ずつ全読みする旧動作）。
 */
var PR_WASH_QUEUE_COUNT_CACHE_ = null;
var PR_WASH_QUEUE_COUNT_MAP_PROP_ = 'PR_WASH_QUEUE_COUNT_MAP';
/**
 * 洗いスライス用の「対象メーカーの行」実行内スナップショット（{parent: {rows, idxK}}）。
 * 旧: 1スライスで 重複整理／未済取得／Keepa済マーク／残数＋ループ前後 の計5〜6回、
 * キュー全行（約26,000行）を読み直していた。実測 1スライス=22.6〜28.8分（件数に無関係）。
 * 切戻し: Property PR_WASH_QUEUE_SNAPSHOT=false（呼ばれるたび読み直す旧動作）。
 */
var PR_WASH_QUEUE_SNAP_ = {};
var PR_WASH_QUEUE_SNAPSHOT_PROP_ = 'PR_WASH_QUEUE_SNAPSHOT';
/**
 * メーカーサマリの「メーカー→規模帯／洗い進捗／Catalog見積」実行内キャッシュ。
 * 旧: prMakerBandFromSum_ が1社ごとに prFindMakerSumRow_（サマリ全読み）＋ヘッダ＋当該行の
 * 3回読みをしていた。選定はキューに出てくる社ごとにこれを呼ぶので、
 * 洗える社が無い（全社後回し）ときの pick が最大11.9分かかり、その間 Keepa は完全に空いていた。
 * 実測 2026-09-08 夜間: 空振り pick 24回。1社見つかる pick は19秒、全社後回しは2〜11.9分。
 * 切戻し: Property PR_MAKER_BAND_MAP=false（1社ずつ読み直す旧動作）。
 */
var PR_MAKER_BAND_CACHE_ = null;
var PR_MAKER_BAND_MAP_PROP_ = 'PR_MAKER_BAND_MAP';
var PR_FROZEN_RE_ = /冷凍|冷蔵|生鮮/;
var PR_SPAPI_LWA_URL_ = 'https://api.amazon.com/auth/o2/token';
var PR_SPAPI_DEFAULT_ENDPOINT_ = 'https://sellingpartnerapi-fe.amazon.com';
var PR_SPAPI_DEFAULT_MP_ = 'A1VC38T7YXB528';
var PR_MAKER_WASH_ENABLED_PROP_ = 'PR_MAKER_WASH_ENABLED';
var PR_MAKER_WASH_CAP_PROP_ = 'PR_MAKER_WASH_CAP';
/** メーカー洗い1回あたりの Keepa＋門件数（既定20）。Property で上書き。上限50。 */
var PR_MAKER_WASH_CAP_DEFAULT_ = 30;
var PR_MAKER_WASH_CAP_MAX_ = 50;
/** 洗い完走（案A+C）。未設定=ON。false で旧 page0+cap に戻す。 */
var PR_MAKER_WASH_FULL_PROP_ = 'PR_MAKER_WASH_FULL_ENABLED';
var PR_WASH_FULL_TRIGGER_FN_ = 'prRunWashFullFromTrigger';
var PR_WASH_FULL_TRIGGER_PROP_ = 'PR_WASH_FULL_TRIGGER_ENABLED';
var PR_WASH_FULL_TRIGGER_AFTER_PROP_ = 'PR_WASH_FULL_TRIGGER_AFTER_MIN';
var PR_WASH_FULL_TRIGGER_AFTER_DEFAULT_ = 5;
/** 洗いトリガー診断: 予約時の発火予定（ms）・arm メタ。未発火切り分け用。 */
var PR_WASH_TRIGGER_EXPECT_FIRE_MS_PROP_ = 'PR_WASH_TRIGGER_EXPECT_FIRE_MS';
var PR_WASH_TRIGGER_ARM_TAG_PROP_ = 'PR_WASH_TRIGGER_ARM_TAG';
var PR_WASH_TRIGGER_ARM_PARENT_PROP_ = 'PR_WASH_TRIGGER_ARM_PARENT';
var PR_WASH_TRIGGER_ARM_RUNID_PROP_ = 'PR_WASH_TRIGGER_ARM_RUNID';
var PR_WASH_TRIGGER_LAST_CLEAR_PROP_ = 'PR_WASH_TRIGGER_LAST_CLEAR';
/** C4 完走連鎖（§6.4）。未設定=ON。false で完走後 D1→D2→心拍なし。 */
var PR_CYCLE_C4_ENABLED_PROP_ = 'PR_CYCLE_C4_ENABLED';
/** 旧: C4 が次社洗いを直接 arm するときのメーカー名（成功連鎖ではセットしない。残骸消化用）。 */
var PR_CYCLE_C4_ARM_MAKER_PROP_ = 'PR_CYCLE_C4_ARM_MAKER';
/** 直近完走 parent（C4 手動・ログ用）。 */
var PR_CYCLE_C4_LAST_PARENT_PROP_ = 'PR_CYCLE_C4_LAST_PARENT';
/** 洗いトリガー実行中（expect より先に立てる。沈黙番犬用）。 */
var PR_WASH_INFLIGHT_RUNID_PROP_ = 'PR_WASH_INFLIGHT_RUNID';
var PR_WASH_INFLIGHT_MS_PROP_ = 'PR_WASH_INFLIGHT_MS';
var PR_WASH_INFLIGHT_PHASE_PROP_ = 'PR_WASH_INFLIGHT_PHASE';
/** inflight を期限切れとみなす（30分壁＋余裕）。 */
var PR_WASH_INFLIGHT_STALE_MS_ = 35 * 60 * 1000;
/** ckpt/overlap_wait で wait_shortened を無効（2026-09-01）。false で旧挙動。 */
var PR_WASH_CKPT_NO_SHORTEN_PROP_ = 'PR_WASH_CKPT_NO_SHORTEN';
/** overlap 再予約 afterMin を inflight stale 寄せ（2026-09-01）。false で8分固定。 */
var PR_WASH_OVERLAP_STALE_AWARE_PROP_ = 'PR_WASH_OVERLAP_STALE_AWARE';
/**
 * 「実残トークンが満タンなら待ちは2分まで」の整流を、効率ループとは別に持つ。
 * 2026-09-07: PR_WASH_EFF_LOOP_ENABLED=false（ckpt を戻す）にした副作用で、
 * この上限も一緒に消えて overlap 再予約が8〜35分待ちになり、満タンのまま空白が出た。
 * ckpt（同一実行の安全ネット）は対象外＝短縮しない。切戻し: PR_WASH_ARM_FULL_CAP=false。
 */
var PR_WASH_ARM_FULL_CAP_PROP_ = 'PR_WASH_ARM_FULL_CAP';
/**
 * 効率ループ整流（P-eff-1/2・2026-09-03）。未設定=ON。
 * ON: 実行中 ckpt 廃止・満タン afterMin≤2・WAIT_TOKEN 前 inflight clear・overlap 短TTL/即takeover。
 * 切戻し false → 旧 ckpt0（GAS残＋2）／overlap stale 寄せ。
 */
var PR_WASH_EFF_LOOP_ENABLED_PROP_ = 'PR_WASH_EFF_LOOP_ENABLED';
/** overlap 即 takeover の年齢下限（分）。 */
var PR_WASH_OVERLAP_TAKEOVER_MIN_ = 8;
/**
 * 店レーン譲り（2026-09-03）。未設定=ON。
 * 洗い→心拍/店へ譲るとき inflight を捨て短予約。店作業中は busy で二重発火を弾く。
 * 切戻し false → 旧（心拍中も洗い inflight 残し → overdue≈66分）。
 */
var PR_SHOP_LANE_CLEAR_INFLIGHT_PROP_ = 'PR_SHOP_LANE_CLEAR_INFLIGHT';
var PR_SHOP_HB_BUSY_MS_PROP_ = 'PR_SHOP_HB_BUSY_MS';
var PR_SHOP_HB_BUSY_RUN_PROP_ = 'PR_SHOP_HB_BUSY_RUN';
/** 店/心拍作業中とみなす上限（分）。超えたら busy 失効。 */
var PR_SHOP_HB_BUSY_TTL_MIN_ = 25;
/** 譲り時の安全ネット予約（分）。alight 完了で 1 分に上書き。 */
var PR_SHOP_YIELD_ARM_MIN_ = 3;
/**
 * 効率ループ合否ログ（①ログ「効率合否」）。未設定=ON。切戻し false。
 * 比較用: gap / afterMin / liveLeft / inflightAge / cause。
 */
var PR_EFF_JUDGE_ENABLED_PROP_ = 'PR_EFF_JUDGE_ENABLED';
var PR_EFF_LAST_SLICE_MS_PROP_ = 'PR_EFF_LAST_SLICE_MS';
var PR_EFF_LAST_SLICE_PARENT_PROP_ = 'PR_EFF_LAST_SLICE_PARENT';
/** 切片間隔: これ以下 PASS（分）。理想≈14。 */
var PR_EFF_GAP_PASS_MAX_MIN_ = 25;
/** 切片間隔: 超えたら FAIL（分）。空洞扱い。 */
var PR_EFF_GAP_FAIL_MIN_ = 90;
/** 予約: liveLeft がこれ以上なのに afterMin が長い → FAIL。 */
var PR_EFF_ARM_FULL_LEFT_ = 280;
var PR_EFF_ARM_LONG_AFTER_MIN_ = 10;
/**
 * ckpt 予約の正常上限に足す余裕（分）。
 * ckpt は `prCheckpointWashFullTrigger_` が「残り予算+2分」で貼る＝最大 予算+2（既定22分）。
 * 同一実行と重ならないための安全ネットなので、その範囲の待ちは異常ではない
 * （満タン2分上限も `prScheduleWashFullTrigger_` で ckpt を除外している）。
 * 切戻し: Property PR_EFF_CKPT_BUDGET_RULE=false（固定15分で WARN の旧判定）。
 */
var PR_EFF_CKPT_MARGIN_MIN_ = 3;
var PR_EFF_CKPT_BUDGET_RULE_PROP_ = 'PR_EFF_CKPT_BUDGET_RULE';
/** overlap: age がこれ以上（分）なら FAIL（即 takeover すべき）。 */
var PR_EFF_OVERLAP_FAIL_AGE_MIN_ = 8;
/** C4 完走連鎖を洗い完走と別実行にする時間トリガー（30分壁回避）。 */
var PR_CYCLE_C4_TRIGGER_FN_ = 'prRunCycleC4FromTrigger';
var PR_CYCLE_C4_TRIGGER_AFTER_DEFAULT_ = 1;
var PR_CYCLE_C4_PENDING_PARENT_PROP_ = 'PR_CYCLE_C4_PENDING_PARENT';
var PR_CYCLE_C4_PENDING_RUNID_PROP_ = 'PR_CYCLE_C4_PENDING_RUNID';
var PR_CYCLE_C4_PENDING_TAG_PROP_ = 'PR_CYCLE_C4_PENDING_TAG';
var PR_CYCLE_C4_TRIGGER_EXPECT_MS_PROP_ = 'PR_CYCLE_C4_TRIGGER_EXPECT_MS';
/** 心拍生存番犬（15分周期）。時計ゼロ＋inflight期限切れを再武装。 */
var PR_CYCLE_LIVENESS_FN_ = 'prRunCycleLivenessFromTrigger';
var PR_CYCLE_LIVENESS_EVERY_MIN_ = 15;
/** inflight だけで生きとみなす上限（洗い stale 35分より短い。発火死の早期検知）。 */
var PR_CYCLE_LIVENESS_SILENCE_MS_ = 20 * 60 * 1000;
/**
 * 次実行待ちの腐り判定: Keepa バケツ満タン回復（ceil(MAX/rate)+1）＋α 分。
 * 予定時刻がこれ以上過ぎても発火しない／実行中印だけ残る → リセットして再開。
 */
var PR_CYCLE_STALE_ALPHA_MIN_ = 5;
/**
 * 実行中印（inflight）だけの腐り判定は **30分壁基準の35分**にする（2026-09-08）。
 * トークン回復基準の66分を待つと、選定で沈黙死した実行を番犬が拾うまで70分空く
 * （実測 09/08 00:02 発火→pick_begin のまま沈黙→01:12 hard_reset＝179分の空白）。
 * GAS は約30分で殺すので、35分を超えた実行中印は生きている可能性がない。
 * 予定時刻超過（expect_overdue）は従来どおりトークン回復基準のまま。
 * 切戻し: Property PR_CYCLE_INFLIGHT_STALE_FAST=false。
 */
var PR_CYCLE_INFLIGHT_STALE_FAST_PROP_ = 'PR_CYCLE_INFLIGHT_STALE_FAST';
/**
 * 規模帯の確定（Catalog列挙）を1実行で何社まとめるか（2026-09-08）。
 * 列挙そのものは16〜48秒だが、1社ごとに 洗い→店→心拍 の受け渡しをやり直すため
 * 実測1社あたり7分かかっていた（内訳: 選定1.6分＋yield2分＋店busy2分＋往復1分）。
 * 未判定が実測244社あり、1社7分では消化に約28時間。その間 Keepa は材料切れで空く
 * （2026-09-08 夜間: 10.3時間のうち61%が Keepa作業ゼロ・空振り選定24回）。
 * 打ち切り: 予算残り不足／小中が1社出た時点。1 にすると旧動作（1社ずつ）。
 * 切戻し: Property PR_CYCLE_ENUM_BATCH_N=1。
 */
var PR_CYCLE_ENUM_BATCH_N_PROP_ = 'PR_CYCLE_ENUM_BATCH_N';
var PR_CYCLE_ENUM_BATCH_N_DEFAULT_ = 8;
var PR_CYCLE_ENUM_BATCH_N_MAX_ = 20;
/** まとめ列挙を打ち切る予算残り（予約とログのために残す）。 */
var PR_CYCLE_ENUM_BATCH_RESERVE_MS_ = 5 * 60 * 1000;
/** タイトル含有0がこの回数続いたら保留（title0_skip）して心拍へ。 */
var PR_WASH_TITLE0_SKIP_AFTER_ = 3;
var PR_WASH_TITLE0_FAIL_PROP_ = 'PR_WASH_TITLE0_FAIL_JSON';
/** title0_skip 済み parent（| 区切り）。再選定・再試行しない。 */
var PR_WASH_TITLE0_SKIP_SET_PROP_ = 'PR_WASH_TITLE0_SKIP_SET';
/** 洗い完走済みで D1 未の parent 列（| 区切り）。A-1／番犬の C4 recover 用。 */
var PR_CYCLE_C4_NEEDS_D1_PROP_ = 'PR_CYCLE_C4_NEEDS_D1';
/** D1 完走済み parent（| 区切り）。再 C4 しない。 */
var PR_CYCLE_C4_D1_DONE_PROP_ = 'PR_CYCLE_C4_D1_DONE';
/**
 * Keepa スライス途中でも門通過→品番リスト転記（CYCLE §6.4bis）。未設定=ON。
 * 完走待ちで大／特大の通過がリストに載らない漏れを防ぐ。C4 D1_DONE は付けない。
 */
var PR_CYCLE_MID_D1_ENABLED_PROP_ = 'PR_CYCLE_MID_D1_ENABLED';
/** 週次転記チェック（CYCLE §6.4ter）。未設定=ON。 */
var PR_CYCLE_LIST_CHECK_ENABLED_PROP_ = 'PR_CYCLE_LIST_CHECK_ENABLED';
var PR_CYCLE_LIST_CHECK_DAYS_PROP_ = 'PR_CYCLE_LIST_CHECK_DAYS';
var PR_CYCLE_LIST_CHECK_DAYS_DEFAULT_ = 7;
var PR_CYCLE_LIST_CHECK_FN_ = 'prRunListTransferWeeklyCheck';
/** 小・中 Keepa 完走後に店レーン1回を挟む（CYCLE §6.3bis 案B）。値=完走メーカー名。 */
var PR_CYCLE_SHOP_AFTER_SM_PROP_ = 'PR_CYCLE_SHOP_AFTER_SM';
/** 大・特大の週間上限（未設定=2）。 */
var PR_CYCLE_XL_WEEKLY_CAP_PROP_ = 'PR_CYCLE_XL_WEEKLY_CAP';
var PR_CYCLE_XL_WEEKLY_CAP_DEFAULT_ = 2;
var PR_CYCLE_XL_WEEK_KEY_PROP_ = 'PR_CYCLE_XL_WEEK_KEY';
var PR_CYCLE_XL_WEEK_COUNT_PROP_ = 'PR_CYCLE_XL_WEEK_COUNT';
/** 今週すでに枠を消費した大・特大メーカー（| 区切り）。 */
var PR_CYCLE_XL_WEEK_MAKERS_PROP_ = 'PR_CYCLE_XL_WEEK_MAKERS';
/** ブランド自動追加閾値（CYCLE §6.6）。 */
var PR_BRAND_HINT_RAW_MIN_DEFAULT_ = 30;
var PR_BRAND_HINT_RATIO_DEFAULT_ = 0.4;
var PR_BRAND_HINT_ABS_DEFAULT_ = 10;
var PR_BRAND_HINT_RAW_MIN_PROP_ = 'PR_BRAND_HINT_RAW_MIN';
var PR_BRAND_HINT_RATIO_PROP_ = 'PR_BRAND_HINT_RATIO';
var PR_BRAND_HINT_ABS_PROP_ = 'PR_BRAND_HINT_ABS';
/** 洗い完走の単一実行ロック（UTC ms 期限）。TTL 既定30分。 */
var PR_WASH_FULL_LOCK_PROP_ = 'PR_WASH_FULL_LOCK_UNTIL_MS';
var PR_WASH_FULL_LOCK_TTL_MS_ = 30 * 60 * 1000;
/** A-1 等で上げる。ロック取得時の世代より新しければ旧洗いを打ち切る。 */
var PR_WASH_ABORT_EPOCH_PROP_ = 'PR_WASH_ABORT_EPOCH';
/** 現在の洗いロックが取得された時点の abort 世代。 */
var PR_WASH_LOCK_EPOCH_PROP_ = 'PR_WASH_LOCK_EPOCH';
var PR_WASH_FULL_LOCK_OWNER_PROP_ = 'PR_WASH_FULL_LOCK_OWNER';
var PR_PROF_SELLER_MIN_HEADER_ = '出品者数下限';
var PR_PROF_SELLER_MAX_HEADER_ = '出品者数上限';
var PR_PROF_SELLER_MIN_DEFAULT_ = 1;
var PR_PROF_SELLER_MAX_DEFAULT_ = 20;
/** 順位番号の下限。1＝ほぼ制限なしだが欠測は落ち。 */
var PR_PROF_RANK_MIN_HEADER_ = '順位下限';
var PR_PROF_RANK_MIN_DEFAULT_ = 1;
/** 税込価格の上限。カテゴリ別変更可。初期は全カテゴリ 100000。 */
var PR_PROF_PRICE_MAX_HEADER_ = '価格上限';
var PR_PROF_PRICE_MAX_DEFAULT_ = 100000;
/** D2 重量上限（g）。欠測は見送りにしない。カテゴリ別可。 */
var PR_PROF_WEIGHT_MAX_HEADER_ = '重量上限g';
var PR_PROF_WEIGHT_MAX_DEFAULT_ = 5000;
/**
 * ①プロファイル列の左寄せ順。価格上下限・順位まわりを隣接。無い列はスキップ。
 * 未知の列は右側に残る。
 */
var PR_PROF_COL_ORDER_ = [
  '調査カテゴリー',
  'AmazonルートID',
  'カテゴリ',
  'マッチ語',
  '価格下限',
  '価格上限',
  '順位段階1',
  '順位段階2',
  '順位段階3',
  '門に使う順位段階',
  '順位下限',
  '本体ルール',
  '本体不在%',
  '冷凍生',
  '大型',
  '重量上限g',
  '出品者1人',
  '出品者数下限',
  '出品者数上限',
  '粗利',
  '有効',
  'メモ',
];
var PR_WASH_QUEUE_SHEET_ = '①洗いキュー';
var PR_WASH_SPLIT_SHEET_ = '①洗い分割語';
var PR_MAKER_MASTER_SHEET_ = 'メーカーマスタ';
/** Keepa済キュー行の保持日数。未設定=90（3か月）。Property PR_WASH_QUEUE_RETENTION_DAYS。 */
var PR_WASH_QUEUE_RETENTION_DAYS_DEFAULT_ = 90;
var PR_WASH_QUEUE_RETENTION_DAYS_PROP_ = 'PR_WASH_QUEUE_RETENTION_DAYS';
var PR_WASH_STATUS_HEADER_ = '洗い状態';
var PR_WASH_PROGRESS_HEADER_ = '洗い進捗';
/** 実態=要確認のときだけ Amazon 検索URL。列名「実態ck」（末尾追加。M列の「実態」は触らない）。 */
var PR_JITTAI_CK_HEADER_ = '実態ck';
var PR_CATALOG_SPLIT_THRESHOLD_ = 500;
var PR_BAND_S_MAX_ = 100;
var PR_BAND_M_MAX_ = 300;
var PR_CYCLE_BAND_SELECT_PROP_ = 'PR_CYCLE_BAND_SELECT';
var PR_CYCLE_HEARTBEAT_PAUSED_PROP_ = 'PR_CYCLE_HEARTBEAT_PAUSED';
var PR_CYCLE_HEARTBEAT_ENABLED_PROP_ = 'PR_CYCLE_HEARTBEAT_ENABLED';
var PR_CYCLE_RESUME_01_FN_ = 'prRunCycleResumeAt01';
var PR_CYCLE_DAILY_MAIL_FN_ = 'prRunCycleDailyMail';
var PR_CYCLE_DAILY_MAIL_ENABLED_PROP_ = 'PR_CYCLE_DAILY_MAIL_ENABLED';
var PR_CYCLE_DAILY_MAIL_TO_PROP_ = 'PR_CYCLE_DAILY_MAIL_TO';
var PR_CYCLE_AUTO_CK_CAP_PROP_ = 'PR_CYCLE_AUTO_CK_DAILY_CAP';
var PR_CYCLE_AUTO_CK_DAY_PROP_ = 'PR_CYCLE_AUTO_CK_DAY';
var PR_CYCLE_AUTO_CK_COUNT_PROP_ = 'PR_CYCLE_AUTO_CK_COUNT';
var PR_CYCLE_K_SMALL_MED_ = 3;
var PR_CYCLE_AUTO_CK_CAP_DEFAULT_ = 8;
var PR_SELLER_SMALL_HIT_COL_ = '小規模出現';
var PR_CATALOG_EST_COL_ = 'Catalog見積';
var PR_BAND_COL_ = '規模帯';
var PR_BAND_S_ = '小';
var PR_BAND_M_ = '中';
var PR_BAND_L_ = '大';
var PR_BAND_XL_ = '特大';
var PR_BAND_UNK_ = '不明';
var PR_CAT_INFER_SAMPLE_MAX_ = 200;
var PR_CAT_INFER_WORD_MAX_ = 15;
var PR_WASH_QUEUE_HEADERS_ = ['メーカー', '入口語', 'ASIN', '列挙日時', 'Keepa済', '門結果', 'メモ'];
var PR_WASH_SPLIT_HEADERS_ = [
  'メーカー',
  '分割語',
  '状態',
  '推定元',
  'サンプル数',
  '列挙ASIN数',
  'nextToken',
  '更新',
];
var PR_WASH_ST_NONE_ = '未着手';
var PR_WASH_ST_ENUM_ = '列挙中';
var PR_WASH_ST_INFER_ = '分割推定中';
var PR_WASH_ST_KEEPA_ = 'Keepa中';
var PR_WASH_ST_NEED_Q2_ = '要分割語';
var PR_WASH_ST_DONE_ = '完了';
var PR_WASH_ST_HOLD_ = '保留';
var PR_SPLIT_ST_PENDING_ = '未';
var PR_SPLIT_ST_ENUM_ = '列挙中';
var PR_SPLIT_ST_DONE_ = '済';
var PR_GEMINI_MODEL_DEFAULT_ = 'gemini-3.6-flash';
var PR_OPENAI_MODEL_DEFAULT_ = 'gpt-4o-mini';
/** Keepa refillRate フォールバック（実測5）。応答優先。Property KEEPA_REFILL_RATE。 */
var PR_KEEPA_REFILL_RATE_DEFAULT_ = 5;
/** Keepa トークンバケツ上限（画面・契約。Property KEEPA_TOKEN_BUCKET_MAX）。 */
var PR_KEEPA_TOKEN_BUCKET_MAX_DEFAULT_ = 300;
var PR_KEEPA_TOKEN_BUCKET_MAX_PROP_ = 'KEEPA_TOKEN_BUCKET_MAX';
/**
 * Keepa 単価はメニュー別ではなく「取得内容（プロファイル）」別。
 * product_offers20 = history=0&stats=90&offers=20（店門・洗い・D1欠落）。実測≈5.5–6。
 * Property KEEPA_PRODUCT_TOKEN_COST で上書き（従来キー流用）。
 */
var PR_KEEPA_COST_PRODUCT_OFFERS20_DEFAULT_ = 6;
var PR_KEEPA_PRODUCT_TOKEN_DEFAULT_ = PR_KEEPA_COST_PRODUCT_OFFERS20_DEFAULT_;
var PR_KEEPA_PRODUCT_TOKEN_PROP_ = 'KEEPA_PRODUCT_TOKEN_COST';
var PR_KEEPA_COST_PROFILE_OFFERS20_ = 'product_offers20';
/**
 * 2段ゲート（2026-09-07）。門は stats（`current[11]`＝新品出品者数）だけで判定でき、
 * offers を要らない。1段目 offers なし＝1トークン/件で落とし、通過だけ offers=20 で取り直す。
 * 実測の通過率は3%（同日 120件中4件）なので 1件あたり約5 → 約1.2 になる。
 * 切戻し: Property PR_KEEPA_TWO_STAGE_GATE=false（従来の一発 offers=20）。
 */
var PR_KEEPA_TWO_STAGE_GATE_PROP_ = 'PR_KEEPA_TWO_STAGE_GATE';
var PR_KEEPA_COST_PROFILE_STATS90_ = 'product_stats90';
var PR_KEEPA_COST_PRODUCT_STATS90_DEFAULT_ = 1;
var PR_KEEPA_STATS90_TOKEN_PROP_ = 'KEEPA_STATS90_TOKEN_COST';
var PR_KEEPA_COST_PROFILE_TWO_STAGE_ = 'product_two_stage';
/** 1段目1＋通過3%×5 ≒ 1.2。afford 見積は安全側に2を既定にする。 */
var PR_KEEPA_COST_PRODUCT_TWO_STAGE_DEFAULT_ = 2;
var PR_KEEPA_TWO_STAGE_TOKEN_PROP_ = 'KEEPA_TWO_STAGE_TOKEN_COST';
/** 門通過なのに Keepaフルの BuyBox が空の行を1回だけ offers=20 で太らせる。未設定=ON。 */
var PR_D1_ENRICH_BLANK_BB_PROP_ = 'PR_D1_ENRICH_BLANK_BB';
var PR_D1_ENRICH_CAP_PROP_ = 'PR_D1_ENRICH_CAP';
var PR_D1_ENRICH_CAP_DEFAULT_ = 20;
var PR_KEEPA_FULL_WRITE_PROP_ = 'PR_KEEPA_FULL_WRITE_ENABLED';
var PR_PURPOSE_RESEARCH_ = 'リサーチ';
var PR_JITTAI_MAKER_ = '製造メーカー';
var PR_JITTAI_BRAND_ = '入口ブランド';
var PR_JITTAI_CHECK_ = '要確認';
/** AI調査前のサマリ初期ラベル（洗わない。実態ck URLも付けない）。 */
var PR_JITTAI_UNCHECKED_ = '未調査';
var PR_JITTAI_EXCLUDE_ = '非仕入先';
/**
 * 実態シード（最小）。非仕入先のみ。他は AI または要確認。
 * tools/competitor_store/maker_alias.py MAKER_JITTAI の非仕入先行と同期。
 */
var PR_MAKER_JITTAI_EXCLUDE_SEED_ = {
  わが街とくさんネット: '非仕入先',
  扶桑社: '非仕入先',
  天然生活: '非仕入先',
  ノーブランド品: '非仕入先',
  Generic: '非仕入先',
  未確定: '非仕入先',
  リフココ: '非仕入先',
  リフ工房: '非仕入先',
};
/** メーカーマスタ旧データクリア対象（実態AI再調査用）。 */
var PR_MAKER_MASTER_STALE_RESET_ = {
  五木食品: true,
  石原水産: true,
  永谷園: true,
  サトウ食品: true,
};
/** メーカーマスタ B–F＋メモ（第2語キャッシュ。洗い完走で使用。列は削除しない）。 */
var PR_MAKER_MASTER_Q2_COLS_ = ['第2クエリ語', '採取元', '公式URL', '取得日時', '再採取しない', 'メモ'];
/** @deprecated PR_MAKER_JITTAI_EXCLUDE_SEED_ を使用 */
var PR_MAKER_JITTAI_SEED_ = PR_MAKER_JITTAI_EXCLUDE_SEED_;
/** 入口語 → 親メーカー（maker_alias.py MAKER_ALIASES 同期）。 */
var PR_MAKER_ALIASES_ = {
  AGF: '味の素',
  エージーエフ: '味の素',
  'AGF(エージーエフ)': '味の素',
  agf: '味の素',
  いしもと: 'いし本食品工業',
  いし本食品工業: 'いし本食品工業',
  'S&B': 'エスビー食品',
  エスビー食品: 'エスビー食品',
  ヱスビー: 'エスビー食品',
  Nagatanien: '永谷園',
  Sato: 'サトウ食品',
};
/** セラーヒット集計から除外（CYCLE §5）。 */
var PR_HIT_EXCLUDE_SET_ = { Generic: true, generic: true, ノーブランド品: true };
/** 入口ブランド型の明示 needle（maker_alias.py と同期。Catalog entry は先頭語）。 */
var PR_BRAND_NEEDLES_ = {
  'サトウ食品': ['Sato', 'サトウの'],
  '味の素': ['AGF', 'エージーエフ'],
  'エスビー食品': ['S&B', 'エスビー'],
  'いし本食品工業': ['いしもと'],
};
/** 実態リゾルバ Property（正本: B_MAKER_JITTA_CLASSIFICATION.md）。 */
var PR_JITTA_RESOLVE_ENABLED_PROP_ = 'PR_JITTA_RESOLVE_ENABLED';
var PR_JITTA_RESOLVE_CAP_PROP_ = 'PR_JITTA_RESOLVE_CAP';
var PR_JITTA_RESOLVE_CAP_DEFAULT_ = 100;
var PR_JITTA_RESOLVE_CAP_MAX_ = 100;
/** 心拍で実態AIを店と分離（省略=ON）。切戻し false で旧: A-light 末尾のみ。 */
var PR_JITTA_HB_LANE_ENABLED_PROP_ = 'PR_JITTA_HB_LANE_ENABLED';
/** 心拍1回あたりの実態AI件数（省略=8・最大20）。 */
var PR_JITTA_HB_CAP_PROP_ = 'PR_JITTA_HB_CAP';
var PR_JITTA_HB_CAP_DEFAULT_ = 8;
var PR_JITTA_HB_CAP_MAX_ = 20;
/** 未調査がこの件数以上なら自動レ点を止める（省略=30）。 */
var PR_JITTA_BACKLOG_GATE_PROP_ = 'PR_JITTA_BACKLOG_GATE';
var PR_JITTA_BACKLOG_GATE_DEFAULT_ = 30;
/** A-light 末尾でも実態AIする（省略=レーンONならOFF／レーンOFFならON）。 */
var PR_JITTA_AFTER_ALIGHT_PROP_ = 'PR_JITTA_AFTER_ALIGHT';
/**
 * 実態AIの実行主体。hb=①心拍（既定）／comp=競合DBスクリプト。
 * comp のとき①は jitta_resolve しない（二重防止）。切戻しは hb または削除。
 */
var PR_JITTA_OWNER_PROP_ = 'PR_JITTA_OWNER';
var PR_JITTA_OWNER_HB_ = 'hb';
var PR_JITTA_OWNER_COMP_ = 'comp';
var PR_JITTA_GYOSHU_UNKNOWN_ = '不明';
var PR_JITTA_GYOSHU_INFERRED_SUFFIX_ = '（推定）';
/** メーカーマスタ実態キャッシュ列（schema.py MAKER_HEADERS 同期）。 */
var PR_MAKER_JITTA_MASTER_COLS_ = [
  '実態候補',
  '実態信頼度',
  '実態採取元',
  '実態根拠URL',
  '実態取得日時',
  '実態再採取しない',
  '業種',
  'Amazonルートカテゴリ',
  'AmazonBrowse棚',
  '入口語候補',
  '判定メモ',
];
/** メーカーサマリミラー列。 */
var PR_MAKER_JITTA_MIRROR_COLS_ = [
  '実態採取元',
  '実態信頼度',
  '実態根拠URL',
  '業種',
  'Amazonルートカテゴリ',
  'AmazonBrowse棚',
];
/** Amazon ルート（セラー台帳列名。schema.py AMAZON_SELLER_CAT_COLS 同期）。 */
var PR_AMAZON_ROOT_CATS_ = [
  '洋書',
  '本',
  'ミュージック',
  'ファッション',
  'DVD',
  'クラシック',
  'ホーム&キッチン',
  'PCソフト',
  'TVゲーム',
  '文房具・オフィス用品',
  '家電&カメラ',
  'ドラッグストア',
  'ビューティー',
  'DIY・工具・ガーデン',
  'おもちゃ',
  '産業・研究開発用品',
  'スポーツ&アウトドア',
  'パソコン・周辺機器',
  '車＆バイク',
  'ホビー',
  'ベビー&マタニティ',
  'ペット用品',
  '楽器',
  '大型家電',
  '食品・飲料・お酒',
  'デジタルミュージック',
  'Kindleストア',
  '不明',
];
/**
 * AI/Web 調査 fixture（tools/competitor_store/maker_jitta_resolve.py AI_FIXTURES 同期）。
 * API 失敗時のフォールバック。
 */
var PR_JITTA_AI_FIXTURES_ = {
  永谷園: {
    jitta_candidate: '製造メーカー',
    entry_preferred: 'メーカー',
    entry_needle: '',
    confidence: 'high',
    source: 'official',
    evidence_urls: [
      'https://www.nagatanien.co.jp/company/profile.html',
      'https://www.nagatanien-hd.co.jp/',
    ],
    gyoshu: '即席食品・中食（食品製造）',
    gyoshu_source: 'official',
    amazon_root_categories: '食品・飲料・お酒',
    amazon_browse_shelves: 'お茶漬け・ふりかけ|即席みそ汁|惣菜・食材・調味料',
    reason: '会社概要に製造販売と明記。棚名も永谷園。',
  },
  味の素: {
    jitta_candidate: '入口ブランド',
    entry_preferred: 'ブランド',
    entry_needle: 'AGF',
    confidence: 'high',
    source: 'official',
    evidence_urls: ['https://www.ajinomoto.co.jp/'],
    gyoshu: '調味料・加工食品',
    gyoshu_source: 'official',
    amazon_root_categories: '食品・飲料・お酒',
    amazon_browse_shelves: '調味料|コーヒー・紅茶|冷凍食品|スープ・みそ汁',
    reason: '親は味の素。棚では AGF・ほんだし等の別ブランド名。',
  },
  あさくま: {
    jitta_candidate: '要確認',
    entry_preferred: '要確認',
    entry_needle: '',
    confidence: 'medium',
    source: 'official',
    evidence_urls: ['https://www.asakuma.co.jp/company.html'],
    gyoshu: '外食・飲食',
    gyoshu_source: 'official',
    amazon_root_categories: '食品・飲料・お酒',
    amazon_browse_shelves: '調味料・ソース',
    reason: '本業はレストラン。ソース販売はあるが卸可否は要確認。',
  },
};
var PR_CAND_SELLER_SPAPI_DONE_ = 'セラー調査済みSP-API';
var PR_CAND_SELLER_SPAPI_AT_ = 'セラー調査日時';
/** ①候補: 門通過行の品番リスト転記状態（CYCLE §6.4ter）。 */
var PR_CAND_LIST_CHECK_ = '転記チェック';
var PR_CAND_LIST_CHECK_MI_ = '未';
var PR_CAND_LIST_CHECK_SUMI_ = '済';
var PR_CAND_LIST_CHECK_MORE_ = '漏れ';
var PR_CAND_HEADERS_ = [
  'メーカー',
  'ASIN',
  '商品名',
  '税込価格',
  '順位90',
  'JAN',
  '発見経路',
  'sellerId',
  '門結果',
  '門理由',
  'runId',
  '取得日時',
  '店名',
  PR_CAND_SELLER_SPAPI_DONE_,
  PR_CAND_SELLER_SPAPI_AT_,
  PR_CAND_LIST_CHECK_,
];

function onOpen() {
  // シンプルトリガーは常に30秒。シートI/O・Keepa・openByIdは禁止。
  // メニュー正: 自動リサーチ → A全自動 / B個別 / C見積枠 / 99運用（HUMAN_RUN §2.1）
  try {
    var ui = SpreadsheetApp.getUi();
    var mA = ui
      .createMenu('A. 全自動')
      .addItem(
        'A-1. 開始／再開（セラー・メーカー・品番リサーチ全自動）',
        'prMenuCycleResumeNow'
      )
      .addItem('A-2. 全診断（時計・沈黙・再武装）', 'prMenuCycleFullDiagnose');
    var mB = ui
      .createMenu('B. 個別調査')
      .addItem('1. レ点の店だけ調べる', 'prMenuAlightHit')
      .addItem('2. レ点のメーカーだけ洗う', 'prMenuCkMakers')
      .addItem('3. 実態だけAI判定（未調査）', 'prMenuResolveMakerJitta');
    var mC = ui
      .createMenu('C. 見積・発注')
      .addItem('（準備中・見積提案〜発注）', 'prMenuEstimateOrderPlaceholder');
    var mTest = ui
      .createMenu('テスト')
      .addItem('T1. 試験 A→D（店1＋メーカー1）', 'prMenuADTrialV1b')
      .addItem('T2. モリタ /query だけ', 'prMenuQueryMorita')
      .addItem('T3. モリタ 取扱→門', 'prMenuMoritaGate')
      .addItem('T4. 尾西を cap20 洗い', 'prMenuMakerWashOnishi')
      .addItem('T5. AGF→味の素を cap20 洗い', 'prMenuMakerWashAgf')
      .addItem('T6. 尾西 列挙だけ（キュー）', 'prMenuWashFullEnumDryOnishi')
      .addItem('T7. 旧: レ点メーカー→リスト', 'prMenuMakerResearchAndListUp')
      .addItem('T8. 旧: 未調査メーカー自動→リスト', 'prMenuAutoMakerResearchAndListUp');
    var m99 = ui
      .createMenu('99. 運用・切り分け')
      .addItem('1. 自動を止める（洗いトリガー停止）', 'prMenuWashFullTriggerStop')
      .addItem('2. 洗いトリガーだけ診断', 'prMenuWashTriggerDiagnose')
      .addItem('3. 心拍だけ診断', 'prMenuCycleLivenessDiagnose')
      .addItem('4. 完走連鎖だけ実行（D1→D2→心拍）', 'prMenuCycleC4AfterComplete')
      .addItem('5. D1だけ（門→品番リスト）', 'prMenuListAppendD1')
      .addItem('6. D2だけ（見積候補に絞る）', 'prMenuListFilterD2')
      .addItem('7. D1→D2 連続', 'prMenuListD1ThenD2')
      .addItem('8. 単品名だけAI充填', 'prMenuSingleNameFill')
      .addItem('9. セラー台帳を足す（C0）', 'prMenuSellerAccum')
      .addItem('10. セラーヒットを再計算', 'prMenuSellerHitBackfill')
      .addItem('11. 洗いを手動で続ける', 'prMenuWashFullContinue')
      .addItem('12. 洗いをループで続ける', 'prMenuWashFullLoopContinue')
      .addItem('13. 洗いトリガーを手動開始', 'prMenuWashFullTriggerStart')
      .addItem('14. 洗いキューの期限切れ削除', 'prMenuWashQueuePurgeExpired')
      .addItem('15. 実態ck（要確認に検索URL）', 'prMenuSyncMakerJittaCkUrls')
      .addItem('16. 実態ラベル移行（要確認→未調査）', 'prMenuMigrateJittaUnchecked_')
      .addItem('17. Keepaフル索引を直す', 'prMenuRebuildKeepaFullIndex')
      .addItem('18. Keepaフルから再門', 'prMenuRegateFromKeepaFull')
      .addItem('19. ①候補の重複ASINまとめ', 'prMenuCompactCand')
      .addItem('20. 偽完了を列挙中に戻す（候補0の完了）', 'prMenuFalseCompleteRewash')
      .addItem('21. マスタ旧データを消して再調査用に', 'prMenuResetStaleMakerMaster')
      .addItem('22. 日次生存メールを今すぐ送る', 'prMenuCycleDailyMailNow')
      .addItem('23. 約1分後にサイクル再開', 'prMenuCycleResume')
      .addItem('24. 自動選定→洗い1スライス（C3）', 'prMenuCycleC3MakerWash')
      .addItem('25. 自動選定→洗い→完走トリガー（C3+T）', 'prMenuCycleC3MakerWashWithTrigger')
      .addItem('26. 店1件通し（レ点1→門→①候補）', 'prMenuSellerOneThrough')
      .addItem('27. レ点の店を調査（旧・複数）', 'prMenuCkSellers')
      .addItem('28. 転記チェック（週次・今すぐ）', 'prMenuListTransferWeeklyCheckNow')
      .addItem('29. D2 backlog解消→C4再開', 'prMenuCycleC4ResumeAfterD2Clear')
      .addSeparator()
      .addSubMenu(mTest);
    ui.createMenu('自動リサーチ')
      .addSubMenu(mA)
      .addSubMenu(mB)
      .addSubMenu(mC)
      .addSubMenu(m99)
      .addToUi();
  } catch (e) {
    Logger.log('[仕入れ検討①] onOpen %s', String((e && e.message) || e).slice(0, 160));
  }
}

function prMenuQueryMorita() {
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=query state=RUNNING seller=%s', runId, PR_SELLER_MORITA_);
  try {
    var result = prKeepaQueryGet_(PR_SELLER_MORITA_, PR_FOOD_ROOT_, 0, 100);
    prAppendLog_([
      runId,
      'モリタ /query GET',
      'DONE',
      new Date(),
      result.totalResults,
      result.asinCount,
      result.tokensConsumed,
      (result.asinList[0] || ''),
    ]);
    Logger.log(
      '[仕入れ検討①] runId=%s state=DONE total=%s n=%s consumed=%s first=%s',
      runId,
      result.totalResults,
      result.asinCount,
      result.tokensConsumed,
      result.asinList[0] || ''
    );
    prMenuToast_(
      '① /query',
      '合計 ' + result.totalResults + ' / asinList ' + result.asinCount +
        ' consumed ' + result.tokensConsumed
    );
  } catch (e) {
    var msg = String((e && e.message) || e);
    msg = prRedactKey_(msg);
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'モリタ /query GET', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

function prKeepaKey_() {
  var key = String(PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY') || '').trim();
  if (!key) {
    throw new Error('Script Property KEEPA_API_KEY が未設定（出品と同じキーを①プロジェクトへ貼る）');
  }
  return key;
}

/**
 * Keepa GET /token（トークン消費0）。予約待ちが妥当か切り分けるための残量スナップショット。
 * @return {{ok:boolean, tokensLeft:*, refillRate:*, refillIn:*, tokensConsumed:*, http:number, error:string}}
 */
function prKeepaTokenStatus_() {
  try {
    var url = PR_KEEPA_TOKEN_ + '?key=' + encodeURIComponent(prKeepaKey_());
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { Accept: 'application/json' },
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText() || '';
    var data = {};
    try {
      data = JSON.parse(body);
    } catch (eParse) {
      return {
        ok: false,
        tokensLeft: '',
        refillRate: '',
        refillIn: '',
        tokensConsumed: '',
        http: code,
        error: 'json ' + String((eParse && eParse.message) || eParse).slice(0, 80),
      };
    }
    var out = {
      ok: code === 200,
      tokensLeft: data.tokensLeft != null ? data.tokensLeft : '',
      refillRate: data.refillRate != null ? data.refillRate : '',
      refillIn: data.refillIn != null ? data.refillIn : '',
      tokensConsumed: data.tokensConsumed != null ? data.tokensConsumed : '',
      http: code,
      error: code === 200 ? '' : 'HTTP ' + code,
    };
    return out;
  } catch (e) {
    return {
      ok: false,
      tokensLeft: '',
      refillRate: '',
      refillIn: '',
      tokensConsumed: '',
      http: 0,
      error: prRedactKey_(String((e && e.message) || e)).slice(0, 120),
    };
  }
}

/**
 * GET /query?domain=5&selection= JSON。POST ボディは使わない。
 * @param {string} sellerId
 * @param {string|string[]} rootCategoryOrIds Amazon ルートID（1つまたは OR 配列）
 * @param {number=} page
 * @param {number=} perPage
 */
function prKeepaQueryGet_(sellerId, rootCategoryOrIds, page, perPage) {
  var roots = prNormalizeRootCategoryIds_(rootCategoryOrIds);
  if (!roots.length) roots = [PR_FOOD_ROOT_];
  var selection = {
    rootCategory: roots,
    sellerIds: [String(sellerId)],
    productType: ['0'],
    sort: [['current_SALES', 'asc'], ['monthlySold', 'desc']],
    page: Number(page) || 0,
    perPage: Number(perPage) || 100,
  };
  var selStr = JSON.stringify(selection);
  var url =
    PR_KEEPA_QUERY_ +
    '?key=' +
    encodeURIComponent(prKeepaKey_()) +
    '&domain=5&selection=' +
    encodeURIComponent(selStr);
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { Accept: 'application/json' },
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '';
  if (code !== 200) {
    throw new Error('Keepa /query HTTP ' + code + ' ' + prRedactKey_(body).slice(0, 180));
  }
  var data = JSON.parse(body);
  var asins = data.asinList || [];
  if (!Array.isArray(asins)) asins = [];
  Logger.log(
    '[仕入れ検討①] query GET ok roots=%s total=%s n=%s consumed=%s',
    roots.join('|'),
    data.totalResults,
    asins.length,
    data.tokensConsumed
  );
  return {
    totalResults: data.totalResults,
    asinCount: asins.length,
    asinList: asins.map(function (a) {
      return String(a);
    }),
    tokensConsumed: data.tokensConsumed,
    rootCategories: roots,
  };
}

function prNormalizeRootCategoryIds_(rootCategoryOrIds) {
  var out = [];
  var seen = {};
  var raw = rootCategoryOrIds;
  if (raw == null || raw === '') return out;
  var arr = Object.prototype.toString.call(raw) === '[object Array]' ? raw : [raw];
  var i;
  for (i = 0; i < arr.length; i++) {
    var id = String(arr[i] || '')
      .trim()
      .replace(/\.0$/, '');
    if (!id || seen[id]) continue;
    seen[id] = true;
    out.push(id);
  }
  return out;
}

function prRedactKey_(s) {
  return String(s || '').replace(/key=[^&"'\s]+/gi, 'key=REDACTED');
}

function prSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function prEnsureSheet_(title, headers) {
  var ss = prSpreadsheet_();
  var sh = ss.getSheetByName(title);
  if (!sh) {
    sh = ss.insertSheet(title);
  }
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function prAppendLog_(row) {
  var headers = ['runId', '内容', 'state', 'at', 'totalResults', 'asinCount', 'tokensConsumed', 'detail'];
  var sh = prEnsureSheet_(PR_LOG_SHEET_, headers);
  sh.appendRow(row);
}

/**
 * 成功・情報用。toast は非ブロッキング。alert は OK までスクリプトが終わらず、
 * 競合DBなど別ブックへの書込が画面に出ないことがある。
 */
function prMenuToast_(title, body, sec) {
  var msg = String(body || '')
    .replace(/\r\n/g, '\n')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (msg.length > 250) msg = msg.slice(0, 247) + '...';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(msg, String(title || '仕入れ検討①').slice(0, 80), sec || 12);
  } catch (e1) {
    Logger.log('[仕入れ検討①] toast %s', String((e1 && e1.message) || e1).slice(0, 120));
  }
}

function prMenuFailAlert_(msg) {
  SpreadsheetApp.getUi().alert('失敗: ' + String(msg || '').slice(0, 300));
}

function prGasStart_() {
  PR_GAS_START_MS_ = new Date().getTime();
  PR_KEEPA_FULL_MAP_CACHE_ = null;
  PR_KEEPA_FULL_INDEX_CACHE_ = null;
  PR_PROF_CACHE_ = null;
  PR_CK_READY_ = {};
  PR_CAND_MEM_ = null;
  PR_CYCLE_POOL_CACHE_ = null;
  PR_WASH_QUEUE_COUNT_CACHE_ = null;
  PR_WASH_QUEUE_SNAP_ = {};
  PR_MAKER_BAND_CACHE_ = null;
  prClearSellerLedgerCache_();
}

function prGasTimeUp_() {
  if (!PR_GAS_START_MS_) return false;
  return new Date().getTime() - PR_GAS_START_MS_ >= PR_GAS_BUDGET_MS_;
}

function prMenuMoritaGate() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=gate state=RUNNING seller=%s', runId, PR_SELLER_MORITA_);
  try {
    var r = prRunSellerQueryGate_(PR_SELLER_MORITA_, runId);
    prAppendLog_([
      runId,
      'モリタ 取扱→門',
      r.status === 'fail' ? 'FAILED' : 'DONE',
      new Date(),
      r.total,
      r.asinCount,
      '',
      r.detail,
    ]);
    Logger.log('[仕入れ検討①] runId=%s state=%s %s', runId, r.status, r.detail);
    prMenuToast_('① 門', r.detail + ' query n=' + r.asinCount);
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'モリタ 取扱→門', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/** 店の取扱見直し間隔（日）。商品Keepa 90日とは別。未確定のため仮=180。レ点は無視して実行。 */
function prSellerRewalkDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_SELLER_REWALK_DAYS_PROP_) || 180);
  if (!(n > 0)) n = 180;
  return Math.floor(n);
}

/** 再調査時は機械欄を今回の値で上書き。切戻しは false（空のときだけ書く）。 */
function prSellerRewalkOverwriteEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_SELLER_REWALK_OVERWRITE_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prCkBatchCap_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_CK_BATCH_CAP_PROP_) || 3);
  if (!(n > 0)) n = 3;
  if (n > 10) n = 10;
  return Math.floor(n);
}

function prAutoMakerEnabled_() {
  return prCkEnabled_(PR_AUTO_MAKER_ENABLED_PROP_);
}

/** 未調査メーカー自動メニュー：1回あたり件数（既定1）。 */
function prAutoMakerCap_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_AUTO_MAKER_CAP_PROP_) || 1);
  if (!(n > 0)) n = 1;
  if (n > 5) n = 5;
  return Math.floor(n);
}

/** メーカー再調査間隔（日）。最終調査日が空またはこの日数より古い行を未調査扱い。既定90。 */
function prAutoMakerRewalkDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_AUTO_MAKER_REWALK_DAYS_PROP_) || 90);
  if (!(n > 0)) n = 90;
  return Math.floor(n);
}

function prMakerJittaAutoEligible_(jitta) {
  var j = String(jitta || '').trim();
  return j !== PR_JITTAI_EXCLUDE_ && j !== PR_JITTAI_CHECK_ && j !== PR_JITTAI_UNCHECKED_;
}

/** 洗い・C3選定で止める実態（未調査＝AI前、要確認＝AI後の人判定待ち）。 */
function prMakerJittaWashBlocked_(jitta) {
  var j = String(jitta || '').trim();
  return j === PR_JITTAI_CHECK_ || j === PR_JITTAI_UNCHECKED_;
}

/** サマリ.実態が AI 調査前のプレースホルダか。 */
function prMakerJittaIsUninvestigatedLabel_(jitta) {
  var j = String(jitta || '').trim();
  return !j || j === PR_JITTAI_UNCHECKED_;
}

/** サマリ行が未調査（実態信頼度が空で、AI前ラベルまたはレガシー要確認）。 */
function prMakerSummaryRowUninvestigated_(rowVals, hdr) {
  if (!rowVals || !hdr) return false;
  var idxConf = prHeaderIndex_(hdr, '実態信頼度');
  if (idxConf >= 0 && String(rowVals[idxConf] || '').trim()) return false;
  var idxJ = prHeaderIndex_(hdr, '実態');
  var j = idxJ >= 0 ? String(rowVals[idxJ] || '').trim() : '';
  if (prMakerJittaIsUninvestigatedLabel_(j)) return true;
  return j === PR_JITTAI_CHECK_;
}

function prCellToMsTokyo_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v.getTime();
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime();
  }
  var t = Date.parse(s);
  return isNaN(t) ? null : t;
}

/**
 * メーカーサマリから未調査メーカーを cap 件選ぶ（C3 自動選定 §6.3）。
 */
function prPickAutoMakerCandidates_(vals, hdr, cap) {
  return prPickCycleMakerSelectRows_(vals, hdr, cap);
}

function prCkEnabled_(propName) {
  var raw = PropertiesService.getScriptProperties().getProperty(propName);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prCkIsTrue_(v) {
  if (v === true || v === 1) return true;
  if (v === false || v === 0) return false;
  var s = String(v == null ? '' : v).trim().toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'YES' || s === 'ON';
}

function prCkEmptyAlert_(title, lastRow, trueN) {
  return (
    '調査CK（レ点）がオンの店がありません。\n' +
    '競合DB「' +
    title +
    '」A列の調査CKにチェックを入れてから再実行してください。\n' +
    '読込=' +
    Math.max(0, lastRow - 1) +
    '行 TRUE=' +
    trueN +
    '\n（完了するとレ点は外れます。入れ直すと行が青になります）'
  );
}

/** 同じ sellerId の別行に店名があれば使う（A①空行と storefront 行の重複）。 */
function prSellerNameFromSameId_(vals, idxId, idxName, sid) {
  var key = String(sid || '').trim().toUpperCase();
  if (!key || idxName < 0) return '';
  var r;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxId] || '').trim().toUpperCase() !== key) continue;
    var n = String(vals[r][idxName] || '').trim();
    if (n) return n;
  }
  return '';
}

function prYmdTokyo_(d) {
  return Utilities.formatDate(d || new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

/** 取得日時・キュー時刻など（日本時間・オフセット付き）。旧 UTC(Z) から切替。 */
function prNowFetchedAt_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function prHeaderIndex_(hdr, name) {
  var i;
  for (i = 0; i < hdr.length; i++) {
    if (String(hdr[i] || '').trim() === name) return i;
  }
  return -1;
}

function prAppendMemo_(old, extra) {
  var s = String(old || '').trim();
  var add = String(extra || '').trim();
  if (!add) return s;
  if (s.indexOf(add) >= 0) return s;
  return s ? s + ' | ' + add : add;
}

/** 入口語ヒットの page0=… メタデータを除去（Catalog entry/needle 汚染防止）。 */
function prCleanEntryToken_(s) {
  var t = String(s || '').trim();
  if (!t) return '';
  return t.replace(/\s*page0=[^\s|/,、／]+/gi, '').trim();
}

function prSplitNeedles_(parent, entryHit) {
  var out = [];
  var seen = {};
  function push(x) {
    var t = prCleanEntryToken_(x);
    if (!t || seen[t]) return;
    seen[t] = true;
    out.push(t);
  }
  push(parent);
  String(entryHit || '')
    .split(/[|,、／/]/)
    .forEach(push);
  return out;
}

/** ①候補 parent 列用 needle。入口語ヒット＋ブランド型の明示語を統合。 */
function prMakerWashNeedles_(parent, entryHit) {
  var needles = prSplitNeedles_(parent, entryHit);
  var extra = PR_BRAND_NEEDLES_[String(parent || '').trim()];
  if (!extra) return needles;
  var seen = {};
  var i;
  for (i = 0; i < needles.length; i++) seen[needles[i]] = true;
  for (i = 0; i < extra.length; i++) {
    var t = prCleanEntryToken_(extra[i]);
    if (t && !seen[t]) {
      seen[t] = true;
      needles.push(t);
    }
  }
  return needles;
}

/** Catalog keywords。ブランド型は入口語（page0 メタデータ除去後の先頭 brand 語）。 */
function prMakerCatalogEntry_(parent, pref, entryHit) {
  var p = String(parent || '').trim();
  var pr = String(pref || '').trim();
  if (pr.indexOf('ブランド') < 0) return p;
  var needles = prSplitNeedles_(parent, entryHit);
  var i;
  for (i = 1; i < needles.length; i++) {
    if (needles[i] && needles[i] !== p) return needles[i];
  }
  var extra = PR_BRAND_NEEDLES_[p];
  if (extra && extra.length) return prCleanEntryToken_(extra[0]);
  return p;
}

/**
 * メーカーサマリにメーカーが無ければ2行目へ1行挿入（同名は足さない）。調査CKの枠を付ける。
 * 挿入／既存とも実態が空ならシードで埋める（Generic→非仕入先など）。
 * @return {number} シート行番号（既存or新規）。失敗時0
 */
function prEnsureMakerSumRowTop_(sh, makerName, opt) {
  opt = opt || {};
  var parent = String(makerName || '').trim();
  if (!sh || !parent) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  if (idxMaker < 0) throw new Error('メーカーサマリに メーカー 列が無い');
  var sheetRow = 0;
  if (sh.getLastRow() >= 2) {
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      if (String(vals[r][idxMaker] || '').trim() === parent) {
        sheetRow = r + 2;
        break;
      }
    }
  }
  if (!sheetRow) {
    sh.insertRowBefore(2);
    sh.getRange(2, idxMaker + 1).setValue(parent);
    if (idxCk >= 0) {
      sh.getRange(2, idxCk + 1).insertCheckboxes();
      if (opt.ckTrue) sh.getRange(2, idxCk + 1).setValue(true);
    }
    sheetRow = 2;
    Logger.log('[仕入れ検討①] メーカーサマリ先頭挿入 maker=%s', parent);
  }
  prFillMakerSumJittaIfEmpty_(sh, sheetRow, parent);
  prEnsureMakerSumJittaDropdown_(sh);
  return sheetRow;
}

/** 非仕入先シードのみ（Generic 等）。 */
function prJittaExcludeSeed_(parent) {
  var p = String(parent || '').trim();
  if (!p) return '';
  if (PR_MAKER_JITTAI_EXCLUDE_SEED_[p]) return PR_MAKER_JITTAI_EXCLUDE_SEED_[p];
  var low = p.toLowerCase();
  var k;
  for (k in PR_MAKER_JITTAI_EXCLUDE_SEED_) {
    if (String(k).toLowerCase() === low) return PR_MAKER_JITTAI_EXCLUDE_SEED_[k];
  }
  return '';
}

/** 実態シード。非仕入先のみ（後方互換）。 */
function prJittaSeed_(parent) {
  return prJittaExcludeSeed_(parent);
}

function prEntryPreferredForJitta_(jitta) {
  if (jitta === PR_JITTAI_BRAND_) return 'ブランド';
  if (jitta === PR_JITTAI_MAKER_) return 'メーカー';
  if (jitta === PR_JITTAI_EXCLUDE_) return '非仕入先';
  if (jitta === PR_JITTAI_CHECK_) return '要確認';
  if (jitta === PR_JITTAI_UNCHECKED_) return '未調査';
  return '未設定';
}

/**
 * サマリ行の実態が空ならシード／未調査を書く。入口推奨も空なら連動。
 * 人が既に実態を書いている行は上書きしない。
 */
function prFillMakerSumJittaIfEmpty_(sh, sheetRow, parent) {
  if (!sh || !(sheetRow >= 2)) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxJ = prHeaderIndex_(hdr, '実態');
  var idxPref = prHeaderIndex_(hdr, '入口推奨');
  if (idxJ < 0) return;
  var cur = String(sh.getRange(sheetRow, idxJ + 1).getValue() || '').trim();
  var jitta = '';
  if (
    cur === PR_JITTAI_MAKER_ ||
    cur === PR_JITTAI_BRAND_ ||
    cur === PR_JITTAI_CHECK_ ||
    cur === PR_JITTAI_UNCHECKED_ ||
    cur === PR_JITTAI_EXCLUDE_
  ) {
    jitta = cur;
  } else {
    jitta = prJittaExcludeSeed_(parent) || PR_JITTAI_UNCHECKED_;
    sh.getRange(sheetRow, idxJ + 1).setValue(jitta);
    Logger.log('[仕入れ検討①] メーカーサマリ実態付与 maker=%s jitta=%s', parent, jitta);
  }
  if (idxPref >= 0) {
    var pref = String(sh.getRange(sheetRow, idxPref + 1).getValue() || '').trim();
    if (!pref) sh.getRange(sheetRow, idxPref + 1).setValue(prEntryPreferredForJitta_(jitta));
  }
  prSyncMakerJittaCkUrlRow_(sh, sheetRow);
}

/** Amazon.co.jp 検索（メーカー名入力済み）。 */
function prAmazonMakerSearchUrl_(makerName) {
  var m = String(makerName || '').trim();
  if (!m) return '';
  return 'https://www.amazon.co.jp/s?k=' + encodeURIComponent(m);
}

/** メーカーサマリ末尾に「実態ck」列を確保（既存列は動かさない）。 */
function prEnsureMakerJittaCkCol_(sh) {
  if (!sh) return -1;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = prHeaderIndex_(hdr, PR_JITTAI_CK_HEADER_);
  if (idx >= 0) return idx;
  sh.getRange(1, lastCol + 1).setValue(PR_JITTAI_CK_HEADER_);
  return lastCol;
}

/** メーカーサマリ末尾に セラーヒット 列を確保（CYCLE §5）。 */
function prEnsureMakerSumHitCol_(sh) {
  if (!sh) return -1;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = prHeaderIndex_(hdr, PR_MAKER_SUM_HIT_COL_);
  if (idx >= 0) return idx;
  sh.getRange(1, lastCol + 1).setValue(PR_MAKER_SUM_HIT_COL_);
  return lastCol;
}

/** maker_alias.resolve_alias 相当。 */
function prResolveMakerAlias_(name) {
  var t = prCleanEntryToken_(name);
  if (!t) return '';
  if (PR_MAKER_ALIASES_[t]) return PR_MAKER_ALIASES_[t];
  var low = t.toLowerCase();
  var k;
  for (k in PR_MAKER_ALIASES_) {
    if (String(k).toLowerCase() === low) return PR_MAKER_ALIASES_[k];
  }
  return t;
}

function prMakersInCell_(cell) {
  var out = [];
  var parts = String(cell || '').split('|');
  var i;
  for (i = 0; i < parts.length; i++) {
    var p = prResolveMakerAlias_(parts[i]);
    if (p) out.push(p);
  }
  return out;
}

function prProductParentName_(p) {
  var mfr = String((p && p.manufacturer) || '').trim();
  var brand = String((p && p.brand) || '').trim();
  return prResolveMakerAlias_(mfr || brand);
}

/** A-light: product cap 件まで manufacturer 抽出（dry_cycle_alight_hit 同期）。 */
function prExtractMakersFromProducts_(products, cap) {
  var seen = {};
  var ordered = [];
  var n = 0;
  var i;
  for (i = 0; i < (products || []).length; i++) {
    if (n >= cap) break;
    var parent = prProductParentName_(products[i]);
    n++;
    if (!parent || seen[parent]) continue;
    seen[parent] = true;
    ordered.push(parent);
  }
  ordered.sort();
  return ordered;
}

function prMergeExtractMakerCell_(old, newMakers) {
  var seen = {};
  var merged = [];
  var all = prMakersInCell_(old).concat(newMakers || []);
  var i;
  for (i = 0; i < all.length; i++) {
    var name = all[i];
    if (!name || seen[name]) continue;
    if (PR_HIT_EXCLUDE_SET_[name] || prJittaForMaker_(name) === PR_JITTAI_EXCLUDE_) continue;
    seen[name] = true;
    merged.push(name);
  }
  merged.sort();
  return merged.join('|');
}

function prHitEligibleMaker_(name) {
  if (!name || PR_HIT_EXCLUDE_SET_[name]) return false;
  if (prJittaForMaker_(name) === PR_JITTAI_EXCLUDE_) return false;
  return true;
}

/**
 * 1行: 実態=要確認なら Amazon 検索URL、それ以外は空。
 * M列「実態」は読取のみ。書込は「実態ck」のみ。
 */
function prSyncMakerJittaCkUrlRow_(sh, sheetRow) {
  if (!sh || !(sheetRow >= 2)) return;
  var idxCkUrl = prEnsureMakerJittaCkCol_(sh);
  if (idxCkUrl < 0) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxJ = prHeaderIndex_(hdr, '実態');
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  if (idxJ < 0 || idxMaker < 0) return;
  var jitta = String(sh.getRange(sheetRow, idxJ + 1).getValue() || '').trim();
  var maker = String(sh.getRange(sheetRow, idxMaker + 1).getValue() || '').trim();
  var url = jitta === PR_JITTAI_CHECK_ ? prAmazonMakerSearchUrl_(maker) : '';
  sh.getRange(sheetRow, idxCkUrl + 1).setValue(url);
}

/** 99. 実態=要確認行だけ「実態ck」に Amazon 検索URLを一括付与／不要行は空に。 */
function prMenuSyncMakerJittaCkUrls() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=jittai_ck_url state=RUNNING', runId);
  try {
    var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
    var idxCkUrl = prEnsureMakerJittaCkCol_(sh);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) {
      prMenuToast_('① 実態ck', 'メーカーサマリにデータ行がありません', 8);
      return;
    }
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxJ = prHeaderIndex_(hdr, '実態');
    var idxMaker = prHeaderIndex_(hdr, 'メーカー');
    if (idxJ < 0 || idxMaker < 0) throw new Error('メーカーサマリに 実態 または メーカー が無い');
    var n = lastRow - 1;
    var vals = sh.getRange(2, 1, n, lastCol).getValues();
    var out = [];
    var filled = 0;
    var r;
    for (r = 0; r < vals.length; r++) {
      var jitta = String(vals[r][idxJ] || '').trim();
      var maker = String(vals[r][idxMaker] || '').trim();
      var url = jitta === PR_JITTAI_CHECK_ ? prAmazonMakerSearchUrl_(maker) : '';
      if (url) filled++;
      out.push([url]);
    }
    sh.getRange(2, idxCkUrl + 1, n, 1).setValues(out);
    var detail = 'filled=' + filled + ' rows=' + n + ' col=' + PR_JITTAI_CK_HEADER_;
    Logger.log('[仕入れ検討①] runId=%s state=DONE %s', runId, detail);
    prAppendLog_([runId, '実態ck付与', 'DONE', new Date(), filled, n, '', detail]);
    prMenuToast_('① 実態ck', '要確認 ' + filled + '件に Amazon 検索URLを付与', 10);
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, '実態ck付与', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/** 洗い対象としてよい実態か（非仕入先・要確認は不可）。 */
function prMakerJittaWashEligible_(jitta) {
  var j = String(jitta || '').trim();
  return j === PR_JITTAI_MAKER_ || j === PR_JITTAI_BRAND_;
}

/** メーカーサマリ.実態 列にプルダウン（M列相当）。 */
function prEnsureMakerSumJittaDropdown_(sh) {
  if (!sh) return;
  var readyKey = PR_COMPETITOR_SS_ + ':' + PR_MAKER_SUM_SHEET_ + ':jitta_dd';
  if (PR_CK_READY_[readyKey]) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxJ = prHeaderIndex_(hdr, '実態');
  if (idxJ < 0) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [PR_JITTAI_MAKER_, PR_JITTAI_BRAND_, PR_JITTAI_CHECK_, PR_JITTAI_UNCHECKED_, PR_JITTAI_EXCLUDE_],
      true
    )
    .setAllowInvalid(false)
    .build();
  var endRow = Math.max(sh.getLastRow(), 2);
  var numRows = Math.max(endRow - 1, 1);
  sh.getRange(2, idxJ + 1, numRows, 1).setDataValidation(rule);
  PR_CK_READY_[readyKey] = true;
  Logger.log('[仕入れ検討①] メーカーサマリ実態プルダウン rows=%s', numRows);
}

/** Catalog 列挙見積。不明は null。サマリ Catalog見積列があれば優先。 */
function prMakerCatalogEstimate_(parent, washProgress, catalogEstCell) {
  parent = String(parent || '').trim();
  if (!parent) return null;
  if (catalogEstCell !== undefined && catalogEstCell !== null && String(catalogEstCell).trim() !== '') {
    var nCell = Number(String(catalogEstCell).replace(/[^\d.-]/g, ''));
    if (!isNaN(nCell) && nCell > 0) return nCell;
  }
  var q = prWashQueueUniqCount_(parent);
  if (q > 0) return q;
  var prog = String(washProgress || '');
  var m = prog.match(/queue=(\d+)/);
  if (m) return Number(m[1]) || null;
  return null;
}

/** 規模帯。Property PR_CYCLE_BAND_SELECT=false で旧 A/B に戻すときの互換も残す。 */
function prMakerCatalogBand_(estimate) {
  if (estimate == null || estimate === '') return PR_BAND_UNK_;
  var n = Number(estimate);
  if (isNaN(n)) return PR_BAND_UNK_;
  if (n <= PR_BAND_S_MAX_) return PR_BAND_S_;
  if (n <= PR_BAND_M_MAX_) return PR_BAND_M_;
  if (n <= PR_CATALOG_SPLIT_THRESHOLD_) return PR_BAND_L_;
  return PR_BAND_XL_;
}

function prMakerCatalogTier_(estimate) {
  if (!prCycleBandSelectEnabled_()) {
    if (estimate == null || estimate === '') return 'B';
    return Number(estimate) <= PR_CATALOG_SPLIT_THRESHOLD_ ? 'A' : 'B';
  }
  var b = prMakerCatalogBand_(estimate);
  return b === PR_BAND_S_ || b === PR_BAND_M_ || b === PR_BAND_L_ ? 'A' : 'B';
}

function prCycleBandSelectEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_CYCLE_BAND_SELECT_PROP_);
  if (raw == null || raw === '') return true;
  return String(raw).toLowerCase() !== 'false';
}

function prCycleHeartbeatPaused_() {
  return String(PropertiesService.getScriptProperties().getProperty(PR_CYCLE_HEARTBEAT_PAUSED_PROP_) || '') === '1';
}

/** 他メニュー開始時: 心拍停止＋次の01:00再開を予約。 */
function prCyclePauseForOtherMenu_(tag) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PR_CYCLE_HEARTBEAT_PAUSED_PROP_, '1');
  props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'false');
  try {
    prClearWashFullTriggers_({ reason: 'pause_other:' + String(tag || ''), log: true });
  } catch (e0) {}
  prCycleArmResume01_();
  Logger.log('[仕入れ検討①] cycle_pause tag=%s', String(tag || '').slice(0, 80));
}

/** 洗い候補なし時: 心拍トリガーを再武装（列挙／自動レ点／C0 等）。 */
function prCycleArmHeartbeatAfterIdle_(runId, tag) {
  if (!prCycleHeartbeatEnabled_()) return 0;
  PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
  return prScheduleWashFullTrigger_(1, {
    runId: runId,
    parent: '',
    tag: tag || 'heartbeat_idle',
  });
}

/**
 * 心拍 enumOnly 完了後: 規模帯だけ残し、Keepa中／列挙中を外す（大特大が小中より先に続き洗いされないため）。
 * キューASINは残す（小中尽きたあとの洗い用）。
 */
function prFinalizeEnumOnlySize_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return { ok: false, queueN: 0, band: '' };
  var queueN = prWashQueueUniqCount_(parent);
  var loc = prFindMakerSumRow_(parent);
  if (!loc || !loc.sh || !(loc.sheetRow > 0)) {
    return { ok: false, queueN: queueN, band: prMakerCatalogBand_(queueN || null) };
  }
  var cut501 = queueN >= PR_CATALOG_SPLIT_THRESHOLD_ + 1;
  prWriteMakerSumBand_(loc.sh, loc.sheetRow, queueN, cut501);
  var band = prMakerCatalogBand_(cut501 ? Math.max(queueN, PR_CATALOG_SPLIT_THRESHOLD_ + 1) : queueN);
  // 続き対象から外す（空＝未着手扱い。完了にはしない＝90日除外にしない）
  if (loc.idxSt >= 0) loc.sh.getRange(loc.sheetRow, loc.idxSt + 1).setValue('');
  if (loc.idxProg >= 0) {
    loc.sh.getRange(loc.sheetRow, loc.idxProg + 1).setValue('enum_done queue=' + queueN + ' band=' + band);
  }
  Logger.log(
    '[仕入れ検討①] enumOnly finalize parent=%s queue=%s band=%s',
    parent,
    queueN,
    band
  );
  return { ok: true, queueN: queueN, band: band };
}

/** Property PR_CYCLE_POOL_REUSE。未設定=ON。false で毎回数え直す旧動作。 */
function prCyclePoolReuseEnabled_() {
  return prCkEnabled_(PR_CYCLE_POOL_REUSE_PROP_);
}

/**
 * 選定プールの件数（原因ログ用）。1回で小中・不明を数える。
 * 2026-09-07: 同一実行内は数え直さない（呼び出しが10か所以上あり、全社ぶんの
 * 規模数え直しが重なって予算20分を食い潰していた）。切戻し PR_CYCLE_POOL_REUSE=false。
 */
function prCyclePoolCounts_() {
  var reuse = prCyclePoolReuseEnabled_();
  if (reuse && PR_CYCLE_POOL_CACHE_) return PR_CYCLE_POOL_CACHE_;
  var t0 = new Date().getTime();
  var pools = prCycleMakerBandPools_();
  var out = {
    smN: pools.byBand['小'].length + pools.byBand['中'].length,
    unkN: pools.unk.length,
    largeN: pools.byBand['大'].length,
    xlN: pools.byBand['特大'].length,
  };
  if (reuse) PR_CYCLE_POOL_CACHE_ = out;
  Logger.log(
    '[仕入れ検討①] pool_counts sm=%s unk=%s large=%s xl=%s elapsedMs=%s reuse=%s',
    out.smN,
    out.unkN,
    out.largeN,
    out.xlN,
    new Date().getTime() - t0,
    reuse ? 'Y' : 'N'
  );
  return out;
}

/** Property PR_MAKER_BAND_MAP。未設定=ON。false で1社ずつ読み直す旧動作。 */
function prMakerBandMapEnabled_() {
  return prCkEnabled_(PR_MAKER_BAND_MAP_PROP_);
}

/**
 * メーカーサマリを1回だけ読み「メーカー→{band, prog, est}」を作る。
 * 同名は先頭行を採る（prFindMakerSumRow_ と同じ）。
 * @return {!Object<string, {band:string, prog:string, est:*}>}
 */
function prMakerBandRowMap_() {
  if (PR_MAKER_BAND_CACHE_) return PR_MAKER_BAND_CACHE_;
  var out = {};
  var t0 = new Date().getTime();
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) {
    PR_MAKER_BAND_CACHE_ = out;
    return out;
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  if (idxM < 0) {
    PR_MAKER_BAND_CACHE_ = out;
    return out;
  }
  var idxBand = prHeaderIndex_(hdr, PR_BAND_COL_);
  var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
  var idxEst = prHeaderIndex_(hdr, PR_CATALOG_EST_COL_);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    var p = String(vals[r][idxM] || '').trim();
    if (!p || out[p]) continue;
    out[p] = {
      band: idxBand >= 0 ? String(vals[r][idxBand] || '').trim() : '',
      prog: idxProg >= 0 ? String(vals[r][idxProg] || '') : '',
      est: idxEst >= 0 ? vals[r][idxEst] : '',
    };
  }
  PR_MAKER_BAND_CACHE_ = out;
  Logger.log(
    '[仕入れ検討①] maker_band_map rows=%s makers=%s elapsedMs=%s',
    vals.length,
    Object.keys(out).length,
    new Date().getTime() - t0
  );
  return out;
}

/** サマリの規模帯／見積を書き換えたあとに呼ぶ（表を作り直させる）。 */
function prMakerBandMapClear_() {
  PR_MAKER_BAND_CACHE_ = null;
}

/** メーカーの規模帯（サマリ列優先）。 */
function prMakerBandFromSum_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return PR_BAND_UNK_;
  if (prMakerBandMapEnabled_()) {
    // 全社ループから呼ばれる。表を1回作って以降は読まない
    var row = prMakerBandRowMap_()[parent];
    if (!row) return PR_BAND_UNK_;
    if (row.band) return row.band;
    return prMakerCatalogBand_(prMakerCatalogEstimate_(parent, row.prog, row.est));
  }
  var loc = prFindMakerSumRow_(parent);
  if (!loc || !loc.sh) return PR_BAND_UNK_;
  var lastCol = Math.max(loc.sh.getLastColumn(), 1);
  var hdr = loc.sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxBand = prHeaderIndex_(hdr, PR_BAND_COL_);
  var idxEst = prHeaderIndex_(hdr, PR_CATALOG_EST_COL_);
  var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
  var row = loc.sh.getRange(loc.sheetRow, 1, 1, lastCol).getValues()[0];
  if (idxBand >= 0 && String(row[idxBand] || '').trim()) {
    return String(row[idxBand] || '').trim();
  }
  return prMakerCatalogBand_(
    prMakerCatalogEstimate_(
      parent,
      idxProg >= 0 ? String(row[idxProg] || '') : '',
      idxEst >= 0 ? row[idxEst] : ''
    )
  );
}

/** メーカー完走後の店レーン予約（小中・大・特大）。次の巨大より店を先に。 */
function prCycleMarkShopAfterSm_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return;
  PropertiesService.getScriptProperties().setProperty(PR_CYCLE_SHOP_AFTER_SM_PROP_, parent);
  Logger.log('[仕入れ検討①] shop_after_sm mark parent=%s', parent);
}

function prCycleShopAfterSmParent_() {
  return String(
    PropertiesService.getScriptProperties().getProperty(PR_CYCLE_SHOP_AFTER_SM_PROP_) || ''
  ).trim();
}

function prCycleClearShopAfterSm_() {
  PropertiesService.getScriptProperties().deleteProperty(PR_CYCLE_SHOP_AFTER_SM_PROP_);
}

function prCycleWeekKeyTokyo_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-ww');
}

function prCycleXlWeeklyCap_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_CYCLE_XL_WEEKLY_CAP_PROP_) ||
      PR_CYCLE_XL_WEEKLY_CAP_DEFAULT_
  );
  if (!(n > 0)) n = PR_CYCLE_XL_WEEKLY_CAP_DEFAULT_;
  return Math.floor(n);
}

/** @return {{key:string, count:number, makers:string[]}} */
function prCycleXlWeeklyState_() {
  var props = PropertiesService.getScriptProperties();
  var key = prCycleWeekKeyTokyo_();
  var cur = String(props.getProperty(PR_CYCLE_XL_WEEK_KEY_PROP_) || '');
  if (cur !== key) {
    props.setProperty(PR_CYCLE_XL_WEEK_KEY_PROP_, key);
    props.setProperty(PR_CYCLE_XL_WEEK_COUNT_PROP_, '0');
    props.deleteProperty(PR_CYCLE_XL_WEEK_MAKERS_PROP_);
    return { key: key, count: 0, makers: [] };
  }
  var makers = prCyclePipeListGet_(PR_CYCLE_XL_WEEK_MAKERS_PROP_);
  var count = Number(props.getProperty(PR_CYCLE_XL_WEEK_COUNT_PROP_) || 0) || 0;
  if (makers.length && count < makers.length) count = makers.length;
  return { key: key, count: count, makers: makers };
}

function prCycleXlWeeklyHasRoom_() {
  var st = prCycleXlWeeklyState_();
  return st.count < prCycleXlWeeklyCap_();
}

/**
 * 大／特大 Keepa 開始時に週間枠を消費。同一メーカーは週1回だけ。
 * @return {boolean} 開始してよい
 */
function prCycleXlWeeklyTryConsume_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return false;
  var st = prCycleXlWeeklyState_();
  if (st.makers.indexOf(parent) >= 0) return true;
  if (st.count >= prCycleXlWeeklyCap_()) {
    Logger.log('[仕入れ検討①] xl_week_cap_block parent=%s count=%s', parent, st.count);
    return false;
  }
  var props = PropertiesService.getScriptProperties();
  prCyclePipeListAdd_(PR_CYCLE_XL_WEEK_MAKERS_PROP_, parent);
  props.setProperty(PR_CYCLE_XL_WEEK_COUNT_PROP_, String(st.count + 1));
  Logger.log('[仕入れ検討①] xl_week_consume parent=%s n=%s', parent, st.count + 1);
  return true;
}

function prBrandHintRawMin_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_BRAND_HINT_RAW_MIN_PROP_) ||
      PR_BRAND_HINT_RAW_MIN_DEFAULT_
  );
  return n > 0 ? Math.floor(n) : PR_BRAND_HINT_RAW_MIN_DEFAULT_;
}

function prBrandHintRatio_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_BRAND_HINT_RATIO_PROP_) ||
      PR_BRAND_HINT_RATIO_DEFAULT_
  );
  if (!(n > 0) || n > 1) return PR_BRAND_HINT_RATIO_DEFAULT_;
  return n;
}

function prBrandHintAbsMin_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_BRAND_HINT_ABS_PROP_) ||
      PR_BRAND_HINT_ABS_DEFAULT_
  );
  return n > 0 ? Math.floor(n) : PR_BRAND_HINT_ABS_DEFAULT_;
}

/**
 * Catalog 生ヒットを最大 maxPages 頁サンプル（タイトルフィルタ前）。
 * @return {{rawN:number, titles:string[], brands:string[]}}
 */
function prCatalogSampleRawTitles_(entry, maxPages) {
  var titles = [];
  var brands = [];
  var rawN = 0;
  var token = '';
  var pages = 0;
  var maxP = maxPages > 0 ? maxPages : 5;
  while (pages < maxP) {
    var cat = prCatalogSearchPage_(entry, token);
    pages++;
    var i;
    for (i = 0; i < (cat.items || []).length; i++) {
      rawN++;
      var tit = String(cat.items[i].title || '').trim();
      if (tit) titles.push(tit);
      var br = String(cat.items[i].brand || '').trim();
      if (br) brands.push(br);
    }
    token = cat.nextToken || '';
    if (!token) break;
    Utilities.sleep(200);
  }
  return { rawN: rawN, titles: titles, brands: brands };
}

/** 法人格・一般語を除いた短縮候補。 */
function prBrandHintStripLegal_(name) {
  var s = String(name || '').trim();
  s = s.replace(/株式会社|有限会社|合同会社|合資会社|一般社団法人|一般財団法人/g, '');
  s = s.replace(/^\s+|\s+$/g, '');
  return s;
}

/**
 * 実態AI／マスタ突合用のメーカーキー（法人格ゆれ除去）。
 * モンテ物産＝モンテ物産(株)。表示名は変えない。
 */
function prNormalizeMakerKey_(name) {
  var s = String(name || '').trim();
  if (!s) return '';
  s = s.replace(/株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|一般財団法人/g, '');
  s = s.replace(/[（(]株[）)]/g, '');
  s = s.replace(/[（(]有[）)]/g, '');
  s = s.replace(/㈱|㈲/g, '');
  s = s.replace(/\s+/g, '');
  return s;
}

function prMakerCanonicalScore_(name) {
  var n = String(name || '').trim();
  var score = 1000 - n.length;
  if (/株式会社|[（(]株[）)]|㈱|有限会社|[（(]有[）)]|㈲|合同会社|合資会社|合名会社/.test(n)) {
    score -= 200;
  }
  return score;
}

function prPickCanonicalMakerName_(names) {
  var best = '';
  var bestScore = -99999;
  var i;
  for (i = 0; i < (names || []).length; i++) {
    var n = String(names[i] || '').trim();
    if (!n) continue;
    var sc = prMakerCanonicalScore_(n);
    if (!best || sc > bestScore) {
      best = n;
      bestScore = sc;
    }
  }
  return best;
}

function prBrandHintIsExcludedToken_(tok) {
  tok = String(tok || '').trim();
  if (!tok) return true;
  var ban = {
    食品: true,
    製菓: true,
    商事: true,
    産業: true,
    工業: true,
    本舗: true,
    商店: true,
    会社: true,
    メーカー: true,
    ブランド: true,
  };
  return !!ban[tok];
}

/**
 * タイトル群から短縮ブランド候補を選ぶ（§6.6 閾値）。
 * @return {{ok:boolean, brand:string, rawN:number, hitN:number, ratio:number, detail:string}}
 */
function prPickBrandHintFromTitles_(legalName, sample) {
  sample = sample || { rawN: 0, titles: [], brands: [] };
  var rawN = Number(sample.rawN) || 0;
  var titles = sample.titles || [];
  var rawMin = prBrandHintRawMin_();
  if (rawN < rawMin) {
    return {
      ok: false,
      brand: '',
      rawN: rawN,
      hitN: 0,
      ratio: 0,
      detail: 'raw_below_min raw=' + rawN + ' need=' + rawMin,
    };
  }
  var candidates = [];
  var stripped = prBrandHintStripLegal_(legalName);
  if (stripped && stripped !== legalName) candidates.push(stripped);
  if (stripped.length >= 2 && stripped.length <= 8) candidates.push(stripped);
  // Catalog brand 列の最多も候補
  var brandCnt = {};
  var bi;
  for (bi = 0; bi < (sample.brands || []).length; bi++) {
    var b = String(sample.brands[bi] || '').trim();
    if (!b || prBrandHintIsExcludedToken_(b)) continue;
    var bLen = b.length;
    if (bLen < 2 || bLen > 8) continue;
    brandCnt[b] = (brandCnt[b] || 0) + 1;
  }
  var bk;
  for (bk in brandCnt) {
    if (Object.prototype.hasOwnProperty.call(brandCnt, bk)) candidates.push(bk);
  }
  var uniq = {};
  var list = [];
  var c;
  for (c = 0; c < candidates.length; c++) {
    var tok = String(candidates[c] || '').trim();
    if (!tok || uniq[tok]) continue;
    if (tok === String(legalName || '').trim()) continue;
    if (prBrandHintIsExcludedToken_(tok)) continue;
    if (tok.length < 2 || tok.length > 8) continue;
    uniq[tok] = true;
    list.push(tok);
  }
  var ratioNeed = prBrandHintRatio_();
  var absNeed = prBrandHintAbsMin_();
  var best = null;
  var ti;
  for (c = 0; c < list.length; c++) {
    var needle = list[c];
    var hit = 0;
    for (ti = 0; ti < titles.length; ti++) {
      if (String(titles[ti] || '').indexOf(needle) >= 0) hit++;
    }
    var ratio = titles.length ? hit / titles.length : 0;
    if (hit < absNeed || ratio < ratioNeed) continue;
    if (!best || hit > best.hitN) {
      best = { brand: needle, hitN: hit, ratio: ratio };
    }
  }
  if (!best) {
    return {
      ok: false,
      brand: '',
      rawN: rawN,
      hitN: 0,
      ratio: 0,
      detail: 'no_token_pass titles=' + titles.length,
    };
  }
  return {
    ok: true,
    brand: best.brand,
    rawN: rawN,
    hitN: best.hitN,
    ratio: best.ratio,
    detail:
      'brand=' +
      best.brand +
      ' hit=' +
      best.hitN +
      '/' +
      titles.length +
      ' raw=' +
      rawN,
  };
}

/**
 * title0 時にブランド行を追加（既存同名は増やさない）。
 * @return {{added:boolean, brand:string, detail:string}}
 */
function prTryAddBrandHintMaker_(legalParent) {
  legalParent = String(legalParent || '').trim();
  if (!legalParent) return { added: false, brand: '', detail: 'no_parent' };
  var entry = prMakerCatalogEntry_(legalParent, '', '');
  var sample;
  try {
    sample = prCatalogSampleRawTitles_(entry, 5);
  } catch (eCat) {
    return {
      added: false,
      brand: '',
      detail: 'catalog_fail ' + String((eCat && eCat.message) || eCat).slice(0, 80),
    };
  }
  var pick = prPickBrandHintFromTitles_(legalParent, sample);
  if (!pick.ok) {
    Logger.log('[仕入れ検討①] brand_hint skip parent=%s %s', legalParent, pick.detail);
    return { added: false, brand: '', detail: pick.detail };
  }
  var brand = pick.brand;
  var existed = !!prFindMakerSumRow_(brand);
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh) return { added: false, brand: brand, detail: 'no_sum' };
  var sheetRow = prEnsureMakerSumRowTop_(sh, brand, {});
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxJ = prHeaderIndex_(hdr, '実態');
  var idxPref = prHeaderIndex_(hdr, '入口推奨');
  var idxMemo = prHeaderIndex_(hdr, 'メモ');
  if (idxJ >= 0) {
    var curJ = String(sh.getRange(sheetRow, idxJ + 1).getValue() || '').trim();
    if (!curJ || curJ === PR_JITTAI_CHECK_ || curJ === PR_JITTAI_UNCHECKED_ || !existed) {
      sh.getRange(sheetRow, idxJ + 1).setValue(PR_JITTAI_BRAND_);
    }
  }
  if (idxPref >= 0) {
    var curP = String(sh.getRange(sheetRow, idxPref + 1).getValue() || '').trim();
    if (!curP || !existed) sh.getRange(sheetRow, idxPref + 1).setValue('ブランド');
  }
  if (idxMemo >= 0) {
    var memo = String(sh.getRange(sheetRow, idxMemo + 1).getValue() || '');
    var tag = 'brand_hint_from:' + legalParent;
    if (memo.indexOf(tag) < 0) {
      sh.getRange(sheetRow, idxMemo + 1).setValue((memo ? memo + ' ' : '') + tag);
    }
  }
  Logger.log(
    '[仕入れ検討①] brand_hint parent=%s → %s row=%s existed=%s %s',
    legalParent,
    brand,
    sheetRow,
    existed ? 'Y' : 'N',
    pick.detail
  );
  return {
    added: !existed,
    brand: brand,
    detail: pick.detail + ' row=' + sheetRow + (existed ? ' existed' : ' new'),
  };
}

/**
 * 大／特大の Keepa 続きを後回しにするか。
 * @return {{defer:boolean, reason:string, smN:number, unkN:number, largeN:number, band:string}}
 */
function prWashContinueDeferInfo_(parent, optCounts) {
  parent = String(parent || '').trim();
  var counts = optCounts || { smN: 0, unkN: 0, largeN: 0, xlN: 0 };
  if (!optCounts) {
    try {
      counts = prCyclePoolCounts_();
    } catch (e0) {
      return {
        defer: false,
        reason: 'pool_err=' + String((e0 && e0.message) || e0).slice(0, 60),
        smN: 0,
        unkN: 0,
        largeN: 0,
        band: '',
      };
    }
  }
  var largeN = Number(counts.largeN) || 0;
  var band = prMakerBandFromSum_(parent);
  var isLarge = band === PR_BAND_L_ || band === PR_BAND_XL_;
  if (!isLarge) {
    return {
      defer: false,
      reason: 'band=' + band + ' not_large allow_continue',
      smN: counts.smN,
      unkN: counts.unkN,
      largeN: largeN,
      band: band,
    };
  }
  // 小・中が残る、または規模不明（未列挙）が残る間は大特大を洗わない
  if (counts.smN > 0) {
    return {
      defer: true,
      reason: 'defer_xl_sm_remain sm=' + counts.smN + ' unk=' + counts.unkN + ' band=' + band,
      smN: counts.smN,
      unkN: counts.unkN,
      largeN: largeN,
      band: band,
    };
  }
  if (counts.unkN > 0) {
    return {
      defer: true,
      reason: 'defer_xl_unk_remain sm=0 unk=' + counts.unkN + ' band=' + band,
      smN: counts.smN,
      unkN: counts.unkN,
      largeN: largeN,
      band: band,
    };
  }
  // §4.4: 小中=0・店ラウンド未尽かつ店レーン材料ありなら大続きを選定しない
  // （材料なしは心拍と同様に大へ進んでよい。pick→defer 禁止）
  if (counts.smN === 0 && prAlightShouldShopBeforeLarge_()) {
    var shopAct = null;
    try {
      shopAct = prCycleShopBeforeLargeAction_({});
    } catch (eShop) {
      shopAct = { action: 'alight' };
    }
    if (shopAct) {
      return {
        defer: true,
        reason:
          'defer_shop_before_large action=' +
          shopAct.action +
          ' rounds=' +
          prAlightRoundCount_() +
          ' done=' +
          prAlightDailyDoneCount_() +
          '/' +
          prAlightDailyMax_() +
          ' band=' +
          band,
        smN: counts.smN,
        unkN: counts.unkN,
        largeN: largeN,
        band: band,
      };
    }
  }
  // 大が洗えるのに特大の続き／キューを先に進めない（小→中→大→特大）
  if (band === PR_BAND_XL_ && largeN > 0) {
    return {
      defer: true,
      reason: 'defer_xl_large_remain large=' + largeN + ' band=' + band,
      smN: counts.smN,
      unkN: counts.unkN,
      largeN: largeN,
      band: band,
    };
  }
  // 週間 N 枠（新規開始相当）。既に枠内の同一メーカーは続き可
  try {
    var xlSt = prCycleXlWeeklyState_();
    var cap = prCycleXlWeeklyCap_();
    if (xlSt.makers.indexOf(parent) < 0 && xlSt.count >= cap) {
      return {
        defer: true,
        reason:
          'defer_xl_week_cap count=' + xlSt.count + '/' + cap + ' band=' + band,
        smN: counts.smN,
        unkN: counts.unkN,
        largeN: largeN,
        band: band,
      };
    }
  } catch (eCap) {
    Logger.log(
      '[仕入れ検討①] WARN xl_week_defer %s',
      String((eCap && eCap.message) || eCap).slice(0, 80)
    );
  }
  return {
    defer: false,
    reason: 'allow_xl sm=0 unk=0 large=' + largeN + ' band=' + band,
    smN: counts.smN,
    unkN: counts.unkN,
    largeN: largeN,
    band: band,
  };
}

/** 小中／不明残りで誤って Keepa中のときだけ帯確定に戻す。週間枠・大残りは進捗を残す。 */
function prWashContinueDeferShouldFinalize_(reason) {
  reason = String(reason || '');
  return (
    reason.indexOf('defer_xl_sm_remain') >= 0 || reason.indexOf('defer_xl_unk_remain') >= 0
  );
}

/** 小・中がまだ洗える間／不明が残る間は、大／特大の Keepa 続きを後回しにする。 */
function prWashContinueShouldDefer_(parent) {
  return !!prWashContinueDeferInfo_(parent).defer;
}

/**
 * 誤って Keepa中／列挙中の大・特大を規模確定扱いに戻す（不明 or 小中が残るとき）。
 * @return {{n:number, detail:string}}
 */
function prSanitizeDeferredLargeKeepaStuck_() {
  var counts = { smN: 0, unkN: 0 };
  try {
    counts = prCyclePoolCounts_();
  } catch (e0) {
    return { n: 0, detail: 'sanitize_skip pool_err' };
  }
  if (counts.smN <= 0 && counts.unkN <= 0) {
    return { n: 0, detail: 'sanitize_skip sm=0 unk=0 no_defer_gate' };
  }
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return { n: 0, detail: 'sanitize_skip no_sum' };
  prEnsureMakerWashStatusCols_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
  if (idxM < 0 || idxSt < 0) return { n: 0, detail: 'sanitize_skip no_cols' };
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var fixed = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    var st = String(vals[r][idxSt] || '').trim();
    if (st !== PR_WASH_ST_KEEPA_ && st !== PR_WASH_ST_ENUM_) continue;
    var parent = String(vals[r][idxM] || '').trim();
    if (!parent) continue;
    var info = prWashContinueDeferInfo_(parent);
    if (!info.defer) continue;
    prFinalizeEnumOnlySize_(parent);
    fixed.push(parent + ':' + info.band);
  }
  var detail =
    'sanitize_fixed n=' +
    fixed.length +
    ' sm=' +
    counts.smN +
    ' unk=' +
    counts.unkN +
    (fixed.length ? ' makers=' + fixed.join('|').slice(0, 120) : '');
  Logger.log('[仕入れ検討①] %s', detail);
  return { n: fixed.length, detail: detail };
}

/** 手動再開: 停止解除・01:00削除・心拍再武装。opt.immediate=true なら可能なら即 prDispatchHeartbeat_。 */
function prCycleResumeManual_(opt) {
  opt = opt || {};
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PR_CYCLE_HEARTBEAT_PAUSED_PROP_);
  prCycleClearResume01_();
  props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
  var runId = (opt && opt.runId) || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var detail = 'manual_resume 01_cleared';
  if (opt.immediate && prCycleHeartbeatEnabled_() && prWashLockLeftSec_() <= 0) {
    prDispatchHeartbeat_(runId);
    detail += ' heartbeat_dispatched';
  } else {
    var afterMin = prScheduleWashFullTrigger_(1, {
      runId: runId,
      parent: '',
      tag: opt.tag || 'cycle_manual_resume',
    });
    detail += ' trigger_armed afterMin=' + afterMin;
    if (prWashLockLeftSec_() > 0) detail += ' lock_busy';
  }
  prCycleArmDailyMailTrigger_();
  prCycleArmLivenessTrigger_();
  prCycleArmListCheckTrigger_();
  prAppendLog_([runId, opt.logLabel || '自動サイクル再開', 'DONE', new Date(), 1, '', '', detail]);
  return { runId: runId, detail: detail };
}


function prMenuCycleResume() {
  prCycleResumeManual_({});
  prMenuToast_('自動サイクル再開', '01:00 トリガーを消し、約1分後に心拍を予約しました', 8);
}

/**
 * A-1: 店レーンの日次カウンタをリセット（自動レ点・A-light完走・ラウンド・新メーカー印）。
 * 週間大特大枠・メーカー洗い進捗・キューは触らない。テスト連打でも店ラウンドをやり直せる。
 * @return {string} ログ用要約
 */
function prCycleA1ResetShopDayCounters_() {
  var props = PropertiesService.getScriptProperties();
  var before =
    'autoCk=' +
    prCycleAutoCkAlreadyToday_() +
    '/' +
    prCycleAutoCkDailyCap_() +
    ' alightDone=' +
    prAlightDailyDoneCount_() +
    '/' +
    prAlightDailyMax_() +
    ' rounds=' +
    prAlightRoundCount_() +
    ' lastNew=' +
    prAlightLastNewMakers_();
  try {
    props.deleteProperty(PR_CYCLE_AUTO_CK_DAY_PROP_);
    props.deleteProperty(PR_CYCLE_AUTO_CK_COUNT_PROP_);
  } catch (e1) {}
  try {
    props.deleteProperty(PR_ALIGHT_DAILY_DONE_DAY_PROP_);
    props.deleteProperty(PR_ALIGHT_DAILY_DONE_COUNT_PROP_);
  } catch (e2) {}
  try {
    props.deleteProperty(PR_ALIGHT_ROUND_DAY_PROP_);
    props.deleteProperty(PR_ALIGHT_ROUND_COUNT_PROP_);
  } catch (e3) {}
  try {
    props.setProperty(PR_ALIGHT_LAST_NEW_PROP_, '0');
  } catch (e4) {}
  try {
    prAlightContSet_(null);
  } catch (e5) {}
  return 'shop_day_reset before={' + before + '} after=0';
}

/**
 * A-1: 実行中を一時停止（a1_park）→ abort 世代↑ →時計リセット→心拍でフロー先頭から再選定。
 * 孤児ロックは破棄。キュー・Keepa済・Catalog見積は残す。D1 未の完走社だけ C4 recover。
 * 店レーン日次カウンタ（自動レ点／A-light／ラウンド）はリセットする。週間大特大枠は残す。
 */
function prMenuCycleResumeNow() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var park = prCycleA1ParkInflight_();
  var abortN = prWashAbortBump_('A1');
  var reset = prCycleHardResetClocks_(runId, 'A1');
  // 旧実行が握ったロック／再取得レースを落とし、即心拍で catalog_first できるようにする
  var rec = prWashFullLockReconcile_('A1');
  try {
    prWashFullUnlock_();
  } catch (eUn) {}
  var shopReset = '';
  try {
    shopReset = prCycleA1ResetShopDayCounters_();
  } catch (eShop) {
    shopReset = 'shop_day_reset_err=' + String((eShop && eShop.message) || eShop).slice(0, 60);
  }
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PR_CYCLE_HEARTBEAT_PAUSED_PROP_);
  prCycleClearResume01_();
  props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
  prCycleArmDailyMailTrigger_();
  prCycleArmLivenessTrigger_();
  prCycleArmListCheckTrigger_();
  var detail =
    'A1_park n=' +
    park.n +
    ' abortEpoch=' +
    abortN +
    ' lock_rec=' +
    (rec.detail || '') +
    ' shop_lane_reprioritize=1 ' +
    shopReset +
    ' alightDone=' +
    prAlightDailyDoneCount_() +
    '/' +
    prAlightDailyMax_() +
    ' alightRounds=' +
    prAlightRoundCount_() +
    ' autoCkToday=' +
    prCycleAutoCkAlreadyToday_() +
    '/' +
    prCycleAutoCkDailyCap_() +
    ' ' +
    (reset.detail || '');
  if (park.detail) detail += ' ' + park.detail;
  var c4Parent = prCycleMaybeArmC4ForLastComplete_(runId, 'A1_c4_recover');
  if (c4Parent) detail += ' c4_armed parent=' + c4Parent;
  else detail += ' c4_skip';
  if (prCycleHeartbeatEnabled_()) {
    // ロックは上で破棄済み。busy でも心拍へ（中で再 reconcile）
    prDispatchHeartbeat_(runId);
    detail += ' heartbeat_dispatched';
  } else {
    var afterMin = prScheduleWashFullTrigger_(1, {
      runId: runId,
      parent: '',
      tag: 'A1_resume',
    });
    detail += ' trigger_armed afterMin=' + afterMin;
  }
  prAppendLog_([runId, 'A-1 全自動開始／再開', 'DONE', new Date(), 1, '', '', detail.slice(0, 200)]);
  prMenuToast_('A-1 全自動', detail.slice(0, 160), 12);
}

/** C. 見積・発注の枠（E〜J 未実装）。 */
function prMenuEstimateOrderPlaceholder() {
  prMenuToast_('C. 見積・発注', '準備中です（見積提案〜発注は次スライス）', 8);
}

/**
 * A-2: 洗いトリガー＋心拍生存をまとめて①ログへ。沈黙なら再武装。
 */
function prMenuCycleFullDiagnose() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var snap = prCycleAliveReason_();
  var clocks = prCycleClockSnapshot_();
  var missed = prDetectWashTriggerMissed_(runId, 'A2_full_diagnose');
  var ensure = prCycleEnsureLiveness_(runId, 'A2_full_diagnose');
  var detail =
    'full_diag via=A2 now=' +
    prNowFetchedAt_() +
    ' alive=' +
    snap.reason +
    ' ensure=' +
    ensure.reason +
    ' rearmed=' +
    (ensure.rearmed ? '1' : '0') +
    (missed ? ' missed=1' : ' missed=0') +
    ' ' +
    clocks;
  prAppendLog_([
    runId,
    'A-2 全診断',
    'DONE',
    new Date(),
    ensure.rearmed ? 1 : 0,
    '',
    '',
    detail.slice(0, 200),
  ]);
  prMenuToast_(
    'A-2 全診断',
    (ensure.rearmed ? '再武装した。' : '') +
      (missed ? '未発火あり。' : '') +
      detail.slice(0, 160),
    12
  );
}

function prMenuCycleDailyMailNow() {
  prRunCycleDailyMail();
  prMenuToast_('日次メール', '送信を試行しました（①ログ確認）', 6);
}

function prMinutesUntilNext01Jst_() {
  var now = new Date();
  var tz = 'Asia/Tokyo';
  var ymd = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var parts = ymd.split('-');
  var target = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2]),
    1,
    0,
    0
  );
  // Utilities のローカル解釈ずれ対策: 東京オフセットで再計算
  var nowTok = Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH:mm:ss");
  var h = Number(Utilities.formatDate(now, tz, 'H'));
  var mi = Number(Utilities.formatDate(now, tz, 'm'));
  var minsNow = h * 60 + mi;
  var mins01 = 60; // 01:00
  var delta = mins01 - minsNow;
  if (delta <= 2) delta += 24 * 60;
  return Math.max(3, delta);
}

function prCycleArmResume01_() {
  prCycleClearResume01_();
  var afterMin = prMinutesUntilNext01Jst_();
  ScriptApp.newTrigger(PR_CYCLE_RESUME_01_FN_)
    .timeBased()
    .after(afterMin * 60 * 1000)
    .create();
  Logger.log('[仕入れ検討①] resume01 armed afterMin=%s', afterMin);
}

function prCycleClearResume01_() {
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === PR_CYCLE_RESUME_01_FN_) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/** 日次生存メールの時計トリガー（1日1回・JST 8時付近）。無ければ張る。 */
function prCycleArmDailyMailTrigger_() {
  if (!prCycleDailyMailEnabled_()) return;
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === PR_CYCLE_DAILY_MAIL_FN_) return;
  }
  ScriptApp.newTrigger(PR_CYCLE_DAILY_MAIL_FN_).timeBased().everyDays(1).atHour(8).create();
  Logger.log('[仕入れ検討①] daily_mail trigger armed atHour=8');
}

function prCycleListCheckEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_CYCLE_LIST_CHECK_ENABLED_PROP_);
  if (raw == null || raw === '') return true;
  return String(raw).toLowerCase() !== 'false';
}

function prCycleListCheckDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_CYCLE_LIST_CHECK_DAYS_PROP_));
  if (!(n > 0) || n > 90) return PR_CYCLE_LIST_CHECK_DAYS_DEFAULT_;
  return Math.floor(n);
}

/** 週次転記チェックの時計（日曜 JST 9時付近）。無ければ張る。 */
function prCycleArmListCheckTrigger_() {
  if (!prCycleListCheckEnabled_()) return;
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === PR_CYCLE_LIST_CHECK_FN_) return;
  }
  ScriptApp.newTrigger(PR_CYCLE_LIST_CHECK_FN_)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(9)
    .create();
  Logger.log('[仕入れ検討①] list_check trigger armed Sunday atHour=9');
}

function prCountHandlerTriggers_(fnName) {
  var n = 0;
  var list = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === fnName) n++;
  }
  return n;
}

function prCountResume01Triggers_() {
  return prCountHandlerTriggers_(PR_CYCLE_RESUME_01_FN_);
}

function prCountLivenessTriggers_() {
  return prCountHandlerTriggers_(PR_CYCLE_LIVENESS_FN_);
}

/** 心拍の時計スナップショット（原因追及用・①ログ1行）。 */
function prCycleClockSnapshot_() {
  return (
    'washN=' +
    prCountWashFullTriggers_() +
    ' c4N=' +
    prCountCycleC4Triggers_() +
    ' resume01N=' +
    prCountResume01Triggers_() +
    ' liveN=' +
    prCountLivenessTriggers_() +
    ' paused=' +
    (prCycleHeartbeatPaused_() ? '1' : '0') +
    ' hb=' +
    (prCycleHeartbeatEnabled_() ? 'ON' : 'OFF') +
    ' washEngine=' +
    (prMakerWashEnabled_() && prMakerWashFullEnabled_() ? 'ON' : 'OFF') +
    ' inflight=' +
    prWashInflightOwner_() +
    ' phase=' +
    String(PropertiesService.getScriptProperties().getProperty(PR_WASH_INFLIGHT_PHASE_PROP_) || '') +
    ' ageMs=' +
    prWashInflightAgeMs_()
  );
}

/**
 * 規模帯の確定（Catalog列挙）を1実行で複数社まとめて回す。
 * 打ち切り: 予算残り不足／小・中が1社出た（洗いを待たせない）／対象切れ。
 * 1社の例外は飛ばして次へ（タイトル含有0は保留処理だけして継続）。
 * @param {string} runId
 * @param {string} firstParent 心拍が選んだ先頭社（必ず最初に処理する）
 * @return {{n:number, detail:string}}
 */
function prCycleRunEnumBatch_(runId, firstParent) {
  var cap = prCycleEnumBatchN_();
  var order = [];
  var seen = {};
  firstParent = String(firstParent || '').trim();
  if (firstParent) {
    order.push(firstParent);
    seen[firstParent] = true;
  }
  var pend = prCycleCatalogPendingList_(cap + 1);
  var i;
  for (i = 0; i < pend.length && order.length < cap; i++) {
    var p = String(pend[i].parent || '').trim();
    if (!p || seen[p]) continue;
    seen[p] = true;
    order.push(p);
  }
  var t0 = new Date().getTime();
  var bands = {};
  var doneN = 0;
  var skipN = 0;
  var stopBy = 'list_end';
  for (i = 0; i < order.length; i++) {
    if (i > 0 && !prCycleEnumBatchHasBudget_()) {
      stopBy = 'budget';
      break;
    }
    var parent = order[i];
    try {
      prRunMakerWashFull_({
        parent: parent,
        entry: prMakerCatalogEntry_(parent, '', ''),
        needles: prMakerWashNeedles_(parent, ''),
        tag: 'hb_enum',
        runId: runId,
        loopKeepa: false,
        quiet: true,
        armTrigger: false,
        enumOnly: true,
      });
      var fin = prFinalizeEnumOnlySize_(parent);
      doneN++;
      var band = String(fin.band || PR_BAND_UNK_);
      bands[band] = (bands[band] || 0) + 1;
      Logger.log(
        '[仕入れ検討①] enum_batch %s/%s parent=%s band=%s queue=%s',
        i + 1,
        order.length,
        parent,
        band,
        fin.queueN
      );
      // 小・中が出たら列挙をやめて洗いへ譲る（小中優先を崩さない）
      if (band === PR_BAND_S_ || band === PR_BAND_M_) {
        stopBy = 'sm_found=' + parent;
        break;
      }
    } catch (eB) {
      var msgB = prRedactKey_(String((eB && eB.message) || eB));
      if (/タイトル含有0/.test(msgB)) {
        try {
          prHandleWashTitleZero_({
            parent: parent,
            runId: runId,
            tag: 'hb_enum_title0',
            armTrigger: false,
            quiet: true,
          });
        } catch (eT) {}
      }
      skipN++;
      Logger.log(
        '[仕入れ検討①] WARN enum_batch_skip parent=%s %s',
        parent,
        msgB.slice(0, 120)
      );
    }
  }
  var bandStr = Object.keys(bands)
    .map(function (k) {
      return k + '=' + bands[k];
    })
    .join(' ');
  return {
    n: doneN,
    detail:
      'enum_batch n=' +
      doneN +
      '/' +
      order.length +
      ' cap=' +
      cap +
      ' skip=' +
      skipN +
      ' stop=' +
      stopBy +
      (bandStr ? ' ' + bandStr : '') +
      ' elapsedMs=' +
      (new Date().getTime() - t0),
  };
}

/** 1実行でまとめて規模帯を確定する社数（Property。1=旧動作の1社ずつ）。 */
function prCycleEnumBatchN_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_CYCLE_ENUM_BATCH_N_PROP_);
  var n = Number(raw);
  if (raw == null || String(raw).trim() === '' || isNaN(n) || n < 1) {
    return PR_CYCLE_ENUM_BATCH_N_DEFAULT_;
  }
  return Math.min(Math.floor(n), PR_CYCLE_ENUM_BATCH_N_MAX_);
}

/** まとめ列挙を続けてよいか（予算残りに予約・ログぶんを残す）。 */
function prCycleEnumBatchHasBudget_() {
  if (!PR_GAS_START_MS_) return true;
  var elapsed = new Date().getTime() - PR_GAS_START_MS_;
  return PR_GAS_BUDGET_MS_ - elapsed > PR_CYCLE_ENUM_BATCH_RESERVE_MS_;
}

/** Keepa 最大回復待ち＋α（分）。腐った予約／実行印のリセット閾値。 */
function prCycleStaleLimitMin_() {
  var maxWait = Math.ceil(prKeepaTokenBucketMax_() / prKeepaRefillRate_()) + 1;
  return maxWait + PR_CYCLE_STALE_ALPHA_MIN_;
}

function prCycleStaleLimitMs_() {
  return prCycleStaleLimitMin_() * 60 * 1000;
}

/** Property PR_CYCLE_INFLIGHT_STALE_FAST。未設定=ON。false で旧66分基準。 */
function prCycleInflightStaleFastEnabled_() {
  return prCkEnabled_(PR_CYCLE_INFLIGHT_STALE_FAST_PROP_);
}

/** 実行中印だけの腐り上限（ms）。30分壁基準の35分（洗い本体の stale と同値）。 */
function prCycleInflightStaleLimitMs_() {
  var tokenBased = prCycleStaleLimitMs_();
  if (!prCycleInflightStaleFastEnabled_()) return tokenBased;
  return Math.min(PR_WASH_INFLIGHT_STALE_MS_, tokenBased);
}

/** 洗い／C4 の次発火予定（ms）。無いときは 0。 */
function prCycleNextExpectedFireMs_() {
  var props = PropertiesService.getScriptProperties();
  var wash = Number(props.getProperty(PR_WASH_TRIGGER_EXPECT_FIRE_MS_PROP_) || 0);
  var c4 = Number(props.getProperty(PR_CYCLE_C4_TRIGGER_EXPECT_MS_PROP_) || 0);
  var next = 0;
  if (wash > 0) next = wash;
  if (c4 > 0 && (next === 0 || c4 < next)) next = c4;
  return next;
}

/**
 * A-1 専用: 実行中メーカーを一時停止し、心拍が catalog_first／帯選定から再選定できるようにする。
 * ①洗いキュー・Keepa済・Catalog見積・意図保留（title0／empty_research 等）は触らない。
 * A-light 続き JSON は消す（店レーンも先頭から）。
 * @return {{n:number, detail:string}}
 */
function prCycleA1ParkInflight_() {
  var parked = [];
  try {
    prAlightContSet_(null);
  } catch (eAl) {}
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) {
    return { n: 0, detail: 'a1_park n=0 no_sum' };
  }
  prEnsureMakerWashStatusCols_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
  var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
  if (idxM < 0 || idxSt < 0) {
    return { n: 0, detail: 'a1_park n=0 no_cols' };
  }
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var inflight = {};
  inflight[PR_WASH_ST_ENUM_] = true;
  inflight[PR_WASH_ST_INFER_] = true;
  inflight[PR_WASH_ST_KEEPA_] = true;
  inflight[PR_WASH_ST_NEED_Q2_] = true;
  var r;
  for (r = 0; r < vals.length; r++) {
    var st = String(vals[r][idxSt] || '').trim();
    if (!inflight[st]) continue;
    var parent = String(vals[r][idxM] || '').trim();
    if (!parent) continue;
    var oldProg = idxProg >= 0 ? String(vals[r][idxProg] || '') : '';
    if (
      oldProg.indexOf('title0_skip') >= 0 ||
      oldProg.indexOf('empty_research') >= 0 ||
      oldProg.indexOf('a1_park') >= 0
    ) {
      continue;
    }
    var left = 0;
    try {
      left = prWashQueuePendingCount_(parent);
    } catch (eL) {}
    var prog =
      'a1_park was=' +
      st +
      (left > 0 ? ' left=' + left : '') +
      (oldProg ? ' | ' + oldProg.slice(0, 200) : '');
    prSetMakerWashStatus_(parent, PR_WASH_ST_NONE_, prog);
    parked.push(parent);
  }
  var detail =
    'a1_park n=' +
    parked.length +
    (parked.length ? ' ' + parked.slice(0, 6).join('|') : '');
  Logger.log('[仕入れ検討①] %s', detail);
  return { n: parked.length, detail: detail };
}

/** 洗い進捗に a1_park がある＝A-1 で一時停止中（続き選定から外す／帯選定では拾う）。 */
function prWashA1Parked_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return false;
  var loc = prFindMakerSumRow_(parent);
  if (!loc || !loc.sh || !(loc.sheetRow > 0) || loc.idxProg < 0) return false;
  var prog = String(loc.sh.getRange(loc.sheetRow, loc.idxProg + 1).getValue() || '');
  return prog.indexOf('a1_park') >= 0;
}

/**
 * 腐った時計・実行印・ロックを消す（A-1／番犬共通）。
 * ※実行中の洗い状態は消さない（A-1 の park は prCycleA1ParkInflight_ のみ）。
 * @return {{detail:string}}
 */
function prCycleHardResetClocks_(runId, via) {
  var rid =
    runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var before = prCycleClockSnapshot_();
  try {
    prWashInflightClear_('hard_reset:' + String(via || ''));
  } catch (e0) {}
  try {
    prShopHbBusyClear_('hard_reset');
  } catch (e0b) {}
  try {
    prWashFullUnlock_();
  } catch (e1) {}
  try {
    prClearWashFullTriggers_({
      reason: 'hard_reset:' + String(via || ''),
      runId: rid,
      log: true,
    });
  } catch (e2) {}
  try {
    prClearCycleC4Triggers_({ reason: 'hard_reset:' + String(via || ''), runId: rid });
  } catch (e3) {}
  var props = PropertiesService.getScriptProperties();
  try {
    props.deleteProperty(PR_CYCLE_C4_PENDING_PARENT_PROP_);
    props.deleteProperty(PR_CYCLE_C4_PENDING_RUNID_PROP_);
    props.deleteProperty(PR_CYCLE_C4_PENDING_TAG_PROP_);
    props.deleteProperty(PR_CYCLE_C4_TRIGGER_EXPECT_MS_PROP_);
  } catch (e4) {}
  props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
  var detail =
    'hard_reset via=' + String(via || '') + ' before=' + before + ' after=' + prCycleClockSnapshot_();
  prAppendLog_([rid, 'heartbeat liveness', 'DONE', new Date(), '', '', '', detail.slice(0, 200)]);
  return { detail: detail };
}

function prCyclePipeListGet_(prop) {
  var raw = String(PropertiesService.getScriptProperties().getProperty(prop) || '').trim();
  if (!raw) return [];
  return raw
    .split('|')
    .map(function (s) {
      return String(s).trim();
    })
    .filter(Boolean);
}

function prCyclePipeListAdd_(prop, parent) {
  parent = String(parent || '').trim();
  if (!parent) return;
  var list = prCyclePipeListGet_(prop);
  if (list.indexOf(parent) >= 0) return;
  list.push(parent);
  PropertiesService.getScriptProperties().setProperty(prop, list.join('|'));
}

function prCyclePipeListRemove_(prop, parent) {
  parent = String(parent || '').trim();
  if (!parent) return;
  var list = prCyclePipeListGet_(prop).filter(function (p) {
    return p !== parent;
  });
  var props = PropertiesService.getScriptProperties();
  if (list.length) props.setProperty(prop, list.join('|'));
  else props.deleteProperty(prop);
}

/** 当該メーカーの①候補・門通過件数（D1 と同じ読取）。 */
function prCandGatePassCountForMaker_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return 0;
  var makerSet = {};
  makerSet[parent] = true;
  var rows = prLoadCandPassByAsin_({ makerSet: makerSet });
  return Object.keys(rows || {}).length;
}

/** 完了だが①候補通過ゼロ → 再洗いへ（偽完了）。自動は1回まで。 */
function prCycleClearFalseComplete_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return { action: 'skip', parent: '' };
  prCyclePipeListRemove_(PR_CYCLE_C4_NEEDS_D1_PROP_, parent);
  var queueN = Object.keys(prWashQueueAsinSet_(parent)).length;
  var pending = prWashQueuePendingCount_(parent);
  var loc = prFindMakerSumRow_(parent);
  var prog =
    loc && loc.idxProg >= 0
      ? String(loc.sh.getRange(loc.sheetRow, loc.idxProg + 1).getValue() || '')
      : '';
  Logger.log(
    '[仕入れ検討①] c4_no_candidates parent=%s queue=%s pending=%s prog=%s',
    parent,
    queueN,
    pending,
    prog.slice(0, 80)
  );
  // 既に1回救済済み → empty_research 保留（無限再洗い防止）
  if (prog.indexOf('false_complete_rewash') >= 0 || prog.indexOf('c4_no_candidates rewash') >= 0) {
    var estCell = null;
    try {
      if (loc) {
        var hdrE = loc.sh.getRange(1, 1, 1, Math.max(loc.sh.getLastColumn(), 1)).getValues()[0];
        var idxE = prHeaderIndex_(hdrE, PR_CATALOG_EST_COL_);
        if (idxE >= 0) estCell = loc.sh.getRange(loc.sheetRow, idxE + 1).getValue();
      }
    } catch (eEst2) {}
    var est = prMakerCatalogEstimate_(parent, prog, estCell);
    var holdProg =
      'empty_research est=' +
      (est != null ? est : '?') +
      ' queue=' +
      queueN +
      ' cand=0 after_rewash';
    prSetMakerWashStatus_(parent, PR_WASH_ST_HOLD_, holdProg);
    return { action: 'empty_research', parent: parent };
  }
  if (queueN > 0 || pending > 0) {
    prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'false_complete_rewash queue=' + queueN);
    return { action: 'keepa', parent: parent };
  }
  prSetMakerWashStatus_(parent, PR_WASH_ST_ENUM_, 'false_complete_rewash');
  return { action: 'enum', parent: parent };
}

/**
 * 偽完了メーカーを列挙中／Keepa中へ戻す（手動・心拍共用）。
 * @param {{ckOnly?:boolean, quiet?:boolean, runId?:string}} opt
 * @return {{n:number, detail:string, parents:string[]}}
 */
function prCycleRewashFalseCompletes_(opt) {
  opt = opt || {};
  var ckOnly = !!opt.ckOnly;
  var quiet = !!opt.quiet;
  var runId =
    opt.runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) {
    return { n: 0, detail: 'no_sum', parents: [] };
  }
  prEnsureMakerWashStatusCols_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  if (idxM < 0 || idxSt < 0) return { n: 0, detail: 'no_cols', parents: [] };
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var parents = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxSt] || '').trim() !== PR_WASH_ST_DONE_) continue;
    if (ckOnly && idxCk >= 0 && !prCkIsTrue_(vals[r][idxCk])) continue;
    var parent = String(vals[r][idxM] || '').trim();
    if (!parent) continue;
    if (prCandGatePassCountForMaker_(parent) > 0) continue;
    var res = prCycleClearFalseComplete_(parent);
    if (res && res.parent) parents.push(parent + ':' + (res.action || ''));
  }
  var detail =
    'false_complete_rewash n=' +
    parents.length +
    (ckOnly ? ' ckOnly' : ' all') +
    (parents.length ? ' ' + parents.join('|').slice(0, 160) : '');
  Logger.log('[仕入れ検討①] runId=%s %s', runId, detail);
  if (parents.length || !quiet) {
    prAppendLog_([runId, '偽完了救済', 'DONE', new Date(), parents.length, '', '', detail.slice(0, 200)]);
  }
  if (!quiet) prMenuToast_('① 偽完了救済', detail.slice(0, 160), 12);
  return { n: parents.length, detail: detail, parents: parents };
}

/** メニュー: 偽完了を列挙中へ（一括 or レ点）。A-1 には載せない。 */
function prMenuFalseCompleteRewash() {
  prGasStart_();
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    '偽完了を列挙中に戻す',
    '①候補門通過が0の「完了」メーカーを列挙中（または Keepa中）へ戻します。\n\n' +
      'はい = 全件\nいいえ = レ点メーカーのみ\nキャンセル = 中止',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (ans === ui.Button.CANCEL) return;
  var ckOnly = ans === ui.Button.NO;
  prCycleRewashFalseCompletes_({ ckOnly: ckOnly, quiet: false });
}

/** 洗い途中の D1 転記（§6.4bis）。未設定=ON。 */
function prCycleMidD1Enabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_CYCLE_MID_D1_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

/**
 * Keepa 1スライス後: 当該メーカーの門通過を品番リストへ（完走を待たない）。
 * C4 の D1_DONE は付けない。失敗しても洗い本線は止めない。
 * @return {{ok:boolean, skip:boolean, detail:string, n?:number}}
 */
function prCycleTryMidWashD1_(opt) {
  opt = opt || {};
  var parent = String(opt.parent || '').trim();
  var runId = String(opt.runId || '').trim();
  var tag = String(opt.tag || 'mid_d1');
  if (!parent) return { ok: true, skip: true, detail: 'mid_d1 parent_empty' };
  if (!prCycleMidD1Enabled_()) return { ok: true, skip: true, detail: 'mid_d1_off' };
  if (!prListAppendEnabled_()) return { ok: true, skip: true, detail: 'list_append_off' };
  if (prGasTimeUp_()) return { ok: true, skip: true, detail: 'mid_d1_timeup' };
  var makerSet = {};
  makerSet[parent] = true;
  try {
    var r = prMenuListAppendD1({
      makerSet: makerSet,
      discoveryTag: tag + '_catalog',
      quiet: true,
      midWash: true,
      runId: runId || undefined,
    });
    var n = r && r.n != null ? Number(r.n) || 0 : 0;
    var detail =
      'mid_d1 parent=' + parent + ' ' + (r && r.detail ? String(r.detail) : 'ok');
    if (runId) {
      prAppendLog_([
        runId,
        tag + ' 途中D1',
        'DONE',
        new Date(),
        '',
        n,
        '',
        detail.slice(0, 200),
      ]);
    }
    Logger.log('[仕入れ検討①] runId=%s step=mid_d1 state=DONE %s', runId || '-', detail.slice(0, 160));
    return { ok: true, skip: false, detail: detail, n: n };
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e)).slice(0, 160);
    if (runId) {
      prAppendLog_([runId, tag + ' 途中D1', 'FAILED', new Date(), '', '', '', msg]);
    }
    Logger.log('[仕入れ検討①] runId=%s step=mid_d1 state=FAILED %s', runId || '-', msg);
    return { ok: false, skip: false, detail: msg };
  }
}

/** Keepa スライス結果から途中 D1 を呼ぶ（完走・wrote0 はスキップ）。 */
function prCycleAfterKeepaSliceMidD1_(kr, parent, runId, tag) {
  if (!kr || kr.complete || !(kr.wrote > 0)) return null;
  return prCycleTryMidWashD1_({ parent: parent, runId: runId, tag: tag });
}

/** 洗い完走済みで D1 未の parent を needs に登録。 */
function prCycleMarkC4NeedsD1_(parent) {
  parent = String(parent || '').trim();
  if (!parent || !prCycleC4Enabled_()) return;
  if (prCyclePipeListGet_(PR_CYCLE_C4_D1_DONE_PROP_).indexOf(parent) >= 0) return;
  if (prCandGatePassCountForMaker_(parent) <= 0) return;
  prCyclePipeListAdd_(PR_CYCLE_C4_NEEDS_D1_PROP_, parent);
}

/** C4 の D1 成功後、needs から外す。 */
function prCycleMarkC4D1Done_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return;
  prCyclePipeListRemove_(PR_CYCLE_C4_NEEDS_D1_PROP_, parent);
  prCyclePipeListAdd_(PR_CYCLE_C4_D1_DONE_PROP_, parent);
}

/**
 * 洗い状態=完了かつ D1 未の parent 一覧（needs prop ＋サマリ走査）。
 * @return {string[]}
 */
function prCycleListNeedsD1_() {
  var needs = prCyclePipeListGet_(PR_CYCLE_C4_NEEDS_D1_PROP_);
  var d1done = prCyclePipeListGet_(PR_CYCLE_C4_D1_DONE_PROP_);
  var merged = {};
  var i;
  for (i = 0; i < needs.length; i++) {
    if (prCandGatePassCountForMaker_(needs[i]) <= 0) {
      prCycleClearFalseComplete_(needs[i]);
    } else {
      merged[needs[i]] = true;
    }
  }
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
    if (sh && sh.getLastRow() >= 2) {
      prEnsureMakerWashStatusCols_(sh);
      var lastCol = Math.max(sh.getLastColumn(), 1);
      var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      var idxM = prHeaderIndex_(hdr, 'メーカー');
      var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
      if (idxM >= 0 && idxSt >= 0) {
        var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
        for (i = 0; i < vals.length; i++) {
          if (String(vals[i][idxSt] || '').trim() !== PR_WASH_ST_DONE_) continue;
          var p = String(vals[i][idxM] || '').trim();
          if (!p) continue;
          if (prCandGatePassCountForMaker_(p) <= 0) {
            prCycleClearFalseComplete_(p);
            continue;
          }
          merged[p] = true;
        }
      }
    }
  } catch (eScan) {
    Logger.log(
      '[仕入れ検討①] WARN c4_needs_scan %s',
      String((eScan && eScan.message) || eScan).slice(0, 120)
    );
  }
  var out = [];
  for (var k in merged) {
    if (Object.prototype.hasOwnProperty.call(merged, k) && d1done.indexOf(k) < 0) out.push(k);
  }
  return out;
}

function prTitleZeroMap_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PR_WASH_TITLE0_FAIL_PROP_);
    if (!raw) return {};
    return JSON.parse(raw) || {};
  } catch (e0) {
    return {};
  }
}

function prTitleZeroSave_(map) {
  PropertiesService.getScriptProperties().setProperty(
    PR_WASH_TITLE0_FAIL_PROP_,
    JSON.stringify(map || {})
  );
}

function prTitleZeroBump_(parent) {
  parent = String(parent || '').trim();
  if (prTitleZeroSkipSetHas_(parent)) return PR_WASH_TITLE0_SKIP_AFTER_;
  var map = prTitleZeroMap_();
  map[parent] = (Number(map[parent]) || 0) + 1;
  prTitleZeroSave_(map);
  return map[parent];
}

function prTitleZeroClear_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return;
  var map = prTitleZeroMap_();
  if (!map[parent]) return;
  delete map[parent];
  prTitleZeroSave_(map);
}

function prTitleZeroSkipSetHas_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return false;
  return prCyclePipeListGet_(PR_WASH_TITLE0_SKIP_SET_PROP_).indexOf(parent) >= 0;
}

function prTitleZeroSkipSetAdd_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return;
  prCyclePipeListAdd_(PR_WASH_TITLE0_SKIP_SET_PROP_, parent);
}

/** 保留(title0_skip)または永久 skip 集合に入っている。 */
function prWashTitleZeroBlocked_(parent, st, prog) {
  parent = String(parent || '').trim();
  if (!parent) return false;
  if (prTitleZeroSkipSetHas_(parent)) return true;
  st = st != null ? String(st).trim() : '';
  prog = prog != null ? String(prog) : '';
  if (st === PR_WASH_ST_HOLD_ && prog.indexOf('title0_skip') >= 0) return true;
  return false;
}

function prWashTitleZeroBlockedForParent_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return false;
  if (prTitleZeroSkipSetHas_(parent)) return true;
  var loc = prFindMakerSumRow_(parent);
  if (!loc || loc.idxSt < 0) return false;
  var st = String(loc.sh.getRange(loc.sheetRow, loc.idxSt + 1).getValue() || '').trim();
  var prog =
    loc.idxProg >= 0
      ? String(loc.sh.getRange(loc.sheetRow, loc.idxProg + 1).getValue() || '')
      : '';
  return prWashTitleZeroBlocked_(parent, st, prog);
}

/**
 * タイトル含有0の連打対策。N 回目で保留＋洗い予約解除＋心拍へ。
 * @return {{skipped:boolean, count:number, detail:string}}
 */
function prHandleWashTitleZero_(opt) {
  opt = opt || {};
  var parent = String(opt.parent || '').trim();
  var runId = String(opt.runId || '');
  var tag = String(opt.tag || 'wash_full');
  var armTrigger = !!opt.armTrigger;
  var quiet = !!opt.quiet;
  var n = prTitleZeroBump_(parent);
  if (prTitleZeroSkipSetHas_(parent)) {
    prSetMakerWashStatus_(parent, PR_WASH_ST_HOLD_, 'title0_skip permanent');
    return {
      skipped: true,
      count: n,
      detail: 'title0_skip permanent parent=' + parent,
      wash_armed: false,
      heartbeat_armed: false,
      status: PR_WASH_ST_HOLD_,
    };
  }
  if (n < PR_WASH_TITLE0_SKIP_AFTER_) {
    if (armTrigger) {
      PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      prScheduleWashFullTrigger_(prWashFullTriggerAfterMin_(5, ''), {
        runId: runId,
        parent: parent,
        tag: tag,
      });
    }
    return {
      skipped: false,
      count: n,
      detail: 'title0_retry n=' + n + ' parent=' + parent,
    };
  }
  prTitleZeroSkipSetAdd_(parent);
  var mapTz = prTitleZeroMap_();
  mapTz[parent] = PR_WASH_TITLE0_SKIP_AFTER_;
  prTitleZeroSave_(mapTz);
  prSetMakerWashStatus_(parent, PR_WASH_ST_HOLD_, 'title0_skip n=' + n);
  prClearWashFullTriggers_({ reason: 'title0_skip:' + parent, runId: runId, log: true });
  PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'false');
  var detail = 'title0_skip parent=' + parent + ' n=' + n;
  try {
    var hint = prTryAddBrandHintMaker_(parent);
    if (hint && hint.brand) {
      detail += ' brand_hint=' + hint.brand + ' ' + (hint.detail || '');
    } else if (hint && hint.detail) {
      detail += ' brand_hint_skip=' + String(hint.detail).slice(0, 80);
    }
  } catch (eHint) {
    detail += ' brand_hint_err=' + String((eHint && eHint.message) || eHint).slice(0, 60);
  }
  if (prCycleHeartbeatEnabled_()) {
    var hb = prCycleArmHeartbeatAfterIdle_(runId, 'title0_skip');
    detail += ' heartbeat_armed afterMin=' + hb;
  }
  prAppendLog_([runId, tag + ' 洗い完走', 'DONE', new Date(), n, parent, '', detail.slice(0, 200)]);
  if (!quiet) prMenuToast_('① 洗い完走', detail.slice(0, 160), 12);
  return { skipped: true, count: n, detail: detail };
}

/**
 * 完走済みで D1 未のメーカーへ C4 を別実行で予約（A-1／番犬／手動 recover）。
 * D1_DONE 済み・候補0・完了以外・既存 C4 予約ありはスキップ（二重作業を避ける）。
 * @return {string} arm した parent。しなければ ''
 */
function prCycleArmC4Recover_(runId, tag) {
  if (!prCycleC4Enabled_()) return '';
  var props = PropertiesService.getScriptProperties();
  if (String(props.getProperty(PR_CYCLE_C4_PENDING_PARENT_PROP_) || '').trim()) return '';
  if (prCountCycleC4Triggers_() > 0) return '';
  if (prCycleD2OverflowActive_()) return '';
  var d1done = prCyclePipeListGet_(PR_CYCLE_C4_D1_DONE_PROP_);
  var needs = prCycleListNeedsD1_();
  if (!needs.length) return '';
  var parent = '';
  var withCand = [];
  var j;
  for (j = 0; j < needs.length; j++) {
    var candP = needs[j];
    if (d1done.indexOf(candP) >= 0) continue;
    if (prCandGatePassCountForMaker_(candP) > 0) withCand.push(candP);
  }
  if (!withCand.length) return '';
  var last = String(props.getProperty(PR_CYCLE_C4_LAST_PARENT_PROP_) || '').trim();
  if (last && withCand.indexOf(last) >= 0) parent = last;
  if (!parent) parent = withCand[0];
  if (d1done.indexOf(parent) >= 0) return '';
  var loc = prFindMakerSumRow_(parent);
  if (!loc || !loc.sh || !(loc.sheetRow > 0) || loc.idxSt < 0) return '';
  var st = String(loc.sh.getRange(loc.sheetRow, loc.idxSt + 1).getValue() || '').trim();
  if (st !== PR_WASH_ST_DONE_) return '';
  prScheduleCycleC4Trigger_(1, {
    runId: runId,
    parent: parent,
    tag: tag || 'c4_recover',
  });
  return parent;
}

/**
 * 直近完走メーカーで C4 未予約なら、リスト転記を別実行で予約。
 * @return {string} arm した parent。しなければ ''
 */
function prCycleMaybeArmC4ForLastComplete_(runId, tag) {
  return prCycleArmC4Recover_(runId, tag);
}

/**
 * 生きているか。
 * 稼働中: 洗い/C4 の「これから来る」時計、または実行中印が新しい。
 * 予定時刻が Keepa最大回復＋α 以上過ぎている／実行中印がその長さ腐っている → 沈黙（リセット対象）。
 */
function prCycleAliveReason_() {
  if (!prCycleHeartbeatEnabled_()) return { alive: true, reason: 'hb_off' };
  if (prCycleHeartbeatPaused_()) {
    var ok01 = prCountResume01Triggers_() > 0;
    return { alive: ok01, reason: ok01 ? 'paused_01' : 'paused_no_01' };
  }
  if (!prMakerWashEnabled_() || !prMakerWashFullEnabled_()) {
    return { alive: true, reason: 'wash_engine_off' };
  }
  var now = new Date().getTime();
  var staleLimit = prCycleStaleLimitMs_();
  var expectMs = prCycleNextExpectedFireMs_();
  var washN = prCountWashFullTriggers_();
  var c4N = prCountCycleC4Triggers_();
  var owner = prWashInflightOwner_();
  var age = prWashInflightAgeMs_();

  if ((washN > 0 || c4N > 0) && expectMs > 0 && now - expectMs >= staleLimit) {
    return {
      alive: false,
      reason: 'expect_overdue',
      overdueMs: now - expectMs,
      staleLimitMin: prCycleStaleLimitMin_(),
    };
  }
  var inflightLimit = prCycleInflightStaleLimitMs_();
  if (owner && age >= inflightLimit) {
    return {
      alive: false,
      reason: 'inflight_overdue',
      overdueMs: age,
      staleLimitMin: Math.round(inflightLimit / 60000),
      phase: String(
        PropertiesService.getScriptProperties().getProperty(PR_WASH_INFLIGHT_PHASE_PROP_) || ''
      ),
    };
  }
  if (washN > 0) return { alive: true, reason: 'wash_trigger' };
  if (c4N > 0) return { alive: true, reason: 'c4_trigger' };
  if (owner && age > 0 && age < PR_CYCLE_LIVENESS_SILENCE_MS_) {
    return { alive: true, reason: 'inflight_fresh' };
  }
  return { alive: false, reason: 'silence' };
}

/** 15分周期の生存番犬。無ければ張る（日次メールと同型）。 */
function prCycleArmLivenessTrigger_() {
  if (!prCycleHeartbeatEnabled_()) return false;
  var list = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === PR_CYCLE_LIVENESS_FN_) return false;
  }
  ScriptApp.newTrigger(PR_CYCLE_LIVENESS_FN_)
    .timeBased()
    .everyMinutes(PR_CYCLE_LIVENESS_EVERY_MIN_)
    .create();
  Logger.log('[仕入れ検討①] liveness trigger armed everyMin=%s', PR_CYCLE_LIVENESS_EVERY_MIN_);
  return true;
}

/**
 * 沈黙／予定超過を検知してリセット＋再武装。
 * @return {{alive:boolean, rearmed:boolean, reason:string, clocks:string}}
 */
function prCycleEnsureLiveness_(runId, via) {
  prCycleArmLivenessTrigger_();
  var snap = prCycleAliveReason_();
  var clocks = prCycleClockSnapshot_();
  if (snap.alive) {
    return { alive: true, rearmed: false, reason: snap.reason, clocks: clocks };
  }
  var rid =
    runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  if (prCycleHeartbeatPaused_()) {
    prCycleArmResume01_();
    prAppendLog_([
      rid,
      'heartbeat liveness',
      'DONE',
      new Date(),
      '',
      '',
      '',
      'liveness_rearm resume01 via=' + String(via || '') + ' ' + clocks,
    ]);
    return { alive: true, rearmed: true, reason: 'paused_rearm_01', clocks: clocks };
  }
  prCycleHardResetClocks_(rid, String(via || '') + ':' + snap.reason);
  var c4Parent = prCycleMaybeArmC4ForLastComplete_(rid, 'liveness_c4_recover');
  var afterMin = prCycleArmHeartbeatAfterIdle_(rid, 'liveness_reset');
  prAppendLog_([
    rid,
    'heartbeat liveness',
    'DONE',
    new Date(),
    afterMin,
    c4Parent || '',
    '',
    'liveness_reset reason=' +
      snap.reason +
      ' via=' +
      String(via || '') +
      ' afterMin=' +
      afterMin +
      (c4Parent ? ' c4=' + c4Parent : '') +
      ' staleLimitMin=' +
      (snap.staleLimitMin != null ? snap.staleLimitMin : prCycleStaleLimitMin_()) +
      (snap.phase ? ' deadPhase=' + snap.phase : '') +
      (snap.overdueMs != null ? ' overdueMin=' + Math.round(snap.overdueMs / 60000) : '') +
      ' ' +
      clocks,
  ]);
  return { alive: false, rearmed: true, reason: snap.reason, clocks: clocks };
}

/** 時間主導: 15分ごとに生存確認。沈黙／予定超過ならリセットして再開。 */
function prRunCycleLivenessFromTrigger() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var before = prCycleClockSnapshot_();
  var r = prCycleEnsureLiveness_(runId, 'clock15');
  prAppendLog_([
    runId,
    'heartbeat liveness',
    'DONE',
    new Date(),
    r.rearmed ? 1 : 0,
    '',
    '',
    'tick via=clock15 reason=' +
      r.reason +
      ' rearmed=' +
      (r.rearmed ? '1' : '0') +
      ' before=' +
      before,
  ]);
}

/** 01:00 再開。手動再開済み・他実行が無ければ再開。 */
function prRunCycleResumeAt01() {
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  if (!prCycleHeartbeatPaused_()) {
    prCycleClearResume01_();
    prAppendLog_([runId, '01:00再開', 'DONE', new Date(), '', '', '', 'noop_not_paused']);
    return;
  }
  if (prWashLockLeftSec_() > 0) {
    ScriptApp.newTrigger(PR_CYCLE_RESUME_01_FN_).timeBased().after(15 * 60 * 1000).create();
    prAppendLog_([runId, '01:00再開', 'DONE', new Date(), 15, '', '', 'defer_busy lock']);
    return;
  }
  prCycleResumeManual_({ runId: runId });
  prAppendLog_([runId, '01:00再開', 'DONE', new Date(), '', '', '', 'resume']);
}

function prCycleDailyMailEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_CYCLE_DAILY_MAIL_ENABLED_PROP_);
  if (raw == null || raw === '') return true;
  return String(raw).toLowerCase() !== 'false';
}

/** 日次生存メール（社内）。問屋宛ではない。停止中・完走0でも送る。 */
function prRunCycleDailyMail() {
  if (!prCycleDailyMailEnabled_()) return;
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var day = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var paused = prCycleHeartbeatPaused_();
  var stats = prCycleDailyMailStats_();
  var body = prBuildCycleDailyMailBody_({
    day: day,
    paused: paused,
    completedMakers: stats.completedMakers,
    topCategories: stats.topCategories,
    gatePassN: stats.gatePassN,
    enumN: stats.enumN,
    alightShops: stats.alightShops,
    c0New: stats.c0New,
    smPool: stats.smPool,
  });
  var to = String(
    PropertiesService.getScriptProperties().getProperty(PR_CYCLE_DAILY_MAIL_TO_PROP_) || Session.getActiveUser().getEmail() || ''
  ).trim();
  if (!to) {
    prAppendLog_([runId, '日次メール', 'FAILED', new Date(), '', '', '', 'no_to']);
    return;
  }
  try {
    MailApp.sendEmail(to, '[①] 自動リサーチ日次 ' + day + (paused ? '（停止中）' : ''), body);
    prAppendLog_([
      runId,
      '日次メール',
      'DONE',
      new Date(),
      stats.gatePassN,
      stats.completedMakers.length,
      '',
      'sent=' + to + ' sm=' + stats.smPool + ' paused=' + (paused ? 'Y' : 'N'),
    ]);
  } catch (e) {
    prAppendLog_([
      runId,
      '日次メール',
      'FAILED',
      new Date(),
      '',
      '',
      '',
      String((e && e.message) || e).slice(0, 160),
    ]);
  }
}

/**
 * ①候補の転記チェック列を ASIN 単位で更新（通過行のみ）。
 * @param {Object} asinSet ASIN→true
 * @param {string} value 未|済|漏れ|空
 */
function prCandMarkListCheck_(asinSet, value) {
  if (!prCycleListCheckEnabled_()) return 0;
  if (!asinSet) return 0;
  var keys = Object.keys(asinSet);
  if (!keys.length) return 0;
  var sh = prEnsureSheet_(PR_CAND_SHEET_, PR_CAND_HEADERS_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxAsin = prHeaderIndex_(hdr, 'ASIN');
  var idxGate = prHeaderIndex_(hdr, '門結果');
  var idxCk = prHeaderIndex_(hdr, PR_CAND_LIST_CHECK_);
  if (idxAsin < 0 || idxCk < 0) return 0;
  if (sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var n = 0;
  var i;
  for (i = 0; i < vals.length; i++) {
    var a = String(vals[i][idxAsin] || '')
      .trim()
      .toUpperCase();
    if (!a || !asinSet[a]) continue;
    if (idxGate >= 0 && String(vals[i][idxGate] || '') !== '通過') continue;
    var cur = String(vals[i][idxCk] || '');
    if (cur === value) continue;
    sh.getRange(i + 2, idxCk + 1).setValue(value);
    n++;
  }
  if (n) PR_CAND_MEM_ = null;
  return n;
}

function prParseCandFetchedAt_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var s = String(v || '').trim();
  if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function prListAsinSetFromSheet_() {
  var out = {};
  var sh = prSpreadsheet_().getSheetByName(PR_LIST_SHEET_);
  if (!sh || sh.getLastRow() < PR_LIST_DATA_START_) return out;
  var headers = prListHeaders_(sh);
  var asinI = headers.indexOf('ASIN');
  if (asinI < 0) return out;
  var vals = sh
    .getRange(PR_LIST_DATA_START_, asinI + 1, sh.getLastRow(), asinI + 1)
    .getValues();
  var i;
  for (i = 0; i < vals.length; i++) {
    var a = String(vals[i][0] || '')
      .trim()
      .toUpperCase();
    if (a && a.charAt(0) === 'B') out[a] = true;
  }
  return out;
}

/**
 * 週次転記チェック（CYCLE §6.4ter）。通過×直近N日＋既存漏れ → リスト照合 → 漏れは D1 → 社内メール。
 * トリガー／メニュー両用。
 */
function prRunListTransferWeeklyCheck(opt) {
  opt = opt || {};
  if (!prCycleListCheckEnabled_()) {
    Logger.log('[仕入れ検討①] list_check skipped enabled=false');
    return { ok: false, detail: 'disabled' };
  }
  prGasStart_();
  var runId =
    (opt.runId && String(opt.runId).trim()) ||
    'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var days = prCycleListCheckDays_();
  var sinceMs = Date.now() - days * 24 * 60 * 60 * 1000;
  var via = opt.via || 'trigger';
  Logger.log('[仕入れ検討①] runId=%s step=list_check state=RUNNING via=%s days=%s', runId, via, days);
  try {
    prRefuseOriginalSs_();
    prEnsureSheet_(PR_CAND_SHEET_, PR_CAND_HEADERS_);
    var listSet = prListAsinSetFromSheet_();
    var sh = prSpreadsheet_().getSheetByName(PR_CAND_SHEET_);
    if (!sh || sh.getLastRow() < 2) {
      prAppendLog_([runId, '週次転記チェック', 'DONE', new Date(), '', '', '', 'no_cand via=' + via]);
      prSendListTransferWeeklyMail_({
        runId: runId,
        via: via,
        days: days,
        scanned: 0,
        okN: 0,
        leakN: 0,
        d1N: 0,
        detail: '①候補なし',
      });
      return { ok: true, detail: 'no_cand' };
    }
    var lastCol = sh.getLastColumn();
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxAsin = prHeaderIndex_(hdr, 'ASIN');
    var idxGate = prHeaderIndex_(hdr, '門結果');
    var idxAt = prHeaderIndex_(hdr, '取得日時');
    var idxCk = prHeaderIndex_(hdr, PR_CAND_LIST_CHECK_);
    if (idxAsin < 0 || idxGate < 0 || idxCk < 0) {
      throw new Error('①候補に ASIN／門結果／転記チェック列がありません');
    }
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var scanned = 0;
    var okN = 0;
    var leakN = 0;
    var leakSet = {};
    var okSet = {};
    var i;
    for (i = 0; i < vals.length; i++) {
      if (String(vals[i][idxGate] || '') !== '通過') continue;
      var asin = String(vals[i][idxAsin] || '')
        .trim()
        .toUpperCase();
      if (!asin) continue;
      var ck = String(vals[i][idxCk] || '').trim();
      var fetched = idxAt >= 0 ? prParseCandFetchedAt_(vals[i][idxAt]) : null;
      var inWindow = !!(fetched && fetched.getTime() >= sinceMs);
      var isLeakMark = ck === PR_CAND_LIST_CHECK_MORE_;
      var isPending = !ck || ck === PR_CAND_LIST_CHECK_MI_;
      var isSumi = ck === PR_CAND_LIST_CHECK_SUMI_;
      // 対象: 漏れ（常時）／列未記入の通過（初回埋め）／未（窓内 or 日時欠落）／済の再確認（窓内）
      var pick = false;
      if (isLeakMark) pick = true;
      else if (!ck) pick = true;
      else if (isPending && (inWindow || !fetched)) pick = true;
      else if (isSumi && inWindow) pick = true;
      if (!pick) continue;
      scanned++;
      if (listSet[asin]) {
        okN++;
        okSet[asin] = true;
      } else {
        leakN++;
        leakSet[asin] = true;
      }
    }
    prCandMarkListCheck_(okSet, PR_CAND_LIST_CHECK_SUMI_);
    prCandMarkListCheck_(leakSet, PR_CAND_LIST_CHECK_MORE_);
    var d1N = 0;
    var d1Detail = '';
    if (Object.keys(leakSet).length) {
      var d1 = prMenuListAppendD1({
        runId: runId,
        asinSet: leakSet,
        quiet: true,
        midWash: true,
      });
      d1N = d1 && d1.n != null ? Number(d1.n) : 0;
      d1Detail = (d1 && d1.detail) || '';
      // D1 成功分は prMenuListAppendD1 内で済。残漏れは再照合
      var still = {};
      var ks = Object.keys(leakSet);
      var list2 = prListAsinSetFromSheet_();
      for (i = 0; i < ks.length; i++) {
        if (!list2[ks[i]]) still[ks[i]] = true;
      }
      prCandMarkListCheck_(still, PR_CAND_LIST_CHECK_MORE_);
      leakN = Object.keys(still).length;
    }
    var detail =
      'via=' +
      via +
      ' days=' +
      days +
      ' scanned=' +
      scanned +
      ' ok=' +
      okN +
      ' leak=' +
      leakN +
      ' d1n=' +
      d1N +
      (d1Detail ? ' ' + d1Detail : '');
    prAppendLog_([runId, '週次転記チェック', 'DONE', new Date(), leakN, d1N, '', detail.slice(0, 220)]);
    prSendListTransferWeeklyMail_({
      runId: runId,
      via: via,
      days: days,
      scanned: scanned,
      okN: okN,
      leakN: leakN,
      d1N: d1N,
      detail: detail,
    });
    Logger.log('[仕入れ検討①] runId=%s step=list_check state=DONE %s', runId, detail);
    return { ok: true, detail: detail, leakN: leakN, d1N: d1N };
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    prAppendLog_([runId, '週次転記チェック', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    try {
      prSendListTransferWeeklyMail_({
        runId: runId,
        via: via,
        days: days,
        scanned: 0,
        okN: 0,
        leakN: 0,
        d1N: 0,
        detail: 'FAILED ' + msg,
        failed: true,
      });
    } catch (e2) {}
    throw e;
  }
}

function prSendListTransferWeeklyMail_(opt) {
  opt = opt || {};
  var to = String(
    PropertiesService.getScriptProperties().getProperty(PR_CYCLE_DAILY_MAIL_TO_PROP_) ||
      Session.getActiveUser().getEmail() ||
      ''
  ).trim();
  if (!to) {
    prAppendLog_([opt.runId || '', '週次転記メール', 'FAILED', new Date(), '', '', '', 'no_to']);
    return;
  }
  var day = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var subject =
    '[①] 転記チェック週次 ' + day + (opt.failed ? '（失敗）' : opt.leakN ? '（漏れあり）' : '');
  var body =
    '転記チェック（週次）\n' +
    '日付: ' +
    day +
    '\n' +
    '実行: ' +
    String(opt.via || '') +
    '\n' +
    '窓: 直近' +
    Number(opt.days || 0) +
    '日（取得日時）＋既存「漏れ」\n' +
    '照合した通過: ' +
    Number(opt.scanned || 0) +
    '\n' +
    'リストあり→済: ' +
    Number(opt.okN || 0) +
    '\n' +
    '漏れ（D1後の残）: ' +
    Number(opt.leakN || 0) +
    '\n' +
    'D1書込件数: ' +
    Number(opt.d1N || 0) +
    '\n' +
    '詳細: ' +
    String(opt.detail || '') +
    '\n';
  try {
    MailApp.sendEmail(to, subject, body);
    prAppendLog_([
      opt.runId || '',
      '週次転記メール',
      'DONE',
      new Date(),
      opt.leakN,
      opt.d1N,
      '',
      'sent=' + to,
    ]);
  } catch (e) {
    prAppendLog_([
      opt.runId || '',
      '週次転記メール',
      'FAILED',
      new Date(),
      '',
      '',
      '',
      String((e && e.message) || e).slice(0, 160),
    ]);
  }
}

function prMenuListTransferWeeklyCheckNow() {
  prGasStart_();
  var r = prRunListTransferWeeklyCheck({ via: 'menu99' });
  prMenuToast_('転記チェック', (r && r.detail) || '完了', 12);
}

function prBuildCycleDailyMailBody_(opt) {
  opt = opt || {};
  var completed = opt.completedMakers || [];
  var cats = opt.topCategories || [];
  return (
    '自動リサーチ日次（生存確認）\n' +
    '日付: ' +
    String(opt.day || '') +
    '\n' +
    '心拍: ' +
    (opt.paused ? '一時停止' : '稼働または待機') +
    '\n' +
    '完走メーカー(' +
    completed.length +
    '): ' +
    (completed.length ? completed.join('、') : '（なし）') +
    '\n' +
    '主要カテゴリ: ' +
    (cats.length ? cats.slice(0, 8).join('、') : '不明') +
    '\n' +
    '門通過数: ' +
    Number(opt.gatePassN || 0) +
    '\n' +
    '参考: 列挙=' +
    Number(opt.enumN || 0) +
    ' A-light店=' +
    Number(opt.alightShops || 0) +
    ' C0新規=' +
    Number(opt.c0New || 0) +
    ' 小中プール=' +
    Number(opt.smPool || 0) +
    '\n'
  );
}

/** 日次メール用の軽い集計（失敗しても空で続行）。 */
function prCycleDailyMailStats_() {
  var out = {
    completedMakers: [],
    topCategories: [],
    gatePassN: 0,
    enumN: 0,
    alightShops: 0,
    c0New: 0,
    smPool: 0,
  };
  try {
    out.gatePassN = prCandPassAsinList_().length;
  } catch (e0) {}
  try {
    out.alightShops = prSellerUncheckedCount_();
  } catch (e1) {}
  try {
    out.smPool = prCountWashableSmallMed_();
  } catch (e2) {}
  try {
    var day = prYmdTokyo_();
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
    if (sh && sh.getLastRow() >= 2) {
      var lastCol = Math.max(sh.getLastColumn(), 1);
      var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      var idxM = prHeaderIndex_(hdr, 'メーカー');
      var idxDay = prHeaderIndex_(hdr, '最終調査日');
      var idxBand = prHeaderIndex_(hdr, PR_BAND_COL_);
      var idxCat = prHeaderIndex_(hdr, 'amazon_browse_shelves');
      if (idxCat < 0) idxCat = prHeaderIndex_(hdr, '主要カテゴリ');
      var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
      var catCount = {};
      var r;
      for (r = 0; r < vals.length; r++) {
        var maker = idxM >= 0 ? String(vals[r][idxM] || '').trim() : '';
        if (!maker) continue;
        var band = idxBand >= 0 ? String(vals[r][idxBand] || '').trim() : '';
        if (band === PR_BAND_UNK_) out.enumN++;
        var last = idxDay >= 0 ? String(vals[r][idxDay] || '').trim() : '';
        if (last.indexOf(day) === 0 || last === day) out.completedMakers.push(maker);
        if (idxCat >= 0) {
          var c = String(vals[r][idxCat] || '').trim();
          if (c) catCount[c] = (catCount[c] || 0) + 1;
        }
      }
      out.topCategories = Object.keys(catCount).sort(function (a, b) {
        return catCount[b] - catCount[a];
      });
    }
  } catch (e3) {}
  try {
    var props = PropertiesService.getScriptProperties();
    out.c0New = Number(props.getProperty('PR_CYCLE_C0_NEW_TODAY') || 0) || 0;
  } catch (e4) {}
  return out;
}

/** メーカーが食品0%店のみ由来なら true（C3 選定除外）。 */
function prMakerFoodZeroOnly_(maker, sellerVals, sellerHdr) {
  maker = String(maker || '').trim();
  if (!maker) return false;
  var idxExt = prHeaderIndex_(sellerHdr, '抽出メーカー');
  var idxFood = prHeaderIndex_(sellerHdr, '食品比率');
  if (idxExt < 0) return false;
  var seen = false;
  var r;
  for (r = 0; r < sellerVals.length; r++) {
    var cell = String(sellerVals[r][idxExt] || '');
    var found = false;
    prSplitPipeValues_(cell).forEach(function (p) {
      if (p === maker) found = true;
    });
    if (!found) continue;
    seen = true;
    var ratio = 0;
    if (idxFood >= 0) {
      var raw = String(sellerVals[r][idxFood] || '').replace('%', '').trim();
      ratio = Number(raw) || 0;
    }
    if (ratio > 0) return false;
  }
  return seen;
}

/**
 * C3 自動選定（CYCLE §6.3 規模帯）。洗い途中があれば続きのみ。不明は選ばない。
 * @return {Array<{parent:string,sheetRow:number,rowIdx:number,jitta:string,pref:string,hit:string,via:string,tier:string,band:string}>}
 */
function prPickCycleMakerSelectRows_(vals, hdr, cap) {
  cap = cap || 1;
  var cont = prPickWashContinueTarget_();
  if (cont && cont.parent) {
    var idxMaker0 = prHeaderIndex_(hdr, 'メーカー');
    var idxHit0 = prHeaderIndex_(hdr, PR_MAKER_SUM_HIT_COL_);
    var idxPref0 = prHeaderIndex_(hdr, '入口推奨');
    var idxJ0 = prHeaderIndex_(hdr, '実態');
    var r0;
    for (r0 = 0; r0 < vals.length; r0++) {
      if (String(vals[r0][idxMaker0] || '').trim() !== cont.parent) continue;
      return [
        {
          parent: cont.parent,
          sheetRow: r0 + 2,
          rowIdx: r0,
          jitta: idxJ0 >= 0 ? String(vals[r0][idxJ0] || '').trim() : '',
          pref: idxPref0 >= 0 ? String(vals[r0][idxPref0] || '').trim() : cont.pref || '',
          hit: idxHit0 >= 0 ? String(vals[r0][idxHit0] || '').trim() : cont.hit || '',
          via: 'wash_continue:' + (cont.via || ''),
          tier: '-',
          band: '-',
        },
      ];
    }
    return [
      {
        parent: cont.parent,
        sheetRow: 0,
        rowIdx: -1,
        jitta: '',
        pref: cont.pref || '',
        hit: cont.hit || '',
        via: 'wash_continue:' + (cont.via || ''),
        tier: '-',
        band: '-',
      },
    ];
  }

  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  var idxJ = prHeaderIndex_(hdr, '実態');
  var idxPref = prHeaderIndex_(hdr, '入口推奨');
  var idxHit = prHeaderIndex_(hdr, PR_MAKER_SUM_HIT_COL_);
  var idxDay = prHeaderIndex_(hdr, '最終調査日');
  var idxMemo = prHeaderIndex_(hdr, 'メモ');
  var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
  var idxEst = prHeaderIndex_(hdr, PR_CATALOG_EST_COL_);
  var idxBand = prHeaderIndex_(hdr, PR_BAND_COL_);
  var idxStC3 = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
  if (idxMaker < 0) return [];

  var sellerVals = [];
  var sellerHdr = [];
  try {
    var sSh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
    if (sSh && sSh.getLastRow() >= 2) {
      var sLastCol = Math.max(sSh.getLastColumn(), 1);
      sellerHdr = sSh.getRange(1, 1, 1, sLastCol).getValues()[0];
      sellerVals = sSh.getRange(2, 1, sSh.getLastRow() - 1, sLastCol).getValues();
    }
  } catch (eS) {
    Logger.log('[仕入れ検討①] C3 seller読取スキップ %s', String((eS && eS.message) || eS).slice(0, 80));
  }

  var cutoffMs = Date.now() - prAutoMakerRewalkDays_() * 24 * 60 * 60 * 1000;
  var byBand = { 小: [], 中: [], 大: [], 特大: [] };
  var seen = {};
  var r;
  for (r = 0; r < vals.length; r++) {
    var parent = String(vals[r][idxMaker] || '').trim();
    if (!parent || seen[parent]) continue;
    seen[parent] = true;
    if (PR_HIT_EXCLUDE_SET_[parent] || prJittaExcludeSeed_(parent)) continue;
    if (idxCk >= 0 && prCkIsTrue_(vals[r][idxCk])) continue;
    var jitta = idxJ >= 0 ? String(vals[r][idxJ] || '').trim() : '';
    if (!prMakerJittaWashEligible_(jitta)) continue;
    var lastMs = prCellToMsTokyo_(idxDay >= 0 ? vals[r][idxDay] : '');
    var never = lastMs == null;
    if (!never && lastMs > cutoffMs) continue;
    if (sellerVals.length && prMakerFoodZeroOnly_(parent, sellerVals, sellerHdr)) continue;
    var prog = idxProg >= 0 ? String(vals[r][idxProg] || '') : '';
    var stC3 = idxStC3 >= 0 ? String(vals[r][idxStC3] || '').trim() : '';
    if (prWashTitleZeroBlocked_(parent, stC3, prog)) continue;
    if (prWashHoldBlockedForSelect_(stC3, prog)) continue;
    var estCell = idxEst >= 0 ? vals[r][idxEst] : '';
    var nEstC3 = Number(String(estCell == null ? '' : estCell).replace(/[^\d.-]/g, ''));
    if (!( !isNaN(nEstC3) && nEstC3 > 0 && String(estCell).trim() !== '' )) continue;
    var est = prMakerCatalogEstimate_(parent, prog, estCell);
    var band =
      idxBand >= 0 && String(vals[r][idxBand] || '').trim()
        ? String(vals[r][idxBand] || '').trim()
        : prMakerCatalogBand_(est);
    if (band === PR_BAND_UNK_ || !byBand[band]) continue;
    var hitN = idxHit >= 0 ? Number(vals[r][idxHit]) || 0 : 0;
    var tier = prMakerCatalogTier_(est);
    byBand[band].push({
      sheetRow: r + 2,
      rowIdx: r,
      parent: parent,
      jitta: jitta,
      pref: idxPref >= 0 ? String(vals[r][idxPref] || '').trim() : '',
      hit: hitN,
      hitOk: hitN >= 2,
      memo: idxMemo >= 0 ? vals[r][idxMemo] : '',
      tier: tier,
      band: band,
      never: never,
      lastMs: never ? 0 : lastMs,
      via: 'c3_band' + band + '_hit' + hitN,
    });
  }

  function sortBand(a, b) {
    if (a.hitOk !== b.hitOk) return a.hitOk ? -1 : 1;
    if (b.hit !== a.hit) return b.hit - a.hit;
    if (a.never !== b.never) return a.never ? -1 : 1;
    if (a.never && b.never) return a.rowIdx - b.rowIdx;
    if (a.lastMs !== b.lastMs) return a.lastMs - b.lastMs;
    return a.rowIdx - b.rowIdx;
  }

  if (!prCycleBandSelectEnabled_()) {
    var flat = byBand['小'].concat(byBand['中'], byBand['大'], byBand['特大']);
    flat.sort(function (a, b) {
      var ta = a.tier === 'A' ? 0 : 1;
      var tb = b.tier === 'A' ? 0 : 1;
      if (ta !== tb) return ta - tb;
      if (b.hit !== a.hit) return b.hit - a.hit;
      if (a.never !== b.never) return a.never ? -1 : 1;
      if (a.never && b.never) return a.rowIdx - b.rowIdx;
      if (a.lastMs !== b.lastMs) return a.lastMs - b.lastMs;
      return a.rowIdx - b.rowIdx;
    });
    return flat.slice(0, cap);
  }

  ['小', '中', '大', '特大'].forEach(function (b) {
    byBand[b].sort(sortBand);
  });
  var ordered;
  if (byBand['小'].length || byBand['中'].length) {
    ordered = byBand['小'].concat(byBand['中']);
  } else if (byBand['大'].length) {
    ordered = byBand['大'];
  } else {
    ordered = byBand['特大'];
  }
  return ordered.slice(0, cap);
}

/**
 * C3 と同じ適格条件で規模帯プールを返す（cap なし）。
 * @return {{byBand:Object, unk:Array}}
 */
function prCycleMakerBandPools_() {
  var empty = { byBand: { 小: [], 中: [], 大: [], 特大: [] }, unk: [] };
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return empty;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  var idxJ = prHeaderIndex_(hdr, '実態');
  var idxPref = prHeaderIndex_(hdr, '入口推奨');
  var idxHit = prHeaderIndex_(hdr, PR_MAKER_SUM_HIT_COL_);
  var idxDay = prHeaderIndex_(hdr, '最終調査日');
  var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
  var idxEst = prHeaderIndex_(hdr, PR_CATALOG_EST_COL_);
  var idxBand = prHeaderIndex_(hdr, PR_BAND_COL_);
  var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
  if (idxMaker < 0) return empty;
  var sellerVals = [];
  var sellerHdr = [];
  try {
    var sSh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
    if (sSh && sSh.getLastRow() >= 2) {
      var sLastCol = Math.max(sSh.getLastColumn(), 1);
      sellerHdr = sSh.getRange(1, 1, 1, sLastCol).getValues()[0];
      sellerVals = sSh.getRange(2, 1, sSh.getLastRow() - 1, sLastCol).getValues();
    }
  } catch (eS) {}
  var cutoffMs = Date.now() - prAutoMakerRewalkDays_() * 24 * 60 * 60 * 1000;
  var byBand = { 小: [], 中: [], 大: [], 特大: [] };
  var unk = [];
  var seen = {};
  var r;
  for (r = 0; r < vals.length; r++) {
    var parent = String(vals[r][idxMaker] || '').trim();
    if (!parent || seen[parent]) continue;
    seen[parent] = true;
    if (PR_HIT_EXCLUDE_SET_[parent] || prJittaExcludeSeed_(parent)) continue;
    if (idxCk >= 0 && prCkIsTrue_(vals[r][idxCk])) continue;
    var jitta = idxJ >= 0 ? String(vals[r][idxJ] || '').trim() : '';
    if (!prMakerJittaWashEligible_(jitta)) continue;
    var lastMs = prCellToMsTokyo_(idxDay >= 0 ? vals[r][idxDay] : '');
    var never = lastMs == null;
    if (!never && lastMs > cutoffMs) continue;
    if (sellerVals.length && prMakerFoodZeroOnly_(parent, sellerVals, sellerHdr)) continue;
    var prog = idxProg >= 0 ? String(vals[r][idxProg] || '') : '';
    var stRow = idxSt >= 0 ? String(vals[r][idxSt] || '').trim() : '';
    if (prWashTitleZeroBlocked_(parent, stRow, prog)) continue;
    if (prWashHoldBlockedForSelect_(stRow, prog)) continue;
    var estCell = idxEst >= 0 ? vals[r][idxEst] : '';
    var est = prMakerCatalogEstimate_(parent, prog, estCell);
    var band =
      idxBand >= 0 && String(vals[r][idxBand] || '').trim()
        ? String(vals[r][idxBand] || '').trim()
        : prMakerCatalogBand_(est);
    var hitN = idxHit >= 0 ? Number(vals[r][idxHit]) || 0 : 0;
    var item = {
      sheetRow: r + 2,
      rowIdx: r,
      parent: parent,
      jitta: jitta,
      pref: idxPref >= 0 ? String(vals[r][idxPref] || '').trim() : '',
      hit: hitN,
      hitOk: hitN >= 2,
      band: band,
      never: never,
      lastMs: never ? 0 : lastMs,
      via: 'hb_band' + band + '_hit' + hitN,
    };
    if (band === PR_BAND_UNK_) {
      unk.push(item);
      continue;
    }
    // Catalog見積未了は Keepa選定プールに入れない（列挙優先）
    var nEst = Number(String(estCell == null ? '' : estCell).replace(/[^\d.-]/g, ''));
    if (!( !isNaN(nEst) && nEst > 0 && String(estCell).trim() !== '' )) {
      unk.push(item);
      continue;
    }
    if (!byBand[band]) continue;
    byBand[band].push(item);
  }
  function sortBand(a, b) {
    if (a.hitOk !== b.hitOk) return a.hitOk ? -1 : 1;
    if (b.hit !== a.hit) return b.hit - a.hit;
    if (a.never !== b.never) return a.never ? -1 : 1;
    if (a.never && b.never) return a.rowIdx - b.rowIdx;
    if (a.lastMs !== b.lastMs) return a.lastMs - b.lastMs;
    return a.rowIdx - b.rowIdx;
  }
  ['小', '中', '大', '特大'].forEach(function (b) {
    byBand[b].sort(sortBand);
  });
  unk.sort(sortBand);
  return { byBand: byBand, unk: unk };
}

function prCountWashableSmallMed_() {
  var p = prCycleMakerBandPools_();
  return p.byBand['小'].length + p.byBand['中'].length;
}

function prPickEnumUnknownRows_(cap) {
  cap = cap || 1;
  return prCycleMakerBandPools_().unk.slice(0, cap);
}

function prCycleHeartbeatEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_CYCLE_HEARTBEAT_ENABLED_PROP_);
  if (raw == null || raw === '') return true;
  return String(raw).toLowerCase() !== 'false';
}

function prCycleAutoCkDailyCap_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_CYCLE_AUTO_CK_CAP_PROP_) ||
      PR_CYCLE_AUTO_CK_CAP_DEFAULT_
  );
  if (!(n > 0)) n = PR_CYCLE_AUTO_CK_CAP_DEFAULT_;
  return Math.floor(n);
}

function prCycleAutoCkAlreadyToday_() {
  var props = PropertiesService.getScriptProperties();
  var day = prYmdTokyo_();
  if (String(props.getProperty(PR_CYCLE_AUTO_CK_DAY_PROP_) || '') !== day) return 0;
  return Number(props.getProperty(PR_CYCLE_AUTO_CK_COUNT_PROP_) || 0) || 0;
}

function prCycleAutoCkBump_(n) {
  var props = PropertiesService.getScriptProperties();
  var day = prYmdTokyo_();
  if (String(props.getProperty(PR_CYCLE_AUTO_CK_DAY_PROP_) || '') !== day) {
    props.setProperty(PR_CYCLE_AUTO_CK_DAY_PROP_, day);
    props.setProperty(PR_CYCLE_AUTO_CK_COUNT_PROP_, String(n));
    return;
  }
  var cur = Number(props.getProperty(PR_CYCLE_AUTO_CK_COUNT_PROP_) || 0) || 0;
  props.setProperty(PR_CYCLE_AUTO_CK_COUNT_PROP_, String(cur + n));
}

/** 調査CK=TRUE かつ未フルの店数（A-light 待ち。調査済みの誤レ点は数えない）。 */
function prSellerUncheckedCount_() {
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
  if (!sh || sh.getLastRow() < 2) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  if (idxCk < 0) return 0;
  var idxWalk = prHeaderIndex_(hdr, '巡回日');
  var idxSrc = prHeaderIndex_(hdr, '採取元');
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var n = 0;
  var r;
  for (r = 0; r < vals.length; r++) {
    if (!prCkIsTrue_(vals[r][idxCk])) continue;
    if (prSellerAlreadyShopResearched_(vals[r], idxWalk, idxSrc)) continue;
    n++;
  }
  return n;
}

function prSellerIsBlockedAutoCk_(row, idxName, idxId, idxHold, idxMemo) {
  var name = idxName >= 0 ? String(row[idxName] || '').trim().toLowerCase() : '';
  var sid = idxId >= 0 ? String(row[idxId] || '').trim().toLowerCase() : '';
  var blocks = ['amazon.co.jp', 'amazon', 'アマゾン', 'amazon.jp'];
  var i;
  for (i = 0; i < blocks.length; i++) {
    if ((name && name.indexOf(blocks[i]) >= 0) || sid === blocks[i]) return true;
  }
  if (idxHold >= 0 && prCkIsTrue_(row[idxHold])) return true;
  if (idxMemo >= 0 && /hold/i.test(String(row[idxMemo] || ''))) return true;
  return false;
}

/** 自動レ点候補数（CK=OFF・調査済み除外・Amazon/hold 除外）。 */
function prSellerAutoCkCandidateCount_() {
  return prPickAutoCkSellerIds_(9999).length;
}

/**
 * 店フル／浅いholdまで終わった店か（再レ点・再A-lightの対象外）。
 * @param {Array} row
 * @param {number} idxWalk 巡回日
 * @param {number} idxSrc 採取元
 * @return {boolean}
 */
function prSellerAlreadyShopResearched_(row, idxWalk, idxSrc) {
  if (idxWalk >= 0 && String(row[idxWalk] || '').trim()) return true;
  if (idxSrc >= 0) {
    var src = String(row[idxSrc] || '');
    if (/alight\+|shop_full|xl_shallow/i.test(src)) return true;
  }
  return false;
}

function prEnsureSellerSmallHitCol_(sh) {
  sh = sh || prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_SELLER_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = prHeaderIndex_(hdr, PR_SELLER_SMALL_HIT_COL_);
  if (idx >= 0) return idx;
  sh.getRange(1, lastCol + 1).setValue(PR_SELLER_SMALL_HIT_COL_);
  return lastCol;
}

/**
 * 小・中の門通過 ASIN から sellerId→小規模出現を全走査再計算。
 * @return {number} 更新した店数
 */
function prRecomputeSmallMakerHits_() {
  var makerBands = {};
  try {
    var mSh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
    if (mSh && mSh.getLastRow() >= 2) {
      var mLast = Math.max(mSh.getLastColumn(), 1);
      var mHdr = mSh.getRange(1, 1, 1, mLast).getValues()[0];
      var idxM = prHeaderIndex_(mHdr, 'メーカー');
      var idxB = prHeaderIndex_(mHdr, PR_BAND_COL_);
      var idxEst = prHeaderIndex_(mHdr, PR_CATALOG_EST_COL_);
      var idxProg = prHeaderIndex_(mHdr, PR_WASH_PROGRESS_HEADER_);
      var mVals = mSh.getRange(2, 1, mSh.getLastRow() - 1, mLast).getValues();
      var mr;
      for (mr = 0; mr < mVals.length; mr++) {
        var mk = idxM >= 0 ? String(mVals[mr][idxM] || '').trim() : '';
        if (!mk) continue;
        var band =
          idxB >= 0 && String(mVals[mr][idxB] || '').trim()
            ? String(mVals[mr][idxB] || '').trim()
            : prMakerCatalogBand_(
                prMakerCatalogEstimate_(
                  mk,
                  idxProg >= 0 ? String(mVals[mr][idxProg] || '') : '',
                  idxEst >= 0 ? mVals[mr][idxEst] : ''
                )
              );
        makerBands[mk] = band;
      }
    }
  } catch (eM) {}
  var hits = {};
  try {
    var cSh = prSpreadsheet_().getSheetByName(PR_CAND_SHEET_);
    if (cSh && cSh.getLastRow() >= 2) {
      var cLast = Math.max(cSh.getLastColumn(), 1);
      var cHdr = cSh.getRange(1, 1, 1, cLast).getValues()[0];
      var idxAsin = prHeaderIndex_(cHdr, 'ASIN');
      var idxGate = prHeaderIndex_(cHdr, '門結果');
      var idxMaker = prHeaderIndex_(cHdr, 'メーカー');
      var idxSid = prHeaderIndex_(cHdr, 'sellerId');
      var cVals = cSh.getRange(2, 1, cSh.getLastRow() - 1, cLast).getValues();
      var cr;
      for (cr = 0; cr < cVals.length; cr++) {
        if (String(cVals[cr][idxGate] || '') !== '通過') continue;
        var maker = idxMaker >= 0 ? String(cVals[cr][idxMaker] || '').trim() : '';
        var band = makerBands[maker] || '';
        if (band !== PR_BAND_S_ && band !== PR_BAND_M_) continue;
        var sid = idxSid >= 0 ? String(cVals[cr][idxSid] || '').trim() : '';
        if (!sid) continue;
        hits[sid] = (hits[sid] || 0) + 1;
      }
    }
  } catch (eC) {}
  var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_SELLER_SHEET_);
  var idxHit = prEnsureSellerSmallHitCol_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxId = prHeaderIndex_(hdr, 'sellerId');
  if (idxId < 0 || sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var updated = 0;
  var r;
  for (r = 0; r < vals.length; r++) {
    var sid2 = String(vals[r][idxId] || '').trim();
    if (!sid2) continue;
    var n = hits[sid2] || 0;
    if (Number(vals[r][idxHit]) === n) continue;
    sh.getRange(r + 2, idxHit + 1).setValue(n);
    updated++;
  }
  return updated;
}

/**
 * 日次cap内で調査CKをON。戻り: 付けた sellerId[]
 * 調査済み（巡回日／alight完走）は付けない。
 */
function prPickAutoCkSellerIds_(capLeft) {
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
  if (!sh || sh.getLastRow() < 2) return [];
  prEnsureSellerSmallHitCol_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  var idxId = prHeaderIndex_(hdr, 'sellerId');
  var idxName = prHeaderIndex_(hdr, '店名');
  var idxHit = prHeaderIndex_(hdr, PR_SELLER_SMALL_HIT_COL_);
  var idxHold = prHeaderIndex_(hdr, 'hold');
  if (idxHold < 0) idxHold = prHeaderIndex_(hdr, '保留');
  var idxMemo = prHeaderIndex_(hdr, 'メモ');
  var idxWalk = prHeaderIndex_(hdr, '巡回日');
  var idxSrc = prHeaderIndex_(hdr, '採取元');
  var idxFood = prHeaderIndex_(hdr, '食品比率');
  if (idxFood < 0) idxFood = prHeaderIndex_(hdr, '食品・飲料・お酒');
  if (idxCk < 0 || idxId < 0) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var pool = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    if (prCkIsTrue_(vals[r][idxCk])) continue;
    var sid = String(vals[r][idxId] || '').trim();
    if (!sid) continue;
    if (prSellerAlreadyShopResearched_(vals[r], idxWalk, idxSrc)) continue;
    if (prSellerIsBlockedAutoCk_(vals[r], idxName, idxId, idxHold, idxMemo)) continue;
    if (idxFood >= 0) {
      var foodRaw = String(vals[r][idxFood] || '').trim();
      var ratio = Number(foodRaw.replace('%', '').replace(/[^\d.\-]/g, '')) || 0;
      if (foodRaw !== '' && ratio <= 0) continue;
    }
    var hit = idxHit >= 0 ? Number(vals[r][idxHit]) || 0 : 0;
    pool.push({ sid: sid, hit: hit, rowIdx: r, sheetRow: r + 2 });
  }
  pool.sort(function (a, b) {
    if (b.hit !== a.hit) return b.hit - a.hit;
    return a.rowIdx - b.rowIdx;
  });
  var left = capLeft == null ? 9999 : Math.max(0, Number(capLeft) || 0);
  return pool.slice(0, left).map(function (x) {
    return x.sid;
  });
}

function prApplyAutoCkSellers_(sellerIds) {
  if (!sellerIds || !sellerIds.length) return 0;
  var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_SELLER_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  var idxId = prHeaderIndex_(hdr, 'sellerId');
  if (idxCk < 0 || idxId < 0) return 0;
  var want = {};
  var i;
  for (i = 0; i < sellerIds.length; i++) want[String(sellerIds[i])] = true;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var n = 0;
  var r;
  for (r = 0; r < vals.length; r++) {
    var sid = String(vals[r][idxId] || '').trim();
    if (!want[sid]) continue;
    if (prCkIsTrue_(vals[r][idxCk])) continue;
    sh.getRange(r + 2, idxCk + 1).setValue(true);
    n++;
  }
  return n;
}

function prCandUntaggedGateCount_() {
  try {
    return prCandPassAsinListForSellerAccum_().length;
  } catch (e) {
    return 0;
  }
}

/**
 * 心拍の次アクション（CYCLE §7 / dry_cycle_heartbeat 同期）。
 * 孤児ロック（A-1 abort／inflight不一致）は先に破棄してから選定する。
 * @return {{action:string, maker?:string, detail:string}}
 */
function prNextHeartbeatAction_() {
  var rec = prWashFullLockReconcile_('hb_action');
  if (prWashLockLeftSec_() > 0) {
    return {
      action: 'other_busy',
      detail: 'lock ' + (rec.detail || prWashLockDiag_()),
    };
  }
  if (prCycleHeartbeatPaused_()) {
    return { action: 'paused', detail: 'await_01_or_manual' };
  }
  var cont = prPickWashContinueTarget_();
  if (cont && cont.parent && !prCycleDeferLargeContinueForShop_(cont.parent)) {
    var contInfo = prWashContinueDeferInfo_(cont.parent);
    return {
      action: 'wash_continue',
      maker: cont.parent,
      detail:
        'in_progress:' +
        (cont.via || '') +
        ' ' +
        contInfo.reason +
        (cont.deferredDetail ? ' | ' + cont.deferredDetail : ''),
    };
  }
  var pools = prCycleMakerBandPools_();
  var deferNote = cont && cont.deferredDetail ? ' | ' + cont.deferredDetail : '';
  // Catalog見積未了（規模不明含む）を Keepa より先に列挙し切る
  var catPend = prCycleFirstCatalogPending_();
  if (catPend && catPend.parent) {
    return {
      action: 'enum_unknown',
      maker: catPend.parent,
      detail: 'catalog_first ' + (catPend.detail || '') + deferNote,
    };
  }
  if (pools.unk.length) {
    return {
      action: 'enum_unknown',
      maker: pools.unk[0].parent,
      detail: 'band_unknown' + deferNote,
    };
  }
  // 店フル途中（トークン／時間続き）はメーカー新規洗いより先
  if (prAlightHasContinue_()) {
    return { action: 'alight', detail: 'alight_continue' };
  }
  // §6.3bis 案B: Keepa 完走直後は店レーンを1回挟む（大・特大は門通過→セラー優先）
  var shopParent = prCycleShopAfterSmParent_();
  if (shopParent) {
    var uncheckedShop = prSellerUncheckedCount_();
    var autoLeftShop = Math.max(0, prCycleAutoCkDailyCap_() - prCycleAutoCkAlreadyToday_());
    var autoCandShopRaw = autoLeftShop > 0 ? prSellerAutoCkCandidateCount_() : 0;
    var autoCandShop = prEffectiveAutoCkCandidates_(autoCandShopRaw);
    var untaggedShop = prCandUntaggedGateCount_();
    var shopBand = prMakerBandFromSum_(shopParent);
    var shopLarge = shopBand === PR_BAND_L_ || shopBand === PR_BAND_XL_;
    if (shopLarge) {
      // 大・特大完走後: 門通過→セラー追加 → レ点 → A-light
      if (untaggedShop > 0) {
        return {
          action: 'c0',
          detail: 'shop_after_large parent=' + shopParent + ' untagged=' + untaggedShop,
        };
      }
      if (autoCandShop > 0) {
        return {
          action: 'auto_ck',
          detail: 'shop_after_large parent=' + shopParent + ' auto_ck=' + autoCandShop,
        };
      }
      if (uncheckedShop > 0) {
        return {
          action: 'alight',
          detail: 'shop_after_large parent=' + shopParent + ' shops=' + uncheckedShop,
        };
      }
    } else {
      // 小中完走後: 自動レ点 → A-light → C0
      if (autoCandShop > 0) {
        return {
          action: 'auto_ck',
          detail: 'shop_after_sm parent=' + shopParent + ' auto_ck=' + autoCandShop,
        };
      }
      if (uncheckedShop > 0) {
        return {
          action: 'alight',
          detail: 'shop_after_sm parent=' + shopParent + ' shops=' + uncheckedShop,
        };
      }
      if (untaggedShop > 0) {
        return {
          action: 'c0',
          detail: 'shop_after_sm parent=' + shopParent + ' untagged=' + untaggedShop,
        };
      }
    }
    prCycleClearShopAfterSm_();
  }
  var smN = pools.byBand['小'].length + pools.byBand['中'].length;
  var k = PR_CYCLE_K_SMALL_MED_;
  var unchecked = prSellerUncheckedCount_();
  var autoLeft = Math.max(0, prCycleAutoCkDailyCap_() - prCycleAutoCkAlreadyToday_());
  var autoCandRaw = autoLeft > 0 ? prSellerAutoCkCandidateCount_() : 0;
  var autoCand = prEffectiveAutoCkCandidates_(autoCandRaw);
  var untagged = prCandUntaggedGateCount_();
  var xlRoom = prCycleXlWeeklyHasRoom_();
  var xlSt = prCycleXlWeeklyState_();
  var jittaPend = 0;
  if (prJittaHbLaneEnabled_() && prJittaResolveEnabled_()) {
    try {
      jittaPend = prCountMakerJittaUninvestigated_();
    } catch (eJp) {
      jittaPend = 0;
    }
  }

  // 小・中を先に（大を変数に載せない＝pick→defer 禁止）
  if (pools.byBand['小'].length) {
    return {
      action: 'wash_small',
      maker: pools.byBand['小'][0].parent,
      detail: 'sm_pool=' + smN,
    };
  }
  if (pools.byBand['中'].length) {
    return {
      action: 'wash_med',
      maker: pools.byBand['中'][0].parent,
      detail: 'sm_pool=' + smN,
    };
  }

  // 小中=0: 未調査の実態AIを店ラウンドより先（洗えるメーカーを補充）
  if (smN === 0 && jittaPend > 0) {
    return {
      action: 'jitta_resolve',
      detail:
        'sm_empty jitta_pending=' +
        jittaPend +
        ' cap=' +
        prJittaHbCap_() +
        (autoCandRaw > autoCand ? ' auto_ck_blocked' : '') +
        deferNote,
    };
  }

  // 小中=0: §4.4 店ラウンドを尽くすまで大・特大に入らない
  if (smN === 0) {
    var shopBefore = prCycleShopBeforeLargeAction_({
      unchecked: unchecked,
      autoCand: autoCand,
      untagged: untagged,
      deferNote: deferNote,
    });
    if (shopBefore) return shopBefore;
  }

  // 補助: 小中が薄いが >0 のときは上で wash 済み。ここは sm=0 のみ
  if (smN === 0 && (pools.byBand['大'].length || pools.byBand['特大'].length)) {
    var shopWhy = prAlightShouldShopBeforeLarge_()
      ? 'shop_material_empty'
      : 'shop_rounds_exhausted';
    if (!xlRoom) {
      if (jittaPend > 0) {
        return {
          action: 'jitta_resolve',
          detail:
            'xl_week_cap jitta_pending=' +
            jittaPend +
            ' ' +
            shopWhy +
            deferNote,
        };
      }
      return {
        action: 'idle',
        detail:
          'xl_week_cap count=' +
          xlSt.count +
          '/' +
          prCycleXlWeeklyCap_() +
          ' ' +
          shopWhy +
          deferNote,
      };
    }
    if (pools.byBand['大'].length) {
      return {
        action: 'wash_large',
        maker: pools.byBand['大'][0].parent,
        detail:
          shopWhy +
          ' sm_empty xl_week=' +
          xlSt.count +
          ' rounds=' +
          prAlightRoundCount_() +
          deferNote,
      };
    }
    return {
      action: 'wash_xl',
      maker: pools.byBand['特大'][0].parent,
      detail:
        shopWhy +
        ' sm_l_empty xl_week=' +
        xlSt.count +
        ' rounds=' +
        prAlightRoundCount_() +
        deferNote,
    };
  }

  if (smN < k && autoCand > 0) {
    return { action: 'auto_ck', detail: 'sm_pool=' + smN + '<' + k + ' auto_ck=' + autoCand };
  }
  if (smN < k && unchecked > 0) {
    return { action: 'alight', detail: 'sm_pool=' + smN + '<' + k + ' shops=' + unchecked };
  }
  if (unchecked === 0 && untagged > 0) {
    return { action: 'c0', detail: 'untagged=' + untagged };
  }
  if (jittaPend > 0) {
    return {
      action: 'jitta_resolve',
      detail: 'drain_pending=' + jittaPend + ' cap=' + prJittaHbCap_() + deferNote,
    };
  }
  return { action: 'idle', detail: 'no_work' + deferNote };
}

/**
 * 小中=0 のとき大の前の店ラウンド1手（CYCLE §4.4）。
 * 順: 自動レ点 → A-light →（材料切れ時のみ）C0。大を選定してから defer しない。
 * 自動レ点は「店ラウンド尽きた」判定より先に評価する（未レ点を残したまま大へ行かない）。
 * @param {{unchecked?:number, autoCand?:number, untagged?:number, deferNote?:string}=} opt
 * @return {{action:string, detail:string}|null}
 */
function prCycleShopBeforeLargeAction_(opt) {
  opt = opt || {};
  var unchecked =
    opt.unchecked != null ? Number(opt.unchecked) : prSellerUncheckedCount_();
  var autoLeft = Math.max(0, prCycleAutoCkDailyCap_() - prCycleAutoCkAlreadyToday_());
  var autoCandRaw =
    opt.autoCand != null
      ? Number(opt.autoCand)
      : autoLeft > 0
        ? prSellerAutoCkCandidateCount_()
        : 0;
  // opt.autoCand が既にゲート済みでも、二重呼び出しは無害（0のまま）
  var autoCand =
    opt.autoCand != null ? autoCandRaw : prEffectiveAutoCkCandidates_(autoCandRaw);
  var untagged =
    opt.untagged != null ? Number(opt.untagged) : prCandUntaggedGateCount_();
  var deferNote = opt.deferNote || '';
  var rounds = prAlightRoundCount_();
  var done = prAlightDailyDoneCount_();
  var maxD = prAlightDailyMax_();
  var shopOpen = prAlightShouldShopBeforeLarge_();

  // §4.4-a: 未レ点＋日次cap残 → 店ラウンド終了フラグより優先
  if (autoCand > 0) {
    return {
      action: 'auto_ck',
      detail:
        'shop_before_large auto_ck=' +
        autoCand +
        ' autoLeft=' +
        autoLeft +
        ' done=' +
        done +
        '/' +
        maxD +
        ' rounds=' +
        rounds +
        ' shopOpen=' +
        (shopOpen ? '1' : '0') +
        deferNote,
    };
  }
  // §4.4-b: レ点済・未フル（CK=ON）
  if (unchecked > 0 && (shopOpen || done < maxD)) {
    return {
      action: 'alight',
      detail:
        'shop_before_large done=' +
        done +
        '/' +
        maxD +
        ' rounds=' +
        rounds +
        ' shops_ck_on=' +
        unchecked +
        deferNote,
    };
  }
  // §4.4-f: 店もレ点候補も無いときだけ C0
  if (untagged > 0 && shopOpen) {
    return {
      action: 'c0',
      detail:
        'shop_before_large untagged=' +
        untagged +
        ' done=' +
        done +
        ' rounds=' +
        rounds +
        deferNote,
    };
  }
  return null;
}

/** 心拍アクションを1回実行し、必要なら次トリガーを張る。終了時は生存番犬で時計を点検。 */
function prDispatchHeartbeat_(runId) {
  if (!prCycleHeartbeatEnabled_()) {
    prAppendLog_([runId, 'heartbeat', 'DONE', new Date(), '', '', '', 'heartbeat_disabled']);
    return { action: 'disabled' };
  }
  var hbAct = { action: '' };
  try {
  var sanDetail = '';
  try {
    var san = prSanitizeDeferredLargeKeepaStuck_();
    sanDetail = san.detail || '';
    // 原因特定のため n=0 でも①ログへ（sm/unk ゲート・解除メーカー名）
    prAppendLog_([runId, 'heartbeat sanitize', 'DONE', new Date(), san.n, '', '', sanDetail]);
  } catch (eSan) {
    sanDetail = 'sanitize_err=' + String((eSan && eSan.message) || eSan).slice(0, 80);
    prAppendLog_([runId, 'heartbeat sanitize', 'FAILED', new Date(), '', '', '', sanDetail]);
  }
  try {
    var fc = prCycleRewashFalseCompletes_({ ckOnly: false, quiet: true, runId: runId });
    if (fc && fc.n > 0) {
      prAppendLog_([
        runId,
        'heartbeat false_complete',
        'DONE',
        new Date(),
        fc.n,
        '',
        '',
        String(fc.detail || '').slice(0, 200),
      ]);
    }
  } catch (eFc) {
    prAppendLog_([
      runId,
      'heartbeat false_complete',
      'FAILED',
      new Date(),
      '',
      '',
      '',
      String((eFc && eFc.message) || eFc).slice(0, 120),
    ]);
  }
  var pool = { smN: 0, unkN: 0, largeN: 0, xlN: 0 };
  try {
    pool = prCyclePoolCounts_();
  } catch (eP) {}
  var contProbe = prPickWashContinueTarget_();
  var contWhy = '';
  if (contProbe && contProbe.parent) {
    var dInfo = prWashContinueDeferInfo_(contProbe.parent);
    contWhy =
      'cont=' +
      contProbe.parent +
      ' via=' +
      (contProbe.via || '') +
      ' ' +
      dInfo.reason;
  } else if (contProbe && contProbe.deferredDetail) {
    contWhy = 'cont=none ' + contProbe.deferredDetail;
  } else {
    contWhy = 'cont=none';
  }
  prAppendLog_([
    runId,
    'heartbeat pool',
    'DONE',
    new Date(),
    pool.smN,
    pool.unkN,
    '',
    'sm=' +
      pool.smN +
      ' unk=' +
      pool.unkN +
      ' L=' +
      pool.largeN +
      ' XL=' +
      pool.xlN +
      ' ' +
      contWhy +
      (sanDetail && sanDetail.indexOf('sanitize_fixed n=0') < 0 && sanDetail.indexOf('n=0') < 0
        ? ' | ' + sanDetail.slice(0, 80)
        : ''),
  ]);

  var act = prNextHeartbeatAction_();
  var actWhy = act.detail || '';
  if (act.maker) {
    var bandAct = prMakerBandFromSum_(act.maker);
    actWhy += ' band=' + bandAct;
    if (act.action === 'wash_continue' || String(act.action).indexOf('wash_') === 0) {
      var whyC = prWashContinueDeferInfo_(act.maker);
      actWhy += ' | ' + whyC.reason;
    }
  }
  prAppendLog_([
    runId,
    'heartbeat',
    'RUNNING',
    new Date(),
    '',
    act.maker || '',
    '',
    'action=' + act.action + ' ' + actWhy,
  ]);
  Logger.log('[仕入れ検討①] runId=%s heartbeat action=%s maker=%s %s', runId, act.action, act.maker || '', actWhy);

  hbAct = act;
  if (act.action === 'paused' || act.action === 'other_busy') {
    if (act.action === 'paused' && prCountResume01Triggers_() < 1) {
      prCycleArmResume01_();
    }
    var busyArm = 0;
    if (act.action === 'other_busy') {
      // 真のビジーでも 30分待ちにせず、短周期で再選定（孤児クリア後の catalog_first 用）
      PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      busyArm = prScheduleWashFullTrigger_(1, {
        runId: runId,
        parent: '',
        tag: 'hb_lock_retry',
      });
    }
    prAppendLog_([
      runId,
      'heartbeat',
      'DONE',
      new Date(),
      busyArm || '',
      '',
      '',
      act.action +
        ' ' +
        act.detail +
        (busyArm ? ' retry_armed afterMin=' + busyArm : '') +
        ' ' +
        prCycleClockSnapshot_(),
    ]);
    return act;
  }

  var props = PropertiesService.getScriptProperties();
  props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');

  if (act.action === 'wash_continue' || act.action.indexOf('wash_') === 0 || act.action === 'enum_unknown') {
    var parent = act.maker;
    if (!parent) {
      var armNm = prCycleArmHeartbeatAfterIdle_(runId, 'hb_no_maker');
      prAppendLog_([
        runId,
        'heartbeat',
        'DONE',
        new Date(),
        armNm,
        '',
        '',
        'no_maker ' + act.action + ' heartbeat_armed afterMin=' + armNm,
      ]);
      return act;
    }
    // 二重ガード: 大特大を誤って wash_continue したとき。小中／不明残りだけ帯確定に戻す
    if (act.action === 'wash_continue') {
      var guard = prWashContinueDeferInfo_(parent);
      if (guard.defer) {
        if (prWashContinueDeferShouldFinalize_(guard.reason)) {
          prFinalizeEnumOnlySize_(parent);
        }
        var armG = prCycleArmHeartbeatAfterIdle_(runId, 'hb_defer_xl');
        prAppendLog_([
          runId,
          'heartbeat defer_xl',
          'DONE',
          new Date(),
          armG,
          parent,
          '',
          guard.reason +
            (prWashContinueDeferShouldFinalize_(guard.reason) ? ' finalize' : ' keep_progress') +
            ' heartbeat_armed afterMin=' +
            armG,
        ]);
        return { action: 'defer_xl', maker: parent, detail: guard.reason };
      }
      var bandCont = prMakerBandFromSum_(parent);
      if (bandCont === PR_BAND_L_ || bandCont === PR_BAND_XL_) {
        if (!prCycleXlWeeklyTryConsume_(parent)) {
          var armCapC = prCycleArmHeartbeatAfterIdle_(runId, 'hb_xl_week_cap_cont');
          prAppendLog_([
            runId,
            'heartbeat',
            'DONE',
            new Date(),
            armCapC,
            parent,
            '',
            'xl_week_cap_block continue heartbeat_armed afterMin=' + armCapC,
          ]);
          return { action: 'xl_week_cap', maker: parent, detail: 'blocked_continue' };
        }
      }
    }
    var needles = prMakerWashNeedles_(parent, '');
    var entry = prMakerCatalogEntry_(parent, '', '');
    var enumOnly = act.action === 'enum_unknown';
    // 規模帯の確定は1実行でまとめる（1社ごとの受け渡し待ちが実測7分だった）
    if (enumOnly && prCycleEnumBatchN_() > 1) {
      var batch = prCycleRunEnumBatch_(runId, parent);
      var hbMinB = prCycleArmHeartbeatAfterIdle_(runId, 'hb_after_enum');
      prAppendLog_([
        runId,
        'heartbeat enum',
        'DONE',
        new Date(),
        batch.n,
        parent,
        '',
        batch.detail + ' heartbeat_armed afterMin=' + hbMinB,
      ]);
      return act;
    }
    if (!enumOnly && (act.action === 'wash_large' || act.action === 'wash_xl')) {
      if (!prCycleXlWeeklyTryConsume_(parent)) {
        var armCap = prCycleArmHeartbeatAfterIdle_(runId, 'hb_xl_week_cap');
        prAppendLog_([
          runId,
          'heartbeat',
          'DONE',
          new Date(),
          armCap,
          parent,
          '',
          'xl_week_cap_block heartbeat_armed afterMin=' + armCap,
        ]);
        return { action: 'xl_week_cap', maker: parent, detail: 'blocked' };
      }
    }
    try {
      var washOut = prRunMakerWashFull_({
        parent: parent,
        entry: entry,
        needles: needles,
        tag: enumOnly ? 'hb_enum' : 'hb_' + act.action,
        runId: runId,
        loopKeepa: !enumOnly,
        quiet: true,
        armTrigger: !enumOnly,
        enumOnly: enumOnly,
      });
      // enumOnly: 規模確定のみ。Keepa中にしない・次心拍へ
      if (enumOnly) {
        prFinalizeEnumOnlySize_(parent);
        var hbMin = prCycleArmHeartbeatAfterIdle_(runId, 'hb_after_enum');
        prAppendLog_([
          runId,
          'heartbeat enum',
          'DONE',
          new Date(),
          hbMin,
          parent,
          '',
          'band_done ' + ((washOut && washOut.detail) || '') + ' heartbeat_armed afterMin=' + hbMin,
        ]);
      }
    } catch (eW) {
      var msgW = prRedactKey_(String((eW && eW.message) || eW));
      if (/タイトル含有0/.test(msgW)) {
        var tzW = prHandleWashTitleZero_({
          parent: parent,
          runId: runId,
          tag: 'hb_title0',
          armTrigger: true,
          quiet: true,
        });
        prAppendLog_([
          runId,
          'heartbeat',
          'DONE',
          new Date(),
          tzW.count,
          parent,
          '',
          tzW.detail.slice(0, 200),
        ]);
        return act;
      }
      prAppendLog_([runId, 'heartbeat', 'FAILED', new Date(), '', parent, '', msgW.slice(0, 200)]);
      prScheduleWashFullTrigger_(prWashFullTriggerAfterMin_(5, ''), {
        runId: runId,
        parent: parent,
        tag: 'hb_retry',
      });
    }
    return act;
  }

  if (act.action === 'auto_ck') {
    try {
      prRecomputeSmallMakerHits_();
    } catch (eR) {}
    var left = Math.max(0, prCycleAutoCkDailyCap_() - prCycleAutoCkAlreadyToday_());
    if (prJittaBacklogBlocksAutoCk_()) {
      prAppendLog_([
        runId,
        'heartbeat auto_ck',
        'DONE',
        new Date(),
        0,
        '',
        '',
        'blocked_jitta_backlog pending>=' + prJittaBacklogGate_(),
      ]);
      prScheduleWashFullTrigger_(1, { runId: runId, parent: '', tag: 'hb_after_auto_ck_blocked' });
      return act;
    }
    var ids = prPickAutoCkSellerIds_(left);
    var n = prApplyAutoCkSellers_(ids);
    if (n > 0) prCycleAutoCkBump_(n);
    try {
      SpreadsheetApp.flush();
    } catch (eFlush) {}
    prAppendLog_([runId, 'heartbeat auto_ck', 'DONE', new Date(), n, '', '', 'ids=' + ids.join('|').slice(0, 120)]);
    // §4.4: レ点直後は同じ心拍で A-light へ（1分トリガーが C4/overlap で欠ける穴を塞ぐ）
    var uncAfter = 0;
    try {
      uncAfter = prSellerUncheckedCount_();
    } catch (eU) {}
    if (n > 0 && uncAfter > 0) {
      prAppendLog_([
        runId,
        'heartbeat chain_alight',
        'RUNNING',
        new Date(),
        uncAfter,
        '',
        '',
        'shop_after_auto_ck n=' + n + ' unchecked=' + uncAfter,
      ]);
      act = {
        action: 'alight',
        detail: 'shop_after_auto_ck n=' + n + ' unchecked=' + uncAfter,
      };
    } else {
      prScheduleWashFullTrigger_(1, { runId: runId, parent: '', tag: 'hb_after_auto_ck' });
      return act;
    }
  }

  if (act.action === 'jitta_resolve') {
    try {
      if (!prJittaResolveEnabled_()) {
        prAppendLog_([
          runId,
          'heartbeat jitta',
          'DONE',
          new Date(),
          0,
          '',
          '',
          'jitta_disabled',
        ]);
      } else {
        var targetsJ = prCollectJittaResolveTargets_({});
        var batchJ = prResolveJittaBatch_(targetsJ, {
          runId: runId,
          cap: prJittaHbCap_(),
        });
        var detailJ =
          'resolve=' +
          batchJ.resolved +
          ' auto=' +
          batchJ.autoN +
          ' skip=' +
          batchJ.skipped +
          ' cap=' +
          batchJ.cap +
          ' pending=' +
          targetsJ.length;
        prAppendLog_([
          runId,
          'heartbeat jitta',
          'DONE',
          new Date(),
          batchJ.resolved,
          batchJ.autoN,
          '',
          detailJ,
        ]);
        Logger.log('[仕入れ検討①] runId=%s step=jitta_resolve state=DONE %s', runId, detailJ);
      }
    } catch (eJitta) {
      prAppendLog_([
        runId,
        'heartbeat jitta',
        'FAILED',
        new Date(),
        '',
        '',
        '',
        prRedactKey_(String((eJitta && eJitta.message) || eJitta)).slice(0, 160),
      ]);
    }
    prScheduleWashFullTrigger_(1, { runId: runId, parent: '', tag: 'hb_after_jitta' });
    return act;
  }

  if (act.action === 'alight') {
    try {
      var ar = prRunAlightHit_({ runId: runId });
      try {
        prRecomputeSmallMakerHits_();
      } catch (eH) {}
      var dAl = String(act.detail || '');
      if (
        dAl.indexOf('shop_after_sm') >= 0 ||
        dAl.indexOf('shop_after_large') >= 0 ||
        dAl.indexOf('shop_after_auto_ck') >= 0 ||
        prCycleShopAfterSmParent_()
      ) {
        // 店フル途中は印を残す。完走／hold のときだけ消す
        if (!(ar && ar.needContinue)) prCycleClearShopAfterSm_();
      }
      prAppendLog_([
        runId,
        'heartbeat alight',
        ar && ar.ok ? 'DONE' : ar && ar.needContinue ? 'RETRYING' : 'FAILED',
        new Date(),
        (ar && ar.doneN) || '',
        (ar && ar.hitMakers) || '',
        '',
        (ar && ar.detail) || '',
      ]);
      if (ar && ar.needContinue) {
        var waitA = prWashFullTriggerAfterMin_(ar.waitMin, ar.refillIn);
        // 続き中も洗い inflight は持たない（死亡時の overdue 防止）。busy は維持。
        if (prShopLaneClearInflightEnabled_()) {
          prWashInflightClear_('alight_continue:' + runId);
          prShopHbBusySet_(runId);
        }
        prScheduleWashFullTrigger_(waitA, {
          runId: runId,
          parent: '',
          tag: 'hb_alight_continue',
          calcTokensLeft: ar.tokensLeft,
        });
        return act;
      }
    } catch (eA) {
      prAppendLog_([
        runId,
        'heartbeat alight',
        'FAILED',
        new Date(),
        '',
        '',
        '',
        prRedactKey_(String((eA && eA.message) || eA)).slice(0, 160),
      ]);
    }
    // 1店完了 → busy 解除＋即次（死亡待ちをしない）
    if (prShopLaneClearInflightEnabled_()) {
      prWashInflightClear_('after_alight:' + runId);
      prShopHbBusyClear_('after_alight');
    }
    prScheduleWashFullTrigger_(1, { runId: runId, parent: '', tag: 'hb_after_alight' });
    return act;
  }

  if (act.action === 'c0') {
    try {
      var c0 = prRunSellerAccum_({ runId: runId });
      var newN = (c0 && c0.newN) || 0;
      if (newN > 0) {
        var day = prYmdTokyo_();
        var p = PropertiesService.getScriptProperties();
        if (String(p.getProperty('PR_CYCLE_C0_DAY') || '') !== day) {
          p.setProperty('PR_CYCLE_C0_DAY', day);
          p.setProperty('PR_CYCLE_C0_NEW_TODAY', String(newN));
        } else {
          var cur = Number(p.getProperty('PR_CYCLE_C0_NEW_TODAY') || 0) || 0;
          p.setProperty('PR_CYCLE_C0_NEW_TODAY', String(cur + newN));
        }
      }
      prAppendLog_([
        runId,
        'heartbeat c0',
        c0 && c0.ok !== false ? 'DONE' : 'FAILED',
        new Date(),
        newN,
        '',
        '',
        (c0 && c0.detail) || '',
      ]);
    } catch (e0) {
      prAppendLog_([
        runId,
        'heartbeat c0',
        'FAILED',
        new Date(),
        '',
        '',
        '',
        prRedactKey_(String((e0 && e0.message) || e0)).slice(0, 160),
      ]);
    }
    prScheduleWashFullTrigger_(1, { runId: runId, parent: '', tag: 'hb_after_c0' });
    return act;
  }

  // idle: 種切れでも短時間後に再予約（沈黙防止）
  prScheduleWashFullTrigger_(prWashFullTriggerAfterMin_(30, ''), {
    runId: runId,
    parent: '',
    tag: 'hb_idle',
  });
  prAppendLog_([runId, 'heartbeat', 'DONE', new Date(), 30, '', '', 'idle rearm']);
  return act;
  } catch (eHb) {
    var msgHb = prRedactKey_(String((eHb && eHb.message) || eHb));
    prAppendLog_([
      runId,
      'heartbeat',
      'FAILED',
      new Date(),
      '',
      '',
      '',
      'outer ' + msgHb.slice(0, 180),
    ]);
    PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
    prScheduleWashFullTrigger_(2, {
      runId: runId,
      parent: '',
      tag: 'hb_outer_rearm',
    });
    hbAct = { action: 'outer_fail' };
    return hbAct;
  } finally {
    try {
      // alight スライス続き中は busy 維持。それ以外は yield で立てた busy を解放
      var keepBusy = !!(hbAct && hbAct.action === 'alight' && prAlightHasContinue_());
      if (!keepBusy) prShopHbBusyClear_('hb_finally:' + ((hbAct && hbAct.action) || ''));
    } catch (eBusy) {}
    try {
      prCycleEnsureLiveness_(runId, 'hb_finally');
      prAppendLog_([
        runId,
        'heartbeat end',
        'DONE',
        new Date(),
        '',
        (hbAct && hbAct.maker) || '',
        '',
        'action=' + ((hbAct && hbAct.action) || '') + ' ' + prCycleClockSnapshot_(),
      ]);
    } catch (eEnd) {}
  }
}

/** メーカーマスタの第2語＋実態キャッシュをクリア（サマリミラーも消す）。 */
function prClearMakerMasterForResample_(maker) {
  maker = String(maker || '').trim();
  if (!maker) return false;
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_MASTER_SHEET_);
  if (!sh || sh.getLastRow() < 2) return false;
  var idx = prEnsureMakerMasterJittaCols_(sh);
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  var cleared = false;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idx['メーカー']] || '').trim() !== maker) continue;
    var sheetRow = r + 2;
    var c;
    for (c = 0; c < PR_MAKER_MASTER_Q2_COLS_.length; c++) {
      var q2 = PR_MAKER_MASTER_Q2_COLS_[c];
      if (idx[q2] != null) sh.getRange(sheetRow, idx[q2] + 1).setValue('');
    }
    for (c = 0; c < PR_MAKER_JITTA_MASTER_COLS_.length; c++) {
      var jc = PR_MAKER_JITTA_MASTER_COLS_[c];
      if (idx[jc] != null) sh.getRange(sheetRow, idx[jc] + 1).setValue('');
    }
    cleared = true;
    break;
  }
  var sumSh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (sumSh) {
    prEnsureMakerSumJittaMirrorCols_(sumSh);
    var sumRow = 0;
    var lastCol2 = Math.max(sumSh.getLastColumn(), 1);
    var hdr = sumSh.getRange(1, 1, 1, lastCol2).getValues()[0];
    var idxMaker = prHeaderIndex_(hdr, 'メーカー');
    if (idxMaker >= 0 && sumSh.getLastRow() >= 2) {
      var svals = sumSh.getRange(2, 1, sumSh.getLastRow() - 1, lastCol2).getValues();
      var sr;
      for (sr = 0; sr < svals.length; sr++) {
        if (String(svals[sr][idxMaker] || '').trim() === maker) {
          sumRow = sr + 2;
          break;
        }
      }
    }
    if (sumRow) {
      var mi;
      for (mi = 0; mi < PR_MAKER_JITTA_MIRROR_COLS_.length; mi++) {
        var mc = PR_MAKER_JITTA_MIRROR_COLS_[mi];
        var idxM = prHeaderIndex_(hdr, mc);
        if (idxM >= 0) sumSh.getRange(sumRow, idxM + 1).setValue('');
      }
    }
  }
  Logger.log('[仕入れ検討①] maker_master_clear maker=%s cleared=%s', maker, cleared);
  return cleared;
}

function prMenuResetStaleMakerMaster() {
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=maker_master_reset state=RUNNING', runId);
  var names = [];
  var k;
  for (k in PR_MAKER_MASTER_STALE_RESET_) {
    if (PR_MAKER_MASTER_STALE_RESET_[k]) names.push(k);
  }
  var n = 0;
  var i;
  for (i = 0; i < names.length; i++) {
    if (prClearMakerMasterForResample_(names[i])) n++;
  }
  var detail = 'cleared=' + n + ' makers=' + names.join('|');
  prAppendLog_([runId, 'マスタ旧データクリア', 'DONE', new Date(), '', n, '', detail]);
  prMenuToast_('マスタ旧データクリア', detail + ' → 実態AI調査を実行してください');
  Logger.log('[仕入れ検討①] runId=%s step=maker_master_reset state=DONE %s', runId, detail);
  return { ok: true, n: n, detail: detail };
}

/** 調査CK 列が無ければ A の前に挿入。チェック枠とレ点時の青字。 */
function prEnsureLeadingCk_(ssId, title) {
  var readyKey = ssId + ':' + title;
  if (PR_CK_READY_[readyKey]) return PR_CK_READY_[readyKey];
  var sh = SpreadsheetApp.openById(ssId).getSheetByName(title);
  if (!sh) throw new Error('シートなし: ' + title);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = prHeaderIndex_(hdr, PR_CK_HEADER_);
  if (idx < 0) {
    sh.insertColumnBefore(1);
    sh.getRange(1, 1).setValue(PR_CK_HEADER_);
    idx = 0;
    lastCol = Math.max(sh.getLastColumn(), 1);
    var lastRow = Math.max(sh.getLastRow(), 2);
    var endRow = Math.min(sh.getMaxRows(), Math.max(lastRow + 80, 120));
    sh.getRange(2, idx + 1, endRow - 1, 1).insertCheckboxes();
    prEnsureCkBlue_(sh, idx, lastCol, endRow);
  }
  PR_CK_READY_[readyKey] = sh;
  return sh;
}

function prColLetter_(n) {
  var s = '';
  var x = n;
  while (x > 0) {
    var r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/** 調査CK が TRUE の行は文字色を青。同じ数式の既存ルールは付け直す。 */
function prEnsureCkBlue_(sh, idxCk, lastCol, endRow) {
  var formula = '=$' + prColLetter_(idxCk + 1) + '2=TRUE';
  var range = sh.getRange(2, 1, endRow - 1, Math.max(lastCol, 1));
  var old = sh.getConditionalFormatRules() || [];
  var kept = [];
  var i;
  for (i = 0; i < old.length; i++) {
    var f = '';
    try {
      var c = old[i].getBooleanCondition();
      var vals = c && c.getCriteriaValues ? c.getCriteriaValues() : [];
      f = String((vals && vals[0]) || '');
    } catch (e1) {
      f = '';
    }
    if (String(f).indexOf(formula) >= 0) continue;
    kept.push(old[i]);
  }
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setFontColor('#0000ff')
    .setRanges([range])
    .build();
  kept.unshift(rule);
  sh.setConditionalFormatRules(kept);
}

/**
 * /query page0。total≥501（特大）は商品Keepaせず hold（呼び出し側が次の店へ）。
 * @return {{status:string, total:number, asinCount:number, pass:number, drop:number, detail:string}}
 */
function prRunSellerQueryGate_(sellerId, runId, shopName) {
  var sid = String(sellerId || '').trim();
  if (!sid) {
    return { status: 'fail', total: 0, asinCount: 0, pass: 0, drop: 0, detail: 'sellerId空' };
  }
  var name = shopName || '';
  var sfName = prEnsureSellerStorefront_(sid, true);
  if (!name) name = sfName;
  var meas = prMeasureSellerSkuByCategory_(sid);
  var total = Number(meas.skuEst || 0);
  if (meas.hold) {
    Logger.log(
      '[仕入れ検討①] runId=%s HOLD skuEst=%s band=%s seller=%s 次の店へ',
      runId,
      total,
      meas.band,
      sid
    );
    return {
      status: 'hold',
      total: total,
      asinCount: 0,
      pass: 0,
      drop: 0,
      passMakers: [],
      skuBreakdown: meas.skuBreakdown,
      bandBreakdown: meas.bandBreakdown,
      detail:
        '人判断待ち skuEst=' +
        total +
        ' band=' +
        meas.band +
        ' ' +
        meas.skuBreakdown +
        ' seller=' +
        sid,
    };
  }
  var roots = meas.walkRoots;
  var q;
  if (meas.page0Asins) {
    q = {
      totalResults: meas.page0Total,
      asinCount: meas.page0Asins.length,
      asinList: meas.page0Asins,
      tokensConsumed: meas.tokensConsumed,
    };
  } else {
    q = prKeepaQueryGet_(sid, roots, 0, 100);
  }
  var prof = prReadProfile_();
  var pack = prKeepaProductStats90_(q.asinList);
  var products = pack.products;
  var pass = 0;
  var drop = 0;
  var leftover = 0;
  var passMakersSet = {};
  var i;
  for (i = 0; i < products.length; i++) {
    if (prGasTimeUp_()) {
      leftover = products.length - i;
      Logger.log('[仕入れ検討①] runId=%s timeup leftover=%s seller=%s', runId, leftover, sid);
      break;
    }
    var row = prGateProductRow_(products[i], prof, sid, runId, name);
    prUpsertCandidate_(row);
    if (row[8] === '通過') {
      pass++;
      var mfr = String(row[0] || '').trim();
      if (mfr) passMakersSet[mfr] = true;
    } else {
      drop++;
    }
  }
  var passMakers = Object.keys(passMakersSet).sort();
  var qTok = Number(q.tokensConsumed);
  if (!(qTok > 0)) qTok = 0;
  var pTok = Number(pack.tokensConsumed);
  if (!(pTok > 0)) pTok = 0;
  return {
    status: leftover ? 'partial' : 'done',
    total: total,
    asinCount: products.length,
    pass: pass,
    drop: drop,
    passMakers: passMakers,
    tokensConsumed: qTok + pTok,
    detail:
      'pass=' +
      pass +
      ' drop=' +
      drop +
      ' leftover=' +
      leftover +
      (leftover ? ' timeup' : '') +
      ' priceMin=' +
      prof.priceMin +
      ' priceMax=' +
      prof.priceMax +
      ' rankMin=' +
      prof.rankMin +
      ' rankMax=' +
      prof.rankMax +
      ' full+=' +
      pack.fullPlus +
      ' need=' +
      pack.need +
      ' skip=' +
      pack.skip +
      ' qTok=' +
      qTok +
      ' pTok=' +
      pTok +
      ' consumed=' +
      (qTok + pTok) +
      ' seller=' +
      sid,
  };
}

/**
 * レ点店を調査。opt.maxTargets で件数上限（未指定は PR_CK_BATCH_CAP）。
 * @param {{maxTargets?:number, toastTitle?:string, logLabel?:string, forceSellerId?:string, quietEmpty?:boolean, emptyAsAlert?:boolean}} opt
 */
function prMenuCkSellers(opt) {
  opt = opt || {};
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var toastTitle = opt.toastTitle || '① レ点の店';
  var logLabel = opt.logLabel || 'レ点の店を調査';
  Logger.log('[仕入れ検討①] runId=%s step=ck_seller state=RUNNING', runId);
  try {
    if (!prCkEnabled_(PR_CK_SELLER_ENABLED_PROP_)) {
      throw new Error(PR_CK_SELLER_ENABLED_PROP_ + '=false（切戻し中）');
    }
    var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_SELLER_SHEET_);
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      prMenuToast_(toastTitle, 'セラーにデータ行がありません', 8);
      return { ok: false, detail: 'no_rows' };
    }
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
    var idxId = prHeaderIndex_(hdr, 'sellerId');
    var idxName = prHeaderIndex_(hdr, '店名');
    var idxMemo = prHeaderIndex_(hdr, 'メモ');
    var idxWalk = prHeaderIndex_(hdr, '巡回日');
    var idxCnt = prHeaderIndex_(hdr, 'asinList件数');
    var idxSrc = prHeaderIndex_(hdr, '採取元');
    var idxExtractMaker = prHeaderIndex_(hdr, '抽出メーカー');
    var idxPick = prHeaderIndex_(hdr, 'ピック');
    if (idxCk < 0 || idxId < 0) throw new Error('セラーに 調査CK または sellerId が無い');
    var vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var targets = [];
    var trueN = 0;
    var forceSid = String(opt.forceSellerId || '').trim();
    var r;
    for (r = 0; r < vals.length; r++) {
      var sid = String(vals[r][idxId] || '').trim();
      if (forceSid) {
        if (sid !== forceSid) continue;
      } else {
        if (!prCkIsTrue_(vals[r][idxCk])) continue;
        trueN++;
      }
      if (!sid) continue;
      var name = idxName >= 0 ? String(vals[r][idxName] || '').trim() : '';
      if (!name && idxName >= 0) name = prSellerNameFromSameId_(vals, idxId, idxName, sid);
      targets.push({
        sheetRow: r + 2,
        sellerId: sid,
        name: name,
        memo: idxMemo >= 0 ? vals[r][idxMemo] : '',
      });
      if (forceSid) break;
    }
    if (!targets.length) {
      var emptyDetail = forceSid
        ? 'forceSellerId=' + forceSid + ' 行なし'
        : 'TRUE=' + trueN + ' rows=' + (lastRow - 1) + ' sheet=' + PR_SELLER_SHEET_;
      Logger.log('[仕入れ検討①] runId=%s state=DONE %s', runId, emptyDetail);
      prAppendLog_([runId, logLabel, 'DONE', new Date(), '', 0, '', emptyDetail]);
      if (!opt.quietEmpty) {
        var emptyMsg = forceSid
          ? emptyDetail
          : prCkEmptyAlert_(PR_SELLER_SHEET_, lastRow, trueN);
        if (opt.emptyAsAlert) {
          prMenuFailAlert_(emptyMsg);
        } else {
          prMenuToast_(toastTitle, emptyMsg, 15);
        }
      }
      return { ok: false, detail: emptyDetail };
    }
    var cap =
      opt.maxTargets != null && Number(opt.maxTargets) > 0
        ? Math.floor(Number(opt.maxTargets))
        : prCkBatchCap_();
    if (cap > 10) cap = 10;
    var leftover = Math.max(0, targets.length - cap);
    var doneN = 0;
    var holdN = 0;
    var failN = 0;
    var partialN = 0;
    var lines = [];
    var lastGate = null;
    var lastTarget = null;
    var i;
    for (i = 0; i < targets.length && i < cap; i++) {
      if (prGasTimeUp_()) {
        leftover = targets.length - i;
        lines.push('timeup leftover_shops=' + leftover);
        break;
      }
      var t = targets[i];
      lastTarget = t;
      Logger.log(
        '[仕入れ検討①] runId=%s ck_seller %s/%s %s %s',
        runId,
        i + 1,
        Math.min(targets.length, cap),
        t.sellerId,
        t.name
      );
      var g;
      try {
        g = prRunSellerQueryGate_(t.sellerId, runId, t.name);
      } catch (e1) {
        failN++;
        g = {
          status: 'fail',
          total: '',
          asinCount: '',
          passMakers: [],
          detail: prRedactKey_(String((e1 && e1.message) || e1)).slice(0, 180),
        };
      }
      lastGate = g;
      if (prSellerRewalkOverwriteEnabled_()) {
        t.sheetRow = prCollapseSellerDupes_(t.sellerId, t.sheetRow, targets);
        if (idxMemo >= 0) t.memo = sh.getRange(t.sheetRow, idxMemo + 1).getValue();
      }
      if (g.status === 'hold') {
        holdN++;
        if (idxMemo >= 0) {
          sh.getRange(t.sheetRow, idxMemo + 1).setValue(
            prAppendMemo_(t.memo, g.detail + ' ' + prYmdTokyo_())
          );
        }
      } else if (g.status === 'partial') {
        partialN++;
        if (idxMemo >= 0) {
          sh.getRange(t.sheetRow, idxMemo + 1).setValue(
            prAppendMemo_(t.memo, g.detail + ' ' + prYmdTokyo_())
          );
        }
      } else if (g.status === 'done') {
        doneN++;
        sh.getRange(t.sheetRow, idxCk + 1).setValue(false);
        if (idxWalk >= 0) sh.getRange(t.sheetRow, idxWalk + 1).setValue(prYmdTokyo_());
        if (idxCnt >= 0) sh.getRange(t.sheetRow, idxCnt + 1).setValue(g.total);
        if (idxSrc >= 0) sh.getRange(t.sheetRow, idxSrc + 1).setValue('query+storefront0');
      } else {
        failN++;
      }

      // 本来の調査: 門通過メーカーを台帳へ書く（空欄のみではなく今回値で上書き可）
      if (g && g.passMakers && g.passMakers.length) {
        var mk = g.passMakers.join('|');
        if (idxExtractMaker >= 0) {
          if (prSellerRewalkOverwriteEnabled_()) {
            sh.getRange(t.sheetRow, idxExtractMaker + 1).setValue(mk);
          } else {
            prWriteIfEmpty_(sh, t.sheetRow, idxExtractMaker, mk);
          }
        }
        if (idxPick >= 0) {
          if (prSellerRewalkOverwriteEnabled_()) {
            sh.getRange(t.sheetRow, idxPick + 1).setValue('通過');
          } else {
            prWriteIfEmpty_(sh, t.sheetRow, idxPick, '通過');
          }
        }
      }
      lines.push(t.sellerId + ':' + g.status + ' ' + g.detail);
      prAppendLog_([
        runId,
        'レ点の店 ' + t.sellerId,
        g.status === 'fail' ? 'FAILED' : 'DONE',
        new Date(),
        g.total,
        g.asinCount,
        g.tokensConsumed != null ? g.tokensConsumed : '',
        g.detail,
      ]);
    }
    var summary =
      'done=' +
      doneN +
      ' hold=' +
      holdN +
      ' partial=' +
      partialN +
      ' fail=' +
      failN +
      ' leftover=' +
      leftover +
      ' cap=' +
      cap +
      ' overwrite=' +
      (prSellerRewalkOverwriteEnabled_() ? '1' : '0') +
      ' rewalkDays=' +
      prSellerRewalkDays_() +
      ' (仮)';
    Logger.log('[仕入れ検討①] runId=%s state=DONE %s', runId, summary);
    prAppendLog_([runId, logLabel, 'DONE', new Date(), '', targets.length, '', summary]);
    prMenuToast_(toastTitle, summary + ' ' + lines.join(' ').slice(0, 200));
    return {
      ok: failN === 0 && holdN === 0,
      detail: summary,
      gate: lastGate,
      target: lastTarget,
      runId: runId,
    };
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, logLabel, 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/** 1. 店1件通し: レ点TRUEの先頭1店だけ query→門→①候補（台帳更新あり）。レ点なしはアラートで終了。 */
function prMenuSellerOneThrough() {
  return prMenuCkSellers({
    maxTargets: 1,
    toastTitle: '① 店1件通し',
    logLabel: '店1件通し',
    emptyAsAlert: true,
  });
}

/**
 * 0. 試験 A→D V1b: レ点1店（無ければモリタ固定）→門→メーカーサマリ追加→最大1メーカー完走→D1→D2。
 * 台帳・サマリへの本来の書き込みはさぼらない。時間切れ時は洗いトリガー予約。
 */
function prMenuADTrialV1b() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=AD_v1b state=RUNNING', runId);
  try {
    var sellerPart = prMenuCkSellers({
      maxTargets: 1,
      toastTitle: '① A→D(店)',
      logLabel: 'A→D店',
      quietEmpty: true,
    });
    if (!sellerPart || !sellerPart.gate) {
      sellerPart = prMenuCkSellers({
        maxTargets: 1,
        forceSellerId: PR_SELLER_MORITA_,
        toastTitle: '① A→D(モリタ)',
        logLabel: 'A→Dモリタ固定',
      });
    }
    if (!sellerPart || !sellerPart.gate) {
      throw new Error('店調査できず（レ点1店またはモリタ行を確認）');
    }
    var g = sellerPart.gate;
    if (g.status === 'hold') {
      prAppendLog_([runId, 'A→D V1b', 'DONE', new Date(), '', '', '', 'hold ' + g.detail]);
      prMenuToast_('① A→D', '店holdのためB以降スキップ ' + g.detail, 12);
      return;
    }
    if (g.status === 'fail') {
      throw new Error('店調査失敗 ' + g.detail);
    }
    var passMakers = g.passMakers || [];
    var sumSh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
    var added = [];
    var mi;
    for (mi = 0; mi < passMakers.length; mi++) {
      var pm = String(passMakers[mi] || '').trim();
      if (!pm) continue;
      prEnsureMakerSumRowTop_(sumSh, pm, { ckTrue: false });
      added.push(pm);
    }
    var makerCap = 1;
    var rawCap = PropertiesService.getScriptProperties().getProperty('PR_AD_TRIAL_MAKER_CAP');
    if (rawCap != null && String(rawCap).trim() !== '') {
      var nCap = Number(String(rawCap).trim());
      if (nCap > 0) makerCap = Math.min(5, Math.floor(nCap));
    }
    // 実態付与後に選定: 製造メーカー→入口ブランド。非仕入先・要確認・Genericは除外
    var ranked = [];
    for (mi = 0; mi < added.length; mi++) {
      var jitta = prJittaForMaker_(added[mi]);
      if (!prMakerJittaWashEligible_(jitta)) {
        Logger.log(
          '[仕入れ検討①] runId=%s A→D skip maker=%s jitta=%s',
          runId,
          added[mi],
          jitta
        );
        continue;
      }
      ranked.push({
        parent: added[mi],
        jitta: jitta,
        pri: jitta === PR_JITTAI_MAKER_ ? 0 : 1,
      });
    }
    ranked.sort(function (a, b) {
      if (a.pri !== b.pri) return a.pri - b.pri;
      return String(a.parent).localeCompare(String(b.parent), 'ja');
    });
    var pick = [];
    for (mi = 0; mi < ranked.length && pick.length < makerCap; mi++) {
      pick.push(ranked[mi].parent);
    }
    if (!pick.length) {
      prAppendLog_([
        runId,
        'A→D V1b',
        'DONE',
        new Date(),
        '',
        '',
        '',
        'no_maker pass=' + passMakers.length + ' added=' + added.length + ' (実態除外後0)',
      ]);
      prMenuToast_('① A→D', '洗えるメーカーなし（Generic等は非仕入先）。①候補のみ', 12);
      return;
    }
    var parent = pick[0];
    var needles = prMakerWashNeedles_(parent, '');
    var entry = prMakerCatalogEntry_(parent, '', '');
    var locPick = prFindMakerSumRow_(parent);
    if (locPick && locPick.idxCk >= 0) {
      locPick.sh.getRange(locPick.sheetRow, locPick.idxCk + 1).setValue(true);
    }
    var wash = prRunMakerWashFull_({
      parent: parent,
      entry: entry,
      needles: needles,
      tag: 'AD_v1b',
      runId: runId,
      loopKeepa: true,
      armTrigger: true,
      quiet: true,
    });
    var washOk = wash && wash.ok && !wash.skip;
    var washComplete = washOk && prMakerWashDetailIsComplete_(wash.detail || '');
    if (!washComplete) {
      PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      // 洗い側で既に予約済みなら二重化しない。未武装ならここでトークン見積→予約。
      var afterMin = wash && wash.afterMin != null ? Number(wash.afterMin) : 0;
      if (!(wash && wash.armed)) {
        var wm = wash && wash.waitMin != null ? wash.waitMin : 5;
        var ri = wash && wash.refillIn != null ? wash.refillIn : '';
        afterMin = prScheduleWashFullTrigger_(prWashFullTriggerAfterMin_(wm, ri), {
          runId: runId,
          parent: parent,
          tag: 'AD_v1b',
        });
      }
      prAppendLog_([
        runId,
        'A→D V1b',
        'DONE',
        new Date(),
        '',
        '',
        '',
        'wash_pending parent=' +
          parent +
          ' afterMin=' +
          afterMin +
          ' ' +
          (wash && wash.detail ? wash.detail : ''),
      ]);
      prMenuToast_(
        '① A→D',
        '店・メーカー開始済。洗い未完了→約' +
          afterMin +
          '分後にトリガー再開。完走後に「5. D1→D2」 parent=' +
          parent,
        15
      );
      return;
    }
    var makerSet = {};
    makerSet[parent] = true;
    prMenuListAppendD1({ makerSet: makerSet });
    prMenuListFilterD2();
    prAppendLog_([
      runId,
      'A→D V1b',
      'DONE',
      new Date(),
      '',
      '',
      '',
      'complete parent=' + parent + ' makers_added=' + added.length,
    ]);
    prMenuToast_('① A→D', '完了 parent=' + parent + '（店→洗い完走→D1→D2）', 12);
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=AD_v1b state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'A→D V1b', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

function prMenuCkMakers() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=ck_maker state=RUNNING', runId);
  try {
    if (!prCkEnabled_(PR_CK_MAKER_ENABLED_PROP_)) {
      throw new Error(PR_CK_MAKER_ENABLED_PROP_ + '=false（切戻し中）');
    }
    var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      prMenuToast_('① レ点のメーカー', 'メーカーサマリにデータ行がありません', 8);
      return;
    }
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
    var idxName = prHeaderIndex_(hdr, 'メーカー');
    var idxJ = prHeaderIndex_(hdr, '実態');
    var idxPref = prHeaderIndex_(hdr, '入口推奨');
    var idxHit = prHeaderIndex_(hdr, '入口語ヒット');
    var idxDay = prHeaderIndex_(hdr, '最終調査日');
    var idxMemo = prHeaderIndex_(hdr, 'メモ');
    if (idxCk < 0 || idxName < 0) throw new Error('メーカーサマリに 調査CK または メーカー が無い');
    var vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var targets = [];
    var r;
    for (r = 0; r < vals.length; r++) {
      if (!prCkIsTrue_(vals[r][idxCk])) continue;
      var parent = String(vals[r][idxName] || '').trim();
      if (!parent) continue;
      targets.push({
        sheetRow: r + 2,
        parent: parent,
        jitta: idxJ >= 0 ? String(vals[r][idxJ] || '').trim() : '',
        pref: idxPref >= 0 ? String(vals[r][idxPref] || '').trim() : '',
        hit: idxHit >= 0 ? String(vals[r][idxHit] || '').trim() : '',
        memo: idxMemo >= 0 ? vals[r][idxMemo] : '',
      });
    }
    if (!targets.length) {
      prMenuToast_('① レ点のメーカー', 'レ点のメーカーがありません（競合DB メーカーサマリ A列）', 12);
      return;
    }
    var cap = prCkBatchCap_();
    var leftover = Math.max(0, targets.length - cap);
    var doneN = 0;
    var skipN = 0;
    var failN = 0;
    var lines = [];
    var i;
    for (i = 0; i < targets.length && i < cap; i++) {
      if (prGasTimeUp_()) {
        leftover = targets.length - i;
        lines.push('timeup leftover_makers=' + leftover);
        break;
      }
      var t = targets[i];
      var needles = prMakerWashNeedles_(t.parent, t.hit);
      var entry = prMakerCatalogEntry_(t.parent, t.pref, t.hit);
      Logger.log(
        '[仕入れ検討①] runId=%s ck_maker %s/%s parent=%s entry=%s',
        runId,
        i + 1,
        Math.min(targets.length, cap),
        t.parent,
        entry
      );
      var w = prRunMakerWash_({
        parent: t.parent,
        entry: entry,
        needles: needles,
        tag: 'ck_maker',
        quiet: true,
        runId: runId,
      });
      if (w.skip) {
        skipN++;
        if (idxMemo >= 0 && w.detail) {
          sh.getRange(t.sheetRow, idxMemo + 1).setValue(prAppendMemo_(t.memo, w.detail));
        }
      } else if (w.ok) {
        doneN++;
        if (prMakerWashFullEnabled_()) {
          if (prMakerWashDetailIsComplete_(w.detail)) {
            sh.getRange(t.sheetRow, idxCk + 1).setValue(false);
          }
        } else {
          sh.getRange(t.sheetRow, idxCk + 1).setValue(false);
          if (idxDay >= 0) sh.getRange(t.sheetRow, idxDay + 1).setValue(prYmdTokyo_());
        }
      } else {
        failN++;
        if (idxMemo >= 0 && w.detail) {
          sh.getRange(t.sheetRow, idxMemo + 1).setValue(prAppendMemo_(t.memo, w.detail));
        }
      }
      lines.push(t.parent + ':' + (w.skip ? 'skip' : w.ok ? 'done' : 'fail') + ' ' + w.detail);
    }
    var summary =
      'done=' +
      doneN +
      ' skip=' +
      skipN +
      ' fail=' +
      failN +
      ' leftover=' +
      leftover +
      ' cap=' +
      cap;
    Logger.log('[仕入れ検討①] runId=%s state=DONE %s', runId, summary);
    prAppendLog_([runId, 'レ点のメーカーを洗い', 'DONE', new Date(), '', targets.length, '', summary]);
    prMenuToast_('① レ点のメーカー', summary + ' ' + lines.join(' ').slice(0, 200));
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'レ点のメーカーを洗い', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/**
 * レ点メーカーリサーチ→商品リストアップ（メーカー洗い→D1まで1メニュー）
 * - メーカサマリの調査CK=TRUE行を対象に洗い
 * - （任意）promptで追加メーカー名を受け取り、メーカーサマリ先頭に行insertして調査CK=TRUEにする
 * - D1は「そのメーカー集合」かつ「発見経路=ck_maker_catalog」のみで生成（混入抑止）
 */
function prMenuMakerResearchAndListUp() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=maker_d1 state=RUNNING', runId);
  try {
    var ui = SpreadsheetApp.getUi();

    // メーカーサマリを確保し、調査CK列/チェック枠を成立させる
    var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      prMenuToast_('① レ点メーカーリサーチ', 'メーカーサマリにデータ行がありません', 8);
      return;
    }

    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
    var idxMaker = prHeaderIndex_(hdr, 'メーカー');
    var idxPref = prHeaderIndex_(hdr, '入口推奨');
    var idxHit = prHeaderIndex_(hdr, '入口語ヒット');
    var idxJ = prHeaderIndex_(hdr, '実態');
    var idxDay = prHeaderIndex_(hdr, '最終調査日');
    var idxMemo = prHeaderIndex_(hdr, 'メモ');
    if (idxCk < 0 || idxMaker < 0) throw new Error('メーカーサマリに 調査CK または メーカー が無い');

    function readMakerInfo_() {
      var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
      var info = {};
      var makerOrder = [];
      var r;
      for (r = 0; r < vals.length; r++) {
        var parent = String(vals[r][idxMaker] || '').trim();
        if (!parent) continue;
        var key = parent;
        if (info[key]) continue;
        info[key] = {
          sheetRow: r + 2,
          parent: parent,
          jitta: idxJ >= 0 ? String(vals[r][idxJ] || '').trim() : '',
          pref: idxPref >= 0 ? String(vals[r][idxPref] || '').trim() : '',
          hit: idxHit >= 0 ? String(vals[r][idxHit] || '').trim() : '',
          memo: idxMemo >= 0 ? vals[r][idxMemo] : '',
        };
        makerOrder.push(parent);
      }
      return { info: info, makerOrder: makerOrder };
    }

    // 既存のレ点（調査CK=TRUE）を選択セットにする
    var vals0 = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var makerSet = {};
    var makerList = [];
    var r0;
    for (r0 = 0; r0 < vals0.length; r0++) {
      if (!prCkIsTrue_(vals0[r0][idxCk])) continue;
      var m0 = String(vals0[r0][idxMaker] || '').trim();
      if (!m0) continue;
      makerSet[m0] = true;
      makerList.push(m0);
    }

    // 既存の調査CK=TRUEがある場合は、promptを出さずそのメーカーのみを調査する
    var extraMakers = [];
    if (!Object.keys(makerSet).length) {
      var extraText = '';
      try {
        var resp = ui.prompt(
          'レ点メーカーリサーチ→商品リストアップ',
          '調査CK=TRUE のメーカーがありません。追加するメーカー名（改行/カンマ/| 区切り）を入力してください。',
          ui.ButtonSet.OK_CANCEL
        );
        if (resp && resp.getSelectedButton && resp.getSelectedButton() === ui.Button.OK) {
          extraText = String(resp.getResponseText() || '');
        }
      } catch (e1) {
        // promptをキャンセル/非表示の場合は追加なしで進める
      }
      extraMakers = prSplitTok_(extraText);
    }

    if (extraMakers.length) {
      // 現時点のメーカー行情報を使って、無ければ先頭へ挿入
      var mInfoBefore = readMakerInfo_();
      for (var i = 0; i < extraMakers.length; i++) {
        var m = extraMakers[i];
        if (!m) continue;
        if (makerSet[m]) continue;
        if (mInfoBefore.info[m]) {
          makerSet[m] = true;
          makerList.push(m);
          continue;
        }
        // 無いメーカーは先頭データ直下（2行目）へinsertして調査CK=TRUE（同名は足さない）
        prEnsureMakerSumRowTop_(sh, m, { ckTrue: true });
        makerSet[m] = true;
        makerList.push(m);
      }
    }

    if (!Object.keys(makerSet).length) {
      prMenuToast_('① レ点メーカーリサーチ', '調査CK=TRUE のメーカーがありません', 10);
      return;
    }

    // 挿入により行番号が変わるので、最終的に再読取
    var mInfo = readMakerInfo_();
    // makerList（入力/既存順）から、実在行のみ抽出（重複除去）
    var chosen = [];
    var chosenSeen = {};
    for (var j = 0; j < makerList.length; j++) {
      var mk = makerList[j];
      if (!makerSet[mk]) continue;
      if (chosenSeen[mk]) continue;
      chosenSeen[mk] = true;
      if (mInfo.info[mk]) chosen.push(mInfo.info[mk]);
    }
    // 万一読み取りの都合で順序が欠けた場合は、残りを補完
    for (var k = 0; k < mInfo.makerOrder.length; k++) {
      var mk2 = mInfo.makerOrder[k];
      if (makerSet[mk2] && !chosenSeen[mk2] && mInfo.info[mk2]) {
        chosenSeen[mk2] = true;
        chosen.push(mInfo.info[mk2]);
      }
    }

    var cap = prCkBatchCap_();
    var leftover = Math.max(0, chosen.length - cap);
    var doneN = 0;
    var skipN = 0;
    var failN = 0;
    var lines = [];

    for (var t = 0; t < chosen.length && t < cap; t++) {
      if (prGasTimeUp_()) {
        leftover = chosen.length - t;
        lines.push('timeup leftover_makers=' + leftover);
        break;
      }
      var one = chosen[t];
      var needles = prMakerWashNeedles_(one.parent, one.hit);
      var entry = prMakerCatalogEntry_(one.parent, one.pref, one.hit);

      Logger.log(
        '[仕入れ検討①] runId=%s maker_d1 %s/%s parent=%s entry=%s',
        runId,
        t + 1,
        Math.min(chosen.length, cap),
        one.parent,
        entry
      );

      var w = prRunMakerWash_({
        parent: one.parent,
        entry: entry,
        needles: needles,
        tag: 'ck_maker',
        quiet: true,
        runId: runId,
      });
      if (w.skip) {
        skipN++;
        if (idxMemo >= 0 && w.detail) {
          sh.getRange(one.sheetRow, idxMemo + 1).setValue(prAppendMemo_(one.memo, w.detail));
        }
      } else if (w.ok) {
        doneN++;
        if (prMakerWashFullEnabled_()) {
          if (prMakerWashDetailIsComplete_(w.detail)) {
            sh.getRange(one.sheetRow, idxCk + 1).setValue(false);
          }
        } else {
          sh.getRange(one.sheetRow, idxCk + 1).setValue(false);
          if (idxDay >= 0) sh.getRange(one.sheetRow, idxDay + 1).setValue(prYmdTokyo_());
        }
      } else {
        failN++;
        if (idxMemo >= 0 && w.detail) {
          sh.getRange(one.sheetRow, idxMemo + 1).setValue(prAppendMemo_(one.memo, w.detail));
        }
      }
      lines.push(one.parent + ':' + (w.skip ? 'skip' : w.ok ? 'done' : 'fail') + ' ' + w.detail);
    }

    var summary = 'done=' + doneN + ' skip=' + skipN + ' fail=' + failN + ' leftover=' + leftover + ' cap=' + cap;
    Logger.log('[仕入れ検討①] runId=%s state=DONE %s', runId, summary);

    // D1: 選択メーカーだけを対象にして先頭挿入（混入抑止）
    prMenuListAppendD1({
      makerSet: makerSet,
      discoveryTag: 'ck_maker_catalog',
    });

    prAppendLog_([runId, 'レ点メーカー→D1', 'DONE', new Date(), '', chosen.length, '', summary]);
    prMenuToast_('① レ点メーカー→商品リスト', summary + ' ' + lines.join(' ').slice(0, 200));
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'レ点メーカー→D1', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/**
 * C3 メニュー（CYCLE §6.3）。本線メニューから実行。実体は prMenuAutoMakerResearchAndListUp。
 */
function prMenuCycleC3MakerWash() {
  prCyclePauseForOtherMenu_('C3');
  return prMenuAutoMakerResearchAndListUp({});
}

/** C3 + 洗い未完なら完走トリガー自動開始（1メニューで無人乾燥）。 */
function prMenuCycleC3MakerWashWithTrigger() {
  PropertiesService.getScriptProperties().deleteProperty(PR_CYCLE_HEARTBEAT_PAUSED_PROP_);
  prCycleClearResume01_();
  return prMenuAutoMakerResearchAndListUp({ armTriggerAfter: true });
}

/**
 * 未調査メーカーを自動選定（C3）→メーカー洗い→D1（レ点不要）。
 * 選定: 洗い途中優先・規模帯 小→中→大→特大（不明は列挙用・洗い選定外）。Prop PR_CYCLE_BAND_SELECT=false で旧ティア。
 * @param {{armTriggerAfter?:boolean}=} opt armTriggerAfter … 洗い途中なら P3 トリガー予約
 */
function prMenuAutoMakerResearchAndListUp(opt) {
  opt = opt || {};
  var armTriggerAfter = !!opt.armTriggerAfter;
  var logLabel = armTriggerAfter ? 'C3+T 自動選定→洗い→D1→完走トリガー' : 'C3 自動選定→洗い→D1';
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  prDetectWashTriggerMissed_(runId, armTriggerAfter ? 'C3+T_start' : 'C3_start');
  Logger.log(
    '[仕入れ検討①] runId=%s step=C3 state=RUNNING functionName=%s armTriggerAfter=%s',
    runId,
    armTriggerAfter ? 'prMenuCycleC3MakerWashWithTrigger' : 'prMenuCycleC3MakerWash',
    armTriggerAfter ? 'Y' : 'N'
  );
  try {
    if (!prAutoMakerEnabled_()) {
      throw new Error(PR_AUTO_MAKER_ENABLED_PROP_ + '=false（切戻し中）');
    }
    if (!prMakerWashEnabled_()) {
      throw new Error(PR_MAKER_WASH_ENABLED_PROP_ + '=false（切戻し中）');
    }
    var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      prMenuToast_('C3 自動選定', 'メーカーサマリにデータ行がありません', 8);
      return;
    }
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxDay = prHeaderIndex_(hdr, '最終調査日');
    var idxMemo = prHeaderIndex_(hdr, 'メモ');
    var vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var cap = prAutoMakerCap_();
    var chosen = prPickAutoMakerCandidates_(vals, hdr, cap);
    if (!chosen.length) {
      prMenuToast_(
        'C3 自動選定',
        '対象なし（洗い途中なし／' +
          prAutoMakerRewalkDays_() +
          '日以内調査済み／非仕入先／要確認／調査CK付きを除外）',
        12
      );
      prAppendLog_([runId, logLabel, 'DONE', new Date(), '', 0, '', 'pick=0']);
      return;
    }

    var makerSet = {};
    var doneN = 0;
    var skipN = 0;
    var failN = 0;
    var lines = [];
    var t;
    for (t = 0; t < chosen.length; t++) {
      if (prGasTimeUp_()) {
        lines.push('timeup at=' + (t + 1));
        break;
      }
      var one = chosen[t];
      makerSet[one.parent] = true;
      var needles = prMakerWashNeedles_(one.parent, one.hit);
      var entry = prMakerCatalogEntry_(one.parent, one.pref, one.hit);
      Logger.log(
        '[仕入れ検討①] runId=%s auto_maker %s/%s parent=%s entry=%s via=%s tier=%s',
        runId,
        t + 1,
        chosen.length,
        one.parent,
        entry,
        one.via || '',
        one.tier || ''
      );
      var w = prRunMakerWash_({
        parent: one.parent,
        entry: entry,
        needles: needles,
        tag: 'auto_maker',
        quiet: true,
        runId: runId,
        armTrigger: armTriggerAfter,
      });
      if (w.skip) {
        skipN++;
        if (idxMemo >= 0 && w.detail) {
          sh.getRange(one.sheetRow, idxMemo + 1).setValue(prAppendMemo_(one.memo, w.detail));
        }
      } else if (w.ok) {
        doneN++;
        // 完走本線: 最終調査日は prWashFullKeepaSlice_ 内のみ
        if (!prMakerWashFullEnabled_() && idxDay >= 0) {
          sh.getRange(one.sheetRow, idxDay + 1).setValue(prYmdTokyo_());
        }
      } else {
        failN++;
        if (idxMemo >= 0 && w.detail) {
          sh.getRange(one.sheetRow, idxMemo + 1).setValue(prAppendMemo_(one.memo, w.detail));
        }
      }
      lines.push(one.parent + ':' + (w.skip ? 'skip' : w.ok ? 'done' : 'fail') + ' ' + w.detail);
    }

    var summary =
      'pick=' +
      chosen.length +
      ' done=' +
      doneN +
      ' skip=' +
      skipN +
      ' fail=' +
      failN +
      ' cap=' +
      cap +
      ' rewalkDays=' +
      prAutoMakerRewalkDays_();
    if (chosen[0] && chosen[0].via) summary += ' via=' + chosen[0].via;
    if (chosen[0] && chosen[0].parent) summary += ' parent=' + chosen[0].parent;
    Logger.log('[仕入れ検討①] runId=%s step=C3 state=DONE %s', runId, summary);

    prMenuListAppendD1({
      makerSet: makerSet,
      discoveryTag: 'auto_maker_catalog',
    });

    if (armTriggerAfter) {
      var trig = prArmWashFullTriggerForContinue_({ runId: runId, tag: 'C3+T' });
      if (trig && trig.armed) {
        summary +=
          ' trigger_armed parent=' + trig.parent + ' afterMin=' + trig.afterMin;
      } else if (trig && trig.reason) {
        summary += ' trigger_skip=' + trig.reason;
      }
    }

    prAppendLog_([runId, logLabel, 'DONE', new Date(), '', chosen.length, '', summary]);
    if (armTriggerAfter) {
      var toastTitle = 'C3+T 自動選定';
      var toastBody = summary + ' ' + lines.join(' ').slice(0, 160);
      if (summary.indexOf('trigger_armed') >= 0) {
        toastBody += ' …約1分後から完走まで自動（停止=99・トリガー停止）';
      } else if (summary.indexOf('trigger_skip=no_continue_target') >= 0) {
        toastBody += ' …洗い完走済み（C4自動）またはトリガー不要';
      }
      prMenuToast_(toastTitle, toastBody, 14);
    } else {
      prMenuToast_('C3 自動選定', summary + ' ' + lines.join(' ').slice(0, 200));
    }
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, logLabel, 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

function prCycleC4Enabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_CYCLE_C4_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prCountCycleC4Triggers_() {
  var list = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === PR_CYCLE_C4_TRIGGER_FN_) n++;
  }
  return n;
}

/**
 * @param {{runId?:string, reason?:string, log?:boolean}=} opt
 * @return {number}
 */
function prClearCycleC4Triggers_(opt) {
  opt = opt || {};
  var before = prCountCycleC4Triggers_();
  var list = ScriptApp.getProjectTriggers();
  var deleted = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === PR_CYCLE_C4_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(list[i]);
      deleted++;
    }
  }
  if (opt.log) {
    var rid =
      opt.runId ||
      'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    prAppendLog_([
      rid,
      'C4トリガー解除',
      'DONE',
      new Date(),
      deleted,
      '',
      '',
      'c4_trigger_cleared reason=' +
        String(opt.reason || 'clear') +
        ' deleted=' +
        deleted +
        ' before=' +
        before +
        ' after=' +
        prCountCycleC4Triggers_() +
        ' now=' +
        prNowFetchedAt_(),
    ]);
  }
  return deleted;
}

/**
 * 洗い完走と別実行で C4（D1→D2→心拍）を走らせる。afterMin 既定1分。
 * @param {number} afterMin
 * @param {{runId?:string, parent?:string, tag?:string}=} logOpt
 * @return {number}
 */
function prScheduleCycleC4Trigger_(afterMin, logOpt) {
  var mins = Number(afterMin);
  if (!(mins > 0)) mins = PR_CYCLE_C4_TRIGGER_AFTER_DEFAULT_;
  mins = Math.floor(mins);
  var rid =
    (logOpt && logOpt.runId) ||
    'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var tag = (logOpt && logOpt.tag) || 'cycle_c4';
  var parent = String((logOpt && logOpt.parent) || '').trim();
  if (!parent) return 0;
  if (prCycleD2OverflowActive_()) {
    var deferRid =
      (logOpt && logOpt.runId) ||
      'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    prAppendLog_([
      deferRid,
      tag + ' C4トリガー延期',
      'DONE',
      new Date(),
      '',
      parent,
      '',
      'deferred_d2_overflow parent=' + parent,
    ]);
    if (prCycleHeartbeatEnabled_()) {
      prCycleArmHeartbeatAfterIdle_(deferRid, 'c4_deferred_overflow');
    }
    return 0;
  }
  var props = PropertiesService.getScriptProperties();
  prClearCycleC4Triggers_({ reason: 'reschedule:' + tag, runId: rid });
  props.setProperty(PR_CYCLE_C4_PENDING_PARENT_PROP_, parent);
  props.setProperty(PR_CYCLE_C4_PENDING_RUNID_PROP_, rid);
  props.setProperty(PR_CYCLE_C4_PENDING_TAG_PROP_, tag);
  var expectMs = new Date().getTime() + mins * 60 * 1000;
  props.setProperty(PR_CYCLE_C4_TRIGGER_EXPECT_MS_PROP_, String(expectMs));
  ScriptApp.newTrigger(PR_CYCLE_C4_TRIGGER_FN_).timeBased().after(mins * 60 * 1000).create();
  var expectedAt = Utilities.formatDate(
    new Date(expectMs),
    'Asia/Tokyo',
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
  prAppendLog_([
    rid,
    tag + ' C4トリガー予約',
    'DONE',
    new Date(),
    mins,
    parent,
    '',
    'c4_trigger_armed fn=' +
      PR_CYCLE_C4_TRIGGER_FN_ +
      ' afterMin=' +
      mins +
      ' expectedAt=' +
      expectedAt +
      ' parent=' +
      parent +
      ' triggerN=' +
      prCountCycleC4Triggers_() +
      ' now=' +
      prNowFetchedAt_(),
  ]);
  return mins;
}

/** 時間主導: 洗い完走後の C4 完走連鎖（UIなし・独立30分枠）。 */
function prRunCycleC4FromTrigger() {
  prGasStart_();
  var fireRunId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var props = PropertiesService.getScriptProperties();
  var parent = String(props.getProperty(PR_CYCLE_C4_PENDING_PARENT_PROP_) || '').trim();
  var chainRunId =
    String(props.getProperty(PR_CYCLE_C4_PENDING_RUNID_PROP_) || '').trim() || fireRunId;
  var tag = String(props.getProperty(PR_CYCLE_C4_PENDING_TAG_PROP_) || 'cycle_c4').trim();
  var expectMs = Number(props.getProperty(PR_CYCLE_C4_TRIGGER_EXPECT_MS_PROP_) || 0);
  var lagMin =
    expectMs > 0 ? Math.round((new Date().getTime() - expectMs) / 60000) : '';
  var expectedAt =
    expectMs > 0
      ? Utilities.formatDate(new Date(expectMs), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX")
      : '';
  prAppendLog_([
    chainRunId,
    'C4 trigger',
    'RUNNING',
    new Date(),
    lagMin,
    parent,
    '',
    'fire now=' +
      prNowFetchedAt_() +
      ' expectedAt=' +
      expectedAt +
      ' lagMin=' +
      lagMin +
      ' parent=' +
      parent +
      ' chainRunId=' +
      chainRunId +
      ' tag=' +
      tag +
      ' triggerN=' +
      prCountCycleC4Triggers_(),
  ]);
  if (prCycleHeartbeatPaused_()) {
    prClearCycleC4Triggers_({ reason: 'paused', runId: chainRunId });
    if (parent && prCycleC4Enabled_()) {
      prScheduleCycleC4Trigger_(5, { runId: chainRunId, parent: parent, tag: tag + '_paused' });
    }
    prAppendLog_([chainRunId, 'C4 trigger', 'DONE', new Date(), 5, parent, '', 'paused re-arm']);
    prCycleEnsureLiveness_(chainRunId, 'c4_paused');
    return;
  }
  props.deleteProperty(PR_CYCLE_C4_TRIGGER_EXPECT_MS_PROP_);
  props.deleteProperty(PR_CYCLE_C4_PENDING_PARENT_PROP_);
  props.deleteProperty(PR_CYCLE_C4_PENDING_RUNID_PROP_);
  props.deleteProperty(PR_CYCLE_C4_PENDING_TAG_PROP_);
  prClearCycleC4Triggers_({ reason: 'fired:' + fireRunId, runId: chainRunId });
  if (!prCycleC4Enabled_()) {
    prAppendLog_([chainRunId, 'C4 trigger', 'DONE', new Date(), '', '', '', 'c4_disabled']);
    prCycleEnsureLiveness_(chainRunId, 'c4_disabled');
    return;
  }
  if (!parent) {
    prAppendLog_([chainRunId, 'C4 trigger', 'DONE', new Date(), '', '', '', 'no_pending_parent']);
    prCycleEnsureLiveness_(chainRunId, 'c4_no_parent');
    return;
  }
  prWashInflightSet_(chainRunId, 'c4');
  try {
    prRunCycleC4Chain_({ parent: parent, runId: chainRunId, tag: tag, quiet: true });
  } catch (c4e) {
    var c4msg = prRedactKey_(String((c4e && c4e.message) || c4e));
    Logger.log('[仕入れ検討①] runId=%s step=C4 trigger state=FAILED %s', chainRunId, c4msg);
    prAppendLog_([chainRunId, 'C4 完走連鎖', 'FAILED', new Date(), '', '', '', c4msg.slice(0, 200)]);
    if (prCycleC4Enabled_() && parent) {
      prScheduleCycleC4Trigger_(5, { runId: chainRunId, parent: parent, tag: tag + '_retry' });
    }
  } finally {
    prWashInflightClear_('c4_exit:' + chainRunId);
    prCycleEnsureLiveness_(chainRunId, 'c4_finally');
  }
}

/**
 * C4 完走連鎖（CYCLE §6.4）。洗い完走後: D1→D2。D成功時は心拍へ（次社選定は心拍）。D失敗時も心拍へ。
 * @return {{dOk:boolean, armNext:boolean, nextMaker:string, detail:string}}
 */
function prRunCycleC4Chain_(opt) {
  opt = opt || {};
  if (!prCycleC4Enabled_()) {
    return { dOk: false, armNext: false, nextMaker: '', detail: 'c4_disabled' };
  }
  var parent = String(opt.parent || '').trim();
  if (!parent) {
    return { dOk: false, armNext: false, nextMaker: '', detail: 'parent_empty' };
  }
  var runId =
    opt.runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var quiet = !!opt.quiet;
  var tag = String(opt.tag || 'cycle_c4');
  Logger.log('[仕入れ検討①] runId=%s step=C4 state=RUNNING parent=%s', runId, parent);
  PropertiesService.getScriptProperties().setProperty(PR_CYCLE_C4_LAST_PARENT_PROP_, parent);
  var makerSet = {};
  makerSet[parent] = true;
  var d1Ok = true;
  var d1Msg = '';
  try {
    prMenuListAppendD1({ makerSet: makerSet, discoveryTag: tag + '_catalog' });
  } catch (e1) {
    d1Ok = false;
    d1Msg = prRedactKey_(String((e1 && e1.message) || e1)).slice(0, 120);
  }
  var d2Ok = false;
  var d2Left = 0;
  var d2Detail = '';
  if (d1Ok) {
    prWashInflightPhase_('c4_d2');
    try {
      var d2r = prRunListFilterD2_({ runId: runId, quiet: true, skipToast: true });
      d2Ok = !!d2r.ok;
      d2Left = Number(d2r.leftover) || 0;
      d2Detail = String(d2r.detail || '');
    } catch (e2) {
      d2Ok = false;
      d2Detail = prRedactKey_(String((e2 && e2.message) || e2)).slice(0, 120);
    }
  }
  var dOk = d1Ok && d2Ok && d2Left <= 0;
  if (!dOk) {
    var reason = !d1Ok ? 'd1_fail' : !d2Ok ? 'd2_fail' : 'd2_leftover=' + d2Left;
    if (d1Ok && d2Ok && d2Left > 0) {
      reason = 'd2_overflow';
      prCyclePipeListAdd_(PR_CYCLE_D2_OVERFLOW_PROP_, parent);
      try {
        var capOv = prListFilterCap_();
        var unclOv = prListCountUnclassified_();
        prSendD2OverflowMail_({
          runId: runId,
          parent: parent,
          leftover: d2Left,
          unclassified: unclOv,
          cap: capOv,
          detail: d2Detail,
        });
      } catch (eOvMail) {
        Logger.log(
          '[仕入れ検討①] runId=%s d2_overflow_mail skip %s',
          runId,
          String((eOvMail && eOvMail.message) || eOvMail).slice(0, 120)
        );
      }
    }
    var stopSummary =
      'stopped parent=' + parent + ' ' + reason + (d1Msg ? ' d1=' + d1Msg : '') + (d2Detail ? ' d2=' + d2Detail : '');
    if (prCycleHeartbeatEnabled_()) {
      var hbFail = prCycleArmHeartbeatAfterIdle_(runId, 'C4_d_fail_heartbeat');
      stopSummary += ' heartbeat_armed afterMin=' + hbFail;
    }
    prAppendLog_([runId, 'C4 完走連鎖', 'DONE', new Date(), '', '', '', stopSummary.slice(0, 200)]);
    Logger.log('[仕入れ検討①] runId=%s step=C4 state=STOPPED %s', runId, stopSummary);
    if (!quiet) prMenuToast_('C4 完走連鎖', 'D失敗→心拍へ ' + parent + ' (' + reason + ')', 12);
    return { dOk: false, armNext: false, nextMaker: '', detail: stopSummary };
  }
  // 小・中完走後の店挟みを次社洗いより優先（§6.3bis）
  var shopPend = prCycleShopAfterSmParent_();
  if (shopPend) {
    var shopSummary = 'd_ok shop_after_sm=' + shopPend + ' parent=' + parent;
    if (prCycleHeartbeatEnabled_()) {
      var hbShop = prCycleArmHeartbeatAfterIdle_(runId, 'C4_shop_after_sm');
      shopSummary += ' heartbeat_armed afterMin=' + hbShop;
    }
    prAppendLog_([runId, 'C4 完走連鎖', 'DONE', new Date(), '', '', '', shopSummary.slice(0, 200)]);
    if (!quiet) prMenuToast_('C4 完走連鎖', 'D OK。店レーンへ（案B） ' + shopPend, 12);
    prCycleMarkC4D1Done_(parent);
    return { dOk: true, armNext: false, nextMaker: '', detail: shopSummary };
  }
  // 次社は心拍に一本化（before_large／Catalog先行／店レーンと二重管理しない）
  var hbSummary = 'd_ok next_via_heartbeat parent=' + parent;
  if (prCycleHeartbeatEnabled_()) {
    var hbNext = prCycleArmHeartbeatAfterIdle_(runId, 'C4_next_heartbeat');
    hbSummary += ' heartbeat_armed afterMin=' + hbNext;
  } else {
    hbSummary += ' heartbeat_off';
  }
  prAppendLog_([runId, 'C4 完走連鎖', 'DONE', new Date(), '', '', '', hbSummary.slice(0, 200)]);
  Logger.log('[仕入れ検討①] runId=%s step=C4 state=DONE %s', runId, hbSummary);
  if (!quiet) {
    prMenuToast_(
      'C4 完走連鎖',
      '完走→D1→D2 OK。次は心拍で選定' +
        (hbSummary.indexOf('heartbeat_armed') >= 0 ? '（約1分後）' : ''),
      12
    );
  }
  prCycleMarkC4D1Done_(parent);
  return { dOk: true, armNext: false, nextMaker: '', detail: hbSummary };
}

/** C4 メニュー: 完走済み（D1 未）parent で D1→D2→心拍。 */
function prMenuCycleC4AfterComplete() {
  prGasStart_();
  var props = PropertiesService.getScriptProperties();
  var needs = prCycleListNeedsD1_();
  var parent = '';
  var last = String(props.getProperty(PR_CYCLE_C4_LAST_PARENT_PROP_) || '').trim();
  if (last && needs.indexOf(last) >= 0) parent = last;
  if (!parent && needs.length) parent = needs[0];
  if (!parent) {
    prMenuFailAlert_('完走済みで D1 未のメーカーなし（洗い状態=完了）');
    return;
  }
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  prRunCycleC4Chain_({ parent: parent, runId: runId, tag: 'C4_menu', quiet: false });
}

/** D2 backlog 解消後に C4 を再開（overflow 停止の救済）。 */
function prMenuCycleC4ResumeAfterD2Clear() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var uncl = prListCountUnclassified_();
  if (uncl > 0) {
    prMenuFailAlert_(
      '未判定 ASIN が ' +
        uncl +
        ' 件残っています。先にメニュー 99→6「D2だけ」で ①ログ leftover=0 にしてください。'
    );
    return;
  }
  var props = PropertiesService.getScriptProperties();
  var pending = prCycleD2OverflowGet_();
  var parent = '';
  var last = String(props.getProperty(PR_CYCLE_C4_LAST_PARENT_PROP_) || '').trim();
  if (last && pending.indexOf(last) >= 0) parent = last;
  if (!parent && pending.length) parent = pending[0];
  if (!parent && last) {
    var needsLast = prCycleListNeedsD1_();
    if (needsLast.indexOf(last) >= 0) parent = last;
  }
  if (!parent) {
    var needs = prCycleListNeedsD1_();
    if (needs.length) parent = needs[0];
  }
  if (!parent) {
    prCycleD2OverflowClear_();
    prMenuFailAlert_('C4 再開対象のメーカーなし（needs 空）。overflow フラグだけ解除しました。');
    return;
  }
  prCycleD2OverflowClear_();
  Logger.log('[仕入れ検討①] runId=%s step=D2_overflow_resume state=RUNNING parent=%s', runId, parent);
  prRunCycleC4Chain_({ parent: parent, runId: runId, tag: 'd2_overflow_menu_c4', quiet: false });
}

/**
 * history=0&stats=90&offers=20。90日内の Keepaフルを先に使う。
 * GET した分だけ Keepaフルへ追記（csv[]落とす・目的=リサーチ・指紋同じなら非書。ただし薄い→offers補強は上書き）。
 * @return {{products:Object[], fullPlus:number, need:number, skip:number, tokensConsumed:number, tokensLeft:*, refillRate:number, refillIn:*, waitMin:number}}
 */
function prKeepaProductStats90_(asins) {
  var packed = prSplitKeepaFull_(asins || []);
  Logger.log('[仕入れ検討①] Keepaフル hit=%s need_get=%s', packed.have.length, packed.need.length);
  var fetchedPack = packed.need.length
    ? prKeepaProductStats90Fetch_(packed.need)
    : {
        products: [],
        tokensConsumed: 0,
        tokensLeft: '',
        refillRate: prKeepaRefillRate_(),
        refillIn: '',
        rateLimited: false,
      };
  var fetched = fetchedPack.products || [];
  var fullPlus = prAppendKeepaFull_(fetched);
  var rate =
    fetchedPack.refillRate > 0 ? fetchedPack.refillRate : prKeepaRefillRate_();
  var waitMin = prKeepaWaitMinutes_(fetchedPack.tokensConsumed, rate);
  if (fetchedPack.rateLimited && Number(fetchedPack.refillIn) > 0) {
    waitMin = Math.max(waitMin, Math.ceil(Number(fetchedPack.refillIn) / 60000) + 1);
  }
  Logger.log(
    '[仕入れ検討①] Keepaフル append=%s consumed=%s left=%s refillRate=%s refillIn=%s waitMin=%s rateLimited=%s',
    fullPlus,
    fetchedPack.tokensConsumed,
    fetchedPack.tokensLeft,
    rate,
    fetchedPack.refillIn,
    waitMin,
    fetchedPack.rateLimited ? 'Y' : 'N'
  );
  return {
    products: packed.have.concat(fetched),
    fullPlus: fullPlus,
    need: packed.need.length,
    skip: packed.have.length,
    tokensConsumed: fetchedPack.tokensConsumed || 0,
    tokensLeft: fetchedPack.tokensLeft,
    refillRate: rate,
    refillIn: fetchedPack.refillIn,
    waitMin: waitMin,
    rateLimited: !!fetchedPack.rateLimited,
  };
}

/** Property PR_KEEPA_TWO_STAGE_GATE。未設定=ON。false で従来の一発 offers=20。 */
function prKeepaTwoStageGateEnabled_() {
  return prCkEnabled_(PR_KEEPA_TWO_STAGE_GATE_PROP_);
}

/**
 * offers=20 で取った形跡があるか。
 * Keepaフルの生JSON は offers を落とす（prProductToFullRec_）ので、
 * `offers` 配列ではなく offers 指定時だけ入る stats の痕跡で判定する。
 * @param {!Object} p
 */
function prKeepaHasOffersInfo_(p) {
  if (p && p.offers && p.offers.length) return true;
  var st = p && p.stats ? p.stats : null;
  if (!st) return false;
  if (String(st.buyBoxSellerId || '').trim()) return true;
  if (Number(st.offerCountFBA) >= 0) return true;
  if (Number(st.offerCountFBM) >= 0) return true;
  return false;
}

/**
 * 洗い用の商品取得（2段ゲート）。`prKeepaProductStats90_` と同じ形で返す。
 *
 * 1段目: offers なし（1トークン/件）。門は stats だけで判定できるので、落ちはここで確定。
 * 2段目: 門を通ったものだけ offers=20（約5トークン/件）で取り直す。
 *
 * Keepaフルへ書くのは「落ちた薄い行」と「通った太い行」だけ。
 * 通ったのに2段目が取れなかった ASIN は `notReady` で返し、Keepa済にしない（次スライスで再試行）。
 * 薄い行を通過のまま確定させると、店レーン・D1 が使う BuyBox が欠けたまま90日キャッシュされる。
 *
 * @param {!Array<string>} asins
 * @param {string} runId
 * @return {{products:!Array<!Object>, fullPlus:number, need:number, skip:number,
 *           tokensConsumed:number, tokensLeft:(number|string), refillRate:number,
 *           refillIn:(number|string), waitMin:number, rateLimited:boolean,
 *           notReady:!Object<string,boolean>, stage1:number, stage2:number}}
 */
function prKeepaWashGateProducts_(asins, runId) {
  var packed = prSplitKeepaFull_(asins || []);
  var rate = prKeepaRefillRate_();
  var consumed = 0;
  var tokensLeft = '';
  var refillIn = '';
  var rateLimited = false;
  var i;

  function absorb_(pack) {
    consumed += Number(pack.tokensConsumed) || 0;
    if (pack.tokensLeft != null && pack.tokensLeft !== '') tokensLeft = pack.tokensLeft;
    if (Number(pack.refillRate) > 0) rate = Number(pack.refillRate);
    if (pack.refillIn != null && pack.refillIn !== '') refillIn = pack.refillIn;
    if (pack.rateLimited) rateLimited = true;
  }

  // 1段目: offers なし
  var s1 = { products: [] };
  if (packed.need.length) {
    s1 = prKeepaProductStats90Fetch_(packed.need, { offers: 0 });
    absorb_(s1);
  }
  var fetched1 = {};
  for (i = 0; i < (s1.products || []).length; i++) {
    var a1 = String(s1.products[i].asin || '').toUpperCase();
    if (a1) fetched1[a1] = true;
  }
  var all = packed.have.concat(s1.products || []);

  // 門の仮判定（消費0）。通過ぶんだけ offers を取り直す
  var prof = prReadProfile_();
  var passSet = {};
  var need2 = [];
  for (i = 0; i < all.length; i++) {
    var a = String(all[i].asin || '').toUpperCase();
    if (!a) continue;
    if (prGateProductRow_(all[i], prof, '', runId)[8] !== '通過') continue;
    passSet[a] = true;
    if (!prKeepaHasOffersInfo_(all[i])) need2.push(a);
  }

  // 2段目: 通過だけ offers=20
  var s2 = { products: [] };
  if (need2.length && !rateLimited) {
    s2 = prKeepaProductStats90Fetch_(need2, { offers: 20 });
    absorb_(s2);
  }
  var fat = {};
  for (i = 0; i < (s2.products || []).length; i++) {
    var a2 = String(s2.products[i].asin || '').toUpperCase();
    if (a2) fat[a2] = s2.products[i];
  }

  var products = [];
  var writeList = [];
  var notReady = {};
  var thinPass = 0;
  for (i = 0; i < all.length; i++) {
    var key = String(all[i].asin || '').toUpperCase();
    if (!key) continue;
    if (fat[key]) {
      products.push(fat[key]);
      writeList.push(fat[key]);
      continue;
    }
    if (passSet[key] && !prKeepaHasOffersInfo_(all[i])) {
      if (rateLimited) {
        // トークン切れで2段目に届かなかっただけ。確定させず次スライスで取り直す
        notReady[key] = true;
        continue;
      }
      // 429 でもないのに offers が来ない。永久保留を避けて薄い行で確定させる。
      // BuyBox は門通過ASIN向けの補強 GET（prKeepaProductEnrichFetch_）が後で埋める
      thinPass++;
    }
    products.push(all[i]);
    if (fetched1[key]) writeList.push(all[i]);
  }

  var fullPlus = prAppendKeepaFull_(writeList);
  var waitMin = prKeepaWaitMinutes_(consumed, rate);
  if (rateLimited && Number(refillIn) > 0) {
    waitMin = Math.max(waitMin, Math.ceil(Number(refillIn) / 60000) + 1);
  }
  Logger.log(
    '[仕入れ検討①] two_stage hit=%s s1=%s pass=%s s2=%s got2=%s notReady=%s thinPass=%s consumed=%s left=%s full+=%s',
    packed.have.length,
    packed.need.length,
    Object.keys(passSet).length,
    need2.length,
    Object.keys(fat).length,
    Object.keys(notReady).length,
    thinPass,
    consumed,
    tokensLeft,
    fullPlus
  );
  return {
    products: products,
    fullPlus: fullPlus,
    need: packed.need.length,
    skip: packed.have.length,
    tokensConsumed: consumed,
    tokensLeft: tokensLeft,
    refillRate: rate,
    refillIn: refillIn,
    waitMin: waitMin,
    rateLimited: rateLimited,
    notReady: notReady,
    stage1: packed.need.length,
    stage2: Object.keys(fat).length,
  };
}

function prParseKeepaFullTime_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) return raw.getTime();
  var s = String(raw || '').trim();
  if (!s) return 0;
  var t = Date.parse(s);
  return isNaN(t) ? 0 : t;
}

function prLoadKeepaFullMap_() {
  if (PR_KEEPA_FULL_MAP_CACHE_) return PR_KEEPA_FULL_MAP_CACHE_;
  var out = {};
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_KEEPA_FULL_SHEET_);
    if (!sh || sh.getLastRow() < 2) return out;
    var lastCol = sh.getLastColumn();
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxAsin = -1;
    var idxAt = -1;
    var idxJson = -1;
    var idxFp = -1;
    var h;
    for (h = 0; h < hdr.length; h++) {
      var name = String(hdr[h] || '').trim();
      if (name === 'ASIN') idxAsin = h;
      if (name === '取得日時') idxAt = h;
      if (name === '生JSON') idxJson = h;
      if (name === '価格指紋') idxFp = h;
    }
    if (idxAsin < 0 || idxJson < 0) return out;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      var asin = String(vals[r][idxAsin] || '').trim().toUpperCase();
      if (!asin) continue;
      // 先頭挿入運用: 上側が最新。先勝ちで上書きしない。
      if (out[asin]) continue;
      var product = null;
      try {
        product = JSON.parse(String(vals[r][idxJson] || ''));
      } catch (e1) {
        product = null;
      }
      out[asin] = {
        at: idxAt >= 0 ? prParseKeepaFullTime_(vals[r][idxAt]) : 0,
        product: product,
        fp: idxFp >= 0 ? String(vals[r][idxFp] || '') : '',
      };
    }
  } catch (e) {
    Logger.log('[仕入れ検討①] Keepaフル読取スキップ %s', String((e && e.message) || e).slice(0, 160));
  }
  PR_KEEPA_FULL_MAP_CACHE_ = out;
  return out;
}

function prSplitKeepaFull_(asins) {
  var have = [];
  var need = [];
  var map = prLoadKeepaFullMap_();
  var now = new Date().getTime();
  var i;
  for (i = 0; i < asins.length; i++) {
    var a = String(asins[i] || '').trim().toUpperCase();
    if (!a) continue;
    var rec = map[a];
    if (rec && rec.product && rec.at && now - rec.at <= PR_FRESH_MS_) {
      have.push(rec.product);
    } else {
      need.push(a);
    }
  }
  return { have: have, need: need };
}

/**
 * @param {!Array<string>} asins
 * @param {{offers?:number}=} opt offers を省略・0 にすると offers を付けない（1トークン/件）
 */
function prKeepaProductStats90Fetch_(asins, opt) {
  var offersN = opt && opt.offers != null ? Number(opt.offers) : 20;
  if (!(offersN > 0)) offersN = 0;
  var mode = offersN > 0 ? 'offers=' + offersN : 'no_offers';
  var out = [];
  var tokensConsumed = 0;
  var tokensLeft = '';
  var refillRate = prKeepaRefillRate_();
  var refillIn = '';
  var rateLimited = false;
  var i;
  for (i = 0; i < asins.length; i += 50) {
    var chunk = asins.slice(i, i + 50);
    var url =
      PR_KEEPA_PRODUCT_ +
      '?key=' +
      encodeURIComponent(prKeepaKey_()) +
      '&domain=5&asin=' +
      encodeURIComponent(chunk.join(',')) +
      '&history=0&stats=90' +
      (offersN > 0 ? '&offers=' + offersN : '');
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { Accept: 'application/json' },
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText() || '';
    if (code !== 200) {
      if (code === 429) {
        rateLimited = true;
        try {
          var data429 = JSON.parse(body);
          if (data429.tokensLeft != null) tokensLeft = data429.tokensLeft;
          if (data429.refillRate != null && Number(data429.refillRate) > 0) {
            refillRate = Number(data429.refillRate);
          }
          if (data429.refillIn != null) refillIn = data429.refillIn;
        } catch (e429) {}
        Logger.log(
          '[仕入れ検討①] product stats=90 %s RATE_LIMIT left=%s refillIn=%s got=%s',
          mode,
          tokensLeft,
          refillIn,
          out.length
        );
        break;
      }
      throw new Error('Keepa /product HTTP ' + code + ' ' + prRedactKey_(body).slice(0, 180));
    }
    var data = JSON.parse(body);
    var cPart = Number(data.tokensConsumed);
    if (cPart > 0) tokensConsumed += cPart;
    if (data.tokensLeft != null) tokensLeft = data.tokensLeft;
    if (data.refillRate != null && Number(data.refillRate) > 0) refillRate = Number(data.refillRate);
    if (data.refillIn != null) refillIn = data.refillIn;
    var waitPart = prKeepaWaitMinutes_(cPart > 0 ? cPart : 0, refillRate);
    Logger.log(
      '[仕入れ検討①] product stats=90 %s n=%s consumed=%s left=%s refillRate=%s refillIn=%s waitMin=%s',
      mode,
      (data.products || []).length,
      data.tokensConsumed,
      data.tokensLeft,
      data.refillRate,
      data.refillIn,
      waitPart
    );
    var products = data.products || [];
    var j;
    for (j = 0; j < products.length; j++) out.push(products[j]);
  }
  return {
    products: out,
    tokensConsumed: tokensConsumed,
    tokensLeft: tokensLeft,
    refillRate: refillRate,
    refillIn: refillIn,
    rateLimited: rateLimited,
  };
}

/**
 * 門通過・D1向け補強 GET（history=0&stats=90&offers=20）。BB／出品者／BuyBox_FBA 用。
 * 生JSONには offers を残さない（prProductToFullRec_ で削除）。
 */
function prKeepaProductEnrichFetch_(asins) {
  var out = [];
  var tokensConsumed = 0;
  var tokensLeft = '';
  var refillRate = prKeepaRefillRate_();
  var refillIn = '';
  var rateLimited = false;
  var i;
  for (i = 0; i < asins.length; i += 50) {
    var chunk = asins.slice(i, i + 50);
    var url =
      PR_KEEPA_PRODUCT_ +
      '?key=' +
      encodeURIComponent(prKeepaKey_()) +
      '&domain=5&asin=' +
      encodeURIComponent(chunk.join(',')) +
      '&history=0&stats=90&offers=20';
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { Accept: 'application/json' },
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText() || '';
    if (code !== 200) {
      if (code === 429) {
        rateLimited = true;
        try {
          var data429 = JSON.parse(body);
          if (data429.tokensLeft != null) tokensLeft = data429.tokensLeft;
          if (data429.refillRate != null && Number(data429.refillRate) > 0) {
            refillRate = Number(data429.refillRate);
          }
          if (data429.refillIn != null) refillIn = data429.refillIn;
        } catch (e429) {}
        Logger.log('[仕入れ検討①] enrich offers=20 RATE_LIMIT left=%s', tokensLeft);
        break;
      }
      throw new Error('Keepa enrich /product HTTP ' + code + ' ' + prRedactKey_(body).slice(0, 180));
    }
    var data = JSON.parse(body);
    var cPart = Number(data.tokensConsumed);
    if (cPart > 0) tokensConsumed += cPart;
    if (data.tokensLeft != null) tokensLeft = data.tokensLeft;
    if (data.refillRate != null && Number(data.refillRate) > 0) refillRate = Number(data.refillRate);
    if (data.refillIn != null) refillIn = data.refillIn;
    var products = data.products || [];
    var j;
    for (j = 0; j < products.length; j++) out.push(products[j]);
  }
  return {
    products: out,
    tokensConsumed: tokensConsumed,
    tokensLeft: tokensLeft,
    refillRate: refillRate,
    refillIn: refillIn,
    rateLimited: rateLimited,
  };
}

function prFullNeedsD1Enrich_(full) {
  // 案A: Keepaフルに行があれば補強しない（価格・順位でリスト化。欠落は洗い／needFlat の offers=20 に任せる）。
  if (!full) return true;
  if (!prD1EnrichBlankBbEnabled_()) return false;
  // 2段ゲート（PR_KEEPA_TWO_STAGE_GATE）以後、門で落ちた ASIN は offers なしの薄い行として残る。
  // 再門で通過に変わった行・2段目が届かなかった行は BuyBox 2列が空になり、店レーンが拾えない。
  // 実測 2026-09-07: 通過782件中5件だけ該当（Keepaフル全体の空1,021件はほぼ落ち行）。
  var bb = String(full['BuyBoxセラー'] || '').trim();
  var fba = String(full['BuyBox_FBA'] || '').trim();
  return !bb && !fba;
}

/** Property PR_D1_ENRICH_BLANK_BB。未設定=ON。false で「行があれば補強しない」旧動作。 */
function prD1EnrichBlankBbEnabled_() {
  return prCkEnabled_(PR_D1_ENRICH_BLANK_BB_PROP_);
}

/** 1回の補強 GET 上限（既定20＝約100トークン）。BuyBox が元々無い商品での空回りを抑える。 */
function prD1EnrichCap_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_D1_ENRICH_CAP_PROP_) ||
      PR_D1_ENRICH_CAP_DEFAULT_
  );
  if (!(n > 0)) n = PR_D1_ENRICH_CAP_DEFAULT_;
  return n;
}

/**
 * 門通過 ASIN のうち Keepaフルに行が無いものだけ offers=20 で取得して upsert。
 * 行がある場合は再GETしない（案A）。
 * @return {{enriched:number, rateLimited:boolean, detail:string}}
 */
function prEnrichKeepaForPassAsins_(asins, flatMap) {
  var need = [];
  var i;
  for (i = 0; i < asins.length; i++) {
    var a = String(asins[i] || '')
      .trim()
      .toUpperCase();
    if (!a) continue;
    if (prFullNeedsD1Enrich_(flatMap[a])) need.push(a);
  }
  if (!need.length) {
    return { enriched: 0, rateLimited: false, detail: 'enrich_skip_planA' };
  }
  var needAll = need.length;
  var enrichCap = prD1EnrichCap_();
  if (need.length > enrichCap) need = need.slice(0, enrichCap);
  var pack = prKeepaProductEnrichFetch_(need);
  var n = prAppendKeepaFull_(pack.products || []);
  Logger.log(
    '[仕入れ検討①] D1 enrich need=%s cap=%s got=%s upsert=%s rateLimited=%s consumed=%s',
    needAll,
    need.length,
    (pack.products || []).length,
    n,
    pack.rateLimited ? 'Y' : 'N',
    pack.tokensConsumed
  );
  return {
    enriched: n,
    rateLimited: !!pack.rateLimited,
    detail:
      'enrich_need=' +
      needAll +
      (needAll > need.length ? ' cap=' + need.length : '') +
      ' upsert=' +
      n +
      (pack.rateLimited ? ' RATE_LIMIT' : ''),
  };
}

/** Script Property PR_KEEPA_FULL_WRITE_ENABLED。未設定=ON。false で切戻し。 */
function prKeepaFullWriteEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_KEEPA_FULL_WRITE_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

/**
 * GET した product だけ Keepaフルへ。同一ASINは上書き（upsert）。新規は2行目へ挿入。
 * 空ASIN行は削除。モールヒット非書。csv[]は倉庫に置かない。
 * @return {number} 新規挿入＋上書き件数
 */
function prAppendKeepaFull_(products) {
  if (!prKeepaFullWriteEnabled_()) {
    Logger.log('[仕入れ検討①] Keepaフル追記スキップ %s=false', PR_KEEPA_FULL_WRITE_PROP_);
    return 0;
  }
  if (!products || !products.length) return 0;
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_KEEPA_FULL_SHEET_);
  if (!sh) throw new Error('Keepaフルシートなし');
  prCleanupKeepaFullBlankRows_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var rowIndex = prKeepaFullAsinRowIndex_(sh, headers);
  var map = prLoadKeepaFullMap_();
  var fetched = prNowFetchedAt_();
  var insertLines = [];
  var wrote = 0;
  var i;
  for (i = 0; i < products.length; i++) {
    var rec = prProductToFullRec_(products[i], fetched);
    var asin = rec.ASIN;
    if (!asin) continue;
    if (rec._csvLeft) {
      Logger.log('[仕入れ検討①] Keepaフル csv残 skip %s', asin);
      continue;
    }
    var line = [];
    var h;
    for (h = 0; h < headers.length; h++) {
      var key = String(headers[h] || '').trim();
      line.push(key && rec[key] != null ? rec[key] : '');
    }
    var prev = map[asin];
    var loc = rowIndex[asin];
    if (loc && loc.rowNum > 0) {
      var sameFp = prev && String(prev.fp || '') === rec['価格指紋'];
      var newHasOffers = !!(products[i].offers && products[i].offers.length);
      var prevHasOffers = !!(prev && prev.product && prev.product.offers && prev.product.offers.length);
      // 指紋同でも薄い行→offers=20 への補強は上書きする
      if (sameFp && (!newHasOffers || prevHasOffers)) {
        Logger.log('[仕入れ検討①] Keepaフル skip_same_fp %s row=%s', asin, loc.rowNum);
        continue;
      }
      sh.getRange(loc.rowNum, 1, 1, line.length).setValues([line]);
      wrote++;
      if (loc.extraRows && loc.extraRows.length) {
        prDeleteSheetRowsDescending_(sh, loc.extraRows);
      }
    } else {
      insertLines.push(line);
      wrote++;
    }
    map[asin] = { at: new Date().getTime(), product: products[i], fp: rec['価格指紋'] };
  }
  if (insertLines.length) prInsertRowsBelowHeader_(sh, insertLines);
  PR_KEEPA_FULL_MAP_CACHE_ = map;
  // 先頭挿入で行番号がずれるため、ASIN列だけ見て索引を自動再構築（人手メンテなし）
  try {
    prRebuildKeepaFullIndex_(sh);
  } catch (idxErr) {
    Logger.log(
      '[仕入れ検討①] Keepaフル索引再構築失敗 %s',
      String((idxErr && idxErr.message) || idxErr).slice(0, 160)
    );
  }
  return wrote;
}

/** ASIN → { rowNum, extraRows[] }。上側が正。 */
function prKeepaFullAsinRowIndex_(sh, headers) {
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var idxAsin = -1;
  var h;
  for (h = 0; h < headers.length; h++) {
    if (String(headers[h] || '').trim() === 'ASIN') idxAsin = h;
  }
  if (idxAsin < 0) return out;
  var n = sh.getLastRow() - 1;
  // ASIN→行 しか作らないので ASIN 列だけ読む（旧: 生JSON を含む全列を読んでいた）
  var vals = sh.getRange(2, idxAsin + 1, n, 1).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    var asin = String(vals[r][0] || '')
      .trim()
      .toUpperCase();
    if (!asin) continue;
    var sheetRow = r + 2;
    if (!out[asin]) out[asin] = { rowNum: sheetRow, extraRows: [] };
    else out[asin].extraRows.push(sheetRow);
  }
  return out;
}

/** 行番号を降順に削除（下から）。 */
function prDeleteSheetRowsDescending_(sh, rowNums) {
  if (!sh || !rowNums || !rowNums.length) return;
  var uniq = {};
  var i;
  for (i = 0; i < rowNums.length; i++) {
    var n = Number(rowNums[i]);
    if (n >= 2) uniq[n] = true;
  }
  var list = Object.keys(uniq)
    .map(function (k) {
      return Number(k);
    })
    .sort(function (a, b) {
      return b - a;
    });
  for (i = 0; i < list.length; i++) sh.deleteRow(list[i]);
}

/** Property PR_KEEPA_FULL_BLANK_SCAN_BATCH。未設定=ON。false で1行ずつ読む旧動作。 */
function prKeepaFullBlankScanBatchEnabled_() {
  return prCkEnabled_(PR_KEEPA_FULL_BLANK_SCAN_BATCH_PROP_);
}

/**
 * ASIN空の完全空行を下から削除。
 * 2026-09-07: ASIN列を1回読んで候補行だけ全列を見る。
 * 旧: 最終行から1行ずつ getRange().getValues()。Keepaフルは実データ約8,400行あるため、
 * 空行が0でも毎回8,400回の個別読み取り＝約22分。Keepa取得のあるスライスが常に
 * 22〜29分（1ASINでも22.9分）になっていた原因。切戻し PR_KEEPA_FULL_BLANK_SCAN_BATCH=false。
 */
function prCleanupKeepaFullBlankRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxAsin = -1;
  var h;
  for (h = 0; h < headers.length; h++) {
    if (String(headers[h] || '').trim() === 'ASIN') idxAsin = h;
  }
  if (idxAsin < 0) return 0;
  var last = sh.getLastRow();
  var deleted = 0;
  var r;
  if (prKeepaFullBlankScanBatchEnabled_()) {
    var t0 = new Date().getTime();
    // ASIN が入っている行は消さない。まず ASIN 列だけ読んで候補を絞る
    var asinCol = sh.getRange(2, idxAsin + 1, last - 1, 1).getValues();
    var cand = [];
    for (r = 0; r < asinCol.length; r++) {
      if (String(asinCol[r][0] || '').trim() === '') cand.push(r + 2);
    }
    var blank = [];
    // 候補だけ全列を見る（生JSON を含む全行読みはしない）。連続行はまとめ読み
    var i = 0;
    var reads = 0;
    while (i < cand.length) {
      var j = i;
      while (j + 1 < cand.length && cand[j + 1] === cand[j] + 1) j++;
      var n = j - i + 1;
      var vals = sh.getRange(cand[i], 1, n, lastCol).getValues();
      reads++;
      var k;
      for (k = 0; k < n; k++) {
        var any = false;
        var c;
        for (c = 0; c < vals[k].length; c++) {
          if (String(vals[k][c] || '').trim() !== '') {
            any = true;
            break;
          }
        }
        if (!any) blank.push(cand[i] + k);
      }
      i = j + 1;
    }
    prDeleteSheetRowsDescending_(sh, blank);
    deleted = blank.length;
    Logger.log(
      '[仕入れ検討①] Keepaフル blank_scan rows=%s cand=%s reads=%s deleted=%s elapsedMs=%s',
      last - 1,
      cand.length,
      reads,
      deleted,
      new Date().getTime() - t0
    );
  } else {
    for (r = last; r >= 2; r--) {
      var row = sh.getRange(r, 1, 1, lastCol).getValues()[0];
      var asin = String(row[idxAsin] || '').trim();
      if (asin) continue;
      var any2 = false;
      var c2;
      for (c2 = 0; c2 < row.length; c2++) {
        if (String(row[c2] || '').trim() !== '') {
          any2 = true;
          break;
        }
      }
      if (!any2) {
        sh.deleteRow(r);
        deleted++;
      }
    }
  }
  if (deleted) {
    PR_KEEPA_FULL_MAP_CACHE_ = null;
    PR_KEEPA_FULL_INDEX_CACHE_ = null;
    Logger.log('[仕入れ検討①] Keepaフル blank_rows_deleted=%s', deleted);
    try {
      prRebuildKeepaFullIndex_(sh);
    } catch (idxErr2) {
      Logger.log(
        '[仕入れ検討①] Keepaフル索引再構築失敗 %s',
        String((idxErr2 && idxErr2.message) || idxErr2).slice(0, 160)
      );
    }
  }
  return deleted;
}

/**
 * データ行をヘッダ直下（2行目）へ挿入。複数行はまとめて差し込み（上が先の行）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {Array<Array<*>>} lines
 */
function prInsertRowsBelowHeader_(sh, lines) {
  if (!sh || !lines || !lines.length) return;
  var n = lines.length;
  var cols = lines[0].length;
  sh.insertRowsBefore(2, n);
  sh.getRange(2, 1, n, cols).setValues(lines);
}

function prStripKeepaCsv_(v) {
  if (v === null || v === undefined) return v;
  if (Object.prototype.toString.call(v) === '[object Array]') {
    var a = [];
    var i;
    for (i = 0; i < v.length; i++) a.push(prStripKeepaCsv_(v[i]));
    return a;
  }
  if (typeof v === 'object') {
    var o = {};
    var k;
    for (k in v) {
      if (!Object.prototype.hasOwnProperty.call(v, k)) continue;
      if (k === 'csv') continue;
      o[k] = prStripKeepaCsv_(v[k]);
    }
    return o;
  }
  return v;
}

function prKeepaSlotStr_(arr, idx) {
  if (!arr || !arr.length || idx >= arr.length) return '';
  var v = arr[idx];
  if (v === null || v === undefined || v === -1) return '';
  return String(v);
}

function prKeepaStatsAt_(stats, key, idx) {
  var st = stats || {};
  return prKeepaSlotStr_(st[key] || [], idx);
}

function prPriceFingerprint_(p) {
  var st = p && p.stats ? p.stats : {};
  var cur = st.current || [];
  return (
    prKeepaSlotStr_(cur, 18) +
    '|' +
    prKeepaSlotStr_(cur, 1) +
    '|' +
    prKeepaSlotStr_(cur, 11) +
    '|' +
    String((p && p.title) || '')
  );
}

function prAmazonDirectLabel_(v) {
  if (v === 0 || v === '0') return 'いる';
  if (v === -1 || v === '-1') return 'いない';
  return '';
}

function prKeepaOos180Str_(stats) {
  var arr = stats && stats.outOfStockPercentage180;
  if (!arr || !arr.length) return '';
  var v = arr[0];
  if (v === null || v === undefined || v === -1) return '';
  return String(v);
}

function prKeepaRatingStr_(v) {
  if (v === null || v === undefined || v === '' || v === -1) return '';
  var n = Number(v);
  if (isNaN(n)) return String(v);
  if (n > 5 && n <= 50) n = n / 10.0;
  if (n <= 0) return '';
  return String(n);
}

function prPackCm_(v) {
  if (v === null || v === undefined || v === '' || v === -1) return '';
  var n = Number(v);
  if (!(n > 0)) return '';
  return String(n / 10);
}

function prBoolJa_(v) {
  if (v === true || v === 'true' || v === 1 || v === '1') return 'はい';
  if (v === false || v === 'false' || v === 0 || v === '0') return 'いいえ';
  return '';
}

function prFbaFeeYen_(p) {
  var ff = p && p.fbaFees;
  if (ff && typeof ff === 'object') {
    var v = ff.pickAndPackFee;
    if (v == null) v = ff.pickAndPack;
    if (v == null || v === '' || v === -1) return '';
    return String(v);
  }
  if (ff == null || ff === '' || ff === -1) return '';
  return String(ff);
}

function prCatRootTree_(p) {
  var tree = (p && p.categoryTree) || [];
  var names = [];
  var i;
  for (i = 0; i < tree.length; i++) {
    if (tree[i] && tree[i].name) names.push(String(tree[i].name));
  }
  return { root: names[0] || '', tree: names.join(' > ') };
}

function prKeepaImageUrls_(p) {
  var out = [];
  var seen = {};
  function add(name) {
    var s = String(name || '').trim();
    if (!s) return;
    var url = s.indexOf('http') === 0 ? s : 'https://m.media-amazon.com/images/I/' + s;
    if (seen[url]) return;
    seen[url] = true;
    out.push(url);
  }
  var arr = (p && p.images) || [];
  var i;
  for (i = 0; i < arr.length; i++) {
    var it = arr[i];
    if (it && typeof it === 'object') add(it.l || it.m);
    else add(it);
  }
  var csv = String((p && (p.imagesCSV || p.image)) || '');
  if (csv) {
    var parts = csv.split(',');
    for (i = 0; i < parts.length; i++) add(parts[i]);
  }
  return out;
}

function prProductToFullRec_(p, fetchedAt) {
  var stripped = prStripKeepaCsv_(p || {});
  if (stripped.offers) delete stripped.offers;
  var asin = String(stripped.asin || '').trim().toUpperCase();
  var stats = stripped.stats && typeof stripped.stats === 'object' ? stripped.stats : {};
  var cur = stats.current || [];
  var eans = stripped.eanList || [];
  var ean = eans.length ? eans[0] : stripped.ean || '';
  var urls = prKeepaImageUrls_(stripped);
  var cat = prCatRootTree_(stripped);
  var bbNow =
    prKeepaSlotStr_(cur, 18) ||
    (stats.buyBoxPrice != null &&
    stats.buyBoxPrice !== '' &&
    stats.buyBoxPrice !== -1 &&
    stats.buyBoxPrice !== -2
      ? String(stats.buyBoxPrice)
      : '');
  var rec = {
    取得日時: fetchedAt,
    目的: PR_PURPOSE_RESEARCH_,
    ASIN: asin,
    '商品コード: EAN': String(ean || ''),
    商品名: String(stripped.title || ''),
    画像: urls[0] ? '=HYPERLINK("' + urls[0].replace(/"/g, '') + '","メイン")' : '',
    サブ画像: urls.slice(1).join('|'),
    製造者: String(stripped.manufacturer || ''),
    ブランド: String(stripped.brand || ''),
    親ASIN: String(stripped.parentAsin || ''),
    'URL: Amazon': asin ? 'https://www.amazon.co.jp/dp/' + asin : '',
    'URL: Keepa': asin ? 'https://keepa.com/#!product/5-' + asin : '',
    'Buy Box: 現在価格': bbNow,
    'Buy Box: 30 日平均': prKeepaStatsAt_(stats, 'avg30', 18),
    'Buy Box: 90 日平均': prKeepaStatsAt_(stats, 'avg90', 18),
    'Amazon: 現在価格': prKeepaSlotStr_(cur, 0),
    'Amazon: 30 日平均': prKeepaStatsAt_(stats, 'avg30', 0),
    'Amazon: 90 日平均': prKeepaStatsAt_(stats, 'avg90', 0),
    '新品: 90 日平均': prKeepaStatsAt_(stats, 'avg90', 1),
    '参考価格: 90 日平均': prKeepaStatsAt_(stats, 'avg90', 4),
    '売れ筋ランキング: 90 日平均': prKeepaStatsAt_(stats, 'avg90', 3),
    月間売上: stripped.monthlySold == null || stripped.monthlySold === '' ? '' : String(stripped.monthlySold),
    出品者数: prKeepaSlotStr_(cur, 11),
    Amazon直販: prAmazonDirectLabel_(stripped.availabilityAmazon),
    '新品: 現在価格': prKeepaSlotStr_(cur, 1),
    '売れ筋ランキング: 現在': prKeepaSlotStr_(cur, 3),
    'Amazon: 180日在庫切れ%': prKeepaOos180Str_(stats),
    'レビュー: 評価': prKeepaRatingStr_(prKeepaSlotStr_(cur, 16) || stripped.rating),
    'レビュー: 評価件数': prKeepaSlotStr_(cur, 17) || (stripped.reviewCount == null || stripped.reviewCount === '' || stripped.reviewCount === -1 ? '' : String(stripped.reviewCount)),
    発売日: String(stripped.releaseDate || stripped.listedSince || ''),
    'カテゴリ: ルート': cat.root,
    'カテゴリ: ツリー': cat.tree,
    アイテム数: stripped.numberOfItems ? String(stripped.numberOfItems) : '',
    パッケージ数量: stripped.packageQuantity ? String(stripped.packageQuantity) : '',
    梱包_L_cm: prPackCm_(stripped.packageLength),
    梱包_W_cm: prPackCm_(stripped.packageWidth),
    梱包_H_cm: prPackCm_(stripped.packageHeight),
    梱包_重量_g: stripped.packageWeight == null || stripped.packageWeight === '' || stripped.packageWeight === -1 ? '' : String(stripped.packageWeight),
    FBA手数料: prFbaFeeYen_(stripped),
    BuyBoxセラー: String(stats.buyBoxSellerId || '').trim(),
    BuyBox_FBA: prBoolJa_(stats.buyBoxIsFBA),
    価格指紋: prPriceFingerprint_(stripped),
    生JSON: JSON.stringify(stripped),
    _csvLeft: stripped.csv != null,
  };
  return rec;
}

function prReadAllProfiles_() {
  if (PR_PROF_CACHE_) return PR_PROF_CACHE_;
  var sh = prSpreadsheet_().getSheetByName(PR_PROF_SHEET_);
  if (!sh || sh.getLastRow() < 2) {
    PR_PROF_CACHE_ = [];
    return PR_PROF_CACHE_;
  }
  prEnsureProfileSellerCountCols_(sh);
  prEnsureProfileSurveyCols_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var vals = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  var headers = vals[0];
  var out = [];
  var r;
  for (r = 1; r < vals.length; r++) {
    var map = {};
    var i;
    for (i = 0; i < headers.length; i++) map[String(headers[i] || '')] = vals[r][i];
    var en = String(map['有効'] || 'はい');
    if (en === 'いいえ' || en === '0') continue;
    out.push(map);
  }
  PR_PROF_CACHE_ = out;
  return out;
}

/** ①プロファイルに門用の上下限列を確保。未設定行は既定埋め。列順は左寄せ。 */
function prEnsureProfileSellerCountCols_(sh) {
  if (!sh || sh.getLastRow() < 1) return;
  var specs = [
    { header: PR_PROF_SELLER_MIN_HEADER_, def: PR_PROF_SELLER_MIN_DEFAULT_ },
    { header: PR_PROF_SELLER_MAX_HEADER_, def: PR_PROF_SELLER_MAX_DEFAULT_ },
    { header: PR_PROF_RANK_MIN_HEADER_, def: PR_PROF_RANK_MIN_DEFAULT_ },
    { header: PR_PROF_PRICE_MAX_HEADER_, def: PR_PROF_PRICE_MAX_DEFAULT_ },
    { header: PR_PROF_WEIGHT_MAX_HEADER_, def: PR_PROF_WEIGHT_MAX_DEFAULT_ },
  ];
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxs = [];
  var s;
  var h;
  for (s = 0; s < specs.length; s++) {
    var idx = -1;
    for (h = 0; h < headers.length; h++) {
      if (String(headers[h] || '').trim() === specs[s].header) {
        idx = h;
        break;
      }
    }
    if (idx < 0) {
      idx = lastCol;
      sh.getRange(1, idx + 1).setValue(specs[s].header);
      headers.push(specs[s].header);
      lastCol++;
    }
    idxs.push(idx);
  }
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var r;
    for (s = 0; s < specs.length; s++) {
      var col = idxs[s] + 1;
      var vals = sh.getRange(2, col, lastRow - 1, 1).getValues();
      var changed = false;
      for (r = 0; r < vals.length; r++) {
        if (vals[r][0] === '' || vals[r][0] == null) {
          vals[r][0] = specs[s].def;
          changed = true;
        }
      }
      if (changed) sh.getRange(2, col, vals.length, 1).setValues(vals);
    }
  }
  prReorderProfileCols_(sh);
}

/**
 * 調査カテゴリー初期行の判定。競合DBルート名「食品・飲料・お酒」が正。
 * カテゴリ名がルートIDだけの行は対象外（誤行）。
 */
function prIsFoodSurveyCategoryName_(cat) {
  var s = String(cat || '').trim();
  if (!s) return false;
  if (/^\d+$/.test(s)) return false;
  if (s === '食品・飲料・お酒') return true;
  if (s === '食品') return true;
  if (s.indexOf('食品・飲料・お酒') === 0) return true;
  return false;
}

/**
 * ①プロファイルに調査カテゴリー・AmazonルートIDを確保。食品行は初期 TRUE + 57239051。
 * CYCLE §4.1bis。
 */
function prEnsureProfileSurveyCols_(sh) {
  if (!sh || sh.getLastRow() < 1) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  function ensureHeader_(name) {
    var h;
    for (h = 0; h < headers.length; h++) {
      if (String(headers[h] || '').trim() === name) return h;
    }
    var idx = lastCol;
    sh.getRange(1, idx + 1).setValue(name);
    headers.push(name);
    lastCol++;
    return idx;
  }
  var idxCk = ensureHeader_(PR_PROF_SURVEY_CK_HEADER_);
  var idxRoot = ensureHeader_(PR_PROF_ROOT_ID_HEADER_);
  var idxCat = -1;
  var h;
  for (h = 0; h < headers.length; h++) {
    if (String(headers[h] || '').trim() === 'カテゴリ') {
      idxCat = h;
      break;
    }
  }
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var n = lastRow - 1;
    var ckVals = sh.getRange(2, idxCk + 1, n, 1).getValues();
    var rootVals = sh.getRange(2, idxRoot + 1, n, 1).getValues();
    var catVals =
      idxCat >= 0 ? sh.getRange(2, idxCat + 1, n, 1).getValues() : null;
    var ckChanged = false;
    var rootChanged = false;
    var r;
    for (r = 0; r < n; r++) {
      var cat = catVals ? String(catVals[r][0] || '').trim() : '';
      var isFood = prIsFoodSurveyCategoryName_(cat);
      if (rootVals[r][0] === '' || rootVals[r][0] == null) {
        if (isFood) {
          rootVals[r][0] = PR_FOOD_ROOT_;
          rootChanged = true;
        }
      } else {
        rootVals[r][0] = String(rootVals[r][0])
          .trim()
          .replace(/\.0$/, '');
      }
      if (ckVals[r][0] === '' || ckVals[r][0] == null) {
        ckVals[r][0] = isFood;
        ckChanged = true;
      }
    }
    if (ckChanged) sh.getRange(2, idxCk + 1, n, 1).setValues(ckVals);
    if (rootChanged) sh.getRange(2, idxRoot + 1, n, 1).setValues(rootVals);
    try {
      sh.getRange(2, idxCk + 1, n, 1).insertCheckboxes();
    } catch (eCk) {}
  }
  prReorderProfileCols_(sh);
  PR_PROF_CACHE_ = null;
}

/** 価格上下限・順位まわりを左で隣接。moveColumns(Range, dest) で書式・プルダウンを列ごと移動。 */
function prReorderProfileCols_(sh) {
  if (!sh) return;
  var order = PR_PROF_COL_ORDER_;
  var k;
  for (k = 0; k < order.length; k++) {
    var want = order[k];
    var dest = k + 1;
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var from = -1;
    var h;
    for (h = 0; h < headers.length; h++) {
      if (String(headers[h] || '').trim() === want) {
        from = h + 1;
        break;
      }
    }
    if (from < 0 || from === dest) continue;
    // 正: moveColumns(Range, destinationIndex)。(number,number) はシグネチャ不一致で落ちる。
    sh.moveColumns(sh.getRange(1, from), dest);
  }
}

function prStageKey_(raw) {
  var s = String(raw == null || raw === '' ? '1' : raw).trim();
  if (/^\d+\.0$/.test(s)) s = s.slice(0, -2);
  if (s === '2' || s === '順位段階2') return '順位段階2';
  if (s === '3' || s === '順位段階3') return '順位段階3';
  return '順位段階1';
}

function prRankLimit_(prof) {
  var key = prStageKey_(prof['門に使う順位段階']);
  var n = Number(prof[key]);
  if (!(n > 0)) n = Number(prof['順位段階1']);
  return n > 0 ? n : 150000;
}

function prCatBlob_(p) {
  var tree = (p && p.categoryTree) || [];
  var names = [];
  var i;
  for (i = 0; i < tree.length; i++) {
    if (tree[i] && tree[i].name) names.push(String(tree[i].name));
  }
  return names.join(' ');
}

function prProfileForProduct_(p, profs) {
  var blob = prCatBlob_(p);
  var root = blob.split(' ')[0] || '';
  var common = null;
  var i;
  for (i = 0; i < profs.length; i++) {
    var m = profs[i];
    var cat = String(m['カテゴリ'] || '');
    if (cat === '共通') {
      common = m;
      continue;
    }
    var needles = String(m['マッチ語'] || '').split('|');
    var j;
    for (j = 0; j < needles.length; j++) {
      if (needles[j] && blob.indexOf(needles[j]) >= 0) return m;
    }
    if (cat && blob.indexOf(cat) >= 0) return m;
    if (cat === '食品' && blob.indexOf('食品') === 0) return m;
  }
  return (
    common ||
    profs[0] || {
      価格下限: 2000,
      価格上限: PR_PROF_PRICE_MAX_DEFAULT_,
      順位段階1: 150000,
      順位下限: PR_PROF_RANK_MIN_DEFAULT_,
      冷凍生: '落とす',
      本体ルール: '見ない',
    }
  );
}

/** D2 用: 品番リストのカテゴリー文字列からプロファイル行を選ぶ。 */
function prProfileForCategoryName_(categoryStr, profs) {
  var parts = String(categoryStr || '')
    .split(/[›>]/)
    .map(function (s) {
      return String(s || '').trim();
    })
    .filter(function (s) {
      return !!s;
    });
  var tree = [];
  var i;
  for (i = 0; i < parts.length; i++) tree.push({ name: parts[i] });
  if (!tree.length && String(categoryStr || '').trim()) {
    tree.push({ name: String(categoryStr).trim() });
  }
  return prProfileForProduct_({ categoryTree: tree }, profs);
}

function prProfileWeightMaxG_(rowProf) {
  var n = Number(rowProf[PR_PROF_WEIGHT_MAX_HEADER_]);
  if (!(n > 0)) n = Number(rowProf['重量上限g']);
  return n > 0 ? n : PR_PROF_WEIGHT_MAX_DEFAULT_;
}

/** 大型列: 落とす（既定）なら D2 で大型・特大ティアを見送り。 */
function prProfileOversizeDrop_(rowProf) {
  var v = String(rowProf['大型'] || '落とす').trim();
  return !(v === '落とさない' || v === 'いいえ' || v === '0' || v === 'OFF');
}

function prParseWeightG_(raw) {
  var s = String(raw || '')
    .trim()
    .replace(/,/g, '');
  if (!s) return NaN;
  var m = s.match(/^([0-9]+(?:\.[0-9]+)?)\s*(kg|g)?$/i);
  if (m) {
    var v = Number(m[1]);
    if (!(v > 0)) return NaN;
    if (String(m[2] || '').toLowerCase() === 'kg') return v * 1000;
    return v;
  }
  var s2 = s.replace(/g$/i, '').trim();
  var n = Number(s2);
  return isFinite(n) && n > 0 ? n : NaN;
}

function prD2IsOversizeTier_(tier) {
  return /^大型|^特大/.test(String(tier || '').trim());
}

function prProfilePriceMax_(rowProf) {
  var n = Number(rowProf[PR_PROF_PRICE_MAX_HEADER_]);
  if (!(n > 0)) n = Number(rowProf['価格上限']);
  return n > 0 ? n : PR_PROF_PRICE_MAX_DEFAULT_;
}

function prProfileRankMin_(rowProf) {
  var n = Number(rowProf[PR_PROF_RANK_MIN_HEADER_]);
  if (!(n >= 0)) n = Number(rowProf['順位下限']);
  return n >= 0 ? n : PR_PROF_RANK_MIN_DEFAULT_;
}

function prReadProfile_() {
  var profs = prReadAllProfiles_();
  var pick = prProfileForProduct_({}, profs);
  return {
    priceMin: Number(pick['価格下限']) > 0 ? Number(pick['価格下限']) : 2000,
    priceMax: prProfilePriceMax_(pick),
    rankMin: prProfileRankMin_(pick),
    rankMax: prRankLimit_(pick),
    raw: pick,
  };
}

function prAmazonOos180_(p) {
  var st = p && p.stats ? p.stats : {};
  var arr = st.outOfStockPercentage180;
  if (!arr || !arr.length) return null;
  var v = arr[0];
  if (v === null || v === undefined || v === -1) return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

function prSellerCountNew_(p) {
  var st = p && p.stats ? p.stats : {};
  var cur = st.current || [];
  if (!cur.length || 11 >= cur.length) return null;
  var v = cur[11];
  if (v === null || v === undefined || v === '' || v === -1) return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

function prGateProductRow_(p, prof, sellerId, runId, shopName) {
  var table = prReadAllProfiles_();
  var rowProf = prProfileForProduct_(p, table);
  var priceMin = Number(rowProf['価格下限']) > 0 ? Number(rowProf['価格下限']) : 2000;
  var priceMax = prProfilePriceMax_(rowProf);
  var rankMin = prProfileRankMin_(rowProf);
  var rankMax = prRankLimit_(rowProf);
  var sellerMin = Number(rowProf[PR_PROF_SELLER_MIN_HEADER_]);
  if (!(sellerMin >= 0)) sellerMin = PR_PROF_SELLER_MIN_DEFAULT_;
  var sellerMax = Number(rowProf[PR_PROF_SELLER_MAX_HEADER_]);
  if (!(sellerMax > 0)) sellerMax = PR_PROF_SELLER_MAX_DEFAULT_;
  var title = String(p.title || '');
  var price = prStatsSlot_(p, 18);
  if (price === null) price = prStatsSlot_(p, 1);
  var rank = prStatsSlot_(p, 3);
  var sellerN = prSellerCountNew_(p);
  if (sellerN === null) sellerN = 0;
  var amazonDirect = prAmazonDirectLabel_(p.availabilityAmazon);
  var st = '通過';
  var why = '門通過';
  if (price === null && rank === null) {
    st = '落ち';
    why = 'stats空';
  } else if (price !== null && price < priceMin) {
    st = '落ち';
    why = '価格<' + priceMin;
  } else if (price !== null && price > priceMax) {
    st = '落ち';
    why = '価格>' + priceMax;
  } else if (rank === null) {
    st = '落ち';
    why = '順位欠測';
  } else if (rank < rankMin) {
    st = '落ち';
    why = '順位<' + rankMin;
  } else if (rank > rankMax) {
    st = '落ち';
    why = '順位>' + rankMax;
  } else if (sellerN < sellerMin) {
    st = '落ち';
    why = '出品者数<' + sellerMin + '(n=' + sellerN + ')';
  } else if (sellerN > sellerMax) {
    st = '落ち';
    why = '出品者数>' + sellerMax + '(n=' + sellerN + ')';
  } else if (String(rowProf['冷凍生'] || '落とす') !== '落とさない' && PR_FROZEN_RE_.test(title)) {
    st = '落ち';
    why = '冷凍冷蔵生鮮';
  } else if (String(rowProf['本体ルール'] || '') === '6ヶ月不在なら通過' && (p.availabilityAmazon === 0 || p.availabilityAmazon === '0')) {
    var oos = prAmazonOos180_(p);
    var thresh = Number(rowProf['本体不在%']);
    if (!(thresh > 0)) thresh = 80;
    if (oos !== null && oos < thresh) {
      st = '落ち';
      why = '本体いる(6ヶ月在庫あり/直販=' + amazonDirect + ')';
    }
  }
  if (st === '通過' && amazonDirect) {
    why = why + '/直販=' + amazonDirect;
  }
  var eans = p.eanList || [];
  var jan = eans.length ? String(eans[0]) : '';
  var mfr = String(p.manufacturer || p.brand || '');
  return [
    mfr,
    String(p.asin || '').toUpperCase(),
    title,
    price === null ? '' : price,
    rank === null ? '' : rank,
    jan,
    'keepa_query',
    sellerId,
    st,
    why,
    runId,
    new Date(),
    shopName || prSellerNameForId_(sellerId),
  ];
}

function prStatsSlot_(p, idx) {
  var st = p && p.stats ? p.stats : {};
  var keys = ['avg90', 'avg30', 'avg'];
  var k;
  for (k = 0; k < keys.length; k++) {
    var arr = st[keys[k]];
    if (!arr || !arr.length || idx >= arr.length) continue;
    var v = arr[idx];
    if (v === null || v === undefined || v === -1) continue;
    var n = Number(v);
    if (!isNaN(n) && n >= 0) return n;
  }
  return null;
}

function prSplitTok_(s) {
  var out = [];
  var seen = {};
  String(s || '')
    .split(/[|,\n、／/]+/)
    .forEach(function (p) {
      var t = String(p || '').trim();
      if (!t) return;
      var k = t.toUpperCase();
      if (seen[k]) return;
      seen[k] = true;
      out.push(t);
    });
  return out;
}

function prJoinTok_(parts) {
  var out = [];
  var seen = {};
  var i;
  for (i = 0; i < (parts || []).length; i++) {
    var t = String(parts[i] || '').trim();
    if (!t) continue;
    var k = t.toUpperCase();
    if (seen[k]) continue;
    seen[k] = true;
    out.push(t);
  }
  return out.join('|');
}

function prMergeTok_(a, b) {
  return prJoinTok_(prSplitTok_(a).concat(prSplitTok_(b)));
}

var PR_SELLER_NAME_CACHE_ = null;
var PR_SELLER_ROW_CACHE_ = null;
var PR_SELLER_ROWS_CACHE_ = null;
var PR_SELLER_STORE_CACHE_ = null;
var PR_SELLER_NAME_GET_N_ = 0;
var PR_SELLER_NAME_MISS_ = {};
var PR_SELLER_STORE_MISS_ = {};
var PR_SELLER_SYNC_SKIP_ = {
  調査CK: true,
  sellerId: true,
  採取元: true,
  メモ: true,
  巡回日: true,
  asinList件数: true,
  抽出メーカー: true,
  ピック: true,
};

function prSellerNameGetEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_SELLER_NAME_GET_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prSellerNameGetCap_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_SELLER_NAME_GET_CAP_PROP_) || 40);
  if (!(n > 0)) n = 40;
  if (n > 80) n = 80;
  return Math.floor(n);
}

function prLoadSellerLedger_() {
  if (PR_SELLER_NAME_CACHE_ && PR_SELLER_ROW_CACHE_ && PR_SELLER_STORE_CACHE_ && PR_SELLER_ROWS_CACHE_) {
    return;
  }
  var names = {};
  var rows = {};
  var allRows = {};
  var stores = {};
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
    if (sh && sh.getLastRow() >= 2) {
      var lastCol = sh.getLastColumn();
      var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      var idxId = prHeaderIndex_(hdr, 'sellerId');
      var idxName = prHeaderIndex_(hdr, '店名');
      var idxStore = prHeaderIndex_(hdr, 'ストアASIN数');
      if (idxId >= 0) {
        var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
        var i;
        for (i = 0; i < vals.length; i++) {
          var sid = String(vals[i][idxId] || '').trim();
          if (!sid) continue;
          var key = sid.toUpperCase();
          var sheetRow = i + 2;
          if (!allRows[key]) allRows[key] = [];
          allRows[key].push(sheetRow);
          if (!rows[key]) rows[key] = sheetRow;
          var name = idxName >= 0 ? String(vals[i][idxName] || '').trim() : '';
          if (name) {
            names[key] = name;
            rows[key] = sheetRow;
          }
          var storeN = idxStore >= 0 ? String(vals[i][idxStore] || '').trim() : '';
          if (storeN) stores[key] = storeN;
        }
      }
    }
  } catch (e) {
    Logger.log('[仕入れ検討①] セラー台帳読取スキップ %s', String((e && e.message) || e).slice(0, 160));
  }
  PR_SELLER_NAME_CACHE_ = names;
  PR_SELLER_ROW_CACHE_ = rows;
  PR_SELLER_ROWS_CACHE_ = allRows;
  PR_SELLER_STORE_CACHE_ = stores;
}

function prSellerNameMap_() {
  prLoadSellerLedger_();
  return PR_SELLER_NAME_CACHE_ || {};
}

/**
 * Keepa GET /seller storefront=0。asinList は取らない。
 * 店名・ストアASIN数・構成比を同じ応答から取る。
 * @return {{name:string, storeN:string, mix:string, foodPct:string, hypo:string, hasAsinList:boolean}|null}
 */
function prKeepaSellerGet_(sellerId) {
  var sid = String(sellerId || '').trim();
  if (!sid) return null;
  var url =
    PR_KEEPA_SELLER_ +
    '?key=' +
    encodeURIComponent(prKeepaKey_()) +
    '&domain=5&seller=' +
    encodeURIComponent(sid) +
    '&storefront=0';
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { Accept: 'application/json' },
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '';
  if (code !== 200) {
    Logger.log(
      '[仕入れ検討①] /seller HTTP %s seller=%s %s',
      code,
      sid,
      prRedactKey_(body).slice(0, 120)
    );
    return null;
  }
  var data = JSON.parse(body);
  Logger.log(
    '[仕入れ検討①] /seller storefront=0 seller=%s consumed=%s left=%s',
    sid,
    data.tokensConsumed,
    data.tokensLeft
  );
  return prParseKeepaSeller_(data, sid);
}

function prParseKeepaSeller_(data, sellerId) {
  var sid = String(sellerId || '').trim();
  var sellers = data && data.sellers;
  var rec = {};
  if (sellers && typeof sellers === 'object' && !Array.isArray(sellers)) {
    rec = sellers[sid] || sellers[sid.toUpperCase()] || {};
    if (!rec || !Object.keys(rec).length) {
      var k;
      for (k in sellers) {
        if (Object.prototype.hasOwnProperty.call(sellers, k)) {
          rec = sellers[k];
          break;
        }
      }
    }
  } else if (Array.isArray(sellers) && sellers.length) {
    rec = sellers[0] || {};
  }
  var stats = (rec && rec.sellerCategoryStatistics) || [];
  var rows = [];
  var i;
  for (i = 0; i < stats.length; i++) {
    var it = stats[i];
    if (!it || typeof it !== 'object') continue;
    var cid = String(it.catId || '').trim();
    var n = Number(it.productCount || 0);
    if (!(n > 0)) n = 0;
    if (cid && n > 0) rows.push([n, cid, it.avg30SalesRank]);
  }
  var tot = 0;
  for (i = 0; i < rows.length; i++) tot += rows[i][0];
  rows.sort(function (a, b) {
    return b[0] - a[0];
  });
  var parts = [];
  var foodN = 0;
  var lim = Math.min(rows.length, 20);
  for (i = 0; i < lim; i++) {
    var pct = tot ? (100.0 * rows[i][0]) / tot : 0;
    if (rows[i][1] === PR_FOOD_ROOT_) foodN += rows[i][0];
    var avg = rows[i][2];
    var avgS = avg === null || avg === undefined || avg === '' || avg === -1 ? '' : String(avg);
    parts.push(
      rows[i][1] + ':' + rows[i][0] + ':' + pct.toFixed(1) + '%' + (avgS ? ':r' + avgS : '')
    );
  }
  var foodPct = tot ? Math.round((1000.0 * foodN) / tot) / 10 : 0;
  var hypo = '構成比なし';
  if (tot && rows.length === 1 && rows[0][1] === PR_FOOD_ROOT_) {
    hypo = '食品のみ（卸の可能性・要人判断）';
  } else if (tot && foodPct >= 90) {
    hypo = '食品偏重（卸の可能性・要人判断）';
  } else if (tot) {
    hypo = '複数カテゴリ';
  }
  var tsf = rec && rec.totalStorefrontAsins;
  var storeN = '';
  if (Object.prototype.toString.call(tsf) === '[object Array]' && tsf.length >= 2) {
    storeN = String(tsf[1]);
  } else if (tsf !== null && tsf !== undefined && tsf !== '') {
    storeN = String(tsf);
  }
  var asinList = (rec && rec.asinList) || [];
  var hasAsinList = Array.isArray(asinList) && asinList.length > 0;
  if (hasAsinList) {
    Logger.log('[仕入れ検討①] /seller asinList unexpected n=%s seller=%s', asinList.length, sid);
  }
  return {
    name: String((rec && (rec.sellerName || rec.shopName)) || '').trim(),
    storeN: storeN,
    mix: parts.join('|'),
    foodPct: tot ? String(foodPct) : '',
    hypo: hypo,
    hasAsinList: hasAsinList,
  };
}

function prCellHas_(v) {
  if (v === 0 || v === false) return true;
  return String(v == null ? '' : v).trim() !== '';
}

function prSellerNameLocked_(memo) {
  return String(memo || '').indexOf('手修正') >= 0;
}

function prWriteLatest_(sh, row, col, value) {
  if (col < 0 || !prCellHas_(value)) return;
  sh.getRange(row, col + 1).setValue(value);
}

function prClearSellerLedgerCache_() {
  PR_SELLER_NAME_CACHE_ = null;
  PR_SELLER_ROW_CACHE_ = null;
  PR_SELLER_ROWS_CACHE_ = null;
  PR_SELLER_STORE_CACHE_ = null;
}

function prWriteIfEmpty_(sh, row, col, value) {
  if (col < 0 || !prCellHas_(value)) return;
  var old = sh.getRange(row, col + 1).getValue();
  if (!prCellHas_(old)) sh.getRange(row, col + 1).setValue(value);
}

/**
 * 同一 sellerId を keepRow にまとめ、余分な行を消す。メモは追記。sellerId・調査CKは残す。
 * @return {number} 残した行番号
 */
function prCollapseSellerDupes_(sellerId, keepRow, targets) {
  var sid = String(sellerId || '').trim();
  var keep = Number(keepRow || 0);
  if (!sid || keep < 2) return keep;
  prLoadSellerLedger_();
  var key = sid.toUpperCase();
  var sheetRows = (PR_SELLER_ROWS_CACHE_[key] || []).slice();
  if (sheetRows.indexOf(keep) < 0) sheetRows.push(keep);
  if (sheetRows.length < 2) return keep;
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
  if (!sh) return keep;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxMemo = prHeaderIndex_(hdr, 'メモ');
  var bodies = {};
  var i;
  var c;
  for (i = 0; i < sheetRows.length; i++) {
    bodies[sheetRows[i]] = sh.getRange(sheetRows[i], 1, 1, lastCol).getValues()[0];
  }
  var keepLine = bodies[keep].slice();
  for (i = 0; i < sheetRows.length; i++) {
    var other = sheetRows[i];
    if (other === keep) continue;
    var extra = bodies[other];
    for (c = 0; c < hdr.length; c++) {
      var hn = String(hdr[c] || '').trim();
      if (hn === '調査CK' || hn === 'sellerId') continue;
      if (hn === 'メモ') {
        keepLine[c] = prAppendMemo_(keepLine[c], extra[c]);
        continue;
      }
      if (PR_SELLER_SYNC_SKIP_[hn]) continue;
      if (prCellHas_(extra[c]) && (prSellerRewalkOverwriteEnabled_() || !prCellHas_(keepLine[c]))) {
        keepLine[c] = extra[c];
      }
    }
  }
  sh.getRange(keep, 1, 1, lastCol).setValues([keepLine]);
  var extras = [];
  for (i = 0; i < sheetRows.length; i++) {
    if (sheetRows[i] !== keep) extras.push(sheetRows[i]);
  }
  extras.sort(function (a, b) {
    return b - a;
  });
  var j;
  for (i = 0; i < extras.length; i++) {
    var del = extras[i];
    sh.deleteRow(del);
    if (keep > del) keep--;
    if (targets) {
      for (j = 0; j < targets.length; j++) {
        if (targets[j].sheetRow > del) targets[j].sheetRow--;
      }
    }
  }
  prClearSellerLedgerCache_();
  Logger.log('[仕入れ検討①] 同一店1行 seller=%s keep=%s del=%s', sid, keep, extras.length);
  return keep;
}

/** 同一 sellerId の空セルへ、店名行の構成比などをコピー。採取元・メモ・巡回日は触らない。 */
function prSyncSellerDupes_(sellerId) {
  var sid = String(sellerId || '').trim();
  if (!sid) return;
  prLoadSellerLedger_();
  var key = sid.toUpperCase();
  var sheetRows = PR_SELLER_ROWS_CACHE_[key] || [];
  if (sheetRows.length < 2) return;
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
  if (!sh) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var bodies = [];
  var i;
  var c;
  for (i = 0; i < sheetRows.length; i++) {
    bodies.push(sh.getRange(sheetRows[i], 1, 1, lastCol).getValues()[0]);
  }
  var best = 0;
  var bestScore = -1;
  for (i = 0; i < bodies.length; i++) {
    var score = 0;
    for (c = 0; c < hdr.length; c++) {
      var h = String(hdr[c] || '').trim();
      if (PR_SELLER_SYNC_SKIP_[h]) continue;
      if (prCellHas_(bodies[i][c])) score++;
    }
    if (score > bestScore) {
      bestScore = score;
      best = i;
    }
  }
  if (bestScore <= 0) return;
  var donor = bodies[best];
  var changed = 0;
  for (i = 0; i < bodies.length; i++) {
    if (i === best) continue;
    var line = bodies[i].slice();
    var rowChanged = false;
    for (c = 0; c < hdr.length; c++) {
      var hn = String(hdr[c] || '').trim();
      if (PR_SELLER_SYNC_SKIP_[hn]) continue;
      if (!prCellHas_(donor[c])) continue;
      if (prCellHas_(line[c])) continue;
      line[c] = donor[c];
      rowChanged = true;
    }
    if (rowChanged) {
      sh.getRange(sheetRows[i], 1, 1, lastCol).setValues([line]);
      changed++;
    }
  }
  if (changed) {
    Logger.log('[仕入れ検討①] 同一店コピー seller=%s fromRow=%s n=%s', sid, sheetRows[best], changed);
  }
}

/** 台帳へ店名・ストアASIN数・構成比。再調査時は今回値で上書き。空の新値では消さない。query件数は書かない。 */
function prWriteSellerStorefront_(sellerId, parsed) {
  var sid = String(sellerId || '').trim();
  if (!sid || !parsed) return;
  var nm = String(parsed.name || '').trim();
  prLoadSellerLedger_();
  var key = sid.toUpperCase();
  if (nm) PR_SELLER_NAME_CACHE_[key] = nm;
  if (parsed.storeN) PR_SELLER_STORE_CACHE_[key] = String(parsed.storeN);
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
  if (!sh) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxId = prHeaderIndex_(hdr, 'sellerId');
  var idxName = prHeaderIndex_(hdr, '店名');
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  var idxStore = prHeaderIndex_(hdr, 'ストアASIN数');
  var idxMix = prHeaderIndex_(hdr, 'カテゴリ構成');
  var idxFood = prHeaderIndex_(hdr, '食品比率');
  var idxHypo = prHeaderIndex_(hdr, '卸仮説');
  var idxMemo = prHeaderIndex_(hdr, 'メモ');
  var idxTargetCat = prHeaderIndex_(hdr, '対象カテゴリ');
  var idxMainCat = prHeaderIndex_(hdr, 'メインカテゴリ');
  var idxFoodCat = prHeaderIndex_(hdr, '食品・飲料・お酒');
  if (idxId < 0) return;
  var sheetRows = (PR_SELLER_ROWS_CACHE_[key] || []).slice();
  var preferred = PR_SELLER_ROW_CACHE_[key];
  if (preferred && sheetRows.indexOf(preferred) < 0) sheetRows.push(preferred);
  var overwrite = prSellerRewalkOverwriteEnabled_();
  var writer = overwrite ? prWriteLatest_ : prWriteIfEmpty_;
  var foodPctNum = Number(parsed.foodPct);
  if (isNaN(foodPctNum)) foodPctNum = 0;
  var mainCatVal = '食品・飲料・お酒 ' + foodPctNum.toFixed(1) + '%';
  if (sheetRows.length) {
    var locked = false;
    var j;
    if (idxMemo >= 0) {
      for (j = 0; j < sheetRows.length; j++) {
        if (prSellerNameLocked_(sh.getRange(sheetRows[j], idxMemo + 1).getValue())) {
          locked = true;
          break;
        }
      }
    }
    for (j = 0; j < sheetRows.length; j++) {
      var sheetRow = sheetRows[j];
      if (idxName >= 0 && !locked) writer(sh, sheetRow, idxName, nm);
      if (idxStore >= 0) writer(sh, sheetRow, idxStore, parsed.storeN);
      if (idxMix >= 0) writer(sh, sheetRow, idxMix, parsed.mix);
      if (idxFood >= 0) writer(sh, sheetRow, idxFood, parsed.foodPct);
      if (idxHypo >= 0) writer(sh, sheetRow, idxHypo, parsed.hypo);
      if (idxTargetCat >= 0) writer(sh, sheetRow, idxTargetCat, PR_FOOD_ROOT_);
      if (idxMainCat >= 0) writer(sh, sheetRow, idxMainCat, mainCatVal);
      if (idxFoodCat >= 0) writer(sh, sheetRow, idxFoodCat, foodPctNum);
    }
    return;
  }
  var line = [];
  var i;
  for (i = 0; i < lastCol; i++) line.push('');
  if (idxCk >= 0) line[idxCk] = false;
  line[idxId] = sid;
  if (idxName >= 0) line[idxName] = nm;
  if (idxStore >= 0) line[idxStore] = parsed.storeN || '';
  if (idxMix >= 0) line[idxMix] = parsed.mix || '';
  if (idxFood >= 0) line[idxFood] = parsed.foodPct || '';
  if (idxHypo >= 0) line[idxHypo] = parsed.hypo || '';
  if (idxTargetCat >= 0) line[idxTargetCat] = PR_FOOD_ROOT_;
  if (idxMainCat >= 0) line[idxMainCat] = mainCatVal;
  if (idxFoodCat >= 0) line[idxFoodCat] = foodPctNum;
  var newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, lastCol).setValues([line]);
  if (idxCk >= 0) sh.getRange(newRow, idxCk + 1).insertCheckboxes();
  PR_SELLER_ROW_CACHE_[key] = newRow;
  PR_SELLER_ROWS_CACHE_[key] = [newRow];
}

/**
 * 店名またはストアASIN数が台帳に無ければ /seller storefront=0 を1回。
 * 同じ応答で構成比も残す。storefront=1 は使わない。
 * @return {string} 店名
 */
function prEnsureSellerStorefront_(sellerId, force) {
  var sid = String(sellerId || '').trim();
  if (!sid) return '';
  var key = sid.toUpperCase();
  var refresh = !!force && prSellerRewalkOverwriteEnabled_();
  prLoadSellerLedger_();
  if (!refresh) prSyncSellerDupes_(sid);
  var haveName = !!(PR_SELLER_NAME_CACHE_[key]);
  var haveStore = !!(PR_SELLER_STORE_CACHE_[key]);
  if (!refresh && haveName && haveStore) return PR_SELLER_NAME_CACHE_[key];
  if (haveName && PR_SELLER_STORE_MISS_[key]) return PR_SELLER_NAME_CACHE_[key];
  if (!haveName && PR_SELLER_NAME_MISS_[key]) return '';
  if (!prSellerNameGetEnabled_()) return haveName ? PR_SELLER_NAME_CACHE_[key] : '';
  if (PR_SELLER_NAME_GET_N_ >= prSellerNameGetCap_()) {
    Logger.log('[仕入れ検討①] 店GET cap 到達 leftover seller=%s', sid);
    return haveName ? PR_SELLER_NAME_CACHE_[key] : '';
  }
  PR_SELLER_NAME_GET_N_++;
  var parsed = null;
  try {
    parsed = prKeepaSellerGet_(sid);
  } catch (e) {
    Logger.log(
      '[仕入れ検討①] 店GET失敗 seller=%s %s',
      sid,
      prRedactKey_(String((e && e.message) || e)).slice(0, 160)
    );
    if (!haveName) PR_SELLER_NAME_MISS_[key] = true;
    PR_SELLER_STORE_MISS_[key] = true;
    return haveName ? PR_SELLER_NAME_CACHE_[key] : '';
  }
  if (!parsed) {
    if (!haveName) PR_SELLER_NAME_MISS_[key] = true;
    PR_SELLER_STORE_MISS_[key] = true;
    return haveName ? PR_SELLER_NAME_CACHE_[key] : '';
  }
  if (!parsed.name && !haveName) PR_SELLER_NAME_MISS_[key] = true;
  if (!parsed.storeN) PR_SELLER_STORE_MISS_[key] = true;
  prWriteSellerStorefront_(sid, parsed);
  prSyncSellerDupes_(sid);
  return (PR_SELLER_NAME_CACHE_[key] || parsed.name || '');
}

function prKeepaSellerNameGet_(sellerId) {
  var parsed = prKeepaSellerGet_(sellerId);
  return parsed && parsed.name ? parsed.name : '';
}

function prWriteSellerName_(sellerId, name) {
  prWriteSellerStorefront_(sellerId, { name: name, storeN: '', mix: '', foodPct: '', hypo: '' });
}

/** 台帳の店名だけ返す。Keepa は呼ばない（まとめ・upsert用）。ストアASIN数はレ点の店で取る。 */
function prEnsureSellerName_(sellerId) {
  var sid = String(sellerId || '').trim();
  if (!sid) return '';
  prLoadSellerLedger_();
  return PR_SELLER_NAME_CACHE_[sid.toUpperCase()] || '';
}

function prSellerNameForId_(sellerId) {
  return prEnsureSellerName_(sellerId);
}

function prNamesForSellerIds_(sellerCell) {
  var ids = prSplitTok_(sellerCell);
  var names = [];
  var i;
  for (i = 0; i < ids.length; i++) {
    var n = prEnsureSellerName_(ids[i]);
    if (n) names.push(n);
  }
  return prJoinTok_(names);
}

function prPreferGate_(oldSt, oldWhy, newSt, newWhy) {
  var o = String(oldSt || '').trim();
  var n = String(newSt || '').trim();
  if (o === '通過' && n !== '通過') return { st: o, why: String(oldWhy || '') };
  if (n) return { st: n, why: String(newWhy || '') };
  return { st: o, why: String(oldWhy || '') };
}

function prRegateFromKeepaFullEnabled_() {
  return prCkEnabled_(PR_REGATE_FROM_FULL_ENABLED_PROP_);
}

/** 既定 true＝門結果空白のみ。Property PR_REGATE_ONLY_BLANK=false で全行再門。 */
function prRegateOnlyBlankGate_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_REGATE_ONLY_BLANK_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

/** Keepaフル 生JSON→product。90日判定なし（再門用）。 */
function prProductFromKeepaFullAsin_(asin) {
  var a = String(asin || '')
    .trim()
    .toUpperCase();
  if (!a) return null;
  var map = prLoadKeepaFullMap_();
  var rec = map[a];
  if (!rec || !rec.product) return null;
  return rec.product;
}

/**
 * ①候補は 1ASIN1行。sellerId は被ったら | で同じセルへ足す。重複行は先頭にまとめ、余りを消す。
 * @return {{kept:number, merged:number, deleted:number}}
 */
function prCompactCandDupes_() {
  var sh = prEnsureSheet_(PR_CAND_SHEET_, PR_CAND_HEADERS_);
  var last = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = {};
  var i;
  for (i = 0; i < hdr.length; i++) idx[String(hdr[i] || '')] = i;
  if (idx['店名'] === undefined) {
    sh.getRange(1, hdr.length + 1).setValue('店名');
    hdr.push('店名');
    idx['店名'] = hdr.length - 1;
    lastCol = hdr.length;
  }
  if (idx['ASIN'] === undefined || last < 2) {
    return { kept: 0, merged: 0, deleted: 0, names: 0, getN: 0 };
  }
  PR_SELLER_NAME_GET_N_ = 0;
  PR_SELLER_NAME_MISS_ = {};
  PR_SELLER_STORE_MISS_ = {};
  var vals = sh.getRange(2, 1, last - 1, lastCol).getValues();
  var groups = {};
  var order = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    var asin = String(vals[r][idx['ASIN']] || '').trim().toUpperCase();
    if (!asin) continue;
    if (!groups[asin]) {
      groups[asin] = [];
      order.push(asin);
    }
    groups[asin].push({ sheetRow: r + 2, line: vals[r].slice() });
  }
  var merged = 0;
  var extraRows = [];
  for (i = 0; i < order.length; i++) {
    var g = groups[order[i]];
    if (g.length < 2) continue;
    merged++;
    var keep = g[0].line;
    var si = idx['sellerId'];
    var ni = idx['店名'];
    var pi = idx['発見経路'];
    var gi = idx['門結果'];
    var wi = idx['門理由'];
    var j;
    for (j = 1; j < g.length; j++) {
      var line = g[j].line;
      if (si !== undefined) keep[si] = prMergeTok_(keep[si], line[si]);
      if (ni !== undefined) keep[ni] = prMergeTok_(keep[ni], line[ni]);
      if (pi !== undefined) keep[pi] = prMergeTok_(keep[pi], line[pi]);
      if (gi !== undefined) {
        var g2 = prPreferGate_(keep[gi], wi !== undefined ? keep[wi] : '', line[gi], wi !== undefined ? line[wi] : '');
        keep[gi] = g2.st;
        if (wi !== undefined) keep[wi] = g2.why;
      }
      extraRows.push(g[j].sheetRow);
    }
    sh.getRange(g[0].sheetRow, 1, 1, keep.length).setValues([keep]);
  }
  extraRows.sort(function (a, b) {
    return b - a;
  });
  for (i = 0; i < extraRows.length; i++) {
    sh.deleteRow(extraRows[i]);
  }
  var filled = prFillCandShopNames_(sh);
  Logger.log(
    '[仕入れ検討①] ①候補 compact kept=%s mergedAsin=%s deleted=%s names=%s getN=%s',
    order.length,
    merged,
    extraRows.length,
    filled,
    PR_SELLER_NAME_GET_N_
  );
  return {
    kept: order.length,
    merged: merged,
    deleted: extraRows.length,
    names: filled,
    getN: PR_SELLER_NAME_GET_N_,
  };
}

function prFillCandShopNames_(sh) {
  var last = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxId = prHeaderIndex_(hdr, 'sellerId');
  var idxName = prHeaderIndex_(hdr, '店名');
  if (idxId < 0 || idxName < 0 || last < 2) return 0;
  var vals = sh.getRange(2, 1, last - 1, lastCol).getValues();
  var n = 0;
  var r;
  for (r = 0; r < vals.length; r++) {
    var looked = prNamesForSellerIds_(vals[r][idxId]);
    var mergedNames = prMergeTok_(vals[r][idxName], looked);
    if (String(mergedNames) !== String(vals[r][idxName] || '')) {
      vals[r][idxName] = mergedNames;
      n++;
    }
  }
  if (n) sh.getRange(2, 1, vals.length, lastCol).setValues(vals);
  return n;
}

function prMenuCompactCand() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  try {
    var r = prCompactCandDupes_();
    var detail =
      'asin=' +
      r.kept +
      ' merged=' +
      r.merged +
      ' deleted=' +
      r.deleted +
      ' names=' +
      (r.names || 0) +
      ' getN=' +
      (r.getN || 0);
    if ((r.getN || 0) >= prSellerNameGetCap_()) detail += ' leftover=cap再実行';
    prAppendLog_([runId, '①候補 ASINまとめ', 'DONE', new Date(), '', r.kept, '', detail]);
    prMenuToast_('①候補 まとめ', detail);
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    prAppendLog_([runId, '①候補 ASINまとめ', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

function prCandMemGet_(sh, headers, lastCol) {
  var colN = Math.max(headers.length, lastCol);
  if (PR_CAND_MEM_ && PR_CAND_MEM_.shId === sh.getSheetId() && PR_CAND_MEM_.colN === colN) {
    return PR_CAND_MEM_;
  }
  var last = sh.getLastRow();
  var vals = last >= 2 ? sh.getRange(2, 1, last - 1, colN).getValues() : [];
  var idx = {};
  var i;
  for (i = 0; i < headers.length; i++) idx[String(headers[i])] = i;
  PR_CAND_MEM_ = { shId: sh.getSheetId(), colN: colN, idx: idx, vals: vals };
  return PR_CAND_MEM_;
}

function prUpsertCandidate_(row, opt) {
  opt = opt || {};
  var sh = prEnsureSheet_(PR_CAND_SHEET_, PR_CAND_HEADERS_);
  var need = PR_CAND_HEADERS_;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var i;
  var headerAdded = false;
  for (i = 0; i < need.length; i++) {
    var found = false;
    var c;
    for (c = 0; c < headers.length; c++) {
      if (String(headers[c]) === need[i]) found = true;
    }
    if (!found) {
      sh.getRange(1, headers.length + 1).setValue(need[i]);
      headers.push(need[i]);
      headerAdded = true;
    }
  }
  if (headerAdded) PR_CAND_MEM_ = null;
  var mem = prCandMemGet_(sh, headers, lastCol);
  var idx = mem.idx;
  var vals = mem.vals;
  var asin = String(row[1] || '').toUpperCase();
  if (!asin) return;
  var line = [];
  for (i = 0; i < headers.length; i++) line.push('');
  var rec = {};
  for (i = 0; i < need.length; i++) rec[need[i]] = row[i];
  for (i = 0; i < need.length; i++) {
    if (idx[need[i]] === undefined) continue;
    line[idx[need[i]]] = rec[need[i]];
  }
  if (vals.length && idx['ASIN'] !== undefined) {
    var hit = -1;
    var extras = [];
    var r;
    for (r = 0; r < vals.length; r++) {
      if (String(vals[r][idx['ASIN']] || '').toUpperCase() !== asin) continue;
      if (hit < 0) hit = r;
      else extras.push(r);
    }
    if (hit >= 0) {
      var old = vals[hit];
      if (idx['sellerId'] !== undefined) {
        line[idx['sellerId']] = prMergeTok_(old[idx['sellerId']], line[idx['sellerId']]);
      }
      if (idx['店名'] !== undefined) {
        line[idx['店名']] = prMergeTok_(
          old[idx['店名']],
          prMergeTok_(line[idx['店名']], prNamesForSellerIds_(line[idx['sellerId']]))
        );
      }
      if (idx['発見経路'] !== undefined) {
        line[idx['発見経路']] = prMergeTok_(old[idx['発見経路']], line[idx['発見経路']]);
      }
      if (idx['門結果'] !== undefined && !opt.forceGate) {
        var g = prPreferGate_(
          old[idx['門結果']],
          idx['門理由'] !== undefined ? old[idx['門理由']] : '',
          line[idx['門結果']],
          idx['門理由'] !== undefined ? line[idx['門理由']] : ''
        );
        line[idx['門結果']] = g.st;
        if (idx['門理由'] !== undefined) line[idx['門理由']] = g.why;
      }
      if (idx['メーカー'] !== undefined) {
        var om = String(old[idx['メーカー']] || '').trim();
        var nm = String(line[idx['メーカー']] || '').trim();
        if (om && !nm) line[idx['メーカー']] = om;
      }
      if (prCycleListCheckEnabled_() && idx[PR_CAND_LIST_CHECK_] !== undefined && idx['門結果'] !== undefined) {
        var gateUp = String(line[idx['門結果']] || '');
        var oldCk = String(old[idx[PR_CAND_LIST_CHECK_]] || '').trim();
        if (gateUp === '通過') {
          if (oldCk === PR_CAND_LIST_CHECK_SUMI_ || oldCk === PR_CAND_LIST_CHECK_MORE_) {
            line[idx[PR_CAND_LIST_CHECK_]] = oldCk;
          } else {
            line[idx[PR_CAND_LIST_CHECK_]] = PR_CAND_LIST_CHECK_MI_;
          }
        } else {
          line[idx[PR_CAND_LIST_CHECK_]] = '';
        }
      }
      var e;
      for (e = 0; e < extras.length; e++) {
        var ex = vals[extras[e]];
        if (idx['sellerId'] !== undefined) {
          line[idx['sellerId']] = prMergeTok_(line[idx['sellerId']], ex[idx['sellerId']]);
        }
        if (idx['店名'] !== undefined) {
          line[idx['店名']] = prMergeTok_(line[idx['店名']], ex[idx['店名']]);
        }
        if (idx['発見経路'] !== undefined) {
          line[idx['発見経路']] = prMergeTok_(line[idx['発見経路']], ex[idx['発見経路']]);
        }
      }
      if (extras.length && idx['店名'] !== undefined && idx['sellerId'] !== undefined) {
        line[idx['店名']] = prMergeTok_(line[idx['店名']], prNamesForSellerIds_(line[idx['sellerId']]));
      }
      sh.getRange(hit + 2, 1, 1, line.length).setValues([line]);
      mem.vals[hit] = line;
      if (extras.length) {
        extras.sort(function (a, b) {
          return b - a;
        });
        for (e = 0; e < extras.length; e++) {
          sh.deleteRow(extras[e] + 2);
        }
        PR_CAND_MEM_ = null;
      }
      return;
    }
  }
  if (idx['店名'] !== undefined && idx['sellerId'] !== undefined) {
    line[idx['店名']] = prMergeTok_(line[idx['店名']], prNamesForSellerIds_(line[idx['sellerId']]));
  }
  if (prCycleListCheckEnabled_() && idx[PR_CAND_LIST_CHECK_] !== undefined && idx['門結果'] !== undefined) {
    if (String(line[idx['門結果']] || '') === '通過') {
      line[idx[PR_CAND_LIST_CHECK_]] = PR_CAND_LIST_CHECK_MI_;
    } else {
      line[idx[PR_CAND_LIST_CHECK_]] = '';
    }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, line.length).setValues([line]);
  mem.vals.push(line);
}

/** Script Property PR_MAKER_WASH_ENABLED。未設定=ON。false で切戻し。 */
function prMakerWashEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_MAKER_WASH_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prMakerWashCap_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_MAKER_WASH_CAP_PROP_) || PR_MAKER_WASH_CAP_DEFAULT_
  );
  if (!(n > 0)) n = PR_MAKER_WASH_CAP_DEFAULT_;
  if (n > PR_MAKER_WASH_CAP_MAX_) n = PR_MAKER_WASH_CAP_MAX_;
  return Math.floor(n);
}

/** Keepa refillRate（/分）。応答優先。フォールバックは実測5。 */
function prKeepaRefillRate_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty('KEEPA_REFILL_RATE') || PR_KEEPA_REFILL_RATE_DEFAULT_
  );
  if (!(n > 0)) n = PR_KEEPA_REFILL_RATE_DEFAULT_;
  return Math.floor(n);
}

function prKeepaTokenBucketMax_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_KEEPA_TOKEN_BUCKET_MAX_PROP_) ||
      PR_KEEPA_TOKEN_BUCKET_MAX_DEFAULT_
  );
  if (!(n > 0)) n = PR_KEEPA_TOKEN_BUCKET_MAX_DEFAULT_;
  return Math.floor(n);
}

function prKeepaProductTokenCost_() {
  return prKeepaCostProductOffers20_();
}

/** offers=20 付き /product の ASIN 単価（既定6・Property 可）。 */
function prKeepaCostProductOffers20_() {
  return prKeepaProfileUnit_(PR_KEEPA_COST_PROFILE_OFFERS20_);
}

/**
 * プロファイル別単価。将来 stats のみ等を足すときはここに分岐。
 * @param {string} profile
 */
function prKeepaProfileUnit_(profile) {
  var p = String(profile || PR_KEEPA_COST_PROFILE_OFFERS20_);
  var props = PropertiesService.getScriptProperties();
  var n;
  if (p === PR_KEEPA_COST_PROFILE_OFFERS20_ || p === 'offers20') {
    n = Number(
      props.getProperty(PR_KEEPA_PRODUCT_TOKEN_PROP_) || PR_KEEPA_COST_PRODUCT_OFFERS20_DEFAULT_
    );
    if (!(n > 0)) n = PR_KEEPA_COST_PRODUCT_OFFERS20_DEFAULT_;
    return n;
  }
  if (p === PR_KEEPA_COST_PROFILE_STATS90_) {
    n = Number(
      props.getProperty(PR_KEEPA_STATS90_TOKEN_PROP_) || PR_KEEPA_COST_PRODUCT_STATS90_DEFAULT_
    );
    if (!(n > 0)) n = PR_KEEPA_COST_PRODUCT_STATS90_DEFAULT_;
    return n;
  }
  if (p === PR_KEEPA_COST_PROFILE_TWO_STAGE_) {
    n = Number(
      props.getProperty(PR_KEEPA_TWO_STAGE_TOKEN_PROP_) || PR_KEEPA_COST_PRODUCT_TWO_STAGE_DEFAULT_
    );
    if (!(n > 0)) n = PR_KEEPA_COST_PRODUCT_TWO_STAGE_DEFAULT_;
    return n;
  }
  return PR_KEEPA_COST_PRODUCT_OFFERS20_DEFAULT_;
}

/**
 * 実測単価 = consumed/件数。異常値はプロファイル既定。
 * @param {number} tokensConsumed
 * @param {number} asinCount
 * @param {string=} profile
 */
function prKeepaMeasuredUnit_(tokensConsumed, asinCount, profile) {
  var n = Number(asinCount);
  var c = Number(tokensConsumed);
  if (!(n > 0) || !(c > 0)) return prKeepaProfileUnit_(profile);
  var u = c / n;
  if (!(u >= 0.5) || u > 30) return prKeepaProfileUnit_(profile);
  return Math.round(u * 10) / 10;
}

/** 待ち分数 = ceil(consumed/refillRate)。最低1。 */
function prKeepaWaitMinutes_(tokensConsumed, refillRate) {
  var rate = Number(refillRate);
  if (!(rate > 0)) rate = prKeepaRefillRate_();
  var consumed = Number(tokensConsumed);
  if (!(consumed > 0)) return 1;
  return Math.max(1, Math.ceil(consumed / rate));
}

/**
 * 洗い完走の再開待ち分。
 * 単価: opt.unitCost → 実測(consumed/asinCount) → プロファイル既定（offers20=6）。
 * need = 次スライス件数(min(残り, cap))×単価（上限バケツMAX）。deficit = need - max(0, tokensLeft)。
 * wait = ceil(deficit/rate)+1。上限 = ceil(MAX/rate)+1。
 * 2026-08-29: 残り全体×単価だと特大で常に MAX61 になるため、次 cap 分のみに変更。
 */
function prKeepaWashResumeWaitMin_(opt) {
  var rate = Number(opt && opt.refillRate);
  if (!(rate > 0)) rate = prKeepaRefillRate_();
  var profile = (opt && opt.profile) || PR_KEEPA_COST_PROFILE_OFFERS20_;
  var unit = Number(opt && opt.unitCost);
  if (!(unit > 0)) {
    unit = prKeepaMeasuredUnit_(opt && opt.tokensConsumed, opt && opt.asinCount, profile);
  }
  if (!(unit > 0)) unit = prKeepaProfileUnit_(profile);
  var bucketMax = prKeepaTokenBucketMax_();
  var pending = Math.max(0, Number(opt && opt.pendingCount) || 0);
  var cap = Number(opt && opt.cap);
  if (!(cap > 0)) cap = prMakerWashCap_();
  if (cap > PR_MAKER_WASH_CAP_MAX_) cap = PR_MAKER_WASH_CAP_MAX_;
  var nextN = pending > 0 ? Math.min(pending, cap) : 0;
  var need = Math.min(nextN * unit, bucketMax);
  var leftRaw = opt && opt.tokensLeft;
  var left = Number(leftRaw);
  if (leftRaw === '' || leftRaw == null || left !== left) left = 0;
  // Keepa 応答の負残は「実質0」として扱う（deficit が MAX に張り付くのを防ぐ）
  if (left < 0) left = 0;
  var deficit = need - left;
  if (deficit < 0) deficit = 0;
  if (deficit > bucketMax) deficit = bucketMax;
  var waitMin = Math.ceil(deficit / rate) + 1;
  if (!(waitMin > 0)) waitMin = 1;
  var maxWait = Math.ceil(bucketMax / rate) + 1;
  if (waitMin > maxWait) waitMin = maxWait;
  if (opt && opt.rateLimited && Number(opt.refillIn) > 0) {
    var fromRefill = Math.ceil(Number(opt.refillIn) / 60000) + 1;
    if (fromRefill > waitMin) waitMin = fromRefill;
    if (waitMin > maxWait) waitMin = maxWait;
  }
  Logger.log(
    '[仕入れ検討①] resumeWait pending=%s nextN=%s cap=%s unit=%s need=%s left=%s deficit=%s rate=%s waitMin=%s max=%s profile=%s',
    pending,
    nextN,
    cap,
    unit,
    need,
    left,
    deficit,
    rate,
    waitMin,
    maxWait,
    profile
  );
  return Math.floor(waitMin);
}

function prMakerWashDetailIsComplete_(detail) {
  var s = String(detail || '');
  return s.indexOf('COMPLETE') >= 0 || s.indexOf('完走 parent=') >= 0;
}

function prMakerWashFullEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_MAKER_WASH_FULL_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prEnsureWashQueueSheet_() {
  return prEnsureSheet_(PR_WASH_QUEUE_SHEET_, PR_WASH_QUEUE_HEADERS_);
}

/** Keepa済キュー行の保持日数（既定90）。0以下は掃除オフ。 */
function prWashQueueRetentionDays_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_WASH_QUEUE_RETENTION_DAYS_PROP_);
  if (raw == null || String(raw).trim() === '') return PR_WASH_QUEUE_RETENTION_DAYS_DEFAULT_;
  var n = Number(String(raw).trim());
  if (!isFinite(n)) return PR_WASH_QUEUE_RETENTION_DAYS_DEFAULT_;
  return Math.floor(n);
}

/** 列挙日時セル → Date。空・不正は null。 */
function prParseWashQueueEnumDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (v == null || v === '') return null;
  if (typeof v === 'number' && isFinite(v)) {
    var epoch = new Date(v);
    return isNaN(epoch.getTime()) ? null : epoch;
  }
  var s = String(v).trim();
  if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  var m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    var d2 = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(d2.getTime()) ? null : d2;
  }
  return null;
}

/**
 * Keepa済かつ列挙日時が保持日数より古い行を削除。未済は残す。列挙日時空は残す。
 * days<=0 なら何もしない。
 * @param {{quiet?:boolean, runId?:string, days?:number}} opt
 * @return {{ok:boolean, deleted:number, days:number, detail:string}}
 */
function prPurgeExpiredWashQueueRows_(opt) {
  var quiet = !!(opt && opt.quiet);
  var runId =
    (opt && opt.runId) || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var days = opt && opt.days != null ? Number(opt.days) : prWashQueueRetentionDays_();
  if (!isFinite(days)) days = PR_WASH_QUEUE_RETENTION_DAYS_DEFAULT_;
  days = Math.floor(days);
  if (days <= 0) {
    var offDetail = 'retention_off days=' + days;
    prAppendLog_([runId, '洗いキュー期限切れ削除', 'DONE', new Date(), '', 0, '', offDetail]);
    return { ok: true, deleted: 0, days: days, detail: offDetail };
  }
  var sh = prEnsureWashQueueSheet_();
  if (sh.getLastRow() < 2) {
    var emptyDetail = 'empty days=' + days;
    if (!quiet) prMenuToast_('① 洗いキュー削除', emptyDetail);
    prAppendLog_([runId, '洗いキュー期限切れ削除', 'DONE', new Date(), '', 0, '', emptyDetail]);
    return { ok: true, deleted: 0, days: days, detail: emptyDetail };
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxK = prHeaderIndex_(hdr, 'Keepa済');
  var idxDt = prHeaderIndex_(hdr, '列挙日時');
  if (idxK < 0 || idxDt < 0) {
    throw new Error('①洗いキューに Keepa済／列挙日時 列がありません');
  }
  var cutoff = new Date();
  cutoff.setTime(cutoff.getTime() - days * 24 * 60 * 60 * 1000);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var delRows = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    if (!prCkIsTrue_(vals[r][idxK])) continue;
    var dt = prParseWashQueueEnumDate_(vals[r][idxDt]);
    if (!dt) continue;
    if (dt.getTime() < cutoff.getTime()) delRows.push(r + 2);
  }
  delRows.sort(function (x, y) {
    return y - x;
  });
  var d;
  for (d = 0; d < delRows.length; d++) {
    sh.deleteRow(delRows[d]);
  }
  if (delRows.length) {
    prWashQueueCountCacheClear_();
    prWashQueueSnapshotClear_();
  }
  var detail = 'deleted=' + delRows.length + ' days=' + days + ' cutoff=' + cutoff.toISOString();
  Logger.log('[仕入れ検討①] runId=%s wash_queue_purge %s', runId, detail);
  prAppendLog_([runId, '洗いキュー期限切れ削除', 'DONE', new Date(), '', delRows.length, '', detail]);
  if (!quiet) prMenuToast_('① 洗いキュー削除', detail);
  return { ok: true, deleted: delRows.length, days: days, detail: detail };
}

function prMenuWashQueuePurgeExpired() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  try {
    prPurgeExpiredWashQueueRows_({ runId: runId, quiet: false });
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    prAppendLog_([runId, '洗いキュー期限切れ削除', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/** メーカーサマリに洗い状態／進捗列を確保。 */
function prEnsureMakerWashStatusCols_(sh) {
  if (!sh) return { idxSt: -1, idxProg: -1 };
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
  var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
  if (idxSt < 0) {
    sh.getRange(1, lastCol + 1).setValue(PR_WASH_STATUS_HEADER_);
    idxSt = lastCol;
    lastCol++;
  }
  if (idxProg < 0) {
    sh.getRange(1, lastCol + 1).setValue(PR_WASH_PROGRESS_HEADER_);
    idxProg = lastCol;
  }
  return { idxSt: idxSt, idxProg: idxProg };
}

function prFindMakerSumRow_(parent) {
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return null;
  var cols = prEnsureMakerWashStatusCols_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  if (idxMaker < 0) return null;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxMaker] || '').trim() === parent) {
      return {
        sh: sh,
        sheetRow: r + 2,
        idxSt: cols.idxSt,
        idxProg: cols.idxProg,
        idxDay: prHeaderIndex_(hdr, '最終調査日'),
        idxCk: prHeaderIndex_(hdr, PR_CK_HEADER_),
      };
    }
  }
  return null;
}

/** 洗い完走時: メーカーサマリの調査CKを外す（トリガー完走含む）。 */
function prClearMakerSumCk_(parent) {
  var loc = prFindMakerSumRow_(parent);
  if (!loc || loc.idxCk < 0) return false;
  loc.sh.getRange(loc.sheetRow, loc.idxCk + 1).setValue(false);
  Logger.log('[仕入れ検討①] maker_ck_clear parent=%s row=%s', parent, loc.sheetRow);
  return true;
}

/**
 * Catalog見積≥1 なのに調査ASIN未済ゼロ → リサーチしない（空ループ防止）。
 * progress に empty_research / done_no_candidates が付いた Keepa中も対象外。
 */
function prWashEmptyResearchBlocked_(st, prog) {
  st = st != null ? String(st).trim() : '';
  prog = prog != null ? String(prog) : '';
  if (st === PR_WASH_ST_HOLD_ && prog.indexOf('empty_research') >= 0) return true;
  if (prog.indexOf('empty_research') >= 0) return true;
  if (prog.indexOf('done_no_candidates') >= 0) return true;
  return false;
}

/** 選定プールから外す保留（empty_research／title0／要確認等）。 */
function prWashHoldBlockedForSelect_(st, prog) {
  st = st != null ? String(st).trim() : '';
  prog = prog != null ? String(prog) : '';
  if (st === PR_WASH_ST_HOLD_) return true;
  if (prWashEmptyResearchBlocked_(st, prog)) return true;
  return false;
}

/**
 * サマリ Catalog見積セルが正の数か（列挙完了の印）。
 * キュー件数だけでは Keepa 開始しない（Catalog先行）。
 */
function prMakerCatalogEstCellOk_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return false;
  var loc = prFindMakerSumRow_(parent);
  if (!loc || !loc.sh || !(loc.sheetRow > 0)) return false;
  var lastCol = Math.max(loc.sh.getLastColumn(), 1);
  var hdr = loc.sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxEst = prHeaderIndex_(hdr, PR_CATALOG_EST_COL_);
  if (idxEst < 0) return false;
  var cell = loc.sh.getRange(loc.sheetRow, idxEst + 1).getValue();
  if (cell === undefined || cell === null || String(cell).trim() === '') return false;
  var n = Number(String(cell).replace(/[^\d.-]/g, ''));
  return !isNaN(n) && n > 0;
}

/**
 * Catalog見積未了の洗えるメーカー先頭（規模不明含む）。Keepa新規より列挙優先。
 * @return {{parent:string, detail:string}|null}
 */
function prCycleFirstCatalogPending_() {
  var list = prCycleCatalogPendingList_(1);
  return list.length ? list[0] : null;
}

/**
 * Catalog見積未了の洗えるメーカーを最大 cap 社（規模不明を先、次に帯あり見積なし）。
 * 絞り込み条件は先頭1社版と同じ（`prCycleFirstCatalogPending_` はこれの cap=1）。
 * @param {number} cap
 * @return {!Array<{parent:string, detail:string}>}
 */
function prCycleCatalogPendingList_(cap) {
  var want = Math.max(1, Number(cap) || 1);
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return [];
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  var idxJ = prHeaderIndex_(hdr, '実態');
  var idxDay = prHeaderIndex_(hdr, '最終調査日');
  var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
  var idxEst = prHeaderIndex_(hdr, PR_CATALOG_EST_COL_);
  var idxBand = prHeaderIndex_(hdr, PR_BAND_COL_);
  var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
  if (idxMaker < 0) return [];
  var sellerVals = [];
  var sellerHdr = [];
  try {
    var sSh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_SELLER_SHEET_);
    if (sSh && sSh.getLastRow() >= 2) {
      var sLastCol = Math.max(sSh.getLastColumn(), 1);
      sellerHdr = sSh.getRange(1, 1, 1, sLastCol).getValues()[0];
      sellerVals = sSh.getRange(2, 1, sSh.getLastRow() - 1, sLastCol).getValues();
    }
  } catch (eS) {}
  var cutoffMs = Date.now() - prAutoMakerRewalkDays_() * 24 * 60 * 60 * 1000;
  var unkList = [];
  var missList = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    var parent = String(vals[r][idxMaker] || '').trim();
    if (!parent) continue;
    if (PR_HIT_EXCLUDE_SET_[parent] || prJittaExcludeSeed_(parent)) continue;
    if (idxCk >= 0 && prCkIsTrue_(vals[r][idxCk])) continue;
    var jitta = idxJ >= 0 ? String(vals[r][idxJ] || '').trim() : '';
    if (!prMakerJittaWashEligible_(jitta)) continue;
    var lastMs = prCellToMsTokyo_(idxDay >= 0 ? vals[r][idxDay] : '');
    if (lastMs != null && lastMs > cutoffMs) continue;
    if (sellerVals.length && prMakerFoodZeroOnly_(parent, sellerVals, sellerHdr)) continue;
    var prog = idxProg >= 0 ? String(vals[r][idxProg] || '') : '';
    var st = idxSt >= 0 ? String(vals[r][idxSt] || '').trim() : '';
    if (prWashTitleZeroBlocked_(parent, st, prog)) continue;
    if (prWashHoldBlockedForSelect_(st, prog)) continue;
    if (st === PR_WASH_ST_DONE_ || st === PR_WASH_ST_KEEPA_) continue;
    var estCell = idxEst >= 0 ? vals[r][idxEst] : '';
    var nEst = Number(String(estCell == null ? '' : estCell).replace(/[^\d.-]/g, ''));
    var estOk = !isNaN(nEst) && nEst > 0 && String(estCell).trim() !== '';
    var band =
      idxBand >= 0 && String(vals[r][idxBand] || '').trim()
        ? String(vals[r][idxBand] || '').trim()
        : prMakerCatalogBand_(prMakerCatalogEstimate_(parent, prog, estCell));
    if (!estOk) {
      if (band === PR_BAND_UNK_ || !band) {
        if (unkList.length < want) unkList.push({ parent: parent, detail: 'band_unknown_no_est' });
      } else if (missList.length < want) {
        missList.push({ parent: parent, detail: 'band=' + band + '_no_catalog_est' });
      }
    }
    if (unkList.length >= want) break;
  }
  // 規模不明を優先。足りない分だけ「帯あり見積なし」で埋める
  return unkList.concat(missList).slice(0, want);
}

/** Keepa中＋候補ゼロ＋未済0 を保留へ（片岡型の空ループ救済）。 */
function prSanitizeEmptyResearchStuck_() {
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return { n: 0, detail: 'empty_res_skip no_sum' };
  prEnsureMakerWashStatusCols_(sh);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
  var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
  var idxEst = prHeaderIndex_(hdr, PR_CATALOG_EST_COL_);
  if (idxM < 0 || idxSt < 0) return { n: 0, detail: 'empty_res_skip no_cols' };
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var fixed = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    var st = String(vals[r][idxSt] || '').trim();
    var prog = idxProg >= 0 ? String(vals[r][idxProg] || '') : '';
    if (st !== PR_WASH_ST_KEEPA_ && st !== PR_WASH_ST_ENUM_) continue;
    if (prog.indexOf('done_no_candidates') < 0 && prog.indexOf('empty_research') < 0) continue;
    var parent = String(vals[r][idxM] || '').trim();
    if (!parent) continue;
    if (prWashQueuePendingCount_(parent) > 0) continue;
    var estCell = idxEst >= 0 ? vals[r][idxEst] : null;
    var est = prMakerCatalogEstimate_(parent, prog, estCell);
    var queueN = Object.keys(prWashQueueAsinSet_(parent)).length;
    var holdProg =
      'empty_research est=' +
      (est != null ? est : '?') +
      ' queue=' +
      queueN +
      ' cand=0';
    prSetMakerWashStatus_(parent, PR_WASH_ST_HOLD_, holdProg);
    fixed.push(parent);
  }
  var detail =
    'empty_res_fixed n=' + fixed.length + (fixed.length ? ' makers=' + fixed.join('|').slice(0, 120) : '');
  if (fixed.length) Logger.log('[仕入れ検討①] %s', detail);
  return { n: fixed.length, detail: detail };
}

function prSetMakerWashStatus_(parent, status, progress) {
  var loc = prFindMakerSumRow_(parent);
  if (!loc) {
    Logger.log('[仕入れ検討①] 洗い状態書込スキップ（サマリ行なし）%s', parent);
    return;
  }
  if (status === PR_WASH_ST_DONE_) {
    var cn = prCandGatePassCountForMaker_(parent);
    if (!(cn > 0)) {
      var pendingN = prWashQueuePendingCount_(parent);
      if (pendingN > 0) {
        Logger.log('[仕入れ検討①] WARN done_no_candidates pending parent=%s n=%s', parent, pendingN);
        status = PR_WASH_ST_KEEPA_;
        progress = 'done_no_candidates pending=' + pendingN;
      } else {
        // Catalog見積≥1 でも未済0・候補0 → 完了扱いにせず保留（再armしない）
        var estCell = null;
        try {
          var hdrE = loc.sh.getRange(1, 1, 1, Math.max(loc.sh.getLastColumn(), 1)).getValues()[0];
          var idxE = prHeaderIndex_(hdrE, PR_CATALOG_EST_COL_);
          if (idxE >= 0) estCell = loc.sh.getRange(loc.sheetRow, idxE + 1).getValue();
        } catch (eEst) {}
        var est = prMakerCatalogEstimate_(parent, progress, estCell);
        var queueN = Object.keys(prWashQueueAsinSet_(parent)).length;
        Logger.log(
          '[仕入れ検討①] WARN empty_research parent=%s est=%s queue=%s',
          parent,
          est != null ? est : '?',
          queueN
        );
        status = PR_WASH_ST_HOLD_;
        progress =
          'empty_research est=' +
          (est != null ? est : '?') +
          ' queue=' +
          queueN +
          ' cand=0';
      }
    }
  }
  if (loc.idxSt >= 0) loc.sh.getRange(loc.sheetRow, loc.idxSt + 1).setValue(status);
  if (loc.idxProg >= 0 && progress != null) {
    loc.sh.getRange(loc.sheetRow, loc.idxProg + 1).setValue(String(progress).slice(0, 500));
  }
  if (status === PR_WASH_ST_DONE_) {
    prCycleMarkC4NeedsD1_(parent);
  }
  // 規模帯をサマリへ（選定用）
  var prog = String(progress || '');
  var mq = prog.match(/queue=(\d+)/);
  if (mq) {
    var qn = Number(mq[1]) || 0;
    prWriteMakerSumBand_(loc.sh, loc.sheetRow, qn, qn >= PR_CATALOG_SPLIT_THRESHOLD_ + 1);
  }
}

/** メーカーサマリに Catalog見積／規模帯列を確保して書く。 */
function prWriteMakerSumBand_(sh, sheetRow, estimate, cut501) {
  if (!sh || !(sheetRow > 0)) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxEst = prHeaderIndex_(hdr, PR_CATALOG_EST_COL_);
  var idxBand = prHeaderIndex_(hdr, PR_BAND_COL_);
  if (idxEst < 0) {
    sh.getRange(1, lastCol + 1).setValue(PR_CATALOG_EST_COL_);
    idxEst = lastCol;
    lastCol++;
    hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  if (idxBand < 0) {
    lastCol = Math.max(sh.getLastColumn(), 1);
    sh.getRange(1, lastCol + 1).setValue(PR_BAND_COL_);
    idxBand = lastCol;
  }
  var n = Number(estimate);
  if (isNaN(n) || n < 0) return;
  var writeN = cut501 ? Math.max(n, PR_CATALOG_SPLIT_THRESHOLD_ + 1) : n;
  sh.getRange(sheetRow, idxEst + 1).setValue(writeN);
  sh.getRange(sheetRow, idxBand + 1).setValue(prMakerCatalogBand_(writeN));
  prMakerBandMapClear_();
}

function prLoadMakerQ2Words_(parent) {
  var out = [];
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_MASTER_SHEET_);
    if (!sh || sh.getLastRow() < 2) return out;
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxM = prHeaderIndex_(hdr, 'メーカー');
    var idxQ = prHeaderIndex_(hdr, '第2クエリ語');
    if (idxM < 0 || idxQ < 0) return out;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      if (String(vals[r][idxM] || '').trim() !== parent) continue;
      var raw = String(vals[r][idxQ] || '').trim();
      if (!raw) return out;
      var parts = raw.split('|');
      var i;
      for (i = 0; i < parts.length; i++) {
        var w = String(parts[i] || '').trim();
        if (w) out.push(w);
      }
      return out;
    }
  } catch (e) {
    Logger.log('[仕入れ検討①] 第2語読取スキップ %s', String((e && e.message) || e).slice(0, 120));
  }
  return out;
}

/** Property PR_WASH_QUEUE_COUNT_MAP。未設定=ON。false で1社ずつ全読みする旧動作。 */
function prWashQueueCountMapEnabled_() {
  return prCkEnabled_(PR_WASH_QUEUE_COUNT_MAP_PROP_);
}

/**
 * 洗いキューを1回だけ読み、メーカー→ユニークASIN数の表にする（実行内キャッシュ）。
 * 規模見積は「行数」しか要らないので、社ごとの全読みをここに1本化する。
 * @return {!Object<string, number>}
 */
function prWashQueueCountMap_() {
  if (PR_WASH_QUEUE_COUNT_CACHE_) return PR_WASH_QUEUE_COUNT_CACHE_;
  var out = {};
  var t0 = new Date().getTime();
  var sh = prEnsureWashQueueSheet_();
  if (sh.getLastRow() < 2) {
    PR_WASH_QUEUE_COUNT_CACHE_ = out;
    return out;
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxA = prHeaderIndex_(hdr, 'ASIN');
  if (idxM < 0 || idxA < 0) {
    PR_WASH_QUEUE_COUNT_CACHE_ = out;
    return out;
  }
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var seen = {};
  var r;
  for (r = 0; r < vals.length; r++) {
    var m = String(vals[r][idxM] || '').trim();
    var a = String(vals[r][idxA] || '').trim().toUpperCase();
    if (!m || !a) continue;
    var key = m + '\t' + a;
    if (seen[key]) continue;
    seen[key] = true;
    out[m] = (out[m] || 0) + 1;
  }
  PR_WASH_QUEUE_COUNT_CACHE_ = out;
  Logger.log(
    '[仕入れ検討①] queue_count_map rows=%s makers=%s elapsedMs=%s',
    vals.length,
    Object.keys(out).length,
    new Date().getTime() - t0
  );
  return out;
}

/** 洗いキューの当該メーカーのユニークASIN数。件数だけ要る場所はこちらを使う。 */
function prWashQueueUniqCount_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return 0;
  // 全社ループから呼ばれる。表ができていれば Property も読まない
  if (PR_WASH_QUEUE_COUNT_CACHE_) return Number(PR_WASH_QUEUE_COUNT_CACHE_[parent]) || 0;
  if (!prWashQueueCountMapEnabled_()) {
    return Object.keys(prWashQueueAsinSet_(parent)).length;
  }
  return Number(prWashQueueCountMap_()[parent]) || 0;
}

/** キューへの追記・削除のあとに呼ぶ（件数表を作り直させる）。 */
function prWashQueueCountCacheClear_() {
  PR_WASH_QUEUE_COUNT_CACHE_ = null;
}

/** Property PR_WASH_QUEUE_SNAPSHOT。未設定=ON。false で毎回読み直す旧動作。 */
function prWashQueueSnapshotEnabled_() {
  return prCkEnabled_(PR_WASH_QUEUE_SNAPSHOT_PROP_);
}

/**
 * 当該メーカーのキュー行を1回だけ読む。行順は維持（未済の取り出し順が変わらないように）。
 * @param {string} parent
 * @return {{rows:!Array<{asin:string, sheetRow:number, done:boolean}>, idxK:number}}
 */
function prWashQueueSnapshot_(parent) {
  parent = String(parent || '').trim();
  var reuse = prWashQueueSnapshotEnabled_();
  if (reuse && PR_WASH_QUEUE_SNAP_[parent]) return PR_WASH_QUEUE_SNAP_[parent];
  var snap = { rows: [], idxK: -1 };
  var t0 = new Date().getTime();
  var sh = prEnsureWashQueueSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxA = prHeaderIndex_(hdr, 'ASIN');
  snap.idxK = prHeaderIndex_(hdr, 'Keepa済');
  if (!parent || idxM < 0 || idxA < 0 || sh.getLastRow() < 2) {
    if (reuse) PR_WASH_QUEUE_SNAP_[parent] = snap;
    return snap;
  }
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxM] || '').trim() !== parent) continue;
    var a = String(vals[r][idxA] || '').trim().toUpperCase();
    if (!a) continue;
    snap.rows.push({
      asin: a,
      sheetRow: r + 2,
      done: snap.idxK >= 0 && prCkIsTrue_(vals[r][snap.idxK]),
    });
  }
  if (reuse) PR_WASH_QUEUE_SNAP_[parent] = snap;
  Logger.log(
    '[仕入れ検討①] queue_snap parent=%s scan=%s rows=%s elapsedMs=%s reuse=%s',
    parent,
    vals.length,
    snap.rows.length,
    new Date().getTime() - t0,
    reuse ? 'Y' : 'N'
  );
  return snap;
}

/** 行の追加・削除のあとに呼ぶ（行番号がずれるので作り直させる）。 */
function prWashQueueSnapshotClear_() {
  PR_WASH_QUEUE_SNAP_ = {};
}

function prWashQueueAsinSet_(parent) {
  var set = {};
  var rows = prWashQueueSnapshot_(parent).rows;
  var r;
  for (r = 0; r < rows.length; r++) set[rows[r].asin] = true;
  return set;
}

function prWashQueueAppendAsins_(parent, entry, asins) {
  if (!asins || !asins.length) return 0;
  var sh = prEnsureWashQueueSheet_();
  var snap = prWashQueueSnapshot_(parent);
  var existing = {};
  var r;
  for (r = 0; r < snap.rows.length; r++) existing[snap.rows[r].asin] = true;
  var now = prNowFetchedAt_();
  var lines = [];
  var newAsins = [];
  var i;
  for (i = 0; i < asins.length; i++) {
    var a = String(asins[i] || '').trim().toUpperCase();
    if (!a || existing[a]) continue;
    existing[a] = true;
    newAsins.push(a);
    lines.push([parent, entry, a, now, false, '', '']);
  }
  if (!lines.length) return 0;
  // 末尾に足すだけなので既存行の行番号は動かない。スナップショットは捨てずに継ぎ足す
  // （旧: 1ページ追記ごとにキュー全読みし直していた）
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, lines.length, PR_WASH_QUEUE_HEADERS_.length).setValues(lines);
  for (i = 0; i < newAsins.length; i++) {
    snap.rows.push({ asin: newAsins[i], sheetRow: startRow + i, done: false });
  }
  if (PR_WASH_QUEUE_COUNT_CACHE_) {
    PR_WASH_QUEUE_COUNT_CACHE_[parent] =
      (Number(PR_WASH_QUEUE_COUNT_CACHE_[parent]) || 0) + newAsins.length;
  }
  return lines.length;
}

/**
 * 同一メーカー内の ASIN 重複行を削除。Keepa済の行を優先残し、なければ先に出た行。
 * @return {number} 削除行数
 */
function prWashQueueDedupeAsins_(parent) {
  var snap = prWashQueueSnapshot_(parent);
  var rows = snap.rows;
  if (!rows.length) return 0;
  var keep = {};
  var r;
  for (r = 0; r < rows.length; r++) {
    var a = rows[r].asin;
    if (!keep[a]) {
      keep[a] = rows[r];
      continue;
    }
    if (rows[r].done && !keep[a].done) keep[a] = rows[r];
  }
  var delRows = [];
  for (r = 0; r < rows.length; r++) {
    if (keep[rows[r].asin].sheetRow !== rows[r].sheetRow) delRows.push(rows[r].sheetRow);
  }
  if (!delRows.length) return 0;
  delRows.sort(function (x, y) {
    return y - x;
  });
  var sh = prEnsureWashQueueSheet_();
  var d;
  for (d = 0; d < delRows.length; d++) {
    sh.deleteRow(delRows[d]);
  }
  prWashQueueCountCacheClear_();
  prWashQueueSnapshotClear_();
  Logger.log('[仕入れ検討①] 洗いキュー重複削除 parent=%s n=%s', parent, delRows.length);
  return delRows.length;
}

function prWashQueuePendingKeepa_(parent, cap) {
  var rows = prWashQueueSnapshot_(parent).rows;
  var out = [];
  var seen = {};
  var r;
  for (r = 0; r < rows.length && out.length < cap; r++) {
    if (rows[r].done) continue;
    if (seen[rows[r].asin]) continue;
    seen[rows[r].asin] = true;
    out.push({ asin: rows[r].asin, sheetRow: rows[r].sheetRow });
  }
  return out;
}

/**
 * 当該メーカーの同一 ASIN 行をすべて Keepa済にする（重複行があっても二重GETしない）。
 * 行が続いているところはまとめ書き（旧: 1行ずつ setValue で件数ぶん往復していた）。
 */
function prWashQueueMarkKeepaDoneByAsins_(parent, asins) {
  if (!asins || !asins.length) return;
  var want = {};
  var i;
  for (i = 0; i < asins.length; i++) {
    var a0 = String(asins[i] || '').trim().toUpperCase();
    if (a0) want[a0] = true;
  }
  var snap = prWashQueueSnapshot_(parent);
  if (snap.idxK < 0) return;
  var hit = [];
  var r;
  for (r = 0; r < snap.rows.length; r++) {
    if (!want[snap.rows[r].asin] || snap.rows[r].done) continue;
    snap.rows[r].done = true;
    hit.push(snap.rows[r].sheetRow);
  }
  prWashQueueWriteKeepaDoneRows_(hit, snap.idxK + 1);
}

/** 指定行の Keepa済 列に true を入れる。連続行はまとめて1回で書く。 */
function prWashQueueWriteKeepaDoneRows_(sheetRows, colK) {
  if (!sheetRows || !sheetRows.length || !(colK > 0)) return;
  var sh = prEnsureWashQueueSheet_();
  var rowsSorted = sheetRows.slice().sort(function (x, y) {
    return x - y;
  });
  var i = 0;
  var writes = 0;
  while (i < rowsSorted.length) {
    var j = i;
    while (j + 1 < rowsSorted.length && rowsSorted[j + 1] === rowsSorted[j] + 1) j++;
    var n = j - i + 1;
    var block = [];
    var k;
    for (k = 0; k < n; k++) block.push([true]);
    sh.getRange(rowsSorted[i], colK, n, 1).setValues(block);
    writes++;
    i = j + 1;
  }
  Logger.log('[仕入れ検討①] queue_mark_done rows=%s writes=%s', rowsSorted.length, writes);
}

function prWashQueueMarkKeepaDone_(sheetRows) {
  if (!sheetRows || !sheetRows.length) return;
  var sh = prEnsureWashQueueSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxK = prHeaderIndex_(hdr, 'Keepa済');
  if (idxK < 0) return;
  prWashQueueWriteKeepaDoneRows_(sheetRows, idxK + 1);
  prWashQueueSnapshotClear_();
}

function prWashQueuePendingCount_(parent) {
  return prWashQueuePendingKeepa_(parent, 999999).length;
}

/**
 * 1クエリを nextToken で列挙しキューへ。timeup で途中保存可。
 * titleSink があればタイトルを最大 PR_CAT_INFER_SAMPLE_MAX_ 件蓄積（分割推定用）。
 * @return {{added:number, pages:number, hits:number, nextToken:string, done:boolean, timeup:boolean}}
 */
function prCatalogEnumerateQueryToQueue_(parent, entry, needles, startToken, maxPages, stopAt, titleSink) {
  var token = String(startToken || '');
  var pages = 0;
  var hits = 0;
  var added = 0;
  var maxP = maxPages > 0 ? maxPages : 40;
  var stop = stopAt > 0 ? stopAt : 0;
  var seen = prWashQueueAsinSet_(parent);
  var sink = titleSink || null;
  var sinkMax = PR_CAT_INFER_SAMPLE_MAX_;
  while (pages < maxP) {
    if (stop > 0 && Object.keys(seen).length >= stop) {
      return { added: added, pages: pages, hits: hits, nextToken: token, done: false, timeup: false };
    }
    if (prGasTimeUp_()) {
      return { added: added, pages: pages, hits: hits, nextToken: token, done: false, timeup: true };
    }
    var cat = prCatalogSearchPage_(entry, token);
    pages++;
    var pageHits = prCatalogTitleHits_(cat.items, needles);
    hits += pageHits.length;
    var batch = [];
    var i;
    for (i = 0; i < pageHits.length; i++) {
      var a = String(pageHits[i].asin || '').toUpperCase();
      if (sink && sink.length < sinkMax) {
        var tit = String(pageHits[i].title || '').trim();
        if (tit) sink.push(tit);
      }
      if (!a || seen[a]) continue;
      seen[a] = true;
      batch.push(a);
    }
    added += prWashQueueAppendAsins_(parent, entry, batch);
    token = cat.nextToken || '';
    if (!token) {
      return { added: added, pages: pages, hits: hits, nextToken: '', done: true, timeup: false };
    }
    Utilities.sleep(200);
  }
  return { added: added, pages: pages, hits: hits, nextToken: token, done: !token, timeup: false };
}

function prEnsureWashSplitSheet_() {
  var sh = prEnsureSheet_(PR_WASH_SPLIT_SHEET_, PR_WASH_SPLIT_HEADERS_);
  prEnsureHeaderNames_(sh, PR_WASH_SPLIT_HEADERS_);
  return sh;
}

/** 既存シートに足りないヘッダ名を末尾追記（列順の厳密一致は求めない）。 */
function prEnsureHeaderNames_(sh, names) {
  if (!sh || !names || !names.length) return;
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, names.length).setValues([names]);
    return;
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var have = {};
  var i;
  for (i = 0; i < hdr.length; i++) {
    var h = String(hdr[i] || '').trim();
    if (h) have[h] = true;
  }
  for (i = 0; i < names.length; i++) {
    var n = String(names[i] || '').trim();
    if (!n || have[n]) continue;
    lastCol = Math.max(sh.getLastColumn(), 1) + 1;
    sh.getRange(1, lastCol).setValue(n);
    have[n] = true;
  }
}

function prLoadPendingSplitWords_(parent) {
  var sh = prEnsureWashSplitSheet_();
  var out = [];
  if (sh.getLastRow() < 2) return out;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxW = prHeaderIndex_(hdr, '分割語');
  var idxS = prHeaderIndex_(hdr, '状態');
  var idxTok = prHeaderIndex_(hdr, 'nextToken');
  var idxN = prHeaderIndex_(hdr, '列挙ASIN数');
  if (idxM < 0 || idxW < 0) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxM] || '').trim() !== parent) continue;
    var st = idxS >= 0 ? String(vals[r][idxS] || '').trim() : '';
    if (st === PR_SPLIT_ST_DONE_) continue;
    var w = String(vals[r][idxW] || '').trim();
    if (!w) continue;
    var addedSoFar = idxN >= 0 ? Number(vals[r][idxN] || 0) : 0;
    if (!(addedSoFar > 0)) addedSoFar = 0;
    out.push({
      word: w,
      sheetRow: r + 2,
      token: idxTok >= 0 ? String(vals[r][idxTok] || '') : '',
      addedSoFar: addedSoFar,
    });
  }
  return out;
}

function prWriteSplitWords_(parent, words, source, sampleN) {
  var sh = prEnsureWashSplitSheet_();
  prClearSplitWordsForMaker_(parent);
  if (!words || !words.length) return 0;
  var now = prNowFetchedAt_();
  var lastCol = Math.max(sh.getLastColumn(), PR_WASH_SPLIT_HEADERS_.length);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxW = prHeaderIndex_(hdr, '分割語');
  var idxS = prHeaderIndex_(hdr, '状態');
  var idxSrc = prHeaderIndex_(hdr, '推定元');
  var idxSample = prHeaderIndex_(hdr, 'サンプル数');
  var idxN = prHeaderIndex_(hdr, '列挙ASIN数');
  var idxTok = prHeaderIndex_(hdr, 'nextToken');
  var idxU = prHeaderIndex_(hdr, '更新');
  if (idxM < 0 || idxW < 0) throw new Error('①洗い分割語ヘッダ不足');
  var lines = [];
  var i;
  var seen = {};
  for (i = 0; i < words.length && lines.length < PR_CAT_INFER_WORD_MAX_; i++) {
    var w = String(words[i] || '').trim();
    if (!w || seen[w]) continue;
    seen[w] = true;
    var row = [];
    var c;
    for (c = 0; c < lastCol; c++) row.push('');
    row[idxM] = parent;
    row[idxW] = w;
    if (idxS >= 0) row[idxS] = PR_SPLIT_ST_PENDING_;
    if (idxSrc >= 0) row[idxSrc] = source || '';
    if (idxSample >= 0) row[idxSample] = sampleN || '';
    if (idxN >= 0) row[idxN] = '';
    if (idxTok >= 0) row[idxTok] = '';
    if (idxU >= 0) row[idxU] = now;
    lines.push(row);
  }
  if (!lines.length) return 0;
  sh.getRange(sh.getLastRow() + 1, 1, lines.length, lastCol).setValues(lines);
  return lines.length;
}

function prSaveSplitWordResume_(sheetRow, nextToken, addedCum) {
  var sh = prEnsureWashSplitSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxS = prHeaderIndex_(hdr, '状態');
  var idxTok = prHeaderIndex_(hdr, 'nextToken');
  var idxN = prHeaderIndex_(hdr, '列挙ASIN数');
  var idxU = prHeaderIndex_(hdr, '更新');
  if (idxS >= 0) sh.getRange(sheetRow, idxS + 1).setValue(PR_SPLIT_ST_ENUM_);
  if (idxTok >= 0) sh.getRange(sheetRow, idxTok + 1).setValue(String(nextToken || ''));
  if (idxN >= 0 && addedCum != null) sh.getRange(sheetRow, idxN + 1).setValue(addedCum);
  if (idxU >= 0) {
    sh.getRange(sheetRow, idxU + 1).setValue(
      prNowFetchedAt_()
    );
  }
}

function prMarkSplitWordDone_(sheetRow, addedN) {
  var sh = prEnsureWashSplitSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxS = prHeaderIndex_(hdr, '状態');
  var idxN = prHeaderIndex_(hdr, '列挙ASIN数');
  var idxTok = prHeaderIndex_(hdr, 'nextToken');
  var idxU = prHeaderIndex_(hdr, '更新');
  if (idxS >= 0) sh.getRange(sheetRow, idxS + 1).setValue(PR_SPLIT_ST_DONE_);
  if (idxTok >= 0) sh.getRange(sheetRow, idxTok + 1).setValue('');
  if (idxN >= 0 && addedN != null) sh.getRange(sheetRow, idxN + 1).setValue(addedN);
  if (idxU >= 0) {
    sh.getRange(sheetRow, idxU + 1).setValue(
      prNowFetchedAt_()
    );
  }
}

function prClearSplitWordsForMaker_(parent) {
  var sh = prEnsureWashSplitSheet_();
  if (sh.getLastRow() < 2) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  if (idxM < 0) return 0;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var del = 0;
  var r;
  for (r = vals.length - 1; r >= 0; r--) {
    if (String(vals[r][idxM] || '').trim() !== parent) continue;
    sh.deleteRow(r + 2);
    del++;
  }
  return del;
}

function prParseCategoryWordList_(text) {
  var raw = String(text || '').trim();
  if (!raw) return [];
  var m = raw.match(/\[[\s\S]*\]/);
  if (m) {
    try {
      var arr = JSON.parse(m[0]);
      if (Object.prototype.toString.call(arr) === '[object Array]') {
        return arr
          .map(function (x) {
            return String(x || '').trim();
          })
          .filter(Boolean);
      }
    } catch (e1) {}
  }
  raw = raw.replace(/^```[\s\S]*?\n/, '').replace(/```$/, '');
  var parts = raw.split(/[\n|｜,、\/／]+/);
  var out = [];
  var seen = {};
  var i;
  for (i = 0; i < parts.length; i++) {
    var w = String(parts[i] || '')
      .replace(/^[\d\.\-\*\s]+/, '')
      .replace(/["「」]/g, '')
      .trim();
    if (!w || w.length > 20 || seen[w]) continue;
    if (/^(json|list|カテゴリ|以下)/i.test(w)) continue;
    seen[w] = true;
    out.push(w);
  }
  return out.slice(0, PR_CAT_INFER_WORD_MAX_);
}

function prBuildCategoryInferPrompt_(parent, titles) {
  var sample = (titles || []).slice(0, PR_CAT_INFER_SAMPLE_MAX_);
  return (
    'あなたは日本の食品ECの商品分類アシスタントです。\n' +
    'メーカー「' +
    parent +
    '」の Amazon 商品名サンプルから、Catalog 再検索用の短いカテゴリ語を最大' +
    PR_CAT_INFER_WORD_MAX_ +
    '個挙げてください。\n' +
    '条件: 日本語の短い名詞句のみ。メーカー名自体は入れない。冷凍・冷蔵・生鮮は入れない。\n' +
    '出力は JSON 配列のみ（例: ["みそ汁","おかゆ","カレー"]）。説明文は不要。\n\n' +
    '商品名サンプル:\n' +
    sample.join('\n')
  );
}

function prGeminiGenerateText_(prompt) {
  var key = String(PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '').trim();
  if (!key) {
    Logger.log('[仕入れ検討①] GEMINI_API_KEY 未設定');
    return null;
  }
  var model =
    String(PropertiesService.getScriptProperties().getProperty('PR_GEMINI_MODEL') || '').trim() ||
    PR_GEMINI_MODEL_DEFAULT_;
  var url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model +
    ':generateContent?key=' +
    encodeURIComponent(key);
  var body = {
    contents: [{ parts: [{ text: String(prompt || '') }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 2048 },
  };
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(body),
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('[仕入れ検討①] Gemini HTTP %s %s', res.getResponseCode(), res.getContentText().slice(0, 200));
    return null;
  }
  var json = JSON.parse(res.getContentText());
  var cand = json.candidates && json.candidates[0];
  var parts = cand && cand.content && cand.content.parts ? cand.content.parts : [];
  var texts = [];
  var pi;
  for (pi = 0; pi < parts.length; pi++) {
    if (parts[pi] && parts[pi].text) texts.push(String(parts[pi].text));
  }
  var text = texts.join('\n').trim();
  return text || null;
}

function prOpenAiGenerateText_(prompt) {
  var key = String(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '').trim();
  if (!key) {
    Logger.log('[仕入れ検討①] OPENAI_API_KEY 未設定');
    return null;
  }
  var model =
    String(PropertiesService.getScriptProperties().getProperty('PR_OPENAI_MODEL') || '').trim() ||
    PR_OPENAI_MODEL_DEFAULT_;
  var res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({
      model: model,
      messages: [{ role: 'user', content: String(prompt || '') }],
      temperature: 0.2,
      max_completion_tokens: 600,
    }),
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('[仕入れ検討①] OpenAI HTTP %s %s', res.getResponseCode(), res.getContentText().slice(0, 200));
    return null;
  }
  var json = JSON.parse(res.getContentText());
  if (json.error) {
    Logger.log('[仕入れ検討①] OpenAI error %s', String(json.error.message || '').slice(0, 160));
    return null;
  }
  var content =
    json.choices && json.choices[0] && json.choices[0].message
      ? String(json.choices[0].message.content || '').trim()
      : '';
  return content || null;
}

/**
 * タイトルから分割カテゴリ推定。Gemini 先、失敗時のみ OpenAI。
 * @return {{words:string[], source:string}}
 */
function prInferSplitCategories_(parent, titles) {
  var prompt = prBuildCategoryInferPrompt_(parent, titles || []);
  var gText = prGeminiGenerateText_(prompt);
  var words = prParseCategoryWordList_(gText);
  if (words.length) {
    Logger.log('[仕入れ検討①] 分割推定 gemini n=%s', words.length);
    return { words: words, source: 'gemini' };
  }
  Logger.log('[仕入れ検討①] 分割推定 Gemini 失敗/空 → OpenAI');
  var oText = prOpenAiGenerateText_(prompt);
  words = prParseCategoryWordList_(oText);
  if (words.length) {
    Logger.log('[仕入れ検討①] 分割推定 openai n=%s', words.length);
    return { words: words, source: 'openai' };
  }
  return { words: [], source: '' };
}

function prSplitWordsAllDone_(parent) {
  var sh = prEnsureWashSplitSheet_();
  if (sh.getLastRow() < 2) return false;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxS = prHeaderIndex_(hdr, '状態');
  if (idxM < 0) return false;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var any = false;
  var r;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxM] || '').trim() !== parent) continue;
    any = true;
    var st = idxS >= 0 ? String(vals[r][idxS] || '').trim() : '';
    if (st !== PR_SPLIT_ST_DONE_) return false;
  }
  return any;
}

/** 未／列挙中の分割語をすべて済にする（キューが既にあるときの Keepa 優先用）。 */
function prForceSplitWordsDone_(parent) {
  var sh = prEnsureWashSplitSheet_();
  if (sh.getLastRow() < 2) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = prHeaderIndex_(hdr, 'メーカー');
  var idxS = prHeaderIndex_(hdr, '状態');
  var idxTok = prHeaderIndex_(hdr, 'nextToken');
  var idxU = prHeaderIndex_(hdr, '更新');
  if (idxM < 0 || idxS < 0) return 0;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var n = 0;
  var now = prNowFetchedAt_();
  var r;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxM] || '').trim() !== parent) continue;
    var st = String(vals[r][idxS] || '').trim();
    if (st === PR_SPLIT_ST_DONE_) continue;
    sh.getRange(r + 2, idxS + 1).setValue(PR_SPLIT_ST_DONE_);
    if (idxTok >= 0) sh.getRange(r + 2, idxTok + 1).setValue('');
    if (idxU >= 0) sh.getRange(r + 2, idxU + 1).setValue(now);
    n++;
  }
  return n;
}

/**
 * >500 時: 分割語シートが無ければ推定して書き、未済語を順に列挙。
 * 既に全語「済」／キューが閾値超で分割行が空なら再推定しない。
 * @return {{ok:boolean, status:string, detail:string, added:number}}
 */
function prEnumerateBySplitWords_(parent, needles, titles, runId) {
  prEnsureWashSplitSheet_();
  var pending = prLoadPendingSplitWords_(parent);
  if (!pending.length) {
    if (prSplitWordsAllDone_(parent)) {
      prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'split_all_done');
      return {
        ok: true,
        status: PR_WASH_ST_KEEPA_,
        detail: '分割列挙済（再推定スキップ） parent=' + parent,
        added: 0,
      };
    }
    var queueN0 = Object.keys(prWashQueueAsinSet_(parent)).length;
    if (queueN0 > PR_CATALOG_SPLIT_THRESHOLD_) {
      prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'queue=' + queueN0);
      return {
        ok: true,
        status: PR_WASH_ST_KEEPA_,
        detail: '分割行なし・キュー充分 queue=' + queueN0 + ' 再推定スキップ',
        added: 0,
      };
    }
    prSetMakerWashStatus_(parent, PR_WASH_ST_INFER_, 'infer');
    var infer = prInferSplitCategories_(parent, titles);
    if (!infer.words.length) {
      prSetMakerWashStatus_(parent, PR_WASH_ST_NEED_Q2_, 'infer_fail');
      return {
        ok: true,
        status: PR_WASH_ST_NEED_Q2_,
        detail: '要分割語 parent=' + parent + ' 推定失敗（Gemini/OpenAI）',
        added: 0,
      };
    }
    prWriteSplitWords_(parent, infer.words, infer.source, (titles && titles.length) || 0);
    pending = prLoadPendingSplitWords_(parent);
  }
  var totalAdded = 0;
  var i;
  for (i = 0; i < pending.length; i++) {
    if (prGasTimeUp_()) {
      prSetMakerWashStatus_(parent, PR_WASH_ST_ENUM_, 'split_i=' + i);
      return {
        ok: true,
        status: PR_WASH_ST_ENUM_,
        detail: '列挙中 分割語 i=' + i + '/' + pending.length + ' added=' + totalAdded,
        added: totalAdded,
      };
    }
    var word = pending[i].word;
    var qEntry = parent + ' ' + word;
    var startTok = String(pending[i].token || '');
    Logger.log(
      '[仕入れ検討①] runId=%s split_enum %s resumeTok=%s',
      runId,
      qEntry,
      startTok ? 'Y' : 'N'
    );
    var sh = prEnsureWashSplitSheet_();
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxS = prHeaderIndex_(hdr, '状態');
    if (idxS >= 0) sh.getRange(pending[i].sheetRow, idxS + 1).setValue(PR_SPLIT_ST_ENUM_);
    var er = prCatalogEnumerateQueryToQueue_(parent, qEntry, needles, startTok, 40, 0, null);
    var cum = (pending[i].addedSoFar || 0) + er.added;
    totalAdded += er.added;
    if (er.timeup || !er.done) {
      prSaveSplitWordResume_(pending[i].sheetRow, er.nextToken || '', cum);
      prSetMakerWashStatus_(parent, PR_WASH_ST_ENUM_, 'split=' + word);
      return {
        ok: true,
        status: PR_WASH_ST_ENUM_,
        detail:
          '列挙中 split=' +
          word +
          ' added=' +
          er.added +
          ' cum=' +
          cum +
          ' tok=' +
          (er.nextToken ? 'Y' : 'N'),
        added: totalAdded,
      };
    }
    prMarkSplitWordDone_(pending[i].sheetRow, cum);
  }
  return { ok: true, status: PR_WASH_ST_KEEPA_, detail: '分割列挙完了 added=' + totalAdded, added: totalAdded };
}

/**
 * Catalog 列挙→①洗いキュー。ユニークが閾値超なら第2語分割。Keepaはしない。
 * page0満杯だけでは分割しない（500超が正）。
 * @return {{ok:boolean, skip:boolean, status:string, detail:string}}
 */
function prEnumerateMakerToQueue_(opt) {
  var parent = String((opt && opt.parent) || '').trim();
  var entry = String((opt && opt.entry) || parent).trim();
  var needles = (opt && opt.needles) || [parent];
  var runId = (opt && opt.runId) || '';
  if (!parent) throw new Error('parent 空');
  prEnsureWashQueueSheet_();
  prWashQueueDedupeAsins_(parent);
  // 分割語の途中再開（推定済・未列挙の語がある）
  var pendingSplit = prLoadPendingSplitWords_(parent);
  if (pendingSplit.length) {
    Logger.log('[仕入れ検討①] runId=%s resume_split n=%s', runId, pendingSplit.length);
    var sr = prEnumerateBySplitWords_(parent, needles, [], runId);
    var uniqR = Object.keys(prWashQueueAsinSet_(parent)).length;
    if (sr.status === PR_WASH_ST_KEEPA_) {
      prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'queue=' + uniqR);
    }
    return {
      ok: true,
      skip: false,
      status: sr.status,
      detail: sr.detail + ' queue=' + uniqR,
    };
  }
  prSetMakerWashStatus_(parent, PR_WASH_ST_ENUM_, 'enum_main');
  var totalAdded = 0;
  var totalPages = 0;
  var titles = [];
  // 本クエリ列挙（閾値超で打ち切り）。タイトルは分割推定用に蓄積。
  var erMain = prCatalogEnumerateQueryToQueue_(
    parent,
    entry,
    needles,
    '',
    200,
    PR_CATALOG_SPLIT_THRESHOLD_ + 1,
    titles
  );
  totalAdded += erMain.added;
  totalPages += erMain.pages;
  var uniq = Object.keys(prWashQueueAsinSet_(parent)).length;
  Logger.log(
    '[仕入れ検討①] runId=%s enum main parent=%s uniq=%s pages=%s titles=%s done=%s timeup=%s',
    runId,
    parent,
    uniq,
    totalPages,
    titles.length,
    erMain.done ? 'Y' : 'N',
    erMain.timeup ? 'Y' : 'N'
  );
  if (erMain.timeup) {
    prSetMakerWashStatus_(parent, PR_WASH_ST_ENUM_, 'tok=' + (erMain.nextToken || '').slice(0, 80));
    return {
      ok: true,
      skip: false,
      status: PR_WASH_ST_ENUM_,
      detail: '列挙中 main added=' + totalAdded + ' uniq=' + uniq,
    };
  }
  if (erMain.done && uniq > 0 && uniq <= PR_CATALOG_SPLIT_THRESHOLD_) {
    prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'queue=' + uniq);
    return {
      ok: true,
      skip: false,
      status: PR_WASH_ST_KEEPA_,
      detail: '列挙完了 parent=' + parent + ' queue=' + uniq + ' added=' + totalAdded + ' pages=' + totalPages,
    };
  }
  // 閾値超 → 分割B（タイトル推定 → ①洗い分割語 → 語ごと列挙）
  if (uniq > PR_CATALOG_SPLIT_THRESHOLD_ || (!erMain.done && uniq >= PR_CATALOG_SPLIT_THRESHOLD_)) {
    var sr2 = prEnumerateBySplitWords_(parent, needles, titles, runId);
    totalAdded += sr2.added;
    uniq = Object.keys(prWashQueueAsinSet_(parent)).length;
    if (sr2.status === PR_WASH_ST_KEEPA_) {
      prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'queue=' + uniq);
    }
    return {
      ok: true,
      skip: false,
      status: sr2.status,
      detail: sr2.detail + ' queue=' + uniq + ' pages=' + totalPages,
    };
  }
  if (!erMain.done) {
    prSetMakerWashStatus_(parent, PR_WASH_ST_ENUM_, 'tok=' + (erMain.nextToken || '').slice(0, 80));
    return {
      ok: true,
      skip: false,
      status: PR_WASH_ST_ENUM_,
      detail: '列挙中 main_cap pages=' + totalPages + ' uniq=' + uniq,
    };
  }
  uniq = Object.keys(prWashQueueAsinSet_(parent)).length;
  if (!uniq) {
    throw new Error('タイトル含有0 entry=' + entry);
  }
  prTitleZeroClear_(parent);
  prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'queue=' + uniq);
  return {
    ok: true,
    skip: false,
    status: PR_WASH_ST_KEEPA_,
    detail: '列挙完了 parent=' + parent + ' queue=' + uniq + ' added=' + totalAdded + ' pages=' + totalPages,
  };
}

/**
 * キュー未済 Keepa→門→①候補。完走時のみ最終調査日。
 */
function prWashFullKeepaSlice_(opt) {
  var parent = String((opt && opt.parent) || '').trim();
  var tag = String((opt && opt.tag) || 'wash_full');
  var runId = (opt && opt.runId) || '';
  var cap = prMakerWashCap_();
  if (opt && Number(opt.cap) > 0) {
    cap = Math.min(cap, Math.floor(Number(opt.cap)));
  }
  if (!(cap > 0)) cap = 1;
  prWashQueueDedupeAsins_(parent);
  var pending = prWashQueuePendingKeepa_(parent, cap);
  if (!pending.length) {
    var left0 = prWashQueuePendingCount_(parent);
    if (left0 === 0) {
      prClearSplitWordsForMaker_(parent);
      // Catalog見積≥1 でも未済0・候補0 なら DONE 経由で empty_research 保留になる
      prSetMakerWashStatus_(parent, PR_WASH_ST_DONE_, 'done');
      var loc = prFindMakerSumRow_(parent);
      var stAfter = '';
      if (loc && loc.idxSt >= 0) {
        stAfter = String(loc.sh.getRange(loc.sheetRow, loc.idxSt + 1).getValue() || '').trim();
      }
      if (stAfter === PR_WASH_ST_DONE_ && loc && loc.idxDay >= 0) {
        loc.sh.getRange(loc.sheetRow, loc.idxDay + 1).setValue(prYmdTokyo_());
      }
      prClearMakerSumCk_(parent);
      try {
        prPurgeExpiredWashQueueRows_({ quiet: true, runId: runId });
      } catch (purgeErr0) {
        Logger.log(
          '[仕入れ検討①] runId=%s WARN queue_purge %s',
          runId,
          String((purgeErr0 && purgeErr0.message) || purgeErr0).slice(0, 120)
        );
      }
      if (stAfter === PR_WASH_ST_HOLD_) {
        return {
          ok: true,
          detail: 'empty_research parent=' + parent + ' queue_done',
          complete: true,
          wrote: 0,
        };
      }
      return { ok: true, detail: '完走 parent=' + parent + ' queue_done', complete: true, wrote: 0 };
    }
    return { ok: true, detail: 'Keepa待ちなし leftover=? ', complete: false, wrote: 0 };
  }
  var want = [];
  var i;
  for (i = 0; i < pending.length; i++) want.push(pending[i].asin);
  var twoStage = prKeepaTwoStageGateEnabled_();
  var pack = twoStage ? prKeepaWashGateProducts_(want, runId) : prKeepaProductStats90_(want);
  var notReady = pack.notReady || {};
  var byAsin = {};
  for (i = 0; i < pack.products.length; i++) {
    var pa = String(pack.products[i].asin || '').toUpperCase();
    if (pa) byAsin[pa] = pack.products[i];
  }
  var pass = 0;
  var drop = 0;
  var wrote = 0;
  var held = 0;
  var doneAsins = [];
  for (i = 0; i < pending.length; i++) {
    if (notReady[pending[i].asin]) {
      // 門は通ったが offers=20 が取れていない。BuyBox 欠けで確定させず次スライスへ回す
      held++;
      continue;
    }
    var p = byAsin[pending[i].asin];
    if (!p) {
      Logger.log('[仕入れ検討①] runId=%s WARN no_product %s', runId, pending[i].asin);
      continue;
    }
    var row = prGateProductRow_(p, prReadProfile_(), '', runId);
    row[0] = parent;
    row[6] = tag + '_catalog';
    prUpsertCandidate_(row);
    wrote++;
    doneAsins.push(pending[i].asin);
    if (row[8] === '通過') pass++;
    else drop++;
  }
  prWashQueueMarkKeepaDoneByAsins_(parent, doneAsins);
  var left = prWashQueuePendingCount_(parent);
  var complete = left === 0;
  if (complete) {
    prClearSplitWordsForMaker_(parent);
    prSetMakerWashStatus_(parent, PR_WASH_ST_DONE_, 'done');
    var loc2 = prFindMakerSumRow_(parent);
    var stAfter2 = '';
    if (loc2 && loc2.idxSt >= 0) {
      stAfter2 = String(loc2.sh.getRange(loc2.sheetRow, loc2.idxSt + 1).getValue() || '').trim();
    }
    if (stAfter2 === PR_WASH_ST_DONE_ && loc2 && loc2.idxDay >= 0) {
      loc2.sh.getRange(loc2.sheetRow, loc2.idxDay + 1).setValue(prYmdTokyo_());
    }
    prClearMakerSumCk_(parent);
    try {
      prPurgeExpiredWashQueueRows_({ quiet: true, runId: runId });
    } catch (purgeErr) {
      Logger.log(
        '[仕入れ検討①] runId=%s WARN queue_purge %s',
        runId,
        String((purgeErr && purgeErr.message) || purgeErr).slice(0, 120)
      );
    }
  } else {
    prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'left=' + left);
  }
  var rateLimited = !!pack.rateLimited;
  var sliceWait = pack.waitMin != null ? pack.waitMin : 1;
  var gotN = want.length > 0 ? want.length : wrote;
  var costProfile = twoStage
    ? PR_KEEPA_COST_PROFILE_TWO_STAGE_
    : PR_KEEPA_COST_PROFILE_OFFERS20_;
  var unitCost = prKeepaMeasuredUnit_(pack.tokensConsumed, gotN, costProfile);
  var resumeWait = prKeepaWashResumeWaitMin_({
    pendingCount: left,
    tokensLeft: pack.tokensLeft,
    refillRate: pack.refillRate,
    refillIn: pack.refillIn,
    rateLimited: rateLimited,
    tokensConsumed: pack.tokensConsumed,
    asinCount: gotN,
    unitCost: unitCost,
    cap: cap,
    profile: costProfile,
  });
  var detail =
    'parent=' +
    parent +
    ' want=' +
    want.length +
    ' wrote=' +
    wrote +
    ' pass=' +
    pass +
    ' drop=' +
    drop +
    ' left=' +
    left +
    ' full+=' +
    pack.fullPlus +
    ' skip=' +
    pack.skip +
    ' consumed=' +
    (pack.tokensConsumed != null ? pack.tokensConsumed : '') +
    ' tokLeft=' +
    (pack.tokensLeft != null ? pack.tokensLeft : '') +
    ' unit=' +
    unitCost +
    ' sliceWait=' +
    sliceWait +
    ' waitMin=' +
    resumeWait +
    (twoStage ? ' two_stage=' + (pack.stage1 || 0) + '/' + (pack.stage2 || 0) : '') +
    (held ? ' held=' + held : '') +
    (rateLimited ? ' RATE_LIMIT' : '') +
    (complete ? ' COMPLETE' : '');
  return {
    ok: true,
    detail: detail,
    complete: complete,
    wrote: wrote,
    rateLimited: rateLimited,
    waitMin: resumeWait,
    sliceWaitMin: sliceWait,
    refillIn: pack.refillIn,
    tokensLeft: pack.tokensLeft,
    tokensConsumed: pack.tokensConsumed,
    unitCost: unitCost,
    asinCount: gotN,
  };
}

function prWashAbortEpoch_() {
  return Number(PropertiesService.getScriptProperties().getProperty(PR_WASH_ABORT_EPOCH_PROP_) || 0) || 0;
}

function prWashLockEpoch_() {
  return Number(PropertiesService.getScriptProperties().getProperty(PR_WASH_LOCK_EPOCH_PROP_) || 0) || 0;
}

/** A-1 等。旧洗いループは次スライスで打ち切り。 */
function prWashAbortBump_(via) {
  var props = PropertiesService.getScriptProperties();
  var n = prWashAbortEpoch_() + 1;
  props.setProperty(PR_WASH_ABORT_EPOCH_PROP_, String(n));
  Logger.log('[仕入れ検討①] wash_abort_bump n=%s via=%s', n, String(via || ''));
  return n;
}

/**
 * 孤児／A-1 後の洗いロックを破棄。生きている同一 inflight のロックは残す。
 * @return {{cleared:boolean, detail:string}}
 */
function prWashFullLockReconcile_(via) {
  via = String(via || '');
  if (prWashLockLeftSec_() <= 0) {
    return { cleared: false, detail: 'no_lock' };
  }
  var props = PropertiesService.getScriptProperties();
  var lockOwner = String(props.getProperty(PR_WASH_FULL_LOCK_OWNER_PROP_) || '');
  var lockRun = lockOwner.split(':')[0];
  var lockEpoch = prWashLockEpoch_();
  var abortEpoch = prWashAbortEpoch_();
  var inflight = prWashInflightOwner_() || '';
  var detail = '';
  if (abortEpoch > lockEpoch) {
    prWashFullUnlock_();
    detail = 'cleared abort_epoch abort=' + abortEpoch + ' lockEpoch=' + lockEpoch + ' via=' + via;
    Logger.log('[仕入れ検討①] wash_lock_reconcile %s', detail);
    return { cleared: true, detail: detail };
  }
  if (!inflight || prWashInflightStale_()) {
    prWashFullUnlock_();
    detail =
      'cleared orphan_inflight owner=' +
      lockOwner +
      ' inflight=' +
      (inflight || '-') +
      ' via=' +
      via;
    Logger.log('[仕入れ検討①] wash_lock_reconcile %s', detail);
    return { cleared: true, detail: detail };
  }
  // full_trigger 選定中に旧 Keepa のロックが残っているケース
  if (lockRun && inflight && lockRun !== inflight) {
    prWashFullUnlock_();
    detail =
      'cleared owner_mismatch lock=' + lockRun + ' inflight=' + inflight + ' via=' + via;
    Logger.log('[仕入れ検討①] wash_lock_reconcile %s', detail);
    return { cleared: true, detail: detail };
  }
  return {
    cleared: false,
    detail: 'lock_live owner=' + lockOwner + ' inflight=' + inflight + ' via=' + via,
  };
}

function prWashFullTryLock_(owner) {
  var props = PropertiesService.getScriptProperties();
  var now = new Date().getTime();
  // 取得前に孤児を掃除（A-1 直後の catalog_first を阻まない）
  prWashFullLockReconcile_('try_lock');
  var until = Number(props.getProperty(PR_WASH_FULL_LOCK_PROP_) || 0);
  if (until > now) {
    Logger.log('[仕入れ検討①] wash_full lock busy until=%s owner_try=%s', until, owner || '');
    return false;
  }
  var epoch = prWashAbortEpoch_();
  props.setProperty(PR_WASH_FULL_LOCK_PROP_, String(now + PR_WASH_FULL_LOCK_TTL_MS_));
  props.setProperty(PR_WASH_FULL_LOCK_OWNER_PROP_, String(owner || ''));
  props.setProperty(PR_WASH_LOCK_EPOCH_PROP_, String(epoch));
  return true;
}

function prWashFullUnlock_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PR_WASH_FULL_LOCK_PROP_);
  props.deleteProperty(PR_WASH_FULL_LOCK_OWNER_PROP_);
  // LOCK_EPOCH は残す（次の abort 比較用）。unlock だけでは世代は上げない。
}

/**
 * 洗い完走本線: 列挙（未なら）→ Keepa 1スライス（loopKeepa なら時間上限まで繰り返し）。
 * armTrigger=true なら未完了時に時間トリガーで再開（完走で解除）。
 */
function prRunMakerWashFull_(opt) {
  if (!PR_GAS_START_MS_) prGasStart_();
  var quiet = !!(opt && opt.quiet);
  var enumOnly = !!(opt && opt.enumOnly);
  var loopKeepa = !!(opt && opt.loopKeepa);
  var armTrigger = !!(opt && opt.armTrigger);
  var runId =
    (opt && opt.runId) || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var lockOwner = runId + (armTrigger ? ':arm' : ':run');
  if (!prWashFullTryLock_(lockOwner)) {
    var busyMsg = '洗い完走ロック中（他実行中）。しばらく待つかトリガー停止後に再実行';
    if (!quiet) prMenuFailAlert_(busyMsg);
    prAppendLog_([
      runId,
      ((opt && opt.tag) || 'wash_full') + ' 洗い完走',
      'DONE',
      new Date(),
      '',
      '',
      '',
      'lock_busy ' + prWashLockDiag_() + ' armTrigger=' + (armTrigger ? 'Y' : 'N'),
    ]);
    // トリガー経路でロック負けすると予約が消え停滞する → ロック残り＋1分で再予約
    if (armTrigger) {
      PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      var leftSecBusy = prWashLockLeftSec_();
      // 30分TTL待ちにしない（孤児は reconcile 済み想定。真の競合も短周期で再試行）
      var busyWait = leftSecBusy > 0 ? Math.min(2, Math.ceil(leftSecBusy / 60) + 1) : 1;
      if (busyWait < 1) busyWait = 1;
      var busyAfter = prScheduleWashFullTrigger_(busyWait, {
        runId: runId,
        parent: String((opt && opt.parent) || '').trim(),
        tag: ((opt && opt.tag) || 'wash_full') + '_lock_busy',
      });
      return {
        ok: false,
        skip: true,
        detail:
          'lock_busy trigger_armed afterMin=' +
          busyAfter +
          ' leftSec=' +
          leftSecBusy +
          ' ' +
          prWashLockDiag_(),
      };
    }
    return { ok: false, skip: true, detail: 'lock_busy' };
  }
  try {
    return prRunMakerWashFullBody_(opt);
  } finally {
    prWashFullUnlock_();
  }
}

function prRunMakerWashFullBody_(opt) {
  var quiet = !!(opt && opt.quiet);
  var enumOnly = !!(opt && opt.enumOnly);
  var loopKeepa = !!(opt && opt.loopKeepa);
  var armTrigger = !!(opt && opt.armTrigger);
  var runId =
    (opt && opt.runId) || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var parent = String((opt && opt.parent) || '').trim();
  var entry = String((opt && opt.entry) || parent).trim();
  var needles = (opt && opt.needles) || [parent];
  var tag = String((opt && opt.tag) || 'wash_full');
  var myLockEpoch = prWashLockEpoch_();
  function aborted_() {
    return prWashAbortEpoch_() > myLockEpoch;
  }
  function finishAbort_(where) {
    var hb = 0;
    if (prCycleHeartbeatEnabled_()) {
      hb = prCycleArmHeartbeatAfterIdle_(runId, 'a1_abort');
    }
    return done_(
      true,
      true,
      'a1_abort at=' +
        String(where || '') +
        ' abortEpoch=' +
        prWashAbortEpoch_() +
        ' lockEpoch=' +
        myLockEpoch +
        (hb ? ' heartbeat_armed afterMin=' + hb : ''),
      'DONE'
    );
  }
  function finishIncomplete_(detail, waitMin, refillIn, calcTokensLeft) {
    if (aborted_()) return finishAbort_('incomplete');
    var afterMin = prWashFullTriggerAfterMin_(waitMin, refillIn);
    var armed = false;
    if (armTrigger) {
      // Prop を先に ON（未設定=OFF だと予約が飛ばない）。その後トークン回復見積で予約。
      PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      // P-eff-2: WAIT_TOKEN 遷移前に inflight を手放す（次発火の overlap 連鎖を止める）
      if (prWashEffLoopEnabled_()) {
        prWashInflightClear_('wait_token_arm:' + runId);
      }
      afterMin = prScheduleWashFullTrigger_(afterMin, {
        runId: runId,
        parent: parent,
        tag: tag,
        calcTokensLeft: calcTokensLeft != null ? calcTokensLeft : '',
      });
      armed = true;
    }
    var d =
      String(detail || '') +
      (armed ? ' trigger_armed afterMin=' + afterMin : ' trigger_not_armed');
    return done_(true, false, d, 'DONE', {
      waitMin: waitMin,
      afterMin: afterMin,
      armed: armed,
      refillIn: refillIn,
    });
  }
  function finishComplete_(detail) {
    // 完走時は常に既存予約を消す（armTrigger=false でも PROP だけ OFF→孤児トリガーが trigger_disabled で止まるのを防ぐ）
    prClearWashFullTriggers_({
      reason: 'complete:' + parent,
      runId: runId,
      log: true,
    });
    PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'false');
    prTitleZeroClear_(parent);
    var locDone = prFindMakerSumRow_(parent);
    var stDone = '';
    var progDone = '';
    if (locDone && locDone.idxSt >= 0) {
      stDone = String(locDone.sh.getRange(locDone.sheetRow, locDone.idxSt + 1).getValue() || '').trim();
    }
    if (locDone && locDone.idxProg >= 0) {
      progDone = String(locDone.sh.getRange(locDone.sheetRow, locDone.idxProg + 1).getValue() || '');
    }
    // empty_research 保留は D1 材料なし → C4 せず心拍へ（同社再arm防止）
    if (prWashEmptyResearchBlocked_(stDone, progDone) || stDone === PR_WASH_ST_HOLD_) {
      var hbEmpty = prCycleHeartbeatEnabled_()
        ? prCycleArmHeartbeatAfterIdle_(runId, 'empty_research_heartbeat')
        : 0;
      var emptyDetail =
        String(detail || '') +
        (hbEmpty ? ' heartbeat_armed afterMin=' + hbEmpty : ' empty_research_no_c4');
      return done_(true, false, emptyDetail, 'DONE');
    }
    // 完走 → 店レーン1回を挟む（小中に加え大・特大も。§6.3bis）
    if (stDone === PR_WASH_ST_DONE_) {
      var bandDone = prMakerBandFromSum_(parent);
      if (
        bandDone === PR_BAND_S_ ||
        bandDone === PR_BAND_M_ ||
        bandDone === PR_BAND_L_ ||
        bandDone === PR_BAND_XL_
      ) {
        prCycleMarkShopAfterSm_(parent);
      }
    }
    prCycleMarkC4NeedsD1_(parent);
    var fullDetail = detail;
    if (prCycleC4Enabled_() && parent) {
      var c4After = prScheduleCycleC4Trigger_(PR_CYCLE_C4_TRIGGER_AFTER_DEFAULT_, {
        runId: runId,
        parent: parent,
        tag: tag,
      });
      fullDetail = String(detail || '') + ' c4_trigger_armed afterMin=' + c4After;
    } else if (prCycleHeartbeatEnabled_()) {
      var hbAfter = prCycleArmHeartbeatAfterIdle_(runId, 'complete_no_c4_heartbeat');
      fullDetail = String(detail || '') + ' heartbeat_armed afterMin=' + hbAfter;
    }
    return done_(true, false, fullDetail, 'DONE');
  }
  function done_(ok, skip, detail, state, extra) {
    prAppendLog_([runId, tag + ' 洗い完走', state || 'DONE', new Date(), '', '', '', detail]);
    Logger.log('[仕入れ検討①] runId=%s state=%s %s', runId, state || 'DONE', detail);
    if (!quiet) {
      if (!ok && !skip) prMenuFailAlert_(detail);
      else if (skip) prMenuToast_('① 洗い完走', '洗わない ' + detail);
      else prMenuToast_('① 洗い完走', detail);
    }
    var out = { ok: ok, skip: skip, detail: detail };
    if (extra) {
      if (extra.waitMin != null) out.waitMin = extra.waitMin;
      if (extra.afterMin != null) out.afterMin = extra.afterMin;
      if (extra.armed != null) out.armed = extra.armed;
      if (extra.refillIn != null) out.refillIn = extra.refillIn;
    }
    return out;
  }
  try {
    if (!prMakerWashEnabled_()) throw new Error(PR_MAKER_WASH_ENABLED_PROP_ + '=false');
    if (!parent) throw new Error('parent 空');
    var jitta = prJittaForMaker_(parent);
    if (jitta === PR_JITTAI_EXCLUDE_) {
      prSetMakerWashStatus_(parent, PR_WASH_ST_HOLD_, '非仕入先');
      return done_(true, true, '実態=非仕入先 parent=' + parent);
    }
    if (prMakerJittaWashBlocked_(jitta)) {
      prSetMakerWashStatus_(parent, PR_WASH_ST_HOLD_, jitta);
      return done_(true, true, '実態=' + jitta + ' parent=' + parent);
    }
    if (prWashTitleZeroBlockedForParent_(parent)) {
      return done_(true, true, 'title0_skip blocked parent=' + parent);
    }
    if (aborted_()) return finishAbort_('start');
    var pending = prWashQueuePendingCount_(parent);
    var queueN = prWashQueueUniqCount_(parent);
    var loc = prFindMakerSumRow_(parent);
    var st = '';
    if (loc && loc.idxSt >= 0) {
      st = String(loc.sh.getRange(loc.sheetRow, loc.idxSt + 1).getValue() || '').trim();
    }
    var hasPendingSplit = prLoadPendingSplitWords_(parent).length > 0;
    if (
      !enumOnly &&
      pending > 0 &&
      queueN > PR_CATALOG_SPLIT_THRESHOLD_ &&
      hasPendingSplit
    ) {
      var forced = prForceSplitWordsDone_(parent);
      Logger.log(
        '[仕入れ検討①] runId=%s force_split_done parent=%s n=%s queue=%s (Keepa優先)',
        runId,
        parent,
        forced,
        queueN
      );
      hasPendingSplit = false;
    }
    var needEnum = false;
    if (hasPendingSplit) {
      needEnum = true;
    } else if (queueN === 0) {
      needEnum = true;
    } else if (pending > 0) {
      needEnum = false;
      if (st !== PR_WASH_ST_KEEPA_ && st !== PR_WASH_ST_DONE_) {
        prSetMakerWashStatus_(parent, PR_WASH_ST_KEEPA_, 'queue_pending=' + pending);
      }
    } else if (st === PR_WASH_ST_NEED_Q2_ || st === PR_WASH_ST_INFER_) {
      needEnum = true;
    }
    if (needEnum && st !== PR_WASH_ST_DONE_) {
      var er = prEnumerateMakerToQueue_({
        parent: parent,
        entry: entry,
        needles: needles,
        runId: runId,
      });
      if (
        er.status === PR_WASH_ST_NEED_Q2_ ||
        er.status === PR_WASH_ST_ENUM_ ||
        er.status === PR_WASH_ST_INFER_
      ) {
        if (enumOnly) {
          // 列挙途中でも規模見積があれば帯を残し、心拍へ戻す（沈黙防止）
          prFinalizeEnumOnlySize_(parent);
          if (armTrigger) {
            var midArm = prCycleArmHeartbeatAfterIdle_(runId, 'hb_enum_partial');
            return done_(true, false, er.detail + ' heartbeat_armed afterMin=' + midArm);
          }
          return done_(true, false, er.detail);
        }
        return finishIncomplete_(er.detail, 2, '');
      }
      if (enumOnly) {
        var fin1 = prFinalizeEnumOnlySize_(parent);
        var d1 = er.detail + ' enumOnly band=' + (fin1.band || '') + ' queue=' + fin1.queueN;
        if (armTrigger) {
          var a1 = prCycleArmHeartbeatAfterIdle_(runId, 'hb_after_enum');
          d1 += ' heartbeat_armed afterMin=' + a1;
        }
        return done_(true, false, d1);
      }
    } else if (enumOnly) {
      var fin2 = prFinalizeEnumOnlySize_(parent);
      var d2 =
        '列挙スキップ queue=' +
        queueN +
        ' pendingKeepa=' +
        pending +
        ' status=' +
        st +
        ' band=' +
        (fin2.band || '');
      if (armTrigger) {
        var a2 = prCycleArmHeartbeatAfterIdle_(runId, 'hb_after_enum');
        d2 += ' heartbeat_armed afterMin=' + a2;
      }
      return done_(true, false, d2);
    }
    if (enumOnly) {
      var fin3 = prFinalizeEnumOnlySize_(parent);
      var d3 = 'enumOnly band=' + (fin3.band || '') + ' queue=' + fin3.queueN;
      if (armTrigger) {
        var a3 = prCycleArmHeartbeatAfterIdle_(runId, 'hb_after_enum');
        d3 += ' heartbeat_armed afterMin=' + a3;
      }
      return done_(true, false, d3);
    }
    // Catalog見積が自社・他社で未了なら Keepa に入らない（Keepa中の続きは可）
    var pendingNow = prWashQueuePendingCount_(parent);
    var isKeepaCont = st === PR_WASH_ST_KEEPA_ && pendingNow > 0;
    if (!isKeepaCont) {
      if (!prMakerCatalogEstCellOk_(parent)) {
        var finCat = prFinalizeEnumOnlySize_(parent);
        var dCat =
          'catalog_first_self band=' + (finCat.band || '') + ' queue=' + finCat.queueN;
        if (armTrigger) {
          var aCat = prCycleArmHeartbeatAfterIdle_(runId, 'catalog_first_self');
          dCat += ' heartbeat_armed afterMin=' + aCat;
        }
        return done_(true, false, dCat);
      }
      var otherCat = prCycleFirstCatalogPending_();
      if (otherCat && otherCat.parent && otherCat.parent !== parent) {
        var dOth =
          'catalog_first_other pending=' +
          otherCat.parent +
          ' ' +
          (otherCat.detail || '');
        if (armTrigger) {
          var aOth = prCycleArmHeartbeatAfterIdle_(runId, 'catalog_first_other');
          dOth += ' heartbeat_armed afterMin=' + aOth;
        }
        return done_(true, false, dOth);
      }
    }
    if (!loopKeepa) {
      if (aborted_()) return finishAbort_('pre_keepa');
      var kr1 = prWashFullKeepaSlice_({ parent: parent, tag: tag, runId: runId });
      if (aborted_()) return finishAbort_('post_keepa');
      if (kr1.complete) return finishComplete_(kr1.detail);
      prCycleAfterKeepaSliceMidD1_(kr1, parent, runId, tag);
      // C3+T 等: 1スライス成功でも未完なら必ず再予約（wrote>0 だけだとトリガーが消えたままになる）
      if (armTrigger) {
        return finishIncomplete_(kr1.detail, kr1.waitMin, kr1.refillIn, kr1.tokensLeft);
      }
      return done_(true, false, kr1.detail, 'DONE');
    }

    var slices = 0;
    var totalWrote = 0;
    var lastDetail = '';
    var lastWait = 1;
    var lastRefill = '';
    var lastTokLeft = '';
    var nextCap = prMakerWashCap_();
    var maxSlices = 80;
    // ループ開始時点でも安全ネット予約（1スライス目の途中死対策）
    if (armTrigger) {
      prCheckpointWashFullTrigger_({
        armTrigger: true,
        runId: runId,
        parent: parent,
        tag: tag,
        waitMin: 1,
        refillIn: '',
        sliceN: 0,
      });
    }
    while (!prGasTimeUp_() && slices < maxSlices) {
      if (aborted_()) return finishAbort_('loop');
      var leftBefore = prWashQueuePendingCount_(parent);
      if (leftBefore === 0) {
        var krDone = prWashFullKeepaSlice_({ parent: parent, tag: tag, runId: runId });
        return finishComplete_(
          'LOOP COMPLETE slices=' + slices + ' wrote=' + totalWrote + ' ' + krDone.detail
        );
      }
      var kr = prWashFullKeepaSlice_({
        parent: parent,
        tag: tag,
        runId: runId,
        cap: nextCap,
      });
      slices++;
      totalWrote += kr.wrote || 0;
      lastDetail = kr.detail || '';
      lastWait = kr.waitMin != null ? kr.waitMin : 1;
      lastRefill = kr.refillIn != null ? kr.refillIn : '';
      lastTokLeft = kr.tokensLeft != null ? kr.tokensLeft : '';
      var leftAfter = prWashQueuePendingCount_(parent);
      Logger.log(
        '[仕入れ検討①] runId=%s loop_slice=%s wrote=%s complete=%s rateLimited=%s unit=%s tokLeft=%s left=%s',
        runId,
        slices,
        kr.wrote,
        kr.complete ? 'Y' : 'N',
        kr.rateLimited ? 'Y' : 'N',
        kr.unitCost != null ? kr.unitCost : '',
        kr.tokensLeft != null ? kr.tokensLeft : '',
        leftAfter
      );
      // ①ログ＋チェックポイント再予約（強制終了でも直前予約が残る）
      var ckAfter = 0;
      if (armTrigger && !kr.complete) {
        ckAfter = prCheckpointWashFullTrigger_({
          armTrigger: true,
          runId: runId,
          parent: parent,
          tag: tag,
          waitMin: lastWait,
          refillIn: lastRefill,
          calcTokensLeft: lastTokLeft,
          sliceN: slices,
        });
      }
      prAppendLog_([
        runId,
        tag + ' 洗いスライス',
        'DONE',
        new Date(),
        slices,
        kr.wrote || 0,
        '',
        'slice=' +
          slices +
          ' wrote=' +
          (kr.wrote || 0) +
          ' left=' +
          leftAfter +
          ' complete=' +
          (kr.complete ? 'Y' : 'N') +
          ' rateLimited=' +
          (kr.rateLimited ? 'Y' : 'N') +
          ' tokLeft=' +
          (kr.tokensLeft != null ? kr.tokensLeft : '') +
          (ckAfter > 0 ? ' checkpoint_armed afterMin=' + ckAfter : '') +
          ' now=' +
          prNowFetchedAt_() +
          ' ' +
          String(lastDetail || '').slice(0, 120),
      ]);
      try {
        prEffJudgeAfterSlice_(runId, parent, {
          wrote: kr.wrote || 0,
          left: leftAfter,
          tokLeft: kr.tokensLeft != null ? kr.tokensLeft : '',
          waitMin: lastWait,
          complete: !!kr.complete,
        });
      } catch (eEff0) {}
      if (kr.complete) {
        return finishComplete_(
          'LOOP COMPLETE slices=' + slices + ' wrote=' + totalWrote + ' ' + lastDetail
        );
      }
      prCycleAfterKeepaSliceMidD1_(kr, parent, runId, tag);
      if (kr.rateLimited) {
        return finishIncomplete_(
          'LOOP stop RATE_LIMIT slices=' + slices + ' wrote=' + totalWrote + ' ' + lastDetail,
          lastWait,
          lastRefill,
          lastTokLeft
        );
      }
      if (!(kr.wrote > 0)) {
        return finishIncomplete_(
          'LOOP stop wrote0 slices=' + slices + ' wrote=' + totalWrote + ' ' + lastDetail,
          lastWait,
          lastRefill,
          lastTokLeft
        );
      }
      // 残トークンで次スライスが賄えるなら同実行内で継続（sliceWait≥5でも即終了しない）
      var unit =
        kr.unitCost > 0
          ? Number(kr.unitCost)
          : prKeepaCostProductOffers20_();
      var tokLeft = Number(kr.tokensLeft);
      if (kr.tokensLeft === '' || kr.tokensLeft == null || tokLeft !== tokLeft) tokLeft = 0;
      var afford = unit > 0 ? Math.floor(tokLeft / unit) : 0;
      if (afford < 1) {
        return finishIncomplete_(
          'LOOP stop lowTok slices=' + slices + ' wrote=' + totalWrote + ' ' + lastDetail,
          lastWait,
          lastRefill,
          lastTokLeft
        );
      }
      nextCap = Math.min(prMakerWashCap_(), afford);
      // 予算残りが次スライスに足りなそうならここで終了して再予約（強制終了を避ける）
      if (prGasTimeUp_()) {
        return finishIncomplete_(
          'LOOP timeup slices=' + slices + ' wrote=' + totalWrote + ' ' + lastDetail,
          1,
          lastRefill,
          lastTokLeft
        );
      }
      Utilities.sleep(300);
    }
    return finishIncomplete_(
      'LOOP timeup slices=' + slices + ' wrote=' + totalWrote + ' ' + lastDetail,
      1,
      lastRefill,
      lastTokLeft
    );
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    if (/タイトル含有0/.test(msg)) {
      var tz = prHandleWashTitleZero_({
        parent: parent,
        runId: runId,
        tag: tag,
        armTrigger: armTrigger,
        quiet: quiet,
      });
      if (quiet) return done_(true, tz.skipped, tz.detail, 'DONE');
      if (tz.skipped) return done_(true, true, tz.detail, 'DONE');
      prAppendLog_([
        runId,
        tag + ' 洗い完走',
        'DONE',
        new Date(),
        tz.count,
        parent,
        '',
        tz.detail.slice(0, 200),
      ]);
      return done_(true, false, tz.detail, 'DONE');
    }
    if (armTrigger) {
      PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      prScheduleWashFullTrigger_(prWashFullTriggerAfterMin_(5, ''), {
        runId: runId,
        parent: parent,
        tag: tag,
      });
    }
    if (quiet) return done_(false, false, msg.slice(0, 200), 'FAILED');
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, tag + ' 洗い完走', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

function prMenuWashFullEnumDryOnishi() {
  prGasStart_();
  prRunMakerWashFull_({
    parent: '尾西食品',
    entry: '尾西食品',
    needles: ['尾西食品'],
    tag: 'full_enum_dry',
    enumOnly: true,
  });
}

function prMenuWashFullContinue() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var target = prPickWashContinueTarget_();
  if (!target || !target.parent) {
    prMenuToast_('① 洗い完走 続き', '列挙中／Keepa中のメーカーなし');
    return;
  }
  var needles = prMakerWashNeedles_(target.parent, target.hit || '');
  var entry = prMakerCatalogEntry_(target.parent, target.pref || '', target.hit || '');
  prRunMakerWashFull_({
    parent: target.parent,
    entry: entry,
    needles: needles,
    tag: 'full_continue',
    runId: runId,
  });
}

/** 1実行内で Keepa スライスを繰り返し（約20分上限）。トリガーは作らない。 */
function prMenuWashFullLoopContinue() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var target = prPickWashContinueTarget_();
  if (!target || !target.parent) {
    prMenuToast_('① 洗い完走 ループ', '列挙中／Keepa中のメーカーなし');
    return;
  }
  var needles = prMakerWashNeedles_(target.parent, target.hit || '');
  var entry = prMakerCatalogEntry_(target.parent, target.pref || '', target.hit || '');
  prRunMakerWashFull_({
    parent: target.parent,
    entry: entry,
    needles: needles,
    tag: 'full_loop',
    runId: runId,
    loopKeepa: true,
  });
}

function prWashFullTriggerEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_WASH_FULL_TRIGGER_PROP_);
  if (raw == null || String(raw).trim() === '') return false;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

/** ckpt / overlap_wait では wait_shortened を掛けない（未設定=ON）。 */
function prWashCkptNoShorten_() {
  return prCkEnabled_(PR_WASH_CKPT_NO_SHORTEN_PROP_);
}

/** 効率ループ整流 ON（未設定=ON）。false で旧 ckpt／overlap。 */
function prWashEffLoopEnabled_() {
  return prCkEnabled_(PR_WASH_EFF_LOOP_ENABLED_PROP_);
}

function prWashTriggerTagNoShorten_(tag) {
  var t = String(tag || '');
  if (t.indexOf('_ckpt') >= 0) return true;
  if (t === 'overlap_wait') return true;
  return false;
}

/** ckpt（同一実行の安全ネット）の予約か。満タン2分上限の除外に使う。 */
function prWashTriggerTagIsCkpt_(tag) {
  return String(tag || '').indexOf('_ckpt') >= 0;
}

/** Property PR_WASH_ARM_FULL_CAP。未設定=ON。false で満タン時の2分上限をやめる。 */
function prWashArmFullCapEnabled_() {
  return prCkEnabled_(PR_WASH_ARM_FULL_CAP_PROP_);
}

/** overlap 再予約の afterMin。効率ループONは短再試行2分。OFF時は stale 寄せ（下限8）。 */
function prOverlapWaitAfterMin_() {
  if (prWashEffLoopEnabled_()) return 2;
  var base = 8;
  if (!prCkEnabled_(PR_WASH_OVERLAP_STALE_AWARE_PROP_)) return base;
  var age = prWashInflightAgeMs_();
  if (!(age > 0)) return base;
  var remainMs = Math.max(0, PR_WASH_INFLIGHT_STALE_MS_ - age);
  var fromStale = Math.max(base, Math.ceil(remainMs / 60000) + 1);
  var maxWait = Math.ceil(prKeepaTokenBucketMax_() / prKeepaRefillRate_()) + 1;
  if (fromStale > maxWait) fromStale = maxWait;
  return Math.floor(fromStale);
}

function prCountWashFullTriggers_() {
  var list = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === PR_WASH_FULL_TRIGGER_FN_) n++;
  }
  return n;
}

function prWashLockDiag_() {
  var props = PropertiesService.getScriptProperties();
  var until = Number(props.getProperty(PR_WASH_FULL_LOCK_PROP_) || 0);
  var owner = String(props.getProperty(PR_WASH_FULL_LOCK_OWNER_PROP_) || '');
  var now = new Date().getTime();
  if (!(until > now)) return 'lock=free';
  var leftSec = Math.ceil((until - now) / 1000);
  return 'lock=busy leftSec=' + leftSec + ' owner=' + owner;
}

function prWashLockLeftSec_() {
  var until = Number(PropertiesService.getScriptProperties().getProperty(PR_WASH_FULL_LOCK_PROP_) || 0);
  var now = new Date().getTime();
  if (!(until > now)) return 0;
  return Math.ceil((until - now) / 1000);
}

/**
 * 洗い未完の安全ネット予約。
 * 効率ループON（P-eff-1）: 実行中は張らない。再開は finishIncomplete の waitMin のみ。
 * OFF: 今の実行が終わる見込み＋余裕で afterMin（旧 ckpt0≈21分）。
 * @return {number} 予約した分数。arm しない場合は 0
 */
function prCheckpointWashFullTrigger_(opt) {
  opt = opt || {};
  if (!opt.armTrigger) return 0;
  var tag = String(opt.tag || 'wash') + '_ckpt';
  if (opt.sliceN != null && String(opt.sliceN) !== '') tag += String(opt.sliceN);
  if (prWashEffLoopEnabled_()) {
    Logger.log(
      '[仕入れ検討①] ckpt_skipped_eff_loop tag=%s sliceN=%s runId=%s',
      tag,
      opt.sliceN != null ? opt.sliceN : '',
      opt.runId || ''
    );
    return 0;
  }
  var elapsed = PR_GAS_START_MS_ ? new Date().getTime() - PR_GAS_START_MS_ : 0;
  var remainMs = Math.max(0, PR_GAS_BUDGET_MS_ - elapsed);
  var afterMin = Math.max(2, Math.ceil(remainMs / 60000) + 2);
  var fromTok = prWashFullTriggerAfterMin_(opt.waitMin, opt.refillIn);
  if (fromTok > afterMin) afterMin = fromTok;
  PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
  return prScheduleWashFullTrigger_(afterMin, {
    runId: opt.runId,
    parent: opt.parent || '',
    tag: tag,
    calcTokensLeft: opt.calcTokensLeft != null ? opt.calcTokensLeft : '',
  });
}

function prWashTriggerExpectClear_(reason) {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PR_WASH_TRIGGER_EXPECT_FIRE_MS_PROP_);
  props.deleteProperty(PR_WASH_TRIGGER_ARM_TAG_PROP_);
  props.deleteProperty(PR_WASH_TRIGGER_ARM_PARENT_PROP_);
  props.deleteProperty(PR_WASH_TRIGGER_ARM_RUNID_PROP_);
  props.setProperty(
    PR_WASH_TRIGGER_LAST_CLEAR_PROP_,
    String(reason || '') + '@' + prNowFetchedAt_()
  );
}

function prWashInflightSet_(runId, phase) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PR_WASH_INFLIGHT_RUNID_PROP_, String(runId || ''));
  props.setProperty(PR_WASH_INFLIGHT_MS_PROP_, String(new Date().getTime()));
  props.setProperty(PR_WASH_INFLIGHT_PHASE_PROP_, String(phase || ''));
}

function prWashInflightPhase_(phase) {
  PropertiesService.getScriptProperties().setProperty(
    PR_WASH_INFLIGHT_PHASE_PROP_,
    String(phase || '')
  );
}

function prWashInflightClear_(reason) {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PR_WASH_INFLIGHT_RUNID_PROP_);
  props.deleteProperty(PR_WASH_INFLIGHT_MS_PROP_);
  props.deleteProperty(PR_WASH_INFLIGHT_PHASE_PROP_);
  Logger.log('[仕入れ検討①] inflight_clear %s', String(reason || ''));
}

function prWashInflightAgeMs_() {
  var ms = Number(
    PropertiesService.getScriptProperties().getProperty(PR_WASH_INFLIGHT_MS_PROP_) || 0
  );
  if (!(ms > 0)) return 0;
  return new Date().getTime() - ms;
}

function prWashInflightOwner_() {
  return String(
    PropertiesService.getScriptProperties().getProperty(PR_WASH_INFLIGHT_RUNID_PROP_) || ''
  ).trim();
}

function prWashInflightStale_() {
  var owner = prWashInflightOwner_();
  if (!owner) return false;
  return prWashInflightAgeMs_() >= PR_WASH_INFLIGHT_STALE_MS_;
}

/** 店レーン譲り ON（未設定=ON）。 */
function prShopLaneClearInflightEnabled_() {
  return prCkEnabled_(PR_SHOP_LANE_CLEAR_INFLIGHT_PROP_);
}

function prShopHbBusySet_(runId) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PR_SHOP_HB_BUSY_MS_PROP_, String(new Date().getTime()));
  props.setProperty(PR_SHOP_HB_BUSY_RUN_PROP_, String(runId || ''));
}

function prShopHbBusyClear_(reason) {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PR_SHOP_HB_BUSY_MS_PROP_);
  props.deleteProperty(PR_SHOP_HB_BUSY_RUN_PROP_);
  Logger.log('[仕入れ検討①] shop_hb_busy_clear %s', String(reason || ''));
}

/** 店/心拍作業中（TTL 内）。二重の洗い発火を抑止する。 */
function prShopHbBusy_() {
  if (!prShopLaneClearInflightEnabled_()) return false;
  var ms = Number(
    PropertiesService.getScriptProperties().getProperty(PR_SHOP_HB_BUSY_MS_PROP_) || 0
  );
  if (!(ms > 0)) return false;
  var age = new Date().getTime() - ms;
  if (age >= PR_SHOP_HB_BUSY_TTL_MIN_ * 60 * 1000) {
    prShopHbBusyClear_('ttl_expired');
    return false;
  }
  return true;
}

/**
 * 洗い続き無し／大後回し → 心拍へ譲る。
 * inflight を捨て、短予約＋busy。死亡時も overdue66分を待たない。
 * @return {number} 予約分数（OFF 時は 0）
 */
function prWashYieldToHeartbeat_(runId, tag) {
  if (!prShopLaneClearInflightEnabled_()) return 0;
  var rid = runId || '';
  var t = String(tag || 'hb_yield');
  prWashInflightClear_('yield_hb:' + t + ':' + rid);
  prShopHbBusySet_(rid);
  if (!prWashFullTriggerEnabled_()) return 0;
  var mins = prScheduleWashFullTrigger_(PR_SHOP_YIELD_ARM_MIN_, {
    runId: rid,
    parent: '',
    tag: t,
  });
  prAppendLog_([
    rid || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss'),
    'shop_lane_yield',
    'DONE',
    new Date(),
    mins,
    '',
    '',
    'yield_hb tag=' +
      t +
      ' afterMin=' +
      mins +
      ' busy=1 now=' +
      prNowFetchedAt_(),
  ]);
  return mins;
}

/** 効率合否ログ ON（未設定=ON）。 */
function prEffJudgeEnabled_() {
  return prCkEnabled_(PR_EFF_JUDGE_ENABLED_PROP_);
}

/**
 * fields を key=value 連結（①ログ detail・後解析用）。
 * @param {Object} fields
 * @return {string}
 */
function prEffJudgeDetail_(fields) {
  fields = fields || {};
  var keys = [];
  var k;
  for (k in fields) {
    if (!Object.prototype.hasOwnProperty.call(fields, k)) continue;
    keys.push(k);
  }
  keys.sort();
  var parts = [];
  var i;
  for (i = 0; i < keys.length; i++) {
    k = keys[i];
    var v = fields[k];
    if (v == null || v === '') continue;
    parts.push(k + '=' + String(v).replace(/\s+/g, '_').slice(0, 80));
  }
  return parts.join(' ');
}

/**
 * ①ログへ効率合否1行。state=PASS|WARN|FAIL。対策立案の比較軸。
 * @param {string} runId
 * @param {string} event slice|arm|overlap|stale|fire
 * @param {string} verdict PASS|WARN|FAIL
 * @param {string} cause 短い原因コード
 * @param {Object=} fields
 */
function prAppendEffJudge_(runId, event, verdict, cause, fields) {
  if (!prEffJudgeEnabled_()) return;
  var f = fields || {};
  f.event = event || '';
  f.verdict = verdict || '';
  f.cause = cause || '';
  f.now = prNowFetchedAt_();
  var detail = prEffJudgeDetail_(f);
  var v = String(verdict || 'WARN').toUpperCase();
  if (v !== 'PASS' && v !== 'WARN' && v !== 'FAIL') v = 'WARN';
  try {
    prAppendLog_([
      runId || '',
      '効率合否',
      v,
      new Date(),
      f.afterMin != null ? f.afterMin : f.gapMin != null ? f.gapMin : '',
      f.wrote != null ? f.wrote : f.liveLeft != null ? f.liveLeft : '',
      f.ageMin != null ? f.ageMin : '',
      detail.slice(0, 200),
    ]);
  } catch (e0) {
    Logger.log(
      '[仕入れ検討①] eff_judge_log_fail %s',
      String((e0 && e0.message) || e0).slice(0, 120)
    );
  }
  Logger.log(
    '[仕入れ検討①] eff_judge event=%s verdict=%s cause=%s %s',
    event,
    v,
    cause,
    detail.slice(0, 160)
  );
}

/**
 * Keepa切片後の間隔合否。前回切片からの gapMin を記録し last_slice を更新。
 */
function prEffJudgeAfterSlice_(runId, parent, opt) {
  opt = opt || {};
  if (!prEffJudgeEnabled_()) return;
  parent = String(parent || '').trim();
  var props = PropertiesService.getScriptProperties();
  var nowMs = new Date().getTime();
  var prevMs = Number(props.getProperty(PR_EFF_LAST_SLICE_MS_PROP_) || 0);
  var prevParent = String(props.getProperty(PR_EFF_LAST_SLICE_PARENT_PROP_) || '');
  var gapMin = '';
  var verdict = 'PASS';
  var cause = 'slice_ok';
  if (prevMs > 0) {
    gapMin = Math.round((nowMs - prevMs) / 60000);
    if (gapMin > PR_EFF_GAP_FAIL_MIN_) {
      verdict = 'FAIL';
      cause = 'gap_hole';
    } else if (gapMin > PR_EFF_GAP_PASS_MAX_MIN_) {
      verdict = 'WARN';
      cause = 'gap_slow';
    } else {
      cause = 'gap_ok';
    }
  } else {
    cause = 'slice_first';
  }
  if (!(Number(opt.wrote) > 0) && !opt.complete) {
    verdict = 'FAIL';
    cause = 'wrote0';
  }
  props.setProperty(PR_EFF_LAST_SLICE_MS_PROP_, String(nowMs));
  if (parent) props.setProperty(PR_EFF_LAST_SLICE_PARENT_PROP_, parent);
  prAppendEffJudge_(runId, 'slice', verdict, cause, {
    parent: parent,
    prevParent: prevParent,
    gapMin: gapMin,
    wrote: opt.wrote != null ? opt.wrote : '',
    left: opt.left != null ? opt.left : '',
    tokLeft: opt.tokLeft != null ? opt.tokLeft : '',
    waitMin: opt.waitMin != null ? opt.waitMin : '',
    complete: opt.complete ? 'Y' : 'N',
    idealGap: PR_EFF_GAP_PASS_MAX_MIN_,
    fixHint:
      verdict === 'FAIL' && cause === 'gap_hole'
        ? 'ckpt_overlap_stale'
        : verdict === 'WARN'
          ? 'shorten_wait_or_clear_inflight'
          : '',
  });
}

/**
 * 洗いトリガー予約の合否（満タン長待ち・ckpt）。
 */
function prEffJudgeAfterArm_(runId, tag, afterMin, liveLeft, calcLeft, parent) {
  if (!prEffJudgeEnabled_()) return;
  var mins = Number(afterMin);
  var left = Number(liveLeft);
  var tagS = String(tag || '');
  var isCkpt = tagS.indexOf('_ckpt') >= 0;
  var isOverlap = tagS === 'overlap_wait' || tagS.indexOf('overlap') >= 0;
  var verdict = 'PASS';
  var cause = 'arm_ok';
  var fixHint = '';
  var ckptMax = prEffCkptNormalAfterMin_() + PR_EFF_CKPT_MARGIN_MIN_;
  if (isCkpt && prEffCkptBudgetRuleEnabled_()) {
    // ckpt は残り予算ぶん待つのが正。トークンが満タンでも短縮しない（自分自身と overlap する）。
    // 予算＋余裕を超えてなお満タンのときだけ、待ち計算式の異常として拾う。
    if (left === left && left >= PR_EFF_ARM_FULL_LEFT_ && mins > ckptMax) {
      verdict = 'WARN';
      cause = 'ckpt_long_despite_tokens';
      fixHint = 'check_waitMin_formula';
    } else {
      cause = 'ckpt_safety_net';
    }
  } else if (left === left && left >= PR_EFF_ARM_FULL_LEFT_ && mins >= PR_EFF_ARM_LONG_AFTER_MIN_) {
    verdict = 'FAIL';
    cause = isCkpt ? 'ckpt_full_bucket_long_wait' : 'arm_full_bucket_long_wait';
    fixHint = 'deficit0_afterMin_le_2';
  } else if (isCkpt && mins >= 15) {
    verdict = 'WARN';
    cause = 'ckpt_long_net';
    fixHint = 'ckpt_only_after_slice_or_short';
  } else if (isOverlap && mins >= 12) {
    verdict = 'WARN';
    cause = 'overlap_long_rearm';
    fixHint = 'takeover_if_dead';
  } else if (mins > 30) {
    verdict = 'WARN';
    cause = 'arm_very_long';
    fixHint = 'check_waitMin_formula';
  }
  prAppendEffJudge_(runId, 'arm', verdict, cause, {
    parent: parent || '',
    tag: tagS.slice(0, 40),
    afterMin: mins === mins ? mins : '',
    liveLeft: left === left ? left : liveLeft,
    calcLeft: calcLeft != null ? calcLeft : '',
    isCkpt: isCkpt ? 'Y' : 'N',
    isOverlap: isOverlap ? 'Y' : 'N',
    ckptMax: isCkpt ? ckptMax : '',
    fixHint: fixHint,
  });
}

/** Property PR_EFF_CKPT_BUDGET_RULE。未設定=ON。false で固定15分の旧判定。 */
function prEffCkptBudgetRuleEnabled_() {
  return prCkEnabled_(PR_EFF_CKPT_BUDGET_RULE_PROP_);
}

/** ckpt 予約の正常上限（分）= 実行予算+2。`prCheckpointWashFullTrigger_` の式と対。 */
function prEffCkptNormalAfterMin_() {
  return Math.ceil(PR_GAS_BUDGET_MS_ / 60000) + 2;
}

/** overlap_skip 合否。 */
function prEffJudgeOverlap_(runId, owner, ageMs, phase) {
  if (!prEffJudgeEnabled_()) return;
  var ageMin = Math.round(Number(ageMs) / 60000);
  var verdict = 'WARN';
  var cause = 'overlap_fresh';
  var fixHint = 'short_retry';
  if (ageMin >= PR_EFF_OVERLAP_FAIL_AGE_MIN_) {
    verdict = 'FAIL';
    cause = 'overlap_inflight_stuck';
    fixHint = 'immediate_takeover';
  }
  prAppendEffJudge_(runId, 'overlap', verdict, cause, {
    owner: owner || '',
    ageMin: ageMin,
    ageMs: ageMs,
    phase: phase || '',
    failAgeMin: PR_EFF_OVERLAP_FAIL_AGE_MIN_,
    fixHint: fixHint,
  });
}

/** stale takeover 合否（常に FAIL＝病態が表面化した印）。 */
function prEffJudgeStale_(runId, owner, ageMs, phase) {
  if (!prEffJudgeEnabled_()) return;
  prAppendEffJudge_(runId, 'stale', 'FAIL', 'inflight_stale_takeover', {
    owner: owner || '',
    ageMin: Math.round(Number(ageMs) / 60000),
    ageMs: ageMs,
    phase: phase || '',
    staleThreshMin: Math.round(PR_WASH_INFLIGHT_STALE_MS_ / 60000),
    fixHint: 'clear_inflight_before_arm_shorten_ttl',
  });
}

function prWashLogPhase_(runId, phase, extra) {
  prWashInflightPhase_(phase);
  prAppendLog_([
    runId,
    'full_trigger phase',
    'RUNNING',
    new Date(),
    '',
    '',
    '',
    'phase=' +
      String(phase || '') +
      (extra ? ' ' + extra : '') +
      ' inflight=' +
      prWashInflightOwner_() +
      ' ageMs=' +
      prWashInflightAgeMs_() +
      ' triggerN=' +
      prCountWashFullTriggers_(),
  ]);
}

/**
 * 予約済みなのに予定時刻＋猶予を過ぎて未発火なら①ログへ trigger_missed。
 * @return {string} 書いた detail。問題なしなら ''
 */
function prDetectWashTriggerMissed_(runId, via) {
  var props = PropertiesService.getScriptProperties();
  var inflightOwner = prWashInflightOwner_();
  var inflightAge = prWashInflightAgeMs_();
  if (inflightOwner && inflightAge >= PR_WASH_INFLIGHT_STALE_MS_) {
    var staleDetail =
      'trigger_missed via=' +
      String(via || '') +
      ' hint=inflight_stale owner=' +
      inflightOwner +
      ' ageMs=' +
      inflightAge +
      ' phase=' +
      String(props.getProperty(PR_WASH_INFLIGHT_PHASE_PROP_) || '') +
      ' triggerN=' +
      prCountWashFullTriggers_() +
      ' now=' +
      prNowFetchedAt_();
    prAppendLog_([
      runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss'),
      '洗いトリガー診断',
      'DONE',
      new Date(),
      Math.floor(inflightAge / 60000),
      '',
      '',
      staleDetail,
    ]);
    return staleDetail;
  }
  if (inflightOwner && inflightAge > 0 && inflightAge < PR_WASH_INFLIGHT_STALE_MS_) {
    return '';
  }
  var expectMs = Number(props.getProperty(PR_WASH_TRIGGER_EXPECT_FIRE_MS_PROP_) || 0);
  if (!(expectMs > 0)) return '';
  var graceMs = 3 * 60 * 1000;
  var now = new Date().getTime();
  if (now < expectMs + graceMs) return '';
  var nTrig = prCountWashFullTriggers_();
  var enabled = prWashFullTriggerEnabled_();
  var expectedAt = Utilities.formatDate(
    new Date(expectMs),
    'Asia/Tokyo',
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
  var overdueMin = Math.floor((now - expectMs) / 60000);
  var hint = '';
  if (nTrig === 0 && enabled) hint = 'prop_on_but_no_trigger';
  else if (nTrig === 0 && !enabled) hint = 'no_trigger_prop_off';
  else if (nTrig > 0 && enabled) hint = 'trigger_still_queued_or_stuck';
  else hint = 'orphan_trigger_prop_off';
  var detail =
    'trigger_missed via=' +
    String(via || '') +
    ' expectedAt=' +
    expectedAt +
    ' overdueMin=' +
    overdueMin +
    ' now=' +
    prNowFetchedAt_() +
    ' prop=' +
    (enabled ? 'ON' : 'OFF') +
    ' triggerN=' +
    nTrig +
    ' armTag=' +
    String(props.getProperty(PR_WASH_TRIGGER_ARM_TAG_PROP_) || '') +
    ' armParent=' +
    String(props.getProperty(PR_WASH_TRIGGER_ARM_PARENT_PROP_) || '') +
    ' armRunId=' +
    String(props.getProperty(PR_WASH_TRIGGER_ARM_RUNID_PROP_) || '') +
    ' lastClear=' +
    String(props.getProperty(PR_WASH_TRIGGER_LAST_CLEAR_PROP_) || '') +
    ' ' +
    prWashLockDiag_() +
    ' hint=' +
    hint;
  prAppendLog_([
    runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss'),
    '洗いトリガー診断',
    'DONE',
    new Date(),
    overdueMin,
    '',
    '',
    detail,
  ]);
  Logger.log('[仕入れ検討①] %s', detail);
  return detail;
}

function prWashFullTriggerAfterMin_(waitMinHint, refillIn) {
  var fromProp = Number(
    PropertiesService.getScriptProperties().getProperty(PR_WASH_FULL_TRIGGER_AFTER_PROP_) ||
      PR_WASH_FULL_TRIGGER_AFTER_DEFAULT_
  );
  if (!(fromProp > 0)) fromProp = PR_WASH_FULL_TRIGGER_AFTER_DEFAULT_;
  var fromWait = Number(waitMinHint) > 0 ? Number(waitMinHint) : 0;
  var fromRefill = 0;
  if (Number(refillIn) > 0) fromRefill = Math.ceil(Number(refillIn) / 60000) + 1;
  var rate = prKeepaRefillRate_();
  var maxWait = Math.ceil(prKeepaTokenBucketMax_() / rate) + 1;
  var n;
  if (fromWait > 0) n = Math.max(fromWait, fromRefill, 1);
  else n = Math.max(fromProp, fromRefill, 1);
  if (n > maxWait) n = maxWait;
  return Math.floor(n);
}

/**
 * @param {{runId?:string, reason?:string, log?:boolean}=} opt
 * @return {number} 削除したトリガー本数
 */
function prClearWashFullTriggers_(opt) {
  opt = opt || {};
  var before = prCountWashFullTriggers_();
  var list = ScriptApp.getProjectTriggers();
  var deleted = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === PR_WASH_FULL_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(list[i]);
      deleted++;
    }
  }
  var reason = String(opt.reason || 'clear');
  prWashTriggerExpectClear_(reason);
  if (opt.log) {
    var rid =
      opt.runId ||
      'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    prAppendLog_([
      rid,
      '洗いトリガー解除',
      'DONE',
      new Date(),
      deleted,
      '',
      '',
      'trigger_cleared reason=' +
        reason +
        ' deleted=' +
        deleted +
        ' before=' +
        before +
        ' after=' +
        prCountWashFullTriggers_() +
        ' prop=' +
        (prWashFullTriggerEnabled_() ? 'ON' : 'OFF') +
        ' now=' +
        prNowFetchedAt_(),
    ]);
  }
  return deleted;
}

/** 既存の洗い完走トリガーを消して afterMin 分後に1本予約。上限=バケツMAX回復+1分。
 * @param {number} afterMin
 * @param {{runId?:string, parent?:string, tag?:string}=} logOpt ①ログへ予約記録
 * @return {number} 実際に予約した分数
 */
function prScheduleWashFullTrigger_(afterMin, logOpt) {
  var mins = Number(afterMin);
  if (!(mins > 0)) mins = PR_WASH_FULL_TRIGGER_AFTER_DEFAULT_;
  var maxWait = Math.ceil(prKeepaTokenBucketMax_() / prKeepaRefillRate_()) + 1;
  if (mins > maxWait) mins = maxWait;
  mins = Math.floor(mins);
  var rid =
    (logOpt && logOpt.runId) ||
    'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var tag = (logOpt && logOpt.tag) || 'wash';
  var parent = (logOpt && logOpt.parent) || '';
  // 予約前: /token で実残量を取り、待ち分数と突き合わせ可能にする（消費0）
  var tokProbe = prKeepaTokenStatus_();
  var calcLeft =
    logOpt && logOpt.calcTokensLeft != null && logOpt.calcTokensLeft !== ''
      ? logOpt.calcTokensLeft
      : '';
  // 実残で次スライス可能・待ち計算側の残が欠落／不足なら長待ちを短縮（61分誤待ち対策）
  var unitCost = prKeepaCostProductOffers20_();
  var liveN = Number(tokProbe.tokensLeft);
  var calcN = Number(calcLeft);
  var calcMissing =
    calcLeft === '' || calcLeft == null || !(calcN === calcN);
  var waitShortened = false;
  var shortenAllowed =
    !(prWashCkptNoShorten_() && prWashTriggerTagNoShorten_(tag));
  if (
    shortenAllowed &&
    tokProbe.ok &&
    liveN === liveN &&
    liveN >= unitCost &&
    mins > 2 &&
    (calcMissing || calcN < unitCost)
  ) {
    mins = 2;
    waitShortened = true;
  }
  // P-eff-1: 満タンなのに長待ちは禁止（overlap/誤 wait 含む）。ckpt は同実行の安全ネットなので対象外
  var effCapped = false;
  if (
    (prWashEffLoopEnabled_() || prWashArmFullCapEnabled_()) &&
    !prWashTriggerTagIsCkpt_(tag) &&
    tokProbe.ok &&
    liveN === liveN &&
    liveN >= PR_EFF_ARM_FULL_LEFT_ &&
    mins > 2
  ) {
    mins = 2;
    effCapped = true;
  }
  var probeDetail =
    'token_probe ok=' +
    (tokProbe.ok ? '1' : '0') +
    ' liveLeft=' +
    tokProbe.tokensLeft +
    ' liveRate=' +
    tokProbe.refillRate +
    ' liveRefillIn=' +
    tokProbe.refillIn +
    ' liveConsumed=' +
    tokProbe.tokensConsumed +
    ' calcLeft=' +
    calcLeft +
    ' afterMin=' +
    mins +
    ' maxWait=' +
    maxWait +
    (waitShortened ? ' wait_shortened=1' : '') +
    (effCapped ? ' eff_loop_cap=1' : '') +
    (tokProbe.error ? ' err=' + tokProbe.error : '') +
    ' now=' +
    prNowFetchedAt_();
  prAppendLog_([
    rid,
    tag + ' Keepa token',
    tokProbe.ok ? 'DONE' : 'FAILED',
    new Date(),
    mins,
    parent,
    '',
    probeDetail.slice(0, 200),
  ]);
  Logger.log('[仕入れ検討①] %s', probeDetail);
  prDetectWashTriggerMissed_(rid, 'before_reschedule:' + tag);
  var beforeN = prCountWashFullTriggers_();
  prClearWashFullTriggers_({ reason: 'reschedule:' + tag, runId: rid });
  ScriptApp.newTrigger(PR_WASH_FULL_TRIGGER_FN_).timeBased().after(mins * 60 * 1000).create();
  var afterN = prCountWashFullTriggers_();
  var expectMs = new Date().getTime() + mins * 60 * 1000;
  var expectedAt = Utilities.formatDate(
    new Date(expectMs),
    'Asia/Tokyo',
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PR_WASH_TRIGGER_EXPECT_FIRE_MS_PROP_, String(expectMs));
  props.setProperty(PR_WASH_TRIGGER_ARM_TAG_PROP_, String(tag));
  props.setProperty(PR_WASH_TRIGGER_ARM_PARENT_PROP_, String(parent));
  props.setProperty(PR_WASH_TRIGGER_ARM_RUNID_PROP_, String(rid));
  props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
  Logger.log(
    '[仕入れ検討①] wash_full trigger afterMin=%s max=%s expectedAt=%s triggerN=%s→%s liveLeft=%s',
    mins,
    maxWait,
    expectedAt,
    beforeN,
    afterN,
    tokProbe.tokensLeft
  );
  prAppendLog_([
    rid,
    tag + ' トリガー予約',
    'DONE',
    new Date(),
    mins,
    '',
    '',
    'trigger_armed fn=' +
      PR_WASH_FULL_TRIGGER_FN_ +
      ' afterMin=' +
      mins +
      ' expectedAt=' +
      expectedAt +
      ' parent=' +
      parent +
      ' maxWait=' +
      maxWait +
      ' triggerN=' +
      afterN +
      ' beforeN=' +
      beforeN +
      ' liveLeft=' +
      tokProbe.tokensLeft +
      ' calcLeft=' +
      calcLeft +
      ' prop=ON now=' +
      prNowFetchedAt_() +
      ' ' +
      prWashLockDiag_(),
  ]);
  try {
    prEffJudgeAfterArm_(rid, tag, mins, tokProbe.tokensLeft, calcLeft, parent);
  } catch (eEffArm) {}
  prCycleArmLivenessTrigger_();
  return mins;
}

/** 時間主導: UIなし。完走でトリガー解除。 */
function prRunWashFullFromTrigger() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  if (prCycleHeartbeatPaused_()) {
    prAppendLog_([runId, 'full_trigger 洗い完走', 'DONE', new Date(), '', '', '', 'paused skip']);
    prCycleEnsureLiveness_(runId, 'ft_paused');
    return;
  }
  // 店/心拍作業中（yield 済み）→ 洗い本処理に入らず短再試行（二重 A-light 防止）
  if (prShopHbBusy_()) {
    prAppendLog_([
      runId,
      'full_trigger shop_busy',
      'DONE',
      new Date(),
      PR_SHOP_YIELD_ARM_MIN_,
      '',
      '',
      'shop_hb_busy_skip owner=' +
        String(
          PropertiesService.getScriptProperties().getProperty(PR_SHOP_HB_BUSY_RUN_PROP_) || ''
        ) +
        ' rearmMin=' +
        PR_SHOP_YIELD_ARM_MIN_,
    ]);
    if (prWashFullTriggerEnabled_()) {
      prScheduleWashFullTrigger_(PR_SHOP_YIELD_ARM_MIN_, {
        runId: runId,
        parent: '',
        tag: 'shop_busy_wait',
      });
    }
    prCycleEnsureLiveness_(runId, 'ft_shop_busy');
    return;
  }
  var inflightOwner = prWashInflightOwner_();
  if (inflightOwner && !prWashInflightStale_()) {
    var ovAgeMs = prWashInflightAgeMs_();
    var ovAgeMin = Math.floor(ovAgeMs / 60000);
    var ovPhase = String(
      PropertiesService.getScriptProperties().getProperty(PR_WASH_INFLIGHT_PHASE_PROP_) || ''
    );
    try {
      prEffJudgeOverlap_(runId, inflightOwner, ovAgeMs, ovPhase);
    } catch (eEffOv) {}
    // P-eff-2: age≥8分なら即 takeover（37分 stale まで待たない）
    if (prWashEffLoopEnabled_() && ovAgeMin >= PR_WASH_OVERLAP_TAKEOVER_MIN_) {
      prAppendLog_([
        runId,
        'full_trigger 洗い完走',
        'RUNNING',
        new Date(),
        ovAgeMin,
        '',
        '',
        'inflight_overlap_takeover owner=' +
          inflightOwner +
          ' ageMs=' +
          ovAgeMs +
          ' phase=' +
          ovPhase +
          ' threshMin=' +
          PR_WASH_OVERLAP_TAKEOVER_MIN_,
      ]);
      prWashInflightClear_('overlap_takeover:' + inflightOwner);
      inflightOwner = '';
      // fall through → 本処理
    } else {
      prAppendLog_([
        runId,
        'full_trigger overlap',
        'DONE',
        new Date(),
        ovAgeMin,
        '',
        '',
        'overlap_skip owner=' +
          inflightOwner +
          ' ageMs=' +
          ovAgeMs +
          ' phase=' +
          ovPhase +
          ' triggerN=' +
          prCountWashFullTriggers_(),
      ]);
      if (prWashFullTriggerEnabled_()) {
        prScheduleWashFullTrigger_(prOverlapWaitAfterMin_(), {
          runId: runId,
          parent: '',
          tag: 'overlap_wait',
        });
      }
      prCycleEnsureLiveness_(runId, 'ft_overlap');
      return;
    }
  }
  if (inflightOwner && prWashInflightStale_()) {
    var staleAge = prWashInflightAgeMs_();
    var stalePhase = String(
      PropertiesService.getScriptProperties().getProperty(PR_WASH_INFLIGHT_PHASE_PROP_) || ''
    );
    prDetectWashTriggerMissed_(runId, 'fire_takeover');
    try {
      prEffJudgeStale_(runId, inflightOwner, staleAge, stalePhase);
    } catch (eEffSt) {}
    prWashInflightClear_('stale_takeover:' + inflightOwner);
    prAppendLog_([
      runId,
      'full_trigger 洗い完走',
      'RUNNING',
      new Date(),
      Math.round(staleAge / 60000),
      '',
      '',
      'inflight_stale_takeover owner=' + inflightOwner + ' ageMs=' + staleAge,
    ]);
  }
  prWashInflightSet_(runId, 'fire');
  var props = PropertiesService.getScriptProperties();
  var expectMs = Number(props.getProperty(PR_WASH_TRIGGER_EXPECT_FIRE_MS_PROP_) || 0);
  var lagMin =
    expectMs > 0 ? Math.round((new Date().getTime() - expectMs) / 60000) : '';
  var expectedAt =
    expectMs > 0
      ? Utilities.formatDate(new Date(expectMs), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX")
      : '';
  prAppendLog_([
    runId,
    'full_trigger 洗い完走',
    'RUNNING',
    new Date(),
    lagMin,
    '',
    '',
    'fire now=' +
      prNowFetchedAt_() +
      ' expectedAt=' +
      expectedAt +
      ' lagMin=' +
      lagMin +
      ' prop=' +
      (prWashFullTriggerEnabled_() ? 'ON' : 'OFF') +
      ' triggerN=' +
      prCountWashFullTriggers_() +
      ' armTag=' +
      String(props.getProperty(PR_WASH_TRIGGER_ARM_TAG_PROP_) || '') +
      ' armParent=' +
      String(props.getProperty(PR_WASH_TRIGGER_ARM_PARENT_PROP_) || '') +
      ' armRunId=' +
      String(props.getProperty(PR_WASH_TRIGGER_ARM_RUNID_PROP_) || '') +
      ' inflight=' +
      runId +
      ' ' +
      prWashLockDiag_(),
  ]);
  // inflight 確保後にだけ expect を消す（先消しすると沈黙番犬が盲目）
  prWashTriggerExpectClear_('fired:' + runId);
  try {
    prWashLogPhase_(runId, 'after_clear');
  if (!prWashFullTriggerEnabled_()) {
    // PROP=OFF でも続き対象があれば再有効化（孤児停止からの自動復帰）
    var contRescue = prPickWashContinueTarget_();
    if (contRescue && contRescue.parent) {
      props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      var rescueInfo = prWashContinueDeferInfo_(contRescue.parent);
      prAppendLog_([
        runId,
        'full_trigger 洗い完走',
        'DONE',
        new Date(),
        '',
        contRescue.parent,
        '',
        'trigger_rescue_on parent=' +
          contRescue.parent +
          ' via=' +
          (contRescue.via || '') +
          ' ' +
          rescueInfo.reason +
          (contRescue.deferredDetail ? ' | ' + contRescue.deferredDetail : '') +
          ' now=' +
          prNowFetchedAt_(),
      ]);
    } else if (prAlightHasContinue_() && prCycleHeartbeatEnabled_()) {
      props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      prAppendLog_([
        runId,
        'full_trigger 洗い完走',
        'DONE',
        new Date(),
        '',
        '',
        '',
        'trigger_rescue_alight_cont now=' + prNowFetchedAt_(),
      ]);
      prDispatchHeartbeat_(runId);
      return;
    } else if (contRescue && contRescue.deferredDetail && prCycleHeartbeatEnabled_()) {
      // 大特大だけ後回し中 → トリガー消さず心拍へ（列挙・小中）
      props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
      prAppendLog_([
        runId,
        'full_trigger 洗い完走',
        'DONE',
        new Date(),
        '',
        '',
        '',
        'trigger_rescue_heartbeat ' + contRescue.deferredDetail + ' now=' + prNowFetchedAt_(),
      ]);
      prDispatchHeartbeat_(runId);
      return;
    } else {
      prClearWashFullTriggers_({ reason: 'trigger_disabled', runId: runId, log: true });
      prAppendLog_([
        runId,
        'full_trigger 洗い完走',
        'DONE',
        new Date(),
        '',
        '',
        '',
        'trigger_disabled prop=OFF no_continue now=' + prNowFetchedAt_(),
      ]);
      return;
    }
  }
  if (!prMakerWashEnabled_() || !prMakerWashFullEnabled_()) {
    prClearWashFullTriggers_({ reason: 'wash_disabled', runId: runId, log: true });
    props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'false');
    prAppendLog_([
      runId,
      'full_trigger 洗い完走',
      'DONE',
      new Date(),
      '',
      '',
      '',
      'wash_disabled makerWash=' +
        (prMakerWashEnabled_() ? 'ON' : 'OFF') +
        ' full=' +
        (prMakerWashFullEnabled_() ? 'ON' : 'OFF') +
        ' now=' +
        prNowFetchedAt_(),
    ]);
    return;
  }
  var armMaker = String(props.getProperty(PR_CYCLE_C4_ARM_MAKER_PROP_) || '').trim();
  if (armMaker) {
    props.deleteProperty(PR_CYCLE_C4_ARM_MAKER_PROP_);
    var armBand = prMakerBandFromSum_(armMaker);
    if (armBand === PR_BAND_L_ || armBand === PR_BAND_XL_) {
      if (!prCycleXlWeeklyTryConsume_(armMaker)) {
        var hbCap = prCycleHeartbeatEnabled_()
          ? prCycleArmHeartbeatAfterIdle_(runId, 'C4_arm_xl_week_cap')
          : 0;
        prAppendLog_([
          runId,
          'C4_arm 洗い完走',
          'DONE',
          new Date(),
          '',
          armMaker,
          '',
          'xl_week_cap_block band=' +
            armBand +
            (hbCap ? ' heartbeat_armed afterMin=' + hbCap : ''),
        ]);
        return;
      }
    }
    var armNeedles = prMakerWashNeedles_(armMaker, '');
    var armEntry = prMakerCatalogEntry_(armMaker, '', '');
    try {
      var armOut = prRunMakerWashFull_({
        parent: armMaker,
        entry: armEntry,
        needles: armNeedles,
        tag: 'C4_arm',
        runId: runId,
        loopKeepa: true,
        quiet: true,
        armTrigger: true,
      });
      if (armOut && armOut.skip && String(armOut.detail || '').indexOf('lock_busy') >= 0) {
        prAppendLog_([
          runId,
          'C4_arm 洗い完走',
          'DONE',
          new Date(),
          '',
          '',
          '',
          String(armOut.detail || 'lock_busy') + ' ' + prWashLockDiag_(),
        ]);
        return;
      }
    } catch (eArm) {
      var msgArm = prRedactKey_(String((eArm && eArm.message) || eArm));
      prAppendLog_([runId, 'C4_arm 洗い完走', 'FAILED', new Date(), '', '', '', msgArm.slice(0, 200)]);
      if (prWashFullTriggerEnabled_()) {
        props.setProperty(PR_CYCLE_C4_ARM_MAKER_PROP_, armMaker);
        prScheduleWashFullTrigger_(prWashFullTriggerAfterMin_(5, ''), {
          runId: runId,
          parent: armMaker,
          tag: 'C4_arm_retry',
        });
      }
    }
    return;
  }
  prWashLogPhase_(runId, 'pick_begin');
  var pickT0 = new Date().getTime();
  var target = prPickWashContinueTarget_();
  prWashLogPhase_(
    runId,
    'pick_end',
    'elapsedMs=' +
      (new Date().getTime() - pickT0) +
      ' parent=' +
      ((target && target.parent) || '') +
      ' via=' +
      ((target && target.via) || '')
  );
  if (!target || !target.parent) {
    if (target && target.deferredDetail) {
      prAppendLog_([
        runId,
        'full_trigger 洗い完走',
        'DONE',
        new Date(),
        '',
        '',
        '',
        'no_continue_deferred ' + target.deferredDetail + ' now=' + prNowFetchedAt_(),
      ]);
    }
    // 洗い続き無し（or 大特大を後回し）→ 心拍完成形
    if (prCycleHeartbeatEnabled_()) {
      prWashLogPhase_(runId, 'heartbeat');
      prWashYieldToHeartbeat_(runId, 'hb_yield_deferred');
      prDispatchHeartbeat_(runId);
      return;
    }
    prClearWashFullTriggers_({ reason: 'no_target', runId: runId, log: true });
    props.setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'false');
    prAppendLog_([
      runId,
      'full_trigger 洗い完走',
      'DONE',
      new Date(),
      '',
      '',
      '',
      'no_target heartbeat_off now=' + prNowFetchedAt_() + ' ' + prWashLockDiag_(),
    ]);
    return;
  }
  var ftDefer = prWashContinueDeferInfo_(target.parent);
  prAppendLog_([
    runId,
    'full_trigger 洗い完走',
    'RUNNING',
    new Date(),
    '',
    target.parent,
    '',
    'target parent=' +
      target.parent +
      ' via=' +
      (target.via || '') +
      ' ' +
      ftDefer.reason +
      (target.deferredDetail ? ' | ' + target.deferredDetail : '') +
      ' now=' +
      prNowFetchedAt_(),
  ]);
  // 二重ガード: 選定漏れで大が来ても target ログ後に Keepa しない（通常は deferInfo で除外済）
  if (!ftDefer.defer && prCycleDeferLargeContinueForShop_(target.parent)) {
    prAppendLog_([
      runId,
      'full_trigger shop_before_large',
      'DONE',
      new Date(),
      '',
      '',
      '',
      'skip_large_continue→heartbeat keep_progress band=' +
        prMakerBandFromSum_(target.parent) +
        ' rounds=' +
        prAlightRoundCount_() +
        ' now=' +
        prNowFetchedAt_(),
    ]);
    if (prCycleHeartbeatEnabled_()) {
      prWashLogPhase_(runId, 'heartbeat');
      prWashYieldToHeartbeat_(runId, 'hb_yield_shop_before');
      prDispatchHeartbeat_(runId);
      return;
    }
    var armShop = prCycleArmHeartbeatAfterIdle_(runId, 'ft_shop_before_large');
    prAppendLog_([
      runId,
      'full_trigger shop_before_large',
      'DONE',
      new Date(),
      armShop,
      '',
      '',
      'heartbeat_armed afterMin=' + armShop,
    ]);
    return;
  }
  // 二重ガード: 選定漏れで大特大が来ても Keepa に入らず心拍へ
  if (ftDefer.defer) {
    if (prWashContinueDeferShouldFinalize_(ftDefer.reason)) {
      prFinalizeEnumOnlySize_(target.parent);
    }
    prAppendLog_([
      runId,
      'full_trigger defer_xl',
      'DONE',
      new Date(),
      '',
      target.parent,
      '',
      ftDefer.reason +
        (prWashContinueDeferShouldFinalize_(ftDefer.reason)
          ? ' finalize→heartbeat'
          : ' keep_progress→heartbeat'),
    ]);
    if (prCycleHeartbeatEnabled_()) {
      prWashLogPhase_(runId, 'heartbeat');
      prWashYieldToHeartbeat_(runId, 'hb_yield_defer_xl');
      prDispatchHeartbeat_(runId);
      return;
    }
    var armFt = prCycleArmHeartbeatAfterIdle_(runId, 'ft_defer_xl');
    prAppendLog_([
      runId,
      'full_trigger defer_xl',
      'DONE',
      new Date(),
      armFt,
      target.parent,
      '',
      'heartbeat_armed afterMin=' + armFt,
    ]);
    return;
  }
  var ftBand = prMakerBandFromSum_(target.parent);
  if (ftBand === PR_BAND_L_ || ftBand === PR_BAND_XL_) {
    if (!prCycleXlWeeklyTryConsume_(target.parent)) {
      prAppendLog_([
        runId,
        'full_trigger xl_week_cap',
        'DONE',
        new Date(),
        '',
        target.parent,
        '',
        'xl_week_cap_block continue→heartbeat band=' + ftBand,
      ]);
      if (prCycleHeartbeatEnabled_()) {
        prWashLogPhase_(runId, 'heartbeat');
        prWashYieldToHeartbeat_(runId, 'hb_yield_xl_cap');
        prDispatchHeartbeat_(runId);
        return;
      }
      var armCapFt = prCycleArmHeartbeatAfterIdle_(runId, 'ft_xl_week_cap');
      prAppendLog_([
        runId,
        'full_trigger xl_week_cap',
        'DONE',
        new Date(),
        armCapFt,
        target.parent,
        '',
        'heartbeat_armed afterMin=' + armCapFt,
      ]);
      return;
    }
  }
  var needles = prMakerWashNeedles_(target.parent, target.hit || '');
  var entry = prMakerCatalogEntry_(target.parent, target.pref || '', target.hit || '');
  try {
    var ftOut = prRunMakerWashFull_({
      parent: target.parent,
      entry: entry,
      needles: needles,
      tag: 'full_trigger',
      runId: runId,
      loopKeepa: true,
      quiet: true,
      armTrigger: true,
    });
    if (ftOut && ftOut.skip && String(ftOut.detail || '').indexOf('lock_busy') >= 0) {
      prAppendLog_([
        runId,
        'full_trigger 洗い完走',
        'DONE',
        new Date(),
        '',
        '',
        '',
        String(ftOut.detail || 'lock_busy') + ' ' + prWashLockDiag_(),
      ]);
      return;
    }
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    prAppendLog_([runId, 'full_trigger 洗い完走', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    if (prWashFullTriggerEnabled_()) {
      prScheduleWashFullTrigger_(prWashFullTriggerAfterMin_(5, ''), {
        runId: runId,
        parent: target.parent,
        tag: 'full_trigger',
      });
    }
  }
  } catch (eOuter) {
    var msgOuter = prRedactKey_(String((eOuter && eOuter.message) || eOuter));
    prAppendLog_([
      runId,
      'full_trigger 洗い完走',
      'FAILED',
      new Date(),
      '',
      '',
      '',
      'outer ' + msgOuter.slice(0, 180),
    ]);
    PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
    prScheduleWashFullTrigger_(2, {
      runId: runId,
      parent: '',
      tag: 'watchdog_rearm',
    });
  } finally {
    prWashInflightClear_('exit:' + runId);
    try {
      prCycleEnsureLiveness_(runId, 'ft_finally');
    } catch (eLiv) {}
  }
}

/**
 * 洗い続き対象があれば P3 完走トリガーを予約。
 * @return {{armed:boolean, parent:string, afterMin:number, reason:string}}
 */
function prArmWashFullTriggerForContinue_(opt) {
  opt = opt || {};
  var runId =
    opt.runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  prDetectWashTriggerMissed_(runId, opt.tag || 'arm_continue');
  if (!prMakerWashEnabled_()) {
    return { armed: false, parent: '', afterMin: 0, reason: 'wash_disabled' };
  }
  if (!prMakerWashFullEnabled_()) {
    return { armed: false, parent: '', afterMin: 0, reason: 'wash_full_disabled' };
  }
  var target = prPickWashContinueTarget_();
  if (!target || !target.parent) {
    return { armed: false, parent: '', afterMin: 0, reason: 'no_continue_target' };
  }
  PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'true');
  var afterMin = prScheduleWashFullTrigger_(1, {
    runId: runId,
    parent: target.parent,
    tag: opt.tag || 'wash_trigger',
  });
  return { armed: true, parent: target.parent, afterMin: afterMin, reason: '' };
}

/** トリガー開始: 有効化＋まもなく1回実行予約のみ（長時間Keepaは時間トリガー側）。 */
function prMenuWashFullTriggerStart() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var trig = prArmWashFullTriggerForContinue_({ runId: runId, tag: 'menu_trigger_start' });
  if (!trig.armed) {
    prMenuFailAlert_('対象メーカーなし（キューKeepa未済／洗い状態を確認）');
    return;
  }
  prMenuToast_(
    '① 洗い完走 トリガー',
    '予約しました parent=' + trig.parent + '（約' + trig.afterMin + '分後に自動実行・完走まで再開）'
  );
}

function prMenuWashFullTriggerStop() {
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  prDetectWashTriggerMissed_(runId, 'menu_trigger_stop');
  PropertiesService.getScriptProperties().setProperty(PR_WASH_FULL_TRIGGER_PROP_, 'false');
  prClearWashFullTriggers_({ reason: 'menu_stop', runId: runId, log: true });
  prWashFullUnlock_();
  prMenuToast_('① 洗い完走 トリガー', '停止しました');
}

/** ①ログへ洗いトリガー状態を書く（未発火・PROP・本数の切り分け用）。 */
function prMenuWashTriggerDiagnose() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var props = PropertiesService.getScriptProperties();
  var expectMs = Number(props.getProperty(PR_WASH_TRIGGER_EXPECT_FIRE_MS_PROP_) || 0);
  var expectedAt =
    expectMs > 0
      ? Utilities.formatDate(new Date(expectMs), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX")
      : '';
  var leftMin =
    expectMs > 0 ? Math.round((expectMs - new Date().getTime()) / 60000) : '';
  var detail =
    'trigger_diag via=menu now=' +
    prNowFetchedAt_() +
    ' prop=' +
    (prWashFullTriggerEnabled_() ? 'ON' : 'OFF') +
    ' triggerN=' +
    prCountWashFullTriggers_() +
    ' expectedAt=' +
    expectedAt +
    ' leftMin=' +
    leftMin +
    ' armTag=' +
    String(props.getProperty(PR_WASH_TRIGGER_ARM_TAG_PROP_) || '') +
    ' armParent=' +
    String(props.getProperty(PR_WASH_TRIGGER_ARM_PARENT_PROP_) || '') +
    ' armRunId=' +
    String(props.getProperty(PR_WASH_TRIGGER_ARM_RUNID_PROP_) || '') +
    ' lastClear=' +
    String(props.getProperty(PR_WASH_TRIGGER_LAST_CLEAR_PROP_) || '') +
    ' inflight=' +
    String(props.getProperty(PR_WASH_INFLIGHT_RUNID_PROP_) || '') +
    ' inflightPhase=' +
    String(props.getProperty(PR_WASH_INFLIGHT_PHASE_PROP_) || '') +
    ' inflightAgeMs=' +
    prWashInflightAgeMs_() +
    ' ' +
    prWashLockDiag_() +
    ' alive=' +
    prCycleAliveReason_().reason +
    ' ' +
    prCycleClockSnapshot_();
  prAppendLog_([runId, '洗いトリガー診断', 'DONE', new Date(), leftMin, '', '', detail]);
  var missed = prDetectWashTriggerMissed_(runId, 'menu_diagnose');
  prMenuToast_(
    '① 洗いトリガー診断',
    (missed ? '未発火あり。' : '') + detail.slice(0, 180),
    12
  );
}

/** ①ログへ心拍の時計・番犬状態を書く。 */
function prMenuCycleLivenessDiagnose() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var before = prCycleClockSnapshot_();
  var snap = prCycleAliveReason_();
  var r = prCycleEnsureLiveness_(runId, 'menu_diagnose');
  var detail =
    'liveness_diag via=menu now=' +
    prNowFetchedAt_() +
    ' reason=' +
    snap.reason +
    ' ensure=' +
    r.reason +
    ' rearmed=' +
    (r.rearmed ? '1' : '0') +
    ' before=' +
    before;
  prAppendLog_([runId, 'heartbeat liveness', 'DONE', new Date(), r.rearmed ? 1 : 0, '', '', detail]);
  prMenuToast_('① 心拍生存診断', detail.slice(0, 180), 12);
}

/**
 * 続き対象: ①サマリの洗い状態 → なければ①洗い分割語の未済 → なければキューKeepa未済。
 * 小・中／規模不明が残る間は大／特大の Keepa 続き・キュー続きを後回し。
 * @return {{parent:string, hit?:string, pref?:string, via:string, deferredDetail?:string}|null}
 */
function prPickWashContinueTarget_() {
  var deferred = [];
  var poolOnce = null;
  function poolCached_() {
    if (!poolOnce) poolOnce = prCyclePoolCounts_();
    return poolOnce;
  }
  function noteDefer_(parent, via) {
    var info = prWashContinueDeferInfo_(parent, poolCached_());
    if (!info.defer) return false;
    if (deferred.length < 8) {
      deferred.push(parent + '@' + via + ':' + info.reason);
    }
    return true;
  }
  function deferredDetail_() {
    return deferred.length
      ? 'deferred_skip n=' + deferred.length + ' ' + deferred.join('|').slice(0, 180)
      : '';
  }
  try {
    prSanitizeEmptyResearchStuck_();
  } catch (eEmpty) {
    Logger.log(
      '[仕入れ検討①] WARN empty_res_sanitize %s',
      String((eEmpty && eEmpty.message) || eEmpty).slice(0, 120)
    );
  }
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (sh && sh.getLastRow() >= 2) {
    prEnsureMakerWashStatusCols_(sh);
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxM = prHeaderIndex_(hdr, 'メーカー');
    var idxSt = prHeaderIndex_(hdr, PR_WASH_STATUS_HEADER_);
    var idxProg = prHeaderIndex_(hdr, PR_WASH_PROGRESS_HEADER_);
    var idxHit = prHeaderIndex_(hdr, '入口語ヒット');
    var idxPref = prHeaderIndex_(hdr, '入口推奨');
    if (idxM >= 0 && idxSt >= 0) {
      var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
      var r;
      for (r = 0; r < vals.length; r++) {
        var st = String(vals[r][idxSt] || '').trim();
        if (
          st === PR_WASH_ST_ENUM_ ||
          st === PR_WASH_ST_INFER_ ||
          st === PR_WASH_ST_KEEPA_ ||
          st === PR_WASH_ST_NEED_Q2_
        ) {
          var pSum = String(vals[r][idxM] || '').trim();
          var progSum = idxProg >= 0 ? String(vals[r][idxProg] || '') : '';
          if (pSum && progSum.indexOf('a1_park') >= 0) continue;
          if (pSum && prWashTitleZeroBlocked_(pSum, st, progSum)) continue;
          if (prWashEmptyResearchBlocked_(st, progSum)) continue;
          if (pSum && noteDefer_(pSum, 'sum')) continue;
          return {
            parent: pSum,
            hit: idxHit >= 0 ? String(vals[r][idxHit] || '') : '',
            pref: idxPref >= 0 ? String(vals[r][idxPref] || '') : '',
            via: 'sum',
            deferredDetail: deferredDetail_(),
          };
        }
      }
    }
  }
  var splitSh = prEnsureWashSplitSheet_();
  if (splitSh.getLastRow() >= 2) {
    var lastCol2 = Math.max(splitSh.getLastColumn(), 1);
    var hdr2 = splitSh.getRange(1, 1, 1, lastCol2).getValues()[0];
    var idxM2 = prHeaderIndex_(hdr2, 'メーカー');
    var idxS2 = prHeaderIndex_(hdr2, '状態');
    if (idxM2 >= 0) {
      var vals2 = splitSh.getRange(2, 1, splitSh.getLastRow() - 1, lastCol2).getValues();
      var r2;
      for (r2 = 0; r2 < vals2.length; r2++) {
        var st2 = idxS2 >= 0 ? String(vals2[r2][idxS2] || '').trim() : '';
        if (st2 === PR_SPLIT_ST_DONE_) continue;
        var p2 = String(vals2[r2][idxM2] || '').trim();
        if (!p2) continue;
        if (prWashA1Parked_(p2)) continue;
        if (prWashTitleZeroBlockedForParent_(p2)) continue;
        if (noteDefer_(p2, 'split')) continue;
        return { parent: p2, hit: '', pref: '', via: 'split', deferredDetail: deferredDetail_() };
      }
    }
  }
  var qSh = prEnsureWashQueueSheet_();
  if (qSh.getLastRow() >= 2) {
    var lastCol3 = Math.max(qSh.getLastColumn(), 1);
    var hdr3 = qSh.getRange(1, 1, 1, lastCol3).getValues()[0];
    var idxM3 = prHeaderIndex_(hdr3, 'メーカー');
    var idxK = prHeaderIndex_(hdr3, 'Keepa済');
    if (idxM3 >= 0) {
      var vals3 = qSh.getRange(2, 1, qSh.getLastRow() - 1, lastCol3).getValues();
      var seenQ = {};
      var r3;
      for (r3 = 0; r3 < vals3.length; r3++) {
        if (idxK >= 0 && prCkIsTrue_(vals3[r3][idxK])) continue;
        var p3 = String(vals3[r3][idxM3] || '').trim();
        if (!p3 || seenQ[p3]) continue;
        seenQ[p3] = true;
        if (prWashA1Parked_(p3)) continue;
        if (prWashTitleZeroBlockedForParent_(p3)) continue;
        if (noteDefer_(p3, 'queue')) continue;
        return { parent: p3, hit: '', pref: '', via: 'queue', deferredDetail: deferredDetail_() };
      }
    }
  }
  if (deferred.length) {
    return { parent: '', via: 'deferred_all', deferredDetail: deferredDetail_() };
  }
  return null;
}

function prMenuMakerWashOnishi() {
  prRunMakerWash_({
    parent: '尾西食品',
    entry: '尾西食品',
    needles: ['尾西食品'],
    tag: 'T1_gas',
    forceLegacy: true,
  });
}

function prMenuMakerWashAgf() {
  prRunMakerWash_({
    parent: '味の素',
    entry: 'AGF',
    needles: ['AGF', 'エージーエフ', '味の素'],
    tag: 'T2_gas',
    forceLegacy: true,
  });
}

/**
 * Catalog page0 → タイトル含有 → Keepa stats=90 → Keepaフル追記 → 門 → ①候補。
 * 非仕入先／要確認は洗わない。モールヒット非書。csv[]は倉庫に置かない。
 * @param {{parent:string, entry:string, needles:string[], tag:string, quiet?:boolean, runId?:string}} opt
 * @return {{ok:boolean, skip:boolean, detail:string}}
 */
function prRunMakerWash_(opt) {
  if (!PR_GAS_START_MS_) prGasStart_();
  var quiet = !!(opt && opt.quiet);
  var runId =
    (opt && opt.runId) || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var parent = String((opt && opt.parent) || '').trim();
  var entry = String((opt && opt.entry) || parent).trim();
  var needles = (opt && opt.needles) || [parent];
  var tag = String((opt && opt.tag) || 'maker_wash');
  var cap = prMakerWashCap_();
  Logger.log(
    '[仕入れ検討①] runId=%s step=maker_wash state=RUNNING parent=%s entry=%s cap=%s quiet=%s',
    runId,
    parent,
    entry,
    cap,
    quiet ? 'Y' : 'N'
  );
  function done_(ok, skip, detail, logLabel, state) {
    prAppendLog_([runId, logLabel || tag + ' Catalog→門', state || 'DONE', new Date(), '', '', '', detail]);
    Logger.log('[仕入れ検討①] runId=%s state=%s %s', runId, state || 'DONE', detail);
    if (!quiet) {
      if (!ok && !skip) prMenuFailAlert_(detail);
      else if (skip) prMenuToast_('① メーカー洗い', '洗わない ' + detail);
      else prMenuToast_('① メーカー洗い', detail);
    }
    return { ok: ok, skip: skip, detail: detail };
  }
  try {
    if (!prMakerWashEnabled_()) {
      throw new Error(PR_MAKER_WASH_ENABLED_PROP_ + '=false（切戻し中）');
    }
    if (prMakerWashFullEnabled_() && !(opt && opt.forceLegacy)) {
      return prRunMakerWashFull_(opt);
    }
    if (!parent) throw new Error('parent 空');
    var jitta = prJittaForMaker_(parent);
    if (jitta === PR_JITTAI_EXCLUDE_) {
      return done_(true, true, '実態=非仕入先のため洗わない parent=' + parent, tag + ' 拒否', 'DONE');
    }
    if (prMakerJittaWashBlocked_(jitta)) {
      var stopTag = jitta === PR_JITTAI_UNCHECKED_ ? '未調査停止' : '要確認停止';
      var stopMsg =
        jitta === PR_JITTAI_UNCHECKED_
          ? '実態=未調査のため洗わない parent=' + parent
          : '実態=要確認のため洗わない（メール下書きのみ・自動送信なし） parent=' + parent;
      return done_(true, true, stopMsg, tag + ' ' + stopTag, 'DONE');
    }
    var cat = prCatalogSearchPage0_(entry);
    var hits = prCatalogTitleHits_(cat.items, needles);
    Logger.log(
      '[仕入れ検討①] runId=%s catalog http=%s title_hit=%s next=%s',
      runId,
      cat.code,
      hits.length,
      cat.nextToken ? 'Y' : 'N'
    );
    if (!hits.length) {
      throw new Error('タイトル含有0 entry=' + entry);
    }
    var want = [];
    var i;
    for (i = 0; i < hits.length && want.length < cap; i++) {
      var a = String(hits[i].asin || '').toUpperCase();
      if (a) want.push(a);
    }
    var pack = prKeepaProductStats90_(want);
    var products = pack.products;
    var byAsin = {};
    for (i = 0; i < products.length; i++) {
      var pa = String(products[i].asin || '').toUpperCase();
      if (pa) byAsin[pa] = products[i];
    }
    var pass = 0;
    var drop = 0;
    var wrote = 0;
    for (i = 0; i < want.length; i++) {
      var p = byAsin[want[i]];
      if (!p) {
        Logger.log('[仕入れ検討①] runId=%s WARN no_product %s', runId, want[i]);
        continue;
      }
      var row = prGateProductRow_(p, prReadProfile_(), '', runId);
      row[0] = parent;
      row[6] = tag + '_catalog';
      prUpsertCandidate_(row);
      wrote++;
      if (row[8] === '通過') pass++;
      else drop++;
    }
    var detail =
      'parent=' +
      parent +
      ' entry=' +
      entry +
      ' jitta=' +
      jitta +
      ' title_hit=' +
      hits.length +
      ' want=' +
      want.length +
      ' wrote=' +
      wrote +
      ' pass=' +
      pass +
      ' drop=' +
      drop +
      ' full+=' +
      pack.fullPlus +
      ' need=' +
      pack.need +
      ' skip=' +
      pack.skip +
      ' consumed=' +
      (pack.tokensConsumed != null ? pack.tokensConsumed : '') +
      ' left=' +
      (pack.tokensLeft != null ? pack.tokensLeft : '') +
      ' refillRate=' +
      (pack.refillRate != null ? pack.refillRate : '') +
      ' refillIn=' +
      (pack.refillIn != null ? pack.refillIn : '') +
      ' waitMin=' +
      (pack.waitMin != null ? pack.waitMin : '');
    return done_(true, false, detail, tag + ' Catalog→門', 'DONE');
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    if (quiet) {
      return done_(false, false, msg.slice(0, 200), tag + ' Catalog→門', 'FAILED');
    }
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, tag + ' Catalog→門', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/**
 * メーカーの実態。maker_alias シード優先 → シート実態 → 未登録は未調査（洗わない）。
 * Generic 等はシードで非仕入先。
 */
function prJittaForMaker_(parent) {
  var p = String(parent || '').trim();
  var seed = prJittaSeed_(p);
  if (seed) return seed;
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
    if (!sh || sh.getLastRow() < 2) return PR_JITTAI_UNCHECKED_;
    var lastCol = sh.getLastColumn();
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxName = -1;
    var idxJ = -1;
    var h;
    for (h = 0; h < hdr.length; h++) {
      var name = String(hdr[h] || '').trim();
      if (name === 'メーカー') idxName = h;
      if (name === '実態') idxJ = h;
    }
    if (idxName < 0 || idxJ < 0) return PR_JITTAI_UNCHECKED_;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      if (String(vals[r][idxName] || '').trim() === p) {
        var j = String(vals[r][idxJ] || '').trim();
        if (
          j === PR_JITTAI_MAKER_ ||
          j === PR_JITTAI_BRAND_ ||
          j === PR_JITTAI_CHECK_ ||
          j === PR_JITTAI_UNCHECKED_ ||
          j === PR_JITTAI_EXCLUDE_
        ) {
          return j;
        }
        return PR_JITTAI_UNCHECKED_;
      }
    }
  } catch (e) {
    Logger.log(
      '[仕入れ検討①] メーカーサマリ読取スキップ %s',
      String((e && e.message) || e).slice(0, 160)
    );
  }
  return PR_JITTAI_UNCHECKED_;
}

function prSpapiAccessToken_() {
  var props = PropertiesService.getScriptProperties();
  var clientId = String(props.getProperty('SPAPI_LWA_CLIENT_ID') || '').trim();
  var clientSecret = String(props.getProperty('SPAPI_LWA_CLIENT_SECRET') || '').trim();
  var refreshToken = String(props.getProperty('SPAPI_REFRESH_TOKEN') || '').trim();
  var missing = [];
  if (!clientId) missing.push('SPAPI_LWA_CLIENT_ID');
  if (!clientSecret) missing.push('SPAPI_LWA_CLIENT_SECRET');
  if (!refreshToken) missing.push('SPAPI_REFRESH_TOKEN');
  if (missing.length) {
    throw new Error('Script Properties 不足: ' + missing.join(', ') + '（秘密はチャットに書かない）');
  }
  var resp = UrlFetchApp.fetch(PR_SPAPI_LWA_URL_, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded;charset=UTF-8',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret,
    },
    muteHttpExceptions: true,
  });
  var code = resp.getResponseCode();
  var text = resp.getContentText() || '';
  if (code !== 200) {
    throw new Error('LWA HTTP ' + code + ' ' + text.slice(0, 180));
  }
  var json = JSON.parse(text);
  var token = String(json.access_token || '').trim();
  if (!token) throw new Error('LWA access_token 空');
  return token;
}

function prSpapiEndpoint_() {
  var raw = String(PropertiesService.getScriptProperties().getProperty('SPAPI_ENDPOINT') || '').trim();
  if (!raw) return PR_SPAPI_DEFAULT_ENDPOINT_;
  raw = raw.replace(/^(https?):\\+/i, '$1://');
  raw = raw.replace(/^(https?):\/(?!\/)/i, '$1://');
  raw = raw.replace(/\\/g, '/').replace(/\/+$/, '');
  if (!/^https?:\/\//i.test(raw)) return PR_SPAPI_DEFAULT_ENDPOINT_;
  return raw;
}

function prSpapiMarketplace_() {
  var mp = String(PropertiesService.getScriptProperties().getProperty('SPAPI_MARKETPLACE_ID') || '').trim();
  return mp || PR_SPAPI_DEFAULT_MP_;
}

/** Catalog Items 1ページ（pageSize≤20）。nextToken で続き。生はモールヒットに書かない。 */
function prCatalogSearchPage_(keywords, pageToken) {
  var endpoint = prSpapiEndpoint_();
  var mp = prSpapiMarketplace_();
  var token = prSpapiAccessToken_();
  var url =
    endpoint +
    '/catalog/2022-04-01/items?' +
    'marketplaceIds=' +
    encodeURIComponent(mp) +
    '&keywords=' +
    encodeURIComponent(String(keywords || '')) +
    '&includedData=summaries&pageSize=20';
  var pt = String(pageToken || '').trim();
  if (pt) url += '&pageToken=' + encodeURIComponent(pt);
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      Accept: 'application/json',
      'x-amz-access-token': token,
      'User-Agent': 'PurchaseResearch/1.0 (GAS)',
    },
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '';
  if (code !== 200) {
    var waitMin = prCatalogHttpWaitMin_(resp, code, body);
    if (code === 429 || code === 503 || /QuotaExceeded|Throttl|Rate/i.test(body)) {
      prNotifyCatalogQuotaMail_({
        code: code,
        waitMin: waitMin,
        keywords: keywords,
        body: body.slice(0, 240),
      });
    }
    throw new Error('Catalog HTTP ' + code + ' waitMin≈' + waitMin + ' ' + body.slice(0, 160));
  }
  var data = JSON.parse(body);
  var items = [];
  var list = data.items || [];
  var i;
  for (i = 0; i < list.length; i++) {
    var it = list[i] || {};
    var asin = String(it.asin || '').trim().toUpperCase();
    var title = '';
    var brand = '';
    var sums = it.summaries || [];
    if (sums.length && sums[0]) {
      title = String(sums[0].itemName || '').trim();
      brand = String(sums[0].brand || '').trim();
    }
    if (asin) items.push({ asin: asin, title: title, brand: brand });
  }
  var nextToken = '';
  if (data.pagination && data.pagination.nextToken) {
    nextToken = String(data.pagination.nextToken);
  }
  return { code: code, items: items, nextToken: nextToken };
}

/** Catalog 失敗時の待ち分（Retry-After 秒／本文／既定60分）。 */
function prCatalogHttpWaitMin_(resp, code, body) {
  try {
    var h = resp && resp.getHeaders ? resp.getHeaders() : {};
    var ra = h['Retry-After'] || h['retry-after'] || '';
    var sec = Number(ra);
    if (sec > 0) return Math.max(1, Math.ceil(sec / 60));
  } catch (eH) {}
  var m = String(body || '').match(/(\d+)\s*(?:seconds?|秒)/i);
  if (m) return Math.max(1, Math.ceil(Number(m[1]) / 60));
  if (code === 429 || code === 503) return 60;
  return 30;
}

var PR_CATALOG_QUOTA_MAIL_MS_PROP_ = 'PR_CATALOG_QUOTA_MAIL_LAST_MS';

/** Catalog 上限・スロットル時の社内メール（クールダウン60分）。Keepaへ逃げない。 */
function prNotifyCatalogQuotaMail_(opt) {
  opt = opt || {};
  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty(PR_CATALOG_QUOTA_MAIL_MS_PROP_) || 0);
  if (last && Date.now() - last < 60 * 60 * 1000) {
    Logger.log('[仕入れ検討①] catalog_quota_mail cooldown skip');
    return;
  }
  var waitMin = Number(opt.waitMin) > 0 ? Math.floor(Number(opt.waitMin)) : 60;
  var to = String(
    props.getProperty(PR_CYCLE_DAILY_MAIL_TO_PROP_) || Session.getActiveUser().getEmail() || ''
  ).trim();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var day = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  var subject = '[①] Catalog上限 HTTP' + (opt.code || '') + ' 約' + waitMin + '分待機';
  var body =
    'Catalog Items API が上限／スロットルです。Keepa には進めず列挙を止めました。\n' +
    '時刻: ' +
    day +
    '\n' +
    'HTTP: ' +
    (opt.code || '') +
    '\n' +
    '回復目安: 約' +
    waitMin +
    '分後に再試行してください（Retry-After／既定）。\n' +
    'keywords: ' +
    String(opt.keywords || '').slice(0, 80) +
    '\n' +
    'body: ' +
    String(opt.body || '').slice(0, 200) +
    '\n';
  if (!to) {
    prAppendLog_([runId, 'Catalog上限メール', 'FAILED', new Date(), waitMin, '', '', 'no_to']);
    return;
  }
  try {
    MailApp.sendEmail(to, subject, body);
    props.setProperty(PR_CATALOG_QUOTA_MAIL_MS_PROP_, String(Date.now()));
    prAppendLog_([
      runId,
      'Catalog上限メール',
      'DONE',
      new Date(),
      waitMin,
      '',
      '',
      'sent=' + to + ' code=' + (opt.code || ''),
    ]);
  } catch (e) {
    prAppendLog_([
      runId,
      'Catalog上限メール',
      'FAILED',
      new Date(),
      '',
      '',
      '',
      String((e && e.message) || e).slice(0, 160),
    ]);
  }
}

/** Catalog Items キーワード page0（pageSize≤20）。生はモールヒットに書かない。 */
function prCatalogSearchPage0_(keywords) {
  return prCatalogSearchPage_(keywords, '');
}

function prCatalogTitleHits_(items, needles) {
  var out = [];
  var i;
  for (i = 0; i < (items || []).length; i++) {
    var r = items[i];
    var blob = String(r.title || '') + ' ' + String(r.brand || '');
    var j;
    var ok = false;
    for (j = 0; j < needles.length; j++) {
      if (needles[j] && blob.indexOf(needles[j]) >= 0) {
        ok = true;
        break;
      }
    }
    if (ok) out.push(r);
  }
  return out;
}

function prRefuseOriginalSs_() {
  if (prSpreadsheet_().getId() === PR_ORIG_SS_) {
    throw new Error('原本スプシでは実行しない');
  }
}

function prListAppendEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_LIST_APPEND_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prListAppendCap_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_LIST_APPEND_CAP_PROP_) || 100);
  if (!(n > 0)) n = 100;
  if (n > 500) n = 500;
  return Math.floor(n);
}

function prListSkipHeadersMap_() {
  var list = [
    'ASIN重複チェック',
    '調査員名',
    'メーカーがマイナーか有名か',
    '商品画像',
    '単品商品名',
    'JAN調査プロンプト',
    'AI調査プロンプト',
    'FBA出品者',
    'メーカーが出品者にいる',
    '1点当たり メーカー定価(税込)',
    '価格が安定しているか',
    '卸商品リスト卸値（税抜き）',
    '卸商品リスト卸値（税込み）',
    '卸値一覧（あれば）sheet商品名VLOOK',
    '卸値一覧（あれば）sheet卸値税抜き',
    '卸値一覧（あれば）sheet卸値税込み',
    '最低利益額',
    '1点販売額',
    '1点FBA手数料',
    '目標仕入額(税込み）',
    '目標仕入額(税抜き）',
    '最低目標掛率',
    '定価と比べてama価格は',
    '割引率',
    '1点仕入れ額（税込み）',
    '1点損益分岐額',
    '1点利益額',
    '粗利益率',
    '順位10000単位',
    '順位対象外',
    '調査年月',
    '件数',
    '調査追加時',
    '合計調査件数',
    '報酬額',
    '単品商品名２ＰＶ用',
    '計算',
    'nn',
  ];
  var o = {};
  var i;
  for (i = 0; i < list.length; i++) o[list[i]] = true;
  return o;
}

function prListDirectWriteMap_() {
  var list = [
    '調査日',
    'メーカー名',
    'ASIN',
    'ランキング',
    '過去1か月販売数（表示あれば）',
    '商品名',
    'JANコード',
    'サイズ',
    '重量',
    '出品者数',
    '画像URL',
    'Keepaグラフ',
    'amazonページ',
    '最安値',
    'FBA手数料',
    'amazon本体が出品者にいる',
    '順位',
    'カテゴリー',
    'クイックショップサイズ表記',
    '自己配送 送料概算',
    'サイズ3辺合計',
    '税率',
    '販売形態',
    '同商品のセット数量',
  ];
  var o = {};
  var i;
  for (i = 0; i < list.length; i++) o[list[i]] = true;
  return o;
}

/** D1 差分判定に使う機械列（利益式・人間列は含めない）。 */
function prListDiffHeaderKeys_() {
  return [
    'メーカー名',
    'ランキング',
    '最安値',
    '順位',
    'カテゴリー',
    '商品名',
    'JANコード',
    'サイズ',
    '重量',
    '出品者数',
    '画像URL',
    'FBA手数料',
    'amazon本体が出品者にいる',
    '過去1か月販売数（表示あれば）',
    'クイックショップサイズ表記',
    '自己配送 送料概算',
    'サイズ3辺合計',
    '販売形態',
    '同商品のセット数量',
  ];
}

function prListNormCell_(v) {
  if (v == null || v === '') return '';
  return String(v).trim().replace(/,/g, '');
}

function prListRecDiffers_(headers, existingVals, newRec) {
  var keys = prListDiffHeaderKeys_();
  var i;
  for (i = 0; i < keys.length; i++) {
    var k = keys[i];
    var idx = headers.indexOf(k);
    if (idx < 0) continue;
    if (prListNormCell_(existingVals[idx]) !== prListNormCell_(newRec[k])) return true;
  }
  return false;
}

function prListWriteRecs_(sh, headers, startRow, recs, formulaCols, directMap, skipMap) {
  var n = recs.length;
  var j;
  var i;
  for (j = 0; j < headers.length; j++) {
    var h = headers[j];
    // 行2テンプレの式列は directMap 以外触らない（skipMap でも空書きしない）
    if (formulaCols[j] && !directMap[h]) continue;
    var col = [];
    for (i = 0; i < n; i++) {
      if (skipMap[h]) {
        col.push(['']);
      } else {
        var val = recs[i][h];
        col.push([val !== '' && val != null ? val : '']);
      }
    }
    sh.getRange(startRow, j + 1, n, 1).setValues(col);
  }
}

function prListHeaders_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var raw = sh.getRange(PR_LIST_HEADER_ROW_, 1, 1, lastCol).getValues()[0];
  var out = [];
  var i;
  for (i = 0; i < raw.length; i++) {
    out.push(String(raw[i] || '').replace(/\n/g, ' ').trim());
  }
  while (out.length && !out[out.length - 1]) out.pop();
  return out;
}

function prListFormulaColSet_(sh, numCols) {
  var set = {};
  if (sh.getLastRow() < PR_LIST_FORMULA_TEMPLATE_ROW_) return set;
  var frow = sh.getRange(PR_LIST_FORMULA_TEMPLATE_ROW_, 1, 1, numCols).getFormulas()[0];
  var i;
  for (i = 0; i < frow.length; i++) {
    if (String(frow[i] || '').indexOf('=') === 0) set[i] = true;
  }
  return set;
}

function prListImageUrlFromFull_(full) {
  var img = String((full && full['画像']) || '').trim();
  if (img.indexOf('http') === 0) return img;
  var m = img.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
  if (m && m[1]) return m[1];
  var sub = String((full && full['サブ画像']) || '').trim();
  if (sub) {
    var parts = sub.split('|');
    var j;
    for (j = 0; j < parts.length; j++) {
      var u = String(parts[j] || '').trim();
      if (u.indexOf('http') === 0) return u;
    }
  }
  return '';
}

function prListImageHyperlink_(url) {
  var u = String(url || '').trim();
  if (!u) return '';
  return '=HYPERLINK("' + u.replace(/"/g, '') + '","メイン")';
}

/** 出品 ASIN貼り付け（Keepa用）fetchKeepaProductsForSheet と同系。コード.js extractSetCount* を①専用に移植。 */
function prNormalizeFullwidthDigits_(s) {
  return String(s || '').replace(/[０-９]/g, function (ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30);
  });
}

function prExtractSetCountCandidatesFromTitle_(title) {
  if (!title || typeof title !== 'string') return [];
  var numMap = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10, '１': 1, '２': 2, '３': 3, '４': 4, '５': 5, '６': 6, '７': 7, '８': 8, '９': 9 };
  var patterns = [
    /(\d+)\s*個\s*セット/i,
    /(\d+)\s*袋\s*セット/i,
    /(\d+)\s*セット/i,
    /(\d+)\s*袋\s*×\s*(\d+)/i,
    /(\d+)\s*袋(?!\s*×)/i,
    /(三|３|五|５|十)\s*袋/i,
    /(五|５)\s*袋\s*入?/i,
    /(\d+)\s*袋\s*入/i,
    /(\d+)\s*パック/i,
    /(\d+)\s*本\s*入り/i,
    /(\d+)\s*個\s*入り/i,
    /(\d+)\s*個(?!\s*[セ入])/i,
    /×\s*(\d+)\s*袋/i,
    /(\d+)\s*P\s*入り/i,
    /(\d+)\s*P\)/i,
    /（\s*(\d+)\s*袋\s*）/,
    /[（(](\d+)[袋個][）)]/i,
  ];
  var seen = {};
  var out = [];
  var p;
  var i;
  for (p = 0; p < patterns.length; p++) {
    var m = title.match(patterns[p]);
    if (m) {
      for (i = 1; i < m.length; i++) {
        if (m[i] == null) continue;
        var n = numMap[m[i]];
        if (n == null) n = parseInt(m[i], 10);
        if (n != null && !isNaN(n) && n >= 1 && !seen[n]) {
          seen[n] = true;
          out.push(n);
        }
      }
    }
  }
  return out;
}

function prExtractSetCountFromTitle_(title) {
  var candidates = prExtractSetCountCandidatesFromTitle_(title);
  return candidates.length > 0 ? candidates[0] : null;
}

/** 計N袋・種×各・5食など ASIN貼り付け regex で拾えない分（parseSetCountFromItemName 系）。 */
function prParseSetCountFromTitleAlt_(title) {
  if (!title) return null;
  var s = prNormalizeFullwidthDigits_(String(title).trim());
  var totalBag = s.match(/(?:合計|計|全)\s*(\d+)\s*袋/i);
  if (totalBag) {
    var tn = parseInt(totalBag[1], 10);
    if (tn >= 1 && tn <= 999) return { setCount: tn, source: 'total_bag' };
  }
  var kindM = s.match(/(?:【)?(\d+)\s*種類?(?:】)?/);
  var eachM = s.match(/各\s*(\d+)\s*袋/);
  if (kindM && eachM) {
    var ke = parseInt(kindM[1], 10) * parseInt(eachM[1], 10);
    if (ke >= 1 && ke <= 999) return { setCount: ke, source: 'kind_x_each' };
  }
  var kindBag = s.match(/(\d+)\s*種\s*[×x]\s*(\d+)\s*袋/i);
  if (kindBag) {
    var kb = parseInt(kindBag[1], 10) * parseInt(kindBag[2], 10);
    if (kb >= 1 && kb <= 999) return { setCount: kb, source: 'kind_x_bag' };
  }
  var eatM = s.match(/(\d+)\s*食(?:\s|$|入)/i);
  if (eatM) {
    var e = parseInt(eatM[1], 10);
    if (e >= 1 && e <= 999) return { setCount: e, source: 'shoku' };
  }
  return null;
}

function prTitleLooksContentOnly_(title) {
  var s = prNormalizeFullwidthDigits_(String(title || '').trim());
  if (!/\d{3,}\s*入/i.test(s)) return false;
  if (/\d+\s*(個|袋|本|食|パック)\s*(セット|入)/i.test(s)) return false;
  return true;
}

function prListSaleFormLabel_(title, qty) {
  var s = String(title || '');
  var mix = /種類|種\s*[×x]|詰め合わせ|飲み比べ|詰合|アソート|よりどり|ミックス|詰合せ/i.test(s);
  var bulk = /\d+\s*袋\s*セット|まとめ|×\s*\d+\s*袋|\d+袋(?:セット|入)/i.test(s);
  if (mix && bulk) return '④別品番とセット売り+まとめ売り';
  if (mix) return '③別品番とセット売り';
  if (qty > 1) return '②同商品まとめ売り';
  return '①単品売り';
}

function prListSetQtyAndSaleForm_(full) {
  var title = String((full && full['商品名']) || '').trim();
  var qty = null;
  var src = '';
  // 種×袋・種類×各・計N袋・N食 は regex 単体より先（×5袋 だけ拾う誤りを防ぐ）
  var alt = prParseSetCountFromTitleAlt_(title);
  if (alt) {
    qty = alt.setCount;
    src = alt.source;
  }
  if (qty == null) {
    qty = prExtractSetCountFromTitle_(title);
    if (qty != null) src = 'asin_paste_regex';
  }
  if (qty == null && prTitleLooksContentOnly_(title)) {
    qty = 1;
    src = 'content_only';
  }
  if (qty == null) {
    var noi = Number(full && full['アイテム数']);
    if (isFinite(noi) && noi >= 1 && noi <= 999) {
      qty = Math.floor(noi);
      src = 'keepa_noi';
    }
  }
  if (qty == null || qty < 1) qty = 1;
  if (qty > 999) qty = 999;
  return { qty: String(qty), form: prListSaleFormLabel_(title, qty), src: src };
}

function prLoadCandPassByAsin_(opt) {
  opt = opt || {};
  var makerSet = opt.makerSet;
  var discoveryTag = opt.discoveryTag;
  var asinSet = opt.asinSet;
  var sh = prSpreadsheet_().getSheetByName(PR_CAND_SHEET_);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = {};
  var i;
  for (i = 0; i < hdr.length; i++) idx[String(hdr[i] || '')] = i;
  if (idx['ASIN'] === undefined || idx['門結果'] === undefined) return out;
  var idxMaker = makerSet ? idx['メーカー'] : -1;
  var idxTag = discoveryTag ? idx['発見経路'] : -1;
  if (makerSet && idxMaker === undefined) return out;
  if (discoveryTag && idxTag === undefined) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var seen = {};
  for (i = 0; i < vals.length; i++) {
    var asin = String(vals[i][idx['ASIN']] || '').trim().toUpperCase();
    if (!asin || seen[asin]) continue;
    if (asinSet && !asinSet[asin]) continue;
    if (makerSet) {
      var mk = String(vals[i][idxMaker] || '').trim();
      if (!mk || !makerSet[mk]) continue;
    }
    if (discoveryTag) {
      var tag = String(vals[i][idxTag] || '').trim();
      if (tag !== discoveryTag) continue;
    }
    if (String(vals[i][idx['門結果']] || '') !== '通過') continue;
    seen[asin] = true;
    var row = {};
    var k;
    for (k in idx) {
      if (Object.prototype.hasOwnProperty.call(idx, k)) row[k] = vals[i][idx[k]];
    }
    out[asin] = row;
  }
  return out;
}

/** ASIN → { rowNum, vals }（先頭行のみ。重複ASINは D1 更新対象を1行に限定）。 */
/** Property PR_LIST_DIFF_USE_FORMULA。未設定=ON。false で数式を表示結果のまま比べる旧動作。 */
function prListDiffUseFormulaEnabled_() {
  return prCkEnabled_(PR_LIST_DIFF_USE_FORMULA_PROP_);
}

function prListExistingAsinRows_(sh, headers) {
  var out = {};
  var asinI = headers.indexOf('ASIN');
  if (asinI < 0 || sh.getLastRow() < PR_LIST_DATA_START_) return out;
  var n = sh.getLastRow() - PR_LIST_DATA_START_ + 1;
  var rng = sh.getRange(PR_LIST_DATA_START_, 1, n, headers.length);
  var vals = rng.getValues();
  // 画像URL は `=HYPERLINK(...,"メイン")` を書くが、getValues は表示結果「メイン」を返す。
  // そのまま比べると必ず差分になり、洗いスライスごとに行の削除→先頭再挿入が起きる
  // （2026-09-07 実測: 画像がある4行が毎回 refresh。人が手入力した列も消えていた）。
  if (prListDiffUseFormulaEnabled_()) {
    var fx = rng.getFormulas();
    var r;
    var c;
    for (r = 0; r < vals.length; r++) {
      for (c = 0; c < headers.length; c++) {
        if (String(fx[r][c] || '').charAt(0) === '=') vals[r][c] = fx[r][c];
      }
    }
  }
  var i;
  for (i = 0; i < vals.length; i++) {
    var a = String(vals[i][asinI] || '').trim().toUpperCase();
    if (a.length === 10 && !out[a]) out[a] = { rowNum: PR_LIST_DATA_START_ + i, vals: vals[i] };
  }
  return out;
}

/** Script Property PR_KEEPA_FULL_INDEX_ENABLED。未設定=ON。false で旧・全件 flat 読込。 */
function prKeepaFullIndexEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_KEEPA_FULL_INDEX_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prEnsureKeepaFullIndexSheet_() {
  var ss = SpreadsheetApp.openById(PR_COMPETITOR_SS_);
  var sh = ss.getSheetByName(PR_KEEPA_FULL_INDEX_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(PR_KEEPA_FULL_INDEX_SHEET_);
    sh.getRange(1, 1, 1, PR_KEEPA_FULL_INDEX_HEADERS_.length).setValues([PR_KEEPA_FULL_INDEX_HEADERS_]);
    Logger.log('[仕入れ検討①] Keepaフル索引シート作成');
  } else if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, PR_KEEPA_FULL_INDEX_HEADERS_.length).setValues([PR_KEEPA_FULL_INDEX_HEADERS_]);
  }
  return sh;
}

/**
 * Keepaフルの ASIN 列だけ読んで索引を全再構築。人手メンテなし。
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} shFull
 * @return {number} 索引件数
 */
function prRebuildKeepaFullIndex_(shFull) {
  if (!prKeepaFullIndexEnabled_()) return 0;
  var sh =
    shFull || SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_KEEPA_FULL_SHEET_);
  if (!sh) throw new Error('Keepaフルシートなし');
  var idxSh = prEnsureKeepaFullIndexSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxAsin = -1;
  var h;
  for (h = 0; h < headers.length; h++) {
    if (String(headers[h] || '').trim() === 'ASIN') idxAsin = h;
  }
  if (idxAsin < 0) throw new Error('Keepaフルに ASIN 列が無い');
  var last = sh.getLastRow();
  var lines = [PR_KEEPA_FULL_INDEX_HEADERS_.slice()];
  var cache = {};
  if (last >= 2) {
    var n = last - 1;
    var asins = sh.getRange(2, idxAsin + 1, n, 1).getValues();
    var r;
    var seen = {};
    for (r = 0; r < asins.length; r++) {
      var a = String(asins[r][0] || '')
        .trim()
        .toUpperCase();
      if (!a || seen[a]) continue;
      seen[a] = true;
      var sheetRow = r + 2;
      lines.push([a, sheetRow]);
      cache[a] = sheetRow;
    }
  }
  idxSh.clear();
  idxSh.getRange(1, 1, lines.length, 2).setValues(lines);
  PR_KEEPA_FULL_INDEX_CACHE_ = cache;
  Logger.log('[仕入れ検討①] Keepaフル索引再構築 n=%s', lines.length - 1);
  return lines.length - 1;
}

/** ASIN → Keepaフル行番号。空なら自動再構築。 */
function prLoadKeepaFullIndexMap_() {
  if (PR_KEEPA_FULL_INDEX_CACHE_) return PR_KEEPA_FULL_INDEX_CACHE_;
  var out = {};
  if (!prKeepaFullIndexEnabled_()) return out;
  try {
    var idxSh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_KEEPA_FULL_INDEX_SHEET_);
    var fullSh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_KEEPA_FULL_SHEET_);
    var fullN = fullSh && fullSh.getLastRow() >= 2 ? fullSh.getLastRow() - 1 : 0;
    var idxN = idxSh && idxSh.getLastRow() >= 2 ? idxSh.getLastRow() - 1 : 0;
    if ((!idxSh || idxN === 0) && fullN > 0) {
      prRebuildKeepaFullIndex_(fullSh);
      return PR_KEEPA_FULL_INDEX_CACHE_ || {};
    }
    if (!idxSh || idxSh.getLastRow() < 2) {
      PR_KEEPA_FULL_INDEX_CACHE_ = out;
      return out;
    }
    var vals = idxSh.getRange(2, 1, idxSh.getLastRow() - 1, 2).getValues();
    var i;
    for (i = 0; i < vals.length; i++) {
      var a = String(vals[i][0] || '')
        .trim()
        .toUpperCase();
      var row = Number(vals[i][1]);
      if (!a || !(row >= 2) || out[a]) continue;
      out[a] = row;
    }
  } catch (e) {
    Logger.log('[仕入れ検討①] Keepaフル索引読取スキップ %s', String((e && e.message) || e).slice(0, 160));
  }
  PR_KEEPA_FULL_INDEX_CACHE_ = out;
  return out;
}

/**
 * 指定 ASIN だけ Keepaフルから読む（索引経由）。D1 用。
 * 索引OFF時は従来の全件 flat。
 */
function prLoadKeepaFullFlatMapForAsins_(asins, opt) {
  opt = opt || {};
  var retried = !!opt.retried;
  var want = [];
  var seen = {};
  var i;
  for (i = 0; i < (asins || []).length; i++) {
    var a = String(asins[i] || '')
      .trim()
      .toUpperCase();
    if (!a || seen[a]) continue;
    seen[a] = true;
    want.push(a);
  }
  if (!want.length) return {};
  if (!prKeepaFullIndexEnabled_()) return prLoadKeepaFullFlatMap_();
  var out = {};
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_KEEPA_FULL_SHEET_);
    if (!sh || sh.getLastRow() < 2) return out;
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxMap = prLoadKeepaFullIndexMap_();
    var rowList = [];
    for (i = 0; i < want.length; i++) {
      var row = idxMap[want[i]];
      if (row >= 2) rowList.push({ asin: want[i], row: row });
    }
    if (!rowList.length && Object.keys(idxMap).length === 0 && sh.getLastRow() >= 2) {
      prRebuildKeepaFullIndex_(sh);
      idxMap = prLoadKeepaFullIndexMap_();
      rowList = [];
      for (i = 0; i < want.length; i++) {
        row = idxMap[want[i]];
        if (row >= 2) rowList.push({ asin: want[i], row: row });
      }
    }
    rowList.sort(function (x, y) {
      return x.row - y.row;
    });
    for (i = 0; i < rowList.length; i++) {
      var vals = sh.getRange(rowList[i].row, 1, 1, lastCol).getValues()[0];
      var rec = {};
      var h;
      for (h = 0; h < hdr.length; h++) {
        var key = String(hdr[h] || '').trim();
        if (!key || key === '生JSON') continue;
        rec[key] = vals[h];
      }
      var got = String(rec['ASIN'] || '')
        .trim()
        .toUpperCase();
      if (got && got !== rowList[i].asin) {
        Logger.log(
          '[仕入れ検討①] Keepaフル索引ずれ検知 want=%s got=%s row=%s retried=%s',
          rowList[i].asin,
          got,
          rowList[i].row,
          retried ? 'Y' : 'N'
        );
        if (!retried) {
          prRebuildKeepaFullIndex_(sh);
          return prLoadKeepaFullFlatMapForAsins_(want, { retried: true });
        }
        return prLoadKeepaFullFlatMap_();
      }
      out[rowList[i].asin] = rec;
    }
  } catch (e) {
    Logger.log(
      '[仕入れ検討①] Keepaフル部分読取失敗→全件 %s',
      String((e && e.message) || e).slice(0, 160)
    );
    return prLoadKeepaFullFlatMap_();
  }
  return out;
}

function prMenuRebuildKeepaFullIndex() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  try {
    var n = prRebuildKeepaFullIndex_();
    prAppendLog_([runId, 'Keepaフル索引再構築', 'DONE', new Date(), n, '', '', 'n=' + n]);
    prMenuToast_('① Keepaフル索引', '再構築完了 ' + n + '件', 10);
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    prAppendLog_([runId, 'Keepaフル索引再構築', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/**
 * Keepaフル 生JSON から門だけ再評価→①候補更新。/product GET なし。
 * 既定は門結果空白行のみ。PR_REGATE_ONLY_BLANK=false で全行上書き（通過→落ちも可）。
 */
function prMenuRegateFromKeepaFull() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=regate_full state=RUNNING', runId);
  try {
    var r = prRegateCandidatesFromKeepaFull_({ runId: runId });
    prAppendLog_([
      runId,
      'Keepaフル再門',
      r.leftover ? 'PARTIAL' : 'DONE',
      new Date(),
      '',
      r.updated,
      '',
      r.detail,
    ]);
    Logger.log('[仕入れ検討①] runId=%s step=regate_full state=%s %s', runId, r.leftover ? 'PARTIAL' : 'DONE', r.detail);
    prMenuToast_('① Keepaフル再門', r.detail, 12);
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=regate_full state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'Keepaフル再門', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/**
 * @param {{runId?:string, onlyBlank?:boolean}} opt
 * @return {{updated:number, pass:number, drop:number, leftover:number, detail:string}}
 */
function prRegateCandidatesFromKeepaFull_(opt) {
  opt = opt || {};
  prRefuseOriginalSs_();
  if (!prRegateFromKeepaFullEnabled_()) {
    throw new Error(PR_REGATE_FROM_FULL_ENABLED_PROP_ + '=false（切戻し中）');
  }
  var runId = opt.runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var onlyBlank = opt.onlyBlank !== false && prRegateOnlyBlankGate_();
  var sh = prSpreadsheet_().getSheetByName(PR_CAND_SHEET_);
  if (!sh || sh.getLastRow() < 2) {
    return { updated: 0, pass: 0, drop: 0, leftover: 0, detail: '候補0 onlyBlank=' + (onlyBlank ? 'Y' : 'N') };
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = {};
  var i;
  for (i = 0; i < hdr.length; i++) idx[String(hdr[i] || '')] = i;
  if (idx['ASIN'] === undefined || idx['門結果'] === undefined) {
    throw new Error('①候補に ASIN／門結果列なし');
  }
  PR_KEEPA_FULL_MAP_CACHE_ = null;
  PR_CAND_MEM_ = null;
  prLoadKeepaFullMap_();
  var prof = prReadProfile_();
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var updated = 0;
  var pass = 0;
  var drop = 0;
  var skipGate = 0;
  var skipNoFull = 0;
  var leftover = 0;
  var r;
  for (r = 0; r < vals.length; r++) {
    if (prGasTimeUp_()) {
      leftover = vals.length - r;
      Logger.log('[仕入れ検討①] runId=%s regate timeup leftover=%s', runId, leftover);
      break;
    }
    var asin = String(vals[r][idx['ASIN']] || '')
      .trim()
      .toUpperCase();
    if (!asin) continue;
    var oldGate = String(vals[r][idx['門結果']] || '').trim();
    if (onlyBlank && oldGate) {
      skipGate++;
      continue;
    }
    var p = prProductFromKeepaFullAsin_(asin);
    if (!p) {
      skipNoFull++;
      continue;
    }
    var sid = String(idx['sellerId'] !== undefined ? vals[r][idx['sellerId']] || '' : '')
      .split('|')[0]
      .trim();
    var shop = idx['店名'] !== undefined ? String(vals[r][idx['店名']] || '') : '';
    var row = prGateProductRow_(p, prof, sid, runId, shop);
    if (idx['メーカー'] !== undefined) {
      var om = String(vals[r][idx['メーカー']] || '').trim();
      if (om) row[0] = om;
    }
    if (idx['発見経路'] !== undefined) {
      var ot = String(vals[r][idx['発見経路']] || '').trim();
      if (ot) row[6] = ot;
    }
    prUpsertCandidate_(row, { forceGate: true });
    updated++;
    if (row[8] === '通過') pass++;
    else drop++;
  }
  var detail =
    'updated=' +
    updated +
    ' pass=' +
    pass +
    ' drop=' +
    drop +
    ' skip_gate=' +
    skipGate +
    ' skip_nofull=' +
    skipNoFull +
    ' leftover=' +
    leftover +
    ' onlyBlank=' +
    (onlyBlank ? 'Y' : 'N') +
    ' priceMin=' +
    prof.priceMin +
    ' priceMax=' +
    prof.priceMax +
    ' rankMin=' +
    prof.rankMin +
    ' rankMax=' +
    prof.rankMax +
    ' GET=0';
  return { updated: updated, pass: pass, drop: drop, leftover: leftover, detail: detail };
}

function prLoadKeepaFullFlatMap_() {
  var out = {};
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_KEEPA_FULL_SHEET_);
    if (!sh || sh.getLastRow() < 2) return out;
    var lastCol = sh.getLastColumn();
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxAsin = -1;
    var h;
    for (h = 0; h < hdr.length; h++) {
      if (String(hdr[h] || '').trim() === 'ASIN') idxAsin = h;
    }
    if (idxAsin < 0) return out;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      var asin = String(vals[r][idxAsin] || '').trim().toUpperCase();
      if (!asin || out[asin]) continue;
      var rec = {};
      for (h = 0; h < hdr.length; h++) {
        var key = String(hdr[h] || '').trim();
        if (!key || key === '生JSON') continue;
        rec[key] = vals[r][h];
      }
      out[asin] = rec;
    }
  } catch (e) {
    Logger.log('[仕入れ検討①] Keepaフル flat 読取スキップ %s', String((e && e.message) || e).slice(0, 160));
  }
  return out;
}

function prFullRecFromProduct_(p) {
  var fetched = prNowFetchedAt_();
  var rec = prProductToFullRec_(p, fetched);
  delete rec._csvLeft;
  return rec;
}

function prCategoryNameFromFull_(full) {
  var cat = String(full['カテゴリ: ルート'] || '').trim();
  if (cat) return cat;
  try {
    var p = JSON.parse(String(full['生JSON'] || ''));
    var tree = p.categoryTree || [];
    if (tree.length && tree[0] && tree[0].name) return String(tree[0].name);
  } catch (e1) {
    return '';
  }
  return '';
}

function prMinCartNewPrice_(full) {
  var vals = [];
  var keys = [
    'Buy Box: 現在価格',
    '新品: 現在価格',
    'Buy Box: 90 日平均',
    '新品: 90 日平均',
    'Buy Box: 30 日平均',
  ];
  var i;
  for (i = 0; i < keys.length; i++) {
    var v = String(full[keys[i]] || '')
      .trim()
      .replace(/,/g, '');
    if (!v) continue;
    var n = Number(v);
    if (!isNaN(n) && n > 0) vals.push(n);
  }
  if (!vals.length) return '';
  var m = Math.min.apply(null, vals);
  return String(m === Math.floor(m) ? Math.floor(m) : m);
}

function prRankDisplay90_(full) {
  var r = String(full['売れ筋ランキング: 90 日平均'] || '').trim();
  if (!r) r = String(full['売れ筋ランキング: 現在'] || '').trim();
  var cat = prCategoryNameFromFull_(full);
  if (r && cat) {
    try {
      return Number(r).toLocaleString('ja-JP') + '位 | ' + cat;
    } catch (e1) {
      return r + ' | ' + cat;
    }
  }
  if (r) {
    try {
      return Number(r).toLocaleString('ja-JP') + '位';
    } catch (e2) {
      return r;
    }
  }
  return r;
}

function prPackSizeFromFull_(full) {
  var l = full['梱包_L_cm'];
  var w = full['梱包_W_cm'];
  var h = full['梱包_H_cm'];
  if (!(l && w && h)) {
    try {
      var p = JSON.parse(String(full['生JSON'] || ''));
      function cm(v) {
        var n = Number(v);
        return n > 0 ? n / 10.0 : null;
      }
      l = cm(p.packageLength);
      w = cm(p.packageWidth);
      h = cm(p.packageHeight);
    } catch (e1) {
      return '';
    }
  }
  if (!(l && w && h)) return '';
  return String(Math.round(Number(l) * 10) / 10) + ' x ' + String(Math.round(Number(w) * 10) / 10) + ' x ' + String(Math.round(Number(h) * 10) / 10) + 'cm';
}

function prPackWeightFromFull_(full) {
  var g = full['梱包_重量_g'];
  if (g) {
    var s = String(g);
    return s.indexOf('g') >= 0 ? s : s + 'g';
  }
  try {
    var p = JSON.parse(String(full['生JSON'] || ''));
    var w = p.packageWeight;
    if (w != null && w !== '' && w !== -1) return String(w) + 'g';
  } catch (e1) {
    return '';
  }
  return '';
}

function prListMakerName_(candRow, full) {
  if (candRow) {
    var m = String(candRow['メーカー'] || '').trim();
    if (m) return m;
  }
  return String(full['製造者'] || full['ブランド'] || '').trim();
}

function prListCellStr_(v) {
  if (v == null || v === '') return '';
  return String(v).trim();
}

function prParseFbaFeeRemarkCriteria_(remark) {
  var out = { mode: 'none', maxL: NaN, maxW: NaN, maxH: NaN, maxSum: NaN, maxWeightG: NaN };
  var s = String(remark || '')
    .replace(/,/g, '')
    .replace(/\s+/g, '');
  if (!s) return out;
  var wMatch =
    s.match(/\/\s*([0-9]+(?:\.[0-9]+)?)\s*(kg|g)/i) ||
    s.match(/([0-9]+(?:\.[0-9]+)?)\s*(kg|g)\s*以下/i) ||
    s.match(/([0-9]+(?:\.[0-9]+)?)\s*(kg|g)/i);
  if (wMatch) {
    var wv = Number(wMatch[1]);
    out.maxWeightG = String(wMatch[2]).toLowerCase() === 'kg' ? wv * 1000 : wv;
  }
  var box = s.match(/([0-9]+(?:\.[0-9]+)?)\s*[x×]\s*([0-9]+(?:\.[0-9]+)?)\s*[x×]\s*([0-9]+(?:\.[0-9]+)?)\s*cm/i);
  if (box) {
    var dims = [Number(box[1]), Number(box[2]), Number(box[3])].sort(function (a, b) {
      return b - a;
    });
    out.mode = 'box';
    out.maxL = dims[0];
    out.maxW = dims[1];
    out.maxH = dims[2];
    return out;
  }
  var sum = s.match(/(?:三辺合計|合計)?\s*([0-9]+(?:\.[0-9]+)?)\s*cm/i);
  if (sum) {
    out.mode = 'sum';
    out.maxSum = Number(sum[1]);
  }
  return out;
}

function prGetFbaFeeTableFromSettings_() {
  if (PR_FBA_TABLE_CACHE_) return PR_FBA_TABLE_CACHE_;
  var list = [];
  try {
    var sheet = SpreadsheetApp.openById(PR_LISTING_SS_).getSheetByName(PR_SETTINGS_MASTER_SHEET_);
    if (!sheet || sheet.getLastRow() < 2) {
      PR_FBA_TABLE_CACHE_ = list;
      return list;
    }
    var data = sheet.getRange(1, 1, sheet.getLastRow(), 6).getValues();
    var ord = 0;
    var r;
    for (r = 0; r < data.length; r++) {
      if (String(data[r][0] || '').trim() !== 'FBA手数料') continue;
      var name = String(data[r][1] || '').trim();
      var fee = Number(String(data[r][3] != null ? data[r][3] : '').replace(/,/g, ''));
      var remark = String(data[r][5] != null ? data[r][5] : '').trim();
      if (!name || !isFinite(fee) || fee < 0) continue;
      var parsed = prParseFbaFeeRemarkCriteria_(remark);
      parsed.name = name;
      parsed.fee = fee;
      parsed.remark = remark;
      parsed.rowOrder = ord++;
      list.push(parsed);
    }
  } catch (e) {
    Logger.log('[仕入れ検討①] FBA設定表読取スキップ %s', String((e && e.message) || e).slice(0, 120));
  }
  PR_FBA_TABLE_CACHE_ = list;
  return list;
}

/** 出品 00_設定マスタ FBA first-fit（L1／コード.js P1b と同系）。 */
function prPickFbaTierFromSettings_(lCm, wCm, hCm, weightG) {
  var sides = [Number(lCm), Number(wCm), Number(hCm)].filter(function (x) {
    return isFinite(x) && x > 0;
  });
  if (sides.length < 3) return { tier: '', fee: NaN, sumCm: NaN };
  sides.sort(function (a, b) {
    return b - a;
  });
  var longest = sides[0];
  var median = sides[1];
  var shortest = sides[2];
  var sumCm = longest + median + shortest;
  var g = isFinite(weightG) ? weightG : NaN;
  var inStd = longest <= 45 && median <= 35 && shortest <= 20 && (!isFinite(g) || g <= 9000);
  var table = prGetFbaFeeTableFromSettings_();
  var i;
  for (i = 0; i < table.length; i++) {
    var t = table[i];
    if (t.mode === 'none') continue;
    if (/^小型|^標準/.test(t.name) && !inStd) continue;
    if (/^大型|^特大/.test(t.name) && inStd) continue;
    if (isFinite(t.maxWeightG) && isFinite(g) && g > t.maxWeightG) continue;
    if (isFinite(t.maxWeightG) && !isFinite(g)) continue;
    if (t.mode === 'box') {
      if (longest <= t.maxL && median <= t.maxW && shortest <= t.maxH) {
        return { tier: t.name, fee: t.fee, sumCm: sumCm };
      }
      continue;
    }
    if (t.mode === 'sum' && isFinite(t.maxSum) && sumCm <= t.maxSum) {
      return { tier: t.name, fee: t.fee, sumCm: sumCm };
    }
  }
  return { tier: '', fee: NaN, sumCm: sumCm };
}

function prListPackWeightG_(full) {
  var g = full['梱包_重量_g'];
  if (g != null && g !== '') {
    var n = Number(String(g).replace(/g/i, '').replace(/,/g, '').trim());
    if (!isNaN(n) && n > 0) return n;
  }
  return NaN;
}

function prListPackDimsFromFull_(full) {
  return {
    l: Number(full['梱包_L_cm']) || NaN,
    w: Number(full['梱包_W_cm']) || NaN,
    h: Number(full['梱包_H_cm']) || NaN,
    g: prListPackWeightG_(full),
  };
}

/** クイックショップ／FBA手数料。Keepaフル L1 列 → 設定マスタ first-fit → Keepa FBA手数料。 */
function prListFbaFieldsFromFull_(full) {
  var tier = prListCellStr_(full['出品FBAティア']);
  var fee = prListCellStr_(full['出品FBA手数料']);
  if (!fee) fee = prListCellStr_(full['FBA手数料']);
  var sumCm = prListCellStr_(full['梱包3辺合計_cm']);
  if (!tier || !fee || !sumCm) {
    var dims = prListPackDimsFromFull_(full);
    if (isFinite(dims.l) && isFinite(dims.w) && isFinite(dims.h)) {
      var picked = prPickFbaTierFromSettings_(dims.l, dims.w, dims.h, dims.g);
      if (!tier && picked.tier) tier = picked.tier;
      if (!fee && isFinite(picked.fee)) fee = String(picked.fee === Math.floor(picked.fee) ? Math.floor(picked.fee) : picked.fee);
      if (!sumCm && isFinite(picked.sumCm)) sumCm = String(Math.round(picked.sumCm * 10) / 10);
    }
  }
  return { tier: tier, fee: fee, sumCm: sumCm };
}

function prListBuildRec_(headers, full, candRow) {
  var rec = {};
  var i;
  for (i = 0; i < headers.length; i++) rec[headers[i]] = '';
  rec['調査日'] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  rec['メーカー名'] = prListMakerName_(candRow, full);
  rec['ASIN'] = String(full['ASIN'] || '').trim().toUpperCase();
  rec['ランキング'] = prRankDisplay90_(full);
  rec['過去1か月販売数（表示あれば）'] = full['月間売上'] != null ? String(full['月間売上']) : '';
  rec['商品名'] = String(full['商品名'] || '');
  rec['JANコード'] = String(full['商品コード: EAN'] || '');
  rec['サイズ'] = prPackSizeFromFull_(full);
  rec['重量'] = prPackWeightFromFull_(full);
  rec['出品者数'] = full['出品者数'] != null ? String(full['出品者数']) : '';
  rec['画像URL'] = prListImageHyperlink_(prListImageUrlFromFull_(full));
  rec['Keepaグラフ'] = String(full['URL: Keepa'] || '');
  rec['amazonページ'] = String(full['URL: Amazon'] || '');
  rec['最安値'] = prMinCartNewPrice_(full);
  var fba = prListFbaFieldsFromFull_(full);
  rec['FBA手数料'] = fba.fee;
  rec['amazon本体が出品者にいる'] = prListCellStr_(full['Amazon直販']);
  var rank90 = full['売れ筋ランキング: 90 日平均'] != null ? String(full['売れ筋ランキング: 90 日平均']) : '';
  if (!String(rank90 || '').trim()) {
    rank90 = full['売れ筋ランキング: 現在'] != null ? String(full['売れ筋ランキング: 現在']) : '';
  }
  rec['順位'] = rank90;
  rec['カテゴリー'] = prCategoryNameFromFull_(full);
  rec['クイックショップサイズ表記'] = fba.tier;
  rec['自己配送 送料概算'] = full['自己発送送料'] != null ? String(full['自己発送送料']) : '';
  rec['サイズ3辺合計'] = fba.sumCm;
  rec['税率'] = 0.08;
  var setPack = prListSetQtyAndSaleForm_(full);
  rec['同商品のセット数量'] = setPack.qty;
  rec['販売形態'] = setPack.form;
  return rec;
}

/**
 * D1→D2 連続。門通過→品番リストのあと機械絞込まで1メニュー。
 * 単独の D1／D2 は再実行・切り分け用に残す。
 */
function prMenuListD1ThenD2() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=D1_D2 state=RUNNING', runId);
  try {
    prMenuListAppendD1({});
    prMenuListFilterD2();
    prAppendLog_([runId, 'D1→D2', 'DONE', new Date(), '', '', '', 'D1_then_D2']);
    prMenuToast_('① D1→D2', '完了（品番リスト→見積候補）');
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=D1_D2 state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'D1→D2', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    // D1／D2 側で既にアラート済み
    throw e;
  }
}

/**
 * D1: ①候補 門通過 → Keepaフル JOIN → 仕入検討品番リスト。
 * 新規は行5から挿入。既存は機械列に差分があれば上書きして行5へ移動。差分なしは触らない。
 * 紐付け正本: dry_t1_list_paste.py（メーカー=①候補、最安=min(BB,新品)、順位=90日平均）。
 * opt.midWash: 洗い途中（予算リセットしない・toast/alert 抑制）。戻り {n, detail}。
 */
function prMenuListAppendD1(opt) {
  opt = opt || {};
  var midWash = !!opt.midWash;
  var quiet = !!opt.quiet || midWash;
  if (!midWash) prGasStart_();
  var runId =
    (opt.runId && String(opt.runId).trim()) ||
    'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=D1 state=RUNNING%s', runId, midWash ? ' midWash' : '');
  try {
    prRefuseOriginalSs_();
    if (!prListAppendEnabled_()) {
      throw new Error(PR_LIST_APPEND_ENABLED_PROP_ + '=false（切戻し中）');
    }
    var sh = prSpreadsheet_().getSheetByName(PR_LIST_SHEET_);
    if (!sh) throw new Error('シートなし: ' + PR_LIST_SHEET_);
    var headers = prListHeaders_(sh);
    if (headers.indexOf('ASIN') < 0) throw new Error('ASIN列なし');
    var formulaCols = prListFormulaColSet_(sh, headers.length);
    var skipMap = prListSkipHeadersMap_();
    var directMap = prListDirectWriteMap_();
    var candMap = prLoadCandPassByAsin_({
      makerSet: opt.makerSet,
      discoveryTag: opt.discoveryTag,
      asinSet: opt.asinSet,
    });
    var passAsins = Object.keys(candMap);
    var existRows = prListExistingAsinRows_(sh, headers);
    var existCount = Object.keys(existRows).length;
    var cap = prListAppendCap_();
    var flatMap = prLoadKeepaFullFlatMapForAsins_(passAsins);
    var needFlat = [];
    var i;
    for (i = 0; i < passAsins.length; i++) {
      if (!flatMap[passAsins[i]]) needFlat.push(passAsins[i]);
    }
    var pack = needFlat.length
      ? prKeepaProductStats90_(needFlat)
      : {
          products: [],
          fullPlus: 0,
          need: 0,
          skip: 0,
          tokensConsumed: 0,
          tokensLeft: '',
          refillRate: prKeepaRefillRate_(),
          refillIn: '',
          waitMin: 1,
        };
    var j;
    for (j = 0; j < pack.products.length; j++) {
      var fr = prFullRecFromProduct_(pack.products[j]);
      var a = String(fr.ASIN || '').trim().toUpperCase();
      if (a) flatMap[a] = fr;
    }
    var enrich = prEnrichKeepaForPassAsins_(passAsins, flatMap);
    if (enrich.enriched > 0) {
      flatMap = prLoadKeepaFullFlatMapForAsins_(passAsins);
    }
    var newAsins = [];
    var refreshAsins = [];
    var skipSame = 0;
    var missFull = 0;
    for (i = 0; i < passAsins.length; i++) {
      var asin = passAsins[i];
      if (!flatMap[asin]) {
        missFull++;
        continue;
      }
      var rec = prListBuildRec_(headers, flatMap[asin], candMap[asin]);
      if (!existRows[asin]) {
        newAsins.push(asin);
      } else if (prListRecDiffers_(headers, existRows[asin].vals, rec)) {
        refreshAsins.push(asin);
      } else {
        skipSame++;
      }
    }
    var topQueue = [];
    for (i = 0; i < passAsins.length; i++) {
      asin = passAsins[i];
      if (newAsins.indexOf(asin) >= 0 || refreshAsins.indexOf(asin) >= 0) topQueue.push(asin);
    }
    var leftover = 0;
    if (topQueue.length > cap) {
      leftover = topQueue.length - cap;
      topQueue = topQueue.slice(0, cap);
    }
    if (!topQueue.length) {
      var emptyDetail =
        'pass=' +
        passAsins.length +
        ' exist=' +
        existCount +
        ' new=' +
        newAsins.length +
        ' refresh=' +
        refreshAsins.length +
        ' skip_same=' +
        skipSame +
        ' append=0 miss_full=' +
        missFull +
        ' ' +
        (enrich && enrich.detail ? enrich.detail : '');
      prAppendLog_([runId, 'D1 品番リスト', 'DONE', new Date(), '', '', '', emptyDetail]);
      if (prCycleListCheckEnabled_()) {
        var markEmpty = {};
        for (i = 0; i < passAsins.length; i++) {
          if (existRows[passAsins[i]]) markEmpty[passAsins[i]] = true;
        }
        prCandMarkListCheck_(markEmpty, PR_CAND_LIST_CHECK_SUMI_);
      }
      if (!quiet) prMenuToast_('D1 品番リスト', '変更なし ' + emptyDetail);
      return { n: 0, detail: emptyDetail };
    }
    if (prGasTimeUp_()) {
      var waitDetail =
        'pass=' +
        passAsins.length +
        ' new=' +
        newAsins.length +
        ' refresh=' +
        refreshAsins.length +
        ' queue=' +
        topQueue.length +
        ' leftover=' +
        leftover +
        ' timeup_before_write cap=' +
        cap;
      prAppendLog_([runId, 'D1 品番リスト', 'DONE', new Date(), '', 0, '', waitDetail]);
      if (!quiet) prMenuToast_('D1 品番リスト', '時間切れ（書込前） ' + waitDetail + ' 再実行', 15);
      return { n: 0, detail: waitDetail };
    }
    var refreshInQueue = [];
    for (i = 0; i < topQueue.length; i++) {
      if (refreshAsins.indexOf(topQueue[i]) >= 0) refreshInQueue.push(topQueue[i]);
    }
    if (refreshInQueue.length) {
      var deleteRows = [];
      for (i = 0; i < refreshInQueue.length; i++) {
        if (existRows[refreshInQueue[i]]) deleteRows.push(existRows[refreshInQueue[i]].rowNum);
      }
      deleteRows.sort(function (x, y) {
        return y - x;
      });
      for (i = 0; i < deleteRows.length; i++) sh.deleteRow(deleteRows[i]);
    }
    var startRow = PR_LIST_DATA_START_;
    var n = topQueue.length;
    sh.insertRowsBefore(startRow, n);
    var tpl = sh.getRange(PR_LIST_FORMULA_TEMPLATE_ROW_, 1, 1, headers.length);
    tpl.copyTo(sh.getRange(startRow, 1, n, headers.length), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    var recs = [];
    for (i = 0; i < n; i++) {
      recs.push(prListBuildRec_(headers, flatMap[topQueue[i]], candMap[topQueue[i]]));
    }
    prListWriteRecs_(sh, headers, startRow, recs, formulaCols, directMap, skipMap);
    var appendNew = 0;
    var refreshN = 0;
    for (i = 0; i < topQueue.length; i++) {
      if (newAsins.indexOf(topQueue[i]) >= 0) appendNew++;
      else refreshN++;
    }
    if (prCycleListCheckEnabled_()) {
      var markDone = {};
      for (i = 0; i < topQueue.length; i++) markDone[topQueue[i]] = true;
      for (i = 0; i < passAsins.length; i++) {
        if (existRows[passAsins[i]] && newAsins.indexOf(passAsins[i]) < 0) {
          markDone[passAsins[i]] = true;
        }
      }
      prCandMarkListCheck_(markDone, PR_CAND_LIST_CHECK_SUMI_);
    }
    var detail =
      'pass=' +
      passAsins.length +
      ' exist=' +
      existCount +
      ' new=' +
      newAsins.length +
      ' refresh=' +
      refreshAsins.length +
      ' skip_same=' +
      skipSame +
      ' append=' +
      appendNew +
      ' refresh_top=' +
      refreshN +
      ' miss_full=' +
      missFull +
      ' insert=top5' +
      ' full+=' +
      pack.fullPlus +
      ' cap=' +
      cap +
      ' ' +
      (enrich && enrich.detail ? enrich.detail : '') +
      (leftover ? ' leftover=' + leftover : '');
    prAppendLog_([runId, 'D1 品番リスト', 'DONE', new Date(), '', n, '', detail]);
    Logger.log('[仕入れ検討①] runId=%s state=DONE %s', runId, detail);
    if (!quiet) prMenuToast_('D1 品番リスト', detail);
    return { n: n, detail: detail };
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=D1 state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'D1 品番リスト', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    if (!quiet) prMenuFailAlert_(msg);
    throw e;
  }
}

var PR_D2_FURUSATO_RE_ = /ふるさと|返礼品/;
var PR_D2_SELECT_RE_ = /よりどり|種類が選べ|選べるセット|から選択/;
var PR_D2_USED_RE_ = /中古品|中古/;
var PR_D2_BOOK_CATS_ = ['本', '洋書', 'ミュージック', 'DVD', 'Kindleストア', 'クラシック', 'デジタルミュージック'];

function prListFilterEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_LIST_FILTER_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prListFilterCap_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_LIST_FILTER_CAP_PROP_) ||
      PR_LIST_FILTER_CAP_DEFAULT_
  );
  if (!(n > 0)) n = PR_LIST_FILTER_CAP_DEFAULT_;
  if (n > 2000) n = 2000;
  return Math.floor(n);
}

/** 既定 true＝見積候補／見送り済み行は D2 で触らない。PR_D2_ONLY_BLANK=false で全 ASIN 再判定。 */
function prD2OnlyBlankEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_D2_ONLY_BLANK_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prListQuoteIsClassified_(flag) {
  var s = String(flag || '').trim();
  return s === '見積候補' || s === '見送り';
}

/** 品番リスト上の未判定 ASIN 件数（見積候補・見送りどちらも空）。 */
function prListCountUnclassified_() {
  prRefuseOriginalSs_();
  var sh = prSpreadsheet_().getSheetByName(PR_LIST_SHEET_);
  if (!sh) return 0;
  var headers = prListHeaders_(sh);
  headers = prListEnsureFilterCols_(sh, headers);
  var idxQuote = headers.indexOf(PR_LIST_COL_QUOTE_);
  var idxAsin = headers.indexOf('ASIN');
  if (idxQuote < 0 || idxAsin < 0) return 0;
  var lastDataRow = prListLastAsinRow_(sh, idxAsin + 1);
  if (lastDataRow < PR_LIST_DATA_START_) return 0;
  var numRows = lastDataRow - PR_LIST_DATA_START_ + 1;
  var vals = sh.getRange(PR_LIST_DATA_START_, 1, numRows, headers.length).getValues();
  var n = 0;
  var i;
  for (i = 0; i < vals.length; i++) {
    var asin = String(vals[i][idxAsin] || '').trim().toUpperCase();
    if (asin.length !== 10) continue;
    if (prListQuoteIsClassified_(vals[i][idxQuote])) continue;
    n++;
  }
  return n;
}

function prCycleD2OverflowGet_() {
  return prCyclePipeListGet_(PR_CYCLE_D2_OVERFLOW_PROP_);
}

function prCycleD2OverflowActive_() {
  return prCycleD2OverflowGet_().length > 0;
}

function prCycleD2OverflowClear_() {
  PropertiesService.getScriptProperties().deleteProperty(PR_CYCLE_D2_OVERFLOW_PROP_);
}

/**
 * D2 で leftover=0・未判定=0 になったら overflow を解消し C4 を1社自動再開。
 * @return {{resumed:boolean, detail:string, parent?:string, dOk?:boolean}}
 */
function prMaybeAutoResumeAfterD2Clear_(runId, d2r) {
  d2r = d2r || {};
  if ((Number(d2r.leftover) || 0) > 0) return { resumed: false, detail: 'leftover>0' };
  if (prListCountUnclassified_() > 0) return { resumed: false, detail: 'unclassified>0' };
  var pending = prCycleD2OverflowGet_();
  if (!pending.length) return { resumed: false, detail: 'no_overflow' };
  var props = PropertiesService.getScriptProperties();
  var parent = pending[0];
  var last = String(props.getProperty(PR_CYCLE_C4_LAST_PARENT_PROP_) || '').trim();
  if (last && pending.indexOf(last) >= 0) parent = last;
  prCycleD2OverflowClear_();
  var needs = prCycleListNeedsD1_();
  if (needs.indexOf(parent) < 0) {
    prAppendLog_([
      runId,
      'D2 overflow解消',
      'DONE',
      new Date(),
      '',
      parent,
      '',
      'cleared no_needs parent=' + parent,
    ]);
    return { resumed: false, detail: 'cleared_no_needs', parent: parent };
  }
  try {
    var r = prRunCycleC4Chain_({
      parent: parent,
      runId: runId,
      tag: 'd2_overflow_auto_c4',
      quiet: true,
    });
    return { resumed: true, detail: String(r.detail || ''), parent: parent, dOk: !!r.dOk };
  } catch (eAuto) {
    var msg = prRedactKey_(String((eAuto && eAuto.message) || eAuto)).slice(0, 160);
    prAppendLog_([runId, 'D2 overflow解消', 'FAILED', new Date(), '', parent, '', msg]);
    return { resumed: false, detail: msg, parent: parent };
  }
}

/** D2 上限超停止時の通知（本文に復帰手順）。 */
function prSendD2OverflowMail_(opt) {
  opt = opt || {};
  var to = String(
    PropertiesService.getScriptProperties().getProperty(PR_CYCLE_DAILY_MAIL_TO_PROP_) ||
      Session.getActiveUser().getEmail() ||
      ''
  ).trim();
  if (!to) {
    prAppendLog_([opt.runId || '', 'D2上限超メール', 'FAILED', new Date(), '', '', '', 'no_to']);
    return;
  }
  var day = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var parent = String(opt.parent || '').trim();
  var leftover = Number(opt.leftover) || 0;
  var unclassified = Number(opt.unclassified) || 0;
  var cap = Number(opt.cap) || prListFilterCap_();
  var processed = Math.max(0, unclassified - leftover);
  var subject = '[①] D2未判定が上限超 parent=' + parent + ' leftover=' + leftover;
  var body =
    'D2 機械絞込が1回あたり上限（cap）を超える未判定 ASIN が残ったため、C4 完走連鎖を停止しました。\n\n' +
    '停止したC4メーカー: ' +
    parent +
    '\n' +
    '未判定ASIN（リスト全体）: ' +
    unclassified +
    ' 件\n' +
    '今回処理: 約 ' +
    processed +
    ' 件 / 残り（leftover）: ' +
    leftover +
    ' 件 / cap=' +
    cap +
    '\n' +
    'runId: ' +
    String(opt.runId || '') +
    '\n' +
    '詳細: ' +
    String(opt.detail || '') +
    '\n\n' +
    '--- 復帰手順 ---\n' +
    '① スプシ「仕入検討品番リスト」で 見積候補 が空の ASIN 行数を確認\n' +
    '② メニュー 99 → 6「D2だけ（見積候補に絞る）」を実行\n' +
    '   ①ログの leftover=0 になるまで繰り返す\n' +
    '   （未判定が多い場合: Script Properties で一時 PR_LIST_FILTER_CAP=2000 → D2 → 1000に戻す）\n' +
    '③ メニュー 99 → 28「D2 backlog解消→C4再開」を1回実行\n' +
    '   （または 4「完走連鎖だけ実行」で needs 先頭のメーカー）\n' +
    '④ ①ログで C4 完走連鎖 d_ok を確認\n\n' +
    '※ 洗いのやり直し・A-1 は不要です。D1 は skip_same で実質スキップされます。\n';
  try {
    MailApp.sendEmail(to, subject, body);
    prAppendLog_([
      opt.runId || '',
      'D2上限超メール',
      'DONE',
      new Date(),
      leftover,
      parent,
      '',
      'to=' + to + ' unclassified=' + unclassified,
    ]);
  } catch (eMail) {
    prAppendLog_([
      opt.runId || '',
      'D2上限超メール',
      'FAILED',
      new Date(),
      leftover,
      parent,
      '',
      prRedactKey_(String((eMail && eMail.message) || eMail)).slice(0, 200),
    ]);
    throw eMail;
  }
}

/** ①候補を ASIN キーで全件（門落ち含む）。 */
function prLoadCandAllByAsin_() {
  var sh = prSpreadsheet_().getSheetByName(PR_CAND_SHEET_);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = {};
  var i;
  for (i = 0; i < hdr.length; i++) idx[String(hdr[i] || '')] = i;
  if (idx['ASIN'] === undefined) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  for (i = 0; i < vals.length; i++) {
    var asin = String(vals[i][idx['ASIN']] || '').trim().toUpperCase();
    if (!asin) continue;
    var row = {};
    var k;
    for (k in idx) {
      if (Object.prototype.hasOwnProperty.call(idx, k)) row[k] = vals[i][idx[k]];
    }
    out[asin] = row;
  }
  return out;
}

/** メーカー名 → 実態。 */
function prLoadMakerJittaMap_() {
  var out = {};
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
    if (!sh || sh.getLastRow() < 2) return out;
    var lastCol = sh.getLastColumn();
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxName = -1;
    var idxJ = -1;
    var h;
    for (h = 0; h < hdr.length; h++) {
      var name = String(hdr[h] || '').trim();
      if (name === 'メーカー') idxName = h;
      if (name === '実態') idxJ = h;
    }
    if (idxName < 0 || idxJ < 0) return out;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      var m = String(vals[r][idxName] || '').trim();
      var j = String(vals[r][idxJ] || '').trim();
      if (m) out[m] = j || PR_JITTAI_MAKER_;
    }
  } catch (e) {
    Logger.log('[仕入れ検討①] メーカーサマリ jitta 読取スキップ %s', String((e && e.message) || e).slice(0, 160));
  }
  return out;
}

function prD2IsBookCategory_(cat) {
  var c = String(cat || '').trim();
  if (!c) return false;
  var i;
  for (i = 0; i < PR_D2_BOOK_CATS_.length; i++) {
    var b = PR_D2_BOOK_CATS_[i];
    if (c === b || c.indexOf(b) === 0) return true;
  }
  return false;
}

/**
 * D2 機械絞込。見積候補／見送り＋理由。粗利は見ない。行は消さない。
 * @param {Object=} listExtra 品番リスト行のクイックショップサイズ表記・重量
 * @return {{flag:string, reason:string}}
 */
function prD2Classify_(asin, maker, title, category, candRow, jitta, listExtra) {
  listExtra = listExtra || {};
  if (jitta === PR_JITTAI_EXCLUDE_) return { flag: '見送り', reason: '非仕入先' };
  if (jitta === PR_JITTAI_UNCHECKED_) return { flag: '見送り', reason: '未調査' };
  if (jitta === PR_JITTAI_CHECK_) return { flag: '見送り', reason: '要確認' };
  if (!candRow) return { flag: '見送り', reason: '①候補なし' };
  var gate = String(candRow['門結果'] || '').trim();
  if (gate === '落ち') return { flag: '見送り', reason: String(candRow['門理由'] || '門落ち').slice(0, 80) };
  if (gate !== '通過') return { flag: '見送り', reason: '門未通過' };
  var t = String(title || '');
  if (PR_D2_FURUSATO_RE_.test(t)) return { flag: '見送り', reason: 'ふるさと' };
  if (PR_D2_SELECT_RE_.test(t)) return { flag: '見送り', reason: '選択式' };
  if (PR_D2_USED_RE_.test(t)) return { flag: '見送り', reason: '中古' };
  if (prD2IsBookCategory_(category)) return { flag: '見送り', reason: '本系カテゴリ' };
  var profs = prReadAllProfiles_();
  var rowProf = prProfileForCategoryName_(category, profs);
  var qsTier = String(listExtra.qsTier || '');
  if (prProfileOversizeDrop_(rowProf) && prD2IsOversizeTier_(qsTier)) {
    return { flag: '見送り', reason: '大型サイズ' };
  }
  var wg = prParseWeightG_(listExtra.weight);
  var wmax = prProfileWeightMaxG_(rowProf);
  if (isFinite(wg) && wg >= wmax) {
    return { flag: '見送り', reason: '重量>=' + wmax + 'g' };
  }
  return { flag: '見積候補', reason: '' };
}

function prListEnsureFilterCols_(sh, headers) {
  var out = headers.slice();
  var need = [PR_LIST_COL_QUOTE_, PR_LIST_COL_SKIP_REASON_];
  var i;
  for (i = 0; i < need.length; i++) {
    if (out.indexOf(need[i]) < 0) {
      out.push(need[i]);
      sh.getRange(PR_LIST_HEADER_ROW_, out.length).setValue(need[i]);
    }
  }
  return out;
}

function prListLastAsinRow_(sh, asinCol1Based) {
  var last = sh.getLastRow();
  if (last < PR_LIST_DATA_START_) return PR_LIST_DATA_START_ - 1;
  var col = sh.getRange(PR_LIST_DATA_START_, asinCol1Based, last - PR_LIST_DATA_START_ + 1, 1).getValues();
  var i;
  for (i = col.length - 1; i >= 0; i--) {
    var a = String(col[i][0] || '').trim().toUpperCase();
    if (a.length === 10) return PR_LIST_DATA_START_ + i;
  }
  return PR_LIST_DATA_START_ - 1;
}

/**
 * D2: 仕入検討品番リスト全行に見積候補／見送りを書く。①候補・メーカーサマリ参照。
 * @return {{ok:boolean, leftover:number, quote:number, skip:number, nAsin:number, detail:string}}
 */
function prRunListFilterD2_(opt) {
  opt = opt || {};
  var runId =
    opt.runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var quiet = !!opt.quiet;
  prRefuseOriginalSs_();
  if (!prListFilterEnabled_()) {
    throw new Error(PR_LIST_FILTER_ENABLED_PROP_ + '=false（切戻し中）');
  }
  var sh = prSpreadsheet_().getSheetByName(PR_LIST_SHEET_);
  if (!sh) throw new Error('シートなし: ' + PR_LIST_SHEET_);
  var headers = prListHeaders_(sh);
  headers = prListEnsureFilterCols_(sh, headers);
  var idxQuote = headers.indexOf(PR_LIST_COL_QUOTE_);
  var idxReason = headers.indexOf(PR_LIST_COL_SKIP_REASON_);
  var idxAsin = headers.indexOf('ASIN');
  var idxMaker = headers.indexOf('メーカー名');
  var idxTitle = headers.indexOf('商品名');
  var idxCat = headers.indexOf('カテゴリー');
  var idxQs = headers.indexOf('クイックショップサイズ表記');
  var idxWeight = headers.indexOf('重量');
  if (idxQuote < 0 || idxReason < 0 || idxAsin < 0) {
    throw new Error('見積候補列またはASIN列なし');
  }
  var lastDataRow = prListLastAsinRow_(sh, idxAsin + 1);
  if (lastDataRow < PR_LIST_DATA_START_) {
    throw new Error('ASINデータ行なし');
  }
  var numRows = lastDataRow - PR_LIST_DATA_START_ + 1;
  var cap = prListFilterCap_();
  var onlyBlank = prD2OnlyBlankEnabled_();
  var vals = sh.getRange(PR_LIST_DATA_START_, 1, numRows, headers.length).getValues();
  var candMap = prLoadCandAllByAsin_();
  var jittaMap = prLoadMakerJittaMap_();
  var flags = [];
  var reasons = [];
  var nQuote = 0;
  var nSkip = 0;
  var nBlank = 0;
  var nAsin = 0;
  var nWork = 0;
  var nProcessed = 0;
  var nSkippedClassified = 0;
  var leftover = 0;
  var i;
  for (i = 0; i < vals.length; i++) {
    var row = vals[i];
    var asin = String(row[idxAsin] || '').trim().toUpperCase();
    var prevFlag = String(row[idxQuote] || '');
    var prevReason = String(row[idxReason] || '');
    if (asin.length !== 10) {
      flags.push([prevFlag]);
      reasons.push([prevReason]);
      nBlank++;
      continue;
    }
    nAsin++;
    if (onlyBlank && prListQuoteIsClassified_(prevFlag)) {
      flags.push([prevFlag]);
      reasons.push([prevReason]);
      nSkippedClassified++;
      continue;
    }
    nWork++;
    if (nWork > cap) {
      flags.push([prevFlag]);
      reasons.push([prevReason]);
      leftover++;
      continue;
    }
    var maker = idxMaker >= 0 ? String(row[idxMaker] || '').trim() : '';
    var title = idxTitle >= 0 ? String(row[idxTitle] || '') : '';
    var cat = idxCat >= 0 ? String(row[idxCat] || '') : '';
    var jitta = jittaMap[maker] || PR_JITTAI_MAKER_;
    var cls = prD2Classify_(asin, maker, title, cat, candMap[asin], jitta, {
      qsTier: idxQs >= 0 ? String(row[idxQs] || '') : '',
      weight: idxWeight >= 0 ? String(row[idxWeight] || '') : '',
    });
    flags.push([cls.flag]);
    reasons.push([cls.reason]);
    nProcessed++;
    if (cls.flag === '見積候補') nQuote++;
    else nSkip++;
  }
  sh.getRange(PR_LIST_DATA_START_, idxQuote + 1, flags.length, 1).setValues(flags);
  sh.getRange(PR_LIST_DATA_START_, idxReason + 1, flags.length, 1).setValues(reasons);
  var detail =
    'sheet_rows=' +
    vals.length +
    ' asin=' +
    nAsin +
    ' quote=' +
    nQuote +
    ' skip=' +
    nSkip +
    ' leftover=' +
    leftover +
    ' blank_rows=' +
    nBlank +
    ' cap=' +
    cap +
    ' onlyBlank=' +
    (onlyBlank ? 'Y' : 'N') +
    ' work=' +
    nWork +
    ' processed=' +
    nProcessed +
    ' skip_classified=' +
    nSkippedClassified;
  if (!opt.skipLog) {
    prAppendLog_([runId, 'D2 機械絞込', 'DONE', new Date(), '', nQuote + nSkip, '', detail]);
  }
  if (!quiet) {
    Logger.log('[仕入れ検討①] runId=%s state=DONE %s', runId, detail);
  }
  if (!quiet && !opt.skipToast) {
    prMenuToast_('D2 機械絞込', detail);
  }
  // 自動化: 見積候補が付いたあと単品名を埋める（未充填の続き。切戻し PR_SINGLE_NAME_AFTER_D2=false）
  if (!opt.skipSingleName && prSingleNameAfterD2Enabled_()) {
    try {
      if (!PR_GAS_START_MS_) prGasStart_();
      var sn = prRunSingleNameFill_({
        runId: runId,
        quiet: true,
        skipToast: true,
        fromD2: true,
      });
      detail += ' | S ' + (sn.detail || '');
      if (!quiet && !opt.skipToast && sn.leftover > 0) {
        prMenuToast_('S 単品名', '残り leftover=' + sn.leftover + ' → Sメニュー再実行', 10);
      }
    } catch (eSn) {
      var snMsg = prRedactKey_(String((eSn && eSn.message) || eSn)).slice(0, 120);
      Logger.log('[仕入れ検討①] runId=%s single_name_after_D2 skip %s', runId, snMsg);
      detail += ' | S_skip=' + snMsg;
    }
  }
  return {
    ok: true,
    leftover: leftover,
    quote: nQuote,
    skip: nSkip,
    nAsin: nAsin,
    nWork: nWork,
    nProcessed: nProcessed,
    detail: detail,
  };
}

function prMenuListFilterD2() {
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=D2 state=RUNNING', runId);
  try {
    var hadOverflow = prCycleD2OverflowActive_();
    var d2r = prRunListFilterD2_({ runId: runId, quiet: false });
    var ar = prMaybeAutoResumeAfterD2Clear_(runId, d2r);
    if (ar.resumed) {
      var toastMsg =
        'overflow解消後 C4 自動再開 parent=' +
        (ar.parent || '') +
        (ar.dOk ? ' d_ok' : ' d_fail');
      prMenuToast_('D2→C4再開', toastMsg, 12);
    } else if (
      hadOverflow &&
      !(Number(d2r.leftover) || 0) &&
      prListCountUnclassified_() === 0 &&
      ar.detail !== 'no_overflow'
    ) {
      if (ar.detail === 'cleared_no_needs') {
        prMenuToast_('D2 overflow解消', 'overflow 解除（C4 needs 対象外）', 10);
      } else {
        prMenuToast_('D2 backlog解消', 'メニュー 99→28「D2 backlog解消→C4再開」を実行してください', 12);
      }
    }
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=D2 state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'D2 機械絞込', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/** 単品商品名 AI 充填。切戻し PR_SINGLE_NAME_AI_ENABLED=false。 */
function prSingleNameEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_SINGLE_NAME_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prSingleNameCap_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_SINGLE_NAME_CAP_PROP_) || PR_SINGLE_NAME_CAP_DEFAULT_);
  if (!(n > 0)) n = PR_SINGLE_NAME_CAP_DEFAULT_;
  if (n > 300) n = 300;
  return Math.floor(n);
}

/** D2 後に単品名充填。未設定=ON。 */
function prSingleNameAfterD2Enabled_() {
  if (!prSingleNameEnabled_()) return false;
  var raw = PropertiesService.getScriptProperties().getProperty(PR_SINGLE_NAME_AFTER_D2_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prSingleNameStripHint_(title) {
  var s = String(title || '');
  s = s.replace(/\d+\s*個\s*(セット|入)/gi, ' ');
  s = s.replace(/\d+\s*袋\s*(セット|入)/gi, ' ');
  s = s.replace(/[×xX]\s*\d+\s*(袋|個|本)/gi, ' ');
  s = s.replace(/\d+\s*食(?:入|セット)?/gi, ' ');
  s = s.replace(/まとめ|お買い得|送料無料|【[^】]*】/g, ' ');
  return s.replace(/\s+/g, ' ').trim();
}

function prSingleNameSaleCode_(raw) {
  var s = String(raw || '');
  if (s.indexOf('④') >= 0) return '④';
  if (s.indexOf('③') >= 0 || (s.indexOf('別品番') >= 0 && s.indexOf('セット') >= 0)) return '③別品番とセット売り';
  if (s.indexOf('詰め合わせ') >= 0 || s.indexOf('アソート') >= 0) return '③別品番とセット売り';
  if (s.indexOf('②') >= 0 || s.indexOf('まとめ') >= 0) return '②';
  if (s.indexOf('①') >= 0 || s.indexOf('単品') >= 0) return '①';
  return '①';
}

function prSingleNameParseQty_(raw) {
  var n = Number(String(raw || '').replace(/,/g, ''));
  if (!(n >= 1)) return 1;
  return Math.floor(n);
}

function prSingleNameParseJans_(raw) {
  var s = String(raw || '').trim();
  if (!s) return [];
  var parts = s.split(/[,|｜\s]+/);
  var out = [];
  var i;
  for (i = 0; i < parts.length; i++) {
    var p = String(parts[i] || '').replace(/\D/g, '');
    if (p.length >= 8 && out.indexOf(p) < 0) out.push(p);
  }
  return out;
}

function prSingleNameTitleMatch_(draft, title) {
  var d = String(draft || '').replace(/\s+/g, '');
  var t = String(title || '').replace(/\s+/g, '');
  if (!d || !t) return 'low';
  if (t.indexOf(d) >= 0 || d.indexOf(t) >= 0) return 'high';
  var hit = 0;
  var i;
  for (i = 0; i < d.length - 1; i++) {
    if (t.indexOf(d.slice(i, i + 2)) >= 0) hit++;
  }
  var n = d.length - 1;
  return n > 0 && hit / n >= 0.45 ? 'high' : 'low';
}

/** Catalog 未接続時の簡易存在。仮名がタイトルに近ければ ok。 */
function prSingleNameSoftExistence_(draft, title, maker) {
  if (!String(draft || '').trim()) return 'ng';
  if (prSingleNameTitleMatch_(draft, title) === 'high') return 'ok';
  var m = String(maker || '').replace(/\s+/g, '');
  var t = String(title || '').replace(/\s+/g, '');
  var d = String(draft || '');
  if (m && t.indexOf(m) >= 0 && d.length >= 2) {
    var i;
    for (i = 0; i < d.length - 1; i++) {
      if (t.indexOf(d.slice(i, i + 2).replace(/\s+/g, '')) >= 0) return 'ok';
    }
  }
  return 'ng';
}

function prSingleNameForceMix_(title) {
  var t = String(title || '');
  var mixCue = /食べ比べ|各\d+\s*袋|詰め合わせ|アソート|よりどり/.test(t);
  var twoSku =
    /チーズかつお[\s\S]{0,40}まぐろ/.test(t) ||
    /まぐろチーズ[\s\S]{0,40}チーズかつお/.test(t) ||
    /かつお[\s\S]{0,20}まぐろ[\s\S]{0,20}ぶり/.test(t);
  if (mixCue && twoSku) return true;
  if (t.indexOf('＋') >= 0 && mixCue) return true;
  return false;
}

function prSingleNameLooksCanon_(name) {
  var s = String(name || '').trim();
  if (!s || s.indexOf('要確認') >= 0 || s.indexOf('保留') === 0) return false;
  if (s.length > 36) return false;
  if (/お魚屋さん|大人のおつまみ|徳用|大容量|業務用/.test(s)) return false;
  return true;
}

/** 人が直す行の単品商品名に載せる文言。見積には使わない。 */
function prSingleNameHumanLabel_(pattern) {
  var p = String(pattern || '');
  if (p === 'P-MIX') return '要確認（異種セット）';
  if (p === 'E1') return '要確認（AI迷い）';
  if (p === 'D1') return '要確認（存在NG）';
  if (p === 'F2') return '要確認（タイトル不一致）';
  if (p === 'F3') return '要確認（別ASIN疑い）';
  if (p === 'D2') return '保留（存在不明）';
  return '要確認（人判定）';
}

/**
 * 要確認・保留セルを目立たせる。異種=赤白、その他=黄。
 * 確定名の行は色を戻す。
 */
function prSingleNameApplyHumanStyles_(sh, vals, startRow, idxQuote, idxSingle, idxTrust, idxPat) {
  var i;
  for (i = 0; i < vals.length; i++) {
    var row = vals[i];
    if (idxQuote >= 0 && String(row[idxQuote] || '').trim() !== '見積候補') continue;
    var trust = idxTrust >= 0 ? String(row[idxTrust] || '').trim() : '';
    var pat = idxPat >= 0 ? String(row[idxPat] || '').trim() : '';
    var name = String(row[idxSingle] || '').trim();
    var human =
      trust === '要確認' ||
      trust === '保留' ||
      name.indexOf('要確認') === 0 ||
      name.indexOf('保留') === 0;
    var sheetRow = startRow + i;
    var nameCell = sh.getRange(sheetRow, idxSingle + 1);
    var trustCell = idxTrust >= 0 ? sh.getRange(sheetRow, idxTrust + 1) : null;
    if (human) {
      if (!name || (name.indexOf('要確認') !== 0 && name.indexOf('保留') !== 0)) {
        name = prSingleNameHumanLabel_(pat || (trust === '保留' ? 'D2' : ''));
        row[idxSingle] = name;
        nameCell.setValue(name);
      }
      if (pat === 'P-MIX' || name.indexOf('異種セット') >= 0) {
        nameCell.setBackground('#d32f2f').setFontColor('#ffffff').setFontWeight('bold');
        if (trustCell) trustCell.setBackground('#d32f2f').setFontColor('#ffffff').setFontWeight('bold');
      } else {
        nameCell.setBackground('#ffeb3b').setFontColor('#000000').setFontWeight('bold');
        if (trustCell) trustCell.setBackground('#ffeb3b').setFontColor('#000000').setFontWeight('bold');
      }
    } else if (name) {
      nameCell.setBackground(null).setFontColor(null).setFontWeight('normal');
      if (trustCell) trustCell.setBackground(null).setFontColor(null).setFontWeight('normal');
    }
  }
}

function prSingleNameAiDraft_(title, maker, qty, sale) {
  if (prSingleNameForceMix_(title)) {
    return { kind: 'mix', name: '', confidence: 'low', source: 'mix_guard' };
  }
  var prompt =
    'あなたは食品の仕入れ担当です。Amazonタイトルから、メーカー・問屋への見積に出す「単品商品名」だけをJSONで返してください。\n' +
    '目的: 同じ中身の商品は、袋数・ケース数・販促語・競合ショップのカタログ表記が違っても、必ず同じ最短名にする。異種の詰め合わせは見積の単品にしない。\n' +
    'メーカー: ' +
    maker +
    '\n販売形態: ' +
    sale +
    '\nセット数量: ' +
    qty +
    '\nタイトル: ' +
    title +
    '\n' +
    '規則:\n' +
    '1. 出力はJSONのみ。形式: {"kind":"single または mix","name":"..."}\n' +
    '2. kind=single の name は「味・品目＋内容量」の最少形。例: チーズかつお 220g\n' +
    '3. name に入れない: メーカー名、ショップ名、送料無料、ポイント、大容量、徳用、業務用、おつまみ、個包装、受賞、袋数（×5袋、2ケース）、【】飾り\n' +
    '4. 内容量が違うものは別商品（220g と 1kg は別 name）\n' +
    '5. 味・具材が2種以上、または「食べ比べ」「詰め合わせ」「各1袋」「＋」で別品 → kind=mix, name=""\n' +
    '6. 同じ品のまとめ売り（220g×5袋）は kind=single。袋数は落とす\n' +
    '7. メーカー名は name に入れない。型番が無い食品は型番を作らない\n' +
    '8. 品目名（食べるおだし）と味の固有名（ピングレソルティ、はしっこ）は残す\n' +
    '例: 「チーズかつお 220ｇ × 5袋」→ {"kind":"single","name":"チーズかつお 220g"}\n' +
    '例: 「お魚屋さんがつくった大人のおつまみ チーズかつお 1kg 業務用」→ {"kind":"single","name":"チーズかつお 1kg"}\n' +
    '例: 「チーズかつお220g＋まぐろ190g 各3袋」→ {"kind":"mix","name":""}\n' +
    '例: 「チーズかつお まぐろチーズ 業務用 1kg 食べ比べ」→ {"kind":"mix","name":""}\n' +
    '例: 「カルピスウォーター 300gパウチ×30本入×(2ケース)」→ {"kind":"single","name":"カルピスウォーター 300g"}\n' +
    '例: 「食べるおだし (150g, まぐろ はしっこ)」→ {"kind":"single","name":"食べるおだし まぐろ はしっこ 150g"}\n' +
    '例: 「おいしい水 天然水 ピングレソルティ 600ml×24本」→ {"kind":"single","name":"おいしい水 天然水 ピングレソルティ 600ml"}\n' +
    '例: 「おいしい水 天然水 シンプルecoラベル 600ml×24本」→ {"kind":"single","name":"おいしい水 天然水 600ml"}\n';
  var text = prOpenAiGenerateText_(prompt);
  if (!text) text = prGeminiGenerateText_(prompt);
  var fallback = prSingleNameStripHint_(title);
  if (!text) return { kind: 'single', name: fallback, confidence: 'low', source: 'strip' };
  var m = String(text).match(/\{[\s\S]*\}/);
  if (!m) return { kind: 'single', name: fallback, confidence: 'low', source: 'strip' };
  try {
    var obj = JSON.parse(m[0]);
    var kind = String(obj.kind || 'single').trim().toLowerCase();
    var name = String(obj.name || '').trim();
    if (kind === 'mix') return { kind: 'mix', name: '', confidence: 'low', source: 'ai' };
    if (kind !== 'single') kind = name ? 'single' : 'mix';
    if (kind === 'mix') return { kind: 'mix', name: '', confidence: 'low', source: 'ai' };
    return { kind: 'single', name: name || fallback, confidence: 'high', source: 'ai' };
  } catch (e1) {
    return { kind: 'single', name: fallback, confidence: 'low', source: 'strip' };
  }
}

/**
 * 単品名・JAN パターン判定（Python single_name_classify と同型）。
 * @return {{pattern:string, single_trust:string, jan_trust:string, write:boolean, for_estimate:boolean, need_human:boolean}}
 */
function prSingleNameClassify_(inp) {
  var sale = String(inp.sale_form || '①');
  var qty = Number(inp.set_qty) || 1;
  var draft = String(inp.draft_name || '').trim();
  var existence = String(inp.existence || 'ok').toLowerCase();
  var aiConf = String(inp.ai_confidence || 'high').toLowerCase();
  var titleMatch = String(inp.title_match || 'high').toLowerCase();
  var hitAsin = String(inp.hit_asin || '').toUpperCase();
  var rowAsin = String(inp.row_asin || '').toUpperCase();
  var keepaJans = inp.keepa_jans || [];
  var searchJans = inp.search_jans || [];
  var nk = keepaJans.length;
  var ns = searchJans.length;

  function pack(pattern, singleTrust, janTrust, needHuman, write) {
    var w = write !== false && !!draft && singleTrust !== '要確認' && singleTrust !== '保留';
    return {
      pattern: pattern,
      single_trust: singleTrust,
      jan_trust: janTrust,
      write: w,
      for_estimate: w && singleTrust !== '要確認' && singleTrust !== '保留',
      need_human: !!needHuman,
    };
  }

  if (String(inp.kind || '') === 'mix' || sale.indexOf('③') === 0 || sale.indexOf('④') === 0 || sale.indexOf('詰め合わせ') >= 0) {
    return pack('P-MIX', '要確認', '触らない', true, false);
  }
  if (existence === 'unknown' || existence === '不明') {
    return pack('D2', '保留', '触らない', false, false);
  }
  if (existence === 'ng') {
    return pack('D1', '要確認', '触らない', true, false);
  }
  if (hitAsin && rowAsin && hitAsin !== rowAsin) {
    return pack('F3', '要確認', 'low', true, false);
  }
  if (sale.indexOf('①') === 0 && titleMatch === 'low') {
    return pack('F2', '要確認', 'low', true, false);
  }
  if (aiConf === 'low' || aiConf === '低い') {
    return pack('E1', '要確認', 'low', true, false);
  }

  var f1 = sale.indexOf('②') === 0 && qty > 1 && titleMatch === 'low';
  function finish(pattern, singleTrust, janTrust, needHuman) {
    if (f1) return pack('F1', 'mid', janTrust, needHuman, true);
    return pack(pattern, singleTrust, janTrust, needHuman, true);
  }

  if (nk === 1) {
    var kj = keepaJans[0];
    if (ns === 0) return finish('A3', 'high', 'mid', false);
    if (ns === 1 && searchJans[0] === kj) return finish('A1', 'high', 'high', false);
    if (searchJans.indexOf(kj) >= 0) return finish('A1', 'high', 'high', false);
    return finish('A2', 'mid', 'low', false);
  }
  if (nk === 0) {
    if (ns === 1) return finish('B1', 'high', '候補', true);
    if (ns > 1) return finish('B2', 'mid', '複数', true);
    return finish('B3', 'high', '欠測', false);
  }
  if (ns === 1) return finish('C1', 'high', '候補', false);
  return finish('C2', 'mid', '複数', true);
}

function prListEnsureSingleNameCols_(sh, headers) {
  var out = headers.slice();
  var need = [PR_LIST_COL_SINGLE_NAME_, PR_LIST_COL_SINGLE_TRUST_, PR_LIST_COL_SINGLE_PATTERN_];
  var i;
  for (i = 0; i < need.length; i++) {
    if (out.indexOf(need[i]) < 0) {
      out.push(need[i]);
      sh.getRange(PR_LIST_HEADER_ROW_, out.length).setValue(need[i]);
    }
  }
  return out;
}

/**
 * 見積候補かつ単品商品名が空の行を AI 充填（リスト上の未充填を順に消化）。
 * 既存の単品商品名は上書きしない。cap／時間予算で leftover。
 * @return {{ok:boolean, wrote:number, skip:number, human:number, leftover:number, processed:number, detail:string}}
 */
function prRunSingleNameFill_(opt) {
  opt = opt || {};
  var runId = opt.runId || 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  if (!PR_GAS_START_MS_) prGasStart_();
  prRefuseOriginalSs_();
  if (!prSingleNameEnabled_()) {
    throw new Error(PR_SINGLE_NAME_ENABLED_PROP_ + '=false（切戻し中）');
  }
  var sh = prSpreadsheet_().getSheetByName(PR_LIST_SHEET_);
  if (!sh) throw new Error('シートなし: ' + PR_LIST_SHEET_);
  var headers = prListHeaders_(sh);
  headers = prListEnsureSingleNameCols_(sh, headers);
  var idxAsin = headers.indexOf('ASIN');
  var idxTitle = headers.indexOf('商品名');
  var idxMaker = headers.indexOf('メーカー名');
  var idxJan = headers.indexOf('JANコード');
  var idxSale = headers.indexOf('販売形態');
  var idxQty = headers.indexOf('同商品のセット数量');
  var idxQuote = headers.indexOf(PR_LIST_COL_QUOTE_);
  var idxSingle = headers.indexOf(PR_LIST_COL_SINGLE_NAME_);
  var idxTrust = headers.indexOf(PR_LIST_COL_SINGLE_TRUST_);
  var idxPat = headers.indexOf(PR_LIST_COL_SINGLE_PATTERN_);
  if (idxAsin < 0 || idxTitle < 0 || idxSingle < 0) {
    throw new Error('ASIN/商品名/単品商品名 列なし');
  }
  var lastDataRow = prListLastAsinRow_(sh, idxAsin + 1);
  if (lastDataRow < PR_LIST_DATA_START_) throw new Error('ASINデータ行なし');
  var numRows = lastDataRow - PR_LIST_DATA_START_ + 1;
  var vals = sh.getRange(PR_LIST_DATA_START_, 1, numRows, headers.length).getValues();
  var cap = prSingleNameCap_();
  var wrote = 0;
  var skip = 0;
  var human = 0;
  var processed = 0;
  var leftover = 0;
  var i;
  var janCanon = {};
  for (i = 0; i < vals.length; i++) {
    var jans0 = prSingleNameParseJans_(idxJan >= 0 ? vals[i][idxJan] : '');
    var nm0 = String(vals[i][idxSingle] || '').trim();
    if (!prSingleNameLooksCanon_(nm0)) continue;
    var j0;
    for (j0 = 0; j0 < jans0.length; j0++) {
      var jn = jans0[j0];
      if (!janCanon[jn] || nm0.length < janCanon[jn].length) janCanon[jn] = nm0;
    }
  }
  for (i = 0; i < vals.length; i++) {
    var row = vals[i];
    var asin = String(row[idxAsin] || '').trim().toUpperCase();
    if (asin.length !== 10) continue;
    if (idxQuote >= 0 && String(row[idxQuote] || '').trim() !== '見積候補') {
      skip++;
      continue;
    }
    var existing = String(row[idxSingle] || '').trim();
    if (existing) {
      skip++;
      continue;
    }
    var prevTrust = idxTrust >= 0 ? String(row[idxTrust] || '').trim() : '';
    if (prevTrust === '要確認') {
      skip++;
      continue;
    }
    if (processed >= cap || prGasTimeUp_()) {
      leftover++;
      continue;
    }
    processed++;
    var title = idxTitle >= 0 ? String(row[idxTitle] || '') : '';
    var maker = idxMaker >= 0 ? String(row[idxMaker] || '').trim() : '';
    var sale = prSingleNameSaleCode_(idxSale >= 0 ? row[idxSale] : '');
    var qty = prSingleNameParseQty_(idxQty >= 0 ? row[idxQty] : 1);
    var keepaJans = prSingleNameParseJans_(idxJan >= 0 ? row[idxJan] : '');
    var reused = '';
    var j1;
    for (j1 = 0; j1 < keepaJans.length; j1++) {
      if (janCanon[keepaJans[j1]]) {
        reused = janCanon[keepaJans[j1]];
        break;
      }
    }
    var ai;
    if (reused) {
      ai = { kind: 'single', name: reused, confidence: 'high', source: 'jan_unify' };
    } else {
      ai = prSingleNameAiDraft_(title, maker, qty, sale);
      if (ai.source === 'ai') Utilities.sleep(400);
    }
    var tmatch = prSingleNameTitleMatch_(ai.name, title);
    var existence = ai.kind === 'mix' ? 'ok' : prSingleNameSoftExistence_(ai.name, title, maker);
    var cls = prSingleNameClassify_({
      draft_name: ai.name,
      kind: ai.kind,
      sale_form: sale,
      set_qty: qty,
      existence: existence,
      ai_confidence: ai.confidence,
      title_match: ai.kind === 'mix' ? 'high' : tmatch,
      row_asin: asin,
      hit_asin: existence === 'ok' ? asin : '',
      keepa_jans: keepaJans,
      search_jans: keepaJans.slice(),
    });
    Logger.log(
      '[仕入れ検討①] runId=%s single_name asin=%s pattern=%s kind=%s src=%s write=%s draft=%s',
      runId,
      asin,
      cls.pattern,
      ai.kind,
      ai.source,
      cls.write,
      String(ai.name).slice(0, 40)
    );
    if (cls.need_human) {
      human++;
      if (!cls.write) {
        row[idxSingle] = prSingleNameHumanLabel_(cls.pattern);
      }
    }
    if (cls.write) {
      row[idxSingle] = ai.name;
      wrote++;
      for (j1 = 0; j1 < keepaJans.length; j1++) {
        if (!janCanon[keepaJans[j1]] || ai.name.length < janCanon[keepaJans[j1]].length) {
          janCanon[keepaJans[j1]] = ai.name;
        }
      }
    }
    if (idxTrust >= 0) row[idxTrust] = cls.single_trust;
    if (idxPat >= 0) row[idxPat] = cls.pattern;
    vals[i] = row;
  }
  if (wrote > 0 || human > 0) {
    var names = [];
    var trusts = [];
    var pats = [];
    for (i = 0; i < vals.length; i++) {
      names.push([vals[i][idxSingle]]);
      if (idxTrust >= 0) trusts.push([vals[i][idxTrust]]);
      if (idxPat >= 0) pats.push([vals[i][idxPat]]);
    }
    sh.getRange(PR_LIST_DATA_START_, idxSingle + 1, names.length, 1).setValues(names);
    if (idxTrust >= 0) sh.getRange(PR_LIST_DATA_START_, idxTrust + 1, trusts.length, 1).setValues(trusts);
    if (idxPat >= 0) sh.getRange(PR_LIST_DATA_START_, idxPat + 1, pats.length, 1).setValues(pats);
  }
  // 既存の要確認空欄にも文言＋色（異種=赤白／その他=黄）
  prSingleNameApplyHumanStyles_(
    sh,
    vals,
    PR_LIST_DATA_START_,
    idxQuote,
    idxSingle,
    idxTrust,
    idxPat
  );
  var detail =
    'processed=' +
    processed +
    ' wrote=' +
    wrote +
    ' skip=' +
    skip +
    ' human=' +
    human +
    ' leftover=' +
    leftover +
    ' cap=' +
    cap +
    (opt.fromD2 ? ' via=D2' : '');
  if (!opt.skipLog) {
    prAppendLog_([runId, 'S 単品商品名AI充填', leftover > 0 ? 'DONE_LEFTOVER' : 'DONE', new Date(), '', wrote, '', detail]);
  }
  return {
    ok: true,
    wrote: wrote,
    skip: skip,
    human: human,
    leftover: leftover,
    processed: processed,
    detail: detail,
  };
}

function prMenuSingleNameFill() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=single_name state=RUNNING', runId);
  try {
    var r = prRunSingleNameFill_({ runId: runId });
    var msg = r.detail;
    if (r.leftover > 0) msg += ' → もう一度 S で続き';
    prMenuToast_('① 単品商品名', msg, 14);
  } catch (e) {
    var msg2 = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=single_name state=FAILED %s', runId, msg2);
    prAppendLog_([runId, 'S 単品商品名AI充填', 'FAILED', new Date(), '', '', '', msg2.slice(0, 200)]);
    prMenuFailAlert_(msg2);
    throw e;
  }
}

function prSellerAccumEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_SELLER_ACCUM_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prSellerAccumCap_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_SELLER_ACCUM_CAP_PROP_) ||
      PR_SELLER_ACCUM_CAP_DEFAULT_
  );
  if (!(n > 0)) n = PR_SELLER_ACCUM_CAP_DEFAULT_;
  if (n > PR_SELLER_ACCUM_CAP_MAX_) n = PR_SELLER_ACCUM_CAP_MAX_;
  return Math.floor(n);
}

function prSellerAccumSleepMs_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_SELLER_ACCUM_SLEEP_MS_PROP_) ||
      PR_SELLER_ACCUM_SLEEP_DEFAULT_MS_
  );
  if (!(n >= 500)) n = PR_SELLER_ACCUM_SLEEP_DEFAULT_MS_;
  if (n > 5000) n = 5000;
  return Math.floor(n);
}

function prSellerAccumBatchEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_SELLER_ACCUM_BATCH_PROP_);
  if (raw == null || String(raw).trim() === '') return false;
  var s = String(raw).trim().toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'ON' || s === 'はい';
}

function prSellerAccumLoopEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_SELLER_ACCUM_LOOP_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prSellerAccum429WaitMs_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_SELLER_ACCUM_429_WAIT_MS_PROP_) ||
      PR_SELLER_ACCUM_429_WAIT_DEFAULT_MS_
  );
  if (!(n >= 10000)) n = PR_SELLER_ACCUM_429_WAIT_DEFAULT_MS_;
  return Math.floor(n);
}

function prSellerAccum429RetryMax_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_SELLER_ACCUM_429_RETRY_MAX_PROP_) ||
      PR_SELLER_ACCUM_429_RETRY_DEFAULT_
  );
  if (!(n >= 0)) n = PR_SELLER_ACCUM_429_RETRY_DEFAULT_;
  return Math.floor(n);
}

function prSellerAccumOkSid_(sid) {
  var s = String(sid || '').trim();
  if (s.length < 10 || PR_SELLER_ACCUM_SKIP_IDS_[s]) return false;
  return true;
}

function prSellerAccumHave_(have, sid) {
  return !!have[String(sid || '').trim().toUpperCase()];
}

function prSellerAccumMark_(have, sid) {
  have[String(sid || '').trim().toUpperCase()] = true;
}

/** 門通過 ASIN を出現順（重複除く）。 */
function prCandPassAsinList_(opt) {
  opt = opt || {};
  var sh = prSpreadsheet_().getSheetByName(PR_CAND_SHEET_);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxAsin = prHeaderIndex_(hdr, 'ASIN');
  var idxGate = prHeaderIndex_(hdr, '門結果');
  if (idxAsin < 0 || idxGate < 0) return out;
  var idxSpapiDone = opt.skipSellerSpapiDone ? prHeaderIndex_(hdr, PR_CAND_SELLER_SPAPI_DONE_) : -1;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var seen = {};
  var i;
  for (i = 0; i < vals.length; i++) {
    if (String(vals[i][idxGate] || '') !== '通過') continue;
    if (idxSpapiDone >= 0 && prCkIsTrue_(vals[i][idxSpapiDone])) continue;
    var asin = String(vals[i][idxAsin] || '').trim().toUpperCase();
    if (asin.length !== 10 || seen[asin]) continue;
    seen[asin] = true;
    out.push(asin);
  }
  return out;
}

/** C0 用: 門通過のうち セラー調査済みSP-API でない ASIN。 */
function prCandPassAsinListForSellerAccum_() {
  return prCandPassAsinList_({ skipSellerSpapiDone: true });
}

/** ①候補に セラー調査済みSP-API 列を確保（チェック枠付き）。 */
function prCandEnsureSellerSpapiHeaders_(sh) {
  var need = [PR_CAND_SELLER_SPAPI_DONE_, PR_CAND_SELLER_SPAPI_AT_];
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var addedDone = false;
  var i;
  var c;
  for (i = 0; i < need.length; i++) {
    var found = false;
    for (c = 0; c < headers.length; c++) {
      if (String(headers[c]) === need[i]) found = true;
    }
    if (!found) {
      sh.getRange(1, headers.length + 1).setValue(need[i]);
      headers.push(need[i]);
      if (need[i] === PR_CAND_SELLER_SPAPI_DONE_) addedDone = true;
    }
  }
  if (addedDone) PR_CAND_MEM_ = null;
  var idxDone = prHeaderIndex_(headers, PR_CAND_SELLER_SPAPI_DONE_);
  var idxAt = prHeaderIndex_(headers, PR_CAND_SELLER_SPAPI_AT_);
  if (addedDone && idxDone >= 0 && sh.getLastRow() >= 2) {
    var numRows = sh.getLastRow() - 1;
    sh.getRange(2, idxDone + 1, numRows, 1).insertCheckboxes();
  }
  if (idxAt >= 0) {
    var trimCol;
    for (trimCol = headers.length - 1; trimCol > idxAt; trimCol--) {
      if (String(headers[trimCol] || '').trim() !== '') break;
      sh.deleteColumn(trimCol + 1);
      headers.pop();
    }
  }
  return {
    headers: headers,
    idxDone: idxDone,
    idxAt: idxAt,
    lastCol: headers.length,
  };
}

/** SP-API 完了後、①候補の当該 ASIN 行に セラー調査済みSP-API を付ける。 */
function prMarkCandSellerSpapiDone_(asin, opt) {
  opt = opt || {};
  var a = String(asin || '').trim().toUpperCase();
  if (a.length !== 10) return false;
  var sh = prSpreadsheet_().getSheetByName(PR_CAND_SHEET_);
  if (!sh || sh.getLastRow() < 2) return false;
  var meta = prCandEnsureSellerSpapiHeaders_(sh);
  var idxAsin = prHeaderIndex_(meta.headers, 'ASIN');
  if (idxAsin < 0 || meta.idxDone < 0) return false;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, meta.lastCol).getValues();
  var r;
  var hit = -1;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxAsin] || '').trim().toUpperCase() === a) {
      hit = r;
      break;
    }
  }
  if (hit < 0) return false;
  var rowNum = hit + 2;
  sh.getRange(rowNum, meta.idxDone + 1).setValue(true);
  if (meta.idxAt >= 0) {
    sh.getRange(rowNum, meta.idxAt + 1).setValue(opt.at || new Date());
  }
  PR_CAND_MEM_ = null;
  return true;
}

/** 門通過 ASIN だけ Keepaフル BB 列（新規 GET なし）。 */
function prKeepaFullBbMapForAsins_(asins) {
  var flat = prLoadKeepaFullFlatMapForAsins_(asins);
  var out = {};
  var i;
  for (i = 0; i < asins.length; i++) {
    var a = asins[i];
    var rec = flat[a];
    if (!rec) continue;
    var sid = String(rec['BuyBoxセラー'] || '').trim();
    if (prSellerAccumOkSid_(sid)) out[a] = sid;
  }
  return out;
}

function prSpapiOfferSellerIdOf_(o) {
  if (!o || typeof o !== 'object') return '';
  var s = o.SellerId || o.sellerId || '';
  if (s && typeof s === 'object') s = s.SellerId || s.sellerId || '';
  return String(s || '').trim();
}

/** SP-API getItemOffers（新品）。429/http 失敗は stopped。1 ASIN あたり最大20 sellerId。 */
function prSpapiItemOffersNew_(asin, opt) {
  opt = opt || {};
  var capPerAsin = opt.capPerAsin || 20;
  var endpoint = prSpapiEndpoint_();
  var mp = prSpapiMarketplace_();
  var token = prSpapiAccessToken_();
  var a = String(asin || '').trim().toUpperCase();
  var seen = [];
  var have = {};
  var pages = 0;
  var https = [];
  var nextToken = '';
  var err = '';
  var stopped = '';
  var maxPages = 20;
  while (seen.length < capPerAsin && pages < maxPages) {
    pages++;
    var q = 'MarketplaceId=' + encodeURIComponent(mp) + '&ItemCondition=New';
    if (nextToken) q += '&NextToken=' + encodeURIComponent(nextToken);
    var url =
      endpoint +
      '/products/pricing/v0/items/' +
      encodeURIComponent(a) +
      '/offers?' +
      q;
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        Accept: 'application/json',
        'x-amz-access-token': token,
        'User-Agent': 'PurchaseResearch/1.0 (GAS)',
      },
    });
    var code = resp.getResponseCode();
    https.push(code);
    if (code === 429) {
      stopped = '429';
      break;
    }
    var body = resp.getContentText() || '';
    if (code !== 200) {
      stopped = 'http';
      try {
        var ej = JSON.parse(body);
        err = JSON.stringify(ej.errors || ej).slice(0, 240);
      } catch (e1) {
        err = body.slice(0, 200);
      }
      break;
    }
    var data = JSON.parse(body);
    var payload = data.payload || data;
    var offers = payload.Offers || payload.offers || [];
    var oi;
    for (oi = 0; oi < offers.length; oi++) {
      var sid = prSpapiOfferSellerIdOf_(offers[oi]);
      if (sid && !have[sid]) {
        have[sid] = true;
        seen.push(sid);
        if (seen.length >= capPerAsin) break;
      }
    }
    var summary = payload.Summary || payload.summary || {};
    nextToken = payload.NextToken || payload.nextToken || summary.NextToken || '';
    if (!nextToken || !offers.length) break;
    Utilities.sleep(500);
  }
  return { http: https, ids: seen, err: err, stopped: stopped, pages: pages };
}

/** getItemOffersBatch 最大20 ASIN。 */
function prSpapiItemOffersBatch_(asins) {
  var endpoint = prSpapiEndpoint_();
  var mp = prSpapiMarketplace_();
  var token = prSpapiAccessToken_();
  var chunk = asins.slice(0, 20);
  var reqs = [];
  var i;
  for (i = 0; i < chunk.length; i++) {
    reqs.push({
      uri: '/products/pricing/v0/items/' + chunk[i] + '/offers',
      method: 'GET',
      MarketplaceId: mp,
      ItemCondition: 'New',
    });
  }
  var url = endpoint + '/batches/products/pricing/v0/itemOffers';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    followRedirects: true,
    payload: JSON.stringify({ requests: reqs }),
    headers: {
      Accept: 'application/json',
      'x-amz-access-token': token,
      'User-Agent': 'PurchaseResearch/1.0 (GAS)',
    },
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '';
  var by = {};
  var codes = [];
  var err = '';
  if (code === 429) {
    return { http: code, err: '429', by: by, codes: [429], stopped: '429', n: chunk.length };
  }
  var data;
  try {
    data = JSON.parse(body);
  } catch (e2) {
    return { http: code, err: body.slice(0, 200), by: by, codes: [], stopped: 'http', n: chunk.length };
  }
  if (code !== 200) {
    return {
      http: code,
      err: JSON.stringify(data.errors || data).slice(0, 240),
      by: by,
      codes: [],
      stopped: 'http',
      n: chunk.length,
    };
  }
  var rows = data.responses || [];
  for (i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var st = row.status && row.status.statusCode;
    codes.push(st);
    if (st === 429) {
      return { http: code, err: '429 in batch', by: by, codes: codes, stopped: '429', n: chunk.length };
    }
    var b = row.body || {};
    var payload = b.payload || b;
    var ident = payload.Identifier || payload.identifier || {};
    var a = String(ident.ASIN || ident.asin || payload.ASIN || '').trim().toUpperCase();
    var offers = payload.Offers || payload.offers || [];
    var ids = [];
    var oi;
    for (oi = 0; oi < offers.length; oi++) {
      var sid = prSpapiOfferSellerIdOf_(offers[oi]);
      if (sid) ids.push(sid);
    }
    if (a) by[a] = ids;
    if (st && Number(st) >= 400 && !err) err = String(b.errors || st).slice(0, 120);
  }
  return { http: code, err: err, by: by, codes: codes, stopped: '', n: chunk.length };
}

/** 競合DB セラーへ C0 行を末尾追加。調査CK は OFF。 */
function prAppendSellerAccumRows_(sellerIds) {
  if (!sellerIds || !sellerIds.length) return 0;
  var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_SELLER_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxId = prHeaderIndex_(hdr, 'sellerId');
  var idxSrc = prHeaderIndex_(hdr, '採取元');
  var idxMemo = prHeaderIndex_(hdr, 'メモ');
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  if (idxId < 0) throw new Error('セラー sellerId 列なし');
  var rows = [];
  var i;
  var j;
  for (i = 0; i < sellerIds.length; i++) {
    var line = [];
    for (j = 0; j < lastCol; j++) line.push('');
    if (idxCk >= 0) line[idxCk] = false;
    line[idxId] = sellerIds[i];
    if (idxSrc >= 0) line[idxSrc] = 'bb+spapi_new';
    if (idxMemo >= 0) line[idxMemo] = 'C0貯め 構成比は店洗い時';
    rows.push(line);
  }
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, lastCol).setValues(rows);
  if (idxCk >= 0) {
    sh.getRange(startRow, idxCk + 1, rows.length, 1).insertCheckboxes();
  }
  prClearSellerLedgerCache_();
  return rows.length;
}

/**
 * C0 セラー貯め本体。門通過 ASIN → BB 和集合 + SP-API 新品。新規100/回。
 * 429 時は PR_SELLER_ACCUM_AFTER_ASIN に last ASIN を保存して再開。
 */
function prRunSellerAccum_(opt) {
  opt = opt || {};
  prRefuseOriginalSs_();
  if (!prSellerAccumEnabled_()) {
    return { ok: true, skipped: true, newN: 0, detail: 'PR_SELLER_ACCUM_ENABLED=false' };
  }
  var props = PropertiesService.getScriptProperties();
  var cap = prSellerAccumCap_();
  if (opt.capLeft != null && opt.capLeft >= 0 && opt.capLeft < cap) cap = opt.capLeft;
  var sleepMs = prSellerAccumSleepMs_();
  var batchMode = prSellerAccumBatchEnabled_();
  var passAll = prCandPassAsinList_();
  var asins = prCandPassAsinListForSellerAccum_();
  var passPending = asins.length;
  prLoadSellerLedger_();
  var have = {};
  var ek;
  for (ek in PR_SELLER_ROWS_CACHE_) {
    if (Object.prototype.hasOwnProperty.call(PR_SELLER_ROWS_CACHE_, ek)) have[ek] = true;
  }
  var sellerN = Object.keys(have).length;
  if (!asins.length) {
    return {
      ok: true,
      newN: 0,
      detail:
        'pass=' +
        passAll.length +
        ' pending=' +
        passPending +
        ' seller_n=' +
        sellerN +
        ' remain_asin=0 KeepaGETなし',
    };
  }
  var bb = prKeepaFullBbMapForAsins_(asins);
  var newIds = [];
  var nApi = 0;
  var nMarked = 0;
  var stopped = '';
  var lastA = '';
  var runId = String(opt.runId || '');
  var ai;
  if (batchMode) {
    var bi = 0;
    while (bi < asins.length && newIds.length < cap) {
      if (prGasTimeUp_()) {
        stopped = 'time';
        break;
      }
      var chunk = asins.slice(bi, bi + 20);
      bi += 20;
      Utilities.sleep(Math.max(2000, sleepMs));
      nApi++;
      lastA = chunk[chunk.length - 1];
      var br = prSpapiItemOffersBatch_(chunk);
      Logger.log(
        '[仕入れ検討①] C0 batch http=%s n=%s codes=%s err=%s',
        br.http,
        br.n,
        (br.codes || []).join(','),
        br.err || ''
      );
      if (br.stopped === '429') {
        stopped = '429';
        props.setProperty(PR_SELLER_ACCUM_AFTER_ASIN_PROP_, lastA);
        break;
      }
      if (br.stopped === 'http') {
        stopped = 'http';
        break;
      }
      var ci;
      for (ci = 0; ci < chunk.length; ci++) {
        var ca = chunk[ci];
        var cids = br.by && br.by[ca] ? br.by[ca].slice() : [];
        if (bb[ca]) cids.unshift(bb[ca]);
        var cj;
        for (cj = 0; cj < cids.length; cj++) {
          var cs = cids[cj];
          if (!prSellerAccumOkSid_(cs) || prSellerAccumHave_(have, cs)) continue;
          prSellerAccumMark_(have, cs);
          newIds.push(cs);
          if (newIds.length >= cap) {
            stopped = 'cap';
            break;
          }
        }
        if (stopped) break;
      }
      var cm;
      for (cm = 0; cm < chunk.length; cm++) {
        if (prMarkCandSellerSpapiDone_(chunk[cm])) nMarked++;
      }
      if (stopped) break;
    }
  } else {
    for (ai = 0; ai < asins.length; ai++) {
      if (newIds.length >= cap) {
        stopped = 'cap';
        break;
      }
      if (prGasTimeUp_()) {
        stopped = 'time';
        break;
      }
      var a = asins[ai];
      lastA = a;
      var ids = [];
      if (bb[a]) ids.push(bb[a]);
      Utilities.sleep(sleepMs);
      nApi++;
      var r = prSpapiItemOffersNew_(a, { capPerAsin: 20 });
      if (r.stopped === '429') {
        stopped = '429';
        props.setProperty(PR_SELLER_ACCUM_AFTER_ASIN_PROP_, a);
        Logger.log('[仕入れ検討①] C0 STOP 429 after api=%s asin=%s', nApi, a);
        break;
      }
      if (r.stopped === 'http') {
        stopped = 'http';
        Logger.log(
          '[仕入れ検討①] C0 STOP http asin=%s http=%s err=%s',
          a,
          (r.http || []).join(','),
          r.err
        );
        break;
      }
      ids = ids.concat(r.ids || []);
      var si;
      for (si = 0; si < ids.length; si++) {
        var s = ids[si];
        if (!prSellerAccumOkSid_(s) || prSellerAccumHave_(have, s)) continue;
        prSellerAccumMark_(have, s);
        newIds.push(s);
        if (newIds.length >= cap) {
          stopped = 'cap';
          break;
        }
      }
      if (prMarkCandSellerSpapiDone_(a)) nMarked++;
    }
  }
  var remainPending = Math.max(0, passPending - nMarked);
  if (stopped === '429' || stopped === 'time') {
    if (lastA) props.setProperty(PR_SELLER_ACCUM_AFTER_ASIN_PROP_, lastA);
  } else {
    props.deleteProperty(PR_SELLER_ACCUM_AFTER_ASIN_PROP_);
  }
  var hit = 0;
  if (newIds.length) hit = prAppendSellerAccumRows_(newIds);
  var detail =
    'pass=' +
    passAll.length +
    ' pending=' +
    passPending +
    ' marked=' +
    nMarked +
    ' remain_asin=' +
    remainPending +
    ' seller_n=' +
    sellerN +
    ' new=' +
    hit +
    ' api=' +
    nApi +
    ' last=' +
    (lastA || '-') +
    ' stop=' +
    (stopped || 'ok') +
    ' sleep_ms=' +
    sleepMs +
    ' batch=' +
    (batchMode ? '1' : '0') +
    ' KeepaGETなし';
  var ok = stopped !== 'http' && hit === newIds.length;
  return { ok: ok, newN: hit, api: nApi, stopped: stopped, lastAsin: lastA, detail: detail };
}

function prMenuSellerAccum() {
  prCyclePauseForOtherMenu_('C0');
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=C0 state=RUNNING functionName=prMenuSellerAccum', runId);
  try {
    var capTotal = prSellerAccumCap_();
    var loopOn = prSellerAccumLoopEnabled_();
    var wait429 = prSellerAccum429WaitMs_();
    var max429 = prSellerAccum429RetryMax_();
    var totalNew = 0;
    var rounds = 0;
    var retry429 = 0;
    var lastR = null;
    var lastDetail = '';
    while (totalNew < capTotal) {
      if (prGasTimeUp_()) {
        lastDetail = (lastDetail || '') + ' stop=time_outer';
        break;
      }
      if (!prCandPassAsinListForSellerAccum_().length) break;
      rounds++;
      lastR = prRunSellerAccum_({ runId: runId, capLeft: capTotal - totalNew });
      if (lastR.skipped) {
        lastDetail = lastR.detail || '';
        prAppendLog_([runId, 'C0 セラー貯め', 'DONE', new Date(), '', 0, '', lastDetail]);
        prMenuToast_('C0 セラー貯め', lastDetail, 8);
        return;
      }
      totalNew += lastR.newN || 0;
      lastDetail = lastR.detail || '';
      if (lastR.stopped === 'http') break;
      if (lastR.stopped === 'time') break;
      if (lastR.stopped === '429') {
        if (!loopOn || retry429 >= max429 || prGasTimeUp_()) break;
        retry429++;
        Logger.log(
          '[仕入れ検討①] C0 429 wait=%sms retry=%s/%s runId=%s',
          wait429,
          retry429,
          max429,
          runId
        );
        Utilities.sleep(wait429);
        continue;
      }
      retry429 = 0;
      if (totalNew >= capTotal) break;
      if (!prCandPassAsinListForSellerAccum_().length) break;
      if (!loopOn) break;
    }
    var state = 'DONE';
    if (lastR && (lastR.stopped === '429' || lastR.stopped === 'time')) state = 'RETRYING';
    else if (lastR && lastR.stopped === 'http') state = 'FAILED';
    else if (lastR && !lastR.ok && lastR.stopped !== 'cap') state = 'FAILED';
    var detail =
      (lastDetail || '') +
      ' loop=' +
      (loopOn ? '1' : '0') +
      ' rounds=' +
      rounds +
      ' total_new=' +
      totalNew +
      ' retry429=' +
      retry429;
    prAppendLog_([runId, 'C0 セラー貯め', state, new Date(), '', totalNew, '', detail]);
    Logger.log('[仕入れ検討①] runId=%s step=C0 state=%s %s', runId, state, detail);
    prMenuToast_('C0 セラー貯め', detail);
    if (state === 'FAILED') prMenuFailAlert_(detail);
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=C0 state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'C0 セラー貯め', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

function prAlightEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_ALIGHT_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prAlightShopCap_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_ALIGHT_SHOP_CAP_PROP_) ||
      PR_ALIGHT_SHOP_CAP_DEFAULT_
  );
  if (!(n > 0)) n = PR_ALIGHT_SHOP_CAP_DEFAULT_;
  if (n > 10) n = 10;
  return Math.floor(n);
}

function prAlightProductCap_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_ALIGHT_PRODUCT_CAP_PROP_) ||
      PR_ALIGHT_PRODUCT_CAP_DEFAULT_
  );
  if (!(n > 0)) n = PR_ALIGHT_PRODUCT_CAP_DEFAULT_;
  if (n > PR_ALIGHT_PRODUCT_CAP_MAX_) n = PR_ALIGHT_PRODUCT_CAP_MAX_;
  return Math.floor(n);
}

function prAlightDailyTarget_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_ALIGHT_DAILY_TARGET_PROP_) ||
      PR_ALIGHT_DAILY_TARGET_DEFAULT_
  );
  if (!(n > 0)) n = PR_ALIGHT_DAILY_TARGET_DEFAULT_;
  return Math.floor(n);
}

function prAlightDailyMax_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_ALIGHT_DAILY_MAX_PROP_) ||
      PR_ALIGHT_DAILY_MAX_DEFAULT_
  );
  if (!(n > 0)) n = PR_ALIGHT_DAILY_MAX_DEFAULT_;
  return Math.floor(n);
}

function prAlightNoMakerRoundsMax_() {
  var n = Number(
    PropertiesService.getScriptProperties().getProperty(PR_ALIGHT_NO_MAKER_ROUNDS_MAX_PROP_) ||
      PR_ALIGHT_NO_MAKER_ROUNDS_DEFAULT_
  );
  if (!(n > 0)) n = PR_ALIGHT_NO_MAKER_ROUNDS_DEFAULT_;
  return Math.floor(n);
}

function prSellerBandFromSkuN_(n) {
  n = Number(n) || 0;
  if (n <= 0) return PR_BAND_UNK_;
  if (n <= PR_SELLER_BAND_S_MAX_) return PR_BAND_S_;
  if (n <= PR_SELLER_BAND_M_MAX_) return PR_BAND_M_;
  if (n <= PR_SELLER_BAND_L_MAX_) return PR_BAND_L_;
  return PR_BAND_XL_;
}

function prSellerBandRank_(band) {
  var b = String(band || '');
  if (b === PR_BAND_S_) return 1;
  if (b === PR_BAND_M_) return 2;
  if (b === PR_BAND_L_) return 3;
  if (b === PR_BAND_XL_) return 4;
  return 0;
}

/**
 * ①プロファイルの調査カテゴリー=TRUE 行から AmazonルートID を返す（CYCLE §4.1bis）。
 * @return {{id:string, label:string}[]}
 */
function prSurveyRootCategories_() {
  var profs = prReadAllProfiles_();
  var out = [];
  var seen = {};
  var i;
  for (i = 0; i < profs.length; i++) {
    var m = profs[i];
    var cat = String(m['カテゴリ'] || '').trim();
    if (cat === '共通') continue;
    if (/^\d+$/.test(cat)) continue;
    if (!prCkIsTrue_(m[PR_PROF_SURVEY_CK_HEADER_])) continue;
    var id = String(m[PR_PROF_ROOT_ID_HEADER_] || '')
      .trim()
      .replace(/\.0$/, '');
    if (!id || seen[id]) continue;
    seen[id] = true;
    out.push({ id: id, label: cat || id });
  }
  if (!out.length) {
    Logger.log('[仕入れ検討①] survey roots empty → fallback food %s', PR_FOOD_ROOT_);
    out.push({ id: PR_FOOD_ROOT_, label: '食品・飲料・お酒' });
  }
  return out;
}

function prSurveyRootCategoryIds_() {
  return prSurveyRootCategories_().map(function (c) {
    return c.id;
  });
}

/**
 * レ点カテゴリごとに query page0 → SKU見積・帯。代表帯=max。hold=いずれか≥501。
 * @return {{
 *   parts:Array, skuEst:number, band:string, hold:boolean,
 *   skuBreakdown:string, bandBreakdown:string, tokensConsumed:number,
 *   walkRoots:string[], page0Asins:string[]|null, page0Total:number
 * }}
 */
function prMeasureSellerSkuByCategory_(sellerId) {
  var sid = String(sellerId || '').trim();
  var cats = prSurveyRootCategories_();
  var parts = [];
  var tokens = 0;
  var skuEst = 0;
  var band = PR_BAND_UNK_;
  var hold = false;
  var i;
  for (i = 0; i < cats.length; i++) {
    var q = prKeepaQueryGet_(sid, cats[i].id, 0, 100);
    tokens += Number(q.tokensConsumed) || 0;
    var n = Number(q.totalResults || 0);
    var b = prSellerBandFromSkuN_(n);
    parts.push({
      id: cats[i].id,
      label: cats[i].label,
      n: n,
      band: b,
      asinList: (q.asinList || []).slice(),
    });
    if (n > skuEst) skuEst = n;
    if (prSellerBandRank_(b) > prSellerBandRank_(band)) band = b;
    if (n >= PR_QUERY_TOTAL_HOLD_) hold = true;
  }
  var skuParts = [];
  var bandParts = [];
  for (i = 0; i < parts.length; i++) {
    skuParts.push(parts[i].label + ':' + parts[i].n);
    bandParts.push(parts[i].label + ':' + parts[i].band);
  }
  var walkRoots = cats.map(function (c) {
    return c.id;
  });
  var page0Asins = null;
  var page0Total = skuEst;
  if (parts.length === 1) {
    page0Asins = parts[0].asinList.slice();
    page0Total = parts[0].n;
  }
  return {
    parts: parts,
    skuEst: skuEst,
    band: band,
    hold: hold || band === PR_BAND_XL_,
    skuBreakdown: skuParts.join('|'),
    bandBreakdown: bandParts.join('|'),
    tokensConsumed: tokens,
    walkRoots: walkRoots,
    page0Asins: page0Asins,
    page0Total: page0Total,
  };
}

function prAlightDailyDoneCount_() {
  var props = PropertiesService.getScriptProperties();
  var day = prYmdTokyo_();
  if (String(props.getProperty(PR_ALIGHT_DAILY_DONE_DAY_PROP_) || '') !== day) return 0;
  return Number(props.getProperty(PR_ALIGHT_DAILY_DONE_COUNT_PROP_) || 0) || 0;
}

function prAlightDailyDoneBump_(n) {
  n = Number(n) || 0;
  if (n <= 0) return;
  var props = PropertiesService.getScriptProperties();
  var day = prYmdTokyo_();
  if (String(props.getProperty(PR_ALIGHT_DAILY_DONE_DAY_PROP_) || '') !== day) {
    props.setProperty(PR_ALIGHT_DAILY_DONE_DAY_PROP_, day);
    props.setProperty(PR_ALIGHT_DAILY_DONE_COUNT_PROP_, String(n));
    return;
  }
  var cur = Number(props.getProperty(PR_ALIGHT_DAILY_DONE_COUNT_PROP_) || 0) || 0;
  props.setProperty(PR_ALIGHT_DAILY_DONE_COUNT_PROP_, String(cur + n));
}

function prAlightRoundCount_() {
  var props = PropertiesService.getScriptProperties();
  var day = prYmdTokyo_();
  if (String(props.getProperty(PR_ALIGHT_ROUND_DAY_PROP_) || '') !== day) return 0;
  return Number(props.getProperty(PR_ALIGHT_ROUND_COUNT_PROP_) || 0) || 0;
}

function prAlightRoundBump_() {
  var props = PropertiesService.getScriptProperties();
  var day = prYmdTokyo_();
  if (String(props.getProperty(PR_ALIGHT_ROUND_DAY_PROP_) || '') !== day) {
    props.setProperty(PR_ALIGHT_ROUND_DAY_PROP_, day);
    props.setProperty(PR_ALIGHT_ROUND_COUNT_PROP_, '1');
    return 1;
  }
  var cur = Number(props.getProperty(PR_ALIGHT_ROUND_COUNT_PROP_) || 0) || 0;
  cur += 1;
  props.setProperty(PR_ALIGHT_ROUND_COUNT_PROP_, String(cur));
  return cur;
}

function prAlightSetLastNewMakers_(n) {
  PropertiesService.getScriptProperties().setProperty(
    PR_ALIGHT_LAST_NEW_PROP_,
    String(Number(n) || 0)
  );
}

function prAlightLastNewMakers_() {
  return Number(PropertiesService.getScriptProperties().getProperty(PR_ALIGHT_LAST_NEW_PROP_) || -1);
}

function prAlightContGet_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PR_ALIGHT_CONT_PROP_) || '';
    if (!raw) return null;
    var o = JSON.parse(raw);
    if (!o || !o.sid) return null;
    return o;
  } catch (e) {
    return null;
  }
}

function prAlightContSet_(obj) {
  if (!obj) {
    PropertiesService.getScriptProperties().deleteProperty(PR_ALIGHT_CONT_PROP_);
    return;
  }
  PropertiesService.getScriptProperties().setProperty(
    PR_ALIGHT_CONT_PROP_,
    JSON.stringify(obj).slice(0, 8500)
  );
}

function prAlightHasContinue_() {
  return !!prAlightContGet_();
}

/**
 * 小中=0のとき大の前に店ラウンドがまだ必要か（A-light 日次／追加round）。
 * 自動レ点の可否はここでは見ない（prCycleShopBeforeLargeAction_ が先に評価）。
 * @return {boolean}
 */
function prAlightShouldShopBeforeLarge_() {
  if (prAlightHasContinue_()) return true;
  var done = prAlightDailyDoneCount_();
  var maxD = prAlightDailyMax_();
  if (done >= maxD) return false;
  var target = prAlightDailyTarget_();
  var rounds = prAlightRoundCount_();
  var roundMax = prAlightNoMakerRoundsMax_();
  // 未設定(-1)は「新メーカー0」扱い＝追加ラウンド継続可（旧バグ: -1 のまま即終了）
  var lastNew = prAlightLastNewMakers_();
  if (lastNew < 0) lastNew = 0;
  if (done < target) return true;
  if (lastNew === 0 && rounds < roundMax) return true;
  return false;
}

/**
 * sm=0 のとき、大/特大 Keepa続きより店レーンを先に（§4.4・A-1再開）。
 * 自動レ点候補が残っていれば必ず true（店ラウンド尽きた判定より優先）。
 * 進捗 left= は残す。
 * @param {string} parent
 * @return {boolean}
 */
function prCycleDeferLargeContinueForShop_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return false;
  var band = prMakerBandFromSum_(parent);
  if (band !== PR_BAND_L_ && band !== PR_BAND_XL_) return false;
  try {
    var pools = prCycleMakerBandPools_();
    var smN = pools.byBand['小'].length + pools.byBand['中'].length;
    if (smN > 0) return false;
  } catch (e) {
    return false;
  }
  try {
    return !!prCycleShopBeforeLargeAction_({});
  } catch (e2) {
    return true;
  }
}

/** A-light 用 /product GET（stats=0・offers なし）。Keepaフル upsert しない。 */
function prKeepaProductMakerFetch_(asins) {
  var out = [];
  var tokensConsumed = 0;
  var tokensLeft = '';
  var refillIn = '';
  var rateLimited = false;
  var i;
  for (i = 0; i < asins.length; i += 50) {
    var chunk = asins.slice(i, i + 50);
    var url =
      PR_KEEPA_PRODUCT_ +
      '?key=' +
      encodeURIComponent(prKeepaKey_()) +
      '&domain=5&asin=' +
      encodeURIComponent(chunk.join(',')) +
      '&history=0&stats=0';
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { Accept: 'application/json' },
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText() || '';
    if (code !== 200) {
      if (code === 429) {
        rateLimited = true;
        try {
          var data429 = JSON.parse(body);
          if (data429.tokensLeft != null) tokensLeft = data429.tokensLeft;
          if (data429.refillIn != null) refillIn = data429.refillIn;
        } catch (e429) {}
        Logger.log('[仕入れ検討①] A-light product RATE_LIMIT left=%s', tokensLeft);
        break;
      }
      throw new Error('Keepa A-light /product HTTP ' + code + ' ' + prRedactKey_(body).slice(0, 180));
    }
    var data = JSON.parse(body);
    var cPart = Number(data.tokensConsumed);
    if (cPart > 0) tokensConsumed += cPart;
    if (data.tokensLeft != null) tokensLeft = data.tokensLeft;
    if (data.refillIn != null) refillIn = data.refillIn;
    var products = data.products || [];
    var j;
    for (j = 0; j < products.length; j++) out.push(products[j]);
  }
  return {
    products: out,
    tokensConsumed: tokensConsumed,
    tokensLeft: tokensLeft,
    refillIn: refillIn,
    rateLimited: rateLimited,
  };
}

/** Keepaフルキャッシュ＋軽量 GET。prAppendKeepaFull_ は呼ばない。 */
function prKeepaProductMakerLight_(asins, productCap) {
  var list = (asins || []).slice(0, productCap);
  var packed = prSplitKeepaFull_(list);
  Logger.log('[仕入れ検討①] A-light product hit=%s need_get=%s cap=%s', packed.have.length, packed.need.length, productCap);
  var fetchedPack = packed.need.length
    ? prKeepaProductMakerFetch_(packed.need)
    : { products: [], tokensConsumed: 0, rateLimited: false, refillIn: '' };
  return {
    products: packed.have.concat(fetchedPack.products || []),
    tokensConsumed: fetchedPack.tokensConsumed || 0,
    rateLimited: !!fetchedPack.rateLimited,
    refillIn: fetchedPack.refillIn || '',
    need: packed.need.length,
    skip: packed.have.length,
  };
}

/** セラー全行から セラーヒット を再計算（+1 加算禁止）。 */
function prRecomputeSellerHits_(sh) {
  var hits = {};
  if (!sh || sh.getLastRow() < 2) return hits;
  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxId = prHeaderIndex_(hdr, 'sellerId');
  var idxExtract = prHeaderIndex_(hdr, '抽出メーカー');
  if (idxId < 0 || idxExtract < 0) return hits;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    var sid = String(vals[r][idxId] || '').trim();
    if (!sid) continue;
    var seenInShop = {};
    var makers = prMakersInCell_(vals[r][idxExtract]);
    var mi;
    for (mi = 0; mi < makers.length; mi++) {
      var parent = makers[mi];
      if (!prHitEligibleMaker_(parent)) continue;
      if (seenInShop[parent]) continue;
      seenInShop[parent] = true;
      hits[parent] = (hits[parent] || 0) + 1;
    }
  }
  return hits;
}

/** メーカーサマリ.セラーヒット を全店再走査結果で上書き。 */
function prSyncMakerSumHits_(hits) {
  hits = hits || {};
  var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
  prEnsureMakerSumHitCol_(sh);
  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  var idxHit = prHeaderIndex_(hdr, PR_MAKER_SUM_HIT_COL_);
  if (idxMaker < 0 || idxHit < 0) throw new Error('メーカーサマリ列不足（メーカー/セラーヒット）');
  var k;
  for (k in hits) {
    if (!Object.prototype.hasOwnProperty.call(hits, k)) continue;
    if (!prHitEligibleMaker_(k)) continue;
    prEnsureMakerSumRowTop_(sh, k, {});
  }
  lastCol = sh.getLastColumn();
  hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  idxHit = prHeaderIndex_(hdr, PR_MAKER_SUM_HIT_COL_);
  idxMaker = prHeaderIndex_(hdr, 'メーカー');
  var nSet = 0;
  if (sh.getLastRow() >= 2) {
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      var name = String(vals[r][idxMaker] || '').trim();
      if (!name) continue;
      var count = hits[name] != null ? hits[name] : 0;
      sh.getRange(r + 2, idxHit + 1).setValue(count);
      if (count > 0) prFillMakerSumJittaIfEmpty_(sh, r + 2, name);
      nSet++;
    }
  }
  Logger.log('[仕入れ検討①] セラーヒット sync rows=%s hit_makers=%s', nSet, Object.keys(hits).length);
  return Object.keys(hits).length;
}

function prJittaResolveEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_JITTA_RESOLVE_ENABLED_PROP_);
  if (raw == null || raw === '') return true;
  return String(raw).toLowerCase() !== 'false';
}

/** 実態AI実行主体。未設定=hb（①心拍）。 */
function prJittaOwner_() {
  var o = String(
    PropertiesService.getScriptProperties().getProperty(PR_JITTA_OWNER_PROP_) || ''
  )
    .trim()
    .toLowerCase();
  if (o === PR_JITTA_OWNER_COMP_) return PR_JITTA_OWNER_COMP_;
  return PR_JITTA_OWNER_HB_;
}

function prJittaOwnerIsComp_() {
  return prJittaOwner_() === PR_JITTA_OWNER_COMP_;
}

function prJittaResolveCap_() {
  var n = parseInt(
    PropertiesService.getScriptProperties().getProperty(PR_JITTA_RESOLVE_CAP_PROP_) ||
      PR_JITTA_RESOLVE_CAP_DEFAULT_,
    10
  );
  if (!(n > 0)) n = PR_JITTA_RESOLVE_CAP_DEFAULT_;
  if (n > PR_JITTA_RESOLVE_CAP_MAX_) n = PR_JITTA_RESOLVE_CAP_MAX_;
  return n;
}

/** 心拍レーンで実態AI（店と分離）。省略=ON。OWNER=comp のときは OFF（競合側に委譲）。 */
function prJittaHbLaneEnabled_() {
  if (prJittaOwnerIsComp_()) return false;
  var raw = PropertiesService.getScriptProperties().getProperty(PR_JITTA_HB_LANE_ENABLED_PROP_);
  if (raw == null || raw === '') return true;
  return String(raw).toLowerCase() !== 'false';
}

function prJittaHbCap_() {
  var n = parseInt(
    PropertiesService.getScriptProperties().getProperty(PR_JITTA_HB_CAP_PROP_) ||
      PR_JITTA_HB_CAP_DEFAULT_,
    10
  );
  if (!(n > 0)) n = PR_JITTA_HB_CAP_DEFAULT_;
  if (n > PR_JITTA_HB_CAP_MAX_) n = PR_JITTA_HB_CAP_MAX_;
  return n;
}

function prJittaBacklogGate_() {
  var n = parseInt(
    PropertiesService.getScriptProperties().getProperty(PR_JITTA_BACKLOG_GATE_PROP_) ||
      PR_JITTA_BACKLOG_GATE_DEFAULT_,
    10
  );
  if (!(n > 0)) n = PR_JITTA_BACKLOG_GATE_DEFAULT_;
  return n;
}

/**
 * A-light 末尾の実態AI。明示 Prop があればそれに従う。
 * OWNER=comp または未設定時レーンONなら OFF。レーンOFFなら ON（旧挙動）。
 */
function prJittaAfterAlightEnabled_() {
  if (prJittaOwnerIsComp_()) return false;
  var raw = PropertiesService.getScriptProperties().getProperty(PR_JITTA_AFTER_ALIGHT_PROP_);
  if (raw != null && String(raw).trim() !== '') {
    return String(raw).toLowerCase() !== 'false';
  }
  return !prJittaHbLaneEnabled_();
}

/** メーカーサマリの未調査（AI前）件数。1回の読取。 */
function prCountMakerJittaUninvestigated_() {
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var n = 0;
  var r;
  for (r = 0; r < vals.length; r++) {
    if (prMakerSummaryRowUninvestigated_(vals[r], hdr)) n++;
  }
  return n;
}

/** 未調査がゲート以上なら自動レ点を止める（山を増やさない）。OWNER=comp でも件数ゲートは維持。 */
function prJittaBacklogBlocksAutoCk_() {
  if (prJittaOwnerIsComp_()) {
    return prCountMakerJittaUninvestigated_() >= prJittaBacklogGate_();
  }
  if (!prJittaHbLaneEnabled_()) return false;
  if (!prJittaResolveEnabled_()) return false;
  return prCountMakerJittaUninvestigated_() >= prJittaBacklogGate_();
}

/** 自動レ点候補数。バックログ中は 0。 */
function prEffectiveAutoCkCandidates_(rawCand) {
  var n = Number(rawCand) || 0;
  if (n <= 0) return 0;
  if (prJittaBacklogBlocksAutoCk_()) return 0;
  return n;
}

/** サマリで AI 確定済み（実態信頼度=high）。シードの実態だけでは確定扱いにしない。 */
function prMakerJittaSummaryAiConfirmed_(sh, sheetRow, hdr) {
  if (!sh || !(sheetRow >= 2)) return false;
  hdr = hdr || sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  var idxConf = prHeaderIndex_(hdr, '実態信頼度');
  if (idxConf < 0) return false;
  return String(sh.getRange(sheetRow, idxConf + 1).getValue() || '').trim() === 'high';
}

function prMakerMasterJittaCached_(master) {
  if (!master) return false;
  return (
    !!String(master['実態候補'] || '').trim() && !!String(master['実態信頼度'] || '').trim()
  );
}

/** 実態 AI 調査がまだ必要か。スキップ対象は false。 */
function prMakerNeedsJittaResolve_(maker) {
  maker = String(maker || '').trim();
  if (!maker || !prHitEligibleMaker_(maker)) return false;
  var master = prLoadMakerMasterJittaRow_(maker);
  if (master && prCkIsTrue_(master['実態再採取しない'])) return false;
  // マスタキャッシュがあっても同一キーの未調査サマリは heal 対象
  var sumSh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sumSh) return true;
  var sumRow = 0;
  var lastCol = Math.max(sumSh.getLastColumn(), 1);
  var hdr = sumSh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  if (idxMaker >= 0 && sumSh.getLastRow() >= 2) {
    var vals = sumSh.getRange(2, 1, sumSh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      if (String(vals[r][idxMaker] || '').trim() === maker) {
        sumRow = r + 2;
        break;
      }
    }
  }
  if (sumRow && prMakerJittaSummaryAiConfirmed_(sumSh, sumRow, hdr)) return false;
  return true;
}

/**
 * 未調査メーカー名リスト。hits があればセラーヒット降順、なければサマリのセラーヒット列。
 * 法人格ゆれは同一キーにまとめ、正名候補を1件だけ返す。
 */
function prCollectJittaResolveTargets_(hits) {
  hits = hits || {};
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return [];
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  var idxHit = prHeaderIndex_(hdr, PR_MAKER_SUM_HIT_COL_);
  if (idxMaker < 0) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var byKey = {};
  var r;
  for (r = 0; r < vals.length; r++) {
    var name = String(vals[r][idxMaker] || '').trim();
    if (!name) continue;
    var hitN = hits[name] != null ? Number(hits[name]) || 0 : Number(vals[r][idxHit]) || 0;
    if (!prMakerNeedsJittaResolve_(name)) continue;
    if (!prMakerSummaryRowUninvestigated_(vals[r], hdr)) continue;
    var key = prNormalizeMakerKey_(name) || name;
    var prev = byKey[key];
    if (
      !prev ||
      hitN > prev.hit ||
      (hitN === prev.hit && prMakerCanonicalScore_(name) > prMakerCanonicalScore_(prev.maker))
    ) {
      byKey[key] = { maker: name, hit: hitN };
    }
  }
  var rows = [];
  var k;
  for (k in byKey) {
    if (Object.prototype.hasOwnProperty.call(byKey, k)) rows.push(byKey[k]);
  }
  rows.sort(function (a, b) {
    return b.hit - a.hit;
  });
  var out = [];
  for (r = 0; r < rows.length; r++) out.push(rows[r].maker);
  return out;
}

/**
 * 未調査メーカーを AI 調査。スキップは cap に数えない。実 API 呼び出しのみカウント。
 * @return {{resolved:number, skipped:number, autoN:number, cap:number}}
 */
function prResolveJittaBatch_(makers, opt) {
  opt = opt || {};
  var cap = opt.cap != null ? opt.cap : prJittaResolveCap_();
  var resolved = 0;
  var skipped = 0;
  var autoN = 0;
  var i;
  for (i = 0; i < makers.length && resolved < cap; i++) {
    if (prGasTimeUp_()) break;
    var maker = makers[i];
    if (!prMakerNeedsJittaResolve_(maker)) {
      skipped++;
      continue;
    }
    var res = prResolveMakerJitta_(maker, opt);
    if (res.skipped) {
      skipped++;
      continue;
    }
    resolved++;
    if (res.autoConfirm) autoN++;
  }
  if (resolved || skipped) {
    Logger.log(
      '[仕入れ検討①] jitta_resolve_batch resolved=%s skipped=%s auto=%s cap=%s',
      resolved,
      skipped,
      autoN,
      cap
    );
  }
  return { resolved: resolved, skipped: skipped, autoN: autoN, cap: cap };
}

function prAmazonRootCatSet_() {
  var set = {};
  var i;
  for (i = 0; i < PR_AMAZON_ROOT_CATS_.length; i++) set[PR_AMAZON_ROOT_CATS_[i]] = true;
  return set;
}

function prEnsureSheetCols_(sh, colNames) {
  if (!sh) return {};
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = {};
  var i;
  var c;
  for (i = 0; i < colNames.length; i++) {
    c = colNames[i];
    var j = prHeaderIndex_(hdr, c);
    if (j < 0) {
      lastCol++;
      sh.getRange(1, lastCol).setValue(c);
      hdr.push(c);
      j = lastCol - 1;
    }
    idx[c] = j;
  }
  return idx;
}

function prEnsureMakerMasterJittaCols_(sh) {
  return prEnsureSheetCols_(sh, ['メーカー'].concat(PR_MAKER_JITTA_MASTER_COLS_));
}

function prEnsureMakerSumJittaMirrorCols_(sh) {
  return prEnsureSheetCols_(sh, PR_MAKER_JITTA_MIRROR_COLS_);
}

function prSplitPipeValues_(text) {
  var out = [];
  var seen = {};
  String(text || '')
    .split('|')
    .forEach(function (p) {
      var t = String(p || '').trim();
      if (!t || seen[t]) return;
      seen[t] = true;
      out.push(t);
    });
  return out;
}

function prJoinPipeValues_(arr) {
  return prSplitPipeValues_(arr.join('|')).join('|');
}

function prNormalizeGyoshu_(resolved) {
  resolved = resolved || {};
  var raw = String(resolved.gyoshu || '').trim();
  var src = String(resolved.gyoshu_source || '').trim();
  if (raw) {
    if (
      (src === 'gemini' || src === 'openai' || src === 'product_inferred') &&
      raw.indexOf(PR_JITTA_GYOSHU_INFERRED_SUFFIX_) < 0
    ) {
      return raw + PR_JITTA_GYOSHU_INFERRED_SUFFIX_;
    }
    return raw;
  }
  if (src === 'none') return PR_JITTA_GYOSHU_UNKNOWN_;
  return '';
}

function prNormalizeAmazonRootCategories_(text) {
  var allowed = prAmazonRootCatSet_();
  var kept = [];
  prSplitPipeValues_(text).forEach(function (v) {
    if (allowed[v]) kept.push(v);
  });
  return prJoinPipeValues_(kept);
}

function prEvidenceUrlsText_(urls) {
  if (typeof urls === 'string') return String(urls || '').trim();
  if (!urls || !urls.length) return '';
  var out = [];
  var i;
  for (i = 0; i < urls.length; i++) {
    var u = String(urls[i] || '').trim();
    if (u) out.push(u);
  }
  return out.join('|');
}

function prJittaResolvedToMasterPatch_(resolved) {
  var now = prNowFetchedAt_();
  return {
    実態候補: String(resolved.jitta_candidate || '').trim(),
    実態信頼度: String(resolved.confidence || '').trim(),
    実態採取元: String(resolved.source || '').trim(),
    実態根拠URL: prEvidenceUrlsText_(resolved.evidence_urls),
    実態取得日時: now,
    業種: prNormalizeGyoshu_(resolved),
    Amazonルートカテゴリ: prNormalizeAmazonRootCategories_(resolved.amazon_root_categories),
    AmazonBrowse棚: prJoinPipeValues_(prSplitPipeValues_(resolved.amazon_browse_shelves)),
    入口語候補: String(resolved.entry_needle || '').trim(),
    判定メモ: String(resolved.reason || '').trim().slice(0, 500),
  };
}

function prMirrorJittaPatch_(patch) {
  patch = patch || {};
  return {
    実態採取元: patch['実態採取元'] || '',
    実態信頼度: patch['実態信頼度'] || '',
    実態根拠URL: patch['実態根拠URL'] || '',
    業種: patch['業種'] || '',
    Amazonルートカテゴリ: patch['Amazonルートカテゴリ'] || '',
    AmazonBrowse棚: patch['AmazonBrowse棚'] || '',
  };
}

function prShouldAutoConfirmJitta_(summaryJitta, resolved) {
  var human = String(summaryJitta || '').trim();
  if (human === PR_JITTAI_EXCLUDE_) return false;
  var cand = String((resolved && resolved.jitta_candidate) || '').trim();
  var conf = String((resolved && resolved.confidence) || '').trim();
  return conf === 'high' && cand && cand !== PR_JITTAI_CHECK_;
}

function prParseJittaResolveJson_(text) {
  var raw = String(text || '').trim();
  raw = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```$/, '');
  var start = raw.indexOf('{');
  var end = raw.lastIndexOf('}');
  if (start < 0 || end <= start) return null;
  try {
    return JSON.parse(raw.slice(start, end + 1));
  } catch (e) {
    return null;
  }
}

function prBuildJittaResolvePrompt_(maker) {
  var roots = PR_AMAZON_ROOT_CATS_.join(' / ');
  return (
    'あなたは日本のEC仕入れ調査アシスタントです。メーカー「' +
    maker +
    '」について次を JSON 1オブジェクトのみで返してください。\n' +
    'キー: jitta_candidate, entry_preferred, entry_needle, confidence, source, evidence_urls, gyoshu, gyoshu_source, amazon_root_categories, amazon_browse_shelves, reason\n' +
    'jitta_candidate は次のいずれか: 製造メーカー / 入口ブランド / 要確認 / 非仕入先\n' +
    'entry_preferred: メーカー / ブランド / 要確認 / 非仕入先\n' +
    'confidence: high / medium / low\n' +
    'source: official / gemini / openai / 人手\n' +
    'evidence_urls: 根拠URLの配列\n' +
    'gyoshu: 一般的な業種（東証・日経・会社HPベース）。取れなければ空文字\n' +
    'gyoshu_source: official / gemini / openai / product_inferred / none\n' +
    'amazon_root_categories: Amazonルート（|区切り）。語彙は次のみ: ' +
    roots +
    '\n' +
    'amazon_browse_shelves: ルート内のBrowse棚名（|区切り）\n' +
    '判断の芯: パッケージに載る名前は会社名か別ブランド名か。発売元商社でも自社製造なら製造メーカー。\n' +
    '説明文は不要。JSONのみ。'
  );
}

function prLookupJittaAiFixture_(maker) {
  var key = String(maker || '').trim();
  if (!key || !PR_JITTA_AI_FIXTURES_[key]) return null;
  var raw = PR_JITTA_AI_FIXTURES_[key];
  return JSON.parse(JSON.stringify(raw));
}

function prFetchJittaResolveAi_(maker) {
  var prompt = prBuildJittaResolvePrompt_(maker);
  var gText = prGeminiGenerateText_(prompt);
  var parsed = prParseJittaResolveJson_(gText);
  if (parsed) {
    parsed.source = parsed.source || 'gemini';
    if (!parsed.gyoshu_source) parsed.gyoshu_source = 'gemini';
    return parsed;
  }
  var oText = prOpenAiGenerateText_(prompt);
  parsed = prParseJittaResolveJson_(oText);
  if (parsed) {
    parsed.source = parsed.source || 'openai';
    if (!parsed.gyoshu_source) parsed.gyoshu_source = 'openai';
    return parsed;
  }
  return prLookupJittaAiFixture_(maker);
}

function prLoadMakerMasterJittaRow_(maker) {
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_MASTER_SHEET_);
  if (!sh || sh.getLastRow() < 2) return null;
  var idx = prEnsureMakerMasterJittaCols_(sh);
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var key = prNormalizeMakerKey_(maker);
  var exact = null;
  var fuzzy = null;
  var r;
  for (r = 0; r < vals.length; r++) {
    var nm = String(vals[r][idx['メーカー']] || '').trim();
    if (!nm) continue;
    var row = { sheetRow: r + 2, sh: sh, idx: idx, メーカー: nm };
    var i;
    for (i = 0; i < PR_MAKER_JITTA_MASTER_COLS_.length; i++) {
      var c = PR_MAKER_JITTA_MASTER_COLS_[i];
      row[c] = vals[r][idx[c]];
    }
    if (nm === maker) {
      exact = row;
      break;
    }
    if (!fuzzy && key && prNormalizeMakerKey_(nm) === key) fuzzy = row;
  }
  return exact || fuzzy;
}

function prFindMakerSumRowsByKey_(sh, maker) {
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var key = prNormalizeMakerKey_(maker);
  if (!key) return out;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = prHeaderIndex_(hdr, 'メーカー');
  if (idx < 0) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    var nm = String(vals[r][idx] || '').trim();
    if (!nm) continue;
    if (prNormalizeMakerKey_(nm) === key) out.push({ sheetRow: r + 2, name: nm });
  }
  return out;
}

function prCanonicalMakerForJittaUpsert_(maker) {
  maker = String(maker || '').trim();
  var loc = prLoadMakerMasterJittaRow_(maker);
  if (loc && loc['メーカー']) return String(loc['メーカー']).trim();
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_SUM_SHEET_);
  var names = [];
  if (sh) {
    prFindMakerSumRowsByKey_(sh, maker).forEach(function (x) {
      names.push(x.name);
    });
  }
  names.push(maker);
  var picked = prPickCanonicalMakerName_(names);
  var key = prNormalizeMakerKey_(maker);
  if (picked && prNormalizeMakerKey_(picked) === key) return picked;
  return key || maker;
}

function prUpsertMakerMasterJitta_(maker, patch) {
  var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_MAKER_MASTER_SHEET_);
  if (!sh) throw new Error('メーカーマスタが無い');
  var idx = prEnsureMakerMasterJittaCols_(sh);
  var canon = prCanonicalMakerForJittaUpsert_(maker);
  var loc = prLoadMakerMasterJittaRow_(maker);
  var sheetRow;
  if (loc) {
    sheetRow = loc.sheetRow;
  } else {
    sheetRow = sh.getLastRow() + 1;
    sh.getRange(sheetRow, idx['メーカー'] + 1).setValue(canon);
  }
  var c;
  for (c in patch) {
    if (!Object.prototype.hasOwnProperty.call(patch, c)) continue;
    if (idx[c] == null) continue;
    sh.getRange(sheetRow, idx[c] + 1).setValue(patch[c]);
  }
  return sheetRow;
}

function prApplyJittaMirrorToSummaryRow_(sh, sheetRow, patch, resolved, opt) {
  opt = opt || {};
  var idxMirror = prEnsureMakerSumJittaMirrorCols_(sh);
  var mirror = prMirrorJittaPatch_(patch);
  var c;
  for (c in mirror) {
    if (!Object.prototype.hasOwnProperty.call(mirror, c)) continue;
    if (idxMirror[c] == null) continue;
    sh.getRange(sheetRow, idxMirror[c] + 1).setValue(mirror[c]);
  }
  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxJ = prHeaderIndex_(hdr, '実態');
  var idxPref = prHeaderIndex_(hdr, '入口推奨');
  var curJ = idxJ >= 0 ? String(sh.getRange(sheetRow, idxJ + 1).getValue() || '').trim() : '';
  var auto = !opt.skipAutoConfirm && prShouldAutoConfirmJitta_(curJ, resolved);
  var appliedJitta = curJ;
  if (auto && idxJ >= 0) {
    appliedJitta = String(resolved.jitta_candidate || '').trim();
    sh.getRange(sheetRow, idxJ + 1).setValue(appliedJitta);
    if (idxPref >= 0) {
      sh.getRange(sheetRow, idxPref + 1).setValue(
        String(resolved.entry_preferred || '').trim() || prEntryPreferredForJitta_(appliedJitta)
      );
    }
  } else if (idxJ >= 0) {
    var cand = String(resolved.jitta_candidate || '').trim();
    var conf = String(resolved.confidence || patch['実態信頼度'] || '').trim();
    var promoteUninvestigated = !!cand && !!conf && prMakerJittaIsUninvestigatedLabel_(curJ);
    var promoteLegacyCheck =
      !!cand && !!conf && !opt.skipAutoConfirm && curJ === PR_JITTAI_CHECK_;
    if (promoteUninvestigated || promoteLegacyCheck) {
      appliedJitta = cand;
      sh.getRange(sheetRow, idxJ + 1).setValue(cand);
      if (idxPref >= 0) {
        sh.getRange(sheetRow, idxPref + 1).setValue(
          String(resolved.entry_preferred || '').trim() || prEntryPreferredForJitta_(cand)
        );
      }
    }
  }
  prSyncMakerJittaCkUrlRow_(sh, sheetRow);
  return { sheetRow: sheetRow, autoConfirm: auto, appliedJitta: appliedJitta };
}

function prApplyJittaMirrorToSummary_(maker, patch, resolved, opt) {
  opt = opt || {};
  var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
  var rows = prFindMakerSumRowsByKey_(sh, maker);
  if (!rows.length) {
    var sheetRow = prEnsureMakerSumRowTop_(sh, maker, {});
    rows = [{ sheetRow: sheetRow, name: maker }];
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var last = { sheetRow: 0, autoConfirm: false, appliedJitta: '', mirrored: 0 };
  var n = 0;
  var i;
  for (i = 0; i < rows.length; i++) {
    var rowVals = sh.getRange(rows[i].sheetRow, 1, rows[i].sheetRow, lastCol).getValues()[0];
    if (!prMakerSummaryRowUninvestigated_(rowVals, hdr)) continue;
    last = prApplyJittaMirrorToSummaryRow_(sh, rows[i].sheetRow, patch, resolved, opt);
    n++;
  }
  // 対象が無かった場合（すべて調査済）は先頭行の状態だけ返す
  if (!n && rows.length) {
    last.sheetRow = rows[0].sheetRow;
    var idxJ = prHeaderIndex_(hdr, '実態');
    last.appliedJitta =
      idxJ >= 0 ? String(sh.getRange(rows[0].sheetRow, idxJ + 1).getValue() || '').trim() : '';
  }
  last.mirrored = n;
  return last;
}

/**
 * 1メーカーの実態リゾルバ。人入力済みサマリ.実態は触らない。
 * 法人格ゆれの別名行へも同一結果をミラーする。
 * @return {{skipped:boolean, maker:string, source:string, autoConfirm:boolean}}
 */
function prResolveMakerJitta_(maker, opt) {
  opt = opt || {};
  maker = String(maker || '').trim();
  if (!maker) return { skipped: true, maker: '', reason: 'empty' };
  var sumSh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
  var sumRow = prEnsureMakerSumRowTop_(sumSh, maker, {});
  var lastCol = sumSh.getLastColumn();
  var hdr = sumSh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxJ = prHeaderIndex_(hdr, '実態');
  if (prMakerJittaSummaryAiConfirmed_(sumSh, sumRow, hdr)) {
    return {
      skipped: true,
      maker: maker,
      reason: 'ai_confirmed',
      jitta: idxJ >= 0 ? String(sumSh.getRange(sumRow, idxJ + 1).getValue() || '').trim() : '',
    };
  }
  var master = prLoadMakerMasterJittaRow_(maker);
  if (
    master &&
    prCkIsTrue_(master['実態再採取しない']) &&
    !opt.forceResample
  ) {
    var cachedNo = {};
    var iNo;
    for (iNo = 0; iNo < PR_MAKER_JITTA_MASTER_COLS_.length; iNo++) {
      var colNo = PR_MAKER_JITTA_MASTER_COLS_[iNo];
      cachedNo[colNo] = master[colNo];
    }
    prApplyJittaMirrorToSummary_(
      maker,
      cachedNo,
      {
        jitta_candidate: master['実態候補'],
        confidence: master['実態信頼度'],
        entry_preferred: prEntryPreferredForJitta_(String(master['実態候補'] || '').trim()),
      },
      { skipAutoConfirm: true }
    );
    return { skipped: true, maker: maker, reason: 'cache_no_resample', jitta: String(master['実態候補'] || '') };
  }
  if (prMakerMasterJittaCached_(master) && !opt.forceResample && !opt.skipCacheHeal) {
    var cached = {};
    var i;
    for (i = 0; i < PR_MAKER_JITTA_MASTER_COLS_.length; i++) {
      var col = PR_MAKER_JITTA_MASTER_COLS_[i];
      cached[col] = master[col];
    }
    var healed = prApplyJittaMirrorToSummary_(
      maker,
      cached,
      {
        jitta_candidate: master['実態候補'],
        confidence: master['実態信頼度'],
        entry_preferred: prEntryPreferredForJitta_(String(master['実態候補'] || '').trim()),
      },
      opt
    );
    Logger.log(
      '[仕入れ検討①] jitta_heal key=%s maker=%s cand=%s conf=%s rows=%s',
      prNormalizeMakerKey_(maker),
      maker,
      master['実態候補'],
      master['実態信頼度'],
      healed.mirrored || 1
    );
    return {
      skipped: false,
      fromCache: true,
      maker: maker,
      source: String(master['実態採取元'] || ''),
      jitta: healed.appliedJitta || String(master['実態候補'] || ''),
      autoConfirm: !!healed.autoConfirm,
      confidence: String(master['実態信頼度'] || ''),
    };
  }
  var promptMaker = prCanonicalMakerForJittaUpsert_(maker);
  var resolved = opt.resolved || null;
  if (!resolved && !opt.skipAi) resolved = prFetchJittaResolveAi_(promptMaker);
  if (!resolved) {
    return {
      skipped: false,
      maker: maker,
      reason: 'no_resolve',
      jitta: prJittaSeed_(maker) || PR_JITTAI_UNCHECKED_,
    };
  }
  var patch = prJittaResolvedToMasterPatch_(resolved);
  prUpsertMakerMasterJitta_(maker, patch);
  var applied = prApplyJittaMirrorToSummary_(maker, patch, resolved, opt);
  Logger.log(
    '[仕入れ検討①] jitta_resolve maker=%s key=%s cand=%s conf=%s auto=%s applied=%s rows=%s',
    maker,
    prNormalizeMakerKey_(maker),
    patch['実態候補'],
    patch['実態信頼度'],
    applied.autoConfirm,
    applied.appliedJitta,
    applied.mirrored || 1
  );
  return {
    skipped: false,
    fromCache: false,
    maker: maker,
    source: patch['実態採取元'],
    jitta:
      applied.appliedJitta ||
      (applied.autoConfirm ? patch['実態候補'] : prJittaSeed_(maker) || PR_JITTAI_UNCHECKED_),
    autoConfirm: applied.autoConfirm,
    confidence: patch['実態信頼度'],
  };
}

/** C1+C2 後: ヒットメーカーのうち未調査を可能な限り AI 調査。 */
function prResolveJittaForHitMakers_(hits, opt) {
  opt = opt || {};
  if (!prJittaResolveEnabled_()) return 0;
  var batch = prResolveJittaBatch_(prCollectJittaResolveTargets_(hits || {}), opt);
  return batch.resolved;
}

/** メニュー: サマリの未調査メーカーを AI 調査。 */
function prMenuResolveMakerJitta() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=jitta_resolve state=RUNNING functionName=prMenuResolveMakerJitta', runId);
  try {
    if (!prJittaResolveEnabled_()) throw new Error(PR_JITTA_RESOLVE_ENABLED_PROP_ + '=false');
    var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
    prEnsureMakerSumJittaMirrorCols_(sh);
    prEnsureMakerSumJittaDropdown_(sh);
    if (sh.getLastRow() < 2) {
      prMenuToast_('実態AI調査', 'メーカーサマリにデータ行がありません', 8);
      return { ok: false, n: 0 };
    }
    var targets = prCollectJittaResolveTargets_({});
    var batch = prResolveJittaBatch_(targets, {});
    var detail =
      'resolve=' +
      batch.resolved +
      ' auto=' +
      batch.autoN +
      ' skip=' +
      batch.skipped +
      ' cap=' +
      batch.cap +
      ' uninvestigated=' +
      targets.length;
    prAppendLog_([runId, '実態AI調査', 'DONE', new Date(), '', batch.resolved, batch.autoN, detail]);
    prMenuToast_('実態AI調査', detail);
    Logger.log('[仕入れ検討①] runId=%s step=jitta_resolve state=DONE %s', runId, detail);
    return { ok: true, n: batch.resolved, autoN: batch.autoN, detail: detail };
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=jitta_resolve state=FAILED %s', runId, msg);
    prAppendLog_([runId, '実態AI調査', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/**
 * レガシー: 実態=要確認かつ実態信頼度が空 → 未調査へ移行（AI前プレースホルダ）。
 * AI済みの要確認（信頼度あり）は触らない。
 */
function prMenuMigrateJittaUnchecked_() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=jitta_migrate_unchecked state=RUNNING', runId);
  try {
    var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
    prEnsureMakerSumJittaMirrorCols_(sh);
    prEnsureMakerSumJittaDropdown_(sh);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) {
      prMenuToast_('実態ラベル移行', 'データ行がありません', 8);
      return { ok: false, n: 0 };
    }
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxJ = prHeaderIndex_(hdr, '実態');
    var idxPref = prHeaderIndex_(hdr, '入口推奨');
    if (idxJ < 0) throw new Error('メーカーサマリに 実態 列が無い');
    var vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var n = 0;
    var r;
    for (r = 0; r < vals.length; r++) {
      if (!prMakerSummaryRowUninvestigated_(vals[r], hdr)) continue;
      if (String(vals[r][idxJ] || '').trim() !== PR_JITTAI_CHECK_) continue;
      sh.getRange(r + 2, idxJ + 1).setValue(PR_JITTAI_UNCHECKED_);
      if (idxPref >= 0) sh.getRange(r + 2, idxPref + 1).setValue('未調査');
      prSyncMakerJittaCkUrlRow_(sh, r + 2);
      n++;
    }
    var detail = 'migrated=' + n;
    prAppendLog_([runId, '実態ラベル移行', 'DONE', new Date(), '', n, '', detail]);
    prMenuToast_('実態ラベル移行', detail);
    Logger.log('[仕入れ検討①] runId=%s step=jitta_migrate_unchecked state=DONE %s', runId, detail);
    return { ok: true, n: n, detail: detail };
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=jitta_migrate_unchecked state=FAILED %s', runId, msg);
    prAppendLog_([runId, '実態ラベル移行', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/**
 * C1 単店 A-light。小中大は query ページ送り＋product スライス（続きあり）。
 * 特大(いずれかカテゴリ≥501)は浅い抽出のみ hold。門・①候補・Keepaフル upsert なし。
 * 規模＝カテゴリ別 query（CYCLE §4.1bis）。ASIN 取得＝レ点ルート OR。
 * @param {string} sellerId
 * @param {string} runId
 * @param {string} shopName
 * @param {{productCap?:number, oldExtract?:string, cont?:Object}=} opt
 */
function prRunAlightShop_(sellerId, runId, shopName, opt) {
  opt = opt || {};
  var sid = String(sellerId || '').trim();
  var productCap = opt.productCap != null ? opt.productCap : prAlightProductCap_();
  var name = shopName || '';
  var cont = opt.cont || null;
  var page = cont ? Number(cont.page) || 0 : 0;
  var asinOff = cont ? Number(cont.asinOff) || 0 : 0;
  var processed = cont ? Number(cont.processed) || 0 : 0;
  var total = cont ? Number(cont.total) || 0 : -1;
  var band = cont ? String(cont.band || '') : '';
  var qTok = cont ? Number(cont.qTok) || 0 : 0;
  var pTok = cont ? Number(cont.pTok) || 0 : 0;
  var pageAsins = cont && Array.isArray(cont.pageAsins) ? cont.pageAsins.slice() : null;
  var walkRoots =
    cont && Array.isArray(cont.walkRoots) && cont.walkRoots.length
      ? cont.walkRoots.slice()
      : null;
  var skuEst = cont && cont.skuEst != null ? Number(cont.skuEst) : -1;
  var skuBreakdown = cont ? String(cont.skuBreakdown || '') : '';
  var bandBreakdown = cont ? String(cont.bandBreakdown || '') : '';
  var oldExtract = opt.oldExtract || '';

  prEnsureSellerStorefront_(sid, true);
  if (!name) name = prSellerNameForId_(sid);

  function packOut_(status, pack, newMakers, merged, extra) {
    extra = extra || {};
    var detail =
      'alight status=' +
      status +
      ' band=' +
      (band || '') +
      ' skuEst=' +
      skuEst +
      ' total=' +
      total +
      ' processed=' +
      processed +
      ' page=' +
      page +
      ' asinOff=' +
      asinOff +
      ' roots=' +
      (walkRoots || []).join('|') +
      ' product_n=' +
      ((pack && pack.products && pack.products.length) || 0) +
      ' makers=' +
      (newMakers || []).length +
      ' qTok=' +
      qTok +
      ' pTok=' +
      pTok +
      ' seller=' +
      sid +
      (extra.detailSuffix || '');
    return {
      status: status,
      total: skuEst >= 0 ? skuEst : total,
      walkTotal: total,
      hold: status === 'hold',
      band: band,
      skuEst: skuEst,
      skuBreakdown: skuBreakdown,
      bandBreakdown: bandBreakdown,
      newMakers: newMakers || [],
      merged: merged || oldExtract,
      asinCount: processed,
      tokensConsumed: qTok + pTok,
      tokensLeft: (pack && pack.tokensLeft) || extra.tokensLeft || '',
      refillIn: (pack && pack.refillIn) || extra.refillIn || '',
      waitMin: extra.waitMin || 0,
      needContinue: status === 'continue' || status === 'rate_limit',
      detail: detail,
    };
  }

  function contPayload_() {
    return {
      sid: sid,
      page: page,
      asinOff: asinOff,
      processed: processed,
      total: total,
      band: band,
      qTok: qTok,
      pTok: pTok,
      pageAsins: pageAsins || [],
      walkRoots: walkRoots || [],
      skuEst: skuEst,
      skuBreakdown: skuBreakdown,
      bandBreakdown: bandBreakdown,
    };
  }

  // 初回: カテゴリ別規模測定。特大は浅いだけ
  if (!cont) {
    var meas = prMeasureSellerSkuByCategory_(sid);
    qTok += Number(meas.tokensConsumed) || 0;
    skuEst = Number(meas.skuEst) || 0;
    band = meas.band;
    skuBreakdown = meas.skuBreakdown;
    bandBreakdown = meas.bandBreakdown;
    walkRoots = meas.walkRoots.slice();
    page = 0;
    asinOff = 0;
    processed = 0;
    if (meas.hold) {
      var asinXl;
      if (meas.page0Asins) {
        asinXl = meas.page0Asins.slice(0, productCap);
        total = meas.page0Total;
      } else {
        var qXl = prKeepaQueryGet_(sid, walkRoots, 0, 100);
        qTok += Number(qXl.tokensConsumed) || 0;
        asinXl = (qXl.asinList || []).slice(0, productCap);
        total = Number(qXl.totalResults || 0);
      }
      pageAsins = asinXl.slice();
      var packXl = prKeepaProductMakerLight_(asinXl, productCap);
      pTok += Number(packXl.tokensConsumed) || 0;
      if (packXl.rateLimited) {
        return packOut_('rate_limit', packXl, [], oldExtract, {
          detailSuffix: ' xl_shallow',
          waitMin: 5,
        });
      }
      var makersXl = prExtractMakersFromProducts_(packXl.products, productCap);
      var mergedXl = prMergeExtractMakerCell_(oldExtract, makersXl);
      processed = asinXl.length;
      return packOut_('hold', packXl, makersXl, mergedXl, { detailSuffix: ' xl_no_full' });
    }
    if (meas.page0Asins) {
      pageAsins = meas.page0Asins.slice();
      total = meas.page0Total;
    } else {
      var q0 = prKeepaQueryGet_(sid, walkRoots, 0, 100);
      qTok += Number(q0.tokensConsumed) || 0;
      pageAsins = (q0.asinList || []).slice();
      total = Number(q0.totalResults || 0);
    }
  }

  if (!walkRoots || !walkRoots.length) {
    walkRoots = prSurveyRootCategoryIds_();
  }

  if (prGasTimeUp_()) {
    prAlightContSet_(contPayload_());
    return packOut_('continue', null, [], oldExtract, {
      detailSuffix: ' timeup',
      waitMin: 1,
    });
  }

  // ページ枯渇 → 次ページ
  if (!pageAsins || asinOff >= pageAsins.length) {
    if (pageAsins && asinOff >= pageAsins.length && processed > 0) {
      page += 1;
      asinOff = 0;
      pageAsins = null;
    }
    if (!pageAsins) {
      if (total >= 0 && processed >= total) {
        return packOut_('done', null, [], oldExtract, { detailSuffix: ' complete' });
      }
      var qn = prKeepaQueryGet_(sid, walkRoots, page, 100);
      qTok += Number(qn.tokensConsumed) || 0;
      pageAsins = (qn.asinList || []).slice();
      if (!pageAsins.length) {
        return packOut_('done', null, [], oldExtract, { detailSuffix: ' page_empty' });
      }
      if (total < 0) {
        total = Number(qn.totalResults || 0);
        if (skuEst < 0) skuEst = total;
        if (!band) band = prSellerBandFromSkuN_(skuEst >= 0 ? skuEst : total);
      }
    }
  }

  var slice = pageAsins.slice(asinOff, asinOff + productCap);
  if (!slice.length) {
    return packOut_('done', null, [], oldExtract, { detailSuffix: ' slice_empty' });
  }
  var pack = prKeepaProductMakerLight_(slice, productCap);
  pTok += Number(pack.tokensConsumed) || 0;
  if (pack.rateLimited) {
    prAlightContSet_(contPayload_());
    return packOut_('rate_limit', pack, [], oldExtract, {
      detailSuffix: ' slice',
      waitMin: 5,
    });
  }
  var newMakers = prExtractMakersFromProducts_(pack.products, slice.length);
  var merged = prMergeExtractMakerCell_(oldExtract, newMakers);
  asinOff += slice.length;
  processed += slice.length;

  var shopDone = processed >= total || (asinOff >= pageAsins.length && pageAsins.length < 100);
  if (!shopDone && asinOff >= pageAsins.length) {
    if (pageAsins.length < 100 || processed >= total) shopDone = true;
  }

  if (shopDone) {
    prAlightContSet_(null);
    return packOut_('done', pack, newMakers, merged, { detailSuffix: ' shop_done' });
  }

  prAlightContSet_(contPayload_());
  return packOut_('continue', pack, newMakers, merged, {
    detailSuffix: ' slice_ok',
    waitMin: 1,
  });
}

function prEnsureSellerBandCols_(sh) {
  sh = sh || prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_SELLER_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  function ensureCol_(name) {
    var idx = prHeaderIndex_(hdr, name);
    if (idx >= 0) return idx;
    sh.getRange(1, lastCol + 1).setValue(name);
    lastCol++;
    hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    return lastCol - 1;
  }
  return {
    idxBand: ensureCol_(PR_SELLER_BAND_COL_),
    idxSku: ensureCol_(PR_SELLER_SKU_COL_),
    idxSkuBd: ensureCol_(PR_SELLER_SKU_BREAKDOWN_COL_),
    idxBandBd: ensureCol_(PR_SELLER_BAND_BREAKDOWN_COL_),
  };
}

function prMakerSumNameSet_() {
  var set = {};
  var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_MAKER_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return set;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxMaker = prHeaderIndex_(hdr, 'メーカー');
  if (idxMaker < 0) return set;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    var name = String(vals[r][idxMaker] || '').trim();
    if (name) set[name] = true;
  }
  return set;
}

function prCountNewMakersVsSet_(makers, set) {
  var n = 0;
  var i;
  for (i = 0; i < (makers || []).length; i++) {
    var m = makers[i];
    if (!m || !prHitEligibleMaker_(m)) continue;
    if (set[m]) continue;
    set[m] = true;
    n++;
  }
  return n;
}

/** C1 バッチ + C2 セラーヒット再計算。1実行1店フル（続きあり）。 */
function prRunAlightHit_(opt) {
  opt = opt || {};
  prRefuseOriginalSs_();
  if (!prAlightEnabled_()) {
    return { ok: true, skipped: true, detail: 'PR_ALIGHT_ENABLED=false' };
  }
  if (!opt.skipCkCheck && !prCkEnabled_(PR_CK_SELLER_ENABLED_PROP_)) {
    throw new Error(PR_CK_SELLER_ENABLED_PROP_ + '=false（切戻し中）');
  }
  var runId = String(opt.runId || '');
  var sh = prEnsureLeadingCk_(PR_COMPETITOR_SS_, PR_SELLER_SHEET_);
  var bandCols = prEnsureSellerBandCols_(sh);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return { ok: false, detail: 'no_rows', hitMakers: 0, doneN: 0, holdN: 0, failN: 0 };
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxCk = prHeaderIndex_(hdr, PR_CK_HEADER_);
  var idxId = prHeaderIndex_(hdr, 'sellerId');
  var idxName = prHeaderIndex_(hdr, '店名');
  var idxMemo = prHeaderIndex_(hdr, 'メモ');
  var idxWalk = prHeaderIndex_(hdr, '巡回日');
  var idxCnt = prHeaderIndex_(hdr, 'asinList件数');
  var idxSrc = prHeaderIndex_(hdr, '採取元');
  var idxExtract = prHeaderIndex_(hdr, '抽出メーカー');
  var idxBand = bandCols.idxBand;
  var idxSku = bandCols.idxSku;
  var idxSkuBd = bandCols.idxSkuBd;
  var idxBandBd = bandCols.idxBandBd;
  if (idxCk < 0 || idxId < 0) throw new Error('セラーに 調査CK または sellerId が無い');
  var vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var productCap = prAlightProductCap_();
  var shopCap = opt.alightOnly === false ? 0 : prAlightShopCap_();
  var cont = prAlightContGet_();
  var targets = [];
  var trueN = 0;
  var r;

  if (cont && cont.sid) {
    for (r = 0; r < vals.length; r++) {
      var sidC = String(vals[r][idxId] || '').trim();
      if (sidC !== cont.sid) continue;
      targets.push({
        sheetRow: r + 2,
        sellerId: sidC,
        name: idxName >= 0 ? String(vals[r][idxName] || '').trim() : '',
        memo: idxMemo >= 0 ? vals[r][idxMemo] : '',
        oldExtract: idxExtract >= 0 ? String(vals[r][idxExtract] || '') : '',
        cont: cont,
      });
      break;
    }
    if (!targets.length) {
      prAlightContSet_(null);
      Logger.log('[仕入れ検討①] alight_cont orphan cleared sid=%s', cont.sid);
    }
  } else if (shopCap > 0) {
    for (r = 0; r < vals.length; r++) {
      if (!prCkIsTrue_(vals[r][idxCk])) continue;
      // 調査済みなのにレ点が残っていたら外してスキップ（再調査しない）
      if (prSellerAlreadyShopResearched_(vals[r], idxWalk, idxSrc)) {
        sh.getRange(r + 2, idxCk + 1).setValue(false);
        continue;
      }
      trueN++;
      var sid = String(vals[r][idxId] || '').trim();
      if (!sid) continue;
      var name = idxName >= 0 ? String(vals[r][idxName] || '').trim() : '';
      if (!name && idxName >= 0) name = prSellerNameFromSameId_(vals, idxId, idxName, sid);
      targets.push({
        sheetRow: r + 2,
        sellerId: sid,
        name: name,
        memo: idxMemo >= 0 ? vals[r][idxMemo] : '',
        oldExtract: idxExtract >= 0 ? String(vals[r][idxExtract] || '') : '',
        cont: null,
      });
      if (targets.length >= shopCap) break;
    }
  }

  var doneN = 0;
  var holdN = 0;
  var failN = 0;
  var rateN = 0;
  var contN = 0;
  var lines = [];
  var needContinue = false;
  var waitMin = 0;
  var refillIn = '';
  var tokensLeft = '';
  var sumSet = prMakerSumNameSet_();
  var newMakerN = 0;

  if (targets.length) {
    var i;
    for (i = 0; i < targets.length; i++) {
      if (prGasTimeUp_()) {
        lines.push('timeup leftover_shops=' + (targets.length - i));
        needContinue = !!prAlightContGet_();
        break;
      }
      var t = targets[i];
      Logger.log(
        '[仕入れ検討①] runId=%s step=C1 alight %s/%s %s %s cont=%s',
        runId,
        i + 1,
        targets.length,
        t.sellerId,
        t.name,
        t.cont ? '1' : '0'
      );
      // 1店をスライス完走まで同一実行で回す（時間／429で抜ける）
      var g = null;
      var loopGuard = 0;
      var curCont = t.cont;
      var curExtract = t.oldExtract;
      while (loopGuard < 40) {
        loopGuard++;
        if (prGasTimeUp_()) {
          if (!prAlightContGet_() && curCont) {
            prAlightContSet_(curCont);
          }
          needContinue = true;
          waitMin = Math.max(waitMin, 1);
          lines.push(t.sellerId + ':timeup_mid');
          break;
        }
        try {
          g = prRunAlightShop_(t.sellerId, runId, t.name, {
            productCap: productCap,
            oldExtract: curExtract,
            cont: curCont,
          });
        } catch (e1) {
          failN++;
          g = { status: 'fail', detail: prRedactKey_(String((e1 && e1.message) || e1)).slice(0, 180) };
        }
        if (g.merged) curExtract = g.merged;
        if (idxExtract >= 0 && g.merged) {
          sh.getRange(t.sheetRow, idxExtract + 1).setValue(g.merged);
        }
        if (idxBand >= 0 && g.band) {
          sh.getRange(t.sheetRow, idxBand + 1).setValue(g.band);
        }
        var skuWrite = g.skuEst != null && g.skuEst >= 0 ? g.skuEst : g.total;
        if (idxSku >= 0 && skuWrite != null && skuWrite >= 0) {
          sh.getRange(t.sheetRow, idxSku + 1).setValue(skuWrite);
        }
        if (idxSkuBd >= 0 && g.skuBreakdown) {
          sh.getRange(t.sheetRow, idxSkuBd + 1).setValue(g.skuBreakdown);
        }
        if (idxBandBd >= 0 && g.bandBreakdown) {
          sh.getRange(t.sheetRow, idxBandBd + 1).setValue(g.bandBreakdown);
        }
        if (g.status === 'rate_limit') {
          needContinue = true;
          contN++;
          rateN++;
          waitMin = Math.max(waitMin, Number(g.waitMin) || 5);
          refillIn = g.refillIn || refillIn;
          tokensLeft = g.tokensLeft || tokensLeft;
          lines.push(t.sellerId + ':rate_limit ' + (g.detail || '').slice(0, 80));
          prAppendLog_([
            runId,
            'C1 A-light ' + t.sellerId,
            'RETRYING',
            new Date(),
            g.total,
            g.asinCount,
            g.tokensConsumed,
            g.detail,
          ]);
          break;
        }
        if (g.status === 'continue') {
          curCont = prAlightContGet_();
          contN++;
          // 同一実行で次スライスへ（時間切れ時のみトリガー続き）
          continue;
        }
        if (g.status === 'hold' || g.status === 'done') {
          prAlightContSet_(null);
          if (g.status === 'hold') {
            holdN++;
            if (idxMemo >= 0) {
              sh.getRange(t.sheetRow, idxMemo + 1).setValue(
                prAppendMemo_(t.memo, 'A-light HOLD total=' + g.total + ' ' + prYmdTokyo_())
              );
            }
          } else {
            doneN++;
          }
          sh.getRange(t.sheetRow, idxCk + 1).setValue(false);
          if (idxWalk >= 0) sh.getRange(t.sheetRow, idxWalk + 1).setValue(prYmdTokyo_());
          if (idxCnt >= 0 && g.total != null) sh.getRange(t.sheetRow, idxCnt + 1).setValue(g.total);
          if (idxSrc >= 0) {
            sh.getRange(t.sheetRow, idxSrc + 1).setValue(
              g.status === 'hold' ? 'alight+xl_shallow' : 'alight+shop_full'
            );
          }
          var makersFinal = prMakersInCell_(g.merged || curExtract);
          var shopNew = prCountNewMakersVsSet_(makersFinal, sumSet);
          newMakerN += shopNew;
          prAlightDailyDoneBump_(1);
          prAlightRoundBump_();
          prAlightSetLastNewMakers_(shopNew);
          lines.push(
            t.sellerId +
              ':' +
              g.status +
              ' makers=' +
              makersFinal.length +
              ' new=' +
              shopNew
          );
          prAppendLog_([
            runId,
            'C1 A-light ' + t.sellerId,
            'DONE',
            new Date(),
            g.total,
            g.asinCount,
            g.tokensConsumed,
            g.detail || '',
          ]);
          break;
        }
        failN++;
        lines.push(t.sellerId + ':fail ' + (g.detail || ''));
        prAppendLog_([
          runId,
          'C1 A-light ' + t.sellerId,
          'FAILED',
          new Date(),
          g.total,
          g.asinCount,
          g.tokensConsumed,
          g.detail || '',
        ]);
        break;
      }
      // continue 中は1店のみ（次店へ進まない）
      if (needContinue) break;
    }
    prClearSellerLedgerCache_();
  } else if (shopCap > 0) {
    lines.push('ck_true=' + trueN + ' targets=0');
  }

  var hits = prRecomputeSellerHits_(sh);
  var hitMakers = prSyncMakerSumHits_(hits);
  var jittaN = 0;
  var jittaNote = '';
  try {
    if (prJittaAfterAlightEnabled_()) {
      jittaN = prResolveJittaForHitMakers_(hits, { runId: runId });
    } else {
      jittaNote = ' jitta_deferred_to_hb';
    }
  } catch (eJ) {
    Logger.log(
      '[仕入れ検討①] jitta_resolve after C2 skip %s',
      String((eJ && eJ.message) || eJ).slice(0, 120)
    );
  }
  var summary =
    'C1 done=' +
    doneN +
    ' hold=' +
    holdN +
    ' fail=' +
    failN +
    ' rate=' +
    rateN +
    ' cont=' +
    contN +
    ' new_makers=' +
    newMakerN +
    ' daily_done=' +
    prAlightDailyDoneCount_() +
    ' rounds=' +
    prAlightRoundCount_() +
    ' C2 hit_makers=' +
    hitMakers +
    ' jitta=' +
    jittaN +
    jittaNote +
    ' shop_cap=' +
    shopCap +
    ' product_cap=' +
    productCap +
    ' ck_true=' +
    trueN;
  Logger.log('[仕入れ検討①] runId=%s step=C1+C2 state=DONE %s', runId, summary);
  return {
    ok: failN === 0 && rateN === 0 && !needContinue,
    needContinue: needContinue,
    waitMin: waitMin,
    refillIn: refillIn,
    tokensLeft: tokensLeft,
    detail: summary + ' ' + lines.join(' ').slice(0, 200),
    hitMakers: hitMakers,
    doneN: doneN,
    holdN: holdN,
    failN: failN,
    rateN: rateN,
    newMakerN: newMakerN,
    lines: lines,
  };
}

/** C1+C2: レ点店 A-light → セラーヒット全店再計算。 */
function prMenuAlightHit() {
  prCyclePauseForOtherMenu_('C1+C2');
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=C1+C2 state=RUNNING functionName=prMenuAlightHit', runId);
  try {
    var r = prRunAlightHit_({ runId: runId });
    if (r.skipped) {
      prMenuToast_('C1+C2 A-light', r.detail);
      return r;
    }
    var state = r.ok ? 'DONE' : r.rateN ? 'RETRYING' : 'FAILED';
    prAppendLog_([runId, 'C1+C2 A-light＋セラーヒット', state, new Date(), '', r.hitMakers, '', r.detail]);
    prMenuToast_('C1+C2 A-light', r.detail);
    if (state === 'FAILED') prMenuFailAlert_(r.detail);
    return r;
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=C1+C2 state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'C1+C2 A-light＋セラーヒット', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}

/** C2 のみ: 既存 セラー.抽出メーカー から セラーヒット バックフィル。 */
function prMenuSellerHitBackfill() {
  prGasStart_();
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=C2 state=RUNNING functionName=prMenuSellerHitBackfill', runId);
  try {
    var r = prRunAlightHit_({ runId: runId, alightOnly: false, skipCkCheck: true });
    var detail = 'C2 backfill hit_makers=' + r.hitMakers + ' ' + (r.detail || '');
    prAppendLog_([runId, 'C2 セラーヒット再計算', 'DONE', new Date(), '', r.hitMakers, '', detail]);
    prMenuToast_('C2 セラーヒット', detail);
    return r;
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s step=C2 state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'C2 セラーヒット再計算', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    prMenuFailAlert_(msg);
    throw e;
  }
}
