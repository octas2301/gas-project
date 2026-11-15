/**
 * 未広告ASIN掲載（第1スライス）
 * - SP-APIで出品一覧（出品開始日＝createdDate 降順）＋親ASIN（Listings → Catalog）
 * - Drive「📥広告中の商品レポート」最新CSVで広告あり／なし・既存キャンペーン名
 * - Drive「📥一括キャンペーンDL」最新xlsx/CSV → シート「ID参照用」を自動更新（既存足し用）
 * - 親キャンペーン（オート＋Manual_）／広告グループは子ごと／既存親へSKU・AG追加
 * - キャンペーン名は可読性優先（硬上限128・ASIN末尾識別なし）
 *
 * 正本: docs/広告運用_未広告ASIN掲載_第1スライス要件.md
 *
 * Script Properties:
 *   UNADVERTISED_ASIN_ENABLED … 未設定=ON（falseで停止）
 *   UNADVERTISED_ASIN_DAILY_BUDGET … 既定 1000
 *   UNADVERTISED_ASIN_DEFAULT_BID … 既定 11
 *   UNADVERTISED_ASIN_MAX_LISTINGS … 既定 200（1実行の上限）
 *   UNADVERTISED_ASIN_CSV_MAX_AGE_DAYS … 既定 14
 *   UNADVERTISED_ASIN_ID_AUTO_FROM_BULK … 未設定=ON（falseでID参照用の自動更新を停止）
 *   LISTING_MASTER_SS_ID … 出品マスタSS（未設定時は既定ID）
 *   UNADVERTISED_ASIN_OFFER_COUNT_ENABLED … 未設定=ON（falseでセラー数APIスキップ）
 *   UNADVERTISED_ASIN_OFFER_SLEEP_MS … 既定 3000（batch間隔）
 *   UNADVERTISED_ASIN_OFFER_429_WAIT_MS … 既定 90000
 *   UNADVERTISED_ASIN_OFFER_429_RETRIES … 既定 5
 *   SPAPI_LWA_CLIENT_ID / SPAPI_LWA_CLIENT_SECRET / SPAPI_REFRESH_TOKEN
 *   SPAPI_SELLER_ID / SPAPI_MARKETPLACE_ID / SPAPI_ENDPOINT
 */

var UA_ENABLED_PROP_ = 'UNADVERTISED_ASIN_ENABLED';
var UA_BUDGET_PROP_ = 'UNADVERTISED_ASIN_DAILY_BUDGET';
var UA_BID_PROP_ = 'UNADVERTISED_ASIN_DEFAULT_BID';
var UA_MAX_LIST_PROP_ = 'UNADVERTISED_ASIN_MAX_LISTINGS';
var UA_CSV_AGE_PROP_ = 'UNADVERTISED_ASIN_CSV_MAX_AGE_DAYS';
var UA_ID_AUTO_PROP_ = 'UNADVERTISED_ASIN_ID_AUTO_FROM_BULK';
var UA_MASTER_SS_PROP_ = 'LISTING_MASTER_SS_ID';
/** tools/common_ss_view/registry.json 出品SSと同値 */
var UA_MASTER_SS_DEFAULT_ = '1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28';
var UA_MASTER_SHEET_ = '▼商品マスタ(人間作業用)';

var UA_SHEET_LIST_ = '未広告ASIN一覧';
/** 一覧更新直前のスナップショット（合否の前後比較用） */
var UA_SHEET_LIST_PREV_ = '未広告ASIN一覧_前回';
var UA_SHEET_CSV_REF_ = '広告中商品_参照';
var UA_SHEET_SETTINGS_ = '未広告ASIN_設定';
var UA_SHEET_ID_ = 'ID参照用';
var UA_ADV_FOLDER_NAME_ = '📥広告中の商品レポート';
var UA_BULK_ID_FOLDER_NAME_ = '📥一括キャンペーンDL';

var UA_SPAPI_LWA_CLIENT_ID_ = 'SPAPI_LWA_CLIENT_ID';
var UA_SPAPI_LWA_CLIENT_SECRET_ = 'SPAPI_LWA_CLIENT_SECRET';
var UA_SPAPI_REFRESH_TOKEN_ = 'SPAPI_REFRESH_TOKEN';
var UA_SPAPI_SELLER_ID_ = 'SPAPI_SELLER_ID';
var UA_SPAPI_MARKETPLACE_ID_ = 'SPAPI_MARKETPLACE_ID';
var UA_SPAPI_ENDPOINT_ = 'SPAPI_ENDPOINT';
var UA_SPAPI_LWA_URL_ = 'https://api.amazon.com/auth/o2/token';
var UA_SPAPI_DEFAULT_ENDPOINT_ = 'https://sellingpartnerapi-fe.amazon.com';
var UA_SPAPI_DEFAULT_MARKETPLACE_ = 'A1VC38T7YXB528';
var UA_SPAPI_UA_ = 'OctasUnadvertisedAsin/1.13 (Language=GoogleAppsScript)';

var UA_BID_STRATEGY_ = '動的な入札 - ダウンのみ';
var UA_NEG_MATCH_ = '除外キーワードの完全一致';
var UA_POS_MATCH_EXACT_ = '完全一致';
/** アカウント実DLのエンティティ表記（商品広告≠プロダクト広告） */
var UA_ENT_PRODUCT_AD_ = 'プロダクト広告';
var UA_ENT_CAMP_NEG_KW_ = 'キャンペーン除外キーワード';
var UA_ENT_AG_NEG_KW_ = '除外キーワード';
var UA_ENT_BID_ADJ_ = '入札調整';
/** 掲載枠調整0%（実DLと同じ4枠） */
var UA_PLACEMENTS_ = [
  '掲載枠の商品ページ',
  '掲載枠のトップ',
  '掲載枠のその他の検索結果',
  'Amazonビジネスの掲載枠'
];

/**
 * Amazon Ads 公式一括テンプレ（AdvertisingBulksheetTemplate-seller.xlsx）
 * シート「スポンサープロダクト広告キャンペーン」のヘッダー31列と同一。
 */
var UA_SP_BULK_SHEET_NAME_ = 'スポンサープロダクト広告キャンペーン';
var UA_SP_BULK_HEADERS_ = [
  'プロダクト', 'エンティティ', '操作',
  'キャンペーンID', '広告グループID', 'ポートフォリオID',
  '広告ID', 'キーワードID', '商品ターゲティングID',
  'キャンペーン名', '広告グループ名',
  '開始日', '終了日', 'ターゲティングのタイプ', 'ステータス', '1日の予算',
  'SKU', '広告グループの入札額の初期値', '入札額',
  'キーワードテキスト', 'ネイティブ言語のキーワード', 'ネイティブ言語の地域',
  'マッチタイプ', '入札戦略', '掲載枠', '割合',
  '商品ターゲティング式', 'オーディエンスID', '購入者のコホート割合',
  'お客様コホートタイプ', 'サイト'
];

var UA_LIST_HEADERS_ = [
  'レ点',
  '親ASIN',
  '親ブロック／キャンペーン予定名',
  'マニュアル予定名',
  '広告グループ予定名',
  '掲載方針',
  '既存オート名',
  '既存マニュアル名',
  'ASIN',
  'SKU',
  '出品開始日',
  'ブランド',
  '商品名',
  '商品名ベース',
  'バリ名',
  '広告状態',
  '新品セラー数',
  '備考'
];

/** 新品オファー数 ≥ この値なら相乗り（複数セラー） */
var UA_MULTI_SELLER_MIN_ = 2;
/** キャンペーン／AG名の硬上限（Manual_接頭辞は含まない） */
var UA_NAME_MAX_ = 128;
var UA_BRAND_NONE_ = '（なし）';

/**
 * メニュー: 一覧更新（SP-API × 広告中CSV × 出品マスタ）
 */
function menuUnadvertisedAsinRefreshList() {
  var runId = 'UA_LIST_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
  Logger.log('[未広告ASIN] runId=' + runId + ' state=RUNNING step=refreshList');
  try {
    if (!uaEnabled_()) {
      throw new Error(UA_ENABLED_PROP_ + '=false のため停止。Script Properties で true にするかキー削除。');
    }
    uaEnsureSettingsSheet_();
    // 前回親キャッシュはスナップショット前に読む（直後に現一覧で上書きするため）
    var prevParentMap = uaLoadPrevParentMap_();
    uaSnapshotListSheetBeforeRefresh_();
    var idRefresh = uaRefreshIdSheetFromBulkFolder_();
    var adv = uaLoadAdvertisedAsinSet_();
    var listings = uaFetchListingsFromSpapi_();
    var masterMap = uaLoadMasterByAsin_();
    var parentResolved = uaResolveParentAsins_(listings, prevParentMap);
    var parentMap = parentResolved.map || parentResolved;
    var parentSource = parentResolved.source || {};
    var idNames = uaLoadExistingCampaignNamesFromIdSheet_();
    var offerMap = uaFetchNewOfferCountsByAsin_(listings);
    var rows = uaBuildListRows_(listings, parentMap, adv, masterMap, idNames, offerMap, parentSource);
    uaWriteListSheet_(rows);
    uaWriteCsvRefSheet_(adv.headers, adv.sampleRows, adv.fileName, adv.fileDate);
    var adStateIdx = UA_LIST_HEADERS_.indexOf('広告状態');
    var noneCount = rows.filter(function (r) { return r[adStateIdx] === 'なし'; }).length;
    var multiN = 0;
    var offerOk = 0;
    Object.keys(offerMap).forEach(function (a) {
      if (offerMap[a] && offerMap[a].ok) {
        offerOk++;
        if (offerMap[a].count >= UA_MULTI_SELLER_MIN_) multiN++;
      }
    });
    var src = parentResolved.source || {};
    var srcList = 0;
    var srcCat = 0;
    var srcSelf = 0;
    Object.keys(src).forEach(function (k) {
      if (src[k] === 'listings') srcList++;
      else if (src[k] === 'catalog') srcCat++;
      else srcSelf++;
    });
    var msg = '一覧を更新しました。\n' +
      '出品=' + listings.length + '／広告ありASIN=' + adv.asinSet.length +
      '／広告なし行=' + noneCount +
      '\n親解決 Listings=' + srcList + '／Catalog=' + srcCat +
      '／前回親=' + (function () {
        var n = 0;
        Object.keys(src).forEach(function (k) { if (src[k] === 'prevCache') n++; });
        return n;
      })() +
      '／子自身=' + srcSelf +
      '\nセラー数取得=' + offerOk + 'ASIN／複数セラー(≥' + UA_MULTI_SELLER_MIN_ + ')=' + multiN +
      '\n前回一覧→シート「' + UA_SHEET_LIST_PREV_ + '」（合否の前後比較用）' +
      '\nCSV: ' + adv.fileName +
      '\nID参照用: ' + (idRefresh && idRefresh.summary ? idRefresh.summary : '（未実行）') +
      '\nrunId=' + runId;
    Logger.log('[未広告ASIN] runId=' + runId + ' state=DONE ' + msg.replace(/\n/g, ' | '));
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log('[未広告ASIN] runId=' + runId + ' state=FAILED err=' + (e && e.message ? e.message : e));
    SpreadsheetApp.getUi().alert('一覧更新に失敗しました。\n' + (e && e.message ? e.message : e));
  }
}

/**
 * メニュー: レ点行 → 親キャンペーン／子AGバルク
 */
function menuUnadvertisedAsinGenerateBulk() {
  var runId = 'UA_BULK_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
  Logger.log('[未広告ASIN] runId=' + runId + ' state=RUNNING step=generateBulk');
  try {
    if (!uaEnabled_()) {
      throw new Error(UA_ENABLED_PROP_ + '=false のため停止。');
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(UA_SHEET_LIST_);
    if (!sh) throw new Error('シート「' + UA_SHEET_LIST_ + '」がありません。先に一覧更新を実行してください。');
    var data = sh.getDataRange().getValues();
    if (data.length < 2) throw new Error('一覧が空です。');
    var header = data[0].map(function (h) { return String(h || '').trim(); });
    var col = uaHeaderIndexMap_(header);
    uaRequireCols_(col, ['レ点', '親ASIN', 'ASIN', 'SKU', '広告状態', '親ブロック／キャンペーン予定名', 'マニュアル予定名', '広告グループ予定名', '掲載方針']);

    var targets = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!uaIsChecked_(r[col['レ点']])) continue;
      var asin = String(r[col['ASIN']] || '').trim().toUpperCase();
      var sku = String(r[col['SKU']] || '').trim();
      var parentAsin = String(r[col['親ASIN']] || '').trim().toUpperCase();
      var adState = String(r[col['広告状態']] || '').trim();
      var autoName = String(r[col['親ブロック／キャンペーン予定名']] || '').trim();
      var manName = String(r[col['マニュアル予定名']] || '').trim();
      var agName = String(r[col['広告グループ予定名']] || '').trim();
      var policy = String(r[col['掲載方針']] || '').trim();
      var existAuto = col['既存オート名'] >= 0 ? String(r[col['既存オート名']] || '').trim() : '';
      var existMan = col['既存マニュアル名'] >= 0 ? String(r[col['既存マニュアル名']] || '').trim() : '';
      var nameBase = col['商品名ベース'] >= 0 ? String(r[col['商品名ベース']] || '').trim() : '';
      var note = col['備考'] >= 0 ? String(r[col['備考']] || '').trim() : '';
      if (!asin || !sku) throw new Error((i + 1) + '行目: ASINまたはSKUが空です。');
      if (!parentAsin) throw new Error((i + 1) + '行目: 親ASINが空です。一覧を更新してください。');
      if (adState === 'あり') {
        throw new Error((i + 1) + '行目: 広告状態が「あり」です。レ点を外すか一覧を更新してください。');
      }
      if (note.indexOf('名前衝突') !== -1) {
        throw new Error((i + 1) + '行目: 名前衝突のためスキップ対象です。');
      }
      if (!autoName || !manName || !agName) {
        throw new Error((i + 1) + '行目: キャンペーン名／マニュアル名／広告グループ名が空です。');
      }
      if (!nameBase) nameBase = autoName;
      targets.push({
        asin: asin,
        sku: sku,
        parentAsin: parentAsin,
        autoName: autoName,
        manName: manName,
        agName: agName,
        policy: policy,
        existAuto: existAuto,
        existMan: existMan,
        nameBase: nameBase
      });
    }
    if (!targets.length) {
      throw new Error('レ点が付いた行がありません。広告なしの行にレ点を付けてから実行してください。');
    }

    var ui = SpreadsheetApp.getUi();
    var conf = ui.alert(
      '未広告ASINバルク生成',
      targets.length + '件（子）を親キャンペーン／子AG仕様で作成します。\n続けますか？',
      ui.ButtonSet.YES_NO
    );
    if (conf !== ui.Button.YES) {
      Logger.log('[未広告ASIN] runId=' + runId + ' state=CANCELLED');
      return;
    }

    var settings = uaReadSettings_();
    var budget = settings.budget;
    var bid = settings.bid;
    var startDate = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');
    var allChildrenByParent = uaCollectAllChildrenByParentFromList_(data, col);
    var idRefresh = uaRefreshIdSheetFromBulkFolder_();
    var idLookup = uaLoadIdLookupMaps_();
    Logger.log('[未広告ASIN] runId=' + runId +
      ' idRefresh=' + (idRefresh && idRefresh.summary ? idRefresh.summary : '') +
      ' idCamps=' + Object.keys((idLookup && idLookup.campaignMap) || {}).length +
      ' idAgs=' + Object.keys((idLookup && idLookup.adGroupMap) || {}).length +
      ' idPa=' + Object.keys((idLookup && idLookup.existingPa) || {}).length);
    uaLogIdSheetNameProbes_(idLookup && idLookup.campaignMap);
    var built = uaBuildBulkRowsParentChild_(targets, allChildrenByParent, idLookup, budget, bid, startDate);
    var bulkRows = built.rows;
    if (!bulkRows || bulkRows.length < 2) {
      throw new Error(
        '出力できる行がありません。' +
        (built.skipped && built.skipped.length
          ? '\nスキップ: ' + built.skipped.join('\n')
          : '') +
        '\n既存キャンペーンは Drive「' + UA_BULK_ID_FOLDER_NAME_ +
        '」に一括操作のカスタムシートを置いてから再実行してください。'
      );
    }
    var fileName = uaSaveBulkXlsx_(bulkRows, runId);
    Logger.log('[未広告ASIN] runId=' + runId + ' state=DONE targets=' + targets.length +
      ' rows=' + (bulkRows.length - 1) + ' skipped=' + (built.skipped || []).length +
      ' file=' + fileName +
      ' idRefresh=' + (idRefresh && idRefresh.summary ? idRefresh.summary : ''));
    var skipMsg = (built.skipped && built.skipped.length)
      ? '\n\nスキップした親（ID参照用に数値IDなし）:\n' + built.skipped.slice(0, 8).join('\n') +
        (built.skipped.length > 8 ? '\n…他' + (built.skipped.length - 8) + '件' : '') +
        '\n→ 「' + UA_BULK_ID_FOLDER_NAME_ + '」のファイルを取り直して再実行'
      : '';
    ui.alert('バルクを保存しました。\n子件数=' + targets.length + '\n行数=' + (bulkRows.length - 1) +
      '\n' + fileName +
      '\nID参照用: ' + (idRefresh && idRefresh.summary ? idRefresh.summary : '（未実行）') +
      '\nAmazon広告の一括操作からアップロードしてください。\nrunId=' + runId + skipMsg);
  } catch (e) {
    Logger.log('[未広告ASIN] runId=' + runId + ' state=FAILED err=' + (e && e.message ? e.message : e));
    SpreadsheetApp.getUi().alert('バルク生成に失敗しました。\n' + (e && e.message ? e.message : e));
  }
}

// --- 一覧組み立て ---

/**
 * @param {Object[]} listings
 * @param {Object} parentMap asin -> parentAsin
 * @param {Object} adv uaLoadAdvertisedAsinSet_ 戻り
 * @param {Object} masterMap asin -> {nameBase, varName}
 * @param {string[]} idNames ID参照用のキャンペーン名
 * @param {Object} offerMap asin -> {ok:boolean, count:number, err?:string}
 * @param {Object=} parentSource asin -> listings|catalog|self
 */
function uaBuildListRows_(listings, parentMap, adv, masterMap, idNames, offerMap, parentSource) {
  offerMap = offerMap || {};
  parentSource = parentSource || {};
  var advMap = {};
  for (var a = 0; a < adv.asinSet.length; a++) advMap[adv.asinSet[a]] = true;

  var childrenByParent = {};
  for (var i = 0; i < listings.length; i++) {
    var asin0 = String(listings[i].asin || '').trim().toUpperCase();
    if (!asin0) continue;
    var p0 = parentMap[asin0] || asin0;
    if (!childrenByParent[p0]) childrenByParent[p0] = [];
    childrenByParent[p0].push(asin0);
  }

  var parentHasAd = {};
  var parentExistAuto = {};
  var parentExistMan = {};
  Object.keys(childrenByParent).forEach(function (p) {
    var kids = childrenByParent[p];
    var has = false;
    var autos = [];
    var mans = [];
    for (var k = 0; k < kids.length; k++) {
      if (advMap[kids[k]]) {
        has = true;
        var camps = adv.campaignsByAsin[kids[k]] || [];
        for (var c = 0; c < camps.length; c++) {
          var cn = camps[c];
          if (/^Manual_/i.test(cn)) mans.push(cn);
          else autos.push(cn);
        }
      }
    }
    parentHasAd[p] = has;
    parentExistAuto[p] = uaPickPreferredName_(autos);
    parentExistMan[p] = uaPickPreferredName_(mans);
  });

  var nameMap = {};
  for (var n = 0; n < idNames.length; n++) nameMap[uaNormalizeName_(idNames[n])] = true;
  Object.keys(adv.campaignNameSet || {}).forEach(function (cn) {
    nameMap[uaNormalizeName_(cn)] = true;
  });

  var parentNewest = {};
  for (var s0 = 0; s0 < listings.length; s0++) {
    var a0 = String(listings[s0].asin || '').trim().toUpperCase();
    if (!a0) continue;
    var p0b = parentMap[a0] || a0;
    var d0 = listings[s0].createdDate || '';
    if (!parentNewest[p0b] || d0 > parentNewest[p0b]) parentNewest[p0b] = d0;
  }
  var sorted = listings.slice().sort(function (x, y) {
    var ax = String(x.asin || '').toUpperCase();
    var ay = String(y.asin || '').toUpperCase();
    var px = parentMap[ax] || ax;
    var py = parentMap[ay] || ay;
    var nx = parentNewest[px] || '';
    var ny = parentNewest[py] || '';
    if (nx !== ny) return nx < ny ? 1 : -1;
    if (px !== py) return String(px).localeCompare(String(py));
    var dx = x.createdDate || '';
    var dy = y.createdDate || '';
    if (dx === dy) return String(ay).localeCompare(String(ax));
    return dx < dy ? 1 : -1;
  });

  var rows = [];
  var parentCampaignBase = {};
  var parentCampCollision = {};
  var usedCampNames = {};
  Object.keys(nameMap).forEach(function (n) { usedCampNames[n] = true; });
  var usedAgNamesByParent = {};

  for (var j = 0; j < sorted.length; j++) {
    var it = sorted[j];
    var asin = String(it.asin || '').trim().toUpperCase();
    if (!asin) continue;
    var parentAsin = parentMap[asin] || asin;
    var title = String(it.title || '').trim();
    var sku = String(it.sku || '').trim();
    var m = uaLookupMaster_(masterMap, asin, sku);
    var nameBase = String(m.nameBase || '').trim();
    var varName = String(m.varName || '').trim();
    var brand = uaResolveBrandLabel_(it.brand, m.makerName);
    var offer = offerMap[asin] || {};
    var offerKnown = !!offer.ok;
    var offerCount = offerKnown ? Number(offer.count) : '';
    var isAinori = false;
    var ainoriReason = '';
    if (offerKnown) {
      isAinori = offerCount >= UA_MULTI_SELLER_MIN_;
      if (isAinori) ainoriReason = '相乗り(セラー' + offerCount + ')';
    } else {
      isAinori = !!m.isAinori || uaGuessAinoriSku_(sku, asin);
      if (isAinori) ainoriReason = m.isAinori ? '相乗り(マスタSKU)' : '相乗り(SKU形式)';
    }
    var notes = [];
    if (!nameBase) {
      nameBase = uaSanitizeName_(title || asin);
      notes.push('商品名ベース欠落→タイトル代用');
    }
    if (!offerKnown) notes.push('セラー数未取得');
    if (isAinori) notes.push(ainoriReason || '相乗り');

    var policy = parentHasAd[parentAsin] ? '既存に追加' : '新規ペア';
    var campBase;
    if (parentHasAd[parentAsin] && parentExistAuto[parentAsin]) {
      campBase = parentExistAuto[parentAsin];
      parentCampaignBase[parentAsin] = campBase;
    } else {
      if (!parentCampaignBase[parentAsin]) {
        // 相乗りはマスタ商品名ベースが自社と同一になりやすい → SP-API商品名（出品タイトル）を使う
        var stem = uaCampaignStem_(nameBase, title, isAinori);
        if (isAinori && title) notes.push('相乗り名=出品タイトル');
        var reg = uaRegisterReadableName_(stem, usedCampNames);
        parentCampaignBase[parentAsin] = reg.name;
        parentCampCollision[parentAsin] = reg.collision;
        if (!reg.collision) {
          usedCampNames[uaNormalizeName_('Manual_' + reg.name)] = true;
        }
      }
      campBase = parentCampaignBase[parentAsin];
      if (parentCampCollision[parentAsin]) notes.push('名前衝突');
    }

    var autoName = parentHasAd[parentAsin] && parentExistAuto[parentAsin]
      ? parentExistAuto[parentAsin]
      : campBase;
    var manName = parentHasAd[parentAsin] && parentExistMan[parentAsin]
      ? parentExistMan[parentAsin]
      : ('Manual_' + campBase);
    if (!/^Manual_/i.test(manName) && !(parentHasAd[parentAsin] && parentExistMan[parentAsin])) {
      manName = 'Manual_' + campBase;
    }

    if (!usedAgNamesByParent[parentAsin]) usedAgNamesByParent[parentAsin] = {};
    var agReg = uaBuildAdGroupName_(campBase, varName, title, usedAgNamesByParent[parentAsin]);
    var agName = agReg.name;
    if (agReg.collision) notes.push('AG名衝突');

    var adState = advMap[asin] ? 'あり' : 'なし';
    if (!sku) notes.push('SKU空');
    if (!parentAsin || parentAsin === asin || parentSource[asin] === 'self') {
      notes.push('親ASINフォールバック=子');
    } else if (parentSource[asin] === 'listings') {
      notes.push('親=Listings');
    } else if (parentSource[asin] === 'catalog' || parentSource[asin] === 'catalogRetry') {
      notes.push('親=Catalog');
    } else if (parentSource[asin] === 'prevCache') {
      notes.push('親=前回一覧');
    }
    if (policy === '新規ペア') {
      if (nameMap[uaNormalizeName_(autoName)] || nameMap[uaNormalizeName_(manName)]) {
        if (notes.indexOf('名前衝突') === -1) notes.push('名前衝突');
      }
    }
    if (it.status) {
      var st = String(it.status).toUpperCase();
      if (st.indexOf('DELETED') !== -1 || st.indexOf('INACTIVE') !== -1) notes.push('状態:' + it.status);
    }
    var checked = (adState === 'なし' && notes.indexOf('名前衝突') === -1 && sku && !isAinori);
    rows.push([
      !!checked,
      parentAsin,
      autoName,
      manName,
      agName,
      policy,
      parentExistAuto[parentAsin] || '',
      parentExistMan[parentAsin] || '',
      asin,
      sku,
      it.createdDate || '',
      brand,
      title,
      nameBase,
      varName,
      adState,
      offerKnown ? offerCount : '',
      notes.join('／')
    ]);
  }
  return rows;
}

function uaPickPreferredName_(arr) {
  if (!arr || !arr.length) return '';
  var seen = {};
  var uniq = [];
  for (var i = 0; i < arr.length; i++) {
    var s = String(arr[i] || '').trim();
    if (!s || seen[s]) continue;
    seen[s] = true;
    uniq.push(s);
  }
  if (!uniq.length) return '';
  uniq.sort(function (a, b) { return a.length - b.length; });
  return uniq[0];
}

/** キャンペーン名の種（商品名ベース優先。短名強制なし） */
function uaCampaignNameFromBase_(nameBase) {
  return uaClipName_(uaSanitizeName_(nameBase || ''), UA_NAME_MAX_);
}

/**
 * キャンペーン名の種。
 * 相乗りはマスタ商品名ベースが自社バリエと同一になりやすいため、出品タイトル（SP-API）を使う。
 */
function uaCampaignStem_(nameBase, title, isAinori) {
  if (isAinori) {
    var t = uaSanitizeName_(title || '');
    if (t) return uaClipName_(t, UA_NAME_MAX_);
  }
  var b = uaSanitizeName_(nameBase || '');
  if (b) return uaClipName_(b, UA_NAME_MAX_);
  return uaClipName_(uaSanitizeName_(title || ''), UA_NAME_MAX_);
}

/**
 * 可読名を登録。ASIN末尾は付けない。既存と同一なら collision=true（名前はそのまま）。
 * @return {{name:string, collision:boolean}}
 */
function uaRegisterReadableName_(stem, usedMap) {
  var name = uaClipName_(uaSanitizeName_(stem || ''), UA_NAME_MAX_);
  if (!name) name = '未設定';
  var norm = uaNormalizeName_(name);
  if (!usedMap[norm]) {
    usedMap[norm] = true;
    return { name: name, collision: false };
  }
  return { name: name, collision: true };
}

/**
 * AG名。バリ名・パック表記で差が付く場合だけ足す。ASIN末尾は付けない。
 * @return {{name:string, collision:boolean}}
 */
function uaBuildAdGroupName_(campBase, varName, title, usedMap) {
  var base = uaClipName_(uaSanitizeName_(campBase || ''), UA_NAME_MAX_);
  if (!base) base = 'AG';
  var candidates = [base];
  var v = uaSanitizeName_(varName || '');
  if (v && base.indexOf(v) === -1) {
    candidates.push(uaClipName_(uaSanitizeName_(base + ' ' + v), UA_NAME_MAX_));
  }
  var pack = uaExtractPackHint_(title);
  if (pack && base.indexOf(pack) === -1) {
    candidates.push(uaClipName_(uaSanitizeName_(base + ' ' + pack), UA_NAME_MAX_));
  }
  for (var i = 0; i < candidates.length; i++) {
    var cand = candidates[i];
    var cn = uaNormalizeName_(cand);
    if (!usedMap[cn]) {
      usedMap[cn] = true;
      return { name: cand, collision: false };
    }
  }
  return { name: candidates[candidates.length - 1] || base, collision: true };
}

/** タイトルから「N缶」「N個」等の短い差を拾う */
function uaExtractPackHint_(title) {
  var t = String(title || '');
  var m = t.match(/(\d+\s*缶)/) || t.match(/(\d+\s*個)/) || t.match(/(\d+\s*食)/) ||
    t.match(/\((\d+食[^)]*)\)/) || t.match(/\((\d+個[^)]*)\)/);
  if (!m) return '';
  return uaSanitizeName_(m[1] || m[0]).substring(0, 24);
}

function uaClipName_(s, maxLen) {
  var t = String(s || '');
  var lim = maxLen != null ? maxLen : UA_NAME_MAX_;
  if (t.length > lim) t = t.substring(0, lim);
  return t;
}

function uaSanitizeName_(s) {
  var t = String(s || '').replace(/\s+/g, ' ').trim();
  t = t.replace(/[\\\/:\*\?"<>\|]/g, '');
  if (t.length > UA_NAME_MAX_) t = t.substring(0, UA_NAME_MAX_);
  return t;
}

/**
 * マニュアル初回KW用。Amazonの「Keyword is invalid」を避けるため短く整える。
 * 括弧・×等を落とし、語数≤10・文字数≤80。
 */
function uaSanitizeSeedKeyword_(s) {
  var t = String(s || '');
  t = t.replace(/（[^）]*）/g, ' ').replace(/\([^)]*\)/g, ' ');
  t = t.replace(/[×✕✖･・\/／|｜,，、;；:：\[\]【】「」『』"'`]/g, ' ');
  t = t.replace(/[\\\/\*\?"<>]/g, ' ');
  t = t.replace(/\s+/g, ' ').trim();
  if (!t) return '';
  var words = t.split(' ').filter(function (w) { return !!w; });
  if (words.length > 10) words = words.slice(0, 10);
  t = words.join(' ');
  if (t.length > 80) t = t.substring(0, 80).replace(/\s+\S*$/, '').trim();
  if (t.length < 2) return '';
  return t;
}

function uaResolveBrandLabel_(spapiBrand, makerName) {
  var b = String(spapiBrand || '').trim();
  if (b) return b;
  var m = String(makerName || '').trim();
  if (m) return m;
  return UA_BRAND_NONE_;
}

/**
 * 相乗りSKUの機械判定（マスタの Amazon相乗りSKU が取れないときの補助）
 * 正: 自社カタログSKUは JAN 起点などで ASIN を含めないことが多い。
 * 相乗りSKUは ASIN を埋め込むことが多い（例: Sakae-B0C52R16P1-S / -B0CY4H1TJD-S / sanky-B0…-19as13）
 */
function uaGuessAinoriSku_(sku, asin) {
  var s = String(sku || '');
  if (!s) return false;
  var a = String(asin || '').trim().toUpperCase();
  if (a && s.toUpperCase().indexOf(a) !== -1) return true;
  if (/B0[A-Z0-9]{8}/i.test(s)) return true;
  if (/B0[A-Z0-9]{8}.*as\d+/i.test(s)) return true;
  if (/as\d{1,3}$/i.test(s) && /B0[A-Z0-9]{8}/i.test(s)) return true;
  return false;
}

function uaLookupMaster_(masterMap, asin, sku) {
  // SKU優先（相乗り判定はSKU単位）。ASINヒットは名前だけ借り、isAinoriは付けない
  if (sku && masterMap.bySku && masterMap.bySku[sku]) return masterMap.bySku[sku];
  if (asin && masterMap.byAsin && masterMap.byAsin[asin]) {
    var rec = masterMap.byAsin[asin];
    return {
      nameBase: rec.nameBase || '',
      varName: rec.varName || '',
      makerName: rec.makerName || '',
      isAinori: false
    };
  }
  return { nameBase: '', varName: '', makerName: '', isAinori: false };
}

function uaNormalizeName_(s) {
  s = String(s || '');
  if (typeof s.normalize === 'function') {
    try { s = s.normalize('NFKC'); } catch (ignoreNfkc) {}
  }
  return s.replace(/\s+/g, '').replace(/　/g, '');
}

/**
 * 新規作成用。× など拒否文字だけ直す。
 * 括弧の全角半角は触らない（NFKCすると既存キャンペーンと別名扱いになり already exists が増える）。
 */
function uaSanitizeAmzBulkName_(s) {
  s = String(s || '');
  s = s.replace(/[×✕✖✗☓✕]/g, 'x');
  s = s.replace(/[<>:"/\\|?*]/g, '');
  return s.replace(/\s+/g, ' ').trim();
}

/** キャンペーン名→数値ID（元名と×除去名の両方で探す） */
function uaLookupCampId_(campaignMap, name) {
  if (!name || !campaignMap) return '';
  var candidates = [String(name), uaSanitizeAmzBulkName_(name)];
  for (var i = 0; i < candidates.length; i++) {
    var id = campaignMap[uaNormalizeName_(candidates[i])] || '';
    if (uaLooksLikeAmzEntityId_(id)) return String(id);
  }
  return '';
}

/**
 * ID照合の内訳（原因切り分け用）。ヒット時は id、ミス時は近傍キー候補。
 * @return {{id:string, hitHow:string, near:string}}
 */
function uaCampIdLookupDetail_(campaignMap, name) {
  var out = { id: '', hitHow: 'miss', near: '' };
  if (!name || !campaignMap) return out;
  var raw = String(name);
  var san = uaSanitizeAmzBulkName_(raw);
  var keys = [
    { how: 'raw', key: uaNormalizeName_(raw) },
    { how: 'sanitize', key: uaNormalizeName_(san) }
  ];
  for (var i = 0; i < keys.length; i++) {
    var id = campaignMap[keys[i].key] || '';
    if (uaLooksLikeAmzEntityId_(id)) {
      out.id = String(id);
      out.hitHow = keys[i].how;
      return out;
    }
  }
  // 近傍: 正規化名の先頭12文字が一致するキー（最大3）
  var stem = uaNormalizeName_(san || raw).substring(0, 12);
  if (stem.length >= 4) {
    var near = [];
    var mapKeys = Object.keys(campaignMap);
    for (var k = 0; k < mapKeys.length && near.length < 3; k++) {
      if (mapKeys[k].indexOf(stem) !== -1) {
        near.push(mapKeys[k].substring(0, 40) + '=' + String(campaignMap[mapKeys[k]]).substring(0, 18));
      }
    }
    out.near = near.join(';');
  }
  return out;
}

/**
 * レ点親の解決経路を1行で残す（新規作成／昇格／スキップの根拠）。
 */
function uaLogBulkParentResolve_(opt) {
  opt = opt || {};
  Logger.log(
    '[未広告ASIN] bulk_resolve parent=' + (opt.parentAsin || '') +
    ' policy=' + (opt.policy || '') +
    ' path=' + (opt.path || '') +
    ' autoName=' + String(opt.autoName || '').substring(0, 60) +
    ' manName=' + String(opt.manName || '').substring(0, 60) +
    ' autoId=' + (opt.autoId || '') +
    ' manId=' + (opt.manId || '') +
    ' autoHit=' + (opt.autoHit || '') +
    ' manHit=' + (opt.manHit || '') +
    ' autoNear=' + String(opt.autoNear || '').substring(0, 80) +
    ' manNear=' + String(opt.manNear || '').substring(0, 80) +
    ' skus=' + (opt.skuCount || 0)
  );
}

/** ID参照用に特定語が何件あるか（一括DL欠け／名ゆれ切り分け） */
function uaLogIdSheetNameProbes_(campaignMap) {
  campaignMap = campaignMap || {};
  var keys = Object.keys(campaignMap);
  var probes = ['うず巻き', '唐辛子', '七味', '食べるおだし', '麩'];
  var parts = [];
  for (var i = 0; i < probes.length; i++) {
    var p = probes[i];
    var n = 0;
    var sample = '';
    for (var k = 0; k < keys.length; k++) {
      if (keys[k].indexOf(p) === -1) continue;
      n++;
      if (!sample) sample = keys[k].substring(0, 36);
    }
    parts.push(p + '=' + n + (sample ? '(' + sample + ')' : ''));
  }
  Logger.log('[未広告ASIN] id_probe camps=' + keys.length + ' ' + parts.join(' '));
}

function uaIsChecked_(v) {
  return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === '1';
}

/**
 * 更新前の一覧を「未広告ASIN一覧_前回」へ退避（合否の前後比較用）。
 */
function uaSnapshotListSheetBeforeRefresh_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UA_SHEET_LIST_);
  if (!sh || sh.getLastRow() < 1) {
    Logger.log('[未広告ASIN] snapshot skip（一覧が空または無し）');
    return;
  }
  var old = ss.getSheetByName(UA_SHEET_LIST_PREV_);
  if (old) {
    ss.deleteSheet(old);
  }
  var copied = sh.copyTo(ss);
  copied.setName(UA_SHEET_LIST_PREV_);
  Logger.log('[未広告ASIN] snapshot OK -> ' + UA_SHEET_LIST_PREV_ +
    ' rows=' + sh.getLastRow());
}

function uaWriteListSheet_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UA_SHEET_LIST_);
  if (!sh) sh = ss.insertSheet(UA_SHEET_LIST_);
  sh.clear();
  var out = [UA_LIST_HEADERS_].concat(rows);
  sh.getRange(1, 1, out.length, UA_LIST_HEADERS_.length).setValues(out);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).insertCheckboxes();
  }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, Math.min(8, UA_LIST_HEADERS_.length));
}

function uaWriteCsvRefSheet_(headers, sampleRows, fileName, fileDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UA_SHEET_CSV_REF_);
  if (!sh) sh = ss.insertSheet(UA_SHEET_CSV_REF_);
  sh.clear();
  sh.getRange(1, 1, 2, 2).setValues([
    ['元ファイル', fileName || ''],
    ['更新日時', fileDate || '']
  ]);
  if (headers && headers.length) {
    var body = [headers].concat(sampleRows || []);
    var numRows = body.length;
    var numCols = headers.length;
    var normalized = body.map(function (r) {
      var out = [];
      for (var c = 0; c < numCols; c++) out.push(r[c] !== undefined && r[c] !== null ? r[c] : '');
      return out;
    });
    sh.getRange(3, 1, numRows, numCols).setValues(normalized);
  }
}

function uaHeaderIndexMap_(header) {
  var map = {};
  for (var i = 0; i < header.length; i++) {
    map[header[i]] = i;
  }
  return map;
}

function uaRequireCols_(col, names) {
  for (var i = 0; i < names.length; i++) {
    if (col[names[i]] === undefined || col[names[i]] < 0) {
      throw new Error('一覧に列「' + names[i] + '」がありません。一覧更新を実行してください。');
    }
  }
}

// --- 設定シート ---

function uaEnsureSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UA_SHEET_SETTINGS_);
  if (sh && sh.getLastRow() >= 2) return;
  if (!sh) sh = ss.insertSheet(UA_SHEET_SETTINGS_);
  sh.clear();
  sh.getRange(1, 1, 10, 2).setValues([
    ['キー', '値'],
    ['日予算', 1000],
    ['デフォルト入札', 11],
    ['入札戦略', UA_BID_STRATEGY_],
    ['除外にASINを入れる', 'TRUE'],
    ['除外マッチ', UA_NEG_MATCH_],
    ['除外の置き場所', 'キャンペーン＋広告グループ'],
    ['マニュアル接頭辞', 'Manual_'],
    ['マニュアル初回KW', '商品名ベース・完全一致1語'],
    ['相乗り判定', '新品セラー数≥2（SP-API getItemOffersBatch）']
  ]);
}

function uaReadSettings_() {
  uaEnsureSettingsSheet_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UA_SHEET_SETTINGS_);
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[String(data[i][0] || '').trim()] = data[i][1];
  }
  var budget = Number(map['日予算']);
  var bid = Number(map['デフォルト入札']);
  if (!isFinite(budget) || budget <= 0) budget = uaGetNumberProp_(UA_BUDGET_PROP_, 1000);
  if (!isFinite(bid) || bid <= 0) bid = uaGetNumberProp_(UA_BID_PROP_, 11);
  return { budget: budget, bid: bid, strategy: String(map['入札戦略'] || UA_BID_STRATEGY_) };
}

// --- 広告中CSV ---

/**
 * @return {{asinSet:string[], campaignsByAsin:Object, campaignNameSet:Object, headers:string[], sampleRows:string[][], fileName:string, fileDate:string}}
 */
function uaLoadAdvertisedAsinSet_() {
  if (typeof AMZ_FOLDER_ID === 'undefined' || !AMZ_FOLDER_ID) {
    throw new Error('AMZ_FOLDER_ID が未定義です。');
  }
  var parent = DriveApp.getFolderById(AMZ_FOLDER_ID);
  var folders = parent.getFoldersByName(UA_ADV_FOLDER_NAME_);
  if (!folders.hasNext()) {
    throw new Error(
      'フォルダ「' + UA_ADV_FOLDER_NAME_ + '」がありません。\n' +
      'Amazon用Driveフォルダ内に作成し、「広告中の商品」CSVを保存してから再実行してください。'
    );
  }
  var folder = folders.next();
  var files = folder.getFiles();
  var latest = null;
  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName() || '';
    if (!/\.csv$/i.test(name)) continue;
    if (!latest || f.getLastUpdated().getTime() > latest.getLastUpdated().getTime()) latest = f;
  }
  if (!latest) {
    throw new Error(
      '「' + UA_ADV_FOLDER_NAME_ + '」にCSVがありません。\n' +
      'Adsの「広告中の商品」レポートをダウンロードして保存してください。'
    );
  }
  var ageDays = uaGetNumberProp_(UA_CSV_AGE_PROP_, 14);
  var ageMs = Date.now() - latest.getLastUpdated().getTime();
  if (ageMs > ageDays * 24 * 60 * 60 * 1000) {
    var ui = SpreadsheetApp.getUi();
    var ans = ui.alert(
      '広告中CSVが古い可能性があります',
      'ファイル: ' + latest.getName() + '\n最終更新: ' + latest.getLastUpdated() +
      '\n' + ageDays + '日を超えています。このまま続けますか？',
      ui.ButtonSet.YES_NO
    );
    if (ans !== ui.Button.YES) throw new Error('ユーザーが中止しました（CSV古い）。');
  }
  var text = latest.getBlob().getDataAsString('UTF-8');
  if (text.indexOf('\ufffd') !== -1 || text.indexOf('ã') !== -1) {
    try { text = latest.getBlob().getDataAsString('Shift_JIS'); } catch (ignore) {}
  }
  var parsed = uaParseCsv_(text);
  if (!parsed.length) throw new Error('広告中CSVが空です: ' + latest.getName());
  var header = parsed[0].map(function (h) { return String(h || '').trim(); });
  // 親IDはASIN列に使わない（誤検出防止）
  var asinCol = uaFindCol_(header, [
    '広告対象商品ID', 'Advertised Product ID', 'Advertised product ID',
    'ASIN', 'Asin', 'asin',
    '広告掲載商品のASIN', '広告対象ASIN', 'Advertised ASIN', '商品ASIN'
  ]);
  if (asinCol < 0) {
    asinCol = uaFindColAsinLoose_(header);
  }
  if (asinCol < 0) {
    throw new Error('広告中CSVにASIN列が見つかりません。ヘッダ=' + header.slice(0, 12).join('|'));
  }
  var campCol = uaFindCol_(header, [
    'キャンペーン名', 'Campaign Name', 'campaign name', '広告キャンペーン名'
  ]);
  var set = {};
  var list = [];
  var campaignsByAsin = {};
  var campaignNameSet = {};
  for (var i = 1; i < parsed.length; i++) {
    var asin = String(parsed[i][asinCol] || '').trim().toUpperCase();
    if (!(/^B0[A-Z0-9]{8}$/i.test(asin) || /^[A-Z0-9]{10}$/.test(asin))) continue;
    if (!set[asin]) {
      set[asin] = true;
      list.push(asin);
    }
    if (campCol >= 0) {
      var cname = String(parsed[i][campCol] || '').trim();
      // 一覧表示で後ろに付くキャンペーンIDを落とす
      cname = cname.replace(/\s+A[0-9][A-Z0-9]{18,}$/i, '').trim();
      if (cname) {
        if (!campaignsByAsin[asin]) campaignsByAsin[asin] = [];
        campaignsByAsin[asin].push(cname);
        campaignNameSet[cname] = true;
      }
    }
  }
  var sample = parsed.slice(1, Math.min(21, parsed.length));
  Logger.log('[未広告ASIN] CSV asinCol=' + asinCol + ' campCol=' + campCol + ' asins=' + list.length);
  return {
    asinSet: list,
    campaignsByAsin: campaignsByAsin,
    campaignNameSet: campaignNameSet,
    headers: header,
    sampleRows: sample,
    fileName: latest.getName(),
    fileDate: Utilities.formatDate(latest.getLastUpdated(), 'JST', 'yyyy-MM-dd HH:mm')
  };
}

function uaFindCol_(header, names) {
  for (var i = 0; i < header.length; i++) {
    for (var j = 0; j < names.length; j++) {
      if (header[i] === names[j]) return i;
    }
  }
  return -1;
}

/** ASINっぽい列。親ID列は除外 */
function uaFindColAsinLoose_(header) {
  for (var k = 0; k < header.length; k++) {
    var h = String(header[k] || '');
    if (/親/.test(h) && /ID|ASIN/i.test(h)) continue;
    if (/Parent/i.test(h) && /ID|ASIN/i.test(h)) continue;
    if (/ASIN/i.test(h) || /広告対象商品ID/i.test(h) || /Advertised Product ID/i.test(h)) return k;
  }
  return -1;
}

function uaParseCsv_(text) {
  var rows = [];
  var row = [];
  var cur = '';
  var inQ = false;
  var s = String(text || '').replace(/^\uFEFF/, '');
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    if (inQ) {
      if (c === '"') {
        if (s.charAt(i + 1) === '"') { cur += '"'; i++; }
        else inQ = false;
      } else cur += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === ',') { row.push(cur); cur = ''; }
      else if (c === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
      else if (c === '\r') { /* skip */ }
      else cur += c;
    }
  }
  if (cur.length || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

// --- 出品マスタ ---

/**
 * 出品マスタは先頭行がヘッダではない。列は「列項目名」で先頭〜40行を横断ヒット。
 * 親くくりには使わない（商品名ベース・バリ・メーカー名・相乗りSKUのみ）。
 * @return {{byAsin:Object, bySku:Object, cols:Object}}
 */
function uaLoadMasterByAsin_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = String(props.getProperty(UA_MASTER_SS_PROP_) || '').trim() || UA_MASTER_SS_DEFAULT_;
  var byAsin = {};
  var bySku = {};
  var cols = {
    asin: -1, base: -1, varName: -1, childSku: -1, maker: -1, rideSku: -1, dataStart: 0
  };
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var sh = ss.getSheetByName(UA_MASTER_SHEET_);
    if (!sh) {
      Logger.log('[未広告ASIN] マスタシートなし name=' + UA_MASTER_SHEET_ + ' ss=' + ssId);
      return { byAsin: byAsin, bySku: bySku, cols: cols };
    }
    var data = sh.getDataRange().getValues();
    if (data.length < 2) {
      return { byAsin: byAsin, bySku: bySku, cols: cols };
    }

    var nameHits = uaCollectMasterHeaderNameHits_(data, 40);
    var anchorRow = uaFindMasterHeaderRow_(data);
    var headerBottom = anchorRow;
    ['ASINコード', '対象ASIN', 'ASIN', '商品名ベース', '子SKU', 'バリエーション値',
      'Amazon相乗りSKU', '相乗りSKU', 'バリエーション', 'バリ名', 'メーカー名'].forEach(function (nm) {
      var rows = nameHits._rows && nameHits._rows[nm];
      if (!rows) return;
      for (var ri = 0; ri < rows.length; ri++) {
        if (rows[ri] > headerBottom) headerBottom = rows[ri];
      }
    });
    var dataStart = headerBottom >= 0 ? headerBottom + 1 : 1;

    var colAsin = uaPickMasterColByNames_(nameHits, ['ASINコード', '対象ASIN', 'ASIN'], data, dataStart, 'asin');
    var colBase = uaPickMasterColByNames_(nameHits, ['商品名ベース'], data, dataStart, 'any');
    var colVar = uaPickMasterColByNames_(nameHits, ['バリエーション値', 'バリエーション', 'バリ名'], data, dataStart, 'any');
    var colChildSku = uaPickMasterColByNames_(nameHits, ['子SKU'], data, dataStart, 'childSku');
    var colMaker = uaPickMasterColByNames_(nameHits, ['メーカー名'], data, dataStart, 'any');
    var colRideSku = uaPickMasterColByNames_(nameHits, ['Amazon相乗りSKU', '相乗りSKU'], data, dataStart, 'any');

    cols = {
      asin: colAsin,
      base: colBase,
      varName: colVar,
      childSku: colChildSku,
      maker: colMaker,
      rideSku: colRideSku,
      dataStart: dataStart
    };

    if (colAsin < 0 && colChildSku < 0) {
      Logger.log('[未広告ASIN] マスタにASIN/子SKU列なし（列項目名ヒット失敗） ss=' + ssId);
      return { byAsin: byAsin, bySku: bySku, cols: cols };
    }

    var lastBase = '';
    var lastVar = '';
    var lastMaker = '';
    for (var i = dataStart; i < data.length; i++) {
      var nameBase = colBase >= 0 ? String(data[i][colBase] || '').trim() : '';
      var varName = colVar >= 0 ? String(data[i][colVar] || '').trim() : '';
      var makerName = colMaker >= 0 ? String(data[i][colMaker] || '').trim() : '';
      if (nameBase) lastBase = nameBase;
      if (varName) lastVar = varName;
      if (makerName) lastMaker = makerName;
      var asin = colAsin >= 0 ? String(data[i][colAsin] || '').trim().toUpperCase() : '';
      var childSku = colChildSku >= 0 ? String(data[i][colChildSku] || '').trim() : '';
      var rideSku = colRideSku >= 0 ? String(data[i][colRideSku] || '').trim() : '';
      var useBase = nameBase || lastBase;
      var useVar = varName || lastVar;
      var useMaker = makerName || lastMaker;
      if (!asin && !childSku && !rideSku) continue;
      if (asin && (!byAsin[asin] || useBase)) {
        byAsin[asin] = {
          nameBase: useBase, varName: useVar, makerName: useMaker, isAinori: false
        };
      }
      if (childSku) {
        bySku[childSku] = {
          nameBase: useBase, varName: useVar, makerName: useMaker, isAinori: false
        };
      }
      if (rideSku) {
        bySku[rideSku] = {
          nameBase: useBase, varName: useVar, makerName: useMaker, isAinori: true
        };
      }
    }
    Logger.log('[未広告ASIN] master cols asin=' + (colAsin + 1) +
      ' childSku=' + (colChildSku + 1) +
      ' maker=' + (colMaker + 1) +
      ' base=' + (colBase + 1) +
      ' dataStart=' + (dataStart + 1) +
      ' byAsin=' + Object.keys(byAsin).length +
      ' bySku=' + Object.keys(bySku).length +
      ' ss=' + ssId);
  } catch (e) {
    Logger.log('[未広告ASIN] master load FAILED ' + (e && e.message ? e.message : e));
  }
  return { byAsin: byAsin, bySku: bySku, cols: cols };
}

/**
 * 先頭 maxScan 行を走査し、列項目名 → 列index（0始まり）のヒットを集める。
 * @return {Object} name -> number[] ／ 内部キー _rows: name -> rowIndex[]
 */
function uaCollectMasterHeaderNameHits_(data, maxScan) {
  var hits = {};
  var rowsOf = {};
  var lim = Math.min(data.length, maxScan || 40);
  for (var r = 0; r < lim; r++) {
    var row = data[r] || [];
    for (var c = 0; c < row.length; c++) {
      var v = uaNormalizeMasterHeaderName_(row[c]);
      if (!v) continue;
      if (!hits[v]) hits[v] = [];
      if (hits[v].indexOf(c) === -1) hits[v].push(c);
      if (!rowsOf[v]) rowsOf[v] = [];
      if (rowsOf[v].indexOf(r) === -1) rowsOf[v].push(r);
    }
  }
  hits._rows = rowsOf;
  return hits;
}

function uaNormalizeMasterHeaderName_(v) {
  return String(v == null ? '' : v)
    .replace(/^\uFEFF/, '')
    .replace(/^["'\s\u3000]+|["'\s\u3000]+$/g, '')
    .trim();
}

/**
 * 列項目名ヒットから列を1つ選ぶ。親SKU／子SKUはデータ形状で採否。
 * @param {string} kind asin|childSku|parentSku|any
 */
function uaPickMasterColByNames_(nameHits, names, data, dataStart, kind) {
  var cands = [];
  for (var i = 0; i < names.length; i++) {
    var cols = nameHits[names[i]] || [];
    for (var j = 0; j < cols.length; j++) {
      if (cands.indexOf(cols[j]) === -1) cands.push(cols[j]);
    }
  }
  if (!cands.length) return -1;
  if (kind === 'any' || kind === 'asin') {
    return cands[cands.length - 1];
  }
  var best = -1;
  var bestScore = -1;
  for (var k = 0; k < cands.length; k++) {
    var sc = kind === 'parentSku'
      ? uaScoreMasterParentSkuCol_(data, cands[k], dataStart)
      : uaScoreMasterChildSkuCol_(data, cands[k], dataStart);
    Logger.log('[未広告ASIN] masterCol score kind=' + kind +
      ' nameHitCol=' + (cands[k] + 1) + ' score=' + sc.toFixed(3));
    if (sc > bestScore) {
      bestScore = sc;
      best = cands[k];
    }
  }
  var minScore = kind === 'parentSku' ? 0.3 : 0.2;
  if (bestScore < minScore) {
    Logger.log('[未広告ASIN] masterCol REJECT kind=' + kind +
      ' bestCol=' + (best + 1) + ' score=' + bestScore.toFixed(3) +
      ' （列項目名はヒットしたがデータ形状が親子SKUに見えない）');
    return -1;
  }
  return best;
}

/** 親SKU列らしさ: SKU様のうち末尾 -oya / -oyaN の割合 */
function uaScoreMasterParentSkuCol_(data, col, dataStart) {
  var filled = 0;
  var oya = 0;
  var start = Math.max(0, dataStart | 0);
  var lim = Math.min(data.length, start + 800);
  for (var i = start; i < lim; i++) {
    var v = String((data[i] && data[i][col]) || '').trim();
    if (!v || v.indexOf('#REF') !== -1) continue;
    if (v.length > 80 || /\s/.test(v)) continue;
    filled++;
    if (/-oya\d*$/i.test(v)) oya++;
  }
  return filled ? (oya / filled) : 0;
}

/** 子SKU列らしさ: prefix-… 形式かつ末尾 -oya でない割合 */
function uaScoreMasterChildSkuCol_(data, col, dataStart) {
  var filled = 0;
  var good = 0;
  var oya = 0;
  var start = Math.max(0, dataStart | 0);
  var lim = Math.min(data.length, start + 800);
  for (var i = start; i < lim; i++) {
    var v = String((data[i] && data[i][col]) || '').trim();
    if (!v || v.indexOf('#REF') !== -1) continue;
    if (v.length > 80 || /\s/.test(v)) continue;
    filled++;
    if (/-oya\d*$/i.test(v)) oya++;
    if (/^[A-Za-z][A-Za-z0-9]*-.+/.test(v)) good++;
  }
  if (!filled) return 0;
  if (oya > good) return 0;
  return good / filled;
}

/** 出品マスタのヘッダ行（0始まり）。出品GASの ANCHOR=ASINコード に合わせる */
function uaFindMasterHeaderRow_(data) {
  var maxScan = Math.min(data.length, 40);
  var fallback = -1;
  for (var r = 0; r < maxScan; r++) {
    var row = data[r] || [];
    var hasAsin = false;
    var hasBase = false;
    for (var c = 0; c < row.length; c++) {
      var v = uaNormalizeMasterHeaderName_(row[c]);
      if (v === 'ASINコード' || v === '対象ASIN') hasAsin = true;
      if (v === '商品名ベース') hasBase = true;
    }
    if (hasAsin && hasBase) return r;
    if (hasAsin && fallback < 0) fallback = r;
  }
  return fallback;
}

// --- SP-API ---

function uaFetchListingsFromSpapi_() {
  var creds = uaLoadSpapiCreds_();
  var token = uaRequestLwa_(creds.clientId, creds.clientSecret, creds.refreshToken);
  var maxN = Math.floor(uaGetNumberProp_(UA_MAX_LIST_PROP_, 200));
  if (maxN < 1) maxN = 1;
  if (maxN > 2000) maxN = 2000;
  var pageSize = 20;
  var pageToken = '';
  var out = [];
  var page = 0;
  while (out.length < maxN) {
    page++;
    var url = creds.endpoint.replace(/\/$/, '') +
      '/listings/2021-08-01/items/' + encodeURIComponent(creds.sellerId) +
      '?marketplaceIds=' + encodeURIComponent(creds.marketplaceId) +
      '&includedData=summaries,relationships' +
      '&pageSize=' + pageSize +
      '&withStatus=BUYABLE' +
      '&sortBy=createdDate&sortDirection=DESC';
    if (pageToken) url += '&pageToken=' + encodeURIComponent(pageToken);
    Logger.log('[未広告ASIN] SP-API listings page=' + page + ' have=' + out.length);
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'x-amz-access-token': token,
        'user-agent': UA_SPAPI_UA_,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var text = resp.getContentText() || '';
    if (code === 401 || code === 403) {
      throw new Error('SP-API認証エラー HTTP ' + code + '。Script Properties の SPAPI_* を確認してください。本文=' + text.substring(0, 200));
    }
    if (code !== 200) {
      if (page === 1 && (code === 400 || /sortBy|withStatus|InvalidInput|relationships/i.test(text))) {
        Logger.log('[未広告ASIN] listingsパラメータ非対応の可能性。フォールバック body=' + text.substring(0, 160));
        return uaFetchListingsFromSpapiNoSort_(creds, token, maxN);
      }
      throw new Error('SP-API listings HTTP ' + code + ' ' + text.substring(0, 300));
    }
    var body = JSON.parse(text);
    var items = body.items || body.listings || [];
    for (var i = 0; i < items.length && out.length < maxN; i++) {
      var parsed = uaParseListingItem_(items[i]);
      if (parsed) out.push(parsed);
    }
    pageToken = body.pagination && body.pagination.nextToken ? body.pagination.nextToken : '';
    if (!pageToken) break;
    Utilities.sleep(350);
  }
  Logger.log('[未広告ASIN] SP-API listings done count=' + out.length);
  return out;
}

function uaFetchListingsFromSpapiNoSort_(creds, token, maxN) {
  var pageSize = 20;
  var pageToken = '';
  var out = [];
  while (out.length < maxN) {
    var url = creds.endpoint.replace(/\/$/, '') +
      '/listings/2021-08-01/items/' + encodeURIComponent(creds.sellerId) +
      '?marketplaceIds=' + encodeURIComponent(creds.marketplaceId) +
      '&includedData=summaries' +
      '&pageSize=' + pageSize +
      '&withStatus=BUYABLE';
    if (pageToken) url += '&pageToken=' + encodeURIComponent(pageToken);
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'x-amz-access-token': token,
        'user-agent': UA_SPAPI_UA_,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var text = resp.getContentText() || '';
    if (code !== 200) throw new Error('SP-API listings(noSort) HTTP ' + code + ' ' + text.substring(0, 300));
    var body = JSON.parse(text);
    var items = body.items || [];
    for (var i = 0; i < items.length && out.length < maxN; i++) {
      var parsed = uaParseListingItem_(items[i]);
      if (parsed) out.push(parsed);
    }
    pageToken = body.pagination && body.pagination.nextToken ? body.pagination.nextToken : '';
    if (!pageToken) break;
    Utilities.sleep(350);
  }
  return out;
}

function uaParseListingItem_(item) {
  if (!item) return null;
  var sku = String(item.sku || item.sellerSku || '').trim();
  var summaries = item.summaries || [];
  var sum = summaries.length ? summaries[0] : {};
  var asin = String(sum.asin || item.asin || '').trim().toUpperCase();
  var title = String(sum.itemName || sum.title || '').trim();
  var created = String(sum.createdDate || sum.created_at || item.createdDate || '').trim();
  if (created.indexOf('T') !== -1) created = created.substring(0, 10);
  var status = '';
  if (sum.status) status = Array.isArray(sum.status) ? sum.status.join(',') : String(sum.status);
  else if (item.status) status = String(item.status);
  var brand = String(sum.brandName || sum.brand || '').trim();
  if (!asin && !sku) return null;
  return {
    asin: asin,
    sku: sku,
    title: title,
    brand: brand,
    createdDate: created,
    status: status,
    relationships: item.relationships || []
  };
}

/**
 * 子→親ASIN解決。まとめキーは常に Amazon 親ASIN（Q1）。
 * 順: Listings → Catalogバッチ → Catalog単発再試行 → 前回一覧キャッシュ。
 * @param {Object=} prevParentMap childAsin -> parentAsin（未広告ASIN一覧_前回）
 * @return {{map:Object, source:Object}}
 */
function uaResolveParentAsins_(listings, prevParentMap) {
  prevParentMap = prevParentMap || {};
  var map = {};
  var source = {};
  var parentSkuOf = {};
  var asins = [];

  for (var i = 0; i < listings.length; i++) {
    var a0 = String(listings[i].asin || '').trim().toUpperCase();
    if (!a0) continue;
    if (!map[a0]) {
      map[a0] = a0;
      source[a0] = 'self';
      asins.push(a0);
    }

    var fromList = uaExtractListingParentRef_(listings[i].relationships);
    if (fromList.parentAsin) {
      map[a0] = fromList.parentAsin;
      source[a0] = 'listings';
    }
    if (fromList.parentSku) {
      parentSkuOf[a0] = fromList.parentSku;
    }
  }

  var byParentSku = {};
  Object.keys(parentSkuOf).forEach(function (asin) {
    var ps = parentSkuOf[asin];
    if (!byParentSku[ps]) byParentSku[ps] = [];
    byParentSku[ps].push(asin);
  });

  var creds = null;
  var token = '';
  var skuAsinCache = {};
  try {
    creds = uaLoadSpapiCreds_();
    token = uaRequestLwa_(creds.clientId, creds.clientSecret, creds.refreshToken);
  } catch (eCred) {
    Logger.log('[未広告ASIN] resolveParent creds FAILED ' +
      (eCred && eCred.message ? eCred.message : eCred));
  }

  Object.keys(byParentSku).forEach(function (psku) {
    var group = byParentSku[psku];
    var known = '';
    for (var g = 0; g < group.length; g++) {
      if (map[group[g]] && map[group[g]] !== group[g]) {
        known = map[group[g]];
        break;
      }
    }
    if (!known && creds && token) {
      known = uaFetchAsinForSellerSku_(creds, token, psku, skuAsinCache);
    }
    if (!known) return;
    for (var g2 = 0; g2 < group.length; g2++) {
      var ca = group[g2];
      if (map[ca] !== ca && source[ca] !== 'self') continue;
      map[ca] = known;
      source[ca] = 'listings';
    }
  });

  var needCat = [];
  for (var ci = 0; ci < asins.length; ci++) {
    if (map[asins[ci]] === asins[ci]) needCat.push(asins[ci]);
  }
  if (needCat.length && creds && token) {
    uaApplyCatalogParentBatch_(creds, token, needCat, map, source, 'catalog');
  }

  // バッチで取りこぼした子だけ単発再試行
  var needRetry = [];
  for (var rj = 0; rj < asins.length; rj++) {
    if (map[asins[rj]] === asins[rj]) needRetry.push(asins[rj]);
  }
  if (needRetry.length && creds && token) {
    var retryOk = 0;
    for (var ri = 0; ri < needRetry.length; ri++) {
      var one = needRetry[ri];
      var got = uaFetchCatalogParentAsinOne_(creds, token, one);
      if (got && got !== one) {
        map[one] = got;
        source[one] = 'catalogRetry';
        retryOk++;
      }
      Utilities.sleep(250);
    }
    Logger.log('[未広告ASIN] catalogRetry tried=' + needRetry.length + ' ok=' + retryOk);
  }

  // それでも self → 前回一覧の親（子自身は採用しない）
  var cacheOk = 0;
  for (var cj = 0; cj < asins.length; cj++) {
    var ca2 = asins[cj];
    if (map[ca2] !== ca2) continue;
    var prevP = String(prevParentMap[ca2] || '').trim().toUpperCase();
    if (prevP && prevP !== ca2) {
      map[ca2] = prevP;
      source[ca2] = 'prevCache';
      cacheOk++;
    }
  }
  Logger.log('[未広告ASIN] prevCache applied=' + cacheOk);

  var parents = {};
  var srcCount = { listings: 0, catalog: 0, catalogRetry: 0, prevCache: 0, self: 0 };
  Object.keys(map).forEach(function (k) {
    parents[map[k]] = true;
    var s = source[k] || 'self';
    srcCount[s] = (srcCount[s] || 0) + 1;
  });
  Logger.log('[未広告ASIN] parentMap children=' + asins.length +
    ' uniqueParents=' + Object.keys(parents).length +
    ' src=' + JSON.stringify(srcCount));
  return { map: map, source: source };
}

/** 前回一覧シートから 子ASIN→親ASIN */
function uaLoadPrevParentMap_() {
  var out = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(UA_SHEET_LIST_PREV_);
    if (!sh || sh.getLastRow() < 2) return out;
    var data = sh.getDataRange().getValues();
    var col = uaHeaderIndexMap_(data[0]);
    if (col['ASIN'] == null || col['親ASIN'] == null) return out;
    for (var i = 1; i < data.length; i++) {
      var a = String(data[i][col['ASIN']] || '').trim().toUpperCase();
      var p = String(data[i][col['親ASIN']] || '').trim().toUpperCase();
      if (a && p) out[a] = p;
    }
    Logger.log('[未広告ASIN] prevParentMap size=' + Object.keys(out).length);
  } catch (e) {
    Logger.log('[未広告ASIN] prevParentMap FAILED ' + (e && e.message ? e.message : e));
  }
  return out;
}

function uaHeaderIndexMap_(headerRow) {
  var m = {};
  for (var i = 0; i < headerRow.length; i++) {
    var k = String(headerRow[i] || '').trim();
    if (k && m[k] === undefined) m[k] = i;
  }
  return m;
}

function uaApplyCatalogParentBatch_(creds, token, asins, map, source, sourceTag) {
  var batchSize = 20;
  for (var b = 0; b < asins.length; b += batchSize) {
    var chunk = asins.slice(b, b + batchSize);
    try {
      var url = creds.endpoint.replace(/\/$/, '') +
        '/catalog/2022-04-01/items' +
        '?marketplaceIds=' + encodeURIComponent(creds.marketplaceId) +
        '&identifiersType=ASIN' +
        '&identifiers=' + encodeURIComponent(chunk.join(',')) +
        '&includedData=relationships';
      var resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: {
          'x-amz-access-token': token,
          'user-agent': UA_SPAPI_UA_,
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      var text = resp.getContentText() || '';
      if (code !== 200) {
        Logger.log('[未広告ASIN] catalog batch HTTP ' + code + ' body=' + text.substring(0, 160));
        Utilities.sleep(400);
        continue;
      }
      var body = JSON.parse(text);
      var items = body.items || [];
      for (var j = 0; j < items.length; j++) {
        var item = items[j];
        var childAsin = String(item.asin || '').trim().toUpperCase();
        if (!childAsin || map[childAsin] !== childAsin) continue;
        var parentAsin = uaExtractParentAsinFromCatalogItem_(item);
        if (parentAsin) {
          map[childAsin] = parentAsin;
          source[childAsin] = sourceTag || 'catalog';
        }
      }
    } catch (e) {
      Logger.log('[未広告ASIN] catalog batch FAILED ' + (e && e.message ? e.message : e));
    }
    Utilities.sleep(350);
  }
}

function uaFetchCatalogParentAsinOne_(creds, token, childAsin) {
  var a = String(childAsin || '').trim().toUpperCase();
  if (!a) return '';
  try {
    var url = creds.endpoint.replace(/\/$/, '') +
      '/catalog/2022-04-01/items/' + encodeURIComponent(a) +
      '?marketplaceIds=' + encodeURIComponent(creds.marketplaceId) +
      '&includedData=relationships';
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'x-amz-access-token': token,
        'user-agent': UA_SPAPI_UA_,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return '';
    var body = JSON.parse(resp.getContentText() || '{}');
    return uaExtractParentAsinFromCatalogItem_(body) || '';
  } catch (e) {
    return '';
  }
}

/** Listings relationships から親ASIN／親SKUを抜く */
function uaExtractListingParentRef_(relBlocks) {
  var out = { parentAsin: '', parentSku: '' };
  if (!relBlocks) return out;
  var blocks = Array.isArray(relBlocks) ? relBlocks : [relBlocks];
  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i] || {};
    var rels = block.relationships || (Array.isArray(block) ? block : null);
    if (!rels && (block.parentAsins || block.parentSkus || block.parentASINs)) {
      rels = [block];
    }
    if (!rels) continue;
    for (var j = 0; j < rels.length; j++) {
      var rel = rels[j] || {};
      if (!out.parentAsin) {
        var pas = rel.parentAsins || rel.parentASINs || [];
        if (pas && pas.length) out.parentAsin = String(pas[0] || '').trim().toUpperCase();
      }
      if (!out.parentSku) {
        var psk = rel.parentSkus || rel.parentSKUs || [];
        if (psk && psk.length) out.parentSku = String(psk[0] || '').trim();
      }
      if (out.parentAsin && out.parentSku) return out;
    }
  }
  return out;
}

/**
 * セラーSKUの Listings から ASIN を取得（親SKU → 親ASIN用）。キャッシュ付き。
 */
function uaFetchAsinForSellerSku_(creds, token, sellerSku, cache) {
  var sku = String(sellerSku || '').trim();
  if (!sku) return '';
  if (cache && cache[sku] !== undefined) return cache[sku];
  var asin = '';
  try {
    var url = creds.endpoint.replace(/\/$/, '') +
      '/listings/2021-08-01/items/' + encodeURIComponent(creds.sellerId) +
      '/' + encodeURIComponent(sku) +
      '?marketplaceIds=' + encodeURIComponent(creds.marketplaceId) +
      '&includedData=summaries';
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'x-amz-access-token': token,
        'user-agent': UA_SPAPI_UA_,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var text = resp.getContentText() || '';
    if (code === 200) {
      var body = JSON.parse(text);
      var sums = body.summaries || [];
      if (sums.length) asin = String(sums[0].asin || '').trim().toUpperCase();
      if (!asin) asin = String(body.asin || '').trim().toUpperCase();
    } else {
      Logger.log('[未広告ASIN] listingsBySku HTTP ' + code +
        ' sku=' + sku.substring(0, 40) + ' body=' + text.substring(0, 120));
    }
    Utilities.sleep(300);
  } catch (e) {
    Logger.log('[未広告ASIN] listingsBySku FAILED sku=' + sku.substring(0, 40) +
      ' ' + (e && e.message ? e.message : e));
  }
  if (cache) cache[sku] = asin;
  return asin;
}

function uaExtractParentAsinFromCatalogItem_(item) {
  var relBlocks = item.relationships || [];
  for (var i = 0; i < relBlocks.length; i++) {
    var block = relBlocks[i];
    var rels = block.relationships || (Array.isArray(block) ? block : [block]);
    for (var j = 0; j < rels.length; j++) {
      var rel = rels[j] || {};
      var parents = rel.parentAsins || rel.parentASINs || [];
      if (parents && parents.length) {
        return String(parents[0] || '').trim().toUpperCase();
      }
    }
  }
  return '';
}

/**
 * SP-API getItemOffersBatch で新品 TotalOfferCount を取得。
 * 429時は待機リトライ（打ち切りしない）。最大20 ASIN/回。
 * @return {Object} asin -> {ok:boolean, count:number, err?:string}
 */
function uaFetchNewOfferCountsByAsin_(listings) {
  var out = {};
  var enabled = PropertiesService.getScriptProperties().getProperty('UNADVERTISED_ASIN_OFFER_COUNT_ENABLED');
  if (enabled !== null && String(enabled).trim() !== '' &&
      /^(false|0|no|off)$/i.test(String(enabled).trim())) {
    Logger.log('[未広告ASIN] offerCount skipped UNADVERTISED_ASIN_OFFER_COUNT_ENABLED=false');
    return out;
  }
  var asins = [];
  var seen = {};
  for (var i = 0; i < listings.length; i++) {
    var a = String(listings[i].asin || '').trim().toUpperCase();
    if (!a || seen[a]) continue;
    seen[a] = true;
    asins.push(a);
  }
  if (!asins.length) return out;

  var creds;
  var token;
  try {
    creds = uaLoadSpapiCreds_();
    token = uaRequestLwa_(creds.clientId, creds.clientSecret, creds.refreshToken);
  } catch (e) {
    Logger.log('[未広告ASIN] offerCount auth FAIL ' + (e && e.message ? e.message : e));
    return out;
  }

  var batchSize = 20;
  // batch API 目安 0.5回/秒。既定は余裕を見て 3秒
  var sleepMs = Math.max(1000, Math.floor(uaGetNumberProp_('UNADVERTISED_ASIN_OFFER_SLEEP_MS', 3000)));
  var wait429 = Math.max(10000, Math.floor(uaGetNumberProp_('UNADVERTISED_ASIN_OFFER_429_WAIT_MS', 90000)));
  var maxRetry = Math.max(1, Math.floor(uaGetNumberProp_('UNADVERTISED_ASIN_OFFER_429_RETRIES', 5)));
  var batchIndex = 0;

  for (var b = 0; b < asins.length; b += batchSize) {
    batchIndex++;
    var chunk = asins.slice(b, b + batchSize);
    if (batchIndex > 1) Utilities.sleep(sleepMs);

    var pending = chunk.slice();
    var attempt = 0;
    while (pending.length && attempt <= maxRetry) {
      attempt++;
      try {
        var part = uaSpapiItemOffersBatchCounts_(creds, token, pending);
        var got429 = part._stopped === '429';
        Object.keys(part).forEach(function (k) {
          if (k.charAt(0) === '_') return;
          out[k] = part[k];
        });
        // まだ取れていないASINだけ残す
        pending = pending.filter(function (asin) {
          return !(out[asin] && (out[asin].ok || out[asin].err));
        });
        if (!pending.length) break;
        if (got429 || part._http === 429) {
          Logger.log('[未広告ASIN] offerCount 429 batch=' + batchIndex +
            ' attempt=' + attempt + '/' + maxRetry +
            ' remain=' + pending.length + ' waitMs=' + wait429);
          if (attempt > maxRetry) break;
          Utilities.sleep(wait429);
          try {
            token = uaRequestLwa_(creds.clientId, creds.clientSecret, creds.refreshToken);
          } catch (eTok) {
            Logger.log('[未広告ASIN] offerCount re-auth FAIL ' + (eTok && eTok.message ? eTok.message : eTok));
          }
          continue;
        }
        // 429以外で未取得が残る → 短待機して1回だけ再試行、だめなら諦めて次バッチ
        if (pending.length && attempt <= maxRetry) {
          Logger.log('[未広告ASIN] offerCount incomplete batch=' + batchIndex +
            ' remain=' + pending.length + ' retry');
          Utilities.sleep(sleepMs);
          continue;
        }
        break;
      } catch (e2) {
        Logger.log('[未広告ASIN] offerCount batch FAIL batch=' + batchIndex +
          ' ' + (e2 && e2.message ? e2.message : e2));
        if (attempt > maxRetry) break;
        Utilities.sleep(sleepMs);
      }
    }
    if (pending.length) {
      Logger.log('[未広告ASIN] offerCount give up batch=' + batchIndex +
        ' asins=' + pending.join(','));
    }
  }
  var okN = 0;
  Object.keys(out).forEach(function (k) { if (out[k] && out[k].ok) okN++; });
  Logger.log('[未広告ASIN] offerCount done uniqueAsin=' + asins.length + ' ok=' + okN);
  return out;
}

/**
 * @return {Object} asin -> {ok, count, err?} ＋ _stopped / _http
 */
function uaSpapiItemOffersBatchCounts_(creds, token, asins) {
  var mp = creds.marketplaceId;
  var reqs = [];
  for (var i = 0; i < asins.length; i++) {
    reqs.push({
      uri: '/products/pricing/v0/items/' + asins[i] + '/offers',
      method: 'GET',
      MarketplaceId: mp,
      ItemCondition: 'New'
    });
  }
  var url = creds.endpoint.replace(/\/$/, '') + '/batches/products/pricing/v0/itemOffers';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({ requests: reqs }),
    headers: {
      Accept: 'application/json',
      'x-amz-access-token': token,
      'User-Agent': UA_SPAPI_UA_
    }
  });
  var code = resp.getResponseCode();
  var text = resp.getContentText() || '';
  var by = { _http: code };
  if (code === 429) {
    by._stopped = '429';
    return by;
  }
  var data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    Logger.log('[未広告ASIN] offerCount parse FAIL http=' + code + ' body=' + text.substring(0, 160));
    return by;
  }
  if (code !== 200) {
    Logger.log('[未広告ASIN] offerCount HTTP ' + code + ' ' + text.substring(0, 200));
    return by;
  }
  var rows = data.responses || [];
  var any429 = false;
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || {};
    var st = row.status && row.status.statusCode;
    var body = row.body || {};
    var payload = body.payload || body;
    var ident = payload.Identifier || payload.identifier || {};
    var asin = String(ident.ASIN || ident.asin || payload.ASIN || asins[r] || '').trim().toUpperCase();
    if (!asin) continue;
    if (st === 429) {
      any429 = true;
      continue; // 成功分は残し、このASINは未取得のままリトライ
    }
    if (st && Number(st) !== 200) {
      by[asin] = { ok: false, count: 0, err: 'http' + st };
      continue;
    }
    var count = uaExtractNewOfferCount_(payload);
    by[asin] = { ok: true, count: count };
  }
  if (any429) by._stopped = '429';
  return by;
}

/** Summary.TotalOfferCount 優先。無ければ NumberOfOffers の New 合計 */
function uaExtractNewOfferCount_(payload) {
  var summary = payload.Summary || payload.summary || {};
  var total = summary.TotalOfferCount;
  if (total === undefined || total === null) total = summary.totalOfferCount;
  if (total !== undefined && total !== null && isFinite(Number(total))) {
    return Math.max(0, Math.floor(Number(total)));
  }
  var arr = summary.NumberOfOffers || summary.numberOfOffers || [];
  var sum = 0;
  var found = false;
  for (var i = 0; i < arr.length; i++) {
    var o = arr[i] || {};
    var cond = String(o.condition || o.Condition || '').toLowerCase();
    if (cond && cond !== 'new') continue;
    var n = o.OfferCount !== undefined ? o.OfferCount : o.offerCount;
    if (n !== undefined && n !== null && isFinite(Number(n))) {
      sum += Math.floor(Number(n));
      found = true;
    }
  }
  return found ? sum : 0;
}

/**
 * 親ASIN → 全子ASIN（除外KW用）。Catalogの子一覧＋一覧シート上の兄弟をマージ。
 */
function uaFetchChildAsinsForParent_(creds, token, parentAsin) {
  var out = {};
  out[parentAsin] = true;
  try {
    var url = creds.endpoint.replace(/\/$/, '') +
      '/catalog/2022-04-01/items/' + encodeURIComponent(parentAsin) +
      '?marketplaceIds=' + encodeURIComponent(creds.marketplaceId) +
      '&includedData=relationships';
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'x-amz-access-token': token,
        'user-agent': UA_SPAPI_UA_,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 200) {
      var body = JSON.parse(resp.getContentText() || '{}');
      var relBlocks = body.relationships || [];
      for (var i = 0; i < relBlocks.length; i++) {
        var block = relBlocks[i];
        var rels = block.relationships || [];
        for (var j = 0; j < rels.length; j++) {
          var children = rels[j].childAsins || [];
          for (var c = 0; c < children.length; c++) {
            var ca = String(children[c] || '').trim().toUpperCase();
            if (ca) out[ca] = true;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('[未広告ASIN] childAsins parent=' + parentAsin + ' err=' + (e && e.message ? e.message : e));
  }
  return Object.keys(out);
}

function uaLoadSpapiCreds_() {
  var props = PropertiesService.getScriptProperties();
  var clientId = String(props.getProperty(UA_SPAPI_LWA_CLIENT_ID_) || '').trim();
  var clientSecret = String(props.getProperty(UA_SPAPI_LWA_CLIENT_SECRET_) || '').trim();
  var refreshToken = String(props.getProperty(UA_SPAPI_REFRESH_TOKEN_) || '').trim();
  var sellerId = String(props.getProperty(UA_SPAPI_SELLER_ID_) || '').trim();
  var marketplaceId = String(props.getProperty(UA_SPAPI_MARKETPLACE_ID_) || '').trim() || UA_SPAPI_DEFAULT_MARKETPLACE_;
  var endpoint = String(props.getProperty(UA_SPAPI_ENDPOINT_) || '').trim() || UA_SPAPI_DEFAULT_ENDPOINT_;
  if (!clientId || !clientSecret || !refreshToken || !sellerId) {
    throw new Error(
      'SP-APIのScript Propertiesが不足しています。\n' +
      '必要: SPAPI_LWA_CLIENT_ID / SPAPI_LWA_CLIENT_SECRET / SPAPI_REFRESH_TOKEN / SPAPI_SELLER_ID\n' +
      '（出品GASと同じ値を広告運用GASに設定）'
    );
  }
  return {
    clientId: clientId,
    clientSecret: clientSecret,
    refreshToken: refreshToken,
    sellerId: sellerId,
    marketplaceId: marketplaceId,
    endpoint: endpoint
  };
}

function uaRequestLwa_(clientId, clientSecret, refreshToken) {
  var payload = {
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret
  };
  var resp = UrlFetchApp.fetch(UA_SPAPI_LWA_URL_, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded;charset=UTF-8',
    payload: payload,
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var text = resp.getContentText() || '';
  if (code !== 200) throw new Error('LWA token HTTP ' + code + ' ' + text.substring(0, 300));
  var body = JSON.parse(text);
  if (!body.access_token) throw new Error('LWA応答に access_token がありません');
  return String(body.access_token);
}

// --- ID参照／既存名 ---

function uaIdAutoFromBulkEnabled_() {
  var v = PropertiesService.getScriptProperties().getProperty(UA_ID_AUTO_PROP_);
  if (v === null || v === undefined || String(v).trim() === '') return true;
  var s = String(v).trim().toLowerCase();
  return !(s === 'false' || s === '0' || s === 'no' || s === 'off');
}

/**
 * Drive「📥一括キャンペーンDL」の最新xlsx/CSVから ID参照用 を書き換える。
 * ファイルが無い／失敗時は既存シートを残し、summary だけ返す（一覧・バルクは続行可）。
 * @return {{ok:boolean, summary:string, fileName?:string, campaignCount?:number, adGroupCount?:number}}
 */
function uaRefreshIdSheetFromBulkFolder_() {
  if (!uaIdAutoFromBulkEnabled_()) {
    Logger.log('[未広告ASIN] ID参照用自動更新スキップ prop=' + UA_ID_AUTO_PROP_ + '=false');
    return { ok: false, summary: '自動更新OFF（既存シートを使用）' };
  }
  try {
    var latest = uaFindLatestBulkIdFile_();
    if (!latest) {
      Logger.log('[未広告ASIN] ID参照用自動更新: フォルダにxlsx/CSVなし name=' + UA_BULK_ID_FOLDER_NAME_);
      return {
        ok: false,
        summary: '「' + UA_BULK_ID_FOLDER_NAME_ + '」にファイルなし（既存シートを使用）'
      };
    }
    var ageDays = uaGetNumberProp_(UA_CSV_AGE_PROP_, 14);
    var ageMs = Date.now() - latest.getLastUpdated().getTime();
    if (ageMs > ageDays * 24 * 60 * 60 * 1000) {
      Logger.log('[未広告ASIN] ID参照用ソースが古い file=' + latest.getName() +
        ' updated=' + latest.getLastUpdated());
    }
    var table = uaReadBulkIdTableFromFile_(latest);
    var extracted = uaExtractIdMapsFromBulkTable_(table.headers, table.rows);
    if (!extracted.campaignCount) {
      return {
        ok: false,
        summary: latest.getName() + ' に数値キャンペーンIDなし（既存シートを使用）',
        fileName: latest.getName()
      };
    }
    uaWriteIdSheetFromExtracted_(extracted);
    var summary = latest.getName() +
      ' → キャンペーン' + extracted.campaignCount +
      '／AG' + extracted.adGroupCount;
    Logger.log('[未広告ASIN] ID参照用更新OK ' + summary +
      ' sheet=' + (table.sheetName || 'csv'));
    return {
      ok: true,
      summary: summary,
      fileName: latest.getName(),
      campaignCount: extracted.campaignCount,
      adGroupCount: extracted.adGroupCount
    };
  } catch (e) {
    var err = e && e.message ? e.message : String(e);
    Logger.log('[未広告ASIN] ID参照用自動更新FAILED err=' + err);
    return { ok: false, summary: '自動更新失敗: ' + err + '（既存シートを使用）' };
  }
}

/** @return {GoogleAppsScript.Drive.File|null} */
function uaFindLatestBulkIdFile_() {
  if (typeof AMZ_FOLDER_ID === 'undefined' || !AMZ_FOLDER_ID) {
    throw new Error('AMZ_FOLDER_ID が未定義です。');
  }
  var folder = (typeof getOrCreateSubFolder === 'function')
    ? getOrCreateSubFolder(AMZ_FOLDER_ID, UA_BULK_ID_FOLDER_NAME_)
    : uaGetOrCreateSubFolder_(AMZ_FOLDER_ID, UA_BULK_ID_FOLDER_NAME_);
  var files = folder.getFiles();
  var latest = null;
  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName() || '';
    if (name.indexOf('~$') === 0) continue;
    if (!/\.(csv|xlsx|xls)$/i.test(name)) continue;
    if (!latest || f.getLastUpdated().getTime() > latest.getLastUpdated().getTime()) latest = f;
  }
  return latest;
}

/**
 * @param {GoogleAppsScript.Drive.File} file
 * @return {{headers:string[], rows:string[][], sheetName:string}}
 */
function uaReadBulkIdTableFromFile_(file) {
  var name = file.getName() || '';
  var mime = String(file.getMimeType() || '');
  if (/\.csv$/i.test(name) || mime.indexOf('csv') !== -1) {
    var text = file.getBlob().getDataAsString('UTF-8');
    if (text.indexOf('\ufffd') !== -1 || text.indexOf('ã') !== -1) {
      try { text = file.getBlob().getDataAsString('Shift_JIS'); } catch (ignore) {}
    }
    var parsed = uaParseCsv_(text);
    if (!parsed.length) throw new Error('一括DLのCSVが空です: ' + name);
    return {
      headers: parsed[0].map(function (h) { return String(h || '').trim(); }),
      rows: parsed.slice(1),
      sheetName: 'csv'
    };
  }
  if (mime === MimeType.GOOGLE_SHEETS) {
    return uaReadSpBulkSheetValues_(SpreadsheetApp.openById(file.getId()), name);
  }
  var tempId = uaConvertExcelBlobToGoogleSheetId_(file.getBlob(), 'tmp_ua_id_' + Utilities.getUuid());
  try {
    return uaReadSpBulkSheetValues_(SpreadsheetApp.openById(tempId), name);
  } finally {
    try { DriveApp.getFileById(tempId).setTrashed(true); } catch (ignore2) {}
  }
}

/**
 * xlsx/xls を Googleスプレッドシートへ変換（Drive upload API）。
 * @return {string} fileId
 */
function uaConvertExcelBlobToGoogleSheetId_(blob, title) {
  var boundary = 'uaid' + Utilities.getUuid().replace(/-/g, '');
  var contentType = blob.getContentType() ||
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var meta = JSON.stringify({
    name: title,
    mimeType: MimeType.GOOGLE_SHEETS
  });
  var prefix = '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    meta + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: ' + contentType + '\r\n\r\n';
  var suffix = '\r\n--' + boundary + '--';
  var payload = Utilities.newBlob(prefix).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(suffix).getBytes());
  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
    {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'multipart/related; boundary=' + boundary
      },
      payload: payload,
      muteHttpExceptions: true
    }
  );
  var code = res.getResponseCode();
  var bodyText = res.getContentText() || '';
  if (code < 200 || code >= 300) {
    throw new Error('xlsx→Sheets変換失敗 HTTP' + code + ' ' + bodyText.substring(0, 280));
  }
  var body = JSON.parse(bodyText);
  if (!body.id) throw new Error('xlsx→Sheets変換応答にidなし');
  return body.id;
}

/**
 * @return {{headers:string[], rows:string[][], sheetName:string}}
 */
function uaReadSpBulkSheetValues_(ss, sourceName) {
  var sh = uaPickSpBulkSheet_(ss);
  var values = sh.getDataRange().getValues();
  if (!values.length) throw new Error('一括DLシートが空です: ' + sourceName + ' / ' + sh.getName());
  return {
    headers: values[0].map(function (h) { return String(h || '').trim(); }),
    rows: values.slice(1),
    sheetName: sh.getName()
  };
}

function uaPickSpBulkSheet_(ss) {
  var prefer = [
    UA_SP_BULK_SHEET_NAME_,
    'Sponsored Products Campaigns',
    'Sponsored Products',
    'スポンサープロダクト広告'
  ];
  for (var i = 0; i < prefer.length; i++) {
    var sh = ss.getSheetByName(prefer[i]);
    if (sh) return sh;
  }
  var sheets = ss.getSheets();
  for (var j = 0; j < sheets.length; j++) {
    var n = sheets[j].getName() || '';
    if (/スポンサープロダクト|Sponsored\s*Products/i.test(n)) return sheets[j];
  }
  return sheets[0];
}

/**
 * Amazon一括DLでは広告グループ行の「キャンペーン名」が空のことが多い。
 * キャンペーンID＋AG名でも数値AG IDを拾う（空名スキップだと already exists のループになる）。
 * @return {{campaignMap:Object, adGroupMap:Object, campaignCount:number, adGroupCount:number}}
 */
function uaExtractIdMapsFromBulkTable_(headers, rows) {
  var cIdCol = uaFindCol_(headers, [
    'キャンペーンID', 'Campaign ID', 'campaign id', 'CampaignId'
  ]);
  var agIdCol = uaFindCol_(headers, [
    '広告グループID', 'Ad Group ID', 'Ad group ID', 'ad group id', 'AdGroupId'
  ]);
  var cNameCol = uaFindCol_(headers, [
    'キャンペーン名', 'Campaign Name', 'campaign name', '広告キャンペーン名'
  ]);
  var agNameCol = uaFindCol_(headers, [
    '広告グループ名', 'Ad Group Name', 'Ad group Name', 'ad group name', 'AdGroupName'
  ]);
  if (cIdCol < 0 || cNameCol < 0) {
    throw new Error(
      '一括DLにキャンペーンID／キャンペーン名列がありません。ヘッダ=' +
      headers.slice(0, 14).join('|')
    );
  }
  var skuCol = uaFindCol_(headers, ['SKU', 'Sku', 'sku']);
  var kwCol = uaFindCol_(headers, [
    'キーワードテキスト', 'Keyword Text', 'keyword text', 'KeywordText'
  ]);
  var entCol = uaFindCol_(headers, ['エンティティ', 'Entity', 'entity']);
  var campaignMap = {};
  var campIdToName = {};
  var adGroupMap = {};
  var productAds = {};
  var negKeys = {};
  var i;
  var row;
  var cName;
  var cId;
  var normC;
  // 1) キャンペーン名が付いている行から ID↔名
  for (i = 0; i < rows.length; i++) {
    row = rows[i];
    cName = String(row[cNameCol] || '').trim();
    cId = String(row[cIdCol] || '').trim();
    if (!cName || !uaLooksLikeAmzEntityId_(cId)) continue;
    normC = uaNormalizeName_(cName);
    if (!normC) continue;
    campaignMap[normC] = { name: cName, id: cId };
    campIdToName[cId] = cName;
  }
  // 2) AG行（キャンペーン名が空でもキャンペーンIDがあれば取る）
  if (agIdCol >= 0 && agNameCol >= 0) {
    for (i = 0; i < rows.length; i++) {
      row = rows[i];
      cId = String(row[cIdCol] || '').trim();
      cName = String(row[cNameCol] || '').trim() || campIdToName[cId] || '';
      var agName = String(row[agNameCol] || '').trim();
      var agId = String(row[agIdCol] || '').trim();
      if (!uaLooksLikeAmzEntityId_(cId) || !agName || !uaLooksLikeAmzEntityId_(agId)) continue;
      if (!cName) cName = campIdToName[cId] || '';
      if (!cName) {
        // 名がファイル内に一度も出ないときも AG は残す（ID参照はキャンペーンID照合）
        cName = 'Campaign_' + cId;
      }
      if (!campIdToName[cId]) campIdToName[cId] = cName;
      normC = uaNormalizeName_(cName);
      if (normC && !campaignMap[normC]) {
        campaignMap[normC] = { name: cName, id: cId };
      }
      var agKey = (normC || cId) + '|' + uaNormalizeName_(agName);
      adGroupMap[agKey] = {
        campaignName: cName,
        campaignId: cId,
        agName: agName,
        agId: agId
      };
    }
  }
  // 3) 既存の商品広告SKU・除外KW（再ULの already exists を避ける）
  for (i = 0; i < rows.length; i++) {
    row = rows[i];
    cId = String(row[cIdCol] || '').trim();
    var agId2 = agIdCol >= 0 ? String(row[agIdCol] || '').trim() : '';
    if (!uaLooksLikeAmzEntityId_(cId)) continue;
    var ent = entCol >= 0 ? String(row[entCol] || '').trim() : '';
    if (skuCol >= 0) {
      var sku = String(row[skuCol] || '').trim();
      if (sku && uaLooksLikeAmzEntityId_(agId2)) {
        productAds[cId + '|' + agId2 + '|' + sku.toUpperCase()] = {
          campaignId: cId,
          agId: agId2,
          sku: sku,
          campaignName: campIdToName[cId] || String(row[cNameCol] || '').trim(),
          agName: agNameCol >= 0 ? String(row[agNameCol] || '').trim() : ''
        };
      }
    }
    if (kwCol >= 0 && (/除外/.test(ent) || /negative/i.test(ent))) {
      var kw = String(row[kwCol] || '').trim().toUpperCase();
      if (kw && /^B0[A-Z0-9]{8}$/i.test(kw)) {
        if (uaLooksLikeAmzEntityId_(agId2)) negKeys[cId + '|' + agId2 + '|' + kw] = true;
        else negKeys[cId + '||' + kw] = true;
      }
    }
  }
  Logger.log('[未広告ASIN] ID抽出 campaigns=' + Object.keys(campaignMap).length +
    ' adGroups=' + Object.keys(adGroupMap).length +
    ' productAds=' + Object.keys(productAds).length +
    ' negKeys=' + Object.keys(negKeys).length +
    ' campIdNames=' + Object.keys(campIdToName).length);
  return {
    campaignMap: campaignMap,
    adGroupMap: adGroupMap,
    productAds: productAds,
    negKeys: negKeys,
    campaignCount: Object.keys(campaignMap).length,
    adGroupCount: Object.keys(adGroupMap).length
  };
}

/** 抽出結果を公式31列レイアウトで ID参照用 に書く（KWレーン互換） */
function uaWriteIdSheetFromExtracted_(extracted) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UA_SHEET_ID_);
  if (!sh) sh = ss.insertSheet(UA_SHEET_ID_);
  var out = [UA_SP_BULK_HEADERS_.slice()];
  var campKeys = Object.keys(extracted.campaignMap);
  campKeys.sort();
  for (var i = 0; i < campKeys.length; i++) {
    var c = extracted.campaignMap[campKeys[i]];
    out.push(uaSpBulkRow_({
      'プロダクト': 'スポンサープロダクト広告',
      'エンティティ': 'キャンペーン',
      'キャンペーンID': c.id,
      'キャンペーン名': c.name
    }));
  }
  var agKeys = Object.keys(extracted.adGroupMap);
  agKeys.sort();
  for (var j = 0; j < agKeys.length; j++) {
    var a = extracted.adGroupMap[agKeys[j]];
    out.push(uaSpBulkRow_({
      'プロダクト': 'スポンサープロダクト広告',
      'エンティティ': '広告グループ',
      'キャンペーンID': a.campaignId,
      '広告グループID': a.agId,
      'キャンペーン名': a.campaignName,
      '広告グループ名': a.agName
    }));
  }
  var paKeys = Object.keys(extracted.productAds || {});
  paKeys.sort();
  for (var p = 0; p < paKeys.length; p++) {
    var pa = extracted.productAds[paKeys[p]];
    out.push(uaSpBulkRow_({
      'プロダクト': 'スポンサープロダクト広告',
      'エンティティ': UA_ENT_PRODUCT_AD_,
      'キャンペーンID': pa.campaignId,
      '広告グループID': pa.agId,
      'キャンペーン名': pa.campaignName || '',
      '広告グループ名': pa.agName || '',
      'SKU': pa.sku
    }));
  }
  sh.clear();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  // 除外KWは行数爆発しやすいので Script Properties に要約キーだけ残す
  try {
    PropertiesService.getScriptProperties().setProperty(
      'UNADVERTISED_ASIN_EXISTING_NEG_KEYS',
      JSON.stringify(Object.keys(extracted.negKeys || {}).slice(0, 20000))
    );
  } catch (negPropErr) {
    Logger.log('[未広告ASIN] negKeys prop skip ' + (negPropErr && negPropErr.message ? negPropErr.message : negPropErr));
  }
  Logger.log('[未広告ASIN] ID参照用書込 rows=' + (out.length - 1) +
    ' campaigns=' + extracted.campaignCount +
    ' adGroups=' + extracted.adGroupCount +
    ' productAds=' + paKeys.length);
}

function uaLoadExistingCampaignNamesFromIdSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UA_SHEET_ID_);
  var names = [];
  if (!sh) return names;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var cName = data[i][9];
    if (cName) names.push(String(cName));
  }
  return names;
}

function uaLoadIdLookupMaps_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UA_SHEET_ID_);
  var campaignMap = {};
  var campIdToName = {};
  var adGroupMap = {};
  var adGroupByCampId = {};
  var existingPa = {};
  var existingNeg = {};
  if (!sh) {
    return {
      campaignMap: campaignMap,
      campIdToName: campIdToName,
      adGroupMap: adGroupMap,
      adGroupByCampId: adGroupByCampId,
      existingPa: existingPa,
      existingNeg: existingNeg
    };
  }
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var ent = String(data[i][1] || '').trim();
    var cName = data[i][9];
    var cId = String(data[i][3] || '').trim();
    var agId = String(data[i][4] || '').trim();
    var agName = data[i][10];
    var sku = String(data[i][16] || '').trim();
    var norm = uaNormalizeName_(cName);
    if (norm && cId) {
      if (uaLooksLikeAmzEntityId_(cId)) {
        campaignMap[norm] = cId;
        campIdToName[cId] = String(cName || campIdToName[cId] || '');
      } else if (!campaignMap[norm]) {
        campaignMap[norm] = cId;
      }
    }
    if (agName && uaLooksLikeAmzEntityId_(agId)) {
      var nAg = uaNormalizeName_(agName);
      if (norm) adGroupMap[norm + '|' + nAg] = agId;
      if (uaLooksLikeAmzEntityId_(cId)) {
        adGroupByCampId[cId + '|' + nAg] = agId;
      }
    }
    if (sku && uaLooksLikeAmzEntityId_(cId) && uaLooksLikeAmzEntityId_(agId) &&
        (ent === UA_ENT_PRODUCT_AD_ || ent.indexOf('プロダクト') !== -1 || /Product\s*Ad/i.test(ent))) {
      existingPa[cId + '|' + agId + '|' + sku.toUpperCase()] = true;
    }
  }
  try {
    var negJson = PropertiesService.getScriptProperties().getProperty('UNADVERTISED_ASIN_EXISTING_NEG_KEYS');
    if (negJson) {
      var negArr = JSON.parse(negJson);
      for (var n = 0; n < negArr.length; n++) existingNeg[negArr[n]] = true;
    }
  } catch (negLoadErr) {
    Logger.log('[未広告ASIN] negKeys load skip');
  }
  return {
    campaignMap: campaignMap,
    campIdToName: campIdToName,
    adGroupMap: adGroupMap,
    adGroupByCampId: adGroupByCampId,
    existingPa: existingPa,
    existingNeg: existingNeg
  };
}

/** キャンペーン名または数値キャンペーンID＋AG名で数値AG IDを探す */
function uaResolveAgId_(idLookup, campName, campId, agName) {
  var nAg = uaNormalizeName_(agName);
  if (!nAg) return '';
  var nCamp = uaNormalizeName_(campName);
  var id = '';
  if (nCamp) id = (idLookup.adGroupMap || {})[nCamp + '|' + nAg] || '';
  if (uaLooksLikeAmzEntityId_(id)) return String(id);
  var cId = String(campId || '').trim();
  if (uaLooksLikeAmzEntityId_(cId)) {
    id = (idLookup.adGroupByCampId || {})[cId + '|' + nAg] || '';
    if (uaLooksLikeAmzEntityId_(id)) return String(id);
  }
  return '';
}

function uaCollectAllChildrenByParentFromList_(data, col) {
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var p = String(data[i][col['親ASIN']] || '').trim().toUpperCase();
    var a = String(data[i][col['ASIN']] || '').trim().toUpperCase();
    if (!p || !a) continue;
    if (!map[p]) map[p] = {};
    map[p][a] = true;
  }
  var out = {};
  Object.keys(map).forEach(function (p) {
    out[p] = Object.keys(map[p]);
  });
  return out;
}

// --- バルク（公式一括テンプレ SP シート準拠） ---

/** ヘッダー名→値のマップを公式31列の1行にする */
function uaSpBulkRow_(fieldMap) {
  var row = [];
  for (var i = 0; i < UA_SP_BULK_HEADERS_.length; i++) {
    var key = UA_SP_BULK_HEADERS_[i];
    var v = fieldMap[key];
    row.push(v === undefined || v === null ? '' : v);
  }
  return row;
}

/** AmazonのエンティティID（数値）か */
function uaLooksLikeAmzEntityId_(id) {
  return /^\d{6,}$/.test(String(id || '').trim());
}

/**
 * ファイル内で一意な広告グループ仮ID（表示名は別列）。
 * オートとManual_で同じバリ名でも衝突しない。
 */
function uaBulkDraftAgId_(campaignKey, agName) {
  var a = String(campaignKey || '').trim();
  var b = String(agName || '').trim();
  var id = a + ' | ' + b;
  if (id.length <= 128) return id;
  var keep = 128 - 3 - Math.min(b.length, 60);
  if (keep < 8) keep = 8;
  return a.substring(0, keep) + ' | ' + b.substring(0, 60);
}

/**
 * @return {{rows: string[][], skipped: string[]}}
 */
function uaBuildBulkRowsParentChild_(targets, allChildrenByParent, idLookup, budget, bid, startDate) {
  var rows = [UA_SP_BULK_HEADERS_.slice()];
  var skipped = [];
  var byParent = {};
  for (var i = 0; i < targets.length; i++) {
    var t = targets[i];
    if (!byParent[t.parentAsin]) byParent[t.parentAsin] = [];
    byParent[t.parentAsin].push(t);
  }

  var creds = null;
  var token = '';
  try {
    creds = uaLoadSpapiCreds_();
    token = uaRequestLwa_(creds.clientId, creds.clientSecret, creds.refreshToken);
  } catch (e) {
    Logger.log('[未広告ASIN] bulk catalog skip ' + (e && e.message ? e.message : e));
  }

  var createdAgKeys = {};
  var seenCampNeg = {};
  var seenAgNeg = {};
  var seenSeedKw = {};

  var parentKeys = Object.keys(byParent);
  for (var p = 0; p < parentKeys.length; p++) {
    var parentAsin = parentKeys[p];
    var group = byParent[parentAsin];
    var sample = group[0];
    var policy0 = String(sample.policy || '');
    var isExisting = policy0.indexOf('既存') !== -1;
    var autoName = isExisting && sample.existAuto ? sample.existAuto : sample.autoName;
    var manName = isExisting && sample.existMan ? sample.existMan : sample.manName;
    if (!/^Manual_/i.test(manName)) manName = 'Manual_' + manName.replace(/^Manual_/i, '');

    // ID参照用に数値IDがある名前は「新規」でも既存扱いにする（already exists 防止）
    var autoDet = uaCampIdLookupDetail_(idLookup.campaignMap, autoName);
    var manDet = uaCampIdLookupDetail_(idLookup.campaignMap, manName);
    var autoIdMapped = autoDet.id;
    var manIdMapped = manDet.id;
    var resolvePath = isExisting ? 'policy_existing' : 'policy_new';
    if (!isExisting &&
        uaLooksLikeAmzEntityId_(autoIdMapped) &&
        uaLooksLikeAmzEntityId_(manIdMapped)) {
      isExisting = true;
      resolvePath = 'promote_before_sanitize';
      // Amazon上の正式名があればそれに合わせる
      if (idLookup.campIdToName && idLookup.campIdToName[autoIdMapped]) {
        autoName = idLookup.campIdToName[autoIdMapped];
      }
      if (idLookup.campIdToName && idLookup.campIdToName[manIdMapped]) {
        manName = idLookup.campIdToName[manIdMapped];
      }
      Logger.log('[未広告ASIN] bulk promote to existing by ID sheet parent=' + parentAsin +
        ' auto=' + autoName + ' man=' + manName);
    }

    var autoId;
    var manId;
    if (isExisting) {
      autoId = autoIdMapped || uaLookupCampId_(idLookup.campaignMap, autoName);
      manId = manIdMapped || uaLookupCampId_(idLookup.campaignMap, manName);
      if (!uaLooksLikeAmzEntityId_(autoId) || !uaLooksLikeAmzEntityId_(manId)) {
        var skipLabel = autoName + ' / ' + manName +
          '（既存だがID参照用に数値キャンペーンIDなし。「' +
          UA_BULK_ID_FOLDER_NAME_ + '」を取り直して再実行）';
        skipped.push(skipLabel);
        uaLogBulkParentResolve_({
          parentAsin: parentAsin,
          policy: policy0,
          path: 'skip_existing_no_id',
          autoName: autoName,
          manName: manName,
          autoId: autoId,
          manId: manId,
          autoHit: autoDet.hitHow,
          manHit: manDet.hitHow,
          autoNear: autoDet.near,
          manNear: manDet.near,
          skuCount: group.length
        });
        continue;
      }
      if (idLookup.campIdToName && idLookup.campIdToName[autoId]) {
        autoName = idLookup.campIdToName[autoId];
      }
      if (idLookup.campIdToName && idLookup.campIdToName[manId]) {
        manName = idLookup.campIdToName[manId];
      }
      if (resolvePath === 'policy_existing') resolvePath = 'existing_with_id';
    } else {
      // 新規名から Amazon 拒否文字（× 等）を除去（括弧の全角半角は変えない）
      autoName = uaSanitizeAmzBulkName_(autoName);
      manName = uaSanitizeAmzBulkName_(manName);
      if (!/^Manual_/i.test(manName)) manName = 'Manual_' + manName.replace(/^Manual_/i, '');
      // sanitize 後にもう一度既存照合
      autoDet = uaCampIdLookupDetail_(idLookup.campaignMap, autoName);
      manDet = uaCampIdLookupDetail_(idLookup.campaignMap, manName);
      autoIdMapped = autoDet.id;
      manIdMapped = manDet.id;
      if (uaLooksLikeAmzEntityId_(autoIdMapped) && uaLooksLikeAmzEntityId_(manIdMapped)) {
        isExisting = true;
        autoId = autoIdMapped;
        manId = manIdMapped;
        resolvePath = 'promote_after_sanitize';
        if (idLookup.campIdToName && idLookup.campIdToName[autoId]) {
          autoName = idLookup.campIdToName[autoId];
        }
        if (idLookup.campIdToName && idLookup.campIdToName[manId]) {
          manName = idLookup.campIdToName[manId];
        }
      } else {
        autoId = autoName;
        manId = manName;
        resolvePath = 'create_by_name';
      }
    }
    uaLogBulkParentResolve_({
      parentAsin: parentAsin,
      policy: policy0,
      path: resolvePath,
      autoName: autoName,
      manName: manName,
      autoId: uaLooksLikeAmzEntityId_(autoId) ? autoId : '(name)',
      manId: uaLooksLikeAmzEntityId_(manId) ? manId : '(name)',
      autoHit: autoDet.hitHow,
      manHit: manDet.hitHow,
      autoNear: autoDet.near,
      manNear: manDet.near,
      skuCount: group.length
    });

    var negAsins = (allChildrenByParent[parentAsin] || []).slice();
    if (creds && token) {
      var fromCat = uaFetchChildAsinsForParent_(creds, token, parentAsin);
      var merged = {};
      negAsins.concat(fromCat).forEach(function (x) { if (x) merged[x] = true; });
      group.forEach(function (g) { merged[g.asin] = true; });
      negAsins = Object.keys(merged);
      Utilities.sleep(300);
    } else {
      group.forEach(function (g) {
        if (negAsins.indexOf(g.asin) === -1) negAsins.push(g.asin);
      });
    }

    if (!isExisting) {
      rows.push(uaSpBulkRow_({
        'プロダクト': 'スポンサープロダクト広告',
        'エンティティ': 'キャンペーン',
        '操作': '作成',
        'キャンペーンID': autoId,
        'キャンペーン名': autoName,
        '開始日': startDate,
        'ターゲティングのタイプ': 'オート',
        'ステータス': '有効',
        '1日の予算': budget,
        '入札戦略': UA_BID_STRATEGY_
      }));
      rows.push(uaSpBulkRow_({
        'プロダクト': 'スポンサープロダクト広告',
        'エンティティ': 'キャンペーン',
        '操作': '作成',
        'キャンペーンID': manId,
        'キャンペーン名': manName,
        '開始日': startDate,
        'ターゲティングのタイプ': 'マニュアル',
        'ステータス': '有効',
        '1日の予算': budget,
        '入札戦略': UA_BID_STRATEGY_
      }));
      for (var pl = 0; pl < UA_PLACEMENTS_.length; pl++) {
        var place = UA_PLACEMENTS_[pl];
        rows.push(uaSpBulkRow_({
          'プロダクト': 'スポンサープロダクト広告',
          'エンティティ': UA_ENT_BID_ADJ_,
          '操作': '作成',
          'キャンペーンID': autoId,
          'キャンペーン名': autoName,
          '入札戦略': UA_BID_STRATEGY_,
          '掲載枠': place,
          '割合': 0
        }));
        rows.push(uaSpBulkRow_({
          'プロダクト': 'スポンサープロダクト広告',
          'エンティティ': UA_ENT_BID_ADJ_,
          '操作': '作成',
          'キャンペーンID': manId,
          'キャンペーン名': manName,
          '入札戦略': UA_BID_STRATEGY_,
          '掲載枠': place,
          '割合': 0
        }));
      }
    }

    // キャンペーン単位の除外（親×除外テキストで1回）
    var existingNeg = idLookup.existingNeg || {};
    for (var n0 = 0; n0 < negAsins.length; n0++) {
      var negCamp = String(negAsins[n0] || '').trim();
      if (!negCamp) continue;
      var negCampU = negCamp.toUpperCase();
      var ckAuto = String(autoId) + '\t' + negCamp;
      var ckMan = String(manId) + '\t' + negCamp;
      if (!seenCampNeg[ckAuto] && !existingNeg[String(autoId) + '||' + negCampU]) {
        seenCampNeg[ckAuto] = true;
        rows.push(uaSpBulkRow_({
          'プロダクト': 'スポンサープロダクト広告',
          'エンティティ': UA_ENT_CAMP_NEG_KW_,
          '操作': '作成',
          'キャンペーンID': autoId,
          'キャンペーン名': autoName,
          'ステータス': '有効',
          'キーワードテキスト': negCamp,
          'マッチタイプ': UA_NEG_MATCH_
        }));
      }
      if (!seenCampNeg[ckMan] && !existingNeg[String(manId) + '||' + negCampU]) {
        seenCampNeg[ckMan] = true;
        rows.push(uaSpBulkRow_({
          'プロダクト': 'スポンサープロダクト広告',
          'エンティティ': UA_ENT_CAMP_NEG_KW_,
          '操作': '作成',
          'キャンペーンID': manId,
          'キャンペーン名': manName,
          'ステータス': '有効',
          'キーワードテキスト': negCamp,
          'マッチタイプ': UA_NEG_MATCH_
        }));
      }
    }

    for (var g = 0; g < group.length; g++) {
      var child = group[g];
      var agName = child.agName;
      if (!isExisting) agName = uaSanitizeAmzBulkName_(agName) || agName;
      var agIdAutoLookup = uaResolveAgId_(idLookup, autoName, autoId, agName);
      var agIdManLookup = uaResolveAgId_(idLookup, manName, manId, agName);
      // 数値AG IDがある → 既存AGへ紐づけ（作成しない＝already exists 防止）
      // 無い → 仮IDで広告グループ行を作成（名前だけIDは禁止）
      var agIdAuto;
      var agIdMan;
      var agNewAuto;
      var agNewMan;
      if (uaLooksLikeAmzEntityId_(agIdAutoLookup)) {
        agIdAuto = agIdAutoLookup;
        agNewAuto = false;
      } else {
        agIdAuto = uaBulkDraftAgId_(autoId, agName);
        agNewAuto = true;
      }
      if (uaLooksLikeAmzEntityId_(agIdManLookup)) {
        agIdMan = agIdManLookup;
        agNewMan = false;
      } else {
        agIdMan = uaBulkDraftAgId_(manId, agName);
        agNewMan = true;
      }

      var agCreateKeyAuto = String(autoId) + '\t' + String(agIdAuto);
      var agCreateKeyMan = String(manId) + '\t' + String(agIdMan);

      if (agNewAuto && !createdAgKeys[agCreateKeyAuto]) {
        createdAgKeys[agCreateKeyAuto] = true;
        rows.push(uaSpBulkRow_({
          'プロダクト': 'スポンサープロダクト広告',
          'エンティティ': '広告グループ',
          '操作': '作成',
          'キャンペーンID': autoId,
          'キャンペーン名': autoName,
          '広告グループID': agIdAuto,
          '広告グループ名': agName,
          'ステータス': '有効',
          '広告グループの入札額の初期値': bid
        }));
      }
      var existingPa = idLookup.existingPa || {};
      var skuU = String(child.sku || '').trim().toUpperCase();
      var paKeyAuto = String(autoId) + '|' + String(agIdAuto) + '|' + skuU;
      var paKeyMan = String(manId) + '|' + String(agIdMan) + '|' + skuU;
      if (skuU && !existingPa[paKeyAuto]) {
        rows.push(uaSpBulkRow_({
          'プロダクト': 'スポンサープロダクト広告',
          'エンティティ': UA_ENT_PRODUCT_AD_,
          '操作': '作成',
          'キャンペーンID': autoId,
          'キャンペーン名': autoName,
          '広告グループID': agIdAuto,
          '広告グループ名': agName,
          'ステータス': '有効',
          'SKU': child.sku
        }));
      }
      if (agNewMan && !createdAgKeys[agCreateKeyMan]) {
        createdAgKeys[agCreateKeyMan] = true;
        rows.push(uaSpBulkRow_({
          'プロダクト': 'スポンサープロダクト広告',
          'エンティティ': '広告グループ',
          '操作': '作成',
          'キャンペーンID': manId,
          'キャンペーン名': manName,
          '広告グループID': agIdMan,
          '広告グループ名': agName,
          'ステータス': '有効',
          '広告グループの入札額の初期値': bid
        }));
      }
      if (skuU && !existingPa[paKeyMan]) {
        rows.push(uaSpBulkRow_({
          'プロダクト': 'スポンサープロダクト広告',
          'エンティティ': UA_ENT_PRODUCT_AD_,
          '操作': '作成',
          'キャンペーンID': manId,
          'キャンペーン名': manName,
          '広告グループID': agIdMan,
          '広告グループ名': agName,
          'ステータス': '有効',
          'SKU': child.sku
        }));
      }

      var seedKw = uaSanitizeSeedKeyword_(child.nameBase || autoName);
      if (seedKw) {
        var sk = String(manId) + '\t' + String(agIdMan) + '\t' + seedKw;
        if (!seenSeedKw[sk]) {
          seenSeedKw[sk] = true;
          rows.push(uaSpBulkRow_({
            'プロダクト': 'スポンサープロダクト広告',
            'エンティティ': 'キーワード',
            '操作': '作成',
            'キャンペーンID': manId,
            'キャンペーン名': manName,
            '広告グループID': agIdMan,
            '広告グループ名': agName,
            'ステータス': '有効',
            '入札額': bid,
            'キーワードテキスト': seedKw,
            'マッチタイプ': UA_POS_MATCH_EXACT_
          }));
        }
      }

      for (var n = 0; n < negAsins.length; n++) {
        var neg = String(negAsins[n] || '').trim();
        if (!neg) continue;
        var negU = neg.toUpperCase();
        var nkAuto = String(autoId) + '\t' + String(agIdAuto) + '\t' + neg;
        var nkMan = String(manId) + '\t' + String(agIdMan) + '\t' + neg;
        var exNegAuto = existingNeg[String(autoId) + '|' + String(agIdAuto) + '|' + negU];
        var exNegMan = existingNeg[String(manId) + '|' + String(agIdMan) + '|' + negU];
        if (!seenAgNeg[nkAuto] && !exNegAuto) {
          seenAgNeg[nkAuto] = true;
          rows.push(uaSpBulkRow_({
            'プロダクト': 'スポンサープロダクト広告',
            'エンティティ': UA_ENT_AG_NEG_KW_,
            '操作': '作成',
            'キャンペーンID': autoId,
            'キャンペーン名': autoName,
            '広告グループID': agIdAuto,
            '広告グループ名': agName,
            'ステータス': '有効',
            'キーワードテキスト': neg,
            'マッチタイプ': UA_NEG_MATCH_
          }));
        }
        if (!seenAgNeg[nkMan] && !exNegMan) {
          seenAgNeg[nkMan] = true;
          rows.push(uaSpBulkRow_({
            'プロダクト': 'スポンサープロダクト広告',
            'エンティティ': UA_ENT_AG_NEG_KW_,
            '操作': '作成',
            'キャンペーンID': manId,
            'キャンペーン名': manName,
            '広告グループID': agIdMan,
            '広告グループ名': agName,
            'ステータス': '有効',
            'キーワードテキスト': neg,
            'マッチタイプ': UA_NEG_MATCH_
          }));
        }
      }
    }
  }
  Logger.log('[未広告ASIN] bulk rows=' + (rows.length - 1) +
    ' skippedParents=' + skipped.length +
    ' headerCols=' + UA_SP_BULK_HEADERS_.length +
    ' sheet=' + UA_SP_BULK_SHEET_NAME_);
  return { rows: rows, skipped: skipped };
}

function uaSaveBulkXlsx_(bulkRows, runId) {
  var folderId = (typeof AMZ_FOLDER_ID !== 'undefined') ? AMZ_FOLDER_ID : '';
  if (!folderId) throw new Error('AMZ_FOLDER_ID がありません');
  var outName = (typeof OUTPUT_FOLDER_NAME !== 'undefined') ? OUTPUT_FOLDER_NAME : '📦生成されたバルクファイル';
  var sub = (typeof getOrCreateSubFolder === 'function')
    ? getOrCreateSubFolder(folderId, outName)
    : uaGetOrCreateSubFolder_(folderId, outName);
  var tempSS = SpreadsheetApp.create('temp_ua_' + runId);
  var tempSheet = tempSS.getSheets()[0];
  tempSheet.setName(UA_SP_BULK_SHEET_NAME_);
  tempSheet.getRange(1, 1, bulkRows.length, bulkRows[0].length).setValues(bulkRows);
  SpreadsheetApp.flush();
  var stamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');
  var fileName = 'UnadvertisedAsinBulk_' + stamp + '.xlsx';
  var blob = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
  ).getBlob().setName(fileName);
  sub.createFile(blob);
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  Logger.log('[未広告ASIN] bulk saved name=' + fileName + ' sheet=' + UA_SP_BULK_SHEET_NAME_);
  return fileName;
}

function uaGetOrCreateSubFolder_(parentId, name) {
  var p = DriveApp.getFolderById(parentId);
  var fs = p.getFoldersByName(name);
  return fs.hasNext() ? fs.next() : p.createFolder(name);
}

// --- props ---

function uaEnabled_() {
  var v = PropertiesService.getScriptProperties().getProperty(UA_ENABLED_PROP_);
  if (v === null || v === undefined || String(v).trim() === '') return true;
  var s = String(v).trim().toLowerCase();
  return !(s === 'false' || s === '0' || s === 'no' || s === 'off');
}

function uaGetNumberProp_(key, defVal) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  if (v === null || v === undefined || String(v).trim() === '') return defVal;
  var n = Number(v);
  return isFinite(n) ? n : defVal;
}
