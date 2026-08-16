/**
 * SP-API v1.4 — GAS から Listings Items 直呼び
 *   第1段: 子SKUレ点（21-⑩ dry_run／21-⑪ prod）
 *   第2段: 承認①済 Amazon（21-⑫ dry_run／21-⑬ prod）※21-⑨と同抽出
 *
 * - dry_run: VALIDATION_PREVIEW（永続化しない）
 * - prod: 実 PUT（ALLOW_PROD 必須）
 * 承認: docs/org/LV4_SPAPI_GAS_PUT_APPROVAL.md
 *       docs/org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md
 * 手順: docs/org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md
 *
 * Script Properties（第1段・第2段で共用）— 本番常時ONセット（2026-08-10）:
 *   APPROVAL_AMAZON_SPAPI_PUT_ENABLED … 未設定時 **true**（明示 false で緊急停止）
 *   APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD … 未設定時 **true**（開始前確認は残す）
 *   APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY … 既定 **false**（承認②・毎回ONにしない）
 *   APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS … 既定 10（1〜50。未設定時はこの既定）
 *   APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_0 … 既定 true（マスタqty経路ON時のみ緩和）
 * デュアル Phase1: 自己発=Amazon相乗りSKU／FBA=Amazon相乗りSKU_FBA（列分離・1実行1系統）
 *   APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_ATTRS … 既定 true（FBA時に電池・危険物属性）
 *   SPAPI_LWA_CLIENT_ID / SPAPI_LWA_CLIENT_SECRET / SPAPI_REFRESH_TOKEN
 *   SPAPI_SELLER_ID / SPAPI_MARKETPLACE_ID（既定 JP: A1VC38T7YXB528）
 *   SPAPI_ENDPOINT（既定 https://sellingpartnerapi-fe.amazon.com）
 *
 * ※ Script Properties に既に false が入っている場合は、キー削除（未設定＝ON）か true に1回書き換え。
 */

var APPROVAL_AMAZON_SPAPI_PUT_PROP = 'APPROVAL_AMAZON_SPAPI_PUT_ENABLED';
var APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD_PROP = 'APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD';
var APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY_PROP = 'APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY';
var APPROVAL_AMAZON_SPAPI_PUT_MAX_PROP = 'APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS';
var APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_PROP = 'APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_0';
/** 本番常時ONセット: 未設定時 true（明示 false でOFF） */
var AMAZON_PROD_DEFAULT_PUT_ENABLED_ = true;
var AMAZON_PROD_DEFAULT_ALLOW_PROD_ = true;
/** FBA時の電池・危険物属性付与（既定 true。false で旧body） */
var APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_PROP =
  'APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_ATTRS';
/** 自己発（MFN）専用。FBA実行では読まない・書かない（デュアル Phase1） */
var AMAZON_OFFER_SELLER_SKU_HEADER_ = 'Amazon相乗りSKU';
/** FBA専用。自己発実行では読まない・書かない */
var AMAZON_OFFER_SELLER_SKU_FBA_HEADER_ = 'Amazon相乗りSKU_FBA';
var AMAZON_OFFER_ASIN_HEADER_ = 'ASINコード';

/**
 * 系統別の相乗りSKU保存列名（デュアル Phase1）。
 * @param {string=} fulfillment mfn | fba
 * @return {string} 'Amazon相乗りSKU' | 'Amazon相乗りSKU_FBA'
 */
function amazonSpapiPutOfferSellerSkuHeader_(fulfillment) {
  return String(fulfillment || 'mfn').toLowerCase() === 'fba'
    ? AMAZON_OFFER_SELLER_SKU_FBA_HEADER_
    : AMAZON_OFFER_SELLER_SKU_HEADER_;
}

/**
 * D相乗りの系統リストを正規化。
 * @param {string=} raw 'mfn' | 'fba' | 'mfn,fba' | 'both'
 * @return {string[]} 順序は常に自己発→FBA（選んだものだけ）
 */
function amazonSpapiPutParseOfferFulfillments_(raw) {
  var s = String(raw || 'mfn').toLowerCase().replace(/\s+/g, '');
  var wantMfn = false;
  var wantFba = false;
  if (s === 'both') {
    wantMfn = true;
    wantFba = true;
  } else {
    var parts = s.split(',');
    for (var i = 0; i < parts.length; i++) {
      if (parts[i] === 'mfn' || parts[i] === 'self') wantMfn = true;
      if (parts[i] === 'fba') wantFba = true;
    }
  }
  if (!wantMfn && !wantFba) wantMfn = true;
  var out = [];
  if (wantMfn) out.push('mfn');
  if (wantFba) out.push('fba');
  return out;
}

var SPAPI_LWA_CLIENT_ID_PROP = 'SPAPI_LWA_CLIENT_ID';
var SPAPI_LWA_CLIENT_SECRET_PROP = 'SPAPI_LWA_CLIENT_SECRET';
var SPAPI_REFRESH_TOKEN_PROP = 'SPAPI_REFRESH_TOKEN';
var SPAPI_SELLER_ID_PROP = 'SPAPI_SELLER_ID';
var SPAPI_MARKETPLACE_ID_PROP = 'SPAPI_MARKETPLACE_ID';
var SPAPI_ENDPOINT_PROP = 'SPAPI_ENDPOINT';

var SPAPI_LWA_TOKEN_URL_ = 'https://api.amazon.com/auth/o2/token';
var SPAPI_DEFAULT_ENDPOINT_ = 'https://sellingpartnerapi-fe.amazon.com';
var SPAPI_DEFAULT_MARKETPLACE_ = 'A1VC38T7YXB528';
var SPAPI_USER_AGENT_ = 'OctasSpapiGasPut/1.4 (Language=GoogleAppsScript)';

/**
 * メニュー 21-⑩: 子SKUレ点 → Listings PUT dry_run（VALIDATION_PREVIEW）
 */
function menuAmazonSpapiPutDryRun() {
  return menuAmazonSpapiPutListings_({ mode: 'dry_run', source: 'child_ck' });
}

/**
 * メニュー 21-⑪: 子SKUレ点 → Listings PUT prod（ALLOW_PROD 必須）
 */
function menuAmazonSpapiPutProd() {
  return menuAmazonSpapiPutListings_({ mode: 'prod', source: 'child_ck' });
}

/**
 * メニュー 21-⑫: 承認①済 Amazon → Listings PUT dry_run（第2段）
 */
function menuAmazonSpapiPutApprovedDryRun() {
  return menuAmazonSpapiPutListings_({ mode: 'dry_run', source: 'approved' });
}

/**
 * メニュー 21-⑬: 承認①済 Amazon → Listings PUT prod（第2段・ALLOW_PROD 必須）
 */
function menuAmazonSpapiPutApprovedProd() {
  return menuAmazonSpapiPutListings_({ mode: 'prod', source: 'approved' });
}

/**
 * @param {{mode:string, source?:string, silent?:boolean, skipProdConfirmation?:boolean, offerFulfillment?:string, inventoryMode?:string}} opt
 *   source: 'child_ck'（第1段）| 'approved'（第2段）| 'offer_ck'（Dレ点相乗り）
 *   offerFulfillment: 'mfn' | 'fba'（D選択。相乗り自己発／相乗りFBA）
 *   inventoryMode: 'ZERO'（既定）| 'MASTER'（承認②・ALLOW_MASTER_QTY必須）
 * @return {{ok:boolean, reason?:string, runId?:string, count?:number, fail?:number, batchId?:string}}
 */
function menuAmazonSpapiPutListings_(opt) {
  opt = opt || {};
  var mode = String(opt.mode || 'dry_run');
  var source = String(opt.source || 'child_ck');
  var isApproved = source === 'approved';
  var isOfferCheckbox = source === 'offer_ck';
  var silent = !!opt.silent;
  var skipProdConfirmation = !!opt.skipProdConfirmation;
  var offerFulfillment = String(opt.offerFulfillment || 'mfn').toLowerCase() === 'fba' ? 'fba' : 'mfn';
  var inventoryMode = String(opt.inventoryMode || 'ZERO').toUpperCase() === 'MASTER' ? 'MASTER' : 'ZERO';
  var isProd = mode === 'prod';
  var stepName;
  var functionName;
  var runIdPrefix;
  if (isApproved) {
    stepName = isProd ? 'AmazonSpapiPutApprovedProd' : 'AmazonSpapiPutApprovedDryRun';
    functionName = isProd ? 'menuAmazonSpapiPutApprovedProd' : 'menuAmazonSpapiPutApprovedDryRun';
    runIdPrefix = 'SPAPI_PUT_APPR_' + (isProd ? 'PROD_' : 'DRY_');
  } else if (isOfferCheckbox) {
    stepName = isProd ? 'AmazonSpapiPutOfferCheckboxProd' : 'AmazonSpapiPutOfferCheckboxDryRun';
    functionName = 'menuAmazonSpapiPutListings_';
    runIdPrefix = 'SPAPI_PUT_OFFER_CK_' + (isProd ? 'PROD_' : 'DRY_');
  } else {
    stepName = isProd ? 'AmazonSpapiPutProd' : 'AmazonSpapiPutDryRun';
    functionName = isProd ? 'menuAmazonSpapiPutProd' : 'menuAmazonSpapiPutDryRun';
    runIdPrefix = 'SPAPI_PUT_' + (isProd ? 'PROD_' : 'DRY_');
  }
  var runId = runIdPrefix +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' +
    String(Utilities.getUuid()).replace(/-/g, '').substring(0, 6);

  Logger.log('[' + stepName + '] runId=' + runId + ' functionName=' + functionName +
    ' state=PENDING mode=' + mode + ' source=' + source +
    ' inventoryMode=' + inventoryMode +
    (isOfferCheckbox ? ' offerFulfillment=' + offerFulfillment : ''));

  if (!getBoolScriptProperty_(
        APPROVAL_AMAZON_SPAPI_PUT_PROP,
        (typeof AMAZON_PROD_DEFAULT_PUT_ENABLED_ !== 'undefined')
          ? AMAZON_PROD_DEFAULT_PUT_ENABLED_
          : true
      )) {
    var off = 'SP-API GAS直呼びは無効です。Script Properties の ' +
      APPROVAL_AMAZON_SPAPI_PUT_PROP +
      ' を true にするか、キーを削除してください（未設定時はON・明示falseで緊急停止）。';
    return amazonSpapiPutFail_(stepName, runId, off, silent);
  }

  if (isProd && !getBoolScriptProperty_(
        APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD_PROP,
        (typeof AMAZON_PROD_DEFAULT_ALLOW_PROD_ !== 'undefined')
          ? AMAZON_PROD_DEFAULT_ALLOW_PROD_
          : true
      )) {
    var noProd = 'prod には ' + APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD_PROP +
      '=true が必要です（未設定時はON。明示falseで停止。開始前確認は残ります）。';
    return amazonSpapiPutFail_(stepName, runId, noProd, silent);
  }

  if (inventoryMode === 'MASTER' &&
      !getBoolScriptProperty_(APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY_PROP, false)) {
    var noMaster = 'マスタ在庫送信には ' + APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY_PROP +
      '=true が必要です（承認②・既定は無効）。在庫0で出すか Property を有効にしてください。';
    return amazonSpapiPutFail_(stepName, runId, noMaster, silent);
  }

  var maxItems = Math.floor(getNumberScriptProperty_(APPROVAL_AMAZON_SPAPI_PUT_MAX_PROP, 10));
  if (maxItems < 1) maxItems = 1;
  if (maxItems > 50) maxItems = 50;
  // 相乗り: 既定 quantity=0。マスタqty経路ON時のみ FORCE_QTY_0 を緩和
  var useMasterQty = isOfferCheckbox && inventoryMode === 'MASTER';
  var forceQty0 = useMasterQty
    ? false
    : (isOfferCheckbox
      ? true
      : getBoolScriptProperty_(APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_PROP, true));

  var creds;
  try {
    creds = amazonSpapiPutLoadCreds_();
  } catch (eCred) {
    return amazonSpapiPutFail_(stepName, runId,
      String(eCred && eCred.message ? eCred.message : eCred), silent);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterCtx;
  try {
    masterCtx = amazonApprovalLv4LoadMasterContext_(ss);
  } catch (eLoad) {
    return amazonSpapiPutFail_(stepName, runId,
      String(eLoad && eLoad.message ? eLoad.message : eLoad), silent);
  }

  var collected;
  try {
    if (isApproved) {
      collected = amazonSpapiPutCollectApprovedItems_(masterCtx, forceQty0);
    } else if (isOfferCheckbox) {
      collected = amazonSpapiPutCollectOfferCheckboxItems_(
        masterCtx, forceQty0, isProd, offerFulfillment, inventoryMode);
    } else {
      collected = amazonSpapiPutCollectChildCkItems_(masterCtx, forceQty0);
    }
  } catch (eCol) {
    return amazonSpapiPutFail_(stepName, runId,
      String(eCol && eCol.message ? eCol.message : eCol), silent);
  }

  var batchId = collected.batchId ? String(collected.batchId) : '';
  var parentSkip = isApproved
    ? (collected.parentOnlySkipped || 0)
    : (collected.parentCkOnly || 0);

  if (!collected.items.length) {
    var none;
    if (isApproved) {
      none = '承認①済 Amazon から有効な出品候補が0件です。\n' +
        'batchId=' + (batchId || '(なし)') +
        ' / スキップ=' + collected.skipped.length +
        ' / 親行のみスキップ=' + parentSkip + '\n\n' +
        (typeof amazonSpapiExportFormatSkipDetails_ === 'function'
          ? amazonSpapiExportFormatSkipDetails_(collected.skipped)
          : '') +
        '\n先に承認キューで amazon を承認①してください。親行のみ（子SKU空）はスキップされます。';
    } else {
      none = '出品CK付きの子SKU行から有効な出品候補が0件です。\n' +
        'レ点子行=' + collected.targetRows1.length +
        ' / スキップ=' + collected.skipped.length +
        ' / 親レ点のみ除外=' + parentSkip + '\n\n' +
        (typeof amazonSpapiExportFormatSkipDetails_ === 'function'
          ? amazonSpapiExportFormatSkipDetails_(collected.skipped)
          : '');
    }
    return amazonSpapiPutFail_(stepName, runId, none, silent);
  }

  if (collected.items.length > maxItems) {
    var overHint = isApproved
      ? '対象を絞るか Property '
      : 'レ点を減らすか Property ';
    var over = '候補が max_items=' + maxItems + ' を超えています（' + collected.items.length +
      '件）。' + overHint + APPROVAL_AMAZON_SPAPI_PUT_MAX_PROP + ' を見直してください。' +
      (batchId ? '\nbatchId=' + batchId : '');
    return amazonSpapiPutFail_(stepName, runId, over, silent);
  }

  if (isProd && !silent && !skipProdConfirmation) {
    var ui = SpreadsheetApp.getUi();
    var confTitle = isApproved
      ? 'SP-API prod 確認（承認①済）'
      : (isOfferCheckbox ? 'D 相乗りprod確認（人間レ点）' : 'SP-API prod 確認');
    var confBody = '本番 PUT を実行します。件数=' + collected.items.length +
      (batchId ? '\nbatchId=' + batchId : '') +
      '\n在庫FORCE_0=' + forceQty0 +
      '\ninventoryMode=' + inventoryMode +
      (isOfferCheckbox
        ? '\n相乗り=' + (offerFulfillment === 'fba' ? 'FBA' : '自己発') +
          '／保存列=' + amazonSpapiPutOfferSellerSkuHeader_(offerFulfillment)
        : '') +
      '\nSKU例=' + collected.items[0].sku +
      (useMasterQty
        ? '\n' + amazonSpapiPutFormatQtyConfirm_(collected.items, offerFulfillment)
        : '') +
      (parentSkip ? '\n親行のみスキップ=' + parentSkip : '') +
      '\n続行しますか？';
    var conf = ui.alert(confTitle, confBody, ui.ButtonSet.OK_CANCEL);
    if (conf !== ui.Button.OK) {
      Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED cancelled_by_user');
      return { ok: false, reason: 'ユーザー取消', runId: runId, batchId: batchId };
    }
  }

  Logger.log('[' + stepName + '] runId=' + runId + ' state=RUNNING items=' + collected.items.length +
    ' maxItems=' + maxItems + ' forceQty0=' + forceQty0 +
    ' inventoryMode=' + inventoryMode +
    (batchId ? ' batchId=' + batchId : '') +
    ' parentSkip=' + parentSkip);

  var accessToken;
  try {
    accessToken = amazonSpapiPutRequestLwa_(creds.clientId, creds.clientSecret, creds.refreshToken);
  } catch (eLwa) {
    return amazonSpapiPutFail_(stepName, runId,
      'LWA失敗: ' + String(eLwa && eLwa.message ? eLwa.message : eLwa), silent);
  }
  Logger.log('[' + stepName + '] runId=' + runId + ' LWA OK');

  var okCount = 0;
  var failCount = 0;
  var lines = [];
  var adviceParts = [];
  var validationPreview = !isProd;

  for (var i = 0; i < collected.items.length; i++) {
    var item = collected.items[i];
    Logger.log('[' + stepName + '] --- item ' + (i + 1) + '/' + collected.items.length +
      ' sku=' + item.sku + ' asin=' + item.asin +
      (item.note ? ' note=' + item.note : ''));
    try {
      var one = amazonSpapiPutProcessOne_(creds, accessToken, item, validationPreview);
      if (one.ok && validationPreview && isOfferCheckbox &&
          (String(one.status || '').toUpperCase() !== 'VALID' || Number(one.issueCount || 0) !== 0)) {
        one = {
          ok: false,
          reason: 'dry_runは status=VALID かつ issues=0 のみ保存可（status=' +
            String(one.status || '') + ' issues=' + Number(one.issueCount || 0) + ')',
          advice: one.advice || ''
        };
      }
      if (one.ok) {
        // dry_run=VALID時／prod=成功時に相乗りSKU列へ保存（通常運用はprod直可）
        if (isOfferCheckbox) {
          amazonSpapiPutPersistOfferSellerSku_(masterCtx, item, runId);
        }
        okCount++;
        lines.push('OK ' + item.sku + ' status=' + one.status + ' issues=' + one.issueCount);
      } else {
        failCount++;
        lines.push('FAIL ' + item.sku + ' ' + one.reason);
        if (one.advice) adviceParts.push(String(one.advice));
      }
      Logger.log('[' + stepName + '] ' + lines[lines.length - 1]);
      if (one.advice) {
        Logger.log('[' + stepName + '] advice sku=' + item.sku + ' ' +
          String(one.advice).replace(/\n/g, ' | '));
      }
    } catch (eOne) {
      failCount++;
      var er = String(eOne && eOne.message ? eOne.message : eOne);
      lines.push('FAIL ' + item.sku + ' ' + er);
      Logger.log('[' + stepName + '] FAIL sku=' + item.sku + ' ' + er);
    }
    Utilities.sleep(300);
  }

  var adviceBlock = adviceParts.length
    ? amazonSpapiPutDedupeAdvice_(adviceParts).join('\n')
    : '';
  var sourceLabel = isApproved ? '承認①済' : (isOfferCheckbox ? '相乗り子SKUレ点' : '子SKUレ点');
  var doneMsg = 'SP-API GAS ' + mode + '（' + sourceLabel + '）完了。\n' +
    (batchId ? 'batchId=' + batchId + '\n' : '') +
    'ok=' + okCount + ' fail=' + failCount + ' total=' + collected.items.length +
    (collected.skipped.length ? ' / スキップ' + collected.skipped.length : '') +
    (parentSkip ? ' / 親行スキップ' + parentSkip : '') +
    '\nrunId=' + runId + '\n\n' + lines.slice(0, 12).join('\n') +
    (adviceBlock ? '\n\n【おすすめ・次アクション】\n' + adviceBlock : '') +
    '\n\n本番は ' + APPROVAL_AMAZON_SPAPI_PUT_PROP +
    (isProd ? '／' + APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD_PROP : '') +
    ' を常時ON（未設定=ON）で運用可。緊急停止時のみ明示 false。';

  Logger.log('[' + stepName + '] runId=' + runId + ' state=DONE ok=' + okCount +
    ' fail=' + failCount + (batchId ? ' batchId=' + batchId : ''));
  if (!silent) {
    try { SpreadsheetApp.getUi().alert(doneMsg); } catch (eUi) {}
  }

  var failReason = '';
  if (!(failCount === 0 && okCount > 0)) {
    failReason = '既存相乗り PUT 失敗（ok=' + okCount + ' fail=' + failCount + '）\n' +
      'runId=' + runId + '\n' + lines.slice(0, 8).join('\n') +
      (adviceBlock ? '\n\n【おすすめ・次アクション】\n' + adviceBlock : '');
  }

  return {
    ok: failCount === 0 && okCount > 0,
    runId: runId,
    count: okCount,
    fail: failCount,
    skipped: collected.skipped.length,
    parentSkip: parentSkip,
    batchId: batchId || undefined,
    mode: mode,
    source: source,
    reason: failReason || undefined,
    advice: adviceBlock || undefined,
    detail: doneMsg
  };
}

/**
 * @return {{items:Array, skipped:Array, targetRows1:Array, parentCkOnly:number}}
 */
function amazonSpapiPutCollectChildCkItems_(masterCtx, forceQty0) {
  var ckName = (typeof CHECKBOX_HEADER_NAME !== 'undefined') ? CHECKBOX_HEADER_NAME : '出品CK';
  var iCk = masterCtx.col[ckName];
  var iChild = masterCtx.col['子SKU'];
  if (iCk == null) throw new Error('マスタに「' + ckName + '」列がありません。');
  if (iChild == null) throw new Error('マスタに「子SKU」列がありません。');

  var targetRows1 = [];
  var parentCkOnly = 0;
  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    var row = masterCtx.values[r] || [];
    if (!amazonSpapiPutCheckboxIsTrue_(row[iCk])) continue;
    var childSku = String(row[iChild] != null ? row[iChild] : '').trim();
    if (!childSku) {
      parentCkOnly++;
      continue;
    }
    targetRows1.push(r + 1);
  }

  var items = [];
  var skipped = [];
  for (var i = 0; i < targetRows1.length; i++) {
    var row1 = targetRows1[i];
    var built = amazonSpapiExportBuildItemFromRow_(masterCtx, row1 - 1, forceQty0);
    if (built.ok) {
      items.push(built.item);
    } else {
      skipped.push({ row1: row1, reason: built.reason });
    }
  }
  return {
    items: items,
    skipped: skipped,
    targetRows1: targetRows1,
    parentCkOnly: parentCkOnly
  };
}

/**
 * Dレ点本線の既存相乗り行だけを収集する。
 * SKU列空なら dry_run／prod とも生成。既存値の as/af 正規化も両モード可。
 * 列への保存は PUT 成功後（dry_run=VALID／prod=成功）。
 * inventoryMode=ZERO（既定）: quantity=0。MASTER: マスタ「在庫数」生値（不正は送信停止）。
 * FBA は body で quantity 非送信（item.quantity はログ用）。
 * @param {string=} offerFulfillment mfn | fba
 * @param {string=} inventoryMode ZERO | MASTER
 */
function amazonSpapiPutCollectOfferCheckboxItems_(masterCtx, forceQty0, isProd, offerFulfillment, inventoryMode) {
  if (masterCtx.col[AMAZON_OFFER_ASIN_HEADER_] == null) {
    throw new Error('マスタに「' + AMAZON_OFFER_ASIN_HEADER_ + '」列がありません。');
  }
  if (typeof amazonCheckboxMainlineInspect_ !== 'function') {
    throw new Error('レ点本線の行分類関数がありません。AmazonApprovalExport.js の反映を確認してください。');
  }
  var useMasterQty = String(inventoryMode || 'ZERO').toUpperCase() === 'MASTER';
  forceQty0 = !useMasterQty;
  var fulfillment = String(offerFulfillment || 'mfn').toLowerCase() === 'fba' ? 'fba' : 'mfn';
  var skuHeader = amazonSpapiPutOfferSellerSkuHeader_(fulfillment);
  if (masterCtx.col[skuHeader] == null) {
    throw new Error('マスタに「' + skuHeader + '」列がありません。' +
      (fulfillment === 'fba'
        ? ' ヘッダに Amazon相乗りSKU_FBA を追加してください（デュアル Phase1・FBA専用）。'
        : ' ヘッダに Amazon相乗りSKU を確認してください（自己発専用）。'));
  }
  Logger.log('[AmazonSpapiPut] offer_ck collect fulfillment=' + fulfillment +
    ' skuHeader=' + skuHeader + ' isProd=' + !!isProd);
  var inspected = amazonCheckboxMainlineInspect_(masterCtx, {
    includeNew: false,
    includeOffer: true
  });

  var items = [];
  var skipped = [];
  var targetRows1 = [];
  var qtyErrors = [];
  for (var i = 0; i < inspected.offerRows.length; i++) {
    var one = inspected.offerRows[i];
    var rowIndex0 = one.rowIndex0;
    one.offerFulfillment = fulfillment;

    var built = amazonSpapiExportBuildItemFromRow_(masterCtx, rowIndex0, true);
    if (!built.ok) {
      skipped.push({ row1: one.row1, reason: built.reason });
      continue;
    }
    var asin = amazonSpapiPutResolveDirectOfferAsin_(masterCtx, rowIndex0, one.parentSku);
    if (!asin) {
      skipped.push({ row1: one.row1, reason: 'ASINコード空/不正（競合店ASIN・URLは使用しません）' });
      continue;
    }

    var qty = 0;
    if (useMasterQty) {
      // FBAはAPIへquantityを送らないが、確認ダイアログ用にマスタ生値を読む。不正は停止。
      var qtyRes = amazonSpapiPutReadMasterQtyStrict_(masterCtx, rowIndex0);
      if (!qtyRes.ok) {
        qtyErrors.push({ row1: one.row1, reason: qtyRes.reason, sku: one.childSku || '' });
        continue;
      }
      qty = qtyRes.qty;
    }

    // 系統別列のみ読取（自己発=NF／FBA=_FBA。他系統列は触らない）
    // 通常運用はprod直可。列空なら生成、s残存なら as/af 正規化（dry_runも同等）
    var savedSku = amazonApprovalLv4Cell_(masterCtx, rowIndex0, skuHeader);
    var sellerSku = savedSku;
    if (!sellerSku) {
      var generated = amazonSpapiPutBuildOfferSellerSku_(masterCtx, rowIndex0, one, asin);
      if (!generated.ok) {
        skipped.push({ row1: one.row1, reason: generated.reason });
        continue;
      }
      sellerSku = generated.sku;
    } else {
      var normalizedSaved = amazonSpapiPutEnsureOfferFulfillmentSuffix_(
        sellerSku, fulfillment);
      if (!normalizedSaved.ok) {
        skipped.push({ row1: one.row1, reason: normalizedSaved.reason });
        continue;
      }
      sellerSku = normalizedSaved.sku;
      var matchFx = amazonSpapiPutOfferSkuMatchesFulfillment_(sellerSku, fulfillment, skuHeader);
      if (!matchFx.ok) {
        skipped.push({ row1: one.row1, reason: matchFx.reason });
        continue;
      }
      var checked = amazonSpapiPutValidateSavedOfferSellerSku_(sellerSku, asin, skuHeader);
      if (!checked.ok) {
        skipped.push({ row1: one.row1, reason: checked.reason });
        continue;
      }
      // 記号だけ直した場合は成功後に再保存できるよう「未保存扱い」にする
      if (sellerSku !== savedSku) {
        savedSku = '';
      }
    }

    built.item.sku = sellerSku;
    built.item.asin = asin;
    built.item.quantity = qty;
    built.item.fulfillmentChannel = fulfillment === 'fba' ? 'AMAZON_JP' : 'DEFAULT';
    built.item.rowIndex0 = rowIndex0;
    built.item.offerSkuWasSaved = !!savedSku;
    built.item.offerSellerSkuHeader = skuHeader;
    built.item.note = (built.item.note ? built.item.note + ';' : '') +
      'source=offer_ck;masterRow=' + one.row1 + ';fulfillment=' + fulfillment +
      ';skuHeader=' + skuHeader +
      ';inventoryMode=' + (useMasterQty ? 'MASTER' : 'ZERO') +
      ';forceQty0=' + (forceQty0 ? '1' : '0');
    items.push(built.item);
    targetRows1.push(one.row1);
  }
  if (qtyErrors.length) {
    throw new Error('マスタ在庫（在庫数）が不正のため送信停止しました。\n' +
      amazonSpapiExportFormatSkipDetails_(qtyErrors));
  }
  return {
    items: items,
    skipped: skipped,
    targetRows1: targetRows1,
    parentCkOnly: inspected.parentCkOnly
  };
}

/**
 * マスタ「在庫数」生値を厳密読取。空・非数・負は不正（0は合法）。
 * @return {{ok:boolean, qty?:number, reason?:string}}
 */
function amazonSpapiPutReadMasterQtyStrict_(masterCtx, rowIndex0) {
  if (masterCtx.col['在庫数'] == null) {
    return { ok: false, reason: 'マスタに「在庫数」列がありません' };
  }
  var qRaw = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '在庫数');
  if (qRaw === '' || qRaw == null) {
    return { ok: false, reason: '在庫数空（マスタ在庫送信時は必須）' };
  }
  var qNum = Number(qRaw);
  if (isNaN(qNum)) {
    return { ok: false, reason: '在庫数非数 raw=' + String(qRaw) };
  }
  if (qNum < 0) {
    return { ok: false, reason: '在庫数負 raw=' + String(qRaw) };
  }
  return { ok: true, qty: Math.floor(qNum) };
}

/**
 * 開始前確認用: 件数・SKU例・qty合計／内訳。FBAは「API非送信」と注記。
 * @param {Array} items
 * @param {string} fulfillment mfn|fba
 */
function amazonSpapiPutFormatQtyConfirm_(items, fulfillment) {
  var isFba = String(fulfillment || '').toLowerCase() === 'fba';
  var sum = 0;
  var lines = [];
  var n = Math.min(items.length, 8);
  for (var i = 0; i < items.length; i++) {
    var q = Math.floor(Number(items[i].quantity) || 0);
    sum += q;
    if (i < n) {
      lines.push((i + 1) + '. ' + items[i].sku + ' qty=' + q);
    }
  }
  if (items.length > n) lines.push('…他' + (items.length - n) + '件');
  return '送信qty合計=' + sum +
    (isFba ? '（FBAのためAPIにはquantity非送信・確認用）' : '（自己発・マスタ在庫数生値）') +
    '\n内訳:\n' + lines.join('\n');
}

/** 相乗り先ASINはN列「ASINコード」のみ。子空時だけ同じ親の親行へフォールバック。 */
function amazonSpapiPutResolveDirectOfferAsin_(masterCtx, rowIndex0, parentSku) {
  var asin = amazonApprovalLv4Cell_(masterCtx, rowIndex0, AMAZON_OFFER_ASIN_HEADER_);
  if (amazonApprovalLv4LooksLikeAsin_(asin)) return String(asin).toUpperCase();
  if (parentSku) {
    var parentRow = amazonSpapiExportFindParentRow_(masterCtx, parentSku);
    if (parentRow != null) {
      asin = amazonApprovalLv4Cell_(masterCtx, parentRow, AMAZON_OFFER_ASIN_HEADER_);
      if (amazonApprovalLv4LooksLikeAsin_(asin)) return String(asin).toUpperCase();
    }
  }
  return '';
}

function amazonSpapiPutBuildOfferSellerSku_(masterCtx, rowIndex0, one, asin) {
  var childSku = String(one.childSku || '').trim();
  var a = String(asin || '').trim().toUpperCase();
  var fulfillment = amazonSpapiPutResolveOfferFulfillment_(one);

  // 中央がすでに今回のASINならJAN置換はせず、発送記号だけ as/af に揃える
  // 例: sanky-B01N5A6ESU-19s13 → sanky-B01N5A6ESU-19as13
  if (childSku.indexOf('-' + a + '-') >= 0) {
    var ensuredReuse = amazonSpapiPutEnsureOfferFulfillmentSuffix_(childSku, fulfillment);
    if (!ensuredReuse.ok) return ensuredReuse;
    var headerReuse = amazonSpapiPutOfferSellerSkuHeader_(fulfillment);
    var reused = amazonSpapiPutValidateSavedOfferSellerSku_(ensuredReuse.sku, a, headerReuse);
    if (!reused.ok) return reused;
    return { ok: true, sku: ensuredReuse.sku, reused: true };
  }

  var jan = amazonApprovalLv4Cell_(masterCtx, rowIndex0, 'JANコード');
  var original = amazonApprovalLv4Cell_(masterCtx, rowIndex0, 'オリジナルカタログ商品名');
  var tokens = [];
  if (jan) tokens.push(jan);
  if (original && original !== 'JAN重複時に手入力' && original !== jan) tokens.push(original);
  var sku = '';
  var hitCount = 0;
  for (var i = 0; i < tokens.length; i++) {
    var needle = '-' + tokens[i] + '-';
    var pos = childSku.indexOf(needle);
    var nextPos = pos;
    while (nextPos >= 0) {
      hitCount++;
      nextPos = childSku.indexOf(needle, nextPos + needle.length);
    }
    if (pos >= 0) {
      sku = childSku.substring(0, pos + 1) + a +
        childSku.substring(pos + needle.length - 1);
    }
  }
  if (hitCount !== 1 || !sku) {
    return { ok: false, reason: '子SKU中央識別値の置換対象が' + hitCount +
      '件（JAN/オリジナル名を完全一致で特定できません） sku=' + childSku };
  }

  var ensured = amazonSpapiPutEnsureOfferFulfillmentSuffix_(sku, fulfillment);
  if (!ensured.ok) return ensured;
  var header = amazonSpapiPutOfferSellerSkuHeader_(fulfillment);
  var valid = amazonSpapiPutValidateSavedOfferSellerSku_(ensured.sku, a, header);
  if (!valid.ok) return valid;
  return { ok: true, sku: ensured.sku };
}

/** @return {string} mfn | fba */
function amazonSpapiPutResolveOfferFulfillment_(one) {
  var fulfillment = String((one && one.offerFulfillment) || '').toLowerCase();
  if (fulfillment === 'fba' || fulfillment === 'mfn') return fulfillment;
  var shipping = String((one && one.shipping) || '').trim().toLowerCase();
  return shipping === '相乗りfba' ? 'fba' : 'mfn';
}

/**
 * 相乗りSKUの発送記号を as/af に揃える。
 * 既に as/af があれば維持。s13/f13 などは D選択に従い as13/af13 へ。
 * @return {{ok:boolean, sku?:string, reason?:string}}
 */
function amazonSpapiPutEnsureOfferFulfillmentSuffix_(sku, fulfillment) {
  var s = String(sku || '').trim();
  if (!s) return { ok: false, reason: 'Amazon相乗りSKUが空です' };
  if (/a[fs]\d+/i.test(s)) return { ok: true, sku: s };

  var wantAf = String(fulfillment || '').toLowerCase() === 'fba';
  var re = /([0-9])([sf])([0-9]+)/gi;
  var last = null;
  var match;
  while ((match = re.exec(s)) !== null) {
    last = match;
  }
  if (!last) {
    return { ok: false, reason: '子SKUの発送記号 s/f を特定できません: ' + s };
  }
  var letter = wantAf ? 'af' : 'as';
  var replacement = last[1] + letter + last[3];
  var out = s.substring(0, last.index) + replacement + s.substring(last.index + last[0].length);
  return { ok: true, sku: out };
}

/**
 * 保存済みSKUの as/af が系統と一致するか（デュアル Phase1）。
 * @return {{ok:boolean, reason?:string}}
 */
function amazonSpapiPutOfferSkuMatchesFulfillment_(sku, fulfillment, skuHeader) {
  var s = String(sku || '');
  var header = String(skuHeader || amazonSpapiPutOfferSellerSkuHeader_(fulfillment));
  var wantAf = String(fulfillment || '').toLowerCase() === 'fba';
  var hasAf = /af\d+/i.test(s);
  var hasAs = /as\d+/i.test(s);
  if (wantAf && hasAs && !hasAf) {
    return {
      ok: false,
      reason: header + 'に自己発記号(as)のSKUがあります。FBA列へ移すか削除してください: ' + s
    };
  }
  if (!wantAf && hasAf && !hasAs) {
    return {
      ok: false,
      reason: header + 'にFBA記号(af)のSKUがあります。Amazon相乗りSKU_FBA へ移すか削除してください: ' + s
    };
  }
  return { ok: true };
}

function amazonSpapiPutValidateSavedOfferSellerSku_(sku, asin, skuHeader) {
  var s = String(sku || '').trim();
  var a = String(asin || '').trim().toUpperCase();
  var header = String(skuHeader || AMAZON_OFFER_SELLER_SKU_HEADER_);
  if (s.indexOf('-' + a + '-') < 0) {
    return { ok: false, reason: header +
      'のASINが今回のASINコードと不一致: ' + s + ' / ' + a };
  }
  if (s.length > 40) return { ok: false, reason: header + 'が40文字超: ' + s.length };
  if (!/^[\x21-\x7E]+$/.test(s)) return { ok: false, reason: header + 'に全角/空白あり: ' + s };
  // as1 / af1 / as12 / af12 などを許容
  if (!/a[fs]\d+/i.test(s)) {
    return { ok: false, reason: header + 'に as/af 発送記号がありません: ' + s };
  }
  return { ok: true, sku: s };
}

function amazonSpapiPutPersistOfferSellerSku_(masterCtx, item, runId) {
  var fulfill = String(item.fulfillmentChannel || '') === 'AMAZON_JP' ? 'fba' : 'mfn';
  var header = String(item.offerSellerSkuHeader || amazonSpapiPutOfferSellerSkuHeader_(fulfill));
  var col = masterCtx.col[header];
  if (col == null || item.rowIndex0 == null) {
    throw new Error(header + 'の保存先を解決できません。');
  }
  var old = String(masterCtx.values[item.rowIndex0][col] == null
    ? '' : masterCtx.values[item.rowIndex0][col]).trim();
  if (old && old !== item.sku) {
    var upgraded = amazonSpapiPutEnsureOfferFulfillmentSuffix_(old, fulfill);
    if (!(upgraded.ok && upgraded.sku === item.sku)) {
      throw new Error(header + 'が実行中に変更されました。行' +
        (item.rowIndex0 + 1) + ' old=' + old + ' new=' + item.sku);
    }
  }
  if (!old || old !== item.sku) {
    masterCtx.sheet.getRange(item.rowIndex0 + 1, col + 1).setValue(item.sku);
    masterCtx.values[item.rowIndex0][col] = item.sku;
    Logger.log('[AmazonSpapiPut] runId=' + runId + ' state=DONE action=SAVE_OFFER_SKU row=' +
      (item.rowIndex0 + 1) + ' header=' + header + ' sku=' + item.sku + ' asin=' + item.asin +
      ' fulfillment=' + fulfill +
      (old && old !== item.sku ? ' upgradedFrom=' + old : ''));
  }
}

/**
 * 第2段: 21-⑨（menuAmazonSpapiExportApprovedItemsCsv）と同一の対象抽出。
 * @return {{items:Array, skipped:Array, batchId:string, parentOnlySkipped:number, targetRows1:Array}}
 */
function amazonSpapiPutCollectApprovedItems_(masterCtx, forceQty0) {
  if (typeof approvalQueueGetLatestApprovedAmazon_ !== 'function') {
    throw new Error('approvalQueueGetLatestApprovedAmazon_ がありません。');
  }
  var loaded = approvalQueueGetLatestApprovedAmazon_();
  var lines = (loaded && loaded.lines) ? loaded.lines : [];
  var batchId = (loaded && loaded.batch && loaded.batch.batchId)
    ? String(loaded.batch.batchId)
    : '';
  if (!loaded || !loaded.found || !lines.length) {
    return {
      items: [],
      skipped: [{ reason: 'APPROVED の Amazon 明細なし' }],
      batchId: batchId,
      parentOnlySkipped: 0,
      targetRows1: []
    };
  }

  var items = [];
  var skipped = [];
  var seenSku = {};
  var parentOnlySkipped = 0;
  var targetRows1 = [];

  for (var i = 0; i < lines.length; i++) {
    var L = lines[i];
    if (String(L.mall) !== 'amazon' || String(L.lineStatus) !== 'APPROVED') continue;
    var parentSku = String(L.parentSku || '').trim();
    var childSku = String(L.childSku || '').trim();
    if (!childSku) {
      parentOnlySkipped++;
      skipped.push({ parentSku: parentSku, reason: '親行のみ（子SKU空）' });
      continue;
    }
    if (seenSku[childSku]) continue;
    seenSku[childSku] = true;

    var mr = amazonApprovalLv4FindMasterRow_(masterCtx, parentSku, childSku);
    if (!mr) {
      skipped.push({ parentSku: parentSku, childSku: childSku, reason: 'マスタ行なし' });
      continue;
    }
    var built = amazonSpapiExportBuildItemFromRow_(masterCtx, mr.rowIndex0, forceQty0);
    if (built.ok) {
      built.item.note = (built.item.note ? built.item.note + ';' : '') + 'approved=' + batchId;
      items.push(built.item);
      targetRows1.push(mr.rowIndex0 + 1);
    } else {
      skipped.push({
        parentSku: parentSku,
        childSku: childSku,
        reason: built.reason || 'build失敗'
      });
    }
  }

  return {
    items: items,
    skipped: skipped,
    batchId: batchId,
    parentOnlySkipped: parentOnlySkipped,
    targetRows1: targetRows1
  };
}

function amazonSpapiPutCheckboxIsTrue_(v) {
  if (typeof amazonSpapiExportCheckboxIsTrue_ === 'function') {
    return amazonSpapiExportCheckboxIsTrue_(v);
  }
  return v === true || String(v).toUpperCase() === 'TRUE' || String(v).trim() === '1';
}

/**
 * @return {{clientId:string, clientSecret:string, refreshToken:string, sellerId:string, marketplaceId:string, endpoint:string}}
 */
function amazonSpapiPutLoadCreds_() {
  var props = PropertiesService.getScriptProperties();
  var clientId = String(props.getProperty(SPAPI_LWA_CLIENT_ID_PROP) || '').trim();
  var clientSecret = String(props.getProperty(SPAPI_LWA_CLIENT_SECRET_PROP) || '').trim();
  var refreshToken = String(props.getProperty(SPAPI_REFRESH_TOKEN_PROP) || '').trim();
  var sellerId = String(props.getProperty(SPAPI_SELLER_ID_PROP) || '').trim();
  var marketplaceId = String(props.getProperty(SPAPI_MARKETPLACE_ID_PROP) || '').trim() ||
    SPAPI_DEFAULT_MARKETPLACE_;
  var endpointRaw = String(props.getProperty(SPAPI_ENDPOINT_PROP) || '').trim();
  var endpoint = amazonSpapiPutNormalizeEndpoint_(endpointRaw || SPAPI_DEFAULT_ENDPOINT_);

  var missing = [];
  if (!clientId) missing.push(SPAPI_LWA_CLIENT_ID_PROP);
  if (!clientSecret) missing.push(SPAPI_LWA_CLIENT_SECRET_PROP);
  if (!refreshToken) missing.push(SPAPI_REFRESH_TOKEN_PROP);
  if (!sellerId) missing.push(SPAPI_SELLER_ID_PROP);
  if (missing.length) {
    throw new Error('Script Properties 不足: ' + missing.join(', ') +
      '（秘密はチャット・Gitに書かない）');
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

/**
 * SPAPI_ENDPOINT 正規化。
 * Excel／手入力で https:\host になる誤記を https://host に直す（UrlFetch「無効な引数」対策）。
 * @param {string} endpoint
 * @return {string}
 */
function amazonSpapiPutNormalizeEndpoint_(endpoint) {
  var s = String(endpoint || '').trim();
  if (!s) s = SPAPI_DEFAULT_ENDPOINT_;
  // https:\host や http:\host → https://host
  s = s.replace(/^(https?):\\+/i, '$1://');
  // https:/host（スラッシュ1本）→ https://host
  s = s.replace(/^(https?):\/(?!\/)/i, '$1://');
  // バックスラッシュ残りを除去
  s = s.replace(/\\/g, '/');
  // 末尾スラッシュ除去
  s = s.replace(/\/+$/, '');
  if (!/^https?:\/\//i.test(s)) {
    s = SPAPI_DEFAULT_ENDPOINT_;
  }
  return s;
}

/**
 * @return {string} access_token
 */
function amazonSpapiPutRequestLwa_(clientId, clientSecret, refreshToken) {
  var payload = {
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret
  };
  var resp = UrlFetchApp.fetch(SPAPI_LWA_TOKEN_URL_, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded;charset=UTF-8',
    payload: payload,
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var text = resp.getContentText() || '';
  if (code !== 200) {
    throw new Error('HTTP ' + code + ' ' + text.substring(0, 300));
  }
  var body = JSON.parse(text);
  if (!body.access_token) throw new Error('access_token なし');
  return String(body.access_token);
}

/**
 * P4b等の読取共有: LWA＋creds。PUT経路は変更しない。
 * @return {{creds:Object, accessToken:string}}
 */
function amazonSpapiPutAcquireAccess_() {
  var creds = amazonSpapiPutLoadCreds_();
  var accessToken = amazonSpapiPutRequestLwa_(
    creds.clientId, creds.clientSecret, creds.refreshToken);
  return { creds: creds, accessToken: accessToken };
}

/**
 * SP-API GET（Listings以外）。path は先頭スラッシュ付き。
 * @param {Object} creds
 * @param {string} accessToken
 * @param {string} path
 * @param {Object=} query
 * @return {{code:number, text:string, json:?Object}}
 */
function amazonSpapiPutHttpGet_(creds, accessToken, path, query) {
  var qsParts = [];
  query = query || {};
  for (var k in query) {
    if (!query.hasOwnProperty(k)) continue;
    if (query[k] == null || query[k] === '') continue;
    qsParts.push(encodeURIComponent(k) + '=' + encodeURIComponent(String(query[k])));
  }
  var url = String(creds.endpoint || '').replace(/\/+$/, '') + path +
    (qsParts.length ? ('?' + qsParts.join('&')) : '');
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'x-amz-access-token': accessToken,
      'user-agent': SPAPI_USER_AGENT_,
      'accept': 'application/json'
    },
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var text = resp.getContentText() || '';
  var json = null;
  if (code === 200 && text) {
    try { json = JSON.parse(text); } catch (eParse) { json = null; }
  }
  return { code: code, text: text, json: json };
}

/**
 * @param {Object} creds
 * @param {string} accessToken
 * @param {{sku:string, asin:string, price:number, quantity:number}} item
 * @param {boolean} validationPreview
 * @return {{ok:boolean, status?:string, issueCount?:number, reason?:string, http?:number}}
 */
function amazonSpapiPutProcessOne_(creds, accessToken, item, validationPreview) {
  var sku = String(item.sku || '').trim();
  var asin = String(item.asin || '').trim().toUpperCase();
  var price = Number(item.price);
  var quantity = Math.floor(Number(item.quantity));
  if (!sku) return { ok: false, reason: 'sku空' };
  if (!/^[A-Z0-9]{10}$/.test(asin)) return { ok: false, reason: 'ASIN不正' };
  if (!(price > 0)) return { ok: false, reason: 'price不正' };
  if (quantity < 0) return { ok: false, reason: 'quantity不正' };

  // 事前 GET（存在確認・ログ用。404でも PUT は続行）
  var getRes = amazonSpapiPutFetchListings_(creds, accessToken, sku, 'get');
  Logger.log('[AmazonSpapiPut] GET HTTP ' + getRes.code + ' sku=' + sku);

  var body = amazonSpapiPutBuildOfferBody_(
    creds.marketplaceId,
    asin,
    price,
    quantity,
    item.fulfillmentChannel || 'DEFAULT'
  );
  var putRes = amazonSpapiPutFetchListings_(creds, accessToken, sku, 'put', body, validationPreview);
  var status = '';
  var issueCount = 0;
  var putBody = null;
  try {
    putBody = putRes.text ? JSON.parse(putRes.text) : null;
  } catch (eParse) {
    putBody = null;
  }
  if (putBody) {
    status = String(putBody.status || '');
    var issues = putBody.issues || [];
    issueCount = issues.length;
    var hasError = false;
    for (var i = 0; i < issues.length; i++) {
      var sev = String((issues[i] && issues[i].severity) || '').toUpperCase();
      if (sev === 'ERROR') {
        hasError = true;
        break;
      }
    }
    if (putRes.code >= 200 && putRes.code < 300 && !hasError) {
      return { ok: true, status: status || 'OK', issueCount: issueCount, http: putRes.code };
    }
    var firstErr = '';
    for (var j = 0; j < Math.min(issues.length, 3); j++) {
      var iss = issues[j] || {};
      firstErr += (firstErr ? '; ' : '') + String(iss.code || '') + ':' + String(iss.message || '');
    }
    var advice = amazonSpapiPutBuildIssuesAdvice_(issues, item.fulfillmentChannel);
    return {
      ok: false,
      status: status,
      issueCount: issueCount,
      http: putRes.code,
      reason: 'HTTP ' + putRes.code + ' status=' + status + (firstErr ? ' ' + firstErr : ''),
      advice: advice
    };
  }
  if (putRes.code >= 200 && putRes.code < 300) {
    return { ok: true, status: 'OK', issueCount: 0, http: putRes.code };
  }
  return {
    ok: false,
    http: putRes.code,
    reason: 'HTTP ' + putRes.code + ' ' + String(putRes.text || '').substring(0, 200)
  };
}

/**
 * FBA compliance 属性を付けるか（既定 ON）。
 */
function amazonSpapiPutFbaComplianceEnabled_() {
  return getBoolScriptProperty_(APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_PROP, true);
}

/**
 * FBA（AMAZON_JP）時に電池・危険物の安全既定を attributes へ追加。
 * @param {Object} attributes
 * @param {string} marketplaceId
 */
function amazonSpapiPutAppendFbaComplianceAttrs_(attributes, marketplaceId) {
  if (!attributes || !marketplaceId) return;
  // 公式 safety_and_compliance 系。JP enum は dry_run で確定／追随。
  attributes.batteries_required = [
    { value: false, marketplace_id: marketplaceId }
  ];
  attributes.batteries_included = [
    { value: false, marketplace_id: marketplaceId }
  ];
  attributes.supplier_declared_dg_hz_regulation = [
    { value: 'not_applicable', marketplace_id: marketplaceId }
  ];
}

/**
 * issues から許可リスト内のおすすめ文を作る。未知は原文のみ促す。
 * @param {Array} issues
 * @param {string=} fulfillmentChannel
 * @return {string}
 */
function amazonSpapiPutBuildIssuesAdvice_(issues, fulfillmentChannel) {
  var isFba = String(fulfillmentChannel || '').toUpperCase() === 'AMAZON_JP' ||
    String(fulfillmentChannel || '').toLowerCase() === 'fba';
  var tips = [];
  var unknown = [];
  var list = issues || [];
  for (var i = 0; i < list.length; i++) {
    var iss = list[i] || {};
    var code = String(iss.code || '');
    var msg = String(iss.message || '');
    var attrs = iss.attributeNames || [];
    var blob = (code + ' ' + msg + ' ' + attrs.join(' ')).toLowerCase();
    var known = false;
    if (code === '90220' || /電池|batter/.test(blob)) {
      tips.push('・電池: おすすめ batteries_required=false / batteries_included=false（非電池商品の安全既定）');
      known = true;
    }
    if (code === '90220' || /危険物|hazmat|dg_hz|supplier_declared/.test(blob)) {
      tips.push('・危険物: おすすめ supplier_declared_dg_hz_regulation=not_applicable（該当なし相当）');
      known = true;
    }
    if (!known && (code || msg)) {
      unknown.push(code + ':' + msg);
    }
  }
  var out = [];
  if (tips.length) {
    out.push(amazonSpapiPutDedupeAdvice_(tips).join('\n'));
    if (isFba) {
      out.push('※FBA経路では compliance 属性を自動付与しています（Property ' +
        APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_PROP + '）。');
      out.push('※値を確認したうえで、必ず再度 dry_run → VALID 後にのみ prod（折衷運用）。');
    } else {
      out.push('※自己発で同エラーなら、FBA用属性の要否を切り分けてください。');
    }
  }
  if (unknown.length) {
    out.push('・許可リスト外のエラー（自動補完しません）: ' + unknown.slice(0, 3).join(' / '));
  }
  if (!out.length && list.length) {
    out.push('・issues あり。原文を確認し、dry_run で再検証してください（自動補完なし）。');
  }
  return out.join('\n');
}

function amazonSpapiPutDedupeAdvice_(parts) {
  var seen = {};
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var p = String(parts[i] || '').trim();
    if (!p || seen[p]) continue;
    seen[p] = true;
    out.push(p);
  }
  return out;
}

function amazonSpapiPutBuildOfferBody_(marketplaceId, asin, price, quantity, fulfillmentChannel) {
  var channel = String(fulfillmentChannel || 'DEFAULT').trim().toUpperCase();
  if (channel !== 'AMAZON_JP') channel = 'DEFAULT';
  var fulfillment = {
    fulfillment_channel_code: channel,
    marketplace_id: marketplaceId
  };
  // FBA在庫はAmazon管理。quantityはMFNだけに送る。
  if (channel === 'DEFAULT') fulfillment.quantity = quantity;
  var attributes = {
    condition_type: [
      { value: 'new_new', marketplace_id: marketplaceId }
    ],
    merchant_suggested_asin: [
      { value: asin, marketplace_id: marketplaceId }
    ],
    purchasable_offer: [
      {
        currency: 'JPY',
        our_price: [{ schedule: [{ value_with_tax: price }] }],
        marketplace_id: marketplaceId
      }
    ],
    fulfillment_availability: [fulfillment]
  };
  if (channel === 'AMAZON_JP' && amazonSpapiPutFbaComplianceEnabled_()) {
    amazonSpapiPutAppendFbaComplianceAttrs_(attributes, marketplaceId);
    Logger.log('[AmazonSpapiPut] FBA compliance attrs ON' +
      ' (batteries_required/included=false, dg_hz=not_applicable)');
  }
  return {
    productType: 'PRODUCT',
    requirements: 'LISTING_OFFER_ONLY',
    attributes: attributes
  };
}

/**
 * @return {{code:number, text:string}}
 */
function amazonSpapiPutFetchListings_(creds, accessToken, sku, method, body, validationPreview) {
  var path = '/listings/2021-08-01/items/' +
    amazonSpapiPutEncodePath_(creds.sellerId) + '/' +
    amazonSpapiPutEncodePath_(sku);
  var qs = 'marketplaceIds=' + encodeURIComponent(creds.marketplaceId);
  if (method === 'get') {
    qs += '&includedData=' + encodeURIComponent('summaries,attributes,issues');
  } else if (validationPreview) {
    qs += '&mode=VALIDATION_PREVIEW';
  }
  var url = creds.endpoint + path + '?' + qs;
  // UrlFetchApp は Host 手動指定不可（「無効な引数」になる）。LWA は x-amz-access-token のみ必須。
  var headers = {
    'x-amz-access-token': accessToken,
    'user-agent': SPAPI_USER_AGENT_,
    'accept': 'application/json',
    'content-type': 'application/json'
  };
  var opts = {
    method: method === 'get' ? 'get' : 'put',
    headers: headers,
    muteHttpExceptions: true
  };
  if (method !== 'get') {
    opts.payload = JSON.stringify(body);
  }
  var resp = UrlFetchApp.fetch(url, opts);
  return { code: resp.getResponseCode(), text: resp.getContentText() || '' };
}

/** RFC3986: -_.~ はエンコードしない（Python quote(safe="-_.~") 相当） */
function amazonSpapiPutEncodePath_(s) {
  return encodeURIComponent(String(s))
    .replace(/[!'()*]/g, function (c) {
      return '%' + c.charCodeAt(0).toString(16).toUpperCase();
    });
}

function amazonSpapiPutFail_(stepName, runId, reason, silent) {
  Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' +
    String(reason).replace(/\n/g, ' | '));
  if (!silent) {
    try { SpreadsheetApp.getUi().alert(String(reason)); } catch (e0) {}
  }
  return { ok: false, reason: reason, runId: runId };
}

/**
 * 競合定時用: Amazon 在庫を sellerSku 1件だけ GET。マスタ非書込。トリガーなし。
 * Listings fulfillmentAvailability を正。取れなければ FBA summaries。
 */
function menuReadAmazonInventoryOneSku() {
  var ui = bTryUi_();
  if (!ui) {
    Logger.log('[競合ストア] Amazon在庫1SKUはUIから指定');
    return;
  }
  var resp = ui.prompt(
    'Amazon在庫 1 SKU 読取',
    '子SKU（sellerSku）を1つ。マスタには書きません。全件取得しません。',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var sku = String(resp.getResponseText() || '').trim().split(/[\s,;]+/)[0];
  if (!sku) {
    ui.alert('SKUが空です。');
    return;
  }
  var out = amazonSpapiReadInventoryOneSku_(sku);
  Logger.log('[競合ストア] Amazon在庫1SKU sku=' + sku + ' qty=' + out.qty +
    ' source=' + out.source + ' http=' + out.http + ' マスタ非書');
  ui.alert(
    'SKU=' + sku + '\n在庫=' + (out.qty == null ? '(なし)' : out.qty) +
    '\n取得元=' + out.source + '\nHTTP=' + out.http +
    '\nマスタは未変更です。'
  );
}

function amazonSpapiReadInventoryOneSku_(sku) {
  var acc = amazonSpapiPutAcquireAccess_();
  var creds = acc.creds;
  var token = acc.accessToken;
  var listingsPath = '/listings/2021-08-01/items/' +
    amazonSpapiPutEncodePath_(creds.sellerId) + '/' +
    amazonSpapiPutEncodePath_(sku);
  var listings = amazonSpapiPutHttpGet_(creds, token, listingsPath, {
    marketplaceIds: creds.marketplaceId,
    includedData: 'fulfillmentAvailability,summaries'
  });
  var qty = amazonSpapiInventoryQtyFromListingsJson_(listings.json);
  if (qty != null) {
    return { sku: sku, qty: qty, source: 'listings', http: listings.code };
  }
  var fba = amazonSpapiPutHttpGet_(creds, token, '/fba/inventory/v1/summaries', {
    details: 'true',
    granularityType: 'Marketplace',
    granularityId: creds.marketplaceId,
    marketplaceIds: creds.marketplaceId,
    sellerSkus: sku
  });
  var fbaQty = amazonSpapiInventoryQtyFromFbaJson_(fba.json, sku);
  return {
    sku: sku,
    qty: fbaQty,
    source: fbaQty != null ? 'fba_summaries' : 'none',
    http: String(listings.code) + '/' + String(fba.code)
  };
}

function amazonSpapiInventoryQtyFromListingsJson_(json) {
  if (!json) return null;
  var arr = json.fulfillmentAvailability || json.fulfillment_availability || [];
  var i;
  var found = null;
  for (i = 0; i < arr.length; i++) {
    var q = arr[i] && arr[i].quantity;
    if (q == null || q === '') continue;
    var n = Number(q);
    if (!isFinite(n)) continue;
    if (found == null || n > found) found = n;
  }
  return found;
}

function amazonSpapiInventoryQtyFromFbaJson_(json, sku) {
  if (!json) return null;
  var rows = json.inventorySummaries || [];
  var i;
  for (i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    if (sku && String(row.sellerSku || '').trim() !== String(sku).trim()) continue;
    var q = row.totalQuantity;
    if (q == null || q === '') continue;
    var n = Number(q);
    if (isFinite(n)) return n;
  }
  return null;
}
