/**
 * SP-API スプシ橋渡し — マスタ子SKUレ点 / 承認①済 → Drive CSV
 *
 * - GAS から SP-API は呼ばない（CSV書き出しのみ）
 * - 21-⑧: 出品CK付きの**子SKU行のみ**（親レ点のみでは出さない。行選択は廃止）
 * - ローカル tools/spapi_listings_write が dry_run/prod（v1.3 は Drive 自動取得可）
 * 承認: docs/org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md／LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md
 * 手順: docs/org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md
 *
 * Script Properties:
 *   APPROVAL_AMAZON_SPAPI_EXPORT_ENABLED … 既定 false
 *   APPROVAL_AMAZON_SPAPI_EXPORT_MAX_ITEMS … 既定 5（1〜50）
 *   APPROVAL_AMAZON_SPAPI_EXPORT_FORCE_QTY_0 … 既定 true（在庫0強制）
 *   APPROVAL_AMAZON_SPAPI_EXPORT_FOLDER_ID … 空なら Lv4 GENERATED フォルダ流用
 */

var APPROVAL_AMAZON_SPAPI_EXPORT_PROP = 'APPROVAL_AMAZON_SPAPI_EXPORT_ENABLED';
var APPROVAL_AMAZON_SPAPI_EXPORT_MAX_PROP = 'APPROVAL_AMAZON_SPAPI_EXPORT_MAX_ITEMS';
var APPROVAL_AMAZON_SPAPI_EXPORT_FORCE_QTY_PROP = 'APPROVAL_AMAZON_SPAPI_EXPORT_FORCE_QTY_0';
var APPROVAL_AMAZON_SPAPI_EXPORT_FOLDER_PROP = 'APPROVAL_AMAZON_SPAPI_EXPORT_FOLDER_ID';

/**
 * メニュー 21-⑧: マスタの出品CK付き**子SKU行** → SP-API items CSV（Drive）
 * 親レ点のみ・行選択は対象外。出品＝レ点（子）。
 * @param {{silent?:boolean}=} opt
 * @return {{ok:boolean, reason?:string, fileUrl?:string, runId?:string, count?:number}}
 */
function menuAmazonSpapiExportItemsCsv(opt) {
  opt = opt || {};
  var silent = !!opt.silent;
  var stepName = 'AmazonSpapiExportItemsCsv';
  var functionName = 'menuAmazonSpapiExportItemsCsv';
  var runId = 'SPAPI_EXPORT_' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' +
    String(Utilities.getUuid()).replace(/-/g, '').substring(0, 6);

  Logger.log('[' + stepName + '] runId=' + runId + ' functionName=' + functionName + ' state=PENDING');

  if (!getBoolScriptProperty_(APPROVAL_AMAZON_SPAPI_EXPORT_PROP, false)) {
    var off = 'SP-API CSV出力は無効です。Script Properties の ' +
      APPROVAL_AMAZON_SPAPI_EXPORT_PROP + ' を true にしてください（既定は無効）。' +
      'SP-API直呼びは対象外です。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + off);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    }
    return { ok: false, reason: off, runId: runId };
  }

  var maxItems = Math.floor(getNumberScriptProperty_(APPROVAL_AMAZON_SPAPI_EXPORT_MAX_PROP, 5));
  if (maxItems < 1) maxItems = 1;
  if (maxItems > 50) maxItems = 50;
  var forceQty0 = getBoolScriptProperty_(APPROVAL_AMAZON_SPAPI_EXPORT_FORCE_QTY_PROP, true);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterCtx;
  try {
    masterCtx = amazonApprovalLv4LoadMasterContext_(ss);
  } catch (eLoad) {
    var loadFail = String(eLoad && eLoad.message ? eLoad.message : eLoad);
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + loadFail);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(loadFail); } catch (e1) {}
    }
    return { ok: false, reason: loadFail, runId: runId };
  }

  var ckName = (typeof CHECKBOX_HEADER_NAME !== 'undefined') ? CHECKBOX_HEADER_NAME : '出品CK';
  var iCk = masterCtx.col[ckName];
  var iChild = masterCtx.col['子SKU'];
  if (iCk == null) {
    var noCk = 'マスタに「' + ckName + '」列がありません。出品＝レ点のため必須です。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + noCk);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(noCk); } catch (eCk) {}
    }
    return { ok: false, reason: noCk, runId: runId };
  }
  if (iChild == null) {
    var noChild = 'マスタに「子SKU」列がありません。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + noChild);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(noChild); } catch (eCh) {}
    }
    return { ok: false, reason: noChild, runId: runId };
  }

  // マスタ全行から「子SKUあり＋出品CK」のみ（親レ点のみは出さない）
  var targetRows1 = [];
  var parentCkOnly = 0;
  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    var row = masterCtx.values[r] || [];
    if (!amazonSpapiExportCheckboxIsTrue_(row[iCk])) continue;
    var childSku = String(row[iChild] != null ? row[iChild] : '').trim();
    if (!childSku) {
      parentCkOnly++;
      continue;
    }
    targetRows1.push(r + 1);
  }

  if (!targetRows1.length) {
    var noSel = '出品CK付きの子SKU行がありません。\n' +
      '・出品対象の子行に「' + ckName + '」を付けてから 21-⑧ を実行\n' +
      '・親行だけのレ点では出力しません（親レ点のみ検知=' + parentCkOnly + '）\n' +
      '・行選択は不要（マスタ全行を走査）';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + noSel.replace(/\n/g, ' | '));
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(noSel); } catch (e3) {}
    }
    return { ok: false, reason: noSel, runId: runId };
  }

  Logger.log('[' + stepName + '] runId=' + runId + ' state=RUNNING childCkRows=' + targetRows1.length +
    ' parentCkOnlySkipped=' + parentCkOnly + ' maxItems=' + maxItems + ' forceQty0=' + forceQty0);

  var items = [];
  var skipped = [];
  for (var i = 0; i < targetRows1.length; i++) {
    var row1 = targetRows1[i];
    var rowIndex0 = row1 - 1;
    var built = amazonSpapiExportBuildItemFromRow_(masterCtx, rowIndex0, forceQty0);
    if (built.ok) {
      items.push(built.item);
    } else {
      skipped.push({ row1: row1, reason: built.reason });
    }
  }

  if (items.length > maxItems) {
    var over = '出力候補が max_items=' + maxItems + ' を超えています（' + items.length +
      '件）。レ点を減らすか Property ' + APPROVAL_AMAZON_SPAPI_EXPORT_MAX_PROP + ' を見直してください。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + over);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(over); } catch (e4) {}
    }
    return { ok: false, reason: over, runId: runId, count: items.length };
  }

  if (!items.length) {
    var none = '有効な行が0件です（レ点付き子行はあったが ASIN／販売価格amazon 不足）。\n' +
      'レ点子行数=' + targetRows1.length + ' / スキップ=' + skipped.length +
      (parentCkOnly ? (' / 親レ点のみ除外=' + parentCkOnly) : '') + '\n\n' +
      amazonSpapiExportFormatSkipDetails_(skipped) +
      '\n必要な列: 子SKU＋出品CK ／ ASINコードor競合店ASIN ／ 販売価格amazon（正の数）';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + none.replace(/\n/g, ' | '));
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(none); } catch (e5) {}
    }
    return { ok: false, reason: none, runId: runId };
  }

  var csv = amazonSpapiExportItemsToCsv_(items);
  var fileName = runId + '_SPAPI_ITEMS.csv';
  var folder = amazonSpapiExportGetFolder_(ss);
  var file = folder.createFile(Utilities.newBlob(csv, 'text/csv', fileName));
  var fileUrl = file.getUrl();

  Logger.log('[' + stepName + '] runId=' + runId + ' state=DONE count=' + items.length +
    ' skipped=' + skipped.length + ' parentCkOnly=' + parentCkOnly +
    ' fileName=' + fileName + ' fileUrl=' + fileUrl);
  for (var s = 0; s < Math.min(skipped.length, 10); s++) {
    Logger.log('[' + stepName + '] skipped row1=' + skipped[s].row1 + ' reason=' + skipped[s].reason);
  }

  var msg = 'SP-API items CSV を Drive に出力しました（子SKUレ点のみ）。\n' +
    '件数: ' + items.length + '（スキップ ' + skipped.length +
    (parentCkOnly ? '／親レ点のみ除外 ' + parentCkOnly : '') + '）\n' +
    'ファイル: ' + fileName + '\n' +
    'URL: ' + fileUrl + '\n\n' +
    (skipped.length
      ? ('スキップ詳細:\n' + amazonSpapiExportFormatSkipDetails_(skipped) + '\n')
      : '') +
    '次: ローカルで\n' +
    '  python spapi_listings_write.py --fetch-drive --mode dry_run\n' +
    '（GAS から SP-API は呼びません）';
  if (!silent) {
    try { SpreadsheetApp.getUi().alert(msg); } catch (e6) {}
  }

  return {
    ok: true,
    runId: runId,
    count: items.length,
    skipped: skipped.length,
    parentCkOnlySkipped: parentCkOnly,
    fileName: fileName,
    fileUrl: fileUrl
  };
}

/**
 * 出品CK判定（boolean TRUE / 文字列 TRUE 両対応）
 */
function amazonSpapiExportCheckboxIsTrue_(v) {
  if (typeof amazonAiCheckboxIsTrue_ === 'function') {
    return amazonAiCheckboxIsTrue_(v);
  }
  return v === true || String(v).toUpperCase() === 'TRUE' || String(v).trim() === '1';
}

/**
 * メニュー 21-⑨: 最新承認①済 Amazon 行 → SP-API items CSV（Drive）
 * ASIN 必須（相乗り／既存オファー向け）。新規カタログ専用行はスキップ。
 * @param {{silent?:boolean}=} opt
 * @return {{ok:boolean, reason?:string, fileUrl?:string, runId?:string, count?:number}}
 */
function menuAmazonSpapiExportApprovedItemsCsv(opt) {
  opt = opt || {};
  var silent = !!opt.silent;
  var stepName = 'AmazonSpapiExportApprovedItemsCsv';
  var functionName = 'menuAmazonSpapiExportApprovedItemsCsv';
  var runId = 'SPAPI_EXPORT_APPR_' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' +
    String(Utilities.getUuid()).replace(/-/g, '').substring(0, 6);

  Logger.log('[' + stepName + '] runId=' + runId + ' functionName=' + functionName + ' state=PENDING');

  if (!getBoolScriptProperty_(APPROVAL_AMAZON_SPAPI_EXPORT_PROP, false)) {
    var off = 'SP-API CSV出力は無効です。Script Properties の ' +
      APPROVAL_AMAZON_SPAPI_EXPORT_PROP + ' を true にしてください（既定は無効）。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + off);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    }
    return { ok: false, reason: off, runId: runId };
  }

  if (typeof approvalQueueGetLatestApprovedAmazon_ !== 'function') {
    var noAq = 'approvalQueueGetLatestApprovedAmazon_ がありません。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + noAq);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(noAq); } catch (e1) {}
    }
    return { ok: false, reason: noAq, runId: runId };
  }

  var maxItems = Math.floor(getNumberScriptProperty_(APPROVAL_AMAZON_SPAPI_EXPORT_MAX_PROP, 5));
  if (maxItems < 1) maxItems = 1;
  if (maxItems > 50) maxItems = 50;
  var forceQty0 = getBoolScriptProperty_(APPROVAL_AMAZON_SPAPI_EXPORT_FORCE_QTY_PROP, true);

  var loaded;
  try {
    loaded = approvalQueueGetLatestApprovedAmazon_();
  } catch (eLoadAq) {
    var aqFail = String(eLoadAq && eLoadAq.message ? eLoadAq.message : eLoadAq);
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + aqFail);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(aqFail); } catch (e2) {}
    }
    return { ok: false, reason: aqFail, runId: runId };
  }

  var lines = (loaded && loaded.lines) ? loaded.lines : [];
  var batchId = (loaded && loaded.batch && loaded.batch.batchId)
    ? String(loaded.batch.batchId)
    : '';
  if (!loaded || !loaded.found || !lines.length) {
    var noneAq = 'APPROVED の Amazon 明細がありません。先に承認キューで amazon を承認①してください。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + noneAq);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(noneAq); } catch (e3) {}
    }
    return { ok: false, reason: noneAq, runId: runId };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterCtx;
  try {
    masterCtx = amazonApprovalLv4LoadMasterContext_(ss);
  } catch (eLoad) {
    var loadFail = String(eLoad && eLoad.message ? eLoad.message : eLoad);
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + loadFail);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(loadFail); } catch (e4) {}
    }
    return { ok: false, reason: loadFail, runId: runId };
  }

  Logger.log('[' + stepName + '] runId=' + runId + ' state=RUNNING batchId=' + batchId +
    ' lines=' + lines.length + ' maxItems=' + maxItems + ' forceQty0=' + forceQty0);

  var items = [];
  var skipped = [];
  var seenSku = {};

  for (var i = 0; i < lines.length; i++) {
    var L = lines[i];
    if (String(L.mall) !== 'amazon' || String(L.lineStatus) !== 'APPROVED') continue;
    var parentSku = String(L.parentSku || '').trim();
    var childSku = String(L.childSku || '').trim();
    // 親行のみ（子SKU空）はバリエーション親のことが多い → スキップ（子行を待つ）
    if (!childSku) {
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
    } else {
      skipped.push({
        parentSku: parentSku,
        childSku: childSku,
        reason: built.reason || 'build失敗'
      });
    }
  }

  if (items.length > maxItems) {
    var over = '承認①済の出力候補が max_items=' + maxItems + ' を超えています（' + items.length +
      '件）。Property ' + APPROVAL_AMAZON_SPAPI_EXPORT_MAX_PROP + ' を見直すか対象を絞ってください。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + over);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(over); } catch (e5) {}
    }
    return { ok: false, reason: over, runId: runId, count: items.length };
  }

  if (!items.length) {
    var none = '有効な承認①済行が0件です。\n' +
      'batchId=' + batchId + '\n' +
      'スキップ=' + skipped.length + '\n\n' +
      amazonSpapiExportFormatSkipDetails_(skipped) +
      '\n必要な列: 子SKU ／ ASIN ／ 販売価格amazon（正の数）。親行のみはスキップされます。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + none.replace(/\n/g, ' | '));
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(none); } catch (e6) {}
    }
    return { ok: false, reason: none, runId: runId };
  }

  var csv = amazonSpapiExportItemsToCsv_(items);
  var fileName = runId + '_SPAPI_ITEMS.csv';
  var folder = amazonSpapiExportGetFolder_(ss);
  var file = folder.createFile(Utilities.newBlob(csv, 'text/csv', fileName));
  var fileUrl = file.getUrl();

  Logger.log('[' + stepName + '] runId=' + runId + ' state=DONE count=' + items.length +
    ' skipped=' + skipped.length + ' fileName=' + fileName + ' fileUrl=' + fileUrl);

  var msg = '承認①済 Amazon → SP-API items CSV を Drive に出力しました。\n' +
    'batchId: ' + batchId + '\n' +
    '件数: ' + items.length + '（スキップ ' + skipped.length + '）\n' +
    'ファイル: ' + fileName + '\n' +
    'URL: ' + fileUrl + '\n\n' +
    (skipped.length
      ? ('スキップ詳細:\n' + amazonSpapiExportFormatSkipDetails_(skipped) + '\n')
      : '') +
    '次: ローカルで\n' +
    '  python spapi_listings_write.py --fetch-drive --mode dry_run\n' +
    '（または CSV を手動配置）';
  if (!silent) {
    try { SpreadsheetApp.getUi().alert(msg); } catch (e7) {}
  }

  return {
    ok: true,
    runId: runId,
    batchId: batchId,
    count: items.length,
    skipped: skipped.length,
    fileName: fileName,
    fileUrl: fileUrl
  };
}

/**
 * @return {{ok:boolean, item?:Object, reason?:string, missing?:Array<string>}}
 */
function amazonSpapiExportBuildItemFromRow_(masterCtx, rowIndex0, forceQty0) {
  var parentSku = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '親SKU');
  var childSku = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '子SKU');
  var sellerSku = childSku || parentSku;
  var missing = [];
  var hints = [];

  if (!sellerSku) {
    missing.push('親SKU');
    missing.push('子SKU');
  }

  var asin = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, rowIndex0);
  var asinFromParent = false;
  if (!asin && parentSku) {
    var parentRow = amazonSpapiExportFindParentRow_(masterCtx, parentSku);
    if (parentRow != null) {
      asin = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, parentRow);
      if (asin) asinFromParent = true;
    }
  }
  if (!asin) {
    missing.push('ASINコード/競合店ASIN/URL');
  }

  var priceRaw = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '販売価格amazon');
  var priceNum = Number(priceRaw);
  var priceFromParent = false;
  if (priceRaw === '' || isNaN(priceNum) || priceNum <= 0) {
    if (parentSku) {
      var pr = amazonSpapiExportFindParentRow_(masterCtx, parentSku);
      if (pr != null) {
        priceRaw = amazonApprovalLv4Cell_(masterCtx, pr, '販売価格amazon');
        priceNum = Number(priceRaw);
        if (priceRaw !== '' && !isNaN(priceNum) && priceNum > 0) priceFromParent = true;
      }
    }
  }
  if (priceRaw === '' || isNaN(priceNum) || priceNum <= 0) {
    missing.push('販売価格amazon');
    if (masterCtx.col['販売価格amazon'] == null) {
      hints.push('ヘッダ「販売価格amazon」列が見つかりません');
    } else {
      hints.push('販売価格amazonが空または0以下（卸値列では代用不可）');
    }
  }

  if (missing.length) {
    var reason = '空/不正: ' + missing.join('・');
    if (sellerSku) reason += ' / sku=' + sellerSku;
    if (parentSku && !childSku) reason += '（親行の可能性）';
    if (hints.length) reason += ' / ' + hints.join('; ');
    return { ok: false, reason: reason, missing: missing };
  }

  var qty = 0;
  if (!forceQty0) {
    var qRaw = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '在庫数');
    var qNum = Number(qRaw);
    if (qRaw !== '' && !isNaN(qNum) && qNum >= 0) qty = Math.floor(qNum);
  }

  var note = parentSku ? ('parent=' + parentSku) : '';
  if (asinFromParent) note += (note ? ';' : '') + 'asinFromParent';
  if (priceFromParent) note += (note ? ';' : '') + 'priceFromParent';
  return {
    ok: true,
    item: {
      sku: sellerSku,
      asin: String(asin).toUpperCase(),
      price: priceNum,
      quantity: qty,
      note: note
    }
  };
}

/**
 * スキップ一覧をダイアログ用に整形（最大8件）。
 * @param {Array} skipped
 * @return {string}
 */
function amazonSpapiExportFormatSkipDetails_(skipped) {
  if (!skipped || !skipped.length) return '(スキップ詳細なし)';
  var lines = [];
  var n = Math.min(skipped.length, 8);
  for (var i = 0; i < n; i++) {
    var s = skipped[i] || {};
    var label = '';
    if (s.row1 != null) label = '行' + s.row1;
    else if (s.childSku || s.parentSku) {
      label = 'sku=' + (s.childSku || s.parentSku);
    } else {
      label = '#' + (i + 1);
    }
    lines.push('- ' + label + ': ' + (s.reason || '(理由不明)'));
  }
  if (skipped.length > 8) {
    lines.push('…他 ' + (skipped.length - 8) + ' 件');
  }
  return lines.join('\n');
}

/** @return {number|null} rowIndex0 */
function amazonSpapiExportFindParentRow_(masterCtx, parentSku) {
  var want = String(parentSku || '').trim();
  if (!want) return null;
  var iParent = masterCtx.col['親SKU'];
  var iChild = masterCtx.col['子SKU'];
  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    var row = masterCtx.values[r] || [];
    var p = iParent != null ? String(row[iParent] != null ? row[iParent] : '').trim() : '';
    var c = iChild != null ? String(row[iChild] != null ? row[iChild] : '').trim() : '';
    if (p === want && !c) return r;
  }
  return null;
}

function amazonSpapiExportItemsToCsv_(items) {
  var lines = ['sku,asin,price,quantity,note'];
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    lines.push([
      amazonSpapiExportCsvEscape_(it.sku),
      amazonSpapiExportCsvEscape_(it.asin),
      String(it.price),
      String(it.quantity),
      amazonSpapiExportCsvEscape_(it.note || '')
    ].join(','));
  }
  return lines.join('\r\n') + '\r\n';
}

function amazonSpapiExportCsvEscape_(v) {
  var s = String(v == null ? '' : v);
  if (/[",\r\n]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function amazonSpapiExportGetFolder_(ss) {
  var folderId = PropertiesService.getScriptProperties()
    .getProperty(APPROVAL_AMAZON_SPAPI_EXPORT_FOLDER_PROP);
  if (folderId) {
    try {
      return DriveApp.getFolderById(String(folderId).trim());
    } catch (e) {
      Logger.log('[amazonSpapiExportGetFolder_] invalid folder id, fallback Lv4 folder');
    }
  }
  if (typeof amazonApprovalLv4GetOrCreateFolder_ === 'function') {
    return amazonApprovalLv4GetOrCreateFolder_(ss);
  }
  var name = 'Lv4_Amazon_SPAPI_ITEMS';
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}
