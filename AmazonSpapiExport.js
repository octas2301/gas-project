/**
 * SP-API スプシ橋渡し v1.2 — マスタ選択行 → Drive CSV（Listings用）
 *
 * - GAS から SP-API は呼ばない（CSV書き出しのみ）
 * - ローカル tools/spapi_listings_write が dry_run/prod
 * 承認: docs/org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md
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
 * メニュー 21-⑧: 選択行を SP-API items CSV として Drive へ出力
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

  var sheet = masterCtx.sheet;
  if (ss.getActiveSheet().getSheetId() !== sheet.getSheetId()) {
    var wrong = 'マスタ「' + sheet.getName() + '」を開き、対象行を選択してから実行してください。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + wrong);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(wrong); } catch (e2) {}
    }
    return { ok: false, reason: wrong, runId: runId };
  }

  var ranges = sheet.getActiveRangeList()
    ? sheet.getActiveRangeList().getRanges()
    : [sheet.getActiveRange()];
  var rowSet = {};
  for (var ri = 0; ri < ranges.length; ri++) {
    var rg = ranges[ri];
    if (!rg) continue;
    var start = rg.getRow();
    var n = rg.getNumRows();
    for (var rr = 0; rr < n; rr++) {
      var row1 = start + rr;
      if (row1 <= masterCtx.headerRowIdx + 1) continue;
      rowSet[row1] = true;
    }
  }
  var selectedRows1 = Object.keys(rowSet).map(function (k) { return Number(k); }).sort(function (a, b) {
    return a - b;
  });

  if (!selectedRows1.length) {
    var noSel = 'データ行が選択されていません（ヘッダ以外を選択）。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + noSel);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(noSel); } catch (e3) {}
    }
    return { ok: false, reason: noSel, runId: runId };
  }

  Logger.log('[' + stepName + '] runId=' + runId + ' state=RUNNING selected=' + selectedRows1.length +
    ' maxItems=' + maxItems + ' forceQty0=' + forceQty0);

  var items = [];
  var skipped = [];
  for (var i = 0; i < selectedRows1.length; i++) {
    var row1 = selectedRows1[i];
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
      '件）。選択を減らすか Property ' + APPROVAL_AMAZON_SPAPI_EXPORT_MAX_PROP + ' を見直してください。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + over);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(over); } catch (e4) {}
    }
    return { ok: false, reason: over, runId: runId, count: items.length };
  }

  if (!items.length) {
    var none = '有効な行が0件です（子SKUまたは親SKU・ASIN・販売価格amazon が必要）。skipped=' +
      skipped.length;
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + none +
      ' detail=' + JSON.stringify(skipped).substring(0, 500));
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
    ' skipped=' + skipped.length + ' fileName=' + fileName + ' fileUrl=' + fileUrl);
  for (var s = 0; s < Math.min(skipped.length, 10); s++) {
    Logger.log('[' + stepName + '] skipped row1=' + skipped[s].row1 + ' reason=' + skipped[s].reason);
  }

  var msg = 'SP-API items CSV を Drive に出力しました。\n' +
    '件数: ' + items.length + '（スキップ ' + skipped.length + '）\n' +
    'ファイル: ' + fileName + '\n' +
    'URL: ' + fileUrl + '\n\n' +
    '次: ローカルでダウンロード → tools/spapi_listings_write/items.csv に配置 → dry_run → prod\n' +
    '（GAS から SP-API は呼びません）';
  if (!silent) {
    try { SpreadsheetApp.getUi().alert(msg); } catch (e6) {}
  }

  return {
    ok: true,
    runId: runId,
    count: items.length,
    skipped: skipped.length,
    fileName: fileName,
    fileUrl: fileUrl
  };
}

/**
 * @return {{ok:boolean, item?:Object, reason?:string}}
 */
function amazonSpapiExportBuildItemFromRow_(masterCtx, rowIndex0, forceQty0) {
  var parentSku = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '親SKU');
  var childSku = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '子SKU');
  var sellerSku = childSku || parentSku;
  if (!sellerSku) {
    return { ok: false, reason: '親SKU/子SKU空' };
  }

  var asin = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, rowIndex0);
  if (!asin && parentSku) {
    var parentRow = amazonSpapiExportFindParentRow_(masterCtx, parentSku);
    if (parentRow != null) {
      asin = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, parentRow);
    }
  }
  if (!asin) {
    return { ok: false, reason: 'ASIN無し（ASINコード／競合店ASIN／URL） sku=' + sellerSku };
  }

  var priceRaw = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '販売価格amazon');
  var priceNum = Number(priceRaw);
  if (priceRaw === '' || isNaN(priceNum) || priceNum <= 0) {
    if (parentSku) {
      var pr = amazonSpapiExportFindParentRow_(masterCtx, parentSku);
      if (pr != null) {
        priceRaw = amazonApprovalLv4Cell_(masterCtx, pr, '販売価格amazon');
        priceNum = Number(priceRaw);
      }
    }
  }
  if (priceRaw === '' || isNaN(priceNum) || priceNum <= 0) {
    return { ok: false, reason: '販売価格amazon不正 sku=' + sellerSku };
  }

  var qty = 0;
  if (!forceQty0) {
    var qRaw = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '在庫数');
    var qNum = Number(qRaw);
    if (qRaw !== '' && !isNaN(qNum) && qNum >= 0) qty = Math.floor(qNum);
  }

  var note = parentSku ? ('parent=' + parentSku) : '';
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
