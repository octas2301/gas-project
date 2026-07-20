/**
 * Lv3 承認①済 → Yahoo日中掲載（既存経路オーケストレーション）
 * 要件: docs/org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md
 *
 * - Yahoo.js の出品API／Builder 本体は改変しない（runYahooExport 呼出のみ）
 * - 既存メニュー「▶ 出品実行」は触らない（手動逃げ道）
 * - レ点は案A: 実行直前に承認済み子SKUのみON（他OFF）、終了後に復元
 * - 分割: 主=実働25分／副=ユニーク画像≤50（楽天運用揃え。Yahoo公式バッチ上限ではない）
 *
 * Script Properties:
 *   APPROVAL_YAHOO_LV3_ENABLED … 既定 false
 *   APPROVAL_YAHOO_LV3_APPLY_STOCK … 既定 true（対象子の在庫数を 0/1 に合わせる）
 *   APPROVAL_YAHOO_LV3_SKIP_EXPORT … 既定 false（true なら runYahooExport をスキップ＝ドライラン）
 *   APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_LIMIT … 既定 50（副制約。運用揃え）
 *   APPROVAL_YAHOO_LV3_STATE … レジューム用（自動）
 */

var APPROVAL_YAHOO_LV3_PROP = 'APPROVAL_YAHOO_LV3_ENABLED';
var APPROVAL_YAHOO_LV3_APPLY_STOCK_PROP = 'APPROVAL_YAHOO_LV3_APPLY_STOCK';
var APPROVAL_YAHOO_LV3_SKIP_EXPORT_PROP = 'APPROVAL_YAHOO_LV3_SKIP_EXPORT';
var APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_LIMIT_PROP = 'APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_LIMIT';
var APPROVAL_YAHOO_LV3_STATE_PROP = 'APPROVAL_YAHOO_LV3_STATE';
var APPROVAL_YAHOO_LV3_TIME_MS = 25 * 60 * 1000;
var APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_MAX_DEFAULT = 50;
var APPROVAL_YAHOO_LV3_MAX_TRIGGER_RUNS = 40;
var APPROVAL_YAHOO_LV3_TRIGGER_FN = 'runApprovalYahooLv3FromTrigger';

/**
 * メニュー 20-①: 最新 APPROVED バッチの Yahoo 子をサブバッチ分割して runYahooExport を呼ぶ。
 */
function menuApprovalYahooLv3Run() {
  var fn = 'menuApprovalYahooLv3Run';
  if (!getBoolScriptProperty_(APPROVAL_YAHOO_LV3_PROP, false)) {
    var off = 'Lv3 Yahooは無効です。Script Properties の ' + APPROVAL_YAHOO_LV3_PROP + ' を true にしてください。';
    Logger.log('[' + fn + '] state=FAILED ' + off);
    try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    return;
  }
  try {
    var ui = SpreadsheetApp.getUi();
    var skipExport = getBoolScriptProperty_(APPROVAL_YAHOO_LV3_SKIP_EXPORT_PROP, false);
    var imgLimit = yahooApprovalLv3UniqueImageLimit_();
    var res = ui.alert(
      'Lv3 Yahoo日中掲載',
      '最新の承認①（APPROVED）の Yahoo 子SKUを対象に、既存 runYahooExport を呼び出します。\n' +
        '・手動メニュー「▶ 出品実行」はそのまま残っています\n' +
        '・レ点は一時的に「承認済み子のみ」ONにし、終了後に復元します（案A）\n' +
        '・分割: 主=実働約25分／副=ユニーク画像≤' + imgLimit + '（楽天運用揃え）\n' +
        (skipExport ? '・SKIP_EXPORT=true のため出品APIは呼びません（ドライラン）\n' : '・Yahoo API（画像・editItem・setStock）まで進みます\n') +
        '\n実行しますか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (res !== ui.Button.OK) return;
  } catch (eUi) {}

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var runId = 'LV3_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' +
    ('000000' + Math.floor(Math.random() * 1e6)).slice(-6);
  Logger.log('[' + fn + '] state=RUNNING runId=' + runId);
  try {
    var summary = yahooApprovalLv3Run_(ss, runId, 0, null);
    Logger.log('[' + fn + '] state=DONE runId=' + runId + ' ' + JSON.stringify(summary));
    try {
      SpreadsheetApp.getUi().alert(
        'Lv3 実行結果',
        'runId=' + runId +
          '\nbatchId=' + (summary.batchId || '') +
          '\nサブバッチ完了=' + summary.subBatchesDone +
          '\n子SKU成功=' + summary.childrenDone +
          '\nスキップ=' + summary.skipped +
          '\n継続待ち=' + (summary.willResume ? 'YES（約1分後に自動再開）' : 'NO') +
          '\nメッセージ=' + (summary.message || ''),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e1) {}
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED runId=' + runId + ' ' + ((err && err.message) || err));
    yahooApprovalLv3Mail_('【Lv3 Yahoo】実行失敗', 'runId=' + runId + '\n' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert('Lv3失敗: ' + ((err && err.message) || err)); } catch (e2) {}
  }
}

/** メニュー 20-②: レジューム状態クリア */
function menuApprovalYahooLv3ClearState() {
  try {
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_YAHOO_LV3_STATE_PROP);
    yahooApprovalLv3DeleteTriggers_();
    SpreadsheetApp.getUi().alert('Lv3 実行状態をクリアしました。');
  } catch (e) {
    try { SpreadsheetApp.getUi().alert(String(e && e.message || e)); } catch (e2) {}
  }
}

/** トリガー再開 */
function runApprovalYahooLv3FromTrigger() {
  var fn = APPROVAL_YAHOO_LV3_TRIGGER_FN;
  var stateJson = PropertiesService.getScriptProperties().getProperty(APPROVAL_YAHOO_LV3_STATE_PROP);
  if (!stateJson) return;
  yahooApprovalLv3DeleteTriggers_();
  var state;
  try {
    state = JSON.parse(stateJson);
  } catch (e) {
    return;
  }
  var runCount = Number(state.triggerRunCount || 0) + 1;
  state.triggerRunCount = runCount;
  PropertiesService.getScriptProperties().setProperty(APPROVAL_YAHOO_LV3_STATE_PROP, JSON.stringify(state));
  if (runCount > APPROVAL_YAHOO_LV3_MAX_TRIGGER_RUNS) {
    Logger.log('[' + fn + '] state=FAILED maxTriggerRuns runId=' + state.runId);
    yahooApprovalLv3Mail_('【Lv3 Yahoo】自動再開上限', 'runId=' + state.runId + ' batchId=' + state.batchId);
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_YAHOO_LV3_STATE_PROP);
    return;
  }
  if (!getBoolScriptProperty_(APPROVAL_YAHOO_LV3_PROP, false)) {
    Logger.log('[' + fn + '] state=FAILED disabled');
    return;
  }
  var ss = SpreadsheetApp.openById(state.spreadsheetId);
  Logger.log('[' + fn + '] state=RUNNING runId=' + state.runId + ' nextSub=' + state.nextSubBatchIndex);
  try {
    yahooApprovalLv3Run_(ss, state.runId, state.nextSubBatchIndex || 0, state);
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    yahooApprovalLv3Mail_('【Lv3 Yahoo】再開失敗', 'runId=' + state.runId + '\n' + ((err && err.message) || err));
  }
}

/**
 * @param {Spreadsheet} ss
 * @param {string} runId
 * @param {number} startSubBatchIndex
 * @param {Object|null} resumeState
 * @return {Object}
 */
function yahooApprovalLv3Run_(ss, runId, startSubBatchIndex, resumeState) {
  var fn = 'yahooApprovalLv3Run_';
  var startedAt = Date.now();
  var loaded = (typeof approvalQueueGetLatestApprovedYahoo_ === 'function')
    ? approvalQueueGetLatestApprovedYahoo_()
    : { found: false };
  if (!loaded.found || !loaded.lines || !loaded.lines.length) {
    throw new Error('APPROVED の Yahoo 明細がありません。先に承認キューで承認①してください。');
  }
  var batch = loaded.batch;
  var batchId = batch.batchId;
  var inventoryMode = String(batch.inventoryMode || 'ZERO').toUpperCase() === 'ONE' ? 'ONE' : 'ZERO';
  var stockVal = inventoryMode === 'ONE' ? 1 : 0;

  var masterCtx = yahooApprovalLv3LoadMasterContext_(ss);
  var resolved = yahooApprovalLv3ResolveChildren_(masterCtx, loaded.lines, resumeState && resumeState.doneChildren);
  Logger.log('[' + fn + '] runId=' + runId + ' batchId=' + batchId +
    ' candidates=' + resolved.children.length + ' skipped=' + resolved.skipped.length +
    ' inventoryMode=' + inventoryMode);

  var imgLimit = yahooApprovalLv3UniqueImageLimit_();
  var subBatches = yahooApprovalLv3BuildSubBatches_(masterCtx, resolved.children, imgLimit);
  if (!subBatches.length) {
    return {
      batchId: batchId,
      subBatchesDone: 0,
      childrenDone: 0,
      skipped: resolved.skipped.length,
      willResume: false,
      message: '実行対象の子SKUがありません（スキップのみ）'
    };
  }

  var doneChildren = (resumeState && resumeState.doneChildren) ? resumeState.doneChildren.slice() : [];
  var subBatchesDone = Number(resumeState && resumeState.subBatchesDone) || 0;
  var checkboxSnapshot = null;
  var willResume = false;
  var message = '';

  try {
    checkboxSnapshot = yahooApprovalLv3SnapshotCheckboxes_(masterCtx);
    for (var s = startSubBatchIndex; s < subBatches.length; s++) {
      if ((Date.now() - startedAt) > APPROVAL_YAHOO_LV3_TIME_MS) {
        yahooApprovalLv3SaveState_({
          spreadsheetId: ss.getId(),
          runId: runId,
          batchId: batchId,
          inventoryMode: inventoryMode,
          nextSubBatchIndex: s,
          subBatchesDone: subBatchesDone,
          doneChildren: doneChildren,
          triggerRunCount: resumeState ? resumeState.triggerRunCount : 0
        });
        yahooApprovalLv3SetTrigger_();
        willResume = true;
        message = '時間予算のため中断。サブバッチ ' + (s + 1) + '/' + subBatches.length + ' から再開予定';
        Logger.log('[' + fn + '] state=RETRYING runId=' + runId + ' ' + message);
        break;
      }

      var sub = subBatches[s];
      var subBatchId = batchId + '_Y' + (s + 1);
      Logger.log('[' + fn + '] state=RUNNING runId=' + runId + ' batchId=' + batchId +
        ' subBatchId=' + subBatchId + ' children=' + sub.children.length +
        ' uniqueImages=' + sub.uniqueImageCount + ' uniqueLimit=' + imgLimit);

      try {
        yahooApprovalLv3ApplyPlanACheckboxes_(masterCtx, sub.children);
        if (getBoolScriptProperty_(APPROVAL_YAHOO_LV3_APPLY_STOCK_PROP, true)) {
          yahooApprovalLv3ApplyStock_(masterCtx, sub.children, stockVal);
        }
        SpreadsheetApp.flush();

        if (!getBoolScriptProperty_(APPROVAL_YAHOO_LV3_SKIP_EXPORT_PROP, false)) {
          if (typeof runYahooExport !== 'function') {
            throw new Error('runYahooExport が見つかりません');
          }
          runYahooExport(ss);
        } else {
          Logger.log('[' + fn + '] SKIP_EXPORT=true subBatchId=' + subBatchId);
        }

        for (var c = 0; c < sub.children.length; c++) {
          doneChildren.push(sub.children[c].childSku);
        }
        subBatchesDone++;
        Logger.log('[' + fn + '] state=DONE subBatchId=' + subBatchId + ' children=' + sub.children.length);
      } catch (subErr) {
        Logger.log('[' + fn + '] state=FAILED subBatchId=' + subBatchId + ' ' + ((subErr && subErr.message) || subErr));
        yahooApprovalLv3Mail_(
          '【Lv3 Yahoo】サブバッチ失敗',
          'runId=' + runId + '\nbatchId=' + batchId + '\nsubBatchId=' + subBatchId + '\n' + ((subErr && subErr.message) || subErr)
        );
        throw subErr;
      } finally {
        try {
          yahooApprovalLv3RestoreCheckboxes_(masterCtx, checkboxSnapshot);
          SpreadsheetApp.flush();
        } catch (restErr) {
          Logger.log('[' + fn + '] state=FAILED checkboxRestore ' + ((restErr && restErr.message) || restErr));
          yahooApprovalLv3Mail_(
            '【Lv3 Yahoo】レ点復元失敗',
            'runId=' + runId + '\nbatchId=' + batchId + '\n' + ((restErr && restErr.message) || restErr) +
              '\nマスタの出品CKを目視確認してください。'
          );
          throw restErr;
        }
      }
    }

    if (!willResume) {
      PropertiesService.getScriptProperties().deleteProperty(APPROVAL_YAHOO_LV3_STATE_PROP);
      yahooApprovalLv3DeleteTriggers_();
      message = message || ('全サブバッチ完了 ' + subBatchesDone + '/' + subBatches.length);
    }
  } catch (outer) {
    if (checkboxSnapshot) {
      try {
        yahooApprovalLv3RestoreCheckboxes_(masterCtx, checkboxSnapshot);
        SpreadsheetApp.flush();
      } catch (e3) {
        yahooApprovalLv3Mail_('【Lv3 Yahoo】レ点復元失敗', 'runId=' + runId + '\n' + ((e3 && e3.message) || e3));
      }
    }
    throw outer;
  }

  return {
    batchId: batchId,
    subBatchesDone: subBatchesDone,
    childrenDone: doneChildren.length,
    skipped: resolved.skipped.length,
    willResume: willResume,
    message: message
  };
}

/** @return {number} */
function yahooApprovalLv3UniqueImageLimit_() {
  var n = (typeof getNumberScriptProperty_ === 'function')
    ? getNumberScriptProperty_(APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_LIMIT_PROP, APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_MAX_DEFAULT)
    : APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_MAX_DEFAULT;
  if (!n || n < 1) n = APPROVAL_YAHOO_LV3_UNIQUE_IMAGE_MAX_DEFAULT;
  return n;
}

/** @return {{sheet:Sheet, values:Array, headerRowIdx:number, col:Object, ckName:string}} */
function yahooApprovalLv3LoadMasterContext_(ss) {
  var masterName = (typeof MASTER_SHEET_NAME !== 'undefined')
    ? MASTER_SHEET_NAME
    : ((typeof SHEET_NAME_MASTER !== 'undefined') ? SHEET_NAME_MASTER : '▼商品マスタ(人間作業用)');
  var ckName = (typeof CHECKBOX_HEADER_NAME !== 'undefined') ? CHECKBOX_HEADER_NAME : '出品CK';
  var sheet = ss.getSheetByName(masterName);
  if (!sheet) throw new Error('マスタシートが見つかりません: ' + masterName);
  var actualName = sheet.getName();
  Logger.log(
    '[yahooApprovalLv3LoadMasterContext_] requestedSheet=' + masterName +
      ' actualSheet=' + actualName +
      ' nameMatch=' + (masterName === actualName) +
      ' spreadsheetId=' + ss.getId()
  );
  var values = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  var limit = Math.min(values.length, 25);
  for (var r = 0; r < limit; r++) {
    var row = values[r] || [];
    if (row.indexOf('親SKU') !== -1 && row.indexOf(ckName) !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error('マスタヘッダーが見つかりません');
  var headers = values[headerRowIdx];
  var col = {};
  for (var c = 0; c < headers.length; c++) {
    var key = String(headers[c] || '').trim();
    if (key && col[key] == null) col[key] = c;
  }
  if (col['親SKU'] == null || col[ckName] == null || col['子SKU'] == null) {
    throw new Error('必須列がありません（親SKU / 子SKU / 出品CK）');
  }
  Logger.log(
    '[yahooApprovalLv3LoadMasterContext_] headerRow1Based=' + (headerRowIdx + 1) +
      ' col親SKU=' + col['親SKU'] +
      ' col子SKU=' + col['子SKU'] +
      ' col在庫数=' + (col['在庫数'] != null ? col['在庫数'] : 'MISSING') +
      ' col出品CK=' + col[ckName]
  );
  return { sheet: sheet, values: values, headerRowIdx: headerRowIdx, col: col, ckName: ckName };
}

/**
 * @param {Object} masterCtx
 * @param {Array} lines
 * @param {Array<string>|null} doneChildren
 * @return {{children:Array, skipped:Array}}
 */
function yahooApprovalLv3ResolveChildren_(masterCtx, lines, doneChildren) {
  var doneMap = {};
  if (doneChildren) {
    for (var d = 0; d < doneChildren.length; d++) doneMap[String(doneChildren[d])] = true;
  }
  var col = masterCtx.col;
  var values = masterCtx.values;
  var iParent = col['親SKU'];
  var iChild = col['子SKU'];
  var iStock = col['在庫数'];
  var sheetName = '';
  try { sheetName = masterCtx.sheet.getName(); } catch (eName) { sheetName = '(unknown)'; }
  var children = [];
  var skipped = [];
  var seen = {};

  Logger.log(
    '[yahooApprovalLv3ResolveChildren_] sheet=' + sheetName +
      ' stockColIndex0=' + (iStock != null ? iStock : 'MISSING') +
      ' yahooApprovedLines=' + lines.length
  );

  for (var i = 0; i < lines.length; i++) {
    var L = lines[i];
    if (String(L.mall) !== 'yahoo' || String(L.lineStatus) !== 'APPROVED') continue;
    var childSku = String(L.childSku || '').trim();
    var parentSku = String(L.parentSku || '').trim();
    if (!childSku) {
      skipped.push({ reason: 'SKIPPED_ORPHAN', childSku: '', detail: 'childSku空', queueMasterRow: L.masterRow });
      Logger.log('[yahooApprovalLv3ResolveChildren_] SKIPPED_ORPHAN childSku空 lineId=' + L.lineId);
      continue;
    }
    if (seen[childSku] || doneMap[childSku]) {
      Logger.log('[yahooApprovalLv3ResolveChildren_] SKIP_SEEN_OR_DONE childSku=' + childSku);
      continue;
    }

    var rowIdx = -1;
    var rowMatchHow = '';
    if (L.masterRow != null && Number(L.masterRow) >= 1) {
      var cand = Number(L.masterRow) - 1;
      if (cand > masterCtx.headerRowIdx && cand < values.length) {
        var childAt = String(values[cand][iChild] || '').trim();
        if (childAt === childSku) {
          rowIdx = cand;
          rowMatchHow = 'queueMasterRow';
        } else {
          Logger.log(
            '[yahooApprovalLv3ResolveChildren_] masterRow不一致 childSku=' + childSku +
              ' queueMasterRow=' + L.masterRow +
              ' cellChildSku="' + childAt + '"'
          );
        }
      }
    }
    if (rowIdx < 0) {
      for (var r = masterCtx.headerRowIdx + 1; r < values.length; r++) {
        if (String(values[r][iChild] || '').trim() === childSku) {
          rowIdx = r;
          rowMatchHow = 'scanFirst';
          break;
        }
      }
    }
    if (rowIdx < 0) {
      skipped.push({
        reason: 'SKIPPED_ORPHAN',
        childSku: childSku,
        detail: 'マスタに子行なし',
        queueMasterRow: L.masterRow
      });
      Logger.log(
        '[yahooApprovalLv3ResolveChildren_] SKIPPED_ORPHAN sheet=' + sheetName +
          ' childSku=' + childSku +
          ' queueMasterRow=' + L.masterRow +
          ' detail=マスタに子行なし'
      );
      continue;
    }
    if (!parentSku) {
      parentSku = String(values[rowIdx][iParent] || '').trim();
    }

    var stockRaw = iStock != null ? values[rowIdx][iStock] : '';
    var stockNum = (stockRaw === '' || stockRaw == null) ? null : Number(stockRaw);
    var stockType = (stockRaw === null || stockRaw === undefined) ? String(stockRaw) : typeof stockRaw;
    if (stockNum != null && !isNaN(stockNum) && stockNum > 0) {
      skipped.push({
        reason: 'SKIPPED_IN_STOCK',
        childSku: childSku,
        detail: '在庫数=' + stockNum,
        row1: rowIdx + 1,
        queueMasterRow: L.masterRow,
        stockRaw: stockRaw,
        stockType: stockType,
        rowMatchHow: rowMatchHow
      });
      Logger.log(
        '[yahooApprovalLv3ResolveChildren_] SKIPPED_IN_STOCK sheet=' + sheetName +
          ' childSku=' + childSku +
          ' row1=' + (rowIdx + 1) +
          ' queueMasterRow=' + L.masterRow +
          ' rowMatchHow=' + rowMatchHow +
          ' stockRaw=' + stockRaw +
          ' stockType=' + stockType +
          ' stockNum=' + stockNum
      );
      continue;
    }

    seen[childSku] = true;
    children.push({
      childSku: childSku,
      parentSku: parentSku,
      rowIndex0: rowIdx,
      masterRow: rowIdx + 1,
      lineId: L.lineId
    });
    Logger.log(
      '[yahooApprovalLv3ResolveChildren_] CANDIDATE sheet=' + sheetName +
        ' childSku=' + childSku +
        ' row1=' + (rowIdx + 1) +
        ' rowMatchHow=' + rowMatchHow +
        ' stockRaw=' + stockRaw +
        ' stockType=' + stockType +
        ' stockNum=' + stockNum
    );
  }

  var orphanN = 0;
  var inStockN = 0;
  for (var s = 0; s < skipped.length; s++) {
    if (skipped[s].reason === 'SKIPPED_ORPHAN') orphanN++;
    if (skipped[s].reason === 'SKIPPED_IN_STOCK') inStockN++;
  }
  Logger.log(
    '[yahooApprovalLv3ResolveChildren_] summary sheet=' + sheetName +
      ' candidates=' + children.length +
      ' skipped=' + skipped.length +
      ' SKIPPED_ORPHAN=' + orphanN +
      ' SKIPPED_IN_STOCK=' + inStockN
  );
  return { children: children, skipped: skipped };
}

/**
 * @param {Object} masterCtx
 * @param {Array} children
 * @param {number} uniqueMax
 * @return {Array<{children:Array, uniqueImageCount:number}>}
 */
function yahooApprovalLv3BuildSubBatches_(masterCtx, children, uniqueMax) {
  var batches = [];
  var current = [];
  var keySet = {};

  function flush() {
    if (!current.length) return;
    batches.push({
      children: current.slice(),
      uniqueImageCount: Object.keys(keySet).length
    });
    current = [];
    keySet = {};
  }

  for (var i = 0; i < children.length; i++) {
    var ch = children[i];
    // 現バッチ内の同一親の他子も兄弟画像として見積もる（runYahooExport と同型の近似）
    var keys = yahooApprovalLv3CollectImageKeysForChild_(masterCtx, ch, current);
    var trial = {};
    for (var k in keySet) {
      if (Object.prototype.hasOwnProperty.call(keySet, k)) trial[k] = true;
    }
    for (var j = 0; j < keys.length; j++) trial[keys[j]] = true;
    var trialCount = Object.keys(trial).length;
    if (current.length && trialCount > uniqueMax) {
      flush();
      keys = yahooApprovalLv3CollectImageKeysForChild_(masterCtx, ch, []);
      trial = {};
      for (var j2 = 0; j2 < keys.length; j2++) trial[keys[j2]] = true;
    }
    current.push(ch);
    keySet = trial;
  }
  flush();
  return batches;
}

/**
 * YahooDataBuilder に近い画像集合のユニークキー（自身メイン＋サブ＋同一親でバッチ内の他子メイン）
 * @param {Object} masterCtx
 * @param {{parentSku:string, rowIndex0:number, childSku:string}} child
 * @param {Array} batchPeers 既に同一サブバッチに入っている子
 * @return {Array<string>}
 */
function yahooApprovalLv3CollectImageKeysForChild_(masterCtx, child, batchPeers) {
  var col = masterCtx.col;
  var values = masterCtx.values;
  var iParent = col['親SKU'];
  var iChild = col['子SKU'];
  var keys = {};
  var parentRowIdx = -1;
  for (var r = masterCtx.headerRowIdx + 1; r < values.length; r++) {
    if (String(values[r][iParent] || '').trim() !== child.parentSku) continue;
    var cAt = String(values[r][iChild] || '').trim();
    if (!cAt) {
      parentRowIdx = r;
      break;
    }
  }

  function addCell(rowIdx, colName) {
    var idx = col[colName];
    if (idx == null || rowIdx < 0) return;
    var key = yahooApprovalLv3ImageKey_(values[rowIdx][idx]);
    if (key) keys[key] = true;
  }

  addCell(child.rowIndex0, '楽天メイン画像1');
  for (var m = 1; m <= 8; m++) {
    var subName = '楽天サブ画像' + m;
    var subIdx = col[subName];
    var hasChildSub = false;
    if (subIdx != null) {
      var cv = String(values[child.rowIndex0][subIdx] || '').trim();
      if (cv) {
        addCell(child.rowIndex0, subName);
        hasChildSub = true;
      }
    }
    if (!hasChildSub) addCell(parentRowIdx, subName);
  }

  var peers = batchPeers || [];
  for (var p = 0; p < peers.length; p++) {
    if (peers[p].parentSku !== child.parentSku) continue;
    if (peers[p].childSku === child.childSku) continue;
    addCell(peers[p].rowIndex0, '楽天メイン画像1');
  }
  return Object.keys(keys);
}

/** DriveファイルID優先、なければ正規化URL */
function yahooApprovalLv3ImageKey_(cell) {
  var s = String(cell == null ? '' : cell).trim();
  if (!s) return '';
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/) ||
    s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/) ||
    s.match(/file\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m) return 'drive:' + m[1];
  var bare = s;
  var q = bare.indexOf('?');
  if (q >= 0) bare = bare.substring(0, q);
  return 'url:' + bare.toLowerCase();
}

/**
 * 出品CK列の全データ行スナップショット
 * @return {Array<{row1:number, value:*}>}
 */
function yahooApprovalLv3SnapshotCheckboxes_(masterCtx) {
  var ck = masterCtx.col[masterCtx.ckName];
  var snap = [];
  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    snap.push({ row1: r + 1, value: masterCtx.values[r][ck] });
  }
  return snap;
}

/**
 * 案A: 承認済み子SKU行のみ TRUE、他は FALSE（親行も含めOFF）
 * @param {Object} masterCtx
 * @param {Array<{rowIndex0:number, childSku:string}>} children
 */
function yahooApprovalLv3ApplyPlanACheckboxes_(masterCtx, children) {
  var target = {};
  for (var i = 0; i < children.length; i++) {
    target[children[i].rowIndex0] = true;
  }
  var ck = masterCtx.col[masterCtx.ckName];
  var col1 = ck + 1;
  var lastRow = masterCtx.values.length;
  var startRow = masterCtx.headerRowIdx + 2;
  var numRows = lastRow - masterCtx.headerRowIdx - 1;
  if (numRows <= 0) return;

  var out = [];
  for (var r = masterCtx.headerRowIdx + 1; r < lastRow; r++) {
    out.push([target[r] ? true : false]);
  }
  masterCtx.sheet.getRange(startRow, col1, numRows, 1).setValues(out);
  for (var r2 = masterCtx.headerRowIdx + 1; r2 < lastRow; r2++) {
    masterCtx.values[r2][ck] = target[r2] ? true : false;
  }
  Logger.log(
    '[yahooApprovalLv3ApplyPlanACheckboxes_] childrenOn=' + children.length + ' totalOn=' + children.length
  );
}

function yahooApprovalLv3RestoreCheckboxes_(masterCtx, snapshot) {
  if (!snapshot || !snapshot.length) return;
  var ck = masterCtx.col[masterCtx.ckName];
  var col1 = ck + 1;
  var startRow = snapshot[0].row1;
  var values = [];
  for (var i = 0; i < snapshot.length; i++) {
    values.push([snapshot[i].value]);
  }
  masterCtx.sheet.getRange(startRow, col1, values.length, 1).setValues(values);
  for (var j = 0; j < snapshot.length; j++) {
    var r0 = snapshot[j].row1 - 1;
    if (r0 >= 0 && r0 < masterCtx.values.length) {
      masterCtx.values[r0][ck] = snapshot[j].value;
    }
  }
}

function yahooApprovalLv3ApplyStock_(masterCtx, children, stockVal) {
  var iStock = masterCtx.col['在庫数'];
  if (iStock == null) {
    Logger.log('[yahooApprovalLv3ApplyStock_] 在庫数列なしのためスキップ');
    return;
  }
  var col1 = iStock + 1;
  for (var i = 0; i < children.length; i++) {
    var row1 = children[i].rowIndex0 + 1;
    masterCtx.sheet.getRange(row1, col1).setValue(stockVal);
    masterCtx.values[children[i].rowIndex0][iStock] = stockVal;
  }
  Logger.log('[yahooApprovalLv3ApplyStock_] children=' + children.length + ' stock=' + stockVal);
}

function yahooApprovalLv3SaveState_(state) {
  state.updatedAt = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(APPROVAL_YAHOO_LV3_STATE_PROP, JSON.stringify(state));
}

function yahooApprovalLv3SetTrigger_() {
  yahooApprovalLv3DeleteTriggers_();
  ScriptApp.newTrigger(APPROVAL_YAHOO_LV3_TRIGGER_FN).timeBased().after(1 * 60 * 1000).create();
}

function yahooApprovalLv3DeleteTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === APPROVAL_YAHOO_LV3_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function yahooApprovalLv3Mail_(subject, body) {
  try {
    var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    if (!email) return;
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('[yahooApprovalLv3Mail_] skip ' + ((e && e.message) || e));
  }
}
