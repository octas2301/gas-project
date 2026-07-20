/**
 * Lv2 承認①済 → 楽天日中掲載（聖域内オーケストレーション）
 * 要件: docs/org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md
 *
 * - generateRakutenCSV 本体は改変しない（呼出のみ）
 * - 既存メニュー「8. 楽天CSV出力」「一括出品」は触らない（手動逃げ道）
 * - レ点は案A: 実行直前に承認済み親＋紐づく子をON（他OFF）、終了後に復元
 *   （親だけONだと generateRakutenCSV がシングルSKU扱い→バリエーション親で32バイト超過になり得る）
 *
 * Script Properties:
 *   APPROVAL_RAKUTEN_LV2_ENABLED … 既定 false
 *   APPROVAL_RAKUTEN_LV2_APPLY_STOCK … 既定 true（対象親の在庫数を 0/1 に合わせる）
 *   APPROVAL_RAKUTEN_LV2_SKIP_CSV … 既定 false（true なら CSV 呼出をスキップ＝ドライラン）
 *   APPROVAL_RAKUTEN_LV2_STATE … レジューム用（自動）
 */

var APPROVAL_RAKUTEN_LV2_PROP = 'APPROVAL_RAKUTEN_LV2_ENABLED';
var APPROVAL_RAKUTEN_LV2_APPLY_STOCK_PROP = 'APPROVAL_RAKUTEN_LV2_APPLY_STOCK';
var APPROVAL_RAKUTEN_LV2_SKIP_CSV_PROP = 'APPROVAL_RAKUTEN_LV2_SKIP_CSV';
var APPROVAL_RAKUTEN_LV2_STATE_PROP = 'APPROVAL_RAKUTEN_LV2_STATE';
var APPROVAL_RAKUTEN_LV2_TIME_MS = 25 * 60 * 1000;
var APPROVAL_RAKUTEN_LV2_UNIQUE_IMAGE_MAX = 50;
var APPROVAL_RAKUTEN_LV2_MAX_TRIGGER_RUNS = 40;
var APPROVAL_RAKUTEN_LV2_TRIGGER_FN = 'runApprovalRakutenLv2FromTrigger';

/**
 * メニュー 19-①: 最新 APPROVED バッチの楽天行をサブバッチ分割して generateRakutenCSV を呼ぶ。
 */
function menuApprovalRakutenLv2Run() {
  var fn = 'menuApprovalRakutenLv2Run';
  if (!getBoolScriptProperty_(APPROVAL_RAKUTEN_LV2_PROP, false)) {
    var off = 'Lv2楽天は無効です。Script Properties の ' + APPROVAL_RAKUTEN_LV2_PROP + ' を true にしてください。';
    Logger.log('[' + fn + '] state=FAILED ' + off);
    try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    return;
  }
  try {
    var ui = SpreadsheetApp.getUi();
    var skipCsv = getBoolScriptProperty_(APPROVAL_RAKUTEN_LV2_SKIP_CSV_PROP, false);
    var res = ui.alert(
      'Lv2 楽天日中掲載',
      '最新の承認①（APPROVED）の楽天親を対象に、既存 generateRakutenCSV を呼び出します。\n' +
        '・手動メニュー「楽天CSV出力」はそのまま残っています\n' +
        '・レ点は一時的に「対象親＋その子」のみONにし、終了後に復元します（案A・通常出品と同じ親子レ点）\n' +
        (skipCsv ? '・SKIP_CSV=true のため CSV は呼びません（ドライラン）\n' : '・MAKE/FTP まで既存経路どおり進む場合があります\n') +
        '\n実行しますか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (res !== ui.Button.OK) return;
  } catch (eUi) {}

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var runId = 'LV2_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' +
    ('000000' + Math.floor(Math.random() * 1e6)).slice(-6);
  Logger.log('[' + fn + '] state=RUNNING runId=' + runId);
  try {
    var summary = rakutenApprovalLv2Run_(ss, runId, 0, null);
    Logger.log('[' + fn + '] state=DONE runId=' + runId + ' ' + JSON.stringify(summary));
    try {
      SpreadsheetApp.getUi().alert(
        'Lv2 実行結果',
        'runId=' + runId +
          '\nbatchId=' + (summary.batchId || '') +
          '\nサブバッチ完了=' + summary.subBatchesDone +
          '\n親SKU成功=' + summary.parentsDone +
          '\nスキップ=' + summary.skipped +
          '\n継続待ち=' + (summary.willResume ? 'YES（約1分後に自動再開）' : 'NO') +
          '\nメッセージ=' + (summary.message || ''),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e1) {}
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED runId=' + runId + ' ' + ((err && err.message) || err));
    rakutenApprovalLv2Mail_('【Lv2楽天】実行失敗', 'runId=' + runId + '\n' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert('Lv2失敗: ' + ((err && err.message) || err)); } catch (e2) {}
  }
}

/** メニュー 19-②: レジューム状態クリア */
function menuApprovalRakutenLv2ClearState() {
  try {
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_RAKUTEN_LV2_STATE_PROP);
    rakutenApprovalLv2DeleteTriggers_();
    SpreadsheetApp.getUi().alert('Lv2 実行状態をクリアしました。');
  } catch (e) {
    try { SpreadsheetApp.getUi().alert(String(e && e.message || e)); } catch (e2) {}
  }
}

/** トリガー再開 */
function runApprovalRakutenLv2FromTrigger() {
  var fn = APPROVAL_RAKUTEN_LV2_TRIGGER_FN;
  var stateJson = PropertiesService.getScriptProperties().getProperty(APPROVAL_RAKUTEN_LV2_STATE_PROP);
  if (!stateJson) return;
  rakutenApprovalLv2DeleteTriggers_();
  var state;
  try {
    state = JSON.parse(stateJson);
  } catch (e) {
    return;
  }
  var runCount = Number(state.triggerRunCount || 0) + 1;
  state.triggerRunCount = runCount;
  PropertiesService.getScriptProperties().setProperty(APPROVAL_RAKUTEN_LV2_STATE_PROP, JSON.stringify(state));
  if (runCount > APPROVAL_RAKUTEN_LV2_MAX_TRIGGER_RUNS) {
    Logger.log('[' + fn + '] state=FAILED maxTriggerRuns runId=' + state.runId);
    rakutenApprovalLv2Mail_('【Lv2楽天】自動再開上限', 'runId=' + state.runId + ' batchId=' + state.batchId);
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_RAKUTEN_LV2_STATE_PROP);
    return;
  }
  if (!getBoolScriptProperty_(APPROVAL_RAKUTEN_LV2_PROP, false)) {
    Logger.log('[' + fn + '] state=FAILED disabled');
    return;
  }
  var ss = SpreadsheetApp.openById(state.spreadsheetId);
  Logger.log('[' + fn + '] state=RUNNING runId=' + state.runId + ' nextSub=' + state.nextSubBatchIndex);
  try {
    rakutenApprovalLv2Run_(ss, state.runId, state.nextSubBatchIndex || 0, state);
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    rakutenApprovalLv2Mail_('【Lv2楽天】再開失敗', 'runId=' + state.runId + '\n' + ((err && err.message) || err));
  }
}

/**
 * @param {Spreadsheet} ss
 * @param {string} runId
 * @param {number} startSubBatchIndex
 * @param {Object|null} resumeState
 * @return {Object}
 */
function rakutenApprovalLv2Run_(ss, runId, startSubBatchIndex, resumeState) {
  var fn = 'rakutenApprovalLv2Run_';
  var startedAt = Date.now();
  var loaded = (typeof approvalQueueGetLatestApprovedRakuten_ === 'function')
    ? approvalQueueGetLatestApprovedRakuten_()
    : { found: false };
  if (!loaded.found || !loaded.lines || !loaded.lines.length) {
    throw new Error('APPROVED の楽天明細がありません。先に承認キューで承認①してください。');
  }
  var batch = loaded.batch;
  var batchId = batch.batchId;
  var inventoryMode = String(batch.inventoryMode || 'ZERO').toUpperCase() === 'ONE' ? 'ONE' : 'ZERO';
  var stockVal = inventoryMode === 'ONE' ? 1 : 0;

  var masterCtx = rakutenApprovalLv2LoadMasterContext_(ss);
  var resolved = rakutenApprovalLv2ResolveParents_(masterCtx, loaded.lines, resumeState && resumeState.doneParents);
  Logger.log('[' + fn + '] runId=' + runId + ' batchId=' + batchId +
    ' candidates=' + resolved.parents.length + ' skipped=' + resolved.skipped.length +
    ' inventoryMode=' + inventoryMode);

  var subBatches = rakutenApprovalLv2BuildSubBatches_(masterCtx, resolved.parents);
  if (!subBatches.length) {
    return {
      batchId: batchId,
      subBatchesDone: 0,
      parentsDone: 0,
      skipped: resolved.skipped.length,
      willResume: false,
      message: '実行対象の親がありません（スキップのみ）'
    };
  }

  var doneParents = (resumeState && resumeState.doneParents) ? resumeState.doneParents.slice() : [];
  var subBatchesDone = Number(resumeState && resumeState.subBatchesDone) || 0;
  var checkboxSnapshot = null;
  var willResume = false;
  var message = '';

  try {
    checkboxSnapshot = rakutenApprovalLv2SnapshotCheckboxes_(masterCtx);
    for (var s = startSubBatchIndex; s < subBatches.length; s++) {
      if ((Date.now() - startedAt) > APPROVAL_RAKUTEN_LV2_TIME_MS) {
        rakutenApprovalLv2SaveState_({
          spreadsheetId: ss.getId(),
          runId: runId,
          batchId: batchId,
          inventoryMode: inventoryMode,
          nextSubBatchIndex: s,
          subBatchesDone: subBatchesDone,
          doneParents: doneParents,
          triggerRunCount: resumeState ? resumeState.triggerRunCount : 0
        });
        rakutenApprovalLv2SetTrigger_();
        willResume = true;
        message = '時間予算のため中断。サブバッチ ' + (s + 1) + '/' + subBatches.length + ' から再開予定';
        Logger.log('[' + fn + '] state=RETRYING runId=' + runId + ' ' + message);
        break;
      }

      var sub = subBatches[s];
      var subBatchId = batchId + '_R' + (s + 1);
      Logger.log('[' + fn + '] state=RUNNING runId=' + runId + ' batchId=' + batchId +
        ' subBatchId=' + subBatchId + ' parents=' + sub.parents.length +
        ' uniqueImages=' + sub.uniqueImageCount);

      try {
        rakutenApprovalLv2ApplyPlanACheckboxes_(masterCtx, checkboxSnapshot, sub.parents);
        if (getBoolScriptProperty_(APPROVAL_RAKUTEN_LV2_APPLY_STOCK_PROP, true)) {
          rakutenApprovalLv2ApplyStock_(masterCtx, sub.parents, stockVal);
        }
        SpreadsheetApp.flush();

        if (!getBoolScriptProperty_(APPROVAL_RAKUTEN_LV2_SKIP_CSV_PROP, false)) {
          if (typeof generateRakutenCSV !== 'function') {
            throw new Error('generateRakutenCSV が見つかりません');
          }
          generateRakutenCSV(true, ss);
        } else {
          Logger.log('[' + fn + '] SKIP_CSV=true subBatchId=' + subBatchId);
        }

        for (var p = 0; p < sub.parents.length; p++) {
          doneParents.push(sub.parents[p].parentSku);
        }
        subBatchesDone++;
        Logger.log('[' + fn + '] state=DONE subBatchId=' + subBatchId + ' parents=' + sub.parents.length);
      } catch (subErr) {
        Logger.log('[' + fn + '] state=FAILED subBatchId=' + subBatchId + ' ' + ((subErr && subErr.message) || subErr));
        rakutenApprovalLv2Mail_(
          '【Lv2楽天】サブバッチ失敗',
          'runId=' + runId + '\nbatchId=' + batchId + '\nsubBatchId=' + subBatchId + '\n' + ((subErr && subErr.message) || subErr)
        );
        throw subErr;
      } finally {
        try {
          rakutenApprovalLv2RestoreCheckboxes_(masterCtx, checkboxSnapshot);
          SpreadsheetApp.flush();
        } catch (restErr) {
          Logger.log('[' + fn + '] state=FAILED checkboxRestore ' + ((restErr && restErr.message) || restErr));
          rakutenApprovalLv2Mail_(
            '【Lv2楽天】レ点復元失敗',
            'runId=' + runId + '\nbatchId=' + batchId + '\n' + ((restErr && restErr.message) || restErr) +
              '\nマスタの出品CKを目視確認してください。'
          );
          throw restErr;
        }
      }
    }

    if (!willResume) {
      PropertiesService.getScriptProperties().deleteProperty(APPROVAL_RAKUTEN_LV2_STATE_PROP);
      rakutenApprovalLv2DeleteTriggers_();
      message = message || ('全サブバッチ完了 ' + subBatchesDone + '/' + subBatches.length);
    }
  } catch (outer) {
    if (checkboxSnapshot) {
      try {
        rakutenApprovalLv2RestoreCheckboxes_(masterCtx, checkboxSnapshot);
        SpreadsheetApp.flush();
      } catch (e3) {
        rakutenApprovalLv2Mail_('【Lv2楽天】レ点復元失敗', 'runId=' + runId + '\n' + ((e3 && e3.message) || e3));
      }
    }
    throw outer;
  }

  return {
    batchId: batchId,
    subBatchesDone: subBatchesDone,
    parentsDone: doneParents.length,
    skipped: resolved.skipped.length,
    willResume: willResume,
    message: message
  };
}

/** @return {{sheet:Sheet, values:Array, headerRowIdx:number, col:Object, ckName:string}} */
function rakutenApprovalLv2LoadMasterContext_(ss) {
  var masterName = (typeof MASTER_SHEET_NAME !== 'undefined') ? MASTER_SHEET_NAME : '▼商品マスタ(人間作業用)';
  var ckName = (typeof CHECKBOX_HEADER_NAME !== 'undefined') ? CHECKBOX_HEADER_NAME : '出品CK';
  var sheet = ss.getSheetByName(masterName);
  if (!sheet) throw new Error('マスタシートが見つかりません: ' + masterName);
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
  if (col['親SKU'] == null || col[ckName] == null) {
    throw new Error('必須列がありません（親SKU / 出品CK）');
  }
  if (col['子SKU'] == null) {
    Logger.log('[rakutenApprovalLv2LoadMasterContext_] 警告: 子SKU列なし → 案Aは親のみレ点になる');
  }
  return { sheet: sheet, values: values, headerRowIdx: headerRowIdx, col: col, ckName: ckName };
}

/**
 * @param {Object} masterCtx
 * @param {Array} lines
 * @param {Array<string>|null} doneParents
 * @return {{parents:Array, skipped:Array}}
 */
function rakutenApprovalLv2ResolveParents_(masterCtx, lines, doneParents) {
  var doneMap = {};
  if (doneParents) {
    for (var d = 0; d < doneParents.length; d++) doneMap[String(doneParents[d])] = true;
  }
  var col = masterCtx.col;
  var values = masterCtx.values;
  var iParent = col['親SKU'];
  var iChild = col['子SKU'];
  var iStock = col['在庫数'];
  var parents = [];
  var skipped = [];
  var seen = {};

  for (var i = 0; i < lines.length; i++) {
    var L = lines[i];
    if (String(L.mall) !== 'rakuten' || String(L.lineStatus) !== 'APPROVED') continue;
    var parentSku = String(L.parentSku || '').trim();
    if (!parentSku) {
      skipped.push({ reason: 'SKIPPED_ORPHAN', parentSku: '', detail: 'parentSku空' });
      continue;
    }
    if (seen[parentSku] || doneMap[parentSku]) continue;

    var rowIdx = -1;
    if (L.masterRow != null && Number(L.masterRow) >= 1) {
      var cand = Number(L.masterRow) - 1;
      if (cand > masterCtx.headerRowIdx && cand < values.length) {
        var childAt = iChild != null ? String(values[cand][iChild] || '').trim() : '';
        var parentAt = String(values[cand][iParent] || '').trim();
        if (parentAt === parentSku && !childAt) rowIdx = cand;
      }
    }
    if (rowIdx < 0) {
      for (var r = masterCtx.headerRowIdx + 1; r < values.length; r++) {
        var child = iChild != null ? String(values[r][iChild] || '').trim() : '';
        if (String(values[r][iParent] || '').trim() === parentSku && !child) {
          rowIdx = r;
          break;
        }
      }
    }
    if (rowIdx < 0) {
      skipped.push({ reason: 'SKIPPED_ORPHAN', parentSku: parentSku, detail: 'マスタに親行なし' });
      continue;
    }

    var stockRaw = iStock != null ? values[rowIdx][iStock] : '';
    var stockNum = (stockRaw === '' || stockRaw == null) ? null : Number(stockRaw);
    if (stockNum != null && !isNaN(stockNum) && stockNum > 0) {
      skipped.push({ reason: 'SKIPPED_IN_STOCK', parentSku: parentSku, detail: '在庫数=' + stockNum });
      continue;
    }

    seen[parentSku] = true;
    parents.push({
      parentSku: parentSku,
      rowIndex0: rowIdx,
      masterRow: rowIdx + 1,
      lineId: L.lineId
    });
  }
  return { parents: parents, skipped: skipped };
}

/**
 * @param {Object} masterCtx
 * @param {Array} parents
 * @return {Array<{parents:Array, uniqueImageCount:number}>}
 */
function rakutenApprovalLv2BuildSubBatches_(masterCtx, parents) {
  var batches = [];
  var current = [];
  var keySet = {};

  function flush() {
    if (!current.length) return;
    batches.push({
      parents: current.slice(),
      uniqueImageCount: Object.keys(keySet).length
    });
    current = [];
    keySet = {};
  }

  for (var i = 0; i < parents.length; i++) {
    var p = parents[i];
    var keys = rakutenApprovalLv2CollectImageKeysForParent_(masterCtx, p.parentSku);
    var trial = {};
    for (var k in keySet) {
      if (Object.prototype.hasOwnProperty.call(keySet, k)) trial[k] = true;
    }
    for (var j = 0; j < keys.length; j++) trial[keys[j]] = true;
    var trialCount = Object.keys(trial).length;
    // 追加するとユニーク50超になるなら、先に現バッチを確定（1親単独で50超はそのまま1バッチ）
    if (current.length && trialCount > APPROVAL_RAKUTEN_LV2_UNIQUE_IMAGE_MAX) {
      flush();
      trial = {};
      for (var j2 = 0; j2 < keys.length; j2++) trial[keys[j2]] = true;
    }
    current.push(p);
    keySet = trial;
  }
  flush();
  return batches;
}

/** @return {Array<string>} */
function rakutenApprovalLv2CollectImageKeysForParent_(masterCtx, parentSku) {
  var col = masterCtx.col;
  var values = masterCtx.values;
  var iParent = col['親SKU'];
  var keys = {};
  var imgNames = [];
  for (var m = 1; m <= 10; m++) {
    imgNames.push('楽天メイン画像' + m);
    imgNames.push('楽天サブ画像' + m);
  }
  for (var r = masterCtx.headerRowIdx + 1; r < values.length; r++) {
    if (String(values[r][iParent] || '').trim() !== parentSku) continue;
    for (var n = 0; n < imgNames.length; n++) {
      var idx = col[imgNames[n]];
      if (idx == null) continue;
      var key = rakutenApprovalLv2ImageKey_(values[r][idx]);
      if (key) keys[key] = true;
    }
  }
  return Object.keys(keys);
}

/** DriveファイルID優先、なければ正規化URL */
function rakutenApprovalLv2ImageKey_(cell) {
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
function rakutenApprovalLv2SnapshotCheckboxes_(masterCtx) {
  var ck = masterCtx.col[masterCtx.ckName];
  var snap = [];
  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    snap.push({ row1: r + 1, value: masterCtx.values[r][ck] });
  }
  return snap;
}

/**
 * 案A: 対象親＋同一親SKUの子行を TRUE、他は FALSE（実行用）。
 * 通常の手動出品（親子レ点）と同じ形にし、generateRakutenCSV をマルチSKU経路へ乗せる。
 * @param {Object} masterCtx
 * @param {Array} snapshot 未使用（署名互換）。復元は別関数。
 * @param {Array<{rowIndex0:number, parentSku:string}>} parents
 */
function rakutenApprovalLv2ApplyPlanACheckboxes_(masterCtx, snapshot, parents) {
  var target = {};
  var parentSkuSet = {};
  for (var i = 0; i < parents.length; i++) {
    target[parents[i].rowIndex0] = true;
    var ps = String(parents[i].parentSku || '').trim();
    if (ps) parentSkuSet[ps] = true;
  }
  var iParent = masterCtx.col['親SKU'];
  var iChild = masterCtx.col['子SKU'];
  var ck = masterCtx.col[masterCtx.ckName];
  var col1 = ck + 1;
  var lastRow = masterCtx.values.length;
  var startRow = masterCtx.headerRowIdx + 2; // 1-based first data
  var numRows = lastRow - masterCtx.headerRowIdx - 1;
  if (numRows <= 0) return;

  var childOn = 0;
  if (iParent != null) {
    for (var rScan = masterCtx.headerRowIdx + 1; rScan < lastRow; rScan++) {
      if (target[rScan]) continue;
      var pAt = String(masterCtx.values[rScan][iParent] || '').trim();
      if (!pAt || !parentSkuSet[pAt]) continue;
      var cAt = (iChild != null) ? String(masterCtx.values[rScan][iChild] || '').trim() : '';
      if (!cAt) continue; // 親行は親SKU一致でも子SKU空 → 既に target 済み想定
      target[rScan] = true;
      childOn++;
    }
  }

  var out = [];
  for (var r = masterCtx.headerRowIdx + 1; r < lastRow; r++) {
    out.push([target[r] ? true : false]);
  }
  masterCtx.sheet.getRange(startRow, col1, numRows, 1).setValues(out);
  for (var r2 = masterCtx.headerRowIdx + 1; r2 < lastRow; r2++) {
    masterCtx.values[r2][ck] = target[r2] ? true : false;
  }
  Logger.log(
    '[rakutenApprovalLv2ApplyPlanACheckboxes_] parents=' + parents.length +
      ' childrenOn=' + childOn + ' totalOn=' + (parents.length + childOn)
  );
}

function rakutenApprovalLv2RestoreCheckboxes_(masterCtx, snapshot) {
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

function rakutenApprovalLv2ApplyStock_(masterCtx, parents, stockVal) {
  var iStock = masterCtx.col['在庫数'];
  if (iStock == null) {
    Logger.log('[rakutenApprovalLv2ApplyStock_] 在庫数列なしのためスキップ');
    return;
  }
  var col1 = iStock + 1;
  for (var i = 0; i < parents.length; i++) {
    var row1 = parents[i].rowIndex0 + 1;
    masterCtx.sheet.getRange(row1, col1).setValue(stockVal);
    masterCtx.values[parents[i].rowIndex0][iStock] = stockVal;
  }
  Logger.log('[rakutenApprovalLv2ApplyStock_] parents=' + parents.length + ' stock=' + stockVal);
}

function rakutenApprovalLv2SaveState_(state) {
  state.updatedAt = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(APPROVAL_RAKUTEN_LV2_STATE_PROP, JSON.stringify(state));
}

function rakutenApprovalLv2SetTrigger_() {
  rakutenApprovalLv2DeleteTriggers_();
  ScriptApp.newTrigger(APPROVAL_RAKUTEN_LV2_TRIGGER_FN).timeBased().after(1 * 60 * 1000).create();
}

function rakutenApprovalLv2DeleteTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === APPROVAL_RAKUTEN_LV2_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function rakutenApprovalLv2Mail_(subject, body) {
  try {
    var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    if (!email) return;
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('[rakutenApprovalLv2Mail_] skip ' + ((e && e.message) || e));
  }
}
