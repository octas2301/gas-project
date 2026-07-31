/**
 * U2: Amazon MAIN／PT — マッチング sheet 作業面 + マスタ永続 + Drive 02 コピー出口
 * 要件: docs/org/D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md
 *
 * - 楽天 1〜25 列アップロード経路は触らない（Amazon 枠は col 76+）
 * - generateRakutenCSV / Yahoo.js 非改変
 * - 02 出力はコピー（楽天 setName/archive と競合しない）
 *
 * Script Properties:
 *   AMAZON_IMAGE_U2_ENABLED … 既定 false
 *   AMAZON_IMAGE_CANDIDATE_FOLDER_ID … 白抜き候補（楽天ソースと分離）
 *   AMAZON_DRIVE_IMAGE_FOLDER_ID … 02 出口（空なら既定 ID）
 *   AMAZON_IMAGE_CANDIDATE_ARCHIVE_ENABLED … ④成功分を 07/アップロード済み画像 へ退避（未設定=true／false でOFF）
 */

var AMAZON_IMAGE_U2_PROP = 'AMAZON_IMAGE_U2_ENABLED';
var AMAZON_IMAGE_CANDIDATE_FOLDER_PROP = 'AMAZON_IMAGE_CANDIDATE_FOLDER_ID';
var AMAZON_IMAGE_DRIVE02_PROP = 'AMAZON_DRIVE_IMAGE_FOLDER_ID';
var AMAZON_IMAGE_DRIVE02_DEFAULT = '1T6_E6T-qd9whSF8Re8lyRVB2n-P4BM84';
var AMAZON_IMAGE_CANDIDATE_ARCHIVE_PROP = 'AMAZON_IMAGE_CANDIDATE_ARCHIVE_ENABLED';
var AMAZON_IMAGE_CANDIDATE_ARCHIVE_FOLDER_NAME = 'アップロード済み画像';

/** 1-based。楽天候補は 26〜75 → Amazon は 76 以降 */
var AMAZON_MX_COL_MAIN = 76;
var AMAZON_MX_COL_MODE = 77;
var AMAZON_MX_COL_PT1 = 78;
var AMAZON_MX_PT_COUNT = 8;
var AMAZON_MX_COL_CAND1 = 86;
var AMAZON_MX_CAND_COUNT = 20;

var AMAZON_MASTER_COL_MODE = 'Amazon画像モード';
var AMAZON_MASTER_COL_MAIN = 'Amazon MAIN 参照';
var AMAZON_MASTER_COL_PT = 'Amazon PT 参照';

/**
 * generateAiImageMatrix 完了後フック（sheet.clear 後の枠＋マスタ復元）
 */
function amazonImageMatrixOnAfterGenerate_(ss, matrixSheet) {
  var fn = 'amazonImageMatrixOnAfterGenerate_';
  if (!getBoolScriptProperty_(AMAZON_IMAGE_U2_PROP, false)) {
    Logger.log('[' + fn + '] state=SKIPPED U2 disabled');
    return;
  }
  Logger.log('[' + fn + '] state=RUNNING');
  try {
    amazonImageMatrixEnsureZone_(matrixSheet);
    var n = amazonImageMatrixRestoreFromMaster_(ss, matrixSheet, { silent: true });
    Logger.log('[' + fn + '] state=DONE restoredRows=' + n);
  } catch (e) {
    Logger.log('[' + fn + '] state=FAILED ' + ((e && e.message) || e));
  }
}

/** メニュー: 枠追加＆マスタから復元 */
function menuAmazonImageMatrixRestore() {
  var fn = 'menuAmazonImageMatrixRestore';
  if (!amazonImageU2Guard_(fn)) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_MATRIX);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「' + SHEET_NAME_MATRIX + '」がありません。先に C でマトリクスを作成してください。');
    return;
  }
  Logger.log('[' + fn + '] state=RUNNING');
  amazonImageMatrixEnsureZone_(sheet);
  var n = amazonImageMatrixRestoreFromMaster_(ss, sheet, { silent: false });
  Logger.log('[' + fn + '] state=DONE restoredRows=' + n);
  SpreadsheetApp.getUi().alert('Amazon枠を整え、マスタから復元しました（対象行≈' + n + '）。\n列' + AMAZON_MX_COL_MAIN + '以降が Amazon 用です（楽天1〜25列は非改変）。');
}

/** メニュー: sheet → マスタ永続化 */
function menuAmazonImageMatrixPersist() {
  var fn = 'menuAmazonImageMatrixPersist';
  if (!amazonImageU2Guard_(fn)) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_MATRIX);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('マトリクスシートがありません。');
    return;
  }
  Logger.log('[' + fn + '] state=RUNNING');
  var summary = amazonImageMatrixPersistToMaster_(ss, sheet);
  Logger.log('[' + fn + '] state=DONE ' + JSON.stringify(summary));
  SpreadsheetApp.getUi().alert(
    'マスタへ保存しました。\n更新行=' + summary.updated +
      '\nスキップ=' + summary.skipped +
      '\n列追加=' + (summary.colsAdded || []).join(',')
  );
}

/** メニュー: Drive 02 へコピー出力 */
function menuAmazonImageMatrixExportTo02() {
  var fn = 'menuAmazonImageMatrixExportTo02';
  if (!amazonImageU2Guard_(fn)) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_MATRIX);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('マトリクスシートがありません。');
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var conf = ui.alert(
    'Amazon → Drive 02',
    'マトリクス／マスタの Amazon MAIN（ONLY時は PT）を Drive 02 へコピーします。\n' +
      '同名ファイルは置き換えます。楽天アップは実行しません。\n実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (conf !== ui.Button.OK) return;
  Logger.log('[' + fn + '] state=RUNNING');
  var summary = amazonImageMatrixExportTo02_(ss, sheet);
  Logger.log('[' + fn + '] state=DONE ' + JSON.stringify(summary));
  ui.alert(
    'Drive 02 出力完了\nMAIN成功=' + summary.mainOk +
      '\nPT成功=' + summary.ptOk +
      '\n失敗=' + summary.failed +
      '\n候補退避=' + (summary.archived != null ? summary.archived : 0) +
      '\n' + (summary.message || '')
  );
}

/** メニュー: Amazon 候補フォルダの画像を右端に並べる */
function menuAmazonImageMatrixLoadCandidates() {
  var fn = 'menuAmazonImageMatrixLoadCandidates';
  if (!amazonImageU2Guard_(fn)) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_MATRIX);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('マトリクスシートがありません。');
    return;
  }
  var selection = amazonImageMatrixValidateCurrentSelection_(ss, sheet);
  if (!selection.ok) {
    Logger.log('[' + fn + '] state=FAILED reason=stale_matrix ' + selection.message);
    SpreadsheetApp.getUi().alert(
      '安全停止: マッチングsheetが現在のレ点対象と一致しません。\n' +
        selection.message + '\n\n先に C を再実行してください。'
    );
    return;
  }
  var folderId = String(PropertiesService.getScriptProperties().getProperty(AMAZON_IMAGE_CANDIDATE_FOLDER_PROP) || '').trim();
  if (!folderId) {
    SpreadsheetApp.getUi().alert(AMAZON_IMAGE_CANDIDATE_FOLDER_PROP + ' が未設定です。Amazon用白抜き候補フォルダの ID を設定してください。');
    return;
  }
  Logger.log('[' + fn + '] state=RUNNING folderId=' + folderId);
  amazonImageMatrixEnsureZone_(sheet);
  var n = amazonImageMatrixLoadCandidates_(sheet, folderId);
  Logger.log('[' + fn + '] state=DONE placed=' + n);
  SpreadsheetApp.getUi().alert('Amazon候補を列' + AMAZON_MX_COL_CAND1 + '以降に並べました（' + n + '枚）。子SKU行の Amazon MAIN／PT へドラッグしてください。');
}

function amazonImageU2Guard_(fn) {
  if (!getBoolScriptProperty_(AMAZON_IMAGE_U2_PROP, false)) {
    var msg = 'U2 Amazon画像は無効です。Script Properties の ' + AMAZON_IMAGE_U2_PROP + ' を true にしてください。';
    Logger.log('[' + fn + '] state=FAILED ' + msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e0) {}
    return false;
  }
  return true;
}

/**
 * C-Amazon②の誤紐付け防止。現在のレ点選定とマッチングsheetの親子SKUを照合する。
 * C本体と同じ契約: 子レ点あり=親+レ点子、子レ点なしで親レ点=親+全子。
 * @return {{ok:boolean,message:string}}
 */
function amazonImageMatrixValidateCurrentSelection_(ss, matrixSheet) {
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) return { ok: false, message: 'マスタがありません。' };

  var values = masterSheet.getDataRange().getValues();
  var headerRowIdx = getAnchorRowIndex(values);
  if (headerRowIdx < 0) return { ok: false, message: 'マスタヘッダーが見つかりません。' };

  var colMap = getColumnIndexMap(values[headerRowIdx]);
  var idxP = colMap[COL_NAME_PARENT_SKU];
  var idxC = colMap[COL_NAME_CHILD_SKU];
  var idxCheck = colMap[COL_NAME_CHECK];
  if (idxP === undefined || idxCheck === undefined) {
    return { ok: false, message: '必須列（親SKU または 出品CK）が見つかりません。' };
  }

  function isChecked_(v) {
    return v === true || v === 1 || String(v).toUpperCase() === 'TRUE';
  }
  function key_(p, c) {
    return String(p).trim() + '\t' + String(c || '').trim();
  }

  var groups = {};
  for (var r = headerRowIdx + 1; r < values.length; r++) {
    var row = values[r];
    var pCode = String(row[idxP] || '').trim();
    var cCode = idxC !== undefined ? String(row[idxC] || '').trim() : '';
    if (!pCode) continue;
    if (!groups[pCode]) groups[pCode] = { parentChecked: false, children: [], checkedChildren: [] };
    if (!cCode) {
      groups[pCode].parentChecked = isChecked_(row[idxCheck]);
    } else {
      groups[pCode].children.push(cCode);
      if (isChecked_(row[idxCheck])) groups[pCode].checkedChildren.push(cCode);
    }
  }

  var expected = {};
  Object.keys(groups).forEach(function (pCode) {
    var group = groups[pCode];
    var children = group.checkedChildren.length > 0
      ? group.checkedChildren
      : (group.parentChecked ? group.children : []);
    if (children.length === 0 && !group.parentChecked) return;
    expected[key_(pCode, '')] = true;
    children.forEach(function (cCode) { expected[key_(pCode, cCode)] = true; });
  });

  var actual = {};
  var lastRow = matrixSheet.getLastRow();
  if (lastRow >= 3) {
    var sheetRows = matrixSheet.getRange(3, 1, lastRow - 2, 2).getValues();
    sheetRows.forEach(function (row) {
      var pCode = String(row[0] || '').trim();
      if (pCode) actual[key_(pCode, row[1])] = true;
    });
  }

  var expectedKeys = Object.keys(expected).sort();
  var actualKeys = Object.keys(actual).sort();
  if (expectedKeys.join('\n') === actualKeys.join('\n')) {
    return { ok: true, message: '一致' };
  }

  var missing = expectedKeys.filter(function (k) { return !actual[k]; });
  var stale = actualKeys.filter(function (k) { return !expected[k]; });
  function display_(keys) {
    return keys.slice(0, 3).map(function (k) { return k.replace('\t', ' / '); }).join(', ') || 'なし';
  }
  return {
    ok: false,
    message: '不足=' + display_(missing) + '\n旧・余分=' + display_(stale)
  };
}

function amazonImageMatrixEnsureZone_(sheet) {
  var needCols = AMAZON_MX_COL_CAND1 + AMAZON_MX_CAND_COUNT - 1;
  if (sheet.getMaxColumns() < needCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols - sheet.getMaxColumns());
  }
  sheet.getRange(1, AMAZON_MX_COL_MAIN).setValue('▼Amazon MAIN/PT（U2・楽天青枠外）').setFontWeight('bold').setFontColor('white').setBackground('#b45f06');
  sheet.getRange(1, AMAZON_MX_COL_MODE).setValue('モード').setBackground('#b45f06').setFontColor('white');
  var i;
  for (i = 0; i < AMAZON_MX_PT_COUNT; i++) {
    sheet.getRange(1, AMAZON_MX_COL_PT1 + i).setBackground('#b45f06');
  }
  sheet.getRange(1, AMAZON_MX_COL_CAND1).setValue('▼Amazon候補（右からMAIN/PTへ）').setBackground('#7f6000').setFontColor('white');

  sheet.getRange(2, AMAZON_MX_COL_MAIN).setValue('Amazon MAIN').setBackground('#fce5cd').setFontWeight('bold');
  sheet.getRange(2, AMAZON_MX_COL_MODE).setValue('Amazonモード').setBackground('#fce5cd').setFontWeight('bold');
  for (i = 0; i < AMAZON_MX_PT_COUNT; i++) {
    sheet.getRange(2, AMAZON_MX_COL_PT1 + i).setValue('Amazon PT' + ('0' + (i + 1)).slice(-2)).setBackground('#fce5cd').setFontWeight('bold');
  }
  for (i = 0; i < AMAZON_MX_CAND_COUNT; i++) {
    sheet.getRange(2, AMAZON_MX_COL_CAND1 + i).setValue('Amz候補' + (i + 1)).setBackground('#fff2cc');
  }
  sheet.setColumnWidths(AMAZON_MX_COL_MAIN, 1 + AMAZON_MX_PT_COUNT, 100);
  sheet.setColumnWidth(AMAZON_MX_COL_MODE, 90);
  sheet.setColumnWidths(AMAZON_MX_COL_CAND1, AMAZON_MX_CAND_COUNT, 60);
}

/**
 * @return {number} 復元した行数
 */
function amazonImageMatrixRestoreFromMaster_(ss, matrixSheet, opts) {
  opts = opts || {};
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) throw new Error('マスタがありません');
  var mValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = getAnchorRowIndex(mValues);
  if (headerRowIdx < 0) throw new Error('マスタヘッダーが見つかりません');
  var colMap = getColumnIndexMap(mValues[headerRowIdx]);
  var idxP = colMap['親SKU'];
  var idxC = colMap['子SKU'];
  var idxMode = colMap[AMAZON_MASTER_COL_MODE];
  var idxMain = colMap[AMAZON_MASTER_COL_MAIN];
  var idxPt = colMap[AMAZON_MASTER_COL_PT];
  if (idxMain === undefined) {
    if (!opts.silent) {
      try {
        SpreadsheetApp.getUi().alert('マスタに「' + AMAZON_MASTER_COL_MAIN + '」列がありません。先に「sheet→マスタ保存」で列追加するか、手で追加してください。');
      } catch (eA) {}
    }
    return 0;
  }

  var byKey = {};
  var r;
  for (r = headerRowIdx + 1; r < mValues.length; r++) {
    var p = String(mValues[r][idxP] || '').trim();
    if (!p) continue;
    var c = idxC !== undefined ? String(mValues[r][idxC] || '').trim() : '';
    byKey[p + '\t' + c] = {
      mode: idxMode !== undefined ? String(mValues[r][idxMode] || '').trim() : '',
      mainId: amazonImageExtractDriveId_(mValues[r][idxMain]),
      ptIds: idxPt !== undefined ? amazonImageSplitPtIds_(mValues[r][idxPt]) : []
    };
  }

  var lastRow = matrixSheet.getLastRow();
  if (lastRow < 3) return 0;
  var restored = 0;
  for (r = 3; r <= lastRow; r++) {
    var pCode = String(matrixSheet.getRange(r, 1).getValue() || '').trim();
    var cCode = String(matrixSheet.getRange(r, 2).getValue() || '').trim();
    if (!pCode) continue;
    var hit = byKey[pCode + '\t' + cCode];
    if (!hit) continue;
    if (hit.mode) matrixSheet.getRange(r, AMAZON_MX_COL_MODE).setValue(hit.mode);
    if (hit.mainId) {
      matrixSheet.getRange(r, AMAZON_MX_COL_MAIN).setFormula(amazonImageFormula_(hit.mainId));
      restored++;
    }
    var pi;
    for (pi = 0; pi < AMAZON_MX_PT_COUNT; pi++) {
      if (hit.ptIds[pi]) {
        matrixSheet.getRange(r, AMAZON_MX_COL_PT1 + pi).setFormula(amazonImageFormula_(hit.ptIds[pi]));
      }
    }
  }
  return restored;
}

function amazonImageMatrixPersistToMaster_(ss, matrixSheet) {
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) throw new Error('マスタがありません');
  var mValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = getAnchorRowIndex(mValues);
  if (headerRowIdx < 0) throw new Error('マスタヘッダーが見つかりません');

  var ensure = amazonImageEnsureMasterColumns_(masterSheet, headerRowIdx);
  mValues = masterSheet.getDataRange().getValues();
  var colMap = getColumnIndexMap(mValues[headerRowIdx]);
  var idxP = colMap['親SKU'];
  var idxC = colMap['子SKU'];
  var idxMode = colMap[AMAZON_MASTER_COL_MODE];
  var idxMain = colMap[AMAZON_MASTER_COL_MAIN];
  var idxPt = colMap[AMAZON_MASTER_COL_PT];

  var rowIndexByKey = {};
  var r;
  for (r = headerRowIdx + 1; r < mValues.length; r++) {
    var p = String(mValues[r][idxP] || '').trim();
    if (!p) continue;
    var c = idxC !== undefined ? String(mValues[r][idxC] || '').trim() : '';
    rowIndexByKey[p + '\t' + c] = r;
  }

  var lastRow = matrixSheet.getLastRow();
  var num = Math.max(lastRow - 2, 0);
  if (num < 1) return { updated: 0, skipped: 0, colsAdded: ensure.added };

  var keys = matrixSheet.getRange(3, 1, num, 2).getValues();
  var mainF = matrixSheet.getRange(3, AMAZON_MX_COL_MAIN, num, 1).getFormulas();
  var modes = matrixSheet.getRange(3, AMAZON_MX_COL_MODE, num, 1).getValues();
  var ptF = matrixSheet.getRange(3, AMAZON_MX_COL_PT1, num, AMAZON_MX_PT_COUNT).getFormulas();

  var updated = 0;
  var skipped = 0;
  var i;
  for (i = 0; i < keys.length; i++) {
    var pCode = String(keys[i][0] || '').trim();
    var cCode = String(keys[i][1] || '').trim();
    if (!pCode) {
      skipped++;
      continue;
    }
    var mRow = rowIndexByKey[pCode + '\t' + cCode];
    if (mRow === undefined) {
      skipped++;
      continue;
    }
    var mainId = amazonImageExtractDriveId_(mainF[i][0]);
    var mode = String(modes[i][0] || '').trim();
    if (!mode) mode = amazonImageInferMode_(mValues[mRow], colMap);

    var ptIds = [];
    var t;
    for (t = 0; t < AMAZON_MX_PT_COUNT; t++) {
      var id = amazonImageExtractDriveId_(ptF[i][t]);
      if (id) ptIds.push(id);
    }

    if (idxMode !== undefined) masterSheet.getRange(mRow + 1, idxMode + 1).setValue(mode);
    if (idxMain !== undefined) masterSheet.getRange(mRow + 1, idxMain + 1).setValue(mainId || '');
    if (idxPt !== undefined) masterSheet.getRange(mRow + 1, idxPt + 1).setValue(ptIds.join('|'));
    updated++;
  }
  return { updated: updated, skipped: skipped, colsAdded: ensure.added };
}

function amazonImageMatrixExportTo02_(ss, matrixSheet) {
  var folderId = String(PropertiesService.getScriptProperties().getProperty(AMAZON_IMAGE_DRIVE02_PROP) || '').trim() ||
    AMAZON_IMAGE_DRIVE02_DEFAULT;
  var dest = DriveApp.getFolderById(folderId);
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  var mValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = getAnchorRowIndex(mValues);
  var colMap = getColumnIndexMap(mValues[headerRowIdx]);

  // 先に永続化してから出力（作業面と正を揃える）
  amazonImageMatrixPersistToMaster_(ss, matrixSheet);

  var lastRow = matrixSheet.getLastRow();
  var num = Math.max(lastRow - 2, 0);
  if (num < 1) return { mainOk: 0, ptOk: 0, failed: 0, message: 'データ行なし' };

  var keys = matrixSheet.getRange(3, 1, num, 2).getValues();
  var mainF = matrixSheet.getRange(3, AMAZON_MX_COL_MAIN, num, 1).getFormulas();
  var modes = matrixSheet.getRange(3, AMAZON_MX_COL_MODE, num, 1).getValues();
  var ptF = matrixSheet.getRange(3, AMAZON_MX_COL_PT1, num, AMAZON_MX_PT_COUNT).getFormulas();

  var mainOk = 0;
  var ptOk = 0;
  var failed = 0;
  var notes = [];
  var usedIds = {};
  var i;
  for (i = 0; i < keys.length; i++) {
    var pCode = String(keys[i][0] || '').trim();
    var cCode = String(keys[i][1] || '').trim();
    if (!pCode) continue;
    // 親行（子SKU空）は MAIN 出力対象外（子SKU＝sellerSku が正）
    if (!cCode) continue;

    var sellerSku = cCode;
    var mode = String(modes[i][0] || '').trim().toUpperCase();
    if (!mode) {
      var mRow = amazonImageFindMasterRow_(mValues, headerRowIdx, colMap, pCode, cCode);
      mode = mRow >= 0 ? amazonImageInferMode_(mValues[mRow], colMap) : 'AMAZON_ONLY';
    }
    if (mode.indexOf('REUSE') >= 0) mode = 'REUSE_RAKUTEN';
    else mode = 'AMAZON_ONLY';

    var mainId = amazonImageExtractDriveId_(mainF[i][0]);
    if (!mainId) {
      failed++;
      notes.push(sellerSku + ': MAINなし');
      continue;
    }
    try {
      amazonImageCopyTo02_(dest, mainId, sellerSku + '.MAIN.jpg');
      mainOk++;
      usedIds[mainId] = true;
    } catch (eM) {
      failed++;
      notes.push(sellerSku + ' MAIN: ' + ((eM && eM.message) || eM));
      continue;
    }

    if (mode === 'AMAZON_ONLY') {
      var t;
      var ptSeq = 0;
      for (t = 0; t < AMAZON_MX_PT_COUNT; t++) {
        var ptId = amazonImageExtractDriveId_(ptF[i][t]);
        if (!ptId) continue;
        ptSeq++;
        var ptName = sellerSku + '.PT' + ('0' + ptSeq).slice(-2) + '.jpg';
        try {
          amazonImageCopyTo02_(dest, ptId, ptName);
          ptOk++;
          usedIds[ptId] = true;
        } catch (eP) {
          failed++;
          notes.push(sellerSku + ' ' + ptName + ': ' + ((eP && eP.message) || eP));
        }
      }
    }
  }

  var archiveResult = amazonImageArchiveUsedCandidates_(Object.keys(usedIds));
  if (archiveResult.note) notes.push(archiveResult.note);
  Logger.log(
    '[amazonImageMatrixExportTo02_] archive moved=' +
      archiveResult.moved +
      ' skipped=' +
      archiveResult.skipped +
      ' err=' +
      archiveResult.errors
  );

  return {
    mainOk: mainOk,
    ptOk: ptOk,
    failed: failed,
    archived: archiveResult.moved,
    message: notes.slice(0, 8).join('\n') + (notes.length > 8 ? '\n…' : '')
  };
}

/**
 * ④で 02 コピー成功したファイルだけ、候補フォルダ(07)直下にあれば「アップロード済み画像」へ退避。
 * Property AMAZON_IMAGE_CANDIDATE_ARCHIVE_ENABLED=false で無効（未設定=有効）。
 */
function amazonImageArchiveUsedCandidates_(fileIds) {
  var out = { moved: 0, skipped: 0, errors: 0, note: '' };
  if (!fileIds || !fileIds.length) return out;
  if (!getBoolScriptProperty_(AMAZON_IMAGE_CANDIDATE_ARCHIVE_PROP, true)) {
    out.note = '候補退避スキップ(Property OFF)';
    return out;
  }
  var candId = String(PropertiesService.getScriptProperties().getProperty(AMAZON_IMAGE_CANDIDATE_FOLDER_PROP) || '').trim();
  if (!candId) {
    out.note = '候補退避スキップ(候補フォルダ未設定)';
    return out;
  }
  var candFolder;
  try {
    candFolder = DriveApp.getFolderById(candId);
  } catch (eF) {
    out.note = '候補退避失敗: フォルダID無効';
    out.errors++;
    return out;
  }
  var archiveFolder = amazonImageGetOrCreateArchiveFolder_(candFolder);
  if (!archiveFolder) {
    out.note = '候補退避失敗: 退避フォルダ作成不可';
    out.errors++;
    return out;
  }
  var archiveId = archiveFolder.getId();
  var i;
  for (i = 0; i < fileIds.length; i++) {
    var fid = fileIds[i];
    if (!fid) continue;
    try {
      var file = DriveApp.getFileById(fid);
      var parents = file.getParents();
      var inCandRoot = false;
      var alreadyArchived = false;
      while (parents.hasNext()) {
        var p = parents.next();
        var pid = p.getId();
        if (pid === candId) inCandRoot = true;
        if (pid === archiveId) alreadyArchived = true;
      }
      if (alreadyArchived || !inCandRoot) {
        out.skipped++;
        continue;
      }
      file.moveTo(archiveFolder);
      out.moved++;
    } catch (eA) {
      out.errors++;
      Logger.log('[amazonImageArchiveUsedCandidates_] id=' + fid + ' err=' + ((eA && eA.message) || eA));
    }
  }
  if (out.errors) out.note = '候補退避エラー=' + out.errors;
  return out;
}

function amazonImageGetOrCreateArchiveFolder_(candFolder) {
  var it = candFolder.getFoldersByName(AMAZON_IMAGE_CANDIDATE_ARCHIVE_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  try {
    return candFolder.createFolder(AMAZON_IMAGE_CANDIDATE_ARCHIVE_FOLDER_NAME);
  } catch (eC) {
    Logger.log('[amazonImageGetOrCreateArchiveFolder_] ' + ((eC && eC.message) || eC));
    return null;
  }
}

function amazonImageCopyTo02_(destFolder, fileId, destName) {
  var src = DriveApp.getFileById(fileId);
  var existing = destFolder.getFilesByName(destName);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }
  src.makeCopy(destName, destFolder);
}

function amazonImageMatrixLoadCandidates_(sheet, folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var list = [];
  while (files.hasNext() && list.length < AMAZON_MX_CAND_COUNT) {
    var f = files.next();
    if (String(f.getMimeType() || '').indexOf('image/') === 0) list.push(f);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0;
  var j;
  for (j = 0; j < list.length; j++) {
    try { list[j].setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (eS) {}
    var formula = amazonImageFormula_(list[j].getId());
    var rr;
    for (rr = 3; rr <= lastRow; rr++) {
      sheet.getRange(rr, AMAZON_MX_COL_CAND1 + j).setFormula(formula);
    }
  }
  return list.length;
}

function amazonImageEnsureMasterColumns_(masterSheet, headerRowIdx) {
  var headers = masterSheet.getRange(headerRowIdx + 1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
  var colMap = getColumnIndexMap(headers);
  var needed = [AMAZON_MASTER_COL_MODE, AMAZON_MASTER_COL_MAIN, AMAZON_MASTER_COL_PT];
  var added = [];
  var lastCol = headers.length;
  var i;
  for (i = 0; i < needed.length; i++) {
    if (colMap[needed[i]] === undefined) {
      lastCol++;
      masterSheet.getRange(headerRowIdx + 1, lastCol).setValue(needed[i]).setFontWeight('bold');
      added.push(needed[i]);
      colMap[needed[i]] = lastCol - 1;
    }
  }
  return { added: added, colMap: colMap };
}

function amazonImageInferMode_(masterRow, colMap) {
  var i;
  for (i = 1; i <= 10; i++) {
    var idx = colMap['楽天サブ画像' + i];
    if (idx !== undefined && masterRow[idx] != null && String(masterRow[idx]).trim() !== '') {
      return 'REUSE_RAKUTEN';
    }
  }
  return 'AMAZON_ONLY';
}

function amazonImageFindMasterRow_(mValues, headerRowIdx, colMap, pCode, cCode) {
  var idxP = colMap['親SKU'];
  var idxC = colMap['子SKU'];
  var r;
  for (r = headerRowIdx + 1; r < mValues.length; r++) {
    if (String(mValues[r][idxP] || '').trim() !== pCode) continue;
    var c = idxC !== undefined ? String(mValues[r][idxC] || '').trim() : '';
    if (c === cCode) return r;
  }
  return -1;
}

function amazonImageExtractDriveId_(formulaOrVal) {
  var s = String(formulaOrVal == null ? '' : formulaOrVal).trim();
  if (!s) return '';
  var m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;
  return '';
}

function amazonImageSplitPtIds_(val) {
  return String(val || '')
    .split(/[|,;\s]+/)
    .map(function (x) { return amazonImageExtractDriveId_(x); })
    .filter(function (x) { return !!x; });
}

function amazonImageFormula_(fileId) {
  return '=IMAGE("https://drive.google.com/uc?export=view&id=' + fileId + '", 2)';
}
