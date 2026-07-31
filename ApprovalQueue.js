/**
 * Lv1 承認キュー（出品①）— EC API / 楽天CSV / マスタ書込なし
 * 要件: docs/org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md
 *
 * Script Properties:
 *   APPROVAL_QUEUE_V1_ENABLED … 既定 false（true で候補作成可）
 *   APPROVAL_UI_ALLOWED_EMAILS … 承認操作を許可するメール（カンマ区切り）
 */

var APPROVAL_QUEUE_V1_PROP = 'APPROVAL_QUEUE_V1_ENABLED';
var APPROVAL_UI_ALLOWED_EMAILS_PROP = 'APPROVAL_UI_ALLOWED_EMAILS';
var APPROVAL_QUEUE_SHEET_NAME = '▼承認キュー(出品①)';
var APPROVAL_QUEUE_SCHEMA_VERSION = 'lv1-1';

var APPROVAL_QUEUE_HEADERS = [
  'recordType', 'batchId', 'createdAt', 'createdBy', 'status',
  'approvedAt', 'approvedBy', 'cancelledAt', 'inventoryMode', 'note',
  'sourceSummary', 'schemaVersion',
  'lineId', 'mall', 'masterRow', 'parentSku', 'childSku', 'productName',
  'checkboxCol', 'lineStatus', 'previewFlag', 'rejectReason'
];

/**
 * メニュー: レ点から承認候補バッチを作成（PENDING_APPROVAL）。ECは呼ばない。
 */
function menuApprovalQueueCreateCandidates() {
  var step = 'ApprovalQueueCreate';
  if (!getBoolScriptProperty_(APPROVAL_QUEUE_V1_PROP, false)) {
    var msgOff = '承認キューは無効です。Script Properties の ' + APPROVAL_QUEUE_V1_PROP + ' を true にしてください。';
    Logger.log('[' + step + '] state=FAILED ' + msgOff);
    try { SpreadsheetApp.getUi().alert(msgOff); } catch (e0) {}
    return;
  }
  try {
    Logger.log('[' + step + '] state=RUNNING');
    var result = approvalQueueCreateCandidates_();
    Logger.log('[' + step + '] state=DONE batchId=' + result.batchId +
      ' rakuten=' + result.rakutenCount + ' yahoo=' + result.yahooCount +
      ' amazon=' + result.amazonCount);
    try {
      SpreadsheetApp.getUi().alert(
        '承認候補を作成しました',
        'batchId=' + result.batchId +
          '\n楽天(親レ点)=' + result.rakutenCount +
          '\nYahoo(子レ点)=' + result.yahooCount +
          '\nAmazon(親+子加算)=' + result.amazonCount +
          '\nシート「' + APPROVAL_QUEUE_SHEET_NAME + '」を確認し、Webで承認してください。' +
          '\n（モールへの出品はしていません）',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e1) {}
  } catch (err) {
    Logger.log('[' + step + '] state=FAILED ' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert('候補作成失敗: ' + ((err && err.message) || err)); } catch (e2) {}
  }
}

/**
 * 承認WebのURL（未デプロイ等なら空文字）。
 * @return {string}
 */
function approvalQueueBuildWebUrl_() {
  var url = '';
  try {
    url = ScriptApp.getService().getUrl() || '';
  } catch (e) {
    url = '';
  }
  if (!url) return '';
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + 'action=approval_queue';
}

/**
 * メニュー: 承認WebのURL案内（デプロイは人間作業）
 */
function menuApprovalQueueShowWebHelp() {
  var webUrl = approvalQueueBuildWebUrl_();
  var body =
    '【Lv1 承認キュー】\n' +
    '1) Script Properties:\n' +
    '   ' + APPROVAL_QUEUE_V1_PROP + '=true\n' +
    '   ' + APPROVAL_UI_ALLOWED_EMAILS_PROP + '=あなたのGoogleメール\n' +
    '2) メニュー「承認候補を作成」を実行\n' +
    '3) 既存の削除用Webアプリを再デプロイ（または新規）\n' +
    '4) スマホで次のURLを開く（末尾のクエリ必須）:\n' +
    '   {ウェブアプリURL}?action=approval_queue\n\n' +
    (webUrl
      ? ('承認用URL:\n' + webUrl)
      : '（未デプロイ、またはURL取得不可）');
  try {
    SpreadsheetApp.getUi().alert('承認Webの使い方', body, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e3) {
    Logger.log('[ApprovalQueueHelp] ' + body);
  }
}

/**
 * Web表示用 HTML（Yahoo.js の doGet から action=approval_queue で呼ぶ）。
 * ※ doGet は Yahoo 削除Webと共存するため本ファイルでは定義しない。
 * @return {string}
 */
function approvalQueueGetHtmlForWebApp() {
  return approvalQueueBuildHtml_();
}

/** @return {{batchId:string, rakutenCount:number, yahooCount:number, amazonCount:number}} */
function approvalQueueCreateCandidates_() {
  var extracted = approvalQueueExtractCandidates_();
  if (!extracted.lines.length) {
    throw new Error('レ点付きの候補がありません（楽天=親レ点 / Yahoo=子レ点 / Amazon=親+子加算）。');
  }
  var batchId = approvalQueueNewBatchId_();
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
  var createdBy = '';
  try {
    createdBy = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'menu';
  } catch (e) {
    createdBy = 'menu';
  }
  var summary = 'rakutenParents=' + extracted.rakutenCount +
    ';yahooChildren=' + extracted.yahooCount +
    ';amazonLines=' + extracted.amazonCount;
  var sh = approvalQueueEnsureSheet_();
  var headerRow = [
    'HEADER', batchId, now, createdBy, 'PENDING_APPROVAL',
    '', '', '', 'ZERO', '',
    summary, APPROVAL_QUEUE_SCHEMA_VERSION,
    '', '', '', '', '', '',
    '', '', '', ''
  ];
  var rows = [headerRow];
  for (var i = 0; i < extracted.lines.length; i++) {
    var L = extracted.lines[i];
    rows.push([
      'LINE', batchId, '', '', '',
      '', '', '', '', '',
      '', '',
      L.lineId, L.mall, L.masterRow, L.parentSku, L.childSku, L.productName,
      L.checkboxCol, 'CANDIDATE', L.previewFlag || '', ''
    ]);
  }
  var start = sh.getLastRow() + 1;
  if (start < 2) start = 2;
  sh.getRange(start, 1, rows.length, APPROVAL_QUEUE_HEADERS.length).setValues(rows);
  return {
    batchId: batchId,
    rakutenCount: extracted.rakutenCount,
    yahooCount: extracted.yahooCount,
    amazonCount: extracted.amazonCount
  };
}

/**
 * Web / google.script.run 用: 最新 PENDING バッチを返す
 * @return {Object}
 */
function approvalQueueApiGetLatestPending() {
  approvalQueueAssertAllowedOrThrow_();
  return approvalQueueLoadLatestByStatus_('PENDING_APPROVAL');
}

/**
 * Lv2用: 最新 APPROVED バッチのうち mall=rakuten かつ lineStatus=APPROVED のみ。
 * @return {{found:boolean, batch:Object|null, lines:Array}}
 */
function approvalQueueGetLatestApprovedRakuten_() {
  var res = approvalQueueLoadLatestByStatus_('APPROVED');
  if (!res.found || !res.batch) return { found: false, batch: null, lines: [] };
  var lines = [];
  for (var i = 0; i < res.lines.length; i++) {
    var L = res.lines[i];
    if (String(L.mall) === 'rakuten' && String(L.lineStatus) === 'APPROVED') {
      lines.push(L);
    }
  }
  return { found: lines.length > 0, batch: res.batch, lines: lines };
}

/**
 * Lv3用: 最新 APPROVED バッチのうち mall=yahoo かつ lineStatus=APPROVED のみ。
 * @return {{found:boolean, batch:Object|null, lines:Array}}
 */
function approvalQueueGetLatestApprovedYahoo_() {
  var res = approvalQueueLoadLatestByStatus_('APPROVED');
  if (!res.found || !res.batch) return { found: false, batch: null, lines: [] };
  var lines = [];
  for (var i = 0; i < res.lines.length; i++) {
    var L = res.lines[i];
    if (String(L.mall) === 'yahoo' && String(L.lineStatus) === 'APPROVED') {
      lines.push(L);
    }
  }
  return { found: lines.length > 0, batch: res.batch, lines: lines };
}

/**
 * Lv4用: 最新 APPROVED バッチのうち mall=amazon かつ lineStatus=APPROVED のみ。
 * @return {{found:boolean, batch:Object|null, lines:Array}}
 */
function approvalQueueGetLatestApprovedAmazon_() {
  var res = approvalQueueLoadLatestByStatus_('APPROVED');
  if (!res.found || !res.batch) return { found: false, batch: null, lines: [] };
  var lines = [];
  for (var i = 0; i < res.lines.length; i++) {
    var L = res.lines[i];
    if (String(L.mall) === 'amazon' && String(L.lineStatus) === 'APPROVED') {
      lines.push(L);
    }
  }
  return { found: lines.length > 0, batch: res.batch, lines: lines };
}

/**
 * Web: 最新 APPROVED も参照用に返す（任意）
 * @return {Object}
 */
function approvalQueueApiGetLatestApproved() {
  approvalQueueAssertAllowedOrThrow_();
  return approvalQueueLoadLatestByStatus_('APPROVED');
}

/**
 * Web: 一括承認
 * @param {string} batchId
 * @return {{ok:boolean, message:string}}
 */
function approvalQueueApiApproveBatch(batchId) {
  approvalQueueAssertAllowedOrThrow_();
  var email = approvalQueueCurrentEmail_();
  var sh = approvalQueueEnsureSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: false, message: 'データがありません' };
  var id = String(batchId || '').trim();
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
  var found = false;
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[1]) !== id) continue;
    found = true;
    if (String(row[0]) === 'HEADER') {
      if (String(row[4]) === 'CANCELLED') {
        return { ok: false, message: '取消済みバッチは承認できません' };
      }
      row[4] = 'APPROVED';
      row[5] = now;
      row[6] = email;
    } else if (String(row[0]) === 'LINE') {
      if (String(row[19]) === 'CANDIDATE') {
        row[19] = 'APPROVED';
      }
    }
  }
  if (!found) return { ok: false, message: 'batchId が見つかりません' };
  sh.getRange(1, 1, data.length, APPROVAL_QUEUE_HEADERS.length).setValues(data);
  Logger.log('[ApprovalQueueApprove] batchId=' + id + ' by=' + email + ' state=DONE');
  return { ok: true, message: '承認しました: ' + id };
}

/**
 * Web: 行否認
 * @param {string} batchId
 * @param {string} lineId
 * @param {string} reason
 * @return {{ok:boolean, message:string}}
 */
function approvalQueueApiRejectLine(batchId, lineId, reason) {
  approvalQueueAssertAllowedOrThrow_();
  var sh = approvalQueueEnsureSheet_();
  var data = sh.getDataRange().getValues();
  var id = String(batchId || '').trim();
  var lid = String(lineId || '').trim();
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[0]) === 'LINE' && String(row[1]) === id && String(row[12]) === lid) {
      row[19] = 'REJECTED';
      row[21] = String(reason || '').substring(0, 200);
      sh.getRange(r + 1, 1, 1, APPROVAL_QUEUE_HEADERS.length).setValues([row]);
      return { ok: true, message: '否認: ' + lid };
    }
  }
  return { ok: false, message: '行が見つかりません' };
}

/**
 * Web: バッチ取消
 * @param {string} batchId
 * @return {{ok:boolean, message:string}}
 */
function approvalQueueApiCancelBatch(batchId) {
  approvalQueueAssertAllowedOrThrow_();
  var email = approvalQueueCurrentEmail_();
  var sh = approvalQueueEnsureSheet_();
  var data = sh.getDataRange().getValues();
  var id = String(batchId || '').trim();
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
  var found = false;
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[0]) === 'HEADER' && String(row[1]) === id) {
      row[4] = 'CANCELLED';
      row[7] = now;
      found = true;
      sh.getRange(r + 1, 1, 1, APPROVAL_QUEUE_HEADERS.length).setValues([row]);
      break;
    }
  }
  if (!found) return { ok: false, message: 'batchId が見つかりません' };
  Logger.log('[ApprovalQueueCancel] batchId=' + id + ' by=' + email + ' state=DONE');
  return { ok: true, message: '取消しました: ' + id };
}

/**
 * Web: inventoryMode 変更（ONE は確認済み前提でクライアントから呼ぶ）
 * @param {string} batchId
 * @param {string} mode ZERO|ONE
 * @return {{ok:boolean, message:string}}
 */
function approvalQueueApiSetInventoryMode(batchId, mode) {
  approvalQueueAssertAllowedOrThrow_();
  var m = String(mode || '').toUpperCase();
  if (m !== 'ZERO' && m !== 'ONE') return { ok: false, message: 'mode は ZERO または ONE' };
  var sh = approvalQueueEnsureSheet_();
  var data = sh.getDataRange().getValues();
  var id = String(batchId || '').trim();
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[0]) === 'HEADER' && String(row[1]) === id) {
      if (String(row[4]) === 'CANCELLED') {
        return { ok: false, message: '取消済みです' };
      }
      row[8] = m;
      sh.getRange(r + 1, 1, 1, APPROVAL_QUEUE_HEADERS.length).setValues([row]);
      return { ok: true, message: 'inventoryMode=' + m };
    }
  }
  return { ok: false, message: 'batchId が見つかりません' };
}

/** @return {{ok:boolean, email:string, allowed:boolean}} */
function approvalQueueApiWhoAmI() {
  var email = approvalQueueCurrentEmail_();
  return { ok: true, email: email, allowed: approvalQueueIsEmailAllowed_(email) };
}

// ----- internals -----

function approvalQueueNewBatchId_() {
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var rnd = String(Utilities.getUuid()).replace(/-/g, '').substring(0, 6);
  return 'A1_' + ts + '_' + rnd;
}

function approvalQueueCurrentEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function approvalQueueIsEmailAllowed_(email) {
  var raw = '';
  try {
    raw = PropertiesService.getScriptProperties().getProperty(APPROVAL_UI_ALLOWED_EMAILS_PROP) || '';
  } catch (e) {
    raw = '';
  }
  var list = String(raw).split(/[,;\s]+/).map(function (s) {
    return String(s || '').trim().toLowerCase();
  }).filter(Boolean);
  if (!list.length) return false;
  var em = String(email || '').trim().toLowerCase();
  if (!em) return false;
  for (var i = 0; i < list.length; i++) {
    if (list[i] === em) return true;
  }
  return false;
}

function approvalQueueAssertAllowedOrThrow_() {
  var email = approvalQueueCurrentEmail_();
  if (!approvalQueueIsEmailAllowed_(email)) {
    throw new Error('承認操作が許可されていません。Script Properties の ' +
      APPROVAL_UI_ALLOWED_EMAILS_PROP + ' にログイン中のメールを設定してください。');
  }
}

function approvalQueueEnsureSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートがありません');
  var name = APPROVAL_QUEUE_SHEET_NAME;
  if (name !== '▼承認キュー(出品①)') {
    throw new Error('承認キューシート名定数が不正です');
  }
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, APPROVAL_QUEUE_HEADERS.length).setValues([APPROVAL_QUEUE_HEADERS]);
    Logger.log('[ApprovalQueue] created sheet=' + name);
  } else {
    var h = String(sh.getRange(1, 1).getValue() || '').trim();
    if (h !== 'recordType') {
      if (sh.getLastRow() === 0) {
        sh.getRange(1, 1, 1, APPROVAL_QUEUE_HEADERS.length).setValues([APPROVAL_QUEUE_HEADERS]);
      } else {
        throw new Error('承認キューシートのヘッダーが recordType ではありません');
      }
    }
  }
  return sh;
}

/**
 * @return {{lines:Array<Object>, rakutenCount:number, yahooCount:number, amazonCount:number}}
 */
function approvalQueueExtractCandidates_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterName = (typeof MASTER_SHEET_NAME !== 'undefined') ? MASTER_SHEET_NAME : '▼商品マスタ(人間作業用)';
  var ckName = (typeof CHECKBOX_HEADER_NAME !== 'undefined') ? CHECKBOX_HEADER_NAME : '出品CK';
  var sh = ss.getSheetByName(masterName);
  if (!sh) throw new Error('マスタシートが見つかりません: ' + masterName);
  var values = sh.getDataRange().getValues();
  if (!values.length) throw new Error('マスタが空です');

  var headerRowIdx = -1;
  var limit = Math.min(values.length, 25);
  for (var r = 0; r < limit; r++) {
    var row = values[r] || [];
    if (row.indexOf('親SKU') !== -1 && row.indexOf(ckName) !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error('マスタヘッダー（親SKU・出品CK）が見つかりません');

  var headers = values[headerRowIdx];
  var col = {};
  for (var c = 0; c < headers.length; c++) {
    var key = String(headers[c] || '').trim();
    if (key && col[key] == null) col[key] = c;
  }
  var iParent = col['親SKU'];
  var iChild = col['子SKU'];
  var iCk = col[ckName];
  var iName = col['商品名'] != null ? col['商品名'] : col['商品名ベース'];
  var iStock = col['在庫数'];
  if (iParent == null || iCk == null) {
    throw new Error('必須列がありません（親SKU / 出品CK）');
  }

  var isTrue = (typeof yahooMasterCheckboxIsTrue_ === 'function')
    ? yahooMasterCheckboxIsTrue_
    : function (cell) { return cell === true || String(cell).toUpperCase() === 'TRUE'; };

  var lines = [];
  var rakutenCount = 0;
  var yahooCount = 0;
  var amazonCount = 0;
  var lineSeq = 0;

  for (var i = headerRowIdx + 1; i < values.length; i++) {
    var dataRow = values[i] || [];
    if (!isTrue(dataRow[iCk])) continue;
    var parentSku = String(dataRow[iParent] != null ? dataRow[iParent] : '').trim();
    var childSku = iChild != null ? String(dataRow[iChild] != null ? dataRow[iChild] : '').trim() : '';
    var productName = iName != null ? String(dataRow[iName] != null ? dataRow[iName] : '').trim() : '';
    var stock = iStock != null ? dataRow[iStock] : '';
    var stockNum = (stock === '' || stock == null) ? null : Number(stock);
    var preview = (stockNum != null && !isNaN(stockNum) && stockNum > 0) ? 'MAY_SKIP_IN_STOCK' : '';
    var masterRow = i + 1; // 1-based

    // 楽天: 親行（子SKU空）かつレ点
    if (!childSku) {
      lineSeq++;
      lines.push({
        lineId: 'L' + lineSeq,
        mall: 'rakuten',
        masterRow: masterRow,
        parentSku: parentSku,
        childSku: '',
        productName: productName,
        checkboxCol: ckName,
        previewFlag: preview
      });
      rakutenCount++;
      // Amazon: 親行を加算（同一バッチ・個別承認）
      lineSeq++;
      lines.push({
        lineId: 'L' + lineSeq,
        mall: 'amazon',
        masterRow: masterRow,
        parentSku: parentSku,
        childSku: '',
        productName: productName,
        checkboxCol: ckName,
        previewFlag: preview
      });
      amazonCount++;
    } else {
      // Yahoo: 子SKU行のレ点のみ
      lineSeq++;
      lines.push({
        lineId: 'L' + lineSeq,
        mall: 'yahoo',
        masterRow: masterRow,
        parentSku: parentSku,
        childSku: childSku,
        productName: productName,
        checkboxCol: ckName,
        previewFlag: preview
      });
      yahooCount++;
      // Amazon: 子行を加算（同一バッチ・個別承認）
      lineSeq++;
      lines.push({
        lineId: 'L' + lineSeq,
        mall: 'amazon',
        masterRow: masterRow,
        parentSku: parentSku,
        childSku: childSku,
        productName: productName,
        checkboxCol: ckName,
        previewFlag: preview
      });
      amazonCount++;
    }
  }

  return {
    lines: lines,
    rakutenCount: rakutenCount,
    yahooCount: yahooCount,
    amazonCount: amazonCount
  };
}

/**
 * @param {string} status
 * @return {Object}
 */
function approvalQueueLoadLatestByStatus_(status) {
  var sh = approvalQueueEnsureSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return { found: false, batch: null, lines: [] };
  }
  var want = String(status || '');
  var header = null;
  for (var r = data.length - 1; r >= 1; r--) {
    var row = data[r];
    if (String(row[0]) === 'HEADER' && String(row[4]) === want) {
      header = {
        batchId: String(row[1]),
        createdAt: row[2],
        createdBy: row[3],
        status: String(row[4]),
        approvedAt: row[5],
        approvedBy: row[6],
        cancelledAt: row[7],
        inventoryMode: String(row[8] || 'ZERO'),
        note: row[9],
        sourceSummary: row[10],
        schemaVersion: row[11]
      };
      break;
    }
  }
  if (!header) return { found: false, batch: null, lines: [] };
  var lines = [];
  for (var j = 1; j < data.length; j++) {
    var lr = data[j];
    if (String(lr[0]) === 'LINE' && String(lr[1]) === header.batchId) {
      lines.push({
        lineId: String(lr[12]),
        mall: String(lr[13]),
        masterRow: lr[14],
        parentSku: String(lr[15] || ''),
        childSku: String(lr[16] || ''),
        productName: String(lr[17] || ''),
        checkboxCol: String(lr[18] || ''),
        lineStatus: String(lr[19] || ''),
        previewFlag: String(lr[20] || ''),
        rejectReason: String(lr[21] || '')
      });
    }
  }
  return { found: true, batch: header, lines: lines };
}

function approvalQueueBuildHtml_() {
  return [
    '<!DOCTYPE html><html><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width,initial-scale=1">',
    '<style>',
    'body{font-family:sans-serif;margin:12px;background:#111;color:#eee;}',
    'button{margin:4px 4px 4px 0;padding:10px 12px;font-size:15px;}',
    '.ok{background:#2d6a4f;color:#fff;border:0;border-radius:8px;}',
    '.warn{background:#9a3412;color:#fff;border:0;border-radius:8px;}',
    '.neutral{background:#334155;color:#fff;border:0;border-radius:8px;}',
    'table{border-collapse:collapse;width:100%;font-size:13px;margin-top:8px;}',
    'th,td{border:1px solid #333;padding:6px;text-align:left;}',
    'th{background:#1e293b;}',
    '.meta{font-size:13px;color:#94a3b8;margin:6px 0;}',
    '.err{color:#fca5a5;}',
    '.flag{color:#fbbf24;}',
    '</style></head><body>',
    '<h2>承認キュー（出品①）</h2>',
    '<p class="meta">モールへは出品しません。レ点候補の朝承認のみ。</p>',
    '<div id="who" class="meta">認証確認中…</div>',
    '<div id="msg"></div>',
    '<div id="actions"></div>',
    '<div id="batch"></div>',
    '<script>',
    'function setMsg(t, isErr){var m=document.getElementById("msg");m.textContent=t||"";m.className=isErr?"err":"meta";}',
    'function load(){',
    '  google.script.run.withSuccessHandler(function(w){',
    '    document.getElementById("who").textContent="ログイン: "+(w.email||"(不明)")+" / 許可="+(w.allowed?"YES":"NO");',
    '    if(!w.allowed){setMsg("許可メールに含まれていません。Script Properties を確認してください。",true);return;}',
    '    refreshPending();',
    '  }).withFailureHandler(function(e){setMsg(String(e&&e.message||e),true);}).approvalQueueApiWhoAmI();',
    '}',
    'function refreshPending(){',
    '  google.script.run.withSuccessHandler(render).withFailureHandler(function(e){setMsg(String(e&&e.message||e),true);}).approvalQueueApiGetLatestPending();',
    '}',
    'function render(res){',
    '  var a=document.getElementById("actions"); var b=document.getElementById("batch");',
    '  if(!res||!res.found){',
    '    a.innerHTML="<p>PENDING のバッチがありません。スプシメニューで「承認候補を作成」を実行してください。</p>";',
    '    b.innerHTML=""; return;',
    '  }',
    '  var bat=res.batch; var lines=res.lines||[];',
    '  a.innerHTML="";',
    '  var bApprove=document.createElement("button"); bApprove.className="ok"; bApprove.textContent="一括承認";',
    '  bApprove.onclick=function(){',
    '    if(!confirm("このバッチを承認しますか？\\n"+bat.batchId))return;',
    '    google.script.run.withSuccessHandler(function(r){setMsg(r.message,!r.ok);refreshPending();}).approvalQueueApiApproveBatch(bat.batchId);',
    '  };',
    '  var bCancel=document.createElement("button"); bCancel.className="warn"; bCancel.textContent="バッチ取消";',
    '  bCancel.onclick=function(){',
    '    if(!confirm("取消しますか？\\n"+bat.batchId))return;',
    '    google.script.run.withSuccessHandler(function(r){setMsg(r.message,!r.ok);refreshPending();}).approvalQueueApiCancelBatch(bat.batchId);',
    '  };',
    '  var bMode=document.createElement("button"); bMode.className="neutral"; bMode.textContent="在庫モード→ONE";',
    '  bMode.onclick=function(){',
    '    if(!confirm("在庫1フォールバック(ONE)にしますか？画面明示が必要です。"))return;',
    '    google.script.run.withSuccessHandler(function(r){setMsg(r.message,!r.ok);refreshPending();}).approvalQueueApiSetInventoryMode(bat.batchId,"ONE");',
    '  };',
    '  var bZero=document.createElement("button"); bZero.className="neutral"; bZero.textContent="在庫モード→ZERO";',
    '  bZero.onclick=function(){',
    '    google.script.run.withSuccessHandler(function(r){setMsg(r.message,!r.ok);refreshPending();}).approvalQueueApiSetInventoryMode(bat.batchId,"ZERO");',
    '  };',
    '  var bReload=document.createElement("button"); bReload.className="neutral"; bReload.textContent="再読込";',
    '  bReload.onclick=refreshPending;',
    '  a.appendChild(bApprove); a.appendChild(bCancel); a.appendChild(bMode); a.appendChild(bZero); a.appendChild(bReload);',
    '  var html="<div class=\\"meta\\">batchId=<b>"+bat.batchId+"</b><br>status="+bat.status+',
    '    " / inventoryMode="+bat.inventoryMode+"<br>"+(bat.sourceSummary||"")+"</div>";',
    '  html+="<table><tr><th>mall</th><th>SKU</th><th>名前</th><th>status</th><th>flag</th><th></th></tr>";',
    '  for(var i=0;i<lines.length;i++){',
    '    var L=lines[i];',
    '    var sku=(L.parentSku||"")+(L.childSku?("-"+L.childSku):"");',
    '    html+="<tr><td>"+L.mall+"</td><td>"+sku+"</td><td>"+(L.productName||"")+"</td><td>"+L.lineStatus+"</td>";',
    '    html+="<td class=\\"flag\\">"+(L.previewFlag||"")+"</td><td>";',
    '    if(L.lineStatus==="CANDIDATE"){',
    '      html+="<button data-lid=\\""+L.lineId+"\\" class=\\"rej\\">否認</button>";',
    '    }',
    '    html+="</td></tr>";',
    '  }',
    '  html+="</table>";',
    '  b.innerHTML=html;',
    '  var btns=b.querySelectorAll("button.rej");',
    '  for(var k=0;k<btns.length;k++){',
    '    btns[k].onclick=function(ev){',
    '      var lid=ev.target.getAttribute("data-lid");',
    '      var reason=prompt("否認理由（任意）","")||"";',
    '      google.script.run.withSuccessHandler(function(r){setMsg(r.message,!r.ok);refreshPending();}).approvalQueueApiRejectLine(bat.batchId,lid,reason);',
    '    };',
    '  }',
    '}',
    'load();',
    '</script></body></html>'
  ].join('');
}
