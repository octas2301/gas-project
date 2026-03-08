// ==========================================
// 在庫管理アプリ (Octas) - エントリ・共通
// ==========================================

var SPREADSHEET_ID_DEFAULT = '1_kqBcQTcL0Q5--KsAdxMYbPMsaToLEXxOMuDw_z-kwE';

/** デバッグログ用シート名 */
var DEBUG_LOG_SHEET_NAME = 'デバッグログ';

/**
 * GAS側のログをスプレッドシート「デバッグログ」に追記する。調査用。失敗しても doGet は続行する。
 * @param {string} msg メッセージ
 */
function debugLogToSheet(msg) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(DEBUG_LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(DEBUG_LOG_SHEET_NAME);
      sheet.getRange(1, 1, 1, 3).setValues([['日時', 'メッセージ', '備考']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    var row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(new Date());
    sheet.getRange(row, 2).setValue(String(msg));
  } catch (e) {
    // ログ失敗は無視
  }
}

/**
 * 使用するスプレッドシートを取得する。
 * Script Properties に SPREADSHEET_ID が設定されていればそれを使用、なければデフォルトIDを使用。
 */
function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || SPREADSHEET_ID_DEFAULT;
  return SpreadsheetApp.openById(id);
}

/** キャッシュ有効期限（秒）。1～2分の体感改善用。 */
var CACHE_TTL_SEC = 90;

/**
 * ドキュメントキャッシュを取得する。
 * @return {Cache} CacheService.getDocumentCache()
 */
function getAppCache() {
  return CacheService.getDocumentCache();
}

/**
 * 入荷・出荷・棚卸・確定後修正など書き込み時に、差異・在庫一覧のキャッシュを無効化する。
 */
function invalidateListCachesOnWrite() {
  try {
    var c = getAppCache();
    c.remove('varianceList');
    c.remove('currentStockList');
  } catch (e) {}
}

/**
 * 現在のユーザー（ログイン中のメールアドレス）がアクセス許可シートに含まれるかチェックする。
 * Web アプリでは getActiveUser().getEmail() が空になることがあるため、空の場合は getEffectiveUser()（実行ユーザー＝オーナー）でフォールバックする。
 * doGet の phase_init 短縮のため、結果をキャッシュする（TTL 60秒）。アクセス許可シートを変更した場合は最大60秒で反映される。
 * @return {boolean} 許可されていれば true
 */
function isAccessAllowed() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) userEmail = Session.getEffectiveUser().getEmail();
    if (!userEmail) return false;
    var cacheKey = 'accessAllowed_' + userEmail.replace(/[^a-zA-Z0-9@._-]/g, '_').slice(0, 200);
    var cache = getAppCache();
    var cached = cache.get(cacheKey);
    if (cached !== null) return cached === '1';
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('アクセス許可');
    if (!sheet) {
      try { cache.put(cacheKey, '0', 60); } catch (e) {}
      return false;
    }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      try { cache.put(cacheKey, '0', 60); } catch (e) {}
      return false;
    }
    var header = data[0];
    var emailCol = header.indexOf('メールアドレス');
    if (emailCol < 0) {
      try { cache.put(cacheKey, '0', 60); } catch (e) {}
      return false;
    }
    for (var i = 1; i < data.length; i++) {
      var email = (data[i][emailCol] || '').toString().trim().toLowerCase();
      if (email && userEmail.toLowerCase() === email) {
        try { cache.put(cacheKey, '1', 60); } catch (e) {}
        return true;
      }
    }
    try { cache.put(cacheKey, '0', 60); } catch (e) {}
    return false;
  } catch (e) {
    return false;
  }
}

/**
 * Web アプリとして配信。常に index.html を返す。アクセス未許可の場合はエラーHTMLを返す。
 */
function sanitizeJsonpCallbackName(name) {
  var callback = (name || '').toString().trim();
  if (!callback) return '';
  if (!/^[A-Za-z0-9_$.]+$/.test(callback)) return '';
  return callback;
}

function createJsonpOutput(callbackName, payload) {
  var callback = sanitizeJsonpCallbackName(callbackName) || 'callback';
  var body = callback + '(' + JSON.stringify(payload || {}) + ');';
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function handleExternalScanApiGet(params) {
  var api = (params.api || '').toString().trim();
  var accessAllowed = isAccessAllowed();
  if (!accessAllowed) {
    if (api === 'externalLookupJsonp') {
      return createJsonpOutput(params.callback, { ok: false, error: 'access denied' });
    }
    return ContentService.createTextOutput('access denied').setMimeType(ContentService.MimeType.TEXT);
  }

  if (api === 'externalLookupJsonp') {
    var jan = (params.jan || '').toString().trim();
    var product = jan ? findProductByJan(jan, jan) : null;
    return createJsonpOutput(params.callback, {
      ok: true,
      jan: jan,
      product: product ? {
        singleJAN: product.singleJAN,
        boxJAN: product.boxJAN,
        productName: product.productName,
        cost: product.cost,
        avgCost: product.avgCost,
        boxInCount: product.boxInCount
      } : null
    });
  }

  if (api === 'externalScanLog') {
    var msg = (params.msg || '').toString();
    debugLogToSheet('[外部スキャンPoC] ' + msg.substring(0, 1000));
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  }

  return ContentService.createTextOutput('unsupported api').setMimeType(ContentService.MimeType.TEXT);
}

function doGet(e) {
  var serverLog = [];
  function log(msg) {
    var line = (new Date()).toISOString().slice(11, 23) + ' ' + String(msg);
    serverLog.push(line);
    try { debugLogToSheet(msg); } catch (z) {}
  }
  function escapeForHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  var errHtml = function(msg, loginInfo, logs) {
    var body = '<h1>' + (msg || 'アクセス権がありません') + '</h1><p>このアプリを利用するには、管理者にメールアドレスを「アクセス許可」シートに登録してもらってください。</p>' +
      '<p>ログイン中: ' + (loginInfo || '取得できません') + '</p>';
    if (logs && logs.length > 0) {
      body += '<p style="text-align:left;margin:1rem;font-size:12px;font-family:monospace;white-space:pre-wrap;background:#fff3cd;padding:8px;">【サーバー調査ログ】\n' + logs.map(escapeForHtml).join('\n') + '</p>';
    }
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<style>body{font-family:sans-serif;padding:2rem;text-align:center;color:#c00;} h1{font-size:1.2rem;}</style></head><body>' + body + '</body></html>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).setTitle('在庫管理アプリ');
  };
  try {
    var t0DoGet = new Date().getTime();
    log('doGet start');
    var params = e && e.parameter ? e.parameter : {};
    var view = params.view || params.v || '';
    log('request view=' + (view || '(なし)'));
    var loginInfo = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '取得できません';
    log('loginInfo=' + loginInfo);

    if (params.api) {
      return handleExternalScanApiGet(params);
    }

    if (!isAccessAllowed()) {
      log('access denied');
      var doGetMsDeny = new Date().getTime() - t0DoGet;
      log('doGet end ' + doGetMsDeny + 'ms');
      try { appendPerfLog('doGet', doGetMsDeny, ''); } catch (z) {}
      return errHtml('アクセス権がありません', loginInfo, serverLog);
    }
    log('access ok');
    var t1AfterAccess = new Date().getTime();
    try { appendPerfLog('doGet_phase_init', t1AfterAccess - t0DoGet, ''); } catch (z) {}

    var baseUrl = '';
    try {
      var svc = ScriptApp.getService();
      if (svc) baseUrl = (svc.getUrl() || '').toString().trim();
      if (baseUrl && baseUrl.indexOf('?') >= 0) baseUrl = baseUrl.split('?')[0];
    } catch (z) {}
    if (!baseUrl) baseUrl = '';

    var gasExecUrl = baseUrl;
    var externalScanBaseUrl = (PropertiesService.getScriptProperties().getProperty('EXTERNAL_SCAN_BASE_URL') || '').toString().trim();
    if (!externalScanBaseUrl) {
      externalScanBaseUrl = 'https://octas2301.github.io/gas-project/inventory-app/external-scan-poc.html';
    }
    var sep = (gasExecUrl.indexOf('?') >= 0) ? '&' : '?';
    var returnUrlShipping = gasExecUrl + sep + 'view=shipping';
    var returnUrlMove = gasExecUrl + sep + 'view=move';
    var returnUrlReceiving = gasExecUrl + sep + 'view=receiving';
    var returnUrlCount = gasExecUrl + sep + 'view=count';
    var returnUrlAdjust = gasExecUrl + sep + 'view=count&adjust=1';
    var returnUrlProducts = gasExecUrl + sep + 'view=products';

    if (params.test === '1' || params.view === 'scantest') {
      var jan = (params.jan || '').toString().trim();
      var callbackUrl = baseUrl + (baseUrl.indexOf('?') >= 0 ? '&' : '?') + 'test=1&view=scantest&jan=EAN';
      var encodedCallback = encodeURIComponent(callbackUrl);
      var callbackUrlQrbot = baseUrl + (baseUrl.indexOf('?') >= 0 ? '&' : '?') + 'test=1&view=scantest&jan={CODE}';
      var encodedCallbackQrbot = encodeURIComponent(callbackUrlQrbot);
      var scantestHtml = HtmlService.createTemplateFromFile('scantest');
      scantestHtml.baseUrl = baseUrl;
      scantestHtml.baseUrlEscaped = escapeForHtml(baseUrl || '(取得できません)');
      scantestHtml.jan = jan;
      scantestHtml.janEscaped = escapeForHtml(jan);
      scantestHtml.encodedCallback = encodedCallback;
      scantestHtml.encodedCallbackQrbot = encodedCallbackQrbot;
      return scantestHtml.evaluate()
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .setTitle('pic2shop 動作確認テスト')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    var safeView = ['shipping', 'move', 'receiving', 'count', 'variance', 'products', 'stock', 'settings'].indexOf(String(view)) >= 0 ? view : 'shipping';
    log('safeView=' + safeView);
    var toastMsg = params.toast || '';
    var janFromUrl = (params.jan || '').toString().trim();
    var adjustFromUrl = (params.adjust === '1' || params.adjust === 'true');
    // 初回表示を軽くするためマスタは doGet では取得せず、クライアントの loadMasters() で非同期取得
    var locationsOptionsHtml = '';
    var reasonsOptionsHtml = '';
    var tantouOptionsHtml = '';

    var html = HtmlService.createTemplateFromFile('index');
    html.view = safeView;
    html.viewAttr = escapeForHtml(String(safeView || 'shipping'));
    html.baseUrl = baseUrl;
    html.serverLogHtml = serverLog.map(escapeForHtml).join('\n');
    html.showLog = (params.log === '1' || params.debug === '1');
    html.locationsOptionsHtml = locationsOptionsHtml;
    html.reasonsOptionsHtml = reasonsOptionsHtml;
    html.tantouOptionsHtml = tantouOptionsHtml;
    html.toastMsg = escapeForHtml(toastMsg);
    html.janFromUrl = janFromUrl;
    html.janFromUrlEscaped = escapeForHtml(janFromUrl);
    html.adjustFromUrl = adjustFromUrl;
    html.gasExecUrl = gasExecUrl;
    html.externalScanBaseUrl = externalScanBaseUrl;
    html.externalScanBaseUrlForAttr = escapeForHtml(externalScanBaseUrl);
    html.returnUrlShippingEnc = encodeURIComponent(returnUrlShipping);
    html.returnUrlMoveEnc = encodeURIComponent(returnUrlMove);
    html.returnUrlReceivingEnc = encodeURIComponent(returnUrlReceiving);
    html.returnUrlCountEnc = encodeURIComponent(returnUrlCount);
    html.returnUrlAdjustEnc = encodeURIComponent(returnUrlAdjust);
    html.returnUrlProductsEnc = encodeURIComponent(returnUrlProducts);
    var t2TemplateVars = new Date().getTime();
    log('template vars set');
    try { appendPerfLog('doGet_phase_templateVars', t2TemplateVars - t1AfterAccess, ''); } catch (z) {}

    var t3BeforeEvaluate = new Date().getTime();
    var out = html.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('在庫管理アプリ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
    var t4AfterEvaluate = new Date().getTime();
    try { appendPerfLog('doGet_phase_evaluate', t4AfterEvaluate - t3BeforeEvaluate, ''); } catch (z) {}
    var doGetMs = new Date().getTime() - t0DoGet;
    log('doGet end ' + doGetMs + 'ms');
    try { appendPerfLog('doGet', doGetMs, ''); } catch (z) {}
    return out;
  } catch (err) {
    log('ERR: ' + (err && err.message ? err.message : String(err)));
    var doGetMs = new Date().getTime() - t0DoGet;
    log('doGet end ' + doGetMs + 'ms');
    try { appendPerfLog('doGet', doGetMs, ''); } catch (z) {}
    return errHtml('エラーが発生しました', (err && err.message) ? String(err.message) : '', serverLog);
  }
}

/**
 * POST はフォーム送信を処理。出荷はクライアント側 apiShipping 推奨。POST が来た場合もここで受け、処理後に doGet で同じ画面を返す。
 */
function doPost(e) {
  e = e || {};
  var params = e.parameter || {};
  try {
    if (params.action === 'shipping') {
      var jan = (params.ship_jan || '').toString().trim();
      var from = (params.ship_from || '').toString().trim();
      if (jan && from) {
        var p = findProductByJan(jan, jan);
        var productName = (p && p.productName) ? p.productName : jan;
        var tantou = apiGetTantou();
        var result = appendShippingRow(jan, productName, from, tantou);
        if (result.ok) appendLog(getCurrentUserEmail(), tantou, '出荷', '登録', jan);
        return doGet({ parameter: { view: 'shipping', toast: result.ok ? '登録しました' : (result.error || 'エラー') } });
      }
      return doGet({ parameter: { view: 'shipping', toast: 'JANと出荷元を入力してください' } });
    }
    return doGet(e);
  } catch (err) {
    var msg = (err && (err.message || String(err))) || '不明なエラー';
    return doGet({ parameter: { view: 'shipping', toast: 'エラー: ' + msg } });
  }
}

/**
 * 現在のユーザーのメールアドレスを返す（クライアント用）。
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail() || '';
}

/**
 * アクセス許可チェックをクライアントから呼ぶ用（doGet で既にチェック済みだが、API としても使う）。
 */
function checkAccess() {
  return isAccessAllowed();
}

/**
 * スキャン診断ログをクライアントから記録する。デバッグログシートに [スキャン診断] 付きで追記。
 * @param {string} msg 診断メッセージ（BarcodeDetector有無・UA・getUserMedia結果・エラー内容など）
 */
function apiLogScanDiagnostics(msg) {
  try {
    debugLogToSheet('[スキャン診断] ' + (msg || ''));
  } catch (e) {}
}

// --- クライアント用 API（google.script.run から呼ぶ）---

function apiGetTantou() {
  var t0 = new Date().getTime();
  var result = PropertiesService.getUserProperties().getProperty('CURRENT_TANTOU') || '';
  appendPerfLog('apiGetTantou', new Date().getTime() - t0, '');
  return result;
}

function apiSetTantou(name) {
  var t0 = new Date().getTime();
  PropertiesService.getUserProperties().setProperty('CURRENT_TANTOU', (name || '').toString().trim());
  appendPerfLog('apiSetTantou', new Date().getTime() - t0, name || '');
}

function apiLookupProduct(jan) {
  var t0 = new Date().getTime();
  var p = findProductByJan(jan, jan);
  var result = p ? { singleJAN: p.singleJAN, boxJAN: p.boxJAN, productName: p.productName, cost: p.cost, avgCost: p.avgCost, boxInCount: p.boxInCount } : null;
  appendPerfLog('apiLookupProduct', new Date().getTime() - t0, '');
  return result;
}

function apiCheckDuplicate(singleJan, boxJan) {
  return checkProductDuplicate(singleJan, boxJan);
}

function apiGetLocations() {
  var cache = getAppCache();
  var cached = cache.get('locations');
  if (cached !== null) { try { return JSON.parse(cached); } catch (e) {} }
  var t0 = new Date().getTime();
  var result = getMasterList('場所マスタ');
  try { cache.put('locations', JSON.stringify(result), CACHE_TTL_SEC); } catch (e) {}
  appendPerfLog('apiGetLocations', new Date().getTime() - t0, '');
  return result;
}
function apiGetReasons() {
  var cache = getAppCache();
  var cached = cache.get('reasons');
  if (cached !== null) { try { return JSON.parse(cached); } catch (e) {} }
  var t0 = new Date().getTime();
  var result = getMasterList('理由マスタ');
  try { cache.put('reasons', JSON.stringify(result), CACHE_TTL_SEC); } catch (e) {}
  appendPerfLog('apiGetReasons', new Date().getTime() - t0, '');
  return result;
}
function apiGetTantouList() {
  var cache = getAppCache();
  var cached = cache.get('tantouList');
  if (cached !== null) { try { return JSON.parse(cached); } catch (e) {} }
  var t0 = new Date().getTime();
  var result = getMasterList('担当者マスタ');
  try { cache.put('tantouList', JSON.stringify(result), CACHE_TTL_SEC); } catch (e) {}
  appendPerfLog('apiGetTantouList', new Date().getTime() - t0, '');
  return result;
}

/**
 * 場所・理由・担当者マスタと現在担当者を1回の呼び出しで返す。初回表示の往復回数を減らす。
 * @return {{ locations: Array, reasons: Array, tantouList: Array, currentTantou: string }}
 */
function apiGetMasters() {
  var t0 = new Date().getTime();
  var cache = getAppCache();
  var locations = null, reasons = null, tantouList = null;
  try {
    var cLoc = cache.get('locations');
    var cRea = cache.get('reasons');
    var cTan = cache.get('tantouList');
    if (cLoc !== null && cRea !== null && cTan !== null) {
      locations = JSON.parse(cLoc);
      reasons = JSON.parse(cRea);
      tantouList = JSON.parse(cTan);
    }
  } catch (e) {}
  if (locations === null || reasons === null || tantouList === null) {
    locations = getMasterList('場所マスタ') || [];
    reasons = getMasterList('理由マスタ') || [];
    tantouList = getMasterList('担当者マスタ') || [];
    try {
      cache.put('locations', JSON.stringify(locations), CACHE_TTL_SEC);
      cache.put('reasons', JSON.stringify(reasons), CACHE_TTL_SEC);
      cache.put('tantouList', JSON.stringify(tantouList), CACHE_TTL_SEC);
    } catch (e) {}
  }
  var currentTantou = (PropertiesService.getUserProperties().getProperty('CURRENT_TANTOU') || '').toString().trim();
  appendPerfLog('apiGetMasters', new Date().getTime() - t0, '');
  return { locations: locations, reasons: reasons, tantouList: tantouList, currentTantou: currentTantou };
}

/** クライアントから呼ばれる。ページ表示〜マスタ取得までの時間をパフォーマンスログに追記する。 */
function apiPerfLogClient(t_domMs, t_mastersMs) {
  try { appendPerfLogClient(t_domMs, t_mastersMs); } catch (e) {}
}

/** クライアントから呼ばれる。区間別の詳細パフォーマンスをパフォーマンスログに追記する。 */
function apiPerfLogClientDetail(t_responseMs, t_domMs, t_apiGetMastersMs, t_mastersTotalMs) {
  try { appendPerfLogClientDetail(t_responseMs, t_domMs, t_apiGetMastersMs, t_mastersTotalMs); } catch (e) {}
}

function apiGetVarianceList() {
  var cache = getAppCache();
  var cached = cache.get('varianceList');
  if (cached !== null) { try { return JSON.parse(cached); } catch (e) {} }
  var t0 = new Date().getTime();
  var result = getVarianceList();
  try { cache.put('varianceList', JSON.stringify(result), CACHE_TTL_SEC); } catch (e) {}
  appendPerfLog('apiGetVarianceList', new Date().getTime() - t0, '');
  return result;
}

function apiCountStart() {
  var t0 = new Date().getTime();
  try {
    var active = getActiveCountSession();
    if (active && active.sessionId) {
      debugLogToSheet('apiCountStart: blocking, existing sessionId=' + (active.sessionId || '').toString().substring(0, 8) + '...');
      appendPerfLog('apiCountStart', new Date().getTime() - t0, '');
      return { ok: false, error: '未完了の棚卸があります。続きはこのまま入力して完了を押すか、先に完了してからスタートしてください。' };
    }
    var result = createCountSession();
    if (result && result.ok && result.sessionId) {
      debugLogToSheet('apiCountStart: created sessionId=' + (result.sessionId || '').toString().substring(0, 8) + '...');
    }
    appendPerfLog('apiCountStart', new Date().getTime() - t0, '');
    return result;
  } catch (e) {
    debugLogToSheet('apiCountStart error: ' + (e && e.message ? e.message : String(e)));
    appendPerfLog('apiCountStart', new Date().getTime() - t0, '');
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function apiCountComplete(sessionId) {
  var t0 = new Date().getTime();
  try {
    var result = completeCountSession(sessionId);
    if (result && result.ok) invalidateListCachesOnWrite();
    appendPerfLog('apiCountComplete', new Date().getTime() - t0, '');
    return result;
  } catch (e) {
    appendPerfLog('apiCountComplete', new Date().getTime() - t0, '');
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function apiGetCountSession() {
  var t0 = new Date().getTime();
  try {
    var session = getActiveCountSession();
    if (session && session.sessionId) {
      debugLogToSheet('apiGetCountSession: sessionId=' + (session.sessionId || '').toString().substring(0, 8) + '...');
      appendPerfLog('apiGetCountSession', new Date().getTime() - t0, '');
      return {
        sessionId: (session.sessionId || '').toString(),
        startTime: session.startTime instanceof Date ? Utilities.formatDate(session.startTime, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss") : (session.startTime || '').toString()
      };
    }
    debugLogToSheet('apiGetCountSession: null');
    appendPerfLog('apiGetCountSession', new Date().getTime() - t0, '');
    return session;
  } catch (e) {
    debugLogToSheet('apiGetCountSession error: ' + (e && e.message ? e.message : String(e)));
    appendPerfLog('apiGetCountSession', new Date().getTime() - t0, '');
    return null;
  }
}

function apiGetCurrentStockList() {
  var appCache = getAppCache();
  var cached = appCache.get('currentStockList');
  if (cached !== null) { try { return JSON.parse(cached); } catch (e) {} }
  var t0 = new Date().getTime();
  var sheet = getSheetByName('商品マスタ');
  if (!sheet) {
    appendPerfLog('apiGetCurrentStockList', new Date().getTime() - t0, '');
    return [];
  }
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iSingle = h.indexOf('単品JAN');
  var iName = h.indexOf('商品名');
  if (iSingle < 0) {
    appendPerfLog('apiGetCurrentStockList', new Date().getTime() - t0, '');
    return [];
  }
  var janList = [];
  for (var r = 1; r < data.length; r++) {
    var j = (data[r][iSingle] || '').toString().trim();
    if (j) janList.push(j);
  }
  var stockBatch = getCurrentStockBatch(janList);
  var locationList = getMasterList('場所マスタ') || [];
  var cache = getInventorySheetCache();
  var lastCountDate = getLastCountDateFromCache(cache);
  var stockBatchByLoc = getCurrentStockBatchByLocation(janList, locationList, cache, lastCountDate);
  var list = [];
  for (var r = 1; r < data.length; r++) {
    var jan = (data[r][iSingle] || '').toString().trim();
    if (!jan) continue;
    var stock = (stockBatch[jan] !== undefined) ? stockBatch[jan] : getCurrentStockByJan(jan);
    var byLocation = [];
    for (var L = 0; L < locationList.length; L++) {
      var loc = (locationList[L] || '').toString().trim();
      var qty = (stockBatchByLoc[jan] && stockBatchByLoc[jan][loc] !== undefined) ? stockBatchByLoc[jan][loc] : 0;
      byLocation.push({ location: loc, qty: qty });
    }
    list.push({ singleJAN: jan, productName: (data[r][iName] || '').toString(), stock: stock, byLocation: byLocation });
  }
  var result = { list: list, locationList: locationList };
  try { appCache.put('currentStockList', JSON.stringify(result), CACHE_TTL_SEC); } catch (e) {}
  appendPerfLog('apiGetCurrentStockList', new Date().getTime() - t0, '');
  return result;
}

function apiShipping(singleJAN, productName, locationFrom) {
  var t0 = new Date().getTime();
  var tantou = apiGetTantou();
  var result = appendShippingRow(singleJAN, productName, locationFrom, tantou);
  if (result.ok) {
    appendLog(getCurrentUserEmail(), tantou, '出荷', '登録', singleJAN);
    invalidateListCachesOnWrite();
  }
  appendPerfLog('apiShipping', new Date().getTime() - t0, tantou || '');
  return result;
}

function apiCount(singleJAN, productName, cost, boxInCount, boxCount, loose, quantity, location, sessionId) {
  var t0 = new Date().getTime();
  var tantou = apiGetTantou();
  var result = appendCountRow(singleJAN, productName, cost, boxInCount, boxCount, loose, quantity, location, tantou, sessionId);
  if (result.ok) {
    appendLog(getCurrentUserEmail(), tantou, '棚卸', '登録', singleJAN);
    invalidateListCachesOnWrite();
  }
  appendPerfLog('apiCount', new Date().getTime() - t0, tantou || '');
  return result;
}

function apiReceiving(singleJAN, productName, cost, boxInCount, boxCount, loose, quantity, location, unitCost, reason) {
  var t0 = new Date().getTime();
  try {
    var tantou = apiGetTantou();
    var jan = (singleJAN || '').toString().trim();
    var p = findProductByJan(jan, jan);
    if (!p) {
      var addResult = appendProductRow(jan, jan, productName || jan, cost, boxInCount);
      if (!addResult || !addResult.ok) {
        appendPerfLog('apiReceiving', new Date().getTime() - t0, '');
        return addResult || { ok: false, error: '商品マスタ登録に失敗しました' };
      }
    }
    var currentQtyForAvg = null;
    if (unitCost && quantity) {
      var cache = getInventorySheetCache();
      var lastCountDate = getLastCountDateFromCache(cache);
      currentQtyForAvg = getCurrentStockByJanFromCache(jan, cache, lastCountDate);
    }
    var result = appendReceivingRow(singleJAN, productName, cost, boxInCount, boxCount, loose, quantity, location, tantou, unitCost, reason);
    if (result.ok) {
      try {
        if (currentQtyForAvg !== null) updateAverageCost(singleJAN, quantity, unitCost, currentQtyForAvg);
      } catch (e) {}
      try {
        appendLog(getCurrentUserEmail(), tantou, '入荷', '登録', singleJAN);
      } catch (e) {}
      invalidateListCachesOnWrite();
    }
    appendPerfLog('apiReceiving', new Date().getTime() - t0, tantou || '');
    return result;
  } catch (err) {
    appendPerfLog('apiReceiving', new Date().getTime() - t0, '');
    return { ok: false, error: (err && (err.message || String(err))) || '登録に失敗しました' };
  }
}

function apiMove(singleJAN, productName, fromLocation, toLocation, qty) {
  var t0 = new Date().getTime();
  var tantou = apiGetTantou();
  var result = appendMoveRow(singleJAN, productName, fromLocation, toLocation, qty, tantou);
  if (result.ok) {
    appendLog(getCurrentUserEmail(), tantou, '在庫移動', '登録', singleJAN);
    invalidateListCachesOnWrite();
  }
  appendPerfLog('apiMove', new Date().getTime() - t0, tantou || '');
  return result;
}

function apiProductAdd(singleJAN, boxJAN, productName, cost, boxInCount) {
  var t0 = new Date().getTime();
  var result = appendProductRow(singleJAN, boxJAN, productName, cost, boxInCount);
  if (result.ok) {
    appendLog(getCurrentUserEmail(), apiGetTantou(), '商品情報', '新規登録', singleJAN);
    invalidateListCachesOnWrite();
  }
  appendPerfLog('apiProductAdd', new Date().getTime() - t0, apiGetTantou() || '');
  return result;
}

/**
 * 棚卸確定後修正を棚卸スキャンリストに記録する。
 * adjustType: '加算'（在庫増）または '減算'（在庫減）。
 * qty は正の整数で渡す（符号はサーバー側で付与）。
 */
function apiCountAdjust(singleJAN, productName, adjustType, qty, location, memo) {
  var t0 = new Date().getTime();
  try {
    var tantou = apiGetTantou();
    var result = appendCountAdjustRow(singleJAN, productName, adjustType, qty, location, tantou, memo);
    if (result.ok) {
      appendLog(getCurrentUserEmail(), tantou, '棚卸確定後修正', adjustType, singleJAN + (memo ? ' ' + memo : ''));
      invalidateListCachesOnWrite();
    }
    appendPerfLog('apiCountAdjust', new Date().getTime() - t0, tantou || '');
    return result;
  } catch (e) {
    appendPerfLog('apiCountAdjust', new Date().getTime() - t0, '');
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}
