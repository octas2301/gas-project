// ==========================================
// 在庫管理アプリ - スプレッドシート読書き
// ==========================================

/**
 * シート名でシートを取得。無ければ null。
 */
function getSheetByName(name) {
  var ss = getSpreadsheet();
  return ss.getSheetByName(name);
}

/**
 * ヘッダー行（1行目）から列名のインデックス（0始まり）を取得。
 */
function getColumnIndex(sheet, columnName) {
  var data = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < data.length; i++) {
    if ((data[i] || '').toString().trim() === columnName) return i;
  }
  return -1;
}

/**
 * 商品マスタ: 単品JAN または 箱JAN で行を検索。見つかれば行データ（オブジェクト）、無ければ null。
 * 同じJANで複数行ある場合は先頭を返す。
 */
function findProductByJan(singleJan, boxJan) {
  var sheet = getSheetByName('商品マスタ');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var h = data[0];
  var iSingle = h.indexOf('単品JAN');
  var iBox = h.indexOf('箱JAN');
  var iName = h.indexOf('商品名');
  var iCost = h.indexOf('商品原価');
  var iAvgCost = h.indexOf('平均原価');
  var iBoxQty = h.indexOf('箱入数');
  if (iSingle < 0) return null;
  var searchSingle = (singleJan || '').toString().trim();
  var searchBox = (boxJan || '').toString().trim();
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var s = (row[iSingle] || '').toString().trim();
    var b = (iBox >= 0 ? (row[iBox] || '').toString().trim() : '');
    if (searchSingle && s === searchSingle) {
      return { rowIndex: r + 1, singleJAN: s, boxJAN: b, productName: row[iName], cost: row[iCost], avgCost: row[iAvgCost], boxInCount: row[iBoxQty], values: row };
    }
    if (searchBox && b === searchBox) {
      return { rowIndex: r + 1, singleJAN: s, boxJAN: b, productName: row[iName], cost: row[iCost], avgCost: row[iAvgCost], boxInCount: row[iBoxQty], values: row };
    }
  }
  return null;
}

/**
 * 商品マスタに同じ単品JANまたは同じ箱JANが既に存在するか（二重登録チェック）。
 * 返り値: { duplicate: true/false, message: '...' }
 */
function checkProductDuplicate(singleJan, boxJan, excludeRowIndex) {
  var sheet = getSheetByName('商品マスタ');
  if (!sheet) return { duplicate: false };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { duplicate: false };
  var h = data[0];
  var iSingle = h.indexOf('単品JAN');
  var iBox = h.indexOf('箱JAN');
  var iName = h.indexOf('商品名');
  var sIn = (singleJan || '').toString().trim();
  var bIn = (boxJan || '').toString().trim();
  for (var r = 1; r < data.length; r++) {
    if (excludeRowIndex && r + 1 === excludeRowIndex) continue;
    var s = (data[r][iSingle] || '').toString().trim();
    var b = (iBox >= 0 ? (data[r][iBox] || '').toString().trim() : '');
    if (sIn && s === sIn) {
      return { duplicate: true, message: '既に登録済みです（単品JAN: ' + sIn + '）' };
    }
    if (bIn && b === bIn) {
      return { duplicate: true, message: '既に登録済みです（箱JAN: ' + bIn + '）' };
    }
  }
  return { duplicate: false };
}

/**
 * 商品マスタに1行追加。平均原価は引数で渡すか、商品原価で初期化。
 */
function appendProductRow(singleJAN, boxJAN, productName, cost, boxInCount) {
  var sheet = getSheetByName('商品マスタ');
  if (!sheet) return { ok: false, error: 'シートが見つかりません' };
  var dup = checkProductDuplicate(singleJAN, boxJAN);
  if (dup.duplicate) return { ok: false, error: dup.message };
  var avgCost = cost !== undefined && cost !== '' ? Number(cost) : 0;
  sheet.appendRow([
    (singleJAN || '').toString().trim(),
    (boxJAN || '').toString().trim(),
    (productName || '').toString().trim(),
    cost !== undefined && cost !== '' ? Number(cost) : 0,
    avgCost,
    boxInCount !== undefined && boxInCount !== '' ? Number(boxInCount) : 0
  ]);
  return { ok: true };
}

/**
 * 棚卸セッションを1件作成（スタート押下時）。返り値: { ok: true, sessionId: '...' } または { ok: false, error: '...' }
 */
function createCountSession() {
  var sheet = getSheetByName('棚卸セッション');
  if (!sheet) return { ok: false, error: '棚卸セッションシートが見つかりません' };
  var sessionId = Utilities.getUuid();
  var now = new Date();
  sheet.appendRow([sessionId, now, null]);
  return { ok: true, sessionId: sessionId };
}

/**
 * 未完了の棚卸セッションが存在すればその1件を返す。なければ null。
 * 返り値: { sessionId, startTime } または null
 */
function getActiveCountSession() {
  var sheet = getSheetByName('棚卸セッション');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var h = data[0];
  var iId = h.indexOf('セッションID');
  var iStart = h.indexOf('開始日時');
  var iEnd = h.indexOf('完了日時');
  if (iId < 0 || iStart < 0 || iEnd < 0) return null;
  for (var r = data.length - 1; r >= 1; r--) {
    var completed = data[r][iEnd];
    if (completed === null || completed === '' || (typeof completed === 'string' && completed.trim() === '')) {
      return { sessionId: (data[r][iId] || '').toString().trim(), startTime: data[r][iStart] };
    }
  }
  return null;
}

/**
 * 指定セッションを完了にする（完了日時を記録）。
 */
function completeCountSession(sessionId) {
  var sheet = getSheetByName('棚卸セッション');
  if (!sheet) return { ok: false, error: '棚卸セッションシートが見つかりません' };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: false, error: 'セッションがありません' };
  var h = data[0];
  var iId = h.indexOf('セッションID');
  var iEnd = h.indexOf('完了日時');
  if (iId < 0 || iEnd < 0) return { ok: false, error: '列が見つかりません' };
  var sid = (sessionId || '').toString().trim();
  if (!sid) return { ok: false, error: 'セッションIDがありません' };
  var now = new Date();
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iId] || '').toString().trim() === sid) {
      sheet.getRange(r + 1, iEnd + 1).setValue(now);
      return { ok: true };
    }
  }
  return { ok: false, error: '該当セッションが見つかりません' };
}

/**
 * 棚卸スキャンリストに1行追加。sessionId は省略可（棚卸期間外のときは空）。
 * 種別は常に「通常スキャン」で記録する。
 */
function appendCountRow(singleJAN, productName, cost, boxInCount, boxCount, loose, quantity, location, tantouName, sessionId) {
  var sheet = getSheetByName('棚卸スキャンリスト');
  if (!sheet) return { ok: false, error: 'シートが見つかりません' };
  var now = new Date();
  var yyyymm = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM');
  var totalCost = (quantity || 0) * (cost || 0);
  var sid = (sessionId !== undefined && sessionId !== null && (sessionId || '').toString().trim() !== '')
    ? (sessionId || '').toString().trim() : '';
  sheet.appendRow([
    Utilities.getUuid(),
    now,
    (singleJAN || '').toString().trim(),
    (productName || '').toString().trim(),
    cost !== undefined && cost !== '' ? Number(cost) : 0,
    boxInCount !== undefined && boxInCount !== '' ? Number(boxInCount) : 0,
    boxCount !== undefined && boxCount !== '' ? Number(boxCount) : 0,
    loose !== undefined && loose !== '' ? Number(loose) : 0,
    quantity !== undefined && quantity !== '' ? Number(quantity) : 0,
    (location || '').toString().trim(),
    (tantouName || '').toString().trim(),
    totalCost,
    yyyymm,
    sid,
    '通常スキャン',
    ''
  ]);
  return { ok: true };
}

/**
 * 棚卸確定後修正を棚卸スキャンリストに1行追加。
 * adjustType: '加算' または '減算'。数量は加算=正、減算=負で記録する。
 */
function appendCountAdjustRow(singleJAN, productName, adjustType, qty, location, tantouName, memo) {
  var sheet = getSheetByName('棚卸スキャンリスト');
  if (!sheet) return { ok: false, error: 'シートが見つかりません' };
  var now = new Date();
  var yyyymm = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM');
  var absQty = Math.abs(Number(qty) || 0);
  if (absQty <= 0) return { ok: false, error: '数量を入力してください' };
  var signedQty = (adjustType === '加算') ? absQty : -absQty;
  sheet.appendRow([
    Utilities.getUuid(),
    now,
    (singleJAN || '').toString().trim(),
    (productName || '').toString().trim(),
    0,
    0,
    0,
    0,
    signedQty,
    (location || '').toString().trim(),
    (tantouName || '').toString().trim(),
    0,
    yyyymm,
    '',
    '確定後修正',
    (memo || '').toString().trim()
  ]);
  return { ok: true };
}

/**
 * 入荷リストに1行追加。商品マスタの平均原価を更新するのは InventoryLogic 側で呼ぶ。
 */
function appendReceivingRow(singleJAN, productName, cost, boxInCount, boxCount, loose, quantity, location, tantouName, unitCost, reason) {
  var sheet = getSheetByName('入荷リスト');
  if (!sheet) return { ok: false, error: 'シートが見つかりません' };
  var now = new Date();
  var yyyymm = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM');
  var totalCost = (quantity || 0) * (cost || 0);
  sheet.appendRow([
    Utilities.getUuid(),
    now,
    (singleJAN || '').toString().trim(),
    (productName || '').toString().trim(),
    cost !== undefined && cost !== '' ? Number(cost) : 0,
    boxInCount !== undefined && boxInCount !== '' ? Number(boxInCount) : 0,
    boxCount !== undefined && boxCount !== '' ? Number(boxCount) : 0,
    loose !== undefined && loose !== '' ? Number(loose) : 0,
    quantity !== undefined && quantity !== '' ? Number(quantity) : 0,
    (location || '').toString().trim(),
    (tantouName || '').toString().trim(),
    totalCost,
    yyyymm,
    unitCost !== undefined && unitCost !== '' ? Number(unitCost) : 0,
    (reason || '').toString().trim()
  ]);
  return { ok: true };
}

/**
 * 出荷リストに1行追加（数量1）。
 */
function appendShippingRow(singleJAN, productName, locationFrom, tantouName) {
  var sheet = getSheetByName('出荷リスト');
  if (!sheet) return { ok: false, error: 'シートが見つかりません' };
  var now = new Date();
  var yyyymm = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM');
  sheet.appendRow([
    Utilities.getUuid(),
    now,
    (singleJAN || '').toString().trim(),
    (productName || '').toString().trim(),
    1,
    (locationFrom || '').toString().trim(),
    (tantouName || '').toString().trim(),
    yyyymm
  ]);
  return { ok: true };
}

/**
 * 在庫移動リストに1行追加。
 */
function appendMoveRow(singleJAN, productName, fromLocation, toLocation, qty, tantouName) {
  var sheet = getSheetByName('在庫移動リスト');
  if (!sheet) return { ok: false, error: 'シートが見つかりません' };
  var now = new Date();
  var moveDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
  sheet.appendRow([
    Utilities.getUuid(),
    now,
    (singleJAN || '').toString().trim(),
    (productName || '').toString().trim(),
    (fromLocation || '').toString().trim(),
    (toLocation || '').toString().trim(),
    qty !== undefined && qty !== '' ? Number(qty) : 0,
    (tantouName || '').toString().trim(),
    moveDate
  ]);
  return { ok: true };
}

/** パフォーマンスログ用シート名 */
var PERF_LOG_SHEET_NAME = 'パフォーマンスログ';

/**
 * パフォーマンスログに1行追加。APIの処理時間を記録する。失敗しても呼び出し元には影響しない。
 * @param {string} apiName API名（例: apiCount, apiGetVarianceList）
 * @param {number} elapsedMs 処理時間（ミリ秒）
 * @param {string} [tantou] 担当者名（省略可）
 */
function appendPerfLog(apiName, elapsedMs, tantou) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(PERF_LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(PERF_LOG_SHEET_NAME);
      sheet.getRange(1, 1, 1, 4).setValues([['日時', 'API名', '処理時間(ms)', '担当者']]);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    }
    sheet.appendRow([
      new Date(),
      String(apiName || ''),
      Math.round(Number(elapsedMs) || 0),
      String(tantou || '').trim()
    ]);
  } catch (e) {
    // ログ失敗は無視
  }
}

/**
 * 操作ログに1行追加。
 */
function appendLog(email, tantouName, screen, action, memo) {
  var sheet = getSheetByName('操作ログ');
  if (!sheet) return;
  sheet.appendRow([
    new Date(),
    (email || '').toString().trim(),
    (tantouName || '').toString().trim(),
    (screen || '').toString().trim(),
    (action || '').toString().trim(),
    (memo || '').toString().trim()
  ]);
}

/**
 * マスタ一覧取得（場所・理由・担当者）。キー列は1列目想定。
 */
function getMasterList(sheetName) {
  var sheet = getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var colName = (data[0][0] || '').toString().trim();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var v = (data[i][0] || '').toString().trim();
    if (v) list.push(v);
  }
  return list;
}

/**
 * 設定値の取得。キーで検索し値のセルを返す。
 */
function getSettingValue(key) {
  var sheet = getSheetByName('設定');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === key) {
      return (data[i][1] || '').toString().trim();
    }
  }
  return null;
}

/**
 * 日報宛先リスト（メールアドレス配列）。
 */
function getReportEmailList() {
  var sheet = getSheetByName('日報宛先');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var list = [];
  var col = 0;
  if (data[0][0] === 'メールアドレス') col = 0;
  for (var i = 1; i < data.length; i++) {
    var v = (data[i][col] || '').toString().trim();
    if (v && v.indexOf('@') > 0) list.push(v);
  }
  return list;
}
