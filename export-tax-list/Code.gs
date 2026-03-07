/**
 * 消費税輸出免除不適格用連絡一覧表 - 設定読み込み・フォルダ準備
 */

/**
 * 設定シート（A列: 項目名, B列: 値）から設定をオブジェクトで取得
 */
function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('設定');
  if (!sheet) {
    throw new Error('「設定」シートが見つかりません。');
  }
  var data = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 0; i < data.length; i++) {
    var key = data[i][0];
    var val = data[i][1];
    if (key && key.toString().trim() !== '') {
      config[key.toString().trim()] = val != null ? val.toString().trim() : '';
    }
  }
  return config;
}

/**
 * 監視フォルダ直下に「転記済み」「転記失敗」フォルダがなければ作成する
 */
function ensureFolders() {
  var config = getConfig();
  var folderId = config['監視フォルダID'];
  if (!folderId) {
    throw new Error('設定シートに「監視フォルダID」がありません。');
  }
  var parentFolder = DriveApp.getFolderById(folderId);
  var completedName = config['転記済みフォルダ名'] || '転記済み';
  var failedName = config['転記失敗フォルダ名'] || '転記失敗';

  var completedFolder = getOrCreateSubFolder(parentFolder, completedName);
  var failedFolder = getOrCreateSubFolder(parentFolder, failedName);

  return { completed: completedFolder, failed: failedFolder };
}

function getOrCreateSubFolder(parentFolder, name) {
  var iter = parentFolder.getFoldersByName(name);
  if (iter.hasNext()) {
    return iter.next();
  }
  return parentFolder.createFolder(name);
}

/** 一覧シートのヘッダー行（要件どおり＋6項目追加＋重複防止用元ファイル名＋保存先＋元ファイル名重複） */
var HEADER_ROW = ['No.', '海　外　客　先', '取引年月日', '輸出金額', 'Invoice No.', '品名', '貨物個数', '貨物重量', '通貨', '通貨レート', 'FOB価格', '要確認', '元ファイル名', '保存先', '元ファイル名重複'];

/**
 * 西暦年（例: 2025）のシートを取得。なければ作成する。
 */
function getOrCreateYearSheet(ss, year) {
  var sheetName = year + '年';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * シートの1行目にヘッダーがない場合だけ書き込む（既存なら上書きしない）
 */
function ensureHeader(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol >= HEADER_ROW.length) {
    var firstCell = sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0];
    var hasHeader = firstCell[0] === HEADER_ROW[0] && firstCell[1] === HEADER_ROW[1];
    if (hasHeader) return;
  }
  sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
}

/**
 * 監視フォルダ直下のPDFファイル一覧を取得（サブフォルダはのぞく）
 */
function getPdfFilesInFolder(folder) {
  var mimeType = MimeType.PDF;
  var files = [];
  var iter = folder.getFilesByType(mimeType);
  while (iter.hasNext()) {
    files.push(iter.next());
  }
  return files;
}

/**
 * 1分後に自分自身を実行するトリガーを1つだけ作成
 */
function scheduleResumeInOneMinute() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'run') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('run')
    .timeBased()
    .after(60 * 1000)
    .create();
}

/**
 * 時間驅動トリガーを1つ作成（1分ごとに runScheduled を実行）
 * 実行時間帯が1枠も設定されていない場合はトリガーを登録せず終了。
 * 設定されている場合は既存の runScheduled トリガーを1つにまとめて登録する。
 */
function createScheduleTrigger() {
  var config;
  try {
    config = getConfig();
  } catch (e) {
    Logger.log('設定を読み込めませんでした。トリガーは登録しません。');
    return;
  }
  var slots = getRunTimeSlots(config);
  if (slots.length === 0) {
    Logger.log('実行時間帯が1つも設定されていません。トリガーを登録しません。設定シートに「実行時間帯1開始」「実行時間帯1終了」などを追加してください。');
    return;
  }
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runScheduled') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runScheduled')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('トリガーを登録しました: 1分ごとに runScheduled を実行（設定された時間帯の間のみ処理）');
}

/**
 * 設定から実行可能な時間帯を最大5枠取得する
 * 設定例: 実行時間帯1開始=1, 実行時間帯1終了=7, 実行時間帯2開始=10, 実行時間帯2終了=12, ...
 * 戻り値: [[1,7], [10,12], ...] の形式（開始時≤現在時<終了時 で判定）
 * 1枠も設定がない場合は [] を返す（トリガーは動くが run は実行しない）
 */
function getRunTimeSlots(config) {
  var slots = [];
  for (var i = 1; i <= 5; i++) {
    var startVal = config['実行時間帯' + i + '開始'];
    var endVal = config['実行時間帯' + i + '終了'];
    var start = parseInt(startVal, 10);
    var end = parseInt(endVal, 10);
    if (isNaN(start) || isNaN(end) || start < 0 || start > 23 || end < 0 || end > 23) continue;
    slots.push([start, end]);
  }
  return slots;
}

/**
 * 現在時刻が実行時間帯のいずれかに含まれるか
 */
function isWithinRunTimeSlots(hour, slots) {
  for (var i = 0; i < slots.length; i++) {
    if (hour >= slots[i][0] && hour < slots[i][1]) return true;
  }
  return false;
}

/**
 * 時間驅動トリガー用の入口（設定シートの時間帯5枠のいずれかの間だけ run を実行）
 * トリガーはこの関数を「1分ごと」で指定する。
 * 実行時間帯が1枠も設定されていない場合は何もせず return（テスト中はトリガーを有効にしても安全）
 */
function runScheduled() {
  var config;
  try {
    config = getConfig();
  } catch (e) {
    return;
  }
  var slots = getRunTimeSlots(config);
  if (slots.length === 0) return;
  var now = new Date();
  var hour = now.getHours();
  if (!isWithinRunTimeSlots(hour, slots)) return;
  run();
}

/**
 * ログシートを取得または作成し、1行追記する
 * 5列目「メール送信済み」は空欄で追加（サマリーメール送信時に「済」を入れる）
 */
function appendLog(ss, runAt, result, count, message) {
  var logSheet = ss.getSheetByName('ログ');
  if (!logSheet) logSheet = ss.insertSheet('ログ');
  var lastRow = logSheet.getLastRow();
  if (lastRow === 0) {
    logSheet.getRange(1, 1, 1, 5).setValues([['実行日時', '結果', '処理PDF数', 'メッセージ', 'メール送信済み']]);
    lastRow = 1;
  }
  var insertRow = lastRow + 1;
  logSheet.getRange(insertRow, 1, 1, 5).setValues([[runAt, result, count, message || '', '']]);
}

/**
 * 実行結果を通知メールアドレスに送信する（設定にある場合のみ）
 * extra: { processedCount, errorCount, remainingCount, estimatedEndTime } を渡すと本文に追記
 */
function sendResultEmail(config, runAt, result, count, message, extra) {
  var to = (config['通知メールアドレス'] || '').toString().trim();
  if (!to) return;
  var subject = '[消費税輸出免除一覧] ' + result + ' - ' + (runAt ? Utilities.formatDate(runAt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '');
  var body = '結果: ' + result + '\nメッセージ: ' + (message || '') + '\n\n実行日時: ' + (runAt ? Utilities.formatDate(runAt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') : '');
  if (result === '途中終了') {
    body += '\n※残りは1分後に自動で再実行され、全て完了した時点で「成功」のメールをお送りします。';
  }
  if (extra) {
    body += '\n\n--- 今回の実行 ---\n';
    if (extra.processedCount !== undefined) body += '①処理できた件数: ' + extra.processedCount + '\n';
    if (extra.errorCount !== undefined) body += '②エラーが起きた件数: ' + extra.errorCount + '\n';
    if (extra.remainingCount !== undefined) body += '③残りの件数: ' + extra.remainingCount + '\n';
    if (extra.estimatedEndTime) {
      var now = new Date();
      var diffMs = extra.estimatedEndTime.getTime() - now.getTime();
      var totalMinutes = Math.max(0, Math.round(diffMs / 60000));
      var timeStr = totalMinutes >= 60 ? (Math.floor(totalMinutes / 60) + '時間' + (totalMinutes % 60 > 0 ? (totalMinutes % 60) + '分' : '')) : (totalMinutes + '分');
      body += '④終了予測時間: ' + timeStr + '\n';
    }
  }
  try {
    MailApp.sendEmail(to, subject, body);
    Logger.log('メール送信完了: ' + to);
  } catch (e) {
    Logger.log('メール送信失敗: ' + e.message);
  }
}

/**
 * メイン処理の入口（手動実行・トリガー共通）
 */
function run() {
  var runAt = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config;
  try {
    config = getConfig();
  } catch (e) {
    appendLog(ss, runAt, 'エラー', 0, '設定読み込み失敗: ' + e.message);
    Logger.log('エラー: ' + e.message);
    return;
  }

  var maxRunSec = parseInt(config['最大実行時間秒'] || '240', 10) || 240;
  var maxRunMs = maxRunSec * 1000;
  var startTime = new Date().getTime();
  var GAS_LIMIT_MS = 6 * 60 * 1000;
  var CLEANUP_BUFFER_MS = 90 * 1000;
  var stopAtMs = Math.min(maxRunMs, GAS_LIMIT_MS - CLEANUP_BUFFER_MS);

  var folderId = config['監視フォルダID'];
  if (!folderId) {
    appendLog(ss, runAt, 'エラー', 0, '監視フォルダIDが設定されていません');
    Logger.log('エラー: 監視フォルダIDが設定されていません。');
    return;
  }

  var parentFolder, folders;
  try {
    parentFolder = DriveApp.getFolderById(folderId);
    folders = ensureFolders();
  } catch (e) {
    appendLog(ss, runAt, 'エラー', 0, 'フォルダ取得失敗: ' + e.message);
    Logger.log('エラー: ' + e.message);
    return;
  }

  var year = config['処理対象年'] ? parseInt(config['処理対象年'], 10) : new Date().getFullYear();
  var sheet = getOrCreateYearSheet(ss, year);
  ensureHeader(sheet);

  var pdfs = getPdfFilesInFolder(parentFolder);
  if (pdfs.length === 0) {
    appendLog(ss, runAt, '対象なし', 0, '処理対象のPDFはありません');
    Logger.log('処理対象のPDFはありません。');
    return;
  }

  Logger.log('未処理PDF: ' + pdfs.length + ' 件');

  var processedCount = 0;
  var errorCount = 0;

  for (var i = 0; i < pdfs.length; i++) {
    if (new Date().getTime() - startTime > stopAtMs) {
      Logger.log('制限時間のため停止。1分後に再実行します。');
      scheduleResumeInOneMinute();
      var remainingPdfs = getPdfFilesInFolder(parentFolder).length;
      appendLog(ss, runAt, '途中終了', pdfs.length, '制限時間のため停止。1分後に再実行（処理済み' + processedCount + '／エラー' + errorCount + '／残り' + remainingPdfs + '）');
      return;
    }

    var file = pdfs[i];
    var res = processOnePdf(file, sheet, folders, config, startTime, stopAtMs);
    if (res === 'timeout') {
      Logger.log('制限時間のため停止。1分後に再実行します。');
      scheduleResumeInOneMinute();
      var remainingPdfs = getPdfFilesInFolder(parentFolder).length;
      appendLog(ss, runAt, '途中終了', pdfs.length, '制限時間のため停止。1分後に再実行（処理済み' + processedCount + '／エラー' + errorCount + '／残り' + remainingPdfs + '）');
      return;
    }
    if (res === 'completed') processedCount++; else errorCount++;
  }

  appendLog(ss, runAt, '成功', pdfs.length, '全件処理しました（処理済み' + processedCount + '／エラー' + errorCount + '）');
  Logger.log('今回の実行で全件処理しました。');
}

/**
 * PDF 1件の処理：読み取り → 転記 → 転記済み or 転記失敗へ移動
 * startTime, stopAtMs を渡すと、残りが stopAt を切っている場合は処理せず 'timeout' を返す（強制終了・メール未送信を防ぐ）
 */
function processOnePdf(file, sheet, folders, config, startTime, stopAtMs) {
  if (startTime != null && stopAtMs != null) {
    var remainMs = stopAtMs - (new Date().getTime() - startTime);
    if (remainMs < 90 * 1000) return 'timeout';
  }
  var fileName = file.getName();
  var apiKey = (config['GeminiAPIキー'] || '').toString().trim();
  if (!apiKey) {
    Logger.log('転記失敗(理由): GeminiAPIキーが設定されていません: ' + fileName);
    moveToFolder(file, folders.failed);
    return 'failed';
  }

  var blob = file.getBlob();
  var numChecks = parseInt(config['複数AI検証回数'] || '0', 10) || 0;
  var extracted;
  try {
    if (numChecks >= 2) {
      try {
        extracted = extractFromPdfWithGeminiMultiple(blob, apiKey, numChecks);
      } catch (eMulti) {
        Logger.log('複数回抽出が失敗したため1回のみで再試行: ' + fileName);
        extracted = extractFromPdfWithGemini(blob, apiKey);
      }
    } else {
      extracted = extractFromPdfWithGemini(blob, apiKey);
    }
  } catch (e) {
    Logger.log('転記失敗(理由): 読み取りエラー: ' + fileName + ' - ' + e.message);
    moveToFolder(file, folders.failed);
    return 'failed';
  }

  if (!extracted || extracted.length === 0) {
    Logger.log('転記失敗(理由): 抽出件数が0件でした: ' + fileName);
    moveToFolder(file, folders.failed);
    return 'failed';
  }

  var anyWritten = false;
  var firstRowForThisFile = null;
  for (var i = 0; i < extracted.length; i++) {
    var r = extracted[i];
    var invoiceNo = (r.申告番号 != null ? r.申告番号 : '').toString().trim();
    if (isDuplicate(sheet, invoiceNo, fileName)) continue;

    var nextRow = sheet.getLastRow() + 1;
    if (firstRowForThisFile === null) firstRowForThisFile = nextRow;
    var sameFileAsRow = (firstRowForThisFile === nextRow) ? null : firstRowForThisFile;

    var needCheck = (r.仕向人 === '' || r.申告年月日 === '' || r.申告番号 === '' || r.申告価格 === null || r.申告価格 === '');
    var completedLabel = (config['転記済みフォルダ名'] || '').toString().trim() || '転記済み';
    appendRecord(sheet, r, fileName, needCheck, completedLabel, sameFileAsRow);
    anyWritten = true;
  }

  if (anyWritten) {
    moveToFolder(file, folders.completed);
    return 'completed';
  }
  // 全件重複＝既にシートにデータがあるため転記済み扱いにする（転記失敗にすると件数不整合・「転記されているのに転記失敗」になる）
  Logger.log('全件重複のため新規転記は0件ですが、既存データのため転記済みへ移動: ' + fileName);
  moveToFolder(file, folders.completed);
  return 'completed';
}

function moveToFolder(file, folder) {
  folder.addFile(file);
  file.getParents().next().removeFile(file);
}

/**
 * シートで重複判定（申告番号 or 元ファイル名が既に存在するか）
 * Invoice No.は5列目（E列）、元ファイル名は13列目（M列）
 */
function isDuplicate(sheet, invoiceNo, fileName) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var colE = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  var colM = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  for (var i = 0; i < colE.length; i++) {
    if ((colE[i][0] != null && colE[i][0].toString().trim() === invoiceNo) ||
        (colM[i][0] != null && colM[i][0].toString().trim() === fileName)) {
      return true;
    }
  }
  return false;
}

/**
 * 1件分のレコードをシート末尾に追加（No. は行番号-1で連番・要確認フラグ・元ファイル名・保存先・元ファイル名重複）
 * 取引年月日は YYYY/MM/DD 形式なら日付型で格納
 * sameFileAsRow: 同じPDFの2行目以降なら、先頭行番号を渡すと「○行目と同じ元ファイル名」を記入
 */
function appendRecord(sheet, r, fileName, needCheck, saveDestination, sameFileAsRow) {
  var lastRow = sheet.getLastRow();
  var insertRow = lastRow + 1;
  var no = insertRow - 1;
  var consignee = (r.仕向人 != null ? r.仕向人 : '').toString().trim();
  var dateStr = (r.申告年月日 != null ? r.申告年月日 : '').toString().trim();
  var dateVal = dateStr;
  if (dateStr && /^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(dateStr.replace(/\s/g, ''))) {
    var parts = dateStr.replace(/\s/g, '').split(/[\/\-]/);
    dateVal = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  var price = r.申告価格;
  if (price !== null && price !== '') {
    if (typeof price === 'string') price = price.replace(/[¥,]/g, '');
    price = parseInt(price, 10);
  }
  if (isNaN(price)) price = '';
  var invoiceNo = (r.申告番号 != null ? r.申告番号 : '').toString().trim();

  var productName = (r.品名 != null ? r.品名 : '').toString().trim();
  var cargoQty = r.貨物個数;
  if (cargoQty !== null && cargoQty !== '') {
    cargoQty = parseInt(cargoQty, 10);
    if (isNaN(cargoQty)) cargoQty = '';
  } else {
    cargoQty = '';
  }
  var cargoWeight = r.貨物重量;
  if (cargoWeight !== null && cargoWeight !== '') {
    cargoWeight = parseFloat(cargoWeight);
    if (isNaN(cargoWeight)) cargoWeight = '';
  } else {
    cargoWeight = '';
  }
  var currency = (r.通貨 != null ? r.通貨 : '').toString().trim();
  var currencyRate = r.通貨レート;
  if (currencyRate !== null && currencyRate !== '') {
    currencyRate = parseFloat(currencyRate);
    if (isNaN(currencyRate)) currencyRate = '';
  } else {
    currencyRate = '';
  }
  var fobPrice = r.FOB価格;
  if (fobPrice !== null && fobPrice !== '') {
    fobPrice = parseFloat(fobPrice);
    if (isNaN(fobPrice)) fobPrice = '';
  } else {
    fobPrice = '';
  }

  var checkFlag = needCheck ? '要確認' : '';
  var dest = (saveDestination != null && saveDestination !== '') ? String(saveDestination) : '転記済み';
  var dupNote = (sameFileAsRow != null && sameFileAsRow > 0) ? (sameFileAsRow + '行目と同じ元ファイル名') : '';
  var row = [no, consignee, dateVal, price, invoiceNo, productName, cargoQty, cargoWeight, currency, currencyRate, fobPrice, checkFlag, fileName, dest, dupNote];
  sheet.getRange(insertRow, 1, 1, HEADER_ROW.length).setValues([row]);
  if (dateVal instanceof Date) {
    sheet.getRange(insertRow, 3).setNumberFormat('yyyy/mm/dd');
  }
  if (cargoWeight !== '' && typeof cargoWeight === 'number') {
    sheet.getRange(insertRow, 8).setNumberFormat('0.0');
  }
  if (currencyRate !== '' && typeof currencyRate === 'number') {
    sheet.getRange(insertRow, 10).setNumberFormat('0.00');
  }
  if (fobPrice !== '' && typeof fobPrice === 'number') {
    sheet.getRange(insertRow, 11).setNumberFormat('0.00');
  }
}

/**
 * Gemini 用のリクエストURLとペイロードを組み立てる（並列呼び出し用）
 */
function buildGeminiRequest(blob, apiKey) {
  var config = getConfig();
  var modelName = (config['Geminiモデル名'] || '').toString().trim() || 'gemini-3-flash-preview';
  var base64 = Utilities.base64Encode(blob.getBytes());
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + modelName + ':generateContent?key=' + encodeURIComponent(apiKey);
  var payload = {
    contents: [{
      parts: [
        { inline_data: { mime_type: 'application/pdf', data: base64 } },
        { text: 'このPDFは輸出許可書です。次の項目を抽出し、JSON配列で返してください。複数ページの場合はページごとに1件ずつオブジェクトを入れてください。\n項目:\n- 仕向人\n- 申告年月日(YYYY/MM/DD形式)\n- 申告価格(数値のみ・カンマ・円記号なし)\n- 申告番号\n- 品名(そのまま文字列)\n- 貨物個数(整数、単位「個」は除く)\n- 貨物重量(小数点以下を含む数値、単位「KGM」は除く。例:1.6)\n- 通貨(通貨レート欄の通貨コード部分、例: THB, USD)\n- 通貨レート(小数点第2位まで含む数値。例:4.45)\n- FOB価格(小数点第2位まで含む数値。例:1029.00)\n形式: [{"仕向人":"...","申告年月日":"...","申告価格":数値,"申告番号":"...","品名":"...","貨物個数":整数,"貨物重量":小数,"通貨":"...","通貨レート":小数,"FOB価格":小数}, ...]' }
      ]
    }],
    generationConfig: { response_mime_type: 'application/json', temperature: 0.1 }
  };
  return { url: url, payload: JSON.stringify(payload) };
}

/**
 * Gemini の応答本文をパースしてレコード配列を返す
 */
function parseGeminiResponseBody(body) {
  var json = JSON.parse(body);
  var text = (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) ? json.candidates[0].content.parts[0].text : '';
  if (!text || text.trim() === '') throw new Error('抽出結果が空です');
  text = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');
  var parsed = JSON.parse(text);
  if (!Array.isArray(parsed)) parsed = [parsed];
  return parsed;
}

/**
 * Gemini API で PDF から 仕向人・申告年月日・申告価格・申告番号 を抽出（1回・リトライあり）
 */
function extractFromPdfWithGemini(blob, apiKey) {
  var req = buildGeminiRequest(blob, apiKey);
  var options = { method: 'post', contentType: 'application/json', payload: req.payload, muteHttpExceptions: true };
  var maxRetries = 3;
  var res, code, body;
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    res = UrlFetchApp.fetch(req.url, options);
    code = res.getResponseCode();
    body = res.getContentText();
    if (code === 200) break;
    if ((code === 503 || code === 429) && attempt < maxRetries) {
      var waitSec = (code === 429) ? 60 : 2;
      Logger.log('Gemini API ' + code + ' (試行 ' + attempt + '/' + maxRetries + ')。' + waitSec + '秒後に再試行します。');
      Utilities.sleep(waitSec * 1000);
    } else {
      throw new Error('Gemini API エラー: ' + code + ' ' + body);
    }
  }
  return parseGeminiResponseBody(body);
}

/**
 * 複数回 Gemini を並列呼び出しし、多数決で1件の結果にまとめる（1回分の時間で完了）
 */
function extractFromPdfWithGeminiMultiple(blob, apiKey, numCalls) {
  var req = buildGeminiRequest(blob, apiKey);
  var requests = [];
  for (var c = 0; c < numCalls; c++) {
    requests.push({
      url: req.url,
      method: 'post',
      contentType: 'application/json',
      payload: req.payload,
      muteHttpExceptions: true
    });
  }
  var responses = UrlFetchApp.fetchAll(requests);
  var allResults = [];
  for (var i = 0; i < responses.length; i++) {
    if (responses[i].getResponseCode() !== 200) continue;
    try {
      var one = parseGeminiResponseBody(responses[i].getContentText());
      if (one && one.length > 0) allResults.push(one);
    } catch (e) {}
  }
  if (allResults.length === 0) throw new Error('複数回抽出いずれも結果なし');
  return mergeWithMajorityVote(allResults);
}

/**
 * 複数回の抽出結果をレコードごとに多数決して1つの配列にまとめる
 */
function mergeWithMajorityVote(allResults) {
  var n = allResults[0].length;
  var merged = [];
  for (var i = 0; i < n; i++) {
    var consigneeVals = [], dateVals = [], priceVals = [], invoiceVals = [];
    var productNameVals = [], cargoQtyVals = [], cargoWeightVals = [];
    var currencyVals = [], currencyRateVals = [], fobPriceVals = [];
    for (var j = 0; j < allResults.length; j++) {
      var r = allResults[j][i];
      if (!r) continue;
      consigneeVals.push((r.仕向人 != null ? r.仕向人 : '').toString().trim());
      dateVals.push((r.申告年月日 != null ? r.申告年月日 : '').toString().trim());
      var p = r.申告価格;
      if (p !== null && p !== undefined && p !== '') {
        if (typeof p === 'string') p = p.replace(/[¥,]/g, '');
        p = parseInt(p, 10);
      }
      priceVals.push(isNaN(p) ? null : p);
      invoiceVals.push((r.申告番号 != null ? r.申告番号 : '').toString().trim());

      productNameVals.push((r.品名 != null ? r.品名 : '').toString().trim());
      var qty = r.貨物個数;
      if (qty !== null && qty !== undefined && qty !== '') qty = parseInt(qty, 10);
      cargoQtyVals.push(isNaN(qty) ? null : qty);
      var wt = r.貨物重量;
      if (wt !== null && wt !== undefined && wt !== '') wt = parseFloat(wt);
      cargoWeightVals.push(isNaN(wt) ? null : wt);
      currencyVals.push((r.通貨 != null ? r.通貨 : '').toString().trim());
      var rate = r.通貨レート;
      if (rate !== null && rate !== undefined && rate !== '') rate = parseFloat(rate);
      currencyRateVals.push(isNaN(rate) ? null : rate);
      var fob = r.FOB価格;
      if (fob !== null && fob !== undefined && fob !== '') fob = parseFloat(fob);
      fobPriceVals.push(isNaN(fob) ? null : fob);
    }
    merged.push({
      仕向人: majorityVoteString(consigneeVals),
      申告年月日: majorityVoteString(dateVals),
      申告価格: majorityVoteNumber(priceVals),
      申告番号: majorityVoteString(invoiceVals),
      品名: majorityVoteString(productNameVals),
      貨物個数: majorityVoteNumber(cargoQtyVals),
      貨物重量: majorityVoteNumber(cargoWeightVals),
      通貨: majorityVoteString(currencyVals),
      通貨レート: majorityVoteNumber(currencyRateVals),
      FOB価格: majorityVoteNumber(fobPriceVals)
    });
  }
  return merged;
}

function majorityVoteString(vals) {
  var counts = {};
  var first = '';
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i] != null ? vals[i].toString().trim() : '';
    if (first === '' && v !== '') first = v;
    if (v === '') continue;
    counts[v] = (counts[v] || 0) + 1;
  }
  var maxCount = 0, result = first;
  for (var k in counts) {
    if (counts[k] > maxCount) { maxCount = counts[k]; result = k; }
  }
  return result;
}

function majorityVoteNumber(vals) {
  var nums = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i] !== null && vals[i] !== undefined && vals[i] !== '') {
      var n = parseFloat(vals[i]);
      if (!isNaN(n)) nums.push(n);
    }
  }
  if (nums.length === 0) return null;
  var counts = {};
  for (var j = 0; j < nums.length; j++) {
    var x = nums[j].toString();
    counts[x] = (counts[x] || 0) + 1;
  }
  var maxCount = 0, result = nums[0];
  for (var k in counts) {
    if (counts[k] > maxCount) { maxCount = counts[k]; result = parseFloat(k); }
  }
  return result;
}

/**
 * 動作確認用：設定読み込みとフォルダ作成を実行
 */
function testSetup() {
  var config = getConfig();
  Logger.log(config);
  var folders = ensureFolders();
  Logger.log('転記済みフォルダ: ' + folders.completed.getId());
  Logger.log('転記失敗フォルダ: ' + folders.failed.getId());
}

/**
 * 動作確認用：今年のシートとヘッダーを用意
 */
function testSheetSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = new Date().getFullYear();
  var sheet = getOrCreateYearSheet(ss, year);
  ensureHeader(sheet);
  Logger.log(year + '年シートのヘッダーを設定しました。');
}

/**
 * メール送信の権限確認用（1回実行すると「権限を確認」が出る場合があります）
 */
function testSendEmail() {
  var config = getConfig();
  var to = (config['通知メールアドレス'] || '').toString().trim();
  if (!to) {
    Logger.log('設定シートに「通知メールアドレス」を設定してください。');
    return;
  }
  MailApp.sendEmail(to, '[テスト] 消費税輸出免除一覧', 'メール送信のテストです。届いていれば権限は有効です。');
  Logger.log('送信しました: ' + to);
}

/**
 * 30分ごとに呼び出され、未送信のログをまとめてメール送信する
 */
function sendSummaryEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config;
  try {
    config = getConfig();
  } catch (e) {
    Logger.log('設定読み込み失敗: ' + e.message);
    return;
  }

  var to = (config['通知メールアドレス'] || '').toString().trim();
  if (!to) {
    Logger.log('通知メールアドレスが設定されていません。');
    return;
  }

  var logSheet = ss.getSheetByName('ログ');
  if (!logSheet) {
    Logger.log('ログシートがありません。');
    return;
  }

  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('ログデータがありません。');
    return;
  }

  var data = logSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var unsentRows = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][4] !== '済') {
      unsentRows.push({ rowIndex: i + 2, runAt: data[i][0], result: data[i][1], count: data[i][2], message: data[i][3] });
    }
  }

  if (unsentRows.length === 0) {
    Logger.log('未送信のログはありません。');
    return;
  }

  var totalProcessed = 0, totalError = 0, lastRemaining = 0;
  var details = [];
  for (var j = 0; j < unsentRows.length; j++) {
    var row = unsentRows[j];
    var timeStr = row.runAt ? Utilities.formatDate(new Date(row.runAt), 'Asia/Tokyo', 'HH:mm:ss') : '';
    details.push('- ' + timeStr + ' ' + row.result + ': ' + row.message);
    var match = row.message.match(/処理済み(\d+)／エラー(\d+)/);
    if (match) {
      totalProcessed += parseInt(match[1], 10);
      totalError += parseInt(match[2], 10);
    }
    var remainMatch = row.message.match(/残り(\d+)/);
    if (remainMatch) lastRemaining = parseInt(remainMatch[1], 10);
  }

  var now = new Date();
  var subject = '[消費税輸出免除一覧] 30分サマリー - ' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var body = '【30分間の処理サマリー】\n\n';
  body += '実行回数: ' + unsentRows.length + '回\n';
  body += '処理PDF合計: ' + totalProcessed + '件\n';
  body += 'エラー合計: ' + totalError + '件\n';
  body += '残り: ' + lastRemaining + '件\n';
  if (totalProcessed > 0 && lastRemaining > 0) {
    var estimatedMin = Math.round(lastRemaining / totalProcessed * unsentRows.length * 5);
    if (estimatedMin >= 60) {
      body += '終了予測: 約' + Math.floor(estimatedMin / 60) + '時間' + (estimatedMin % 60 > 0 ? (estimatedMin % 60) + '分' : '') + '\n';
    } else {
      body += '終了予測: 約' + estimatedMin + '分\n';
    }
  }
  body += '\n--- 詳細 ---\n';
  body += details.join('\n');

  try {
    MailApp.sendEmail(to, subject, body);
    Logger.log('サマリーメール送信完了: ' + to);
    for (var k = 0; k < unsentRows.length; k++) {
      logSheet.getRange(unsentRows[k].rowIndex, 5).setValue('済');
    }
  } catch (e) {
    Logger.log('サマリーメール送信失敗: ' + e.message);
  }
}

/**
 * 30分ごとに sendSummaryEmail を実行するトリガーを登録する
 * 既存の sendSummaryEmail トリガーがあれば削除して再登録
 */
function createSummaryEmailTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendSummaryEmail') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('sendSummaryEmail')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('トリガーを登録しました: 30分ごとに sendSummaryEmail を実行');
}

/**
 * sendSummaryEmail トリガーを削除する
 */
function deleteSummaryEmailTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendSummaryEmail') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }
  Logger.log('sendSummaryEmail トリガーを ' + deleted + ' 件削除しました。');
}
