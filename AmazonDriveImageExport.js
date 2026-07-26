/**
 * Lv4 T2: Drive `04\02` の MAIN 1枚 → Cloudflare R2 Put（PoC）
 * 設計: docs/org/LV4_R2_IMAGE_PIPELINE_POC.md §7.1
 * 承認: docs/org/LV4_T2_IMPLEMENTATION_APPROVAL.md（T2のみ・T3以降は別）
 *
 * - 楽天03・Yahoo.js・マスタ・xlsm・全件ループは触らない
 * - 1実行あたり MAIN 1枚のみ
 * - Secret は Script Properties のみ（コードに書かない）
 *
 * Script Properties:
 *   AMAZON_DRIVE_R2_UPLOAD_ENABLED … 既定 false
 *   AMAZON_DRIVE_IMAGE_FOLDER_ID … 02 Folder ID（空なら下記デフォルト）
 *   AMAZON_DRIVE_R2_POC_SKU … 例: lifec-4560151300139-oya（必須・{SKU}.MAIN.jpg を探す）
 *   R2_ACCOUNT_ID / R2_ACCESS_KEY_ID / R2_SECRET_ACCESS_KEY
 *   R2_BUCKET … 既定 octas-amazon-imag
 *   R2_PUBLIC_BASE … 既定 https://pub-d974bd81c7d84f9bbc65f8479d3f85d4.r2.dev
 */

var AMAZON_DRIVE_R2_UPLOAD_PROP = 'AMAZON_DRIVE_R2_UPLOAD_ENABLED';
var AMAZON_DRIVE_IMAGE_FOLDER_ID_PROP = 'AMAZON_DRIVE_IMAGE_FOLDER_ID';
var AMAZON_DRIVE_R2_POC_SKU_PROP = 'AMAZON_DRIVE_R2_POC_SKU';
/** Drive 04\02 の Folder ID（docs DRIVE_04_FOLDER_IDS）。Property で上書き可 */
var AMAZON_DRIVE_IMAGE_FOLDER_ID_DEFAULT = '1T6_E6T-qd9whSF8Re8lyRVB2n-P4BM84';
var AMAZON_R2_BUCKET_DEFAULT = 'octas-amazon-imag';
var AMAZON_R2_PUBLIC_BASE_DEFAULT = 'https://pub-d974bd81c7d84f9bbc65f8479d3f85d4.r2.dev';
/** PoC: 1実行1枚に固定 */
var AMAZON_DRIVE_R2_POC_MAX_FILES = 1;

/**
 * メニュー 21-⑥: Drive→R2 MAIN 1枚 PoC
 */
function menuAmazonDriveR2UploadPoc() {
  var stepName = 'AmazonDriveR2UploadPoc';
  var functionName = 'menuAmazonDriveR2UploadPoc';
  var runId = 'R2T2_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') +
    '_' + String(Utilities.getUuid()).replace(/-/g, '').substring(0, 6);

  Logger.log('[' + stepName + '] runId=' + runId + ' functionName=' + functionName + ' state=PENDING');

  if (!getBoolScriptProperty_(AMAZON_DRIVE_R2_UPLOAD_PROP, false)) {
    var off = 'Drive→R2 PoC は無効です。Script Properties の ' + AMAZON_DRIVE_R2_UPLOAD_PROP +
      ' を true にしてください（既定は無効）。T3/ZIP量産は別承認です。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + off);
    try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var pocSku = String(props.getProperty(AMAZON_DRIVE_R2_POC_SKU_PROP) || '').trim();
  if (!pocSku) {
    var noSku = AMAZON_DRIVE_R2_POC_SKU_PROP + ' が未設定です。例: lifec-4560151300139-oya';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + noSku);
    try { SpreadsheetApp.getUi().alert(noSku); } catch (e1) {}
    return;
  }

  var folderId = String(props.getProperty(AMAZON_DRIVE_IMAGE_FOLDER_ID_PROP) || '').trim() ||
    AMAZON_DRIVE_IMAGE_FOLDER_ID_DEFAULT;
  var accountId = String(props.getProperty('R2_ACCOUNT_ID') || '').trim();
  var accessKey = String(props.getProperty('R2_ACCESS_KEY_ID') || '').trim();
  var secretKey = String(props.getProperty('R2_SECRET_ACCESS_KEY') || '').trim();
  var bucket = String(props.getProperty('R2_BUCKET') || '').trim() || AMAZON_R2_BUCKET_DEFAULT;
  var publicBase = String(props.getProperty('R2_PUBLIC_BASE') || '').trim() || AMAZON_R2_PUBLIC_BASE_DEFAULT;
  publicBase = publicBase.replace(/\/$/, '');

  if (!accountId || !accessKey || !secretKey) {
    var noAuth = 'R2_ACCOUNT_ID / R2_ACCESS_KEY_ID / R2_SECRET_ACCESS_KEY のいずれかが未設定です（値はログに出しません）。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED authMissing=true');
    try { SpreadsheetApp.getUi().alert(noAuth); } catch (e2) {}
    return;
  }

  var fileName = pocSku + '.MAIN.jpg';
  Logger.log('[' + stepName + '] runId=' + runId + ' state=RUNNING folderId=' + folderId +
    ' fileName=' + fileName + ' maxFiles=' + AMAZON_DRIVE_R2_POC_MAX_FILES);

  try {
    var ui = SpreadsheetApp.getUi();
    var conf = ui.alert(
      '21-⑥ Drive→R2（T2 PoC・1枚）',
      '対象: ' + fileName + '\n' +
        'Drive folderId: ' + folderId + '\n' +
        'R2 bucket: ' + bucket + '\n' +
        '公開URL予定: ' + publicBase + '/' + fileName + '\n\n' +
        '楽天フォルダ・マスタ・xlsmは触りません。\n実行しますか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (conf !== ui.Button.OK) {
      Logger.log('[' + stepName + '] runId=' + runId + ' state=CANCELLED');
      return;
    }
  } catch (eUi) {}

  var file = amazonDriveFindFileByName_(folderId, fileName);
  if (!file) {
    var nf = 'Drive 上に ' + fileName + ' が見つかりません（02配下を再帰検索）。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + nf);
    try { SpreadsheetApp.getUi().alert(nf); } catch (e3) {}
    return;
  }

  var blob = file.getBlob();
  var bytes = blob.getBytes();
  Logger.log('[' + stepName + '] runId=' + runId + ' state=RUNNING driveFileId=' + file.getId() +
    ' sizeBytes=' + (bytes ? bytes.length : 0));

  var putRes = amazonR2PutObject_(accountId, accessKey, secretKey, bucket, fileName, bytes, 'image/jpeg');
  var publicUrl = publicBase + '/' + fileName;
  Logger.log('[' + stepName + '] runId=' + runId + ' stepName=' + stepName +
    ' functionName=' + functionName + ' sku=' + pocSku +
    ' http=' + putRes.code + ' url=' + publicUrl +
    ' state=' + (putRes.ok ? 'DONE' : 'FAILED') +
    ' errSummary=' + (putRes.error || ''));

  if (!putRes.ok) {
    try {
      SpreadsheetApp.getUi().alert(
        'R2 Put 失敗',
        'http=' + putRes.code + '\n' + (putRes.error || '').substring(0, 400),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e4) {}
    return;
  }

  try {
    SpreadsheetApp.getUi().alert(
      'Drive→R2 PoC 完了',
      '成功\nrunId=' + runId + '\nSKU=' + pocSku +
        '\nURL=\n' + publicUrl +
        '\n\nブラウザで URL が 200 か確認してください。\n' +
        '終わったら ' + AMAZON_DRIVE_R2_UPLOAD_PROP + ' を false に戻してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e5) {}
}

/**
 * folderId 配下（1階層のサブフォルダ含む）からファイル名完全一致を探す。
 * @param {string} folderId
 * @param {string} fileName
 * @return {GoogleAppsScript.Drive.File|null}
 */
function amazonDriveFindFileByName_(folderId, fileName) {
  var root = DriveApp.getFolderById(folderId);
  var hit = amazonDriveFindInFolder_(root, fileName, 0);
  return hit;
}

/**
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @param {string} fileName
 * @param {number} depth
 * @return {GoogleAppsScript.Drive.File|null}
 */
function amazonDriveFindInFolder_(folder, fileName, depth) {
  if (depth > 3) return null;
  var files = folder.getFilesByName(fileName);
  if (files.hasNext()) return files.next();
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    var sub = subs.next();
    var found = amazonDriveFindInFolder_(sub, fileName, depth + 1);
    if (found) return found;
  }
  return null;
}

/**
 * Cloudflare R2 へ S3互換 PUT（SigV4・UNSIGNED-PAYLOAD）
 * @return {{ok:boolean, code:number, error:string}}
 */
function amazonR2PutObject_(accountId, accessKey, secretKey, bucket, objectKey, bytes, contentType) {
  var host = accountId + '.r2.cloudflarestorage.com';
  // path-style: /{bucket}/{key}
  var url = 'https://' + host + '/' + bucket + '/' + objectKey;

  var region = 'auto';
  var service = 's3';
  var now = new Date();
  var amzDate = Utilities.formatDate(now, 'GMT', "yyyyMMdd'T'HHmmss'Z'");
  var dateStamp = Utilities.formatDate(now, 'GMT', 'yyyyMMdd');
  var payloadHash = 'UNSIGNED-PAYLOAD';

  var canonicalUri = '/' + bucket + '/' + objectKey;
  var canonicalHeaders = 'content-type:' + contentType + '\n' +
    'host:' + host + '\n' +
    'x-amz-content-sha256:' + payloadHash + '\n' +
    'x-amz-date:' + amzDate + '\n';
  var signedHeaders = 'content-type;host;x-amz-content-sha256;x-amz-date';
  var canonicalRequest = [
    'PUT',
    canonicalUri,
    '',
    canonicalHeaders,
    signedHeaders,
    payloadHash
  ].join('\n');

  var algorithm = 'AWS4-HMAC-SHA256';
  var credentialScope = dateStamp + '/' + region + '/' + service + '/aws4_request';
  var stringToSign = [
    algorithm,
    amzDate,
    credentialScope,
    amazonR2Sha256Hex_(canonicalRequest)
  ].join('\n');

  var signingKey = amazonR2GetSignatureKey_(secretKey, dateStamp, region, service);
  var signature = amazonR2HmacHex_(signingKey, stringToSign);
  var authorization = algorithm + ' Credential=' + accessKey + '/' + credentialScope +
    ', SignedHeaders=' + signedHeaders + ', Signature=' + signature;

  var res = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: contentType,
    payload: bytes,
    muteHttpExceptions: true,
    headers: {
      'Authorization': authorization,
      'x-amz-content-sha256': payloadHash,
      'x-amz-date': amzDate
    }
  });
  var code = res.getResponseCode();
  var body = res.getContentText() || '';
  if (code >= 200 && code < 300) {
    return { ok: true, code: code, error: '' };
  }
  return {
    ok: false,
    code: code,
    error: body.substring(0, 500)
  };
}

/** @param {string} s @return {string} */
function amazonR2Sha256Hex_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return amazonR2BytesToHex_(raw);
}

/** @param {string} s @return {Byte[]} */
function amazonR2Utf8Bytes_(s) {
  return Utilities.newBlob(String(s)).getBytes();
}

/**
 * GAS は computeHmacSha256Signature(String, Byte[]) を受け付けない。
 * 常に (Byte[], Byte[]) で呼ぶ（第1引数=value、第2=key）。
 * @param {string|Byte[]} key
 * @param {string|Byte[]} data
 * @return {Byte[]}
 */
function amazonR2HmacBytes_(key, data) {
  var keyBytes = (typeof key === 'string') ? amazonR2Utf8Bytes_(key) : key;
  var dataBytes = (typeof data === 'string') ? amazonR2Utf8Bytes_(data) : data;
  return Utilities.computeHmacSha256Signature(dataBytes, keyBytes);
}

/** @param {Byte[]} key @param {string} data @return {string} */
function amazonR2HmacHex_(key, data) {
  return amazonR2BytesToHex_(amazonR2HmacBytes_(key, data));
}

/**
 * @param {string} key
 * @param {string} dateStamp
 * @param {string} region
 * @param {string} service
 * @return {Byte[]}
 */
function amazonR2GetSignatureKey_(key, dateStamp, region, service) {
  var kDate = amazonR2HmacBytes_('AWS4' + key, dateStamp);
  var kRegion = amazonR2HmacBytes_(kDate, region);
  var kService = amazonR2HmacBytes_(kRegion, service);
  return amazonR2HmacBytes_(kService, 'aws4_request');
}

/** @param {Byte[]} bytes @return {string} */
function amazonR2BytesToHex_(bytes) {
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i];
    if (b < 0) b += 256;
    var h = b.toString(16);
    if (h.length === 1) out += '0';
    out += h;
  }
  return out;
}

// ----- U4: Drive02 → R2 → マスタ URL（2026-07-26 承認） -----
var AMAZON_U4_URL_EMBED_PROP = 'AMAZON_U4_URL_EMBED_ENABLED';
var AMAZON_U4_MAX_SKUS_PROP = 'AMAZON_U4_MAX_SKUS';
var AMAZON_U4_SKU_LIST_PROP = 'AMAZON_U4_SKU_LIST';
var AMAZON_U4_MASTER_COL_MAIN_URL = 'Amazon MAIN URL';
var AMAZON_U4_MASTER_COL_PT_URL = 'Amazon PT URL';
var AMAZON_U4_MAX_SKUS_DEFAULT = 20;

/**
 * メニュー 21-⑦: 対象子SKUの MAIN（＋ONLY時PT）を R2 へ上げ、マスタに URL を書く。
 */
function menuAmazonU4UrlEmbed() {
  var stepName = 'AmazonU4UrlEmbed';
  var functionName = 'menuAmazonU4UrlEmbed';
  var runId = 'U4_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') +
    '_' + String(Utilities.getUuid()).replace(/-/g, '').substring(0, 6);
  Logger.log('[' + stepName + '] runId=' + runId + ' functionName=' + functionName + ' state=PENDING');

  if (!getBoolScriptProperty_(AMAZON_U4_URL_EMBED_PROP, false)) {
    var off = 'U4 は無効です。Script Properties の ' + AMAZON_U4_URL_EMBED_PROP + ' を true にしてください（既定は無効）。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + off);
    try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var folderId = String(props.getProperty(AMAZON_DRIVE_IMAGE_FOLDER_ID_PROP) || '').trim() ||
    AMAZON_DRIVE_IMAGE_FOLDER_ID_DEFAULT;
  var accountId = String(props.getProperty('R2_ACCOUNT_ID') || '').trim();
  var accessKey = String(props.getProperty('R2_ACCESS_KEY_ID') || '').trim();
  var secretKey = String(props.getProperty('R2_SECRET_ACCESS_KEY') || '').trim();
  var bucket = String(props.getProperty('R2_BUCKET') || '').trim() || AMAZON_R2_BUCKET_DEFAULT;
  var publicBase = String(props.getProperty('R2_PUBLIC_BASE') || '').trim() || AMAZON_R2_PUBLIC_BASE_DEFAULT;
  publicBase = publicBase.replace(/\/$/, '');
  var maxSkus = parseInt(String(props.getProperty(AMAZON_U4_MAX_SKUS_PROP) || ''), 10);
  if (!(maxSkus > 0)) maxSkus = AMAZON_U4_MAX_SKUS_DEFAULT;

  if (!accountId || !accessKey || !secretKey) {
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED authMissing=true');
    try {
      SpreadsheetApp.getUi().alert('R2_ACCOUNT_ID / R2_ACCESS_KEY_ID / R2_SECRET_ACCESS_KEY のいずれかが未設定です。');
    } catch (e1) {}
    return;
  }

  var targets = amazonU4CollectTargetSkus_(maxSkus);
  if (!targets.length) {
    var msg =
      '対象SKUがありません。\n' +
      '・マッチングsheetの子行に Amazon MAIN がある\n' +
      '・またはマスタに Amazon MAIN 参照がある子\n' +
      '・または Property ' + AMAZON_U4_SKU_LIST_PROP + '（カンマ区切り）';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED noTargets');
    try { SpreadsheetApp.getUi().alert(msg); } catch (e2) {}
    return;
  }

  try {
    var ui = SpreadsheetApp.getUi();
    var conf = ui.alert(
      '21-⑦ U4 Drive02→R2→マスタURL',
      '対象SKU数=' + targets.length + '（上限' + maxSkus + '）\n' +
        '例: ' + targets.slice(0, 3).join(', ') + (targets.length > 3 ? '…' : '') + '\n' +
        'マスタに「' + AMAZON_U4_MASTER_COL_MAIN_URL + '」等を書きます（在庫・JANは触りません）。\n' +
        'xlsm直編集・ZIP・楽天CSVはしません。\n実行しますか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (conf !== ui.Button.OK) {
      Logger.log('[' + stepName + '] runId=' + runId + ' state=CANCELLED');
      return;
    }
  } catch (eUi) {}

  Logger.log('[' + stepName + '] runId=' + runId + ' state=RUNNING n=' + targets.length);
  var summary = amazonU4EmbedUrls_(runId, targets, folderId, accountId, accessKey, secretKey, bucket, publicBase);
  Logger.log('[' + stepName + '] runId=' + runId + ' state=DONE ' + JSON.stringify(summary));

  try {
    SpreadsheetApp.getUi().alert(
      'U4 完了',
      'runId=' + runId +
        '\nMAIN成功=' + summary.mainOk +
        '\nPT成功=' + summary.ptOk +
        '\n失敗=' + summary.failed +
        '\nマスタ更新SKU=' + summary.masterUpdated +
        '\n' + (summary.message || '') +
        '\n\n終わったら ' + AMAZON_U4_URL_EMBED_PROP + ' を false に戻してください。\n' +
        '次: 21-①／D Amazon で GENERATED（Amazon URL優先）。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e3) {}
}

/**
 * U4 コア（確認・成功ダイアログなし）。Eコース用。
 * @return {{ok:boolean, runId?:string, summary?:Object, error?:string}}
 */
function amazonU4UrlEmbedSilent_() {
  var stepName = 'AmazonU4UrlEmbedSilent';
  var runId = 'U4_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') +
    '_' + String(Utilities.getUuid()).replace(/-/g, '').substring(0, 6);
  Logger.log('[' + stepName + '] runId=' + runId + ' state=PENDING');

  if (!getBoolScriptProperty_(AMAZON_U4_URL_EMBED_PROP, false)) {
    var off = 'U4 は無効です。Script Properties の ' + AMAZON_U4_URL_EMBED_PROP + ' を true にしてください（既定は無効）。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + off);
    return { ok: false, runId: runId, error: off };
  }

  var props = PropertiesService.getScriptProperties();
  var folderId = String(props.getProperty(AMAZON_DRIVE_IMAGE_FOLDER_ID_PROP) || '').trim() ||
    AMAZON_DRIVE_IMAGE_FOLDER_ID_DEFAULT;
  var accountId = String(props.getProperty('R2_ACCOUNT_ID') || '').trim();
  var accessKey = String(props.getProperty('R2_ACCESS_KEY_ID') || '').trim();
  var secretKey = String(props.getProperty('R2_SECRET_ACCESS_KEY') || '').trim();
  var bucket = String(props.getProperty('R2_BUCKET') || '').trim() || AMAZON_R2_BUCKET_DEFAULT;
  var publicBase = String(props.getProperty('R2_PUBLIC_BASE') || '').trim() || AMAZON_R2_PUBLIC_BASE_DEFAULT;
  publicBase = publicBase.replace(/\/$/, '');
  var maxSkus = parseInt(String(props.getProperty(AMAZON_U4_MAX_SKUS_PROP) || ''), 10);
  if (!(maxSkus > 0)) maxSkus = AMAZON_U4_MAX_SKUS_DEFAULT;

  if (!accountId || !accessKey || !secretKey) {
    var authErr = 'R2_ACCOUNT_ID / R2_ACCESS_KEY_ID / R2_SECRET_ACCESS_KEY のいずれかが未設定です。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED authMissing=true');
    return { ok: false, runId: runId, error: authErr };
  }

  var targets = amazonU4CollectTargetSkus_(maxSkus);
  if (!targets.length) {
    var msg =
      '対象SKUがありません。マッチングsheetの子に Amazon MAIN／マスタ Amazon MAIN 参照／' +
      AMAZON_U4_SKU_LIST_PROP + ' を確認してください。';
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED noTargets');
    return { ok: false, runId: runId, error: msg };
  }

  Logger.log('[' + stepName + '] runId=' + runId + ' state=RUNNING n=' + targets.length);
  try {
    var summary = amazonU4EmbedUrls_(runId, targets, folderId, accountId, accessKey, secretKey, bucket, publicBase);
    Logger.log('[' + stepName + '] runId=' + runId + ' state=DONE ' + JSON.stringify(summary));
    if (summary && Number(summary.failed) > 0 && !(Number(summary.mainOk) > 0)) {
      return {
        ok: false,
        runId: runId,
        summary: summary,
        error: 'U4失敗: MAIN成功=0 失敗=' + summary.failed + ' ' + (summary.message || '')
      };
    }
    return { ok: true, runId: runId, summary: summary };
  } catch (eRun) {
    var err = String((eRun && eRun.message) || eRun);
    Logger.log('[' + stepName + '] runId=' + runId + ' state=FAILED ' + err);
    return { ok: false, runId: runId, error: err };
  }
}

/**
 * @param {number} maxSkus
 * @return {string[]}
 */
function amazonU4CollectTargetSkus_(maxSkus) {
  var seen = {};
  var out = [];
  function add(sku) {
    sku = String(sku || '').trim();
    if (!sku || seen[sku]) return;
    if (out.length >= maxSkus) return;
    seen[sku] = true;
    out.push(sku);
  }

  var props = PropertiesService.getScriptProperties();
  var listProp = String(props.getProperty(AMAZON_U4_SKU_LIST_PROP) || '').trim();
  if (listProp) {
    listProp.split(/[,，\s]+/).forEach(function (s) { add(s); });
    if (out.length) return out;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var matrixName = (typeof SHEET_NAME_MATRIX !== 'undefined') ? SHEET_NAME_MATRIX : '★画像AIマッチング(操作用)';
  var mx = ss.getSheetByName(matrixName);
  if (mx && mx.getLastRow() >= 3) {
    var num = mx.getLastRow() - 2;
    var keys = mx.getRange(3, 1, num, 2).getValues();
    var mainF = mx.getRange(3, 76, num, 1).getFormulas();
    var i;
    for (i = 0; i < keys.length && out.length < maxSkus; i++) {
      var cCode = String(keys[i][1] || '').trim();
      if (!cCode) continue;
      if (String(mainF[i][0] || '').trim()) add(cCode);
    }
    if (out.length) return out;
  }

  var masterName = (typeof MASTER_SHEET_NAME !== 'undefined') ? MASTER_SHEET_NAME : '▼商品マスタ(人間作業用)';
  var master = ss.getSheetByName(masterName);
  if (!master) return out;
  var mValues = master.getDataRange().getValues();
  var headerRowIdx = getAnchorRowIndex(mValues);
  var colMap = getColumnIndexMap(mValues[headerRowIdx]);
  var idxC = colMap['子SKU'];
  var idxRef = colMap['Amazon MAIN 参照'];
  if (idxC === undefined) return out;
  var r;
  for (r = headerRowIdx + 1; r < mValues.length && out.length < maxSkus; r++) {
    var child = String(mValues[r][idxC] || '').trim();
    if (!child) continue;
    if (idxRef !== undefined && String(mValues[r][idxRef] || '').trim()) add(child);
  }
  return out;
}

/**
 * @return {{mainOk:number, ptOk:number, failed:number, masterUpdated:number, message:string}}
 */
function amazonU4EmbedUrls_(runId, skus, folderId, accountId, accessKey, secretKey, bucket, publicBase) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterName = (typeof MASTER_SHEET_NAME !== 'undefined') ? MASTER_SHEET_NAME : '▼商品マスタ(人間作業用)';
  var master = ss.getSheetByName(masterName);
  var mValues = master.getDataRange().getValues();
  var headerRowIdx = getAnchorRowIndex(mValues);
  var ensure = amazonU4EnsureMasterUrlColumns_(master, headerRowIdx);
  mValues = master.getDataRange().getValues();
  headerRowIdx = getAnchorRowIndex(mValues);
  var colMap = getColumnIndexMap(mValues[headerRowIdx]);
  var idxP = colMap['親SKU'];
  var idxC = colMap['子SKU'];
  var idxMode = colMap['Amazon画像モード'];
  var idxMainUrl = colMap[AMAZON_U4_MASTER_COL_MAIN_URL];
  var idxPtUrl = colMap[AMAZON_U4_MASTER_COL_PT_URL];
  var idxPtRef = colMap['Amazon PT 参照'];

  var mainOk = 0;
  var ptOk = 0;
  var failed = 0;
  var masterUpdated = 0;
  var notes = [];
  if (ensure.added.length) notes.push('列追加=' + ensure.added.join(','));

  var s;
  for (s = 0; s < skus.length; s++) {
    var sku = skus[s];
    var mainName = sku + '.MAIN.jpg';
    var file = amazonDriveFindFileByName_(folderId, mainName);
    if (!file) {
      failed++;
      notes.push(sku + ': MAINファイルなし');
      Logger.log('[U4] runId=' + runId + ' sku=' + sku + ' state=FAILED noMainFile');
      continue;
    }
    var bytes = file.getBlob().getBytes();
    var putMain = amazonR2PutObject_(accountId, accessKey, secretKey, bucket, mainName, bytes, 'image/jpeg');
    var mainUrl = publicBase + '/' + mainName;
    Logger.log(
      '[U4] runId=' + runId + ' sku=' + sku + ' http=' + putMain.code +
        ' url=' + mainUrl + ' state=' + (putMain.ok ? 'DONE' : 'FAILED')
    );
    if (!putMain.ok) {
      failed++;
      notes.push(sku + ' MAIN Put http=' + putMain.code);
      continue;
    }
    mainOk++;

    var ptUrls = [];
    var mRow = amazonU4FindMasterChildRow_(mValues, headerRowIdx, idxP, idxC, sku);
    var mode = '';
    if (mRow >= 0 && idxMode !== undefined) mode = String(mValues[mRow][idxMode] || '').trim().toUpperCase();
    if (mode.indexOf('ONLY') >= 0 && mRow >= 0 && idxPtRef !== undefined) {
      var ptRefs = String(mValues[mRow][idxPtRef] || '').split('|');
      var pi;
      var ptSeq = 0;
      for (pi = 0; pi < ptRefs.length; pi++) {
        var ptId = String(ptRefs[pi] || '').trim();
        if (!ptId) continue;
        ptSeq++;
        var ptName = sku + '.PT' + ('0' + ptSeq).slice(-2) + '.jpg';
        try {
          var ptFile = DriveApp.getFileById(ptId);
          var ptBytes = ptFile.getBlob().getBytes();
          var putPt = amazonR2PutObject_(accountId, accessKey, secretKey, bucket, ptName, ptBytes, 'image/jpeg');
          if (putPt.ok) {
            ptOk++;
            ptUrls.push(publicBase + '/' + ptName);
          } else {
            failed++;
            notes.push(sku + ' ' + ptName + ' http=' + putPt.code);
          }
        } catch (ePt) {
          failed++;
          notes.push(sku + ' PT: ' + ((ePt && ePt.message) || ePt));
        }
      }
    }

    if (mRow >= 0 && idxMainUrl !== undefined) {
      master.getRange(mRow + 1, idxMainUrl + 1).setValue(mainUrl);
      if (idxPtUrl !== undefined) master.getRange(mRow + 1, idxPtUrl + 1).setValue(ptUrls.join('|'));
      masterUpdated++;
    } else {
      notes.push(sku + ': マスタ子行なし（R2のみ成功）');
    }
  }

  // 子に書いた MAIN URL を、同じ親の親行（子SKU空）へ空欄時のみコピー（Lv4 Build / C1 用）
  var prop = amazonU4PropagateMainUrlToParents_(master, mValues, headerRowIdx, idxP, idxC, idxMainUrl);
  if (prop && prop.copied > 0) {
    masterUpdated += prop.copied;
    notes.push('親MAIN URLコピー=' + prop.copied);
    Logger.log('[U4] runId=' + runId + ' parentMainUrlCopied=' + prop.copied);
  }

  return {
    mainOk: mainOk,
    ptOk: ptOk,
    failed: failed,
    masterUpdated: masterUpdated,
    message: notes.slice(0, 10).join('\n') + (notes.length > 10 ? '\n…' : ''),
    parentMainUrlCopied: prop ? prop.copied : 0
  };
}

/**
 * 親行の Amazon MAIN URL が空のとき、同一親の子行から先頭の非空URLをコピーする。
 * @return {{copied:number}}
 */
function amazonU4PropagateMainUrlToParents_(master, mValues, headerRowIdx, idxP, idxC, idxMainUrl) {
  var out = { copied: 0 };
  if (idxP == null || idxC == null || idxMainUrl == null || !master) return out;
  // 子への setValue 直後は引数 mValues が古いことがあるので再読込
  mValues = master.getDataRange().getValues();
  var byParent = {};
  var r;
  for (r = headerRowIdx + 1; r < mValues.length; r++) {
    var p = String(mValues[r][idxP] != null ? mValues[r][idxP] : '').trim();
    if (!p) continue;
    var c = String(mValues[r][idxC] != null ? mValues[r][idxC] : '').trim();
    var url = String(mValues[r][idxMainUrl] != null ? mValues[r][idxMainUrl] : '').trim();
    if (!byParent[p]) byParent[p] = { parentRow: -1, parentUrl: '', childUrl: '' };
    if (!c) {
      byParent[p].parentRow = r;
      byParent[p].parentUrl = url;
    } else if (url && !byParent[p].childUrl) {
      byParent[p].childUrl = url;
    }
  }
  var keys = Object.keys(byParent);
  var k;
  for (k = 0; k < keys.length; k++) {
    var g = byParent[keys[k]];
    if (g.parentRow < 0 || !g.childUrl) continue;
    if (g.parentUrl) continue;
    master.getRange(g.parentRow + 1, idxMainUrl + 1).setValue(g.childUrl);
    out.copied++;
    Logger.log('[U4] parentMainUrlCopy parentSku=' + keys[k] + ' row1=' + (g.parentRow + 1));
  }
  return out;
}

function amazonU4EnsureMasterUrlColumns_(masterSheet, headerRowIdx) {
  var headers = masterSheet.getRange(headerRowIdx + 1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
  var colMap = getColumnIndexMap(headers);
  var needed = [AMAZON_U4_MASTER_COL_MAIN_URL, AMAZON_U4_MASTER_COL_PT_URL];
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

function amazonU4FindMasterChildRow_(mValues, headerRowIdx, idxP, idxC, childSku) {
  if (idxC === undefined) return -1;
  var r;
  for (r = headerRowIdx + 1; r < mValues.length; r++) {
    if (String(mValues[r][idxC] || '').trim() === childSku) return r;
  }
  return -1;
}
