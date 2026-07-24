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
