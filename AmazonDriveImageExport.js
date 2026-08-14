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

/**
 * R2 Put を最大 attempts 回（一時的な「使用できないアドレス」対策）。
 * @return {{ok:boolean, code:number, error:string, attempts:number}}
 */
function amazonR2PutObjectRetry_(accountId, accessKey, secretKey, bucket, objectKey, bytes, contentType, attempts) {
  attempts = attempts > 0 ? attempts : AMAZON_U4_PUT_RETRY_DEFAULT;
  var last = { ok: false, code: 0, error: '', attempts: 0 };
  var i;
  for (i = 1; i <= attempts; i++) {
    try {
      last = amazonR2PutObject_(accountId, accessKey, secretKey, bucket, objectKey, bytes, contentType);
      last.attempts = i;
      if (last.ok) return last;
      Logger.log('[U4] putRetry key=' + objectKey + ' attempt=' + i + '/' + attempts +
        ' http=' + last.code);
    } catch (ePut) {
      last = {
        ok: false,
        code: 0,
        error: String((ePut && ePut.message) || ePut).substring(0, 500),
        attempts: i
      };
      Logger.log('[U4] putRetry key=' + objectKey + ' attempt=' + i + '/' + attempts +
        ' throw=' + last.error);
    }
    if (i < attempts) {
      Utilities.sleep(1500 * i);
    }
  }
  return last;
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

// ----- U4: Drive02 → R2 → マスタ URL（2026-07-26 承認／2026-08-08 resume） -----
var AMAZON_U4_URL_EMBED_PROP = 'AMAZON_U4_URL_EMBED_ENABLED';
var AMAZON_U4_MAX_SKUS_PROP = 'AMAZON_U4_MAX_SKUS';
var AMAZON_U4_SKU_LIST_PROP = 'AMAZON_U4_SKU_LIST';
var AMAZON_U4_FORCE_REUPLOAD_PROP = 'AMAZON_U4_FORCE_REUPLOAD';
var AMAZON_U4_SLICE_MS_PROP = 'AMAZON_U4_SLICE_MS';
var AMAZON_U4_RESUME_STATE_PROP = 'AMAZON_U4_RESUME_STATE';
var AMAZON_U4_RESUME_TRIGGER_FN = 'runAmazonU4ResumeFromTrigger';
var AMAZON_U4_MASTER_COL_MAIN_URL = 'Amazon MAIN URL';
var AMAZON_U4_MASTER_COL_PT_URL = 'Amazon PT URL';
var AMAZON_U4_MAX_SKUS_DEFAULT = 20;
/** 1実行の実働上限（未設定時）。UrlFetch多発のため短め。残りはトリガー再開 */
var AMAZON_U4_SLICE_MS_DEFAULT = 270000;
var AMAZON_U4_PUT_RETRY_DEFAULT = 3;

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
        '\nスキップ(充足)=' + (summary.skippedFull || 0) +
        '／MAINのみ=' + (summary.skippedMain || 0) +
        '／PTのみ=' + (summary.skippedPt || 0) +
        '\n失敗=' + summary.failed +
        '\nマスタ更新SKU=' + summary.masterUpdated +
        (summary.continued
          ? ('\n\n※時間スライス: 残り' + (summary.remaining || 0) + '件を約1分後に自動再開します')
          : '') +
        '\n' + (summary.message || '') +
        '\n\n終わったら ' + AMAZON_U4_URL_EMBED_PROP + ' を false に戻してください。\n' +
        '強制再上げは ' + AMAZON_U4_FORCE_REUPLOAD_PROP + '=true（終わったら false）。\n' +
        '次: 21-①／D Amazon で GENERATED（Amazon URL優先）。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e3) {}
}

/**
 * U4 コア（確認・成功ダイアログなし）。Eコース・D新規の自動実行用。
 * @param {{skus?:Array<string>, force?:boolean}=} opts
 *   skus 指定時はレ点由来の対象だけを処理する。force=true は手動運用トグルを迂回する（D自動実行用）。
 * @return {{ok:boolean, runId?:string, summary?:Object, error?:string}}
 */
function amazonU4UrlEmbedSilent_(opts) {
  opts = opts || {};
  var stepName = 'AmazonU4UrlEmbedSilent';
  var runId = (opts.resumeRunId && String(opts.resumeRunId).trim())
    ? String(opts.resumeRunId).trim()
    : ('U4_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') +
      '_' + String(Utilities.getUuid()).replace(/-/g, '').substring(0, 6));
  Logger.log('[' + stepName + '] runId=' + runId + ' state=PENDING resume=' +
    (opts.resumeRunId ? '1' : '0'));

  if (!opts.force && !getBoolScriptProperty_(AMAZON_U4_URL_EMBED_PROP, false)) {
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

  var targets = [];
  if (opts.skus && opts.skus.length) {
    var seenSku = {};
    // 再開時はキュー全体を処理（スライスが再度分割）。初回収集時のみ maxSkus で切る
    var cap = opts.resumeRunId ? opts.skus.length : maxSkus;
    for (var si = 0; si < opts.skus.length && targets.length < cap; si++) {
      var one = String(opts.skus[si] || '').trim();
      if (!one || seenSku[one]) continue;
      seenSku[one] = true;
      targets.push(one);
    }
  } else {
    targets = amazonU4CollectTargetSkus_(maxSkus);
  }
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
    Logger.log('[' + stepName + '] runId=' + runId + ' state=' +
      (summary && summary.continued ? 'CONTINUE' : 'DONE') + ' ' + JSON.stringify(summary));
    if (summary && Number(summary.failed) > 0 && !(Number(summary.mainOk) > 0) &&
        !(Number(summary.skippedFull) > 0) && !(Number(summary.skippedMain) > 0) &&
        !summary.continued) {
      return {
        ok: false,
        runId: runId,
        summary: summary,
        error: 'U4失敗: MAIN成功=0 失敗=' + summary.failed + ' ' + (summary.message || '') +
          '／次: C「Ama新カタログ②：楽天サブ→マスタ反映→URLをR2へUP」（02手置きは復旧例外のみ）'
      };
    }
    return { ok: true, runId: runId, summary: summary, continued: !!(summary && summary.continued) };
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
 * 公開HTTPSとして使えるか（API用 cloudflarestorage は不可）。
 * @param {string} url
 * @return {boolean}
 */
function amazonU4IsUsablePublicUrl_(url) {
  var u = String(url || '').trim();
  if (!u || u.indexOf('https://') !== 0) return false;
  if (u.indexOf('r2.cloudflarestorage.com') >= 0) return false;
  if (u.indexOf('使用できない') >= 0) return false;
  return true;
}

/**
 * @return {{mainOk:number, ptOk:number, failed:number, skippedFull:number, skippedMain:number, skippedPt:number,
 *   masterUpdated:number, message:string, continued:boolean, remaining:number, parentMainUrlCopied:number, siblingPtUrlCopied:number}}
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

  var props = PropertiesService.getScriptProperties();
  var forceReupload = getBoolScriptProperty_(AMAZON_U4_FORCE_REUPLOAD_PROP, false);
  var sliceMs = parseInt(String(props.getProperty(AMAZON_U4_SLICE_MS_PROP) || ''), 10);
  if (!(sliceMs > 0)) sliceMs = AMAZON_U4_SLICE_MS_DEFAULT;
  var startedAt = Date.now();

  var mainOk = 0;
  var ptOk = 0;
  var failed = 0;
  var skippedFull = 0;
  var skippedMain = 0;
  var skippedPt = 0;
  var masterUpdated = 0;
  var notes = [];
  var continued = false;
  var remainingSkus = [];
  if (ensure.added.length) notes.push('列追加=' + ensure.added.join(','));
  if (forceReupload) notes.push('FORCE_REUPLOAD=true');

  var s;
  for (s = 0; s < skus.length; s++) {
    if (s > 0 && (Date.now() - startedAt) >= sliceMs) {
      remainingSkus = skus.slice(s);
      continued = true;
      amazonU4ScheduleResume_({
        runId: runId,
        skus: remainingSkus,
        startedAtIso: new Date().toISOString()
      });
      notes.push('SLICE残り=' + remainingSkus.length + '（約1分後トリガー再開）');
      Logger.log('[U4] runId=' + runId + ' state=SLICE remaining=' + remainingSkus.length +
        ' elapsedMs=' + (Date.now() - startedAt));
      break;
    }

    var sku = skus[s];
    try {
      var mRow = amazonU4FindMasterChildRow_(mValues, headerRowIdx, idxP, idxC, sku);
      var mode = '';
      if (mRow >= 0 && idxMode !== undefined) {
        mode = String(mValues[mRow][idxMode] || '').trim().toUpperCase();
      }
      var existingMainUrl = (mRow >= 0 && idxMainUrl !== undefined)
        ? String(mValues[mRow][idxMainUrl] || '').trim() : '';
      var existingPtUrl = (mRow >= 0 && idxPtUrl !== undefined)
        ? String(mValues[mRow][idxPtUrl] || '').trim() : '';
      var mainUsable = amazonU4IsUsablePublicUrl_(existingMainUrl);
      var ptUsable = amazonU4IsUsablePublicUrl_(existingPtUrl);

      if (!forceReupload && mainUsable && ptUsable) {
        skippedFull++;
        Logger.log('[U4] runId=' + runId + ' sku=' + sku + ' state=SKIP_FULL reason=urls_present');
        continue;
      }

      var mainName = sku + '.MAIN.jpg';
      var file = amazonDriveFindFileByName_(folderId, mainName);
      var mainUrl = '';
      if (!forceReupload && mainUsable) {
        mainUrl = existingMainUrl;
        skippedMain++;
        mainOk++;
        Logger.log('[U4] runId=' + runId + ' sku=' + sku + ' state=SKIP_MAIN url=' + mainUrl);
      } else if (file) {
        var bytes = file.getBlob().getBytes();
        var putMain = amazonR2PutObjectRetry_(
          accountId, accessKey, secretKey, bucket, mainName, bytes, 'image/jpeg', AMAZON_U4_PUT_RETRY_DEFAULT
        );
        mainUrl = publicBase + '/' + mainName;
        Logger.log(
          '[U4] runId=' + runId + ' sku=' + sku + ' http=' + putMain.code +
            ' attempts=' + (putMain.attempts || 1) +
            ' url=' + mainUrl + ' state=' + (putMain.ok ? 'DONE' : 'FAILED') +
            (putMain.error ? ' err=' + String(putMain.error).substring(0, 120) : '')
        );
        if (!putMain.ok) {
          failed++;
          notes.push(sku + ' MAIN Put http=' + putMain.code + ' ' + (putMain.error || ''));
          continue;
        }
        mainOk++;
      } else if (amazonU4IsUsablePublicUrl_(existingMainUrl)) {
        mainUrl = existingMainUrl;
        Logger.log('[U4] runId=' + runId + ' sku=' + sku + ' state=REUSE_EXISTING_MAIN');
        mainOk++;
      } else {
        failed++;
        notes.push(sku + ': MAINファイルなし（次: C-1→ドラッグ→C-2。02手置きは復旧例外のみ）');
        Logger.log('[U4] runId=' + runId + ' sku=' + sku + ' state=FAILED noMainFile');
        continue;
      }

      var ptUrls = [];
      var ptSeq = 0;
      if (!forceReupload && ptUsable) {
        skippedPt++;
        Logger.log('[U4] runId=' + runId + ' sku=' + sku + ' state=SKIP_PT urlPresent=1');
      } else {
        // 1) Amazon PT 参照（Drive ID）があれば優先（画像モードに関わらず）
        if (mRow >= 0 && idxPtRef !== undefined) {
          var ptRefs = String(mValues[mRow][idxPtRef] || '').split('|');
          var pi;
          for (pi = 0; pi < ptRefs.length; pi++) {
            var ptId = String(ptRefs[pi] || '').trim();
            if (!ptId) continue;
            ptSeq++;
            var ptName = sku + '.PT' + ('0' + ptSeq).slice(-2) + '.jpg';
            try {
              var ptFile = DriveApp.getFileById(ptId);
              var ptBytes = ptFile.getBlob().getBytes();
              var putPt = amazonR2PutObjectRetry_(
                accountId, accessKey, secretKey, bucket, ptName, ptBytes, 'image/jpeg', AMAZON_U4_PUT_RETRY_DEFAULT
              );
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

        // 2) PT参照が空／失敗0件 → 楽天サブ画像1〜8を取得して R2 へ（AMAZON_ONLY は除外）
        if (!ptUrls.length && mode.indexOf('ONLY') < 0 && mRow >= 0) {
          var rakutenSubs = amazonU4CollectRakutenSubUrls_(mValues, headerRowIdx, colMap, mRow, idxP);
          var ri;
          for (ri = 0; ri < rakutenSubs.length; ri++) {
            ptSeq++;
            var ptNameR = sku + '.PT' + ('0' + ptSeq).slice(-2) + '.jpg';
            try {
              var rBytes = amazonU4FetchImageBytesRetry_(rakutenSubs[ri]);
              if (!rBytes || !rBytes.length) {
                failed++;
                notes.push(sku + ' ' + ptNameR + ': 楽天画像取得失敗');
                continue;
              }
              var putPtR = amazonR2PutObjectRetry_(
                accountId, accessKey, secretKey, bucket, ptNameR, rBytes, 'image/jpeg', AMAZON_U4_PUT_RETRY_DEFAULT
              );
              if (putPtR.ok) {
                ptOk++;
                ptUrls.push(publicBase + '/' + ptNameR);
              } else {
                failed++;
                notes.push(sku + ' ' + ptNameR + ' http=' + putPtR.code);
              }
            } catch (eR) {
              failed++;
              notes.push(sku + ' 楽天PT: ' + ((eR && eR.message) || eR));
            }
          }
          if (rakutenSubs.length) {
            Logger.log('[U4] runId=' + runId + ' sku=' + sku +
              ' rakutenSubSources=' + rakutenSubs.length + ' ptUploaded=' + ptUrls.length);
          }
        }
      }

      Logger.log('[U4] runId=' + runId + ' sku=' + sku + ' imageMode=' + (mode || '(空)') +
        ' ptOkThisSku=' + ptUrls.length);

      if (mRow >= 0 && idxMainUrl !== undefined) {
        master.getRange(mRow + 1, idxMainUrl + 1).setValue(mainUrl);
        if (idxPtUrl !== undefined && ptUrls.length) {
          master.getRange(mRow + 1, idxPtUrl + 1).setValue(ptUrls.join('|'));
        }
        masterUpdated++;
        mValues[mRow][idxMainUrl] = mainUrl;
        if (ptUrls.length && idxPtUrl !== undefined) {
          mValues[mRow][idxPtUrl] = ptUrls.join('|');
        }
      } else {
        notes.push(sku + ': マスタ子行なし（R2のみ成功）');
      }
    } catch (eSku) {
      failed++;
      var em = String((eSku && eSku.message) || eSku);
      notes.push(sku + ': ' + em);
      Logger.log('[U4] runId=' + runId + ' sku=' + sku + ' state=FAILED_CONTINUE ' + em);
    }
  }

  var prop = { copied: 0 };
  var propPt = { copied: 0 };
  var siblingPt = { copied: 0 };
  if (!continued) {
    prop = amazonU4PropagateMainUrlToParents_(master, mValues, headerRowIdx, idxP, idxC, idxMainUrl);
    if (prop && prop.copied > 0) {
      masterUpdated += prop.copied;
      notes.push('親MAIN URLコピー=' + prop.copied);
      Logger.log('[U4] runId=' + runId + ' parentMainUrlCopied=' + prop.copied);
    }
    propPt = amazonU4PropagateMainUrlToParents_(master, mValues, headerRowIdx, idxP, idxC, idxPtUrl, 'PT');
    if (propPt && propPt.copied > 0) {
      masterUpdated += propPt.copied;
      notes.push('親PT URLコピー=' + propPt.copied);
      Logger.log('[U4] runId=' + runId + ' parentPtUrlCopied=' + propPt.copied);
    }
    siblingPt = amazonU4PropagatePtUrlToTargetSiblings_(
      master, headerRowIdx, idxP, idxC, idxPtUrl, skus, runId
    );
    if (siblingPt && siblingPt.copied > 0) {
      masterUpdated += siblingPt.copied;
      notes.push('兄弟PT URLコピー=' + siblingPt.copied);
      Logger.log('[U4] runId=' + runId + ' siblingPtUrlCopied=' + siblingPt.copied);
    }
    amazonU4ClearResumeState_();
  }

  return {
    mainOk: mainOk,
    ptOk: ptOk,
    failed: failed,
    skippedFull: skippedFull,
    skippedMain: skippedMain,
    skippedPt: skippedPt,
    masterUpdated: masterUpdated,
    message: notes.slice(0, 12).join('\n') + (notes.length > 12 ? '\n…' : ''),
    continued: continued,
    remaining: remainingSkus.length,
    parentMainUrlCopied: prop ? prop.copied : 0,
    siblingPtUrlCopied: siblingPt ? siblingPt.copied : 0
  };
}

function amazonU4ScheduleResume_(state) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(AMAZON_U4_RESUME_STATE_PROP, JSON.stringify(state || {}));
  amazonU4DeleteResumeTriggers_();
  ScriptApp.newTrigger(AMAZON_U4_RESUME_TRIGGER_FN).timeBased().after(60 * 1000).create();
  Logger.log('[U4] resume scheduled skus=' + ((state && state.skus && state.skus.length) || 0) +
    ' runId=' + ((state && state.runId) || ''));
}

function amazonU4ClearResumeState_() {
  PropertiesService.getScriptProperties().deleteProperty(AMAZON_U4_RESUME_STATE_PROP);
  amazonU4DeleteResumeTriggers_();
}

function amazonU4DeleteResumeTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === AMAZON_U4_RESUME_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * U4 時間スライス再開（1分後トリガー）。人間操作不要。
 */
function runAmazonU4ResumeFromTrigger() {
  var fn = 'runAmazonU4ResumeFromTrigger';
  var props = PropertiesService.getScriptProperties();
  var raw = String(props.getProperty(AMAZON_U4_RESUME_STATE_PROP) || '').trim();
  if (!raw) {
    Logger.log('[' + fn + '] state=SKIP no_resume_state');
    amazonU4DeleteResumeTriggers_();
    return;
  }
  var state;
  try {
    state = JSON.parse(raw);
  } catch (eJ) {
    Logger.log('[' + fn + '] state=FAILED bad_json');
    amazonU4ClearResumeState_();
    return;
  }
  var skus = (state && state.skus) ? state.skus : [];
  if (!skus.length) {
    Logger.log('[' + fn + '] state=SKIP empty_skus');
    amazonU4ClearResumeState_();
    return;
  }
  Logger.log('[' + fn + '] state=RUNNING n=' + skus.length + ' parentRunId=' + (state.runId || ''));
  var res = amazonU4UrlEmbedSilent_({
    skus: skus,
    force: true,
    resumeRunId: state.runId || ''
  });
  Logger.log('[' + fn + '] state=' + (res && res.ok ? 'DONE' : 'FAILED') +
    ' ' + JSON.stringify((res && res.summary) || {}) +
    (res && res.error ? ' err=' + res.error : ''));
}

/**
 * 外部画像URL取得（リトライ付き）。
 * @return {Byte[]|null}
 */
function amazonU4FetchImageBytesRetry_(url) {
  var i;
  for (i = 1; i <= AMAZON_U4_PUT_RETRY_DEFAULT; i++) {
    try {
      var bytes = amazonU4FetchImageBytes_(url);
      if (bytes && bytes.length) return bytes;
    } catch (eF) {
      Logger.log('[U4] fetchImageRetry attempt=' + i + ' ' + ((eF && eF.message) || eF));
    }
    if (i < AMAZON_U4_PUT_RETRY_DEFAULT) Utilities.sleep(1000 * i);
  }
  return null;
}

/**
 * 親行の Amazon MAIN URL が空のとき、同一親の子行から先頭の非空URLをコピーする。
 * @return {{copied:number}}
 */
function amazonU4PropagateMainUrlToParents_(master, mValues, headerRowIdx, idxP, idxC, idxMainUrl, label) {
  var out = { copied: 0 };
  label = label || 'MAIN';
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
    Logger.log('[U4] parentUrlCopy kind=' + label + ' parentSku=' + keys[k] + ' row1=' + (g.parentRow + 1));
  }
  return out;
}

/**
 * 今回のU4対象子SKUで Amazon PT URL が空なら、同一親内で上から最初の
 * PT URL 非空子SKUの値をコピーする。既存値と対象外SKUは変更しない。
 * @return {{copied:number}}
 */
function amazonU4PropagatePtUrlToTargetSiblings_(
  master, headerRowIdx, idxP, idxC, idxPtUrl, targetSkus, runId
) {
  var out = { copied: 0 };
  if (idxP == null || idxC == null || idxPtUrl == null || !master) return out;

  var targetSet = {};
  var i;
  for (i = 0; i < (targetSkus || []).length; i++) {
    var targetSku = String(targetSkus[i] || '').trim();
    if (targetSku) targetSet[targetSku] = true;
  }
  if (!Object.keys(targetSet).length) return out;

  var values = master.getDataRange().getValues();
  var byParent = {};
  var r;
  for (r = headerRowIdx + 1; r < values.length; r++) {
    var parentSku = String(values[r][idxP] != null ? values[r][idxP] : '').trim();
    var childSku = String(values[r][idxC] != null ? values[r][idxC] : '').trim();
    if (!parentSku || !childSku) continue;
    var ptUrl = String(values[r][idxPtUrl] != null ? values[r][idxPtUrl] : '').trim();
    if (!byParent[parentSku]) {
      byParent[parentSku] = { sourceSku: '', sourceUrl: '', emptyTargets: [] };
    }
    if (ptUrl && !byParent[parentSku].sourceUrl) {
      byParent[parentSku].sourceSku = childSku;
      byParent[parentSku].sourceUrl = ptUrl;
    }
    if (targetSet[childSku] && !ptUrl) {
      byParent[parentSku].emptyTargets.push({ row: r, sku: childSku });
    }
  }

  var parents = Object.keys(byParent);
  for (i = 0; i < parents.length; i++) {
    var group = byParent[parents[i]];
    if (!group.sourceUrl || !group.emptyTargets.length) continue;
    for (var j = 0; j < group.emptyTargets.length; j++) {
      var dest = group.emptyTargets[j];
      master.getRange(dest.row + 1, idxPtUrl + 1).setValue(group.sourceUrl);
      out.copied++;
      Logger.log(
        '[U4] runId=' + runId + ' siblingPtUrlCopy parentSku=' + parents[i] +
        ' sourceSku=' + group.sourceSku + ' targetSku=' + dest.sku +
        ' row1=' + (dest.row + 1)
      );
    }
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

/**
 * 楽天サブ画像1〜8を収集（子優先、空なら同一親の親行）。最大8件。絶対URLへ正規化済み。
 * @return {string[]}
 */
function amazonU4CollectRakutenSubUrls_(mValues, headerRowIdx, colMap, childRowIdx, idxP) {
  var out = [];
  var parentSku = idxP !== undefined
    ? String(mValues[childRowIdx][idxP] || '').trim() : '';
  var parentRow = -1;
  if (parentSku && idxP !== undefined) {
    var idxC = colMap['子SKU'];
    var r;
    for (r = headerRowIdx + 1; r < mValues.length; r++) {
      if (String(mValues[r][idxP] || '').trim() !== parentSku) continue;
      var c = idxC !== undefined ? String(mValues[r][idxC] || '').trim() : '';
      if (!c) {
        parentRow = r;
        break;
      }
    }
  }
  var i;
  for (i = 1; i <= 8; i++) {
    var idx = colMap['楽天サブ画像' + i];
    if (idx === undefined) continue;
    var raw = String(mValues[childRowIdx][idx] || '').trim();
    if (!raw && parentRow >= 0) raw = String(mValues[parentRow][idx] || '').trim();
    if (!raw) continue;
    var abs = amazonU4NormalizeRakutenImageUrl_(raw);
    if (abs && out.indexOf(abs) < 0) out.push(abs);
  }
  return out;
}

/**
 * 楽天Cabinet相対パス → 公開https。既にhttpならそのまま。
 * Yahoo.js と同型: `/…` → `https://image.rakuten.co.jp/octas/cabinet` + path
 */
function amazonU4NormalizeRakutenImageUrl_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  if (/^https?:\/\//i.test(s)) return s;
  if (s.indexOf('//') === 0) return 'https:' + s;
  if (s.charAt(0) !== '/') s = '/' + s;
  // R-Cabinet保存後は `/123/img.jpg` 形式が多い。octas店舗固定（Yahoo.js準拠）
  return 'https://image.rakuten.co.jp/octas/cabinet' + s;
}

/**
 * 外部画像URLを取得してバイト配列を返す。失敗時は null。
 * @return {Byte[]|null}
 */
function amazonU4FetchImageBytes_(url) {
  var abs = amazonU4NormalizeRakutenImageUrl_(url);
  if (!abs) return null;
  var res = UrlFetchApp.fetch(abs, {
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true
  });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('[U4] fetchImage http=' + code + ' url=' + abs.substring(0, 120));
    return null;
  }
  var blob = res.getBlob();
  var bytes = blob.getBytes();
  if (!bytes || !bytes.length) return null;
  return bytes;
}

function amazonU4FindMasterChildRow_(mValues, headerRowIdx, idxP, idxC, childSku) {
  if (idxC === undefined) return -1;
  var r;
  for (r = headerRowIdx + 1; r < mValues.length; r++) {
    if (String(mValues[r][idxC] || '').trim() === childSku) return r;
  }
  return -1;
}
