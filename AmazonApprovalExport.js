/**
 * Lv4 承認①済 → Amazon バルク埋め用データ（GENERATED）
 * 要件: docs/org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md
 *
 * - 楽天 generateRakutenCSV / Yahoo.js 出品APIは触らない
 * - マスタ「在庫数」「JAN」への書込禁止（読取のみ）
 * - 純正 .xlsm 編集・SC自動UPはしない（PACKAGED/UPは人間＋ローカル）
 * - TRACK 未設定は実行しない
 *
 * Script Properties:
 *   APPROVAL_AMAZON_LV4_ENABLED … 既定 false
 *   APPROVAL_AMAZON_LV4_TRACK … A | B | BOTH（必須。空＝実行しない）
 *   APPROVAL_AMAZON_LV4_SKIP_EXPORT … 既定 false（true=Drive書込スキップのドライラン）
 *   APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE … 既定 送料無料パターン
 *   APPROVAL_AMAZON_LV4_FOLDER_ID … GENERATED 保存先（空ならスプレッドシート親にフォルダ作成）
 *   APPROVAL_AMAZON_LV4_PARENTS_PER_SUB … 副制約・親件数/サブバッチ（既定 5）
 *   APPROVAL_AMAZON_LV4_LEAD_TIME_DAYS … エラー時のみ後付け用（初版は未使用）
 *   APPROVAL_AMAZON_LV4_STATE … レジューム用（自動）
 *   APPROVAL_AMAZON_LV4_BRAND_GATE_MODE … M2(A): manual_ok 必須（人間が制限なし確認後）。未設定＝SKIPPED_BRAND_GATE
 */

var APPROVAL_AMAZON_LV4_PROP = 'APPROVAL_AMAZON_LV4_ENABLED';
var APPROVAL_AMAZON_LV4_TRACK_PROP = 'APPROVAL_AMAZON_LV4_TRACK';
var APPROVAL_AMAZON_LV4_SKIP_EXPORT_PROP = 'APPROVAL_AMAZON_LV4_SKIP_EXPORT';
var APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE_PROP = 'APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE';
var APPROVAL_AMAZON_LV4_FOLDER_ID_PROP = 'APPROVAL_AMAZON_LV4_FOLDER_ID';
var APPROVAL_AMAZON_LV4_PARENTS_PER_SUB_PROP = 'APPROVAL_AMAZON_LV4_PARENTS_PER_SUB';
var APPROVAL_AMAZON_LV4_STATE_PROP = 'APPROVAL_AMAZON_LV4_STATE';
var APPROVAL_AMAZON_LV4_BRAND_GATE_MODE_PROP = 'APPROVAL_AMAZON_LV4_BRAND_GATE_MODE';
var APPROVAL_AMAZON_LV4_TIME_MS = 25 * 60 * 1000;
var APPROVAL_AMAZON_LV4_MAX_TRIGGER_RUNS = 40;
var APPROVAL_AMAZON_LV4_TRIGGER_FN = 'runApprovalAmazonLv4FromTrigger';
var APPROVAL_AMAZON_LV4_LOG_SHEET = '▼Lv4実行ログ(Amazon)';
var APPROVAL_AMAZON_LV4_DEFAULT_SHIPPING = '送料無料パターン';
var APPROVAL_AMAZON_LV4_BRAND = 'ノーブランド品';

var APPROVAL_AMAZON_LV4_LOG_HEADERS = [
  'recordType', 'runId', 'batchId', 'subBatchId', 'track', 'parentSkus', 'childCount',
  'status', 'fileName', 'fileUrl', 'generatedAt', 'packagedAt', 'uploadedOkAt',
  'note', 'category', 'brand', 'exemptionDate', 'evidenceUrl'
];

/**
 * メニュー 21-①: 最新 APPROVED の amazon を親単位サブバッチで GENERATED する。
 */
/**
 * メニュー 21-① / D×Amazon ファサード共用。
 * @param {{silent?:boolean}=} opts silent=true なら確認・成功ダイアログなし（エラー時のみ alert）
 * @return {{ok:boolean, cancelled?:boolean, runId?:string, track?:string, summary?:Object, reason?:string, error?:string}}
 */
function menuApprovalAmazonLv4Run(opts) {
  var fn = 'menuApprovalAmazonLv4Run';
  opts = (opts && typeof opts === 'object') ? opts : {};
  var silent = !!opts.silent;
  if (!getBoolScriptProperty_(APPROVAL_AMAZON_LV4_PROP, false)) {
    var off = 'Lv4 Amazonは無効です。Script Properties の ' + APPROVAL_AMAZON_LV4_PROP + ' を true にしてください。';
    Logger.log('[' + fn + '] state=FAILED ' + off);
    try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    return { ok: false, reason: off };
  }
  var track = amazonApprovalLv4NormalizeTrack_(
    PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_TRACK_PROP)
  );
  if (!track) {
    var noTrack = 'TRACK 未設定のため実行しません。' + APPROVAL_AMAZON_LV4_TRACK_PROP +
      ' に A / B / BOTH のいずれかを設定してください（M1は B）。';
    Logger.log('[' + fn + '] state=FAILED ' + noTrack);
    try { SpreadsheetApp.getUi().alert(noTrack); } catch (eT) {}
    return { ok: false, reason: noTrack };
  }
  if (!silent) {
    try {
      var ui = SpreadsheetApp.getUi();
      var skip = getBoolScriptProperty_(APPROVAL_AMAZON_LV4_SKIP_EXPORT_PROP, false);
      var ship = amazonApprovalLv4ShippingTemplate_();
      var res = ui.alert(
        'Lv4 Amazonバルク（GENERATED）',
        '最新の承認①（APPROVED）の Amazon 明細を対象に、埋め用データを Drive へ出力します。\n' +
          '・TRACK=' + track + '\n' +
          '・配送テンプレ=' + ship + '\n' +
          '・マスタ在庫・JANは書き込みません\n' +
          '・純正 .xlsm / Seller Central UP は人間作業です\n' +
          (skip ? '・SKIP_EXPORT=true（ドライラン・Drive未書込）\n' : '') +
          '\n実行しますか？',
        ui.ButtonSet.OK_CANCEL
      );
      if (res !== ui.Button.OK) return { ok: false, cancelled: true };
    } catch (eUi) {}
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var runId = 'LV4_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' +
    ('000000' + Math.floor(Math.random() * 1e6)).slice(-6);
  Logger.log('[' + fn + '] state=RUNNING runId=' + runId + ' track=' + track + ' silent=' + silent);
  try {
    var summary = amazonApprovalLv4Run_(ss, runId, 0, null, track);
    Logger.log('[' + fn + '] state=DONE runId=' + runId + ' ' + JSON.stringify(summary));
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(
          'Lv4 実行結果',
          'runId=' + runId +
            '\nbatchId=' + (summary.batchId || '') +
            '\ntrack=' + track +
            '\nサブバッチ完了=' + summary.subBatchesDone +
            '\n親成功=' + summary.parentsDone +
            '\nスキップ=' + summary.skipped +
            '\n継続待ち=' + (summary.willResume ? 'YES' : 'NO') +
            '\n' + (summary.message || ''),
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } catch (e1) {}
    }
    return { ok: true, runId: runId, track: track, summary: summary };
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED runId=' + runId + ' ' + ((err && err.message) || err));
    amazonApprovalLv4Mail_('【Lv4 Amazon】実行失敗', 'runId=' + runId + '\n' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert('Lv4失敗: ' + ((err && err.message) || err)); } catch (e2) {}
    return { ok: false, runId: runId, track: track, error: String((err && err.message) || err) };
  }
}

/** メニュー 21-②: レジューム状態クリア */
function menuApprovalAmazonLv4ClearState() {
  try {
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_AMAZON_LV4_STATE_PROP);
    amazonApprovalLv4DeleteTriggers_();
    SpreadsheetApp.getUi().alert('Lv4 実行状態をクリアしました。');
  } catch (e) {
    try { SpreadsheetApp.getUi().alert(String(e && e.message || e)); } catch (e2) {}
  }
}

/**
 * メニュー 21-③: アップロード成功を記録（UPLOADED_OK）
 */
function menuApprovalAmazonLv4MarkUploadedOk() {
  var fn = 'menuApprovalAmazonLv4MarkUploadedOk';
  try {
    var ui = SpreadsheetApp.getUi();
    var ans = ui.prompt(
      'Lv4 アップロード成功を記録',
      'subBatchId を入力（例: {batchId}_B1）',
      ui.ButtonSet.OK_CANCEL
    );
    if (ans.getSelectedButton() !== ui.Button.OK) return;
    var subBatchId = String(ans.getResponseText() || '').trim();
    if (!subBatchId) {
      ui.alert('subBatchId が空です');
      return;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var updated = amazonApprovalLv4MarkStatus_(ss, subBatchId, 'UPLOADED_OK', '');
    Logger.log('[' + fn + '] state=DONE subBatchId=' + subBatchId + ' updated=' + updated);
    ui.alert('UPLOADED_OK を記録しました（更新行=' + updated + '）\n※在庫0/1の実機見え方未検証なら「掲載完了」と断言しないでください');
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert(String(err && err.message || err)); } catch (e2) {}
  }
}

/**
 * メニュー 21-④: アップロード失敗を記録（UPLOAD_FAILED）→ 同一 subBatchId で再生成可
 */
function menuApprovalAmazonLv4MarkUploadFailed() {
  var fn = 'menuApprovalAmazonLv4MarkUploadFailed';
  try {
    var ui = SpreadsheetApp.getUi();
    var ans = ui.prompt(
      'Lv4 アップロード失敗を記録',
      'subBatchId と理由（任意）を「subBatchId|理由」形式で入力',
      ui.ButtonSet.OK_CANCEL
    );
    if (ans.getSelectedButton() !== ui.Button.OK) return;
    var raw = String(ans.getResponseText() || '').trim();
    var parts = raw.split('|');
    var subBatchId = String(parts[0] || '').trim();
    var note = parts.length > 1 ? String(parts.slice(1).join('|')).trim() : '';
    if (!subBatchId) {
      ui.alert('subBatchId が空です');
      return;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var updated = amazonApprovalLv4MarkStatus_(ss, subBatchId, 'UPLOAD_FAILED', note);
    Logger.log('[' + fn + '] state=DONE subBatchId=' + subBatchId + ' updated=' + updated);
    ui.alert('UPLOAD_FAILED を記録しました（更新行=' + updated + '）。同一 subBatchId で再生成できます。');
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert(String(err && err.message || err)); } catch (e2) {}
  }
}

/** トリガー再開 */
function runApprovalAmazonLv4FromTrigger() {
  var fn = APPROVAL_AMAZON_LV4_TRIGGER_FN;
  var stateJson = PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_STATE_PROP);
  if (!stateJson) return;
  amazonApprovalLv4DeleteTriggers_();
  var state;
  try {
    state = JSON.parse(stateJson);
  } catch (e) {
    return;
  }
  var runCount = Number(state.triggerRunCount || 0) + 1;
  state.triggerRunCount = runCount;
  PropertiesService.getScriptProperties().setProperty(APPROVAL_AMAZON_LV4_STATE_PROP, JSON.stringify(state));
  if (runCount > APPROVAL_AMAZON_LV4_MAX_TRIGGER_RUNS) {
    Logger.log('[' + fn + '] state=FAILED maxTriggerRuns runId=' + state.runId);
    amazonApprovalLv4Mail_('【Lv4 Amazon】自動再開上限', 'runId=' + state.runId + ' batchId=' + state.batchId);
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_AMAZON_LV4_STATE_PROP);
    return;
  }
  if (!getBoolScriptProperty_(APPROVAL_AMAZON_LV4_PROP, false)) {
    Logger.log('[' + fn + '] state=FAILED disabled');
    return;
  }
  var track = amazonApprovalLv4NormalizeTrack_(state.track ||
    PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_TRACK_PROP));
  if (!track) {
    Logger.log('[' + fn + '] state=FAILED trackUnset');
    return;
  }
  var ss = SpreadsheetApp.openById(state.spreadsheetId);
  Logger.log('[' + fn + '] state=RUNNING runId=' + state.runId + ' resume remaining (index0 after doneParents filter)');
  try {
    // doneParents 除外後にサブバッチ再採番するため startIndex は常に 0
    amazonApprovalLv4Run_(ss, state.runId, 0, state, track);
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    amazonApprovalLv4Mail_('【Lv4 Amazon】再開失敗', 'runId=' + state.runId + '\n' + ((err && err.message) || err));
  }
}

/**
 * @param {Spreadsheet} ss
 * @param {string} runId
 * @param {number} startSubBatchIndex 再開時は無視し0から（doneParents除外後に再採番するため）
 * @param {Object|null} resumeState
 * @param {string} track
 * @return {Object}
 */
function amazonApprovalLv4Run_(ss, runId, startSubBatchIndex, resumeState, track) {
  var fn = 'amazonApprovalLv4Run_';
  var startedAt = Date.now();
  var loaded = (typeof approvalQueueGetLatestApprovedAmazon_ === 'function')
    ? approvalQueueGetLatestApprovedAmazon_()
    : { found: false };
  if (!loaded.found || !loaded.lines || !loaded.lines.length) {
    throw new Error('APPROVED の Amazon 明細がありません。先に承認キューで amazon 親＋子を承認①してください。');
  }
  var batch = loaded.batch;
  var batchId = batch.batchId;
  if (resumeState && resumeState.batchId && String(resumeState.batchId) !== String(batchId)) {
    throw new Error(
      '再開stateのbatchId不一致: state=' + resumeState.batchId + ' latestAPPROVED=' + batchId +
      '。21-②で状態クリアしてから再実行してください。'
    );
  }
  var inventoryMode = String(batch.inventoryMode || 'ZERO').toUpperCase() === 'ONE' ? 'ONE' : 'ZERO';
  var stockOut = inventoryMode === 'ONE' ? 1 : 0;
  var shipping = amazonApprovalLv4ShippingTemplate_();
  var skipExport = getBoolScriptProperty_(APPROVAL_AMAZON_LV4_SKIP_EXPORT_PROP, false);

  amazonApprovalLv4EnsureLogSheet_(ss);
  var masterCtx = amazonApprovalLv4LoadMasterContext_(ss);
  var exemptions = amazonApprovalLv4LoadExemptions_(ss);

  // 冪等: GENERATED/UPLOADED_OK/PACKAGED 済み親は除外。DRY_RUN・UPLOAD_FAILED はブロックしない
  var blockedMap = amazonApprovalLv4LoadBlockedParents_(ss, batchId, track);
  var excludeParents = [];
  if (resumeState && resumeState.doneParents) {
    excludeParents = resumeState.doneParents.slice();
  }
  var blockedKeys = Object.keys(blockedMap);
  for (var bi = 0; bi < blockedKeys.length; bi++) {
    excludeParents.push(blockedKeys[bi]);
  }
  // 冪等除外は recordType=SKIP（RUN の latest を汚染しない）
  var idempotentSkipped = [];
  for (var ik = 0; ik < blockedKeys.length; ik++) {
    idempotentSkipped.push({
      parentSku: blockedKeys[ik],
      reason: 'SKIPPED_IDEMPOTENT',
      detail: 'latest=' + blockedMap[blockedKeys[ik]] + '（UPLOAD_FAILEDのみ再生成可）'
    });
  }
  amazonApprovalLv4LogSkippedParents_(ss, runId, batchId, track, idempotentSkipped);

  var resolved = amazonApprovalLv4ResolveParents_(
    masterCtx, loaded.lines, track, excludeParents, exemptions
  );
  amazonApprovalLv4LogSkippedParents_(ss, runId, batchId, track, resolved.skipped);

  Logger.log('[' + fn + '] runId=' + runId + ' batchId=' + batchId +
    ' track=' + track + ' parents=' + resolved.parents.length +
    ' skipped=' + resolved.skipped.length +
    ' blockedIdempotent=' + blockedKeys.length +
    ' inventoryMode=' + inventoryMode +
    ' skipExport=' + skipExport);

  var perSub = amazonApprovalLv4ParentsPerSub_();
  var subBatches = amazonApprovalLv4BuildSubBatches_(resolved.parents, perSub);
  if (!subBatches.length) {
    return {
      batchId: batchId,
      subBatchesDone: 0,
      parentsDone: 0,
      skipped: resolved.skipped.length,
      willResume: false,
      message: '実行対象の親がありません（スキップ/冪等除外）: ' +
        amazonApprovalLv4SkipSummary_(resolved.skipped) +
        ';idempotentBlocked=' + blockedKeys.length
    };
  }

  // 再開時: 残り親は index0 から。subBatchId はシート上の最大番号から単調増加（衝突防止）
  var loopStart = 0;
  var doneParents = (resumeState && resumeState.doneParents) ? resumeState.doneParents.slice() : [];
  var subBatchesDone = Number(resumeState && resumeState.subBatchesDone) || 0;
  var trackTag = track === 'A' ? 'A' : (track === 'BOTH' ? 'X' : 'B');
  var nextSubSeq = amazonApprovalLv4NextSubBatchSeq_(ss, batchId, trackTag, resumeState);
  var willResume = false;
  var message = '';
  var folder = skipExport ? null : amazonApprovalLv4GetOrCreateFolder_(ss);

  for (var s = loopStart; s < subBatches.length; s++) {
    if ((Date.now() - startedAt) > APPROVAL_AMAZON_LV4_TIME_MS) {
      amazonApprovalLv4SaveState_({
        spreadsheetId: ss.getId(),
        runId: runId,
        batchId: batchId,
        track: track,
        inventoryMode: inventoryMode,
        nextSubBatchIndex: s,
        subBatchesDone: subBatchesDone,
        nextSubSeq: nextSubSeq,
        doneParents: doneParents,
        triggerRunCount: resumeState ? resumeState.triggerRunCount : 0
      });
      amazonApprovalLv4SetTrigger_();
      willResume = true;
      message = '時間予算のため中断。残り親から再開予定（subBatchIdは単調増加）';
      Logger.log('[' + fn + '] state=RETRYING runId=' + runId + ' ' + message);
      break;
    }

    var sub = subBatches[s];
    var subBatchId = batchId + '_' + trackTag + nextSubSeq;
    nextSubSeq++;
    Logger.log('[' + fn + '] state=RUNNING runId=' + runId + ' batchId=' + batchId +
      ' subBatchId=' + subBatchId + ' parents=' + sub.length + ' track=' + track);

    try {
      var built = amazonApprovalLv4BuildRows_(masterCtx, sub, track, stockOut, shipping);
      // Build 内で想定外スキップがあればログ（Resolve 済みなら通常0件）
      if (built.skipped && built.skipped.length) {
        amazonApprovalLv4LogSkippedParents_(ss, runId, batchId, track, built.skipped);
      }
      if (!built.rows.length || !built.okParents.length) {
        amazonApprovalLv4AppendLog_(ss, {
          recordType: 'RUN',
          runId: runId,
          batchId: batchId,
          subBatchId: subBatchId,
          track: track,
          parentSkus: amazonApprovalLv4ParentSkuList_(sub),
          childCount: 0,
          status: 'SKIPPED_EMPTY',
          fileName: '',
          fileUrl: '',
          note: '検証後に出力行なし（親はdoneにしない）'
        });
        continue;
      }

      var outStatus = skipExport ? 'DRY_RUN' : 'GENERATED';
      var fileName = '';
      var fileUrl = '';
      if (!skipExport) {
        var csv = amazonApprovalLv4RowsToCsv_(built.headers, built.rows);
        var blob = Utilities.newBlob(csv, 'text/csv', subBatchId + '_GENERATED.csv');
        amazonApprovalLv4ArchiveSameName_(folder, blob.getName());
        var file = folder.createFile(blob);
        fileName = file.getName();
        fileUrl = file.getUrl();
        var meta = {
          runId: runId,
          batchId: batchId,
          subBatchId: subBatchId,
          track: track,
          inventoryMode: inventoryMode,
          stockOut: stockOut,
          shippingTemplate: shipping,
          brand: APPROVAL_AMAZON_LV4_BRAND,
          parents: amazonApprovalLv4ParentSkuList_(built.okParents),
          rowCount: built.rows.length,
          note: 'GAS GENERATED only. PACKAGED=.xlsm is local. Do not wipe master JAN/stock.'
        };
        var metaName = subBatchId + '_GENERATED.meta.json';
        amazonApprovalLv4ArchiveSameName_(folder, metaName);
        folder.createFile(Utilities.newBlob(
          JSON.stringify(meta, null, 2),
          'application/json',
          metaName
        ));
      } else {
        fileName = '(SKIP_EXPORT)';
        fileUrl = '';
      }

      amazonApprovalLv4AppendLog_(ss, {
        recordType: 'RUN',
        runId: runId,
        batchId: batchId,
        subBatchId: subBatchId,
        track: track,
        parentSkus: amazonApprovalLv4ParentSkuList_(built.okParents),
        childCount: built.childCount,
        status: outStatus,
        fileName: fileName,
        fileUrl: fileUrl,
        note: skipExport ? 'DRY_RUN（冪等ブロック対象外。本番はSKIP_EXPORT=falseでGENERATED）' : ''
      });

      for (var p = 0; p < built.okParents.length; p++) {
        doneParents.push(built.okParents[p].parentSku);
      }
      subBatchesDone++;
      Logger.log('[' + fn + '] state=DONE subBatchId=' + subBatchId +
        ' status=' + outStatus +
        ' fileName=' + fileName + ' rows=' + built.rows.length +
        ' okParents=' + built.okParents.length);
    } catch (subErr) {
      Logger.log('[' + fn + '] state=FAILED subBatchId=' + subBatchId + ' ' + ((subErr && subErr.message) || subErr));
      amazonApprovalLv4AppendLog_(ss, {
        recordType: 'RUN',
        runId: runId,
        batchId: batchId,
        subBatchId: subBatchId,
        track: track,
        parentSkus: amazonApprovalLv4ParentSkuList_(sub),
        childCount: 0,
        status: 'FAILED',
        fileName: '',
        fileUrl: '',
        note: String((subErr && subErr.message) || subErr)
      });
      amazonApprovalLv4Mail_(
        '【Lv4 Amazon】サブバッチ失敗',
        'runId=' + runId + '\nsubBatchId=' + subBatchId + '\n' + ((subErr && subErr.message) || subErr)
      );
      throw subErr;
    }
  }

  if (!willResume) {
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_AMAZON_LV4_STATE_PROP);
    amazonApprovalLv4DeleteTriggers_();
    message = message || (
      (skipExport
        ? 'DRY_RUN完了（本番GENERATEDはSKIP_EXPORT=falseで再実行可）。'
        : 'GENERATED完了。次はローカルでPACKAGED→SC手動UP→21-③。') +
      'スキップ=' + amazonApprovalLv4SkipSummary_(resolved.skipped)
    );
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

/**
 * @param {string|null} raw
 * @return {string} A|B|BOTH|''
 */
function amazonApprovalLv4NormalizeTrack_(raw) {
  var t = String(raw || '').trim().toUpperCase();
  if (t === 'A' || t === 'B' || t === 'BOTH') return t;
  return '';
}

function amazonApprovalLv4ShippingTemplate_() {
  var v = PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE_PROP);
  v = String(v || '').trim();
  return v || APPROVAL_AMAZON_LV4_DEFAULT_SHIPPING;
}

function amazonApprovalLv4ParentsPerSub_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_PARENTS_PER_SUB_PROP));
  if (!n || isNaN(n) || n < 1) return 5;
  return Math.floor(n);
}

/** @return {{sheet:Sheet, values:Array, headerRowIdx:number, col:Object}} */
function amazonApprovalLv4LoadMasterContext_(ss) {
  var masterName = (typeof MASTER_SHEET_NAME !== 'undefined')
    ? MASTER_SHEET_NAME
    : ((typeof SHEET_NAME_MASTER !== 'undefined') ? SHEET_NAME_MASTER : '▼商品マスタ(人間作業用)');
  var sheet = ss.getSheetByName(masterName);
  if (!sheet) throw new Error('マスタシートが見つかりません: ' + masterName);
  var values = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  var limit = Math.min(values.length, 25);
  for (var r = 0; r < limit; r++) {
    var row = values[r] || [];
    if (row.indexOf('親SKU') !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error('マスタヘッダー（親SKU）が見つかりません');
  var headers = values[headerRowIdx];
  var col = {};
  for (var c = 0; c < headers.length; c++) {
    var key = String(headers[c] || '').trim();
    if (key && col[key] == null) col[key] = c;
  }
  if (col['親SKU'] == null || col['子SKU'] == null) {
    throw new Error('必須列がありません（親SKU / 子SKU）');
  }
  Logger.log('[amazonApprovalLv4LoadMasterContext_] headerRow1Based=' + (headerRowIdx + 1) +
    ' col販売価格amazon=' + (col['販売価格amazon'] != null ? col['販売価格amazon'] : 'MISSING') +
    ' col在庫数=' + (col['在庫数'] != null ? col['在庫数'] : 'MISSING'));
  return { sheet: sheet, values: values, headerRowIdx: headerRowIdx, col: col };
}

/**
 * 承認 amazon 行を親単位にまとめ、TRACK/在庫/免除でフィルタ。
 * @return {{parents:Array, skipped:Array}}
 */
function amazonApprovalLv4ResolveParents_(masterCtx, lines, track, doneParents, exemptions) {
  var doneMap = {};
  if (doneParents) {
    for (var d = 0; d < doneParents.length; d++) doneMap[String(doneParents[d])] = true;
  }
  var byParent = {};
  for (var i = 0; i < lines.length; i++) {
    var L = lines[i];
    if (String(L.mall) !== 'amazon' || String(L.lineStatus) !== 'APPROVED') continue;
    var parentSku = String(L.parentSku || '').trim();
    if (!parentSku) continue;
    if (!byParent[parentSku]) {
      byParent[parentSku] = { parentSku: parentSku, parentLine: null, childLines: [] };
    }
    var childSku = String(L.childSku || '').trim();
    if (!childSku) {
      byParent[parentSku].parentLine = L;
    } else {
      byParent[parentSku].childLines.push(L);
    }
  }

  var parents = [];
  var skipped = [];
  exemptions = exemptions || [];

  var keys = Object.keys(byParent);
  for (var k = 0; k < keys.length; k++) {
    var g = byParent[keys[k]];
    if (doneMap[g.parentSku]) continue;

    if (track === 'B' || track === 'BOTH') {
      // M1/B: 親承認＋子1件以上必須。単品親は対象外
      if (!g.parentLine) {
        skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_INCOMPLETE_VARIATION', detail: '親行未承認' });
        continue;
      }
      if (!g.childLines.length) {
        skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_INCOMPLETE_VARIATION', detail: '承認済み子0' });
        continue;
      }
    }

    var masterParent = amazonApprovalLv4FindMasterRow_(masterCtx, g.parentSku, '');
    if (!masterParent) {
      skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_ORPHAN', detail: '親がマスタに無い' });
      continue;
    }

    // 販売中スキップ: 親または承認済み子のいずれか在庫>0
    var stockHit = amazonApprovalLv4AnyInStock_(masterCtx, g);
    if (stockHit) {
      skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_IN_STOCK', detail: stockHit });
      continue;
    }

    var resolvedTrack = track;
    if (track === 'BOTH') {
      var asin = amazonApprovalLv4Cell_(masterCtx, masterParent.rowIndex0, 'ASINコード');
      if (asin) {
        resolvedTrack = 'A';
      } else {
        resolvedTrack = 'B';
      }
    }

    if (resolvedTrack === 'B') {
      var catResolved = amazonApprovalLv4ResolveCategory_(masterCtx, masterParent.rowIndex0, g.parentSku);
      var cat = catResolved.value;
      if (!cat) {
        skipped.push({
          parentSku: g.parentSku,
          reason: 'SKIPPED_NEED_HUMAN',
          detail: 'amazonカテゴリ空（GTIN証跡照合不可） sourceTried=' + catResolved.sourceTried
        });
        continue;
      }
      if (!amazonApprovalLv4HasGtinExemption_(exemptions, cat)) {
        skipped.push({
          parentSku: g.parentSku,
          reason: 'SKIPPED_GTIN_EXEMPTION',
          detail: 'カテゴリ未証跡:' + cat + ' via=' + catResolved.source
        });
        continue;
      }
      var mainImg = amazonApprovalLv4ResolveMainImageUrl_(masterCtx, masterParent.rowIndex0);
      if (!mainImg) {
        var anyChildImg = false;
        for (var ci0 = 0; ci0 < g.childLines.length; ci0++) {
          var mr0 = amazonApprovalLv4FindMasterRow_(masterCtx, g.parentSku, String(g.childLines[ci0].childSku || '').trim());
          if (mr0 && amazonApprovalLv4ResolveMainImageUrl_(masterCtx, mr0.rowIndex0)) {
            anyChildImg = true;
            break;
          }
        }
        if (!anyChildImg) {
          skipped.push({
            parentSku: g.parentSku,
            reason: 'SKIPPED_NEED_HUMAN',
            detail: 'メイン画像URL空'
          });
          continue;
        }
      }
    }

    if (resolvedTrack === 'A') {
      // ASIN: 親または承認済み子のいずれか（ASINコード／競合店ASIN／URL）
      var asinProbe = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, masterParent.rowIndex0);
      if (!asinProbe) {
        for (var ai = 0; ai < g.childLines.length && !asinProbe; ai++) {
          var mrA = amazonApprovalLv4FindMasterRow_(
            masterCtx, g.parentSku, String(g.childLines[ai].childSku || '').trim()
          );
          if (mrA) asinProbe = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, mrA.rowIndex0);
        }
      }
      if (!asinProbe) {
        skipped.push({
          parentSku: g.parentSku,
          reason: 'SKIPPED_NEED_HUMAN',
          detail: 'AトラックだがASIN無し（ASINコード／競合店ASINコード／競合URL）'
        });
        continue;
      }
      var gateMode = String(
        PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_BRAND_GATE_MODE_PROP) || ''
      ).trim().toLowerCase();
      if (gateMode !== 'manual_ok') {
        skipped.push({
          parentSku: g.parentSku,
          reason: 'SKIPPED_BRAND_GATE',
          detail: 'ブランド／出品制限を人間確認後、' + APPROVAL_AMAZON_LV4_BRAND_GATE_MODE_PROP +
            '=manual_ok を設定してください（現在=' + (gateMode || '空') + '）'
        });
        continue;
      }
    }

    // 子をマスタ照合（価格フォールバックより先に解決）
    var children = [];
    var orphanChild = false;
    for (var c = 0; c < g.childLines.length; c++) {
      var cl = g.childLines[c];
      var childSku2 = String(cl.childSku || '').trim();
      var mr = amazonApprovalLv4FindMasterRow_(masterCtx, g.parentSku, childSku2);
      if (!mr) {
        skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_ORPHAN', detail: '子無し:' + childSku2 });
        orphanChild = true;
        break;
      }
      children.push({ line: cl, rowIndex0: mr.rowIndex0, childSku: childSku2 });
    }
    if (orphanChild) continue;

    if (resolvedTrack === 'B' && !children.length) {
      skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_INCOMPLETE_VARIATION', detail: '子解決0' });
      continue;
    }

    var priceRes = amazonApprovalLv4ResolvePriceAmazon_(
      masterCtx, g.parentSku, masterParent.rowIndex0, children
    );
    if (!priceRes.value) {
      skipped.push({
        parentSku: g.parentSku,
        reason: 'SKIPPED_NEED_HUMAN',
        detail: '販売価格amazon不正 tried=' + priceRes.sourceTried
      });
      continue;
    }

    parents.push({
      parentSku: g.parentSku,
      parentLine: g.parentLine,
      parentRowIndex0: masterParent.rowIndex0,
      children: children,
      resolvedTrack: resolvedTrack,
      priceAmazon: priceRes.value,
      priceSource: priceRes.source
    });
  }

  return { parents: parents, skipped: skipped };
}

function amazonApprovalLv4AnyInStock_(masterCtx, group) {
  var iStock = masterCtx.col['在庫数'];
  if (iStock == null) return '';
  function stockOf(rowIndex0) {
    var v = masterCtx.values[rowIndex0][iStock];
    var n = (v === '' || v == null) ? 0 : Number(v);
    return (!isNaN(n) && n > 0) ? n : 0;
  }
  var parentRow = amazonApprovalLv4FindMasterRow_(masterCtx, group.parentSku, '');
  if (parentRow && stockOf(parentRow.rowIndex0) > 0) {
    return '親在庫>0';
  }
  for (var i = 0; i < group.childLines.length; i++) {
    var cs = String(group.childLines[i].childSku || '').trim();
    var cr = amazonApprovalLv4FindMasterRow_(masterCtx, group.parentSku, cs);
    if (cr && stockOf(cr.rowIndex0) > 0) return '子在庫>0:' + cs;
  }
  return '';
}

function amazonApprovalLv4FindMasterRow_(masterCtx, parentSku, childSku) {
  var iParent = masterCtx.col['親SKU'];
  var iChild = masterCtx.col['子SKU'];
  var wantChild = String(childSku || '').trim();
  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    var row = masterCtx.values[r] || [];
    if (String(row[iParent] != null ? row[iParent] : '').trim() !== parentSku) continue;
    var c = iChild != null ? String(row[iChild] != null ? row[iChild] : '').trim() : '';
    if (c === wantChild) return { rowIndex0: r };
  }
  return null;
}

/**
 * M2(A)用 ASIN 解決。優先: ASINコード → 競合店ASINコード → 競合URL系から /dp/ASIN
 * @return {string}
 */
function amazonApprovalLv4ResolveAsinForOffer_(masterCtx, rowIndex0) {
  var direct = amazonApprovalLv4Cell_(masterCtx, rowIndex0, 'ASINコード');
  if (amazonApprovalLv4LooksLikeAsin_(direct)) return direct.toUpperCase();
  var rival = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '競合店ASINコード');
  if (amazonApprovalLv4LooksLikeAsin_(rival)) return rival.toUpperCase();
  var urlCols = ['競合AmazonページURL', '競合URL', 'Amazon URL', '商品URL'];
  for (var i = 0; i < urlCols.length; i++) {
    var u = amazonApprovalLv4Cell_(masterCtx, rowIndex0, urlCols[i]);
    var fromUrl = amazonApprovalLv4AsinFromUrl_(u);
    if (fromUrl) return fromUrl;
  }
  return '';
}

function amazonApprovalLv4LooksLikeAsin_(s) {
  return /^B0[A-Z0-9]{8}$/i.test(String(s || '').trim());
}

function amazonApprovalLv4AsinFromUrl_(url) {
  var m = String(url || '').match(/\/(?:dp|gp\/product|product)\/(B0[A-Z0-9]{8})/i);
  return m ? String(m[1]).toUpperCase() : '';
}

function amazonApprovalLv4Cell_(masterCtx, rowIndex0, colName) {
  var idx = masterCtx.col[colName];
  if (idx == null) return '';
  return String(masterCtx.values[rowIndex0][idx] != null ? masterCtx.values[rowIndex0][idx] : '').trim();
}

/**
 * 販売価格amazon 解決。親が空／不正なら承認済み子を順に見る。
 * Logger: [Lv4Price]
 * @param {Array<{rowIndex0:number, childSku:string}>} children
 * @return {{value:string, source:string, sourceTried:string}}
 */
function amazonApprovalLv4ResolvePriceAmazon_(masterCtx, parentSku, parentRowIndex0, children) {
  var tried = [];
  var parentRaw = amazonApprovalLv4Cell_(masterCtx, parentRowIndex0, '販売価格amazon');
  var parentNum = Number(parentRaw);
  tried.push('parent(v=' + parentRaw + ')');
  if (parentRaw !== '' && !isNaN(parentNum) && parentNum > 0) {
    Logger.log('[Lv4Price] HIT parentSku=' + parentSku + ' source=parent value=' + parentRaw);
    return { value: String(parentRaw), source: 'parent', sourceTried: tried.join(';') };
  }
  children = children || [];
  for (var i = 0; i < children.length; i++) {
    var ch = children[i];
    var raw = amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, '販売価格amazon');
    var n = Number(raw);
    tried.push('child:' + String(ch.childSku || '') + '(v=' + raw + ')');
    if (raw !== '' && !isNaN(n) && n > 0) {
      Logger.log('[Lv4Price] HIT parentSku=' + parentSku +
        ' source=child:' + String(ch.childSku || '') + ' value=' + raw);
      return { value: String(raw), source: 'child:' + String(ch.childSku || ''), sourceTried: tried.join(';') };
    }
  }
  Logger.log('[Lv4Price] EMPTY parentSku=' + parentSku + ' tried=' + tried.join(';'));
  return { value: '', source: '', sourceTried: tried.join(';') };
}

/**
 * Amazonカテゴリ解決。優先順:
 * 1) amazon カテゴリー  2) amazonカテゴリー  3) カテゴリー（T列。見出しに改行※注記があっても可）
 * Logger に source / 列index / 値プレビュー / 失敗時の候補ヘッダを残す。
 * @return {{value:string, source:string, sourceTried:string}}
 */
function amazonApprovalLv4ResolveCategory_(masterCtx, rowIndex0, parentSku) {
  var preferred = ['amazon カテゴリー', 'amazonカテゴリー', 'カテゴリー'];
  var tried = [];
  var i;
  var name;
  var colIdx;
  var raw;
  var val;

  for (i = 0; i < preferred.length; i++) {
    name = preferred[i];
    colIdx = masterCtx.col[name];
    raw = colIdx != null
      ? String(masterCtx.values[rowIndex0][colIdx] != null ? masterCtx.values[rowIndex0][colIdx] : '')
      : '';
    val = String(raw || '').trim();
    tried.push(name + '(exists=' + (colIdx != null) + ',idx=' + (colIdx != null ? colIdx : -1) +
      ',len=' + raw.length + ',v=' + (val ? val.substring(0, 24) : '') + ')');
    if (val) {
      Logger.log('[Lv4Cat] HIT parentSku=' + parentSku + ' row1=' + (rowIndex0 + 1) +
        ' source=' + name + ' colIndex0=' + colIdx + ' valuePreview=' + val.substring(0, 40));
      return { value: val, source: name, sourceTried: tried.join(';') };
    }
  }

  // 見出しが「カテゴリー\n※…」や「カテゴリー ※…」でも拾う（amazon系を優先）
  var fuzzyAmazon = null;
  var fuzzyPlain = null;
  for (var k in masterCtx.col) {
    if (!Object.prototype.hasOwnProperty.call(masterCtx.col, k)) continue;
    var first = String(k).split(/\r?\n/)[0].trim();
    var compact = first.replace(/\s+/g, ' ');
    if (compact === 'amazon カテゴリー' || compact === 'amazonカテゴリー' ||
        compact.indexOf('amazon') === 0 && compact.indexOf('カテゴリ') >= 0) {
      if (!fuzzyAmazon) fuzzyAmazon = k;
    } else if (compact === 'カテゴリー' || compact.indexOf('カテゴリー') === 0) {
      if (!fuzzyPlain) fuzzyPlain = k;
    }
  }
  var fuzzyOrder = [];
  if (fuzzyAmazon) fuzzyOrder.push(fuzzyAmazon);
  if (fuzzyPlain) fuzzyOrder.push(fuzzyPlain);

  for (i = 0; i < fuzzyOrder.length; i++) {
    name = fuzzyOrder[i];
    colIdx = masterCtx.col[name];
    raw = String(masterCtx.values[rowIndex0][colIdx] != null ? masterCtx.values[rowIndex0][colIdx] : '');
    val = String(raw || '').trim();
    tried.push('fuzzy:' + name.replace(/\r?\n/g, '\\n') + '(idx=' + colIdx + ',len=' + raw.length +
      ',v=' + (val ? val.substring(0, 24) : '') + ')');
    if (val) {
      Logger.log('[Lv4Cat] HIT_FUZZY parentSku=' + parentSku + ' row1=' + (rowIndex0 + 1) +
        ' sourceHeader=' + JSON.stringify(name) + ' colIndex0=' + colIdx +
        ' valuePreview=' + val.substring(0, 40));
      return { value: val, source: 'fuzzy:' + String(name).split(/\r?\n/)[0], sourceTried: tried.join(';') };
    }
  }

  var catKeys = [];
  for (var k2 in masterCtx.col) {
    if (!Object.prototype.hasOwnProperty.call(masterCtx.col, k2)) continue;
    if (String(k2).indexOf('カテゴリ') >= 0) {
      catKeys.push(JSON.stringify(k2) + '@' + masterCtx.col[k2]);
    }
  }
  Logger.log('[Lv4Cat] EMPTY parentSku=' + parentSku + ' row1=' + (rowIndex0 + 1) +
    ' tried=' + tried.join(';') + ' headersWithカテゴリ=' + catKeys.join('|'));
  return { value: '', source: '', sourceTried: tried.join(';') };
}

function amazonApprovalLv4BuildSubBatches_(parents, perSub) {
  var out = [];
  for (var i = 0; i < parents.length; i += perSub) {
    out.push(parents.slice(i, i + perSub));
  }
  return out;
}

function amazonApprovalLv4ParentSkuList_(parents) {
  var a = [];
  for (var i = 0; i < parents.length; i++) a.push(parents[i].parentSku);
  return a.join(',');
}

/**
 * Resolve 済み親を行化。価格/画像不足はここでは起きない想定。
 * 万一不足時は skipped に入れ、okParents には含めない。
 * @return {{headers:Array<string>, rows:Array<Array>, childCount:number, okParents:Array, skipped:Array}}
 */
function amazonApprovalLv4BuildRows_(masterCtx, parents, trackProp, stockOut, shipping) {
  var headers = [
    'track', 'parentSku', 'childSku', 'sellerSku', 'manufacturerPart',
    'productName', 'brand', 'priceAmazon', 'inventory', 'gtin', 'asin',
    'mainImageUrl', 'subImageUrls', 'amazonCategory', 'setCount',
    'shippingTemplate', 'variationRole'
  ];
  var rows = [];
  var childCount = 0;
  var okParents = [];
  var skipped = [];

  for (var p = 0; p < parents.length; p++) {
    var parent = parents[p];
    var rTrack = parent.resolvedTrack || trackProp;
    var priceRes2 = parent.priceAmazon
      ? { value: parent.priceAmazon, source: parent.priceSource || 'cached', sourceTried: '' }
      : amazonApprovalLv4ResolvePriceAmazon_(
          masterCtx, parent.parentSku, parent.parentRowIndex0, parent.children || []
        );
    var price = priceRes2.value;
    if (!price) {
      skipped.push({
        parentSku: parent.parentSku,
        reason: 'SKIPPED_NEED_HUMAN',
        detail: 'Build時:販売価格amazon不正 tried=' + priceRes2.sourceTried
      });
      continue;
    }
    var catResolved2 = amazonApprovalLv4ResolveCategory_(masterCtx, parent.parentRowIndex0, parent.parentSku);
    var cat = catResolved2.value;
    var catalogName = amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, 'オリジナルカタログ商品名');
    var asin = amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, 'ASINコード');
    var mainImg = amazonApprovalLv4ResolveMainImageUrl_(masterCtx, parent.parentRowIndex0);
    // Resolve と同様: 親空なら承認済み子の MAIN URL を親行用にフォールバック
    if (!mainImg && parent.children && parent.children.length) {
      var ciFb;
      for (ciFb = 0; ciFb < parent.children.length; ciFb++) {
        var chFb = parent.children[ciFb];
        mainImg = amazonApprovalLv4ResolveMainImageUrl_(masterCtx, chFb.rowIndex0);
        if (mainImg) {
          Logger.log('[Lv4Build] parentMainFallback parentSku=' + parent.parentSku +
            ' fromChild=' + String(chFb.childSku || ''));
          break;
        }
      }
    }

    if (rTrack === 'B') {
      if (!cat) {
        skipped.push({
          parentSku: parent.parentSku,
          reason: 'SKIPPED_NEED_HUMAN',
          detail: 'Build時:カテゴリ空 tried=' + catResolved2.sourceTried
        });
        continue;
      }
      if (!mainImg) {
        skipped.push({
          parentSku: parent.parentSku,
          reason: 'SKIPPED_NEED_HUMAN',
          detail: 'Build時:メイン画像空（親・子とも Amazon MAIN URL／楽天メイン画像1 なし）'
        });
        continue;
      }
      rows.push([
        'B', parent.parentSku, '', parent.parentSku, '',
        catalogName || amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, '商品名'),
        APPROVAL_AMAZON_LV4_BRAND, price, stockOut, '', '',
        mainImg, amazonApprovalLv4SubImages_(masterCtx, parent.parentRowIndex0),
        cat, amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, 'A.セット商品数'),
        shipping, 'parent'
      ]);
      for (var c = 0; c < parent.children.length; c++) {
        var ch = parent.children[c];
        var childMain = amazonApprovalLv4ResolveMainImageUrl_(masterCtx, ch.rowIndex0) || mainImg;
        var mfr = amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, 'メーカー品番');
        if (!mfr) mfr = amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, '型番');
        if (!mfr) mfr = ch.childSku;
        var childPrice = amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, '販売価格amazon') || price;
        var childName = catalogName ||
          amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, '商品名') ||
          amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, '商品名');
        rows.push([
          'B', parent.parentSku, ch.childSku, ch.childSku, mfr,
          childName, APPROVAL_AMAZON_LV4_BRAND, childPrice, stockOut, '', '',
          childMain, amazonApprovalLv4SubImages_(masterCtx, ch.rowIndex0),
          cat, amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, 'A.セット商品数'),
          shipping, 'child'
        ]);
        childCount++;
      }
      okParents.push(parent);
    } else {
      var targets = parent.children.length
        ? parent.children
        : [{ childSku: parent.parentSku, rowIndex0: parent.parentRowIndex0, line: parent.parentLine }];
      for (var a = 0; a < targets.length; a++) {
        var t = targets[a];
        var sku = t.childSku || parent.parentSku;
        var mfrA = amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, 'メーカー品番');
        if (!mfrA) mfrA = amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, '型番');
        if (!mfrA) mfrA = sku;
        var jan = amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, '商品コード(JANコード等)');
        if (!jan) jan = amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, 'JANコード');
        var asinOffer = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, t.rowIndex0) ||
          amazonApprovalLv4ResolveAsinForOffer_(masterCtx, parent.parentRowIndex0) ||
          asin;
        if (!asinOffer) {
          skipped.push({
            parentSku: parent.parentSku,
            reason: 'SKIPPED_NEED_HUMAN',
            detail: 'Build時:Aトラック ASIN無し sku=' + sku
          });
          continue;
        }
        rows.push([
          'A', parent.parentSku, t.childSku || '', sku, mfrA,
          amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, '商品名'),
          '', amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, '販売価格amazon') || price,
          stockOut, jan, asinOffer,
          amazonApprovalLv4ResolveMainImageUrl_(masterCtx, t.rowIndex0),
          amazonApprovalLv4SubImages_(masterCtx, t.rowIndex0),
          cat, amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, 'A.セット商品数'),
          shipping, 'offer'
        ]);
        childCount++;
      }
      okParents.push(parent);
    }
  }

  return {
    headers: headers,
    rows: rows,
    childCount: childCount,
    okParents: okParents,
    skipped: skipped
  };
}

function amazonApprovalLv4SubImages_(masterCtx, rowIndex0) {
  var amazonPt = amazonApprovalLv4Cell_(masterCtx, rowIndex0, 'Amazon PT URL');
  if (amazonPt) return amazonPt;
  var urls = [];
  for (var i = 1; i <= 8; i++) {
    var u = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '楽天サブ画像' + i);
    if (u) urls.push(u);
  }
  return urls.join('|');
}

/** U4: Amazon MAIN URL があれば優先、なければ楽天メイン画像1 */
function amazonApprovalLv4ResolveMainImageUrl_(masterCtx, rowIndex0) {
  var amz = amazonApprovalLv4Cell_(masterCtx, rowIndex0, 'Amazon MAIN URL');
  if (amz) return amz;
  return amazonApprovalLv4Cell_(masterCtx, rowIndex0, '楽天メイン画像1') || '';
}

function amazonApprovalLv4RowsToCsv_(headers, rows) {
  function esc(v) {
    var s = String(v == null ? '' : v);
    if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }
  var lines = [headers.map(esc).join(',')];
  for (var i = 0; i < rows.length; i++) {
    lines.push(rows[i].map(esc).join(','));
  }
  return lines.join('\r\n');
}

function amazonApprovalLv4EnsureLogSheet_(ss) {
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(APPROVAL_AMAZON_LV4_LOG_SHEET);
    sh.getRange(1, 1, 1, APPROVAL_AMAZON_LV4_LOG_HEADERS.length).setValues([APPROVAL_AMAZON_LV4_LOG_HEADERS]);
    sh.appendRow([
      'EXEMPTION', '', '', '', '', '', '',
      '', '', '', '', '', '',
      '（この行を編集: カテゴリ実値または * ・承認日・証跡URLをすべて記入。brand=ノーブランド品）',
      '（カテゴリ or *）', APPROVAL_AMAZON_LV4_BRAND, '', ''
    ]);
    return sh;
  }
  var h = String(sh.getRange(1, 1).getValue() || '');
  if (h !== 'recordType') {
    // 証跡・履歴消失防止のため clear しない
    throw new Error(
      'シート「' + APPROVAL_AMAZON_LV4_LOG_SHEET + '」の1行目が recordType ではありません。' +
      '手動でヘッダーを直すか、別名退避してから再実行してください（自動clearはしません）。'
    );
  }
  return sh;
}

function amazonApprovalLv4LoadExemptions_(ss) {
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[0]) !== 'EXEMPTION') continue;
    out.push({
      category: String(row[14] || '').trim(),
      brand: String(row[15] || '').trim(),
      exemptionDate: String(row[16] || '').trim(),
      evidenceUrl: String(row[17] || '').trim()
    });
  }
  return out;
}

/**
 * GTIN免除: カテゴリ空は不合格。ブランドは「ノーブランド品」完全一致。
 * 承認日・証跡URL必須。カテゴリは完全一致または * のみ。
 */
function amazonApprovalLv4HasGtinExemption_(exemptions, category) {
  var cat = String(category || '').trim();
  if (!cat) return false;
  for (var i = 0; i < exemptions.length; i++) {
    var e = exemptions[i];
    if (String(e.brand || '').trim() !== APPROVAL_AMAZON_LV4_BRAND) continue;
    if (!e.exemptionDate || !e.evidenceUrl) continue;
    if (e.category === '（カテゴリ or *）') continue;
    if (e.category === '*') return true;
    if (e.category && e.category === cat) return true;
  }
  return false;
}

/**
 * 同一 batchId+track で、意味ある状態の最新が GENERATED/UPLOADED_OK/PACKAGED の親 → ブロック。
 * SKIPPED_* / DRY_RUN / FAILED は latest 判定に使わない（汚染防止）。
 * UPLOAD_FAILED が最新ならブロックしない。
 * @return {Object<string,string>} parentSku -> status
 */
function amazonApprovalLv4LoadBlockedParents_(ss, batchId, track) {
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  var latest = {};
  if (!sh || sh.getLastRow() < 2) return {};
  var data = sh.getDataRange().getValues();
  var wantBatch = String(batchId || '');
  var wantTrack = String(track || '');
  var meaningful = {
    GENERATED: true,
    UPLOADED_OK: true,
    PACKAGED: true,
    UPLOAD_FAILED: true
  };
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[0]) !== 'RUN') continue;
    if (wantBatch && String(row[2]) !== wantBatch) continue;
    var rowTrack = String(row[4] || '');
    if (wantTrack && rowTrack && rowTrack !== wantTrack && wantTrack !== 'BOTH') continue;
    var status = String(row[7] || '');
    if (!meaningful[status]) continue;
    var parents = String(row[5] || '').split(',');
    for (var p = 0; p < parents.length; p++) {
      var ps = String(parents[p] || '').trim();
      if (!ps) continue;
      latest[ps] = status;
    }
  }
  var blocked = {};
  var keys = Object.keys(latest);
  for (var k = 0; k < keys.length; k++) {
    var st = latest[keys[k]];
    if (st === 'GENERATED' || st === 'UPLOADED_OK' || st === 'PACKAGED') {
      blocked[keys[k]] = st;
    }
  }
  return blocked;
}

/**
 * スキップ行は recordType=SKIP（冪等 latest を汚染しない）。
 */
function amazonApprovalLv4LogSkippedParents_(ss, runId, batchId, track, skipped) {
  if (!skipped || !skipped.length) return;
  for (var i = 0; i < skipped.length; i++) {
    var s = skipped[i];
    amazonApprovalLv4AppendLog_(ss, {
      recordType: 'SKIP',
      runId: runId,
      batchId: batchId,
      subBatchId: '',
      track: track,
      parentSkus: s.parentSku || '',
      childCount: 0,
      status: s.reason || 'SKIPPED_NEED_HUMAN',
      fileName: '',
      fileUrl: '',
      note: s.detail || ''
    });
  }
}

/**
 * 同一 batchId + trackTag の subBatchId 連番の次番号（1始まり・単調増加）。
 */
function amazonApprovalLv4NextSubBatchSeq_(ss, batchId, trackTag, resumeState) {
  var maxSeq = 0;
  if (resumeState && resumeState.nextSubSeq) {
    maxSeq = Math.max(maxSeq, Number(resumeState.nextSubSeq) - 1);
  }
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues();
    var prefix = String(batchId) + '_' + String(trackTag);
    var re = new RegExp('^' + prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\d+)$');
    for (var r = 1; r < data.length; r++) {
      var id = String(data[r][3] || '');
      var m = id.match(re);
      if (m) {
        var n = Number(m[1]);
        if (!isNaN(n) && n > maxSeq) maxSeq = n;
      }
    }
  }
  return maxSeq + 1;
}

function amazonApprovalLv4ArchiveSameName_(folder, fileName) {
  var existing = folder.getFilesByName(fileName);
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  while (existing.hasNext()) {
    var old = existing.next();
    try {
      old.setName(fileName.replace(/(\.[^.]+)$/, '_archived_' + stamp + '$1'));
    } catch (e) {
      throw new Error(
        '同名ファイルのアーカイブ改名に失敗（trashしない）: ' + fileName + ' / ' +
        ((e && e.message) || e)
      );
    }
  }
}

function amazonApprovalLv4AppendLog_(ss, obj) {
  var sh = amazonApprovalLv4EnsureLogSheet_(ss);
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
  var status = String(obj.status || '');
  var generatedAt = (status === 'GENERATED') ? now : '';
  var uploadedAt = (status === 'UPLOADED_OK') ? now : '';
  sh.appendRow([
    obj.recordType || 'RUN',
    obj.runId || '',
    obj.batchId || '',
    obj.subBatchId || '',
    obj.track || '',
    obj.parentSkus || '',
    obj.childCount != null ? obj.childCount : '',
    status,
    obj.fileName || '',
    obj.fileUrl || '',
    generatedAt,
    '',
    uploadedAt,
    obj.note || '',
    obj.category || '',
    obj.brand || '',
    obj.exemptionDate || '',
    obj.evidenceUrl || ''
  ]);
}

/**
 * 追記専用。既存行は上書きしない。
 * UPLOADED_OK は同一 subBatchId の GENERATED が無いと拒否。
 */
function amazonApprovalLv4MarkStatus_(ss, subBatchId, status, note) {
  var sh = amazonApprovalLv4EnsureLogSheet_(ss);
  var data = sh.getDataRange().getValues();
  var who = '';
  try {
    who = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (e) {}
  var hasGenerated = false;
  var parentSkus = '';
  var batchId = '';
  var track = '';
  var runId = '';
  var fileName = '';
  var fileUrl = '';
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) !== 'RUN') continue;
    if (String(data[r][3]) !== subBatchId) continue;
    if (String(data[r][7]) === 'GENERATED') {
      hasGenerated = true;
      parentSkus = String(data[r][5] || '');
      batchId = String(data[r][2] || '');
      track = String(data[r][4] || '');
      runId = String(data[r][1] || '');
      fileName = String(data[r][8] || '');
      fileUrl = String(data[r][9] || '');
    }
  }
  if (status === 'UPLOADED_OK' && !hasGenerated) {
    throw new Error('subBatchId=' + subBatchId + ' の GENERATED 行がありません。任意文字列では完了記録できません。');
  }
  if (status === 'UPLOAD_FAILED' && !hasGenerated) {
    Logger.log('[amazonApprovalLv4MarkStatus_] WARN no GENERATED for ' + subBatchId + ' — append anyway');
  }
  amazonApprovalLv4AppendLog_(ss, {
    recordType: 'RUN',
    runId: runId,
    batchId: batchId,
    subBatchId: subBatchId,
    track: track,
    parentSkus: parentSkus,
    childCount: '',
    status: status,
    fileName: fileName,
    fileUrl: fileUrl,
    note: (note || '') + ' by=' + who + ' (append-only)'
  });
  return 1;
}

function amazonApprovalLv4GetOrCreateFolder_(ss) {
  var folderId = PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_FOLDER_ID_PROP);
  if (folderId) {
    try {
      return DriveApp.getFolderById(String(folderId).trim());
    } catch (e) {
      Logger.log('[amazonApprovalLv4GetOrCreateFolder_] invalid folder id, fallback create');
    }
  }
  var name = 'Lv4_Amazon_GENERATED';
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function amazonApprovalLv4SkipSummary_(skipped) {
  if (!skipped || !skipped.length) return '0';
  var counts = {};
  for (var i = 0; i < skipped.length; i++) {
    var r = skipped[i].reason || 'OTHER';
    counts[r] = (counts[r] || 0) + 1;
  }
  var parts = [];
  var keys = Object.keys(counts);
  for (var k = 0; k < keys.length; k++) {
    parts.push(keys[k] + '=' + counts[keys[k]]);
  }
  return parts.join(',');
}

function amazonApprovalLv4SaveState_(state) {
  PropertiesService.getScriptProperties().setProperty(APPROVAL_AMAZON_LV4_STATE_PROP, JSON.stringify(state));
}

function amazonApprovalLv4SetTrigger_() {
  amazonApprovalLv4DeleteTriggers_();
  ScriptApp.newTrigger(APPROVAL_AMAZON_LV4_TRIGGER_FN).timeBased().after(60 * 1000).create();
}

function amazonApprovalLv4DeleteTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === APPROVAL_AMAZON_LV4_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function amazonApprovalLv4Mail_(subject, body) {
  try {
    var email = Session.getEffectiveUser().getEmail();
    if (!email) return;
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('[amazonApprovalLv4Mail_] ' + ((e && e.message) || e));
  }
}
