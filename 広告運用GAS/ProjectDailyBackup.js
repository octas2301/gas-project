/**
 * プロジェクト別 Drive 自動バックアップ（広告運用）
 * 正本: docs/org/PROJECT_DAILY_BACKUP_REQUIREMENTS.md
 *
 * 切戻し: ADS_PROJECT_DAILY_BACKUP_ENABLED=false
 * トリガー: adsRunProjectDailyBackupFromTrigger（毎日 8:00 Asia/Tokyo）
 */

var ADS_PDB_ENABLED_PROP_ = 'ADS_PROJECT_DAILY_BACKUP_ENABLED';
var ADS_PDB_ROOT_FOLDER_ID_PROP_ = 'ADS_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID';
var ADS_PDB_PJ_FOLDER_ID_PROP_ = 'ADS_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID';
var ADS_PDB_TARGET_SS_ID_PROP_ = 'ADS_PROJECT_DAILY_BACKUP_TARGET_SS_ID';
var ADS_PDB_RETENTION_DAYS_PROP_ = 'ADS_PROJECT_DAILY_BACKUP_RETENTION_DAYS';
var ADS_PDB_SHRINK_RATIO_PROP_ = 'ADS_PROJECT_DAILY_BACKUP_SHRINK_RATIO';
var ADS_PDB_MAIL_TO_PROP_ = 'ADS_PROJECT_DAILY_BACKUP_MAIL_TO';
var ADS_PDB_REQUIRED_SHEETS_PROP_ = 'ADS_PROJECT_DAILY_BACKUP_REQUIRED_SHEETS';

var ADS_PDB_ROOT_FOLDER_NAME_ = '99.バックアップ';
var ADS_PDB_PJ_FOLDER_NAME_ = '広告';
var ADS_PDB_COPY_PREFIX_ = '広告_backup_';
var ADS_PDB_RETENTION_DEFAULT_ = 14;
var ADS_PDB_SHRINK_DEFAULT_ = 0.8;
var ADS_PDB_TRIGGER_FN_ = 'adsRunProjectDailyBackupFromTrigger';
/** 分析本線で常にある想定。無いときは Property で差し替え可 */
var ADS_PDB_REQUIRED_SHEETS_DEFAULT_ = '元データ_Amazon|履歴データ';
var ADS_PDB_TZ_ = 'Asia/Tokyo';
/** 正スプシ既定（Property 空のときの Active／この定数） */
var ADS_PDB_TARGET_SS_DEFAULT_ = '17jwaao79ApQrHzA27BPVIl1sQ9G2tVAqWPkVvrGGelo';

function adsPdbEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(ADS_PDB_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toLowerCase();
  return !(s === 'false' || s === '0' || s === 'off' || s === 'no');
}

function adsPdbRetentionDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(ADS_PDB_RETENTION_DAYS_PROP_) || '');
  if (!(n > 0)) n = ADS_PDB_RETENTION_DEFAULT_;
  return Math.floor(n);
}

function adsPdbShrinkRatio_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(ADS_PDB_SHRINK_RATIO_PROP_) || '');
  if (!(n > 0) || n >= 1) n = ADS_PDB_SHRINK_DEFAULT_;
  return n;
}

/** 昨日8:00〜今日8:00（実行日の朝8時を終点）。 */
function adsPdbWindowBounds_() {
  var now = new Date();
  var todayYmd = Utilities.formatDate(now, ADS_PDB_TZ_, 'yyyy-MM-dd');
  var end = Utilities.parseDate(todayYmd + ' 08:00:00', ADS_PDB_TZ_, 'yyyy-MM-dd HH:mm:ss');
  var start = new Date(end.getTime() - 24 * 60 * 60 * 1000);
  return { start: start, end: end, ymd: todayYmd };
}

function adsPdbFileUpdatedInWindow_(file, win) {
  if (!file || !win) return false;
  var t = file.getLastUpdated();
  if (!t) return false;
  var ms = t.getTime();
  return ms >= win.start.getTime() && ms < win.end.getTime() + 60 * 1000;
}

function adsPdbResolveFolderByName_(parent, name) {
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return null;
}

function adsPdbEnsurePjFolder_() {
  var props = PropertiesService.getScriptProperties();
  var pjId = String(props.getProperty(ADS_PDB_PJ_FOLDER_ID_PROP_) || '').trim();
  if (pjId) {
    try {
      return DriveApp.getFolderById(pjId);
    } catch (e1) {
      throw new Error('PJフォルダID不正: ' + ADS_PDB_PJ_FOLDER_ID_PROP_);
    }
  }
  var rootId = String(props.getProperty(ADS_PDB_ROOT_FOLDER_ID_PROP_) || '').trim();
  var root = null;
  if (rootId) {
    try {
      root = DriveApp.getFolderById(rootId);
    } catch (e2) {
      throw new Error('親フォルダID不正: ' + ADS_PDB_ROOT_FOLDER_ID_PROP_);
    }
  } else {
    root = adsPdbResolveFolderByName_(null, ADS_PDB_ROOT_FOLDER_NAME_);
    if (!root) throw new Error('Driveに「' + ADS_PDB_ROOT_FOLDER_NAME_ + '」が無い。作成するか ROOT_FOLDER_ID を設定');
  }
  var pj = adsPdbResolveFolderByName_(root, ADS_PDB_PJ_FOLDER_NAME_);
  if (!pj) {
    pj = root.createFolder(ADS_PDB_PJ_FOLDER_NAME_);
    Logger.log('[広告日次バックアップ] created pj folder=%s', pj.getId());
  }
  return pj;
}

function adsPdbTargetFile_() {
  var id = String(PropertiesService.getScriptProperties().getProperty(ADS_PDB_TARGET_SS_ID_PROP_) || '').trim();
  if (!id) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) id = ss.getId();
    } catch (eAct) {
      /* clasp / trigger */
    }
  }
  if (!id) id = ADS_PDB_TARGET_SS_DEFAULT_;
  return DriveApp.getFileById(id);
}

function adsPdbRequiredSheetNames_() {
  var raw = String(PropertiesService.getScriptProperties().getProperty(ADS_PDB_REQUIRED_SHEETS_PROP_) || '').trim();
  if (!raw) raw = ADS_PDB_REQUIRED_SHEETS_DEFAULT_;
  return raw
    .split('|')
    .map(function (s) {
      return String(s || '').trim();
    })
    .filter(function (s) {
      return !!s;
    });
}

/**
 * @return {{ok:boolean, reason:string}}
 */
function adsPdbCorruptSuspect_(file, folder) {
  try {
    var ss = SpreadsheetApp.openById(file.getId());
    var need = adsPdbRequiredSheetNames_();
    var i;
    for (i = 0; i < need.length; i++) {
      if (!ss.getSheetByName(need[i])) {
        return { ok: false, reason: 'required_sheet_missing=' + need[i] };
      }
    }
  } catch (eOpen) {
    return {
      ok: false,
      reason: 'open_failed=' + String((eOpen && eOpen.message) || eOpen).slice(0, 120)
    };
  }
  var prev = adsPdbLatestGoodCopy_(folder, file.getId());
  if (prev) {
    var ratio = adsPdbShrinkRatio_();
    var liveSize = Number(file.getSize()) || 0;
    var prevSize = Number(prev.getSize()) || 0;
    if (prevSize >= 50000 && liveSize < prevSize * ratio) {
      return {
        ok: false,
        reason: 'shrink live=' + liveSize + ' prev=' + prevSize + ' ratio<' + ratio
      };
    }
  }
  return { ok: true, reason: '' };
}

function adsPdbLatestGoodCopy_(folder, liveId) {
  if (!folder) return null;
  var it = folder.getFiles();
  var best = null;
  var bestName = '';
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(ADS_PDB_COPY_PREFIX_) !== 0) continue;
    if (f.isTrashed()) continue;
    if (!best || name > bestName) {
      best = f;
      bestName = name;
    }
  }
  return best;
}

function adsPdbPurgeOldCopies_(folder, liveId, retentionDays) {
  if (!folder || !(retentionDays > 0)) return 0;
  var cutoff = new Date().getTime() - retentionDays * 24 * 60 * 60 * 1000;
  var it = folder.getFiles();
  var n = 0;
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(ADS_PDB_COPY_PREFIX_) !== 0) continue;
    var created = f.getDateCreated();
    if (created && created.getTime() < cutoff) {
      f.setTrashed(true);
      n++;
    }
  }
  return n;
}

function adsPdbAppendLog_(state, detail) {
  try {
    var file = adsPdbTargetFile_();
    var ss = SpreadsheetApp.openById(file.getId());
    var sh = ss.getSheetByName('履歴データ');
    if (!sh) return;
    sh.appendRow([
      new Date(),
      '日次バックアップ',
      state,
      String(detail || '').slice(0, 900)
    ]);
  } catch (eLog) {
    Logger.log('[広告日次バックアップ] log_fail %s', String((eLog && eLog.message) || eLog).slice(0, 120));
  }
}

function adsPdbReportHuman_(subject, body) {
  adsPdbAppendLog_('REPORT', subject + ' | ' + body);
  Logger.log('[広告日次バックアップ] REPORT %s %s', subject, body);
  var to = String(PropertiesService.getScriptProperties().getProperty(ADS_PDB_MAIL_TO_PROP_) || '').trim();
  if (!to) return;
  try {
    MailApp.sendEmail(to, '[広告バックアップ] ' + subject, body);
  } catch (eMail) {
    Logger.log(
      '[広告日次バックアップ] mail_fail %s',
      String((eMail && eMail.message) || eMail).slice(0, 160)
    );
    adsPdbAppendLog_('MAIL_FAIL', String((eMail && eMail.message) || eMail).slice(0, 200));
  }
}

/**
 * @return {{ok:boolean, action:string, detail:string}}
 */
function adsRunProjectDailyBackupCore_() {
  if (!adsPdbEnabled_()) {
    return { ok: true, action: 'disabled', detail: ADS_PDB_ENABLED_PROP_ + '=false' };
  }
  var win = adsPdbWindowBounds_();
  var file = adsPdbTargetFile_();
  var liveId = file.getId();
  if (!adsPdbFileUpdatedInWindow_(file, win)) {
    var skipDetail =
      'no_update_in_window lastUpdated=' +
      Utilities.formatDate(file.getLastUpdated(), ADS_PDB_TZ_, "yyyy-MM-dd'T'HH:mm:ss") +
      ' window=' +
      Utilities.formatDate(win.start, ADS_PDB_TZ_, "yyyy-MM-dd'T'HH:mm") +
      '..' +
      Utilities.formatDate(win.end, ADS_PDB_TZ_, "yyyy-MM-dd'T'HH:mm");
    adsPdbAppendLog_('SKIP', skipDetail);
    return { ok: true, action: 'skip_no_update', detail: skipDetail };
  }
  var folder = adsPdbEnsurePjFolder_();
  var corrupt = adsPdbCorruptSuspect_(file, folder);
  if (!corrupt.ok) {
    var msg = 'corrupt_skip ' + corrupt.reason + ' liveId=' + liveId;
    adsPdbReportHuman_('破損疑いのためコピーしませんでした', msg);
    return { ok: true, action: 'skip_corrupt', detail: msg };
  }
  var copyName = ADS_PDB_COPY_PREFIX_ + win.ymd;
  var existing = folder.getFilesByName(copyName);
  if (existing.hasNext()) {
    var existDetail = 'already_exists name=' + copyName;
    adsPdbAppendLog_('SKIP', existDetail);
    return { ok: true, action: 'skip_exists', detail: existDetail };
  }
  var copy = file.makeCopy(copyName, folder);
  var purged = adsPdbPurgeOldCopies_(folder, liveId, adsPdbRetentionDays_());
  var doneDetail =
    'copied id=' +
    copy.getId() +
    ' name=' +
    copyName +
    ' purged=' +
    purged +
    ' folder=' +
    folder.getId();
  adsPdbAppendLog_('DONE', doneDetail);
  Logger.log('[広告日次バックアップ] %s', doneDetail);
  return { ok: true, action: 'copied', detail: doneDetail };
}

/** 時間主導: 毎日8時。 */
function adsRunProjectDailyBackupFromTrigger() {
  try {
    var r = adsRunProjectDailyBackupCore_();
    Logger.log('[広告日次バックアップ] trigger action=%s %s', r.action, r.detail);
  } catch (e) {
    var msg = String((e && e.message) || e).slice(0, 400);
    adsPdbReportHuman_('日次バックアップ失敗', msg);
    throw e;
  }
}

/** 手動: 今すぐ（更新窓あり）。 */
function menuAdsProjectDailyBackupRunNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = adsRunProjectDailyBackupCore_();
    ui.alert('広告 日次バックアップ', r.action + '\n' + r.detail, ui.ButtonSet.OK);
  } catch (e) {
    var msg = String((e && e.message) || e);
    adsPdbReportHuman_('手動バックアップ失敗', msg);
    ui.alert('失敗', msg, ui.ButtonSet.OK);
  }
}

/** 検収用: 更新窓を無視して1回コピー。 */
function menuAdsProjectDailyBackupForceCopy() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    '強制バックアップ',
    '更新窓を無視して今すぐコピーします（破損疑いならコピーしません）。実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;
  try {
    if (!adsPdbEnabled_()) {
      ui.alert('OFF', ADS_PDB_ENABLED_PROP_ + '=false', ui.ButtonSet.OK);
      return;
    }
    var file = adsPdbTargetFile_();
    var folder = adsPdbEnsurePjFolder_();
    var corrupt = adsPdbCorruptSuspect_(file, folder);
    if (!corrupt.ok) {
      adsPdbReportHuman_('強制コピー中止（破損疑い）', corrupt.reason);
      ui.alert('中止', corrupt.reason, ui.ButtonSet.OK);
      return;
    }
    var ymd = Utilities.formatDate(new Date(), ADS_PDB_TZ_, 'yyyyMMdd_HHmmss');
    var copyName = ADS_PDB_COPY_PREFIX_ + 'force_' + ymd;
    var copy = file.makeCopy(copyName, folder);
    var detail = 'force id=' + copy.getId() + ' name=' + copyName;
    adsPdbAppendLog_('DONE', detail);
    ui.alert('完了', detail, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function menuAdsProjectDailyBackupInstallTrigger() {
  var ui = SpreadsheetApp.getUi();
  adsPdbRemoveTriggers_();
  ScriptApp.newTrigger(ADS_PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(ADS_PDB_TZ_).create();
  ui.alert('広告 日次バックアップ', '毎朝8時トリガーを設置しました（' + ADS_PDB_TRIGGER_FN_ + '）', ui.ButtonSet.OK);
}

function menuAdsProjectDailyBackupRemoveTrigger() {
  var ui = SpreadsheetApp.getUi();
  var n = adsPdbRemoveTriggers_();
  ui.alert('広告 日次バックアップ', 'トリガー削除 n=' + n, ui.ButtonSet.OK);
}

/**
 * 初回セットアップ（clasp run / エディタから1回）。
 * PJフォルダID・対象SS・ROOT を Property に書き、トリガーを設置する。
 * @param {string=} pjFolderId
 */
function adsSetupProjectDailyBackup_(pjFolderId) {
  var props = PropertiesService.getScriptProperties();
  var rootId = '1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE';
  var target = ADS_PDB_TARGET_SS_DEFAULT_;
  var pj = String(pjFolderId || '').trim();
  if (!pj) {
    var folder = adsPdbEnsurePjFolder_();
    pj = folder.getId();
  }
  props.setProperties(
    {
      ADS_PROJECT_DAILY_BACKUP_ENABLED: 'true',
      ADS_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID: rootId,
      ADS_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID: pj,
      ADS_PROJECT_DAILY_BACKUP_TARGET_SS_ID: target
    },
    false
  );
  adsPdbRemoveTriggers_();
  ScriptApp.newTrigger(ADS_PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(ADS_PDB_TZ_).create();
  var out = {
    ok: true,
    rootId: rootId,
    pjFolderId: pj,
    targetSsId: target,
    trigger: ADS_PDB_TRIGGER_FN_
  };
  Logger.log('[広告日次バックアップ] setup %s', JSON.stringify(out));
  return out;
}

/** UI: 初回 Property＋トリガー。 */
function menuAdsProjectDailyBackupSetup() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = adsSetupProjectDailyBackup_();
    ui.alert(
      '広告 日次バックアップ セットアップ',
      'Property と毎朝8時トリガーを設定しました。\npjFolder=' + r.pjFolderId + '\nss=' + r.targetSsId,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

/**
 * clasp run 用: 更新窓無視の強制コピー（UIなし）。
 * @return {{ok:boolean, action:string, detail:string}}
 */
function adsForceProjectDailyBackupNoUi() {
  if (!adsPdbEnabled_()) {
    return { ok: true, action: 'disabled', detail: ADS_PDB_ENABLED_PROP_ + '=false' };
  }
  var file = adsPdbTargetFile_();
  var folder = adsPdbEnsurePjFolder_();
  var corrupt = adsPdbCorruptSuspect_(file, folder);
  if (!corrupt.ok) {
    adsPdbReportHuman_('強制コピー中止（破損疑い）', corrupt.reason);
    return { ok: true, action: 'skip_corrupt', detail: corrupt.reason };
  }
  var ymd = Utilities.formatDate(new Date(), ADS_PDB_TZ_, 'yyyyMMdd_HHmmss');
  var copyName = ADS_PDB_COPY_PREFIX_ + 'force_' + ymd;
  var copy = file.makeCopy(copyName, folder);
  var detail = 'force id=' + copy.getId() + ' name=' + copyName + ' folder=' + folder.getId();
  adsPdbAppendLog_('DONE', detail);
  Logger.log('[広告日次バックアップ] %s', detail);
  return { ok: true, action: 'copied', detail: detail };
}

function adsPdbRemoveTriggers_() {
  var list = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === ADS_PDB_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(list[i]);
      n++;
    }
  }
  return n;
}
