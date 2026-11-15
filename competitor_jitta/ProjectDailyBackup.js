/**
 * プロジェクト別 Drive 自動バックアップ（競合Jitta＝競合DBスプシ）
 * 正本: docs/org/PROJECT_DAILY_BACKUP_REQUIREMENTS.md
 *
 * 対象は競合DB（①の競合DB控えと同じ正本。フォルダは PJ 別）。
 * 切戻し: CJ_PROJECT_DAILY_BACKUP_ENABLED=false
 * トリガー: cjRunProjectDailyBackupFromTrigger（毎日 8:00 Asia/Tokyo）
 */

var CJ_PDB_ENABLED_PROP_ = 'CJ_PROJECT_DAILY_BACKUP_ENABLED';
var CJ_PDB_ROOT_FOLDER_ID_PROP_ = 'CJ_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID';
var CJ_PDB_PJ_FOLDER_ID_PROP_ = 'CJ_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID';
var CJ_PDB_TARGET_SS_ID_PROP_ = 'CJ_PROJECT_DAILY_BACKUP_TARGET_SS_ID';
var CJ_PDB_RETENTION_DAYS_PROP_ = 'CJ_PROJECT_DAILY_BACKUP_RETENTION_DAYS';
var CJ_PDB_SHRINK_RATIO_PROP_ = 'CJ_PROJECT_DAILY_BACKUP_SHRINK_RATIO';
var CJ_PDB_MAIL_TO_PROP_ = 'CJ_PROJECT_DAILY_BACKUP_MAIL_TO';
var CJ_PDB_REQUIRED_SHEETS_PROP_ = 'CJ_PROJECT_DAILY_BACKUP_REQUIRED_SHEETS';

var CJ_PDB_ROOT_FOLDER_NAME_ = '99.バックアップ';
var CJ_PDB_PJ_FOLDER_NAME_ = '競合Jitta';
var CJ_PDB_COPY_PREFIX_ = '競合Jitta_backup_';
var CJ_PDB_RETENTION_DEFAULT_ = 14;
var CJ_PDB_SHRINK_DEFAULT_ = 0.8;
var CJ_PDB_TRIGGER_FN_ = 'cjRunProjectDailyBackupFromTrigger';
var CJ_PDB_REQUIRED_SHEETS_DEFAULT_ = 'メーカーサマリ|Keepaフル';
var CJ_PDB_TZ_ = 'Asia/Tokyo';
var CJ_PDB_TARGET_SS_DEFAULT_ = '1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs';

function cjPdbEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(CJ_PDB_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toLowerCase();
  return !(s === 'false' || s === '0' || s === 'off' || s === 'no');
}

function cjPdbRetentionDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(CJ_PDB_RETENTION_DAYS_PROP_) || '');
  if (!(n > 0)) n = CJ_PDB_RETENTION_DEFAULT_;
  return Math.floor(n);
}

function cjPdbShrinkRatio_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(CJ_PDB_SHRINK_RATIO_PROP_) || '');
  if (!(n > 0) || n >= 1) n = CJ_PDB_SHRINK_DEFAULT_;
  return n;
}

function cjPdbWindowBounds_() {
  var now = new Date();
  var todayYmd = Utilities.formatDate(now, CJ_PDB_TZ_, 'yyyy-MM-dd');
  var end = Utilities.parseDate(todayYmd + ' 08:00:00', CJ_PDB_TZ_, 'yyyy-MM-dd HH:mm:ss');
  var start = new Date(end.getTime() - 24 * 60 * 60 * 1000);
  return { start: start, end: end, ymd: todayYmd };
}

function cjPdbFileUpdatedInWindow_(file, win) {
  if (!file || !win) return false;
  var t = file.getLastUpdated();
  if (!t) return false;
  var ms = t.getTime();
  return ms >= win.start.getTime() && ms < win.end.getTime() + 60 * 1000;
}

function cjPdbResolveFolderByName_(parent, name) {
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return null;
}

function cjPdbEnsurePjFolder_() {
  var props = PropertiesService.getScriptProperties();
  var pjId = String(props.getProperty(CJ_PDB_PJ_FOLDER_ID_PROP_) || '').trim();
  if (pjId) {
    try {
      return DriveApp.getFolderById(pjId);
    } catch (e1) {
      throw new Error('PJフォルダID不正: ' + CJ_PDB_PJ_FOLDER_ID_PROP_);
    }
  }
  var rootId = String(props.getProperty(CJ_PDB_ROOT_FOLDER_ID_PROP_) || '').trim();
  var root = null;
  if (rootId) {
    try {
      root = DriveApp.getFolderById(rootId);
    } catch (e2) {
      throw new Error('親フォルダID不正: ' + CJ_PDB_ROOT_FOLDER_ID_PROP_);
    }
  } else {
    root = cjPdbResolveFolderByName_(null, CJ_PDB_ROOT_FOLDER_NAME_);
    if (!root) throw new Error('Driveに「' + CJ_PDB_ROOT_FOLDER_NAME_ + '」が無い。作成するか ROOT_FOLDER_ID を設定');
  }
  var pj = cjPdbResolveFolderByName_(root, CJ_PDB_PJ_FOLDER_NAME_);
  if (!pj) {
    pj = root.createFolder(CJ_PDB_PJ_FOLDER_NAME_);
    Logger.log('[競合Jitta日次バックアップ] created pj folder=%s', pj.getId());
  }
  return pj;
}

function cjPdbTargetFile_() {
  var id = String(PropertiesService.getScriptProperties().getProperty(CJ_PDB_TARGET_SS_ID_PROP_) || '').trim();
  if (!id) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) id = ss.getId();
    } catch (eAct) {
      /* trigger */
    }
  }
  if (!id) id = CJ_PDB_TARGET_SS_DEFAULT_;
  return DriveApp.getFileById(id);
}

function cjPdbRequiredSheetNames_() {
  var raw = String(PropertiesService.getScriptProperties().getProperty(CJ_PDB_REQUIRED_SHEETS_PROP_) || '').trim();
  if (!raw) raw = CJ_PDB_REQUIRED_SHEETS_DEFAULT_;
  return raw
    .split('|')
    .map(function (s) {
      return String(s || '').trim();
    })
    .filter(function (s) {
      return !!s;
    });
}

function cjPdbCorruptSuspect_(file, folder) {
  try {
    var ss = SpreadsheetApp.openById(file.getId());
    var need = cjPdbRequiredSheetNames_();
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
  var prev = cjPdbLatestGoodCopy_(folder, file.getId());
  if (prev) {
    var ratio = cjPdbShrinkRatio_();
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

function cjPdbLatestGoodCopy_(folder, liveId) {
  if (!folder) return null;
  var it = folder.getFiles();
  var best = null;
  var bestName = '';
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(CJ_PDB_COPY_PREFIX_) !== 0) continue;
    if (f.isTrashed()) continue;
    if (!best || name > bestName) {
      best = f;
      bestName = name;
    }
  }
  return best;
}

function cjPdbPurgeOldCopies_(folder, liveId, retentionDays) {
  if (!folder || !(retentionDays > 0)) return 0;
  var cutoff = new Date().getTime() - retentionDays * 24 * 60 * 60 * 1000;
  var it = folder.getFiles();
  var n = 0;
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(CJ_PDB_COPY_PREFIX_) !== 0) continue;
    var created = f.getDateCreated();
    if (created && created.getTime() < cutoff) {
      f.setTrashed(true);
      n++;
    }
  }
  return n;
}

function cjPdbAppendLog_(state, detail) {
  try {
    var file = cjPdbTargetFile_();
    var ss = SpreadsheetApp.openById(file.getId());
    var sh = ss.getSheetByName('実態AIログ');
    if (!sh) return;
    sh.appendRow([
      new Date(),
      '日次バックアップ',
      state,
      String(detail || '').slice(0, 900)
    ]);
  } catch (eLog) {
    Logger.log('[競合Jitta日次バックアップ] log_fail %s', String((eLog && eLog.message) || eLog).slice(0, 120));
  }
}

function cjPdbReportHuman_(subject, body) {
  cjPdbAppendLog_('REPORT', subject + ' | ' + body);
  Logger.log('[競合Jitta日次バックアップ] REPORT %s %s', subject, body);
  var to = String(PropertiesService.getScriptProperties().getProperty(CJ_PDB_MAIL_TO_PROP_) || '').trim();
  if (!to) return;
  try {
    MailApp.sendEmail(to, '[競合Jittaバックアップ] ' + subject, body);
  } catch (eMail) {
    Logger.log(
      '[競合Jitta日次バックアップ] mail_fail %s',
      String((eMail && eMail.message) || eMail).slice(0, 160)
    );
    cjPdbAppendLog_('MAIL_FAIL', String((eMail && eMail.message) || eMail).slice(0, 200));
  }
}

/**
 * @return {{ok:boolean, action:string, detail:string}}
 */
function cjRunProjectDailyBackupCore_() {
  if (!cjPdbEnabled_()) {
    return { ok: true, action: 'disabled', detail: CJ_PDB_ENABLED_PROP_ + '=false' };
  }
  var win = cjPdbWindowBounds_();
  var file = cjPdbTargetFile_();
  var liveId = file.getId();
  if (!cjPdbFileUpdatedInWindow_(file, win)) {
    var skipDetail =
      'no_update_in_window lastUpdated=' +
      Utilities.formatDate(file.getLastUpdated(), CJ_PDB_TZ_, "yyyy-MM-dd'T'HH:mm:ss") +
      ' window=' +
      Utilities.formatDate(win.start, CJ_PDB_TZ_, "yyyy-MM-dd'T'HH:mm") +
      '..' +
      Utilities.formatDate(win.end, CJ_PDB_TZ_, "yyyy-MM-dd'T'HH:mm");
    cjPdbAppendLog_('SKIP', skipDetail);
    return { ok: true, action: 'skip_no_update', detail: skipDetail };
  }
  var folder = cjPdbEnsurePjFolder_();
  var corrupt = cjPdbCorruptSuspect_(file, folder);
  if (!corrupt.ok) {
    var msg = 'corrupt_skip ' + corrupt.reason + ' liveId=' + liveId;
    cjPdbReportHuman_('破損疑いのためコピーしませんでした', msg);
    return { ok: true, action: 'skip_corrupt', detail: msg };
  }
  var copyName = CJ_PDB_COPY_PREFIX_ + win.ymd;
  var existing = folder.getFilesByName(copyName);
  if (existing.hasNext()) {
    var existDetail = 'already_exists name=' + copyName;
    cjPdbAppendLog_('SKIP', existDetail);
    return { ok: true, action: 'skip_exists', detail: existDetail };
  }
  var copy = file.makeCopy(copyName, folder);
  var purged = cjPdbPurgeOldCopies_(folder, liveId, cjPdbRetentionDays_());
  var doneDetail =
    'copied id=' +
    copy.getId() +
    ' name=' +
    copyName +
    ' purged=' +
    purged +
    ' folder=' +
    folder.getId();
  cjPdbAppendLog_('DONE', doneDetail);
  Logger.log('[競合Jitta日次バックアップ] %s', doneDetail);
  return { ok: true, action: 'copied', detail: doneDetail };
}

function cjRunProjectDailyBackupFromTrigger() {
  try {
    var r = cjRunProjectDailyBackupCore_();
    Logger.log('[競合Jitta日次バックアップ] trigger action=%s %s', r.action, r.detail);
  } catch (e) {
    var msg = String((e && e.message) || e).slice(0, 400);
    cjPdbReportHuman_('日次バックアップ失敗', msg);
    throw e;
  }
}

function menuCjProjectDailyBackupRunNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = cjRunProjectDailyBackupCore_();
    ui.alert('競合Jitta 日次バックアップ', r.action + '\n' + r.detail, ui.ButtonSet.OK);
  } catch (e) {
    var msg = String((e && e.message) || e);
    cjPdbReportHuman_('手動バックアップ失敗', msg);
    ui.alert('失敗', msg, ui.ButtonSet.OK);
  }
}

function menuCjProjectDailyBackupForceCopy() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    '強制バックアップ',
    '更新窓を無視して今すぐコピーします（破損疑いならコピーしません）。実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;
  try {
    if (!cjPdbEnabled_()) {
      ui.alert('OFF', CJ_PDB_ENABLED_PROP_ + '=false', ui.ButtonSet.OK);
      return;
    }
    var file = cjPdbTargetFile_();
    var folder = cjPdbEnsurePjFolder_();
    var corrupt = cjPdbCorruptSuspect_(file, folder);
    if (!corrupt.ok) {
      cjPdbReportHuman_('強制コピー中止（破損疑い）', corrupt.reason);
      ui.alert('中止', corrupt.reason, ui.ButtonSet.OK);
      return;
    }
    var ymd = Utilities.formatDate(new Date(), CJ_PDB_TZ_, 'yyyyMMdd_HHmmss');
    var copyName = CJ_PDB_COPY_PREFIX_ + 'force_' + ymd;
    var copy = file.makeCopy(copyName, folder);
    var detail = 'force id=' + copy.getId() + ' name=' + copyName;
    cjPdbAppendLog_('DONE', detail);
    ui.alert('完了', detail, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function menuCjProjectDailyBackupInstallTrigger() {
  var ui = SpreadsheetApp.getUi();
  cjPdbRemoveTriggers_();
  ScriptApp.newTrigger(CJ_PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(CJ_PDB_TZ_).create();
  ui.alert('競合Jitta 日次バックアップ', '毎朝8時トリガーを設置しました（' + CJ_PDB_TRIGGER_FN_ + '）', ui.ButtonSet.OK);
}

function menuCjProjectDailyBackupRemoveTrigger() {
  var ui = SpreadsheetApp.getUi();
  var n = cjPdbRemoveTriggers_();
  ui.alert('競合Jitta 日次バックアップ', 'トリガー削除 n=' + n, ui.ButtonSet.OK);
}

/**
 * @param {string=} pjFolderId
 */
function cjSetupProjectDailyBackup_(pjFolderId) {
  var props = PropertiesService.getScriptProperties();
  var rootId = '1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE';
  var target = CJ_PDB_TARGET_SS_DEFAULT_;
  var pj = String(pjFolderId || '').trim();
  if (!pj) {
    var folder = cjPdbEnsurePjFolder_();
    pj = folder.getId();
  }
  props.setProperties(
    {
      CJ_PROJECT_DAILY_BACKUP_ENABLED: 'true',
      CJ_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID: rootId,
      CJ_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID: pj,
      CJ_PROJECT_DAILY_BACKUP_TARGET_SS_ID: target
    },
    false
  );
  cjPdbRemoveTriggers_();
  ScriptApp.newTrigger(CJ_PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(CJ_PDB_TZ_).create();
  var out = {
    ok: true,
    rootId: rootId,
    pjFolderId: pj,
    targetSsId: target,
    trigger: CJ_PDB_TRIGGER_FN_
  };
  Logger.log('[競合Jitta日次バックアップ] setup %s', JSON.stringify(out));
  return out;
}

function menuCjProjectDailyBackupSetup() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = cjSetupProjectDailyBackup_();
    ui.alert(
      '競合Jitta 日次バックアップ セットアップ',
      'Property と毎朝8時トリガーを設定しました。\npjFolder=' + r.pjFolderId + '\nss=' + r.targetSsId,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function cjPdbRemoveTriggers_() {
  var list = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === CJ_PDB_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(list[i]);
      n++;
    }
  }
  return n;
}
