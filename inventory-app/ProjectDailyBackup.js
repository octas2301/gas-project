/**
 * プロジェクト別 Drive 自動バックアップ（在庫管理）
 * 正本: docs/org/PROJECT_DAILY_BACKUP_REQUIREMENTS.md
 *
 * 切戻し: INV_PROJECT_DAILY_BACKUP_ENABLED=false
 * トリガー: invRunProjectDailyBackupFromTrigger（毎日 8:00 Asia/Tokyo）
 */

var INV_PDB_ENABLED_PROP_ = 'INV_PROJECT_DAILY_BACKUP_ENABLED';
var INV_PDB_ROOT_FOLDER_ID_PROP_ = 'INV_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID';
var INV_PDB_PJ_FOLDER_ID_PROP_ = 'INV_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID';
var INV_PDB_TARGET_SS_ID_PROP_ = 'INV_PROJECT_DAILY_BACKUP_TARGET_SS_ID';
var INV_PDB_RETENTION_DAYS_PROP_ = 'INV_PROJECT_DAILY_BACKUP_RETENTION_DAYS';
var INV_PDB_SHRINK_RATIO_PROP_ = 'INV_PROJECT_DAILY_BACKUP_SHRINK_RATIO';
var INV_PDB_MAIL_TO_PROP_ = 'INV_PROJECT_DAILY_BACKUP_MAIL_TO';
var INV_PDB_REQUIRED_SHEETS_PROP_ = 'INV_PROJECT_DAILY_BACKUP_REQUIRED_SHEETS';

var INV_PDB_ROOT_FOLDER_NAME_ = '99.バックアップ';
var INV_PDB_PJ_FOLDER_NAME_ = '在庫';
var INV_PDB_COPY_PREFIX_ = '在庫_backup_';
var INV_PDB_RETENTION_DEFAULT_ = 14;
var INV_PDB_SHRINK_DEFAULT_ = 0.8;
var INV_PDB_TRIGGER_FN_ = 'invRunProjectDailyBackupFromTrigger';
var INV_PDB_REQUIRED_SHEETS_DEFAULT_ = '商品マスタ|在庫移動リスト';
var INV_PDB_TZ_ = 'Asia/Tokyo';
var INV_PDB_TARGET_SS_DEFAULT_ = '1_kqBcQTcL0Q5--KsAdxMYbPMsaToLEXxOMuDw_z-kwE';

function invPdbEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(INV_PDB_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toLowerCase();
  return !(s === 'false' || s === '0' || s === 'off' || s === 'no');
}

function invPdbRetentionDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(INV_PDB_RETENTION_DAYS_PROP_) || '');
  if (!(n > 0)) n = INV_PDB_RETENTION_DEFAULT_;
  return Math.floor(n);
}

function invPdbShrinkRatio_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(INV_PDB_SHRINK_RATIO_PROP_) || '');
  if (!(n > 0) || n >= 1) n = INV_PDB_SHRINK_DEFAULT_;
  return n;
}

function invPdbWindowBounds_() {
  var now = new Date();
  var todayYmd = Utilities.formatDate(now, INV_PDB_TZ_, 'yyyy-MM-dd');
  var end = Utilities.parseDate(todayYmd + ' 08:00:00', INV_PDB_TZ_, 'yyyy-MM-dd HH:mm:ss');
  var start = new Date(end.getTime() - 24 * 60 * 60 * 1000);
  return { start: start, end: end, ymd: todayYmd };
}

function invPdbFileUpdatedInWindow_(file, win) {
  if (!file || !win) return false;
  var t = file.getLastUpdated();
  if (!t) return false;
  var ms = t.getTime();
  return ms >= win.start.getTime() && ms < win.end.getTime() + 60 * 1000;
}

function invPdbResolveFolderByName_(parent, name) {
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return null;
}

function invPdbEnsurePjFolder_() {
  var props = PropertiesService.getScriptProperties();
  var pjId = String(props.getProperty(INV_PDB_PJ_FOLDER_ID_PROP_) || '').trim();
  if (pjId) {
    try {
      return DriveApp.getFolderById(pjId);
    } catch (e1) {
      throw new Error('PJフォルダID不正: ' + INV_PDB_PJ_FOLDER_ID_PROP_);
    }
  }
  var rootId = String(props.getProperty(INV_PDB_ROOT_FOLDER_ID_PROP_) || '').trim();
  var root = null;
  if (rootId) {
    try {
      root = DriveApp.getFolderById(rootId);
    } catch (e2) {
      throw new Error('親フォルダID不正: ' + INV_PDB_ROOT_FOLDER_ID_PROP_);
    }
  } else {
    root = invPdbResolveFolderByName_(null, INV_PDB_ROOT_FOLDER_NAME_);
    if (!root) throw new Error('Driveに「' + INV_PDB_ROOT_FOLDER_NAME_ + '」が無い。作成するか ROOT_FOLDER_ID を設定');
  }
  var pj = invPdbResolveFolderByName_(root, INV_PDB_PJ_FOLDER_NAME_);
  if (!pj) {
    pj = root.createFolder(INV_PDB_PJ_FOLDER_NAME_);
    Logger.log('[在庫日次バックアップ] created pj folder=%s', pj.getId());
  }
  return pj;
}

function invPdbTargetSsId_() {
  var id = String(PropertiesService.getScriptProperties().getProperty(INV_PDB_TARGET_SS_ID_PROP_) || '').trim();
  if (id) return id;
  var existing = String(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '').trim();
  if (existing) return existing;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss.getId();
  } catch (eAct) {
    /* web / trigger */
  }
  return INV_PDB_TARGET_SS_DEFAULT_;
}

function invPdbTargetFile_() {
  return DriveApp.getFileById(invPdbTargetSsId_());
}

function invPdbRequiredSheetNames_() {
  var raw = String(PropertiesService.getScriptProperties().getProperty(INV_PDB_REQUIRED_SHEETS_PROP_) || '').trim();
  if (!raw) raw = INV_PDB_REQUIRED_SHEETS_DEFAULT_;
  return raw
    .split('|')
    .map(function (s) {
      return String(s || '').trim();
    })
    .filter(function (s) {
      return !!s;
    });
}

function invPdbCorruptSuspect_(file, folder) {
  try {
    var ss = SpreadsheetApp.openById(file.getId());
    var need = invPdbRequiredSheetNames_();
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
  var prev = invPdbLatestGoodCopy_(folder, file.getId());
  if (prev) {
    var ratio = invPdbShrinkRatio_();
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

function invPdbLatestGoodCopy_(folder, liveId) {
  if (!folder) return null;
  var it = folder.getFiles();
  var best = null;
  var bestName = '';
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(INV_PDB_COPY_PREFIX_) !== 0) continue;
    if (f.isTrashed()) continue;
    if (!best || name > bestName) {
      best = f;
      bestName = name;
    }
  }
  return best;
}

function invPdbPurgeOldCopies_(folder, liveId, retentionDays) {
  if (!folder || !(retentionDays > 0)) return 0;
  var cutoff = new Date().getTime() - retentionDays * 24 * 60 * 60 * 1000;
  var it = folder.getFiles();
  var n = 0;
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(INV_PDB_COPY_PREFIX_) !== 0) continue;
    var created = f.getDateCreated();
    if (created && created.getTime() < cutoff) {
      f.setTrashed(true);
      n++;
    }
  }
  return n;
}

function invPdbAppendLog_(state, detail) {
  try {
    var ss = SpreadsheetApp.openById(invPdbTargetSsId_());
    var sh = ss.getSheetByName('操作ログ');
    if (!sh) return;
    sh.appendRow([
      new Date(),
      '日次バックアップ',
      state,
      String(detail || '').slice(0, 900)
    ]);
  } catch (eLog) {
    Logger.log('[在庫日次バックアップ] log_fail %s', String((eLog && eLog.message) || eLog).slice(0, 120));
  }
}

function invPdbReportHuman_(subject, body) {
  invPdbAppendLog_('REPORT', subject + ' | ' + body);
  Logger.log('[在庫日次バックアップ] REPORT %s %s', subject, body);
  var to = String(PropertiesService.getScriptProperties().getProperty(INV_PDB_MAIL_TO_PROP_) || '').trim();
  if (!to) return;
  try {
    MailApp.sendEmail(to, '[在庫バックアップ] ' + subject, body);
  } catch (eMail) {
    Logger.log(
      '[在庫日次バックアップ] mail_fail %s',
      String((eMail && eMail.message) || eMail).slice(0, 160)
    );
    invPdbAppendLog_('MAIL_FAIL', String((eMail && eMail.message) || eMail).slice(0, 200));
  }
}

/**
 * @return {{ok:boolean, action:string, detail:string}}
 */
function invRunProjectDailyBackupCore_() {
  if (!invPdbEnabled_()) {
    return { ok: true, action: 'disabled', detail: INV_PDB_ENABLED_PROP_ + '=false' };
  }
  var win = invPdbWindowBounds_();
  var file = invPdbTargetFile_();
  var liveId = file.getId();
  if (!invPdbFileUpdatedInWindow_(file, win)) {
    var skipDetail =
      'no_update_in_window lastUpdated=' +
      Utilities.formatDate(file.getLastUpdated(), INV_PDB_TZ_, "yyyy-MM-dd'T'HH:mm:ss") +
      ' window=' +
      Utilities.formatDate(win.start, INV_PDB_TZ_, "yyyy-MM-dd'T'HH:mm") +
      '..' +
      Utilities.formatDate(win.end, INV_PDB_TZ_, "yyyy-MM-dd'T'HH:mm");
    invPdbAppendLog_('SKIP', skipDetail);
    return { ok: true, action: 'skip_no_update', detail: skipDetail };
  }
  var folder = invPdbEnsurePjFolder_();
  var corrupt = invPdbCorruptSuspect_(file, folder);
  if (!corrupt.ok) {
    var msg = 'corrupt_skip ' + corrupt.reason + ' liveId=' + liveId;
    invPdbReportHuman_('破損疑いのためコピーしませんでした', msg);
    return { ok: true, action: 'skip_corrupt', detail: msg };
  }
  var copyName = INV_PDB_COPY_PREFIX_ + win.ymd;
  var existing = folder.getFilesByName(copyName);
  if (existing.hasNext()) {
    var existDetail = 'already_exists name=' + copyName;
    invPdbAppendLog_('SKIP', existDetail);
    return { ok: true, action: 'skip_exists', detail: existDetail };
  }
  var copy = file.makeCopy(copyName, folder);
  var purged = invPdbPurgeOldCopies_(folder, liveId, invPdbRetentionDays_());
  var doneDetail =
    'copied id=' +
    copy.getId() +
    ' name=' +
    copyName +
    ' purged=' +
    purged +
    ' folder=' +
    folder.getId();
  invPdbAppendLog_('DONE', doneDetail);
  Logger.log('[在庫日次バックアップ] %s', doneDetail);
  return { ok: true, action: 'copied', detail: doneDetail };
}

function invRunProjectDailyBackupFromTrigger() {
  try {
    var r = invRunProjectDailyBackupCore_();
    Logger.log('[在庫日次バックアップ] trigger action=%s %s', r.action, r.detail);
  } catch (e) {
    var msg = String((e && e.message) || e).slice(0, 400);
    invPdbReportHuman_('日次バックアップ失敗', msg);
    throw e;
  }
}

function menuInvProjectDailyBackupRunNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = invRunProjectDailyBackupCore_();
    ui.alert('在庫 日次バックアップ', r.action + '\n' + r.detail, ui.ButtonSet.OK);
  } catch (e) {
    var msg = String((e && e.message) || e);
    invPdbReportHuman_('手動バックアップ失敗', msg);
    ui.alert('失敗', msg, ui.ButtonSet.OK);
  }
}

function menuInvProjectDailyBackupForceCopy() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    '強制バックアップ',
    '更新窓を無視して今すぐコピーします（破損疑いならコピーしません）。実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;
  try {
    if (!invPdbEnabled_()) {
      ui.alert('OFF', INV_PDB_ENABLED_PROP_ + '=false', ui.ButtonSet.OK);
      return;
    }
    var file = invPdbTargetFile_();
    var folder = invPdbEnsurePjFolder_();
    var corrupt = invPdbCorruptSuspect_(file, folder);
    if (!corrupt.ok) {
      invPdbReportHuman_('強制コピー中止（破損疑い）', corrupt.reason);
      ui.alert('中止', corrupt.reason, ui.ButtonSet.OK);
      return;
    }
    var ymd = Utilities.formatDate(new Date(), INV_PDB_TZ_, 'yyyyMMdd_HHmmss');
    var copyName = INV_PDB_COPY_PREFIX_ + 'force_' + ymd;
    var copy = file.makeCopy(copyName, folder);
    var detail = 'force id=' + copy.getId() + ' name=' + copyName;
    invPdbAppendLog_('DONE', detail);
    ui.alert('完了', detail, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function menuInvProjectDailyBackupInstallTrigger() {
  var ui = SpreadsheetApp.getUi();
  invPdbRemoveTriggers_();
  ScriptApp.newTrigger(INV_PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(INV_PDB_TZ_).create();
  ui.alert('在庫 日次バックアップ', '毎朝8時トリガーを設置しました（' + INV_PDB_TRIGGER_FN_ + '）', ui.ButtonSet.OK);
}

function menuInvProjectDailyBackupRemoveTrigger() {
  var ui = SpreadsheetApp.getUi();
  var n = invPdbRemoveTriggers_();
  ui.alert('在庫 日次バックアップ', 'トリガー削除 n=' + n, ui.ButtonSet.OK);
}

/**
 * @param {string=} pjFolderId
 */
function invSetupProjectDailyBackup_(pjFolderId) {
  var props = PropertiesService.getScriptProperties();
  var rootId = '1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE';
  var target = invPdbTargetSsId_();
  var pj = String(pjFolderId || '').trim();
  if (!pj) {
    var folder = invPdbEnsurePjFolder_();
    pj = folder.getId();
  }
  props.setProperties(
    {
      INV_PROJECT_DAILY_BACKUP_ENABLED: 'true',
      INV_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID: rootId,
      INV_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID: pj,
      INV_PROJECT_DAILY_BACKUP_TARGET_SS_ID: target
    },
    false
  );
  invPdbRemoveTriggers_();
  ScriptApp.newTrigger(INV_PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(INV_PDB_TZ_).create();
  var out = {
    ok: true,
    rootId: rootId,
    pjFolderId: pj,
    targetSsId: target,
    trigger: INV_PDB_TRIGGER_FN_
  };
  Logger.log('[在庫日次バックアップ] setup %s', JSON.stringify(out));
  return out;
}

function menuInvProjectDailyBackupSetup() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = invSetupProjectDailyBackup_();
    ui.alert(
      '在庫 日次バックアップ セットアップ',
      'Property と毎朝8時トリガーを設定しました。\npjFolder=' + r.pjFolderId + '\nss=' + r.targetSsId,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function invPdbRemoveTriggers_() {
  var list = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === INV_PDB_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(list[i]);
      n++;
    }
  }
  return n;
}
