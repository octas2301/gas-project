/**
 * プロジェクト別 Drive 自動バックアップ（税務＝消費税輸出免税一覧）
 * 正本: docs/org/PROJECT_DAILY_BACKUP_REQUIREMENTS.md
 *
 * 切戻し: TAX_PROJECT_DAILY_BACKUP_ENABLED=false
 * トリガー: taxRunProjectDailyBackupFromTrigger（毎日 8:00 Asia/Tokyo）
 */

var TAX_PDB_ENABLED_PROP_ = 'TAX_PROJECT_DAILY_BACKUP_ENABLED';
var TAX_PDB_ROOT_FOLDER_ID_PROP_ = 'TAX_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID';
var TAX_PDB_PJ_FOLDER_ID_PROP_ = 'TAX_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID';
var TAX_PDB_TARGET_SS_ID_PROP_ = 'TAX_PROJECT_DAILY_BACKUP_TARGET_SS_ID';
var TAX_PDB_RETENTION_DAYS_PROP_ = 'TAX_PROJECT_DAILY_BACKUP_RETENTION_DAYS';
var TAX_PDB_SHRINK_RATIO_PROP_ = 'TAX_PROJECT_DAILY_BACKUP_SHRINK_RATIO';
var TAX_PDB_MAIL_TO_PROP_ = 'TAX_PROJECT_DAILY_BACKUP_MAIL_TO';
var TAX_PDB_REQUIRED_SHEETS_PROP_ = 'TAX_PROJECT_DAILY_BACKUP_REQUIRED_SHEETS';

var TAX_PDB_ROOT_FOLDER_NAME_ = '99.バックアップ';
var TAX_PDB_PJ_FOLDER_NAME_ = '税務';
var TAX_PDB_COPY_PREFIX_ = '税務_backup_';
var TAX_PDB_RETENTION_DEFAULT_ = 14;
var TAX_PDB_SHRINK_DEFAULT_ = 0.8;
var TAX_PDB_TRIGGER_FN_ = 'taxRunProjectDailyBackupFromTrigger';
var TAX_PDB_REQUIRED_SHEETS_DEFAULT_ = '設定|ログ';
var TAX_PDB_TZ_ = 'Asia/Tokyo';
var TAX_PDB_TARGET_SS_DEFAULT_ = '1gCA0xK71arBrAy_vMJPETD20H2rvvoN0arGMG9t3GcM';

function taxPdbEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(TAX_PDB_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toLowerCase();
  return !(s === 'false' || s === '0' || s === 'off' || s === 'no');
}

function taxPdbRetentionDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(TAX_PDB_RETENTION_DAYS_PROP_) || '');
  if (!(n > 0)) n = TAX_PDB_RETENTION_DEFAULT_;
  return Math.floor(n);
}

function taxPdbShrinkRatio_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(TAX_PDB_SHRINK_RATIO_PROP_) || '');
  if (!(n > 0) || n >= 1) n = TAX_PDB_SHRINK_DEFAULT_;
  return n;
}

function taxPdbWindowBounds_() {
  var now = new Date();
  var todayYmd = Utilities.formatDate(now, TAX_PDB_TZ_, 'yyyy-MM-dd');
  var end = Utilities.parseDate(todayYmd + ' 08:00:00', TAX_PDB_TZ_, 'yyyy-MM-dd HH:mm:ss');
  var start = new Date(end.getTime() - 24 * 60 * 60 * 1000);
  return { start: start, end: end, ymd: todayYmd };
}

function taxPdbFileUpdatedInWindow_(file, win) {
  if (!file || !win) return false;
  var t = file.getLastUpdated();
  if (!t) return false;
  var ms = t.getTime();
  return ms >= win.start.getTime() && ms < win.end.getTime() + 60 * 1000;
}

function taxPdbResolveFolderByName_(parent, name) {
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return null;
}

function taxPdbEnsurePjFolder_() {
  var props = PropertiesService.getScriptProperties();
  var pjId = String(props.getProperty(TAX_PDB_PJ_FOLDER_ID_PROP_) || '').trim();
  if (pjId) {
    try {
      return DriveApp.getFolderById(pjId);
    } catch (e1) {
      throw new Error('PJフォルダID不正: ' + TAX_PDB_PJ_FOLDER_ID_PROP_);
    }
  }
  var rootId = String(props.getProperty(TAX_PDB_ROOT_FOLDER_ID_PROP_) || '').trim();
  var root = null;
  if (rootId) {
    try {
      root = DriveApp.getFolderById(rootId);
    } catch (e2) {
      throw new Error('親フォルダID不正: ' + TAX_PDB_ROOT_FOLDER_ID_PROP_);
    }
  } else {
    root = taxPdbResolveFolderByName_(null, TAX_PDB_ROOT_FOLDER_NAME_);
    if (!root) throw new Error('Driveに「' + TAX_PDB_ROOT_FOLDER_NAME_ + '」が無い。作成するか ROOT_FOLDER_ID を設定');
  }
  var pj = taxPdbResolveFolderByName_(root, TAX_PDB_PJ_FOLDER_NAME_);
  if (!pj) {
    pj = root.createFolder(TAX_PDB_PJ_FOLDER_NAME_);
    Logger.log('[税務日次バックアップ] created pj folder=%s', pj.getId());
  }
  return pj;
}

function taxPdbTargetFile_() {
  var id = String(PropertiesService.getScriptProperties().getProperty(TAX_PDB_TARGET_SS_ID_PROP_) || '').trim();
  if (!id) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) id = ss.getId();
    } catch (eAct) {
      /* trigger */
    }
  }
  if (!id) id = TAX_PDB_TARGET_SS_DEFAULT_;
  return DriveApp.getFileById(id);
}

function taxPdbRequiredSheetNames_() {
  var raw = String(PropertiesService.getScriptProperties().getProperty(TAX_PDB_REQUIRED_SHEETS_PROP_) || '').trim();
  if (!raw) raw = TAX_PDB_REQUIRED_SHEETS_DEFAULT_;
  return raw
    .split('|')
    .map(function (s) {
      return String(s || '').trim();
    })
    .filter(function (s) {
      return !!s;
    });
}

function taxPdbCorruptSuspect_(file, folder) {
  try {
    var ss = SpreadsheetApp.openById(file.getId());
    var need = taxPdbRequiredSheetNames_();
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
  var prev = taxPdbLatestGoodCopy_(folder, file.getId());
  if (prev) {
    var ratio = taxPdbShrinkRatio_();
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

function taxPdbLatestGoodCopy_(folder, liveId) {
  if (!folder) return null;
  var it = folder.getFiles();
  var best = null;
  var bestName = '';
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(TAX_PDB_COPY_PREFIX_) !== 0) continue;
    if (f.isTrashed()) continue;
    if (!best || name > bestName) {
      best = f;
      bestName = name;
    }
  }
  return best;
}

function taxPdbPurgeOldCopies_(folder, liveId, retentionDays) {
  if (!folder || !(retentionDays > 0)) return 0;
  var cutoff = new Date().getTime() - retentionDays * 24 * 60 * 60 * 1000;
  var it = folder.getFiles();
  var n = 0;
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(TAX_PDB_COPY_PREFIX_) !== 0) continue;
    var created = f.getDateCreated();
    if (created && created.getTime() < cutoff) {
      f.setTrashed(true);
      n++;
    }
  }
  return n;
}

function taxPdbAppendLog_(state, detail) {
  try {
    var file = taxPdbTargetFile_();
    var ss = SpreadsheetApp.openById(file.getId());
    var sh = ss.getSheetByName('ログ');
    if (!sh) return;
    sh.appendRow([
      new Date(),
      '日次バックアップ',
      state,
      String(detail || '').slice(0, 900)
    ]);
  } catch (eLog) {
    Logger.log('[税務日次バックアップ] log_fail %s', String((eLog && eLog.message) || eLog).slice(0, 120));
  }
}

function taxPdbReportHuman_(subject, body) {
  taxPdbAppendLog_('REPORT', subject + ' | ' + body);
  Logger.log('[税務日次バックアップ] REPORT %s %s', subject, body);
  var to = String(PropertiesService.getScriptProperties().getProperty(TAX_PDB_MAIL_TO_PROP_) || '').trim();
  if (!to) return;
  try {
    MailApp.sendEmail(to, '[税務バックアップ] ' + subject, body);
  } catch (eMail) {
    Logger.log(
      '[税務日次バックアップ] mail_fail %s',
      String((eMail && eMail.message) || eMail).slice(0, 160)
    );
    taxPdbAppendLog_('MAIL_FAIL', String((eMail && eMail.message) || eMail).slice(0, 200));
  }
}

/**
 * @return {{ok:boolean, action:string, detail:string}}
 */
function taxRunProjectDailyBackupCore_() {
  if (!taxPdbEnabled_()) {
    return { ok: true, action: 'disabled', detail: TAX_PDB_ENABLED_PROP_ + '=false' };
  }
  var win = taxPdbWindowBounds_();
  var file = taxPdbTargetFile_();
  var liveId = file.getId();
  if (!taxPdbFileUpdatedInWindow_(file, win)) {
    var skipDetail =
      'no_update_in_window lastUpdated=' +
      Utilities.formatDate(file.getLastUpdated(), TAX_PDB_TZ_, "yyyy-MM-dd'T'HH:mm:ss") +
      ' window=' +
      Utilities.formatDate(win.start, TAX_PDB_TZ_, "yyyy-MM-dd'T'HH:mm") +
      '..' +
      Utilities.formatDate(win.end, TAX_PDB_TZ_, "yyyy-MM-dd'T'HH:mm");
    taxPdbAppendLog_('SKIP', skipDetail);
    return { ok: true, action: 'skip_no_update', detail: skipDetail };
  }
  var folder = taxPdbEnsurePjFolder_();
  var corrupt = taxPdbCorruptSuspect_(file, folder);
  if (!corrupt.ok) {
    var msg = 'corrupt_skip ' + corrupt.reason + ' liveId=' + liveId;
    taxPdbReportHuman_('破損疑いのためコピーしませんでした', msg);
    return { ok: true, action: 'skip_corrupt', detail: msg };
  }
  var copyName = TAX_PDB_COPY_PREFIX_ + win.ymd;
  var existing = folder.getFilesByName(copyName);
  if (existing.hasNext()) {
    var existDetail = 'already_exists name=' + copyName;
    taxPdbAppendLog_('SKIP', existDetail);
    return { ok: true, action: 'skip_exists', detail: existDetail };
  }
  var copy = file.makeCopy(copyName, folder);
  var purged = taxPdbPurgeOldCopies_(folder, liveId, taxPdbRetentionDays_());
  var doneDetail =
    'copied id=' +
    copy.getId() +
    ' name=' +
    copyName +
    ' purged=' +
    purged +
    ' folder=' +
    folder.getId();
  taxPdbAppendLog_('DONE', doneDetail);
  Logger.log('[税務日次バックアップ] %s', doneDetail);
  return { ok: true, action: 'copied', detail: doneDetail };
}

function taxRunProjectDailyBackupFromTrigger() {
  try {
    var r = taxRunProjectDailyBackupCore_();
    Logger.log('[税務日次バックアップ] trigger action=%s %s', r.action, r.detail);
  } catch (e) {
    var msg = String((e && e.message) || e).slice(0, 400);
    taxPdbReportHuman_('日次バックアップ失敗', msg);
    throw e;
  }
}

function menuTaxProjectDailyBackupRunNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = taxRunProjectDailyBackupCore_();
    ui.alert('税務 日次バックアップ', r.action + '\n' + r.detail, ui.ButtonSet.OK);
  } catch (e) {
    var msg = String((e && e.message) || e);
    taxPdbReportHuman_('手動バックアップ失敗', msg);
    ui.alert('失敗', msg, ui.ButtonSet.OK);
  }
}

function menuTaxProjectDailyBackupForceCopy() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    '強制バックアップ',
    '更新窓を無視して今すぐコピーします（破損疑いならコピーしません）。実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;
  try {
    if (!taxPdbEnabled_()) {
      ui.alert('OFF', TAX_PDB_ENABLED_PROP_ + '=false', ui.ButtonSet.OK);
      return;
    }
    var file = taxPdbTargetFile_();
    var folder = taxPdbEnsurePjFolder_();
    var corrupt = taxPdbCorruptSuspect_(file, folder);
    if (!corrupt.ok) {
      taxPdbReportHuman_('強制コピー中止（破損疑い）', corrupt.reason);
      ui.alert('中止', corrupt.reason, ui.ButtonSet.OK);
      return;
    }
    var ymd = Utilities.formatDate(new Date(), TAX_PDB_TZ_, 'yyyyMMdd_HHmmss');
    var copyName = TAX_PDB_COPY_PREFIX_ + 'force_' + ymd;
    var copy = file.makeCopy(copyName, folder);
    var detail = 'force id=' + copy.getId() + ' name=' + copyName;
    taxPdbAppendLog_('DONE', detail);
    ui.alert('完了', detail, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function menuTaxProjectDailyBackupInstallTrigger() {
  var ui = SpreadsheetApp.getUi();
  taxPdbRemoveTriggers_();
  ScriptApp.newTrigger(TAX_PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(TAX_PDB_TZ_).create();
  ui.alert('税務 日次バックアップ', '毎朝8時トリガーを設置しました（' + TAX_PDB_TRIGGER_FN_ + '）', ui.ButtonSet.OK);
}

function menuTaxProjectDailyBackupRemoveTrigger() {
  var ui = SpreadsheetApp.getUi();
  var n = taxPdbRemoveTriggers_();
  ui.alert('税務 日次バックアップ', 'トリガー削除 n=' + n, ui.ButtonSet.OK);
}

/**
 * @param {string=} pjFolderId
 */
function taxSetupProjectDailyBackup_(pjFolderId) {
  var props = PropertiesService.getScriptProperties();
  var rootId = '1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE';
  var target = TAX_PDB_TARGET_SS_DEFAULT_;
  var pj = String(pjFolderId || '').trim();
  if (!pj) {
    var folder = taxPdbEnsurePjFolder_();
    pj = folder.getId();
  }
  props.setProperties(
    {
      TAX_PROJECT_DAILY_BACKUP_ENABLED: 'true',
      TAX_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID: rootId,
      TAX_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID: pj,
      TAX_PROJECT_DAILY_BACKUP_TARGET_SS_ID: target
    },
    false
  );
  taxPdbRemoveTriggers_();
  ScriptApp.newTrigger(TAX_PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(TAX_PDB_TZ_).create();
  var out = {
    ok: true,
    rootId: rootId,
    pjFolderId: pj,
    targetSsId: target,
    trigger: TAX_PDB_TRIGGER_FN_
  };
  Logger.log('[税務日次バックアップ] setup %s', JSON.stringify(out));
  return out;
}

function menuTaxProjectDailyBackupSetup() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = taxSetupProjectDailyBackup_();
    ui.alert(
      '税務 日次バックアップ セットアップ',
      'Property と毎朝8時トリガーを設定しました。\npjFolder=' + r.pjFolderId + '\nss=' + r.targetSsId,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function taxPdbRemoveTriggers_() {
  var list = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === TAX_PDB_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(list[i]);
      n++;
    }
  }
  return n;
}
