/**
 * プロジェクト別 Drive 自動バックアップ（出品・第1版）
 * 正本: docs/org/PROJECT_DAILY_BACKUP_REQUIREMENTS.md
 *
 * 切戻し: PROJECT_DAILY_BACKUP_ENABLED=false
 * トリガー: runProjectDailyBackupFromTrigger（毎日 8:00 Asia/Tokyo）
 */

var PDB_ENABLED_PROP_ = 'PROJECT_DAILY_BACKUP_ENABLED';
var PDB_ROOT_FOLDER_ID_PROP_ = 'PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID';
var PDB_PJ_FOLDER_ID_PROP_ = 'PROJECT_DAILY_BACKUP_PJ_FOLDER_ID';
var PDB_TARGET_SS_ID_PROP_ = 'PROJECT_DAILY_BACKUP_TARGET_SS_ID';
var PDB_RETENTION_DAYS_PROP_ = 'PROJECT_DAILY_BACKUP_RETENTION_DAYS';
var PDB_SHRINK_RATIO_PROP_ = 'PROJECT_DAILY_BACKUP_SHRINK_RATIO';
var PDB_MAIL_TO_PROP_ = 'PROJECT_DAILY_BACKUP_MAIL_TO';
var PDB_REQUIRED_SHEETS_PROP_ = 'PROJECT_DAILY_BACKUP_REQUIRED_SHEETS';

var PDB_ROOT_FOLDER_NAME_ = '99.バックアップ';
var PDB_PJ_FOLDER_NAME_ = '出品';
var PDB_COPY_PREFIX_ = '出品_backup_';
var PDB_RETENTION_DEFAULT_ = 14;
var PDB_SHRINK_DEFAULT_ = 0.8;
var PDB_TRIGGER_FN_ = 'runProjectDailyBackupFromTrigger';
var PDB_REQUIRED_SHEETS_DEFAULT_ = '▼商品マスタ(人間作業用)';
var PDB_TZ_ = 'Asia/Tokyo';

function pdbEnabled_() {
  if (typeof getBoolScriptProperty_ === 'function') {
    return getBoolScriptProperty_(PDB_ENABLED_PROP_, true);
  }
  var raw = PropertiesService.getScriptProperties().getProperty(PDB_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toLowerCase();
  return !(s === 'false' || s === '0' || s === 'off' || s === 'no');
}

function pdbRetentionDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PDB_RETENTION_DAYS_PROP_) || '');
  if (!(n > 0)) n = PDB_RETENTION_DEFAULT_;
  return Math.floor(n);
}

function pdbShrinkRatio_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PDB_SHRINK_RATIO_PROP_) || '');
  if (!(n > 0) || n >= 1) n = PDB_SHRINK_DEFAULT_;
  return n;
}

/** 昨日8:00〜今日8:00（実行日の朝8時を終点）。 */
function pdbWindowBounds_() {
  var now = new Date();
  var todayYmd = Utilities.formatDate(now, PDB_TZ_, 'yyyy-MM-dd');
  var end = Utilities.parseDate(todayYmd + ' 08:00:00', PDB_TZ_, 'yyyy-MM-dd HH:mm:ss');
  var start = new Date(end.getTime() - 24 * 60 * 60 * 1000);
  return { start: start, end: end, ymd: todayYmd };
}

function pdbFileUpdatedInWindow_(file, win) {
  if (!file || !win) return false;
  var t = file.getLastUpdated();
  if (!t) return false;
  var ms = t.getTime();
  return ms >= win.start.getTime() && ms < win.end.getTime() + 60 * 1000;
}

function pdbResolveFolderByName_(parent, name) {
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return null;
}

function pdbEnsurePjFolder_() {
  var props = PropertiesService.getScriptProperties();
  var pjId = String(props.getProperty(PDB_PJ_FOLDER_ID_PROP_) || '').trim();
  if (pjId) {
    try {
      return DriveApp.getFolderById(pjId);
    } catch (e1) {
      throw new Error('PJフォルダID不正: ' + PDB_PJ_FOLDER_ID_PROP_);
    }
  }
  var rootId = String(props.getProperty(PDB_ROOT_FOLDER_ID_PROP_) || '').trim();
  var root = null;
  if (rootId) {
    try {
      root = DriveApp.getFolderById(rootId);
    } catch (e2) {
      throw new Error('親フォルダID不正: ' + PDB_ROOT_FOLDER_ID_PROP_);
    }
  } else {
    root = pdbResolveFolderByName_(null, PDB_ROOT_FOLDER_NAME_);
    if (!root) throw new Error('Driveに「' + PDB_ROOT_FOLDER_NAME_ + '」が無い。作成するか ROOT_FOLDER_ID を設定');
  }
  var pj = pdbResolveFolderByName_(root, PDB_PJ_FOLDER_NAME_);
  if (!pj) {
    pj = root.createFolder(PDB_PJ_FOLDER_NAME_);
    Logger.log('[日次バックアップ] created pj folder=%s', pj.getId());
  }
  return pj;
}

function pdbTargetFile_() {
  var id = String(PropertiesService.getScriptProperties().getProperty(PDB_TARGET_SS_ID_PROP_) || '').trim();
  if (!id) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('対象スプシID未設定かつ Active なし');
    id = ss.getId();
  }
  return DriveApp.getFileById(id);
}

function pdbRequiredSheetNames_() {
  var raw = String(PropertiesService.getScriptProperties().getProperty(PDB_REQUIRED_SHEETS_PROP_) || '').trim();
  if (!raw) raw = PDB_REQUIRED_SHEETS_DEFAULT_;
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
function pdbCorruptSuspect_(file, folder) {
  var reason = '';
  try {
    var ss = SpreadsheetApp.openById(file.getId());
    var need = pdbRequiredSheetNames_();
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
  var prev = pdbLatestGoodCopy_(folder, file.getId());
  if (prev) {
    var ratio = pdbShrinkRatio_();
    var liveSize = Number(file.getSize()) || 0;
    var prevSize = Number(prev.getSize()) || 0;
    if (prevSize >= 50000 && liveSize < prevSize * ratio) {
      return {
        ok: false,
        reason:
          'shrink live=' + liveSize + ' prev=' + prevSize + ' ratio<' + ratio
      };
    }
  }
  return { ok: true, reason: '' };
}

function pdbLatestGoodCopy_(folder, liveId) {
  if (!folder) return null;
  var it = folder.getFiles();
  var best = null;
  var bestName = '';
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(PDB_COPY_PREFIX_) !== 0) continue;
    if (f.isTrashed()) continue;
    if (!best || name > bestName) {
      best = f;
      bestName = name;
    }
  }
  return best;
}

function pdbPurgeOldCopies_(folder, liveId, retentionDays) {
  if (!folder || !(retentionDays > 0)) return 0;
  var cutoff = new Date().getTime() - retentionDays * 24 * 60 * 60 * 1000;
  var it = folder.getFiles();
  var n = 0;
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(PDB_COPY_PREFIX_) !== 0) continue;
    var created = f.getDateCreated();
    if (created && created.getTime() < cutoff) {
      f.setTrashed(true);
      n++;
    }
  }
  return n;
}

function pdbAppendSystemLog_(state, detail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      var id = String(PropertiesService.getScriptProperties().getProperty(PDB_TARGET_SS_ID_PROP_) || '').trim();
      if (id) ss = SpreadsheetApp.openById(id);
    }
    if (!ss) return;
    var sh = ss.getSheetByName('▼ログ(システム用)');
    if (!sh) return;
    sh.appendRow([new Date(), '日次バックアップ', state, String(detail || '').slice(0, 900)]);
  } catch (eLog) {
    Logger.log('[日次バックアップ] log_fail %s', String((eLog && eLog.message) || eLog).slice(0, 120));
  }
}

function pdbReportHuman_(subject, body) {
  pdbAppendSystemLog_('REPORT', subject + ' | ' + body);
  Logger.log('[日次バックアップ] REPORT %s %s', subject, body);
  var to = String(PropertiesService.getScriptProperties().getProperty(PDB_MAIL_TO_PROP_) || '').trim();
  if (!to) return;
  try {
    MailApp.sendEmail(to, '[出品バックアップ] ' + subject, body);
  } catch (eMail) {
    Logger.log(
      '[日次バックアップ] mail_fail %s',
      String((eMail && eMail.message) || eMail).slice(0, 160)
    );
    pdbAppendSystemLog_('MAIL_FAIL', String((eMail && eMail.message) || eMail).slice(0, 200));
  }
}

/**
 * 本体。戻り値は検収・手動実行用。
 * @return {{ok:boolean, action:string, detail:string}}
 */
function runProjectDailyBackupCore_() {
  if (!pdbEnabled_()) {
    return { ok: true, action: 'disabled', detail: PDB_ENABLED_PROP_ + '=false' };
  }
  var win = pdbWindowBounds_();
  var file = pdbTargetFile_();
  var liveId = file.getId();
  if (!pdbFileUpdatedInWindow_(file, win)) {
    var skipDetail =
      'no_update_in_window lastUpdated=' +
      Utilities.formatDate(file.getLastUpdated(), PDB_TZ_, "yyyy-MM-dd'T'HH:mm:ss") +
      ' window=' +
      Utilities.formatDate(win.start, PDB_TZ_, "yyyy-MM-dd'T'HH:mm") +
      '..' +
      Utilities.formatDate(win.end, PDB_TZ_, "yyyy-MM-dd'T'HH:mm");
    pdbAppendSystemLog_('SKIP', skipDetail);
    return { ok: true, action: 'skip_no_update', detail: skipDetail };
  }
  var folder = pdbEnsurePjFolder_();
  var corrupt = pdbCorruptSuspect_(file, folder);
  if (!corrupt.ok) {
    var msg = 'corrupt_skip ' + corrupt.reason + ' liveId=' + liveId;
    pdbReportHuman_('破損疑いのためコピーしませんでした', msg);
    return { ok: true, action: 'skip_corrupt', detail: msg };
  }
  var copyName = PDB_COPY_PREFIX_ + win.ymd;
  var existing = folder.getFilesByName(copyName);
  if (existing.hasNext()) {
    var existDetail = 'already_exists name=' + copyName;
    pdbAppendSystemLog_('SKIP', existDetail);
    return { ok: true, action: 'skip_exists', detail: existDetail };
  }
  var copy = file.makeCopy(copyName, folder);
  var purged = pdbPurgeOldCopies_(folder, liveId, pdbRetentionDays_());
  var doneDetail =
    'copied id=' +
    copy.getId() +
    ' name=' +
    copyName +
    ' purged=' +
    purged +
    ' folder=' +
    folder.getId();
  pdbAppendSystemLog_('DONE', doneDetail);
  Logger.log('[日次バックアップ] %s', doneDetail);
  return { ok: true, action: 'copied', detail: doneDetail };
}

/** 時間主導: 毎日8時。 */
function runProjectDailyBackupFromTrigger() {
  try {
    var r = runProjectDailyBackupCore_();
    Logger.log('[日次バックアップ] trigger action=%s %s', r.action, r.detail);
  } catch (e) {
    var msg = String((e && e.message) || e).slice(0, 400);
    pdbReportHuman_('日次バックアップ失敗', msg);
    throw e;
  }
}

/** 手動: 今すぐ（検収用）。更新窓外でもコピーしたいときは Property で窓を広げない。手動は窓判定を通す。 */
function menuProjectDailyBackupRunNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = runProjectDailyBackupCore_();
    ui.alert('日次バックアップ', r.action + '\n' + r.detail, ui.ButtonSet.OK);
  } catch (e) {
    var msg = String((e && e.message) || e);
    pdbReportHuman_('手動バックアップ失敗', msg);
    ui.alert('失敗', msg, ui.ButtonSet.OK);
  }
}

/** 検収用: 更新窓を無視して1回コピー（破損チェックはする）。 */
function menuProjectDailyBackupForceCopy() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    '強制バックアップ',
    '更新窓を無視して今すぐコピーします（破損疑いならコピーしません）。実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;
  try {
    if (!pdbEnabled_()) {
      ui.alert('OFF', PDB_ENABLED_PROP_ + '=false', ui.ButtonSet.OK);
      return;
    }
    var file = pdbTargetFile_();
    var folder = pdbEnsurePjFolder_();
    var corrupt = pdbCorruptSuspect_(file, folder);
    if (!corrupt.ok) {
      pdbReportHuman_('強制コピー中止（破損疑い）', corrupt.reason);
      ui.alert('中止', corrupt.reason, ui.ButtonSet.OK);
      return;
    }
    var ymd = Utilities.formatDate(new Date(), PDB_TZ_, 'yyyyMMdd_HHmmss');
    var copyName = PDB_COPY_PREFIX_ + 'force_' + ymd;
    var copy = file.makeCopy(copyName, folder);
    var detail = 'force id=' + copy.getId() + ' name=' + copyName;
    pdbAppendSystemLog_('DONE', detail);
    ui.alert('完了', detail, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function menuProjectDailyBackupInstallTrigger() {
  var ui = SpreadsheetApp.getUi();
  pdbRemoveTriggers_();
  ScriptApp.newTrigger(PDB_TRIGGER_FN_).timeBased().atHour(8).everyDays(1).inTimezone(PDB_TZ_).create();
  ui.alert('日次バックアップ', '毎朝8時トリガーを設置しました（' + PDB_TRIGGER_FN_ + '）', ui.ButtonSet.OK);
}

function menuProjectDailyBackupRemoveTrigger() {
  var ui = SpreadsheetApp.getUi();
  var n = pdbRemoveTriggers_();
  ui.alert('日次バックアップ', 'トリガー削除 n=' + n, ui.ButtonSet.OK);
}

function pdbRemoveTriggers_() {
  var list = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === PDB_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(list[i]);
      n++;
    }
  }
  return n;
}
