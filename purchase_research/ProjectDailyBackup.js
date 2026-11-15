/**
 * プロジェクト別 Drive 自動バックアップ（仕入れ検討①・第1版）
 * 正本: docs/org/PROJECT_DAILY_BACKUP_REQUIREMENTS.md
 *
 * 対象: ①調査複製 ＋ 競合DB（各 Property）
 * 切戻し: PR_PROJECT_DAILY_BACKUP_ENABLED=false
 */

var PR_PDB_ENABLED_PROP_ = 'PR_PROJECT_DAILY_BACKUP_ENABLED';
var PR_PDB_ROOT_FOLDER_ID_PROP_ = 'PR_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID';
var PR_PDB_PJ_FOLDER_ID_PROP_ = 'PR_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID';
var PR_PDB_RESEARCH_SS_ID_PROP_ = 'PR_PROJECT_DAILY_BACKUP_RESEARCH_SS_ID';
var PR_PDB_COMPETITOR_SS_ID_PROP_ = 'PR_PROJECT_DAILY_BACKUP_COMPETITOR_SS_ID';
var PR_PDB_RETENTION_DAYS_PROP_ = 'PR_PROJECT_DAILY_BACKUP_RETENTION_DAYS';
var PR_PDB_SHRINK_RATIO_PROP_ = 'PR_PROJECT_DAILY_BACKUP_SHRINK_RATIO';
var PR_PDB_MAIL_TO_PROP_ = 'PR_PROJECT_DAILY_BACKUP_MAIL_TO';

var PR_PDB_ROOT_FOLDER_NAME_ = '99.バックアップ';
var PR_PDB_PJ_FOLDER_NAME_ = '仕入れ検討①';
var PR_PDB_RETENTION_DEFAULT_ = 14;
var PR_PDB_SHRINK_DEFAULT_ = 0.8;
var PR_PDB_TRIGGER_FN_ = 'prRunProjectDailyBackupFromTrigger';
var PR_PDB_TZ_ = 'Asia/Tokyo';

/** 調査複製の既定必須シート */
var PR_PDB_RESEARCH_SHEETS_DEFAULT_ = '①ログ|①候補';
/** 競合DBの既定必須シート */
var PR_PDB_COMP_SHEETS_DEFAULT_ = 'Keepaフル|メーカーサマリ';

function prPdbEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR_PDB_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return true;
  var s = String(raw).trim().toUpperCase();
  return !(s === 'FALSE' || s === '0' || s === 'OFF' || s === 'いいえ');
}

function prPdbRetentionDays_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_PDB_RETENTION_DAYS_PROP_) || '');
  if (!(n > 0)) n = PR_PDB_RETENTION_DEFAULT_;
  return Math.floor(n);
}

function prPdbShrinkRatio_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(PR_PDB_SHRINK_RATIO_PROP_) || '');
  if (!(n > 0) || n >= 1) n = PR_PDB_SHRINK_DEFAULT_;
  return n;
}

function prPdbWindowBounds_() {
  var now = new Date();
  var todayYmd = Utilities.formatDate(now, PR_PDB_TZ_, 'yyyy-MM-dd');
  var end = Utilities.parseDate(todayYmd + ' 08:00:00', PR_PDB_TZ_, 'yyyy-MM-dd HH:mm:ss');
  var start = new Date(end.getTime() - 24 * 60 * 60 * 1000);
  return { start: start, end: end, ymd: todayYmd };
}

function prPdbFileUpdatedInWindow_(file, win) {
  if (!file || !win) return false;
  var t = file.getLastUpdated();
  if (!t) return false;
  var ms = t.getTime();
  return ms >= win.start.getTime() && ms < win.end.getTime() + 60 * 1000;
}

function prPdbResolveFolderByName_(parent, name) {
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return null;
}

function prPdbEnsurePjFolder_() {
  var props = PropertiesService.getScriptProperties();
  var pjId = String(props.getProperty(PR_PDB_PJ_FOLDER_ID_PROP_) || '').trim();
  if (pjId) {
    try {
      return DriveApp.getFolderById(pjId);
    } catch (e1) {
      throw new Error('PJフォルダID不正: ' + PR_PDB_PJ_FOLDER_ID_PROP_);
    }
  }
  var rootId = String(props.getProperty(PR_PDB_ROOT_FOLDER_ID_PROP_) || '').trim();
  var root = null;
  if (rootId) {
    try {
      root = DriveApp.getFolderById(rootId);
    } catch (e2) {
      throw new Error('親フォルダID不正: ' + PR_PDB_ROOT_FOLDER_ID_PROP_);
    }
  } else {
    root = prPdbResolveFolderByName_(null, PR_PDB_ROOT_FOLDER_NAME_);
    if (!root) throw new Error('Driveに「' + PR_PDB_ROOT_FOLDER_NAME_ + '」が無い');
  }
  var pj = prPdbResolveFolderByName_(root, PR_PDB_PJ_FOLDER_NAME_);
  if (!pj) pj = root.createFolder(PR_PDB_PJ_FOLDER_NAME_);
  return pj;
}

function prPdbResearchSsId_() {
  var id = String(PropertiesService.getScriptProperties().getProperty(PR_PDB_RESEARCH_SS_ID_PROP_) || '').trim();
  if (id) return id;
  // ①は bound script。未設定なら Active（調査複製）を正とする
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss.getId();
  throw new Error(PR_PDB_RESEARCH_SS_ID_PROP_ + ' 未設定かつ Active なし');
}

function prPdbCompetitorSsId_() {
  var id = String(PropertiesService.getScriptProperties().getProperty(PR_PDB_COMPETITOR_SS_ID_PROP_) || '').trim();
  if (id) return id;
  if (typeof PR_COMPETITOR_SS_ !== 'undefined' && PR_COMPETITOR_SS_) return String(PR_COMPETITOR_SS_);
  throw new Error(PR_PDB_COMPETITOR_SS_ID_PROP_ + ' 未設定');
}

function prPdbSplitSheets_(raw, fallback) {
  var s = String(raw || '').trim() || fallback;
  return s
    .split('|')
    .map(function (x) {
      return String(x || '').trim();
    })
    .filter(function (x) {
      return !!x;
    });
}

function prPdbCorruptSuspect_(file, folder, prefix, requiredSheets) {
  try {
    var ss = SpreadsheetApp.openById(file.getId());
    var i;
    for (i = 0; i < requiredSheets.length; i++) {
      if (!ss.getSheetByName(requiredSheets[i])) {
        return { ok: false, reason: 'required_sheet_missing=' + requiredSheets[i] };
      }
    }
  } catch (eOpen) {
    return {
      ok: false,
      reason: 'open_failed=' + String((eOpen && eOpen.message) || eOpen).slice(0, 120)
    };
  }
  var prev = prPdbLatestGoodCopy_(folder, file.getId(), prefix);
  if (prev) {
    var ratio = prPdbShrinkRatio_();
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

function prPdbLatestGoodCopy_(folder, liveId, prefix) {
  if (!folder) return null;
  var it = folder.getFiles();
  var best = null;
  var bestName = '';
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(prefix) !== 0) continue;
    if (f.isTrashed()) continue;
    if (!best || name > bestName) {
      best = f;
      bestName = name;
    }
  }
  return best;
}

function prPdbPurgeOldCopies_(folder, liveId, prefix, retentionDays) {
  if (!folder || !(retentionDays > 0)) return 0;
  var cutoff = new Date().getTime() - retentionDays * 24 * 60 * 60 * 1000;
  var it = folder.getFiles();
  var n = 0;
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() === liveId) continue;
    var name = String(f.getName() || '');
    if (name.indexOf(prefix) !== 0) continue;
    var created = f.getDateCreated();
    if (created && created.getTime() < cutoff) {
      f.setTrashed(true);
      n++;
    }
  }
  return n;
}

function prPdbAppendLog_(state, detail) {
  try {
    if (typeof prAppendLog_ === 'function') {
      var runId =
        'pr_bak_' + Utilities.formatDate(new Date(), PR_PDB_TZ_, 'yyyyMMdd_HHmmss');
      prAppendLog_([runId, '日次バックアップ', state, new Date(), '', '', '', String(detail || '').slice(0, 200)]);
      return;
    }
  } catch (e1) {}
  Logger.log('[①日次バックアップ] %s %s', state, detail);
}

function prPdbReportHuman_(subject, body) {
  prPdbAppendLog_('REPORT', subject + ' | ' + body);
  Logger.log('[①日次バックアップ] REPORT %s %s', subject, body);
  var to = String(PropertiesService.getScriptProperties().getProperty(PR_PDB_MAIL_TO_PROP_) || '').trim();
  if (!to) return;
  try {
    MailApp.sendEmail(to, '[①バックアップ] ' + subject, body);
  } catch (eMail) {
    prPdbAppendLog_('MAIL_FAIL', String((eMail && eMail.message) || eMail).slice(0, 200));
  }
}

/**
 * 1ファイル分。
 * @return {{ok:boolean, action:string, detail:string}}
 */
function prPdbBackupOne_(opt) {
  opt = opt || {};
  var label = String(opt.label || 'file');
  var prefix = String(opt.prefix || 'backup_');
  var fileId = String(opt.fileId || '').trim();
  var sheets = opt.sheets || [];
  var win = opt.win;
  var folder = opt.folder;
  var force = !!opt.force;

  var file = DriveApp.getFileById(fileId);
  if (!force && !prPdbFileUpdatedInWindow_(file, win)) {
    var skip =
      label +
      ' skip_no_update last=' +
      Utilities.formatDate(file.getLastUpdated(), PR_PDB_TZ_, "yyyy-MM-dd'T'HH:mm:ss");
    prPdbAppendLog_('SKIP', skip);
    return { ok: true, action: 'skip_no_update', detail: skip };
  }
  var corrupt = prPdbCorruptSuspect_(file, folder, prefix, sheets);
  if (!corrupt.ok) {
    var msg = label + ' corrupt_skip ' + corrupt.reason;
    prPdbReportHuman_('破損疑いのためコピーしませんでした', msg);
    return { ok: true, action: 'skip_corrupt', detail: msg };
  }
  var copyName = force
    ? prefix + 'force_' + Utilities.formatDate(new Date(), PR_PDB_TZ_, 'yyyyMMdd_HHmmss')
    : prefix + win.ymd;
  if (!force) {
    var existing = folder.getFilesByName(copyName);
    if (existing.hasNext()) {
      var ex = label + ' already_exists ' + copyName;
      prPdbAppendLog_('SKIP', ex);
      return { ok: true, action: 'skip_exists', detail: ex };
    }
  }
  var copy = file.makeCopy(copyName, folder);
  var purged = prPdbPurgeOldCopies_(folder, fileId, prefix, prPdbRetentionDays_());
  var done =
    label +
    ' copied id=' +
    copy.getId() +
    ' name=' +
    copyName +
    ' purged=' +
    purged;
  prPdbAppendLog_('DONE', done);
  return { ok: true, action: 'copied', detail: done };
}

function prRunProjectDailyBackupCore_(opt) {
  opt = opt || {};
  var force = !!opt.force;
  if (!prPdbEnabled_()) {
    return { ok: true, action: 'disabled', detail: PR_PDB_ENABLED_PROP_ + '=false', results: [] };
  }
  var win = prPdbWindowBounds_();
  var folder = prPdbEnsurePjFolder_();
  var results = [];
  results.push(
    prPdbBackupOne_({
      label: '①調査',
      prefix: '①調査_backup_',
      fileId: prPdbResearchSsId_(),
      sheets: prPdbSplitSheets_('', PR_PDB_RESEARCH_SHEETS_DEFAULT_),
      win: win,
      folder: folder,
      force: force
    })
  );
  results.push(
    prPdbBackupOne_({
      label: '競合DB',
      prefix: '競合DB_backup_',
      fileId: prPdbCompetitorSsId_(),
      sheets: prPdbSplitSheets_('', PR_PDB_COMP_SHEETS_DEFAULT_),
      win: win,
      folder: folder,
      force: force
    })
  );
  return { ok: true, action: 'batch', detail: 'n=' + results.length, results: results };
}

function prRunProjectDailyBackupFromTrigger() {
  try {
    var r = prRunProjectDailyBackupCore_({});
    Logger.log('[①日次バックアップ] trigger %s', JSON.stringify(r.results || []));
  } catch (e) {
    prPdbReportHuman_('日次バックアップ失敗', String((e && e.message) || e).slice(0, 400));
    throw e;
  }
}

function prMenuProjectDailyBackupRunNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = prRunProjectDailyBackupCore_({});
    var lines = (r.results || [])
      .map(function (x) {
        return x.action + ': ' + x.detail;
      })
      .join('\n');
    ui.alert('①日次バックアップ', lines || r.detail, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function prMenuProjectDailyBackupForceCopy() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    '強制バックアップ',
    '更新窓を無視して①調査＋競合DBをコピーします。実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;
  try {
    var r = prRunProjectDailyBackupCore_({ force: true });
    var lines = (r.results || [])
      .map(function (x) {
        return x.action + ': ' + x.detail;
      })
      .join('\n');
    ui.alert('完了', lines, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function prMenuProjectDailyBackupInstallTrigger() {
  var ui = SpreadsheetApp.getUi();
  prPdbRemoveTriggers_();
  ScriptApp.newTrigger(PR_PDB_TRIGGER_FN_)
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone(PR_PDB_TZ_)
    .create();
  ui.alert('①日次バックアップ', '毎朝8時トリガーを設置しました', ui.ButtonSet.OK);
}

function prMenuProjectDailyBackupRemoveTrigger() {
  var ui = SpreadsheetApp.getUi();
  var n = prPdbRemoveTriggers_();
  ui.alert('①日次バックアップ', 'トリガー削除 n=' + n, ui.ButtonSet.OK);
}

function prPdbRemoveTriggers_() {
  var list = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === PR_PDB_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(list[i]);
      n++;
    }
  }
  return n;
}
