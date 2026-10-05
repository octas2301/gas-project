/**
 * 販促施策（楽天・Yahooクーポン等）期限リマインド
 * - シート「▼販促施策リマインド」に2四半期分を事前登録
 * - 発行・延長自体は人手（APIなし）。忘れ防止のメールのみ
 * - 宛先: ▼設定(Yahooマッピング) B17 → なければ実行ユーザー
 *
 * 復元: 本ファイル削除＋メニュー削除／トリガー削除。Git revert。
 */
var PROMO_REMIND_SHEET_ = '▼販促施策リマインド';
var PROMO_REMIND_TRIGGER_HANDLER_ = 'runPromoRemindFromTrigger';
var PROMO_REMIND_MAP_SHEET_ = '▼設定(Yahooマッピング)';
var PROMO_REMIND_MAIL_CELL_ = 'B17';
var PROMO_REMIND_LAST_KEY_ = 'PROMO_REMIND_LAST_MAIL_DAY';
var PROMO_REMIND_HEADERS_ = [
  'モール',
  '施策ID',
  '施策名',
  '種別',
  '期間ラベル',
  '開始',
  '終了',
  '状態',
  '発行済',
  '警告日前',
  '最終通知日',
  'メモ',
  '獲得URL'
];

/**
 * シート作成＋2026Q3（〜10/1 0時）と2026Q4（〜2027/1/1 0時）を投入。
 * 既存シートがある場合は確認のうえ再シード可。
 */
function menuPromoRemindInitSheet() {
  var ui = SpreadsheetApp.getUi();
  var go = ui.alert(
    '販促リマインド',
    'クーポン期限のシートを作ります。\nキャンセルなら作りません。',
    ui.ButtonSet.OK_CANCEL
  );
  if (go !== ui.Button.OK) {
    Logger.log('[menuPromoRemindInitSheet] state=CANCEL');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(PROMO_REMIND_SHEET_);
  if (existing) {
    var ans = ui.alert(
      '販促リマインド',
      '「' + PROMO_REMIND_SHEET_ + '」は既にあります。\n' +
        '先頭ヘッダ＋2四半期シードで上書きしますか？\n（既存データは消えます）',
      ui.ButtonSet.YES_NO
    );
    if (ans !== ui.Button.YES) return;
    ss.deleteSheet(existing);
  }
  var sh = ss.insertSheet(PROMO_REMIND_SHEET_);
  promoRemindWriteSeed_(sh);
  Logger.log('[menuPromoRemindInitSheet] state=DONE rows=' + (sh.getLastRow() - 1));
  ui.alert(
    '作成しました',
    PROMO_REMIND_SHEET_ + '\n' +
      '・2026Q3（終了 2026/10/01 0:00）= 稼働中・発行済\n' +
      '・2026Q4（終了 2027/01/01 0:00）= 次期予定・未発行\n' +
      '忘れても次期行が残るので再発行の控えになります。\n' +
      '続けて「日次トリガー設置」を実行してください。',
    ui.ButtonSet.OK
  );
}

/** 今すぐ判定＋メール（テスト可） */
function menuPromoRemindRunNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = promoRemindRun_(false);
    Logger.log('[menuPromoRemindRunNow] ' + JSON.stringify(r));
    ui.alert(
      '販促リマインド実行',
      '対象行: ' + r.hit + ' / 全データ: ' + r.total + '\n' +
        'メール: ' + (r.mailOk ? '送信OK' : (r.mailSkip || r.mailError || 'なし')) +
        (r.to ? '\n宛先: ' + r.to : ''),
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log('[menuPromoRemindRunNow] FAILED ' + ((e && e.message) || e));
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

/** 強制テスト: 警告条件を無視して稼働中・次期予定をメールに載せる */
function menuPromoRemindTestMailForce() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = promoRemindRun_(true);
    Logger.log('[menuPromoRemindTestMailForce] ' + JSON.stringify(r));
    ui.alert(
      '強制テストメール',
      '掲載行: ' + r.hit + '\nメール: ' + (r.mailOk ? '送信OK' : (r.mailError || '失敗')) +
        (r.to ? '\n宛先: ' + r.to : ''),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('失敗', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

function menuPromoRemindInstallTrigger() {
  var ui = SpreadsheetApp.getUi();
  promoRemindRemoveTriggers_();
  ScriptApp.newTrigger(PROMO_REMIND_TRIGGER_HANDLER_)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  Logger.log('[menuPromoRemindInstallTrigger] state=DONE');
  ui.alert('日次トリガーを設置しました（毎日9時前後・' + PROMO_REMIND_TRIGGER_HANDLER_ + '）。');
}

function menuPromoRemindRemoveTrigger() {
  var n = promoRemindRemoveTriggers_();
  Logger.log('[menuPromoRemindRemoveTrigger] removed=' + n);
  try {
    SpreadsheetApp.getUi().alert('販促リマインドトリガーを削除しました（' + n + '件）。');
  } catch (e) {}
}

/** トリガー本体 */
function runPromoRemindFromTrigger() {
  var fn = 'runPromoRemindFromTrigger';
  Logger.log('[' + fn + '] state=PENDING');
  try {
    if (!promoRemindEnabled_()) {
      Logger.log('[' + fn + '] state=SKIPPED reason=disabled');
      return;
    }
    var r = promoRemindRun_(false);
    Logger.log('[' + fn + '] state=DONE hit=' + r.hit + ' mailOk=' + r.mailOk + ' skip=' + (r.mailSkip || ''));
  } catch (e) {
    Logger.log('[' + fn + '] state=FAILED ' + ((e && e.message) || e));
  }
}

function promoRemindEnabled_() {
  var v = String(PropertiesService.getScriptProperties().getProperty('PROMO_REMIND_ENABLED') || 'true').trim().toLowerCase();
  return v !== 'false' && v !== '0' && v !== 'off';
}

function promoRemindRemoveTriggers_() {
  var n = 0;
  var list = ScriptApp.getProjectTriggers();
  for (var i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === PROMO_REMIND_TRIGGER_HANDLER_) {
      ScriptApp.deleteTrigger(list[i]);
      n++;
    }
  }
  return n;
}

function promoRemindWriteSeed_(sh) {
  sh.clear();
  sh.getRange(1, 1, 1, PROMO_REMIND_HEADERS_.length).setValues([PROMO_REMIND_HEADERS_]);
  sh.setFrozenRows(1);
  sh.getRange('F:G').setNumberFormat('yyyy/mm/dd hh:mm');

  var q3Start = new Date(2026, 8, 13, 13, 0, 0); // month 0-index
  var q3End = new Date(2026, 9, 1, 0, 0, 0); // 2026-10-01 00:00
  var q4Start = new Date(2026, 9, 1, 0, 0, 0);
  var q4End = new Date(2027, 0, 1, 0, 0, 0); // 2027-01-01 00:00

  var defs = [
    ['Yahoo', 'YH_CONV_50', '1500円以上で50円OFF', 'クーポン', '転換・公開円引き'],
    ['Yahoo', 'YH_MATOME_2_5', '2点以上で5%OFF（上限150円）', 'クーポン', '客単価・まとめ'],
    ['Yahoo', 'YH_FIRST_3', '初来店3%OFF（上限150円）', 'クーポン', '新規・未購入'],
    ['Yahoo', 'YH_THANKS_3', 'サンキュー3%クーポン', 'クーポン', 'リピート・非表示'],
    ['Yahoo', 'YH_PR_REVIEW', 'PRオプション見直し（パッケージ2%相当＋有料は主力のみ）', '広告', '有料一律は原則なし'],
    ['楽天', 'RK_COUPON_REVIEW', '楽天クーポン／販促の四半期見直し', 'クーポン', '中身は都度記入']
  ];

  var rows = [];
  for (var i = 0; i < defs.length; i++) {
    var d = defs[i];
    rows.push([d[0], d[1], d[2], d[3], '2026Q3', q3Start, q3End, '稼働中', true, 7, '', d[4], '']);
    rows.push([d[0], d[1], d[2], d[3], '2026Q4', q4Start, q4End, '次期予定', false, 7, '', '未発行・Q3終了前に発行', '']);
  }
  sh.getRange(2, 1, rows.length, PROMO_REMIND_HEADERS_.length).setValues(rows);
  sh.autoResizeColumns(1, PROMO_REMIND_HEADERS_.length);
}

/**
 * @param {boolean} forceAll 警告日数を無視して稼働中・次期予定を載せる
 */
function promoRemindRun_(forceAll) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PROMO_REMIND_SHEET_);
  if (!sh) throw new Error('シート「' + PROMO_REMIND_SHEET_ + '」がありません。先に初期化してください。');

  var last = sh.getLastRow();
  if (last < 2) return { total: 0, hit: 0, mailOk: false, mailSkip: 'empty' };

  var values = sh.getRange(2, 1, last, PROMO_REMIND_HEADERS_.length).getValues();
  var today = promoRemindTodayTokyo_();
  var hits = [];
  var notifyCol = []; // row index in sheet (1-based) for K update

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = String(row[7] || '').trim();
    if (status !== '稼働中' && status !== '次期予定') continue;
    var end = promoRemindToDate_(row[6]);
    if (!end) continue;
    var warnDays = parseInt(String(row[9] != null && row[9] !== '' ? row[9] : 7), 10);
    if (!(warnDays >= 0)) warnDays = 7;
    var daysLeft = promoRemindDaysLeft_(today, end);
    var hit = forceAll || daysLeft <= warnDays;
    if (!hit) continue;
    hits.push({
      mall: row[0],
      id: row[1],
      name: row[2],
      kind: row[3],
      label: row[4],
      status: status,
      issued: row[8] === true || String(row[8]).toUpperCase() === 'TRUE',
      end: end,
      daysLeft: daysLeft,
      memo: row[11],
      sheetRow: i + 2
    });
    notifyCol.push(i + 2);
  }

  if (!hits.length) {
    return { total: values.length, hit: 0, mailOk: false, mailSkip: 'no_hits' };
  }

  var to = promoRemindMailTo_();
  if (!to) return { total: values.length, hit: hits.length, mailOk: false, mailError: '宛先なし（マッピングB17）' };

  if (!forceAll) {
    var props = PropertiesService.getScriptProperties();
    var dayKey = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
    var sig = dayKey + ':' + hits.map(function (h) { return h.id + h.label + h.status; }).join('|');
    if (props.getProperty(PROMO_REMIND_LAST_KEY_) === sig) {
      return { total: values.length, hit: hits.length, mailOk: true, mailSkip: 'suppressed_same_day', to: to };
    }
    props.setProperty(PROMO_REMIND_LAST_KEY_, sig);
  }

  var body = promoRemindBuildBody_(hits, today);
  try {
    MailApp.sendEmail(to, '【販促リマインド】クーポン／施策の期限確認', body);
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    for (var j = 0; j < notifyCol.length; j++) {
      sh.getRange(notifyCol[j], 11).setValue(nowStr); // K 最終通知日
    }
    return { total: values.length, hit: hits.length, mailOk: true, to: to };
  } catch (e) {
    return { total: values.length, hit: hits.length, mailOk: false, mailError: String((e && e.message) || e), to: to };
  }
}

function promoRemindBuildBody_(hits, today) {
  var lines = [];
  lines.push('楽天・Yahoo! 販促施策の期限リマインドです（発行の自動延長はできません）。');
  lines.push('基準日: ' + Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd'));
  lines.push('');
  lines.push('■ 通知対象');
  for (var i = 0; i < hits.length; i++) {
    var h = hits[i];
    lines.push(
      '- [' + h.mall + '] ' + h.name +
        ' / ' + h.label +
        ' / ' + h.status +
        ' / 発行済=' + (h.issued ? 'Y' : 'N') +
        ' / 終了=' + Utilities.formatDate(h.end, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') +
        ' / 残=' + h.daysLeft + '日' +
        (h.memo ? ' / ' + h.memo : '')
    );
  }
  lines.push('');
  lines.push('■ やること');
  lines.push('1. 稼働中で残日数が少ない → ストアで条件どおり再発行（または次期のクーポンを発行）');
  lines.push('2. 次期予定かつ発行済=N → 今のうちに発行し、シートの発行済をTRUE・状態を稼働中へ');
  lines.push('3. 終わった四半期は状態=終了。次の次の四半期行を1行足す（常に2回分残す）');
  lines.push('');
  lines.push('シート: ' + PROMO_REMIND_SHEET_);
  lines.push('メニュー: 🛒 Yahoo!出品 → 販促リマインド…');
  return lines.join('\n');
}

function promoRemindMailTo_() {
  var to = '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var map = ss.getSheetByName(PROMO_REMIND_MAP_SHEET_);
    if (map) to = String(map.getRange(PROMO_REMIND_MAIL_CELL_).getValue() || '').trim();
  } catch (e0) {}
  if (!to) {
    try {
      to = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
    } catch (e1) {}
  }
  return to;
}

function promoRemindTodayTokyo_() {
  var s = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var p = s.split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function promoRemindToDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var s = String(v || '').trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[ T](\d{1,2}):(\d{1,2}))?/);
  if (!m) return null;
  return new Date(
    Number(m[1]),
    Number(m[2]) - 1,
    Number(m[3]),
    m[4] != null ? Number(m[4]) : 0,
    m[5] != null ? Number(m[5]) : 0,
    0
  );
}

/** 終了日時までの暦日（終了日当日=0、過ぎたら負） */
function promoRemindDaysLeft_(today, end) {
  var e = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.round((e.getTime() - today.getTime()) / 86400000);
}
