// ==========================================
// 在庫管理アプリ - 在庫移動日報メール（毎日20:00）
// ==========================================

/**
 * 本日の在庫移動リストを取得し、1件以上あれば日報メールを送信する。
 * 宛先は「日報宛先」シートから取得。送信時刻は 20:00 Asia/Tokyo でトリガー登録する想定。
 */
function sendDailyMoveReport() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('在庫移動リスト');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var h = data[0];
  var iDate = h.indexOf('移動日時');
  var iJan = h.indexOf('単品JAN');
  var iName = h.indexOf('商品名');
  var iFrom = h.indexOf('移動元');
  var iTo = h.indexOf('移動先');
  var iQty = h.indexOf('移動数量');
  var iTantou = h.indexOf('担当者名');
  var iMoveDay = h.indexOf('移動日');
  if (iDate < 0) iDate = 1;
  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
  var rows = [];
  for (var r = 1; r < data.length; r++) {
    var d = data[r][iDate];
    var dayStr = d instanceof Date ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd') : (data[r][iMoveDay] || '').toString();
    if (dayStr !== todayStr) continue;
    rows.push({
      time: d instanceof Date ? Utilities.formatDate(d, 'Asia/Tokyo', 'HH:mm') : '',
      jan: (data[r][iJan] || '').toString(),
      name: (data[r][iName] || '').toString(),
      from: (data[r][iFrom] || '').toString(),
      to: (data[r][iTo] || '').toString(),
      qty: (data[r][iQty] || '').toString(),
      tantou: (data[r][iTantou] || '').toString()
    });
  }
  if (rows.length === 0) return;
  var toList = getReportEmailList();
  if (toList.length === 0) return;
  var body = '本日の在庫移動リストです。<br><br><table border="1" cellpadding="4" cellspacing="0"><tr><th>時間</th><th>JAN</th><th>商品名</th><th>数量</th><th>移動元</th><th>移動先</th><th>担当</th></tr>';
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    body += '<tr><td>' + row.time + '</td><td>' + row.jan + '</td><td>' + row.name + '</td><td>' + row.qty + '</td><td>' + row.from + '</td><td>' + row.to + '</td><td>' + row.tantou + '</td></tr>';
  }
  body += '</table>';
  MailApp.sendEmail({
    to: toList.join(','),
    subject: '【在庫移動】本日の移動報告 ' + todayStr,
    htmlBody: body
  });
}

/**
 * 日次トリガーを登録する（毎日 20:00 Asia/Tokyo）。
 * 手動で1回実行するか、GAS エディタの「トリガー」から設定する。
 */
function installDailyReportTrigger() {
  var existing = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'sendDailyMoveReport';
  });
  if (existing.length > 0) {
    existing.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  }
  ScriptApp.newTrigger('sendDailyMoveReport')
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .inTimezone('Asia/Tokyo')
    .create();
}
