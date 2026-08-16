function copyAndMoveCursor() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sourceRange = sheet.getRange("I5:N9");
  var targetColumn = 2; // Column B
  var targetStartRow = 15;

  // 空白がある行の検索
  for (var row = targetStartRow; row <= 10016; row++) {
    var cellValue = sheet.getRange(row, targetColumn).getValue();
    if (cellValue === "") {
      targetStartRow = row;
      break;
    }
  }

  // ソースの値を取得
  var sourceValues = sourceRange.getValues();

  // ターゲット範囲にソースの値を複製
  sheet.getRange(targetStartRow, targetColumn, sourceValues.length, sourceValues[0].length).setValues(sourceValues);

  // 貼り付けたセルにカーソルを移動
  sheet.getRange(targetStartRow + sourceValues.length, targetColumn).activate();
}

