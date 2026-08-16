function copyAndMoveCursor2() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Amazonデータ抽出sheet');
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('商品入力シート');
  
  var sourceRange = sourceSheet.getRange("C3:P7");
  var targetColumn = 1; // Column A

  // 空白がある行の検索
  var targetStartRow = findEmptyRow(targetSheet, targetColumn);

  // ソースの値を取得
  var sourceValues = sourceRange.getValues();

  // ターゲット範囲にソースの値を複製
  targetSheet.getRange(targetStartRow, targetColumn, sourceValues.length, sourceValues[0].length).setValues(sourceValues);

  // 貼り付けたセルにカーソルを移動
  targetSheet.getRange(targetStartRow, targetColumn).activate();
}

// 空白がある行を検索する関数
function findEmptyRow(sheet, column) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 6) {
    return 6;
  }
  
  var data = sheet.getRange(6, column, lastRow - 5, 1).getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "") {
      return i + 6;
    }
  }
  
  return lastRow + 1; // 空白が見つからなかった場合は次の行を返す
}
