/** @OnlyCurrentDoc */

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D10:E181').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('N10:O181').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('T10:U181').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('Z10:AA181').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('AF10:AG181').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('J3:J7').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('D10:D10').activate();
};

function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('W:W').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('W5').activate();
};

function myFunction2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('T5').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  const SCEll = 'AC3'    //開始セル
  const sh = SpreadsheetApp.getActive()
  sh.getRange(SCEll).activate()
  sh.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate() 
  // set values
  rng.offset(1, 0).activate();
  rng.offset(0, -1).activate();
  spreadsheet.getRange('T5:Y44').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  spreadsheet.getRange('AC8197').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('AC18:AG8197').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AC8197'));
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('AG37').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('AA18:AG37').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AG37'));
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('AA11000').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('AA30').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('AA13:AA30').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AA30'));
};

function myFunction3() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2:G41').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getRange('AC5:AH44').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AH44'));
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレース計算'), true);
  spreadsheet.getRange('B2').activate();
  spreadsheet.getRange('\'在庫トレースsheet\'!AC5:AH44').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};

function myFunction4() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getRange('AC5:AI44').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレース計算'), true);
  spreadsheet.getRange('\'在庫トレースsheet\'!AC5:AI44').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};

function myFunction5() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレース計算'), true);
  spreadsheet.getRange('A2:G41').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getRange('A15').activate();
  spreadsheet.getRange('\'在庫トレース計算\'!A2:G41').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};

function myFunction6() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B8').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレース計算'), true);
  spreadsheet.getRange('A:G').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('G1'));
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function myFunction7() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B1:C1').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('C1'));
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getRange('W:W').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('W5').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレース計算'), true);
  spreadsheet.getRange('A:G').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1').activate();
};

function myFunction8() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getRange('W:W').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレース計算'), true);
  spreadsheet.getRange('A:G').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getRange('W5').activate();
};

function myFunction9() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C2:G9').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function myFunction10() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C2:G9').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('N5:N9').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C2').activate();
};

function myFunction11() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K4:O8').activate();
  spreadsheet.getRange('K2:O2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('B15:AD1013').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B15').activate();
};

function myFunction12() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D4:AZ8').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Amazonデータ抽出sheet（集約）'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('仕入れ検討リスト（集約）'), true);
  spreadsheet.getRange('D4').activate();
  spreadsheet.getRange('\'Amazonデータ抽出sheet（集約）\'!D4:AZ8').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Amazonデータ抽出sheet（集約）'), true);
  spreadsheet.getRange('D2').activate();
};



/**
 * Amazonのデータをフィルタリングして仕入れ検討リストに転記し、指定のルールで転記元をリセットする関数
 */
function transferDataAndResetAdvanced() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Amazonデータ抽出sheet（集約）');
  const destinationSheet = ss.getSheetByName('仕入れ検討リスト（集約）');

  // シートが見つからない場合はエラーメッセージを出して終了
  if (!sourceSheet || !destinationSheet) {
    SpreadsheetApp.getUi().alert('必要なシートが見つかりませんでした。\n・Amazonデータ抽出sheet（集約）\n・仕入れ検討リスト（集約）');
    return;
  }

  // === STEP 1: データを取得し、条件に合う行だけを抽出 ===

  // E4:AZ8のデータを全て取得
  const sourceDataRange = sourceSheet.getRange('E4:AZ8');
  const allSourceData = sourceDataRange.getValues();

  // 転記するデータだけを入れるための空の配列を用意
  const dataToPaste = [];

  // 取得したデータを1行ずつチェック
  for (const row of allSourceData) {
    // E列から数えてK列は7番目(インデックス6)、P列は12番目(インデックス11)
    const kValue = row[6]; 
    const pValue = row[11];

    // P列(pValue)とK列(kValue)の両方に値があれば、その行を転記対象として追加
    if (kValue && pValue) {
      dataToPaste.push(row);
    }
  }

  // フィルタリングの結果、転記するデータが1行もなかった場合は処理を終了
  if (dataToPaste.length === 0) {
    SpreadsheetApp.getUi().alert('転記対象のデータが見つかりませんでした。\n(P列とK列の両方にデータがある行がありませんでした)');
    return;
  }

  // === STEP 2: フィルタリングしたデータを転記 ===

  const lastRow = destinationSheet.getLastRow();
  const startRow = Math.max(3, lastRow + 1);
  const numRowsToCopy = dataToPaste.length;
  const numColsToCopy = dataToPaste[0].length;

  // 行が不足する場合は、最終行の前に挿入
  const maxRows = destinationSheet.getMaxRows();
  if (startRow + numRowsToCopy - 1 > maxRows) {
    const rowsNeeded = (startRow + numRowsToCopy - 1) - maxRows;
    destinationSheet.insertRowsBefore(maxRows, rowsNeeded);
  }

  // 抽出したデータをE列以降に貼り付け
  destinationSheet.getRange(startRow, 5, numRowsToCopy, numColsToCopy).setValues(dataToPaste);


  // === STEP 3: 数式をコピー（転記先のA列～C列） ===
  
  const sourceFormulaRange = destinationSheet.getRange('A2:C2');
  // 転記した行数分だけ、A列～C列に数式をコピー
  const destinationFormulaRange = destinationSheet.getRange(startRow, 1, numRowsToCopy, 3);
  sourceFormulaRange.copyTo(destinationFormulaRange);


  // === STEP 4: 転記元のデータをリセット ===
  
  // K2:O2の数式をK4:O8へコピー
  sourceSheet.getRange('K2:O2').copyTo(sourceSheet.getRange('K4:O8'));
  
  // B15:AD1013の「値のみ」をクリア
  sourceSheet.getRange('B15:AD1013').clearContent();


  // === STEP 5: 最終処理（カーソル移動と通知） ===

  ss.setActiveSheet(sourceSheet);
  sourceSheet.getRange('E4').activate();
  
  SpreadsheetApp.getUi().alert('フィルタリングしたデータの転記とリセットが完了しました。');
}

function myFunction13() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4:AI11026').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function myFunction14() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4:AI11026').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A4').activate();
};

function myFunction15() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3:F9992').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A3').activate();
};

function myFunction16() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I17').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('①Keepa（抽出ｄ貼付）'), true);
  spreadsheet.getRange('A4:AI11026').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('③PV'), true);
  spreadsheet.getRange('A4:A9547').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A4').activate();  
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('①Keepa（抽出ｄ貼付）'), true);
  spreadsheet.getRange('A4').activate();
};