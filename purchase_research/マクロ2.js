function CopyTitleOnLast() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getRange('AC5:AI44').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレース計算'), true);
  spreadsheet.getRange('\'在庫トレースsheet\'!AC5:AI44').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const shSource = ss.getSheetByName("在庫トレース計算")
  const lclm = shSource.getLastColumn()

  console.log(lclm)

  const sdata = shSource.getRange(1,1,41,lclm)  //コピー元のデータ

   console.log(sdata)

  const shDest = ss.getSheetByName("在庫トレースsheet")
  let lRow = shDest.getLastRow()

  console.log(lRow)

  // console.log( sh_day.getRange(2,2).getDisplayValue())


  //A列をすべて取得する
  const dataA = shDest.getRange(1,1,lRow,1).getValues()

  console.log(dataA)

  //最下行から入力されている行を取得する
  let insertRow = 0


  console.log("B1はdataA[0][0]→",dataA[lRow-1][0])

  for(i=lRow-1;i>=0;i--){
    if(dataA[i][0] != ""){
      insertRow = i+1+1   //0オリジンのため+1、さらに2行下なので+2する
      console.log(i)
      break
    }
  }
      console.log(insertRow)

  //タイトル行のコピー
  sdata.copyTo(shDest.getRange(insertRow,1))

  //today()に今日の日付を上書きする
  shDest.getRange(insertRow+1,1).setValue(new Date())

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('在庫トレースsheet'), true);
  spreadsheet.getRange('W5').activate();

}
