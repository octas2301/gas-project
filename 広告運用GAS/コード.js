/**
 * 【Amazon/楽天 完全統合・最終決定版】広告運用システム
 * * アップデート内容：
 * 1. Amazon分析ロジックを「SUM（合計）集計」に変更。マッチタイプごとの行を合算して評価。
 * 2. 📌運用マニュアル作成機能 (createManualSheet) を復活・統合。
 * 3. 重複排除、スペース無視、親子リレーID取得、指列表記防止の全機能を完備。
 */

// --- 設定 ---
const CONFIG = {
  AMZ_ROAS_GOOD: 1500,      
  AMZ_ROAS_KEEP: 1000,      
  CLICK_THRESHOLD: 2,       
  RAK_KW_MIN_BID: 41, 
  RAK_PROD_MIN_BID: 21,
  CLEANUP_DAYS: 60,
  TARGET_CVR: 10,
  RAK_ROAS_GOOD: 2000,
  RAK_ROAS_LIMIT: 1000 
};

const AMZ_FOLDER_ID = "1-N2z4FlJtqgegLOyV7_7PHQSFo4akQjx";
const RAK_FOLDER_ID = "1R8cIvrwU2PVW1tfGNrfOl5LyevmPGOZb";
const OUTPUT_FOLDER_NAME = "📦生成されたバルクファイル";

const PRODUCT_MAP = {
  "五色あられ 小袋ミニ 200g 業務用 GFC  湿気知らずの小容量パック ぶぶあられ キャラ弁 飾り 弁当 雛祭り": "gfc-nam-200g",
  "石原水産 まぐろチーズ 220g おつまみ 個包装 お菓子": "ishihara-maguro-220",
  "石原水産 チーズかつお 245g おつまみ 個包装 お菓子": "ishihara-katsuo-245",
  "サバトン マロンクリーム 1kg 製菓材料 プロの味を再現 モンブラン 栗 缶詰 フランス ケーキ スイーツ": "sabaton-m-1000",
  "ストレス イライラ リラックス 度チェッカー 5枚 ライフケア技研 [富山大医学部 共同開発] 解消 グッズ": "stress-checker-5"
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 広告運用メニュー')
    .addItem('1-2. 分析＆バルク生成を実行', 'runOptimizationAndBulk')
    .addItem('1. 分析を実行する', 'runOptimization')
    .addSeparator()
    .addItem('2. バルク生成＆ドライブ保存', 'generateBulkFiles')
    .addItem('🔍 バルクファイル事前診断', 'checkBulkFileErrors')
    .addSeparator()
    .addItem('3. 【新機能】お宝キーワード広告CSV作成', 'generateRakutenKeywordAdCSV')
    .addSeparator()
    .addItem('🔧 操作マニュアルを作成', 'createManualSheet')
    .addItem('🧹 古いファイルの掃除', 'runManualCleanup')
    .addToUi();
}

/**
 * 1. 分析実行 (AmazonはSUM集計方式)
 */
function runOptimization(skipUi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  const sHist = getSheetByNameSafe(ss, "履歴データ");
  if (sHist.getLastRow() === 0) sHist.appendRow(["実行日", "モール", "識別子", "ROAS", "CVR", "ステータス"]);
  const histMap = getHistoryMap(sHist);

  const sheets = {
    amzNeg: getSheetByNameSafe(ss, "除外候補（Amazon）").clear(),
    amzPos: getSheetByNameSafe(ss, "入札調整候補（Amazon）").clear(), 
    rakNeg: getSheetByNameSafe(ss, "除外候補（楽天）").clear(),
    rakPos: getSheetByNameSafe(ss, "入札強化候補（楽天）").clear(),
    amzRaw: getSheetByNameSafe(ss, "元データ_Amazon").clear(),
    rakRaw: getSheetByNameSafe(ss, "元データ_楽天").clear()
  };

  let amzNegData = [], amzPosData = [], rakNegData = [], rakPosData = [], histNew = [];

  // --- Amazon処理（SUM集計方式） ---
  try {
    const amzFiles = DriveApp.getFolderById(AMZ_FOLDER_ID).getFiles();
    while(amzFiles.hasNext()){
      let f = amzFiles.next();
      if(f.getMimeType() !== MimeType.CSV && !f.getName().endsWith(".csv")) continue;
      
      let csvString = f.getBlob().getDataAsString("UTF-8");
      if (csvString.indexOf("キャンペーン") === -1 && csvString.indexOf("Campaign") === -1) {
         csvString = f.getBlob().getDataAsString("Shift-JIS");
      }
      const csvData = Utilities.parseCsv(csvString);
      if (csvData.length > 0) sheets.amzRaw.getRange(1, 1, Math.min(csvData.length, 20000), csvData[0].length).setValues(csvData.slice(0, 20000));

      const headerRow = csvData[0];
      const col = {
        term: findColIndex(headerRow, ["カスタマー検索用語", "Customer Search Term"]),
        clicks: findColIndex(headerRow, ["クリック数", "Clicks"]),
        imp: findColIndex(headerRow, ["インプレッション", "Impressions"]),
        sales: findColIndex(headerRow, ["売上", "Sales"]),
        orders: findColIndex(headerRow, ["注文", "Orders"]),
        cost: findColIndex(headerRow, ["支出", "Spend", "費用"]),
        cId: findColIndex(headerRow, ["キャンペーンID", "Campaign ID"]),
        agId: findColIndex(headerRow, ["広告グループID", "Ad Group ID"]),
        cName: findColIndex(headerRow, ["キャンペーン名", "Campaign Name"]),
        agName: findColIndex(headerRow, ["広告グループ名", "Ad Group Name"]),
        kwText: findColIndex(headerRow, ["キーワードテキスト", "Keyword Text"]),
        kwId: findColIndex(headerRow, ["キーワードID", "Keyword Id"])
      };

      // SUM集計用オブジェクト
      let amzSumMap = {};

      csvData.slice(1).forEach(row => {
        let campName = col.cName !== -1 ? row[col.cName] : "";
        let adGroupName = col.agName !== -1 ? row[col.agName] : "";
        let isManual = (campName.indexOf("Manual") !== -1 || campName.indexOf("マニュアル") !== -1);
        let kw = isManual ? (row[col.kwText] || row[col.term]) : row[col.term];
        if (!kw) return;

        let key = adGroupName + "|||" + kw; // 広告グループごとにキーワードを合算
        if (!amzSumMap[key]) {
          amzSumMap[key] = {
            campName: campName, adGroupName: adGroupName, kw: kw,
            imp: 0, clicks: 0, orders: 0, sales: 0, cost: 0,
            campId: row[col.cId], adGroupId: row[col.agId], kwId: row[col.kwId],
            isManual: isManual
          };
        }
        amzSumMap[key].imp += parseInt(row[col.imp]) || 0;
        amzSumMap[key].clicks += parseInt(row[col.clicks]) || 0;
        amzSumMap[key].orders += parseInt(row[col.orders]) || 0;
        amzSumMap[key].sales += cleanNum(row[col.sales]);
        amzSumMap[key].cost += cleanNum(row[col.cost]);
      });

      // 集計後のデータで判定
      for (let key in amzSumMap) {
        let d = amzSumMap[key];
        if (d.clicks === 0 || (histMap[d.kw] && histMap[d.kw].status === "処理済")) continue;

        let roasCheck = d.cost > 0 ? (d.sales / d.cost) * 100 : (d.orders > 0 ? 9999 : 0);
        let cpc = d.clicks > 0 ? roundNum(d.cost / d.clicks, 0) : 0;
        let cvr = d.clicks > 0 ? roundNum(d.orders / d.clicks, 4) : 0;
        let ctr = d.imp > 0 ? roundNum(d.clicks / d.imp, 4) : 0;

        let actionType = "維持", targetBid = "", note = "";

        if (roasCheck < CONFIG.AMZ_ROAS_KEEP) {
          actionType = "除外";
        } else if (d.clicks >= CONFIG.CLICK_THRESHOLD) {
          if (!d.isManual) {
            actionType = "昇格";
            targetBid = calculateSmartBid(d.sales, d.clicks, 1000, cpc);
            note = "オート➡︎マニュアル移動";
          } else {
            actionType = "調整";
            targetBid = roasCheck >= CONFIG.AMZ_ROAS_GOOD ? calculateSmartBid(d.sales, d.clicks, CONFIG.AMZ_ROAS_GOOD, cpc) : Math.max(2, Math.floor(cpc * 0.9));
            note = roasCheck >= CONFIG.AMZ_ROAS_GOOD ? "入札アップ(攻め)" : "入札ダウン(利益確保)";
          }
        }

        let outRow = [true, today, d.campName, d.adGroupName, d.kw, d.imp, d.clicks, d.orders, d.sales, d.orders>0?roundNum(d.sales/d.orders,0):0, d.cost, cpc, ctr, cvr, roundNum(roasCheck/100,2), note, targetBid, d.campId, d.adGroupId, d.isManual?"完全一致":"オート集計", d.kwId];
        if (actionType === "除外") amzNegData.push(outRow);
        else if (actionType !== "維持") amzPosData.push(outRow);
        histNew.push([new Date(), "Amazon", d.kw, roasCheck, cvr, "未処理"]);
      }
      break; 
    }
  } catch(e) { console.log("Amz Error: " + e); }

  // --- 楽天処理 (以前と同様) ---
  try {
    const rakFiles = DriveApp.getFolderById(RAK_FOLDER_ID).getFiles();
    while(rakFiles.hasNext()) {
      const file = rakFiles.next();
      if (file.getMimeType() !== MimeType.CSV && !file.getName().endsWith(".csv")) continue;
      const blob = file.getBlob();
      const rawLines = blob.getDataAsString("Shift-JIS").split("\n");
      let startIdx = -1;
      for(let i=0; i<Math.min(rawLines.length, 20); i++) if (rawLines[i].indexOf("商品管理番号") !== -1) { startIdx = i; break; }
      if (startIdx !== -1) {
        const csvData = Utilities.parseCsv(rawLines.slice(startIdx).join("\n"));
        if(sheets.rakRaw) sheets.rakRaw.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
        const headerRow = csvData[0];
        const idxId = findColIndex(headerRow, ["商品管理番号"]);
        const idxClicks = findColIndex(headerRow, ["クリック数"]);
        const idxRoas = findColIndex(headerRow, ["ROAS(合計720時間)", "ROAS"]); 
        const idxSales = findColIndex(headerRow, ["売上金額(合計720時間)", "売上金額"]);
        const idxCvr = findColIndex(headerRow, ["CVR(合計720時間)", "CVR", "転換率"]);
        if (idxId === -1 || idxClicks === -1) continue;
        csvData.slice(1).forEach(row => {
          let id = row[idxId], clicks = parseInt(row[idxClicks]) || 0;
          if (!id || clicks === 0 || (histMap[id] && histMap[id].status === "処理済")) return;
          let roasCheck = cleanNum(row[idxRoas]), sales = cleanNum(row[idxSales]);
          let res = analyzeTrendCore(id, roasCheck, CONFIG.RAK_ROAS_LIMIT, histMap[id]?.values || []);
          let outRow = [true, today, id, clicks, cleanNum(row[findColIndex(headerRow,["実績額"])]), sales, row[findColIndex(headerRow,["売上件数"])], roundNum(cleanNum(row[idxCvr])/100,4), roundNum(roasCheck/100,4), roundNum(cleanNum(row[findColIndex(headerRow,["CPC"])]),0), roundNum(cleanNum(row[findColIndex(headerRow,["CTR"])])/100,4), res.trend, res.action];
          if (roasCheck >= CONFIG.RAK_ROAS_GOOD) { let p = [...outRow]; p[0] = true; p[12] = `目標:${calculateSmartBid(sales, clicks, CONFIG.RAK_ROAS_GOOD, outRow[9])}円`; rakPosData.push(p); }
          if (res.action !== "維持") rakNegData.push(outRow);
          histNew.push([new Date(), "楽天", id, roasCheck, outRow[7], "未処理"]);
        });
        break;
      }
    }
  } catch(e) { console.log("Rak Error: " + e); }

  const amzHead = ["判定", "分析日", "キャンペーン名", "広告グループ名", "検索キーワード", "インプレッション", "クリック数", "注文", "売上", "購入単価", "支出", "CPC", "CTR", "CVR", "ROAS", "判定結果", "推奨入札額", "キャンペーンID", "広告グループID", "マッチタイプ", "キーワードID"];
  const rakHead = ["判定", "分析日", "商品管理番号", "クリック数(合計)", "実績額(合計)", "売上金額(合計720時間)", "売上件数(合計720時間)", "CVR(合計720時間)(%)", "ROAS(合計720時間)(%)", "CPC実績(合計)", "CTR(%)", "傾向", "推奨アクション"];
  writeToSheetFast(sheets.amzNeg, amzHead, amzNegData, 14); writeToSheetFast(sheets.amzPos, amzHead, amzPosData, 14);
  writeToSheetFast(sheets.rakNeg, rakHead, rakNegData, 8); writeToSheetFast(sheets.rakPos, rakHead, rakPosData, 8);
  applyAllFormats(sheets, sHist);
  if (histNew.length > 0) updateHistorySmart(sHist, histNew);
  if (typeof skipUi === "undefined" || !skipUi) Browser.msgBox("分析完了。");
  else Logger.log("分析完了。");
}

/**
 * 1-2. 分析＆バルク生成を連続実行（メニュー統合）
 * @param {boolean} skipUi - true のときダイアログを出さない（トリガー実行用）
 */
function runOptimizationAndBulk(skipUi) {
  runOptimization(skipUi);
  generateBulkFiles(skipUi);
}

/**
 * 隔週金曜 3:00 トリガー用。第2・第4金曜のときだけ「分析＆バルク生成」を実行する。
 * トリガー設定：毎日 午前3時～4時 に実行。
 */
function runOptimizationAndBulkOnSchedule() {
  const now = new Date();
  const jst = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm");
  const dayOfWeek = now.getDay();
  const dayOfMonth = now.getDate();
  const weekOfMonth = Math.ceil(dayOfMonth / 7);
  const isFriday = (dayOfWeek === 5);
  const is2ndOr4thFriday = isFriday && (weekOfMonth === 2 || weekOfMonth === 4);
  if (!is2ndOr4thFriday) {
    Logger.log("[広告運用] " + jst + " 実行スキップ（隔週金曜以外）");
    return;
  }
  Logger.log("[広告運用] " + jst + " 隔週金曜のため分析＆バルク生成を実行");
  try {
    runOptimizationAndBulk(true);
    Logger.log("[広告運用] 分析＆バルク生成 完了");
  } catch (e) {
    Logger.log("[広告運用] エラー: " + e.message);
  }
}

/**
 * 2. バルク生成 (ID自動補完・スペース無視)
 */
function generateBulkFiles(skipUi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeStamp = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm");
  const amzSubFolder = getOrCreateSubFolder(AMZ_FOLDER_ID, OUTPUT_FOLDER_NAME);
  const rakSubFolder = getOrCreateSubFolder(RAK_FOLDER_ID, OUTPUT_FOLDER_NAME);
  
  let amzRows = [["プロダクト","エンティティ","操作","キャンペーンID","キャンペーン名","広告グループID","広告グループ名","ステータス","キーワードテキスト","マッチタイプ","入札額","キーワードID"]];
  let rakRows = [["コントロールカラム","商品管理番号","商品CPC"]];
  let processedKeys = [];

  let campaignMap = {}, campaignIdToAdGroupId = {}, adGroupMap = {};
  const idSheet = ss.getSheetByName("ID参照用");
  if (idSheet) {
    const idData = idSheet.getDataRange().getValues();
    idData.slice(1).forEach(row => {
      let cName = row[9], cId = row[3], agId = row[4], agName = row[10], entity = row[1];
      let normCName = normalizeKey(cName);
      if (normCName && cId) campaignMap[normCName] = cId;
      if (cId && agId) {
        if (agName) adGroupMap[normalizeKey(cName + "_" + agName)] = agId;
        if (entity.indexOf("広告グループ") !== -1 || entity.indexOf("Ad Group") !== -1) campaignIdToAdGroupId[cId] = agId;
      }
    });
  }

  const amzPosSheet = ss.getSheetByName("入札調整候補（Amazon）");
  if (amzPosSheet) {
    amzPosSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] === true) {
        processedKeys.push(r[4]);
        if (r[15].indexOf("オート➡︎マニュアル") !== -1) {
          let tCName = "Manual_" + r[2], tAgName = "Manual_" + r[3];
          let fCId = campaignMap[normalizeKey(tCName)] || "";
          let fAgId = adGroupMap[normalizeKey(tCName + "_" + tAgName)] || campaignIdToAdGroupId[fCId] || "";
          amzRows.push(["スポンサープロダクト広告", "キーワード", "作成", fCId, tCName, fAgId, tAgName, "有効", r[4], "完全一致", r[16], ""]);
          amzRows.push(["スポンサープロダクト広告", "除外キーワード", "作成", r[17], r[2], r[18], r[3], "有効", r[4], "除外キーワードの完全一致", "", ""]);
        } else {
          amzRows.push(["スポンサープロダクト広告", "キーワード", "更新", r[17], r[2], r[18], r[3], "有効", r[4], "完全一致", r[16], r[20]]);
        }
      }
    });
  }

  const amzNegSheet = ss.getSheetByName("除外候補（Amazon）");
  if (amzNegSheet) {
    amzNegSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] === true) {
        processedKeys.push(r[4]);
        amzRows.push(["スポンサープロダクト広告", "除外キーワード", "作成", r[17], r[2], r[18], r[3], "有効", r[4], "除外キーワードの完全一致", "", ""]);
      }
    });
  }

  // 楽天処理
  ["除外候補（楽天）", "入札強化候補（楽天）"].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] === true) {
        let isNeg = name.includes("除外"), bid = r[12] ? r[12].match(/目標:(\d+)円/) : null;
        rakRows.push([isNeg ? "d" : "u", r[2], isNeg ? "" : (bid ? bid[1] : CONFIG.RAK_PROD_MIN_BID)]);
      }
    });
  });

  if (amzRows.length > 1) {
    const tempSS = SpreadsheetApp.create("temp_amz");
    const tempSheet = tempSS.getSheets()[0];
    tempSheet.getRange(1, 1, amzRows.length, amzRows[0].length).setValues(amzRows);
    [4, 6, 12].forEach(c => tempSheet.getRange(2, c, amzRows.length-1, 1).setNumberFormat("0"));
    SpreadsheetApp.flush();
    const blob = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx", {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}}).getBlob().setName(`AmazonBulk_${timeStamp}.xlsx`);
    amzSubFolder.createFile(blob); DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  }
  if (rakRows.length > 1) {
    const csv = rakRows.map(row => row.join(",")).join("\r\n");
    rakSubFolder.createFile(Utilities.newBlob("").setDataFromString(csv, "Shift-JIS").setName(`RakutenBulk_${timeStamp}.csv`));
  }
  if (processedKeys.length > 0) {
    const hSheet = ss.getSheetByName("履歴データ");
    const hData = hSheet.getDataRange().getValues();
    hData.forEach(r => { if (processedKeys.indexOf(r[2]) !== -1) r[5] = "処理済"; });
    hSheet.getDataRange().setValues(hData);
  }
  if (typeof skipUi === "undefined" || !skipUi) Browser.msgBox("保存完了。");
  else Logger.log("保存完了。");
}

/**
 * 🔧 操作マニュアル作成
 */
function createManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = getSheetByNameSafe(ss, "📌運用マニュアル").clear();
  sheet.setColumnWidth(2, 180); sheet.setColumnWidth(3, 500);
  const setH = (r, t, c) => { sheet.getRange(r, 2, 1, 2).merge().setValue(t).setBackground(c).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center"); };
  let r = 2;
  sheet.getRange(r, 2, 1, 2).merge().setValue("🚀 広告運用システム 完全マニュアル").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  r += 3;
  setH(r, "【重要】名前のルール", "#EA4335"); r++;
  sheet.getRange(r, 2, 2, 2).setValues([["名前の変更", "オート名の「日付」を消してスッキリさせる。例：「GFC 五色あられ」"], ["ペアの作成", "マニュアル名は「Manual_」＋オート名。例：「Manual_GFC 五色あられ」"]]).setWrap(true);
  r += 4;
  setH(r, "STEP 1：データの準備", "#FF9900"); r++;
  sheet.getRange(r, 2, 2, 2).setValues([["1. 分析データ", "「SP検索ワードレポート」をAmazonフォルダへ保存。"], ["2. ID用データ", "「SPキャンペーンレポート」を全コピーして「ID参照用」シートへ。"]]).setWrap(true);
  r += 4;
  setH(r, "STEP 2：実行と反映", "#34A853"); r++;
  sheet.getRange(r, 2, 2, 2).setValues([["1. 分析・生成", "メニューから「分析」→「バルク生成」の順に実行。"], ["2. アップロード", "生成されたExcelをAmazonの「一括操作」画面からアップ。"]]).setWrap(true);
  Browser.msgBox("マニュアルを更新しました。");
}

// --- 共通ヘルパー ---
function normalizeKey(s) { return s ? s.toString().replace(/\s+/g, '').replace(/　/g, '') : ""; }
function checkBulkFileErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("チェック用");
  if (!sheet) { Browser.msgBox("エラー：「チェック用」シートがありません。"); return; }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const h = data[0];
  const colMap = {}; h.forEach((name, i) => colMap[name.trim()] = i);
  let errorMsg = ""; let errCount = 0;
  const entityCol = colMap["エンティティ"] ?? colMap["Entity"], matchCol = colMap["マッチタイプ"] ?? colMap["Match Type"], bidCol = colMap["入札額"] ?? colMap["Bid"], opCol = colMap["操作"] ?? colMap["Operation"], idCol = colMap["キーワードID"] ?? colMap["Keyword Id"];
  if (entityCol === undefined || matchCol === undefined) { Browser.msgBox("エラー：必須列が見つかりません。"); return; }
  for (let i = 1; i < data.length; i++) {
    let row = data[i], entity = row[entityCol], matchType = row[matchCol], op = row[opCol], id = row[idCol];
    if ((entity === "除外キーワード" || entity === "Negative Keyword") && matchType.indexOf("除外") === -1) { errCount++; errorMsg += `[行${i+1}] 矛盾：除外なのにマッチタイプが『${matchType}』\n`; }
    if (op === "更新" && !id) { errCount++; errorMsg += `[行${i+1}] エラー：更新操作にはIDが必須です\n`; }
  }
  if (errCount > 0) Browser.msgBox(`⚠️ ${errCount} 件のエラー：\n` + errorMsg.substring(0,500));
  else Browser.msgBox("✅ データ整合性OK");
}
function applyAllFormats(sheets, sHist) {
  const fmtPct = "0.00%", fmtNum = "#,##0";
  [sheets.amzNeg, sheets.amzPos, sheets.rakNeg, sheets.rakPos].forEach(sheet => {
    if (sheet.getLastRow() > 1) {
      let lr = sheet.getLastRow() - 1;
      sheet.getRange(2, 6, lr, 6).setNumberFormat(fmtNum);
      sheet.getRange(2, 13, lr, 3).setNumberFormat(fmtPct);
    }
  });
}
function findColIndex(header, kws) { for (let i = 0; i < header.length; i++) { let c = header[i].toString(); if (kws.some(k => c.indexOf(k) !== -1)) return i; } return -1; }
function getOrCreateSubFolder(pId, n) { const p = DriveApp.getFolderById(pId), fs = p.getFoldersByName(n); return fs.hasNext() ? fs.next() : p.createFolder(n); }
function generateRakutenKeywordAdCSV() {
  const folder = DriveApp.getFolderById(RAK_FOLDER_ID), files = folder.getFiles();
  let targetFile = null; while (files.hasNext()) { let f = files.next(); if (f.getName().indexOf("act_") !== -1) { targetFile = f; break; } }
  if (!targetFile) { Browser.msgBox("エラー：act_CSVが見つかりません。"); return; }
  const csvData = Utilities.parseCsv(targetFile.getBlob().getDataAsString("Shift-JIS"));
  const header = csvData[2], col = {}; header.forEach((name, i) => { col[name.replace(/"/g, "")] = i; });
  let outputRows = [["コントロールカラム", "商品管理番号", "キーワード", "入札単価", "マッチタイプ"]];
  let count = 0;
  csvData.slice(3).forEach(row => {
    let productName = row[col["商品名"]], kw = row[col["検索キーワード"]], cvr = parseFloat(row[col["転換率"]]) || 0, itemCode = PRODUCT_MAP[productName];
    if (itemCode && kw && cvr >= CONFIG.TARGET_CVR) { outputRows.push(["u", itemCode, kw, CONFIG.RAK_KW_MIN_BID, 3]); count++; }
  });
  if (count > 0) {
    const csv = outputRows.map(r => r.join(",")).join("\r\n");
    folder.createFile(Utilities.newBlob("").setDataFromString(csv, "Shift-JIS").setName(`楽天キーワード広告入札_${Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm")}.csv`));
    Browser.msgBox(count + "件抽出しました。");
  } else { Browser.msgBox("条件に合うキーワードなし。"); }
}
function cleanupOldFiles(folder, days) {
  const files = folder.getFiles(); const threshold = new Date(); threshold.setDate(threshold.getDate() - days);
  while (files.hasNext()) { const file = files.next(); if (file.getDateCreated() < threshold) file.setTrashed(true); }
}
function runManualCleanup() { 
  const amzSub = getOrCreateSubFolder(AMZ_FOLDER_ID, OUTPUT_FOLDER_NAME), rakSub = getOrCreateSubFolder(RAK_FOLDER_ID, OUTPUT_FOLDER_NAME);
  cleanupOldFiles(amzSub, CONFIG.CLEANUP_DAYS); cleanupOldFiles(rakSub, CONFIG.CLEANUP_DAYS); Browser.msgBox("掃除完了。"); 
}
function getHistoryMap(sh) { const data = sh.getDataRange().getValues(); let m = {}; if (data.length < 2) return m; data.slice(1).forEach(r => { let hKey = r[2]; if (!m[hKey]) m[hKey] = { values: [], status: "未処理" }; m[hKey].values.push(r[3]); if (r[5] === "処理済") m[hKey].status = "処理済"; }); return m; }
function updateHistorySmart(sh, n) { const d = sh.getDataRange().getValues(), t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd"); let a = []; n.forEach(nr => { let e = false; for (let i = Math.max(1, d.length - 1000); i < d.length; i++) if (Utilities.formatDate(d[i][0], "JST", "yyyy/MM/dd") === t && d[i][2] === nr[2]) { sh.getRange(i + 1, 4, 1, 2).setValues([[nr[3], nr[4]]]); e = true; break; } if (!e) a.push(nr); }); if (a.length > 0) sh.getRange(sh.getLastRow() + 1, 1, a.length, 6).setValues(a); }
function writeToSheetFast(sh, h, d, i) { sh.clear().appendRow(h); if (d.length > 0) { d.sort((a, b) => b[i] - a[i]); sh.getRange(2, 1, d.length, h.length).setValues(d); sh.getRange(2, 1, d.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()); } }
function cleanNum(v) { return v ? parseFloat(v.toString().replace(/[￥, %]/g, "")) || 0 : 0; }
function roundNum(val, digits) { return parseFloat(val.toFixed(digits)); }
function calculateSmartBid(s, c, t, cr) { return Math.round(Math.min(Math.max((s/c)/(t/100), cr*0.7), cr*1.3)); }
function analyzeTrendCore(id, r, l, h) { let a = "維持", t = "ー"; if (r < l) { if (h.length > 0 && r > h[h.length-1]) t = "↗︎ 向上中"; else a = "除外"; } return { action: a, trend: t }; }
function getSheetByNameSafe(ss, n) { let s = ss.getSheetByName(n); if (!s) s = ss.insertSheet(n); return s; }