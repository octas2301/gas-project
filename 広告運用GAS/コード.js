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

  ui.createMenu('🚀 広告運用')
    .addItem('1. 広告レポートを分析し、除外・入札案を出す', 'runOptimization')
    .addItem('1-② 分析して広告バルクまで一気に作る', 'runOptimizationAndBulk')
    .addSeparator()
    .addItem('2. 広告バルクをDriveに保存する', 'generateBulkFiles')
    .addItem('2-② 広告バルクの誤りをチェックする', 'checkBulkFileErrors')
    .addSeparator()
    .addItem('3. 未広告ASIN：一覧を更新する', 'menuUnadvertisedAsinRefreshList')
    .addItem('3-② 未広告ASIN：レ点からバルクを作る', 'menuUnadvertisedAsinGenerateBulk')
    .addItem('3-③ 未広告ASIN：掲載マニュアルを作る（画像枠つき）', 'createUnadvertisedAsinManualSheet')
    .addSeparator()
    .addItem('4. 楽天お宝キーワードの広告CSVを作る', 'generateRakutenKeywordAdCSV')
    .addSeparator()
    .addSubMenu(ui.createMenu('99.テスト用')
      .addItem('99-① 操作マニュアルシートを作る', 'createManualSheet')
      .addItem('99-② 古いDriveファイルを削除する', 'runManualCleanup'))
    .addToUi();

  ui.createMenu('⏱ Amazonタイムセール・ポイント減衰設定')
    .addItem('1. 公式タイムセールの提出xlsxを作る（Agent依頼文）', 'menuTimeSaleP1bCursorPrompt')
    .addItem('1-② 提出xlsxを作り直す（修正後・Agent依頼文）', 'menuTimeSaleP1bRebuildPrompt')
    .addSeparator()
    .addItem('2. セール開始の21日前・14日前：数量確認メールを送る（下書き送信）', 'menuTimeSaleQtyConfirmSend')
    .addSeparator()
    .addItem('3. セール開始前日／終了翌日：ポイント作業の催促メールを送る（下書き送信）', 'menuTimeSalePointsRemindSend')
    .addItem('4. セール開始時：期間中ポイント%をAmazonに載せる（Agent依頼文）', 'menuTimeSalePointsApplyCursorPrompt')
    .addItem('5. セール終了後：減衰中ポイント%をAmazonに戻す（Agent依頼文）', 'menuTimeSalePointsRestoreCursorPrompt')
    .addSeparator()
    .addItem('6. 新しいSKUに、何日かけて何%ずつ下げるかを記入する（選択行）', 'menuProposePriceRecoverySelection')
    .addItem('6-② 予定日より前に、今すぐ1段下げたいSKUに印を付ける（選択行）', 'menuRequestTaperRunSelection')
    .addSeparator()
    .addSubMenu(ui.createMenu('99.テスト用')
      .addItem('99-① 【危険】タイムセールシートを初期化する', 'bootstrapTimeSaleSheetsP0')
      .addItem('99-② 列の並びを目的グループに直す', 'menuRealignMasterColumns')
      .addItem('99-③ ヘッダ色・入力黄・計算式を付け直す', 'menuApplyMasterHeaderColors')
      .addItem('99-④ 期間・間隔・実行依頼の入力規則を直す', 'menuFixRecoveryValidations')
      .addItem('99-⑤ 計画が空の有効SKUへ、減衰の進め方を一括記入する（初回向け）', 'menuProposePriceRecoveryEligible')
      .addItem('99-⑥ 提出xlsxは作らず、タイムセール表だけ最新にする（Agent依頼文）', 'menuTimeSaleSyncCursorPrompt')
      .addItem('99-⑦ 減衰を手動で1回回す（日次失敗時・Agent依頼文）', 'menuTimeSalePriceRecoverySendCursorPrompt'))
    .addToUi();
}

/**
 * P0: マスタ名簿（有効SKU）＋施策「タイムセール」。旧実行/期間値下げは案内のみ。
 */
function bootstrapTimeSaleSheetsP0() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterHeaders = [
    'SKU', 'ASIN', '親ASIN', '商品名', '画像URL', 'marketplace', '通貨', '有効',
    '出品者価格_SC', 'タイムセール価格_SC', '販売商品数_SC', 'V30', 'Q_fba', '原価U', 'メモ'
  ];
  const saleHeaders = [
    'sale_id', 'レーン', 'SKU', 'ASIN', '親ASIN', '商品名', '画像', 'marketplace', '通貨',
    '有効', '承認済', '種別', 'スケジュール', '開始日', '終了日',
    '出品者価格_SC', 'タイムセール価格_SC', 'タイムセール価格_確定',
    '販売商品数_SC', 'V30', '販売商品数_確定', '通常価格', 'セール価格', '目標販売', '想定利益',
    '原価U', '提出対象', '状態', '更新日時', 'runId', 'メッセージ', 'メモ'
  ];

  const master = getSheetByNameSafe(ss, 'タイムセール_マスタ');
  if (master.getLastRow() === 0 || String(master.getRange(1, 1).getValue()) !== 'SKU') {
    // 案内のみ／空ならヘッダだけ用意（名簿は人が入れる／syncが施策から復元可）
    const wasGuide = String(master.getRange(1, 1).getValue()) === '案内';
    const keep = [];
    if (!wasGuide && master.getLastRow() >= 2) {
      const vals = master.getRange(2, 1, master.getLastRow(), Math.min(2, master.getLastColumn())).getValues();
      vals.forEach(function (row) {
        const a = String(row[0] || '').trim();
        const b = String(row[1] || '').trim();
        if (/^B0/i.test(a)) keep.push(['', a]);
        else if (a || b) keep.push([a, b]);
      });
    }
    master.clear();
    master.appendRow(masterHeaders);
    master.getRange(1, 1, 1, masterHeaders.length).setFontWeight('bold');
    master.setFrozenRows(1);
    keep.forEach(function (pair) {
      master.appendRow([
        pair[0], pair[1], '', '', 'JP', 'JPY', 'TRUE', '', '', '', '', ''
      ]);
    });
  }

  const sh = getSheetByNameSafe(ss, 'タイムセール');
  if (sh.getLastRow() === 0 || String(sh.getRange(1, 1).getValue()) !== 'sale_id') {
    sh.clear();
    sh.appendRow(saleHeaders);
    sh.getRange(1, 1, 1, saleHeaders.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  ['タイムセール_実行', 'タイムセール_期間値下げ'].forEach(function (name) {
    const leg = getSheetByNameSafe(ss, name);
    leg.clear();
    leg.getRange(1, 1, 2, 1).setValues([
      ['案内'],
      ['対象はタイムセール_マスタ（有効のみ）。施策はタイムセール。sync_master_and_exec.py --write']
    ]);
  });

  const analysis = getSheetByNameSafe(ss, 'タイムセール_分析');
  if (analysis.getLastRow() === 0) {
    analysis.getRange(1, 1, 3, 1).setValues([
      ['タイムセール_分析'],
      ['ソース: タイムセール／対象名簿=タイムセール_マスタ'],
      ['B=名付き公式のみバルク／A=月・カスタム相当']
    ]);
  }

  SpreadsheetApp.getUi().alert('マスタ名簿＋施策シートを用意しました。次: sync --write');
  try {
    if (typeof applyMasterHeaderGroupColors_ === 'function') {
      applyMasterHeaderGroupColors_(master);
    }
  } catch (e) { /* optional */ }
  Logger.log(JSON.stringify({ stepName: 'bootstrapTimeSaleSheetsP0', state: 'DONE' }));
}

function menuTimeSaleSyncCursorPrompt() {
  const prompt = [
    'gas-project で Amazon タイムセール同期を実行してください。',
    '',
    'ルール: 対象はタイムセール_マスタの有効SKUのみ。名付き公式=Bバルク。月/カスタム=A。並び=開始日降順→商品名昇順。',
    '',
    '1. cd tools/amazon_deals_bulk',
    '2. python sync_master_and_exec.py',
    '3. python sync_master_and_exec.py --write',
    '4. マスタ件数とタイムセールA/B件数を報告',
    '',
    '要件§2: docs/org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md'
  ].join('\n');
  showTimeSaleCursorDialog_('99-⑥ タイムセール表だけ最新にする — Agent依頼文', prompt, 'menuTimeSaleSyncCursorPrompt');
}

// menuTimeSaleP1bCursorPrompt / menuTimeSaleP1bRebuildPrompt → TimeSaleP1bMenus.js

function showTimeSaleCursorDialog_(title, prompt, stepName) {
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:sans-serif;font-size:13px;padding:12px;}' +
    'textarea{width:100%;height:280px;font-family:monospace;font-size:11px;}</style></head><body>' +
    '<h3 style="margin-top:0;">' + title + '</h3>' +
    '<p>下の文をコピーし、Cursor Agent に貼って実行してください。</p>' +
    '<textarea id="t">' + prompt.replace(/</g, '&lt;') + '</textarea>' +
    '<p><button onclick="var t=document.getElementById(\'t\');t.select();document.execCommand(\'copy\');">コピー</button></p>' +
    '</body></html>'
  ).setWidth(640).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, title);
  Logger.log(JSON.stringify({ stepName: stepName, state: 'DONE' }));
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

/**
 * 未広告ASIN → スポンサープロダクト広告（人手UL）の手順書。
 * 灰色の枠にスクショを貼れる（挿入 → 画像 → セル内に画像）。
 */
function createUnadvertisedAsinManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '📌未広告ASIN掲載マニュアル';
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  sheet.setFrozenRows(2);
  sheet.setColumnWidth(1, 56);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 520);
  sheet.setColumnWidth(4, 280);
  sheet.setColumnWidth(5, 280);

  const imgHint_ = '【ここにスクショを貼る】\n貼り方: この枠のセルを選択 → 挿入 → 画像 → セル内に画像';
  const paintImg_ = (r, rows, label) => {
    const range = sheet.getRange(r, 4, rows, 2);
    range.merge()
      .setValue(imgHint_ + '\n\n' + label)
      .setBackground('#F1F3F4')
      .setFontColor('#5F6368')
      .setFontSize(10)
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#9AA0A6', SpreadsheetApp.BorderStyle.DASHED);
    for (let i = 0; i < rows; i++) sheet.setRowHeight(r + i, 22);
    sheet.setRowHeight(r, 28);
    sheet.setRowHeight(r + rows - 1, 28);
  };
  const head_ = (r, text, color) => {
    sheet.getRange(r, 1, 1, 5).merge()
      .setValue(text)
      .setBackground(color)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('left');
  };
  const step_ = (r, no, title, body) => {
    sheet.getRange(r, 1).setValue(no).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(r, 2).setValue(title).setFontWeight('bold').setWrap(true);
    sheet.getRange(r, 3).setValue(body).setWrap(true);
    sheet.setRowHeight(r, Math.min(120, 36 + Math.floor(String(body).length / 28) * 14));
  };

  let r = 1;
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue('スポンサープロダクト広告（未広告ASIN）— 手動掲載マニュアル')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  r = 2;
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue('メニュー「🚀 広告運用」の 3 / 3-② を使う手順。灰色枠に自分の画面スクショを貼ると分かりやすい。最終更新: メニュー3-③で再作成可。')
    .setWrap(true).setFontColor('#444444');
  sheet.setRowHeight(r, 48);

  r = 4;
  head_(r, '0. 全体の流れ（どこから何をするか）', '#1A73E8');
  r = 5;
  step_(r, '流れ', '入口',
    'スプレッドシート上部メニュー → 「🚀 広告運用」\n' +
    '・3. 未広告ASIN：一覧を更新する\n' +
    '・3-② 未広告ASIN：レ点からバルクを作る\n' +
    '・3-③ このマニュアルを作り直す\n' +
    'できたxlsx → Amazon広告の「一括操作」からアップロード（Ads API承認前は人手UL）');
  paintImg_(r, 6, '題材例: メニュー「🚀 広告運用」を開いた画面');
  r += 7;

  head_(r, '1. 毎回やること（掲載の本番手順）', '#FF9900');
  r++;
  step_(r, '①', '広告中CSVを取る',
    'Amazon広告 → レポート → 「広告中の商品」（Advertised product）をCSVでDL。\n' +
    'Driveの Amazon用フォルダ内「📥広告中の商品レポート」へ保存（いちばん新しい1本が正）。');
  paintImg_(r, 5, '題材例: 広告中の商品レポートのDL画面');
  r += 6;

  step_(r, '②', '一覧を更新',
    'メニュー → 「3. 未広告ASIN：一覧を更新する」\n' +
    'シート「未広告ASIN一覧」が書き換わる。出品開始が新しい順・最大200件。');
  paintImg_(r, 5, '題材例: 一覧更新後のダイアログ（出品数／あり／なし）');
  r += 6;

  step_(r, '③', 'レ点（載せる子だけON）',
    '「未広告ASIN一覧」で、載せたい子のレ点だけON。\n' +
    '【重要】広告状態が「あり」の行は原則OFF。\n' +
    '先日UL成功でも、レポート未反映だと「なし」＋レ点ONのまま残ることがある → 二重作成になるので、成功分のレ点は外す。\n' +
    '備考に「相乗り」「名前衝突」がある行は既定OFF（必要なら人がON）。');
  paintImg_(r, 6, '題材例: 未広告ASIN一覧（レ点・広告状態・備考）');
  r += 7;

  step_(r, '④', '一括キャンペーンDLを置く（既存に足すとき必須）',
    'Amazon広告 → 左「一括操作」→「カスタマイズしたスプレッドシートを作成しダウンロード」。\n' +
    'スポンサープロダクトにチェック。インプレッション0除外はOFF推奨。\n' +
    '落としたxlsx/CSVを Drive「📥一括キャンペーンDL」へ（最新1本が正）。\n' +
    '一覧更新・3-②のとき機械が「ID参照用」へ数値IDを書き込む（手貼り不要）。\n' +
    'ファイルが無い／IDが無い親はバルク時にスキップ。');
  paintImg_(r, 6, '題材例: 一括操作のカスタムDL／📥一括キャンペーンDLフォルダ');
  r += 7;

  step_(r, '⑤', 'バルクを作る',
    'メニュー → 「3-② 未広告ASIN：レ点からバルクを作る」→ はい。\n' +
    'Drive「📦生成されたバルクファイル」に UnadvertisedAsinBulk_日時.xlsx が保存される。');
  paintImg_(r, 5, '題材例: バルク保存完了ダイアログ');
  r += 6;

  step_(r, '⑥', 'Amazonへアップロード',
    'Amazon広告 → 左メニュー「一括操作」（またはキャンペーン周りの一括操作）。\n' +
    '「キャンペーンのアップロード」→ さきほどのxlsxを選ぶ → アップロード。\n' +
    '※「すべてのツール」一覧には無いことがある。左の一括操作／管理から開く。\n' +
    'ファイル名（オプション）は空でよい。');
  paintImg_(r, 6, '題材例: 一括操作のアップロード画面');
  r += 7;

  step_(r, '⑦', '結果確認',
    '一括操作の履歴で Success / Failed を見る。FailedならエラーレポートをDLして原因を直す。\n' +
    '成功後すぐ「広告中の商品」CSVに出ないことがある → 一覧の「あり」反映は遅れる。\n' +
    '反映前に同じレ点で再ULしない。');
  paintImg_(r, 5, '題材例: 一括操作の Success／Failed 行');
  r += 6;

  head_(r, '2. 一覧の見方（短い表）', '#34A853');
  r++;
  sheet.getRange(r, 1, 1, 5).merge().setValue(
    'レ点=載せる子／掲載方針=新規ペア or 既存に追加／キャンペーン予定名=商品が分かる名前（ASIN末尾を付けない）／' +
    '広告状態=あり・なし／備考=相乗り・名前衝突・親フォールバックなど'
  ).setWrap(true);
  sheet.setRowHeight(r, 56);
  r += 2;

  head_(r, '3. よくある注意', '#EA4335');
  r++;
  sheet.getRange(r, 1, 4, 3).setValues([
    ['注意', '同じメニューを件数分割で何度も押さない', '切れても機械再開が正。人手で3を連打しない。'],
    ['注意', 'UL成功直後の再生成', 'レポート未反映の「なし＋レ点ON」をそのまま3-②すると二重作成エラーになる。'],
    ['注意', '最適化（除外・入札）', 'メニュー1・2は別レーン（従来のKW最適化）。未広告の初回掲載は3と3-②。'],
    ['注意', 'Ads API', 'API承認後はUL自動化を検討。承認前は一括操作の人手ULが正。']
  ]).setWrap(true);
  sheet.getRange(r, 1, 4, 1).setFontWeight('bold');
  r += 5;

  head_(r, '4. 画像の貼り方（共通）', '#5F6368');
  r++;
  sheet.getRange(r, 1, 1, 5).merge().setValue(
    '右側の灰色の点線枠をクリック → メニュー「挿入」→「画像」→「セル内に画像」→ スクショを選ぶ。\n' +
    '枠が小さいときは行の高さをドラッグで広げる。マニュアルを消しそうになったら、メニュー「3-③」で作り直せる。'
  ).setWrap(true);
  sheet.setRowHeight(r, 64);

  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(
    'シート「' + name + '」を作成／更新しました。\n右側の灰色枠にスクショを貼ってください。'
  );
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