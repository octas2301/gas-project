// =======================================================
// Yahoo!ショッピング自動出品システム (Yahoo.gs)
// Ver 3.0 / Phase 3-4: 画像API連携 & CSV出力 (修正版)
// =======================================================

// ■ 定数定義
const SHEET_NAME_MASTER = '▼商品マスタ(人間作業用)';
const SHEET_NAME_YAHOO_MAP = '▼設定(Yahooマッピング)';
const SHEET_NAME_LOG = '▼ログ(システム用)';

// Yahoo! API エンドポイント
const YAHOO_AUTH_URL = "https://auth.login.yahoo.co.jp/yconnect/v2/token";
const YAHOO_IMG_API_URL = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemImage";
const YAHOO_ITEM_API_URL = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/editItem";
const YAHOO_STOCK_API_URL = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/setStock";
const YAHOO_SUBMIT_ITEM_API_URL = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/submitItem"; // ストアクリエイターpro 反映（1件ずつ）
const YAHOO_DELETE_API_URL = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/deleteItem";

// Yahoo!ストアのページURL（削除確認用）
const YAHOO_STORE_BASE_URL = "https://store.shopping.yahoo.co.jp";

// 削除確認WebアプリのURL（完了メールの「削除用」リンクで使用。デプロイし直した場合はここを更新）
const YAHOO_DELETE_WEBAPP_BASE_URL = "https://script.google.com/macros/s/AKfycbz8q0bMB-BAL7L0Y4MtgrWX4iDgeeOmOViQCfGcp4JHsboG5l8mmk_zUaV0saMKPhhd/exec";

// スプレッドシートまとめて削除用：マスタで「削除CK」にレ点を付けた行が削除対象
const DELETE_CHECK_HEADER = '削除CK';

// 【修正】フォルダIDを固定値で定義（シート読み込みエラー回避のため）
const YAHOO_SUBMISSION_FOLDER_ID = "1MbNPCe1N1CBTyx5Rch1Q6JY6k8hjhWBv"; 

// ==========================================
// 0. カスタムメニュー（コード.jsのonOpenから呼び出されるため、ここでは定義しない）
// ==========================================
// Yahoo!メニューはコード.jsの onOpen() 関数に統合済み

// 商品削除：マスタの「削除CK」にレ点を付けた行を対象に、確認のうえYahoo!から削除する
function showDeleteSelectionDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
  const data = masterSheet.getDataRange().getValues();
  const headers = data[7];
  
  const deleteCheckCol = headers.indexOf(DELETE_CHECK_HEADER);
  const yahooIdCol = headers.indexOf('Yahoo出品済ID');
  
  if (deleteCheckCol === -1) {
    ui.alert(
      '❌ 「' + DELETE_CHECK_HEADER + '」列がありません',
      'マスタに「' + DELETE_CHECK_HEADER + '」列を追加し、削除したい行にレ点を付けてから、再度「🗑️ 商品を削除...」を実行してください。',
      ui.ButtonSet.OK
    );
    return;
  }
  if (yahooIdCol === -1) {
    ui.alert('❌ エラー', '「Yahoo出品済ID」列が見つかりません。', ui.ButtonSet.OK);
    return;
  }
  
  // 削除CKにレ点が付いていて、Yahoo出品済IDがある行を対象にする（削除フラグ・削除日時は問わない＝再出品前の行も削除可能）
  const itemCodes = [];
  for (let i = 8; i < data.length; i++) {
    const row = data[i];
    const isDeleteChecked = row[deleteCheckCol] === true || row[deleteCheckCol] === 1 || String(row[deleteCheckCol]).toUpperCase() === 'TRUE';
    const yahooId = row[yahooIdCol];
    if (isDeleteChecked && yahooId && String(yahooId).trim()) {
      itemCodes.push(String(yahooId).trim());
    }
  }
  
  if (itemCodes.length === 0) {
    ui.alert(
      'ℹ️ 削除対象がありません',
      '「' + DELETE_CHECK_HEADER + '」にレ点を付けた行のうち、Yahoo出品済IDが入っている行がありません。\n\n削除したい行にレ点を付けてから再度実行してください。',
      ui.ButtonSet.OK
    );
    return;
  }
  
  const confirmResult = ui.alert(
    '⚠️ Yahoo!商品削除の確認',
    '「' + DELETE_CHECK_HEADER + '」の付いた ' + itemCodes.length + ' 件をYahoo!から削除します。\n\nこの操作は取り消せません。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (confirmResult !== ui.Button.YES) {
    return;
  }
  
  const result = executeDeleteFromDialog(itemCodes);
  // 削除が終わったら削除CKのレ点を自動で外す（次回の誤削除防止）
  const codeSet = new Set(itemCodes);
  for (let i = 8; i < data.length; i++) {
    const yahooId = data[i][yahooIdCol];
    if (yahooId && codeSet.has(String(yahooId).trim())) {
      masterSheet.getRange(i + 1, deleteCheckCol + 1).setValue(false);
    }
  }
  ui.alert(result.success ? '✅ 削除完了' : '⚠️ 削除結果', result.message, ui.ButtonSet.OK);
}

// 削除選択ダイアログのHTML生成
function createDeleteSelectionHtml(items) {
  const itemsJson = JSON.stringify(items);
  return `<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; padding: 10px; }
    .header { margin-bottom: 15px; }
    .item-list { max-height: 280px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; }
    .item { padding: 8px; border-bottom: 1px solid #eee; display: flex; align-items: center; }
    .item:hover { background: #f5f5f5; }
    .item input { margin-right: 10px; width: 18px; height: 18px; }
    .item-code { font-family: monospace; font-size: 12px; color: #666; }
    .item-name { font-size: 13px; margin-left: 10px; color: #333; }
    .buttons { margin-top: 15px; text-align: right; }
    button { padding: 10px 20px; margin-left: 10px; cursor: pointer; border-radius: 4px; }
    .delete-btn { background: #e74c3c; color: white; border: none; }
    .delete-btn:hover { background: #c0392b; }
    .cancel-btn { background: #95a5a6; color: white; border: none; }
    .select-all { margin-bottom: 10px; }
    .count { color: #666; font-size: 12px; }
  </style>
</head>
<body>
  <div class="header">
    <p>削除する商品にチェックを入れてください：</p>
    <label class="select-all">
      <input type="checkbox" id="selectAll" onchange="toggleAll()"> すべて選択
    </label>
    <span class="count" id="count">(0件選択中)</span>
  </div>
  
  <div class="item-list" id="itemList"></div>
  
  <div class="buttons">
    <button class="cancel-btn" onclick="google.script.host.close()">キャンセル</button>
    <button class="delete-btn" onclick="executeDelete()">選択した商品を削除</button>
  </div>
  
  <script>
    const items = ${itemsJson};
    
    function renderItems() {
      const container = document.getElementById('itemList');
      container.innerHTML = items.map((item, idx) => 
        '<div class="item">' +
        '<input type="checkbox" id="item' + idx + '" onchange="updateCount()">' +
        '<div>' +
        '<div class="item-code">' + item.itemCode + '</div>' +
        '<div class="item-name">' + (item.name || '') + '</div>' +
        '</div></div>'
      ).join('');
    }
    
    function toggleAll() {
      const checked = document.getElementById('selectAll').checked;
      items.forEach((_, idx) => {
        document.getElementById('item' + idx).checked = checked;
      });
      updateCount();
    }
    
    function updateCount() {
      const count = items.filter((_, idx) => document.getElementById('item' + idx).checked).length;
      document.getElementById('count').textContent = '(' + count + '件選択中)';
    }
    
    function executeDelete() {
      const selected = items.filter((_, idx) => document.getElementById('item' + idx).checked);
      if (selected.length === 0) {
        alert('削除する商品を選択してください。');
        return;
      }
      if (!confirm('選択した ' + selected.length + ' 件を削除しますか？\\n\\nこの操作は取り消せません。')) {
        return;
      }
      const itemCodes = selected.map(item => item.itemCode);
      google.script.run
        .withSuccessHandler(onSuccess)
        .withFailureHandler(onError)
        .executeDeleteFromDialog(itemCodes);
    }
    
    function onSuccess(result) {
      alert(result.message);
      google.script.host.close();
    }
    
    function onError(error) {
      alert('エラーが発生しました: ' + error.message);
    }
    
    renderItems();
  </script>
</body>
</html>`;
}

// ダイアログから呼び出される削除実行関数
function executeDeleteFromDialog(itemCodes) {
  let successCount = 0;
  let failCount = 0;
  const errors = [];
  
  itemCodes.forEach(itemCode => {
    const result = deleteYahooItem(itemCode);
    if (result.success) {
      successCount++;
    } else {
      failCount++;
      errors.push(itemCode + ': ' + result.message);
    }
  });
  
  let message = '削除完了: ' + successCount + '件';
  if (failCount > 0) {
    message += ' / 失敗: ' + failCount + '件';
    if (errors.length > 0) {
      message += '\n\nエラー:\n' + errors.join('\n');
    }
  }
  
  return { success: successCount > 0, message: message };
}

// 削除ダイアログを表示（手動入力用）
function showDeleteDialog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🗑️ Yahoo!商品削除（手動入力）',
    '削除する商品コードを入力してください：\n（例: sanky-4970107110284-66s6）\n\n※チェックボックス選択がおすすめです',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const itemCode = response.getResponseText().trim();
    if (!itemCode) {
      ui.alert('❌ エラー', '商品コードが入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    // 確認ダイアログ
    const confirm = ui.alert(
      '⚠️ 削除確認',
      `以下の商品を削除しますか？\n\n商品コード: ${itemCode}\n\nこの操作は取り消せません。`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirm === ui.Button.YES) {
      const result = deleteYahooItem(itemCode);
      if (result.success) {
        ui.alert('✅ 削除完了', `${itemCode} を削除しました。\n\nマスタシートで修正後、再出品してください。`, ui.ButtonSet.OK);
      } else {
        ui.alert('❌ 削除失敗', result.message, ui.ButtonSet.OK);
      }
    }
  }
}

// 削除対象（Yahoo出品済ID列にデータがある行）を一覧表示
function listDeletableItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
  const data = masterSheet.getDataRange().getValues();
  const headers = data[7];
  
  const yahooIdCol = headers.indexOf('Yahoo出品済ID');
  const deleteFlagCol = headers.indexOf('Yahoo削除フラグ');
  const nameCol = headers.indexOf('商品名');
  
  if (yahooIdCol === -1) {
    SpreadsheetApp.getUi().alert('❌ エラー', '「Yahoo出品済ID」列が見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  let items = [];
  for (let i = 8; i < data.length; i++) {
    const yahooId = data[i][yahooIdCol];
    const deleted = deleteFlagCol !== -1 ? data[i][deleteFlagCol] : false;
    if (yahooId && !deleted) {
      const name = nameCol !== -1 ? String(data[i][nameCol]).substring(0, 30) : '';
      items.push(`• ${yahooId}\n  ${name}...`);
    }
  }
  
  if (items.length === 0) {
    SpreadsheetApp.getUi().alert('📋 削除対象一覧', '現在、削除可能な出品済み商品はありません。', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('📋 削除対象一覧', `出品済み商品 (${items.length}件):\n\n${items.join('\n\n')}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// 1. メイン実行関数
// ==========================================

/**
 * 【実行エントリーポイント】
 * Yahoo!出品データ生成 & 画像送信 & CSV保存を行う
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ssOverride] - トリガー実行時は openById で開いたスプレッドシートを渡す
 */
function runYahooExport(ssOverride) {
  console.log("=== Yahoo!出品処理 開始 ===");
  const ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    console.error("[Yahoo!出品] スプレッドシートを取得できません（トリガー実行時は ssOverride を渡してください）");
    throw new Error("スプレッドシートを取得できません");
  }

  try {
    // 1. 必要なシートと設定の読み込み
    const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
    const mapSheet = ss.getSheetByName(SHEET_NAME_YAHOO_MAP);
    if (!masterSheet || !mapSheet) throw new Error("必要なシートが見つかりません。");

    // 2. データ構造化 (SKU爆発)
    console.log("1. データを構築中...");
    const builder = new YahooDataBuilder(masterSheet, mapSheet);
    const products = builder.buildProductData();
    
    // ※テスト制限コード（slice）を削除しました
    
    if (products.length === 0) {
      console.warn("⚠️ 出品対象の商品がありません。「出品CK」を確認してください。");
      return;
    }
    console.log(`   -> ${products.length}件の商品データを生成しました。`);

    // 3. APIアクセストークン取得
    console.log("2. APIトークンを取得中...");
    const accessToken = getYahooAccessToken(mapSheet);
    const sellerId = mapSheet.getRange("B15").getValue(); // 【追加】ストアID取得
    console.log("   -> トークン取得成功");

    // 4. 画像をAPIで送信
    console.log("3. 画像をYahoo!へ送信中...");
    const uploader = new YahooImageUploader(accessToken, sellerId); // 【修正】sellerIdを渡す
    let successImgCount = 0;


    
    // 商品ごとに画像を送信
    products.forEach((product, pIdx) => {
      // API制限考慮: 商品ごとの待機時間
      if (pIdx > 0 && pIdx % 10 === 0) Utilities.sleep(2000); 

      product.images.forEach((url, i) => {
        const fileName = product.imageFiles[i]; // code.jpg
        const itemCode = product.code;          // 商品コード
        
        try {
          // 画像送信実行
          const res = uploader.upload(itemCode, fileName, url);
          if (res) successImgCount++;
        } catch (e) {
          console.warn(`   [Skip] 画像送信エラー: ${fileName}`);
          // 画像エラーでもCSV生成は続行する
        }
      });
    });
    console.log(`   -> 画像送信処理完了 (成功: ${successImgCount}枚)`);

    // 5. APIによる商品登録 & 在庫設定
    console.log("4. Yahoo! APIへデータを送信中...");
    let itemClient = new YahooApiClient(accessToken, sellerId, mapSheet);

    let successCount = 0;
    let errorCount = 0;
    const errorItems = [];

    function isAuthError(msg) {
      if (!msg) return false;
      var s = String(msg);
      return /Authentication.*incomplet|incomplet.*request|認証.*パラメータ/.test(s);
    }

    function runOneProduct(client, product) {
      var itemCode = product.code || product.originalRow['子SKU'] || '';
      var itemRes = client.updateItem(product);
      if (!itemRes) throw new Error("商品登録APIでエラーが発生しました");
      logYahooEvent("API", "editItem", itemCode, "OK");
      Utilities.sleep(200);
      var stockRes = client.setStock(product);
      if (!stockRes) throw new Error("在庫設定APIでエラーが発生しました");
      logYahooEvent("API", "setStock", itemCode, "OK");
      Utilities.sleep(1000);
      var submitRes = client.submitItem(product);
      if (!submitRes.ok) {
        var body = submitRes.body || '';
        if (/it-07004|新規ページは個別反映できません/.test(body)) {
          logYahooEvent("API", "submitItem", itemCode, "新規のためAPI反映は未対応。ストアクリエイターproで「反映」ボタンを押してください");
          console.log("   [OK] 出品完了（反映は手動）: " + product.code);
          return true;
        }
        throw new Error("ストアクリエイターpro反映APIでエラー: HTTP " + (submitRes.code || '') + " " + body.substring(0, 200));
      }
      logYahooEvent("API", "submitItem", itemCode, "OK");
      console.log("   [OK] 出品・反映完了: " + product.code);
      return true;
    }

    products.forEach(function (product, pIdx) {
      if (pIdx > 0) Utilities.sleep(500);
      var itemCode = product.code || product.originalRow['子SKU'] || '';

      try {
        runOneProduct(itemClient, product);
        updateMasterYahooId(masterSheet, product.originalRow['子SKU'], product.code);
        successCount++;
      } catch (e) {
        if (isAuthError(e.message)) {
          console.warn("   [Retry] 認証エラーのためトークン再取得して再試行: " + product.code);
          try {
            var newToken = getYahooAccessToken(mapSheet);
            itemClient = new YahooApiClient(newToken, sellerId, mapSheet);
            runOneProduct(itemClient, product);
            updateMasterYahooId(masterSheet, product.originalRow['子SKU'], product.code);
            successCount++;
            console.log("   [OK] 再試行で出品完了: " + product.code);
          } catch (retryErr) {
            console.error("   [NG] 出品失敗 (再試行も失敗) (" + product.code + "): " + retryErr.message);
            logYahooEvent("API_ERROR", product.code || itemCode, retryErr.message, "");
            errorCount++;
            errorItems.push({ code: product.code || itemCode, name: product.name || "", message: retryErr.message || "" });
          }
        } else {
          console.error("   [NG] 出品失敗 (" + product.code + "): " + e.message);
          logYahooEvent("API_ERROR", product.code || itemCode, e.message, "");
          errorCount++;
          errorItems.push({ code: product.code || itemCode, name: product.name || "", message: e.message || "" });
        }
      }
    });

    console.log(`=== 処理完了 ===`);
    console.log(`成功: ${successCount}件 / 失敗: ${errorCount}件`);
    
    // 6. 出品完了メール送信
    if (successCount > 0) {
      sendYahooCompletionEmail(products, sellerId, successCount, errorCount, errorItems);
    }
    
    logYahooEvent("SUCCESS", "BATCH", "処理完了", `成功:${successCount}, 失敗:${errorCount}`);

  } catch (e) {
    console.error("⛔ 致命的なエラーが発生しました");
    console.error(e.stack);
    logYahooEvent("FATAL", "SYSTEM", "処理中断", e.message);
  }
}

// ==========================================
// 2. クラス定義: 画像送信 (Image Uploader)
// ==========================================

class YahooImageUploader {
  // 【修正】sellerId を受け取るように変更
  constructor(token, sellerId) {
    this.token = token;
    this.sellerId = sellerId;
  }

  /**
   * 画像を1枚送信する
   */
  upload(itemCode, fileName, imageUrl) {
    if (!imageUrl || imageUrl === "undefined") return false;

    // 【修正】楽天Cabinetの相対パスを絶対URLに変換
    if (imageUrl.startsWith("/")) {
      imageUrl = "https://image.rakuten.co.jp/octas/cabinet" + imageUrl;
    }

    // 画像データを取得
    let imageBlob;
    try {


      const imgRes = UrlFetchApp.fetch(imageUrl);
      if (imgRes.getResponseCode() !== 200) throw new Error("Image fetch failed");
      imageBlob = imgRes.getBlob().setName(fileName);
    } catch (e) {
      console.warn(`   [Skip] 画像取得失敗: ${imageUrl}`);
      return false;
    }

    // multipart/form-data の構築
    const boundary = "xxxxxxxxxx";
    let requestBody = Utilities.newBlob(
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="item_code"\r\n\r\n${itemCode}\r\n` +
      `--${boundary}\r\n` +
      // 【修正】name="image_file" -> "file" に変更
      `Content-Disposition: form-data; name="file"; filename="${fileName}"\r\n` +
      `Content-Type: image/jpeg\r\n\r\n`
    ).getBytes();
    
    requestBody = requestBody.concat(imageBlob.getBytes());
    requestBody = requestBody.concat(Utilities.newBlob(`\r\n--${boundary}--`).getBytes());

    // 【修正】URLに seller_id をクエリパラメータとして付与
    const uploadUrl = `${YAHOO_IMG_API_URL}?seller_id=${this.sellerId}`;

    const options = {
      method: "post",
      contentType: `multipart/form-data; boundary=${boundary}`,
      headers: { "Authorization": `Bearer ${this.token}` },
      payload: requestBody,
      muteHttpExceptions: true
    };

    // API送信
    const response = UrlFetchApp.fetch(uploadUrl, options);
    const responseCode = response.getResponseCode();




    // 成功判定 (200 OK)
    if (responseCode === 200) {
      console.log(`   [OK] 画像送信: ${fileName}`);
      Utilities.sleep(200); // 連続送信対策
      return true;
    } else {
      const content = response.getContentText();
      // エラーでも止まらずログだけ出す
      console.warn(`   [NG] APIエラー: ${fileName} (${responseCode})`);
      return false;
    }
  }
}

// ==========================================
// 3. クラス定義: CSV生成 (CSV Generator)
// ==========================================

class YahooCSVGenerator {
  constructor(mapSheet) {
    this.mapSheet = mapSheet;
    this.headers = [];
    this.mappings = [];
    this._loadMapping();
  }

  _loadMapping() {
    const data = this.mapSheet.getDataRange().getValues();
    // 1行目: Yahooフィールド名, 2行目: ファイル種別, 4行目: マスタ列名, 5行目: 固定値, 6行目: 出力フラグ
    const yahooFields = data[0]; 
    const fileTypes = data[1];   // 【追加】2行目(ファイル種別)を取得
    const masterCols = data[3];  
    const fixedVals = data[4];   
    const outputFlags = data[5]; 

    // B列以降を読み込む
    for (let i = 1; i < yahooFields.length; i++) {
      // 出力フラグがON、かつ ファイル種別が 'quantity' や 'option' ではない場合のみ追加
      const isOutput = (outputFlags[i] === true || outputFlags[i] === "TRUE");
      const isNotQuantity = (String(fileTypes[i]).toLowerCase().trim() !== 'quantity');
      const isNotOption = (String(fileTypes[i]).toLowerCase().trim() !== 'option');

      if (isOutput && isNotQuantity && isNotOption) {
        this.headers.push(yahooFields[i]);
        this.mappings.push({
          yahooField: yahooFields[i],
          masterCol: masterCols[i],
          fixedVal: fixedVals[i]
        });
      }
    }
  }

  /**
   * data_add.csv 生成
   */
  generateDataAddCsv(products) {
    let csv = [];
    csv.push(this.headers.join(',')); // Header






    // データ行
    products.forEach(p => {
      const row = this.mappings.map(map => {
        // --- 1. システム自動生成値 (最優先) ---
        if (map.yahooField === 'code') return p.code;
        if (map.yahooField === 'name') return this._escapeCsv(p.name);
        
        // グルーピングID (半角英数字・ハイフン以外を除去)
        if (map.yahooField === 'grouping-id') {
          return p.groupingId.replace(/[^a-zA-Z0-9-]/g, '');
        }

        // 画像URL (API送信済みのため空欄), その他不要項目
        if (map.yahooField === 'item-image-urls') return "";
        if (map.yahooField === 'sub-code') return ""; 
        if (map.yahooField === 'options') return ""; 

        // --- 2. マッピングシートの設定（固定値・マスタ参照）を優先 ---
        
        // A. 固定値・特殊タグの処理
        if (map.fixedVal !== "") {
          const fixVal = String(map.fixedVal); // 数値を文字に変換
          
          // [[固定値:XX]] の処理
          const matchFixed = fixVal.match(/\[\[固定値:(.*?)\]\]/);
          if (matchFixed) return matchFixed[1];

          // [[JOIN_BULLETS]] の処理
          if (fixVal === "[[JOIN_BULLETS]]") {
            let bullets = [];
            for (let i = 1; i <= 5; i++) {
              const colName = `商品説明の箇条書き${['①','②','③','④','⑤'][i-1]}`;
              const val = p.originalRow[colName];
              if (val) bullets.push(val);
            }
            return this._escapeCsv(bullets.join('<br>'));
          }
          // [[JOIN_BULLETS_DIAMOND]] の処理（◆付き箇条書き）
          if (fixVal === "[[JOIN_BULLETS_DIAMOND]]") {
            let bullets = [];
            for (let i = 1; i <= 5; i++) {
              const colName = `商品説明の箇条書き${['①','②','③','④','⑤'][i-1]}`;
              const val = p.originalRow[colName];
              if (val) bullets.push('◆ ' + val);
            }
            return this._escapeCsv(bullets.join('<br>'));
          }

          // [[HTML_CLEAN]] の処理
          if (fixVal === "[[HTML_CLEAN]]" && map.masterCol) {
            let val = p.originalRow[map.masterCol] || "";
            val = val.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');
            return this._escapeCsv(val);
          }
          
          if (fixVal === "[[AUTO]]") return ""; 
          return fixVal; // そのまま固定値を返す
        }

        // B. マスタ参照の処理
        if (map.masterCol) {
          const val = p.originalRow[map.masterCol];
          return this._escapeCsv(val !== undefined ? val : "");
        }

        // --- 3. 自動補完（バックアップ処理） ---
        // マッピングシートに設定がない場合のみ発動する、バリエーションの自動入力
        if (map.yahooField === 'variation1-name' && p.variationValue) {
           return this._escapeCsv(p.variationValue); // 例: 3袋
        }
        if (map.yahooField === 'variation1-free-title') {
           return "セット内容"; // デフォルト値
        }
        
        return "";
      });
      csv.push(row.join(','));
    });





    return csv.join('\r\n');
  }

  /**
   * quantity_add.csv 生成
   */
  generateQuantityAddCsv(products) {
    const headers = ["code", "sub-code", "quantity", "mode"];
    let csv = [headers.join(',')];

    products.forEach(p => {
      const row = [p.code, "", p.quantity, ""]; 
      csv.push(row.join(','));
    });

    return csv.join('\r\n');
  }

  _escapeCsv(text) {
    if (!text) return "";
    text = String(text);
    if (text.includes(',') || text.includes('"') || text.includes('\n')) {
      return '"' + text.replace(/"/g, '""') + '"';
    }
    return text;
  }
}

// ==========================================
// 4. データ構築クラス (Builder)
// ==========================================

class YahooDataBuilder {
  constructor(masterSheet, mapSheet) {
    this.masterSheet = masterSheet;
    this.mapSheet = mapSheet;
    this.masterData = null;
    this.headerMap = {};
  }

  buildProductData() {
    this._loadMasterData();
    const groups = this._groupMasterData();
    const yahooProducts = [];

    for (const [pSku, group] of Object.entries(groups)) {
      group.children.forEach(childRow => {
        const product = this._createSingleProduct(pSku, group.parent, childRow, group.siblings);
        if (product) yahooProducts.push(product);
      });
    }
    return yahooProducts;
  }

  _createSingleProduct(parentSku, parentRow, childRow, allChildren) {
    // ID生成: 子SKUをそのまま使用（Yahoo!グルーピング仕様: 子SKU = 商品コード）
    const childSku = String(childRow[this.headerMap['子SKU']]).trim();
    const yahooCode = childSku; // 子SKUをそのまま商品コードとして使用

    // 商品名: 親商品名 + 【バリエーション値】 (75文字制限・単語区切りカット)
    let baseName = String(parentRow[this.headerMap['商品名']]).trim();
    const varValue = this.headerMap['バリエーション値'] ? String(childRow[this.headerMap['バリエーション値']]).trim() : "";
    
    let suffixName = varValue ? ` 【${varValue}】` : "";
    const maxLen = 75 - suffixName.length;

    if (baseName.length > maxLen) {
      let trimmed = baseName.substring(0, maxLen);
      // 末尾がスペースでなければ、最後のスペースまでカット（単語途中での切断防止）
      const lastSpace = trimmed.lastIndexOf(" ");
      if (lastSpace > 0) {
        trimmed = trimmed.substring(0, lastSpace);
      }
      baseName = trimmed;
    }
    const yahooName = baseName + suffixName;

    // --- 画像処理 ---
    // 1. 自分の楽天メイン画像
    const myImage = this._getColValue(childRow, '楽天メイン画像1');

    // 2. 楽天サブ画像 (1-8まで取得。子になければ親を確認)
    const subImages = [];
    for (let i = 1; i <= 8; i++) {
      let url = this._getColValue(childRow, `楽天サブ画像${i}`);
      if (!url) url = this._getColValue(parentRow, `楽天サブ画像${i}`);
      if (url) subImages.push(url);
    }

    // 3. 兄弟SKUのメイン画像を収集（自分以外）してサブ画像の最後に追加
    const siblingImages = [];
    const myChildSku = String(childRow[this.headerMap['子SKU']]).trim();
    allChildren.forEach(sibling => {
      const sibSku = String(sibling[this.headerMap['子SKU']]).trim();
      if (sibSku !== myChildSku) {
        const sibImg = this._getColValue(sibling, '楽天メイン画像1');
        if (sibImg) siblingImages.push(sibImg);
      }
    });

    // 結合: 自分のメイン → サブ画像 → 兄弟のメイン画像
    let imageUrls = [myImage, ...subImages, ...siblingImages].filter(url => url && url !== "" && url !== "undefined");
    imageUrls = [...new Set(imageUrls)]; // 重複除去





    // ファイル名生成 (code.jpg, code_1.jpg...)
    const imageFiles = imageUrls.map((url, idx) => {
      const ext = ".jpg"; 
      if (idx === 0) return `${yahooCode}${ext}`;
      return `${yahooCode}_${idx}${ext}`; 
    });

    // 親データと子データをマージ（子が空なら親の値を使う）
    const mergedData = this._mergeParentChild(parentRow, childRow);

    // 【修正】特定項目は強制的に親の値を適用（子行の入力を無視）
    const forceParentCols = ['商品名', 'キャッチコピー', 'Yahoo!キャッチコピー\nAIから取得', '商品説明の箇条書き①', '商品説明の箇条書き②', '商品説明の箇条書き③', '商品説明の箇条書き④', '商品説明の箇条書き⑤'];
    forceParentCols.forEach(col => {
      const idx = this.headerMap[col];
      if (idx !== undefined) mergedData[col] = String(parentRow[idx]).trim();
    });

    // オブジェクト生成（grouping_id はAPI仕様で半角英数字・ハイフンのみ。同一親＝同一IDに統一）
    const raw = String(parentSku).trim();
    const half = raw.replace(/[０-９Ａ-Ｚａ-ｚ]/g, function (ch) {
      const c = ch.charCodeAt(0);
      if (c >= 0xFF10 && c <= 0xFF19) return String.fromCharCode(c - 0xFEE0); // ０-９→0-9
      if (c >= 0xFF21 && c <= 0xFF3A) return String.fromCharCode(c - 0xFEE0); // Ａ-Ｚ
      if (c >= 0xFF41 && c <= 0xFF5A) return String.fromCharCode(c - 0xFEE0); // ａ-ｚ
      return ch;
    });
    const safeGroupingId = half.replace(/[^a-zA-Z0-9-]/g, '') || raw.replace(/[^a-zA-Z0-9-]/g, '');
    return {
      code: yahooCode,
      name: yahooName,
      groupingId: safeGroupingId || parentSku,
      quantity: childRow[this.headerMap['在庫数']] || 0,
      variationValue: varValue, // 【追加】バリエーション名（例: 3袋）を保持
      originalRow: mergedData, // 【修正】マージ済みデータを使用
      parentRow: parentRow,
      images: imageUrls,
      imageFiles: imageFiles
    };
  }

  /**
   * 【追加】行データ(配列)を、ヘッダー名をキーとする連想配列に変換
   */
  _mapRowToObject(row) {
    const obj = {};
    for (const [key, idx] of Object.entries(this.headerMap)) {
      obj[key] = String(row[idx]).trim();
    }
    return obj;
  }


  /**
   * 【追加】親行と子行をマージする（子行が空欄なら親行の値を使う）
   */
  _mergeParentChild(parentRow, childRow) {
    const merged = {};
    for (const [key, idx] of Object.entries(this.headerMap)) {
      const childVal = String(childRow[idx]).trim();
      const parentVal = String(parentRow[idx]).trim();
      // 子が空なら親を使う、そうでなければ子を使う
      merged[key] = childVal !== "" ? childVal : parentVal;
    }
    return merged;
  }



  // カラム名から安全に値を取得するヘルパー
  _getColValue(row, colName) {
    const idx = this.headerMap[colName];
    if (idx === undefined || idx === -1) return "";
    return String(row[idx]).trim();
  }

  _loadMasterData() {
    const values = this.masterSheet.getDataRange().getValues();
    let headerRowIdx = -1;
    for (let r = 0; r < 20; r++) {
      if (values[r].includes('親SKU') && values[r].includes('出品CK')) {
        headerRowIdx = r;
        break;
      }
    }
    if (headerRowIdx === -1) throw new Error("マスタのヘッダー行が見つかりません。");

    const headers = values[headerRowIdx];
    headers.forEach((h, i) => { this.headerMap[String(h).trim()] = i; });
    this.masterData = values.slice(headerRowIdx + 1);
  }

  _groupMasterData() {
    const groups = {};
    let currentParent = null;
    let currentParentSku = "";
    const colP = this.headerMap['親SKU'];
    const colC = this.headerMap['子SKU'];
    const colCK = this.headerMap['出品CK'];

    this.masterData.forEach(row => {
      const pSku = String(row[colP]).trim();
      const cSku = String(row[colC]).trim();
      const isCheck = row[colCK] === true;

      if (pSku && !cSku) {
        // 親行: グループを用意（親のレ点は出数目の制御には使わず、子のレ点のみで出品する）
        currentParentSku = pSku;
        currentParent = row;
        if (!groups[pSku]) {
          groups[pSku] = { parent: row, children: [], siblings: [] };
        } else {
          groups[pSku].parent = row;
        }
      } else if (pSku && cSku && currentParent && pSku === currentParentSku) {
        // 子行: レ点が付いている行だけ出品対象に追加
        if (isCheck) {
          groups[pSku].children.push(row);
          groups[pSku].siblings.push(row);
        }
      }
    });
    return groups;
  }
}

// ==========================================
// 5. 共通ユーティリティ (Auth/Log)
// ==========================================

function getYahooAccessToken(mapSheet) {
  const clientId = mapSheet.getRange("B12").getValue();
  const clientSecret = mapSheet.getRange("B13").getValue();
  const refreshToken = mapSheet.getRange("B14").getValue();

  if (!refreshToken) throw new Error("リフレッシュトークンがありません。");

  const payload = {
    "grant_type": "refresh_token",
    "client_id": clientId,
    "client_secret": clientSecret,
    "refresh_token": refreshToken
  };
  const options = {
    "method": "post",
    "payload": payload,
    "muteHttpExceptions": true
  };
  
  const res = UrlFetchApp.fetch(YAHOO_AUTH_URL, options);
  if (res.getResponseCode() !== 200) {
    throw new Error(`トークン取得失敗: ${res.getContentText()}`);
  }
  
  const json = JSON.parse(res.getContentText());
  return json.access_token;
}

function logYahooEvent(type, code, message, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (sheet) {
    sheet.appendRow([new Date(), type, code, message, detail]);
  }
}

function setupYahooEnvironment() {
  /* 初期設定済みのため省略 */
}






// ==========================================
// 6. クラス定義: 商品・在庫API (Item/Stock Client)
// ==========================================

class YahooApiClient {
  constructor(token, sellerId, mapSheet) {
    this.token = token;
    this.sellerId = sellerId;
    this.mapSheet = mapSheet;
    this.mapping = this._loadMapping();
  }

  // マッピング設定の読み込み（出力フラグ・ファイル区分も含む）
  _loadMapping() {
    const data = this.mapSheet.getDataRange().getValues();
    const yahooFields = data[0];   // 1行目: Yahoo項目名
    const fileTypes = data[1];     // 2行目: ファイル区分
    const masterCols = data[3];    // 4行目: マスタ参照列
    const fixedVals = data[4];     // 5行目: 固定値/設定
    const outputFlags = data[5];   // 6行目: 出力フラグ
    if (!this._mappingLoadLogged) {
      this._mappingLoadLogged = true;
      console.log("   [DEBUG] マッピング読込: 1行目列数=" + (yahooFields ? yahooFields.length : 0) + "（getDataRangeの範囲内）");
      const leadLike = (yahooFields || []).filter(function(c) {
        const s = String(c).trim();
        return s && /納期|発送|lead/i.test(s);
      });
      if (leadLike.length) console.log("   [DEBUG] 1行目で納期/発送/leadを含む列名: " + leadLike.join(" | "));
      else console.log("   [DEBUG] 1行目に「納期」「発送」「lead」を含む列はありません（発送日列の追加・列名を確認してください）");
    }
    
    const map = {};
    const isLeadTimeKey = function(f) {
      return /lead|納期|発送|instock|outstock/i.test(String(f).trim()) || f === "在庫あり納期" || f === "在庫なし納期";
    };
    for (let i = 1; i < yahooFields.length; i++) {
      let field = String(yahooFields[i]).trim();
      if (field) {
        // 発送日情報: 列名が切れている場合の互換（lead-time-outstoc → lead-time-outstock）
        if (field === "lead-time-outstoc") field = "lead-time-outstock";
        var ft = fileTypes[i] != null ? String(fileTypes[i]).trim().toLowerCase() : "";
        var of = outputFlags[i];
        var outFlag = of === true || (of != null && String(of).trim().toUpperCase() === "TRUE");
        const config = {
          masterCol: masterCols[i],
          fixedVal: fixedVals[i],
          fileType: ft,
          outputFlag: outFlag
        };
        // 発送日関連: 同名列が複数ある場合「data＋出力ON」を優先（後からoption/falseで上書きされないようにする）
        if (isLeadTimeKey(field)) {
          const cur = map[field];
          const curGood = cur && cur.fileType === "data" && cur.outputFlag;
          const newGood = ft === "data" && outFlag;
          if (curGood && !newGood) continue;  // 既にdata+trueがあるなら上書きしない
        }
        map[field] = config;
      }
    }
    return map;
  }

  // Yahoo項目名 → APIパラメータ名の変換（ハイフン→アンダースコア、日本語列名の互換）
  _toApiParamName(yahooField) {
    const f = String(yahooField).trim();
    if (f === "lead-time-outstoc") return "lead_time_outstock";
    if (f === "lead-time-instock") return "lead_time_instock";
    if (f === "lead-time-outstock") return "lead_time_outstock";
    // マッピング1行目が日本語の場合の互換（項目名が違うこともあり得る）
    if (/在庫あり納期|発送日.*在庫あり|lead_time_instock/i.test(f)) return "lead_time_instock";
    if (/在庫なし納期|発送日.*在庫なし|lead_time_outstock/i.test(f)) return "lead_time_outstock";
    return f.replace(/-/g, "_");
  }

  // 商品登録 (editItem) - 出力フラグがTRUEの項目のみ送信
  updateItem(product) {
    this._logLeadTimeMappingKeys();
    const payload = {};
    
    // 出力フラグがTRUE かつ ファイル区分が'data'の項目を動的に追加
    for (const [field, config] of Object.entries(this.mapping)) {
      if (!config.outputFlag) continue;
      if (config.fileType !== 'data') continue;
      
      const apiParam = this._toApiParamName(field);
      let value = this._getValue(field, product);
      
      // 特殊処理が必要な項目
      if (field === 'code') {
        // item_code は子SKU（sub-code）を使用
        value = this._getValue('sub-code', product) || product.originalRow['子SKU'] || product.code;
      } else if (field === 'sub-code') {
        // sub-codeは別途処理されるのでスキップ
        continue;
      } else if (field === 'taxable') {
        // taxableは数値を文字列に変換
        value = String(this._coerceTaxable(value));
      } else if (field === 'name' && !value) {
        // nameが空ならbuilder生成値を使用
        value = product.name;
      } else if (field === 'grouping-id') {
        // grouping_idはマッピングから取得、なければbuilder生成値。APIは半角英数字・ハイフンのみ有効
        value = value || product.groupingId;
        value = String(value).replace(/[^a-zA-Z0-9-]/g, '');
      }
      
      // 空でない値のみ追加
      if (value !== "" && value !== null && value !== undefined) {
        payload[apiParam] = value;
      } else if (apiParam === "lead_time_instock" || apiParam === "lead_time_outstock") {
        console.log("   [DEBUG] 発送日: \"" + field + "\" => API=" + apiParam + " は値が空のため送信しません（_getValue結果: \"" + value + "\"）");
      }
    }
    
    // item_codeは必須なので確実に設定（子SKUを使用）
    if (!payload.item_code) {
      payload.item_code = this._getValue('sub-code', product) || product.originalRow['子SKU'] || product.code;
    }
    
    // nameは必須なので確実に設定
    if (!payload.name) {
      const nameFromMapping = this._getValue('name', product);
      payload.name = nameFromMapping || product.name || product.originalRow['商品名'] || "";
      console.log(`   [DEBUG] name補完: マッピング="${nameFromMapping}", builder="${product.name}", 最終="${payload.name}"`);
    }
    
    // 発送日: fileType=option / outputFlag=false でスキップされていても、マッピングに列があり値があれば送信する
    const leadTimeApiNames = ["lead_time_instock", "lead_time_outstock"];
    for (const ltKey of Object.keys(this.mapping)) {
      const apiParam = this._toApiParamName(ltKey);
      if (leadTimeApiNames.indexOf(apiParam) === -1) continue;
      if (payload[apiParam] !== undefined && payload[apiParam] !== "") continue;
      const val = this._getValue(ltKey, product);
      if (val !== "" && val !== null && val !== undefined) {
        payload[apiParam] = val;
      }
    }
    
    // デバッグログ: 送信するペイロードの内容
    console.log(`   [DEBUG] 送信ペイロード (${Object.keys(payload).length}項目):`);
    for (const [key, val] of Object.entries(payload)) {
      const truncVal = String(val).length > 50 ? String(val).substring(0, 50) + "..." : val;
      console.log(`      ${key} = "${truncVal}"`);
    }
    // 発送日情報: 毎回送信有無をログ（要因切り分け用）。項目名が日本語でもAPI名で送るよう _toApiParamName で変換済み
    console.log(`   [DEBUG] 発送日情報: lead_time_instock=${payload.lead_time_instock !== undefined ? payload.lead_time_instock : "(未送信)"}, lead_time_outstock=${payload.lead_time_outstock !== undefined ? payload.lead_time_outstock : "(未送信)"}`);

    return this._post(YAHOO_ITEM_API_URL, payload);
  }

  // 発送日関連のマッピングキーを診断（初回のみログ。項目名が違う場合の切り分け用）
  _logLeadTimeMappingKeys() {
    if (this._leadTimeKeysLogged) return;
    this._leadTimeKeysLogged = true;
    var keys = Object.keys(this.mapping).filter(function (k) {
      return /lead|納期|発送|instock|outstock/i.test(k) || k === "在庫あり納期" || k === "在庫なし納期";
    });
    console.log("   [DEBUG] 発送日関連とみなしたマッピングの1行目キー: " + (keys.length ? keys.join(", ") : "(なし)"));
    if (keys.length > 0) {
      for (var i = 0; i < keys.length; i++) {
        var c = this.mapping[keys[i]];
        var apiParam = this._toApiParamName(keys[i]);
        console.log("   [DEBUG]   -> \"" + keys[i] + "\" => API: " + apiParam + ", fileType=\"" + (c.fileType || "") + "\", outputFlag=" + c.outputFlag + ", fixedVal=\"" + (c.fixedVal != null ? String(c.fixedVal).substring(0, 30) : "") + "\"");
      }
    } else {
      console.log("   [DEBUG] マッピング総列数(1行目): " + (this.mapping ? Object.keys(this.mapping).length : 0) + "。発送日用の列は1行目に「lead」「納期」「発送」等を含む名前である必要があります。");
    }
  }

  // ===== 調査用: form-urlencoded 文字列化（実際に送る本文の可視化） =====
  _encodeFormPayload(payload) {
    return Object.keys(payload)
      .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(this._toStr(payload[k]))}`)
      .join("&");
  }

  // 調査用: 送信直前の payload を整形して返す（seller_id 付与＋空削除）
  _preparePayloadForPost(payload) {
    const p = Object.assign({}, payload);
    p["seller_id"] = this.sellerId;
    Object.keys(p).forEach(key => {
      if (p[key] === "" || p[key] === null || p[key] === undefined) delete p[key];
    });
    return p;
  }

  // 在庫設定 (setStock) - 出力フラグがTRUE かつ ファイル区分が'quantity'の項目を送信
  setStock(product) {
    const payload = {};
    
    // 出力フラグがTRUE かつ ファイル区分が'quantity'の項目を動的に追加
    for (const [field, config] of Object.entries(this.mapping)) {
      if (!config.outputFlag) continue;
      if (config.fileType !== 'quantity') continue;
      
      const apiParam = this._toApiParamName(field);
      let value = this._getValue(field, product);
      
      // 特殊処理が必要な項目
      if (field === 'code') {
        // item_code は子SKU（sub-code）を使用
        value = this._getValue('sub-code', product) || product.originalRow['子SKU'] || product.code;
      } else if (field === 'sub-code') {
        // sub-codeは別途処理されるのでスキップ
        continue;
      } else if (field === 'quantity') {
        // quantityは整数を文字列に変換
        value = String(Math.floor(Number(value || product.quantity || 0)));
      } else if (field === 'allow-overdraft' || field === 'stock-close') {
        // Boolean値の変換
        value = this._coerceBoolean(value, false);
      }
      
      // 空でない値のみ追加
      if (value !== "" && value !== null && value !== undefined) {
        payload[apiParam] = value;
      }
    }
    
    // item_codeは必須なので確実に設定（子SKUを使用）
    if (!payload.item_code) {
      payload.item_code = this._getValue('sub-code', product) || product.originalRow['子SKU'] || product.code;
    }
    
    // quantityは必須なので確実に設定
    if (!payload.quantity) {
      payload.quantity = String(Math.floor(Number(product.quantity || 0)));
    }
    
    // デバッグログ: 送信するペイロードの内容
    console.log(`   [DEBUG] setStock ペイロード (${Object.keys(payload).length}項目):`);
    for (const [key, val] of Object.entries(payload)) {
      console.log(`      ${key} = "${val}"`);
    }

    return this._post(YAHOO_STOCK_API_URL, payload);
  }

  /**
   * ストアクリエイターpro 反映 (submitItem)
   * editItem で登録した商品をストアに反映する。1クエリ/秒の制限あり。
   * @param {Object} product - 商品オブジェクト（code が item_code として使用される）
   * @returns {{ ok: boolean, code?: number, body?: string }} 成功時 ok:true。失敗時 ok:false と code/body を返す（throw しない）
   */
  submitItem(product) {
    const itemCode = product.code || product.originalRow['子SKU'] || '';
    if (!itemCode) {
      console.warn('   [submitItem] item_code が空のためスキップ');
      return { ok: false, body: 'item_code empty' };
    }
    const payload = {
      seller_id: this.sellerId,
      item_code: itemCode
    };
    return this._postAndReturnBody(YAHOO_SUBMIT_ITEM_API_URL, payload);
  }
  
  // Boolean値の変換ヘルパー（"1", "true", 1, true → true / その他 → defaultVal）
  _coerceBoolean(val, defaultVal) {
    if (val === true || val === 1 || val === "1" || String(val).toLowerCase() === "true") return true;
    if (val === false || val === 0 || val === "0" || String(val).toLowerCase() === "false") return false;
    return defaultVal;
  }

  // 値解決ヘルパー (マッピング or 固定値 or マスタ値)
  _getValue(field, product) {
    const config = this.mapping[field];
    if (!config) return "";

    const fixed = this._toStr(config.fixedVal).trim();
    const masterCol = this._toStr(config.masterCol).trim();

    // 固定値/タグが設定されている場合
    if (fixed !== "") {
      // 自動系タグはここでは空扱い
      if (fixed === "[[AUTO]]" || fixed === "[[IMAGE_AUTO]]") return "";

      // 箇条書き結合（そのまま改行で連結）
      if (fixed === "[[JOIN_BULLETS]]") {
        const bullets = [];
        const marks = ['①','②','③','④','⑤'];
        for (let i = 0; i < marks.length; i++) {
          const v = product.originalRow[`商品説明の箇条書き${marks[i]}`];
          const s = this._toStr(v).trim();
          if (s) bullets.push(s);
        }
        return bullets.join('<br>');
      }
      // 箇条書き＋◆付き（競合のような装飾。caption/explanation向け）
      if (fixed === "[[JOIN_BULLETS_DIAMOND]]") {
        const bullets = [];
        const marks = ['①','②','③','④','⑤'];
        for (let i = 0; i < marks.length; i++) {
          const v = product.originalRow[`商品説明の箇条書き${marks[i]}`];
          const s = this._toStr(v).trim();
          if (s) bullets.push('◆ ' + s);
        }
        return bullets.join('<br>');
      }

      // HTML除去
      if (fixed === "[[HTML_CLEAN]]") {
        const base = masterCol ? product.originalRow[masterCol] : "";
        return this._stripHtml(this._toStr(base)).trim();
      }

      // 定期購入価格: 通常販売価格(price)×0.95 をラウンドダウン。計算元は editItem の price と同一にし表示と一致させる
      if (field === "subscription-price" && (fixed === "[[CALC_SUBSCRIPTION_PRICE]]" || fixed.indexOf("[[CALC_SUBSCRIP") === 0)) {
        const priceConfig = this.mapping["price"];
        let raw = "";
        if (priceConfig) {
          const pFixed = this._toStr(priceConfig.fixedVal).trim();
          const pCol = this._toStr(priceConfig.masterCol).trim();
          if (pFixed !== "" && pFixed !== "[[AUTO]]") {
            const parsed = this._parseFixedValueTag(pFixed);
            raw = parsed !== null ? parsed : pFixed;
          } else if (pCol) {
            raw = product.originalRow[pCol];
          }
        }
        if (raw === "" || raw === undefined) {
          raw = product.originalRow["販売価格"] || product.originalRow["商品の販売価格"] || product.originalRow["Yahoo!価格設定"] || 0;
        }
        const num = String(raw).replace(/[０-９]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); }).replace(/[^0-9.]/g, "");
        const price = Number(num) || 0;
        return String(Math.floor(price * 0.95));
      }

      // [[固定値:XX]] → XX 抽出
      const parsed = this._parseFixedValueTag(fixed);
      if (parsed !== null) return parsed;

      return fixed;
    }

    // 固定値が無ければマスタ参照
    if (masterCol) {
      const v = product.originalRow[masterCol];
      return this._toStr(v).trim();
    }
    return "";
  }

  _toStr(v) {
    if (v === null || v === undefined) return "";
    return String(v);
  }

  _stripHtml(s) {
    return this._toStr(s).replace(/<[^>]*>/g, "");
  }

  _parseFixedValueTag(s) {
    const m = this._toStr(s).match(/\[\[\s*固定値\s*:\s*(.*?)\s*\]\]/);
    if (!m) return null;
    return this._toStr(m[1]).trim();
  }

  _coerceTaxable(raw) {
    // 期待値: 0 or 1。[[固定値:1]] や「課税」なども受けて正規化する。
    const s0 = this._toStr(raw).trim();
    if (s0 === "") return 1; // デフォルトは課税 1

    // タグを剥がす
    const parsed = this._parseFixedValueTag(s0);
    const s1 = (parsed !== null ? parsed : s0).trim();

    // 全角数字→半角
    const normalized = s1.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));

    // 日本語表現
    if (/非課税/.test(normalized)) return 0;
    if (/課税/.test(normalized)) return 1;

    const n = parseInt(normalized, 10);
    if (Number.isNaN(n)) return 1;
    return n === 1 ? 1 : 0;
  }

  // API送信実行
  _post(url, payload) {
    // 【修正】seller_id をPayloadに追加 (URLではなくBodyで送る)
    payload["seller_id"] = this.sellerId;
    const endpoint = url;
    
    // 不要な空文字パラメータを削除
    Object.keys(payload).forEach(key => {
      if (payload[key] === "" || payload[key] === null || payload[key] === undefined) {
        delete payload[key];
      }
    });

    const options = {
      method: "post",
      contentType: "application/x-www-form-urlencoded", // ItemAPIはForm形式推奨
      headers: { "Authorization": `Bearer ${this.token}` },
      payload: payload,
      muteHttpExceptions: true
    };

    const res = UrlFetchApp.fetch(endpoint, options);
    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code === 200) {
      return true;
    } else {
      console.warn(`API Error (${code}): ${body}`);
      throw new Error(`API Error: ${body}`);
    }
  }

  /**
   * POST してレスポンス本文を返す（submitItem の反映結果確認用）
   * @returns {{ ok: boolean, code: number, body: string }}
   */
  _postAndReturnBody(url, payload) {
    payload["seller_id"] = this.sellerId;
    Object.keys(payload).forEach(key => {
      if (payload[key] === "" || payload[key] === null || payload[key] === undefined) delete payload[key];
    });
    const options = {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      headers: { "Authorization": `Bearer ${this.token}` },
      payload: payload,
      muteHttpExceptions: true
    };
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    const body = res.getContentText();
    return { ok: code === 200, code: code, body: body };
  }
}







// ==============================================================================================================================






// ==========================================
// ■ 調査用デバッグ関数 7 (Payload生成診断)
// ==========================================

function debugPayloadConstruction() {
  console.log("=== Payload生成診断 開始 ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('▼商品マスタ(人間作業用)');
  const mapSheet = ss.getSheetByName('▼設定(Yahooマッピング)');

  // 1. データ構築 (Builder利用)
  console.log("1. 商品データを構築中...");
  const builder = new YahooDataBuilder(masterSheet, mapSheet);
  let products = builder.buildProductData();
  
  // テスト対象を1件に絞る
  if (products.length === 0) {
    console.error("❌ 出品対象の商品が見つかりません。出品CKを入れてください。");
    return;
  }
  const targetProduct = products[0];
  console.log(`   対象商品: ${targetProduct.code}`);
  console.log(`   商品名(Builder生成): ${targetProduct.name}`);

  // 2. APIクライアント初期化
  const client = new YahooApiClient("dummy_token", "dummy_seller", mapSheet);
  
  // 3. マッピング設定の確認
  console.log("\n2. マッピング設定のロード確認");
  const mapping = client.mapping;
  const checkFields = ['path', 'name', 'price', 'product-category', 'ship-weight'];
  
  checkFields.forEach(field => {
    if (mapping[field]) {
      console.log(`   [${field}] 設定あり -> マスタ列:「${mapping[field].masterCol}」 / 固定値:「${mapping[field].fixedVal}」`);
    } else {
      console.error(`   ❌ [${field}] のマッピング設定が読み込めていません！シートの1行目を確認してください。`);
    }
  });

  // 4. 値の抽出テスト (_getValueのシミュレーション)
  console.log("\n3. 値の抽出テスト (Payloadに入る予定の値)");
  
  // エラーが出ていた項目を重点チェック
  const testParams = [
    { key: 'path', label: 'パス(path)' },
    { key: 'name', label: '商品名(name)' },
    { key: 'price', label: '価格(price)' },
    { key: 'product-category', label: 'カテゴリ(product-category)' },
    { key: 'ship-weight', label: '重量(ship-weight)' },
    { key: 'taxable', label: '課税(taxable)' }
  ];

  testParams.forEach(param => {
    let val = "";
    const config = mapping[param.key];

    if (config) {
      if (config.fixedVal && config.fixedVal !== "") {
        val = `[固定値] ${config.fixedVal}`;
      } else if (config.masterCol) {
        val = targetProduct.originalRow[config.masterCol];
        if (val === undefined) val = "[Undefined] (列名不一致の可能性)";
      } else {
        val = "[設定なし]";
      }
    } else {
      val = "[マッピングなし]";
    }

    console.log(`   ${param.label}: 「${val}」`);

    if (!val || val === "" || val === "[Undefined] (列名不一致の可能性)") {
      console.warn(`   ⚠️ この項目が空のため、APIエラーになっています。`);
    }
  });

  // 4. taxable の最終評価テスト（_getValue と _coerceTaxable を直接確認）
  console.log("\n4. taxable の最終評価テスト");
  const rawTaxable = client._getValue('taxable', targetProduct);
  const coercedTaxable = client._coerceTaxable(rawTaxable);
  console.log(`   [_getValue] taxable raw = 「${rawTaxable}」 (type: ${typeof rawTaxable})`);
  console.log(`   [_coerceTaxable] taxable coerced = ${coercedTaxable} (type: ${typeof coercedTaxable})`);

  // 5. editItem へ「実際に送る form-urlencoded 本文」を可視化
  console.log("\n5. editItem 送信本文（form-urlencoded）確認");
  const sellerId = mapSheet.getRange("B15").getValue();
  const accessToken = getYahooAccessToken(mapSheet);
  const realClient = new YahooApiClient(accessToken, sellerId, mapSheet);

  // updateItem と同様に payload を組み立て
  const payload = {
    "item_code": targetProduct.code,
    "path": realClient._getValue('path', targetProduct),
    "name": targetProduct.name,
    "price": targetProduct.originalRow['販売価格'] || 0,
    "product_category": realClient._getValue('product-category', targetProduct),
    "jan": realClient._getValue('jan', targetProduct),
    "headline": realClient._getValue('headline', targetProduct),
    "caption": realClient._getValue('caption', targetProduct),
    "explanation": realClient._getValue('explanation', targetProduct),
    "relevant_links": realClient._getValue('relevant-links', targetProduct),
    "ship_weight": realClient._getValue('ship-weight', targetProduct),
    "taxable": String(realClient._coerceTaxable(realClient._getValue('taxable', targetProduct))),
    "release_date": realClient._getValue('release-date', targetProduct),
    "temporary_point_term": realClient._getValue('temporary-point-term', targetProduct),
    "point_code": realClient._getValue('point-code', targetProduct),
    "sale_period_start": realClient._getValue('sale-period-start', targetProduct),
    "sale_period_end": realClient._getValue('sale-period-end', targetProduct),
    "delivery": realClient._getValue('delivery', targetProduct),
    "astk_code": realClient._getValue('astk-code', targetProduct),
    "condition": realClient._getValue('condition', targetProduct),
    "taojapan": realClient._getValue('taojapan', targetProduct),
    "lead_time_instock": realClient._getValue('lead-time-instock', targetProduct),
    "lead_time_outstock": realClient._getValue('lead-time-outstock', targetProduct),
    "brand_code": realClient._getValue('brand-code', targetProduct),
    "grouping_id": targetProduct.groupingId
  };
  const prepared = realClient._preparePayloadForPost(payload);
  const body = realClient._encodeFormPayload(prepared);

  console.log(`   taxable(payload) = ${prepared.taxable} (type: ${typeof prepared.taxable})`);
  console.log(`   taxable(form) = ${body.match(/(?:^|&)taxable=([^&]*)/)?.[1] || "(not found)"}`);
  console.log(`   form body (先頭500文字) = ${body.slice(0, 500)}`);

  console.log("\n=== 診断終了 ===");
}

// ==========================================
// ■ 調査用デバッグ関数 8 (実際のAPIリクエスト・レスポンス詳細調査)
// ==========================================

/**
 * Yahoo! editItem API への実際のリクエストを送信し、
 * リクエスト・レスポンスの詳細をログに出力する調査用関数
 */
function debugYahooApiRequest() {
  console.log("=== Yahoo! API リクエスト詳細調査 開始 ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('▼商品マスタ(人間作業用)');
  const mapSheet = ss.getSheetByName('▼設定(Yahooマッピング)');
  
  if (!masterSheet || !mapSheet) {
    console.error("❌ 必要なシートが見つかりません。");
    return;
  }

  // 1. データ構築
  console.log("\n1. 商品データを構築中...");
  const builder = new YahooDataBuilder(masterSheet, mapSheet);
  const products = builder.buildProductData();
  
  if (products.length === 0) {
    console.error("❌ 出品対象の商品が見つかりません。出品CKを入れてください。");
    return;
  }
  const targetProduct = products[0];
  console.log(`   対象商品: ${targetProduct.code}`);

  // 2. APIクライアント初期化
  console.log("\n2. API認証情報を取得中...");
  const accessToken = getYahooAccessToken(mapSheet);
  const sellerId = mapSheet.getRange("B15").getValue();
  const client = new YahooApiClient(accessToken, sellerId, mapSheet);
  console.log(`   seller_id: ${sellerId}`);
  console.log(`   access_token: ${accessToken.substring(0, 20)}...`);

  // 3. Payload構築（updateItemと同じロジック）
  console.log("\n3. Payload構築中...");
  const payload = {
    "item_code": targetProduct.code,
    "path": client._getValue('path', targetProduct),
    "name": targetProduct.name,
    "price": targetProduct.originalRow['販売価格'] || 0,
    "product_category": client._getValue('product-category', targetProduct),
    "jan": client._getValue('jan', targetProduct),
    "headline": client._getValue('headline', targetProduct),
    "caption": client._getValue('caption', targetProduct),
    "explanation": client._getValue('explanation', targetProduct),
    "relevant_links": client._getValue('relevant-links', targetProduct),
    "ship_weight": client._getValue('ship-weight', targetProduct),
    "taxable": String(client._coerceTaxable(client._getValue('taxable', targetProduct))),
    "release_date": client._getValue('release-date', targetProduct),
    "temporary_point_term": client._getValue('temporary-point-term', targetProduct),
    "point_code": client._getValue('point-code', targetProduct),
    "sale_period_start": client._getValue('sale-period-start', targetProduct),
    "sale_period_end": client._getValue('sale-period-end', targetProduct),
    "delivery": client._getValue('delivery', targetProduct),
    "astk_code": client._getValue('astk-code', targetProduct),
    "condition": client._getValue('condition', targetProduct),
    "taojapan": client._getValue('taojapan', targetProduct),
    "lead_time_instock": client._getValue('lead-time-instock', targetProduct),
    "lead_time_outstock": client._getValue('lead-time-outstock', targetProduct),
    "brand_code": client._getValue('brand-code', targetProduct),
    "grouping_id": targetProduct.groupingId
  };

  // バリエーション設定（仕様: variation1_*）
  if (targetProduct.variationValue) {
    const freeTitle = client._getValue('variation1-free-title', targetProduct) || "セット内容";
    payload["variation1_free_title"] = freeTitle;
    payload["variation1_name"] = targetProduct.variationValue;
  }

  // seller_id追加と空パラメータ削除（_postと同じ処理）
  const preparedPayload = client._preparePayloadForPost(payload);
  const encodedBody = client._encodeFormPayload(preparedPayload);

  // 4. 必須パラメータチェック
  console.log("\n4. 必須パラメータチェック");
  const requiredParams = ['seller_id', 'item_code', 'path', 'name', 'product_category', 'price'];
  requiredParams.forEach(key => {
    const val = preparedPayload[key];
    if (val === undefined || val === null || val === "") {
      console.error(`   ❌ [${key}] が空です（必須パラメータ）`);
    } else {
      console.log(`   ✓ [${key}] = ${val} (type: ${typeof val})`);
    }
  });

  // 5. 全パラメータの一覧と型チェック
  console.log("\n5. 全パラメータ一覧（型・値確認）");
  Object.keys(preparedPayload).sort().forEach(key => {
    const val = preparedPayload[key];
    const type = typeof val;
    const strVal = String(val);
    const len = strVal.length;
    
    // 値が長すぎる場合は切り詰め
    const displayVal = len > 50 ? strVal.substring(0, 50) + "..." : strVal;
    
    console.log(`   [${key}]`);
    console.log(`      type: ${type}`);
    console.log(`      value: ${displayVal}`);
    console.log(`      length: ${len}`);
    
    // 数値型のチェック
    if (type === 'number') {
      if (isNaN(val)) {
        console.warn(`      ⚠️ NaN が検出されました`);
      }
      if (!isFinite(val)) {
        console.warn(`      ⚠️ Infinity が検出されました`);
      }
    }
    
    // taxable の特別チェック
    if (key === 'taxable') {
      console.log(`      [taxable特別チェック]`);
      console.log(`        値: ${val}`);
      console.log(`        型: ${type}`);
      console.log(`        数値チェック: ${val === 0 || val === 1 ? 'OK' : 'NG'}`);
      console.log(`        form-urlencoded: ${encodedBody.match(/(?:^|&)taxable=([^&]*)/)?.[1] || "(not found)"}`);
    }
  });

  // 6. 実際のHTTPリクエスト送信（詳細ログ付き）
  console.log("\n6. 実際のAPIリクエスト送信");
  console.log(`   URL: ${YAHOO_ITEM_API_URL}`);
  console.log(`   Method: POST`);
  console.log(`   Content-Type: application/x-www-form-urlencoded`);
  
  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    headers: { 
      "Authorization": `Bearer ${accessToken}`,
      "User-Agent": "Google Apps Script"
    },
    payload: preparedPayload,
    muteHttpExceptions: true
  };

  console.log(`   Headers:`);
  console.log(`      Authorization: Bearer ${accessToken.substring(0, 20)}...`);
  
  console.log(`   Payload (form-urlencoded, 先頭1000文字):`);
  console.log(`   ${encodedBody.substring(0, 1000)}${encodedBody.length > 1000 ? '...' : ''}`);
  console.log(`   Payload 総文字数: ${encodedBody.length}`);

  try {
    const res = UrlFetchApp.fetch(YAHOO_ITEM_API_URL, options);
    const statusCode = res.getResponseCode();
    const responseText = res.getContentText();
    const responseHeaders = res.getHeaders();

    console.log("\n7. APIレスポンス詳細");
    console.log(`   HTTP Status Code: ${statusCode}`);
    console.log(`   Response Headers:`);
    Object.keys(responseHeaders).forEach(h => {
      console.log(`      ${h}: ${responseHeaders[h]}`);
    });
    
    console.log(`   Response Body:`);
    console.log(`   ${responseText}`);

    // XMLパースしてエラー詳細を抽出
    if (responseText.includes('<Error>')) {
      console.log("\n8. エラー詳細解析");
      const errorMatch = responseText.match(/<Error>[\s\S]*?<Target>(.*?)<\/Target>[\s\S]*?<Code>(.*?)<\/Code>[\s\S]*?<Message><!\[CDATA\[(.*?)\]\]><\/Message>[\s\S]*?<\/Error>/);
      if (errorMatch) {
        console.log(`   エラー対象: ${errorMatch[1]}`);
        console.log(`   エラーコード: ${errorMatch[2]}`);
        console.log(`   エラーメッセージ: ${errorMatch[3]}`);
      }
      
      // 複数エラーがある場合
      const allErrors = responseText.match(/<Error>[\s\S]*?<\/Error>/g);
      if (allErrors && allErrors.length > 1) {
        console.log(`   検出されたエラー数: ${allErrors.length}`);
        allErrors.forEach((err, idx) => {
          const targetMatch = err.match(/<Target>(.*?)<\/Target>/);
          const codeMatch = err.match(/<Code>(.*?)<\/Code>/);
          const msgMatch = err.match(/<Message><!\[CDATA\[(.*?)\]\]><\/Message>/);
          console.log(`   エラー${idx + 1}:`);
          console.log(`      Target: ${targetMatch ? targetMatch[1] : 'N/A'}`);
          console.log(`      Code: ${codeMatch ? codeMatch[1] : 'N/A'}`);
          console.log(`      Message: ${msgMatch ? msgMatch[1] : 'N/A'}`);
        });
      }
    } else if (responseText.includes('<Status>OK</Status>')) {
      console.log("\n8. ✅ APIリクエスト成功");
    }

    // 警告がある場合も抽出
    if (responseText.includes('<Warning>')) {
      console.log("\n9. 警告情報");
      const warnings = responseText.match(/<Warning>[\s\S]*?<\/Warning>/g);
      warnings.forEach((warn, idx) => {
        const targetMatch = warn.match(/<Target>(.*?)<\/Target>/);
        const codeMatch = warn.match(/<Code>(.*?)<\/Code>/);
        const msgMatch = warn.match(/<Message><!\[CDATA\[(.*?)\]\]><\/Message>/);
        console.log(`   警告${idx + 1}:`);
        console.log(`      Target: ${targetMatch ? targetMatch[1] : 'N/A'}`);
        console.log(`      Code: ${codeMatch ? codeMatch[1] : 'N/A'}`);
        console.log(`      Message: ${msgMatch ? msgMatch[1] : 'N/A'}`);
      });
    }

  } catch (e) {
    console.error("\n7. ❌ リクエスト送信エラー");
    console.error(`   エラー: ${e.message}`);
    console.error(`   スタック: ${e.stack}`);
  }

  console.log("\n=== 調査終了 ===");
}

// ==========================================
// ■ 調査用デバッグ関数 9 (バリエーション・グルーピングID 送信内容確認)
// ==========================================

/**
 * it-14165「グルーピングIDとバリエーションは両方を設定してください」の原因特定のため、
 * 実際に送る payload の「grouping_id」「variation 系」パラメータ名・値をログ出力する
 */
function debugVariationAndGrouping() {
  console.log("=== バリエーション・グルーピングID 送信内容調査 開始 ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('▼商品マスタ(人間作業用)');
  const mapSheet = ss.getSheetByName('▼設定(Yahooマッピング)');
  if (!masterSheet || !mapSheet) {
    console.error("❌ 必要なシートが見つかりません。");
    return;
  }

  const builder = new YahooDataBuilder(masterSheet, mapSheet);
  const products = builder.buildProductData();
  if (products.length === 0) {
    console.error("❌ 出品対象の商品がありません。出品CKを入れてください。");
    return;
  }
  const targetProduct = products[0];
  const sellerId = mapSheet.getRange("B15").getValue();
  const client = new YahooApiClient("dummy_token", sellerId, mapSheet);

  // updateItem と同じロジックで payload を組み立て
  const payload = {
    "item_code": targetProduct.code,
    "path": client._getValue('path', targetProduct),
    "name": targetProduct.name,
    "price": targetProduct.originalRow['販売価格'] || 0,
    "product_category": client._getValue('product-category', targetProduct),
    "jan": client._getValue('jan', targetProduct),
    "headline": client._getValue('headline', targetProduct),
    "caption": client._getValue('caption', targetProduct),
    "explanation": client._getValue('explanation', targetProduct),
    "relevant_links": client._getValue('relevant-links', targetProduct),
    "ship_weight": client._getValue('ship-weight', targetProduct),
    "taxable": String(client._coerceTaxable(client._getValue('taxable', targetProduct))),
    "release_date": client._getValue('release-date', targetProduct),
    "temporary_point_term": client._getValue('temporary-point-term', targetProduct),
    "point_code": client._getValue('point-code', targetProduct),
    "sale_period_start": client._getValue('sale-period-start', targetProduct),
    "sale_period_end": client._getValue('sale-period-end', targetProduct),
    "delivery": client._getValue('delivery', targetProduct),
    "astk_code": client._getValue('astk-code', targetProduct),
    "condition": client._getValue('condition', targetProduct),
    "taojapan": client._getValue('taojapan', targetProduct),
    "lead_time_instock": client._getValue('lead-time-instock', targetProduct),
    "lead_time_outstock": client._getValue('lead-time-outstock', targetProduct),
    "brand_code": client._getValue('brand-code', targetProduct),
    "grouping_id": targetProduct.groupingId
  };
  if (targetProduct.variationValue) {
    const freeTitle = client._getValue('variation1-free-title', targetProduct) || "セット内容";
    payload["variation1_free_title"] = freeTitle;
    payload["variation1_name"] = targetProduct.variationValue;
  }
  const prepared = client._preparePayloadForPost(payload);

  console.log("\n1. 対象商品");
  console.log("   item_code: " + targetProduct.code);
  console.log("   groupingId: " + (targetProduct.groupingId || "(空)"));
  console.log("   variationValue: " + (targetProduct.variationValue || "(空)"));

  console.log("\n2. 送信 payload の全キー一覧（アルファベット順）");
  Object.keys(prepared).sort().forEach(k => console.log("   " + k));

  console.log("\n3. グルーピング・バリエーション関連のみ抽出");
  const related = Object.keys(prepared).filter(k =>
    /grouping|variation/i.test(k)
  );
  if (related.length === 0) {
    console.log("   ※ grouping / variation を含むキーが 1 つもありません。");
  } else {
    related.forEach(k => {
      console.log("   [" + k + "]");
      console.log("      type: " + typeof prepared[k]);
      console.log("      value: " + prepared[k]);
    });
  }

  console.log("\n4. Yahoo! editItem 仕様上のパラメータ名（参考）");
  console.log("   grouping_id : グルーピングID（半角英数字・ハイフンのみ）");
  console.log("   variation1_free_title : バリエーション1 - 項目名（※ variation と 1 の間はアンダースコアなし）");
  console.log("   variation1_name       : バリエーション1 - 表示名");
  console.log("   variation1_spec_id    : バリエーション1 - スペックID（任意）");

  console.log("\n5. 差異チェック");
  const hasGrouping = prepared.hasOwnProperty("grouping_id");
  const hasVariation1FreeTitle = prepared.hasOwnProperty("variation1_free_title");
  const hasVariation1Name = prepared.hasOwnProperty("variation1_name");
  const hasVariation_1FreeTitle = prepared.hasOwnProperty("variation_1_free_title");
  const hasVariation_1Name = prepared.hasOwnProperty("variation_1_name");
  console.log("   grouping_id を送信: " + (hasGrouping ? "はい" : "いいえ"));
  console.log("   variation1_free_title を送信: " + (hasVariation1FreeTitle ? "はい" : "いいえ"));
  console.log("   variation1_name を送信: " + (hasVariation1Name ? "はい" : "いいえ"));
  console.log("   variation_1_free_title を送信: " + (hasVariation_1FreeTitle ? "はい" : "いいえ"));
  console.log("   variation_1_name を送信: " + (hasVariation_1Name ? "はい" : "いいえ"));
  if (hasVariation_1FreeTitle && !hasVariation1FreeTitle) {
    console.log("   ⚠️ 現在は variation_1_*（アンダースコア入り）で送っており、仕様の variation1_* と一致していません。");
  }
  if (hasVariation_1Name && !hasVariation1Name) {
    console.log("   ⚠️ 現在は variation_1_name で送っており、仕様の variation1_name と一致していません。");
  }

  console.log("\n=== 調査終了 ===");
}

// ==========================================
// ■ 調査用デバッグ関数 10 (在庫API setStock 送信内容確認)
// ==========================================

/**
 * st-02104「quantity(33.0) is invalid value」の原因特定のため、
 * setStock に渡す quantity の値・型をログ出力する
 */
function debugSetStockQuantity() {
  console.log("=== setStock 送信内容（quantity）調査 開始 ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('▼商品マスタ(人間作業用)');
  const mapSheet = ss.getSheetByName('▼設定(Yahooマッピング)');
  if (!masterSheet || !mapSheet) {
    console.error("❌ 必要なシートが見つかりません。");
    return;
  }
  const builder = new YahooDataBuilder(masterSheet, mapSheet);
  const products = builder.buildProductData();
  if (products.length === 0) {
    console.error("❌ 出品対象の商品がありません。");
    return;
  }

  console.log("\n1. 各商品の quantity の値・型（setStock にそのまま渡しているもの）");
  products.forEach((p, i) => {
    const q = p.quantity;
    console.log(`   商品${i + 1}: item_code=${p.code}`);
    console.log(`      quantity = ${q}`);
    console.log(`      typeof quantity = ${typeof q}`);
    console.log(`      Number.isInteger(quantity) = ${Number.isInteger(q)}`);
    if (typeof q === "number" && !Number.isInteger(q)) {
      console.log("      ⚠️ 小数です。在庫APIは整数のみ受け付ける可能性があります。");
    }
  });

  console.log("\n2. 原因の説明");
  console.log("   エラー st-02104: quantity(33.0) is invalid value.");
  console.log("   Yahoo 在庫API(setStock) は在庫数を「整数」で受け付ける仕様のため、");
  console.log("   33.0 のように小数として送信されるとエラーになります。");
  console.log("   スプレッドシートの「在庫数」が数値型で読み込まれると、");
  console.log("   そのまま product.quantity に入り、送信時に 33.0 になることがあります。");
  console.log("   対処: 送信前に整数に変換する（例: Math.floor(Number(quantity))）。");

  console.log("\n=== 調査終了 ===");
}

// ==========================================
// ■ 調査用デバッグ関数: 価格と画像の調査
// ==========================================

function debugPriceAndImages() {
  console.log("=== 価格と画像の調査 開始 ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('▼商品マスタ(人間作業用)');
  const mapSheet = ss.getSheetByName('▼設定(Yahooマッピング)');

  if (!masterSheet || !mapSheet) {
    console.error("❌ 必要なシートが見つかりません。");
    return;
  }

  // ========== 1. マッピング設定の確認 ==========
  console.log("\n========== 1. マッピングシートの price 設定 ==========");
  const mapData = mapSheet.getDataRange().getValues();
  const yahooFields = mapData[0];  // 1行目: Yahoo項目名
  const masterCols = mapData[3];   // 4行目: マスタ参照列
  const fixedVals = mapData[4];    // 5行目: 固定値/設定

  const priceIdx = yahooFields.indexOf('price');
  if (priceIdx === -1) {
    console.log("   ❌ 'price' フィールドがマッピングシートに見つかりません");
    yahooFields.forEach((f, i) => {
      if (f && String(f).toLowerCase().includes('price')) {
        console.log(`   候補: column ${i + 1} = "${f}"`);
      }
    });
  } else {
    console.log(`   列位置: ${priceIdx + 1}`);
    console.log(`   マスタ参照列 (4行目): "${masterCols[priceIdx]}"`);
    console.log(`   固定値/設定 (5行目): "${fixedVals[priceIdx]}"`);
  }

  // 全てのprice関連フィールドを表示
  console.log("\n   --- price関連フィールド一覧 ---");
  yahooFields.forEach((f, i) => {
    const fStr = String(f).toLowerCase();
    if (fStr.includes('price') || fStr === '販売価格' || fStr === 'yahoo!価格設定') {
      console.log(`   [${i + 1}] "${f}" → マスタ参照: "${masterCols[i]}" / 固定値: "${fixedVals[i]}"`);
    }
  });

  // ========== 2. 商品データの構築と確認 ==========
  console.log("\n========== 2. 商品データの価格確認 ==========");
  const builder = new YahooDataBuilder(masterSheet, mapSheet);
  const products = builder.buildProductData();

  if (products.length === 0) {
    console.error("❌ 出品対象の商品がありません。");
    return;
  }

  // 最初の商品の価格情報を詳しく表示
  const p = products[0];
  console.log(`\n   対象商品: ${p.code}`);
  console.log(`   --- originalRow の価格関連項目 ---`);
  
  const priceRelatedCols = ['販売価格', 'Yahoo!価格設定', 'price', '参考価格', 'セール価格'];
  priceRelatedCols.forEach(col => {
    const val = p.originalRow[col];
    if (val !== undefined) {
      console.log(`   "${col}" = "${val}" (type: ${typeof val})`);
    }
  });

  // APIクライアント経由で _getValue の結果を確認
  console.log("\n   --- _getValue('price', product) の結果 ---");
  const client = new YahooApiClient(mapSheet);
  const priceFromMapping = client._getValue('price', p);
  console.log(`   _getValue('price', p) = "${priceFromMapping}" (type: ${typeof priceFromMapping})`);
  
  // フォールバック値
  const fallbackPrice = p.originalRow['販売価格'];
  console.log(`   p.originalRow['販売価格'] = "${fallbackPrice}" (type: ${typeof fallbackPrice})`);
  
  // 最終的に送信される値
  const finalPrice = priceFromMapping || p.originalRow['販売価格'] || 0;
  console.log(`   最終的に送信される price = "${finalPrice}"`);

  // ========== 3. 画像の確認 ==========
  console.log("\n========== 3. 画像データの確認 ==========");
  products.forEach((prod, idx) => {
    console.log(`\n   商品${idx + 1}: ${prod.code}`);
    console.log(`   画像枚数: ${prod.images.length}`);
    prod.images.forEach((img, imgIdx) => {
      const truncUrl = img.length > 60 ? img.substring(0, 60) + "..." : img;
      console.log(`      [${imgIdx}] ${prod.imageFiles[imgIdx]} ← ${truncUrl}`);
    });
  });

  // 兄弟画像が含まれているか確認
  console.log("\n   --- 兄弟画像の含有チェック ---");
  if (products.length >= 2) {
    const p1Images = products[0].images;
    const p2MainImage = products[1].images[0];
    const p1HasP2Main = p1Images.includes(p2MainImage);
    console.log(`   商品1に商品2のメイン画像が含まれているか: ${p1HasP2Main ? "✅ YES" : "❌ NO"}`);
    
    if (!p1HasP2Main && p2MainImage) {
      console.log(`   商品2のメイン画像URL: ${p2MainImage.substring(0, 60)}...`);
      console.log("   → 兄弟画像の追加ロジックに問題がある可能性");
    }
  }

  console.log("\n=== 調査終了 ===");
}

// ==========================================
// ■ マスタシート更新: Yahoo出品済ID記録
// ==========================================

function updateMasterYahooId(masterSheet, childSku, yahooItemCode) {
  const data = masterSheet.getDataRange().getValues();
  const headers = data[7]; // 8行目がヘッダー
  
  const childSkuCol = headers.indexOf('子SKU');
  const yahooIdCol = headers.indexOf('Yahoo出品済ID');
  
  if (childSkuCol === -1 || yahooIdCol === -1) {
    console.warn("⚠️ マスタシートに「子SKU」または「Yahoo出品済ID」列が見つかりません");
    return;
  }
  
  // 子SKUに一致する行を探して更新
  for (let i = 8; i < data.length; i++) {
    if (String(data[i][childSkuCol]).trim() === String(childSku).trim()) {
      masterSheet.getRange(i + 1, yahooIdCol + 1).setValue(yahooItemCode);
      return;
    }
  }
}

// ==========================================
// ■ 出品完了メール送信
// ==========================================

function sendYahooCompletionEmail(products, sellerId, successCount, errorCount, errorItems) {
  errorItems = errorItems || [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUrl = ss.getUrl();
  const mapSheet = ss.getSheetByName(SHEET_NAME_YAHOO_MAP);

  // 反映用URL: B18 に http で始まるURLがあればそれを使用。それ以外はストアクリエイターpro トップ
  const YAHOO_PRO_STORE_TOP = "https://pro.store.yahoo.co.jp";
  let reflectUrl = `${YAHOO_PRO_STORE_TOP}/pro.${sellerId}`;
  if (mapSheet && mapSheet.getRange("B18").getValue()) {
    const customReflect = String(mapSheet.getRange("B18").getValue()).trim();
    if (customReflect && (customReflect.startsWith("http://") || customReflect.startsWith("https://"))) {
      reflectUrl = customReflect;
    }
  }

  // メール本文を構築
  let body = `Yahoo!出品処理が完了しました。\n\n`;
  body += `成功: ${successCount}件 / 失敗: ${errorCount}件\n\n`;

  if (errorItems.length > 0) {
    body += `--- 失敗した商品（原因） ---\n\n`;
    errorItems.forEach(function (err, i) {
      const truncName = (err.name || "").length > 30 ? (err.name || "").substring(0, 30) + "..." : (err.name || "");
      body += `${i + 1}. 商品コード: ${err.code}\n`;
      body += `   商品名: ${truncName}\n`;
      body += `   エラー: ${(err.message || "").substring(0, 300)}\n\n`;
    });
    body += `上記の商品はマスタで内容を確認・修正し、再度「Yahoo!出品」で出品してください。\n\n`;
  }

  body += `--- 出品済み商品一覧 ---\n\n`;

  const deleteBaseUrl = getWebAppUrl();

  products.forEach((product, idx) => {
    const itemCode = product.code;
    const productName = product.name || "";
    const truncName = productName.length > 30 ? productName.substring(0, 30) + "..." : productName;

    const storeUrl = `${YAHOO_STORE_BASE_URL}/${sellerId}/${itemCode}.html`;
    const deleteUrl = `${deleteBaseUrl}?action=delete&item_code=${encodeURIComponent(itemCode)}&name=${encodeURIComponent(productName)}`;

    body += `${idx + 1}. ${truncName}\n`;
    body += `   商品コード: ${itemCode}\n`;
    body += `   商品ページ: ${storeUrl}\n`;
    body += `   削除用: ${deleteUrl}\n\n`;
  });

  body += `\n--- 新規出品の反映（スマホから手動で） ---\n`;
  body += `新規出品はAPIで自動反映できません。下のリンクをタップ → ログイン → 商品ページ → 「反映」ボタンで一括反映\n`;
  body += `   ${reflectUrl}\n`;
  body += `※開けない場合は「▼設定(Yahooマッピング)」の B18 に開けるURLを設定してください。\n\n`;

  body += `--- 削除が必要な場合 ---\n`;
  body += `【スマホ】各商品の「削除用」のURLをタップ→確認画面で「削除する」で削除できます。\n\n`;
  body += `【PC】スプレッドシートで削除したい行の「削除CK」にレ点を付け、メニュー「🛒 Yahoo!出品」→「🗑️ 商品を削除...」でまとめて削除できます。\n   ${sheetUrl}\n\n`;
  body += `削除後はマスタシートで修正し、再出品してください。\n`;
  
  // メール送信先を取得（スクリプトプロパティ or マッピングシートB16）
  let recipient = "";
  try {
    const props = PropertiesService.getScriptProperties();
    recipient = props.getProperty('NOTIFICATION_EMAIL');
  } catch (e) {}
  
  if (!recipient) {
    const mapSheet = ss.getSheetByName(SHEET_NAME_YAHOO_MAP);
    if (mapSheet) {
      recipient = mapSheet.getRange("B17").getValue(); // 成否メール送信宛先
    }
  }
  
  if (!recipient) {
    console.warn("   ⚠️ 通知メールアドレスが設定されていません（マッピングシートB17「成否メール送信宛先」に設定してください）");
    return;
  }
  
  const subject = `【Yahoo!出品完了】${successCount}件出品 / ${errorCount}件失敗`;
  
  try {
    MailApp.sendEmail(recipient, subject, body);
    console.log(`   -> 完了メールを送信しました: ${recipient}`);
  } catch (e) {
    console.warn(`   ⚠️ メール送信に失敗しました: ${e.message}`);
  }
}

// WebアプリのURLを取得（削除フォームの送信先・完了メールの削除リンク）
function getWebAppUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) return url;
  } catch (err) {
    // エディタ実行時・スプレッドシート実行時は getService() が使えない場合あり
  }
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('YAHOO_DELETE_WEBAPP_URL');
  if (url) return url;
  return YAHOO_DELETE_WEBAPP_BASE_URL;
}

// テスト用: メール送信の権限承認（初回のみ実行）
function testMailPermission() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(SHEET_NAME_YAHOO_MAP);
  const recipient = mapSheet.getRange("B17").getValue();
  
  if (!recipient) {
    console.log("❌ B17にメールアドレスが設定されていません");
    return;
  }
  
  MailApp.sendEmail(recipient, "【テスト】Yahoo!出品システム", "これはメール送信のテストです。このメールが届いていれば、権限の承認が完了しています。");
  console.log(`✅ テストメールを送信しました: ${recipient}`);
}

// テスト用: ダイアログ表示の権限承認（初回のみ実行）
function testDialogPermission() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ テスト成功', 'ダイアログの権限が承認されています。', ui.ButtonSet.OK);
  console.log('✅ ダイアログ表示テスト成功');
}

// WebアプリURLを設定する関数（初回デプロイ後に実行）
function setWebAppUrl(url) {
  PropertiesService.getScriptProperties().setProperty('YAHOO_DELETE_WEBAPP_URL', url);
  console.log(`WebアプリURLを設定しました: ${url}`);
}

// ==========================================
// ■ Webアプリ: 削除確認ページ (doGet)
// ==========================================

function doGet(e) {
  // Webアプリとして呼ばれたときのみ e が渡る（エディタ実行時は undefined）
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action;
  const itemCode = params.item_code;
  const productName = params.name || "";
  const itemUrl = params.item_url;

  if (action === 'delete' && itemCode) {
    const html = getDeleteConfirmHtml(itemCode, productName);
    return HtmlService.createHtmlOutput(html)
      .setTitle('Yahoo!商品削除確認')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action === 'rakuten_delete' && itemUrl) {
    const html = getRakutenDeleteConfirmHtml(itemUrl, productName);
    return HtmlService.createHtmlOutput(html)
      .setTitle('楽天 商品削除確認')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutput('<h1>無効なリクエストです</h1>');
}

// ==========================================
// ■ Webアプリ: 削除実行 (doPost)
// ==========================================

function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action;
  const itemCode = params.item_code;
  const itemUrl = params.item_url;

  if (action === 'rakuten_delete' && itemUrl) {
    try {
      const result = executeRakutenDeleteFromWebApp(itemUrl);
      const html = getRakutenDeleteResultHtml(result.success, itemUrl, result.message);
      return HtmlService.createHtmlOutput(html)
        .setTitle('楽天 削除結果')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (error) {
      const html = getRakutenDeleteResultHtml(false, itemUrl, error.message);
      return HtmlService.createHtmlOutput(html)
        .setTitle('楽天 削除結果')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  if (!itemCode) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: '商品コードが指定されていません'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const result = deleteYahooItem(itemCode);
    const html = getDeleteResultHtml(result.success, itemCode, result.message);
    return HtmlService.createHtmlOutput(html)
      .setTitle('削除結果')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    const html = getDeleteResultHtml(false, itemCode, error.message);
    return HtmlService.createHtmlOutput(html)
      .setTitle('削除結果')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ==========================================
// ■ Yahoo! API: 商品削除
// ==========================================

function deleteYahooItem(itemCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(SHEET_NAME_YAHOO_MAP);
  const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
  
  if (!mapSheet) {
    return { success: false, message: 'マッピングシートが見つかりません' };
  }
  
  // APIトークン取得
  const accessToken = getYahooAccessToken(mapSheet);
  const sellerId = mapSheet.getRange("B15").getValue();
  
  // 削除API実行
  const payload = {
    "seller_id": sellerId,
    "item_code": itemCode
  };
  
  const options = {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/x-www-form-urlencoded"
    },
    payload: payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(YAHOO_DELETE_API_URL, options);
    const code = response.getResponseCode();
    const content = response.getContentText();
    
    if (code === 200) {
      // マスタシートの削除フラグを更新
      updateMasterDeleteFlag(masterSheet, itemCode);
      
      logYahooEvent("DELETE", "API", "削除成功", itemCode);
      return { success: true, message: '商品を削除しました' };
    } else {
      logYahooEvent("DELETE_ERROR", "API", `削除失敗(${code})`, content);
      return { success: false, message: `削除に失敗しました: ${content}` };
    }
  } catch (error) {
    logYahooEvent("DELETE_ERROR", "API", "例外発生", error.message);
    return { success: false, message: `エラーが発生しました: ${error.message}` };
  }
}

// マスタシートの削除フラグを更新
function updateMasterDeleteFlag(masterSheet, yahooItemCode) {
  const data = masterSheet.getDataRange().getValues();
  const headers = data[7]; // 8行目がヘッダー
  
  const yahooIdCol = headers.indexOf('Yahoo出品済ID');
  const deleteFlagCol = headers.indexOf('Yahoo削除フラグ');
  const deleteDateCol = headers.indexOf('Yahoo削除日時');
  
  if (yahooIdCol === -1) {
    console.warn("⚠️「Yahoo出品済ID」列が見つかりません");
    return;
  }
  
  // Yahoo出品済IDに一致する行を探して更新
  for (let i = 8; i < data.length; i++) {
    if (String(data[i][yahooIdCol]).trim() === String(yahooItemCode).trim()) {
      if (deleteFlagCol !== -1) {
        masterSheet.getRange(i + 1, deleteFlagCol + 1).setValue(true);
      }
      if (deleteDateCol !== -1) {
        masterSheet.getRange(i + 1, deleteDateCol + 1).setValue(new Date());
      }
      console.log(`   -> マスタシートの削除フラグを更新: 行${i + 1}`);
      return;
    }
  }
}

// ==========================================
// ■ HTML生成: 削除確認ページ
// ==========================================

function getDeleteConfirmHtml(itemCode, productName) {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: 'Helvetica Neue', Arial, sans-serif;
      max-width: 500px;
      margin: 50px auto;
      padding: 20px;
      background-color: #f5f5f5;
    }
    .card {
      background: white;
      border-radius: 8px;
      padding: 30px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    h1 {
      color: #e74c3c;
      font-size: 24px;
      margin-bottom: 20px;
    }
    .info {
      background: #f8f9fa;
      border-radius: 4px;
      padding: 15px;
      margin: 20px 0;
    }
    .info label {
      font-weight: bold;
      color: #666;
      font-size: 12px;
    }
    .info p {
      margin: 5px 0 15px 0;
      font-size: 14px;
      word-break: break-all;
    }
    .buttons {
      display: flex;
      gap: 10px;
      margin-top: 20px;
    }
    button {
      flex: 1;
      padding: 12px 20px;
      border: none;
      border-radius: 4px;
      font-size: 16px;
      cursor: pointer;
      transition: opacity 0.2s;
    }
    button:hover {
      opacity: 0.8;
    }
    .delete-btn {
      background: #e74c3c;
      color: white;
    }
    .cancel-btn {
      background: #95a5a6;
      color: white;
    }
    .warning {
      color: #e74c3c;
      font-size: 14px;
      margin-top: 15px;
    }
  </style>
</head>
<body>
  <div class="card">
    <h1>⚠️ Yahoo!商品削除確認</h1>
    <p>以下の商品をYahoo!ストアから削除しますか？</p>
    
    <div class="info">
      <label>商品コード</label>
      <p>${itemCode}</p>
      <label>商品名</label>
      <p>${productName || '（取得できませんでした）'}</p>
    </div>
    
    <p class="warning">※この操作は取り消せません。削除後はマスタシートで修正して再出品してください。</p>
    
    <form action="${getWebAppUrl()}" method="post">
      <input type="hidden" name="item_code" value="${itemCode}">
      <div class="buttons">
        <button type="submit" class="delete-btn">削除する</button>
        <button type="button" class="cancel-btn" onclick="document.body.innerHTML='<div class=\\'card\\'><h1 style=\\'color:#95a5a6;\\'>キャンセルしました</h1><p>このタブを閉じてください。</p></div>'">キャンセル</button>
      </div>
    </form>
  </div>
</body>
</html>`;
}

// ==========================================
// ■ HTML生成: 削除結果ページ
// ==========================================

function getDeleteResultHtml(success, itemCode, message) {
  const color = success ? '#27ae60' : '#e74c3c';
  const icon = success ? '✅' : '❌';
  const title = success ? '削除完了' : '削除失敗';
  
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: 'Helvetica Neue', Arial, sans-serif;
      max-width: 500px;
      margin: 50px auto;
      padding: 20px;
      background-color: #f5f5f5;
    }
    .card {
      background: white;
      border-radius: 8px;
      padding: 30px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      text-align: center;
    }
    h1 {
      color: ${color};
      font-size: 24px;
      margin-bottom: 20px;
    }
    .icon {
      font-size: 48px;
      margin-bottom: 20px;
    }
    .message {
      background: #f8f9fa;
      border-radius: 4px;
      padding: 15px;
      margin: 20px 0;
      font-size: 14px;
    }
    .item-code {
      font-family: monospace;
      background: #eee;
      padding: 5px 10px;
      border-radius: 4px;
    }
    .next-steps {
      text-align: left;
      margin-top: 20px;
      padding: 15px;
      background: #fff3cd;
      border-radius: 4px;
    }
    .next-steps h3 {
      margin-top: 0;
      font-size: 14px;
    }
    .next-steps ol {
      margin: 10px 0 0 0;
      padding-left: 20px;
      font-size: 13px;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="icon">${icon}</div>
    <h1>${title}</h1>
    
    <p>商品コード: <span class="item-code">${itemCode}</span></p>
    
    <div class="message">${message}</div>
    
    ${success ? `
    <div class="next-steps">
      <h3>📋 次のステップ</h3>
      <ol>
        <li>スプレッドシートのマスタを開く</li>
        <li>「Yahoo削除フラグ」でフィルタ</li>
        <li>データを修正</li>
        <li>「出品CK」を入れて再出品</li>
      </ol>
    </div>
    ` : ''}
    
    <p style="margin-top: 20px;">
      <a href="javascript:window.close()">このウィンドウを閉じる</a>
    </p>
  </div>
</body>
</html>`;
}

// ==========================================
// ■ HTML生成: 楽天 削除確認・結果ページ
// ==========================================

function getRakutenDeleteConfirmHtml(itemUrl, productName) {
  const postUrl = getWebAppUrl();
  const esc = (s) => (s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; max-width: 500px; margin: 50px auto; padding: 20px; background-color: #f5f5f5; }
    .card { background: white; border-radius: 8px; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
    h1 { color: #e74c3c; font-size: 24px; margin-bottom: 20px; }
    .info { background: #f8f9fa; border-radius: 4px; padding: 15px; margin: 20px 0; }
    .info label { font-weight: bold; color: #666; font-size: 12px; }
    .info p { margin: 5px 0 15px 0; font-size: 14px; word-break: break-all; }
    .buttons { display: flex; gap: 10px; margin-top: 20px; }
    button { flex: 1; padding: 12px 20px; border: none; border-radius: 4px; font-size: 16px; cursor: pointer; }
    .delete-btn { background: #e74c3c; color: white; }
    .cancel-btn { background: #95a5a6; color: white; }
    .warning { color: #e74c3c; font-size: 14px; margin-top: 15px; }
  </style>
</head>
<body>
  <div class="card">
    <h1>⚠️ 楽天 商品削除確認</h1>
    <p>以下の商品を楽天から削除しますか？（item-delete.csv をMAKE経由でFTPに送信します）</p>
    <div class="info">
      <label>商品管理番号</label>
      <p>${esc(itemUrl)}</p>
      <label>商品名</label>
      <p>${esc(productName || '（取得できませんでした）')}</p>
    </div>
    <p class="warning">※この操作は取り消せません。削除後はマスタで楽天削除フラグを確認し、必要に応じて再出品してください。</p>
    <form action="${esc(postUrl)}" method="post">
      <input type="hidden" name="action" value="rakuten_delete">
      <input type="hidden" name="item_url" value="${esc(itemUrl)}">
      <input type="hidden" name="name" value="${esc(productName)}">
      <div class="buttons">
        <button type="submit" class="delete-btn">削除する</button>
        <button type="button" class="cancel-btn" onclick="document.body.innerHTML='<div class=\\'card\\'><h1 style=\\'color:#95a5a6;\\'>キャンセルしました</h1><p>このタブを閉じてください。</p></div>'">キャンセル</button>
      </div>
    </form>
  </div>
</body>
</html>`;
}

function getRakutenDeleteResultHtml(success, itemUrl, message) {
  const color = success ? '#27ae60' : '#e74c3c';
  const icon = success ? '✅' : '❌';
  const title = success ? '削除依頼送信完了' : '削除失敗';
  const esc = (s) => (s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; max-width: 500px; margin: 50px auto; padding: 20px; background-color: #f5f5f5; }
    .card { background: white; border-radius: 8px; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center; }
    h1 { color: ${color}; font-size: 24px; margin-bottom: 20px; }
    .message { background: #f8f9fa; border-radius: 4px; padding: 15px; margin: 20px 0; font-size: 14px; }
    .item-code { font-family: monospace; background: #eee; padding: 5px 10px; border-radius: 4px; }
    .next-steps { text-align: left; margin-top: 20px; padding: 15px; background: #fff3cd; border-radius: 4px; font-size: 13px; }
  </style>
</head>
<body>
  <div class="card">
    <div style="font-size:48px;margin-bottom:20px;">${icon}</div>
    <h1>${title}</h1>
    <p>商品管理番号: <span class="item-code">${esc(itemUrl)}</span></p>
    <div class="message">${esc(message)}</div>
    ${success ? '<div class="next-steps"><strong>📋 次のステップ</strong><br>マスタの「楽天削除フラグ」「楽天削除日時」を確認し、必要に応じて再出品してください。</div>' : ''}
    <p style="margin-top: 20px;"><a href="javascript:window.close()">このウィンドウを閉じる</a></p>
  </div>
</body>
</html>`;
}