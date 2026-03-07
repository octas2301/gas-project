// ==========================================
// 【AI出品ツール Ver 36.0 完全版】 Part 1/3
// ==========================================

// ==========================================
// 0. 設定・グローバル定数
// ==========================================
// APIキーは Script Properties で設定（GAS エディタ → プロジェクトの設定 → スクリプト プロパティ）
function getOpenAiApiKey() { return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || ''; }
function getGeminiApiKey() { return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || ''; }
function getRakutenLicenseKey() { return PropertiesService.getScriptProperties().getProperty('RAKUTEN_LICENSE_KEY') || ''; }
function getRakutenServiceSecret() { return PropertiesService.getScriptProperties().getProperty('RAKUTEN_SERVICE_SECRET') || ''; }

// ★Googleドライブの「フォルダID」
const RAKUTEN_GENRE_FOLDER_ID = '1294Hqc_1bTNu5Q2_ah7sakC3Gin9350c';
const YAHOO_CATEGORY_FOLDER_ID = '1qhxHd1sjaxLfe8IVmTGf38xTE9D9p-pw'; 
const YAHOO_BRAND_FOLDER_ID = '1wZ0yCmp0dmLvq8YdFwZ7ceCAqRSx0npO';    

// ★CSV保存先・移動先のフォルダID
const CSV_SAVE_FOLDER_ID    = '1tcq7Mb9695pOcc236cq0NQpyjl8QAs3X'; 
const CSV_ARCHIVE_FOLDER_ID = '1Fc1KeJS4cmOCWCX62wqWPUW3BjUJYCtb'; 

// ★最新・最適モデルへ変更 (2026/02/10更新)
const MODEL_OPENAI = 'gpt-5.2'; 
const MODEL_GEMINI = 'gemini-3-flash-preview'; 

// ★シート名の設定
const TARGET_SHEET_NAME = 'AI情報取得data';
const MASTER_SHEET_NAME = '▼商品マスタ(人間作業用)';
/** ASIN貼り付け（Keepa用）シート名。RESEARCH_AND_ESTIMATE §8.8.9 運用 */
const ASIN_PASTE_SHEET_NAME = 'ASIN貼り付け（Keepa用）';
const SETTING_SHEET_NAME = '▼設定(マッピング)';
/** 手数料・送料等の参照元。A=設定項目, B=キー, D=値1(料率等)。販売手数料は 設定項目=販売手数料 の行のキーで検索。 */
const SETTINGS_MASTER_SHEET_NAME = '00_設定マスタ';
const SETTINGS_ITEM_SALES_FEE = '販売手数料';
const VERIFY_SHEET_NAME  = 'a.楽天ULファイル確認用';
const RAKUTEN_REFLECTION_LOG_SHEET = '楽天反映確認ログ';
const BATCH_EXPORT_LOG_SHEET = '一括出品ログ';

// ★ヘッダー検索用アンカー（この列がある行をヘッダーとみなす）
const ANCHOR_HEADER_NAME = 'ASINコード';

// ★出力判定用のチェックボックス列名
const CHECKBOX_HEADER_NAME = '出品CK';

// ★出力ファイル名 (楽天仕様: normal-item.csv 固定)
const OUTPUT_FILE_NAME = 'normal-item.csv';

// 楽天削除用WebアプリのURL（完了メールの「削除用」リンク。デプロイ後に更新）
const RAKUTEN_DELETE_WEBAPP_BASE_URL = '';

// ■ 商品リサーチ: 競合価格のマスタ列名（▼商品マスタ(人間作業用)）。スプシヘッダーに合わせて半角スペースなし。
const COL_COMPETITIVE_PRICE_AMAZON = '競合価格amazon';
const COL_COMPETITIVE_PRICE_RAKUTEN = '競合価格楽天';
const COL_COMPETITIVE_PRICE_YAHOO = '競合価格Yahoo!';
// ■ 商品リサーチ: AI提案の書き込み先（販売価格・セット数）。価格・卸値は税込で統一。
const COL_PRICE_AMAZON = '販売価格amazon';
const COL_PRICE_RAKUTEN = '楽天価格設定';
const COL_PRICE_YAHOO = 'Yahoo!価格設定';
const COL_MASTER_TOTAL_QTY = 'A.セット商品数';
const COL_COMPETITOR_URL_AMAZON = '競合AmazonページURL';
/** Yahoo! セット別競合: 商品ページURL・要確認・メモ（▼商品マスタ(人間作業用)）。 */
const COL_COMPETITOR_URL_YAHOO = '競合URLYahoo!';
const COL_REVIEW_STATUS_YAHOO = '要確認Yahoo!';
const COL_REVIEW_MEMO_YAHOO = '確認内容メモYahoo!';
/** 楽天セット別競合: 商品ページURL・要確認・メモ。docs/RAKUTEN_YAHOO_COMPETITIVE_PRICE_REQUIREMENTS.md 参照。 */
const COL_COMPETITOR_URL_RAKUTEN = '競合URL楽天';
const COL_REVIEW_STATUS_RAKUTEN = '要確認楽天';
const COL_REVIEW_MEMO_RAKUTEN = '確認内容メモ楽天';
/** 楽天 商品名一致スコア: この値以上を「同一商品候補」として採用。変更可。ログは [楽天商品名一致] で検索。 */
var RAKUTEN_NAME_MATCH_THRESHOLD = 50;
/** 楽天 商品名一致: 各ルールの得点。変更可。合計は100でキャップ。 */
var RAKUTEN_NAME_MATCH_FULL_MATCH_SCORE = 50;   // 期待名がヒットに含まれる or 逆
var RAKUTEN_NAME_MATCH_WORD_MAX_SCORE = 40;     // 単語一致（最大）
var RAKUTEN_NAME_MATCH_MAKER_SCORE = 10;        // メーカーがヒットに含まれる
var RAKUTEN_NAME_MATCH_CHAR_MAX_SCORE = 20;     // 文字一致ボーナス（最大）
/** 楽天 商品名一致: 各ルールの有効フラグ。false にするとそのルールは加算されない。 */
var RAKUTEN_NAME_MATCH_USE_FULL = true;
var RAKUTEN_NAME_MATCH_USE_WORD = true;
var RAKUTEN_NAME_MATCH_USE_MAKER = true;
var RAKUTEN_NAME_MATCH_USE_CHAR = true;
const COL_COST_TAX_IN = '卸値(税込)';
/** 手数料計算用。マスタに無い行はアラート＋該当セルを赤くする（§8.8.20 T-3）。 */
const COL_AMAZON_CATEGORY = 'amazon カテゴリー';
/** 送料（円）。空欄時は0として計算。人間が設定する運用（§8.8.14 F）。列が無くてもエラーにしない。 */
const COL_SHIPPING = '送料';
/** 確定送料（円）。このフェーズではこちらを優先。列が無い場合は「送料」を参照。 */
const COL_SHIPPING_FIXED = '確定送料';
/** セット卸値（税込み）。行のセット数に対する卸値。列が無い場合は「卸値(税込)」を参照。 */
const COL_COST_SET_TAX_IN = 'セット卸値（税込み）';
/** amazon手数料計（円）。販売価格×手数料率の結果を書き込む。 */
const COL_AMAZON_FEE = 'amazon手数料計';
/** CPO値決めの返答全文を書き込む列。親SKU行にのみ書き込む。 */
const COL_AMAZON_STRATEGY = 'amazon価格戦略';
/** CPOマッピング（プレースホルダ→マスタ列）を記載する 00_設定マスタ の行範囲。 */
const CPO_MAPPING_FIRST_ROW = 92;
const CPO_MAPPING_LAST_ROW  = 100;

/** リサーチでマスタに書き込む際の必須列一覧（実行時チェック用）。コード.js 先頭の定数と docs/MASTER_LINKAGE_TASKS.md「マスタ書き込み用 固定列一覧」と一致させる。 */
const MASTER_WRITE_COLUMNS_RESEARCH = [
  COL_COMPETITIVE_PRICE_AMAZON,
  COL_COMPETITIVE_PRICE_RAKUTEN,
  COL_COMPETITIVE_PRICE_YAHOO,
  COL_PRICE_AMAZON,
  COL_PRICE_RAKUTEN,
  COL_PRICE_YAHOO,
  COL_MASTER_TOTAL_QTY,
  COL_COST_TAX_IN
];
// ■ 優先4で追加予定のマスタ列（列名確定後に定数化し MASTER_WRITE_COLUMNS_RESEARCH に追加する）: 競合セット数, 穴のセット数提案
/** マスタで◎が付いている行を対象とするときの列名。この列のセルに「◎」が含まれる行を対象とする。 */
const COL_MASTER_MARK_TARGET = '評価';
/** 優先4: マスタに書き込む「競合セット数」列名（カンマ区切り等で複数値を書く）。 */
const COL_MASTER_COMPETITIVE_SET_COUNTS = '競合セット数';
/** 優先4: マスタに書き込む「穴のセット数提案」列名。 */
const COL_MASTER_HOLE_SET_PROPOSAL = '穴のセット数提案';

// ■ 比較項目の定義 (AI取得項目)
// COMPARE_ITEMS に '参考情報(画像URL)' を追加してください
const COMPARE_ITEMS = [
  '参考情報(画像URL)', // ★ここに追加
  'メインKW(3つ)', '検索KW(商品名用)', '★Amazon検索KW(150字)', 'CTRコピー',
  '楽天Slug(30字)', 'Yahooキャッチ(20-30字)', '楽天キャッチ(60-80字)',
  '商品説明', '箇条書き1', '箇条書き2', '箇条書き3', '箇条書き4', '箇条書き5',
  '市場価格調査',
  '梱包:幅(cm)', '梱包:奥(cm)', '梱包:高(cm)', '梱包:3辺計(cm)', '梱包:重量(g)', '★配送サイズ(タリフ)',
  '★データ根拠',
  '推奨バリ項目名(軸1)', '推奨バリ項目名(軸2)',
  '原材料(食品)', '賞味期限(食品)', '保存方法(食品)', 
  '素材・材質', '商品本体サイズ', 'カラー', 'その他スペック',
  '★推奨楽天ジャンルID', '楽天ジャンル名', 'シリーズ', 'ブランド', '原産国', '総個数', '総重量', '総容量',
  '★推奨YahooカテゴリID', 'Yahooカテゴリ名', '★推奨Yahooブランドコード',
  'Amz:感熱性', 'Amz:一人分数量', 'Amz:一人分単位', 'Amz:ユニット数単位', 'Amz:危険物規制', 'Amz:液体物含有'
];

// ■ 数値として扱いたい列名のリスト
const NUMERIC_COL_NAMES = [
  '商品の販売価格', '在庫数', '総個数', 
  '卸値(税抜)', '卸値(税込)',
  '梱包:幅(cm)', '梱包:奥(cm)', '梱包:高(cm)', '梱包:3辺計(cm)', '梱包:重量(g)',
  '総重量', '総容量',
  'Amz:一人分数量', 'Amz:ユニット数'
];

// ■ 楽天CSV出力時に、2行目以降(SKU行)でも値を出力すべき項目のキー
const SKU_LEVEL_KEYS = [
  'item_url', 'item_number', 'sku_id', 'system_sku_id',
  'sku_var_key_1', 'sku_var_value_1', 'sku_var_key_2', 'sku_var_value_2',
  'price', 'inventory', 'order_limit', 'catalog_id', 'sku_image_path',
  'sub_price', 'sub_first_price', 'dist_price', 'dist_first_price',
  'normal_price', 'display_price', 'lead_time_instock', 'lead_time_outstock',
  'tax_rate', 'point_rate' // 消費税率やポイント変倍率もSKU単位で持つ場合があるため追加
];

// ==========================================
// 1. メニュー作成
// ==========================================
function onOpen() {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    console.warn('[onOpen] getUi() 失敗: ' + (e && e.message));
    return;
  }
  if (!ui) return;

  // 一括出品でエラーが発生していた場合、スプレッドシートを開いたときにポップアップで案内する
  try {
    var pending = PropertiesService.getScriptProperties().getProperty('BATCH_EXPORT_PENDING_ALERT');
    if (pending) {
      PropertiesService.getScriptProperties().deleteProperty('BATCH_EXPORT_PENDING_ALERT');
      ui.alert('【一括出品】エラーが発生していました', pending, ui.ButtonSet.OK);
    }
  } catch (_) {}

  try {
    // AI出品ツールメニュー
    ui.createMenu('AI出品ツール')
      .addItem('一括出品実行...', 'showBatchExportModal')
      .addItem('一括出品トリガーをすべて削除', 'menuDeleteBatchExportTriggers')
      .addSeparator()
      .addItem('1. シート初期化', 'initializeSheetComparison')
      .addItem('2. 全データ一括生成', 'generateListingDataComparison')
      .addSeparator()
      .addItem('3. マスタへ同期', 'syncAiDataToMaster')
      .addSeparator()
      .addSubMenu(ui.createMenu('商品リサーチ')
        .addSubMenu(ui.createMenu('① 仕入れ検討用')
          .addItem('（準備中）要件定義済み・実装は未実施', 'menuResearchProcurementPlaceholder'))
        .addSeparator()
        .addSubMenu(ui.createMenu('② 出品用')
          .addItem('選択行に競合価格を入力', 'menuWriteCompetitivePricesToSelection')
          .addItem('リサーチCSVから競合価格Amazonをマスタに反映', 'menuImportResearchCsvToMaster')
          .addItem('選択行のASINでKeepaから競合価格を取得', 'menuFetchCompetitivePriceFromKeepa')
          .addItem('選択行のJANでYahoo!から競合価格を取得', 'menuFetchCompetitivePriceFromYahoo')
          .addItem('選択行のJANで楽天から競合価格を取得', 'menuFetchCompetitivePriceFromRakuten')
          .addItem('選択行のJANでYahoo!からセット別競合価格を取得', 'menuFetchYahooSetPricesToMaster')
          .addItem('選択行のJANで楽天からセット別競合価格を取得', 'menuFetchRakutenSetPricesToMaster')
          .addItem('Yahoo! セット別価格 取得テスト', 'menuTestYahooSetPrices')
          .addItem('楽天・Yahoo 競合価格APIテスト', 'menuTestRakutenYahooCompetitivePriceApi')
          .addItem('楽天 90件＋セット数・最安テスト', 'menuTestRakuten90ItemsSetCountMinPrice')
          .addItem('楽天 90件スコア一覧（降順・URL付き）', 'menuTestRakuten90ItemsScoreList')
          .addItem('楽天 セット数 Gemini一括判定テスト', 'menuTestRakutenSetCountBatchGemini')
          .addItem('モール横断 セット数統合判定テスト', 'menuTestCrossMallSetCountJudge')
          .addItem('選択行に価格・セット数提案を反映', 'menuProposePriceAndSetToSelection')
          .addSeparator()
          .addItem('ASIN貼り付け（Keepa用）シートを準備', 'menuPrepareAsinPasteSheet')
          .addItem('ASIN貼り付けシートでKeepa取得（20件まで）', 'menuKeepaFetchAsinPasteSheet20')
          .addItem('ASIN貼り付けシートでKeepa取得（50件まで）', 'menuKeepaFetchAsinPasteSheet50')
          .addItem('ASIN貼り付けシートでKeepa取得（全件）', 'menuKeepaFetchAsinPasteSheetAll')
          .addItem('ASIN貼り付けシートをリセット', 'menuResetAsinPasteSheet')
          .addSeparator()
          .addItem('競合・穴のセット数を提案（ASIN貼り付け）', 'menuProposeSetCountFromAsinPasteSheet')
          .addItem('セット構成提案', 'menuSetCompositionProposal')
          .addItem('競合価格のみ修正を反映', 'menuUpdateCompetitivePriceOnly'))
        .addSeparator()
        .addSubMenu(ui.createMenu('③ 販売価格の調整')
          .addItem('Amazonカテゴリーを自動入力（Gemini）', 'menuFillAmazonCategoryByGemini')
          .addSeparator()
          .addItem('販売価格を提案（セット構成提案後に実行）', 'menuProposeSalesPrices')
          .addItem('CPOで価格提案（Gemini）', 'menuCPOProposePrices')
          .addSeparator()
          .addItem('出品CKをクリア（選択行）', 'menuClearShippingCheckForSelection')))
      .addSeparator()
      .addItem('4. 楽天CSV出力 (診断・保存)', 'generateRakutenCSV')
      .addItem('   🗑️ 楽天 商品を削除...', 'showRakutenDeleteSelectionDialog')
      .addItem('   🔗 削除用URLを今のデプロイに設定', 'menuSetRakutenDeleteWebAppUrl')
      .addItem('   🔗 削除用URLを手動で設定', 'menuSetRakutenDeleteWebAppUrlManual')
      .addItem('   ⏹️ 反映確認を停止', 'stopReflectionCheck')
      .addSeparator()
      .addItem('5. 🤖 AI画像仕分けシート作成', 'generateAiImageMatrix')
      .addItem('6. 🚀 リネーム＆アップロード実行', 'executeRenameAndUploadFromMatrix')
      .addSeparator()
      .addItem('計算式を書き出し（表示中シート）', 'menuExportSheetFormulas')
      .addItem('計算式を書き出し（シート・行範囲を指定）', 'menuExportSheetFormulasWithRange')
      .addSeparator()
      .addItem('★システム診断を実行', 'debugSystemCheck')
      .addItem('documentエラー原因を診断（ログで調査）', 'menuDebugDocumentError')
      .addToUi();

    // Yahoo!出品メニュー
    ui.createMenu('🛒 Yahoo!出品')
      .addItem('▶ 出品実行', 'runYahooExport')
      .addSeparator()
      .addItem('🗑️ 商品を削除...', 'showDeleteSelectionDialog')
      .addToUi();
  } catch (e) {
    console.warn('[onOpen] メニュー作成失敗: ' + (e && e.message));
  }
}

// ==========================================
// 1.0.1 計算式を書き出し（表示中シート）
// ==========================================
/**
 * 【メニュー】表示中のシートの全セルの計算式を取得し、新しいシートに「行・列・列名・式」で書き出す。
 * 自動出品シート（セラセン商品登録一括アップロード等）の計算式を共有・整合性チェックする用途。
 * 使い方: 対象シートを開いた状態で「AI出品ツール」→「計算式を書き出し（表示中シート）」を実行する。
 */
function menuExportSheetFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (!sheet) {
    SpreadsheetApp.getUi().alert('シートを開いてから実行してください。');
    return;
  }
  var result = exportSheetFormulas(sheet);
  if (result.error) {
    SpreadsheetApp.getUi().alert('計算式の書き出しに失敗しました', result.error, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '「' + result.sheetName + '」に ' + result.count + ' 件の式を書き出しました。',
    '計算式を書き出し',
    5
  );
}

/**
 * 【メニュー】シート名と行範囲を指定して計算式を書き出す。共有・確認用。
 * シート名はコピペで指定可能。開始行・終了行を指定するとその行だけ書き出す（例: 1〜2行目のみ）。
 */
function menuExportSheetFormulasWithRange() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var currentName = activeSheet ? activeSheet.getName() : '';

  var resp1 = ui.prompt(
    '計算式を書き出し（シート・行範囲を指定）',
    '対象シート名を入力してください。\n（コピペ可。現在のシート: ' + currentName + '）',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp1.getSelectedButton() !== ui.Button.OK) return;
  var sheetName = (resp1.getResponseText() || '').trim();
  if (!sheetName) {
    ui.alert('シート名が空です。');
    return;
  }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ui.alert('シート「' + sheetName + '」が見つかりません。名前を確認してください。');
    return;
  }

  var resp2 = ui.prompt(
    '開始行',
    '何行目から書き出しますか？（1以上の数字。省略時は 1）',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp2.getSelectedButton() !== ui.Button.OK) return;
  var startInput = (resp2.getResponseText() || '1').trim() || '1';
  var startRow = parseInt(startInput, 10);
  if (isNaN(startRow) || startRow < 1) {
    ui.alert('開始行は1以上の数字を入力してください。');
    return;
  }

  var resp3 = ui.prompt(
    '終了行',
    '何行目まで書き出しますか？（数字。空欄＝データの最後まで）',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp3.getSelectedButton() !== ui.Button.OK) return;
  var endInput = (resp3.getResponseText() || '').trim();
  var endRow = null;
  if (endInput !== '') {
    endRow = parseInt(endInput, 10);
    if (isNaN(endRow) || endRow < 1) {
      ui.alert('終了行は1以上の数字を入力するか、空欄にしてください。');
      return;
    }
    if (endRow < startRow) {
      ui.alert('終了行は開始行以上にしてください。');
      return;
    }
  }

  var result = exportSheetFormulas(sheet, startRow, endRow);
  if (result.error) {
    ui.alert('計算式の書き出しに失敗しました: ' + result.error);
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '「' + result.sheetName + '」に ' + result.count + ' 件の式を書き出しました。（' + sheetName + ' の行' + startRow + (endRow != null ? '〜' + endRow : '〜最後') + '）',
    '計算式を書き出し',
    6
  );
}

/**
 * 指定シートのデータ範囲内の計算式を取得し、新しいシートに書き出す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} [startRow1] 書き出す開始行（1-based）。省略時は1行目から。
 * @param {number|null} [endRow1] 書き出す終了行（1-based）。省略時・null はデータの最後まで。
 * @returns {{ sheetName: string, count: number } | { error: string }}
 */
function exportSheetFormulas(sheet, startRow1, endRow1) {
  try {
    var range = sheet.getDataRange();
    if (!range) return { error: 'データ範囲がありません。' };
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    if (numRows === 0 || numCols === 0) return { error: 'データがありません。' };

    var rowStart = 0;
    var rowEnd = numRows - 1;
    if (startRow1 != null && startRow1 !== undefined) {
      rowStart = Math.max(0, parseInt(startRow1, 10) - 1);
      if (rowStart > rowEnd) return { error: '開始行がデータ範囲を超えています。' };
    }
    if (endRow1 != null && endRow1 !== undefined) {
      rowEnd = Math.min(numRows - 1, parseInt(endRow1, 10) - 1);
      if (rowEnd < rowStart) return { error: '終了行が開始行より小さいです。' };
    }

    var formulas = range.getFormulas();
    var values = range.getValues();
    var headers = (numRows > 0) ? values[0] : [];

    var rangeNote = (startRow1 != null)
      ? (startRow1 + '〜' + (endRow1 != null ? endRow1 : '最後'))
      : '全体';
    var rows = [
      ['# 対象シート', sheet.getName(), '行範囲', rangeNote],
      ['行', '列', '列名', '式']
    ];
    var count = 0;
    for (var r = rowStart; r <= rowEnd; r++) {
      for (var c = 0; c < numCols; c++) {
        var formula = (formulas[r] && formulas[r][c]) ? String(formulas[r][c]).trim() : '';
        if (formula !== '' && formula.indexOf('=') === 0) {
          var colLabel = (headers[c] !== undefined && headers[c] !== null && String(headers[c]).trim() !== '')
            ? String(headers[c]).trim()
            : '列' + (c + 1);
          // 先頭に ' を付けて文字列として書き、シート上で式がそのまま表示されるようにする（コピペで式を共有しやすくする）
          rows.push([r + 1, c + 1, colLabel, "'" + formula]);
          count++;
        }
      }
    }

    var outName = '計算式書き出し_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    var outSheet = sheet.getParent().getSheetByName(outName);
    if (outSheet) outSheet.clear(); else outSheet = sheet.getParent().insertSheet(outName);
    outSheet.getRange(1, 1, rows.length, 4).setValues(rows);
    outSheet.getRange(1, 1, 2, 4).setFontWeight('bold');
    outSheet.autoResizeColumns(1, 4);
    return { sheetName: outName, count: count };
  } catch (e) {
    return { error: (e && e.message) ? e.message : String(e) };
  }
}

// ==========================================
// 1.1 一括出品実行（5.6）モーダル
// ==========================================
function showBatchExportModal() {
  // 未表示の一括出品エラーがあれば先にポップアップで表示する（onOpen で見逃した場合の補完）
  try {
    var pending = PropertiesService.getScriptProperties().getProperty('BATCH_EXPORT_PENDING_ALERT');
    if (pending) {
      PropertiesService.getScriptProperties().deleteProperty('BATCH_EXPORT_PENDING_ALERT');
      SpreadsheetApp.getUi().alert('【一括出品】エラーが発生していました', pending, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (_) {}
  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:sans-serif;padding:16px;">' +
    '<h3>一括出品実行</h3>' +
    '<p>コースを選んで「実行」を押してください。予約後すぐにポップアップは閉じ、約1分後に裏で実行されます。完了時にシート上に通知が出ます。</p>' +
    '<form id="f">' +
    '<label><input type="radio" name="course" value="full" checked> フル（楽天 → Yahoo! の順で実行）</label><br>' +
    '<label><input type="radio" name="course" value="rakuten"> 楽天のみ</label><br>' +
    '<label><input type="radio" name="course" value="yahoo"> Yahoo!のみ</label><br>' +
    '</form>' +
    '<p style="margin-top:16px;"><button id="run">実行</button> <button id="cancel">キャンセル</button></p>' +
    '<p id="msg" style="color:#666;font-size:12px;"></p>' +
    '<p style="color:#888;font-size:11px;margin-top:12px;">※「1分ごと」のトリガーは登録しないでください。予約時に「1分後に1回」が自動作成されます。不要なトリガーはメニュー「一括出品トリガーをすべて削除」で削除できます。</p>' +
    '<script>' +
    'document.getElementById("run").onclick=function(){ var c=document.querySelector("input[name=course]:checked").value; document.getElementById("msg").textContent="予約しました。ポップアップを閉じます..."; google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).withFailureHandler(function(e){ document.getElementById("msg").textContent="エラー: "+e.message; }).scheduleBatchExport(c); };' +
    'document.getElementById("cancel").onclick=function(){ google.script.host.close(); };' +
    '</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(260), '一括出品実行');
}

/**
 * 一括出品を「予約」し、すぐ返るのでポップアップは閉じる。
 * 約1分後に1回だけ runBatchExportFromTrigger を実行するトリガーを自動作成する。
 * トリガー作成に失敗した場合（編集者など）は、オーナーが手動で「runBatchExportFromTrigger」を1分後に1回のトリガーを追加する必要がある。
 */
function scheduleBatchExport(course) {
  if (course !== 'full' && course !== 'rakuten' && course !== 'yahoo') throw new Error('不正なコースです: ' + course);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートを取得できません');
  const props = PropertiesService.getScriptProperties();
  const ssId = ss.getId();

  props.setProperty('BATCH_EXPORT_COURSE', course);
  props.setProperty('BATCH_EXPORT_SS_ID', ssId);
  const readBackCourse = props.getProperty('BATCH_EXPORT_COURSE');
  const readBackSsId = props.getProperty('BATCH_EXPORT_SS_ID');
  if (readBackCourse && readBackSsId) {
    console.log('[一括出品] 予約しました（保存OK） course=' + readBackCourse + ' ssId=' + readBackSsId);
  } else {
    console.warn('[一括出品] 予約を保存しましたが読み返しで取得できませんでした course=' + (readBackCourse || '') + ' ssId=' + (readBackSsId || ''));
  }

  // オーナーなら約1分後に1回だけ実行するトリガーを自動作成。編集者は自動作成できないので、オーナーが登録した「1分ごと」トリガーに実行を任せる。
  try {
    var trigger = ScriptApp.newTrigger('runBatchExportFromTrigger')
      .timeBased()
      .after(60 * 1000)  // 1分後
      .create();
    props.setProperty('BATCH_EXPORT_TRIGGER_ID', trigger.getUniqueId());
    console.log('[一括出品] 1分後の実行トリガーを作成しました');
  } catch (e) {
    console.warn('[一括出品] トリガー自動作成をスキップ: ' + e.message);
    ss.toast('予約は保存しました。オーナーが「runBatchExportFromTrigger」を1分ごとに登録していれば、約1分以内に実行されます。', '一括出品', 10);
  }

  ss.toast('予約しました。約1分以内に裏で実行されます。完了時に通知が出ます。', '一括出品', 8);
}

/** 一括出品用トリガーをすべて削除する（自動作成したトリガー＋オーナー登録分を含む） */
function deleteBatchExportTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'runBatchExportFromTrigger'; });
    var n = triggers.length;
    triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
    return n;
  } catch (e) {
    console.warn("一括出品トリガー削除をスキップ（権限不足）: " + e.message);
    return -1;
  }
}

/** メニュー「一括出品トリガーをすべて削除」用。削除後にトーストで結果を表示する */
function menuDeleteBatchExportTriggers() {
  var n = deleteBatchExportTriggers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    if (n >= 0) ss.toast(n > 0 ? '一括出品のトリガーを' + n + '件削除しました。' : '一括出品のトリガーは登録されていませんでした。', 'トリガー削除', 6);
    else ss.toast('トリガー削除に失敗しました（権限不足の可能性）。', 'トリガー削除', 8);
  }
}

/** 自動作成したトリガーだけを削除する（オーナーが登録した「1分ごと」トリガーは残す）。トリガーIDが Script Properties に無い場合は何もしない。 */
function deleteOwnBatchExportTrigger() {
  var props = PropertiesService.getScriptProperties();
  var triggerId = props.getProperty('BATCH_EXPORT_TRIGGER_ID');
  if (!triggerId) return;
  try {
    var triggers = ScriptApp.getProjectTriggers().filter(function(t) {
      return t.getHandlerFunction() === 'runBatchExportFromTrigger' && t.getUniqueId() === triggerId;
    });
    triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
    props.deleteProperty('BATCH_EXPORT_TRIGGER_ID');
    console.log('[一括出品] 自動作成したトリガーを削除しました');
  } catch (e) {
    console.warn("[一括出品] 自動作成トリガー削除をスキップ: " + e.message);
  }
}

/**
 * トリガーから呼ばれる。予約されたコースで楽天・Yahoo を実行し、完了時にトーストを出す。
 * 予約がなければ自動作成トリガーだけ削除して即終了。実行後も自動作成トリガーだけ削除。
 * ※「1分ごと」の手動トリガーは登録しないこと。予約時に「1分後に1回」が自動作成される。
 */
function runBatchExportFromTrigger() {
  var props = PropertiesService.getScriptProperties();
  var course = props.getProperty('BATCH_EXPORT_COURSE');
  var ssId = props.getProperty('BATCH_EXPORT_SS_ID');

  if (!course || !ssId) {
    console.log('[一括出品] 予約データがありません → 自動作成トリガーのみ削除して終了');
    deleteOwnBatchExportTrigger();
    return;
  }

  // 重複実行防止: 既に別の runBatchExportFromTrigger が実行中なら何もせず終了（メモリ・負荷軽減）
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    console.log('[一括出品] 他で実行中のためスキップ（重複実行防止）');
    return;
  }
  try {
    runBatchExportFromTriggerImpl(course, ssId);
  } finally {
    lock.releaseLock();
  }
}

function runBatchExportFromTriggerImpl(course, ssId) {
  console.log('[一括出品] runBatchExportFromTrigger 開始 course=' + course);
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('BATCH_EXPORT_COURSE');
  props.deleteProperty('BATCH_EXPORT_SS_ID');
  console.log('[一括出品] 予約を検出 実行開始');

  var ssForToast;
  try {
    ssForToast = SpreadsheetApp.openById(ssId);
  } catch (e) {
    console.error('[一括出品] スプレッドシートを開けません ssId=' + ssId + ' err=' + e.message);
    deleteOwnBatchExportTrigger();
    return;
  }

  var errMsg = null;
  try {
    runBatchExport(course, true, ssForToast);
    console.log('[一括出品] runBatchExport 完了');
  } catch (e) {
    errMsg = e.message;
    console.error('[一括出品] エラー: ' + errMsg);
    if (e.stack) console.error(e.stack);
  } finally {
    deleteOwnBatchExportTrigger();
  }

  if (errMsg) {
    // トーストはトリガー実行時に失敗することがあるため、先に保存・メール送信を行う
    try {
      PropertiesService.getScriptProperties().setProperty('BATCH_EXPORT_PENDING_ALERT', errMsg);
      console.log('[一括出品] エラー内容を保存しました（スプレッドシートを開くか「一括出品実行」でポップアップ表示）');
    } catch (_) { console.warn('[一括出品] PENDING_ALERT 保存失敗'); }
    var to = PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL');
    if (to) {
      try {
        MailApp.sendEmail(to, '【一括出品】エラーが発生しました', '一括出品の実行中にエラーが発生しました。\n\n' + errMsg + '\n\nスプレッドシートを開くか「一括出品実行」を開くと、該当商品を案内するポップアップが表示されます。');
        console.log('[一括出品] エラー通知メールを送信しました: ' + to);
      } catch (mailErr) {
        console.warn('[一括出品] エラー通知メール送信失敗: ' + mailErr.message);
      }
    } else {
      console.warn('[一括出品] NOTIFICATION_EMAIL が未設定のためエラー通知メールを送信できません');
    }
    try { ssForToast.toast('一括出品でエラーが発生しました: ' + errMsg, '一括出品', 15); } catch (_) {}
  } else {
    try { ssForToast.toast('一括出品が完了しました。', '一括出品', 10); } catch (_) {}
  }
}

/**
 * 一括出品の途中経過をシート「一括出品ログ」に追記する（トーストはタブを離れると見逃すため、事実確認用）
 */
function logBatchExportProgress(ss, message) {
  if (!ss) return;
  try {
    var sheet = ss.getSheetByName(BATCH_EXPORT_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(BATCH_EXPORT_LOG_SHEET);
      sheet.getRange(1, 1, 1, 2).setValues([['日時', '内容']]).setFontWeight('bold').setBackground('#eee');
    }
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var nextRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(nextRow, 1, 1, 2).setValues([[now, message]]);
  } catch (e) { console.warn('[一括出品ログ] 記録失敗: ' + e.message); }
}

/**
 * 一括出品実行の実処理。コースに応じて楽天CSV出力・Yahoo出品を実行する。
 * 途中経過はトースト＋シート「一括出品ログ」に記録（タブを離れても後から確認可能）。
 * @param {string} course - "full" | "rakuten" | "yahoo"
 * @param {boolean} fromTrigger - トリガーから呼ばれた場合 true（確認ダイアログを出さない）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ssOverride] - トリガー実行時は openById で開いたスプレッドシートを渡す
 */
function runBatchExport(course, fromTrigger, ssOverride) {
  var ss = ssOverride || (function() { try { return SpreadsheetApp.getActiveSpreadsheet(); } catch (_) { return null; } })();
  function toast(msg, title, sec) {
    if (ss) try { ss.toast(msg, title || '一括出品', sec != null ? sec : 4); } catch (_) {}
  }

  if (course === 'rakuten') {
    toast('楽天CSV出力を開始しています...', '一括出品', 5);
    logBatchExportProgress(ss, '楽天CSV出力を開始');
    generateRakutenCSV(!!fromTrigger, ssOverride);
    logBatchExportProgress(ss, '楽天CSV出力完了');
    toast('楽天CSV出力が完了しました。', '一括出品', 5);
    return;
  }
  if (course === 'yahoo') {
    toast('Yahoo!出品を開始しています...', '一括出品', 5);
    logBatchExportProgress(ss, 'Yahoo!出品を開始');
    runYahooExport(ssOverride);
    logBatchExportProgress(ss, 'Yahoo!出品完了');
    toast('Yahoo!出品が完了しました。', '一括出品', 5);
    return;
  }
  if (course === 'full') {
    toast('一括出品を開始しました。（楽天CSV出力中...）', '一括出品', 5);
    logBatchExportProgress(ss, '一括出品開始（楽天CSV出力中）');
    var rakutenError = null;
    try {
      generateRakutenCSV(!!fromTrigger, ssOverride);
      logBatchExportProgress(ss, '楽天CSV出力完了。Yahoo!出品に進む');
      toast('楽天CSV出力が完了しました。Yahoo!出品に進みます...', '一括出品', 5);
    } catch (e) {
      rakutenError = e.message;
      console.error('[一括出品] 楽天CSVでエラー（Yahoo出品は続行）: ' + rakutenError);
      logBatchExportProgress(ss, '楽天CSVでエラー（Yahoo続行）: ' + (e.message || ''));
      toast('楽天CSVでエラーがありました。Yahoo!出品を続行します...', '一括出品', 6);
    }
    logBatchExportProgress(ss, 'Yahoo!出品を開始');
    runYahooExport(ssOverride);
    logBatchExportProgress(ss, '一括出品完了');
    if (rakutenError) throw new Error('楽天CSV: ' + rakutenError + '（Yahoo出品は実行済み）');
    return;
  }
  throw new Error('不正なコースです: ' + course);
}

// ==========================================
// 2. シート初期化
// ==========================================
function initializeSheetComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TARGET_SHEET_NAME);

  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert('初期化します。データは削除されますがよろしいですか？', ui.ButtonSet.YES_NO);
    if (res == ui.Button.NO) return;
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(TARGET_SHEET_NAME);
  }

  const inputHeaders = [
    'メーカー名', '仕入判断', 'JANコード', '商品名', 
    '卸値(税抜)', '卸値(税込)', 'JANコードorオリジナルカタログ商品名', 
    'ASINコード', '参考情報(画像URL)', 'プロンプト', 'Gemini_JSON', 'ChatGPT_JSON'
  ];

  let headers = [...inputHeaders];
  COMPARE_ITEMS.forEach(item => {
    headers.push(`${item}[Gemini]`, `${item}[GPT]`, `${item}[正解]`);
  });
  const masterHeaders = COMPARE_ITEMS.map(item => `▼マスタ(${item})`);
  headers = headers.concat(masterHeaders);

  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight('bold').setBorder(true, true, true, true, true, true);
  hRange.setBackground('#d9ead3');
  
  const masterStartCol = headers.length - masterHeaders.length + 1;
  sheet.getRange(1, masterStartCol, 1, masterHeaders.length).setBackground('#c9daf8');

  const inputColCount = inputHeaders.length; 
  for (let i = 0; i < COMPARE_ITEMS.length; i++) {
    const correctColIndex = inputColCount + (i * 3) + 3;
    sheet.getRange(1, correctColIndex).setBackground('#fff2cc'); 
    sheet.getRange(2, correctColIndex, sheet.getMaxRows()-1, 1).setBackground('#fff2cc');
  }

  sheet.setFrozenColumns(4); 
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(4, 200); 
  
  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns());
  dataRange.setVerticalAlignment('top');
  dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(2, 10, sheet.getMaxRows() - 1, 3).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISBLANK($D2),ISBLANK($C2))') 
    .setBackground('#f4c7c3')
    .setRanges([sheet.getRange(2, 1, sheet.getMaxRows()-1, inputColCount)])
    .build());
  sheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getActive().toast('初期化完了。', '完了');
}

// ==========================================
// 3. AIデータ生成 (メイン)
// ==========================================
function generateListingDataComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  
  if (!sheet || !getOpenAiApiKey() || !getGeminiApiKey()) {
    SpreadsheetApp.getUi().alert('シートまたはAPIキーの設定を確認してください。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const startOutputCol = 13; 
  const maxCols = sheet.getMaxColumns();
  if (maxCols >= startOutputCol) {
    sheet.getRange(2, startOutputCol, lastRow - 1, maxCols - 12).clearContent();
  }


/////

// --- 【修正後】（以下のコードに完全に置き換えてください） ---------
  for (let i = 0; i < lastRow - 1; i++) {
    const r = 2 + i;
    const inputMaker = sheet.getRange(r, 1).getValue(); 
    const inputJan = sheet.getRange(r, 3).getValue();
    const inputName = sheet.getRange(r, 4).getValue();
    const inputAsin = sheet.getRange(r, 8).getValue();
    const inputUrl = sheet.getRange(r, 9).getValue(); 
    
    const promptCol = 10;
    const jsonColG = 11;
    const jsonColO = 12;

    if (!inputName && !inputAsin) continue;

    // 1. CSV検索 (変更なし)
    const rakutenCandidates = searchCandidatesWithScore(inputName, inputMaker, RAKUTEN_GENRE_FOLDER_ID, '', false, 1);
    const yahooCatCandidates = searchCandidatesWithScore(inputName, inputMaker, YAHOO_CATEGORY_FOLDER_ID, '', false, 2); 
    const yahooBrandCandidates = searchCandidatesWithScore(inputName, inputMaker, YAHOO_BRAND_FOLDER_ID, '', true, 1);

    // 2. 画像取得 (変更なし)
    let imgBase64 = null, mime = null;
    if (inputUrl && String(inputUrl).startsWith('http')) {
      try {
        const b = UrlFetchApp.fetch(inputUrl).getBlob();
        imgBase64 = Utilities.base64Encode(b.getBytes());
        mime = b.getContentType();
      } catch (e) {}
    }

    // ★追加: Step 1. GeminiによるWeb検索調査（JSONモードOFF）
    SpreadsheetApp.getActive().toast(`${r}行目: Web検索調査中...`, '進行中');
    let searchResultText = "（検索未実行）";
    try {
      // 調査専用のプロンプトを作成
      const searchPrompt = createSearchPrompt(inputName, inputJan, inputMaker);
      // 検索専用APIを呼び出し（結果はテキスト）
      searchResultText = callGeminiSearchAPI(searchPrompt);
    } catch (e) {
      searchResultText = "（検索エラー: " + e.message + "）";
    }

    // ★修正: Step 2. 調査結果を含めて本番プロンプト生成
    const prompt = createFullPrompt(inputName, inputJan, inputAsin, rakutenCandidates, yahooCatCandidates, yahooBrandCandidates, searchResultText);
    sheet.getRange(r, promptCol).setValue(prompt);

    SpreadsheetApp.getActive().toast(`${r}行目: AI生成中(JSON)...`, '進行中');
    
    let gData = {}, oData = {};
    
    // Gemini生成 (JSONモードON)
    try { 
      gData = callGeminiVisionAPI(prompt, imgBase64, mime);
      gData = refineGenreData(gData, RAKUTEN_GENRE_FOLDER_ID, YAHOO_CATEGORY_FOLDER_ID, YAHOO_BRAND_FOLDER_ID);
      sheet.getRange(r, jsonColG).setValue(JSON.stringify(gData)); 
    } catch (e) { 
      sheet.getRange(r, jsonColG).setValue(e.message); 
    }
    
    // OpenAI生成 (検索結果を流用)
    try { 
      oData = callOpenAIVisionAPI(prompt, imgBase64, mime); 
      oData = refineGenreData(oData, RAKUTEN_GENRE_FOLDER_ID, YAHOO_CATEGORY_FOLDER_ID, YAHOO_BRAND_FOLDER_ID);
      sheet.getRange(r, jsonColO).setValue(JSON.stringify(oData)); 
    } catch (e) { 
      sheet.getRange(r, jsonColO).setValue(e.message); 
    }

    writeSplitted(sheet, r, gData, oData, startOutputCol, inputUrl);
  }



// ------------------------------------------------------------
  



  setMasterFormulas(sheet, startOutputCol);
  SpreadsheetApp.getActive().toast('完了しました。', '完了');
}




// ==============================================================================================================================




// ==========================================
// 【AI出品ツール Ver 38.2】 Part 2/3 (属性吸い上げ強化版)
// ==========================================

// ==========================================
// 4. マスタ同期機能 (変更なし)
// ==========================================
function syncAiDataToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aiSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (!aiSheet || !masterSheet) {
    SpreadsheetApp.getUi().alert(`シートが見つかりません。`);
    return;
  }

  const aiData = aiSheet.getDataRange().getValues();
  const masterRange = masterSheet.getDataRange();
  const masterValues = masterRange.getValues();     

  let masterHeaderRowIndex = -1;
  const searchLimit = Math.min(masterValues.length, 20); 
  for (let r = 0; r < searchLimit; r++) {
    if (masterValues[r].includes(ANCHOR_HEADER_NAME)) {
      masterHeaderRowIndex = r;
      break;
    }
  }

  if (masterHeaderRowIndex === -1) {
    SpreadsheetApp.getUi().alert(`ヘッダー「${ANCHOR_HEADER_NAME}」が見つかりません。`);
    return;
  }

  const aiHeaders = aiData[0]; 
  const masterHeaders = masterValues[masterHeaderRowIndex]; 

  const aiColMap = getColumnIndexMap(aiHeaders);
  const masterColMap = getColumnIndexMap(masterHeaders);

  let aiAsinIdx = aiColMap['ASIN'];
  if (aiAsinIdx === undefined) aiAsinIdx = aiColMap['ASINコード'];
  const aiJanIdx = aiColMap['JANコード'];
  const aiNameIdx = aiColMap['商品名'];

  const masterAsinIdx = masterColMap['ASINコード'];
  const masterJanIdx = masterColMap['JANコード'];
  const masterNameIdx = masterColMap['商品名']; 
  const masterChildSkuIdx = masterColMap['子SKU']; 

  const targetColumns = [];
  COMPARE_ITEMS.forEach(item => {
    const colName = `▼マスタ(${item})`;
    if (masterColMap[colName] !== undefined && aiColMap[colName] !== undefined) {
      targetColumns.push({
        name: colName,
        mIdx: masterColMap[colName],
        aiIdx: aiColMap[colName]
      });
    }
  });

  const normalize = (val) => {
    if (val === null || val === undefined) return "";
    return String(val).trim().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  };
  
  const cleanNumber = (val) => {
    if (val === null || val === undefined || val === "") return "";
    let str = String(val).trim();
    str = str.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    str = str.replace(/[^0-9.]/g, '');
    return str;
  };

  const aiMapByAsin = new Map();
  const aiMapByJan = new Map();
  const aiMapByName = new Map();

  for (let i = 1; i < aiData.length; i++) {
    const row = aiData[i];
    const asin = normalize(row[aiAsinIdx]);
    const jan = normalize(row[aiJanIdx]);
    const name = normalize(row[aiNameIdx]);
    
    if (asin) aiMapByAsin.set(asin, row);
    if (jan) aiMapByJan.set(jan, row);
    if (name) aiMapByName.set(name, row);
  }

  let totalUpdatedRows = 0;
  
  targetColumns.forEach(targetCol => {
    const columnValues = masterValues.map(row => [row[targetCol.mIdx]]); 
    let hasUpdate = false;

    for (let i = masterHeaderRowIndex + 1; i < masterValues.length; i++) {
      const row = masterValues[i];
      const mAsin = normalize(row[masterAsinIdx]);
      const mJan = normalize(row[masterJanIdx]);
      const mName = normalize(row[masterNameIdx]);
      const mChildSku = row[masterChildSkuIdx];

      if (String(mChildSku).trim() !== '') continue; 
      if(!mAsin && !mJan && !mName) continue;

      let aiRow = null;
      if (mAsin && aiMapByAsin.has(mAsin)) aiRow = aiMapByAsin.get(mAsin);
      else if (mJan && aiMapByJan.has(mJan)) aiRow = aiMapByJan.get(mJan);
      else if (mName && aiMapByName.has(mName)) aiRow = aiMapByName.get(mName);

      if (aiRow) {
        let val = aiRow[targetCol.aiIdx];
        const rawColName = targetCol.name.replace('▼マスタ(', '').replace(')', '');
        if (NUMERIC_COL_NAMES.includes(rawColName)) val = cleanNumber(val);

        columnValues[i][0] = val; 
        hasUpdate = true;
        if (targetColumns.indexOf(targetCol) === 0) totalUpdatedRows++;
      }
    }

    if (hasUpdate) {
      masterSheet.getRange(1, targetCol.mIdx + 1, columnValues.length, 1).setValues(columnValues);
    }
  });

  SpreadsheetApp.getActive().toast(`同期完了: ${totalUpdatedRows}件更新`, '成功');
}




// ==========================================
// 3.1 商品リサーチ: ①仕入れ検討用 / ②出品用 の分離
// ==========================================
/**
 * ① 仕入れ検討用商品リサーチ: 要件定義済み・実装は未実施である旨を表示する。
 * RESEARCH_AND_ESTIMATE §1.4.0・§1.4.2 参照。実装は今後の開発で対応する。
 */
function menuResearchProcurementPlaceholder() {
  SpreadsheetApp.getUi().alert(
    '① 仕入れ検討用商品リサーチ',
    '仕入れ検討用のリサーチは要件定義済みです。\n実装は今後の開発で対応します。\n（docs RESEARCH_AND_ESTIMATE §1.4.0・§1.4.2）',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==========================================
// 3.1 商品リサーチ（②出品用）: 競合価格のマスタ書き込み
// ==========================================
/**
 * 【② 出品用】選択行に競合価格（Amazon / 楽天 / Yahoo!）を入力してマスタに書き込む。手動入力用。API連携時は writeCompetitivePricesToMaster を直接呼ぶ。
 */
function menuWriteCompetitivePricesToSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) { SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。'); return; }
  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() !== MASTER_SHEET_NAME) { SpreadsheetApp.getUi().alert('マスタシートを開き、競合価格を書き込みたい行を選択してから実行してください。'); return; }
  const range = activeSheet.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('行を選択してから実行してください。'); return; }
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const rowIndices = []; for (let r = 0; r < numRows; r++) rowIndices.push(startRow + r);
  const masterValues = masterSheet.getDataRange().getValues();
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(masterValues.length, 20); r++) { if (masterValues[r].includes(ANCHOR_HEADER_NAME)) { headerRowIdx = r; break; } }
  if (headerRowIdx === -1) { SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。'); return; }
  const masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  const colAmazon = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  const colRakuten = masterColMap[COL_COMPETITIVE_PRICE_RAKUTEN];
  const colYahoo = masterColMap[COL_COMPETITIVE_PRICE_YAHOO];
  if (colAmazon === undefined || colRakuten === undefined || colYahoo === undefined) {
    const missing = []; if (colAmazon === undefined) missing.push(COL_COMPETITIVE_PRICE_AMAZON); if (colRakuten === undefined) missing.push(COL_COMPETITIVE_PRICE_RAKUTEN); if (colYahoo === undefined) missing.push(COL_COMPETITIVE_PRICE_YAHOO);
    SpreadsheetApp.getUi().alert('マスタに次の列がありません: ' + missing.join('、')); return;
  }
  const ui = SpreadsheetApp.getUi();
  const amazonVal = ui.prompt('競合価格 Amazon', '数値または空欄', ui.ButtonSet.OK_CANCEL); if (amazonVal.getSelectedButton() !== ui.Button.OK) return;
  const rakutenVal = ui.prompt('競合価格 楽天', '数値または空欄', ui.ButtonSet.OK_CANCEL); if (rakutenVal.getSelectedButton() !== ui.Button.OK) return;
  const yahooVal = ui.prompt('競合価格 Yahoo!', '数値または空欄', ui.ButtonSet.OK_CANCEL); if (yahooVal.getSelectedButton() !== ui.Button.OK) return;
  const values = { amazon: String(amazonVal.getResponseText()).trim(), rakuten: String(rakutenVal.getResponseText()).trim(), yahoo: String(yahooVal.getResponseText()).trim() };
  writeCompetitivePricesToMaster(ss, masterSheet, headerRowIdx, rowIndices, values);
  SpreadsheetApp.getActive().toast(rowIndices.length + '行に競合価格を反映しました', '完了');
}
/** 指定したマスタ行に競合価格を書き込む。メニュー・API両方から利用可能。 */
function writeCompetitivePricesToMaster(ss, masterSheet, headerRowIdx, rowIndices1Based, values) {
  const masterValues = masterSheet.getDataRange().getValues();
  const masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  const colAmazon = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  const colRakuten = masterColMap[COL_COMPETITIVE_PRICE_RAKUTEN];
  const colYahoo = masterColMap[COL_COMPETITIVE_PRICE_YAHOO];
  if (colAmazon === undefined || colRakuten === undefined || colYahoo === undefined) return;
  for (let i = 0; i < rowIndices1Based.length; i++) {
    const row = rowIndices1Based[i];
    masterSheet.getRange(row, colAmazon + 1).setValue(values.amazon || '');
    masterSheet.getRange(row, colRakuten + 1).setValue(values.rakuten || '');
    masterSheet.getRange(row, colYahoo + 1).setValue(values.yahoo || '');
  }
}

// ----------------------------------------
// 3.1b 商品リサーチ（②出品用）: リサーチCSV（octas形式等）から競合価格Amazonをマスタに反映
// ----------------------------------------
/**
 * 【② 出品用】選択範囲をリサーチCSVとして解釈し、ASINと価格列でマスタの該当行に競合価格 Amazon を書き込む。
 * 1行目をヘッダーとし、「ASIN」列と「最安価格」または「BuyBox価格」列を自動検出。マスタは ASINコード で突合。
 */
function menuImportResearchCsvToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) { SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。'); return; }
  const range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('リサーチCSVの範囲（ヘッダー含む）を選択してから実行してください。'); return; }
  const data = range.getValues();
  if (data.length < 2) { SpreadsheetApp.getUi().alert('選択範囲は2行以上必要です（1行目=ヘッダー、2行目以降=データ）。'); return; }
  const headers = data[0].map(function(c) { return String(c || '').trim(); });
  let asinCol = -1, priceCol = -1;
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    if (h === 'ASIN' || h.indexOf('ASIN') >= 0) asinCol = c;
    if (h === '最安価格' || h === 'BuyBox価格' || h.indexOf('最安価格') >= 0 || h.indexOf('BuyBox価格') >= 0) priceCol = c;
  }
  if (priceCol === -1) { for (let c = 0; c < headers.length; c++) { if (headers[c].indexOf('Current価格') >= 0) { priceCol = c; break; } } }
  if (asinCol === -1 || priceCol === -1) { SpreadsheetApp.getUi().alert('ヘッダーに「ASIN」列と「最安価格」または「BuyBox価格」列が見つかりません。'); return; }
  const asinToPrice = {};
  const parseNum = function(v) { if (v === undefined || v === null || v === '') return NaN; const n = Number(String(v).replace(/,/g, '')); return isNaN(n) ? NaN : n; };
  for (let r = 1; r < data.length; r++) {
    const asin = String(data[r][asinCol] || '').trim();
    if (!asin) continue;
    const price = parseNum(data[r][priceCol]);
    if (!isNaN(price)) asinToPrice[asin] = price;
  }
  const masterValues = masterSheet.getDataRange().getValues();
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(masterValues.length, 20); r++) { if (masterValues[r].includes(ANCHOR_HEADER_NAME)) { headerRowIdx = r; break; } }
  if (headerRowIdx === -1) { SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。'); return; }
  const masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  const colAsin = masterColMap['ASINコード'];
  const colAmazon = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  if (colAsin === undefined || colAmazon === undefined) { SpreadsheetApp.getUi().alert('マスタに ASINコード または ' + COL_COMPETITIVE_PRICE_AMAZON + ' 列がありません。'); return; }
  let updated = 0;
  for (let r = headerRowIdx + 1; r < masterValues.length; r++) {
    const asin = String(masterValues[r][colAsin] || '').trim();
    if (!asin || asinToPrice[asin] === undefined) continue;
    masterSheet.getRange(r + 1, colAmazon + 1).setValue(asinToPrice[asin]);
    updated++;
  }
  SpreadsheetApp.getActive().toast('競合価格 Amazon を ' + updated + ' 行に反映しました', '完了');
}

// ----------------------------------------
// 3.1c 商品リサーチ（②出品用）: Keepa API で競合価格（送料込み最安値）を取得してマスタに反映
// ----------------------------------------
/** Keepa API のドメイン: 5 = Amazon.co.jp */
const KEEPA_DOMAIN_JP = 5;
/** Keepa の csv 配列: 10=Buy Box(カート価格), 0=Amazon, 7=NEW_FBM等。競合価格はカート価格を最優先で取得 */
const KEEPA_INDEX_BUY_BOX = 10;
const KEEPA_INDEX_NEW_FBM_SHIPPING = 7;
var KEEPA_PRICE_CSV_INDICES = [10, 0, 7, 1, 2];

/**
 * 【② 出品用】選択行の ASIN を使って Keepa API で価格を取得し、競合価格 Amazon に書き込む。
 * Script Properties に KEEPA_API_KEY を設定してください。トークン制限のため、選択行は少なめ（目安20行以内）を推奨。
 */
function menuFetchCompetitivePriceFromKeepa() {
  const key = PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY');
  if (!key || key.trim() === '') { SpreadsheetApp.getUi().alert('Script Properties に KEEPA_API_KEY を設定してください。'); return; }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) { SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。'); return; }
  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() !== MASTER_SHEET_NAME) { SpreadsheetApp.getUi().alert('マスタシートを開き、対象行を選択してから実行してください。'); return; }
  const range = activeSheet.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('行を選択してから実行してください。'); return; }
  const masterValues = masterSheet.getDataRange().getValues();
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(masterValues.length, 20); r++) { if (masterValues[r].includes(ANCHOR_HEADER_NAME)) { headerRowIdx = r; break; } }
  if (headerRowIdx === -1) { SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。'); return; }
  const masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  const colAsin = masterColMap['ASINコード'];
  if (colAsin === undefined) { SpreadsheetApp.getUi().alert('マスタに ASINコード 列がありません。'); return; }
  const startRow = range.getRow(), numRows = range.getNumRows();
  const asinToRow = {};
  for (let i = 0; i < numRows; i++) {
    const rowIdx = startRow + i;
    const dataIdx = rowIdx - 1;
    if (dataIdx <= headerRowIdx) continue;
    const asin = String(masterValues[dataIdx][colAsin] || '').trim();
    if (asin) asinToRow[asin] = rowIdx;
  }
  const asins = Object.keys(asinToRow);
  if (asins.length === 0) { SpreadsheetApp.getUi().alert('選択行に ASIN がありません。'); return; }
  SpreadsheetApp.getActive().toast('Keepa に問い合わせ中...', '処理中');
  const priceMap = fetchCompetitivePricesFromKeepa(key, asins);
  if (!priceMap) { SpreadsheetApp.getUi().alert('Keepa API の取得に失敗しました。キー・ネットワーク・トークン残数を確認してください。'); return; }
  const colAmazon = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  if (colAmazon === undefined) { SpreadsheetApp.getUi().alert('マスタに ' + COL_COMPETITIVE_PRICE_AMAZON + ' 列がありません。'); return; }
  let updated = 0;
  asins.forEach(function(asin) {
    const price = priceMap[asin];
      if (price != null && asinToRow[asin] != null) {
      var actualPrice = Math.round(Number(price) / 100);
      masterSheet.getRange(asinToRow[asin], colAmazon + 1).setValue(actualPrice);
      updated++;
    }
  });
  SpreadsheetApp.getActive().toast('競合価格 Amazon を ' + updated + ' 行に反映しました', '完了');
}

/**
 * 【② 出品用】選択行の JAN を使って Yahoo! ショッピング API で価格を取得し、競合価格 Yahoo! に書き込む。
 * Script Properties に YAHOO_SHOPPING_CLIENT_ID を設定してください。同一 JAN は 1 回だけ API 呼び出しする。
 */
function menuFetchCompetitivePriceFromYahoo() {
  var clientId = (PropertiesService.getScriptProperties().getProperty('YAHOO_SHOPPING_CLIENT_ID') || '').trim();
  if (!clientId) {
    SpreadsheetApp.getUi().alert('Script Properties に YAHOO_SHOPPING_CLIENT_ID を設定してください。');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。');
    return;
  }
  var activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() !== MASTER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('マスタシートを開き、対象行を選択してから実行してください。');
    return;
  }
  var range = activeSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してから実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(masterValues.length, 20); r++) {
    if (masterValues[r].indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colYahoo = masterColMap[COL_COMPETITIVE_PRICE_YAHOO];
  if (colJan === undefined) {
    SpreadsheetApp.getUi().alert('マスタに JANコード 列がありません。');
    return;
  }
  if (colYahoo === undefined) {
    SpreadsheetApp.getUi().alert('マスタに ' + COL_COMPETITIVE_PRICE_YAHOO + ' 列がありません。');
    return;
  }
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var janToRows = {};
  for (var i = 0; i < numRows; i++) {
    var rowIdx = startRow + i;
    var dataIdx = rowIdx - 1;
    if (dataIdx <= headerRowIdx) continue;
    var jan = String(masterValues[dataIdx][colJan] || '').trim();
    if (jan.length >= 8) {
      if (!janToRows[jan]) janToRows[jan] = [];
      janToRows[jan].push(rowIdx);
    }
  }
  var jans = Object.keys(janToRows);
  if (jans.length === 0) {
    SpreadsheetApp.getUi().alert('選択行に有効な JAN がありません。');
    return;
  }
  SpreadsheetApp.getActive().toast('Yahoo! API に問い合わせ中...', '処理中');
  var updated = 0;
  for (var j = 0; j < jans.length; j++) {
    var jan = jans[j];
    var res = fetchYahooShoppingItemPriceByJan(clientId, jan);
    var price = (res.error || res.price == null) ? '' : res.price;
    var rows = janToRows[jan];
    for (var k = 0; k < rows.length; k++) {
      masterSheet.getRange(rows[k], colYahoo + 1).setValue(price);
      if (price !== '') updated++;
    }
    if (j < jans.length - 1) Utilities.sleep(300);
  }
  SpreadsheetApp.getActive().toast('競合価格 Yahoo! を ' + updated + ' 行に反映しました', '完了');
}

/**
 * 【② 出品用】選択行の JAN を使って楽天市場 API で価格を取得し、競合価格楽天に書き込む。
 * Script Properties に RAKUTEN_APP_ID・RAKUTEN_ACCESS_KEY を設定してください。同一 JAN は 1 回だけ API 呼び出しする。
 * 楽天 API が 403 等で未対応の場合は価格は空のまま。サポート対応後に利用可能になる想定。
 */
function menuFetchCompetitivePriceFromRakuten() {
  var props = PropertiesService.getScriptProperties();
  var appId = (props.getProperty('RAKUTEN_APP_ID') || '').trim();
  var accessKey = (props.getProperty('RAKUTEN_ACCESS_KEY') || '').trim();
  if (!appId || !accessKey) {
    SpreadsheetApp.getUi().alert('Script Properties に RAKUTEN_APP_ID と RAKUTEN_ACCESS_KEY を設定してください。');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。');
    return;
  }
  var activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() !== MASTER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('マスタシートを開き、対象行を選択してから実行してください。');
    return;
  }
  var range = activeSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してから実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(masterValues.length, 20); r++) {
    if (masterValues[r].indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colRakuten = masterColMap[COL_COMPETITIVE_PRICE_RAKUTEN];
  var colMaker = masterColMap['メーカー名'];
  var colName = masterColMap['商品名ベース'] != null ? masterColMap['商品名ベース'] : masterColMap['商品名'];
  if (colJan === undefined) {
    SpreadsheetApp.getUi().alert('マスタに JANコード 列がありません。');
    return;
  }
  if (colRakuten === undefined) {
    SpreadsheetApp.getUi().alert('マスタに ' + COL_COMPETITIVE_PRICE_RAKUTEN + ' 列がありません。');
    return;
  }
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var janToRows = {};
  for (var i = 0; i < numRows; i++) {
    var rowIdx = startRow + i;
    var dataIdx = rowIdx - 1;
    if (dataIdx <= headerRowIdx) continue;
    var jan = String(masterValues[dataIdx][colJan] || '').trim();
    if (jan.length >= 8) {
      if (!janToRows[jan]) janToRows[jan] = [];
      janToRows[jan].push(rowIdx);
    }
  }
  var jans = Object.keys(janToRows);
  if (jans.length === 0) {
    SpreadsheetApp.getUi().alert('選択行に有効な JAN がありません。');
    return;
  }
  Logger.log('[楽天競合価格] 選択行: ' + numRows + ' 行、対象JAN数: ' + jans.length + '（' + jans.slice(0, 5).join(', ') + (jans.length > 5 ? ' ...' : '') + '）');
  SpreadsheetApp.getActive().toast('楽天 API に問い合わせ中...', '処理中');
  var updated = 0;
  for (var j = 0; j < jans.length; j++) {
    var jan = jans[j];
    var res = fetchRakutenIchibaItemPrice(appId, accessKey, jan);
    Logger.log('[楽天競合価格] Test1 JAN=' + jan + ' (field=0 Broad): rawCount=' + (res.rawCount || 0) + (res.price != null ? ' price=' + res.price : '') + (res.error ? ' error=' + res.error : ''));
    if ((res.rawCount === 0 || res.error) && (colMaker !== undefined || colName !== undefined)) {
      var firstRowIdx = janToRows[jan][0];
      var firstDataIdx = firstRowIdx - 1;
      var maker = (colMaker !== undefined) ? String(masterValues[firstDataIdx][colMaker] || '').trim() : '';
      var name = (colName !== undefined) ? String(masterValues[firstDataIdx][colName] || '').trim() : '';
      var fallbackKeyword = (maker + ' ' + name).trim();
      if (fallbackKeyword.length >= 2) {
        Utilities.sleep(300);
        res = fetchRakutenIchibaItemPrice(appId, accessKey, fallbackKeyword);
        Logger.log('[楽天競合価格] Test2 メーカー+商品名: "' + fallbackKeyword.substring(0, 40) + (fallbackKeyword.length > 40 ? '...' : '') + '" → rawCount=' + (res.rawCount || 0) + (res.price != null ? ' price=' + res.price : ''));
      }
      if ((res.rawCount === 0 || res.error) && name.length >= 2) {
        Utilities.sleep(300);
        res = fetchRakutenIchibaItemPrice(appId, accessKey, name);
        Logger.log('[楽天競合価格] Test3 商品名のみ: "' + name.substring(0, 40) + (name.length > 40 ? '...' : '') + '" → rawCount=' + (res.rawCount || 0) + (res.price != null ? ' price=' + res.price : ''));
      }
      if ((res.rawCount === 0 || res.error) && fallbackKeyword.length >= 2) {
        Utilities.sleep(300);
        res = fetchRakutenIchibaItemPrice(appId, accessKey, fallbackKeyword, { orFlag: 1 });
        Logger.log('[楽天競合価格] Test4 メーカー+商品名(orFlag=1): "' + fallbackKeyword.substring(0, 40) + (fallbackKeyword.length > 40 ? '...' : '') + '" → rawCount=' + (res.rawCount || 0) + (res.price != null ? ' price=' + res.price : ''));
      }
    }
    var price = (res.error || res.price == null) ? '' : res.price;
    var rows = janToRows[jan];
    for (var k = 0; k < rows.length; k++) {
      masterSheet.getRange(rows[k], colRakuten + 1).setValue(price);
      if (price !== '') updated++;
    }
    if (j < jans.length - 1) Utilities.sleep(300);
  }
  Logger.log('[楽天競合価格] 完了: 価格を反映した行数=' + updated);
  SpreadsheetApp.getActive().toast('競合価格楽天 を ' + updated + ' 行に反映しました', '完了');
}

/** セット別価格で「要確認」とする閾値: 1個あたり換算が競合価格Amazonのこの倍率未満なら要確認。 */
var YAHOO_SET_PRICE_REVIEW_RATIO = 0.3;

/**
 * 【② 出品用】選択行の JAN で Yahoo! API を複数件取得し、セット数ごとの最安価格・URL を 競合価格Yahoo!・競合URLYahoo! に書き込む。
 * 1個あたりが競合価格Amazonの YAHOO_SET_PRICE_REVIEW_RATIO 未満のセットは要確認とし、要確認Yahoo!・確認内容メモYahoo! に記載する。
 * 要確認列は選択式（要確認／確認OK）。Script Properties: YAHOO_SHOPPING_CLIENT_ID。
 */
function menuFetchYahooSetPricesToMaster() {
  var clientId = (PropertiesService.getScriptProperties().getProperty('YAHOO_SHOPPING_CLIENT_ID') || '').trim();
  if (!clientId) {
    SpreadsheetApp.getUi().alert('Script Properties に YAHOO_SHOPPING_CLIENT_ID を設定してください。');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。');
    return;
  }
  var activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() !== MASTER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('マスタシートを開き、対象行を選択してから実行してください。');
    return;
  }
  var range = activeSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してから実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(masterValues.length, 20); r++) {
    if (masterValues[r].indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colYahoo = masterColMap[COL_COMPETITIVE_PRICE_YAHOO];
  var colUrl = masterColMap[COL_COMPETITOR_URL_YAHOO];
  var colReview = masterColMap[COL_REVIEW_STATUS_YAHOO];
  var colMemo = masterColMap[COL_REVIEW_MEMO_YAHOO];
  var colAmazon = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  var colSetQty = masterColMap[COL_MASTER_TOTAL_QTY];
  if (colJan === undefined) {
    SpreadsheetApp.getUi().alert('マスタに JANコード 列がありません。');
    return;
  }
  if (colYahoo === undefined || colUrl === undefined || colReview === undefined || colMemo === undefined) {
    SpreadsheetApp.getUi().alert('マスタに 競合価格Yahoo!・競合URLYahoo!・要確認Yahoo!・確認内容メモYahoo! のいずれかがありません。');
    return;
  }
  if (colSetQty === undefined) {
    SpreadsheetApp.getUi().alert('マスタに ' + COL_MASTER_TOTAL_QTY + ' 列がありません。行ごとのセット数に合わせた最安値を書き込むため必要です。');
    return;
  }
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var janToRowInfos = {};
  for (var i = 0; i < numRows; i++) {
    var rowIdx = startRow + i;
    var dataIdx = rowIdx - 1;
    if (dataIdx <= headerRowIdx) continue;
    var jan = String(masterValues[dataIdx][colJan] || '').trim();
    if (jan.length >= 8) {
      var amazonVal = masterValues[dataIdx][colAmazon];
      var amazonNum = (amazonVal !== '' && amazonVal != null) ? (typeof amazonVal === 'number' ? amazonVal : parseInt(String(amazonVal).replace(/[^0-9]/g, ''), 10)) : 0;
      if (isNaN(amazonNum)) amazonNum = 0;
      var setQtyVal = masterValues[dataIdx][colSetQty];
      var setQty = (setQtyVal !== '' && setQtyVal != null) ? (typeof setQtyVal === 'number' ? setQtyVal : parseInt(String(setQtyVal).replace(/[^0-9]/g, ''), 10)) : null;
      if (isNaN(setQty) || setQty < 1) setQty = null;
      if (!janToRowInfos[jan]) janToRowInfos[jan] = [];
      janToRowInfos[jan].push({ row: rowIdx, amazon: amazonNum, setQty: setQty });
    }
  }
  var jans = Object.keys(janToRowInfos);
  if (jans.length === 0) {
    SpreadsheetApp.getUi().alert('選択行に有効な JAN がありません。');
    return;
  }
  SpreadsheetApp.getActive().toast('Yahoo! API に問い合わせ中...', '処理中');
  var written = 0;
  for (var j = 0; j < jans.length; j++) {
    var jan = jans[j];
    var res = fetchYahooShoppingItemsByJan(clientId, jan, 20);
    if (res.error) {
      Logger.log('[Yahoo!セット別] JAN=' + jan + ' エラー: ' + res.error);
      continue;
    }
    var hits = res.hits || [];
    var setToMinPrice = {};
    var setToUrl = {};
    for (var hi = 0; hi < hits.length; hi++) {
      var name = hits[hi].name;
      var price = hits[hi].price;
      var itemUrl = hits[hi].url || '';
      var setCount = parseSetCountFromItemName(name);
      if (setCount != null && price != null && price > 0) {
        if (setToMinPrice[setCount] == null || price < setToMinPrice[setCount]) {
          setToMinPrice[setCount] = price;
          setToUrl[setCount] = itemUrl;
        }
      }
    }
    var rowInfos = janToRowInfos[jan];
    for (var ri = 0; ri < rowInfos.length; ri++) {
      var info = rowInfos[ri];
      var row = info.row;
      var amazon = info.amazon;
      var setQty = info.setQty;
      var priceString = '';
      var urlString = '';
      var needReview = false;
      var memoText = '';
      if (setQty != null && setToMinPrice[setQty] != null) {
        var price = setToMinPrice[setQty];
        priceString = String(price);
        if (setToUrl[setQty]) urlString = setToUrl[setQty];
        var perUnit = price / setQty;
        if (amazon > 0 && perUnit < amazon * YAHOO_SET_PRICE_REVIEW_RATIO) {
          needReview = true;
          memoText = setQty + '→' + price + '円: 1個あたりがAmazon競合価格比で異常に安い';
        }
      }
      masterSheet.getRange(row, colYahoo + 1).setValue(priceString);
      masterSheet.getRange(row, colUrl + 1).setValue(urlString);
      masterSheet.getRange(row, colReview + 1).setValue(needReview ? '要確認' : '');
      masterSheet.getRange(row, colMemo + 1).setValue(memoText);
      written++;
    }
    if (j < jans.length - 1) Utilities.sleep(300);
  }
  Logger.log('[Yahoo!セット別] 入力規則は設定しない（人間設定を維持）');
  SpreadsheetApp.getActive().toast('Yahoo! セット別競合価格を ' + written + ' 行に反映しました', '完了');
}

/**
 * 【② 出品用】選択行の JAN で楽天 Ichiba API を呼び、セット数ごとの最安値（総量＋ポイント込み）を 競合価格楽天・競合URL楽天・要確認楽天・確認内容メモ楽天 に書き込む。
 * 既存の menuFetchCompetitivePriceFromRakuten は変更しない。切り戻し時はこのメニューを外すこと。docs/RAKUTEN_YAHOO_COMPETITIVE_PRICE_REQUIREMENTS.md 参照。
 * ログは [楽天セット別] で統一し、調査・改善に利用する。
 */
function menuFetchRakutenSetPricesToMaster() {
  var props = PropertiesService.getScriptProperties();
  var appId = (props.getProperty('RAKUTEN_APP_ID') || '').trim();
  var accessKey = (props.getProperty('RAKUTEN_ACCESS_KEY') || '').trim();
  if (!appId || !accessKey) {
    SpreadsheetApp.getUi().alert('Script Properties に RAKUTEN_APP_ID と RAKUTEN_ACCESS_KEY を設定してください。');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。');
    return;
  }
  var activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() !== MASTER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('マスタシートを開き、対象行を選択してから実行してください。');
    return;
  }
  var range = activeSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してから実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(masterValues.length, 20); r++) {
    if (masterValues[r].indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colRakuten = masterColMap[COL_COMPETITIVE_PRICE_RAKUTEN];
  var colUrlRakuten = masterColMap[COL_COMPETITOR_URL_RAKUTEN];
  var colReviewRakuten = masterColMap[COL_REVIEW_STATUS_RAKUTEN];
  var colMemoRakuten = masterColMap[COL_REVIEW_MEMO_RAKUTEN];
  var colAmazon = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  var colSetQty = masterColMap[COL_MASTER_TOTAL_QTY];
  var colMaker = masterColMap['メーカー名'];
  var colName = masterColMap['商品名ベース'] != null ? masterColMap['商品名ベース'] : masterColMap['商品名'];
  if (colJan === undefined) {
    SpreadsheetApp.getUi().alert('マスタに JANコード 列がありません。');
    return;
  }
  if (colRakuten === undefined || colUrlRakuten === undefined || colReviewRakuten === undefined || colMemoRakuten === undefined) {
    SpreadsheetApp.getUi().alert('マスタに 競合価格楽天・競合URL楽天・要確認楽天・確認内容メモ楽天 のいずれかがありません。');
    return;
  }
  if (colSetQty === undefined) {
    SpreadsheetApp.getUi().alert('マスタに ' + COL_MASTER_TOTAL_QTY + ' 列がありません。行ごとのセット数に合わせた最安値を書き込むため必要です。');
    return;
  }
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var janToRowInfos = {};
  for (var i = 0; i < numRows; i++) {
    var rowIdx = startRow + i;
    var dataIdx = rowIdx - 1;
    if (dataIdx <= headerRowIdx) continue;
    var jan = String(masterValues[dataIdx][colJan] || '').trim();
    if (jan.length >= 8) {
      var amazonVal = masterValues[dataIdx][colAmazon];
      var amazonNum = (amazonVal !== '' && amazonVal != null) ? (typeof amazonVal === 'number' ? amazonVal : parseInt(String(amazonVal).replace(/[^0-9]/g, ''), 10)) : 0;
      if (isNaN(amazonNum)) amazonNum = 0;
      var setQtyVal = masterValues[dataIdx][colSetQty];
      var setQty = (setQtyVal !== '' && setQtyVal != null) ? (typeof setQtyVal === 'number' ? setQtyVal : parseInt(String(setQtyVal).replace(/[^0-9]/g, ''), 10)) : null;
      if (isNaN(setQty) || setQty < 1) setQty = null;
      if (!janToRowInfos[jan]) janToRowInfos[jan] = [];
      janToRowInfos[jan].push({ row: rowIdx, amazon: amazonNum, setQty: setQty });
    }
  }
  var jans = Object.keys(janToRowInfos);
  if (jans.length === 0) {
    SpreadsheetApp.getUi().alert('選択行に有効な JAN がありません。');
    return;
  }
  Logger.log('[楽天セット別] 開始 選択行=' + numRows + ' JAN種類=' + jans.length + ' JAN例=' + (jans[0] || ''));
  SpreadsheetApp.getActive().toast('楽天 API に問い合わせ中...', '楽天セット別', 5);
  var written = 0;
  for (var j = 0; j < jans.length; j++) {
    var jan = jans[j];
    var rowInfos = janToRowInfos[jan];
    var firstDataIdx = rowInfos[0].row - 1;
    var maker = (colMaker !== undefined) ? String(masterValues[firstDataIdx][colMaker] || '').trim() : '';
    var name = (colName !== undefined) ? String(masterValues[firstDataIdx][colName] || '').trim() : '';
    var nameForScore = name;
    if (name.length > 0 && getGeminiApiKey()) {
      var parsed = parseProductNameByGemini(name);
      if (parsed && parsed.product && String(parsed.product).trim().length > 0) {
        nameForScore = String(parsed.product).trim();
        Logger.log('[楽天商品名一致] スコア用期待名(コア)="' + nameForScore + '" 元="' + name.substring(0, 40) + (name.length > 40 ? '...' : '') + '"');
      }
      Utilities.sleep(350);
    }
    var fallbackKeyword = (maker + ' ' + name).trim();
    var res = fetchRakutenIchibaItems(appId, accessKey, jan);
    var keywordUsed = jan;
    if (res.error && res.items.length === 0) {
      Logger.log('[楽天セット別] JAN=' + jan + ' 検索エラー: ' + res.error);
    }
    if (res.items.length === 0 && fallbackKeyword.length >= 2) {
      Utilities.sleep(300);
      res = fetchRakutenIchibaItems(appId, accessKey, fallbackKeyword);
      keywordUsed = fallbackKeyword;
      Logger.log('[楽天セット別] JAN=' + jan + ' フォールバック keyword="' + fallbackKeyword.substring(0, 40) + (fallbackKeyword.length > 40 ? '...' : '') + '" 取得件数=' + res.items.length);
    }
    var allItems = (res.items || []).slice();
    for (var page = 2; page <= 3; page++) {
      var nextRes = fetchRakutenIchibaItems(appId, accessKey, keywordUsed, { hits: 30, page: page });
      if (nextRes.items && nextRes.items.length > 0) {
        for (var ni = 0; ni < nextRes.items.length; ni++) allItems.push(nextRes.items[ni]);
      }
      if (page < 3) Utilities.sleep(300);
    }
    Logger.log('[楽天セット別] JAN=' + jan + ' 3ページ合計取得=' + allItems.length + ' 件');
    var th = (typeof RAKUTEN_NAME_MATCH_THRESHOLD !== 'undefined') ? RAKUTEN_NAME_MATCH_THRESHOLD : 50;
    var activeRules = [];
    if (RAKUTEN_NAME_MATCH_USE_FULL) activeRules.push('完全一致');
    if (RAKUTEN_NAME_MATCH_USE_WORD) activeRules.push('単語一致');
    if (RAKUTEN_NAME_MATCH_USE_MAKER) activeRules.push('メーカー一致');
    if (RAKUTEN_NAME_MATCH_USE_CHAR) activeRules.push('文字一致');
    Logger.log('[楽天商品名一致] 閾値=' + th + ' 有効ルール=' + (activeRules.length ? activeRules.join(',') : 'なし'));
    var applyNameFilter = (maker.length > 0 || nameForScore.length > 0);
    if (!applyNameFilter) Logger.log('[楽天商品名一致] 期待メーカー・商品名が両方空のためスコアフィルタなし（全件対象）');
    var nameMatchAdopted = 0;
    var nameMatchExcluded = 0;
    var setToMinPrice = {};
    var setToUrl = {};
    for (var ii = 0; ii < allItems.length; ii++) {
      var it = allItems[ii];
      if (applyNameFilter) {
        var nameMatch = evalRakutenItemNameMatchScore(maker, nameForScore, it.itemName);
        if (nameMatch.score < th) {
          nameMatchExcluded++;
          var ruleStr = nameMatch.usedRules.map(function (r) { return r.rule + ':' + r.score; }).join(',');
          Logger.log('[楽天商品名一致] 除外 score=' + nameMatch.score + ' ルール=' + (ruleStr || 'なし') + ' itemName=' + (it.itemName ? it.itemName.substring(0, 40) : ''));
          continue;
        }
        nameMatchAdopted++;
      }
      var parsedSet = parseSetCountFromItemNameWithSource(it.itemName);
      var setCount = parsedSet ? parsedSet.setCount : null;
      if (parsedSet && parsedSet.fromP && setCount != null && getGeminiApiKey()) {
        var geminiSet = inferSetCountFromItemNameByGemini(it.itemName);
        if (geminiSet != null && geminiSet >= 1) setCount = geminiSet; else setCount = null;
        if (setCount == null) Utilities.sleep(350);
      }
      if (setCount == null) {
        Logger.log('[楽天セット別] セット数不明 itemName先頭50字=' + (it.itemName ? it.itemName.substring(0, 50) : ''));
        continue;
      }
      var effectivePrice = it.itemPrice - Math.round((it.itemPrice * (it.pointRate || 0)) / 100);
      if (it.postageFlag === 1) {
        Logger.log('[楽天セット別] 送料あり セット数=' + setCount + ' itemPrice=' + it.itemPrice + ' ポイント還元後=' + effectivePrice);
      }
      if (setToMinPrice[setCount] == null || effectivePrice < setToMinPrice[setCount]) {
        setToMinPrice[setCount] = effectivePrice;
        setToUrl[setCount] = it.itemUrl;
        if (applyNameFilter && nameMatch && nameMatch.usedRules.length > 0) {
          var ruleStrAdopt = nameMatch.usedRules.map(function (r) { return r.rule + ':' + r.score; }).join(',');
          Logger.log('[楽天商品名一致] 採用 セット数=' + setCount + ' score=' + nameMatch.score + ' ルール=' + ruleStrAdopt);
        }
      }
    }
    Logger.log('[楽天商品名一致] JAN=' + jan + ' 取得=' + allItems.length + ' 閾値以上採用=' + nameMatchAdopted + ' 除外=' + nameMatchExcluded);
    var setCountsFound = Object.keys(setToMinPrice).map(function (k) { return k; });
    Logger.log('[楽天セット別] JAN=' + jan + ' セット数別最安 セット数=' + setCountsFound.join(',') + ' 件数=' + setCountsFound.length);
    for (var ri = 0; ri < rowInfos.length; ri++) {
      var info = rowInfos[ri];
      var row = info.row;
      var amazon = info.amazon;
      var setQty = info.setQty;
      var priceString = '';
      var urlString = '';
      var needReview = false;
      var memoText = '';
      if (setQty != null && setToMinPrice[setQty] != null) {
        var price = setToMinPrice[setQty];
        priceString = String(price);
        if (setToUrl[setQty]) urlString = setToUrl[setQty];
        var perUnit = price / setQty;
        if (amazon > 0 && perUnit < amazon * YAHOO_SET_PRICE_REVIEW_RATIO) {
          needReview = true;
          memoText = setQty + '→' + price + '円: 1個あたりがAmazon競合価格比で異常に安い';
        }
        Logger.log('[楽天セット別] 行' + row + ' セット数=' + setQty + ' 価格=' + price + ' URL=' + (urlString ? 'あり' : 'なし') + (needReview ? ' 要確認' : ''));
      } else {
        Logger.log('[楽天セット別] 行' + row + ' セット数=' + (setQty != null ? setQty : '未設定') + ' 該当なし');
      }
      masterSheet.getRange(row, colRakuten + 1).setValue(priceString);
      masterSheet.getRange(row, colUrlRakuten + 1).setValue(urlString);
      masterSheet.getRange(row, colReviewRakuten + 1).setValue(needReview ? '要確認' : '');
      masterSheet.getRange(row, colMemoRakuten + 1).setValue(memoText);
      written++;
    }
    if (j < jans.length - 1) Utilities.sleep(300);
  }
  Logger.log('[楽天セット別] 完了 反映行数=' + written);
  SpreadsheetApp.getActive().toast('楽天 セット別競合価格を ' + written + ' 行に反映しました', '楽天セット別', 6);
}

/**
 * 要確認Yahoo! 列に「要確認」「確認OK」のリスト入力ルールを設定し、先頭データ行（ヘッダー次の行）からのみ対象にする。
 * プルダウン項目の色は GAS で指定できないため、条件付き書式で「要確認」＝赤、「確認OK」＝青の背景を付ける。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} masterSheet
 * @param {Object.<string, number>} masterColMap
 * @param {number} headerRowIdx - ヘッダー行の 0-based 行インデックス
 */
function setYahooReviewColumnValidation(masterSheet, masterColMap, headerRowIdx) {
  Logger.log('[Yahoo!セット別] setYahooReviewColumnValidation が呼ばれた（入力規則を上書きする）');
  var colReview = masterColMap[COL_REVIEW_STATUS_YAHOO];
  if (colReview === undefined) return;
  var firstDataRow1Based = (headerRowIdx != null && headerRowIdx >= 0) ? headerRowIdx + 2 : 2;
  var lastRow = Math.max(masterSheet.getLastRow(), firstDataRow1Based + 1);
  var numRows = lastRow - firstDataRow1Based + 1;
  if (numRows < 1) return;
  var range = masterSheet.getRange(firstDataRow1Based, colReview + 1, numRows, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['要確認', '確認OK'], true)
    .setAllowInvalid(true)
    .build();
  range.setDataValidation(rule);
  var rules = masterSheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('要確認')
    .setBackground('#c53929')
    .setFontColor('#ffffff')
    .setRanges([range])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('確認OK')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setRanges([range])
    .build());
  masterSheet.setConditionalFormatRules(rules);
}

/**
 * Keepa API で複数 ASIN の送料込み最安値（NEW_FBM_SHIPPING の最新値）を取得する。
 * @param {string} apiKey Keepa API キー
 * @param {string[]} asins ASIN の配列（最大20推奨、トークン消費 = 1/ASIN）
 * @return {Object.<string, number>|null} ASIN → 価格（円）のマップ。失敗時は null
 */
function fetchCompetitivePricesFromKeepa(apiKey, asins) {
  if (!asins || asins.length === 0) return {};
  const url = 'https://api.keepa.com/product?key=' + encodeURIComponent(apiKey) + '&domain=' + KEEPA_DOMAIN_JP + '&asin=' + asins.join(',');
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    const text = res.getContentText();
    if (code !== 200) return null;
    const json = JSON.parse(text);
    if (!json.products) return {};
    const priceMap = {};
    json.products.forEach(function(p) {
      const asin = p.asin;
      if (!asin) return;
      var price = null;
      if (p.csv && p.csv[KEEPA_INDEX_NEW_FBM_SHIPPING]) {
        var arr = p.csv[KEEPA_INDEX_NEW_FBM_SHIPPING];
        for (var i = arr.length - 1; i >= 0; i -= 2) {
          if (i >= 1 && arr[i] !== -1 && arr[i] !== null) { price = arr[i]; break; }
        }
      }
      if (price == null && p.csv && p.csv[0]) {
        var arr0 = p.csv[0];
        for (var j = arr0.length - 1; j >= 0; j -= 2) {
          if (j >= 1 && arr0[j] !== -1 && arr0[j] !== null) { price = arr0[j]; break; }
        }
      }
      if (price != null) priceMap[asin] = price;
    });
    return priceMap;
  } catch (e) {
    console.warn('[Keepa] ' + e.message);
    return null;
  }
}

// ----------------------------------------
// 3.1c-2 楽天・Yahoo! 競合価格API 最小テスト（テスト終了後メニュー削除想定）
// ----------------------------------------
/** テスト用固定 JAN（よくある商品。JAN で 0 件の場合はキーワードで再検索する）。 */
var TEST_RAKUTEN_YAHOO_JAN = '4970107110284';
/** JAN で 0 件だったときのフォールバック用キーワード。 */
var TEST_RAKUTEN_YAHOO_KEYWORD_FALLBACK = 'ティッシュ';

/**
 * 楽天市場商品検索API（Ichiba Item Search 2022-06-01）で 1 件目の価格を取得する。
 * Script Properties: RAKUTEN_APP_ID, RAKUTEN_ACCESS_KEY（application_secret または Access Key）。
 * @param {string} appId - applicationId
 * @param {string} accessKey - accessKey（Bearer で送る）
 * @param {string} keyword - 検索キーワード（JAN または商品名）
 * @param {{ field?: number, orFlag?: number }} [opt] - 省略可。field: 0=Broad/1=Restricted, orFlag: 1でOR検索
 * @return {{ price: number|null, error: string|null, rawCount: number }}
 */
function fetchRakutenIchibaItemPrice(appId, accessKey, keyword, opt) {
  var result = { price: null, error: null, rawCount: 0 };
  if (!appId || !accessKey) {
    result.error = 'RAKUTEN_APP_ID または RAKUTEN_ACCESS_KEY が未設定です。';
    return result;
  }
  var base = 'https://openapi.rakuten.co.jp/ichibams/api/IchibaItem/Search/20220601';
  var field = (opt && opt.field !== undefined) ? opt.field : 0;
  var orFlag = (opt && opt.orFlag === 1) ? 1 : 0;
  var q = 'applicationId=' + encodeURIComponent(appId) + '&accessKey=' + encodeURIComponent(accessKey) + '&keyword=' + encodeURIComponent(keyword) + '&format=json&formatVersion=2&hits=1&elements=itemName,itemPrice&field=' + field + '&availability=0';
  if (orFlag === 1) q += '&orFlag=1';
  var url = base + '?' + q;
  try {
    var options = {
      muteHttpExceptions: true,
      headers: {
        'Authorization': 'Bearer ' + accessKey,
        'Origin': 'https://script.google.com',
        'Referer': 'https://script.google.com/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0'
      }
    };
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    var text = res.getContentText();
    if (code !== 200) {
      result.error = 'HTTP ' + code + ' ' + (text ? text.substring(0, 200) : '');
      Logger.log('[楽天API] ' + result.error);
      return result;
    }
    var json = JSON.parse(text);
    if (json.error) {
      result.error = (json.error_description || json.error) + '';
      Logger.log('[楽天API] 認証またはパラメータエラー: ' + result.error);
      return result;
    }
    var items = (json.items || []);
    result.rawCount = items.length;
    if (items.length > 0 && items[0].itemPrice != null) {
      result.price = parseInt(items[0].itemPrice, 10);
      if (isNaN(result.price)) result.price = null;
    }
    if (items.length === 0) {
      Logger.log('[楽天API] keyword="' + (keyword ? keyword.substring(0, 40) + (keyword.length > 40 ? '...' : '') : '') + '" field=' + field + (orFlag === 1 ? ' orFlag=1' : '') + ' → 0件');
    } else if (field === 0 || orFlag === 1) {
      Logger.log('[楽天API] ヒット: keyword="' + (keyword ? keyword.substring(0, 30) : '') + '..." field=' + field + (orFlag === 1 ? ' orFlag=1' : '') + ' count=' + items.length);
    }
    return result;
  } catch (e) {
    result.error = e.message || String(e);
    Logger.log('[楽天API] 例外: ' + result.error);
    return result;
  }
}

/**
 * 楽天 Ichiba Item Search API で複数件取得し Items 配列を返す。セット別競合価格用（既存の fetchRakutenIchibaItemPrice は変更しない）。
 * レスポンスの配列は json.Items || json.items で取得。docs/RAKUTEN_YAHOO_COMPETITIVE_PRICE_REQUIREMENTS.md 参照。
 * @param {string} appId - applicationId
 * @param {string} accessKey - accessKey
 * @param {string} keyword - 検索キーワード（JAN または メーカー+商品名）
 * @param {{ hits?: number, page?: number, field?: number, orFlag?: number }} [opt] - hits 省略時は 30。page でページ指定（1-based）
 * @return {{ error: string|null, items: Array<{itemName:string, itemPrice:number, itemUrl:string, imageUrl:string, postageFlag:number, pointRate:number}> }}
 */
function fetchRakutenIchibaItems(appId, accessKey, keyword, opt) {
  var result = { error: null, items: [] };
  if (!appId || !accessKey) {
    result.error = 'RAKUTEN_APP_ID または RAKUTEN_ACCESS_KEY が未設定です。';
    return result;
  }
  var base = 'https://openapi.rakuten.co.jp/ichibams/api/IchibaItem/Search/20220601';
  var field = (opt && opt.field !== undefined) ? opt.field : 0;
  var orFlag = (opt && opt.orFlag === 1) ? 1 : 0;
  var hits = (opt && opt.hits > 0) ? Math.min(opt.hits, 30) : 30;
  var page = (opt && opt.page >= 1) ? opt.page : 1;
  var q = 'applicationId=' + encodeURIComponent(appId) + '&accessKey=' + encodeURIComponent(accessKey) + '&keyword=' + encodeURIComponent(keyword) + '&format=json&formatVersion=2&hits=' + hits + '&page=' + page + '&field=' + field + '&availability=0';
  if (orFlag === 1) q += '&orFlag=1';
  var url = base + '?' + q;
  try {
    var options = {
      muteHttpExceptions: true,
      headers: {
        'Authorization': 'Bearer ' + accessKey,
        'Origin': 'https://script.google.com',
        'Referer': 'https://script.google.com/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0'
      }
    };
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    var text = res.getContentText();
    if (code !== 200) {
      result.error = 'HTTP ' + code + ' ' + (text ? text.substring(0, 200) : '');
      Logger.log('[楽天セット別] API keyword="' + (keyword ? String(keyword).substring(0, 30) : '') + '" ' + result.error);
      return result;
    }
    var json = JSON.parse(text);
    if (json.error) {
      result.error = (json.error_description || json.error) + '';
      Logger.log('[楽天セット別] API 認証/パラメータエラー: ' + result.error);
      return result;
    }
    var rawItems = (json.Items || json.items || []);
    for (var i = 0; i < rawItems.length; i++) {
      var it = rawItems[i];
      var price = (it.itemPrice != null) ? parseInt(it.itemPrice, 10) : null;
      if (price == null || isNaN(price)) continue;
      var imageUrl = '';
      if (it.mediumImageUrls && it.mediumImageUrls.length > 0) {
        if (typeof it.mediumImageUrls[0] === 'string') imageUrl = String(it.mediumImageUrls[0] || '');
        else if (it.mediumImageUrls[0] && it.mediumImageUrls[0].imageUrl != null) imageUrl = String(it.mediumImageUrls[0].imageUrl || '');
      }
      if (!imageUrl && it.smallImageUrls && it.smallImageUrls.length > 0) {
        if (typeof it.smallImageUrls[0] === 'string') imageUrl = String(it.smallImageUrls[0] || '');
        else if (it.smallImageUrls[0] && it.smallImageUrls[0].imageUrl != null) imageUrl = String(it.smallImageUrls[0].imageUrl || '');
      }
      result.items.push({
        itemName: (it.itemName != null) ? String(it.itemName) : '',
        itemPrice: price,
        itemUrl: (it.itemUrl != null) ? String(it.itemUrl) : '',
        imageUrl: imageUrl,
        postageFlag: (it.postageFlag != null) ? parseInt(it.postageFlag, 10) : 0,
        pointRate: (it.pointRate != null) ? parseInt(it.pointRate, 10) : 0
      });
    }
    Logger.log('[楽天セット別] API keyword="' + (keyword ? String(keyword).substring(0, 30) : '') + '" 取得件数=' + rawItems.length + ' パース後=' + result.items.length);
    return result;
  } catch (e) {
    result.error = e.message || String(e);
    Logger.log('[楽天セット別] API 例外: ' + result.error);
    return result;
  }
}

/**
 * 楽天 商品価格ナビ製品検索API（Product Search 2025-08-01）で productCode=JAN を指定して検索する。
 * 認証は Ichiba と同じ RAKUTEN_APP_ID / RAKUTEN_ACCESS_KEY を使用。
 * @param {string} appId - applicationId
 * @param {string} accessKey - accessKey（Bearer ヘッダで送る）
 * @param {string} janCode - JAN コード（productCode）
 * @return {{ error: string|null, count: number, productName: string, salesMinPrice: number|null, minPrice: number|null, mediumImageUrl: string, smallImageUrl: string, makerName: string }}
 */
function fetchRakutenProductSearchByJan(appId, accessKey, janCode) {
  var result = { error: null, count: 0, productName: '', salesMinPrice: null, minPrice: null, mediumImageUrl: '', smallImageUrl: '', makerName: '' };
  if (!appId || !accessKey) {
    result.error = 'RAKUTEN_APP_ID または RAKUTEN_ACCESS_KEY が未設定です。';
    return result;
  }
  var base = 'https://openapi.rakuten.co.jp/ichibaproduct/api/Product/Search/20250801';
  var q = 'applicationId=' + encodeURIComponent(appId) + '&accessKey=' + encodeURIComponent(accessKey) + '&productCode=' + encodeURIComponent(String(janCode).trim()) + '&format=json&formatVersion=2&hits=1';
  var url = base + '?' + q;
  try {
    var options = {
      muteHttpExceptions: true,
      headers: {
        'Authorization': 'Bearer ' + accessKey,
        'Origin': 'https://script.google.com',
        'Referer': 'https://script.google.com/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0'
      }
    };
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    var text = res.getContentText();
    if (code !== 200) {
      result.error = 'HTTP ' + code + ' ' + (text ? text.substring(0, 250) : '');
      return result;
    }
    var json = JSON.parse(text);
    if (json.error) {
      result.error = (json.error_description || json.error) + '';
      return result;
    }
    var items = (json.Products || json.items || []);
    result.count = parseInt(json.count, 10) || items.length;
    if (items.length > 0) {
      var it = items[0];
      var row = (it.item != null) ? it.item : it;
      var getStr = function (o, a, b) { var v = (o[a] != null) ? o[a] : (o[b] != null) ? o[b] : null; return (v != null) ? String(v) : ''; };
      var getNum = function (o, a, b) { var v = (o[a] != null) ? o[a] : (o[b] != null) ? o[b] : null; return (v != null && !isNaN(parseInt(v, 10))) ? parseInt(v, 10) : null; };
      result.productName = getStr(row, 'productName', 'product_name');
      result.salesMinPrice = getNum(row, 'salesMinPrice', 'sales_min_price');
      result.minPrice = getNum(row, 'minPrice', 'min_price');
      result.mediumImageUrl = getStr(row, 'mediumImageUrl', 'medium_image_url');
      result.smallImageUrl = getStr(row, 'smallImageUrl', 'small_image_url');
      result.makerName = getStr(row, 'makerName', 'maker_name');
    }
    return result;
  } catch (e) {
    result.error = e.message || String(e);
    return result;
  }
}

/**
 * 商品名を Gemini で分解し、メーカー・コア商品名・内容量・袋数を返す。楽天検索でコア商品名のみ使う用途。
 * @param {string} fullName - 例: "イトク食品 蒸し生姜湯 16g×5袋"
 * @return {{ maker: string, product: string, contentAmount: string, bagCount: string }|null} 失敗時は null
 */
function parseProductNameByGemini(fullName) {
  if (!fullName || typeof fullName !== 'string' || fullName.trim().length === 0) return null;
  var key = getGeminiApiKey();
  if (!key) return null;
  var prompt = '以下の商品名を分解し、JSONのみで答えてください。\n'
    + '・メーカー: メーカー名・ブランド名（無ければ空文字）\n'
    + '・商品: コアとなる商品名（検索キーワードに使う短い名前。例: 蒸し生姜湯）\n'
    + '・内容量: 数値＋単位（例: 16g）。無ければ空文字\n'
    + '・袋数: 袋数・個数（数値のみ、例: 5）。無ければ空文字\n'
    + '回答は次の形式のJSON1つだけ。説明や改行は不要。\n'
    + '{"メーカー":"","商品":"","内容量":"","袋数":""}\n\n'
    + '商品名: ' + fullName.trim();
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (typeof MODEL_GEMINI !== 'undefined' ? MODEL_GEMINI : 'gemini-2.0-flash') + ':generateContent?key=' + encodeURIComponent(key);
    var body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1, maxOutputTokens: 512 } };
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body) });
    var code = res.getResponseCode();
    if (code !== 200) {
      Logger.log('[商品名分解] Gemini HTTP ' + code + ' ' + (res.getContentText() ? res.getContentText().substring(0, 150) : ''));
      return null;
    }
    var json = JSON.parse(res.getContentText());
    var text = (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) ? json.candidates[0].content.parts[0].text : '';
    text = (text || '').replace(/```\w*\n?/g, '').trim();
    if (!text) {
      Logger.log('[商品名分解] Gemini 応答テキストが空');
      return null;
    }
    var start = text.indexOf('{');
    var end = text.lastIndexOf('}');
    if (start < 0 || end <= start) {
      Logger.log('[商品名分解] JSON未検出 応答先頭300字: ' + text.substring(0, 300));
      return null;
    }
    var parsed = JSON.parse(text.substring(start, end + 1));
    if (!parsed || typeof parsed !== 'object') {
      Logger.log('[商品名分解] パース後がオブジェクトでない');
      return null;
    }
    var getStr = function (p, ja, en) {
      var v = (p[ja] != null ? p[ja] : p[en]);
      return (v != null ? String(v).trim() : '') || '';
    };
    var result = {
      maker: getStr(parsed, 'メーカー', 'maker'),
      product: getStr(parsed, '商品', 'product'),
      contentAmount: getStr(parsed, '内容量', 'contentAmount'),
      bagCount: getStr(parsed, '袋数', 'bagCount')
    };
    if (!result.product) {
      Logger.log('[商品名分解] 商品名が空 応答先頭300字: ' + text.substring(0, 300));
    }
    return result;
  } catch (e) {
    Logger.log('[商品名分解] Gemini 例外: ' + (e && e.message));
    return null;
  }
}

/**
 * 楽天 Ichiba API を page=1,2,3 で最大90件取得し、商品名解析・セット数抽出・単価計算・セット数別最安をログに出すテスト。
 */
function menuTestRakuten90ItemsSetCountMinPrice() {
  var props = PropertiesService.getScriptProperties();
  var appId = (props.getProperty('RAKUTEN_APP_ID') || '').trim();
  var accessKey = (props.getProperty('RAKUTEN_ACCESS_KEY') || '').trim();
  if (!appId || !accessKey) {
    SpreadsheetApp.getUi().alert('Script Properties に RAKUTEN_APP_ID と RAKUTEN_ACCESS_KEY を設定してください。');
    return;
  }
  var keyword = '蒸し生姜湯';
  var allItems = [];
  for (var page = 1; page <= 3; page++) {
    var res = fetchRakutenIchibaItems(appId, accessKey, keyword, { hits: 30, page: page });
    if (res.error) {
      Logger.log('[楽天90件テスト] page=' + page + ' エラー: ' + res.error);
      break;
    }
    if (res.items && res.items.length > 0) {
      for (var i = 0; i < res.items.length; i++) allItems.push(res.items[i]);
    }
    if (page < 3) Utilities.sleep(300);
  }
  Logger.log('[楽天90件テスト] keyword="' + keyword + '" 合計取得=' + allItems.length + ' 件');
  var setToMinPrice = {};
  var setToUrl = {};
  var setToCount = {};
  for (var j = 0; j < allItems.length; j++) {
    var it = allItems[j];
    var setCount = parseSetCountFromItemName(it.itemName);
    if (setCount == null) continue;
    var effectivePrice = it.itemPrice - Math.round((it.itemPrice * (it.pointRate || 0)) / 100);
    if (setToMinPrice[setCount] == null || effectivePrice < setToMinPrice[setCount]) {
      setToMinPrice[setCount] = effectivePrice;
      setToUrl[setCount] = it.itemUrl;
    }
    setToCount[setCount] = (setToCount[setCount] || 0) + 1;
  }
  var setCounts = Object.keys(setToMinPrice).map(function (k) { return parseInt(k, 10); }).sort(function (a, b) { return a - b; });
  for (var s = 0; s < setCounts.length; s++) {
    var sc = setCounts[s];
    var price = setToMinPrice[sc];
    var unitPrice = (price / sc).toFixed(0);
    Logger.log('[楽天90件テスト] セット数=' + sc + ' 件数=' + (setToCount[sc] || 0) + ' 最安(税込)=' + price + ' 円 1個単価=' + unitPrice + ' 円');
  }
  var summary = allItems.length + '件取得 セット数別=' + setCounts.length + '種（ログで最安確認）';
  SpreadsheetApp.getActive().toast(summary, '楽天 90件＋セット数・最安テスト', 6);
}

/** 楽天90件スコア一覧を書き出すシート名（テスト用）。人間が目でチェックする。 */
const RAKUTEN_SCORE_LIST_SHEET_NAME = '楽天スコア一覧';
/** 楽天セット数 Gemini一括判定テスト用シート名。 */
const RAKUTEN_SET_COUNT_BATCH_SHEET_NAME = '楽天セット数一括判定';
/** 楽天・Yahoo!・Amazon の候補を1シートに集約し、Gemini の統合判定を出すテスト用シート名。 */
const CROSS_MALL_SET_COUNT_SHEET_NAME = 'モール横断セット数判定';

/**
 * 楽天 90件取得し、商品名一致スコアを降順で一覧出力する（ログ＋シート）。URL 付きで人間が目視確認できる。
 * マスタを開き対象行を選択した状態で実行。選択行の JAN・メーカー・商品名で検索し、90件をスコア降順で「楽天スコア一覧」シートに書き出す。
 */
function menuTestRakuten90ItemsScoreList() {
  var props = PropertiesService.getScriptProperties();
  var appId = (props.getProperty('RAKUTEN_APP_ID') || '').trim();
  var accessKey = (props.getProperty('RAKUTEN_ACCESS_KEY') || '').trim();
  if (!appId || !accessKey) {
    SpreadsheetApp.getUi().alert('Script Properties に RAKUTEN_APP_ID と RAKUTEN_ACCESS_KEY を設定してください。');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。');
    return;
  }
  if (ss.getActiveSheet().getSheetName() !== MASTER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('マスタシートを開き、対象行を1行以上選択してから実行してください。');
    return;
  }
  var range = masterSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してから実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(masterValues.length, 20); r++) {
    if (masterValues[r].indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colMaker = masterColMap['メーカー名'];
  var colName = masterColMap['商品名ベース'] != null ? masterColMap['商品名ベース'] : masterColMap['商品名'];
  if (colJan === undefined) {
    SpreadsheetApp.getUi().alert('マスタに JANコード 列がありません。');
    return;
  }
  var firstRow = range.getRow();
  var dataIdx = firstRow - 1;
  if (dataIdx <= headerRowIdx) {
    SpreadsheetApp.getUi().alert('データ行を選択してください。');
    return;
  }
  var jan = String(masterValues[dataIdx][colJan] || '').trim();
  if (jan.length < 8) {
    SpreadsheetApp.getUi().alert('選択行の JANコード が無効です。');
    return;
  }
  var maker = (colMaker !== undefined) ? String(masterValues[dataIdx][colMaker] || '').trim() : '';
  var name = (colName !== undefined) ? String(masterValues[dataIdx][colName] || '').trim() : '';
  var nameForScore = name;
  if (name.length > 0 && getGeminiApiKey()) {
    var parsed = parseProductNameByGemini(name);
    if (parsed && parsed.product && String(parsed.product).trim().length > 0) {
      nameForScore = String(parsed.product).trim();
      Logger.log('[楽天スコア一覧] スコア用期待名(コア)="' + nameForScore + '" 元="' + name.substring(0, 40) + (name.length > 40 ? '...' : '') + '"');
    }
    Utilities.sleep(350);
  }
  var fallbackKeyword = (maker + ' ' + name).trim();
  SpreadsheetApp.getActive().toast('90件取得・スコア計算中...', '楽天スコア一覧', 3);
  var res = fetchRakutenIchibaItems(appId, accessKey, jan);
  var keywordUsed = jan;
  if (res.items.length === 0 && fallbackKeyword.length >= 2) {
    Utilities.sleep(300);
    res = fetchRakutenIchibaItems(appId, accessKey, fallbackKeyword);
    keywordUsed = fallbackKeyword;
  }
  var allItems = (res.items || []).slice();
  for (var page = 2; page <= 3; page++) {
    var nextRes = fetchRakutenIchibaItems(appId, accessKey, keywordUsed, { hits: 30, page: page });
    if (nextRes.items && nextRes.items.length > 0) {
      for (var ni = 0; ni < nextRes.items.length; ni++) allItems.push(nextRes.items[ni]);
    }
    if (page < 3) Utilities.sleep(300);
  }
  var scored = [];
  for (var i = 0; i < allItems.length; i++) {
    var it = allItems[i];
    var nameMatch = evalRakutenItemNameMatchScore(maker, nameForScore, it.itemName);
    var parsedSet = parseSetCountFromItemNameWithSource(it.itemName);
    var setCount = null;
    var setCountReason = '不明';
    if (parsedSet) {
      setCount = parsedSet.setCount;
      if (parsedSet.fromP && getGeminiApiKey()) {
        var geminiSet = inferSetCountFromItemNameByGemini(it.itemName);
        if (geminiSet != null && geminiSet >= 1) {
          setCount = geminiSet;
          setCountReason = 'Gemini';
        } else {
          setCount = null;
          setCountReason = '正規表現(P)→Gemini不明';
        }
        Utilities.sleep(350);
      } else {
        setCountReason = '正規表現';
      }
    }
    scored.push({
      score: nameMatch.score,
      itemName: it.itemName || '',
      itemUrl: it.itemUrl || '',
      usedRules: nameMatch.usedRules,
      setCount: setCount,
      setCountReason: setCountReason
    });
  }
  scored.sort(function (a, b) { return (b.score - a.score); });
  Logger.log('[楽天スコア一覧] JAN=' + jan + ' 取得=' + allItems.length + ' 件 スコア降順で出力（セット数・根拠付き）');
  for (var r = 0; r < scored.length; r++) {
    var ruleStr = scored[r].usedRules.map(function (x) { return x.rule + ':' + x.score; }).join(',');
    Logger.log('[楽天スコア一覧] ' + (r + 1) + '位 score=' + scored[r].score + ' セット数=' + (scored[r].setCount != null ? scored[r].setCount : '不明') + ' 根拠=' + scored[r].setCountReason + ' ルール=' + ruleStr);
  }
  var sheet = ss.getSheetByName(RAKUTEN_SCORE_LIST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RAKUTEN_SCORE_LIST_SHEET_NAME);
  }
  sheet.clear();
  sheet.getRange(1, 1, 1, 7).setValues([['順位', 'スコア', '商品名', 'URL', '採用ルール', 'セット数', 'セット数根拠']]);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  var rows = [];
  for (var r = 0; r < scored.length; r++) {
    rows.push([
      r + 1,
      scored[r].score,
      (scored[r].itemName || '').substring(0, 200),
      scored[r].itemUrl || '',
      scored[r].usedRules.map(function (x) { return x.rule + ':' + x.score; }).join(','),
      scored[r].setCount != null ? scored[r].setCount : '不明',
      scored[r].setCountReason || '不明'
    ]);
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  SpreadsheetApp.getActive().toast('楽天スコア一覧を「' + RAKUTEN_SCORE_LIST_SHEET_NAME + '」に出力しました（' + scored.length + '件・セット数根拠付き）', '楽天スコア一覧', 6);
}

/**
 * 楽天 セット数 Gemini一括判定テスト。同一JANで取得した商品名をまとめて Gemini に送り、文脈を踏まえて各商品のセット数を一括判定する。
 * 結果を「楽天セット数一括判定」シートに出力。既存の「楽天 90件スコア一覧」のセット数列と比較してテストできる。
 */
function menuTestRakutenSetCountBatchGemini() {
  var props = PropertiesService.getScriptProperties();
  var appId = (props.getProperty('RAKUTEN_APP_ID') || '').trim();
  var accessKey = (props.getProperty('RAKUTEN_ACCESS_KEY') || '').trim();
  if (!appId || !accessKey) {
    SpreadsheetApp.getUi().alert('Script Properties に RAKUTEN_APP_ID と RAKUTEN_ACCESS_KEY を設定してください。');
    return;
  }
  if (!getGeminiApiKey()) {
    SpreadsheetApp.getUi().alert('Script Properties に GEMINI_API_KEY を設定してください。');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。');
    return;
  }
  if (ss.getActiveSheet().getSheetName() !== MASTER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('マスタシートを開き、対象行を1行以上選択してから実行してください。');
    return;
  }
  var range = masterSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してから実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(masterValues.length, 20); r++) {
    if (masterValues[r].indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colMaker = masterColMap['メーカー名'];
  var colName = masterColMap['商品名ベース'] != null ? masterColMap['商品名ベース'] : masterColMap['商品名'];
  if (colJan === undefined) {
    SpreadsheetApp.getUi().alert('マスタに JANコード 列がありません。');
    return;
  }
  var firstRow = range.getRow();
  var dataIdx = firstRow - 1;
  if (dataIdx <= headerRowIdx) {
    SpreadsheetApp.getUi().alert('データ行を選択してください。');
    return;
  }
  var jan = String(masterValues[dataIdx][colJan] || '').trim();
  if (jan.length < 8) {
    SpreadsheetApp.getUi().alert('選択行の JANコード が無効です。');
    return;
  }
  var maker = (colMaker !== undefined) ? String(masterValues[dataIdx][colMaker] || '').trim() : '';
  var name = (colName !== undefined) ? String(masterValues[dataIdx][colName] || '').trim() : '';
  var fallbackKeyword = (maker + ' ' + name).trim();
  SpreadsheetApp.getActive().toast('90件取得・一括判定中...', '楽天セット数一括', 3);
  var res = fetchRakutenIchibaItems(appId, accessKey, jan);
  var keywordUsed = jan;
  if (res.items.length === 0 && fallbackKeyword.length >= 2) {
    Utilities.sleep(300);
    res = fetchRakutenIchibaItems(appId, accessKey, fallbackKeyword);
    keywordUsed = fallbackKeyword;
  }
  var allItems = (res.items || []).slice();
  for (var page = 2; page <= 3; page++) {
    var nextRes = fetchRakutenIchibaItems(appId, accessKey, keywordUsed, { hits: 30, page: page });
    if (nextRes.items && nextRes.items.length > 0) {
      for (var ni = 0; ni < nextRes.items.length; ni++) allItems.push(nextRes.items[ni]);
    }
    if (page < 3) Utilities.sleep(300);
  }
  if (allItems.length === 0) {
    SpreadsheetApp.getUi().alert('取得件数が0件でした。');
    return;
  }
  var itemNames = allItems.map(function (it) { return it.itemName || ''; });
  var batchSetCounts = inferSetCountBatchByGemini(itemNames);
  var sheet = ss.getSheetByName(RAKUTEN_SET_COUNT_BATCH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RAKUTEN_SET_COUNT_BATCH_SHEET_NAME);
  }
  sheet.clear();
  sheet.getRange(1, 1, 1, 4).setValues([['行番号', '商品名', 'URL', 'セット数(一括)']]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  var rows = [];
  for (var r = 0; r < allItems.length; r++) {
    var sc = (r < batchSetCounts.length && batchSetCounts[r] >= 1) ? batchSetCounts[r] : (r < batchSetCounts.length && batchSetCounts[r] === 0 ? '不明' : '不明');
    rows.push([
      r + 1,
      (allItems[r].itemName || '').substring(0, 200),
      allItems[r].itemUrl || '',
      sc
    ]);
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  Logger.log('[楽天セット数一括] JAN=' + jan + ' 取得=' + allItems.length + ' 一括判定完了');
  SpreadsheetApp.getActive().toast('楽天セット数一括判定を「' + RAKUTEN_SET_COUNT_BATCH_SHEET_NAME + '」に出力しました（' + allItems.length + '件）', '楽天セット数一括', 6);
}

/**
 * 楽天・Yahoo!・Amazon の候補をまとめて Gemini に送り、モール横断でセット数を統合判定する。
 * Amazon は strongest evidence、Yahoo! JAN は strong evidence、Rakuten は補助証拠として扱うよう促す。
 * @param {Array<{source:string,title:string,price:number|null,url:string,localSetCount:number|null,localReason:string,score:number}>} candidates
 * @param {{jan:string, maker:string, coreName:string, baseUnitCost:number|null}} context
 * @return {Array<{setCount:number|null, reason:string}>}
 */
function parseCrossMallGeminiArrayText(text) {
  var raw = String(text || '').replace(/```(?:json)?\s*/g, '').trim();
  var start = raw.indexOf('[');
  if (start < 0) return { arr: null, mode: 'missing' };
  var end = raw.lastIndexOf(']');
  var arrayText = (end >= start)
    ? raw.substring(start, end + 1)
    : raw.substring(start);
  arrayText = arrayText
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/,\s*([}\]])/g, '$1');
  try {
    return { arr: JSON.parse(arrayText), mode: 'json' };
  } catch (e) {
    var objects = arrayText.match(/\{[\s\S]*?\}/g) || [];
    var fallback = [];
    for (var i = 0; i < objects.length; i++) {
      var s = objects[i];
      var idxMatch = s.match(/"idx"\s*:\s*(\d+)/) || s.match(/\bidx\b\s*:\s*(\d+)/);
      var setMatch = s.match(/"set"\s*:\s*(\d+)/) || s.match(/\bset\b\s*:\s*(\d+)/);
      var reasonMatch = s.match(/"reason"\s*:\s*"([^"\r\n}]*)/) || s.match(/\breason\b\s*:\s*"([^"\r\n}]*)/);
      var idx = idxMatch ? parseInt(idxMatch[1], 10) : NaN;
      var setVal = setMatch ? parseInt(setMatch[1], 10) : NaN;
      if (isNaN(idx) || idx < 1) continue;
      fallback.push({
        idx: idx,
        set: isNaN(setVal) ? 0 : setVal,
        reason: reasonMatch ? String(reasonMatch[1] || '').trim() : ''
      });
    }
    if (fallback.length > 0) return { arr: fallback, mode: 'regex' };
    return { arr: null, mode: 'error', error: e && e.message ? e.message : String(e), snippet: arrayText.substring(0, 300) };
  }
}

function getGeminiTextFromResponse(json) {
  var parts = (((json || {}).candidates || [])[0] || {}).content;
  parts = parts && parts.parts ? parts.parts : [];
  var texts = [];
  for (var i = 0; i < parts.length; i++) {
    if (parts[i] && parts[i].text != null) texts.push(String(parts[i].text));
  }
  return texts.join('').trim();
}

function buildCrossMallSetCountPrompt(candidates, context, opt) {
  var options = opt || {};
  var titleLimit = options.titleLimit != null ? options.titleLimit : 90;
  var prompt = '以下は同一JAN商品の候補一覧です。楽天・Yahoo!・Amazon の複数モールの証拠を総合して、各候補の販売セット数（何個セット・何袋セットで売っているか）を判定してください。\n'
    + '【前提】\n'
    + '・JAN: ' + ((context && context.jan) || '') + '\n'
    + '・メーカー: ' + ((context && context.maker) || '') + '\n'
    + '・コア商品名: ' + ((context && context.coreName) || '') + '\n'
    + '・基準単品原価（自社の推定1セット原価）: ' + (((context && context.baseUnitCost) != null) ? context.baseUnitCost : '') + '\n'
    + '【重要ルール】\n'
    + '1. Amazon の候補は既存フローで取得した情報のため最重要証拠、Yahoo! の JAN ヒットは強い証拠、Rakuten は補助証拠として扱ってください。\n'
    + '2. 数字のあとに P・袋・包・缶・個 などが複数出る場合、他候補でも同じ数字＋似た単位が多く繰り返されるなら、その数字は内容量（1パッケージ内の数量）を疑ってください。\n'
    + '3. セット数はパッケージ数量（×2個、3個セット、10袋セット等）を優先してください。複数オプション表記で1つに決められない場合は 0 にしてください。\n'
    + '4. 各候補には「比較用補正価格」「候補セット数で割った比較単価」が与えられます。これは競合の真の原価ではなく、送料・モール手数料を仮定差し引きした比較用の参考値です。少量セットは送料比率が高く見えるため、額面価格ではなくこの比較単価を優先して見てください。\n'
    + '5. 比較単価が他候補や基準単品原価と極端に乖離するセット数候補は優先度を下げてください。ただし内容量表記（4P、4袋入など）をセット数と誤認しないことを優先してください。\n'
    + '【回答形式】\n'
    + 'JSON配列のみ。各要素は {"idx":1,"set":10,"reason":"10個セット表記"} の形式。不明は {"idx":1,"set":0,"reason":"不明"}。\n'
    + 'reason は 20文字以内の短文にし、改行・ダブルクォートを含めないでください。配列以外の説明文は不要です。\n\n';
  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i];
    prompt += (i + 1) + '. '
      + 'source=' + c.source
      + ' price=' + (c.price != null ? c.price : '')
      + ' adjusted=' + (c.adjustedPrice != null ? c.adjustedPrice : '')
      + ' localSet=' + (c.localSetCount != null ? c.localSetCount : '')
      + ' localReason=' + (c.localReason || '')
      + ' title=' + (c.title || '').substring(0, titleLimit)
      + '\n';
  }
  return prompt;
}

function inferCrossMallSetCountsByGeminiCore(candidates, context, opt) {
  var options = opt || {};
  var batchLabel = options.batchLabel || 'single';
  if (!candidates || candidates.length === 0) return [];
  var key = getGeminiApiKey();
  if (!key) return { items: [], success: false, mode: 'no_key' };
  var prompt = buildCrossMallSetCountPrompt(candidates, context, options);
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (typeof MODEL_GEMINI !== 'undefined' ? MODEL_GEMINI : 'gemini-2.0-flash') + ':generateContent?key=' + encodeURIComponent(key);
    var body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1, maxOutputTokens: 4096, responseMimeType: 'application/json' } };
    Logger.log('[モール横断セット数] Gemini試行 batch=' + batchLabel + ' 候補=' + candidates.length + ' promptLen=' + prompt.length);
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body) });
    if (res.getResponseCode() !== 200) {
      Logger.log('[モール横断セット数] Gemini HTTP ' + res.getResponseCode() + ' batch=' + batchLabel);
      return { items: [], success: false, mode: 'http_' + res.getResponseCode() };
    }
    var json = JSON.parse(res.getContentText());
    var text = getGeminiTextFromResponse(json);
    var parsedResult = parseCrossMallGeminiArrayText(text);
    if (!parsedResult.arr) {
      Logger.log('[モール横断セット数] JSON未検出 batch=' + batchLabel + ' 応答先頭300字=' + text.substring(0, 300));
      if (parsedResult.error) Logger.log('[モール横断セット数] JSONパース失敗 batch=' + batchLabel + ' error=' + parsedResult.error + ' snippet=' + (parsedResult.snippet || ''));
      return { items: [], success: false, mode: parsedResult.mode || 'parse_failed' };
    }
    var arr = parsedResult.arr;
    if (!Array.isArray(arr)) return { items: [], success: false, mode: 'not_array' };
    Logger.log('[モール横断セット数] Gemini parseMode=' + parsedResult.mode + ' batch=' + batchLabel + ' 件数=' + arr.length);
    var byIdx = {};
    for (var a = 0; a < arr.length; a++) {
      var idx = parseInt(arr[a] && arr[a].idx, 10);
      if (!isNaN(idx) && idx >= 1) byIdx[idx - 1] = arr[a];
    }
    var out = [];
    var filled = 0;
    var noResultReason = (parsedResult.mode === 'regex' && arr.length < candidates.length) ? '（応答切れ）' : '（応答切れ）';
    for (var j = 0; j < candidates.length; j++) {
      var row = byIdx[j] || arr[j] || {};
      var n = parseInt(row.set, 10);
      var hasResult = !!(byIdx[j] || arr[j]);
      var reasonText = (row.reason != null ? String(row.reason).trim() : '') || '';
      if (!hasResult) reasonText = noResultReason;
      else if ((!isNaN(n) && n >= 1 && n <= 999)) { /* 有効なセット数ならそのまま */ }
      else if (!reasonText) reasonText = '（判定不明）';
      if (!isNaN(n) && n >= 1 && n <= 999) filled++;
      out.push({
        setCount: (!isNaN(n) && n >= 1 && n <= 999) ? n : null,
        reason: reasonText
      });
    }
    Logger.log('[モール横断セット数] Gemini採用 batch=' + batchLabel + ' 有効件数=' + filled + '/' + candidates.length);
    return { items: out, success: true, mode: parsedResult.mode || 'json' };
  } catch (e) {
    Logger.log('[モール横断セット数] 例外 batch=' + batchLabel + ' ' + (e && e.message));
    return { items: [], success: false, mode: 'exception' };
  }
}

function buildCrossMallGeminiBatches(candidates, maxPerBatch) {
  var limit = Math.max(6, Number(maxPerBatch) || 8);
  var amazon = [];
  var nonAmazon = [];
  for (var i = 0; i < candidates.length; i++) {
    if (candidates[i].source === 'amazon') amazon.push(i);
    else nonAmazon.push(i);
  }
  if (candidates.length <= limit) return [candidates.map(function (_, idx) { return idx; })];
  var anchorCount = Math.min(amazon.length, Math.max(0, limit - 6));
  var anchors = amazon.slice(0, anchorCount);
  var rest = amazon.slice(anchorCount).concat(nonAmazon);
  var bodyLimit = Math.max(1, limit - anchors.length);
  var batches = [];
  for (var start = 0; start < rest.length; start += bodyLimit) {
    var body = rest.slice(start, start + bodyLimit);
    batches.push(anchors.concat(body));
  }
  if (batches.length === 0) batches.push(anchors.slice());
  return batches;
}

function inferCrossMallSetCountsByGeminiAdaptive(candidates, context) {
  var full = inferCrossMallSetCountsByGeminiCore(candidates, context, { batchLabel: 'single', titleLimit: 90 });
  if (full.success && full.items && full.items.length === candidates.length) {
    var singleFilled = 0;
    for (var si = 0; si < full.items.length; si++) {
      if (full.items[si].setCount != null) singleFilled++;
    }
    var threshold = Math.max(1, Math.floor(candidates.length * 0.8));
    if (singleFilled >= threshold) {
      Logger.log('[モール横断セット数] Gemini戦略=single 候補=' + candidates.length + ' 有効=' + singleFilled);
      return { items: full.items, strategy: 'single' };
    }
  }
  var batches = buildCrossMallGeminiBatches(candidates, 8);
  Logger.log('[モール横断セット数] Gemini戦略=batched 候補=' + candidates.length + ' batchCount=' + batches.length);
  var merged = new Array(candidates.length);
  var retrySubSize = 4; // 応答切れ時リトライの1バッチあたり件数
  for (var b = 0; b < batches.length; b++) {
    var idxList = batches[b];
    var batchCandidates = [];
    for (var x = 0; x < idxList.length; x++) batchCandidates.push(candidates[idxList[x]]);
    var batchRes = inferCrossMallSetCountsByGeminiCore(batchCandidates, context, { batchLabel: 'split-' + (b + 1), titleLimit: 70 });
    if (!batchRes.success || !batchRes.items || batchRes.items.length === 0) continue;
    var hasTruncation = batchRes.items.some(function (item) { return item.reason === '（応答切れ）'; });
    if (hasTruncation && idxList.length > retrySubSize) {
      Logger.log('[モール横断セット数] 応答切れ検出 batch=split-' + (b + 1) + ' リトライ subSize=' + retrySubSize);
      for (var start = 0; start < idxList.length; start += retrySubSize) {
        var subIdx = idxList.slice(start, start + retrySubSize);
        var subCandidates = [];
        for (var sx = 0; sx < subIdx.length; sx++) subCandidates.push(candidates[subIdx[sx]]);
        var subLabel = 'split-' + (b + 1) + '-retry-' + (Math.floor(start / retrySubSize) + 1);
        var subRes = inferCrossMallSetCountsByGeminiCore(subCandidates, context, { batchLabel: subLabel, titleLimit: 70 });
        if (subRes.success && subRes.items) {
          for (var sy = 0; sy < subRes.items.length && sy < subIdx.length; sy++) {
            var origIdx = subIdx[sy];
            merged[origIdx] = subRes.items[sy];
          }
        }
      }
    } else {
      for (var y = 0; y < batchRes.items.length && y < idxList.length; y++) {
        var originalIdx = idxList[y];
        if (!merged[originalIdx]) merged[originalIdx] = batchRes.items[y];
      }
    }
  }
  var out = [];
  var mergedCount = 0;
  for (var i2 = 0; i2 < candidates.length; i2++) {
    var row2 = merged[i2] || { setCount: null, reason: '（未取得）' };
    if (row2.setCount != null) mergedCount++;
    out.push(row2);
  }
  Logger.log('[モール横断セット数] Gemini戦略=batched 完了 有効件数=' + mergedCount + '/' + candidates.length);
  return { items: out, strategy: 'batched' };
}

function inferCrossMallSetCountsByGemini(candidates, context) {
  var adaptive = inferCrossMallSetCountsByGeminiAdaptive(candidates, context);
  return adaptive && adaptive.items ? adaptive.items : [];
}

/**
 * モール横断のセット数判定用に、比較用の前提値（送料・手数料・原価）をマスタから取得する。
 * ここで作る数値は競合の真のコストではなく、候補比較のための仮定値。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Array[]} masterValues
 * @param {Object<string, number>} masterColMap
 * @param {number} headerRowIdx
 * @param {number} dataIdx
 * @param {string} jan
 * @return {{shipping:number, shippingSource:string, shippingBySet:Object<string, number>, shippingSourceBySet:Object<string, string>, costBySet:Object<string, number>, costSourceBySet:Object<string, string>, baseUnitCost:number|null, amazonCategory:string, amazonCategorySource:string, amazonFeeRate:number|null, amazonFeeSource:string, rateMap:Object<string, number>}}
 */
function getCrossMallComparisonContext(ss, masterValues, masterColMap, headerRowIdx, dataIdx, jan) {
  var parseNum = function(v) {
    if (v === undefined || v === null || v === '') return NaN;
    var n = Number(String(v).replace(/,/g, ''));
    return isNaN(n) ? NaN : n;
  };
  var colShipping = masterColMap[COL_SHIPPING_FIXED] !== undefined ? masterColMap[COL_SHIPPING_FIXED] : masterColMap[COL_SHIPPING];
  var colSetQty = masterColMap[COL_MASTER_TOTAL_QTY];
  var colCost = masterColMap[COL_COST_SET_TAX_IN] !== undefined ? masterColMap[COL_COST_SET_TAX_IN] : masterColMap[COL_COST_TAX_IN];
  var colAmazonCategory = masterColMap[COL_AMAZON_CATEGORY];
  var colAmazonFeeRate = masterColMap['amazon手数料率'];
  var colJan = masterColMap['JANコード'];
  var shippingBySet = {};
  var shippingSourceBySet = {};
  var costBySet = {};
  var costSourceBySet = {};
  function findFirstNumericSameJan(colIdx) {
    if (colIdx === undefined) return null;
    for (var rr = headerRowIdx + 1; rr < masterValues.length; rr++) {
      if (String(masterValues[rr][colJan] || '').trim() !== jan) continue;
      var n = parseNum(masterValues[rr][colIdx]);
      if (!isNaN(n) && n >= 0) return { value: n, row: rr + 1 };
    }
    return null;
  }
  function findFirstStringSameJan(colIdx) {
    if (colIdx === undefined) return null;
    for (var rr = headerRowIdx + 1; rr < masterValues.length; rr++) {
      if (String(masterValues[rr][colJan] || '').trim() !== jan) continue;
      var s = String(masterValues[rr][colIdx] || '').trim();
      if (s !== '') return { value: s, row: rr + 1 };
    }
    return null;
  }
  var shipping = (colShipping !== undefined) ? parseNum(masterValues[dataIdx][colShipping]) : NaN;
  var shippingSource = 'selected';
  if (isNaN(shipping) || shipping < 0) {
    var shippingFallback = findFirstNumericSameJan(colShipping);
    if (shippingFallback) {
      shipping = shippingFallback.value;
      shippingSource = 'sameJAN_row' + shippingFallback.row;
    } else {
      shipping = 0;
      shippingSource = 'fallback_zero';
    }
  }
  if (colShipping !== undefined && colSetQty !== undefined) {
    for (var rs = headerRowIdx + 1; rs < masterValues.length; rs++) {
      if (String(masterValues[rs][colJan] || '').trim() !== jan) continue;
      var setNumForShipping = parseNum(masterValues[rs][colSetQty]);
      var shippingNum = parseNum(masterValues[rs][colShipping]);
      if (isNaN(setNumForShipping) || setNumForShipping < 1) continue;
      if (isNaN(shippingNum) || shippingNum < 0) continue;
      var key = String(Math.round(setNumForShipping));
      if (shippingBySet[key] === undefined) {
        shippingBySet[key] = shippingNum;
        shippingSourceBySet[key] = 'sameJAN_set_row' + (rs + 1);
      }
    }
  }
  var baseUnitCost = null;
  if (colCost !== undefined && colSetQty !== undefined) {
    for (var r = headerRowIdx + 1; r < masterValues.length; r++) {
      if (String(masterValues[r][masterColMap['JANコード']] || '').trim() !== jan) continue;
      var setQty = parseNum(masterValues[r][colSetQty]);
      var cost = parseNum(masterValues[r][colCost]);
      if (!isNaN(setQty) && setQty >= 1 && !isNaN(cost) && cost > 0) {
        var costKey = String(Math.round(setQty));
        if (costBySet[costKey] === undefined) {
          costBySet[costKey] = cost;
          costSourceBySet[costKey] = 'sameJAN_set_row' + (r + 1);
        }
      }
      if (baseUnitCost == null && !isNaN(setQty) && setQty === 1 && !isNaN(cost) && cost > 0) {
        baseUnitCost = cost;
      }
    }
    if (baseUnitCost == null) {
      var selectedSet = parseNum(masterValues[dataIdx][colSetQty]);
      var selectedCost = parseNum(masterValues[dataIdx][colCost]);
      if (!isNaN(selectedCost) && selectedCost > 0) {
        if (!isNaN(selectedSet) && selectedSet >= 1) baseUnitCost = selectedCost / selectedSet;
        else baseUnitCost = selectedCost;
      }
    }
  }
  var amazonCategory = (colAmazonCategory !== undefined) ? String(masterValues[dataIdx][colAmazonCategory] || '').trim() : '';
  var amazonCategorySource = 'selected';
  if (!amazonCategory) {
    var categoryFallback = findFirstStringSameJan(colAmazonCategory);
    if (categoryFallback) {
      amazonCategory = categoryFallback.value;
      amazonCategorySource = 'sameJAN_row' + categoryFallback.row;
    } else {
      amazonCategorySource = 'empty';
    }
  }
  var amazonFeeRate = (colAmazonFeeRate !== undefined) ? parseNum(masterValues[dataIdx][colAmazonFeeRate]) : NaN;
  var amazonFeeSource = 'selected';
  if (isNaN(amazonFeeRate) || amazonFeeRate < 0) {
    var feeFallback = findFirstNumericSameJan(colAmazonFeeRate);
    if (feeFallback) {
      amazonFeeRate = feeFallback.value;
      amazonFeeSource = 'sameJAN_row' + feeFallback.row;
    } else {
      amazonFeeRate = null;
      amazonFeeSource = 'empty';
    }
  }
  var rateMap = getCommissionRateMapFromSettingsMaster(ss);
  return {
    shipping: shipping,
    shippingSource: shippingSource,
    shippingBySet: shippingBySet,
    shippingSourceBySet: shippingSourceBySet,
    costBySet: costBySet,
    costSourceBySet: costSourceBySet,
    baseUnitCost: baseUnitCost,
    amazonCategory: amazonCategory,
    amazonCategorySource: amazonCategorySource,
    amazonFeeRate: amazonFeeRate,
    amazonFeeSource: amazonFeeSource,
    rateMap: rateMap
  };
}

function getAssumedShippingForSetCount(setCount, ctx) {
  var out = {
    shipping: (ctx && ctx.shipping != null) ? ctx.shipping : 0,
    source: (ctx && ctx.shippingSource) ? ctx.shippingSource : 'fallback_zero'
  };
  var n = parseInt(setCount, 10);
  if (isNaN(n) || n < 1 || !ctx || !ctx.shippingBySet) return out;
  var key = String(n);
  if (ctx.shippingBySet[key] !== undefined) {
    out.shipping = ctx.shippingBySet[key];
    out.source = (ctx.shippingSourceBySet && ctx.shippingSourceBySet[key]) ? ctx.shippingSourceBySet[key] : ('sameJAN_set_' + key);
  }
  return out;
}

function getAssumedCostForSetCount(setCount, ctx) {
  var out = { cost: null, source: 'unknown' };
  var n = parseInt(setCount, 10);
  if (isNaN(n) || n < 1 || !ctx) return out;
  var key = String(n);
  if (ctx.costBySet && ctx.costBySet[key] !== undefined) {
    out.cost = ctx.costBySet[key];
    out.source = (ctx.costSourceBySet && ctx.costSourceBySet[key]) ? ctx.costSourceBySet[key] : ('sameJAN_set_' + key);
    return out;
  }
  if (ctx.baseUnitCost != null && !isNaN(ctx.baseUnitCost) && ctx.baseUnitCost > 0) {
    out.cost = Math.round(ctx.baseUnitCost * n);
    out.source = 'baseUnitCost*x' + n;
  }
  return out;
}

/**
 * 比較用のモール手数料率を返す。真値ではなく候補比較のための仮定値。
 * Amazon はカテゴリ・価格帯または amazon手数料率、楽天/Yahoo! は 00_設定マスタ のキーを優先し、無ければ 0。
 * @param {string} source
 * @param {number|null} price
 * @param {{rateMap:Object<string, number>, amazonCategory:string, amazonFeeRate:number|null}} ctx
 * @return {number}
 */
function getComparisonFeeRateForSource(source, price, ctx) {
  var rateMap = (ctx && ctx.rateMap) || {};
  if (source === 'amazon') {
    if (ctx && ctx.amazonCategory) {
      var byCategory = getAmazonCommissionRateForPrice(rateMap, ctx.amazonCategory, price);
      if (byCategory != null && !isNaN(byCategory)) return byCategory;
    }
    if (ctx && ctx.amazonFeeRate != null && !isNaN(ctx.amazonFeeRate)) return ctx.amazonFeeRate;
    return 0;
  }
  if (source === 'rakuten') {
    if (rateMap['楽天'] !== undefined) return rateMap['楽天'];
    if (rateMap['Rakuten'] !== undefined) return rateMap['Rakuten'];
    return 0;
  }
  if (source === 'yahoo') {
    if (rateMap['Yahoo!'] !== undefined) return rateMap['Yahoo!'];
    if (rateMap['Yahoo'] !== undefined) return rateMap['Yahoo'];
    if (rateMap['ヤフー'] !== undefined) return rateMap['ヤフー'];
    return 0;
  }
  return 0;
}

function applyImageAiToAmbiguousSetCandidates(candidates, limit) {
  var maxCount = Math.max(0, Number(limit) || 0);
  if (!candidates || candidates.length === 0 || maxCount <= 0 || !getGeminiApiKey()) return 0;
  var updated = 0;
  for (var i = 0; i < candidates.length && updated < maxCount; i++) {
    var c = candidates[i];
    if (!c || c.localReason !== '正規表現(P)' || !c.imageUrl) continue;
    var img = inferSetCountFromImageByGemini(c.imageUrl);
    if (img && img.setCount != null && img.setCount >= 1) {
      c.localSetCount = img.setCount;
      c.localReason = '画像AI';
      if (c.adjustedPrice != null) {
        c.localUnitComparable = Math.round(c.adjustedPrice / img.setCount);
        c.localCostRatio = (c.baseUnitCostForCalc != null && c.baseUnitCostForCalc > 0)
          ? Math.round((c.localUnitComparable / c.baseUnitCostForCalc) * 100) / 100
          : null;
      }
      updated++;
      Logger.log('[モール横断セット数] 画像AI採用 source=' + c.source + ' set=' + img.setCount + ' title=' + String(c.title || '').substring(0, 40));
    } else {
      Logger.log('[モール横断セット数] 画像AI不明 source=' + c.source + ' title=' + String(c.title || '').substring(0, 40));
    }
    Utilities.sleep(350);
  }
  return updated;
}

/**
 * モール横断のセット数統合判定テスト。楽天・Yahoo!・Amazon 候補を1シートに集約し、Gemini の一括判定結果を比較する。
 * 旧テストは比較用に残し、この統合テストを新しい主テストとして追加する。
 */
function menuTestCrossMallSetCountJudge() {
  var props = PropertiesService.getScriptProperties();
  var appId = (props.getProperty('RAKUTEN_APP_ID') || '').trim();
  var accessKey = (props.getProperty('RAKUTEN_ACCESS_KEY') || '').trim();
  var yahooId = (props.getProperty('YAHOO_SHOPPING_CLIENT_ID') || '').trim();
  var keepaKey = (props.getProperty('KEEPA_API_KEY') || '').trim();
  if (!appId || !accessKey || !yahooId) {
    SpreadsheetApp.getUi().alert('Script Properties に RAKUTEN_APP_ID / RAKUTEN_ACCESS_KEY / YAHOO_SHOPPING_CLIENT_ID を設定してください。');
    return;
  }
  if (!getGeminiApiKey()) {
    SpreadsheetApp.getUi().alert('Script Properties に GEMINI_API_KEY を設定してください。');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。');
    return;
  }
  if (ss.getActiveSheet().getSheetName() !== MASTER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('マスタシートを開き、対象行を1行選択してから実行してください。');
    return;
  }
  var range = masterSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してから実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(masterValues.length, 20); hr++) {
    if (masterValues[hr].indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = hr; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colMaker = masterColMap['メーカー名'];
  var colName = masterColMap['商品名ベース'] != null ? masterColMap['商品名ベース'] : masterColMap['商品名'];
  var colAsin = masterColMap['ASINコード'];
  var colCompetitorAsin = masterColMap['競合店ASINコード'];
  if (colJan === undefined) {
    SpreadsheetApp.getUi().alert('マスタに JANコード 列がありません。');
    return;
  }
  var firstRow = range.getRow();
  var dataIdx = firstRow - 1;
  if (dataIdx <= headerRowIdx) {
    SpreadsheetApp.getUi().alert('データ行を選択してください。');
    return;
  }
  var jan = String(masterValues[dataIdx][colJan] || '').trim();
  if (jan.length < 8) {
    SpreadsheetApp.getUi().alert('選択行の JANコード が無効です。');
    return;
  }
  var comparisonCtx = getCrossMallComparisonContext(ss, masterValues, masterColMap, headerRowIdx, dataIdx, jan);
  Logger.log('[モール横断セット数] 比較前提 shipping=' + comparisonCtx.shipping + ' source=' + comparisonCtx.shippingSource
    + ' amazonCategory=' + (comparisonCtx.amazonCategory || '(空)') + ' categorySource=' + comparisonCtx.amazonCategorySource
    + ' amazonFeeRate=' + (comparisonCtx.amazonFeeRate != null ? comparisonCtx.amazonFeeRate : 'null') + ' feeSource=' + comparisonCtx.amazonFeeSource);
  var maker = (colMaker !== undefined) ? String(masterValues[dataIdx][colMaker] || '').trim() : '';
  var name = (colName !== undefined) ? String(masterValues[dataIdx][colName] || '').trim() : '';
  var coreName = name;
  if (name.length > 0 && getGeminiApiKey()) {
    var parsed = parseProductNameByGemini(name);
    if (parsed && parsed.product) coreName = String(parsed.product).trim() || name;
    Utilities.sleep(350);
  }
  SpreadsheetApp.getActive().toast('モール横断候補を集約中...', 'モール横断セット数', 3);

  var candidates = [];

  // Rakuten candidates
  var fallbackKeyword = (maker + ' ' + name).trim();
  var rakutenRes = fetchRakutenIchibaItems(appId, accessKey, jan);
  var rakutenKeyword = jan;
  if (rakutenRes.items.length === 0 && fallbackKeyword.length >= 2) {
    Utilities.sleep(300);
    rakutenRes = fetchRakutenIchibaItems(appId, accessKey, fallbackKeyword);
    rakutenKeyword = fallbackKeyword;
  }
  var rakutenItems = (rakutenRes.items || []).slice();
  for (var rp = 2; rp <= 3; rp++) {
    Utilities.sleep(1000);
    var rakutenPage = fetchRakutenIchibaItems(appId, accessKey, rakutenKeyword, { hits: 30, page: rp });
    if (rakutenPage.error && String(rakutenPage.error).indexOf('429') >= 0) {
      Logger.log('[モール横断セット数] 楽天 429 検出 1秒待機してリトライ page=' + rp);
      Utilities.sleep(1000);
      rakutenPage = fetchRakutenIchibaItems(appId, accessKey, rakutenKeyword, { hits: 30, page: rp });
    }
    if (rakutenPage.items && rakutenPage.items.length > 0) {
      for (var rpi = 0; rpi < rakutenPage.items.length; rpi++) rakutenItems.push(rakutenPage.items[rpi]);
    }
  }
  for (var ri = 0; ri < rakutenItems.length; ri++) {
    var rItem = rakutenItems[ri];
    var rScore = evalRakutenItemNameMatchScore(maker, coreName, rItem.itemName);
    var rParsed = parseSetCountFromItemNameWithSource(rItem.itemName);
    var rShippingInfo = getAssumedShippingForSetCount(rParsed ? rParsed.setCount : null, comparisonCtx);
    var rCostInfo = getAssumedCostForSetCount(rParsed ? rParsed.setCount : null, comparisonCtx);
    var rFeeRate = getComparisonFeeRateForSource('rakuten', rItem.itemPrice, comparisonCtx);
    var rAdjusted = (rItem.itemPrice != null) ? Math.max(0, rItem.itemPrice - rShippingInfo.shipping - (rItem.itemPrice * rFeeRate)) : null;
    var rLocalUnit = (rParsed && rParsed.setCount != null && rParsed.setCount >= 1 && rAdjusted != null) ? (rAdjusted / rParsed.setCount) : null;
    var rLocalCostRatio = (rLocalUnit != null && comparisonCtx.baseUnitCost != null && comparisonCtx.baseUnitCost > 0) ? (rLocalUnit / comparisonCtx.baseUnitCost) : null;
    var rExpectedProfit = (rAdjusted != null && rCostInfo.cost != null) ? Math.round(rAdjusted - rCostInfo.cost) : null;
    candidates.push({
      source: 'rakuten',
      score: rScore.score,
      title: rItem.itemName || '',
      url: rItem.itemUrl || '',
      imageUrl: rItem.imageUrl || '',
      price: rItem.itemPrice != null ? rItem.itemPrice : null,
      localSetCount: rParsed ? rParsed.setCount : null,
      localReason: rParsed ? (rParsed.fromP ? '正規表現(P)' : '正規表現') : '不明',
      localRules: rScore.usedRules.map(function (x) { return x.rule + ':' + x.score; }).join(','),
      assumedShipping: rShippingInfo.shipping,
      assumedShippingSource: rShippingInfo.source,
      feeRate: rFeeRate,
      adjustedPrice: rAdjusted != null ? Math.round(rAdjusted) : null,
      assumedCost: rCostInfo.cost,
      assumedCostSource: rCostInfo.source,
      expectedProfit: rExpectedProfit,
      localUnitComparable: rLocalUnit != null ? Math.round(rLocalUnit) : null,
      localCostRatio: rLocalCostRatio != null ? Math.round(rLocalCostRatio * 100) / 100 : null,
      baseUnitCostForCalc: comparisonCtx.baseUnitCost
    });
  }

  // Yahoo candidates
  var yahooRes = fetchYahooShoppingItemsByJan(yahooId, jan, 50);
  var yahooHits = yahooRes && yahooRes.hits ? yahooRes.hits : [];
  for (var yi = 0; yi < yahooHits.length; yi++) {
    var yItem = yahooHits[yi];
    var yScore = evalRakutenItemNameMatchScore(maker, coreName, yItem.name);
    var yParsed = parseSetCountFromItemNameWithSource(yItem.name);
    var yShippingInfo = getAssumedShippingForSetCount(yParsed ? yParsed.setCount : null, comparisonCtx);
    var yCostInfo = getAssumedCostForSetCount(yParsed ? yParsed.setCount : null, comparisonCtx);
    var yFeeRate = getComparisonFeeRateForSource('yahoo', yItem.price, comparisonCtx);
    var yAdjusted = (yItem.price != null) ? Math.max(0, yItem.price - yShippingInfo.shipping - (yItem.price * yFeeRate)) : null;
    var yLocalUnit = (yParsed && yParsed.setCount != null && yParsed.setCount >= 1 && yAdjusted != null) ? (yAdjusted / yParsed.setCount) : null;
    var yLocalCostRatio = (yLocalUnit != null && comparisonCtx.baseUnitCost != null && comparisonCtx.baseUnitCost > 0) ? (yLocalUnit / comparisonCtx.baseUnitCost) : null;
    var yExpectedProfit = (yAdjusted != null && yCostInfo.cost != null) ? Math.round(yAdjusted - yCostInfo.cost) : null;
    candidates.push({
      source: 'yahoo',
      score: yScore.score,
      title: yItem.name || '',
      url: yItem.url || '',
      imageUrl: yItem.imageUrl || '',
      price: yItem.price != null ? yItem.price : null,
      localSetCount: yParsed ? yParsed.setCount : null,
      localReason: yParsed ? (yParsed.fromP ? '正規表現(P)' : '正規表現') : '不明',
      localRules: yScore.usedRules.map(function (x) { return x.rule + ':' + x.score; }).join(','),
      assumedShipping: yShippingInfo.shipping,
      assumedShippingSource: yShippingInfo.source,
      feeRate: yFeeRate,
      adjustedPrice: yAdjusted != null ? Math.round(yAdjusted) : null,
      assumedCost: yCostInfo.cost,
      assumedCostSource: yCostInfo.source,
      expectedProfit: yExpectedProfit,
      localUnitComparable: yLocalUnit != null ? Math.round(yLocalUnit) : null,
      localCostRatio: yLocalCostRatio != null ? Math.round(yLocalCostRatio * 100) / 100 : null,
      baseUnitCostForCalc: comparisonCtx.baseUnitCost
    });
  }

  // Amazon candidates (from same JAN rows' ASIN / competitor ASIN, fetched by Keepa)
  var asinMap = {};
  if (keepaKey) {
    for (var mr = headerRowIdx + 1; mr < masterValues.length; mr++) {
      if (String(masterValues[mr][colJan] || '').trim() !== jan) continue;
      var asin1 = (colAsin !== undefined) ? String(masterValues[mr][colAsin] || '').trim() : '';
      var asin2 = (colCompetitorAsin !== undefined) ? String(masterValues[mr][colCompetitorAsin] || '').trim() : '';
      if (asin1) asinMap[asin1] = true;
      if (asin2) asinMap[asin2] = true;
    }
    var asins = Object.keys(asinMap);
    if (asins.length > 0) {
      var keepaRes = fetchKeepaProductsForSheet(keepaKey, asins);
      if (keepaRes && keepaRes.list) {
        for (var ai = 0; ai < keepaRes.list.length; ai++) {
          var aItem = keepaRes.list[ai];
          var aScore = evalRakutenItemNameMatchScore(maker, coreName, aItem.title);
          var aShippingInfo = getAssumedShippingForSetCount(aItem.setCount, comparisonCtx);
          var aCostInfo = getAssumedCostForSetCount(aItem.setCount, comparisonCtx);
          var aFeeRate = getComparisonFeeRateForSource('amazon', aItem.price, comparisonCtx);
          var aAdjusted = (aItem.price != null) ? Math.max(0, aItem.price - aShippingInfo.shipping - (aItem.price * aFeeRate)) : null;
          var aLocalUnit = (aItem.setCount != null && aItem.setCount >= 1 && aAdjusted != null) ? (aAdjusted / aItem.setCount) : null;
          var aLocalCostRatio = (aLocalUnit != null && comparisonCtx.baseUnitCost != null && comparisonCtx.baseUnitCost > 0) ? (aLocalUnit / comparisonCtx.baseUnitCost) : null;
          var aExpectedProfit = (aAdjusted != null && aCostInfo.cost != null) ? Math.round(aAdjusted - aCostInfo.cost) : null;
          candidates.push({
            source: 'amazon',
            score: aScore.score,
            title: aItem.title || '',
            url: aItem.asin ? ('https://www.amazon.co.jp/dp/' + aItem.asin) : '',
            imageUrl: aItem.imageUrl || '',
            price: aItem.price != null ? aItem.price : null,
            localSetCount: aItem.setCount != null ? aItem.setCount : null,
            localReason: aItem.setCountReason || '不明',
            localRules: aScore.usedRules.map(function (x) { return x.rule + ':' + x.score; }).join(','),
            assumedShipping: aShippingInfo.shipping,
            assumedShippingSource: aShippingInfo.source,
            feeRate: aFeeRate,
            adjustedPrice: aAdjusted != null ? Math.round(aAdjusted) : null,
            assumedCost: aCostInfo.cost,
            assumedCostSource: aCostInfo.source,
            expectedProfit: aExpectedProfit,
            localUnitComparable: aLocalUnit != null ? Math.round(aLocalUnit) : null,
            localCostRatio: aLocalCostRatio != null ? Math.round(aLocalCostRatio * 100) / 100 : null,
            baseUnitCostForCalc: comparisonCtx.baseUnitCost
          });
        }
      }
    }
  }

  if (candidates.length === 0) {
    SpreadsheetApp.getUi().alert('候補を1件も取得できませんでした。');
    return;
  }

  candidates.sort(function (a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return String(a.source).localeCompare(String(b.source));
  });
  Logger.log('[モール横断セット数] 画像AI前倒し 停止');
  var integrated = [];
  for (var ii = 0; ii < candidates.length; ii++) integrated.push({ setCount: null, reason: '' });
  var geminiTargets = [];
  var geminiIndexMap = [];
  for (var gt = 0; gt < candidates.length; gt++) {
    var candidate = candidates[gt];
    var isAmbiguousP = candidate.localReason === '正規表現(P)';
    var shouldExcludeByProfit = (!isAmbiguousP && candidate.localSetCount != null && candidate.localSetCount >= 1 && candidate.expectedProfit != null && candidate.expectedProfit < 0);
    if (shouldExcludeByProfit) {
      integrated[gt] = { setCount: null, reason: '想定利益マイナスのため統合判定から除外' };
      continue;
    }
    geminiTargets.push(candidate);
    geminiIndexMap.push(gt);
  }
  if (geminiTargets.length > 0) {
    var integratedFiltered = inferCrossMallSetCountsByGemini(geminiTargets, { jan: jan, maker: maker, coreName: coreName, baseUnitCost: comparisonCtx.baseUnitCost });
    for (var gf = 0; gf < integratedFiltered.length; gf++) {
      integrated[geminiIndexMap[gf]] = integratedFiltered[gf];
    }
  }

  var sheet = ss.getSheetByName(CROSS_MALL_SET_COUNT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CROSS_MALL_SET_COUNT_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, 23).setValues([['順位', 'モール', 'スコア', '商品名', 'URL', '価格', '価格+送料', '仮定送料', '仮定手数料率', '比較用補正価格', 'ローカル候補セット数', 'ローカル根拠', 'ローカル比較単価', 'ローカル原価倍率', 'ローカル想定原価', 'ローカル想定利益額', '統合セット数', '統合根拠', '不明理由種別', '統合比較単価', '統合原価倍率', '統合想定原価', '統合想定利益額']]);
  sheet.getRange(1, 1, 1, 23).setFontWeight('bold');
  var rows = [];
  for (var ci = 0; ci < candidates.length; ci++) {
    var integratedSet = (integrated[ci] && integrated[ci].setCount != null && integrated[ci].setCount >= 1) ? integrated[ci].setCount : null;
    var integratedShippingInfo = getAssumedShippingForSetCount(integratedSet, comparisonCtx);
    var integratedCostInfo = getAssumedCostForSetCount(integratedSet, comparisonCtx);
    var integratedAdjusted = (integratedSet != null && candidates[ci].price != null)
      ? Math.round(Math.max(0, candidates[ci].price - integratedShippingInfo.shipping - (candidates[ci].price * (candidates[ci].feeRate || 0))))
      : '';
    var integratedUnit = (integratedSet != null && integratedAdjusted !== '') ? Math.round(integratedAdjusted / integratedSet) : '';
    var integratedCostRatio = (integratedUnit !== '' && comparisonCtx.baseUnitCost != null && comparisonCtx.baseUnitCost > 0) ? Math.round((integratedUnit / comparisonCtx.baseUnitCost) * 100) / 100 : '';
    var integratedProfit = (integratedAdjusted !== '' && integratedCostInfo.cost != null) ? Math.round(integratedAdjusted - integratedCostInfo.cost) : '';
    var priceInclShipping = (candidates[ci].price != null && candidates[ci].assumedShipping != null) ? (candidates[ci].price + candidates[ci].assumedShipping) : (candidates[ci].price != null ? candidates[ci].price : '');
    var reasonText = (integrated[ci] && integrated[ci].reason) ? integrated[ci].reason : '';
    var unknownKind = '';
    if (integratedSet == null) {
      if (reasonText.indexOf('想定利益マイナス') >= 0) unknownKind = '利益除外';
      else if (reasonText === '（応答切れ）') unknownKind = '応答切れ';
      else if (reasonText === '（判定不明）') unknownKind = '判定不明';
      else if (reasonText === '（未取得）') unknownKind = '未取得';
      else if (reasonText.length > 0) unknownKind = 'Gemini判定';
    }
    rows.push([
      ci + 1,
      candidates[ci].source,
      candidates[ci].score,
      (candidates[ci].title || '').substring(0, 200),
      candidates[ci].url || '',
      candidates[ci].price != null ? candidates[ci].price : '',
      priceInclShipping,
      candidates[ci].assumedShipping != null ? candidates[ci].assumedShipping : '',
      candidates[ci].feeRate != null ? candidates[ci].feeRate : '',
      candidates[ci].adjustedPrice != null ? candidates[ci].adjustedPrice : '',
      candidates[ci].localSetCount != null ? candidates[ci].localSetCount : '不明',
      candidates[ci].localReason || '不明',
      candidates[ci].localUnitComparable != null ? candidates[ci].localUnitComparable : '',
      candidates[ci].localCostRatio != null ? candidates[ci].localCostRatio : '',
      candidates[ci].assumedCost != null ? candidates[ci].assumedCost : '',
      candidates[ci].expectedProfit != null ? candidates[ci].expectedProfit : '',
      integratedSet != null ? integratedSet : '不明',
      reasonText,
      unknownKind,
      integratedUnit,
      integratedCostRatio,
      integratedCostInfo.cost != null ? integratedCostInfo.cost : '',
      integratedProfit
    ]);
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 23).setValues(rows);
  }
  if (comparisonCtx.baseUnitCost != null) sheet.getRange(1, 25).setValue('基準単品原価');
  if (comparisonCtx.baseUnitCost != null) sheet.getRange(2, 25).setValue(comparisonCtx.baseUnitCost);
  Logger.log('[モール横断セット数] JAN=' + jan + ' 楽天=' + rakutenItems.length + ' Yahoo=' + yahooHits.length + ' Amazon=' + (candidates.filter(function (x) { return x.source === 'amazon'; }).length) + ' 合計=' + candidates.length + ' Gemini対象=' + geminiTargets.length + ' 利益除外=' + (candidates.length - geminiTargets.length) + ' 比較用送料=' + comparisonCtx.shipping + ' 基準単品原価=' + (comparisonCtx.baseUnitCost != null ? comparisonCtx.baseUnitCost : 'null'));
  SpreadsheetApp.getActive().toast('モール横断セット数判定を「' + CROSS_MALL_SET_COUNT_SHEET_NAME + '」に出力しました（' + candidates.length + '件）', 'モール横断セット数', 8);
}

/**
 * Yahoo! ショッピング 商品検索 API v3 で jan_code 検索し 1 件目の価格を取得する。
 * Script Properties: YAHOO_SHOPPING_CLIENT_ID。
 * @param {string} clientId - appid（Client ID）
 * @param {string} janCode - JAN コード
 * @return {{ price: number|null, error: string|null, rawCount: number }}
 */
function fetchYahooShoppingItemPriceByJan(clientId, janCode) {
  var result = { price: null, error: null, rawCount: 0 };
  if (!clientId || !clientId.trim()) {
    result.error = 'YAHOO_SHOPPING_CLIENT_ID が未設定です。';
    return result;
  }
  var url = 'https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch?appid=' + encodeURIComponent(clientId) + '&jan_code=' + encodeURIComponent(janCode) + '&results=1';
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = res.getResponseCode();
    var text = res.getContentText();
    if (code !== 200) {
      result.error = 'HTTP ' + code + ' ' + (text ? text.substring(0, 200) : '');
      Logger.log('[Yahoo!API] ' + result.error);
      return result;
    }
    var json = JSON.parse(text);
    var hits = (json.hits || []);
    result.rawCount = hits.length;
    if (hits.length > 0 && hits[0].price != null) {
      result.price = parseInt(hits[0].price, 10);
      if (isNaN(result.price)) result.price = null;
    }
    return result;
  } catch (e) {
    result.error = e.message || String(e);
    Logger.log('[Yahoo!API] 例外: ' + result.error);
    return result;
  }
}

/**
 * Yahoo! ショッピング API v3 で jan_code 検索し、複数件（最大 maxResults）を取得する。セット別価格テスト用。
 * @param {string} clientId - appid（Client ID）
 * @param {string} janCode - JAN コード
 * @param {number} [maxResults=20] - 取得件数（最大50）
 * @return {{ error: string|null, hits: Array<{ name: string, price: number|null, url: string, imageUrl: string }> }}
 */
function fetchYahooShoppingItemsByJan(clientId, janCode, maxResults) {
  var result = { error: null, hits: [] };
  if (!clientId || !clientId.trim()) {
    result.error = 'YAHOO_SHOPPING_CLIENT_ID が未設定です。';
    return result;
  }
  var limit = Math.min(Math.max(Number(maxResults) || 20, 1), 50);
  var url = 'https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch?appid=' + encodeURIComponent(clientId) + '&jan_code=' + encodeURIComponent(janCode) + '&results=' + limit;
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = res.getResponseCode();
    var text = res.getContentText();
    if (code !== 200) {
      result.error = 'HTTP ' + code + ' ' + (text ? text.substring(0, 200) : '');
      Logger.log('[Yahoo!API] ' + result.error);
      return result;
    }
    var json = JSON.parse(text);
    var rawHits = (json.hits || []);
    for (var i = 0; i < rawHits.length; i++) {
      var h = rawHits[i];
      var name = (h.name != null) ? String(h.name) : '';
      var price = (h.price != null) ? parseInt(h.price, 10) : null;
      if (isNaN(price)) price = null;
      var itemUrl = (h.url != null) ? String(h.url) : '';
      var imageUrl = '';
      if (h.image && h.image.medium != null) imageUrl = String(h.image.medium || '');
      if (!imageUrl && h.image && h.image.small != null) imageUrl = String(h.image.small || '');
      if (!imageUrl && h.image && h.image.id != null) imageUrl = String(h.image.id || '');
      result.hits.push({ name: name, price: price, url: itemUrl, imageUrl: imageUrl });
    }
    return result;
  } catch (e) {
    result.error = e.message || String(e);
    Logger.log('[Yahoo!API] 例外: ' + result.error);
    return result;
  }
}

/**
 * 商品名からセット数（個数・袋数など）を推定する。表記ゆれに対応するため複数パターンを試す。
 * @param {string} name - 商品名（タイトル）
 * @return {number|null} 推定セット数（1〜999）。取れなければ null
 */
function parseSetCountFromItemName(name) {
  var r = parseSetCountFromItemNameWithSource(name);
  return r ? r.setCount : null;
}

/**
 * 商品名からセット数を推定し、どのパターンでマッチしたかも返す。数字+P（4P等）は内容量の可能性があるため、fromP=true のときは Gemini 判定を推奨。
 * @param {string} name - 商品名（タイトル）
 * @return {{ setCount: number|null, fromP: boolean }|null} 取れなければ null
 */
function parseSetCountFromItemNameWithSource(name) {
  if (!name || typeof name !== 'string') return null;
  var s = name.trim();
  var patterns = [
    { re: /×\s*(\d+)\s*個\s*セット/i, fromP: false },
    { re: /×\s*(\d+)\s*個\s*入/i, fromP: false },
    { re: /(\d+)\s*個\s*セット/i, fromP: false },
    { re: /(\d+)\s*個\s*入/i, fromP: false },
    { re: /(\d+)\s*[Pp]\s*(?:セット)?/, fromP: true }
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = s.match(patterns[i].re);
    if (m) {
      var n = parseInt(m[1], 10);
      if (n >= 1 && n <= 999) return { setCount: n, fromP: patterns[i].fromP };
    }
  }
  if (/\[1袋\]|1\s*袋|1\s*個(?:セット)?/i.test(s)) return { setCount: 1, fromP: false };
  return null;
}

/**
 * 楽天の商品名から「販売単位のセット数」を Gemini で判定する。4P が内容量かセット数かなど曖昧な場合に使用。
 * @param {string} itemName - 楽天 API の itemName
 * @return {number|null} 販売セット数（1〜999）。不明・内容量のみの場合は null
 */
function inferSetCountFromItemNameByGemini(itemName) {
  if (!itemName || typeof itemName !== 'string' || itemName.trim().length === 0) return null;
  var key = getGeminiApiKey();
  if (!key) return null;
  var prompt = '以下の楽天商品名から、**販売単位のセット数**（何個セット・何袋セットで売っているか）を1つの整数だけで答えてください。\n'
    + '・「16g×4P」「4袋入」などは内容量（1袋あたりのパック数等）なのでセット数に含めません。\n'
    + '・「2袋・5袋・10袋」のように複数オプションがある場合は、そのいずれかではなく「この1商品としての最小販売単位」が不明なら 0 と答えてください。\n'
    + '・「×3個セット」「10個セット」のように明確にセット数が一つに決まる場合はその整数を返してください。\n'
    + '回答は数字のみ（0＝不明）。説明は不要。\n\n商品名:\n' + itemName.substring(0, 500);
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (typeof MODEL_GEMINI !== 'undefined' ? MODEL_GEMINI : 'gemini-2.0-flash') + ':generateContent?key=' + encodeURIComponent(key);
    var body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1, maxOutputTokens: 64 } };
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body) });
    if (res.getResponseCode() !== 200) {
      Logger.log('[楽天セット別] inferSetCountFromItemNameByGemini HTTP ' + res.getResponseCode());
      return null;
    }
    var json = JSON.parse(res.getContentText());
    var text = (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) ? json.candidates[0].content.parts[0].text : '';
    text = String(text).replace(/\s/g, '').replace(/[^\d]/g, '');
    var n = parseInt(text, 10);
    if (isNaN(n) || n < 1 || n > 999) return null;
    Logger.log('[楽天セット別] Gemini セット数 商品名先頭40字=' + itemName.substring(0, 40) + ' → ' + n);
    return n;
  } catch (e) {
    Logger.log('[楽天セット別] inferSetCountFromItemNameByGemini 例外: ' + (e && e.message));
    return null;
  }
}

/**
 * 複数の楽天商品名を一括で Gemini に送り、文脈を踏まえて各商品の販売セット数を判定する。内容量とセット数の区別を一覧全体から推測させる。
 * @param {string[]} itemNames - 商品名の配列（同一JANで取得した一覧）
 * @return {number[]} 各商品のセット数（不明は0）。件数は itemNames.length と同一。失敗時は空配列。
 */
function inferSetCountBatchByGemini(itemNames) {
  if (!itemNames || itemNames.length === 0) return [];
  var key = getGeminiApiKey();
  if (!key) {
    Logger.log('[楽天セット数一括] Gemini API キー未設定');
    return [];
  }
  var listText = '';
  for (var i = 0; i < itemNames.length; i++) {
    listText += (i + 1) + '. ' + (itemNames[i] || '').substring(0, 300) + '\n';
  }
  var prompt = '以下は同一JANで取得した楽天の商品名一覧です。各商品について、**販売単位のセット数**（何個セット・何袋セットで売っているか）を、一覧全体の文脈（他商品の表記との整合）を踏まえて判断し、1つの整数で答えてください。\n'
    + '・「16g×4P」「4袋入」などは内容量（1袋あたりのパック数等）のことが多いです。同一シリーズに「10個セット」「2袋・5袋・10袋」などがあれば、4Pは内容量と判断してください。\n'
    + '・不明・複数オプションで特定できない場合は0。\n'
    + '・回答は1行に1つ、行番号と同順で整数のみ。説明や記号は不要。\n\n商品名一覧:\n' + listText + '\n上記1番目〜' + itemNames.length + '番目のセット数（0=不明）を、1行1整数で' + itemNames.length + '行答えてください。';
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (typeof MODEL_GEMINI !== 'undefined' ? MODEL_GEMINI : 'gemini-2.0-flash') + ':generateContent?key=' + encodeURIComponent(key);
    var body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1, maxOutputTokens: 512 } };
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body) });
    if (res.getResponseCode() !== 200) {
      Logger.log('[楽天セット数一括] Gemini HTTP ' + res.getResponseCode());
      return [];
    }
    var json = JSON.parse(res.getContentText());
    var text = (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) ? json.candidates[0].content.parts[0].text : '';
    var lines = String(text).split(/\r?\n/);
    var result = [];
    for (var j = 0; j < itemNames.length; j++) {
      var num = 0;
      if (j < lines.length) {
        var match = lines[j].replace(/\s/g, '').match(/\d+/);
        if (match) {
          var n = parseInt(match[0], 10);
          num = (n >= 1 && n <= 999) ? n : 0;
        }
      }
      result.push(num);
    }
    Logger.log('[楽天セット数一括] Gemini 一括判定 件数=' + itemNames.length + ' 取得=' + result.length);
    return result;
  } catch (e) {
    Logger.log('[楽天セット数一括] 例外: ' + (e && e.message));
    return [];
  }
}

/**
 * メニュー「Yahoo! セット別価格 取得テスト」の処理。
 * 固定 JAN で Yahoo! API を複数件取得し、商品名からセット数を推定してセット数ごとの最安価格をログ・トーストで表示する。
 */
function menuTestYahooSetPrices() {
  var clientId = (PropertiesService.getScriptProperties().getProperty('YAHOO_SHOPPING_CLIENT_ID') || '').trim();
  if (!clientId) {
    SpreadsheetApp.getUi().alert('Script Properties に YAHOO_SHOPPING_CLIENT_ID を設定してください。');
    return;
  }
  var jan = TEST_RAKUTEN_YAHOO_JAN;
  var res = fetchYahooShoppingItemsByJan(clientId, jan, 20);
  if (res.error) {
    Logger.log('[Yahoo!セット別テスト] エラー: ' + res.error);
    SpreadsheetApp.getActive().toast('エラー: ' + res.error, 'Yahoo! セット別テスト', 6);
    return;
  }
  var hits = res.hits || [];
  Logger.log('[Yahoo!セット別テスト] JAN=' + jan + ' 取得件数=' + hits.length);
  var setToMinPrice = {};
  for (var i = 0; i < hits.length; i++) {
    var name = hits[i].name;
    var price = hits[i].price;
    Logger.log('[Yahoo!セット別テスト] hit[' + i + '] name=' + (name ? name.substring(0, 60) : '') + '... price=' + price);
    var setCount = parseSetCountFromItemName(name);
    if (setCount != null && price != null && price > 0) {
      if (setToMinPrice[setCount] == null || price < setToMinPrice[setCount]) {
        setToMinPrice[setCount] = price;
      }
    }
  }
  var setCounts = Object.keys(setToMinPrice).map(function(k) { return parseInt(k, 10); }).sort(function(a, b) { return a - b; });
  var summary = 'JAN=' + jan + ' 取得=' + hits.length + '件 セット数推定=' + setCounts.length + '種';
  for (var j = 0; j < setCounts.length; j++) {
    var sc = setCounts[j];
    summary += ' ' + sc + '→' + setToMinPrice[sc] + '円';
    Logger.log('[Yahoo!セット別テスト] セット数=' + sc + ' 最安=' + setToMinPrice[sc] + '円');
  }
  Logger.log('[Yahoo!セット別テスト] ' + summary);
  SpreadsheetApp.getActive().toast(summary, 'Yahoo! セット別テスト', 10);
}

/**
 * メニュー「楽天・Yahoo 競合価格APIテスト」の処理。
 * 固定 JAN で楽天・Yahoo! を 1 回ずつ呼び、ログに価格を出しトーストで通知。マスタ書き込みはしない。
 * テスト終了後はメニュー項目を削除する想定。
 */
function menuTestRakutenYahooCompetitivePriceApi() {
  var props = PropertiesService.getScriptProperties();
  var appId = (props.getProperty('RAKUTEN_APP_ID') || '').trim();
  var accessKey = (props.getProperty('RAKUTEN_ACCESS_KEY') || '').trim();
  var yahooId = (props.getProperty('YAHOO_SHOPPING_CLIENT_ID') || '').trim();

  var jan = TEST_RAKUTEN_YAHOO_JAN;
  var logLines = [];
  logLines.push('[競合価格APIテスト] JAN=' + jan);

  // 楽天（JAN で 0 件のときだけキーワードで再検索）
  var rRakuten = fetchRakutenIchibaItemPrice(appId, accessKey, jan);
  if (rRakuten.error) {
    logLines.push('楽天: エラー ' + rRakuten.error);
    Logger.log('[競合価格APIテスト] 楽天 エラー: ' + rRakuten.error);
  } else if (rRakuten.rawCount === 0) {
    rRakuten = fetchRakutenIchibaItemPrice(appId, accessKey, TEST_RAKUTEN_YAHOO_KEYWORD_FALLBACK);
    if (rRakuten.price != null) {
      logLines.push('楽天: ' + rRakuten.price + '円（キーワード「' + TEST_RAKUTEN_YAHOO_KEYWORD_FALLBACK + '」）');
    } else {
      logLines.push('楽天: 0件');
    }
  } else {
    logLines.push('楽天: ' + rRakuten.price + '円');
  }

  // Yahoo!
  var rYahoo = fetchYahooShoppingItemPriceByJan(yahooId, jan);
  if (rYahoo.error) {
    logLines.push('Yahoo!: エラー ' + rYahoo.error);
    Logger.log('[競合価格APIテスト] Yahoo! エラー: ' + rYahoo.error);
  } else if (rYahoo.rawCount === 0) {
    logLines.push('Yahoo!: 0件');
  } else {
    logLines.push('Yahoo!: ' + rYahoo.price + '円');
  }

  var summary = logLines.join(' / ');
  Logger.log('[競合価格APIテスト] ' + summary);
  SpreadsheetApp.getActive().toast(summary, '楽天・Yahoo APIテスト', 8);
}

// ----------------------------------------
// 3.1d ASIN貼り付け（Keepa用）シートで Keepa から商品名・価格・セット数を取得
// ----------------------------------------
/** 1商品あたりの列ブロック（画像, ASIN, 商品名, 評価, 競合価格, セット数, 商品URL, セット数_AI根拠, 参考画像URL, 卸値(税込)）。ブロックとAI情報取得dataの行が対応（ブロックb＝AIの2+b行目）。 */
const ASIN_PASTE_BLOCK_HEADERS = ['画像', 'ASIN', '商品名', '評価', '競合価格(Amazon)', 'セット数', '商品URL', 'セット数_AI根拠', '参考画像URL', '卸値(税込)'];
/** シート作成時の商品ブロック数。最大10商品＝60列。1行目にAI情報取得dataの商品名を式で表示 */
const ASIN_PASTE_DEFAULT_BLOCKS = 10;
/** AI情報取得data の「商品名」「JAN」「メーカー」列（1行目の式で参照）。B列＝仕入判断は参照しない。列順変更時は要調整 */
const AI_SHEET_NAME_FOR_ASIN_PASTE = 'AI情報取得data';
const AI_PRODUCT_NAME_COL = 'D';
const AI_JAN_COL = 'C';
const AI_MAKER_COL = 'A';
/** AI情報取得data の「参考情報(画像URL)」列（1-based）。inputHeaders の9列目 */
const AI_REF_IMAGE_URL_COL = 9;
/** 画像一致率を評価に加算する際の閾値（参考画像セット数採用）。80%以上で参考画像セット数を採用。 */
const IMAGE_MATCH_THRESHOLD_REF_SET = 80;
/** 評価を◎にするための画像一致率の下限。これ未満は名前+画像で100超えても◎にせずXX%のまま（別商品の誤◎を防ぐ）。 */
const IMAGE_MATCH_MIN_FOR_CIRCLE = 70;
/** Keepa取得の行ごとログ（URL・画像紐付き確認・要確認行の追跡用）を書き出すシート名 */
const KEEPA_FETCH_LOG_SHEET_NAME = 'Keepa取得_ログ';
/** Keepa取得キャッシュ用シート名。同一ASINは取得から指定日数までは再取得せずトークン節約。Script Property KEEPA_CACHE_DAYS で日数（既定7）、KEEPA_CACHE_ENABLED=false で無効。 */
const KEEPA_CACHE_SHEET_NAME = 'Keepa取得_キャッシュ';
/** キャッシュ保持日数（Script Property 未設定時の既定値） */
const KEEPA_CACHE_DAYS_DEFAULT = 7;
/** キャッシュシートの列。Keepa取得データ（調査用）と同じ項目＋取得日時＋セット数項目。調査・回収用。 */
const KEEPA_CACHE_HEADERS = [
  '取得日時',
  '画像', '商品名', 'ASIN', '商品コード: EAN', '製造者', 'ブランド', 'アイテム数', '発売日',
  'URL: Amazon', 'URL: Keepa', 'カテゴリ: ルート', 'カテゴリ: ツリー',
  '売れ筋ランキング: 90 日平均', 'レビュー: 評価', 'レビュー: 評価件数',
  'Buy Box: 現在価格', 'Buy Box: 30 日平均', 'Buy Box: 90 日平均',
  'Amazon: 現在価格', 'Amazon: 30 日平均', 'Amazon: 90 日平均',
  '新品: 90 日平均', '参考価格: 90 日平均',
  'setCount', 'setCountFromTitle', 'setCountReason'
];
/** キャッシュの「調査用」部分のみのヘッダー（buildKeepaTableRow 用）。KEEPA_CACHE_HEADERS の取得日時と setCount 系を除く。 */
var KEEPA_CACHE_DEBUG_PART_HEADERS = [
  '画像', '商品名', 'ASIN', '商品コード: EAN', '製造者', 'ブランド', 'アイテム数', '発売日',
  'URL: Amazon', 'URL: Keepa', 'カテゴリ: ルート', 'カテゴリ: ツリー',
  '売れ筋ランキング: 90 日平均', 'レビュー: 評価', 'レビュー: 評価件数',
  'Buy Box: 現在価格', 'Buy Box: 30 日平均', 'Buy Box: 90 日平均',
  'Amazon: 現在価格', 'Amazon: 30 日平均', 'Amazon: 90 日平均',
  '新品: 90 日平均', '参考価格: 90 日平均'
];
/** 画像一致スコアの加重平均の重み（影響度が高い順：shape, package, text, color, capacity）。合計1.0 */
const IMAGE_MATCH_WEIGHTS = { shape: 0.28, package: 0.24, text: 0.22, color: 0.16, capacity: 0.10 };

/**
 * シート「ASIN貼り付け（Keepa用）」がなければ作成し、1行目に商品名の式・2行目にヘッダーを設定する。
 * 1行目：各ブロック先頭列に ='AI情報取得data'!D2,C2,A2（商品名・JAN・メーカー）を式で表示。2行目：ASIN, 商品名, 評価, 競合価格(Amazon), セット数, 商品URL×10ブロック。3行目以降にASINを貼る。B列（仕入判断）は参照しない。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureAsinPasteSheet(ss) {
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  if (sheet) return sheet;
  sheet = ss.insertSheet(ASIN_PASTE_SHEET_NAME);
  var blockCols = ASIN_PASTE_BLOCK_HEADERS.length;
  var totalCols = ASIN_PASTE_DEFAULT_BLOCKS * blockCols;
  var row1 = [];
  var row2 = [];
  for (var b = 0; b < ASIN_PASTE_DEFAULT_BLOCKS; b++) {
    var aiRow = 2 + b;
    row1.push('');
    row1.push('');
    row1.push("='" + AI_SHEET_NAME_FOR_ASIN_PASTE + "'!" + AI_PRODUCT_NAME_COL + aiRow);
    row1.push("='" + AI_SHEET_NAME_FOR_ASIN_PASTE + "'!" + AI_JAN_COL + aiRow);
    row1.push("='" + AI_SHEET_NAME_FOR_ASIN_PASTE + "'!" + AI_MAKER_COL + aiRow);
    for (var i = 5; i < blockCols; i++) row1.push('');
    for (var h = 0; h < blockCols; h++) row2.push(ASIN_PASTE_BLOCK_HEADERS[h]);
  }
  sheet.getRange(1, 1, 1, totalCols).setValues([row1]);
  for (var b = 0; b < ASIN_PASTE_DEFAULT_BLOCKS; b++) {
    sheet.getRange(1, b * blockCols + 1, 1, b * blockCols + Math.min(5, blockCols)).setFontWeight('bold');
  }
  sheet.getRange(2, 1, 1, totalCols).setValues([row2]);
  sheet.getRange(2, 1, 1, totalCols).setFontWeight('bold');
  sheet.autoResizeColumns(1, totalCols);
  return sheet;
}

/**
 * 評点式の一致率を計算する。JAN一致＝◎(100%)、否则は商品名(0-50)＋文字一致(0-50)＋メーカー一致ボーナス(最大+10)。
 * @param {string} expectedName 期待商品名（AI情報取得data）
 * @param {string} expectedJAN 期待JAN（同上。空可）
 * @param {string} expectedMaker 期待メーカー名（同上。空可）
 * @param {string} title Keepaのtitle
 * @param {string|null} keepaEan KeepaのEAN/JAN（あれば。空可）
 * @return {{ display: string, score: number }}
 */
function evalMatchScore(expectedName, expectedJAN, expectedMaker, title, keepaEan) {
  var e = (expectedName != null) ? String(expectedName).trim() : '';
  var t = (title != null) ? String(title).trim() : '';
  var norm = function(s) { return (s != null) ? String(s).replace(/\D/g, '') : ''; };
  var j1 = norm(expectedJAN);
  var j2 = norm(keepaEan);
  if (j1.length >= 10 && j2.length >= 10 && j1 === j2) return { display: '◎', score: 100 };

  var nameScore = 0;
  if (e.length > 0 && t.length > 0) {
    if (t.indexOf(e) >= 0 || e.indexOf(t) >= 0) nameScore = 50;
    else {
      var eWords = e.replace(/[\s\u3000]+/g, ' ').split(' ').filter(function(w) { return w.length >= 2; });
      var commonLen = 0;
      for (var i = 0; i < eWords.length; i++) {
        if (t.indexOf(eWords[i]) >= 0) commonLen += eWords[i].length;
      }
      nameScore = (e.length > 0) ? Math.min(50, Math.round(50 * commonLen / e.length)) : 0;
    }
  }
  var charBonus = 0;
  if (e.length > 0 && t.length > 0) {
    var seen = {};
    for (var c = 0; c < e.length; c++) {
      var ch = e.charAt(c);
      if (!seen[ch] && t.indexOf(ch) >= 0) { seen[ch] = true; charBonus++; }
    }
    charBonus = Math.min(50, charBonus);
  }
  var makerBonus = 0;
  var maker = (expectedMaker != null) ? String(expectedMaker).trim() : '';
  if (maker.length >= 1 && t.length > 0 && t.indexOf(maker) >= 0) makerBonus = 10;
  var score = Math.min(99, nameScore + charBonus + makerBonus);
  return { display: score + '%', score: score };
}

/**
 * 楽天ヒットの商品名が「期待する商品」と一致するかをスコア化する。③商品名一致スコア用。
 * 閾値・各ルールの得点は RAKUTEN_NAME_MATCH_* 定数で変更可能。ログ調査は [楽天商品名一致] で検索。
 * @param {string} expectedMaker 期待メーカー名（マスタのメーカー名。空可）
 * @param {string} expectedProductName 期待商品名（マスタの商品名ベース／商品名。空可）
 * @param {string} hitItemName 楽天APIの itemName
 * @return {{ score: number, usedRules: Array<{ rule: string, score: number }> }}
 */
function evalRakutenItemNameMatchScore(expectedMaker, expectedProductName, hitItemName) {
  var usedRules = [];
  var total = 0;
  var e = (expectedProductName != null) ? String(expectedProductName).trim() : '';
  var t = (hitItemName != null) ? String(hitItemName).trim() : '';
  var maker = (expectedMaker != null) ? String(expectedMaker).trim() : '';

  var fullScore = RAKUTEN_NAME_MATCH_FULL_MATCH_SCORE;
  var wordMax = RAKUTEN_NAME_MATCH_WORD_MAX_SCORE;
  var makerScore = RAKUTEN_NAME_MATCH_MAKER_SCORE;
  var charMax = RAKUTEN_NAME_MATCH_CHAR_MAX_SCORE;

  if (RAKUTEN_NAME_MATCH_USE_FULL && e.length > 0 && t.length > 0) {
    var eNorm = e.replace(/[\s\u3000]+/g, '');
    var tNorm = t.replace(/[\s\u3000]+/g, '');
    if (t.indexOf(e) >= 0 || e.indexOf(t) >= 0 || (eNorm.length > 0 && tNorm.length > 0 && (tNorm.indexOf(eNorm) >= 0 || eNorm.indexOf(tNorm) >= 0))) {
      total += fullScore;
      usedRules.push({ rule: '完全一致', score: fullScore });
    }
  }

  var wordScore = 0;
  if (RAKUTEN_NAME_MATCH_USE_WORD && e.length > 0 && t.length > 0) {
    var eWords = e.replace(/[\s\u3000]+/g, ' ').split(' ').filter(function (w) { return w.length >= 2; });
    var commonLen = 0;
    for (var i = 0; i < eWords.length; i++) {
      if (t.indexOf(eWords[i]) >= 0) commonLen += eWords[i].length;
    }
    wordScore = (e.length > 0) ? Math.min(wordMax, Math.round(wordMax * commonLen / e.length)) : 0;
    if (wordScore > 0) {
      total += wordScore;
      usedRules.push({ rule: '単語一致', score: wordScore });
    }
  }

  if (RAKUTEN_NAME_MATCH_USE_MAKER && maker.length >= 1 && t.length > 0 && t.indexOf(maker) >= 0) {
    total += makerScore;
    usedRules.push({ rule: 'メーカー一致', score: makerScore });
  }

  var charBonus = 0;
  if (RAKUTEN_NAME_MATCH_USE_CHAR && e.length > 0 && t.length > 0) {
    var seen = {};
    for (var c = 0; c < e.length; c++) {
      var ch = e.charAt(c);
      if (!seen[ch] && t.indexOf(ch) >= 0) { seen[ch] = true; charBonus++; }
    }
    charBonus = Math.min(charMax, charBonus);
    if (charBonus > 0) {
      total += charBonus;
      usedRules.push({ rule: '文字一致', score: charBonus });
    }
  }

  total = Math.min(100, total);
  return { score: total, usedRules: usedRules };
}

/**
 * 商品名からセット数候補をすべて抽出する（括弧内の「〇袋」と「〇袋セット」等を両方候補にする）。価格判定で最安単価に近いものを後段で採用する。
 * @param {string} title
 * @return {number[]} 抽出した数値の配列（重複除く・1以上）。0件なら []
 */
function extractSetCountCandidatesFromTitle(title) {
  if (!title || typeof title !== 'string') return [];
  var numMap = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10, '１': 1, '２': 2, '３': 3, '４': 4, '５': 5, '６': 6, '７': 7, '８': 8, '９': 9 };
  var patterns = [
    /(\d+)\s*個\s*セット/i,
    /(\d+)\s*袋\s*セット/i,
    /(\d+)\s*セット/i,
    /(\d+)\s*袋\s*×\s*(\d+)/i,
    /(\d+)\s*袋(?!\s*×)/i,
    /(三|３|五|５|十)\s*袋/i,
    /(五|５)\s*袋\s*入?/i,
    /(\d+)\s*袋\s*入/i,
    /(\d+)\s*パック/i,
    /(\d+)\s*本\s*入り/i,
    /(\d+)\s*個\s*入り/i,
    /(\d+)\s*個(?!\s*[セ入])/i,
    /×\s*(\d+)\s*袋/i,
    /(\d+)\s*P\s*入り/i,
    /(\d+)\s*P\)/i,
    /（\s*(\d+)\s*袋\s*）/,
    /[（(](\d+)[袋個][）)]/i
  ];
  var seen = {};
  var out = [];
  for (var p = 0; p < patterns.length; p++) {
    var m = title.match(patterns[p]);
    if (m) {
      for (var i = 1; i < m.length; i++) {
        if (m[i] == null) continue;
        var n = numMap[m[i]];
        if (n == null) n = parseInt(m[i], 10);
        if (n != null && !isNaN(n) && n >= 1 && !seen[n]) {
          seen[n] = true;
          out.push(n);
        }
      }
    }
  }
  return out;
}

/**
 * 商品名（title）からセット数を抽出する。まとめ売りでよく使う表現（セット・袋・個・パック・本入り・×N等）を複数パターンで判定。
 * 複数候補がある場合は先頭1件を返す。候補を全て使いたい場合は extractSetCountCandidatesFromTitle を使用。
 * @param {string} title
 * @return {number|null} 抽出できれば数値、できなければ null
 */
function extractSetCountFromTitle(title) {
  var candidates = extractSetCountCandidatesFromTitle(title);
  return candidates.length > 0 ? candidates[0] : null;
}

/**
 * Keepa API で複数 ASIN の商品名・価格・セット数を取得する。価格は csv の複数インデックス(0,7,1,2)を順に試す。
 * @param {string} apiKey
 * @param {string[]} asins
 * @return {{ list: Array<{asin:string, title:string, price:number|null, setCount:number|null, ean:string|null}>, raw: Array }|null} list=表示用、raw=調査用の生データ
 */
/**
 * Keepa product から画像URLを取得する。image がURLならそのまま、imagesCSV なら Amazon 画像URLを組み立てる。
 * @param {Object} p Keepa product
 * @return {string|null}
 */
function getKeepaProductImageUrl(p) {
  if (!p) return null;
  if (p.image && String(p.image).trim() !== '') {
    var s = String(p.image).trim();
    if (s.indexOf('http') === 0) return s;
    return 'https://m.media-amazon.com/images/I/' + s + '._AC_SL160_.jpg';
  }
  if (p.imagesCSV && String(p.imagesCSV).trim() !== '') {
    var first = String(p.imagesCSV).split(',')[0].trim();
    if (first) return 'https://m.media-amazon.com/images/I/' + first + '._AC_SL160_.jpg';
  }
  return null;
}

/**
 * Keepa の価格生値を円に換算。日本では「円」で返る場合があるため、100～500000 ならそのまま円、それ以外は 1/100 とみなす。
 * @param {number} raw Keepa から取得した価格
 * @return {number|null}
 */
function keepaPriceRawToYen(raw) {
  if (raw == null || raw === -1 || isNaN(Number(raw))) return null;
  var v = Number(raw);
  if (v >= 100 && v <= 500000) return Math.round(v);
  return Math.round(v / 100);
}

/**
 * Keepa product からカート（Buy Box）現在価格を円で取得。競合価格表示・調査用シートで共通利用。
 * 優先: offers(isBuyBox) → stats.buyBoxPrice → stats.current[10] → stats.current[0] → csv[10],csv[0],...
 * stats/csv も日本では円で返る場合があるため keepaPriceRawToYen で統一換算。
 * @param {Object} p Keepa product
 * @return {number|null} 円、無いときは null
 */
function getKeepaBuyBoxCurrentYenFromProduct(p) {
  if (!p) return null;
  var yen = getKeepaBuyBoxPriceFromOffers(p);
  if (yen != null) return yen;
  if (p.stats && p.stats.buyBoxPrice != null && p.stats.buyBoxPrice > 0) return keepaPriceRawToYen(p.stats.buyBoxPrice);
  if (p.stats && (p.stats.current || p.stats.avg)) {
    var st = p.stats.current || p.stats.avg;
    if (st[10] != null && st[10] !== -1) { yen = keepaPriceRawToYen(st[10]); if (yen != null) return yen; }
    if (st[0] != null && st[0] !== -1) { yen = keepaPriceRawToYen(st[0]); if (yen != null) return yen; }
  }
  if (p.csv) {
    var indices = (typeof KEEPA_PRICE_CSV_INDICES !== 'undefined' && KEEPA_PRICE_CSV_INDICES.length) ? KEEPA_PRICE_CSV_INDICES : [10, 0, 7, 1, 2];
    for (var k = 0; k < indices.length; k++) {
      if (p.csv[indices[k]]) {
        var arr = p.csv[indices[k]];
        for (var i = arr.length - 1; i >= 0; i -= 2) {
          if (i >= 1 && arr[i] !== -1 && arr[i] != null) { yen = keepaPriceRawToYen(arr[i]); if (yen != null) return yen; }
        }
      }
    }
  }
  return null;
}

/**
 * Keepa offers からカート（Buy Box）価格を取得。offerCSV は [..., 時刻, 価格, 送料] の並び。
 * 日本ドメインでは価格が「円」で返る場合があるため、100～500000 の範囲ならそのまま円、それ以外は 1/100 として換算。
 * @param {Object} p Keepa product（offers あり）
 * @return {number|null} 円（item + ship）
 */
function getKeepaBuyBoxPriceFromOffers(p) {
  if (!p || !p.offers || !Array.isArray(p.offers)) return null;
  var currentKeepaTime = (new Date().getTime() - new Date(2011, 0, 1).getTime()) / 60000;
  var ACTIVE_THRESHOLD = 525600;
  for (var oi = 0; oi < p.offers.length; oi++) {
    var o = p.offers[oi];
    if (!o || !o.isBuyBox || !o.offerCSV || o.offerCSV.length < 3) continue;
    var arr = o.offerCSV;
    var len = arr.length;
    var lastTime = arr[len - 3];
    var price = arr[len - 2];
    var ship = arr[len - 1];
    if (currentKeepaTime - lastTime >= ACTIVE_THRESHOLD) continue;
    if (price == null || price === -1) continue;
    var shipVal = (ship != null && ship >= 0) ? ship : 0;
    var total = Number(price) + Number(shipVal);
    if (total >= 100 && total <= 500000) return Math.round(total);
    return Math.round(total / 100);
  }
  return null;
}

function fetchKeepaProductsForSheet(apiKey, asins) {
  if (!asins || asins.length === 0) return { list: [], raw: [] };
  var url = 'https://api.keepa.com/product?key=' + encodeURIComponent(apiKey) + '&domain=' + KEEPA_DOMAIN_JP + '&asin=' + asins.join(',') + '&stats=90&offers=20';
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    var json = JSON.parse(res.getContentText());
    if (!json.products) return { list: [], raw: [] };
    var out = [];
    json.products.forEach(function(p) {
      var asin = p.asin;
      if (!asin) return;
      var title = (p.title != null) ? String(p.title) : '';
      var ean = (p.ean != null && p.ean !== '') ? String(p.ean) : null;
      var priceYen = getKeepaBuyBoxCurrentYenFromProduct(p);
      if (priceYen != null && (priceYen > 50000 || (priceYen < 10 && priceYen !== 0))) priceYen = null;
      var setCount = extractSetCountFromTitle(title);
      var setCountReason = (setCount != null) ? '正規表現' : null;
      var imageUrl = getKeepaProductImageUrl(p);
      out.push({ asin: asin, title: title, price: priceYen, setCount: setCount, setCountFromTitle: setCount, setCountReason: setCountReason, ean: ean, imageUrl: imageUrl });
    });
    if (PropertiesService.getScriptProperties().getProperty('KEEPA_USE_AI_SET_COUNT') === 'true' && getGeminiApiKey()) {
      var needAi = [];
      var needAiIdx = [];
      out.forEach(function(o, idx) {
        if (o.setCount == null && o.title) { needAi.push(o.title); needAiIdx.push(idx); }
      });
      if (needAi.length > 0) {
        var aiResult = inferSetCountsByGemini(needAi);
        if (aiResult && aiResult.setCounts && aiResult.setCounts.length === needAiIdx.length) {
          for (var ai = 0; ai < needAiIdx.length; ai++) {
            if (aiResult.setCounts[ai] != null && aiResult.setCounts[ai] > 0) {
              out[needAiIdx[ai]].setCount = aiResult.setCounts[ai];
              out[needAiIdx[ai]].setCountFromTitle = aiResult.setCounts[ai];
              out[needAiIdx[ai]].setCountReason = (aiResult.reasons && aiResult.reasons[ai]) ? aiResult.reasons[ai] : 'AI推測';
            }
          }
        }
      }
    }
    if (PropertiesService.getScriptProperties().getProperty('KEEPA_SET_COUNT_FROM_IMAGE') === 'true' && getGeminiApiKey()) {
      var imageCallLimit = 5;
      var imageCallCount = 0;
      for (var ii = 0; ii < out.length && imageCallCount < imageCallLimit; ii++) {
        if (out[ii].setCount != null || !out[ii].imageUrl) continue;
        var imgResult = inferSetCountFromImageByGemini(out[ii].imageUrl);
        if (imgResult && imgResult.setCount != null) {
          out[ii].setCount = imgResult.setCount;
          out[ii].setCountReason = imgResult.reason || '画像AI';
          imageCallCount++;
        }
        Utilities.sleep(350);
      }
    }
    return { list: out, raw: json.products };
  } catch (e) {
    console.warn('[Keepa] ' + e.message);
    return null;
  }
}

/**
 * 商品名からSEO用・目を引くための余分な文言を除き、商品名として必要な部分とセット数関連（〇袋・〇個セット等）のみ残す。Script Property KEEPA_CLEAN_TITLE_BY_AI=true で有効。
 * @param {string[]} titles 商品名の配列
 * @return {string[]|null} クリーニング後の商品名配列。失敗時は null（元のタイトルをそのまま使う）
 */
function cleanProductTitlesByGemini(titles) {
  if (!titles || titles.length === 0) return [];
  var key = getGeminiApiKey();
  if (!key) return null;
  var prompt = '以下の商品名それぞれについて、次のルールで簡潔に書き直してください。\n'
    + '【残す】メーカー名・ブランド名、商品名の本質部分、セット数や容量のヒントになる数字と単位（〇袋・〇個セット・〇P入り・〇g×〇包等）。\n'
    + '【除く】SEO対策用のキーワード、お客様の目を引くための余分な文言、重複する説明。\n'
    + '回答はJSON配列のみ。各要素は1つの文字列（書き直した商品名）。改行や説明は不要。\n\n';
  titles.forEach(function(t, i) { prompt += (i + 1) + '. ' + t + '\n'; });
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (typeof MODEL_GEMINI !== 'undefined' ? MODEL_GEMINI : 'gemini-2.0-flash') + ':generateContent?key=' + encodeURIComponent(key);
    var body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, maxOutputTokens: 2048 } };
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body) });
    if (res.getResponseCode() !== 200) return null;
    var json = JSON.parse(res.getContentText());
    var text = (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) ? json.candidates[0].content.parts[0].text : '';
    text = text.replace(/```\w*\n?/g, '').trim();
    var start = text.indexOf('[');
    var end = text.lastIndexOf(']');
    if (start < 0 || end <= start) return null;
    var arr = JSON.parse(text.substring(start, end + 1));
    if (!Array.isArray(arr) || arr.length !== titles.length) return null;
    return arr.map(function(x) { return (x != null ? String(x).trim() : '') || ''; });
  } catch (e) {
    console.warn('[Keepa cleanTitle] ' + (e && e.message));
    return null;
  }
}

/**
 * 商品画像URLから、まとめ売りの数量（袋・個・セット等）をGemini Visionで推測する。
 * Script Property KEEPA_SET_COUNT_FROM_IMAGE=true かつ GEMINI_API_KEY があるとき有効。
 * @param {string} imageUrl 商品画像のURL（Keepaの画像URL等）
 * @return {{ setCount: number|null, reason: string }|null} 失敗時は null
 */
function inferSetCountFromImageByGemini(imageUrl) {
  if (!imageUrl || typeof imageUrl !== 'string' || imageUrl.indexOf('http') !== 0) return null;
  var key = getGeminiApiKey();
  if (!key) return null;
  var blob;
  try {
    var resp = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; rv:91.0) Gecko/20100101 Firefox/91.0' } });
    if (resp.getResponseCode() !== 200) return null;
    blob = resp.getBlob();
  } catch (e) {
    return null;
  }
  if (!blob || blob.getBytes().length > 4 * 1024 * 1024) return null;
  var base64 = Utilities.base64Encode(blob.getBytes());
  var mime = (blob.getContentType() && blob.getContentType().indexOf('png') >= 0) ? 'image/png' : 'image/jpeg';
  var prompt = 'この商品画像だけを見て、まとめ売りの数量（何袋・何個セット・何本入り等）を1つの整数だけで答えてください。\n'
    + '【判定のポイント】\n'
    + '・同じ形のパッケージ（スタンドパウチ・袋・ボトル等）が何個写っているか数える。形が同じものをまとめ売りの個数とする。\n'
    + '・重なって写っている場合も、同じ商品のパッケージとして数える。\n'
    + '・「〇袋入」「〇pack」「〇個セット」など画像内の文字で数量が書いてあればそれを優先する。「四袋入」等は1袋あたりの内容なので、外側の袋数と区別する。\n'
    + '・明らかに色・デザインが他と違う別商品が写っている場合は除外し、同じ商品のまとめ売りだけを数える。\n'
    + '該当なし・不明は0。説明は不要。整数1つのみ答えてください。';
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (typeof MODEL_GEMINI !== 'undefined' ? MODEL_GEMINI : 'gemini-2.0-flash') + ':generateContent?key=' + encodeURIComponent(key);
    var body = {
      contents: [{ parts: [{ inlineData: { mimeType: mime, data: base64 } }, { text: prompt }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 32 }
    };
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body) });
    if (res.getResponseCode() !== 200) return null;
    var json = JSON.parse(res.getContentText());
    var text = (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) ? json.candidates[0].content.parts[0].text : '';
    text = String(text).trim().replace(/\D/g, '');
    var n = text ? parseInt(text, 10) : 0;
    if (isNaN(n) || n < 0) n = 0;
    return { setCount: n > 0 ? n : null, reason: n > 0 ? '画像AI' : '' };
  } catch (e) {
    return null;
  }
}

/**
 * AI情報取得data の指定ブロック（0-based）に対応する行の「参考情報(画像URL)」を取得する。
 * ブロック b は AI の行 2+b に対応する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {number} blockIndex ブロック番号（0-based）
 * @return {string} URL または空文字
 */
function getRefImageUrlForAsinPasteBlock(ss, blockIndex) {
  if (!ss || blockIndex < 0) return '';
  var aiSheet = ss.getSheetByName(AI_SHEET_NAME_FOR_ASIN_PASTE);
  if (!aiSheet) return '';
  var row = 2 + blockIndex;
  var col = typeof AI_REF_IMAGE_URL_COL !== 'undefined' ? AI_REF_IMAGE_URL_COL : 9;
  var val = aiSheet.getRange(row, col).getValue();
  return (val != null && typeof val === 'string') ? String(val).trim() : '';
}

/**
 * AI情報取得data の指定ブロック（0-based）に対応する行の「卸値(税込)」を取得する。ブロック b は AI の行 2+b に対応。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {number} blockIndex ブロック番号（0-based）
 * @return {number|null} 卸値(税込)の数値、または取得できないとき null
 */
function getCostTaxInForAsinPasteBlock(ss, blockIndex) {
  if (!ss || blockIndex < 0) return null;
  var aiSheet = ss.getSheetByName(AI_SHEET_NAME_FOR_ASIN_PASTE);
  if (!aiSheet) return null;
  var headerRow = aiSheet.getRange(1, 1, 1, 50).getValues()[0] || [];
  var col = -1;
  for (var c = 0; c < headerRow.length; c++) {
    if (String(headerRow[c]).trim() === '卸値(税込)') { col = c + 1; break; }
  }
  if (col < 0) return null;
  var row = 2 + blockIndex;
  var val = aiSheet.getRange(row, col).getValue();
  if (val == null || val === '') return null;
  if (typeof val === 'number' && !isNaN(val)) return val >= 0 ? val : null;
  var s = String(val).replace(/,/g, '').trim();
  var n = parseFloat(s, 10);
  return (!isNaN(n) && n >= 0) ? n : null;
}

/**
 * 2枚の画像URLを Gemini Vision に渡し、形・色・文字・パッケージ形状・容量表記の一致率（0-100）と総合スコアを返す。
 * 背景は無視し商品部分のみを比較するようプロンプトで指示する。
 * @param {string} refImageUrl 参考画像（AI情報取得data の参考情報(画像URL)）
 * @param {string} keepaImageUrl Keepa の商品画像URL
 * @return {{ shape: number, color: number, text: number, package: number, capacity: number, overall: number }|null} 失敗時は null
 */
function getImageMatchScoreByGemini(refImageUrl, keepaImageUrl) {
  if (!refImageUrl || !keepaImageUrl || refImageUrl.indexOf('http') !== 0 || keepaImageUrl.indexOf('http') !== 0) {
    Logger.log('[getImageMatchScoreByGemini] URL無効 ref=' + (refImageUrl || '(空)') + ' keepa=' + (keepaImageUrl || '(空)'));
    return null;
  }
  var key = getGeminiApiKey();
  if (!key) {
    Logger.log('[getImageMatchScoreByGemini] GeminiAPIキー未設定');
    return null;
  }
  function fetchImageAsBase64(url, label) {
    try {
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36' }, followRedirects: true });
      var code = resp.getResponseCode();
      if (code !== 200) {
        Logger.log('[getImageMatchScoreByGemini] ' + label + ' HTTP ' + code + ' URL=' + url.substring(0, 80));
        return null;
      }
      var blob = resp.getBlob();
      var size = blob ? blob.getBytes().length : 0;
      if (!blob || size === 0) {
        Logger.log('[getImageMatchScoreByGemini] ' + label + ' blob空 URL=' + url.substring(0, 80));
        return null;
      }
      if (size > 4 * 1024 * 1024) {
        Logger.log('[getImageMatchScoreByGemini] ' + label + ' 画像大きすぎ ' + size + 'bytes URL=' + url.substring(0, 80));
        return null;
      }
      Logger.log('[getImageMatchScoreByGemini] ' + label + ' OK ' + size + 'bytes contentType=' + blob.getContentType());
      return { b64: Utilities.base64Encode(blob.getBytes()), mime: (blob.getContentType() && blob.getContentType().indexOf('png') >= 0) ? 'image/png' : 'image/jpeg' };
    } catch (e) {
      Logger.log('[getImageMatchScoreByGemini] ' + label + ' 例外: ' + (e && e.message) + ' URL=' + url.substring(0, 80));
      return null;
    }
  }
  Logger.log('[getImageMatchScoreByGemini] 参考URL=' + refImageUrl.substring(0, 80) + ' KeepaURL=' + keepaImageUrl.substring(0, 80));
  var img1 = fetchImageAsBase64(refImageUrl, '参考画像');
  var img2 = fetchImageAsBase64(keepaImageUrl, 'Keepa画像');
  if (!img1 || !img2) {
    Logger.log('[getImageMatchScoreByGemini] 画像取得失敗 img1=' + (img1 ? 'OK' : 'null') + ' img2=' + (img2 ? 'OK' : 'null'));
    return null;
  }
  var prompt = '2枚の商品画像です。背景は無視し、商品部分のみを比較してください。\n'
    + '以下の5項目それぞれについて、一致度を0～100の整数で答えてください。\n'
    + '1. shape: 商品の形・シルエットの類似度（撮影角度・回転の違いは減点しない。同一商品なら高く評価する）\n2. color: 色・配色の一致度\n3. text: パッケージやラベルの文字・ロゴの一致度\n4. package: パッケージ形状（袋/ボトル/箱等）の一致度（角度・向きの違いは減点しない）\n5. capacity: 容量・本数らしき表記の一致度\n'
    + '回答は必ず1行で、JSONオブジェクトのみを出力してください。説明や改行は不要。例: {"shape":80,"color":70,"text":60,"package":90,"capacity":50}';

  function parseGeminiScoreText(text) {
    if (!text || typeof text !== 'string') return null;
    text = String(text).replace(/```\w*\n?/g, '').trim();
    var mShape = text.match(/"shape":\s*(\d+)/);
    var mColor = text.match(/"color":\s*(\d+)/);
    var mText  = text.match(/"text":\s*(\d+)/);
    var mPkg   = text.match(/"package":\s*(\d+)/);
    var mCap   = text.match(/"capacity":\s*(\d+)/);
    function buildFromRegex() {
      if (!mShape || !mColor || !mText) return null;
      var shape = Math.min(100, Math.max(0, parseInt(mShape[1], 10) || 0));
      var color = Math.min(100, Math.max(0, parseInt(mColor[1], 10) || 0));
      var textScore = Math.min(100, Math.max(0, parseInt(mText[1], 10) || 0));
      var packageScore = mPkg ? Math.min(100, Math.max(0, parseInt(mPkg[1], 10) || 0)) : 50;
      var capacity = mCap ? Math.min(100, Math.max(0, parseInt(mCap[1], 10) || 0)) : 50;
      var w = typeof IMAGE_MATCH_WEIGHTS !== 'undefined' ? IMAGE_MATCH_WEIGHTS : { shape: 0.25, color: 0.25, text: 0.2, package: 0.15, capacity: 0.15 };
      var overall = Math.round(shape * w.shape + color * w.color + textScore * w.text + packageScore * w.package + capacity * w.capacity);
      overall = Math.min(100, Math.max(0, overall));
      return { shape: shape, color: color, text: textScore, package: packageScore, capacity: capacity, overall: overall };
    }
    var start = text.indexOf('{');
    if (start >= 0) {
      var depth = 0, end = -1;
      for (var i = start; i < text.length; i++) {
        if (text.charAt(i) === '{') depth++;
        else if (text.charAt(i) === '}') { depth--; if (depth === 0) { end = i; break; } }
      }
      if (end >= start) {
        try {
          var o = JSON.parse(text.substring(start, end + 1));
          var shape = Math.min(100, Math.max(0, parseInt(o.shape, 10) || 0));
          var color = Math.min(100, Math.max(0, parseInt(o.color, 10) || 0));
          var textScore = Math.min(100, Math.max(0, parseInt(o.text, 10) || 0));
          var packageScore = Math.min(100, Math.max(0, parseInt(o.package, 10) || 0));
          var capacity = Math.min(100, Math.max(0, parseInt(o.capacity, 10) || 0));
          var w = typeof IMAGE_MATCH_WEIGHTS !== 'undefined' ? IMAGE_MATCH_WEIGHTS : { shape: 0.25, color: 0.25, text: 0.2, package: 0.15, capacity: 0.15 };
          var overall = Math.round(shape * w.shape + color * w.color + textScore * w.text + packageScore * w.package + capacity * w.capacity);
          overall = Math.min(100, Math.max(0, overall));
          return { shape: shape, color: color, text: textScore, package: packageScore, capacity: capacity, overall: overall };
        } catch (parseErr) { /* fall through to regex */ }
      }
    }
    return buildFromRegex();
  }

  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (typeof MODEL_GEMINI !== 'undefined' ? MODEL_GEMINI : 'gemini-2.0-flash') + ':generateContent?key=' + encodeURIComponent(key);
    var parts = [
      { inlineData: { mimeType: img1.mime, data: img1.b64 } },
      { inlineData: { mimeType: img2.mime, data: img2.b64 } },
      { text: prompt }
    ];
    var body = { contents: [{ parts: parts }], generationConfig: { temperature: 0.1, maxOutputTokens: 1024 } };
    var lastText = '';
    for (var attempt = 1; attempt <= 2; attempt++) {
      var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body) });
      if (res.getResponseCode() !== 200) {
        Logger.log('[getImageMatchScoreByGemini] Gemini API HTTP ' + res.getResponseCode() + ' body=' + res.getContentText().substring(0, 300));
        return null;
      }
      var apiJson = JSON.parse(res.getContentText());
      lastText = (apiJson.candidates && apiJson.candidates[0] && apiJson.candidates[0].content && apiJson.candidates[0].content.parts && apiJson.candidates[0].content.parts[0]) ? apiJson.candidates[0].content.parts[0].text : '';
      lastText = String(lastText).replace(/```\w*\n?/g, '').trim();
      Logger.log('[getImageMatchScoreByGemini] Gemini応答(' + attempt + '回目)=' + lastText.substring(0, 200));
      var result = parseGeminiScoreText(lastText);
      if (result) {
        Logger.log('[getImageMatchScoreByGemini] shape=' + result.shape + ' color=' + result.color + ' text=' + result.text + ' package=' + result.package + ' capacity=' + result.capacity + ' overall=' + result.overall);
        return result;
      }
      if (attempt === 1) {
        Logger.log('[getImageMatchScoreByGemini] JSON解析失敗のため2.5秒後にリトライします');
        Utilities.sleep(2500);
      }
    }
    Logger.log('[getImageMatchScoreByGemini] JSON解析失敗: 応答にJSONなし text=' + lastText.substring(0, 300));
    return null;
  } catch (e) {
    Logger.log('[getImageMatchScoreByGemini] ' + (e && e.message));
    return null;
  }
}

/**
 * 参考画像（AI情報取得data の参考情報(画像URL)）から、まとめ売りのセット数を Gemini Vision で推測する。
 * @param {string} refImageUrl 参考画像のURL
 * @return {{ setCount: number|null, reason: string }|null}
 */
function inferSetCountFromReferenceImageByGemini(refImageUrl) {
  return inferSetCountFromImageByGemini(refImageUrl);
}

/**
 * 商品名の配列から、まとめ売りのセット数（袋・個・セット等）をGeminiで推測する。Script Property KEEPA_USE_AI_SET_COUNT=true で有効。
 * 推測値と根拠（短い理由）を返す。
 * @param {string[]} titles 商品名の配列
 * @return {{ setCounts: number[], reasons: string[] }|null} 失敗時は null
 */
function inferSetCountsByGemini(titles) {
  if (!titles || titles.length === 0) return { setCounts: [], reasons: [] };
  var key = getGeminiApiKey();
  if (!key) return null;
  var prompt = '以下の商品名それぞれについて、まとめ売りのセット数（何袋・何個セット・何本入り等）を推測し、整数と短い根拠を返してください。\n'
    + '回答はJSON配列のみ。各要素は {"set": 数値, "reason": "根拠（例: 5袋の表記）"} の形。該当なし・不明は {"set":0,"reason":""}。改行や説明文は不要。\n\n';
  titles.forEach(function(t, i) { prompt += (i + 1) + '. ' + t + '\n'; });
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (typeof MODEL_GEMINI !== 'undefined' ? MODEL_GEMINI : 'gemini-2.0-flash') + ':generateContent?key=' + encodeURIComponent(key);
    var body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1, maxOutputTokens: 1024 } };
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body) });
    if (res.getResponseCode() !== 200) return null;
    var json = JSON.parse(res.getContentText());
    var text = (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) ? json.candidates[0].content.parts[0].text : '';
    text = text.replace(/```\w*\n?/g, '').trim();
    var start = text.indexOf('[');
    var end = text.lastIndexOf(']');
    if (start < 0 || end <= start) return null;
    var arr = JSON.parse(text.substring(start, end + 1));
    if (!Array.isArray(arr)) return null;
    var setCounts = [];
    var reasons = [];
    for (var i = 0; i < arr.length; i++) {
      var o = arr[i];
      if (o && typeof o === 'object') {
        var n = parseInt(o.set != null ? o.set : o, 10);
        setCounts.push(isNaN(n) || n < 0 ? 0 : n);
        reasons.push((o.reason != null ? String(o.reason).trim() : '') || '');
      } else {
        setCounts.push(0);
        reasons.push('');
      }
    }
    return { setCounts: setCounts, reasons: reasons };
  } catch (e) {
    console.warn('[Keepa AI setCount] ' + (e && e.message));
    return null;
  }
}

/** Keepa 取得生データを書き出す調査用シート名（テスト・調査用）。項目ごとに列で表形式で出力。 */
const KEEPA_DEBUG_SHEET_NAME = 'Keepa取得データ（調査用）';
/** 手動エクスポート(KeepaExport-2026-02-23.csv)の列名に合わせた主要列。差分確認用。 */
var KEEPA_DEBUG_HEADERS_MANUAL_CSV = [
  '画像', '商品名', 'ASIN', '商品コード: EAN', '製造者', 'ブランド', 'アイテム数', '発売日',
  'URL: Amazon', 'URL: Keepa', 'カテゴリ: ルート', 'カテゴリ: ツリー',
  '売れ筋ランキング: 90 日平均', 'レビュー: 評価', 'レビュー: 評価件数',
  'Buy Box: 現在価格', 'Buy Box: 30 日平均', 'Buy Box: 90 日平均',
  'Amazon: 現在価格', 'Amazon: 30 日平均', 'Amazon: 90 日平均',
  '新品: 90 日平均', '参考価格: 90 日平均'
];
/** Keepa取得項目一覧シート名（どの項目をどこから取得しているか英日・取得元を一覧表示）。 */
const KEEPA_FIELD_LIST_SHEET_NAME = 'Keepa取得項目一覧';
/** Keepa取得実行後の診断ログを追記するシート（結果を残して差分分析用）。 */
const KEEPA_DIAGNOSTIC_SHEET_NAME = 'Keepa取得_診断';
/** document is not defined 原因追究用の診断ログを書き出すシート名 */
const DOCUMENT_ERROR_DIAG_SHEET_NAME = 'documentエラー診断ログ';

/**
 * Keepa 手動DLと一致する項目一覧。[日本語名, 英語名, 取得元, 備考]
 * 日本語名は手動エクスポート KeepaExport-2026-02-23.csv の列名に合わせる（差分比較用）。
 */
var KEEPA_FIELD_MAPPING = [
  ['画像', 'Image', 'product.image / product.imagesCSV', ''],
  ['商品名', 'Title', 'product.title', '手動CSV=商品名'],
  ['ASIN', 'ASIN', 'product.asin', ''],
  ['商品コード: EAN', '商品コード: EAN', 'product.ean', '手動CSV=商品コード: EAN'],
  ['製造者', 'Manufacturer', 'product.manufacturer', '手動CSV=製造者'],
  ['ブランド', 'Brand', 'product.brand', ''],
  ['アイテム数', 'アイテム数', 'product.numberOfItems', ''],
  ['発売日', '発売日', 'product.releaseDate', ''],
  ['URL: Amazon', 'URL: Amazon', '構築: https://www.amazon.co.jp/dp/{ASIN}', ''],
  ['URL: Keepa', 'URL: Keepa', '構築: https://keepa.com/#!product/5-{ASIN}', ''],
  ['カテゴリ: ルート', 'Categories: Root', 'product.categoryTree', ''],
  ['カテゴリ: サブ', 'Categories: Sub', 'product.categoryTree', ''],
  ['カテゴリ: ツリー', 'Categories: Tree', 'product.categoryTree', ''],
  ['売れ筋ランキング: 90 日平均', 'Sales Rank: 90 days avg.', 'stats.avg[3] / csv[3]', ''],
  ['売れ筋ランキング: 365 日平均', 'Sales Rank: 365 days avg.', 'stats', ''],
  ['レビュー: 評価', 'Reviews: Rating', 'csv[16] 直近/10', ''],
  ['レビュー: 評価件数', 'Reviews: Review Count', 'csv[17] 直近', ''],
  ['Buy Box: 現在価格', 'Buy Box: Current', 'offers(isBuyBox) / stats.buyBoxPrice / stats.current[10] / csv[10]', '競合価格で使用'],
  ['Buy Box: 30 日平均', 'Buy Box: 30 days avg.', 'stats.avg[10]', ''],
  ['Buy Box: 90 日平均', 'Buy Box: 90 days avg.', 'stats.avg[10]', ''],
  ['Buy Box: 90日間の下落 %', 'Buy Box: 90 days drop %', 'stats', ''],
  ['Amazon: 現在価格', 'Amazon: Current', 'csv[0] 直近 / stats.current[0]', ''],
  ['Amazon: 30 日平均', 'Amazon: 30 days avg.', 'stats.avg[0]', ''],
  ['Amazon: 90 日平均', 'Amazon: 90 days avg.', 'stats.avg[0]', ''],
  ['Amazon: 365 日平均', 'Amazon: 365 days avg.', 'stats', ''],
  ['Amazon: 90日間の下落 %', 'Amazon: 90 days drop %', 'stats', ''],
  ['新品: 現在価格', 'New: Current', 'csv[1] / stats', ''],
  ['新品: 90 日平均', 'New: 90 days avg.', 'stats.avg[1]', ''],
  ['新品: 365 日平均', 'New: 365 days avg.', 'stats', ''],
  ['参考価格: 90 日平均', 'List Price: 90 days avg.', 'stats.avg[4] / csv[4]', ''],
  ['参考価格: 365 日平均', 'List Price: 365 days avg.', 'stats', ''],
  ['新品出品者数90日平均', 'New Offer Count: 90 days avg.', 'stats', ''],
  ['取得済みライブオファー数（新品FBA）', 'Count of retrieved live offers: New, FBA', 'offers', ''],
  ['取得済みライブオファー数（新品FBM）', 'Count of retrieved live offers: New, FBM', 'offers', ''],
  ['梱包サイズ長さ（cm）', 'Package: Length (cm)', 'product.packageLength', ''],
  ['梱包サイズ幅（cm）', 'Package: Width (cm)', 'product.packageWidth', ''],
  ['梱包サイズ高さ（cm）', 'Package: Height (cm)', 'product.packageHeight', ''],
  ['梱包重量（g）', 'Package: Weight (g)', 'product.packageWeight', ''],
  ['梱包数量', 'Package: Quantity', 'product.packageQuantity', ''],
  ['商品サイズ長さ（cm）', 'Item: Length (cm)', 'product.itemLength', ''],
  ['商品サイズ幅（cm）', 'Item: Width (cm)', 'product.itemWidth', ''],
  ['商品サイズ高さ（cm）', 'Item: Height (cm)', 'product.itemHeight', ''],
  ['商品重量（g）', 'Item: Weight (g)', 'product.itemWeight', ''],
  ['定期おトク便', 'Subscribe and Save', 'API要確認', ''],
  ['ワンタイムクーポン・定期おトク便割引率', 'One Time Coupon: Subscribe & Save %', 'API要確認', ''],
  ['ビジネス割引', 'Business Discount', 'API要確認', ''],
  ['Buy Box: % Amazon 30 日', 'Buy Box: % Amazon 30 days', 'stats', ''],
  ['Buy Box: % Amazon 90 日', 'Buy Box: % Amazon 90 days', 'stats', ''],
  ['Buy Box: % トップセラー 30 日', 'Buy Box: % Top Seller 30 days', 'stats', ''],
  ['Buy Box: % トップセラー 90 日', 'Buy Box: % Top Seller 90 days', 'stats', ''],
  ['Buy Box: 勝者数 30 日', 'Buy Box: Winner Count 30 days', 'stats', ''],
  ['Buy Box: 勝者数 90 日', 'Buy Box: Winner Count 90 days', 'stats', ''],
  ['Buy Box: 標準偏差 30 日', 'Buy Box: Standard Deviation 30 days', 'stats', ''],
  ['Buy Box: 標準偏差 90 日', 'Buy Box: Standard Deviation 90 days', 'stats', ''],
  ['Buy Box: 変動性 30 日', 'Buy Box: Flipability 30 days', 'stats', ''],
  ['Buy Box: 変動性 90 日', 'Buy Box: Flipability 90 days', 'stats', ''],
  ['梱包サイズ長さ（cm）', 'Package: Length (cm)', 'product.packageLength', ''],
  ['梱包サイズ幅（cm）', 'Package: Width (cm)', 'product.packageWidth', ''],
  ['梱包サイズ高さ（cm）', 'Package: Height (cm)', 'product.packageHeight', ''],
  ['梱包重量（g）', 'Package: Weight (g)', 'product.packageWeight', ''],
  ['梱包数量', 'Package: Quantity', 'product.packageQuantity', ''],
  ['商品サイズ長さ（cm）', 'Item: Length (cm)', 'product.itemLength', ''],
  ['商品サイズ幅（cm）', 'Item: Width (cm)', 'product.itemWidth', ''],
  ['商品サイズ高さ（cm）', 'Item: Height (cm)', 'product.itemHeight', ''],
  ['商品重量（g）', 'Item: Weight (g)', 'product.itemWeight', ''],
  ['定期おトク便', 'Subscribe and Save', 'API要確認', ''],
  ['ワンタイムクーポン・定期おトク便割引率', 'One Time Coupon: Subscribe & Save %', 'API要確認', ''],
  ['ビジネス割引', 'Business Discount', 'API要確認', ''],
  ['カート獲得率Amazon30日', 'Buy Box: % Amazon 30 days', 'stats', ''],
  ['カート獲得率Amazon90日', 'Buy Box: % Amazon 90 days', 'stats', ''],
  ['カート獲得率Amazon180日', 'Buy Box: % Amazon 180 days', 'stats', ''],
  ['カート獲得率Amazon365日', 'Buy Box: % Amazon 365 days', 'stats', ''],
  ['カート獲得率トップセラー30日', 'Buy Box: % Top Seller 30 days', 'stats', ''],
  ['カート獲得率トップセラー90日', 'Buy Box: % Top Seller 90 days', 'stats', ''],
  ['カート獲得率トップセラー180日', 'Buy Box: % Top Seller 180 days', 'stats', ''],
  ['カート獲得率トップセラー365日', 'Buy Box: % Top Seller 365 days', 'stats', ''],
  ['カート獲得セラー数30日', 'Buy Box: Winner Count 30 days', 'stats', ''],
  ['カート獲得セラー数90日', 'Buy Box: Winner Count 90 days', 'stats', ''],
  ['カート獲得セラー数180日', 'Buy Box: Winner Count 180 days', 'stats', ''],
  ['カート獲得セラー数365日', 'Buy Box: Winner Count 365 days', 'stats', ''],
  ['カート価格標準偏差30日', 'Buy Box: Standard Deviation 30 days', 'stats', ''],
  ['カート価格標準偏差90日', 'Buy Box: Standard Deviation 90 days', 'stats', ''],
  ['カート価格標準偏差365日', 'Buy Box: Standard Deviation 365 days', 'stats', ''],
  ['カート価格流動性30日', 'Buy Box: Flipability 30 days', 'stats', ''],
  ['カート価格流動性90日', 'Buy Box: Flipability 90 days', 'stats', ''],
  ['カート価格流動性365日', 'Buy Box: Flipability 365 days', 'stats', '']
];

/** Keepa csv 配列のインデックスと日本語ラベル（0=Amazon, 1=新品, 2=中古, 3=セールスランク, 4=参考価格, 7=新品FBM送料込み 等） */
var KEEPA_CSV_INDEX_LABELS = {
  0: 'Amazon現在価格',
  1: '新品価格',
  2: '中古価格',
  3: 'セールスランク',
  4: '参考価格',
  5: 'Warehouse',
  6: '新品(6)',
  7: '新品FBM送料込み',
  8: 'Lightning Deal',
  9: 'Sales Rank',
  10: 'Buy Box（カート価格）',
  11: 'New',
  12: 'Count New',
  13: 'Count Used',
  14: 'Trade In',
  15: 'Rental'
};

/**
 * Keepa の csv[i] 配列から直近の値を取得する。配列は [時刻0, 値0, 時刻1, 値1, ...]。-1 は欠損。
 * @param {Array} arr csv[i]
 * @param {boolean} isPrice 価格系なら true（Keepaは1/100で格納するため100で割る）
 * @return {number|string|null}
 */
function getKeepaCsvLatestValue(arr, isPrice) {
  if (!arr || !Array.isArray(arr) || arr.length < 2) return null;
  for (var i = arr.length - 1; i >= 1; i -= 2) {
    var v = arr[i];
    if (v != null && v !== -1) {
      if (isPrice) return Math.round(Number(v) / 100);
      return v;
    }
  }
  return null;
}

/**
 * Keepa API の product から1行分の配列を組み立てる（表の列順に合わせる）。
 * @param {Object} p product
 * @param {string} asin
 * @param {string[]} headers ヘッダー配列（この順で値を並べる）
 * @return {Array}
 */
function buildKeepaTableRow(p, asin, headers) {
  var row = [];
  var csv = (p && p.csv) ? p.csv : {};
  var stats = (p && p.stats) ? p.stats : {};
  var avg = stats.avg || stats.avg90;
  var priceIndices = { 0: true, 1: true, 2: true, 4: true, 6: true, 7: true, 10: true };
  function statAvg(idx) {
    if (!avg || avg[idx] == null || avg[idx] === -1) return '';
    return priceIndices[idx] ? String(Math.round(Number(avg[idx]) / 100)) : String(avg[idx]);
  }
  for (var h = 0; h < headers.length; h++) {
    var key = headers[h];
    var val = '';
    if (key === '画像') {
      val = (p && (p.image || p.imagesCSV)) ? (p.image || String(p.imagesCSV).split(',')[0] || '') : '';
    } else if (key === 'タイトル' || key === '商品名') {
      val = (p && p.title) ? String(p.title) : '';
    } else if (key === 'ASIN') {
      val = asin || (p && p.asin) || '';
    } else if (key === '商品コード（EAN）' || key === '商品コード: EAN') {
      val = (p && p.ean) ? String(p.ean) : '';
    } else if (key === 'メーカー' || key === '製造者') {
      val = (p && p.manufacturer) ? String(p.manufacturer) : '';
    } else if (key === 'ブランド') {
      val = (p && p.brand) ? String(p.brand) : '';
    } else if (key === 'アイテム数') {
      val = (p && (p.numberOfItems != null || p.itemNumber != null)) ? String(p.numberOfItems != null ? p.numberOfItems : p.itemNumber) : '';
    } else if (key === '発売日') {
      val = (p && p.releaseDate) ? String(p.releaseDate) : '';
    } else if (key === 'AmazonURL' || key === 'URL: Amazon') {
      val = asin ? 'https://www.amazon.co.jp/dp/' + asin : '';
    } else if (key === 'KeepaURL' || key === 'URL: Keepa') {
      val = asin ? 'https://keepa.com/#!product/5-' + asin : '';
    } else if (key === 'カテゴリ（ルート）' || key === 'カテゴリ: ルート' || key === 'カテゴリ（ツリー）' || key === 'カテゴリ: ツリー') {
      if (p && p.categoryTree && Array.isArray(p.categoryTree) && p.categoryTree.length > 0) {
        val = p.categoryTree.map(function(c) {
          if (c == null) return '';
          return (typeof c === 'object' && c.name != null) ? c.name : String(c);
        }).filter(Boolean).join(' > ');
      } else if (p && p.categoryTree) {
        val = typeof p.categoryTree === 'string' ? p.categoryTree : String(p.categoryTree);
      }
    } else if (key === 'Buy Box: 現在価格') {
      var buyBoxYen = getKeepaBuyBoxCurrentYenFromProduct(p);
      val = (buyBoxYen != null) ? String(buyBoxYen) : '';
    } else if (key === 'Buy Box: 30 日平均' || key === 'Buy Box: 90 日平均') {
      val = statAvg(10);
    } else if (key === 'Amazon: 現在価格') {
      var v0 = csv[0] ? getKeepaCsvLatestValue(csv[0], true) : null;
      val = (v0 != null) ? String(v0) : '';
    } else if (key === 'Amazon: 30 日平均' || key === 'Amazon: 90 日平均') {
      val = statAvg(0);
    } else if (key === '新品: 90 日平均') {
      val = statAvg(1);
    } else if (key === '参考価格: 90 日平均') {
      val = statAvg(4);
    } else if (key === '売れ筋ランキング: 90 日平均') {
      val = statAvg(3);
    } else if (key === 'レビュー: 評価') {
      if (csv[16] && csv[16].length >= 2) {
        var r = csv[16][csv[16].length - 1];
        val = (r != null && r !== -1) ? String(Math.round(Number(r) / 10) / 10) : '';
      }
    } else if (key === 'レビュー: 評価件数') {
      if (csv[17] && csv[17].length >= 2) {
        var c = csv[17][csv[17].length - 1];
        val = (c != null && c !== -1) ? String(c) : '';
      }
    } else if (key.indexOf('csv_') === 0) {
      var idx = parseInt(key.replace('csv_', '').split('(')[0], 10);
      if (!isNaN(idx) && csv[idx]) {
        var v = getKeepaCsvLatestValue(csv[idx], priceIndices[idx] === true);
        val = (v != null && v !== '') ? String(v) : '';
      }
    } else {
      if (p && p[key] != null) val = String(p[key]);
    }
    row.push(val === undefined ? '' : val);
  }
  return row;
}

/**
 * Keepa API の生データを「項目ごとの表」で調査用シートに書き出す。1行目＝ヘッダー、2行目以降＝1商品1行。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string[]} asins 今回取得したASIN一覧（順序保持）
 * @param {Array} rawProducts Keepa API の json.products
 */
function writeKeepaRawToDebugSheet(ss, asins, rawProducts) {
  var sheet = ss.getSheetByName(KEEPA_DEBUG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(KEEPA_DEBUG_SHEET_NAME);
  }
  var asinMap = {};
  rawProducts.forEach(function(p) { if (p.asin) asinMap[p.asin] = p; });
  var csvIndices = [];
  rawProducts.forEach(function(p) {
    if (p && p.csv && typeof p.csv === 'object') {
      for (var k in p.csv) if (p.csv.hasOwnProperty(k)) {
        var n = parseInt(k, 10);
        if (!isNaN(n) && csvIndices.indexOf(n) < 0) csvIndices.push(n);
      }
    }
  });
  csvIndices.sort(function(a, b) { return a - b; });
  var labels = (typeof KEEPA_CSV_INDEX_LABELS !== 'undefined') ? KEEPA_CSV_INDEX_LABELS : {};
  var headers = (typeof KEEPA_DEBUG_HEADERS_MANUAL_CSV !== 'undefined' && KEEPA_DEBUG_HEADERS_MANUAL_CSV.length) ? KEEPA_DEBUG_HEADERS_MANUAL_CSV.slice() : [
    '画像', '商品名', 'ASIN', '商品コード: EAN', '製造者', 'ブランド', 'アイテム数', '発売日',
    'URL: Amazon', 'URL: Keepa', 'カテゴリ: ルート', 'カテゴリ: ツリー',
    '売れ筋ランキング: 90 日平均', 'レビュー: 評価', 'レビュー: 評価件数',
    'Buy Box: 現在価格', 'Buy Box: 30 日平均', 'Buy Box: 90 日平均',
    'Amazon: 現在価格', 'Amazon: 30 日平均', 'Amazon: 90 日平均', '新品: 90 日平均', '参考価格: 90 日平均'
  ];
  csvIndices.forEach(function(i) {
    headers.push('csv_' + i + (labels[i] ? '(' + labels[i] + ')' : ''));
  });
  var rows = [headers];
  for (var a = 0; a < asins.length; a++) {
    var p = asinMap[asins[a]];
    rows.push(buildKeepaTableRow(p, asins[a], headers));
  }
  sheet.clear();
  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Keepa取得実行後の診断ログを「Keepa取得_診断」シートに1行追記。結果を残して次回AIに貼り付けて差分分析できるようにする。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} info { totalWritten, evalTop5, titleCleanedRan, setReasonCount, note }
 */
function writeKeepaDiagnosticSheet(ss, info) {
  var sheet = ss.getSheetByName(KEEPA_DIAGNOSTIC_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KEEPA_DIAGNOSTIC_SHEET_NAME);
  var headers = ['実行日時', '書き込み件数', '評価の先頭5件(期待=降順)', '商品名簡潔化', 'セット数根拠件数', '備考'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  var row = [
    info.date ? (info.date instanceof Date ? info.date.toLocaleString('ja-JP') : String(info.date)) : new Date().toLocaleString('ja-JP'),
    info.totalWritten != null ? info.totalWritten : '',
    info.evalTop5 != null ? String(info.evalTop5) : '',
    info.titleCleanedRan ? '実行' : '未実行',
    info.setReasonCount != null ? info.setReasonCount : '',
    info.note != null ? String(info.note) : ''
  ];
  sheet.appendRow(row);
}

/**
 * Keepa取得の行ごとログ用シートを用意する（URL・画像紐付き確認・要確認行の追跡用）。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function ensureKeepaFetchLogSheet(ss) {
  var sheet = ss.getSheetByName(KEEPA_FETCH_LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KEEPA_FETCH_LOG_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    var headers = ['日時', 'ブロック', '行', 'ASIN', '画像URL(先頭80文字)', '商品名(先頭50文字)', '評価', '期待名(先頭30文字)', '備考', '名前評価', '画像overall', '画像shape', '画像color', '画像text', '画像package', '画像capacity'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

/**
 * Keepa取得の1行分のログを「Keepa取得_ログ」シートに追記する。URL・画像紐付き・画像一致内訳の原因追及に利用。
 * rowData に nameScore, imageOverall, imageShape, imageColor, imageText, imagePackage, imageCapacity があれば追記する。
 */
function appendKeepaFetchLogRow(ss, rowData) {
  var sheet = ss.getSheetByName(KEEPA_FETCH_LOG_SHEET_NAME);
  if (!sheet) return;
  var img = (rowData.imageUrl != null) ? String(rowData.imageUrl).substring(0, 80) : '';
  var title = (rowData.title != null) ? String(rowData.title).substring(0, 50) : '';
  var exp = (rowData.expectedName != null) ? String(rowData.expectedName).substring(0, 30) : '';
  var note = (rowData.score != null && rowData.score < 30) ? '要確認(評価低)' : '';
  var row = [
    rowData.date || new Date().toLocaleString('ja-JP'),
    rowData.block != null ? rowData.block : '',
    rowData.rowIndex != null ? rowData.rowIndex : '',
    rowData.asin || '',
    img,
    title,
    rowData.evalDisplay || '',
    exp,
    note
  ];
  row.push(rowData.nameScore != null ? rowData.nameScore : '');
  row.push(rowData.imageOverall != null ? rowData.imageOverall : '');
  row.push(rowData.imageShape != null ? rowData.imageShape : '');
  row.push(rowData.imageColor != null ? rowData.imageColor : '');
  row.push(rowData.imageText != null ? rowData.imageText : '');
  row.push(rowData.imagePackage != null ? rowData.imagePackage : '');
  row.push(rowData.imageCapacity != null ? rowData.imageCapacity : '');
  sheet.appendRow(row);
}

/**
 * Keepa取得キャッシュ用シートを取得または作成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getKeepaCacheSheet(ss) {
  if (!ss) return null;
  var sheet = ss.getSheetByName(KEEPA_CACHE_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KEEPA_CACHE_SHEET_NAME);
  if (sheet.getLastRow() === 0 && typeof KEEPA_CACHE_HEADERS !== 'undefined') {
    sheet.getRange(1, 1, 1, KEEPA_CACHE_HEADERS.length).setValues([KEEPA_CACHE_HEADERS]);
    sheet.getRange(1, 1, 1, KEEPA_CACHE_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * キャッシュシートのうち、指定日数より古い行を削除する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {number} days 保持日数（この日数より古い行を削除）
 * @return {number} 削除した行数
 */
function purgeKeepaCacheOlderThanDays(ss, days) {
  var sheet = getKeepaCacheSheet(ss);
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  var cutoff = new Date().getTime() - days * 24 * 60 * 60 * 1000;
  var data = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues();
  var toDelete = [];
  for (var r = 0; r < data.length; r++) {
    var val = data[r][0];
    var t = (typeof val === 'number') ? val : (typeof val === 'string' ? parseFloat(val, 10) : NaN);
    if (!isNaN(t) && t < cutoff) toDelete.push(r + 2);
  }
  if (toDelete.length === 0) return 0;
  for (var i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
  Logger.log('[Keepa] キャッシュ削除: ' + toDelete.length + ' 行（' + days + '日超過）');
  return toDelete.length;
}

/**
 * 指定ASINのうち、キャッシュにあり且つ有効期限内のデータを返す。
 * キャッシュシートが旧形式（9列）の場合は従来の列位置で読む。新形式（調査用項目）の場合はヘッダー順で読む。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string[]} asins ASINの配列
 * @param {number} days 有効日数（この日数以内の取得のみ有効）
 * @return {Object.<string, {asin:string, title:string, price:number|null, imageUrl:string, setCount:number|null, setCountFromTitle:number|null, setCountReason:string|null, ean:string|null}>} ASINをキーにしたキャッシュオブジェクト
 */
function getKeepaCachedResults(ss, asins, days) {
  var out = {};
  if (!asins || asins.length === 0) return out;
  var sheet = getKeepaCacheSheet(ss);
  if (!sheet || sheet.getLastRow() <= 1) return out;
  var cutoff = new Date().getTime() - days * 24 * 60 * 60 * 1000;
  var asinSet = {};
  asins.forEach(function(a) { if (a) asinSet[a] = true; });
  var numCols = sheet.getLastColumn();
  var rows = sheet.getRange(2, 1, sheet.getLastRow(), numCols).getValues();
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var isOldFormat = (numCols <= 10) || ((row[24] == null || row[24] === '') && (row[25] == null || row[25] === '') && (row[26] == null || row[26] === ''));
    var idxDate = 0, idxAsin, idxTitle, idxPrice, idxImage, idxSetCount, idxSetCountFromTitle, idxSetCountReason, idxEan;
    if (isOldFormat) {
      idxAsin = 1; idxTitle = 2; idxPrice = 3; idxImage = 4; idxSetCount = 5; idxSetCountFromTitle = 6; idxSetCountReason = 7; idxEan = 8;
    } else {
      if (typeof KEEPA_CACHE_HEADERS !== 'undefined') {
        idxAsin = KEEPA_CACHE_HEADERS.indexOf('ASIN');
        idxTitle = KEEPA_CACHE_HEADERS.indexOf('商品名');
        idxPrice = KEEPA_CACHE_HEADERS.indexOf('Buy Box: 現在価格');
        idxImage = KEEPA_CACHE_HEADERS.indexOf('画像');
        idxSetCount = KEEPA_CACHE_HEADERS.indexOf('setCount');
        idxSetCountFromTitle = KEEPA_CACHE_HEADERS.indexOf('setCountFromTitle');
        idxSetCountReason = KEEPA_CACHE_HEADERS.indexOf('setCountReason');
        idxEan = KEEPA_CACHE_HEADERS.indexOf('商品コード: EAN');
      } else {
        idxAsin = 3; idxTitle = 2; idxPrice = 16; idxImage = 1; idxSetCount = 24; idxSetCountFromTitle = 25; idxSetCountReason = 26; idxEan = 4;
      }
    }
    var t = (typeof row[idxDate] === 'number') ? row[idxDate] : parseFloat(row[idxDate], 10);
    if (isNaN(t) || t < cutoff) continue;
    var asin = String((row[idxAsin] != null) ? row[idxAsin] : '').trim();
    if (!asinSet[asin]) continue;
    var priceVal = (row[idxPrice] !== '' && row[idxPrice] != null) ? (typeof row[idxPrice] === 'number' ? row[idxPrice] : parseFloat(String(row[idxPrice]).replace(/,/g, ''), 10)) : null;
    if (typeof priceVal === 'number' && isNaN(priceVal)) priceVal = null;
    var setCount = (row[idxSetCount] !== '' && row[idxSetCount] != null) ? (typeof row[idxSetCount] === 'number' ? row[idxSetCount] : parseInt(row[idxSetCount], 10)) : null;
    if (typeof setCount === 'number' && isNaN(setCount)) setCount = null;
    var setCountFromTitle = (row[idxSetCountFromTitle] !== '' && row[idxSetCountFromTitle] != null) ? (typeof row[idxSetCountFromTitle] === 'number' ? row[idxSetCountFromTitle] : parseInt(row[idxSetCountFromTitle], 10)) : null;
    if (typeof setCountFromTitle === 'number' && isNaN(setCountFromTitle)) setCountFromTitle = null;
    out[asin] = {
      asin: asin,
      title: (row[idxTitle] != null) ? String(row[idxTitle]) : '',
      price: priceVal,
      imageUrl: (row[idxImage] != null) ? String(row[idxImage]) : '',
      setCount: setCount,
      setCountFromTitle: setCountFromTitle,
      setCountReason: (row[idxSetCountReason] != null && row[idxSetCountReason] !== '') ? String(row[idxSetCountReason]) : null,
      ean: (row[idxEan] != null && row[idxEan] !== '') ? String(row[idxEan]) : null
    };
  }
  return out;
}

/**
 * API取得結果をキャッシュシートに書き込む。既存の同ASIN行があればその行を更新、なければ末尾に追記する。
 * 調査用シートと同じ項目を保存し後の調査・回収に利用する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Array<{asin:string, title:string, price:number|null, imageUrl:string, setCount:number|null, setCountFromTitle:number|null, setCountReason:string|null, ean:string|null}>} results fetchKeepaProductsForSheet の list
 * @param {Array|null} rawProducts Keepa API の json.products。省略時は調査用列は空で setCount 等のみ保存（互換用）
 */
function writeKeepaCache(ss, results, rawProducts) {
  if (!results || results.length === 0) return;
  var sheet = getKeepaCacheSheet(ss);
  if (!sheet) return;
  var numCols = typeof KEEPA_CACHE_HEADERS !== 'undefined' ? KEEPA_CACHE_HEADERS.length : 27;
  var lastRow = sheet.getLastRow();
  var asinToRow = {};
  if (lastRow >= 2) {
    var asinColIndex = (typeof KEEPA_CACHE_HEADERS !== 'undefined') ? KEEPA_CACHE_HEADERS.indexOf('ASIN') : 3;
    if (asinColIndex < 0) asinColIndex = 3;
    var cacheData = sheet.getRange(2, 1, lastRow, numCols).getValues();
    for (var d = 0; d < cacheData.length; d++) {
      var asin = String((cacheData[d][asinColIndex] != null) ? cacheData[d][asinColIndex] : '').trim();
      if (asin) asinToRow[asin] = d + 2;
    }
  }
  var nextAppendRow = lastRow + 1;
  var now = new Date().getTime();
  var headersBuild = (typeof KEEPA_DEBUG_HEADERS_MANUAL_CSV !== 'undefined' && KEEPA_DEBUG_HEADERS_MANUAL_CSV && KEEPA_DEBUG_HEADERS_MANUAL_CSV.length) ? KEEPA_DEBUG_HEADERS_MANUAL_CSV.slice() : (typeof KEEPA_CACHE_DEBUG_PART_HEADERS !== 'undefined' ? KEEPA_CACHE_DEBUG_PART_HEADERS.slice() : []);
  var asinToRaw = {};
  if (rawProducts && rawProducts.length > 0) {
    rawProducts.forEach(function(p) { if (p && p.asin) asinToRaw[p.asin] = p; });
  }
  var asinsToWrite = [];
  var rows = [];
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    if (!r || !r.asin) continue;
    var debugRow = [];
    if (headersBuild.length > 0) {
      var p = asinToRaw[r.asin];
      if (p) {
        debugRow = buildKeepaTableRow(p, r.asin, headersBuild);
      } else {
        for (var h = 0; h < headersBuild.length; h++) debugRow.push('');
        if (debugRow.length >= 1) debugRow[0] = (r.imageUrl != null) ? String(r.imageUrl) : '';
        if (debugRow.length >= 2) debugRow[1] = (r.title != null) ? String(r.title) : '';
        if (debugRow.length >= 4) debugRow[3] = (r.ean != null) ? String(r.ean) : '';
        if (debugRow.length >= 16) debugRow[15] = (r.price != null) ? String(r.price) : '';
      }
    }
    var setCountPart = [(r.setCount != null) ? r.setCount : '', (r.setCountFromTitle != null) ? r.setCountFromTitle : '', (r.setCountReason != null) ? String(r.setCountReason) : ''];
    var fullRow = [now].concat(debugRow).concat(setCountPart);
    var idxImage = (typeof KEEPA_CACHE_HEADERS !== 'undefined') ? KEEPA_CACHE_HEADERS.indexOf('画像') : 1;
    if (idxImage >= 0 && fullRow.length > idxImage && r.imageUrl != null && String(r.imageUrl).trim().indexOf('http') === 0) {
      fullRow[idxImage] = String(r.imageUrl).trim();
    }
    asinsToWrite.push(r.asin);
    rows.push(fullRow);
  }
  if (rows.length > 0) {
    var updated = 0;
    var appended = 0;
    Logger.log('[Keepa] キャッシュ書込: results=' + (results ? results.length : 0) + ', 書込行数=' + rows.length + ', 既存ASINマップ件数=' + Object.keys(asinToRow).length);
    for (var i = 0; i < rows.length; i++) {
      if (rows[i] == null || !Array.isArray(rows[i])) {
        Logger.log('[Keepa] キャッシュ 警告: 行' + i + ' が不正 (null or 非配列)');
        continue;
      }
      if (rows[i].length !== numCols) {
        Logger.log('[Keepa] キャッシュ 警告: 行' + i + ' の列数=' + rows[i].length + ', 期待=' + numCols);
      }
      try {
        var rawRow = rows[i];
        for (var c = 0; c < rawRow.length; c++) {
          if (Array.isArray(rawRow[c])) {
            Logger.log('[Keepa] キャッシュ 行' + i + ' 列' + c + ' が配列: length=' + rawRow[c].length);
          }
        }
        var flatRow = [];
        for (var c = 0; c < numCols; c++) {
          var cell = rawRow[c];
          flatRow.push(Array.isArray(cell) ? (cell.length ? cell.join(' ') : '') : (cell != null ? cell : ''));
        }
        var asin = asinsToWrite[i];
        var targetRow = asinToRow[asin];
        if (targetRow != null) {
          sheet.getRange(targetRow, 1, 1, numCols).setValues([flatRow]);
          updated++;
        } else {
          sheet.getRange(nextAppendRow, 1, 1, numCols).setValues([flatRow]);
          asinToRow[asin] = nextAppendRow;
          nextAppendRow++;
          appended++;
        }
      } catch (e) {
        Logger.log('[Keepa] キャッシュ 行' + i + ' 書込失敗: ' + (e && e.message) + ' ASIN=' + (asinsToWrite[i] || ''));
        throw e;
      }
    }
    Logger.log('[Keepa] キャッシュ書込完了: 更新=' + updated + ', 追記=' + appended);
  }
}

/**
 * Keepa取得項目一覧シートを作成・更新する。どの項目をどこから取得しているか（日本語・英語・取得元）を表で表示。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function writeKeepaFieldListSheet(ss) {
  var sheet = ss.getSheetByName(KEEPA_FIELD_LIST_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KEEPA_FIELD_LIST_SHEET_NAME);
  var mapping = (typeof KEEPA_FIELD_MAPPING !== 'undefined' && KEEPA_FIELD_MAPPING.length) ? KEEPA_FIELD_MAPPING : [];
  var headers = ['項目名（日本語）', '項目名（英語）', '取得元（Keepa API）', '備考'];
  var rows = [headers];
  for (var i = 0; i < mapping.length; i++) {
    rows.push(mapping[i].length >= 4 ? mapping[i] : [mapping[i][0] || '', mapping[i][1] || '', mapping[i][2] || '', mapping[i][3] || '']);
  }
  sheet.clear();
  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, 4).setValues(rows);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.autoResizeColumns(1, 4);
  }
}

/** 単品単価の許容幅（上下1段のセット数を含めたグループ内の最低単価からの変動）。0.3 = ±30% */
var SET_COUNT_UNIT_PRICE_TOLERANCE = 0.3;

/** 画像一致率がこの値以上なら単価照合で画像セット数を優先しうる。 */
var SET_COUNT_IMAGE_PRIORITY_THRESHOLD = 95;

/**
 * ブロック内の行データについて、商品名候補と画像セット数から「ブロック内最安単価に最も近い」候補を採用する。
 * 候補が複数ある場合は単価が minInGroup に最も近いものを採用。画像一致率が高い場合は画像を優先しうる。
 * 卸値(税込)が指定されている場合、単価が卸値の0.5倍未満または2倍超の候補Cは「異常」として除外する。
 * @param {Array<{rowIndex:number, asin?:string, price:number|null, setCountName:number|null, setCountNameCandidates:number[], setCountImage:number|null, setCountCurrent:number|null, setCountReasonCurrent:string, imageOverall?:number}>} blockData
 * @param {number|null} costTaxInBlock ブロックの卸値(税込)。指定時は単価がこの0.5倍未満または2倍超の候補を除外
 * @return {Array<{setCount:number|null, setCountReason:string}>} 各行の採用セット数と理由（blockDataと同じ順）
 */
function resolveSetCountByUnitPrice(blockData, costTaxInBlock) {
  if (!blockData || blockData.length === 0) return [];
  var tol = (typeof SET_COUNT_UNIT_PRICE_TOLERANCE !== 'undefined' ? SET_COUNT_UNIT_PRICE_TOLERANCE : 0.3);
  var imgPriorityTh = (typeof SET_COUNT_IMAGE_PRIORITY_THRESHOLD !== 'undefined' ? SET_COUNT_IMAGE_PRIORITY_THRESHOLD : 95);
  var costMin = (costTaxInBlock != null && costTaxInBlock > 0) ? costTaxInBlock * 0.5 : null;
  var costMax = (costTaxInBlock != null && costTaxInBlock > 0) ? costTaxInBlock * 2 : null;
  var result = [];

  function getCandidates(row) {
    var cand = row.setCountNameCandidates && row.setCountNameCandidates.length > 0
      ? row.setCountNameCandidates.slice()
      : (row.setCountName != null ? [row.setCountName] : []);
    if (row.setCountImage != null && cand.indexOf(row.setCountImage) < 0) cand.push(row.setCountImage);
    return cand;
  }

  function isUnitPriceAbnormal(price, setCount) {
    if (costMin == null || costMax == null || !price || setCount < 1) return false;
    var u = price / setCount;
    return u < costMin || u > costMax;
  }

  function filterCandidatesByCost(row, candidates) {
    if (costMin == null || costMax == null || !candidates.length) return candidates;
    var price = (row.price != null && row.price > 0) ? row.price : null;
    if (!price) return candidates;
    var out = [];
    for (var c = 0; c < candidates.length; c++) {
      var sc = candidates[c];
      if (sc >= 1 && !isUnitPriceAbnormal(price, sc)) out.push(sc);
    }
    return out.length > 0 ? out : [];
  }

  var minInGroup = null;
  for (var r = 0; r < blockData.length; r++) {
    var rr = blockData[r];
    var p = (rr.price != null && rr.price > 0) ? rr.price : null;
    if (!p) continue;
    var cand = getCandidates(rr);
    cand = filterCandidatesByCost(rr, cand);
    for (var c = 0; c < cand.length; c++) {
      var sc = cand[c];
      if (sc >= 1) {
        var u = p / sc;
        if (costMin != null && costMax != null && (u < costMin || u > costMax)) continue;
        minInGroup = (minInGroup == null || u < minInGroup) ? u : minInGroup;
      }
    }
  }
  Logger.log('[Keepa] セット数照合 minInGroup(ブロック内最安単価)=' + (minInGroup != null ? Math.round(minInGroup) : 'null') + (costTaxInBlock != null ? ' 卸値(税込)=' + costTaxInBlock + ' 許容0.5~2倍' : ''));

  for (var i = 0; i < blockData.length; i++) {
    var row = blockData[i];
    var rawCandidates = getCandidates(row);
    var candidates = filterCandidatesByCost(row, rawCandidates);
    var price = (row.price != null && row.price > 0) ? row.price : null;

    if (candidates.length === 0) {
      if (rawCandidates.length > 0) {
        // 卸値範囲外でもセット数は書く。取得情報（商品名由来）を優先し、参考画像のセット数は使わない
        var setCountOut = (row.setCountNameCandidates && row.setCountNameCandidates.length > 0) ? row.setCountNameCandidates[0] : rawCandidates[0];
        result.push({ setCount: setCountOut >= 1 ? setCountOut : null, setCountReason: '卸値範囲外' });
      } else {
        result.push({ setCount: row.setCountCurrent, setCountReason: row.setCountReasonCurrent || '' });
      }
      continue;
    }
    if (candidates.length === 1) {
      var use = candidates[0];
      result.push({ setCount: use, setCountReason: row.setCountReasonCurrent || (use === row.setCountImage ? '参考画像' : '正規表現') });
      continue;
    }

    if (price == null) {
      result.push({ setCount: candidates[0], setCountReason: row.setCountReasonCurrent || '正規表現' });
      continue;
    }

    var imageOverall = (row.imageOverall != null) ? row.imageOverall : 0;
    var preferImage = (imageOverall >= imgPriorityTh && row.setCountImage != null && candidates.indexOf(row.setCountImage) >= 0);
    var unitImage = row.setCountImage != null ? price / row.setCountImage : null;
    var withinImage = (preferImage && unitImage != null && minInGroup != null && minInGroup > 0 &&
      unitImage >= minInGroup * (1 - tol) && unitImage <= minInGroup * (1 + tol));

    if (withinImage) {
      Logger.log('[Keepa] セット数照合 行' + row.rowIndex + ' ASIN=' + (row.asin || '') + ' 画像優先(overall=' + imageOverall + ') 採用=' + row.setCountImage + ' 単価=' + Math.round(unitImage));
      result.push({ setCount: row.setCountImage, setCountReason: '参考画像' });
      continue;
    }

    var bestCount = candidates[0];
    var bestDiff = null;
    for (var j = 0; j < candidates.length; j++) {
      var sc = candidates[j];
      var u = price / sc;
      var d = minInGroup != null ? Math.abs(u - minInGroup) : 0;
      if (bestDiff == null || d < bestDiff) {
        bestDiff = d;
        bestCount = sc;
      }
    }
    var reason = (bestCount === row.setCountImage) ? '単価照合(画像)' : '単価照合(名)';
    Logger.log('[Keepa] セット数照合 行' + row.rowIndex + ' ASIN=' + (row.asin || '') + ' price=' + price + ' 候補=' + candidates.join(',') + ' minInGroup=' + (minInGroup != null ? Math.round(minInGroup) : '') + ' 採用=' + bestCount + ' ' + reason);
    result.push({ setCount: bestCount, setCountReason: reason });
  }
  return result;
}

/**
 * ASIN貼り付けシートの指定ブロック（1塊）を「評価」列で降順ソートする。
 * A列を含まない範囲では Range.sort() が失敗するため、getValues/getFormulas→ソート→書き戻しで行う。
 * 式セル（=IMAGE 等）は getFormulas で取得し、書き戻し時に式を維持する（setValues のみだと式が消えてサービスエラーになる）。
 */
function sortAsinPasteSheetBlockByEvalDesc(sheet, firstRow1Based, numRows, colEval0BasedInBlock, startCol1Based, endCol1Based) {
  if (numRows <= 1 || colEval0BasedInBlock < 0) return;
  var startCol = startCol1Based != null ? startCol1Based : 1;
  var numCols = (endCol1Based != null && startCol1Based != null) ? (endCol1Based - startCol + 1) : (endCol1Based || 20);
  var maxRow = sheet.getMaxRows();
  var maxCol = sheet.getMaxColumns();
  var numRowsClamped = Math.min(numRows, maxRow - firstRow1Based + 1);
  var numColsClamped = Math.min(numCols, maxCol - startCol + 1);
  if (numRowsClamped < 1 || numColsClamped < 1) return;

  var range = sheet.getRange(firstRow1Based, startCol, numRowsClamped, numColsClamped);
  var values = range.getValues();
  var formulas = range.getFormulas();
  if (values.length <= 1) return;

  function evalSortKey(val) {
    if (val == null || val === '') return -1;
    var s = String(val).trim();
    if (s === '◎') return 100;
    var m = s.match(/^(\d+)%?$/);
    return m ? parseInt(m[1], 10) : -1;
  }
  var colIdx = colEval0BasedInBlock;
  var rowIndices = [];
  for (var r = 0; r < values.length; r++) rowIndices.push(r);
  rowIndices.sort(function(a, b) {
    var ka = evalSortKey(values[a][colIdx]);
    var kb = evalSortKey(values[b][colIdx]);
    return kb - ka;
  });
  var sortedValues = [];
  var sortedFormulas = [];
  for (var i = 0; i < rowIndices.length; i++) {
    var idx = rowIndices[i];
    sortedValues.push(values[idx].slice());
    sortedFormulas.push(formulas[idx].slice());
  }
  for (var row = 0; row < sortedValues.length; row++) {
    for (var col = 0; col < sortedValues[row].length; col++) {
      var f = sortedFormulas[row][col];
      if (f && typeof f === 'string' && f.trim().indexOf('=') === 0) {
        sortedValues[row][col] = f;
      }
    }
  }
  range.setValues(sortedValues);
}

/**
 * 【② 出品用】ASIN貼り付け（Keepa用）シートがなければ作成し、そのシートに切り替える。既にある場合は切り替えるだけ。
 */
function menuPrepareAsinPasteSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  if (!sheet) {
    sheet = ensureAsinPasteSheet(ss);
    sheet.activate();
    SpreadsheetApp.getUi().alert(
      '「' + ASIN_PASTE_SHEET_NAME + '」シートを作成しました。\n\n' +
      '1行目：AI情報取得dataの商品名・JAN・メーカーが計算式で表示されます。\n' +
      '2行目：ヘッダー。3行目以降に各ブロックの1列目（A列＝1商品目、G列＝2商品目…）へASINを縦に1行1件で貼り付けてから、メニューで「Keepa取得」を実行してください。'
    );
    return;
  }
  sheet.activate();
  SpreadsheetApp.getActive().toast('「' + ASIN_PASTE_SHEET_NAME + '」シートに切り替えました。', '準備', 3);
}

/**
 * Keepa取得項目一覧シートを作成・表示する。どの項目をどこから取得しているか（日本語・英語・取得元）を確認できる。
 */
function menuShowKeepaFieldList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  writeKeepaFieldListSheet(ss);
  var sheet = ss.getSheetByName(KEEPA_FIELD_LIST_SHEET_NAME);
  if (sheet) sheet.activate();
  SpreadsheetApp.getActive().toast('「' + KEEPA_FIELD_LIST_SHEET_NAME + '」を表示しました。', 'Keepa取得項目一覧', 3);
}

/**
 * 【② 出品用】ASIN貼り付け（Keepa用）シートの ASIN 列を読み、指定件数だけ Keepa で取得してシートに書き込む。
 * @param {number} limit 0 のときは全件
 */
function runKeepaFetchAsinPasteSheet(limit) {
  try {
    runKeepaFetchAsinPasteSheetImpl(limit);
  } catch (e) {
    var msg = (e && e.message) ? String(e.message) : String(e);
    var stack = (e && e.stack) ? String(e.stack) : '';
    Logger.log('[runKeepaFetchAsinPasteSheet] ' + msg + '\n' + stack);
    SpreadsheetApp.getUi().alert('Keepa取得でエラーが発生しました。\n\n' + msg + (stack ? '\n\n--- スタック（原因特定用）---\n' + stack : ''));
    throw e;
  }
}

function runKeepaFetchAsinPasteSheetImpl(limit) {
  var key = PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY');
  if (!key || key.trim() === '') {
    SpreadsheetApp.getUi().alert('Script Properties に KEEPA_API_KEY を設定してください。');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  if (!sheet) {
    sheet = ensureAsinPasteSheet(ss);
    sheet.activate();
    SpreadsheetApp.getUi().alert(
      '「' + ASIN_PASTE_SHEET_NAME + '」シートを作成しました。\n\n' +
      '1行目：AI情報取得dataの商品名が計算式で表示されます（そのまま利用してください）。\n' +
      '2行目：ヘッダー。3行目以降に各ブロックの1列目（A列＝1商品目、G列＝2商品目、M列＝…）へASINを縦に1行1件で貼り付けてから、メニューで「Keepa取得」を実行してください。'
    );
    return;
  }
  if (ss.getActiveSheet().getSheetName() !== ASIN_PASTE_SHEET_NAME) {
    sheet.activate();
    SpreadsheetApp.getActive().toast('「' + ASIN_PASTE_SHEET_NAME + '」に切り替えました。ASINを貼り付けてからKeepa取得を実行してください。', '準備', 5);
    return;
  }
  var data = sheet.getDataRange().getValues();
  var headerRowIndex = -1;
  var asinColsFound = [];
  for (var r = 0; r < Math.min(6, data.length); r++) {
    var row = data[r] || [];
    asinColsFound = [];
    for (var c = 0; c < row.length; c++) {
      if (String(row[c]).trim() === 'ASIN') asinColsFound.push(c);
    }
    if (asinColsFound.length > 0) {
      headerRowIndex = r;
      break;
    }
  }
  var productNameRow = (headerRowIndex > 0) ? (data[0] || []) : [];
  var headers = (headerRowIndex >= 0) ? (data[headerRowIndex] || []) : [];
  var dataStartRow = headerRowIndex + 1;
  if (headerRowIndex < 0 || asinColsFound.length === 0) {
    Logger.log('[Keepa] ヘッダー検出失敗: 「ASIN」を含む行がありません。先頭6行の列: ' + JSON.stringify(data.slice(0, 6).map(function(row) { return (row || []).slice(0, 8); })));
    SpreadsheetApp.getUi().alert('「ASIN」というヘッダーが見つかりません。1行目または2行目に「ASIN」列があるシートにしてください。');
    return;
  }
  function colByName(names) {
    for (var n = 0; n < names.length; n++) {
      for (var c = 0; c < headers.length; c++) {
        if (String(headers[c]).trim() === names[n]) return c;
      }
    }
    return -1;
  }
  var colImage = colByName(['画像']);
  var colTitle = colByName(['商品名']);
  var colPrice = colByName(['競合価格(Amazon)', '競合価格']);
  var colSet = colByName(['セット数']);
  var colEval = colByName(['評価']);
  var colUrl = colByName(['商品URL']);
  var colSetReason = colByName(['セット数_AI根拠']);
  Logger.log('[Keepa] レイアウト: ヘッダー行=' + (headerRowIndex + 1) + ', ASIN列=' + asinColsFound.join(',') + ', 画像=' + colImage + ', 商品名=' + colTitle + ', 競合価格=' + colPrice + ', セット数=' + colSet + ', 評価=' + colEval + ', URL=' + colUrl + ', セット数_AI根拠=' + colSetReason);
  var runStartTime = new Date().toLocaleString('ja-JP');
  ensureKeepaFetchLogSheet(ss);
  var cacheEnabled = (PropertiesService.getScriptProperties().getProperty('KEEPA_CACHE_ENABLED') !== 'false');
  var cacheDays = parseInt(PropertiesService.getScriptProperties().getProperty('KEEPA_CACHE_DAYS'), 10);
  if (isNaN(cacheDays) || cacheDays < 1) cacheDays = (typeof KEEPA_CACHE_DAYS_DEFAULT !== 'undefined' ? KEEPA_CACHE_DAYS_DEFAULT : 7);
  if (cacheEnabled) purgeKeepaCacheOlderThanDays(ss, cacheDays);
  var asinCount = 0;
  for (var r = dataStartRow; r < data.length; r++) {
    for (var bi = 0; bi < asinColsFound.length; bi++) {
      var a = String((data[r] || [])[asinColsFound[bi]] || '').trim().replace(/[\s\u3000]/g, '');
      if (a.length >= 10) asinCount++;
    }
  }
  SpreadsheetApp.getActive().toast('ヘッダー' + (headerRowIndex + 1) + '行目・ASIN列検出済み。対象' + asinCount + '件です。', 'Keepa取得', 3);
  var skipRefImage = (PropertiesService.getScriptProperties().getProperty('KEEPA_SKIP_REF_IMAGE_AND_MATCH') === 'true');
  if (getGeminiApiKey() && !skipRefImage) {
    var missingRef = [];
    for (var bi = 0; bi < asinColsFound.length; bi++) {
      var asinRowsB = [];
      for (var r = dataStartRow; r < data.length; r++) {
        var a = String((data[r] || [])[asinColsFound[bi]] || '').trim().replace(/[\s\u3000]/g, '');
        if (a.length >= 10) asinRowsB.push(r);
      }
      if (asinRowsB.length > 0) {
        var refUrl = getRefImageUrlForAsinPasteBlock(ss, bi);
        if (!refUrl || refUrl === '') missingRef.push('ブロック' + (bi + 1));
      }
    }
    if (missingRef.length > 0) {
      SpreadsheetApp.getUi().alert('参考情報(画像URL)が未入力のブロックがあります。\n\n' + missingRef.join('、') + ' に対応する「AI情報取得data」の行に、対象品の商品画像URLを入力してから再度実行してください。');
      return;
    }
  }
  var totalWritten = 0;
  var lastTitleCleanedRan = false;
  var totalSetReasonWritten = 0;
  for (var b = 0; b < asinColsFound.length; b++) {
    var colAsin = asinColsFound[b];
    var blockStartCol = colAsin - 1;
    var colImageB = blockStartCol;
    var colTitleB = blockStartCol + 2;
    var colEvalB = blockStartCol + 3;
    var colPriceB = blockStartCol + 4;
    var colSetB = blockStartCol + 5;
    var colUrlB = blockStartCol + 6;
    var colSetReasonB = blockStartCol + 7;
    var colRefImageUrlB = blockStartCol + 8;
    var colCostTaxInB = blockStartCol + 9;
    var nameCol = (headers[colAsin] === '画像') ? (colAsin + 2) : (colAsin + 1);
    var expectedName = (productNameRow[nameCol] != null) ? String(productNameRow[nameCol]).trim() : '';
    var expectedJAN = (productNameRow[nameCol + 1] != null) ? String(productNameRow[nameCol + 1]).trim() : '';
    var expectedMaker = (productNameRow[nameCol + 2] != null) ? String(productNameRow[nameCol + 2]).trim() : '';
    var asinRows = [];
    for (var r = dataStartRow; r < data.length; r++) {
      var asin = String(data[r][colAsin] || '').trim().replace(/[\s\u3000]/g, '');
      if (asin.length >= 10) asinRows.push({ asin: asin, rowIndex: r + 1 });
    }
    Logger.log('[Keepa] ブロック' + (b + 1) + ': ASIN列=' + (colAsin + 1) + ', データ行数=' + asinRows.length + ', 先頭3件=' + asinRows.slice(0, 3).map(function(x) { return x.asin; }).join(','));
    if (asinRows.length === 0) continue;
    var take = (limit > 0) ? Math.min(limit, asinRows.length) : asinRows.length;
    var asins = asinRows.slice(0, take).map(function(x) { return x.asin; });
    var cachedMap = cacheEnabled ? getKeepaCachedResults(ss, asins, cacheDays) : {};
    var asinsToFetch = asins.filter(function(a) { return !cachedMap[a]; });
    var fetchResult = null;
    if (asinsToFetch.length > 0) {
      SpreadsheetApp.getActive().toast('Keepa に問い合わせ中（ブロック' + (b + 1) + '/' + asinColsFound.length + '、' + asinsToFetch.length + ' 件）...', '処理中', 2);
      fetchResult = fetchKeepaProductsForSheet(key, asinsToFetch);
      if (fetchResult == null) {
        SpreadsheetApp.getUi().alert('Keepa API の取得に失敗しました。キー・ネットワーク・トークン残数を確認してください。');
        return;
      }
      writeKeepaCache(ss, fetchResult.list, fetchResult.raw);
    } else {
      SpreadsheetApp.getActive().toast('ブロック' + (b + 1) + ': 全件キャッシュのためAPIスキップ（' + asins.length + ' 件）', 'Keepa取得', 2);
    }
    var asinToResult = {};
    var idx;
    for (idx = 0; idx < asins.length; idx++) {
      if (cachedMap[asins[idx]]) asinToResult[asins[idx]] = cachedMap[asins[idx]];
    }
    if (fetchResult && fetchResult.list) {
      fetchResult.list.forEach(function(x) { asinToResult[x.asin] = x; });
    }
    var results = asins.map(function(a) { return asinToResult[a]; }).filter(Boolean);
    Logger.log('[Keepa] API応答: results=' + (results ? results.length : 0) + '件');
    if (getGeminiApiKey() && results && results.length > 0) {
      var doClean = PropertiesService.getScriptProperties().getProperty('KEEPA_CLEAN_TITLE_BY_AI');
      if (doClean !== 'false') {
        var titlesToClean = results.map(function(r) { return r.title || ''; });
        var cleaned = cleanProductTitlesByGemini(titlesToClean);
        if (cleaned && cleaned.length === results.length) {
          lastTitleCleanedRan = true;
          for (var ci = 0; ci < results.length; ci++) { if (cleaned[ci] !== undefined && cleaned[ci] !== null) results[ci].title = String(cleaned[ci]).trim() || results[ci].title; }
        }
      }
    }
    if (fetchResult && fetchResult.raw && fetchResult.raw.length > 0) {
      writeKeepaRawToDebugSheet(ss, asinsToFetch, fetchResult.raw);
    }
    writeKeepaFieldListSheet(ss);
    var refImageUrl = skipRefImage ? '' : getRefImageUrlForAsinPasteBlock(ss, b);
    var costTaxInBlock = getCostTaxInForAsinPasteBlock(ss, b);
    Logger.log('[runKeepaFetchAsinPasteSheet] block=' + b + ' refImageUrl=' + (refImageUrl || '(空)') + ' skipRefImage=' + skipRefImage + ' costTaxInBlock=' + (costTaxInBlock != null ? costTaxInBlock : 'null'));
    var refSetCount = null;
    if (refImageUrl && getGeminiApiKey()) {
      var refSet = inferSetCountFromReferenceImageByGemini(refImageUrl);
      if (refSet && refSet.setCount != null) refSetCount = refSet.setCount;
      Utilities.sleep(350);
    }
    Logger.log('[Keepa] ブロック' + (b + 1) + ' refSetCount=' + (refSetCount != null ? refSetCount : 'null'));
    var testModeRefSet = (PropertiesService.getScriptProperties().getProperty('KEEPA_REF_IMAGE_SET_COUNT_TEST_MODE') === 'true');
    var asinToResult = {};
    results.forEach(function(x) { asinToResult[x.asin] = x; });
    var blockWritten = 0;
    var blockData = [];
    for (var i = 0; i < take && i < asinRows.length; i++) {
      var rowIndex = asinRows[i].rowIndex;
      var res = asinToResult[asinRows[i].asin];
      if (!res) {
        Logger.log('[Keepa] 行' + rowIndex + ' ASIN=' + asinRows[i].asin + ' はAPIに含まれずスキップ');
        continue;
      }
      var setCountName = (res.setCountFromTitle != null) ? res.setCountFromTitle : res.setCount;
      var setCountNameCandidates = extractSetCountCandidatesFromTitle(res.title || '');
      if (setCountNameCandidates.length === 0 && setCountName != null) setCountNameCandidates = [setCountName];
      if (colImageB >= 0 && res.imageUrl) {
        sheet.getRange(rowIndex, colImageB + 1).setFormula('=IMAGE("' + res.imageUrl.replace(/"/g, '""') + '", 2)');
      }
      var titleToWrite = (res.title != null && res.title !== '') ? res.title : '';
      if (colTitleB >= 0) sheet.getRange(rowIndex, colTitleB + 1).setValue(titleToWrite);
      var evalResult = evalMatchScore(expectedName, expectedJAN, expectedMaker, res.title, res.ean || null);
      var nameScoreBeforeImage = evalResult.score;
      var imageOverall = 0;
      var imageShape = null, imageColor = null, imageText = null, imagePackage = null, imageCapacity = null;
      if (refImageUrl && res.imageUrl && getGeminiApiKey() && !skipRefImage) {
        Logger.log('[Keepa] 画像マッチ開始 ASIN=' + asinRows[i].asin + ' refURL=' + refImageUrl.substring(0, 60) + ' keepaURL=' + res.imageUrl.substring(0, 60));
        var imgScore = getImageMatchScoreByGemini(refImageUrl, res.imageUrl);
        Logger.log('[Keepa] 画像マッチ結果 ASIN=' + asinRows[i].asin + ' imgScore=' + JSON.stringify(imgScore));
        if (imgScore && imgScore.overall != null) {
          imageOverall = imgScore.overall;
          imageShape = imgScore.shape; imageColor = imgScore.color; imageText = imgScore.text;
          imagePackage = imgScore.package; imageCapacity = imgScore.capacity;
          if (imageShape >= 55 && imageColor >= 55 && imageOverall < 50) {
            imageOverall = Math.max(imageOverall, 50);
          }
          var th = (typeof IMAGE_MATCH_THRESHOLD_REF_SET !== 'undefined' ? IMAGE_MATCH_THRESHOLD_REF_SET : 80);
          if (imageOverall >= th && refSetCount != null) {
            if (!testModeRefSet) {
              res.setCount = refSetCount;
              res.setCountReason = '参考画像';
            } else {
              res.setCountReason = '参考画像';
            }
          }
        }
        Utilities.sleep(350);
      }
      // 評価は画像のみ・重み付き（名前は使わない）。◎でないときも imageOverall を％表示する（0%にしない）
      var effectiveImage = (imageOverall > 70) ? imageOverall : 0;
      var newScore = Math.min(100, effectiveImage);
      // ◎＝画像スコア70以上、または shape が100点なら無条件
      var imgMinForCircle = 70;
      var showCircle = (imageOverall >= imgMinForCircle) || (imageShape === 100);
      var displayPct = Math.min(100, Math.round(imageOverall || 0));
      evalResult = { display: showCircle ? '◎' : displayPct + '%', score: newScore };
      var setCountImage = (imageOverall >= 80 && refSetCount != null) ? refSetCount : (res.setCountReason === '画像AI' ? res.setCount : null);
      if (colEvalB >= 0) sheet.getRange(rowIndex, colEvalB + 1).setValue(evalResult.display);
      if (colPriceB >= 0) sheet.getRange(rowIndex, colPriceB + 1).setValue(res.price != null ? res.price : '');
      blockData.push({
        rowIndex: rowIndex,
        asin: asinRows[i].asin,
        price: res.price,
        setCountName: setCountName,
        setCountNameCandidates: setCountNameCandidates,
        setCountImage: setCountImage,
        setCountCurrent: res.setCount,
        setCountReasonCurrent: res.setCountReason,
        imageOverall: imageOverall
      });
      if (colRefImageUrlB >= 0) sheet.getRange(rowIndex, colRefImageUrlB + 1).setValue(refImageUrl || '');
      if (colCostTaxInB >= 0) sheet.getRange(rowIndex, colCostTaxInB + 1).setValue(costTaxInBlock != null ? costTaxInBlock : '');
      if (colSetB >= 0) {
        var th = (typeof IMAGE_MATCH_THRESHOLD_REF_SET !== 'undefined' ? IMAGE_MATCH_THRESHOLD_REF_SET : 80);
        if (testModeRefSet && refSetCount != null && imageOverall >= th) {
          var colRefSet = colSetReasonB + 1;
          if (headerRowIndex >= 0 && (headers[colRefSet] == null || String(headers[colRefSet]).trim() === '')) {
            sheet.getRange(headerRowIndex + 1, colRefSet + 1).setValue('参考画像セット数');
          }
          sheet.getRange(rowIndex, colRefSet + 1).setValue(refSetCount);
        }
      }
      if (colUrlB >= 0) sheet.getRange(rowIndex, colUrlB + 1).setValue('https://www.amazon.co.jp/dp/' + asinRows[i].asin);
      Logger.log('[Keepa] 行' + rowIndex + ' ASIN=' + asinRows[i].asin + ' 名前=' + nameScoreBeforeImage + ' 画像overall=' + imageOverall + ' shape=' + imageShape + ' color=' + imageColor + ' 最終=' + evalResult.display);
      appendKeepaFetchLogRow(ss, {
        date: runStartTime,
        block: b + 1,
        rowIndex: rowIndex,
        asin: asinRows[i].asin,
        imageUrl: res.imageUrl || '',
        title: res.title || '',
        evalDisplay: evalResult.display || '',
        expectedName: expectedName || '',
        score: evalResult.score != null ? evalResult.score : 0,
        nameScore: nameScoreBeforeImage,
        imageOverall: imageOverall || '',
        imageShape: imageShape,
        imageColor: imageColor,
        imageText: imageText,
        imagePackage: imagePackage,
        imageCapacity: imageCapacity
      });
      blockWritten++;
      totalWritten++;
    }
    for (var di = 0; di < blockData.length; di++) {
      var d = blockData[di];
      Logger.log('[Keepa] セット数照合 行' + d.rowIndex + ' ASIN=' + (d.asin || '') + ' price=' + d.price + ' 候補(名)=' + (d.setCountNameCandidates ? d.setCountNameCandidates.join(',') : '') + ' 画像=' + (d.setCountImage != null ? d.setCountImage : '') + ' imageOverall=' + (d.imageOverall != null ? d.imageOverall : ''));
    }
    var resolved = resolveSetCountByUnitPrice(blockData, costTaxInBlock);
    var reasonCol1Based = colSetReasonB + 1;
    for (var idx = 0; idx < resolved.length; idx++) {
      var r = resolved[idx];
      var dataRow = blockData[idx];
      if (!dataRow) continue;
      if (colSetB >= 0) {
        sheet.getRange(dataRow.rowIndex, colSetB + 1).setValue(r.setCount != null ? r.setCount : '');
      }
      if (r.setCount != null || (r.setCountReason && String(r.setCountReason).trim())) {
        sheet.getRange(dataRow.rowIndex, reasonCol1Based).setValue(r.setCountReason || '正規表現');
        totalSetReasonWritten++;
      }
      // 卸値範囲外の行はセット数_AI根拠のセルを赤背景で目立たせる（評価列は上書きしない）
      if (reasonCol1Based > 0) {
        var reasonCell = sheet.getRange(dataRow.rowIndex, reasonCol1Based);
        if (r.setCountReason && String(r.setCountReason).trim() === '卸値範囲外') {
          reasonCell.setBackground('#ff0000');
        } else {
          reasonCell.setBackground(null);
        }
      }
    }
    if (blockWritten > 0 && colEvalB >= 0) {
      SpreadsheetApp.flush();
      // ④1塊＝AI情報取得datasheetの1行に対応する列塊（画像〜卸値の10列）。塊ごとに評価で降順ソート。
      var blockCols = (typeof ASIN_PASTE_BLOCK_HEADERS !== 'undefined' ? ASIN_PASTE_BLOCK_HEADERS.length : 10);
      if (blockCols < 1) blockCols = 10;
      var sortStartCol1Based = blockStartCol + 1;
      var endCol1Based = blockStartCol + blockCols;
      var colEvalInBlock = 3;
      var sortFirstRow = dataStartRow + 1;
      var sortNumCols = endCol1Based - sortStartCol1Based + 1;
      Logger.log('[Keepa] ソート範囲 firstRow1Based=' + sortFirstRow + ' numRows=' + blockWritten + ' startCol1Based=' + sortStartCol1Based + ' endCol1Based=' + endCol1Based + ' numCols=' + sortNumCols);
      try {
        sortAsinPasteSheetBlockByEvalDesc(sheet, sortFirstRow, blockWritten, colEvalInBlock, sortStartCol1Based, endCol1Based);
      } catch (sortErr) {
        Logger.log('[Keepa] ソート失敗: ' + (sortErr && sortErr.message));
      }
    }
    Logger.log('[Keepa] ブロック' + (b + 1) + ' 書き込み=' + blockWritten + '件');
  }
  Logger.log('[Keepa] 合計書き込み=' + totalWritten + '件');
  SpreadsheetApp.flush();
  var evalTop5 = [];
  if (totalWritten > 0 && asinColsFound.length > 0) {
    var firstBlockEvalCol = asinColsFound[0] + 2;
    var firstDataRow = dataStartRow + 1;
    var lastRow = Math.min(firstDataRow + 4, firstDataRow + totalWritten - 1);
    if (lastRow >= firstDataRow) {
      var evalRange = sheet.getRange(firstDataRow, firstBlockEvalCol + 1, lastRow, firstBlockEvalCol + 1);
      var evals = evalRange.getValues();
      evalTop5 = evals.map(function(r) { return r[0] != null ? String(r[0]) : ''; });
    }
  }
  writeKeepaDiagnosticSheet(ss, {
    date: new Date(),
    totalWritten: totalWritten,
    evalTop5: evalTop5.join(', '),
    titleCleanedRan: lastTitleCleanedRan,
    setReasonCount: totalSetReasonWritten,
    note: 'セット数_AI根拠は1ブロック目ならH列。このシートをコピーしてAIに貼ると差分分析できます。'
  });
  SpreadsheetApp.getActive().toast(totalWritten + ' 件をシートに書き込みました。（評価の降順で並び替え済み）', 'Keepa取得完了', 5);
}

/** ブロック内での「セット数」列のオフセット（0-based）。ASIN_PASTE_BLOCK_HEADERS の並びでセット数は6列目＝インデックス5。 */
var ASIN_PASTE_SET_COL_OFFSET = 5;

/**
 * ASIN貼り付けシートの各ブロックの「セット数」列から数値を集約し、競合セット数一覧と穴のセット数候補を返す。
 * 対象は「最初にセット数が1件以上入っているブロック」のみ（1商品分として扱う）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 「ASIN貼り付け（Keepa用）」シート
 * @return {{ competitive: number[], hole: number[] } | null} 競合セット数（昇順・重複除く）と穴のセット数（最大10件）。データがなければ null
 */
function getAggregatedSetCountsFromAsinPasteSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(data.length, 5); hr++) {
    var rowData = data[hr] || [];
    for (var hc = 0; hc < rowData.length; hc++) {
      if (String(rowData[hc]).trim() === 'ASIN') { headerRowIdx = hr; break; }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) return null;
  var headers = data[headerRowIdx];
  var dataStartRow = headerRowIdx + 1;
  var asinColumns = [];
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'ASIN') asinColumns.push(c);
  }
  if (asinColumns.length === 0) return null;
  var setNumbers = [];
  for (var b = 0; b < asinColumns.length; b++) {
    var colAsin = asinColumns[b];
    var colSet = colAsin + 4;   // セット数 = ASIN+4 (画像,ASIN,商品名,評価,競合価格,セット数)
    var colEval = colAsin + 2;  // 評価 = ASIN+2
    if (colSet >= headers.length) continue;
    for (var r = dataStartRow; r < data.length; r++) {
      if (String(data[r][colEval] || '').trim() !== '◎') continue;
      var val = data[r][colSet];
      if (val === '' || val == null) continue;
      var n = parseInt(String(val).trim(), 10);
      if (!isNaN(n) && n >= 1) setNumbers.push(n);
    }
    if (setNumbers.length > 0) break;
  }
  if (setNumbers.length === 0) return null;
  var competitive = [];
  var seen = {};
  setNumbers.forEach(function(n) { if (!seen[n]) { seen[n] = true; competitive.push(n); } });
  competitive.sort(function(a, b) { return a - b; });
  var maxC = competitive.length ? competitive[competitive.length - 1] : 0;
  var cap = Math.max(maxC + 10, 20);
  var remaining = 10 - competitive.length;
  var hole = [];
  if (remaining > 0) {
    for (var i = 1; i <= cap && hole.length < remaining; i++) {
      if (seen[i]) continue;
      hole.push(i);
    }
  }
  return { competitive: competitive, hole: hole };
}

/**
 * ASIN貼り付けシートの全ブロックについて、各ブロックの「セット数」列から競合セット数・穴のセット数を集約する。
 * 優先4でマスタ◎行とブロックを1対1で対応させるときに使用する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 「ASIN貼り付け（Keepa用）」シート
 * @return {Array<{ competitive: number[], hole: number[] } | null>} ブロック順の配列。データがないブロックは null
 */

/**
 * アンカー数列間の最大ギャップを中点分割で比例配分し、numSlots 個の補間数を返す。
 * @param {number[]} anchors ソート済みのアンカー数列（competitive + [bizSet] など）
 * @param {number} numSlots 追加したいスロット数
 * @return {number[]} 補間数のソート済み配列
 */
function fillGapsProportionally(anchors, numSlots) {
  if (numSlots <= 0 || anchors.length < 2) return [];
  var sorted = anchors.slice().sort(function(a, b) { return a - b; });
  var seen = {};
  sorted.forEach(function(n) { seen[n] = true; });
  var result = [];
  for (var iter = 0; result.length < numSlots && iter < 300; iter++) {
    // 最大ギャップを探す
    var bestSize = -1, bestLo = -1, bestHi = -1;
    for (var k = 0; k + 1 < sorted.length; k++) {
      var g = sorted[k + 1] - sorted[k];
      if (g > bestSize) { bestSize = g; bestLo = sorted[k]; bestHi = sorted[k + 1]; }
    }
    if (bestSize <= 1) break; // 埋めるギャップなし
    var mid = Math.floor((bestLo + bestHi) / 2);
    if (seen[mid]) break;
    seen[mid] = true;
    result.push(mid);
    // ソート済み配列に mid を挿入
    var ins = 0;
    while (ins < sorted.length && sorted[ins] < mid) ins++;
    sorted.splice(ins, 0, mid);
  }
  result.sort(function(a, b) { return a - b; });
  return result;
}

function getAggregatedSetCountsByBlocksFromAsinPasteSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(data.length, 5); hr++) {
    var rowData = data[hr] || [];
    for (var hc = 0; hc < rowData.length; hc++) {
      if (String(rowData[hc]).trim() === 'ASIN') { headerRowIdx = hr; break; }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) return [];
  var headers = data[headerRowIdx];
  var dataStartRow = headerRowIdx + 1;
  var asinColumns = [];
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'ASIN') asinColumns.push(c);
  }
  if (asinColumns.length === 0) return [];
  var results = [];
  for (var b = 0; b < asinColumns.length; b++) {
    var colAsin = asinColumns[b];
    var colSet = colAsin + 4;   // セット数 = ASIN+4
    var colEval = colAsin + 2;  // 評価 = ASIN+2
    if (colSet >= headers.length) {
      results.push(null);
      continue;
    }
    var compSetNumbers = [];
    var seenComp = {};
    for (var r = dataStartRow; r < data.length; r++) {
      // ◎フィルターなし: 全行のセット数を収集（価格は getMinPriceAndAsin~で◎のみ）
      var val = data[r][colSet];
      if (val === '' || val == null) continue;
      var n = parseInt(String(val).trim(), 10);
      if (!isNaN(n) && n >= 1 && !seenComp[n]) {
        seenComp[n] = true;
        compSetNumbers.push(n);
      }
    }
    if (compSetNumbers.length === 0) {
      results.push(null);
      continue;
    }
    compSetNumbers.sort(function(a, b) { return a - b; });
    var maxC = compSetNumbers[compSetNumbers.length - 1];
    // 競合セット数は最大9件に絞る（業者用1枠確保）。小さい順で重要度が高いものを優先
    var compCapped = compSetNumbers.slice(0, 9);
    // hole はmenuSetCompositionProposal内でbizSetを考慮して比例配分で計算する
    results.push({ competitive: compCapped, hole: [], maxCompetitive: maxC });
  }
  return results;
}

/**
 * 商品名から「NP」パターンを読み取り、1袋あたりの包数を返す。
 * 例: "蒸し生姜湯 4P" → 4, "六漢生姜湯 48P（16g×48P）" → 48
 * 見つからない場合は 1 を返す。
 * @param {string} productName 商品名
 * @return {number} 1袋あたりの包数
 */
function extractPacketsPerBag(productName) {
  // 末尾に近い "数字P" を優先して取得（例: 48P が 4P より正確）
  var matches = productName.match(/(\d+)\s*[Pp]/g);
  if (!matches || matches.length === 0) return 1;
  // 複数ある場合（例: "48P（16g×48P）"）は最大値を採用
  var maxVal = 1;
  for (var k = 0; k < matches.length; k++) {
    var n = parseInt(matches[k], 10);
    if (!isNaN(n) && n > maxVal) maxVal = n;
  }
  return maxVal;
}

/**
 * 商品名と1袋包数から、Geminiで1日使用包数を推測し業者用セット数を返す。
 * 計算式: ceil(daily_packets × 60 / packetsPerBag)
 * Gemini APIキーなし or 失敗時は fallback = maxCompetitive × 2。
 * @param {string} productName 商品名
 * @param {number} maxCompetitive 競合の最大セット数
 * @return {number} 業者用セット数
 */
function getBusinessSetCount(productName, maxCompetitive) {
  var fallback = maxCompetitive * 2;
  var packetsPerBag = extractPacketsPerBag(productName);
  Logger.log('[業者用セット数] 商品名=' + productName + ' packetsPerBag=' + packetsPerBag);

  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    Logger.log('[業者用セット数] APIキーなし。fallback=' + fallback);
    return fallback;
  }
  try {
    // Geminiには「1日何包使うか」だけを聞く。包数→袋数換算はコード側で行う
    var prompt =
      '商品「' + productName + '」について、一般的な成人が毎日使用した場合に' +
      '1日あたり何包（1包＝1回分）使うか、整数で推測してください。\n' +
      '回答はJSONのみ: {"daily_packets": N}  ※Nは整数。説明不要。';
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL_GEMINI +
      ':generateContent?key=' + encodeURIComponent(apiKey);
    var body = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 512 }
    };
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText());
    var rawText = ((json.candidates || [])[0] || {});
    rawText = ((rawText.content || {}).parts || [{}])[0].text || '';
    rawText = rawText.replace(/```json|```/g, '').trim();
    Logger.log('[業者用セット数] Gemini rawText=' + rawText.substring(0, 300));
    // daily_packets を正規表現で直接抽出
    var numMatch = rawText.match(/"daily_packets"\s*:\s*(\d+)/) ||
                   rawText.match(/daily_packets[^0-9]*(\d+)/);
    var dailyPackets = numMatch ? parseInt(numMatch[1], 10) : NaN;
    if (isNaN(dailyPackets) || dailyPackets < 1) {
      Logger.log('[業者用セット数] daily_packets抽出失敗。response=' + rawText.substring(0, 200) + ' fallback=' + fallback);
      return fallback;
    }
    // 12ヶ月（360日）消費袋数 = ceil(daily_packets × 360 / packetsPerBag)
    // fallback との max は取らない（非◎行でmaxCompetitiveが膨らむと fallback が過大になるため）
    // bizSet < currentMax の引き上げは menuSetCompositionProposal 側の sixMonthCap 内で行う
    var businessSets = Math.ceil(dailyPackets * 360 / packetsPerBag);
    Logger.log('[業者用セット数] ' + productName +
      ': daily=' + dailyPackets + '包/日' +
      ' packetsPerBag=' + packetsPerBag +
      ' 12ヶ月=' + businessSets + '袋');
    return businessSets;
  } catch (e) {
    Logger.log('[業者用セット数] Gemini呼び出し失敗: ' + e.message + ' fallback=' + fallback);
    return fallback;
  }
}

/**
 * ASIN貼り付けシートの指定ブロックから、◎行のみを対象にセット数別の
 * 最低価格・そのASIN・商品URLを返す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 「ASIN貼り付け（Keepa用）」シート
 * @param {number} blockIndex 0-based ブロック番号
 * @return {Object<number, {price: number, asin: string, url: string}>} セット数 → {price, asin, url}
 */
function getMinPriceAndAsinBySetCountFromAsinPasteBlock(sheet, blockIndex) {
  var data = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(data.length, 5); hr++) {
    var rowData = data[hr] || [];
    for (var hc = 0; hc < rowData.length; hc++) {
      if (String(rowData[hc]).trim() === 'ASIN') { headerRowIdx = hr; break; }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) return {};
  var headers = data[headerRowIdx];
  var dataStartRow = headerRowIdx + 1;
  var asinColumns = [];
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'ASIN') asinColumns.push(c);
  }
  if (blockIndex < 0 || blockIndex >= asinColumns.length) return {};
  var colAsin  = asinColumns[blockIndex];
  var colEval  = colAsin + 2; // 評価（◎フィルター用）
  var colPrice = colAsin + 3; // 競合価格(Amazon)
  var colSet   = colAsin + 4; // セット数
  var colUrl   = colAsin + 5; // 商品URL（ASIN+5）
  if (colSet >= headers.length || colPrice >= headers.length) return {};
  var bySet = {};
  for (var r = dataStartRow; r < data.length; r++) {
    // ◎行のみ対象
    if (String(data[r][colEval] || '').trim() !== '◎') continue;
    var setVal = data[r][colSet];
    var priceVal = data[r][colPrice];
    if (setVal === '' || setVal == null) continue;
    var setNum = parseInt(String(setVal).trim(), 10);
    if (isNaN(setNum) || setNum < 1) continue;
    var price = Number(String(priceVal).replace(/,/g, ''));
    if (isNaN(price) || price < 10 || price > 100000) continue;
    var asin = String(data[r][colAsin] || '').trim();
    var url  = colUrl < data[r].length ? String(data[r][colUrl] || '').trim() : '';
    if (bySet[setNum] === undefined || price < bySet[setNum].price) {
      bySet[setNum] = { price: Math.round(price), asin: asin, url: url };
    }
  }
  Logger.log('[getMinPriceAndAsin] block=' + blockIndex + ' colAsin=' + colAsin +
    ' colPrice=' + colPrice + ' colSet=' + colSet + ' colUrl=' + colUrl +
    ' result=' + JSON.stringify(bySet));
  return bySet;
}

/**
 * 00_設定マスタ から「販売手数料」のキー→料率マップを取得する。
 * シート: A=設定項目, B=キー, D=値1(料率)。設定項目=販売手数料 の行のみ対象。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Object<string, number>} キー → 料率(0.08等)。シートが無い・行が無い場合は空オブジェクト。
 */
function getCommissionRateMapFromSettingsMaster(ss) {
  var sheet = ss.getSheetByName(SETTINGS_MASTER_SHEET_NAME);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var item = String(row[0] || '').trim();
    if (item !== SETTINGS_ITEM_SALES_FEE) continue;
    var key = String(row[1] || '').trim();
    if (key === '') continue;
    var val1 = row[3];
    if (val1 === null || val1 === undefined || val1 === '') continue;
    var rate = Number(String(val1).replace(/,/g, ''));
    if (!isNaN(rate)) map[key] = rate;
  }
  return map;
}

/**
 * Amazon用の販売手数料率を、カテゴリと価格帯から取得する。
 * キーは「食品:1500円以下」「食品:1501円以上」形式。カテゴリのみの場合は価格帯を付与して検索。
 * @param {Object<string, number>} rateMap getCommissionRateMapFromSettingsMaster の戻り値
 * @param {string} categoryValue マスタの「amazon カテゴリー」の値（例: 食品, 食品:1501円以上）
 * @param {number} price 販売価格（税込）。1500円以下/1501円以上でキーが変わる場合に使用
 * @return {number|null} 料率。未定義なら null
 */
function getAmazonCommissionRateForPrice(rateMap, categoryValue, price) {
  if (!categoryValue || String(categoryValue).trim() === '') return null;
  var cat = String(categoryValue).trim();
  if (rateMap[cat] !== undefined) return rateMap[cat];
  var base = cat.indexOf(':') >= 0 ? cat.replace(/:.*$/, '').trim() : cat;
  if (!base) return null;
  var band = (price !== undefined && !isNaN(price) && price <= 1500) ? '1500円以下' : '1501円以上';
  var key = base + ':' + band;
  return rateMap[key] !== undefined ? rateMap[key] : null;
}

/**
 * 商品名から Amazon カテゴリーを Gemini で推測する。
 * 00_設定マスタの有効なベースカテゴリ一覧（価格帯を除いたもの）を候補として渡し、
 * Gemini が最も適切なものを1つ選んで返す。
 * @param {string} productName 商品名（例: 蒸し生姜湯 4P）
 * @param {string[]} baseCategoryList 有効なベースカテゴリ一覧（例: ['食品', 'ドラッグストア', ...]）
 * @return {string} 推測されたベースカテゴリ（例: '食品'）。失敗時は ''
 */
function guessAmazonCategoryByGemini(productName, baseCategoryList) {
  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    Logger.log('[カテゴリ推測] APIキーなし。productName=' + productName);
    return '';
  }
  if (!baseCategoryList || baseCategoryList.length === 0) return '';
  try {
    var prompt =
      '以下の商品名から、Amazon販売手数料カテゴリとして最も適切なものを1つ選んでください。\n' +
      '商品名: 「' + productName + '」\n' +
      '候補カテゴリ: ' + baseCategoryList.join(' / ') + '\n' +
      '回答はJSONのみ: {"category": "カテゴリ名"}  ※候補の中から1つそのままの文字列で。説明不要。';
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL_GEMINI +
      ':generateContent?key=' + encodeURIComponent(apiKey);
    var body = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 512 }
    };
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText());
    var rawText = (((json.candidates || [])[0] || {}).content || {}).parts;
    rawText = rawText ? (rawText[0].text || '') : '';
    rawText = rawText.replace(/```json|```/g, '').trim();
    Logger.log('[カテゴリ推測] rawText=' + rawText.substring(0, 200));
    var match = rawText.match(/"category"\s*:\s*"([^"]+)"/) ||
                rawText.match(/"category"\s*:\s*"([^"]*)/);
    if (!match || !match[1]) return '';
    var guessed = match[1].trim();
    if (baseCategoryList.indexOf(guessed) >= 0) return guessed;
    var lowerGuessed = guessed.toLowerCase();
    for (var i = 0; i < baseCategoryList.length; i++) {
      if (baseCategoryList[i].toLowerCase() === lowerGuessed) return baseCategoryList[i];
    }
    Logger.log('[カテゴリ推測] 候補に一致なし: guessed=' + guessed);
    return '';
  } catch (e) {
    Logger.log('[カテゴリ推測] Gemini呼び出し失敗: ' + e.message);
    return '';
  }
}

/**
 * ベースカテゴリ（例: "食品"）と価格から 00_設定マスタのフルキー（例: "食品:1500円以下"）を返す。
 * 価格帯のないカテゴリ（例: "服&ファッション小物"）はそのまま返す。
 * @param {Object} rateMap getCommissionRateMapFromSettingsMaster の返り値
 * @param {string} baseCategory ベースカテゴリ名
 * @param {number|null} price 競合価格（null/NaN の場合はデフォルト 1501円以上 扱い）
 * @return {string} フルキー
 */
function resolveFullCategoryKey(rateMap, baseCategory, price) {
  if (!baseCategory) return '';
  if (rateMap[baseCategory] !== undefined) return baseCategory;
  var hasBand = (rateMap[baseCategory + ':1500円以下'] !== undefined ||
                 rateMap[baseCategory + ':1501円以上'] !== undefined);
  if (!hasBand) return baseCategory;
  var p = (price !== null && price !== undefined && !isNaN(price) && price > 0) ? price : 9999;
  return p <= 1500 ? baseCategory + ':1500円以下' : baseCategory + ':1501円以上';
}

/**
 * マスタの「amazon カテゴリー」が空の子行を対象に、Gemini で自動推測して書き込む。
 * 同一JANは1回の推測結果を全行に適用する。
 * 価格帯があるカテゴリ（例: 食品）は競合価格でフルキーを確定する。
 * 競合価格がない行は ①同JAN他行の平均 → ②上下±5行 → ③デフォルト(1501円以上) の順で補完する。
 * メニュー「③ 販売価格の調整 → Amazonカテゴリーを自動入力（Gemini）」から実行。
 */
function menuFillAmazonCategoryByGemini() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('「' + MASTER_SHEET_NAME + '」シートが見つかりません。');
    return;
  }
  if (!getGeminiApiKey()) {
    SpreadsheetApp.getUi().alert('GEMINI_API_KEY が Script Properties に設定されていません。');
    return;
  }
  var rateMap = getCommissionRateMapFromSettingsMaster(ss);
  var baseCategoryList = getBaseCategoryList(rateMap);
  if (baseCategoryList.length === 0) {
    SpreadsheetApp.getUi().alert('00_設定マスタ から有効なカテゴリキーが取得できませんでした。');
    return;
  }

  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(masterValues.length, 25); hr++) {
    if ((masterValues[hr] || []).indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = hr; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colSetQty = masterColMap[COL_MASTER_TOTAL_QTY];
  var colCategory = masterColMap[COL_AMAZON_CATEGORY];
  var colName = masterColMap['商品名ベース'] !== undefined ? masterColMap['商品名ベース'] : masterColMap['商品名'];
  var colCheckboxFill = masterColMap[CHECKBOX_HEADER_NAME]; // 出品CKフィルタ用

  if (colCategory === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「' + COL_AMAZON_CATEGORY + '」列がありません。列を追加してから実行してください。');
    return;
  }
  if (colJan === undefined || colSetQty === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「JANコード」または「' + COL_MASTER_TOTAL_QTY + '」列がありません。');
    return;
  }

  var colCompPrice = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];

  // 全行の競合価格を先にスキャン（価格帯判定の補完用）
  var rowToCompPrice = {};   // rowIdx(0-based) → price
  var janToCompPrices = {};  // JAN → [price, ...]
  for (var rp = headerRowIdx + 1; rp < masterValues.length; rp++) {
    if (colCompPrice === undefined) break;
    var cp = parseFloat(masterValues[rp][colCompPrice]);
    if (!isNaN(cp) && cp > 0) {
      rowToCompPrice[rp] = cp;
      var janRp = String(masterValues[rp][colJan] || '').trim();
      if (janRp) {
        if (!janToCompPrices[janRp]) janToCompPrices[janRp] = [];
        janToCompPrices[janRp].push(cp);
      }
    }
  }

  // 空カテゴリの子行を収集し、JAN→商品名のマップを作る（出品CK=TRUEの行のみ対象）
  var janToProductName = {};
  var targetRows = [];
  for (var r = headerRowIdx + 1; r < masterValues.length; r++) {
    var mRow = masterValues[r];
    // 出品CK列がある場合はTRUEの行のみを対象とする
    if (colCheckboxFill !== undefined && mRow[colCheckboxFill] !== true) continue;
    var jan = String(mRow[colJan] || '').trim();
    var setVal = String(mRow[colSetQty] || '').trim();
    if (jan === '' || setVal === '') continue;
    var catVal = String(mRow[colCategory] || '').trim();
    if (catVal !== '') continue;
    targetRows.push(r);
    if (!janToProductName[jan] && colName !== undefined) {
      var pName = String(mRow[colName] || '').trim();
      if (pName) janToProductName[jan] = pName;
    }
  }

  if (targetRows.length === 0) {
    SpreadsheetApp.getActive().toast('「amazon カテゴリー」が空の子行はありませんでした。', '完了', 5);
    return;
  }

  // JAN単位でGemini呼び出し（重複回避）
  var janToGuessed = {};
  var uniqueJans = Object.keys(janToProductName);
  for (var j = 0; j < uniqueJans.length; j++) {
    var jan2 = uniqueJans[j];
    var pName2 = janToProductName[jan2];
    Utilities.sleep(200);
    var guessed = guessAmazonCategoryByGemini(pName2, baseCategoryList);
    janToGuessed[jan2] = guessed;
    Logger.log('[カテゴリ自動入力] JAN=' + jan2 + ' 商品名=' + pName2 + ' ベース推測=' + guessed);
  }

  // 書き込み（価格帯をフルキーに解決して書き込む）
  var filled = 0;
  var failed = 0;
  for (var t = 0; t < targetRows.length; t++) {
    var rIdx = targetRows[t];
    var jan3 = String(masterValues[rIdx][colJan] || '').trim();
    var guessedBase = janToGuessed[jan3] || '';
    if (!guessedBase) {
      masterSheet.getRange(rIdx + 1, colCategory + 1).setBackground('#ffcccc');
      failed++;
      continue;
    }

    // 価格帯判定：競合価格を取得（①自行 → ②同JAN平均 → ③上下±5行 → ④デフォルト）
    var effectivePrice = null;
    var priceSource = '';
    // ① 自行の競合価格
    if (rowToCompPrice[rIdx] !== undefined) {
      effectivePrice = rowToCompPrice[rIdx];
      priceSource = '自行';
    }
    // ② 同JAN他行の平均
    if (effectivePrice === null && jan3 && janToCompPrices[jan3] && janToCompPrices[jan3].length > 0) {
      var prices = janToCompPrices[jan3];
      var sum = 0;
      for (var pi = 0; pi < prices.length; pi++) sum += prices[pi];
      effectivePrice = sum / prices.length;
      priceSource = '同JAN平均';
    }
    // ③ 上下±5行の競合価格（上から優先）
    if (effectivePrice === null) {
      for (var delta = 1; delta <= 5 && effectivePrice === null; delta++) {
        if (rowToCompPrice[rIdx - delta] !== undefined) {
          effectivePrice = rowToCompPrice[rIdx - delta];
          priceSource = '上' + delta + '行';
        } else if (rowToCompPrice[rIdx + delta] !== undefined) {
          effectivePrice = rowToCompPrice[rIdx + delta];
          priceSource = '下' + delta + '行';
        }
      }
    }
    // ④ デフォルト（1501円以上として扱う）
    if (effectivePrice === null) priceSource = 'デフォルト(1501円以上)';

    var fullKey = resolveFullCategoryKey(rateMap, guessedBase, effectivePrice);
    masterSheet.getRange(rIdx + 1, colCategory + 1).setValue(fullKey).setBackground(null);
    Logger.log('[カテゴリ自動入力] 行' + (rIdx + 1) + ' JAN=' + jan3 +
      ' base=' + guessedBase + ' price=' + effectivePrice + '(' + priceSource + ')' +
      ' → ' + fullKey);
    filled++;
  }

  SpreadsheetApp.getActive().toast(
    '自動入力完了。入力済み: ' + filled + '行' +
    (failed > 0 ? '  推測失敗（赤表示）: ' + failed + '行' : ''),
    'Amazonカテゴリー自動入力 完了',
    8
  );
}

/**
 * 出品CK（CHECKBOX_HEADER_NAME）を選択行でクリア（FALSE に）する。
 * 自動出品完了後に手動で実行して出品対象から外す用途。
 * メニュー「③ 販売価格の調整 → 出品CKをクリア（選択行）」から実行。
 */
function menuClearShippingCheckForSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('「' + MASTER_SHEET_NAME + '」シートが見つかりません。');
    return;
  }
  var activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() !== MASTER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) を開き、クリアしたい行を選択してから実行してください。');
    return;
  }
  var range = activeSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してから実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(masterValues.length, 25); hr++) {
    if ((masterValues[hr] || []).indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = hr; break; }
  }
  var masterColMap = headerRowIdx >= 0 ? getColumnIndexMap(masterValues[headerRowIdx]) : {};
  var colCheckbox = masterColMap[CHECKBOX_HEADER_NAME];
  if (colCheckbox === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「' + CHECKBOX_HEADER_NAME + '」列がありません。');
    return;
  }
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var cleared = 0;
  for (var r = 0; r < numRows; r++) {
    var row1Based = startRow + r;
    if (headerRowIdx >= 0 && row1Based <= headerRowIdx + 1) continue;
    masterSheet.getRange(row1Based, colCheckbox + 1).setValue(false);
    cleared++;
  }
  SpreadsheetApp.getActive().toast(
    '選択行 ' + cleared + ' 行の出品CKをクリアしました。',
    '出品CKクリア 完了',
    5
  );
}

/**
 * 手数料レートマップからベースカテゴリ一覧（価格帯を除いた種類）を返す。
 * 例: '食品:1500円以下' → '食品'、'Amazonデバイス用アクセサリ' → そのまま。
 * Rakuten / Yahoo は除外する。
 * @param {Object<string, number>} rateMap
 * @return {string[]} ユニークなベースカテゴリ名の配列
 */
function getBaseCategoryList(rateMap) {
  var set = {};
  var keys = Object.keys(rateMap);
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (k === 'Rakuten' || k === 'Yahoo') continue;
    var base = k.indexOf(':') >= 0 ? k.replace(/:.*$/, '').trim() : k;
    if (base) set[base] = true;
  }
  return Object.keys(set);
}

/**
 * ② 価格設定ロジック（RESEARCH_AND_ESTIMATE §8.8.14 確定版）。
 *
 * 【ceiling の決め方（セット数ごと）】
 *   競合あり: min(競合価格-1, 直近下セット数の競合単価×自セット数-1)
 *   穴(競合なし): 直近下セット数の競合単価×自セット数-1 (なければ上側を使用)
 *
 * 【finalPrice の決め方】
 *   利益率 > 20%  → 20%になるよう切り下げ
 *   利益率 8-20%  → ceiling をそのまま使用
 *   利益率 < 8% かつ利益額 ≥ 200円 → ceiling + 黄色警告
 *   利益額 < 200円 → 最低利益確保価格（競合より高くてもOK）
 *   それでも不可  → スキップ + 赤警告
 *
 * 書き込み: 販売価格amazon および amazon手数料計。楽天・Yahoo!は別フェーズでAPI競合取得後に計算。
 */
function menuProposeSalesPrices() {
  var ui = SpreadsheetApp.getUi();
  if (!ui) return;
  var response = ui.alert('送料は設定しましたか？', '「はい」で実行、「いいえ」でキャンセルします。', ui.ButtonSet.YES_NO);
  if (response === ui.Button.NO) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  var asinSheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  var aiSheet = ss.getSheetByName(TARGET_SHEET_NAME);

  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('「' + MASTER_SHEET_NAME + '」シートが見つかりません。');
    return;
  }
  if (!asinSheet) {
    SpreadsheetApp.getUi().alert('「' + ASIN_PASTE_SHEET_NAME + '」シートが見つかりません。先にKeepa取得まで実行してください。');
    return;
  }
  if (!aiSheet) {
    SpreadsheetApp.getUi().alert('「' + TARGET_SHEET_NAME + '」シートが見つかりません。');
    return;
  }

  var settingsSheet = ss.getSheetByName(SETTINGS_MASTER_SHEET_NAME);
  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('「' + SETTINGS_MASTER_SHEET_NAME + '」シートが見つかりません。手数料の参照ができません。');
    return;
  }

  var rateMap = getCommissionRateMapFromSettingsMaster(ss);
  var baseCategoryList = getBaseCategoryList(rateMap);
  var useGeminiForCategory = (getGeminiApiKey() !== '' && baseCategoryList.length > 0);

  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(masterValues.length, 25); hr++) {
    if ((masterValues[hr] || []).indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = hr; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }

  var masterHeaders = masterValues[headerRowIdx];
  var masterColMap = getColumnIndexMap(masterHeaders);
  var colJan = masterColMap['JANコード'];
  var colSetQty = masterColMap[COL_MASTER_TOTAL_QTY];
  var colCost = masterColMap[COL_COST_SET_TAX_IN] !== undefined ? masterColMap[COL_COST_SET_TAX_IN] : masterColMap[COL_COST_TAX_IN];
  var colCategory = masterColMap[COL_AMAZON_CATEGORY];
  var colShipping = masterColMap[COL_SHIPPING_FIXED] !== undefined ? masterColMap[COL_SHIPPING_FIXED] : masterColMap[COL_SHIPPING];
  var colPriceAmazon = masterColMap[COL_PRICE_AMAZON];
  var colAmazonFee = masterColMap[COL_AMAZON_FEE];
  var colCompetitivePrice = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  var colCheckboxPrice = masterColMap[CHECKBOX_HEADER_NAME]; // 出品CKフィルタ用

  if (colJan === undefined || colSetQty === undefined || colCost === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「JANコード」「' + COL_MASTER_TOTAL_QTY + '」または卸値列（「' + COL_COST_SET_TAX_IN + '」「' + COL_COST_TAX_IN + '」のいずれか）がありません。');
    return;
  }
  if (colPriceAmazon === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「' + COL_PRICE_AMAZON + '」列がありません。');
    return;
  }
  if (colCompetitivePrice === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「' + COL_COMPETITIVE_PRICE_AMAZON + '」列がありません。');
    return;
  }

  var aiData = aiSheet.getDataRange().getValues();
  var aiColMap = getColumnIndexMap(aiData[0] || []);
  var aiJanIdx = aiColMap['JANコード'];
  var janToBlock = {};
  if (aiJanIdx !== undefined) {
    var blockNo = 0;
    for (var ar = 1; ar < aiData.length; ar++) {
      var jan = String(aiData[ar][aiJanIdx] || '').trim();
      if (jan !== '') {
        if (janToBlock[jan] === undefined) janToBlock[jan] = blockNo;
        blockNo++;
      }
    }
  }

  var blockCache = {};
  function getCachedBlock(blockIndex) {
    if (blockCache[blockIndex] === undefined) {
      blockCache[blockIndex] = getMinPriceAndAsinBySetCountFromAsinPasteBlock(asinSheet, blockIndex);
    }
    return blockCache[blockIndex];
  }

  var MIN_PROFIT   = 200;
  var RATE_MAX     = 0.30;  // 利益率ハード上限（超えたら切り下げ）
  var RATE_WARN    = 0.08;  // 利益率警告ライン（下回ったら黄色）

  var parseNum = function(v) {
    if (v === undefined || v === null || v === '') return NaN;
    var n = Number(String(v).replace(/,/g, ''));
    return isNaN(n) ? NaN : n;
  };

  /**
   * ceiling・コスト・手数料率から finalPrice を算出する。
   * @return {{price: number|null, warning: string}} warning: '' | 'yellow' | 'red'
   */
  var calcPrice = function(ceiling, cost, shipping, rate) {
    if (rate === null || isNaN(rate) || isNaN(ceiling) || ceiling < 10) {
      return {price: null, warning: 'red'};
    }

    var fee        = ceiling * rate;
    var profit     = ceiling - cost - shipping - fee;
    var profitRate = profit / ceiling;

    // ① 利益率 > 20% → 20%になるよう価格を切り下げ
    if (profitRate > RATE_MAX) {
      var denom = 1 - rate - RATE_MAX;
      if (denom <= 0) return {price: null, warning: 'red'};
      var capPrice  = Math.floor((cost + shipping) / denom);
      if (capPrice < 10) return {price: null, warning: 'red'};
      var capProfit = capPrice - cost - shipping - capPrice * rate;
      if (capProfit >= MIN_PROFIT) return {price: capPrice, warning: ''};
      // cap後も利益不足ならそのまま下の処理へ（稀なケース）
      profit     = capProfit;
      profitRate = capProfit / capPrice;
    }

    // ② 利益額 ≥ 200円 → ceiling（またはcap後価格）をそのまま採用
    if (profit >= MIN_PROFIT) {
      return {price: ceiling, warning: profitRate < RATE_WARN ? 'yellow' : ''};
    }

    // ③ 利益額不足 → 最低利益確保価格（競合より高くてもOK）
    var denom2 = 1 - rate;
    if (denom2 <= 0) return {price: null, warning: 'red'};
    var minPrice = Math.ceil((cost + shipping + MIN_PROFIT) / denom2);
    if (minPrice < 10) return {price: null, warning: 'red'};
    // 価格帯をまたぐ可能性があるため1回再計算（1500円境界対策）
    var minFee    = minPrice * rate;
    var minProfit = minPrice - cost - shipping - minFee;
    if (minProfit >= MIN_PROFIT) return {price: minPrice, warning: ''};
    // 1回だけ再試行（例: 食品 1499→1501 で料率が変わる場合）
    var minPrice2  = Math.ceil((cost + shipping + MIN_PROFIT) / (1 - rate));
    var minProfit2 = minPrice2 - cost - shipping - minPrice2 * rate;
    if (minProfit2 >= MIN_PROFIT) return {price: minPrice2, warning: ''};

    return {price: null, warning: 'red'};
  };

  var rowsMissingCategory = [];
  var updated = 0;
  var skipped = 0;

  // --- Gemini カテゴリ一括推測（空カテゴリの行のみ・JAN単位で重複回避）---
  var janToGuessedCategory = {};
  if (useGeminiForCategory && colCategory !== undefined) {
    var janToNameForGuess = {};
    for (var pre = headerRowIdx + 1; pre < masterValues.length; pre++) {
      var pRow = masterValues[pre];
      var pJan = String(pRow[colJan] || '').trim();
      var pSet = String(pRow[colSetQty] || '').trim();
      if (pJan === '' || pSet === '') continue;
      var pCat = String(pRow[colCategory] || '').trim();
      if (pCat !== '') continue;
      if (!janToNameForGuess[pJan]) {
        var colNameIdx = masterColMap['商品名ベース'] !== undefined ? masterColMap['商品名ベース'] : masterColMap['商品名'];
        var pName = colNameIdx !== undefined ? String(pRow[colNameIdx] || '').trim() : '';
        if (pName) janToNameForGuess[pJan] = pName;
      }
    }
    var uniqueJansForGuess = Object.keys(janToNameForGuess);
    for (var gj = 0; gj < uniqueJansForGuess.length; gj++) {
      var gJan = uniqueJansForGuess[gj];
      Utilities.sleep(200);
      var guessedCat = guessAmazonCategoryByGemini(janToNameForGuess[gJan], baseCategoryList);
      janToGuessedCategory[gJan] = guessedCat;
      Logger.log('[menuProposeSalesPrices] カテゴリ推測 JAN=' + gJan + ' → ' + guessedCat);
    }
  }

  // --- Step1: 出品CK=TRUE の子行を JAN 単位でグループ化 ---
  var janGroups = {}; // JAN → [{rowIdx, setNum, cost, shipping, categoryVal}]
  for (var r = headerRowIdx + 1; r < masterValues.length; r++) {
    var mRow = masterValues[r];
    if (colCheckboxPrice !== undefined && mRow[colCheckboxPrice] !== true) continue;
    var jan = String(mRow[colJan] || '').trim();
    var setVal = String(mRow[colSetQty] || '').trim();
    if (jan === '' || setVal === '') continue;
    var setNum = parseInt(setVal, 10);
    if (isNaN(setNum) || setNum < 1) continue;

    var cost = parseNum(mRow[colCost]);
    if (isNaN(cost) || cost < 0) { skipped++; continue; }

    var categoryVal = colCategory !== undefined ? String(mRow[colCategory] || '').trim() : '';
    if (!categoryVal && janToGuessedCategory[jan]) {
      categoryVal = janToGuessedCategory[jan];
      masterSheet.getRange(r + 1, colCategory + 1).setValue(categoryVal).setBackground(null);
    }
    if (!categoryVal) {
      rowsMissingCategory.push(r + 1);
      continue;
    }

    var shipping = (colShipping !== undefined && mRow[colShipping] !== undefined && mRow[colShipping] !== '')
      ? parseNum(mRow[colShipping]) : 0;
    if (isNaN(shipping)) shipping = 0;

    if (!janGroups[jan]) janGroups[jan] = [];
    janGroups[jan].push({rowIdx: r, setNum: setNum, cost: cost, shipping: shipping, categoryVal: categoryVal});
  }

  // --- Step2: JAN 単位で ceiling 計算 → 価格算出 → 書き込み ---
  var janList = Object.keys(janGroups);
  for (var jIdx = 0; jIdx < janList.length; jIdx++) {
    var jan2 = janList[jIdx];
    var blockIndex = janToBlock[jan2];
    if (blockIndex === undefined) {
      for (var sk = 0; sk < janGroups[jan2].length; sk++) skipped++;
      continue;
    }

    var compMap = getCachedBlock(blockIndex); // setNum(str) → {price, asin}

    // 競合が存在するセット数の昇順リスト（数値）
    var compSetNums = [];
    var compKeys = Object.keys(compMap);
    for (var ck = 0; ck < compKeys.length; ck++) {
      var cn = parseInt(compKeys[ck], 10);
      if (!isNaN(cn) && cn > 0 && compMap[compKeys[ck]] && compMap[compKeys[ck]].price >= 10) {
        compSetNums.push(cn);
      }
    }
    compSetNums.sort(function(a, b) { return a - b; });

    var rows = janGroups[jan2];
    rows.sort(function(a, b) { return a.setNum - b.setNum; });

    for (var ri = 0; ri < rows.length; ri++) {
      var row = rows[ri];
      var sn  = row.setNum;

      // ---- ceiling_A: 同セット数の競合価格 - 1 ----
      var ceiling_A = Infinity;
      if (compMap[sn] && compMap[sn].price >= 10) {
        ceiling_A = compMap[sn].price - 1;
      }

      // ---- ceiling_C: 直近下の競合単価 × 自セット数 - 1 ----
      var ceiling_C = Infinity;
      var lowerNums = [];
      for (var ci = 0; ci < compSetNums.length; ci++) {
        if (compSetNums[ci] < sn) lowerNums.push(compSetNums[ci]);
      }
      if (lowerNums.length > 0) {
        var nearestLower   = lowerNums[lowerNums.length - 1];
        var lowerUnitPrice = compMap[nearestLower].price / nearestLower;
        ceiling_C = Math.floor(lowerUnitPrice * sn) - 1;
      } else if (ceiling_A === Infinity) {
        // 下に競合なし かつ 自セット数にも競合なし（最小の穴）→ 上の競合単価を使用
        var upperNums = [];
        for (var ci2 = 0; ci2 < compSetNums.length; ci2++) {
          if (compSetNums[ci2] > sn) upperNums.push(compSetNums[ci2]);
        }
        if (upperNums.length > 0) {
          var nearestUpper   = upperNums[0];
          var upperUnitPrice = compMap[nearestUpper].price / nearestUpper;
          ceiling_C = Math.floor(upperUnitPrice * sn) - 1;
        }
      }

      var ceiling = Math.min(ceiling_A, ceiling_C);
      if (!isFinite(ceiling) || ceiling < 10) {
        Logger.log('[価格提案] ceiling未確定 スキップ JAN=' + jan2 + ' set=' + sn);
        skipped++;
        continue;
      }

      // ---- Amazon 価格 ----
      var amzRate   = getAmazonCommissionRateForPrice(rateMap, row.categoryVal, ceiling);
      if (amzRate === null) {
        rowsMissingCategory.push(row.rowIdx + 1);
        continue;
      }
      var amzResult = calcPrice(ceiling, row.cost, row.shipping, amzRate);

      Logger.log('[価格提案] JAN=' + jan2 + ' set=' + sn + ' ceiling=' + ceiling +
        ' amz=' + (amzResult.price !== null ? amzResult.price : 'skip'));

      // ---- 書き込み（Amazon のみ。楽天・Yahoo!は別フェーズでAPI競合取得後に計算）----
      var amzCell = masterSheet.getRange(row.rowIdx + 1, colPriceAmazon + 1);

      if (amzResult.price !== null) {
        amzCell.setValue(amzResult.price)
               .setBackground(amzResult.warning === 'yellow' ? '#fff2cc' : null);
        if (colAmazonFee !== undefined) {
          var rateForFee = getAmazonCommissionRateForPrice(rateMap, row.categoryVal, amzResult.price);
          if (rateForFee !== null) {
            var feeAmount = Math.round(amzResult.price * rateForFee);
            masterSheet.getRange(row.rowIdx + 1, colAmazonFee + 1).setValue(feeAmount);
          }
        }
      } else {
        amzCell.setBackground('#ffcccc');
        skipped++;
        continue;
      }
      updated++;
    }
  }

  // ---- カテゴリ不足の警告 ----
  if (rowsMissingCategory.length > 0) {
    var msg = '「amazon カテゴリー」が空の行があります。該当行を赤くしました。行: ' +
              rowsMissingCategory.slice(0, 20).join(', ');
    if (rowsMissingCategory.length > 20) msg += ' …他' + (rowsMissingCategory.length - 20) + '行';
    for (var mi = 0; mi < rowsMissingCategory.length; mi++) {
      if (colCategory !== undefined) {
        masterSheet.getRange(rowsMissingCategory[mi], colCategory + 1).setBackground('#ffcccc');
      }
    }
    SpreadsheetApp.getUi().alert(msg);
  }

  SpreadsheetApp.getActive().toast(
    '販売価格を提案しました。更新: ' + updated + '行' +
    (skipped > 0 ? '  スキップ: ' + skipped + '行' : ''),
    '販売価格提案 完了',
    8
  );
}

/**
 * CPO値決め用プロンプト本文。docs/CPO_PROMPT.md と同一。{{ }} は getCPOMapping とマスタ値で置換する。
 * @return {string}
 */
function getCPOPromptTemplate() {
  return '# Role\n' +
    'あなたは、年商10億円規模のECブランドを牽引する「最高プライシング責任者（CPO）」です。競合を出し抜き、配送の限界効率を突き詰め、利益を1円単位で最大化する値決めを行います。\n\n' +
    '# Objective\n' +
    '入力された原価・送料・競合価格・市場データを統合し、顧客が「まとめ買いしたくなる（Volume Discount）」且つ「利益が最大化される」完璧なセット数と価格のポートフォリオを提案してください。モール名（Amazon / 楽天 / Yahoo!）ごとに同じルールで価格を算出してください。\n\n' +
    '# Pricing Rules & Logic (絶対厳守)\n' +
    '1. **厳密な利益計算（全て税込）**\n' +
    '   - 利益額 = 販売価格(税込) - セット卸値(税込) - 梱包箱コスト - 確定送料 - (販売価格 × (モール手数料率 + 販促費率))\n' +
    '   - 販促費率：ポイント還元・広告費想定（%）。未指定時は Amazon=3%、楽天=12%、Yahoo=3% とせよ。\n' +
    '   - Amazonは価格帯（1500円以下/以上）で手数料率が変わる場合はそれを考慮すること。\n\n' +
    '2. **利益のガードレール**\n' +
    '   - 利益額：全セットで最低200円/セット。\n' +
    '   - 利益率：最低8%、目安15%、上限30%。\n\n' +
    '3. **配送・セット数**\n' +
    '   - 「送料の壁」があれば、1個あたり配送コストが最安になるセット数を主力候補とする。\n' +
    '   - 1個あたり単価はセット数が増えるごとに必ず安くなる（または同額）グラデーションにすること。単価の逆転は禁止。\n\n' +
    '4. **競合**\n' +
    '   - 競合があるセット数は原則「競合価格 - 1円」。\n' +
    '   - 競合が極端に安くガードレールを満たせない場合は追従せず、他セットへ誘導する設計とする。\n' +
    '   - 競合がいないセット数（穴）は高利益率（20%〜）を狙う。ただし、前のセット数の1個あたり単価を上回ってはならない（単価の逆転は禁止）。\n\n' +
    '5. **心理価格**\n' +
    '   - 理論値を末尾80/90/00円や大台を割る価格（1,980円等）に調整すること。\n\n' +
    '# Input Data\n' +
    '## 基本情報\n' +
    '- 商品名：{{ 商品名 }}\n' +
    '- 対象モール：{{ モール名 }}\n' +
    '- モール販売手数料率：{{ 手数料率 }} %\n' +
    '- 販促費率：{{ 販促費率 }} %\n\n' +
    '## コスト・配送\n' +
    '- 梱包箱コスト（1出荷あたり）：{{ 梱包箱コスト }} 円\n' +
    '- 送料の壁：{{ 送料の壁テキスト }}\n\n' +
    '## セット数別データ\n' +
    '```\n' +
    'セット数\tセット卸値(税込)\t確定送料\t競合価格({{ モール名 }})\n' +
    '{{ セット数別テーブル }}\n' +
    '```\n\n' +
    '# Output (厳守)\n' +
    '**必ず最初に 5. 【JSONデータ】のJSONブロックを出力し、その後に 1～4 の説明を書くこと。** システムは先頭の ```json ブロックのみを読み取り価格を反映する。表だけでは反映できない。JSONブロックを省略しないこと。\n\n' +
    '## 5. 【JSONデータ】（必ず先に出力）\n' +
    '**【重要】単価の逆転は禁止。** セット数が増えるほど1個あたり単価は必ず安くなるか同額にすること。これに反する価格をJSONに含めてはならない。\n\n' +
    '必ず以下のJSONブロックを1つだけ含めること。価格は税込・整数。\n' +
    '```json\n' +
    '{"amazon":[{"setCount":1,"price":680},{"setCount":2,"price":900}],"rakuten":[],"yahoo":[]}\n' +
    '```\n\n' +
    '## 1. 【戦略的価格マトリクス】(Markdown表)\n' +
    '列：セット数 / 販売価格(税込) / 1個あたり単価 / 利益額 / 利益率 / 役割\n\n' +
    '## 2. 【配送コストとマジックナンバーの解説】\n' +
    '簡潔に。\n\n' +
    '## 3. 【対競合戦略と誘導ロジック】\n' +
    '簡潔に。\n\n' +
    '## 4. 【ネクストアクション】\n' +
    '簡潔に。';
}

/**
 * 00_設定マスタ A92～C100 を読み、プレースホルダ名（A列）→ マスタ列名（B列）のマップを返す。
 * A列は "商品名" または "{{ 商品名 }}" の形式。置換時は "{{ キー }}" で検索する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Object<string, string>} プレースホルダキー（スペース含む） → マスタ列名
 */
function getCPOMapping(ss) {
  var sheet = ss.getSheetByName(SETTINGS_MASTER_SHEET_NAME);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < CPO_MAPPING_FIRST_ROW) return {};
  var rng = sheet.getRange(CPO_MAPPING_FIRST_ROW, 1, Math.min(CPO_MAPPING_LAST_ROW, lastRow), 2);
  var rows = rng.getValues();
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var a = String(rows[i][0] || '').trim();
    var b = String(rows[i][1] || '').trim();
    if (a === '' || b === '') continue;
    var key = a.replace(/^\{\{\s*|\s*\}\}$/g, '').trim();
    if (key) map[key] = b;
  }
  return map;
}

/**
 * 1 JAN 分のCPOプロンプトを組み立てる。マッピングに従いマスタから値を取り {{ }} を置換する。
 * @param {Array<Array>} masterValues マスタ全セル
 * @param {number} headerRowIdx ヘッダー行（0-based）
 * @param {Object<string, number>} masterColMap 列名→列インデックス（0-based）
 * @param {string} jan 対象JAN
 * @param {number} parentRowIdx 親行の行インデックス（0-based）
 * @param {Array<{rowIdx: number, setNum: number, cost: number, shipping: number, competitivePrice: number|string, boxCost: number|string, feeRate: number|string}>} childRows 子行情報
 * @param {Object<string, string>} mapping getCPOMapping の戻り値
 * @return {string} 置換済みプロンプト
 */
function buildCPOPromptForJAN(masterValues, headerRowIdx, masterColMap, jan, parentRowIdx, childRows, mapping) {
  var template = getCPOPromptTemplate();
  var parentRow = masterValues[parentRowIdx] || [];

  function getVal(key, row, def) {
    var colName = mapping[key];
    if (!colName) return (def !== undefined ? def : '');
    var colIdx = masterColMap[colName];
    if (colIdx === undefined) return (def !== undefined ? def : '');
    var v = row[colIdx];
    if (v === undefined || v === null) return (def !== undefined ? def : '');
    return String(v).trim();
  }

  var productName = getVal('商品名', parentRow, '');
  var mallName = getVal('モール名', parentRow, 'Amazon');
  if (mallName === '') mallName = 'Amazon';

  var firstChild = childRows.length > 0 ? childRows[0] : null;
  var feeRate = firstChild && firstChild.feeRate !== undefined && firstChild.feeRate !== '' ? firstChild.feeRate : getVal('手数料率', parentRow, '');
  var promoRate = getVal('販促費率', parentRow, '3');
  var boxCost = firstChild && firstChild.boxCost !== undefined && firstChild.boxCost !== '' ? firstChild.boxCost : getVal('梱包箱コスト', parentRow, '0');
  var shippingWall = getVal('送料の壁テキスト', parentRow, '');

  var tableLines = [];
  for (var c = 0; c < childRows.length; c++) {
    var ch = childRows[c];
    var comp = ch.competitivePrice !== undefined && ch.competitivePrice !== '' ? ch.competitivePrice : '';
    tableLines.push(ch.setNum + '\t' + ch.cost + '\t' + ch.shipping + '\t' + comp);
  }
  var setTable = tableLines.join('\n');

  var replacements = {
    '商品名': productName,
    'モール名': mallName,
    '手数料率': feeRate,
    '販促費率': promoRate,
    '梱包箱コスト': boxCost,
    '送料の壁テキスト': shippingWall,
    'セット数別テーブル': setTable
  };
  var result = template;
  for (var key in replacements) {
    if (replacements.hasOwnProperty(key)) {
      result = result.replace(new RegExp('\\{\\{\\s*' + key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*\\}\\}', 'g'), String(replacements[key]));
    }
  }
  return result;
}

/**
 * CPO用プロンプトをGeminiに送り、返答全文を返す。
 * @param {string} prompt 置換済みプロンプト
 * @return {string} 返答テキスト。失敗時は ''
 */
function callGeminiForCPO(prompt) {
  var apiKey = getGeminiApiKey();
  if (!apiKey) return '';
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL_GEMINI +
      ':generateContent?key=' + encodeURIComponent(apiKey);
    var body = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 16384 }
    };
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var bodyText = res.getContentText();
    if (code !== 200) {
      Logger.log('[CPO] API HTTP ' + code + ' body先頭300字=' + (bodyText ? String(bodyText).substring(0, 300) : ''));
      return '';
    }
    var json = JSON.parse(bodyText);
    var candidates = json.candidates;
    if (!candidates || candidates.length === 0) {
      var feedback = (json.promptFeedback || {}).blockReason || (json.promptFeedback || {}).blockReasonMessage || '';
      Logger.log('[CPO] candidatesが空 blockReason等=' + feedback + ' body先頭200字=' + (bodyText ? String(bodyText).substring(0, 200) : ''));
      return '';
    }
    var parts = (candidates[0].content || {}).parts;
    var rawText = parts && parts[0] ? (parts[0].text || '') : '';
    if (!rawText) {
      Logger.log('[CPO] 返答テキストが空 finishReason=' + (candidates[0].finishReason || ''));
    }
    return rawText;
  } catch (e) {
    Logger.log('[CPO] Gemini呼び出し失敗: ' + (e && e.message));
    return '';
  }
}

/**
 * CPO返答テキストから「## 5. 【JSONデータ】」以降のJSONブロックを除去した人間向けテキストと、JSON部分を返す。
 * セルには textForCell を書き、Logger には全文・JSON部分を出すために使う。
 * @param {string} responseText 返答全文
 * @return {{ textForCell: string, jsonPart: string }}
 */
function stripCPOJsonFromResponse(responseText) {
  var textForCell = responseText || '';
  var jsonPart = '';
  var idx5 = textForCell.indexOf('## 5.');
  if (idx5 >= 0) {
    var idxJson = textForCell.indexOf('```json', idx5);
    if (idxJson >= 0) {
      var startBlock = idx5;
      var endBacktick = textForCell.indexOf('```', idxJson + 7);
      if (endBacktick >= 0) {
        jsonPart = textForCell.substring(idxJson + 7, endBacktick).trim();
        textForCell = (textForCell.substring(0, startBlock) + textForCell.substring(endBacktick + 3)).replace(/\n{3,}/g, '\n\n').trim();
      }
    }
  }
  return { textForCell: textForCell, jsonPart: jsonPart };
}

/**
 * CPO返答テキスト末尾のJSONブロックをパースし、amazon配列を返す。
 * プロンプトで「先にJSONを出力」しているため、先頭の ```json を優先。無い場合は末尾の ```json または正規表現で検索。
 * JSONが無い場合は Markdown 表（| セット数 | 販売価格(税込) | ...）からセット数・価格を抽出するフォールバックを行う。
 * @param {string} responseText 返答全文
 * @return {Array<{setCount: number, price: number}>} amazon 配列。パース失敗時は []
 */
function parseCPOJson(responseText) {
  if (!responseText || typeof responseText !== 'string') return [];
  var str = responseText;
  var jsonBlock = '';
  // 先頭の ```json を優先（プロンプトで「先にJSON出力」のため）
  var startFirst = str.indexOf('```json');
  if (startFirst >= 0) {
    var endFirst = str.indexOf('```', startFirst + 7);
    jsonBlock = endFirst >= 0 ? str.substring(startFirst + 7, endFirst).trim() : str.substring(startFirst + 7).trim();
  }
  if (!jsonBlock) {
    var startLast = str.lastIndexOf('```json');
    if (startLast >= 0) {
      var endLast = str.indexOf('```', startLast + 7);
      jsonBlock = endLast >= 0 ? str.substring(startLast + 7, endLast).trim() : str.substring(startLast + 7).trim();
    }
  }
  if (!jsonBlock) {
    var m = str.match(/\{\s*"amazon"\s*:\s*\[[\s\S]*?\](?:\s*,\s*"rakuten"\s*:\s*\[[\s\S]*?\])?(?:\s*,\s*"yahoo"\s*:\s*\[[\s\S]*?\])?\s*\}/);
    if (m) jsonBlock = m[0];
  }
  if (jsonBlock) {
    try {
      var parsed = JSON.parse(jsonBlock);
      if (Array.isArray(parsed.amazon) && parsed.amazon.length > 0) {
        return parsed.amazon;
      }
    } catch (_) {
      Logger.log('[CPO] JSONパース失敗。jsonBlock先頭150字=' + (jsonBlock ? String(jsonBlock).substring(0, 150) : ''));
    }
  }
  // フォールバック: Markdown表から セット数 | 販売価格(税込) の行を抽出（例: | 1 | 780円 | または | 2 | **1,380円** |）
  var fromTable = parseCPOMarkdownTable(str);
  if (fromTable.length > 0) {
    Logger.log('[CPO] JSON未検出のためMarkdown表から復元。件数=' + fromTable.length);
    return fromTable;
  }
  Logger.log('[CPO] JSONブロック未検出。responseText末尾250字=' + (str ? String(str).slice(-250) : ''));
  return [];
}

/**
 * CPO返答のMarkdown表（戦略的価格マトリクス）から setCount と price を抽出する。
 * 行形式: | 1 | 780円 | ... または | 2 | **1,380円** | ...
 * @param {string} str 返答全文
 * @return {Array<{setCount: number, price: number}>}
 */
function parseCPOMarkdownTable(str) {
  var out = [];
  var lines = str.split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    // セット数 | 価格円 の形式（2列目が販売価格）。** やカンマを除去
    var match = line.match(/^\|\s*(\d+)\s*\|\s*(?:\*\*)?([\d,]+)(?:\*\*)?\s*円/);
    if (match) {
      var setCount = parseInt(match[1], 10);
      var priceStr = (match[2] || '').replace(/,/g, '');
      var price = parseInt(priceStr, 10);
      if (!isNaN(setCount) && setCount >= 1 && !isNaN(price) && price >= 0) {
        out.push({ setCount: setCount, price: price });
      }
    }
  }
  return out;
}

// ========== セット数ポートフォリオ提案（Gemini） ==========
// docs/SET_COUNT_PROPOSAL_PROMPT.md と同一内容。AI情報取得data 参照時は
// ▼商品マスタ(人間作業用)の商品名ベースとAI情報取得dataの商品名（または商品名ベース）が一致した行を用いる。

/**
 * セット数ポートフォリオ提案用プロンプト本文。docs/SET_COUNT_PROPOSAL_PROMPT.md と同一。
 * @return {string}
 */
function getSetCountProposalPromptTemplate() {
  return '# Role\n' +
    'あなたは年商10億円超のECブランドを支える「戦略的マーチャンダイザー（MD）」兼「物流コンサルタント」です。知識・配送コストの境界線・競合の隙間を統合し、最も利益率が高く、かつ顧客が納得する「セット数展開」をロジカルに構築します。\n\n' +
    '# Objective\n' +
    '入力された商品情報を元に、以下の【4つのフェーズ】を経て、自社が展開すべき「必勝セット数ラインナップ」を提案してください。\n\n' +
    '# Investigation & Logic (思考プロセス)\n\n' +
    '## Phase 1: 消費サイクル・共有性の徹底解析\n' +
    '- 知識に基づき、当該商品の「1日あたりの標準使用量」および「使用頻度」を特定せよ。\n' +
    '- 【業務用】の定義：1人使用で1ヶ月〜2ヶ月分に相当する量（最低2枠）。【上限】：3ヶ月分を絶対上限とする。\n\n' +
    '## Phase 2: 物流効率（送料の壁）との照合\n' +
    '- 提供された「送料の壁」データを参照し、最も1個あたりの送料が安くなる「物流黄金セット数」を特定せよ。\n\n' +
    '## Phase 3: 競合隙間戦略（Gap Analysis）\n' +
    '- **競合店が出品しているセット数はすべて必ずラインナップに含めること。** 後で価格決定の判断のため、入力の「競合のセット数リスト」のすべてを欠かさず含めて提案すること。競合の隙間を自社独自の高利益枠として設定せよ。\n\n' +
    '## Phase 4: ポートフォリオの最適配置\n' +
    '- 少量（集客）から大量（業務用）まで、バランスよく（緩やかな上昇曲線）配置せよ。\n\n' +
    '# Input Data\n' +
    '- 商品名：{{ 商品名 }}\n' +
    '- カテゴリー：{{ カテゴリー }}\n' +
    '- 商品仕様（重量/サイズ/賞味・使用期限）：{{ 商品仕様 }}\n' +
    '- 配送コストデータ（送料の壁）：{{ 送料の壁テキスト }}\n' +
    '- 競合のセット数リスト：{{ 競合のセット数リスト }}\n' +
    '- ターゲット属性：{{ ターゲット属性 }}\n\n' +
    '# Output\n' +
    '1. 【市場・消費データ分析】 2. 【戦略的セット数提案マトリクス】(表) 3. 【物流・利益効率のアドバイス】 4. 【マーケティングシナリオ】\n\n' +
    '# Constraints\n' +
    '- **競合店が出品しているセット数はすべて必ず含めること。** 入力の競合のセット数リストに記載されたセット数は1つも漏らさず含めること。\n' +
    '- その上でセット数は**5〜8パターン程度**（競合が多い場合はそれを上回ってもよい）。**業務用2枠は必ず含めること。**\n\n' +
    '# JSON出力（厳守）\n' +
    '必ず最後に以下のJSONブロックを1つだけ含めること。\n' +
    '```json\n{"setCounts": [1, 2, 3, 6, 12, 24]}\n```';
}

/**
 * セット数提案プロンプトの {{ }} を置換する。
 * @param {{productName: string, category: string, specs: string, shippingWallText: string, competitiveSetList: Array<number>, targetAttribute: string}} input
 * @return {string}
 */
function buildSetCountProposalPrompt(input) {
  var tpl = getSetCountProposalPromptTemplate();
  var listStr = (input.competitiveSetList && input.competitiveSetList.length) ? input.competitiveSetList.join(', ') : '（なし）';
  var replacements = {
    '商品名': input.productName || '',
    'カテゴリー': input.category || '',
    '商品仕様': input.specs || '',
    '送料の壁テキスト': input.shippingWallText || '',
    '競合のセット数リスト': listStr,
    'ターゲット属性': input.targetAttribute || '単身者/家族'
  };
  var result = tpl;
  for (var key in replacements) {
    if (replacements.hasOwnProperty(key)) {
      result = result.replace(new RegExp('\\{\\{\\s*' + key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*\\}\\}', 'g'), String(replacements[key]));
    }
  }
  return result;
}

/**
 * セット数ポートフォリオをGeminiに問い合わせ、setCounts と返答全文を返す。
 * @param {string} prompt 置換済みプロンプト
 * @return {{ setCounts: number[]|null, responseText: string }} 成功時は setCounts と responseText、失敗時は setCounts: null と responseText（空または生テキスト）
 */
function getSetCountProposalByGemini(prompt) {
  var empty = { setCounts: null, responseText: '' };
  var apiKey = getGeminiApiKey();
  if (!apiKey) return empty;
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL_GEMINI +
      ':generateContent?key=' + encodeURIComponent(apiKey);
    var body = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 8192 }
    };
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText());
    var rawText = (((json.candidates || [])[0] || {}).content || {}).parts;
    rawText = rawText ? (rawText[0].text || '') : '';
    if (!rawText) return empty;
    // JSONブロック抽出（```json ... ``` または "setCounts": [...] を含むブロック）
    var jsonBlock = '';
    var start = rawText.lastIndexOf('```json');
    if (start >= 0) {
      var end = rawText.indexOf('```', start + 7);
      jsonBlock = end >= 0 ? rawText.substring(start + 7, end).trim() : rawText.substring(start + 7).trim();
    } else {
      var m = rawText.match(/\{\s*"setCounts"\s*:\s*\[[\s\S]*?\]\s*\}/);
      if (m) jsonBlock = m[0];
    }
    if (!jsonBlock) {
      Logger.log('[セット数提案] JSONブロック未検出。rawText末尾250字=' + rawText.slice(-250));
      return { setCounts: null, responseText: rawText };
    }
    var parsed = JSON.parse(jsonBlock);
    var arr = parsed.setCounts;
    if (!Array.isArray(arr) || arr.length < 3 || arr.length > 20) return { setCounts: null, responseText: rawText };
    var nums = [];
    for (var i = 0; i < arr.length; i++) {
      var n = parseInt(arr[i], 10);
      if (isNaN(n) || n < 1) return { setCounts: null, responseText: rawText };
      nums.push(n);
    }
    nums.sort(function(a, b) { return a - b; });
    return { setCounts: nums, responseText: rawText };
  } catch (e) {
    Logger.log('[セット数提案] Gemini呼び出し失敗: ' + (e && e.message) + ' jsonBlock先頭100=' + (typeof jsonBlock !== 'undefined' ? String(jsonBlock).substring(0, 100) : ''));
    return empty;
  }
}

/**
 * 出品CK=TRUE の商品を JAN 単位で1つずつ CPO（Gemini）で価格提案し、親行に amazon価格戦略、子行に 販売価格amazon を書き込む。
 * メニュー「③ 販売価格の調整 → CPOで価格提案（Gemini）」から実行。
 */
function menuCPOProposePrices() {
  var ui = SpreadsheetApp.getUi();
  if (!ui) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    ui.alert('「' + MASTER_SHEET_NAME + '」シートが見つかりません。');
    return;
  }
  if (!getGeminiApiKey()) {
    ui.alert('GEMINI_API_KEY が Script Properties に設定されていません。');
    return;
  }
  var settingsSheet = ss.getSheetByName(SETTINGS_MASTER_SHEET_NAME);
  if (!settingsSheet) {
    ui.alert('「' + SETTINGS_MASTER_SHEET_NAME + '」シートが見つかりません。');
    return;
  }

  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(masterValues.length, 25); hr++) {
    if ((masterValues[hr] || []).indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = hr; break; }
  }
  if (headerRowIdx === -1) {
    ui.alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  var colJan = masterColMap['JANコード'];
  var colSetQty = masterColMap[COL_MASTER_TOTAL_QTY];
  var colCost = masterColMap[COL_COST_SET_TAX_IN] !== undefined ? masterColMap[COL_COST_SET_TAX_IN] : masterColMap[COL_COST_TAX_IN];
  var colShipping = masterColMap[COL_SHIPPING_FIXED] !== undefined ? masterColMap[COL_SHIPPING_FIXED] : masterColMap[COL_SHIPPING];
  var colCompetitive = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  var colCheckbox = masterColMap[CHECKBOX_HEADER_NAME];
  var colStrategy = masterColMap[COL_AMAZON_STRATEGY];
  var colPriceAmazon = masterColMap[COL_PRICE_AMAZON];
  var colBoxCost = masterColMap['梱包箱コスト'];
  var colFeeRate = masterColMap['amazon手数料率'];
  var colSizeFba = masterColMap['サイズ＆自己発/FBA'];
  var colBoxSpec = masterColMap['梱包箱指定'];

  if (colJan === undefined || colSetQty === undefined || colCost === undefined) {
    ui.alert('マスタに「JANコード」「' + COL_MASTER_TOTAL_QTY + '」または卸値列がありません。');
    return;
  }
  if (colStrategy === undefined) {
    ui.alert('マスタに「' + COL_AMAZON_STRATEGY + '」列がありません。');
    return;
  }
  if (colPriceAmazon === undefined) {
    ui.alert('マスタに「' + COL_PRICE_AMAZON + '」列がありません。');
    return;
  }

  var mapping = getCPOMapping(ss);
  var parseNum = function(v) {
    if (v === undefined || v === null || v === '') return NaN;
    var n = Number(String(v).replace(/,/g, ''));
    return isNaN(n) ? NaN : n;
  };

  // 出品CK=TRUE の行を JAN 単位でグループ化。親行（A.セット商品数が空）と子行を分ける。
  var janToGroup = {};
  for (var r = headerRowIdx + 1; r < masterValues.length; r++) {
    var row = masterValues[r];
    if (colCheckbox !== undefined && row[colCheckbox] !== true) continue;
    var jan = String(row[colJan] || '').trim();
    if (jan === '') continue;
    var setVal = String(row[colSetQty] || '').trim();
    var cost = parseNum(row[colCost]);
    if (isNaN(cost) || cost < 0) cost = 0;
    var shipping = (colShipping !== undefined && row[colShipping] !== undefined && row[colShipping] !== '')
      ? parseNum(row[colShipping]) : 0;
    if (isNaN(shipping)) shipping = 0;
    var competitive = row[colCompetitive];
    if (competitive !== undefined && competitive !== null && competitive !== '') competitive = String(competitive).trim();
    else competitive = '';
    var boxCost = (colBoxCost !== undefined && row[colBoxCost] !== undefined && row[colBoxCost] !== '')
      ? String(row[colBoxCost]).trim() : '';
    var feeRate = (colFeeRate !== undefined && row[colFeeRate] !== undefined && row[colFeeRate] !== '')
      ? String(row[colFeeRate]).trim() : '';

    if (setVal === '') {
      if (!janToGroup[jan]) janToGroup[jan] = { parentRowIdx: r, children: [] };
      else if (janToGroup[jan].parentRowIdx === undefined) janToGroup[jan].parentRowIdx = r;
    } else {
      var setNum = parseInt(setVal, 10);
      if (isNaN(setNum) || setNum < 1) continue;
      if (!janToGroup[jan]) janToGroup[jan] = { parentRowIdx: undefined, children: [] };
      janToGroup[jan].children.push({
        rowIdx: r,
        setNum: setNum,
        cost: cost,
        shipping: shipping,
        competitivePrice: competitive,
        boxCost: boxCost,
        feeRate: feeRate
      });
    }
  }

  var janList = [];
  for (var j in janToGroup) {
    if (janToGroup[j].parentRowIdx === undefined || janToGroup[j].children.length === 0) continue;
    janList.push(j);
  }
  if (janList.length === 0) {
    ui.alert('出品CK=TRUE かつ 親行（A.セット商品数が空）と子行が揃った JAN が1件もありません。');
    return;
  }

  // 子SKUの「サイズ＆自己発/FBA」「梱包箱指定」が未入力の場合は実行前に確認
  if (colSizeFba !== undefined && colBoxSpec !== undefined) {
    var jansWithBlank = [];
    for (var jb = 0; jb < janList.length; jb++) {
      var jJan = janList[jb];
      var jGroup = janToGroup[jJan];
      for (var jc = 0; jc < jGroup.children.length; jc++) {
        var childRow = masterValues[jGroup.children[jc].rowIdx];
        var v1 = childRow[colSizeFba]; var v2 = childRow[colBoxSpec];
        var blank1 = (v1 === undefined || v1 === null || String(v1).trim() === '');
        var blank2 = (v2 === undefined || v2 === null || String(v2).trim() === '');
        if (blank1 || blank2) {
          if (jansWithBlank.indexOf(jJan) === -1) jansWithBlank.push(jJan);
          break;
        }
      }
    }
    if (jansWithBlank.length > 0) {
      var msg = '「サイズ＆自己発/FBA」または「梱包箱指定」が未入力の子SKUがあるJANがあります。\n\n送料は入力されていますか？\n（いいえ：実行しません　はい：このまま実行します）';
      var result = ui.alert('送料の確認', msg, ui.ButtonSet.YES_NO);
      if (result !== ui.Button.YES) {
        return;
      }
    }
  }

  Logger.log('[CPO] 対象JAN数=' + janList.length);
  var updated = 0;
  var failed = 0;
  for (var ji = 0; ji < janList.length; ji++) {
    var jan = janList[ji];
    var group = janToGroup[jan];
    group.children.sort(function(a, b) { return a.setNum - b.setNum; });
    var prompt = buildCPOPromptForJAN(masterValues, headerRowIdx, masterColMap, jan, group.parentRowIdx, group.children, mapping);
    var responseText = callGeminiForCPO(prompt);
    if (!responseText) {
      Logger.log('[CPO] JAN=' + jan + ' Gemini返答なし');
      failed++;
      continue;
    }
    Logger.log('[CPO] JAN=' + jan + ' 返答長=' + responseText.length);
    var stripped = stripCPOJsonFromResponse(responseText);
    if (stripped.jsonPart) Logger.log('[CPO] JSON部分=' + stripped.jsonPart);
    Logger.log('[CPO] 返答全文(JSON含む)=' + responseText);
    masterSheet.getRange(group.parentRowIdx + 1, colStrategy + 1).setValue(stripped.textForCell);
    var amazonPrices = parseCPOJson(responseText);
    var masterSetNums = group.children.map(function(c) { return c.setNum; });
    var parsedSetCounts = (amazonPrices && amazonPrices.length) ? amazonPrices.map(function(ap) { return ap.setCount; }) : [];
    Logger.log('[CPO] JAN=' + jan + ' パース後価格件数=' + (amazonPrices ? amazonPrices.length : 0) + ' マスタのセット数=' + JSON.stringify(masterSetNums) + ' パース結果のセット数=' + JSON.stringify(parsedSetCounts));
    for (var ai = 0; ai < (amazonPrices || []).length; ai++) {
      var ap = amazonPrices[ai];
      var setCount = ap.setCount;
      var price = ap.price;
      for (var ci = 0; ci < group.children.length; ci++) {
        if (group.children[ci].setNum === setCount) {
          masterSheet.getRange(group.children[ci].rowIdx + 1, colPriceAmazon + 1).setValue(price);
          updated++;
          break;
        }
      }
    }
    for (var ci = 0; ci < group.children.length; ci++) {
      var setNum = group.children[ci].setNum;
      var found = (amazonPrices || []).some(function(ap) { return ap.setCount === setNum; });
      if (!found) Logger.log('[CPO] JAN=' + jan + ' setNum=' + setNum + ' に価格を反映できませんでした（パース結果に setCount=' + setNum + ' が含まれていません）');
    }
    if (ji < janList.length - 1) Utilities.sleep(500);
  }

  SpreadsheetApp.getActive().toast(
    'CPO価格提案を完了しました。更新: ' + updated + '行' + (failed > 0 ? '  失敗: ' + failed + '商品' : ''),
    'CPO 完了',
    8
  );
}

/**
 * ASIN貼り付けシートの指定ブロックから、ASIN列の値（競合ASIN）を取得する。先頭の非空ASINまたはカンマ区切りで返す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 「ASIN貼り付け（Keepa用）」シート
 * @param {number} blockIndex 0-based ブロック番号
 * @return {string} ブロック内のASIN（1件ならそのまま、複数ならカンマ区切り）。無い場合は ''
 */
function getAsinsFromAsinPasteBlock(sheet, blockIndex) {
  var data = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(data.length, 5); hr++) {
    var rowData = data[hr] || [];
    for (var hc = 0; hc < rowData.length; hc++) {
      if (String(rowData[hc]).trim() === 'ASIN') { headerRowIdx = hr; break; }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) return '';
  var headers = data[headerRowIdx];
  var asinColumns = [];
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'ASIN') asinColumns.push(c);
  }
  if (blockIndex < 0 || blockIndex >= asinColumns.length) return '';
  var colAsin = asinColumns[blockIndex];
  var dataStartRow = headerRowIdx + 1;
  var asins = [];
  for (var r = dataStartRow; r < data.length; r++) {
    var v = data[r][colAsin];
    if (v === null || v === undefined) continue;
    var s = String(v).trim();
    if (s !== '') asins.push(s);
  }
  Logger.log('[getAsinsFromBlock] block=' + blockIndex + ' headerRowIdx=' + headerRowIdx +
    ' colAsin=' + colAsin + ' dataStartRow=' + dataStartRow + ' asins=' + JSON.stringify(asins));
  return asins.length ? asins.join(',') : '';
}
function menuProposeSetCountFromAsinPasteSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「' + ASIN_PASTE_SHEET_NAME + '」シートがありません。');
    return;
  }
  if (ss.getActiveSheet().getSheetName() !== ASIN_PASTE_SHEET_NAME) {
    sheet.activate();
  }
  var result = getAggregatedSetCountsFromAsinPasteSheet(sheet);
  if (!result) {
    SpreadsheetApp.getUi().alert('セット数のデータがありません。3行目以降の「セット数」列にKeepa取得結果があるか確認してください。');
    return;
  }
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) lastCol = 1;
  var colCompetitive = lastCol + 1;
  var colHole = lastCol + 2;
  sheet.getRange(2, colCompetitive).setValue('競合セット数');
  sheet.getRange(2, colHole).setValue('穴のセット数提案');
  sheet.getRange(2, colCompetitive, 2, colHole).setFontWeight('bold');
  sheet.getRange(3, colCompetitive).setValue(result.competitive.join(', '));
  sheet.getRange(3, colHole).setValue(result.hole.join(', '));
  SpreadsheetApp.getActive().toast(
    '競合セット数: ' + result.competitive.join(', ') + '／穴の提案: ' + result.hole.join(', '),
    'セット数提案を書き出しました',
    6
  );
}

/**
 * 【優先4】▼商品マスタ(人間作業用) の◎が付いている行を対象に、
 * ASIN貼り付け（Keepa用）シートのブロック順で競合セット数・穴のセット数提案をマスタの固定列に書き込み、
 * 続けて価格・セット数提案を反映する。セット数別最低価格はブロック内データでGASが判定済み（競合価格列を参照して提案）。
 * RESEARCH_AND_ESTIMATE §8.8.16 N-2運用・◎対象。
 */
function menuApplyPriority4SetCountAndPriceToMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('「' + MASTER_SHEET_NAME + '」シートが見つかりません。');
    return;
  }
  var asinSheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  if (!asinSheet) {
    SpreadsheetApp.getUi().alert('「' + ASIN_PASTE_SHEET_NAME + '」シートがありません。先にKeepa取得まで実行してください。');
    return;
  }
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(masterValues.length, 25); r++) {
    if ((masterValues[r] || []).indexOf(ANCHOR_HEADER_NAME) !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var headers = masterValues[headerRowIdx];
  var masterColMap = {};
  for (var h = 0; h < headers.length; h++) {
    masterColMap[String(headers[h]).trim()] = h;
  }
  var colMark = masterColMap[COL_MASTER_MARK_TARGET];
  if (colMark === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「' + COL_MASTER_MARK_TARGET + '」列がありません。◎が付いている行を対象にするため、この列を追加してください。');
    return;
  }
  var colCompetitiveSet = masterColMap[COL_MASTER_COMPETITIVE_SET_COUNTS];
  var colHoleSet = masterColMap[COL_MASTER_HOLE_SET_PROPOSAL];
  if (colCompetitiveSet === undefined || colHoleSet === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「' + COL_MASTER_COMPETITIVE_SET_COUNTS + '」または「' + COL_MASTER_HOLE_SET_PROPOSAL + '」列がありません。列を追加してから実行してください。');
    return;
  }
  var markRowIndices = [];
  for (var r = headerRowIdx + 1; r < masterValues.length; r++) {
    var cellVal = masterValues[r][colMark];
    if (cellVal === null || cellVal === undefined) continue;
    if (String(cellVal).indexOf('◎') !== -1) markRowIndices.push(r + 1);
  }
  if (markRowIndices.length === 0) {
    SpreadsheetApp.getUi().alert('マスタに◎が付いている行がありません。' + COL_MASTER_MARK_TARGET + '列に◎を付けてから実行してください。');
    return;
  }
  var blockResults = getAggregatedSetCountsByBlocksFromAsinPasteSheet(asinSheet);
  var written = 0;
  for (var i = 0; i < markRowIndices.length; i++) {
    var rowIdx = markRowIndices[i];
    if (i < blockResults.length && blockResults[i]) {
      masterSheet.getRange(rowIdx, colCompetitiveSet + 1).setValue(blockResults[i].competitive.join(', '));
      masterSheet.getRange(rowIdx, colHoleSet + 1).setValue(blockResults[i].hole.join(', '));
      written++;
    }
  }
  masterValues = masterSheet.getDataRange().getValues();
  var masterColMapNew = {};
  for (var h = 0; h < (masterValues[headerRowIdx] || []).length; h++) {
    masterColMapNew[String(masterValues[headerRowIdx][h]).trim()] = h;
  }
  var priceUpdated = proposePriceAndSetFromCompetitive(ss, masterSheet, headerRowIdx, markRowIndices, masterValues, masterColMapNew);
  SpreadsheetApp.getActive().toast(
    '◎行 ' + markRowIndices.length + ' 件中、セット数提案を ' + written + ' 件反映し、価格提案を ' + priceUpdated + ' 行に反映しました。',
    '優先4 完了',
    8
  );
}

/** 最終データ行の判定に使う列範囲（G～Z）。1-based のため 7～26。 */
const MASTER_LAST_DATA_COL_START = 7;
const MASTER_LAST_DATA_COL_END = 26;

/**
 * 【セット構成提案】AI情報取得data の記載がある2行目以降の全行を対象に、
 * ASIN貼り付け（Keepa用）のブロックから競合セット数・穴のセット数提案・ASIN・競合価格(Amazon)を取得し、
 * マスタの最終データ行（G～Zのいずれかに値がある最後の行）の次から、
 * 親行（メーカー名ベース・仕入判断・JAN・商品名ベース・卸値・競合店ASINコード・競合価格amazon）＋子行（セット数は数値のみ、競合価格はセット数別最安）を追記する。
 * 書き込み先: ▼商品マスタ(人間作業用)。トリガー: メニュー「セット構成提案」。
 */
function menuSetCompositionProposal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aiSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  var asinSheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);

  if (!aiSheet) {
    SpreadsheetApp.getUi().alert('「' + TARGET_SHEET_NAME + '」シートが見つかりません。');
    return;
  }
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('「' + MASTER_SHEET_NAME + '」シートが見つかりません。');
    return;
  }
  if (!asinSheet) {
    SpreadsheetApp.getUi().alert('「' + ASIN_PASTE_SHEET_NAME + '」シートがありません。先にKeepa取得まで実行してください。');
    return;
  }

  var aiData = aiSheet.getDataRange().getValues();
  if (aiData.length < 2) {
    SpreadsheetApp.getUi().alert('AI情報取得data に2行目以降のデータがありません。');
    return;
  }
  var aiHeaders = aiData[0];
  var aiColMap = getColumnIndexMap(aiHeaders);
  var aiCols = ['メーカー名', '仕入判断', 'JANコード', '商品名', '卸値(税抜)', '卸値(税込)', 'JANコードorオリジナルカタログ商品名', 'ASINコード'];
  var aiIndices = [];
  for (var a = 0; a < aiCols.length; a++) {
    var idx = aiColMap[aiCols[a]];
    if (idx === undefined) aiIndices.push(-1);
    else aiIndices.push(idx);
  }
  var aiDataRows = [];
  for (var r = 1; r < aiData.length; r++) {
    var row = aiData[r];
    var hasContent = false;
    for (var c = 0; c < aiIndices.length; c++) {
      if (aiIndices[c] >= 0) {
        var v = row[aiIndices[c]];
        if (v !== null && v !== undefined && String(v).trim() !== '') { hasContent = true; break; }
      }
    }
    if (hasContent) aiDataRows.push({ rowIndex: r, row: row });
  }
  if (aiDataRows.length === 0) {
    SpreadsheetApp.getUi().alert('AI情報取得data の2行目以降に、メーカー名・仕入判断・JAN・商品名・卸値・ASINのいずれかが入っている行がありません。');
    return;
  }

  var blockResults = getAggregatedSetCountsByBlocksFromAsinPasteSheet(asinSheet);
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(masterValues.length, 25); hr++) {
    if ((masterValues[hr] || []).indexOf(ANCHOR_HEADER_NAME) !== -1) {
      headerRowIdx = hr;
      break;
    }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterHeaders = masterValues[headerRowIdx];
  var numMasterCols = masterHeaders.length;
  var masterColMap = getColumnIndexMap(masterHeaders);
  var colMaker = masterColMap['メーカー名ベース'] !== undefined ? masterColMap['メーカー名ベース'] : masterColMap['メーカー名'];
  var colJudge = masterColMap['仕入判断'];
  var colJan = masterColMap['JANコード'];
  var colName = masterColMap['商品名ベース'] !== undefined ? masterColMap['商品名ベース'] : (masterColMap['商品名'] !== undefined ? masterColMap['商品名'] : masterColMap['オリジナルカタログ商品名']);
  var colCostEx = masterColMap['卸値(税抜)'];
  var colCostIn = masterColMap['卸値(税込)'];
  var colJanOrName = masterColMap['JANコードorオリジナルカタログ商品名'];
  var colCompetitorAsin = masterColMap['競合店ASINコード'];
  var colCompetitorUrl  = masterColMap[COL_COMPETITOR_URL_AMAZON]; // 任意列（なくてもエラーにしない）
  var colCompetitivePrice = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  var colSetQty = masterColMap[COL_MASTER_TOTAL_QTY];
  var colCheckbox = masterColMap[CHECKBOX_HEADER_NAME]; // 出品CK列（書き出し時にTRUEを設定）
  var colStrategy = masterColMap[COL_AMAZON_STRATEGY];  // セット数提案の根拠を子SKU先頭行に書き出す用
  if (colMaker === undefined || colJudge === undefined || colJan === undefined || colName === undefined ||
      colCostEx === undefined || colCostIn === undefined || colCompetitorAsin === undefined) {
    SpreadsheetApp.getUi().alert('マスタに次のいずれかの列がありません: メーカー名ベース(またはメーカー名), 仕入判断, JANコード, 商品名ベース(または商品名), 卸値(税抜), 卸値(税込), 競合店ASINコード');
    return;
  }
  if (colSetQty === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「' + COL_MASTER_TOTAL_QTY + '」列がありません。');
    return;
  }
  // 賞味期限: 「賞味期限（ある場合のみ）分からない場合は今日で入力」を優先し、AI情報取得dataから計算した日付を書き込む列
  var colExpiry = masterColMap['賞味期限（ある場合のみ）分からない場合は今日で入力'] !== undefined
    ? masterColMap['賞味期限（ある場合のみ）分からない場合は今日で入力']
    : (masterColMap['賞味期限(食品)'] !== undefined ? masterColMap['賞味期限(食品)'] : masterColMap['賞味期限']);
  var colOriginalCatalog = masterColMap['オリジナルカタログ商品名'];
  var colPurchaseDate = masterColMap['購入日'];

  var gStart = MASTER_LAST_DATA_COL_START - 1;
  var gEnd = MASTER_LAST_DATA_COL_END - 1;
  if (gEnd >= numMasterCols) gEnd = numMasterCols - 1;
  var lastDataRowIdx = headerRowIdx;
  for (var rr = headerRowIdx + 1; rr < masterValues.length; rr++) {
    var mRow = masterValues[rr];
    var hasData = false;
    for (var cc = gStart; cc <= gEnd; cc++) {
      var val = mRow[cc];
      if (val !== null && val !== undefined && String(val).trim() !== '') { hasData = true; break; }
    }
    if (hasData) lastDataRowIdx = rr;
  }
  var writeStartRow1Based = lastDataRowIdx + 2;

  function getAiVal(row, colKey) {
    var idx = aiColMap[colKey];
    if (idx === undefined) return '';
    var v = row[idx];
    return (v === null || v === undefined) ? '' : String(v).trim();
  }
  /** 賞味期限文言（例: 製造日より24ヶ月）を入力日から起算した日付に変換。解釈できない場合は入力日を返す。 */
  function computeExpiryDateFromInput(expiryStr, inputDate) {
    if (!expiryStr || !inputDate) return inputDate;
    var m = String(expiryStr).match(/(\d+)\s*ヶ?月/);
    if (m) {
      var months = parseInt(m[1], 10);
      if (!isNaN(months)) {
        var d = new Date(inputDate.getTime());
        d.setMonth(d.getMonth() + months);
        return d;
      }
    }
    return inputDate;
  }

  // AI情報取得data の商品名。列が「商品名」または「商品名ベース」のどちらでも取得できるようにする。
  function getAiProductName(aiRow) {
    return getAiVal(aiRow, '商品名') || getAiVal(aiRow, '商品名ベース');
  }
  // AI情報取得data 参照時は、▼商品マスタの商品名ベースとAIの商品名（または商品名ベース）が一致した行を用いる。
  function getMasterRowForAiRow(aiRow) {
    var nameBase = getAiVal(aiRow, '商品名ベース');
    var name = getAiVal(aiRow, '商品名');
    var key = nameBase || name;
    if (!key) return null;
    for (var r = headerRowIdx + 1; r < masterValues.length; r++) {
      var mVal = String((masterValues[r][colName] || '')).trim();
      if (mVal === key || mVal === nameBase || mVal === name) return masterValues[r];
    }
    return null;
  }

  function mergeSetCounts(competitive, hole) {
    var seen = {};
    var out = [];
    if (competitive && competitive.length) {
      for (var k = 0; k < competitive.length; k++) {
        var n = competitive[k];
        if (!seen[n]) { seen[n] = true; out.push(n); }
      }
    }
    if (hole && hole.length) {
      for (var h = 0; h < hole.length; h++) {
        var m = hole[h];
        if (!seen[m]) { seen[m] = true; out.push(m); }
      }
    }
    out.sort(function(a, b) { return a - b; });
    return out;
  }

  var rowsToWrite = [];
  var currentRow = writeStartRow1Based;
  for (var i = 0; i < aiDataRows.length; i++) {
    var aiRow = aiDataRows[i].row;
    var parentVals = [
      getAiVal(aiRow, 'メーカー名'),
      getAiVal(aiRow, '仕入判断'),
      getAiVal(aiRow, 'JANコード'),
      getAiProductName(aiRow),
      getAiVal(aiRow, '卸値(税抜)'),
      getAiVal(aiRow, '卸値(税込)'),
      getAiVal(aiRow, 'JANコードorオリジナルカタログ商品名'),
      ''
    ];
    // セット数別の最低価格とそのASINを取得（◎行のみ）
    var priceAndAsinBySet = getMinPriceAndAsinBySetCountFromAsinPasteBlock(asinSheet, i);
    Logger.log('[セット構成提案] index=' + i + ' 商品名=' + getAiProductName(aiRow));
    Logger.log('[セット構成提案] priceAndAsinBySet=' + JSON.stringify(priceAndAsinBySet));
    // 親行用: 全セット中の最低価格とそのASIN・URL
    var parentCompetitivePrice = '';
    var parentAsin = '';
    var parentUrl  = '';
    var minPriceOverall = Infinity;
    var setKeys = Object.keys(priceAndAsinBySet);
    for (var sk = 0; sk < setKeys.length; sk++) {
      var entry = priceAndAsinBySet[setKeys[sk]];
      if (entry.price < minPriceOverall) {
        minPriceOverall = entry.price;
        parentCompetitivePrice = entry.price;
        parentAsin = entry.asin;
        parentUrl  = entry.url || '';
      }
    }
    var block = i < blockResults.length ? blockResults[i] : null;
    var productName4Gap = getAiProductName(aiRow);
    var packetsPerBag = extractPacketsPerBag(productName4Gap);
    var sixMonthMaxSets = Math.floor(360 / packetsPerBag); // 1日1包想定での12ヶ月分セット数上限（異常値カット用）
    var setCounts = [];
    var geminiResponseText = ''; // セット数提案の根拠（子SKU先頭行のamazon価格戦略に書き出す）

    // セット数ポートフォリオ提案: Gemini を優先。失敗時は従来ロジック（競合＋業者用＋比例配分）にフォールバック。
    if (getGeminiApiKey()) {
      var masterRow = getMasterRowForAiRow(aiRow);
      function fromAiOrMaster(aiCol, masterCol) {
        var v = getAiVal(aiRow, aiCol);
        if (v) return v;
        if (!masterRow) return '';
        var idx = masterColMap[masterCol];
        if (idx === undefined && masterCol) idx = masterColMap['▼マスタ(' + masterCol + ')'];
        if (idx === undefined) return '';
        var x = masterRow[idx];
        return (x !== null && x !== undefined && x !== '') ? String(x).trim() : '';
      }
      var category = fromAiOrMaster('楽天ジャンル名', '楽天ジャンル名') || fromAiOrMaster('Yahooカテゴリ名', 'Yahooカテゴリ名') || '';
      var specParts = [];
      ['梱包:幅(cm)', '梱包:奥(cm)', '梱包:高(cm)', '梱包:重量(g)', '賞味期限(食品)', '保存方法(食品)', '★配送サイズ(タリフ)'].forEach(function(col) {
        var val = fromAiOrMaster(col, col);
        if (val) specParts.push(col + ':' + val);
      });
      var specs = specParts.join(' / ');
      var competitiveList = block ? block.competitive.filter(function(n) { return n <= sixMonthMaxSets; }) : [];
      var prompt = buildSetCountProposalPrompt({
        productName: productName4Gap,
        category: category,
        specs: specs,
        shippingWallText: '', // セット構成提案時は送料の壁は空（精度見直し時に拡張）
        competitiveSetList: competitiveList,
        targetAttribute: '単身者/家族'
      });
      var geminiResult = getSetCountProposalByGemini(prompt);
      if (geminiResult && geminiResult.responseText) geminiResponseText = geminiResult.responseText;
      // Geminiの回答を優先。6ヶ月上限を超える異常値のみカットし、それ以外はそのまま採用する
      if (geminiResult && geminiResult.setCounts && geminiResult.setCounts.length >= 3) {
        var filtered = geminiResult.setCounts.filter(function(n) { return n <= sixMonthMaxSets; });
        if (filtered.length >= 3) {
          setCounts = filtered;
          Logger.log('[セット構成提案] Gemini採用 setCounts=' + JSON.stringify(setCounts));
        }
      }
    }

    if (setCounts.length === 0) {
      // フォールバック: 競合セット数＋業者用＋比例配分
      setCounts = block ? block.competitive.filter(function(n) { return n <= sixMonthMaxSets; }) : [];
      if (block && block.maxCompetitive > 0) {
        var compSetMax = setCounts.length > 0 ? Math.max.apply(null, setCounts) : 1;
        var sixMonthCap = sixMonthMaxSets;
        var bizSet = getBusinessSetCount(productName4Gap, compSetMax);
        if (bizSet > sixMonthCap) { bizSet = sixMonthCap; }
        var currentMax = setCounts.length > 0 ? setCounts[setCounts.length - 1] : 0;
        if (bizSet <= currentMax) bizSet = Math.min(currentMax * 2, sixMonthCap);
        var regularSlots = Math.max(0, 10 - setCounts.length - 1);
        if (regularSlots > 0 && bizSet > currentMax) {
          var anchors = setCounts.concat([bizSet]);
          var holeNums = fillGapsProportionally(anchors, regularSlots);
          var seenComp = {};
          setCounts.forEach(function(n) { seenComp[n] = true; });
          holeNums.forEach(function(n) {
            if (!seenComp[n] && n !== bizSet) { seenComp[n] = true; setCounts.push(n); }
          });
        }
        var seenFinal = {};
        setCounts.forEach(function(n) { seenFinal[n] = true; });
        if (!seenFinal[bizSet]) setCounts.push(bizSet);
      }
      setCounts.sort(function(a, b) { return a - b; });
      if (setCounts.length > 10) setCounts = setCounts.slice(0, 10);
    }
    if (setCounts.length === 0) setCounts.push(1);

    // 競合店が出品しているセット数はすべて必ず含める。Gemini採用時はGeminiをベースに競合の欠けを追加。フォールバック時は現setCountsをベースに競合を追加。最大15種類。6ヶ月上限超は異常値として除外済み。
    var compWithinCap = block ? block.competitive.filter(function(n) { return n <= sixMonthMaxSets; }) : [];
    if (compWithinCap.length > 0) {
      var seenBase = {};
      setCounts.forEach(function(n) { seenBase[n] = true; });
      for (var ci = 0; ci < compWithinCap.length; ci++) {
        var cn = compWithinCap[ci];
        if (!seenBase[cn]) { setCounts.push(cn); seenBase[cn] = true; }
      }
      setCounts.sort(function(a, b) { return a - b; });
      if (setCounts.length > 15) setCounts = setCounts.slice(0, 15);
    }

    Logger.log('[セット構成提案] packetsPerBag=' + packetsPerBag + ' setCounts=' + JSON.stringify(setCounts));

    // 書き込み行データを列単位で収集する構造体に積む
    // { colIdx(0-based): [ row0値, row1値, ... ] } の形で管理し、最後に列単位 setValues する
    var numNewRows = 1 + setCounts.length; // 親行1 + 子行数
    function pushColValues(colMap, colIdx, rowValues) {
      if (colIdx === undefined) return;
      if (!colMap[colIdx]) colMap[colIdx] = [];
      for (var rv = 0; rv < rowValues.length; rv++) colMap[colIdx].push([rowValues[rv]]);
    }

    // 書き込み対象列インデックス→値配列のマップ
    var writeColMap = {};
    // 親行 + 子行ぶんの値を各列に積む
    var makerVals   = [getAiVal(aiRow, 'メーカー名')];
    var judgeVals   = [getAiVal(aiRow, '仕入判断')];
    var janVals     = [getAiVal(aiRow, 'JANコード')];
    var nameVals    = [getAiProductName(aiRow)];
    var costExVals  = [getAiVal(aiRow, '卸値(税抜)')];
    var costInVals  = [getAiVal(aiRow, '卸値(税込)')];
    var janOrNameVals = [getAiVal(aiRow, 'JANコードorオリジナルカタログ商品名')];
    var asinVals    = [parentAsin];
    var urlVals     = [parentUrl];
    var priceVals   = [''];  // 親行の競合価格は空
    var setQtyVals  = [''];  // 親行のセット数は空

    for (var s = 0; s < setCounts.length; s++) {
      var setNum = setCounts[s];
      var childEntry = priceAndAsinBySet[setNum];
      makerVals.push(getAiVal(aiRow, 'メーカー名'));
      judgeVals.push(getAiVal(aiRow, '仕入判断'));
      janVals.push(getAiVal(aiRow, 'JANコード'));
      nameVals.push(getAiProductName(aiRow));
      costExVals.push(getAiVal(aiRow, '卸値(税抜)'));
      costInVals.push(getAiVal(aiRow, '卸値(税込)'));
      janOrNameVals.push(getAiVal(aiRow, 'JANコードorオリジナルカタログ商品名'));
      asinVals.push(childEntry ? childEntry.asin  : '');
      urlVals.push(childEntry  ? (childEntry.url  || '') : '');
      priceVals.push(childEntry ? childEntry.price : '');
      setQtyVals.push(setNum);
    }

    pushColValues(writeColMap, colMaker,          makerVals);
    pushColValues(writeColMap, colJudge,          judgeVals);
    pushColValues(writeColMap, colJan,            janVals);
    pushColValues(writeColMap, colName,           nameVals);
    pushColValues(writeColMap, colCostEx,         costExVals);
    pushColValues(writeColMap, colCostIn,         costInVals);
    pushColValues(writeColMap, colJanOrName,      janOrNameVals);
    pushColValues(writeColMap, colCompetitorAsin, asinVals);
    pushColValues(writeColMap, colCompetitorUrl,  urlVals);
    pushColValues(writeColMap, colCompetitivePrice, priceVals);
    pushColValues(writeColMap, colSetQty,         setQtyVals);

    // 子SKUの先頭行にセット数提案の根拠（Gemini返答全文）を amazon価格戦略 に書き出す
    if (colStrategy !== undefined) {
      var strategyVals = [''];
      for (var sv = 0; sv < setCounts.length; sv++) strategyVals.push(sv === 0 ? geminiResponseText : '');
      pushColValues(writeColMap, colStrategy, strategyVals);
    }

    rowsToWrite.push({ numRows: numNewRows, colMap: writeColMap });
  }

  if (rowsToWrite.length === 0) {
    SpreadsheetApp.getUi().alert('書き込む行がありません。');
    return;
  }

  // 必要行数を確保
  var totalNewRows = rowsToWrite.reduce(function(sum, b) { return sum + b.numRows; }, 0);
  var startRow = writeStartRow1Based;
  if (startRow + totalNewRows - 1 > masterSheet.getMaxRows()) {
    masterSheet.insertRowsAfter(masterSheet.getMaxRows(), (startRow + totalNewRows - 1) - masterSheet.getMaxRows());
  }

  // 書き込み: 列単位で setValues（書き込み対象列以外は一切触らない）
  // Q列(17)～HJ列(218) の範囲はテンプレート行（1行目=親SKU用、2行目=子SKU用）から数式をコピーする
  // amazon価格戦略列・賞味期限列は除外する（戦略列はセット構成提案の根拠、賞味期限はAI情報から計算した日付をStep4で書き込むため）
  var FORMULA_TMPL_COL_START = 17;  // Q列 (1-based)
  var FORMULA_TMPL_COL_END   = Math.min(218, numMasterCols); // HJ列 or シート最大列
  var formulaTmplColCount    = FORMULA_TMPL_COL_END - FORMULA_TMPL_COL_START + 1;
  var doCopyFormulas         = formulaTmplColCount > 0;
  var strategyCol1Based = (colStrategy !== undefined) ? (colStrategy + 1) : -1;
  var expiryCol1Based = (colExpiry !== undefined) ? (colExpiry + 1) : -1;
  var copyRanges = []; // { start: 1-based, numCols: number }（除外列を飛ばした区間の配列）
  if (doCopyFormulas) {
    var excludedCols = [];
    if (strategyCol1Based >= FORMULA_TMPL_COL_START && strategyCol1Based <= FORMULA_TMPL_COL_END) excludedCols.push(strategyCol1Based);
    if (expiryCol1Based >= FORMULA_TMPL_COL_START && expiryCol1Based <= FORMULA_TMPL_COL_END && excludedCols.indexOf(expiryCol1Based) < 0) excludedCols.push(expiryCol1Based);
    excludedCols.sort(function(a, b) { return a - b; });
    var prev = FORMULA_TMPL_COL_START - 1;
    for (var ex = 0; ex < excludedCols.length; ex++) {
      var c = excludedCols[ex];
      if (c > prev + 1) copyRanges.push({ start: prev + 1, numCols: c - prev - 1 });
      prev = c;
    }
    if (FORMULA_TMPL_COL_END > prev) copyRanges.push({ start: prev + 1, numCols: FORMULA_TMPL_COL_END - prev });
  }
  var row1OriginalCatalogValue = (colOriginalCatalog !== undefined) ? masterSheet.getRange(1, colOriginalCatalog + 1).getValue() : null;
  try {
    var rowCursor = startRow;
    for (var bi = 0; bi < rowsToWrite.length; bi++) {
      var block2 = rowsToWrite[bi];
      var nRows = block2.numRows;
      var cMap  = block2.colMap;

      // Step1: テンプレート行からQ-HJ列の数式をコピー（amazon価格戦略列は除外。相対参照を行方向に自動調整）
      // 書き込み対象列は Step2 でデータを上書きするため先にコピーして問題ない
      if (copyRanges.length > 0) {
        for (var cr = 0; cr < copyRanges.length; cr++) {
          var rng = copyRanges[cr];
          // 親行（1行目テンプレート → rowCursor行目へ）
          masterSheet.getRange(1, rng.start, 1, rng.numCols)
            .copyTo(
              masterSheet.getRange(rowCursor, rng.start, 1, rng.numCols),
              SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
              false
            );
          // 子行（2行目テンプレート → rowCursor+1行目以降へ）
          if (nRows > 1) {
            masterSheet.getRange(2, rng.start, 1, rng.numCols)
              .copyTo(
                masterSheet.getRange(rowCursor + 1, rng.start, nRows - 1, rng.numCols),
                SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
                false
              );
          }
        }
      }

      // Step2: 書き込み対象列のデータを上書き（数式より優先）
      var colKeys = Object.keys(cMap);
      for (var ck = 0; ck < colKeys.length; ck++) {
        var colIdx = parseInt(colKeys[ck], 10);
        var vals2d = cMap[colIdx]; // [[v0],[v1],...]
        masterSheet.getRange(rowCursor, colIdx + 1, nRows, 1).setValues(vals2d);
      }
      // Step3: 出品CK列にTRUEを書き込む（セット構成提案で書き出した行を後工程のフィルタ対象とする）
      if (colCheckbox !== undefined) {
        var checkVals = [];
        for (var cv = 0; cv < nRows; cv++) checkVals.push([true]);
        masterSheet.getRange(rowCursor, colCheckbox + 1, nRows, 1).setValues(checkVals);
      }
      // Step4: 賞味期限・オリジナルカタログ商品名・購入日（入力列以外への1・2行目コピー条件）
      var aiRow = aiDataRows[bi].row;
      var inputDate = new Date();
      if (colExpiry !== undefined) {
        var cat = getAiVal(aiRow, '楽天ジャンル名') || getAiVal(aiRow, 'Yahooカテゴリ名') || getAiVal(aiRow, 'カテゴリー') || '';
        var isFood = /食品|飲料|お酒|フード|食料|飲料・お酒/.test(cat);
        if (isFood) {
          var expiryStr = getAiVal(aiRow, '賞味期限(食品)');
          var expiryDate = expiryStr ? computeExpiryDateFromInput(expiryStr, inputDate) : inputDate;
          for (var er = 0; er < nRows; er++) {
            masterSheet.getRange(rowCursor + er, colExpiry + 1).setValue(expiryDate);
          }
        }
      }
      if (colOriginalCatalog !== undefined) {
        masterSheet.getRange(rowCursor, colOriginalCatalog + 1).setValue(row1OriginalCatalogValue);
        for (var oc = 1; oc < nRows; oc++) {
          masterSheet.getRange(rowCursor + oc, colOriginalCatalog + 1).setFormula('=M$1');
        }
      }
      if (colPurchaseDate !== undefined) {
        for (var pd = 0; pd < nRows; pd++) {
          masterSheet.getRange(rowCursor + pd, colPurchaseDate + 1).setValue(inputDate);
        }
      }
      // 親SKU行を行全体で水色にして視覚的に判別しやすくする
      masterSheet.getRange(rowCursor, 1, 1, numMasterCols).setBackground('#ADD8E6');
      rowCursor += nRows;
    }
  } catch (e) {
    var logSheetName = 'セット構成提案_エラーログ';
    var ssLog = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ssLog.getSheetByName(logSheetName);
    if (!logSheet) logSheet = ssLog.insertSheet(logSheetName);
    if (logSheet.getLastRow() === 0) {
      logSheet.getRange(1, 1, 1, 8).setValues([['日時', 'エラーメッセージ', 'startRow', 'totalNewRows', 'numMasterCols', 'lastDataRowIdx', 'aiDataRows数', 'スタック']]);
      logSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }
    var errMsg = (e && e.message) ? String(e.message) : String(e);
    var stack = (e && e.stack) ? String(e.stack).substring(0, 2000) : '';
    logSheet.appendRow([
      new Date().toLocaleString('ja-JP'),
      errMsg,
      startRow,
      totalNewRows,
      numMasterCols,
      lastDataRowIdx,
      aiDataRows.length,
      stack
    ]);
    Logger.log('[セット構成提案] エラー: ' + errMsg + ' startRow=' + startRow + ' totalNewRows=' + totalNewRows);
    Logger.log(stack);
    SpreadsheetApp.getUi().alert('セット構成提案の書き込みでエラーが発生しました。\n\n' + errMsg + '\n\n詳細はシート「' + logSheetName + '」および GAS の「表示→ログ」で確認できます。');
    return;
  }
  SpreadsheetApp.getActive().toast(
    'セット構成提案を ' + totalNewRows + ' 行（親行・子行）マスタに追記しました。開始行: ' + startRow,
    'セット構成提案 完了',
    8
  );
  // 価格提案に必要な情報のため、セット構成提案の最後に Amazonカテゴリー自動入力を実行
  menuFillAmazonCategoryByGemini();
}

/**
 * ASIN貼り付けシートのセット数修正後に、マスタの競合価格・ASIN・URLだけを再反映する。
 * 新規行は追加しない。JANコード＋A.セット商品数が一致する行の3列のみ上書き。
 */
function menuUpdateCompetitivePriceOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aiSheet     = ss.getSheetByName(TARGET_SHEET_NAME);
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  var asinSheet   = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);

  if (!aiSheet)     { SpreadsheetApp.getUi().alert('「' + TARGET_SHEET_NAME + '」シートが見つかりません。'); return; }
  if (!masterSheet) { SpreadsheetApp.getUi().alert('「' + MASTER_SHEET_NAME + '」シートが見つかりません。'); return; }
  if (!asinSheet)   { SpreadsheetApp.getUi().alert('「' + ASIN_PASTE_SHEET_NAME + '」シートが見つかりません。'); return; }

  // AI情報取得data: JANコード → ブロック番号マップ（行順 = ブロック番号）
  var aiData = aiSheet.getDataRange().getValues();
  var aiColMap = getColumnIndexMap(aiData[0] || []);
  var aiJanIdx = aiColMap['JANコード'];
  var janToBlock = {};
  if (aiJanIdx !== undefined) {
    var blockNo = 0;
    for (var ar = 1; ar < aiData.length; ar++) {
      var jan = String(aiData[ar][aiJanIdx] || '').trim();
      if (jan !== '') {
        if (janToBlock[jan] === undefined) janToBlock[jan] = blockNo;
        blockNo++;
      }
    }
  }

  // マスタのヘッダー行検出
  var masterValues = masterSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var hr = 0; hr < Math.min(masterValues.length, 25); hr++) {
    if ((masterValues[hr] || []).indexOf(ANCHOR_HEADER_NAME) !== -1) { headerRowIdx = hr; break; }
  }
  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。');
    return;
  }
  var masterHeaders = masterValues[headerRowIdx];
  var masterColMap  = getColumnIndexMap(masterHeaders);
  var colJan              = masterColMap['JANコード'];
  var colSetQty           = masterColMap[COL_MASTER_TOTAL_QTY];
  var colCompetitivePrice = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  var colCompetitorAsin   = masterColMap['競合店ASINコード'];
  var colCompetitorUrl    = masterColMap[COL_COMPETITOR_URL_AMAZON];

  if (colJan === undefined || colSetQty === undefined) {
    SpreadsheetApp.getUi().alert('マスタに「JANコード」または「' + COL_MASTER_TOTAL_QTY + '」列がありません。');
    return;
  }
  if (colCompetitivePrice === undefined && colCompetitorAsin === undefined && colCompetitorUrl === undefined) {
    SpreadsheetApp.getUi().alert('マスタに更新対象の列（競合価格amazon / 競合店ASINコード / 競合AmazonページURL）が一つも見つかりません。');
    return;
  }

  // ブロックデータをキャッシュ（同一ブロックを複数回読まないように）
  var blockCache = {};
  function getCachedBlock(blockIndex) {
    if (blockCache[blockIndex] === undefined) {
      blockCache[blockIndex] = getMinPriceAndAsinBySetCountFromAsinPasteBlock(asinSheet, blockIndex);
    }
    return blockCache[blockIndex];
  }

  var updatedCount = 0;
  var skippedCount = 0;
  for (var r = headerRowIdx + 1; r < masterValues.length; r++) {
    var mRow = masterValues[r];
    var jan    = String(mRow[colJan]    || '').trim();
    var setVal = String(mRow[colSetQty] || '').trim();
    if (jan === '' || setVal === '') continue; // 子行のみ対象（セット数あり）
    var setNum = parseInt(setVal, 10);
    if (isNaN(setNum) || setNum < 1) continue;

    var blockIndex = janToBlock[jan];
    if (blockIndex === undefined) { skippedCount++; continue; }

    var priceAndAsin = getCachedBlock(blockIndex);
    var entry = priceAndAsin[setNum];

    // 更新対象の値（データなし = 空文字でクリア）
    var newPrice = entry ? entry.price : '';
    var newAsin  = entry ? entry.asin  : '';
    var newUrl   = entry ? (entry.url || '') : '';

    // 変更があったセルだけ書き込む（1セルずつ setValues で一括化する前に収集）
    if (colCompetitivePrice !== undefined) masterSheet.getRange(r + 1, colCompetitivePrice + 1).setValue(newPrice);
    if (colCompetitorAsin   !== undefined) masterSheet.getRange(r + 1, colCompetitorAsin   + 1).setValue(newAsin);
    if (colCompetitorUrl    !== undefined) masterSheet.getRange(r + 1, colCompetitorUrl    + 1).setValue(newUrl);
    updatedCount++;
  }

  SpreadsheetApp.getActive().toast(
    '競合価格・ASIN・URLを更新しました。更新行数: ' + updatedCount +
    (skippedCount > 0 ? '  スキップ（JANマッチなし）: ' + skippedCount + '行' : ''),
    '競合価格のみ修正を反映 完了',
    8
  );
}

function menuKeepaFetchAsinPasteSheet20() { runKeepaFetchAsinPasteSheet(20); }
function menuKeepaFetchAsinPasteSheet50() { runKeepaFetchAsinPasteSheet(50); }
function menuKeepaFetchAsinPasteSheetAll() { runKeepaFetchAsinPasteSheet(0); }

/**
 * 【原因追究】ReferenceError: document is not defined の発生箇所を特定する診断を実行する。
 * 各ステップを try-catch で実行し、どこでエラーになるかを Logger とシート「documentエラー診断ログ」に記録する。
 * メニュー: AI出品ツール → 商品リサーチ → ②出品用 → documentエラー原因を診断（ログで調査）
 */
function menuDebugDocumentError() {
  debugDocumentErrorCause();
}

/**
 * document is not defined 原因追究の本体。ステップごとに実行し、エラーが出たステップとスタックを記録する。
 */
function debugDocumentErrorCause() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(DOCUMENT_ERROR_DIAG_SHEET_NAME);
  if (!logSheet) logSheet = ss.insertSheet(DOCUMENT_ERROR_DIAG_SHEET_NAME);
  var headers = ['日時', 'ステップ', 'ステップ名', '結果', 'エラーメッセージ', 'スタック（抜粋）'];
  if (logSheet.getLastRow() === 0) {
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  var now = new Date().toLocaleString('ja-JP');
  var rows = [];

  function runStep(stepNum, stepName, fn) {
    var result = 'OK';
    var errMsg = '';
    var stack = '';
    try {
      fn();
    } catch (e) {
      result = 'エラー';
      errMsg = (e && e.message) ? String(e.message) : String(e);
      stack = (e && e.stack) ? String(e.stack) : '';
      Logger.log('[document診断] Step' + stepNum + ' ' + stepName + ' でエラー: ' + errMsg);
      Logger.log(stack);
    }
    rows.push([now, stepNum, stepName, result, errMsg, stack.substring(0, 2000)]);
    return result === 'OK';
  }

  // Step 0: グローバルに document が存在するか（参照せず typeof のみ → エラーにはならない）
  runStep(0, 'typeof document 確認', function() {
    var t = typeof document;
    Logger.log('[document診断] Step0: typeof document = ' + t);
    if (t !== 'undefined') Logger.log('[document診断] 警告: GASサーバーでは通常 undefined です');
  });

  // Step 1: getGeminiApiKey
  runStep(1, 'getGeminiApiKey()', function() {
    var key = getGeminiApiKey();
    Logger.log('[document診断] Step1: getGeminiApiKey 長さ=' + (key ? key.length : 0));
  });

  // Step 2: getRefImageUrlForAsinPasteBlock
  runStep(2, 'getRefImageUrlForAsinPasteBlock(ss,0)', function() {
    var url = getRefImageUrlForAsinPasteBlock(ss, 0);
    Logger.log('[document診断] Step2: refUrl 長さ=' + (url ? url.length : 0));
  });

  // Step 3: getImageMatchScoreByGemini（テスト用URLで実行。両方ある場合のみ）
  runStep(3, 'getImageMatchScoreByGemini(テストURL2つ)', function() {
    var refUrl = getRefImageUrlForAsinPasteBlock(ss, 0);
    var testKeepaUrl = 'https://m.media-amazon.com/images/I/71abc123.jpg';
    if (!refUrl || refUrl.indexOf('http') !== 0) {
      Logger.log('[document診断] Step3: 参考画像URLがないためスキップ（エラーではない）');
      return;
    }
    var score = getImageMatchScoreByGemini(refUrl, testKeepaUrl);
    Logger.log('[document診断] Step3: getImageMatchScoreByGemini 結果=' + (score ? score.overall : 'null'));
  });

  // Step 4: inferSetCountFromReferenceImageByGemini
  runStep(4, 'inferSetCountFromReferenceImageByGemini(テストURL)', function() {
    var refUrl = getRefImageUrlForAsinPasteBlock(ss, 0);
    if (!refUrl || refUrl.indexOf('http') !== 0) {
      Logger.log('[document診断] Step4: 参考画像URLがないためスキップ');
      return;
    }
    var res = inferSetCountFromReferenceImageByGemini(refUrl);
    Logger.log('[document診断] Step4: inferSetCount 結果=' + (res && res.setCount != null ? res.setCount : 'null'));
  });

  // Step 5: 本番経路（1件だけ）— ここで落ちればスタックで場所が分かる
  runStep(5, 'runKeepaFetchAsinPasteSheetImpl(1)', function() {
    runKeepaFetchAsinPasteSheetImpl(1);
  });

  // シートに追記（1行ずつ appendRow で範囲・行数不一致を防ぐ）
  var i;
  for (i = 0; i < rows.length; i++) {
    logSheet.appendRow(rows[i]);
  }
  logSheet.activate();
  var errCount = 0;
  for (var i = 0; i < rows.length; i++) { if (rows[i][3] === 'エラー') errCount++; }
  SpreadsheetApp.getActive().toast(
    '診断完了: ' + rows.length + ' ステップ中 ' + errCount + ' 件でエラー。シート「' + DOCUMENT_ERROR_DIAG_SHEET_NAME + '」と実行ログを確認してください。',
    'documentエラー診断',
    8
  );
}

/**
 * 【② 出品用】ASIN貼り付け（Keepa用）シートのデータ（3行目以降）をすべてクリアする。1行目（商品名の式）と2行目（ヘッダー）は残す。
 */
function menuResetAsinPasteSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「' + ASIN_PASTE_SHEET_NAME + '」シートがありません。');
    return;
  }
  var lastRow = sheet.getLastRow();
  var data = sheet.getDataRange().getValues();
  var clearFromRow = 3;
  if (data.length >= 1 && data[0] && String(data[0][0]).trim() === 'ASIN' && (!data[1] || String(data[1][0]).trim() !== 'ASIN')) {
    clearFromRow = 2;
  }
  if (lastRow < clearFromRow) {
    SpreadsheetApp.getActive().toast('クリアするデータがありません。', 'リセット', 3);
    return;
  }
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) lastCol = 1;
  sheet.getRange(clearFromRow, 1, lastRow, lastCol).clearContent();
  SpreadsheetApp.getActive().toast((clearFromRow === 2 ? '2' : '3') + '行目以降をクリアしました。', 'リセット完了', 5);
}

// ----------------------------------------
// 3.2 商品リサーチ（②出品用）: 競合価格・卸値から販売価格・セット数を提案してマスタに反映
// ----------------------------------------
/**
 * 【② 出品用】選択行について、マスタの競合価格（Amazon/楽天/Yahoo!）と卸値(税込)を元に
 * 販売価格amazon・楽天価格設定・Yahoo!価格設定への提案値を算出し、マスタに書き込む。
 * 販売価格・競合価格・卸値はすべて税込で判断する。
 * セット数（A.セット商品数）は、提案値が渡された場合のみ書き込む（省略時は変更しない）。
 * RESEARCH_AND_ESTIMATE §1.1・§1.4 のAI提案フローに相当。
 */
function menuProposePriceAndSetToSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) { SpreadsheetApp.getUi().alert('▼商品マスタ(人間作業用) シートが見つかりません。'); return; }
  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() !== MASTER_SHEET_NAME) { SpreadsheetApp.getUi().alert('マスタシートを開き、提案を反映したい行を選択してから実行してください。'); return; }
  const range = activeSheet.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('行を選択してから実行してください。'); return; }
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const rowIndices = []; for (let r = 0; r < numRows; r++) rowIndices.push(startRow + r);
  const masterValues = masterSheet.getDataRange().getValues();
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(masterValues.length, 20); r++) { if (masterValues[r].includes(ANCHOR_HEADER_NAME)) { headerRowIdx = r; break; } }
  if (headerRowIdx === -1) { SpreadsheetApp.getUi().alert('マスタのヘッダー行（ASINコード）が見つかりません。'); return; }
  const masterColMap = getColumnIndexMap(masterValues[headerRowIdx]);
  const required = [COL_COMPETITIVE_PRICE_AMAZON, COL_COMPETITIVE_PRICE_RAKUTEN, COL_COMPETITIVE_PRICE_YAHOO, COL_PRICE_AMAZON, COL_PRICE_RAKUTEN, COL_PRICE_YAHOO];
  const missing = getMissingMasterWriteColumnsResearch(masterColMap, required);
  if (missing.length) { SpreadsheetApp.getUi().alert('マスタに次の列がありません: ' + missing.join('、')); return; }
  const updated = proposePriceAndSetFromCompetitive(ss, masterSheet, headerRowIdx, rowIndices, masterValues, masterColMap);
  SpreadsheetApp.getActive().toast(updated + '行に価格提案を反映しました', '完了');
}

/**
 * 指定したマスタ行について、競合価格・卸値(税込)から提案価格を算出してマスタに書き込む。価格はすべて税込で統一。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} masterSheet
 * @param {number} headerRowIdx 1-based ではない行インデックス（0-based）
 * @param {number[]} rowIndices1Based 対象行の 1-based 行番号の配列
 * @param {Array[]} masterValues マスタの getDataRange().getValues()
 * @param {Object} masterColMap ヘッダー名 → 0-based 列インデックス
 * @param {{ totalQty?: number|string }} [proposedSetCount] セット数提案（ある場合のみ A.セット商品数 に書き込む）
 * @return {number} 更新した行数
 */
function proposePriceAndSetFromCompetitive(ss, masterSheet, headerRowIdx, rowIndices1Based, masterValues, masterColMap, proposedSetCount) {
  const colAmazon = masterColMap[COL_COMPETITIVE_PRICE_AMAZON];
  const colRakuten = masterColMap[COL_COMPETITIVE_PRICE_RAKUTEN];
  const colYahoo = masterColMap[COL_COMPETITIVE_PRICE_YAHOO];
  const colCost = masterColMap[COL_COST_TAX_IN];
  const colPriceAmazon = masterColMap[COL_PRICE_AMAZON];
  const colPriceRakuten = masterColMap[COL_PRICE_RAKUTEN];
  const colPriceYahoo = masterColMap[COL_PRICE_YAHOO];
  const colTotalQty = masterColMap[COL_MASTER_TOTAL_QTY];
  if (colPriceAmazon === undefined || colPriceRakuten === undefined || colPriceYahoo === undefined) return 0;
  const parseNum = function(v) { if (v === undefined || v === null || v === '') return NaN; const n = Number(String(v).replace(/,/g, '')); return isNaN(n) ? NaN : n; };
  const roundPrice = function(n) { return Math.round(n); };
  const minMarginRate = 1.05; // 卸値の5%以上を確保
  let updated = 0;
  for (let i = 0; i < rowIndices1Based.length; i++) {
    const rowIdx = rowIndices1Based[i];
    const dataRowIdx = rowIdx - 1;
    if (dataRowIdx <= headerRowIdx || dataRowIdx >= masterValues.length) continue;
    const row = masterValues[dataRowIdx];
    const compAmazon = parseNum(row[colAmazon]);
    const compRakuten = parseNum(row[colRakuten]);
    const compYahoo = parseNum(row[colYahoo]);
    const cost = parseNum(row[colCost]);
    const prices = [compAmazon, compRakuten, compYahoo].filter(function(p) { return !isNaN(p); });
    const median = function(arr) {
      if (arr.length === 0) return NaN;
      const a = arr.slice().sort(function(x, y) { return x - y; });
      const m = Math.floor(a.length / 2);
      return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
    };
    const fallback = prices.length ? median(prices) : (isNaN(cost) ? NaN : roundPrice(cost * minMarginRate));
    const ensureMin = function(p) { if (isNaN(p)) return fallback; if (!isNaN(cost) && cost > 0 && p < cost * minMarginRate) return roundPrice(cost * minMarginRate); return roundPrice(p); };
    let priceAmazon = ensureMin(compAmazon);
    let priceRakuten = ensureMin(compRakuten);
    let priceYahoo = ensureMin(compYahoo);
    if (isNaN(priceAmazon)) priceAmazon = isNaN(priceRakuten) ? priceYahoo : (isNaN(priceYahoo) ? priceRakuten : median([priceRakuten, priceYahoo]));
    if (isNaN(priceRakuten)) priceRakuten = isNaN(priceAmazon) ? priceYahoo : (isNaN(priceYahoo) ? priceAmazon : median([priceAmazon, priceYahoo]));
    if (isNaN(priceYahoo)) priceYahoo = isNaN(priceAmazon) ? priceRakuten : (isNaN(priceRakuten) ? priceAmazon : median([priceAmazon, priceRakuten]));
    if (isNaN(priceAmazon) && isNaN(priceRakuten) && isNaN(priceYahoo)) continue; // 競合価格も卸値も無い行はスキップ
    masterSheet.getRange(rowIdx, colPriceAmazon + 1).setValue(isNaN(priceAmazon) ? '' : roundPrice(priceAmazon));
    masterSheet.getRange(rowIdx, colPriceRakuten + 1).setValue(isNaN(priceRakuten) ? '' : roundPrice(priceRakuten));
    masterSheet.getRange(rowIdx, colPriceYahoo + 1).setValue(isNaN(priceYahoo) ? '' : roundPrice(priceYahoo));
    updated++;
    if (proposedSetCount != null && proposedSetCount.totalQty !== undefined && proposedSetCount.totalQty !== '' && colTotalQty !== undefined) {
      const qty = proposedSetCount.totalQty;
      masterSheet.getRange(rowIdx, colTotalQty + 1).setValue(typeof qty === 'number' ? qty : String(qty).trim());
    }
  }
  return updated;
}

// ==========================================
// 5. 楽天市場CSV出力機能 (Ver 49.1 - バラし展開・リッチ化・バグ修正版)
// ==========================================
/**
 * @param {boolean} [skipConfirm] - true のとき確認ダイアログを出さずに開始（トリガー実行用）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ssOverride] - トリガー実行時は openById で開いたスプレッドシートを渡す（nullだとgetActiveSpreadsheetを使用）
 */
function generateRakutenCSV(skipConfirm, ssOverride) {
  const ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    console.error('[楽天CSV] スプレッドシートを取得できません（トリガー実行時は ssOverride を渡してください）');
    throw new Error('スプレッドシートを取得できません');
  }
  if (ssOverride) console.log('[楽天CSV] トリガー経由で実行 ssId=' + ss.getId());
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
  // トリガー実行時は getUi() が使えないため、UI は使わずエラー時は throw する
  const ui = ssOverride ? null : SpreadsheetApp.getUi();

  if (!masterSheet || !settingSheet) {
    if (ui) { ui.alert('シートが見つかりません。'); return; }
    throw new Error('シートが見つかりません。');
  }
  // 反映確認トリガーからログを書くために、スプレッドシートIDを必ず保存する
  try {
    PropertiesService.getScriptProperties().setProperty('RAKUTEN_REFLECTION_LOG_SS_ID', ss.getId());
  } catch (e) { console.warn('RAKUTEN_REFLECTION_LOG_SS_ID 保存失敗: ' + e.message); }

  if (!skipConfirm && ui) {
    const response = ui.alert('確認', '出品商品にレ点は付いていますか？\n(OKで処理を開始します)', ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) return;
  }

  // 設定読み込み
  const settingRaw = settingSheet.getDataRange().getValues();
  const rawHeaders = settingRaw[0].slice(1);
  let validCols = 0;
  for (let i = 0; i < rawHeaders.length; i++) {
    if (rawHeaders[i] && String(rawHeaders[i]).trim() !== "") validCols++;
    else break;
  }

  const csvHeaders = settingRaw[0].slice(1, 1 + validCols);
  const systemKeys = settingRaw[1].slice(1, 1 + validCols);
  const masterCols = settingRaw[2].slice(1, 1 + validCols);
  const configs    = settingRaw.length > 3 ? settingRaw[3].slice(1, 1 + validCols) : [];
  
  const prodFlags   = settingRaw.length > 4 ? settingRaw[4].slice(1, 1 + validCols) : [];
  const optionFlags = settingRaw.length > 5 ? settingRaw[5].slice(1, 1 + validCols) : [];
  const skuFlags    = settingRaw.length > 6 ? settingRaw[6].slice(1, 1 + validCols) : [];

  let catalogIdColIdx = -1;
  let noCatalogIdReasonColIdx = -1;
  for(let c=0; c<systemKeys.length; c++){
    if(systemKeys[c] === 'catalog_id') catalogIdColIdx = c;
    if(systemKeys[c] === 'catalog_id_reason') noCatalogIdReasonColIdx = c;
  }

  // マスタ読み込み
  const masterValues = masterSheet.getDataRange().getValues();
  let masterHeaderRowIdx = -1;
  for (let r = 0; r < 20; r++) {
    if (masterValues[r].includes(ANCHOR_HEADER_NAME)) {
      masterHeaderRowIdx = r;
      break;
    }
  }
  if (masterHeaderRowIdx === -1) {
    if (ui) { ui.alert('マスタシートのヘッダーが見つかりません。'); return; }
    throw new Error('マスタシートのヘッダーが見つかりません。');
  }

  const masterHeaders = masterValues[masterHeaderRowIdx];
  const masterColMap = {};
  masterHeaders.forEach((name, i) => masterColMap[String(name).trim()] = i);

  const parentSkuIdx = masterColMap['親SKU'];
  const childSkuIdx = masterColMap['子SKU'];
  const checkIdx = masterColMap[CHECKBOX_HEADER_NAME];
  const varThemeIdx = masterColMap['バリエーションテーマ'];
  const varValueIdx = masterColMap['バリエーション値'];

  if (parentSkuIdx === undefined || checkIdx === undefined) {
    if (ui) { ui.alert(`マスタに「親SKU」または「${CHECKBOX_HEADER_NAME}」列がありません。`); return; }
    throw new Error(`マスタに「親SKU」または「${CHECKBOX_HEADER_NAME}」列がありません。`);
  }

  // 親行かどうか（子SKUが空なら親行）
  const isParentRow = (row) => {
    if (childSkuIdx === undefined) return false;
    const v = row[childSkuIdx];
    return v === undefined || v === null || String(v).trim() === '';
  };

  // --- CSV生成 (グルーピング) ---
  const groups = {};
  const checkedRows = [];

  for (let i = masterHeaderRowIdx + 1; i < masterValues.length; i++) {
    const row = masterValues[i];
    const isChecked = row[checkIdx] === true || String(row[checkIdx]).toUpperCase() === 'TRUE';
    if (!isChecked) continue; 

    checkedRows.push(row);
    const pSku = String(row[parentSkuIdx]).trim();
    if (!groups[pSku]) groups[pSku] = [];
    groups[pSku].push(row);
  }

  if (checkedRows.length === 0) {
    if (ui) { ui.alert('チェックが付いている商品がありません。'); return; }
    throw new Error('チェックが付いている商品がありません。');
  }

  const csvOutput = [];
  csvOutput.push(csvHeaders);

  const INHERIT_KEYS = [
    'tax_rate', 'warehouse_flg', 'delivery_set_id', 
    'lead_time_instock', 'lead_time_outstock', 
    'search_visiblity', 'consumption_tax',
    'double_price_id', 'order_limit', 'restock_button', 'noshi_flg',
    'restore_inventory_flg', 'backorder_flg', 'delivery_lead_time',
    'sku_warehouse_flg', 'shipping_cost', 'regional_shipping_id', 
    'single_shipping_flg', 'okiphai_flg', 'drop_shipping',
    'delivery_id_instock', 'shipping_class1', 'shipping_class2', 'individual_shipping'
  ];

  // --- HTML生成用ヘルパー ---
  const getMVal = (row, colName) => {
    let idx = masterColMap[`▼マスタ(${colName})`];
    if (idx === undefined) idx = masterColMap[colName];
    return (idx !== undefined) ? row[idx] : "";
  };
  
  const rakutenShopId = (settingRaw.length > 10 && settingRaw[10][1]) ? String(settingRaw[10][1]).trim() : "";

  for (const pSku of Object.keys(groups)) {
    const groupRows = groups[pSku];
    let parentRowOrigin;
    let children;
    if (isParentRow(groupRows[0])) {
      // 先頭が親行 → 従来どおり
      parentRowOrigin = groupRows[0];
      children = groupRows.slice(1);
    } else {
      // 子SKUのレ点のみ: マスタから親行を取得して紐付ける
      let foundParent = null;
      for (let r = masterHeaderRowIdx + 1; r < masterValues.length; r++) {
        const row = masterValues[r];
        if (String(row[parentSkuIdx]).trim() !== pSku) continue;
        if (isParentRow(row)) {
          foundParent = row;
          break;
        }
      }
      parentRowOrigin = foundParent || groupRows[0]; // 親が見つからなければ先頭を暫定使用
      children = groupRows; // レ点の付いた子行すべて
    }

    let varDefinitions = "";
    let varChoices = "";
    const uniqueValues = new Set();
    const orderedValues = [];
    let themeName = "";
    
    // データ吸い上げ
    const commonValues = {}; 
    // バリエーション選択肢は全行（idx=0含む）から収集。子のみレ点のとき先頭行も子なので欠落させない
    // 1セルに「A|B|C」と入っている場合は | で分割して各選択肢として扱う（楽天は各選択肢32バイトまで）
    groupRows.forEach((row) => {
      if (varThemeIdx !== undefined && varValueIdx !== undefined) {
        const th = String(row[varThemeIdx]).trim();
        if (th && !themeName) themeName = th.split(/\r\n|\n/)[0].replace(/[|｜]/g, '');
        const val = String(row[varValueIdx]).trim();
        if (val) {
          var valParts = val.split(/[|｜]/).map(function(s) { return s.trim(); }).filter(Boolean);
          if (valParts.length === 0) valParts = [val];
          valParts.forEach(function(part) {
            if (!uniqueValues.has(part)) {
              uniqueValues.add(part);
              orderedValues.push(part);
            }
          });
        }
      }
      for (let c = 0; c < systemKeys.length; c++) {
        const key = systemKeys[c];
        const mapColName = masterCols[c];
        if (key.includes('product_attr') || key.startsWith('attr_') || INHERIT_KEYS.includes(key)) {
          if (!commonValues[key] && mapColName && masterColMap[mapColName] !== undefined) {
             const val = row[masterColMap[mapColName]];
             if (val !== "" && val !== undefined) commonValues[key] = val;
          }
        }
      }
    });

    if (orderedValues.length > 0) {
      // 楽天: バリエーション項目キー定義と項目名定義は | 区切りで同数必須。1軸のときは名前1個必ず。
      varDefinitions = (themeName && String(themeName).trim()) ? themeName : (orderedValues[0] || "選択肢");
      varChoices = orderedValues.join('|');
    }
    
    let targets = [];
    
    if (children.length > 0) {
      // マルチSKU: 親行ベースで展開
      children.forEach(childRow => {
        const virtualRow = [...parentRowOrigin]; 
        


       // (A) URLと商品名
        const cSkuVal = (masterColMap['子SKU'] !== undefined) ? childRow[masterColMap['子SKU']] : "";
        const varVal = (masterColMap['バリエーション値'] !== undefined) ? childRow[masterColMap['バリエーション値']] : "";
        
        if (cSkuVal) {
           // ★修正: 子SKUの末尾(最後のハイフン以降)のみを結合して短縮する
           let suffix = "_" + cSkuVal; // デフォルト(ハイフンがない場合)は全結合
           const lastHyphen = cSkuVal.lastIndexOf('-');
           if (lastHyphen !== -1) {
             // ハイフンを含む末尾を取得 (例: sankyo-1234-S444 -> -S444)
             suffix = cSkuVal.substring(lastHyphen); 
           }
           const newUrl = `${pSku}${suffix}`;
           
           virtualRow[parentSkuIdx] = newUrl; // URL上書き
        }
        if (varVal) {
           const orgName = parentRowOrigin[masterColMap['商品名']];
           virtualRow[masterColMap['商品名']] = `${orgName}【${varVal}】`;
        }




        // (B) メイン画像 (子行に画像があれば、親のメイン画像をそれで差し替える)
        const mainImgIdx = masterColMap['楽天メイン画像1'];
        if (mainImgIdx !== undefined) {
           const childImg = childRow[mainImgIdx];
           if (childImg && String(childImg).trim() !== "") {
              virtualRow[mainImgIdx] = childImg; // 画像上書き
           }
        }

        // ★追加: 他のバリエーションの画像を「サブ画像」の空き枠に追加登録する
        // 1. まず、既存のサブ画像（共通画像）がどこまで埋まっているか確認
        let subCounter = 1;
        while (subCounter <= 10) {
           const subColName = `楽天サブ画像${subCounter}`;
           const subIdx = masterColMap[subColName];
           // 値が入っていれば次へ、空欄ならそこが開始位置
           if (subIdx !== undefined && virtualRow[subIdx] && String(virtualRow[subIdx]).trim() !== "") {
              subCounter++;
           } else {
              break; 
           }
        }

        // 2. 他の子SKU（兄弟）の画像を順番に埋め込む
        children.forEach(sibling => {
           if (sibling === childRow) return; // 自分自身はスキップ
           if (subCounter > 10) return;      // 枠がいっぱい(10枚)なら終了

           const sibImgVal = (mainImgIdx !== undefined) ? sibling[mainImgIdx] : "";
           if (sibImgVal && String(sibImgVal).trim() !== "") {
              const targetSubCol = masterColMap[`楽天サブ画像${subCounter}`];
              if (targetSubCol !== undefined) {
                 virtualRow[targetSubCol] = sibImgVal;
                 subCounter++;
              }
           }
        });

        targets.push({
          isSplit: true,
          row: virtualRow,
          originalChildRow: childRow
        });
      });
    } else {
      // シングルSKU
      targets.push({
        isSplit: false,
        row: groupRows[0],
        originalChildRow: groupRows[0]
      });
    }

    // CSV行生成
    targets.forEach(target => {
      const currentRow = target.row;
      const parentDataValues = {};

      // (A) HTML生成 (リッチ化対応)
      const lpTitle = currentRow[masterColMap['商品名']] || "";
      const lpCatch = getMVal(currentRow, '楽天キャッチ(60-80字)') || ""; 
      
      let lpImagesHtml = "";
      // ★修正: 説明文(HTML)の先頭にメイン画像を含める（元に戻す）
      const lpTargetCols = ['楽天メイン画像1'];
      for(let i=1; i<=10; i++) lpTargetCols.push(`楽天サブ画像${i}`);

      lpTargetCols.forEach(colName => {
         let path = getMVal(currentRow, colName);
         if (path && String(path).trim() !== "") {
           if (!path.startsWith('/')) path = '/' + path;
           let fullSrc = rakutenShopId ? `https://image.rakuten.co.jp/${rakutenShopId}/cabinet${path}` : `/cabinet${path}`;
           lpImagesHtml += `<div class="img-box"><img src="${fullSrc}" alt="商品画像" /></div>`;
         }
      });

      const specName = getMVal(currentRow, 'ブランド') || getMVal(currentRow, 'シリーズ') || "商品";
      const specIngredients = getMVal(currentRow, '原材料(食品)');
      const specExpiry = getMVal(currentRow, '賞味期限(食品)');
      const specStorage = getMVal(currentRow, '保存方法(食品)');

      const smartphoneTableHtml = `<br><br><table border="0" cellspacing="1" cellpadding="7" width="100%" bgcolor="#ccc"><tr><th colspan="2" bgcolor="#eee">商品情報</th></tr><tr><th width="30%" bgcolor="#eee">名称</th><td bgcolor="#fff">${specName}</td></tr><tr><th width="30%" bgcolor="#eee">原材料名</th><td bgcolor="#fff">${specIngredients}</td></tr><tr><th width="30%" bgcolor="#eee">賞味期限</th><td bgcolor="#fff">${specExpiry}</td></tr><tr><th width="30%" bgcolor="#eee">保存方法</th><td bgcolor="#fff">${specStorage}</td></tr></table>`;
      const pcSpecTableHtml = `<style type='text/css'> <!-- table.spec { font-size: 13px; color: #333; line-height: 1.4em; min-width: 350px; border: 1px solid #c1c1c1; border-collapse: collapse; margin: 0 0 20px 0; } table.spec th { border: 1px solid #c1c1c1; background: #eee; padding: 8px; font-weight: bold; text-align: center; background: -moz-linear-gradient(top, rgba(255,255,255,1) 0%, rgba(248,248,248,1) 100%); background: -webkit-gradient(linear, left top, left bottom, color-stop(0%,rgba(255,255,255,1)), color-stop(100%,rgba(248,248,248,1))); filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#ffffff', endColorstr='#f8f8f8',GradientType=0 ); } table.spec td { width: 70%; padding: 10px; border: 1px solid #c1c1c1; text-align: left; background: #fff; } --> </style><table class="table-spec spec"><thead><tr><th colspan="2">商品情報</th></tr></thead><tbody><tr><th>名称</th><td>${specName}</td></tr><tr><th>原材料名</th><td>${specIngredients}</td></tr><tr><th>賞味期限</th><td>${specExpiry}</td></tr><tr><th>保存方法</th><td>${specStorage}</td></tr></tbody></table>`;
      const pcLpHtml = `<style type="text/css"> <!-- #item-cont-re { font-family:"メイリオ", Meiryo, "ヒラギノ角ゴPro W3", "Hiragino Kaku Gothic Pro",Osaka, "ＭＳ Ｐゴシック", sans-serif; font-size:100.01%; width: 615px; overflow:hidden; margin: 0 auto; text-align: center; background: #fff; } #item-cont-re .img-box { width: 615px; padding: 0; margin: 0 auto 16px; text-align: center; } #item-cont-re .img-box img { max-width: 100%; min-width: 500px; } #item-cont-re .img-box a img:hover { opacity: 0.8; } #item-cont-re .title-box { width: 500px; margin: 0 auto; text-align: center; } #item-cont-re .title-box h2 { font-size: 24px; color: #333; border-bottom: 1px solid #ddd; margin: 0; padding: 0 0 8px; } #item-cont-re .title-box p { font-size: 16px; color: #666; margin: 16px auto; padding: 0; } #item-cont-re .item-caption { width: 500px; margin: 0 auto; } #item-cont-re .item-caption p { font-size: 14px; color: #555; line-height: 1.8em; text-align: left; display: inline-block; } --> </style><div id="item-cont-re">${lpImagesHtml}<div class="title-box"> <h2>${lpTitle}</h2> <p>${lpCatch}<br></p> </div><div class="item-caption"> <p></p> </div></div><div class="template_id" style="display:none;">5</div>`;

      // (B) 商品行
      const productRow = createRow(currentRow, prodFlags, true, false, false);
      for(let c=0; c<systemKeys.length; c++) parentDataValues[systemKeys[c]] = productRow.rawValue[c];
      csvOutput.push(productRow.csvLine);

      // (C) オプション行
      const targetOptHeader = 'オプション名(特典)';
      const targetOptIdx = masterColMap[targetOptHeader];
      const hasOptionData = (targetOptIdx !== undefined && currentRow[targetOptIdx] && String(currentRow[targetOptIdx]).trim() !== "");
      
      if (hasOptionData && optionFlags.some(f => f == 1 || f === true || String(f).toUpperCase() === 'TRUE')) {
        const optionRow = createRow(currentRow, optionFlags, false, true, false);
        csvOutput.push(optionRow.csvLine);
      }

      // (D) SKU行（同一親に紐づく他子SKU。子のレ点のみのときも children で一貫）
      const allSiblings = children;
      
      if (allSiblings.length > 0) {
        allSiblings.forEach(siblingRow => {
          const skuRow = createRow(siblingRow, skuFlags, false, false, true);
          // SKU行の親URL書き換え
          const urlIdx = systemKeys.indexOf('item_url');
          if (urlIdx > -1) {
             const newParentUrl = productRow.rawValue[urlIdx];
             skuRow.csvLine[urlIdx] = newParentUrl;
          }
          csvOutput.push(skuRow.csvLine);
        });
      } else {
        const skuRow = createRow(currentRow, skuFlags, false, false, true);
        csvOutput.push(skuRow.csvLine);
      }

      // --- 内部関数 ---
      function createRow(row, flags, isProduct, isOption, isSku) {
        const line = [];
        const rawValues = []; 

        for (let c = 0; c < csvHeaders.length; c++) {
          const key = systemKeys[c];
          const mapColName = masterCols[c];
          const configVal = configs[c];
          const flag = flags[c];
          const isTarget = (flag === true || String(flag).toUpperCase() === 'TRUE' || flag == 1);
          const hasConfig = (configVal !== "" && configVal !== undefined && configVal !== null);

          let value = "";
          
          if (mapColName && masterColMap[mapColName] !== undefined) value = row[masterColMap[mapColName]];
          
          if ((isProduct || isOption) && (value === "" || value === undefined)) {
             if (commonValues[key]) value = commonValues[key];
          }

          if (isSku) {
            const isAttribute = (key.includes('product_attr') || key.startsWith('attr_'));
            if (INHERIT_KEYS.includes(key) || isAttribute) {
              if (parentDataValues[key] !== undefined) value = parentDataValues[key];
            }
          }

          // ★修正: 宣言順序を直しました
          const isCalculation = hasConfig && (String(configVal).includes('[[') || (!isNaN(parseFloat(configVal)) && mapColName));
          if (hasConfig && !isCalculation && !mapColName) value = configVal;
          
          if (hasConfig) {
            const sConf = String(configVal);
            if (sConf.includes('[[TAX_AUTO]]')) {
              if (value === 'A_GEN_REDUCED') value = 0.08;
              else if (value === 'A_GEN_STANDARD') value = 0.1;
            } else if (sConf.startsWith('[[固定値:')) {
              value = sConf.replace('[[固定値:', '').replace(']]', '');
            
            
            
            } else if (!isNaN(parseFloat(sConf)) && mapColName) {
              const numValue = parseFloat(String(value).replace(/,/g, ''));
              if (!isNaN(numValue)) value = Math.floor(numValue * parseFloat(sConf));
            }
          }

          // ★修正: 画像パスの補完 (/cabinet は付けないのが正解)
          if ((key.includes('image_path') || key === 'sku_image_path' || key === 'white_bg_path') && value && String(value).trim() !== "") {
             let path = String(value).trim();
             
             // httpで始まらない（R-Cabinet内画像）の場合
             if (!path.startsWith('http')) {
                // もし "/cabinet/" が付いていたら削除する (重複防止)
                if (path.startsWith('/cabinet/')) {
                   path = path.replace('/cabinet', '');
                }
                // 先頭に "/" がなければ付ける
                if (!path.startsWith('/')) {
                   path = '/' + path;
                }
                value = path; // 例: /12644827/img.jpg
             }
          }




          if (isProduct) {
            if (key === 'mobile_item_desc') value = value + smartphoneTableHtml;
            if (key === 'item_desc') value = value + pcSpecTableHtml;
            if (key === 'sales_desc') value = pcLpHtml; 
          }

          if (key === 'item_url' && !value && isSku) value = parentDataValues['item_url'];
          
          if ((isProduct || isOption)) {
            if (key === 'variation_names') value = varDefinitions;
            if (key === 'variation_choices_1') value = varChoices;
            // 楽天: キーと項目名は同数必須。1軸のときは Key0 と名前1個をセットで出力
            if (key === 'variation_keys' && (varDefinitions || varChoices)) value = "Key0";
          }
          if (isSku && uniqueValues.size > 0) {
              if (key === 'sku_var_key_1') value = "Key0";
              if (key === 'sku_var_value_1' && varValueIdx !== undefined) value = row[varValueIdx];
          }

          if (isProduct && key === 'catalog_id') value = "";
          if (['jan','catalog_id','isbn'].some(k => key.includes(k)) && value) value = convertScientificToDecimal(value);
          if (key === 'catalog_id_reason') {
               const janIdx = masterColMap['JANコード'];
               const janVal = (janIdx !== undefined) ? row[janIdx] : "";
               if (janVal && String(janVal).trim() !== "") value = ""; 
          }

          rawValues.push(value); 
          if (!isTarget && key !== 'item_url') value = "";
          line.push(value);
        }

        // 自動クリーニング
        for (let c = 0; c < systemKeys.length; c++) {
          const key = systemKeys[c];
          const val = line[c];
          if (key.startsWith('image_type_') && val) {
            const suffix = key.replace('image_type_', '');
            if (!line[systemKeys.indexOf('image_path_' + suffix)]) line[c] = '';
          }
          if (key === 'sku_image_type' && val && !line[systemKeys.indexOf('sku_image_path')]) line[c] = '';
          if (key === 'white_bg_type' && val && !line[systemKeys.indexOf('white_bg_path')]) line[c] = '';
          if (key === 'choice_type' && val && !line[systemKeys.indexOf('option_name')]) line[c] = '';
          if (key.includes('attr_value') && !val) {
            const unitIdx = systemKeys.indexOf(key.replace('value', 'unit'));
            if (unitIdx > -1) line[unitIdx] = '';
            const nameIdx = systemKeys.indexOf(key.replace('value', 'name'));
            if (nameIdx > -1) line[nameIdx] = '';
          }
        }
        
        if (isSku && catalogIdColIdx > -1 && noCatalogIdReasonColIdx > -1) {
          const id = line[catalogIdColIdx];
          const isReasonTarget = (skuFlags[noCatalogIdReasonColIdx] == 1 || skuFlags[noCatalogIdReasonColIdx] === true);
          if (isReasonTarget) {
            if (id && String(id).trim() !== "") line[noCatalogIdReasonColIdx] = ""; 
          } else {
            line[noCatalogIdReasonColIdx] = "";
          }
        }

        return { csvLine: line, rawValue: rawValues };
      }
    });
  }

  const settingDataSliced = [csvHeaders, systemKeys, masterCols, configs];
  writeToVerificationSheetWithDiagnosis(csvOutput, settingDataSliced);

  // 楽天「バリエーション1選択肢定義」は「各選択肢」（|で区切った1つ1つ）が32バイトまで。SKU行の「バリエーション項目選択肢1」も同様
  const varChoices1Idx = systemKeys.indexOf('variation_choices_1');
  const skuVarValue1Idx = systemKeys.indexOf('sku_var_value_1');
  if (varChoices1Idx >= 0) {
    const overRows = [];
    const addedRows = {}; // 同一行の重複表示を防ぐ
    for (let r = 1; r < csvOutput.length; r++) {
      const line = csvOutput[r];
      const val = line[varChoices1Idx];
      if (val && String(val).trim()) {
        const parts = String(val).split('|');
        for (let i = 0; i < parts.length; i++) {
          const seg = parts[i].trim();
          if (seg && truncateToByteLength(seg, 32) !== seg) {
            const itemUrl = line[systemKeys.indexOf('item_url')];
            const key = (itemUrl || '行' + (r + 1));
            if (!addedRows[key]) { addedRows[key] = true; overRows.push(key + '（選択肢: ' + (seg.length > 20 ? seg.substring(0, 20) + '…' : seg) + '）'); }
            break;
          }
        }
      }
      if (skuVarValue1Idx >= 0) {
        const skuVal = line[skuVarValue1Idx];
        if (skuVal && String(skuVal).trim() && truncateToByteLength(String(skuVal).trim(), 32) !== String(skuVal).trim()) {
          const itemUrl = line[systemKeys.indexOf('item_url')];
          const key = (itemUrl || '行' + (r + 1));
          if (!addedRows[key]) { addedRows[key] = true; overRows.push(key + '（SKU選択肢が32バイト超過）'); }
        }
      }
    }
    if (overRows.length > 0) {
      const allItems = overRows.join(', ');
      const msg = '「バリエーション1選択肢定義」のいずれかの選択肢（|で区切った1つ1つ）、またはSKU行のバリエーション値が32バイトを超えています。\n\n【該当】\n' + allItems + '\n\nマスタで Ctrl+F などで上記のいずれかを検索し、該当の選択肢を32バイト以内に修正してください。';
      if (ui) {
        ui.alert('楽天アップロードを中止しました', msg, ui.ButtonSet.OK);
        return;
      }
      throw new Error(msg);
    }
  }

  // DriveにCSVを保存し、ファイルIDを取得
  const fileId = saveCsvWithArchive(csvOutput, OUTPUT_FILE_NAME);
  
  // MAKE経由で楽天FTPへ自動アップロード（Google Drive経由）
  if (fileId) {
    try { PropertiesService.getScriptProperties().setProperty('RAKUTEN_REFLECTION_LOG_SS_ID', ss.getId()); } catch (_) {}
    logRakutenReflection('CSV送信', `Drive保存済み、ファイルID: ${fileId}`);
    try {
      sendFileToRakutenFtpViaDrive(fileId, OUTPUT_FILE_NAME);
    } catch (e) {
      logRakutenReflection('CSV送信', `MAKE送信エラー: ${e.message}`);
      throw e;
    }
    // 反映確認を開始（5分おきに出品URLへアクセスして200なら反映済み。完了時にメール送信）
    const productsForCheck = extractProductsFromCsvOutput(csvOutput, systemKeys);
    if (productsForCheck.length > 0) {
      startReflectionCheck(productsForCheck, rakutenShopId);
    } else {
      console.warn("   ⚠️ 反映確認: 対象商品が0件のためスキップ。完了メールは送信されません。");
      logRakutenReflection('反映確認スキップ', '対象商品0件');
    }
  } else {
    logRakutenReflection('CSV送信スキップ', 'Drive保存失敗のためFTP送信なし');
    console.error("❌ CSVファイルの保存に失敗したため、FTPアップロードをスキップします");
  }
}

/**
 * csvOutput から商品一覧（code, name）を抽出（反映確認・完了メール用）
 * 設定シートのシステムキーが item_name または product_name のどちらでも対応
 */
function extractProductsFromCsvOutput(csvOutput, systemKeys) {
  const itemUrlIdx = systemKeys.indexOf('item_url');
  const itemNameIdx = systemKeys.indexOf('item_name') >= 0 ? systemKeys.indexOf('item_name') : systemKeys.indexOf('product_name');
  if (itemUrlIdx === -1 || itemNameIdx === -1) return [];
  const seen = {};
  const products = [];
  for (let i = 1; i < csvOutput.length; i++) {
    const line = csvOutput[i];
    const itemUrl = line[itemUrlIdx];
    const itemName = (line[itemNameIdx] !== undefined && line[itemNameIdx] !== null) ? String(line[itemNameIdx]).trim() : "";
    if (itemUrl && String(itemUrl).trim() && !seen[itemUrl]) {
      seen[itemUrl] = true;
      products.push({ code: String(itemUrl).trim(), name: itemName });
    }
  }
  return products;
}

/**
 * 楽天出品完了メール送信（反映確認完了後に送信。全件200なら「反映完了」、タイムアウトなら「M/N件反映」）
 * 要件: 改修前と同様にプレーンテキストのみ送信する。Gmail等がURLを自動でリンク化するため、商品ページ・削除用がクリック可能になる。
 * トリガー実行時は getActiveSpreadsheet() が null になるため、RAKUTEN_REFLECTION_LOG_SS_ID でスプレッドシートを開く。
 * @param {Object} checkData - 反映確認データ { shopId, items: [{ url, name, reflected }] }
 * @param {string} reason - "complete" | "timeout"
 */
function sendRakutenCompletionEmailWithReflectionResult(checkData, reason) {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (_) {}
  if (!ss) {
    try {
      var id = PropertiesService.getScriptProperties().getProperty('RAKUTEN_REFLECTION_LOG_SS_ID');
      if (id) ss = SpreadsheetApp.openById(id);
    } catch (_) {}
  }
  var sheetUrl = (ss && ss.getUrl) ? ss.getUrl() : "";

  const shopId = checkData.shopId || "";
  const items = checkData.items || [];
  const reflectedCount = items.filter(i => i.reflected).length;
  const totalCount = items.length;

  var deleteBaseUrl = getRakutenDeleteWebAppUrl();
  var body = "";
  if (reason === "complete") {
    body = "楽天への出品が完了しました（全" + totalCount + "件、商品ページが表示されることを確認しました）。\n\n";
  } else {
    body = "楽天への反映確認がタイムアウトしました。" + reflectedCount + "/" + totalCount + "件が反映済みです。\n\n";
  }
  body += "--- 出品商品一覧 ---\n\n";
  items.forEach(function(item, idx) {
    var status = item.reflected ? "✅反映済" : "⏳未確認";
    var truncName = (item.name || "").length > 30 ? (item.name || "").substring(0, 30) + "..." : (item.name || "");
    var storeUrl = shopId ? "https://item.rakuten.co.jp/" + shopId + "/" + item.url + "/" : "";
    var deleteUrl = deleteBaseUrl ? deleteBaseUrl + (deleteBaseUrl.indexOf("?") >= 0 ? "&" : "?") + "action=rakuten_delete&item_url=" + encodeURIComponent(item.url) + "&name=" + encodeURIComponent(item.name || "") : "";
    body += (idx + 1) + ". " + truncName + " [" + status + "]\n";
    body += "   商品管理番号: " + item.url + "\n";
    if (storeUrl) body += "   商品ページ: " + storeUrl + "\n";
    if (deleteUrl) body += "   削除用: " + deleteUrl + "\n";
    body += "\n";
  });
  body += "\n--- 削除が必要な場合 ---\n";
  if (deleteBaseUrl) {
    body += "【スマホ】各商品の「削除用」のURLをタップ→確認画面で「削除する」で削除できます。\n\n";
  } else {
    body += "【スマホ】削除用URLを利用するには、GASの「デプロイ」でWebアプリをデプロイし、Script Properties に「RAKUTEN_DELETE_WEBAPP_URL」でそのURLを設定してください。\n\n";
  }
  body += "【PC】スプレッドシートで削除したい行の「楽天削除CK」にレ点を付け、メニュー「AI出品ツール」→「楽天 商品を削除...」でまとめて削除できます。\n   " + sheetUrl + "\n\n";
  body += "削除後はマスタシートで修正し、再出品してください。\n";

  var recipient = "";
  try { recipient = PropertiesService.getScriptProperties().getProperty("NOTIFICATION_EMAIL") || ""; } catch (e) {}
  if (!recipient && ss) {
    try {
      var rakutenSetting = ss.getSheetByName(SETTING_SHEET_NAME);
      if (rakutenSetting) recipient = rakutenSetting.getRange("B17").getValue() || "";
    } catch (_) {}
  }
  if (!recipient && ss) {
    try {
      var yahooSetting = ss.getSheetByName("▼設定(Yahooマッピング)");
      if (yahooSetting) recipient = yahooSetting.getRange("B17").getValue() || "";
    } catch (_) {}
  }
  if (!recipient) {
    console.warn("   ⚠️ 楽天完了メール: 通知先が未設定です");
    return;
  }
  var subject = reason === "complete"
    ? "【楽天出品完了】全" + totalCount + "件が反映されました"
    : "【楽天反映確認】" + reflectedCount + "/" + totalCount + "件反映（タイムアウト）";
  try {
    MailApp.sendEmail(recipient, subject, body);
    console.log("   -> 楽天完了メールを送信しました: " + recipient);
    logRakutenReflection("完了メール送信", "件名: " + subject + "、宛先: " + recipient);
  } catch (e) {
    console.warn("   ⚠️ 楽天完了メール送信に失敗しました: " + e.message);
    logRakutenReflection("完了メール送信失敗", e.message);
  }
}

/**
 * 楽天削除用WebアプリのURL。完了メールの「削除用」リンク（1品ずつクリックして出品停止する用）に使う。
 * 優先: (1) ScriptApp.getService().getUrl() (2) Script Properties の RAKUTEN_DELETE_WEBAPP_URL (3) 定数。
 * 完了メールはトリガーから送られるため getUrl() が空になり、削除用がメールに出ない。メニュー「削除用URLを今のデプロイに設定」で保存すると解消する。
 */
function getRakutenDeleteWebAppUrl() {
  try {
    var url = ScriptApp.getService().getUrl();
    if (url && String(url).indexOf("http") === 0) return String(url);
  } catch (err) {}
  try {
    var url = (PropertiesService.getScriptProperties().getProperty("RAKUTEN_DELETE_WEBAPP_URL") || "").trim();
    if (url && url !== "RAKUTEN_DELETE_WEBAPP_URL" && url.indexOf("http") === 0) return url;
  } catch (e) {}
  return (RAKUTEN_DELETE_WEBAPP_BASE_URL && String(RAKUTEN_DELETE_WEBAPP_BASE_URL).indexOf("http") === 0) ? RAKUTEN_DELETE_WEBAPP_BASE_URL : "";
}

/**
 * メニュー「削除用URLを今のデプロイに設定」用。WebアプリURLを取得して保存。取得できない場合は手動入力に誘導する。
 */
function menuSetRakutenDeleteWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var url = "";
  try {
    url = ScriptApp.getService().getUrl();
  } catch (e) {}
  if (!url || String(url).indexOf("http") !== 0) {
    var msg = "このプロジェクトはまだ「ウェブアプリ」としてデプロイされていません。\n\n";
    msg += "【方法A】デプロイしてから再度このメニューを実行\n";
    msg += "1. スプレッドシートの「拡張機能」→「Apps Script」でエディタを開く\n";
    msg += "2. 右上「デプロイ」→「新しいデプロイ」\n";
    msg += "3. 種類で「ウェブアプリ」を選び、説明を入力して「デプロイ」\n";
    msg += "4. 「ウェブアプリのURL」をコピーできるので、あとで使う場合は控えておく\n";
    msg += "5. このスプレッドシートに戻り、もう一度「削除用URLを今のデプロイに設定」を実行\n\n";
    msg += "【方法B】すでにデプロイ済みのURLがある場合\n";
    msg += "「OK」を押したあと、メニュー「削除用URLを手動で設定」からURLを貼り付けて保存できます。";
    ui.alert("削除用URLを取得できませんでした", msg, ui.ButtonSet.OK);
    return;
  }
  try {
    PropertiesService.getScriptProperties().setProperty("RAKUTEN_DELETE_WEBAPP_URL", url);
    SpreadsheetApp.getActiveSpreadsheet().toast("削除用URLを保存しました。今後の楽天出品完了メールに「削除用」リンクが付きます。", "設定完了", 8);
    ui.alert("設定完了", "削除用URLを保存しました。\n\n今後の楽天出品完了メールに、各商品ごとの「削除用」リンクが表示されます。メールからクリックすると削除確認画面が開き、1品ずつ出品を停止できます。", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("保存に失敗しました", (e.message || ""), ui.ButtonSet.OK);
  }
}

/**
 * メニュー「削除用URLを手動で設定」用。デプロイ済みのウェブアプリURLを貼り付けて保存する。
 * デプロイ後に「ウェブアプリのURL」（…/exec で終わるもの）をコピーしてここに貼り付ける。
 */
function menuSetRakutenDeleteWebAppUrlManual() {
  var ui = SpreadsheetApp.getUi();
  var current = "";
  try {
    current = PropertiesService.getScriptProperties().getProperty("RAKUTEN_DELETE_WEBAPP_URL") || "";
  } catch (_) {}
  var result = ui.prompt(
    "削除用URLを手動で設定",
    "GASの「デプロイ」→「ウェブアプリ」でデプロイしたときのURL（https://script.google.com/.../exec）を貼り付けてください。\n\n現在: " + (current ? current.substring(0, 50) + "…" : "未設定"),
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var url = (result.getResponseText() || "").trim();
  if (!url) {
    ui.alert("URLが入力されていません。", ui.ButtonSet.OK);
    return;
  }
  if (url.indexOf("http") !== 0) {
    ui.alert("URLは https:// で始まる形式で入力してください。", ui.ButtonSet.OK);
    return;
  }
  try {
    PropertiesService.getScriptProperties().setProperty("RAKUTEN_DELETE_WEBAPP_URL", url);
    SpreadsheetApp.getActiveSpreadsheet().toast("削除用URLを保存しました。", "設定完了", 6);
    ui.alert("設定完了", "削除用URLを保存しました。今後の楽天出品完了メールに「削除用」リンクが付きます。", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("保存に失敗しました", (e.message || ""), ui.ButtonSet.OK);
  }
}

// 楽天削除用CSVファイル名（楽天FTPは item-delete.csv で削除処理）
const RAKUTEN_DELETE_CSV_NAME = 'item-delete.csv';

/**
 * 楽天 商品を削除（スプレッドシートの「楽天削除CK」で選択 → 確認 → item-delete.csv 生成 → MAKE送信 → フラグ更新・CK解除）
 */
function showRakutenDeleteSelectionDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    ui.alert('❌ エラー', 'マスタシートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  const data = masterSheet.getDataRange().getValues();
  // ヘッダー行を「楽天削除CK」と「親SKU」が含まれる行で検出（8行目固定だとずれる場合があるため）
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(data.length, 15); r++) {
    const row = data[r].map(c => String(c).trim());
    if (row.includes('楽天削除CK') && row.includes('親SKU')) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx === -1) {
    ui.alert('❌ エラー', 'マスタに「楽天削除CK」と「親SKU」の両方がある行（ヘッダー行）が見つかりません。', ui.ButtonSet.OK);
    return;
  }
  const headers = data[headerRowIdx].map(h => String(h).trim());
  const colRakutenDeleteCK = headers.indexOf('楽天削除CK');
  const colRakutenDeleteFlag = headers.indexOf('楽天削除フラグ');
  const colRakutenDeleteDate = headers.indexOf('楽天削除日時');
  const colRakutenItemUrl = headers.indexOf('楽天商品管理番号');
  const colParentSku = headers.indexOf('親SKU');
  const colChildSku = headers.indexOf('子SKU');

  if (colRakutenDeleteCK === -1) {
    ui.alert('❌ 「楽天削除CK」列がありません', 'マスタに「楽天削除CK」列を追加し、削除したい行にレ点を付けてから再度実行してください。', ui.ButtonSet.OK);
    return;
  }
  if (colParentSku === -1) {
    ui.alert('❌ エラー', '「親SKU」列が見つかりません。', ui.ButtonSet.OK);
    return;
  }

  const isChecked = (val) => val === true || val === 1 || String(val).toUpperCase() === 'TRUE';
  const getItemUrl = (row) => {
    if (colRakutenItemUrl !== -1 && row[colRakutenItemUrl] && String(row[colRakutenItemUrl]).trim()) {
      return String(row[colRakutenItemUrl]).trim();
    }
    const pSku = String(row[colParentSku]).trim();
    if (!pSku) return "";
    const cSku = (colChildSku !== -1 && row[colChildSku] != null) ? String(row[colChildSku]).trim() : "";
    if (!cSku) return pSku;
    const lastHyphen = cSku.lastIndexOf('-');
    const suffix = lastHyphen !== -1 ? cSku.substring(lastHyphen) : "_" + cSku;
    return pSku + suffix;
  };

  const itemUrls = [];
  const rowIndices = [];
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!isChecked(row[colRakutenDeleteCK])) continue;
    const itemUrl = getItemUrl(row);
    if (!itemUrl) continue;
    itemUrls.push(itemUrl);
    rowIndices.push(i);
  }

  if (itemUrls.length === 0) {
    ui.alert(
      'ℹ️ 削除対象がありません',
      '「楽天削除CK」にレ点を付けた行のいずれにも、商品管理番号が取れませんでした。\n\n' +
      '・削除したい行に「親SKU」を入力してください（子SKUがある場合は子SKUも）。\n' +
      '・または「楽天商品管理番号」列に、楽天に登録している商品管理番号（item_url）を直接入力してください。\n\n' +
      '過去出品分を削除する場合、楽天の商品一覧などで確認した商品管理番号を「楽天商品管理番号」列に貼り付けてから、その行に楽天削除CKを付けて再度実行してください。',
      ui.ButtonSet.OK
    );
    return;
  }

  const confirmed = ui.alert('⚠️ 楽天 商品削除の確認', '「楽天削除CK」の付いた ' + itemUrls.length + ' 件を楽天から削除します。\n\nitem-delete.csv をMAKE経由でFTPに送信します。この操作は取り消せません。よろしいですか？', ui.ButtonSet.YES_NO);
  if (confirmed !== ui.Button.YES) return;

  const csvRows = [['商品管理番号（商品URL）']];
  itemUrls.forEach(url => csvRows.push([url]));
  const csvString = csvRows.map(row => row.map(cell => {
    const s = cell === null || cell === undefined ? "" : String(cell);
    return (s.indexOf('"') >= 0 || s.indexOf(',') >= 0 || s.indexOf('\n') >= 0) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(',')).join('\r\n');

  let fileId = null;
  try {
    const blob = Utilities.newBlob('', 'text/csv', RAKUTEN_DELETE_CSV_NAME).setDataFromString(csvString, 'Shift_JIS');
    const saveFolder = DriveApp.getFolderById(CSV_SAVE_FOLDER_ID);
    const file = saveFolder.createFile(blob);
    fileId = file.getId();
  } catch (e) {
    ui.alert('❌ 保存エラー', 'item-delete.csv の保存に失敗しました: ' + e.message, ui.ButtonSet.OK);
    return;
  }

  const sent = sendRakutenDeleteFileToMake(fileId);
  if (!sent) {
    ui.alert('⚠️ 送信結果', 'MAKEへの送信に失敗しました。Driveには保存されています。', ui.ButtonSet.OK);
  }

  const now = new Date();
  for (let i = 0; i < rowIndices.length; i++) {
    const r = rowIndices[i];
    if (colRakutenDeleteFlag !== -1) masterSheet.getRange(r + 1, colRakutenDeleteFlag + 1).setValue(true);
    if (colRakutenDeleteDate !== -1) masterSheet.getRange(r + 1, colRakutenDeleteDate + 1).setValue(now);
    masterSheet.getRange(r + 1, colRakutenDeleteCK + 1).setValue(false);
  }
  ui.alert('✅ 削除依頼を送信しました', (itemUrls.length) + ' 件の item-delete.csv をMAKEへ送信しました。楽天FTPで処理されます。\n\n楽天削除フラグ・楽天削除日時を更新し、楽天削除CKのレ点を外しました。', ui.ButtonSet.OK);
}

/**
 * マスタで商品管理番号に一致する行の楽天削除フラグ・楽天削除日時を更新（Webアプリの1件削除用）
 */
function updateRakutenDeleteFlagInMaster(itemUrl) {
  if (!itemUrl || !String(itemUrl).trim()) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) return;
  const data = masterSheet.getDataRange().getValues();
  const headerRowIdx = 7;
  if (data.length <= headerRowIdx) return;
  const headers = data[headerRowIdx].map(h => String(h).trim());
  const colFlag = headers.indexOf('楽天削除フラグ');
  const colDate = headers.indexOf('楽天削除日時');
  const colRakutenItemUrl = headers.indexOf('楽天商品管理番号');
  const colParentSku = headers.indexOf('親SKU');
  const colChildSku = headers.indexOf('子SKU');
  const getRowItemUrl = (row) => {
    if (colRakutenItemUrl !== -1 && row[colRakutenItemUrl] && String(row[colRakutenItemUrl]).trim()) return String(row[colRakutenItemUrl]).trim();
    const pSku = String(row[colParentSku] || "").trim();
    if (!pSku) return "";
    const cSku = (colChildSku !== -1 && row[colChildSku] != null) ? String(row[colChildSku]).trim() : "";
    if (!cSku) return pSku;
    const lastHyphen = cSku.lastIndexOf('-');
    return pSku + (lastHyphen !== -1 ? cSku.substring(lastHyphen) : "_" + cSku);
  };
  const target = String(itemUrl).trim();
  const now = new Date();
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    if (getRowItemUrl(data[i]) === target) {
      if (colFlag !== -1) masterSheet.getRange(i + 1, colFlag + 1).setValue(true);
      if (colDate !== -1) masterSheet.getRange(i + 1, colDate + 1).setValue(now);
      return;
    }
  }
}

/**
 * Webアプリから呼ばれる楽天1件削除：item-delete.csv を1行で生成 → Drive保存 → MAKE送信 → マスタ更新
 */
function executeRakutenDeleteFromWebApp(itemUrl) {
  if (!itemUrl || !String(itemUrl).trim()) {
    return { success: false, message: '商品管理番号が指定されていません' };
  }
  const csvRows = [['商品管理番号（商品URL）'], [String(itemUrl).trim()]];
  const csvString = csvRows.map(row => row.map(cell => (cell === null || cell === undefined ? "" : String(cell))).join(',')).join('\r\n');
  try {
    const blob = Utilities.newBlob('', 'text/csv', RAKUTEN_DELETE_CSV_NAME).setDataFromString(csvString, 'Shift_JIS');
    const saveFolder = DriveApp.getFolderById(CSV_SAVE_FOLDER_ID);
    const file = saveFolder.createFile(blob);
    const sent = sendRakutenDeleteFileToMake(file.getId());
    updateRakutenDeleteFlagInMaster(String(itemUrl).trim());
    return { success: true, message: sent ? '削除依頼を送信しました。楽天FTPで処理されます。' : 'CSVを保存しましたがMAKEへの送信に失敗しました。' };
  } catch (e) {
    return { success: false, message: '処理に失敗しました: ' + e.message };
  }
}

/**
 * 削除用CSVをMAKEのWebhookに送信（action: delete で削除用と判別）
 */
function sendRakutenDeleteFileToMake(fileId) {
  const MAKE_WEBHOOK_URL = "https://hook.eu1.make.com/sbxvjnqvn7b23l673crdgkb7j1s4sl4w";
  const payload = {
    action: "delete",
    fileId: fileId,
    fileName: RAKUTEN_DELETE_CSV_NAME,
    source: "googleDrive"
  };
  try {
    const response = UrlFetchApp.fetch(MAKE_WEBHOOK_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    if (code === 200) {
      console.log("✅ 楽天削除CSVをMAKEへ送信しました");
      SpreadsheetApp.getActive().toast('楽天削除CSVをMAKEへ送信しました', '楽天削除');
      return true;
    }
    console.error("⚠️ MAKE送信失敗: " + code + " " + response.getContentText());
    return false;
  } catch (e) {
    console.error("❌ 楽天削除MAKE送信エラー: " + e.message);
    return false;
  }
}








// ==========================================
// 6. 確認用シート出力＆自動診断 (変更なし)
// ==========================================
function writeToVerificationSheetWithDiagnosis(csvOutput, settingData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(VERIFY_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(VERIFY_SHEET_NAME);
  
  sheet.clear();
  const headers = settingData[0];
  const systemKeys = settingData[1];
  
  const displayRows = [];
  const headerRow = ['No', '楽天項目名', 'SysKey', '【自動診断】判定コメント'];
  
  const numDataRows = Math.min(csvOutput.length - 1, 10);
  for (let i = 0; i < numDataRows; i++) headerRow.push(`Data${i+1}`);
  displayRows.push(headerRow);

  for (let colIdx = 0; colIdx < headers.length; colIdx++) {
    const itemName = headers[colIdx];
    const sysKey = systemKeys[colIdx];
    
    const rowData = [colIdx + 1, itemName, sysKey, ""];
    let diagnosis = "";
    const valuesToCheck = [];
    
    for (let i = 0; i < numDataRows; i++) {
      const val = csvOutput[i + 1][colIdx]; 
      valuesToCheck.push(val);
      rowData.push(val);
    }

    if (String(valuesToCheck[0]).includes('E+') || String(valuesToCheck[0]).includes('e+')) diagnosis += "⚠️指数;";
    if (sysKey === 'variation_names' && valuesToCheck[0] && valuesToCheck[0].includes('\n')) diagnosis += "❌改行;";
    if (['item_url','item_name','genre_id'].includes(systemKeys[colIdx]) && !valuesToCheck[0]) diagnosis += "❌必須空;";

    rowData[3] = diagnosis || "OK";
    displayRows.push(rowData);
  }

  if (displayRows.length > 0) {
    sheet.getRange(1, 1, displayRows.length, displayRows[0].length).setValues(displayRows);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(4);
    sheet.getRange(1, 1, 1, displayRows[0].length).setFontWeight('bold').setBackground('#d9ead3');
    sheet.getRange(2, 4, displayRows.length - 1, 1).setFontColor('red'); 
    sheet.autoResizeColumns(1, 4);
  }
  SpreadsheetApp.getActive().toast('確認用シートを更新しました。', '完了');
}

// ==========================================
// 7. ユーティリティ関数
// ==========================================
function getColumnIndexMap(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h).trim()] = i);
  return map;
}

/**
 * リサーチでマスタに書き込む際、必須列がマスタに存在するかチェックする。
 * @param {Object} masterColMap ヘッダー名 → 0-based 列インデックス（getColumnIndexMap の戻り値）
 * @param {string[]} [columnNames] チェックする列名の配列。省略時は MASTER_WRITE_COLUMNS_RESEARCH を使用
 * @returns {string[]} 存在しない列名の配列。空ならすべて存在する
 */
function getMissingMasterWriteColumnsResearch(masterColMap, columnNames) {
  const names = columnNames || MASTER_WRITE_COLUMNS_RESEARCH;
  return names.filter(function(c) { return masterColMap[c] === undefined; });
}

function convertScientificToDecimal(num) {
  let str = String(num);
  if (str.indexOf('E') === -1 && str.indexOf('e') === -1) return str.replace(/[^0-9]/g, '');
  let n = Number(num);
  return n.toLocaleString('fullwide', { useGrouping: false });
}

/** 楽天「バリエーション1選択肢定義」用: Shift_JIS換算で maxBytes を超えないよう先頭で切り詰める */
function truncateToByteLength(str, maxBytes) {
  if (str == null || str === '' || maxBytes <= 0) return str === undefined || str === null ? '' : String(str);
  const s = String(str);
  let bytes = 0;
  let i = 0;
  for (; i < s.length; i++) {
    const b = s.charCodeAt(i) < 128 ? 1 : 2;
    if (bytes + b > maxBytes) break;
    bytes += b;
  }
  return s.substring(0, i);
}

function saveCsvWithArchive(dataArray, fileName) {
  try {
    const saveFolder = DriveApp.getFolderById(CSV_SAVE_FOLDER_ID);
    const archiveFolder = DriveApp.getFolderById(CSV_ARCHIVE_FOLDER_ID);
    
    const existingFiles = saveFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      const timestamp = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd_HHmm_");
      file.setName(timestamp + fileName); 
      file.moveTo(archiveFolder); 
    }

    const csvString = dataArray.map(row => row.map(f => {
      let str = f === null || f === undefined ? "" : String(f);
      return str.includes('"') || str.includes(',') || str.includes('\n') ? '"' + str.replace(/"/g, '""') + '"' : str;
    }).join(',')).join('\r\n');
    
    const newFile = saveFolder.createFile(Utilities.newBlob('', 'text/csv', fileName).setDataFromString(csvString, 'Shift_JIS'));
    SpreadsheetApp.getActive().toast(`CSVを保存しました: ${fileName}`, '完了');
    
    // ファイルIDを返す（MAKE連携用）
    return newFile.getId();
  } catch (e) {
    SpreadsheetApp.getUi().alert(`保存エラー: ${e.message}`);
    return null;
  }
}

function debugSystemCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  results.push(`マスタシート: ${masterSheet ? 'OK' : '❌ なし'}`);
  if (masterSheet) {
    const data = masterSheet.getDataRange().getValues();
    let headerRow = -1;
    for(let r=0; r<20; r++) if(data[r].includes(ANCHOR_HEADER_NAME)) { headerRow = r; break; }
    results.push(`ヘッダー行: ${headerRow !== -1 ? (headerRow + 1) + '行目' : '❌ なし'}`);
  }
  SpreadsheetApp.getUi().alert("【診断結果】\n" + results.join('\n'));
}





// ==============================================================================================================================






//パート3
// ==========================================
// 【AI出品ツール Ver 36.0 完全版】 Part 3/3
// ==========================================

// Gemini API呼び出し
// Gemini API呼び出し (JSON生成用・検索OFF)
function callGeminiVisionAPI(promptText, base64Image, mimeType) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_GEMINI}:generateContent?key=${getGeminiApiKey()}`;
  
  const parts = [{ text: promptText }];
  if (base64Image && mimeType) {
    parts.push({
      inline_data: { mime_type: mimeType, data: base64Image }
    });
  }
  
  const payload = {
    contents: [{ parts: parts }],
    // ★修正: 本番生成時は検索OFF (JSONモード優先)
    generationConfig: { response_mime_type: "application/json" }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  
  if (json.error) throw new Error(json.error.message);
  
  try {
    const rawText = json.candidates[0].content.parts[0].text;
    return JSON.parse(rawText);
  } catch (e) {
    return { error: "JSON parse error", raw: json };
  }
}

// Gemini API呼び出し (検索調査用・JSONモードOFF)
function callGeminiSearchAPI(promptText) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_GEMINI}:generateContent?key=${getGeminiApiKey()}`;
  
  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    // ★修正: 最新モデルに対応したシンプル構文に変更 (これでエラーが消えます)
    tools: [{ google_search: {} }],
    generationConfig: {} 
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  
  if (json.error) throw new Error("Search Error: " + json.error.message);
  
  try {
    return json.candidates[0].content.parts[0].text;
  } catch (e) {
    return "検索結果取得失敗";
  }
}

// OpenAI API呼び出し (復活)
function callOpenAIVisionAPI(promptText, base64Image, mimeType) {
  const url = 'https://api.openai.com/v1/chat/completions';
  const content = [{ type: "text", text: promptText }];
  if (base64Image && mimeType) {
    content.push({
      type: "image_url",
      image_url: { url: `data:${mimeType};base64,${base64Image}` }
    });
  }
  const payload = {
    model: MODEL_OPENAI,
    messages: [{ role: "user", content: content }],
    response_format: { type: "json_object" }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${getOpenAiApiKey()}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error(json.error.message);
  return JSON.parse(json.choices[0].message.content);
}




// ==============================================================================================================================



// ==========================================
// Part 3/3：プロンプト生成関数 (あなたの指示 + 検索結果統合版)
// ==========================================

// ★変更点1: 引数の最後に 'searchResult' を追加
function createFullPrompt(name, jan, asin, rCand, yCatCand, yBrandCand, searchResult) {
  return `
あなたは、楽天市場、Yahoo!ショッピング、Amazonの各プラットフォームのアルゴリズムとユーザー行動を熟知した、ECマーケティング界のトップコンサルタントです。

【目的】
以下の商品情報と画像(あれば)、および**[Web検索による調査結果]**を統合し、顧客の購買心理を刺激し、かつ各サイトの検索SEOを最大化させる出品用JSONデータを作成してください。

[商品情報]
商品名: ${name}
JAN: ${jan}
ASIN: ${asin}

[参考カテゴリ候補(CSV検索結果)]
楽天: ${rCand}
Yahoo: ${yCatCand}
Yahooブランド: ${yBrandCand}

// ★変更点2: ここにWeb検索の結果を埋め込みます（これがないと正確なスペックが反映されません）
[Web検索による調査結果 (最新のスペック情報)]
${searchResult}
※注意: この調査結果には、現在の正確な市場価格、梱包サイズ、原材料などが含まれています。
推測ではなく、この調査結果の数値を最優先でJSONに反映させてください。

【STEP 1：商品分析（思考プロセス）】
データを生成する前に、以下の3点を内部で深く分析してください（出力は不要ですが、結果をすべての項目に反映させてください）。
1. USP（独自の強み）: 競合と比較して、この商品が選ばれる決定的な理由は何か？
2. ペインポイント: 顧客がこの商品を買うことで解決される悩みは何か？
3. 検索意図: ユーザーが検索窓に打ち込む「悩み」や「理想の状態」は何か？

【STEP 2：データ生成ルール】
以下のJSONキーに対して、指定された「マスタ項目ルール」を厳守して値を出力してください。
出力はJSON形式のみです。

{
  "main_kw": ["KW1", "KW2", "KW3"],
  // ▼ルール: 商品そのものや使用感をイメージできるメインキーワード候補3つ。「商品名」ではなく「価値」を抽出してください。ユーザーがこの商品を使って「最高だ」と感じる瞬間を象徴する、インパクトの強いキーワードを3つ選定してください。

  "search_kw_8": "スペース区切り8単語",
  // ▼ルール: Web検索されやすいキーワードを検索されやすい順に8つ半角スペースで区切って並べて。検索上位を狙うためのSEO特化型キーワードです。ユーザーが「今すぐ欲しい」ときに検索窓に入れる言葉を、検索頻度が高いと推測される順に、必ず「半角スペース区切り」で8つ並べてください。抽象的な言葉（例：最高、良い、おすすめ）は避け、具体的な特徴や用途（例：軽量、大容量、通勤用）に絞ってください。

  "amazon_kw": "Amazon用検索語句(150字以内)",
  // ▼ルール: 検索ワードを140文字～150文字までで半角スペースを間に入れて作成して。商品名・メーカー名・ブランド名は厳禁。140〜150文字を使い切り、類義語や利用シーンで埋めること。

  "ctr_copy": {
    "char_8": ["8文字以下案1", "案2", "案3"],
    // ▼ルール: アイコンやバナー上の「一言」用。最も強い「単語」をぶつける。ターゲットがスマホでスクロールする手を止め、思わずタップしてしまうコピー。

    "char_10": ["10文字以下案1", "案2", "案3"],
    // ▼ルール: 商品名冒頭の【 】内用。一瞬でメリットを伝える。

    "char_15": ["15文字以下案1", "案2", "案3"],
    // ▼ルール: 広告の1行目用。ベネフィットを具体化する。

    "char_20": ["20文字以下案1", "案2", "案3"]
    // ▼ルール: 訴求の完成形。悩み解決＋αの付加価値を伝える。
  },
  // ※CTRコピー全体の指針: 以下の4つの切り口をバランスよく取り入れること。
  // 「即時性・希少性」（例：本日終了、在庫僅か）
  // 「具体的便益」（例：朝の時短5分、腰が楽に）
  // 「信頼・実績」（例：楽天1位、満足度98%）
  // 「自分事化」（例：30代の乾燥肌に、キャンプ初心者に）

  "rakuten_slug": "楽天用URL末尾",
  // ▼ルール: 楽天compassに登録する商品URLを半角2~30文字でSEO効果が望めるよう設定して下さい。SEOを意識し、英数字とハイフンで意味のある英単語（例：water-bottle-stainless）を設定すること。

  "yahoo_catch_copy": ["キャッチ1", "キャッチ2", "キャッチ3"],
  // ▼ルール: 【重要：文字数厳守】「20文字以上、30文字以内」で記述すること。
  //
  // 【新・作成指針：ワンワード・インパクト】
  // スマホの一瞬のスクロールで「私のことだ」と思わせる、「直感的なベネフィット」を提示してください。
  //
  // × 悪い例（機能のみ）：「国産生姜100%使用の生姜湯」
  // 〇 良い例（感覚＋機能）：「一口でポカポカ！濃厚な国産生姜100%」
  //
  // 指示：
  // 1. 左側10文字に、商品を使った瞬間の「快感」や「解決」を表す言葉（ポカポカ、スッキリ、モチモチ等）を置くこと。
  // 2. 「実績」よりも「実利（どうなるか）」を優先し、短い言葉で畳みかけるリズムを作ること。

  "catch_copy": ["キャッチ1", "キャッチ2", "キャッチ3"],
  // ▼ルール: 【重要：文字数厳守】「60文字以上、80文字以内」で記述すること。
  //
  // 【新・作成指針：ミニ・ストーリーの見出し化】
  // 単語の羅列ではなく、一つの「短い文章」として読ませ、クリックしたくなる構成にしてください。
  //
  // 指示：
  // 1. 【隅付き括弧】の中身は、単なるキーワードではなく「ターゲットの呼びかけ」にする。
  //    （例：【冷え性のあなたへ】【週末のご褒美に】など）
  // 2. SEOキーワード（商品名、成分など）は必須だが、文脈の中で自然に繋げること。
  //    無理やり詰め込まず、「〇〇だから、××な時に最適」という論理構成で文章を繋ぐこと。
  // 3. 読んだ人が「あ、その手があったか」と膝を打つような、具体的な利用シーン（ギフト、夜食、朝活など）を必ず一つ提案すること。

  "description": "商品説明文",
  // ▼ルール: 【目的】 読者に「これは私のための商品だ」と確信させ、購入ボタンを押させること。
  // 【重要：文字数厳守】 句読点を含め、必ず300文字以上、350文字以内で記述すること（300文字未満はNG）。
  //
  // 【ライティング構成：シズル感＆ベネフィット融合型】
  // 1. ファーストタッチ（五感への訴求）:
  //    冒頭で商品の最大の特徴を伝えつつ、それを手にした時の「感覚（香り、手触り、喉越し、視覚的変化）」を情緒的に表現してください。
  //    （例：一口飲めば、ピリッとした刺激と深いコクが広がり…）
  //
  // 2. ロジカル・エビデンス（納得の根拠）:
  //    なぜその体験ができるのか、素材・製法・成分などの「動かぬ証拠」をプロの視点で解説してください。
  //    （例：蒸して乾燥させることで成分を引き出した「蒸し生姜」を100%使用）
  //
  // 3. 「不」の解消とプラスアルファ（ベネフィット）:
  //    その機能が、日常のどんな「不（不安、不満、不便）」を解消し、どんな素敵な時間をもたらすかを具体的に描写してください。
  //    （例：冷えが気になる季節のセルフケアや、寝る前のリラックスタイムに）
  //
  // 4. クロージング（自分事化）:
  //    「贈り物としても喜ばれる」「常備しておきたい逸品」など、購入後の具体的な活用イメージで締めくくってください。
  //
  // 【禁止事項】
  // ・「ぜひ検討してください」「おすすめです」といった、売り手側の使い古された言葉は使わない。
  // ・事実の羅列（スペック表）にしない。必ずその事実がもたらす「恩恵（ベネフィット）」とセットで書くこと。

  "bullets": ["箇条書き1", "箇条書き2", "箇条書き3", "箇条書き4", "箇条書き5"],
  // ▼ルール: 【重要：文字数厳守】各「80文字以上、100文字以内」で記述すること。
  //
  // 【全体指針：5つの視点で「欲しい」を確信させる】
  // 以下の5つのテーマに基づいて、それぞれ異なる角度から商品の魅力を「短い物語」として伝えてください。
  // 下記1～5のテーマごとに一目で後述の文章のイメージが伝わる見出しを【】中に入れて文頭につけて下さい。
  //
  // 1. [テーマ：圧倒的な差別化ポイント]：他社製品と決定的に違う「最大のウリ」は何か？
  //    見出しのヒント：その商品にしかない「強み」を凝縮した言葉 (例：【通常の生姜とは別格の温め力】）
  //    （文章例：通常の生姜ではなく「蒸し生姜」だからこそ実現した、驚きの温めパワー。）
  //    
  // 2. [テーマ：リアルな感覚と体験]：使った瞬間の「感覚（味、香り、肌触り）」はどうなるか？
  //    見出しのヒント：五感（味、香り、触感）に訴えかける言葉 （例：【心までほどける芳醇な香りと喉越し】）
  //    （文章例：お湯を注いだ瞬間に広がる芳醇な香り。とろりとした喉越しが心まで満たします。）
  //    
  // 3. [テーマ：安心の根拠と誠実さ]：なぜ安全なのか？（産地、成分、無添加など）
  //    見出しのヒント：安全性やこだわりを象徴する言葉 （例：【体への優しさを追求した国産無添加】）
  //    （文章例：毎日体に入れるものだから、余計なものは一切不使用。国産素材にこだわり抜きました。）
  //    
  // 4. [テーマ：利用、解決シーンの具体化]：日常のどんな場面で役立つか？
  //    見出しのヒント：どんな時、どんな場所で役立つかを示す言葉 （例：【冷えたオフィスでホッと一息つく魔法】）
  //    （文章例：冷房で冷え切ったオフィスの休憩時間に。ホッと一息つく魔法のような一杯。）
  //    
  // 5. [テーマ：最高の自分へのご褒美・大切な人へのギフト、贈答]：誰におすすめか？どんな気持ちになれるか？
  //    見出しのヒント：誰に向けた、どんな想いのギフトかを示す言葉 （例：【大切な人の体を労わる、いたわりの一箱】）
  //    （文章例：頑張っている自分へのご褒美や、大切な方への体を気遣う贈り物としても最適です。）
  //      
  //
  // 【記述ルール】
  // - 文頭には【 】で、中身を要約した「雑誌の見出し」のようなキャッチーなフレーズを付けること。
  // - 「～です。～ます。」調で、語りかけるような丁寧なトーンで統一すること。

  "market_price_research": "市場価格帯の調査結果",
  // ★変更点3: ここだけ変えます。調査結果を使うように指示。
  // ▼ルール: 上記の [Web検索による調査結果] に価格情報がある場合は、その数値を採用し「(Web調査済み)」と記載すること。
  // 情報がない場合のみ、ブランド分析や競合比較からの推測を行うこと。

  "package": { "width": 0, "depth": 0, "height": 0, "weight": 0 },
  // ★変更点4: ここだけ変えます。調査結果を使うように指示。
  // ▼ルール: 上記の [Web検索による調査結果] にサイズ・重量情報がある場合は、その数値をそのまま採用すること。
  // 情報がない場合のみ、メール便/宅急便コンパクト/大型の推論フローを適用すること。

  "data_source": "情報の取得元(画像/JAN検索/推測/Web検索)",
  // ▼ルール: Google検索を使用した場合は「Web検索」、画像のみの場合は「画像解析」と明記すること。

  // 以下は属性抽出用
  "variation_suggestions": [
    {"axis1": "カラー", "axis2": "サイズ", "value": "赤 S"}
  ],
  "specs": {
    "ingredients": "原材料", "expiry": "賞味期限", "storage": "保存方法",
    // ★変更点5: ここだけ変えます。調査結果を使うように指示。
    // ▼ルール: [Web検索による調査結果] に原材料や保存方法の記載がある場合は、一字一句正確に転記すること。
    // 画像裏面やWeb情報がない場合に限り「不明」と回答可。
    "material": "素材", "dimensions": "本体サイズ", "color": "カラー", "other": "その他"
  },
  
  "rakuten_attributes": {
    "genre_suggestions": [{"id": "...", "name": "..."}],
    "series": "...", "brand": "...", "origin": "...",
    "total_qty": 0, "total_weight_g": 0, "total_vol_ml": 0
  },
  "yahoo_attributes": {
    "category_suggestions": [{"id": "...", "name": "..."}],
    // ▼ルール: 【重要：逆引き検索用】
    // 候補に適切なものがない場合、idを空欄にし、nameには「詳細なカテゴリ名」だけでなく、
    // 「ソフトドリンク」「健康食品」「調味料」のような【大分類のカテゴリ名】も必ず提案してください。
    // （システムがその名前を使って再検索するため、ヒットしやすい一般的な名称を含めることが重要です）

    "brand_suggestions": [{"id": "...", "name": "..."}]
  },
  "amazon_attributes": {
    "is_heat_sensitive": false, "serving_size": "...", "serving_size_unit": "...",
    "unit_count_type": "...", "dangerous_goods": "...", "contains_liquid": false
  }
}
`;
}

// 調査専用プロンプト生成関数 (復活)
function createSearchPrompt(name, jan, maker) {
  return `
以下の商品のスペック情報をWeb検索して調査し、結果を簡潔なテキストで教えてください。

商品名: ${name}
JAN: ${jan}
メーカー: ${maker}

【調査項目】
1. 現在の市場価格（Amazon, 楽天, 公式サイト等の販売価格）
2. 梱包サイズ（幅 x 奥行 x 高さ cm）と梱包重量（g）
   ※梱包サイズ不明な場合は「商品本体サイズ」と「内容量」
3. 原材料名（食品の場合。パッケージ裏面の記載内容）
4. 賞味期限・保存方法
5. 正確な商品カテゴリ（例：ソフトドリンク、調味料など）

【出力形式】
箇条書きで事実のみを出力してください。挨拶や前置きは不要です。
`;
}




// ==============================================================================================================================



// 候補検索

//楽天とYahoo!のジャンIDやカテゴリーIDを検索するコード




// --- 【修正後】（以下のコードに完全に置き換えてください） ---------
function searchCandidatesWithScore(productName, makerName, folderId, fileNamePrefix, isBrandSearch = false, targetColIndex = 1) {
  if (!folderId) return "（フォルダ設定なし）";
  
  // 1. キーワード分割 (漢字・カタカナ・英数単位)
  // 例: "六漢生姜湯" -> "六漢", "生姜湯"
  const rawText = (String(productName) + " " + String(makerName)).replace(/　/g, ' ');
  const keywords = rawText.match(/[一-龠]+|[ァ-ヴー]+|[a-zA-Z0-9]+|[ぁ-ん]+/g) || [];
  const validKeywords = [...new Set(keywords.filter(w => w.length >= 2))];

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let targetFile = null;
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();
      if (fileNamePrefix) { if (name.indexOf(fileNamePrefix) === 0) { targetFile = f; break; } }
      else { if (name.endsWith('.csv') || name.indexOf('.') === -1) { targetFile = f; break; } }
    }
    if (!targetFile) return "（CSVが見つかりません）";

    const blob = targetFile.getBlob();
    let text = "";
    try { text = blob.getDataAsString('Shift_JIS'); } catch(e) { text = blob.getDataAsString('UTF-8'); }
    const lines = text.split(/\r\n|\n/);
    
    // ★希少性ロジック: 全行スキャンして各キーワードのヒット数(登場頻度)をカウント
    const keywordHitMap = {}; 
    validKeywords.forEach(k => keywordHitMap[k] = 0);

    for (let i = 1; i < lines.length; i++) {
      const line = lines[i];
      if (!line) continue;
      const parts = line.split(',');
      if (parts.length <= targetColIndex) continue;
      
      const targetText = parts[targetColIndex] || ""; 
      const searchText = isBrandSearch ? line : targetText; 

      validKeywords.forEach(k => {
        if (searchText.includes(k)) {
          keywordHitMap[k]++;
        }
      });
    }

    // ★スコア計算: ヒット数が少ない(希少な)キーワードほど高得点にする
    // 例: 食品(1000件ヒット) -> 0.1点 / 生姜湯(5件ヒット) -> 20点
    const lineScores = {}; 

    validKeywords.forEach(k => {
      const hitsCount = keywordHitMap[k];
      if (hitsCount > 0) {
        const scorePerHit = 100 / hitsCount; // 希少性スコア係数
        
        // 再スキャンしてスコア加算 (メモリ効率のため簡易実装)
        // ※データ量が多い場合はMap活用推奨だが、数万行ならこのループでも動作範囲内
        for (let i = 1; i < lines.length; i++) {
          const line = lines[i];
          if (!line) continue;
          const parts = line.split(',');
          if (parts.length <= targetColIndex) continue;
          const targetText = parts[targetColIndex] || "";
          const searchText = isBrandSearch ? line : targetText;

          if (searchText.includes(k)) {
            if (!lineScores[i]) lineScores[i] = 0;
            lineScores[i] += scorePerHit;
          }
        }
      }
    });

    // 結果生成
    const scoredCandidates = [];
    for (const [lineIdx, score] of Object.entries(lineScores)) {
      const line = lines[lineIdx];
      const parts = line.split(',');
      const id = parts[0];
      const targetText = parts[targetColIndex] || "";
      scoredCandidates.push({ str: `${id}: ${targetText}`, score: score });
    }
    
    scoredCandidates.sort((a, b) => b.score - a.score);
    const topCandidates = scoredCandidates.slice(0, 20).map(c => c.str);
    
    return topCandidates.length === 0 ? "（一致候補なし：AIによる推測を推奨）" : topCandidates.join('\n');

  } catch (e) { return `（検索エラー: ${e.message}）`; }
}


// 【新規追加】 AI推測ジャンル名からの逆引き補完関数
// (Part 3/3 のコードの一番下に追加してください)
function refineGenreData(jsonData, rFolderId, yCatFolderId, yBrandFolderId) {
  if (!jsonData) return jsonData;

  // 汎用逆引き関数
  const doReverseSearch = (suggestionName, folderId, isBrand, targetCol) => {
    if (!suggestionName) return null;
    // AIの言葉をそのままキーワードとして検索 (メーカー名は空でOK)
    const hit = searchCandidatesWithScore(suggestionName, "", folderId, "", isBrand, targetCol);
    if (hit && !hit.includes('一致候補なし') && !hit.includes('エラー')) {
      const topHit = hit.split('\n')[0]; // スコア1位
      const match = topHit.match(/^(\d+):/);
      if (match) return match[1]; // IDを返す
    }
    return null;
  };

  // 1. 楽天ジャンル補完
  if (jsonData.rakuten_attributes && jsonData.rakuten_attributes.genre_suggestions) {
    jsonData.rakuten_attributes.genre_suggestions.forEach(sug => {
      if (!sug.id || sug.id.includes('不明')) {
        const newId = doReverseSearch(sug.name, rFolderId, false, 1);
        if (newId) sug.id = newId;
      }
    });
  }

  // 2. Yahooカテゴリ補完
  if (jsonData.yahoo_attributes && jsonData.yahoo_attributes.category_suggestions) {
    jsonData.yahoo_attributes.category_suggestions.forEach(sug => {
      if (!sug.id || sug.id.includes('不明')) {
        const newId = doReverseSearch(sug.name, yCatFolderId, false, 2);
        if (newId) sug.id = newId;
      }
    });
  }

  // 3. Yahooブランド補完
  if (jsonData.yahoo_attributes && jsonData.yahoo_attributes.brand_suggestions) {
    jsonData.yahoo_attributes.brand_suggestions.forEach(sug => {
      if (!sug.id || sug.id.includes('不明')) {
        const newId = doReverseSearch(sug.name, yBrandFolderId, true, 1);
        if (newId) sug.id = newId;
      }
    });
  }

  return jsonData;
}


// ------------------------------------------------------------



// 取得データ書き込み
function writeSplitted(sheet, r, g, o, startCol, inputUrl) {
  let col = startCol; 
  const set3 = (valG, valO, mergeMode = 'none', highlightDiff = false) => {
    const vG = (valG === undefined || valG === null) ? '' : String(valG);
    const vO = (valO === undefined || valO === null) ? '' : String(valO);
    sheet.getRange(r, col).setValue(vG);
    sheet.getRange(r, col + 1).setValue(vO);
    let valCorrect = '';
    if (mergeMode === 'merge_comma') valCorrect = mergeAndDedup(vG, vO, /[,、\n]+/);
    else if (mergeMode === 'merge_space') valCorrect = mergeAndDedup(vG, vO, /[\s　]+/);
    else if (mergeMode === 'merge_lines') valCorrect = mergeAndDedup(vG, vO, /\n/);
    else if (mergeMode === 'gemini') valCorrect = vG ? vG : vO;
    else if (mergeMode === 'expiry') valCorrect = selectSpecificExpiry(vG, vO);
    else valCorrect = (vG.length >= vO.length) ? vG : vO;
    sheet.getRange(r, col + 2).setValue(valCorrect);
    if (highlightDiff && vG !== vO) { sheet.getRange(r, col).setBackground('#fce5cd'); sheet.getRange(r, col + 1).setBackground('#fce5cd'); }
    col += 3;
  };

  // URL書き込み
  set3(inputUrl, inputUrl);

  set3((g.main_kw||[]).join(','), (o.main_kw||[]).join(','), 'merge_comma');
  set3(g.search_kw_8, o.search_kw_8, 'merge_space');
  set3(g.amazon_kw, o.amazon_kw, 'none', true);
  
  const fmtCtr = d => {
    if (!d || !d.ctr_copy) return '';
    let lines = [];
    for (const [k, v] of Object.entries(d.ctr_copy)) { if (Array.isArray(v)) v.forEach(copy => lines.push(`[${k}] ${copy}`)); }
    return lines.join('\n');
  };
  set3(fmtCtr(g), fmtCtr(o), 'merge_lines');
  
  set3(g.rakuten_slug, o.rakuten_slug);
  set3((g.yahoo_catch_copy||[]).join('\n'), (o.yahoo_catch_copy||[]).join('\n'));
  set3((g.catch_copy||[]).join('\n\n'), (o.catch_copy||[]).join('\n\n'));
  set3(g.description, o.description, 'gemini');
  for(let i=0; i<5; i++){ set3((g.bullets?.[i] || ''), (o.bullets?.[i] || ''), 'gemini'); }
  set3(g.market_price_research, o.market_price_research);
  
  const pG = g.package || {}; const pO = o.package || {};
  set3(pG.width, pO.width); set3(pG.depth, pO.depth); set3(pG.height, pO.height);
  const sumG = (parseFloat(pG.width)||0) + (parseFloat(pG.depth)||0) + (parseFloat(pG.height)||0);
  const sumO = (parseFloat(pO.width)||0) + (parseFloat(pO.depth)||0) + (parseFloat(pO.height)||0);
  set3(sumG || '', sumO || ''); set3(pG.weight, pO.weight);
  const cG = calcShippingTariff(sumG, pG.weight); const cO = calcShippingTariff(sumO, pO.weight);
  set3(cG, cO);
  
  set3(g.data_source, o.data_source);
  
  const fmtVar = (suggestions, key) => {
    if (!suggestions || !Array.isArray(suggestions)) return '';
    return suggestions.map((s, i) => `${i+1}. ${s[key] || ''}`).join('\n');
  };
  set3(fmtVar(g.variation_suggestions, 'axis1'), fmtVar(o.variation_suggestions, 'axis1'), 'gemini');
  set3(fmtVar(g.variation_suggestions, 'axis2'), fmtVar(o.variation_suggestions, 'axis2'), 'gemini');
  
  const sG = g.specs || {}; const sO = o.specs || {};
  set3(sG.ingredients, sO.ingredients); set3(sG.expiry, sO.expiry, 'expiry'); set3(sG.storage, sO.storage);
  set3(sG.material, sO.material); set3(sG.dimensions, sO.dimensions); set3(sG.color, sO.color); set3(sG.other, sO.other);
  
  const rG = g.rakuten_attributes || {}; const rO = o.rakuten_attributes || {};
  const fmt3 = (list, key) => (list && Array.isArray(list)) ? list.map((x,i)=>`${i+1}. ${x[key]||''}`).join('\n') : '';
  
  // ★修正箇所: g.genre... を rG.genre... に変更 (これで正しくIDが出ます)
  set3(fmt3(rG.genre_suggestions, 'id'), fmt3(rO.genre_suggestions, 'id'), 'gemini');
  set3(fmt3(rG.genre_suggestions, 'name'), fmt3(rO.genre_suggestions, 'name'), 'gemini');
  
  set3(rG.series, rO.series); set3(rG.brand, rO.brand); set3(rG.origin, rO.origin);
  set3(rG.total_qty, rO.total_qty); set3(rG.total_weight_g, rO.total_weight_g); set3(rG.total_vol_ml, rO.total_vol_ml);
  
  const yG = g.yahoo_attributes || {}; const yO = o.yahoo_attributes || {};
  // Yahoo側も yG を使うよう念のため確認・統一
  set3(fmt3(yG.category_suggestions, 'id'), fmt3(yO.category_suggestions, 'id'), 'gemini');
  set3(fmt3(yG.category_suggestions, 'name'), fmt3(yO.category_suggestions, 'name'), 'gemini');
  set3(fmt3(yG.brand_suggestions, 'id'), fmt3(yO.brand_suggestions, 'id'), 'gemini');
  
  const aG = g.amazon_attributes || {}; const aO = o.amazon_attributes || {};
  set3(aG.is_heat_sensitive, aO.is_heat_sensitive); set3(aG.serving_size, aO.serving_size);
  set3(aG.serving_size_unit, aO.serving_size_unit); set3(aG.unit_count_type, aO.unit_count_type);
  set3(aG.dangerous_goods, aO.dangerous_goods); set3(aG.contains_liquid, aO.contains_liquid);
}


// ------------------------------------------------------------



function mergeAndDedup(str1, str2, splitter) {
  const arr1 = str1 ? str1.split(splitter) : [];
  const arr2 = str2 ? str2.split(splitter) : [];
  const unique = [...new Set([...arr1, ...arr2].map(s => s.trim()).filter(s => s))];
  return unique.join(String(splitter).includes('s') ? ' ' : (String(splitter).includes('n') ? '\n' : ','));
}
function selectSpecificExpiry(valG, valO) {
  const ngWords = ['不明', '未記載', '判読', '確認でき', '読み取れ', 'unknown'];
  const isNg = (txt) => ngWords.some(w => txt.includes(w));
  if (isNg(valG) && !isNg(valO)) return valO;
  if (!isNg(valG) && isNg(valO)) return valG;
  return valG ? valG : valO;
}
function setMasterFormulas(sheet, startCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const masterStartCol = startCol + (COMPARE_ITEMS.length * 3);
  for (let i = 0; i < COMPARE_ITEMS.length; i++) {
    const letter = columnToLetter(startCol + (i * 3) + 2);
    sheet.getRange(2, masterStartCol + i, lastRow - 1, 1).setFormula(`=${letter}2`);
  }
}
function columnToLetter(c) {
  let t, l = ''; while (c > 0) { t = (c - 1) % 26; l = String.fromCharCode(t + 65) + l; c = (c - t - 1) / 26; } return l;
}
function calcShippingTariff(size, weight) {
  if (!size && !weight) return '';
  const s = parseFloat(size) || 0;
  const w = parseFloat(weight) || 0;
  if (s <= 60 && w <= 2000) return '60サイズ';
  if (s <= 80 && w <= 5000) return '80サイズ';
  if (s <= 100 && w <= 10000) return '100サイズ';
  return '大型';
}




// ==============================================================================================================================




// ==========================================
// 【Ver 47.0】AI画像自動仕分け (数字マッチング・サブ共通化・ゴミ排除)
// ==========================================

// --- 1. 設定・定数 ---

// ★RMS認証情報は Script Properties（RAKUTEN_LICENSE_KEY, RAKUTEN_SERVICE_SECRET）で設定

// ★Googleドライブ設定 (URL末尾の長いID)
const DRIVE_IMAGE_SOURCE_FOLDER_ID = '1_YNjcEfNATrgr0J5ci_Lp3wC_r2RTRsz'; // 未整理画像のフォルダID
const DRIVE_IMAGE_ARCHIVE_FOLDER_ID = '1GS5LL833jZ4eUPcCS4JonBRA0HGb04u2'; // 完了後の移動先ID








// ★AI設定
const GEMINI_VISION_MODEL = 'models/gemini-2.0-flash';

// シート設定
const SHEET_NAME_MATRIX = '★画像AIマッチング(操作用)';
const WORK_FOLDER_NAME = '★[システム用]一時作業フォルダ_削除禁止';

// マスタ列名
const COL_NAME_PARENT_SKU = '親SKU'; 
const COL_NAME_CHILD_SKU = '子SKU';
const COL_NAME_ITEM_NAME = '商品名';
const COL_NAME_VAR_NAME = 'バリエーション値';
const COL_NAME_REF_IMAGE_PART = '参考情報(画像URL';
const COL_NAME_CHECK = '出品CK';

// 画像列名定義 (メイン10列 + サブ10列 = 計20列)
const RAKUTEN_IMAGE_COLS = [
  '楽天メイン画像1', '楽天メイン画像2', '楽天メイン画像3', '楽天メイン画像4', '楽天メイン画像5',
  '楽天メイン画像6', '楽天メイン画像7', '楽天メイン画像8', '楽天メイン画像9', '楽天メイン画像10',
  '楽天サブ画像1', '楽天サブ画像2', '楽天サブ画像3', '楽天サブ画像4', '楽天サブ画像5',
  '楽天サブ画像6', '楽天サブ画像7', '楽天サブ画像8', '楽天サブ画像9', '楽天サブ画像10'
];

// --- メニュー --- (重複のためコメントアウト: 80行目のonOpenを使用)
// function onOpen() {
//   SpreadsheetApp.getUi().createMenu('AI出品ツール')
//     .addItem('1. シート初期化', 'initializeSheetComparison')
//     .addItem('2. 全データ一括生成', 'generateListingDataComparison')
//     .addSeparator()
//     .addItem('3. マスタへ同期', 'syncAiDataToMaster')
//     .addSeparator()
//     .addItem('4. 楽天CSV出力 (診断・保存)', 'generateRakutenCSV')
//     .addSeparator()
//     .addItem('5. 🤖 AI画像仕分けシート作成', 'generateAiImageMatrix')
//     .addItem('6. 🚀 リネーム＆アップロード実行', 'executeRenameAndUploadFromMatrix')
//     .addSeparator()
//     .addItem('★システム診断を実行', 'debugSystemCheck')
//     .addToUi();
// }




/**
 * 機能1: 画像をAI解析し、数字ロジックで振り分ける (メイン10列+サブ10列 安全版)
 */
function generateAiImageMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  // --- 1. マスタデータの読み込み ---
  const mValues = masterSheet.getDataRange().getValues();
  const headerRowIdx = getAnchorRowIndex(mValues); 

  if (headerRowIdx === -1) { 
    ui.alert(`マスタシートのヘッダー行が見つかりません。\n「${COL_NAME_PARENT_SKU}」と「${COL_NAME_CHECK}」の両方を含む行が必要です。`); 
    return; 
  }

  const headers = mValues[headerRowIdx];
  const colMap = getColumnIndexMap(headers);
  
  const idxP = colMap[COL_NAME_PARENT_SKU];
  const idxC = colMap[COL_NAME_CHILD_SKU];
  const idxName = colMap[COL_NAME_ITEM_NAME];
  const idxVar = colMap[COL_NAME_VAR_NAME];
  const idxRef = colMap['▼マスタ(参考情報(画像URL))'] !== undefined ? colMap['▼マスタ(参考情報(画像URL))'] : -1;
  const idxCheck = colMap[COL_NAME_CHECK];

  if (idxP === undefined || idxCheck === undefined) { 
    ui.alert('必須列（親SKU または 出品CK）が見つかりません'); return; 
  }

  // --- 商品構造リスト作成 ---
  const productGroups = {}; 
  const displayRows = [];   
  let currentPCode = "";

  // ★追加: 候補エリアの上限数 (50枚)
  const CANDIDATE_LIMIT = 50;
  
  for (let i = headerRowIdx + 1; i < mValues.length; i++) {
    const row = mValues[i];
    const pCode = String(row[idxP]).trim();
    const cCode = (idxC !== undefined) ? String(row[idxC]).trim() : "";
    
    if (pCode === "") continue;

    if (cCode === "") {
      const checkVal = row[idxCheck];
      const isChecked = (checkVal === true || checkVal === 1 || String(checkVal).toUpperCase() === 'TRUE');
      
      if (isChecked) {
        currentPCode = pCode;
        const name = (idxName !== undefined) ? String(row[idxName]).trim() : "";
        let refUrl = (idxRef !== -1) ? String(row[idxRef]).trim() : "";
        
        productGroups[pCode] = {
          type: 'Parent',
          pCode: pCode,
          cCode: '',
          name: name,
          varName: '',
          refImg: refUrl,
          assignedMains: new Array(10).fill(''), 
          assignedSubs: new Array(10).fill(''), 
          assignedCandidates: new Array(CANDIDATE_LIMIT).fill(''), // ★候補枠を追加
          children: [] 
        };

        
        displayRows.push(productGroups[pCode]);
      } else {
        currentPCode = "";
      }
    } else {
      // ★修正: 親が処理対象(レ点あり)の場合のみ子を追加。親が対象外ならスキップする安全装置
      if (currentPCode === pCode && productGroups[pCode]) {
        const varName = (idxVar !== undefined) ? String(row[idxVar]).trim() : "";
        const childObj = {
          type: 'Child',
          pCode: pCode,
          cCode: cCode,
          name: productGroups[pCode].name,
          varName: varName,
          refImg: '', 
          assignedMains: new Array(10).fill(''), 
          assignedSubs: new Array(10).fill(''),
          assignedCandidates: new Array(CANDIDATE_LIMIT).fill('') // ★候補枠を追加
        };
        productGroups[pCode].children.push(childObj);
        displayRows.push(childObj);
      }
    }
  }

  if (displayRows.length === 0) { ui.alert('出品対象の商品がありません。'); return; }

  // --- 2. 画像ファイルの準備 ---
  let sourceFolder, workFolder;
  try {
    sourceFolder = DriveApp.getFolderById(DRIVE_IMAGE_SOURCE_FOLDER_ID);
    const folders = sourceFolder.getFoldersByName(WORK_FOLDER_NAME);
    if (folders.hasNext()) workFolder = folders.next();
    else workFolder = sourceFolder.createFolder(WORK_FOLDER_NAME);
    try { workFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
  } catch (e) {
    ui.alert(`フォルダエラー: ${e.message}`); return;
  }

  const files = sourceFolder.getFiles();
  const fileList = [];
  let limitCounter = 0;
  const PROCESS_LIMIT = 50; 

  while (files.hasNext()) {
    if (limitCounter >= PROCESS_LIMIT) break;
    const f = files.next();
    if (f.getMimeType().startsWith('image/')) {
      try {
        f.moveTo(workFolder);
        fileList.push(f);
      } catch(e) {
        fileList.push(f.makeCopy(f.getName(), workFolder));
      }
      limitCounter++;
    }
  }
  if (fileList.length === 0) {
    const wFiles = workFolder.getFiles();
    while (wFiles.hasNext()) {
      if (limitCounter >= PROCESS_LIMIT) break;
      fileList.push(wFiles.next());
      limitCounter++;
    }
  }
  if (fileList.length === 0) { ui.alert('画像が見つかりません'); return; }

  // --- 3. シート準備 ---
  let sheet = ss.getSheetByName(SHEET_NAME_MATRIX);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(SHEET_NAME_MATRIX);

  // ヘッダー作成 (固定枠20列 + 候補枠50列)
  const candidateHeaders = Array.from({length: CANDIDATE_LIMIT}, (_, i) => `候補${i+1}`);
  
  // 1行目: エリア説明
  const header1 = ['商品情報', '', '', '', '', '▼アップロード対象 (ここに移動)', ...new Array(19).fill(''), '▼未整理・候補 (ここから左へドラッグ)', ...new Array(CANDIDATE_LIMIT-1).fill('')];
  
  // 2行目: 項目名
  const header2 = ['親SKU', '子SKU', '商品名', 'バリエーション値', '参考画像', ...RAKUTEN_IMAGE_COLS, ...candidateHeaders];
  const totalCols = header2.length;

  sheet.getRange(1, 1, 1, totalCols).setValues([header1]).setFontWeight('bold').setFontColor('white');
  sheet.getRange(1, 1, 1, 5).setBackground('#4c1130'); // 商品情報
  sheet.getRange(1, 6, 1, 20).setBackground('#0b5394'); // アップロード対象
  sheet.getRange(1, 26, 1, CANDIDATE_LIMIT).setBackground('#666666'); // 候補エリア

  sheet.getRange(2, 1, 1, totalCols).setValues([header2]).setBackground('#d9d9d9').setFontWeight('bold');
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(5);

  // --- 4. AI解析 & 振り分け ---
  const parents = Object.values(productGroups);
  const productListText = parents.map((p, i) => `${i}: ${p.name} (ID: ${p.pCode})`).join('\n');

  for (const file of fileList) {
    const fileName = file.getName();
    SpreadsheetApp.getActive().toast(`AI解析中... ${fileName}`);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}

    const imageUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;
    const imageFormula = `=IMAGE("${imageUrl}", 2)`; 

    const prompt = `
    Analyze this product image.
    Match it to one of the products below.
    [Product List]
    ${productListText}
    Task:
    1. Identify the matching product index.
    2. Does the image contain a NUMBER indicating quantity/set count? (e.g. "3" in "3袋", "10" in "10 sets").
    3. Is it a "Main" image (package/front) or "Sub" image (detail/cooking/back)?
    Output JSON: {"matched_index": 0, "quantity_number": 3, "type": "Main" or "Sub"}
    `;

    let assigned = false;
    try {
      const blob = file.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      const res = callGeminiVision(prompt, base64, blob.getContentType());
      const result = JSON.parse(res.replace(/```json|```/g, '').trim());

      if (result.matched_index !== -1 && parents[result.matched_index]) {
        const parent = parents[result.matched_index];
        let targetObj = parent; // デフォルトは親

        // 数字判定があれば子SKUへターゲットを変更
        if (result.quantity_number !== null && String(result.quantity_number) !== "" && result.type === 'Main') {
          const targetChild = parent.children.find(c => extractNumbers(c.varName).includes(String(result.quantity_number)));
          if (targetChild) targetObj = targetChild;
        }

        // 配置ロジック: Main/Sub判定を尊重しつつ、空いている青枠に積極的に入れる
        let placed = false;
        
        // 1. Main判定の場合: メイン枠 -> サブ枠(親) の順でトライ
        if (result.type === 'Main') {
           placed = pushToEmptySlot(targetObj.assignedMains, imageFormula);
           if (!placed) placed = pushToEmptySlot(parent.assignedSubs, imageFormula);
        } 
        // 2. Sub判定(または不明)の場合: サブ枠(親) -> メイン枠 の順でトライ
        else {
           placed = pushToEmptySlot(parent.assignedSubs, imageFormula);
           if (!placed) placed = pushToEmptySlot(targetObj.assignedMains, imageFormula);
        }

        // 3. それでも満杯なら候補枠へ
        if (!placed) {
           pushToEmptySlot(targetObj.assignedCandidates, imageFormula);
        }
        assigned = true;
      }
    } catch (e) { console.warn(`AI Error: ${e.message}`); }

    // マッチしない画像は先頭行の「青枠(メイン/サブ)」へ優先的に入れる
    if (!assigned && displayRows.length > 0) {
       const firstRow = displayRows[0];
       // メイン -> サブ -> 候補 の順で空きを探して入れる
       if (!pushToEmptySlot(firstRow.assignedMains, imageFormula)) {
          if (!pushToEmptySlot(firstRow.assignedSubs, imageFormula)) {
             pushToEmptySlot(firstRow.assignedCandidates, imageFormula);
          }
       }
    }
  }

  // --- 5. 出力 ---
  const outputRows = displayRows.map(row => {
    let refFunc = '';
    if (row.type === 'Parent' && row.refImg) refFunc = `=IMAGE("${row.refImg}", 2)`;
    else if (row.type === 'Child' && productGroups[row.pCode].refImg) refFunc = `=IMAGE("${productGroups[row.pCode].refImg}", 2)`;
    
    return [
      row.pCode, row.cCode, row.name, row.varName, refFunc, 
      ...row.assignedMains, 
      ...row.assignedSubs,
      ...row.assignedCandidates // 右端に追加
    ];
  });

  if (outputRows.length > 0) {
    const range = sheet.getRange(3, 1, outputRows.length, outputRows[0].length);
    range.setValues(outputRows);
    
    sheet.setRowHeights(3, outputRows.length, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidths(6, 20, 100); // メイン・サブエリア
    sheet.setColumnWidths(26, CANDIDATE_LIMIT, 60); // 候補エリアは狭く
    sheet.getRange(3, 1, outputRows.length, 4).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
    // 青枠でアップロード対象を囲む
    sheet.getRange(3, 6, outputRows.length, 20).setBorder(true, true, true, true, true, true, '#0000ff', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  ui.alert('AI仕分け完了。\n・青枠内：アップロード対象\n・右側のグレー枠：候補・未整理\n\n必要な画像を右側から青枠内へドラッグ移動してください。');
}


/**
 * 機能2: リネーム＆アップロード実行 (デバッグログ強化版)
 */
function executeRenameAndUploadFromMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrixSheet = ss.getSheetByName(SHEET_NAME_MATRIX);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  console.log("🚀 アップロード処理を開始します...");

  if (!matrixSheet) { ui.alert('【エラー】画像マッチングシートが見つかりません'); return; }

  let rcabinetId = '0';
  if (settingSheet) {
    const v = settingSheet.getRange("B9").getValue();
    if (v) rcabinetId = String(v).trim();
  }
  console.log(`R-CabinetフォルダID: ${rcabinetId}`);

  if (ui.alert('実行確認', '画像をアップロードし、マスタへURLを書き込みます。\nよろしいですか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // マスタの列特定
  const mValues = masterSheet.getDataRange().getValues();
  const mHeaderIdx = getAnchorRowIndex(mValues);
  if (mHeaderIdx === -1) { ui.alert("マスタのヘッダーが見つかりません"); return; }
  
  const mHeaders = mValues[mHeaderIdx];
  const mColMap = getColumnIndexMap(mHeaders);
  const mPCodeIdx = mColMap[COL_NAME_PARENT_SKU];
  const mCCodeIdx = mColMap[COL_NAME_CHILD_SKU];

  // 行マップ作成
  const rowMap = {};
  for(let i=mHeaderIdx+1; i<mValues.length; i++){
    const p = String(mValues[i][mPCodeIdx]).trim();
    const c = (mCCodeIdx!==undefined) ? String(mValues[i][mCCodeIdx]).trim() : "";
    if(p) {
      const key = c ? `${p}_${c}` : p;
      rowMap[key] = i + 1;
    }
  }

  // 画像シート読み込み
  const lastRow = matrixSheet.getLastRow();
  if (lastRow < 3) { ui.alert("画像シートにデータがありません"); return; }

  // A列～Y列(25列分)を取得
  const formulas = matrixSheet.getRange(3, 1, lastRow - 2, 25).getFormulas(); 
  const values = matrixSheet.getRange(3, 1, lastRow - 2, 25).getValues();

  console.log(`対象行数: ${formulas.length} 行`);

  const archiveFolder = DriveApp.getFolderById(DRIVE_IMAGE_ARCHIVE_FOLDER_ID);
  let successCount = 0;
  let errorCount = 0;
  let skipCount = 0;
  const processedFileIds = new Set(); 
  
  // ★追加: 重複アップロード回避用のURLキャッシュ
  const uploadedUrlMap = new Map();

  // 行ループ
  for (let r = 0; r < formulas.length; r++) {
    const pCode = String(values[r][0]).trim();
    const cCode = String(values[r][1]).trim();
    if (!pCode) {
      console.log(`行${r+3}: 親SKUがないためスキップ`);
      continue;
    }

    const isChild = (cCode !== "");
    const rowKey = isChild ? `${pCode}_${cCode}` : pCode;
    
    // 画像列ループ (F列=Index5 ～ Y列=Index24)
    for (let c = 5; c < 25; c++) {
      const cellFormula = formulas[r][c];
      
      // 画像判定
      if (!cellFormula || !cellFormula.includes('IMAGE')) {
        continue; // 画像がないセルは無視
      }

      const match = cellFormula.match(/id=([a-zA-Z0-9_-]+)/);
      if (!match) {
        console.warn(`行${r+3}列${c+1}: 数式はあるがID抽出に失敗 -> ${cellFormula}`);
        skipCount++;
        continue;
      }
      
      const fileId = match[1];
      const slotIndex = c - 5; // 0=メイン1, 1=メイン2..., 10=サブ1...

      console.log(`>> 処理中: 行${r+3} / 画像枠${slotIndex} (FileID: ${fileId})`);

      // ファイル名・列名決定
      let targetFileName = "";
      let targetColName = ""; 

      if (slotIndex < 10) {
        // メイン画像 (0~9)
        const suffix = slotIndex; 
        if (isChild) targetFileName = (suffix === 0) ? `${cCode}.jpg` : `${cCode}_${suffix}.jpg`;
        else targetFileName = (suffix === 0) ? `${pCode}.jpg` : `${pCode}_${suffix}.jpg`;
        targetColName = `楽天メイン画像${slotIndex + 1}`; 
      } else {
        // サブ画像 (10~19)
        const subIdx = slotIndex - 9; // 1~10
        targetFileName = `${pCode}_sub${subIdx}.jpg`; 
        targetColName = `楽天サブ画像${subIdx}`;
      }

      // API実行
      try {
        const file = DriveApp.getFileById(fileId);
        
        // ★追加: 重複回避ロジック
        let imageUrl = "";
        let isNewUpload = false; // 新規アップロードかどうか

        if (uploadedUrlMap.has(fileId)) {
           // A. 既に同じ画像をアップロード済みの場合 -> 使い回す
           imageUrl = uploadedUrlMap.get(fileId);
           console.log(`   ♻️ キャッシュ利用(アップロード省略): ${imageUrl}`);
        } else {
           // B. 初めての画像の場合 -> 新規アップロード
           file.setName(targetFileName); 
           const res = callRCabinetInsertFile(file.getBlob(), targetFileName, rcabinetId);
           
           if (!res.success) {
             // 失敗時
             console.error(`   ❌ アップロード失敗: ${res.message}`);
             matrixSheet.getRange(r + 3, c + 1).setBackground('#f4cccc'); // 赤
             matrixSheet.getRange(r + 3, c + 1).setNote(res.message);
             errorCount++;
             continue; // 次の画像の処理へ
           }
           
           // 成功時
           imageUrl = res.imageUrl;
           uploadedUrlMap.set(fileId, imageUrl); // キャッシュに登録
           isNewUpload = true;
           console.log(`   ✅ 新規アップロード成功: ${imageUrl}`);
        }

        // --- 以下、マスタへの書き込み処理 (共通) ---
        
        // 1. 本来の列への書き込み
        const targetColIdx = mColMap[targetColName];
        const writeRow = rowMap[rowKey];
        
        if (writeRow && targetColIdx !== undefined) {
          masterSheet.getRange(writeRow, targetColIdx + 1).setValue(imageUrl);
          console.log(`      -> マスタ(${writeRow}行目, ${targetColName})に保存`);
        } else {
          console.warn(`      -> ⚠️ マスタに行または列が見つかりません (Row:${writeRow}, Col:${targetColName})`);
        }

        // 2. 白背景画像への自動コピー
        if (targetColName === '楽天メイン画像1') {
           const whiteBgIdx = mColMap['楽天白背景画像'];
           if (writeRow && whiteBgIdx !== undefined) {
              masterSheet.getRange(writeRow, whiteBgIdx + 1).setValue(imageUrl);
              console.log(`      -> 白背景列にも保存`);
           }
        }
        
        // 3. アーカイブ移動 (新規アップロード時のみ)
        if (isNewUpload && !processedFileIds.has(fileId)) {
           try { file.moveTo(archiveFolder); } catch(e){}
           processedFileIds.add(fileId);
        }
        
        matrixSheet.getRange(r + 3, c + 1).setBackground('#d9ead3'); // 緑
        successCount++;
        
        // 連続アクセス制限回避 (新規アップロード時のみ待機)
        if (isNewUpload) Utilities.sleep(200);

      } catch (e) {
        console.error(`   ❌ システムエラー: ${e.message}`);
        matrixSheet.getRange(r + 3, c + 1).setBackground('#f4cccc');
        errorCount++;
      }
    }
  }
  
  const resultMsg = `処理完了\n成功: ${successCount} 件\n失敗: ${errorCount} 件\nスキップ(ID取得不可など): ${skipCount} 件`;
  console.log(resultMsg);
  ui.alert(resultMsg);
}

// ------------------------------------------
// 共通関数
// ------------------------------------------
function pushToEmptySlot(array, val) {
  for (let k = 0; k < array.length; k++) {
    if (array[k] === '') {
      array[k] = val;
      return true;
    }
  }
  return false;
}
function extractNumbers(str) {
  // 文字列から数字のみを配列で抽出 ("3袋" -> ["3"])
  const matches = str.match(/\d+/g);
  return matches ? matches : [];
}
function getAnchorRowIndex(values) {
  const targetKey1 = COL_NAME_PARENT_SKU; 
  const targetKey2 = COL_NAME_CHECK;      
  for (let r = 0; r < 20; r++) { 
    if (!values[r]) continue;
    let has1 = false, has2 = false;
    for (let c = 0; c < values[r].length; c++) {
      const v = String(values[r][c]).trim();
      if(v === targetKey1) has1 = true;
      if(v === targetKey2) has2 = true;
    }
    if (has1 && has2) return r; 
  }
  return -1; 
}
function getColumnIndexMap(headers) {
  const m = {};
  headers.forEach((v, i) => m[String(v).trim()] = i);
  return m;
}
function callGeminiVision(prompt, base64, mime) {
  const url = `https://generativelanguage.googleapis.com/v1beta/${GEMINI_VISION_MODEL}:generateContent?key=${getGeminiApiKey()}`;
  const payload = { contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mime, data: base64 } }] }] };
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  return res.getContentText();
}



/**
 * 楽天R-Cabinetへの画像アップロード関数 (完全版)
 * 1. マルチパート形式でアップロード
 * 2. 返却されたFileIdを使って検索APIを叩く
 * 3. 画像URL(FileUrl)を取得して返す
 */

function callRCabinetInsertFile(blob, name, folderId) {
  const authKey = Utilities.base64Encode(`${getRakutenServiceSecret()}:${getRakutenLicenseKey()}`);
  const headers = { 'Authorization': `ESA ${authKey}` };

  // Step 1. 画像のアップロード
  const uploadUrl = 'https://api.rms.rakuten.co.jp/es/1.0/cabinet/file/insert';
  
  const uploadXml = `<?xml version="1.0" encoding="UTF-8"?>
<request>
  <fileInsertRequest>
    <file>
      <fileName>${name}</fileName>
      <folderId>${folderId}</folderId>
      <overWrite>true</overWrite>
    </file>
  </fileInsertRequest>
</request>`;

  const uploadPayload = { "xml": uploadXml, "file": blob };
  let fileId = "";

  try {
    const res = UrlFetchApp.fetch(uploadUrl, {
      method: 'post',
      headers: headers,
      payload: uploadPayload,
      muteHttpExceptions: true
    });
    
    const text = res.getContentText();
    const code = res.getResponseCode();

    if (code === 200 && text.includes('<resultCode>0</resultCode>')) {
      const match = text.match(/<FileId>(\d+)<\/FileId>/);
      if (match) fileId = match[1];
      else return { success: false, message: "Upload OK but FileId not found" };
    } else {
      const errMatch = text.match(/<message>(.*?)<\/message>/) || text.match(/<faultstring>(.*?)<\/faultstring>/);
      return { success: false, message: errMatch ? errMatch[1] : text };
    }

  } catch (e) {
    return { success: false, message: "Upload Error: " + e.message };
  }

  Utilities.sleep(500);

  // Step 2. URLの取得
  const searchUrl = `https://api.rms.rakuten.co.jp/es/1.0/cabinet/files/search?fileId=${fileId}`;

  try {
    const res = UrlFetchApp.fetch(searchUrl, {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    });

    const text = res.getContentText();
    const match = text.match(/<FileUrl>(.*?)<\/FileUrl>/);

    if (match) {
      let url = match[1];
      
      // ★修正箇所: /cabinet/ は消すが、先頭のスラッシュは残すため「+8」にする
      // 例: https://.../cabinet/123/img.jpg -> /123/img.jpg
      if (url.includes('/cabinet/')) {
        url = url.substring(url.indexOf('/cabinet/') + 8); 
      }
      
      return { success: true, imageUrl: url };
    } else {
      return { success: false, message: "Search OK but FileUrl not found" };
    }

  } catch (e) {
    return { success: false, message: "Search Error: " + e.message };
  }
}






// ==============================================================================================================================






function debugPriceCalculation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('▼商品マスタ(人間作業用)');
  const settingSheet = ss.getSheetByName('▼設定(マッピング)');
  
  console.log("★価格計算ロジックの精密診断");

  // --- 1. 設定シートの確認 ---
  const setValues = settingSheet.getDataRange().getValues();
  const sysKeys = setValues[1]; // 2行目(システムキー)
  const mapNames = setValues[2]; // 3行目(マスタ引用元)
  const configs = setValues[3]; // 4行目(係数)

  // 調査対象のキー
  const targetKeys = [
    { key: 'price', label: '通常購入販売価格' },
    { key: 'sub_price', label: '定期購入販売価格' },
    { key: 'sub_first_price', label: '定期用初回価格' }
  ];

  const rules = {};

  console.log("\n【設定シートのルール確認】");
  targetKeys.forEach(t => {
    const idx = sysKeys.indexOf(t.key);
    if (idx === -1) {
      console.error(`❌ 設定シートに ${t.key} が見つかりません`);
      return;
    }
    rules[t.key] = {
      mapName: mapNames[idx],
      config: configs[idx]
    };
    console.log(`■ ${t.label} (${t.key})`);
    console.log(`   引用元列名: [${rules[t.key].mapName}]`);
    console.log(`   計算係数  : [${rules[t.key].config}]`);
  });

  // --- 2. マスタデータの特定 ---
  const mValues = masterSheet.getDataRange().getValues();
  let hRow = -1;
  for(let i=0; i<20; i++){ if(mValues[i].includes('親SKU')){ hRow = i; break; } }
  
  const mHeaders = mValues[hRow];
  const colMap = {};
  mHeaders.forEach((h, i) => colMap[String(h).trim()] = i);

  // ターゲットSKU (66s6) を含む行を探す
  let targetRow = null;
  const cSkuIdx = colMap['子SKU'];
  
  for (let i = hRow + 1; i < mValues.length; i++) {
    const cSku = String(mValues[i][cSkuIdx]);
    if (cSku.includes('66s6')) { // 部分一致で探す
      targetRow = mValues[i];
      console.log(`\n【診断対象行を発見】: 行${i+1} (子SKU: ${cSku})`);
      break;
    }
  }

  if (!targetRow) { console.error("❌ 診断対象のSKU (66s6) がマスタに見つかりません"); return; }

  // --- 3. 計算シミュレーション ---
  console.log(`\n【計算結果の検証】`);

  targetKeys.forEach(t => {
    const rule = rules[t.key];
    if (!rule) return;

    // A. マスタからの生の値
    const colIdx = colMap[rule.mapName];
    let rawValue = "";
    if (colIdx !== undefined) {
      rawValue = targetRow[colIdx];
    } else {
      console.error(`   ❌ マスタに列「${rule.mapName}」が存在しません`);
    }

    console.log(`\n■ ${t.label}`);
    console.log(`   参照した列: ${rule.mapName}`);
    console.log(`   マスタの値: ${rawValue} (円)`);
    console.log(`   係数(設定): ${rule.config}`);

    // B. 計算実行
    let calculated = rawValue;
    if (rule.config && !isNaN(parseFloat(rule.config))) {
      // 数値化して計算
      const numVal = parseFloat(String(rawValue).replace(/,/g, ''));
      const numConf = parseFloat(String(rule.config));
      
      if (!isNaN(numVal)) {
        calculated = Math.floor(numVal * numConf);
        console.log(`   計算式    : ${numVal} × ${numConf} = ${numVal * numConf}`);
        console.log(`   最終結果  : ${calculated} (円)`);
      } else {
        console.warn(`   ⚠️ 計算不可: マスタの値が数値ではありません`);
      }
    } else {
      console.log(`   (計算なし、そのまま出力)`);
    }
  });
}

// ============================================================
// MAKE Webhook テスト関数（楽天FTPアップロード連携用）
// ============================================================

/**
 * MAKEのWebhookにテストデータを送信する関数（Google Drive経由）
 * MAKEで「Redetermine data structure」を実行して待機状態にしてから実行してください
 */
function testMakeWebhookDrive() {
  const webhookUrl = "https://hook.eu1.make.com/sbxvjnqvn7b23l673crdgkb7j1s4sl4w";
  
  // 楽天の正しいファイル名（normal-item.csv）
  const testFileName = "normal-item.csv";
  
  // テスト用CSVデータ（楽天フォーマット準拠）
  const csvRows = [
    ["コントロールカラム", "商品管理番号（商品URL）", "商品番号", "全商品ディレクトリID", "タグID", "商品名", "販売価格"],
    ["n", "test-item-001", "TEST001", "12345", "", "テスト商品", "1000"]
  ];
  
  // CSV文字列を生成（カンマ区切り、CRLF改行）
  const testCsvData = csvRows.map(row => row.join(",")).join("\r\n");
  
  // テスト用ファイルをDriveに作成（Shift-JIS）
  const saveFolder = DriveApp.getFolderById(CSV_SAVE_FOLDER_ID);
  
  // 既存の同名ファイルを削除
  const existingFiles = saveFolder.getFilesByName(testFileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }
  
  // Shift-JISでエンコードしてファイル作成
  const blob = Utilities.newBlob("", "text/csv", testFileName).setDataFromString(testCsvData, "Shift_JIS");
  const newFile = saveFolder.createFile(blob);
  const fileId = newFile.getId();
  
  console.log("=== MAKE Webhook テスト開始 (Google Drive経由) ===");
  console.log("テストファイルを作成しました:");
  console.log("  ファイル名: " + testFileName);
  console.log("  ファイルID: " + fileId);
  console.log("  CSV内容:");
  console.log(testCsvData);
  console.log("  バイト数: " + blob.getBytes().length);
  
  const payload = {
    fileId: fileId,
    fileName: "normal-item.csv",
    source: "googleDrive"
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  console.log("\n送信先URL: " + webhookUrl);
  console.log("ペイロード: " + JSON.stringify(payload, null, 2));
  
  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log("\n--- レスポンス ---");
    console.log("ステータスコード: " + responseCode);
    console.log("レスポンス内容: " + responseText);
    
    if (responseCode === 200) {
      console.log("\n✅ テスト成功！MAKEがデータを受信しました。");
      console.log("MAKEでGoogle Drive接続を設定し、このファイルIDを使ってダウンロードするように設定してください。");
    } else {
      console.log("\n⚠️ 予期しないレスポンスコード: " + responseCode);
    }
  } catch (e) {
    console.error("❌ エラー発生: " + e.message);
  }
  
  console.log("\n=== テスト完了 ===");
}

// ============================================================
// 楽天FTP 反映確認機能
// ============================================================

const RAKUTEN_SHOP_ID = "octas"; // 楽天店舗ID（必要に応じて設定シートから取得）

/** テスト用：反映確認・完了メールのログをシートに追記（後から確認できる） */
function logRakutenReflection(action, detail) {
  const detailStr = String(detail || '').substring(0, 500);
  try {
    const id = PropertiesService.getScriptProperties().getProperty('RAKUTEN_REFLECTION_LOG_SS_ID');
    let ss = null;
    if (id) {
      try { ss = SpreadsheetApp.openById(id); } catch (_) {}
    }
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.warn('楽天反映ログ: スプレッドシートを特定できません。RAKUTEN_REFLECTION_LOG_SS_ID=' + (id || '未設定'));
      return;
    }
    let sheet = ss.getSheetByName(RAKUTEN_REFLECTION_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(RAKUTEN_REFLECTION_LOG_SHEET);
      sheet.getRange(1, 1, 1, 4).setValues([['日時', 'アクション', '詳細', '']]);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#eee');
    }
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const nextRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(nextRow, 1, 1, 3).setValues([[now, String(action), detailStr]]);
  } catch (e) {
    console.warn('楽天反映ログ記録失敗: action=' + action + ', err=' + e.message);
  }
}

const REFLECTION_CHECK_INTERVAL_MIN = 5; // チェック間隔（分）
const REFLECTION_CHECK_MAX_DURATION_MIN = 30; // 最大チェック時間（分）

/**
 * 楽天商品ページの反映確認を開始する（5分おきにURLへアクセスして200なら反映済み。完了時にメール送信）
 * @param {Array<{code:string, name:string}>} products - 商品一覧（商品管理番号・商品名）
 * @param {string} shopId - 楽天店舗ID（商品URL・完了メール用）
 */
function startReflectionCheck(products, shopId) {
  if (!products || products.length === 0) {
    console.log("反映確認: 対象商品がありません");
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  
  const checkData = {
    startTime: startTime,
    shopId: shopId || "",
    items: products.map(p => ({ url: p.code, name: p.name || "", reflected: false }))
  };
  props.setProperty('REFLECTION_CHECK_DATA', JSON.stringify(checkData));
  // トリガー実行時にログを書くため、スプレッドシートIDを保存
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) props.setProperty('RAKUTEN_REFLECTION_LOG_SS_ID', ss.getId());
  } catch (_) {}

  console.log(`=== 反映確認を開始 ===`);
  console.log(`対象商品数: ${products.length}`);
  logRakutenReflection('反映確認開始', `商品数: ${products.length}、間隔: ${REFLECTION_CHECK_INTERVAL_MIN}分、最大: ${REFLECTION_CHECK_MAX_DURATION_MIN}分`);
  
  deleteReflectionCheckTriggers();
  try {
    ScriptApp.newTrigger('runReflectionCheck')
      .timeBased()
      .everyMinutes(REFLECTION_CHECK_INTERVAL_MIN)
      .create();
  } catch (e) {
    // 権限不足で自動作成できない場合。手動トリガーが1本あれば反映確認は動くため、シートには書かず console のみ
    console.warn('反映確認トリガー自動作成をスキップ: ' + e.message);
  }
  
  SpreadsheetApp.getActive().toast(
    `${products.length}件の反映確認を開始しました（${REFLECTION_CHECK_INTERVAL_MIN}分おき）。完了時にメール送信します。`,
    '楽天FTP'
  );
}

/**
 * 反映確認を実行（トリガーから呼び出される）
 */
function runReflectionCheck() {
  const props = PropertiesService.getScriptProperties();
  const checkDataStr = props.getProperty('REFLECTION_CHECK_DATA');
  
  if (!checkDataStr) {
    console.log("反映確認: データがありません。トリガーを削除します。");
    deleteReflectionCheckTriggers();
    return;
  }
  
  const checkData = JSON.parse(checkDataStr);
  const now = new Date().getTime();
  const elapsedMin = (now - checkData.startTime) / (1000 * 60);
  
  console.log(`=== 反映確認チェック (経過: ${Math.round(elapsedMin)}分) ===`);
  logRakutenReflection('チェック実行', `経過: ${Math.round(elapsedMin)}分`);
  
  // 最大時間を超えた場合は終了
  if (elapsedMin >= REFLECTION_CHECK_MAX_DURATION_MIN) {
    console.log("最大待機時間を超過。反映確認を終了します。");
    finishReflectionCheck(checkData, "timeout");
    return;
  }
  
  // 店舗IDを取得（設定シートから）
  let shopId = RAKUTEN_SHOP_ID;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
    if (settingSheet) {
      const settingRaw = settingSheet.getDataRange().getValues();
      if (settingRaw.length > 10 && settingRaw[10][1]) {
        shopId = String(settingRaw[10][1]).trim();
      }
    }
  } catch (e) {
    console.log("設定シートから店舗ID取得失敗、デフォルト使用: " + shopId);
  }
  
  let allReflected = true;
  let reflectedCount = 0;
  
  // 各商品をチェック
  checkData.items.forEach((item, index) => {
    if (item.reflected) {
      reflectedCount++;
      return; // 既に反映済み
    }
    
    const productUrl = `https://item.rakuten.co.jp/${shopId}/${item.url}/`;
    
    try {
      const response = UrlFetchApp.fetch(productUrl, {
        muteHttpExceptions: true,
        followRedirects: false
      });
      const statusCode = response.getResponseCode();
      
      if (statusCode === 200) {
        console.log(`✅ 反映確認: ${item.url} (HTTP ${statusCode})`);
        item.reflected = true;
        reflectedCount++;
        updateMasterSheetFtpStatus(item.url, "反映済");
        logRakutenReflection('URL確認OK', `${item.url} → 200`);
      } else {
        console.log(`⏳ 未反映: ${item.url} (HTTP ${statusCode})`);
        logRakutenReflection('URL未反映', `${item.url} → ${statusCode}`);
        allReflected = false;
      }
    } catch (e) {
      console.log(`❌ チェックエラー: ${item.url} - ${e.message}`);
      logRakutenReflection('URLチェックエラー', `${item.url} - ${e.message}`);
      allReflected = false;
    }
    
    // API制限対策: 少し待機
    Utilities.sleep(500);
  });
  
  console.log(`反映状況: ${reflectedCount}/${checkData.items.length} 件`);
  logRakutenReflection('反映状況', `${reflectedCount}/${checkData.items.length} 件`);
  
  // 全て反映済みなら終了
  if (allReflected) {
    console.log("全商品の反映を確認。反映確認を終了します。");
    finishReflectionCheck(checkData, "complete");
    return;
  }
  
  // データを更新して保存
  props.setProperty('REFLECTION_CHECK_DATA', JSON.stringify(checkData));
}

/**
 * マスタシートの楽天FTPステータスを更新
 * @param {string} itemUrl - 商品管理番号
 * @param {string} status - ステータス（送信済/反映済/Error）
 */
function updateMasterSheetFtpStatus(itemUrl, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return;
    
    const data = masterSheet.getDataRange().getValues();
    const headers = data[0];
    
    // 楽天FTPステータス列を探す
    let ftpStatusCol = headers.indexOf('楽天FTP');
    if (ftpStatusCol === -1) {
      ftpStatusCol = headers.indexOf('楽天FTPステータス');
    }
    
    // 商品管理番号列を探す（複数の候補）
    let itemUrlCol = headers.indexOf('商品管理番号（商品URL）');
    if (itemUrlCol === -1) itemUrlCol = headers.indexOf('商品URL');
    if (itemUrlCol === -1) itemUrlCol = headers.indexOf('item_url');
    if (itemUrlCol === -1) itemUrlCol = headers.indexOf('親SKU');
    
    if (ftpStatusCol === -1 || itemUrlCol === -1) {
      console.log(`ステータス更新スキップ: 列が見つかりません (FTP列:${ftpStatusCol}, URL列:${itemUrlCol})`);
      return;
    }
    
    // 該当行を探して更新
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][itemUrlCol]).trim() === String(itemUrl).trim()) {
        masterSheet.getRange(i + 1, ftpStatusCol + 1).setValue(status);
        console.log(`マスタ更新: 行${i + 1}, ${itemUrl} -> ${status}`);
        break;
      }
    }
  } catch (e) {
    console.error("マスタシート更新エラー: " + e.message);
  }
}

/**
 * 反映確認を終了する。トリガーを先に削除して実行を「完了」にし、メール送信後に即終了する。
 * @param {Object} checkData - チェックデータ
 * @param {string} reason - 終了理由（complete/timeout）
 */
function finishReflectionCheck(checkData, reason) {
  deleteReflectionCheckTriggers();

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('REFLECTION_CHECK_DATA');

  const reflectedCount = checkData.items.filter(i => i.reflected).length;
  const totalCount = checkData.items.length;

  let message = "";
  if (reason === "complete") {
    message = `✅ 全${totalCount}件の商品が楽天に反映されました`;
  } else {
    message = `⏱️ タイムアウト: ${reflectedCount}/${totalCount}件が反映済み`;
  }

  console.log(message);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, '楽天FTP反映確認', 8);
  } catch (_) {
    try {
      var id = props.getProperty('RAKUTEN_REFLECTION_LOG_SS_ID');
      if (id) SpreadsheetApp.openById(id).toast(message, '楽天FTP反映確認', 8);
    } catch (_) {}
  }
  logRakutenReflection(reason === 'complete' ? '反映確認完了' : '反映確認タイムアウト', `${reflectedCount}/${totalCount} 件`);

  if (reason === "timeout") {
    checkData.items.forEach(item => {
      if (!item.reflected) {
        updateMasterSheetFtpStatus(item.url, "未反映");
      }
    });
  }

  // 反映確認の結果を完了メールで送信（Yahoo同様の内容：一覧・商品ページURL・削除用URL・削除案内）
  sendRakutenCompletionEmailWithReflectionResult(checkData, reason);
}

/**
 * 反映確認用のトリガーを削除
 * トリガー実行時は script.scriptapp 権限がないため失敗する。その場合は try-catch で握りつぶし処理は継続する。
 */
function deleteReflectionCheckTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runReflectionCheck') {
        ScriptApp.deleteTrigger(trigger);
        console.log("反映確認トリガーを削除しました");
      }
    });
  } catch (e) {
    console.warn("反映確認トリガー削除をスキップ（権限不足。トリガー実行時は発生します）: " + e.message);
  }
}

/**
 * 反映確認を手動で停止する
 */
function stopReflectionCheck() {
  deleteReflectionCheckTriggers();
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('REFLECTION_CHECK_DATA');
  SpreadsheetApp.getActive().toast('反映確認を停止しました', '楽天FTP');
  console.log("反映確認を手動で停止しました");
}

/**
 * Google DriveのファイルをMAKE経由で楽天FTPにアップロードする
 * @param {string} fileId - Google DriveのファイルID
 * @param {string} fileName - アップロードするファイル名
 * @returns {boolean} - 成功時true、失敗時false
 */
function sendFileToRakutenFtpViaDrive(fileId, fileName) {
  const MAKE_WEBHOOK_URL = "https://hook.eu1.make.com/sbxvjnqvn7b23l673crdgkb7j1s4sl4w";
  
  const payload = {
    fileId: fileId,
    fileName: fileName,
    source: "googleDrive"  // MAKEにデータソースを伝える
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  console.log("=== 楽天FTPアップロード開始 (Google Drive経由) ===");
  console.log("ファイルID: " + fileId);
  console.log("ファイル名: " + fileName);
  
  try {
    const response = UrlFetchApp.fetch(MAKE_WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      console.log("✅ MAKEへの送信成功 - 楽天FTPにアップロードされます");
      SpreadsheetApp.getActive().toast('楽天FTPへアップロード中...', 'MAKE連携');
      return true;
    } else {
      console.error("⚠️ MAKEへの送信失敗 - ステータスコード: " + responseCode);
      console.error("レスポンス: " + response.getContentText());
      return false;
    }
  } catch (e) {
    console.error("❌ FTPアップロードエラー: " + e.message);
    return false;
  }
}

/**
 * CSVデータをMAKE経由で楽天FTPにアップロードする（旧方式・Base64）
 * @param {string} csvData - CSV形式の文字列データまたはBase64エンコード済みデータ
 * @param {string} fileName - アップロードするファイル名
 * @param {boolean} enableReflectionCheck - 反映確認を有効にするか（デフォルト: true）
 * @param {boolean} isBase64 - データがBase64エンコード済みか（デフォルト: false）
 * @returns {boolean} - 成功時true、失敗時false
 */
function sendCsvToRakutenFtp(csvData, fileName, enableReflectionCheck = true, isBase64 = false) {
  const MAKE_WEBHOOK_URL = "https://hook.eu1.make.com/sbxvjnqvn7b23l673crdgkb7j1s4sl4w";
  
  const payload = {
    csvData: csvData,
    fileName: fileName,
    isBase64: isBase64  // MAKEにBase64かどうかを伝える
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  console.log("=== 楽天FTPアップロード開始 ===");
  console.log("ファイル名: " + fileName);
  console.log("データサイズ: " + csvData.length + " 文字");
  console.log("エンコーディング: " + (isBase64 ? "Shift-JIS (Base64)" : "UTF-8"));
  
  try {
    const response = UrlFetchApp.fetch(MAKE_WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      console.log("✅ MAKEへの送信成功 - 楽天FTPにアップロードされます");
      SpreadsheetApp.getActive().toast('楽天FTPへアップロード中...', 'MAKE連携');
      
      // 反映確認を開始（Base64の場合はデコードしてからURL抽出）
      if (enableReflectionCheck) {
        let csvForParsing = csvData;
        if (isBase64) {
          // Base64デコードしてShift-JISからUTF-8に変換
          try {
            const decoded = Utilities.newBlob(Utilities.base64Decode(csvData)).getDataAsString('Shift_JIS');
            csvForParsing = decoded;
          } catch (e) {
            console.log("Base64デコード失敗、反映確認をスキップ: " + e.message);
            csvForParsing = null;
          }
        }
        
        if (csvForParsing) {
          const itemUrls = extractItemUrlsFromCsv(csvForParsing);
          if (itemUrls.length > 0) {
            const products = itemUrls.map(url => ({ code: url, name: "" }));
            let shopId = "";
            try {
              const ss = SpreadsheetApp.getActiveSpreadsheet();
              const settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
              if (settingSheet) {
                const raw = settingSheet.getDataRange().getValues();
                if (raw.length > 10 && raw[10][1]) shopId = String(raw[10][1]).trim();
              }
            } catch (e) {}
            startReflectionCheck(products, shopId);
          }
        }
      }
      
      return true;
    } else {
      console.error("⚠️ MAKEへの送信失敗 - ステータスコード: " + responseCode);
      console.error("レスポンス: " + response.getContentText());
      return false;
    }
  } catch (e) {
    console.error("❌ FTPアップロードエラー: " + e.message);
    return false;
  }
}

/**
 * CSVデータから商品管理番号（item_url）を抽出する
 * @param {string} csvData - CSV形式の文字列
 * @returns {string[]} - 商品管理番号の配列（重複なし）
 */
function extractItemUrlsFromCsv(csvData) {
  const lines = csvData.split(/\r\n|\n/);
  if (lines.length < 2) return [];
  
  // ヘッダー行から item_url（商品管理番号）の列インデックスを探す
  const headers = parseCSVLine(lines[0]);
  let itemUrlColIdx = headers.findIndex(h => 
    h === '商品管理番号（商品URL）' || 
    h === 'item_url' || 
    h === '商品URL' ||
    h === '商品管理番号'
  );
  
  if (itemUrlColIdx === -1) {
    console.log("商品管理番号列が見つかりません");
    return [];
  }
  
  // データ行から商品管理番号を抽出（重複除去）
  const itemUrls = new Set();
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const row = parseCSVLine(lines[i]);
    const itemUrl = row[itemUrlColIdx];
    if (itemUrl && itemUrl.trim()) {
      itemUrls.add(itemUrl.trim());
    }
  }
  
  const result = Array.from(itemUrls);
  console.log(`抽出した商品管理番号: ${result.length}件`);
  return result;
}

/**
 * CSV行をパースする（カンマ区切り、ダブルクォート対応）
 * @param {string} line - CSV行
 * @returns {string[]} - フィールドの配列
 */
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (inQuotes) {
      if (char === '"') {
        if (line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === ',') {
        result.push(current);
        current = '';
      } else {
        current += char;
      }
    }
  }
  result.push(current);
  
  return result;
}