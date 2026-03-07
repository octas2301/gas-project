// ==========================================
// 在庫管理アプリ - スプレッドシート初期化（初回1回実行）
// ==========================================

/**
 * メニューに「スプレッドシートを初期化」を追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('在庫管理アプリ')
    .addItem('スプレッドシートを初期化（シート作成・初期データ）', 'initializeSpreadsheet')
    .addToUi();
}

/**
 * 指定したスプレッドシートに全シートを作成し、見出しと初期データを投入する。
 * 既存シートは上書きしない（存在する場合はスキップ）。
 */
function initializeSpreadsheet() {
  var ss = getSpreadsheet();
  var sheetNames = [
    'アクセス許可',
    '設定',
    '日報宛先',
    '商品マスタ',
    '場所マスタ',
    '理由マスタ',
    '担当者マスタ',
    '棚卸セッション',
    '棚卸スキャンリスト',
    '入荷リスト',
    '出荷リスト',
    '在庫移動リスト',
    '操作ログ',
    'パフォーマンスログ'
  ];

  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    // 見出しと初期データはシートごとに設定
    setupSheet(sheet, name);
  }

  SpreadsheetApp.getUi().alert('初期化が完了しました。');
}

function setupSheet(sheet, name) {
  sheet.clear();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 1) return; // 既にデータがある場合はスキップ

  switch (name) {
    case 'アクセス許可':
      sheet.getRange(1, 1, 1, 2).setValues([['メールアドレス', '備考']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      break;
    case '設定':
      sheet.getRange(1, 1, 1, 2).setValues([['キー', '値']]);
      sheet.getRange(2, 1).setValue('棚卸アラート月数');
      sheet.getRange(2, 2).setValue(3);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      break;
    case '日報宛先':
      sheet.getRange(1, 1).setValues([['メールアドレス']]);
      sheet.getRange(1, 1).setFontWeight('bold');
      break;
    case '商品マスタ':
      sheet.getRange(1, 1, 1, 6).setValues([['単品JAN', '箱JAN', '商品名', '商品原価', '平均原価', '箱入数']]);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      break;
    case '場所マスタ':
      sheet.getRange(1, 1).setValues([['場所名']]);
      sheet.getRange(1, 1).setFontWeight('bold');
      sheet.getRange(2, 1, 7, 1).setValues([
        ['オフィス1階（自宅）'], ['オフィス2階（自宅）'], ['ガレージ（車庫）'],
        ['倉庫（BENALO賃貸）'], ['ナーシングサポート前後'], ['返品'], ['その他（ここに直接入力）']
      ]);
      break;
    case '理由マスタ':
      sheet.getRange(1, 1).setValues([['理由名']]);
      sheet.getRange(1, 1).setFontWeight('bold');
      sheet.getRange(2, 1, 6, 1).setValues([
        ['仕入（通常の発注・購入）'], ['返品（お客様からの返品）'], ['棚卸調整増（棚卸で見つかった在庫）'],
        ['サンプル受入（無償提供品など）'], ['初期在庫（システム稼働時の開始残高）'], ['その他']
      ]);
      break;
    case '担当者マスタ':
      sheet.getRange(1, 1).setValues([['担当者名']]);
      sheet.getRange(1, 1).setFontWeight('bold');
      sheet.getRange(2, 1, 13, 1).setValues([
        ['Octas渡邊'], ['Octas古畑'], ['Octas矢野'],
        ['ナーシングA'], ['ナーシングB'], ['ナーシングC'], ['ナーシングD'], ['ナーシングE'],
        ['Octas-TA'], ['Octas-CO'], ['Octas-NA'], ['Octas-HA'], ['Octas-MO']
      ]);
      break;
    case '棚卸セッション':
      sheet.getRange(1, 1, 1, 3).setValues([['セッションID', '開始日時', '完了日時']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      break;
    case '棚卸スキャンリスト':
      sheet.getRange(1, 1, 1, 16).setValues([[
        'ID', '日時', '単品JAN', '商品名', '原価', '箱入数', '箱数', '端数', '数量', 'ロケーション', '担当者名', '総原価', '年月', 'セッションID', '種別', '備考'
      ]]);
      sheet.getRange(1, 1, 1, 16).setFontWeight('bold');
      break;
    case '入荷リスト':
      sheet.getRange(1, 1, 1, 15).setValues([[
        'ID', '日時', '単品JAN', '商品名', '原価', '箱入数', '箱数', '端数', '数量', 'ロケーション', '担当者名', '総原価', '年月', '仕入単価', '入荷理由'
      ]]);
      sheet.getRange(1, 1, 1, 15).setFontWeight('bold');
      break;
    case '出荷リスト':
      sheet.getRange(1, 1, 1, 8).setValues([['ID', '日時', '単品JAN', '商品名', '数量', '出荷元', '担当者名', '年月']]);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
      break;
    case '在庫移動リスト':
      sheet.getRange(1, 1, 1, 9).setValues([['ID', '移動日時', '単品JAN', '商品名', '移動元', '移動先', '移動数量', '担当者名', '移動日']]);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
      break;
    case '操作ログ':
      sheet.getRange(1, 1, 1, 6).setValues([['日時', 'メールアドレス', '担当者名', '画面', '操作', '備考']]);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      break;
    case 'パフォーマンスログ':
      sheet.getRange(1, 1, 1, 4).setValues([['日時', 'API名', '処理時間(ms)', '担当者']]);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      break;
    default:
      break;
  }
}
