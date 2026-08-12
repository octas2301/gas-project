/**
 * P1b: ②最新バルク確認 → Cursor指示（下にUL固定手順）→ 再build指示
 * フォルダIDは tools/amazon_deals_bulk/config.example.json と同期
 */

var TS_DEALS_FOLDER_02_ID_ = '1OOYfcgiskqahZImJfzXNM3TW1HwLQN2R';
var TS_DEALS_FOLDER_03_ID_ = '16R6IjbAOkafsPJK1Du7SSYQ9tRpo4IHH';
var TS_DEALS_FOLDER_02_NAME_ = '02.amazon公式タイムセールバルクファイル保存（人間が保存）';
var TS_DEALS_FOLDER_03_NAME_ = '03.amazon公式タイムセールバルクファイル作成（Python作成⇒人間がUL）';
var TS_DEALS_SC_BULK_URL_ =
  'https://sellercentral.amazon.co.jp/promotion-central/manage-in-bulk?tab=upload';

var TS_DEALS_UL_GUIDE_FIXED_ = [
  '【SCアップロード手順（Agent完了後）】',
  '出力フォルダ: 「' + TS_DEALS_FOLDER_03_NAME_ + '」の最新 xlsx',
  '1. ③の最新提出xlsxを開く／ダウンロード',
  '2. タイムセール価格がバルク最大以下か、参加行だけか確認',
  '3. Seller Central → 広告 → タイムセール → 一括管理',
  '   ' + TS_DEALS_SC_BULK_URL_,
  '4. ファイルをアップロード（xlsx・サイズ上限に注意）',
  '5. 処理成功を確認 → 一覧で「近日開始／参加中」を目視',
  '注意: ②最新から作った③だけ使う／既登録Saleの再ULは任意／名付きは候補に出たら即登録'
].join('\n');

/**
 * メニュー: ②保存確認のうえ Cursor 指示を出す（UL手順はダイアログ下部の固定文）
 */
function menuTimeSaleP1bCursorPrompt() {
  const ui = SpreadsheetApp.getUi();
  const check = checkTimeSaleFolder02LatestXlsx_();
  const folder02Name = TS_DEALS_FOLDER_02_NAME_;
  const folder03Name = TS_DEALS_FOLDER_03_NAME_;

  var confirmMsg =
    '【必須】提出xlsxは必ず「②の最新推奨バルク」から作ります（価格・候補を最新にするため）。\n\n' +
    '保存先フォルダ名:\n「' + folder02Name + '」\n\n';
  if (check.ok) {
    confirmMsg +=
      'Drive上の最新xlsx:\n' + check.name + '\n更新: ' + check.updated + '\n\n' +
      'このファイルを②に保存済みで、Cursorに提出xlsx作成を依頼しますか？';
  } else {
    confirmMsg +=
      '⚠ ' + (check.message || '②フォルダにxlsxが見つかりません') + '\n\n' +
      'SCから推奨バルクをDLし、上記フォルダへ保存してから「はい」を押してください。\n' +
      '保存済みなら「はい」で続行（ローカル同期のみの場合あり）しますか？';
  }

  const ans = ui.alert('P1b 公式B提出 — ②確認', confirmMsg, ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) {
    ui.alert('中止しました。先に②へ最新バルクを保存してください。');
    return;
  }

  const prompt = [
    'gas-project で Amazon 公式B提出xlsx（P1b）を実行してください。',
    '',
    '方針: ②最新バルクの当該SKU行を正。カタログだけのBF等は入れない。',
    '既登録（UL済/実施中）は再ULしない。カスタムは公式と非重なりで②の日付のまま提出対象へ。',
    '',
    '【人が済んでいること】',
    '- SCから推奨バルクをDLし、次のフォルダへ保存済み:',
    '  ' + folder02Name,
    check.ok
      ? ('- Drive上の最新: ' + check.name + '（' + check.updated + '）')
      : '- （Drive未検出時はローカル②の最新を使う）',
    '',
    '【Agent手順】',
    '1. cd tools/amazon_deals_bulk',
    '2. ②の最新xlsxを確認（paths.latest_xlsx / config local_02）',
    '3. python sync_master_and_exec.py --write  （SKUローカル候補→スプシ転記・提出対象自動）',
    '4. dry: python build_submit_xlsx.py',
    '5. --write: python build_submit_xlsx.py --write → フォルダ③',
    '   出力先: ' + folder03Name,
    '6. 出力パス・opt_in件数・使った②ファイル名・スケジュール一覧を報告',
    '',
    'HUMAN_RUN: docs/org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md'
  ].join('\n');

  showTimeSaleP1bDialog_('1. 公式タイムセールの提出xlsxを作る — Agent依頼文', prompt, TS_DEALS_UL_GUIDE_FIXED_, 'menuTimeSaleP1bCursorPrompt');
}

/**
 * メニュー: 人が③を直した／スプシ修正後の再build（Cursor指示）
 */
function menuTimeSaleP1bRebuildPrompt() {
  const check03 = checkTimeSaleFolder03LatestXlsx_();
  const prompt = [
    'gas-project で Amazon 公式B提出xlsx（P1b）を再作成してください。',
    '',
    '前提: 人がタイムセールシートまたは要件を修正済み。②は前回と同じ最新でよい（価格刷新が必要なら先に②を差し替え）。',
    check03.ok ? ('現在の③最新: ' + check03.name + '（' + check03.updated + '）') : '③に既存xlsxなし',
    '',
    '【Agent手順】',
    '1. cd tools/amazon_deals_bulk',
    '2. （シートだけ直した場合）python build_submit_xlsx.py --write',
    '3. （②も取り直した／候補を再転記する場合）python sync_master_and_exec.py --write → build_submit_xlsx.py --write',
    '4. 出力パス・opt_in・スケジュールを報告',
    '',
    'HUMAN_RUN: docs/org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md'
  ].join('\n');

  showTimeSaleP1bDialog_('1-② 提出xlsxを作り直す — Agent依頼文', prompt, TS_DEALS_UL_GUIDE_FIXED_, 'menuTimeSaleP1bRebuildPrompt');
}

/** @deprecated メニュー削除済み。互換のため残す場合は再buildへ誘導 */
function menuTimeSaleP1bUploadGuide() {
  SpreadsheetApp.getUi().alert(
    'このメニューは廃止しました。\n' +
    'UL手順は「1. 公式タイムセールの提出xlsxを作る（Agent依頼文）」ダイアログ下部の固定文を見てください。\n' +
    '修正後の再作成は「1-② 提出xlsxを作り直す（修正後・Agent依頼文）」を使います。'
  );
}

function showTimeSaleP1bDialog_(title, prompt, ulGuide, stepName) {
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:sans-serif;font-size:13px;padding:12px;}' +
    'textarea{width:100%;height:240px;font-family:monospace;font-size:11px;}' +
    'pre.ul{white-space:pre-wrap;background:#f6f8fa;border:1px solid #ddd;padding:10px;' +
    'font-size:12px;line-height:1.45;margin-top:8px;}</style></head><body>' +
    '<h3 style="margin-top:0;">' + title + '</h3>' +
    '<p>下の文をコピーし、Cursor Agent に貼って実行してください。</p>' +
    '<textarea id="t">' + String(prompt).replace(/</g, '&lt;') + '</textarea>' +
    '<p><button onclick="var t=document.getElementById(\'t\');t.select();document.execCommand(\'copy\');">コピー</button></p>' +
    '<hr>' +
    '<pre class="ul">' + String(ulGuide).replace(/</g, '&lt;') + '</pre>' +
    '</body></html>'
  ).setWidth(680).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, title);
  Logger.log(JSON.stringify({ stepName: stepName, state: 'DONE' }));
}

function checkTimeSaleFolder02LatestXlsx_() {
  return checkTimeSaleDealsFolderLatestXlsx_(TS_DEALS_FOLDER_02_ID_, TS_DEALS_FOLDER_02_NAME_);
}

function checkTimeSaleFolder03LatestXlsx_() {
  return checkTimeSaleDealsFolderLatestXlsx_(TS_DEALS_FOLDER_03_ID_, TS_DEALS_FOLDER_03_NAME_);
}

function checkTimeSaleDealsFolderLatestXlsx_(folderId, folderLabel) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    var best = null;
    var bestMs = 0;
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName() || '';
      if (!/\.xlsx?$/i.test(name) || name.indexOf('~$') === 0) continue;
      const ms = f.getLastUpdated().getTime();
      if (ms >= bestMs) {
        bestMs = ms;
        best = f;
      }
    }
    if (!best) {
      return {
        ok: false,
        message: '「' + folderLabel + '」に xlsx がありません'
      };
    }
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    return {
      ok: true,
      name: best.getName(),
      updated: Utilities.formatDate(best.getLastUpdated(), tz, 'yyyy-MM-dd HH:mm'),
      url: best.getUrl(),
      folderName: folder.getName()
    };
  } catch (e) {
    return {
      ok: false,
      message: 'フォルダ参照エラー（' + folderLabel + '）: ' + e
    };
  }
}
