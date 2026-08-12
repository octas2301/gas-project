/**
 * タイムセール_マスタ: 実質戻し（最終売価固定＋ポイント減衰）スケジュール提案。
 * 人必須: 目標売価円（最終売価）＋販促ポイント%。円・実質は計算表示。
 * 採用＝シートへ即書き込み。減衰期間／減衰段%／減衰間隔を埋める。
 * Python price_recovery_logic.propose_points_taper と同式。
 */

var TS_RECOVERY_PERIODS_ = ['1か月', '2か月', '3か月', '4か月', '5か月', '6か月'];
var TS_RECOVERY_INTERVALS_ = ['1週間', '2週間', '1か月', '2か月'];
var TS_PROMO_PCT_MAX_ = 50;
var TS_HUMAN_INPUT_BG_ = '#FFF2CC';

/** 実質戻しブロック（人入力→表示→状態）。sheet_schema.PRICE_RECOVERY_COLS と同順 */
var TS_RECOVERY_HEADERS_ = [
  '目標売価円',
  '販促ポイント%',
  '減衰期間',
  '減衰段%',
  '減衰間隔',
  '減衰開始日',
  '減衰実行依頼',
  '販促ポイント円',
  '実質価格円',
  '減衰中ポイント%',
  '次回減衰後%',
  '減衰進捗',
  '減衰状態',
  '次回減衰日',
  '最終減衰実行日時',
  '現在売価円'
];

/** 行2以降を黄にする人入力列（sheet_schema.MASTER_HUMAN_INPUT_COLS） */
var TS_HUMAN_INPUT_COLS_ = [
  '有効',
  '期間中ポイント%',
  'ポイントメモ',
  '目標売価円',
  '販促ポイント%',
  '減衰期間',
  '減衰段%',
  '減衰間隔',
  '減衰開始日',
  '減衰実行依頼',
  'メモ'
];

/** ヘッダ色（Python sheet_schema.MASTER_HEADER_COLOR_GROUPS と同色） */
var TS_MASTER_HEADER_COLOR_GROUPS_ = [
  {
    label: '基本',
    color: '#E8EAED',
    names: ['SKU', 'ASIN', '親ASIN', '商品名', '画像URL', 'marketplace', '通貨', '有効']
  },
  {
    label: 'SC取込',
    color: '#D2E3FC',
    names: ['出品者価格_SC', 'タイムセール価格_SC', '販売商品数_SC', 'V30', 'Q_fba', '原価U']
  },
  {
    label: 'ポイント',
    color: '#E8DEF8',
    names: [
      '期間中ポイント%', 'ポイントメモ',
      '期間中ポイント円',
      'セール前ポイント%', 'セール前ポイント円',
      '出品者ポイント現在%', '出品者ポイント現在円',
      'ポイント状態'
    ]
  },
  {
    label: '実質戻し',
    color: '#FCE8C3',
    names: TS_RECOVERY_HEADERS_
  },
  {
    label: 'レーンA実績',
    color: '#C8E6C9',
    names: ['A実施', 'A最終送付日時', 'A期間', 'A価格円', 'Aログ参照']
  },
  {
    label: 'メモ',
    color: '#FFF9C4',
    names: ['メモ']
  }
];

/**
 * メニュー: マスタ1行目グループ色＋人入力列の黄セル
 */
function menuApplyMasterHeaderColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('タイムセール_マスタ');
  if (!sh) {
    SpreadsheetApp.getUi().alert('タイムセール_マスタ がありません');
    return;
  }
  ensureMasterRecoveryColumns_(sh);
  renameLegacyFinalPriceHeader_(sh);
  const n = applyMasterHeaderGroupColors_(sh);
  const y = applyMasterHumanInputYellow_(sh);
  const f = applyMasterDisplayFormulas_(sh);
  SpreadsheetApp.getUi().alert(
    'マスタヘッダ色付け',
    '色付け列数: ' + n + '／人入力黄: ' + y + '／表示数式: ' + f +
      '\n基本=灰／SC=青／ポイント=紫／実質戻し=琥珀／A=緑／メモ=ヘッダ黄' +
      '\n人入力セル(行2〜)=黄 #FFF2CC' +
      '\n販促ポイント円・実質価格円=数式（販促売価円列は削除）',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  Logger.log(JSON.stringify({
    stepName: 'menuApplyMasterHeaderColors',
    state: 'DONE',
    painted: n,
    humanYellow: y,
    displayFormulas: f
  }));
}

/**
 * Python price_recovery_logic 側と同式のヘッダ色分け
 */
function applyMasterHeaderGroupColors_(sh) {
  const headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  const colorByName = {};
  TS_MASTER_HEADER_COLOR_GROUPS_.forEach(function (g) {
    g.names.forEach(function (name) {
      colorByName[name] = g.color;
    });
  });
  var painted = 0;
  for (var i = 0; i < headers.length; i++) {
    const name = String(headers[i] || '').trim();
    const color = colorByName[name];
    if (!color) continue;
    sh.getRange(1, i + 1)
      .setBackground(color)
      .setFontWeight('bold')
      .setFontColor('#202124');
    painted++;
  }
  return painted;
}

/**
 * 人入力列の行2〜を黄色
 */
function applyMasterHumanInputYellow_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const lastRow = Math.max(1, sh.getLastRow());
  if (lastRow < 2) return 0;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const human = {};
  TS_HUMAN_INPUT_COLS_.forEach(function (n) { human[n] = true; });
  var painted = 0;
  const numRows = lastRow - 1;
  for (var i = 0; i < headers.length; i++) {
    const name = String(headers[i] || '').trim();
    if (!human[name]) continue;
    sh.getRange(2, i + 1, numRows, 1).setBackground(TS_HUMAN_INPUT_BG_);
    painted++;
  }
  return painted;
}

/**
 * 販促ポイント円 = ROUND(目標×販促%/100)
 * 実質価格円 = 目標 − 販促ポイント円
 * 販促売価円列は削除済み。
 */
function applyMasterDisplayFormulas_(sh) {
  const map = headerIndexMap_(sh);
  const cTarget = map['目標売価円'];
  const cPct = map['販促ポイント%'];
  const cYen = map['販促ポイント円'];
  const cEff = map['実質価格円'];
  if (!cTarget || !cPct || !cYen || !cEff) return 0;
  const lastRow = Math.max(sh.getLastRow(), 50);
  const numRows = lastRow - 1;
  if (numRows < 1) return 0;
  const yenFormulas = [];
  const effFormulas = [];
  const nextFormulas = [];
  const cActive = map['減衰中ポイント%'];
  const cStep = map['減衰段%'];
  const cNext = map['次回減衰後%'];
  const cBefore = map['セール前ポイント%'];
  for (var r = 2; r <= lastRow; r++) {
    const tA1 = sh.getRange(r, cTarget).getA1Notation();
    const pA1 = sh.getRange(r, cPct).getA1Notation();
    const yA1 = sh.getRange(r, cYen).getA1Notation();
    yenFormulas.push(['=IF(OR(' + tA1 + '="",' + pA1 + '=""),"",ROUND(' + tA1 + '*' + pA1 + '/100,0))']);
    effFormulas.push(['=IF(OR(' + tA1 + '="",' + yA1 + '=""),"",' + tA1 + '-' + yA1 + ')']);
    if (cActive && cStep && cNext) {
      const aA1 = sh.getRange(r, cActive).getA1Notation();
      const sA1 = sh.getRange(r, cStep).getA1Notation();
      const bA1 = cBefore ? sh.getRange(r, cBefore).getA1Notation() : '';
      const beforeExpr = bA1
        ? ('IF(OR(' + bA1 + '="",' + bA1 + '=0),1,' + bA1 + ')')
        : '1';
      nextFormulas.push([
        '=IF(OR(' + sA1 + '="",AND(' + aA1 + '="",' + pA1 + '="")),"",MAX(' +
          beforeExpr + ',IF(' + aA1 + '="",' + pA1 + ',' + aA1 + ')-' + sA1 + '))'
      ]);
    }
  }
  sh.getRange(2, cYen, numRows, 1).setFormulas(yenFormulas);
  sh.getRange(2, cEff, numRows, 1).setFormulas(effFormulas);
  var n = numRows * 2;
  if (cNext && nextFormulas.length) {
    sh.getRange(2, cNext, numRows, 1).setFormulas(nextFormulas);
    n += numRows;
  }
  return n;
}

/**
 * 選択行の減衰実行依頼=TRUE（ポーラー／taper_send --poll が拾う）
 */
function menuRequestTaperRunSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('タイムセール_マスタ');
  if (!sh) {
    SpreadsheetApp.getUi().alert('タイムセール_マスタ がありません');
    return;
  }
  ensureMasterRecoveryColumns_(sh);
  const map = headerIndexMap_(sh);
  if (!map['減衰実行依頼']) {
    SpreadsheetApp.getUi().alert('減衰実行依頼 列がありません。列並べ替え後に再実行してください');
    return;
  }
  const range = sh.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('行を選択してください');
    return;
  }
  var n = 0;
  for (var r = range.getRow(); r <= range.getLastRow(); r++) {
    if (r < 2) continue;
    setCell_(sh, r, map['減衰実行依頼'], true);
    n++;
  }
  SpreadsheetApp.getUi().alert(
    '6-② 今すぐ1段下げたいSKUに印',
    n + ' 行に印を付けました。\n次の日次 taper_send.py --poll が拾います。急ぐときは 99-⑦ で手動1回。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  Logger.log(JSON.stringify({ stepName: 'menuRequestTaperRunSelection', state: 'DONE', rows: n }));
}

/** 旧ヘッダ「最終売価円」→「目標売価円」 */
function renameLegacyFinalPriceHeader_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === '最終売価円') {
      sh.getRange(1, i + 1).setValue('目標売価円');
      Logger.log(JSON.stringify({
        stepName: 'renameLegacyFinalPriceHeader_',
        state: 'DONE',
        col: i + 1
      }));
    }
  }
}

/**
 * メニュー: 選択行に提案→確認後書き込み
 */
function menuProposePriceRecoverySelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('タイムセール_マスタ');
  if (!sh) {
    SpreadsheetApp.getUi().alert('タイムセール_マスタ がありません');
    return;
  }
  ensureMasterRecoveryColumns_(sh);
  const map = headerIndexMap_(sh);
  const range = sh.getActiveRange();
  if (!range || range.getSheet().getName() !== 'タイムセール_マスタ') {
    SpreadsheetApp.getUi().alert('タイムセール_マスタ上で行を選択してください');
    return;
  }
  const start = Math.max(2, range.getRow());
  const end = range.getLastRow();
  const proposals = [];
  for (var r = start; r <= end; r++) {
    const p = buildRecoveryProposalForRow_(sh, r, map);
    if (p.error) {
      proposals.push({ row: r, error: p.error });
    } else {
      proposals.push({ row: r, proposal: p.proposal, sku: p.sku, asin: p.asin });
    }
  }
  if (!proposals.length) {
    SpreadsheetApp.getUi().alert('対象行がありません');
    return;
  }
  const ok = proposals.filter(function (x) { return x.proposal; });
  const ng = proposals.filter(function (x) { return x.error; });
  var msg = '提案 ' + ok.length + ' 件';
  if (ng.length) msg += '／スキップ ' + ng.length + ' 件\n';
  else msg += '\n';
  ok.slice(0, 8).forEach(function (x) {
    const p = x.proposal;
    msg +=
      '\n行' + x.row + ' ' + (x.sku || x.asin || '') +
      '\n  販促' + p.promoPct + '%（' + p.promoYen + '円）実質' + p.effective +
      '／段−' + p.stepPct + '% × ' + p.interval + ' ／ 期間' + p.period +
      '（' + p.steps + '段）\n';
  });
  if (ok.length > 8) msg += '\n…他 ' + (ok.length - 8) + ' 件\n';
  ng.slice(0, 5).forEach(function (x) {
    msg += '\n行' + x.row + ': ' + x.error;
  });
  msg += '\n\n採用してシートに書き込みますか？';
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('ポイント減衰スケジュール提案', msg, ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) {
    Logger.log(JSON.stringify({ stepName: 'menuProposePriceRecoverySelection', state: 'CANCELLED' }));
    return;
  }
  var written = 0;
  ok.forEach(function (x) {
    writeRecoveryProposal_(sh, x.row, map, x.proposal);
    written++;
  });
  applyRecoveryValidations_(sh, map);
  ui.alert('書き込み完了: ' + written + ' 件');
  Logger.log(JSON.stringify({
    stepName: 'menuProposePriceRecoverySelection',
    state: 'DONE',
    written: written,
    skipped: ng.length
  }));
}

/**
 * メニュー: 有効かつ目標・販促あり・戻し未設定の行を一括提案
 */
function menuProposePriceRecoveryEligible() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('タイムセール_マスタ');
  if (!sh) {
    SpreadsheetApp.getUi().alert('タイムセール_マスタ がありません');
    return;
  }
  ensureMasterRecoveryColumns_(sh);
  const map = headerIndexMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert('データ行がありません');
    return;
  }
  const proposals = [];
  for (var r = 2; r <= last; r++) {
    if (!isTruthyCell_(sh, r, map['有効'])) continue;
    const periodVal = cellStr_(sh, r, map['減衰期間']);
    if (periodVal) continue; // 未設定のみ
    const p = buildRecoveryProposalForRow_(sh, r, map);
    if (p.error) continue;
    proposals.push({ row: r, proposal: p.proposal, sku: p.sku, asin: p.asin });
  }
  if (!proposals.length) {
    SpreadsheetApp.getUi().alert(
      '対象なし（有効・目標売価円・販促ポイント%あり・減衰期間が空の行）'
    );
    return;
  }
  const ui = SpreadsheetApp.getUi();
  var msg = proposals.length + ' 件に提案を書き込みますか？\n';
  proposals.slice(0, 10).forEach(function (x) {
    msg +=
      '\n行' + x.row + ' ' + (x.sku || x.asin) +
      ': 販促' + x.proposal.promoPct + '%／段−' + x.proposal.stepPct +
      '%/' + x.proposal.interval + '/' + x.proposal.period;
  });
  if (ui.alert('ポイント減衰一括提案', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  proposals.forEach(function (x) {
    writeRecoveryProposal_(sh, x.row, map, x.proposal);
  });
  applyRecoveryValidations_(sh, map);
  ui.alert('書き込み完了: ' + proposals.length + ' 件');
  Logger.log(JSON.stringify({
    stepName: 'menuProposePriceRecoveryEligible',
    state: 'DONE',
    written: proposals.length
  }));
}

function buildRecoveryProposalForRow_(sh, row, map) {
  const sku = cellStr_(sh, row, map['SKU']);
  const asin = cellStr_(sh, row, map['ASIN']);
  const target = toNumber_(cellStr_(sh, row, map['目標売価円']));
  const promoPct = toNumber_(cellStr_(sh, row, map['販促ポイント%']));
  if (target === null || promoPct === null) {
    return { error: '目標売価円と販促ポイント%を数値で入力してください', sku: sku, asin: asin };
  }
  if (target <= 0) {
    return { error: '目標売価円は正の数にしてください', sku: sku, asin: asin };
  }
  if (promoPct < 0 || promoPct > TS_PROMO_PCT_MAX_) {
    return { error: '販促ポイント%は 0〜' + TS_PROMO_PCT_MAX_ + ' です', sku: sku, asin: asin };
  }
  var endPct = toNumber_(cellStr_(sh, row, map['セール前ポイント%']));
  if (endPct === null) endPct = 1;
  if (promoPct < endPct) {
    return { error: '販促ポイント%は最終終着%（セール前ポイント%）以上にしてください', sku: sku, asin: asin };
  }
  return {
    sku: sku,
    asin: asin,
    proposal: computePointsTaperProposal_(target, promoPct, endPct)
  };
}

/**
 * 最終売価＋販促ポイント% → 円表示＋減衰スケジュール。
 * Python propose_points_taper と同式。
 */
function computePointsTaperProposal_(finalPrice, promoPct, endPct) {
  const yen = Math.round(finalPrice * promoPct / 100);
  const effective = Math.round(finalPrice - yen);
  var interval = '2週間';
  var iw = 2;
  var months = periodMonthsFromPromoPct_(promoPct);
  var weeksBudget = Math.max(iw, Math.round(months * 4.345));
  var nSteps = Math.max(2, Math.floor(weeksBudget / iw));
  var delta = promoPct - endPct;
  var stepPct = Math.max(1, Math.round(delta / nSteps));
  var guard = 0;
  while (stepPct * nSteps < delta - 0.5 && months < 6 && guard < 4) {
    months++;
    weeksBudget = Math.max(iw, Math.round(months * 4.345));
    nSteps = Math.max(2, Math.floor(weeksBudget / iw));
    stepPct = Math.max(1, Math.round(delta / nSteps));
    guard++;
  }
  const period = months + 'か月';
  const progress =
    '未開始｜最終売価' + finalPrice + '／販促ポイント' + promoPct +
    '%（' + yen + '円）／実質' + effective +
    '｜終着' + endPct + '%｜段−' + stepPct + '%×' + interval +
    '｜目安' + nSteps + '段・期間' + period;
  return {
    period: period,
    stepPct: stepPct,
    interval: interval,
    steps: nSteps,
    weeks: weeksBudget,
    progress: progress,
    current: finalPrice,
    promoPct: promoPct,
    promoYen: yen,
    effective: effective,
    endPct: endPct
  };
}

/** @deprecated 互換: 旧名 */
function computeRecoveryProposal_(target, promoPct) {
  return computePointsTaperProposal_(target, promoPct, 1);
}

function periodMonthsFromPromoPct_(pct) {
  if (pct <= 10) return 2;
  if (pct <= 20) return 3;
  if (pct <= 35) return 3;
  if (pct <= 45) return 4;
  return 5;
}

function writeRecoveryProposal_(sh, row, map, p) {
  // 円・実質・次回減衰後%はシート数式
  setCell_(sh, row, map['減衰期間'], p.period);
  setCell_(sh, row, map['減衰段%'], p.stepPct); // 段%
  setCell_(sh, row, map['減衰間隔'], p.interval);
  if (map['減衰中ポイント%']) setCell_(sh, row, map['減衰中ポイント%'], p.promoPct);
  setCell_(sh, row, map['減衰進捗'], p.progress);
  setCell_(sh, row, map['減衰状態'], '未開始');
  setCell_(sh, row, map['現在売価円'], p.current);
}

function ensureMasterRecoveryColumns_(sh) {
  renameLegacyFinalPriceHeader_(sh);
  const headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var missing = [];
  TS_RECOVERY_HEADERS_.forEach(function (name) {
    if (headers.indexOf(name) < 0) missing.push(name);
  });
  if (missing.length) {
    var col = headers.length + 1;
    missing.forEach(function (name) {
      sh.getRange(1, col).setValue(name).setFontWeight('bold');
      col++;
    });
  }
  applyMasterHeaderGroupColors_(sh);
  applyMasterHumanInputYellow_(sh);
  applyMasterDisplayFormulas_(sh);
}

/**
 * メニュー: 列を schema 順に並べ替え（人入力を各グループ先頭）＋色／黄セル
 * ※ データは列名で保全。最終売価円は目標売価円へ改名。
 */
function menuRealignMasterColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('タイムセール_マスタ');
  if (!sh) {
    SpreadsheetApp.getUi().alert('タイムセール_マスタ がありません');
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const conf = ui.alert(
    'マスタ列の並べ替え',
    '目的グループ順に並べ替えます（人入力列を先頭）。\n' +
      '最終売価円→目標売価円。数式は保全します。\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (conf !== ui.Button.YES) {
    Logger.log(JSON.stringify({ stepName: 'menuRealignMasterColumns', state: 'CANCELLED' }));
    return;
  }
  const stats = realignMasterColumnsToSchema_(sh);
  applyMasterHeaderGroupColors_(sh);
  const y = applyMasterHumanInputYellow_(sh);
  const f = applyMasterDisplayFormulas_(sh);
  const map = headerIndexMap_(sh);
  applyRecoveryValidations_(sh, map);
  ui.alert(
    '並べ替え完了',
    '列数: ' + stats.cols + '／行: ' + stats.rows +
      '／人入力黄: ' + y + '／表示数式: ' + f +
      (stats.extras.length ? '\n末尾保全: ' + stats.extras.join(', ') : ''),
    ui.ButtonSet.OK
  );
  Logger.log(JSON.stringify({
    stepName: 'menuRealignMasterColumns',
    state: 'DONE',
    cols: stats.cols,
    rows: stats.rows,
    extras: stats.extras
  }));
}

/** sheet_schema.MASTER_HEADERS と同順（GAS側コピー） */
function tsMasterHeadersCanonical_() {
  return [
    'SKU', 'ASIN', '親ASIN', '商品名', '画像URL', 'marketplace', '通貨', '有効',
    '出品者価格_SC', 'タイムセール価格_SC', '販売商品数_SC', 'V30', 'Q_fba', '原価U',
    '期間中ポイント%', 'ポイントメモ', '期間中ポイント円',
    'セール前ポイント%', 'セール前ポイント円',
    '出品者ポイント現在%', '出品者ポイント現在円', 'ポイント状態'
  ].concat(TS_RECOVERY_HEADERS_).concat([
    'A実施', 'A最終送付日時', 'A期間', 'A価格円', 'Aログ参照', 'メモ'
  ]);
}

function realignMasterColumnsToSchema_(sh) {
  renameLegacyFinalPriceHeader_(sh);
  const lastCol = Math.max(1, sh.getLastColumn());
  const lastRow = Math.max(1, sh.getLastRow());
  const width = lastCol;
  const height = lastRow;
  const values = sh.getRange(1, 1, height, width).getValues();
  const formulas = sh.getRange(1, 1, height, width).getFormulas();
  const oldHeaders = values[0].map(function (h) { return String(h || '').trim(); });
  const colIndex = {};
  oldHeaders.forEach(function (h, i) {
    if (!h) return;
    if (colIndex[h] === undefined) colIndex[h] = i;
  });
  // 別名
  if (colIndex['目標売価円'] === undefined && colIndex['最終売価円'] !== undefined) {
    colIndex['目標売価円'] = colIndex['最終売価円'];
  }

  const canonical = tsMasterHeadersCanonical_();
  const extras = [];
  oldHeaders.forEach(function (h) {
    if (!h) return;
    if (h === '最終売価円') return;
    if (canonical.indexOf(h) < 0 && extras.indexOf(h) < 0) extras.push(h);
  });
  const want = canonical.concat(extras);

  const out = [];
  out.push(want);
  for (var r = 1; r < height; r++) {
    const row = [];
    var any = false;
    for (var c = 0; c < want.length; c++) {
      const name = want[c];
      const src = colIndex[name];
      var cell = '';
      if (src !== undefined) {
        const f = formulas[r][src];
        cell = (f && String(f).charAt(0) === '=') ? f : values[r][src];
      }
      if (cell !== '' && cell !== null && cell !== undefined) any = true;
      row.push(cell === null || cell === undefined ? '' : cell);
    }
    if (any) out.push(row);
  }

  sh.clear();
  if (out.length && out[0].length) {
    sh.getRange(1, 1, out.length, out[0].length).setValues(
      out.map(function (row) {
        return row.map(function (v) {
          // setValues は数式文字列も値として入る → 後で setFormula
          return (typeof v === 'string' && v.charAt(0) === '=') ? '' : v;
        });
      })
    );
    for (var rr = 0; rr < out.length; rr++) {
      for (var cc = 0; cc < out[rr].length; cc++) {
        const v = out[rr][cc];
        if (typeof v === 'string' && v.charAt(0) === '=') {
          sh.getRange(rr + 1, cc + 1).setFormula(v);
        }
      }
    }
  }
  return { cols: want.length, rows: Math.max(0, out.length - 1), extras: extras };
}

/**
 * 減衰期間／減衰間隔のみプルダウン。
 * ※ Sheet.getRange(r,c,numRows,numCols) の第3・4は行数・列数（最終座標ではない）。
 * 旧バグで横に広がった検証が残るため、戻しブロック全体を先にクリアする。
 */
function applyRecoveryValidations_(sh, map) {
  const lastDataRow = Math.max(2, sh.getLastRow());
  const numDataRows = lastDataRow - 1; // 行2〜lastDataRow
  const clearRows = Math.max(numDataRows, 200);

  var startCol = map['目標売価円'] || map['減衰期間'] || 1;
  var endCol = map['現在売価円'] || map['減衰間隔'] || sh.getLastColumn();
  if (endCol < startCol) endCol = sh.getLastColumn();
  // 旧バグ（numCols=列番号）で右側まで汚染されていることがあるので余裕を見る
  endCol = Math.min(sh.getLastColumn(), Math.max(endCol, startCol) + 15);
  sh.getRange(2, startCol, clearRows, endCol - startCol + 1).clearDataValidations();

  if (map['減衰期間'] && numDataRows > 0) {
    sh.getRange(2, map['減衰期間'], numDataRows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(TS_RECOVERY_PERIODS_, true)
        .setAllowInvalid(false)
        .build()
    );
  }
  if (map['減衰間隔'] && numDataRows > 0) {
    sh.getRange(2, map['減衰間隔'], numDataRows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(TS_RECOVERY_INTERVALS_, true)
        .setAllowInvalid(false)
        .build()
    );
  }
  if (map['減衰実行依頼'] && numDataRows > 0) {
    sh.getRange(2, map['減衰実行依頼'], numDataRows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build()
    );
  }
}

/**
 * メニュー: 戻し列のズレたプルダウンを掃除して付け直し
 */
function menuFixRecoveryValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('タイムセール_マスタ');
  if (!sh) {
    SpreadsheetApp.getUi().alert('タイムセール_マスタ がありません');
    return;
  }
  ensureMasterRecoveryColumns_(sh);
  const map = headerIndexMap_(sh);
  applyRecoveryValidations_(sh, map);
  SpreadsheetApp.getUi().alert(
    'プルダウン修正',
    '減衰期間＝1〜6か月／減衰間隔＝週・月／減衰実行依頼＝チェック。\n他列の検証はクリアしました。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  Logger.log(JSON.stringify({
    stepName: 'menuFixRecoveryValidations',
    state: 'DONE',
    periodCol: map['減衰期間'] || null,
    intervalCol: map['減衰間隔'] || null
  }));
}

function headerIndexMap_(sh) {
  const headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  const map = {};
  headers.forEach(function (h, i) {
    const name = String(h || '').trim();
    if (name) map[name] = i + 1;
  });
  // 旧ヘッダ互換
  if (!map['目標売価円'] && map['最終売価円']) map['目標売価円'] = map['最終売価円'];
  if (!map['減衰期間'] && map['戻し期間']) map['減衰期間'] = map['戻し期間'];
  if (!map['減衰段%'] && map['戻し価格円']) map['減衰段%'] = map['戻し価格円'];
  if (!map['減衰間隔'] && map['戻し間隔']) map['減衰間隔'] = map['戻し間隔'];
  if (!map['減衰進捗'] && map['戻し進捗']) map['減衰進捗'] = map['戻し進捗'];
  if (!map['減衰状態'] && map['戻し状態']) map['減衰状態'] = map['戻し状態'];
  if (!map['次回減衰日'] && map['次回戻し日']) map['次回減衰日'] = map['次回戻し日'];
  return map;
}

function cellStr_(sh, row, col) {
  if (!col) return '';
  return String(sh.getRange(row, col).getValue() || '').trim();
}

function setCell_(sh, row, col, value) {
  if (!col) return;
  sh.getRange(row, col).setValue(value);
}

function toNumber_(s) {
  if (s === '' || s === null || s === undefined) return null;
  const n = Number(String(s).replace(/[,円%]/g, ''));
  return isFinite(n) ? n : null;
}

function isTruthyCell_(sh, row, col) {
  if (!col) return true;
  const v = sh.getRange(row, col).getValue();
  if (v === true) return true;
  const s = String(v || '').trim().toUpperCase();
  return s === '' || s === 'TRUE' || s === 'はい' || s === 'YES' || s === '1' || s === '○';
}
