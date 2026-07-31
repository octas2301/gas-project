/**
 * Lv4 承認①済 → Amazon バルク埋め用データ（GENERATED）
 * 要件: docs/org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md
 *
 * - 楽天 generateRakutenCSV / Yahoo.js 出品APIは触らない
 * - マスタ「在庫数」「JAN」への書込禁止（読取のみ）
 * - 純正 .xlsm 編集・SC自動UPはしない（PACKAGED/UPは人間＋ローカル）
 * - TRACK 未設定は実行しない
 *
 * Script Properties:
 *   APPROVAL_AMAZON_LV4_ENABLED … 既定 false
 *   APPROVAL_AMAZON_LV4_TRACK … A | B | BOTH（必須。空＝実行しない）
 *   APPROVAL_AMAZON_LV4_SKIP_EXPORT … 既定 false（true=Drive書込スキップのドライラン）
 *   APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE … 既定 送料無料パターン
 *   APPROVAL_AMAZON_LV4_FOLDER_ID … GENERATED 保存先（空ならスプレッドシート親にフォルダ作成）
 *   APPROVAL_AMAZON_LV4_PARENTS_PER_SUB … 副制約・親件数/サブバッチ（既定 5）
 *   APPROVAL_AMAZON_LV4_LEAD_TIME_DAYS … エラー時のみ後付け用（初版は未使用）
 *   APPROVAL_AMAZON_LV4_STATE … レジューム用（自動）
 *   APPROVAL_AMAZON_LV4_BRAND_GATE_MODE … M2(A): manual_ok 必須（人間が制限なし確認後）。未設定＝SKIPPED_BRAND_GATE
 *   APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK … Dレ点新規のみ在庫>0でも続行（未設定＝true。false で旧 SKIPPED_IN_STOCK に戻す）
 *   APPROVAL_AMAZON_LV4_EXEMPTION_ALL_CATEGORIES … 21-⑭ を全カテゴリ `*` で記録（既定 false＝レ点カテゴリのみ）
 *   APPROVAL_AMAZON_LV4_SC_SUMMARY_ENABLED … SC処理サマリ自動記録の有効化（既定 false＝fail-closed）
 *   APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_ID … サマリ監視フォルダ（必須。未設定なら実行しない）
 *   APPROVAL_AMAZON_LV4_SC_SUMMARY_INTERVAL_MIN … 監視トリガー間隔分（5/10/15/30/60。既定 15）
 *   APPROVAL_AMAZON_LV4_SC_SUMMARY_SS_ID … トリガー用スプレッドシートID（21-⑯設置時に自動保存）
 */

var APPROVAL_AMAZON_LV4_PROP = 'APPROVAL_AMAZON_LV4_ENABLED';
var APPROVAL_AMAZON_LV4_TRACK_PROP = 'APPROVAL_AMAZON_LV4_TRACK';
var APPROVAL_AMAZON_LV4_SKIP_EXPORT_PROP = 'APPROVAL_AMAZON_LV4_SKIP_EXPORT';
var APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE_PROP = 'APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE';
var APPROVAL_AMAZON_LV4_FOLDER_ID_PROP = 'APPROVAL_AMAZON_LV4_FOLDER_ID';
var APPROVAL_AMAZON_LV4_PARENTS_PER_SUB_PROP = 'APPROVAL_AMAZON_LV4_PARENTS_PER_SUB';
var APPROVAL_AMAZON_LV4_STATE_PROP = 'APPROVAL_AMAZON_LV4_STATE';
var APPROVAL_AMAZON_LV4_BRAND_GATE_MODE_PROP = 'APPROVAL_AMAZON_LV4_BRAND_GATE_MODE';
var APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK_PROP = 'APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK';
var APPROVAL_AMAZON_LV4_EXEMPTION_ALL_PROP = 'APPROVAL_AMAZON_LV4_EXEMPTION_ALL_CATEGORIES';
var APPROVAL_AMAZON_LV4_SC_SUMMARY_PROP = 'APPROVAL_AMAZON_LV4_SC_SUMMARY_ENABLED';
var APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_PROP = 'APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_ID';
var APPROVAL_AMAZON_LV4_SC_SUMMARY_INTERVAL_PROP = 'APPROVAL_AMAZON_LV4_SC_SUMMARY_INTERVAL_MIN';
var APPROVAL_AMAZON_LV4_SC_SUMMARY_SS_PROP = 'APPROVAL_AMAZON_LV4_SC_SUMMARY_SS_ID';
var APPROVAL_AMAZON_LV4_SC_SUMMARY_TRIGGER_FN = 'runApprovalAmazonLv4ScSummaryFromTrigger';
var APPROVAL_AMAZON_LV4_SC_SUMMARY_DONE_FOLDER = '_処理済';
var APPROVAL_AMAZON_LV4_SC_SUMMARY_MAX_FILES = 20;
var APPROVAL_AMAZON_LV4_TIME_MS = 25 * 60 * 1000;
var APPROVAL_AMAZON_LV4_MAX_TRIGGER_RUNS = 40;
var APPROVAL_AMAZON_LV4_TRIGGER_FN = 'runApprovalAmazonLv4FromTrigger';
var APPROVAL_AMAZON_LV4_LOG_SHEET = '▼Lv4実行ログ(Amazon)';
var APPROVAL_AMAZON_LV4_DEFAULT_SHIPPING = '送料無料パターン';
var APPROVAL_AMAZON_LV4_BRAND = 'ノーブランド品';
var AMAZON_CHECKBOX_SHIPPING_HEADER_ = '自己発送';

var APPROVAL_AMAZON_LV4_LOG_HEADERS = [
  'recordType', 'runId', 'batchId', 'subBatchId', 'track', 'parentSkus', 'childCount',
  'status', 'fileName', 'fileUrl', 'generatedAt', 'packagedAt', 'uploadedOkAt',
  'note', 'category', 'brand', 'exemptionDate', 'evidenceUrl'
];

/**
 * メニュー 21-①: 最新 APPROVED の amazon を親単位サブバッチで GENERATED する。
 */
/**
 * メニュー 21-① / D×Amazon ファサード共用。
 * @param {{silent?:boolean, source?:string, track?:string, includeOffer?:boolean}=} opts
 *   source='child_ck' は人間レ点の新規カタログ行のみ。既定は承認①済。
 * @return {{ok:boolean, cancelled?:boolean, runId?:string, track?:string, summary?:Object, reason?:string, error?:string}}
 */
function menuApprovalAmazonLv4Run(opts) {
  var fn = 'menuApprovalAmazonLv4Run';
  opts = (opts && typeof opts === 'object') ? opts : {};
  var silent = !!opts.silent;
  var source = String(opts.source || 'approved') === 'child_ck' ? 'child_ck' : 'approved';
  var includeOfferForSplit = !!opts.includeOffer;
  if (!getBoolScriptProperty_(APPROVAL_AMAZON_LV4_PROP, false)) {
    var off = 'Lv4 Amazonは無効です。Script Properties の ' + APPROVAL_AMAZON_LV4_PROP + ' を true にしてください。';
    Logger.log('[' + fn + '] state=FAILED ' + off);
    try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    return { ok: false, reason: off };
  }
  var track = source === 'child_ck'
    ? 'B'
    : amazonApprovalLv4NormalizeTrack_(opts.track ||
        PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_TRACK_PROP));
  if (!track) {
    var noTrack = 'TRACK 未設定のため実行しません。' + APPROVAL_AMAZON_LV4_TRACK_PROP +
      ' に A / B / BOTH のいずれかを設定してください（M1は B）。';
    Logger.log('[' + fn + '] state=FAILED ' + noTrack);
    try { SpreadsheetApp.getUi().alert(noTrack); } catch (eT) {}
    return { ok: false, reason: noTrack };
  }
  if (!silent) {
    try {
      var ui = SpreadsheetApp.getUi();
      var skip = getBoolScriptProperty_(APPROVAL_AMAZON_LV4_SKIP_EXPORT_PROP, false);
      var ship = amazonApprovalLv4ShippingTemplate_();
      var res = ui.alert(
        'Lv4 Amazonバルク（GENERATED）',
        (source === 'child_ck'
          ? '人間がレ点を付けた新規カタログ子SKUを対象に、埋め用データを Drive へ出力します。\n'
          : '最新の承認①（APPROVED）の Amazon 明細を対象に、埋め用データを Drive へ出力します。\n') +
          '・TRACK=' + track + '\n' +
          '・配送テンプレ=' + ship + '\n' +
          '・マスタ在庫・JANは書き込みません\n' +
          '・純正 .xlsm / Seller Central UP は人間作業です\n' +
          (skip ? '・SKIP_EXPORT=true（ドライラン・Drive未書込）\n' : '') +
          '\n実行しますか？',
        ui.ButtonSet.OK_CANCEL
      );
      if (res !== ui.Button.OK) return { ok: false, cancelled: true };
    } catch (eUi) {}
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var runId = 'LV4_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' +
    ('000000' + Math.floor(Math.random() * 1e6)).slice(-6);
  Logger.log('[' + fn + '] state=RUNNING runId=' + runId + ' track=' + track +
    ' source=' + source + ' includeOffer=' + includeOfferForSplit + ' silent=' + silent);
  try {
    var summary = amazonApprovalLv4Run_(ss, runId, 0, null, track, source, {
      includeOffer: includeOfferForSplit
    });
    Logger.log('[' + fn + '] state=DONE runId=' + runId + ' ' + JSON.stringify(summary));
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(
          'Lv4 実行結果',
          'runId=' + runId +
            '\nbatchId=' + (summary.batchId || '') +
            '\ntrack=' + track +
            '\nサブバッチ完了=' + summary.subBatchesDone +
            '\n親成功=' + summary.parentsDone +
            '\nスキップ=' + summary.skipped +
            '\n継続待ち=' + (summary.willResume ? 'YES' : 'NO') +
            '\n' + (summary.message || ''),
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } catch (e1) {}
    }
    return { ok: true, runId: runId, track: track, summary: summary };
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED runId=' + runId + ' ' + ((err && err.message) || err));
    amazonApprovalLv4Mail_('【Lv4 Amazon】実行失敗', 'runId=' + runId + '\n' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert('Lv4失敗: ' + ((err && err.message) || err)); } catch (e2) {}
    return { ok: false, runId: runId, track: track, error: String((err && err.message) || err) };
  }
}

/**
 * メニュー 21-⑭: GTIN免除証跡（EXEMPTION）を記録する。
 * レ点新規カタログのカテゴリを自動検出し、人間確認のうえ状態シートへ追記する。
 * 全カテゴリ（`*`）は Property で明示的に選んだ場合のみ。
 */
function menuApprovalAmazonLv4RecordGtinExemption() {
  var fn = 'menuApprovalAmazonLv4RecordGtinExemption';
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (eUi) {
    Logger.log('[' + fn + '] state=FAILED UI なし（メニューから実行してください）');
    return;
  }
  try {
    Logger.log('[' + fn + '] state=RUNNING');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allMode = getBoolScriptProperty_(APPROVAL_AMAZON_LV4_EXEMPTION_ALL_PROP, false);
    var exemptions = amazonApprovalLv4LoadExemptions_(ss);

    var targets = [];
    if (allMode) {
      var warn = ui.alert(
        'GTIN免除証跡（全カテゴリ）',
        APPROVAL_AMAZON_LV4_EXEMPTION_ALL_PROP + '=true のため、カテゴリ `*`（全カテゴリ）で記録します。\n' +
          '以後、どのカテゴリの新規カタログでもGTIN免除ゲートを通過します。\n' +
          '免除が使えないカテゴリでもSCアップロードまで気付けません。\n\n' +
          'カテゴリ別に記録する場合はキャンセルし、Property を false にしてください。\n\n続行しますか？',
        ui.ButtonSet.OK_CANCEL
      );
      if (warn !== ui.Button.OK) {
        Logger.log('[' + fn + '] state=DONE cancelled=all_mode_warning');
        return;
      }
      targets = ['*'];
    } else {
      var detected = amazonApprovalLv4DetectCheckboxCategories_(ss);
      if (detected.note) Logger.log('[' + fn + '] detect ' + detected.note);
      if (!detected.categories.length) {
        ui.alert(
          'GTIN免除証跡を記録できません',
          '対象カテゴリを検出できませんでした。\n' +
            '・新規カタログにする子SKU行へ出品CKを付ける\n' +
            '・その親行のカテゴリ列（amazon カテゴリー／カテゴリー）を埋める\n' +
            (detected.note ? '\n検出ログ: ' + detected.note : ''),
          ui.ButtonSet.OK
        );
        Logger.log('[' + fn + '] state=FAILED カテゴリ検出0件');
        return;
      }
      targets = detected.categories;
    }

    var pending = [];
    var already = [];
    var t;
    for (t = 0; t < targets.length; t++) {
      var exists = targets[t] === '*'
        ? amazonApprovalLv4HasStarExemption_(exemptions)
        : amazonApprovalLv4HasGtinExemption_(exemptions, targets[t]);
      if (exists) {
        already.push(targets[t]);
      } else {
        pending.push(targets[t]);
      }
    }
    if (!pending.length) {
      ui.alert(
        'GTIN免除証跡は登録済み',
        '次のカテゴリは有効な証跡があります。追記しません。\n・' + already.join('\n・'),
        ui.ButtonSet.OK
      );
      Logger.log('[' + fn + '] state=DONE 既存証跡のみ already=' + already.join('|'));
      return;
    }

    var evidence = '';
    var evidenceSource = '';
    var suggested = amazonApprovalLv4SuggestGtinEvidence_(ss, pending);
    Logger.log('[' + fn + '] suggest ' + suggested.note);
    if (suggested.text) {
      var useSug = ui.alert(
        'おすすめ証跡（マスタASIN）',
        '同カテゴリのマスタ行から過去成功ASIN候補を生成しました。\n' +
          '（SC申請URL／ケースIDは自動取得できません）\n\n' +
          suggested.text +
          '\n\nOK＝この内容で進む\nキャンセル＝手入力する',
        ui.ButtonSet.OK_CANCEL
      );
      if (useSug === ui.Button.OK) {
        evidence = suggested.text;
        evidenceSource = 'suggested_asin';
      }
    }
    if (!evidence) {
      var ans = ui.prompt(
        'GTIN免除証跡を記録',
        (suggested.text
          ? 'おすすめを使わず手入力します。\n'
          : 'マスタに同カテゴリのASIN候補がありませんでした。\n') +
          '証跡（SC申請URL／ケースID／過去成功ASIN など根拠）を入力してください。\n' +
          '対象カテゴリ: ' + pending.join(' / ') +
          (already.length ? '\n（登録済みで対象外: ' + already.join(' / ') + '）' : ''),
        ui.ButtonSet.OK_CANCEL
      );
      if (ans.getSelectedButton() !== ui.Button.OK) {
        Logger.log('[' + fn + '] state=DONE cancelled=evidence_prompt');
        return;
      }
      evidence = String(ans.getResponseText() || '').trim();
      evidenceSource = 'manual';
    }
    if (!evidence) {
      ui.alert('証跡が空です。承認日と証跡の両方が無い行は無効判定になります。');
      Logger.log('[' + fn + '] state=FAILED 証跡が空');
      return;
    }

    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    var confirm = ui.alert(
      'GTIN免除証跡を追記します',
      'シート: ' + APPROVAL_AMAZON_LV4_LOG_SHEET + '\n' +
        'カテゴリ:\n・' + pending.join('\n・') + '\n' +
        'ブランド: ' + APPROVAL_AMAZON_LV4_BRAND + '\n' +
        '承認日: ' + today + '\n' +
        '証跡: ' + evidence + '\n\n' +
        'この記録は「人間がGTIN免除を確認した」という宣言です。よろしいですか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (confirm !== ui.Button.OK) {
      Logger.log('[' + fn + '] state=DONE cancelled=final_confirm');
      return;
    }

    var added = [];
    for (t = 0; t < pending.length; t++) {
      amazonApprovalLv4AppendLog_(ss, {
        recordType: 'EXEMPTION',
        note: '人間確認により記録（21-⑭' +
          (evidenceSource === 'suggested_asin' ? '・おすすめASIN' : '') + '）',
        category: pending[t],
        brand: APPROVAL_AMAZON_LV4_BRAND,
        exemptionDate: today,
        evidenceUrl: evidence
      });
      added.push(pending[t]);
    }
    Logger.log('[' + fn + '] state=DONE added=' + added.join('|') +
      ' already=' + already.join('|') + ' allMode=' + allMode +
      ' evidenceSource=' + evidenceSource);
    ui.alert(
      'GTIN免除証跡を記録しました',
      '追記=' + added.length + '件\n・' + added.join('\n・') +
        (already.length ? '\n\n登録済みで対象外=' + already.join(' / ') : '') +
        '\n\nこの後 D（新規カタログ）を実行できます。',
      ui.ButtonSet.OK
    );
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    try { ui.alert('GTIN免除証跡の記録に失敗: ' + ((err && err.message) || err)); } catch (e2) {}
  }
}

/** 全カテゴリ（`*`）の有効な免除証跡が既にあるか。 */
function amazonApprovalLv4HasStarExemption_(exemptions) {
  for (var i = 0; i < exemptions.length; i++) {
    var e = exemptions[i];
    if (String(e.brand || '').trim() !== APPROVAL_AMAZON_LV4_BRAND) continue;
    if (!e.exemptionDate || !e.evidenceUrl) continue;
    if (e.category === '*') return true;
  }
  return false;
}

/**
 * レ点新規カタログ対象の親から、GTIN免除ゲートで使われるカテゴリ実値を集める。
 * @return {{categories:Array<string>, note:string}}
 */
function amazonApprovalLv4DetectCheckboxCategories_(ss) {
  var masterCtx = amazonApprovalLv4LoadMasterContext_(ss);
  var inspected = amazonCheckboxMainlineInspect_(masterCtx, {
    includeNew: true,
    includeOffer: false
  });
  if (!inspected.newRows.length) {
    return { categories: [], note: 'レ点子SKU=0（親のみレ点=' + inspected.parentCkOnly + '）' };
  }
  var seenParent = {};
  var seenCat = {};
  var categories = [];
  var missing = [];
  for (var i = 0; i < inspected.newRows.length; i++) {
    var parentSku = String(inspected.newRows[i].parentSku || '').trim();
    if (!parentSku || seenParent[parentSku]) continue;
    seenParent[parentSku] = true;
    var parentRow = amazonApprovalLv4FindMasterRow_(masterCtx, parentSku, '');
    if (!parentRow) {
      missing.push(parentSku + '(親行なし)');
      continue;
    }
    var cat = amazonApprovalLv4ResolveCategory_(masterCtx, parentRow.rowIndex0, parentSku).value;
    if (!cat) {
      missing.push(parentSku + '(カテゴリ空)');
      continue;
    }
    if (!seenCat[cat]) {
      seenCat[cat] = true;
      categories.push(cat);
    }
  }
  return {
    categories: categories,
    note: 'レ点子SKU=' + inspected.newRows.length + ' 親=' + Object.keys(seenParent).length +
      ' カテゴリ=' + categories.length + (missing.length ? ' 未解決=' + missing.join(',') : '')
  };
}

/**
 * マスタ同カテゴリの ASINコード から証跡おすすめ文を作る。
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {Array<string>} categories 記録対象カテゴリ（`*` 可）
 * @return {{text:string, candidates:Array, note:string}}
 */
function amazonApprovalLv4SuggestGtinEvidence_(ss, categories) {
  categories = categories || [];
  var masterCtx = amazonApprovalLv4LoadMasterContext_(ss);
  var iParent = masterCtx.col['親SKU'];
  var iChild = masterCtx.col['子SKU'];
  var iAsin = masterCtx.col['ASINコード'];
  if (iAsin == null || iParent == null) {
    return { text: '', candidates: [], note: 'ASIN列または親SKU列なし' };
  }
  var wantAny = false;
  var want = {};
  var c;
  for (c = 0; c < categories.length; c++) {
    if (categories[c] === '*') wantAny = true;
    else if (categories[c]) want[categories[c]] = true;
  }

  var parentCatCache = {};
  var seenAsin = {};
  var perCat = {};
  var extras = [];
  var maxPerCat = 1;
  var maxTotal = 3;

  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    var row = masterCtx.values[r] || [];
    var asin = String(row[iAsin] == null ? '' : row[iAsin]).trim().toUpperCase();
    if (!amazonApprovalLv4LooksLikeAsin_(asin) || seenAsin[asin]) continue;
    var parentSku = String(row[iParent] == null ? '' : row[iParent]).trim();
    if (!parentSku) continue;

    var cat = parentCatCache[parentSku];
    if (cat === undefined) {
      var parentRow = amazonApprovalLv4FindMasterRow_(masterCtx, parentSku, '');
      if (!parentRow) {
        parentCatCache[parentSku] = '';
        continue;
      }
      cat = amazonApprovalLv4ResolveCategory_(masterCtx, parentRow.rowIndex0, parentSku).value || '';
      parentCatCache[parentSku] = cat;
    }
    if (!cat) continue;
    if (!wantAny && !want[cat]) continue;

    seenAsin[asin] = true;
    var one = {
      asin: asin,
      parentSku: parentSku,
      category: cat,
      url: 'https://www.amazon.co.jp/dp/' + asin
    };
    if (!perCat[cat]) perCat[cat] = [];
    if (perCat[cat].length < maxPerCat) {
      perCat[cat].push(one);
    } else {
      extras.push(one);
    }
  }

  var candidates = [];
  var catKeys = Object.keys(perCat);
  for (c = 0; c < catKeys.length; c++) {
    candidates = candidates.concat(perCat[catKeys[c]]);
  }
  for (c = 0; c < extras.length && candidates.length < maxTotal; c++) {
    candidates.push(extras[c]);
  }
  candidates = candidates.slice(0, maxTotal);
  if (!candidates.length) {
    return {
      text: '',
      candidates: [],
      note: '候補0 wantAny=' + wantAny + ' cats=' + categories.join('|')
    };
  }

  var lines = [];
  for (c = 0; c < candidates.length; c++) {
    var hit = candidates[c];
    lines.push(
      '過去成功ASIN: ' + hit.asin +
        '（親SKU=' + hit.parentSku + ' / カテゴリ=' + hit.category + '）\n' +
        hit.url
    );
  }
  return {
    text: lines.join('\n'),
    candidates: candidates,
    note: '候補=' + candidates.length + ' cats=' + catKeys.join('|')
  };
}

/**
 * メニュー 21-⑮: SC処理サマリを検知して UPLOADED_OK / UPLOAD_FAILED を自動記録（手動実行）。
 * 判定はファイル名のみ。中身（成功件数）は読まない。
 */
function menuApprovalAmazonLv4ScanScSummaries() {
  var fn = 'menuApprovalAmazonLv4ScanScSummaries';
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (eUi) {}
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var res = amazonApprovalLv4ScanScSummaries_(ss, fn);
    if (ui) {
      ui.alert(
        'SC処理サマリの取り込み',
        '記録=' + res.recorded.length + '件\n' +
          (res.recorded.length ? '・' + res.recorded.join('\n・') + '\n' : '') +
          '既に同じ状態=' + res.already.length + '件\n' +
          '対象外ファイル=' + res.ignored + '件\n' +
          'エラー=' + res.errors.length + '件' +
          (res.errors.length ? '\n・' + res.errors.join('\n・') : ''),
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    if (ui) {
      try { ui.alert('SC処理サマリの取り込みに失敗: ' + ((err && err.message) || err)); } catch (e2) {}
    }
  }
}

/** 時間主導トリガー用。UI を使わない。 */
function runApprovalAmazonLv4ScSummaryFromTrigger() {
  var fn = APPROVAL_AMAZON_LV4_SC_SUMMARY_TRIGGER_FN;
  try {
    var ssId = PropertiesService.getScriptProperties()
      .getProperty(APPROVAL_AMAZON_LV4_SC_SUMMARY_SS_PROP);
    var ss = ssId
      ? SpreadsheetApp.openById(String(ssId).trim())
      : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('[' + fn + '] state=FAILED スプレッドシート未解決（' +
        APPROVAL_AMAZON_LV4_SC_SUMMARY_SS_PROP + ' 未設定）');
      return;
    }
    amazonApprovalLv4ScanScSummaries_(ss, fn);
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    amazonApprovalLv4Mail_('【Lv4 Amazon】SCサマリ監視で失敗', String((err && err.message) || err));
  }
}

/**
 * 監視フォルダのSC処理サマリを1件ずつ状態記録し、処理済へ退避する。
 * 記録できなかったファイルは移動しない（人間が直せるように残す）。
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {string} fn 呼び出し元名（ログ用）
 * @return {{recorded:Array<string>, already:Array<string>, ignored:number, errors:Array<string>}}
 */
function amazonApprovalLv4ScanScSummaries_(ss, fn) {
  if (!getBoolScriptProperty_(APPROVAL_AMAZON_LV4_SC_SUMMARY_PROP, false)) {
    throw new Error('SC処理サマリ自動記録は無効です。' + APPROVAL_AMAZON_LV4_SC_SUMMARY_PROP + ' を true にしてください。');
  }
  var folderId = String(PropertiesService.getScriptProperties()
    .getProperty(APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_PROP) || '').trim();
  if (!folderId) {
    throw new Error(APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_PROP + ' が未設定です。サマリ監視フォルダのIDを設定してください。');
  }
  var folder = DriveApp.getFolderById(folderId);
  var doneFolder = amazonApprovalLv4ScSummaryDoneFolder_(folder);

  var out = { recorded: [], already: [], ignored: 0, errors: [] };
  var files = folder.getFiles();
  var seen = 0;
  Logger.log('[' + fn + '] state=RUNNING folderId=' + folderId +
    ' maxFiles=' + APPROVAL_AMAZON_LV4_SC_SUMMARY_MAX_FILES);

  while (files.hasNext() && seen < APPROVAL_AMAZON_LV4_SC_SUMMARY_MAX_FILES) {
    var file = files.next();
    var name = file.getName();
    if (!/-processing-summary/i.test(name)) {
      out.ignored++;
      continue;
    }
    seen++;
    var subBatchId = amazonApprovalLv4ScSummarySubBatchId_(name);
    if (!subBatchId) {
      out.errors.push(name + '（subBatchIdを名前から取れません）');
      Logger.log('[' + fn + '] state=FAILED subBatchId解析不可 file=' + name);
      continue;
    }
    var status = amazonApprovalLv4ScSummaryDecideStatus_(name);
    var latest = amazonApprovalLv4SubBatchLatestStatus_(ss, subBatchId);
    if (latest === status) {
      out.already.push(subBatchId + '=' + status);
      Logger.log('[' + fn + '] already subBatchId=' + subBatchId + ' status=' + status);
      amazonApprovalLv4ScSummaryMove_(file, doneFolder, fn);
      continue;
    }
    try {
      amazonApprovalLv4MarkStatus_(ss, subBatchId, status, 'SC処理サマリ検知（21-⑮・ファイル名判定）: ' + name);
    } catch (eMark) {
      out.errors.push(subBatchId + '（' + ((eMark && eMark.message) || eMark) + '）');
      Logger.log('[' + fn + '] state=FAILED mark subBatchId=' + subBatchId +
        ' ' + ((eMark && eMark.message) || eMark));
      continue;
    }
    out.recorded.push(subBatchId + '=' + status);
    Logger.log('[' + fn + '] state=DONE recorded subBatchId=' + subBatchId +
      ' status=' + status + ' file=' + name);
    amazonApprovalLv4ScSummaryMove_(file, doneFolder, fn);
  }

  Logger.log('[' + fn + '] state=DONE recorded=' + out.recorded.length +
    ' already=' + out.already.length + ' ignored=' + out.ignored +
    ' errors=' + out.errors.length);
  if (out.recorded.length || out.errors.length) {
    amazonApprovalLv4Mail_(
      '【Lv4 Amazon】SCサマリ自動記録 ' + out.recorded.length + '件',
      '記録:\n' + (out.recorded.join('\n') || 'なし') +
        '\n\nエラー:\n' + (out.errors.join('\n') || 'なし') +
        '\n\n※ファイル名のみで判定しています。掲載完了の断定には実機確認が必要です。'
    );
  }
  return out;
}

/**
 * `{subBatchId}_PACKAGED_…-processing-summary.xlsm` から subBatchId を取り出す。
 * subBatchId 自体に `_` を含むため `_PACKAGED_` の直前までを採用する。
 */
function amazonApprovalLv4ScSummarySubBatchId_(fileName) {
  var m = String(fileName || '').match(/^(.+?)_PACKAGED_/);
  return m ? String(m[1]).trim() : '';
}

/** ファイル名に失敗マーカーがあれば UPLOAD_FAILED。既定は UPLOADED_OK。 */
function amazonApprovalLv4ScSummaryDecideStatus_(fileName) {
  return /(_NG|UPLOAD_FAILED)/i.test(String(fileName || ''))
    ? 'UPLOAD_FAILED'
    : 'UPLOADED_OK';
}

/** 監視フォルダ直下の処理済フォルダ（無ければ作成）。 */
function amazonApprovalLv4ScSummaryDoneFolder_(folder) {
  var it = folder.getFoldersByName(APPROVAL_AMAZON_LV4_SC_SUMMARY_DONE_FOLDER);
  return it.hasNext() ? it.next() : folder.createFolder(APPROVAL_AMAZON_LV4_SC_SUMMARY_DONE_FOLDER);
}

/** 処理済へ退避。失敗しても記録は有効なので停止させない（trashしない）。 */
function amazonApprovalLv4ScSummaryMove_(file, doneFolder, fn) {
  try {
    file.moveTo(doneFolder);
  } catch (e) {
    Logger.log('[' + fn + '] WARN 処理済への移動に失敗 file=' + file.getName() +
      ' ' + ((e && e.message) || e) + '（二重記録は状態一致チェックで抑止）');
  }
}

/** 同一 subBatchId の RUN 行のうち、最後に記録された status。 */
function amazonApprovalLv4SubBatchLatestStatus_(ss, subBatchId) {
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return '';
  var data = sh.getDataRange().getValues();
  var latest = '';
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) !== 'RUN') continue;
    if (String(data[r][3]) !== subBatchId) continue;
    latest = String(data[r][7] || '');
  }
  return latest;
}

/** メニュー 21-⑯: SCサマリ監視トリガーを設置 */
function menuApprovalAmazonLv4InstallScSummaryTrigger() {
  var fn = 'menuApprovalAmazonLv4InstallScSummaryTrigger';
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (eUi) {}
  try {
    var props = PropertiesService.getScriptProperties();
    if (!getBoolScriptProperty_(APPROVAL_AMAZON_LV4_SC_SUMMARY_PROP, false)) {
      throw new Error(APPROVAL_AMAZON_LV4_SC_SUMMARY_PROP + ' が false です。先に true にしてください。');
    }
    if (!String(props.getProperty(APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_PROP) || '').trim()) {
      throw new Error(APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_PROP + ' が未設定です。');
    }
    var minutes = amazonApprovalLv4ScSummaryInterval_();
    amazonApprovalLv4ScSummaryDeleteTriggers_();
    var builder = ScriptApp.newTrigger(APPROVAL_AMAZON_LV4_SC_SUMMARY_TRIGGER_FN).timeBased();
    if (minutes >= 60) builder.everyHours(1);
    else builder.everyMinutes(minutes);
    builder.create();
    props.setProperty(APPROVAL_AMAZON_LV4_SC_SUMMARY_SS_PROP,
      SpreadsheetApp.getActiveSpreadsheet().getId());
    Logger.log('[' + fn + '] state=DONE intervalMin=' + minutes);
    if (ui) {
      ui.alert(
        'SCサマリ監視トリガーを設置しました',
        '間隔=' + (minutes >= 60 ? '1時間' : minutes + '分') + '\n' +
          'SCからダウンロードしたサマリを監視フォルダへ置くだけで、\n' +
          'UPLOADED_OK が自動記録されます（メニュー操作不要）。\n' +
          '※ファイル名のみで判定します。',
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    if (ui) {
      try { ui.alert(String((err && err.message) || err)); } catch (e2) {}
    }
  }
}

/** メニュー 21-⑰: SCサマリ監視トリガーを削除 */
function menuApprovalAmazonLv4RemoveScSummaryTrigger() {
  var fn = 'menuApprovalAmazonLv4RemoveScSummaryTrigger';
  try {
    var removed = amazonApprovalLv4ScSummaryDeleteTriggers_();
    Logger.log('[' + fn + '] state=DONE removed=' + removed);
    SpreadsheetApp.getUi().alert('SCサマリ監視トリガーを削除しました（削除数=' + removed + '）');
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert(String((err && err.message) || err)); } catch (e2) {}
  }
}

/** everyMinutes が受け付ける値へ丸める（5/10/15/30、60以上は1時間）。 */
function amazonApprovalLv4ScSummaryInterval_() {
  var raw = Number(PropertiesService.getScriptProperties()
    .getProperty(APPROVAL_AMAZON_LV4_SC_SUMMARY_INTERVAL_PROP));
  if (!isFinite(raw) || raw <= 0) return 15;
  var allowed = [5, 10, 15, 30];
  for (var i = 0; i < allowed.length; i++) {
    if (raw <= allowed[i]) return allowed[i];
  }
  return 60;
}

function amazonApprovalLv4ScSummaryDeleteTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === APPROVAL_AMAZON_LV4_SC_SUMMARY_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  return removed;
}

/** メニュー 21-②: レジューム状態クリア */
function menuApprovalAmazonLv4ClearState() {
  try {
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_AMAZON_LV4_STATE_PROP);
    amazonApprovalLv4DeleteTriggers_();
    SpreadsheetApp.getUi().alert('Lv4 実行状態をクリアしました。');
  } catch (e) {
    try { SpreadsheetApp.getUi().alert(String(e && e.message || e)); } catch (e2) {}
  }
}

/**
 * メニュー 21-③: アップロード成功を記録（UPLOADED_OK）
 */
function menuApprovalAmazonLv4MarkUploadedOk() {
  var fn = 'menuApprovalAmazonLv4MarkUploadedOk';
  try {
    var ui = SpreadsheetApp.getUi();
    var ans = ui.prompt(
      'Lv4 アップロード成功を記録',
      'subBatchId を入力（例: {batchId}_B1）',
      ui.ButtonSet.OK_CANCEL
    );
    if (ans.getSelectedButton() !== ui.Button.OK) return;
    var subBatchId = String(ans.getResponseText() || '').trim();
    if (!subBatchId) {
      ui.alert('subBatchId が空です');
      return;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var updated = amazonApprovalLv4MarkStatus_(ss, subBatchId, 'UPLOADED_OK', '');
    Logger.log('[' + fn + '] state=DONE subBatchId=' + subBatchId + ' updated=' + updated);
    ui.alert('UPLOADED_OK を記録しました（更新行=' + updated + '）\n※在庫0/1の実機見え方未検証なら「掲載完了」と断言しないでください');
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert(String(err && err.message || err)); } catch (e2) {}
  }
}

/**
 * メニュー 21-④: アップロード失敗を記録（UPLOAD_FAILED）→ 同一 subBatchId で再生成可
 */
function menuApprovalAmazonLv4MarkUploadFailed() {
  var fn = 'menuApprovalAmazonLv4MarkUploadFailed';
  try {
    var ui = SpreadsheetApp.getUi();
    var ans = ui.prompt(
      'Lv4 アップロード失敗を記録',
      'subBatchId と理由（任意）を「subBatchId|理由」形式で入力',
      ui.ButtonSet.OK_CANCEL
    );
    if (ans.getSelectedButton() !== ui.Button.OK) return;
    var raw = String(ans.getResponseText() || '').trim();
    var parts = raw.split('|');
    var subBatchId = String(parts[0] || '').trim();
    var note = parts.length > 1 ? String(parts.slice(1).join('|')).trim() : '';
    if (!subBatchId) {
      ui.alert('subBatchId が空です');
      return;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var updated = amazonApprovalLv4MarkStatus_(ss, subBatchId, 'UPLOAD_FAILED', note);
    Logger.log('[' + fn + '] state=DONE subBatchId=' + subBatchId + ' updated=' + updated);
    ui.alert('UPLOAD_FAILED を記録しました（更新行=' + updated + '）。同一 subBatchId で再生成できます。');
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    try { SpreadsheetApp.getUi().alert(String(err && err.message || err)); } catch (e2) {}
  }
}

/** トリガー再開 */
function runApprovalAmazonLv4FromTrigger() {
  var fn = APPROVAL_AMAZON_LV4_TRIGGER_FN;
  var stateJson = PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_STATE_PROP);
  if (!stateJson) return;
  amazonApprovalLv4DeleteTriggers_();
  var state;
  try {
    state = JSON.parse(stateJson);
  } catch (e) {
    return;
  }
  var runCount = Number(state.triggerRunCount || 0) + 1;
  state.triggerRunCount = runCount;
  PropertiesService.getScriptProperties().setProperty(APPROVAL_AMAZON_LV4_STATE_PROP, JSON.stringify(state));
  if (runCount > APPROVAL_AMAZON_LV4_MAX_TRIGGER_RUNS) {
    Logger.log('[' + fn + '] state=FAILED maxTriggerRuns runId=' + state.runId);
    amazonApprovalLv4Mail_('【Lv4 Amazon】自動再開上限', 'runId=' + state.runId + ' batchId=' + state.batchId);
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_AMAZON_LV4_STATE_PROP);
    return;
  }
  if (!getBoolScriptProperty_(APPROVAL_AMAZON_LV4_PROP, false)) {
    Logger.log('[' + fn + '] state=FAILED disabled');
    return;
  }
  var track = amazonApprovalLv4NormalizeTrack_(state.track ||
    PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_TRACK_PROP));
  if (!track) {
    Logger.log('[' + fn + '] state=FAILED trackUnset');
    return;
  }
  var ss = SpreadsheetApp.openById(state.spreadsheetId);
  Logger.log('[' + fn + '] state=RUNNING runId=' + state.runId + ' resume remaining (index0 after doneParents filter)');
  try {
    // doneParents 除外後にサブバッチ再採番するため startIndex は常に 0
    amazonApprovalLv4Run_(ss, state.runId, 0, state, track, state.source || 'approved', {
      includeOffer: !!state.includeOffer
    });
  } catch (err) {
    Logger.log('[' + fn + '] state=FAILED ' + ((err && err.message) || err));
    amazonApprovalLv4Mail_('【Lv4 Amazon】再開失敗', 'runId=' + state.runId + '\n' + ((err && err.message) || err));
  }
}

/**
 * @param {Spreadsheet} ss
 * @param {string} runId
 * @param {number} startSubBatchIndex 再開時は無視し0から（doneParents除外後に再採番するため）
 * @param {Object|null} resumeState
 * @param {string} track
 * @param {string=} source approved | child_ck
 * @param {{includeOffer?:boolean}=} runOpts
 * @return {Object}
 */
function amazonApprovalLv4Run_(ss, runId, startSubBatchIndex, resumeState, track, source, runOpts) {
  var fn = 'amazonApprovalLv4Run_';
  var startedAt = Date.now();
  runOpts = runOpts || {};
  source = String(source || 'approved') === 'child_ck' ? 'child_ck' : 'approved';
  var loaded = source === 'child_ck'
    ? amazonApprovalLv4LoadCheckboxNewAmazon_(ss, { includeOffer: !!runOpts.includeOffer })
    : ((typeof approvalQueueGetLatestApprovedAmazon_ === 'function')
        ? approvalQueueGetLatestApprovedAmazon_()
        : { found: false });
  if (!loaded.found || !loaded.lines || !loaded.lines.length) {
    throw new Error(source === 'child_ck'
      ? 'レ点付きの新規カタログ子SKUがありません。子SKUレ点を確認してください。'
      : 'APPROVED の Amazon 明細がありません。先に承認キューで amazon 親＋子を承認①してください。');
  }
  var batch = loaded.batch;
  var batchId = batch.batchId;
  if (resumeState && resumeState.batchId && String(resumeState.batchId) !== String(batchId)) {
    throw new Error(
      '再開stateのbatchId不一致: state=' + resumeState.batchId + ' latestAPPROVED=' + batchId +
      '。21-②で状態クリアしてから再実行してください。'
    );
  }
  var inventoryMode = String(batch.inventoryMode || 'ZERO').toUpperCase() === 'ONE' ? 'ONE' : 'ZERO';
  var stockOut = inventoryMode === 'ONE' ? 1 : 0;
  var shipping = amazonApprovalLv4ShippingTemplate_();
  var skipExport = getBoolScriptProperty_(APPROVAL_AMAZON_LV4_SKIP_EXPORT_PROP, false);

  amazonApprovalLv4EnsureLogSheet_(ss);
  var masterCtx = amazonApprovalLv4LoadMasterContext_(ss);
  var exemptions = amazonApprovalLv4LoadExemptions_(ss);

  // 冪等: GENERATED/UPLOADED_OK/PACKAGED 済み親は除外。DRY_RUN・UPLOAD_FAILED はブロックしない
  // 人間レ点本線はレ点構成が変わってbatchIdが変わっても、親SKU横断で成功済みを再生成しない。
  var blockedMap = amazonApprovalLv4LoadBlockedParents_(
    ss,
    source === 'child_ck' ? '' : batchId,
    track
  );
  var excludeParents = [];
  if (resumeState && resumeState.doneParents) {
    excludeParents = resumeState.doneParents.slice();
  }
  var blockedKeys = Object.keys(blockedMap);
  for (var bi = 0; bi < blockedKeys.length; bi++) {
    excludeParents.push(blockedKeys[bi]);
  }
  // 冪等除外は recordType=SKIP（RUN の latest を汚染しない）
  var idempotentSkipped = [];
  for (var ik = 0; ik < blockedKeys.length; ik++) {
    idempotentSkipped.push({
      parentSku: blockedKeys[ik],
      reason: 'SKIPPED_IDEMPOTENT',
      detail: 'latest=' + blockedMap[blockedKeys[ik]] + '（UPLOAD_FAILEDのみ再生成可）'
    });
  }
  amazonApprovalLv4LogSkippedParents_(ss, runId, batchId, track, idempotentSkipped);

  // 在庫>0での続行はDレ点新規（child_ck）限定。承認①経路は従来どおりスキップ。
  var ckAllowInStock = source === 'child_ck' &&
    getBoolScriptProperty_(APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK_PROP, true);

  var resolved = amazonApprovalLv4ResolveParents_(
    masterCtx, loaded.lines, track, excludeParents, exemptions,
    { allowInStock: ckAllowInStock }
  );
  amazonApprovalLv4LogSkippedParents_(ss, runId, batchId, track, resolved.skipped);

  Logger.log('[' + fn + '] runId=' + runId + ' batchId=' + batchId +
    ' track=' + track + ' parents=' + resolved.parents.length +
    ' skipped=' + resolved.skipped.length +
    ' blockedIdempotent=' + blockedKeys.length +
    ' inventoryMode=' + inventoryMode + ' source=' + source +
    ' ckAllowInStock=' + ckAllowInStock +
    ' skipExport=' + skipExport);

  var perSub = amazonApprovalLv4ParentsPerSub_();
  var subBatches = amazonApprovalLv4BuildSubBatches_(resolved.parents, perSub);
  if (!subBatches.length) {
    return {
      batchId: batchId,
      subBatchesDone: 0,
      parentsDone: 0,
      skipped: resolved.skipped.length,
      willResume: false,
      message: '実行対象の親がありません（スキップ/冪等除外）: ' +
        amazonApprovalLv4SkipSummary_(resolved.skipped) +
        ';idempotentBlocked=' + blockedKeys.length
    };
  }

  // 再開時: 残り親は index0 から。subBatchId はシート上の最大番号から単調増加（衝突防止）
  var loopStart = 0;
  var doneParents = (resumeState && resumeState.doneParents) ? resumeState.doneParents.slice() : [];
  var subBatchesDone = Number(resumeState && resumeState.subBatchesDone) || 0;
  var trackTag = track === 'A' ? 'A' : (track === 'BOTH' ? 'X' : 'B');
  var nextSubSeq = amazonApprovalLv4NextSubBatchSeq_(ss, batchId, trackTag, resumeState);
  var willResume = false;
  var message = '';
  var folder = skipExport ? null : amazonApprovalLv4GetOrCreateFolder_(ss);

  for (var s = loopStart; s < subBatches.length; s++) {
    if ((Date.now() - startedAt) > APPROVAL_AMAZON_LV4_TIME_MS) {
      amazonApprovalLv4SaveState_({
        spreadsheetId: ss.getId(),
        runId: runId,
        batchId: batchId,
        track: track,
        source: source,
        includeOffer: !!runOpts.includeOffer,
        inventoryMode: inventoryMode,
        nextSubBatchIndex: s,
        subBatchesDone: subBatchesDone,
        nextSubSeq: nextSubSeq,
        doneParents: doneParents,
        triggerRunCount: resumeState ? resumeState.triggerRunCount : 0
      });
      amazonApprovalLv4SetTrigger_();
      willResume = true;
      message = '時間予算のため中断。残り親から再開予定（subBatchIdは単調増加）';
      Logger.log('[' + fn + '] state=RETRYING runId=' + runId + ' ' + message);
      break;
    }

    var sub = subBatches[s];
    var subBatchId = batchId + '_' + trackTag + nextSubSeq;
    nextSubSeq++;
    Logger.log('[' + fn + '] state=RUNNING runId=' + runId + ' batchId=' + batchId +
      ' subBatchId=' + subBatchId + ' parents=' + sub.length + ' track=' + track);

    try {
      var built = amazonApprovalLv4BuildRows_(masterCtx, sub, track, stockOut, shipping);
      // Build 内で想定外スキップがあればログ（Resolve 済みなら通常0件）
      if (built.skipped && built.skipped.length) {
        amazonApprovalLv4LogSkippedParents_(ss, runId, batchId, track, built.skipped);
      }
      if (!built.rows.length || !built.okParents.length) {
        amazonApprovalLv4AppendLog_(ss, {
          recordType: 'RUN',
          runId: runId,
          batchId: batchId,
          subBatchId: subBatchId,
          track: track,
          parentSkus: amazonApprovalLv4ParentSkuList_(sub),
          childCount: 0,
          status: 'SKIPPED_EMPTY',
          fileName: '',
          fileUrl: '',
          note: '検証後に出力行なし（親はdoneにしない）'
        });
        continue;
      }

      var outStatus = skipExport ? 'DRY_RUN' : 'GENERATED';
      var fileName = '';
      var fileUrl = '';
      if (!skipExport) {
        var csv = amazonApprovalLv4RowsToCsv_(built.headers, built.rows);
        var blob = Utilities.newBlob(csv, 'text/csv', subBatchId + '_GENERATED.csv');
        amazonApprovalLv4ArchiveSameName_(folder, blob.getName());
        var file = folder.createFile(blob);
        fileName = file.getName();
        fileUrl = file.getUrl();
        var meta = {
          runId: runId,
          batchId: batchId,
          subBatchId: subBatchId,
          track: track,
          inventoryMode: inventoryMode,
          stockOut: stockOut,
          shippingTemplate: shipping,
          brand: APPROVAL_AMAZON_LV4_BRAND,
          parents: amazonApprovalLv4ParentSkuList_(built.okParents),
          rowCount: built.rows.length,
          note: 'GAS GENERATED only. PACKAGED=.xlsm is local. Do not wipe master JAN/stock.'
        };
        var metaName = subBatchId + '_GENERATED.meta.json';
        amazonApprovalLv4ArchiveSameName_(folder, metaName);
        folder.createFile(Utilities.newBlob(
          JSON.stringify(meta, null, 2),
          'application/json',
          metaName
        ));
      } else {
        fileName = '(SKIP_EXPORT)';
        fileUrl = '';
      }

      amazonApprovalLv4AppendLog_(ss, {
        recordType: 'RUN',
        runId: runId,
        batchId: batchId,
        subBatchId: subBatchId,
        track: track,
        parentSkus: amazonApprovalLv4ParentSkuList_(built.okParents),
        childCount: built.childCount,
        status: outStatus,
        fileName: fileName,
        fileUrl: fileUrl,
        note: skipExport ? 'DRY_RUN（冪等ブロック対象外。本番はSKIP_EXPORT=falseでGENERATED）' : ''
      });

      for (var p = 0; p < built.okParents.length; p++) {
        doneParents.push(built.okParents[p].parentSku);
      }
      subBatchesDone++;
      Logger.log('[' + fn + '] state=DONE subBatchId=' + subBatchId +
        ' status=' + outStatus +
        ' fileName=' + fileName + ' rows=' + built.rows.length +
        ' okParents=' + built.okParents.length);
    } catch (subErr) {
      Logger.log('[' + fn + '] state=FAILED subBatchId=' + subBatchId + ' ' + ((subErr && subErr.message) || subErr));
      amazonApprovalLv4AppendLog_(ss, {
        recordType: 'RUN',
        runId: runId,
        batchId: batchId,
        subBatchId: subBatchId,
        track: track,
        parentSkus: amazonApprovalLv4ParentSkuList_(sub),
        childCount: 0,
        status: 'FAILED',
        fileName: '',
        fileUrl: '',
        note: String((subErr && subErr.message) || subErr)
      });
      amazonApprovalLv4Mail_(
        '【Lv4 Amazon】サブバッチ失敗',
        'runId=' + runId + '\nsubBatchId=' + subBatchId + '\n' + ((subErr && subErr.message) || subErr)
      );
      throw subErr;
    }
  }

  if (!willResume) {
    PropertiesService.getScriptProperties().deleteProperty(APPROVAL_AMAZON_LV4_STATE_PROP);
    amazonApprovalLv4DeleteTriggers_();
    message = message || (
      (skipExport
        ? 'DRY_RUN完了（本番GENERATEDはSKIP_EXPORT=falseで再実行可）。'
        : 'GENERATED完了。次はローカルでPACKAGED→SC手動UP→21-③。') +
      'スキップ=' + amazonApprovalLv4SkipSummary_(resolved.skipped)
    );
  }

  return {
    batchId: batchId,
    subBatchesDone: subBatchesDone,
    parentsDone: doneParents.length,
    skipped: resolved.skipped.length,
    willResume: willResume,
    message: message
  };
}

/**
 * @param {string|null} raw
 * @return {string} A|B|BOTH|''
 */
function amazonApprovalLv4NormalizeTrack_(raw) {
  var t = String(raw || '').trim().toUpperCase();
  if (t === 'A' || t === 'B' || t === 'BOTH') return t;
  return '';
}

function amazonApprovalLv4ShippingTemplate_() {
  var v = PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_SHIPPING_TEMPLATE_PROP);
  v = String(v || '').trim();
  return v || APPROVAL_AMAZON_LV4_DEFAULT_SHIPPING;
}

function amazonApprovalLv4ParentsPerSub_() {
  var n = Number(PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_PARENTS_PER_SUB_PROP));
  if (!n || isNaN(n) || n < 1) return 5;
  return Math.floor(n);
}

/** @return {{sheet:Sheet, values:Array, headerRowIdx:number, col:Object}} */
function amazonApprovalLv4LoadMasterContext_(ss) {
  var masterName = (typeof MASTER_SHEET_NAME !== 'undefined')
    ? MASTER_SHEET_NAME
    : ((typeof SHEET_NAME_MASTER !== 'undefined') ? SHEET_NAME_MASTER : '▼商品マスタ(人間作業用)');
  var sheet = ss.getSheetByName(masterName);
  if (!sheet) throw new Error('マスタシートが見つかりません: ' + masterName);
  var values = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  var limit = Math.min(values.length, 25);
  for (var r = 0; r < limit; r++) {
    var row = values[r] || [];
    if (row.indexOf('親SKU') !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error('マスタヘッダー（親SKU）が見つかりません');
  var headers = values[headerRowIdx];
  var col = {};
  for (var c = 0; c < headers.length; c++) {
    var key = String(headers[c] || '').trim();
    if (key && col[key] == null) col[key] = c;
  }
  if (col['親SKU'] == null || col['子SKU'] == null) {
    throw new Error('必須列がありません（親SKU / 子SKU）');
  }
  Logger.log('[amazonApprovalLv4LoadMasterContext_] headerRow1Based=' + (headerRowIdx + 1) +
    ' col販売価格amazon=' + (col['販売価格amazon'] != null ? col['販売価格amazon'] : 'MISSING') +
    ' col在庫数=' + (col['在庫数'] != null ? col['在庫数'] : 'MISSING'));
  return { sheet: sheet, values: values, headerRowIdx: headerRowIdx, col: col };
}

/**
 * Dレ点本線の出品方式を正規化する（互換・ログ用）。振り分けの正はD選択。
 * @return {string} new | offer | ''
 */
function amazonCheckboxMainlineRouteFromShipping_(raw) {
  var v = String(raw == null ? '' : raw).trim().toLowerCase();
  if (v === 'fba' || v === '自己発送') return 'new';
  if (v === '相乗りfba' || v === '相乗り自己発' || v === '相乗り自己発送') return 'offer';
  return '';
}

/**
 * レ点付き子SKUを新規／相乗りの対象へ分類する。
 * X列は新規SKU式用のため、振り分けには使わない。
 * - 相乗りのみ: レ点子SKUすべて → offer
 * - 新規のみ: レ点子SKUすべて → new
 * - 両方: 同じレ点子SKUを新規と相乗りの両方へ出す（相互排他しない）
 *   新規=子SKU／相乗り=Amazon相乗りSKU。相乗りは後段でN列ASIN必須。
 * @param {Object} masterCtx
 * @param {{includeNew?:boolean, includeOffer?:boolean}=} opts
 * @return {{newRows:Array, offerRows:Array, unknown:Array, parentCkOnly:number}}
 */
function amazonCheckboxMainlineInspect_(masterCtx, opts) {
  opts = opts || {};
  var includeNew = opts.includeNew !== false;
  var includeOffer = opts.includeOffer !== false;
  var ckName = (typeof CHECKBOX_HEADER_NAME !== 'undefined') ? CHECKBOX_HEADER_NAME : '出品CK';
  var iCk = masterCtx.col[ckName];
  var iChild = masterCtx.col['子SKU'];
  var iParent = masterCtx.col['親SKU'];
  var iShipping = masterCtx.col[AMAZON_CHECKBOX_SHIPPING_HEADER_];
  var iAsin = masterCtx.col['ASINコード'];
  if (iCk == null || iChild == null || iParent == null) {
    throw new Error('Dレ点本線の必須列がありません（' + ckName + '／親SKU／子SKU）');
  }
  var out = { newRows: [], offerRows: [], unknown: [], parentCkOnly: 0 };
  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    var row = masterCtx.values[r] || [];
    var checked = row[iCk] === true || String(row[iCk] || '').trim().toUpperCase() === 'TRUE';
    if (!checked) continue;
    var childSku = String(row[iChild] == null ? '' : row[iChild]).trim();
    if (!childSku) {
      out.parentCkOnly++;
      continue;
    }
    var parentSku = String(row[iParent] == null ? '' : row[iParent]).trim();
    var asinRaw = iAsin == null ? '' : String(row[iAsin] == null ? '' : row[iAsin]).trim();
    var hasAsin = amazonApprovalLv4LooksLikeAsin_(asinRaw);
    if (!hasAsin && parentSku) {
      var parentRow = typeof amazonSpapiExportFindParentRow_ === 'function'
        ? amazonSpapiExportFindParentRow_(masterCtx, parentSku)
        : null;
      if (parentRow != null && iAsin != null) {
        var pAsin = String(masterCtx.values[parentRow][iAsin] == null
          ? '' : masterCtx.values[parentRow][iAsin]).trim();
        hasAsin = amazonApprovalLv4LooksLikeAsin_(pAsin);
      }
    }
    var one = {
      rowIndex0: r,
      row1: r + 1,
      parentSku: parentSku,
      childSku: childSku,
      shipping: iShipping == null ? '' : String(row[iShipping] == null ? '' : row[iShipping]).trim(),
      hasAsin: hasAsin
    };
    // Dで選んだ経路へ独立に載せる（同じ行を新規＋相乗りへ同時出品可）
    if (includeNew) out.newRows.push(one);
    if (includeOffer) out.offerRows.push(one);
  }
  return out;
}

/** 人間レ点の新規カタログ行を承認①互換の一時明細へ変換する。 */
function amazonApprovalLv4LoadCheckboxNewAmazon_(ss, opts) {
  opts = opts || {};
  var masterCtx = amazonApprovalLv4LoadMasterContext_(ss);
  var inspected = amazonCheckboxMainlineInspect_(masterCtx, {
    includeNew: true,
    includeOffer: !!opts.includeOffer
  });
  if (!inspected.newRows.length) return { found: false, lines: [] };

  var lines = [];
  var parentSeen = {};
  var ids = [];
  for (var i = 0; i < inspected.newRows.length; i++) {
    var one = inspected.newRows[i];
    if (!one.parentSku) {
      throw new Error('新規カタログ対象の親SKUが空です。行' + one.row1);
    }
    if (!parentSeen[one.parentSku]) {
      lines.push({
        mall: 'amazon',
        lineStatus: 'APPROVED',
        parentSku: one.parentSku,
        childSku: ''
      });
      parentSeen[one.parentSku] = true;
    }
    lines.push({
      mall: 'amazon',
      lineStatus: 'APPROVED',
      parentSku: one.parentSku,
      childSku: one.childSku
    });
    ids.push(one.parentSku + '/' + one.childSku);
  }
  ids.sort();
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    ids.join('\n'),
    Utilities.Charset.UTF_8
  );
  var hex = '';
  for (var d = 0; d < digest.length; d++) {
    hex += ('0' + ((digest[d] + 256) % 256).toString(16)).slice(-2);
  }
  return {
    found: true,
    batch: { batchId: 'CK_' + hex.substring(0, 12), inventoryMode: 'ZERO' },
    lines: lines
  };
}

/**
 * 承認 amazon 行を親単位にまとめ、TRACK/在庫/免除でフィルタ。
 * @param {{allowInStock?:boolean}=} opts allowInStock=true で在庫>0でも新規を続行（Dレ点新規のみ）
 * @return {{parents:Array, skipped:Array}}
 */
function amazonApprovalLv4ResolveParents_(masterCtx, lines, track, doneParents, exemptions, opts) {
  opts = opts || {};
  var allowInStock = !!opts.allowInStock;
  var doneMap = {};
  if (doneParents) {
    for (var d = 0; d < doneParents.length; d++) doneMap[String(doneParents[d])] = true;
  }
  var byParent = {};
  for (var i = 0; i < lines.length; i++) {
    var L = lines[i];
    if (String(L.mall) !== 'amazon' || String(L.lineStatus) !== 'APPROVED') continue;
    var parentSku = String(L.parentSku || '').trim();
    if (!parentSku) continue;
    if (!byParent[parentSku]) {
      byParent[parentSku] = { parentSku: parentSku, parentLine: null, childLines: [] };
    }
    var childSku = String(L.childSku || '').trim();
    if (!childSku) {
      byParent[parentSku].parentLine = L;
    } else {
      byParent[parentSku].childLines.push(L);
    }
  }

  var parents = [];
  var skipped = [];
  exemptions = exemptions || [];

  var keys = Object.keys(byParent);
  for (var k = 0; k < keys.length; k++) {
    var g = byParent[keys[k]];
    if (doneMap[g.parentSku]) continue;

    if (track === 'B' || track === 'BOTH') {
      // M1/B: 親承認＋子1件以上必須。単品親は対象外
      if (!g.parentLine) {
        skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_INCOMPLETE_VARIATION', detail: '親行未承認' });
        continue;
      }
      if (!g.childLines.length) {
        skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_INCOMPLETE_VARIATION', detail: '承認済み子0' });
        continue;
      }
    }

    var masterParent = amazonApprovalLv4FindMasterRow_(masterCtx, g.parentSku, '');
    if (!masterParent) {
      skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_ORPHAN', detail: '親がマスタに無い' });
      continue;
    }

    // 販売中スキップ: 親または承認済み子のいずれか在庫>0
    // Dレ点新規は別カタログ（ノーブランドセット）を作るため、在庫>0でも続行できる。
    // マスタ在庫は読取のみで、GENERATED の送信在庫は inventoryMode に従い常に 0/1。
    var stockHit = amazonApprovalLv4AnyInStock_(masterCtx, g);
    if (stockHit) {
      if (!allowInStock) {
        skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_IN_STOCK', detail: stockHit });
        continue;
      }
      Logger.log('[amazonApprovalLv4ResolveParents_] allowInStock parent=' + g.parentSku +
        ' detail=' + stockHit + '（マスタ在庫は非改変・送信在庫はinventoryMode準拠）');
    }

    var resolvedTrack = track;
    if (track === 'BOTH') {
      var asin = amazonApprovalLv4Cell_(masterCtx, masterParent.rowIndex0, 'ASINコード');
      if (asin) {
        resolvedTrack = 'A';
      } else {
        resolvedTrack = 'B';
      }
    }

    if (resolvedTrack === 'B') {
      var catResolved = amazonApprovalLv4ResolveCategory_(masterCtx, masterParent.rowIndex0, g.parentSku);
      var cat = catResolved.value;
      if (!cat) {
        skipped.push({
          parentSku: g.parentSku,
          reason: 'SKIPPED_NEED_HUMAN',
          detail: 'amazonカテゴリ空（GTIN証跡照合不可） sourceTried=' + catResolved.sourceTried
        });
        continue;
      }
      if (!amazonApprovalLv4HasGtinExemption_(exemptions, cat)) {
        skipped.push({
          parentSku: g.parentSku,
          reason: 'SKIPPED_GTIN_EXEMPTION',
          detail: 'カテゴリ未証跡:' + cat + ' via=' + catResolved.source
        });
        continue;
      }
      var mainImg = amazonApprovalLv4ResolveMainImageUrl_(masterCtx, masterParent.rowIndex0);
      if (!mainImg) {
        var anyChildImg = false;
        for (var ci0 = 0; ci0 < g.childLines.length; ci0++) {
          var mr0 = amazonApprovalLv4FindMasterRow_(masterCtx, g.parentSku, String(g.childLines[ci0].childSku || '').trim());
          if (mr0 && amazonApprovalLv4ResolveMainImageUrl_(masterCtx, mr0.rowIndex0)) {
            anyChildImg = true;
            break;
          }
        }
        if (!anyChildImg) {
          skipped.push({
            parentSku: g.parentSku,
            reason: 'SKIPPED_NEED_HUMAN',
            detail: 'メイン画像URL空'
          });
          continue;
        }
      }
    }

    if (resolvedTrack === 'A') {
      // ASIN: 親または承認済み子のいずれか（ASINコード／競合店ASIN／URL）
      var asinProbe = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, masterParent.rowIndex0);
      if (!asinProbe) {
        for (var ai = 0; ai < g.childLines.length && !asinProbe; ai++) {
          var mrA = amazonApprovalLv4FindMasterRow_(
            masterCtx, g.parentSku, String(g.childLines[ai].childSku || '').trim()
          );
          if (mrA) asinProbe = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, mrA.rowIndex0);
        }
      }
      if (!asinProbe) {
        skipped.push({
          parentSku: g.parentSku,
          reason: 'SKIPPED_NEED_HUMAN',
          detail: 'AトラックだがASIN無し（ASINコード／競合店ASINコード／競合URL）'
        });
        continue;
      }
      var gateMode = String(
        PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_BRAND_GATE_MODE_PROP) || ''
      ).trim().toLowerCase();
      if (gateMode !== 'manual_ok') {
        skipped.push({
          parentSku: g.parentSku,
          reason: 'SKIPPED_BRAND_GATE',
          detail: 'ブランド／出品制限を人間確認後、' + APPROVAL_AMAZON_LV4_BRAND_GATE_MODE_PROP +
            '=manual_ok を設定してください（現在=' + (gateMode || '空') + '）'
        });
        continue;
      }
    }

    // 子をマスタ照合（価格フォールバックより先に解決）
    var children = [];
    var orphanChild = false;
    for (var c = 0; c < g.childLines.length; c++) {
      var cl = g.childLines[c];
      var childSku2 = String(cl.childSku || '').trim();
      var mr = amazonApprovalLv4FindMasterRow_(masterCtx, g.parentSku, childSku2);
      if (!mr) {
        skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_ORPHAN', detail: '子無し:' + childSku2 });
        orphanChild = true;
        break;
      }
      children.push({ line: cl, rowIndex0: mr.rowIndex0, childSku: childSku2 });
    }
    if (orphanChild) continue;

    if (resolvedTrack === 'B' && !children.length) {
      skipped.push({ parentSku: g.parentSku, reason: 'SKIPPED_INCOMPLETE_VARIATION', detail: '子解決0' });
      continue;
    }

    var priceRes = amazonApprovalLv4ResolvePriceAmazon_(
      masterCtx, g.parentSku, masterParent.rowIndex0, children
    );
    if (!priceRes.value) {
      skipped.push({
        parentSku: g.parentSku,
        reason: 'SKIPPED_NEED_HUMAN',
        detail: '販売価格amazon不正 tried=' + priceRes.sourceTried
      });
      continue;
    }

    parents.push({
      parentSku: g.parentSku,
      parentLine: g.parentLine,
      parentRowIndex0: masterParent.rowIndex0,
      children: children,
      resolvedTrack: resolvedTrack,
      priceAmazon: priceRes.value,
      priceSource: priceRes.source
    });
  }

  return { parents: parents, skipped: skipped };
}

function amazonApprovalLv4AnyInStock_(masterCtx, group) {
  var iStock = masterCtx.col['在庫数'];
  if (iStock == null) return '';
  function stockOf(rowIndex0) {
    var v = masterCtx.values[rowIndex0][iStock];
    var n = (v === '' || v == null) ? 0 : Number(v);
    return (!isNaN(n) && n > 0) ? n : 0;
  }
  var parentRow = amazonApprovalLv4FindMasterRow_(masterCtx, group.parentSku, '');
  if (parentRow && stockOf(parentRow.rowIndex0) > 0) {
    return '親在庫>0';
  }
  for (var i = 0; i < group.childLines.length; i++) {
    var cs = String(group.childLines[i].childSku || '').trim();
    var cr = amazonApprovalLv4FindMasterRow_(masterCtx, group.parentSku, cs);
    if (cr && stockOf(cr.rowIndex0) > 0) return '子在庫>0:' + cs;
  }
  return '';
}

function amazonApprovalLv4FindMasterRow_(masterCtx, parentSku, childSku) {
  var iParent = masterCtx.col['親SKU'];
  var iChild = masterCtx.col['子SKU'];
  var wantChild = String(childSku || '').trim();
  for (var r = masterCtx.headerRowIdx + 1; r < masterCtx.values.length; r++) {
    var row = masterCtx.values[r] || [];
    if (String(row[iParent] != null ? row[iParent] : '').trim() !== parentSku) continue;
    var c = iChild != null ? String(row[iChild] != null ? row[iChild] : '').trim() : '';
    if (c === wantChild) return { rowIndex0: r };
  }
  return null;
}

/**
 * M2(A)用 ASIN 解決。優先: ASINコード → 競合店ASINコード → 競合URL系から /dp/ASIN
 * @return {string}
 */
function amazonApprovalLv4ResolveAsinForOffer_(masterCtx, rowIndex0) {
  var direct = amazonApprovalLv4Cell_(masterCtx, rowIndex0, 'ASINコード');
  if (amazonApprovalLv4LooksLikeAsin_(direct)) return direct.toUpperCase();
  var rival = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '競合店ASINコード');
  if (amazonApprovalLv4LooksLikeAsin_(rival)) return rival.toUpperCase();
  var urlCols = ['競合AmazonページURL', '競合URL', 'Amazon URL', '商品URL'];
  for (var i = 0; i < urlCols.length; i++) {
    var u = amazonApprovalLv4Cell_(masterCtx, rowIndex0, urlCols[i]);
    var fromUrl = amazonApprovalLv4AsinFromUrl_(u);
    if (fromUrl) return fromUrl;
  }
  return '';
}

function amazonApprovalLv4LooksLikeAsin_(s) {
  return /^B0[A-Z0-9]{8}$/i.test(String(s || '').trim());
}

function amazonApprovalLv4AsinFromUrl_(url) {
  var m = String(url || '').match(/\/(?:dp|gp\/product|product)\/(B0[A-Z0-9]{8})/i);
  return m ? String(m[1]).toUpperCase() : '';
}

function amazonApprovalLv4Cell_(masterCtx, rowIndex0, colName) {
  var idx = masterCtx.col[colName];
  if (idx == null) return '';
  return String(masterCtx.values[rowIndex0][idx] != null ? masterCtx.values[rowIndex0][idx] : '').trim();
}

/**
 * 販売価格amazon 解決。親が空／不正なら承認済み子を順に見る。
 * Logger: [Lv4Price]
 * @param {Array<{rowIndex0:number, childSku:string}>} children
 * @return {{value:string, source:string, sourceTried:string}}
 */
function amazonApprovalLv4ResolvePriceAmazon_(masterCtx, parentSku, parentRowIndex0, children) {
  var tried = [];
  var parentRaw = amazonApprovalLv4Cell_(masterCtx, parentRowIndex0, '販売価格amazon');
  var parentNum = Number(parentRaw);
  tried.push('parent(v=' + parentRaw + ')');
  if (parentRaw !== '' && !isNaN(parentNum) && parentNum > 0) {
    Logger.log('[Lv4Price] HIT parentSku=' + parentSku + ' source=parent value=' + parentRaw);
    return { value: String(parentRaw), source: 'parent', sourceTried: tried.join(';') };
  }
  children = children || [];
  for (var i = 0; i < children.length; i++) {
    var ch = children[i];
    var raw = amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, '販売価格amazon');
    var n = Number(raw);
    tried.push('child:' + String(ch.childSku || '') + '(v=' + raw + ')');
    if (raw !== '' && !isNaN(n) && n > 0) {
      Logger.log('[Lv4Price] HIT parentSku=' + parentSku +
        ' source=child:' + String(ch.childSku || '') + ' value=' + raw);
      return { value: String(raw), source: 'child:' + String(ch.childSku || ''), sourceTried: tried.join(';') };
    }
  }
  Logger.log('[Lv4Price] EMPTY parentSku=' + parentSku + ' tried=' + tried.join(';'));
  return { value: '', source: '', sourceTried: tried.join(';') };
}

/**
 * Amazonカテゴリ解決。優先順:
 * 1) amazon カテゴリー  2) amazonカテゴリー  3) カテゴリー（T列。見出しに改行※注記があっても可）
 * Logger に source / 列index / 値プレビュー / 失敗時の候補ヘッダを残す。
 * @return {{value:string, source:string, sourceTried:string}}
 */
function amazonApprovalLv4ResolveCategory_(masterCtx, rowIndex0, parentSku) {
  var preferred = ['amazon カテゴリー', 'amazonカテゴリー', 'カテゴリー'];
  var tried = [];
  var i;
  var name;
  var colIdx;
  var raw;
  var val;

  for (i = 0; i < preferred.length; i++) {
    name = preferred[i];
    colIdx = masterCtx.col[name];
    raw = colIdx != null
      ? String(masterCtx.values[rowIndex0][colIdx] != null ? masterCtx.values[rowIndex0][colIdx] : '')
      : '';
    val = String(raw || '').trim();
    tried.push(name + '(exists=' + (colIdx != null) + ',idx=' + (colIdx != null ? colIdx : -1) +
      ',len=' + raw.length + ',v=' + (val ? val.substring(0, 24) : '') + ')');
    if (val) {
      Logger.log('[Lv4Cat] HIT parentSku=' + parentSku + ' row1=' + (rowIndex0 + 1) +
        ' source=' + name + ' colIndex0=' + colIdx + ' valuePreview=' + val.substring(0, 40));
      return { value: val, source: name, sourceTried: tried.join(';') };
    }
  }

  // 見出しが「カテゴリー\n※…」や「カテゴリー ※…」でも拾う（amazon系を優先）
  var fuzzyAmazon = null;
  var fuzzyPlain = null;
  for (var k in masterCtx.col) {
    if (!Object.prototype.hasOwnProperty.call(masterCtx.col, k)) continue;
    var first = String(k).split(/\r?\n/)[0].trim();
    var compact = first.replace(/\s+/g, ' ');
    if (compact === 'amazon カテゴリー' || compact === 'amazonカテゴリー' ||
        compact.indexOf('amazon') === 0 && compact.indexOf('カテゴリ') >= 0) {
      if (!fuzzyAmazon) fuzzyAmazon = k;
    } else if (compact === 'カテゴリー' || compact.indexOf('カテゴリー') === 0) {
      if (!fuzzyPlain) fuzzyPlain = k;
    }
  }
  var fuzzyOrder = [];
  if (fuzzyAmazon) fuzzyOrder.push(fuzzyAmazon);
  if (fuzzyPlain) fuzzyOrder.push(fuzzyPlain);

  for (i = 0; i < fuzzyOrder.length; i++) {
    name = fuzzyOrder[i];
    colIdx = masterCtx.col[name];
    raw = String(masterCtx.values[rowIndex0][colIdx] != null ? masterCtx.values[rowIndex0][colIdx] : '');
    val = String(raw || '').trim();
    tried.push('fuzzy:' + name.replace(/\r?\n/g, '\\n') + '(idx=' + colIdx + ',len=' + raw.length +
      ',v=' + (val ? val.substring(0, 24) : '') + ')');
    if (val) {
      Logger.log('[Lv4Cat] HIT_FUZZY parentSku=' + parentSku + ' row1=' + (rowIndex0 + 1) +
        ' sourceHeader=' + JSON.stringify(name) + ' colIndex0=' + colIdx +
        ' valuePreview=' + val.substring(0, 40));
      return { value: val, source: 'fuzzy:' + String(name).split(/\r?\n/)[0], sourceTried: tried.join(';') };
    }
  }

  var catKeys = [];
  for (var k2 in masterCtx.col) {
    if (!Object.prototype.hasOwnProperty.call(masterCtx.col, k2)) continue;
    if (String(k2).indexOf('カテゴリ') >= 0) {
      catKeys.push(JSON.stringify(k2) + '@' + masterCtx.col[k2]);
    }
  }
  Logger.log('[Lv4Cat] EMPTY parentSku=' + parentSku + ' row1=' + (rowIndex0 + 1) +
    ' tried=' + tried.join(';') + ' headersWithカテゴリ=' + catKeys.join('|'));
  return { value: '', source: '', sourceTried: tried.join(';') };
}

function amazonApprovalLv4BuildSubBatches_(parents, perSub) {
  var out = [];
  for (var i = 0; i < parents.length; i += perSub) {
    out.push(parents.slice(i, i + perSub));
  }
  return out;
}

function amazonApprovalLv4ParentSkuList_(parents) {
  var a = [];
  for (var i = 0; i < parents.length; i++) a.push(parents[i].parentSku);
  return a.join(',');
}

/**
 * Resolve 済み親を行化。価格/画像不足はここでは起きない想定。
 * 万一不足時は skipped に入れ、okParents には含めない。
 * @return {{headers:Array<string>, rows:Array<Array>, childCount:number, okParents:Array, skipped:Array}}
 */
function amazonApprovalLv4BuildRows_(masterCtx, parents, trackProp, stockOut, shipping) {
  var headers = [
    'track', 'parentSku', 'childSku', 'sellerSku', 'manufacturerPart',
    'productName', 'brand', 'priceAmazon', 'inventory', 'gtin', 'asin',
    'mainImageUrl', 'subImageUrls', 'amazonCategory', 'setCount',
    'shippingTemplate', 'variationRole'
  ];
  var rows = [];
  var childCount = 0;
  var okParents = [];
  var skipped = [];

  for (var p = 0; p < parents.length; p++) {
    var parent = parents[p];
    var rTrack = parent.resolvedTrack || trackProp;
    var priceRes2 = parent.priceAmazon
      ? { value: parent.priceAmazon, source: parent.priceSource || 'cached', sourceTried: '' }
      : amazonApprovalLv4ResolvePriceAmazon_(
          masterCtx, parent.parentSku, parent.parentRowIndex0, parent.children || []
        );
    var price = priceRes2.value;
    if (!price) {
      skipped.push({
        parentSku: parent.parentSku,
        reason: 'SKIPPED_NEED_HUMAN',
        detail: 'Build時:販売価格amazon不正 tried=' + priceRes2.sourceTried
      });
      continue;
    }
    var catResolved2 = amazonApprovalLv4ResolveCategory_(masterCtx, parent.parentRowIndex0, parent.parentSku);
    var cat = catResolved2.value;
    var catalogName = amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, 'オリジナルカタログ商品名');
    var asin = amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, 'ASINコード');
    var mainImg = amazonApprovalLv4ResolveMainImageUrl_(masterCtx, parent.parentRowIndex0);
    // Resolve と同様: 親空なら承認済み子の MAIN URL を親行用にフォールバック
    if (!mainImg && parent.children && parent.children.length) {
      var ciFb;
      for (ciFb = 0; ciFb < parent.children.length; ciFb++) {
        var chFb = parent.children[ciFb];
        mainImg = amazonApprovalLv4ResolveMainImageUrl_(masterCtx, chFb.rowIndex0);
        if (mainImg) {
          Logger.log('[Lv4Build] parentMainFallback parentSku=' + parent.parentSku +
            ' fromChild=' + String(chFb.childSku || ''));
          break;
        }
      }
    }

    if (rTrack === 'B') {
      if (!cat) {
        skipped.push({
          parentSku: parent.parentSku,
          reason: 'SKIPPED_NEED_HUMAN',
          detail: 'Build時:カテゴリ空 tried=' + catResolved2.sourceTried
        });
        continue;
      }
      if (!mainImg) {
        skipped.push({
          parentSku: parent.parentSku,
          reason: 'SKIPPED_NEED_HUMAN',
          detail: 'Build時:メイン画像空（親・子とも Amazon MAIN URL／楽天メイン画像1 なし）'
        });
        continue;
      }
      rows.push([
        'B', parent.parentSku, '', parent.parentSku, '',
        catalogName || amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, '商品名'),
        APPROVAL_AMAZON_LV4_BRAND, price, stockOut, '', '',
        mainImg, amazonApprovalLv4SubImages_(masterCtx, parent.parentRowIndex0),
        cat, amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, 'A.セット商品数'),
        shipping, 'parent'
      ]);
      for (var c = 0; c < parent.children.length; c++) {
        var ch = parent.children[c];
        var childMain = amazonApprovalLv4ResolveMainImageUrl_(masterCtx, ch.rowIndex0) || mainImg;
        var mfr = amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, 'メーカー品番');
        if (!mfr) mfr = amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, '型番');
        if (!mfr) mfr = ch.childSku;
        var childPrice = amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, '販売価格amazon') || price;
        var childName = catalogName ||
          amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, '商品名') ||
          amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, '商品名');
        rows.push([
          'B', parent.parentSku, ch.childSku, ch.childSku, mfr,
          childName, APPROVAL_AMAZON_LV4_BRAND, childPrice, stockOut, '', '',
          childMain, amazonApprovalLv4SubImages_(masterCtx, ch.rowIndex0),
          cat, amazonApprovalLv4Cell_(masterCtx, ch.rowIndex0, 'A.セット商品数'),
          shipping, 'child'
        ]);
        childCount++;
      }
      okParents.push(parent);
    } else {
      var targets = parent.children.length
        ? parent.children
        : [{ childSku: parent.parentSku, rowIndex0: parent.parentRowIndex0, line: parent.parentLine }];
      for (var a = 0; a < targets.length; a++) {
        var t = targets[a];
        var sku = t.childSku || parent.parentSku;
        var mfrA = amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, 'メーカー品番');
        if (!mfrA) mfrA = amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, '型番');
        if (!mfrA) mfrA = sku;
        var jan = amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, '商品コード(JANコード等)');
        if (!jan) jan = amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, 'JANコード');
        var asinOffer = amazonApprovalLv4ResolveAsinForOffer_(masterCtx, t.rowIndex0) ||
          amazonApprovalLv4ResolveAsinForOffer_(masterCtx, parent.parentRowIndex0) ||
          asin;
        if (!asinOffer) {
          skipped.push({
            parentSku: parent.parentSku,
            reason: 'SKIPPED_NEED_HUMAN',
            detail: 'Build時:Aトラック ASIN無し sku=' + sku
          });
          continue;
        }
        rows.push([
          'A', parent.parentSku, t.childSku || '', sku, mfrA,
          amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, '商品名'),
          '', amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, '販売価格amazon') || price,
          stockOut, jan, asinOffer,
          amazonApprovalLv4ResolveMainImageUrl_(masterCtx, t.rowIndex0),
          amazonApprovalLv4SubImages_(masterCtx, t.rowIndex0),
          cat, amazonApprovalLv4Cell_(masterCtx, t.rowIndex0, 'A.セット商品数'),
          shipping, 'offer'
        ]);
        childCount++;
      }
      okParents.push(parent);
    }
  }

  return {
    headers: headers,
    rows: rows,
    childCount: childCount,
    okParents: okParents,
    skipped: skipped
  };
}

function amazonApprovalLv4SubImages_(masterCtx, rowIndex0) {
  var amazonPt = amazonApprovalLv4Cell_(masterCtx, rowIndex0, 'Amazon PT URL');
  if (amazonPt) return amazonPt;
  var urls = [];
  for (var i = 1; i <= 8; i++) {
    var u = amazonApprovalLv4Cell_(masterCtx, rowIndex0, '楽天サブ画像' + i);
    if (u) urls.push(u);
  }
  return urls.join('|');
}

/** U4: Amazon MAIN URL があれば優先、なければ楽天メイン画像1 */
function amazonApprovalLv4ResolveMainImageUrl_(masterCtx, rowIndex0) {
  var amz = amazonApprovalLv4Cell_(masterCtx, rowIndex0, 'Amazon MAIN URL');
  if (amz) return amz;
  return amazonApprovalLv4Cell_(masterCtx, rowIndex0, '楽天メイン画像1') || '';
}

function amazonApprovalLv4RowsToCsv_(headers, rows) {
  function esc(v) {
    var s = String(v == null ? '' : v);
    if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }
  var lines = [headers.map(esc).join(',')];
  for (var i = 0; i < rows.length; i++) {
    lines.push(rows[i].map(esc).join(','));
  }
  return lines.join('\r\n');
}

function amazonApprovalLv4EnsureLogSheet_(ss) {
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(APPROVAL_AMAZON_LV4_LOG_SHEET);
    sh.getRange(1, 1, 1, APPROVAL_AMAZON_LV4_LOG_HEADERS.length).setValues([APPROVAL_AMAZON_LV4_LOG_HEADERS]);
    sh.appendRow([
      'EXEMPTION', '', '', '', '', '', '',
      '', '', '', '', '', '',
      '（この行を編集: カテゴリ実値または * ・承認日・証跡URLをすべて記入。brand=ノーブランド品）',
      '（カテゴリ or *）', APPROVAL_AMAZON_LV4_BRAND, '', ''
    ]);
    return sh;
  }
  var h = String(sh.getRange(1, 1).getValue() || '');
  if (h !== 'recordType') {
    // 証跡・履歴消失防止のため clear しない
    throw new Error(
      'シート「' + APPROVAL_AMAZON_LV4_LOG_SHEET + '」の1行目が recordType ではありません。' +
      '手動でヘッダーを直すか、別名退避してから再実行してください（自動clearはしません）。'
    );
  }
  return sh;
}

function amazonApprovalLv4LoadExemptions_(ss) {
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[0]) !== 'EXEMPTION') continue;
    out.push({
      category: String(row[14] || '').trim(),
      brand: String(row[15] || '').trim(),
      exemptionDate: String(row[16] || '').trim(),
      evidenceUrl: String(row[17] || '').trim()
    });
  }
  return out;
}

/**
 * GTIN免除: カテゴリ空は不合格。ブランドは「ノーブランド品」完全一致。
 * 承認日・証跡URL必須。カテゴリは完全一致または * のみ。
 */
function amazonApprovalLv4HasGtinExemption_(exemptions, category) {
  var cat = String(category || '').trim();
  if (!cat) return false;
  for (var i = 0; i < exemptions.length; i++) {
    var e = exemptions[i];
    if (String(e.brand || '').trim() !== APPROVAL_AMAZON_LV4_BRAND) continue;
    if (!e.exemptionDate || !e.evidenceUrl) continue;
    if (e.category === '（カテゴリ or *）') continue;
    if (e.category === '*') return true;
    if (e.category && e.category === cat) return true;
  }
  return false;
}

/**
 * 同一 batchId+track で、意味ある状態の最新が GENERATED/UPLOADED_OK/PACKAGED の親 → ブロック。
 * SKIPPED_* / DRY_RUN / FAILED は latest 判定に使わない（汚染防止）。
 * UPLOAD_FAILED が最新ならブロックしない。
 * @return {Object<string,string>} parentSku -> status
 */
function amazonApprovalLv4LoadBlockedParents_(ss, batchId, track) {
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  var latest = {};
  if (!sh || sh.getLastRow() < 2) return {};
  var data = sh.getDataRange().getValues();
  var wantBatch = String(batchId || '');
  var wantTrack = String(track || '');
  var meaningful = {
    GENERATED: true,
    UPLOADED_OK: true,
    PACKAGED: true,
    UPLOAD_FAILED: true
  };
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[0]) !== 'RUN') continue;
    if (wantBatch && String(row[2]) !== wantBatch) continue;
    var rowTrack = String(row[4] || '');
    if (wantTrack && rowTrack && rowTrack !== wantTrack && wantTrack !== 'BOTH') continue;
    var status = String(row[7] || '');
    if (!meaningful[status]) continue;
    var parents = String(row[5] || '').split(',');
    for (var p = 0; p < parents.length; p++) {
      var ps = String(parents[p] || '').trim();
      if (!ps) continue;
      latest[ps] = status;
    }
  }
  var blocked = {};
  var keys = Object.keys(latest);
  for (var k = 0; k < keys.length; k++) {
    var st = latest[keys[k]];
    if (st === 'GENERATED' || st === 'UPLOADED_OK' || st === 'PACKAGED') {
      blocked[keys[k]] = st;
    }
  }
  return blocked;
}

/**
 * スキップ行は recordType=SKIP（冪等 latest を汚染しない）。
 */
function amazonApprovalLv4LogSkippedParents_(ss, runId, batchId, track, skipped) {
  if (!skipped || !skipped.length) return;
  for (var i = 0; i < skipped.length; i++) {
    var s = skipped[i];
    amazonApprovalLv4AppendLog_(ss, {
      recordType: 'SKIP',
      runId: runId,
      batchId: batchId,
      subBatchId: '',
      track: track,
      parentSkus: s.parentSku || '',
      childCount: 0,
      status: s.reason || 'SKIPPED_NEED_HUMAN',
      fileName: '',
      fileUrl: '',
      note: s.detail || ''
    });
  }
}

/**
 * 同一 batchId + trackTag の subBatchId 連番の次番号（1始まり・単調増加）。
 */
function amazonApprovalLv4NextSubBatchSeq_(ss, batchId, trackTag, resumeState) {
  var maxSeq = 0;
  if (resumeState && resumeState.nextSubSeq) {
    maxSeq = Math.max(maxSeq, Number(resumeState.nextSubSeq) - 1);
  }
  var sh = ss.getSheetByName(APPROVAL_AMAZON_LV4_LOG_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues();
    var prefix = String(batchId) + '_' + String(trackTag);
    var re = new RegExp('^' + prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\d+)$');
    for (var r = 1; r < data.length; r++) {
      var id = String(data[r][3] || '');
      var m = id.match(re);
      if (m) {
        var n = Number(m[1]);
        if (!isNaN(n) && n > maxSeq) maxSeq = n;
      }
    }
  }
  return maxSeq + 1;
}

function amazonApprovalLv4ArchiveSameName_(folder, fileName) {
  var existing = folder.getFilesByName(fileName);
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  while (existing.hasNext()) {
    var old = existing.next();
    try {
      old.setName(fileName.replace(/(\.[^.]+)$/, '_archived_' + stamp + '$1'));
    } catch (e) {
      throw new Error(
        '同名ファイルのアーカイブ改名に失敗（trashしない）: ' + fileName + ' / ' +
        ((e && e.message) || e)
      );
    }
  }
}

function amazonApprovalLv4AppendLog_(ss, obj) {
  var sh = amazonApprovalLv4EnsureLogSheet_(ss);
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
  var status = String(obj.status || '');
  var generatedAt = (status === 'GENERATED') ? now : '';
  var uploadedAt = (status === 'UPLOADED_OK') ? now : '';
  sh.appendRow([
    obj.recordType || 'RUN',
    obj.runId || '',
    obj.batchId || '',
    obj.subBatchId || '',
    obj.track || '',
    obj.parentSkus || '',
    obj.childCount != null ? obj.childCount : '',
    status,
    obj.fileName || '',
    obj.fileUrl || '',
    generatedAt,
    '',
    uploadedAt,
    obj.note || '',
    obj.category || '',
    obj.brand || '',
    obj.exemptionDate || '',
    obj.evidenceUrl || ''
  ]);
}

/**
 * 追記専用。既存行は上書きしない。
 * UPLOADED_OK は同一 subBatchId の GENERATED が無いと拒否。
 */
function amazonApprovalLv4MarkStatus_(ss, subBatchId, status, note) {
  var sh = amazonApprovalLv4EnsureLogSheet_(ss);
  var data = sh.getDataRange().getValues();
  var who = '';
  try {
    who = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (e) {}
  var hasGenerated = false;
  var parentSkus = '';
  var batchId = '';
  var track = '';
  var runId = '';
  var fileName = '';
  var fileUrl = '';
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) !== 'RUN') continue;
    if (String(data[r][3]) !== subBatchId) continue;
    if (String(data[r][7]) === 'GENERATED') {
      hasGenerated = true;
      parentSkus = String(data[r][5] || '');
      batchId = String(data[r][2] || '');
      track = String(data[r][4] || '');
      runId = String(data[r][1] || '');
      fileName = String(data[r][8] || '');
      fileUrl = String(data[r][9] || '');
    }
  }
  if (status === 'UPLOADED_OK' && !hasGenerated) {
    throw new Error('subBatchId=' + subBatchId + ' の GENERATED 行がありません。任意文字列では完了記録できません。');
  }
  if (status === 'UPLOAD_FAILED' && !hasGenerated) {
    Logger.log('[amazonApprovalLv4MarkStatus_] WARN no GENERATED for ' + subBatchId + ' — append anyway');
  }
  amazonApprovalLv4AppendLog_(ss, {
    recordType: 'RUN',
    runId: runId,
    batchId: batchId,
    subBatchId: subBatchId,
    track: track,
    parentSkus: parentSkus,
    childCount: '',
    status: status,
    fileName: fileName,
    fileUrl: fileUrl,
    note: (note || '') + ' by=' + who + ' (append-only)'
  });
  return 1;
}

function amazonApprovalLv4GetOrCreateFolder_(ss) {
  var folderId = PropertiesService.getScriptProperties().getProperty(APPROVAL_AMAZON_LV4_FOLDER_ID_PROP);
  if (folderId) {
    try {
      return DriveApp.getFolderById(String(folderId).trim());
    } catch (e) {
      Logger.log('[amazonApprovalLv4GetOrCreateFolder_] invalid folder id, fallback create');
    }
  }
  var name = 'Lv4_Amazon_GENERATED';
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function amazonApprovalLv4SkipSummary_(skipped) {
  if (!skipped || !skipped.length) return '0';
  var counts = {};
  for (var i = 0; i < skipped.length; i++) {
    var r = skipped[i].reason || 'OTHER';
    counts[r] = (counts[r] || 0) + 1;
  }
  var parts = [];
  var keys = Object.keys(counts);
  for (var k = 0; k < keys.length; k++) {
    parts.push(keys[k] + '=' + counts[keys[k]]);
  }
  return parts.join(',');
}

function amazonApprovalLv4SaveState_(state) {
  PropertiesService.getScriptProperties().setProperty(APPROVAL_AMAZON_LV4_STATE_PROP, JSON.stringify(state));
}

function amazonApprovalLv4SetTrigger_() {
  amazonApprovalLv4DeleteTriggers_();
  ScriptApp.newTrigger(APPROVAL_AMAZON_LV4_TRIGGER_FN).timeBased().after(60 * 1000).create();
}

function amazonApprovalLv4DeleteTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === APPROVAL_AMAZON_LV4_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function amazonApprovalLv4Mail_(subject, body) {
  try {
    var email = Session.getEffectiveUser().getEmail();
    if (!email) return;
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('[amazonApprovalLv4Mail_] ' + ((e && e.message) || e));
  }
}
