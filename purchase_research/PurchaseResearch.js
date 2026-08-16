/**
 * 仕入れ検討①。出品 コード.js には置かない。
 * /query は GET + selection= のみ。POST {"selection":} は使うな。
 */
var PR_LOG_SHEET_ = '①ログ';
var PR_CAND_SHEET_ = '①候補';
var PR_PROF_SHEET_ = '①プロファイル';
var PR_FOOD_ROOT_ = '57239051';
var PR_SELLER_MORITA_ = 'AYC4Z8PML8T30';
var PR_KEEPA_QUERY_ = 'https://api.keepa.com/query';
var PR_KEEPA_PRODUCT_ = 'https://api.keepa.com/product';
var PR_COMPETITOR_SS_ = '1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs';
var PR_KEEPA_FULL_SHEET_ = 'Keepaフル';
var PR_FRESH_MS_ = 90 * 24 * 60 * 60 * 1000;
var PR_FROZEN_RE_ = /冷凍|冷蔵|生鮮/;
var PR_CAND_HEADERS_ = [
  'メーカー',
  'ASIN',
  '商品名',
  '税込価格',
  '順位90',
  'JAN',
  '発見経路',
  'sellerId',
  '門結果',
  '門理由',
  'runId',
  '取得日時',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('仕入れ検討①')
    .addItem('モリタ /query ドライラン', 'prMenuQueryMorita')
    .addItem('モリタ 取扱→門', 'prMenuMoritaGate')
    .addToUi();
}

function prMenuQueryMorita() {
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=query state=RUNNING seller=%s', runId, PR_SELLER_MORITA_);
  try {
    var result = prKeepaQueryGet_(PR_SELLER_MORITA_, PR_FOOD_ROOT_, 0, 100);
    prAppendLog_([
      runId,
      'モリタ /query GET',
      'DONE',
      new Date(),
      result.totalResults,
      result.asinCount,
      result.tokensConsumed,
      (result.asinList[0] || ''),
    ]);
    Logger.log(
      '[仕入れ検討①] runId=%s state=DONE total=%s n=%s consumed=%s first=%s',
      runId,
      result.totalResults,
      result.asinCount,
      result.tokensConsumed,
      result.asinList[0] || ''
    );
    SpreadsheetApp.getUi().alert(
      '① /query GET\n合計 ' + result.totalResults + ' / asinList ' + result.asinCount +
        '\nconsumed ' + result.tokensConsumed + '\n先頭 ' + (result.asinList[0] || '')
    );
  } catch (e) {
    var msg = String((e && e.message) || e);
    msg = prRedactKey_(msg);
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'モリタ /query GET', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    SpreadsheetApp.getUi().alert('失敗: ' + msg.slice(0, 300));
    throw e;
  }
}

function prKeepaKey_() {
  var key = String(PropertiesService.getScriptProperties().getProperty('KEEPA_API_KEY') || '').trim();
  if (!key) {
    throw new Error('Script Property KEEPA_API_KEY が未設定（出品と同じキーを①プロジェクトへ貼る）');
  }
  return key;
}

/** GET /query?domain=5&selection= JSON。POST ボディは使わない。 */
function prKeepaQueryGet_(sellerId, rootCategory, page, perPage) {
  var selection = {
    rootCategory: [String(rootCategory)],
    sellerIds: [String(sellerId)],
    productType: ['0'],
    sort: [['current_SALES', 'asc'], ['monthlySold', 'desc']],
    page: Number(page) || 0,
    perPage: Number(perPage) || 100,
  };
  var selStr = JSON.stringify(selection);
  var url =
    PR_KEEPA_QUERY_ +
    '?key=' +
    encodeURIComponent(prKeepaKey_()) +
    '&domain=5&selection=' +
    encodeURIComponent(selStr);
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { Accept: 'application/json' },
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '';
  if (code !== 200) {
    throw new Error('Keepa /query HTTP ' + code + ' ' + prRedactKey_(body).slice(0, 180));
  }
  var data = JSON.parse(body);
  var asins = data.asinList || [];
  if (!Array.isArray(asins)) asins = [];
  Logger.log(
    '[仕入れ検討①] query GET ok total=%s n=%s consumed=%s',
    data.totalResults,
    asins.length,
    data.tokensConsumed
  );
  return {
    totalResults: data.totalResults,
    asinCount: asins.length,
    asinList: asins.map(function (a) {
      return String(a);
    }),
    tokensConsumed: data.tokensConsumed,
  };
}

function prRedactKey_(s) {
  return String(s || '').replace(/key=[^&"'\s]+/gi, 'key=REDACTED');
}

function prSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function prEnsureSheet_(title, headers) {
  var ss = prSpreadsheet_();
  var sh = ss.getSheetByName(title);
  if (!sh) {
    sh = ss.insertSheet(title);
  }
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function prAppendLog_(row) {
  var headers = ['runId', '内容', 'state', 'at', 'totalResults', 'asinCount', 'tokensConsumed', 'detail'];
  var sh = prEnsureSheet_(PR_LOG_SHEET_, headers);
  sh.appendRow(row);
}

function prMenuMoritaGate() {
  var runId = 'pr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  Logger.log('[仕入れ検討①] runId=%s step=gate state=RUNNING seller=%s', runId, PR_SELLER_MORITA_);
  try {
    var q = prKeepaQueryGet_(PR_SELLER_MORITA_, PR_FOOD_ROOT_, 0, 100);
    var prof = prReadProfile_();
    var products = prKeepaProductStats90_(q.asinList);
    var pass = 0;
    var drop = 0;
    var i;
    for (i = 0; i < products.length; i++) {
      var row = prGateProductRow_(products[i], prof, PR_SELLER_MORITA_, runId);
      prUpsertCandidate_(row);
      if (row[8] === '通過') pass++;
      else drop++;
    }
    var detail =
      'pass=' + pass + ' drop=' + drop + ' priceMin=' + prof.priceMin + ' rankMax=' + prof.rankMax;
    prAppendLog_([
      runId,
      'モリタ 取扱→門',
      'DONE',
      new Date(),
      q.totalResults,
      products.length,
      '',
      detail,
    ]);
    Logger.log('[仕入れ検討①] runId=%s state=DONE %s', runId, detail);
    SpreadsheetApp.getUi().alert('① 門\n' + detail + '\nquery n=' + q.asinCount);
  } catch (e) {
    var msg = prRedactKey_(String((e && e.message) || e));
    Logger.log('[仕入れ検討①] runId=%s state=FAILED %s', runId, msg);
    prAppendLog_([runId, 'モリタ 取扱→門', 'FAILED', new Date(), '', '', '', msg.slice(0, 200)]);
    SpreadsheetApp.getUi().alert('失敗: ' + msg.slice(0, 300));
    throw e;
  }
}

/** history=0 では current/csv が空。stats=90（offers なし）。90日内の Keepaフルを先に使う。 */
function prKeepaProductStats90_(asins) {
  var packed = prSplitKeepaFull_(asins || []);
  Logger.log('[仕入れ検討①] Keepaフル hit=%s need_get=%s', packed.have.length, packed.need.length);
  var fetched = packed.need.length ? prKeepaProductStats90Fetch_(packed.need) : [];
  return packed.have.concat(fetched);
}

function prParseKeepaFullTime_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) return raw.getTime();
  var s = String(raw || '').trim();
  if (!s) return 0;
  var t = Date.parse(s);
  return isNaN(t) ? 0 : t;
}

function prLoadKeepaFullMap_() {
  var out = {};
  try {
    var sh = SpreadsheetApp.openById(PR_COMPETITOR_SS_).getSheetByName(PR_KEEPA_FULL_SHEET_);
    if (!sh || sh.getLastRow() < 2) return out;
    var lastCol = sh.getLastColumn();
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var idxAsin = -1;
    var idxAt = -1;
    var idxJson = -1;
    var h;
    for (h = 0; h < hdr.length; h++) {
      var name = String(hdr[h] || '').trim();
      if (name === 'ASIN') idxAsin = h;
      if (name === '取得日時') idxAt = h;
      if (name === '生JSON') idxJson = h;
    }
    if (idxAsin < 0 || idxJson < 0) return out;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      var asin = String(vals[r][idxAsin] || '').trim().toUpperCase();
      if (!asin) continue;
      var product = null;
      try {
        product = JSON.parse(String(vals[r][idxJson] || ''));
      } catch (e1) {
        product = null;
      }
      out[asin] = {
        at: idxAt >= 0 ? prParseKeepaFullTime_(vals[r][idxAt]) : 0,
        product: product,
      };
    }
  } catch (e) {
    Logger.log('[仕入れ検討①] Keepaフル読取スキップ %s', String((e && e.message) || e).slice(0, 160));
  }
  return out;
}

function prSplitKeepaFull_(asins) {
  var have = [];
  var need = [];
  var map = prLoadKeepaFullMap_();
  var now = new Date().getTime();
  var i;
  for (i = 0; i < asins.length; i++) {
    var a = String(asins[i] || '').trim().toUpperCase();
    if (!a) continue;
    var rec = map[a];
    if (rec && rec.product && rec.at && now - rec.at <= PR_FRESH_MS_) {
      have.push(rec.product);
    } else {
      need.push(a);
    }
  }
  return { have: have, need: need };
}

function prKeepaProductStats90Fetch_(asins) {
  var out = [];
  var i;
  for (i = 0; i < asins.length; i += 50) {
    var chunk = asins.slice(i, i + 50);
    var url =
      PR_KEEPA_PRODUCT_ +
      '?key=' +
      encodeURIComponent(prKeepaKey_()) +
      '&domain=5&asin=' +
      encodeURIComponent(chunk.join(',')) +
      '&history=0&stats=90';
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { Accept: 'application/json' },
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText() || '';
    if (code !== 200) {
      throw new Error('Keepa /product HTTP ' + code + ' ' + prRedactKey_(body).slice(0, 180));
    }
    var data = JSON.parse(body);
    Logger.log(
      '[仕入れ検討①] product stats=90 n=%s consumed=%s left=%s',
      (data.products || []).length,
      data.tokensConsumed,
      data.tokensLeft
    );
    var products = data.products || [];
    var j;
    for (j = 0; j < products.length; j++) out.push(products[j]);
  }
  return out;
}

function prReadProfile_() {
  var def = { priceMin: 2000, rankMax: 150000 };
  var sh = prSpreadsheet_().getSheetByName(PR_PROF_SHEET_);
  if (!sh || sh.getLastRow() < 2) return def;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {};
  var i;
  for (i = 0; i < headers.length; i++) map[String(headers[i] || '')] = row[i];
  var price = Number(map['価格下限']);
  var rank = Number(map['順位段階1']);
  if (price > 0) def.priceMin = price;
  if (rank > 0) def.rankMax = rank;
  return def;
}

function prStatsSlot_(p, idx) {
  var st = p && p.stats ? p.stats : {};
  var keys = ['avg90', 'avg30', 'avg'];
  var k;
  for (k = 0; k < keys.length; k++) {
    var arr = st[keys[k]];
    if (!arr || !arr.length || idx >= arr.length) continue;
    var v = arr[idx];
    if (v === null || v === undefined || v === -1) continue;
    var n = Number(v);
    if (!isNaN(n) && n >= 0) return n;
  }
  return null;
}

function prGateProductRow_(p, prof, sellerId, runId) {
  var title = String(p.title || '');
  var price = prStatsSlot_(p, 18);
  if (price === null) price = prStatsSlot_(p, 1);
  var rank = prStatsSlot_(p, 3);
  var st = '通過';
  var why = '門通過';
  if (price === null && rank === null) {
    st = '落ち';
    why = 'stats空';
  } else if (price !== null && price < prof.priceMin) {
    st = '落ち';
    why = '価格<' + prof.priceMin;
  } else if (rank !== null && rank > prof.rankMax) {
    st = '落ち';
    why = '順位>' + prof.rankMax;
  } else if (PR_FROZEN_RE_.test(title)) {
    st = '落ち';
    why = '冷凍冷蔵生鮮';
  }
  var eans = p.eanList || [];
  var jan = eans.length ? String(eans[0]) : '';
  var mfr = String(p.manufacturer || p.brand || '');
  return [
    mfr,
    String(p.asin || '').toUpperCase(),
    title,
    price === null ? '' : price,
    rank === null ? '' : rank,
    jan,
    'keepa_query',
    sellerId,
    st,
    why,
    runId,
    new Date(),
  ];
}

function prUpsertCandidate_(row) {
  var sh = prEnsureSheet_(PR_CAND_SHEET_, PR_CAND_HEADERS_);
  var need = PR_CAND_HEADERS_;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var i;
  for (i = 0; i < need.length; i++) {
    var found = false;
    var c;
    for (c = 0; c < headers.length; c++) {
      if (String(headers[c]) === need[i]) found = true;
    }
    if (!found) {
      sh.getRange(1, headers.length + 1).setValue(need[i]);
      headers.push(need[i]);
    }
  }
  var idx = {};
  for (i = 0; i < headers.length; i++) idx[String(headers[i])] = i;
  var last = sh.getLastRow();
  var asin = String(row[1] || '').toUpperCase();
  if (!asin) return;
  var line = [];
  for (i = 0; i < headers.length; i++) line.push('');
  var rec = {};
  for (i = 0; i < need.length; i++) rec[need[i]] = row[i];
  for (i = 0; i < need.length; i++) {
    if (idx[need[i]] === undefined) continue;
    line[idx[need[i]]] = rec[need[i]];
  }
  if (last >= 2 && idx['ASIN'] !== undefined) {
    var vals = sh.getRange(2, idx['ASIN'] + 1, last - 1, 1).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      if (String(vals[r][0] || '').toUpperCase() === asin) {
        sh.getRange(r + 2, 1, 1, line.length).setValues([line]);
        return;
      }
    }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, line.length).setValues([line]);
}
