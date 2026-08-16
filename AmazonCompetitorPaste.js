/**
 * Amazon 競合貼付 P1: Catalog 階段で ASIN 空欄のみ埋める。
 * 正本: docs/org/B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md
 * 領域1 clasp 禁止。HTML禁止。brandNames 禁止。C′は本線に出さない。
 * 評価列・既存ASIN・◎行は上書きしない。マスタ子SKUは回さない。
 */

var AMAZON_PASTE_CATALOG_MAX_FILL_ = 15;

function amazonPasteNormMaker_(s) {
  return String(s || '').replace(/[\s\u3000]/g, '').toLowerCase();
}

function amazonPasteTitleHasMaker_(title, maker) {
  var m = amazonPasteNormMaker_(maker);
  if (!m) return true;
  return amazonPasteNormMaker_(title).indexOf(m) >= 0;
}

function amazonPasteShortCore_(name, maker) {
  var s = String(name || '');
  var mk = String(maker || '').trim();
  if (mk) s = s.split(mk).join(' ');
  var words = String(s).split(/[\s\u3000]+/).filter(function (w) { return w.length >= 2; });
  if (words.length) return words.slice(0, 2).join(' ');
  return s.replace(/\s+/g, '').substring(0, 12);
}

function amazonPasteParseCatalogItems_(json) {
  var items = (json && json.items) ? json.items : [];
  var out = [];
  var i;
  for (i = 0; i < items.length; i++) {
    var it = items[i] || {};
    var asin = String(it.asin || it.ASIN || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{10}$/.test(asin)) continue;
    var title = '';
    var sums = it.summaries || [];
    if (sums[0]) title = String(sums[0].itemName || sums[0].name || '').trim();
    out.push({ asin: asin, title: title });
  }
  return out;
}

function amazonPasteCatalogSearch_(creds, token, stage, jan, maker, name) {
  var q = { marketplaceIds: creds.marketplaceId, includedData: 'summaries,identifiers', pageSize: '20' };
  if (stage === 'A_id') {
    q.identifiers = String(jan || '').replace(/\D/g, '');
    q.identifiersType = 'JAN';
  } else if (stage === 'B_kw_jan') {
    q.keywords = String(jan || '').replace(/\D/g, '');
  } else if (stage === 'D_maker_name') {
    q.keywords = (String(maker || '').trim() + ' ' + String(name || '').trim()).substring(0, 80);
  } else if (stage === 'E_core') {
    q.keywords = amazonPasteShortCore_(name, maker).substring(0, 80);
  } else {
    return [];
  }
  if (q.identifiers && String(q.identifiers).length < 8) return [];
  if (q.keywords && !String(q.keywords).trim()) return [];
  var res = amazonSpapiPutHttpGet_(creds, token, '/catalog/2022-04-01/items', q);
  Logger.log('[AmazonPasteP1] stage=' + stage + ' http=' + res.code + ' jan=' + jan);
  if (res.code !== 200 || !res.json) return [];
  return amazonPasteParseCatalogItems_(res.json);
}

/**
 * @param {boolean} write
 * @return {{ok:boolean, msg:string, filled:number}}
 */
function amazonPasteCatalogFillEmptyAsinsImpl_(write) {
  var runId = 'pasteP1-' + Utilities.getUuid().slice(0, 8);
  Logger.log('[AmazonPasteP1] runId=' + runId + ' state=RUNNING write=' + !!write);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  if (!sheet) {
    return { ok: false, msg: 'ASIN貼り付けシートなし', filled: 0 };
  }
  var creds = amazonSpapiPutLoadCreds_();
  var acc = amazonSpapiPutAcquireAccess_();
  var token = acc && acc.accessToken;
  if (!token) return { ok: false, msg: 'SP-API tokenなし', filled: 0 };

  var blockCols = ASIN_PASTE_BLOCK_HEADERS.length;
  var nBlocks = ASIN_PASTE_DEFAULT_BLOCKS;
  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) values = [[], []];
  var filledTotal = 0;
  var notes = [];
  var janCache = {};
  var b;
  for (b = 0; b < nBlocks; b++) {
    var jan = getJanFromAiDataForAsinPasteBlock_(ss, b);
    var janDigits = String(jan || '').replace(/\D/g, '');
    if (janDigits.length < 8) continue;
    var startCol = b * blockCols;
    var colAsin = startCol + 1;
    var colName = startCol + 2;
    var colEval = startCol + 3;
    var existing = {};
    var emptyRows = [];
    var lastUsed = 1;
    var r;
    for (r = 2; r < values.length; r++) {
      var row = values[r] || [];
      var blockHas = false;
      var c;
      for (c = 0; c < blockCols; c++) {
        if (String(row[startCol + c] || '').trim()) blockHas = true;
      }
      if (blockHas) lastUsed = r;
      var a0 = String(row[colAsin] || '').trim().toUpperCase();
      if (/^[A-Z0-9]{10}$/.test(a0)) existing[a0] = true;
    }
    for (r = 2; r <= lastUsed; r++) {
      var row2 = values[r] || [];
      var a1 = String(row2[colAsin] || '').trim();
      var ev = String(row2[colEval] || '').trim();
      if (!a1 && ev !== '◎') emptyRows.push(r);
    }
    if (!emptyRows.length) {
      notes.push('b' + b + ' jan=' + jan + ' stage=skip_no_empty fill=0 lastUsed=' + (lastUsed + 1));
      Logger.log('[AmazonPasteP1] runId=' + runId + ' ' + notes[notes.length - 1]);
      continue;
    }
    var aiSheet = ss.getSheetByName(AI_SHEET_NAME_FOR_ASIN_PASTE);
    var aiRow = 2 + b;
    var maker = '';
    var name = '';
    if (aiSheet) {
      var map = getAiColumnIndexMap_(aiSheet.getRange(1, 1, 1, aiSheet.getLastColumn()).getValues()[0]);
      if (map['メーカー名'] !== undefined) maker = String(aiSheet.getRange(aiRow, map['メーカー名'] + 1).getValue() || '').trim();
      if (map['商品名'] !== undefined) name = String(aiSheet.getRange(aiRow, map['商品名'] + 1).getValue() || '').trim();
    }
    var hits;
    var stageUsed;
    if (janCache[janDigits]) {
      hits = janCache[janDigits].hits;
      stageUsed = janCache[janDigits].stage + '(reuse)';
      Logger.log('[AmazonPasteP1] runId=' + runId + ' b=' + b + ' jan=' + jan + ' reuse');
    } else {
      hits = [];
      stageUsed = '';
      var stages = ['A_id', 'B_kw_jan', 'D_maker_name', 'E_core'];
      var si;
      for (si = 0; si < stages.length; si++) {
        Logger.log('[AmazonPasteP1] runId=' + runId + ' b=' + b + ' stage=' + stages[si] + ' jan=' + jan);
        var raw = amazonPasteCatalogSearch_(creds, token, stages[si], jan, maker, name);
        var kept = [];
        var hi;
        for (hi = 0; hi < raw.length; hi++) {
          if (!amazonPasteTitleHasMaker_(raw[hi].title, maker)) continue;
          kept.push(raw[hi]);
        }
        if (kept.length) {
          hits = kept;
          stageUsed = stages[si];
          break;
        }
      }
      janCache[janDigits] = { hits: hits, stage: stageUsed || 'none' };
    }
    var filledB = 0;
    var slotI = 0;
    var hi2;
    for (hi2 = 0; hi2 < hits.length && slotI < emptyRows.length && filledB < AMAZON_PASTE_CATALOG_MAX_FILL_; hi2++) {
      var h = hits[hi2];
      if (existing[h.asin]) continue;
      var rr = emptyRows[slotI];
      slotI++;
      if (write) {
        sheet.getRange(rr + 1, colAsin + 1).setValue(h.asin);
        if (!String((values[rr] || [])[colName] || '').trim()) {
          sheet.getRange(rr + 1, colName + 1).setValue(h.title);
        }
      }
      existing[h.asin] = true;
      filledB++;
      filledTotal++;
    }
    notes.push('b' + b + ' jan=' + jan + ' stage=' + (stageUsed || 'none') + ' fill=' + filledB);
    Logger.log('[AmazonPasteP1] runId=' + runId + ' ' + notes[notes.length - 1]);
  }
  Logger.log('[AmazonPasteP1] runId=' + runId + ' state=DONE filled=' + filledTotal);
  return { ok: true, msg: notes.join('\n'), filled: filledTotal };
}

function menuAmazonPasteCatalogFillEmptyAsinsDryRun() {
  menuAPrepAsinExtractDryRun();
}

function menuAmazonPasteCatalogFillEmptyAsins() {
  menuAPrepAsinExtractApply();
}

function menuAPrepAsinExtractDryRun() {
  var out = amazonPasteCatalogFillEmptyAsinsImpl_(false);
  SpreadsheetApp.getUi().alert(
    'A.準備 dry_run（未書込）\n空欄だけ計画。◎・既存ASINは触らない。\nfilled概算=' + out.filled + '\n\n' + out.msg
  );
}

function menuAPrepAsinExtractApply() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.alert(
    'A.準備（ASIN抽出）\nCatalog階段で空欄だけ埋めます。◎・既存ASINは上書きしません。HTMLは使いません。続けますか？',
    ui.ButtonSet.YES_NO
  );
  if (a !== ui.Button.YES) return;
  var out = amazonPasteCatalogFillEmptyAsinsImpl_(true);
  ui.alert('A.準備 書込 filled=' + out.filled + '\n\n' + out.msg + '\n\n次は A. Keepa取得。並べ替えはA末尾（P5）。');
}

/** SEO区切り。半角スペースを前後に挟む。再実行しても二重にしない。 */
var A_PREP_SEO_OPEN_ = '（(';
var A_PREP_SEO_CLOSE_ = '）)';
var A_PREP_SEO_MID_ = '・/／×+＋|｜、';

function menuAPrepSeoNameSpacesDryRun() {
  var out = aPrepSeoNameSpacesImpl_(false);
  SpreadsheetApp.getUi().alert('A.準備 商品名SEOスペース dry_run（未書込）\n' + out.msg);
}

function menuAPrepSeoNameSpacesApply() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.alert(
    'AI情報取得data の「商品名」に、（ ）・ などの前後へ半角スペースを入れます。マスタは触りません。続けますか？',
    ui.ButtonSet.YES_NO
  );
  if (a !== ui.Button.YES) return;
  var out = aPrepSeoNameSpacesImpl_(true);
  ui.alert('A.準備 商品名SEOスペース 書込\n' + out.msg);
}

function aPrepSeoPadName_(raw) {
  var s = String(raw == null ? '' : raw);
  if (!s) return s;
  var open = A_PREP_SEO_OPEN_;
  var close = A_PREP_SEO_CLOSE_;
  var mid = A_PREP_SEO_MID_;
  var out = '';
  var i;
  function lastSpace() {
    return out.length > 0 && out.charAt(out.length - 1) === ' ';
  }
  for (i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    var next = i + 1 < s.length ? s.charAt(i + 1) : '';
    if (open.indexOf(c) >= 0) {
      out += c;
      if (next && next !== ' ' && close.indexOf(next) < 0) out += ' ';
    } else if (close.indexOf(c) >= 0) {
      if (out.length && !lastSpace() && open.indexOf(out.charAt(out.length - 1)) < 0) out += ' ';
      out += c;
      if (next && next !== ' ' && close.indexOf(next) < 0 && open.indexOf(next) < 0) out += ' ';
    } else if (mid.indexOf(c) >= 0) {
      if (out.length && !lastSpace()) out += ' ';
      out += c;
      if (next && next !== ' ') out += ' ';
    } else {
      out += c;
    }
  }
  return out;
}

/**
 * @param {boolean} write
 * @return {{changed:number, skipped:number, msg:string}}
 */
function aPrepSeoNameSpacesImpl_(write) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TARGET_SHEET_NAME);
  var stats = { changed: 0, skipped: 0, msg: '' };
  if (!sh) {
    stats.msg = TARGET_SHEET_NAME + ' がありません。';
    return stats;
  }
  var last = sh.getLastRow();
  if (last < 2) {
    stats.msg = 'データ行がありません。';
    return stats;
  }
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = -1;
  var h;
  for (h = 0; h < hdr.length; h++) {
    if (String(hdr[h] || '').trim() === '商品名') { col = h; break; }
  }
  if (col < 0) {
    stats.msg = 'ヘッダー「商品名」がありません。';
    return stats;
  }
  var rng = sh.getRange(2, col + 1, last - 1, 1);
  var vals = rng.getValues();
  var formulas = rng.getFormulas();
  var examples = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    if (formulas[r][0]) {
      stats.skipped++;
      continue;
    }
    var cur = vals[r][0];
    if (cur == null || String(cur).trim() === '') {
      stats.skipped++;
      continue;
    }
    var next = aPrepSeoPadName_(cur);
    if (next === String(cur)) {
      stats.skipped++;
      continue;
    }
    stats.changed++;
    if (examples.length < 5) examples.push('行' + (r + 2) + ': ' + String(cur) + ' → ' + next);
    if (write) vals[r][0] = next;
    Logger.log('[A.準備SEO] 行' + (r + 2) + ' "' + String(cur) + '" → "' + next + '"');
  }
  if (write && stats.changed) rng.setValues(vals);
  stats.msg = (write ? '書込' : 'dry') + ' changed=' + stats.changed + ' skip=' + stats.skipped +
    '\n区切り=（）()・/／×+＋|｜、\n' + examples.join('\n');
  Logger.log('[A.準備SEO] ' + stats.msg.replace(/\n/g, ' | '));
  return stats;
}

function amazonPasteP2EvalScore_(ev) {
  var s = String(ev || '').trim();
  if (s === '◎') return 100;
  var m = s.match(/^(\d+)%?$/);
  return m ? parseInt(m[1], 10) : null;
}

function amazonPasteP2HasKeepa_(row, colEval, colPrice, colSet) {
  if (amazonPasteP2EvalScore_(row[colEval]) != null) return true;
  if (String(row[colPrice] || '').trim()) return true;
  if (String(row[colSet] || '').trim()) return true;
  return false;
}

function amazonPasteP2Classify_(title, ev, asin, row, colEval, colPrice, colSet, maker) {
  if (String(asin || '').trim() && !amazonPasteP2HasKeepa_(row, colEval, colPrice, colSet)) {
    return { kind: '未属性', reason: 'Keepa未取得' };
  }
  var t = (typeof normalizeFullwidthDigits_ === 'function')
    ? normalizeFullwidthDigits_(String(title || '')) : String(title || '');
  if (/ふるさと納税|返礼品/.test(t)) return { kind: '非候補', reason: 'ふるさと' };
  if (/よりどり|種類が選べ|選べるセット|から選択/.test(t)) return { kind: '非候補', reason: '選択式' };
  if (/中古/.test(t)) return { kind: '非候補', reason: '中古' };
  var parsed = (typeof parseSetCountFromItemNameWithSource === 'function')
    ? parseSetCountFromItemNameWithSource(t) : null;
  if (/各\s*\d+\s*袋/.test(t) && !(parsed && parsed.setCount)) return { kind: '非候補', reason: '各N袋' };
  if (!amazonPasteTitleHasMaker_(t, maker)) return { kind: '非候補', reason: 'メーカー無し' };
  return { kind: '候補', reason: '' };
}

function amazonPasteP2Bag_(title, setCell) {
  var t = (typeof normalizeFullwidthDigits_ === 'function')
    ? normalizeFullwidthDigits_(String(title || '')) : String(title || '');
  var parsed = (typeof parseSetCountFromItemNameWithSource === 'function')
    ? parseSetCountFromItemNameWithSource(t) : null;
  if (parsed && parsed.setCount) return parsed.setCount;
  if (/各\s*\d+\s*袋/.test(t)) return null;
  var n = parseFloat(String(setCell || '').replace(/,/g, ''));
  return n >= 1 ? n : null;
}

function amazonPasteP2Tag_(old, kind, reason) {
  var s = String(old || '').replace(/^\[機械[^\]]*\]\s*/, '');
  var tag = '[機械]' + kind + (reason ? ':' + reason : '');
  return s ? (tag + ' ' + s) : tag;
}

/**
 * @param {boolean} write
 * @return {{ok:boolean, msg:string}}
 */
function amazonPasteP5RankAfterAEnabled_() {
  return getBoolScriptProperty_(PROP_AMAZON_PASTE_P5_RANK_AFTER_A_ENABLED, true);
}

function keepaPasteP5RankAfterA_(ss) {
  var runId = 'pasteP5-' + Utilities.getUuid().slice(0, 8);
  if (!amazonPasteP5RankAfterAEnabled_()) {
    Logger.log('[AmazonPasteP5] runId=' + runId + ' state=DONE skip property off');
    return;
  }
  Logger.log('[AmazonPasteP5] runId=' + runId + ' state=RUNNING write=true');
  var out = amazonPasteP2RankImpl_(true, ss);
  Logger.log('[AmazonPasteP5] runId=' + runId + ' state=DONE ok=' + !!(out && out.ok) + ' ' + ((out && out.msg) || ''));
}

/**
 * @param {boolean} write
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} optSs
 * @return {{ok:boolean, msg:string}}
 */
function amazonPasteP2RankImpl_(write, optSs) {
  var runId = 'pasteP2-' + Utilities.getUuid().slice(0, 8);
  Logger.log('[AmazonPasteP2] runId=' + runId + ' state=RUNNING write=' + !!write);
  var ss = optSs || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  if (!sheet) return { ok: false, msg: 'ASIN貼り付けシートなし' };
  var blockCols = ASIN_PASTE_BLOCK_HEADERS.length;
  var values = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();
  var notes = [];
  var b;
  for (b = 0; b < ASIN_PASTE_DEFAULT_BLOCKS; b++) {
    var start = b * blockCols;
    var colAsin = start + 1;
    var colName = start + 2;
    var colEval = start + 3;
    var colPrice = start + 4;
    var colSet = start + 5;
    var colTag = start + 7;
    var maker = '';
    var jan = getJanFromAiDataForAsinPasteBlock_(ss, b);
    var aiSheet = ss.getSheetByName(AI_SHEET_NAME_FOR_ASIN_PASTE);
    if (aiSheet) {
      var map = getAiColumnIndexMap_(aiSheet.getRange(1, 1, 1, aiSheet.getLastColumn()).getValues()[0]);
      if (map['メーカー名'] !== undefined) {
        maker = String(aiSheet.getRange(2 + b, map['メーカー名'] + 1).getValue() || '').trim();
      }
    }
    var lastUsed = 1;
    var r;
    for (r = 2; r < values.length; r++) {
      var row = values[r] || [];
      var has = false;
      var c;
      for (c = 0; c < blockCols; c++) {
        if (String(row[start + c] || '').trim()) has = true;
      }
      if (has) lastUsed = r;
    }
    if (lastUsed < 2) {
      notes.push('b' + b + ' skip empty');
      continue;
    }
    var items = [];
    for (r = 2; r <= lastUsed; r++) {
      var rec = {
        idx: r,
        asin: String((values[r] || [])[colAsin] || '').trim(),
        title: String((values[r] || [])[colName] || ''),
        ev: (values[r] || [])[colEval],
        price: (values[r] || [])[colPrice],
        setc: (values[r] || [])[colSet]
      };
      if (!rec.asin) {
        rec.kind = 'blank';
        rec.score = -1;
        rec.bag = null;
      } else {
        var cl = amazonPasteP2Classify_(rec.title, rec.ev, rec.asin, values[r] || [], colEval, colPrice, colSet, maker);
        rec.kind = cl.kind;
        rec.reason = cl.reason;
        rec.score = amazonPasteP2EvalScore_(rec.ev);
        rec.bag = amazonPasteP2Bag_(rec.title, rec.setc);
      }
      items.push(rec);
    }
    var byBag = {};
    items.forEach(function (it) {
      if (it.kind !== '候補' || !it.bag) return;
      var p = parseFloat(String(it.price || '').replace(/,/g, ''));
      if (!(p > 0)) return;
      it.unit = p / it.bag;
      if (!byBag[it.bag]) byBag[it.bag] = [];
      byBag[it.bag].push(it);
    });
    Object.keys(byBag).forEach(function (k) {
      var grp = byBag[k];
      if (grp.length < 2) return;
      grp.forEach(function (it) {
        var others = [];
        grp.forEach(function (x) { if (x !== it && x.unit) others.push(x.unit); });
        if (!others.length) return;
        others.sort(function (a, b) { return a - b; });
        var med = others.length % 2 ? others[(others.length - 1) / 2]
          : (others[others.length / 2 - 1] + others[others.length / 2]) / 2;
        if (med && it.unit >= med * 2) {
          it.kind = '非候補';
          it.reason = '単価2倍';
        }
      });
    });
    items.sort(function (a, b) {
      function rank(it) {
        if (it.kind === '候補') return it.bag ? 0 : 1;
        if (it.kind === '未属性') return 2;
        if (it.kind === '非候補') return 3;
        return 4;
      }
      var ra = rank(a);
      var rb = rank(b);
      if (ra !== rb) return ra - rb;
      if (ra === 0 && a.bag !== b.bag) return a.bag - b.bag;
      var sa = a.score != null ? a.score : -1;
      var sb = b.score != null ? b.score : -1;
      return sb - sa;
    });
    var nCand = 0;
    var nBad = 0;
    var nPend = 0;
    if (write) {
      var blockW = [];
      var blockF = [];
      var i;
      for (i = 0; i < items.length; i++) {
        var src = items[i].idx;
        var vrow = (values[src] || []).slice(start, start + blockCols);
        var frow = (formulas[src] || []).slice(start, start + blockCols);
        var kind = items[i].kind;
        if (kind === '候補') nCand++;
        else if (kind === '非候補') nBad++;
        else if (kind === '未属性') nPend++;
        if (kind && kind !== 'blank') {
          vrow[7] = amazonPasteP2Tag_(vrow[7], kind, items[i].reason || '');
        }
        var col;
        for (col = 0; col < blockCols; col++) {
          if (frow[col] && String(frow[col]).trim().indexOf('=') === 0) vrow[col] = frow[col];
        }
        blockW.push(vrow);
      }
      sheet.getRange(3, start + 1, blockW.length, blockCols).setValues(blockW);
    } else {
      items.forEach(function (it) {
        if (it.kind === '候補') nCand++;
        else if (it.kind === '非候補') nBad++;
        else if (it.kind === '未属性') nPend++;
      });
    }
    var line = 'b' + b + ' jan=' + jan + ' cand=' + nCand + ' pending=' + nPend + ' reject=' + nBad;
    notes.push(line);
    Logger.log('[AmazonPasteP2] runId=' + runId + ' ' + line);
  }
  Logger.log('[AmazonPasteP2] runId=' + runId + ' state=DONE');
  return { ok: true, msg: notes.join('\n') };
}

function menuAmazonPasteP2RankDryRun() {
  var out = amazonPasteP2RankImpl_(false);
  SpreadsheetApp.getUi().alert('P2 dry_run（未並べ替え）\n\n' + out.msg);
}

function menuAmazonPasteP2RankApply() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('ブロック内で候補を上・非候補を下へ並べます。評価◎は消しません。行は削除しません。', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var out = amazonPasteP2RankImpl_(true);
  ui.alert('P2 書込\n\n' + out.msg);
}

function amazonPasteP3JanDigits_(v) {
  var s = String(v || '').trim();
  if (s.slice(-2) === '.0') s = s.slice(0, -2);
  return s.replace(/\D/g, '');
}

function amazonPasteP3ExcludeTitle_(title) {
  var t = String(title || '');
  return /ふるさと納税|返礼品|よりどり|種類が選べ|選べるセット|中古品|中古/.test(t);
}

function amazonPasteP3IsKindMix_(title) {
  var t = (typeof normalizeFullwidthDigits_ === 'function')
    ? normalizeFullwidthDigits_(String(title || '')) : String(title || '');
  var m = t.match(/(?:【)?(\d+)\s*種類?(?:】)?/);
  return !!(m && parseInt(m[1], 10) >= 2);
}

function amazonPasteP3ParseMasterSetQty_(v) {
  var m = String(v || '').trim().match(/^(\d+)/);
  if (!m) return null;
  var n = parseInt(m[1], 10);
  return n >= 1 ? n : null;
}

function amazonPasteP3CkTrue_(v) {
  if (v === true) return true;
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === '1';
}

/**
 * 貼付◎を JAN＋袋数で束ねる。同一JANはブロック横断で最安。
 */
function amazonPasteP3BuildClusters_(ss) {
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  var clusters = {};
  if (!sheet) return clusters;
  var blockCols = ASIN_PASTE_BLOCK_HEADERS.length;
  var values = sheet.getDataRange().getValues();
  var b;
  for (b = 0; b < ASIN_PASTE_DEFAULT_BLOCKS; b++) {
    var jan = amazonPasteP3JanDigits_(getJanFromAiDataForAsinPasteBlock_(ss, b));
    if (jan.length < 8) continue;
    var start = b * blockCols;
    var lastUsed = 1;
    var r;
    for (r = 2; r < values.length; r++) {
      var row = values[r] || [];
      var has = false;
      var c;
      for (c = 0; c < blockCols; c++) {
        if (String(row[start + c] || '').trim()) has = true;
      }
      if (has) lastUsed = r;
    }
    if (lastUsed < 2) continue;
    if (!clusters[jan]) clusters[jan] = {};
    for (r = 2; r <= lastUsed; r++) {
      var rec = values[r] || [];
      var ev = String(rec[start + 3] || '').trim();
      if (ev !== '◎') continue;
      var title = String(rec[start + 2] || '');
      if (amazonPasteP3ExcludeTitle_(title)) continue;
      if (amazonPasteP3IsKindMix_(title)) continue;
      var asin = String(rec[start + 1] || '').trim().toUpperCase();
      var price = parseFloat(String(rec[start + 4] || '').replace(/,/g, ''));
      if (!asin || !(price > 0)) continue;
      var bag = amazonPasteP2Bag_(title, rec[start + 5]);
      if (!bag) continue;
      var key = String(bag);
      var url = String(rec[start + 6] || '').trim() || ('https://www.amazon.co.jp/dp/' + asin);
      var prev = clusters[jan][key];
      if (prev && prev.priceIncl <= price) continue;
      clusters[jan][key] = { priceIncl: Math.round(price), asin: asin, url: url };
    }
  }
  return clusters;
}

function amazonPasteP3Plan_(ss, clusters) {
  var master = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!master) return { ok: false, msg: 'マスタなし', planned: [] };
  var vals = master.getDataRange().getValues();
  var headerRowIdx = -1;
  var hr;
  for (hr = 0; hr < Math.min(vals.length, 20); hr++) {
    if ((vals[hr] || []).indexOf(ANCHOR_HEADER_NAME) !== -1) {
      headerRowIdx = hr;
      break;
    }
  }
  if (headerRowIdx < 0) return { ok: false, msg: 'マスタヘッダーなし', planned: [] };
  var cmap = getColumnIndexMap(vals[headerRowIdx]);
  var colJan = cmap['JANコード'];
  var colSet = cmap[COL_MASTER_TOTAL_QTY];
  var colCk = cmap[CHECKBOX_HEADER_NAME];
  var colAmz = cmap[COL_COMPETITIVE_PRICE_AMAZON];
  var colAsin = cmap['競合店ASINコード'];
  var colUrl = cmap[COL_COMPETITOR_URL_AMAZON];
  if (colJan === undefined || colSet === undefined || colCk === undefined ||
      colAmz === undefined || colAsin === undefined || colUrl === undefined) {
    return { ok: false, msg: 'マスタ列不足（JAN/セット/CK/amazon/ASIN/URL）', planned: [] };
  }
  var planned = [];
  var mr;
  for (mr = headerRowIdx + 1; mr < vals.length; mr++) {
    if (!amazonPasteP3CkTrue_(vals[mr][colCk])) continue;
    var jan = amazonPasteP3JanDigits_(vals[mr][colJan]);
    var setQty = amazonPasteP3ParseMasterSetQty_(vals[mr][colSet]);
    if (!setQty) continue;
    var hit = (clusters[jan] || {})[String(setQty)];
    if (!hit) continue;
    planned.push({
      row: mr + 1,
      jan: jan,
      setQty: setQty,
      current: vals[mr][colAmz],
      newPrice: hit.priceIncl,
      newAsin: hit.asin,
      newUrl: hit.url,
      colAmz: colAmz + 1,
      colAsin: colAsin + 1,
      colUrl: colUrl + 1
    });
  }
  return { ok: true, msg: '', planned: planned, master: master };
}

/** 出品マスタ（Python MASTER_SS_ID と同じ） */
var AMAZON_PASTE_MASTER_SS_ID_ = '1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28';

/**
 * メニューは Active。clasp run は Active が無いのでマスタを openById。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ssOverride
 */
function amazonPasteGetSpreadsheet_(ssOverride) {
  if (ssOverride) return ssOverride;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  return SpreadsheetApp.openById(AMAZON_PASTE_MASTER_SS_ID_);
}

/**
 * @param {boolean} write
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ssOverride
 */
function amazonPasteP3ApplyImpl_(write, ssOverride) {
  var runId = 'pasteP3-' + Utilities.getUuid().slice(0, 8);
  Logger.log('[AmazonPasteP3] runId=' + runId + ' state=RUNNING write=' + !!write);
  var ss = amazonPasteGetSpreadsheet_(ssOverride);
  var clusters = amazonPasteP3BuildClusters_(ss);
  var jans = Object.keys(clusters);
  var i;
  for (i = 0; i < jans.length; i++) {
    var keys = Object.keys(clusters[jans[i]]).sort();
    Logger.log('[AmazonPasteP3] runId=' + runId + ' cluster ' + jans[i] + ' sets=' + keys.join(','));
  }
  var plan = amazonPasteP3Plan_(ss, clusters);
  if (!plan.ok) {
    Logger.log('[AmazonPasteP3] runId=' + runId + ' state=FAILED ' + plan.msg);
    return { ok: false, msg: plan.msg, n: 0 };
  }
  var lines = ['would_write=' + plan.planned.length];
  var p;
  for (p = 0; p < plan.planned.length && p < 40; p++) {
    var it = plan.planned[p];
    var line = 'row=' + it.row + ' JAN=' + it.jan + ' set=' + it.setQty +
      ' now=' + it.current + ' -> ' + it.newPrice + ' ' + it.newAsin;
    lines.push(line);
    Logger.log('[AmazonPasteP3] runId=' + runId + ' ' + line);
  }
  if (write) {
    var enabled = getBoolScriptProperty_(PROP_AMAZON_PASTE_P3_WRITE_ENABLED, false);
    if (!enabled) {
      Logger.log('[AmazonPasteP3] runId=' + runId + ' state=DONE write blocked property');
      return { ok: true, msg: 'AMAZON_PASTE_P3_WRITE_ENABLED が false（未設定含む）。マスタ未書。\n\n' + lines.join('\n'), n: 0 };
    }
    for (p = 0; p < plan.planned.length; p++) {
      it = plan.planned[p];
      plan.master.getRange(it.row, it.colAmz).setValue(it.newPrice);
      plan.master.getRange(it.row, it.colAsin).setValue(it.newAsin);
      plan.master.getRange(it.row, it.colUrl).setValue(it.newUrl);
    }
    Logger.log('[AmazonPasteP3] runId=' + runId + ' state=DONE written=' + plan.planned.length);
    return { ok: true, msg: 'written=' + plan.planned.length + '\n' + lines.join('\n'), n: plan.planned.length };
  }
  Logger.log('[AmazonPasteP3] runId=' + runId + ' state=DONE write=false');
  return { ok: true, msg: lines.join('\n'), n: plan.planned.length };
}

function menuAmazonPasteP3MasterDryRun() {
  var out = amazonPasteP3ApplyImpl_(false);
  SpreadsheetApp.getUi().alert('P3 dry_run（マスタ未書）\n\n' + out.msg);
}

function menuAmazonPasteP3MasterApply() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('出品CK行の競合価格amazon / 競合店ASIN / URL を書きます。楽天Yahoo列は触りません。', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var out = amazonPasteP3ApplyImpl_(true);
  ui.alert('P3 書込\n\n' + out.msg);
}

var KEEPA_FULL_HEADERS_ = [
  '取得日時', '目的', 'ASIN', '商品コード: EAN', '商品名', '画像', '製造者', 'ブランド', '親ASIN',
  'URL: Amazon', 'URL: Keepa',
  'Buy Box: 現在価格', 'Buy Box: 30 日平均', 'Buy Box: 90 日平均',
  'Amazon: 現在価格', 'Amazon: 30 日平均', 'Amazon: 90 日平均',
  '新品: 90 日平均', '参考価格: 90 日平均', '売れ筋ランキング: 90 日平均',
  '月間売上', 'レビュー: 評価', 'レビュー: 評価件数', '発売日',
  'カテゴリ: ルート', 'カテゴリ: ツリー', 'アイテム数', 'パッケージ数量',
  '梱包_L_cm', '梱包_W_cm', '梱包_H_cm', '梱包_重量_g',
  'FBA手数料', 'BuyBoxセラー', 'BuyBox_FBA', '価格指紋', '生JSON'
];

function keepaFullWriteEnabled_() {
  return getBoolScriptProperty_(PROP_KEEPA_FULL_WRITE_ENABLED, false);
}

function keepaFullStripCsv_(v) {
  var i;
  var k;
  if (v && typeof v === 'object') {
    if (Object.prototype.toString.call(v) === '[object Array]') {
      var a = [];
      for (i = 0; i < v.length; i++) a.push(keepaFullStripCsv_(v[i]));
      return a;
    }
    var o = {};
    for (k in v) {
      if (!Object.prototype.hasOwnProperty.call(v, k) || k === 'csv') continue;
      o[k] = keepaFullStripCsv_(v[k]);
    }
    return o;
  }
  return v;
}

function keepaFullFingerprint_(p) {
  var stats = (p && p.stats) || {};
  var cur = stats.current || [];
  var amazon = cur.length > 0 ? cur[0] : '';
  var buybox = cur.length > 18 ? cur[18] : '';
  return String(buybox) + '|' + String(amazon) + '|' + String((p && p.title) || '');
}

function keepaFullPlanActions_(existingRows, products) {
  var rows = (existingRows || []).slice();
  var actions = [];
  var i;
  for (i = 0; i < (products || []).length; i++) {
    var p = products[i] || {};
    var asin = String(p.asin || '').trim().toUpperCase();
    if (!asin) {
      actions.push('skip_no_asin');
      continue;
    }
    var fp = keepaFullFingerprint_(p);
    var latest = null;
    var j;
    for (j = 0; j < rows.length; j++) {
      if (String((rows[j] && rows[j].ASIN) || '').trim().toUpperCase() === asin) latest = rows[j];
    }
    if (latest && String(latest['価格指紋'] || '') === fp) {
      actions.push('skip_same_fp');
      continue;
    }
    actions.push('append');
    rows.push({ ASIN: asin, '価格指紋': fp });
  }
  return actions;
}

function keepaFullSelfTest_() {
  var p1 = { asin: 'B0LIST0001', title: '出品A', csv: [[1]], stats: { current: [10].concat([-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1], [11]) } };
  var pBad = { title: 'noasin' };
  var p2 = { asin: 'B0LIST0001', title: '出品A', csv: [[2]], stats: { current: [99].concat([-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1], [11]) } };
  var stripped = keepaFullStripCsv_(p1);
  if (stripped.csv) throw new Error('csv remained');
  var acts = keepaFullPlanActions_([], [p1, p1, pBad, p2]);
  var want = ['append', 'skip_same_fp', 'skip_no_asin', 'append'];
  if (acts.join(',') !== want.join(',')) throw new Error('plan=' + acts.join(','));
  var live = [];
  var hi;
  for (hi = 0; hi < KEEPA_FULL_HEADERS_.length; hi++) {
    if (KEEPA_FULL_HEADERS_[hi] !== '目的') live.push(KEEPA_FULL_HEADERS_[hi]);
  }
  live.push('目的');
  var rec = keepaFullProductToRec_(p1, 't', '出品');
  var aligned = keepaFullRowForHeaders_(rec, live);
  if (aligned[1] !== 'B0LIST0001') throw new Error('asin_col');
  if (aligned[aligned.length - 1] !== '出品') throw new Error('purpose_last');
  return 'ok ' + acts.join(',') + ' hdr=purpose_last';
}

function keepaFullAfterKeepaCache_(rawProducts) {
  var runId = 'keepaFull-' + Utilities.getUuid().slice(0, 8);
  if (!keepaFullWriteEnabled_()) {
    Logger.log('[KeepaFull] runId=' + runId + ' state=DONE skip property off');
    return;
  }
  Logger.log('[KeepaFull] runId=' + runId + ' state=RUNNING write=true n=' + ((rawProducts && rawProducts.length) || 0));
  var out = keepaFullUpsertImpl_(rawProducts, true, '出品');
  Logger.log('[KeepaFull] runId=' + runId + ' state=DONE ' + out.msg);
}

function keepaFullUpsertImpl_(rawProducts, write, purpose) {
  if (!rawProducts || !rawProducts.length) {
    return { ok: true, msg: 'rawなし', n: 0 };
  }
  var props = PropertiesService.getScriptProperties();
  var id = String(props.getProperty(PROP_COMPETITOR_SS_ID) || '').trim();
  if (!id) return { ok: false, msg: 'COMPETITOR_SS_ID なし', n: 0 };
  var destSs = SpreadsheetApp.openById(id);
  var dest = destSs.getSheetByName(COMPETITOR_KEEPA_FULL_SHEET_);
  if (!dest) dest = destSs.insertSheet(COMPETITOR_KEEPA_FULL_SHEET_);
  if (dest.getLastRow() === 0) {
    dest.getRange(1, 1, 1, KEEPA_FULL_HEADERS_.length).setValues([KEEPA_FULL_HEADERS_]);
    dest.getRange(1, 1, 1, KEEPA_FULL_HEADERS_.length).setFontWeight('bold');
  }
  var lastCol = Math.max(dest.getLastColumn(), KEEPA_FULL_HEADERS_.length);
  var hdr = dest.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxAsin = -1;
  var idxFp = -1;
  var h;
  for (h = 0; h < hdr.length; h++) {
    if (String(hdr[h] || '').trim() === 'ASIN') idxAsin = h;
    if (String(hdr[h] || '').trim() === '価格指紋') idxFp = h;
  }
  var existing = [];
  if (dest.getLastRow() >= 2 && idxAsin >= 0) {
    var vals = dest.getRange(2, 1, dest.getLastRow() - 1, lastCol).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      existing.push({
        ASIN: String(vals[r][idxAsin] || ''),
        '価格指紋': idxFp >= 0 ? String(vals[r][idxFp] || '') : ''
      });
    }
  }
  var actions = keepaFullPlanActions_(existing, rawProducts);
  var i;
  var n = 0;
  if (write) {
    var now = Utilities.formatDate(new Date(), 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'");
    for (i = 0; i < rawProducts.length; i++) {
      if (actions[i] !== 'append') continue;
      var recObj = keepaFullProductToRec_(rawProducts[i], now, purpose || '出品');
      var padded = keepaFullRowForHeaders_(recObj, hdr);
      dest.getRange(dest.getLastRow() + 1, 1, 1, hdr.length).setValues([padded]);
      n++;
    }
  }
  var msg = 'write=' + !!write + ' append=' + n + ' plan=' + actions.join(',');
  Logger.log('[KeepaFull] ' + msg);
  return { ok: true, msg: msg, n: n };
}

function keepaFullStat_(cur, idx) {
  if (!cur || cur.length <= idx) return '';
  var v = cur[idx];
  if (v == null || v === -1) return '';
  return String(v);
}

function keepaFullRowForHeaders_(rec, hdr) {
  var out = [];
  var i;
  for (i = 0; i < (hdr || []).length; i++) {
    var k = String(hdr[i] || '').trim();
    out.push(k && rec && rec[k] != null ? rec[k] : '');
  }
  return out;
}

function keepaFullProductToRec_(p, fetchedAt, purpose) {
  p = p || {};
  var stats = p.stats || {};
  var cur = stats.current || [];
  var asin = String(p.asin || '').trim().toUpperCase();
  var eans = p.eanList || [];
  var ean = eans.length ? eans[0] : (p.ean || '');
  var raw = JSON.stringify(keepaFullStripCsv_(p));
  if (raw.length > 45000) raw = raw.substring(0, 45000);
  var rec = {};
  rec['取得日時'] = fetchedAt;
  rec['目的'] = purpose || '出品';
  rec['ASIN'] = asin;
  rec['商品コード: EAN'] = ean == null ? '' : String(ean);
  rec['商品名'] = String(p.title || '');
  rec['製造者'] = String(p.manufacturer || '');
  rec['ブランド'] = String(p.brand || '');
  rec['親ASIN'] = String(p.parentAsin || '');
  rec['URL: Amazon'] = asin ? 'https://www.amazon.co.jp/dp/' + asin : '';
  rec['URL: Keepa'] = asin ? 'https://keepa.com/#!product/5-' + asin : '';
  rec['Buy Box: 現在価格'] = keepaFullStat_(cur, 18);
  rec['Amazon: 現在価格'] = keepaFullStat_(cur, 0);
  rec['価格指紋'] = keepaFullFingerprint_(p);
  rec['生JSON'] = raw;
  return rec;
}

function keepaFullProductToRow_(p, fetchedAt, purpose, hdr) {
  return keepaFullRowForHeaders_(keepaFullProductToRec_(p, fetchedAt, purpose), hdr || KEEPA_FULL_HEADERS_);
}

function menuKeepaFullUpsertDryRun() {
  var msg = keepaFullSelfTest_();
  Logger.log('[KeepaFull] dry self-test ' + msg);
  SpreadsheetApp.getUi().alert('Keepaフル dry（専用非書）\nProperty 未設定=OFF。A経由の書込のみ。\nself-test: ' + msg);
}

function menuKeepaFullUpsertApply() {
  var ui = SpreadsheetApp.getUi();
  if (!keepaFullWriteEnabled_()) {
    ui.alert('KEEPA_FULL_WRITE_ENABLED が false（未設定含む）。専用非書。Aのあとも書きません。');
    return;
  }
  ui.alert('書込はメニューA成功後の自動のみです。ダミー商品は専用へ書きません。Property=ON なら次のAで追記します。');
}

function keepaFullAcquiredTime_(raw) {
  if (raw == null || raw === '') return null;
  if (Object.prototype.toString.call(raw) === '[object Date]' && !isNaN(raw.getTime())) return raw;
  if (typeof raw === 'number' && isFinite(raw)) {
    if (raw > 1e11) return new Date(raw);
    if (raw > 20000 && raw < 80000) return new Date(Math.round((raw - 25569) * 86400000));
  }
  var s = String(raw).trim();
  var t = Date.parse(s);
  if (!isNaN(t)) return new Date(t);
  return null;
}

function keepaFullIsGetNeeded_(acquired, now) {
  var dt = keepaFullAcquiredTime_(acquired);
  if (!dt) return true;
  now = now || new Date();
  return (now.getTime() - dt.getTime()) > 90 * 24 * 60 * 60 * 1000;
}

function keepaFullHasGateStats_(rawJson) {
  try {
    var p = JSON.parse(String(rawJson || '{}'));
    var stats = (p && p.stats) || {};
    var keys = ['avg90', 'avg30', 'avg'];
    var ki;
    for (ki = 0; ki < keys.length; ki++) {
      var arr = stats[keys[ki]];
      if (!arr || !arr.length) continue;
      var idx;
      var idxs = [18, 1, 3];
      for (idx = 0; idx < idxs.length; idx++) {
        var i = idxs[idx];
        if (i >= arr.length) continue;
        var v = arr[i];
        if (v == null || v === -1) continue;
        if (Number(v) >= 0) return true;
      }
    }
  } catch (e) {}
  return false;
}

function keepaFullWarehouseGetNeeded_(row, now) {
  if (!row || keepaFullIsGetNeeded_(row['取得日時'], now)) return true;
  return !keepaFullHasGateStats_(row['生JSON']);
}

function keepaFullClassifyNeedGet_(asins, existingRows, now) {
  var need = [];
  var skip = [];
  var seen = {};
  var i;
  for (i = 0; i < (asins || []).length; i++) {
    var asin = String(asins[i] || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{10}$/.test(asin) || seen[asin]) continue;
    seen[asin] = true;
    var latest = null;
    var j;
    for (j = 0; j < (existingRows || []).length; j++) {
      if (String((existingRows[j] && existingRows[j].ASIN) || '').trim().toUpperCase() === asin) {
        latest = existingRows[j];
      }
    }
    if (latest && !keepaFullWarehouseGetNeeded_(latest, now)) skip.push(asin);
    else need.push(asin);
  }
  return { need_get: need, skip_fresh: skip };
}

function keepaFullCollectPasteAsins_(ss) {
  var sheet = ss.getSheetByName(ASIN_PASTE_SHEET_NAME);
  var out = [];
  if (!sheet) return out;
  var values = sheet.getDataRange().getValues();
  var blockCols = ASIN_PASTE_BLOCK_HEADERS.length;
  var b;
  var seen = {};
  for (b = 0; b < ASIN_PASTE_DEFAULT_BLOCKS; b++) {
    var col = b * blockCols + 1;
    var r;
    for (r = 2; r < values.length; r++) {
      var a = String((values[r] || [])[col] || '').trim().toUpperCase();
      if (!/^[A-Z0-9]{10}$/.test(a) || seen[a]) continue;
      seen[a] = true;
      out.push(a);
    }
  }
  return out;
}

function keepaFullW4SelfTest_() {
  var now = new Date('2026-08-15T12:00:00Z');
  var okj = '{"stats":{"avg90":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,100]}}';
  var rows = [
    { ASIN: 'B0FRESH001', '取得日時': '2026-08-01T00:00:00Z', '商品名': 'フル新鮮', '生JSON': okj },
    { ASIN: 'B0STALE001', '取得日時': '2026-04-01T00:00:00Z' },
    { ASIN: 'B0EMPTY001', '取得日時': '2026-08-01T00:00:00Z', '生JSON': '{}' }
  ];
  var out = keepaFullClassifyNeedGet_(['B0FRESH001', 'B0FRESH001', 'B0STALE001', 'B0MISS0001'], rows, now);
  if (out.skip_fresh.join(',') !== 'B0FRESH001') throw new Error('skip');
  if (out.need_get.join(',') !== 'B0STALE001,B0MISS0001') throw new Error('need');
  var plan = keepaFullPlanAFetch_(['B0CACHE000', 'B0FRESH001', 'B0STALE001', 'B0MISS0001', 'B0EMPTY001'], rows, { B0CACHE000: true }, now);
  if (plan.hydrate.join(',') !== 'B0FRESH001') throw new Error('hydrate=' + plan.hydrate.join(','));
  if (plan.fetch.join(',') !== 'B0STALE001,B0MISS0001,B0EMPTY001') throw new Error('fetch=' + plan.fetch.join(','));
  return 'ok skip=1 need=2 hydrate=1 fetch=3';
}

function keepaFullReadEnabled_() {
  return getBoolScriptProperty_(PROP_KEEPA_FULL_READ_ENABLED, false);
}

function keepaFullPlanAFetch_(asins, existingRows, cacheMap, now) {
  var hydrate = [];
  var fetch = [];
  var seen = {};
  var i;
  for (i = 0; i < (asins || []).length; i++) {
    var asin = String(asins[i] || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{10}$/.test(asin) || seen[asin]) continue;
    seen[asin] = true;
    if (cacheMap && cacheMap[asin]) continue;
    var latest = null;
    var j;
    for (j = 0; j < (existingRows || []).length; j++) {
      if (String((existingRows[j] && existingRows[j].ASIN) || '').trim().toUpperCase() === asin) latest = existingRows[j];
    }
    if (latest && !keepaFullWarehouseGetNeeded_(latest, now)) hydrate.push(asin);
    else fetch.push(asin);
  }
  return { hydrate: hydrate, fetch: fetch };
}

function keepaFullRowToCacheEntry_(row) {
  row = row || {};
  var priceRaw = row['Buy Box: 現在価格'];
  var price = (priceRaw === '' || priceRaw == null) ? null : parseFloat(String(priceRaw).replace(/,/g, ''), 10);
  if (typeof price === 'number' && isNaN(price)) price = null;
  return {
    asin: String(row.ASIN || '').trim().toUpperCase(),
    title: String(row['商品名'] || ''),
    price: price,
    imageUrl: String(row['画像'] || ''),
    setCount: null,
    setCountFromTitle: null,
    setCountReason: 'keepa_full',
    ean: row['商品コード: EAN'] ? String(row['商品コード: EAN']) : null,
    brand: String(row['ブランド'] || ''),
    manufacturer: String(row['製造者'] || ''),
    packageLCm: row['梱包_L_cm'] || '',
    packageWCm: row['梱包_W_cm'] || '',
    packageHCm: row['梱包_H_cm'] || '',
    packageWeightG: row['梱包_重量_g'] || ''
  };
}

var keepaFullRowsMemo_ = null;

function keepaFullLoadExistingRows_() {
  if (keepaFullRowsMemo_) return keepaFullRowsMemo_;
  keepaFullRowsMemo_ = [];
  var id = String(PropertiesService.getScriptProperties().getProperty(PROP_COMPETITOR_SS_ID) || '').trim();
  if (!id) return keepaFullRowsMemo_;
  var dest = SpreadsheetApp.openById(id).getSheetByName(COMPETITOR_KEEPA_FULL_SHEET_);
  if (!dest || dest.getLastRow() < 2) return keepaFullRowsMemo_;
  var lastCol = dest.getLastColumn();
  var hdr = dest.getRange(1, 1, 1, lastCol).getValues()[0];
  var vals = dest.getRange(2, 1, dest.getLastRow() - 1, lastCol).getValues();
  var r;
  var h;
  for (r = 0; r < vals.length; r++) {
    var rec = {};
    for (h = 0; h < hdr.length; h++) {
      rec[String(hdr[h] || '').trim()] = vals[r][h];
    }
    keepaFullRowsMemo_.push(rec);
  }
  return keepaFullRowsMemo_;
}

function keepaFullMaybeHydrateCachedMap_(asins, cachedMap) {
  if (!keepaFullReadEnabled_()) {
    Logger.log('[KeepaFullW4] hydrate skip property off');
    return 0;
  }
  var existing = keepaFullLoadExistingRows_();
  var plan = keepaFullPlanAFetch_(asins, existing, cachedMap, new Date());
  var n = 0;
  var i;
  for (i = 0; i < plan.hydrate.length; i++) {
    var asin = plan.hydrate[i];
    var latest = null;
    var j;
    for (j = 0; j < existing.length; j++) {
      if (String((existing[j] && existing[j].ASIN) || '').trim().toUpperCase() === asin) latest = existing[j];
    }
    if (!latest) continue;
    cachedMap[asin] = keepaFullRowToCacheEntry_(latest);
    n++;
  }
  Logger.log('[KeepaFullW4] hydrate=' + n + ' fetch=' + plan.fetch.length);
  return n;
}

function menuKeepaFullW4ReadDryRun() {
  var runId = 'keepaW4-' + Utilities.getUuid().slice(0, 8);
  Logger.log('[KeepaFullW4] runId=' + runId + ' state=RUNNING write=false get=false');
  var st = keepaFullW4SelfTest_();
  var ss = amazonPasteGetSpreadsheet_();
  var asins = keepaFullCollectPasteAsins_(ss);
  keepaFullRowsMemo_ = null;
  var existing = keepaFullLoadExistingRows_();
  var plan = keepaFullClassifyNeedGet_(asins, existing, new Date());
  var planA = keepaFullPlanAFetch_(asins, existing, {}, new Date());
  var msg = st + ' paste=' + asins.length + ' full=' + existing.length +
    ' skip_fresh=' + plan.skip_fresh.length + ' need_get=' + plan.need_get.length +
    ' hydrate=' + planA.hydrate.length + ' fetch=' + planA.fetch.length;
  Logger.log('[KeepaFullW4] runId=' + runId + ' state=DONE ' + msg);
  SpreadsheetApp.getUi().alert('W4 dry（GETしない・専用非書）\n' + msg);
}
