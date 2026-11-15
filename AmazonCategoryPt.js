/**
 * P4b — Amazon Product Type / Browse Node 提案書込
 *
 * 方針: docs/org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md §2
 * - 手数料列 amazon カテゴリーは触らない
 * - Keepa 新規取得なし
 * - 本線: 貼付◎ → ログ◎ → マスタ他店 → JAN複数ヒット → 楽天/Yahoo/商品名
 * - 同点は空欄にしない（楽天語→商品名→名称投票→残票）
 * - 自ASINは単品では数えない（ミックス親だけ可）
 * - Browse未確定なら PT を書かない（手がかり0のときだけ）
 * - 空セルのみ書込。トグル既定 false
 *
 * Property:
 *   APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED … 既定 false
 *   APPROVAL_AMAZON_P4B_PT_MAX_PARENTS … 既定 3（1〜10）
 *   APPROVAL_AMAZON_P4B_PT_MAX_ASINS … 貼付◎ON時 既定10／OFF時5
 *   APPROVAL_AMAZON_P4B_PASTE_CIRCLE_VOTE_ENABLED … 未設定=true。falseで旧マスタ多数決
 */

var APPROVAL_AMAZON_P4B_PT_WRITE_PROP = 'APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED';
var APPROVAL_AMAZON_P4B_PT_MAX_PROP = 'APPROVAL_AMAZON_P4B_PT_MAX_PARENTS';
var APPROVAL_AMAZON_P4B_PT_MAX_ASINS_PROP = 'APPROVAL_AMAZON_P4B_PT_MAX_ASINS';
var APPROVAL_AMAZON_P4B_PASTE_CIRCLE_VOTE_PROP = 'APPROVAL_AMAZON_P4B_PASTE_CIRCLE_VOTE_ENABLED';
var AMAZON_P4B_SHELF_VOTE_THRESHOLD_ = 3;
var AMAZON_P4B_PT_HEADER_ = 'Amazon Product Type';
var AMAZON_P4B_BROWSE_HEADER_ = 'Amazon Browse Node';
/**
 * Catalog PT → 純正xlsm／棚にある PT（フォールバック）。
 * 本線は browseNodeId → SHELF browseIndex の preferredProductType。
 */
var AMAZON_P4B_PT_SHELF_ALIASES_ = {
  MEAT: 'GROCERY'
};

/** browse 表示文字列または ID から Node ID を抜く */
function amazonP4bExtractBrowseNodeId_(browseText) {
  var s = String(browseText || '').trim();
  if (!s) return '';
  var m = s.match(/\((\d{6,})\)\s*$/);
  if (m) return m[1];
  m = s.match(/\b(\d{6,})\b/);
  return m ? m[1] : '';
}

/**
 * @param {string} nodeId
 * @param {{byBrowseNode?:Object}=} shelf
 * @return {Object|null} browseIndex 行
 */
function amazonP4bResolvePreferredPtFromBrowse_(nodeId, shelf) {
  var nid = String(nodeId || '').trim();
  if (!nid) return null;
  var map = (shelf && shelf.byBrowseNode) ? shelf.byBrowseNode : null;
  if (!map) {
    try {
      if (typeof batchExportAmazonLoadShelfRegistry_ === 'function') {
        var loaded = batchExportAmazonLoadShelfRegistry_();
        if (loaded && loaded.ok) map = loaded.byBrowseNode || null;
      }
    } catch (e0) {
      map = null;
    }
  }
  if (!map || !map[nid]) return null;
  return map[nid];
}
var AMAZON_P4B_RIVAL_ASIN_HEADER_ = '競合店ASINコード';
var AMAZON_P4B_URL_COLS_ = ['競合AmazonページURL', '競合URL', 'Amazon URL', '商品URL'];

/**
 * メニュー 21-⑱: レ点親へ PT／browse 提案（空セルのみ）
 */
function menuAmazonP4bSuggestProductTypeBrowse() {
  return menuAmazonP4bSuggestProductTypeBrowse_({});
}

/**
 * D新規ゲート用: 指定親のうち PT 空 **または Browse 空** のものだけ silent で P4b。
 * Property 自動 ON はしない。対象がありトグル false なら ok=false。
 * @param {Object} masterCtx
 * @param {Array<{parentSku:string, parentRowIndex0:number, parentRow1:number, sampleChildRowIndex0:number, childRowIndexes0:number[], itemName:string}>} parents
 * @return {{ok:boolean, reason?:string, runId?:string, wrote?:number, skipped?:number, details?:Array, emptyCount?:number, ran?:boolean}}
 */
function amazonP4bSuggestEmptyPtParentsForNewGate_(masterCtx, parents) {
  var fn = 'amazonP4bSuggestEmptyPtParentsForNewGate_';
  var colPt = masterCtx.col[AMAZON_P4B_PT_HEADER_];
  var colBrowse = masterCtx.col[AMAZON_P4B_BROWSE_HEADER_];
  if (colPt == null || colBrowse == null) {
    return {
      ok: false,
      reason: 'マスタに「' + AMAZON_P4B_PT_HEADER_ + '」と「' +
        AMAZON_P4B_BROWSE_HEADER_ + '」列を追加してください。'
    };
  }
  var emptyParents = [];
  var i;
  for (i = 0; i < (parents || []).length; i++) {
    var p = parents[i];
    var pt = String(masterCtx.values[p.parentRowIndex0][colPt] || '').trim();
    var browse = String(masterCtx.values[p.parentRowIndex0][colBrowse] || '').trim();
    if (!pt || !browse) emptyParents.push(p);
  }
  if (!emptyParents.length) {
    Logger.log('[' + fn + '] state=DONE ran=0 reason=all_pt_browse_filled n=' + parents.length);
    return { ok: true, ran: false, emptyCount: 0, wrote: 0, skipped: 0, details: [] };
  }
  if (!getBoolScriptProperty_(APPROVAL_AMAZON_P4B_PT_WRITE_PROP, false)) {
    var off = '新規ゲート: PTまたはBrowseが空の親が' + emptyParents.length +
      '件あります。Script Property ' + APPROVAL_AMAZON_P4B_PT_WRITE_PROP +
      '=true にして再実行するか、親行の「Amazon Product Type」「Amazon Browse Node」を埋めてください。例=' +
      emptyParents[0].parentSku;
    Logger.log('[' + fn + '] state=FAILED ' + off);
    return { ok: false, reason: off, emptyCount: emptyParents.length, ran: false };
  }
  var maxParents = Math.max(1, Math.min(10,
    Math.floor(getNumberScriptProperty_(APPROVAL_AMAZON_P4B_PT_MAX_PROP, 3))));
  if (emptyParents.length > maxParents) {
    var over = '新規ゲート: PT/Browse未充足の親が' + emptyParents.length +
      '件で max=' + maxParents + '超。' + APPROVAL_AMAZON_P4B_PT_MAX_PROP +
      ' を上げるかレ点を減らしてください。';
    Logger.log('[' + fn + '] state=FAILED ' + over);
    return { ok: false, reason: over, emptyCount: emptyParents.length, ran: false };
  }
  return menuAmazonP4bSuggestProductTypeBrowse_({
    silent: true,
    parentsOverride: emptyParents,
    masterCtxOverride: masterCtx
  });
}

/**
 * @param {{silent?:boolean, parentsOverride?:Array, masterCtxOverride?:Object}=} opt
 */
function menuAmazonP4bSuggestProductTypeBrowse_(opt) {
  opt = opt || {};
  var silent = !!opt.silent;
  var fn = 'menuAmazonP4bSuggestProductTypeBrowse';
  var runId = amazonP4bNewRunId_();
  Logger.log('[' + fn + '] runId=' + runId + ' state=PENDING');

  if (!getBoolScriptProperty_(APPROVAL_AMAZON_P4B_PT_WRITE_PROP, false)) {
    var off = 'P4b書込が無効です。Script Property ' +
      APPROVAL_AMAZON_P4B_PT_WRITE_PROP + '=true にして再実行してください。';
    Logger.log('[' + fn + '] runId=' + runId + ' state=FAILED ' + off);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(off); } catch (e0) {}
    }
    return { ok: false, reason: off, runId: runId };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterCtx = opt.masterCtxOverride || amazonApprovalLv4LoadMasterContext_(ss);
  var colPt = masterCtx.col[AMAZON_P4B_PT_HEADER_];
  var colBrowse = masterCtx.col[AMAZON_P4B_BROWSE_HEADER_];
  if (colPt == null || colBrowse == null) {
    var noCol = 'マスタに「' + AMAZON_P4B_PT_HEADER_ + '」と「' +
      AMAZON_P4B_BROWSE_HEADER_ + '」列を追加してください。';
    Logger.log('[' + fn + '] runId=' + runId + ' state=FAILED ' + noCol);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(noCol); } catch (e1) {}
    }
    return { ok: false, reason: noCol, runId: runId };
  }

  var parents = opt.parentsOverride || amazonP4bCollectCheckedParents_(masterCtx);
  var maxParents = Math.max(1, Math.min(10,
    Math.floor(getNumberScriptProperty_(APPROVAL_AMAZON_P4B_PT_MAX_PROP, 3))));
  if (!opt.parentsOverride && parents.length > maxParents) {
    parents = parents.slice(0, maxParents);
  }
  if (!parents.length) {
    var noParent = '出品CK付きの子SKU（親付き）がありません。';
    Logger.log('[' + fn + '] runId=' + runId + ' state=FAILED ' + noParent);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(noParent); } catch (e2) {}
    }
    return { ok: false, reason: noParent, runId: runId };
  }

  var auth;
  try {
    auth = amazonSpapiPutAcquireAccess_();
  } catch (eAuth) {
    var authFail = 'LWA失敗: ' + String(eAuth && eAuth.message ? eAuth.message : eAuth);
    Logger.log('[' + fn + '] runId=' + runId + ' state=FAILED ' + authFail);
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(authFail); } catch (e3) {}
    }
    return { ok: false, reason: authFail, runId: runId };
  }

  var wrote = 0;
  var skipped = 0;
  var details = [];

  for (var i = 0; i < parents.length; i++) {
    var p = parents[i];
    var result = amazonP4bSuggestOneParent_(
      masterCtx, auth.creds, auth.accessToken, p, colPt, colBrowse, runId);
    details.push(result);
    if (result.wrote) wrote++;
    else skipped++;
    Logger.log('[' + fn + '] runId=' + runId +
      ' parent=' + p.parentSku + ' row=' + p.parentRow1 +
      ' wrote=' + (result.wrote ? 1 : 0) +
      ' pt=' + (result.productType || '') +
      ' browse=' + String(result.browse || '').substring(0, 80) +
      ' sources=' + (result.sources || []).join('+') +
      ' refAsin=' + (result.refAsin || '') +
      ' skip=' + (result.skipReason || '') +
      ' warns=' + ((result.warns || []).join(';') || ''));
  }

  var msg = 'P4b提案完了 runId=' + runId +
    '\n対象親=' + parents.length + ' 書込=' + wrote + ' スキップ=' + skipped +
    '\n（空セルのみ。手数料 amazon カテゴリーは未変更）';
  Logger.log('[' + fn + '] runId=' + runId + ' state=DONE wrote=' + wrote +
    ' skipped=' + skipped);
  if (!silent) {
    try { SpreadsheetApp.getUi().alert(msg); } catch (e4) {}
  }
  return {
    ok: true,
    runId: runId,
    wrote: wrote,
    skipped: skipped,
    details: details,
    ran: true,
    emptyCount: parents.length
  };
}

function amazonP4bNewRunId_() {
  var d = new Date();
  var p = function(n) { return (n < 10 ? '0' : '') + n; };
  var stamp = d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' +
    p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds());
  var hex = Utilities.getUuid().replace(/-/g, '').substring(0, 6);
  return 'P4B_PT_' + stamp + '_' + hex;
}

/**
 * @return {Array<{parentSku:string, parentRowIndex0:number, parentRow1:number, sampleChildRowIndex0:number, childRowIndexes0:number[], itemName:string}>}
 */
function amazonP4bCollectCheckedParents_(masterCtx) {
  if (typeof amazonCheckboxMainlineInspect_ !== 'function') {
    throw new Error('amazonCheckboxMainlineInspect_ がありません');
  }
  var inspected = amazonCheckboxMainlineInspect_(masterCtx, {
    includeNew: true,
    includeOffer: true
  });
  var byParent = {};
  var order = [];
  var list = inspected.newRows || [];
  for (var i = 0; i < list.length; i++) {
    var one = list[i];
    var ps = String(one.parentSku || '').trim();
    if (!ps) continue;
    if (!byParent[ps]) {
      byParent[ps] = {
        parentSku: ps,
        sampleChildRowIndex0: one.rowIndex0,
        childRowIndexes0: []
      };
      order.push(ps);
    }
    byParent[ps].childRowIndexes0.push(one.rowIndex0);
  }

  var out = [];
  for (var o = 0; o < order.length; o++) {
    var pack = byParent[order[o]];
    var parentRow = (typeof amazonSpapiExportFindParentRow_ === 'function')
      ? amazonSpapiExportFindParentRow_(masterCtx, pack.parentSku)
      : null;
    if (parentRow == null) {
      Logger.log('[AmazonP4b] parent row missing parentSku=' + pack.parentSku);
      continue;
    }
    var itemName = '';
    var nameCols = ['最終商品名amazon', '商品名amazon', '商品名案(Amazon)', '商品名'];
    for (var c = 0; c < nameCols.length; c++) {
      itemName = amazonApprovalLv4Cell_(masterCtx, parentRow, nameCols[c]);
      if (itemName) break;
    }
    if (!itemName) {
      itemName = amazonApprovalLv4Cell_(masterCtx, pack.sampleChildRowIndex0, '最終商品名amazon') ||
        amazonApprovalLv4Cell_(masterCtx, pack.sampleChildRowIndex0, '商品名');
    }
    out.push({
      parentSku: pack.parentSku,
      parentRowIndex0: parentRow,
      parentRow1: parentRow + 1,
      sampleChildRowIndex0: pack.sampleChildRowIndex0,
      childRowIndexes0: pack.childRowIndexes0,
      itemName: itemName
    });
  }
  return out;
}

function amazonP4bLooksLikeAsin_(s) {
  var t = String(s || '').trim().toUpperCase();
  if (typeof amazonApprovalLv4LooksLikeAsin_ === 'function' && amazonApprovalLv4LooksLikeAsin_(t)) {
    return true;
  }
  return /^[A-Z0-9]{10}$/.test(t);
}

function amazonP4bAsinFromUrl_(url) {
  if (typeof amazonApprovalLv4AsinFromUrl_ === 'function') {
    return amazonApprovalLv4AsinFromUrl_(url) || '';
  }
  var m = String(url || '').match(/\/(?:dp|gp\/product|product)\/([A-Z0-9]{10})/i);
  return m ? String(m[1]).toUpperCase() : '';
}

function amazonP4bFirstAsinCell_(masterCtx, rowIndex0, colName) {
  var raw = amazonApprovalLv4Cell_(masterCtx, rowIndex0, colName);
  if (amazonP4bLooksLikeAsin_(raw)) return String(raw).trim().toUpperCase();
  return '';
}

function amazonP4bFirstUrlOnRow_(masterCtx, rowIndex0) {
  for (var i = 0; i < AMAZON_P4B_URL_COLS_.length; i++) {
    var u = amazonApprovalLv4Cell_(masterCtx, rowIndex0, AMAZON_P4B_URL_COLS_[i]);
    if (u) return { url: u, col: AMAZON_P4B_URL_COLS_[i] };
  }
  return { url: '', col: '' };
}

/**
 * §2 #4: 競合店ASIN → 競合URL → 自ASIN
 * @return {{asin:string, source:string, rivalAsin:string, urlAsin:string, url:string, humanUrl:boolean, warns:string[]}}
 */
function amazonP4bResolveReferenceAsin_(masterCtx, parent) {
  var warns = [];
  var rows = [parent.parentRowIndex0].concat(parent.childRowIndexes0 || []);
  var seen = {};
  var uniqRows = [];
  for (var r = 0; r < rows.length; r++) {
    var ri = rows[r];
    if (seen[ri]) continue;
    seen[ri] = true;
    uniqRows.push(ri);
  }

  var rivalAsin = '';
  var rivalRow = -1;
  var i;
  for (i = 0; i < uniqRows.length; i++) {
    var rawRival = amazonApprovalLv4Cell_(masterCtx, uniqRows[i], AMAZON_P4B_RIVAL_ASIN_HEADER_);
    if (!rawRival) continue;
    if (amazonP4bLooksLikeAsin_(rawRival)) {
      rivalAsin = String(rawRival).trim().toUpperCase();
      rivalRow = uniqRows[i];
      break;
    }
    warns.push('rival_asin_format_invalid:' + String(rawRival).substring(0, 20));
  }

  var urlInfo = { url: '', col: '' };
  var urlAsin = '';
  for (i = 0; i < uniqRows.length; i++) {
    urlInfo = amazonP4bFirstUrlOnRow_(masterCtx, uniqRows[i]);
    if (!urlInfo.url) continue;
    urlAsin = amazonP4bAsinFromUrl_(urlInfo.url);
    if (urlAsin) break;
    warns.push('rival_url_no_asin:' + urlInfo.col);
  }

  if (rivalAsin && urlAsin && rivalAsin !== urlAsin) {
    warns.push('rival_asin_url_mismatch:' + rivalAsin + '!=' + urlAsin);
  }

  if (rivalAsin) {
    return {
      asin: rivalAsin,
      source: 'rival_asin',
      rivalAsin: rivalAsin,
      urlAsin: urlAsin,
      url: urlInfo.url || '',
      humanUrl: !!urlInfo.url,
      warns: warns,
      rowHint: rivalRow
    };
  }
  if (urlAsin) {
    return {
      asin: urlAsin,
      source: 'rival_url',
      rivalAsin: '',
      urlAsin: urlAsin,
      url: urlInfo.url || '',
      humanUrl: true,
      warns: warns,
      rowHint: -1
    };
  }

  var ownAsin = '';
  for (i = 0; i < uniqRows.length; i++) {
    ownAsin = amazonP4bFirstAsinCell_(masterCtx, uniqRows[i], 'ASINコード');
    if (ownAsin) break;
  }
  if (ownAsin) {
    return {
      asin: ownAsin,
      source: 'own_asin',
      rivalAsin: '',
      urlAsin: urlAsin,
      url: urlInfo.url || '',
      humanUrl: !!urlInfo.url,
      warns: warns,
      rowHint: -1
    };
  }

  return {
    asin: '',
    source: '',
    rivalAsin: '',
    urlAsin: '',
    url: urlInfo.url || '',
    humanUrl: !!urlInfo.url,
    warns: warns,
    rowHint: -1
  };
}

function amazonP4bNormalizeDigits_(s) {
  return String(s || '').replace(/\D/g, '');
}

/**
 * 簡易タイトル類似（WARN用）。共通の4文字以上の連続部分があればOK扱い。
 */
function amazonP4bTitleLooksRelated_(masterName, catalogTitle) {
  var a = String(masterName || '').replace(/\s+/g, '').toLowerCase();
  var b = String(catalogTitle || '').replace(/\s+/g, '').toLowerCase();
  if (!a || !b) return true;
  if (a.length < 4 || b.length < 4) return true;
  var short = a.length <= b.length ? a : b;
  var long = a.length <= b.length ? b : a;
  for (var len = Math.min(8, short.length); len >= 4; len--) {
    for (var i = 0; i + len <= short.length; i++) {
      if (long.indexOf(short.substring(i, i + len)) >= 0) return true;
    }
  }
  return false;
}

function amazonP4bNormalizePtForShelf_(pt) {
  var n = String(pt || '').trim().toUpperCase();
  if (!n) return '';
  if (AMAZON_P4B_PT_SHELF_ALIASES_[n]) return AMAZON_P4B_PT_SHELF_ALIASES_[n];
  return n;
}

/**
 * Catalog／Definitions 候補から棚にある PT を優先。無ければ先頭をエイリアス正規化。
 */
function amazonP4bFirstProductType_(candidates) {
  var norms = [];
  var i;
  for (i = 0; i < (candidates || []).length; i++) {
    var n = amazonP4bNormalizePtForShelf_(candidates[i]);
    if (n) norms.push(n);
  }
  if (!norms.length) return '';
  var shelf = null;
  try {
    if (typeof batchExportAmazonLoadShelfRegistry_ === 'function') {
      shelf = batchExportAmazonLoadShelfRegistry_();
    }
  } catch (eShelf) {
    shelf = null;
  }
  if (shelf && shelf.ok && shelf.byPt) {
    for (i = 0; i < norms.length; i++) {
      if (shelf.byPt[norms[i]]) return norms[i];
    }
  }
  return norms[0];
}

function amazonP4bSuggestOneParent_(masterCtx, creds, accessToken, parent, colPt, colBrowse, runId) {
  var row = masterCtx.values[parent.parentRowIndex0] || [];
  var existingPt = String(row[colPt] != null ? row[colPt] : '').trim();
  var existingBrowse = String(row[colBrowse] != null ? row[colBrowse] : '').trim();
  var existingPtNorm = amazonP4bNormalizePtForShelf_(existingPt);
  if (existingPt && existingPtNorm && existingPtNorm !== existingPt.toUpperCase()) {
    var sheetAlias = masterCtx.sheet;
    var rAlias = parent.parentRowIndex0 + 1;
    sheetAlias.getRange(rAlias, colPt + 1).setValue(existingPtNorm);
    masterCtx.values[parent.parentRowIndex0][colPt] = existingPtNorm;
    Logger.log('[AmazonP4b] runId=' + runId + ' state=DONE action=ALIAS_PT row=' + rAlias +
      ' parent=' + parent.parentSku + ' from=' + existingPt + ' to=' + existingPtNorm);
    existingPt = existingPtNorm;
    if (existingBrowse) {
      return {
        parentSku: parent.parentSku,
        wrote: true,
        productType: existingPt,
        browse: existingBrowse,
        sources: ['shelf_alias'],
        skipReason: '',
        aliasRemap: true
      };
    }
  }
  if (existingPt && existingBrowse) {
    return {
      parentSku: parent.parentSku,
      wrote: false,
      skipReason: 'both_cells_nonempty',
      sources: []
    };
  }

  var jan = amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, 'JANコード') ||
    amazonApprovalLv4Cell_(masterCtx, parent.sampleChildRowIndex0, 'JANコード');
  var warns = [];
  var sources = [];
  var decided = null;
  var signals = amazonP4bCollectTextSignals_(masterCtx, parent);
  var isMix = amazonP4bParentIsMix_(masterCtx, parent);
  var excludeOwn = isMix ? {} : amazonP4bOwnAsinSet_(masterCtx, parent);
  var ss = (masterCtx.sheet && typeof masterCtx.sheet.getParent === 'function')
    ? masterCtx.sheet.getParent()
    : SpreadsheetApp.getActiveSpreadsheet();
  var maxAsins = amazonP4bMaxAsins_();
  var lastVote = { voteMap: {}, voteMeta: {}, refAsin: '' };

  function tryStage(cands, stageName) {
    if (!cands || !cands.length) {
      Logger.log('[AmazonP4b] runId=' + runId + ' stage=' + stageName + ' n=0');
      return null;
    }
    var built = amazonP4bVoteAsins_(creds, accessToken, parent, jan, cands, runId, warns, sources);
    if (built.refAsin) lastVote.refAsin = built.refAsin;
    if (!built.hasAccept) {
      Logger.log('[AmazonP4b] runId=' + runId + ' stage=' + stageName +
        ' n=' + cands.length + ' accept=0');
      return null;
    }
    lastVote.voteMap = built.voteMap;
    lastVote.voteMeta = built.voteMeta;
    var picked = amazonP4bPickVoteWinnerWithTieBreak_(built.voteMap, built.voteMeta, signals);
    Logger.log('[AmazonP4b] runId=' + runId + ' stage=' + stageName +
      ' accept=' + built.acceptCount + ' picked=' + ((picked && picked.nodeId) || 'tie'));
    return picked;
  }

  if (amazonP4bPasteCircleVoteEnabled_()) {
    var pasteBlock = amazonP4bFindPasteBlockIndexForParent_(ss, jan, parent.itemName, isMix);
    Logger.log('[AmazonP4b] runId=' + runId + ' parent=' + parent.parentSku +
      ' mix=' + isMix + ' pasteBlock=' + pasteBlock);
    decided = tryStage(
      amazonP4bCollectPasteCircleAsins_(ss, pasteBlock, maxAsins, excludeOwn),
      'paste_circle');
    if (!amazonP4bHasNode_(decided)) {
      decided = tryStage(
        amazonP4bCollectKeepaLogCircleAsins_(ss, jan, parent.itemName, isMix, pasteBlock, maxAsins, excludeOwn),
        'keepa_log_circle');
    }
    if (!amazonP4bHasNode_(decided)) {
      decided = tryStage(
        amazonP4bCollectMasterRivalAsins_(masterCtx, parent, maxAsins, excludeOwn, isMix),
        'master_rival');
    }
    if (!amazonP4bHasNode_(decided)) {
      decided = tryStage(
        amazonP4bCollectJanCatalogAsins_(creds, accessToken, jan, maxAsins, excludeOwn),
        'jan_catalog_multi');
    }
  } else {
    decided = tryStage(amazonP4bCollectCandidateAsins_(masterCtx, parent), 'master_legacy');
    if (!amazonP4bHasNode_(decided) && amazonP4bNormalizeDigits_(jan).length >= 8) {
      var janCat = amazonP4bFetchCatalogByJan_(creds, accessToken, jan);
      if (janCat && janCat.refAsin) {
        decided = tryStage([{ asin: janCat.refAsin, source: 'jan_catalog', cat: janCat }], 'jan_catalog');
      } else {
        warns.push('jan_catalog_miss');
      }
    }
  }

  if (decided && decided.needsShelf) {
    var tiedResolved = amazonP4bResolveTieViaShelf_(signals, decided.tiedKeys, lastVote.voteMeta);
    decided = tiedResolved || decided.fallback || null;
    if (decided) sources.push('tie_shelf');
  }

  if (amazonP4bHasNode_(decided)) {
    sources.push(decided.source || 'asin_majority');
    return amazonP4bWritePtBrowseDecision_(masterCtx, parent, colPt, colBrowse, runId,
      existingPt, existingBrowse, decided, sources, warns, jan);
  }

  var wWinner = amazonP4bPickShelfAlways_(signals);
  if (wWinner) {
    sources.push('shelf_weighted');
    decided = {
      nodeId: wWinner.nodeId,
      browse: wWinner.browse,
      asin: '',
      source: 'shelf_weighted',
      productTypes: []
    };
    return amazonP4bWritePtBrowseDecision_(masterCtx, parent, colPt, colBrowse, runId,
      existingPt, existingBrowse, decided, sources, warns, jan);
  }

  decided = amazonP4bPickVoteWinner_(lastVote.voteMap, lastVote.voteMeta);
  if (amazonP4bHasNode_(decided)) {
    sources.push('asin_majority_last');
    return amazonP4bWritePtBrowseDecision_(masterCtx, parent, colPt, colBrowse, runId,
      existingPt, existingBrowse, decided, sources, warns, jan);
  }

  if (warns.length) {
    Logger.log('[AmazonP4b] runId=' + runId + ' WARN parent=' + parent.parentSku +
      ' warns=' + warns.join(';'));
  }
  Logger.log('[AmazonP4b] runId=' + runId + ' state=DONE action=NO_WRITE parent=' +
    parent.parentSku + ' reason=no_signal_at_all');
  return {
    parentSku: parent.parentSku,
    wrote: false,
    skipReason: 'no_signal_at_all',
    sources: amazonP4bUniqSources_(sources),
    productType: existingPt || '',
    browse: existingBrowse || '',
    browseNodeId: '',
    refAsin: lastVote.refAsin || '',
    jan: jan || '',
    warns: warns
  };
}

function amazonP4bUniqSources_(sources) {
  var seen = {};
  var out = [];
  var i;
  for (i = 0; i < (sources || []).length; i++) {
    var s = sources[i];
    if (!s || seen[s]) continue;
    seen[s] = true;
    out.push(s);
  }
  return out;
}

function amazonP4bPasteCircleVoteEnabled_() {
  return getBoolScriptProperty_(APPROVAL_AMAZON_P4B_PASTE_CIRCLE_VOTE_PROP, true);
}

function amazonP4bMaxAsins_() {
  var def = amazonP4bPasteCircleVoteEnabled_() ? 10 : 5;
  return Math.max(1, Math.min(10,
    Math.floor(getNumberScriptProperty_(APPROVAL_AMAZON_P4B_PT_MAX_ASINS_PROP, def))));
}

function amazonP4bHasNode_(decided) {
  return !!(decided && (decided.nodeId || decided.needsShelf));
}

function amazonP4bParentIsMix_(masterCtx, parent) {
  var v = amazonApprovalLv4Cell_(masterCtx, parent.parentRowIndex0, 'セット構成2_親SKU');
  return !!String(v || '').trim();
}

function amazonP4bOwnAsinSet_(masterCtx, parent) {
  var seen = {};
  var rows = [parent.parentRowIndex0].concat(parent.childRowIndexes0 || []);
  var i;
  for (i = 0; i < rows.length; i++) {
    var own = amazonP4bFirstAsinCell_(masterCtx, rows[i], 'ASINコード');
    if (own) seen[own] = true;
  }
  return seen;
}

function amazonP4bLooksLikeMixTitle_(s) {
  var t = String(s || '');
  return /・|セット|詰合|詰め合|アソート|2種|3種|食べ比べ/.test(t);
}

function amazonP4bCommonLen_(a, b) {
  var sa = String(a || '');
  var sb = String(b || '');
  if (!sa || !sb) return 0;
  var shortS = sa.length <= sb.length ? sa : sb;
  var longS = sa.length <= sb.length ? sb : sa;
  var ln;
  var i;
  for (ln = Math.min(12, shortS.length); ln >= 3; ln--) {
    for (i = 0; i + ln <= shortS.length; i++) {
      if (longS.indexOf(shortS.substr(i, ln)) >= 0) return ln;
    }
  }
  return 0;
}

function amazonP4bFindPasteBlockIndexForParent_(ss, jan, itemName, isMix) {
  if (!ss) return -1;
  var sheetName = (typeof ASIN_PASTE_SHEET_NAME !== 'undefined')
    ? ASIN_PASTE_SHEET_NAME : 'ASIN貼り付け（Keepa用）';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return -1;
  var data = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  var hr;
  var hc;
  for (hr = 0; hr < Math.min(data.length, 5); hr++) {
    var rowData = data[hr] || [];
    for (hc = 0; hc < rowData.length; hc++) {
      if (String(rowData[hc]).trim() === 'ASIN') {
        headerRowIdx = hr;
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) return -1;
  var headers = data[headerRowIdx] || [];
  var asinColumns = [];
  var c;
  for (c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'ASIN') asinColumns.push(c);
  }
  var janDigits = amazonP4bNormalizeDigits_(jan);
  var nameRow = data[0] || [];
  var bestIdx = -1;
  var bestScore = -1;
  var bi;
  for (bi = 0; bi < asinColumns.length; bi++) {
    var colAsin = asinColumns[bi];
    var blockName = String(nameRow[colAsin + 1] || '').trim();
    var blockJan = amazonP4bNormalizeDigits_(nameRow[colAsin + 2]);
    if (janDigits && blockJan && blockJan !== janDigits) continue;
    var mixTitle = amazonP4bLooksLikeMixTitle_(blockName);
    if (isMix && !mixTitle) continue;
    if (!isMix && mixTitle) continue;
    var score = 0;
    if (janDigits && blockJan && blockJan === janDigits) score += 8;
    if (amazonP4bTitleLooksRelated_(itemName, blockName)) score += 10;
    var na = amazonP4bNormalizeHaystack_(itemName);
    var nb = amazonP4bNormalizeHaystack_(blockName);
    if (na && nb && (na.indexOf(nb) >= 0 || nb.indexOf(na) >= 0)) score += 5;
    score += Math.min(8, amazonP4bCommonLen_(na, nb));
    if (score > bestScore) {
      bestScore = score;
      bestIdx = bi;
    }
  }
  return bestIdx;
}

function amazonP4bCollectPasteCircleAsins_(ss, blockIndex, maxAsins, excludeOwn) {
  var out = [];
  if (!ss || blockIndex < 0) return out;
  var sheetName = (typeof ASIN_PASTE_SHEET_NAME !== 'undefined')
    ? ASIN_PASTE_SHEET_NAME : 'ASIN貼り付け（Keepa用）';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return out;
  var data = sheet.getDataRange().getValues();
  var headerRowIdx = -1;
  var hr;
  var hc;
  for (hr = 0; hr < Math.min(data.length, 5); hr++) {
    var rowData = data[hr] || [];
    for (hc = 0; hc < rowData.length; hc++) {
      if (String(rowData[hc]).trim() === 'ASIN') {
        headerRowIdx = hr;
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) return out;
  var headers = data[headerRowIdx] || [];
  var asinColumns = [];
  var c;
  for (c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'ASIN') asinColumns.push(c);
  }
  if (blockIndex >= asinColumns.length) return out;
  var colAsin = asinColumns[blockIndex];
  var colEval = colAsin + 2;
  var colReason = colAsin + 6;
  var machineOk = [];
  var allCircle = [];
  var r;
  for (r = headerRowIdx + 1; r < data.length; r++) {
    if (String(data[r][colEval] || '').trim() !== '◎') continue;
    var asin = String(data[r][colAsin] || '').trim().toUpperCase();
    if (!amazonP4bLooksLikeAsin_(asin)) continue;
    if (excludeOwn && excludeOwn[asin]) continue;
    var rec = { asin: asin, source: 'paste_circle' };
    allCircle.push(rec);
    var reason = String(data[r][colReason] || '');
    if (reason.indexOf('非候補') < 0) machineOk.push(rec);
  }
  var use = machineOk.length ? machineOk : allCircle;
  return amazonP4bCapAsins_(use, maxAsins);
}

function amazonP4bCollectKeepaLogCircleAsins_(ss, jan, itemName, isMix, pasteBlock, maxAsins, excludeOwn) {
  var j = String(jan || '').trim();
  if (j.length < 8 || !ss) return [];
  var sheetName = (typeof KEEPA_FETCH_LOG_SHEET_NAME !== 'undefined')
    ? KEEPA_FETCH_LOG_SHEET_NAME : 'Keepa取得_ログ';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0] || [];
  var colJan = -1;
  var colAsin = -1;
  var colEval = -1;
  var colDate = -1;
  var colExpect = -1;
  var colBlock = -1;
  var c;
  for (c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (h === 'JAN') colJan = c;
    else if (h === 'ASIN') colAsin = c;
    else if (h === '評価') colEval = c;
    else if (h === '日時') colDate = c;
    else if (h.indexOf('期待名') === 0) colExpect = c;
    else if (h === 'ブロック') colBlock = c;
  }
  if (colJan < 0 || colAsin < 0 || colEval < 0) return [];
  var wantBlock = (pasteBlock >= 0 && colBlock >= 0) ? String(pasteBlock + 1) : '';
  var latestDate = '';
  var r;
  for (r = data.length - 1; r >= 1; r--) {
    if (String(data[r][colJan] || '').trim() !== j) continue;
    latestDate = colDate >= 0 ? String(data[r][colDate] || '').trim() : '';
    break;
  }
  if (colDate >= 0 && !latestDate) return [];
  var out = [];
  var seen = {};
  for (r = 1; r < data.length; r++) {
    if (String(data[r][colJan] || '').trim() !== j) continue;
    if (colDate >= 0 && String(data[r][colDate] || '').trim() !== latestDate) continue;
    if (String(data[r][colEval] || '').trim() !== '◎') continue;
    if (wantBlock && String(data[r][colBlock] || '').trim() !== wantBlock) continue;
    var expectName = colExpect >= 0 ? String(data[r][colExpect] || '') : '';
    if (!wantBlock) {
      var mixTitle = amazonP4bLooksLikeMixTitle_(expectName);
      if (isMix && !mixTitle) continue;
      if (!isMix && mixTitle) continue;
      if (expectName && !amazonP4bTitleLooksRelated_(itemName, expectName)) continue;
    }
    var asin = String(data[r][colAsin] || '').trim().toUpperCase();
    if (!amazonP4bLooksLikeAsin_(asin) || seen[asin]) continue;
    if (excludeOwn && excludeOwn[asin]) continue;
    seen[asin] = true;
    out.push({ asin: asin, source: 'keepa_log_circle' });
  }
  return amazonP4bCapAsins_(out, maxAsins);
}

function amazonP4bCollectMasterRivalAsins_(masterCtx, parent, maxAsins, excludeOwn, isMix) {
  var rows = [parent.parentRowIndex0].concat(parent.childRowIndexes0 || []);
  var seenRow = {};
  var uniqRows = [];
  var r;
  for (r = 0; r < rows.length; r++) {
    if (seenRow[rows[r]]) continue;
    seenRow[rows[r]] = true;
    uniqRows.push(rows[r]);
  }
  var out = [];
  var seenAsin = {};
  function add(asin, source) {
    asin = String(asin || '').trim().toUpperCase();
    if (!amazonP4bLooksLikeAsin_(asin) || seenAsin[asin]) return;
    if (excludeOwn && excludeOwn[asin]) return;
    if (out.length >= maxAsins) return;
    seenAsin[asin] = true;
    out.push({ asin: asin, source: source });
  }
  var i;
  for (i = 0; i < uniqRows.length; i++) {
    var rawRival = amazonApprovalLv4Cell_(masterCtx, uniqRows[i], AMAZON_P4B_RIVAL_ASIN_HEADER_);
    if (amazonP4bLooksLikeAsin_(rawRival)) add(rawRival, 'rival_asin');
  }
  for (i = 0; i < uniqRows.length; i++) {
    var urlInfo = amazonP4bFirstUrlOnRow_(masterCtx, uniqRows[i]);
    if (!urlInfo.url) continue;
    var ua = amazonP4bAsinFromUrl_(urlInfo.url);
    if (ua) add(ua, 'rival_url');
  }
  if (isMix) {
    for (i = 0; i < uniqRows.length; i++) {
      var own = amazonP4bFirstAsinCell_(masterCtx, uniqRows[i], 'ASINコード');
      if (own) add(own, 'own_asin');
    }
  }
  return out;
}

function amazonP4bCapAsins_(list, maxAsins) {
  var cap = Math.max(1, Number(maxAsins) || 10);
  if (list.length <= cap) return list;
  return list.slice(0, cap);
}

function amazonP4bCollectJanCatalogAsins_(creds, accessToken, jan, maxAsins, excludeOwn) {
  var digits = amazonP4bNormalizeDigits_(jan);
  if (digits.length < 8) return [];
  var res = amazonSpapiPutHttpGet_(
    creds,
    accessToken,
    '/catalog/2022-04-01/items',
    {
      marketplaceIds: creds.marketplaceId,
      identifiers: digits,
      identifiersType: 'JAN',
      includedData: 'summaries,productTypes,classifications,identifiers'
    }
  );
  if (res.code !== 200 || !res.json) {
    Logger.log('[AmazonP4b] jan search HTTP ' + res.code + ' jan=' + digits);
    return [];
  }
  var items = res.json.items || [];
  var out = [];
  var seen = {};
  var i;
  for (i = 0; i < items.length; i++) {
    var wrapped = amazonP4bWrapCatalogItem_(items[i]);
    var asin = String(wrapped.refAsin || '').toUpperCase();
    if (!amazonP4bLooksLikeAsin_(asin) || seen[asin]) continue;
    if (excludeOwn && excludeOwn[asin]) continue;
    seen[asin] = true;
    out.push({ asin: asin, source: 'jan_catalog', cat: wrapped });
    if (out.length >= maxAsins) break;
  }
  return out;
}

function amazonP4bWrapCatalogItem_(item) {
  var wrapped = {
    productTypes: [],
    browse: amazonP4bBrowseFromClassifications_(item && item.classifications),
    identifierDigits: [],
    itemName: '',
    brand: '',
    manufacturer: '',
    refAsin: String((item && (item.asin || item.ASIN)) || '').trim().toUpperCase()
  };
  if (!item) return wrapped;
  var pts = item.productTypes || [];
  var i;
  for (i = 0; i < pts.length; i++) {
    var one = pts[i];
    var name = '';
    if (typeof one === 'string') name = one;
    else if (one && typeof one === 'object') name = one.productType || one.name || '';
    if (name) wrapped.productTypes.push(String(name));
  }
  var summaries = item.summaries || [];
  if (summaries.length && summaries[0]) {
    wrapped.itemName = String(summaries[0].itemName || summaries[0].name || '').trim();
    wrapped.brand = String(summaries[0].brand || summaries[0].brandName || '').trim();
    wrapped.manufacturer = String(summaries[0].manufacturer || '').trim();
  }
  var idBlocks = item.identifiers || [];
  var idSeen = {};
  var b;
  for (b = 0; b < idBlocks.length; b++) {
    var block = idBlocks[b];
    var ids = (block && block.identifiers) ? block.identifiers : [];
    var k;
    for (k = 0; k < ids.length; k++) {
      var idObj = ids[k];
      if (!idObj) continue;
      var dig = amazonP4bNormalizeDigits_(idObj.identifier || idObj.value || '');
      if (dig.length >= 8 && !idSeen[dig]) {
        idSeen[dig] = true;
        wrapped.identifierDigits.push(dig);
      }
    }
  }
  return wrapped;
}

function amazonP4bVoteAsins_(creds, accessToken, parent, jan, cands, runId, warns, sources) {
  var voteMap = {};
  var voteMeta = {};
  var hasAccept = false;
  var acceptCount = 0;
  var refAsin = '';
  var ai;
  for (ai = 0; ai < (cands || []).length; ai++) {
    var cand = cands[ai];
    sources.push(cand.source);
    if (!refAsin) refAsin = cand.asin;
    var cat = cand.cat || amazonP4bFetchCatalogHints_(creds, accessToken, cand.asin);
    var chk = amazonP4bCatalogConsistencyCheck_(parent.itemName, jan, cat, cand.asin);
    if (!chk.ok) {
      warns.push.apply(warns, chk.reasons);
      Logger.log('[AmazonP4b] runId=' + runId + ' REJECT asin=' + cand.asin +
        ' source=' + cand.source + ' reasons=' + chk.reasons.join(';'));
      continue;
    }
    var browseHint = cat.browse || '';
    var nodeId = amazonP4bExtractBrowseNodeId_(browseHint);
    if (!nodeId) {
      warns.push('catalog_no_browse_node:' + cand.asin);
      continue;
    }
    if (!voteMap[nodeId]) voteMap[nodeId] = 0;
    voteMap[nodeId]++;
    if (!voteMeta[nodeId] || voteMap[nodeId] >= (voteMeta[nodeId].votes || 0)) {
      voteMeta[nodeId] = {
        votes: voteMap[nodeId],
        browse: browseHint,
        asin: cand.asin,
        source: cand.source,
        productTypes: cat.productTypes || [],
        itemName: cat.itemName || ''
      };
    }
    hasAccept = true;
    acceptCount++;
    Logger.log('[AmazonP4b] runId=' + runId + ' ACCEPT asin=' + cand.asin +
      ' node=' + nodeId + ' source=' + cand.source);
  }
  return {
    voteMap: voteMap,
    voteMeta: voteMeta,
    hasAccept: hasAccept,
    acceptCount: acceptCount,
    refAsin: refAsin
  };
}

function amazonP4bVoteMetaToDecided_(nodeId, voteMap, voteMeta) {
  var meta = (voteMeta && voteMeta[nodeId]) || {};
  return {
    nodeId: nodeId,
    browse: meta.browse || nodeId,
    asin: meta.asin || '',
    source: meta.source || 'asin_majority',
    productTypes: meta.productTypes || [],
    votes: voteMap ? voteMap[nodeId] : 0
  };
}

function amazonP4bTopVoteKeys_(voteMap) {
  var keys = Object.keys(voteMap || {});
  if (!keys.length) return [];
  var max = 0;
  var i;
  for (i = 0; i < keys.length; i++) {
    if (voteMap[keys[i]] > max) max = voteMap[keys[i]];
  }
  var top = [];
  for (i = 0; i < keys.length; i++) {
    if (voteMap[keys[i]] === max) top.push(keys[i]);
  }
  top.sort();
  return top;
}

function amazonP4bFilterKeysByHaystack_(keys, voteMeta, haystack) {
  var hs = amazonP4bNormalizeHaystack_(haystack);
  if (!hs || hs.length < 2) return [];
  var hit = [];
  var i;
  for (i = 0; i < (keys || []).length; i++) {
    var nid = keys[i];
    var browse = String((voteMeta[nid] && voteMeta[nid].browse) || '');
    var leaf = browse.replace(/\s*\(\d+\)\s*$/, '');
    var parts = leaf.split('>');
    var last = amazonP4bNormalizeHaystack_(parts[parts.length - 1] || leaf);
    if (last.length < 2) continue;
    if (hs.indexOf(last) >= 0 || last.indexOf(hs) >= 0) {
      hit.push(nid);
      continue;
    }
    var tokenHit = false;
    var chunks = last.split(/[・･\/]/);
    var c;
    for (c = 0; c < chunks.length; c++) {
      var ck = amazonP4bNormalizeHaystack_(chunks[c]);
      if (ck.length >= 2 && hs.indexOf(ck) >= 0) tokenHit = true;
    }
    if (tokenHit) hit.push(nid);
  }
  return hit;
}

function amazonP4bPickVoteWinnerWithTieBreak_(voteMap, voteMeta, signals) {
  var top = amazonP4bTopVoteKeys_(voteMap);
  if (!top.length) return null;
  if (top.length === 1) return amazonP4bVoteMetaToDecided_(top[0], voteMap, voteMeta);
  var mallHs = [signals.rakuten, signals.yahoo].join(' ');
  var mall = amazonP4bFilterKeysByHaystack_(top, voteMeta, mallHs);
  if (mall.length === 1) {
    var dMall = amazonP4bVoteMetaToDecided_(mall[0], voteMap, voteMeta);
    dMall.source = 'tie_mall';
    return dMall;
  }
  var pool = mall.length ? mall : top;
  var nameHs = [signals.itemName, signals.nameBase].join(' ');
  var byName = amazonP4bFilterKeysByHaystack_(pool, voteMeta, nameHs);
  if (byName.length === 1) {
    var dName = amazonP4bVoteMetaToDecided_(byName[0], voteMap, voteMeta);
    dName.source = 'tie_name';
    return dName;
  }
  var remain = byName.length ? byName : pool;
  return {
    needsShelf: true,
    tiedKeys: remain,
    fallback: amazonP4bVoteMetaToDecided_(remain[0], voteMap, voteMeta)
  };
}

function amazonP4bResolveTieViaShelf_(signals, tiedKeys, voteMeta) {
  var shelf = amazonP4bPickShelfAlways_(signals);
  if (!shelf || !shelf.nodeId) return null;
  var i;
  for (i = 0; i < (tiedKeys || []).length; i++) {
    if (String(tiedKeys[i]) === String(shelf.nodeId)) {
      return amazonP4bVoteMetaToDecided_(tiedKeys[i], {}, voteMeta);
    }
  }
  return {
    nodeId: shelf.nodeId,
    browse: shelf.browse,
    asin: '',
    source: 'tie_shelf',
    productTypes: []
  };
}

function amazonP4bPickShelfAlways_(signals) {
  var wVotes = {};
  var wMeta = {};
  amazonP4bAddShelfVotesFromHaystack_(signals.itemName, 3, wVotes, wMeta, 'item_name');
  amazonP4bAddShelfVotesFromHaystack_(signals.nameBase, 1, wVotes, wMeta, 'name_base');
  amazonP4bAddShelfVotesFromHaystack_(signals.maker, 1, wVotes, wMeta, 'maker');
  amazonP4bAddShelfVotesFromHaystack_(signals.rakuten, 2, wVotes, wMeta, 'rakuten_genre');
  amazonP4bAddShelfVotesFromHaystack_(signals.yahoo, 2, wVotes, wMeta, 'yahoo_category');
  var hit = amazonP4bPickWeightedWinner_(wVotes, wMeta, 1);
  if (hit) return hit;
  return amazonP4bPickWeightedWinner_(wVotes, wMeta, 0);
}

function amazonP4bCollectCandidateAsins_(masterCtx, parent) {
  var maxAsins = amazonP4bMaxAsins_();
  var rows = [parent.parentRowIndex0].concat(parent.childRowIndexes0 || []);
  var seenRow = {};
  var uniqRows = [];
  var r;
  for (r = 0; r < rows.length; r++) {
    if (seenRow[rows[r]]) continue;
    seenRow[rows[r]] = true;
    uniqRows.push(rows[r]);
  }
  var out = [];
  var seenAsin = {};
  function add(asin, source) {
    asin = String(asin || '').trim().toUpperCase();
    if (!amazonP4bLooksLikeAsin_(asin) || seenAsin[asin]) return;
    if (out.length >= maxAsins) return;
    seenAsin[asin] = true;
    out.push({ asin: asin, source: source });
  }
  var i;
  for (i = 0; i < uniqRows.length; i++) {
    var rawRival = amazonApprovalLv4Cell_(masterCtx, uniqRows[i], AMAZON_P4B_RIVAL_ASIN_HEADER_);
    if (amazonP4bLooksLikeAsin_(rawRival)) add(rawRival, 'rival_asin');
  }
  for (i = 0; i < uniqRows.length; i++) {
    var urlInfo = amazonP4bFirstUrlOnRow_(masterCtx, uniqRows[i]);
    if (!urlInfo.url) continue;
    var ua = amazonP4bAsinFromUrl_(urlInfo.url);
    if (ua) add(ua, 'rival_url');
  }
  for (i = 0; i < uniqRows.length; i++) {
    var own = amazonP4bFirstAsinCell_(masterCtx, uniqRows[i], 'ASINコード');
    if (own) add(own, 'own_asin');
  }
  return out;
}

function amazonP4bCatalogConsistencyCheck_(itemName, jan, cat, asin) {
  var reasons = [];
  if (!cat) {
    return { ok: false, reasons: ['catalog_empty:' + asin] };
  }
  var hasPt = cat.productTypes && cat.productTypes.length;
  var hasBrowse = !!String(cat.browse || '').trim();
  if (!hasPt && !hasBrowse) {
    reasons.push('catalog_http_or_empty:' + asin);
    return { ok: false, reasons: reasons };
  }
  if (!amazonP4bTitleLooksRelated_(itemName, cat.itemName)) {
    reasons.push('title_mismatch:' + asin);
  }
  var janDigits = amazonP4bNormalizeDigits_(jan);
  if (janDigits.length >= 8 && cat.identifierDigits && cat.identifierDigits.length) {
    var janHit = false;
    var j;
    for (j = 0; j < cat.identifierDigits.length; j++) {
      if (cat.identifierDigits[j] === janDigits ||
          cat.identifierDigits[j].indexOf(janDigits) >= 0 ||
          janDigits.indexOf(cat.identifierDigits[j]) >= 0) {
        janHit = true;
        break;
      }
    }
    if (!janHit) reasons.push('jan_mismatch:' + asin);
  }
  var conflict = amazonP4bBrowseNameCategoryConflict_(itemName, cat.browse || '');
  if (conflict) reasons.push(conflict);
  return { ok: reasons.length === 0, reasons: reasons };
}

function amazonP4bBrowseNameCategoryConflict_(itemName, browseText) {
  var n = String(itemName || '');
  var b = String(browseText || '');
  var nameMeat = /肉|焼鳥|焼き鳥|牛|豚|鶏|コンビーフ|ハム|ソーセージ/.test(n);
  var nameFish = /魚|鮭|マグロ|ツナ|いわし|鯖|さんま|海産/.test(n);
  var browseMeat = /肉の缶詰/.test(b);
  var browseFish = /魚介の缶詰/.test(b);
  if (nameMeat && browseFish) return 'category_conflict:name_meat_browse_fish';
  if (nameFish && browseMeat) return 'category_conflict:name_fish_browse_meat';
  return '';
}

function amazonP4bPickVoteWinner_(voteMap, voteMeta) {
  var keys = Object.keys(voteMap || {});
  if (!keys.length) return null;
  keys.sort(function (a, b) {
    var d = voteMap[b] - voteMap[a];
    if (d) return d;
    return a < b ? -1 : 1;
  });
  var top = keys[0];
  var meta = voteMeta[top] || {};
  return {
    nodeId: top,
    browse: meta.browse || top,
    asin: meta.asin || '',
    source: meta.source || 'asin_majority',
    productTypes: meta.productTypes || [],
    votes: voteMap[top]
  };
}

function amazonP4bPickWeightedWinner_(voteMap, voteMeta, threshold) {
  var keys = Object.keys(voteMap || {});
  if (!keys.length) return null;
  keys.sort(function (a, b) {
    var d = voteMap[b] - voteMap[a];
    if (d) return d;
    return a < b ? -1 : 1;
  });
  var top = keys[0];
  if (voteMap[top] < threshold) return null;
  var meta = voteMeta[top] || {};
  return {
    nodeId: top,
    browse: meta.browse || amazonP4bFormatBrowseFromShelf_(top),
    score: voteMap[top]
  };
}

function amazonP4bFormatBrowseFromShelf_(nodeId) {
  var hit = amazonP4bResolvePreferredPtFromBrowse_(nodeId, null);
  if (hit && hit.browsePath) {
    return String(hit.browsePath).trim() + ' (' + nodeId + ')';
  }
  return String(nodeId);
}

function amazonP4bCollectTextSignals_(masterCtx, parent) {
  function cell() {
    var names = Array.prototype.slice.call(arguments);
    var rows = [parent.parentRowIndex0].concat(parent.childRowIndexes0 || []);
    var ri;
    var ni;
    for (ri = 0; ri < rows.length; ri++) {
      for (ni = 0; ni < names.length; ni++) {
        var v = amazonApprovalLv4Cell_(masterCtx, rows[ri], names[ni]);
        if (v) return String(v).trim();
      }
    }
    return '';
  }
  return {
    itemName: String(parent.itemName || '').trim(),
    nameBase: cell('商品名ベース', '商品名', '商品名amazon'),
    maker: cell('メーカー名ベース', 'メーカー名'),
    rakuten: cell('楽天ジャンル名', '要確認_楽天ジャンル', '★推奨楽天ジャンルID'),
    yahoo: cell('Yahooカテゴリ名', '要確認_Yahooカテゴリ', '★推奨YahooカテゴリID')
  };
}

function amazonP4bNormalizeHaystack_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, '').replace(/　/g, '');
}

function amazonP4bAddShelfVotesFromHaystack_(haystack, weight, voteMap, voteMeta, sourceTag) {
  var hs = amazonP4bNormalizeHaystack_(haystack);
  if (!hs || hs.length < 2 || !(weight > 0)) return;
  var map = null;
  try {
    if (typeof batchExportAmazonLoadShelfRegistry_ === 'function') {
      var loaded = batchExportAmazonLoadShelfRegistry_();
      if (loaded && loaded.ok) map = loaded.byBrowseNode || null;
    }
  } catch (e0) {
    map = null;
  }
  if (!map) return;
  var nids = Object.keys(map);
  var i;
  for (i = 0; i < nids.length; i++) {
    var nid = nids[i];
    var brow = map[nid] || {};
    var path = String(brow.browsePath || '').trim();
    if (!path) continue;
    var parts = path.split('>');
    var leaf = String(parts[parts.length - 1] || '').trim();
    var leafN = amazonP4bNormalizeHaystack_(leaf);
    if (leafN.length < 4) continue;
    var hit = hs.indexOf(leafN) >= 0;
    if (!hit && /肉の缶詰/.test(leafN) && /肉|焼鳥|焼き鳥|コンビーフ|ハム|ソーセージ/.test(hs)) hit = true;
    if (!hit && /魚介の缶詰/.test(leafN) && /魚|鮭|マグロ|ツナ|いわし|鯖|さんま/.test(hs)) hit = true;
    if (!hit) {
      var chunks = leafN.split(/[・･]/);
      var ci;
      for (ci = 0; ci < chunks.length; ci++) {
        var ck = String(chunks[ci] || '');
        if (ck.length >= 3 && hs.indexOf(ck) >= 0) {
          hit = true;
          break;
        }
      }
    }
    if (!hit) continue;
    if (!voteMap[nid]) voteMap[nid] = 0;
    voteMap[nid] += weight;
    voteMeta[nid] = {
      browse: path + ' (' + nid + ')',
      source: sourceTag,
      path: path
    };
  }
}

function amazonP4bFetchCatalogByJan_(creds, accessToken, jan) {
  var digits = amazonP4bNormalizeDigits_(jan);
  if (digits.length < 8) return null;
  var res = amazonSpapiPutHttpGet_(
    creds,
    accessToken,
    '/catalog/2022-04-01/items',
    {
      marketplaceIds: creds.marketplaceId,
      identifiers: digits,
      identifiersType: 'JAN',
      includedData: 'summaries,productTypes,classifications,identifiers'
    }
  );
  if (res.code !== 200 || !res.json) {
    Logger.log('[AmazonP4b] jan search HTTP ' + res.code + ' jan=' + digits);
    return null;
  }
  var items = res.json.items || [];
  if (!items.length) return null;
  return amazonP4bWrapCatalogItem_(items[0]);
}

function amazonP4bWritePtBrowseDecision_(masterCtx, parent, colPt, colBrowse, runId,
    existingPt, existingBrowse, decided, sources, warns, jan) {
  var nodeId = String(decided.nodeId || '').trim();
  var browseHint = String(decided.browse || '').trim();
  if (!browseHint && nodeId) browseHint = amazonP4bFormatBrowseFromShelf_(nodeId);
  if (!amazonP4bExtractBrowseNodeId_(browseHint) && nodeId) {
    browseHint = browseHint ? (browseHint + ' (' + nodeId + ')') : nodeId;
  }
  if (!amazonP4bExtractBrowseNodeId_(browseHint)) {
    Logger.log('[AmazonP4b] runId=' + runId + ' NO_WRITE missing_browse parent=' + parent.parentSku);
    return {
      parentSku: parent.parentSku,
      wrote: false,
      skipReason: 'decision_missing_browse',
      sources: amazonP4bUniqSources_(sources),
      warns: warns || []
    };
  }

  var shelfHit = amazonP4bResolvePreferredPtFromBrowse_(
    amazonP4bExtractBrowseNodeId_(browseHint) || nodeId, null);
  var pickedPt = '';
  if (shelfHit && shelfHit.preferredProductType) {
    pickedPt = String(shelfHit.preferredProductType).trim().toUpperCase();
    sources.push('shelf_browse');
    Logger.log('[AmazonP4b] runId=' + runId + ' browse_route node=' +
      (amazonP4bExtractBrowseNodeId_(browseHint) || nodeId) +
      ' preferred=' + pickedPt + ' template=' + String(shelfHit.templateFile || ''));
  } else if (decided.productTypes && decided.productTypes.length) {
    pickedPt = amazonP4bNormalizePtForShelf_(amazonP4bFirstProductType_(decided.productTypes));
  }

  var writePt = '';
  if (pickedPt) {
    if (!existingPt || String(existingPt).toUpperCase() !== pickedPt) {
      if (!existingPt || shelfHit || AMAZON_P4B_PT_SHELF_ALIASES_[String(existingPt).toUpperCase()]) {
        writePt = pickedPt;
      }
    }
  }
  var writeBrowse = (!existingBrowse && browseHint) ? browseHint : '';
  if (!writeBrowse && !existingBrowse) {
    return {
      parentSku: parent.parentSku,
      wrote: false,
      skipReason: 'browse_required_not_writable',
      sources: amazonP4bUniqSources_(sources),
      warns: warns || []
    };
  }
  if (!writePt && !writeBrowse) {
    return {
      parentSku: parent.parentSku,
      wrote: false,
      skipReason: existingPt ? 'pt_filled_browse_empty_no_hint' : 'nothing_to_write',
      sources: amazonP4bUniqSources_(sources),
      productType: pickedPt || existingPt || '',
      browse: browseHint || existingBrowse || '',
      warns: warns || []
    };
  }

  var sheet = masterCtx.sheet;
  var r1 = parent.parentRowIndex0 + 1;
  if (writePt) {
    sheet.getRange(r1, colPt + 1).setValue(writePt);
    masterCtx.values[parent.parentRowIndex0][colPt] = writePt;
  }
  if (writeBrowse) {
    sheet.getRange(r1, colBrowse + 1).setValue(writeBrowse);
    masterCtx.values[parent.parentRowIndex0][colBrowse] = writeBrowse;
  }
  var srcUniq = amazonP4bUniqSources_(sources);
  Logger.log('[AmazonP4b] runId=' + runId + ' state=DONE action=WRITE_PT_BROWSE row=' + r1 +
    ' parent=' + parent.parentSku + ' pt=' + writePt + ' browseLen=' + writeBrowse.length +
    ' node=' + (amazonP4bExtractBrowseNodeId_(writeBrowse || browseHint) || nodeId) +
    ' sources=' + srcUniq.join('+') + ' refAsin=' + (decided.asin || ''));
  if (warns && warns.length) {
    Logger.log('[AmazonP4b] runId=' + runId + ' WARN parent=' + parent.parentSku +
      ' warns=' + warns.join(';'));
  }
  return {
    parentSku: parent.parentSku,
    wrote: true,
    productType: writePt || existingPt,
    browse: writeBrowse || existingBrowse,
    browseNodeId: amazonP4bExtractBrowseNodeId_(writeBrowse || browseHint) || nodeId,
    sources: srcUniq,
    refAsin: decided.asin || '',
    jan: jan || '',
    warns: warns || []
  };
}

/**
 * @return {{productTypes:string[], browse:string, identifierDigits:string[], itemName:string}}
 */
function amazonP4bFetchCatalogHints_(creds, accessToken, asin) {
  var out = { productTypes: [], browse: '', identifierDigits: [], itemName: '', brand: '', manufacturer: '' };
  var res = amazonSpapiPutHttpGet_(
    creds,
    accessToken,
    '/catalog/2022-04-01/items/' + encodeURIComponent(asin),
    {
      marketplaceIds: creds.marketplaceId,
      includedData: 'summaries,productTypes,classifications,identifiers'
    }
  );
  if (res.code !== 200 || !res.json) {
    Logger.log('[AmazonP4b] catalog HTTP ' + res.code + ' asin=' + asin);
    return out;
  }
  var body = res.json;
  var pts = body.productTypes || [];
  for (var i = 0; i < pts.length; i++) {
    var one = pts[i];
    var name = '';
    if (typeof one === 'string') name = one;
    else if (one && typeof one === 'object') {
      name = one.productType || one.name || '';
    }
    if (name) out.productTypes.push(String(name));
  }
  out.browse = amazonP4bBrowseFromClassifications_(body.classifications);

  var summaries = body.summaries || [];
  if (summaries.length && summaries[0]) {
    out.itemName = String(summaries[0].itemName || summaries[0].name || '').trim();
    out.brand = String(
      summaries[0].brand || summaries[0].brandName || ''
    ).trim();
    out.manufacturer = String(summaries[0].manufacturer || '').trim();
  }

  var idBlocks = body.identifiers || [];
  var idSeen = {};
  for (var b = 0; b < idBlocks.length; b++) {
    var block = idBlocks[b];
    var ids = (block && block.identifiers) ? block.identifiers : [];
    for (var k = 0; k < ids.length; k++) {
      var idObj = ids[k];
      if (!idObj) continue;
      var dig = amazonP4bNormalizeDigits_(idObj.identifier || idObj.value || '');
      if (dig.length >= 8 && !idSeen[dig]) {
        idSeen[dig] = true;
        out.identifierDigits.push(dig);
      }
    }
  }
  return out;
}

function amazonP4bBrowseFromClassifications_(classifications) {
  if (!classifications || !classifications.length) return '';
  var bestPath = '';
  var bestId = '';
  function walk(nodes, pathNames) {
    if (!nodes || !nodes.length) return;
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      if (!n || typeof n !== 'object') continue;
      var display = String(n.displayName || n.name || '').trim();
      var cid = String(n.classificationId || n.id || '').trim();
      var nextPath = pathNames.slice();
      if (display) nextPath.push(display);
      var kids = n.classifications || n.children || n.parent || null;
      var nested = n.classifications;
      if (nested && nested.length) {
        walk(nested, nextPath);
      } else if (display || cid) {
        var path = nextPath.join(' > ');
        if (cid) path = path + ' (' + cid + ')';
        if (path.length >= bestPath.length) {
          bestPath = path;
          bestId = cid;
        }
      }
      if (kids && kids !== nested && Array.isArray(kids)) walk(kids, nextPath);
    }
  }
  for (var c = 0; c < classifications.length; c++) {
    var block = classifications[c];
    if (block && block.classifications) walk(block.classifications, []);
    else if (Array.isArray(block)) walk(block, []);
    else if (block && (block.displayName || block.classificationId)) walk([block], []);
  }
  return bestPath || bestId;
}

/**
 * @return {string[]}
 */
function amazonP4bSearchDefinitions_(creds, accessToken, keywords) {
  var res = amazonSpapiPutHttpGet_(
    creds,
    accessToken,
    '/definitions/2020-09-01/productTypes',
    {
      marketplaceIds: creds.marketplaceId,
      keywords: String(keywords || '').substring(0, 80)
    }
  );
  if (res.code !== 200 || !res.json) {
    Logger.log('[AmazonP4b] definitions search HTTP ' + res.code);
    return [];
  }
  var names = [];
  var arr = res.json.productTypes || res.json.ProductTypes || [];
  for (var i = 0; i < arr.length; i++) {
    var one = arr[i];
    var name = '';
    if (typeof one === 'string') name = one;
    else if (one && typeof one === 'object') {
      name = one.name || one.productType || '';
    }
    if (name) names.push(String(name));
  }
  return names;
}
