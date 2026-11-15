/**
 * 競合DB — Catalog列挙（改訂P2 増員並走）
 * 正本: docs/org/B_COMPETITOR_JITTA_REQUIREMENTS.md §P2
 * Keepaはやらない。①は沈黙しない。同一親は共有シートロック。
 * 切戻し: CJ_CATALOG_ENABLED=false（未設定=OFF。SP-API鍵を置いてから ON）
 * 第1スライス: 本クエリで規模確定（501超は特大見積で打ち切り。分割語フルは①洗い側）。
 * 2026-11-01: 洗いキューは1実行1回読み＋追記はメモリ継ぎ足し（① SNAPSHOT と同思想）。
 * 失敗時は comp_enum を残さず catalog_skip を実態AIログへ。バッチ先頭で stuck heal。
 * 2026-11-01: 落ちあと拾い強化 — 起動先頭で次予約／実態AI・ヒットから再武装／updated古ロック掃除。
 * 2026-11-10: 未済行が上限以上なら新規社への追記を止める（① PR_WASH_QUEUE_PENDING_ROWS_MAX と同値）。
 * 2026-11-10: 追記0のまま seen 件数で規模帯を書かない（幽霊規模帯）。シートへ catalog_pending_cap。既存幽霊は heal。
 * 2026-11-11: 未済≥上限なら社選定・API前にバッチ全体スキップ（途中列挙＝キューあり社だけ続行）。時計は間隔維持。
 */
var CJ_CATALOG_ENABLED_PROP_ = 'CJ_CATALOG_ENABLED';
/** ①と同名。未設定=50000。0でオフ。 */
var CJ_WASH_QUEUE_PENDING_ROWS_MAX_PROP_ = 'PR_WASH_QUEUE_PENDING_ROWS_MAX';
var CJ_WASH_QUEUE_PENDING_ROWS_MAX_DEFAULT_ = 50000;
var CJ_CATALOG_CAP_PROP_ = 'CJ_CATALOG_CAP';
var CJ_CATALOG_CAP_DEFAULT_ = 3;
var CJ_CATALOG_CAP_MAX_ = 8;
var CJ_CATALOG_INTERVAL_MIN_PROP_ = 'CJ_CATALOG_INTERVAL_MIN';
var CJ_CATALOG_INTERVAL_MIN_DEFAULT_ = 10;
var CJ_CATALOG_TRIGGER_HANDLER_ = 'cjRunCatalogFromTrigger';
var CJ_CATALOG_EXPECT_FIRE_MS_PROP_ = 'CJ_CATALOG_EXPECT_FIRE_MS';
var CJ_CATALOG_INFLIGHT_UNTIL_PROP_ = 'CJ_CATALOG_INFLIGHT_UNTIL';
var CJ_CATALOG_STALE_GRACE_MIN_PROP_ = 'CJ_CATALOG_STALE_GRACE_MIN';
var CJ_CATALOG_STALE_GRACE_MIN_DEFAULT_ = 10;
var CJ_CATALOG_REARM_MIN_PROP_ = 'CJ_CATALOG_REARM_MIN';
var CJ_CATALOG_REARM_MIN_DEFAULT_ = 5;
var CJ_CATALOG_LOCK_STALE_MIN_PROP_ = 'CJ_CATALOG_LOCK_STALE_MIN';
var CJ_CATALOG_LOCK_STALE_MIN_DEFAULT_ = 35;

var CJ_SHARED_LOCK_SHEET_ = 'Catalog親ロック';
var CJ_SHARED_LOCK_TTL_MS_ = 30 * 60 * 1000;
var CJ_SHARED_LOCK_OWNER_ = 'comp';

var CJ_WASH_QUEUE_SHEET_ = '①洗いキュー';
var CJ_WASH_QUEUE_HEADERS_ = ['メーカー', '入口語', 'ASIN', '列挙日時', 'Keepa済', '門結果', 'メモ'];
var CJ_PROF_SHEET_ = '①プロファイル';
var CJ_PROF_SURVEY_CK_ = '調査カテゴリー';
var CJ_PROF_ROOT_ID_ = 'AmazonルートID';
var CJ_FOOD_ROOT_ = '57239051';

var CJ_CATALOG_SPLIT_THRESHOLD_ = 500;
var CJ_BAND_S_MAX_ = 100;
var CJ_BAND_M_MAX_ = 300;
var CJ_BAND_S_ = '小';
var CJ_BAND_M_ = '中';
var CJ_BAND_L_ = '大';
var CJ_BAND_XL_ = '特大';
var CJ_BAND_UNK_ = '不明';
var CJ_CATALOG_EST_COL_ = 'Catalog見積';
var CJ_BAND_COL_ = '規模帯';
var CJ_WASH_STATUS_HEADER_ = '洗い状態';
var CJ_WASH_PROGRESS_HEADER_ = '洗い進捗';
var CJ_CK_HEADER_ = '調査CK';
var CJ_WASH_ST_ENUM_ = '列挙中';

var CJ_SPAPI_LWA_URL_ = 'https://api.amazon.com/auth/o2/token';
var CJ_SPAPI_DEFAULT_ENDPOINT_ = 'https://sellingpartnerapi-fe.amazon.com';
var CJ_SPAPI_DEFAULT_MP_ = 'A1VC38T7YXB528';
var CJ_GAS_BUDGET_MS_ = 20 * 60 * 1000;
var CJ_CATALOG_GAS_START_MS_ = 0;
var CJ_SURVEY_ROOT_SETS_CACHE_ = null;
/**
 * ① PR_WASH_QUEUE_COUNT_MAP / SNAPSHOT と同じ思想。
 * 1実行で洗いキューを1回だけ読み、メーカー→ASIN集合。追記はメモリ継ぎ足し（全行再読しない）。
 * null=未読込。
 */
var CJ_WASH_QUEUE_ASIN_MAP_ = null;
/** 未済行数（AsinMap 読込時に同時計測。null=未計測）。 */
var CJ_WASH_QUEUE_PENDING_ROWS_ = null;

function cjWashQueuePendingRowsMax_() {
  var raw = PropertiesService.getScriptProperties().getProperty(CJ_WASH_QUEUE_PENDING_ROWS_MAX_PROP_);
  if (raw == null || String(raw).trim() === '') return CJ_WASH_QUEUE_PENDING_ROWS_MAX_DEFAULT_;
  var n = Number(String(raw).trim());
  if (!isFinite(n) || n < 0) return CJ_WASH_QUEUE_PENDING_ROWS_MAX_DEFAULT_;
  return Math.floor(n);
}

function cjCatalogEnabled_() {
  var raw = PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_ENABLED_PROP_);
  if (raw == null || String(raw).trim() === '') return false;
  return String(raw).toLowerCase() !== 'false';
}

function cjCatalogCap_() {
  var n = parseInt(
    PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_CAP_PROP_) ||
      CJ_CATALOG_CAP_DEFAULT_,
    10
  );
  if (!(n > 0)) n = CJ_CATALOG_CAP_DEFAULT_;
  if (n > CJ_CATALOG_CAP_MAX_) n = CJ_CATALOG_CAP_MAX_;
  return n;
}

function cjCatalogIntervalMin_() {
  var n = parseInt(
    PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_INTERVAL_MIN_PROP_) ||
      CJ_CATALOG_INTERVAL_MIN_DEFAULT_,
    10
  );
  if (n <= 1) return 1;
  if (n <= 5) return 5;
  if (n <= 10) return 10;
  if (n <= 15) return 15;
  return 30;
}

function cjCatalogRearmMin_() {
  var n = parseInt(
    PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_REARM_MIN_PROP_) ||
      CJ_CATALOG_REARM_MIN_DEFAULT_,
    10
  );
  if (!(n > 0)) n = CJ_CATALOG_REARM_MIN_DEFAULT_;
  return Math.min(Math.max(1, n), 30);
}

function cjCatalogStaleGraceMin_() {
  var n = parseInt(
    PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_STALE_GRACE_MIN_PROP_) ||
      CJ_CATALOG_STALE_GRACE_MIN_DEFAULT_,
    10
  );
  if (!(n > 0)) n = CJ_CATALOG_STALE_GRACE_MIN_DEFAULT_;
  return n;
}

/** 親ロック updated がこの分より古ければ死んだ実行とみなし掃除（untilMs 破損にも対応）。 */
function cjCatalogLockStaleMin_() {
  var n = parseInt(
    PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_LOCK_STALE_MIN_PROP_) ||
      CJ_CATALOG_LOCK_STALE_MIN_DEFAULT_,
    10
  );
  if (!(n > 0)) n = CJ_CATALOG_LOCK_STALE_MIN_DEFAULT_;
  return Math.min(Math.max(15, n), 120);
}

function cjParseFetchedAtMs_(s) {
  s = String(s || '').trim();
  if (!s) return 0;
  try {
    var d = Utilities.parseDate(s, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var t = d && d.getTime ? d.getTime() : 0;
    return t > 0 ? t : 0;
  } catch (e) {
    return 0;
  }
}

function cjResearchSs_() {
  return SpreadsheetApp.openById(CJ_RESEARCH_SS_);
}

function cjCatalogGasStart_() {
  CJ_CATALOG_GAS_START_MS_ = new Date().getTime();
  CJ_SURVEY_ROOT_SETS_CACHE_ = null;
  CJ_WASH_QUEUE_ASIN_MAP_ = null;
  CJ_WASH_QUEUE_PENDING_ROWS_ = null;
}

function cjCatalogGasTimeUp_() {
  if (!CJ_CATALOG_GAS_START_MS_) return false;
  return new Date().getTime() - CJ_CATALOG_GAS_START_MS_ >= CJ_GAS_BUDGET_MS_;
}

function cjNowFetchedAt_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

function cjCatalogBand_(estimate) {
  if (estimate == null || estimate === '') return CJ_BAND_UNK_;
  var n = Number(estimate);
  if (isNaN(n)) return CJ_BAND_UNK_;
  if (n <= CJ_BAND_S_MAX_) return CJ_BAND_S_;
  if (n <= CJ_BAND_M_MAX_) return CJ_BAND_M_;
  if (n <= CJ_CATALOG_SPLIT_THRESHOLD_) return CJ_BAND_L_;
  return CJ_BAND_XL_;
}

/** ---- 共有親ロック（競合スプシ。①も同シート） ---- */
function cjEnsureSharedLockSheet_() {
  var ss = cjSs_();
  var sh = ss.getSheetByName(CJ_SHARED_LOCK_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(CJ_SHARED_LOCK_SHEET_);
    sh.getRange(1, 1, 1, 6).setValues([['メーカー', 'owner', 'lane', 'untilMs', 'runId', 'updated']]);
  }
  return sh;
}

/**
 * @return {{ok:boolean, why:string}}
 */
function cjSharedCatalogParentTryLock_(parent, owner, runId) {
  parent = String(parent || '').trim();
  owner = String(owner || CJ_SHARED_LOCK_OWNER_).trim() || CJ_SHARED_LOCK_OWNER_;
  if (!parent) return { ok: true, why: 'empty_parent' };
  var sh = cjEnsureSharedLockSheet_();
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var now = new Date().getTime();
    var lastRow = sh.getLastRow();
    var map = {};
    var rowOf = {};
    if (lastRow >= 2) {
      var vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
      var r;
      for (r = 0; r < vals.length; r++) {
        var p = String(vals[r][0] || '').trim();
        if (!p) continue;
        var until = Number(vals[r][3]) || 0;
        if (until <= now) continue;
        map[p] = {
          owner: String(vals[r][1] || ''),
          lane: String(vals[r][2] || ''),
          until: until,
          runId: String(vals[r][4] || ''),
        };
        rowOf[p] = r + 2;
      }
    }
    var cur = map[parent];
    if (cur && cur.until > now) {
      if (String(cur.owner) === owner && String(cur.lane) === 'enum') {
        var rowSame = rowOf[parent];
        if (rowSame) {
          sh.getRange(rowSame, 4, 1, 3).setValues([
            [now + CJ_SHARED_LOCK_TTL_MS_, String(runId || ''), cjNowFetchedAt_()],
          ]);
        }
        return { ok: true, why: 'renew' };
      }
      return {
        ok: false,
        why: 'busy_by_' + cur.owner + '_until=' + cur.until,
      };
    }
    var row = rowOf[parent];
    if (row) {
      sh.getRange(row, 1, 1, 6).setValues([
        [parent, owner, 'enum', now + CJ_SHARED_LOCK_TTL_MS_, String(runId || ''), cjNowFetchedAt_()],
      ]);
    } else {
      sh.appendRow([
        parent,
        owner,
        'enum',
        now + CJ_SHARED_LOCK_TTL_MS_,
        String(runId || ''),
        cjNowFetchedAt_(),
      ]);
    }
    SpreadsheetApp.flush();
    return { ok: true, why: 'acquired' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * until切れ、または updated が stale 分より古い行を空ける（途中死の残骸）。
 * @return {number}
 */
function cjHealStaleSharedLocks_(runId) {
  var sh = cjEnsureSharedLockSheet_();
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  var n = 0;
  try {
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return 0;
    var now = new Date().getTime();
    var staleMs = cjCatalogLockStaleMin_() * 60 * 1000;
    var vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      var parent = String(vals[r][0] || '').trim();
      if (!parent) continue;
      var until = Number(vals[r][3]) || 0;
      var updatedMs = cjParseFetchedAtMs_(vals[r][5]);
      var expired = !(until > now);
      var staleUpdated = updatedMs > 0 && now - updatedMs >= staleMs;
      if (!expired && !staleUpdated) continue;
      sh.getRange(r + 2, 1, 1, 6).setValues([['', '', '', '', '', '']]);
      n++;
    }
    if (n > 0) {
      SpreadsheetApp.flush();
      cjAppendLog_(runId || '', 'catalog_lock_heal', n, 0, 'cleared_stale_lock n=' + n);
      Logger.log('[競合Catalog] heal_stale_lock n=%s', n);
    }
    return n;
  } finally {
    lock.releaseLock();
  }
}

function cjSharedCatalogParentUnlock_(parent, owner) {
  parent = String(parent || '').trim();
  owner = String(owner || CJ_SHARED_LOCK_OWNER_).trim() || CJ_SHARED_LOCK_OWNER_;
  if (!parent) return;
  var sh = cjEnsureSharedLockSheet_();
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    var vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    var r;
    for (r = 0; r < vals.length; r++) {
      if (String(vals[r][0] || '').trim() !== parent) continue;
      if (String(vals[r][1] || '') !== owner) return;
      sh.getRange(r + 2, 1, 1, 6).setValues([['', '', '', '', '', '']]);
      return;
    }
  } finally {
    lock.releaseLock();
  }
}

function cjSharedCatalogBusyParents_() {
  var sh = cjSs_().getSheetByName(CJ_SHARED_LOCK_SHEET_);
  if (!sh || sh.getLastRow() < 2) return {};
  var now = new Date().getTime();
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  var out = {};
  var r;
  for (r = 0; r < vals.length; r++) {
    var p = String(vals[r][0] || '').trim();
    var until = Number(vals[r][3]) || 0;
    if (!p || until <= now) continue;
    out[p] = String(vals[r][1] || '');
  }
  return out;
}

/** ---- Catalog待ち選定（① prCycleCatalogPendingList_ と同趣旨） ---- */
function cjCatalogPendingList_(cap) {
  var want = Math.max(1, Number(cap) || 1);
  var sh = cjSs_().getSheetByName(CJ_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return [];
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var idxMaker = cjHeaderIndex_(hdr, 'メーカー');
  var idxCk = cjHeaderIndex_(hdr, CJ_CK_HEADER_);
  var idxJ = cjHeaderIndex_(hdr, '実態');
  var idxProg = cjHeaderIndex_(hdr, CJ_WASH_PROGRESS_HEADER_);
  var idxEst = cjHeaderIndex_(hdr, CJ_CATALOG_EST_COL_);
  var idxBand = cjHeaderIndex_(hdr, CJ_BAND_COL_);
  var idxSt = cjHeaderIndex_(hdr, CJ_WASH_STATUS_HEADER_);
  if (idxMaker < 0) return [];
  var busy = cjSharedCatalogBusyParents_();
  var unkList = [];
  var missList = [];
  var r;
  for (r = 0; r < vals.length; r++) {
    var parent = String(vals[r][idxMaker] || '').trim();
    if (!parent) continue;
    if (CJ_HIT_EXCLUDE_[parent] || CJ_EXCLUDE_SEED_[parent]) continue;
    if (idxCk >= 0 && cjCkIsTrue_(vals[r][idxCk])) continue;
    var jitta = idxJ >= 0 ? String(vals[r][idxJ] || '').trim() : '';
    if (jitta !== CJ_JITTAI_MAKER_ && jitta !== CJ_JITTAI_BRAND_) continue;
    var st = idxSt >= 0 ? String(vals[r][idxSt] || '').trim() : '';
    if (st === '完了' || st === 'Keepa中') continue;
    var prog = idxProg >= 0 ? String(vals[r][idxProg] || '') : '';
    if (/title0_skip/.test(prog)) continue;
    if (busy[parent] && busy[parent] !== CJ_SHARED_LOCK_OWNER_) continue;
    var estCell = idxEst >= 0 ? vals[r][idxEst] : '';
    var nEst = Number(String(estCell == null ? '' : estCell).replace(/[^\d.-]/g, ''));
    var estOk = !isNaN(nEst) && nEst > 0 && String(estCell).trim() !== '';
    var band = idxBand >= 0 ? String(vals[r][idxBand] || '').trim() : '';
    if (!estOk) {
      if (!band || band === CJ_BAND_UNK_) {
        if (unkList.length < want) unkList.push(parent);
      } else if (missList.length < want) {
        missList.push(parent);
      }
    }
    if (unkList.length >= want) break;
  }
  return unkList.concat(missList).slice(0, want);
}

function cjCatalogEntry_(parent) {
  parent = String(parent || '').trim();
  var key = cjNormalizeMakerKey_(parent) || parent;
  return key || parent;
}

/** ---- 調査ルート（①プロファイルを読取） ---- */
function cjSurveyRootMatchSets_() {
  if (CJ_SURVEY_ROOT_SETS_CACHE_) return CJ_SURVEY_ROOT_SETS_CACHE_;
  var ids = {};
  var labels = {};
  try {
    var sh = cjResearchSs_().getSheetByName(CJ_PROF_SHEET_);
    if (sh && sh.getLastRow() >= 2) {
      var lastCol = Math.max(sh.getLastColumn(), 1);
      var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
      var idxCat = cjHeaderIndex_(hdr, 'カテゴリ');
      var idxCk = cjHeaderIndex_(hdr, CJ_PROF_SURVEY_CK_);
      var idxId = cjHeaderIndex_(hdr, CJ_PROF_ROOT_ID_);
      var r;
      for (r = 0; r < vals.length; r++) {
        var cat = idxCat >= 0 ? String(vals[r][idxCat] || '').trim() : '';
        if (cat === '共通' || /^\d+$/.test(cat)) continue;
        if (idxCk >= 0 && !cjCkIsTrue_(vals[r][idxCk])) continue;
        var id = idxId >= 0 ? String(vals[r][idxId] || '').trim().replace(/\.0$/, '') : '';
        if (id) ids[id] = true;
        var lab = String(cat || '')
          .replace(/[＆&]/g, '・')
          .replace(/\s+/g, '')
          .toLowerCase();
        if (lab) labels[lab] = true;
      }
    }
  } catch (eProf) {
    Logger.log('[競合Catalog] survey_prof_err %s', String((eProf && eProf.message) || eProf).slice(0, 120));
  }
  if (!Object.keys(ids).length) {
    ids[CJ_FOOD_ROOT_] = true;
    labels['食品・飲料・お酒'.replace(/[＆&]/g, '・').replace(/\s+/g, '').toLowerCase()] = true;
  }
  CJ_SURVEY_ROOT_SETS_CACHE_ = { ids: ids, labels: labels };
  return CJ_SURVEY_ROOT_SETS_CACHE_;
}

function cjSurveyRootIdMatches_(raw) {
  var id = String(raw == null ? '' : raw)
    .trim()
    .replace(/\.0$/, '');
  return !!(id && cjSurveyRootMatchSets_().ids[id]);
}

function cjSurveyRootLabelMatches_(name) {
  var lab = String(name || '')
    .replace(/[＆&]/g, '・')
    .replace(/\s+/g, '')
    .toLowerCase();
  return !!(lab && cjSurveyRootMatchSets_().labels[lab]);
}

function cjCatalogClassifNodeId_(n) {
  if (!n || typeof n !== 'object') return '';
  return String(n.classificationId || n.id || '')
    .trim()
    .replace(/\.0$/, '');
}

function cjCatalogClassifNodeName_(n) {
  if (!n || typeof n !== 'object') return '';
  return String(n.displayName || n.name || '').trim();
}

function cjCatalogClassifParentChain_(node) {
  var out = [];
  var cur = node;
  var guard = 0;
  while (cur && typeof cur === 'object' && guard < 24) {
    out.push(cur);
    cur = cur.parent;
    guard++;
  }
  return out;
}

function cjCatalogClassifNodesFromBlock_(block) {
  if (!block) return [];
  var nodes = block.classifications;
  if (nodes && nodes.length) return nodes;
  if (block.displayName || block.classificationId || block.name || block.id) return [block];
  if (Array.isArray(block)) return block;
  return [];
}

function cjCatalogRootFromClassifications_(classifications) {
  if (!classifications || !classifications.length) return { rootId: '', rootName: '' };
  var fallback = null;
  var c;
  for (c = 0; c < classifications.length; c++) {
    var nodes = cjCatalogClassifNodesFromBlock_(classifications[c]);
    var ni;
    for (ni = 0; ni < nodes.length; ni++) {
      var chain = cjCatalogClassifParentChain_(nodes[ni]);
      if (!chain.length) continue;
      var k;
      for (k = chain.length - 1; k >= 0; k--) {
        var id = cjCatalogClassifNodeId_(chain[k]);
        var nm = cjCatalogClassifNodeName_(chain[k]);
        if (cjSurveyRootIdMatches_(id) || cjSurveyRootLabelMatches_(nm)) {
          return { rootId: id, rootName: nm };
        }
      }
      var tip = chain[chain.length - 1];
      if (!fallback) {
        fallback = {
          rootId: cjCatalogClassifNodeId_(tip),
          rootName: cjCatalogClassifNodeName_(tip),
        };
      }
    }
  }
  return fallback || { rootId: '', rootName: '' };
}

function cjCatalogFilterSurveyRoot_(items) {
  var list = items || [];
  var keep = [];
  var drop = 0;
  var miss = 0;
  var i;
  for (i = 0; i < list.length; i++) {
    var it = list[i] || {};
    var rootId = String(it.rootId || '').trim().replace(/\.0$/, '');
    var rootName = String(it.rootName || '').trim();
    if (!rootId && !rootName) {
      miss++;
      continue;
    }
    if (cjSurveyRootIdMatches_(rootId) || cjSurveyRootLabelMatches_(rootName)) {
      keep.push(it);
    } else {
      drop++;
    }
  }
  return { keep: keep, drop: drop, miss: miss };
}

/** ---- SP-API Catalog ---- */
function cjSpapiAccessToken_() {
  var props = PropertiesService.getScriptProperties();
  var clientId = String(props.getProperty('SPAPI_LWA_CLIENT_ID') || '').trim();
  var clientSecret = String(props.getProperty('SPAPI_LWA_CLIENT_SECRET') || '').trim();
  var refreshToken = String(props.getProperty('SPAPI_REFRESH_TOKEN') || '').trim();
  var missing = [];
  if (!clientId) missing.push('SPAPI_LWA_CLIENT_ID');
  if (!clientSecret) missing.push('SPAPI_LWA_CLIENT_SECRET');
  if (!refreshToken) missing.push('SPAPI_REFRESH_TOKEN');
  if (missing.length) {
    throw new Error('Script Properties 不足: ' + missing.join(', '));
  }
  var resp = UrlFetchApp.fetch(CJ_SPAPI_LWA_URL_, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded;charset=UTF-8',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret,
    },
    muteHttpExceptions: true,
  });
  var code = resp.getResponseCode();
  var text = resp.getContentText() || '';
  if (code !== 200) throw new Error('LWA HTTP ' + code + ' ' + text.slice(0, 180));
  var json = JSON.parse(text);
  var token = String(json.access_token || '').trim();
  if (!token) throw new Error('LWA access_token 空');
  return token;
}

function cjSpapiEndpoint_() {
  var raw = String(PropertiesService.getScriptProperties().getProperty('SPAPI_ENDPOINT') || '').trim();
  if (!raw) return CJ_SPAPI_DEFAULT_ENDPOINT_;
  raw = raw.replace(/^(https?):\\+/i, '$1://');
  raw = raw.replace(/^(https?):\/(?!\/)/i, '$1://');
  raw = raw.replace(/\\/g, '/').replace(/\/+$/, '');
  if (!/^https?:\/\//i.test(raw)) return CJ_SPAPI_DEFAULT_ENDPOINT_;
  return raw;
}

function cjSpapiMarketplace_() {
  var mp = String(PropertiesService.getScriptProperties().getProperty('SPAPI_MARKETPLACE_ID') || '').trim();
  return mp || CJ_SPAPI_DEFAULT_MP_;
}

function cjCatalogSearchPage_(keywords, pageToken) {
  var endpoint = cjSpapiEndpoint_();
  var mp = cjSpapiMarketplace_();
  var token = cjSpapiAccessToken_();
  var url =
    endpoint +
    '/catalog/2022-04-01/items?' +
    'marketplaceIds=' +
    encodeURIComponent(mp) +
    '&keywords=' +
    encodeURIComponent(String(keywords || '')) +
    '&includedData=' +
    encodeURIComponent('summaries,classifications,salesRanks') +
    '&pageSize=20';
  var pt = String(pageToken || '').trim();
  if (pt) url += '&pageToken=' + encodeURIComponent(pt);
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      Accept: 'application/json',
      'x-amz-access-token': token,
      'User-Agent': 'CompetitorJittaCatalog/1.0 (GAS)',
    },
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '';
  if (code !== 200) {
    throw new Error('Catalog HTTP ' + code + ' ' + body.slice(0, 160));
  }
  var data = JSON.parse(body);
  var items = [];
  var list = data.items || [];
  var i;
  for (i = 0; i < list.length; i++) {
    var it = list[i] || {};
    var asin = String(it.asin || '')
      .trim()
      .toUpperCase();
    var title = '';
    var brand = '';
    var sums = it.summaries || [];
    if (sums.length && sums[0]) {
      title = String(sums[0].itemName || '').trim();
      brand = String(sums[0].brand || '').trim();
    }
    if (asin) {
      var root = cjCatalogRootFromClassifications_(it.classifications);
      items.push({
        asin: asin,
        title: title,
        brand: brand,
        rootId: root.rootId,
        rootName: root.rootName,
      });
    }
  }
  var nextToken = '';
  if (data.pagination && data.pagination.nextToken) {
    nextToken = String(data.pagination.nextToken);
  }
  return { items: items, nextToken: nextToken };
}

/** ---- 洗いキュー（①調査スプシ・① SNAPSHOT/COUNT_MAP と同思想） ---- */
function cjEnsureWashQueueSheet_() {
  var ss = cjResearchSs_();
  var sh = ss.getSheetByName(CJ_WASH_QUEUE_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(CJ_WASH_QUEUE_SHEET_);
    sh.getRange(1, 1, 1, CJ_WASH_QUEUE_HEADERS_.length).setValues([CJ_WASH_QUEUE_HEADERS_]);
  }
  return sh;
}

/** 洗いキュー全行を1回だけ読み、メーカー→ASIN集合。未済行数も同時計測。 */
function cjWashQueueAsinMapLoad_() {
  if (CJ_WASH_QUEUE_ASIN_MAP_) return CJ_WASH_QUEUE_ASIN_MAP_;
  var t0 = new Date().getTime();
  var map = {};
  var pendingRows = 0;
  var sh = cjEnsureWashQueueSheet_();
  if (sh.getLastRow() < 2) {
    CJ_WASH_QUEUE_ASIN_MAP_ = map;
    CJ_WASH_QUEUE_PENDING_ROWS_ = 0;
    return map;
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = cjHeaderIndex_(hdr, 'メーカー');
  var idxA = cjHeaderIndex_(hdr, 'ASIN');
  var idxK = cjHeaderIndex_(hdr, 'Keepa済');
  if (idxM < 0 || idxA < 0) {
    CJ_WASH_QUEUE_ASIN_MAP_ = map;
    CJ_WASH_QUEUE_PENDING_ROWS_ = 0;
    return map;
  }
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    var p = String(vals[r][idxM] || '').trim();
    var a = String(vals[r][idxA] || '')
      .trim()
      .toUpperCase();
    if (!p || !a) continue;
    if (!map[p]) map[p] = {};
    map[p][a] = true;
    if (idxK < 0 || !cjCkIsTrue_(vals[r][idxK])) pendingRows++;
  }
  CJ_WASH_QUEUE_ASIN_MAP_ = map;
  CJ_WASH_QUEUE_PENDING_ROWS_ = pendingRows;
  Logger.log(
    '[競合Catalog] queue_asin_map rows=%s pendingRows=%s makers=%s elapsedMs=%s',
    vals.length,
    pendingRows,
    Object.keys(map).length,
    new Date().getTime() - t0
  );
  return map;
}

/** 未済が上限以上かつ当該社に既存行が無いとき追記禁止。 */
function cjWashQueuePendingRowsCapBlocksAppend_(parent) {
  var maxN = cjWashQueuePendingRowsMax_();
  if (!(maxN > 0)) return false;
  cjWashQueueAsinMapLoad_();
  var pendingRows = Number(CJ_WASH_QUEUE_PENDING_ROWS_) || 0;
  if (pendingRows < maxN) return false;
  parent = String(parent || '').trim();
  if (!parent) return true;
  return cjWashQueueSheetCount_(parent) === 0;
}

/** メモリマップ上の当該社 ASIN 数（シート追記済みぶん）。 */
function cjWashQueueSheetCount_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return 0;
  var map = CJ_WASH_QUEUE_ASIN_MAP_ || {};
  var existing = map[parent];
  if (!existing) return 0;
  return Object.keys(existing).length;
}

/** 当該メーカーの ASIN 集合（マップからのコピー。列挙中の seen は呼び出し側が更新）。 */
function cjWashQueueAsinSet_(parent) {
  parent = String(parent || '').trim();
  var map = cjWashQueueAsinMapLoad_();
  var src = (parent && map[parent]) || {};
  var seen = {};
  var a;
  for (a in src) {
    if (Object.prototype.hasOwnProperty.call(src, a)) seen[a] = true;
  }
  return seen;
}

/**
 * キュー末尾へ追記。既存判定はメモリマップのみ（全行再読しない）。
 * @return {number} 追加行数
 */
function cjWashQueueAppendAsins_(parent, entry, asins) {
  if (!asins || !asins.length) return 0;
  parent = String(parent || '').trim();
  if (cjWashQueuePendingRowsCapBlocksAppend_(parent)) {
    Logger.log(
      '[競合Catalog] pending_rows_cap skip_append parent=%s n=%s pendingRows=%s max=%s',
      parent,
      asins.length,
      Number(CJ_WASH_QUEUE_PENDING_ROWS_) || 0,
      cjWashQueuePendingRowsMax_()
    );
    return 0;
  }
  var map = cjWashQueueAsinMapLoad_();
  if (!map[parent]) map[parent] = {};
  var existing = map[parent];
  var sh = cjEnsureWashQueueSheet_();
  var now = cjNowFetchedAt_();
  var lines = [];
  var i;
  for (i = 0; i < asins.length; i++) {
    var a = String(asins[i] || '')
      .trim()
      .toUpperCase();
    if (!a || existing[a]) continue;
    existing[a] = true;
    lines.push([parent, entry, a, now, false, '', 'comp_catalog']);
  }
  if (!lines.length) return 0;
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, lines.length, CJ_WASH_QUEUE_HEADERS_.length).setValues(lines);
  if (CJ_WASH_QUEUE_PENDING_ROWS_ != null) {
    CJ_WASH_QUEUE_PENDING_ROWS_ = Number(CJ_WASH_QUEUE_PENDING_ROWS_) + lines.length;
  }
  return lines.length;
}

function cjCatalogEnumerateToQueue_(parent, entry) {
  var token = '';
  var pages = 0;
  var added = 0;
  var seen = cjWashQueueAsinSet_(parent);
  var stopAt = CJ_CATALOG_SPLIT_THRESHOLD_ + 1;
  while (pages < 200) {
    if (Object.keys(seen).length >= stopAt) {
      return { added: added, pages: pages, queueN: Object.keys(seen).length, done: false, cut501: true };
    }
    if (cjCatalogGasTimeUp_()) {
      return { added: added, pages: pages, queueN: Object.keys(seen).length, done: false, timeup: true };
    }
    var cat = cjCatalogSearchPage_(entry, token);
    pages++;
    var kept = cjCatalogFilterSurveyRoot_(cat.items).keep;
    var batch = [];
    var i;
    for (i = 0; i < kept.length; i++) {
      var a = String(kept[i].asin || '').toUpperCase();
      if (!a || seen[a]) continue;
      seen[a] = true;
      batch.push(a);
    }
    if (batch.length && cjWashQueuePendingRowsCapBlocksAppend_(parent)) {
      Logger.log(
        '[競合Catalog] pending_rows_cap stop_enum parent=%s pendingRows=%s max=%s',
        parent,
        Number(CJ_WASH_QUEUE_PENDING_ROWS_) || 0,
        cjWashQueuePendingRowsMax_()
      );
      // queueN はシート実在だけ（seen は幽霊規模帯の原因になるので使わない）
      return {
        added: added,
        pages: pages,
        queueN: cjWashQueueSheetCount_(parent),
        seenN: Object.keys(seen).length,
        done: false,
        cut501: false,
        pendingCap: true,
      };
    }
    added += cjWashQueueAppendAsins_(parent, entry, batch);
    token = cat.nextToken || '';
    if (!token) {
      return { added: added, pages: pages, queueN: Object.keys(seen).length, done: true, cut501: false };
    }
    Utilities.sleep(200);
  }
  return {
    added: added,
    pages: pages,
    queueN: Object.keys(seen).length,
    done: !token,
    cut501: Object.keys(seen).length >= stopAt,
  };
}

function cjFindMakerSumRow_(parent) {
  parent = String(parent || '').trim();
  var sh = cjSs_().getSheetByName(CJ_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return null;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = cjHeaderIndex_(hdr, 'メーカー');
  if (idxM < 0) return null;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var r;
  for (r = 0; r < vals.length; r++) {
    if (String(vals[r][idxM] || '').trim() === parent) {
      return {
        sh: sh,
        sheetRow: r + 2,
        hdr: hdr,
        idxEst: cjHeaderIndex_(hdr, CJ_CATALOG_EST_COL_),
        idxBand: cjHeaderIndex_(hdr, CJ_BAND_COL_),
        idxSt: cjHeaderIndex_(hdr, CJ_WASH_STATUS_HEADER_),
        idxProg: cjHeaderIndex_(hdr, CJ_WASH_PROGRESS_HEADER_),
      };
    }
  }
  return null;
}

function cjWriteMakerSumBand_(parent, queueN, cut501) {
  var loc = cjFindMakerSumRow_(parent);
  if (!loc) return { ok: false, band: cjCatalogBand_(queueN) };
  var cols = cjEnsureCols_(loc.sh, [CJ_CATALOG_EST_COL_, CJ_BAND_COL_, CJ_WASH_STATUS_HEADER_, CJ_WASH_PROGRESS_HEADER_]);
  var writeN = cut501 ? Math.max(queueN, CJ_CATALOG_SPLIT_THRESHOLD_ + 1) : queueN;
  var band = cjCatalogBand_(writeN);
  loc.sh.getRange(loc.sheetRow, cols[CJ_CATALOG_EST_COL_] + 1).setValue(writeN);
  loc.sh.getRange(loc.sheetRow, cols[CJ_BAND_COL_] + 1).setValue(band);
  loc.sh.getRange(loc.sheetRow, cols[CJ_WASH_STATUS_HEADER_] + 1).setValue('');
  loc.sh
    .getRange(loc.sheetRow, cols[CJ_WASH_PROGRESS_HEADER_] + 1)
    .setValue('enum_done queue=' + queueN + ' band=' + band + ' via=comp');
  return { ok: true, band: band, queueN: queueN };
}

function cjMarkTitle0_(parent, detail) {
  var loc = cjFindMakerSumRow_(parent);
  if (!loc) return;
  var cols = cjEnsureCols_(loc.sh, [CJ_WASH_STATUS_HEADER_, CJ_WASH_PROGRESS_HEADER_]);
  loc.sh.getRange(loc.sheetRow, cols[CJ_WASH_STATUS_HEADER_] + 1).setValue('');
  loc.sh
    .getRange(loc.sheetRow, cols[CJ_WASH_PROGRESS_HEADER_] + 1)
    .setValue(('title0_skip ' + String(detail || '')).slice(0, 180));
}

/** 失敗時に列挙中／comp_enum を残さない（同じ3社ループ防止）。 */
function cjClearCompEnumBadge_(parent) {
  parent = String(parent || '').trim();
  if (!parent) return;
  var loc = cjFindMakerSumRow_(parent);
  if (!loc) return;
  var cols = cjEnsureCols_(loc.sh, [CJ_WASH_STATUS_HEADER_, CJ_WASH_PROGRESS_HEADER_]);
  var prog = String(
    loc.sh.getRange(loc.sheetRow, cols[CJ_WASH_PROGRESS_HEADER_] + 1).getValue() || ''
  );
  if (prog.indexOf('title0_skip') >= 0 || prog.indexOf('enum_done') >= 0) return;
  loc.sh.getRange(loc.sheetRow, cols[CJ_WASH_STATUS_HEADER_] + 1).setValue('');
  loc.sh.getRange(loc.sheetRow, cols[CJ_WASH_PROGRESS_HEADER_] + 1).setValue('');
}

/**
 * キュー0なのに via=comp の enum_done／規模帯が残っている幽霊を消す。
 * 完了／Keepa中／調査対象外／要修復は触らない（完走後の済削除でキュー0は正当）。
 * @return {number}
 */
function cjHealPhantomCompBands_(runId) {
  cjWashQueueAsinMapLoad_();
  var map = CJ_WASH_QUEUE_ASIN_MAP_ || {};
  var sh = cjSs_().getSheetByName(CJ_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = cjHeaderIndex_(hdr, 'メーカー');
  var idxSt = cjHeaderIndex_(hdr, CJ_WASH_STATUS_HEADER_);
  var idxProg = cjHeaderIndex_(hdr, CJ_WASH_PROGRESS_HEADER_);
  var idxEst = cjHeaderIndex_(hdr, CJ_CATALOG_EST_COL_);
  var idxBand = cjHeaderIndex_(hdr, CJ_BAND_COL_);
  if (idxM < 0 || idxProg < 0) return 0;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var n = 0;
  var r;
  for (r = 0; r < vals.length; r++) {
    var parent = String(vals[r][idxM] || '').trim();
    if (!parent) continue;
    var prog = String(vals[r][idxProg] || '');
    if (prog.indexOf('via=comp') < 0 || prog.indexOf('enum_done') < 0) continue;
    var st = idxSt >= 0 ? String(vals[r][idxSt] || '').trim() : '';
    if (st === '完了' || st === '調査対象外' || st === '要修復' || st === 'Keepa中') continue;
    var existing = map[parent];
    if (existing && Object.keys(existing).length > 0) continue;
    if (idxEst >= 0) sh.getRange(r + 2, idxEst + 1).setValue('');
    if (idxBand >= 0) sh.getRange(r + 2, idxBand + 1).setValue('');
    if (idxSt >= 0) sh.getRange(r + 2, idxSt + 1).setValue('');
    sh.getRange(r + 2, idxProg + 1).setValue('');
    n++;
  }
  if (n > 0) {
    cjAppendLog_(runId || '', 'catalog_phantom_heal', n, 0, 'cleared_phantom_band n=' + n);
    Logger.log('[競合Catalog] phantom_band_heal n=%s', n);
  }
  return n;
}

/**
 * 前回落ちで残った comp_enum／列挙中を掃除してから選定する。
 * @return {number}
 */
function cjHealStuckCompEnum_(runId) {
  var sh = cjSs_().getSheetByName(CJ_SUM_SHEET_);
  if (!sh || sh.getLastRow() < 2) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxM = cjHeaderIndex_(hdr, 'メーカー');
  var idxSt = cjHeaderIndex_(hdr, CJ_WASH_STATUS_HEADER_);
  var idxProg = cjHeaderIndex_(hdr, CJ_WASH_PROGRESS_HEADER_);
  if (idxM < 0 || idxProg < 0) return 0;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  var n = 0;
  var r;
  for (r = 0; r < vals.length; r++) {
    var prog = String(vals[r][idxProg] || '');
    if (prog.indexOf('comp_enum') < 0) continue;
    if (prog.indexOf('enum_done') >= 0) continue;
    var st = idxSt >= 0 ? String(vals[r][idxSt] || '').trim() : '';
    sh.getRange(r + 2, idxProg + 1).setValue('');
    if (idxSt >= 0 && (st === CJ_WASH_ST_ENUM_ || st === '列挙中')) {
      sh.getRange(r + 2, idxSt + 1).setValue('');
    }
    n++;
  }
  if (n > 0) {
    cjAppendLog_(runId || '', 'catalog_heal', n, 0, 'cleared_comp_enum n=' + n);
    Logger.log('[競合Catalog] heal_comp_enum n=%s', n);
  }
  return n;
}

/**
 * 1社列挙（規模確定）。
 * @return {{ok:boolean, skip?:boolean, band?:string, queueN?:number, detail:string}}
 */
function cjRunCatalogOne_(parent, runId) {
  parent = String(parent || '').trim();
  if (!parent) return { ok: false, detail: 'parent_empty' };
  var lock = cjSharedCatalogParentTryLock_(parent, CJ_SHARED_LOCK_OWNER_, runId);
  if (!lock.ok) {
    return { ok: false, skip: true, detail: 'lock_' + lock.why };
  }
  var badgeOn = false;
  try {
    var entry = cjCatalogEntry_(parent);
    var loc0 = cjFindMakerSumRow_(parent);
    if (loc0) {
      var cols0 = cjEnsureCols_(loc0.sh, [CJ_WASH_STATUS_HEADER_, CJ_WASH_PROGRESS_HEADER_]);
      loc0.sh.getRange(loc0.sheetRow, cols0[CJ_WASH_STATUS_HEADER_] + 1).setValue(CJ_WASH_ST_ENUM_);
      loc0.sh.getRange(loc0.sheetRow, cols0[CJ_WASH_PROGRESS_HEADER_] + 1).setValue('comp_enum');
      badgeOn = true;
    }
    var er = cjCatalogEnumerateToQueue_(parent, entry);
    var queueN = Number(er.queueN) || 0;
    if (er.pendingCap) {
      if (badgeOn) cjClearCompEnumBadge_(parent);
      badgeOn = false;
      var capDetail =
        'pending_cap parent=' +
        parent +
        ' pendingRows=' +
        (Number(CJ_WASH_QUEUE_PENDING_ROWS_) || 0) +
        ' max=' +
        cjWashQueuePendingRowsMax_() +
        ' seen=' +
        (Number(er.seenN) || 0) +
        ' added=' +
        (Number(er.added) || 0) +
        ' sheetQueue=' +
        cjWashQueueSheetCount_(parent);
      cjAppendLog_(runId || '', 'catalog_pending_cap', 0, 0, capDetail.slice(0, 500));
      Logger.log('[競合Catalog] %s', capDetail);
      return {
        ok: true,
        skip: true,
        pendingCap: true,
        queueN: 0,
        detail: capDetail,
      };
    }
    if (!queueN && er.done) {
      var alt = cjNormalizeMakerKey_(parent);
      if (alt && alt !== entry) {
        er = cjCatalogEnumerateToQueue_(parent, alt);
        queueN = Number(er.queueN) || 0;
        entry = alt;
        if (er.pendingCap) {
          if (badgeOn) cjClearCompEnumBadge_(parent);
          badgeOn = false;
          var capDetail2 =
            'pending_cap parent=' +
            parent +
            ' pendingRows=' +
            (Number(CJ_WASH_QUEUE_PENDING_ROWS_) || 0) +
            ' max=' +
            cjWashQueuePendingRowsMax_() +
            ' seen=' +
            (Number(er.seenN) || 0) +
            ' added=' +
            (Number(er.added) || 0) +
            ' sheetQueue=' +
            cjWashQueueSheetCount_(parent);
          cjAppendLog_(runId || '', 'catalog_pending_cap', 0, 0, capDetail2.slice(0, 500));
          return {
            ok: true,
            skip: true,
            pendingCap: true,
            queueN: 0,
            detail: capDetail2,
          };
        }
      }
    }
    if (!queueN && er.done) {
      cjMarkTitle0_(parent, 'entry=' + entry + ' pages=' + er.pages);
      badgeOn = false;
      return { ok: true, skip: true, queueN: 0, band: CJ_BAND_UNK_, detail: 'title0 entry=' + entry };
    }
    // シートに1行も無いのに規模帯を書かない（seen だけの幽霊を防ぐ）
    var sheetN = cjWashQueueSheetCount_(parent);
    if (!(sheetN > 0) && !(Number(er.added) > 0)) {
      if (badgeOn) cjClearCompEnumBadge_(parent);
      badgeOn = false;
      var noQ =
        'no_sheet_queue parent=' +
        parent +
        ' seen=' +
        (Number(er.queueN) || 0) +
        ' added=' +
        (Number(er.added) || 0);
      cjAppendLog_(runId || '', 'catalog_skip', 0, 0, noQ.slice(0, 500));
      return { ok: true, skip: true, queueN: 0, detail: noQ };
    }
    queueN = sheetN > 0 ? sheetN : queueN;
    var fin = cjWriteMakerSumBand_(
      parent,
      queueN,
      !!er.cut501 || queueN >= CJ_CATALOG_SPLIT_THRESHOLD_ + 1
    );
    badgeOn = false;
    return {
      ok: true,
      band: fin.band,
      queueN: queueN,
      detail:
        'band=' +
        fin.band +
        ' queue=' +
        queueN +
        ' pages=' +
        er.pages +
        (er.cut501 ? ' cut501=1' : '') +
        (er.timeup ? ' timeup=1' : ''),
    };
  } catch (eOne) {
    var msg = String((eOne && eOne.message) || eOne).slice(0, 160);
    if (badgeOn) cjClearCompEnumBadge_(parent);
    cjAppendLog_(runId || '', 'catalog_skip', 0, 0, (parent + ' err=' + msg).slice(0, 500));
    return { ok: false, skip: true, detail: 'err=' + msg };
  } finally {
    cjSharedCatalogParentUnlock_(parent, CJ_SHARED_LOCK_OWNER_);
  }
}

function cjClearCatalogTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === CJ_CATALOG_TRIGGER_HANDLER_) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function cjScheduleCatalogTrigger_(afterMin, tag) {
  cjClearCatalogTriggers_();
  var mins = Math.max(1, Math.floor(Number(afterMin) || cjCatalogRearmMin_()));
  var expect = new Date().getTime() + mins * 60 * 1000;
  PropertiesService.getScriptProperties().setProperty(CJ_CATALOG_EXPECT_FIRE_MS_PROP_, String(expect));
  ScriptApp.newTrigger(CJ_CATALOG_TRIGGER_HANDLER_).timeBased().after(mins * 60 * 1000).create();
  Logger.log('[競合Catalog] arm afterMin=%s tag=%s expectMs=%s', mins, tag || '', expect);
}

function cjCatalogInflightUntil_() {
  return Number(
    PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_INFLIGHT_UNTIL_PROP_) || 0
  );
}

function cjCatalogSetInflight_(untilMs) {
  PropertiesService.getScriptProperties().setProperty(
    CJ_CATALOG_INFLIGHT_UNTIL_PROP_,
    String(untilMs)
  );
}

function cjCatalogClearInflight_() {
  PropertiesService.getScriptProperties().deleteProperty(CJ_CATALOG_INFLIGHT_UNTIL_PROP_);
}

function cjCountCatalogTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  var n = 0;
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === CJ_CATALOG_TRIGGER_HANDLER_) n++;
  }
  return n;
}

/**
 * 予約ありだが発火予定が無い／過ぎている＝幽霊。
 * 予定なし＋トリガーありも幽霊（①案P1と同型）。
 */
function cjCatalogTriggerIsStale_() {
  var n = cjCountCatalogTriggers_();
  if (!(n > 0)) return false;
  var expect = Number(
    PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_EXPECT_FIRE_MS_PROP_) || 0
  );
  if (!(expect > 0)) return true;
  var graceMs = cjCatalogStaleGraceMin_() * 60 * 1000;
  return new Date().getTime() > expect + graceMs;
}

/**
 * Catalog時計が無いか幽霊なら張り直す。実行中リース／未来の発火予定があるときは触らない。
 * 実態AI・ヒット・失敗 finally から呼ぶ（新番犬時計は増やさない）。
 * @return {boolean} 予約したか
 */
function cjMaybeArmCatalogRecovery_(runId, tag) {
  if (!cjCatalogEnabled_()) return false;
  var now = new Date().getTime();
  if (cjCatalogInflightUntil_() > now) return false;
  var expect = Number(
    PropertiesService.getScriptProperties().getProperty(CJ_CATALOG_EXPECT_FIRE_MS_PROP_) || 0
  );
  if (expect > now) return false;
  var n = cjCountCatalogTriggers_();
  if (n > 0 && !cjCatalogTriggerIsStale_()) return false;
  if (n > 0) cjClearCatalogTriggers_();
  var armTag = String(tag || 'recovery');
  cjScheduleCatalogTrigger_(cjCatalogRearmMin_(), armTag);
  cjAppendLog_(runId || '', 'catalog_rearm', 0, 0, armTag);
  Logger.log('[競合Catalog] recovery_arm tag=%s runId=%s', armTag, runId || '');
  return true;
}

/**
 * Catalog待ちを cap 社まで列挙。
 */
function cjRunCatalogBatch_(runId) {
  if (!cjCatalogEnabled_()) {
    return { n: 0, detail: 'catalog_off' };
  }
  cjCatalogGasStart_();
  runId = String(runId || 'cj_cat_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss'));
  var healed = 0;
  try {
    cjHealStaleSharedLocks_(runId);
  } catch (eLockHeal) {
    Logger.log(
      '[競合Catalog] WARN lock_heal %s',
      String((eLockHeal && eLockHeal.message) || eLockHeal).slice(0, 120)
    );
  }
  try {
    healed = cjHealStuckCompEnum_(runId);
  } catch (eHeal) {
    Logger.log(
      '[競合Catalog] WARN heal %s',
      String((eHeal && eHeal.message) || eHeal).slice(0, 120)
    );
  }
  // キューはバッチ先頭で1回（以降はメモリ）。失敗しても列挙自体は試す。
  try {
    cjWashQueueAsinMapLoad_();
  } catch (eMap) {
    var mapMsg = String((eMap && eMap.message) || eMap).slice(0, 160);
    cjAppendLog_(runId, 'catalog_skip', 0, 0, 'queue_map_load err=' + mapMsg);
    return { n: 0, detail: 'queue_map_fail ' + mapMsg, left: true };
  }
  var phantomCleared = 0;
  try {
    phantomCleared = cjHealPhantomCompBands_(runId);
  } catch (ePh) {
    Logger.log(
      '[競合Catalog] WARN phantom_heal %s',
      String((ePh && ePh.message) || ePh).slice(0, 120)
    );
  }
  var cap = cjCatalogCap_();
  var parents = cjCatalogPendingList_(cap);
  var pendingRows = Number(CJ_WASH_QUEUE_PENDING_ROWS_) || 0;
  var pendingMax = cjWashQueuePendingRowsMax_();
  // 未済が上限以上: 新規社は追記できないので API しない。キューあり（途中列挙）だけ残す
  if (pendingMax > 0 && pendingRows >= pendingMax) {
    parents = parents.filter(function (p) {
      return cjWashQueueSheetCount_(p) > 0;
    });
    if (!parents.length) {
      var batchSkip =
        'batch_skip pendingRows=' +
        pendingRows +
        ' max=' +
        pendingMax +
        (phantomCleared ? ' phantom=' + phantomCleared : '') +
        (healed ? ' heal=' + healed : '');
      cjAppendLog_(runId, 'catalog_pending_cap', 0, 0, batchSkip.slice(0, 500));
      Logger.log('[競合Catalog] pending_cap %s', batchSkip);
      return {
        n: 0,
        detail: 'catalog_pending_cap ' + batchSkip,
        left: true,
        pendingCapBatch: true,
      };
    }
  }
  var doneN = 0;
  var skipN = 0;
  var bands = {};
  var i;
  for (i = 0; i < parents.length; i++) {
    if (cjCatalogGasTimeUp_()) break;
    var parent = parents[i];
    var r = cjRunCatalogOne_(parent, runId);
    if (r.skip) {
      skipN++;
      if (r.detail && String(r.detail).indexOf('err=') === 0) {
        // catalog_skip は cjRunCatalogOne_ 内で済
      } else if (r.detail && String(r.detail).indexOf('lock_') === 0) {
        cjAppendLog_(runId, 'catalog_skip', 0, 0, (parent + ' ' + r.detail).slice(0, 500));
      }
    } else if (r.ok) {
      doneN++;
      bands[r.band || '?'] = (bands[r.band || '?'] || 0) + 1;
    }
    Logger.log('[競合Catalog] %s/%s parent=%s %s', i + 1, parents.length, parent, r.detail);
  }
  var bandStr = Object.keys(bands)
    .map(function (k) {
      return k + '=' + bands[k];
    })
    .join(' ');
  var detail =
    'catalog_batch n=' +
    doneN +
    '/' +
    parents.length +
    ' skip=' +
    skipN +
    (healed ? ' heal=' + healed : '') +
    (phantomCleared ? ' phantom=' + phantomCleared : '') +
    (bandStr ? ' ' + bandStr : '');
  cjAppendLog_(runId, 'catalog_batch', doneN, skipN, detail);
  return { n: doneN, detail: detail, left: cjCatalogPendingList_(1).length > 0 };
}

function cjRunCatalogFromTrigger() {
  var runId = 'cj_cat_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var heldInflight = false;
  try {
    if (!cjCatalogEnabled_()) {
      cjClearCatalogTriggers_();
      PropertiesService.getScriptProperties().deleteProperty(CJ_CATALOG_EXPECT_FIRE_MS_PROP_);
      cjCatalogClearInflight_();
      cjAppendLog_(runId, 'catalog_off', 0, 0, 'trigger');
      return;
    }
    // 先頭で次を予約（途中死でも鎖が残る。①案P1と同思想）
    cjScheduleCatalogTrigger_(cjCatalogIntervalMin_(), 'start_of_run');
    var now = new Date().getTime();
    if (cjCatalogInflightUntil_() > now) {
      cjAppendLog_(runId, 'catalog_overlap', 0, 0, 'inflight_skip');
      return;
    }
    // 予算＋猶予。途中死でも重複起動を抑える
    cjCatalogSetInflight_(now + CJ_GAS_BUDGET_MS_ + cjCatalogStaleGraceMin_() * 60 * 1000);
    heldInflight = true;
    var out = cjRunCatalogBatch_(runId);
    if (out.left || out.n > 0) {
      cjScheduleCatalogTrigger_(cjCatalogIntervalMin_(), 'after_batch');
    } else {
      cjScheduleCatalogTrigger_(cjCatalogIntervalMin_(), 'idle_poll');
    }
  } catch (e) {
    cjAppendLog_(runId, 'catalog_fail', 0, 0, String((e && e.message) || e).slice(0, 200));
    try {
      cjMaybeArmCatalogRecovery_(runId, 'after_fail');
    } catch (e2) {}
    throw e;
  } finally {
    if (heldInflight) cjCatalogClearInflight_();
  }
}

function cjMenuCatalogNow() {
  if (!cjCatalogEnabled_()) {
    SpreadsheetApp.getUi().alert('CJ_CATALOG_ENABLED が OFF です。Property を true にしてから実行。');
    return;
  }
  var runId = 'cj_cat_menu_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var out = cjRunCatalogBatch_(runId);
  SpreadsheetApp.getUi().alert('Catalog列挙: ' + out.detail);
}

function cjMenuArmCatalogTrigger() {
  if (!cjCatalogEnabled_()) {
    SpreadsheetApp.getUi().alert('CJ_CATALOG_ENABLED が OFF です。');
    return;
  }
  cjScheduleCatalogTrigger_(1, 'menu_arm');
  SpreadsheetApp.getUi().alert('Catalog時計を予約しました（約1分後〜）。');
}

function cjMenuDisarmCatalogTrigger() {
  cjClearCatalogTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(CJ_CATALOG_EXPECT_FIRE_MS_PROP_);
  SpreadsheetApp.getUi().alert('Catalog時計を停止しました。');
}
