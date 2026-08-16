/**
 * 楽天・Yahoo 競合検索ベイクオフ A〜E（マスタ非書込）。
 * 結果: スプシ「競合検索ベイクオフ」へ書き、Agent が CSV 化する。
 */
var RAKUTEN_YAHOO_BAKEOFF_SHEET_ = '競合検索ベイクオフ';
var RAKUTEN_YAHOO_BAKEOFF_SS_PROD_ = '1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28';

function runRakutenYahooBakeoffToSheet() {
  var props = PropertiesService.getScriptProperties();
  var appId = (props.getProperty('RAKUTEN_APP_ID') || '').trim();
  var accessKey = (props.getProperty('RAKUTEN_ACCESS_KEY') || '').trim();
  var yahooId = (props.getProperty('YAHOO_SHOPPING_CLIENT_ID') || '').trim();
  var cases = rakutenYahooBakeoffCases_();
  var rows = [['mall', 'case_id', 'jan', 'maker', 'ai_name', 'query_id', 'query_label', 'query_text', 'api_param', 'hit_rank', 'hit_name', 'hit_price', 'hit_url', 'postage_or_ship', 'error', 'hit_count']];
  var ci;
  for (ci = 0; ci < cases.length; ci++) {
    var c = cases[ci];
    rows = rows.concat(rakutenYahooBakeoffRunMall_('rakuten', appId, accessKey, yahooId, c));
    rows = rows.concat(rakutenYahooBakeoffRunMall_('yahoo', appId, accessKey, yahooId, c));
  }
  var ss;
  try {
    ss = SpreadsheetApp.openById(RAKUTEN_YAHOO_BAKEOFF_SS_PROD_);
  } catch (eOpen) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sh = ss.getSheetByName(RAKUTEN_YAHOO_BAKEOFF_SHEET_);
  if (!sh) sh = ss.insertSheet(RAKUTEN_YAHOO_BAKEOFF_SHEET_);
  sh.clear();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  Logger.log('[ベイクオフ] rows=' + (rows.length - 1) + ' sheet=' + RAKUTEN_YAHOO_BAKEOFF_SHEET_);
  return { rows: rows.length - 1, sheet: RAKUTEN_YAHOO_BAKEOFF_SHEET_, ssId: ss.getId() };
}

function rakutenYahooBakeoffCases_() {
  var maker = '石原水産';
  return [
    { id: 'katsuobushi', jan: '4906283045119', maker: maker, ai_name: '食べるおだし（かつお）50ｇ', core: '石原水産 食べるおだし かつお 50g' },
    { id: 'maguro', jan: '4906283045317', maker: maker, ai_name: '食べるおだし（まぐろ）35ｇ', core: '石原水産 食べるおだし まぐろ 35g' },
    { id: 'buri', jan: '4906283047410', maker: maker, ai_name: '食べるおだし（ぶり）40ｇ', core: '石原水産 食べるおだし ぶり 40g' },
    { id: 'set3', jan: '4906283045119', maker: maker, ai_name: '食べるおだし（かつお・まぐろ・ぶり）', core: '石原水産 食べるおだし かつお まぐろ ぶり' }
  ];
}

function rakutenYahooBakeoffQueries_(c) {
  var jan = String(c.jan || '').trim();
  var name = String(c.ai_name || '').trim();
  var makerName = (String(c.maker || '').trim() + ' ' + name).trim();
  return [
    { id: 'A', label: '現行JAN', text: jan },
    { id: 'B', label: 'JANをキーワード', text: jan },
    { id: 'C', label: 'AI商品名', text: name },
    { id: 'D', label: 'メーカー+商品名', text: makerName },
    { id: 'E', label: '短い核', text: String(c.core || '').trim() }
  ];
}

function rakutenYahooBakeoffRunMall_(mall, appId, accessKey, yahooId, c) {
  var out = [];
  var qs = rakutenYahooBakeoffQueries_(c);
  var qi;
  for (qi = 0; qi < qs.length; qi++) {
    var q = qs[qi];
    var got = { error: '', items: [], param: '' };
    try {
      if (mall === 'rakuten') {
        got.param = 'keyword';
        if (!appId || !accessKey) {
          got.error = 'RAKUTEN_APP_ID/ACCESS_KEY missing';
        } else {
          Utilities.sleep(400);
          var rr = fetchRakutenIchibaItems(appId, accessKey, q.text, { hits: 10, page: 1 });
          got.error = rr.error || '';
          got.items = (rr.items || []).slice(0, 10).map(function (it) {
            return { name: it.itemName || '', price: it.itemPrice, url: it.itemUrl || '', extra: it.postageFlag };
          });
        }
      } else {
        if (!yahooId) {
          got.error = 'YAHOO_SHOPPING_CLIENT_ID missing';
        } else if (q.id === 'A') {
          got.param = 'jan_code';
          Utilities.sleep(200);
          var yj = fetchYahooShoppingItemsByJan(yahooId, q.text, 10);
          got.error = yj.error || '';
          got.items = (yj.hits || []).slice(0, 10).map(function (h) {
            return { name: h.name || '', price: h.price, url: h.url || '', extra: h.shippingCode };
          });
        } else {
          got.param = 'query';
          Utilities.sleep(200);
          var yq = fetchYahooShoppingItemsByQuery(yahooId, q.text, 10);
          got.error = yq.error || '';
          got.items = (yq.hits || []).slice(0, 10).map(function (h) {
            return { name: h.name || '', price: h.price, url: h.url || '', extra: h.shippingCode };
          });
        }
      }
    } catch (eRun) {
      got.error = eRun && eRun.message ? eRun.message : String(eRun);
    }
    if (!got.items.length) {
      out.push(rakutenYahooBakeoffRow_(mall, c, q, got.param, 0, '', '', '', '', got.error, 0));
    } else {
      var hi;
      for (hi = 0; hi < got.items.length; hi++) {
        var it = got.items[hi];
        out.push(rakutenYahooBakeoffRow_(mall, c, q, got.param, hi + 1, it.name, it.price, it.url, it.extra, got.error, got.items.length));
      }
    }
  }
  return out;
}

function rakutenYahooBakeoffRow_(mall, c, q, param, rank, name, price, url, extra, err, n) {
  return [
    mall, c.id, c.jan, c.maker, c.ai_name, q.id, q.label, q.text, param,
    rank, name, price != null ? price : '', url, extra != null ? extra : '', err || '', n
  ];
}
