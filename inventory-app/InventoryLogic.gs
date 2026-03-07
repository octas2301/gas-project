// ==========================================
// 在庫管理アプリ - 現在庫・平均原価・差異・アラート
// ==========================================

/**
 * 入荷登録後に商品マスタの平均原価を移動平均で更新する。
 * currentQtyOptional を渡すと getCurrentStockByJan を呼ばずにそれを使う（入荷APIからキャッシュで計算した値を渡して高速化）。
 */
function updateAverageCost(singleJAN, receivedQty, unitCost, currentQtyOptional) {
  var sheet = getSheetByName('商品マスタ');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iSingle = h.indexOf('単品JAN');
  var iAvg = h.indexOf('平均原価');
  var iCost = h.indexOf('商品原価');
  if (iSingle < 0 || iAvg < 0) return;
  var jan = (singleJAN || '').toString().trim();
  var qty = Number(receivedQty) || 0;
  var cost = Number(unitCost) || 0;
  if (!jan || qty <= 0) return;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iSingle] || '').toString().trim() !== jan) continue;
    var currentAvg = Number(data[r][iAvg]) || 0;
    var currentCost = Number(data[r][iCost]) || 0;
    if (currentAvg === 0) currentAvg = currentCost;
    var currentQty = (currentQtyOptional !== undefined && currentQtyOptional !== null && !isNaN(currentQtyOptional))
      ? Number(currentQtyOptional) : getCurrentStockByJan(jan);
    var newTotal = currentAvg * currentQty + cost * qty;
    var newQty = currentQty + qty;
    var newAvg = newQty > 0 ? newTotal / newQty : currentAvg;
    sheet.getRange(r + 1, iAvg + 1).setValue(Math.round(newAvg * 100) / 100);
    break;
  }
}

/**
 * 単品JAN ごとの現在庫（前回棚卸日の棚卸実数 + 確定後修正合計 + 翌日以降の入荷 - 翌日以降の出荷）。全社合計。
 * 前回棚卸日と同じ日の入荷・出荷は含めない（二重計上を防ぐ）。
 */
function getCurrentStockByJan(singleJAN) {
  var jan = (singleJAN || '').toString().trim();
  if (!jan) return 0;
  var lastCountDate = getLastCountDate();
  if (!lastCountDate) return 0;
  var countOnLastDate = getCountTotalOnDate(jan, lastCountDate);
  var adjustTotal = getCountAdjustTotal(jan, lastCountDate, null);
  var nextDayStart = new Date(lastCountDate.getFullYear(), lastCountDate.getMonth(), lastCountDate.getDate() + 1);
  var inQty = getReceivingTotalAfterDate(jan, nextDayStart);
  var outQty = getShippingTotalAfterDate(jan, nextDayStart);
  return countOnLastDate + adjustTotal + inQty - outQty;
}

/**
 * 指定日付の棚卸数量合計（その日の棚卸実数）。棚卸セッションがある場合はその日に完了したセッションのスキャンのみ。
 * 種別が「確定後修正」の行は除外する（通常スキャンまたは空の行のみ集計）。
 */
function getCountTotalOnDate(singleJAN, date) {
  var sheet = getSheetByName('棚卸スキャンリスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var iKind = h.indexOf('種別');
  var jan = (singleJAN || '').toString().trim();
  var dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
  var dayEnd = dayStart + 24 * 60 * 60 * 1000;
  var iSession = h.indexOf('セッションID');
  var completedOnDate = null;
  if (iSession >= 0) {
    var sessionSheet = getSheetByName('棚卸セッション');
    if (sessionSheet) {
      var sData = sessionSheet.getDataRange().getValues();
      if (sData.length >= 2) {
        var sh = sData[0];
        var siId = sh.indexOf('セッションID');
        var siEnd = sh.indexOf('完了日時');
        if (siId >= 0 && siEnd >= 0) {
          completedOnDate = {};
          for (var r = 1; r < sData.length; r++) {
            var endDt = sData[r][siEnd];
            if (!(endDt instanceof Date)) continue;
            var endDay = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate()).getTime();
            if (endDay >= dayStart && endDay < dayEnd) {
              completedOnDate[(sData[r][siId] || '').toString().trim()] = true;
            }
          }
        }
      }
    }
  }
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if (iKind >= 0 && (data[r][iKind] || '').toString().trim() === '確定後修正') continue;
    var d = data[r][iDate];
    if (!(d instanceof Date)) continue;
    var t = d.getTime();
    if (t < dayStart || t >= dayEnd) continue;
    if (iSession >= 0 && completedOnDate !== null) {
      var sid = (data[r][iSession] || '').toString().trim();
      if (sid && !completedOnDate[sid]) continue;
    }
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * 棚卸スキャンリストの最も新しい日付（日のみ）。棚卸セッションがあれば「最後に完了した日」を優先。
 */
function getLastCountDate() {
  var sessionSheet = getSheetByName('棚卸セッション');
  if (sessionSheet) {
    var sData = sessionSheet.getDataRange().getValues();
    if (sData.length >= 2) {
      var sh = sData[0];
      var iEnd = sh.indexOf('完了日時');
      if (iEnd >= 0) {
        var maxCompleted = null;
        for (var r = 1; r < sData.length; r++) {
          var endDt = sData[r][iEnd];
          if (endDt instanceof Date) {
            var day = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate());
            if (!maxCompleted || day > maxCompleted) maxCompleted = day;
          }
        }
        if (maxCompleted) return maxCompleted;
      }
    }
  }
  var sheet = getSheetByName('棚卸スキャンリスト');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var h = data[0];
  var iDate = h.indexOf('日時');
  if (iDate < 0) return null;
  var maxDate = null;
  for (var r = 1; r < data.length; r++) {
    var d = data[r][iDate];
    if (d instanceof Date) {
      var day = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      if (!maxDate || day > maxDate) maxDate = day;
    }
  }
  return maxDate;
}

/**
 * 指定日時 asOf 時点で「最後に完了した棚卸日」を返す。棚卸セッションの完了日時が asOf 以前のもののみ対象。
 */
function getLastCountDateAsOf(asOfDateTime) {
  var sessionSheet = getSheetByName('棚卸セッション');
  if (!sessionSheet) return null;
  var data = sessionSheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var h = data[0];
  var iEnd = h.indexOf('完了日時');
  if (iEnd < 0) return null;
  var asOfTime = (asOfDateTime instanceof Date) ? asOfDateTime.getTime() : new Date(asOfDateTime).getTime();
  var maxDay = null;
  for (var r = 1; r < data.length; r++) {
    var endDt = data[r][iEnd];
    if (!(endDt instanceof Date)) continue;
    if (endDt.getTime() > asOfTime) continue;
    var day = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate());
    if (!maxDay || day > maxDay) maxDay = day;
  }
  return maxDay;
}

/**
 * 指定日付の棚卸数量合計。棚卸セッションを使う場合、その日に完了したセッションに属するスキャンのみ集計。
 * asOfDateTime 省略時は getCountTotalOnDate と同じ（全スキャンで日付一致）。
 */
function getCountTotalOnDateAsOf(singleJAN, date, asOfDateTime) {
  var sheet = getSheetByName('棚卸スキャンリスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
  var dayEnd = dayStart + 24 * 60 * 60 * 1000;
  var iSession = h.indexOf('セッションID');
  var completedSessionsOnDate = null;
  if (asOfDateTime != null && iSession >= 0) {
    var sessionSheet = getSheetByName('棚卸セッション');
    if (sessionSheet) {
      var sData = sessionSheet.getDataRange().getValues();
      if (sData.length >= 2) {
        var sh = sData[0];
        var siId = sh.indexOf('セッションID');
        var siEnd = sh.indexOf('完了日時');
        if (siId >= 0 && siEnd >= 0) {
          completedSessionsOnDate = {};
          var asOfTime = (asOfDateTime instanceof Date) ? asOfDateTime.getTime() : new Date(asOfDateTime).getTime();
          for (var r = 1; r < sData.length; r++) {
            var endDt = sData[r][siEnd];
            if (!(endDt instanceof Date) || endDt.getTime() > asOfTime) continue;
            var endDay = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate()).getTime();
            if (endDay >= dayStart && endDay < dayEnd) {
              completedSessionsOnDate[(sData[r][siId] || '').toString().trim()] = true;
            }
          }
        }
      }
    }
  }
  var iKindAsOf = h.indexOf('種別');
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if (iKindAsOf >= 0 && (data[r][iKindAsOf] || '').toString().trim() === '確定後修正') continue;
    var d = data[r][iDate];
    if (!(d instanceof Date)) continue;
    var t = d.getTime();
    if (t < dayStart || t >= dayEnd) continue;
    if (asOfDateTime != null && iSession >= 0) {
      var sid = (data[r][iSession] || '').toString().trim();
      if (sid && completedSessionsOnDate && !completedSessionsOnDate[sid]) continue;
    }
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * 入荷リストで、日時が afterDt より後かつ beforeDt 以前の数量合計。
 */
function getReceivingTotalBetween(singleJAN, afterDt, beforeDt) {
  var sheet = getSheetByName('入荷リスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var afterTime = (afterDt instanceof Date) ? afterDt.getTime() : new Date(afterDt).getTime();
  var beforeTime = (beforeDt instanceof Date) ? beforeDt.getTime() : new Date(beforeDt).getTime();
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    var d = data[r][iDate];
    if (d instanceof Date) {
      var t = d.getTime();
      if (t > afterTime && t <= beforeTime) total += Number(data[r][iQty]) || 0;
    }
  }
  return total;
}

/**
 * 出荷リストで、日時が afterDt より後かつ beforeDt 以前の数量合計。
 */
function getShippingTotalBetween(singleJAN, afterDt, beforeDt) {
  var sheet = getSheetByName('出荷リスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var afterTime = (afterDt instanceof Date) ? afterDt.getTime() : new Date(afterDt).getTime();
  var beforeTime = (beforeDt instanceof Date) ? beforeDt.getTime() : new Date(beforeDt).getTime();
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    var d = data[r][iDate];
    if (d instanceof Date) {
      var t = d.getTime();
      if (t > afterTime && t <= beforeTime) total += Number(data[r][iQty]) || 0;
    }
  }
  return total;
}

/**
 * 指定日時 asOf 時点での理論在庫（JAN別）。スタート時点理論の計算に使用。
 * 確定後修正（lastCountDate〜asOfDateTime の範囲）も反映する。
 */
function getTheoryAsOf(singleJAN, asOfDateTime) {
  var jan = (singleJAN || '').toString().trim();
  if (!jan) return 0;
  var lastCountDate = getLastCountDateAsOf(asOfDateTime);
  if (!lastCountDate) return 0;
  var countOnLastDate = getCountTotalOnDateAsOf(jan, lastCountDate, asOfDateTime);
  var asOfDt = (asOfDateTime instanceof Date) ? asOfDateTime : new Date(asOfDateTime);
  var adjustTotal = getCountAdjustTotal(jan, lastCountDate, asOfDt);
  var nextDayStart = new Date(lastCountDate.getFullYear(), lastCountDate.getMonth(), lastCountDate.getDate() + 1);
  var inQty = 0;
  var outQty = 0;
  var asOfTime = (asOfDateTime instanceof Date) ? asOfDateTime.getTime() : new Date(asOfDateTime).getTime();
  var recSheet = getSheetByName('入荷リスト');
  if (recSheet) {
    var rData = recSheet.getDataRange().getValues();
    var rh = rData[0];
    var riDate = rh.indexOf('日時');
    var riJan = rh.indexOf('単品JAN');
    var riQty = rh.indexOf('数量');
    if (riDate >= 0 && riJan >= 0 && riQty >= 0) {
      var nextTime = nextDayStart.getTime();
      for (var r = 1; r < rData.length; r++) {
        if ((rData[r][riJan] || '').toString().trim() !== jan) continue;
        var rd = rData[r][riDate];
        if (rd instanceof Date) {
          var rt = rd.getTime();
          if (rt > nextTime && rt <= asOfTime) inQty += Number(rData[r][riQty]) || 0;
        }
      }
    }
  }
  var shipSheet = getSheetByName('出荷リスト');
  if (shipSheet) {
    var sData = shipSheet.getDataRange().getValues();
    var sh = sData[0];
    var siDate = sh.indexOf('日時');
    var siJan = sh.indexOf('単品JAN');
    var siQty = sh.indexOf('数量');
    if (siDate >= 0 && siJan >= 0 && siQty >= 0) {
      var nextTime = nextDayStart.getTime();
      for (var r = 1; r < sData.length; r++) {
        if ((sData[r][siJan] || '').toString().trim() !== jan) continue;
        var sd = sData[r][siDate];
        if (sd instanceof Date) {
          var st = sd.getTime();
          if (st > nextTime && st <= asOfTime) outQty += Number(sData[r][siQty]) || 0;
        }
      }
    }
  }
  return countOnLastDate + adjustTotal + inQty - outQty;
}

/**
 * 指定セッションに属するスキャン数量合計（JAN別）。棚卸スキャンリストのセッションID列で絞る。
 * セッションID列がない・または sessionId が空の場合は 0 を返す（期間外の従来集計は getCountTotalThisMonth を使用）。
 */
function getCountTotalForSession(singleJAN, sessionId) {
  var sid = (sessionId || '').toString().trim();
  if (!sid) return 0;
  var sheet = getSheetByName('棚卸スキャンリスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iSession = h.indexOf('セッションID');
  if (iSession < 0 && data[0] && data[0].length >= 14) iSession = 13;
  if (iJan < 0 || iQty < 0) return 0;
  if (iSession < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    var rowSession = (data[r][iSession] || '').toString().trim();
    if (rowSession !== sid) continue;
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * 棚卸スキャンリストの「確定後修正」行の数量合計。
 * fromDate: 対象開始日時（inclusive、null = 制限なし）。
 * toDate: 対象終了日時（inclusive、null = 制限なし）。
 * 数量は加算=正・減算=負で記録されているため、合計がそのまま差分値になる。
 */
function getCountAdjustTotal(singleJAN, fromDate, toDate) {
  var sheet = getSheetByName('棚卸スキャンリスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iKind = h.indexOf('種別');
  if (iDate < 0 || iJan < 0 || iQty < 0 || iKind < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var fromTime = fromDate ? fromDate.getTime() : -Infinity;
  var toTime = toDate ? toDate.getTime() : Infinity;
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if ((data[r][iKind] || '').toString().trim() !== '確定後修正') continue;
    var d = data[r][iDate];
    if (!(d instanceof Date)) continue;
    var t = d.getTime();
    if (t < fromTime || t > toTime) continue;
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

function getReceivingTotalAfterDate(singleJAN, afterDate) {
  var sheet = getSheetByName('入荷リスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var total = 0;
  var afterTime = afterDate.getTime();
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    var d = data[r][iDate];
    if (d instanceof Date && d.getTime() > afterTime) {
      total += Number(data[r][iQty]) || 0;
    }
  }
  return total;
}

function getShippingTotalAfterDate(singleJAN, afterDate) {
  var sheet = getSheetByName('出荷リスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var total = 0;
  var afterTime = afterDate.getTime();
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    var d = data[r][iDate];
    if (d instanceof Date && d.getTime() > afterTime) {
      total += Number(data[r][iQty]) || 0;
    }
  }
  return total;
}

// --- バッチ計算用：シートを1回ずつ読んで全JANの現在庫・棚卸今月分を一括計算 ---

/**
 * 棚卸・入荷・出荷シートのデータを1回ずつ読み、キャッシュオブジェクトで返す。
 */
function getInventorySheetCache() {
  var cache = {};
  var sessionSheet = getSheetByName('棚卸セッション');
  cache.sessionData = sessionSheet ? sessionSheet.getDataRange().getValues() : [];
  var countSheet = getSheetByName('棚卸スキャンリスト');
  cache.countData = countSheet ? countSheet.getDataRange().getValues() : [];
  var recSheet = getSheetByName('入荷リスト');
  cache.receivingData = recSheet ? recSheet.getDataRange().getValues() : [];
  var shipSheet = getSheetByName('出荷リスト');
  cache.shippingData = shipSheet ? shipSheet.getDataRange().getValues() : [];
  return cache;
}

/**
 * キャッシュから「最後に完了した棚卸日」を算出する。
 */
function getLastCountDateFromCache(cache) {
  if (cache.sessionData && cache.sessionData.length >= 2) {
    var sh = cache.sessionData[0];
    var iEnd = sh.indexOf('完了日時');
    if (iEnd >= 0) {
      var maxCompleted = null;
      for (var r = 1; r < cache.sessionData.length; r++) {
        var endDt = cache.sessionData[r][iEnd];
        if (endDt instanceof Date) {
          var day = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate());
          if (!maxCompleted || day > maxCompleted) maxCompleted = day;
        }
      }
      if (maxCompleted) return maxCompleted;
    }
  }
  if (!cache.countData || cache.countData.length < 2) return null;
  var h = cache.countData[0];
  var iDate = h.indexOf('日時');
  if (iDate < 0) return null;
  var maxDate = null;
  for (var r = 1; r < cache.countData.length; r++) {
    var d = cache.countData[r][iDate];
    if (d instanceof Date) {
      var day = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      if (!maxDate || day > maxDate) maxDate = day;
    }
  }
  return maxDate;
}

/**
 * キャッシュを使って指定日付の棚卸数量合計を計算。
 */
function getCountTotalOnDateFromCache(singleJAN, date, cache) {
  var data = cache.countData;
  if (!data || data.length < 2) return 0;
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var iKind = h.indexOf('種別');
  var jan = (singleJAN || '').toString().trim();
  var dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
  var dayEnd = dayStart + 24 * 60 * 60 * 1000;
  var iSession = h.indexOf('セッションID');
  var completedOnDate = null;
  if (iSession >= 0 && cache.sessionData && cache.sessionData.length >= 2) {
    var sh = cache.sessionData[0];
    var siId = sh.indexOf('セッションID');
    var siEnd = sh.indexOf('完了日時');
    if (siId >= 0 && siEnd >= 0) {
      completedOnDate = {};
      for (var r = 1; r < cache.sessionData.length; r++) {
        var endDt = cache.sessionData[r][siEnd];
        if (!(endDt instanceof Date)) continue;
        var endDay = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate()).getTime();
        if (endDay >= dayStart && endDay < dayEnd) {
          completedOnDate[(cache.sessionData[r][siId] || '').toString().trim()] = true;
        }
      }
    }
  }
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if (iKind >= 0 && (data[r][iKind] || '').toString().trim() === '確定後修正') continue;
    var d = data[r][iDate];
    if (!(d instanceof Date)) continue;
    var t = d.getTime();
    if (t < dayStart || t >= dayEnd) continue;
    if (iSession >= 0 && completedOnDate !== null) {
      var sid = (data[r][iSession] || '').toString().trim();
      if (sid && !completedOnDate[sid]) continue;
    }
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * キャッシュを使って確定後修正の数量合計を計算。
 */
function getCountAdjustTotalFromCache(singleJAN, fromDate, toDate, cache) {
  var data = cache.countData;
  if (!data || data.length < 2) return 0;
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iKind = h.indexOf('種別');
  if (iDate < 0 || iJan < 0 || iQty < 0 || iKind < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var fromTime = fromDate ? fromDate.getTime() : -Infinity;
  var toTime = toDate ? toDate.getTime() : Infinity;
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if ((data[r][iKind] || '').toString().trim() !== '確定後修正') continue;
    var d = data[r][iDate];
    if (!(d instanceof Date)) continue;
    var t = d.getTime();
    if (t < fromTime || t > toTime) continue;
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * キャッシュを使って翌日以降の入荷数量合計を計算。
 */
function getReceivingTotalAfterDateFromCache(singleJAN, afterDate, cache) {
  var data = cache.receivingData;
  if (!data || data.length < 2) return 0;
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var afterTime = afterDate.getTime();
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    var d = data[r][iDate];
    if (d instanceof Date && d.getTime() > afterTime) total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * キャッシュを使って翌日以降の出荷数量合計を計算。
 */
function getShippingTotalAfterDateFromCache(singleJAN, afterDate, cache) {
  var data = cache.shippingData;
  if (!data || data.length < 2) return 0;
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var afterTime = afterDate.getTime();
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    var d = data[r][iDate];
    if (d instanceof Date && d.getTime() > afterTime) total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * キャッシュから1JANの現在庫を計算。
 */
function getCurrentStockByJanFromCache(jan, cache, lastCountDate) {
  if (!jan || !lastCountDate) return 0;
  var countOnLastDate = getCountTotalOnDateFromCache(jan, lastCountDate, cache);
  var adjustTotal = getCountAdjustTotalFromCache(jan, lastCountDate, null, cache);
  var nextDayStart = new Date(lastCountDate.getFullYear(), lastCountDate.getMonth(), lastCountDate.getDate() + 1);
  var inQty = getReceivingTotalAfterDateFromCache(jan, nextDayStart, cache);
  var outQty = getShippingTotalAfterDateFromCache(jan, nextDayStart, cache);
  return countOnLastDate + adjustTotal + inQty - outQty;
}

/**
 * キャッシュから今月の棚卸数量合計を計算（1JAN）。
 */
function getCountTotalThisMonthFromCache(singleJAN, yyyymm, countData) {
  if (!countData || countData.length < 2) return 0;
  var h = countData[0];
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iYm = h.indexOf('年月');
  var iKind = h.indexOf('種別');
  if (iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var total = 0;
  for (var r = 1; r < countData.length; r++) {
    if ((countData[r][iJan] || '').toString().trim() !== jan) continue;
    if (iKind >= 0 && (countData[r][iKind] || '').toString().trim() === '確定後修正') continue;
    if (iYm >= 0) {
      var cellYm = countData[r][iYm];
      var cellYmStr = (cellYm instanceof Date) ? Utilities.formatDate(cellYm, 'Asia/Tokyo', 'yyyy/MM') : (cellYm || '').toString().trim();
      if (cellYmStr !== yyyymm) continue;
    }
    total += Number(countData[r][iQty]) || 0;
  }
  return total;
}

/**
 * 複数JANの現在庫を一括計算（シートは各1回のみ読み）。返り値: { jan: 在庫数, ... }
 */
function getCurrentStockBatch(janList) {
  var result = {};
  if (!janList || janList.length === 0) return result;
  var cache = getInventorySheetCache();
  var lastCountDate = getLastCountDateFromCache(cache);
  if (!lastCountDate) return result;
  for (var i = 0; i < janList.length; i++) {
    var jan = (janList[i] || '').toString().trim();
    if (!jan) continue;
    result[jan] = getCurrentStockByJanFromCache(jan, cache, lastCountDate);
  }
  return result;
}

/**
 * 複数JANの「今月棚卸数量」を一括計算（countData はキャッシュの countData を渡す）。返り値: { jan: 数量, ... }
 */
function getCountTotalThisMonthBatch(janList, yyyymm, countData) {
  var result = {};
  if (!janList || janList.length === 0 || !countData) return result;
  for (var i = 0; i < janList.length; i++) {
    var jan = (janList[i] || '').toString().trim();
    if (!jan) continue;
    result[jan] = getCountTotalThisMonthFromCache(jan, yyyymm, countData);
  }
  return result;
}

// --- 場所別の現在庫・棚卸数量（差異・在庫の場所別表示用）---

/**
 * キャッシュから指定ロケーションの「指定日付の棚卸数量」を計算。
 */
function getCountTotalOnDateByLocationFromCache(singleJAN, date, location, cache) {
  var data = cache.countData;
  if (!data || data.length < 2) return 0;
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iLoc = h.indexOf('ロケーション');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var iKind = h.indexOf('種別');
  var jan = (singleJAN || '').toString().trim();
  var locStr = (location || '').toString().trim();
  var dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
  var dayEnd = dayStart + 24 * 60 * 60 * 1000;
  var iSession = h.indexOf('セッションID');
  var completedOnDate = null;
  if (iSession >= 0 && cache.sessionData && cache.sessionData.length >= 2) {
    var sh = cache.sessionData[0];
    var siId = sh.indexOf('セッションID');
    var siEnd = sh.indexOf('完了日時');
    if (siId >= 0 && siEnd >= 0) {
      completedOnDate = {};
      for (var r = 1; r < cache.sessionData.length; r++) {
        var endDt = cache.sessionData[r][siEnd];
        if (!(endDt instanceof Date)) continue;
        var endDay = new Date(endDt.getFullYear(), endDt.getMonth(), endDt.getDate()).getTime();
        if (endDay >= dayStart && endDay < dayEnd) {
          completedOnDate[(cache.sessionData[r][siId] || '').toString().trim()] = true;
        }
      }
    }
  }
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if (iLoc >= 0 && (data[r][iLoc] || '').toString().trim() !== locStr) continue;
    if (iKind >= 0 && (data[r][iKind] || '').toString().trim() === '確定後修正') continue;
    var d = data[r][iDate];
    if (!(d instanceof Date)) continue;
    var t = d.getTime();
    if (t < dayStart || t >= dayEnd) continue;
    if (iSession >= 0 && completedOnDate !== null) {
      var sid = (data[r][iSession] || '').toString().trim();
      if (sid && !completedOnDate[sid]) continue;
    }
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * キャッシュから指定ロケーションの確定後修正数量を計算。
 */
function getCountAdjustTotalByLocationFromCache(singleJAN, fromDate, toDate, location, cache) {
  var data = cache.countData;
  if (!data || data.length < 2) return 0;
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iKind = h.indexOf('種別');
  var iLoc = h.indexOf('ロケーション');
  if (iDate < 0 || iJan < 0 || iQty < 0 || iKind < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var locStr = (location || '').toString().trim();
  var fromTime = fromDate ? fromDate.getTime() : -Infinity;
  var toTime = toDate ? toDate.getTime() : Infinity;
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if (iLoc >= 0 && (data[r][iLoc] || '').toString().trim() !== locStr) continue;
    if ((data[r][iKind] || '').toString().trim() !== '確定後修正') continue;
    var d = data[r][iDate];
    if (!(d instanceof Date)) continue;
    var t = d.getTime();
    if (t < fromTime || t > toTime) continue;
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * キャッシュから指定ロケーションの翌日以降入荷数量を計算。
 */
function getReceivingTotalAfterDateByLocationFromCache(singleJAN, afterDate, location, cache) {
  var data = cache.receivingData;
  if (!data || data.length < 2) return 0;
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iLoc = h.indexOf('ロケーション');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var locStr = (location || '').toString().trim();
  var afterTime = afterDate.getTime();
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if (iLoc >= 0 && (data[r][iLoc] || '').toString().trim() !== locStr) continue;
    var d = data[r][iDate];
    if (d instanceof Date && d.getTime() > afterTime) total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * キャッシュから指定出荷元（ロケーション）の翌日以降出荷数量を計算。
 */
function getShippingTotalAfterDateFromLocationFromCache(singleJAN, afterDate, fromLocation, cache) {
  var data = cache.shippingData;
  if (!data || data.length < 2) return 0;
  var h = data[0];
  var iDate = h.indexOf('日時');
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iFrom = h.indexOf('出荷元');
  if (iDate < 0 || iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var locStr = (fromLocation || '').toString().trim();
  var afterTime = afterDate.getTime();
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if (iFrom >= 0 && (data[r][iFrom] || '').toString().trim() !== locStr) continue;
    var d = data[r][iDate];
    if (d instanceof Date && d.getTime() > afterTime) total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * キャッシュから1JAN・1ロケーションの現在庫を計算。
 */
function getCurrentStockByJanFromCacheByLocation(jan, cache, lastCountDate, location) {
  if (!jan || !lastCountDate) return 0;
  var countOnLast = getCountTotalOnDateByLocationFromCache(jan, lastCountDate, location, cache);
  var adjustTotal = getCountAdjustTotalByLocationFromCache(jan, lastCountDate, null, location, cache);
  var nextDayStart = new Date(lastCountDate.getFullYear(), lastCountDate.getMonth(), lastCountDate.getDate() + 1);
  var inQty = getReceivingTotalAfterDateByLocationFromCache(jan, nextDayStart, location, cache);
  var outQty = getShippingTotalAfterDateFromLocationFromCache(jan, nextDayStart, location, cache);
  return countOnLast + adjustTotal + inQty - outQty;
}

/**
 * キャッシュから今月の棚卸数量をロケーション別に計算（1JAN）。
 */
function getCountTotalThisMonthByLocationFromCache(singleJAN, yyyymm, location, countData) {
  if (!countData || countData.length < 2) return 0;
  var h = countData[0];
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iYm = h.indexOf('年月');
  var iKind = h.indexOf('種別');
  var iLoc = h.indexOf('ロケーション');
  if (iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var locStr = (location || '').toString().trim();
  var total = 0;
  for (var r = 1; r < countData.length; r++) {
    if ((countData[r][iJan] || '').toString().trim() !== jan) continue;
    if (iLoc >= 0 && (countData[r][iLoc] || '').toString().trim() !== locStr) continue;
    if (iKind >= 0 && (countData[r][iKind] || '').toString().trim() === '確定後修正') continue;
    if (iYm >= 0) {
      var cellYm = countData[r][iYm];
      var cellYmStr = (cellYm instanceof Date) ? Utilities.formatDate(cellYm, 'Asia/Tokyo', 'yyyy/MM') : (cellYm || '').toString().trim();
      if (cellYmStr !== yyyymm) continue;
    }
    total += Number(countData[r][iQty]) || 0;
  }
  return total;
}

/**
 * 複数JANの現在庫をロケーション別に一括計算。返り値: { jan: { location: qty, ... }, ... }
 */
function getCurrentStockBatchByLocation(janList, locationList, cache, lastCountDate) {
  var result = {};
  if (!janList || janList.length === 0 || !locationList || locationList.length === 0) return result;
  if (!cache || !lastCountDate) return result;
  for (var i = 0; i < janList.length; i++) {
    var jan = (janList[i] || '').toString().trim();
    if (!jan) continue;
    result[jan] = {};
    for (var L = 0; L < locationList.length; L++) {
      var loc = (locationList[L] || '').toString().trim();
      result[jan][loc] = getCurrentStockByJanFromCacheByLocation(jan, cache, lastCountDate, loc);
    }
  }
  return result;
}

/**
 * 複数JANの今月棚卸数量をロケーション別に一括計算。返り値: { jan: { location: qty, ... }, ... }
 */
function getCountTotalThisMonthBatchByLocation(janList, yyyymm, countData, locationList) {
  var result = {};
  if (!janList || janList.length === 0 || !countData || !locationList || locationList.length === 0) return result;
  for (var i = 0; i < janList.length; i++) {
    var jan = (janList[i] || '').toString().trim();
    if (!jan) continue;
    result[jan] = {};
    for (var L = 0; L < locationList.length; L++) {
      var loc = (locationList[L] || '').toString().trim();
      result[jan][loc] = getCountTotalThisMonthByLocationFromCache(jan, yyyymm, loc, countData);
    }
  }
  return result;
}

/**
 * 差異一覧。未完了の棚卸セッションがあれば「期間中」（入出荷を加味した理論で突合）を返す。
 * 返り値: { list: [{ singleJAN, productName, theoryStock, countQty, variance }, ...], sessionActive: true/false }
 */
function getVarianceList() {
  var active = getActiveCountSession();
  if (active && active.sessionId && active.startTime) {
    var sessionList = getVarianceListDuringSession(active.sessionId, active.startTime);
    return { list: sessionList, sessionActive: true };
  }
  var productSheet = getSheetByName('商品マスタ');
  if (!productSheet) return { list: [], sessionActive: false };
  var data = productSheet.getDataRange().getValues();
  var h = data[0];
  var iSingle = h.indexOf('単品JAN');
  var iName = h.indexOf('商品名');
  if (iSingle < 0) return { list: [], sessionActive: false };
  var janList = [];
  for (var r = 1; r < data.length; r++) {
    var j = (data[r][iSingle] || '').toString().trim();
    if (j) janList.push(j);
  }
  var cache = getInventorySheetCache();
  var lastCountDate = getLastCountDateFromCache(cache);
  var stockBatch = getCurrentStockBatch(janList);
  var thisMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');
  var countBatch = getCountTotalThisMonthBatch(janList, thisMonth, cache.countData);
  var locationList = getMasterList('場所マスタ') || [];
  var stockBatchByLoc = getCurrentStockBatchByLocation(janList, locationList, cache, lastCountDate);
  var countBatchByLoc = getCountTotalThisMonthBatchByLocation(janList, thisMonth, cache.countData, locationList);
  var list = [];
  for (var r = 1; r < data.length; r++) {
    var jan = (data[r][iSingle] || '').toString().trim();
    if (!jan) continue;
    var theory = (stockBatch[jan] !== undefined) ? stockBatch[jan] : getCurrentStockByJan(jan);
    var countQty = (countBatch[jan] !== undefined) ? countBatch[jan] : getCountTotalThisMonth(jan, thisMonth);
    var variance = countQty - theory;
    var byLocation = [];
    for (var L = 0; L < locationList.length; L++) {
      var loc = (locationList[L] || '').toString().trim();
      var tLoc = (stockBatchByLoc[jan] && stockBatchByLoc[jan][loc] !== undefined) ? stockBatchByLoc[jan][loc] : 0;
      var cLoc = (countBatchByLoc[jan] && countBatchByLoc[jan][loc] !== undefined) ? countBatchByLoc[jan][loc] : 0;
      byLocation.push({ location: loc, theory: tLoc, count: cLoc, variance: cLoc - tLoc });
    }
    list.push({
      singleJAN: jan,
      productName: (data[r][iName] || '').toString().trim(),
      theoryStock: theory,
      countQty: countQty,
      variance: variance,
      byLocation: byLocation
    });
  }
  return { list: list, sessionActive: false, locationList: locationList };
}

/**
 * 棚卸期間中の差異一覧。理論在庫（期間中）＝スタート時点理論＋期間中入荷－期間中出荷。差異＝棚卸_期間中－理論_期間中。
 */
function getVarianceListDuringSession(sessionId, startTime) {
  var productSheet = getSheetByName('商品マスタ');
  if (!productSheet) return [];
  var data = productSheet.getDataRange().getValues();
  var h = data[0];
  var iSingle = h.indexOf('単品JAN');
  var iName = h.indexOf('商品名');
  if (iSingle < 0) return [];
  var now = new Date();
  var list = [];
  for (var r = 1; r < data.length; r++) {
    var jan = (data[r][iSingle] || '').toString().trim();
    if (!jan) continue;
    var theoryStart = getTheoryAsOf(jan, startTime);
    var inDuring = getReceivingTotalBetween(jan, startTime, now);
    var outDuring = getShippingTotalBetween(jan, startTime, now);
    var theoryDuring = theoryStart + inDuring - outDuring;
    var countDuring = getCountTotalForSession(jan, sessionId);
    var variance = countDuring - theoryDuring;
    list.push({
      singleJAN: jan,
      productName: (data[r][iName] || '').toString().trim(),
      theoryStock: theoryDuring,
      countQty: countDuring,
      variance: variance
    });
  }
  return list;
}

function getCountTotalThisMonth(singleJAN, yyyymm) {
  var sheet = getSheetByName('棚卸スキャンリスト');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var iJan = h.indexOf('単品JAN');
  var iQty = h.indexOf('数量');
  var iYm = h.indexOf('年月');
  var iKind = h.indexOf('種別');
  if (iJan < 0 || iQty < 0) return 0;
  var jan = (singleJAN || '').toString().trim();
  var total = 0;
  for (var r = 1; r < data.length; r++) {
    if ((data[r][iJan] || '').toString().trim() !== jan) continue;
    if (iKind >= 0 && (data[r][iKind] || '').toString().trim() === '確定後修正') continue;
    if (iYm >= 0) {
      var cellYm = data[r][iYm];
      var cellYmStr = (cellYm instanceof Date) ? Utilities.formatDate(cellYm, 'Asia/Tokyo', 'yyyy/MM') : (cellYm || '').toString().trim();
      if (cellYmStr !== yyyymm) continue;
    }
    total += Number(data[r][iQty]) || 0;
  }
  return total;
}

/**
 * 棚卸アラート月数を設定から取得。デフォルト3。
 */
function getStockAlertMonthCount() {
  var v = getSettingValue('棚卸アラート月数');
  var n = parseInt(v, 10);
  return isNaN(n) || n < 1 ? 3 : n;
}
