/**
 * overstock-dashboard.gs
 *
 * Populates the "Ecom_short_expiry_overstock" tab of the war-room-report Google Sheet.
 *
 * COMPUTATION: snapshot-delta approach — daily cleared AED = yesterday's overstock AED minus
 * today's overstock AED (both read from col DI, filtered to Order Interval = "Christmas/Easter").
 *
 * Source file: "Active SKUs – Short Expiry, overstock, re-order"
 *   File ID : 12TOeabCNl2YUImUCoyDYOmpi-_CCCFos2n6gukLNEm4
 *   Tab format: one tab per day, named DD.MM.YY (e.g. "23.04.26")
 *
 * Target file: war-room-report
 *   File ID : 1Z3Sn4_zPV_VSYSPjZgMEoyYBnzrHpEI35xJB8XAQYu0
 *   Tab name : Ecom_Short_Expiry_Overstock_daily
 */

// ─── CONFIGURATION ────────────────────────────────────────────────────────────

var CFG = {
  SOURCE_FILE_ID : '12TOeabCNl2YUImUCoyDYOmpi-_CCCFos2n6gukLNEm4',
  TARGET_FILE_ID : '1Z3Sn4_zPV_VSYSPjZgMEoyYBnzrHpEI35xJB8XAQYu0',
  TARGET_TAB     : 'Ecom_Short_Expiry_Overstock_daily',  // exact capitalisation required
  NETLIFY_URL    : 'https://peppy-bubblegum-32ada2.netlify.app/',
  BASELINE_DATE  : new Date(2026, 3, 14),  // 14 Apr 2026 (month is 0-indexed)
  DEADLINE_DATE  : new Date(2026, 4, 31),  // 31 May 2026
  FILTER_VALUE   : 'Christmas/Easter',
  STALE_THRESHOLD: 2,                      // flag if latest tab is >2 days behind today
};

// Column header names — matched case-insensitively after normalisation.
// These are the raw values as they appear in the Google Sheet (no markdown escaping).
var HEADERS = {
  ITEM_CODE       : 'item code',
  DESCRIPTION     : 'description',
  SUPPLIER        : 'supplier name',
  SOH             : 'sellable on hand',
  ORDER_INTERVAL  : 'order interval',
  ITEM_COST       : 'item cost',
  OVERSTOCK_VALUE : 'overstock value based on item cost (soh+ in transit)>120 days',
  OVERSTOCK_QTY   : 'overstock qty (considering soh + in transit)qty>120 days',
};

// ─── MENU & TRIGGER ENTRY POINTS ─────────────────────────────────────────────

/**
 * Adds "Casinetto > Refresh overstock dashboard" to the spreadsheet menu.
 * Attach this script to the TARGET file (war-room-report) in the Apps Script editor.
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Casinetto', [
    { name: 'Refresh overstock dashboard', functionName: 'runDashboard' },
  ]);
}

/**
 * Main entry point.
 * Can be called from the Casinetto menu or from a time-based trigger.
 */
function runDashboard() {
  var runStart = new Date();
  var today    = _midnightToday();

  try {
    var src = SpreadsheetApp.openById(CFG.SOURCE_FILE_ID);
    var tgt = SpreadsheetApp.openById(CFG.TARGET_FILE_ID);

    // Safety: require the target tab to already exist with the exact name.
    // Do NOT create it automatically — prevents silent writes to wrong tabs.
    var out = tgt.getSheetByName(CFG.TARGET_TAB);
    if (!out) {
      console.error(
        '[Overstock Dashboard] ABORT — target tab "' + CFG.TARGET_TAB +
        '" not found in file ' + CFG.TARGET_FILE_ID +
        '. Create the tab manually and re-run.'
      );
      return;
    }

    // ── Locate today's tab (with date-descending fallback) ───────────────────
    var todayResult = _findMostRecentTab(src, today);
    var todaySheet  = todayResult.sheet;
    var todayDate   = todayResult.date;
    var isStale     = todayResult.isStale;

    if (!todaySheet) {
      console.error('[Overstock Dashboard] CRITICAL: no dated tabs found in source file.');
      return;
    }

    // ── Baseline tab (14 Apr 2026) ────────────────────────────────────────────
    var baselineSheet = _getTabByDate(src, CFG.BASELINE_DATE);

    // ── Read today's filtered rows ────────────────────────────────────────────
    var todayColInfo = _columnIndices(todaySheet);
    var todayRows    = _filteredRows(todaySheet, todayColInfo);
    var todayAED   = _sum(todayRows, 'overstockValue');
    var todayUnits = _sum(todayRows, 'overstockQty');
    var skuCount   = todayRows.length;

    // ── Stackdriver audit log ─────────────────────────────────────────────────
    console.log(
      '[Overstock Dashboard] Run started at ' + runStart.toISOString() + '. ' +
      'Today tab: ' + _fmtTab(todayDate) + '. ' +
      'Baseline tab (14.04.26) found: ' + (baselineSheet ? 'yes' : 'no') + '. ' +
      'SKU count for Christmas/Easter filter: ' + skuCount
    );

    // ── Yesterday ─────────────────────────────────────────────────────────────
    var yesterday     = new Date(todayDate);
    yesterday.setDate(yesterday.getDate() - 1);
    var ydayResult = _findMostRecentTab(src, yesterday);
    var ydaySheet  = ydayResult.sheet;
    var ydayColInfo = ydaySheet ? _columnIndices(ydaySheet) : { idx: {}, headerRow: 1 };
    var ydayRows    = ydaySheet ? _filteredRows(ydaySheet, ydayColInfo) : [];
    var ydayAED    = _sum(ydayRows, 'overstockValue');
    var ydayUnits  = _sum(ydayRows, 'overstockQty');

    // positive = stock was cleared, negative = overstock grew
    var dailyClearedAED   = ydayAED   - todayAED;
    var dailyClearedUnits = ydayUnits - todayUnits;

    // ── Baseline AED ──────────────────────────────────────────────────────────
    var baselineAED = 0;
    if (baselineSheet) {
      var bColInfo = _columnIndices(baselineSheet);
      var bRows    = _filteredRows(baselineSheet, bColInfo);
      baselineAED = _sum(bRows, 'overstockValue');
    }

    var cumulativeCleared = baselineAED - todayAED;
    var clearedPct        = baselineAED > 0 ? cumulativeCleared / baselineAED : 0;

    // ── Deadline maths ────────────────────────────────────────────────────────
    var daysLeft    = Math.max(1, Math.ceil((CFG.DEADLINE_DATE - today) / 86400000));
    var requiredBurn = todayAED / daysLeft;

    // ── Trailing 7-day average ────────────────────────────────────────────────
    var t7          = _trailing7(src, todayDate);
    var avgBurn7d   = t7.avgBurn;
    var deltas      = t7.deltas;

    // ── Week-over-week ────────────────────────────────────────────────────────
    var weeks = _weeklyComparison(src, todayDate);

    // ── Top-risk SKU (highest remaining overstock AED) ───────────────────────
    var topSku = todayRows.slice()
      .sort(function(a, b) { return b.overstockValue - a.overstockValue; })[0] || null;

    // ── Weekly inventory trend from source Summary tab ────────────────────────
    var trend = _inventoryTrend(src);

    // ── Staleness flag ────────────────────────────────────────────────────────
    var daysBehind = Math.round((today - todayDate) / 86400000);

    // ── Write dashboard ───────────────────────────────────────────────────────
    _writeDashboard(out, {
      runStart       : runStart,
      today          : today,
      todayDate      : todayDate,
      isStale        : isStale,
      daysBehind     : daysBehind,
      todayAED       : todayAED,
      todayUnits     : todayUnits,
      ydayAED        : ydayAED,
      ydayUnits      : ydayUnits,
      dailyClearedAED   : dailyClearedAED,
      dailyClearedUnits : dailyClearedUnits,
      baselineAED    : baselineAED,
      baselineFound  : !!baselineSheet,
      cumulativeCleared : cumulativeCleared,
      clearedPct     : clearedPct,
      daysLeft       : daysLeft,
      requiredBurn   : requiredBurn,
      avgBurn7d      : avgBurn7d,
      deltas         : deltas,
      lastWeek       : weeks.lastWeek,
      thisWeek       : weeks.thisWeek,
      topSku         : topSku,
      trend          : trend,
      skuCount       : skuCount,
    });

    console.log('[Overstock Dashboard] Completed at ' + new Date().toISOString());

  } catch (err) {
    console.error('[Overstock Dashboard] Unhandled error: ' + err.stack);
    throw err;
  }
}

// ─── TAB NAVIGATION ───────────────────────────────────────────────────────────

function _midnightToday() {
  var d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

/** Formats a Date as DD.MM.YY for tab lookup (e.g. "23.04.26"). */
function _fmtTab(date) {
  return _pad2(date.getDate()) + '.' + _pad2(date.getMonth() + 1) + '.' +
         String(date.getFullYear()).slice(2);
}

function _pad2(n) {
  return n < 10 ? '0' + n : String(n);
}

/** Parses DD.MM.YY tab name into a Date at midnight, or null if not that format. */
function _parseTab(name) {
  var m = name.match(/^(\d{2})\.(\d{2})\.(\d{2})$/);
  if (!m) return null;
  return new Date(2000 + parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
}

/** Returns all dated tabs sorted newest → oldest. */
function _allDatedTabs(ss) {
  return ss.getSheets()
    .map(function(s) { return { sheet: s, date: _parseTab(s.getName()) }; })
    .filter(function(x) { return x.date !== null; })
    .sort(function(a, b) { return b.date - a.date; });
}

/** Returns a specific tab by date, or null. */
function _getTabByDate(ss, date) {
  return ss.getSheetByName(_fmtTab(date)) || null;
}

/**
 * Finds today's tab. If missing, falls back to the most recent dated tab on or before
 * targetDate and sets isStale = true.
 */
function _findMostRecentTab(ss, targetDate) {
  var exact = _getTabByDate(ss, targetDate);
  if (exact) return { sheet: exact, date: targetDate, isStale: false };

  var tabs = _allDatedTabs(ss);
  for (var i = 0; i < tabs.length; i++) {
    if (tabs[i].date <= targetDate) {
      return { sheet: tabs[i].sheet, date: tabs[i].date, isStale: true };
    }
  }
  return { sheet: null, date: null, isStale: true };
}

// ─── COLUMN INDEX DISCOVERY ───────────────────────────────────────────────────

/**
 * Normalises a header string: lowercase, strip backslashes, collapse whitespace.
 * Used for both the map keys and lookup keys so matching is resilient to
 * minor formatting differences.
 */
function _norm(s) {
  return String(s).trim().replace(/\\/g, '').replace(/\s+/g, ' ').toLowerCase();
}

/**
 * Scans the first 5 rows of a sheet to find the row that contains BOTH
 * "Item Code" AND "Order Interval". Tabs have a merged group-label row
 * above the real header row, so we cannot assume row 1 is the header.
 *
 * Returns the 1-based row number of the best match.
 * Falls back to the row with the most non-empty cells if sentinels aren't found.
 */
function _findHeaderRow(sheet) {
  var maxScan = Math.min(5, sheet.getLastRow());
  if (maxScan < 1) return 1;

  var rows = sheet.getRange(1, 1, maxScan, sheet.getLastColumn()).getValues();
  var sentinels = [_norm(HEADERS.ITEM_CODE), _norm(HEADERS.ORDER_INTERVAL)];

  for (var r = 0; r < rows.length; r++) {
    var normed = rows[r].map(function(h) { return _norm(h); });
    var hits   = sentinels.filter(function(s) { return normed.indexOf(s) >= 0; }).length;
    if (hits === sentinels.length) return r + 1;  // 1-based
  }

  // Fallback: use the row with the most non-empty header cells
  var best = 0, bestRow = 1;
  for (var r = 0; r < rows.length; r++) {
    var count = rows[r].filter(function(h) { return _norm(h) !== ''; }).length;
    if (count > best) { best = count; bestRow = r + 1; }
  }
  return bestRow;
}

/**
 * Finds the header row and builds a column-index map.
 * Returns { idx, headerRow } where:
 *   idx       — map of normalised-header → 0-based column index
 *   headerRow — 1-based row number where headers were found
 */
function _columnIndices(sheet) {
  var headerRow = _findHeaderRow(sheet);
  var headers   = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = {};
  headers.forEach(function(h, i) {
    var key = _norm(h);
    if (key) idx[key] = i;
  });
  return { idx: idx, headerRow: headerRow };
}

/** Looks up a 0-based column index by header name (normalised). Returns -1 if not found. */
function _col(idx, headerConstant) {
  var key = _norm(headerConstant);
  return (key in idx) ? idx[key] : -1;
}

// ─── DATA EXTRACTION ──────────────────────────────────────────────────────────

/**
 * Reads all data rows from a sheet and returns only those where
 * Order Interval = "Christmas/Easter".
 *
 * @param {Sheet}  sheet    — the source sheet
 * @param {Object} colInfo  — { idx, headerRow } returned by _columnIndices()
 *
 * Each returned object: { itemCode, description, supplier, soh, overstockValue, overstockQty }
 */
function _filteredRows(sheet, colInfo) {
  try {
    var idx       = colInfo.idx       || {};
    var headerRow = colInfo.headerRow || 1;
    var dataStart = headerRow + 1;    // first data row (1-based)

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < dataStart || lastCol < 1) return [];

    var oiCol  = _col(idx, HEADERS.ORDER_INTERVAL);
    var ovvCol = _col(idx, HEADERS.OVERSTOCK_VALUE);
    var ovqCol = _col(idx, HEADERS.OVERSTOCK_QTY);
    var icCol  = _col(idx, HEADERS.ITEM_CODE);
    var dCol   = _col(idx, HEADERS.DESCRIPTION);
    var sCol   = _col(idx, HEADERS.SUPPLIER);
    var sohCol = _col(idx, HEADERS.SOH);
    var icostCol = _col(idx, HEADERS.ITEM_COST);

    // ── Diagnostic log for every tab processed ───────────────────────────────
    console.info(
      '[Overstock Dashboard] Tab "' + sheet.getName() + '": ' +
      'header row=' + headerRow + ', ' +
      'Order Interval col=' + (oiCol  >= 0 ? oiCol  + 1 : 'NOT FOUND') + ' (idx ' + oiCol  + '), ' +
      'Item cost col='      + (icostCol >= 0 ? icostCol + 1 : 'NOT FOUND') + ' (idx ' + icostCol + '), ' +
      'Overstock AED col='  + (ovvCol >= 0 ? ovvCol + 1 : 'NOT FOUND') + ' (idx ' + ovvCol + '), ' +
      'Overstock Qty col='  + (ovqCol >= 0 ? ovqCol + 1 : 'NOT FOUND') + ' (idx ' + ovqCol + ')'
    );

    var data = sheet.getRange(dataStart, 1, lastRow - headerRow, lastCol).getValues();

    // ── Log first 3 Order Interval raw values (catches hidden spaces / alt spellings) ──
    if (oiCol >= 0) {
      var samples = data.slice(0, 3).map(function(row) {
        return JSON.stringify(String(row[oiCol]));
      });
      console.info(
        '[Overstock Dashboard] Tab "' + sheet.getName() + '" first 3 Order Interval values: ' +
        samples.join(', ')
      );
    } else {
      console.warn(
        '[Overstock Dashboard] Tab "' + sheet.getName() + '": Order Interval column NOT FOUND — ' +
        'available headers: ' + Object.keys(idx).slice(0, 20).join(', ')
      );
    }

    var filterNorm = _norm(CFG.FILTER_VALUE);

    return data
      .filter(function(row) {
        return oiCol >= 0 && _norm(row[oiCol]) === filterNorm;
      })
      .map(function(row) {
        return {
          itemCode      : icCol  >= 0 ? row[icCol]  : '',
          description   : dCol   >= 0 ? row[dCol]   : '',
          supplier      : sCol   >= 0 ? row[sCol]   : '',
          soh           : sohCol >= 0 ? _n(row[sohCol])  : 0,
          overstockValue: ovvCol >= 0 ? _n(row[ovvCol]) : 0,
          overstockQty  : ovqCol >= 0 ? _n(row[ovqCol]) : 0,
        };
      });

  } catch (e) {
    console.error('[Overstock Dashboard] _filteredRows error on "' + sheet.getName() + '": ' + e);
    return [];
  }
}

function _sum(rows, key) {
  return rows.reduce(function(s, r) { return s + (r[key] || 0); }, 0);
}

/** Safely parses a cell value to a number. */
function _n(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

// ─── TRAILING 7-DAY AVERAGE ───────────────────────────────────────────────────

/**
 * Computes the average daily AED cleared over the last 7 available day-pairs.
 * Returns { avgBurn: number, deltas: [{date, cleared}] }.
 * Skips any tab-pair where reading fails, with a log warning.
 */
function _trailing7(src, todayDate) {
  var tabs = _allDatedTabs(src)
    .filter(function(t) { return t.date <= todayDate; })
    .slice(0, 8);  // need up to 8 tabs to get 7 deltas

  var deltas = [];
  for (var i = 0; i < tabs.length - 1; i++) {
    try {
      var nColInfo = _columnIndices(tabs[i].sheet);
      var oColInfo = _columnIndices(tabs[i + 1].sheet);
      var nAED = _sum(_filteredRows(tabs[i].sheet, nColInfo), 'overstockValue');
      var oAED = _sum(_filteredRows(tabs[i + 1].sheet, oColInfo), 'overstockValue');
      deltas.push({ date: tabs[i].date, cleared: oAED - nAED });
    } catch (e) {
      console.warn('[Overstock Dashboard] _trailing7: skipping ' +
                   tabs[i].sheet.getName() + ': ' + e);
    }
  }

  var last7    = deltas.slice(0, 7);
  var total    = last7.reduce(function(s, d) { return s + d.cleared; }, 0);
  var avgBurn  = last7.length > 0 ? total / last7.length : 0;
  return { avgBurn: avgBurn, deltas: last7 };
}

// ─── WEEK-OVER-WEEK ───────────────────────────────────────────────────────────

/** Returns the Monday–Sunday bounds of the week containing `date`. */
function _weekBounds(date) {
  var d   = new Date(date);
  var day = d.getDay();  // 0 = Sun
  var mon = new Date(d);
  mon.setDate(d.getDate() + (day === 0 ? -6 : 1 - day));
  mon.setHours(0, 0, 0, 0);
  var sun = new Date(mon);
  sun.setDate(mon.getDate() + 6);
  return { start: mon, end: sun };
}

/**
 * Returns overstock-cleared totals for last complete week and current partial week.
 * Each result: { cleared, tabCount, flag, range }
 * cleared = null when there are fewer than 2 tabs in range (flag = true).
 */
function _weeklyComparison(src, todayDate) {
  var thisW   = _weekBounds(todayDate);
  var prevEnd = new Date(thisW.start);
  prevEnd.setDate(prevEnd.getDate() - 1);
  var lastW = _weekBounds(prevEnd);

  function weekStats(range) {
    var tabs = _allDatedTabs(src)
      .filter(function(t) { return t.date >= range.start && t.date <= range.end; });

    if (tabs.length < 2) {
      return { cleared: null, tabCount: tabs.length, flag: true, range: range };
    }
    try {
      var newest = tabs[0];
      var oldest = tabs[tabs.length - 1];
      var nAED = _sum(_filteredRows(newest.sheet, _columnIndices(newest.sheet)), 'overstockValue');
      var oAED = _sum(_filteredRows(oldest.sheet, _columnIndices(oldest.sheet)), 'overstockValue');
      return { cleared: oAED - nAED, tabCount: tabs.length, flag: false, range: range };
    } catch (e) {
      console.warn('[Overstock Dashboard] _weeklyComparison error: ' + e);
      return { cleared: null, tabCount: tabs.length, flag: true, range: range };
    }
  }

  return { lastWeek: weekStats(lastW), thisWeek: weekStats(thisW) };
}

// ─── INVENTORY TREND ──────────────────────────────────────────────────────────

/**
 * Reads the weekly inventory health tracker from the source file's Summary tab.
 * Returns [{week, totalInv, overstock}] for Wk15 onwards.
 */
function _inventoryTrend(src) {
  try {
    var summarySheet = src.getSheets().filter(function(s) {
      var n = s.getName().toLowerCase();
      return _parseTab(s.getName()) === null &&
             n.indexOf('raw')     === -1 &&
             n.indexOf('daily')   === -1 &&
             n.indexOf('arrival') === -1 &&
             n.indexOf('expir')   === -1 &&
             n.indexOf('weekly sale') === -1;
    })[0];

    if (!summarySheet) return [];

    var vals = summarySheet.getDataRange().getValues();
    var rows = [];
    for (var i = 0; i < vals.length; i++) {
      var lbl = String(vals[i][0]).replace(/\n/g, ' ').trim();
      if (/^Wk(1[5-9]|[2-9]\d)/.test(lbl)) {
        var inv  = _n(vals[i][1]);
        var over = _n(vals[i][2]);
        if (inv > 0 || over > 0) {
          rows.push({ week: lbl, totalInv: inv, overstock: over });
        }
      }
    }
    return rows;
  } catch (e) {
    console.warn('[Overstock Dashboard] _inventoryTrend: ' + e);
    return [];
  }
}

// ─── DASHBOARD WRITER ─────────────────────────────────────────────────────────

/**
 * Clears and rewrites the target sheet with the full dashboard layout.
 * 6-column layout: A–F.  No merged cells except section title rows.
 *
 * SAFETY: verifies the sheet name matches CFG.TARGET_TAB exactly before clearing.
 * If the name does not match, logs an error and aborts without touching anything.
 */
function _writeDashboard(sh, d) {
  // ── Safety guard: confirm we are writing to the correct tab ───────────────
  if (sh.getName() !== CFG.TARGET_TAB) {
    console.error(
      '[Overstock Dashboard] ABORT — sheet name mismatch. ' +
      'Expected "' + CFG.TARGET_TAB + '", got "' + sh.getName() + '". No changes made.'
    );
    return;
  }

  // Full clear: content, formats, and any existing merges
  sh.clear();

  var COLS    = 6;
  var builder = [];  // [{cells, merge, bg, fg, bold, labelBold, wow, wowCol, fontSize, paceOk}]

  // ── Formatting helpers (local) ───────────────────────────────────────────
  var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  function fAED(v) {
    if (v == null || isNaN(v)) return '—';
    return 'AED ' + _fmtN(Math.round(v));
  }
  function fNum(v) {
    if (v == null || isNaN(v)) return '—';
    return _fmtN(Math.round(v));
  }
  function fPct(v) {
    if (v == null || isNaN(v)) return '—';
    return (v * 100).toFixed(1) + '%';
  }
  function fDate(v) {
    if (!(v instanceof Date) || isNaN(v)) return v || '—';
    return _pad2(v.getDate()) + '-' + MONTHS[v.getMonth()] + '-' + String(v.getFullYear()).slice(2);
  }
  function fDT(v) {
    if (!(v instanceof Date) || isNaN(v)) return '—';
    return fDate(v) + ' ' + _pad2(v.getHours()) + ':' + _pad2(v.getMinutes());
  }

  function row(cells, opts) {
    while (cells.length < COLS) cells.push('');
    builder.push(Object.assign({ cells: cells }, opts || {}));
  }

  // ── Netlify dashboard link (row 1, always present) ───────────────────────

  row(['Dashboard: ' + CFG.NETLIFY_URL, '', '', '', '', ''],
      { merge: true, bg: '#F0F4FF', fg: '#1155CC', bold: false, fontSize: 10 });

  // ── Banner rows ───────────────────────────────────────────────────────────

  if (d.isStale || d.daysBehind > CFG.STALE_THRESHOLD) {
    row(
      ['⚠  DATA STALE — ' + d.daysBehind + ' day(s) behind today. ' +
       'Latest available tab: ' + _fmtTab(d.todayDate) + '. Today\'s tab not found.'],
      { merge: true, bg: '#BF0000', fg: '#FFFFFF', bold: true }
    );
  }

  if (d.skuCount === 0) {
    row(
      ['⚠  ERROR — 0 "Christmas/Easter" SKUs found in tab ' + _fmtTab(d.todayDate) +
       '. Check the Order Interval column or source sheet structure.'],
      { merge: true, bg: '#BF0000', fg: '#FFFFFF', bold: true }
    );
  }

  // ── Title block ───────────────────────────────────────────────────────────

  row(
    ['CHRISTMAS / EASTER OVERSTOCK CLEARANCE TRACKER'],
    { merge: true, bg: '#1A1A2E', fg: '#FFFFFF', bold: true, fontSize: 13 }
  );
  row(
    ['Source tab: ' + _fmtTab(d.todayDate) +
     '   ·   Run: ' + fDT(d.runStart) +
     '   ·   Deadline: 31-May-26' +
     '   ·   ' + d.daysLeft + ' days remaining'],
    { merge: true, bg: '#16213E', fg: '#9999BB', fontSize: 9 }
  );
  row([], {});  // spacer

  // ── Clearance snapshot ────────────────────────────────────────────────────

  row(['CLEARANCE SNAPSHOT'],
      { merge: true, bg: '#0F3460', fg: '#FFFFFF', bold: true });

  row(['Cleared %',          fPct(d.clearedPct),
       'Avg 7d Burn / Day',  fAED(d.avgBurn7d),
       'Required Burn / Day', fAED(d.requiredBurn)],
      { boldCols: [0, 2, 4] });

  var onPace = d.avgBurn7d >= d.requiredBurn;
  row(['Pace',
       onPace ? '✓  ON PACE' : '✗  BEHIND PACE',
       '7d avg  vs  required',
       fAED(d.avgBurn7d) + '  vs  ' + fAED(d.requiredBurn),
       'SKUs tracked', d.skuCount],
      { boldCols: [0, 2, 4], paceOk: onPace });

  row([], {});  // spacer

  // ── Yesterday's performance ───────────────────────────────────────────────

  var yday = new Date(d.todayDate);
  yday.setDate(yday.getDate() - 1);

  row(["YESTERDAY'S PERFORMANCE"],
      { merge: true, bg: '#0F3460', fg: '#FFFFFF', bold: true });
  row(['Date', fDate(yday)], { labelBold: true });
  row(['AED Cleared (overstock Δ)',   fAED(d.dailyClearedAED)],
      { labelBold: true, wow: d.dailyClearedAED,   wowCol: 2 });
  row(['Units Cleared (overstock Δ)', fNum(d.dailyClearedUnits)],
      { labelBold: true, wow: d.dailyClearedUnits, wowCol: 2 });

  // WoW vs same weekday last week (delta[6] if we have 7 data points)
  if (d.deltas.length >= 7) {
    var wowBase = d.deltas[6].cleared;
    var wowDiff = d.dailyClearedAED - wowBase;
    row(['WoW vs same weekday last week', fAED(wowDiff)],
        { labelBold: true, wow: wowDiff, wowCol: 2 });
  }

  row([], {});  // spacer

  // ── Three side-by-side cards ──────────────────────────────────────────────

  var lw  = d.lastWeek;
  var tw  = d.thisWeek;
  var sk  = d.topSku;

  var lwLabel = lw.range ? fDate(lw.range.start) + ' – ' + fDate(lw.range.end) : '—';
  var twLabel = tw.range ? fDate(tw.range.start) + ' – ' + fDate(tw.range.end) : '—';

  row(['LAST COMPLETE WEEK', '', 'CURRENT PARTIAL WEEK', '', 'TOP-RISK SKU', ''],
      { bg: '#0F3460', fg: '#FFFFFF', bold: true });

  row([lwLabel, '', twLabel, '', sk ? sk.itemCode : '—', ''],
      { boldCols: [0, 2, 4] });

  row(['Cleared AED',
       lw.cleared != null ? fAED(lw.cleared) : (lw.flag ? '⚠ missing tabs' : '—'),
       'Cleared AED',
       tw.cleared != null ? fAED(tw.cleared) : (tw.flag ? '⚠ missing tabs' : '—'),
       'Description',
       sk ? _trunc(sk.description, 32) : '—'],
      { boldCols: [0, 2, 4] });

  var cardWoW = (lw.cleared != null && tw.cleared != null) ? tw.cleared - lw.cleared : null;
  row(['Tabs in range', lw.tabCount,
       'WoW vs last week', cardWoW != null ? fAED(cardWoW) : '—',
       'Overstock AED', sk ? fAED(sk.overstockValue) : '—'],
      { boldCols: [0, 2, 4], wow: cardWoW, wowCol: 4 });

  row(['', '', '', '',
       'Overstock Units', sk ? fNum(sk.overstockQty) : '—'],
      { boldCols: [4] });

  row(['', '', '', '',
       'SOH (units)', sk ? fNum(sk.soh) : '—'],
      { boldCols: [4] });

  row([], {});  // spacer

  // ── Yesterday's close ─────────────────────────────────────────────────────

  row(["YESTERDAY'S CLOSE"],
      { merge: true, bg: '#0F3460', fg: '#FFFFFF', bold: true });
  row(['Closing Overstock AED',            fAED(d.todayAED)],
      { labelBold: true });
  row(['Closing Overstock Units',          fNum(d.todayUnits)],
      { labelBold: true });
  row(['Cleared Yesterday (AED)',          fAED(d.dailyClearedAED)],
      { labelBold: true, wow: d.dailyClearedAED,   wowCol: 2 });
  row(['Cleared Yesterday (Units)',        fNum(d.dailyClearedUnits)],
      { labelBold: true, wow: d.dailyClearedUnits, wowCol: 2 });
  row(['Cumulative Cleared since 14-Apr-26 (AED)', fAED(d.cumulativeCleared)],
      { labelBold: true });

  row([], {});  // spacer

  // ── Cumulative since 14 Apr 2026 ──────────────────────────────────────────

  row(['CUMULATIVE SINCE 14-APR-2026'],
      { merge: true, bg: '#0F3460', fg: '#FFFFFF', bold: true });
  row(['Baseline Overstock AED (14-Apr-26)',
       fAED(d.baselineAED),
       d.baselineFound ? '' : '⚠  Baseline tab (14.04.26) not found'],
      { labelBold: true });
  row(['Current Overstock AED',               fAED(d.todayAED)],   { labelBold: true });
  row(['Current Overstock Units',             fNum(d.todayUnits)], { labelBold: true });
  row(['Total Cleared AED',                   fAED(d.cumulativeCleared)], { labelBold: true });
  row(['Cleared %',                           fPct(d.clearedPct)], { labelBold: true });
  row(['Avg Daily Burn (7d trailing)',         fAED(d.avgBurn7d)],  { labelBold: true });
  row(['Required Burn / Day (to 31-May-26)',   fAED(d.requiredBurn)], { labelBold: true });
  row(['Days Remaining to Deadline',           d.daysLeft],         { labelBold: true });

  row([], {});  // spacer

  // ── Weekly inventory trend ────────────────────────────────────────────────

  if (d.trend && d.trend.length > 0) {
    row(['WEEKLY INVENTORY TREND (Wk15 onwards — from Inventory Health Tracker)'],
        { merge: true, bg: '#0F3460', fg: '#FFFFFF', bold: true });
    row(['Week', 'Total Inventory AED', 'Overstock >90d AED'],
        { bg: '#162040', fg: '#FFFFFF', boldCols: [0, 1, 2] });
    for (var ti = 0; ti < d.trend.length; ti++) {
      row([d.trend[ti].week, fAED(d.trend[ti].totalInv), fAED(d.trend[ti].overstock)],
          { labelBold: true });
    }
  }

  // ── Write values ──────────────────────────────────────────────────────────

  if (builder.length === 0) return;

  var values = builder.map(function(r) { return r.cells.slice(0, COLS); });
  sh.getRange(1, 1, values.length, COLS).setValues(values);

  // Merge pass — must happen AFTER setValues
  for (var mi = 0; mi < builder.length; mi++) {
    if (builder[mi].merge) {
      sh.getRange(mi + 1, 1, 1, COLS).merge();
    }
  }

  // Formatting pass
  _applyFormatting(sh, builder, COLS);

  // Column widths: A=label, B=value, C=label, D=value, E=label, F=value
  var widths = [300, 160, 300, 160, 200, 160];
  for (var wi = 0; wi < widths.length; wi++) {
    sh.setColumnWidth(wi + 1, widths[wi]);
  }

  sh.setFrozenRows(2);  // freeze title + subtitle

  SpreadsheetApp.flush();
}

function _applyFormatting(sh, builder, COLS) {
  var FONT      = 'Arial';
  var BASE_SIZE = 10;

  for (var i = 0; i < builder.length; i++) {
    var rn  = i + 1;
    var b   = builder[i];
    var r6  = sh.getRange(rn, 1, 1, COLS);

    r6.setFontFamily(FONT)
      .setFontSize(b.fontSize || BASE_SIZE)
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    if (b.bg) r6.setBackground(b.bg);
    if (b.fg) r6.setFontColor(b.fg);
    if (b.bold === true) r6.setFontWeight('bold');
    if (b.labelBold) sh.getRange(rn, 1).setFontWeight('bold');

    // Per-column bold
    if (b.boldCols) {
      b.boldCols.forEach(function(ci) {
        sh.getRange(rn, ci + 1).setFontWeight('bold');
      });
    }

    // Alignment: odd cols (A,C,E) = left; even cols (B,D,F) = right
    [1, 3, 5].forEach(function(c) { sh.getRange(rn, c).setHorizontalAlignment('left'); });
    [2, 4, 6].forEach(function(c) { sh.getRange(rn, c).setHorizontalAlignment('right'); });

    // WoW conditional colour
    if ('wow' in b && b.wow != null && b.wowCol) {
      var cell = sh.getRange(rn, b.wowCol);
      if      (b.wow > 0) cell.setFontColor('#007A33').setFontWeight('bold');
      else if (b.wow < 0) cell.setFontColor('#CC0000').setFontWeight('bold');
    }

    // Pace pill styling (col B)
    if ('paceOk' in b) {
      sh.getRange(rn, 2)
        .setBackground(b.paceOk ? '#007A33' : '#CC0000')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
  }
}

// ─── FORMATTING HELPERS ───────────────────────────────────────────────────────

/** UK-style number formatting: 1,234,567 */
function _fmtN(v) {
  // toLocaleString is available in Apps Script V8
  return Math.round(v).toLocaleString('en-GB');
}

function _trunc(s, n) {
  if (!s) return '';
  return s.length > n ? s.slice(0, n - 1) + '\u2026' : s;
}
