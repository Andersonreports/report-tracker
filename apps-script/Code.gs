function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Default to "Sheet1" so the existing dashboard fetch (no query param) is
  // unchanged. The Re-analysis tab calls this with ?sheet=Re-analysis.
  var sheetName = (e && e.parameter && e.parameter.sheet) ? e.parameter.sheet : "Sheet1";
  var sheet = ss.getSheetByName(sheetName);

  // getSheetByName is case-sensitive; fall back to a case-insensitive match so
  // "RE-ANALYSIS" / "Re-analysis" etc. all resolve to the same tab.
  if (!sheet) {
    var wanted = String(sheetName).toLowerCase();
    var allSheets = ss.getSheets();
    for (var i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getName().toLowerCase() === wanted) { sheet = allSheets[i]; break; }
    }
  }

  // Unknown sheet → return an empty grid instead of throwing, so a missing
  // Re-analysis sheet never breaks the caller.
  var data = sheet ? readGridDayFirst(sheet) : [];

  var output = JSON.stringify(data);

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// Returns the sheet as text, but with every genuine date cell written day-first as
// dd-MM-yyyy regardless of how that cell happens to be formatted.
//
// getDisplayValues() alone returns each cell exactly as the sheet renders it, and the
// date columns do not share one format: a comparison against a pre-edit copy of the
// workbook found 1797 date cells rendering day-first and 700 rendering month-first, all
// 700 in ANALYZE DATE. The frontend reads a slash date day-first, so those 700 came
// through with day and month swapped — 04/09/2026 read as 4 Sep when the cell held
// 9 April. Nothing in the sheet was wrong; only the format differed.
//
// Formatting the real Date objects here removes the ambiguity at the source, so no cell
// format can mislead the app again in any column. dd-MM-yyyy is deliberate: it is what
// the frontend already expects, so this needs no matching change there. Cells holding
// text rather than a date are passed through untouched.
function readGridDayFirst(sheet) {
  var range   = sheet.getDataRange();
  var display = range.getDisplayValues();   // text, as rendered
  var typed   = range.getValues();          // real types, so Dates are recognisable
  var tz      = sheet.getParent().getSpreadsheetTimeZone();

  for (var r = 0; r < typed.length; r++) {
    for (var c = 0; c < typed[r].length; c++) {
      var v = typed[r][c];
      if (v instanceof Date && !isNaN(v.getTime())) {
        display[r][c] = Utilities.formatDate(v, tz, 'dd-MM-yyyy');
      }
    }
  }
  return display;
}

// Routes writes coming from the Release Date picker on the Dashboard and the
// Remove action on the Released tab (front end/report_tracker.html:
// markSelectedAsReleased, applyBulkRemark, saveRowRemark, removeFromReleased).
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === 'markReleased') return handleMarkReleased(payload);
    if (payload.action === 'unmarkReleased') return handleUnmarkReleased(payload);
    if (payload.action === 'updateRemarks') return handleUpdateRemarks(payload);
    return jsonOutput({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

function handleMarkReleased(payload) {
  var dateStr = String(payload.date || '').trim();
  var items = payload.items || [];
  if (!dateStr || !items.length) {
    return jsonOutput({ ok: false, error: 'Missing date or items' });
  }

  // Incoming date is YYYY-MM-DD (parsed from the free-text Release Date field);
  // the sheet's other date columns read as DD-MM-YYYY text, so write in the same format.
  var dp = dateStr.split('-');
  var relDateText = dp.length === 3 ? (dp[2] + '-' + dp[1] + '-' + dp[0]) : dateStr;

  var result = writeColumnForItems('rel_date', items, function () { return relDateText; });
  return jsonOutput(result);
}

function handleUnmarkReleased(payload) {
  var items = payload.items || [];
  if (!items.length) {
    return jsonOutput({ ok: false, error: 'Missing items' });
  }
  var result = writeColumnForItems('rel_date', items, function () { return ''; });
  return jsonOutput(result);
}

function handleUpdateRemarks(payload) {
  var items = payload.items || [];
  if (!items.length) {
    return jsonOutput({ ok: false, error: 'Missing items' });
  }
  var result = writeColumnForItems('remark', items, function (item) { return String(item.remark || ''); });
  return jsonOutput(result);
}

// Mirrors the same header-alias and run-section detection the frontend uses
// (COL_MAP, parseGrid) so a row found in the UI maps back to the same row in the
// sheet, then writes valueFn(item) into the given column for every matched item.
function writeColumnForItems(colKey, items, valueFn) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  var grid = sheet.getDataRange().getDisplayValues();

  var HEADER_ALIASES = {
    sno:      ['s.no', 'sno', 'serial', '#'],
    gen_id:   ['gen_id', 'genid', 'genetic id', 'gen id'],
    and_id:   ['anderson_id', 'andersonid', 'anderson id'],
    rel_date: ['report release date', 'release date', 'released on', 'report_release_date'],
    remark:   ['remark', 'remarks', 'note', 'notes']
  };

  function normCell(v) { return String(v == null ? '' : v).trim(); }

  function isHeaderRow(row) {
    var first = '';
    for (var i = 0; i < row.length; i++) {
      if (normCell(row[i]) !== '') { first = normCell(row[i]); break; }
    }
    first = first.toLowerCase().replace(/[\s.]/g, '');
    return ['sno', 'srno', 'slno', 'no', '#', 'serial'].indexOf(first) !== -1;
  }

  function findRunCellText(row) {
    for (var i = 0; i < row.length; i++) {
      if (/Run-?\d+/i.test(normCell(row[i]))) return normCell(row[i]);
    }
    return null;
  }

  function colIndexFor(headers, aliases) {
    var lower = headers.map(function (h) { return normCell(h).toLowerCase(); });
    for (var a = 0; a < aliases.length; a++) {
      for (var i = 0; i < lower.length; i++) {
        if (lower[i] === aliases[a] || lower[i].indexOf(aliases[a]) !== -1) return i;
      }
    }
    return -1;
  }

  var headers = null;
  var colIdx = {};
  var currentRunTxt = null;

  for (var r = 0; r < grid.length; r++) {
    var row = grid[r];
    var nonEmptyCount = 0;
    for (var c = 0; c < row.length; c++) { if (normCell(row[c]) !== '') nonEmptyCount++; }
    if (nonEmptyCount === 0) continue;

    var rowIsHeader = isHeaderRow(row);
    var runCellTxt = findRunCellText(row);

    // Section marker row, e.g. "SURFseq- Run-77" — not a data row.
    if (runCellTxt && !rowIsHeader && nonEmptyCount < 5) {
      currentRunTxt = runCellTxt;
      continue;
    }

    if (rowIsHeader) {
      headers = row;
      colIdx = {
        sno:      colIndexFor(headers, HEADER_ALIASES.sno),
        gen_id:   colIndexFor(headers, HEADER_ALIASES.gen_id),
        and_id:   colIndexFor(headers, HEADER_ALIASES.and_id),
        rel_date: colIndexFor(headers, HEADER_ALIASES.rel_date),
        remark:   colIndexFor(headers, HEADER_ALIASES.remark)
      };
      continue;
    }

    if (!headers || colIdx[colKey] === -1) continue;

    var rowSno    = colIdx.sno    !== -1 ? normCell(row[colIdx.sno])    : '';
    var rowGenId  = colIdx.gen_id !== -1 ? normCell(row[colIdx.gen_id]) : '';
    var rowAndId  = colIdx.and_id !== -1 ? normCell(row[colIdx.and_id]) : '';
    var rowRunTxt = currentRunTxt || '—';

    for (var k = 0; k < items.length; k++) {
      var it = items[k];
      if (it._matched) continue;
      var runMatch = normCell(it.run_text || '—') === rowRunTxt;
      var snoMatch = normCell(it.sno) === rowSno;
      var idMatch = (it.gen_id && rowGenId && normCell(it.gen_id) === rowGenId) ||
                    (it.and_id && rowAndId && normCell(it.and_id) === rowAndId);
      if (runMatch && snoMatch && idMatch) {
        sheet.getRange(r + 1, colIdx[colKey] + 1).setValue(valueFn(it));
        it._matched = true;
      }
    }
  }

  var updated = 0, notFound = [];
  items.forEach(function (it) {
    if (it._matched) updated++; else notFound.push(it);
  });

  return { ok: true, updated: updated, notFound: notFound };
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
