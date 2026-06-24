function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1");
  var data = sheet.getDataRange().getDisplayValues();

  var output = JSON.stringify(data);

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
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
