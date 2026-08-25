/**
 * Create / repair Version 2 sheet tabs and headers.
 * Location is validated in Apps Script (not via sheet dropdowns),
 * because sheet data-validation rules drift onto wrong columns after header changes.
 */

function setupInventorySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Open a Google Spreadsheet and run this from Extensions → Apps Script.");
  }
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());

  var names = Object.keys(SHEET_HEADERS_);
  names.forEach(function (sheetName) {
    ensureSheetWithHeaders_(ss, sheetName, SHEET_HEADERS_[sheetName]);
  });

  // Always strip every dropdown rule — leftover Location lists on Student Name / KE Number / Issued By / Date cause E2 errors.
  clearAllSheetValidations_();
  migrateInventoryToAddLedger_();

  return {
    success: true,
    message: "Sheets ready. Inventory is an add-log (one row per stock-in). On-hand = adds minus issues.",
    spreadsheetId: ss.getId(),
    sheets: names
  };
}

/** Menu / manual: clear bad dropdowns without rewriting headers. */
function repairSheetValidations() {
  clearAllSheetValidations_();
  return {
    success: true,
    message: "Cleared all data-validation dropdowns on Track v2 sheets. Try Issue again."
  };
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }

  clearSheetValidations_(sh);

  // Drop obsolete "Order ID" column on ISSUE_REGISTER if it still exists
  if (sheetName === KE.SHEETS.ISSUE) {
    removeColumnByHeader_(sh, "Order ID");
  }

  var lastCol = Math.max(sh.getLastColumn(), headers.length);
  // Clear old header cells beyond current schema
  if (lastCol > headers.length) {
    sh.getRange(1, headers.length + 1, 1, lastCol).clearContent();
  }

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1a2235")
    .setFontColor("#C9A84C");
  sh.setFrozenRows(1);
  return sh;
}

function clearAllSheetValidations_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = getSS_();
  }
  Object.keys(SHEET_HEADERS_).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      clearSheetValidations_(sh);
    }
  });
}

/**
 * Clear validations column-by-column (more reliable than one huge range).
 */
function clearSheetValidations_(sh) {
  if (!sh) {
    return;
  }
  var maxCol = Math.max(sh.getMaxColumns(), sh.getLastColumn(), 20);
  var maxRow = Math.max(sh.getMaxRows(), sh.getLastRow(), 2);
  // Cap to avoid timeouts on huge unused grids
  maxCol = Math.min(maxCol, 30);
  maxRow = Math.min(maxRow, 2000);
  for (var c = 1; c <= maxCol; c++) {
    try {
      sh.getRange(1, c, maxRow, c).clearDataValidations();
    } catch (e) { /* continue */ }
  }
  try {
    sh.getDataRange().clearDataValidations();
  } catch (e2) { /* ignore */ }
}

function removeColumnByHeader_(sh, headerName) {
  if (!sh || sh.getLastColumn() < 1) {
    return;
  }
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (normalize_(headers[i]) === headerName) {
      sh.deleteColumn(i + 1);
      return;
    }
  }
}
