/**
 * Creates / repairs sheet tabs and headers from the Excel template.
 * Run setupInventorySheets() once from the script editor.
 */

function setupInventorySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || getSS_();
  Object.keys(SHEET_HEADERS_).forEach(function (name) {
    ensureSheetWithHeaders_(ss, name, SHEET_HEADERS_[name]);
  });
  ensurePasswordColumn_(ss);
  ensureInventoryAdvancedColumns_(ss);
  ensureProcurementRegisterColumns_(ss);
  applyValidations_(ss);
  return "Inventory sheets ready: " + Object.keys(SHEET_HEADERS_).join(", ");
}

/** Adds Vendor column on REQUEST_REGISTER / PAYMENTS when upgrading older sheets. */
function ensureProcurementRegisterColumns_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet() || getSS_();
  var req = ss.getSheetByName(KE.SHEETS.REQUEST);
  if (req) {
    var reqHeaders = readHeaderRow_(req).map(function (h) { return String(h || "").trim(); });
    if (findCol_(reqHeaders, ["Vendor", "Vendor Name"]) < 0) {
      var custCol = findCol_(reqHeaders, ["Customer Name", "Customer"]);
      var insertAt = custCol >= 0 ? custCol + 1 : 3;
      req.insertColumnBefore(insertAt);
      req.getRange(1, insertAt).setValue("Vendor");
      req.getRange(1, insertAt)
        .setFontWeight("bold")
        .setBackground("#1A1A2E")
        .setFontColor("#F5E6B8");
    }
  }
  var pay = ss.getSheetByName(KE.SHEETS.PAYMENT);
  if (pay) {
    var payHeaders = readHeaderRow_(pay).map(function (h) { return String(h || "").trim(); });
    var payVendor = findCol_(payHeaders, ["Vendor", "Vendor Name"]);
    var payCustomer = findCol_(payHeaders, ["Customer", "Customer Name"]);
    if (payVendor < 0 && payCustomer >= 0) {
      pay.getRange(1, payCustomer + 1).setValue("Vendor");
    } else if (payVendor < 0) {
      pay.insertColumnAfter(2);
      pay.getRange(1, 3).setValue("Vendor");
    }
  }
}

/**
 * Adds the new columns introduced for advanced features without losing existing data.
 */
function ensureInventoryAdvancedColumns_(ss) {
  var sh = ss.getSheetByName(KE.SHEETS.INVENTORY);
  if (!sh) {
    return;
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h || "").trim(); });
  var required = ["Inventory Type", "Allocated To", "Reserved Qty", "Cost Price", "Last Movement", "Image URL"];
  required.forEach(function (col) {
    if (findCol_(existing, col) >= 0) {
      return;
    }
    lastCol += 1;
    sh.getRange(1, lastCol).setValue(col);
    existing.push(col);
  });
}

function ensurePasswordColumn_(ss) {
  var sh = ss.getSheetByName(KE.SHEETS.USER);
  if (!sh) {
    return;
  }
  var lastCol = Math.max(sh.getLastColumn(), 5);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (findCol_(headers, ["Password", "password"]) >= 0) {
    return;
  }
  sh.getRange(1, lastCol + 1).setValue("Password");
}

function ensureSheetWithHeaders_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  var existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var blank = existing.every(function (c) { return !c; });
  if (blank) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1A1A2E")
      .setFontColor("#F5E6B8");
  }
  return sh;
}

function applyValidations_(ss) {
  var locRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(KE.LOCATIONS, true)
    .setAllowInvalid(false)
    .build();

  var roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      KE.ROLES.ADMIN,
      KE.ROLES.ACCOUNTS,
      KE.ROLES.MANAGER,
      KE.ROLES.TRAINER
    ], true)
    .setAllowInvalid(false)
    .build();

  var inv = ss.getSheetByName(KE.SHEETS.INVENTORY);
  if (inv && inv.getLastRow() > 1) {
    inv.getRange(2, 6, Math.max(inv.getLastRow(), 500), 1).setDataValidation(locRule);
  }

  var ven = ss.getSheetByName(KE.SHEETS.VENDOR);
  if (ven && ven.getLastRow() > 1) {
    ven.getRange(2, 6, Math.max(ven.getLastRow(), 200), 1).setDataValidation(locRule);
  }

  var usr = ss.getSheetByName(KE.SHEETS.USER);
  if (usr) {
    usr.getRange(2, 3, Math.max(usr.getLastRow(), 50), 1).setDataValidation(roleRule);
    usr.getRange(2, 4, Math.max(usr.getLastRow(), 50), 1).setDataValidation(locRule);
  }

  applyAdvancedValidations_(ss, locRule);
}

/**
 * Builds and applies validations for the new advanced sheets safely.
 * Each rule is created only when (a) its source values list is non-empty and
 * (b) the target sheet/column actually exists. This prevents
 * "Argument cannot be null: values" when Apps Script hasn't synced the
 * latest Config.gs yet.
 */
function applyAdvancedValidations_(ss, locRule) {
  var typeRule = buildListValidation_(KE && KE.INVENTORY_TYPES);
  var transferStatusRule = buildListValidation_(objectValues_(KE && KE.TRANSFER_STATUS));
  var sampleStatusRule = buildListValidation_(objectValues_(KE && KE.SAMPLE_STATUS));

  var invSh = KE && KE.SHEETS && KE.SHEETS.INVENTORY
    ? ss.getSheetByName(KE.SHEETS.INVENTORY)
    : null;
  if (invSh && typeRule) {
    var invHeaders = readHeaderRow_(invSh);
    var typeCol = findCol_(invHeaders, "Inventory Type");
    if (typeCol >= 0) {
      invSh.getRange(2, typeCol + 1, Math.max(invSh.getLastRow(), 500), 1).setDataValidation(typeRule);
    }
  }

  var trSh = KE && KE.SHEETS && KE.SHEETS.TRANSFER
    ? ss.getSheetByName(KE.SHEETS.TRANSFER)
    : null;
  if (trSh) {
    var trHeaders = readHeaderRow_(trSh);
    var trFrom = findCol_(trHeaders, "From Location");
    var trTo = findCol_(trHeaders, "To Location");
    var trStatus = findCol_(trHeaders, "Status");
    if (trFrom >= 0 && locRule) {
      trSh.getRange(2, trFrom + 1, Math.max(trSh.getLastRow(), 200), 1).setDataValidation(locRule);
    }
    if (trTo >= 0 && locRule) {
      trSh.getRange(2, trTo + 1, Math.max(trSh.getLastRow(), 200), 1).setDataValidation(locRule);
    }
    if (trStatus >= 0 && transferStatusRule) {
      trSh.getRange(2, trStatus + 1, Math.max(trSh.getLastRow(), 200), 1).setDataValidation(transferStatusRule);
    }
  }

  var smpSh = KE && KE.SHEETS && KE.SHEETS.SAMPLE
    ? ss.getSheetByName(KE.SHEETS.SAMPLE)
    : null;
  if (smpSh) {
    var smpHeaders = readHeaderRow_(smpSh);
    var smpLoc = findCol_(smpHeaders, "Location");
    var smpStatus = findCol_(smpHeaders, "Status");
    if (smpLoc >= 0 && locRule) {
      smpSh.getRange(2, smpLoc + 1, Math.max(smpSh.getLastRow(), 200), 1).setDataValidation(locRule);
    }
    if (smpStatus >= 0 && sampleStatusRule) {
      smpSh.getRange(2, smpStatus + 1, Math.max(smpSh.getLastRow(), 200), 1).setDataValidation(sampleStatusRule);
    }
  }
}

function buildListValidation_(values) {
  if (!values || !values.length) {
    return null;
  }
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
}

function objectValues_(obj) {
  if (!obj) {
    return [];
  }
  return Object.keys(obj).map(function (k) { return obj[k]; }).filter(function (v) {
    return v !== undefined && v !== null && v !== "";
  });
}

function readHeaderRow_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1, 1, 1, lastCol).getValues()[0];
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Kings Inventory")
    .addItem("Open dashboard", "openDashboardSidebar")
    .addItem("Open login (web app)", "openInventoryWebLogin")
    .addItem("Setup sheets", "setupInventorySheets")
    .addItem("Load sample demo data", "seedSampleData")
    .addItem("Diagnose sheet row counts", "diagnoseInventorySheets")
    .addItem("Save spreadsheet ID", "setSpreadsheetId")
    .addToUi();
}

function openInventoryWebLogin() {
  var url = getWebAppUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      "Deploy the web app first, then run setSpreadsheetId()."
    );
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<p>Opening login…</p><script>window.open("' + url + '?page=login", "_blank");</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, "Inventory login");
}

function openDashboardSidebar() {
  var t = HtmlService.createTemplateFromFile("Dashboard");
  t.urlToken = "";
  var html = t.evaluate()
    .setTitle("Kings Equestrian Inventory")
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}
