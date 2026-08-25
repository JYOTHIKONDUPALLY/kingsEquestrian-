/**
 * Demo rows for dashboard testing. Run seedSampleData() from Apps Script or sheet menu.
 */

function seedSampleData(force) {
  setupInventorySheets();
  if (!force && PropertiesService.getScriptProperties().getProperty("KE_SAMPLE_SEEDED")) {
    ensureAdvancedSampleData_();
    return "Sample data already loaded. Ensured advanced demo rows for pricing/finance/transfer.";
  }

  var today = todayStr_();

  appendRowsIfEmpty_(KE.SHEETS.USER, [
    ["admin@ke.demo", "KE Admin", "Admin", "Farm", "admin@123"],
    ["accounts@ke.demo", "Accounts User", "Accounts", "Bangalore", "admin@123"],
    ["trainer@ke.demo", "Farm Trainer", "Trainer", "Farm", "admin@123"],
    ["manager@ke.demo", "Farm Manager", "Manager", "Farm", "admin@123"]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.VENDOR, [
    ["VEN-0001", "Decathlon India", "DEC", "9800000001", "vendor@decathlon.in", "Bangalore"],
    ["VEN-0002", "GPA Equestrian", "GPA", "9800000002", "sales@gpa.in", "Farm"],
    ["VEN-0003", "Saif", "SA", "9800000003", "vendor@saif.in", "Hyderabad"]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.INVENTORY, [
    ["KE-DEC-HEL-PRO1", "Riding Helmet Pro", "HEL", "PRO1", "Decathlon India", "Farm",      8, 5, "Farm Store A", today, "Business Inventory",       "",          0, 3500, today],
    ["KE-GPA-SAD-DRS01","Dressage Saddle",   "SAD", "DRS01","GPA Equestrian",  "Farm",      2, 3, "Farm Store B", today, "Operational Inventory",    "",          0, 85000, today],
    ["KE-DEC-BOO-RID02","Riding Boots",      "BOO", "RID02","Decathlon India", "Bangalore",12, 4, "BLR Rack 1",   today, "Business Inventory",       "",          0, 4500, today],
    ["KE-GPA-HEL-STD01","Standard Helmet",   "HEL", "STD01","GPA Equestrian",  "Bangalore", 1, 5, "BLR Rack 2",   today, "Procurement Inventory", "",          0, 3200, today],
    ["KE-SA-HEL-STD01", "Helmet",            "HEL", "STD01","Saif",            "Hyderabad", 5, 5, "HYD Shelf 1",  today, "Business Inventory",       "",          0, 3500, today],
    ["KE-DEC-GLO-WIN01","Winter Gloves",     "GLO", "WIN01","Decathlon India", "Farm",     15,10, "Farm Store A", today, "Sample / Demo Inventory",  "",          0, 1200, today]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.REQUEST, [
    ["REQ-0001", today, "Saif", "Helmet", 20, "Hyderabad", "Farm Trainer", "Unpaid", "Pending"],
    ["REQ-0002", today, "Decathlon India", "Riding Boots", 10, "Bangalore", "KE Admin", "Paid", "Approved"],
    ["REQ-0003", today, "GPA Equestrian", "Dressage Saddle", 1, "Farm", "Farm Trainer", "Paid", "Ordered"],
    ["REQ-0004", today, "GPA Equestrian", "Standard Helmet", 5, "Bangalore", "Accounts User", "Paid", "Received"]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.PAYMENT, [
    ["PAY-0001", "REQ-0002", "Decathlon India", 45000, "Bank", "Recorded", today],
    ["PAY-0002", "REQ-0003", "GPA Equestrian", 85000, "Bank", "Recorded", today],
    ["PAY-0003", "REQ-0004", "GPA Equestrian", 16000, "Bank", "Recorded", today]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.ORDER, [
    ["ORD-0001", "REQ-0002", "Decathlon India", "Riding Boots", 1, "Bangalore", today, "Placed"],
    ["ORD-0002", "REQ-0003", "GPA Equestrian", "Dressage Saddle", 1, "Farm", today, "Received"],
    ["ORD-0003", "REQ-0004", "GPA Equestrian", "Standard Helmet", 1, "Bangalore", today, "Received"]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.RECEIVED, [
    ["GRN-0001", "ORD-0002", "Dressage Saddle", 1, "Farm", "Farm Manager", "Farm Store B", today],
    ["GRN-0002", "ORD-0003", "Standard Helmet", 1, "Bangalore", "KE Admin", "BLR Rack 2", today]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.ISSUE, [
    ["ISS-0001", "REQ-0004", "Standard Helmet", 1, "Vihaan Mehta", "Bangalore", "KE Admin", today]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.TRANSFER, [
    ["TRF-0001", "KE-DEC-GLO-WIN01", "Winter Gloves", "Farm", "Bangalore", 2, "KE Admin", "KE Admin", today, "Completed", "Initial stock balance"]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.SAMPLE, [
    ["SMP-0001", "KE-DEC-GLO-WIN01", "Winter Gloves", "Farm", "Farm Trainer", "Demo fitting at clinic", today, today, "", "Out", "Returned check pending"]
  ]);

  appendRowsIfEmpty_(KE.SHEETS.PRICING, [
    ["KE-DEC-HEL-PRO1",  "Riding Helmet Pro", 3500, 4900, 5500, 28.57, "Decathlon India", 3500, today],
    ["KE-GPA-SAD-DRS01", "Dressage Saddle",  85000,99000,110000, 14.14,"GPA Equestrian", 85000, today],
    ["KE-DEC-BOO-RID02", "Riding Boots",      4500, 5990, 6500, 24.87, "Decathlon India",4500, today]
  ]);

  ensureAdvancedSampleData_();

  PropertiesService.getScriptProperties().setProperty("KE_SAMPLE_SEEDED", "1");
  return "Sample data loaded. Login: KE Admin / admin@123 — then open dashboard.";
}

function ensureAdvancedSampleData_() {
  ensureRowsByKey_(KE.SHEETS.STORAGE, "Storage Name", [
    ["STO-0001", "Farm Store A",  "Farm",      "Main farm equipment room",  "KE Admin", todayStr_()],
    ["STO-0002", "Farm Store B",  "Farm",      "Secondary saddle room",     "KE Admin", todayStr_()],
    ["STO-0003", "BLR Rack 1",   "Bangalore", "Riding gear rack – floor 1","KE Admin", todayStr_()],
    ["STO-0004", "BLR Rack 2",   "Bangalore", "Helmets shelf – floor 1",   "KE Admin", todayStr_()],
    ["STO-0005", "HYD Shelf 1",  "Hyderabad", "General merchandise shelf",  "KE Admin", todayStr_()],
    ["STO-0006", "Pune Cabinet", "Pune",      "Locked equipment cabinet",   "KE Admin", todayStr_()]
  ]);
  ensureRowsByKey_(KE.SHEETS.PRICING, "Item Code", [
    ["KE-DEC-HEL-PRO1",  "Riding Helmet Pro", 3500, 4900, 5500, 28.57, "Decathlon India", 3500, todayStr_()],
    ["KE-GPA-SAD-DRS01", "Dressage Saddle",  85000,99000,110000, 14.14,"GPA Equestrian", 85000, todayStr_()],
    ["KE-DEC-BOO-RID02", "Riding Boots",      4500, 5990, 6500, 24.87, "Decathlon India",4500, todayStr_()]
  ]);
  ensureRowsByKey_(KE.SHEETS.TRANSFER, "Transfer ID", [
    ["TRF-0001", "KE-DEC-GLO-WIN01", "Winter Gloves", "Farm", "Bangalore", 2, "KE Admin", "KE Admin", todayStr_(), "Completed", "Initial stock balance"]
  ]);
  ensureRowsByKey_(KE.SHEETS.SAMPLE, "Sample ID", [
    ["SMP-0001", "KE-DEC-GLO-WIN01", "Winter Gloves", "Farm", "Farm Trainer", "Demo fitting at clinic", todayStr_(), todayStr_(), "", "Out", "Returned check pending"]
  ]);
}

function ensureRowsByKey_(sheetName, keyColName, rows) {
  var sh = getSheet_(sheetName);
  var data = getSheetData_(sheetName);
  var keyCol = findCol_(data.headers, keyColName);
  if (keyCol < 0) {
    return 0;
  }
  var existing = {};
  for (var i = 0; i < data.rows.length; i++) {
    existing[normalize_(data.rows[i][keyCol])] = true;
  }
  var added = 0;
  rows.forEach(function (row) {
    var key = normalize_(row[keyCol]);
    if (!key || existing[key]) {
      return;
    }
    sh.appendRow(row);
    existing[key] = true;
    added++;
  });
  return added;
}

function appendRowsIfEmpty_(sheetName, rows) {
  var sh = getSheet_(sheetName);
  if (sh.getLastRow() > 1) {
    Logger.log("seedSampleData: skipped " + sheetName + " (already has data)");
    return false;
  }
  rows.forEach(function (row) {
    sh.appendRow(row);
  });
  Logger.log("seedSampleData: loaded " + rows.length + " rows into " + sheetName);
  return true;
}

/** Run from Apps Script to see row counts per tab (fixes “no data” debugging). */
function diagnoseInventorySheets() {
  var lines = [];
  Object.keys(KE.SHEETS).forEach(function (key) {
    var name = KE.SHEETS[key];
    try {
      var data = getSheetData_(name);
      lines.push(name + ": " + data.rows.length + " data row(s)");
    } catch (e) {
      lines.push(name + ": MISSING or error — " + e.message);
    }
  });
  var msg = lines.join("\n");
  Logger.log(msg);
  return msg;
}

/** Clears data rows (keeps headers) on all inventory sheets — use with care. */
function clearAllInventoryData() {
  Object.keys(SHEET_HEADERS_).forEach(function (name) {
    var sh = getSS_().getSheetByName(name);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last > 1) {
      sh.deleteRows(2, last - 1);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty("KE_SAMPLE_SEEDED");
  return "Cleared. Run seedSampleData() to reload demo rows.";
}
