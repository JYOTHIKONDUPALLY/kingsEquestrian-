/**
 * Seed demo rows for Version 2. Run after setupInventorySheets().
 */

function seedSampleData() {
  setupInventorySheets();
  var ss = getSS_();

  Object.keys(SHEET_HEADERS_).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow(), sh.getLastColumn()).clearContent();
    }
  });

  appendRow_(KE.SHEETS.USER, [
    "admin@ke.demo", "KE Admin", KE.ROLES.ADMIN, "Farm", "admin@123"
  ]);
  appendRow_(KE.SHEETS.USER, [
    "manager.blr@ke.demo", "BLR Manager", KE.ROLES.MANAGER, "Bangalore", "manager@123"
  ]);
  appendRow_(KE.SHEETS.USER, [
    "trainer.farm@ke.demo", "Farm Trainer", KE.ROLES.TRAINER, "Farm", "trainer@123"
  ]);

  appendRow_(KE.SHEETS.VENDOR, [
    "VEN-0001", "EquiSupply India", "EQS", "9876500001", "sales@eqs.demo", KE.LOCATION_ALL,
    "KE Admin", todayStr_()
  ]);
  appendRow_(KE.SHEETS.VENDOR, [
    "VEN-0002", "Bangalore Tack Co", "BTC", "9876500002", "btc@demo", "Bangalore",
    "KE Admin", todayStr_()
  ]);
  appendRow_(KE.SHEETS.VENDOR, [
    "VEN-0003", "Farm Feed Mart", "FFM", "9876500003", "feed@demo", "Farm",
    "KE Admin", todayStr_()
  ]);

  appendRow_(KE.SHEETS.ITEM, [
    "KE-EQS-HEL-RIDING", "Riding Helmet", "Helmet", "EquiSupply India", KE.LOCATION_ALL, 5,
    "KE Admin", todayStr_(), ""
  ]);
  appendRow_(KE.SHEETS.ITEM, [
    "KE-EQS-GLO-LEATH", "Leather Gloves", "Gloves", "EquiSupply India", KE.LOCATION_ALL, 10,
    "KE Admin", todayStr_(), ""
  ]);
  appendRow_(KE.SHEETS.ITEM, [
    "KE-BTC-BOO-SHORT", "Short Boots", "Boots", "Bangalore Tack Co", "Bangalore", 3,
    "KE Admin", todayStr_(), ""
  ]);
  appendRow_(KE.SHEETS.ITEM, [
    "KE-FFM-FEE-PELLET", "Feed Pellets 25kg", "Feed", "Farm Feed Mart", "Farm", 8,
    "KE Admin", todayStr_(), ""
  ]);

  appendRow_(KE.SHEETS.INVENTORY, [
    "KE-EQS-HEL-RIDING", "Riding Helmet", "Bangalore", 12, nowStr_(), "KE Admin"
  ]);
  appendRow_(KE.SHEETS.INVENTORY, [
    "KE-EQS-HEL-RIDING", "Riding Helmet", "Farm", 4, nowStr_(), "KE Admin"
  ]);
  appendRow_(KE.SHEETS.INVENTORY, [
    "KE-EQS-GLO-LEATH", "Leather Gloves", "Farm", 6, nowStr_(), "KE Admin"
  ]);
  appendRow_(KE.SHEETS.INVENTORY, [
    "KE-BTC-BOO-SHORT", "Short Boots", "Bangalore", 2, nowStr_(), "KE Admin"
  ]);
  appendRow_(KE.SHEETS.INVENTORY, [
    "KE-FFM-FEE-PELLET", "Feed Pellets 25kg", "Farm", 20, nowStr_(), "KE Admin"
  ]);

  appendRow_(KE.SHEETS.STUDENT, [
    "KE1001", "Aarav Sharma", "Ravi Sharma", "aarav@demo", "9000000001",
    "School", "3.5", "5", "A", "Bangalore", "KE Admin", todayStr_()
  ]);
  appendRow_(KE.SHEETS.STUDENT, [
    "KE1002", "Meera Iyer", "Suresh Iyer", "meera@demo", "9000000002",
    "School", "4.0", "6", "B", "Farm", "KE Admin", todayStr_()
  ]);
  appendRow_(KE.SHEETS.STUDENT, [
    "KE1003", "Rohan Patel", "Anita Patel", "rohan@demo", "9000000003",
    "School", "2.8", "4", "A", "Farm", "KE Admin", todayStr_()
  ]);
  appendRow_(KE.SHEETS.STUDENT, [
    "KE1004", "Aarav Sharma", "Vikram Sharma", "aarav.h@demo", "9000000004",
    "School", "3.2", "5", "C", "Hyderabad", "KE Admin", todayStr_()
  ]);

  appendRow_(KE.SHEETS.REQUEST, [
    "REQ-0001", todayStr_(), "EquiSupply India", "Riding Helmet", 5,
    "Farm", "Farm Trainer", KE.REQUEST_STATUS.PENDING, "Need for new batch"
  ]);

  appendRow_(KE.SHEETS.ACTIVITY, [
    nowStr_(), "Seed", "System", "—", "Sample data loaded", "Farm", "KE Admin", "Admin"
  ]);

  var props = PropertiesService.getScriptProperties();
  props.setProperty("V2_SEQ_VEN", "3");
  props.setProperty("V2_SEQ_REQ", "1");
  props.setProperty("V2_SEQ_ISS", "0");

  return {
    success: true,
    message: "Sample data loaded. Login: KE Admin / admin@123"
  };
}
