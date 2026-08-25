/**
 * Core operations – vendors, items, inventory, requests, issues, students.
 */

function generateID(prefix) {
  return generateID_(prefix || "ID");
}

// ─── Vendors ─────────────────────────────────────────────────

function addVendor(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = assertCan_("master", payload.location, prep.token);
  var name = normalize_(payload.vendorName);
  var code = normalize_(payload.code).toUpperCase();
  var loc = validateMasterLocation_(payload.location);
  if (!name || !code) {
    throw new Error("Vendor name and code are required.");
  }
  var id = generateID_("VEN");
  var by = actorName_(user);
  appendRow_(KE.SHEETS.VENDOR, [
    id, name, code,
    normalize_(payload.phone),
    normalize_(payload.email),
    loc,
    by,
    todayStr_()
  ]);
  logActivity_(user, "Added", "Vendor", id, name + " (" + code + ")", loc);
  return ok_("Vendor added.", { vendorId: id });
}

function updateVendor(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var vendorId = normalize_(payload.vendorId);
  if (!vendorId) {
    throw new Error("Vendor ID is required to edit.");
  }
  var found = findSheetRowById_(KE.SHEETS.VENDOR, "Vendor ID", vendorId);
  if (!found) {
    throw new Error("Vendor not found: " + vendorId);
  }
  var oldName = normalize_(found.obj["Vendor Name"]);
  var name = normalize_(payload.vendorName) || oldName;
  var code = normalize_(payload.code).toUpperCase() || normalize_(found.obj["Code"]);
  var loc = validateMasterLocation_(payload.location != null ? payload.location : found.obj["Location"]);
  var user = assertCan_("master", loc, prep.token);
  if (!name || !code) {
    throw new Error("Vendor name and code are required.");
  }
  setSheetRowFields_(KE.SHEETS.VENDOR, found.sheetRow, found.headers, {
    "Vendor Name": name,
    "Code": code,
    "Phone": normalize_(payload.phone),
    "Email": normalize_(payload.email),
    "Location": loc
  });
  // Keep item master vendor name in sync if renamed
  if (oldName && name !== oldName) {
    renameVendorOnItems_(oldName, name);
  }
  logActivity_(user, "Updated", "Vendor", vendorId, name + " (" + code + ")", loc);
  return ok_("Vendor updated.", { vendorId: vendorId });
}

function renameVendorOnItems_(oldName, newName) {
  var sh = getSheet_(KE.SHEETS.ITEM);
  var data = getSheetData_(KE.SHEETS.ITEM);
  var cVendor = findCol_(data.headers, "Vendor");
  if (cVendor < 0) {
    return;
  }
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cVendor]) === oldName) {
      sh.getRange(i + 2, cVendor + 1).setValue(newName);
    }
  }
}

function listVendors(token, viewLocation, payload) {
  var q = listQueryFromPayload_(payload || {}, token, viewLocation);
  var data = getSheetData_(KE.SHEETS.VENDOR);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", q.view);
  var objects = rows.map(function (row) {
    return rowToObject_(data.headers, row);
  }).reverse();
  objects = filterBySearch_(objects, q.search, [
    "Vendor ID", "Vendor Name", "Code", "Phone", "Email", "Location", "Added By"
  ]);
  objects = filterByDateField_(objects, "Date Added", q.dateFrom, q.dateTo);
  return paginateLatest_(objects, q.limit);
}

function getVendorCode_(vendorName) {
  var data = getSheetData_(KE.SHEETS.VENDOR);
  var cName = findCol_(data.headers, "Vendor Name");
  var cCode = findCol_(data.headers, "Code");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cName]) === vendorName) {
      return normalize_(data.rows[i][cCode]).toUpperCase() || "GEN";
    }
  }
  throw new Error("Vendor not found: " + vendorName);
}

function assertVendorExists_(vendorName) {
  var data = getSheetData_(KE.SHEETS.VENDOR);
  var cName = findCol_(data.headers, "Vendor Name");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cName]) === vendorName) {
      return;
    }
  }
  throw new Error("Vendor not found: " + vendorName);
}

// ─── Items ───────────────────────────────────────────────────

function addItem(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = assertCan_("master", payload.location, prep.token);
  var itemName = normalize_(payload.itemName);
  var category = normalize_(payload.category) || "General";
  var vendorName = normalize_(payload.vendor);
  var loc = validateMasterLocation_(payload.location);
  var minLevel = Number(payload.minLevel) || 0;
  if (!itemName || !vendorName) {
    throw new Error("Item name and vendor are required.");
  }
  assertVendorExists_(vendorName);
  var vendorCode = getVendorCode_(vendorName);
  var itemCode = normalize_(payload.itemCode);
  if (!itemCode) {
    itemCode = buildUniqueItemCode_(vendorCode, category, itemName);
  } else if (findItemByCode_(itemCode)) {
    throw new Error("Item code already exists: " + itemCode + ". Choose a different code or leave blank to auto-generate.");
  }
  var by = actorName_(user);
  var imageUrl = resolveItemImageUrl_(payload, "");
  appendRow_(KE.SHEETS.ITEM, [
    itemCode, itemName, category, vendorName, loc, minLevel, by, todayStr_(), imageUrl
  ]);
  logActivity_(user, "Added", "Item", itemCode, itemName + " @ " + loc, loc);
  return ok_("Item added.", { itemCode: itemCode });
}

function resolveItemImageUrl_(payload, existingUrl) {
  payload = payload || {};
  var imageUrl = normalize_(payload.imageUrl);
  if (payload.imageData) {
    imageUrl = saveItemImageToDrive_(payload.imageData, payload.imageName, payload.imageMimeType);
  }
  return imageUrl || normalize_(existingUrl) || "";
}

function saveItemImageToDrive_(dataUrl, fileName, mimeType) {
  dataUrl = normalize_(dataUrl);
  if (!dataUrl) {
    return "";
  }
  var m = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
  if (!m) {
    throw new Error("Invalid image. Use camera or gallery and try again.");
  }
  var type = normalize_(mimeType) || normalize_(m[1]) || "image/jpeg";
  var bytes = Utilities.base64Decode(m[2]);
  var ext = "jpg";
  if (type.indexOf("png") >= 0) ext = "png";
  else if (type.indexOf("webp") >= 0) ext = "webp";
  else if (type.indexOf("gif") >= 0) ext = "gif";
  var safeName = normalize_(fileName).replace(/[^\w.\-]/g, "_");
  if (!safeName) {
    safeName = "item-image-" + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyyMMdd-HHmmss"
    ) + "." + ext;
  }
  if (safeName.toLowerCase().indexOf("." + ext) < 0) {
    safeName += "." + ext;
  }
  try {
    var blob = Utilities.newBlob(bytes, type, safeName);
    var folder = getItemImagesFolder_();
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(file.getId()) + "&sz=w1000";
  } catch (e) {
    throw new Error(
      "Drive permission is required for image upload. " +
      "In Apps Script run authorizeImageUploadDriveAccess() once, allow access, then redeploy. " +
      (e && e.message ? e.message : e)
    );
  }
}

function getItemImagesFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("ITEM_IMAGES_FOLDER_ID");
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) { /* recreate */ }
  }
  var folder = DriveApp.createFolder("KE Track v2 – Item images");
  props.setProperty("ITEM_IMAGES_FOLDER_ID", folder.getId());
  return folder;
}

function authorizeImageUploadDriveAccess() {
  var folder = getItemImagesFolder_();
  return "Drive image upload authorized. Folder: " + folder.getName();
}

function updateItem(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var itemCode = normalize_(payload.itemCode);
  if (!itemCode) {
    throw new Error("Item Code is required to edit.");
  }
  var found = findSheetRowById_(KE.SHEETS.ITEM, "Item Code", itemCode);
  if (!found) {
    throw new Error("Item not found: " + itemCode);
  }
  var oldName = normalize_(found.obj["Item Name"]);
  var itemName = normalize_(payload.itemName) || oldName;
  var category = normalize_(payload.category) || normalize_(found.obj["Category"]) || "General";
  var vendorName = normalize_(payload.vendor) || normalize_(found.obj["Vendor"]);
  var loc = validateMasterLocation_(payload.location != null ? payload.location : found.obj["Location"]);
  var minLevel = payload.minLevel != null && payload.minLevel !== ""
    ? Number(payload.minLevel)
    : Number(found.obj["Min Level"]) || 0;
  var user = assertCan_("master", loc, prep.token);
  if (!itemName || !vendorName) {
    throw new Error("Item name and vendor are required.");
  }
  assertVendorExists_(vendorName);
  var imageUrl = resolveItemImageUrl_(payload, found.obj["Image URL"]);
  setSheetRowFields_(KE.SHEETS.ITEM, found.sheetRow, found.headers, {
    "Item Name": itemName,
    "Category": category,
    "Vendor": vendorName,
    "Location": loc,
    "Min Level": minLevel,
    "Image URL": imageUrl
  });
  if (oldName && itemName !== oldName) {
    syncInventoryItemName_(itemCode, itemName);
  }
  logActivity_(user, "Updated", "Item", itemCode, itemName + " @ " + loc, loc);
  return ok_("Item updated.", { itemCode: itemCode });
}

function syncInventoryItemName_(itemCode, itemName) {
  var sh = getSheet_(KE.SHEETS.INVENTORY);
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var cCode = findCol_(data.headers, "Item Code");
  var cName = findCol_(data.headers, "Item Name");
  if (cCode < 0 || cName < 0) {
    return;
  }
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cCode]) === itemCode) {
      sh.getRange(i + 2, cName + 1).setValue(itemName);
    }
  }
}

/**
 * Build a unique item code from vendor + category + name.
 * Uses up to 12 name chars; if still taken, appends -2, -3, …
 */
function buildUniqueItemCode_(vendorCode, category, itemName) {
  var cat = String(category || "GEN").replace(/[^A-Za-z0-9]/g, "").substring(0, 3).toUpperCase() || "GEN";
  var cleanName = String(itemName || "ITEM").replace(/[^A-Za-z0-9]/g, "").toUpperCase() || "ITEM";
  var namePart = cleanName.substring(0, 12);
  var base = "KE-" + vendorCode + "-" + cat + "-" + namePart;
  if (!findItemByCode_(base)) {
    return base;
  }
  // Same short prefix (e.g. SHORTSHOESNO2 vs SHORTSHOESNO3) — keep uniqueness
  var n = 2;
  while (n < 1000) {
    var candidate = base + "-" + n;
    if (!findItemByCode_(candidate)) {
      return candidate;
    }
    n++;
  }
  return base + "-" + Utilities.getUuid().replace(/-/g, "").substring(0, 6).toUpperCase();
}

function findItemByCode_(itemCode) {
  itemCode = normalize_(itemCode);
  var data = getSheetData_(KE.SHEETS.ITEM);
  var cCode = findCol_(data.headers, "Item Code");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cCode]) === itemCode) {
      return rowToObject_(data.headers, data.rows[i]);
    }
  }
  return null;
}

function findItemByNameOrCode_(item) {
  item = normalize_(item);
  var data = getSheetData_(KE.SHEETS.ITEM);
  var cCode = findCol_(data.headers, "Item Code");
  var cName = findCol_(data.headers, "Item Name");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cCode]) === item || normalize_(data.rows[i][cName]) === item) {
      return rowToObject_(data.headers, data.rows[i]);
    }
  }
  return null;
}

function listItems(token, viewLocation, payload) {
  var q = listQueryFromPayload_(payload || {}, token, viewLocation);
  var data = getSheetData_(KE.SHEETS.ITEM);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", q.view);
  var objects = rows.map(function (row) {
    return rowToObject_(data.headers, row);
  }).reverse();
  objects = filterBySearch_(objects, q.search, [
    "Item Code", "Item Name", "Category", "Vendor", "Location", "Added By"
  ]);
  objects = filterByDateField_(objects, "Date Added", q.dateFrom, q.dateTo);
  return paginateLatest_(objects, q.limit);
}

function itemAvailableAtLocation_(itemRow, location) {
  var itemLoc = resolveLocationName_(itemRow["Location"] || itemRow.Location);
  if (itemLoc === KE.LOCATION_ALL || isAllLocations_(itemLoc)) {
    return true;
  }
  return matchesLocation_(itemLoc, location);
}

// ─── Inventory ───────────────────────────────────────────────

function inventoryQtyCol_(headers) {
  return findCol_(headers, ["Qty Added", "Current Qty"]);
}

function inventoryDateCol_(headers) {
  return findCol_(headers, ["Date Added", "Last Updated"]);
}

function inventoryByCol_(headers) {
  return findCol_(headers, ["Added By", "Updated By"]);
}

function inventoryAddQtyFromObj_(obj) {
  return Number(obj["Qty Added"] != null && obj["Qty Added"] !== "" ? obj["Qty Added"] : obj["Current Qty"]) || 0;
}

function inventoryDateFromObj_(obj) {
  return obj["Date Added"] || obj["Last Updated"] || "";
}

function inventoryByFromObj_(obj) {
  return obj["Added By"] || obj["Updated By"] || "";
}

function normalizeInventoryAddObj_(obj) {
  obj["Qty Added"] = inventoryAddQtyFromObj_(obj);
  obj["Date Added"] = inventoryDateFromObj_(obj);
  obj["Added By"] = inventoryByFromObj_(obj);
  obj["Current Qty"] = obj["Qty Added"];
  obj["Last Updated"] = obj["Date Added"];
  obj["Updated By"] = obj["Added By"];
  return obj;
}

function issuedKey_(itemRef, location) {
  return String(itemRef || "").toLowerCase() + "||" + resolveLocationName_(location);
}

function getIssuedQtyMap_() {
  var data = getSheetData_(KE.SHEETS.ISSUE);
  var cItem = findCol_(data.headers, "Item");
  var cQty = findCol_(data.headers, "Qty");
  var cLoc = findCol_(data.headers, "Location");
  var map = {};
  if (cItem < 0 || cQty < 0 || cLoc < 0) {
    return map;
  }
  for (var i = 0; i < data.rows.length; i++) {
    var loc = resolveLocationName_(data.rows[i][cLoc]);
    var item = normalize_(data.rows[i][cItem]);
    if (!item) {
      continue;
    }
    var key = issuedKey_(item, loc);
    map[key] = (map[key] || 0) + (Number(data.rows[i][cQty]) || 0);
  }
  return map;
}

function issuedQtyForItem_(item, location, issuedMap) {
  if (!item) {
    return 0;
  }
  issuedMap = issuedMap || getIssuedQtyMap_();
  var loc = resolveLocationName_(location);
  var byName = issuedMap[issuedKey_(item["Item Name"], loc)] || 0;
  var byCode = issuedMap[issuedKey_(item["Item Code"], loc)] || 0;
  return Math.max(byName, byCode);
}

function getAddedQty_(itemCode, location) {
  itemCode = normalize_(itemCode);
  location = resolveLocationName_(location);
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var cCode = findCol_(data.headers, "Item Code");
  var cLoc = findCol_(data.headers, "Location");
  var cQty = inventoryQtyCol_(data.headers);
  var total = 0;
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cCode]) === itemCode &&
        resolveLocationName_(data.rows[i][cLoc]) === location) {
      total += Number(data.rows[i][cQty]) || 0;
    }
  }
  return total;
}

/**
 * One-time: old INVENTORY qty was on-hand (adds minus issues).
 * Convert to add-log so on-hand = SUM(adds) − SUM(issues).
 */
function migrateInventoryToAddLedger_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("KE_INVENTORY_ADD_LEDGER") === "1") {
    return;
  }
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var cCode = findCol_(data.headers, "Item Code");
  var cName = findCol_(data.headers, "Item Name");
  var cLoc = findCol_(data.headers, "Location");
  var cQty = inventoryQtyCol_(data.headers);
  if (cCode < 0 || cLoc < 0 || cQty < 0) {
    props.setProperty("KE_INVENTORY_ADD_LEDGER", "1");
    return;
  }
  var issuedMap = getIssuedQtyMap_();
  var firstRowByKey = {};
  var itemByKey = {};
  for (var i = 0; i < data.rows.length; i++) {
    var code = normalize_(data.rows[i][cCode]);
    var loc = resolveLocationName_(data.rows[i][cLoc]);
    if (!code || !loc) {
      continue;
    }
    var key = code + "||" + loc;
    if (firstRowByKey[key] == null) {
      firstRowByKey[key] = i;
      itemByKey[key] = {
        "Item Code": code,
        "Item Name": normalize_(data.rows[i][cName])
      };
    }
  }
  var sh = getSheet_(KE.SHEETS.INVENTORY);
  Object.keys(firstRowByKey).forEach(function (key) {
    var parts = key.split("||");
    var item = itemByKey[key];
    var issued = issuedQtyForItem_(item, parts[1], issuedMap);
    if (issued <= 0) {
      return;
    }
    var rowIdx = firstRowByKey[key];
    var newQty = (Number(data.rows[rowIdx][cQty]) || 0) + issued;
    sh.getRange(rowIdx + 2, cQty + 1).setValue(newQty);
  });
  props.setProperty("KE_INVENTORY_ADD_LEDGER", "1");
}

function listInventory(token, viewLocation, payload) {
  var q = listQueryFromPayload_(payload || {}, token, viewLocation);
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", q.view);
  var minMap = getItemMinLevelMap_();
  var issuedMap = getIssuedQtyMap_();
  var onHandCache = {};
  var objects = rows.map(function (row) {
    var obj = normalizeInventoryAddObj_(rowToObject_(data.headers, row));
    var code = normalize_(obj["Item Code"]);
    var loc = resolveLocationName_(obj["Location"]);
    var cacheKey = code + "||" + loc;
    if (onHandCache[cacheKey] == null) {
      var item = { "Item Code": code, "Item Name": obj["Item Name"] };
      onHandCache[cacheKey] = Math.max(0, getAddedQty_(code, loc) - issuedQtyForItem_(item, loc, issuedMap));
    }
    var onHand = onHandCache[cacheKey];
    var min = minMap[code] != null ? Number(minMap[code]) : 0;
    obj.onHand = onHand;
    obj.minLevel = min;
    obj.stockStatus = stockStatus_(onHand, min);
    return obj;
  });
  objects.sort(function (a, b) {
    return String(inventoryDateFromObj_(b)).localeCompare(String(inventoryDateFromObj_(a)));
  });
  objects = filterBySearch_(objects, q.search, [
    "Item Code", "Item Name", "Location", "Added By", "Updated By", "stockStatus"
  ]);
  objects = filterByDateField_(objects, "Date Added", q.dateFrom, q.dateTo);
  if (q.status) {
    objects = objects.filter(function (o) {
      return normalize_(o.stockStatus) === q.status;
    });
  }
  return paginateLatest_(objects, q.limit);
}

function getItemMinLevelMap_() {
  var data = getSheetData_(KE.SHEETS.ITEM);
  var cCode = findCol_(data.headers, "Item Code");
  var cMin = findCol_(data.headers, "Min Level");
  var map = {};
  for (var i = 0; i < data.rows.length; i++) {
    map[normalize_(data.rows[i][cCode])] = Number(data.rows[i][cMin]) || 0;
  }
  return map;
}

function stockStatus_(qty, min) {
  qty = Number(qty) || 0;
  min = Number(min) || 0;
  if (qty <= 0) {
    return "Out of stock";
  }
  if (min > 0 && qty <= min) {
    return "Low stock";
  }
  if (min > 0 && qty <= min * 1.5) {
    return "Near min";
  }
  return "Healthy";
}

function addInventoryQty(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var loc = validateSiteLocation_(payload.location);
  var user = assertCan_("inventory", loc, prep.token);
  var itemRef = normalize_(payload.itemCode || payload.item);
  var addQty = Number(payload.qty);
  if (!itemRef) {
    throw new Error("Item is required.");
  }
  if (!addQty || addQty < 1 || Math.floor(addQty) !== addQty) {
    throw new Error("Enter a whole number of 1 or more to add. Stock is not overwritten.");
  }
  var item = findItemByNameOrCode_(itemRef);
  if (!item) {
    throw new Error("Item not found: " + itemRef);
  }
  if (!itemAvailableAtLocation_(item, loc)) {
    throw new Error("Item " + item["Item Name"] + " is not available for " + loc + ".");
  }
  var code = item["Item Code"];
  var name = item["Item Name"];
  var by = actorName_(user);
  var newQty = appendInventoryAdd_(code, name, loc, addQty, by);
  logActivity_(user, "Added stock", "Inventory", code, name + " +" + addQty + " → " + newQty, loc);
  return ok_("Added " + addQty + ". On hand now: " + newQty + ".", {
    itemCode: code,
    location: loc,
    added: addQty,
    qty: newQty
  });
}

/** Kept for older UI calls — always adds; never overwrites current qty. */
function setInventoryQty(payload) {
  return addInventoryQty(payload);
}

function appendInventoryAdd_(itemCode, itemName, location, addQty, addedBy) {
  addQty = Number(addQty);
  if (!addQty || addQty < 1 || Math.floor(addQty) !== addQty) {
    throw new Error("Qty to add must be a whole number of 1 or more.");
  }
  location = validateSiteLocation_(location);
  appendRow_(KE.SHEETS.INVENTORY, [
    itemCode, itemName, location, addQty, nowStr_(), addedBy || ""
  ]);
  return getStockQty_(itemCode, location);
}

function adjustInventoryQty_(itemCodeOrName, location, delta, updatedBy) {
  var item = findItemByNameOrCode_(itemCodeOrName);
  if (!item) {
    throw new Error("Item not found: " + itemCodeOrName);
  }
  if (delta < 0) {
    throw new Error("Inventory rows are add-only. Issues reduce on-hand without editing add records.");
  }
  return appendInventoryAdd_(item["Item Code"], item["Item Name"], location, delta, updatedBy);
}

function getStockQty_(itemCodeOrName, location) {
  var item = findItemByNameOrCode_(itemCodeOrName);
  if (!item) {
    return 0;
  }
  location = resolveLocationName_(location);
  var added = getAddedQty_(item["Item Code"], location);
  var issued = issuedQtyForItem_(item, location);
  return Math.max(0, added - issued);
}

// ─── Requests (Excel tab: REQUEST_REGISTER) ──────────────────

function createRequest(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = assertCan_("request", payload.location, prep.token);
  var vendor = normalize_(payload.vendorName || payload.vendor);
  var itemRef = normalize_(payload.item);
  var qty = Number(payload.qty);
  var loc = validateSiteLocation_(payload.location);
  var notes = normalize_(payload.notes);
  if (!vendor || !itemRef || !qty || qty <= 0) {
    throw new Error("Vendor, item, and quantity are required.");
  }
  assertVendorExists_(vendor);
  var item = findItemByNameOrCode_(itemRef);
  if (!item) {
    throw new Error("Item not found: " + itemRef);
  }
  if (!itemAvailableAtLocation_(item, loc)) {
    throw new Error("Item is not available for location " + loc + ".");
  }
  var id = generateID_("REQ");
  appendRow_(KE.SHEETS.REQUEST, [
    id, todayStr_(), vendor, item["Item Name"], qty, loc,
    actorName_(user),
    KE.REQUEST_STATUS.PENDING,
    notes
  ]);
  logActivity_(user, "Created", "Request", id, item["Item Name"] + " x" + qty, loc);
  return ok_("Request added to REQUEST_REGISTER.", { requestId: id });
}

function updateRequestStatus(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var requestId = normalize_(payload.requestId);
  var status = normalize_(payload.status);
  if (!requestId || !status) {
    throw new Error("Request ID and status are required.");
  }
  var allowed = [
    KE.REQUEST_STATUS.PENDING,
    KE.REQUEST_STATUS.ORDERED,
    KE.REQUEST_STATUS.RECEIVED,
    KE.REQUEST_STATUS.CANCELLED
  ];
  if (allowed.indexOf(status) < 0) {
    throw new Error("Invalid status. Use: " + allowed.join(", "));
  }
  var found = findRequestRow_(requestId);
  var user = assertCan_("master", found.location, prep.token);
  var sh = getSheet_(KE.SHEETS.REQUEST);
  var cStatus = findCol_(found.headers, "Status");
  if (cStatus < 0) {
    throw new Error("Status column missing on REQUEST_REGISTER.");
  }
  sh.getRange(found.sheetRow, cStatus + 1).setValue(status);

  if (status === KE.REQUEST_STATUS.RECEIVED && payload.addStock) {
    var receiveQty = Number(payload.receiveQty) || found.qty;
    adjustInventoryQty_(found.item, found.location, receiveQty, actorName_(user));
    logActivity_(user, "Received + stock", "Request", requestId,
      found.item + " +" + receiveQty, found.location);
  } else {
    logActivity_(user, "Status → " + status, "Request", requestId, found.item, found.location);
  }
  return ok_("Request status updated.", { requestId: requestId, status: status });
}

function findRequestRow_(requestId) {
  requestId = normalize_(requestId);
  var data = getSheetData_(KE.SHEETS.REQUEST);
  var cId = findCol_(data.headers, "Request ID");
  var cItem = findCol_(data.headers, "Item");
  var cLoc = findCol_(data.headers, "Location");
  var cQty = findCol_(data.headers, "Qty");
  var cStatus = findCol_(data.headers, "Status");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cId]) === requestId) {
      return {
        sheetRow: i + 2,
        headers: data.headers,
        row: data.rows[i],
        requestId: requestId,
        item: normalize_(data.rows[i][cItem]),
        location: normalize_(data.rows[i][cLoc]),
        qty: Number(data.rows[i][cQty]) || 0,
        status: cStatus >= 0 ? normalize_(data.rows[i][cStatus]) : ""
      };
    }
  }
  throw new Error("Request not found: " + requestId);
}

function listRequests(token, viewLocation, payload) {
  var q = listQueryFromPayload_(payload || {}, token, viewLocation);
  var data = getSheetData_(KE.SHEETS.REQUEST);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", q.view);
  var objects = rows.map(function (row) {
    return rowToObject_(data.headers, row);
  }).reverse();
  objects = filterBySearch_(objects, q.search, [
    "Request ID", "Vendor", "Item", "Location", "Requested By", "Status", "Notes"
  ]);
  objects = filterByDateField_(objects, "Date", q.dateFrom, q.dateTo);
  objects = filterByStatus_(objects, "Status", q.status);
  return paginateLatest_(objects, q.limit);
}

// ─── Students ────────────────────────────────────────────────

function studentIdFromPayload_(payload) {
  return normalize_(payload.studentId || payload.keNumber || payload.Student_ID);
}

function studentNameFromRow_(row) {
  return normalize_(row["Student_Name"] || row["Student Name"] || "");
}

function studentIdFromRow_(row) {
  return normalize_(row["Student_ID"] || row["KE Number"] || "");
}

function addStudent(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var loc = payload.location ? validateSiteLocation_(payload.location) : "";
  var user = assertCan_("issue", loc || null, prep.token);
  var studentId = studentIdFromPayload_(payload);
  var studentName = normalize_(payload.studentName || payload.Student_Name);
  if (!studentId || !studentName) {
    throw new Error("Student_ID and Student_Name are required.");
  }
  if (findStudentById_(studentId)) {
    throw new Error("Student_ID already exists: " + studentId);
  }
  if (!loc) {
    loc = resolveLocationName_(user.location) || KE.DEFAULT_ADMIN_LOCATION;
    if (loc === KE.LOCATION_ALL) {
      loc = KE.DEFAULT_ADMIN_LOCATION;
    }
  }
  var by = actorName_(user);
  appendRow_(KE.SHEETS.STUDENT, [
    studentId,
    studentName,
    normalize_(payload.parentName || payload.Parent_Name),
    normalize_(payload.email || payload.Email),
    normalize_(payload.phone || payload.Phone),
    normalize_(payload.program || payload.Program),
    normalize_(payload.skillRidingAvg || payload.Skill_Riding_Avg),
    normalize_(payload.grade || payload.Grade),
    normalize_(payload.section || payload.Section),
    loc,
    by,
    todayStr_()
  ]);
  logActivity_(user, "Added", "Student", studentId, studentName, loc);
  return ok_("Student added.", { studentId: studentId, studentName: studentName });
}

function updateStudent(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var studentId = studentIdFromPayload_(payload);
  if (!studentId) {
    throw new Error("Student_ID is required to edit.");
  }
  var found = findSheetRowById_(KE.SHEETS.STUDENT, ["Student_ID", "KE Number", "Student ID"], studentId);
  if (!found) {
    throw new Error("Student not found: " + studentId);
  }
  var loc = payload.location
    ? validateSiteLocation_(payload.location)
    : validateSiteLocation_(found.obj["Location"] || KE.DEFAULT_ADMIN_LOCATION);
  var user = assertCan_("issue", loc, prep.token);
  var studentName = normalize_(payload.studentName || payload.Student_Name) || studentNameFromRow_(found.obj);
  if (!studentName) {
    throw new Error("Student_Name is required.");
  }
  setSheetRowFields_(KE.SHEETS.STUDENT, found.sheetRow, found.headers, {
    "Student_Name": studentName,
    "Parent_Name": normalize_(payload.parentName || payload.Parent_Name),
    "Email": normalize_(payload.email || payload.Email),
    "Phone": normalize_(payload.phone || payload.Phone),
    "Program": normalize_(payload.program || payload.Program),
    "Skill_Riding_Avg": normalize_(payload.skillRidingAvg || payload.Skill_Riding_Avg),
    "Grade": normalize_(payload.grade || payload.Grade),
    "Section": normalize_(payload.section || payload.Section),
    "Location": loc
  });
  logActivity_(user, "Updated", "Student", studentId, studentName, loc);
  return ok_("Student updated.", { studentId: studentId, studentName: studentName });
}

function findStudentById_(studentId) {
  studentId = normalize_(studentId);
  if (!studentId) {
    return null;
  }
  var data = getSheetData_(KE.SHEETS.STUDENT);
  var cId = findCol_(data.headers, ["Student_ID", "KE Number", "Student ID"]);
  if (cId < 0) {
    return null;
  }
  var want = studentId.toUpperCase();
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cId]).toUpperCase() === want) {
      return rowToObject_(data.headers, data.rows[i]);
    }
  }
  return null;
}

/** @deprecated use findStudentById_ */
function findStudentByKe_(keNumber) {
  return findStudentById_(keNumber);
}

function listStudents(token, viewLocation, payload) {
  var q = listQueryFromPayload_(payload || {}, token, viewLocation);
  var data = getSheetData_(KE.SHEETS.STUDENT);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", q.view);
  var objects = rows.map(function (row) {
    return rowToObject_(data.headers, row);
  });
  // Prefer newest added first when Date Added exists
  objects.sort(function (a, b) {
    var da = String(b["Date Added"] || "");
    var db = String(a["Date Added"] || "");
    if (da !== db) {
      return da.localeCompare(db);
    }
    return studentNameFromRow_(a).localeCompare(studentNameFromRow_(b));
  });
  objects = filterBySearch_(objects, q.search, [
    "Student_ID", "Student_Name", "Parent_Name", "Email", "Phone",
    "Program", "Grade", "Section", "Location", "Added By"
  ]);
  objects = filterByDateField_(objects, "Date Added", q.dateFrom, q.dateTo);
  return paginateLatest_(objects, q.limit);
}

function appendStudentIfMissing_(studentId, studentName, loc, payload, addedBy) {
  if (findStudentById_(studentId)) {
    return;
  }
  appendRow_(KE.SHEETS.STUDENT, [
    studentId,
    studentName,
    normalize_(payload && (payload.parentName || payload.Parent_Name)),
    normalize_(payload && (payload.email || payload.Email)),
    normalize_(payload && (payload.phone || payload.Phone)),
    normalize_(payload && (payload.program || payload.Program)),
    normalize_(payload && (payload.skillRidingAvg || payload.Skill_Riding_Avg)),
    normalize_(payload && (payload.grade || payload.Grade)),
    normalize_(payload && (payload.section || payload.Section)),
    loc,
    addedBy || "",
    todayStr_()
  ]);
}

// ─── Issues ──────────────────────────────────────────────────

function issueItem(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var loc = validateSiteLocation_(payload.location);
  var user = assertCan_("issue", loc, prep.token);
  var itemRef = normalize_(payload.item);
  var qty = Number(payload.qty);
  var studentName = normalize_(payload.studentName || payload.Student_Name);
  var keNumber = studentIdFromPayload_(payload);

  if (!itemRef || !qty || qty <= 0) {
    throw new Error("Item and quantity are required.");
  }

  // Resolve name from master when Student_ID selected from dropdown
  if (keNumber && !studentName) {
    var master = findStudentById_(keNumber);
    if (master) {
      studentName = studentNameFromRow_(master);
    }
  }

  if (!studentName || !keNumber) {
    throw new Error("Select a student (Name - Student_ID) from the dropdown.");
  }

  var stock = getStockQty_(itemRef, loc);
  if (stock < qty) {
    throw new Error("Cannot issue: stock (" + stock + ") is less than qty (" + qty + ").");
  }

  var item = findItemByNameOrCode_(itemRef);

  var issueId = generateID_("ISS");
  appendRow_(KE.SHEETS.ISSUE, [
    issueId,
    item["Item Name"],
    qty,
    studentName,
    keNumber,
    loc,
    actorName_(user),
    todayStr_()
  ]);
  logActivity_(user, "Issued", "Issue", issueId,
    item["Item Name"] + " x" + qty + " → " + studentName + " (" + keNumber + ")", loc);

  return ok_("Item issued.", {
    issueId: issueId,
    studentName: studentName,
    keNumber: keNumber,
    studentId: keNumber,
    remainingQty: stock - qty
  });
}

function listIssues(token, viewLocation, payload) {
  var q = listQueryFromPayload_(payload || {}, token, viewLocation);
  var data = getSheetData_(KE.SHEETS.ISSUE);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", q.view);
  var objects = rows.map(function (row) {
    return rowToObject_(data.headers, row);
  }).reverse();
  objects = filterBySearch_(objects, q.search, [
    "Issue ID", "Item", "Student Name", "KE Number", "Location", "Issued By"
  ]);
  objects = filterByDateField_(objects, "Date", q.dateFrom, q.dateTo);
  return paginateLatest_(objects, q.limit);
}

// ─── Dashboard summary / low stock ───────────────────────────

function getDashboardSummary_(viewLocation, user) {
  var inventory = getInventoryRowsForView_(viewLocation);
  var lowStock = inventory.filter(function (r) {
    return r.stockStatus === "Low stock" || r.stockStatus === "Out of stock";
  });
  var requests = getRequestRowsForView_(viewLocation);
  var pending = requests.filter(function (r) {
    return r["Status"] === KE.REQUEST_STATUS.PENDING;
  });
  var issuesData = getSheetData_(KE.SHEETS.ISSUE);
  var issueRows = filterRowsByLocationColumn_(
    issuesData.rows, issuesData.headers, "Location", viewLocation
  ).map(function (row) {
    return rowToObject_(issuesData.headers, row);
  }).reverse();
  var today = todayStr_();
  var issuedToday = issueRows.filter(function (r) {
    return String(r["Date"] || "").indexOf(today) === 0;
  });
  var totalQty = 0;
  inventory.forEach(function (r) {
    totalQty += Number(r["Current Qty"]) || 0;
  });

  var byLocation = {};
  KE.LOCATIONS.forEach(function (loc) {
    byLocation[loc] = { rows: 0, qty: 0, low: 0 };
  });
  inventory.forEach(function (r) {
    var loc = resolveLocationName_(r["Location"]);
    if (!byLocation[loc]) {
      byLocation[loc] = { rows: 0, qty: 0, low: 0 };
    }
    byLocation[loc].rows += 1;
    byLocation[loc].qty += Number(r["Current Qty"]) || 0;
    if (r.stockStatus === "Low stock" || r.stockStatus === "Out of stock") {
      byLocation[loc].low += 1;
    }
  });

  var recentActivity = getRecentActivityRows_(viewLocation, 12);

  return {
    totalItems: inventory.length,
    totalQty: totalQty,
    lowStockCount: lowStock.length,
    lowStockItems: lowStock.slice(0, 10),
    pendingRequests: pending.length,
    totalRequests: requests.length,
    totalIssues: issueRows.length,
    issuedToday: issuedToday.length,
    recentIssues: issueRows.slice(0, 8),
    recentActivity: recentActivity,
    locationBreakdown: KE.LOCATIONS.map(function (loc) {
      var b = byLocation[loc] || { rows: 0, qty: 0, low: 0 };
      return { location: loc, rows: b.rows, qty: b.qty, low: b.low };
    }),
    stockByItem: buildStockByItem_(inventory),
    viewLabel: isAllLocations_(viewLocation) ? "All Locations" : viewLocation
  };
}

function buildStockByItem_(inventory) {
  var map = {};
  (inventory || []).forEach(function (r) {
    var code = normalize_(r["Item Code"]) || normalize_(r["Item Name"]);
    if (!code) {
      return;
    }
    if (!map[code]) {
      var locs = {};
      KE.LOCATIONS.forEach(function (loc) { locs[loc] = 0; });
      map[code] = {
        itemCode: r["Item Code"] || code,
        itemName: r["Item Name"] || code,
        totalQty: 0,
        locations: locs
      };
    }
    var qty = Number(r["Current Qty"]) || 0;
    map[code].totalQty += qty;
    var loc = resolveLocationName_(r["Location"]);
    if (loc) {
      map[code].locations[loc] = (map[code].locations[loc] || 0) + qty;
    }
  });
  return Object.keys(map).map(function (k) { return map[k]; }).sort(function (a, b) {
    var byName = String(a.itemName).localeCompare(String(b.itemName));
    if (byName !== 0) {
      return byName;
    }
    return String(a.itemCode).localeCompare(String(b.itemCode));
  });
}

function getRecentActivityRows_(viewLocation, limit) {
  limit = limit || 10;
  try {
    var data = getSheetData_(KE.SHEETS.ACTIVITY);
    var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", viewLocation);
    return rows.map(function (row) {
      return rowToObject_(data.headers, row);
    }).reverse().slice(0, limit);
  } catch (e) {
    return [];
  }
}

function getInventoryRowsForView_(viewLocation) {
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", viewLocation);
  var minMap = getItemMinLevelMap_();
  var issuedMap = getIssuedQtyMap_();
  var groups = {};
  rows.forEach(function (row) {
    var obj = normalizeInventoryAddObj_(rowToObject_(data.headers, row));
    var code = normalize_(obj["Item Code"]);
    var loc = resolveLocationName_(obj["Location"]);
    if (!code || !loc) {
      return;
    }
    var key = code + "||" + loc;
    if (!groups[key]) {
      groups[key] = {
        "Item Code": obj["Item Code"],
        "Item Name": obj["Item Name"],
        "Location": loc,
        added: 0
      };
    }
    groups[key].added += inventoryAddQtyFromObj_(obj);
    if (obj["Item Name"]) {
      groups[key]["Item Name"] = obj["Item Name"];
    }
  });
  return Object.keys(groups).map(function (key) {
    var g = groups[key];
    var item = { "Item Code": g["Item Code"], "Item Name": g["Item Name"] };
    var onHand = Math.max(0, g.added - issuedQtyForItem_(item, g["Location"], issuedMap));
    var min = minMap[normalize_(g["Item Code"])] != null
      ? Number(minMap[normalize_(g["Item Code"])]) : 0;
    g["Current Qty"] = onHand;
    g["Qty Added"] = g.added;
    g.minLevel = min;
    g.stockStatus = stockStatus_(onHand, min);
    return g;
  });
}

function getRequestRowsForView_(viewLocation) {
  var data = getSheetData_(KE.SHEETS.REQUEST);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", viewLocation);
  return rows.map(function (row) {
    return rowToObject_(data.headers, row);
  });
}

function getMasterLists_(viewLocation) {
  var vendors = listVendorNamesForView_(viewLocation);
  var items = listItemOptionsForView_(viewLocation);
  var students = listStudentOptionsForView_(viewLocation);
  var openRequests = getRequestRowsForView_(viewLocation).filter(function (r) {
    return r["Status"] !== KE.REQUEST_STATUS.CANCELLED;
  }).map(function (r) {
    return {
      id: r["Request ID"],
      label: r["Request ID"] + " – " + r["Item"] + " (" + r["Status"] + ")"
    };
  });
  return {
    vendors: vendors,
    items: items,
    students: students,
    requests: openRequests,
    locations: KE.LOCATIONS,
    masterLocations: [KE.LOCATION_ALL].concat(KE.LOCATIONS),
    requestStatuses: [
      KE.REQUEST_STATUS.PENDING,
      KE.REQUEST_STATUS.ORDERED,
      KE.REQUEST_STATUS.RECEIVED,
      KE.REQUEST_STATUS.CANCELLED
    ]
  };
}

function listVendorNamesForView_(viewLocation) {
  var data = getSheetData_(KE.SHEETS.VENDOR);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", viewLocation);
  var cName = findCol_(data.headers, "Vendor Name");
  return uniqueSorted_(rows.map(function (r) { return r[cName]; }));
}

function listItemOptionsForView_(viewLocation) {
  var data = getSheetData_(KE.SHEETS.ITEM);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", viewLocation);
  return rows.map(function (row) {
    var o = rowToObject_(data.headers, row);
    return {
      code: o["Item Code"],
      name: o["Item Name"],
      label: o["Item Name"] + " (" + o["Item Code"] + ")",
      vendor: o["Vendor"],
      location: o["Location"],
      minLevel: o["Min Level"],
      imageUrl: o["Image URL"] || ""
    };
  }).sort(function (a, b) {
    return a.name.localeCompare(b.name);
  });
}

function listStudentOptionsForView_(viewLocation) {
  var data = getSheetData_(KE.SHEETS.STUDENT);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", viewLocation);
  return rows.map(function (row) {
    var o = rowToObject_(data.headers, row);
    var id = studentIdFromRow_(o);
    var name = studentNameFromRow_(o);
    return {
      value: id,
      studentId: id,
      keNumber: id,
      studentName: name,
      location: normalize_(o["Location"]),
      label: name + " - " + id
    };
  }).sort(function (a, b) {
    var byName = a.studentName.localeCompare(b.studentName);
    if (byName !== 0) {
      return byName;
    }
    return a.studentId.localeCompare(b.studentId);
  });
}
