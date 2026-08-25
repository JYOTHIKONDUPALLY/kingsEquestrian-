/**
 * Core inventory operations – masters, requests, payments, orders, receive, issue
 */

function generateID(prefix) {
  return generateID_(prefix || "ID");
}

function prepPayload_(payload) {
  payload = payload || {};
  var token = tokenFromPayload_(payload);
  var viewLocation = payload.viewLocation != null ? payload.viewLocation : "";
  delete payload.viewLocation;
  payload = enforcePayloadLocation_(payload, token);
  return { token: token, viewLocation: viewLocation, payload: payload };
}

// ─── Masters ─────────────────────────────────────────────────

function addVendor(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  assertCan_("master", payload.location, prep.token);
  var name = normalize_(payload.vendorName);
  var code = normalize_(payload.code).toUpperCase();
  var loc = normalize_(payload.location);
  if (!name || !code) {
    throw new Error("Vendor name and code are required.");
  }
  validateLocation_(loc);
  var id = generateID_("VEN");
  appendRow_(KE.SHEETS.VENDOR, [
    id, name, code,
    normalize_(payload.phone),
    normalize_(payload.email),
    loc
  ]);
  return ok_("Vendor added.", { vendorId: id });
}

function addCustomer(payload) {
  throw new Error("Customer master is disabled. Use vendor procurement requests instead.");
}

function addStudent(payload) {
  throw new Error("Student master is disabled. Use vendor procurement requests instead.");
}

function addNewItem(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = assertCan_("master", payload.location, prep.token);
  var itemName = normalize_(payload.itemName);
  var category = normalize_(payload.category);
  var model = normalize_(payload.model);
  var vendorName = normalize_(payload.vendor);
  var loc = normalize_(payload.location);
  var qty = Number(payload.qty) || 0;
  var minLevel = Number(payload.minLevel) || 0;
  var storage = normalize_(payload.storageLocation);
  var inventoryType = normalize_(payload.inventoryType) || KE.DEFAULT_INVENTORY_TYPE;
  var allocatedTo = normalize_(payload.allocatedTo);
  var reservedQty = Number(payload.reservedQty) || 0;
  var imageUrl = normalize_(payload.imageUrl);
  // Optional: upload selected camera/gallery image to Drive and store generated URL.
  if (payload.imageData) {
    imageUrl = saveItemImageToDrive_(payload.imageData, payload.imageName, payload.imageMimeType);
  }
  // Cost price is admin-only input. Silently ignore if non-admin tries to set it.
  var costPrice = 0;
  if (canViewFinancials_(user)) {
    costPrice = Number(payload.costPrice) || 0;
  }
  if (!itemName || !category || !model || !vendorName) {
    throw new Error("Item name, category, model, and vendor are required.");
  }
  if (KE.INVENTORY_TYPES.indexOf(inventoryType) < 0) {
    inventoryType = KE.DEFAULT_INVENTORY_TYPE;
  }
  validateLocation_(loc);
  var vendorCode = getVendorCode_(vendorName);
  var itemCode = buildItemCode_(vendorCode, category, model);
  appendRow_(KE.SHEETS.INVENTORY, [
    itemCode, itemName, category, model, vendorName, loc,
    qty, minLevel, storage, nowStr_(),
    inventoryType, allocatedTo, reservedQty, costPrice, nowStr_(), imageUrl
  ]);
  return ok_("Item added.", { itemCode: itemCode });
}

/**
 * Saves base64 image data to Google Drive and returns a public thumbnail URL.
 * Expected input format:
 *   data:image/jpeg;base64,/9j/4AAQSk...
 */
function saveItemImageToDrive_(dataUrl, fileName, mimeType) {
  dataUrl = normalize_(dataUrl);
  if (!dataUrl) {
    return "";
  }
  var m = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
  if (!m) {
    throw new Error("Invalid image payload.");
  }
  var type = normalize_(mimeType) || normalize_(m[1]) || "image/jpeg";
  var b64 = m[2];
  var bytes = Utilities.base64Decode(b64);
  var ext = "jpg";
  if (type.indexOf("png") >= 0) ext = "png";
  else if (type.indexOf("webp") >= 0) ext = "webp";
  else if (type.indexOf("gif") >= 0) ext = "gif";
  var safeName = normalize_(fileName).replace(/[^\w.\-]/g, "_");
  if (!safeName) {
    safeName = "item-image-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + "." + ext;
  }
  if (safeName.toLowerCase().indexOf("." + ext) < 0) {
    safeName += "." + ext;
  }
  try {
    var blob = Utilities.newBlob(bytes, type, safeName);
    var file = DriveApp.createFile(blob);
    // So thumbnails can render for app users via shared link.
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var id = file.getId();
    return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(id) + "&sz=w1000";
  } catch (e) {
    throw new Error(
      "Drive permission is required for image upload. " +
      "Run authorizeImageUploadDriveAccess() once from Apps Script editor, allow permissions, then redeploy web app. " +
      "Original error: " + (e && e.message ? e.message : e)
    );
  }
}

/**
 * One-time auth helper:
 * Run this manually from Apps Script editor to grant Drive scope used
 * by saveItemImageToDrive_().
 */
function authorizeImageUploadDriveAccess() {
  var folder = DriveApp.getRootFolder();
  return "Drive access authorized. Root folder: " + folder.getName();
}

function buildItemCode_(vendorCode, category, model) {
  var cat = category.replace(/[^A-Za-z0-9]/g, "").substring(0, 3).toUpperCase();
  var mod = model.replace(/[^A-Za-z0-9]/g, "").substring(0, 6).toUpperCase();
  return "KE-" + vendorCode + "-" + cat + "-" + mod;
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
  throw new Error("Vendor not found in Vendor Master: " + vendorName);
}

function assertMasterExists_(sheet, colName, value) {
  var data = getSheetData_(sheet);
  var col = findCol_(data.headers, colName);
  if (col < 0) {
    return;
  }
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][col]) === value) {
      return;
    }
  }
  throw new Error("Value not found in " + sheet + ": " + value);
}

// ─── Request lifecycle ─────────────────────────────────────────

function createRequest(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = assertCan_("request", payload.location, prep.token);
  var vendor = normalize_(payload.vendorName || payload.vendor);
  var item = normalize_(payload.item);
  var qty = Number(payload.qty);
  var loc = normalize_(payload.location);
  if (!vendor || !item || !qty || qty <= 0) {
    throw new Error("Vendor, item, and quantity are required.");
  }
  validateLocation_(loc);
  assertMasterExists_(KE.SHEETS.VENDOR, "Vendor Name", vendor);
  assertItemExists_(item, loc);
  var id = generateID_("REQ");
  appendRow_(KE.SHEETS.REQUEST, [
    id, todayStr_(), vendor, item, qty, loc,
    user.name || user.email,
    "Unpaid",
    KE.REQUEST_STATUS.PENDING
  ]);
  return ok_("Procurement request created.", { requestId: id });
}

/**
 * Increment / decrement the "Reserved Qty" column for the matching
 * INVENTORY_MASTER row. Used by createRequest (positive delta) and
 * issueItem / cancelRequest (negative delta) so available-for-sale
 * is always accurate.
 */
function adjustReservedQty_(itemNameOrCode, location, delta) {
  if (!delta) {
    return;
  }
  var sh = getSheet_(KE.SHEETS.INVENTORY);
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var headers = data.headers;
  var cCode = findCol_(headers, "Item Code");
  var cName = findCol_(headers, "Item Name");
  var cLoc = findCol_(headers, "Location");
  var cResv = findCol_(headers, "Reserved Qty");
  if (cResv < 0) {
    return;
  }
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var match = (normalize_(row[cCode]) === itemNameOrCode) ||
                (normalize_(row[cName]) === itemNameOrCode);
    if (match && normalize_(row[cLoc]) === location) {
      var cur = Number(row[cResv]) || 0;
      var next = Math.max(0, cur + delta);
      sh.getRange(i + 2, cResv + 1).setValue(next);
      return next;
    }
  }
}

function cancelRequest(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var requestId = normalize_(payload.requestId);
  if (!requestId) {
    throw new Error("Request ID is required.");
  }
  var req = findRequestRow_(requestId);
  assertRecordLocationAccess_(prep.token, req.location, prep.viewLocation);
  assertCan_("request", req.location, prep.token);
  if (req.status === KE.REQUEST_STATUS.COMPLETED ||
      req.status === KE.REQUEST_STATUS.ISSUED) {
    throw new Error("Cannot cancel a completed/issued request.");
  }
  if (req.status === KE.REQUEST_STATUS.CANCELLED) {
    return ok_("Already cancelled.", { requestId: requestId });
  }
  updateRequestFields_(req.sheetRow, req.headers, {
    "Status": KE.REQUEST_STATUS.CANCELLED
  });
  return ok_("Request cancelled.", { requestId: requestId });
}

function assertItemExists_(itemNameOrCode, location) {
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var cCode = findCol_(data.headers, "Item Code");
  var cName = findCol_(data.headers, "Item Name");
  var cLoc = findCol_(data.headers, "Location");
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var match = normalize_(row[cCode]) === itemNameOrCode ||
      normalize_(row[cName]) === itemNameOrCode;
    if (match && (!location || normalize_(row[cLoc]) === location)) {
      return;
    }
  }
  throw new Error("Item not found in inventory master: " + itemNameOrCode);
}

function findRequestRow_(requestId) {
  var data = getSheetData_(KE.SHEETS.REQUEST);
  var cId = findCol_(data.headers, "Request ID");
  var cStatus = findCol_(data.headers, "Status");
  var cPay = findCol_(data.headers, "Payment Status");
  var cLoc = findCol_(data.headers, "Location");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cId]) === requestId) {
      return {
        sheetRow: i + 2,
        status: normalize_(data.rows[i][cStatus]),
        paymentStatus: normalize_(data.rows[i][cPay]),
        location: normalize_(data.rows[i][cLoc]),
        row: data.rows[i],
        headers: data.headers
      };
    }
  }
  throw new Error("Request not found: " + requestId);
}

function updateRequestFields_(sheetRow, headers, updates) {
  var sh = getSheet_(KE.SHEETS.REQUEST);
  Object.keys(updates).forEach(function (field) {
    var col = findCol_(headers, field) + 1;
    if (col > 0) {
      sh.getRange(sheetRow, col).setValue(updates[field]);
    }
  });
}

function readRequestField_(req, fieldNames) {
  var col = findCol_(req.headers, fieldNames);
  return col >= 0 ? normalize_(req.row[col]) : "";
}

// NO PAYMENT → NO APPROVAL
function recordPayment(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  assertCan_("payment", null, prep.token);
  var requestId = normalize_(payload.requestId);
  var amount = Number(payload.amount);
  var mode = normalize_(payload.mode) || "Bank";
  if (!requestId || !amount || amount <= 0) {
    throw new Error("Request ID and amount are required.");
  }
  var req = findRequestRow_(requestId);
  assertRecordLocationAccess_(prep.token, req.location, prep.viewLocation);
  if (req.status === KE.REQUEST_STATUS.CANCELLED) {
    throw new Error("Cannot pay a cancelled request.");
  }
  var vendor = normalize_(payload.vendor || payload.vendorName) ||
    readRequestField_(req, ["Vendor", "Vendor Name"]);
  if (!vendor) {
    throw new Error("Vendor is required for vendor payment.");
  }
  var payId = generateID_("PAY");
  appendRow_(KE.SHEETS.PAYMENT, [
    payId, requestId, vendor,
    amount, mode, KE.PAYMENT_STATUS.RECORDED, todayStr_()
  ]);
  updateRequestFields_(req.sheetRow, req.headers, {
    "Payment Status": "Paid",
    "Status": KE.REQUEST_STATUS.PAID
  });
  updateRequestFields_(req.sheetRow, req.headers, {
    "Status": KE.REQUEST_STATUS.APPROVED
  });
  return ok_("Vendor payment recorded and request approved.", { paymentId: payId });
}

// NO APPROVAL → NO ORDER
function placeOrder(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var req = findRequestRow_(normalize_(payload.requestId));
  assertRecordLocationAccess_(prep.token, req.location, prep.viewLocation);
  assertCan_("order", req.location, prep.token);
  if (req.status !== KE.REQUEST_STATUS.APPROVED &&
      req.status !== KE.REQUEST_STATUS.PAID) {
    throw new Error("NO APPROVAL → NO ORDER. Request must be Approved (vendor payment recorded).");
  }
  var vendor = normalize_(payload.vendor || payload.vendorName) ||
    readRequestField_(req, ["Vendor", "Vendor Name"]);
  var item = normalize_(payload.item) || readRequestField_(req, "Item");
  var qty = Number(payload.qty) || Number(req.row[findCol_(req.headers, "Qty")]);
  var loc = normalize_(payload.location) || req.location;
  if (!vendor || !item || !qty) {
    throw new Error("Vendor, item, and quantity are required.");
  }
  assertMasterExists_(KE.SHEETS.VENDOR, "Vendor Name", vendor);
  var orderId = generateID_("ORD");
  appendRow_(KE.SHEETS.ORDER, [
    orderId, normalize_(payload.requestId), vendor, item, qty, loc,
    todayStr_(), KE.ORDER_STATUS.PLACED
  ]);
  updateRequestFields_(req.sheetRow, req.headers, {
    "Status": KE.REQUEST_STATUS.ORDERED
  });
  return ok_("Vendor order placed.", { orderId: orderId });
}

// Receive goods → inventory +
function receiveGoods(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var orderId = normalize_(payload.orderId);
  var qty = Number(payload.qty);
  var loc = normalize_(payload.location);
  var storage = normalize_(payload.storageLocation);
  var receivedBy = normalize_(payload.receivedBy) || getActiveEmail_();
  if (!orderId || !qty || qty <= 0) {
    throw new Error("Order ID and quantity are required.");
  }
  var order = findOrderRow_(orderId);
  assertRecordLocationAccess_(prep.token, order.location || loc, prep.viewLocation);
  assertCan_("receive", order.location || loc, prep.token);
  var item = order.item;
  validateLocation_(loc || order.location);
  var grnId = generateID_("GRN");
  appendRow_(KE.SHEETS.RECEIVED, [
    grnId, orderId, item, qty, loc || order.location,
    receivedBy, storage, todayStr_()
  ]);
  adjustInventoryQty_(item, loc || order.location, qty);
  var reqId = order.requestId;
  if (reqId) {
    try {
      var req = findRequestRow_(reqId);
      updateRequestFields_(req.sheetRow, req.headers, {
        "Status": KE.REQUEST_STATUS.RECEIVED
      });
    } catch (e) { /* ignore */ }
  }
  updateOrderStatus_(order.sheetRow, KE.ORDER_STATUS.RECEIVED);
  return ok_("Goods received and inventory updated.", { grnId: grnId });
}

function findOrderRow_(orderId) {
  var data = getSheetData_(KE.SHEETS.ORDER);
  var cId = findCol_(data.headers, "Order ID");
  var cReq = findCol_(data.headers, "Request ID");
  var cItem = findCol_(data.headers, "Item");
  var cLoc = findCol_(data.headers, "Location");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cId]) === orderId) {
      return {
        sheetRow: i + 2,
        requestId: normalize_(data.rows[i][cReq]),
        item: normalize_(data.rows[i][cItem]),
        location: normalize_(data.rows[i][cLoc]),
        row: data.rows[i],
        headers: data.headers
      };
    }
  }
  throw new Error("Order not found: " + orderId);
}

function updateOrderStatus_(sheetRow, status) {
  var sh = getSheet_(KE.SHEETS.ORDER);
  var headers = getSheetData_(KE.SHEETS.ORDER).headers;
  var col = findCol_(headers, "Status") + 1;
  if (col > 0) {
    sh.getRange(sheetRow, col).setValue(status);
  }
}

function adjustInventoryQty_(item, location, delta) {
  var sh = getSheet_(KE.SHEETS.INVENTORY);
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var cCode = findCol_(data.headers, "Item Code");
  var cName = findCol_(data.headers, "Item Name");
  var cLoc = findCol_(data.headers, "Location");
  var cQty = findCol_(data.headers, "Current Qty");
  var cUpd = findCol_(data.headers, "Last Updated");
  var cMov = findCol_(data.headers, "Last Movement");
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var match = normalize_(row[cCode]) === item || normalize_(row[cName]) === item;
    if (match && normalize_(row[cLoc]) === location) {
      var newQty = Number(row[cQty]) + delta;
      if (newQty < 0) {
        throw new Error("Insufficient stock for " + item + " at " + location);
      }
      sh.getRange(i + 2, cQty + 1).setValue(newQty);
      if (cUpd >= 0) {
        sh.getRange(i + 2, cUpd + 1).setValue(nowStr_());
      }
      if (cMov >= 0) {
        sh.getRange(i + 2, cMov + 1).setValue(nowStr_());
      }
      return newQty;
    }
  }
  throw new Error("Inventory row not found for " + item + " @ " + location);
}

// NO RECEIPT → NO ISSUE
function issueItem(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var requestId = normalize_(payload.requestId);
  var qty = Number(payload.qty);
  var issuedTo = normalize_(payload.issuedTo);
  var loc = normalize_(payload.location);
  var issuedBy = normalize_(payload.issuedBy) || getActiveEmail_();
  if (!requestId || !qty || qty <= 0) {
    throw new Error("Request ID and quantity are required.");
  }
  var req = findRequestRow_(requestId);
  assertRecordLocationAccess_(prep.token, req.location || loc, prep.viewLocation);
  assertCan_("issue", req.location || loc, prep.token);
  if (!hasReceiptForRequest_(requestId)) {
    throw new Error("NO RECEIPT → NO ISSUE. Receive goods against the order first.");
  }
  var itemCol = findCol_(req.headers, "Item");
  var locCol = findCol_(req.headers, "Location");
  var item = normalize_(payload.item) || normalize_(req.row[itemCol]);
  loc = loc || normalize_(req.row[locCol]);
  var stock = getStockQty_(item, loc);
  if (stock < qty) {
    throw new Error("Cannot issue: stock (" + stock + ") is less than requested qty (" + qty + ").");
  }
  adjustInventoryQty_(item, loc, -qty);
  var issueId = generateID_("ISS");
  appendRow_(KE.SHEETS.ISSUE, [
    issueId, requestId, item, qty, issuedTo,
    loc, issuedBy, todayStr_()
  ]);
  updateRequestFields_(req.sheetRow, req.headers, {
    "Status": KE.REQUEST_STATUS.COMPLETED
  });
  return ok_("Item issued and inventory reduced.", { issueId: issueId });
}

function hasReceiptForRequest_(requestId) {
  var orders = getSheetData_(KE.SHEETS.ORDER);
  var cReq = findCol_(orders.headers, "Request ID");
  var cOrd = findCol_(orders.headers, "Order ID");
  var orderIds = [];
  for (var i = 0; i < orders.rows.length; i++) {
    if (normalize_(orders.rows[i][cReq]) === requestId) {
      orderIds.push(normalize_(orders.rows[i][cOrd]));
    }
  }
  if (!orderIds.length) {
    return false;
  }
  var grn = getSheetData_(KE.SHEETS.RECEIVED);
  var cGrnOrd = findCol_(grn.headers, "Order ID");
  for (var j = 0; j < grn.rows.length; j++) {
    if (orderIds.indexOf(normalize_(grn.rows[j][cGrnOrd])) >= 0) {
      return true;
    }
  }
  return false;
}

function getStockQty_(item, location) {
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var cCode = findCol_(data.headers, "Item Code");
  var cName = findCol_(data.headers, "Item Name");
  var cLoc = findCol_(data.headers, "Location");
  var cQty = findCol_(data.headers, "Current Qty");
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var match = normalize_(row[cCode]) === item || normalize_(row[cName]) === item;
    if (match && normalize_(row[cLoc]) === location) {
      return Number(row[cQty]) || 0;
    }
  }
  return 0;
}
