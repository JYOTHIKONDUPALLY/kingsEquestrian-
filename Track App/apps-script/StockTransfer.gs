/**
 * Stock transfer workflow between locations.
 *
 * Roles allowed: Admin (any), Manager (only when their location is From or To).
 * When a transfer is "Completed" the source inventory is reduced and the
 * destination row gets the qty (created if it doesn't exist yet).
 */

function createStockTransfer(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = requireUserFromToken_(prep.token);
  if (!canTransferStock_(user)) {
    throw new Error("Your role (" + user.role + ") cannot move stock between locations.");
  }

  var itemRef = normalize_(payload.item);
  var fromLoc = normalize_(payload.fromLocation);
  var toLoc = normalize_(payload.toLocation);
  var qty = Number(payload.qty);
  var notes = normalize_(payload.notes);
  var autoComplete = payload.autoComplete !== false; // default true – move immediately

  if (!itemRef || !fromLoc || !toLoc || !qty || qty <= 0) {
    throw new Error("Item, from-location, to-location and quantity are required.");
  }
  if (fromLoc === toLoc) {
    throw new Error("From and To locations must be different.");
  }
  validateLocation_(fromLoc);
  validateLocation_(toLoc);

  // Managers can only initiate transfers that involve their own location.
  if (!isAdminRole_(user.role)) {
    var userLoc = resolveLocationName_(user.location);
    if (userLoc !== fromLoc && userLoc !== toLoc) {
      throw new Error("You can only transfer to/from your assigned location (" + userLoc + ").");
    }
  }

  var inv = resolveInventoryRow_(itemRef, fromLoc);
  if (!inv) {
    throw new Error("Inventory not found for " + itemRef + " at " + fromLoc + ".");
  }
  var stock = Number(inv.row[findCol_(inv.headers, "Current Qty")]) || 0;
  if (stock < qty) {
    throw new Error("Source stock (" + stock + ") is less than transfer qty (" + qty + ").");
  }

  var status = autoComplete ? KE.TRANSFER_STATUS.COMPLETED : KE.TRANSFER_STATUS.REQUESTED;
  var transferId = generateID_("TRF");
  appendRow_(KE.SHEETS.TRANSFER, [
    transferId,
    inv.itemCode,
    inv.itemName,
    fromLoc,
    toLoc,
    qty,
    user.name || user.email,
    autoComplete ? (user.name || user.email) : "",
    todayStr_(),
    status,
    notes
  ]);

  if (autoComplete) {
    moveInventoryBetweenLocations_(inv, toLoc, qty);
  }
  return ok_("Transfer " + transferId + " " + status.toLowerCase() + ".", {
    transferId: transferId,
    status: status
  });
}

function approveStockTransfer(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = requireUserFromToken_(prep.token);
  if (!canTransferStock_(user)) {
    throw new Error("Your role cannot approve transfers.");
  }
  var transferId = normalize_(payload.transferId);
  if (!transferId) {
    throw new Error("Transfer ID is required.");
  }

  var sh = getSheet_(KE.SHEETS.TRANSFER);
  var data = getSheetData_(KE.SHEETS.TRANSFER);
  var headers = data.headers;
  var cId = findCol_(headers, "Transfer ID");
  var cItemCode = findCol_(headers, "Item Code");
  var cFrom = findCol_(headers, "From Location");
  var cTo = findCol_(headers, "To Location");
  var cQty = findCol_(headers, "Qty");
  var cStatus = findCol_(headers, "Status");
  var cApprover = findCol_(headers, "Approved By");

  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cId]) !== transferId) {
      continue;
    }
    var row = data.rows[i];
    var current = normalize_(row[cStatus]);
    if (current === KE.TRANSFER_STATUS.COMPLETED) {
      throw new Error("Transfer already completed.");
    }
    if (current === KE.TRANSFER_STATUS.CANCELLED) {
      throw new Error("Transfer was cancelled.");
    }

    var item = normalize_(row[cItemCode]);
    var fromLoc = normalize_(row[cFrom]);
    var toLoc = normalize_(row[cTo]);
    var qty = Number(row[cQty]) || 0;
    var inv = resolveInventoryRow_(item, fromLoc);
    if (!inv) {
      throw new Error("Source inventory missing for " + item + " @ " + fromLoc);
    }
    moveInventoryBetweenLocations_(inv, toLoc, qty);
    var sheetRow = i + 2;
    sh.getRange(sheetRow, cStatus + 1).setValue(KE.TRANSFER_STATUS.COMPLETED);
    if (cApprover >= 0) {
      sh.getRange(sheetRow, cApprover + 1).setValue(user.name || user.email);
    }
    return ok_("Transfer " + transferId + " completed.", { transferId: transferId });
  }
  throw new Error("Transfer not found: " + transferId);
}

function cancelStockTransfer(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  assertAdmin_(prep.token);
  var transferId = normalize_(payload.transferId);
  var sh = getSheet_(KE.SHEETS.TRANSFER);
  var data = getSheetData_(KE.SHEETS.TRANSFER);
  var headers = data.headers;
  var cId = findCol_(headers, "Transfer ID");
  var cStatus = findCol_(headers, "Status");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cId]) === transferId) {
      sh.getRange(i + 2, cStatus + 1).setValue(KE.TRANSFER_STATUS.CANCELLED);
      return ok_("Transfer cancelled.");
    }
  }
  throw new Error("Transfer not found: " + transferId);
}

/**
 * Locates an INVENTORY_MASTER row by Item Code (preferred) or Item Name for a location.
 */
function resolveInventoryRow_(itemRef, location) {
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var headers = data.headers;
  var cCode = findCol_(headers, "Item Code");
  var cName = findCol_(headers, "Item Name");
  var cLoc = findCol_(headers, "Location");
  var loc = resolveLocationName_(location);
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var match = normalize_(row[cCode]) === itemRef ||
                normalize_(row[cName]) === itemRef;
    if (match && resolveLocationName_(row[cLoc]) === loc) {
      return {
        sheetRow: i + 2,
        row: row,
        headers: headers,
        itemCode: normalize_(row[cCode]),
        itemName: normalize_(row[cName])
      };
    }
  }
  return null;
}

/**
 * Moves qty from src inventory row into target location. Creates destination row
 * if it doesn't exist (copies catalog metadata from source).
 */
function moveInventoryBetweenLocations_(srcInv, toLoc, qty) {
  var sh = getSheet_(KE.SHEETS.INVENTORY);
  var headers = srcInv.headers;
  var cQty = findCol_(headers, "Current Qty");
  var cLastMov = findCol_(headers, "Last Movement");
  var srcQty = Number(srcInv.row[cQty]) || 0;
  if (srcQty < qty) {
    throw new Error("Source stock has changed; insufficient qty.");
  }
  sh.getRange(srcInv.sheetRow, cQty + 1).setValue(srcQty - qty);
  if (cLastMov >= 0) {
    sh.getRange(srcInv.sheetRow, cLastMov + 1).setValue(nowStr_());
  }

  var dest = resolveInventoryRow_(srcInv.itemCode, toLoc);
  if (dest) {
    var destQty = Number(dest.row[cQty]) || 0;
    sh.getRange(dest.sheetRow, cQty + 1).setValue(destQty + qty);
    if (cLastMov >= 0) {
      sh.getRange(dest.sheetRow, cLastMov + 1).setValue(nowStr_());
    }
    return;
  }

  // Create destination row by cloning source values, replacing Location and Qty.
  var newRow = [];
  for (var i = 0; i < headers.length; i++) {
    newRow.push(srcInv.row[i]);
  }
  var cLoc = findCol_(headers, "Location");
  var cMin = findCol_(headers, "Min Level");
  var cStore = findCol_(headers, "Storage Location");
  var cResv = findCol_(headers, "Reserved Qty");
  var cAlloc = findCol_(headers, "Allocated To");
  if (cLoc >= 0) {
    newRow[cLoc] = toLoc;
  }
  if (cQty >= 0) {
    newRow[cQty] = qty;
  }
  if (cMin >= 0) {
    newRow[cMin] = newRow[cMin] || 0;
  }
  if (cStore >= 0) {
    newRow[cStore] = "";
  }
  if (cResv >= 0) {
    newRow[cResv] = 0;
  }
  if (cAlloc >= 0) {
    newRow[cAlloc] = "";
  }
  if (cLastMov >= 0) {
    newRow[cLastMov] = nowStr_();
  }
  sh.appendRow(newRow);
}
