/**
 * Sample / demo movement workflow.
 *
 *  - createSampleMovement(payload)  : issue a sample to a customer / trainer.
 *                                     Reduces source inventory by qty.
 *  - returnSampleMovement(payload)  : mark a sample as returned and add qty
 *                                     back into the same location.
 *  - markSampleLost(payload)        : mark as lost (no inventory adjustment).
 *  - getSampleMovements(payload)    : list all sample rows (location filtered).
 *
 * Roles allowed for issue / return / lost: Admin + Manager (or anyone with
 * canManageSamples_). Listing is open to all logged-in users in their view.
 */

function createSampleMovement(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = requireUserFromToken_(prep.token);
  if (!canManageSamples_(user)) {
    throw new Error("Your role (" + user.role + ") cannot issue samples.");
  }

  var itemRef = normalize_(payload.item);
  var loc     = normalize_(payload.location);
  var issuedTo = normalize_(payload.issuedTo);
  var qty     = Number(payload.qty);
  var purpose = normalize_(payload.purpose);
  var expectedReturn = normalize_(payload.expectedReturn);
  var notes   = normalize_(payload.notes);

  if (!itemRef || !loc || !issuedTo || !qty || qty <= 0) {
    throw new Error("Item, location, issued-to and quantity are required.");
  }
  validateLocation_(loc);
  if (!isAdminRole_(user.role)) {
    var userLoc = resolveLocationName_(user.location);
    if (userLoc !== loc) {
      throw new Error("You can only issue samples from your assigned location (" + userLoc + ").");
    }
  }

  var inv = resolveInventoryRow_(itemRef, loc);
  if (!inv) {
    throw new Error("Inventory not found for " + itemRef + " at " + loc + ".");
  }
  var cQty = findCol_(inv.headers, "Current Qty");
  var stock = Number(inv.row[cQty]) || 0;
  if (stock < qty) {
    throw new Error("Cannot issue sample: stock (" + stock + ") is less than qty (" + qty + ").");
  }

  // Reduce inventory
  adjustInventoryQty_(inv.itemCode, loc, -qty);

  var sampleId = generateID_("SMP");
  appendRow_(KE.SHEETS.SAMPLE, [
    sampleId,
    inv.itemCode,
    inv.itemName,
    loc,
    issuedTo,
    purpose,
    todayStr_(),         // Date Out
    expectedReturn,      // Expected Return
    "",                  // Date Returned (blank until return)
    KE.SAMPLE_STATUS.OUT,
    notes
  ]);

  return ok_("Sample issued.", { sampleId: sampleId, qty: qty });
}

function returnSampleMovement(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = requireUserFromToken_(prep.token);
  if (!canManageSamples_(user)) {
    throw new Error("Your role (" + user.role + ") cannot manage samples.");
  }

  var sampleId = normalize_(payload.sampleId);
  if (!sampleId) {
    throw new Error("Sample ID is required.");
  }
  var qtyReturning = Number(payload.qty);

  var sample = findSampleRow_(sampleId);
  if (sample.status === KE.SAMPLE_STATUS.RETURNED) {
    throw new Error("Sample already returned.");
  }
  if (sample.status === KE.SAMPLE_STATUS.LOST) {
    throw new Error("Sample is marked as lost.");
  }
  // If qty omitted, default to original issue qty (we don't store that —
  // so we trust the caller, or default to 1).
  if (!qtyReturning || qtyReturning <= 0) {
    qtyReturning = 1;
  }

  // Add qty back to inventory at the same location
  try {
    adjustInventoryQty_(sample.itemCode || sample.itemName, sample.location, qtyReturning);
  } catch (e) {
    throw new Error("Could not update inventory: " + e.message);
  }

  updateSampleFields_(sample.sheetRow, sample.headers, {
    "Status": KE.SAMPLE_STATUS.RETURNED,
    "Date Returned": todayStr_()
  });

  return ok_("Sample returned and inventory restored.", { sampleId: sampleId, qty: qtyReturning });
}

function markSampleLost(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  var user = requireUserFromToken_(prep.token);
  if (!canManageSamples_(user)) {
    throw new Error("Your role (" + user.role + ") cannot manage samples.");
  }
  var sampleId = normalize_(payload.sampleId);
  if (!sampleId) {
    throw new Error("Sample ID is required.");
  }
  var sample = findSampleRow_(sampleId);
  if (sample.status === KE.SAMPLE_STATUS.RETURNED) {
    throw new Error("Sample was already returned.");
  }
  updateSampleFields_(sample.sheetRow, sample.headers, {
    "Status": KE.SAMPLE_STATUS.LOST,
    "Notes": (sample.notes ? sample.notes + " | " : "") + "Marked lost on " + todayStr_()
  });
  return ok_("Sample marked as lost.", { sampleId: sampleId });
}

function getSampleMovements(payload) {
  var prep = prepPayload_(payload);
  validateSessionToken_(prep.token);
  var viewLocation = normalize_((prep.payload && prep.payload.viewLocation) || "");
  var data = getSheetData_(KE.SHEETS.SAMPLE);
  var rows = data.rows.map(function (r) {
    var obj = {};
    data.headers.forEach(function (h, i) { obj[h] = serializeCellValue_(r[i]); });
    return obj;
  });
  if (viewLocation && !isAllLocations_(viewLocation)) {
    rows = rows.filter(function (r) { return r["Location"] === viewLocation; });
  }
  // newest first
  rows.sort(function (a, b) {
    var da = a["Date Out"] ? new Date(a["Date Out"]).getTime() : 0;
    var db = b["Date Out"] ? new Date(b["Date Out"]).getTime() : 0;
    return db - da;
  });
  return { success: true, samples: rows };
}

function findSampleRow_(sampleId) {
  var data = getSheetData_(KE.SHEETS.SAMPLE);
  var cId   = findCol_(data.headers, "Sample ID");
  var cCode = findCol_(data.headers, "Item Code");
  var cName = findCol_(data.headers, "Item Name");
  var cLoc  = findCol_(data.headers, "Location");
  var cStat = findCol_(data.headers, "Status");
  var cNote = findCol_(data.headers, "Notes");
  for (var i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i][cId]) === sampleId) {
      return {
        sheetRow: i + 2,
        headers: data.headers,
        itemCode: normalize_(data.rows[i][cCode]),
        itemName: normalize_(data.rows[i][cName]),
        location: normalize_(data.rows[i][cLoc]),
        status:   normalize_(data.rows[i][cStat]),
        notes:    cNote >= 0 ? normalize_(data.rows[i][cNote]) : ""
      };
    }
  }
  throw new Error("Sample not found: " + sampleId);
}

function updateSampleFields_(sheetRow, headers, updates) {
  var sh = getSheet_(KE.SHEETS.SAMPLE);
  Object.keys(updates).forEach(function (field) {
    var col = findCol_(headers, field);
    if (col < 0) {
      return;
    }
    sh.getRange(sheetRow, col + 1).setValue(updates[field]);
  });
}
