/**
 * Inventory analytics for the dashboard:
 *  - Inventory Control Center (multi-location pivot)
 *  - Low stock + health status
 *  - Dead stock + fast moving
 *  - Location analytics
 *  - True available stock
 *  - Inventory ownership analytics (by Inventory Type)
 *
 * Cost / value fields are returned only when the caller is allowed to see them
 * (typically Admin). Use sanitizeFinancialsForUser_() to strip values for non-admins.
 */

function getInventoryControlCenter_(viewLocation, includeFinancials) {
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var headers = data.headers;
  var rows = applySheetLocationFilter_(KE.SHEETS.INVENTORY, data.rows, headers, viewLocation);

  var cCode = findCol_(headers, "Item Code");
  var cName = findCol_(headers, "Item Name");
  var cCat = findCol_(headers, "Category");
  var cVen = findCol_(headers, "Vendor");
  var cLoc = findCol_(headers, "Location");
  var cQty = findCol_(headers, "Current Qty");
  var cMin = findCol_(headers, "Min Level");
  var cType = findCol_(headers, "Inventory Type");
  var cAlloc = findCol_(headers, "Allocated To");
  var cResv = findCol_(headers, "Reserved Qty");
  var cCost = findCol_(headers, "Cost Price");
  var cImg = findCol_(headers, "Image URL");

  var pivot = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var code = normalize_(row[cCode]);
    if (!code) {
      continue;
    }
    if (!pivot[code]) {
      pivot[code] = {
        itemCode: code,
        itemName: cName >= 0 ? normalize_(row[cName]) : "",
        category: cCat >= 0 ? normalize_(row[cCat]) : "",
        vendor: cVen >= 0 ? normalize_(row[cVen]) : "",
        minLevel: 0,
        locations: {},
        totalQty: 0,
        reservedQty: 0,
        allocatedTo: "",
        inventoryType: "",
        costPrice: 0,
        inventoryValue: 0,
        availableForSale: 0
        ,imageUrl: ""
      };
    }
    var entry = pivot[code];
    var loc = cLoc >= 0 ? normalize_(row[cLoc]) : "";
    var qty = Number(cQty >= 0 ? row[cQty] : 0) || 0;
    var resv = Number(cResv >= 0 ? row[cResv] : 0) || 0;
    var min = Number(cMin >= 0 ? row[cMin] : 0) || 0;
    var cost = Number(cCost >= 0 ? row[cCost] : 0) || 0;
    var type = cType >= 0 ? normalize_(row[cType]) : "";
    var alloc = cAlloc >= 0 ? normalize_(row[cAlloc]) : "";
    var imageUrl = cImg >= 0 ? normalize_(row[cImg]) : "";

    if (loc) {
      entry.locations[loc] = (entry.locations[loc] || 0) + qty;
    }
    entry.totalQty += qty;
    entry.reservedQty += resv;
    if (min > entry.minLevel) {
      entry.minLevel = min;
    }
    if (type && !entry.inventoryType) {
      entry.inventoryType = type;
    }
    if (alloc && !entry.allocatedTo) {
      entry.allocatedTo = alloc;
    }
    if (imageUrl && !entry.imageUrl) {
      entry.imageUrl = imageUrl;
    }
    if (cost > 0 && !entry.costPrice) {
      entry.costPrice = cost;
    }
    if (cost > 0) {
      entry.inventoryValue += qty * cost;
    }
    if (type === "Business Inventory" || !type) {
      entry.availableForSale += Math.max(0, qty - resv - (alloc ? 0 : 0));
    }
  }

  var items = Object.keys(pivot).map(function (k) {
    var e = pivot[k];
    e.locationsList = KE.LOCATIONS.map(function (loc) {
      return { location: loc, qty: e.locations[loc] || 0 };
    });
    e.status = stockHealthStatus_(e.totalQty, e.minLevel);
    if (!includeFinancials) {
      e.costPrice = null;
      e.inventoryValue = null;
    }
    return e;
  });

  items.sort(function (a, b) {
    return a.itemName.localeCompare(b.itemName);
  });

  return items;
}

function stockHealthStatus_(qty, min) {
  qty = Number(qty) || 0;
  min = Number(min) || 0;
  if (min <= 0 && qty > 0) {
    return "Healthy";
  }
  if (qty <= 0) {
    return "Out of stock";
  }
  if (qty <= min) {
    return "Low stock";
  }
  if (qty <= min * 1.5) {
    return "Near min";
  }
  return "Healthy";
}

/**
 * Location-wise rollup of units and (admin-only) value.
 */
function getLocationInventoryAnalytics_(includeFinancials) {
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var headers = data.headers;
  var cLoc = findCol_(headers, "Location");
  var cQty = findCol_(headers, "Current Qty");
  var cCost = findCol_(headers, "Cost Price");
  var cType = findCol_(headers, "Inventory Type");

  var map = {};
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var loc = cLoc >= 0 ? normalize_(row[cLoc]) : "";
    if (!loc) {
      continue;
    }
    var qty = Number(cQty >= 0 ? row[cQty] : 0) || 0;
    var cost = Number(cCost >= 0 ? row[cCost] : 0) || 0;
    var type = cType >= 0 ? normalize_(row[cType]) : "";
    if (!map[loc]) {
      map[loc] = { location: loc, units: 0, value: 0, types: {} };
    }
    map[loc].units += qty;
    if (cost > 0) {
      map[loc].value += qty * cost;
    }
    if (type) {
      map[loc].types[type] = (map[loc].types[type] || 0) + qty;
    }
  }

  var out = KE.LOCATIONS.map(function (loc) {
    var m = map[loc] || { location: loc, units: 0, value: 0, types: {} };
    if (!includeFinancials) {
      m.value = null;
    }
    return m;
  });

  return out;
}

/**
 * Inventory rollup by ownership type (Business / Customer / Sample / Operational).
 */
function getInventoryByType_(viewLocation, includeFinancials) {
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  data.rows = applySheetLocationFilter_(KE.SHEETS.INVENTORY, data.rows, data.headers, viewLocation);

  var cQty = findCol_(data.headers, "Current Qty");
  var cType = findCol_(data.headers, "Inventory Type");
  var cCost = findCol_(data.headers, "Cost Price");

  var map = {};
  KE.INVENTORY_TYPES.forEach(function (t) {
    map[t] = { type: t, qty: 0, value: 0 };
  });

  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var type = (cType >= 0 ? normalize_(row[cType]) : "") || KE.DEFAULT_INVENTORY_TYPE;
    var qty = Number(cQty >= 0 ? row[cQty] : 0) || 0;
    var cost = Number(cCost >= 0 ? row[cCost] : 0) || 0;
    if (!map[type]) {
      map[type] = { type: type, qty: 0, value: 0 };
    }
    map[type].qty += qty;
    if (cost > 0) {
      map[type].value += qty * cost;
    }
  }

  return KE.INVENTORY_TYPES.map(function (t) {
    var m = map[t];
    if (!includeFinancials) {
      m.value = null;
    }
    return m;
  });
}

/**
 * Compute fast moving and dead stock from ISSUE_REGISTER.
 *  fast moving = highest issue counts (top 5)
 *  dead stock = items with no issue in last 90 days
 */
function getMovementAnalytics_(viewLocation) {
  var iss = getSheetData_(KE.SHEETS.ISSUE);
  iss.rows = applySheetLocationFilter_(KE.SHEETS.ISSUE, iss.rows, iss.headers, viewLocation);
  var cItem = findCol_(iss.headers, "Item");
  var cQty = findCol_(iss.headers, "Qty");
  var cDate = findCol_(iss.headers, "Date");

  var inv = getSheetData_(KE.SHEETS.INVENTORY);
  inv.rows = applySheetLocationFilter_(KE.SHEETS.INVENTORY, inv.rows, inv.headers, viewLocation);
  var iName = findCol_(inv.headers, "Item Name");
  var iCode = findCol_(inv.headers, "Item Code");

  var counts = {};
  var lastIssued = {};
  for (var i = 0; i < iss.rows.length; i++) {
    var key = normalize_(iss.rows[i][cItem]);
    if (!key) {
      continue;
    }
    counts[key] = (counts[key] || 0) + (Number(iss.rows[i][cQty]) || 1);
    var d = iss.rows[i][cDate];
    var dt = d instanceof Date ? d : new Date(String(d || ""));
    if (!isNaN(dt.getTime())) {
      if (!lastIssued[key] || dt.getTime() > lastIssued[key]) {
        lastIssued[key] = dt.getTime();
      }
    }
  }

  var fast = Object.keys(counts).map(function (k) {
    return { item: k, count: counts[k] };
  }).sort(function (a, b) { return b.count - a.count; }).slice(0, 5);

  var deadCutoff = Date.now() - KE.DEAD_STOCK_DAYS * 24 * 60 * 60 * 1000;
  var dead = [];
  for (var j = 0; j < inv.rows.length; j++) {
    var nm = iName >= 0 ? normalize_(inv.rows[j][iName]) : "";
    var cd = iCode >= 0 ? normalize_(inv.rows[j][iCode]) : "";
    var key2 = nm || cd;
    if (!key2) {
      continue;
    }
    var last = lastIssued[nm] || lastIssued[cd] || 0;
    if (!last || last < deadCutoff) {
      dead.push({ itemCode: cd, itemName: nm, lastIssuedMs: last });
    }
  }

  return {
    fastMoving: fast,
    deadStock: dead.slice(0, 10),
    deadStockCount: dead.length
  };
}

/**
 * Convenience wrapper used by WebApp/getDashboardSummary_.
 * Returns a single payload with everything analytics needs, gated by financial access.
 */
function buildInventoryAnalyticsPayload_(viewLocation, user) {
  var includeFinancials = canViewFinancials_(user);
  var items = getInventoryControlCenter_(viewLocation, includeFinancials);
  var byLocation = getLocationInventoryAnalytics_(includeFinancials);
  var byType = getInventoryByType_(viewLocation, includeFinancials);
  var movement = getMovementAnalytics_(viewLocation);

  var totalUnits = items.reduce(function (sum, it) { return sum + (it.totalQty || 0); }, 0);
  var totalValue = includeFinancials
    ? items.reduce(function (sum, it) { return sum + (Number(it.inventoryValue) || 0); }, 0)
    : null;
  var availableForSale = items.reduce(function (sum, it) { return sum + (Number(it.availableForSale) || 0); }, 0);
  var lowStock = items.filter(function (it) { return it.status === "Low stock" || it.status === "Out of stock"; });

  return {
    items: items,
    byLocation: byLocation,
    byType: byType,
    movement: movement,
    totals: {
      totalUnits: totalUnits,
      totalValue: totalValue,
      availableForSale: availableForSale,
      lowStockCount: lowStock.length,
      deadStockCount: movement.deadStockCount,
      includeFinancials: !!includeFinancials
    }
  };
}
