/**
 * Admin-only pricing master and financial KPIs.
 *
 * All exposed functions enforce assertAdmin_() so the row data (cost, margin,
 * profitability) never reaches non-admin clients. Operations users can still
 * see Selling Price / MRP indirectly via the Inventory Control Center, but
 * never the cost/margins.
 */

function getPricingList(token) {
  assertAdmin_(token);
  var data = getSheetData_(KE.SHEETS.PRICING);
  var headers = data.headers;
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    out.push(rowToObject_(headers, data.rows[i]));
  }
  return sanitizeForClient_(out);
}

function upsertPricing(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  assertAdmin_(prep.token);

  var itemCode = normalize_(payload.itemCode);
  var itemName = normalize_(payload.itemName);
  var cost = Number(payload.costPrice) || 0;
  var sell = Number(payload.sellingPrice) || 0;
  var mrp = Number(payload.mrp) || 0;
  var vendor = normalize_(payload.preferredVendor);
  var lastPurchase = Number(payload.lastPurchaseCost) || cost;

  if (!itemCode && !itemName) {
    throw new Error("Item Code or Item Name is required.");
  }
  if (cost <= 0 && sell <= 0 && mrp <= 0) {
    throw new Error("Enter at least one of cost, selling, or MRP.");
  }
  var margin = sell > 0 && cost > 0 ? Math.round(((sell - cost) / sell) * 10000) / 100 : 0;

  var sh = getSheet_(KE.SHEETS.PRICING);
  var data = getSheetData_(KE.SHEETS.PRICING);
  var headers = data.headers;
  var cCode = findCol_(headers, "Item Code");
  var cName = findCol_(headers, "Item Name");

  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    if ((itemCode && normalize_(row[cCode]) === itemCode) ||
        (!itemCode && itemName && normalize_(row[cName]) === itemName)) {
      var sheetRow = i + 2;
      updatePricingRow_(sh, sheetRow, headers, {
        "Item Code": itemCode || normalize_(row[cCode]),
        "Item Name": itemName || normalize_(row[cName]),
        "Cost Price": cost,
        "Selling Price": sell,
        "MRP": mrp,
        "Profit Margin %": margin,
        "Preferred Vendor": vendor || normalize_(row[findCol_(headers, "Preferred Vendor")]),
        "Last Purchase Cost": lastPurchase,
        "Last Updated": nowStr_()
      });
      mirrorCostToInventory_(itemCode || normalize_(row[cCode]), itemName || normalize_(row[cName]), cost);
      return ok_("Pricing updated.", { itemCode: itemCode || normalize_(row[cCode]) });
    }
  }

  appendRow_(KE.SHEETS.PRICING, [
    itemCode, itemName, cost, sell, mrp, margin, vendor, lastPurchase, nowStr_()
  ]);
  mirrorCostToInventory_(itemCode, itemName, cost);
  return ok_("Pricing added.", { itemCode: itemCode });
}

function updatePricingRow_(sh, sheetRow, headers, updates) {
  Object.keys(updates).forEach(function (field) {
    var col = findCol_(headers, field);
    if (col < 0) {
      return;
    }
    sh.getRange(sheetRow, col + 1).setValue(updates[field]);
  });
}

/**
 * Copies cost price into INVENTORY_MASTER for matching rows (so admin analytics
 * can compute inventory value from inventory directly).
 */
function mirrorCostToInventory_(itemCode, itemName, cost) {
  if (!cost || cost <= 0) {
    return;
  }
  var sh = getSheet_(KE.SHEETS.INVENTORY);
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  var headers = data.headers;
  var cCode = findCol_(headers, "Item Code");
  var cName = findCol_(headers, "Item Name");
  var cCost = findCol_(headers, "Cost Price");
  if (cCost < 0) {
    return;
  }
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var match = (itemCode && normalize_(row[cCode]) === itemCode) ||
                (!itemCode && itemName && normalize_(row[cName]) === itemName);
    if (match) {
      sh.getRange(i + 2, cCost + 1).setValue(cost);
    }
  }
}

/**
 * Returns financial KPIs for the dashboard (admin only).
 */
function getFinancialKpis(token, viewLocation) {
  var user = assertAdmin_(token);
  var view = resolveViewLocation_(user, viewLocation);

  // Inventory value (uses INVENTORY_MASTER cost price)
  var inv = getSheetData_(KE.SHEETS.INVENTORY);
  inv.rows = applySheetLocationFilter_(KE.SHEETS.INVENTORY, inv.rows, inv.headers, view);
  var iQty = findCol_(inv.headers, "Current Qty");
  var iCost = findCol_(inv.headers, "Cost Price");
  var iType = findCol_(inv.headers, "Inventory Type");
  var totalValue = 0;
  var businessValue = 0;
  var sampleValue = 0;
  var operationalValue = 0;
  for (var i = 0; i < inv.rows.length; i++) {
    var qty = Number(inv.rows[i][iQty]) || 0;
    var cost = Number(inv.rows[i][iCost]) || 0;
    var type = iType >= 0 ? normalize_(inv.rows[i][iType]) : "";
    var v = qty * cost;
    totalValue += v;
    if (type === "Business Inventory") businessValue += v;
    else if (type === "Sample / Demo Inventory") sampleValue += v;
    else if (type === "Operational Inventory") operationalValue += v;
  }

  // Revenue (PAYMENTS sheet, sum of Amount)
  var pay = getSheetData_(KE.SHEETS.PAYMENT);
  pay.rows = applySheetLocationFilter_(KE.SHEETS.PAYMENT, pay.rows, pay.headers, view);
  var cAmt = findCol_(pay.headers, "Amount");
  var revenue = 0;
  for (var p = 0; p < pay.rows.length; p++) {
    revenue += Number(pay.rows[p][cAmt]) || 0;
  }

  // Vendor payables (placed orders - using selling price proxy if no cost)
  var ord = getSheetData_(KE.SHEETS.ORDER);
  ord.rows = applySheetLocationFilter_(KE.SHEETS.ORDER, ord.rows, ord.headers, view);
  var oQty = findCol_(ord.headers, "Qty");
  var oStatus = findCol_(ord.headers, "Status");
  var vendorOrders = 0;
  var vendorPending = 0;
  for (var o = 0; o < ord.rows.length; o++) {
    var oq = Number(ord.rows[o][oQty]) || 0;
    vendorOrders += oq;
    if (normalize_(ord.rows[o][oStatus]) === KE.ORDER_STATUS.PLACED) {
      vendorPending += oq;
    }
  }

  // Approx gross profit using pricing master
  var pricing = getSheetData_(KE.SHEETS.PRICING);
  var pcCode = findCol_(pricing.headers, "Item Code");
  var pcSell = findCol_(pricing.headers, "Selling Price");
  var pcCost = findCol_(pricing.headers, "Cost Price");
  var priceMap = {};
  for (var pp = 0; pp < pricing.rows.length; pp++) {
    var code = normalize_(pricing.rows[pp][pcCode]);
    if (!code) continue;
    priceMap[code] = {
      sell: Number(pricing.rows[pp][pcSell]) || 0,
      cost: Number(pricing.rows[pp][pcCost]) || 0
    };
  }

  return sanitizeForClient_({
    revenue: revenue,
    inventoryValue: totalValue,
    businessInventoryValue: businessValue,
    sampleInventoryValue: sampleValue,
    operationalInventoryValue: operationalValue,
    vendorOrderUnits: vendorOrders,
    vendorPendingUnits: vendorPending,
    pricingRows: pricing.rows.length,
    pricingCoverage: Object.keys(priceMap).length
  });
}
