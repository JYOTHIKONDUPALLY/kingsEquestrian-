/**
 * Weekly inventory & operations report.
 *
 *  - sendWeeklyInventoryReport()  : sends one email summarising the last 7 days
 *                                   to every Admin user in USER_MASTER.
 *  - installWeeklyReportTrigger() : installs a Monday-morning time-based trigger
 *                                   that calls sendWeeklyInventoryReport().
 *  - uninstallWeeklyReportTrigger(): removes the trigger.
 *
 * Admin-only data (cost, inventory value, vendor profitability) is included
 * because the recipients are always Admin users.
 */

var KE_REPORT_TRIGGER_FN_ = "sendWeeklyInventoryReport";

function installWeeklyReportTrigger() {
  uninstallWeeklyReportTrigger();
  ScriptApp.newTrigger(KE_REPORT_TRIGGER_FN_)
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  return "Weekly report trigger installed (Mondays at 08:00).";
}

function uninstallWeeklyReportTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === KE_REPORT_TRIGGER_FN_) {
      ScriptApp.deleteTrigger(tr);
      removed++;
    }
  });
  return removed + " trigger(s) removed.";
}

/**
 * Build the report payload and send the email.
 * Returns a small status object so it's also callable from the Apps Script UI.
 */
function sendWeeklyInventoryReport() {
  var recipients = getAdminEmailRecipients_();
  if (!recipients.length) {
    return { sent: 0, message: "No Admin recipients found in USER_MASTER." };
  }
  var report = buildWeeklyReportPayload_();
  var subject = "Kings Equestrian – Weekly Report (" + report.period.start + " → " + report.period.end + ")";
  var html = renderWeeklyReportHtml_(report);

  MailApp.sendEmail({
    to: recipients.join(","),
    subject: subject,
    htmlBody: html
  });

  return { sent: recipients.length, recipients: recipients };
}

/**
 * Returns the email addresses of every Admin user in USER_MASTER.
 */
function getAdminEmailRecipients_() {
  var data = getSheetData_(KE.SHEETS.USER);
  var cEmail = findCol_(data.headers, "Email");
  var cRole  = findCol_(data.headers, "Role");
  var out = [];
  if (cEmail < 0) {
    return out;
  }
  for (var i = 0; i < data.rows.length; i++) {
    var role = cRole >= 0 ? normalize_(data.rows[i][cRole]) : "";
    if (role !== KE.ROLES.ADMIN) {
      continue;
    }
    var email = normalize_(data.rows[i][cEmail]);
    if (email && email.indexOf("@") > 0 && out.indexOf(email) < 0) {
      out.push(email);
    }
  }
  return out;
}

function buildWeeklyReportPayload_() {
  var now = new Date();
  var weekStart = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  var fmt = function (d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd");
  };
  var period = { start: fmt(weekStart), end: fmt(now), startMs: weekStart.getTime() };

  // Inventory summary across all locations (admin-level)
  var inv = getInventoryControlCenter_("", true);
  var totalUnits = 0;
  var totalValue = 0;
  var lowStock = [];
  inv.forEach(function (it) {
    totalUnits += it.totalQty || 0;
    totalValue += Number(it.inventoryValue) || 0;
    if (it.status === "Low stock" || it.status === "Out of stock") {
      lowStock.push(it);
    }
  });
  lowStock.sort(function (a, b) { return (a.totalQty || 0) - (b.totalQty || 0); });

  // 7-day activity
  var weekActivity = {
    requests:  countSheetRowsSince_(KE.SHEETS.REQUEST, "Date", period.startMs),
    payments:  countSheetRowsSince_(KE.SHEETS.PAYMENT, "Date", period.startMs),
    orders:    countSheetRowsSince_(KE.SHEETS.ORDER,  "Order Date", period.startMs),
    received:  countSheetRowsSince_(KE.SHEETS.RECEIVED, "Date", period.startMs),
    issues:    countSheetRowsSince_(KE.SHEETS.ISSUE,  "Date", period.startMs),
    transfers: countSheetRowsSince_(KE.SHEETS.TRANSFER, "Date", period.startMs)
  };

  // Revenue this week (sum of payments amount where date >= weekStart)
  var pay = getSheetData_(KE.SHEETS.PAYMENT);
  var cAmt = findCol_(pay.headers, "Amount");
  var cDate = findCol_(pay.headers, "Date");
  var revenue = 0;
  for (var i = 0; i < pay.rows.length; i++) {
    var d = pay.rows[i][cDate];
    var dt = d instanceof Date ? d : new Date(String(d || ""));
    if (!isNaN(dt.getTime()) && dt.getTime() >= period.startMs) {
      revenue += Number(pay.rows[i][cAmt]) || 0;
    }
  }

  // Location-wise brief summary (all locations)
  var locationUnits = getLocationInventoryAnalytics_(true).map(function (row) {
    return {
      location: row.location,
      units: Number(row.units) || 0,
      value: Number(row.value) || 0
    };
  });

  // Movement (fast moving / dead stock)
  var movement = getMovementAnalytics_("");
  var deadByLocation = getDeadStockByLocation_();
  var locationWeekly = buildLocationWeeklySummary_(period.startMs, deadByLocation);
  var locationMovementBrief = buildLocationMovementBrief_();

  // Vendor profitability snapshot
  var vendorProfit = computeVendorProfitability_();

  return {
    period: period,
    totals: {
      totalUnits: totalUnits,
      totalValue: totalValue,
      lowStockCount: lowStock.length,
      deadStockCount: movement.deadStockCount
    },
    weekActivity: weekActivity,
    weekRevenue: revenue,
    locationUnits: locationUnits,
    deadByLocation: deadByLocation,
    locationWeekly: locationWeekly,
    locationMovementBrief: locationMovementBrief,
    lowStock: lowStock.slice(0, 10),
    fastMoving: movement.fastMoving || [],
    deadStock: (movement.deadStock || []).slice(0, 10),
    vendorProfit: vendorProfit.slice(0, 10)
  };
}

/**
 * Location-wise brief details for Low stock, Fast moving and Dead stock.
 * This makes movement sections explicitly location-based in the weekly email.
 */
function buildLocationMovementBrief_() {
  return KE.LOCATIONS.map(function (loc) {
    var low = getLowStockForLocation_(loc, 3);
    var fast = getFastMovingForLocation_(loc, 3);
    var dead = getDeadStockForLocation_(loc, 3);
    return {
      location: loc,
      lowCount: low.total,
      lowTop: low.items,
      fastTop: fast,
      deadCount: dead.total,
      deadTop: dead.items
    };
  });
}

function getLowStockForLocation_(location, limit) {
  var items = getInventoryControlCenter_(location, true).filter(function (it) {
    return it.status === "Low stock" || it.status === "Out of stock";
  });
  items.sort(function (a, b) {
    return (a.totalQty || 0) - (b.totalQty || 0);
  });
  return {
    total: items.length,
    items: items.slice(0, limit || 3).map(function (it) {
      return (it.itemName || it.itemCode || "—") + " (" + (it.totalQty || 0) + ")";
    })
  };
}

function getFastMovingForLocation_(location, limit) {
  var iss = getSheetData_(KE.SHEETS.ISSUE);
  iss.rows = applySheetLocationFilter_(KE.SHEETS.ISSUE, iss.rows, iss.headers, location);
  var cItem = findCol_(iss.headers, "Item");
  var cQty = findCol_(iss.headers, "Qty");
  var counts = {};
  for (var i = 0; i < iss.rows.length; i++) {
    var item = normalize_(iss.rows[i][cItem]);
    if (!item) {
      continue;
    }
    counts[item] = (counts[item] || 0) + (Number(iss.rows[i][cQty]) || 1);
  }
  return Object.keys(counts)
    .map(function (k) { return { item: k, count: counts[k] }; })
    .sort(function (a, b) { return b.count - a.count; })
    .slice(0, limit || 3)
    .map(function (x) { return x.item + " (" + x.count + ")"; });
}

function getDeadStockForLocation_(location, limit) {
  var movement = getMovementAnalytics_(location);
  var dead = movement.deadStock || [];
  return {
    total: movement.deadStockCount || dead.length,
    items: dead.slice(0, limit || 3).map(function (d) {
      return d.itemName || d.itemCode || "—";
    })
  };
}

/**
 * Full location-wise weekly summary so report can show complete counts per location.
 */
function buildLocationWeeklySummary_(sinceMs, deadByLocation) {
  var deadMap = {};
  (deadByLocation || []).forEach(function (d) {
    deadMap[d.location] = Number(d.count) || 0;
  });
  return KE.LOCATIONS.map(function (loc) {
    var invLoc = getInventoryControlCenter_(loc, true);
    var units = 0;
    var lowOut = 0;
    invLoc.forEach(function (it) {
      units += Number(it.totalQty) || 0;
      if (it.status === "Low stock" || it.status === "Out of stock") {
        lowOut++;
      }
    });
    return {
      location: loc,
      units: units,
      lowOut: lowOut,
      deadStock: deadMap[loc] || 0,
      requests: countSheetRowsSinceForLocation_(KE.SHEETS.REQUEST, "Date", sinceMs, loc),
      payments: countSheetRowsSinceForLocation_(KE.SHEETS.PAYMENT, "Date", sinceMs, loc),
      orders: countSheetRowsSinceForLocation_(KE.SHEETS.ORDER, "Order Date", sinceMs, loc),
      received: countSheetRowsSinceForLocation_(KE.SHEETS.RECEIVED, "Date", sinceMs, loc),
      issues: countSheetRowsSinceForLocation_(KE.SHEETS.ISSUE, "Date", sinceMs, loc),
      transfers: countSheetRowsSinceForLocation_(KE.SHEETS.TRANSFER, "Date", sinceMs, loc),
      revenue: sumPaymentsSinceForLocation_(sinceMs, loc)
    };
  });
}

/**
 * Dead stock by location (all-location report section).
 * Dead = no issue movement in last KE.DEAD_STOCK_DAYS.
 */
function getDeadStockByLocation_() {
  var iss = getSheetData_(KE.SHEETS.ISSUE);
  var cIssueItem = findCol_(iss.headers, "Item");
  var cIssueDate = findCol_(iss.headers, "Date");
  var lastIssued = {};
  for (var i = 0; i < iss.rows.length; i++) {
    var item = normalize_(iss.rows[i][cIssueItem]);
    if (!item) {
      continue;
    }
    var d = iss.rows[i][cIssueDate];
    var dt = d instanceof Date ? d : new Date(String(d || ""));
    if (isNaN(dt.getTime())) {
      continue;
    }
    var ms = dt.getTime();
    if (!lastIssued[item] || ms > lastIssued[item]) {
      lastIssued[item] = ms;
    }
  }

  var inv = getSheetData_(KE.SHEETS.INVENTORY);
  var cInvCode = findCol_(inv.headers, "Item Code");
  var cInvName = findCol_(inv.headers, "Item Name");
  var cInvLoc = findCol_(inv.headers, "Location");
  var cutoff = Date.now() - KE.DEAD_STOCK_DAYS * 24 * 60 * 60 * 1000;
  var map = {};
  KE.LOCATIONS.forEach(function (loc) {
    map[loc] = { location: loc, count: 0 };
  });
  for (var j = 0; j < inv.rows.length; j++) {
    var loc = resolveLocationName_(inv.rows[j][cInvLoc]);
    if (!loc) {
      continue;
    }
    if (!map[loc]) {
      map[loc] = { location: loc, count: 0 };
    }
    var code = normalize_(inv.rows[j][cInvCode]);
    var name = normalize_(inv.rows[j][cInvName]);
    var last = lastIssued[name] || lastIssued[code] || 0;
    if (!last || last < cutoff) {
      map[loc].count += 1;
    }
  }

  return Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) { return a.location.localeCompare(b.location); });
}

function countSheetRowsSince_(sheetName, dateField, sinceMs) {
  try {
    var data = getSheetData_(sheetName);
    var c = findCol_(data.headers, dateField);
    if (c < 0) {
      return data.rows.length;
    }
    var count = 0;
    for (var i = 0; i < data.rows.length; i++) {
      var d = data.rows[i][c];
      var dt = d instanceof Date ? d : new Date(String(d || ""));
      if (!isNaN(dt.getTime()) && dt.getTime() >= sinceMs) {
        count++;
      }
    }
    return count;
  } catch (e) {
    return 0;
  }
}

function countSheetRowsSinceForLocation_(sheetName, dateField, sinceMs, viewLocation) {
  try {
    var data = getSheetData_(sheetName);
    data.rows = applySheetLocationFilter_(sheetName, data.rows, data.headers, viewLocation);
    var c = findCol_(data.headers, dateField);
    if (c < 0) {
      return data.rows.length;
    }
    var count = 0;
    for (var i = 0; i < data.rows.length; i++) {
      var d = data.rows[i][c];
      var dt = d instanceof Date ? d : new Date(String(d || ""));
      if (!isNaN(dt.getTime()) && dt.getTime() >= sinceMs) {
        count++;
      }
    }
    return count;
  } catch (e) {
    return 0;
  }
}

function sumPaymentsSinceForLocation_(sinceMs, viewLocation) {
  try {
    var pay = getSheetData_(KE.SHEETS.PAYMENT);
    pay.rows = applySheetLocationFilter_(KE.SHEETS.PAYMENT, pay.rows, pay.headers, viewLocation);
    var cAmt = findCol_(pay.headers, "Amount");
    var cDate = findCol_(pay.headers, "Date");
    var total = 0;
    for (var i = 0; i < pay.rows.length; i++) {
      var d = pay.rows[i][cDate];
      var dt = d instanceof Date ? d : new Date(String(d || ""));
      if (!isNaN(dt.getTime()) && dt.getTime() >= sinceMs) {
        total += Number(pay.rows[i][cAmt]) || 0;
      }
    }
    return total;
  } catch (e) {
    return 0;
  }
}

/**
 * Vendor profitability: revenue (selling price * issued qty) minus cost
 * (cost price * issued qty) per vendor. Uses the Pricing master for prices.
 * Used by both the weekly report and the Finance tab.
 */
function computeVendorProfitability_() {
  // 1. Build item → vendor map (from inventory master)
  var inv = getSheetData_(KE.SHEETS.INVENTORY);
  var iCode = findCol_(inv.headers, "Item Code");
  var iName = findCol_(inv.headers, "Item Name");
  var iVen  = findCol_(inv.headers, "Vendor");
  var iCost = findCol_(inv.headers, "Cost Price");
  var itemToVendor = {};
  var itemCostFallback = {};
  for (var i = 0; i < inv.rows.length; i++) {
    var code = normalize_(inv.rows[i][iCode]);
    var name = normalize_(inv.rows[i][iName]);
    var vendor = iVen >= 0 ? normalize_(inv.rows[i][iVen]) : "";
    var cost = iCost >= 0 ? (Number(inv.rows[i][iCost]) || 0) : 0;
    if (code) {
      itemToVendor[code] = vendor;
      if (cost > 0) itemCostFallback[code] = cost;
    }
    if (name) {
      if (!itemToVendor[name]) itemToVendor[name] = vendor;
      if (cost > 0 && !itemCostFallback[name]) itemCostFallback[name] = cost;
    }
  }

  // 2. Build code → {cost, sell} map (from pricing master)
  var pricing = getSheetData_(KE.SHEETS.PRICING);
  var pCode = findCol_(pricing.headers, "Item Code");
  var pName = findCol_(pricing.headers, "Item Name");
  var pCost = findCol_(pricing.headers, "Cost Price");
  var pSell = findCol_(pricing.headers, "Selling Price");
  var priceByKey = {};
  for (var j = 0; j < pricing.rows.length; j++) {
    var pc = normalize_(pricing.rows[j][pCode]);
    var pn = normalize_(pricing.rows[j][pName]);
    var entry = {
      cost: Number(pricing.rows[j][pCost]) || 0,
      sell: Number(pricing.rows[j][pSell]) || 0
    };
    if (pc) priceByKey[pc] = entry;
    if (pn) priceByKey[pn] = entry;
  }

  // 3. Walk ISSUE_REGISTER and aggregate by vendor
  var iss = getSheetData_(KE.SHEETS.ISSUE);
  var isItem = findCol_(iss.headers, "Item");
  var isQty  = findCol_(iss.headers, "Qty");
  var agg = {};
  for (var k = 0; k < iss.rows.length; k++) {
    var key = normalize_(iss.rows[k][isItem]);
    if (!key) {
      continue;
    }
    var qty = Number(iss.rows[k][isQty]) || 0;
    var vendor = itemToVendor[key] || "—";
    var price = priceByKey[key] || { cost: itemCostFallback[key] || 0, sell: 0 };
    var revenue = qty * price.sell;
    var cogs = qty * price.cost;
    if (!agg[vendor]) {
      agg[vendor] = { vendor: vendor, units: 0, revenue: 0, cogs: 0, profit: 0, margin: 0 };
    }
    agg[vendor].units += qty;
    agg[vendor].revenue += revenue;
    agg[vendor].cogs += cogs;
  }

  var out = Object.keys(agg).map(function (v) {
    var row = agg[v];
    row.profit = row.revenue - row.cogs;
    row.margin = row.revenue > 0 ? Math.round((row.profit / row.revenue) * 10000) / 100 : 0;
    return row;
  });
  out.sort(function (a, b) { return b.profit - a.profit; });
  return out;
}

/**
 * Admin-only wrapper, callable from the dashboard.
 */
function getVendorProfitability(token) {
  assertAdmin_(token);
  return sanitizeForClient_({ success: true, vendors: computeVendorProfitability_() });
}

function renderWeeklyReportHtml_(r) {
  var money = function (n) {
    var v = Number(n) || 0;
    return "&#8377;" + v.toLocaleString("en-IN", { maximumFractionDigits: 0 });
  };
  var th = function (txt) { return '<th style="padding:6px 10px;border-bottom:1px solid #ddd;text-align:left;font-size:12px">' + txt + '</th>'; };
  var td = function (txt, align) {
    return '<td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;' +
      (align ? 'text-align:' + align + ';' : '') + '">' + (txt == null ? "—" : txt) + '</td>';
  };

  var html = '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;color:#222">';
  html += '<h2 style="color:#1f4e79;margin:0 0 4px">📊 Kings Equestrian – Weekly Report</h2>';
  html += '<p style="color:#666;margin:0 0 16px">Period: <strong>' + r.period.start + '</strong> to <strong>' + r.period.end + '</strong></p>';

  // KPI Grid
  html += '<table style="border-collapse:collapse;width:100%;margin-bottom:16px">';
  html += '<tr>';
  html += '<td style="padding:10px;background:#eaf3fb;border-radius:6px;width:25%;text-align:center"><div style="font-size:11px;color:#666">Total units</div><div style="font-size:20px;font-weight:600">' + r.totals.totalUnits + '</div></td>';
  html += '<td style="padding:10px;background:#fff5e0;border-radius:6px;width:25%;text-align:center"><div style="font-size:11px;color:#666">Inventory value</div><div style="font-size:20px;font-weight:600">' + money(r.totals.totalValue) + '</div></td>';
  html += '<td style="padding:10px;background:#fce8e8;border-radius:6px;width:25%;text-align:center"><div style="font-size:11px;color:#666">Low / out</div><div style="font-size:20px;font-weight:600">' + r.totals.lowStockCount + '</div></td>';
  html += '<td style="padding:10px;background:#fdf0e1;border-radius:6px;width:25%;text-align:center"><div style="font-size:11px;color:#666">Dead stock</div><div style="font-size:20px;font-weight:600">' + r.totals.deadStockCount + '</div></td>';
  html += '</tr></table>';

  // Full all-location breakdown
  if (r.locationWeekly && r.locationWeekly.length) {
    html += '<h3 style="color:#1f4e79;margin:18px 0 6px">📍 Location-wise full summary</h3>';
    html += '<table style="border-collapse:collapse;width:100%">';
    html += '<tr>'
      + th("Location")
      + th("Units")
      + th("Low/Out")
      + th("Dead")
      + th("Req")
      + th("Pay")
      + th("Ord")
      + th("Rec")
      + th("Iss")
      + th("Trf")
      + th("Revenue")
      + '</tr>';
    r.locationWeekly.forEach(function (loc) {
      html += '<tr>'
        + td(loc.location)
        + td(loc.units, "right")
        + td(loc.lowOut, "right")
        + td(loc.deadStock, "right")
        + td(loc.requests, "right")
        + td(loc.payments, "right")
        + td(loc.orders, "right")
        + td(loc.received, "right")
        + td(loc.issues, "right")
        + td(loc.transfers, "right")
        + td(money(loc.revenue), "right")
        + '</tr>';
    });
    html += '<tr style="background:#fafafa;font-weight:600">'
      + td("TOTAL")
      + td(r.totals.totalUnits, "right")
      + td(r.totals.lowStockCount, "right")
      + td(r.totals.deadStockCount, "right")
      + td(r.weekActivity.requests, "right")
      + td(r.weekActivity.payments, "right")
      + td(r.weekActivity.orders, "right")
      + td(r.weekActivity.received, "right")
      + td(r.weekActivity.issues, "right")
      + td(r.weekActivity.transfers, "right")
      + td(money(r.weekRevenue), "right")
      + '</tr>';
    html += '</table>';
  }

  // Explicit location-wise movement briefing for low/fast/dead
  if (r.locationMovementBrief && r.locationMovementBrief.length) {
    html += '<h3 style="color:#1f4e79;margin:18px 0 6px">📌 Location-wise low / fast / dead briefing</h3>';
    html += '<table style="border-collapse:collapse;width:100%">';
    html += '<tr>'
      + th("Location")
      + th("Low/Out (count)")
      + th("Top low items")
      + th("Top fast-moving")
      + th("Dead (count)")
      + th("Top dead items")
      + '</tr>';
    r.locationMovementBrief.forEach(function (row) {
      html += '<tr>'
        + td(row.location)
        + td(row.lowCount, "right")
        + td((row.lowTop && row.lowTop.length) ? row.lowTop.join(", ") : "—")
        + td((row.fastTop && row.fastTop.length) ? row.fastTop.join(", ") : "—")
        + td(row.deadCount, "right")
        + td((row.deadTop && row.deadTop.length) ? row.deadTop.join(", ") : "—")
        + '</tr>';
    });
    html += '</table>';
  }

  // Activity & revenue
  html += '<h3 style="color:#1f4e79;margin:18px 0 6px">📈 This Week</h3>';
  html += '<table style="border-collapse:collapse;width:100%">';
  html += '<tr>' + th("Requests") + th("Payments") + th("Orders") + th("Received") + th("Issued") + th("Transfers") + th("Revenue") + '</tr>';
  html += '<tr>'
    + td(r.weekActivity.requests, "center")
    + td(r.weekActivity.payments, "center")
    + td(r.weekActivity.orders, "center")
    + td(r.weekActivity.received, "center")
    + td(r.weekActivity.issues, "center")
    + td(r.weekActivity.transfers, "center")
    + td(money(r.weekRevenue), "right")
    + '</tr></table>';

  // Low stock
  if (r.lowStock && r.lowStock.length) {
    html += '<h3 style="color:#a91e1e;margin:18px 0 6px">⚠️ Low / out of stock (top ' + r.lowStock.length + ')</h3>';
    html += '<table style="border-collapse:collapse;width:100%">';
    html += '<tr>' + th("Code") + th("Item") + th("Qty") + th("Min") + th("Status") + '</tr>';
    r.lowStock.forEach(function (it) {
      html += '<tr>' + td(it.itemCode) + td(it.itemName) + td(it.totalQty, "right") + td(it.minLevel, "right") + td(it.status) + '</tr>';
    });
    html += '</table>';
  }

  // Vendor profitability (admin-only data)
  if (r.vendorProfit && r.vendorProfit.length) {
    html += '<h3 style="color:#1f4e79;margin:18px 0 6px">🏭 Vendor profitability (top 10)</h3>';
    html += '<table style="border-collapse:collapse;width:100%">';
    html += '<tr>' + th("Vendor") + th("Units") + th("Revenue") + th("COGS") + th("Profit") + th("Margin %") + '</tr>';
    r.vendorProfit.forEach(function (v) {
      html += '<tr>'
        + td(v.vendor)
        + td(v.units, "right")
        + td(money(v.revenue), "right")
        + td(money(v.cogs), "right")
        + td(money(v.profit), "right")
        + td((v.margin || 0) + "%", "right")
        + '</tr>';
    });
    html += '</table>';
  }

  // Fast moving + Dead stock
  if (r.fastMoving && r.fastMoving.length) {
    html += '<h3 style="color:#1f7a3a;margin:18px 0 6px">🔥 Fast moving</h3>';
    html += '<table style="border-collapse:collapse;width:100%">';
    html += '<tr>' + th("Item") + th("Issued units") + '</tr>';
    r.fastMoving.forEach(function (f) {
      html += '<tr>' + td(f.item) + td(f.count, "right") + '</tr>';
    });
    html += '</table>';
  }

  if (r.deadStock && r.deadStock.length) {
    html += '<h3 style="color:#a55401;margin:18px 0 6px">💀 Dead stock (no movement &gt; 90d)</h3>';
    html += '<table style="border-collapse:collapse;width:100%">';
    html += '<tr>' + th("Code") + th("Item") + '</tr>';
    r.deadStock.forEach(function (d) {
      html += '<tr>' + td(d.itemCode) + td(d.itemName) + '</tr>';
    });
    html += '</table>';
  }

  html += '<p style="color:#666;font-size:11px;margin-top:24px">Automated report from the Kings Equestrian inventory system. Reply to this email if any number looks off.</p>';
  html += '</div>';
  return html;
}
