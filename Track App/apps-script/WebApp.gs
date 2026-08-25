/**
 * Web app + dashboard API (password session + location scope)
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : "login";
  if (page === "login") {
    var loginT = HtmlService.createTemplateFromFile("Login");
    try {
      loginT.webAppUrl = getWebAppUrl();
    } catch (err) {
      loginT.webAppUrl = "";
    }
    return loginT.evaluate()
      .setTitle("Kings Equestrian Inventory – Login")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var dashT = HtmlService.createTemplateFromFile("Dashboard");
  dashT.urlToken = (e && e.parameter && e.parameter.token) ? String(e.parameter.token) : "";
  return dashT.evaluate()
    .setTitle("Kings Equestrian Inventory")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  return url || "";
}

function buildAppPageUrl(page, token) {
  var url = getWebAppUrl();
  if (!url) {
    throw new Error("Deploy the web app first (Deploy → New deployment).");
  }
  var dest = url + "?page=" + encodeURIComponent(page || "dashboard");
  if (token) {
    dest += "&token=" + encodeURIComponent(token);
  }
  return dest;
}

function getAppContext(token, viewLocation) {
  return sanitizeForClient_(getLocationContext_(token, viewLocation));
}

function getInventoryControlCenterForUser_(token, viewLocation) {
  var user = validateSessionToken_(token);
  var view = resolveViewLocation_(user, viewLocation);
  return sanitizeForClient_(buildInventoryAnalyticsPayload_(view, user));
}

/**
 * Lightweight payload for post-save UI refresh.
 * Faster than full bootstrap because it avoids nav/permissions re-hydration.
 */
function getPostSaveRefresh(token, viewLocation) {
  var user = validateSessionToken_(token);
  var locCtx = getLocationContext_(token, viewLocation);
  var view = locCtx.viewLocation;
  return sanitizeForClient_({
    lists: getMasterLists_(view),
    summary: getDashboardSummary_(view, user),
    location: locCtx
  });
}

function getDashboardBootstrap(token, viewLocation) {
  try {
    var user = validateSessionToken_(token);
    var locCtx = getLocationContext_(token, viewLocation);
    var view = locCtx.viewLocation;
    var isAdmin = isAdminRole_(user.role);
    var payload = {
      authenticated: true,
      user: user,
      location: locCtx,
      locations: KE.LOCATIONS,
      permissions: {
        master: canAccess_("master", user),
        request: canAccess_("request", user),
        payment: canAccess_("payment", user),
        order: canAccess_("order", user),
        receive: canAccess_("receive", user),
        issue: canAccess_("issue", user),
        transfer: canTransferStock_(user),
        sample: canManageSamples_(user),
        pricing: isAdmin,
        finance: canViewFinancials_(user),
        admin: isAdmin
      },
      inventoryTypes: KE.INVENTORY_TYPES,
      lists: getMasterLists_(view),
      // Registers are fetched lazily per tab to reduce initial dashboard load time.
      registers: {},
      summary: getDashboardSummary_(view, user),
      webAppUrl: getWebAppUrl()
    };
    return sanitizeForClient_(payload);
  } catch (err) {
    return {
      authenticated: false,
      message: err.message || "Please sign in."
    };
  }
}

function getMasterLists_(viewLocation) {
  var vendors = sheetColumnForView_(KE.SHEETS.VENDOR, "Vendor Name", viewLocation);
  var items = sheetColumnForView_(KE.SHEETS.INVENTORY, "Item Name", viewLocation);
  var itemCodes = sheetColumnForView_(KE.SHEETS.INVENTORY, "Item Code", viewLocation);
  var itemOptions = getItemOptionsForView_(viewLocation);
  var requests = getOpenRequests_(viewLocation);
  var orders = getOrdersForReceiveDropdown_(viewLocation);
  var storageLocations = getStorageLocations_(viewLocation);
  return {
    vendors: vendors,
    items: items,
    itemCodes: itemCodes,
    itemOptions: itemOptions,
    requests: requests,
    orders: orders,
    storageLocations: storageLocations
  };
}

function getTransferLog_(payload) {
  var prep = prepPayload_(payload);
  validateSessionToken_(prep.token);
  var viewLocation = normalize_((prep.payload && prep.payload.viewLocation) || "");
  var data = getSheetData_(KE.SHEETS.TRANSFER);
  var rows = data.rows.map(function(r) {
    var obj = {};
    data.headers.forEach(function(h, i) { obj[h] = serializeCellValue_(r[i]); });
    return obj;
  });
  // Filter by location if not admin all-locations
  if (viewLocation && !isAllLocations_(viewLocation)) {
    rows = rows.filter(function(r) {
      return r["From Location"] === viewLocation || r["To Location"] === viewLocation;
    });
  }
  // Sort newest first
  rows.sort(function(a, b) {
    var da = a["Date"] ? new Date(a["Date"]).getTime() : 0;
    var db = b["Date"] ? new Date(b["Date"]).getTime() : 0;
    return db - da;
  });
  return { success: true, transfers: rows };
}

function getStorageLocations_(viewLocation) {
  try {
    var data = getSheetData_(KE.SHEETS.STORAGE);
    var cName = findCol_(data.headers, "Storage Name");
    var cLoc  = findCol_(data.headers, "Location");
    var out = [];
    var seen = {};
    for (var i = 0; i < data.rows.length; i++) {
      var name = cName >= 0 ? normalize_(data.rows[i][cName]) : "";
      var loc  = cLoc  >= 0 ? normalize_(data.rows[i][cLoc])  : "";
      if (!name) continue;
      if (viewLocation && !isAllLocations_(viewLocation) && loc && loc !== viewLocation) continue;
      if (seen[name + "|" + loc]) continue;
      seen[name + "|" + loc] = true;
      out.push({ name: name, location: loc, label: loc ? name + " (" + loc + ")" : name });
    }
    return out.sort(function(a, b){ return a.name.localeCompare(b.name); });
  } catch(e) {
    return [];
  }
}

function addStorageLocation(payload) {
  var prep = prepPayload_(payload);
  payload = prep.payload;
  assertCan_("master", payload.location, prep.token);
  var name = normalize_(payload.storageName);
  var loc  = normalize_(payload.location);
  var desc = normalize_(payload.description);
  if (!name) throw new Error("Storage name is required.");
  validateLocation_(loc);
  var user = validateSessionToken_(prep.token);
  var id = generateID_("STO");
  appendRow_(KE.SHEETS.STORAGE, [id, name, loc, desc, user.name || user.email, todayStr_()]);
  return ok_("Storage location added.", { storageId: id });
}

function getItemOptionsForView_(viewLocation) {
  var data = getSheetData_(KE.SHEETS.INVENTORY);
  data.rows = applySheetLocationFilter_(KE.SHEETS.INVENTORY, data.rows, data.headers, viewLocation);
  var cCode = findCol_(data.headers, "Item Code");
  var cName = findCol_(data.headers, "Item Name");
  var cImg  = findCol_(data.headers, "Image URL");
  var seenKey = {};
  var firstImage = {};
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    var code = cCode >= 0 ? normalize_(data.rows[i][cCode]) : "";
    var name = cName >= 0 ? normalize_(data.rows[i][cName]) : "";
    var img  = cImg  >= 0 ? normalize_(data.rows[i][cImg])  : "";
    if (!code && !name) continue;
    // Track first non-empty image per code/name so even later blank rows still resolve to a thumbnail.
    if (img) {
      if (code && !firstImage[code]) firstImage[code] = img;
      if (name && !firstImage[name]) firstImage[name] = img;
    }
    var key = (code + "|" + name).toLowerCase();
    if (seenKey[key]) continue;
    seenKey[key] = true;
    out.push({
      code: code,
      name: name,
      imageUrl: img,
      label: name && code ? (name + " (" + code + ")") : (name || code)
    });
  }
  // Backfill imageUrl where the duplicate row had the picture but the kept row didn't.
  out.forEach(function (it) {
    if (!it.imageUrl) {
      it.imageUrl = firstImage[it.code] || firstImage[it.name] || "";
    }
  });
  out.sort(function (a, b) { return (a.name || a.code).localeCompare(b.name || b.code); });
  return out;
}

function getOrdersForReceiveDropdown_(viewLocation) {
  var data = getSheetData_(KE.SHEETS.ORDER);
  data.rows = applySheetLocationFilter_(KE.SHEETS.ORDER, data.rows, data.headers, viewLocation);
  var cId = findCol_(data.headers, "Order ID");
  var cSt = findCol_(data.headers, "Status");
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    var id = normalize_(data.rows[i][cId]);
    var st = cSt >= 0 ? normalize_(data.rows[i][cSt]) : "";
    if (!id) {
      continue;
    }
    if (st === KE.ORDER_STATUS.RECEIVED || st === KE.ORDER_STATUS.CANCELLED) {
      continue;
    }
    out.push({ id: id, status: st || "Placed" });
  }
  return out;
}

function getOpenRequests_(viewLocation) {
  var data = getSheetData_(KE.SHEETS.REQUEST);
  data.rows = applySheetLocationFilter_(KE.SHEETS.REQUEST, data.rows, data.headers, viewLocation);
  var cId = findCol_(data.headers, "Request ID");
  var cStatus = findCol_(data.headers, "Status");
  var cVendor = findCol_(data.headers, ["Vendor", "Vendor Name"]);
  var cItem = findCol_(data.headers, "Item");
  var cQty = findCol_(data.headers, "Qty");
  var cLoc = findCol_(data.headers, "Location");
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    var st = normalize_(data.rows[i][cStatus]);
    if (st !== KE.REQUEST_STATUS.COMPLETED && st !== KE.REQUEST_STATUS.CANCELLED) {
      out.push({
        id: normalize_(data.rows[i][cId]),
        status: st,
        vendor: cVendor >= 0 ? normalize_(data.rows[i][cVendor]) : "",
        item: cItem >= 0 ? normalize_(data.rows[i][cItem]) : "",
        qty: cQty >= 0 ? Number(data.rows[i][cQty]) || 0 : 0,
        location: cLoc >= 0 ? normalize_(data.rows[i][cLoc]) : ""
      });
    }
  }
  return out;
}

function getDashboardSummary_(viewLocation, user) {
  var inv = getSheetData_(KE.SHEETS.INVENTORY);
  inv.rows = applySheetLocationFilter_(KE.SHEETS.INVENTORY, inv.rows, inv.headers, viewLocation);
  var cQty = findCol_(inv.headers, "Current Qty");
  var cMin = findCol_(inv.headers, "Min Level");
  var cName = findCol_(inv.headers, "Item Name");
  var cCode = findCol_(inv.headers, "Item Code");
  var cLoc = findCol_(inv.headers, "Location");
  var cImg = findCol_(inv.headers, "Image URL");
  var lowStock = [];
  var skuItems = [];
  var totalSkus = inv.rows.length;
  for (var i = 0; i < inv.rows.length; i++) {
    var qty = Number(inv.rows[i][cQty]) || 0;
    var min = Number(inv.rows[i][cMin]) || 0;
    var nameI = cName >= 0 ? normalize_(inv.rows[i][cName]) : "";
    var imgI  = cImg  >= 0 ? normalize_(inv.rows[i][cImg])  : "";
    skuItems.push({
      itemCode: cCode >= 0 ? normalize_(inv.rows[i][cCode]) : "",
      itemName: nameI,
      qty: qty,
      location: cLoc >= 0 ? normalize_(inv.rows[i][cLoc]) : "",
      minLevel: min,
      imageUrl: imgI
    });
    if (min > 0 && qty <= min) {
      lowStock.push({ itemName: nameI, imageUrl: imgI });
    }
  }

  var req = getSheetData_(KE.SHEETS.REQUEST);
  req.rows = applySheetLocationFilter_(KE.SHEETS.REQUEST, req.rows, req.headers, viewLocation);
  var cSt = findCol_(req.headers, "Status");
  var pending = 0;
  var approved = 0;
  for (var j = 0; j < req.rows.length; j++) {
    var s = normalize_(req.rows[j][cSt]);
    if (s === KE.REQUEST_STATUS.PENDING || s === "Unpaid") {
      pending++;
    }
    if (s === KE.REQUEST_STATUS.APPROVED || s === KE.REQUEST_STATUS.ORDERED) {
      approved++;
    }
  }

  var pay = getSheetData_(KE.SHEETS.PAYMENT);
  pay.rows = applySheetLocationFilter_(KE.SHEETS.PAYMENT, pay.rows, pay.headers, viewLocation);
  var ord = getSheetData_(KE.SHEETS.ORDER);
  ord.rows = applySheetLocationFilter_(KE.SHEETS.ORDER, ord.rows, ord.headers, viewLocation);
  var grn = getSheetData_(KE.SHEETS.RECEIVED);
  grn.rows = applySheetLocationFilter_(KE.SHEETS.RECEIVED, grn.rows, grn.headers, viewLocation);
  var iss = getSheetData_(KE.SHEETS.ISSUE);
  iss.rows = applySheetLocationFilter_(KE.SHEETS.ISSUE, iss.rows, iss.headers, viewLocation);

  var analytics = null;
  try {
    analytics = buildInventoryAnalyticsPayload_(viewLocation, user || {});
  } catch (e) {
    analytics = null;
  }

  return {
    totalSkus: totalSkus,
    lowStockCount: lowStock.length,
    lowStockItems: lowStock.slice(0, 10),
    pendingRequests: pending,
    activeRequests: approved,
    paymentCount: pay.rows.length,
    orderCount: ord.rows.length,
    receivedCount: grn.rows.length,
    issueCount: iss.rows.length,
    viewLocation: viewLocation || "All",
    skuItems: skuItems.slice(0, 12),
    analytics: analytics
  };
}

function apiCall(action, payload, token) {
  try {
    payload = payload || {};
    token = token || payload.token || "";
    var viewLocation = payload.viewLocation != null ? payload.viewLocation : "";
    if (payload.viewLocation != null) {
      delete payload.viewLocation;
    }
    switch (action) {
      case "logout": return logoutUser(token);
      case "refresh":
        return sanitizeForClient_(ok_("Refreshed", {
          bootstrap: getDashboardBootstrap(token, viewLocation)
        }));
      case "getRegisterPage":
        return ok_("OK", {
          register: getRegisterPage(
            token,
            viewLocation,
            payload.registerKey,
            payload.page,
            payload.dateFrom,
            payload.dateTo,
            payload.status
          )
        });
      case "getInventoryControlCenter":
        return ok_("OK", {
          analytics: getInventoryControlCenterForUser_(token, viewLocation)
        });
      case "getFinancialKpis":
        return ok_("OK", {
          finance: getFinancialKpis(token, viewLocation)
        });
      case "getPricingList":
        return ok_("OK", {
          pricing: getPricingList(token)
        });
      case "getVendorProfitability":
        return getVendorProfitability(token);
      case "sendWeeklyReportNow":
        assertAdmin_(token);
        return sanitizeForClient_(ok_("Report sent", sendWeeklyInventoryReport()));
      default:
        payload.token = token;
        payload.viewLocation = viewLocation;
        switch (action) {
          case "addVendor": return addVendor(payload);
          case "addNewItem": return addNewItem(payload);
          case "createRequest": return createRequest(payload);
          case "cancelRequest": return cancelRequest(payload);
          case "recordPayment": return recordPayment(payload);
          case "placeOrder": return placeOrder(payload);
          case "receiveGoods": return receiveGoods(payload);
          case "issueItem": return issueItem(payload);
          case "createStockTransfer": return createStockTransfer(payload);
          case "approveStockTransfer": return approveStockTransfer(payload);
          case "cancelStockTransfer": return cancelStockTransfer(payload);
          case "upsertPricing": return upsertPricing(payload);
          case "addStorageLocation": return addStorageLocation(payload);
          case "getTransferLog": return getTransferLog_(payload);
          case "createSampleMovement": return createSampleMovement(payload);
          case "returnSampleMovement": return returnSampleMovement(payload);
          case "markSampleLost": return markSampleLost(payload);
          case "getSampleMovements": return getSampleMovements(payload);
          default:
            return fail_("Unknown action: " + action);
        }
    }
  } catch (err) {
    return fail_(err.message || String(err));
  }
}
