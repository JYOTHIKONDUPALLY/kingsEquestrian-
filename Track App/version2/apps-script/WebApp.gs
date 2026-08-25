/**
 * Web app entry + API router (Version 2)
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
      .setTitle("KE Track v2 – Login")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var dashT = HtmlService.createTemplateFromFile("Dashboard");
  dashT.urlToken = (e && e.parameter && e.parameter.token) ? String(e.parameter.token) : "";
  return dashT.evaluate()
    .setTitle("KE Track v2")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl() || "";
}

function getDashboardBootstrap(token, viewLocation) {
  try {
    var user = validateSessionToken_(token);
    var locCtx = getLocationContext_(token, viewLocation);
    var view = locCtx.viewLocation;
    var isAdmin = isAdminRole_(user.role);
    return sanitizeForClient_({
      authenticated: true,
      user: user,
      location: locCtx,
      locations: KE.LOCATIONS,
      permissions: {
        master: canAccess_("master", user),
        request: canAccess_("request", user),
        inventory: canAccess_("inventory", user),
        issue: canAccess_("issue", user),
        admin: isAdmin,
        view: true
      },
      lists: getMasterLists_(view),
      summary: getDashboardSummary_(view, user),
      webAppUrl: getWebAppUrl()
    });
  } catch (err) {
    return {
      authenticated: false,
      message: err.message || "Please sign in."
    };
  }
}

function apiCall(action, payload, token) {
  try {
    payload = payload || {};
    token = token || payload.token || "";
    var viewLocation = payload.viewLocation != null ? payload.viewLocation : "";
    if (payload.viewLocation != null) {
      delete payload.viewLocation;
    }
    payload.token = token;
    payload.viewLocation = viewLocation;

    switch (action) {
      case "logout":
        return logoutUser(token);
      case "refresh":
        return sanitizeForClient_(ok_("Refreshed", {
          bootstrap: getDashboardBootstrap(token, viewLocation)
        }));
      case "addVendor":
        return sanitizeForClient_(addVendor(payload));
      case "updateVendor":
        return sanitizeForClient_(updateVendor(payload));
      case "listVendors":
        return sanitizeForClient_(ok_("OK", listVendors(token, viewLocation, payload)));
      case "addItem":
        return sanitizeForClient_(addItem(payload));
      case "updateItem":
        return sanitizeForClient_(updateItem(payload));
      case "listItems":
        return sanitizeForClient_(ok_("OK", listItems(token, viewLocation, payload)));
      case "listInventory":
        return sanitizeForClient_(ok_("OK", listInventory(token, viewLocation, payload)));
      case "setInventoryQty":
      case "addInventoryQty":
        return sanitizeForClient_(addInventoryQty(payload));
      case "getStockQty":
        validateSessionToken_(token);
        return sanitizeForClient_(ok_("OK", {
          qty: getStockQty_(payload.itemCode || payload.item, payload.location)
        }));
      case "createRequest":
        return sanitizeForClient_(createRequest(payload));
      case "listRequests":
        return sanitizeForClient_(ok_("OK", listRequests(token, viewLocation, payload)));
      case "updateRequestStatus":
        return sanitizeForClient_(updateRequestStatus(payload));
      case "issueItem":
        return sanitizeForClient_(issueItem(payload));
      case "listIssues":
        return sanitizeForClient_(ok_("OK", listIssues(token, viewLocation, payload)));
      case "addStudent":
        return sanitizeForClient_(addStudent(payload));
      case "updateStudent":
        return sanitizeForClient_(updateStudent(payload));
      case "listStudents": {
        var user = validateSessionToken_(token);
        var studentLoc = normalize_(payload.filterLocation || payload.location) || viewLocation;
        if (!isAdminRole_(user.role)) {
          studentLoc = resolveViewLocation_(user, viewLocation);
        } else if (normalize_(payload.filterLocation || payload.location)) {
          studentLoc = resolveLocationName_(payload.filterLocation || payload.location);
        } else {
          studentLoc = resolveViewLocation_(user, viewLocation);
        }
        var studentList = listStudents(token, studentLoc, payload);
        return sanitizeForClient_(ok_("OK", {
          rows: studentList.rows,
          total: studentList.total,
          shown: studentList.shown,
          limit: studentList.limit,
          students: listStudentOptionsForView_(studentLoc)
        }));
      }
      case "listActivity":
        return sanitizeForClient_(ok_("OK", listActivity(token, viewLocation, payload)));
      case "sendWeeklyReport":
        return sanitizeForClient_(sendWeeklyReportNow(token));
      default:
        return fail_("Unknown action: " + action);
    }
  } catch (err) {
    return fail_(err.message || String(err));
  }
}

/** Sidebar helper when bound to a sheet */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Login")
    .setTitle("KE Track v2");
  SpreadsheetApp.getUi().showSidebar(html);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("KE Track v2")
    .addItem("Open Login Sidebar", "showSidebar")
    .addItem("Setup sheets", "setupInventorySheets")
    .addItem("Repair sheet validations (fix Issue E2 error)", "repairSheetValidations")
    .addItem("Seed sample data", "seedSampleData")
    .addItem("Link spreadsheet ID", "setSpreadsheetId")
    .addSeparator()
    .addItem("Send weekly report now", "sendWeeklyInventoryReport")
    .addItem("Install Monday 8am weekly report", "installWeeklyReportTrigger")
    .addItem("Remove weekly report trigger", "uninstallWeeklyReportTrigger")
    .addItem("Authorize item image upload (Drive)", "authorizeImageUploadDriveAccess")
    .addToUi();
}
