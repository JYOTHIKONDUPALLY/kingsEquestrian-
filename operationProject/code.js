// ============================================================
// KINGS EQUESTRIAN – COMPLETE APPS SCRIPT
// ============================================================

var SESSION_PREFIX = "ops_sess_";
var SESSION_TTL_SEC = 21600;
var __dashActivityCache_ = null;
var __locationsRecordsCache_ = null;
var __locationResolveCache_ = {};
var __sheetHeaderCache_ = {};
var __inventoryMinMapCache_ = null;

/** Max data rows to read per sheet on dashboard build (avoids huge getDataRange timeouts). */
var DASH_SHEET_ROW_LIMITS_ = {
  activity: 4000,
  health: 3000,
  inventory: 800,
  shoeing: 3000,
  vaccination: 3000,
  medical: 400
};

/** Use the spreadsheet this script is bound to, or SPREADSHEET_ID in Script Properties. */
function getSS_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }
  var id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  throw new Error(
    "No spreadsheet linked. Open your Kings Equestrian Google Sheet, " +
    "then in Apps Script run setSpreadsheetId() once, and redeploy."
  );
}

/** Run once from the script editor while your ops sheet is open. */
function setSpreadsheetId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Open the Kings Equestrian spreadsheet first.");
  }
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
  Logger.log("Saved SPREADSHEET_ID: " + ss.getId());
  return ss.getId();
}

// ─── WEB APP ENTRY POINT ─────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page ? e.parameter.page : "login").toLowerCase();
  if (page === "photo") {
    return serveProfilePhoto_(e);
  }
  var tab = (e && e.parameter && e.parameter.tab ? e.parameter.tab : "").toLowerCase();
  var pageMap = {
    login: "Login",
    dashboard: "Dashboard",
    "form-daily-stock-entry": "form-daily-stock-entry",
    "form-health-check": "form-health-check",
    "form-health-medical": "form-health-medical",
    "form-horse-activity": "form-daily-routine",
    "form-daily-routine": "form-daily-routine",
    "form-medical-record": "form-health-medical",
    "form-shoeing-record": "form-shoeing-record",
    "form-vaccination": "form-vaccination",
    "form-add-horse": "form-add-horse",
    "form-add-trainer": "form-add-trainer",
    "form-add-groom": "form-add-groom"
  };
  var fileName = pageMap[page] || "Login";

  if (page === "login") {
    var loginTemplate = HtmlService.createTemplateFromFile(fileName);
    try {
      loginTemplate.webAppUrl = getWebAppUrl();
    } catch (urlErr) {
      loginTemplate.webAppUrl = "";
    }
    return wrapHtmlOutput_(loginTemplate.evaluate());
  }

  var t = HtmlService.createTemplateFromFile(fileName);
  t.appPage = page;
  t.appTab = tab || (page === "dashboard" ? "dashboard" : "forms");
  var params = (e && e.parameter) ? e.parameter : {};
  t.editHorseId = String(params.horseId || params.horseid || params.HorseId || "").trim();
  t.editTrainerName = String(params.trainerName || params.trainername || "").trim();
  t.editGroomName = String(params.groomName || params.groomname || "").trim();
  return wrapHtmlOutput_(t.evaluate());
}

function wrapHtmlOutput_(output) {
  return output
    .setTitle("Kings Equestrian Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover");
}

/** Canonical deployed URL (use for redirects; iframe location is not /exec). */
function getWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    throw new Error("Web app URL unavailable. Deploy as web app first.");
  }
  return url;
}

function buildAppPageUrl(page, token, nextPage, tab, extraQuery) {
  var url = getWebAppUrl();
  var dest = url + "?page=" + encodeURIComponent(page || "login");
  if (token) {
    dest += "&token=" + encodeURIComponent(token);
  }
  if (nextPage) {
    dest += "&next=" + encodeURIComponent(nextPage);
  }
  if (tab) {
    dest += "&tab=" + encodeURIComponent(tab);
  }
  if (extraQuery && typeof extraQuery === "object") {
    var ek;
    for (ek in extraQuery) {
      if (extraQuery.hasOwnProperty(ek) && extraQuery[ek] !== undefined && extraQuery[ek] !== "") {
        dest += "&" + encodeURIComponent(ek) + "=" + encodeURIComponent(String(extraQuery[ek]));
      }
    }
  }
  return dest;
}

function getLoginUrl(nextPage) {
  return buildAppPageUrl("login", "", nextPage || "");
}

// ─── PASSWORD HASH (must match generateHash / Users sheet) ────
function hashPassword(plain) {
  plain = (plain || "").toString();
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain)
  );
}

// ─── USER LOGIN ───────────────────────────────────────────────
/** True if stored value is hash or plain text (migration-friendly). */
function passwordMatchesStored_(stored, plain, inputHash) {
  stored = String(stored || "").trim();
  plain = String(plain || "");
  if (!stored || !plain) {
    return false;
  }
  if (stored === inputHash) {
    return true;
  }
  if (stored === plain) {
    return true;
  }
  return false;
}

/** Some sheets put a readable password in an extra column (e.g. column H). */
function findMatchingPasswordInRow_(row, plain, inputHash, primaryCol) {
  if (primaryCol >= 0 && passwordMatchesStored_(row[primaryCol], plain, inputHash)) {
    return true;
  }
  for (var c = 0; c < row.length; c++) {
    if (c === primaryCol) {
      continue;
    }
    if (passwordMatchesStored_(row[c], plain, inputHash)) {
      return true;
    }
  }
  return false;
}

function createSessionForUser_(email, name, role, location) {
  var token = Utilities.getUuid();
  var locRec = resolveLocationKey_(location);
  var userObj = {
    email: String(email || "").trim(),
    name: String(name || "").trim(),
    role: role,
    location: locRec.name || String(location || "").trim(),
    locationId: locRec.id || ""
  };
  Logger.log("[KingsOps login session] raw Location cell=" + JSON.stringify(location) +
    " → stored location=" + userObj.location + " locationId=" + userObj.locationId);
  var sessionJson = JSON.stringify(userObj);
  CacheService.getScriptCache().put(SESSION_PREFIX + token, sessionJson, SESSION_TTL_SEC);
  try {
    PropertiesService.getScriptProperties().setProperty(SESSION_PREFIX + token, sessionJson);
  } catch (e) {
    Logger.log("Session property store: " + e);
  }
  return { token: token, user: userObj };
}

function loginUser(identifier, password) {
  identifier = (identifier || "").toString().trim().toLowerCase();
  password = (password || "").toString();
  if (!identifier || !password) {
    throw new Error("Please enter username/email and password.");
  }

  var usersSheet;
  try {
    usersSheet = getSS_().getSheetByName("Users");
  } catch (e) {
    throw new Error(
      (e && e.message ? e.message : String(e)) +
      " Run setSpreadsheetId() in Apps Script with your sheet open, then redeploy."
    );
  }
  if (!usersSheet) {
    throw new Error("Users sheet not found in spreadsheet.");
  }

  var inputHash = hashPassword(password);
  var data = usersSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error("No users in Users sheet.");
  }

  var map = getHeaderIndexMap_(usersSheet);
  var colEmail = findCol_(map, ["Email", "email"]);
  var colName = findCol_(map, ["Name", "name"]);
  var colRole = findCol_(map, ["Role", "role"]);
  var colLocation = findCol_(map, ["Location", "location"]);
  var colPassword = findCol_(map, ["Password", "password"]);
  if (colEmail < 0) {
    colEmail = 0;
  }
  if (colName < 0) {
    colName = 1;
  }
  if (colRole < 0) {
    colRole = 2;
  }
  if (colLocation < 0) {
    colLocation = 3;
  }
  if (colPassword < 0) {
    colPassword = 4;
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var email = (row[colEmail] || "").toString().trim();
    var name = (row[colName] || "").toString().trim();
    if (!email && !name) {
      continue;
    }
    var matchesIdentifier =
      email.toLowerCase() === identifier ||
      name.toLowerCase() === identifier ||
      name.toLowerCase().replace(/\s+/g, "") === identifier.replace(/\s+/g, "");
    if (!matchesIdentifier) {
      continue;
    }
    if (!findMatchingPasswordInRow_(row, password, inputHash, colPassword)) {
      throw new Error(
        "Invalid password. Use the password for this account, or ask admin to run hashUserPassword() in Apps Script."
      );
    }
    return createSessionForUser_(
      email,
      name,
      row[colRole],
      row[colLocation]
    );
  }

  throw new Error(
    "User not found. Try your full email (e.g. admin@kingsequestrian.com) or your name as in the Users sheet."
  );
}

/**
 * Run in Apps Script editor to store a hashed password in the Password column.
 * Example: hashUserPassword("admin@kingsequestrian.com", "admin@123");
 */
function hashUserPassword(userEmailOrName, plainPassword) {
  plainPassword = String(plainPassword || "");
  userEmailOrName = String(userEmailOrName || "").trim().toLowerCase();
  if (!userEmailOrName || !plainPassword) {
    throw new Error("Provide email/name and plain password.");
  }
  var sheet = getSS_().getSheetByName("Users");
  if (!sheet) {
    throw new Error("Users sheet not found.");
  }
  var hash = hashPassword(plainPassword);
  var data = sheet.getDataRange().getValues();
  var map = getHeaderIndexMap_(sheet);
  var colEmail = findCol_(map, ["Email", "email"]);
  var colName = findCol_(map, ["Name", "name"]);
  var colPassword = findCol_(map, ["Password", "password"]);
  if (colPassword < 0) {
    colPassword = 4;
  }
  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][colEmail >= 0 ? colEmail : 0] || "").trim().toLowerCase();
    var name = String(data[i][colName >= 0 ? colName : 1] || "").trim().toLowerCase();
    if (email === userEmailOrName || name === userEmailOrName) {
      sheet.getRange(i + 1, colPassword + 1).setValue(hash);
      Logger.log("Updated password hash for row " + (i + 1));
      return { ok: true, hash: hash };
    }
  }
  throw new Error("User not found in Users sheet.");
}

function validateSessionToken_(token) {
  token = (token || "").toString().trim();
  if (!token) {
    throw new Error("Unauthorized");
  }
  var raw = CacheService.getScriptCache().get(SESSION_PREFIX + token);
  if (!raw) {
    raw = PropertiesService.getScriptProperties().getProperty(SESSION_PREFIX + token);
  }
  if (!raw) {
    throw new Error("Unauthorized");
  }
  return JSON.parse(raw);
}

function getSessionUser(token) {
  var u = validateSessionToken_(token);
  var locRec = resolveLocationKey_(u.location || u.locationId);
  return {
    email: u.email,
    name: u.name,
    role: u.role,
    location: locRec.name || u.location || "",
    locationId: locRec.id || u.locationId || "",
    isAdmin: isAdminRole_(u.role)
  };
}

function isAdminRole_(role) {
  return String(role || "").trim().toLowerCase() === "admin";
}

function isTrainerRole_(role) {
  var r = String(role || "").trim().toLowerCase();
  return r.indexOf("trainer") >= 0 && !isAdminRole_(role);
}

function isGroomRole_(role) {
  var r = String(role || "").trim().toLowerCase();
  return r.indexOf("groom") >= 0 && !isAdminRole_(role);
}

function normalizeStaffName_(name) {
  return String(name || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function staffNameMatches_(assignedName, staffName) {
  var a = normalizeStaffName_(assignedName);
  var s = normalizeStaffName_(staffName);
  if (!a || !s) {
    return false;
  }
  return a === s;
}

/** Admin = all horses at location; Trainer/Groom = only HorseProfile assignments. */
function getUserHorseAssignmentMode_(user) {
  user = user || {};
  if (isAdminRole_(user.role)) {
    return { mode: "all", staffName: "", roleLabel: "Admin" };
  }
  if (isTrainerRole_(user.role)) {
    return { mode: "trainer", staffName: String(user.name || "").trim(), roleLabel: "Trainer" };
  }
  if (isGroomRole_(user.role)) {
    return { mode: "groom", staffName: String(user.name || "").trim(), roleLabel: "Groom" };
  }
  return { mode: "all", staffName: "", roleLabel: String(user.role || "") };
}

function horseAssignedToUser_(horseProf, assignMode) {
  if (!assignMode || assignMode.mode === "all") {
    return true;
  }
  if (assignMode.mode === "trainer") {
    return staffNameMatches_(horseProf.trainer, assignMode.staffName);
  }
  if (assignMode.mode === "groom") {
    return staffNameMatches_(horseProf.groom, assignMode.staffName);
  }
  return true;
}

function filterHorseProfilesForUser_(profiles, user) {
  var mode = getUserHorseAssignmentMode_(user);
  if (mode.mode === "all") {
    return profiles;
  }
  return (profiles || []).filter(function(h) {
    return horseAssignedToUser_(h, mode);
  });
}

function assertHorseAccessForUser_(horseName, location, user) {
  var profiles = filterHorseProfilesForUser_(getHorseProfilesForLocation_(location), user);
  var key = String(horseName || "").trim().toLowerCase();
  var i;
  for (i = 0; i < profiles.length; i++) {
    if (String(profiles[i].name || "").trim().toLowerCase() === key) {
      return profiles[i];
    }
  }
  var mode = getUserHorseAssignmentMode_(user);
  if (mode.mode === "trainer") {
    throw new Error('Horse "' + horseName + '" is not assigned to you as trainer.');
  }
  if (mode.mode === "groom") {
    throw new Error('Horse "' + horseName + '" is not assigned to you as groom.');
  }
  throw new Error('Horse "' + horseName + '" is not available at this location.');
}

function isAllLocations_(loc) {
  var s = String(loc || "").trim().toLowerCase();
  return !s || s === "all" || s === "all locations";
}

function resolveViewLocation_(user, requestedView) {
  if (isAdminRole_(user.role)) {
    if (isAllLocations_(requestedView)) {
      return "";
    }
    var resolved = resolveLocationName_(requestedView);
    if (!resolved && String(requestedView || "").trim()) {
      resolved = String(requestedView).trim();
    }
    return resolved;
  }
  return resolveLocationName_(user.location || user.locationId);
}

function assertLocationAccess_(user, requestedLocation) {
  if (isAdminRole_(user.role)) {
    return resolveLocationName_(requestedLocation);
  }
  var userLoc = resolveLocationName_(user.location || user.locationId);
  if (!userLoc) {
    throw new Error("Your account has no location assigned. Contact an administrator.");
  }
  if (requestedLocation && !matchesLocationFilter_(requestedLocation, userLoc)) {
    throw new Error("You can only access data for your assigned location.");
  }
  return userLoc;
}

function enforcePayloadLocation_(payload, token) {
  var user = validateSessionToken_(token);
  if (!isAdminRole_(user.role)) {
    var loc = resolveLocationName_(user.location || user.locationId);
    if (!loc) {
      throw new Error("Your account has no location assigned.");
    }
    payload.Location = loc;
  } else if (payload.Location) {
    payload.Location = resolveLocationName_(payload.Location);
  }
  return payload;
}

function getAppContext(token, viewLocation) {
  try {
    return getAppContext_(token, viewLocation);
  } catch (err) {
    Logger.log("getAppContext: " + err);
    throw new Error(err && err.message ? err.message : String(err));
  }
}

function getAppContext_(token, viewLocation) {
  var user = validateSessionToken_(token);
  Logger.log("[KingsOps getAppContext] session user: " + JSON.stringify({
    email: user.email,
    role: user.role,
    location: user.location,
    locationId: user.locationId
  }));
  var isAdmin = isAdminRole_(user.role);
  var userLocRec = resolveLocationKey_(user.location || user.locationId);
  var userLocation = userLocRec.name || "";
  var userLocationLabel = userLocRec.label || userLocation || "—";
  Logger.log("[KingsOps getAppContext] resolved location name=" + userLocation + " label=" + userLocationLabel);
  var effectiveView = resolveViewLocation_(user, viewLocation);
  var records = getLocationsRecords_();
  var adminOptions = [];
  if (isAdmin) {
    adminOptions.push({ value: "", label: "All Locations", id: "", name: "All Locations", manager: "", phone: "" });
    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      adminOptions.push({
        value: rec.name,
        id: rec.id,
        name: rec.name,
        label: rec.id ? rec.id + " · " + rec.name : rec.name,
        manager: rec.manager || "",
        phone: rec.phone || ""
      });
    }
  }
  var viewRec = resolveLocationKey_(effectiveView);
  var viewLabel = isAllLocations_(effectiveView)
    ? "All Locations"
    : viewRec.label || effectiveView || getDefaultAdminLocationName_() || "—";
  return {
    email: user.email,
    name: user.name,
    role: user.role,
    location: userLocation,
    locationId: userLocRec.id || "",
    locationLabel: userLocationLabel,
    isAdmin: isAdmin,
    locationLocked: !isAdmin,
    viewLocation: effectiveView,
    viewLocationId: viewRec.id || "",
    viewLabel: viewLabel,
    adminOptions: adminOptions,
    defaultAdminLocationId: getDefaultAdminLocationId_(),
    defaultAdminLocation: getDefaultAdminLocationName_(),
    locations: records
  };
}

/**
 * Debug helper — call from browser console via google.script.run or auto from AppLocation.
 * View server logs: Apps Script → Executions.
 */
function debugAppContext(token, viewLocation) {
  var out = {
    ok: false,
    tokenPresent: !!(token && String(token).trim()),
    viewLocationParam: viewLocation,
    sessionUser: null,
    contextSummary: null,
    locationsSheetCount: 0,
    error: ""
  };
  try {
    var user = validateSessionToken_(token);
    out.sessionUser = {
      email: user.email,
      name: user.name,
      role: user.role,
      location: user.location,
      locationId: user.locationId,
      isAdmin: isAdminRole_(user.role)
    };
    var records = getLocationsRecords_();
    out.locationsSheetCount = records.length;
    var ctx = getAppContext_(token, viewLocation);
    out.contextSummary = {
      role: ctx.role,
      location: ctx.location,
      locationId: ctx.locationId,
      locationLabel: ctx.locationLabel,
      isAdmin: ctx.isAdmin,
      viewLocation: ctx.viewLocation,
      viewLabel: ctx.viewLabel,
      defaultAdminLocation: ctx.defaultAdminLocation
    };
    out.ok = true;
  } catch (err) {
    out.error = err && err.message ? err.message : String(err);
    Logger.log("[KingsOps debugAppContext] error: " + out.error);
  }
  Logger.log("[KingsOps debugAppContext] " + JSON.stringify(out));
  return out;
}

function buildHorseLocationMap_(horses) {
  var map = {};
  for (var i = 0; i < horses.length; i++) {
    var name = String(horses[i].name || "").trim().toLowerCase();
    if (name) {
      map[name] = normalizeSheetLocation_(horses[i].location || "");
    }
  }
  return map;
}

function horseMatchesLocationFilter_(horseName, locationFilter, horseLocMap) {
  if (isAllLocations_(locationFilter) || !String(locationFilter || "").trim()) {
    return true;
  }
  var horse = String(horseName || "").trim().toLowerCase();
  return matchesLocationFilter_(horseLocMap[horse], locationFilter);
}

function filterItemsByLocation_(items, locationFilter, locKeys) {
  if (isAllLocations_(locationFilter) || !locationFilter) {
    return items;
  }
  locKeys = locKeys || ["location", "Location"];
  return items.filter(function(item) {
    for (var k = 0; k < locKeys.length; k++) {
      if (item[locKeys[k]] !== undefined && matchesLocationFilter_(item[locKeys[k]], locationFilter)) {
        return true;
      }
    }
    return false;
  });
}

function filterRowsByHorseLocation_(rows, locationFilter, horseLocMap, horseKeys) {
  if (isAllLocations_(locationFilter) || !locationFilter) {
    return rows;
  }
  horseKeys = horseKeys || ["Horse", "horse"];
  return rows.filter(function(row) {
    for (var k = 0; k < horseKeys.length; k++) {
      var horse = String(row[horseKeys[k]] || "").trim().toLowerCase();
      if (horse && horseMatchesLocationFilter_(horse, locationFilter, horseLocMap)) {
        return true;
      }
    }
    return false;
  });
}

function filterDashboardPayload_(payload, locationFilter) {
  if (isAllLocations_(locationFilter) || !locationFilter) {
    payload.viewLocation = "";
    payload.viewLocationLabel = "All Locations";
    Logger.log("[KingsOps filterDashboard] All locations — horses=" + (payload.horses || []).length);
    return payload;
  }
  Logger.log("[KingsOps filterDashboard] filter by: " + locationFilter);
  var locName = resolveLocationName_(locationFilter);
  var locRec = resolveLocationKey_(locationFilter);
  var horseListBefore = (payload.horses || []).slice();
  var horseLocMap = buildHorseLocationMap_(horseListBefore);

  var horsesBefore = horseListBefore.length;
  payload._unfilteredHorseCount = horsesBefore;
  payload._unfilteredTrainerCount = (payload.trainers || []).length;
  payload._unfilteredGroomCount = (payload.grooms || []).length;
  payload.horses = filterItemsByLocation_(horseListBefore, locName, ["location", "Location"]);
  if (horsesBefore > 0 && payload.horses.length === 0) {
    var locSamples = {};
    var hi;
    for (hi = 0; hi < horseListBefore.length; hi++) {
      var hl = String(horseListBefore[hi].location || "").trim() || "(empty)";
      locSamples[hl] = (locSamples[hl] || 0) + 1;
    }
    payload._locationSamples = locSamples;
    payload._locationFilterWarning =
      "No horses matched location \"" + locName + "\" (" + horsesBefore + " horses in sheet). " +
      "Check Horses → Location uses the same names as Locations tab, or pick All Locations in the header.";
    Logger.log("[KingsOps filterDashboard] WARNING: 0/" + horsesBefore + " for \"" + locName + "\" samples=" + JSON.stringify(locSamples));
  }
  payload.trainers = filterItemsByLocation_(payload.trainers || [], locName, ["location", "Location"]);
  payload.grooms = filterItemsByLocation_(payload.grooms || [], locName, ["location", "Location"]);
  payload.inventory = filterItemsByLocation_(payload.inventory || [], locName, ["Location", "location"]);
  payload.todayActivity = filterItemsByLocation_(payload.todayActivity || [], locName, ["Location", "location"]);
  payload.health = filterRowsByHorseLocation_(payload.health || [], locName, horseLocMap);
  payload.vaccinations = filterRowsByHorseLocation_(payload.vaccinations || [], locName, horseLocMap);
  payload.medicalHistory = filterRowsByHorseLocation_(payload.medicalHistory || [], locName, horseLocMap);
  payload.shoeing = filterRowsByHorseLocation_(payload.shoeing || [], locName, horseLocMap);
  if (payload._shoeingStatus) {
    payload._shoeingStatus = payload._shoeingStatus.filter(function(s) {
      return horseMatchesLocationFilter_(s.horse, locName, horseLocMap);
    });
  }
  if (payload._vaccineStatus) {
    payload._vaccineStatus = payload._vaccineStatus.filter(function(v) {
      return horseMatchesLocationFilter_(v.horse, locName, horseLocMap);
    });
  }

  var alerts = [];
  for (var a = 0; a < (payload.alerts || []).length; a++) {
    var alert = payload.alerts[a];
    var title = String(alert.title || "");
    if (alert.type === "stock") {
      var atIdx = title.indexOf(" @ ");
      if (atIdx >= 0) {
        var stockLoc = title.substring(atIdx + 3).trim();
        if (matchesLocationFilter_(stockLoc, locName)) {
          alerts.push(alert);
        }
      }
    } else {
      var horseName = title.split(" — ")[0].trim().toLowerCase();
      if (horseMatchesLocationFilter_(horseName, locName, horseLocMap)) {
        alerts.push(alert);
      }
    }
  }
  payload.alerts = alerts;

  var stockAlerts = 0;
  for (var i = 0; i < payload.inventory.length; i++) {
    var st = String(payload.inventory[i]["Entry Type"] || payload.inventory[i].status || "");
    if (st.indexOf("Low") >= 0 || st.indexOf("High") >= 0 || st.indexOf("Check") >= 0) {
      stockAlerts++;
    }
  }

  var healthSummaryRows = [];
  var healthSource = payload._healthSummary;
  if (!healthSource || !healthSource.length) {
    try {
      healthSource = getHealthSummary();
    } catch (eHealthSrc) {
      healthSource = [];
    }
  }
  healthSummaryRows = healthSource.filter(function(h) {
    var horse = String(h.horse || "").trim().toLowerCase();
    return horseMatchesLocationFilter_(horse, locName, horseLocMap);
  });
  var healthyN = countHealthStatus_(healthSummaryRows, "🟢 Healthy");
  var watchN = countHealthStatus_(healthSummaryRows, "🟡 Watch");
  var vetN = countHealthStatus_(healthSummaryRows, "🔴 Attention");

  var week = getWeekActivityStats_(locName);
  payload.reports = payload.reports || {};
  payload.reports.sessionsThisWeek = week.sessions;
  payload.reports.avgDuration = week.avgDuration;
  payload.reports.periodLabel = week.periodLabel;
  payload.reports.activityMix = week.activityMix;
  payload.reports.healthSummary = { healthy: healthyN, watch: watchN, vetCare: vetN };

  var invConsumption = {};
  for (var c = 0; c < payload.inventory.length; c++) {
    var inv = payload.inventory[c];
    var it = String(inv.Item || inv.item || "Item").trim() || "Item";
    invConsumption[it] = (invConsumption[it] || 0) + (parseFloat(inv.Consumption) || 0);
  }
  payload.reports.inventoryConsumption = invConsumption;

  payload.stats = payload.stats || {};
  payload.stats.totalHorses = payload.horses.length;
  payload.stats.staff = payload.trainers.length + payload.grooms.length;
  payload.stats.stockAlerts = stockAlerts;
  payload.stats.locationCount = 1;
  payload.stats.healthy = healthyN;
  payload.viewLocation = locName;
  payload.viewLocationLabel = locRec.label || locName;
  payload.locations = [locName];
  var invForStock = payload._dailyInventoryRows || [];
  if (locName) {
    invForStock = filterItemsByLocation_(invForStock, locName, ["location", "Location"]);
  }
  payload.widgets = getDashboardWidgets_(payload.horses, payload.trainers, payload.grooms, locName, {
    shoeing: payload._shoeingStatus || [],
    vaccines: payload._vaccineStatus || [],
    stockDays: getStockDaysRemainingFast_(invForStock, ""),
    skipProfiles: true
  });
  payload.stats.rehabilitation = payload.widgets.horses.rehabilitation;
  payload.stats.healthyCount = payload.widgets.horses.healthy;
  payload.stats.watch = payload.widgets.horses.watch;
  payload.stats.trainersActive = payload.widgets.trainers.active;
  payload.stats.trainersLeave = payload.widgets.trainers.leave;
  payload.stats.groomsActive = payload.widgets.grooms.active;
  payload.stats.groomsLeave = payload.widgets.grooms.leave;

  Logger.log("[KingsOps filterDashboard] filtered — horses=" + payload.horses.length +
    " trainers=" + payload.trainers.length + " alerts=" + payload.alerts.length);
  return payload;
}

function logoutUser(token) {
  token = (token || "").toString().trim();
  if (token) {
    CacheService.getScriptCache().remove(SESSION_PREFIX + token);
    try {
      PropertiesService.getScriptProperties().deleteProperty(SESSION_PREFIX + token);
    } catch (e) {}
  }
  return { ok: true };
}

/** Run in Apps Script editor (no token) to verify sheet reads. View → Execution log. */
function testDashboardSheetReads() {
  var ss = getSS_();
  Logger.log("Spreadsheet: " + ss.getName() + " (" + ss.getId() + ")");
  Logger.log("Horses: " + getHorsesList_().length);
  Logger.log("Locations: " + getLocationsList_().length);
  Logger.log("Trainers: " + getTrainersList_().length);
  Logger.log("Grooms: " + getGroomsList_().length);
  Logger.log("Health rows: " + getHealthSummary().length);
  Logger.log("Daily inventory rows: " + getDailyInventoryRows_().length);
  return getDashboardCounts();
}

function getTodayActivityRows_() {
  return readActivityBundle_().todayActivity;
}

function getWeekActivityStats_(locationFilter) {
  var sessions = 0;
  var totalMins = 0;
  var mix = {};
  var cached = getActivitySheetData_();
  var data = cached.data;
  var map = cached.map;
  if (!data || data.length < 2) {
    return { sessions: 0, avgDuration: 0, activityMix: mix, periodLabel: "This Week" };
  }
  var iDate = findCol_(map, ["Date", "date"]);
  var iActivity = findCol_(map, ["Activity", "activity"]);
  var iDuration = findCol_(map, ["Duration", "duration"]);
  var iLoc = findCol_(map, ["Location", "location"]);
  var today = new Date();
  var seven = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  var useAllTime = false;
  var filterLoc = isAllLocations_(locationFilter) ? "" : String(locationFilter || "").trim();

  function countRow(row) {
    if (filterLoc && iLoc >= 0 && !matchesLocationFilter_(row[iLoc], filterLoc)) {
      return;
    }
    sessions++;
    totalMins += parseRoutineDurationToMinutes_(row[iDuration]);
    var act = (row[iActivity] || "").toString() || "Unknown";
    mix[act] = (mix[act] || 0) + 1;
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dt = new Date(row[iDate]);
    if (isNaN(dt.getTime()) || dt < seven) {
      continue;
    }
    countRow(row);
  }

  if (sessions === 0) {
    useAllTime = true;
    for (var j = 1; j < data.length; j++) {
      var r = data[j];
      var d = new Date(r[iDate]);
      if (isNaN(d.getTime())) {
        continue;
      }
      countRow(r);
    }
  }

  return {
    sessions: sessions,
    avgDuration: sessions ? Math.round(totalMins / sessions) : 0,
    activityMix: mix,
    periodLabel: useAllTime ? "All Time" : "This Week"
  };
}

function summarizeDashboardPayload_(payload) {
  payload = payload || {};
  return {
    viewLocation: payload.viewLocation || "",
    viewLocationLabel: payload.viewLocationLabel || "",
    horses: (payload.horses || []).length,
    trainers: (payload.trainers || []).length,
    grooms: (payload.grooms || []).length,
    alerts: (payload.alerts || []).length,
    todayActivity: (payload.todayActivity || []).length,
    inventory: (payload.inventory || []).length,
    health: (payload.health || []).length,
    stats: payload.stats || {}
  };
}

function clearDashRequestCache_() {
  __dashActivityCache_ = null;
  __locationsRecordsCache_ = null;
  __locationResolveCache_ = {};
  __sheetHeaderCache_ = {};
  __inventoryMinMapCache_ = null;
}

function getSheetDataBounded_(sheet, maxDataRows) {
  if (!sheet) {
    return { header: [], rows: [], map: {}, rowOffset: 0 };
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { header: [], rows: [], map: {}, rowOffset: 0 };
  }
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  var c;
  for (c = 0; c < header.length; c++) {
    var key = normalizeHeaderKey_(header[c]);
    if (key) {
      map[key] = c;
    }
  }
  var startRow = 2;
  if (maxDataRows > 0 && lastRow - 1 > maxDataRows) {
    startRow = lastRow - maxDataRows + 1;
  }
  var rows = startRow <= lastRow
    ? sheet.getRange(startRow, 1, lastRow, lastCol).getValues()
    : [];
  return { header: header, rows: rows, map: map, rowOffset: startRow };
}

function getActivitySheetData_() {
  if (__dashActivityCache_) {
    return __dashActivityCache_;
  }
  var sheet = getSS_().getSheetByName("ActivityData");
  if (!sheet) {
    __dashActivityCache_ = { data: [], map: {}, bounded: true };
    return __dashActivityCache_;
  }
  var bounded = getSheetDataBounded_(sheet, DASH_SHEET_ROW_LIMITS_.activity);
  var data = [bounded.header].concat(bounded.rows);
  __dashActivityCache_ = { data: data, map: bounded.map, bounded: true };
  return __dashActivityCache_;
}

function buildHorseHealthStatusMap_(healthRows) {
  var m = {};
  var i;
  for (i = 0; i < (healthRows || []).length; i++) {
    m[healthRows[i].horse] = healthRows[i].status;
  }
  return m;
}

function readActivityBundle_() {
  var cached = getActivitySheetData_();
  var data = cached.data;
  var map = cached.map;
  if (!data || data.length < 2) {
    return {
      todayActivity: [],
      week: { sessions: 0, avgDuration: 0, activityMix: {}, periodLabel: "This Week" }
    };
  }
  var iDate = findCol_(map, ["Date", "date"]);
  var iLoc = findCol_(map, ["Location", "location"]);
  var iHorse = findCol_(map, ["Horse", "Horse_Name"]);
  var iActivity = findCol_(map, ["Activity", "activity"]);
  var iDuration = findCol_(map, ["Duration", "duration"]);
  var iDurationLabel = findCol_(map, ["Duration_Label", "Duration Label"]);
  var iTrainer = findCol_(map, ["Trainer", "Trainer_Name"]);
  var iGroom = findCol_(map, ["Groom", "Groom_Name"]);
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var seven = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  var todayRows = [];
  var latestDate = null;
  var ri;
  for (ri = 1; ri < data.length; ri++) {
    var row = data[ri];
    var dt = new Date(row[iDate]);
    if (isNaN(dt.getTime())) {
      continue;
    }
    var d0 = new Date(dt);
    d0.setHours(0, 0, 0, 0);
    if (d0.getTime() === today.getTime()) {
      todayRows.push(row);
    }
    if (!latestDate || d0 > latestDate) {
      latestDate = d0;
    }
  }
  var targetDate = today;
  var sourceRows = todayRows;
  if (!todayRows.length && latestDate) {
    targetDate = latestDate;
    sourceRows = [];
    for (ri = 1; ri < data.length; ri++) {
      var r2 = data[ri];
      var rd = new Date(r2[iDate]);
      if (isNaN(rd.getTime())) {
        continue;
      }
      var rd0 = new Date(rd);
      rd0.setHours(0, 0, 0, 0);
      if (rd0.getTime() === latestDate.getTime()) {
        sourceRows.push(r2);
      }
    }
  }
  var actLocCache = {};
  function normActLoc(raw) {
    var s = String(raw || "").trim();
    if (!s) {
      return "";
    }
    var k = s.toLowerCase();
    if (!actLocCache[k]) {
      actLocCache[k] = normalizeSheetLocation_(s);
    }
    return actLocCache[k];
  }
  var todayActivity = [];
  for (ri = 0; ri < sourceRows.length; ri++) {
    var sr = sourceRows[ri];
    var durMins = parseRoutineDurationToMinutes_(sr[iDuration]);
    var durLabel = iDurationLabel >= 0 ? String(sr[iDurationLabel] || "").trim() : "";
    if (!durLabel) {
      durLabel = formatRoutineDurationLabel_(durMins);
    }
    todayActivity.push({
      Horse: sr[iHorse],
      Activity: sr[iActivity],
      Duration: durMins,
      DurationLabel: durLabel,
      Trainer: sr[iTrainer],
      Groom: iGroom >= 0 ? sr[iGroom] : "",
      Location: iLoc >= 0 ? normActLoc(sr[iLoc]) : "",
      _date: targetDate ? targetDate.toISOString() : ""
    });
  }
  var sessions = 0;
  var totalMins = 0;
  var mix = {};
  var useAllTime = false;
  function countWeekRow(wrow) {
    sessions++;
    totalMins += parseRoutineDurationToMinutes_(wrow[iDuration]);
    var act = (wrow[iActivity] || "").toString() || "Unknown";
    mix[act] = (mix[act] || 0) + 1;
  }
  for (ri = 1; ri < data.length; ri++) {
    var wrow = data[ri];
    var wdt = new Date(wrow[iDate]);
    if (isNaN(wdt.getTime()) || wdt < seven) {
      continue;
    }
    countWeekRow(wrow);
  }
  if (sessions === 0) {
    useAllTime = true;
    for (ri = 1; ri < data.length; ri++) {
      var ar = data[ri];
      var ad = new Date(ar[iDate]);
      if (isNaN(ad.getTime())) {
        continue;
      }
      countWeekRow(ar);
    }
  }
  return {
    todayActivity: todayActivity,
    week: {
      sessions: sessions,
      avgDuration: sessions ? Math.round(totalMins / sessions) : 0,
      activityMix: mix,
      periodLabel: useAllTime ? "All Time" : "This Week"
    }
  };
}

/** Normalize Daily Inventory row (internal or UI-shaped) for stock-days widget. */
function normalizeInventoryRowForStock_(row) {
  row = row || {};
  var loc = row.location !== undefined && row.location !== null && row.location !== ""
    ? row.location
    : row.Location;
  var closing = row.closing !== undefined && row.closing !== null && row.closing !== ""
    ? row.closing
    : (row.Closing !== undefined && row.Closing !== null && row.Closing !== ""
      ? row.Closing
      : row.Quantity);
  var consumption = row.consumption !== undefined && row.consumption !== null && row.consumption !== ""
    ? row.consumption
    : row.Consumption;
  return {
    item: String(row.item || row.Item || "").trim(),
    location: loc !== undefined && loc !== null ? String(loc).trim() : "",
    closing: parseFloat(closing) || 0,
    consumption: parseFloat(consumption) || 0
  };
}

function getStockDaysRemainingFast_(invRows, locationFilter) {
  invRows = invRows || [];
  var out = [];
  var seen = {};
  var i;
  for (i = 0; i < invRows.length; i++) {
    var norm = normalizeInventoryRowForStock_(invRows[i]);
    if (locationFilter && norm.location && !matchesLocationFilter_(norm.location, locationFilter)) {
      continue;
    }
    if (!norm.item || seen[norm.item + "|" + norm.location]) {
      continue;
    }
    seen[norm.item + "|" + norm.location] = true;
    var daysLeft = norm.consumption > 0 ? Math.floor(norm.closing / norm.consumption) : null;
    out.push({
      item: norm.item,
      location: norm.location,
      stock: norm.closing,
      dailyUse: Math.round(norm.consumption * 100) / 100,
      daysLeft: daysLeft,
      label: daysLeft !== null ? daysLeft + " days stock" : "—"
    });
  }
  out.sort(function(a, b) {
    var da = a.daysLeft === null ? 9999 : a.daysLeft;
    var db = b.daysLeft === null ? 9999 : b.daysLeft;
    return da - db;
  });
  return out.slice(0, 8);
}

function getDashboardData(token, viewLocation) {
  clearDashRequestCache_();
  var t0 = new Date().getTime();
  var user = validateSessionToken_(token);
  var locFilter = resolveViewLocation_(user, viewLocation);
  Logger.log("[KingsOps getDashboardData] START viewParam=" + JSON.stringify(viewLocation) +
    " locFilter=" + JSON.stringify(locFilter));

  try {
    var raw = buildDashboardPayload_();
    Logger.log("[KingsOps getDashboardData] buildDashboardPayload ms=" + (new Date().getTime() - t0) +
      " horses=" + (raw.horses || []).length);
    var payload = filterDashboardPayload_(raw, locFilter);
    delete payload._healthSummary;
    delete payload._shoeingStatus;
    delete payload._vaccineStatus;
    delete payload._dailyInventoryRows;
    Logger.log("[KingsOps getDashboardData] DONE ms=" + (new Date().getTime() - t0) +
      " " + JSON.stringify(summarizeDashboardPayload_(payload)) +
      (payload._locationFilterWarning ? " WARN:" + payload._locationFilterWarning : ""));
    return payload;
  } catch (err) {
    Logger.log("getDashboardData error: " + err);
    var msg = String(err && err.message ? err.message : err);
    var payload = emptyDashboardPayload_(msg);
    try {
      payload.horses = getHorsesList_({});
      payload.trainers = getTrainersList_();
      payload.grooms = getGroomsList_();
      payload.stats.totalHorses = payload.horses.length;
      payload.stats.staff = payload.trainers.length + payload.grooms.length;
    } catch (e2) {
      Logger.log("getDashboardData partial: " + e2);
    }
    if (locFilter) {
      payload = filterDashboardPayload_(payload, locFilter);
    }
    Logger.log("[KingsOps getDashboardData] fallback OK " + JSON.stringify(summarizeDashboardPayload_(payload)));
    return payload;
  }
}

function countHealthStatus_(rows, status) {
  var n = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].status === status) {
      n++;
    }
  }
  return n;
}

function emptyDashboardPayload_(message) {
  return {
    stats: { totalHorses: 0, healthy: 0, staff: 0, stockAlerts: 0, locationCount: 0 },
    horses: [],
    trainers: [],
    grooms: [],
    locations: [],
    alerts: message ? [{ type: "stock", title: "Dashboard data error", sub: message }] : [],
    _loadError: message || "",
    todayActivity: [],
    inventory: [],
    health: [],
    vaccinations: [],
    medicalHistory: [],
    shoeing: [],
    reports: {
      sessionsThisWeek: 0,
      avgDuration: 0,
      activityMix: {},
      healthSummary: { healthy: 0, watch: 0, vetCare: 0 },
      inventoryConsumption: {}
    }
  };
}

function buildDashboardPayload_() {
  var tBuild = new Date().getTime();
  var horses = [];
  var trainers = [];
  var grooms = [];
  var locations = [];
  var healthRows = [];
  getLocationsRecords_();
  try {
    healthRows = getHealthSummary(true);
    Logger.log("[KingsOps build] health ms=" + (new Date().getTime() - tBuild));
  } catch (eHealth) {
    Logger.log("getHealthSummary: " + eHealth);
  }
  try {
    horses = getHorsesList_(buildHorseHealthStatusMap_(healthRows));
    Logger.log("[KingsOps build] horses ms=" + (new Date().getTime() - tBuild));
  } catch (e) {
    Logger.log("getHorsesList_: " + e);
  }
  try {
    trainers = getTrainersList_();
    grooms = getGroomsList_();
    Logger.log("[KingsOps build] staff ms=" + (new Date().getTime() - tBuild));
  } catch (e) {
    Logger.log("getTrainersList_/getGroomsList_: " + e);
  }
  try {
    locations = getLocationsList_();
  } catch (e) {
    Logger.log("getLocationsList_: " + e);
  }
  var staffCount = trainers.length + grooms.length;

  var invRows = [];
  try {
    invRows = getDailyInventoryRows_(DASH_SHEET_ROW_LIMITS_.inventory);
    Logger.log("[KingsOps build] inventory rows=" + invRows.length + " ms=" + (new Date().getTime() - tBuild));
  } catch (e) {
    invRows = [];
  }
  var minMap = getInventoryMinMapCached_();
  var stockAlerts = 0;
  for (var ir = 0; ir < invRows.length; ir++) {
    var st = (invRows[ir].status || "").toString();
    var itemName = String(invRows[ir].item || "").trim();
    var closingQty = parseFloat(invRows[ir].closing);
    var minLevel = minMap[itemName];
    if (st.indexOf("High") >= 0 || st.indexOf("Check") >= 0) {
      stockAlerts++;
    } else if (minLevel && !isNaN(closingQty) && closingQty < minLevel) {
      stockAlerts++;
      invRows[ir].status = "🔴 Low stock";
    }
  }

  var inventoryList = [];
  for (var j = 0; j < invRows.length && j < 25; j++) {
    var r = invRows[j];
    var itemVal = r.item !== undefined ? r.item : r.Item;
    var locVal = r.location !== undefined ? r.location : r.Location;
    var closeVal = r.closing !== undefined ? r.closing : r.Closing;
    inventoryList.push({
      Item: itemVal,
      Location: locVal,
      Quantity: closeVal,
      Opening: r.opening !== undefined ? r.opening : r.Opening,
      Added: r.added !== undefined ? r.added : r.Added,
      Consumption: r.consumption !== undefined ? r.consumption : r.Consumption,
      "Entry Type": r.status || "Summary"
    });
  }

  var healthUi = [];
  for (var hi = 0; hi < healthRows.length && hi < 40; hi++) {
    var h = healthRows[hi];
    healthUi.push({
      Horse: h.horse,
      Date: h.lastCheck,
      Location: h.location || "",
      Status: h.status || "",
      Temp: h.temp,
      Temperature_C: h.temperatureC || h.temp,
      Heart_Rate: h.heartRate || "",
      Respiratory_Rate: h.respiratoryRate || "",
      Appetite: h.appetite,
      Water: h.water,
      Drinking: h.water,
      Dung: h.dung,
      Droppings_Consistency: h.dung,
      Injury: h.injury,
      Medication: h.medication,
      Demeanour: h.demeanour || "",
      Movement: h.movement || "",
      Eyes: h.eyes || "",
      Body_Condition: h.bodyCondition || "",
      Vet_Check_Date: h.vetCheckDate || "",
      Dental_Check_Date: h.dentalCheckDate || "",
      Checked_By: h.checkedBy || "",
      Notes: h.notes
    });
  }

  var shoeing = getShoeingStatus(true);
  Logger.log("[KingsOps build] shoeing ms=" + (new Date().getTime() - tBuild));
  var shoeingUi = [];
  for (var si = 0; si < shoeing.length; si++) {
    var s = shoeing[si];
    shoeingUi.push({
      Horse: s.horse,
      "Shoeing Type": s.type,
      Farrier: s.farrier,
      Date: s.lastDate,
      "Next Due": s.nextDue
    });
  }

  var vaccines = getVaccinationStatus(true);
  Logger.log("[KingsOps build] vaccines ms=" + (new Date().getTime() - tBuild));
  var vaccUi = [];
  for (var vi = 0; vi < vaccines.length; vi++) {
    var v = vaccines[vi];
    vaccUi.push({
      Vaccine: v.vaccine,
      Horse: v.horse,
      "Next Due": v.nextDue,
      Vet: v.vet
    });
  }

  var med = getMedicalHistory(25, true);
  Logger.log("[KingsOps build] medical ms=" + (new Date().getTime() - tBuild));
  var medUi = [];
  for (var mi = 0; mi < med.length; mi++) {
    var m = med[mi];
    medUi.push({
      Horse: m.horse,
      Issue: m.issue,
      Date: m.date,
      Vet: m.vet,
      Status: m.status
    });
  }

  var alerts = [];
  shoeing
    .filter(function(x) {
      return x.alert !== "✅ OK";
    })
    .slice(0, 5)
    .forEach(function(x) {
      alerts.push({
        type: "shoeing",
        title: x.horse + " — shoeing",
        sub: x.alert + " · next " + x.nextDue
      });
    });
  vaccines
    .filter(function(x) {
      return x.alert !== "✅ OK";
    })
    .slice(0, 5)
    .forEach(function(x) {
      alerts.push({
        type: "vaccination",
        title: x.horse + " — " + x.vaccine,
        sub: x.alert + " · due " + x.nextDue
      });
    });
  healthRows
    .filter(function(h) {
      return h.status && h.status !== "🟢 Healthy";
    })
    .slice(0, 5)
    .forEach(function(h) {
      alerts.push({
        type: "health",
        title: h.horse + " — health check",
        sub: (h.status || "Review") + " · " + (h.lastCheck || "")
      });
    });
  invRows
    .filter(function(row) {
      var st = (row.status || "").toString();
      return st.indexOf("High") >= 0 || st.indexOf("Check") >= 0;
    })
    .slice(0, 3)
    .forEach(function(row) {
      alerts.push({
        type: "stock",
        title: row.item + " @ " + row.location,
        sub: row.status || "Review stock"
      });
    });

  var activityBundle = readActivityBundle_();
  Logger.log("[KingsOps build] activity ms=" + (new Date().getTime() - tBuild));
  var week = activityBundle.week;
  var healthyN = healthRows.filter(function(h) {
    return h.status === "🟢 Healthy";
  }).length;
  var watchN = healthRows.filter(function(h) {
    return h.status === "🟡 Watch";
  }).length;
  var vetN = healthRows.filter(function(h) {
    return h.status === "🔴 Attention";
  }).length;

  var invConsumption = {};
  for (var ci = 0; ci < inventoryList.length; ci++) {
    var it = (inventoryList[ci].Item || "").toString() || "Item";
    invConsumption[it] = (invConsumption[it] || 0) + (parseFloat(inventoryList[ci].Consumption) || 0);
  }

  var stockDaysFast = getStockDaysRemainingFast_(invRows, "");
  var widgets = getDashboardWidgets_(horses, trainers, grooms, "", {
    shoeing: shoeing,
    vaccines: vaccines,
    stockDays: stockDaysFast,
    skipProfiles: true
  });

  return {
    viewLocation: "",
    viewLocationLabel: "All Locations",
    _healthSummary: healthRows,
    _shoeingStatus: shoeing,
    _vaccineStatus: vaccines,
    stats: {
      totalHorses: horses.length,
      healthy: healthyN,
      healthyCount: widgets.horses.healthy,
      rehabilitation: widgets.horses.rehabilitation,
      watch: widgets.horses.watch,
      staff: staffCount,
      stockAlerts: stockAlerts,
      locationCount: locations.length,
      trainersActive: widgets.trainers.active,
      trainersLeave: widgets.trainers.leave,
      groomsActive: widgets.grooms.active,
      groomsLeave: widgets.grooms.leave
    },
    widgets: widgets,
    locations: locations,
    horses: horses,
    trainers: trainers,
    grooms: grooms,
    alerts: alerts,
    todayActivity: activityBundle.todayActivity,
    _dailyInventoryRows: invRows,
    inventory: inventoryList,
    health: healthUi,
    vaccinations: vaccUi,
    medicalHistory: medUi,
    shoeing: shoeingUi,
    reports: {
      sessionsThisWeek: week.sessions || 0,
      avgDuration: week.avgDuration || 0,
      periodLabel: week.periodLabel || "This Week",
      activityMix: week.activityMix || {},
      healthSummary: {
        healthy: healthyN,
        watch: watchN,
        vetCare: vetN
      },
      inventoryConsumption: invConsumption
    }
  };
  Logger.log("[KingsOps build] TOTAL ms=" + (new Date().getTime() - tBuild) + " horses=" + horses.length);
}

function getInventoryMinMapCached_() {
  if (__inventoryMinMapCache_) {
    return __inventoryMinMapCache_;
  }
  __inventoryMinMapCache_ = getInventoryMinMap_();
  return __inventoryMinMapCache_;
}

function getUser() {
  var email = Session.getActiveUser().getEmail();
  var data = getSS_().getSheetByName("Users").getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase()) {
      return { email: data[i][0], name: data[i][1], role: data[i][2], location: data[i][3] };
    }
  }
  return { role: "NO_ACCESS" };
}

// ─── GET ALL HORSES ───────────────────────────────────────────
function getHorses() {
  return getHorsesList_();
}

// ─── INVENTORY CALCULATION ────────────────────────────────────
var DAILY_INVENTORY_HEADERS_ = ["Date", "Location", "Item", "Opening", "Added", "Closing", "Consumption"];

function normalizeInventoryDateKey_(dateVal) {
  if (dateVal === null || dateVal === undefined || dateVal === "") {
    return "";
  }
  var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) {
    return String(dateVal).trim();
  }
  return Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd");
}

function inventoryAggregateKey_(dateKey, loc, item) {
  return dateKey + "|" + String(loc || "").trim().toLowerCase() + "|" + String(item || "").trim().toLowerCase();
}

function newInventoryAggregate_(dateVal, loc, item) {
  return { date: dateVal, loc: loc, item: item, open: 0, add: 0, close: 0, consumed: 0 };
}

function accumulateRawInventoryRow_(agg, entryType, qty) {
  var typeLower = String(entryType || "").trim().toLowerCase();
  if (typeLower === "opening") {
    agg.open += qty;
  } else if (typeLower === "added") {
    agg.add += qty;
  } else if (typeLower === "closing") {
    agg.close += qty;
  } else if (typeLower.indexOf("consumption") >= 0) {
    agg.consumed += qty;
  }
}

function inventoryRowValuesFromAggregate_(m) {
  var consumed = m.consumed || 0;
  var opening = m.open || 0;
  var added = m.add || 0;
  var manualClose = m.close || 0;
  var closing;
  var consumption;
  if (manualClose > 0) {
    closing = manualClose;
    consumption = opening + added - closing;
  } else if (consumed > 0) {
    consumption = consumed;
    closing = Math.max(0, opening + added - consumed);
  } else {
    closing = manualClose;
    consumption = opening + added - closing;
  }
  return [m.date, m.loc, m.item, opening, added, closing, consumption];
}

/**
 * Scan Raw Data once. If keyFilter is set, only aggregates those date|location|item keys.
 * @param {Object|null} keyFilter map of inventoryAggregateKey_ -> aggregate stub
 */
function scanRawInventoryAggregates_(rawSheet, keyFilter) {
  var raw = rawSheet.getDataRange().getValues();
  if (raw.length < 2) {
    return {};
  }
  var map = getHeaderIndexMap_(rawSheet);
  var cDate = findCol_(map, ["Date", "date"]);
  var cLoc = findCol_(map, ["Location", "location"]);
  var cItem = findCol_(map, ["Item", "item"]);
  var cType = findCol_(map, ["Entry Type", "Entry_Type", "entry_type"]);
  var cQty = findCol_(map, ["Quantity", "quantity"]);
  var grouped = keyFilter || {};
  var filterMode = !!keyFilter;
  var i;
  for (i = 1; i < raw.length; i++) {
    var row = raw[i];
    var loc = cLoc >= 0 ? normalizeSheetLocation_(row[cLoc]) : "";
    var item = cItem >= 0 ? String(row[cItem] || "").trim() : "";
    if (!loc || !item) {
      continue;
    }
    var dateKey = normalizeInventoryDateKey_(row[cDate]);
    var key = inventoryAggregateKey_(dateKey, loc, item);
    if (filterMode && !keyFilter[key]) {
      continue;
    }
    if (!grouped[key]) {
      grouped[key] = newInventoryAggregate_(row[cDate], loc, item);
    }
    accumulateRawInventoryRow_(grouped[key], row[cType], parseFloat(row[cQty]) || 0);
  }
  return grouped;
}

function ensureDailyInventorySheet_(outSheet) {
  if (outSheet.getLastRow() < 1) {
    outSheet.getRange(1, 1, 1, DAILY_INVENTORY_HEADERS_.length).setValues([DAILY_INVENTORY_HEADERS_]);
  }
}

/** Update or append only the given aggregates on Daily Inventory (fast path). */
function upsertDailyInventoryAggregates_(outSheet, grouped) {
  ensureDailyInventorySheet_(outSheet);
  var lastRow = outSheet.getLastRow();
  var lastCol = Math.max(outSheet.getLastColumn(), DAILY_INVENTORY_HEADERS_.length);
  var data = lastRow >= 1 ? outSheet.getRange(1, 1, lastRow, lastCol).getValues() : [DAILY_INVENTORY_HEADERS_.slice()];
  if (!data.length) {
    data = [DAILY_INVENTORY_HEADERS_.slice()];
  }
  var hdrMap = {};
  var hc;
  for (hc = 0; hc < data[0].length; hc++) {
    var hk = normalizeHeaderKey_(data[0][hc]);
    if (hk) {
      hdrMap[hk] = hc;
    }
  }
  var cDate = hdrMap.date !== undefined ? hdrMap.date : 0;
  var cLoc = hdrMap.location !== undefined ? hdrMap.location : 1;
  var cItem = hdrMap.item !== undefined ? hdrMap.item : 2;
  var rowIndex = {};
  var ri;
  for (ri = 1; ri < data.length; ri++) {
    var idxKey = inventoryAggregateKey_(
      normalizeInventoryDateKey_(data[ri][cDate]),
      data[ri][cLoc],
      data[ri][cItem]
    );
    rowIndex[idxKey] = ri + 1;
  }
  var updates = [];
  var appends = [];
  var k;
  for (k in grouped) {
    if (!grouped.hasOwnProperty(k)) {
      continue;
    }
    var vals = inventoryRowValuesFromAggregate_(grouped[k]);
    if (rowIndex[k]) {
      updates.push({ row: rowIndex[k], values: vals });
    } else {
      appends.push(vals);
    }
  }
  var ui;
  for (ui = 0; ui < updates.length; ui++) {
    outSheet.getRange(updates[ui].row, 1, 1, updates[ui].values.length).setValues([updates[ui].values]);
  }
  if (appends.length) {
    var startRow = outSheet.getLastRow() + 1;
    outSheet.getRange(startRow, 1, startRow + appends.length - 1, appends[0].length).setValues(appends);
  }
  return { updated: updates.length, appended: appends.length };
}

/**
 * After routine save: one Raw Data scan for consumed items only, upsert Daily Inventory rows.
 */
function updateDailyInventoryForConsumption_(location, consumptionByItem, logDate) {
  var rawSheet = getSheetByNameSafe_(["Raw Data", "RawData"]);
  var outSheet = getSheetByNameSafe_(["Daily Inventory", "Daily Inventory Summary"]);
  if (!rawSheet || !outSheet) {
    Logger.log("updateDailyInventoryForConsumption_: missing sheets");
    return { updated: 0, appended: 0, skipped: true };
  }
  var locName = resolveLocationName_(location);
  var dateKey = normalizeInventoryDateKey_(logDate || new Date());
  var entryDate = logDate ? new Date(logDate) : new Date();
  if (isNaN(entryDate.getTime())) {
    entryDate = new Date();
  }
  var keyFilter = {};
  var itemName;
  for (itemName in consumptionByItem) {
    if (!consumptionByItem.hasOwnProperty(itemName)) {
      continue;
    }
    if ((parseFloat(consumptionByItem[itemName]) || 0) <= 0) {
      continue;
    }
    var fk = inventoryAggregateKey_(dateKey, locName, itemName);
    keyFilter[fk] = newInventoryAggregate_(entryDate, locName, itemName);
  }
  if (!Object.keys(keyFilter).length) {
    return { updated: 0, appended: 0 };
  }
  var grouped = scanRawInventoryAggregates_(rawSheet, keyFilter);
  var result = upsertDailyInventoryAggregates_(outSheet, grouped);
  Logger.log("[KingsOps] incremental inventory " + JSON.stringify(result));
  return result;
}

/** Full rebuild of Daily Inventory from all Raw Data (slow — use for triggers / manual repair). */
function generateInventory() {
  var rawSheet = getSheetByNameSafe_(["Raw Data", "RawData"]);
  var outSheet = getSheetByNameSafe_(["Daily Inventory", "Daily Inventory Summary"]);
  if (!rawSheet) {
    throw new Error("Raw Data sheet not found.");
  }
  if (!outSheet) {
    throw new Error("Daily Inventory sheet not found.");
  }
  var grouped = scanRawInventoryAggregates_(rawSheet, null);
  var result = [DAILY_INVENTORY_HEADERS_];
  var k;
  for (k in grouped) {
    if (grouped.hasOwnProperty(k)) {
      result.push(inventoryRowValuesFromAggregate_(grouped[k]));
    }
  }
  outSheet.clearContents();
  if (result.length > 1) {
    outSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
  } else {
    ensureDailyInventorySheet_(outSheet);
  }
  return { success: true, rows: result.length - 1 };
}

// ─── GET INVENTORY SUMMARY ───────────────────────────────────
/** Rebuilds Daily Inventory from Raw Data — slow; do not call from getDashboardData. */
function getInventorySummary(runGenerate) {
  if (runGenerate !== false) {
    try {
      generateInventory();
    } catch (e) {
      Logger.log("generateInventory: " + e);
    }
  }
  return getDailyInventoryRows_();
}

// ─── SHOEING STATUS ───────────────────────────────────────────
function getShoeingStatus(forDashboard) {
  var sheet = getSS_().getSheetByName("ShoeingData");
  if (!sheet) {
    return [];
  }
  var map;
  var rows;
  if (forDashboard) {
    var bounded = getSheetDataBounded_(sheet, DASH_SHEET_ROW_LIMITS_.shoeing);
    map = bounded.map;
    rows = bounded.rows;
  } else {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    map = getHeaderIndexMap_(sheet);
    rows = data.slice(1);
  }
  if (!rows.length) {
    return [];
  }
  var cDate = findCol_(map, ["Date", "date"]);
  var cHorse = findCol_(map, ["Horse", "horse"]);
  var cType = findCol_(map, ["Type", "Shoeing Type"]);
  var cFarrier = findCol_(map, ["Farrier", "farrier"]);
  var cNext = findCol_(map, ["NextDue", "Next Due"]);
  var latest = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var horse = String(row[cHorse] || "").trim();
    if (!horse) {
      continue;
    }
    var date = new Date(row[cDate]);
    var next = new Date(row[cNext]);
    if (!latest[horse] || latest[horse].date < date) {
      latest[horse] = {
        horse: horse,
        type: cType >= 0 ? row[cType] : "",
        date: date,
        next: next,
        farrier: cFarrier >= 0 ? row[cFarrier] : ""
      };
    }
  }
  var out = [];
  var today = new Date();
  for (var h in latest) {
    var m = latest[h];
    var diff = Math.round((m.next - today) / (1000 * 60 * 60 * 24));
    var alert = diff < 0 ? "🔴 OVERDUE" : diff <= 7 ? "🟡 DUE SOON" : "✅ OK";
    out.push({
      horse: h,
      lastDate: m.date.toDateString(),
      nextDue: m.next.toDateString(),
      daysLeft: diff,
      farrier: m.farrier,
      type: m.type,
      alert: alert
    });
  }
  out.sort(function(a, b) {
    return a.daysLeft - b.daysLeft;
  });
  return out;
}

// ─── VACCINATION STATUS ───────────────────────────────────────
function getVaccinationStatus(forDashboard) {
  var sheet = getSS_().getSheetByName("VaccinationData");
  if (!sheet) {
    return [];
  }
  var map;
  var rows;
  if (forDashboard) {
    var bounded = getSheetDataBounded_(sheet, DASH_SHEET_ROW_LIMITS_.vaccination);
    map = bounded.map;
    rows = bounded.rows;
  } else {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    map = getHeaderIndexMap_(sheet);
    rows = data.slice(1);
  }
  if (!rows.length) {
    return [];
  }
  var cHorse = findCol_(map, ["Horse", "horse"]);
  var cVaccine = findCol_(map, ["Vaccine", "vaccine"]);
  var cGiven = findCol_(map, ["DateGiven", "Date Given"]);
  var cNext = findCol_(map, ["NextDue", "Next Due"]);
  var latest = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var horse = String(row[cHorse] || "").trim();
    var vaccine = String(row[cVaccine] || "").trim();
    if (!horse || !vaccine) {
      continue;
    }
    var key = horse + "|" + vaccine;
    var next = new Date(row[cNext]);
    var given = row[cGiven];
    if (!latest[key] || latest[key].next < next) {
      latest[key] = { horse: horse, vaccine: vaccine, given: given, next: next };
    }
  }
  var out = [];
  var today = new Date();
  for (var k in latest) {
    var m = latest[k];
    var diff = Math.round((m.next - today) / (1000 * 60 * 60 * 24));
    var alert = diff < 0 ? "🔴 OVERDUE" : diff <= 14 ? "🟡 DUE SOON" : "✅ OK";
    out.push({
      horse: m.horse,
      vaccine: m.vaccine,
      lastGiven: new Date(m.given).toDateString(),
      nextDue: m.next.toDateString(),
      daysLeft: diff,
      vet: "",
      alert: alert
    });
  }
  out.sort(function(a, b) {
    return a.daysLeft - b.daysLeft;
  });
  return out;
}

function healthStatusFromValues_(temp, appetite, water, dung, injury, notes) {
  var t = String(temp || "").toLowerCase();
  var a = String(appetite || "").toLowerCase();
  var w = String(water || "").toLowerCase();
  var d = String(dung || "").toLowerCase();
  var inj = String(injury || "").toLowerCase();
  var n = String(notes || "").toLowerCase();
  if (inj.indexOf("serious") >= 0 || a.indexOf("not eating") >= 0 || a.indexOf("decreased") >= 0 ||
      t === "high" || t.indexOf("abnormal") >= 0 || n.indexOf("vet") >= 0) {
    return "🔴 Attention";
  }
  if (inj.indexOf("minor") >= 0 || a.indexOf("reduced") >= 0 || w === "low" || w.indexOf("decreased") >= 0 ||
      d.indexOf("loose") >= 0 || d.indexOf("diarrhoea") >= 0 || d.indexOf("soft") >= 0 ||
      t.indexOf("slight") >= 0 || t.indexOf("increased") >= 0) {
    return "🟡 Watch";
  }
  return "🟢 Healthy";
}

function healthStatusFromChecklist_(row) {
  row = row || {};
  if (String(row.Status || "").trim()) {
    return String(row.Status).trim();
  }
  var tempNum = parseFloat(String(row.Temperature_C || row.Temperature || row.Temp || "").replace(/[^\d.]/g, ""));
  var tempLabel = "";
  if (!isNaN(tempNum)) {
    if (tempNum < 37.5 || tempNum > 38.5) {
      tempLabel = tempNum > 38.5 ? "High" : "Abnormal";
    } else {
      tempLabel = "Normal";
    }
  } else {
    tempLabel = String(row.Temp || row.Temperature || "");
  }
  var abnormal = function(v) {
    v = String(v || "").toLowerCase();
    return v.indexOf("abnormal") >= 0 || v.indexOf("increased") >= 0 || v.indexOf("decreased") >= 0 ||
      v.indexOf("diarrhoea") >= 0 || v.indexOf("soft") >= 0 || v.indexOf("hard") >= 0;
  };
  if (abnormal(row.Urination) || abnormal(row.Demeanour) || abnormal(row.Skin) ||
      abnormal(row.Feet_Shoes) || abnormal(row.Movement) || abnormal(row.Eyes) ||
      abnormal(row.Skin_Check_Thorough) || abnormal(row.Digital_Pulse) || abnormal(row.Hoof_Temperature)) {
    return "🔴 Attention";
  }
  return healthStatusFromValues_(
    tempLabel,
    row.Appetite,
    row.Drinking || row.Water,
    row.Droppings_Consistency || row.Dung,
    row.Movement,
    row.Notes
  );
}

function getHorseHealthCheckHeaders_() {
  return [
    "Date", "Location", "Horse", "Horse_Age", "Horse_Colour", "Horse_Sex",
    "Temperature_C", "Heart_Rate", "Respiratory_Rate",
    "Droppings_Amount", "Droppings_Consistency",
    "Appetite", "Drinking", "Urination", "Urination_Details",
    "Demeanour", "Demeanour_Details", "Skin", "Skin_Details",
    "Digital_Pulse", "Hoof_Temperature", "Feet_Shoes", "Feet_Shoes_Details",
    "Movement", "Movement_Details", "Eyes", "Eyes_Details",
    "Body_Condition", "Skin_Check_Thorough", "Skin_Check_Details",
    "Vet_Check_Date", "Dental_Check_Date",
    "Vaccination_Tetanus", "Vaccination_Influenza", "Vaccination_Other", "Vaccination_Notes",
    "Checked_By", "Status",
    "Temp", "Water", "Dung", "Injury", "Medication", "Notes"
  ];
}

function ensureHorseHealthSheet_() {
  var sheet = getSheetByNameSafe_(["HorseHealthData"]);
  if (!sheet) {
    sheet = getSS_().insertSheet("HorseHealthData");
  }
  var required = getHorseHealthCheckHeaders_();
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  var c;
  for (c = 0; c < existing.length; c++) {
    var h = String(existing[c] || "").trim();
    if (h) {
      map[h.toLowerCase()] = true;
    }
  }
  var merged = existing.slice();
  for (c = 0; c < required.length; c++) {
    if (!map[String(required[c]).toLowerCase()]) {
      merged.push(required[c]);
    }
  }
  while (merged.length && !String(merged[merged.length - 1] || "").trim()) {
    merged.pop();
  }
  sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
  return sheet;
}

function normalizeHealthCheckPayload_(payload) {
  payload = payload || {};
  var tempC = payload.Temperature_C || payload.Temperature || "";
  var tempNum = parseFloat(String(tempC).replace(/[^\d.]/g, ""));
  var tempLegacy = "";
  if (!isNaN(tempNum)) {
    if (tempNum > 38.5) {
      tempLegacy = "High";
    } else if (tempNum < 37.5) {
      tempLegacy = "Slight High";
    } else {
      tempLegacy = "Normal";
    }
  }
  payload.Temp = tempLegacy || payload.Temp || "";
  payload.Water = payload.Drinking || payload.Water || "";
  payload.Dung = payload.Droppings_Consistency || payload.Dung || "";
  var inj = "None";
  if (String(payload.Movement || "").toLowerCase().indexOf("abnormal") >= 0 ||
      String(payload.Feet_Shoes || "").toLowerCase().indexOf("abnormal") >= 0) {
    inj = "Minor";
  }
  payload.Injury = inj;
  payload.Medication = payload.Vaccination_Notes || payload.Medication || "";
  var noteParts = [];
  ["Urination_Details", "Demeanour_Details", "Skin_Details", "Feet_Shoes_Details",
    "Movement_Details", "Eyes_Details", "Skin_Check_Details", "Vaccination_Notes", "Droppings_Amount"].forEach(function(k) {
    if (payload[k]) {
      noteParts.push(k.replace(/_/g, " ") + ": " + payload[k]);
    }
  });
  if (payload.Heart_Rate) {
    noteParts.push("HR: " + payload.Heart_Rate);
  }
  if (payload.Respiratory_Rate) {
    noteParts.push("RR: " + payload.Respiratory_Rate);
  }
  if (payload.Body_Condition) {
    noteParts.push("Body: " + payload.Body_Condition);
  }
  payload.Notes = noteParts.join(" | ");
  payload.Status = healthStatusFromChecklist_(payload);
  if (payload.Vaccination_Tetanus === "on" || payload.Vaccination_Tetanus === true) {
    payload.Vaccination_Tetanus = "Yes";
  }
  if (payload.Vaccination_Influenza === "on" || payload.Vaccination_Influenza === true) {
    payload.Vaccination_Influenza = "Yes";
  }
  return payload;
}

function getHorseDetailsForHealthCheck_(horseName, token) {
  validateSessionToken_(token);
  horseName = String(horseName || "").trim();
  if (!horseName) {
    return { age: "", colour: "", sex: "" };
  }
  var horses = getHorsesList_();
  var i;
  for (i = 0; i < horses.length; i++) {
    if (String(horses[i].name || "").trim().toLowerCase() === horseName.toLowerCase()) {
      return {
        age: horses[i].age || "",
        colour: horses[i].breed || "",
        sex: horses[i].gender || ""
      };
    }
  }
  return { age: "", colour: "", sex: "" };
}

// ─── HEALTH SUMMARY ───────────────────────────────────────────
function getHealthSummary(forDashboard) {
  var sheet = getSS_().getSheetByName("HorseHealthData");
  if (!sheet) {
    return [];
  }
  var map;
  var rows;
  if (forDashboard) {
    var bounded = getSheetDataBounded_(sheet, DASH_SHEET_ROW_LIMITS_.health);
    map = bounded.map;
    rows = bounded.rows;
  } else {
    var full = sheet.getDataRange().getValues();
    if (full.length < 2) {
      return [];
    }
    map = getHeaderIndexMap_(sheet);
    rows = full.slice(1);
  }
  if (!rows.length) {
    return [];
  }
  var cDate = findCol_(map, ["Date", "date"]);
  if (cDate < 0) {
    cDate = 0;
  }
  var cHorse = findCol_(map, ["Horse", "horse"]);
  if (cHorse < 0) {
    cHorse = 2;
  }
  function colVal_(row, names) {
    var col = findCol_(map, names);
    return col >= 0 ? row[col] : "";
  }
  var latest = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var horse = String(row[cHorse] || "").trim();
    if (!horse) {
      continue;
    }
    var date = new Date(row[cDate]);
    if (!latest[horse] || latest[horse].date < date) {
      latest[horse] = {
        horse: horse,
        date: date,
        location: colVal_(row, ["Location", "location"]),
        temp: colVal_(row, ["Temp", "Temperature", "Temperature_C"]),
        temperatureC: colVal_(row, ["Temperature_C", "Temperature"]),
        heartRate: colVal_(row, ["Heart_Rate", "Heart Rate"]),
        respiratoryRate: colVal_(row, ["Respiratory_Rate", "Respiratory Rate"]),
        appetite: colVal_(row, ["Appetite", "appetite"]),
        water: colVal_(row, ["Drinking", "Water", "water"]),
        dung: colVal_(row, ["Droppings_Consistency", "Dung", "dung"]),
        droppingsAmount: colVal_(row, ["Droppings_Amount"]),
        injury: colVal_(row, ["Injury", "injury"]),
        medication: colVal_(row, ["Medication", "medication"]),
        notes: colVal_(row, ["Notes", "notes"]),
        statusRaw: colVal_(row, ["Status", "status"]),
        demeanour: colVal_(row, ["Demeanour", "Demeanour"]),
        movement: colVal_(row, ["Movement", "Movement"]),
        eyes: colVal_(row, ["Eyes", "eyes"]),
        bodyCondition: colVal_(row, ["Body_Condition", "Body Condition"]),
        vetCheckDate: colVal_(row, ["Vet_Check_Date", "Vet Check Date"]),
        dentalCheckDate: colVal_(row, ["Dental_Check_Date", "Dental Check Date"]),
        checkedBy: colVal_(row, ["Checked_By", "Checked By"])
      };
    }
  }
  var out = [];
  for (var h in latest) {
    var m = latest[h];
    var status = healthStatusFromChecklist_({
      Status: m.statusRaw,
      Temp: m.temp,
      Temperature_C: m.temperatureC,
      Appetite: m.appetite,
      Drinking: m.water,
      Droppings_Consistency: m.dung,
      Movement: m.movement,
      Demeanour: m.demeanour,
      Eyes: m.eyes,
      Notes: m.notes
    });
    var score = status === "🟢 Healthy" ? 90 : status === "🟡 Watch" ? 65 : 40;
    out.push({
      horse: h,
      score: score,
      status: status,
      lastCheck: m.date.toDateString(),
      date: m.date,
      location: m.location,
      temp: m.temp,
      temperatureC: m.temperatureC,
      heartRate: m.heartRate,
      respiratoryRate: m.respiratoryRate,
      appetite: m.appetite,
      water: m.water,
      dung: m.dung,
      droppingsAmount: m.droppingsAmount,
      injury: m.injury,
      medication: m.medication,
      notes: m.notes,
      demeanour: m.demeanour,
      movement: m.movement,
      eyes: m.eyes,
      bodyCondition: m.bodyCondition,
      vetCheckDate: m.vetCheckDate,
      dentalCheckDate: m.dentalCheckDate,
      checkedBy: m.checkedBy
    });
  }
  out.sort(function(a, b) {
    return a.score - b.score;
  });
  return out;
}

// ─── ACTIVITY SUMMARY (Last 7 days) ──────────────────────────
function getActivitySummary() {
  var sheet = getSS_().getSheetByName("ActivityData");
  if (!sheet) {
    return [];
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var cDate = findCol_(map, ["Date", "date"]);
  var cHorse = findCol_(map, ["Horse", "horse"]);
  var cActivity = findCol_(map, ["Activity", "activity"]);
  var cDuration = findCol_(map, ["Duration", "duration"]);
  var today = new Date();
  var sevenDaysAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  var horses = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var date = new Date(row[cDate]);
    if (isNaN(date.getTime()) || date < sevenDaysAgo) {
      continue;
    }
    var horse = String(row[cHorse] || "").trim();
    if (!horse) {
      continue;
    }
    if (!horses[horse]) {
      horses[horse] = { horse: horse, sessions: 0, totalMins: 0, activities: {} };
    }
    horses[horse].sessions++;
    horses[horse].totalMins += parseRoutineDurationToMinutes_(row[cDuration]);
    var act = String(row[cActivity] || "Unknown");
    horses[horse].activities[act] = (horses[horse].activities[act] || 0) + 1;
  }
  var out = [];
  for (var h in horses) {
    var m = horses[h];
    var topActivity = "None";
    var acts = m.activities;
    for (var a in acts) {
      if (!topActivity || acts[a] > (acts[topActivity] || 0)) {
        topActivity = a;
      }
    }
    out.push({ horse: h, sessions: m.sessions, totalMins: m.totalMins, topActivity: topActivity });
  }
  out.sort(function(a, b) {
    return b.sessions - a.sessions;
  });
  return out;
}

// ─── MEDICAL HISTORY ─────────────────────────────────────────
function getMedicalHistory(maxRows, forDashboard) {
  var sheet = getSS_().getSheetByName("MedicalHistory");
  if (!sheet) {
    return [];
  }
  var map;
  var rows;
  if (forDashboard) {
    var rowLimit = Math.max((maxRows || 25) * 4, DASH_SHEET_ROW_LIMITS_.medical);
    var bounded = getSheetDataBounded_(sheet, rowLimit);
    map = bounded.map;
    rows = bounded.rows;
  } else {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    map = getHeaderIndexMap_(sheet);
    rows = data.slice(1);
  }
  if (!rows.length) {
    return [];
  }
  var cDate = findCol_(map, ["Date", "date"]);
  var cHorse = findCol_(map, ["Horse", "horse"]);
  var cIssue = findCol_(map, ["Issue", "issue"]);
  var cDiag = findCol_(map, ["Diagnosis", "diagnosis"]);
  var cTreat = findCol_(map, ["Treatment", "treatment"]);
  var cMed = findCol_(map, ["Medication", "Medicines"]);
  var cVet = findCol_(map, ["Vet", "vet"]);
  var cFollow = findCol_(map, ["FollowUp", "Follow Up", "Follow-up"]);
  var cNotes = findCol_(map, ["Notes", "notes"]);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var horse = String(row[cHorse] || "").trim();
    if (!horse) {
      continue;
    }
    out.push({
      date: row[cDate] ? new Date(row[cDate]).toDateString() : "",
      horse: horse,
      issue: cIssue >= 0 ? row[cIssue] : "",
      diagnosis: cDiag >= 0 ? row[cDiag] : "",
      treatment: cTreat >= 0 ? row[cTreat] : "",
      medicines: cMed >= 0 ? row[cMed] : "",
      followUp: cFollow >= 0 ? row[cFollow] : "",
      vet: cVet >= 0 ? row[cVet] : "",
      status: cNotes >= 0 ? row[cNotes] : ""
    });
  }
  out.sort(function(a, b) {
    return new Date(b.date) - new Date(a.date);
  });
  if (maxRows && out.length > maxRows) {
    return out.slice(0, maxRows);
  }
  return out;
}

// ─── DASHBOARD COUNTS ─────────────────────────────────────────
function getDashboardCounts() {
  var horses = getHorsesList_().length;
  var health = getHealthSummary();
  var shoeing = getShoeingStatus();
  var vaccines = getVaccinationStatus();
  var alerts = 0;
  var i;
  for (i = 0; i < shoeing.length; i++) {
    if (shoeing[i].alert !== "✅ OK") {
      alerts++;
    }
  }
  for (i = 0; i < vaccines.length; i++) {
    if (vaccines[i].alert !== "✅ OK") {
      alerts++;
    }
  }
  for (i = 0; i < health.length; i++) {
    if (health[i].status === "🔴 Attention") {
      alerts++;
    }
  }
  return {
    totalHorses: horses,
    totalAlerts: alerts,
    healthyHorses: countHealthStatus_(health, "🟢 Healthy"),
    shoeingDue: countAlertsNotOk_(shoeing),
    vaccineDue: countAlertsNotOk_(vaccines)
  };
}

function countAlertsNotOk_(rows) {
  var n = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].alert !== "✅ OK") {
      n++;
    }
  }
  return n;
}

// ─── DAILY EMAIL REPORT ───────────────────────────────────────
function getUsersSheetEmails_() {
  var sheet = getSheetByNameSafe_(["Users"]);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  var data = sheet.getDataRange().getValues();
  var map = getHeaderIndexMap_(sheet);
  var colEmail = findCol_(map, ["Email", "email"]);
  var colName = findCol_(map, ["Name", "name"]);
  var colRole = findCol_(map, ["Role", "role"]);
  var out = [];
  var i;
  for (i = 1; i < data.length; i++) {
    var row = data[i];
    var email = String(row[colEmail >= 0 ? colEmail : 0] || "").trim();
    if (!email || email.indexOf("@") < 0) {
      continue;
    }
    out.push({
      email: email,
      name: colName >= 0 ? String(row[colName] || "").trim() : "",
      role: colRole >= 0 ? String(row[colRole] || "").trim() : ""
    });
  }
  return out;
}

function resolveUserEmailFromSheet_(user) {
  user = user || {};
  var email = String(user.email || "").trim();
  if (email && email.indexOf("@") >= 0) {
    return email;
  }
  var nameKey = String(user.name || "").trim().toLowerCase();
  var list = getUsersSheetEmails_();
  var i;
  for (i = 0; i < list.length; i++) {
    if (nameKey && list[i].name.toLowerCase() === nameKey) {
      return list[i].email;
    }
    if (email && list[i].email.toLowerCase() === email.toLowerCase()) {
      return list[i].email;
    }
  }
  return email;
}

function formatOperationsReportEmailBody_(report) {
  report = report || {};
  var s = report.summary || {};
  var body = "KINGS EQUESTRIAN – " + String(report.periodLabel || "Report").toUpperCase() + "\n";
  body += "Location: " + (report.location || "All Locations") + "\n";
  body += "Period: " + (report.dateRange || "") + "\n";
  body += "Generated: " + (report.generatedAt || "") + "\n\n";

  body += "SUMMARY\n";
  body += "Horses: " + (s.horses || 0) + "\n";
  body += "Logs recorded: " + (s.logsRecorded || 0) + "\n";
  body += "Total feed (kg): " + (s.totalFeedKg || 0) + "\n";
  body += "Total water (L): " + (s.totalWaterL || 0) + "\n";
  body += "Avg feed/day (kg): " + (s.avgFeedPerDay || 0) + "\n";
  body += "Avg water/day (L): " + (s.avgWaterPerDay || 0) + "\n";
  body += "Activity sessions: " + (s.activitySessions || 0) + "\n";
  body += "Activity hours: " + (s.activityHours || 0) + "\n";
  body += "Alerts: " + (s.alertCount || 0) + "\n\n";

  var mi;
  for (mi = 0; mi < (report.metrics || []).length; mi++) {
    var m = report.metrics[mi];
    body += "• " + m.label + ": " + m.value + "\n";
  }
  body += "\n";

  if ((report.alerts || []).length) {
    body += "ALERTS\n";
    var ai;
    for (ai = 0; ai < report.alerts.length && ai < 20; ai++) {
      body += "• " + report.alerts[ai].title + " — " + report.alerts[ai].sub + "\n";
    }
    body += "\n";
  }

  if ((report.horseRows || []).length) {
    body += "HORSE DETAILS\n";
    var hi;
    for (hi = 0; hi < report.horseRows.length && hi < 30; hi++) {
      var h = report.horseRows[hi];
      body += h.horse + " | Feed " + h.feedTotal + " kg | Water " + h.waterTotal + " L | Status: " + h.status + "\n";
    }
    body += "\n";
  }

  if ((report.insights || []).length) {
    body += "INSIGHTS\n";
    var ii;
    for (ii = 0; ii < report.insights.length; ii++) {
      body += "• " + report.insights[ii].label + ": " + report.insights[ii].value + "\n";
    }
    body += "\n";
  }

  body += "— Kings Equestrian Operations\n";
  return body;
}

function sendOperationsReportEmail(period, location, token) {
  var user = validateSessionToken_(token);
  var recipientEmail = resolveUserEmailFromSheet_(user);
  if (!recipientEmail || recipientEmail.indexOf("@") < 0) {
    throw new Error("No email address found in Users sheet for your account. Add an Email in the Users tab.");
  }
  location = String(location || "").trim();
  if (location && !isAllLocations_(location)) {
    location = assertLocationAccess_(user, location);
  } else if (!isAdminRole_(user.role)) {
    location = resolveLocationName_(user.location || user.locationId);
  }
  var report = buildOperationsReport_(period || "daily", location);
  var subject = "Kings Equestrian " + report.periodLabel + " Report – " + report.location;
  var body = formatOperationsReportEmailBody_(report);
  MailApp.sendEmail(recipientEmail, subject, body);
  return { success: true, sentTo: recipientEmail, period: report.period, periodLabel: report.periodLabel };
}

function dailyEmailReport() {
  var admins = getUsersSheetEmails_().filter(function(u) {
    return isAdminRole_(u.role);
  });
  if (!admins.length) {
    admins = getUsersSheetEmails_();
  }
  if (!admins.length) {
    Logger.log("dailyEmailReport: no emails in Users sheet");
    return;
  }
  var report = buildOperationsReport_("daily", "");
  var subject = "Kings Equestrian Daily Report – " + new Date().toDateString();
  var body = formatOperationsReportEmailBody_(report);
  var ei;
  for (ei = 0; ei < admins.length; ei++) {
    try {
      MailApp.sendEmail(admins[ei].email, subject, body);
    } catch (mailErr) {
      Logger.log("dailyEmailReport " + admins[ei].email + ": " + mailErr);
    }
  }
}
function generateHash() {
  var password = "admin@123"; // change to the password you want, then run hashUserPassword
  var hash = hashPassword(password);
  Logger.log("Hash for Password column: " + hash);
  Logger.log('Or run: hashUserPassword("admin@kingsequestrian.com", "admin@123")');
  return hash;
}

// ─── SHEET HELPERS (header-based columns) ─────────────────────
function normalizeHeaderKey_(header) {
  return String(header || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function clearSheetHeaderCache_(sheet) {
  if (!sheet) {
    return;
  }
  var cacheKey = sheet.getParent().getId() + "|" + sheet.getName();
  delete __sheetHeaderCache_[cacheKey];
}

function getHeaderIndexMap_(sheet) {
  var cacheKey = sheet.getParent().getId() + "|" + sheet.getName();
  if (__sheetHeaderCache_[cacheKey]) {
    return __sheetHeaderCache_[cacheKey];
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var c = 0; c < headers.length; c++) {
    var key = normalizeHeaderKey_(headers[c]);
    if (key) {
      map[key] = c;
    }
  }
  __sheetHeaderCache_[cacheKey] = map;
  return map;
}

function findCol_(map, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var key = normalizeHeaderKey_(candidates[i]);
    if (Object.prototype.hasOwnProperty.call(map, key)) {
      return map[key];
    }
  }
  return -1;
}

/** Find column by header name; optional default index (e.g. 1 for Horse_Name in col B). */
function resolveColumn_(sheet, candidates, defaultIndex) {
  var map = getHeaderIndexMap_(sheet);
  var col = findCol_(map, candidates);
  if (col >= 0) {
    return col;
  }
  if (typeof defaultIndex === "number" && defaultIndex >= 0) {
    return defaultIndex;
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var h = 0; h < headers.length; h++) {
    var key = normalizeHeaderKey_(headers[h]);
    for (var c = 0; c < candidates.length; c++) {
      var want = normalizeHeaderKey_(candidates[c]);
      if (key && want && (key === want || key.indexOf(want) >= 0)) {
        return h;
      }
    }
  }
  return -1;
}

function sameLocation_(a, b) {
  if (isAllLocations_(a) || isAllLocations_(b)) {
    return isAllLocations_(a) && isAllLocations_(b);
  }
  var ra = resolveLocationKey_(a);
  var rb = resolveLocationKey_(b);
  if (ra.name && rb.name &&
      String(ra.name).trim().toLowerCase() === String(rb.name).trim().toLowerCase()) {
    return true;
  }
  if (ra.id && rb.id &&
      String(ra.id).trim().toLowerCase() === String(rb.id).trim().toLowerCase()) {
    return true;
  }
  return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
}

function getLocationsRecords_() {
  if (__locationsRecordsCache_) {
    return __locationsRecordsCache_;
  }
  var sheet = getSheetByNameSafe_(["Locations"]);
  if (!sheet) {
    __locationsRecordsCache_ = getActivityLocationsFromHorses_().map(function(name, idx) {
      return { id: "LOC" + String(idx + 1).padStart(3, "0"), name: name, manager: "", phone: "" };
    });
    return __locationsRecordsCache_;
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var idCol = findCol_(map, ["Location_ID", "Location ID", "location_id"]);
  var nameCol = resolveColumn_(sheet, ["Location_Name", "Location Name", "Location"], 1);
  var mgrCol = findCol_(map, ["Manager", "manager"]);
  var phoneCol = findCol_(map, ["Phone", "phone"]);
  if (nameCol < 0) {
    return [];
  }
  var seen = {};
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[nameCol] || "").trim();
    if (!name || seen[name.toLowerCase()]) {
      continue;
    }
    seen[name.toLowerCase()] = true;
    var id = idCol >= 0 ? String(row[idCol] || "").trim() : "";
    if (!id) {
      id = "LOC" + String(out.length + 1).padStart(3, "0");
    }
    out.push({
      id: id,
      name: name,
      manager: mgrCol >= 0 ? String(row[mgrCol] || "").trim() : "",
      phone: phoneCol >= 0 ? String(row[phoneCol] || "").trim() : ""
    });
  }
  out.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });
  __locationsRecordsCache_ = out;
  return out;
}

function resolveLocationKey_(ref) {
  ref = String(ref || "").trim();
  if (!ref) {
    return { id: "", name: "", label: "" };
  }
  var cacheKey = ref.toLowerCase();
  if (__locationResolveCache_[cacheKey]) {
    return __locationResolveCache_[cacheKey];
  }
  var records = getLocationsRecords_();
  var lower = ref.toLowerCase();
  var i;
  for (i = 0; i < records.length; i++) {
    var rec = records[i];
    if (String(rec.id || "").trim().toLowerCase() === lower) {
      var byId = {
        id: rec.id,
        name: rec.name,
        label: rec.id ? rec.id + " · " + rec.name : rec.name
      };
      __locationResolveCache_[cacheKey] = byId;
      return byId;
    }
    if (String(rec.name || "").trim().toLowerCase() === lower) {
      var byName = {
        id: rec.id,
        name: rec.name,
        label: rec.id ? rec.id + " · " + rec.name : rec.name
      };
      __locationResolveCache_[cacheKey] = byName;
      return byName;
    }
  }
  var fallback = { id: ref, name: ref, label: ref };
  __locationResolveCache_[cacheKey] = fallback;
  return fallback;
}

function resolveLocationName_(ref) {
  return resolveLocationKey_(ref).name || String(ref || "").trim();
}

/** Normalize a Location cell (ID or name) to canonical Location_Name for matching. */
function normalizeSheetLocation_(ref) {
  var s = String(ref || "").trim();
  if (!s) {
    return "";
  }
  return resolveLocationName_(s) || s;
}

function getDefaultAdminLocationId_() {
  return "LOC004";
}

function getDefaultAdminLocationName_() {
  var rec = resolveLocationKey_(getDefaultAdminLocationId_());
  return rec.name || "";
}

function matchesLocationFilter_(rowLoc, filterRef) {
  if (isAllLocations_(filterRef) || !String(filterRef || "").trim()) {
    return true;
  }
  return sameLocation_(rowLoc, filterRef);
}

function getSheetByNameSafe_(names) {
  for (var i = 0; i < names.length; i++) {
    var sheet = getSS_().getSheetByName(names[i]);
    if (sheet) {
      return sheet;
    }
  }
  return null;
}

function getLocationsList_() {
  return getLocationsRecords_().map(function(rec) {
    return rec.name;
  });
}

function getActivityLocationsFromHorses_() {
  return getActivityLocations_();
}

function getHorsesList_(healthStatusByName) {
  var sheet = getSheetByNameSafe_(["Horses"]);
  if (!sheet) {
    return [];
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var nameCol = resolveColumn_(sheet, ["Horse_Name", "Horse Name", "Horse"], 1);
  if (nameCol < 0) {
    return [];
  }
  var breedCol = findCol_(map, ["Breed", "breed"]);
  var ageCol = findCol_(map, ["Age", "age"]);
  var locCol = findCol_(map, ["Location", "location"]);
  var ownerCol = findCol_(map, ["Owner", "owner"]);
  var statusCol = findCol_(map, ["Status", "status"]);
  var genderCol = findCol_(map, ["Gender", "gender"]);
  var idCol = findCol_(map, ["Horse_ID", "Horse ID", "id"]);
  var photoCol = findCol_(map, ["Photo_URL", "Photo", "Photo Url"]);
  var photoIdCol = findCol_(map, ["Photo_File_ID", "Photo File ID"]);
  var chipCol = findCol_(map, ["Chip_No", "Chip No.", "Chip Number"]);
  var efiCol = findCol_(map, ["EFI_ID", "EFI ID"]);
  var healthMap = healthStatusByName || {};
  if (healthStatusByName === undefined) {
    try {
      getHealthSummary().forEach(function(h) {
        healthMap[h.horse] = h.status;
      });
    } catch (e) {}
  }
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[nameCol] || "").trim();
    if (!name) {
      continue;
    }
    var sheetStatus = statusCol >= 0 ? String(row[statusCol] || "").trim() : "";
    out.push({
      id: idCol >= 0 ? String(row[idCol] || "").trim() : "",
      name: name,
      breed: breedCol >= 0 ? String(row[breedCol] || "").trim() : "",
      age: ageCol >= 0 ? String(row[ageCol] || "").trim() : "",
      location: locCol >= 0 ? normalizeSheetLocation_(row[locCol]) : "",
      owner: ownerCol >= 0 ? String(row[ownerCol] || "").trim() : "",
      status: sheetStatus,
      gender: genderCol >= 0 ? String(row[genderCol] || "").trim() : "",
      healthStatus: healthMap[name] || "",
      weight: "",
      facilityMultiplier: "",
      leaseRider: "",
      leaseDate: "",
      presentMedication: "",
      trainer: "",
      groom: "",
      chipNo: chipCol >= 0 ? String(row[chipCol] || "").trim() : "",
      efiId: efiCol >= 0 ? String(row[efiCol] || "").trim() : "",
      photoFileId: photoIdCol >= 0 ? String(row[photoIdCol] || "").trim() : "",
      photoUrl: buildProfilePhotoDisplayUrl_(
        photoCol >= 0 ? String(row[photoCol] || "").trim() : "",
        photoIdCol >= 0 ? String(row[photoIdCol] || "").trim() : ""
      )
    });
  }
  out.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });
  return enrichHorsesFromProfiles_(out);
}

function getTrainersList_() {
  var sheet = getSheetByNameSafe_(["Trainers"]);
  if (!sheet) {
    return [];
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var nameCol = resolveColumn_(sheet, ["Trainer_Name", "Trainer Name", "Trainer"], 1);
  if (nameCol < 0) {
    return [];
  }
  var specCol = findCol_(map, ["Specialization", "specialization"]);
  var locCol = findCol_(map, ["Location", "location"]);
  var phoneCol = findCol_(map, ["Phone", "phone"]);
  var expCol = findCol_(map, ["Experience_Years", "Experience Years"]);
  var photoCol = findCol_(map, ["Photo_URL", "Photo", "Photo Url"]);
  var photoIdCol = findCol_(map, ["Photo_File_ID", "Photo File ID"]);
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[nameCol] || "").trim();
    if (!name) {
      continue;
    }
    var statusCol = findCol_(map, ["Status", "status"]);
    var photoFileId = photoIdCol >= 0 ? String(row[photoIdCol] || "").trim() : "";
    if (!photoFileId && photoCol >= 0) {
      photoFileId = extractDriveFileId_(String(row[photoCol] || "").trim());
    }
    out.push({
      name: name,
      specialization: specCol >= 0 ? String(row[specCol] || "").trim() : "",
      location: locCol >= 0 ? normalizeSheetLocation_(row[locCol]) : "",
      phone: phoneCol >= 0 ? String(row[phoneCol] || "").trim() : "",
      experienceYears: expCol >= 0 ? String(row[expCol] || "").trim() : "",
      status: statusCol >= 0 ? String(row[statusCol] || "").trim() || "Active" : "Active",
      photoFileId: photoFileId,
      photoUrl: buildProfilePhotoDisplayUrl_(
        photoCol >= 0 ? String(row[photoCol] || "").trim() : "",
        photoFileId
      )
    });
  }
  return out;
}

function getGroomsList_() {
  var sheet = getSheetByNameSafe_(["Grooms"]);
  if (!sheet) {
    return [];
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var nameCol = resolveColumn_(sheet, ["Groom_Name", "Groom Name", "Groom"], 1);
  if (nameCol < 0) {
    return [];
  }
  var locCol = findCol_(map, ["Location", "location"]);
  var phoneCol = findCol_(map, ["Phone", "phone"]);
  var shiftCol = findCol_(map, ["Shift", "shift"]);
  var statusCol = findCol_(map, ["Status", "status"]);
  var photoCol = findCol_(map, ["Photo_URL", "Photo", "Photo Url"]);
  var photoIdCol = findCol_(map, ["Photo_File_ID", "Photo File ID"]);
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[nameCol] || "").trim();
    if (!name) {
      continue;
    }
    var photoFileId = photoIdCol >= 0 ? String(row[photoIdCol] || "").trim() : "";
    if (!photoFileId && photoCol >= 0) {
      photoFileId = extractDriveFileId_(String(row[photoCol] || "").trim());
    }
    out.push({
      name: name,
      location: locCol >= 0 ? normalizeSheetLocation_(row[locCol]) : "",
      phone: phoneCol >= 0 ? String(row[phoneCol] || "").trim() : "",
      shift: shiftCol >= 0 ? String(row[shiftCol] || "").trim() : "",
      status: statusCol >= 0 ? String(row[statusCol] || "").trim() || "Active" : "Active",
      photoFileId: photoFileId,
      photoUrl: buildProfilePhotoDisplayUrl_(
        photoCol >= 0 ? String(row[photoCol] || "").trim() : "",
        photoFileId
      )
    });
  }
  return out;
}

function getInventoryMinMap_() {
  var sheet = getSheetByNameSafe_(["Inventory Items"]);
  if (!sheet) {
    return {};
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return {};
  }
  var map = getHeaderIndexMap_(sheet);
  var cName = findCol_(map, ["Item_Name", "Item Name", "Item"]);
  var cMin = findCol_(map, ["Min_Stock_Level", "Min Stock Level"]);
  var out = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[cName] || "").trim();
    if (!name) {
      continue;
    }
    out[name] = cMin >= 0 ? parseFloat(row[cMin]) || 0 : 0;
  }
  return out;
}

function getDailyInventoryRows_(maxDataRows) {
  var sheet = getSheetByNameSafe_(["Daily Inventory", "Daily Inventory Summary"]);
  if (!sheet) {
    return [];
  }
  var map;
  var rows;
  if (maxDataRows) {
    var bounded = getSheetDataBounded_(sheet, maxDataRows);
    map = bounded.map;
    rows = bounded.rows;
  } else {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    map = getHeaderIndexMap_(sheet);
    rows = data.slice(1);
  }
  if (!rows.length) {
    return [];
  }
  function cell(row, col) {
    return col >= 0 ? row[col] : "";
  }
  var cDate = findCol_(map, ["Date", "date"]);
  var cLoc = findCol_(map, ["Location", "location"]);
  var cItem = findCol_(map, ["Item", "item"]);
  var cOpen = findCol_(map, ["Opening", "opening"]);
  var cAdd = findCol_(map, ["Added", "added"]);
  var cClose = findCol_(map, ["Closing", "closing"]);
  var cCons = findCol_(map, ["Consumption", "consumption"]);
  var cStatus = findCol_(map, ["Status", "status"]);
  var locCache = {};
  function normLoc(raw) {
    var s = String(raw || "").trim();
    if (!s) {
      return "";
    }
    var k = s.toLowerCase();
    if (locCache[k]) {
      return locCache[k];
    }
    locCache[k] = normalizeSheetLocation_(s);
    return locCache[k];
  }
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    out.push({
      date: cell(row, cDate),
      location: cLoc >= 0 ? normLoc(cell(row, cLoc)) : "",
      item: cell(row, cItem),
      opening: cell(row, cOpen),
      added: cell(row, cAdd),
      closing: cell(row, cClose),
      consumption: cell(row, cCons),
      status: cell(row, cStatus)
    });
  }
  return out;
}

// ─── HORSE ACTIVITY FORM DROPDOWNS ────────────────────────────
function getActivityFormLocations(token) {
  var user = validateSessionToken_(token);
  if (isAdminRole_(user.role)) {
    return getActivityLocations_();
  }
  var loc = String(user.location || "").trim();
  return loc ? [loc] : [];
}

function getActivityHorsesForLocationForUser_(location, user) {
  var profiles = filterHorseProfilesForUser_(getHorseProfilesForLocation_(location), user);
  var out = [];
  var seen = {};
  var i;
  for (i = 0; i < profiles.length; i++) {
    var h = profiles[i];
    var name = String(h.name || "").trim();
    if (!name || seen[name.toLowerCase()]) {
      continue;
    }
    seen[name.toLowerCase()] = true;
    var breed = String(h.breed || "").trim();
    out.push({
      value: name,
      label: breed ? name + " (" + breed + ")" : name
    });
  }
  out.sort(function(a, b) {
    return a.label.localeCompare(b.label);
  });
  return out;
}

function getActivityOptionsForLocation(location, token) {
  var user = validateSessionToken_(token);
  location = assertLocationAccess_(user, location);
  if (!location) {
    return { horses: [], trainers: [] };
  }
  var assign = getUserHorseAssignmentMode_(user);
  return {
    horses: getActivityHorsesForLocationForUser_(location, user),
    trainers: getActivityTrainersForLocation_(location),
    assignMode: assign.mode,
    staffName: assign.staffName
  };
}

/** Trainers and grooms at a location — for Add Horse form dropdowns. */
function getHorseFormStaffForLocation(location, token) {
  var user = validateSessionToken_(token);
  location = assertLocationAccess_(user, location);
  if (!location) {
    return { trainers: [], grooms: [] };
  }
  return {
    trainers: getActivityTrainersForLocation_(location),
    grooms: getStockGroomsForLocation_(location)
  };
}

function getActivityLocations_() {
  var fromLocations = getLocationsList_();
  if (fromLocations.length) {
    return fromLocations;
  }
  var sheet = getSS_().getSheetByName("Horses");
  if (!sheet) {
    throw new Error("Horses sheet not found.");
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var locCol = findCol_(map, ["Location", "location"]);
  if (locCol < 0) {
    throw new Error("Location column not found on Horses sheet.");
  }
  var seen = {};
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var loc = String(data[i][locCol] || "").trim();
    if (!loc || seen[loc]) {
      continue;
    }
    seen[loc] = true;
    out.push(loc);
  }
  out.sort(function(a, b) {
    return a.localeCompare(b);
  });
  return out;
}

function getActivityHorsesForLocation_(location) {
  var sheet = getSS_().getSheetByName("Horses");
  if (!sheet) {
    throw new Error("Horses sheet not found.");
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var locCol = findCol_(map, ["Location", "location"]);
  var nameCol = findCol_(map, ["Horse_Name", "Horse Name", "name"]);
  var breedCol = findCol_(map, ["Breed", "breed"]);
  if (locCol < 0 || nameCol < 0) {
    throw new Error("Required columns not found on Horses sheet (Location, Horse_Name).");
  }
  var out = [];
  var seen = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!sameLocation_(row[locCol], location)) {
      continue;
    }
    var name = String(row[nameCol] || "").trim();
    if (!name || seen[name]) {
      continue;
    }
    seen[name] = true;
    var breed = breedCol >= 0 ? String(row[breedCol] || "").trim() : "";
    var label = breed ? name + " (" + breed + ")" : name;
    out.push({ value: name, label: label });
  }
  out.sort(function(a, b) {
    return a.label.localeCompare(b.label);
  });
  return out;
}

function getActivityTrainersForLocation_(location) {
  var sheet = getSS_().getSheetByName("Trainers");
  if (!sheet) {
    throw new Error("Trainers sheet not found.");
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var locCol = findCol_(map, ["Location", "location"]);
  var nameCol = findCol_(map, ["Trainer_Name", "Trainer Name", "name"]);
  if (locCol < 0 || nameCol < 0) {
    throw new Error("Required columns not found on Trainers sheet (Location, Trainer_Name).");
  }
  var seen = {};
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!sameLocation_(row[locCol], location)) {
      continue;
    }
    var name = String(row[nameCol] || "").trim();
    if (!name || seen[name]) {
      continue;
    }
    seen[name] = true;
    out.push(name);
  }
  out.sort(function(a, b) {
    return a.localeCompare(b);
  });
  return out;
}

// ─── STOCK ENTRY FORM DROPDOWNS ───────────────────────────────
function getStockFormBootstrap(token) {
  var user = validateSessionToken_(token);
  var locations = getStockLocationsFromTrainers_();
  if (!isAdminRole_(user.role)) {
    var userLoc = String(user.location || "").trim();
    locations = userLoc ? locations.filter(function(loc) {
      return sameLocation_(loc, userLoc);
    }) : [];
  }
  return withFormLocationOptions_(token, {
    locations: locations,
    items: getStockInventoryItems_()
  });
}

function getStockTrainersForLocation(location, token) {
  var user = validateSessionToken_(token);
  location = assertLocationAccess_(user, location);
  if (!location) {
    return { trainers: [] };
  }
  return { trainers: getActivityTrainersForLocation_(location) };
}

function getStockGroomsForLocation(location, token) {
  var user = validateSessionToken_(token);
  location = assertLocationAccess_(user, location);
  if (!location) {
    return { grooms: [] };
  }
  return { grooms: getStockGroomsForLocation_(location) };
}

function getStockLocationsFromTrainers_() {
  var sheet = getSS_().getSheetByName("Trainers");
  if (!sheet) {
    throw new Error("Trainers sheet not found.");
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var locCol = findCol_(map, ["Location", "location"]);
  if (locCol < 0) {
    throw new Error("Location column not found on Trainers sheet.");
  }
  var seen = {};
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var loc = String(data[i][locCol] || "").trim();
    if (!loc || seen[loc]) {
      continue;
    }
    seen[loc] = true;
    out.push(loc);
  }
  out.sort(function(a, b) {
    return a.localeCompare(b);
  });
  return out;
}

function getStockLocationsFromGrooms_() {
  var sheet = getSS_().getSheetByName("Grooms");
  if (!sheet) {
    throw new Error("Grooms sheet not found.");
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var locCol = findCol_(map, ["Location", "location"]);
  if (locCol < 0) {
    throw new Error("Location column not found on Grooms sheet.");
  }
  var seen = {};
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var loc = String(data[i][locCol] || "").trim();
    if (!loc || seen[loc]) {
      continue;
    }
    seen[loc] = true;
    out.push(loc);
  }
  out.sort(function(a, b) {
    return a.localeCompare(b);
  });
  return out;
}

function getStockGroomsForLocation_(location) {
  var sheet = getSS_().getSheetByName("Grooms");
  if (!sheet) {
    throw new Error("Grooms sheet not found.");
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var locCol = findCol_(map, ["Location", "location"]);
  var nameCol = findCol_(map, ["Groom_Name", "Groom Name", "name"]);
  if (locCol < 0 || nameCol < 0) {
    throw new Error("Required columns not found on Grooms sheet (Location, Groom_Name).");
  }
  var seen = {};
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!sameLocation_(row[locCol], location)) {
      continue;
    }
    var name = String(row[nameCol] || "").trim();
    if (!name || seen[name]) {
      continue;
    }
    seen[name] = true;
    out.push(name);
  }
  out.sort(function(a, b) {
    return a.localeCompare(b);
  });
  return out;
}

function getStockInventoryItems_() {
  var sheet = getSS_().getSheetByName("Inventory Items");
  if (!sheet) {
    throw new Error("Inventory Items sheet not found.");
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return [];
  }
  var map = getHeaderIndexMap_(sheet);
  var itemCol = findCol_(map, ["Item_Name", "Item Name", "item"]);
  if (itemCol < 0) {
    throw new Error("Item_Name column not found on Inventory Items sheet.");
  }
  var seen = {};
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var item = String(data[i][itemCol] || "").trim();
    if (!item || seen[item]) {
      continue;
    }
    seen[item] = true;
    out.push(item);
  }
  out.sort(function(a, b) {
    return a.localeCompare(b);
  });
  return out;
}

// ─── HEALTH CHECK FORM DROPDOWNS ──────────────────────────────
function getHealthFormBootstrap(token) {
  return withFormLocationOptions_(token, { locations: getActivityFormLocations(token) });
}

function getHealthOptionsForLocation(location, token) {
  var user = validateSessionToken_(token);
  location = assertLocationAccess_(user, location);
  if (!location) {
    return { horses: [], grooms: [] };
  }
  return {
    horses: getActivityHorsesForLocation_(location),
    grooms: getStockGroomsForLocation_(location)
  };
}

function getHorseDetailsForHealthCheck(horseName, token) {
  return getHorseDetailsForHealthCheck_(horseName, token);
}

// ─── VACCINATION FORM DROPDOWNS ───────────────────────────────
function getVaccinationFormBootstrap(token) {
  return withFormLocationOptions_(token, { locations: getActivityFormLocations(token) });
}

function getVaccinationHorsesForLocation(location, token) {
  var user = validateSessionToken_(token);
  location = assertLocationAccess_(user, location);
  if (!location) {
    return { horses: [] };
  }
  return { horses: getActivityHorsesForLocation_(location) };
}

function getShoeingHorsesForLocation(location, token) {
  var user = validateSessionToken_(token);
  location = assertLocationAccess_(user, location);
  if (!location) {
    return { horses: [] };
  }
  return { horses: getActivityHorsesForLocation_(location) };
}

function getHorsesDropdownForUser_(token) {
  var user = validateSessionToken_(token);
  if (isAdminRole_(user.role)) {
    return getHorsesList_().map(function(h) {
      return h.name;
    });
  }
  var loc = String(user.location || "").trim();
  if (!loc) {
    return [];
  }
  return getActivityHorsesForLocation_(loc).map(function(h) {
    return h.value;
  });
}

function withFormLocationOptions_(token, options) {
  var user = validateSessionToken_(token);
  options = options || {};
  if (!isAdminRole_(user.role)) {
    var loc = resolveLocationName_(user.location || user.locationId);
    options.locationLocked = true;
    options.defaultLocation = loc;
    options.Location = loc ? [loc] : [];
  } else {
    options.locationLocked = false;
  }
  if (options.locations !== undefined || options.Location) {
    options.locations = options.Location || options.locations || [];
  }
  return options;
}

function getDropdownOptions(formKey, token) {
  validateSessionToken_(token);
  formKey = String(formKey || "").toLowerCase();
  if (formKey === "activity") {
    return withFormLocationOptions_(token, { Location: getActivityFormLocations(token) });
  }
  if (formKey === "stock") {
    var boot = getStockFormBootstrap(token);
    return withFormLocationOptions_(token, {
      Location: boot.locations,
      Item: boot.items,
      "Entered By": []
    });
  }
  if (formKey === "health") {
    return withFormLocationOptions_(token, {
      Location: getActivityFormLocations(token),
      Horse: [],
      "Checked By": []
    });
  }
  if (formKey === "vaccination") {
    return withFormLocationOptions_(token, {
      Location: getActivityFormLocations(token),
      Horse: []
    });
  }
  if (formKey === "horse" || formKey === "trainer" || formKey === "groom") {
    return withFormLocationOptions_(token, { Location: getActivityFormLocations(token) });
  }
  if (formKey === "routine" || formKey === "activity") {
    return withFormLocationOptions_(token, { Location: getActivityFormLocations(token) });
  }
  if (formKey === "shoeing") {
    return withFormLocationOptions_(token, {
      Location: getActivityFormLocations(token)
    });
  }
  if (formKey === "medical") {
    return withFormLocationOptions_(token, { Horse: getHorsesDropdownForUser_(token) });
  }
  return withFormLocationOptions_(token, {});
}

function submitFormData(formKey, payload, token) {
  validateSessionToken_(token);
  formKey = String(formKey || "").toLowerCase();
  payload = enforcePayloadLocation_(payload || {}, token);
  if (formKey === "activity" || formKey === "routine") {
    return submitDailyRoutineBatch_(payload, token);
  }
  if (formKey === "stock") {
    return submitStockForm_(payload);
  }
  if (formKey === "health") {
    return submitHealthForm_(payload);
  }
  if (formKey === "vaccination") {
    return submitVaccinationForm_(payload);
  }
  if (formKey === "shoeing") {
    return submitShoeingForm_(payload);
  }
  if (formKey === "medical") {
    return submitMedicalForm_(payload);
  }
  if (formKey === "horse") {
    if (String(payload._mode || "").toLowerCase() === "update") {
      return updateHorseForm_(payload);
    }
    return submitHorseForm_(payload);
  }
  if (formKey === "trainer") {
    if (String(payload._mode || "").toLowerCase() === "update") {
      return updateTrainerForm_(payload);
    }
    return submitTrainerForm_(payload);
  }
  if (formKey === "groom") {
    if (String(payload._mode || "").toLowerCase() === "update") {
      return updateGroomForm_(payload);
    }
    return submitGroomForm_(payload);
  }
  throw new Error("Form submit not implemented for: " + formKey);
}

var PROFILE_PHOTOS_FOLDER_NAME_ = "Kings Equestrian Profile Photos";

function getProfilePhotosFolder_() {
  var folders = DriveApp.getFoldersByName(PROFILE_PHOTOS_FOLDER_NAME_);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(PROFILE_PHOTOS_FOLDER_NAME_);
}

function buildProfilePhotoThumbnailUrl_(fileId) {
  fileId = String(fileId || "").trim();
  if (!fileId) {
    return "";
  }
  return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(fileId) + "&sz=w200";
}

function normalizePhotoResult_(photoResult) {
  if (!photoResult) {
    return { url: "", fileId: "" };
  }
  if (typeof photoResult === "string") {
    var fileIdFromUrl = extractDriveFileId_(photoResult);
    return {
      url: fileIdFromUrl ? buildProfilePhotoThumbnailUrl_(fileIdFromUrl) : photoResult,
      fileId: fileIdFromUrl
    };
  }
  var fileId = String(photoResult.fileId || "").trim() || extractDriveFileId_(photoResult.url);
  var url = String(photoResult.url || "").trim();
  if (fileId && (!url || url.indexOf("page=photo") >= 0)) {
    url = buildProfilePhotoThumbnailUrl_(fileId);
  }
  return { url: url, fileId: fileId };
}

function applyPhotoResultToPayload_(payload, photoResult, urlKey, fileIdKey) {
  payload = payload || {};
  var result = normalizePhotoResult_(photoResult);
  var urlField = urlKey || "Photo_URL";
  var fileField = fileIdKey || "Photo_File_ID";
  if (result.url) {
    payload[urlField] = result.url;
  }
  if (result.fileId) {
    payload[fileField] = result.fileId;
  }
  return result;
}

function saveNamedPhotoFromPayload_(payload, dataKey, urlKey, fileIdKey, entityType, entityId, entityName) {
  var stub = {
    Photo_Data: payload[dataKey] || "",
    Photo_URL: payload[urlKey] || "",
    Photo_File_ID: payload[fileIdKey] || ""
  };
  return saveProfilePhotoFromPayload_(stub, entityType, entityId, entityName);
}

function saveProfilePhotoFromPayload_(payload, entityType, entityId, entityName) {
  payload = payload || {};
  var dataUrl = String(payload.Photo_Data || payload.photo_data || "").trim();
  var existing = String(payload.Photo_URL || payload.Photo || payload.photo_url || "").trim();
  var existingFileId = String(payload.Photo_File_ID || payload.photo_file_id || "").trim() ||
    extractDriveFileId_(existing);
  if (!dataUrl) {
    return {
      url: existingFileId ? buildProfilePhotoThumbnailUrl_(existingFileId) : existing,
      fileId: existingFileId
    };
  }
  var mime = "image/jpeg";
  var b64 = dataUrl;
  var match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (match) {
    mime = match[1];
    b64 = match[2];
  }
  var bytes = Utilities.base64Decode(b64);
  var ext = mime.indexOf("png") >= 0 ? "png" : "jpg";
  var safeName = String(entityName || entityId || "profile").replace(/[^\w\-]+/g, "_").slice(0, 48);
  var fileName = String(entityType || "profile") + "_" + safeName + "_" +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyyMMdd_HHmmss") + "." + ext;
  var folder = getProfilePhotosFolder_();
  var file = folder.createFile(Utilities.newBlob(bytes, mime, fileName));
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareErr) {
    Logger.log("saveProfilePhotoFromPayload_ sharing: " + shareErr);
  }
  return {
    url: buildProfilePhotoThumbnailUrl_(file.getId()),
    fileId: file.getId()
  };
}

function getProfilePhotoDataUrl_(fileId) {
  fileId = String(fileId || "").trim();
  if (!fileId) {
    return "";
  }
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var bytes = blob.getBytes();
  if (bytes.length > 400000) {
    return buildProfilePhotoThumbnailUrl_(fileId);
  }
  return "data:" + (blob.getContentType() || "image/jpeg") + ";base64," + Utilities.base64Encode(bytes);
}

function getProfilePhotoDataUrl(fileId, token) {
  validateSessionToken_(token);
  return getProfilePhotoDataUrl_(fileId);
}

function getHorsePhotoDataUrls(fileIds, token) {
  validateSessionToken_(token);
  var out = {};
  var ids = fileIds || [];
  var i;
  for (i = 0; i < ids.length; i++) {
    var fid = String(ids[i] || "").trim();
    if (!fid || out[fid]) {
      continue;
    }
    try {
      out[fid] = getProfilePhotoDataUrl_(fid);
    } catch (e) {
      Logger.log("getHorsePhotoDataUrls " + fid + ": " + e);
    }
  }
  return out;
}

function appendRowByAliases_(sheet, fieldValues) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = getHeaderIndexMap_(sheet);
  var row = [];
  for (var c = 0; c < headers.length; c++) {
    row.push("");
  }
  for (var i = 0; i < fieldValues.length; i++) {
    var item = fieldValues[i];
    var col = findCol_(map, item.aliases);
    if (col >= 0 && item.value !== undefined && item.value !== null) {
      row[col] = item.value;
    }
  }
  sheet.appendRow(row);
  return { success: true };
}

function updateRowByAliases_(sheet, rowIndex1, fieldValues) {
  if (!sheet || rowIndex1 < 2) {
    return { success: false };
  }
  var map = getHeaderIndexMap_(sheet);
  for (var i = 0; i < fieldValues.length; i++) {
    var item = fieldValues[i];
    var col = findCol_(map, item.aliases);
    if (col >= 0 && item.value !== undefined) {
      sheet.getRange(rowIndex1, col + 1).setValue(item.value);
    }
  }
  return { success: true };
}

function findStaffDataRowIndex_(sheet, nameAliases, staffName) {
  if (!sheet || sheet.getLastRow() < 2) {
    return -1;
  }
  var data = sheet.getDataRange().getValues();
  var nameCol = resolveColumn_(sheet, nameAliases, 1);
  var key = String(staffName || "").trim().toLowerCase();
  if (!key) {
    return -1;
  }
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameCol] || "").trim().toLowerCase() === key) {
      return i + 1;
    }
  }
  return -1;
}

function readSheetCellByAliases_(row, map, aliases) {
  var col = findCol_(map, aliases);
  if (col < 0) {
    return "";
  }
  return row[col];
}

function formatSheetDateForEdit_(val) {
  if (!val) {
    return "";
  }
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd");
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return s.slice(0, 10);
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? "" : Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd");
}

function findHorseDataRowIndex_(sheet, horseId, horseName) {
  if (!sheet || sheet.getLastRow() < 2) {
    return -1;
  }
  var data = sheet.getDataRange().getValues();
  var map = getHeaderIndexMap_(sheet);
  var idCol = findCol_(map, ["Horse_ID", "Horse ID", "id"]);
  var nameCol = resolveColumn_(sheet, ["Horse_Name", "Horse Name", "Horse"], 1);
  var hid = String(horseId || "").trim().toLowerCase();
  var hname = String(horseName || "").trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (hid && idCol >= 0 && String(data[i][idCol] || "").trim().toLowerCase() === hid) {
      return i + 1;
    }
    if (hname && String(data[i][nameCol] || "").trim().toLowerCase() === hname) {
      return i + 1;
    }
  }
  return -1;
}

/** Feed times per day in Daily Routine (7AM, 1PM, 6PM, Extra). */
var ROUTINE_FEED_TIMES_PER_DAY_ = 4;

/** Daily feed (kg) = Standard per day × Facility Multiplier × Horse Weight ÷ 100 */
function computeFeedDailyKg_(standardQtyPerDay, facilityMultiplier, weightKg) {
  var std = parseFloat(standardQtyPerDay);
  var mult = parseFloat(facilityMultiplier);
  var w = parseFloat(weightKg);
  if (!isFinite(std) || std <= 0) {
    return "";
  }
  if (!isFinite(mult) || mult <= 0) {
    mult = 1;
  }
  if (!isFinite(w) || w <= 0) {
    return "";
  }
  return Math.round((std * mult * w / 100) * 100) / 100;
}

/** Per feed-time (kg) = daily amount ÷ feed times (standard qty is per day). */
function computeFeedPerSlotKg_(standardQtyPerDay, facilityMultiplier, weightKg) {
  var daily = computeFeedDailyKg_(standardQtyPerDay, facilityMultiplier, weightKg);
  if (daily !== "") {
    return Math.round((daily / ROUTINE_FEED_TIMES_PER_DAY_) * 100) / 100;
  }
  var std = parseFloat(standardQtyPerDay);
  if (isFinite(std) && std > 0) {
    return Math.round((std / ROUTINE_FEED_TIMES_PER_DAY_) * 100) / 100;
  }
  return "";
}

/** Per feed-time water (L) = daily standard ÷ feed times. */
function computeWaterPerSlotLiters_(dailyLiters) {
  var d = parseFloat(dailyLiters);
  if (!isFinite(d) || d <= 0) {
    return "";
  }
  return Math.round((d / ROUTINE_FEED_TIMES_PER_DAY_) * 100) / 100;
}

/** @deprecated use computeFeedPerSlotKg_ */
function computeFeedDefaultKg_(standardQty, facilityMultiplier, weightKg) {
  return computeFeedPerSlotKg_(standardQty, facilityMultiplier, weightKg);
}

function enrichHorsesFromProfiles_(horses) {
  if (!horses || !horses.length) {
    return horses;
  }
  ensureRoutineSheets_();
  var profileSheet = getSheetByNameSafe_(["HorseProfile"]);
  if (!profileSheet || profileSheet.getLastRow() < 2) {
    return horses;
  }
  var pdata = profileSheet.getDataRange().getValues();
  var pmap = getHeaderIndexMap_(profileSheet);
  var pName = resolveColumn_(profileSheet, ["Horse_Name", "Horse Name", "Horse"], 1);
  var pId = findCol_(pmap, ["Horse_ID", "Horse ID", "id"]);
  var pWeight = findCol_(pmap, ["Weight_Kg", "Weight (kg)", "Weight", "weight"]);
  var pMult = findCol_(pmap, ["Facility_Multiplier", "Facility Multiplier"]);
  var pLeaseR = findCol_(pmap, ["Lease_Rider", "Lease Rider"]);
  var pLeaseD = findCol_(pmap, ["Lease_Date", "Lease Date"]);
  var pMed = findCol_(pmap, ["Present_Medication", "Present Medication", "Current_Medication"]);
  var pTrainer = findCol_(pmap, ["Trainer", "trainer"]);
  var pGroom = findCol_(pmap, ["Groom", "groom"]);
  var pPhoto = findCol_(pmap, ["Photo_URL", "Photo", "Photo Url"]);
  var pPhotoId = findCol_(pmap, ["Photo_File_ID", "Photo File ID"]);
  var pChip = findCol_(pmap, ["Chip_No", "Chip No.", "Chip Number"]);
  var pEfi = findCol_(pmap, ["EFI_ID", "EFI ID"]);
  var byId = {};
  var byName = {};
  for (var pi = 1; pi < pdata.length; pi++) {
    var prow = pdata[pi];
    var pname = String(prow[pName] || "").trim();
    if (!pname) {
      continue;
    }
    var prof = {
      weight: pWeight >= 0 ? prow[pWeight] : "",
      facilityMultiplier: pMult >= 0 ? prow[pMult] : "",
      leaseRider: pLeaseR >= 0 ? String(prow[pLeaseR] || "").trim() : "",
      leaseDate: pLeaseD >= 0 ? prow[pLeaseD] : "",
      presentMedication: pMed >= 0 ? String(prow[pMed] || "").trim() : "",
      trainer: pTrainer >= 0 ? String(prow[pTrainer] || "").trim() : "",
      groom: pGroom >= 0 ? String(prow[pGroom] || "").trim() : "",
      chipNo: pChip >= 0 ? String(prow[pChip] || "").trim() : "",
      efiId: pEfi >= 0 ? String(prow[pEfi] || "").trim() : "",
      photoUrl: pPhoto >= 0 ? String(prow[pPhoto] || "").trim() : "",
      photoFileId: pPhotoId >= 0 ? String(prow[pPhotoId] || "").trim() : ""
    };
    if (pId >= 0 && String(prow[pId] || "").trim()) {
      byId[String(prow[pId] || "").trim().toLowerCase()] = prof;
    }
    byName[pname.toLowerCase()] = prof;
  }
  for (var hi = 0; hi < horses.length; hi++) {
    var h = horses[hi];
    var prof2 = (h.id && byId[String(h.id).trim().toLowerCase()]) || byName[String(h.name || "").trim().toLowerCase()];
    if (!prof2) {
      continue;
    }
    if (prof2.weight !== "" && prof2.weight !== undefined) {
      h.weight = prof2.weight;
    }
    if (prof2.facilityMultiplier !== "" && prof2.facilityMultiplier !== undefined) {
      h.facilityMultiplier = prof2.facilityMultiplier;
    }
    h.leaseRider = prof2.leaseRider || h.leaseRider || "";
    h.leaseDate = prof2.leaseDate || h.leaseDate || "";
    h.presentMedication = prof2.presentMedication || h.presentMedication || "";
    if (!h.trainer && prof2.trainer) {
      h.trainer = prof2.trainer;
    }
    if (!h.groom && prof2.groom) {
      h.groom = prof2.groom;
    }
    if (prof2.chipNo) {
      h.chipNo = prof2.chipNo;
    }
    if (prof2.efiId) {
      h.efiId = prof2.efiId;
    }
    if (prof2.photoFileId) {
      h.photoFileId = prof2.photoFileId;
    } else if (!h.photoFileId) {
      h.photoFileId = extractDriveFileId_(prof2.photoUrl || h.photoUrl || "");
    }
    if (prof2.photoUrl || prof2.photoFileId || h.photoUrl || h.photoFileId) {
      h.photoUrl = buildProfilePhotoDisplayUrl_(prof2.photoUrl || h.photoUrl, h.photoFileId || prof2.photoFileId);
    }
  }
  return horses;
}

function horseNameExists_(name) {
  var norm = String(name || "").trim().toLowerCase();
  if (!norm) {
    return false;
  }
  var horses = getHorsesList_();
  for (var i = 0; i < horses.length; i++) {
    if (String(horses[i].name || "").trim().toLowerCase() === norm) {
      return true;
    }
  }
  return false;
}

function trainerNameExists_(name) {
  var norm = String(name || "").trim().toLowerCase();
  if (!norm) {
    return false;
  }
  var trainers = getTrainersList_();
  for (var i = 0; i < trainers.length; i++) {
    if (String(trainers[i].name || "").trim().toLowerCase() === norm) {
      return true;
    }
  }
  return false;
}

function groomNameExists_(name) {
  var norm = String(name || "").trim().toLowerCase();
  if (!norm) {
    return false;
  }
  var grooms = getGroomsList_();
  for (var i = 0; i < grooms.length; i++) {
    if (String(grooms[i].name || "").trim().toLowerCase() === norm) {
      return true;
    }
  }
  return false;
}

function generateNextHorseId_(sheet) {
  var map = getHeaderIndexMap_(sheet);
  var idCol = findCol_(map, ["Horse_ID", "Horse ID", "id"]);
  if (idCol < 0) {
    return "";
  }
  var data = sheet.getDataRange().getValues();
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var raw = String(data[i][idCol] || "").trim();
    var match = raw.match(/(\d+)/);
    if (match) {
      var n = parseInt(match[1], 10);
      if (!isNaN(n) && n > maxNum) {
        maxNum = n;
      }
    }
  }
  var next = maxNum + 1;
  return "H" + (next < 1000 ? ("000" + next).slice(-3) : String(next));
}

function submitHorseForm_(payload) {
  var sheet = getSheetByNameSafe_(["Horses"]);
  if (!sheet) {
    throw new Error("Horses sheet not found.");
  }
  var name = String(payload.Horse_Name || payload["Horse Name"] || payload.Horse || "").trim();
  if (!name) {
    throw new Error("Horse name is required.");
  }
  var loc = String(payload.Location || "").trim();
  if (!loc) {
    throw new Error("Location is required.");
  }
  payload.Location = loc;
  if (horseNameExists_(name)) {
    throw new Error('A horse named "' + name + '" already exists.');
  }
  var horseId = String(payload.Horse_ID || payload["Horse ID"] || "").trim();
  if (!horseId) {
    horseId = generateNextHorseId_(sheet);
  }
  var status = String(payload.Status || "Active").trim() || "Active";
  ensureProfilePhotoColumns_();
  var photoResult = applyPhotoResultToPayload_(payload, saveProfilePhotoFromPayload_(payload, "horse", horseId, name));
  appendRowByAliases_(sheet, [
    { aliases: ["Horse_Name", "Horse Name", "Horse"], value: name },
    { aliases: ["Horse_ID", "Horse ID", "id"], value: horseId },
    { aliases: ["Breed", "breed"], value: String(payload.Breed || "").trim() },
    { aliases: ["Age", "age"], value: payload.Age || "" },
    { aliases: ["Location", "location"], value: loc },
    { aliases: ["Owner", "owner"], value: String(payload.Owner || "").trim() },
    { aliases: ["Status", "status"], value: status },
    { aliases: ["Gender", "gender"], value: String(payload.Gender || "").trim() },
    { aliases: ["Weight_Kg", "Weight (kg)", "Weight"], value: payload.Weight_Kg || payload.Weight || "" },
    { aliases: ["Lease_Rider", "Lease Rider"], value: String(payload.Lease_Rider || payload["Lease Rider"] || "").trim() },
    { aliases: ["Lease_Date", "Lease Date"], value: payload.Lease_Date || payload["Lease Date"] || "" },
    { aliases: ["Chip_No", "Chip No.", "Chip Number"], value: String(payload.Chip_No || payload["Chip No."] || "").trim() },
    { aliases: ["EFI_ID", "EFI ID"], value: String(payload.EFI_ID || payload["EFI ID"] || "").trim() },
    { aliases: ["Photo_URL", "Photo", "Photo Url"], value: photoResult.url || "" },
    { aliases: ["Photo_File_ID", "Photo File ID"], value: photoResult.fileId || "" }
  ]);
  try {
    ensureRoutineSheets_();
    var profile = getSS_().getSheetByName("HorseProfile");
    appendRowByAliases_(profile, getHorseProfileFieldValues_(horseId, name, loc, status, payload));
  } catch (eProf) {
    Logger.log("HorseProfile sync: " + eProf);
  }
  return { success: true, horseId: horseId, name: name };
}

function getHorseProfileFieldValues_(horseId, name, loc, status, payload) {
  payload = payload || {};
  return [
    { aliases: ["Horse_ID", "Horse ID"], value: horseId },
    { aliases: ["Horse_Name", "Horse Name"], value: name },
    { aliases: ["Location", "location"], value: loc },
    { aliases: ["Trainer", "trainer"], value: String(payload.Trainer || "").trim() },
    { aliases: ["Groom", "groom"], value: String(payload.Groom || "").trim() },
    { aliases: ["Status", "status"], value: status },
    { aliases: ["Breed", "breed"], value: String(payload.Breed || "").trim() },
    { aliases: ["Age", "age"], value: payload.Age || "" },
    { aliases: ["Gender", "gender"], value: String(payload.Gender || "").trim() },
    { aliases: ["Owner", "owner"], value: String(payload.Owner || "").trim() },
    { aliases: ["Weight_Kg", "Weight (kg)", "Weight"], value: payload.Weight_Kg || payload.Weight || "" },
    { aliases: ["Facility_Multiplier", "Facility Multiplier"], value: payload.Facility_Multiplier || payload["Facility Multiplier"] || "" },
    { aliases: ["Lease_Rider", "Lease Rider"], value: String(payload.Lease_Rider || payload["Lease Rider"] || "").trim() },
    { aliases: ["Lease_Date", "Lease Date"], value: payload.Lease_Date || payload["Lease Date"] || "" },
    { aliases: ["Present_Medication", "Present Medication", "Current_Medication"], value: String(payload.Present_Medication || payload["Present Medication"] || "").trim() },
    { aliases: ["Default_Wet_Grass"], value: payload.Default_Wet_Grass || payload["Default Wet Grass"] || "" },
    { aliases: ["Default_Dry_Grass"], value: payload.Default_Dry_Grass || "" },
    { aliases: ["Default_Feed_Mixed"], value: payload.Default_Feed_Mixed || "" },
    { aliases: ["Default_Barley"], value: payload.Default_Barley || "" },
    { aliases: ["Default_Oats"], value: payload.Default_Oats || "" },
    { aliases: ["Default_Water_Liters"], value: payload.Default_Water_Liters || "" },
    { aliases: ["Chip_No", "Chip No.", "Chip Number"], value: String(payload.Chip_No || payload["Chip No."] || "").trim() },
    { aliases: ["EFI_ID", "EFI ID"], value: String(payload.EFI_ID || payload["EFI ID"] || "").trim() },
    { aliases: ["Photo_URL", "Photo", "Photo Url"], value: String(payload.Photo_URL || payload.Photo || "").trim() },
    { aliases: ["Photo_File_ID", "Photo File ID"], value: String(payload.Photo_File_ID || payload.photo_file_id || "").trim() }
  ];
}

function getHorseForEdit(horseId, token) {
  var user = validateSessionToken_(token);
  horseId = String(horseId || "").trim();
  if (!horseId) {
    throw new Error("Horse ID is required.");
  }
  var horses = enrichHorsesFromProfiles_(getHorsesList_());
  var match = null;
  for (var i = 0; i < horses.length; i++) {
    if (String(horses[i].id || "").trim() === horseId) {
      match = horses[i];
      break;
    }
  }
  if (!match) {
    throw new Error("Horse not found: " + horseId);
  }
  assertLocationAccess_(user, match.location);
  var profileSheet = getSheetByNameSafe_(["HorseProfile"]);
  var defaults = {};
  if (profileSheet && profileSheet.getLastRow() > 1) {
    var pdata = profileSheet.getDataRange().getValues();
    var pmap = getHeaderIndexMap_(profileSheet);
    var pName = resolveColumn_(profileSheet, ["Horse_Name", "Horse Name", "Horse"], 1);
    var pId = findCol_(pmap, ["Horse_ID", "Horse ID", "id"]);
    var photoCol = findCol_(pmap, ["Photo_URL", "Photo", "Photo Url"]);
    var photoIdCol = findCol_(pmap, ["Photo_File_ID", "Photo File ID"]);
    var chipCol = findCol_(pmap, ["Chip_No", "Chip No.", "Chip Number"]);
    var efiCol = findCol_(pmap, ["EFI_ID", "EFI ID"]);
    var defs = {
      Default_Wet_Grass: findCol_(pmap, ["Default_Wet_Grass", "Default Wet Grass"]),
      Default_Dry_Grass: findCol_(pmap, ["Default_Dry_Grass", "Default Dry Grass"]),
      Default_Feed_Mixed: findCol_(pmap, ["Default_Feed_Mixed", "Default Feed Mixed"]),
      Default_Barley: findCol_(pmap, ["Default_Barley", "Default Barley"]),
      Default_Oats: findCol_(pmap, ["Default_Oats", "Default Oats"]),
      Default_Water_Liters: findCol_(pmap, ["Default_Water_Liters", "Default Water"])
    };
    for (var pi = 1; pi < pdata.length; pi++) {
      var prow = pdata[pi];
      var pid = pId >= 0 ? String(prow[pId] || "").trim() : "";
      var pname = String(prow[pName] || "").trim();
      if (pid !== horseId && pname.toLowerCase() !== String(match.name || "").trim().toLowerCase()) {
        continue;
      }
      var dk;
      for (dk in defs) {
        if (defs.hasOwnProperty(dk) && defs[dk] >= 0) {
          defaults[dk] = prow[defs[dk]];
        }
      }
      if (photoCol >= 0) {
        defaults.Photo_URL = prow[photoCol];
      }
      if (photoIdCol >= 0) {
        defaults.Photo_File_ID = prow[photoIdCol];
      }
      if (chipCol >= 0) {
        defaults.Chip_No = prow[chipCol];
      }
      if (efiCol >= 0) {
        defaults.EFI_ID = prow[efiCol];
      }
      break;
    }
  }
  return {
    horseId: match.id,
    Horse_Name: match.name,
    Chip_No: defaults.Chip_No || match.chipNo || "",
    EFI_ID: defaults.EFI_ID || match.efiId || "",
    Breed: match.breed,
    Age: match.age,
    Location: match.location,
    Owner: match.owner,
    Status: match.status,
    Gender: match.gender,
    Trainer: match.trainer || "",
    Groom: match.groom || "",
    Weight_Kg: match.weight || "",
    Facility_Multiplier: match.facilityMultiplier || "",
    Lease_Rider: match.leaseRider || "",
    Lease_Date: match.leaseDate || "",
    Present_Medication: match.presentMedication || "",
    Default_Wet_Grass: defaults.Default_Wet_Grass || "",
    Default_Dry_Grass: defaults.Default_Dry_Grass || "",
    Default_Feed_Mixed: defaults.Default_Feed_Mixed || "",
    Default_Barley: defaults.Default_Barley || "",
    Default_Oats: defaults.Default_Oats || "",
    Default_Water_Liters: defaults.Default_Water_Liters || "",
    Photo_URL: buildProfilePhotoDisplayUrl_(
      defaults.Photo_URL || match.photoUrl || "",
      defaults.Photo_File_ID || match.photoFileId || ""
    ),
    Photo_File_ID: defaults.Photo_File_ID || match.photoFileId || extractDriveFileId_(defaults.Photo_URL || match.photoUrl || "")
  };
}

function updateHorseForm_(payload) {
  var sheet = getSheetByNameSafe_(["Horses"]);
  if (!sheet) {
    throw new Error("Horses sheet not found.");
  }
  var horseId = String(payload.Horse_ID || payload["Horse ID"] || payload._horseId || "").trim();
  var name = String(payload.Horse_Name || payload["Horse Name"] || "").trim();
  if (!horseId && !name) {
    throw new Error("Horse ID or name is required to update.");
  }
  var rowIdx = findHorseDataRowIndex_(sheet, horseId, name);
  if (rowIdx < 0) {
    throw new Error("Horse not found in Horses sheet.");
  }
  var loc = String(payload.Location || "").trim();
  if (!loc) {
    throw new Error("Location is required.");
  }
  var newName = name;
  var dataCheck = sheet.getDataRange().getValues();
  var mapCheck = getHeaderIndexMap_(sheet);
  var idColCheck = findCol_(mapCheck, ["Horse_ID", "Horse ID", "id"]);
  var nameColCheck = resolveColumn_(sheet, ["Horse_Name", "Horse Name", "Horse"], 1);
  var existingName = String(dataCheck[rowIdx - 1][nameColCheck] || "").trim();
  if (newName && newName.toLowerCase() !== existingName.toLowerCase() && horseNameExists_(newName)) {
    throw new Error('A horse named "' + newName + '" already exists.');
  }
  if (!horseId && idColCheck >= 0) {
    horseId = String(dataCheck[rowIdx - 1][idColCheck] || "").trim();
  }
  if (!newName) {
    newName = existingName;
  }
  var status = String(payload.Status || "Active").trim() || "Active";
  ensureProfilePhotoColumns_();
  var photoResult = applyPhotoResultToPayload_(payload, saveProfilePhotoFromPayload_(payload, "horse", horseId, newName));
  updateRowByAliases_(sheet, rowIdx, [
    { aliases: ["Horse_Name", "Horse Name", "Horse"], value: newName },
    { aliases: ["Breed", "breed"], value: String(payload.Breed || "").trim() },
    { aliases: ["Age", "age"], value: payload.Age || "" },
    { aliases: ["Location", "location"], value: loc },
    { aliases: ["Owner", "owner"], value: String(payload.Owner || "").trim() },
    { aliases: ["Status", "status"], value: status },
    { aliases: ["Gender", "gender"], value: String(payload.Gender || "").trim() },
    { aliases: ["Weight_Kg", "Weight (kg)", "Weight"], value: payload.Weight_Kg || payload.Weight || "" },
    { aliases: ["Lease_Rider", "Lease Rider"], value: String(payload.Lease_Rider || payload["Lease Rider"] || "").trim() },
    { aliases: ["Lease_Date", "Lease Date"], value: payload.Lease_Date || payload["Lease Date"] || "" },
    { aliases: ["Chip_No", "Chip No.", "Chip Number"], value: String(payload.Chip_No || payload["Chip No."] || "").trim() },
    { aliases: ["EFI_ID", "EFI ID"], value: String(payload.EFI_ID || payload["EFI ID"] || "").trim() },
    { aliases: ["Photo_URL", "Photo", "Photo Url"], value: photoResult.url || String(payload.Photo_URL || "").trim() },
    { aliases: ["Photo_File_ID", "Photo File ID"], value: photoResult.fileId || String(payload.Photo_File_ID || "").trim() }
  ]);
  ensureRoutineSheets_();
  var profile = getSS_().getSheetByName("HorseProfile");
  var pRow = findHorseDataRowIndex_(profile, horseId, newName);
  var profVals = getHorseProfileFieldValues_(horseId, newName, loc, status, payload);
  if (pRow >= 0) {
    updateRowByAliases_(profile, pRow, profVals);
  } else {
    appendRowByAliases_(profile, profVals);
  }
  return { success: true, horseId: horseId, name: newName, updated: true };
}

function syncHorseProfileMedication_(horseId, horseName, medication) {
  var profile = getSheetByNameSafe_(["HorseProfile"]);
  if (!profile) {
    return;
  }
  var rowIdx = findHorseDataRowIndex_(profile, horseId, horseName);
  if (rowIdx < 0) {
    return;
  }
  updateRowByAliases_(profile, rowIdx, [
    { aliases: ["Present_Medication", "Present Medication", "Current_Medication"], value: String(medication || "").trim() }
  ]);
}

function submitTrainerForm_(payload) {
  var sheet = getSheetByNameSafe_(["Trainers"]);
  if (!sheet) {
    throw new Error("Trainers sheet not found.");
  }
  var name = String(payload.Trainer_Name || payload["Trainer Name"] || payload.Trainer || "").trim();
  if (!name) {
    throw new Error("Trainer name is required.");
  }
  var loc = String(payload.Location || "").trim();
  if (!loc) {
    throw new Error("Location is required.");
  }
  if (trainerNameExists_(name)) {
    throw new Error('A trainer named "' + name + '" already exists.');
  }
  ensureProfilePhotoColumns_();
  var trainerPhoto = saveProfilePhotoFromPayload_(payload, "trainer", name, name);
  var passportPhoto = saveNamedPhotoFromPayload_(
    payload, "Passport_Photo_Data", "Passport_Photo_URL", "Passport_Photo_File_ID", "trainer_passport", name, name
  );
  appendRowByAliases_(sheet, [
    { aliases: ["Trainer_Name", "Trainer Name", "Trainer"], value: name },
    { aliases: ["Specialization", "specialization"], value: String(payload.Specialization || "").trim() },
    { aliases: ["Location", "location"], value: loc },
    { aliases: ["Phone", "phone"], value: String(payload.Phone || "").trim() },
    { aliases: ["Experience_Years", "Experience Years", "Experience"], value: payload.Experience_Years || payload["Experience Years"] || "" },
    { aliases: ["Aadhar_Number", "Aadhar Number", "Aadhar"], value: String(payload.Aadhar_Number || payload["Aadhar Number"] || "").trim() },
    { aliases: ["Date_Of_Joining", "Date of Joining", "Joining_Date"], value: payload.Date_Of_Joining || payload["Date of Joining"] || "" },
    { aliases: ["Photo_URL", "Photo", "Photo Url"], value: trainerPhoto.url || "" },
    { aliases: ["Photo_File_ID", "Photo File ID"], value: trainerPhoto.fileId || "" },
    { aliases: ["Passport_Photo_URL", "Passport Photo URL", "Passport Photo"], value: passportPhoto.url || "" },
    { aliases: ["Passport_Photo_File_ID", "Passport Photo File ID"], value: passportPhoto.fileId || "" }
  ]);
  return { success: true, name: name };
}

function getTrainerForEdit(trainerName, token) {
  var user = validateSessionToken_(token);
  trainerName = String(trainerName || "").trim();
  if (!trainerName) {
    throw new Error("Trainer name is required.");
  }
  var sheet = getSheetByNameSafe_(["Trainers"]);
  if (!sheet) {
    throw new Error("Trainers sheet not found.");
  }
  var rowIdx = findStaffDataRowIndex_(sheet, ["Trainer_Name", "Trainer Name", "Trainer"], trainerName);
  if (rowIdx < 0) {
    throw new Error("Trainer not found: " + trainerName);
  }
  var data = sheet.getDataRange().getValues();
  var row = data[rowIdx - 1];
  var map = getHeaderIndexMap_(sheet);
  var location = normalizeSheetLocation_(readSheetCellByAliases_(row, map, ["Location", "location"]));
  assertLocationAccess_(user, location);
  var photoUrl = String(readSheetCellByAliases_(row, map, ["Photo_URL", "Photo", "Photo Url"]) || "").trim();
  var photoFileId = String(readSheetCellByAliases_(row, map, ["Photo_File_ID", "Photo File ID"]) || "").trim();
  var passportUrl = String(readSheetCellByAliases_(row, map, ["Passport_Photo_URL", "Passport Photo URL", "Passport Photo"]) || "").trim();
  var passportFileId = String(readSheetCellByAliases_(row, map, ["Passport_Photo_File_ID", "Passport Photo File ID"]) || "").trim();
  return {
    trainerName: trainerName,
    Trainer_Name: String(readSheetCellByAliases_(row, map, ["Trainer_Name", "Trainer Name", "Trainer"]) || "").trim(),
    Specialization: String(readSheetCellByAliases_(row, map, ["Specialization", "specialization"]) || "").trim(),
    Location: location,
    Phone: String(readSheetCellByAliases_(row, map, ["Phone", "phone"]) || "").trim(),
    Experience_Years: readSheetCellByAliases_(row, map, ["Experience_Years", "Experience Years", "Experience"]),
    Aadhar_Number: String(readSheetCellByAliases_(row, map, ["Aadhar_Number", "Aadhar Number", "Aadhar"]) || "").trim(),
    Date_Of_Joining: formatSheetDateForEdit_(readSheetCellByAliases_(row, map, ["Date_Of_Joining", "Date of Joining", "Joining_Date"])),
    Photo_URL: buildProfilePhotoDisplayUrl_(photoUrl, photoFileId),
    Photo_File_ID: photoFileId || extractDriveFileId_(photoUrl),
    Passport_Photo_URL: buildProfilePhotoDisplayUrl_(passportUrl, passportFileId),
    Passport_Photo_File_ID: passportFileId || extractDriveFileId_(passportUrl)
  };
}

function updateTrainerForm_(payload) {
  var sheet = getSheetByNameSafe_(["Trainers"]);
  if (!sheet) {
    throw new Error("Trainers sheet not found.");
  }
  var originalName = String(payload._originalName || payload.trainerName || "").trim();
  var newName = String(payload.Trainer_Name || payload["Trainer Name"] || payload.Trainer || "").trim();
  if (!originalName && !newName) {
    throw new Error("Trainer name is required to update.");
  }
  if (!newName) {
    newName = originalName;
  }
  var rowIdx = findStaffDataRowIndex_(sheet, ["Trainer_Name", "Trainer Name", "Trainer"], originalName || newName);
  if (rowIdx < 0) {
    throw new Error("Trainer not found in Trainers sheet.");
  }
  var loc = String(payload.Location || "").trim();
  if (!loc) {
    throw new Error("Location is required.");
  }
  if (newName.toLowerCase() !== String(originalName || newName).trim().toLowerCase() && trainerNameExists_(newName)) {
    throw new Error('A trainer named "' + newName + '" already exists.');
  }
  ensureProfilePhotoColumns_();
  var trainerPhoto = saveProfilePhotoFromPayload_(payload, "trainer", newName, newName);
  var passportPhoto = saveNamedPhotoFromPayload_(
    payload, "Passport_Photo_Data", "Passport_Photo_URL", "Passport_Photo_File_ID", "trainer_passport", newName, newName
  );
  updateRowByAliases_(sheet, rowIdx, [
    { aliases: ["Trainer_Name", "Trainer Name", "Trainer"], value: newName },
    { aliases: ["Specialization", "specialization"], value: String(payload.Specialization || "").trim() },
    { aliases: ["Location", "location"], value: loc },
    { aliases: ["Phone", "phone"], value: String(payload.Phone || "").trim() },
    { aliases: ["Experience_Years", "Experience Years", "Experience"], value: payload.Experience_Years || payload["Experience Years"] || "" },
    { aliases: ["Aadhar_Number", "Aadhar Number", "Aadhar"], value: String(payload.Aadhar_Number || payload["Aadhar Number"] || "").trim() },
    { aliases: ["Date_Of_Joining", "Date of Joining", "Joining_Date"], value: payload.Date_Of_Joining || payload["Date of Joining"] || "" },
    { aliases: ["Photo_URL", "Photo", "Photo Url"], value: trainerPhoto.url || String(payload.Photo_URL || "").trim() },
    { aliases: ["Photo_File_ID", "Photo File ID"], value: trainerPhoto.fileId || String(payload.Photo_File_ID || "").trim() },
    { aliases: ["Passport_Photo_URL", "Passport Photo URL", "Passport Photo"], value: passportPhoto.url || String(payload.Passport_Photo_URL || "").trim() },
    { aliases: ["Passport_Photo_File_ID", "Passport Photo File ID"], value: passportPhoto.fileId || String(payload.Passport_Photo_File_ID || "").trim() }
  ]);
  return { success: true, name: newName, updated: true };
}

function getGroomBankFieldValues_(payload, name, passbookPhoto) {
  passbookPhoto = passbookPhoto || {};
  return [
    { aliases: ["Bank_Name", "Bank Name"], value: String(payload.Bank_Name || payload["Bank Name"] || "").trim() },
    { aliases: ["Bank_Account_No", "Account No.", "Account Number", "Bank Account"], value: String(payload.Bank_Account_No || payload["Account No."] || payload["Account Number"] || "").trim() },
    { aliases: ["IFSC_Code", "IFSC", "IFSC Code"], value: String(payload.IFSC_Code || payload.IFSC || payload["IFSC Code"] || "").trim().toUpperCase() },
    { aliases: ["Passbook_Photo_URL", "Passbook Photo URL", "Passbook Photo"], value: passbookPhoto.url || String(payload.Passbook_Photo_URL || "").trim() },
    { aliases: ["Passbook_Photo_File_ID", "Passbook Photo File ID"], value: passbookPhoto.fileId || String(payload.Passbook_Photo_File_ID || "").trim() }
  ];
}

function submitGroomForm_(payload) {
  var sheet = getSheetByNameSafe_(["Grooms"]);
  if (!sheet) {
    throw new Error("Grooms sheet not found.");
  }
  var name = String(payload.Groom_Name || payload["Groom Name"] || payload.Groom || "").trim();
  if (!name) {
    throw new Error("Groom name is required.");
  }
  var loc = String(payload.Location || "").trim();
  if (!loc) {
    throw new Error("Location is required.");
  }
  if (groomNameExists_(name)) {
    throw new Error('A groom named "' + name + '" already exists.');
  }
  ensureProfilePhotoColumns_();
  var groomPhoto = saveProfilePhotoFromPayload_(payload, "groom", name, name);
  var passportPhoto = saveNamedPhotoFromPayload_(
    payload, "Passport_Photo_Data", "Passport_Photo_URL", "Passport_Photo_File_ID", "groom_passport", name, name
  );
  var passbookPhoto = saveNamedPhotoFromPayload_(
    payload, "Passbook_Photo_Data", "Passbook_Photo_URL", "Passbook_Photo_File_ID", "groom_passbook", name, name
  );
  appendRowByAliases_(sheet, [
    { aliases: ["Groom_Name", "Groom Name", "Groom"], value: name },
    { aliases: ["Location", "location"], value: loc },
    { aliases: ["Phone", "phone"], value: String(payload.Phone || "").trim() },
    { aliases: ["Shift", "shift"], value: String(payload.Shift || "").trim() },
    { aliases: ["Status", "status"], value: String(payload.Status || "Active").trim() || "Active" },
    { aliases: ["Aadhar_Number", "Aadhar Number", "Aadhar"], value: String(payload.Aadhar_Number || payload["Aadhar Number"] || "").trim() },
    { aliases: ["Date_Of_Joining", "Date of Joining", "Joining_Date"], value: payload.Date_Of_Joining || payload["Date of Joining"] || "" },
    { aliases: ["Photo_URL", "Photo", "Photo Url"], value: groomPhoto.url || "" },
    { aliases: ["Photo_File_ID", "Photo File ID"], value: groomPhoto.fileId || "" },
    { aliases: ["Passport_Photo_URL", "Passport Photo URL", "Passport Photo"], value: passportPhoto.url || "" },
    { aliases: ["Passport_Photo_File_ID", "Passport Photo File ID"], value: passportPhoto.fileId || "" }
  ].concat(getGroomBankFieldValues_(payload, name, passbookPhoto)));
  return { success: true, name: name };
}

function getGroomForEdit(groomName, token) {
  var user = validateSessionToken_(token);
  groomName = String(groomName || "").trim();
  if (!groomName) {
    throw new Error("Groom name is required.");
  }
  var sheet = getSheetByNameSafe_(["Grooms"]);
  if (!sheet) {
    throw new Error("Grooms sheet not found.");
  }
  var rowIdx = findStaffDataRowIndex_(sheet, ["Groom_Name", "Groom Name", "Groom"], groomName);
  if (rowIdx < 0) {
    throw new Error("Groom not found: " + groomName);
  }
  var data = sheet.getDataRange().getValues();
  var row = data[rowIdx - 1];
  var map = getHeaderIndexMap_(sheet);
  var location = normalizeSheetLocation_(readSheetCellByAliases_(row, map, ["Location", "location"]));
  assertLocationAccess_(user, location);
  var photoUrl = String(readSheetCellByAliases_(row, map, ["Photo_URL", "Photo", "Photo Url"]) || "").trim();
  var photoFileId = String(readSheetCellByAliases_(row, map, ["Photo_File_ID", "Photo File ID"]) || "").trim();
  var passportUrl = String(readSheetCellByAliases_(row, map, ["Passport_Photo_URL", "Passport Photo URL", "Passport Photo"]) || "").trim();
  var passportFileId = String(readSheetCellByAliases_(row, map, ["Passport_Photo_File_ID", "Passport Photo File ID"]) || "").trim();
  var passbookUrl = String(readSheetCellByAliases_(row, map, ["Passbook_Photo_URL", "Passbook Photo URL", "Passbook Photo"]) || "").trim();
  var passbookFileId = String(readSheetCellByAliases_(row, map, ["Passbook_Photo_File_ID", "Passbook Photo File ID"]) || "").trim();
  return {
    groomName: groomName,
    Groom_Name: String(readSheetCellByAliases_(row, map, ["Groom_Name", "Groom Name", "Groom"]) || "").trim(),
    Location: location,
    Phone: String(readSheetCellByAliases_(row, map, ["Phone", "phone"]) || "").trim(),
    Shift: String(readSheetCellByAliases_(row, map, ["Shift", "shift"]) || "").trim(),
    Status: String(readSheetCellByAliases_(row, map, ["Status", "status"]) || "").trim() || "Active",
    Aadhar_Number: String(readSheetCellByAliases_(row, map, ["Aadhar_Number", "Aadhar Number", "Aadhar"]) || "").trim(),
    Date_Of_Joining: formatSheetDateForEdit_(readSheetCellByAliases_(row, map, ["Date_Of_Joining", "Date of Joining", "Joining_Date"])),
    Bank_Name: String(readSheetCellByAliases_(row, map, ["Bank_Name", "Bank Name"]) || "").trim(),
    Bank_Account_No: String(readSheetCellByAliases_(row, map, ["Bank_Account_No", "Account No.", "Account Number", "Bank Account"]) || "").trim(),
    IFSC_Code: String(readSheetCellByAliases_(row, map, ["IFSC_Code", "IFSC", "IFSC Code"]) || "").trim(),
    Photo_URL: buildProfilePhotoDisplayUrl_(photoUrl, photoFileId),
    Photo_File_ID: photoFileId || extractDriveFileId_(photoUrl),
    Passport_Photo_URL: buildProfilePhotoDisplayUrl_(passportUrl, passportFileId),
    Passport_Photo_File_ID: passportFileId || extractDriveFileId_(passportUrl),
    Passbook_Photo_URL: buildProfilePhotoDisplayUrl_(passbookUrl, passbookFileId),
    Passbook_Photo_File_ID: passbookFileId || extractDriveFileId_(passbookUrl)
  };
}

function updateGroomForm_(payload) {
  var sheet = getSheetByNameSafe_(["Grooms"]);
  if (!sheet) {
    throw new Error("Grooms sheet not found.");
  }
  var originalName = String(payload._originalName || payload.groomName || "").trim();
  var newName = String(payload.Groom_Name || payload["Groom Name"] || payload.Groom || "").trim();
  if (!originalName && !newName) {
    throw new Error("Groom name is required to update.");
  }
  if (!newName) {
    newName = originalName;
  }
  var rowIdx = findStaffDataRowIndex_(sheet, ["Groom_Name", "Groom Name", "Groom"], originalName || newName);
  if (rowIdx < 0) {
    throw new Error("Groom not found in Grooms sheet.");
  }
  var loc = String(payload.Location || "").trim();
  if (!loc) {
    throw new Error("Location is required.");
  }
  if (newName.toLowerCase() !== String(originalName || newName).trim().toLowerCase() && groomNameExists_(newName)) {
    throw new Error('A groom named "' + newName + '" already exists.');
  }
  ensureProfilePhotoColumns_();
  var groomPhoto = saveProfilePhotoFromPayload_(payload, "groom", newName, newName);
  var passportPhoto = saveNamedPhotoFromPayload_(
    payload, "Passport_Photo_Data", "Passport_Photo_URL", "Passport_Photo_File_ID", "groom_passport", newName, newName
  );
  var passbookPhoto = saveNamedPhotoFromPayload_(
    payload, "Passbook_Photo_Data", "Passbook_Photo_URL", "Passbook_Photo_File_ID", "groom_passbook", newName, newName
  );
  updateRowByAliases_(sheet, rowIdx, [
    { aliases: ["Groom_Name", "Groom Name", "Groom"], value: newName },
    { aliases: ["Location", "location"], value: loc },
    { aliases: ["Phone", "phone"], value: String(payload.Phone || "").trim() },
    { aliases: ["Shift", "shift"], value: String(payload.Shift || "").trim() },
    { aliases: ["Status", "status"], value: String(payload.Status || "Active").trim() || "Active" },
    { aliases: ["Aadhar_Number", "Aadhar Number", "Aadhar"], value: String(payload.Aadhar_Number || payload["Aadhar Number"] || "").trim() },
    { aliases: ["Date_Of_Joining", "Date of Joining", "Joining_Date"], value: payload.Date_Of_Joining || payload["Date of Joining"] || "" },
    { aliases: ["Photo_URL", "Photo", "Photo Url"], value: groomPhoto.url || String(payload.Photo_URL || "").trim() },
    { aliases: ["Photo_File_ID", "Photo File ID"], value: groomPhoto.fileId || String(payload.Photo_File_ID || "").trim() },
    { aliases: ["Passport_Photo_URL", "Passport Photo URL", "Passport Photo"], value: passportPhoto.url || String(payload.Passport_Photo_URL || "").trim() },
    { aliases: ["Passport_Photo_File_ID", "Passport Photo File ID"], value: passportPhoto.fileId || String(payload.Passport_Photo_File_ID || "").trim() }
  ].concat(getGroomBankFieldValues_(payload, newName, passbookPhoto)));
  return { success: true, name: newName, updated: true };
}

function submitShoeingForm_(payload) {
  var sheet = getSS_().getSheetByName("ShoeingData");
  if (!sheet) {
    throw new Error("ShoeingData sheet not found.");
  }
  var loc = String(payload.Location || payload.location || "").trim();
  if (!loc) {
    throw new Error("Location is required.");
  }
  ensureSheetColumn_(sheet, "Location", ["Location", "location"]);
  payload.Location = loc;
  if (payload.NextDue && !payload["Next Due"]) {
    payload["Next Due"] = payload.NextDue;
  }
  return appendRowByHeaders_(sheet, payload);
}

function submitMedicalForm_(payload) {
  var sheet = getSS_().getSheetByName("MedicalHistory");
  if (!sheet) {
    throw new Error("MedicalHistory sheet not found.");
  }
  return appendRowByHeaders_(sheet, payload);
}

function appendRowByHeaders_(sheet, payload) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var c = 0; c < headers.length; c++) {
    var header = String(headers[c] || "").trim();
    row.push(payload.hasOwnProperty(header) ? payload[header] : "");
  }
  sheet.appendRow(row);
  return { success: true };
}

function submitVaccinationForm_(payload) {
  var sheet = getSS_().getSheetByName("VaccinationData");
  if (!sheet) {
    throw new Error("VaccinationData sheet not found.");
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var c = 0; c < headers.length; c++) {
    var header = String(headers[c] || "").trim();
    row.push(payload.hasOwnProperty(header) ? payload[header] : "");
  }
  sheet.appendRow(row);
  return { success: true };
}

function submitHealthForm_(payload) {
  var sheet = ensureHorseHealthSheet_();
  payload = normalizeHealthCheckPayload_(payload);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  var c;
  for (c = 0; c < headers.length; c++) {
    var header = String(headers[c] || "").trim();
    row.push(payload.hasOwnProperty(header) ? payload[header] : "");
  }
  sheet.appendRow(row);
  return { success: true, status: payload.Status };
}

function submitStockForm_(payload) {
  var sheet = getSheetByNameSafe_(["Raw Data", "RawData"]);
  if (!sheet) {
    throw new Error("Raw Data sheet not found.");
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var c = 0; c < headers.length; c++) {
    var header = String(headers[c] || "").trim();
    row.push(payload.hasOwnProperty(header) ? payload[header] : "");
  }
  sheet.appendRow(row);
  try {
    var stockLoc = payload.Location || payload.location;
    var stockItem = payload.Item || payload.item;
    var stockDate = payload.Date || payload.date || payload.Timestamp;
    if (stockLoc && stockItem) {
      var stub = {};
      stub[String(stockItem).trim()] = 1;
      updateDailyInventoryForConsumption_(stockLoc, stub, stockDate);
    }
  } catch (e) {
    Logger.log("incremental inventory after stock submit: " + e);
  }
  return { success: true };
}

function submitActivityForm_(payload) {
  var sheet = getSS_().getSheetByName("ActivityData");
  if (!sheet) {
    throw new Error("ActivityData sheet not found.");
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var c = 0; c < headers.length; c++) {
    var header = String(headers[c] || "").trim();
    row.push(payload.hasOwnProperty(header) ? payload[header] : "");
  }
  sheet.appendRow(row);
  return { success: true };
}

// ─── SET UP TRIGGERS ──────────────────────────────────────────
function setupTriggers() {
  // Delete existing triggers
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  // Daily inventory at 9 AM
  ScriptApp.newTrigger("generateInventory").timeBased().everyDays(1).atHour(9).create();
  // Daily email at 8 PM
  ScriptApp.newTrigger("dailyEmailReport").timeBased().everyDays(1).atHour(20).create();
}



// ─── DIAGNOSTIC TESTS — run each from Apps Script editor ──────

function diagSheetNames() {
  var ss = getSS_();
  var sheets = ss.getSheets();
  Logger.log("Spreadsheet: " + ss.getName());
  Logger.log("Total sheets: " + sheets.length);
  sheets.forEach(function(s) {
    Logger.log("  Sheet: [" + s.getName() + "] rows=" + s.getLastRow());
  });
}

// ─── DAILY ROUTINE LOG & HORSE PROFILE (readme ERP) ───────────
var ROUTINE_FEED_ITEMS_ = [
  { key: "wet_grass", label: "Wet Grass", inventory: ["Wet Grass", "Wet grass"] },
  { key: "dry_grass", label: "Dry Grass", inventory: ["Dry Grass", "Dry grass"] },
  { key: "feed_mixed", label: "Feed (Mixed)", inventory: ["Feed (Mixed)", "Feed Mixed", "Feed"] },
  { key: "barley", label: "Barley", inventory: ["Barley"] },
  { key: "oats", label: "Oats", inventory: ["Oats"] }
];
var ROUTINE_FEED_SLOTS_ = ["7AM", "1PM", "6PM", "Extra"];

/** Dropdown labels → minutes for reports and ActivityData sync. */
var ROUTINE_DURATION_MINUTES_MAP_ = {
  "15": 15,
  "30": 30,
  "45": 45,
  "1hr": 60,
  "1.50": 90,
  "1.5": 90,
  "2hr": 120
};

function parseRoutineDurationToMinutes_(val) {
  var s = String(val || "").trim();
  if (!s) {
    return 0;
  }
  if (ROUTINE_DURATION_MINUTES_MAP_.hasOwnProperty(s)) {
    return ROUTINE_DURATION_MINUTES_MAP_[s];
  }
  var lower = s.toLowerCase();
  if (ROUTINE_DURATION_MINUTES_MAP_.hasOwnProperty(lower)) {
    return ROUTINE_DURATION_MINUTES_MAP_[lower];
  }
  if (lower.indexOf("hr") >= 0) {
    var hours = parseFloat(lower.replace(/hr/g, "").trim());
    return isNaN(hours) ? 0 : Math.round(hours * 60);
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : Math.round(n);
}

function formatRoutineDurationLabel_(minutes) {
  var m = Math.round(parseFloat(minutes) || 0);
  if (m <= 0) {
    return "";
  }
  if (m === 60) {
    return "1hr";
  }
  if (m === 90) {
    return "1.50";
  }
  if (m === 120) {
    return "2hr";
  }
  return String(m);
}

function syncRoutineActivitiesToActivityData_(meta, horseName, horseRow) {
  var sheet = getSheetByNameSafe_(["ActivityData"]);
  if (!sheet) {
    return;
  }
  ensureSheetColumn_(sheet, "Duration_Label", ["Duration_Label", "Duration Label"]);
  var slots = [
    { activity: horseRow.activity1, duration: horseRow.duration1 },
    { activity: horseRow.activity2, duration: horseRow.duration2 }
  ];
  var si;
  for (si = 0; si < slots.length; si++) {
    var slot = slots[si];
    var act = String(slot.activity || "").trim();
    if (!act) {
      continue;
    }
    var label = String(slot.duration || "").trim();
    var mins = parseRoutineDurationToMinutes_(label);
    appendRowByAliases_(sheet, [
      { aliases: ["Date", "date"], value: meta.date },
      { aliases: ["Location", "location"], value: meta.location },
      { aliases: ["Horse", "Horse_Name", "Horse Name"], value: horseName },
      { aliases: ["Activity", "activity"], value: act },
      { aliases: ["Duration", "duration"], value: mins },
      { aliases: ["Duration_Label", "Duration Label"], value: label || formatRoutineDurationLabel_(mins) },
      { aliases: ["Trainer", "Trainer_Name", "Trainer Name"], value: meta.trainer || "" },
      { aliases: ["Groom", "Groom_Name", "Groom Name"], value: meta.groom || "" }
    ]);
  }
}

function getRoutineLogHeaders_() {
  var headers = [
    "Date", "Location", "Horse_ID", "Horse_Name", "Trainer", "Groom",
    "Activity_1", "Duration_1", "Activity_2", "Duration_2"
  ];
  var fi;
  for (fi = 0; fi < ROUTINE_FEED_ITEMS_.length; fi++) {
    var item = ROUTINE_FEED_ITEMS_[fi];
    var si;
    for (si = 0; si < ROUTINE_FEED_SLOTS_.length; si++) {
      headers.push(item.label + "_" + ROUTINE_FEED_SLOTS_[si]);
    }
    headers.push(item.label + "_Total");
  }
  var wi;
  for (wi = 0; wi < ROUTINE_FEED_SLOTS_.length; wi++) {
    headers.push("Water_" + ROUTINE_FEED_SLOTS_[wi]);
  }
  headers.push("Water_Total", "Present_Medication", "Notes", "Logged_By");
  return headers;
}

function ensureSheetHeaders_(sheetName, headers) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function ensureSheetColumn_(sheet, headerName, aliases) {
  if (!sheet || !headerName) {
    return;
  }
  var map = getHeaderIndexMap_(sheet);
  var list = aliases || [headerName];
  if (findCol_(map, list) >= 0) {
    return;
  }
  var col = sheet.getLastColumn() + 1;
  if (col < 1) {
    col = 1;
  }
  sheet.getRange(1, col).setValue(headerName);
  clearSheetHeaderCache_(sheet);
}

function ensureStaffSheetColumns_() {
  var trainers = getSheetByNameSafe_(["Trainers"]);
  var grooms = getSheetByNameSafe_(["Grooms"]);
  var staffCols = [
    ["Aadhar_Number", ["Aadhar_Number", "Aadhar Number", "Aadhar"]],
    ["Date_Of_Joining", ["Date_Of_Joining", "Date of Joining", "Joining_Date"]],
    ["Passport_Photo_URL", ["Passport_Photo_URL", "Passport Photo URL", "Passport Photo"]],
    ["Passport_Photo_File_ID", ["Passport_Photo_File_ID", "Passport Photo File ID"]]
  ];
  var i;
  for (i = 0; i < staffCols.length; i++) {
    if (trainers) {
      ensureSheetColumn_(trainers, staffCols[i][0], staffCols[i][1]);
    }
    if (grooms) {
      ensureSheetColumn_(grooms, staffCols[i][0], staffCols[i][1]);
    }
  }
  var groomBankCols = [
    ["Bank_Name", ["Bank_Name", "Bank Name"]],
    ["Bank_Account_No", ["Bank_Account_No", "Account No.", "Account Number", "Bank Account"]],
    ["IFSC_Code", ["IFSC_Code", "IFSC", "IFSC Code"]],
    ["Passbook_Photo_URL", ["Passbook_Photo_URL", "Passbook Photo URL", "Passbook Photo"]],
    ["Passbook_Photo_File_ID", ["Passbook_Photo_File_ID", "Passbook Photo File ID"]]
  ];
  if (grooms) {
    for (i = 0; i < groomBankCols.length; i++) {
      ensureSheetColumn_(grooms, groomBankCols[i][0], groomBankCols[i][1]);
    }
  }
}

function ensureProfilePhotoColumns_() {
  var horses = getSheetByNameSafe_(["Horses"]);
  var trainers = getSheetByNameSafe_(["Trainers"]);
  var grooms = getSheetByNameSafe_(["Grooms"]);
  var profile = getSheetByNameSafe_(["HorseProfile"]);
  ensureSheetColumn_(horses, "Photo_URL", ["Photo_URL", "Photo", "Photo Url"]);
  ensureSheetColumn_(horses, "Photo_File_ID", ["Photo_File_ID", "Photo File ID"]);
  ensureSheetColumn_(trainers, "Photo_URL", ["Photo_URL", "Photo", "Photo Url"]);
  ensureSheetColumn_(trainers, "Photo_File_ID", ["Photo_File_ID", "Photo File ID"]);
  ensureSheetColumn_(grooms, "Photo_URL", ["Photo_URL", "Photo", "Photo Url"]);
  ensureSheetColumn_(grooms, "Photo_File_ID", ["Photo_File_ID", "Photo File ID"]);
  ensureSheetColumn_(profile, "Photo_URL", ["Photo_URL", "Photo", "Photo Url"]);
  ensureSheetColumn_(profile, "Photo_File_ID", ["Photo_File_ID", "Photo File ID"]);
  ensureSheetColumn_(horses, "Chip_No", ["Chip_No", "Chip No.", "Chip Number"]);
  ensureSheetColumn_(horses, "EFI_ID", ["EFI_ID", "EFI ID"]);
  ensureSheetColumn_(profile, "Chip_No", ["Chip_No", "Chip No.", "Chip Number"]);
  ensureSheetColumn_(profile, "EFI_ID", ["EFI_ID", "EFI ID"]);
  ensureStaffSheetColumns_();
  ensureRoutineSheets_();
}

function extractDriveFileId_(url) {
  var s = String(url || "").trim();
  if (!s) {
    return "";
  }
  if (s.indexOf("page=photo") >= 0) {
    var proxyMatch = s.match(/[?&]fileId=([a-zA-Z0-9_-]+)/i);
    if (proxyMatch) {
      return proxyMatch[1];
    }
  }
  var driveMatch = s.match(/(?:[?&]id=|\/d\/)([a-zA-Z0-9_-]+)/);
  return driveMatch ? driveMatch[1] : "";
}

function buildProfilePhotoDisplayUrl_(storedUrl, fileId) {
  var fid = String(fileId || "").trim() || extractDriveFileId_(storedUrl);
  if (fid) {
    return buildProfilePhotoThumbnailUrl_(fid);
  }
  return String(storedUrl || "").trim();
}

function serveProfilePhoto_(e) {
  var fileId = String((e && e.parameter && (e.parameter.fileId || e.parameter.fileid)) || "").trim();
  if (!fileId) {
    return ContentService.createTextOutput("Photo not found").setMimeType(ContentService.MimeType.TEXT);
  }
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return ContentService.create(blob).setMimeType(blob.getContentType());
  } catch (err) {
    Logger.log("serveProfilePhoto_: " + err);
    return ContentService.createTextOutput("Photo unavailable").setMimeType(ContentService.MimeType.TEXT);
  }
}

function ensureRoutineSheets_() {
  ensureSheetHeaders_("DailyRoutineLog", getRoutineLogHeaders_());
  ensureSheetHeaders_("HorseProfile", [
    "Horse_ID", "Horse_Name", "Location", "Trainer", "Groom", "Status", "Breed", "Age", "Gender", "Owner",
    "Weight_Kg", "Facility_Multiplier", "Lease_Rider", "Lease_Date", "Present_Medication",
    "Default_Wet_Grass", "Default_Dry_Grass", "Default_Feed_Mixed", "Default_Barley", "Default_Oats",
    "Default_Water_Liters", "Chip_No", "EFI_ID", "Photo_URL", "Photo_File_ID"
  ]);
  ensureSheetHeaders_("inventory_transaction", [
    "Timestamp", "Location", "Item", "Quantity", "Type", "Reference"
  ]);
}

function getHorseProfilesForLocation_(location) {
  var horses = getHorsesList_();
  var profileSheet = getSheetByNameSafe_(["HorseProfile"]);
  var profileMap = {};
  if (profileSheet && profileSheet.getLastRow() > 1) {
    var pdata = profileSheet.getDataRange().getValues();
    var pmap = getHeaderIndexMap_(profileSheet);
    var pName = resolveColumn_(profileSheet, ["Horse_Name", "Horse Name", "Horse"], 1);
    var pId = findCol_(pmap, ["Horse_ID", "Horse ID", "id"]);
    var pLoc = findCol_(pmap, ["Location", "location"]);
    var pTrainer = findCol_(pmap, ["Trainer", "trainer"]);
    var pGroom = findCol_(pmap, ["Groom", "groom"]);
    var pStatus = findCol_(pmap, ["Status", "status"]);
    var pWeight = findCol_(pmap, ["Weight_Kg", "Weight (kg)", "Weight", "weight"]);
    var pMult = findCol_(pmap, ["Facility_Multiplier", "Facility Multiplier"]);
    var pMed = findCol_(pmap, ["Present_Medication", "Present Medication", "Current_Medication"]);
    var defs = {
      wet_grass: findCol_(pmap, ["Default_Wet_Grass", "Default Wet Grass"]),
      dry_grass: findCol_(pmap, ["Default_Dry_Grass", "Default Dry Grass"]),
      feed_mixed: findCol_(pmap, ["Default_Feed_Mixed", "Default Feed Mixed"]),
      barley: findCol_(pmap, ["Default_Barley", "Default Barley"]),
      oats: findCol_(pmap, ["Default_Oats", "Default Oats"]),
      water: findCol_(pmap, ["Default_Water_Liters", "Default Water"])
    };
    for (var pi = 1; pi < pdata.length; pi++) {
      var prow = pdata[pi];
      var pname = String(prow[pName] || "").trim();
      if (!pname) continue;
      var weightVal = pWeight >= 0 ? prow[pWeight] : "";
      var multVal = pMult >= 0 ? prow[pMult] : "";
      var stdDefs = {
        wet_grass: defs.wet_grass >= 0 ? prow[defs.wet_grass] : "",
        dry_grass: defs.dry_grass >= 0 ? prow[defs.dry_grass] : "",
        feed_mixed: defs.feed_mixed >= 0 ? prow[defs.feed_mixed] : "",
        barley: defs.barley >= 0 ? prow[defs.barley] : "",
        oats: defs.oats >= 0 ? prow[defs.oats] : "",
        water: defs.water >= 0 ? prow[defs.water] : ""
      };
      var calcDefs = {};
      var dk;
      for (dk in stdDefs) {
        if (!stdDefs.hasOwnProperty(dk)) {
          continue;
        }
        if (dk === "water") {
          calcDefs[dk] = computeWaterPerSlotLiters_(stdDefs[dk]);
        } else {
          calcDefs[dk] = computeFeedPerSlotKg_(stdDefs[dk], multVal, weightVal);
        }
      }
      profileMap[pname.toLowerCase()] = {
        id: pId >= 0 ? String(prow[pId] || "").trim() : "",
        name: pname,
        location: pLoc >= 0 ? String(prow[pLoc] || "").trim() : "",
        trainer: pTrainer >= 0 ? String(prow[pTrainer] || "").trim() : "",
        groom: pGroom >= 0 ? String(prow[pGroom] || "").trim() : "",
        status: pStatus >= 0 ? String(prow[pStatus] || "").trim() : "",
        weight: weightVal,
        facilityMultiplier: multVal,
        presentMedication: pMed >= 0 ? String(prow[pMed] || "").trim() : "",
        standardDefaults: stdDefs,
        defaults: calcDefs
      };
    }
  }
  var out = [];
  for (var hi = 0; hi < horses.length; hi++) {
    var h = horses[hi];
    if (location && !matchesLocationFilter_(h.location, location)) continue;
    var key = String(h.name || "").trim().toLowerCase();
    var prof = profileMap[key] || {};
    out.push({
      serial: hi + 1,
      id: prof.id || h.id || ("H" + String(hi + 1).padStart(3, "0")),
      name: h.name,
      location: h.location,
      breed: h.breed,
      status: prof.status || h.status || "Active",
      trainer: prof.trainer || "",
      groom: prof.groom || "",
      weight: prof.weight || "",
      facilityMultiplier: prof.facilityMultiplier || "",
      presentMedication: prof.presentMedication || "",
      standardDefaults: prof.standardDefaults || {},
      defaults: prof.defaults || {
        wet_grass: "", dry_grass: "", feed_mixed: "", barley: "", oats: "", water: ""
      }
    });
  }
  return out;
}

function sumSlotValues_(row, prefix) {
  var total = 0;
  var si;
  for (si = 0; si < ROUTINE_FEED_SLOTS_.length; si++) {
    total += parseFloat(row[prefix + "_" + ROUTINE_FEED_SLOTS_[si]]) || 0;
  }
  return total;
}

function buildRoutineRowObject_(meta, horseRow) {
  var out = {
    Date: meta.date,
    Location: meta.location,
    Horse_ID: horseRow.horseId || horseRow.id,
    Horse_Name: horseRow.horseName || horseRow.name,
    Trainer: horseRow.trainer || meta.trainer || "",
    Groom: horseRow.groom || meta.groom || "",
    Activity_1: horseRow.activity1 || "",
    Duration_1: horseRow.duration1 || "",
    Activity_2: horseRow.activity2 || "",
    Duration_2: horseRow.duration2 || "",
    Present_Medication: horseRow.presentMedication || horseRow.present_medication || "",
    Notes: horseRow.notes || "",
    Logged_By: meta.loggedBy || ""
  };
  var fi;
  for (fi = 0; fi < ROUTINE_FEED_ITEMS_.length; fi++) {
    var item = ROUTINE_FEED_ITEMS_[fi];
    var si;
    var feedTotal = 0;
    for (si = 0; si < ROUTINE_FEED_SLOTS_.length; si++) {
      var slot = ROUTINE_FEED_SLOTS_[si];
      var val = parseFloat(horseRow[item.key + "_" + slot] || horseRow[item.key + "_" + slot.toLowerCase()] || 0) || 0;
      out[item.label + "_" + slot] = val;
      feedTotal += val;
    }
    out[item.label + "_Total"] = feedTotal;
  }
  var waterTotal = 0;
  for (si = 0; si < ROUTINE_FEED_SLOTS_.length; si++) {
    var wslot = ROUTINE_FEED_SLOTS_[si];
    var wv = parseFloat(horseRow["water_" + wslot] || horseRow["water_" + wslot.toLowerCase()] || 0) || 0;
    out["Water_" + wslot] = wv;
    waterTotal += wv;
  }
  out.Water_Total = waterTotal;
  return out;
}

function getDailyRoutineBootstrap(token, location, dateStr) {
  var user = validateSessionToken_(token);
  location = assertLocationAccess_(user, location);
  if (!location) {
    throw new Error("Location is required.");
  }
  ensureRoutineSheets_();
  var assign = getUserHorseAssignmentMode_(user);
  var allAtLoc = getHorseProfilesForLocation_(location);
  var horses = filterHorseProfilesForUser_(allAtLoc, user);
  var trainers = getActivityTrainersForLocation_(location);
  var defaultTrainer = "";
  if (assign.mode === "trainer") {
    defaultTrainer = assign.staffName;
    if (trainers.indexOf(defaultTrainer) < 0 && defaultTrainer) {
      trainers = [defaultTrainer].concat(trainers);
    }
  }
  return {
    date: dateStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd"),
    location: location,
    horses: horses,
    horsesAtLocation: allAtLoc.length,
    trainers: trainers,
    assignMode: assign.mode,
    staffName: assign.staffName,
    roleLabel: assign.roleLabel,
    defaultTrainer: defaultTrainer,
    feedItems: ROUTINE_FEED_ITEMS_.map(function(x) { return x.label; }),
    feedSlots: ROUTINE_FEED_SLOTS_,
    activities: ["Flatwork", "Pole work", "Lunging", "Dressage", "Jumping", "Horse safari", "Photography", "Training", "Exercise", "Rest", "Grooming", "Turnout"],
    durationOptions: ["15", "30", "45", "1hr", "1.50", "2hr"],
    feedKgOptions: ["1", "1.5", "2", "2.5", "3", "4"],
    waterLtrOptions: ["10", "20", "30", "40"],
    feedFormulaHint: "Per feed time (kg) = (Standard per day × Facility Multiplier × Weight ÷ 100) ÷ " +
      ROUTINE_FEED_TIMES_PER_DAY_ + " feed times. Water (L) per time = daily standard ÷ " + ROUTINE_FEED_TIMES_PER_DAY_
  };
}

function submitDailyRoutineBatch_(payload, token) {
  var user = validateSessionToken_(token);
  payload = payload || {};
  var location = assertLocationAccess_(user, payload.location || payload.Location);
  if (!location) {
    throw new Error("Location is required.");
  }
  var dateVal = payload.date || payload.Date || new Date();
  var rows = payload.rows || [];
  if (!rows.length) {
    throw new Error("No horse rows to save.");
  }
  ensureRoutineSheets_();
  var sheet = getSS_().getSheetByName("DailyRoutineLog");
  var headers = getRoutineLogHeaders_();
  var consumption = {};
  var saved = 0;
  var meta = {
    date: dateVal,
    location: location,
    trainer: payload.trainer || payload.Trainer || "",
    groom: payload.groom || payload.Groom || "",
    loggedBy: user.name || user.email || ""
  };
  var ri;
  for (ri = 0; ri < rows.length; ri++) {
    var horseRow = rows[ri];
    if (!horseRow || !(horseRow.horseName || horseRow.name)) continue;
    var horseName = String(horseRow.horseName || horseRow.name || "").trim();
    assertHorseAccessForUser_(horseName, location, user);
    var hasData = horseRow.activity1 || horseRow.activity2 || horseRow.notes ||
      horseRow.presentMedication || horseRow.present_medication;
    if (!hasData) {
      var fk;
      for (fk = 0; fk < ROUTINE_FEED_ITEMS_.length; fk++) {
        var ik = ROUTINE_FEED_ITEMS_[fk].key;
        var sk;
        for (sk = 0; sk < ROUTINE_FEED_SLOTS_.length; sk++) {
          if (parseFloat(horseRow[ik + "_" + ROUTINE_FEED_SLOTS_[sk]] || 0) > 0) {
            hasData = true;
            break;
          }
        }
        if (hasData) break;
      }
    }
    if (!hasData) continue;
    var rowObj = buildRoutineRowObject_(meta, horseRow);
    var rowArr = [];
    var hi;
    for (hi = 0; hi < headers.length; hi++) {
      rowArr.push(rowObj[headers[hi]] !== undefined ? rowObj[headers[hi]] : "");
    }
    sheet.appendRow(rowArr);
    saved++;
    try {
      syncRoutineActivitiesToActivityData_(meta, horseName, horseRow);
    } catch (eAct) {
      Logger.log("syncRoutineActivitiesToActivityData_: " + eAct);
    }
    var medNote = String(horseRow.presentMedication || horseRow.present_medication || "").trim();
    if (medNote) {
      try {
        syncHorseProfileMedication_(horseRow.horseId || horseRow.id, horseName, medNote);
      } catch (eMed) {
        Logger.log("syncHorseProfileMedication_: " + eMed);
      }
    }
    var fj;
    for (fj = 0; fj < ROUTINE_FEED_ITEMS_.length; fj++) {
      var fitem = ROUTINE_FEED_ITEMS_[fj];
      var ft = parseFloat(rowObj[fitem.label + "_Total"]) || 0;
      if (ft > 0) {
        consumption[fitem.label] = (consumption[fitem.label] || 0) + ft;
      }
    }
  }
  if (!saved) {
    throw new Error("Enter activity, feed, or water for at least one horse.");
  }
  applyInventoryConsumptionFromRoutine_(location, consumption,
    "DailyRoutineLog " + Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd"),
    dateVal);
  var invResult = { updated: 0, appended: 0 };
  try {
    invResult = updateDailyInventoryForConsumption_(location, consumption, dateVal);
  } catch (eInv) {
    Logger.log("updateDailyInventoryForConsumption_ after routine: " + eInv);
  }
  clearDashRequestCache_();
  return { success: true, saved: saved, consumption: consumption, inventory: invResult };
}

function applyInventoryConsumptionFromRoutine_(location, consumptionByItem, reference, logDate) {
  var txnSheet = ensureSheetHeaders_("inventory_transaction", [
    "Timestamp", "Location", "Item", "Quantity", "Type", "Reference"
  ]);
  var rawSheet = getSheetByNameSafe_(["Raw Data", "RawData"]);
  if (!rawSheet) {
    Logger.log("Raw Data sheet missing; skipping inventory deduction.");
    return;
  }
  var now = new Date();
  var entryDate = logDate ? new Date(logDate) : now;
  if (isNaN(entryDate.getTime())) {
    entryDate = now;
  }
  var itemName;
  for (itemName in consumptionByItem) {
    if (!consumptionByItem.hasOwnProperty(itemName)) continue;
    var qty = parseFloat(consumptionByItem[itemName]) || 0;
    if (qty <= 0) continue;
    appendRowByAliases_(rawSheet, [
      { aliases: ["Date", "date"], value: entryDate },
      { aliases: ["Location", "location"], value: location },
      { aliases: ["Item", "item"], value: itemName },
      { aliases: ["Entry Type", "Entry_Type", "entry_type"], value: "Consumption" },
      { aliases: ["Quantity", "quantity"], value: qty }
    ]);
    txnSheet.appendRow([now, location, itemName, qty, "CONSUMPTION", reference || "DailyRoutineLog"]);
  }
}

function getRoutineConsumptionByItem_(locationFilter, days) {
  days = days || 7;
  var sheet = getSheetByNameSafe_(["DailyRoutineLog"]);
  if (!sheet || sheet.getLastRow() < 2) {
    return {};
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  var totals = {};
  var r;
  for (r = 1; r < data.length; r++) {
    var row = data[r];
    var rowDate = new Date(row[0]);
    if (isNaN(rowDate.getTime()) || rowDate < cutoff) continue;
    if (locationFilter && row[1] && !matchesLocationFilter_(row[1], locationFilter)) continue;
    var c;
    for (c = 0; c < headers.length; c++) {
      var h = String(headers[c] || "");
      if (h.indexOf("_Total") >= 0 && h.indexOf("Water") < 0) {
        var itemLabel = h.replace("_Total", "");
        var v = parseFloat(row[c]) || 0;
        if (v > 0) {
          totals[itemLabel] = (totals[itemLabel] || 0) + v;
        }
      }
    }
  }
  var avg = {};
  var dayCount = Math.max(1, days);
  for (var k in totals) {
    if (totals.hasOwnProperty(k)) {
      avg[k] = totals[k] / dayCount;
    }
  }
  return avg;
}

function getStockDaysRemaining_(locationFilter) {
  var invRows = [];
  try {
    invRows = getDailyInventoryRows_();
  } catch (e) {
    return [];
  }
  var avgConsumption = getRoutineConsumptionByItem_(locationFilter, 7);
  var out = [];
  var seen = {};
  var i;
  for (i = 0; i < invRows.length; i++) {
    var row = invRows[i];
    if (locationFilter && row.location && !matchesLocationFilter_(row.location, locationFilter)) continue;
    var item = String(row.item || "").trim();
    if (!item || seen[item + "|" + row.location]) continue;
    seen[item + "|" + row.location] = true;
    var stock = parseFloat(row.closing) || 0;
    var dailyUse = 0;
    var key;
    for (key in avgConsumption) {
      if (avgConsumption.hasOwnProperty(key) && key.toLowerCase() === item.toLowerCase()) {
        dailyUse = avgConsumption[key];
        break;
      }
    }
    if (!dailyUse) {
      dailyUse = parseFloat(row.consumption) || 0;
    }
    var daysLeft = dailyUse > 0 ? Math.floor(stock / dailyUse) : null;
    out.push({
      item: item,
      location: row.location,
      stock: stock,
      dailyUse: Math.round(dailyUse * 100) / 100,
      daysLeft: daysLeft,
      label: daysLeft !== null ? daysLeft + " days stock" : "—"
    });
  }
  out.sort(function(a, b) {
    var da = a.daysLeft === null ? 9999 : a.daysLeft;
    var db = b.daysLeft === null ? 9999 : b.daysLeft;
    return da - db;
  });
  return out.slice(0, 8);
}

function isRehabStatus_(status) {
  var s = String(status || "").toLowerCase();
  return s.indexOf("rehab") >= 0;
}

function isLeaveStatus_(status) {
  var s = String(status || "").toLowerCase();
  return s.indexOf("leave") >= 0 || s.indexOf("off") >= 0 || s.indexOf("inactive") >= 0;
}

function getStaffStatusCounts_(staffList) {
  var active = 0;
  var leave = 0;
  for (var i = 0; i < staffList.length; i++) {
    if (isLeaveStatus_(staffList[i].status)) {
      leave++;
    } else {
      active++;
    }
  }
  return { active: active, leave: leave, total: staffList.length };
}

function getHorsesPerStaff_(horses, field) {
  var map = {};
  var i;
  for (i = 0; i < horses.length; i++) {
    var name = String(horses[i][field] || "").trim();
    if (!name) continue;
    map[name] = (map[name] || 0) + 1;
  }
  var out = [];
  for (var k in map) {
    if (map.hasOwnProperty(k)) {
      out.push({ name: k, count: map[k] });
    }
  }
  out.sort(function(a, b) { return b.count - a.count; });
  return out;
}

function getDashboardWidgets_(horses, trainers, grooms, locationFilter, widgetOpts) {
  widgetOpts = widgetOpts || {};
  horses = horses || [];
  trainers = trainers || [];
  grooms = grooms || [];
  var healthy = 0;
  var rehab = 0;
  var watch = 0;
  var hi;
  for (hi = 0; hi < horses.length; hi++) {
    var hs = String(horses[hi].healthStatus || horses[hi].status || "");
    if (isRehabStatus_(hs)) {
      rehab++;
    } else if (hs.indexOf("🟡") >= 0 || /watch/i.test(hs)) {
      watch++;
    } else if (hs.indexOf("🔴") >= 0 || /vet|attention/i.test(hs)) {
      watch++;
    } else {
      healthy++;
    }
  }
  var trainerStats = getStaffStatusCounts_(trainers);
  var groomStats = getStaffStatusCounts_(grooms);
  var profiles = [];
  if (!widgetOpts.skipProfiles) {
    profiles = getHorseProfilesForLocation_(locationFilter || "");
  }
  var stockDays = widgetOpts.stockDays;
  if (!stockDays) {
    stockDays = getStockDaysRemainingFast_(getDailyInventoryRows_(), locationFilter);
  }
  var shoeing = widgetOpts.shoeing;
  if (!shoeing) {
    shoeing = [];
    try {
      shoeing = getShoeingStatus();
    } catch (e1) {}
  }
  var vaccines = widgetOpts.vaccines;
  if (!vaccines) {
    vaccines = [];
    try {
      vaccines = getVaccinationStatus();
    } catch (e2) {}
  }
  if (locationFilter && !widgetOpts.shoeing && !widgetOpts.vaccines) {
    var horseMap = buildHorseLocationMap_(horses);
    shoeing = shoeing.filter(function(s) {
      return horseMatchesLocationFilter_(s.horse, locationFilter, horseMap);
    });
    vaccines = vaccines.filter(function(v) {
      return horseMatchesLocationFilter_(v.horse, locationFilter, horseMap);
    });
  }
  var shoeingDue = null;
  var siPick;
  for (siPick = 0; siPick < shoeing.length; siPick++) {
    var sPick = shoeing[siPick];
    if (sPick.daysLeft == null || isNaN(sPick.daysLeft)) continue;
    if (!shoeingDue || sPick.daysLeft < shoeingDue.daysLeft) {
      shoeingDue = sPick;
    }
  }
  var vetDue = null;
  var dewormDue = null;
  var vi;
  for (vi = 0; vi < vaccines.length; vi++) {
    var v = vaccines[vi];
    if (/deworm/i.test(v.vaccine) && (!dewormDue || v.daysLeft < dewormDue.daysLeft)) {
      dewormDue = v;
    }
    if (!vetDue || v.daysLeft < vetDue.daysLeft) {
      vetDue = v;
    }
  }
  return {
    horses: { total: horses.length, healthy: healthy, rehabilitation: rehab, watch: watch },
    trainers: trainerStats,
    grooms: groomStats,
    horsesPerTrainer: getHorsesPerStaff_(profiles, "trainer"),
    horsesPerGroom: getHorsesPerStaff_(profiles, "groom"),
    stockDays: stockDays,
    healthDue: {
      vetCheck: vetDue ? { label: "Vet check", sub: vetDue.vaccine + " · " + vetDue.horse, days: vetDue.daysLeft } : null,
      deworming: dewormDue ? { label: "De-worming", sub: dewormDue.horse, days: dewormDue.daysLeft } : null,
      shoeing: shoeingDue ? { label: "Shoeing", sub: shoeingDue.horse, days: shoeingDue.daysLeft } : null
    }
  };
}

function diagHorses() {
  var rows = getHorsesList_();
  Logger.log("Horses count: " + rows.length);
  rows.forEach(function(h) { Logger.log("  " + JSON.stringify(h)); });
}

function diagHealth() {
  var rows = getHealthSummary();
  Logger.log("Health rows: " + rows.length);
  rows.forEach(function(h) { Logger.log("  " + JSON.stringify(h)); });
}

function diagActivity() {
  var sheet = getSS_().getSheetByName("ActivityData");
  if (!sheet) { Logger.log("ActivityData sheet NOT FOUND"); return; }
  var data = sheet.getDataRange().getValues();
  Logger.log("ActivityData total rows (incl header): " + data.length);
  Logger.log("Headers: " + JSON.stringify(data[0]));
  if (data.length > 1) Logger.log("Row 1: " + JSON.stringify(data[1]));

  var today = new Date();
  today.setHours(0,0,0,0);
  Logger.log("Today (server time): " + today.toISOString());

  var map = getHeaderIndexMap_(sheet);
  var iDate = findCol_(map, ["Date","date"]);
  Logger.log("Date column index: " + iDate);
  for (var i = 1; i < data.length; i++) {
    var raw = data[i][iDate];
    var parsed = new Date(raw);
    var d0 = new Date(parsed); d0.setHours(0,0,0,0);
    Logger.log("  Row " + i + " raw date=[" + raw + "] parsed=[" + parsed + "] matchesToday=" + (d0.getTime() === today.getTime()));
  }
}

function diagInventory() {
  var raw = getSheetByNameSafe_(["Raw Data","RawData"]);
  Logger.log("Raw Data sheet: " + (raw ? "FOUND rows=" + raw.getLastRow() : "NOT FOUND"));
  var inv = getSheetByNameSafe_(["Daily Inventory","Daily Inventory Summary"]);
  Logger.log("Daily Inventory sheet: " + (inv ? "FOUND rows=" + inv.getLastRow() : "NOT FOUND"));

  var rows = getDailyInventoryRows_();
  Logger.log("getDailyInventoryRows_ count: " + rows.length);
  rows.slice(0,5).forEach(function(r) { Logger.log("  " + JSON.stringify(r)); });
}

function diagFullPayload() {
  var payload = buildDashboardPayload_();
  Logger.log("stats: " + JSON.stringify(payload.stats));
  Logger.log("horses: " + payload.horses.length);
  Logger.log("todayActivity: " + payload.todayActivity.length);
  Logger.log("inventory: " + payload.inventory.length);
  Logger.log("health: " + payload.health.length);
  Logger.log("shoeing: " + payload.shoeing.length);
  Logger.log("vaccinations: " + payload.vaccinations.length);
  Logger.log("medicalHistory: " + payload.medicalHistory.length);
  Logger.log("reports: " + JSON.stringify(payload.reports));
}

// ─── OPERATIONS REPORTS (Daily / Weekly / Monthly) ─────────────
var REPORT_SLOT_CHECK_HOUR_ = { "7AM": 8, "1PM": 14, "6PM": 19, "Extra": 22 };
var REPORT_REORDER_DAYS_LIMIT_ = 14;

function parseReportDateOnly_(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    var d = new Date(val);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  var s = String(val || "").trim();
  if (!s) return null;
  var d2 = new Date(s);
  if (isNaN(d2.getTime())) return null;
  d2.setHours(0, 0, 0, 0);
  return d2;
}

function getReportPeriodRange_(period) {
  var tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var start = new Date(today);
  var end = new Date(today);
  var days = 1;
  var label = "Today";
  period = String(period || "daily").toLowerCase();
  if (period === "weekly") {
    days = 7;
    start.setDate(start.getDate() - 6);
    label = "Last 7 Days";
  } else if (period === "monthly") {
    days = 30;
    start.setDate(start.getDate() - 29);
    label = "Last 30 Days";
  }
  return {
    period: period,
    periodLabel: label,
    days: days,
    start: start,
    end: end,
    startStr: Utilities.formatDate(start, tz, "yyyy-MM-dd"),
    endStr: Utilities.formatDate(end, tz, "yyyy-MM-dd")
  };
}

function getHorseDefaultDailyTotals_(prof) {
  prof = prof || {};
  var std = prof.standardDefaults || {};
  var mult = parseFloat(prof.facilityMultiplier) || 1;
  var w = parseFloat(prof.weight) || 0;
  var feedDefault = 0;
  var keys = ["wet_grass", "dry_grass", "feed_mixed", "barley", "oats"];
  var ki;
  for (ki = 0; ki < keys.length; ki++) {
    var daily = computeFeedDailyKg_(std[keys[ki]], mult, w);
    if (daily === "" && std[keys[ki]]) {
      daily = parseFloat(std[keys[ki]]) || 0;
    }
    feedDefault += parseFloat(daily) || 0;
  }
  return {
    feedDefault: Math.round(feedDefault * 100) / 100,
    waterDefault: Math.round((parseFloat(std.water) || 0) * 100) / 100
  };
}

function getRoutineSlotFeedSum_(row, map, slot) {
  var sum = 0;
  var fi;
  for (fi = 0; fi < ROUTINE_FEED_ITEMS_.length; fi++) {
    var col = findCol_(map, [ROUTINE_FEED_ITEMS_[fi].label + "_" + slot]);
    if (col >= 0) {
      sum += parseFloat(row[col]) || 0;
    }
  }
  return sum;
}

function getRoutineSlotWater_(row, map, slot) {
  var col = findCol_(map, ["Water_" + slot]);
  return col >= 0 ? parseFloat(row[col]) || 0 : 0;
}

function parseRoutineLogRowFromSheet_(row, map) {
  var horseName = String(readSheetCellByAliases_(row, map, ["Horse_Name", "Horse Name", "Horse"]) || "").trim();
  var feedTotal = 0;
  var fi;
  for (fi = 0; fi < ROUTINE_FEED_ITEMS_.length; fi++) {
    var tCol = findCol_(map, [ROUTINE_FEED_ITEMS_[fi].label + "_Total"]);
    if (tCol >= 0) {
      feedTotal += parseFloat(row[tCol]) || 0;
    }
  }
  var wCol = findCol_(map, ["Water_Total"]);
  var waterTotal = wCol >= 0 ? parseFloat(row[wCol]) || 0 : 0;
  var act1 = String(readSheetCellByAliases_(row, map, ["Activity_1", "Activity 1"]) || "").trim();
  var act2 = String(readSheetCellByAliases_(row, map, ["Activity_2", "Activity 2"]) || "").trim();
  var slotFeed = {};
  var slotWater = {};
  var si;
  for (si = 0; si < ROUTINE_FEED_SLOTS_.length; si++) {
    var slot = ROUTINE_FEED_SLOTS_[si];
    slotFeed[slot] = getRoutineSlotFeedSum_(row, map, slot);
    slotWater[slot] = getRoutineSlotWater_(row, map, slot);
  }
  return {
    date: parseReportDateOnly_(readSheetCellByAliases_(row, map, ["Date", "date"])),
    location: String(readSheetCellByAliases_(row, map, ["Location", "location"]) || "").trim(),
    horseName: horseName,
    trainer: String(readSheetCellByAliases_(row, map, ["Trainer", "trainer"]) || "").trim(),
    groom: String(readSheetCellByAliases_(row, map, ["Groom", "groom"]) || "").trim(),
    feedTotal: Math.round(feedTotal * 100) / 100,
    waterTotal: Math.round(waterTotal * 100) / 100,
    hasActivity: !!(act1 || act2),
    activityMinutes: parseRoutineDurationToMinutes_(readSheetCellByAliases_(row, map, ["Duration_1"])) +
      parseRoutineDurationToMinutes_(readSheetCellByAliases_(row, map, ["Duration_2"])),
    slotFeed: slotFeed,
    slotWater: slotWater
  };
}

function readRoutineLogsInRange_(locationFilter, startDate, endDate) {
  var sheet = getSheetByNameSafe_(["DailyRoutineLog"]);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  var data = sheet.getDataRange().getValues();
  var map = getHeaderIndexMap_(sheet);
  var out = [];
  var i;
  for (i = 1; i < data.length; i++) {
    var parsed = parseRoutineLogRowFromSheet_(data[i], map);
    if (!parsed.horseName || !parsed.date) {
      continue;
    }
    if (parsed.date < startDate || parsed.date > endDate) {
      continue;
    }
    if (locationFilter && parsed.location && !matchesLocationFilter_(parsed.location, locationFilter)) {
      continue;
    }
    out.push(parsed);
  }
  return out;
}

function getActivityStatsInRange_(locationFilter, startDate, endDate) {
  var sessions = 0;
  var totalMins = 0;
  var mix = {};
  var cached = getActivitySheetData_();
  var data = cached.data;
  var map = cached.map;
  if (!data || data.length < 2) {
    return { sessions: 0, totalMinutes: 0, activityHours: 0, avgDuration: 0, activityMix: mix };
  }
  var iDate = findCol_(map, ["Date", "date"]);
  var iActivity = findCol_(map, ["Activity", "activity"]);
  var iDuration = findCol_(map, ["Duration", "duration"]);
  var iLoc = findCol_(map, ["Location", "location"]);
  var i;
  for (i = 1; i < data.length; i++) {
    var row = data[i];
    var dt = parseReportDateOnly_(row[iDate]);
    if (!dt || dt < startDate || dt > endDate) {
      continue;
    }
    if (locationFilter && iLoc >= 0 && !matchesLocationFilter_(row[iLoc], locationFilter)) {
      continue;
    }
    sessions++;
    totalMins += parseRoutineDurationToMinutes_(row[iDuration]);
    var act = String(row[iActivity] || "").trim() || "Unknown";
    mix[act] = (mix[act] || 0) + 1;
  }
  return {
    sessions: sessions,
    totalMinutes: totalMins,
    activityHours: Math.round((totalMins / 60) * 10) / 10,
    avgDuration: sessions ? Math.round(totalMins / sessions) : 0,
    activityMix: mix
  };
}

function detectMissedRoutineSlots_(logRow, nowHour) {
  var missedFeed = [];
  var missedWater = [];
  var si;
  for (si = 0; si < ROUTINE_FEED_SLOTS_.length; si++) {
    var slot = ROUTINE_FEED_SLOTS_[si];
    var checkHour = REPORT_SLOT_CHECK_HOUR_[slot];
    if (nowHour >= checkHour) {
      if ((logRow.slotFeed[slot] || 0) <= 0) {
        missedFeed.push(slot);
      }
      if ((logRow.slotWater[slot] || 0) <= 0) {
        missedWater.push(slot);
      }
    }
  }
  return { missedFeed: missedFeed, missedWater: missedWater };
}

function buildOperationsReport_(period, locationFilter) {
  var range = getReportPeriodRange_(period);
  var loc = isAllLocations_(locationFilter) ? "" : String(locationFilter || "").trim();
  var profiles = getHorseProfilesForLocation_(loc);
  var profByName = {};
  var pi;
  for (pi = 0; pi < profiles.length; pi++) {
    profByName[String(profiles[pi].name || "").trim().toLowerCase()] = profiles[pi];
  }
  var logs = readRoutineLogsInRange_(loc, range.start, range.end);
  var activity = getActivityStatsInRange_(loc, range.start, range.end);
  var vaccines = getVaccinationStatus(true);
  var now = new Date();
  var nowHour = now.getHours();
  var isDaily = range.period === "daily";

  var horseAgg = {};
  var li;
  for (li = 0; li < logs.length; li++) {
    var log = logs[li];
    var hKey = log.horseName.toLowerCase();
    if (!horseAgg[hKey]) {
      horseAgg[hKey] = {
        horse: log.horseName,
        feedTotal: 0,
        waterTotal: 0,
        activeDays: {},
        activityMinutes: 0,
        lowFeedDays: 0,
        lowWaterDays: 0,
        inactiveDays: 0,
        logs: 0,
        trainer: log.trainer,
        groom: log.groom,
        flags: []
      };
    }
    var ha = horseAgg[hKey];
    ha.feedTotal += log.feedTotal;
    ha.waterTotal += log.waterTotal;
    ha.activityMinutes += log.activityMinutes;
    ha.logs++;
    if (log.date) {
      ha.activeDays[log.date.getTime()] = true;
    }
    if (!log.hasActivity) {
      ha.inactiveDays++;
    }
    var prof = profByName[hKey] || {};
    var defs = getHorseDefaultDailyTotals_(prof);
    if (defs.feedDefault > 0 && log.feedTotal < defs.feedDefault * 0.85) {
      ha.lowFeedDays++;
    }
    if (defs.waterDefault > 0 && log.waterTotal < defs.waterDefault * 0.80) {
      ha.lowWaterDays++;
    }
    if (isDaily && log.date && log.date.getTime() === range.end.getTime()) {
      var missed = detectMissedRoutineSlots_(log, nowHour);
      if (missed.missedFeed.length) {
        ha.flags.push("Feed Missed (" + missed.missedFeed.join(", ") + ")");
      }
      if (missed.missedWater.length) {
        ha.flags.push("Water Missed (" + missed.missedWater.join(", ") + ")");
      }
    }
  }

  var horseRows = [];
  var alerts = [];
  var totalFeed = 0;
  var totalWater = 0;
  var hk;
  for (hk in horseAgg) {
    if (!horseAgg.hasOwnProperty(hk)) continue;
    var h = horseAgg[hk];
    var prof2 = profByName[hk] || {};
    var def2 = getHorseDefaultDailyTotals_(prof2);
    var expectedFeed = def2.feedDefault * (isDaily ? 1 : range.days);
    var expectedWater = def2.waterDefault * (isDaily ? 1 : range.days);
    totalFeed += h.feedTotal;
    totalWater += h.waterTotal;
    var status = "OK";
    if (h.lowFeedDays > 0) {
      status = "Low Feed Intake";
      alerts.push({ type: "feed", title: h.horse + " — Low Feed", sub: "Actual " + h.feedTotal + " kg vs expected " + Math.round(expectedFeed * 100) / 100 + " kg" });
    }
    if (h.lowWaterDays > 0) {
      status = status === "OK" ? "Low Water Intake" : status + " / Low Water";
      alerts.push({ type: "water", title: h.horse + " — Low Water", sub: "Actual " + h.waterTotal + " L vs expected " + Math.round(expectedWater * 100) / 100 + " L" });
    }
    if (h.inactiveDays > 0 && !h.activityMinutes) {
      status = status === "OK" ? "Inactive" : status;
      alerts.push({ type: "activity", title: h.horse + " — Inactive", sub: "No activity logged in period" });
    }
    if (range.period === "weekly" && (h.lowFeedDays >= 3 || h.lowWaterDays >= 3)) {
      h.flags.push("Health Observation");
      alerts.push({ type: "health", title: h.horse + " — Health Observation", sub: "Low feed/water on 3+ days" });
    }
    var utilization = range.days ? Math.round((Object.keys(h.activeDays).length / range.days) * 100) : 0;
    horseRows.push({
      horse: h.horse,
      trainer: h.trainer || prof2.trainer || "",
      groom: h.groom || prof2.groom || "",
      feedTotal: Math.round(h.feedTotal * 100) / 100,
      waterTotal: Math.round(h.waterTotal * 100) / 100,
      feedDefault: Math.round(expectedFeed * 100) / 100,
      waterDefault: Math.round(expectedWater * 100) / 100,
      feedVariance: Math.round((h.feedTotal - expectedFeed) * 100) / 100,
      waterVariance: Math.round((h.waterTotal - expectedWater) * 100) / 100,
      activityMinutes: h.activityMinutes,
      activityHours: Math.round((h.activityMinutes / 60) * 10) / 10,
      utilizationPct: utilization,
      status: status,
      flags: h.flags,
      healthScore: Math.max(0, 100 - h.lowFeedDays * 10 - h.lowWaterDays * 10 - h.inactiveDays * 5)
    });
  }
  horseRows.sort(function(a, b) { return a.horse.localeCompare(b.horse); });

  for (pi = 0; pi < profiles.length; pi++) {
    var pname = profiles[pi].name;
    if (!horseAgg[String(pname).toLowerCase()]) {
      horseRows.push({
        horse: pname,
        trainer: profiles[pi].trainer || "",
        groom: profiles[pi].groom || "",
        feedTotal: 0,
        waterTotal: 0,
        feedDefault: getHorseDefaultDailyTotals_(profiles[pi]).feedDefault * (isDaily ? 1 : range.days),
        waterDefault: getHorseDefaultDailyTotals_(profiles[pi]).waterDefault * (isDaily ? 1 : range.days),
        feedVariance: 0,
        waterVariance: 0,
        activityMinutes: 0,
        activityHours: 0,
        utilizationPct: 0,
        status: isDaily ? "No log today" : "No logs",
        flags: isDaily ? ["Inactive"] : [],
        healthScore: 0
      });
    }
  }
  horseRows.sort(function(a, b) { return a.horse.localeCompare(b.horse); });

  for (vi = 0; vi < vaccines.length; vi++) {
    var v = vaccines[vi];
    if (loc) {
      var vHorse = String(v.horse || "").trim().toLowerCase();
      if (!profByName[vHorse]) continue;
    }
    if (v.daysLeft <= 15) {
      alerts.push({
        type: "vaccination",
        title: v.horse + " — " + v.vaccine,
        sub: (v.daysLeft < 0 ? "Overdue" : "Due in " + v.daysLeft + " days")
      });
    }
  }

  var avgFeedPerDay = range.days ? Math.round((totalFeed / range.days) * 100) / 100 : 0;
  var avgWaterPerDay = range.days ? Math.round((totalWater / range.days) * 100) / 100 : 0;
  var nextWeekFeed = Math.round(avgFeedPerDay * 7 * 100) / 100;

  var stockDays = getStockDaysRemaining_(loc);
  var inventory = [];
  var invi;
  for (invi = 0; invi < stockDays.length; invi++) {
    var st = stockDays[invi];
    inventory.push({
      item: st.item,
      stock: st.stock,
      dailyUse: st.dailyUse,
      daysRemaining: st.daysLeft,
      reorder: st.daysLeft !== null && st.daysLeft <= REPORT_REORDER_DAYS_LIMIT_
    });
    if (st.daysLeft !== null && st.daysLeft <= REPORT_REORDER_DAYS_LIMIT_) {
      alerts.push({ type: "stock", title: st.item + " — Reorder", sub: st.daysLeft + " days stock remaining" });
    }
  }

  var expectedLogs = Math.max(1, profiles.length * range.days);
  var stableEfficiency = Math.round((logs.length / expectedLogs) * 100);

  var insights = [];
  if (range.period === "monthly" || range.period === "weekly") {
    var sortedFeed = horseRows.slice().sort(function(a, b) { return b.feedTotal - a.feedTotal; });
    var sortedAct = horseRows.slice().sort(function(a, b) { return b.activityHours - a.activityHours; });
    if (sortedFeed.length) {
      insights.push({ label: "Highest feed usage", value: sortedFeed[0].horse + " (" + sortedFeed[0].feedTotal + " kg)" });
      insights.push({ label: "Lowest intake", value: sortedFeed[sortedFeed.length - 1].horse + " (" + sortedFeed[sortedFeed.length - 1].feedTotal + " kg)" });
    }
    if (sortedAct.length && sortedAct[0].activityHours > 0) {
      insights.push({ label: "Highest activity", value: sortedAct[0].horse + " (" + sortedAct[0].activityHours + " hr)" });
    }
    if (inventory.length) {
      insights.push({ label: "Feed likely to finish first", value: inventory[0].item + " (" + (inventory[0].daysRemaining != null ? inventory[0].daysRemaining + " days" : "—") + ")" });
    }
  }

  var vaccDue = vaccines.filter(function(x) { return x.daysLeft <= 15; }).length;
  var vaccTotal = vaccines.length || 1;
  var vaccCompliance = Math.round(((vaccTotal - vaccDue) / vaccTotal) * 100);

  return {
    period: range.period,
    periodLabel: range.periodLabel,
    dateRange: range.startStr + " to " + range.endStr,
    location: loc || "All Locations",
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm"),
    summary: {
      horses: profiles.length,
      logsRecorded: logs.length,
      totalFeedKg: Math.round(totalFeed * 100) / 100,
      totalWaterL: Math.round(totalWater * 100) / 100,
      avgFeedPerDay: avgFeedPerDay,
      avgWaterPerDay: avgWaterPerDay,
      activitySessions: activity.sessions,
      activityHours: activity.activityHours,
      avgDurationMin: activity.avgDuration,
      stableEfficiencyPct: stableEfficiency,
      vaccinationCompliancePct: vaccCompliance,
      nextWeekFeedKg: nextWeekFeed,
      alertCount: alerts.length
    },
    metrics: buildReportMetrics_(range.period, {
      totalFeed: totalFeed,
      totalWater: totalWater,
      avgFeedPerDay: avgFeedPerDay,
      avgWaterPerDay: avgWaterPerDay,
      activity: activity,
      stableEfficiency: stableEfficiency,
      vaccCompliance: vaccCompliance,
      nextWeekFeed: nextWeekFeed,
      profiles: profiles.length
    }),
    horseRows: horseRows,
    alerts: alerts.slice(0, 30),
    insights: insights,
    inventory: inventory,
    activityMix: activity.activityMix
  };
}

function buildReportMetrics_(period, data) {
  var m = [];
  if (period === "daily") {
    m.push({ label: "Daily feed total (all horses)", value: data.totalFeed + " kg", formula: "wet_grass + dry_grass + feed_mix + barley + oats" });
    m.push({ label: "Daily water total", value: data.totalWater + " L", formula: "sum(all water servings)" });
    m.push({ label: "Activity sessions", value: String(data.activity.sessions), formula: "logged activities today" });
    m.push({ label: "Health checks on file", value: String(data.profiles), formula: "horses at location" });
  } else if (period === "weekly") {
    m.push({ label: "Weekly activity hours", value: data.activity.activityHours + " hr", formula: "sum(all daily activity hours)" });
    m.push({ label: "Avg feed per day", value: data.avgFeedPerDay + " kg", formula: "weekly_feed_total ÷ 7" });
    m.push({ label: "Avg water per day", value: data.avgWaterPerDay + " L", formula: "weekly_water_total ÷ 7" });
    m.push({ label: "Next week feed forecast", value: data.nextWeekFeed + " kg", formula: "avg_daily_consumption × 7" });
  } else {
    m.push({ label: "Monthly feed total", value: data.totalFeed + " kg", formula: "sum(all daily feed totals)" });
    m.push({ label: "Monthly water total", value: data.totalWater + " L", formula: "sum(all daily water totals)" });
    m.push({ label: "Avg daily consumption", value: data.avgFeedPerDay + " kg feed · " + data.avgWaterPerDay + " L water", formula: "monthly_total ÷ active_days" });
    m.push({ label: "Stable efficiency", value: data.stableEfficiency + "%", formula: "completed_logs ÷ expected_logs × 100" });
    m.push({ label: "Vaccination compliance", value: data.vaccCompliance + "%", formula: "completed ÷ scheduled × 100" });
  }
  return m;
}

function getOperationsReport(period, location, token) {
  var user = validateSessionToken_(token);
  location = String(location || "").trim();
  if (location && !isAllLocations_(location)) {
    location = assertLocationAccess_(user, location);
  } else if (!isAdminRole_(user.role)) {
    location = resolveLocationName_(user.location || user.locationId);
  }
  return buildOperationsReport_(period || "daily", location);
}



function diagClientPayload() {
  // Simulates exactly what the browser receives — checks for Date objects
  // that break JSON serialization over google.script.run
  var payload = buildDashboardPayload_();
  
  function checkDates(obj, path) {
    if (obj instanceof Date) {
      Logger.log("DATE OBJECT at " + path + " = " + obj);
      return;
    }
    if (Array.isArray(obj)) {
      obj.forEach(function(item, i) { checkDates(item, path + "[" + i + "]"); });
      return;
    }
    if (obj && typeof obj === "object") {
      Object.keys(obj).forEach(function(k) { checkDates(obj[k], path + "." + k); });
    }
  }
  
  checkDates(payload, "payload");
  Logger.log("Done scanning for Date objects");
  
  // Also log the first inventory row and first shoeing row raw
  if (payload.inventory.length) Logger.log("inventory[0]: " + JSON.stringify(payload.inventory[0]));
  if (payload.shoeing.length) Logger.log("shoeing[0]: " + JSON.stringify(payload.shoeing[0]));
  if (payload.health.length) Logger.log("health[0]: " + JSON.stringify(payload.health[0]));
  if (payload.vaccinations.length) Logger.log("vaccinations[0]: " + JSON.stringify(payload.vaccinations[0]));
  if (payload.medicalHistory.length) Logger.log("medicalHistory[0]: " + JSON.stringify(payload.medicalHistory[0]));
}