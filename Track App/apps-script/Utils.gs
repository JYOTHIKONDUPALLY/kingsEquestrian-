/** Spreadsheet helpers */

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
    "No spreadsheet linked. Open the inventory Google Sheet, run setSpreadsheetId() once, then redeploy."
  );
}

function setSpreadsheetId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Open the Kings Equestrian inventory spreadsheet first.");
  }
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
  return ss.getId();
}

function getSheet_(name) {
  var ss = getSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error("Missing sheet: " + name + ". Run setupInventorySheets() first.");
  }
  return sh;
}

function getSheetData_(name) {
  var sh = getSheet_(name);
  var values = sh.getDataRange().getValues();
  if (!values.length) {
    return { headers: [], rows: [] };
  }
  var headers = values[0].map(function (h) { return String(h || "").trim(); });
  var rows = values.slice(1).filter(function (row) {
    for (var i = 0; i < row.length; i++) {
      if (normalize_(row[i]) !== "") {
        return true;
      }
    }
    return false;
  });
  return { headers: headers, rows: rows };
}

function getRegisterPreview_(sheetName, limit) {
  limit = limit || 25;
  try {
    var data = getSheetData_(sheetName);
    if (!data.rows.length) {
      return { headers: data.headers, rows: [], sheetName: sheetName, count: 0 };
    }
    var slice = data.rows.length > limit
      ? data.rows.slice(data.rows.length - limit)
      : data.rows;
    var rows = slice.map(function (row) {
      return rowToObject_(data.headers, row);
    }).reverse();
    return {
      headers: data.headers,
      rows: rows,
      sheetName: sheetName,
      count: data.rows.length
    };
  } catch (e) {
    return {
      headers: [],
      rows: [],
      sheetName: sheetName,
      count: 0,
      error: e.message || String(e)
    };
  }
}

function findCol_(headers, names) {
  var list = Array.isArray(names) ? names : [names];
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i].toLowerCase();
    for (var j = 0; j < list.length; j++) {
      if (h === String(list[j]).toLowerCase()) {
        return i;
      }
    }
  }
  return -1;
}

function serializeCellValue_(v) {
  if (v == null || v === "") {
    return "";
  }
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(
      v,
      Session.getScriptTimeZone() || "Asia/Kolkata",
      "yyyy-MM-dd HH:mm"
    );
  }
  if (typeof v === "number" || typeof v === "boolean") {
    return v;
  }
  return String(v);
}

function rowToObject_(headers, row) {
  var o = {};
  for (var i = 0; i < headers.length; i++) {
    var key = headers[i];
    if (!key) {
      continue;
    }
    o[key] = serializeCellValue_(row[i]);
  }
  return o;
}

/** Ensures google.script.run can serialize the payload (avoids null responses). */
function sanitizeForClient_(value) {
  if (value == null) {
    return "";
  }
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return serializeCellValue_(value);
  }
  if (Array.isArray(value)) {
    return value.map(sanitizeForClient_);
  }
  if (typeof value === "object") {
    var out = {};
    Object.keys(value).forEach(function (k) {
      out[k] = sanitizeForClient_(value[k]);
    });
    return out;
  }
  return value;
}

function appendRow_(sheetName, row) {
  getSheet_(sheetName).appendRow(row);
}

function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd");
}

function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm");
}

function normalize_(s) {
  return String(s || "").trim();
}

function emailLocalPart_(email) {
  email = normalize_(email).toLowerCase();
  var at = email.indexOf("@");
  return at >= 0 ? email.substring(0, at) : email;
}

/** Match full emails; also allow USER_MASTER local-part only (e.g. kingsequestrianfoundation). */
function emailsMatch_(stored, active) {
  stored = normalize_(stored).toLowerCase();
  active = normalize_(active).toLowerCase();
  if (!stored || !active) {
    return false;
  }
  if (stored === active) {
    return true;
  }
  if (stored.indexOf("@") < 0 && active.indexOf("@") >= 0) {
    return emailLocalPart_(active) === stored;
  }
  if (active.indexOf("@") < 0 && stored.indexOf("@") >= 0) {
    return emailLocalPart_(stored) === active;
  }
  return false;
}

function validateLocation_(loc) {
  if (KE.LOCATIONS.indexOf(loc) < 0) {
    throw new Error("Invalid location. Use: " + KE.LOCATIONS.join(", "));
  }
}

function generateID_(prefix) {
  var sh = getSS_();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var key = "SEQ_" + prefix;
    var props = PropertiesService.getScriptProperties();
    var n = parseInt(props.getProperty(key) || "0", 10) + 1;
    props.setProperty(key, String(n));
    return prefix + "-" + ("0000" + n).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function getActiveEmail_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "";
  } catch (e) {
    return "";
  }
}

function ok_(message, extra) {
  var r = { success: true, message: message || "OK" };
  if (extra) {
    Object.keys(extra).forEach(function (k) { r[k] = extra[k]; });
  }
  return r;
}

function fail_(message) {
  return { success: false, message: message };
}
