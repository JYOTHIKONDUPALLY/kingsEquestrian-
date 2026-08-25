/**
 * Location scope – USER_MASTER location for staff; Admin header picker (default Farm).
 */

function isAdminRole_(role) {
  return normalize_(role) === KE.ROLES.ADMIN;
}

function isAllLocations_(loc) {
  var s = normalize_(loc).toLowerCase();
  return !s || s === "all" || s === "all locations";
}

function resolveLocationName_(loc) {
  loc = normalize_(loc);
  if (!loc) {
    return "";
  }
  for (var i = 0; i < KE.LOCATIONS.length; i++) {
    if (KE.LOCATIONS[i].toLowerCase() === loc.toLowerCase()) {
      return KE.LOCATIONS[i];
    }
  }
  return loc;
}

function resolveViewLocation_(user, requestedView) {
  if (isAdminRole_(user.role)) {
    if (isAllLocations_(requestedView)) {
      return "";
    }
    return resolveLocationName_(requestedView) || KE.DEFAULT_ADMIN_LOCATION;
  }
  var userLoc = resolveLocationName_(user.location);
  if (!userLoc) {
    throw new Error("Your account has no location in USER_MASTER. Contact an administrator.");
  }
  return userLoc;
}

function getLocationContext_(token, requestedView) {
  var user = validateSessionToken_(token);
  var isAdmin = isAdminRole_(user.role);
  if (isAdmin && (requestedView === undefined || requestedView === null)) {
    requestedView = KE.DEFAULT_ADMIN_LOCATION;
  }
  var userLocation = resolveLocationName_(user.location);
  var viewLocation = resolveViewLocation_(user, requestedView);
  var viewLabel = isAllLocations_(viewLocation)
    ? "All Locations"
    : (viewLocation || userLocation || "—");

  var options = [];
  if (isAdmin) {
    options.push({ value: "", label: "All Locations" });
    KE.LOCATIONS.forEach(function (loc) {
      options.push({ value: loc, label: loc });
    });
  }

  return {
    userLocation: userLocation,
    viewLocation: viewLocation,
    viewLabel: viewLabel,
    isAdmin: isAdmin,
    locationLocked: !isAdmin,
    defaultAdminLocation: KE.DEFAULT_ADMIN_LOCATION,
    locationOptions: options
  };
}

function matchesLocation_(cellValue, viewLocation) {
  if (isAllLocations_(viewLocation)) {
    return true;
  }
  return resolveLocationName_(cellValue) === resolveLocationName_(viewLocation);
}

function filterRowsByLocationColumn_(rows, headers, columnName, viewLocation) {
  if (isAllLocations_(viewLocation)) {
    return rows;
  }
  var col = findCol_(headers, columnName);
  if (col < 0) {
    return rows;
  }
  return rows.filter(function (row) {
    return matchesLocation_(row[col], viewLocation);
  });
}

function getRequestIdsForLocation_(viewLocation) {
  var map = {};
  if (isAllLocations_(viewLocation)) {
    return null;
  }
  var data = getSheetData_(KE.SHEETS.REQUEST);
  var cId = findCol_(data.headers, "Request ID");
  var cLoc = findCol_(data.headers, "Location");
  for (var i = 0; i < data.rows.length; i++) {
    if (matchesLocation_(data.rows[i][cLoc], viewLocation)) {
      map[normalize_(data.rows[i][cId])] = true;
    }
  }
  return map;
}

function filterRowsByRequestLocation_(rows, headers, viewLocation) {
  if (isAllLocations_(viewLocation)) {
    return rows;
  }
  var allowed = getRequestIdsForLocation_(viewLocation);
  var cReq = findCol_(headers, "Request ID");
  if (cReq < 0) {
    return [];
  }
  return rows.filter(function (row) {
    return allowed[normalize_(row[cReq])];
  });
}

function applySheetLocationFilter_(sheetName, rows, headers, viewLocation) {
  if (sheetName === KE.SHEETS.PAYMENT) {
    return filterRowsByRequestLocation_(rows, headers, viewLocation);
  }
  return filterRowsByLocationColumn_(rows, headers, "Location", viewLocation);
}

function sheetColumnForView_(sheetName, colName, viewLocation) {
  var data = getSheetData_(sheetName);
  data.rows = applySheetLocationFilter_(sheetName, data.rows, data.headers, viewLocation);
  var col = findCol_(data.headers, colName);
  if (col < 0) {
    return [];
  }
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    var v = normalize_(data.rows[i][col]);
    if (v && out.indexOf(v) < 0) {
      out.push(v);
    }
  }
  return out.sort();
}

function enforcePayloadLocation_(payload, token) {
  payload = payload || {};
  var user = validateSessionToken_(token);
  if (!isAdminRole_(user.role)) {
    var loc = resolveLocationName_(user.location);
    if (!loc) {
      throw new Error("Your account has no location assigned.");
    }
    payload.location = loc;
  } else if (payload.location) {
    payload.location = resolveLocationName_(payload.location);
    validateLocation_(payload.location);
  }
  return payload;
}

function assertRecordLocationAccess_(token, recordLocation, viewLocation) {
  var user = validateSessionToken_(token);
  var view = resolveViewLocation_(user, viewLocation);
  if (isAdminRole_(user.role) && isAllLocations_(view)) {
    return;
  }
  if (!matchesLocation_(recordLocation, view)) {
    throw new Error("This record is not in your selected location (" + view + ").");
  }
}
