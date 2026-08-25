/**
 * Audit trail + list query helpers (search / filters / latest N).
 */

function actorName_(user) {
  user = user || {};
  return normalize_(user.name) || normalize_(user.email) || "Unknown";
}

function logActivity_(user, action, entityType, entityId, details, location) {
  try {
    appendRow_(KE.SHEETS.ACTIVITY, [
      nowStr_(),
      normalize_(action),
      normalize_(entityType),
      normalize_(entityId),
      normalize_(details),
      normalize_(location),
      actorName_(user),
      normalize_(user.role)
    ]);
  } catch (e) {
    Logger.log("Activity log failed: " + e);
  }
}

/**
 * Build list query from API payload.
 * Default: latest 10. With search/filters: up to 50 matches.
 */
function listQueryFromPayload_(payload, token, viewLocation) {
  payload = payload || {};
  var user = validateSessionToken_(token);
  var view = resolveViewLocation_(user, viewLocation);
  var filterLoc = normalize_(payload.filterLocation);
  var effectiveLoc = view;

  if (isAdminRole_(user.role)) {
    if (filterLoc) {
      effectiveLoc = isAllLocations_(filterLoc) ? "" : resolveLocationName_(filterLoc);
    }
  } else {
    effectiveLoc = view;
  }

  var search = normalize_(payload.search).toLowerCase();
  var dateFrom = normalize_(payload.dateFrom);
  var dateTo = normalize_(payload.dateTo);
  var status = normalize_(payload.status);
  var hasFilter = !!(search || filterLoc || dateFrom || dateTo || status);
  var limit = Number(payload.limit);
  if (!limit || limit < 1) {
    limit = hasFilter ? 50 : 10;
  }
  if (limit > 100) {
    limit = 100;
  }

  return {
    user: user,
    view: effectiveLoc,
    search: search,
    dateFrom: dateFrom,
    dateTo: dateTo,
    status: status,
    limit: limit,
    hasFilter: hasFilter
  };
}

function objectMatchesSearch_(obj, search, fields) {
  if (!search) {
    return true;
  }
  for (var i = 0; i < fields.length; i++) {
    var val = String(obj[fields[i]] != null ? obj[fields[i]] : "").toLowerCase();
    if (val.indexOf(search) >= 0) {
      return true;
    }
  }
  return false;
}

function filterBySearch_(objects, search, fields) {
  if (!search) {
    return objects;
  }
  return objects.filter(function (o) {
    return objectMatchesSearch_(o, search, fields);
  });
}

function filterByDateField_(objects, dateField, dateFrom, dateTo) {
  if (!dateFrom && !dateTo) {
    return objects;
  }
  return objects.filter(function (o) {
    var raw = String(o[dateField] || "");
    var d = raw.substring(0, 10);
    if (!d) {
      return false;
    }
    if (dateFrom && d < dateFrom) {
      return false;
    }
    if (dateTo && d > dateTo) {
      return false;
    }
    return true;
  });
}

function filterByStatus_(objects, statusField, status) {
  if (!status) {
    return objects;
  }
  return objects.filter(function (o) {
    return normalize_(o[statusField]) === status;
  });
}

/** rows newest-first → take limit; returns { rows, total, shown, limit } */
function paginateLatest_(objectsNewestFirst, limit) {
  var total = objectsNewestFirst.length;
  var rows = objectsNewestFirst.slice(0, limit);
  return {
    rows: rows,
    total: total,
    shown: rows.length,
    limit: limit
  };
}

function listActivity(token, viewLocation, payload) {
  var q = listQueryFromPayload_(payload || {}, token, viewLocation);
  var data = getSheetData_(KE.SHEETS.ACTIVITY);
  var rows = filterRowsByLocationColumn_(data.rows, data.headers, "Location", q.view);
  var objects = rows.map(function (row) {
    return rowToObject_(data.headers, row);
  }).reverse();
  objects = filterBySearch_(objects, q.search, [
    "Action", "Entity Type", "Entity ID", "Details", "User", "Location"
  ]);
  objects = filterByDateField_(objects, "Date", q.dateFrom, q.dateTo);
  return paginateLatest_(objects, q.limit);
}
