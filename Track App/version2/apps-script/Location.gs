/**
 * Location scope – staff locked to USER_MASTER location; Admin can pick / All.
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
  if (loc.toLowerCase() === "all" || loc === KE.LOCATION_ALL) {
    return KE.LOCATION_ALL;
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
  if (!userLoc || userLoc === KE.LOCATION_ALL) {
    throw new Error("Your account has no site location in USER_MASTER. Contact an administrator.");
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
    locationOptions: options,
    masterLocationOptions: [{ value: KE.LOCATION_ALL, label: "All Locations" }].concat(
      KE.LOCATIONS.map(function (loc) {
        return { value: loc, label: loc };
      })
    )
  };
}

function matchesLocation_(cellValue, viewLocation) {
  if (isAllLocations_(viewLocation)) {
    return true;
  }
  var cell = resolveLocationName_(cellValue);
  if (cell === KE.LOCATION_ALL) {
    return true;
  }
  return cell === resolveLocationName_(viewLocation);
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

function enforcePayloadLocation_(payload, token) {
  payload = payload || {};
  var user = validateSessionToken_(token);
  if (isAdminRole_(user.role)) {
    return payload;
  }
  var userLoc = resolveLocationName_(user.location);
  if (userLoc && userLoc !== KE.LOCATION_ALL) {
    if (payload.location != null && normalize_(payload.location) !== "") {
      var requested = resolveLocationName_(payload.location);
      if (requested !== KE.LOCATION_ALL && requested !== userLoc) {
        throw new Error("You can only create records for " + userLoc + ".");
      }
    }
    payload.location = userLoc;
  }
  return payload;
}

function prepPayload_(payload) {
  payload = payload || {};
  var token = tokenFromPayload_(payload);
  var viewLocation = payload.viewLocation != null ? payload.viewLocation : "";
  delete payload.viewLocation;
  payload = enforcePayloadLocation_(payload, token);
  return { token: token, viewLocation: viewLocation, payload: payload };
}
