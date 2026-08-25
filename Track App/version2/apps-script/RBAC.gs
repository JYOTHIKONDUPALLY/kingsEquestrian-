/**
 * Role-based access – session user, not Google account.
 */

function requireUserFromToken_(token) {
  var user = validateSessionToken_(token);
  if (!user || !user.role) {
    throw new Error("Access denied. Invalid session.");
  }
  return user;
}

function canAccess_(action, user, recordLocation) {
  user = user || {};
  var role = user.role;
  if (role === KE.ROLES.ADMIN) {
    return true;
  }
  if (action === "request" || action === "view") {
    if (role === KE.ROLES.TRAINER || role === KE.ROLES.MANAGER) {
      return true;
    }
  }
  if (action === "issue" || action === "master" || action === "inventory") {
    if (role === KE.ROLES.MANAGER || role === KE.ROLES.ADMIN) {
      if (!recordLocation || !user.location || isAllLocations_(recordLocation)) {
        return true;
      }
      return matchesLocation_(user.location, recordLocation) ||
        matchesLocation_(recordLocation, user.location);
    }
  }
  return false;
}

function assertCan_(action, recordLocation, token) {
  var user = requireUserFromToken_(token);
  if (!canAccess_(action, user, recordLocation)) {
    throw new Error("Your role (" + user.role + ") cannot perform this action.");
  }
  if (recordLocation && !isAdminRole_(user.role) && !isAllLocations_(recordLocation)) {
    var userLoc = resolveLocationName_(user.location);
    if (userLoc && !matchesLocation_(recordLocation, userLoc)) {
      throw new Error("You can only work with data for your location (" + userLoc + ").");
    }
  }
  return user;
}

function assertAdmin_(token) {
  var user = requireUserFromToken_(token);
  if (!isAdminRole_(user.role)) {
    throw new Error("This area is restricted to administrators.");
  }
  return user;
}
