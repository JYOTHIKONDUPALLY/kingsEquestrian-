/**
 * Role-based access – user comes from login session token, not Google account.
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
    if (role === KE.ROLES.TRAINER) {
      return true;
    }
  }
  if (action === "payment" || action === "approve") {
    if (role === KE.ROLES.ACCOUNTS || role === KE.ROLES.ADMIN) {
      return true;
    }
  }
  if (action === "order" || action === "receive" || action === "issue" || action === "master") {
    if (role === KE.ROLES.MANAGER || role === KE.ROLES.ADMIN) {
      if (!recordLocation || !user.location) {
        return true;
      }
      return user.location === recordLocation;
    }
  }
  return false;
}

function assertCan_(action, recordLocation, token, viewLocation) {
  var user = requireUserFromToken_(token);
  if (!canAccess_(action, user, recordLocation)) {
    throw new Error("Your role (" + user.role + ") cannot perform this action.");
  }
  if (recordLocation && !isAdminRole_(user.role)) {
    var userLoc = resolveLocationName_(user.location);
    if (userLoc && !matchesLocation_(recordLocation, userLoc)) {
      throw new Error("You can only work with data for your location (" + userLoc + ").");
    }
  }
  return user;
}

/**
 * True when the role can view full financial data (cost price, margins, vendor cost, etc.).
 * Admin only — Accounts can see payment status but not margins.
 */
function canViewFinancials_(user) {
  if (!user || !user.role) {
    return false;
  }
  return isAdminRole_(user.role);
}

function assertAdmin_(token) {
  var user = requireUserFromToken_(token);
  if (!isAdminRole_(user.role)) {
    throw new Error("This area is restricted to administrators.");
  }
  return user;
}

function canTransferStock_(user) {
  if (!user || !user.role) {
    return false;
  }
  return user.role === KE.ROLES.ADMIN || user.role === KE.ROLES.MANAGER;
}

function canManageSamples_(user) {
  if (!user || !user.role) {
    return false;
  }
  return user.role === KE.ROLES.ADMIN || user.role === KE.ROLES.MANAGER || user.role === KE.ROLES.TRAINER;
}
