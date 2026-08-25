/**
 * Username/email + password login (session token).
 */

var SESSION_PREFIX = "ke_v2_sess_";
var SESSION_TTL_SEC = 21600;

function hashPassword(plain) {
  plain = (plain || "").toString();
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain)
  );
}

function passwordMatchesStored_(stored, plain, inputHash) {
  stored = String(stored || "").trim();
  plain = String(plain || "");
  if (!stored || !plain) {
    return false;
  }
  return stored === inputHash || stored === plain;
}

function createSessionForUser_(email, name, role, location) {
  var token = Utilities.getUuid();
  var userObj = {
    email: String(email || "").trim(),
    name: String(name || "").trim(),
    role: String(role || "").trim(),
    location: String(location || "").trim()
  };
  var sessionJson = JSON.stringify(userObj);
  CacheService.getScriptCache().put(SESSION_PREFIX + token, sessionJson, SESSION_TTL_SEC);
  try {
    PropertiesService.getScriptProperties().setProperty(SESSION_PREFIX + token, sessionJson);
  } catch (e) {
    Logger.log("Session property store: " + e);
  }
  return sanitizeForClient_({ token: token, user: userObj });
}

function loginUser(identifier, password) {
  identifier = (identifier || "").toString().trim().toLowerCase();
  password = (password || "").toString();
  if (!identifier || !password) {
    throw new Error("Please enter username/email and password.");
  }

  var data = getSheetData_(KE.SHEETS.USER);
  if (!data.rows.length) {
    throw new Error("No users in USER_MASTER. Run seedSampleData() or add a user row.");
  }

  var cEmail = findCol_(data.headers, "Email");
  var cName = findCol_(data.headers, "Name");
  var cRole = findCol_(data.headers, "Role");
  var cLoc = findCol_(data.headers, "Location");
  var cPass = findCol_(data.headers, ["Password", "password"]);
  if (cEmail < 0 || cPass < 0) {
    throw new Error("USER_MASTER must have Email and Password columns.");
  }

  var inputHash = hashPassword(password);

  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var email = normalize_(row[cEmail]);
    var name = cName >= 0 ? normalize_(row[cName]) : "";
    if (!email && !name) {
      continue;
    }
    var matches =
      email.toLowerCase() === identifier ||
      (name && name.toLowerCase() === identifier) ||
      (name && name.toLowerCase().replace(/\s+/g, "") === identifier.replace(/\s+/g, "")) ||
      emailsMatch_(email, identifier);
    if (!matches) {
      continue;
    }
    if (!passwordMatchesStored_(row[cPass], password, inputHash)) {
      throw new Error("Invalid password.");
    }
    return createSessionForUser_(
      email,
      name || email,
      cRole >= 0 ? row[cRole] : "",
      cLoc >= 0 ? row[cLoc] : ""
    );
  }

  throw new Error(
    "User not found. Use your Email or Name from USER_MASTER (not your Google sign-in)."
  );
}

function validateSessionToken_(token) {
  token = (token || "").toString().trim();
  if (!token) {
    throw new Error("Please sign in.");
  }
  var raw = CacheService.getScriptCache().get(SESSION_PREFIX + token);
  if (!raw) {
    raw = PropertiesService.getScriptProperties().getProperty(SESSION_PREFIX + token);
  }
  if (!raw) {
    throw new Error("Session expired. Please sign in again.");
  }
  return JSON.parse(raw);
}

function getSessionUser(token) {
  return validateSessionToken_(token);
}

function logoutUser(token) {
  token = (token || "").toString().trim();
  if (token) {
    CacheService.getScriptCache().remove(SESSION_PREFIX + token);
    try {
      PropertiesService.getScriptProperties().deleteProperty(SESSION_PREFIX + token);
    } catch (e) { /* ignore */ }
  }
  return { success: true };
}

function hashUserPassword(userEmailOrName, plainPassword) {
  plainPassword = String(plainPassword || "");
  userEmailOrName = String(userEmailOrName || "").trim().toLowerCase();
  if (!userEmailOrName || !plainPassword) {
    throw new Error("Provide email/name and plain password.");
  }
  var sh = getSheet_(KE.SHEETS.USER);
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h || "").trim(); });
  var cEmail = findCol_(headers, "Email");
  var cName = findCol_(headers, "Name");
  var cPass = findCol_(headers, ["Password", "password"]);
  if (cPass < 0) {
    throw new Error("Add a Password column to USER_MASTER.");
  }
  var hash = hashPassword(plainPassword);
  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][cEmail] || "").trim().toLowerCase();
    var name = cName >= 0 ? String(data[i][cName] || "").trim().toLowerCase() : "";
    if (email === userEmailOrName || name === userEmailOrName) {
      sh.getRange(i + 1, cPass + 1).setValue(hash);
      return { ok: true, hash: hash };
    }
  }
  throw new Error("User not found in USER_MASTER.");
}

function tokenFromPayload_(payload) {
  if (!payload) {
    return "";
  }
  var t = payload.token;
  if (t) {
    delete payload.token;
  }
  return normalize_(t);
}
