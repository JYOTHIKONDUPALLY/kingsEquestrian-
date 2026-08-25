// ============================================================
// KINGS EQUESTRIAN — TRAINER LOGIN & AUDIT
// File: TrainerAuth.gs
// TRAINERS sheet (attendance app login) + password hashing +
// login verification + "who booked / who scored" audit helpers.
// ============================================================

// ────────────────────────────────────────────────────────────
//  PASSWORD HASHING (SHA-256, hex)
// ────────────────────────────────────────────────────────────

function _hashPassword_(plain) {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(plain || ''),
    Utilities.Charset.UTF_8
  );
  return raw.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// ────────────────────────────────────────────────────────────
//  TRAINERS SHEET — SETUP + SEED
// ────────────────────────────────────────────────────────────

function _getOrCreateTrainersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.TRAINERS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.TRAINERS);
    var headers = ['Name', 'Email', 'Phone', 'Aadhar', 'Username', 'Password', 'Password Hash', 'Active', 'Role'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(7, 260);
  }
  _ensureTrainerRoleColumn_(sheet);
  return sheet;
}

function _ensureTrainerRoleColumn_(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TRAINERS);
  if (!sheet) return;
  var roleCol = CONFIG.TRAINER_COLS.ROLE + 1;
  if (sheet.getMaxColumns() < roleCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), roleCol - sheet.getMaxColumns());
  }
  if (!String(sheet.getRange(1, roleCol).getValue() || '').trim()) {
    sheet.getRange(1, roleCol).setValue('Role')
      .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
  }
  if (sheet.getLastRow() > 1) {
    var vals = sheet.getRange(2, roleCol, sheet.getLastRow() - 1, 1).getValues();
    var usernames = sheet.getRange(
      2, CONFIG.TRAINER_COLS.USERNAME + 1, sheet.getLastRow() - 1, 1
    ).getValues();
    var changed = false;
    vals.forEach(function (row, i) {
      if (!String(row[0] || '').trim()) {
        row[0] = String(usernames[i][0] || '').trim().toLowerCase() === 'aarti' ? 'Admin' : 'Trainer';
        changed = true;
      }
    });
    if (changed) sheet.getRange(2, roleCol, vals.length, 1).setValues(vals);
  }
}

/**
 * Menu entry — creates the TRAINERS sheet if missing and seeds a few
 * demo logins. Re-running only adds seed rows that don't already exist
 * (matched by username) and refreshes password hashes.
 */
function setupTrainersSheet() {
  var res = _seedTrainerLogins_();

  try { _ensureScheduleAuditColumns_(); } catch (e) {}

  SpreadsheetApp.getUi().alert(
    'TRAINERS sheet ready.\n\nAdded: ' + res.added + ' trainer(s)\nRefreshed: ' + res.refreshed +
    '\n\nSeed logins (username / password):\n' +
    '  aarti / aarti@123\n  rahul / rahul@123\n  sana / sana@123\n\n' +
    'You can add more trainers directly in the sheet. To set a password, ' +
    'type it in the Password column and run this menu item again to (re)generate the hash.'
  );
}

/**
 * Create/seed the demo trainer logins WITHOUT showing a UI alert.
 * Safe to re-run: updates existing rows (matched by username) and
 * refreshes password hashes. Returns { added, refreshed }.
 */
function _seedTrainerLogins_() {
  var sheet = _getOrCreateTrainersSheet_();

  var seed = [
    // Name, Email, Phone, Aadhar, Username, Password, Role
    ['Aarti Menon', 'aarti.trainer@kingsequestrian.com', '9980011111', '1111-2222-3333', 'aarti', 'aarti@123', 'Admin'],
    ['Rahul Verma', 'rahul.trainer@kingsequestrian.com', '9980022222', '2222-3333-4444', 'rahul', 'rahul@123', 'Trainer'],
    ['Sana Iqbal', 'sana.trainer@kingsequestrian.com', '9980033333', '3333-4444-5555', 'sana', 'sana@123', 'Trainer']
  ];

  var data = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    existing[String(data[i][CONFIG.TRAINER_COLS.USERNAME] || '').trim().toLowerCase()] = i + 1;
  }

  var added = 0, refreshed = 0;
  seed.forEach(function (t) {
    var uname = String(t[4] || '').trim().toLowerCase();
    var hash = _hashPassword_(t[5]);
    var row = [t[0], t[1], t[2], t[3], t[4], t[5], hash, 'Yes', t[6] || 'Trainer'];
    if (existing[uname]) {
      sheet.getRange(existing[uname], 1, 1, row.length).setValues([row]);
      refreshed++;
    } else {
      sheet.appendRow(row);
      added++;
    }
  });

  return { added: added, refreshed: refreshed };
}

/**
 * Regenerate Password Hash for every row from the plaintext Password
 * column. Useful after editing passwords directly in the sheet.
 */
function refreshTrainerPasswordHashes() {
  var sheet = _getOrCreateTrainersSheet_();
  var data = sheet.getDataRange().getValues();
  var updated = 0;
  for (var i = 1; i < data.length; i++) {
    var pw = String(data[i][CONFIG.TRAINER_COLS.PASSWORD] || '');
    if (!pw) continue;
    sheet.getRange(i + 1, CONFIG.TRAINER_COLS.PASS_HASH + 1).setValue(_hashPassword_(pw));
    updated++;
  }
  SpreadsheetApp.getUi().alert('Refreshed password hashes for ' + updated + ' trainer(s).');
}

// ────────────────────────────────────────────────────────────
//  LOGIN (called from attendance PWA)
// ────────────────────────────────────────────────────────────

/**
 * Verify a trainer username + password against the TRAINERS sheet.
 * Returns { success, token, name, username, email } or { success:false, message }.
 */
function trainerLogin(username, password) {
  try {
    var uname = String(username || '').trim().toLowerCase();
    var pw = String(password || '');
    if (!uname || !pw) return { success: false, message: 'Enter username and password.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.TRAINERS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, message: 'No trainers set up yet. Ask admin to run "Setup Trainers Sheet".' };
    }
    _ensureTrainerRoleColumn_(sheet);

    var data = sheet.getDataRange().getValues();
    var hash = _hashPassword_(pw);
    for (var i = 1; i < data.length; i++) {
      var rowUser = String(data[i][CONFIG.TRAINER_COLS.USERNAME] || '').trim().toLowerCase();
      if (rowUser !== uname) continue;

      var active = String(data[i][CONFIG.TRAINER_COLS.ACTIVE] || '').trim().toLowerCase();
      if (active === 'no' || active === 'false' || active === '0') {
        return { success: false, message: 'This account is inactive. Contact admin.' };
      }

      var storedHash = String(data[i][CONFIG.TRAINER_COLS.PASS_HASH] || '').trim().toLowerCase();
      var storedPlain = String(data[i][CONFIG.TRAINER_COLS.PASSWORD] || '');
      var ok = storedHash ? (storedHash === hash) : (storedPlain === pw);
      if (!ok) return { success: false, message: 'Incorrect password.' };

      // Backfill hash if it was missing.
      if (!storedHash && storedPlain) {
        try { sheet.getRange(i + 1, CONFIG.TRAINER_COLS.PASS_HASH + 1).setValue(hash); } catch (e) {}
      }

      var name = String(data[i][CONFIG.TRAINER_COLS.NAME] || rowUser).trim();
      var role = String(data[i][CONFIG.TRAINER_COLS.ROLE] || 'Trainer').trim() || 'Trainer';
      return {
        success: true,
        token: _makeTrainerToken_(rowUser),
        name: name,
        username: rowUser,
        email: String(data[i][CONFIG.TRAINER_COLS.EMAIL] || '').trim(),
        role: role
      };
    }
    return { success: false, message: 'No trainer found with that username.' };
  } catch (e) {
    Logger.log('trainerLogin error: ' + e);
    return { success: false, message: 'Login error. Please try again.' };
  }
}

/** Lightweight signed token so a saved session can be re-validated. */
function _makeTrainerToken_(username) {
  var secret = String(ScriptApp.getScriptId() || 'ke-attendance');
  return _hashPassword_(username + '|' + secret).substring(0, 32);
}

/** Validate a stored session token and return the trainer, or null. */
function validateTrainerToken(username, token) {
  try {
    var uname = String(username || '').trim().toLowerCase();
    if (!uname || !token) return { valid: false };
    if (_makeTrainerToken_(uname) !== String(token)) return { valid: false };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.TRAINERS);
    if (!sheet || sheet.getLastRow() < 2) return { valid: false };
    _ensureTrainerRoleColumn_(sheet);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][CONFIG.TRAINER_COLS.USERNAME] || '').trim().toLowerCase() !== uname) continue;
      var active = String(data[i][CONFIG.TRAINER_COLS.ACTIVE] || '').trim().toLowerCase();
      if (active === 'no' || active === 'false' || active === '0') return { valid: false };
      return {
        valid: true,
        name: String(data[i][CONFIG.TRAINER_COLS.NAME] || uname).trim(),
        username: uname,
        email: String(data[i][CONFIG.TRAINER_COLS.EMAIL] || '').trim(),
        role: String(data[i][CONFIG.TRAINER_COLS.ROLE] || 'Trainer').trim() || 'Trainer'
      };
    }
    return { valid: false };
  } catch (e) {
    Logger.log('validateTrainerToken error: ' + e);
    return { valid: false };
  }
}

/** Validate an authenticated Attendance user and require Admin role. */
function _requireShopAdmin_(username, token) {
  var v = validateTrainerToken(username, token);
  if (!v || !v.valid) throw new Error('Your staff session has expired. Sign in again.');
  if (String(v.role || '').trim().toLowerCase() !== 'admin') {
    throw new Error('Admin access is required to manage shop orders.');
  }
  return v;
}

// ────────────────────────────────────────────────────────────
//  SCHEDULE AUDIT COLUMNS (Booked By / Scored By)
// ────────────────────────────────────────────────────────────

/** Ensure the Schedule sheet has the Booked By / Scored By columns + headers. */
function _ensureScheduleAuditColumns_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sched) return;
  var needCols = CONFIG.SCHED_COLS.SCORED_BY + 1; // 1-based count required
  if (sched.getMaxColumns() < needCols) {
    sched.insertColumnsAfter(sched.getMaxColumns(), needCols - sched.getMaxColumns());
  }
  var headerRow = sched.getRange(1, 1, 1, needCols).getValues()[0];
  if (!String(headerRow[CONFIG.SCHED_COLS.BOOKED_BY] || '').trim()) {
    sched.getRange(1, CONFIG.SCHED_COLS.BOOKED_BY + 1).setValue('Booked By');
  }
  if (!String(headerRow[CONFIG.SCHED_COLS.SCORED_BY] || '').trim()) {
    sched.getRange(1, CONFIG.SCHED_COLS.SCORED_BY + 1).setValue('Scored By');
  }
}

/**
 * Human-readable trainer label for a schedule row.
 * Shows the booker, and if a different trainer scored, appends the scorer last.
 * Rider self-bookings show "Self"; unknown bookers fall back to "Staff".
 */
function _formatSessionTrainerLabel_(bookedBy, scoredBy, source) {
  var b = String(bookedBy || '').trim();
  var s = String(scoredBy || '').trim();
  var src = String(source || '').trim().toLowerCase();

  var bookedLabel;
  if (b) bookedLabel = (b.toLowerCase() === 'self') ? 'Self' : b;
  else if (src === 'rider-portal' || src === 'booking-form') bookedLabel = 'Self';
  else bookedLabel = 'Staff';

  if (s && s.toLowerCase() !== bookedLabel.toLowerCase()) {
    return bookedLabel + ' · Scored by ' + s;
  }
  return s ? ('Scored by ' + s) : bookedLabel;
}
