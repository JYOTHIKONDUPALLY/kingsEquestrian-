// ============================================================
// KINGS EQUESTRIAN — STAFF / GROOMER DIRECTORY & ATTENDANCE
// ============================================================

var STAFF_MONTHLY_LEAVE_CREDITS = 4;

var GROOMER_HEADERS = [
  'Employee ID', 'Name', 'Photo URL', 'Aadhaar Number', 'Mobile Number',
  'Designation', 'Joining Date', 'Status', 'Leave Balance', 'Last Credit Month',
  'Added At', 'Updated At', 'Updated By', 'Notes',
  'Photo_File_ID', 'Passport_Photo_URL', 'Passport_Photo_File_ID',
  'Bank_Name', 'Bank_Account_No', 'IFSC_Code',
  'Passbook_Photo_URL', 'Passbook_Photo_File_ID', 'Uniform_Issued_Date'
];
var GROOMER_ATTENDANCE_HEADERS = [
  'Date', 'Staff ID', 'Name Snapshot', 'Status', 'Marked At', 'Marked By', 'Notes'
];
var GROOMER_LEAVE_HEADERS = [
  'Leave ID', 'Staff ID', 'Start Date', 'End Date', 'Reason',
  'Applied At', 'Applied By', 'Status', 'Days Deducted'
];

function _styleGroomerHeader_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/** One-time remap from older 9-column GROOMERS layout into the HR schema. */
function _migrateLegacyGroomersSheet_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  var isLegacy = headers[2] === 'Phone' || headers[4] === 'Active'
    || (headers[0] === 'Staff ID' && headers.indexOf('Photo URL') < 0);
  if (!isLegacy) {
    if (sheet.getMaxColumns() < GROOMER_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), GROOMER_HEADERS.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, GROOMER_HEADERS.length).setValues([GROOMER_HEADERS]);
    _styleGroomerHeader_(sheet, GROOMER_HEADERS);
    return;
  }
  var data = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  var out = [GROOMER_HEADERS];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!String(r[0] || '').trim() && !String(r[1] || '').trim()) continue;
    var active = String(r[4] || 'Active').trim().toLowerCase();
    var status = (active === 'yes' || active === 'active' || active === 'true' || active === '1')
      ? 'Active' : 'Inactive';
    out.push([
      String(r[0] || '').trim(),
      String(r[1] || '').trim(),
      '',
      '',
      String(r[2] || '').trim(),
      String(r[3] || 'Groomer').trim() || 'Groomer',
      r[5] || '',
      status,
      STAFF_MONTHLY_LEAVE_CREDITS,
      _groomerMonthKey_(new Date()),
      r[5] || '',
      r[6] || '',
      r[7] || '',
      String(r[8] || '').trim()
    ].concat(new Array(Math.max(0, GROOMER_HEADERS.length - 14)).fill('')));
  }
  sheet.clear();
  if (sheet.getMaxColumns() < GROOMER_HEADERS.length) {
    sheet.insertColumnsAfter(Math.max(sheet.getMaxColumns(), 1), GROOMER_HEADERS.length - Math.max(sheet.getMaxColumns(), 1));
  }
  sheet.getRange(1, 1, out.length, GROOMER_HEADERS.length).setValues(out);
  _styleGroomerHeader_(sheet, GROOMER_HEADERS);
}

function _ensureGroomerSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function ensure(name, headers, migrateFn) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      _styleGroomerHeader_(sheet, headers);
    } else if (typeof migrateFn === 'function') {
      migrateFn(sheet);
    } else {
      var existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      headers.forEach(function (header, i) {
        if (!String(existing[i] || '').trim()) sheet.getRange(1, i + 1).setValue(header);
      });
      _styleGroomerHeader_(sheet, headers);
    }
    return sheet;
  }
  return {
    ss: ss,
    groomers: ensure(CONFIG.SHEETS.GROOMERS, GROOMER_HEADERS, _migrateLegacyGroomersSheet_),
    attendance: ensure(CONFIG.SHEETS.GROOMER_ATTENDANCE, GROOMER_ATTENDANCE_HEADERS),
    leaves: ensure(CONFIG.SHEETS.GROOMER_LEAVES, GROOMER_LEAVE_HEADERS)
  };
}

function _requireGroomerTrainer_(username, token) {
  var trainer = validateTrainerToken(username, token);
  if (!trainer || !trainer.valid) throw new Error('Your trainer session has expired. Sign in again.');
  return trainer;
}

/** Create/repair all staff sheets automatically after trainer login. */
function ensureGroomerAttendanceSetup(username, token) {
  try {
    _requireGroomerTrainer_(username, token);
    _ensureGroomerSheets_();
    return { success: true };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

function _groomerYmd_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var text = String(value).trim();
  var match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return match[1] + '-' + match[2] + '-' + match[3];
  var date = new Date(value);
  return isNaN(date.getTime()) ? '' : Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _groomerDate_(ymdValue) {
  var ymd = _groomerYmd_(ymdValue);
  if (!ymd) return null;
  var parts = ymd.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
}

function _groomerMonthKey_(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'string') {
    var text = String(value).trim();
    // Accept LM-2026-07 / CREDIT:2026-07 / 2026-07 (Sheets date coercion safe token)
    var match = text.match(/(?:LM-|CREDIT:)?(\d{4})-(\d{2})/);
    if (match) return match[1] + '-' + match[2];
  }
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  var ymd = _groomerYmd_(value);
  return ymd ? ymd.substring(0, 7) : '';
}

function _groomerCreditToken_(monthKey) {
  return 'LM-' + String(monthKey || _groomerMonthKey_(new Date()));
}

function _monthsInclusive_(fromYm, toYm) {
  if (!fromYm || !toYm) return 1;
  var from = String(fromYm).split('-');
  var to = String(toYm).split('-');
  if (from.length < 2 || to.length < 2) return 1;
  var n = (Number(to[0]) - Number(from[0])) * 12 + (Number(to[1]) - Number(from[1])) + 1;
  return Math.max(1, n);
}

/** Total leave days deducted for a staff member (active leave rows). */
function _staffLeaveDaysDeductedIndex_(leavesSheet) {
  var map = {};
  if (!leavesSheet || leavesSheet.getLastRow() < 2) return map;
  var data = leavesSheet.getDataRange().getValues();
  var lc = CONFIG.GROOMER_LEAVE_COLS;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][lc.STAFF_ID] || '').trim();
    if (!id) continue;
    if (String(data[i][lc.STATUS] || 'Active').trim().toLowerCase() === 'cancelled') continue;
    var days = Number(data[i][lc.DAYS_DEDUCTED] || 0);
    if (!(days > 0)) {
      days = _countInclusiveDays_(
        _groomerYmd_(data[i][lc.START_DATE]),
        _groomerYmd_(data[i][lc.END_DATE])
      );
    }
    map[id] = (map[id] || 0) + days;
  }
  return map;
}

function _staffLeaveDaysDeducted_(leavesSheet, staffId) {
  var id = String(staffId || '').trim();
  if (!id) return 0;
  return Number(_staffLeaveDaysDeductedIndex_(leavesSheet)[id] || 0);
}

function _ensureMonthlyLeaveCredits_(sheet, rowIndex, row, leavesSheet, deductedIndex) {
  var c = CONFIG.GROOMER_COLS;
  var month = _groomerMonthKey_(new Date());
  var joiningMonth = _groomerMonthKey_(row[c.JOINING_DATE] || row[c.ADDED_AT]) || month;
  if (joiningMonth > month) joiningMonth = month;
  var staffId = String(row[c.STAFF_ID] || '').trim();
  var deducted = deductedIndex
    ? Number(deductedIndex[staffId] || 0)
    : _staffLeaveDaysDeducted_(leavesSheet, staffId);
  var months = _monthsInclusive_(joiningMonth, month);
  var correct = Math.max(0, months * STAFF_MONTHLY_LEAVE_CREDITS - deducted);
  var balance = Number(row[c.LEAVE_BALANCE] || 0);
  var token = _groomerCreditToken_(month);
  var changed = balance !== correct || String(row[c.LAST_CREDIT_MONTH] || '').trim() !== token;

  if (changed) {
    row[c.LEAVE_BALANCE] = correct;
    row[c.LAST_CREDIT_MONTH] = token;
    sheet.getRange(rowIndex, c.LEAVE_BALANCE + 1).setValue(correct);
    sheet.getRange(rowIndex, c.LAST_CREDIT_MONTH + 1).setNumberFormat('@').setValue(token);
  }
  return { balance: correct, credited: changed, months: months, deducted: deducted };
}

function _groomerIsActiveStatus_(value) {
  var text = String(value == null ? 'Active' : value).trim().toLowerCase();
  return text === 'active' || text === 'yes' || text === 'true' || text === '1';
}

function _countInclusiveDays_(startYmd, endYmd) {
  var start = _groomerDate_(startYmd);
  var end = _groomerDate_(endYmd);
  if (!start || !end || end < start) return 0;
  return Math.round((end.getTime() - start.getTime()) / 86400000) + 1;
}

function _findGroomerRow_(sheet, staffId) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var ids = sheet.getRange(2, CONFIG.GROOMER_COLS.STAFF_ID + 1, sheet.getLastRow(), 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === String(staffId || '').trim()) return i + 2;
  }
  return 0;
}

function _nextGroomerId_(sheet, designation) {
  var prefix = String(designation || '').toLowerCase().indexOf('groom') >= 0 ? 'GRM-' : 'EMP-';
  var max = 0;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, CONFIG.GROOMER_COLS.STAFF_ID + 1, sheet.getLastRow(), 1)
      .getValues().forEach(function (row) {
        var id = String(row[0] || '').trim().toUpperCase();
        if (id.indexOf(prefix) !== 0) return;
        max = Math.max(max, Number(id.substring(prefix.length)) || 0);
      });
  }
  return prefix + String(max + 1).padStart(3, '0');
}

function _nextGroomerLeaveId_(sheet) {
  var max = 0;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, CONFIG.GROOMER_LEAVE_COLS.LEAVE_ID + 1, sheet.getLastRow(), 1)
      .getValues().forEach(function (row) {
        var digits = String(row[0] || '').replace(/\D/g, '');
        max = Math.max(max, Number(digits) || 0);
      });
  }
  return 'LV-' + String(max + 1).padStart(4, '0');
}

function _getStaffPhotosFolder_() {
  var rootName = (typeof CONFIG !== 'undefined' && CONFIG.DRIVE_ROOT_FOLDER) ? CONFIG.DRIVE_ROOT_FOLDER : 'Kings Equestrian';
  var rootIter = DriveApp.getFoldersByName(rootName);
  var root = rootIter.hasNext() ? rootIter.next() : DriveApp.createFolder(rootName);
  var subName = 'Staff Documents';
  var subIter = root.getFoldersByName(subName);
  return subIter.hasNext() ? subIter.next() : root.createFolder(subName);
}

function _staffPhotoViewUrl_(fileId, existingUrl) {
  var url = String(existingUrl || '').trim();
  if (url) return url;
  var id = String(fileId || '').trim();
  return id ? ('https://drive.google.com/uc?export=view&id=' + id) : '';
}

function uploadStaffPhoto(payload) {
  payload = payload || {};
  try {
    _requireGroomerTrainer_(payload.username, payload.token);
    var raw = String(payload.base64Data || payload.data || '').trim();
    if (!raw) throw new Error('No photo selected.');
    var mime = String(payload.mimeType || 'image/jpeg').trim();
    var b64 = raw;
    var match = raw.match(/^data:([^;]+);base64,(.+)$/);
    if (match) {
      mime = match[1];
      b64 = match[2];
    }
    var bytes = Utilities.base64Decode(b64);
    if (!bytes || !bytes.length) throw new Error('Could not read photo data.');
    if (bytes.length > 8 * 1024 * 1024) throw new Error('Photo is too large. Try again closer or from gallery.');
    var photoType = String(payload.photoType || 'profile').trim().toLowerCase();
    var typeLabel = photoType === 'passport' ? 'Passport' : (photoType === 'passbook' ? 'Passbook' : 'Profile');
    var staffLabel = String(payload.staffName || payload.staffId || 'staff').trim().replace(/[^\w\-]+/g, '_') || 'staff';
    var ext = mime.indexOf('png') >= 0 ? 'png' : 'jpg';
    var fileName = 'Staff_' + typeLabel + '_' + staffLabel + '_'
      + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.' + ext;
    var file = _getStaffPhotosFolder_().createFile(Utilities.newBlob(bytes, mime, fileName));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('uploadStaffPhoto sharing: ' + shareErr);
    }
    var fileId = file.getId();
    var photoUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
    var oldId = String(payload.existingFileId || '').trim();
    if (oldId && oldId !== fileId) {
      try { DriveApp.getFileById(oldId).setTrashed(true); } catch (ignore) {}
    }
    return { success: true, fileId: fileId, photoUrl: photoUrl, photoType: photoType, message: typeLabel + ' photo uploaded.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

function _staffHrFieldsFromPayload_(payload) {
  payload = payload || {};
  return {
    photoUrl: String(payload.photoUrl || '').trim(),
    photoFileId: String(payload.photoFileId || '').trim(),
    passportPhotoUrl: String(payload.passportPhotoUrl || '').trim(),
    passportPhotoFileId: String(payload.passportPhotoFileId || '').trim(),
    bankName: String(payload.bankName || '').trim(),
    bankAccountNo: String(payload.bankAccountNo || '').trim(),
    ifscCode: String(payload.ifscCode || '').trim(),
    passbookPhotoUrl: String(payload.passbookPhotoUrl || '').trim(),
    passbookPhotoFileId: String(payload.passbookPhotoFileId || '').trim()
  };
}

function _writeStaffHrFields_(row, payload) {
  var c = CONFIG.GROOMER_COLS;
  var h = _staffHrFieldsFromPayload_(payload);
  row[c.PHOTO_URL] = h.photoUrl;
  row[c.PHOTO_FILE_ID] = h.photoFileId;
  row[c.PASSPORT_PHOTO_URL] = h.passportPhotoUrl;
  row[c.PASSPORT_PHOTO_FILE_ID] = h.passportPhotoFileId;
  row[c.BANK_NAME] = h.bankName;
  row[c.BANK_ACCOUNT_NO] = h.bankAccountNo;
  row[c.IFSC_CODE] = h.ifscCode;
  row[c.PASSBOOK_PHOTO_URL] = h.passbookPhotoUrl;
  row[c.PASSBOOK_PHOTO_FILE_ID] = h.passbookPhotoFileId;
}

function _staffRowFromRow_(row) {
  var c = CONFIG.GROOMER_COLS;
  var statusRaw = String(row[c.STATUS] || row[c.ACTIVE] || 'Active').trim();
  var status = _groomerIsActiveStatus_(statusRaw) ? 'Active' : 'Inactive';
  return {
    staffId: String(row[c.STAFF_ID] || '').trim(),
    name: String(row[c.NAME] || '').trim(),
    photoUrl: _staffPhotoViewUrl_(row[c.PHOTO_FILE_ID], row[c.PHOTO_URL]),
    photoFileId: String(row[c.PHOTO_FILE_ID] || '').trim(),
    aadhaar: String(row[c.AADHAAR] || '').trim(),
    phone: String(row[c.PHONE] || '').trim(),
    designation: String(row[c.DESIGNATION] || row[c.ROLE] || 'Staff').trim() || 'Staff',
    joiningDate: _groomerYmd_(row[c.JOINING_DATE] || row[c.ADDED_AT]),
    employmentStatus: status,
    leaveBalance: Number(row[c.LEAVE_BALANCE] || 0),
    lastCreditMonth: String(row[c.LAST_CREDIT_MONTH] || '').trim(),
    notes: String(row[c.NOTES] || '').trim(),
    passportPhotoUrl: _staffPhotoViewUrl_(row[c.PASSPORT_PHOTO_FILE_ID], row[c.PASSPORT_PHOTO_URL]),
    passportPhotoFileId: String(row[c.PASSPORT_PHOTO_FILE_ID] || '').trim(),
    bankName: String(row[c.BANK_NAME] || '').trim(),
    bankAccountNo: String(row[c.BANK_ACCOUNT_NO] || '').trim(),
    ifscCode: String(row[c.IFSC_CODE] || '').trim(),
    passbookPhotoUrl: _staffPhotoViewUrl_(row[c.PASSBOOK_PHOTO_FILE_ID], row[c.PASSBOOK_PHOTO_URL]),
    passbookPhotoFileId: String(row[c.PASSBOOK_PHOTO_FILE_ID] || '').trim(),
    uniformIssuedDate: _groomerYmd_(row[c.UNIFORM_ISSUED_DATE])
  };
}

/**
 * Leave balance = (months since joining, inclusive × 4) − leave days used.
 * Implementation: _ensureMonthlyLeaveCredits_ (defined above; accepts optional deductedIndex).
 */

function _setLeaveBalance_(sheet, rowIndex, balance) {
  sheet.getRange(rowIndex, CONFIG.GROOMER_COLS.LEAVE_BALANCE + 1).setValue(Number(balance) || 0);
}

function _findAttendanceRow_(attendanceSheet, staffId, dateYmd) {
  if (!attendanceSheet || attendanceSheet.getLastRow() < 2) return 0;
  var data = attendanceSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (_groomerYmd_(data[i][CONFIG.GROOMER_ATTENDANCE_COLS.DATE]) === dateYmd
        && String(data[i][CONFIG.GROOMER_ATTENDANCE_COLS.STAFF_ID] || '').trim() === String(staffId || '').trim()) {
      return i + 1;
    }
  }
  return 0;
}

function getGroomerStaffDay(dateYmd, username, token) {
  try {
    _requireGroomerTrainer_(username, token);
    var env = _ensureGroomerSheets_();
    var target = _groomerYmd_(dateYmd) || _groomerYmd_(new Date());
    var today = _groomerYmd_(new Date());
    var attendanceMap = {};
    if (env.attendance.getLastRow() > 1) {
      var attData = env.attendance.getDataRange().getValues();
      for (var a = 1; a < attData.length; a++) {
        if (_groomerYmd_(attData[a][CONFIG.GROOMER_ATTENDANCE_COLS.DATE]) !== target) continue;
        attendanceMap[String(attData[a][CONFIG.GROOMER_ATTENDANCE_COLS.STAFF_ID] || '').trim()] = {
          status: String(attData[a][CONFIG.GROOMER_ATTENDANCE_COLS.STATUS] || '').trim(),
          markedBy: String(attData[a][CONFIG.GROOMER_ATTENDANCE_COLS.MARKED_BY] || '').trim(),
          notes: String(attData[a][CONFIG.GROOMER_ATTENDANCE_COLS.NOTES] || '').trim()
        };
      }
    }

    var leaveMap = {};
    if (env.leaves.getLastRow() > 1) {
      var leaveData = env.leaves.getDataRange().getValues();
      for (var l = 1; l < leaveData.length; l++) {
        if (String(leaveData[l][CONFIG.GROOMER_LEAVE_COLS.STATUS] || 'Active').trim().toLowerCase() === 'cancelled') continue;
        var start = _groomerYmd_(leaveData[l][CONFIG.GROOMER_LEAVE_COLS.START_DATE]);
        var end = _groomerYmd_(leaveData[l][CONFIG.GROOMER_LEAVE_COLS.END_DATE]);
        if (start && end && target >= start && target <= end) {
          leaveMap[String(leaveData[l][CONFIG.GROOMER_LEAVE_COLS.STAFF_ID] || '').trim()] = {
            startDate: start,
            endDate: end,
            reason: String(leaveData[l][CONFIG.GROOMER_LEAVE_COLS.REASON] || '').trim()
          };
        }
      }
    }

    var staff = [];
    var deductedIndex = _staffLeaveDaysDeductedIndex_(env.leaves);
    if (env.groomers.getLastRow() > 1) {
      var data = env.groomers.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowIndex = i + 1;
        var person = _staffRowFromRow_(data[i]);
        if (!person.staffId) continue;
        if (person.employmentStatus !== 'Active') continue;
        if (person.joiningDate && target < person.joiningDate) continue;
        var credit = _ensureMonthlyLeaveCredits_(env.groomers, rowIndex, data[i], env.leaves, deductedIndex);
        person.leaveBalance = credit.balance;
        var marked = attendanceMap[person.staffId] || null;
        var leave = leaveMap[person.staffId] || null;
        var status = marked ? marked.status : (leave ? 'Leave' : (target < today ? 'Absent' : 'Present'));
        person.status = status;
        person.attendanceLocked = !!(marked && marked.status === 'Absent');
        person.defaultedPresent = !marked && !leave && target >= today && status === 'Present';
        person.attendanceNotes = marked ? marked.notes : '';
        person.markedBy = marked ? marked.markedBy : '';
        person.leave = leave;
        staff.push(person);
      }
    }
    staff.sort(function (x, y) { return x.name.localeCompare(y.name); });
    var stats = { total: staff.length, present: 0, absent: 0, leave: 0, notMarked: 0 };
    staff.forEach(function (person) {
      if (person.status === 'Present') stats.present++;
      else if (person.status === 'Absent') stats.absent++;
      else if (person.status === 'Leave') stats.leave++;
      else stats.notMarked++;
    });
    return {
      success: true,
      date: target,
      staff: staff,
      stats: stats,
      monthlyLeaveCredits: STAFF_MONTHLY_LEAVE_CREDITS
    };
  } catch (e) {
    return { success: false, message: String(e.message || e), staff: [] };
  }
}

function addGroomerStaff(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireGroomerTrainer_(payload.username, payload.token);
    var name = String(payload.name || '').trim();
    if (!name) throw new Error('Name is required.');
    var designation = String(payload.designation || payload.role || 'Staff').trim() || 'Staff';
    var employmentStatus = String(payload.employmentStatus || payload.status || 'Active').trim() || 'Active';
    if (employmentStatus.toLowerCase() === 'yes') employmentStatus = 'Active';
    if (employmentStatus.toLowerCase() === 'no') employmentStatus = 'Inactive';
    var env = _ensureGroomerSheets_();
    var staffId = String(payload.staffId || payload.employeeId || '').trim().toUpperCase();
    if (!staffId) staffId = _nextGroomerId_(env.groomers, designation);
    if (_findGroomerRow_(env.groomers, staffId)) throw new Error('Employee ID already exists.');
    var joiningDate = _groomerYmd_(payload.joiningDate) || _groomerYmd_(new Date());
    var now = new Date();
    var month = _groomerMonthKey_(now);
    var c = CONFIG.GROOMER_COLS;
    var row = new Array(GROOMER_HEADERS.length).fill('');
    row[c.STAFF_ID] = staffId;
    row[c.NAME] = name;
    _writeStaffHrFields_(row, payload);
    row[c.AADHAAR] = String(payload.aadhaar || '').trim();
    row[c.PHONE] = String(payload.phone || payload.mobile || '').trim();
    row[c.DESIGNATION] = designation;
    row[c.JOINING_DATE] = _groomerDate_(joiningDate);
    row[c.STATUS] = employmentStatus;
    row[c.LEAVE_BALANCE] = STAFF_MONTHLY_LEAVE_CREDITS;
    row[c.LAST_CREDIT_MONTH] = _groomerCreditToken_(month);
    row[c.ADDED_AT] = now;
    row[c.UPDATED_AT] = now;
    row[c.UPDATED_BY] = trainer.name || trainer.username;
    row[c.NOTES] = String(payload.notes || '').trim();
    row[c.UNIFORM_ISSUED_DATE] = _groomerDate_(_groomerYmd_(payload.uniformIssuedDate)) || '';
    env.groomers.appendRow(row);
    var newRow = env.groomers.getLastRow();
    env.groomers.getRange(newRow, c.LAST_CREDIT_MONTH + 1).setNumberFormat('@').setValue(_groomerCreditToken_(month));
    env.groomers.getRange(newRow, c.LEAVE_BALANCE + 1).setValue(STAFF_MONTHLY_LEAVE_CREDITS);
    return { success: true, message: name + ' added.', staffId: staffId };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function updateGroomerStaff(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireGroomerTrainer_(payload.username, payload.token);
    var env = _ensureGroomerSheets_();
    var rowIndex = _findGroomerRow_(env.groomers, payload.staffId);
    if (!rowIndex) throw new Error('Staff member not found.');
    var name = String(payload.name || '').trim();
    if (!name) throw new Error('Name is required.');
    var employmentStatus = String(payload.employmentStatus || payload.status || 'Active').trim() || 'Active';
    if (employmentStatus.toLowerCase() === 'yes') employmentStatus = 'Active';
    if (employmentStatus.toLowerCase() === 'no') employmentStatus = 'Inactive';
    var joiningDate = _groomerYmd_(payload.joiningDate) || _groomerYmd_(new Date());
    var c = CONFIG.GROOMER_COLS;
    var hr = _staffHrFieldsFromPayload_(payload);
    env.groomers.getRange(rowIndex, c.NAME + 1).setValue(name);
    env.groomers.getRange(rowIndex, c.PHOTO_URL + 1).setValue(hr.photoUrl);
    env.groomers.getRange(rowIndex, c.PHOTO_FILE_ID + 1).setValue(hr.photoFileId);
    env.groomers.getRange(rowIndex, c.PASSPORT_PHOTO_URL + 1).setValue(hr.passportPhotoUrl);
    env.groomers.getRange(rowIndex, c.PASSPORT_PHOTO_FILE_ID + 1).setValue(hr.passportPhotoFileId);
    env.groomers.getRange(rowIndex, c.BANK_NAME + 1).setValue(hr.bankName);
    env.groomers.getRange(rowIndex, c.BANK_ACCOUNT_NO + 1).setValue(hr.bankAccountNo);
    env.groomers.getRange(rowIndex, c.IFSC_CODE + 1).setValue(hr.ifscCode);
    env.groomers.getRange(rowIndex, c.PASSBOOK_PHOTO_URL + 1).setValue(hr.passbookPhotoUrl);
    env.groomers.getRange(rowIndex, c.PASSBOOK_PHOTO_FILE_ID + 1).setValue(hr.passbookPhotoFileId);
    env.groomers.getRange(rowIndex, c.AADHAAR + 1).setValue(String(payload.aadhaar || '').trim());
    env.groomers.getRange(rowIndex, c.PHONE + 1).setValue(String(payload.phone || payload.mobile || '').trim());
    env.groomers.getRange(rowIndex, c.DESIGNATION + 1).setValue(String(payload.designation || payload.role || 'Staff').trim() || 'Staff');
    env.groomers.getRange(rowIndex, c.JOINING_DATE + 1).setValue(_groomerDate_(joiningDate));
    env.groomers.getRange(rowIndex, c.STATUS + 1).setValue(employmentStatus);
    env.groomers.getRange(rowIndex, c.NOTES + 1).setValue(String(payload.notes || '').trim());
    env.groomers.getRange(rowIndex, c.UNIFORM_ISSUED_DATE + 1).setValue(_groomerDate_(_groomerYmd_(payload.uniformIssuedDate)) || '');
    env.groomers.getRange(rowIndex, c.UPDATED_AT + 1).setValue(new Date());
    env.groomers.getRange(rowIndex, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
    return { success: true, message: name + ' updated.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function removeGroomerStaff(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireGroomerTrainer_(payload.username, payload.token);
    var env = _ensureGroomerSheets_();
    var rowIndex = _findGroomerRow_(env.groomers, payload.staffId);
    if (!rowIndex) throw new Error('Staff member not found.');
    env.groomers.getRange(rowIndex, CONFIG.GROOMER_COLS.STATUS + 1).setValue('Inactive');
    env.groomers.getRange(rowIndex, CONFIG.GROOMER_COLS.UPDATED_AT + 1).setValue(new Date());
    env.groomers.getRange(rowIndex, CONFIG.GROOMER_COLS.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
    return { success: true, message: 'Removed from the active staff list.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function saveGroomerAttendance(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireGroomerTrainer_(payload.username, payload.token);
    var status = String(payload.status || '').trim();
    if (status !== 'Present' && status !== 'Absent' && status !== 'Leave') {
      throw new Error('Choose Present, Absent or Leave.');
    }
    var target = _groomerYmd_(payload.date);
    if (!target) throw new Error('Choose a valid attendance date.');
    var env = _ensureGroomerSheets_();
    var staffRow = _findGroomerRow_(env.groomers, payload.staffId);
    if (!staffRow) throw new Error('Staff member not found.');
    var rowVals = env.groomers.getRange(staffRow, 1, 1, GROOMER_HEADERS.length).getValues()[0];
    var credit = _ensureMonthlyLeaveCredits_(env.groomers, staffRow, rowVals, env.leaves);
    var name = String(rowVals[CONFIG.GROOMER_COLS.NAME] || '').trim();
    var staffId = String(payload.staffId || '').trim();
    var existingAttRow = _findAttendanceRow_(env.attendance, staffId, target);
    var previousStatus = '';
    if (existingAttRow) {
      previousStatus = String(env.attendance.getRange(existingAttRow, CONFIG.GROOMER_ATTENDANCE_COLS.STATUS + 1).getValue() || '').trim();
    }
    // Once Absent is recorded for the day, Present is locked (avoids flipping marks)
    if (previousStatus === 'Absent' && status === 'Present') {
      throw new Error(name + ' is already marked Absent for this day. Clear it from the sheet if it was a mistake.');
    }

    if (status === 'Leave' && previousStatus !== 'Leave') {
      if (credit.balance < 1) throw new Error('No leave balance left for ' + name + '.');
      env.leaves.appendRow([
        _nextGroomerLeaveId_(env.leaves), staffId,
        _groomerDate_(target), _groomerDate_(target),
        String(payload.notes || 'Single-day leave').trim(),
        new Date(), trainer.name || trainer.username, 'Active', 1
      ]);
    } else if (previousStatus === 'Leave' && status !== 'Leave') {
      _cancelSingleDayLeave_(env.leaves, staffId, target);
    }

    var attRow = [
      _groomerDate_(target), staffId, name, status,
      new Date(), trainer.name || trainer.username, String(payload.notes || '').trim()
    ];
    if (existingAttRow) env.attendance.getRange(existingAttRow, 1, 1, attRow.length).setValues([attRow]);
    else env.attendance.appendRow(attRow);

    rowVals = env.groomers.getRange(staffRow, 1, 1, GROOMER_HEADERS.length).getValues()[0];
    var refreshed = _ensureMonthlyLeaveCredits_(env.groomers, staffRow, rowVals, env.leaves);
    return {
      success: true,
      message: name + ' marked ' + status + '.',
      leaveBalance: refreshed.balance
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function _cancelSingleDayLeave_(leavesSheet, staffId, dateYmd) {
  if (!leavesSheet || leavesSheet.getLastRow() < 2) return;
  var data = leavesSheet.getDataRange().getValues();
  var lc = CONFIG.GROOMER_LEAVE_COLS;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][lc.STAFF_ID] || '').trim() !== String(staffId || '').trim()) continue;
    if (String(data[i][lc.STATUS] || 'Active').trim().toLowerCase() === 'cancelled') continue;
    var start = _groomerYmd_(data[i][lc.START_DATE]);
    var end = _groomerYmd_(data[i][lc.END_DATE]);
    if (start === dateYmd && end === dateYmd) {
      leavesSheet.getRange(i + 1, lc.STATUS + 1).setValue('Cancelled');
    }
  }
}

function applyGroomerLeave(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireGroomerTrainer_(payload.username, payload.token);
    var start = _groomerYmd_(payload.startDate);
    var end = _groomerYmd_(payload.endDate);
    if (!start || !end) throw new Error('Choose valid leave dates.');
    if (end < start) throw new Error('Leave end date cannot be before start date.');
    var days = _countInclusiveDays_(start, end);
    if (days < 1) throw new Error('Invalid leave duration.');
    var env = _ensureGroomerSheets_();
    var staffRow = _findGroomerRow_(env.groomers, payload.staffId);
    if (!staffRow) throw new Error('Staff member not found.');
    var rowVals = env.groomers.getRange(staffRow, 1, 1, GROOMER_HEADERS.length).getValues()[0];
    var credit = _ensureMonthlyLeaveCredits_(env.groomers, staffRow, rowVals, env.leaves);
    if (credit.balance < days) {
      throw new Error('Need ' + days + ' leave credit(s). Available: ' + credit.balance + '.');
    }
    var name = String(rowVals[CONFIG.GROOMER_COLS.NAME] || '').trim();
    var leaveId = _nextGroomerLeaveId_(env.leaves);
    env.leaves.appendRow([
      leaveId, String(payload.staffId || '').trim(), _groomerDate_(start), _groomerDate_(end),
      String(payload.reason || '').trim(), new Date(), trainer.name || trainer.username, 'Active', days
    ]);

    // Mark each day as Leave in attendance (idempotent).
    var cursor = _groomerDate_(start);
    var endDate = _groomerDate_(end);
    while (cursor && endDate && cursor <= endDate) {
      var ymd = _groomerYmd_(cursor);
      var existing = _findAttendanceRow_(env.attendance, payload.staffId, ymd);
      var attRow = [
        _groomerDate_(ymd), String(payload.staffId || '').trim(), name, 'Leave',
        new Date(), trainer.name || trainer.username, String(payload.reason || '').trim()
      ];
      if (existing) env.attendance.getRange(existing, 1, 1, attRow.length).setValues([attRow]);
      else env.attendance.appendRow(attRow);
      cursor.setDate(cursor.getDate() + 1);
    }

    rowVals = env.groomers.getRange(staffRow, 1, 1, GROOMER_HEADERS.length).getValues()[0];
    var refreshed = _ensureMonthlyLeaveCredits_(env.groomers, staffRow, rowVals, env.leaves);
    return {
      success: true,
      message: 'Leave applied (' + days + ' day' + (days === 1 ? '' : 's') + '). Balance: ' + refreshed.balance + '.',
      leaveId: leaveId,
      leaveBalance: refreshed.balance,
      daysDeducted: days
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}
