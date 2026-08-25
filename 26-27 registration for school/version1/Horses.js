// ============================================================
// KINGS EQUESTRIAN — HORSE STABLE REGISTER
// ============================================================

var HORSE_HEADERS = [
  'Horse_ID', 'Horse_Name', 'Location', 'Trainer', 'Groom', 'Status', 'Breed', 'Age',
  'Gender', 'Owner', 'Weight_Kg', 'Facility_Multiplier', 'Lease_Rider', 'Lease_Date',
  'Photo_URL', 'Photo_File_ID', 'Chip_No', 'EFI_ID', 'Date of Birth', 'Colour',
  'Vaccination Date', 'Deworming Date', 'Farrier Date', 'Vet Notes',
  'Added At', 'Updated At', 'Updated By',
  'Vaccination Status', 'Vaccination Postponed To',
  'Deworming Status', 'Deworming Postponed To',
  'Farrier Status', 'Farrier Postponed To'
];

var HORSE_STATUS_OPTIONS = ['Active', 'Leased', 'Rehab', 'Lame', 'Retired'];
var HORSE_GENDER_OPTIONS = ['Stallion', 'Mare', 'Gelding'];
var HORSE_CARE_STATUS_OPTIONS = ['Done', 'Not Done', 'Postponed'];

function _styleHorseHeader_(sheet) {
  sheet.getRange(1, 1, 1, HORSE_HEADERS.length)
    .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/** Remap older HORSES layouts into the expanded register schema. */
function _migrateLegacyHorsesSheet_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  var isLegacy = headers[0] === 'KE Horse ID'
    || (headers[0] === 'Horse_ID' && headers.indexOf('Location') < 0 && headers.length < HORSE_HEADERS.length);
  if (!isLegacy) {
    if (sheet.getMaxColumns() < HORSE_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), HORSE_HEADERS.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, HORSE_HEADERS.length).setValues([HORSE_HEADERS]);
    _styleHorseHeader_(sheet);
    return;
  }
  var data = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  var out = [HORSE_HEADERS];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!String(r[0] || '').trim() && !String(r[1] || '').trim()) continue;
    var dob = r[3] || '';
    var age = _horseAgeLabel_(dob);
    out.push([
      String(r[0] || '').trim(),
      String(r[1] || '').trim(),
      '',
      String(r[7] || '').trim(),
      String(r[8] || '').trim(),
      String(r[9] || 'Active').trim() || 'Active',
      String(r[4] || '').trim(),
      age,
      String(r[6] || '').trim(),
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      String(r[2] || '').trim(),
      dob,
      String(r[5] || '').trim(),
      r[10] || '',
      r[11] || '',
      r[12] || '',
      String(r[13] || '').trim(),
      r[14] || '',
      r[15] || '',
      r[16] || ''
    ]);
  }
  sheet.clear();
  if (sheet.getMaxColumns() < HORSE_HEADERS.length) {
    sheet.insertColumnsAfter(Math.max(sheet.getMaxColumns(), 1), HORSE_HEADERS.length - Math.max(sheet.getMaxColumns(), 1));
  }
  sheet.getRange(1, 1, out.length, HORSE_HEADERS.length).setValues(out);
  _styleHorseHeader_(sheet);
}

function _ensureHorsesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.HORSES);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.HORSES);
  if (sheet.getMaxColumns() < HORSE_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HORSE_HEADERS.length - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HORSE_HEADERS.length).setValues([HORSE_HEADERS]);
    _styleHorseHeader_(sheet);
  } else {
    _migrateLegacyHorsesSheet_(sheet);
  }
  return sheet;
}

function _requireHorseTrainer_(username, token) {
  var trainer = validateTrainerToken(username, token);
  if (!trainer || !trainer.valid) throw new Error('Your trainer session has expired. Sign in again.');
  return trainer;
}

function ensureHorsesSetup(username, token) {
  try {
    _requireHorseTrainer_(username, token);
    _ensureHorsesSheet_();
    return { success: true };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

function _horseYmd_(value) {
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

function _horseDate_(ymdValue) {
  var ymd = _horseYmd_(ymdValue);
  if (!ymd) return '';
  var parts = ymd.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
}

function _horseAgeLabel_(dobYmd) {
  var dob = _horseDate_(dobYmd);
  if (!dob) return '';
  var now = new Date();
  var years = now.getFullYear() - dob.getFullYear();
  var months = now.getMonth() - dob.getMonth();
  if (now.getDate() < dob.getDate()) months -= 1;
  if (months < 0) { years -= 1; months += 12; }
  if (years < 0) return '';
  if (years === 0) return months + ' mo';
  if (months === 0) return years + (years === 1 ? ' yr' : ' yrs');
  return years + (years === 1 ? ' yr ' : ' yrs ') + months + ' mo';
}

function _addDaysYmd_(ymd, days) {
  var date = _horseDate_(ymd);
  if (!date) return '';
  date.setDate(date.getDate() + Number(days || 0));
  return _horseYmd_(date);
}

function _daysFromToday_(ymd) {
  var date = _horseDate_(ymd);
  if (!date) return null;
  var today = _horseDate_(_horseYmd_(new Date()));
  return Math.round((date.getTime() - today.getTime()) / 86400000);
}

function _normalizeCareStatus_(value) {
  var s = String(value || '').trim();
  if (!s) return 'Not Done';
  var lower = s.toLowerCase();
  if (lower === 'done' || lower === 'completed' || lower === 'yes') return 'Done';
  if (lower === 'postponed' || lower === 'postpone' || lower === 'deferred') return 'Postponed';
  if (lower === 'not done' || lower === 'pending' || lower === 'no') return 'Not Done';
  if (HORSE_CARE_STATUS_OPTIONS.indexOf(s) >= 0) return s;
  return 'Not Done';
}

function _healthDueLabel_(daysUntil) {
  if (daysUntil == null) return 'Not recorded';
  if (daysUntil < 0) return 'Overdue by ' + Math.abs(daysUntil) + ' day' + (Math.abs(daysUntil) === 1 ? '' : 's');
  if (daysUntil === 0) return 'Due today';
  if (daysUntil <= 14) return 'Due in ' + daysUntil + ' day' + (daysUntil === 1 ? '' : 's');
  return 'Due in ' + daysUntil + ' days';
}

function _horseCareItem_(key, label, lastDate, status, postponedTo, interval) {
  var careStatus = _normalizeCareStatus_(status);
  var item = {
    key: key,
    label: label,
    lastDate: lastDate || '',
    status: careStatus,
    postponedTo: postponedTo || '',
    interval: interval,
    nextDue: '',
    daysUntil: null,
    dueLabel: 'Not recorded',
    urgency: 9999
  };
  if (careStatus === 'Not Done') {
    item.dueLabel = 'Not done';
    item.daysUntil = -999;
    item.urgency = -999;
    return item;
  }
  if (careStatus === 'Postponed') {
    item.nextDue = postponedTo || '';
    item.daysUntil = item.nextDue ? _daysFromToday_(item.nextDue) : null;
    item.dueLabel = item.nextDue
      ? ('Postponed to ' + item.nextDue + ' · ' + _healthDueLabel_(item.daysUntil))
      : 'Postponed (no new date)';
    item.urgency = item.daysUntil == null ? -500 : item.daysUntil;
    return item;
  }
  // Done — schedule next due from last date + interval
  item.nextDue = lastDate ? _addDaysYmd_(lastDate, interval) : '';
  item.daysUntil = item.nextDue ? _daysFromToday_(item.nextDue) : null;
  item.dueLabel = lastDate ? _healthDueLabel_(item.daysUntil) : 'Done (no date)';
  item.urgency = item.daysUntil == null ? 9999 : item.daysUntil;
  return item;
}

function _horseHealthSummary_(vaccinationYmd, dewormingYmd, farrierYmd, careMeta) {
  careMeta = careMeta || {};
  var intervals = CONFIG.HORSE_HEALTH_INTERVALS || {};
  var items = [
    _horseCareItem_('vaccination', 'Vaccination', vaccinationYmd,
      careMeta.vaccinationStatus, careMeta.vaccinationPostponedTo, intervals.VACCINATION || 365),
    _horseCareItem_('deworming', 'Deworming', dewormingYmd,
      careMeta.dewormingStatus, careMeta.dewormingPostponedTo, intervals.DEWORMING || 90),
    _horseCareItem_('farrier', 'Farrier', farrierYmd,
      careMeta.farrierStatus, careMeta.farrierPostponedTo, intervals.FARRIER || 42)
  ];
  var nearest = null;
  items.forEach(function (item) {
    if (nearest == null || item.urgency < nearest.urgency) nearest = item;
  });
  return {
    vaccination: items[0],
    deworming: items[1],
    farrier: items[2],
    nearest: nearest
  };
}

function _findHorseRow_(sheet, horseId) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var ids = sheet.getRange(2, CONFIG.HORSE_COLS.HORSE_ID + 1, sheet.getLastRow(), 1).getValues();
  var target = String(horseId || '').trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === target) return i + 2;
  }
  return 0;
}

function _nextKeHorseId_(sheet) {
  var prefix = 'KE-H-';
  var max = 0;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, CONFIG.HORSE_COLS.HORSE_ID + 1, sheet.getLastRow(), 1)
      .getValues().forEach(function (row) {
        var id = String(row[0] || '').trim().toUpperCase();
        if (id.indexOf(prefix) !== 0) return;
        max = Math.max(max, Number(id.substring(prefix.length)) || 0);
      });
  }
  return prefix + String(max + 1).padStart(3, '0');
}

function _horseFromRow_(row) {
  var c = CONFIG.HORSE_COLS;
  var dob = _horseYmd_(row[c.DOB]);
  var computedAge = _horseAgeLabel_(dob);
  var storedAge = String(row[c.AGE] || '').trim();
  var vaccination = _horseYmd_(row[c.VACCINATION_DATE]);
  var deworming = _horseYmd_(row[c.DEWORMING_DATE]);
  var farrier = _horseYmd_(row[c.FARRIER_DATE]);
  var careMeta = {
    vaccinationStatus: row[c.VACCINATION_STATUS],
    vaccinationPostponedTo: _horseYmd_(row[c.VACCINATION_POSTPONED_TO]),
    dewormingStatus: row[c.DEWORMING_STATUS],
    dewormingPostponedTo: _horseYmd_(row[c.DEWORMING_POSTPONED_TO]),
    farrierStatus: row[c.FARRIER_STATUS],
    farrierPostponedTo: _horseYmd_(row[c.FARRIER_POSTPONED_TO])
  };
  // Infer Done when a date exists but the status column is empty (legacy rows).
  // The *Assumed flags let the UI show that nobody confirmed the status yet.
  var vaccAssumed = !String(row[c.VACCINATION_STATUS] || '').trim();
  var dewormAssumed = !String(row[c.DEWORMING_STATUS] || '').trim();
  var farrierAssumed = !String(row[c.FARRIER_STATUS] || '').trim();
  if (vaccAssumed && vaccination) careMeta.vaccinationStatus = 'Done';
  if (dewormAssumed && deworming) careMeta.dewormingStatus = 'Done';
  if (farrierAssumed && farrier) careMeta.farrierStatus = 'Done';
  var health = _horseHealthSummary_(vaccination, deworming, farrier, careMeta);
  return {
    horseId: String(row[c.HORSE_ID] || '').trim(),
    keHorseId: String(row[c.HORSE_ID] || '').trim(),
    name: String(row[c.NAME] || '').trim(),
    location: String(row[c.LOCATION] || '').trim(),
    trainer: String(row[c.TRAINER] || '').trim(),
    groom: String(row[c.GROOM] || '').trim(),
    assignedTrainer: String(row[c.TRAINER] || '').trim(),
    assignedGroomer: String(row[c.GROOM] || '').trim(),
    status: String(row[c.STATUS] || 'Active').trim() || 'Active',
    breed: String(row[c.BREED] || '').trim(),
    age: computedAge || storedAge,
    ageStored: storedAge,
    gender: String(row[c.GENDER] || '').trim(),
    owner: String(row[c.OWNER] || '').trim(),
    weightKg: row[c.WEIGHT_KG] === '' || row[c.WEIGHT_KG] == null ? '' : Number(row[c.WEIGHT_KG]),
    facilityMultiplier: row[c.FACILITY_MULTIPLIER] === '' || row[c.FACILITY_MULTIPLIER] == null
      ? '' : Number(row[c.FACILITY_MULTIPLIER]),
    leaseRider: String(row[c.LEASE_RIDER] || '').trim(),
    leaseDate: _horseYmd_(row[c.LEASE_DATE]),
    photoUrl: _horsePhotoViewUrl_(row[c.PHOTO_FILE_ID], row[c.PHOTO_URL]),
    photoFileId: String(row[c.PHOTO_FILE_ID] || '').trim(),
    chipNo: String(row[c.CHIP_NO] || '').trim(),
    efiId: String(row[c.EFI_ID] || '').trim(),
    dob: dob,
    colour: String(row[c.COLOUR] || '').trim(),
    vaccinationDate: vaccination,
    dewormingDate: deworming,
    farrierDate: farrier,
    vaccinationStatus: health.vaccination.status,
    vaccinationPostponedTo: health.vaccination.postponedTo,
    dewormingStatus: health.deworming.status,
    dewormingPostponedTo: health.deworming.postponedTo,
    farrierStatus: health.farrier.status,
    farrierPostponedTo: health.farrier.postponedTo,
    vaccinationAssumed: vaccAssumed,
    dewormingAssumed: dewormAssumed,
    farrierAssumed: farrierAssumed,
    careAssumedCount: (vaccAssumed ? 1 : 0) + (dewormAssumed ? 1 : 0) + (farrierAssumed ? 1 : 0),
    vetNotes: String(row[c.VET_NOTES] || '').trim(),
    health: health,
    nearestCare: health.nearest ? {
      type: health.nearest.label,
      lastDate: health.nearest.lastDate,
      nextDue: health.nearest.nextDue,
      daysUntil: health.nearest.daysUntil,
      dueLabel: health.nearest.dueLabel,
      status: health.nearest.status
    } : null
  };
}

function _getHorsePhotosFolder_() {
  var rootName = (typeof CONFIG !== 'undefined' && CONFIG.DRIVE_ROOT_FOLDER) ? CONFIG.DRIVE_ROOT_FOLDER : 'Kings Equestrian';
  var rootIter = DriveApp.getFoldersByName(rootName);
  var root = rootIter.hasNext() ? rootIter.next() : DriveApp.createFolder(rootName);
  var subName = 'Horse Photos';
  var subIter = root.getFoldersByName(subName);
  return subIter.hasNext() ? subIter.next() : root.createFolder(subName);
}

function _horsePhotoViewUrl_(fileId, existingUrl) {
  var url = String(existingUrl || '').trim();
  if (url) return url;
  var id = String(fileId || '').trim();
  return id ? ('https://drive.google.com/uc?export=view&id=' + id) : '';
}

function uploadHorsePhoto(payload) {
  payload = payload || {};
  try {
    _requireHorseTrainer_(payload.username, payload.token);
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
    var label = String(payload.horseName || payload.horseId || 'horse').trim().replace(/[^\w\-]+/g, '_') || 'horse';
    var ext = mime.indexOf('png') >= 0 ? 'png' : 'jpg';
    var fileName = 'Horse_' + label + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.' + ext;
    var file = _getHorsePhotosFolder_().createFile(Utilities.newBlob(bytes, mime, fileName));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('uploadHorsePhoto sharing: ' + shareErr);
    }
    var fileId = file.getId();
    var photoUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
    var oldId = String(payload.existingFileId || '').trim();
    if (oldId && oldId !== fileId) {
      try { DriveApp.getFileById(oldId).setTrashed(true); } catch (ignore) {}
    }
    return { success: true, fileId: fileId, photoUrl: photoUrl, message: 'Photo uploaded.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

function _horsePayloadToRow_(payload, trainer, existingRow) {
  var c = CONFIG.HORSE_COLS;
  var row = existingRow ? existingRow.slice() : new Array(HORSE_HEADERS.length).fill('');
  var dob = _horseYmd_(payload.dob);
  var age = _horseAgeLabel_(dob) || String(payload.age || '').trim();
  var horseId = String(payload.horseId || payload.keHorseId || row[c.HORSE_ID] || '').trim().toUpperCase();
  var name = String(payload.name || '').trim();
  var gender = String(payload.gender || '').trim();
  var status = String(payload.status || 'Active').trim() || 'Active';
  if (gender && HORSE_GENDER_OPTIONS.indexOf(gender) < 0) throw new Error('Choose Stallion, Mare or Gelding.');
  if (HORSE_STATUS_OPTIONS.indexOf(status) < 0) throw new Error('Invalid status.');
  var now = new Date();
  row[c.HORSE_ID] = horseId;
  row[c.NAME] = name;
  row[c.LOCATION] = String(payload.location || '').trim();
  row[c.TRAINER] = String(payload.trainer || payload.assignedTrainer || '').trim();
  row[c.GROOM] = String(payload.groom || payload.assignedGroomer || '').trim();
  row[c.STATUS] = status;
  row[c.BREED] = String(payload.breed || '').trim();
  row[c.AGE] = age;
  row[c.GENDER] = gender;
  row[c.OWNER] = String(payload.owner || '').trim();
  row[c.WEIGHT_KG] = payload.weightKg === '' || payload.weightKg == null ? '' : Number(payload.weightKg);
  row[c.FACILITY_MULTIPLIER] = payload.facilityMultiplier === '' || payload.facilityMultiplier == null
    ? '' : Number(payload.facilityMultiplier);
  row[c.LEASE_RIDER] = String(payload.leaseRider || '').trim();
  row[c.LEASE_DATE] = _horseDate_(payload.leaseDate) || '';
  row[c.PHOTO_URL] = String(payload.photoUrl || '').trim();
  row[c.PHOTO_FILE_ID] = String(payload.photoFileId || '').trim();
  row[c.CHIP_NO] = String(payload.chipNo || '').trim();
  row[c.EFI_ID] = String(payload.efiId || '').trim();
  row[c.DOB] = _horseDate_(dob) || '';
  row[c.COLOUR] = String(payload.colour || payload.color || '').trim();
  row[c.VACCINATION_DATE] = _horseDate_(payload.vaccinationDate) || '';
  row[c.DEWORMING_DATE] = _horseDate_(payload.dewormingDate) || '';
  row[c.FARRIER_DATE] = _horseDate_(payload.farrierDate) || '';
  row[c.VET_NOTES] = String(payload.vetNotes || '').trim();
  var vStatus = _normalizeCareStatus_(payload.vaccinationStatus);
  var dStatus = _normalizeCareStatus_(payload.dewormingStatus);
  var fStatus = _normalizeCareStatus_(payload.farrierStatus);
  if (!String(payload.vaccinationStatus || '').trim() && payload.vaccinationDate) vStatus = 'Done';
  if (!String(payload.dewormingStatus || '').trim() && payload.dewormingDate) dStatus = 'Done';
  if (!String(payload.farrierStatus || '').trim() && payload.farrierDate) fStatus = 'Done';
  row[c.VACCINATION_STATUS] = vStatus;
  row[c.VACCINATION_POSTPONED_TO] = vStatus === 'Postponed' ? (_horseDate_(payload.vaccinationPostponedTo) || '') : '';
  row[c.DEWORMING_STATUS] = dStatus;
  row[c.DEWORMING_POSTPONED_TO] = dStatus === 'Postponed' ? (_horseDate_(payload.dewormingPostponedTo) || '') : '';
  row[c.FARRIER_STATUS] = fStatus;
  row[c.FARRIER_POSTPONED_TO] = fStatus === 'Postponed' ? (_horseDate_(payload.farrierPostponedTo) || '') : '';
  if (!existingRow) row[c.ADDED_AT] = now;
  row[c.UPDATED_AT] = now;
  row[c.UPDATED_BY] = trainer.name || trainer.username;
  return row;
}

function _listActiveStaffForHorses_() {
  var trainers = [];
  var groomers = [];
  try {
    if (typeof _ensureGroomerSheets_ === 'function') {
      var env = _ensureGroomerSheets_();
      if (env.groomers.getLastRow() > 1) {
        var data = env.groomers.getDataRange().getValues();
        var c = CONFIG.GROOMER_COLS;
        for (var i = 1; i < data.length; i++) {
          var name = String(data[i][c.NAME] || '').trim();
          if (!name) continue;
          var status = String(data[i][c.STATUS] || data[i][c.ACTIVE] || 'Active').trim().toLowerCase();
          if (status !== 'active' && status !== 'yes' && status !== 'true' && status !== '1') continue;
          var designation = String(data[i][c.DESIGNATION] || data[i][c.ROLE] || '').trim().toLowerCase();
          var label = name + ' (' + String(data[i][c.STAFF_ID] || '').trim() + ')';
          if (designation.indexOf('groom') >= 0) groomers.push(label);
          else trainers.push(label);
        }
      }
    }
  } catch (ignore) {}
  return { trainers: trainers, groomers: groomers };
}

function getHorsesList(username, token, statusFilter) {
  try {
    _requireHorseTrainer_(username, token);
    var sheet = _ensureHorsesSheet_();
    var filter = String(statusFilter || 'all').trim().toLowerCase();
    var horses = [];
    var stats = { total: 0, active: 0, leased: 0, rehab: 0, lame: 0, retired: 0, overdue: 0 };
    var customCareIndex = _customCareByHorseIndex_();
    if (sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var horse = _horseFromRow_(data[i]);
        if (!horse.horseId && !horse.name) continue;
        stats.total++;
        var st = horse.status.toLowerCase();
        if (st === 'active') stats.active++;
        else if (st === 'leased') stats.leased++;
        else if (st === 'rehab') stats.rehab++;
        else if (st === 'lame') stats.lame++;
        else if (st === 'retired') stats.retired++;
        if (horse.nearestCare && horse.nearestCare.daysUntil != null && horse.nearestCare.daysUntil < 0) {
          stats.overdue++;
        }
        if (filter && filter !== 'all' && st !== filter) continue;
        horse.customCare = customCareIndex[horse.horseId] || [];
        horses.push(horse);
      }
    }
    horses.sort(function (a, b) {
      var ad = (a.nearestCare && a.nearestCare.daysUntil != null) ? a.nearestCare.daysUntil : 9999;
      var bd = (b.nearestCare && b.nearestCare.daysUntil != null) ? b.nearestCare.daysUntil : 9999;
      if (ad !== bd) return ad - bd;
      return String(a.name || '').localeCompare(String(b.name || ''));
    });
    var staff = _listActiveStaffForHorses_();
    return {
      success: true,
      horses: horses,
      stats: stats,
      statusOptions: HORSE_STATUS_OPTIONS,
      genderOptions: HORSE_GENDER_OPTIONS,
      careStatusOptions: HORSE_CARE_STATUS_OPTIONS,
      careTypeOptions: _listHorseCareTypes_(),
      trainerOptions: staff.trainers,
      groomerOptions: staff.groomers,
      intervals: CONFIG.HORSE_HEALTH_INTERVALS
    };
  } catch (e) {
    return { success: false, message: String(e.message || e), horses: [] };
  }
}

function addHorse(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireHorseTrainer_(payload.username, payload.token);
    var name = String(payload.name || '').trim();
    if (!name) throw new Error('Horse name is required.');
    var sheet = _ensureHorsesSheet_();
    var horseId = String(payload.horseId || payload.keHorseId || '').trim().toUpperCase();
    if (!horseId) horseId = _nextKeHorseId_(sheet);
    if (_findHorseRow_(sheet, horseId)) throw new Error('Horse ID already exists.');
    payload.horseId = horseId;
    var row = _horsePayloadToRow_(payload, trainer, null);
    sheet.appendRow(row);
    _logHorseActivity_({
      horseId: horseId,
      horseName: name,
      activityType: 'Added',
      date: _horseYmd_(new Date()),
      title: 'Horse added to register',
      details: String(payload.status || 'Active'),
      recordedBy: trainer.name || trainer.username
    });
    _syncBuiltInCareActivities_(row, null, trainer);
    return { success: true, message: name + ' added.', horseId: horseId, keHorseId: horseId };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function updateHorse(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireHorseTrainer_(payload.username, payload.token);
    var sheet = _ensureHorsesSheet_();
    var horseId = String(payload.horseId || payload.keHorseId || '').trim();
    var rowIndex = _findHorseRow_(sheet, horseId);
    if (!rowIndex) throw new Error('Horse not found.');
    var existing = sheet.getRange(rowIndex, 1, 1, HORSE_HEADERS.length).getValues()[0];
    payload.horseId = horseId;
    var row = _horsePayloadToRow_(payload, trainer, existing);
    sheet.getRange(rowIndex, 1, 1, HORSE_HEADERS.length).setValues([row]);
    _syncBuiltInCareActivities_(row, existing, trainer);
    _maybeLogLeaseChange_(row, existing, trainer);
    if (String(existing[CONFIG.HORSE_COLS.STATUS] || '') !== String(row[CONFIG.HORSE_COLS.STATUS] || '')) {
      _logHorseActivity_({
        horseId: horseId,
        horseName: String(row[CONFIG.HORSE_COLS.NAME] || '').trim(),
        activityType: 'Status',
        date: _horseYmd_(new Date()),
        title: 'Status → ' + row[CONFIG.HORSE_COLS.STATUS],
        details: 'Was ' + (existing[CONFIG.HORSE_COLS.STATUS] || '—'),
        recordedBy: trainer.name || trainer.username
      });
    }
    return { success: true, message: String(payload.name || row[CONFIG.HORSE_COLS.NAME]).trim() + ' updated.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function removeHorse(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireHorseTrainer_(payload.username, payload.token);
    var sheet = _ensureHorsesSheet_();
    var horseId = String(payload.horseId || payload.keHorseId || '').trim();
    var rowIndex = _findHorseRow_(sheet, horseId);
    if (!rowIndex) throw new Error('Horse not found.');
    sheet.getRange(rowIndex, CONFIG.HORSE_COLS.STATUS + 1).setValue('Retired');
    sheet.getRange(rowIndex, CONFIG.HORSE_COLS.UPDATED_AT + 1).setValue(new Date());
    sheet.getRange(rowIndex, CONFIG.HORSE_COLS.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
    try {
      _logHorseActivity_({
        horseId: horseId,
        horseName: String(sheet.getRange(rowIndex, CONFIG.HORSE_COLS.NAME + 1).getValue() || '').trim(),
        activityType: 'Status',
        date: _horseYmd_(new Date()),
        title: 'Marked Retired',
        details: '',
        recordedBy: trainer.name || trainer.username
      });
    } catch (ignore) {}
    return { success: true, message: 'Horse marked Retired.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


// ── Custom care types + activity history ─────────────────────
var HORSE_CARE_TYPE_HEADERS = ['Type ID', 'Label', 'Interval Days', 'Active', 'Added At', 'Added By'];
var HORSE_ACTIVITY_HEADERS = ['Activity ID', 'Horse ID', 'Horse Name', 'Activity Type', 'Date', 'Title', 'Details', 'Recorded At', 'Recorded By'];
var HORSE_CUSTOM_CARE_HEADERS = ['Row ID', 'Horse ID', 'Type ID', 'Type Label', 'Last Date', 'Status', 'Postponed To', 'Notes', 'Updated At', 'Updated By'];

function _ensureHorseCareSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function ensure(name, headers) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  var types = ensure(CONFIG.SHEETS.HORSE_CARE_TYPES, HORSE_CARE_TYPE_HEADERS);
  if (types.getLastRow() === 1) {
    var now = new Date();
    types.getRange(2, 1, 3, HORSE_CARE_TYPE_HEADERS.length).setValues([
      ['CARE-SHOEING', 'Shoeing', 42, 'Yes', now, 'System'],
      ['CARE-DENTAL', 'Dental', 365, 'Yes', now, 'System'],
      ['CARE-EXTRA-VACC', 'Booster Vaccination', 180, 'Yes', now, 'System']
    ]);
  }
  return {
    types: types,
    activity: ensure(CONFIG.SHEETS.HORSE_ACTIVITY_LOG, HORSE_ACTIVITY_HEADERS),
    custom: ensure(CONFIG.SHEETS.HORSE_CUSTOM_CARE, HORSE_CUSTOM_CARE_HEADERS)
  };
}

function _nextHorseMetaId_(sheet, col, prefix) {
  var max = 0;
  var last = sheet.getLastRow();
  if (last > 1) {
    sheet.getRange(2, col + 1, last, 1).getValues().forEach(function (r) {
      var n = Number(String(r[0] || '').replace(/\D/g, '')) || 0;
      if (n > max) max = n;
    });
  }
  return prefix + String(max + 1).padStart(3, '0');
}

function _listHorseCareTypes_() {
  var sheets = _ensureHorseCareSheets_();
  var out = [];
  if (sheets.types.getLastRow() < 2) return out;
  var c = CONFIG.HORSE_CARE_TYPE_COLS;
  var data = sheets.types.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var active = String(data[i][c.ACTIVE] || 'Yes').trim().toLowerCase();
    if (active === 'no' || active === 'inactive' || active === 'false') continue;
    out.push({
      typeId: String(data[i][c.TYPE_ID] || '').trim(),
      label: String(data[i][c.LABEL] || '').trim(),
      intervalDays: Number(data[i][c.INTERVAL_DAYS] || 0) || 0
    });
  }
  return out;
}

function addHorseCareType(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireHorseTrainer_(payload.username, payload.token);
    var label = String(payload.label || payload.name || '').trim();
    if (!label) throw new Error('Care activity name is required.');
    var sheets = _ensureHorseCareSheets_();
    var existing = _listHorseCareTypes_();
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].label.toLowerCase() === label.toLowerCase()) {
        throw new Error('That care activity already exists.');
      }
    }
    var typeId = String(payload.typeId || '').trim().toUpperCase()
      || _nextHorseMetaId_(sheets.types, CONFIG.HORSE_CARE_TYPE_COLS.TYPE_ID, 'CARE-');
    sheets.types.appendRow([
      typeId, label, Number(payload.intervalDays || 0) || 0, 'Yes',
      new Date(), trainer.name || trainer.username
    ]);
    return { success: true, message: label + ' added.', typeId: typeId, careTypeOptions: _listHorseCareTypes_() };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function _customCareByHorseIndex_() {
  var sheets = _ensureHorseCareSheets_();
  var index = {};
  if (sheets.custom.getLastRow() < 2) return index;
  var c = CONFIG.HORSE_CUSTOM_CARE_COLS;
  var data = sheets.custom.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var horseId = String(data[i][c.HORSE_ID] || '').trim();
    if (!horseId) continue;
    var last = _horseYmd_(data[i][c.LAST_DATE]);
    var status = _normalizeCareStatus_(data[i][c.STATUS]);
    if (!String(data[i][c.STATUS] || '').trim() && last) status = 'Done';
    if (!index[horseId]) index[horseId] = [];
    index[horseId].push({
      rowId: String(data[i][c.ROW_ID] || '').trim(),
      typeId: String(data[i][c.TYPE_ID] || '').trim(),
      label: String(data[i][c.TYPE_LABEL] || '').trim(),
      lastDate: last,
      status: status,
      postponedTo: _horseYmd_(data[i][c.POSTPONED_TO]),
      notes: String(data[i][c.NOTES] || '').trim()
    });
  }
  return index;
}

function _customCareForHorse_(horseId) {
  horseId = String(horseId || '').trim();
  if (!horseId) return [];
  return _customCareByHorseIndex_()[horseId] || [];
}

function upsertHorseCustomCare(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireHorseTrainer_(payload.username, payload.token);
    var horseId = String(payload.horseId || '').trim();
    var typeId = String(payload.typeId || '').trim();
    var label = String(payload.label || payload.typeLabel || '').trim();
    if (!horseId) throw new Error('Horse is required.');
    if (!typeId && !label) throw new Error('Choose a care activity.');
    var sheets = _ensureHorseCareSheets_();
    if (!label || !typeId) {
      var types = _listHorseCareTypes_();
      for (var t = 0; t < types.length; t++) {
        if (typeId && types[t].typeId === typeId) { label = types[t].label; break; }
        if (!typeId && types[t].label.toLowerCase() === label.toLowerCase()) {
          typeId = types[t].typeId; break;
        }
      }
    }
    if (!label) throw new Error('Care activity not found.');
    var status = _normalizeCareStatus_(payload.status);
    if (!String(payload.status || '').trim() && payload.lastDate) status = 'Done';
    var lastDate = _horseYmd_(payload.lastDate) || '';
    var postponed = status === 'Postponed' ? (_horseYmd_(payload.postponedTo) || '') : '';
    if (status === 'Done' && !lastDate) throw new Error('Add the date this care was done.');
    if (status === 'Postponed' && !postponed) throw new Error('Add the postponed-to date.');

    var c = CONFIG.HORSE_CUSTOM_CARE_COLS;
    var rowIndex = 0;
    if (sheets.custom.getLastRow() > 1) {
      var data = sheets.custom.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][c.HORSE_ID] || '').trim() === horseId
          && String(data[i][c.TYPE_ID] || '').trim() === typeId) {
          rowIndex = i + 1; break;
        }
      }
    }
    var now = new Date();
    var horseName = '';
    try {
      var hSheet = _ensureHorsesSheet_();
      var hRow = _findHorseRow_(hSheet, horseId);
      if (hRow) horseName = String(hSheet.getRange(hRow, CONFIG.HORSE_COLS.NAME + 1).getValue() || '').trim();
    } catch (ignore) {}

    if (rowIndex) {
      sheets.custom.getRange(rowIndex, c.TYPE_LABEL + 1).setValue(label);
      sheets.custom.getRange(rowIndex, c.LAST_DATE + 1).setValue(_horseDate_(lastDate) || '');
      sheets.custom.getRange(rowIndex, c.STATUS + 1).setValue(status);
      sheets.custom.getRange(rowIndex, c.POSTPONED_TO + 1).setValue(_horseDate_(postponed) || '');
      sheets.custom.getRange(rowIndex, c.NOTES + 1).setValue(String(payload.notes || '').trim());
      sheets.custom.getRange(rowIndex, c.UPDATED_AT + 1).setValue(now);
      sheets.custom.getRange(rowIndex, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
    } else {
      sheets.custom.appendRow([
        _nextHorseMetaId_(sheets.custom, c.ROW_ID, 'HCR-'),
        horseId, typeId, label,
        _horseDate_(lastDate) || '', status, _horseDate_(postponed) || '',
        String(payload.notes || '').trim(), now, trainer.name || trainer.username
      ]);
    }

    _logHorseActivity_({
      horseId: horseId,
      horseName: horseName,
      activityType: 'Care',
      date: lastDate || postponed || _horseYmd_(now),
      title: label + ' · ' + status,
      details: postponed ? ('Postponed to ' + postponed) : String(payload.notes || '').trim(),
      recordedBy: trainer.name || trainer.username
    });
    return { success: true, message: label + ' saved for this horse.', customCare: _customCareForHorse_(horseId) };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function _logHorseActivity_(entry) {
  entry = entry || {};
  try {
    var sheets = _ensureHorseCareSheets_();
    sheets.activity.appendRow([
      _nextHorseMetaId_(sheets.activity, CONFIG.HORSE_ACTIVITY_COLS.ACTIVITY_ID, 'HA-'),
      String(entry.horseId || '').trim(),
      String(entry.horseName || '').trim(),
      String(entry.activityType || 'Note').trim(),
      _horseDate_(entry.date) || new Date(),
      String(entry.title || '').trim(),
      String(entry.details || '').trim(),
      new Date(),
      String(entry.recordedBy || '').trim()
    ]);
  } catch (e) {
    Logger.log('Horse activity log failed: ' + e);
  }
}

/**
 * If care dates exist on the horse but never made it into HORSE_ACTIVITY_LOG
 * (e.g. earlier ID-bug), write one Care line per built-in care item once.
 */
function _backfillBuiltInCareActivity_(horse) {
  horse = horse || {};
  var horseId = String(horse.horseId || horse.keHorseId || '').trim();
  if (!horseId) return;
  var existing = _getHorseActivity_(horseId, 200);
  var titles = {};
  existing.forEach(function (a) {
    if (String(a.activityType || '') === 'Care') {
      titles[String(a.title || '').toLowerCase()] = true;
    }
  });
  function need(label, date, status, postponed) {
    if (!date && status !== 'Postponed' && status !== 'Done' && status !== 'Not Done') return;
    if (!date && !status && !postponed) return;
    // Skip empty Not Done with no date
    if ((!date && !postponed) && (status === 'Not Done' || !status)) return;
    var keyPrefix = String(label || '').toLowerCase() + ' ·';
    for (var t in titles) {
      if (t.indexOf(keyPrefix) === 0) return;
    }
    _logHorseActivity_({
      horseId: horseId,
      horseName: horse.name || '',
      activityType: 'Care',
      date: date || postponed || _horseYmd_(new Date()),
      title: label + ' · ' + (status || 'Updated'),
      details: postponed ? ('Postponed to ' + postponed) : (date ? ('Date ' + date) : ''),
      recordedBy: 'Backfill'
    });
  }
  need('Vaccination', horse.vaccinationDate, horse.vaccinationStatus, horse.vaccinationPostponedTo);
  need('Deworming', horse.dewormingDate, horse.dewormingStatus, horse.dewormingPostponedTo);
  need('Farrier', horse.farrierDate, horse.farrierStatus, horse.farrierPostponedTo);
  (horse.customCare || []).forEach(function (c) {
    need(c.label || c.typeId, c.lastDate, c.status, c.postponedTo);
  });
}

function _getHorseActivity_(horseId, limit) {
  horseId = String(horseId || '').trim();
  var cap = Math.max(1, Number(limit) || 10);
  var sheets = _ensureHorseCareSheets_();
  var out = [];
  if (sheets.activity.getLastRow() < 2) return out;
  var c = CONFIG.HORSE_ACTIVITY_COLS;
  var data = sheets.activity.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (horseId && String(data[i][c.HORSE_ID] || '').trim() !== horseId) continue;
    out.push({
      activityId: String(data[i][c.ACTIVITY_ID] || '').trim(),
      horseId: String(data[i][c.HORSE_ID] || '').trim(),
      horseName: String(data[i][c.HORSE_NAME] || '').trim(),
      activityType: String(data[i][c.ACTIVITY_TYPE] || '').trim(),
      date: _horseYmd_(data[i][c.DATE]),
      title: String(data[i][c.TITLE] || '').trim(),
      details: String(data[i][c.DETAILS] || '').trim(),
      recordedAt: data[i][c.RECORDED_AT] instanceof Date ? data[i][c.RECORDED_AT].toISOString() : String(data[i][c.RECORDED_AT] || ''),
      recordedBy: String(data[i][c.RECORDED_BY] || '').trim()
    });
  }
  out.sort(function (a, b) {
    var ad = a.date || '';
    var bd = b.date || '';
    if (ad !== bd) return bd.localeCompare(ad);
    return String(b.recordedAt || '').localeCompare(String(a.recordedAt || ''));
  });
  return out.slice(0, cap);
}

function getHorseActivity(username, token, horseId, limit) {
  try {
    _requireHorseTrainer_(username, token);
    var id = String(horseId || '').trim();
    var acts = _getHorseActivity_(id, Math.max(Number(limit) || 10, 50));
    var hasCare = acts.some(function (a) { return String(a.activityType || '') === 'Care'; });
    if (!hasCare) {
      var sheet = _ensureHorsesSheet_();
      var rowIndex = _findHorseRow_(sheet, id);
      if (rowIndex) {
        var row = sheet.getRange(rowIndex, 1, 1, HORSE_HEADERS.length).getValues()[0];
        var horse = _horseFromRow_(row);
        horse.customCare = _customCareForHorse_(id);
        _backfillBuiltInCareActivity_(horse);
        acts = _getHorseActivity_(id, Math.max(Number(limit) || 10, 50));
      }
    }
    return { success: true, activities: acts.slice(0, Number(limit) || 10) };
  } catch (e) {
    return { success: false, message: String(e.message || e), activities: [] };
  }
}

function _syncBuiltInCareActivities_(row, existing, trainer) {
  var c = CONFIG.HORSE_COLS;
  var horseId = String(row[c.HORSE_ID] || '').trim();
  var horseName = String(row[c.NAME] || '').trim();
  var by = trainer.name || trainer.username;
  function check(label, dateCol, statusCol, postCol) {
    var date = _horseYmd_(row[dateCol]);
    var status = String(row[statusCol] || '').trim();
    var post = _horseYmd_(row[postCol]);
    var oldDate = existing ? _horseYmd_(existing[dateCol]) : '';
    var oldStatus = existing ? String(existing[statusCol] || '').trim() : '';
    var oldPost = existing ? _horseYmd_(existing[postCol]) : '';
    if (date === oldDate && status === oldStatus && post === oldPost) return;
    if (!date && !status && !post) return;
    _logHorseActivity_({
      horseId: horseId,
      horseName: horseName,
      activityType: 'Care',
      date: date || post || _horseYmd_(new Date()),
      title: label + ' · ' + (status || 'Updated'),
      details: post ? ('Postponed to ' + post) : (date ? ('Date ' + date) : ''),
      recordedBy: by
    });
  }
  check('Vaccination', c.VACCINATION_DATE, c.VACCINATION_STATUS, c.VACCINATION_POSTPONED_TO);
  check('Deworming', c.DEWORMING_DATE, c.DEWORMING_STATUS, c.DEWORMING_POSTPONED_TO);
  check('Farrier', c.FARRIER_DATE, c.FARRIER_STATUS, c.FARRIER_POSTPONED_TO);
}

function _maybeLogLeaseChange_(row, existing, trainer) {
  if (!existing) return;
  var c = CONFIG.HORSE_COLS;
  var newRider = String(row[c.LEASE_RIDER] || '').trim();
  var oldRider = String(existing[c.LEASE_RIDER] || '').trim();
  var newDate = _horseYmd_(row[c.LEASE_DATE]);
  var oldDate = _horseYmd_(existing[c.LEASE_DATE]);
  if (newRider === oldRider && newDate === oldDate) return;
  var horseId = String(row[c.HORSE_ID] || '').trim();
  var horseName = String(row[c.NAME] || '').trim();
  var by = trainer.name || trainer.username;
  if (newRider && !oldRider) {
    _logHorseActivity_({
      horseId: horseId, horseName: horseName, activityType: 'Lease',
      date: newDate || _horseYmd_(new Date()),
      title: 'Lease out to ' + newRider,
      details: newDate ? ('From ' + newDate) : '',
      recordedBy: by
    });
  } else if (!newRider && oldRider) {
    _logHorseActivity_({
      horseId: horseId, horseName: horseName, activityType: 'Lease',
      date: _horseYmd_(new Date()),
      title: 'Lease returned from ' + oldRider,
      details: oldDate ? ('Was leased from ' + oldDate) : '',
      recordedBy: by
    });
  } else if (newRider !== oldRider || newDate !== oldDate) {
    _logHorseActivity_({
      horseId: horseId, horseName: horseName, activityType: 'Lease',
      date: newDate || _horseYmd_(new Date()),
      title: 'Lease updated → ' + (newRider || '—'),
      details: (oldRider ? ('Was ' + oldRider) : '') + (newDate ? (' · from ' + newDate) : ''),
      recordedBy: by
    });
  }
}
