// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 6_WebApp.gs
// Web app entry point + rider portal backend
// ============================================================

function doGet(e) {
  const app = (e && e.parameter && e.parameter.app) ? String(e.parameter.app) : 'attendance';
  if (app === 'portal') {
    return HtmlService
      .createHtmlOutput(getRiderPortalHtml())
      .setTitle('My Rides · Kings Equestrian')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService
    .createTemplate(getAttendanceAppHtml())
    .evaluate()
    .setTitle('KE Attendance')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ────────────────────────────────────────────────────────────
//  PORTAL BACKEND — getRiderData
//  Change 6: returns { multiProfile: true, profiles: [...] }
//  when multiple riders share the same phone number
// ────────────────────────────────────────────────────────────

function getRiderData(identifier) {
  try {
    identifier = String(identifier || '').trim();
    if (!identifier) return { found: false, error: 'Please enter your phone number or KE Number.' };

    if (identifier.toUpperCase().startsWith('KE')) {
      // Lookup by KE No — always single profile
      const rider = findRiderByKENo(identifier.toUpperCase());
      if (!rider) return { found: false, error: 'No account found. Contact us at +91-9980895533.' };
      return _buildRiderData(rider);
    }

    // Phone lookup — may return multiple profiles
    const riders = findAllRidersByPhone(identifier);
    if (!riders || !riders.length) {
      return { found: false, error: 'No account found. Contact us at +91-9980895533.' };
    }

    if (riders.length === 1) {
      return _buildRiderData(riders[0]);
    }

    // Multiple profiles share this phone — return profile picker
    return {
      found        : true,
      multiProfile : true,
      profiles     : riders.map(r => ({
        keNo  : String(r.row[CONFIG.RIDER_COLS.KE_NO]   || '').trim(),
        name  : r.row[CONFIG.RIDER_COLS.NAME]            || '',
        services: r.row[CONFIG.RIDER_COLS.SERVICES]      || ''
      }))
    };
  } catch (err) {
    Logger.log('getRiderData error: ' + err);
    return { found: false, error: 'Something went wrong. Please try again.' };
  }
}

// Build full rider data object for a single rider
function _buildRiderData(rider) {
  const r    = rider.row;
  const keNo = String(r[CONFIG.RIDER_COLS.KE_NO] || '').trim();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();

  const payments        = _getPaymentsForRider(ss, keNo);
  const sessions        = getSessionsForRider(keNo);
  const classesAttended = _countClassesAttended(ss, keNo);
  const noShowCount     = sessions.filter(function(s) { return String(s.attendance || '').toLowerCase() === 'no-show'; }).length;
  const totalParticipants = sessions.reduce(function(acc, s) {
    return acc + (Number(s.participants) || 1);
  }, 0);

  return {
    found            : true,
    keNo             : keNo,
    name             : r[CONFIG.RIDER_COLS.NAME]         || '',
    phone            : String(r[CONFIG.RIDER_COLS.PHONE] || ''),
    email            : r[CONFIG.RIDER_COLS.EMAIL]        || '',
    services         : r[CONFIG.RIDER_COLS.SERVICES]     || '',
    participants     : r[CONFIG.RIDER_COLS.PARTICIPANTS]  || 1,
    registeredOn     : r[CONFIG.RIDER_COLS.REGISTERED]   ? fmtDate(new Date(r[CONFIG.RIDER_COLS.REGISTERED])) : '',
    payments         : payments,
    sessions         : sessions,
    classesAttended  : classesAttended,
    noShowCount      : noShowCount,
    totalParticipants: totalParticipants
  };
}

// ────────────────────────────────────────────────────────────
//  Change 6: find ALL riders matching a phone number
// ────────────────────────────────────────────────────────────

function findAllRidersByPhone(phone) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return [];
  const target = normalisePhone(phone);
  if (!target || target.length < 10) return [];
  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (normalisePhone(data[i][CONFIG.RIDER_COLS.PHONE]) === target) {
      result.push({ rowIndex: i + 1, row: data[i] });
    }
  }
  return result;
}