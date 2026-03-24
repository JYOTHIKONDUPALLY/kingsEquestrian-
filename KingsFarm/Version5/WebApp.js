// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 6_WebApp.gs
// Web app entry point + rider portal backend
// ============================================================

// ────────────────────────────────────────────────────────────
//  doGet — routes between Attendance App and Rider Portal
//  ?app=portal  → Rider Portal
//  (default)    → Staff Attendance App
// ────────────────────────────────────────────────────────────

function doGet(e) {
  const app = (e && e.parameter && e.parameter.app) ? String(e.parameter.app) : 'attendance';
  if (app === 'portal') {
    return HtmlService
      .createHtmlOutput(getRiderPortalHtml())
      .setTitle('My Rides · Kings Equestrian')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // Default: attendance app
  return HtmlService
    .createTemplate(getAttendanceAppHtml())
    .evaluate()
    .setTitle('KE Attendance')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ────────────────────────────────────────────────────────────
//  PORTAL BACKEND — getRiderData
//  Called via google.script.run from the Rider Portal
// ────────────────────────────────────────────────────────────

function getRiderData(identifier) {
  // identifier can be phone number or KE No (KER...)
  try {
    identifier = String(identifier || '').trim();
    if (!identifier) return { found: false, error: 'Please enter your phone number or KE Number.' };

    let rider = null;
    if (identifier.toUpperCase().startsWith('KER')) {
      rider = findRiderByKENo(identifier.toUpperCase());
    } else {
      rider = findRiderByPhone(identifier);
    }

    if (!rider) return { found: false, error: 'No account found. Contact us at +91-9980895533.' };

    const r    = rider.row;
    const keNo = String(r[CONFIG.RIDER_COLS.KE_NO] || '').trim();

    // Payments — individual transactions only (no totals)
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const payments = _getPaymentsForRider(ss, keNo);

    // Sessions — all (past + upcoming)
    const sessions = getSessionsForRider(keNo);

    // Classes attended count
    const classesAttended = _countClassesAttended(ss, keNo);

    return {
      found          : true,
      keNo           : keNo,
      name           : r[CONFIG.RIDER_COLS.NAME]         || '',
      phone          : String(r[CONFIG.RIDER_COLS.PHONE] || ''),
      email          : r[CONFIG.RIDER_COLS.EMAIL]        || '',
      services       : r[CONFIG.RIDER_COLS.SERVICES]     || '',
      participants   : r[CONFIG.RIDER_COLS.PARTICIPANTS]  || 1,
      registeredOn   : r[CONFIG.RIDER_COLS.REGISTERED]   ? fmtDate(new Date(r[CONFIG.RIDER_COLS.REGISTERED])) : '',
      payments       : payments,        // [{amount, payDate, txnRef, receiptNo, paidOn}]
      sessions       : sessions,        // [{rowIndex, service, date, rawDate, timeSlot, status, attendance, isFuture}]
      classesAttended: classesAttended
    };
  } catch (err) {
    Logger.log('getRiderData error: ' + err);
    return { found: false, error: 'Something went wrong. Please try again.' };
  }
}