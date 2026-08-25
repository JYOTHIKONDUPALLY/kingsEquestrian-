// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 4_Schedule.gs
// All sessions (regular + one-time) go into single Schedule sheet
// ============================================================

// ────────────────────────────────────────────────────────────
//  ADD SESSION TO SCHEDULE
// ────────────────────────────────────────────────────────────

function addSessionToSchedule(d) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sheet) { Logger.log('Schedule sheet not found'); return null; }

  if (typeof _ensureScheduleAuditColumns_ === 'function') _ensureScheduleAuditColumns_();
  const newRow = new Array(CONFIG.SCHED_COLS.SCORED_BY + 1).fill('');
  newRow[CONFIG.SCHED_COLS.KE_NO]        = d.keNo        || '';
  newRow[CONFIG.SCHED_COLS.NAME]         = d.name        || '';
  newRow[CONFIG.SCHED_COLS.PHONE]        = d.phone       || '';
  newRow[CONFIG.SCHED_COLS.EMAIL]        = d.email       || '';
  newRow[CONFIG.SCHED_COLS.SERVICE]      = d.service     || '';
  newRow[CONFIG.SCHED_COLS.DATE]         = d.date        || '';
  newRow[CONFIG.SCHED_COLS.TIME_SLOT]    = d.timeSlot    || '';
  newRow[CONFIG.SCHED_COLS.PARTICIPANTS] = d.participants || 1;
  newRow[CONFIG.SCHED_COLS.STATUS]       = d.status      || 'Scheduled';
  newRow[CONFIG.SCHED_COLS.ATTENDANCE]   = '';
  newRow[CONFIG.SCHED_COLS.STAFF_NOTES]  = '';
  newRow[CONFIG.SCHED_COLS.CAL_EVENT_ID] = '';
  newRow[CONFIG.SCHED_COLS.SOURCE]       = d.source      || 'booking-form';

  const src = String(d.source || '').toLowerCase();
  const bookedBy = d.bookedBy || ((src === 'rider-portal' || src === 'booking-form') ? 'Self' : 'Staff');
  newRow[CONFIG.SCHED_COLS.BOOKED_BY]    = bookedBy;
  newRow[CONFIG.SCHED_COLS.SCORED_BY]    = '';

  sheet.appendRow(newRow);
  const lr = sheet.getLastRow();
  if (d.date) sheet.getRange(lr, CONFIG.SCHED_COLS.DATE + 1).setNumberFormat('dd-MMM-yyyy');

  if (d.date && d.timeSlot) {
    try {
      const calId = _createCalEvent(d);
      if (calId) sheet.getRange(lr, CONFIG.SCHED_COLS.CAL_EVENT_ID + 1).setValue(calId);
    } catch (calErr) {
      Logger.log('Calendar event failed (non-fatal): ' + calErr);
    }
  }

  if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();
  Logger.log('Session added to Schedule: ' + d.keNo + ' | ' + fmtDate(d.date));
  return lr;
}

// ────────────────────────────────────────────────────────────
//  GET SESSIONS FOR DATE
// ────────────────────────────────────────────────────────────

function getSessionsForDate(dateStr) {
  const tz = Session.getScriptTimeZone();
  let target;
  if (dateStr === 'today')         { target = new Date(); }
  else if (dateStr === 'tomorrow') { target = new Date(); target.setDate(target.getDate() + 1); }
  else                             { target = new Date(dateStr); }
  const targetYMD = Utilities.formatDate(target, tz, 'yyyy-MM-dd');

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const date = row[CONFIG.SCHED_COLS.DATE];
    if (!date) continue;
    let rowYMD;
    try { rowYMD = Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd'); } catch(e) { continue; }
    if (rowYMD !== targetYMD) continue;
    if (String(row[CONFIG.SCHED_COLS.STATUS] || '').toLowerCase() === 'cancelled') continue;

    const keNo     = String(row[CONFIG.SCHED_COLS.KE_NO] || '').trim();
    const payments = _getPaymentsForRider(ss, keNo);

    results.push({
      rowIndex    : i + 1,
      keNo        : keNo,
      name        : row[CONFIG.SCHED_COLS.NAME]         || '',
      phone       : String(row[CONFIG.SCHED_COLS.PHONE] || ''),
      email       : row[CONFIG.SCHED_COLS.EMAIL]        || '',
      service     : row[CONFIG.SCHED_COLS.SERVICE]      || '',
      timeSlot    : row[CONFIG.SCHED_COLS.TIME_SLOT]    || '',
      participants: row[CONFIG.SCHED_COLS.PARTICIPANTS]  || 1,
      status      : row[CONFIG.SCHED_COLS.STATUS]       || '',
      attendance  : row[CONFIG.SCHED_COLS.ATTENDANCE]   || '',
      staffNotes  : row[CONFIG.SCHED_COLS.STAFF_NOTES]  || '',
      source      : row[CONFIG.SCHED_COLS.SOURCE]       || '',
      payments    : payments,
      classesAttended: _countClassesAttended(ss, keNo)
    });
  }

  results.sort((a, b) => (a.timeSlot || '').localeCompare(b.timeSlot || ''));
  return results;
}

// ────────────────────────────────────────────────────────────
//  GET ALL RIDERS WITH STATS
// ────────────────────────────────────────────────────────────

function getAllRidersWithStats() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const ridersSheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!ridersSheet) return [];

  const data    = ridersSheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const keNo = String(row[CONFIG.RIDER_COLS.KE_NO] || '').trim();
    if (!keNo) continue;

    const payments        = _getPaymentsForRider(ss, keNo);
    const classesAttended = _countClassesAttended(ss, keNo);
    const nextSession     = _getNextSession(ss, keNo);

    results.push({
      keNo        : keNo,
      name        : row[CONFIG.RIDER_COLS.NAME]         || '',
      phone       : String(row[CONFIG.RIDER_COLS.PHONE] || ''),
      email       : row[CONFIG.RIDER_COLS.EMAIL]        || '',
      services    : row[CONFIG.RIDER_COLS.SERVICES]     || '',
      participants: row[CONFIG.RIDER_COLS.PARTICIPANTS]  || 1,
      registeredOn: row[CONFIG.RIDER_COLS.REGISTERED]   ? fmtDate(new Date(row[CONFIG.RIDER_COLS.REGISTERED])) : '',
      payments    : payments,
      classesAttended,
      nextSession
    });
  }

  results.sort((a, b) => a.name.localeCompare(b.name));
  return results;
}

// ────────────────────────────────────────────────────────────
//  SAVE ATTENDANCE — Change 5: sends Present / No-Show emails
// ────────────────────────────────────────────────────────────

function saveAttendance(rowIndex, status, note) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet) return { success: false, error: 'Schedule sheet not found' };

    const attCell = sheet.getRange(rowIndex, CONFIG.SCHED_COLS.ATTENDANCE + 1);
    attCell.setValue(status);
    _colourAttCell(attCell, status);

    const statusCell = sheet.getRange(rowIndex, CONFIG.SCHED_COLS.STATUS + 1);
    if (status === 'Present')      statusCell.setValue('Completed').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    else if (status === 'No-Show') statusCell.setValue('No-Show').setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');

    if (note !== null && note !== undefined) {
      sheet.getRange(rowIndex, CONFIG.SCHED_COLS.STAFF_NOTES + 1).setValue(note);
    }

    // Send attendance email (Present OR No-Show)
    let emailSent = false;
    try {
      const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
      let email = String(rowData[CONFIG.SCHED_COLS.EMAIL] || '').trim();
      if (!email) {
        // fallback: scan riders sheet for email by KE No
        const keNo = String(rowData[CONFIG.SCHED_COLS.KE_NO] || '').trim();
        if (keNo) {
          const ridersSheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
          if (ridersSheet) {
            const ridersData = ridersSheet.getDataRange().getValues();
            for (let ri = 1; ri < ridersData.length; ri++) {
              if (String(ridersData[ri][CONFIG.RIDER_COLS.KE_NO] || '').trim() === keNo) {
                email = String(ridersData[ri][CONFIG.RIDER_COLS.EMAIL] || '').trim();
                break;
              }
            }
          }
        }
      }

      const emailData = {
        name        : rowData[CONFIG.SCHED_COLS.NAME]         || '',
        email       : email,
        keNo        : rowData[CONFIG.SCHED_COLS.KE_NO]        || '',
        service     : rowData[CONFIG.SCHED_COLS.SERVICE]      || '',
        timeSlot    : rowData[CONFIG.SCHED_COLS.TIME_SLOT]    || '',
        participants: rowData[CONFIG.SCHED_COLS.PARTICIPANTS]  || 1
      };

      if (status === 'Present')      emailSent = sendPresentEmail(emailData);
      else if (status === 'No-Show') emailSent = sendNoShowEmail(emailData);

      if (!emailSent) {
        Logger.log('Attendance email was not sent (maybe invalid/missing recipient) for row ' + rowIndex + ', status ' + status + ', email ' + email);
      }
    } catch (mailErr) {
      Logger.log('Attendance email failed (non-fatal): ' + mailErr);
    }

    Logger.log('Attendance saved: row ' + rowIndex + ' → ' + status + (emailSent ? ' (email sent)' : ' (email not sent)'));
    return { success: true, emailSent: emailSent };
  } catch (err) {
    Logger.log('saveAttendance error: ' + err);
    return { success: false, error: err.message };
  }
}

function saveAssessmentFromAttendance(rowIndex, scores) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sh) return { success: false, error: 'Schedule sheet not found' };
    var row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
    var keNo = String(row[CONFIG.SCHED_COLS.KE_NO] || '').trim();
    if (!keNo) return { success: false, error: 'Missing KE No' };

    var studentId = _findStudentIdByKENo_(keNo);
    if (!studentId) return { success: false, error: 'Student_ID mapping not found' };

    var bookings = ss.getSheetByName('BOOKINGS');
    var bookingId = _findLatestBookingIdForStudent_(bookings, studentId);
    if (!bookingId) return { success: false, error: 'Booking_ID not found for student' };

    var as = ss.getSheetByName('ASSESSMENT');
    if (!as) return { success: false, error: 'ASSESSMENT sheet not found' };

    var safety = Number(scores && scores.safety || 0);
    var riding = Number(scores && scores.riding || 0);
    var knowledge = Number(scores && scores.knowledge || 0);
    var attitude = Number(scores && scores.attitude || 0);
    var avg = Math.round(((safety + riding + knowledge + attitude) / 4) * 100) / 100;
    var passFail = avg >= 2 ? 'Pass' : 'Repeat';

    as.appendRow([new Date(), bookingId, studentId, safety, riding, knowledge, attitude, avg, passFail, 'Submitted via attendance UI']);
    processAssessmentByBooking_2627(bookingId);
    return { success: true, avg: avg, passFail: passFail };
  } catch (err) {
    Logger.log('saveAssessmentFromAttendance error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}

function _findStudentIdByKENo_(keNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var riders = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  var students = ss.getSheetByName('STUDENTS');
  if (!riders || !students) return '';
  var rData = riders.getDataRange().getValues();
  var name = '';
  for (var i = 1; i < rData.length; i++) {
    if (String(rData[i][CONFIG.RIDER_COLS.KE_NO] || '').trim() === String(keNo).trim()) {
      name = String(rData[i][CONFIG.RIDER_COLS.NAME] || '').trim().toLowerCase();
      break;
    }
  }
  if (!name) return '';
  var sData = students.getDataRange().getValues();
  for (var j = 1; j < sData.length; j++) {
    if (String(sData[j][1] || '').trim().toLowerCase() === name) return String(sData[j][0] || '').trim();
  }
  return '';
}

function _findLatestBookingIdForStudent_(bookingsSheet, studentId) {
  if (!bookingsSheet || bookingsSheet.getLastRow() < 2) return '';
  var data = bookingsSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][2] || '').trim() === String(studentId).trim()) return String(data[i][1] || '').trim();
  }
  return '';
}

function _colourAttCell(cell, status) {
  if (status === 'Present')      cell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
  else if (status === 'No-Show') cell.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
  else                           cell.setBackground('#ffffff').setFontColor('#333333').setFontWeight('normal');
}

// ────────────────────────────────────────────────────────────
//  PAYMENT LOOKUPS
// ────────────────────────────────────────────────────────────

function _getPaymentsForRider(ss, keNo) {
  if (!keNo) return [];
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const out  = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][CONFIG.LEDGER_COLS.KE_NO] || '').trim() !== keNo) continue;
      const paymentDateValue = data[i][CONFIG.LEDGER_COLS.PAY_DATE] || data[i][CONFIG.LEDGER_COLS.SENT_AT];
      out.push({
        amount     : Number(data[i][CONFIG.LEDGER_COLS.AMOUNT]) || 0,
        payDate    : fmtDate(paymentDateValue),
        txnRef     : String(data[i][CONFIG.LEDGER_COLS.TXN_REF]    || ''),
        receiptNo  : String(data[i][CONFIG.LEDGER_COLS.RECEIPT_NO] || ''),
        paymentType: String(data[i][CONFIG.LEDGER_COLS.PAYMENT_TYPE] || 'Riding Classes'),
        paidOn     : fmtDateTime(data[i][CONFIG.LEDGER_COLS.SENT_AT] || paymentDateValue),
        filterDate : ymd(paymentDateValue)
      });
    }
    out.reverse(); // newest first
    return out;
  } catch (e) {
    Logger.log('_getPaymentsForRider error: ' + e);
    return [];
  }
}

// ────────────────────────────────────────────────────────────
//  Change 1: Class count — each 30 minutes = 1 class unit
// ────────────────────────────────────────────────────────────

function _countClassesAttended(ss, keNo) {
  if (!keNo) return 0;
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet) return 0;
    const data = sheet.getDataRange().getValues();
    let count  = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][CONFIG.SCHED_COLS.KE_NO] || '').trim() !== keNo) continue;
      if (String(data[i][CONFIG.SCHED_COLS.ATTENDANCE] || '').toLowerCase() === 'present') {
        count += _calculate30MinBlocks(String(data[i][CONFIG.SCHED_COLS.TIME_SLOT] || ''));
      }
    }
    return count;
  } catch (e) { return 0; }
}

// Each 30-minute window = 1 class block
function _calculate30MinBlocks(timeSlot) {
  if (!timeSlot) return 1;
  // 24-hr format: "HH:MM - HH:MM"
  const m24 = timeSlot.match(/(\d{1,2}):(\d{2})\s*[-–]\s*(\d{1,2}):(\d{2})(?!\s*[AaPp])/);
  if (m24) {
    const dur = (parseInt(m24[3]) * 60 + parseInt(m24[4])) - (parseInt(m24[1]) * 60 + parseInt(m24[2]));
    return dur > 0 ? Math.max(1, Math.round(dur / 30)) : 1;
  }
  // AM/PM format: "H:MM AM - H:MM PM"
  const mAP = timeSlot.match(/(\d{1,2}):(\d{2})\s*(AM|PM)\s*[-–]\s*(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (mAP) {
    let sh = parseInt(mAP[1]), sm = parseInt(mAP[2]);
    let eh = parseInt(mAP[4]), em = parseInt(mAP[5]);
    if (mAP[3].toUpperCase() === 'PM' && sh !== 12) sh += 12;
    if (mAP[3].toUpperCase() === 'AM' && sh === 12) sh = 0;
    if (mAP[6].toUpperCase() === 'PM' && eh !== 12) eh += 12;
    if (mAP[6].toUpperCase() === 'AM' && eh === 12) eh = 0;
    const dur = (eh * 60 + em) - (sh * 60 + sm);
    return dur > 0 ? Math.max(1, Math.round(dur / 30)) : 1;
  }
  return 1;
}

function _getNextSession(ss, keNo) {
  if (!keNo) return null;
  try {
    const tz    = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet) return null;
    const data  = sheet.getDataRange().getValues();
    const upcoming = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][CONFIG.SCHED_COLS.KE_NO] || '').trim() !== keNo) continue;
      const att = String(data[i][CONFIG.SCHED_COLS.ATTENDANCE] || '').toLowerCase();
      if (att === 'present' || att === 'no-show') continue;
      const dt = data[i][CONFIG.SCHED_COLS.DATE];
      if (!dt) continue;
      let dStr;
      try { dStr = Utilities.formatDate(new Date(dt), tz, 'yyyy-MM-dd'); } catch(e) { continue; }
      if (dStr >= today) upcoming.push({ dateStr: dStr, row: data[i] });
    }
    if (!upcoming.length) return null;
    upcoming.sort((a, b) => a.dateStr.localeCompare(b.dateStr));
    const r = upcoming[0].row;
    return { date: fmtDate(r[CONFIG.SCHED_COLS.DATE]), timeSlot: r[CONFIG.SCHED_COLS.TIME_SLOT] || '', service: r[CONFIG.SCHED_COLS.SERVICE] || '' };
  } catch (e) { return null; }
}

// ────────────────────────────────────────────────────────────
//  GET SESSIONS FOR A SPECIFIC RIDER  (portal)
// ────────────────────────────────────────────────────────────

function getSessionsForRider(keNo) {
  if (!keNo) return [];
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet) return [];
    const tz   = Session.getScriptTimeZone();
    const data  = sheet.getDataRange().getValues();
    const out   = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][CONFIG.SCHED_COLS.KE_NO] || '').trim() !== keNo) continue;
      const dt = data[i][CONFIG.SCHED_COLS.DATE];
      let dateStr = '', rawDate = '';
      if (dt) {
        try { dateStr = fmtDate(new Date(dt)); rawDate = Utilities.formatDate(new Date(dt), tz, 'yyyy-MM-dd'); } catch(e){}
      }
      var staffNotes = String(data[i][CONFIG.SCHED_COLS.STAFF_NOTES] || '');
      var delivered = (typeof _parseDeliveredClassFromNotes_ === 'function')
        ? _parseDeliveredClassFromNotes_(staffNotes) : null;
      var bookedBy = String(data[i][CONFIG.SCHED_COLS.BOOKED_BY] || '').trim();
      var scoredBy = String(data[i][CONFIG.SCHED_COLS.SCORED_BY] || '').trim();
      var source   = String(data[i][CONFIG.SCHED_COLS.SOURCE] || '').trim();

      var level = delivered ? delivered.level : '';
      var classNumber = delivered ? delivered.classNumber : '';
      var classTitle = delivered ? delivered.title : '';
      var docLink = '';

      // Enrich from curriculum when possible (doc link + fill blanks for future sessions).
      try {
        if (typeof _loadCurriculumEngineCache_ === 'function') {
          var cache = _loadCurriculumEngineCache_();
          if ((!level || !classNumber) && typeof _curriculumProgressFromCache_ === 'function') {
            var cur = _curriculumProgressFromCache_(keNo, cache);
            level = level || cur.currentLevel || '';
            classNumber = classNumber || cur.currentClassNumber || '';
            classTitle = classTitle || cur.currentTitle || '';
          }
          if (level && classNumber) {
            var key = level + '|' + classNumber;
            if (!classTitle && cache.titleByKey) classTitle = cache.titleByKey[key] || '';
            if (cache.docLinkByKey) docLink = cache.docLinkByKey[key] || '';
            if (!docLink && cache.curriculumItems) {
              for (var ci = 0; ci < cache.curriculumItems.length; ci++) {
                var it = cache.curriculumItems[ci];
                if (String(it.level) === String(level) && String(it.classNumber) === String(classNumber)) {
                  docLink = it.docLink || '';
                  break;
                }
              }
            }
          }
        }
      } catch (ce) {}

      out.push({
        rowIndex    : i + 1,
        service     : data[i][CONFIG.SCHED_COLS.SERVICE]      || '',
        date        : dateStr,
        rawDate     : rawDate,
        timeSlot    : data[i][CONFIG.SCHED_COLS.TIME_SLOT]    || '',
        participants: data[i][CONFIG.SCHED_COLS.PARTICIPANTS]  || 1,
        status      : data[i][CONFIG.SCHED_COLS.STATUS]       || '',
        attendance  : data[i][CONFIG.SCHED_COLS.ATTENDANCE]   || '',
        level       : level,
        classNumber : classNumber,
        classTitle  : classTitle,
        docLink     : docLink,
        bookedBy    : bookedBy,
        scoredBy    : scoredBy,
        trainerLabel: (typeof _formatSessionTrainerLabel_ === 'function')
          ? _formatSessionTrainerLabel_(bookedBy, scoredBy, source) : (scoredBy || bookedBy || ''),
        isFuture    : rawDate >= Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')
      });
    }
    out.sort((a, b) => b.rawDate.localeCompare(a.rawDate));
    return out;
  } catch (e) {
    Logger.log('getSessionsForRider error: ' + e);
    return [];
  }
}

// ────────────────────────────────────────────────────────────
//  RESCHEDULE A SESSION  (portal)
// ────────────────────────────────────────────────────────────

function rescheduleSession(keNo, schedRowIndex, newDate, newTime, reason) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet) return { success: false, error: 'Schedule sheet not found' };

    const rowData = sheet.getRange(schedRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (String(rowData[CONFIG.SCHED_COLS.KE_NO] || '').trim() !== String(keNo).trim())
      return { success: false, error: 'KE Number does not match this session.' };

    const att    = String(rowData[CONFIG.SCHED_COLS.ATTENDANCE] || '').toLowerCase();
    const status = String(rowData[CONFIG.SCHED_COLS.STATUS]     || '').toLowerCase();
    if (att === 'present' || status === 'completed' || status === 'cancelled')
      return { success: false, error: 'This session cannot be rescheduled (status: ' + rowData[CONFIG.SCHED_COLS.STATUS] + ').' };

    const newDateObj = new Date(newDate);
    if (isNaN(newDateObj.getTime())) return { success: false, error: 'Invalid new date.' };
    const today = new Date(); today.setHours(0,0,0,0);
    if (newDateObj <= today) return { success: false, error: 'New date must be in the future.' };

    sheet.getRange(schedRowIndex, CONFIG.SCHED_COLS.DATE + 1).setValue(newDateObj).setNumberFormat('dd-MMM-yyyy');
    sheet.getRange(schedRowIndex, CONFIG.SCHED_COLS.TIME_SLOT + 1).setValue(newTime || rowData[CONFIG.SCHED_COLS.TIME_SLOT]);
    sheet.getRange(schedRowIndex, CONFIG.SCHED_COLS.STATUS + 1).setValue('Rescheduled').setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
    sheet.getRange(schedRowIndex, CONFIG.SCHED_COLS.STAFF_NOTES + 1)
      .setValue('Rescheduled: ' + (reason || 'No reason') + ' (was: ' + fmtDate(rowData[CONFIG.SCHED_COLS.DATE]) + ')');

    const calId = String(rowData[CONFIG.SCHED_COLS.CAL_EVENT_ID] || '').trim();
    if (calId) {
      try {
        const ev = CalendarApp.getDefaultCalendar().getEventById(calId);
        if (ev) {
          const tSlot = newTime || String(rowData[CONFIG.SCHED_COLS.TIME_SLOT] || '');
          const tm    = tSlot.match(/(\d+):(\d+)\s*(AM|PM)?/i);
          if (tm) {
            let h = parseInt(tm[1]), m = parseInt(tm[2]);
            if (tm[3] && tm[3].toUpperCase() === 'PM' && h !== 12) h += 12;
            if (tm[3] && tm[3].toUpperCase() === 'AM' && h === 12) h = 0;
            const ns = new Date(newDateObj); ns.setHours(h,m,0,0);
            const ne = new Date(ns); ne.setMinutes(ne.getMinutes() + 30); // 30-min
            ev.setTime(ns, ne);
          }
        }
      } catch (calErr) { Logger.log('Calendar reschedule failed (non-fatal): ' + calErr); }
    }

    try {
      const rider = findRiderByKENo(keNo);
      if (rider) {
        const email = String(rider.row[CONFIG.RIDER_COLS.EMAIL] || '').trim();
        if (email) {
          const name = rider.row[CONFIG.RIDER_COLS.NAME] || '';
          const subjResched = 'Session Rescheduled: ' + fmtDate(newDateObj) + ' (' + keNo + ')';
          try {
            const rs = sendMailKE_(email, subjResched,
              '<div style="font-family:Arial,sans-serif;padding:20px;max-width:540px"><h2 style="color:#1f4e3d">Session Rescheduled</h2><p>Hi <strong>' + name + '</strong>, your session has been moved.</p><p><strong>New Date:</strong> ' + fmtDate(newDateObj) + '<br><strong>New Time:</strong> ' + (newTime || rowData[CONFIG.SCHED_COLS.TIME_SLOT] || 'TBD') + '<br><strong>Reason:</strong> ' + (reason || 'Not specified') + '</p><p>Kings Equestrian Foundation</p></div>',
              {});
            logEmail('Welcome-reschedule', email, '', subjResched, keNo, 'Sent', 'via ' + rs.provider);
          } catch (e) {
            logEmailFailed('Welcome-reschedule', email, '', subjResched, keNo, String(e));
            throw e;
          }
        }
      }
    } catch (mailErr) { Logger.log('Reschedule email failed (non-fatal): ' + mailErr); }

    if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();
    return { success: true, message: 'Session rescheduled to ' + fmtDate(newDateObj) + '!' };
  } catch (err) {
    Logger.log('rescheduleSession error: ' + err);
    return { success: false, error: err.message };
  }
}

// ────────────────────────────────────────────────────────────
//  BOOK MULTIPLE SESSIONS  — Change 4: booking confirmation email
// ────────────────────────────────────────────────────────────

/** Parse portal date as calendar day in script timezone (avoids UTC-only YYYY-MM-DD shifting). */
function _parsePortalSessionDate_(reqDate) {
  if (reqDate instanceof Date && !isNaN(reqDate.getTime())) return reqDate;
  const s = String(reqDate || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    const d = parseInt(m[3], 10);
    return new Date(y, mo, d);
  }
  const dt = new Date(s);
  return dt;
}

function bookMultipleSessions(keNo, sessionRequests) {
  try {
    const rider = findRiderByKENo(keNo);
    if (!rider) return { success: false, error: 'Rider not found for KE No: ' + keNo };

    const r      = rider.row;
    const name   = r[CONFIG.RIDER_COLS.NAME]  || '';
    const email  = r[CONFIG.RIDER_COLS.EMAIL] || '';
    const phone  = String(r[CONFIG.RIDER_COLS.PHONE] || '');
    const added  = [];
    const errors = [];
    Logger.log(JSON.stringify(sessionRequests));

    sessionRequests.forEach((req, idx) => {
      try {
        if (!req.date || !req.service) throw new Error('Date and service are required');
        const newDateObj = _parsePortalSessionDate_(req.date);
        if (isNaN(newDateObj.getTime())) throw new Error('Invalid date');
        // const today = new Date(); today.setHours(0,0,0,0);
        // if (newDateObj <= today) throw new Error('Date must be in the future');
        const riderPax = Number(r[CONFIG.RIDER_COLS.PARTICIPANTS]) || 1;
        addSessionToSchedule({ keNo, name, phone, email, service: req.service, date: newDateObj, timeSlot: req.timeSlot || '', participants: req.participants || 1, source: 'rider-portal', status: 'Scheduled' });
        added.push({ label: fmtDate(newDateObj) + ' — ' + req.service, service: req.service, date: fmtDate(newDateObj), timeSlot: req.timeSlot || '' });
      } catch (e) { errors.push('Request ' + (idx+1) + ': ' + e.message); }
    });

    // Change 4: Send proper booking confirmation email
    if (email && added.length > 0) {
      try { sendBookingConfirmationEmail({ name, email, keNo, added, errors }); }
      catch (mailErr) { Logger.log('Booking email failed (non-fatal): ' + mailErr); }
    }

    // Admin notification
   const adminEmails = getAdminEmails();
if (adminEmails.length && added.length > 0) {
  const adminSessionRows = added.map(a =>
    '<tr style="border-bottom:1px solid #eee">'
    + '<td style="padding:9px 12px">' + (a.service || a.label) + '</td>'
    + '<td style="padding:9px 12px">' + (a.date || '—') + '</td>'
    + '<td style="padding:9px 12px">' + (a.timeSlot || '—') + '</td>'
    + '</tr>'
  ).join('');

  const adminHtmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:0;color:#333">'
    + '<div style="max-width:620px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.1)">'
    + '  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:26px 30px;text-align:center;color:#fff">'
    + '    <img src="https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb"'
+  '   style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px"> '
    + '    <h1 style="margin:0;font-size:22px">New Portal Booking</h1>'
    + '    <p style="margin:6px 0 0;font-size:13px;opacity:.9">Kings Equestrian Foundation — Admin Alert</p>'
    + '  </div>'
    + '  <div style="padding:26px 30px">'
    + '    <p style="font-size:15px">A new booking has been submitted by <strong>' + name + '</strong>.</p>'
    + '    <div style="background:#d4edda;border-left:4px solid #28a745;padding:14px 18px;border-radius:6px;margin:16px 0">'
    + '      <strong style="color:#155724">' + added.length + ' session' + (added.length !== 1 ? 's' : '') + ' booked</strong><br>'
    + '      <span style="font-size:12px;color:#1e7e34">KE No: ' + keNo + '</span>'
    + '    </div>'
    + '    <table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">'
    + '      <thead><tr style="background:#1f4e3d;color:#fff">'
    + '        <th style="padding:9px 12px;text-align:left">Service</th>'
    + '        <th style="padding:9px 12px;text-align:left">Date</th>'
    + '        <th style="padding:9px 12px;text-align:left">Time Slot</th>'
    + '      </tr></thead>'
    + '      <tbody>' + adminSessionRows + '</tbody>'
    + '    </table>'
    + '    <div style="background:#fff8e6;border-left:4px solid #f0a500;padding:13px;border-radius:4px;font-size:12px;color:#7a5000;margin-top:16px">'
    + '      <strong>Action Required:</strong> Please review and confirm this booking in the system.'
    + '    </div>'
    + '  </div>'
    + '  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">'
    + '    ' + emailFooterHtml_()
    + '  </div>'
    + '</div>'
    + '</body></html>';

  // GmailApp.sendEmail(
  //   adminEmails.join(','),
  //   'New Portal Booking: ' + name + ' (' + keNo + ')',
  //   name + ' (' + keNo + ') booked ' + added.length + ' session(s):\n\n' + added.map(a => a.label).join('\n'),
  //   { name: 'Kings Equestrian System', htmlBody: adminHtmlBody }
  // );
}

    return { success: added.length > 0, added: added.length, failed: errors.length, errors, message: added.length + ' session(s) booked successfully!' };
  } catch (err) {
    Logger.log('bookMultipleSessions error: ' + err);
    return { success: false, error: err.message };
  }
}

// ────────────────────────────────────────────────────────────
//  GOOGLE CALENDAR — Change 1: 30-minute events
// ────────────────────────────────────────────────────────────

function _createCalEvent(d) {
  const calendar = CalendarApp.getDefaultCalendar();
  const date     = new Date(d.date);
  if (isNaN(date.getTime())) return null;

  const tSlot = String(d.timeSlot || '').trim();
  const tm    = tSlot.match(/(\d+):(\d+)\s*(AM|PM)?/i);
  if (!tm) return null;

  let h = parseInt(tm[1]), m = parseInt(tm[2]);
  if (tm[3] && tm[3].toUpperCase() === 'PM' && h !== 12) h += 12;
  if (tm[3] && tm[3].toUpperCase() === 'AM' && h === 12) h = 0;

  const start = new Date(date); start.setHours(h, m, 0, 0);
  const end   = new Date(start); end.setMinutes(end.getMinutes() + 30); // 30-min slots

  const pax  = d.participants || 1;
  const desc = 'KE No: ' + d.keNo + '\nService: ' + d.service + '\nParticipants: ' + pax + '\nPhone: ' + d.phone + '\n\n' + emailFooterPlain_();

  const event = calendar.createEvent(
    'KE — ' + d.name + (pax > 1 ? ' ×' + pax : '') + ' (' + d.keNo + ')',
    start, end,
    { description: desc, location: String(CONFIG.BUSINESS_NAME || 'Kings Equestrian') + ', ' + schoolLocationShort_(), guests: d.email || '', sendInvites: false }
  );

  event.removeAllReminders();
  event.addEmailReminder(1440);
  event.addEmailReminder(60);
  event.addPopupReminder(15);

  const svc = String(d.service || '').toLowerCase();
  if (svc.includes('trek'))        event.setColor(CalendarApp.EventColor.GREEN);
  else if (svc.includes('photo'))  event.setColor(CalendarApp.EventColor.YELLOW);
  else if (svc.includes('camp'))   event.setColor(CalendarApp.EventColor.ORANGE);
  else                             event.setColor(CalendarApp.EventColor.CYAN);

  return event.getId();
}

// ============================================================
// OPTIMIZED FUNCTIONS — used only by the Attendance PWA
// All original functions above are completely unchanged.
// These _Fast variants read each sheet once and build lookup
// maps instead of calling the original helpers in a loop.
// ============================================================

/**
 * Drop-in fast replacement for getSessionsForDate, called by the PWA.
 * Reads Schedule + Payments once each instead of N times per session.
 */
function getSessionsForDate_Fast(dateStr) {
  const tz = Session.getScriptTimeZone();
  let target;
  if (dateStr === 'today')         { target = new Date(); }
  else if (dateStr === 'tomorrow') { target = new Date(); target.setDate(target.getDate() + 1); }
  else                             { target = new Date(dateStr); }
  const targetYMD = Utilities.formatDate(target, tz, 'yyyy-MM-dd');

  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const schedSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!schedSheet) return [];
  const schedData = schedSheet.getDataRange().getValues();

  const paySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
  const payData  = paySheet ? paySheet.getDataRange().getValues() : [];

  const payMap = _buildPaymentMap_Fast(payData);
  const attMap = _buildAttendanceMap_Fast(schedData);

  const results = [];
  for (let i = 1; i < schedData.length; i++) {
    const row  = schedData[i];
    const date = row[CONFIG.SCHED_COLS.DATE];
    if (!date) continue;
    let rowYMD;
    try { rowYMD = Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd'); } catch(e) { continue; }
    if (rowYMD !== targetYMD) continue;
    if (String(row[CONFIG.SCHED_COLS.STATUS] || '').toLowerCase() === 'cancelled') continue;

    const keNo = String(row[CONFIG.SCHED_COLS.KE_NO] || '').trim();
    results.push({
      rowIndex        : i + 1,
      keNo            : keNo,
      name            : row[CONFIG.SCHED_COLS.NAME]         || '',
      phone           : String(row[CONFIG.SCHED_COLS.PHONE] || ''),
      email           : row[CONFIG.SCHED_COLS.EMAIL]        || '',
      service         : row[CONFIG.SCHED_COLS.SERVICE]      || '',
      timeSlot        : row[CONFIG.SCHED_COLS.TIME_SLOT]    || '',
      participants    : row[CONFIG.SCHED_COLS.PARTICIPANTS]  || 1,
      status          : row[CONFIG.SCHED_COLS.STATUS]       || '',
      attendance      : row[CONFIG.SCHED_COLS.ATTENDANCE]   || '',
      staffNotes      : row[CONFIG.SCHED_COLS.STAFF_NOTES]  || '',
      source          : row[CONFIG.SCHED_COLS.SOURCE]       || '',
      payments        : payMap[keNo]  || [],
      classesAttended : attMap[keNo]  || 0
    });
  }

  results.sort((a, b) => (a.timeSlot || '').localeCompare(b.timeSlot || ''));
  return results;
}

/**
 * Drop-in fast replacement for getAllRidersWithStats, called by the PWA.
 * Reads Riders, Payments, Schedule once each instead of 3× per rider.
 */
function getAllRidersWithStats_Fast(pre) {
  pre = pre || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ridersSheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!ridersSheet) return [];
  const ridersData = pre.ridersData || ridersSheet.getDataRange().getValues();

  const paySheet  = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
  const payData   = pre.payData || (paySheet ? paySheet.getDataRange().getValues() : []);

  const schedSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  const schedData  = pre.schedData || (schedSheet ? schedSheet.getDataRange().getValues() : []);

  const payMap  = _buildPaymentMap_Fast(payData);
  const attMap  = _buildAttendanceMap_Fast(schedData);
  const nextMap = _buildNextSessionMap_Fast(schedData);

  const results = [];
  for (let i = 1; i < ridersData.length; i++) {
    const row  = ridersData[i];
    const keNo = String(row[CONFIG.RIDER_COLS.KE_NO] || '').trim();
    if (!keNo) continue;

    results.push({
      keNo            : keNo,
      name            : row[CONFIG.RIDER_COLS.NAME]         || '',
      phone           : String(row[CONFIG.RIDER_COLS.PHONE] || ''),
      email           : row[CONFIG.RIDER_COLS.EMAIL]        || '',
      services        : row[CONFIG.RIDER_COLS.SERVICES]     || '',
      participants    : row[CONFIG.RIDER_COLS.PARTICIPANTS]  || 1,
      registeredOn    : row[CONFIG.RIDER_COLS.REGISTERED]   ? fmtDate(new Date(row[CONFIG.RIDER_COLS.REGISTERED])) : '',
      payments        : payMap[keNo]  || [],
      classesAttended : attMap[keNo]  || 0,
      nextSession     : nextMap[keNo] || null
    });
  }

  results.sort((a, b) => a.name.localeCompare(b.name));
  return results;
}

// ── Internal map builders — zero sheet I/O, accept pre-loaded arrays ──

/** keNo → payment[] built from pre-loaded payments data. */
function _buildPaymentMap_Fast(payData) {
  const map = {};
  for (let i = 1; i < payData.length; i++) {
    const keNo = String(payData[i][CONFIG.LEDGER_COLS.KE_NO] || '').trim();
    if (!keNo) continue;
    if (!map[keNo]) map[keNo] = [];
    const paymentDateValue = payData[i][CONFIG.LEDGER_COLS.PAY_DATE] || payData[i][CONFIG.LEDGER_COLS.SENT_AT];
    map[keNo].push({
      amount     : Number(payData[i][CONFIG.LEDGER_COLS.AMOUNT]) || 0,
      payDate    : fmtDate(paymentDateValue),
      txnRef     : String(payData[i][CONFIG.LEDGER_COLS.TXN_REF]    || ''),
      receiptNo  : String(payData[i][CONFIG.LEDGER_COLS.RECEIPT_NO] || ''),
      paymentType: String(payData[i][CONFIG.LEDGER_COLS.PAYMENT_TYPE] || 'Riding Classes'),
      paidOn     : fmtDateTime(payData[i][CONFIG.LEDGER_COLS.SENT_AT] || paymentDateValue),
      filterDate : ymd(paymentDateValue)
    });
  }
  Object.keys(map).forEach(k => map[k].reverse()); // newest first — matches original
  return map;
}

/** keNo → classesAttended built from pre-loaded schedule data. */
function _buildAttendanceMap_Fast(schedData) {
  const map = {};
  for (let i = 1; i < schedData.length; i++) {
    const keNo = String(schedData[i][CONFIG.SCHED_COLS.KE_NO] || '').trim();
    if (!keNo) continue;
    if (String(schedData[i][CONFIG.SCHED_COLS.ATTENDANCE] || '').toLowerCase() === 'present') {
      map[keNo] = (map[keNo] || 0) + _calculate30MinBlocks(String(schedData[i][CONFIG.SCHED_COLS.TIME_SLOT] || ''));
    }
  }
  return map;
}

/** keNo → next upcoming session built from pre-loaded schedule data. */
function _buildNextSessionMap_Fast(schedData) {
  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const best  = {}; // keNo → { dateStr, row } — earliest upcoming per rider

  for (let i = 1; i < schedData.length; i++) {
    const row  = schedData[i];
    const keNo = String(row[CONFIG.SCHED_COLS.KE_NO] || '').trim();
    if (!keNo) continue;
    const att = String(row[CONFIG.SCHED_COLS.ATTENDANCE] || '').toLowerCase();
    if (att === 'present' || att === 'no-show') continue;
    const dt = row[CONFIG.SCHED_COLS.DATE];
    if (!dt) continue;
    let dStr;
    try { dStr = Utilities.formatDate(new Date(dt), tz, 'yyyy-MM-dd'); } catch(e) { continue; }
    if (dStr < today) continue;
    if (!best[keNo] || dStr < best[keNo].dateStr) {
      best[keNo] = { dateStr: dStr, row: row };
    }
  }

  const result = {};
  Object.keys(best).forEach(keNo => {
    const r = best[keNo].row;
    result[keNo] = {
      date     : fmtDate(r[CONFIG.SCHED_COLS.DATE]),
      timeSlot : r[CONFIG.SCHED_COLS.TIME_SLOT] || '',
      service  : r[CONFIG.SCHED_COLS.SERVICE]   || ''
    };
  });
  return result;
}