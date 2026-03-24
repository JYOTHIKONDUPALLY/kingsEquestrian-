// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 4_Schedule.gs
// All sessions (regular + one-time) go into single Schedule sheet
// ============================================================

// ────────────────────────────────────────────────────────────
//  ADD SESSION TO SCHEDULE
// ────────────────────────────────────────────────────────────

function addSessionToSchedule(d) {
  // d: { keNo, name, phone, email, service, date, timeSlot,
  //      participants, source, status }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sheet) { Logger.log('Schedule sheet not found'); return null; }

  const newRow = new Array(13).fill('');
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

  sheet.appendRow(newRow);
  const lr = sheet.getLastRow();
  if (d.date) sheet.getRange(lr, CONFIG.SCHED_COLS.DATE + 1).setNumberFormat('dd-MMM-yyyy');

  // Create calendar event (non-fatal)
  if (d.date && d.timeSlot) {
    try {
      const calId = _createCalEvent(d);
      if (calId) sheet.getRange(lr, CONFIG.SCHED_COLS.CAL_EVENT_ID + 1).setValue(calId);
    } catch (calErr) {
      Logger.log('Calendar event failed (non-fatal): ' + calErr);
    }
  }

  Logger.log('Session added to Schedule: ' + d.keNo + ' | ' + fmtDate(d.date));
  return lr;
}

// ────────────────────────────────────────────────────────────
//  GET SESSIONS FOR DATE  (used by attendance app + daily summary)
// ────────────────────────────────────────────────────────────

function getSessionsForDate(dateStr) {
  const tz = Session.getScriptTimeZone();
  let target;
  if (dateStr === 'today')    { target = new Date(); }
  else if (dateStr === 'tomorrow') { target = new Date(); target.setDate(target.getDate() + 1); }
  else { target = new Date(dateStr); }
  const targetYMD = Utilities.formatDate(target, tz, 'yyyy-MM-dd');

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sheet) return [];

  const data    = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const date = row[CONFIG.SCHED_COLS.DATE];
    if (!date) continue;
    let rowYMD;
    try { rowYMD = Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd'); } catch(e) { continue; }
    if (rowYMD !== targetYMD) continue;

    const status = String(row[CONFIG.SCHED_COLS.STATUS] || '').toLowerCase();
    if (status === 'cancelled') continue;

    const keNo   = String(row[CONFIG.SCHED_COLS.KE_NO] || '').trim();

    // Fetch payment history for this rider
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
      payments    : payments,  // array of payment objects for this rider
      classesAttended: _countClassesAttended(ss, keNo)
    });
  }

  results.sort((a, b) => (a.timeSlot || '').localeCompare(b.timeSlot || ''));
  return results;
}

// ────────────────────────────────────────────────────────────
//  GET ALL RIDERS  (for "All Riders" tab in attendance app)
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

    const payments       = _getPaymentsForRider(ss, keNo);
    const classesAttended = _countClassesAttended(ss, keNo);
    const nextSession    = _getNextSession(ss, keNo);

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
//  SAVE ATTENDANCE  (called by attendance web app)
// ────────────────────────────────────────────────────────────

function saveAttendance(rowIndex, status, note) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet) return { success: false, error: 'Schedule sheet not found' };

    // Write attendance
    const attCell = sheet.getRange(rowIndex, CONFIG.SCHED_COLS.ATTENDANCE + 1);
    attCell.setValue(status);
    _colourAttCell(attCell, status);

    // Update status column too
    const statusCell = sheet.getRange(rowIndex, CONFIG.SCHED_COLS.STATUS + 1);
    if (status === 'Present')  { statusCell.setValue('Completed').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold'); }
    else if (status === 'No-Show') { statusCell.setValue('No-Show').setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold'); }

    // Write staff note
    if (note !== null && note !== undefined) {
      sheet.getRange(rowIndex, CONFIG.SCHED_COLS.STAFF_NOTES + 1).setValue(note);
    }

    // Send attendance ack email when marked Present
    if (status === 'Present') {
      try {
        const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
        sendAttendanceAckEmail({
          name        : rowData[CONFIG.SCHED_COLS.NAME]         || '',
          email       : rowData[CONFIG.SCHED_COLS.EMAIL]        || '',
          keNo        : rowData[CONFIG.SCHED_COLS.KE_NO]        || '',
          service     : rowData[CONFIG.SCHED_COLS.SERVICE]      || '',
          timeSlot    : rowData[CONFIG.SCHED_COLS.TIME_SLOT]    || '',
          participants: rowData[CONFIG.SCHED_COLS.PARTICIPANTS]  || 1
        });
      } catch (mailErr) {
        Logger.log('Attendance ack email failed (non-fatal): ' + mailErr);
      }
    }

    Logger.log('Attendance saved: row ' + rowIndex + ' → ' + status);
    return { success: true };
  } catch (err) {
    Logger.log('saveAttendance error: ' + err);
    return { success: false, error: err.message };
  }
}

function _colourAttCell(cell, status) {
  if (status === 'Present')     cell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
  else if (status === 'No-Show')cell.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
  else                          cell.setBackground('#ffffff').setFontColor('#333333').setFontWeight('normal');
}

// ────────────────────────────────────────────────────────────
//  PAYMENT LOOKUPS (for attendance cards + portal)
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
      out.push({
        amount    : Number(data[i][CONFIG.LEDGER_COLS.AMOUNT]) || 0,
        payDate   : fmtDate(data[i][CONFIG.LEDGER_COLS.PAY_DATE]),
        txnRef    : String(data[i][CONFIG.LEDGER_COLS.TXN_REF]    || ''),
        receiptNo : String(data[i][CONFIG.LEDGER_COLS.RECEIPT_NO] || ''),
        paidOn    : fmtDateTime(data[i][CONFIG.LEDGER_COLS.TIMESTAMP])
      });
    }
    // newest first
    out.reverse();
    return out;
  } catch (e) {
    Logger.log('_getPaymentsForRider error: ' + e);
    return [];
  }
}

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
        const timeSlot = String(data[i][CONFIG.SCHED_COLS.TIME_SLOT] || '');
        const slots = _calculateSlotsFromTimeSlot(timeSlot);
        count += slots;
      }
    }
    return count;
  } catch (e) { return 0; }
}

function _calculateSlotsFromTimeSlot(timeSlot) {
  if (!timeSlot) return 1; // default to 1 if no time slot
  const match = timeSlot.match(/(\d{1,2}):(\d{2})\s*(AM|PM)\s*-\s*(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!match) return 1; // if can't parse, default to 1
  const startH = parseInt(match[1]), startM = parseInt(match[2]), startAP = match[3].toUpperCase();
  const endH = parseInt(match[4]), endM = parseInt(match[5]), endAP = match[6].toUpperCase();
  
  // Convert to 24-hour
  let startMin = startH * 60 + startM;
  if (startAP === 'PM' && startH !== 12) startMin += 12 * 60;
  if (startAP === 'AM' && startH === 12) startMin = startM;
  
  let endMin = endH * 60 + endM;
  if (endAP === 'PM' && endH !== 12) endMin += 12 * 60;
  if (endAP === 'AM' && endH === 12) endMin = endM;
  
  const durationMin = endMin - startMin;
  if (durationMin <= 0) return 1;
  return Math.ceil(durationMin / 30); // each 30 min is a class
}

function _getNextSession(ss, keNo) {
  if (!keNo) return null;
  try {
    const tz    = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet) return null;
    const data  = sheet.getDataRange().getValues();
    let upcoming = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][CONFIG.SCHED_COLS.KE_NO] || '').trim() !== keNo) continue;
      const att    = String(data[i][CONFIG.SCHED_COLS.ATTENDANCE] || '').toLowerCase();
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
    return {
      date    : fmtDate(r[CONFIG.SCHED_COLS.DATE]),
      timeSlot: r[CONFIG.SCHED_COLS.TIME_SLOT] || '',
      service : r[CONFIG.SCHED_COLS.SERVICE]   || ''
    };
  } catch (e) { return null; }
}

// ────────────────────────────────────────────────────────────
//  GET SESSIONS FOR A SPECIFIC RIDER  (for portal)
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
        try {
          dateStr = fmtDate(new Date(dt));
          rawDate = Utilities.formatDate(new Date(dt), tz, 'yyyy-MM-dd');
        } catch(e){}
      }
      out.push({
        rowIndex    : i + 1,
        service     : data[i][CONFIG.SCHED_COLS.SERVICE]      || '',
        date        : dateStr,
        rawDate     : rawDate,
        timeSlot    : data[i][CONFIG.SCHED_COLS.TIME_SLOT]    || '',
        participants: data[i][CONFIG.SCHED_COLS.PARTICIPANTS]  || 1,
        status      : data[i][CONFIG.SCHED_COLS.STATUS]       || '',
        attendance  : data[i][CONFIG.SCHED_COLS.ATTENDANCE]   || '',
        isFuture    : rawDate >= Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')
      });
    }
    // sort by date descending (most recent first)
    out.sort((a, b) => b.rawDate.localeCompare(a.rawDate));
    return out;
  } catch (e) {
    Logger.log('getSessionsForRider error: ' + e);
    return [];
  }
}

// ────────────────────────────────────────────────────────────
//  RESCHEDULE A SESSION  (called from portal)
// ────────────────────────────────────────────────────────────

function rescheduleSession(keNo, schedRowIndex, newDate, newTime, reason) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet) return { success: false, error: 'Schedule sheet not found' };

    const rowData = sheet.getRange(schedRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Verify ownership
    if (String(rowData[CONFIG.SCHED_COLS.KE_NO] || '').trim() !== String(keNo).trim()) {
      return { success: false, error: 'KE Number does not match this session.' };
    }

    // Can't reschedule completed or cancelled
    const att    = String(rowData[CONFIG.SCHED_COLS.ATTENDANCE] || '').toLowerCase();
    const status = String(rowData[CONFIG.SCHED_COLS.STATUS]     || '').toLowerCase();
    if (att === 'present' || status === 'completed' || status === 'cancelled') {
      return { success: false, error: 'This session cannot be rescheduled (status: ' + rowData[CONFIG.SCHED_COLS.STATUS] + ').' };
    }

    const newDateObj = new Date(newDate);
    if (isNaN(newDateObj.getTime())) return { success: false, error: 'Invalid new date.' };
    const today = new Date(); today.setHours(0,0,0,0);
    if (newDateObj <= today)    return { success: false, error: 'New date must be in the future.' };

    // Update
    sheet.getRange(schedRowIndex, CONFIG.SCHED_COLS.DATE + 1)
      .setValue(newDateObj).setNumberFormat('dd-MMM-yyyy');
    sheet.getRange(schedRowIndex, CONFIG.SCHED_COLS.TIME_SLOT + 1).setValue(newTime || rowData[CONFIG.SCHED_COLS.TIME_SLOT]);
    sheet.getRange(schedRowIndex, CONFIG.SCHED_COLS.STATUS + 1)
      .setValue('Rescheduled').setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
    sheet.getRange(schedRowIndex, CONFIG.SCHED_COLS.STAFF_NOTES + 1)
      .setValue('Rescheduled: ' + (reason || 'No reason') + ' (was: ' + fmtDate(rowData[CONFIG.SCHED_COLS.DATE]) + ')');

    // Update calendar event
    const calId = String(rowData[CONFIG.SCHED_COLS.CAL_EVENT_ID] || '').trim();
    if (calId) {
      try {
        const ev = CalendarApp.getDefaultCalendar().getEventById(calId);
        if (ev) {
          const tSlot = newTime || String(rowData[CONFIG.SCHED_COLS.TIME_SLOT] || '');
          const tm    = tSlot.match(/(\d+):(\d+)\s*(AM|PM)/i);
          if (tm) {
            let h = parseInt(tm[1]), m = parseInt(tm[2]);
            if (tm[3].toUpperCase() === 'PM' && h !== 12) h += 12;
            if (tm[3].toUpperCase() === 'AM' && h === 12) h = 0;
            const ns = new Date(newDateObj); ns.setHours(h,m,0,0);
            const ne = new Date(ns); ne.setMinutes(ne.getMinutes() + 60);
            ev.setTime(ns, ne);
          }
        }
      } catch (calErr) { Logger.log('Calendar reschedule failed (non-fatal): ' + calErr); }
    }

    // Notify rider
    try {
      const rider = findRiderByKENo(keNo);
      if (rider) {
        const email = String(rider.row[CONFIG.RIDER_COLS.EMAIL] || '').trim();
        if (email) {
          const name = rider.row[CONFIG.RIDER_COLS.NAME] || '';
          GmailApp.sendEmail(
            email,
            'Session Rescheduled: ' + fmtDate(newDateObj) + ' (' + keNo + ')',
            'Your session has been rescheduled to ' + fmtDate(newDateObj) + '. New time: ' + (newTime || rowData[CONFIG.SCHED_COLS.TIME_SLOT] || 'TBD'),
            {
              htmlBody : '<div style="font-family:Arial,sans-serif;padding:20px;max-width:540px">'
                + '<h2 style="color:#1f4e3d">Session Rescheduled</h2>'
                + '<p>Hi <strong>' + name + '</strong>, your session has been moved.</p>'
                + '<p><strong>New Date:</strong> ' + fmtDate(newDateObj) + '<br>'
                + '<strong>New Time:</strong> ' + (newTime || rowData[CONFIG.SCHED_COLS.TIME_SLOT] || 'TBD') + '<br>'
                + '<strong>Reason:</strong> ' + (reason || 'Not specified') + '</p>'
                + '<p>Kings Equestrian Foundation</p></div>',
              name : 'Kings Equestrian Foundation'
            }
          );
        }
      }
    } catch (mailErr) { Logger.log('Reschedule email failed (non-fatal): ' + mailErr); }

    Logger.log('Session rescheduled: row ' + schedRowIndex + ' → ' + fmtDate(newDateObj));
    return { success: true, message: 'Session rescheduled to ' + fmtDate(newDateObj) + '!' };
  } catch (err) {
    Logger.log('rescheduleSession error: ' + err);
    return { success: false, error: err.message };
  }
}

// ────────────────────────────────────────────────────────────
//  BOOK MULTIPLE SESSIONS  (called from portal)
// ────────────────────────────────────────────────────────────

function bookMultipleSessions(keNo, sessionRequests) {
  // sessionRequests: [{ service, date, timeSlot, participants }]
  try {
    const rider = findRiderByKENo(keNo);
    if (!rider) return { success: false, error: 'Rider not found for KE No: ' + keNo };

    const r       = rider.row;
    const name    = r[CONFIG.RIDER_COLS.NAME]  || '';
    const email   = r[CONFIG.RIDER_COLS.EMAIL] || '';
    const phone   = String(r[CONFIG.RIDER_COLS.PHONE] || '');
    const added   = [];
    const errors  = [];

    sessionRequests.forEach((req, idx) => {
      try {
        if (!req.date || !req.service) throw new Error('Date and service are required');
        const newDateObj = new Date(req.date);
        if (isNaN(newDateObj.getTime())) throw new Error('Invalid date');
        const today = new Date(); today.setHours(0,0,0,0);
        if (newDateObj <= today) throw new Error('Date must be in the future');

        addSessionToSchedule({
          keNo, name, phone, email,
          service     : req.service,
          date        : newDateObj,
          timeSlot    : req.timeSlot    || '',
          participants: req.participants || 1,
          source      : 'rider-portal',
          status      : 'Scheduled'
        });
        added.push(fmtDate(newDateObj) + ' — ' + req.service);
      } catch (e) {
        errors.push('Request ' + (idx+1) + ': ' + e.message);
      }
    });

    // Send confirmation email
    if (email && added.length > 0) {
      const sessionList = added.map(s => '<li>' + s + '</li>').join('');
      GmailApp.sendEmail(
        email,
        added.length + ' Session(s) Booked - Kings Equestrian (' + keNo + ')',
        added.join("\n"),
        {
          htmlBody : '<div style="font-family:Arial,sans-serif;padding:20px;max-width:560px;color:#333">'
            + '<div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:20px;text-align:center;color:#fff;border-radius:10px 10px 0 0">'
            + '<h2 style="margin:0">Sessions Booked!</h2></div>'
            + '<div style="border:1px solid #e0e0e0;border-top:none;padding:20px;border-radius:0 0 10px 10px">'
            + '<p>Hi <strong>' + name + '</strong>, your sessions have been booked:</p>'
            + '<ul style="font-size:13px;line-height:2">' + sessionList + '</ul>'
            + '<p style="font-size:13px;color:#555">Please ensure your payments are up to date.</p>'
            + '<p><a href="' + CONFIG.PAYMENT_FORM_LINK + '" style="background:#1f4e3d;color:#fff;padding:10px 20px;text-decoration:none;border-radius:6px;font-size:13px">Pay Now</a></p>'
            + (errors.length ? '<p style="color:#c62828;font-size:12px">Some requests had issues: ' + errors.join(', ') + '</p>' : '')
            + '</div></div>',
          name : 'Kings Equestrian Foundation'
        }
      );
    }

    // Notify admin
    const adminEmails = getAdminEmails();
    if (adminEmails.length && added.length > 0) {
      GmailApp.sendEmail(
        adminEmails.join(','),
        'New Portal Booking: ' + name + ' (' + keNo + ')',
        name + ' (' + keNo + ') has booked ' + added.length + ' session(s) via the Rider Portal:\n\n' + added.join('\n'),
        { name : 'Kings Equestrian System' }
      );
    }

    return {
      success: added.length > 0,
      added: added.length,
      failed: errors.length,
      errors,
      message: added.length + ' session(s) booked successfully!'
    };
  } catch (err) {
    Logger.log('bookMultipleSessions error: ' + err);
    return { success: false, error: err.message };
  }
}

// ────────────────────────────────────────────────────────────
//  GOOGLE CALENDAR
// ────────────────────────────────────────────────────────────

function _createCalEvent(d) {
  const calendar = CalendarApp.getDefaultCalendar();
  const date     = new Date(d.date);
  if (isNaN(date.getTime())) return null;

  const tSlot   = String(d.timeSlot || '').trim();
  const tm      = tSlot.match(/(\d+):(\d+)\s*(AM|PM)/i);
  if (!tm) return null;

  let h = parseInt(tm[1]), m = parseInt(tm[2]);
  if (tm[3].toUpperCase() === 'PM' && h !== 12) h += 12;
  if (tm[3].toUpperCase() === 'AM' && h === 12) h = 0;

  const start = new Date(date); start.setHours(h, m, 0, 0);
  const end   = new Date(start); end.setMinutes(end.getMinutes() + 60);

  const pax    = d.participants || 1;
  const desc   = 'KE No: ' + d.keNo + '\nService: ' + d.service
    + '\nParticipants: ' + pax + '\nPhone: ' + d.phone
    + '\n\nKings Equestrian Foundation | Karnataka | +91-9980895533';

  const event  = calendar.createEvent(
    'KE — ' + d.name + (pax > 1 ? ' ×' + pax : '') + ' (' + d.keNo + ')',
    start, end,
    { description: desc, location: 'Kings Equestrian Foundation, Karnataka', guests: d.email || '', sendInvites: false }
  );

  event.removeAllReminders();
  event.addEmailReminder(1440);  // 24h
  event.addEmailReminder(60);    //  1h
  event.addPopupReminder(30);    // 30 min

  // Color by service type
  const svc = String(d.service || '').toLowerCase();
  if (svc.includes('trek'))   event.setColor(CalendarApp.EventColor.GREEN);
  else if (svc.includes('photo')) event.setColor(CalendarApp.EventColor.YELLOW);
  else if (svc.includes('camp'))  event.setColor(CalendarApp.EventColor.ORANGE);
  else event.setColor(CalendarApp.EventColor.CYAN);

  return event.getId();
}