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
      out.push({
        amount    : Number(data[i][CONFIG.LEDGER_COLS.AMOUNT]) || 0,
        payDate   : fmtDate(data[i][CONFIG.LEDGER_COLS.PAY_DATE]),
        txnRef    : String(data[i][CONFIG.LEDGER_COLS.TXN_REF]    || ''),
        receiptNo : String(data[i][CONFIG.LEDGER_COLS.RECEIPT_NO] || ''),
        paidOn    : fmtDateTime(data[i][CONFIG.LEDGER_COLS.TIMESTAMP])
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
          GmailApp.sendEmail(email, 'Session Rescheduled: ' + fmtDate(newDateObj) + ' (' + keNo + ')', '',
            { htmlBody: '<div style="font-family:Arial,sans-serif;padding:20px;max-width:540px"><h2 style="color:#1f4e3d">Session Rescheduled</h2><p>Hi <strong>' + name + '</strong>, your session has been moved.</p><p><strong>New Date:</strong> ' + fmtDate(newDateObj) + '<br><strong>New Time:</strong> ' + (newTime || rowData[CONFIG.SCHED_COLS.TIME_SLOT] || 'TBD') + '<br><strong>Reason:</strong> ' + (reason || 'Not specified') + '</p><p>Kings Equestrian Foundation</p></div>', name: 'Kings Equestrian Foundation' });
        }
      }
    } catch (mailErr) { Logger.log('Reschedule email failed (non-fatal): ' + mailErr); }

    return { success: true, message: 'Session rescheduled to ' + fmtDate(newDateObj) + '!' };
  } catch (err) {
    Logger.log('rescheduleSession error: ' + err);
    return { success: false, error: err.message };
  }
}

// ────────────────────────────────────────────────────────────
//  BOOK MULTIPLE SESSIONS  — Change 4: booking confirmation email
// ────────────────────────────────────────────────────────────

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

    sessionRequests.forEach((req, idx) => {
      try {
        if (!req.date || !req.service) throw new Error('Date and service are required');
        const newDateObj = new Date(req.date);
        if (isNaN(newDateObj.getTime())) throw new Error('Invalid date');
        const today = new Date(); today.setHours(0,0,0,0);
        if (newDateObj <= today) throw new Error('Date must be in the future');
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
      GmailApp.sendEmail(adminEmails.join(','), 'New Portal Booking: ' + name + ' (' + keNo + ')',
        name + ' (' + keNo + ') booked ' + added.length + ' session(s):\n\n' + added.map(a => a.label).join('\n'),
        { name: 'Kings Equestrian System' });
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
  const desc = 'KE No: ' + d.keNo + '\nService: ' + d.service + '\nParticipants: ' + pax + '\nPhone: ' + d.phone + '\n\nKings Equestrian Foundation | Karnataka | +91-9980895533';

  const event = calendar.createEvent(
    'KE — ' + d.name + (pax > 1 ? ' ×' + pax : '') + ' (' + d.keNo + ')',
    start, end,
    { description: desc, location: 'Kings Equestrian Foundation, Karnataka', guests: d.email || '', sendInvites: false }
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