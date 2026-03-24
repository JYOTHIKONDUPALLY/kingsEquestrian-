// ============================================================
// KINGS EQUESTRIAN — ATTENDANCE V2
// File: 6_AttendanceV2.gs
// ============================================================
//
// This file EXTENDS the existing attendance app (Features.gs) without
// touching it. It introduces:
//
//   getBookingsForDateV2(dateStr, filter)
//     — merges one-time bookings + regular schedule slots
//     — filter: 'all' | 'one-time' | 'regular'
//
//   saveAttendanceV2(rowIndex, sheetType, status, note)
//     — routes save to correct sheet
//     — 'booking' → Booking Form Response
//     — 'schedule' → Regular Schedule (+ increments classes completed)
//
//   fixMissingCalendarEvents()
//     — menu-callable backfill for bookings where calendar event was
//       never created (because date/time was blank at payment time)
//
// IMPORTANT: Deploy this project as TWO separate Web Apps:
//   Web App 1 (doGet in Features.gs)  → Staff Attendance URL
//   Web App 2 (doGetPortal in 7_RiderPortal.gs) → Rider Portal URL
//
// The attendance app HTML now calls getBookingsForDateV2 instead of
// getBookingsForDate. To use it, re-deploy and point the attendance
// app HTML to call getBookingsForDateV2.
// ============================================================

// ─────────────────────────────────────────────────────────────
//  SCHEDULE SHEET COLUMN INDICES (Regular Schedule)
//  Matches the sheet created by createRegularStudentRecord()
//  A:Reg No | B:Student Name | C:Email | D:Phone | E:Program
//  F:Class No | G:Scheduled Date | H:Time Slot | I:Status
//  J:Calendar Event ID | K:Reminder Sent | L:Reschedule Count
//  M:Original Date | N:Notes | O:Attendance | P:Staff Notes
// ─────────────────────────────────────────────────────────────
const SCHED_COLS_V2 = {
    REG_NO:            0,
    STUDENT_NAME:      1,
    EMAIL:             2,
    PHONE:             3,
    PROGRAM:           4,
    CLASS_NO:          5,
    SCHEDULED_DATE:    6,
    TIME_SLOT:         7,
    STATUS:            8,
    CALENDAR_EVENT_ID: 9,
    REMINDER_SENT:     10,
    RESCHEDULE_COUNT:  11,
    ORIGINAL_DATE:     12,
    NOTES:             13,
    ATTENDANCE:        14,  // written by saveAttendanceV2
    STAFF_NOTES:       15   // written by saveAttendanceV2
};

// Student sheet cols (mirrors Code.gs STUDENT_COLS)
const STUDENT_COLS_V2 = {
    REG_NO:            0,
    NAME:              1,
    EMAIL:             2,
    PHONE:             3,
    PROGRAM:           4,
    PARTICIPANTS:      5,
    TOTAL_CLASSES:     6,
    CLASSES_COMPLETED: 7,
    PAYMENT_STATUS:    8,
    AMOUNT_PAID:       9,
    TOTAL_AMOUNT:      10,
    PAYMENT_REF:       11,
    ENROLLED_ON:       12,
    STATUS:            13
};

// ─────────────────────────────────────────────────────────────
//  MERGED BOOKING FETCH
// ─────────────────────────────────────────────────────────────

/**
 * Returns a unified list of all rider activity for a given date.
 * Merges one-time bookings and regular schedule slots.
 *
 * @param {string} dateStr  'today' | 'tomorrow' | 'YYYY-MM-DD'
 * @param {string} filter   'all' | 'one-time' | 'regular'
 * @returns {Array}  Unified sorted array of booking objects
 */
function getBookingsForDateV2(dateStr, filter) {
    filter = filter || 'all';

    const tz = Session.getScriptTimeZone();
    let target;
    if (dateStr === 'today') {
        target = new Date();
    } else if (dateStr === 'tomorrow') {
        target = new Date();
        target.setDate(target.getDate() + 1);
    } else {
        target = new Date(dateStr);
    }
    const targetStr = Utilities.formatDate(target, tz, 'yyyy-MM-dd');

    const results = [];

    if (filter === 'all' || filter === 'one-time') {
        const oneTime = _getOneTimeBookingsForDate(targetStr, tz);
        results.push.apply(results, oneTime);
    }

    if (filter === 'all' || filter === 'regular') {
        const regular = _getRegularScheduleForDate(targetStr, tz);
        results.push.apply(results, regular);
    }

    // Sort by time slot string (lexicographic is fine for HH:MM AM/PM)
    results.sort(function(a, b) {
        return (a.timeSlot || '').localeCompare(b.timeSlot || '');
    });

    return results;
}

/**
 * Reads one-time bookings from Booking Form Response for a given date.
 * Enriches with payment status. Adds riderType:'one-time'.
 */
function _getOneTimeBookingsForDate(targetDateStr, tz) {
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    if (!bookingSheet) return [];

    // Build paid-references set
    const paidRefs = new Set();
    const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    if (paymentSheet) {
        const pData = paymentSheet.getDataRange().getValues();
        for (var i = 1; i < pData.length; i++) {
            if (String(pData[i][CONFIG.PAYMENT_COLS.RECEIPT_SENT] || '').toLowerCase() === 'yes') {
                paidRefs.add(String(pData[i][CONFIG.PAYMENT_COLS.REGISTRATION_NO] || '').trim());
            }
        }
    }

    const data    = bookingSheet.getDataRange().getValues();
    const headers = bookingSheet.getRange(1, 1, 1, bookingSheet.getLastColumn()).getValues()[0];

    // Ensure Attendance and Staff Notes columns exist
    let attColIndex = headers.indexOf('Attendance');
    if (attColIndex === -1) {
        attColIndex = bookingSheet.getLastColumn();
        bookingSheet.getRange(1, attColIndex + 1).setValue('Attendance');
        headers.push('Attendance');
    }
    let notesColIndex = headers.indexOf('Staff Notes');
    if (notesColIndex === -1) {
        notesColIndex = bookingSheet.getLastColumn();
        bookingSheet.getRange(1, notesColIndex + 1).setValue('Staff Notes');
        headers.push('Staff Notes');
    }

    const results = [];
    for (var j = 1; j < data.length; j++) {
        var row      = data[j];
        var prefDate = row[CONFIG.BOOKING_COLS.PREFERRED_SERVICE_DATE];
        if (!prefDate) continue;

        var rowDateStr;
        try {
            rowDateStr = Utilities.formatDate(new Date(prefDate), tz, 'yyyy-MM-dd');
        } catch (e) { continue; }

        if (rowDateStr !== targetDateStr) continue;

        var ref    = String(row[CONFIG.BOOKING_COLS.REFERENCE] || '').trim();
        var isPaid = paidRefs.has(ref);

        results.push({
            rowIndex      : j + 1,
            sheetType     : 'booking',
            riderType     : 'one-time',
            name          : row[CONFIG.BOOKING_COLS.NAME] || '',
            phone         : String(row[CONFIG.BOOKING_COLS.PHONE_NUMBER] || ''),
            email         : row[CONFIG.BOOKING_COLS.EMAIL_ID] || '',
            services      : row[CONFIG.BOOKING_COLS.OUR_SERVICES] || '',
            timeSlot      : row[CONFIG.BOOKING_COLS.PREFERRED_TIME_SLOT] || '',
            participants  : row[CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS] || 1,
            reference     : ref,
            regNo         : ref,
            paymentStatus : isPaid ? 'Paid' : 'Pending',
            attendance    : row[attColIndex] || '',
            notes         : row[notesColIndex] || '',
            // Regular-specific fields (empty for one-time)
            programName   : '',
            classNo       : '',
            classesDone   : 0,
            classesRemaining: 0,
            totalClasses  : 0
        });
    }
    return results;
}

/**
 * Reads Regular Schedule slots for a given date.
 * Enriches each slot with student payment/class data from Regular Students.
 * Adds riderType:'regular'.
 */
function _getRegularScheduleForDate(targetDateStr, tz) {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var schedSheet    = ss.getSheetByName('Regular Schedule');
    var studentSheet  = ss.getSheetByName('Regular Students');
    if (!schedSheet) return [];

    // Build student lookup map: regNo → student row
    var studentMap = {};
    if (studentSheet) {
        var sData = studentSheet.getDataRange().getValues();
        for (var s = 1; s < sData.length; s++) {
            var sRegNo = String(sData[s][STUDENT_COLS_V2.REG_NO] || '').trim();
            if (sRegNo) studentMap[sRegNo] = sData[s];
        }
    }

    // Ensure Attendance + Staff Notes columns exist in Regular Schedule
    var headers = schedSheet.getRange(1, 1, 1, schedSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('Attendance') === -1) {
        schedSheet.getRange(1, schedSheet.getLastColumn() + 1).setValue('Attendance');
        headers.push('Attendance');
    }
    if (headers.indexOf('Staff Notes') === -1) {
        schedSheet.getRange(1, schedSheet.getLastColumn() + 1).setValue('Staff Notes');
        headers.push('Staff Notes');
    }
    var attIdx   = headers.indexOf('Attendance');
    var notesIdx = headers.indexOf('Staff Notes');

    var data    = schedSheet.getDataRange().getValues();
    var results = [];

    for (var i = 1; i < data.length; i++) {
        var row       = data[i];
        var schedDate = row[SCHED_COLS_V2.SCHEDULED_DATE];
        if (!schedDate) continue;

        var rowDateStr;
        try {
            rowDateStr = Utilities.formatDate(new Date(schedDate), tz, 'yyyy-MM-dd');
        } catch (e) { continue; }

        if (rowDateStr !== targetDateStr) continue;

        // Skip cancelled slots
        var slotStatus = String(row[SCHED_COLS_V2.STATUS] || '').toLowerCase();
        if (slotStatus === 'cancelled') continue;

        var regNo   = String(row[SCHED_COLS_V2.REG_NO] || '').trim();
        var student = studentMap[regNo] || null;

        var amtPaid     = student ? Number(student[STUDENT_COLS_V2.AMOUNT_PAID]) || 0 : 0;
        var totalAmt    = student ? Number(student[STUDENT_COLS_V2.TOTAL_AMOUNT]) || 0 : 0;
        // Derive payment status from amounts so attendance always reflects reality
        var payStatus   = 'Pending';
        if (totalAmt > 0) {
            if (amtPaid >= totalAmt) payStatus = 'Paid';
            else if (amtPaid > 0)    payStatus = 'Partial';
        } else {
            // If total is unknown/0 (e.g., membership), show Paid only if some payment exists
            payStatus = amtPaid > 0 ? 'Paid' : (student ? String(student[STUDENT_COLS_V2.PAYMENT_STATUS] || 'Pending') : 'Pending');
        }
        var totalCls    = student ? Number(student[STUDENT_COLS_V2.TOTAL_CLASSES]) || 0 : 0;
        var doneCls     = student ? Number(student[STUDENT_COLS_V2.CLASSES_COMPLETED]) || 0 : 0;
        var remaining   = Math.max(0, totalCls - doneCls);
        var classNo     = row[SCHED_COLS_V2.CLASS_NO] || '';
        var program     = row[SCHED_COLS_V2.PROGRAM] || '';

        results.push({
            rowIndex          : i + 1,
            sheetType         : 'schedule',
            riderType         : 'regular',
            name              : row[SCHED_COLS_V2.STUDENT_NAME] || '',
            phone             : String(row[SCHED_COLS_V2.PHONE] || ''),
            email             : row[SCHED_COLS_V2.EMAIL] || '',
            services          : program,
            timeSlot          : row[SCHED_COLS_V2.TIME_SLOT] || '',
            participants      : 1,
            reference         : regNo,
            regNo             : regNo,
            paymentStatus     : payStatus,
            amountPaid        : amtPaid,
            totalAmount       : totalAmt,
            attendance        : row[attIdx] || '',
            notes             : row[notesIdx] || '',
            slotStatus        : row[SCHED_COLS_V2.STATUS] || '',
            // Regular-specific
            programName       : program,
            classNo           : classNo + ' of ' + totalCls,
            classesDone       : doneCls,
            classesRemaining  : remaining,
            totalClasses      : totalCls,
            rescheduleCount   : Number(row[SCHED_COLS_V2.RESCHEDULE_COUNT]) || 0
        });
    }
    return results;
}

// ─────────────────────────────────────────────────────────────
//  SAVE ATTENDANCE V2
// ─────────────────────────────────────────────────────────────

/**
 * Saves attendance to the correct sheet based on riderType.
 *
 * @param {number} rowIndex   1-based sheet row
 * @param {string} sheetType  'booking' | 'schedule'
 * @param {string} status     'Present' | 'No-Show' | 'Rescheduled' | ''
 * @param {string} note       Optional staff note
 * @param {string} regNo      KER reg no (only needed for schedule type to increment count)
 */
function saveAttendanceV2(rowIndex, sheetType, status, note, regNo) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();

        if (sheetType === 'schedule') {
            return _saveRegularAttendance(ss, rowIndex, status, note, regNo);
        } else {
            return _saveOneTimeAttendance(ss, rowIndex, status, note);
        }
    } catch (err) {
        Logger.log('saveAttendanceV2 error: ' + err);
        return { success: false, error: err.message };
    }
}

function _saveOneTimeAttendance(ss, rowIndex, status, note) {
    var bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    if (!bookingSheet) return { success: false, error: 'Booking sheet not found' };

    var headers = bookingSheet.getRange(1, 1, 1, bookingSheet.getLastColumn()).getValues()[0];
    var attCol  = headers.indexOf('Attendance');
    if (attCol === -1) {
        attCol = bookingSheet.getLastColumn();
        bookingSheet.getRange(1, attCol + 1).setValue('Attendance');
    }
    var notesCol = headers.indexOf('Staff Notes');
    if (notesCol === -1) {
        notesCol = bookingSheet.getLastColumn();
        bookingSheet.getRange(1, notesCol + 1).setValue('Staff Notes');
    }

    var attCell = bookingSheet.getRange(rowIndex, attCol + 1);
    attCell.setValue(status);
    _colorAttendanceCell(attCell, status);

    if (note !== null && note !== undefined) {
        bookingSheet.getRange(rowIndex, notesCol + 1).setValue(note);
    }

    // Send acknowledgment email when marked Present (existing behaviour)
    if (status === 'Present') {
        try {
            var rowData = bookingSheet.getRange(rowIndex, 1, 1, bookingSheet.getLastColumn()).getValues()[0];
            sendAttendanceAcknowledgmentEmail(rowData);
        } catch (e) {
            Logger.log('Acknowledgment email failed (non-fatal): ' + e);
        }
    }

    Logger.log('One-time attendance saved: row ' + rowIndex + ' → ' + status);
    return { success: true };
}

function _saveRegularAttendance(ss, rowIndex, status, note, regNo) {
    var schedSheet = ss.getSheetByName('Regular Schedule');
    if (!schedSheet) return { success: false, error: 'Regular Schedule sheet not found' };

    var headers = schedSheet.getRange(1, 1, 1, schedSheet.getLastColumn()).getValues()[0];
    var attCol  = headers.indexOf('Attendance');
    if (attCol === -1) {
        attCol = schedSheet.getLastColumn();
        schedSheet.getRange(1, attCol + 1).setValue('Attendance');
    }
    var notesCol = headers.indexOf('Staff Notes');
    if (notesCol === -1) {
        notesCol = schedSheet.getLastColumn();
        schedSheet.getRange(1, notesCol + 1).setValue('Staff Notes');
    }

    var attCell = schedSheet.getRange(rowIndex, attCol + 1);
    attCell.setValue(status);
    _colorAttendanceCell(attCell, status);

    // Also update the STATUS column
    var statusCell = schedSheet.getRange(rowIndex, SCHED_COLS_V2.STATUS + 1);
    if (status === 'Present') {
        statusCell.setValue('Completed').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
        // Increment classes completed for this student
        if (regNo) _incrementClassesCompleted(ss, regNo);

        // Send acknowledgment email to regular rider (non-fatal)
        try {
            var rowData = schedSheet.getRange(rowIndex, 1, 1, schedSheet.getLastColumn()).getValues()[0];
            var tz = Session.getScriptTimeZone();
            var dateLabel = '';
            try {
                if (rowData[SCHED_COLS_V2.SCHEDULED_DATE]) {
                    dateLabel = Utilities.formatDate(new Date(rowData[SCHED_COLS_V2.SCHEDULED_DATE]), tz, 'EEEE, dd MMM yyyy');
                }
            } catch (e) {}

            sendRegularAttendanceAcknowledgmentEmail({
                email: rowData[SCHED_COLS_V2.EMAIL] || '',
                name: rowData[SCHED_COLS_V2.STUDENT_NAME] || '',
                regNo: String(rowData[SCHED_COLS_V2.REG_NO] || regNo || '').trim(),
                program: rowData[SCHED_COLS_V2.PROGRAM] || '',
                classNo: rowData[SCHED_COLS_V2.CLASS_NO] || '',
                scheduledDateLabel: dateLabel,
                timeSlot: rowData[SCHED_COLS_V2.TIME_SLOT] || ''
            });
        } catch (e) {
            Logger.log('Regular acknowledgment email failed (non-fatal): ' + e);
        }
    } else if (status === 'No-Show') {
        statusCell.setValue('No-Show').setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
    }

    if (note !== null && note !== undefined) {
        schedSheet.getRange(rowIndex, notesCol + 1).setValue(note);
    }

    Logger.log('Regular attendance saved: row ' + rowIndex + ' → ' + status + ' (reg: ' + regNo + ')');
    return { success: true };
}

function _colorAttendanceCell(cell, status) {
    if (status === 'Present') {
        cell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    } else if (status === 'No-Show') {
        cell.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
    } else if (status === 'Rescheduled') {
        cell.setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
    } else {
        cell.setBackground('#ffffff').setFontColor('#333333').setFontWeight('normal');
    }
}

function _incrementClassesCompleted(ss, regNo) {
    var studentSheet = ss.getSheetByName('Regular Students');
    if (!studentSheet) return;
    var data = studentSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (String(data[i][STUDENT_COLS_V2.REG_NO] || '').trim() === regNo) {
            var current  = Number(data[i][STUDENT_COLS_V2.CLASSES_COMPLETED]) || 0;
            var total    = Number(data[i][STUDENT_COLS_V2.TOTAL_CLASSES]) || 0;
            var newCount = current + 1;
            studentSheet.getRange(i + 1, STUDENT_COLS_V2.CLASSES_COMPLETED + 1).setValue(newCount);
            if (total > 0 && newCount >= total) {
                studentSheet.getRange(i + 1, STUDENT_COLS_V2.STATUS + 1)
                    .setValue('Completed')
                    .setBackground('#cce5ff')
                    .setFontColor('#004085')
                    .setFontWeight('bold');
                Logger.log('Student ' + regNo + ' completed all ' + total + ' classes.');
            }
            return;
        }
    }
}

// ─────────────────────────────────────────────────────────────
//  CALENDAR BACKFILL FIX
// ─────────────────────────────────────────────────────────────

/**
 * Menu-callable function.
 * Scans Booking Form Response for rows where:
 *   - Preferred Date AND Preferred Time Slot are filled
 *   - Calendar Event ID column is empty
 *   - A payment receipt has been sent (booking is confirmed)
 *
 * Creates the missing calendar events and writes the IDs back.
 *
 * ROOT CAUSE NOTE:
 * Calendar events are created inside sendReceiptForRow() only when
 * BOTH preferredDate AND preferredTimeSlots are present. If either
 * was blank when the payment was processed, no event was created.
 *
 * PERMANENT FIX: Make "Preferred Date" and "Preferred Time Slot"
 * required fields in your booking Google Form.
 */
function fixMissingCalendarEvents() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var bookingSheet  = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    var paymentSheet  = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);

    if (!bookingSheet) { ui.alert('Booking Form sheet not found'); return; }

    // Build set of references with confirmed receipts
    var paidRefs = new Set();
    if (paymentSheet) {
        var pData = paymentSheet.getDataRange().getValues();
        for (var p = 1; p < pData.length; p++) {
            if (String(pData[p][CONFIG.PAYMENT_COLS.RECEIPT_SENT] || '').toLowerCase() === 'yes') {
                paidRefs.add(String(pData[p][CONFIG.PAYMENT_COLS.REGISTRATION_NO] || '').trim());
            }
        }
    }

    var headers  = bookingSheet.getRange(1, 1, 1, bookingSheet.getLastColumn()).getValues()[0];
    var calColIdx = headers.indexOf('Calendar Event ID');
    if (calColIdx === -1) {
        calColIdx = bookingSheet.getLastColumn();
        bookingSheet.getRange(1, calColIdx + 1).setValue('Calendar Event ID');
        headers.push('Calendar Event ID');
    }

    var data     = bookingSheet.getDataRange().getValues();
    var fixed    = 0;
    var skipped  = 0;
    var errors   = [];

    for (var i = 1; i < data.length; i++) {
        var row          = data[i];
        var ref          = String(row[CONFIG.BOOKING_COLS.REFERENCE] || '').trim();
        var existingCalId = String(row[calColIdx] || '').trim();
        var prefDate     = row[CONFIG.BOOKING_COLS.PREFERRED_SERVICE_DATE];
        var prefTime     = row[CONFIG.BOOKING_COLS.PREFERRED_TIME_SLOT];

        // Skip if already has calendar event
        if (existingCalId) { skipped++; continue; }
        // Skip if no date or time
        if (!prefDate || !prefTime) { skipped++; continue; }
        // Skip if not paid (no confirmed booking)
        if (ref && !paidRefs.has(ref)) { skipped++; continue; }

        var name         = row[CONFIG.BOOKING_COLS.NAME] || '';
        var email        = row[CONFIG.BOOKING_COLS.EMAIL_ID] || '';
        var phone        = String(row[CONFIG.BOOKING_COLS.PHONE_NUMBER] || '');
        var services     = row[CONFIG.BOOKING_COLS.OUR_SERVICES] || '';
        var participants = Number(row[CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS]) || 1;

        try {
            var calId = createBookingCalendarEvent({
                name: name, email: email, phone: phone,
                services: services, date: prefDate, timeSlots: prefTime,
                reference: ref, participants: participants
            });
            if (calId) {
                bookingSheet.getRange(i + 1, calColIdx + 1).setValue(calId);
                fixed++;
                Logger.log('Fixed calendar event for row ' + (i + 1) + ' ref: ' + ref);
                Utilities.sleep(200); // avoid calendar quota
            }
        } catch (err) {
            errors.push('Row ' + (i + 1) + ': ' + err.message);
            Logger.log('Calendar fix failed row ' + (i + 1) + ': ' + err.message);
        }
    }

    var msg = '✅ Calendar Backfill Complete!\n\n'
        + 'Created: ' + fixed + ' new events\n'
        + 'Skipped: ' + skipped + ' rows\n'
        + (errors.length ? '\nErrors:\n' + errors.slice(0, 5).join('\n') : '');

    if (fixed === 0 && errors.length === 0) {
        msg = '✅ All confirmed bookings with date+time already have calendar events!';
    }

    msg += '\n\n📌 To prevent this going forward:\nMake "Preferred Date" and "Preferred Time Slot" '
         + 'REQUIRED fields in your booking Google Form.';

    ui.alert(msg);
}

// ─────────────────────────────────────────────────────────────
//  ENHANCED ATTENDANCE APP HTML
//  The existing doGet() in Features.gs stays untouched.
//  This new HTML is served by the same doGet but can be switched
//  by changing getAttendanceAppHtml() → getAttendanceAppV2Html()
//  in Features.gs when ready.
//  For now, call getAttendanceAppV2Html() from a new doGet if you
//  want to deploy a second attendance URL.
// ─────────────────────────────────────────────────────────────

/**
 * Returns the enhanced attendance app HTML with:
 *  - Filter bar: All | One-Time | Regular
 *  - Regular rider cards showing program, class no, classes done/remaining
 *  - Different card border colours per rider type
 *  - Richer summary stats
 */
function getAttendanceAppV2Html() {
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="KE Attendance">
<meta name="theme-color" content="#1f4e3d">
<link rel="apple-touch-icon" href="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg">
<title>KE Attendance v2</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#f0f4f0;min-height:100vh}
header{background:linear-gradient(135deg,#1f4e3d,#4f9c7a);color:#fff;padding:14px 18px;display:flex;align-items:center;gap:10px;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.2)}
header img{width:40px;height:40px;border-radius:50%;border:2px solid rgba(255,255,255,.4);flex-shrink:0}
header h1{font-size:17px;font-weight:700;line-height:1.1}
header p{font-size:11px;opacity:.85}
.refresh-btn{background:rgba(255,255,255,.2);border:none;color:#fff;width:34px;height:34px;border-radius:50%;cursor:pointer;font-size:17px;margin-left:auto;flex-shrink:0;display:flex;align-items:center;justify-content:center}
.refresh-btn:active{background:rgba(255,255,255,.35)}

/* Date tabs */
.date-tabs{display:flex;background:#fff;border-bottom:2px solid #e0e0e0;position:sticky;top:68px;z-index:99}
.dtab{flex:1;padding:11px 6px;text-align:center;font-size:12px;font-weight:600;color:#666;cursor:pointer;border-bottom:3px solid transparent;transition:all .2s}
.dtab.active{color:#1f4e3d;border-bottom-color:#1f4e3d;background:#f9fffe}

/* Filter bar */
.filter-bar{display:flex;gap:6px;padding:10px 14px;background:#fff;border-bottom:1px solid #e8eee8}
.filter-btn{padding:5px 14px;border:1.5px solid #c5dece;border-radius:20px;background:#fff;font-size:12px;font-weight:600;color:#5a7a61;cursor:pointer;transition:all .15s}
.filter-btn.active{background:#1f4e3d;color:#fff;border-color:#1f4e3d}

/* Summary bar */
.summary-bar{background:#1f4e3d;color:#fff;padding:12px 16px;display:flex;gap:10px;flex-wrap:wrap;overflow-x:auto}
.stat{text-align:center;min-width:50px;flex:1}
.stat-num{font-size:20px;font-weight:700}
.stat-lbl{font-size:10px;opacity:.8;margin-top:1px}
.stat.regular-stat{background:rgba(255,255,255,.1);border-radius:8px;padding:6px 8px}

/* Content */
.content{padding:12px;max-width:720px;margin:0 auto}

/* Cards */
.card{background:#fff;border-radius:12px;padding:14px;margin-bottom:10px;box-shadow:0 2px 6px rgba(0,0,0,.07);border-left:4px solid #ccc;transition:box-shadow .2s}
.card.present{border-left-color:#28a745}
.card.no-show{border-left-color:#dc3545}
.card.rescheduled{border-left-color:#ffc107}
.card.regular-card{border-left-color:#4f9c7a}
.card.regular-card.present{border-left-color:#28a745}
.card.regular-card.no-show{border-left-color:#dc3545}

/* Rider type badge */
.type-badge{display:inline-flex;align-items:center;gap:4px;font-size:10px;font-weight:700;padding:2px 8px;border-radius:12px;margin-bottom:6px}
.type-one-time{background:#e8f4fd;color:#1565c0;border:1px solid #bbdefb}
.type-regular{background:#e8f5e9;color:#1b5e20;border:1px solid #c8e6c9}

/* Card content */
.card-top{display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:6px}
.card-name{font-size:15px;font-weight:700;color:#1f4e3d}
.card-time{font-size:12px;color:#555;margin:2px 0}
.card-service{font-size:11px;color:#888;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.card-ref{font-size:10px;color:#aaa;margin-top:4px}

/* Payment badge */
.pay-badge{font-size:10px;font-weight:600;padding:3px 8px;border-radius:12px;flex-shrink:0}
.pay-paid{background:#d4edda;color:#155724}
.pay-partial{background:#cce5ff;color:#004085}
.pay-pending{background:#fff3cd;color:#856404}

/* Regular progress row */
.prog-row{display:flex;align-items:center;gap:8px;margin:7px 0;padding:7px 10px;background:#f0f9f5;border-radius:8px;border:1px solid #c8e6c9}
.prog-item{display:flex;flex-direction:column;align-items:center;flex:1;min-width:0}
.prog-val{font-size:16px;font-weight:700;color:#1f4e3d}
.prog-label{font-size:9px;color:#5a7a61;text-transform:uppercase;letter-spacing:.05em;margin-top:1px}
.prog-divider{width:1px;height:28px;background:#c8e6c9}
.prog-class-no{font-size:11px;color:#2e7d32;font-weight:600}

/* Action buttons */
.btn-row{display:flex;gap:6px;margin-top:10px;flex-wrap:wrap}
.btn{flex:1;min-width:60px;padding:8px 4px;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;transition:all .15s;display:flex;align-items:center;justify-content:center;gap:3px}
.btn-present{background:#d4edda;color:#155724}
.btn-present.active,.btn-present:active{background:#28a745;color:#fff}
.btn-noshow{background:#f8d7da;color:#721c24}
.btn-noshow.active,.btn-noshow:active{background:#dc3545;color:#fff}
.btn-resched{background:#fff3cd;color:#856404}
.btn-resched.active,.btn-resched:active{background:#ffc107;color:#333}
.btn-clear{background:#f0f0f0;color:#666;flex:0 0 auto;min-width:38px}
.btn-clear:active{background:#ccc}

/* Note */
.note-area{width:100%;margin-top:8px;padding:8px 10px;border:1px solid #ddd;border-radius:8px;font-size:12px;font-family:inherit;resize:vertical;min-height:46px;display:none}
.note-area.show{display:block}
.save-note-btn{display:none;margin-top:5px;padding:6px 14px;background:#1f4e3d;color:#fff;border:none;border-radius:8px;font-size:12px;cursor:pointer;font-weight:600}
.save-note-btn.show{display:inline-block}

/* Date picker */
.custom-date-wrap{padding:8px 14px;background:#fff;border-bottom:1px solid #e8eee8;display:none}
input[type=date]{border:1px solid #c5dece;border-radius:8px;padding:6px 10px;font-size:12px;color:#333;background:#fff}

/* Empty / loading */
.empty{text-align:center;padding:36px 16px;color:#999}
.empty-icon{font-size:44px;margin-bottom:10px}
.loading{text-align:center;padding:36px 16px;color:#1f4e3d;font-size:13px}

/* Toast */
#toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%) translateY(50px);background:#1f4e3d;color:#fff;padding:9px 20px;border-radius:22px;font-size:12px;font-weight:600;z-index:9999;opacity:0;transition:all .3s cubic-bezier(.34,1.3,.64,1);pointer-events:none;white-space:nowrap;border:1px solid rgba(93,202,165,.2)}
#toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
</style>
</head>
<body>

<header>
  <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="KE">
  <div>
    <h1>KE Attendance</h1>
    <p id="headerDate">Loading…</p>
  </div>
  <button class="refresh-btn" onclick="loadData()" title="Refresh">⟳</button>
</header>

<!-- Date tabs -->
<div class="date-tabs">
  <div class="dtab active" onclick="switchDateTab('today',this)">Today</div>
  <div class="dtab" onclick="switchDateTab('tomorrow',this)">Tomorrow</div>
  <div class="dtab" onclick="switchDateTab('custom',this)">📅 Date</div>
</div>

<!-- Custom date picker -->
<div class="custom-date-wrap" id="customDateWrap">
  <input type="date" id="customDate" onchange="loadData()">
</div>

<!-- Filter bar -->
<div class="filter-bar">
  <button class="filter-btn active" onclick="applyFilter('all',this)">All</button>
  <button class="filter-btn" onclick="applyFilter('one-time',this)">🎯 One-Time</button>
  <button class="filter-btn" onclick="applyFilter('regular',this)">🐴 Regular</button>
</div>

<!-- Summary -->
<div class="summary-bar" id="summaryBar" style="display:none">
  <div class="stat"><div class="stat-num" id="sTotal">0</div><div class="stat-lbl">Total</div></div>
  <div class="stat"><div class="stat-num" id="sPresent">0</div><div class="stat-lbl">Present</div></div>
  <div class="stat"><div class="stat-num" id="sNoShow">0</div><div class="stat-lbl">No-Show</div></div>
  <div class="stat"><div class="stat-num" id="sUnmarked">0</div><div class="stat-lbl">Unmarked</div></div>
  <div class="stat regular-stat"><div class="stat-num" id="sRegular">0</div><div class="stat-lbl">Regular</div></div>
  <div class="stat regular-stat"><div class="stat-num" id="sOneTime">0</div><div class="stat-lbl">One-Time</div></div>
</div>

<div class="content">
  <div id="bookingList"><div class="loading">⏳ Loading…</div></div>
</div>

<div id="toast"></div>

<script>
var currentDateTab = 'today';
var currentFilter  = 'all';
var allBookings    = [];
var displayed      = [];

function switchDateTab(tab, el) {
  currentDateTab = tab;
  document.querySelectorAll('.dtab').forEach(function(t){ t.classList.remove('active'); });
  el.classList.add('active');
  document.getElementById('customDateWrap').style.display = (tab === 'custom') ? 'block' : 'none';
  loadData();
}

function applyFilter(f, el) {
  currentFilter = f;
  document.querySelectorAll('.filter-btn').forEach(function(b){ b.classList.remove('active'); });
  el.classList.add('active');
  renderDisplay();
}

function getDateParam() {
  if (currentDateTab === 'custom') {
    var v = document.getElementById('customDate').value;
    return v || 'today';
  }
  return currentDateTab;
}

function loadData() {
  document.getElementById('bookingList').innerHTML = '<div class="loading">⏳ Loading bookings…</div>';
  document.getElementById('summaryBar').style.display = 'none';
  google.script.run
    .withSuccessHandler(function(data) {
      allBookings = data || [];
      renderDisplay();
    })
    .withFailureHandler(function(err) {
      document.getElementById('bookingList').innerHTML =
        '<div class="empty"><div class="empty-icon">⚠️</div><p>' + err.message + '</p></div>';
    })
    .getBookingsForDateV2(getDateParam(), 'all');
}

function renderDisplay() {
  if (currentFilter === 'one-time') {
    displayed = allBookings.filter(function(b){ return b.riderType === 'one-time'; });
  } else if (currentFilter === 'regular') {
    displayed = allBookings.filter(function(b){ return b.riderType === 'regular'; });
  } else {
    displayed = allBookings.slice();
  }

  updateSummary();

  var list = document.getElementById('bookingList');
  if (!displayed.length) {
    list.innerHTML = '<div class="empty"><div class="empty-icon">🐴</div><p>No bookings found</p></div>';
    return;
  }

  list.innerHTML = displayed.map(function(b, idx) {
    return buildCard(b, idx);
  }).join('');

  updateHeaderDate();
}

function buildCard(b, idx) {
  var attClass = b.attendance === 'Present' ? 'present'
               : b.attendance === 'No-Show'  ? 'no-show'
               : b.attendance === 'Rescheduled' ? 'rescheduled' : '';
  var isReg = b.riderType === 'regular';
  var cardClass = 'card ' + (isReg ? 'regular-card ' : '') + attClass;

  var typeBadge = isReg
    ? '<span class="type-badge type-regular">🐴 Regular Rider</span>'
    : '<span class="type-badge type-one-time">🎯 One-Time</span>';

  var payClass = b.paymentStatus === 'Paid' ? 'pay-paid'
               : b.paymentStatus === 'Partial' ? 'pay-partial' : 'pay-pending';

  var payLabel = b.paymentStatus;
  if (isReg && b.paymentStatus === 'Partial' && (Number(b.amountPaid)||0) > 0 && (Number(b.totalAmount)||0) > 0) {
    payLabel = 'Partial (₹' + Number(b.amountPaid||0).toLocaleString('en-IN') + '/₹' + Number(b.totalAmount||0).toLocaleString('en-IN') + ')';
  }

  var progRow = '';
  if (isReg) {
    progRow = '<div class="prog-row">'
      + '<div class="prog-item"><div class="prog-val">' + b.totalClasses + '</div><div class="prog-label">Total</div></div>'
      + '<div class="prog-divider"></div>'
      + '<div class="prog-item"><div class="prog-val" style="color:#28a745">' + b.classesDone + '</div><div class="prog-label">Done</div></div>'
      + '<div class="prog-divider"></div>'
      + '<div class="prog-item"><div class="prog-val" style="color:#e67e00">' + b.classesRemaining + '</div><div class="prog-label">Left</div></div>'
      + '<div class="prog-divider"></div>'
      + '<div class="prog-item"><div class="prog-class-no">Class ' + b.classNo + '</div><div class="prog-label">This class</div></div>'
      + '</div>';
  }

  var btnBase = 'btn btn-present ' + (b.attendance==='Present'?'active':'');
  var btnNS   = 'btn btn-noshow '  + (b.attendance==='No-Show' ?'active':'');
  var btnRS   = 'btn btn-resched ' + (b.attendance==='Rescheduled'?'active':'');

  return '<div class="' + cardClass + '" id="dcard-' + idx + '">'
    + typeBadge
    + '<div class="card-top">'
    +   '<div style="min-width:0;flex:1">'
    +     '<div class="card-name">' + b.name + (b.participants > 1 ? ' <span style="font-size:12px;color:#888">×' + b.participants + '</span>' : '') + '</div>'
    +     '<div class="card-time">🕐 ' + (b.timeSlot || 'Time TBD') + '</div>'
    +     '<div class="card-service" title="' + b.services + '">' + b.services + '</div>'
    +   '</div>'
    +   '<span class="pay-badge ' + payClass + '">' + payLabel + '</span>'
    + '</div>'
    + progRow
    + '<div class="card-ref">📋 ' + b.regNo + ' &nbsp;|&nbsp; 📞 ' + b.phone + '</div>'
    + '<div class="btn-row">'
    +   '<button class="' + btnBase + '" onclick="markAtt(' + idx + ',&quot;Present&quot;)">✅ Present</button>'
    +   '<button class="' + btnNS   + '" onclick="markAtt(' + idx + ',&quot;No-Show&quot;)">❌ No-Show</button>'
    +   '<button class="' + btnRS   + '" onclick="markAtt(' + idx + ',&quot;Rescheduled&quot;)">🔄 Resched</button>'
    +   '<button class="btn btn-clear" onclick="markAtt(' + idx + ',&quot;&quot;)">✕</button>'
    + '</div>'
    + '<textarea class="note-area ' + (b.notes?'show':'') + '" id="note-' + idx + '" placeholder="Staff note…">' + (b.notes||'') + '</textarea>'
    + '<button class="save-note-btn ' + (b.notes?'show':'') + '" onclick="saveNote(' + idx + ')">💾 Save Note</button>'
    + '</div>';
}

function updateSummary() {
  var total    = displayed.length;
  var present  = displayed.filter(function(b){ return b.attendance==='Present'; }).length;
  var noshow   = displayed.filter(function(b){ return b.attendance==='No-Show'; }).length;
  var unmarked = displayed.filter(function(b){ return !b.attendance; }).length;
  var regular  = allBookings.filter(function(b){ return b.riderType==='regular'; }).length;
  var oneTime  = allBookings.filter(function(b){ return b.riderType==='one-time'; }).length;

  document.getElementById('sTotal').textContent   = total;
  document.getElementById('sPresent').textContent = present;
  document.getElementById('sNoShow').textContent  = noshow;
  document.getElementById('sUnmarked').textContent= unmarked;
  document.getElementById('sRegular').textContent = regular;
  document.getElementById('sOneTime').textContent = oneTime;
  document.getElementById('summaryBar').style.display = 'flex';
}

function markAtt(idx, status) {
  var b = displayed[idx];
  b.attendance = status;

  var card = document.getElementById('dcard-' + idx);
  if (card) {
    var base = 'card ' + (b.riderType==='regular' ? 'regular-card ' : '');
    card.className = base + (status==='Present' ? 'present' : status==='No-Show' ? 'no-show' : status==='Rescheduled' ? 'rescheduled' : '');
  }

  var noteEl  = document.getElementById('note-' + idx);
  var saveBtn = noteEl ? noteEl.nextElementSibling : null;
  if (status && noteEl) { noteEl.classList.add('show'); if(saveBtn) saveBtn.classList.add('show'); }

  google.script.run
    .withSuccessHandler(function(){ showToast(status ? '✅ ' + status : '↩️ Cleared'); updateSummary(); })
    .withFailureHandler(function(e){ showToast('❌ ' + e.message); })
    .saveAttendanceV2(b.rowIndex, b.sheetType, status, null, b.regNo);
}

function saveNote(idx) {
  var b    = displayed[idx];
  var note = document.getElementById('note-' + idx).value;
  google.script.run
    .withSuccessHandler(function(){ showToast('💾 Note saved'); })
    .withFailureHandler(function(e){ showToast('❌ ' + e.message); })
    .saveAttendanceV2(b.rowIndex, b.sheetType, b.attendance, note, b.regNo);
}

function showToast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(function(){ t.classList.remove('show'); }, 2400);
}

function updateHeaderDate() {
  var now  = new Date();
  var opts = { weekday:'long', day:'numeric', month:'short' };
  document.getElementById('headerDate').textContent = now.toLocaleDateString('en-IN', opts);
}

updateHeaderDate();
loadData();
</script>
</body>
</html>`;
}