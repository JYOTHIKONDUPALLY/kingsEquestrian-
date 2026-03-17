// ============================================================
// KINGS EQUESTRIAN — UNIVERSAL RIDER PORTAL
// File: 7_RiderPortal.gs
// ============================================================
// HOW IT WORKS:
//   - Rider enters phone number
//   - System checks Regular Students sheet first (KER record)
//   - Then checks Booking Form Response for ALL their one-time bookings
//   - If both → shows Regular Dashboard + one-time history section
//   - If regular only → shows full Regular Dashboard
//   - If one-time only → shows booking history with reschedule options
//
// BACKEND FUNCTIONS (called via google.script.run):
//   getStudentData(phone)           → unified data for this phone
//   submitClassBookingRequest(...)  → request a new class slot
//   submitRescheduleRequest(...)    → request reschedule for a slot
// ============================================================

// ─────────────────────────────────────────────────────────────
//  doGetPortal — entry point for the rider portal web app
// ─────────────────────────────────────────────────────────────

function doGetPortal(e) {
    return HtmlService
        .createHtmlOutput(getRiderPortalHtml())
        .setTitle('My Rides · Kings Equestrian')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────────────────────
//  DATA FETCHER — called by the portal via google.script.run
// ─────────────────────────────────────────────────────────────

/**
 * Unified student data lookup by phone number.
 * Returns everything needed to render the portal dashboard.
 *
 * @param {string} phone  Raw phone number entered by the rider
 * @returns {object|null}
 *   {
 *     found: true/false,
 *     riderType: 'regular' | 'one-time' | 'both' | 'not-found',
 *     name: string,
 *     phone: string,
 *     regular: { regNo, program, paymentStatus, amountPaid, totalAmount,
 *                totalClasses, classesDone, classesRemaining,
 *                enrolledOn, status, schedule[] },
 *     oneTimeBookings: [{ reference, services, date, timeSlot,
 *                         paymentStatus, participants, attendance }]
 *   }
 */
function getStudentData(phone) {
    var normalizedPhone = String(phone || '').replace(/\D/g, '').slice(-10);
    if (!normalizedPhone || normalizedPhone.length < 10) {
        return { found: false, riderType: 'not-found', error: 'Invalid phone number' };
    }
     var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = {
        found: false,
        riderType: 'not-found',
        name: '',
        phone: normalizedPhone,
        regular: null,
        oneTimeBookings: []
    };

    // ── 1. Check Regular Students ──────────────────────────────
    var regularData = _getRegularStudentData(ss, normalizedPhone);
    


    // ── 2. Check One-Time Bookings (ALL bookings for this phone) ─
    var oneTimeData = _getAllOneTimeBookings(ss, normalizedPhone);

    if (regularData) {
        result.found    = true;
        result.name     = regularData.name;
        result.regular  = regularData;
        result.riderType = oneTimeData.length > 0 ? 'both' : 'regular';
    }

    if (oneTimeData.length > 0) {
        result.found          = true;
        result.oneTimeBookings = oneTimeData;
        if (!result.name) result.name = oneTimeData[0].name;
        if (result.riderType === 'not-found') result.riderType = 'one-time';
    }

    return result;
}

function _getRegularStudentData(ss, normalizedPhone) {
    var studentSheet = ss.getSheetByName('Regular Students');
    if (!studentSheet) return null;

    var data = studentSheet.getDataRange().getValues();
    var studentRow = null;
    var studentRowIdx = -1;

    for (var i = 1; i < data.length; i++) {
        var rowPhone = String(data[i][STUDENT_COLS_V2.PHONE] || '').replace(/\D/g, '').slice(-10);
        if (rowPhone === normalizedPhone) {
            studentRow    = data[i];
            studentRowIdx = i + 1;
            break;
        }
    }
    if (!studentRow) return null;

    var regNo        = String(studentRow[STUDENT_COLS_V2.REG_NO] || '').trim();
    var totalCls     = Number(studentRow[STUDENT_COLS_V2.TOTAL_CLASSES]) || 0;
    var doneCls      = Number(studentRow[STUDENT_COLS_V2.CLASSES_COMPLETED]) || 0;
    var amtPaid      = Number(studentRow[STUDENT_COLS_V2.AMOUNT_PAID]) || 0;
    var totalAmt     = Number(studentRow[STUDENT_COLS_V2.TOTAL_AMOUNT]) || 0;

    // Build payment progress text
    var balanceDue = Math.max(0, totalAmt - amtPaid);

    // Load schedule
    var schedule = _getRegularScheduleForStudent(ss, regNo);

    return {
        rowIndex         : studentRowIdx,
        regNo            : regNo,
        name             : studentRow[STUDENT_COLS_V2.NAME] || '',
        email            : studentRow[STUDENT_COLS_V2.EMAIL] || '',
        phone            : studentRow[STUDENT_COLS_V2.PHONE] || '',
        program          : studentRow[STUDENT_COLS_V2.PROGRAM] || '',
        participants     : Number(studentRow[STUDENT_COLS_V2.PARTICIPANTS]) || 1,
        totalClasses     : totalCls,
        classesDone      : doneCls,
        classesRemaining : Math.max(0, totalCls - doneCls),
        paymentStatus    : studentRow[STUDENT_COLS_V2.PAYMENT_STATUS] || 'Pending',
        amountPaid       : amtPaid,
        totalAmount      : totalAmt,
        balanceDue       : balanceDue,
        enrolledOn       : studentRow[STUDENT_COLS_V2.ENROLLED_ON] ? _fmtDate(new Date(studentRow[STUDENT_COLS_V2.ENROLLED_ON])) : '',
        status           : studentRow[STUDENT_COLS_V2.STATUS] || 'Active',
        schedule         : schedule
    };
}

function _getRegularScheduleForStudent(ss, regNo) {
    var schedSheet = ss.getSheetByName('Regular Schedule');
    if (!schedSheet) return [];

    var tz   = Session.getScriptTimeZone();
    var data = schedSheet.getDataRange().getValues();
    var rows = [];

    for (var i = 1; i < data.length; i++) {
        if (String(data[i][SCHED_COLS_V2.REG_NO] || '').trim() !== regNo) continue;

        var schedDate = data[i][SCHED_COLS_V2.SCHEDULED_DATE];
        var dateStr   = '';
        try { dateStr = _fmtDate(new Date(schedDate)); } catch(e){}

        var headers = schedSheet.getRange(1, 1, 1, schedSheet.getLastColumn()).getValues()[0];
        var attIdx  = headers.indexOf('Attendance');
        var attendance = attIdx >= 0 ? (data[i][attIdx] || '') : '';

        rows.push({
            rowIndex         : i + 1,
            classNo          : data[i][SCHED_COLS_V2.CLASS_NO] || '',
            scheduledDate    : dateStr,
            scheduledDateRaw : schedDate ? Utilities.formatDate(new Date(schedDate), tz, 'yyyy-MM-dd') : '',
            timeSlot         : data[i][SCHED_COLS_V2.TIME_SLOT] || '',
            status           : data[i][SCHED_COLS_V2.STATUS] || '',
            attendance       : attendance,
            rescheduleCount  : Number(data[i][SCHED_COLS_V2.RESCHEDULE_COUNT]) || 0,
            originalDate     : data[i][SCHED_COLS_V2.ORIGINAL_DATE] ? _fmtDate(new Date(data[i][SCHED_COLS_V2.ORIGINAL_DATE])) : '',
            notes            : data[i][SCHED_COLS_V2.NOTES] || ''
        });
    }

    // Sort by date
    rows.sort(function(a, b) { return (a.scheduledDateRaw || '').localeCompare(b.scheduledDateRaw || ''); });
    return rows;
}

function _getAllOneTimeBookings(ss, normalizedPhone) {
    var bookingSheet  = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    if (!bookingSheet) return [];

    // Build paid refs
    var paidRefs     = new Set();
    var receiptNums  = {};
    var paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    if (paymentSheet) {
        var pData = paymentSheet.getDataRange().getValues();
        for (var p = 1; p < pData.length; p++) {
            if (String(pData[p][CONFIG.PAYMENT_COLS.RECEIPT_SENT] || '').toLowerCase() === 'yes') {
                var pRef = String(pData[p][CONFIG.PAYMENT_COLS.REGISTRATION_NO] || '').trim();
                paidRefs.add(pRef);
                receiptNums[pRef] = String(pData[p][CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_NO] || '');
            }
        }
    }

    var data    = bookingSheet.getDataRange().getValues();
    var headers = bookingSheet.getRange(1, 1, 1, bookingSheet.getLastColumn()).getValues()[0];
    var attIdx  = headers.indexOf('Attendance');
    var results = [];

    for (var i = 1; i < data.length; i++) {
        var row       = data[i];
        var rowPhone  = String(row[CONFIG.BOOKING_COLS.PHONE_NUMBER] || '').replace(/\D/g, '').slice(-10);
        if (rowPhone !== normalizedPhone) continue;

        var ref       = String(row[CONFIG.BOOKING_COLS.REFERENCE] || '').trim();
        var isPaid    = paidRefs.has(ref);
        var prefDate  = row[CONFIG.BOOKING_COLS.PREFERRED_SERVICE_DATE];
        var dateStr   = '';
        try { if (prefDate) dateStr = _fmtDate(new Date(prefDate)); } catch(e){}

        // Is this date in the future?
        var isFuture  = false;
        if (prefDate) {
            try { isFuture = new Date(prefDate) > new Date(); } catch(e){}
        }

        results.push({
            rowIndex     : i + 1,
            name         : row[CONFIG.BOOKING_COLS.NAME] || '',
            reference    : ref,
            services     : row[CONFIG.BOOKING_COLS.OUR_SERVICES] || '',
            date         : dateStr,
            timeSlot     : row[CONFIG.BOOKING_COLS.PREFERRED_TIME_SLOT] || '',
            participants : Number(row[CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS]) || 1,
            paymentStatus: isPaid ? 'Paid' : 'Pending',
            receiptNo    : receiptNums[ref] || '',
            attendance   : attIdx >= 0 ? (row[attIdx] || '') : '',
            isFuture     : isFuture,
            timestamp    : row[CONFIG.BOOKING_COLS.TIMESTAMP] ? _fmtDate(new Date(row[CONFIG.BOOKING_COLS.TIMESTAMP])) : ''
        });
    }

    // Sort newest first
    results.reverse();
    return results;
}

// ─────────────────────────────────────────────────────────────
//  ACTIONS — called by portal via google.script.run
// ─────────────────────────────────────────────────────────────

/**
 * Submits a new class booking request for a regular student.
 * Writes a row to Regular Schedule with status 'Requested'.
 * Notifies admin by email.
 */
function submitClassBookingRequest(phone, date, timeSlot, notes) {
    try {
        var normalizedPhone = String(phone || '').replace(/\D/g, '').slice(-10);
        var ss              = SpreadsheetApp.getActiveSpreadsheet();

        // Find student
        var studentSheet = ss.getSheetByName('Regular Students');
        if (!studentSheet) return { success: false, error: 'Regular Students sheet not found' };

        var sData      = studentSheet.getDataRange().getValues();
        var studentRow = null;
        for (var i = 1; i < sData.length; i++) {
            var rowPhone = String(sData[i][STUDENT_COLS_V2.PHONE] || '').replace(/\D/g, '').slice(-10);
            if (rowPhone === normalizedPhone) { studentRow = sData[i]; break; }
        }
        if (!studentRow) return { success: false, error: 'Student not found. Please check your phone number.' };

        var regNo    = String(studentRow[STUDENT_COLS_V2.REG_NO]).trim();
        var name     = studentRow[STUDENT_COLS_V2.NAME];
        var email    = studentRow[STUDENT_COLS_V2.EMAIL];
        var program  = studentRow[STUDENT_COLS_V2.PROGRAM];
        var totalCls = Number(studentRow[STUDENT_COLS_V2.TOTAL_CLASSES]) || 0;
        var doneCls  = Number(studentRow[STUDENT_COLS_V2.CLASSES_COMPLETED]) || 0;

        // Check remaining classes
        if (totalCls > 0 && doneCls >= totalCls) {
            return { success: false, error: 'You have completed all your classes for this program.' };
        }

        // Count existing scheduled/requested slots
        var schedSheet = ss.getSheetByName('Regular Schedule');
        var bookedCount = 0;
        if (schedSheet) {
            var scData = schedSheet.getDataRange().getValues();
            for (var s = 1; s < scData.length; s++) {
                if (String(scData[s][SCHED_COLS_V2.REG_NO] || '').trim() === regNo) {
                    var st = String(scData[s][SCHED_COLS_V2.STATUS] || '').toLowerCase();
                    if (st === 'scheduled' || st === 'requested' || st === 'rescheduled') bookedCount++;
                }
            }
        }

        // Calculate class number
        var classNo = doneCls + bookedCount + 1;
        if (totalCls > 0 && classNo > totalCls) {
            return { success: false, error: 'You have already booked all remaining class slots.' };
        }

        // Validate date is in future
        var requestedDate = new Date(date);
        if (isNaN(requestedDate.getTime())) return { success: false, error: 'Invalid date selected.' };
        if (requestedDate <= new Date()) return { success: false, error: 'Please select a future date.' };

        // Write to Regular Schedule
        if (!schedSheet) {
            schedSheet = ss.insertSheet('Regular Schedule');
        }

        var newRow = new Array(16).fill('');
        newRow[SCHED_COLS_V2.REG_NO]       = regNo;
        newRow[SCHED_COLS_V2.STUDENT_NAME] = name;
        newRow[SCHED_COLS_V2.EMAIL]        = email;
        newRow[SCHED_COLS_V2.PHONE]        = phone;
        newRow[SCHED_COLS_V2.PROGRAM]      = program;
        newRow[SCHED_COLS_V2.CLASS_NO]     = classNo;
        newRow[SCHED_COLS_V2.SCHEDULED_DATE] = requestedDate;
        newRow[SCHED_COLS_V2.TIME_SLOT]    = timeSlot;
        newRow[SCHED_COLS_V2.STATUS]       = 'Requested';
        newRow[SCHED_COLS_V2.NOTES]        = notes || '';

        schedSheet.appendRow(newRow);
        var newRowIdx = schedSheet.getLastRow();
        schedSheet.getRange(newRowIdx, SCHED_COLS_V2.SCHEDULED_DATE + 1)
            .setNumberFormat('dd-MMM-yyyy');

        // Notify admin
        _notifyAdminClassRequest(name, email, regNo, program, date, timeSlot, classNo, totalCls, notes);

        // Confirm to student
        _sendBookingRequestConfirmation(name, email, regNo, program, date, timeSlot, classNo, totalCls);

        Logger.log('Class booking request: ' + regNo + ' → ' + date + ' ' + timeSlot + ' (Class ' + classNo + ')');
        return { success: true, classNo: classNo, message: 'Your class request has been submitted! We will confirm within 24 hours.' };

    } catch (err) {
        Logger.log('submitClassBookingRequest error: ' + err);
        return { success: false, error: 'Something went wrong. Please try again or contact us.' };
    }
}

/**
 * Submits a reschedule request for an existing Regular Schedule slot.
 * Updates the row status and sends confirmation.
 */
function submitRescheduleRequest(phone, scheduleRowIndex, newDate, newTime, reason) {
    try {
        var ss         = SpreadsheetApp.getActiveSpreadsheet();
        var schedSheet = ss.getSheetByName('Regular Schedule');
        if (!schedSheet) return { success: false, error: 'Schedule sheet not found' };

        var rowData  = schedSheet.getRange(scheduleRowIndex, 1, 1, schedSheet.getLastColumn()).getValues()[0];
        var regNo    = String(rowData[SCHED_COLS_V2.REG_NO] || '').trim();
        var name     = rowData[SCHED_COLS_V2.STUDENT_NAME] || '';
        var email    = rowData[SCHED_COLS_V2.EMAIL] || '';
        var program  = rowData[SCHED_COLS_V2.PROGRAM] || '';
        var classNo  = rowData[SCHED_COLS_V2.CLASS_NO] || '';
        var oldDate  = rowData[SCHED_COLS_V2.SCHEDULED_DATE];
        var oldTime  = rowData[SCHED_COLS_V2.TIME_SLOT] || '';

        // Validate ownership
        var normalizedPhone = String(phone || '').replace(/\D/g, '').slice(-10);
        var rowPhone        = String(rowData[SCHED_COLS_V2.PHONE] || '').replace(/\D/g, '').slice(-10);
        if (rowPhone !== normalizedPhone) {
            return { success: false, error: 'Phone number does not match this booking.' };
        }

        // Validate new date
        var newDateObj = new Date(newDate);
        if (isNaN(newDateObj.getTime())) return { success: false, error: 'Invalid new date.' };
        if (newDateObj <= new Date()) return { success: false, error: 'New date must be in the future.' };

        // Can't reschedule if already completed
        var status = String(rowData[SCHED_COLS_V2.STATUS] || '').toLowerCase();
        if (status === 'completed' || status === 'cancelled') {
            return { success: false, error: 'This class cannot be rescheduled (status: ' + rowData[SCHED_COLS_V2.STATUS] + ').' };
        }

        var reschedCount = (Number(rowData[SCHED_COLS_V2.RESCHEDULE_COUNT]) || 0) + 1;
        var oldDateStr   = oldDate ? _fmtDate(new Date(oldDate)) : 'N/A';

        // Update the schedule row
        schedSheet.getRange(scheduleRowIndex, SCHED_COLS_V2.SCHEDULED_DATE + 1)
            .setValue(newDateObj).setNumberFormat('dd-MMM-yyyy');
        schedSheet.getRange(scheduleRowIndex, SCHED_COLS_V2.TIME_SLOT + 1).setValue(newTime || oldTime);
        schedSheet.getRange(scheduleRowIndex, SCHED_COLS_V2.STATUS + 1)
            .setValue('Rescheduled').setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
        schedSheet.getRange(scheduleRowIndex, SCHED_COLS_V2.RESCHEDULE_COUNT + 1).setValue(reschedCount);
        schedSheet.getRange(scheduleRowIndex, SCHED_COLS_V2.ORIGINAL_DATE + 1)
            .setValue(oldDate || '').setNumberFormat('dd-MMM-yyyy');
        schedSheet.getRange(scheduleRowIndex, SCHED_COLS_V2.REMINDER_SENT + 1).setValue('No'); // reset reminder
        schedSheet.getRange(scheduleRowIndex, SCHED_COLS_V2.NOTES + 1)
            .setValue('Rescheduled #' + reschedCount + ': ' + (reason || 'No reason given'));

        // Update calendar event if exists
        var calEventId = String(rowData[SCHED_COLS_V2.CALENDAR_EVENT_ID] || '').trim();
        if (calEventId) {
            try {
                var event = CalendarApp.getDefaultCalendar().getEventById(calEventId);
                if (event) {
                    var m = (newTime || oldTime).match(/(\d+):(\d+)\s*(AM|PM)/i);
                    if (m) {
                        var h = parseInt(m[1]);
                        var min = parseInt(m[2]);
                        if (m[3].toUpperCase() === 'PM' && h !== 12) h += 12;
                        if (m[3].toUpperCase() === 'AM' && h === 12) h = 0;
                        var newStart = new Date(newDateObj);
                        newStart.setHours(h, min, 0, 0);
                        var newEnd = new Date(newStart);
                        newEnd.setMinutes(newEnd.getMinutes() + 60);
                        event.setTime(newStart, newEnd);
                    }
                }
            } catch (calErr) {
                Logger.log('Calendar update failed (non-fatal): ' + calErr.message);
            }
        }

        // Send confirmation
        _sendRescheduleConfirmation(name, email, regNo, program, classNo, oldDateStr, oldTime, newDate, newTime || oldTime, reason, reschedCount);

        Logger.log('Reschedule: ' + regNo + ' class ' + classNo + ' from ' + oldDateStr + ' to ' + _fmtDate(newDateObj));
        return { success: true, message: 'Class rescheduled successfully! Confirmation sent to your email.' };

    } catch (err) {
        Logger.log('submitRescheduleRequest error: ' + err);
        return { success: false, error: 'Something went wrong. Please contact us.' };
    }
}

/**
 * Submits a reschedule request for a ONE-TIME booking (future booking only).
 * Just sends an email to admin — no automatic change since date comes from form field.
 */
function submitOneTimeRescheduleRequest(phone, reference, currentDate, currentTime, newDate, newTime, reason) {
    try {
        var normalizedPhone = String(phone || '').replace(/\D/g, '').slice(-10);
        var ss              = SpreadsheetApp.getActiveSpreadsheet();
        var bookingSheet    = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
        if (!bookingSheet) return { success: false, error: 'Booking sheet not found' };

        // Find booking row
        var data     = bookingSheet.getDataRange().getValues();
        var foundRow = -1;
        var name = '', email = '';
        for (var i = 1; i < data.length; i++) {
            var rowRef   = String(data[i][CONFIG.BOOKING_COLS.REFERENCE] || '').trim();
            var rowPhone = String(data[i][CONFIG.BOOKING_COLS.PHONE_NUMBER] || '').replace(/\D/g, '').slice(-10);
            if (rowRef === reference && rowPhone === normalizedPhone) {
                foundRow = i + 1;
                name     = data[i][CONFIG.BOOKING_COLS.NAME] || '';
                email    = data[i][CONFIG.BOOKING_COLS.EMAIL_ID] || '';
                break;
            }
        }
        if (foundRow < 0) return { success: false, error: 'Booking not found.' };

        // Notify admin
        var adminEmails = getAdminEmails();
        if (adminEmails.length > 0) {
            var subject = '🔁 Reschedule Request — One-Time Booking (' + reference + ')';
            var body    = 'Reschedule request received:\n\n'
                + 'Name: ' + name + '\nPhone: ' + phone + '\nRef: ' + reference
                + '\n\nCurrent: ' + currentDate + ' ' + currentTime
                + '\nRequested New: ' + newDate + ' ' + newTime
                + '\nReason: ' + (reason || 'Not specified')
                + '\n\nPlease update the booking and confirm with the customer.';
            MailApp.sendEmail({ to: adminEmails.join(','), subject: subject, body: body, name: 'Kings Equestrian Portal' });
        }

        // Confirm to customer
        if (email) {
            var custSubject = '🔁 Reschedule Request Received (' + reference + ')';
            var custBody    = 'Hi ' + name + ',\n\nWe have received your reschedule request:\n\n'
                + 'Current: ' + currentDate + ' ' + currentTime
                + '\nRequested: ' + newDate + ' ' + newTime
                + '\n\nWe will confirm the change within 24 hours.\n\n'
                + 'Kings Equestrian Foundation\n+91 99807 71166';
            MailApp.sendEmail({ to: email, subject: custSubject, body: custBody, name: 'Kings Equestrian Foundation' });
        }

        return { success: true, message: 'Reschedule request sent! We will confirm within 24 hours.' };
    } catch (err) {
        Logger.log('submitOneTimeRescheduleRequest error: ' + err);
        return { success: false, error: 'Something went wrong. Please contact us.' };
    }
}

// ─────────────────────────────────────────────────────────────
//  NOTIFICATION HELPERS
// ─────────────────────────────────────────────────────────────

function _notifyAdminClassRequest(name, email, regNo, program, date, timeSlot, classNo, totalCls, notes) {
    var adminEmails = getAdminEmails();
    if (!adminEmails || !adminEmails.length) return;
    var subject = '📅 New Class Booking Request — ' + name + ' (' + regNo + ')';
    var body    = 'A regular rider has requested a class slot:\n\n'
        + 'Name: ' + name + '\nReg No: ' + regNo + '\nProgram: ' + program
        + '\nDate: ' + date + '\nTime: ' + timeSlot
        + '\nClass No: ' + classNo + ' of ' + totalCls
        + (notes ? '\nNotes: ' + notes : '')
        + '\n\nPlease confirm in the Regular Schedule sheet.\n\nKings Equestrian System';
    try {
        MailApp.sendEmail({ to: adminEmails.join(','), subject: subject, body: body, name: 'Kings Equestrian Portal' });
    } catch (e) { Logger.log('Admin notification failed: ' + e); }
}

function _sendBookingRequestConfirmation(name, email, regNo, program, date, timeSlot, classNo, totalCls) {
    if (!email) return;
    var subject = '📅 Class Booking Request Received — Class ' + classNo + ' of ' + totalCls;
    var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
        + '<body style="font-family:\'Segoe UI\',sans-serif;background:#f0f4f1;margin:0;padding:0">'
        + '<div style="max-width:560px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(4,52,44,0.10)">'
        + '<div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:24px;text-align:center">'
        + '<img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" style="width:60px;height:60px;border-radius:50%;border:2px solid rgba(255,255,255,0.4);display:block;margin:0 auto 10px">'
        + '<h1 style="margin:0;color:#C8EFE3;font-size:18px">Class Request Received!</h1></div>'
        + '<div style="padding:24px">'
        + '<p>Hi <strong>' + name + '</strong>,</p>'
        + '<p style="font-size:13px;color:#5a7a61;">Your class booking request has been received. We will confirm within 24 hours.</p>'
        + '<div style="background:#f0f9f5;border:1px solid #9FE1CB;border-radius:8px;padding:16px;margin:16px 0">'
        + '<div style="font-size:13px;line-height:2">'
        + '<div><strong>Program:</strong> ' + program + '</div>'
        + '<div><strong>Requested Date:</strong> ' + _fmtDate(new Date(date)) + '</div>'
        + '<div><strong>Time Slot:</strong> ' + timeSlot + '</div>'
        + '<div><strong>Class:</strong> ' + classNo + ' of ' + totalCls + '</div>'
        + '</div></div>'
        + '<div style="background:#fff3cd;border-left:4px solid #ffc107;padding:12px;border-radius:4px;font-size:12px;color:#856404">'
        + '⚠️ This is a request, not a confirmed booking. You will receive a confirmation email once admin approves.'
        + '</div></div>'
        + '<div style="background:#1f4e3d;color:rgba(255,255,255,0.8);padding:14px;text-align:center;font-size:11px">'
        + 'Kings Equestrian Foundation · Karnataka · +91 99807 71166</div>'
        + '</div></body></html>';
    try {
        MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody, name: 'Kings Equestrian Foundation' });
    } catch (e) { Logger.log('Booking confirmation email failed: ' + e); }
}

function _sendRescheduleConfirmation(name, email, regNo, program, classNo, oldDate, oldTime, newDate, newTime, reason, reschedCount) {
    if (!email) return;
    var subject  = '🔁 Reschedule Confirmed — Class ' + classNo + ' moved to ' + _fmtDate(new Date(newDate));
    var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
        + '<body style="font-family:\'Segoe UI\',sans-serif;background:#f0f4f1;margin:0;padding:0">'
        + '<div style="max-width:560px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(4,52,44,0.10)">'
        + '<div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:24px;text-align:center">'
        + '<img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" style="width:60px;height:60px;border-radius:50%;border:2px solid rgba(255,255,255,0.4);display:block;margin:0 auto 10px">'
        + '<h1 style="margin:0;color:#C8EFE3;font-size:18px">Class Rescheduled</h1></div>'
        + '<div style="padding:24px">'
        + '<p>Hi <strong>' + name + '</strong>, your class has been rescheduled.</p>'
        + '<table style="width:100%;border-collapse:collapse;font-size:13px;margin:16px 0">'
        + '<tr style="background:#f8d7da"><td style="padding:9px;font-weight:600;color:#721c24">❌ Old Date</td><td style="padding:9px;color:#721c24;text-decoration:line-through">' + oldDate + ' · ' + oldTime + '</td></tr>'
        + '<tr style="background:#d4edda"><td style="padding:9px;font-weight:600;color:#155724">✅ New Date</td><td style="padding:9px;color:#155724;font-weight:bold">' + _fmtDate(new Date(newDate)) + ' · ' + newTime + '</td></tr>'
        + '<tr style="background:#f9f9f9"><td style="padding:9px;color:#666">Reason</td><td style="padding:9px">' + (reason || 'Not specified') + '</td></tr>'
        + '</table>'
        + '<div style="background:#f0f9f5;border-left:4px solid #4caf50;padding:12px;border-radius:4px;font-size:12px;color:#155724">'
        + 'Calendar invite updated automatically. Reschedule #' + reschedCount + ' for this slot.'
        + '</div></div>'
        + '<div style="background:#1f4e3d;color:rgba(255,255,255,0.8);padding:14px;text-align:center;font-size:11px">'
        + 'Kings Equestrian Foundation · Karnataka · +91 99807 71166</div>'
        + '</div></body></html>';
    try {
        MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody, name: 'Kings Equestrian Foundation' });
    } catch (e) { Logger.log('Reschedule confirmation email failed: ' + e); }
}

// ─────────────────────────────────────────────────────────────
//  UTILITY
// ─────────────────────────────────────────────────────────────

function _fmtDate(d) {
    if (!d || isNaN(d.getTime())) return 'N/A';
    try { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy'); }
    catch(e) { return String(d); }
}

// ─────────────────────────────────────────────────────────────
//  RIDER PORTAL HTML — served from RiderPortal.html file
//
//  In Apps Script editor:
//    File → New → HTML file → name it exactly: RiderPortal
//    Paste the contents of RiderPortal.html provided separately.
//
//  This avoids template literal conflicts in .gs files.
// ─────────────────────────────────────────────────────────────

function getRiderPortalHtml() {
    var paymentFormLink = (typeof CONFIG !== 'undefined' && CONFIG.PAYMENT_FORM_LINK)
        ? CONFIG.PAYMENT_FORM_LINK : '#';
    try {
        var t = HtmlService.createTemplateFromFile('RiderPortal');
        t.paymentFormLink = paymentFormLink;
        return t.evaluate().getContent();
    } catch (e) {
        Logger.log('getRiderPortalHtml: could not load RiderPortal.html — ' + e.message);
        return '<html><body style="font-family:sans-serif;padding:2rem;text-align:center">'
            + '<h2>Setup Required</h2>'
            + '<p>Create a file named <strong>RiderPortal</strong> (HTML type) in the Apps Script editor.</p>'
            + '<p>Paste the contents of RiderPortal.html into it, then redeploy.</p>'
            + '</body></html>';
    }
}