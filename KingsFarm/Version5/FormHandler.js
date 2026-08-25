// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 2_FormHandlers.gs
// ============================================================

// ────────────────────────────────────────────────────────────
//  BOOKING FORM SUBMIT
// ────────────────────────────────────────────────────────────

function onBookingFormSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEETS.BOOKING_FORM) return;

    const row  = e.range.getRow();
    const vals = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    const name         = String(vals[CONFIG.BOOKING_COLS.NAME]         || '').trim();
    const email        = String(vals[CONFIG.BOOKING_COLS.EMAIL]        || '').trim();
    const phone        = String(vals[CONFIG.BOOKING_COLS.PHONE]        || '').trim();
    const services     = String(vals[CONFIG.BOOKING_COLS.SERVICES]     || '').trim();
    const participants = Number(vals[CONFIG.BOOKING_COLS.PARTICIPANTS]) || 1;
    const prefDate     = vals[CONFIG.BOOKING_COLS.PREF_DATE];
    const prefTime     = String(vals[CONFIG.BOOKING_COLS.PREF_TIME]    || '').trim();
    const bookingDate  = vals[CONFIG.BOOKING_COLS.TIMESTAMP];

    if (!email || !phone || !name) {
      Logger.log('onBookingFormSubmit: missing required fields at row ' + row);
      return;
    }

    // Check if rider already exists
    let existing = findRiderByPhoneAndName(phone, name);
    let keNo;
    let isFirstTime = false;

    if (existing) {
      keNo = String(existing.row[CONFIG.RIDER_COLS.KE_NO] || '').trim();
    } else {
      keNo        = generateKENo();
      isFirstTime = true;
      _createRiderRecord({ keNo, name, email, phone, services, participants });
    }

    // Write KE No back to Booking Form row
    sheet.getRange(row, CONFIG.BOOKING_COLS.KE_NO + 1).setValue(keNo);

    // Add session to Schedule if date + time provided
    if (prefDate && prefTime) {
      addSessionToSchedule({
        keNo, name, phone, email, service: services,
        date: prefDate, timeSlot: prefTime, participants,
        source: 'booking-form'
      });
    }

    // Send welcome email
    const amount  = CONFIG.ADVANCE_BOOKING_AMOUNT;
    const upiLink = createUPILink(amount, keNo);
    const qrCode  = createQRCode(upiLink);

    sendWelcomeEmail({
      name, email, phone, services, participants,
      amount, keNo, upiLink, qrCode,prefTime,
      prefDate, isFirstTime,
      sheet, row
    });

    Logger.log('Booking processed: ' + keNo + ' (' + name + ') isFirstTime=' + isFirstTime);
  } catch (err) {
    Logger.log('onBookingFormSubmit ERROR: ' + err + '\n' + err.stack);
  }
}

// ────────────────────────────────────────────────────────────
//  CREATE RIDER RECORD
// ────────────────────────────────────────────────────────────

function _createRiderRecord(d) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) { Logger.log('Riders sheet not found'); return; }

  const newRow = new Array(8).fill('');
  newRow[CONFIG.RIDER_COLS.KE_NO]        = d.keNo;
  newRow[CONFIG.RIDER_COLS.NAME]         = d.name;
  newRow[CONFIG.RIDER_COLS.EMAIL]        = d.email;
  newRow[CONFIG.RIDER_COLS.PHONE]        = d.phone;
  newRow[CONFIG.RIDER_COLS.SERVICES]     = d.services;
  newRow[CONFIG.RIDER_COLS.PARTICIPANTS] = d.participants;
  newRow[CONFIG.RIDER_COLS.REGISTERED]   = new Date();

  sheet.appendRow(newRow);
  sheet.getRange(sheet.getLastRow(), CONFIG.RIDER_COLS.REGISTERED + 1)
    .setNumberFormat('dd-MMM-yyyy');
  Logger.log('Rider record created: ' + d.keNo);
}

// ────────────────────────────────────────────────────────────
//  PAYMENT FORM SUBMIT
// ────────────────────────────────────────────────────────────

function onPaymentFormSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEETS.PAYMENT_FORM) return;

    const row = e.range.getRow();
    Logger.log('Payment form submit at row ' + row);

    // Auto-verify
    sheet.getRange(row, CONFIG.PAYMENT_COLS.VERIFIED + 1)
      .setValue('Yes')
      .setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');

    Utilities.sleep(300);

    // Duplicate check
    const vals    = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const phone   = String(vals[CONFIG.PAYMENT_COLS.PHONE]   || '').trim();
    const amount  = Number(vals[CONFIG.PAYMENT_COLS.AMOUNT]);
    const payDate = vals[CONFIG.PAYMENT_COLS.PAY_DATE];   // col 5
    const ts      = vals[CONFIG.PAYMENT_COLS.TIMESTAMP];

    if (_isDuplicatePayment(phone, amount, payDate, ts)) {
      Logger.log('Duplicate payment detected for phone ' + phone + '. Skipping.');
      sheet.getRange(row, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
        .setValue('Duplicate - Skipped')
        .setBackground('#fff3cd').setFontColor('#856404');
      return;
    }

    sendReceiptForRow(row);
  } catch (err) {
    Logger.log('onPaymentFormSubmit ERROR: ' + err);
  }
}

function _isDuplicatePayment(phone, amount, payDate, currentTs) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    if (!sheet) return false;
    const data      = sheet.getDataRange().getValues();
    const normPh    = normalisePhone(phone);
    const normDate  = ymd(payDate);
    const normCurTs = ymd(currentTs);
    for (let i = 1; i < data.length; i++) {
      if (ymd(data[i][CONFIG.PAYMENT_COLS.TIMESTAMP]) === normCurTs) continue;
      const rowPh  = normalisePhone(data[i][CONFIG.PAYMENT_COLS.PHONE]);
      const rowAmt = Number(data[i][CONFIG.PAYMENT_COLS.AMOUNT]);
      const rowDt  = ymd(data[i][CONFIG.PAYMENT_COLS.PAY_DATE]);   // col 5
      const sent   = String(data[i][CONFIG.PAYMENT_COLS.RECEIPT_SENT] || '').toLowerCase();
      if (rowPh === normPh && rowAmt === amount && rowDt === normDate && sent === 'yes') return true;
    }
    return false;
  } catch (e) {
    Logger.log('_isDuplicatePayment error: ' + e);
    return false;
  }
}

// ────────────────────────────────────────────────────────────
//  SEND RECEIPT FOR A PAYMENT FORM ROW
// ────────────────────────────────────────────────────────────

function sendReceiptForRow(rowIndex) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
  if (!paymentSheet) throw new Error('Payment Form Response sheet not found');

  const vals = paymentSheet.getRange(rowIndex, 1, 1, paymentSheet.getLastColumn()).getValues()[0];

  // Read all fields using updated PAYMENT_COLS
  const keNo      = String(vals[CONFIG.PAYMENT_COLS.KE_NO]     || '').trim();
  const phone     = String(vals[CONFIG.PAYMENT_COLS.PHONE]     || '').trim();
  const amount    = Number(vals[CONFIG.PAYMENT_COLS.AMOUNT]);
  const screenshot= String(vals[CONFIG.PAYMENT_COLS.SCREENSHOT]|| '').trim(); // col 4
  const txnRef    = String(vals[CONFIG.PAYMENT_COLS.TXN_REF]   || '').trim(); // col 6
  const pan       = String(vals[CONFIG.PAYMENT_COLS.PAN]       || '').trim(); // col 7
  const verified  = String(vals[CONFIG.PAYMENT_COLS.VERIFIED]  || '').toLowerCase(); // col 8
  const timestamp = vals[CONFIG.PAYMENT_COLS.TIMESTAMP]; // col 0

  // Resolve payment date — use PAY_DATE (col 5) if valid, else fall back to form TIMESTAMP
  let payDate = vals[CONFIG.PAYMENT_COLS.PAY_DATE]; // col 5
  if (!payDate || fmtDate(payDate) === '') {
    payDate = timestamp || new Date();
    Logger.log('sendReceiptForRow: PAY_DATE was empty/invalid, using TIMESTAMP instead');
  }

  if (verified !== 'yes') throw new Error('Payment not verified — set Verified to Yes first');
  if (!amount || isNaN(amount)) throw new Error('Amount is missing or invalid');

  // Resolve rider
  let rider = keNo ? findRiderByKENo(keNo) : null;
  if (!rider && phone) rider = findRiderByPhone(phone);
  if (!rider) throw new Error('No rider found for KE No [' + keNo + '] or phone [' + phone + ']');

  const riderRow = rider.row;
  const rName    = String(riderRow[CONFIG.RIDER_COLS.NAME]  || '').trim();
  const rEmail   = String(riderRow[CONFIG.RIDER_COLS.EMAIL] || '').trim();
  const rKeNo    = String(riderRow[CONFIG.RIDER_COLS.KE_NO] || keNo).trim();
  const rPhone   = String(riderRow[CONFIG.RIDER_COLS.PHONE] || phone);

  if (!rEmail) throw new Error('No email found for rider ' + rKeNo);

  const receiptNo  = _generateReceiptNo(rKeNo);
  const receiptPDF = generate80GReceipt(rName, pan, amount, txnRef, receiptNo, payDate);

  // Store in Drive — non-fatal
  let driveInfo = null;
  try { driveInfo = storeReceiptInDrive(receiptPDF, rName, receiptNo); }
  catch (driveErr) { Logger.log('storeReceiptInDrive skipped: ' + driveErr); }

  // Append to Payments Ledger with updated 10-column schema
  try {
    _appendToLedger(ss, {
      keNo      : rKeNo,
      name      : rName,
      phone     : rPhone,
      amount,
      screenshot,           // new — col 4 in ledger
      payDate,              // col 5 in ledger
      scheduleDate : '',    // col 6 — leave blank here; set manually if needed
      txnRef,               // col 7 in ledger
      receiptNo             // col 8 in ledger
    });
  } catch (ledgerErr) { Logger.log('_appendToLedger error: ' + ledgerErr); }

  // Send receipt email
  const ccEmails = getCCRecipients('receipt');
  GmailApp.sendEmail(
    rEmail,
    'Payment Receipt - ' + rName + ' (' + rKeNo + ')',
    '',
    {
      htmlBody    : buildReceiptEmailHTML({ name: rName, keNo: rKeNo, amount, txnRef, payDate, receiptNo }),
      attachments : [receiptPDF],
      cc          : ccEmails.join(','),
      name        : 'Kings Equestrian Foundation'
    }
  );

  // Mark Payment Form row as done
  paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
    .setValue('Yes').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
  paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_AT + 1)
    .setValue(new Date()).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
  paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_NO + 1).setValue(receiptNo);
  if (driveInfo) {
    paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.DRIVE_LINK + 1).setValue(driveInfo.fileUrl);
  }

  Logger.log('Receipt sent: ' + receiptNo + ' to ' + rEmail);
  return true;
}

function _generateReceiptNo(keNo) {
  const ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmm');
  const last4 = String(keNo).replace(/\D/g,'').slice(-4) || '0000';
  return keNo + '/' + last4 + '/' + ts.slice(-4);
}

// ────────────────────────────────────────────────────────────
//  APPEND TO PAYMENTS LEDGER
//  Updated for 10-column schema:
//  [KE_NO, NAME, PHONE, AMOUNT, SCREENSHOT, PAY_DATE,
//   SCHEDULE_DATE, TXN_REF, RECEIPT_NO, SENT_AT]
//   col 0    1     2      3        4        5
//            6         7          8         9
// ────────────────────────────────────────────────────────────

function _appendToLedger(ss, d) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
    if (!sheet) {
      Logger.log('_appendToLedger: Payments Ledger sheet not found — skipping');
      return;
    }

    // 10-element row matching updated LEDGER_COLS
    const row = new Array(10).fill('');
    row[CONFIG.LEDGER_COLS.KE_NO]         = d.keNo         || '';
    row[CONFIG.LEDGER_COLS.NAME]          = d.name         || '';
    row[CONFIG.LEDGER_COLS.PHONE]         = d.phone        || '';
    row[CONFIG.LEDGER_COLS.AMOUNT]        = d.amount       || 0;
    row[CONFIG.LEDGER_COLS.SCREENSHOT]    = d.screenshot   || ''; // col 4
    row[CONFIG.LEDGER_COLS.PAY_DATE]      = d.payDate      || ''; // col 5
    row[CONFIG.LEDGER_COLS.SCHEDULE_DATE] = d.scheduleDate || ''; // col 6
    row[CONFIG.LEDGER_COLS.TXN_REF]       = d.txnRef       || ''; // col 7
    row[CONFIG.LEDGER_COLS.RECEIPT_NO]    = d.receiptNo    || ''; // col 8
    row[CONFIG.LEDGER_COLS.SENT_AT]       = new Date();            // col 9

    sheet.appendRow(row);

    const lr = sheet.getLastRow();

    // Format date columns
    if (d.payDate && fmtDate(d.payDate) !== '') {
      sheet.getRange(lr, CONFIG.LEDGER_COLS.PAY_DATE + 1).setNumberFormat('dd-MMM-yyyy');
    }
    if (d.scheduleDate && fmtDate(d.scheduleDate) !== '') {
      sheet.getRange(lr, CONFIG.LEDGER_COLS.SCHEDULE_DATE + 1).setNumberFormat('dd-MMM-yyyy');
    }
    sheet.getRange(lr, CONFIG.LEDGER_COLS.SENT_AT + 1).setNumberFormat('dd-MMM-yyyy HH:mm');
    sheet.getRange(lr, 1, 1, 10).setBackground('#f0faf5');

    Logger.log('_appendToLedger: row ' + lr + ' for ' + d.keNo + ' — Rs.' + d.amount + ' on ' + fmtDate(d.payDate));
  } catch (e) {
    Logger.log('_appendToLedger ERROR: ' + e + '\n' + e.stack);
  }
}

// ────────────────────────────────────────────────────────────
//  PORTAL PAYMENT SUBMIT (replaces Google Form)
// ────────────────────────────────────────────────────────────

function getPaymentUploadsFolder() {
  let main = DriveApp.getFoldersByName('Kings Equestrian Payment Screenshots');
  main = main.hasNext() ? main.next() : DriveApp.createFolder('Kings Equestrian Payment Screenshots');
  return main;
}

/** Last known PAN/Aadhaar from Payment Form Response for this KE / phone. */
function _lookupPriorPanForRider_(keNo, phone) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    if (!sheet || sheet.getLastRow() < 2) return '';
    const data = sheet.getDataRange().getValues();
    const keNorm = String(keNo || '').trim().toUpperCase();
    const phNorm = normalisePhone(phone);
    for (let i = data.length - 1; i >= 1; i--) {
      const rowKe = String(data[i][CONFIG.PAYMENT_COLS.KE_NO] || '').trim().toUpperCase();
      const rowPh = normalisePhone(data[i][CONFIG.PAYMENT_COLS.PHONE]);
      const pan = String(data[i][CONFIG.PAYMENT_COLS.PAN] || '').trim();
      if (!pan) continue;
      if ((keNorm && rowKe === keNorm) || (phNorm && rowPh === phNorm)) return pan;
    }
  } catch (e) {
    Logger.log('_lookupPriorPanForRider_ ' + e);
  }
  return '';
}

function getPortalPaymentPrefill(keNo) {
  try {
    keNo = String(keNo || '').trim().toUpperCase();
    if (!keNo) return { success: false, pan: '', phone: '' };
    const rider = findRiderByKENo(keNo);
    const phone = rider ? String(rider.row[CONFIG.RIDER_COLS.PHONE] || '').trim() : '';
    const pan = _lookupPriorPanForRider_(keNo, phone);
    return { success: true, keNo: keNo, phone: phone, pan: pan };
  } catch (e) {
    Logger.log('getPortalPaymentPrefill ' + e);
    return { success: false, pan: '', phone: '' };
  }
}

function _savePaymentScreenshots(files, keNo) {
  const folder = getPaymentUploadsFolder();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const urls = [];
  (files || []).forEach(function(f, i) {
    if (!f || !f.data) return;
    const bytes = Utilities.base64Decode(f.data);
    const blob = Utilities.newBlob(bytes, f.mimeType || 'application/octet-stream', f.name || ('screenshot_' + (i + 1)));
    const file = folder.createFile(blob);
    const safeName = String(f.name || 'screenshot').replace(/[\\/:*?"<>|]/g, '_');
    file.setName(keNo + '_' + ts + '_' + (i + 1) + '_' + safeName);
    urls.push(file.getUrl());
  });
  return urls.join('\n');
}

function _ensurePaymentFormSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
  if (sheet) return sheet;
  sheet = ss.insertSheet(CONFIG.SHEETS.PAYMENT_FORM);
  const headers = [
    'Timestamp', 'KE No', 'Phone', 'Amount Paid', 'Screenshot',
    'Payment Date', 'Transaction Ref', 'PAN/Aadhaar', 'Verified',
    'Receipt Sent', 'Receipt Sent At', 'Receipt No', 'Drive Link'
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Called from My Rides. payload:
 * { keNo, phone, amount, payDate (yyyy-mm-dd), txnRef, pan,
 *   files: [{ name, mimeType, data (base64) }] }
 */
function submitPaymentFromPortal(p) {
  try {
    p = p || {};
    const keNo  = String(p.keNo || '').trim().toUpperCase();
    const phone = String(p.phone || '').trim();
    const amount = Number(p.amount);
    const pan    = String(p.pan || '').trim();
    const txnRef = String(p.txnRef || '').trim();
    const payDateStr = String(p.payDate || '').trim();
    const files = Array.isArray(p.files) ? p.files : [];

    if (!keNo) {
      return { success: false, error: 'Please enter your Registration Number provided in your email.' };
    }
    if (!normalisePhone(phone) || normalisePhone(phone).length < 10) {
      return { success: false, error: 'Please enter your registered phone number.' };
    }
    if (!amount || isNaN(amount) || amount <= 0) {
      return { success: false, error: 'Please enter the exact amount paid.' };
    }
    if (!files.length) {
      return { success: false, error: 'Please upload a clear screenshot or receipt of the completed payment.' };
    }
    if (files.length > 5) {
      return { success: false, error: 'You can upload up to 5 files.' };
    }
    if (!pan) {
      return { success: false, error: 'PAN / Aadhaar number is required for issuing official payment receipts.' };
    }

    const rider = findRiderByKENo(keNo);
    if (!rider) {
      return { success: false, error: 'Registration number not found. Please use the KE Number from your email.' };
    }
    if (normalisePhone(rider.row[CONFIG.RIDER_COLS.PHONE]) !== normalisePhone(phone)) {
      return { success: false, error: 'Phone number does not match this Registration Number.' };
    }

    let payDate = '';
    if (payDateStr) {
      const dt = new Date(payDateStr + 'T00:00:00');
      if (!isNaN(dt.getTime())) payDate = dt;
    }

    const screenshot = _savePaymentScreenshots(files, keNo);
    if (!screenshot) {
      return { success: false, error: 'Could not save the screenshot. Please try again with a smaller file.' };
    }

    const sheet = _ensurePaymentFormSheet();
    const row = new Array(13).fill('');
    const now = new Date();
    row[CONFIG.PAYMENT_COLS.TIMESTAMP] = now;
    row[CONFIG.PAYMENT_COLS.KE_NO]     = keNo;
    row[CONFIG.PAYMENT_COLS.PHONE]     = phone;
    row[CONFIG.PAYMENT_COLS.AMOUNT]    = amount;
    row[CONFIG.PAYMENT_COLS.SCREENSHOT]= screenshot;
    row[CONFIG.PAYMENT_COLS.PAY_DATE]  = payDate || '';
    row[CONFIG.PAYMENT_COLS.TXN_REF]   = txnRef;
    row[CONFIG.PAYMENT_COLS.PAN]       = pan;
    sheet.appendRow(row);
    const rowIndex = sheet.getLastRow();

    sheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.VERIFIED + 1)
      .setValue('Yes')
      .setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    if (payDate) {
      sheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.PAY_DATE + 1).setNumberFormat('dd-MMM-yyyy');
    }

    if (_isDuplicatePayment(phone, amount, payDate || now, now)) {
      sheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
        .setValue('Duplicate - Skipped')
        .setBackground('#fff3cd').setFontColor('#856404');
      return {
        success: true,
        message: 'This payment looks like a duplicate of one already receipted. Your submission has been recorded.'
      };
    }

    try {
      sendReceiptForRow(rowIndex);
      return { success: true, message: 'Payment submitted. Your official receipt will be emailed shortly.' };
    } catch (receiptErr) {
      Logger.log('submitPaymentFromPortal receipt: ' + receiptErr);
      return {
        success: true,
        message: 'Payment submitted for verification. If you do not receive a receipt, please contact us at +91-9980895533.'
      };
    }
  } catch (err) {
    Logger.log('submitPaymentFromPortal ERROR: ' + err + '\n' + (err.stack || ''));
    return { success: false, error: 'Could not submit payment. Please try again.' };
  }
}

// ────────────────────────────────────────────────────────────
//  MENU ACTIONS
// ────────────────────────────────────────────────────────────

function sendPaymentReceiptMenu() {
  const ui           = SpreadsheetApp.getUi();
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
  if (!paymentSheet) { ui.alert('Payment Form Response sheet not found'); return; }

  const sel = paymentSheet.getActiveRange();
  if (!sel || sel.getRow() === 1) { ui.alert('Select valid rows (not the header row)'); return; }

  const startRow = sel.getRow();
  const numRows  = sel.getNumRows();
  const resp     = ui.alert('Send receipts for ' + numRows + ' row(s)?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  let ok = 0, fail = 0, errors = [];
  for (let i = 0; i < numRows; i++) {
    try {
      sendReceiptForRow(startRow + i);
      ok++;
      Utilities.sleep(800);
    } catch (err) {
      fail++;
      errors.push('Row ' + (startRow + i) + ': ' + err.message);
    }
  }
  ui.alert('Done!\n' + ok + ' sent\n' + fail + ' failed'
    + (errors.length ? '\n\n' + errors.slice(0, 5).join('\n') : ''));
}

function resendWelcomeEmail() {
  const ui           = SpreadsheetApp.getUi();
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
  if (!bookingSheet) { ui.alert('Booking Form Response sheet not found'); return; }

  const sel     = bookingSheet.getActiveRange();
  if (!sel || sel.getRow() === 1) { ui.alert('Select valid rows'); return; }
  const start   = sel.getRow();
  const numRows = sel.getNumRows();
  const resp    = ui.alert('Resend welcome emails for ' + numRows + ' row(s)?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  let ok = 0, fail = 0;
  for (let i = 0; i < numRows; i++) {
    try {
      const rowIdx       = start + i;
      const vals         = bookingSheet.getRange(rowIdx, 1, 1, bookingSheet.getLastColumn()).getValues()[0];
      const name         = String(vals[CONFIG.BOOKING_COLS.NAME]         || '').trim();
      const email        = String(vals[CONFIG.BOOKING_COLS.EMAIL]        || '').trim();
      const phone        = String(vals[CONFIG.BOOKING_COLS.PHONE]        || '').trim();
      const services     = String(vals[CONFIG.BOOKING_COLS.SERVICES]     || '').trim();
      const participants = Number(vals[CONFIG.BOOKING_COLS.PARTICIPANTS]) || 1;
      const prefTime     = String(vals[CONFIG.BOOKING_COLS.PREF_TIME]    || '').trim();
      const keNo         = String(vals[CONFIG.BOOKING_COLS.KE_NO]        || '').trim();
      const prefDate  = vals[CONFIG.BOOKING_COLS.PREF_DATE];
      if (!email || !keNo) throw new Error('Missing email or KE No');
      const amount  = CONFIG.ADVANCE_BOOKING_AMOUNT;
      const upiLink = createUPILink(amount, keNo);
      const qrCode  = createQRCode(upiLink);
      sendWelcomeEmail({
        name, email, phone, services, participants,
        amount, keNo, upiLink, qrCode,prefTime,
        prefDate, isFirstTime: false,
        sheet: bookingSheet, row: rowIdx
      });
      ok++;
      Utilities.sleep(800);
    } catch (err) {
      fail++;
      Logger.log('resendWelcomeEmail row ' + (start + i) + ': ' + err.message);
    }
  }
  ui.alert('Done!\n' + ok + ' sent\n' + fail + ' failed');
}