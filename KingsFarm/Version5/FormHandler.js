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
    let existing  = findRiderByPhone(phone);
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
      amount, keNo, upiLink, qrCode,
      bookingDate, isFirstTime,
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
    const payDate = vals[CONFIG.PAYMENT_COLS.PAY_DATE];
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
    const data    = sheet.getDataRange().getValues();
    const normPh  = normalisePhone(phone);
    const normDate = ymd(payDate);
    const normCurTs = ymd(currentTs);
    for (let i = 1; i < data.length; i++) {
      if (ymd(data[i][CONFIG.PAYMENT_COLS.TIMESTAMP]) === normCurTs) continue;
      const rowPh  = normalisePhone(data[i][CONFIG.PAYMENT_COLS.PHONE]);
      const rowAmt = Number(data[i][CONFIG.PAYMENT_COLS.AMOUNT]);
      const rowDt  = ymd(data[i][CONFIG.PAYMENT_COLS.PAY_DATE]);
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

  const vals     = paymentSheet.getRange(rowIndex, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
  const keNo     = String(vals[CONFIG.PAYMENT_COLS.KE_NO]   || '').trim();
  const phone    = String(vals[CONFIG.PAYMENT_COLS.PHONE]   || '').trim();
  const amount   = Number(vals[CONFIG.PAYMENT_COLS.AMOUNT]);
  const payDate  = vals[CONFIG.PAYMENT_COLS.PAY_DATE];
  const txnRef   = String(vals[CONFIG.PAYMENT_COLS.TXN_REF] || '').trim();
  const pan      = String(vals[CONFIG.PAYMENT_COLS.PAN]     || '').trim();
  const verified = String(vals[CONFIG.PAYMENT_COLS.VERIFIED]|| '').toLowerCase();

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

  const receiptNo = _generateReceiptNo(rKeNo);

  // Build 80G PDF using DocumentApp (no DriveApp.createFile needed)
  const receiptPDF = generate80GReceipt(rName, pan, amount, txnRef, receiptNo);

  // Store in Drive — non-fatal, email still sends if Drive scope missing
  let driveInfo = null;
  try { driveInfo = storeReceiptInDrive(receiptPDF, rName, receiptNo); }
  catch (driveErr) { Logger.log('storeReceiptInDrive skipped: ' + driveErr); }

  // Append to Payments Ledger — non-fatal
  try { _appendToLedger(ss, { keNo: rKeNo, name: rName, phone: rPhone, amount, payDate, txnRef, receiptNo }); }
  catch (ledgerErr) { Logger.log('_appendToLedger error: ' + ledgerErr); }

  // Send receipt email — always runs
  const ccEmails = getCCRecipients('receipt');
  GmailApp.sendEmail(
    rEmail,
    'Payment Receipt — ' + rName + ' (' + rKeNo + ')',
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
  const ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmm');
  const last4 = String(keNo).replace(/\D/g,'').slice(-4) || '0000';
  return keNo + '/' + last4 + '/' + ts.slice(-4);
}

// ────────────────────────────────────────────────────────────
//  APPEND TO PAYMENTS LEDGER  (cell-by-cell — avoids array
//  length mismatch if sheet has extra/missing columns)
// ────────────────────────────────────────────────────────────

function _appendToLedger(ss, d) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
    if (!sheet) {
      Logger.log('_appendToLedger: Payments Ledger sheet not found — skipping');
      return;
    }

    // appendRow with a plain array that has exactly 8 items
    const row = [
      d.keNo      || '',
      d.name      || '',
      d.phone     || '',
      d.amount    || 0,
      d.payDate   || '',
      d.txnRef    || '',
      d.receiptNo || '',
      new Date()
    ];

    sheet.appendRow(row);

    // Format date columns in the new row
    const lr = sheet.getLastRow();
    if (d.payDate) {
      sheet.getRange(lr, CONFIG.LEDGER_COLS.PAY_DATE + 1).setNumberFormat('dd-MMM-yyyy');
    }
    sheet.getRange(lr, CONFIG.LEDGER_COLS.SENT_AT + 1).setNumberFormat('dd-MMM-yyyy HH:mm');

    // Colour the row green so it's easy to spot
    sheet.getRange(lr, 1, 1, 8).setBackground('#f0faf5');

    Logger.log('_appendToLedger: row ' + lr + ' written for ' + d.keNo + ' — ₹' + d.amount);
  } catch (e) {
    Logger.log('_appendToLedger ERROR: ' + e + '\n' + e.stack);
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
      errors.push('Row ' + (startRow+i) + ': ' + err.message);
    }
  }
  ui.alert('Done!\n\u2705 ' + ok + ' sent\n\u274c ' + fail + ' failed'
    + (errors.length ? '\n\n' + errors.slice(0,5).join('\n') : ''));
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
      const keNo         = String(vals[CONFIG.BOOKING_COLS.KE_NO]        || '').trim();
      const bookingDate  = vals[CONFIG.BOOKING_COLS.TIMESTAMP];
      if (!email || !keNo) throw new Error('Missing email or KE No');
      const amount  = CONFIG.ADVANCE_BOOKING_AMOUNT;
      const upiLink = createUPILink(amount, keNo);
      const qrCode  = createQRCode(upiLink);
      sendWelcomeEmail({ name, email, phone, services, participants,
        amount, keNo, upiLink, qrCode, bookingDate, isFirstTime: false,
        sheet: bookingSheet, row: rowIdx });
      ok++;
      Utilities.sleep(800);
    } catch (err) {
      fail++;
      Logger.log('resendWelcomeEmail row ' + (start+i) + ': ' + err.message);
    }
  }
  ui.alert('Done!\n\u2705 ' + ok + ' sent\n\u274c ' + fail + ' failed');
}