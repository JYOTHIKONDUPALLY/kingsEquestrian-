// ============================================================
// INDUS EQUESTRIAN — SCHOOL SYSTEM
// File: 2_FormHandlers.gs
// Registration form submit + Payment form submit handlers
// ============================================================

// ────────────────────────────────────────────────────────────
//  REGISTRATION FORM SUBMIT
//  Triggered when a student submits the school registration form.
//  • Creates/finds rider record (participants always = 1)
//  • Writes KE No back to the sheet
//  • Sends welcome email with appropriate consent PDF
// ────────────────────────────────────────────────────────────

function onRegistrationFormSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    if (!isRegistrationSheetName_(sheet.getName())) return;

    const row  = e.range.getRow();
    const vals = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    // ── pull all fields ──
    const studentName      = String(vals[CONFIG.REG_COLS.STUDENT]          || '').trim();
    const parentName       = String(vals[CONFIG.REG_COLS.PARENT]           || '').trim();
    const email            = String(vals[CONFIG.REG_COLS.EMAIL]            || '').trim();
    const phone            = String(vals[CONFIG.REG_COLS.PHONE]            || '').trim();
    const grade            = String(vals[CONFIG.REG_COLS.GRADE]            || '').trim();
    const section          = String(vals[CONFIG.REG_COLS.SECTION]          || '').trim();
    const gradeDisplay     = section ? (grade + ' · ' + section) : grade;
    const program          = String(vals[CONFIG.REG_COLS.PROGRAM]          || 'school').trim().toLowerCase();
    const horseLease       = vals[CONFIG.REG_COLS.HORSE_LEASE];
    const dob              = vals[CONFIG.REG_COLS.DOB];
    const address          = String(vals[CONFIG.REG_COLS.ADDRESS]          || '').trim();
    const motherName       = String(vals[CONFIG.REG_COLS.MOTHER_NAME]      || '').trim();
    const fatherName       = String(vals[CONFIG.REG_COLS.FATHER_NAME]      || '').trim();
    const motherContact    = String(vals[CONFIG.REG_COLS.MOTHER_CONTACT]   || '').trim();
    const motherWhatsApp   = String(vals[CONFIG.REG_COLS.MOTHER_WHATSAPP]  || '').trim();
    const fatherContact    = String(vals[CONFIG.REG_COLS.FATHER_CONTACT]   || '').trim();
    const fatherWhatsApp   = String(vals[CONFIG.REG_COLS.FATHER_WHATSAPP]  || '').trim();
    const emergencyContact = String(vals[CONFIG.REG_COLS.EMERGENCY_CONTACT]|| '').trim();
    const relationship     = String(vals[CONFIG.REG_COLS.RELATIONSHIP]     || 'ward').trim();
    const consentDate      = vals[CONFIG.REG_COLS.CONSENT_DATE] || vals[CONFIG.REG_COLS.TIMESTAMP];
    const timestamp        = vals[CONFIG.REG_COLS.TIMESTAMP];

    // Derive service/program label for display + consent form checkboxes
    // The form's PROGRAM field may contain the full service string
    // e.g. "2 classes per week" or "3 classes per week" or "summer"
    const serviceProgram   = program;   // re-used for checkbox selection in consent PDF

    if (!email || !phone || !studentName) {
      Logger.log('onRegistrationFormSubmit: missing required fields at row ' + row);
      return;
    }

    // ── find or create rider ──
    let existing    = findRiderByPhoneAndName(phone, studentName);
    let keNo;
    let isFirstTime = false;

    if (existing) {
      keNo = String(existing.row[CONFIG.RIDER_COLS.KE_NO] || '').trim();
      // Update service if changed
      _updateRiderService(existing.rowIndex, serviceProgram);
    } else {
      keNo        = generateKENo();
      isFirstTime = true;
      _createRiderRecord({
        keNo,
        name    : studentName,
        email,
        phone,
        services: serviceProgram,
        // participants always 1 for school
      });
    }

    // ── write KE No and program track back to sheet ──
    _ensureExtraRegHeaders(sheet);
    sheet.getRange(row, CONFIG.REG_COLS.REG_REF + 1).setValue(keNo);
    sheet.getRange(row, CONFIG.REG_COLS.PROGRAM_TRACK + 1).setValue(program);
    sheet.getRange(row, CONFIG.REG_COLS.KE_NO + 1).setValue(keNo);

    // ── training sheets: STUDENTS + BOOKINGS (Booking_ID for trainer app / curriculum) ──
    try {
      if (typeof ensureTrainingRecordsForSchoolRegistration === 'function') {
        ensureTrainingRecordsForSchoolRegistration({
          keNo        : keNo,
          studentName : studentName,
          parentName  : parentName,
          email       : email,
          phone       : phone,
          program     : serviceProgram || program,
          grade       : grade,
          section     : section
        });
      }
    } catch (syncErr) {
      Logger.log('ensureTrainingRecordsForSchoolRegistration: ' + syncErr);
    }

    // Grade/section and registration time drive the late-join makeup backlog.
    // Clear rider/session caches so the new rider appears immediately.
    try {
      if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();
    } catch (cacheErr) {
      Logger.log('registration cache invalidation: ' + cacheErr);
    }

    // ── build payment form URL (pre-filled) ──
    const payFormUrl = buildPaymentFormUrl({
      regRef : keNo,
      student: studentName,
      parent : parentName,
      email,
      phone,
      grade: gradeDisplay
    });

    // ── send welcome email ──
    sendIndusSchoolWelcomeEmail({
      studentName, parentName, email, phone, grade: gradeDisplay,
      program, serviceProgram, horseLease,
      dob, address,
      motherName, fatherName,
      motherContact, motherWhatsApp,
      fatherContact, fatherWhatsApp,
      emergencyContact, relationship,
      consentDate, timestamp,
      keNo, payFormUrl, isFirstTime,
      sheet, row
    });

    Logger.log('Registration processed: ' + keNo + ' (' + studentName + ') program=' + program + ' isFirstTime=' + isFirstTime);
  } catch (err) {
    Logger.log('onRegistrationFormSubmit ERROR: ' + err + '\n' + err.stack);
  }
}

// ────────────────────────────────────────────────────────────
//  ENSURE EXTRA HEADERS EXIST
// ────────────────────────────────────────────────────────────

function _ensureExtraRegHeaders(sheet) {
  try {
    const headers      = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const existing     = headers.map(function(h) { return String(h || '').trim(); });
    const toAdd        = CONFIG.EXTRA_REG_HEADERS.filter(function(h) { return existing.indexOf(h) === -1; });
    if (toAdd.length) {
      const nextCol = sheet.getLastColumn() + 1;
      toAdd.forEach(function(h, i) {
        sheet.getRange(1, nextCol + i).setValue(h)
          .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
      });
    }
  } catch (e) {
    Logger.log('_ensureExtraRegHeaders error: ' + e);
  }
}

// ────────────────────────────────────────────────────────────
//  CREATE RIDER RECORD  (participants fixed at 1 for school)
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
  newRow[CONFIG.RIDER_COLS.SERVICES]     = d.services || '';
  newRow[CONFIG.RIDER_COLS.PARTICIPANTS] = 1;             // always 1 for school
  newRow[CONFIG.RIDER_COLS.REGISTERED]   = new Date();

  sheet.appendRow(newRow);
  sheet.getRange(sheet.getLastRow(), CONFIG.RIDER_COLS.REGISTERED + 1)
    .setNumberFormat('dd-MMM-yyyy');
  Logger.log('Rider record created: ' + d.keNo + ' (school — 1 participant)');
}

function _updateRiderService(rowIndex, services) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
    if (!sheet || !services) return;
    sheet.getRange(rowIndex, CONFIG.RIDER_COLS.SERVICES + 1).setValue(services);
  } catch (e) {
    Logger.log('_updateRiderService error: ' + e);
  }
}

// ────────────────────────────────────────────────────────────
//  PAYMENT SHEET HELPERS (column-safe for removed / optional fields)
// ────────────────────────────────────────────────────────────

function _normalisePaymentHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function _paymentHeaderAliases_(key) {
  const aliases = {
    TIMESTAMP: ['Timestamp'],
    REG_REF: ['Registration No', 'Registration Number', 'KE No'],
    PHONE: ['Phone number', 'Phone'],
    AMOUNT: ['Amount Paid (₹)', 'Amount Paid', 'Amount'],
    SCREENSHOT: ['ScreenShot', 'Screenshot'],
    PAY_DATE: ['Payment Date'],
    TXN_REF: ['Transcation Reference Number', 'Transaction Reference Number', 'Transaction ID'],
    PAN: ['Pan / AAdhar Number', 'PAN / Aadhaar Number'],
    MODE: ['Mode of Payment', 'Payment Mode'],
    PAYMENT_FOR: ['Payment For', 'Payment for', 'Purpose of Payment'],
    RECEIPT_SENT: ['Receipt Sent'],
    RECEIPT_AT: ['Receipt Sent At'],
    RECEIPT_NO: ['Receipt No'],
    RECEIPT_LINK: ['Receipt Link']
  };
  return aliases[key] || [];
}

function _paymentColIndex_(sheetOrHeaders, key) {
  let headers = sheetOrHeaders;
  if (sheetOrHeaders && typeof sheetOrHeaders.getRange === 'function') {
    headers = sheetOrHeaders.getRange(1, 1, 1, sheetOrHeaders.getLastColumn()).getValues()[0];
  }
  headers = headers || [];
  const wanted = _paymentHeaderAliases_(key).map(_normalisePaymentHeader_);
  for (let i = 0; i < headers.length; i++) {
    if (wanted.indexOf(_normalisePaymentHeader_(headers[i])) >= 0) return i;
  }
  const fallback = CONFIG.PAYMENT_COLS[key];
  return fallback == null ? -1 : fallback;
}

function _paymentVal_(vals, key, headers) {
  const idx = headers ? _paymentColIndex_(headers, key) : CONFIG.PAYMENT_COLS[key];
  if (idx == null || idx < 0) return '';
  const v = vals[idx];
  return v === undefined || v === null ? '' : v;
}

function _normalisePaymentType_(value) {
  const text = String(value || '').trim().toLowerCase();
  return (text.indexOf('shop') >= 0 || text.indexOf('kit') >= 0 || text.indexOf('equipment') >= 0)
    ? 'Shopping Kit / Equipment' : 'Riding Classes';
}

/**
 * Load student / parent / email / grade from registration sheet.
 * Match order: Registration No (KE) if present, else normalized phone.
 */
function lookupRegistrationForPayment_(regRef, paymentPhone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRegistrationSheet_(ss);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const keNorm = String(regRef || '').trim().toUpperCase();
  const phNorm = normalisePhone(paymentPhone);

  function packRow(i) {
    const row = data[i];
    return {
      student: String(row[CONFIG.REG_COLS.STUDENT] || '').trim(),
      parent: String(row[CONFIG.REG_COLS.PARENT] || '').trim(),
      email: String(row[CONFIG.REG_COLS.EMAIL] || '').trim(),
      grade: String(row[CONFIG.REG_COLS.GRADE] || '').trim(),
      section: String(row[CONFIG.REG_COLS.SECTION] || '').trim(),
      keNo: String(row[CONFIG.REG_COLS.REG_REF] || row[CONFIG.REG_COLS.KE_NO] || '').trim()
    };
  }

  if (keNorm) {
    for (let i = 1; i < data.length; i++) {
      const rowKe = String(data[i][CONFIG.REG_COLS.REG_REF] || data[i][CONFIG.REG_COLS.KE_NO] || '')
        .trim().toUpperCase();
      if (rowKe && rowKe === keNorm) return packRow(i);
    }
  }
  if (phNorm && phNorm.length >= 10) {
    for (let i = 1; i < data.length; i++) {
      if (normalisePhone(data[i][CONFIG.REG_COLS.PHONE]) === phNorm) return packRow(i);
    }
  }
  return null;
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
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const receiptSentCol = _paymentColIndex_(headers, 'RECEIPT_SENT');

    // Auto-verify
    sheet.getRange(row, receiptSentCol + 1)
      .setValue('Verifying')
      .setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');

    Utilities.sleep(300);

    const vals    = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const regRef  = String(_paymentVal_(vals, 'REG_REF', headers) || '').trim();
    const phone   = String(_paymentVal_(vals, 'PHONE', headers) || '').trim();
    const amount  = Number(_paymentVal_(vals, 'AMOUNT', headers));
    const payDate = _paymentVal_(vals, 'PAY_DATE', headers);
    const ts      = _paymentVal_(vals, 'TIMESTAMP', headers);
    const paymentType = _normalisePaymentType_(_paymentVal_(vals, 'PAYMENT_FOR', headers));

    if (_isDuplicatePayment(phone, amount, payDate, ts, paymentType)) {
      Logger.log('Duplicate payment detected for phone ' + phone + '. Skipping.');
      sheet.getRange(row, receiptSentCol + 1)
        .setValue('Duplicate - Skipped')
        .setBackground('#fff3cd').setFontColor('#856404');
      return;
    }

    sendReceiptForRow(row);
  } catch (err) {
    Logger.log('onPaymentFormSubmit ERROR: ' + err);
  }
}

function _isDuplicatePayment(phone, amount, payDate, currentTs, paymentType) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheet  = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    if (!sheet) return false;
    const data      = sheet.getDataRange().getValues();
    const headers   = data[0] || [];
    const normPh    = normalisePhone(phone);
    const normDate  = ymd(payDate);
    const normCurTs = ymd(currentTs);
    for (let i = 1; i < data.length; i++) {
      const rowVals = data[i];
      if (ymd(_paymentVal_(rowVals, 'TIMESTAMP', headers)) === normCurTs) continue;
      const rowPh  = normalisePhone(_paymentVal_(rowVals, 'PHONE', headers));
      const rowAmt = Number(_paymentVal_(rowVals, 'AMOUNT', headers));
      const rowDt  = ymd(_paymentVal_(rowVals, 'PAY_DATE', headers));
      const sent   = String(_paymentVal_(rowVals, 'RECEIPT_SENT', headers) || '').toLowerCase();
      const rowType = _normalisePaymentType_(_paymentVal_(rowVals, 'PAYMENT_FOR', headers));
      if (rowPh === normPh && rowAmt === amount && rowDt === normDate
          && rowType === paymentType && sent === 'yes') return true;
    }
    return false;
  } catch (e) {
    Logger.log('_isDuplicatePayment error: ' + e);
    return false;
  }
}

// ────────────────────────────────────────────────────────────
//  SEND RECEIPT FOR A PAYMENT ROW
// ────────────────────────────────────────────────────────────

function sendReceiptForRow(rowIndex) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
  if (!paymentSheet) throw new Error('Payment Form Response sheet not found');

  const vals      = paymentSheet.getRange(rowIndex, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
  const headers   = paymentSheet.getRange(1, 1, 1, paymentSheet.getLastColumn()).getValues()[0];

  const regRef    = String(_paymentVal_(vals, 'REG_REF', headers) || '').trim();
  const phone     = String(_paymentVal_(vals, 'PHONE', headers) || '').trim();
  const amount    = Number(_paymentVal_(vals, 'AMOUNT', headers));
  const txnRef    = String(_paymentVal_(vals, 'TXN_REF', headers) || '').trim();
  const pan       = String(_paymentVal_(vals, 'PAN', headers) || '').trim();
  const timestamp = _paymentVal_(vals, 'TIMESTAMP', headers);
  const paymentType = _normalisePaymentType_(_paymentVal_(vals, 'PAYMENT_FOR', headers));

  let payDate = _paymentVal_(vals, 'PAY_DATE', headers);
  if (!payDate || fmtDate(payDate) === '') {
    payDate = timestamp || new Date();
    Logger.log('sendReceiptForRow: PAY_DATE empty, using TIMESTAMP');
  }

  const reg = lookupRegistrationForPayment_(regRef, phone);
  let keNo = (regRef || (reg && reg.keNo) || '').trim();

  // Receipt / 80G donor name: parent from registration first, then student
  let rName  = reg ? (reg.parent || reg.student) : '';
  let rEmail = reg ? reg.email : '';

  const rider = findRiderByKENo(keNo || regRef) || findRiderByPhone(phone);
  if (rider) {
    if (!rEmail) rEmail = String(rider.row[CONFIG.RIDER_COLS.EMAIL] || '').trim();
    if (!rName) rName = String(rider.row[CONFIG.RIDER_COLS.NAME] || '').trim();
    if (!keNo) keNo = String(rider.row[CONFIG.RIDER_COLS.KE_NO] || '').trim();
  }

  if (!rEmail) {
    throw new Error('No email found for Registration No ' + (regRef || '—') + ' / phone ' + (phone || '—'));
  }
  if (!keNo) {
    throw new Error('No KE No — enter Registration No on the payment row or use a phone that matches registration.');
  }

  const receiptNo  = _generateReceiptNo(keNo);
  const receiptPDF = generate80GReceipt(rName, pan, amount, txnRef, receiptNo, payDate);

  // Store in Drive
  let driveInfo = null;
  try { driveInfo = storeReceiptInDrive(receiptPDF, rName, receiptNo); } catch (driveErr) { Logger.log('Drive store skipped: ' + driveErr); }

  // Append to Payments Ledger
  try {
    _appendToLedger(ss, {
      keNo, name: rName, parent: reg ? reg.parent : rName,
      student: reg ? reg.student : (rider ? String(rider.row[CONFIG.RIDER_COLS.NAME] || '') : ''),
      grade: reg ? reg.grade : '', section: reg ? reg.section : '',
      phone, amount, payDate, txnRef, receiptNo, paymentType
    });
    if (paymentType === 'Shopping Kit / Equipment' && typeof syncShoppingPaymentsForRider_ === 'function') {
      syncShoppingPaymentsForRider_(keNo);
    }
  } catch (ledgerErr) { Logger.log('_appendToLedger error: ' + ledgerErr); }

  // Send receipt email
  const ccEmails = getCCRecipients('receipt');
  const receiptSubject = paymentType + ' Payment Receipt - ' + rName + ' (' + keNo + ')';
  const receiptSend = sendMailKE_(rEmail, receiptSubject,
    buildReceiptEmailHTML({ name: rName, keNo, amount, txnRef, payDate, receiptNo }),
    {
      attachments : [receiptPDF],
      cc          : ccEmails.join(',')
    }
  );
  try { logEmail('Receipt', rEmail, ccEmails.join(','), receiptSubject, keNo, 'Sent', 'via ' + receiptSend.provider); } catch (logErr) {}

  // Mark row done
  paymentSheet.getRange(rowIndex, _paymentColIndex_(headers, 'RECEIPT_SENT') + 1)
    .setValue('Yes').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
  paymentSheet.getRange(rowIndex, _paymentColIndex_(headers, 'RECEIPT_AT') + 1)
    .setValue(new Date()).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
  paymentSheet.getRange(rowIndex, _paymentColIndex_(headers, 'RECEIPT_NO') + 1).setValue(receiptNo);
  if (driveInfo) {
    paymentSheet.getRange(rowIndex, _paymentColIndex_(headers, 'RECEIPT_LINK') + 1).setValue(driveInfo.fileUrl);
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
// ────────────────────────────────────────────────────────────

function _ensurePaymentLedgerSchema_(sheet) {
  const headers = [
    'KE No', 'Name', 'Phone', 'Amount', 'Payment Date', 'Txn Ref', 'Receipt No', 'Sent At',
    'Student Name', 'Grade', 'Section', 'Payment Type', 'Parent Name'
  ];
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (header, i) {
    if (!String(current[i] || '').trim()) sheet.getRange(1, i + 1).setValue(header);
  });
}

function _appendToLedger(ss, d) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
    if (!sheet) { Logger.log('_appendToLedger: Payments Ledger sheet not found'); return; }
    _ensurePaymentLedgerSchema_(sheet);

    const row = new Array(13).fill('');
    row[CONFIG.LEDGER_COLS.KE_NO]      = d.keNo      || '';
    row[CONFIG.LEDGER_COLS.NAME]       = d.name      || '';
    row[CONFIG.LEDGER_COLS.PHONE]      = d.phone     || '';
    row[CONFIG.LEDGER_COLS.AMOUNT]     = d.amount    || 0;
    row[CONFIG.LEDGER_COLS.PAY_DATE]   = d.payDate   || '';
    row[CONFIG.LEDGER_COLS.TXN_REF]    = d.txnRef    || '';
    row[CONFIG.LEDGER_COLS.RECEIPT_NO] = d.receiptNo || '';
    row[CONFIG.LEDGER_COLS.SENT_AT]    = new Date();
    row[CONFIG.LEDGER_COLS.STUDENT]    = d.student    || '';
    row[CONFIG.LEDGER_COLS.GRADE]      = d.grade      || '';
    row[CONFIG.LEDGER_COLS.SECTION]    = d.section    || '';
    row[CONFIG.LEDGER_COLS.PAYMENT_TYPE] = d.paymentType || 'Riding Classes';
    row[CONFIG.LEDGER_COLS.PARENT]     = d.parent     || d.name || '';

    sheet.appendRow(row);
    const lr = sheet.getLastRow();

    if (d.payDate && fmtDate(d.payDate) !== '') {
      sheet.getRange(lr, CONFIG.LEDGER_COLS.PAY_DATE + 1).setNumberFormat('dd-MMM-yyyy');
    }
    sheet.getRange(lr, CONFIG.LEDGER_COLS.SENT_AT + 1).setNumberFormat('dd-MMM-yyyy HH:mm');
    sheet.getRange(lr, 1, 1, row.length).setBackground('#f0faf5');

    Logger.log('_appendToLedger: ' + d.keNo + ' — Rs.' + d.amount + ' on ' + fmtDate(d.payDate));
  } catch (e) {
    Logger.log('_appendToLedger ERROR: ' + e);
  }
}

// ────────────────────────────────────────────────────────────
//  MENU ACTIONS
// ────────────────────────────────────────────────────────────

function sendPaymentReceiptMenu() {
  const ui           = SpreadsheetApp.getUi();
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
  if (!paymentSheet) { ui.alert('Payment Responses sheet not found'); return; }

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
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRegistrationSheet_(ss);
  if (!sheet) { ui.alert('Registration Response sheet not found'); return; }

  const sel     = sheet.getActiveRange();
  if (!sel || sel.getRow() === 1) { ui.alert('Select valid rows'); return; }
  const start   = sel.getRow();
  const numRows = sel.getNumRows();
  const resp    = ui.alert('Resend welcome emails for ' + numRows + ' row(s)?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  let ok = 0, fail = 0;
  for (let i = 0; i < numRows; i++) {
    try {
      const rowIdx = start + i;
      const vals   = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];

      const studentName      = String(vals[CONFIG.REG_COLS.STUDENT]           || '').trim();
      const parentName       = String(vals[CONFIG.REG_COLS.PARENT]            || '').trim();
      const email            = String(vals[CONFIG.REG_COLS.EMAIL]             || '').trim();
      const phone            = String(vals[CONFIG.REG_COLS.PHONE]             || '').trim();
      const grade            = String(vals[CONFIG.REG_COLS.GRADE]             || '').trim();
      const section          = String(vals[CONFIG.REG_COLS.SECTION]           || '').trim();
      const gradeDisplay     = section ? (grade + ' · ' + section) : grade;
      const program          = String(vals[CONFIG.REG_COLS.PROGRAM]           || 'school').trim().toLowerCase();
      const horseLease       = vals[CONFIG.REG_COLS.HORSE_LEASE];
      const dob              = vals[CONFIG.REG_COLS.DOB];
      const address          = String(vals[CONFIG.REG_COLS.ADDRESS]           || '').trim();
      const motherName       = String(vals[CONFIG.REG_COLS.MOTHER_NAME]       || '').trim();
      const fatherName       = String(vals[CONFIG.REG_COLS.FATHER_NAME]       || '').trim();
      const motherContact    = String(vals[CONFIG.REG_COLS.MOTHER_CONTACT]    || '').trim();
      const motherWhatsApp   = String(vals[CONFIG.REG_COLS.MOTHER_WHATSAPP]   || '').trim();
      const fatherContact    = String(vals[CONFIG.REG_COLS.FATHER_CONTACT]    || '').trim();
      const fatherWhatsApp   = String(vals[CONFIG.REG_COLS.FATHER_WHATSAPP]   || '').trim();
      const emergencyContact = String(vals[CONFIG.REG_COLS.EMERGENCY_CONTACT] || '').trim();
      const relationship     = String(vals[CONFIG.REG_COLS.RELATIONSHIP]      || 'ward').trim();
      const consentDate      = vals[CONFIG.REG_COLS.CONSENT_DATE] || vals[CONFIG.REG_COLS.TIMESTAMP];
      const keNo             = String(vals[CONFIG.REG_COLS.KE_NO]             || '').trim();

      if (!email || !keNo) throw new Error('Missing email or KE No at row ' + rowIdx);

      const payFormUrl = buildPaymentFormUrl({ regRef: keNo, student: studentName, parent: parentName, email, phone, grade: gradeDisplay });

      sendIndusSchoolWelcomeEmail({
        studentName, parentName, email, phone, grade: gradeDisplay,
        program, serviceProgram: program, horseLease,
        dob, address, motherName, fatherName,
        motherContact, motherWhatsApp, fatherContact, fatherWhatsApp,
        emergencyContact, relationship, consentDate,
        keNo, payFormUrl, isFirstTime: false,
        sheet, row: rowIdx
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