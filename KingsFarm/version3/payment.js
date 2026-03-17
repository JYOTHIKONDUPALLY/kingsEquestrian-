// ============================================
// KINGS EQUESTRIAN - PAYMENT LOGIC
// File: 3_Payment.js
// Depends on: 1_Config.js, 2_Booking.js (for createBookingCalendarEvent)
// ============================================

// --------------- PAYMENT FORM SUBMIT HANDLER ---------------
//
// CHANGE: The "Registration No" field in the payment form now accepts
// a phone number (or optionally a KE-reference).
// Logic:
//   1. Read the raw value from PHONE_OR_REG column.
//   2. Call findBookingByPhoneOrRef() which tries KE-ref match first,
//      then falls back to phone-number match.
//   3. If a booking is found, use its stored KE-reference for the receipt.
//   4. If no booking found, log & flag the row — do NOT silently fail.
//
// ---------------------------------------------------------------

function onPaymentFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        if (sheet.getName() !== CONFIG.SHEETS.PAYMENT_FORM) {
            Logger.log('onPaymentFormSubmit: Skipping — wrong sheet: ' + sheet.getName());
            return;
        }
        const row = e.range.getRow();
        Logger.log(`Payment form submitted at row ${row}`);

        const submittedValue = String(sheet.getRange(row, CONFIG.PAYMENT_COLS.PHONE_OR_REG + 1).getValue() || '').trim();
        const amount         = Number(sheet.getRange(row, CONFIG.PAYMENT_COLS.AMOUNT_PAID + 1).getValue());
        const paymentDate    = sheet.getRange(row, CONFIG.PAYMENT_COLS.PAYMENT_DATE + 1).getValue();
        const timestamp      = sheet.getRange(row, CONFIG.PAYMENT_COLS.TIMESTAMP + 1).getValue();

        if (!submittedValue) {
            Logger.log('No phone/registration value found in payment form submission');
            flagPaymentRow(sheet, row, '⚠️ Missing phone/reg', '#f8d7da');
            return;
        }

        Logger.log(`Submitted identifier: "${submittedValue}"`);

        // Resolve booking
        const bookingMatch = findBookingByPhoneOrRef(submittedValue);
        if (!bookingMatch) {
            Logger.log(`No booking found for "${submittedValue}"`);
            flagPaymentRow(sheet, row, `❌ No booking: ${submittedValue}`, '#f8d7da');
            return;
        }

        // Use the authoritative KE-reference from the booking sheet
        const referenceNumber = String(bookingMatch.row[CONFIG.BOOKING_COLS.REFERENCE] || '').trim();
        Logger.log(`Resolved booking reference: ${referenceNumber} for submitted: ${submittedValue}`);

        // Write resolved reference back into the payment sheet for traceability
        sheet.getRange(row, CONFIG.PAYMENT_COLS.PHONE_OR_REG + 1).setValue(referenceNumber);

        // Check duplicate
        const duplicateInfo = findDuplicateReceipt(referenceNumber, amount, paymentDate, timestamp);

        if (duplicateInfo.isDuplicate) {
            Logger.log(`Duplicate receipt detected for ${referenceNumber}. Resending existing receipt.`);
            sheet.getRange(row, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
                .setValue('Duplicate - Resent')
                .setBackground('#fff3cd')
                .setFontColor('#856404');
            resendExistingReceipt(row, duplicateInfo.existingRow, bookingMatch);
            return;
        }

        // Auto-verify & send receipt
        sheet.getRange(row, CONFIG.PAYMENT_COLS.TRANSACTION_VERIFIED + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        Logger.log('Transaction auto-verified, proceeding to send receipt');
        Utilities.sleep(500);
        sendReceiptForRow(row);

    } catch (error) {
        Logger.log('Error in onPaymentFormSubmit: ' + error);
        Logger.log('Stack: ' + error.stack);
    }
}

/** Marks a payment row with a status message and background colour. */
function flagPaymentRow(sheet, row, message, bgColor) {
    sheet.getRange(row, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
        .setValue(message)
        .setBackground(bgColor)
        .setFontColor('#333333');
}

// --------------- DUPLICATE DETECTION ---------------

function findDuplicateReceipt(referenceNumber, amount, paymentDate, currentTimestamp) {
    try {
        const ss           = SpreadsheetApp.getActiveSpreadsheet();
        const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);

        if (!paymentSheet) return { isDuplicate: false, existingRow: null };

        const data                    = paymentSheet.getDataRange().getValues();
        const normalizedDate          = normalizeDate(paymentDate);
        const normalizedCurrentTimestamp = normalizeDate(currentTimestamp);

        for (let i = 1; i < data.length; i++) {
            const rowRef        = String(data[i][CONFIG.PAYMENT_COLS.PHONE_OR_REG] || '').trim();
            const rowAmount     = Number(data[i][CONFIG.PAYMENT_COLS.AMOUNT_PAID]);
            const rowDate       = data[i][CONFIG.PAYMENT_COLS.PAYMENT_DATE];
            const rowTimestamp  = data[i][CONFIG.PAYMENT_COLS.TIMESTAMP];
            const rowReceiptSent = String(data[i][CONFIG.PAYMENT_COLS.RECEIPT_SENT] || '').trim();

            // Skip the current row itself
            if (normalizeDate(rowTimestamp) === normalizedCurrentTimestamp) continue;

            if (rowRef === String(referenceNumber).trim() &&
                rowAmount === amount &&
                normalizeDate(rowDate) === normalizedDate &&
                rowReceiptSent.toLowerCase() === 'yes') {

                Logger.log(`Found existing receipt at row ${i + 1}`);
                return { isDuplicate: true, existingRow: i + 1 };
            }
        }

        return { isDuplicate: false, existingRow: null };

    } catch (error) {
        Logger.log('Error checking for duplicate: ' + error);
        return { isDuplicate: false, existingRow: null };
    }
}

// --------------- RESEND EXISTING RECEIPT ---------------

/**
 * @param {number} currentRow        - Row index in payment sheet being processed.
 * @param {number} existingRow       - Row index of the already-receipted duplicate.
 * @param {object|null} bookingMatch - Pre-resolved booking match (optional — will re-resolve if null).
 */
function resendExistingReceipt(currentRow, existingRow, bookingMatch) {
    try {
        const ss           = SpreadsheetApp.getActiveSpreadsheet();
        const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
        const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);

        if (!paymentSheet || !bookingSheet) {
            Logger.log('Required sheets not found');
            return false;
        }

        const existingData          = paymentSheet.getRange(existingRow, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
        const existingReceiptNumber = existingData[CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_NO];
        const existingDriveLink     = existingData[CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_DRIVER_LINK];

        const currentData     = paymentSheet.getRange(currentRow, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
        const referenceNumber = currentData[CONFIG.PAYMENT_COLS.PHONE_OR_REG];
        const amount          = Number(currentData[CONFIG.PAYMENT_COLS.AMOUNT_PAID]);
        const transactionId   = currentData[CONFIG.PAYMENT_COLS.TRANSACTION_REFERENCE_NUMBER] || '';
        const pan             = currentData[CONFIG.PAYMENT_COLS.PAN_AADHAAR] || '';

        // Use provided match or resolve again
        if (!bookingMatch) {
            bookingMatch = findBookingByPhoneOrRef(referenceNumber);
        }
        if (!bookingMatch) throw new Error(`Booking not found for reference ${referenceNumber}`);

        const riderName        = bookingMatch.row[CONFIG.BOOKING_COLS.NAME];
        const email            = bookingMatch.row[CONFIG.BOOKING_COLS.EMAIL_ID];
        const preferredDate    = bookingMatch.row[CONFIG.BOOKING_COLS.PREFERRED_SERVICE_DATE];
        const preferredTimeSlots = bookingMatch.row[CONFIG.BOOKING_COLS.PREFERRED_TIME_SLOT];

        if (!email) throw new Error('Email not found in booking');

        const receiptPDF = generate80GReceipt(riderName, pan, amount, transactionId, existingReceiptNumber);

        const subject  = `Payment Receipt - ${riderName} - Ref: ${referenceNumber}`;
        const htmlBody = buildReceiptEmailHTML(riderName, referenceNumber, existingReceiptNumber, amount, transactionId, preferredDate, preferredTimeSlots);

        const ccEmails = getCCRecipients('Receipt Mail');

        MailApp.sendEmail({
            to: email,
            cc: ccEmails.join(','),
            subject: subject,
            htmlBody: htmlBody,
            attachments: [receiptPDF],
            name: 'Kings Equestrian Foundation'
        });

        paymentSheet.getRange(currentRow, CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_NO + 1).setValue(existingReceiptNumber);
        paymentSheet.getRange(currentRow, CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_DRIVER_LINK + 1).setValue(existingDriveLink);
        paymentSheet.getRange(currentRow, CONFIG.PAYMENT_COLS.RECEIPT_SENT_TIMESTAMP + 1)
            .setValue(new Date())
            .setNumberFormat('dd-MMM-yyyy HH:mm:ss');

        Logger.log(`Existing receipt ${existingReceiptNumber} resent to: ${email}`);
        return true;

    } catch (error) {
        Logger.log(`Error resending existing receipt: ${error.message}`);
        return false;
    }
}

// --------------- SEND RECEIPT FOR SPECIFIC ROW ---------------

function sendReceiptForRow(rowIndex) {
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);

    if (!paymentSheet) {
        Logger.log('Payment sheet not found');
        return false;
    }

    let email           = '';
    let referenceNumber = '';

    try {
        const row            = paymentSheet.getRange(rowIndex, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
        const submittedValue = String(row[CONFIG.PAYMENT_COLS.PHONE_OR_REG] || '').trim();

        if (!submittedValue) throw new Error('Phone/registration value missing in payment row');

        // Resolve booking
        const bookingMatch = findBookingByPhoneOrRef(submittedValue);
        if (!bookingMatch) throw new Error(`No booking found for submitted value: ${submittedValue}`);

        // Use authoritative KE-reference from booking
        referenceNumber = String(bookingMatch.row[CONFIG.BOOKING_COLS.REFERENCE] || '').trim();

        const riderName        = bookingMatch.row[CONFIG.BOOKING_COLS.NAME];
        email                  = bookingMatch.row[CONFIG.BOOKING_COLS.EMAIL_ID];
        const phone            = bookingMatch.row[CONFIG.BOOKING_COLS.PHONE_NUMBER];
        const services         = bookingMatch.row[CONFIG.BOOKING_COLS.OUR_SERVICES];
        const participants     = bookingMatch.row[CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS] || 1;
        const preferredDate    = bookingMatch.row[CONFIG.BOOKING_COLS.PREFERRED_SERVICE_DATE];
        const preferredTimeSlots = bookingMatch.row[CONFIG.BOOKING_COLS.PREFERRED_TIME_SLOT];

        if (!email) throw new Error('Email not found in booking');

        const amount = Number(row[CONFIG.PAYMENT_COLS.AMOUNT_PAID]);
        if (!amount || Number.isNaN(amount)) throw new Error('Valid amount is required');

        const transactionId      = row[CONFIG.PAYMENT_COLS.TRANSACTION_REFERENCE_NUMBER] || '';
        const pan                = row[CONFIG.PAYMENT_COLS.PAN_AADHAAR] || '';
        const transactionVerified = row[CONFIG.PAYMENT_COLS.TRANSACTION_VERIFIED];

        if (String(transactionVerified || '').toLowerCase() !== 'yes') {
            throw new Error('Transaction not verified. Please verify first.');
        }

        // Generate a unique receipt number based on the KE-reference
        const receiptNumber = generateReceiptNumber(referenceNumber);
        const receiptPDF    = generate80GReceipt(riderName, pan, amount, transactionId, receiptNumber);

        // Store in Drive
        const driveInfo = storeReceiptInDrive(receiptPDF, riderName, receiptNumber, referenceNumber);
        if (driveInfo) {
            Logger.log(`Receipt stored in Drive: ${driveInfo.fileUrl}`);
            paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_DRIVER_LINK + 1).setValue(driveInfo.fileUrl);
        }

        // Send email
        const subject  = `Payment Receipt - ${riderName} - Ref: ${referenceNumber}`;
        const htmlBody = buildReceiptEmailHTML(riderName, referenceNumber, receiptNumber, amount, transactionId, preferredDate, preferredTimeSlots);

        const ccEmails = getCCRecipients('Receipt Mail');

        MailApp.sendEmail({
            to: email,
            cc: ccEmails.join(','),
            subject: subject,
            htmlBody: htmlBody,
            attachments: [receiptPDF],
            name: 'Kings Equestrian Foundation'
        });

        // Update payment sheet
        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_SENT_TIMESTAMP + 1)
            .setValue(new Date())
            .setNumberFormat('dd-MMM-yyyy HH:mm:ss');

        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_NO + 1).setValue(receiptNumber);

        // Write resolved reference back (in case it was a phone-number submission)
        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.PHONE_OR_REG + 1).setValue(referenceNumber);

        // Calendar event
        if (preferredDate && preferredTimeSlots) {
            try {
                const calId = createBookingCalendarEvent({
                    name: riderName, email, phone, services,
                    date: preferredDate, timeSlots: preferredTimeSlots,
                    reference: referenceNumber, participants
                });
                if (calId) Logger.log(`Calendar event created: ${calId} for ${referenceNumber}`);
            } catch (calError) {
                Logger.log(`Warning: Calendar event failed for ${referenceNumber}: ${calError.message}`);
            }
        }

        Logger.log(`Receipt sent to: ${email} for ${referenceNumber} — CC: ${ccEmails.join(', ')}`);
        return true;

    } catch (error) {
        Logger.log(`Receipt failed for row ${rowIndex} (Ref: ${referenceNumber || 'N/A'}, Email: ${email || 'N/A'}): ${error.message}`);
        throw error;
    }
}

// --------------- RECEIPT NUMBER GENERATION ---------------

/**
 * Generates a receipt number from the KE-reference.
 * Format: KE250314xxxx/xxxx
 * If no valid KE-reference is available, falls back to a timestamp-based number.
 */
function generateReceiptNumber(referenceNumber) {
    if (referenceNumber && isKEReference(referenceNumber)) {
        const serialMatch = referenceNumber.match(/\d{4}$/);
        const serial      = serialMatch ? serialMatch[0] : '0000';
        return `${referenceNumber}/${serial}`;
    }
    // Fallback: timestamp-based receipt number (for edge cases)
    const now    = new Date();
    const ts     = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMddHHmm');
    const random = Math.floor(Math.random() * 9000) + 1000;
    return `KE-RCP-${ts}-${random}`;
}

// --------------- RECEIPT EMAIL HTML BUILDER ---------------

function buildReceiptEmailHTML(riderName, referenceNumber, receiptNumber, amount, transactionId, preferredDate, preferredTimeSlots) {
    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif; background-color:#f4f4f4; margin:0; padding:0; color:#333; }
    .container { max-width:650px; margin:20px auto; background:#fff; border-radius:12px; overflow:hidden; box-shadow:0 4px 6px rgba(0,0,0,0.1); }
    .header { background:linear-gradient(135deg,#1f4e3d 0%,#4f9c7a 100%); padding:30px; text-align:center; color:#fff; }
    .header img { width:70px; height:70px; border-radius:10%; margin-bottom:10px; }
    .header h1 { margin:0; font-size:26px; }
    .content { padding:35px 30px; }
    .greeting { font-size:18px; font-weight:600; margin-bottom:15px; color:#1f4e3d; }
    .success-banner { text-align:center; margin:25px 0; }
    .info-box { background:#f8f8f8; border-left:4px solid #1f4e3d; padding:18px; margin:20px 0; border-radius:6px; font-size:14px; }
    .footer { background-color:#1f4e3d; color:#fff; padding:25px; text-align:center; font-size:13px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian Logo">
      <h1>Kings Equestrian Foundation</h1>
      <p style="margin:8px 0 0;">Where horses don't just carry you - they change you</p>
    </div>
    <div class="content">
      <div class="greeting">Dear ${riderName},</div>
      <div class="success-banner">
        <img src="https://i.pinimg.com/736x/69/3c/20/693c200ad675967032f941cf76953b3e.jpg" alt="Payment Successful" width="200" height="150">
        <div style="font-size:18px;font-weight:600;color:#1f7a3f;margin-top:10px;">✅ Payment Confirmed - Booking Complete!</div>
        <p>Thank you for your payment. Your booking is confirmed.</p>
      </div>
      <div class="info-box">
        <strong>Your Payment Receipt (80G)</strong> is attached to this email for tax deduction purposes.
      </div>
      <p>
        <strong>Payment Details:</strong><br>
        Booking Reference: ${referenceNumber}<br>
        Receipt No: ${receiptNumber}<br>
        Amount Paid: ₹${amount.toLocaleString('en-IN')}<br>
        ${transactionId ? `Transaction ID: ${transactionId}<br>` : ''}
        ${preferredDate ? `Scheduled Date: ${formatDate(preferredDate)}<br>` : ''}
        ${preferredTimeSlots ? `Time Slot: ${preferredTimeSlots}<br>` : ''}
      </p>
      <p style="margin-top:20px;font-size:14px;">
        We look forward to welcoming you at Kings Equestrian. Please arrive 15 minutes before your scheduled time.
      </p>
      <p style="margin-top:15px;font-size:13px;color:#666;">
        <strong>What to bring:</strong><br>
        • Comfortable clothing<br>
        • Closed-toe shoes<br>
        • Your booking reference: ${referenceNumber}
      </p>
    </div>
    <div class="footer">
      <p><strong>Kings Equestrian Foundation</strong></p>
      <p>Karnataka, India</p>
      <p>+91-9980895533 | info@kingsequestrian.com</p>
    </div>
  </div>
</body>
</html>`;
}

// --------------- SEND PAYMENT RECEIPT MENU FUNCTION ---------------

function SendPaymentReceipt() {
    const ui           = SpreadsheetApp.getUi();
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);

    if (!paymentSheet) {
        ui.alert('Payment Form Response sheet not found');
        return;
    }

    const selection = paymentSheet.getActiveRange();
    if (!selection || selection.getRow() === 1) {
        ui.alert('Please select valid rows to send receipts (not header row)');
        return;
    }

    const startRow = selection.getRow();
    const numRows  = selection.getNumRows();

    const response = ui.alert('Send Payment Receipts', `Send receipts for ${numRows} row(s)?`, ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    let successCount = 0;
    let failCount    = 0;
    const errors     = [];

    for (let i = 0; i < numRows; i++) {
        const rowIndex = startRow + i;
        try {
            sendReceiptForRow(rowIndex);
            successCount++;
            Utilities.sleep(1000);
        } catch (error) {
            failCount++;
            errors.push(`Row ${rowIndex}: ${error.message}`);
        }
    }

    let message = `Complete!\n✅ Sent: ${successCount}\n❌ Failed: ${failCount}`;
    if (errors.length > 0) {
        message += '\n\nErrors:\n' + errors.slice(0, 5).join('\n');
        if (errors.length > 5) message += `\n... and ${errors.length - 5} more`;
    }
    ui.alert(message);
}