// ============================================
// KINGS EQUESTRIAN - ENHANCED BOOKING SYSTEM
// ============================================

// --------------- CONFIG ---------------

const CONFIG = {
    UPI_ID: "vyapar.176548151976@hdfcbank",
    BUSINESS_NAME: "KingsEquestrian",
    PAYMENT_FORM_LINK: "https://forms.gle/WxskpjCcDQWkA7L57",
    EMAIL_TEMPLATE_DOC_ID: "d/17t23GLXC8g8MkdCDx4BmMswuzrdYM_x0",
    TERMS_CONDITIONS_DOC_ID: "1QbJHA5keyTLvgw-5stTY74i92BQ89TYya-NvtJ4YGx4",
    ADVANCE_BOOKING_AMOUNT: 1000,
    webAppUrl: "https://script.google.com/macros/s/AKfycbxGNi137N_vvd6kFWe0CL2clALwKLp7QKsLgiWUd9fGcvYhTlaeQIy15n2vai_1g-PIig/exec",

    // ── NEW: Regular rider portal URL (same web app, student opens with ?reg=KERxxx) ──
    RIDER_PORTAL_URL: "https://script.google.com/macros/s/AKfycbxGNi137N_vvd6kFWe0CL2clALwKLp7QKsLgiWUd9fGcvYhTlaeQIy15n2vai_1g-PIig/exec",
    // ── NEW: Reschedule form link (create form and replace) ──
    RESCHEDULE_FORM_LINK: "https://forms.gle/REPLACE_WITH_RESCHEDULE_FORM",

    SHEETS: {
        BOOKING_FORM: "Booking Form Response",
        PAYMENT_FORM: "Payment Form Response",
        PRICING: "Pricing",
        MAIL_INFO: "Mail Info",
        // ── NEW sheets for regular riders ──
        REGULAR_BOOKING_FORM: "RegularRidersBooking",  // new Google Form responses land here
        REGULAR_STUDENTS:     "Regular Students",               // master record per enrolled student
    },

    BOOKING_COLS: {
        TIMESTAMP: 0,
        NAME: 2,
        EMAIL_ID: 3,
        PHONE_NUMBER: 4,
        OUR_SERVICES: 5,
        NUMBER_OF_PARTICIPANTS: 6,
        PREFERRED_SERVICE_DATE: 7,
        PREFERRED_TIME_SLOT: 8,
        CONSENT: 9,
        REFERENCE: 10,
        WELCOME_EMAIL_SENT: 11,
        WELCOME_EMAIL_TIMESTAMP: 12,
    },

    PAYMENT_COLS: {
        TIMESTAMP: 0,
        REGISTRATION_NO: 2,    // col C — KE or KER ref (optional, may be blank)
        AMOUNT_PAID: 3,
        SCREENSHOT: 4,
        PAYMENT_DATE: 5,
        TRANSACTION_REFERENCE_NUMBER: 6,
        PAN_AADHAAR: 7,
        PHONE_NUMBER: 8,       // col I — primary lookup key
        TRANSACTION_VERIFIED: 9,
        RECEIPT_SENT: 10,
        RECEIPT_SENT_TIMESTAMP: 11,
        PAYMENT_RECEIPT_NO: 12,
        PAYMENT_RECEIPT_DRIVER_LINK: 13
    },

    // ── NEW: Regular Booking Form columns (your actual Google Form column order) ──
    // Form fields: Timestamp | Name | Email ID | Phone Number | Our Services | No. of Participants | Consent
    // Script writes into cols 8, 9, 10 (H, I, J) after submission
    REG_BOOKING_COLS: {
        TIMESTAMP:              0,
        NAME:                   1,
        EMAIL_ID:               2,
        PHONE_NUMBER:           3,
        OUR_SERVICES:           4,
        NUMBER_OF_PARTICIPANTS: 5,
        CONSENT:                6,
        REG_NO:                 7,   // written by script
        REG_EMAIL_SENT:         8,   // written by script
        REG_EMAIL_TIMESTAMP:    9    // written by script
    },

    // ── NEW: Regular Students sheet columns ──
    // A:Reg No | B:Name | C:Email | D:Phone | E:Program | F:Participants
    // G:Total Classes | H:Classes Done | I:Payment Status
    // J:Amount Paid | K:Total Amount | L:Payment Ref | M:Enrolled On | N:Status
    STUDENT_COLS: {
        REG_NO:             0,
        NAME:               1,
        EMAIL:              2,
        PHONE:              3,
        PROGRAM:            4,
        PARTICIPANTS:       5,
        TOTAL_CLASSES:      6,
        CLASSES_COMPLETED:  7,
        PAYMENT_STATUS:     8,   // Pending / Partial / Paid
        AMOUNT_PAID:        9,
        TOTAL_AMOUNT:       10,
        PAYMENT_REF:        11,
        ENROLLED_ON:        12,
        STATUS:             13   // Active / Completed
    }
};

// --------------- UTILITY FUNCTIONS ---------------

function generateReference() {
    const date = new Date();
    const year = date.getFullYear().toString().substr(-2);
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const random = Math.floor(Math.random() * 9000) + 1000;
    return `KE${year}${month}${day}${random}`;
}

// ── NEW ──
function generateRegularRef() {
    const date = new Date();
    const year = date.getFullYear().toString().substr(-2);
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const random = Math.floor(Math.random() * 9000) + 1000;
    return `KER${year}${month}${day}${random}`;
}

function createUPILink(amount, reference) {
    return `upi://pay?pa=${CONFIG.UPI_ID}&pn=${encodeURIComponent(CONFIG.BUSINESS_NAME)}&am=${amount}&cu=INR&tn=${encodeURIComponent(reference)}`;
}

function getPdfTemplate() {
    try {
        const file = DriveApp.getFileById(CONFIG.EMAIL_TEMPLATE_PDF_ID);
        const blob = file.getBlob();
        return blob;
    } catch (error) {
        Logger.log('Error fetching PDF template: ' + error);
        return null;
    }
}

function createQRCode(link) {
    return `https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(link)}`;
}

function getPricingData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pricingSheet = ss.getSheetByName(CONFIG.SHEETS.PRICING);
    if (!pricingSheet) {
        throw new Error('Pricing sheet not found');
    }
    const data = pricingSheet.getDataRange().getValues();
    const pricingMap = {};
    for (let i = 1; i < data.length; i++) {
        const serviceId = data[i][0];
        const service = data[i][1];
        const pricePerHalfHour = data[i][2];
        const docId = data[i][3];
        if (service) {
            pricingMap[service] = {
                price: pricePerHalfHour,
                docId: docId
            };
        }
    }
    return pricingMap;
}

function getServiceDetailsFromDoc(docId) {
    try {
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();
        const text = body.getText();
        const summaryMatch = text.match(/summary[:\s]*([\s\S]*?)(?=\n\n|$)/i);
        const summary = summaryMatch ? summaryMatch[1].trim() : '';
        return { summary: summary, fullText: text };
    } catch (error) {
        Logger.log('Error fetching service details: ' + error);
        return { summary: 'Professional equestrian service at Kings Equestrian.', fullText: '' };
    }
}

function getServicePDF(docId, serviceName) {
    try {
        const doc = DocumentApp.openById(docId);
        const blob = doc.getAs('application/pdf');
        blob.setName(`${serviceName.replace(/\s+/g, '_')}_Details.pdf`);
        return blob;
    } catch (error) {
        Logger.log('Error creating PDF: ' + error);
        return null;
    }
}

function getTermsAndConditionsPDF() {
    try {
        const doc = DocumentApp.openById(CONFIG.TERMS_CONDITIONS_DOC_ID);
        const blob = doc.getAs('application/pdf');
        blob.setName('Terms_and_Conditions.pdf');
        return blob;
    } catch (error) {
        Logger.log('Error creating T&C PDF: ' + error);
        return null;
    }
}

function getCCRecipients(mailType) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const mailInfoSheet = ss.getSheetByName(CONFIG.SHEETS.MAIL_INFO);
        if (!mailInfoSheet) {
            Logger.log('Mail Info sheet not found');
            return [];
        }
        const data = mailInfoSheet.getDataRange().getValues();
        const ccEmails = [];
        for (let i = 1; i < data.length; i++) {
            const email = data[i][0];
            const type = data[i][1];
            if (email && type && type.toLowerCase().includes(mailType.toLowerCase())) {
                ccEmails.push(email);
            }
        }
        return ccEmails;
    } catch (error) {
        Logger.log('Error getting CC recipients: ' + error);
        return [];
    }
}

// --------------- BOOKING LOOKUP BY PHONE (latest first) ---------------

function findLatestBookingByPhone(phoneNumber, bookingValues) {
    const normalizedPhone = String(phoneNumber || '').trim().replace(/\D/g, '');
    if (!normalizedPhone) return null;

    for (let j = bookingValues.length - 1; j >= 1; j--) {
        const rowPhone = String(bookingValues[j][CONFIG.BOOKING_COLS.PHONE_NUMBER] || '')
            .trim()
            .replace(/\D/g, '');
        if (rowPhone === normalizedPhone) {
            return { rowIndex: j + 1, row: bookingValues[j] };
        }
    }
    return null;
}

// ── NEW: Find regular student by phone (last 10 digits match) ──
function findRegularStudentByPhone(phoneNumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.REGULAR_STUDENTS);
    if (!sheet) return null;

    const normalizedPhone = String(phoneNumber || '').replace(/\D/g, '').slice(-10);
    if (!normalizedPhone || normalizedPhone.length < 10) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        const rowPhone = String(data[i][CONFIG.STUDENT_COLS.PHONE] || '').replace(/\D/g, '').slice(-10);
        if (rowPhone === normalizedPhone) {
            return { rowIndex: i + 1, row: data[i] };
        }
    }
    return null;
}

// ── NEW: Find regular student by KER ref ──
function findRegularStudentByRef(regNo) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.REGULAR_STUDENTS);
    if (!sheet) return null;

    const trimmed = String(regNo || '').trim().toUpperCase();
    if (!trimmed.startsWith('KER')) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][CONFIG.STUDENT_COLS.REG_NO] || '').trim().toUpperCase() === trimmed) {
            return { rowIndex: i + 1, row: data[i] };
        }
    }
    return null;
}

// --------------- MAIN FORM SUBMIT HANDLER ---------------

function onBookingFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        if (sheet.getName() !== CONFIG.SHEETS.BOOKING_FORM) {
            Logger.log('onBookingFormSubmit: Skipping — wrong sheet: ' + sheet.getName());
            return;
        }
        const row = e.range.getRow();

        const name = sheet.getRange(row, CONFIG.BOOKING_COLS.NAME + 1).getValue();
        const email = sheet.getRange(row, CONFIG.BOOKING_COLS.EMAIL_ID + 1).getValue();
        const phone = sheet.getRange(row, CONFIG.BOOKING_COLS.PHONE_NUMBER + 1).getValue();
        const services = sheet.getRange(row, CONFIG.BOOKING_COLS.OUR_SERVICES + 1).getValue();
        const participants = Number(sheet.getRange(row, CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS + 1).getValue()) || 1;
        const bookingDate = sheet.getRange(row, CONFIG.BOOKING_COLS.TIMESTAMP + 1).getValue();

        const amount = CONFIG.ADVANCE_BOOKING_AMOUNT;
        const reference = generateReference();
        const upiLink = createUPILink(amount, reference);
        const qrCode = createQRCode(upiLink);

        sheet.getRange(row, CONFIG.BOOKING_COLS.REFERENCE + 1).setValue(reference);

        sendWelcomeEmail({
            name: name,
            email: email,
            phone: phone,
            services: services,
            participants: participants,
            amount: amount,
            reference: reference,
            upiLink: upiLink,
            qrCode: qrCode,
            row: row,
            sheet: sheet,
            bookingDate: bookingDate
        });

        Logger.log(`Booking processed successfully for ${name} - Reference: ${reference}`);
    } catch (error) {
        Logger.log('Error in onBookingFormSubmit: ' + error);
        Logger.log('Stack trace: ' + error.stack);
        SpreadsheetApp.getUi().alert('Error processing booking: ' + error.message);
    }
}

// ── NEW: Regular Booking Form Submit Handler ──────────────────────────────
//
// Triggered when a student submits the Regular Booking Google Form.
// Sheet: "Regular Booking Form Response"
// Columns match REG_BOOKING_COLS above:
//   0:Timestamp | 1:Name | 2:Email ID | 3:Phone Number |
//   4:Our Services | 5:No. of Participants | 6:Consent
//   7:Reg No (script writes) | 8:Email Sent (script writes) | 9:Email Timestamp (script writes)
//
function onRegularBookingFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        if (sheet.getName() !== CONFIG.SHEETS.REGULAR_BOOKING_FORM) {
            Logger.log('onRegularBookingFormSubmit: Skipping — wrong sheet: ' + sheet.getName());
            return;
        }

        const row = e.range.getRow();
        Logger.log('Regular booking form submitted at row ' + row);

        const C = CONFIG.REG_BOOKING_COLS;
        const name         = String(sheet.getRange(row, C.NAME + 1).getValue() || '').trim();
        const email        = String(sheet.getRange(row, C.EMAIL_ID + 1).getValue() || '').trim();
        const phone        = String(sheet.getRange(row, C.PHONE_NUMBER + 1).getValue() || '').trim();
        const serviceRaw   = String(sheet.getRange(row, C.OUR_SERVICES + 1).getValue() || '').trim();
        const participants = Number(sheet.getRange(row, C.NUMBER_OF_PARTICIPANTS + 1).getValue()) || 1;
        const timestamp    = sheet.getRange(row, C.TIMESTAMP + 1).getValue();

        if (!name || !email || !phone) {
            Logger.log('Regular booking: missing required fields at row ' + row);
            sheet.getRange(row, C.REG_NO + 1).setValue('⚠️ Missing fields');
            return;
        }

        // Generate KER registration number
        const regNo = generateRegularRef();
        Logger.log('Generated KER ref: ' + regNo + ' for ' + name);

        // Get pricing to find total amount for this program
        let totalAmount = 0;
        try {
            const pricingData = getPricingData();
            // Try to match service name against pricing sheet (case-insensitive partial match)
            for (const key of Object.keys(pricingData)) {
                if (serviceRaw.toLowerCase().includes(key.toLowerCase()) ||
                    key.toLowerCase().includes(serviceRaw.toLowerCase())) {
                    totalAmount = pricingData[key].price || 0;
                    break;
                }
            }
        } catch (pErr) {
            Logger.log('Could not fetch pricing: ' + pErr);
        }

        // Write to Regular Students sheet
        createRegularStudentRecord(regNo, name, email, phone, serviceRaw, participants, totalAmount, timestamp);

        // Write reg no back to the booking form sheet
        sheet.getRange(row, C.REG_NO + 1).setValue(regNo);

        // Send registration confirmation email
        sendRegularRegistrationEmail(name, email, phone, regNo, serviceRaw, participants, totalAmount);

        // Mark email sent
        sheet.getRange(row, C.REG_EMAIL_SENT + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');
        sheet.getRange(row, C.REG_EMAIL_TIMESTAMP + 1)
            .setValue(new Date())
            .setNumberFormat('dd-MMM-yyyy HH:mm:ss');

        Logger.log('Regular booking processed: ' + name + ' — ' + regNo + ' — ' + serviceRaw);

    } catch (err) {
        Logger.log('Error in onRegularBookingFormSubmit: ' + err);
        Logger.log('Stack: ' + err.stack);
    }
}

// ── NEW: Create Regular Student Record ──────────────────────────────────
function createRegularStudentRecord(regNo, name, email, phone, program, participants, totalAmount, enrolledOn) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.REGULAR_STUDENTS);

    // Auto-create sheet with headers if it doesn't exist
    if (!sheet) {
        sheet = ss.insertSheet(CONFIG.SHEETS.REGULAR_STUDENTS);
        const headers = [
            'Reg No', 'Name', 'Email', 'Phone', 'Program', 'Participants',
            'Total Classes', 'Classes Completed', 'Payment Status',
            'Amount Paid (₹)', 'Total Amount (₹)', 'Payment Ref', 'Enrolled On', 'Status'
        ];
        sheet.getRange(1, 1, 1, headers.length)
            .setValues([headers])
            .setFontWeight('bold')
            .setBackground('#1f4e3d')
            .setFontColor('#ffffff');
        sheet.setFrozenRows(1);
        sheet.setColumnWidths(1, headers.length, 150);
        Logger.log('Created Regular Students sheet');
    }

    const C = CONFIG.STUDENT_COLS;
    const rowData = new Array(14).fill('');
    rowData[C.REG_NO]            = regNo;
    rowData[C.NAME]              = name;
    rowData[C.EMAIL]             = email;
    rowData[C.PHONE]             = phone;
    rowData[C.PROGRAM]           = program;
    rowData[C.PARTICIPANTS]      = participants;
    rowData[C.TOTAL_CLASSES]     = '';      // filled once admin sets the program
    rowData[C.CLASSES_COMPLETED] = 0;
    rowData[C.PAYMENT_STATUS]    = 'Pending';
    rowData[C.AMOUNT_PAID]       = 0;
    rowData[C.TOTAL_AMOUNT]      = totalAmount || '';
    rowData[C.PAYMENT_REF]       = '';
    rowData[C.ENROLLED_ON]       = enrolledOn || new Date();
    rowData[C.STATUS]            = 'Active';

    sheet.appendRow(rowData);

    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, C.ENROLLED_ON + 1).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
    sheet.getRange(newRow, C.PAYMENT_STATUS + 1)
        .setBackground('#fff3cd')
        .setFontColor('#856404')
        .setFontWeight('bold');

    Logger.log('Regular student record created: ' + regNo + ' — ' + name);
}

// ── NEW: Registration Confirmation Email ────────────────────────────────
function sendRegularRegistrationEmail(name, email, phone, regNo, program, participants, totalAmount) {
    const upiLink = totalAmount > 0 ? createUPILink(totalAmount, regNo) : '';
    const qrCode  = upiLink ? createQRCode(upiLink) : '';
    const subject = '🐴 Registration Confirmed — Kings Equestrian | Reg No: ' + regNo;

    const paymentBlock = totalAmount > 0 ? `
    <div style="background:#e8f5e9;border:2px solid #4caf50;padding:24px;border-radius:10px;margin:24px 0;">
      <h3 style="margin:0 0 8px;color:#1b5e20;font-size:17px;">💳 Complete Your Payment</h3>
      <p style="color:#2e7d32;font-size:13px;margin:0 0 16px;">
        Pay the full program fee (or a partial amount) to activate your account and start booking classes.
        Partial payments are also accepted.
      </p>
      <p style="text-align:center;font-size:36px;font-weight:700;color:#0F6E56;margin:0 0 20px;">
        ₹${totalAmount.toLocaleString('en-IN')}
      </p>
      ${qrCode ? `
      <div style="text-align:center;margin-bottom:20px;">
        <p style="font-size:13px;font-weight:600;color:#1b5e20;margin:0 0 10px;">Scan to Pay via UPI</p>
        <img src="${qrCode}" alt="UPI QR Code" style="width:160px;height:160px;border:2px solid #c8e6c9;border-radius:8px;">
      </div>` : ''}
      <div style="text-align:center;margin-bottom:16px;">
        <a href="${CONFIG.PAYMENT_FORM_LINK}"
           style="display:inline-block;background:#0F6E56;color:white;padding:13px 28px;text-decoration:none;border-radius:7px;font-weight:600;font-size:14px;">
          📝 Submit Payment Confirmation
        </a>
      </div>
      <div style="background:#fff8e1;border-left:4px solid #ffc107;padding:12px 14px;border-radius:0 6px 6px 0;font-size:12px;color:#5d4037;line-height:1.6;">
        <strong>⚠️ Important:</strong> After paying via UPI, submit the payment form above with your
        <strong>Registration No. ${regNo}</strong> (or your phone number ${phone}).
        Your 80G receipt will be sent automatically.
      </div>
    </div>` : `
    <div style="background:#fff3cd;border-left:4px solid #ffc107;padding:14px;border-radius:6px;margin:20px 0;">
      <p style="margin:0;font-size:13px;color:#856404;">
        <strong>Payment details</strong> will be shared with you shortly.
        Please contact us at +91 99807 71166 for pricing information.
      </p>
    </div>`;

    const htmlBody = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{margin:0;padding:0;background:#f0f4f1;font-family:'Segoe UI',Arial,sans-serif;color:#1a2e1e}
  .wrap{max-width:620px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(4,52,44,0.10)}
  .hdr{background:linear-gradient(135deg,#1f4e3d 0%,#4f9c7a 100%);padding:32px 28px 24px;text-align:center}
  .hdr img{width:80px;height:80px;border-radius:50%;border:3px solid rgba(255,255,255,0.4);margin-bottom:14px;display:block;margin-left:auto;margin-right:auto}
  .hdr h1{margin:0;color:#C8EFE3;font-size:24px;font-weight:700}
  .hdr p{margin:8px 0 0;color:rgba(255,255,255,0.88);font-size:13px}
  .body{padding:28px}
  .reg-box{background:#f0f9f5;border:2px solid #9FE1CB;border-radius:10px;padding:20px;text-align:center;margin:0 0 24px}
  .reg-label{font-size:11px;letter-spacing:0.1em;text-transform:uppercase;color:#5a7a61;margin-bottom:6px}
  .reg-no{font-family:Georgia,serif;font-size:32px;font-weight:700;color:#04342C;letter-spacing:0.04em;margin:0 0 6px}
  .reg-hint{font-size:12px;color:#7a9a7e;line-height:1.6}
  h2{color:#1f4e3d;font-size:17px;font-weight:700;margin:24px 0 10px;border-bottom:2px solid #e8f5e9;padding-bottom:8px}
  table.prog{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px}
  table.prog tr{border-bottom:1px solid #f0f4f1}
  table.prog td{padding:8px 4px}
  .prog-lbl{color:#7a9a7e;width:40%}
  .portal-box{background:#1f4e3d;border-radius:10px;padding:20px 22px;margin:24px 0;text-align:center}
  .portal-box p{color:#9FE1CB;font-size:13px;margin:0 0 14px;line-height:1.5}
  .portal-btn{display:inline-block;background:#1D9E75;color:#fff;padding:12px 26px;text-decoration:none;border-radius:7px;font-weight:600;font-size:14px}
  .next-steps{background:#f7faf8;border-radius:8px;padding:16px 18px;margin:20px 0}
  .next-steps ol{margin:8px 0 0;padding-left:20px;font-size:13px;line-height:2.1;color:#3a5a3e}
  .ftr{background:#1f4e3d;padding:16px;text-align:center;font-size:11px;color:rgba(255,255,255,0.75);line-height:1.7}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian">
    <h1>Registration Confirmed! 🎉</h1>
    <p>Kings Equestrian Foundation — Regular Rider Program</p>
  </div>
  <div class="body">
    <p style="font-size:15px;">Dear <strong>${name}</strong>,</p>
    <p style="font-size:13px;color:#5a7a61;line-height:1.6;">
      Welcome to Kings Equestrian Foundation! Your registration has been received and
      your unique Registration Number has been generated. Please save this number —
      you will use it for payments, class bookings, and logging into the rider portal.
    </p>

    <div class="reg-box">
      <div class="reg-label">Your Registration Number</div>
      <div class="reg-no">${regNo}</div>
      <div class="reg-hint">
        Use this when submitting payments, booking classes, or logging into the portal.<br>
        Your phone <strong>${phone}</strong> also works everywhere.
      </div>
    </div>

    <h2>📋 Enrollment Details</h2>
    <table class="prog">
      <tr><td class="prog-lbl">Program</td><td><strong>${program}</strong></td></tr>
      <tr><td class="prog-lbl">Participants</td><td>${participants}</td></tr>
      ${totalAmount > 0 ? `<tr><td class="prog-lbl">Program Fee</td><td><strong style="color:#0F6E56;font-size:16px;">₹${totalAmount.toLocaleString('en-IN')}</strong></td></tr>` : ''}
    </table>

    ${paymentBlock}

    <div class="portal-box">
      <p>Once payment is made, log in to your <strong style="color:#C8EFE3;">Rider Portal</strong>
         to book class slots, view your schedule, and reschedule classes.</p>
      <a href="${CONFIG.RIDER_PORTAL_URL}" class="portal-btn">🐴 Open Rider Portal</a>
    </div>

    <div class="next-steps">
      <strong style="color:#1f4e3d;font-size:13px;">📌 What Happens Next?</strong>
      <ol>
        <li>Pay via UPI using the QR code above</li>
        <li>Submit the <a href="${CONFIG.PAYMENT_FORM_LINK}" style="color:#0F6E56;">payment form</a> with Reg No: <strong>${regNo}</strong></li>
        <li>Receive your 80G tax receipt instantly by email</li>
        <li>Log into the <a href="${CONFIG.RIDER_PORTAL_URL}" style="color:#0F6E56;">Rider Portal</a> to book class slots</li>
        <li>Get automatic reminders before each class</li>
        <li>Need to move a class? Use the reschedule option in the portal</li>
      </ol>
    </div>

    <p style="font-size:12px;color:#7a9a7e;margin-top:20px;">
      Questions? Call or WhatsApp us at <strong>+91 99807 71166</strong> or email
      <a href="mailto:info@kingsequestrian.com" style="color:#0F6E56;">info@kingsequestrian.com</a>
    </p>
  </div>
  <div class="ftr">
    <strong style="color:#fff;">Kings Equestrian Foundation</strong><br>
    Just 30 mins from Bengaluru · Karnataka, India<br>
    📞 +91 99807 71166 · 📸 @kingsequestrianfoundation<br>
    <a href="${CONFIG.RIDER_PORTAL_URL}" style="color:#9FE1CB;">Rider Portal</a> ·
    <a href="${CONFIG.PAYMENT_FORM_LINK}" style="color:#9FE1CB;">Payment Form</a>
  </div>
</div>
</body>
</html>`;

    const ccEmails = getCCRecipients('Regular Registration');
    MailApp.sendEmail({
        to: email,
        cc: ccEmails.join(','),
        subject: subject,
        htmlBody: htmlBody,
        name: 'Kings Equestrian Foundation'
    });
    Logger.log('Registration email sent: ' + name + ' (' + regNo + ') → ' + email);
}

// --------------- PAYMENT FORM SUBMIT HANDLER ---------------

function onPaymentFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        if (sheet.getName() !== CONFIG.SHEETS.PAYMENT_FORM) {
            Logger.log('onPaymentFormSubmit: Skipping — wrong sheet: ' + sheet.getName());
            return;
        }
        const row = e.range.getRow();

        Logger.log(`Payment form submitted at row ${row}`);

        const phoneNumber    = String(sheet.getRange(row, CONFIG.PAYMENT_COLS.PHONE_NUMBER + 1).getValue() || '').trim();
        const regNoSubmitted = String(sheet.getRange(row, CONFIG.PAYMENT_COLS.REGISTRATION_NO + 1).getValue() || '').trim();
        const amount         = Number(sheet.getRange(row, CONFIG.PAYMENT_COLS.AMOUNT_PAID + 1).getValue());
        const paymentDate    = sheet.getRange(row, CONFIG.PAYMENT_COLS.PAYMENT_DATE + 1).getValue();
        const timestamp      = sheet.getRange(row, CONFIG.PAYMENT_COLS.TIMESTAMP + 1).getValue();

        if (!phoneNumber && !regNoSubmitted) {
            Logger.log('No phone or reg no found in payment form submission');
            return;
        }

        Logger.log(`Processing payment — phone: ${phoneNumber}, regNo: ${regNoSubmitted}, amount: ${amount}`);

        // ── Determine if this is a regular student payment ──────────────────
        // Check 1: If they submitted a KER ref → definitely regular
        // Check 2: If phone matches a record in Regular Students sheet → regular
        // Otherwise → existing one-off booking path
        let isRegularStudent = false;
        let regularMatch     = null;

        if (regNoSubmitted.toUpperCase().startsWith('KER')) {
            regularMatch = findRegularStudentByRef(regNoSubmitted);
            if (regularMatch) {
                isRegularStudent = true;
                Logger.log('Regular student identified by KER ref: ' + regNoSubmitted);
            }
        }

        if (!isRegularStudent && phoneNumber) {
            regularMatch = findRegularStudentByPhone(phoneNumber);
            if (regularMatch) {
                isRegularStudent = true;
                Logger.log('Regular student identified by phone: ' + phoneNumber);
            }
        }

        // ── Duplicate check (existing logic, unchanged) ──────────────────────
        const duplicateInfo = findDuplicateReceipt(phoneNumber, amount, paymentDate, timestamp);

        if (duplicateInfo.isDuplicate) {
            Logger.log(`Duplicate receipt detected for phone ${phoneNumber}. Resending existing receipt.`);
            sheet.getRange(row, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
                .setValue('Duplicate - Resent')
                .setBackground('#fff3cd')
                .setFontColor('#856404');
            resendExistingReceipt(row, duplicateInfo.existingRow);
            return;
        }

        // ── Auto-verify ───────────────────────────────────────────────────────
        sheet.getRange(row, CONFIG.PAYMENT_COLS.TRANSACTION_VERIFIED + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        Logger.log('Transaction auto-verified, proceeding to send receipt');
        Utilities.sleep(500);
        sendReceiptForRow(row);

        // ── If regular student: update payment balance in Regular Students sheet ──
        if (isRegularStudent && regularMatch) {
            updateRegularStudentPayment(regularMatch, amount);
        }

    } catch (error) {
        Logger.log('Error in onPaymentFormSubmit: ' + error);
        Logger.log('Stack trace: ' + error.stack);
    }
}

// ── NEW: Update Regular Student Payment Balance ──────────────────────────
function updateRegularStudentPayment(studentMatch, amountPaid) {
    try {
        const ss    = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEETS.REGULAR_STUDENTS);
        if (!sheet) return;

        const C           = CONFIG.STUDENT_COLS;
        const sheetRow    = studentMatch.rowIndex;
        const rowData     = studentMatch.row;

        const previouslyPaid = Number(rowData[C.AMOUNT_PAID]) || 0;
        const totalAmount    = Number(rowData[C.TOTAL_AMOUNT]) || 0;
        const newAmountPaid  = previouslyPaid + amountPaid;

        let newStatus;
        if (totalAmount <= 0) {
            newStatus = 'Paid';
        } else if (newAmountPaid >= totalAmount) {
            newStatus = 'Paid';
        } else {
            newStatus = 'Partial';
        }

        const colors = {
            'Paid':    { bg: '#d4edda', fg: '#155724' },
            'Partial': { bg: '#fff3cd', fg: '#856404' },
            'Pending': { bg: '#f8d7da', fg: '#721c24' }
        };
        const color = colors[newStatus];

        sheet.getRange(sheetRow, C.AMOUNT_PAID + 1).setValue(newAmountPaid);
        sheet.getRange(sheetRow, C.PAYMENT_STATUS + 1)
            .setValue(newStatus)
            .setBackground(color.bg)
            .setFontColor(color.fg)
            .setFontWeight('bold');

        Logger.log('Regular student payment updated: ₹' + previouslyPaid + ' + ₹' + amountPaid
            + ' = ₹' + newAmountPaid + ' / ₹' + totalAmount + ' (' + newStatus + ')');

        // If fully paid and this is the first full payment — send a "you're all set" email
        if (newStatus === 'Paid' && previouslyPaid < totalAmount && totalAmount > 0) {
            const studentEmail = rowData[C.EMAIL];
            const studentName  = rowData[C.NAME];
            const regNo        = rowData[C.REG_NO];
            const program      = rowData[C.PROGRAM];
            if (studentEmail) {
                sendRegularPaymentCompleteEmail(studentName, studentEmail, regNo, program, newAmountPaid);
            }
        }

    } catch (err) {
        Logger.log('updateRegularStudentPayment error: ' + err);
    }
}

// ── NEW: Full Payment Complete Email for Regular Students ────────────────
function sendRegularPaymentCompleteEmail(name, email, regNo, program, amountPaid) {
    const subject = '✅ Full Payment Received — Start Booking Your Classes! (' + regNo + ')';
    const htmlBody = `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  body{margin:0;padding:0;background:#f0f4f1;font-family:'Segoe UI',Arial,sans-serif}
  .wrap{max-width:560px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(4,52,44,0.10)}
  .hdr{background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:26px;text-align:center}
  .hdr img{width:68px;height:68px;border-radius:50%;border:2px solid rgba(255,255,255,0.4);display:block;margin:0 auto 10px}
  .hdr h1{margin:0;color:#C8EFE3;font-size:20px}
  .body{padding:24px}
  .banner{background:#e8f5e9;border:2px solid #4caf50;border-radius:10px;padding:16px;text-align:center;margin:0 0 20px}
  .banner h2{margin:0 0 4px;color:#1b5e20;font-size:18px}
  .dr{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #f0f4f1;font-size:13px}
  .portal-box{background:#1f4e3d;border-radius:10px;padding:18px;text-align:center;margin:20px 0}
  .portal-box p{color:#9FE1CB;font-size:13px;margin:0 0 12px}
  .portal-btn{display:inline-block;background:#1D9E75;color:#fff;padding:11px 24px;text-decoration:none;border-radius:7px;font-weight:600;font-size:13px}
  .ftr{background:#1f4e3d;padding:14px;text-align:center;font-size:11px;color:rgba(255,255,255,0.75);border-top:none}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian">
    <h1>Payment Complete! 🎉</h1>
  </div>
  <div class="body">
    <p>Dear <strong>${name}</strong>,</p>
    <p style="font-size:13px;color:#5a7a61;">Your full payment has been received. You are all set to start booking your classes!</p>
    <div class="banner">
      <h2>🎉 You're Fully Paid!</h2>
      <p style="margin:0;font-size:13px;color:#2e7d32;">Your ${program} account is now fully active.</p>
    </div>
    <div class="dr"><span style="color:#7a9a7e">Registration No.</span><strong>${regNo}</strong></div>
    <div class="dr"><span style="color:#7a9a7e">Program</span><span>${program}</span></div>
    <div class="dr"><span style="color:#7a9a7e">Total Paid</span><strong style="color:#0F6E56">₹${amountPaid.toLocaleString('en-IN')}</strong></div>
    <div class="portal-box">
      <p>Log in to your <strong style="color:#C8EFE3;">Rider Portal</strong> to book your class slots now.</p>
      <a href="${CONFIG.RIDER_PORTAL_URL}" class="portal-btn">🐴 Book My Classes</a>
    </div>
    <p style="font-size:12px;color:#7a9a7e;">
      Need help? Call +91 99807 71166 or email
      <a href="mailto:info@kingsequestrian.com" style="color:#0F6E56;">info@kingsequestrian.com</a>
    </p>
  </div>
  <div class="ftr">Kings Equestrian Foundation · Karnataka · +91 99807 71166</div>
</div>
</body>
</html>`;

    MailApp.sendEmail({ to: email, subject, htmlBody, name: 'Kings Equestrian Foundation' });
    Logger.log('Payment complete email sent to ' + email + ' (' + regNo + ')');
}

// --------------- DUPLICATE DETECTION ---------------

function findDuplicateReceipt(phoneNumber, amount, paymentDate, currentTimestamp) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);

        if (!paymentSheet) {
            return { isDuplicate: false, existingRow: null };
        }

        const data = paymentSheet.getDataRange().getValues();
        const normalizedDate = normalizeDate(paymentDate);
        const normalizedCurrentTimestamp = normalizeDate(currentTimestamp);
        const normalizedPhone = String(phoneNumber || '').trim().replace(/\D/g, '');

        for (let i = 1; i < data.length; i++) {
            const rowPhone = String(data[i][CONFIG.PAYMENT_COLS.PHONE_NUMBER] || '')
                .trim()
                .replace(/\D/g, '');
            const rowAmount = Number(data[i][CONFIG.PAYMENT_COLS.AMOUNT_PAID]);
            const rowDate = data[i][CONFIG.PAYMENT_COLS.PAYMENT_DATE];
            const rowTimestamp = data[i][CONFIG.PAYMENT_COLS.TIMESTAMP];
            const rowReceiptSent = String(data[i][CONFIG.PAYMENT_COLS.RECEIPT_SENT] || '').trim();

            if (normalizeDate(rowTimestamp) === normalizedCurrentTimestamp) {
                continue;
            }

            if (rowPhone === normalizedPhone &&
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

function normalizeDate(dateValue) {
    if (!dateValue) return '';
    try {
        const date = new Date(dateValue);
        if (isNaN(date.getTime())) return String(dateValue);
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch (e) {
        return String(dateValue);
    }
}

// --------------- RESEND EXISTING RECEIPT ---------------

function resendExistingReceipt(currentRow, existingRow) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
        const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);

        if (!paymentSheet || !bookingSheet) {
            Logger.log('Required sheets not found');
            return false;
        }

        const existingData = paymentSheet.getRange(existingRow, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
        const existingReceiptNumber = existingData[CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_NO];
        const existingDriveLink = existingData[CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_DRIVER_LINK];

        const currentData = paymentSheet.getRange(currentRow, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
        const phoneNumber = String(currentData[CONFIG.PAYMENT_COLS.PHONE_NUMBER] || '').trim();
        const amount = Number(currentData[CONFIG.PAYMENT_COLS.AMOUNT_PAID]);
        const transactionId = currentData[CONFIG.PAYMENT_COLS.TRANSACTION_REFERENCE_NUMBER] || '';
        const pan = currentData[CONFIG.PAYMENT_COLS.PAN_AADHAAR] || '';

        const bookingValues = bookingSheet.getDataRange().getValues();
        const bookingMatch = findLatestBookingByPhone(phoneNumber, bookingValues);

        if (!bookingMatch) throw new Error(`Booking not found for phone ${phoneNumber}`);

        const riderName = bookingMatch.row[CONFIG.BOOKING_COLS.NAME];
        const email = bookingMatch.row[CONFIG.BOOKING_COLS.EMAIL_ID];
        const referenceNumber = bookingMatch.row[CONFIG.BOOKING_COLS.REFERENCE];
        const preferredDate = bookingMatch.row[CONFIG.BOOKING_COLS.PREFERRED_SERVICE_DATE];
        const preferredTimeSlots = bookingMatch.row[CONFIG.BOOKING_COLS.PREFERRED_TIME_SLOT];

        if (!email) throw new Error('Email not found in booking');

        const receiptPDF = generate80GReceipt(riderName, pan, amount, transactionId, existingReceiptNumber);

        const subject = `Payment Receipt - ${riderName} - Ref: ${referenceNumber}`;
        const htmlBody = `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head><body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; max-width: 650px; margin: 0 auto; padding: 20px;"><div style="text-align: center; padding: 30px 0; background: linear-gradient(135deg, #4caf50 0%, #45a049 100%); border-radius: 12px 12px 0 0;"><h1 style="color: white; margin: 0; font-size: 28px;">Kings Equestrian Foundation</h1><p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0; font-style: italic;">Where horses don't just carry you - they change you</p></div><div style="background: white; padding: 30px; border: 1px solid #e0e0e0; border-top: none;"><p style="font-size: 16px; margin-bottom: 25px;">Dear <strong>${riderName}</strong>,</p><div style="background: #e8f5e9; border-left: 4px solid #4caf50; padding: 20px; margin: 20px 0; border-radius: 4px; text-align: center;"><h2 style="color: #2e7d32; margin: 0 0 10px 0;">✅ Payment Confirmed - Booking Complete!</h2><p style="margin: 0; font-size: 14px;">Thank you for your payment. Your booking is confirmed.</p></div><p style="font-size: 14px; margin: 20px 0;"><strong>Your Payment Receipt (80G) is attached to this email for tax deduction purposes.</strong></p><h3 style="color: #2c3e50; border-bottom: 2px solid #4caf50; padding-bottom: 10px; margin-top: 25px;">Payment Details:</h3><table style="width: 100%; margin: 15px 0;"><tr><td style="padding: 8px 0; color: #666;">Booking Reference:</td><td style="padding: 8px 0; font-weight: bold;">${referenceNumber}</td></tr><tr><td style="padding: 8px 0; color: #666;">Receipt No:</td><td style="padding: 8px 0; font-weight: bold;">${existingReceiptNumber}</td></tr><tr><td style="padding: 8px 0; color: #666;">Amount Paid:</td><td style="padding: 8px 0; font-weight: bold; color: #4caf50; font-size: 18px;">₹${amount.toLocaleString('en-IN')}</td></tr>${transactionId ? `<tr><td style="padding: 8px 0; color: #666;">Transaction ID:</td><td style="padding: 8px 0; font-weight: bold;">${transactionId}</td></tr>` : ''}${preferredDate ? `<tr><td style="padding: 8px 0; color: #666;">Scheduled Date:</td><td style="padding: 8px 0; font-weight: bold;">${formatDate(preferredDate)}</td></tr>` : ''}${preferredTimeSlots ? `<tr><td style="padding: 8px 0; color: #666;">Time Slot:</td><td style="padding: 8px 0; font-weight: bold;">${preferredTimeSlots}</td></tr>` : ''}</table><div style="background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 25px 0; border-radius: 4px;"><p style="margin: 0; color: #856404;"><strong>We look forward to welcoming you at Kings Equestrian. Please arrive 15 minutes before your scheduled time.</strong></p></div><h4 style="color: #2c3e50; margin-top: 25px;">What to bring:</h4><ul style="margin: 10px 0; padding-left: 20px; color: #666;"><li>Comfortable clothing</li><li>Closed-toe shoes</li><li>Your booking reference: <strong>${referenceNumber}</strong></li></ul></div><div style="background: #f8f9fa; padding: 20px; text-align: center; border-radius: 0 0 12px 12px; border: 1px solid #e0e0e0; border-top: none;"><p style="margin: 5px 0; color: #666; font-size: 14px;"><strong>Kings Equestrian Foundation</strong></p><p style="margin: 5px 0; color: #666; font-size: 13px;">Karnataka, India</p><p style="margin: 5px 0; color: #666; font-size: 13px;">+91-9980895533 | info@kingsequestrian.com</p></div></body></html>`;

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

// --------------- ENHANCED EMAIL FUNCTIONS ---------------

function sendWelcomeEmail(data) {
    const subject = `Welcome to Kings Equestrian - Booking ${data.reference}`;
    const participants = data.participants || 1;

    const attachments = [];

    const termsPDF = getTermsAndConditionsPDF();
    if (termsPDF) attachments.push(termsPDF);

    const detailsPdf = getPdfTemplate();
    if (detailsPdf) attachments.push(detailsPdf);

    try {
        const consentPDF = generateConsentPDF(data.name, data.email, data.phone, data.bookingDate);
        if (consentPDF) {
            attachments.push(consentPDF);
            Logger.log('Consent form PDF generated and added to attachments');
        }
    } catch (error) {
        Logger.log('Error generating consent PDF: ' + error);
    }

    const pricingData = getPricingData();

    const rawServices = Array.isArray(data.services)
        ? data.services.join(', ')
        : String(data.services || '');

    const serviceList = Object.keys(pricingData).filter(key =>
        rawServices.toLowerCase().includes(key.toLowerCase())
    );

    serviceList.forEach(key => {
        const pricing = pricingData[key];
        if (pricing && pricing.docId) {
            const pdf = getServicePDF(pricing.docId, key);
            if (pdf) attachments.push(pdf);
        }
    });

    const servicesHTML = serviceList.map(s => `<li>${s}</li>`).join('');

    const serviceDetailsHTML = ` <div style="margin: 15px 0; padding: 15px; background: #f9f9f9; border-radius: 8px;">
      <h3 style="color: #2c5f2d; margin: 0 0 10px 0;">Selected Services</h3>
      <ul style="margin: 0; padding-left: 20px; font-size: 14px;">
                ${servicesHTML}
            </ul>
            <p style="color: #666; font-size: 14px; margin-bottom: 0;">See attached PDFs for detailed service information</p>
        </div>
    `;

    const paymentSection = `
         <div style="background: #e8f5e9; border: 2px solid #4caf50; padding: 20px; border-radius: 8px; margin: 20px 0;">
      <h3 style="margin-top: 0; color: #2e7d32;">💳 Reserve Your Slot</h3>
      <p style="font-size: 16px;">To confirm your booking, please pay the advance amount:</p>
      <p style="text-align: center; font-size: 32px; font-weight: bold; color: #2c5f2d; margin: 15px 0;">
        ₹${data.amount.toLocaleString('en-IN')}
      </p>
      <p style="text-align: center; font-size: 11px; color: #666; margin: 10px 0; font-style: italic;">
        This advance amount is non-refundable and can be used towards any Kings Equestrian service.
      </p>
      
      <div style=" margin: 25px 0;">
        <div style="flex: 1; min-width: 180px; background: white; padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
          <p style="margin: 0 0 12px 0; font-weight: bold; font-size: 14px; color: #2c5f2d;">Scan to Pay</p>
          <img src="${data.qrCode}" alt="QR Code" style="width: 150px; height: 150px; border: 2px solid #e0e0e0; border-radius: 4px;">
        </div>
        
        <div style="flex: 1; min-width: 180px; text-align: center;">
          <p style="margin: 0 0 15px 0; font-size: 14px; color: #333;">After making payment:</p>
          <a href="${CONFIG.PAYMENT_FORM_LINK}" 
             style="display: inline-block; background: #2c5f2d; color: white; padding: 14px 28px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 15px; box-shadow: 0 3px 6px rgba(44,95,45,0.3); transition: all 0.3s;">
            📝 Submit Payment & Select Slot
          </a>
          <p style="margin: 12px 0 0 0; font-size: 11px; color: #666; font-style: italic;">
            Don't forget to select your preferred date & time!
          </p>
        </div>
      </div>
      
      <div style="background: #fff3cd; padding: 15px; border-radius: 5px; margin-top: 15px; border-left: 4px solid #ffc107;">
        <p style="margin: 0; font-size: 13px; line-height: 1.6;">
          <strong>⚠️ Important:</strong> After scanning the QR code and making payment, click the button above to submit your payment screenshot, transaction details, and select your preferred date & time slot.
        </p>
      </div>
    </div>
  `;

    const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6; margin: 0; padding: 0; background: #f5f5f5;">
      <div style="max-width: 650px; margin: 20px auto; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
        
        <div style="background: linear-gradient(135deg, #1f4e3d 0%, #4f9c7a 100%); padding: 30px; text-align: center; color: white;">
          <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian" style="width: 80px; height: 80px; border-radius: 50%; margin-bottom: 15px;">
          <h1 style="margin: 0; font-size: 28px;">Welcome to Kings Equestrian!</h1>
          <p style="margin: 10px 0 0 0; font-size: 14px; opacity: 0.9;">Where horses don't just carry you — they change you</p>
        </div>
        
        <div style="padding: 30px;">
          <h2 style="color: #2c5f2d; margin-top: 0;">Hello ${data.name}! 👋</h2>
          
          <p>Thank you for choosing Kings Equestrian Foundation. Your booking request has been received.</p>
          
          <div style="background: #f0f8ff; border-left: 4px solid #2c5f2d; padding: 15px; margin: 20px 0;">
            <p style="margin: 0; font-size: 14px;">
              <strong>Booking Reference:</strong> <span style="font-size: 18px; color: #2c5f2d; font-weight: bold;">${data.reference}</span><br>
              <strong>Participants:</strong> ${participants}
            </p>
          </div>
          
          <h3 style="color: #2c5f2d; border-bottom: 2px solid #2c5f2d; padding-bottom: 10px;">📋 Service Details</h3>
          ${serviceDetailsHTML}
          
          ${paymentSection}
          
          <div style="background: #f9f9f9; padding: 20px; border-radius: 8px; margin-top: 20px;">
            <h4 style="margin: 0 0 10px 0; color: #2c5f2d;">📌 What's Next?</h4>
            <ul style="margin: 0; padding-left: 20px;">
              <li>Pay the advance booking fee of ₹${data.amount.toLocaleString('en-IN')}</li>
              <li>Submit payment confirmation and select your preferred date & time through the form</li>
              <li>Review the Terms & Conditions (attached)</li>
              <li>Wait for our confirmation email with your receipt</li>
              <li>Arrive 15 minutes before your scheduled time</li>
            </ul>
          </div>
          
          <p style="margin-top: 20px; font-size: 14px; color: #666;">
            If you have any questions, feel free to reach out to us anytime.
          </p>
        </div>
        
        <div style="background: #1f4e3d; color: white; padding: 20px; text-align: center; font-size: 13px;">
          <p style="margin: 0 0 10px 0;"><strong>Kings Equestrian Foundation</strong></p>
          <p style="margin: 0;">📍 Karnataka, India</p>
          <p style="margin: 5px 0;">📞 +91-9980895533 | ✉️ info@kingsequestrian.com</p>
          <p style="margin: 10px 0 0 0; opacity: 0.8; font-size: 11px;">
            © ${new Date().getFullYear()} Kings Equestrian Foundation. All rights reserved.
          </p>
        </div>
      </div>
    </body>
    </html>
  `;

    const plainBody = `
Welcome to Kings Equestrian Foundation!

Dear ${data.name},

Your booking reference: ${data.reference}

BOOKING DETAILS:
Name: ${data.name}
Contact: ${data.phone}
Services: ${data.services}
Participants: ${participants}

ADVANCE BOOKING AMOUNT: ₹${data.amount.toLocaleString('en-IN')}
(Non-refundable - Can be used towards any Kings Equestrian service)

PAYMENT INSTRUCTIONS:
1. Pay ₹${data.amount.toLocaleString('en-IN')} using UPI
2. Scan QR code or use UPI link
3. Submit payment details and select your preferred date & time: ${CONFIG.PAYMENT_FORM_LINK}

WHAT'S NEXT:
- Pay the advance booking fee
- Submit payment confirmation through the form
- Select your preferred date and time slot
- Review the Terms & Conditions (attached)
- Wait for our confirmation email with receipt
- Arrive 15 minutes before your scheduled time

Kings Equestrian Foundation
Karnataka, India
+91-9980895533 | info@kingsequestrian.com
  `;

    const ccEmails = getCCRecipients('Welcome Mail');

    MailApp.sendEmail({
        to: data.email,
        cc: ccEmails.join(','),
        subject: subject,
        body: plainBody,
        htmlBody: htmlBody,
        attachments: attachments,
        name: 'Kings Equestrian Foundation'
    });

    if (data.sheet && data.row) {
        data.sheet.getRange(data.row, CONFIG.BOOKING_COLS.WELCOME_EMAIL_SENT + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        data.sheet.getRange(data.row, CONFIG.BOOKING_COLS.WELCOME_EMAIL_TIMESTAMP + 1)
            .setValue(new Date())
            .setNumberFormat('dd-MMM-yyyy HH:mm:ss');
    }

    Logger.log(`Welcome email sent to: ${data.email} with CC to: ${ccEmails.join(', ')}`);
}

// --------------- RECEIPT GENERATION ---------------

function generateReceiptNumber(referenceNumber) {
    const serialMatch = referenceNumber.match(/\d{4}$/);
    const serial = serialMatch ? serialMatch[0] : '0000';
    return `${referenceNumber}/${serial}`;
}

function getImageAsBase64(fileId) {
    try {
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        const base64 = Utilities.base64Encode(blob.getBytes());
        const mimeType = blob.getContentType();
        return `data:${mimeType};base64,${base64}`;
    } catch (error) {
        Logger.log('Error getting image: ' + error);
        return '';
    }
}

function getImageFromUrlAsBase64(url) {
    try {
        const response = UrlFetchApp.fetch(url);
        const blob = response.getBlob();
        const base64 = Utilities.base64Encode(blob.getBytes());
        const mimeType = blob.getContentType();
        return `data:${mimeType};base64,${base64}`;
    } catch (error) {
        Logger.log('Error fetching image from URL: ' + error);
        return '';
    }
}

function numberToWords(num) {
    const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

    function convert(num) {
        if (num === 0) return 'Zero';
        if (num < 10) return ones[num];
        if (num < 20) return teens[num - 10];
        if (num < 100) return tens[Math.floor(num / 10)] + (num % 10 ? ' ' + ones[num % 10] : '');
        if (num < 1000) return ones[Math.floor(num / 100)] + ' Hundred' + (num % 100 ? ' ' + convert(num % 100) : '');
        if (num < 100000) return convert(Math.floor(num / 1000)) + ' Thousand' + (num % 1000 ? ' ' + convert(num % 1000) : '');
        if (num < 10000000) return convert(Math.floor(num / 100000)) + ' Lakh' + (num % 100000 ? ' ' + convert(num % 100000) : '');
        return convert(Math.floor(num / 10000000)) + ' Crore' + (num % 10000000 ? ' ' + convert(num % 10000000) : '');
    }

    return convert(num).trim() + ' Rupees';
}

function generate80GReceipt(riderName, pan, amount, transactionRef, receiptNumber) {
    Logger.log('Converting logo to base64...');
    const logoBase64 = getImageFromUrlAsBase64('https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg');
    Logger.log('Converting stamp to base64...');
    const stampBase64 = getImageAsBase64('1fQVqA1ABWCaTJs4uJVxiNqIGhl5iWugJ');
    Logger.log('Converting signature to base64...');
    const signBase64 = getImageAsBase64('1CI6H0JgysxanA0RimUwu7QwSSRospSwc');

    const htmlContent = createReceiptHTML(riderName, pan, amount, transactionRef, receiptNumber, logoBase64, stampBase64, signBase64);

    const htmlFile = DriveApp.createFile(`receipt_temp_${new Date().getTime()}.html`, htmlContent, MimeType.HTML);
    const blob = htmlFile.getAs('application/pdf');
    blob.setName(`80G_Receipt_${riderName.replace(/\s+/g, '_')}_${receiptNumber.replace(/\//g, '_')}.pdf`);

    htmlFile.setTrashed(true);

    return blob;
}

function createReceiptHTML(donorName, pan, amount, transactionRef, receiptNumber, logoBase64, stampBase64, signBase64) {
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy');
    const amountInWords = numberToWords(amount);

    const html = `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
    @page { size: A4; margin: 0; }
    body { font-family: "Times New Roman", serif; margin: 0; padding: 25px; background: #fff; }
    .receipt-container { border: 2px solid #000; border-radius: 35px; padding: 25px 30px; max-width: 800px; margin: auto; position: relative; }
    .header { display: flex; align-items: flex-start; }
    .logo-section { width: 140px; text-align: center; }
    .logo-img { width: 110px; }
    .header-center { flex: 1; text-align: center; }
    .org-name { font-size: 28px; font-weight: bold; margin-bottom: 5px; }
    .registration-info { font-size: 13px; }
    .registration-subdetails { font-size: 13px; margin-top: 3px; }
    .subtext { margin-top: 10px; font-style: italic; font-weight: bold; text-decoration: underline; }
    .receipt-number { position: absolute; right: 30px; top: 15px; font-size: 16px; font-weight: bold; color: red; }
    .receipt-box { border: 2px solid #000; border-radius: 12px; text-align: center; padding: 10px; margin: 20px 0 10px; }
    .receipt-title { font-size: 20px; font-weight: bold; }
    .receipt-subtitle { font-size: 12px; }
    .date-row { text-align: right; font-size: 14px; margin-bottom: 10px; }
    .section-title { font-weight: bold; margin: 12px 0 6px; font-size: 15px; }
    .main-content { display: flex; gap: 30px; margin-top: 10px; }
    .left-column, .right-column { flex: 1; font-size: 14px; }
    .checkbox-item { margin: 5px 0; }
    .checkbox { display: inline-block; width: 13px; height: 13px; border: 1px solid #000; margin-right: 6px; vertical-align: middle; }
    .checkbox.checked { background: #000; position: relative; }
    .checkbox.checked::after { content: "✓"; color: #fff; font-size: 11px; position: absolute; left: 1px; top: -2px; }
    .detail-row { margin: 8px 0; }
    .detail-label { font-weight: bold; }
    .amount-section { border: 2px solid #000; margin: 20px 0; padding: 18px; position: relative; text-align: center; }
    .rupee-symbol { position: absolute; left: 20px; top: 50%; transform: translateY(-50%); font-size: 40px; color: goldenrod; font-weight: bold; }
    .amount-value { font-size: 34px; font-weight: bold; }
    .payment-mode { font-size: 14px; margin-top: 10px; }
    .declaration-section { margin-top: 15px; font-size: 13px; text-align: justify; }
    .signature-section { margin-top: 36px; text-align: right; }
    .org-label { font-weight: bold; margin-bottom: 5px; }
    .stamp-and-sign { position: relative; height: 120px; }
    .sign-img { width: 110px; }
    .stamp-img { width: 120px; }
    .authorized-text { margin-top: 90px; text-decoration: underline; font-size: 14px; }
</style>
</head>
<body>
<div class="receipt-container">
    <div class="receipt-number">${receiptNumber}</div>
    <div class="header">
        <div class="logo-section">
            <img src="${logoBase64}" class="logo-img" />
        </div>
        <div class="header-center">
            <div class="org-name">Kings Equestrian Foundation</div>
            <div class="registration-info">Registered u/s 80G of Income-tax Act Rg no:AAJCK7191GE20231, 1961, PAN: AAJCK7191G</div>
            <div class="registration-subdetails">K202, Tower-6, Jacaranda Block, Devarabisanahalli, Bellandur S.O, Bengaluru – 560103 Karnataka, India<br>kingsequestrianfoundation@gmail.com, kingsequestrianfoundation.com</div>
            <div class="subtext">We gratefully acknowledge your generous contribution in support of our programmes promoting education, well-being, and personal development through sport and experiential learning.</div>
        </div>
    </div>
    <div class="receipt-box">
        <div class="receipt-title">Receipt</div>
        <div class="receipt-subtitle">This receipt is issued in compliance with Rule 18AB and Form 10BD requirements</div>
    </div>
    <div class="date-row"><strong>Date:</strong> ${currentDate}</div>
    <div class="main-content">
        <div class="left-column">
            <div class="section-title">Donor Category (✓ Tick Applicable)</div>
            <div class="checkbox-item"><span class="checkbox checked"></span> Resident Indian Donor</div>
            <div class="checkbox-item"><span class="checkbox"></span> Non-Resident Indian (NRI)</div>
        </div>
        <div class="right-column">
            <div class="section-title">Donor Details</div>
            <div class="detail-row"><span class="detail-label">Name of Donor:</span> ${donorName}</div>
            <div class="detail-row"><span class="detail-label">PAN / Aadhaar:</span> ${pan}</div>
            <div class="detail-row"><span class="detail-label">Amount in Words:</span> ${amountInWords}</div>
        </div>
    </div>
    <div class="amount-section">
        <span class="rupee-symbol">₹</span>
        <div class="amount-value">${amount.toLocaleString('en-IN')}</div>
    </div>
    <div class="payment-mode">
        <strong>Mode of Payment:</strong> Cheque / DD / NEFT / RTGS / UPI (Cash not eligible u/s 80G)<br><br>
        ${transactionRef && transactionRef !== 'N/A' ? `Transaction Reference No.: <strong>${transactionRef}</strong><br><br>` : ''}
        <strong>Amount in Words:</strong> ${amountInWords}
    </div>
    <div class="declaration-section">
        Certified that the above donation is received by trust for charitable purposes only.
        This donation is eligible for deduction under Section 80G of the Income Tax Act, 1961.
        This receipt will be reported in Form 10BD and Form 10BE will be issued to the donor.
    </div>
    <div class="signature-section">
        <div class="org-label">For Kings Equestrian Foundation</div>
        <div class="stamp-and-sign">
            <img src="${signBase64}" class="sign-img" />
            <img src="${stampBase64}" class="stamp-img" />
        </div>
    </div>
</div>
</body>
</html>`;

    return html;
}

// --------------- SEND RECEIPT FOR SPECIFIC ROW ---------------

function sendReceiptForRow(rowIndex) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);

    if (!paymentSheet || !bookingSheet) {
        Logger.log('Required sheets not found');
        return false;
    }

    const bookingValues = bookingSheet.getDataRange().getValues();

    let email = '';
    let phoneNumber = '';

    try {
        const row = paymentSheet.getRange(rowIndex, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
        phoneNumber = String(row[CONFIG.PAYMENT_COLS.PHONE_NUMBER] || '').trim();
        const regNoSubmitted = String(row[CONFIG.PAYMENT_COLS.REGISTRATION_NO] || '').trim();

        // ── Try to match a booking ─────────────────────────────────────────────
        // Regular student with KER ref → look up in Regular Students sheet for email
        // otherwise fall through to existing phone→booking lookup
        let riderName, referenceNumber, preferredDate, preferredTimeSlots, services, participants;

        const isKERRef = regNoSubmitted.toUpperCase().startsWith('KER');
        if (isKERRef) {
            const studentMatch = findRegularStudentByRef(regNoSubmitted);
            if (studentMatch) {
                riderName         = studentMatch.row[CONFIG.STUDENT_COLS.NAME];
                email             = studentMatch.row[CONFIG.STUDENT_COLS.EMAIL];
                referenceNumber   = studentMatch.row[CONFIG.STUDENT_COLS.REG_NO];
                services          = studentMatch.row[CONFIG.STUDENT_COLS.PROGRAM];
                participants      = studentMatch.row[CONFIG.STUDENT_COLS.PARTICIPANTS] || 1;
                preferredDate     = '';
                preferredTimeSlots = '';
            }
        }

        // If not resolved by KER ref, try phone → one-off booking (existing logic)
        if (!email) {
            if (!phoneNumber) throw new Error('Phone number missing in payment form');

            // Also try regular student by phone
            const regularByPhone = findRegularStudentByPhone(phoneNumber);
            if (regularByPhone) {
                riderName          = regularByPhone.row[CONFIG.STUDENT_COLS.NAME];
                email              = regularByPhone.row[CONFIG.STUDENT_COLS.EMAIL];
                referenceNumber    = regularByPhone.row[CONFIG.STUDENT_COLS.REG_NO];
                services           = regularByPhone.row[CONFIG.STUDENT_COLS.PROGRAM];
                participants       = regularByPhone.row[CONFIG.STUDENT_COLS.PARTICIPANTS] || 1;
                preferredDate      = '';
                preferredTimeSlots = '';
            }
        }

        // Still not found → fall back to one-off booking by phone (existing behaviour)
        if (!email) {
            const bookingMatch = findLatestBookingByPhone(phoneNumber, bookingValues);
            if (!bookingMatch) throw new Error(`Booking not found for phone ${phoneNumber}`);

            riderName          = bookingMatch.row[CONFIG.BOOKING_COLS.NAME];
            email              = bookingMatch.row[CONFIG.BOOKING_COLS.EMAIL_ID];
            const phone        = bookingMatch.row[CONFIG.BOOKING_COLS.PHONE_NUMBER];
            services           = bookingMatch.row[CONFIG.BOOKING_COLS.OUR_SERVICES];
            participants       = bookingMatch.row[CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS] || 1;
            preferredDate      = bookingMatch.row[CONFIG.BOOKING_COLS.PREFERRED_SERVICE_DATE];
            preferredTimeSlots = bookingMatch.row[CONFIG.BOOKING_COLS.PREFERRED_TIME_SLOT];
            referenceNumber    = bookingMatch.row[CONFIG.BOOKING_COLS.REFERENCE];
        }

        if (!email) throw new Error('Email not found');

        const amount = Number(row[CONFIG.PAYMENT_COLS.AMOUNT_PAID]);
        if (!amount || Number.isNaN(amount)) throw new Error('Valid amount is required');

        const transactionId = row[CONFIG.PAYMENT_COLS.TRANSACTION_REFERENCE_NUMBER] || '';
        const pan = row[CONFIG.PAYMENT_COLS.PAN_AADHAAR] || '';
        const transactionVerified = row[CONFIG.PAYMENT_COLS.TRANSACTION_VERIFIED];

        if (String(transactionVerified || '').toLowerCase() !== 'yes') {
            throw new Error('Transaction not verified. Please verify first.');
        }

        const receiptNumber = generateReceiptNumber(referenceNumber);
        const receiptPDF = generate80GReceipt(riderName, pan, amount, transactionId, receiptNumber);

        const driveInfo = storeReceiptInDrive(receiptPDF, riderName, receiptNumber, referenceNumber);
        if (driveInfo) {
            Logger.log(`Receipt stored in Drive: ${driveInfo.fileUrl}`);
            paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_DRIVER_LINK + 1).setValue(driveInfo.fileUrl);
        }

        const subject = `Payment Receipt - ${riderName} - Ref: ${referenceNumber}`;
        const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f4; margin: 0; padding: 0; color: #333; }
    .container { max-width: 650px; margin: 20px auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #1f4e3d 0%, #4f9c7a 100%); padding: 30px; text-align: center; color: #fff; }
    .header img { width: 70px; height: 70px; border-radius: 10%; margin-bottom: 10px; }
    .header h1 { margin: 0; font-size: 26px; }
    .content { padding: 35px 30px; }
    .greeting { font-size: 18px; font-weight: 600; margin-bottom: 15px; color: #1f4e3d; }
    .success-banner { text-align: center; margin: 25px 0; }
    .info-box { background: #f8f8f8; border-left: 4px solid #1f4e3d; padding: 18px; margin: 20px 0; border-radius: 6px; font-size: 14px; }
    .footer { background-color: #1f4e3d; color: #fff; padding: 25px; text-align: center; font-size: 13px; }
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
        <img src="https://i.pinimg.com/736x/69/3c/20/693c200ad675967032f941cf76953b3e.jpg" alt="Payment Successful" width="200" height="150" />
        <div style="font-size:18px; font-weight:600; color:#1f7a3f; margin-top:10px;">✅ Payment Confirmed - Booking Complete!</div>
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
      <p style="margin-top: 20px; font-size: 14px;">
        We look forward to welcoming you at Kings Equestrian. Please arrive 15 minutes before your scheduled time.
      </p>
      <p style="margin-top: 15px; font-size: 13px; color: #666;">
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

        const ccEmails = getCCRecipients('Receipt Mail');

        MailApp.sendEmail({
            to: email,
            cc: ccEmails.join(','),
            subject: subject,
            htmlBody: htmlBody,
            attachments: [receiptPDF],
            name: 'Kings Equestrian Foundation'
        });

        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_SENT + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');

        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.RECEIPT_SENT_TIMESTAMP + 1)
            .setValue(new Date())
            .setNumberFormat('dd-MMM-yyyy HH:mm:ss');

        paymentSheet.getRange(rowIndex, CONFIG.PAYMENT_COLS.PAYMENT_RECEIPT_NO + 1)
            .setValue(receiptNumber);

        // Calendar event — only for one-off bookings with date+time set
        if (preferredDate && preferredTimeSlots) {
            try {
                const calendarEventId = createBookingCalendarEvent({
                    name: riderName,
                    email: email,
                    phone: phoneNumber,
                    services: services,
                    date: preferredDate,
                    timeSlots: preferredTimeSlots,
                    reference: referenceNumber,
                    participants: participants
                });
                if (calendarEventId) {
                    Logger.log(`Calendar event created: ${calendarEventId} for ${referenceNumber}`);
                }
            } catch (calError) {
                Logger.log(`Warning: Calendar event creation failed for ${referenceNumber}: ${calError.message}`);
            }
        }

        Logger.log(`Receipt sent to: ${email} for ref ${referenceNumber} with CC to: ${ccEmails.join(', ')}`);
        return true;

    } catch (error) {
        Logger.log(`Receipt failed for row ${rowIndex} (Phone: ${phoneNumber || 'N/A'}, Email: ${email || 'N/A'}): ${error.message}`);
        throw error;
    }
}

// --------------- PAYMENT RECEIPT MENU FUNCTION ---------------

function SendPaymentReceipt() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    const numRows = selection.getNumRows();

    const response = ui.alert('Send Payment Receipts', `Send receipts for ${numRows} row(s)?`, ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    let successCount = 0;
    let failCount = 0;
    const errors = [];

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

// --------------- RESEND WELCOME EMAIL MENU FUNCTION ---------------

function ResendWelcomeEmail() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);

    if (!bookingSheet) {
        ui.alert('❌ Booking Form Response sheet not found');
        return;
    }

    const selection = bookingSheet.getActiveRange();
    if (!selection) {
        ui.alert('Please select rows to resend welcome emails');
        return;
    }

    const startRow = selection.getRow();
    const numRows = selection.getNumRows();

    if (startRow === 1) {
        ui.alert('Cannot send emails for header row');
        return;
    }

    const response = ui.alert('Resend Welcome Emails', `Resend welcome emails for ${numRows} row(s)?`, ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < numRows; i++) {
        const rowIndex = startRow + i;
        try {
            const name = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.NAME + 1).getValue();
            const email = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.EMAIL_ID + 1).getValue();
            const phone = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.PHONE_NUMBER + 1).getValue();
            const services = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.OUR_SERVICES + 1).getValue();
            const reference = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.REFERENCE + 1).getValue();
            const participants = Number(bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS + 1).getValue()) || 1;
            const bookingDate = bookingSheet.getRange(rowIndex, CONFIG.BOOKING_COLS.TIMESTAMP + 1).getValue();

            if (!email || !reference) throw new Error('Missing email or reference');

            const amount = CONFIG.ADVANCE_BOOKING_AMOUNT;
            const upiLink = createUPILink(amount, reference);
            const qrCode = createQRCode(upiLink);

            sendWelcomeEmail({
                name, email, phone, services, participants,
                amount, reference, upiLink, qrCode,
                row: rowIndex, sheet: bookingSheet, bookingDate
            });

            successCount++;
            Utilities.sleep(1000);
        } catch (error) {
            failCount++;
            Logger.log(`❌ Error at row ${rowIndex}: ${error.message}`);
        }
    }

    ui.alert(`Complete!\n✅ Sent: ${successCount}\n❌ Failed: ${failCount}`);
}

// --------------- HELPER FUNCTIONS ---------------

function formatDate(date) {
    if (!date) return 'N/A';
    if (typeof date === 'string') return date;
    try {
        return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd MMM yyyy');
    } catch (e) {
        return String(date);
    }
}

// --------------- GOOGLE CALENDAR INTEGRATION ---------------

function createCalendarEvent(bookingData) {
    try {
        const calendar = CalendarApp.getDefaultCalendar();
        const date = new Date(bookingData.date);
        const timeSlots = String(bookingData.timeSlots).split(',');
        const firstSlot = timeSlots[0].trim();
        const timeParts = firstSlot.match(/(\d+):(\d+)\s*(AM|PM)/i);

        if (!timeParts) {
            Logger.log('Invalid time format: ' + firstSlot);
            return null;
        }

        let hours = parseInt(timeParts[1]);
        const minutes = parseInt(timeParts[2]);
        const period = timeParts[3].toUpperCase();

        if (period === 'PM' && hours !== 12) hours += 12;
        if (period === 'AM' && hours === 12) hours = 0;

        const startTime = new Date(date);
        startTime.setHours(hours, minutes, 0);

        const endTime = new Date(startTime);
        endTime.setMinutes(endTime.getMinutes() + (timeSlots.length * 30));

        const participants = bookingData.participants || 1;
        const participantText = participants > 1 ? ` (${participants} participants)` : '';

        const event = calendar.createEvent(
            `Kings Equestrian - ${bookingData.name}${participantText} (${bookingData.reference})`,
            startTime,
            endTime, {
                description: `Service: ${bookingData.services}\nParticipants: ${participants}\nReference: ${bookingData.reference}\nPhone: ${bookingData.phone}\nEmail: ${bookingData.email}`,
                location: 'Kings Equestrian Foundation, Karnataka',
                guests: bookingData.email,
                sendInvites: true
            }
        );

        Logger.log('Calendar event created: ' + event.getId());
        return event.getId();
    } catch (error) {
        Logger.log('Error creating calendar event: ' + error);
        return null;
    }
}

// --------------- CONSENT PDF GENERATION ---------------

function generateConsentPDF(name, email, phone, bookingDate) {

    const LABEL_FONT = 'Arial';
    const FONT_SIZE = 11;

    const doc = DocumentApp.create('Consent Form - ' + (name || 'Participant'));
    const body = doc.getBody();
    body.clear();

    body.setMarginTop(40);
    body.setMarginBottom(40);
    body.setMarginLeft(50);
    body.setMarginRight(50);

    function paragraph(textStr, size = FONT_SIZE, bold = false, spacing = 6, align = null) {
        const p = body.appendParagraph(textStr);
        const t = p.editAsText();
        t.setFontFamily(LABEL_FONT).setFontSize(size).setBold(bold);
        if (align) p.setAlignment(align);
        p.setSpacingAfter(spacing);
        return p;
    }

    function formatValue(textObj, fullText, value) {
        if (!value || value.toString().trim().length === 0) return;
        if (!fullText) return;
        const valStr = value.toString();
        const start = fullText.indexOf(valStr);
        if (start === -1) return;
        const end = start + valStr.length - 1;
        if (end >= start && start >= 0 && end < fullText.length) {
            textObj.setBold(start, end, true);
            textObj.setUnderline(start, end, true);
        }
    }

    function formatDateOnly(dateValue) {
        if (!dateValue) return null;
        if (typeof dateValue === 'string') {
            try {
                const parsed = new Date(dateValue);
                if (!isNaN(parsed.getTime())) { dateValue = parsed; } else { return dateValue; }
            } catch (e) { return dateValue; }
        }
        if (dateValue instanceof Date) {
            const day = String(dateValue.getDate()).padStart(2, '0');
            const month = String(dateValue.getMonth() + 1).padStart(2, '0');
            const year = dateValue.getFullYear();
            return `${day}/${month}/${year}`;
        }
        return dateValue.toString();
    }

    const logoUrl = 'https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg';
    let logoBlob;
    try {
        logoBlob = UrlFetchApp.fetch(logoUrl).getBlob();
    } catch (e) {
        Logger.log('Error fetching logo: ' + e);
    }

    if (logoBlob) {
        const logoPara = body.appendParagraph('');
        logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        const logoImg = logoPara.appendInlineImage(logoBlob);
        logoImg.setWidth(120);
        logoImg.setHeight(120);
        logoPara.setSpacingAfter(20);
    }

    paragraph('KINGS EQUESTRIAN FOUNDATION', 16, true, 5, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('Acknowledgement & Consent Form – Horse Riding Participants', 13, true, 3, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('(Applicable for Individual / Group / Family Participants)', 10, false, 25, DocumentApp.HorizontalAlignment.CENTER);

    paragraph('Kings Equestrian Foundation offers horse riding programs and related activities, which may include casual riding, dressage, jumping, workshops, clinics, and equine interaction.', 11, false, 12);
    paragraph('I/we understand and acknowledge that participation in equestrian activities involves inherent risks, including but not limited to falls, bruises, muscle strain, fractures, head injuries, or other serious injuries. I/we further acknowledge that horses are live animals and their behaviour can be unpredictable.', 11, false, 12);
    paragraph('I/we also acknowledge that Kings Equestrian Foundation follows reasonable safety precautions, provides trained supervision, and enforces established safety guidelines. However, despite all precautions, accidents may occasionally occur.', 11, false, 20);

    let sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    paragraph('Medical Fitness & Insurance Declaration', 12, true, 12);
    paragraph('I/we hereby declare that I / my child / all participants covered under this consent are medically fit to participate in horse riding and equestrian-related activities. To the best of my/our knowledge, there are no undisclosed medical conditions, injuries, or health concerns that would prevent safe participation, except those disclosed in writing to Kings Equestrian Foundation prior to participation.', 11, false, 12);
    paragraph('I/we further confirm that I / my child / all participants are covered by valid medical and/or personal accident insurance, which will cover any injuries, medical treatment, or emergencies arising from participation.', 11, false, 12);
    paragraph('I/we understand and agree that Kings Equestrian Foundation is not responsible for medical expenses, and all such costs shall be borne by the participant(s) or covered under their insurance.', 11, false, 20);

    sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    paragraph('Acknowledgement & Agreement', 12, true, 12);
    paragraph('I/we confirm that:', 11, false, 8);

    const bulletPoints = [
        'I/we have carefully read and fully understood this consent form.',
        'I/we understand the nature of equestrian activities and the associated risks.',
        'I/we voluntarily consent to participation.',
        'For participants under 18 years of age, I/we am/are the parent(s) or legal guardian(s) and authorised to provide consent.',
        'All participants agree to follow safety instructions, rules, and guidelines issued by Kings Equestrian Foundation and its instructors at all times.'
    ];

    bulletPoints.forEach(point => {
        const p = body.appendParagraph('• ' + point);
        p.editAsText().setFontFamily(LABEL_FONT).setFontSize(11);
        p.setSpacingAfter(6);
        p.setIndentStart(20);
        p.setIndentFirstLine(0);
    });

    body.appendParagraph('').setSpacingAfter(8);
    paragraph('I/we agree that Kings Equestrian Foundation, its trainers, staff, and associates shall not be held responsible for injuries arising from participation, except in cases of proven negligence.', 11, false, 20);

    sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    paragraph('Primary Contact / Parent / Guardian Details', 12, true, 12);

    let p = body.appendParagraph('');
    let t = p.editAsText();
    const nameSpaced = name ? `  ${name}  ` : '___________________________________';
    const nameLine = `Name: ${nameSpaced}`;
    t.setText(nameLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (name) formatValue(t, nameLine, nameSpaced);
    p.setSpacingAfter(12);

    p = body.appendParagraph('');
    t = p.editAsText();
    const phoneSpaced = phone ? `  ${phone}  ` : '___________________________________';
    const phoneLine = `Contact Number: ${phoneSpaced}`;
    t.setText(phoneLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (phone) formatValue(t, phoneLine, phoneSpaced);
    p.setSpacingAfter(12);

    p = body.appendParagraph('');
    t = p.editAsText();
    const emailSpaced = email ? `  ${email}  ` : '___________________________________';
    const emailLine = `Email ID: ${emailSpaced}`;
    t.setText(emailLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (email) formatValue(t, emailLine, emailSpaced);
    p.setSpacingAfter(25);

    sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(25);

    p = body.appendParagraph('');
    t = p.editAsText();
    const signatureSpaced = name ? `  ${name}  ` : '___________________________________';
    const dateFormatted = formatDateOnly(bookingDate);
    const dateSpaced = dateFormatted ? `  ${dateFormatted}  ` : '_______________';
    const signatureLine = `Signature of Participant / Parent / Guardian: ${signatureSpaced}     Date: ${dateSpaced}`;
    t.setText(signatureLine).setFontFamily(LABEL_FONT).setFontSize(11);

    if (name) {
        const sigStart = signatureLine.indexOf(signatureSpaced);
        if (sigStart !== -1) {
            const sigEnd = sigStart + signatureSpaced.length - 1;
            if (sigEnd >= sigStart && sigStart >= 0) {
                t.setFontFamily(sigStart, sigEnd, 'Dancing Script');
                t.setFontSize(sigStart, sigEnd, 16);
                t.setBold(sigStart, sigEnd, false);
                t.setUnderline(sigStart, sigEnd, false);
            }
        }
    }

    if (dateFormatted) formatValue(t, signatureLine, dateSpaced);
    p.setSpacingAfter(30);

    const footerPara = paragraph('Kings Equestrian Foundation | Karnataka, India | +91-9980895533 | info@kingsequestrian.com', 9, false, 0, DocumentApp.HorizontalAlignment.CENTER);
    footerPara.editAsText().setForegroundColor('#666666');

    doc.saveAndClose();

    const pdf = doc.getAs('application/pdf');
    pdf.setName(`Consent_Form_${(name || 'Participant').replace(/\s+/g, '_')}.pdf`);

    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return pdf;
}

// --------------- DRIVE STORAGE ---------------

function getKingsFarmFolder() {
    const folderName = "Kings Farm Receipts";
    const year = new Date().getFullYear();

    let mainFolder = DriveApp.getFoldersByName(folderName);
    if (!mainFolder.hasNext()) {
        mainFolder = DriveApp.createFolder(folderName);
    } else {
        mainFolder = mainFolder.next();
    }

    const yearFolders = mainFolder.getFoldersByName(year.toString());
    if (yearFolders.hasNext()) {
        return yearFolders.next();
    } else {
        return mainFolder.createFolder(year.toString());
    }
}

function storeReceiptInDrive(receiptBlob, riderName, receiptNumber, referenceNumber) {
    try {
        const folder = getKingsFarmFolder();
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
        const fileName = `Receipt_${receiptNumber.replace(/\//g, '-')}_${riderName.replace(/\s+/g, '_')}_${timestamp}.pdf`;

        const file = folder.createFile(receiptBlob);
        file.setName(fileName);
        file.setDescription(`Receipt for ${riderName} | Reference: ${referenceNumber} | Receipt No: ${receiptNumber}`);

        Logger.log(`Receipt saved to Drive: ${fileName}`);

        return {
            fileId: file.getId(),
            fileUrl: file.getUrl(),
            fileName: fileName,
            folderId: folder.getId(),
            folderUrl: folder.getUrl()
        };
    } catch (error) {
        Logger.log('Error storing receipt in Drive: ' + error);
        return null;
    }
}

// --------------- MENU SETUP ---------------

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🎠 Kings Equestrian')
        .addItem('📧 Resend Welcome Email', 'ResendWelcomeEmail')
        .addItem('🧾 Send Payment Receipt', 'SendPaymentReceipt')
        .addSeparator()
        .addItem('⚙️ Setup Triggers', 'setupTriggers')
        .addSeparator()
        .addItem('📅 Send Daily Summary Now', 'testSendDailySummaryNow')
        .addItem('🧪 Test Summary (Dry Run)', 'testDailySummaryDryRun')
        .addItem('⚙️ Setup New Feature Triggers', 'setupNewFeaturesTriggers')
        .addToUi();
}

function setupTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    ScriptApp.newTrigger('onBookingFormSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();

    // ── NEW: trigger for the regular riders booking form ──
    ScriptApp.newTrigger('onRegularBookingFormSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();

    ScriptApp.newTrigger('onPaymentFormSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();

    SpreadsheetApp.getUi().alert('✅ Triggers set up successfully!\n\n' +
        'The system will now automatically:\n' +
        '- Generate reference numbers and send welcome emails on booking\n' +
        '- Generate KER reg numbers and send registration emails for regular riders\n' +
        '- Auto-send receipts when payment form is submitted\n' +
        '- Resend existing receipts for duplicate submissions\n' +
        '- Store receipts in Google Drive');
}