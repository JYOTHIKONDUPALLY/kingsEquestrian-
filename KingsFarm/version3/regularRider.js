// ============================================
// KINGS EQUESTRIAN - REGULAR BOOKING FORM HANDLER
// File: 5_RegularBooking.js
// Depends on: 1_Config.js, 4_Receipt.js
// ============================================
//
// FLOW FOR REGULAR RIDERS:
//
//  [1] Student submits the Regular Booking Google Form
//        → onRegularBookingFormSubmit() fires
//        → Generates KER ref
//        → Creates row in Regular Students sheet (Payment Status: Pending)
//        → Sends Registration Confirmation Email:
//            - Their KER reg number
//            - Program details + pricing
//            - QR code to pay (full or partial)
//            - Link to payment form
//            - Link to rider portal
//
//  [2] Student pays via the Payment Form (existing form)
//        → They enter their KER ref OR phone number
//        → onPaymentFormSubmit() in 3_Payment.js handles it
//        → Detects KER ref → routes to regular student path
//        → Sends 80G receipt for the amount paid
//        → Updates Amount Paid + Payment Status in Regular Students sheet
//        → If balance is now zero → marks Paid
//        → If partial → marks Partial
//
//  [3] Rider Portal (web app)
//        → Student logs in with phone or KER ref
//        → Sees their dashboard, books class slots, reschedules
//
// ═════════════════════════════════════════════════════════════════════════════════

// ────────────────────────────────────────────────────────────────────────
//  FORM SUBMIT TRIGGER
// ────────────────────────────────────────────────────────────────────────

/**
 * Triggered when a student submits the Regular Booking Google Form.
 * Sheet name: "Regular Booking Form Response"
 * Columns (0-indexed):
 *   0: Timestamp
 *   1: Name
 *   2: Email ID
 *   3: Phone Number
 *   4: Our Services  (maps to a program name in Regular Programs sheet)
 *   5: Number of Participants
 *   6: Consent
 *   7: (written by script) Reg No
 *   8: (written by script) Reg Email Sent
 *   9: (written by script) Reg Email Timestamp
 */
function onRegularBookingFormSubmit(e) {
    try {
        const sheet = e.range.getSheet();
        if (sheet.getName() !== CONFIG.SHEETS.REGULAR_BOOKING_FORM) {
            Logger.log('onRegularBookingFormSubmit: Skipping — wrong sheet: ' + sheet.getName());
            return;
        }

        const row = e.range.getRow();
        Logger.log('Regular booking form submitted at row ' + row);

        // ── Read form data ──
        const C    = CONFIG.REG_BOOKING_COLS;
        const name         = String(sheet.getRange(row, C.NAME + 1).getValue() || '').trim();
        const email        = String(sheet.getRange(row, C.EMAIL_ID + 1).getValue() || '').trim();
        const phone        = String(sheet.getRange(row, C.PHONE_NUMBER + 1).getValue() || '').trim();
        const serviceRaw   = String(sheet.getRange(row, C.OUR_SERVICES + 1).getValue() || '').trim();
        const participants = Number(sheet.getRange(row, C.NUMBER_OF_PARTICIPANTS + 1).getValue()) || 1;
        const timestamp    = sheet.getRange(row, C.TIMESTAMP + 1).getValue();

        if (!name || !email || !phone) {
            Logger.log('Regular booking: Missing required fields at row ' + row);
            sheet.getRange(row, C.REG_NO + 1).setValue('⚠️ Missing fields');
            return;
        }

        // ── Resolve program from service selection ──
        // The form's "Our Services" field value might be the full program name
        // or a partial match. We do a best-effort lookup.
        const programs  = getRegularProgramsData();
        const program   = resolveProgram(serviceRaw, programs);
        const progData  = program ? programs[program] : null;

        // ── Generate KER registration number ──
        const regNo = generateRegularRef();
        Logger.log(`Generated KER ref: ${regNo} for ${name}`);

        // ── Create Regular Students record ──
        createRegularStudentRecord({
            regNo, name, email, phone, program: program || serviceRaw,
            participants, totalAmount: progData ? (progData.totalPrice || progData.pricePerMonth) : 0,
            enrolledOn: timestamp
        });

        // ── Write reg no back into the booking form sheet ──
        sheet.getRange(row, C.REG_NO + 1).setValue(regNo);

        // ── Send registration confirmation email ──
        sendRegularRegistrationEmail({
            name, email, phone, regNo,
            program: program || serviceRaw,
            progData, participants
        });

        // ── Mark email sent ──
        sheet.getRange(row, C.REG_EMAIL_SENT + 1)
            .setValue('Yes')
            .setBackground('#d4edda')
            .setFontColor('#155724')
            .setFontWeight('bold');
        sheet.getRange(row, C.REG_EMAIL_TIMESTAMP + 1)
            .setValue(new Date())
            .setNumberFormat('dd-MMM-yyyy HH:mm:ss');

        Logger.log(`Regular booking processed: ${name} — ${regNo} — ${program || serviceRaw}`);

    } catch (err) {
        Logger.log('Error in onRegularBookingFormSubmit: ' + err);
        Logger.log('Stack: ' + err.stack);
    }
}

// ───────────────────────────────────────────────────────────
//  PROGRAM RESOLVER
// ───────────────────────────────────────────────────────────

/**
 * Tries to match the raw form value to a program name in the programs map.
 * Tries: exact → case-insensitive exact → partial match.
 */
function resolveProgram(raw, programs) {
    if (!raw) return null;
    // Exact match
    if (programs[raw]) return raw;
    // Case-insensitive exact
    const lc = raw.toLowerCase();
    for (const key of Object.keys(programs)) {
        if (key.toLowerCase() === lc) return key;
    }
    // Partial match (program name contains the raw value or vice versa)
    for (const key of Object.keys(programs)) {
        if (key.toLowerCase().includes(lc) || lc.includes(key.toLowerCase())) return key;
    }
    Logger.log(`resolveProgram: No match for "${raw}" — will store as-is`);
    return null;
}

// ───────────────────────────────────────────────────────────
//  CREATE REGULAR STUDENT RECORD
// ───────────────────────────────────────────────────────────

/**
 * Writes a new row into the Regular Students sheet.
 * Payment Status starts as "Pending", Amount Paid = 0.
 */
function createRegularStudentRecord(data) {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.REGULAR_STUDENTS);

    if (!sheet) {
        // Auto-create if missing
        setupRegularSheets();
        return createRegularStudentRecord(data); // retry once
    }

    // Check for duplicate (same phone already enrolled in same program)
    const existing = findRegularStudentByPhoneOrRef(data.phone);
    if (existing) {
        const existingProgram = String(existing.row[CONFIG.STUDENT_COLS.PROGRAM] || '');
        if (existingProgram.toLowerCase() === (data.program || '').toLowerCase()) {
            Logger.log(`Duplicate regular booking detected: phone ${data.phone} already enrolled in ${data.program}`);
            // Update the reg no in the booking form but do NOT create a duplicate student record
            return existing.row[CONFIG.STUDENT_COLS.REG_NO];
        }
    }

    const C = CONFIG.STUDENT_COLS;
    const programs = (() => { try { return getRegularProgramsData(); } catch(e) { return {}; } })();
    const prog = programs[data.program] || {};

    const rowData = new Array(15).fill('');
    rowData[C.REG_NO]            = data.regNo;
    rowData[C.NAME]              = data.name;
    rowData[C.EMAIL]             = data.email;
    rowData[C.PHONE]             = data.phone;
    rowData[C.PROGRAM]           = data.program || '';
    rowData[C.START_DATE]        = '';          // filled when first class is booked
    rowData[C.END_DATE]          = '';
    rowData[C.TOTAL_CLASSES]     = prog.totalClasses || '';
    rowData[C.CLASSES_COMPLETED] = 0;
    rowData[C.PAYMENT_STATUS]    = 'Pending';
    rowData[C.AMOUNT_PAID]       = 0;
    rowData[C.TOTAL_AMOUNT]      = data.totalAmount || prog.totalPrice || prog.pricePerMonth || '';
    rowData[C.PAYMENT_REF]       = '';
    rowData[C.ENROLLED_ON]       = data.enrolledOn || new Date();
    rowData[C.STATUS]            = 'Active';

    sheet.appendRow(rowData);

    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, C.ENROLLED_ON + 1).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
    sheet.getRange(newRow, C.PAYMENT_STATUS + 1)
        .setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');

    Logger.log(`Regular student record created: ${data.regNo} — ${data.name}`);
    return data.regNo;
}

// ───────────────────────────────────────────────────────────
//  REGISTRATION CONFIRMATION EMAIL
// ───────────────────────────────────────────────────────────

/**
 * Sends the registration confirmation email to a new regular rider.
 * Includes:
 *  - Their KER registration number (prominent)
 *  - Program summary (classes, price)
 *  - UPI QR code to make payment
 *  - Link to payment form
 *  - Link to rider portal
 */
function sendRegularRegistrationEmail(data) {
    const { name, email, phone, regNo, program, progData, participants } = data;

    const totalAmount   = progData ? (progData.totalPrice || progData.pricePerMonth || 0) : 0;
    const totalClasses  = progData ? progData.totalClasses : '—';
    const upiLink       = totalAmount > 0 ? createUPILink(totalAmount, regNo) : '';
    const qrCode        = upiLink ? createQRCode(upiLink) : '';

    // Program detail rows for the email table
    const progRows = [
        progData && progData.totalClasses    ? `<tr><td class="dl">Total Classes</td><td><strong>${progData.totalClasses}</strong></td></tr>` : '',
        progData && progData.classesPerWeek  ? `<tr><td class="dl">Classes per Week</td><td>${progData.classesPerWeek}</td></tr>` : '',
        progData && progData.preferredDays   ? `<tr><td class="dl">Preferred Days</td><td>${progData.preferredDays.join(', ')}</td></tr>` : '',
        progData && progData.defaultTime     ? `<tr><td class="dl">Default Time</td><td>${progData.defaultTime}</td></tr>` : '',
        progData && progData.pricePerMonth   ? `<tr><td class="dl">Price per Month</td><td>₹${progData.pricePerMonth.toLocaleString('en-IN')}</td></tr>` : '',
        totalAmount > 0                      ? `<tr><td class="dl">Total Program Fee</td><td><strong style="color:#0F6E56;font-size:18px;">₹${totalAmount.toLocaleString('en-IN')}</strong></td></tr>` : '',
    ].filter(Boolean).join('');

    const paymentBlock = totalAmount > 0 ? `
    <div style="background:#e8f5e9;border:2px solid #4caf50;padding:24px;border-radius:10px;margin:24px 0;">
        <h3 style="margin:0 0 8px;color:#1b5e20;font-size:17px;">💳 Complete Your Payment</h3>
        <p style="color:#2e7d32;font-size:13px;margin:0 0 16px;">Pay the full program fee (or a partial amount) to activate your account and start booking classes.</p>

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
               style="display:inline-block;background:#0F6E56;color:white;padding:13px 28px;text-decoration:none;border-radius:7px;font-weight:600;font-size:14px;letter-spacing:0.03em;">
                📝 Submit Payment Confirmation
            </a>
        </div>

        <div style="background:#fff8e1;border-left:4px solid #ffc107;padding:12px 14px;border-radius:0 6px 6px 0;font-size:12px;color:#5d4037;line-height:1.6;">
            <strong>⚠️ Important:</strong> After paying via UPI, submit the payment form above with your
            <strong>Registration No. ${regNo}</strong> and transaction screenshot.
            Your 80G receipt will be emailed automatically. Partial payments are also accepted — you can pay in installments.
        </div>
    </div>` : `
    <div style="background:#fff3cd;border-left:4px solid #ffc107;padding:14px;border-radius:6px;margin:20px 0;">
        <p style="margin:0;font-size:13px;color:#856404;">
            <strong>Payment details</strong> will be shared with you shortly. Please contact us at
            +91 99807 71166 for pricing information.
        </p>
    </div>`;

    const subject = `🐴 Registration Confirmed — Kings Equestrian | Reg No: ${regNo}`;

    const htmlBody = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{margin:0;padding:0;background:#f0f4f1;font-family:'Segoe UI',Arial,sans-serif;color:#1a2e1e}
  .wrap{max-width:620px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(4,52,44,0.10)}
  .hdr{background:linear-gradient(135deg,#04342C 0%,#0F6E56 60%,#1D9E75 100%);padding:32px 28px 24px;text-align:center}
  .hdr img{width:80px;height:80px;border-radius:50%;border:3px solid rgba(93,202,165,0.4);margin-bottom:14px}
  .hdr h1{margin:0;color:#C8EFE3;font-size:26px;font-weight:600;letter-spacing:-0.01em}
  .hdr p{margin:8px 0 0;color:#9FE1CB;font-size:13px;opacity:0.9}
  .body{padding:28px}
  .reg-box{background:#f0f9f5;border:2px solid #9FE1CB;border-radius:10px;padding:20px;text-align:center;margin:0 0 24px}
  .reg-label{font-size:11px;letter-spacing:0.1em;text-transform:uppercase;color:#5a7a61;margin-bottom:6px}
  .reg-no{font-family:Georgia,serif;font-size:34px;font-weight:700;color:#04342C;letter-spacing:0.04em;margin:0 0 6px}
  .reg-hint{font-size:11px;color:#7a9a7e;line-height:1.5}
  h2{color:#04342C;font-size:18px;font-weight:600;margin:24px 0 10px;border-bottom:2px solid #e8f5e9;padding-bottom:8px}
  table.prog{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px}
  table.prog tr{border-bottom:1px solid #f0f4f1}
  table.prog td{padding:8px 6px}
  td.dl{color:#7a9a7e;width:45%}
  .portal-box{background:#04342C;border-radius:10px;padding:20px 22px;margin:24px 0;text-align:center}
  .portal-box p{color:#9FE1CB;font-size:13px;margin:0 0 14px;line-height:1.5}
  .portal-btn{display:inline-block;background:#1D9E75;color:#fff;padding:12px 26px;text-decoration:none;border-radius:7px;font-weight:600;font-size:14px}
  .next-steps{background:#f7faf8;border-radius:8px;padding:16px 18px;margin:20px 0}
  .next-steps ol{margin:8px 0 0;padding-left:20px;font-size:13px;line-height:2;color:#3a5a3e}
  .ftr{background:#f0f4f1;padding:16px;text-align:center;font-size:11px;color:#7a9a7e;line-height:1.7;border-top:1px solid #dce8de}
</style>
</head>
<body>
<div class="wrap">

  <div class="hdr">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian">
    <h1>Registration Confirmed!</h1>
    <p>Kings Equestrian Foundation · Regular Rider Program</p>
  </div>

  <div class="body">
    <p style="font-size:15px;">Dear <strong>${name}</strong>,</p>
    <p style="font-size:13px;color:#5a7a61;line-height:1.6;">
      Welcome to Kings Equestrian Foundation! Your registration has been received and your
      unique Registration Number has been generated. Please save this number — you will use it
      for all future interactions including payments, class bookings, and the rider portal.
    </p>

    <!-- Registration Number — the most important element -->
    <div class="reg-box">
      <div class="reg-label">Your Registration Number</div>
      <div class="reg-no">${regNo}</div>
      <div class="reg-hint">
        Use this number when submitting payments, booking classes, or logging into the portal.<br>
        Your phone number <strong>${phone}</strong> also works everywhere.
      </div>
    </div>

    <!-- Program summary -->
    <h2>📋 Program Details</h2>
    <table class="prog">
      <tr><td class="dl">Program</td><td><strong>${program}</strong></td></tr>
      <tr><td class="dl">Participants</td><td>${participants}</td></tr>
      ${progRows}
    </table>

    <!-- Payment section -->
    ${paymentBlock}

    <!-- Portal link -->
    <div class="portal-box">
      <p>Once registered and payment is made, log in to your <strong style="color:#C8EFE3;">Rider Portal</strong>
      to book your class slots, view your schedule, and reschedule classes.</p>
      <a href="${CONFIG.RIDER_PORTAL_URL}?reg=${encodeURIComponent(regNo)}" class="portal-btn">
        🐴 Open Rider Portal
      </a>
    </div>

    <!-- Next steps -->
    <div class="next-steps">
      <strong style="color:#04342C;font-size:13px;">📌 What Happens Next?</strong>
      <ol>
        <li>Pay via UPI using the QR code above</li>
        <li>Submit the <a href="${CONFIG.PAYMENT_FORM_LINK}" style="color:#0F6E56;">payment confirmation form</a> with your Reg No: <strong>${regNo}</strong></li>
        <li>Receive your 80G tax receipt instantly by email</li>
        <li>Log into the <a href="${CONFIG.RIDER_PORTAL_URL}" style="color:#0F6E56;">Rider Portal</a> to book your class slots</li>
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
    <strong style="color:#5a7a61;">Kings Equestrian Foundation</strong><br>
    Just 30 mins from Bengaluru · Karnataka, India<br>
    📞 +91 99807 71166 · 📸 @kingsequestrianfoundation<br>
    <a href="${CONFIG.RIDER_PORTAL_URL}" style="color:#0F6E56;">Rider Portal</a> ·
    <a href="${CONFIG.PAYMENT_FORM_LINK}" style="color:#0F6E56;">Payment Form</a>
  </div>
</div>
</body>
</html>`;

    const ccEmails = getCCRecipients('Regular Registration');

    MailApp.sendEmail({
        to: email,
        cc: ccEmails.join(','),
        subject,
        htmlBody,
        name: 'Kings Equestrian Foundation'
    });

    Logger.log(`Registration email sent: ${name} (${regNo}) → ${email}`);
}

// ───────────────────────────────────────────────────────────
//  PAYMENT HANDLING FOR REGULAR STUDENTS
// ───────────────────────────────────────────────────────────

/**
 * Called from 3_Payment.js after a receipt is issued for a KER ref.
 * Updates the Regular Students sheet:
 *   - Adds the amount to Amount Paid
 *   - Updates Payment Status: Pending → Partial → Paid
 *   - Records the payment ref
 *
 * @param {string} regNo        KER registration number
 * @param {number} amountPaid   Amount in this payment transaction
 */
function updateRegularStudentPayment(regNo, amountPaid) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.REGULAR_STUDENTS);
    if (!sheet) { Logger.log('Regular Students sheet not found'); return; }

    const data = sheet.getDataRange().getValues();
    const C    = CONFIG.STUDENT_COLS;

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][C.REG_NO] || '').trim() !== String(regNo).trim()) continue;

        const previouslyPaid = Number(data[i][C.AMOUNT_PAID]) || 0;
        const totalAmount    = Number(data[i][C.TOTAL_AMOUNT]) || 0;
        const newAmountPaid  = previouslyPaid + amountPaid;
        const balance        = totalAmount > 0 ? totalAmount - newAmountPaid : -1;

        // Determine new payment status
        let newStatus;
        if (totalAmount <= 0) {
            newStatus = 'Paid'; // no total defined, treat as paid
        } else if (newAmountPaid >= totalAmount) {
            newStatus = 'Paid';
        } else if (newAmountPaid > 0) {
            newStatus = 'Partial';
        } else {
            newStatus = 'Pending';
        }

        // Set colours based on status
        const statusColors = {
            'Paid':    { bg: '#d4edda', fg: '#155724' },
            'Partial': { bg: '#fff3cd', fg: '#856404' },
            'Pending': { bg: '#f8d7da', fg: '#721c24' }
        };
        const colors = statusColors[newStatus];

        const sheetRow = i + 1;
        sheet.getRange(sheetRow, C.AMOUNT_PAID + 1).setValue(newAmountPaid);
        sheet.getRange(sheetRow, C.PAYMENT_REF + 1).setValue(regNo);
        sheet.getRange(sheetRow, C.PAYMENT_STATUS + 1)
            .setValue(newStatus)
            .setBackground(colors.bg)
            .setFontColor(colors.fg)
            .setFontWeight('bold');

        Logger.log(`Payment updated for ${regNo}: previously ₹${previouslyPaid} + ₹${amountPaid} = ₹${newAmountPaid} / ₹${totalAmount} (${newStatus})`);

        // If fully paid, send a "payment complete" confirmation email
        if (newStatus === 'Paid' && previouslyPaid < totalAmount) {
            const studentName  = data[i][C.NAME];
            const studentEmail = data[i][C.EMAIL];
            sendPaymentCompleteEmail(studentName, studentEmail, regNo,
                data[i][C.PROGRAM], newAmountPaid, totalAmount);
        }

        return;
    }

    Logger.log(`updateRegularStudentPayment: student ${regNo} not found in Regular Students sheet`);
}

/**
 * Sent when a regular student's balance reaches zero (fully paid).
 * Encourages them to book classes via the portal.
 */
function sendPaymentCompleteEmail(name, email, regNo, program, amountPaid, totalAmount) {
    const subject = `✅ Full Payment Received — Start Booking Your Classes! (${regNo})`;

    const htmlBody = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{margin:0;padding:0;background:#f0f4f1;font-family:'Segoe UI',Arial,sans-serif}
  .wrap{max-width:580px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(4,52,44,0.10)}
  .hdr{background:linear-gradient(135deg,#04342C,#1D9E75);padding:28px;text-align:center}
  .hdr img{width:70px;height:70px;border-radius:50%;border:2px solid rgba(93,202,165,0.4);margin-bottom:12px}
  .hdr h1{margin:0;color:#C8EFE3;font-size:22px;font-weight:600}
  .body{padding:26px}
  .banner{background:#e8f5e9;border:2px solid #4caf50;border-radius:10px;padding:18px;text-align:center;margin:0 0 22px}
  .banner h2{margin:0 0 6px;color:#1b5e20;font-size:20px}
  .banner p{margin:0;font-size:13px;color:#2e7d32}
  .detail-row{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #f0f4f1;font-size:13px}
  .portal-box{background:#04342C;border-radius:10px;padding:20px;text-align:center;margin:22px 0}
  .portal-box p{color:#9FE1CB;font-size:13px;margin:0 0 14px}
  .portal-btn{display:inline-block;background:#1D9E75;color:#fff;padding:12px 26px;text-decoration:none;border-radius:7px;font-weight:600;font-size:14px}
  .ftr{background:#f0f4f1;padding:14px;text-align:center;font-size:11px;color:#7a9a7e;border-top:1px solid #dce8de}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian">
    <h1>Payment Complete!</h1>
  </div>
  <div class="body">
    <p>Dear <strong>${name}</strong>,</p>
    <p style="font-size:13px;color:#5a7a61;">Your full payment has been received. You are all set to start booking your classes!</p>

    <div class="banner">
      <h2>🎉 You're Fully Paid!</h2>
      <p>Your ${program} account is now fully active.</p>
    </div>

    <div class="detail-row"><span style="color:#7a9a7e">Registration No.</span><strong>${regNo}</strong></div>
    <div class="detail-row"><span style="color:#7a9a7e">Program</span><span>${program}</span></div>
    <div class="detail-row"><span style="color:#7a9a7e">Total Amount Paid</span><strong style="color:#0F6E56">₹${amountPaid.toLocaleString('en-IN')}</strong></div>

    <div class="portal-box">
      <p>Log in to your <strong style="color:#C8EFE3;">Rider Portal</strong> to book your class slots now.</p>
      <a href="${CONFIG.RIDER_PORTAL_URL}?reg=${encodeURIComponent(regNo)}" class="portal-btn">🐴 Book My Classes</a>
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
    Logger.log(`Payment complete email sent to ${email} (${regNo})`);
}

/**
 * Sent after each partial payment — shows updated balance.
 */
function sendPartialPaymentEmail(name, email, regNo, program, amountPaidNow, totalAmountPaid, totalAmount) {
    const balance = totalAmount - totalAmountPaid;
    const subject = `💳 Payment Received — ₹${amountPaidNow.toLocaleString('en-IN')} | Balance: ₹${balance.toLocaleString('en-IN')} (${regNo})`;

    const htmlBody = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8">
<style>
  body{margin:0;padding:0;background:#f0f4f1;font-family:'Segoe UI',Arial,sans-serif}
  .wrap{max-width:560px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(4,52,44,0.08)}
  .hdr{background:#04342C;padding:24px;text-align:center}
  .hdr img{width:60px;height:60px;border-radius:50%;border:2px solid rgba(93,202,165,0.3);margin-bottom:10px}
  .hdr h1{margin:0;color:#C8EFE3;font-size:20px}
  .body{padding:24px}
  .detail-row{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid #f0f4f1;font-size:13px}
  .bal-box{background:#fff8e1;border:1px solid #ffc107;border-radius:8px;padding:14px;text-align:center;margin:18px 0}
  .bal-box p{margin:0;font-size:13px;color:#856404}
  .bal-amt{font-size:28px;font-weight:700;color:#BA7517;margin:4px 0}
  .ftr{background:#f0f4f1;padding:12px;text-align:center;font-size:11px;color:#7a9a7e;border-top:1px solid #dce8de}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="Kings Equestrian">
    <h1>Payment Received</h1>
  </div>
  <div class="body">
    <p>Dear <strong>${name}</strong>, we've received your payment.</p>
    <div class="detail-row"><span style="color:#7a9a7e">Reg No.</span><strong>${regNo}</strong></div>
    <div class="detail-row"><span style="color:#7a9a7e">Program</span><span>${program}</span></div>
    <div class="detail-row"><span style="color:#7a9a7e">Paid This Time</span><strong style="color:#0F6E56">₹${amountPaidNow.toLocaleString('en-IN')}</strong></div>
    <div class="detail-row"><span style="color:#7a9a7e">Total Paid So Far</span><strong>₹${totalAmountPaid.toLocaleString('en-IN')}</strong></div>
    <div class="detail-row"><span style="color:#7a9a7e">Program Total</span><span>₹${totalAmount.toLocaleString('en-IN')}</span></div>

    <div class="bal-box">
      <p>Remaining balance</p>
      <div class="bal-amt">₹${balance.toLocaleString('en-IN')}</div>
      <p>Pay the remaining amount at your convenience via the
        <a href="${CONFIG.PAYMENT_FORM_LINK}" style="color:#BA7517;">payment form</a>.
      </p>
    </div>

    <p style="font-size:12px;color:#7a9a7e;">
      You can continue using the <a href="${CONFIG.RIDER_PORTAL_URL}" style="color:#0F6E56;">Rider Portal</a>
      to book your classes. Full payment is not required to book slots.
    </p>
  </div>
  <div class="ftr">Kings Equestrian Foundation · Karnataka · +91 99807 71166</div>
</div>
</body>
</html>`;

    MailApp.sendEmail({ to: email, subject, htmlBody, name: 'Kings Equestrian Foundation' });
    Logger.log(`Partial payment email sent to ${email}: ₹${amountPaidNow} paid, ₹${balance} remaining`);
}

// ───────────────────────────────────────────────────────────
//  MENU HELPER — Resend Registration Email
// ───────────────────────────────────────────────────────────

/**
 * Admin can select a row in "Regular Booking Form Response" and
 * manually resend the registration email.
 */
function menuResendRegularRegistrationEmail() {
    const ui    = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.REGULAR_BOOKING_FORM);
    if (!sheet) { ui.alert('Regular Booking Form Response sheet not found'); return; }

    const sel = sheet.getActiveRange();
    if (!sel || sel.getRow() <= 1) { ui.alert('Please select a student row (not the header)'); return; }

    const row = sel.getRow();
    const C   = CONFIG.REG_BOOKING_COLS;

    const name    = sheet.getRange(row, C.NAME + 1).getValue();
    const email   = sheet.getRange(row, C.EMAIL_ID + 1).getValue();
    const phone   = sheet.getRange(row, C.PHONE_NUMBER + 1).getValue();
    const service = sheet.getRange(row, C.OUR_SERVICES + 1).getValue();
    let   regNo   = sheet.getRange(row, C.REG_NO + 1).getValue();

    if (!email) { ui.alert('No email found in selected row'); return; }

    // If no reg no yet, generate one now
    if (!regNo || String(regNo).startsWith('⚠️')) {
        regNo = generateRegularRef();
        sheet.getRange(row, C.REG_NO + 1).setValue(regNo);
        Logger.log(`Generated new KER ref on resend: ${regNo}`);
    }

    const programs = (() => { try { return getRegularProgramsData(); } catch(e) { return {}; } })();
    const program  = resolveProgram(String(service), programs) || String(service);
    const progData = programs[program];

    try {
        sendRegularRegistrationEmail({ name, email, phone, regNo, program, progData, participants: 1 });
        sheet.getRange(row, C.REG_EMAIL_SENT + 1).setValue('Resent').setBackground('#d4edda').setFontColor('#155724');
        sheet.getRange(row, C.REG_EMAIL_TIMESTAMP + 1).setValue(new Date()).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
        ui.alert(`✅ Registration email resent to ${email}\nReg No: ${regNo}`);
    } catch (err) {
        ui.alert(`❌ Failed to send: ${err.message}`);
    }
}