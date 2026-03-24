// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 3_Emails.gs
// All email construction and sending functions
// ============================================================

// ────────────────────────────────────────────────────────────
//  WELCOME EMAIL
//  Sent on every new booking form submission (first-time or not)
// ────────────────────────────────────────────────────────────

function sendWelcomeEmail(d) {
  // d: { name, email, phone, services, participants, amount, keNo,
  //      upiLink, qrCode, bookingDate, isFirstTime, sheet, row }

  const attachments = [];

  // Terms PDF
  const termsPDF = getTermsPDF();
  if (termsPDF) attachments.push(termsPDF);
   const additionalPDF = getAdditionalPDF();
  if (additionalPDF) {
    attachments.push(additionalPDF);
    Logger("There is a PDF")
  }

  // Consent form
  try {
    const consentPDF = generateConsentPDF(d.name, d.email, d.phone, d.bookingDate);
    if (consentPDF) attachments.push(consentPDF);
    
  } catch (e) { Logger.log('Consent PDF error: ' + e); }

  // Service-specific PDFs
  try {
    const pricing = getPricingData();
    Object.keys(pricing).forEach(svc => {
      if (String(d.services).toLowerCase().includes(svc.toLowerCase())) {
        const pdf = getServicePDF(pricing[svc].docId, svc);
        if (pdf) attachments.push(pdf);
      }
    });
  } catch (e) { Logger.log('Service PDF error: ' + e); }

  const subject = d.isFirstTime
    ? 'Welcome to Kings Equestrian! Your KE No: ' + d.keNo
    : 'New Booking Confirmed — Kings Equestrian (' + d.keNo + ')';

  const greeting = d.isFirstTime
    ? '<h2 style="color:#1f4e3d;margin:0 0 8px">Welcome, ' + d.name + '! </h2><p>You have been registered with Kings Equestrian. Your <strong>KE Number is ' + d.keNo + '</strong> — keep this safe, you\'ll need it for future bookings and payments.</p>'
    : '<h2 style="color:#1f4e3d;margin:0 0 8px">Hi ' + d.name + '!</h2><p>A new booking has been received for your account <strong>(' + d.keNo + ')</strong>.</p>';

  const htmlBody = `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:0;color:#333">
<div style="max-width:640px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.1)">
  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;text-align:center;color:#fff">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" style="width:72px;height:72px;border-radius:50%;border:3px solid rgba(255,255,255,.4);margin-bottom:12px">
    <h1 style="margin:0;font-size:24px">Kings Equestrian Foundation</h1>
    <p style="margin:6px 0 0;font-size:13px;opacity:.9">Where horses don't just carry you — they change you</p>
  </div>
  <div style="padding:28px 30px">
    ${greeting}
    <div style="background:#f0f8f0;border-left:4px solid #2c5f2d;padding:14px 18px;margin:20px 0;border-radius:4px">
      <p style="margin:0;font-size:13px"><strong>KE Number:</strong> <span style="font-size:20px;color:#1f4e3d;font-weight:bold">${d.keNo}</span><br>
      <strong>Service:</strong> ${d.services}<br>
      <strong>Participants:</strong> ${d.participants}</p>
    </div>
    <div style="background:#e8f5e9;border:2px solid #4caf50;padding:20px;border-radius:8px;margin:20px 0">
      <h3 style="color:#2e7d32;margin:0 0 12px">Pay Advance — ₹${d.amount.toLocaleString('en-IN')}</h3>
      <p style="font-size:13px;color:#555;margin:0 0 16px">Scan the QR code below and then submit the payment confirmation form.</p>
      <div style="text-align:center;margin:16px 0">
        <img src="${d.qrCode}" style="width:160px;height:160px;border:2px solid #e0e0e0;border-radius:6px">
      </div>
      <div style="text-align:center;margin-top:14px">
        <a href="${CONFIG.PAYMENT_FORM_LINK}" style="background:#1f4e3d;color:#fff;padding:12px 26px;text-decoration:none;border-radius:6px;font-weight:bold;font-size:14px;display:inline-block"> Submit Payment</a>
      </div>
      <p style="font-size:11px;color:#777;margin:12px 0 0;text-align:center">After paying, click the button to upload your screenshot and select date/time</p>
    </div>
    <div style="background:#f9f9f9;padding:16px;border-radius:8px">
      <h4 style="color:#1f4e3d;margin:0 0 10px"> What's Next</h4>
      <ol style="margin:0;padding-left:20px;font-size:13px;color:#555;line-height:1.9">
        <li>Pay ₹${d.amount.toLocaleString('en-IN')} advance via the QR code above</li>
        <li>Submit payment via the form and choose your date &amp; time</li>
        <li>Review the attached Terms &amp; Conditions${additionalPDF ? ', Additional Information' : ''}, and Consent Form</li>
        <li>Await your payment receipt &amp; confirmation email</li>
        <li>Arrive 15 min before your slot — wear comfortable shoes!</li>
      </ol>
    </div>
  </div>
  <div style="background:#1f4e3d;color:#fff;padding:18px 30px;text-align:center;font-size:12px">
    <strong>Kings Equestrian Foundation</strong><br>Karnataka, India<br> +91-9980895533 &nbsp;|&nbsp; info@kingsequestrian.com
  </div>
</div>
</body></html>`;

  const ccEmails = getCCRecipients('welcome');
  GmailApp.sendEmail(
    d.email,
    subject,
    '',
    {
      htmlBody    : htmlBody,
      attachments : attachments,
      cc          : ccEmails.join(','),
      name        : 'Kings Equestrian Foundation'
    }
  );

  // Mark sent in sheet
  if (d.sheet && d.row) {
    d.sheet.getRange(d.row, CONFIG.BOOKING_COLS.WELCOME_SENT + 1)
      .setValue('Yes').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    d.sheet.getRange(d.row, CONFIG.BOOKING_COLS.WELCOME_AT + 1)
      .setValue(new Date()).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
  }

  Logger.log('Welcome email sent to ' + d.email + ' (KE: ' + d.keNo + ')');
}

// ────────────────────────────────────────────────────────────
//  RECEIPT EMAIL HTML BUILDER
// ────────────────────────────────────────────────────────────

function buildReceiptEmailHTML(d) {
  // d: { name, keNo, amount, txnRef, payDate, receiptNo }
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:0;color:#333">
<div style="max-width:620px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.1)">
  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;text-align:center;color:#fff">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" style="width:64px;height:64px;border-radius:50%;border:3px solid rgba(255,255,255,.4);margin-bottom:10px">
    <h1 style="margin:0;font-size:22px">Kings Equestrian Foundation</h1>
  </div>
  <div style="padding:28px 30px">
    <p style="font-size:16px">Dear <strong>${d.name}</strong>,</p>
    <div style="background:#d4edda;border-left:4px solid #28a745;padding:16px;border-radius:6px;margin:18px 0;text-align:center">
      <div style="font-size:20px;font-weight:bold;color:#155724">✅ Payment Confirmed!</div>
      <p style="margin:6px 0 0;font-size:13px">Your 80G receipt is attached to this email.</p>
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;margin:16px 0">
      <tr><td style="padding:7px 0;color:#666;width:40%">KE Number</td><td style="padding:7px 0;font-weight:600;color:#1f4e3d">${d.keNo}</td></tr>
      <tr><td style="padding:7px 0;color:#666">Receipt No</td><td style="padding:7px 0;font-weight:600">${d.receiptNo}</td></tr>
      <tr><td style="padding:7px 0;color:#666">Amount Paid</td><td style="padding:7px 0;font-size:18px;font-weight:700;color:#28a745">₹${Number(d.amount).toLocaleString('en-IN')}</td></tr>
      ${d.txnRef ? `<tr><td style="padding:7px 0;color:#666">Transaction ID</td><td style="padding:7px 0;font-weight:600">${d.txnRef}</td></tr>` : ''}
      ${d.payDate ? `<tr><td style="padding:7px 0;color:#666">Payment Date</td><td style="padding:7px 0">${fmtDate(d.payDate)}</td></tr>` : ''}
    </table>
    <div style="background:#fff8e1;border-left:4px solid #ffc107;padding:13px;border-radius:4px;font-size:12px;color:#856404">
      This receipt is eligible for deduction under Section 80G of the Income Tax Act, 1961.
    </div>
  </div>
  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">
    <strong>Kings Equestrian Foundation</strong><br>Karnataka, India &nbsp;|&nbsp; +91-9980895533 &nbsp;|&nbsp; info@kingsequestrian.com
  </div>
</div>
</body></html>`;
}

// ────────────────────────────────────────────────────────────
//  ATTENDANCE ACKNOWLEDGMENT EMAIL  (sent when marked Present)
// ────────────────────────────────────────────────────────────

function sendAttendanceAckEmail(d) {
  // d: { name, email, keNo, service, timeSlot, date, participants }
  if (!d.email || !d.email.includes('@')) return;

  const firstName = d.name.split(' ')[0];
  const subject   = '🐴 Attendance Confirmed — Kings Equestrian (' + d.keNo + ')';

  const htmlBody = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f4f6f4;margin:0;padding:0;color:#333">
<div style="max-width:580px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.08)">
  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:26px 28px;text-align:center;color:#fff">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" style="width:68px;height:68px;border-radius:50%;border:3px solid rgba(255,255,255,.4);margin-bottom:10px">
    <h1 style="margin:0;font-size:21px">Welcome, ${firstName}! 🐴</h1>
  </div>
  <div style="padding:26px 28px">
    <div style="background:#d4edda;border-left:4px solid #28a745;border-radius:6px;padding:14px 18px;margin-bottom:20px">
      <div style="font-weight:bold;color:#155724;font-size:15px">✅ Attendance Confirmed — Present</div>
      <div style="color:#1e7e34;font-size:12px;margin-top:4px">${fmtDate(new Date())}</div>
    </div>
    <p style="font-size:14px;line-height:1.7">Hi <strong>${d.name}</strong>, we're happy you're here today! Your attendance has been recorded.</p>
    <div style="background:#f8faf8;border:1px solid #c8e6c9;border-radius:8px;padding:16px;margin:16px 0;font-size:13px">
      <div style="margin-bottom:6px"><strong>KE No:</strong> ${d.keNo}</div>
      <div style="margin-bottom:6px"><strong>Service:</strong> ${d.service}</div>
      ${d.timeSlot ? '<div style="margin-bottom:6px"><strong>Time:</strong> ' + d.timeSlot + '</div>' : ''}
      ${d.participants > 1 ? '<div><strong>Participants:</strong> ' + d.participants + '</div>' : ''}
    </div>
    <div style="background:#fff8e6;border-left:4px solid #f0a500;padding:14px;border-radius:4px;font-size:12px;color:#7a5000">
      <strong>Tips:</strong> Stay calm around the horses, follow your instructor, wear your helmet, and enjoy!
    </div>
  </div>
  <div style="background:#1f4e3d;color:#fff;padding:16px 28px;text-align:center;font-size:12px">
    Kings Equestrian Foundation &nbsp;|&nbsp; Karnataka &nbsp;|&nbsp; +91-9980895533
  </div>
</div>
</body></html>`;

  const ccEmails = getCCRecipients('welcome');
  GmailApp.sendEmail(
    d.email,
    subject,
    '',
    {
      htmlBody : htmlBody,
      cc       : ccEmails.join(','),
      name     : 'Kings Equestrian Foundation'
    }
  );
  Logger.log('Attendance ack sent to ' + d.email);
}

// ────────────────────────────────────────────────────────────
//  80G RECEIPT PDF GENERATOR
// ────────────────────────────────────────────────────────────

// ────────────────────────────────────────────────────────────
//  80G RECEIPT PDF GENERATOR
//  Uses DocumentApp (Docs scope) instead of DriveApp.createFile
//  so it works without the Drive scope.
//  Flow: create a Google Doc → write content → export as PDF
//        → trash the temp Doc immediately
// ────────────────────────────────────────────────────────────

function generate80GReceipt(riderName, pan, amount, txnRef, receiptNo) {
  const logoB64  = imgBase64FromUrl('https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg');
  const stampB64 = imgBase64FromDrive('1fQVqA1ABWCaTJs4uJVxiNqIGhl5iWugJ');
  const signB64  = imgBase64FromDrive('1CI6H0JgysxanA0RimUwu7QwSSRospSwc');
  const dateStr  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const words    = numberToWords(amount);

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
@page{size:A4;margin:0}
body{font-family:"Times New Roman",serif;margin:0;padding:24px;background:#fff}
.box{border:2px solid #000;border-radius:32px;padding:24px 28px;max-width:800px;margin:auto;position:relative}
.hdr{display:flex;align-items:flex-start;gap:14px}
.logo{width:100px;text-align:center}
.logo img{width:90px}
.hdr-c{flex:1;text-align:center}
.org{font-size:26px;font-weight:bold}
.reg{font-size:12px;margin-top:4px}
.sub{font-size:12px;margin-top:3px}
.tagline{margin-top:8px;font-style:italic;font-weight:bold;text-decoration:underline;font-size:12px}
.rno{position:absolute;right:28px;top:14px;font-size:15px;font-weight:bold;color:red}
.rbox{border:2px solid #000;border-radius:10px;text-align:center;padding:8px;margin:18px 0 8px}
.rtitle{font-size:19px;font-weight:bold}
.rsub{font-size:11px}
.date-r{text-align:right;font-size:13px;margin-bottom:8px}
.mcols{display:flex;gap:28px;margin-top:8px}
.col{flex:1;font-size:13px}
.cb{display:inline-block;width:12px;height:12px;border:1px solid #000;margin-right:5px;vertical-align:middle}
.cb.on{background:#000;position:relative}
.cb.on::after{content:"✓";color:#fff;font-size:10px;position:absolute;left:0px;top:-3px}
.dr{margin:6px 0}
.dl{font-weight:bold}
.amt-box{border:2px solid #000;margin:18px 0;padding:16px;position:relative;text-align:center}
.rs{position:absolute;left:18px;top:50%;transform:translateY(-50%);font-size:38px;color:goldenrod;font-weight:bold}
.av{font-size:32px;font-weight:bold}
.pm{font-size:13px;margin-top:8px}
.decl{margin-top:14px;font-size:12px;text-align:justify}
.sig{margin-top:32px;text-align:right}
.org-lbl{font-weight:bold;margin-bottom:4px}
.sign-area{position:relative;height:110px}
.sign-area img.sig-img{width:100px}
.sign-area img.stp-img{width:110px}
.auth{margin-top:84px;text-decoration:underline;font-size:13px}
</style></head>
<body>
<div class="box">
  <div class="rno">${receiptNo}</div>
  <div class="hdr">
    <div class="logo"><img src="${logoB64}"></div>
    <div class="hdr-c">
      <div class="org">Kings Equestrian Foundation</div>
      <div class="reg">Registered u/s 80G | Reg No: AAJCK7191GE20231 | PAN: AAJCK7191G</div>
      <div class="sub">K202, Tower-6, Jacaranda Block, Devarabisanahalli, Bellandur S.O, Bengaluru – 560103 Karnataka, India</div>
      <div class="sub">kingsequestrianfoundation@gmail.com | kingsequestrianfoundation.com</div>
      <div class="tagline">We gratefully acknowledge your generous contribution in support of our programmes.</div>
    </div>
  </div>
  <div class="rbox">
    <div class="rtitle">Receipt</div>
    <div class="rsub">Issued in compliance with Rule 18AB and Form 10BD requirements</div>
  </div>
  <div class="date-r"><strong>Date:</strong> ${dateStr}</div>
  <div class="mcols">
    <div class="col">
      <div style="font-weight:bold;margin-bottom:8px">Donor Category (✓ Tick Applicable)</div>
      <div><span class="cb on"></span> Resident Indian Donor</div>
      <div style="margin-top:5px"><span class="cb"></span> Non-Resident Indian (NRI)</div>
    </div>
    <div class="col">
      <div style="font-weight:bold;margin-bottom:8px">Donor Details</div>
      <div class="dr"><span class="dl">Name:</span> ${riderName}</div>
      <div class="dr"><span class="dl">PAN / Aadhaar:</span> ${pan || 'N/A'}</div>
      <div class="dr"><span class="dl">Amount in Words:</span> ${words}</div>
    </div>
  </div>
  <div class="amt-box">
    <span class="rs">₹</span>
    <div class="av">${Number(amount).toLocaleString('en-IN')}</div>
  </div>
  <div class="pm">
    <strong>Mode of Payment:</strong> NEFT / RTGS / UPI (Cash not eligible u/s 80G)<br><br>
    ${txnRef && txnRef !== 'N/A' ? '<strong>Transaction Reference:</strong> ' + txnRef + '<br><br>' : ''}
    <strong>Amount in Words:</strong> ${words}
  </div>
  <div class="decl">
    Certified that the above donation is received by trust for charitable purposes only.
    This donation is eligible for deduction under Section 80G of the Income Tax Act, 1961.
    This receipt will be reported in Form 10BD and Form 10BE will be issued to the donor.
  </div>
  <div class="sig">
    <div class="org-lbl">For Kings Equestrian Foundation</div>
    <div class="sign-area">
      <img class="sig-img" src="${signB64}">
      <img class="stp-img" src="${stampB64}">
    </div>
  </div>
</div>
</body></html>`;

  const tmp  = DriveApp.createFile('receipt_temp_' + Date.now() + '.html', html, MimeType.HTML);
  const blob = tmp.getAs('application/pdf');
  blob.setName('80G_Receipt_' + riderName.replace(/\s+/g,'_') + '_' + receiptNo.replace(/\//g,'-') + '.pdf');
  tmp.setTrashed(true);
  return blob;
}