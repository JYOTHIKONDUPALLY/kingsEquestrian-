// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 3_Emails.gs
// All email construction and sending functions
// ============================================================
function formatPrefDate(dateVal) {
  if (!dateVal) return '';
  var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return String(dateVal);
  var days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var months = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  return days[d.getDay()] + ', ' + months[d.getMonth()] + ' ' + d.getDate() + ' ' + d.getFullYear();
}

function formatPrefTime(timeVal) {
  if (!timeVal) return '';
  // handles "18:00 - 18:30" or "18:00-18:30" style strings
  return String(timeVal).replace(/(\d{1,2}):(\d{2})/g, function(_, h, m) {
    var hour = parseInt(h, 10);
    var ampm = hour >= 12 ? 'PM' : 'AM';
    var h12  = hour % 12 || 12;
    return h12 + ':' + m + ' ' + ampm;
  });
}
// ────────────────────────────────────────────────────────────
//  WELCOME EMAIL
//  FIX #7/#9: Removed getAdditionalPDF() attachment entirely.
//             The services brochure is shown as a clickable button/link only.
//             This avoids the SlidesApp/DocumentApp permission error.
//  FIX #1: Removed emojis from email subject lines.
//  FIX #6: All sends use GmailApp (not MailApp) to match script permissions.
// ────────────────────────────────────────────────────────────

function sendWelcomeEmail(d) {
  // d: { name, email, phone, services, participants, amount, keNo,
  //      upiLink, qrCode, prefDate, isFirstTime, sheet, row }

  const attachments = [];
  // Terms PDF
  try {
    const termsPDF = getTermsPDF();
    if (termsPDF) attachments.push(termsPDF);
  } catch (e) { Logger.log('Terms PDF error: ' + e); }

  // Consent form
  try {
    const consentPDF = generateConsentPDF(d.name, d.email, d.phone, d.prefDate);
    if (consentPDF) attachments.push(consentPDF);
  } catch (e) { Logger.log('Consent PDF error: ' + e); }

  // Service-specific PDFs from Pricing sheet
  try {
    const pricing = getPricingData();
    Object.keys(pricing).forEach(svc => {
      if (String(d.services).toLowerCase().includes(svc.toLowerCase())) {
        const pdf = getServicePDF(pricing[svc].docId, svc);
        if (pdf) attachments.push(pdf);
      }
    });
  } catch (e) { Logger.log('Service PDF error: ' + e); }

  // FIX #7/#9: Use the direct Drive link — no PDF download needed
  const servicesBrochureLink = CONFIG.ADDITIONAL_PDF_DRIVE_LINK || '';

  const Myrides=CONFIG.MYRIDES;

const servicesBrochureBlock = servicesBrochureLink
  ? '<div style="background:#e8f4ff;border-left:4px solid #2196f3;padding:13px 18px;margin:16px 0;border-radius:4px">'
    + '<p style="margin:0;font-size:13px"><strong>Learn about our services:</strong><br>'
    + '<a href="' + servicesBrochureLink + '" style="color:#1565c0;font-weight:600" target="_blank">View Our Services Guide</a>'
    + '<br><br><strong>Book your future rides here:</strong><br>'
    + '<a href="' + Myrides + '" style="color:#1565c0;font-weight:600" target="_blank">My Rides</a>'
    + '</p></div>'
  : '';


  // FIX #1: No emojis in subject
  const subject = d.isFirstTime
    ? 'Welcome to Kings Equestrian! Your KE No: ' + d.keNo
    : 'New Booking Confirmed - Kings Equestrian (' + d.keNo + ')';

  const greeting = d.isFirstTime
    ? '<h2 style="color:#1f4e3d;margin:0 0 8px">Welcome, ' + d.name + '!</h2>'
      + '<p>You have been registered with Kings Equestrian. Your <strong>KE Number is ' + d.keNo + '</strong> — keep this safe for future bookings and payments.</p>'
    : '<h2 style="color:#1f4e3d;margin:0 0 8px">Hi ' + d.name + '!</h2>'
      + '<p>A new booking has been received for your account <strong>(' + d.keNo + ')</strong>.</p>';

const htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
+ '<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:0;color:#333">'
+ '<div style="max-width:640px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.1)">'
+ '  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;text-align:center;color:#fff">'
+ '   <img src="https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb"'
+  '   style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px"> '
+ '    <h1 style="margin:0;font-size:24px">Kings Equestrian Foundation</h1>'
+ '    <p style="margin:6px 0 0;font-size:13px;opacity:.9">Where Nature Connects. Riders Transform.</p>'
+ '  </div>'
+ '  <div style="padding:28px 30px">'
+ '    ' + greeting
+ '    <div style="background:#f0f8f0;border-left:4px solid #2c5f2d;padding:14px 18px;margin:20px 0;border-radius:4px">'
+ '      <p style="margin:0;font-size:13px">'
+ '      <strong>KE Number:</strong> <span style="font-size:20px;color:#1f4e3d;font-weight:bold">' + d.keNo + '</span><br>'
+ '      <strong>Service:</strong> ' + d.services + '<br>'
+ '      <strong>Participants:</strong> ' + d.participants + '<br>'
+ '      <strong>Preferred Date:</strong> <br>'+formatPrefDate(d.prefDate)
+ '      <strong>Time Slot:</strong> '+formatPrefTime(d.prefTime)
+ '      </p>'
+ '    </div>'

  // ── Kings Equestrian Experience Details Block ──────────────────────────────
  + '    <div style="background:#f4faf4;border:1px solid #c8e6c9;border-radius:8px;padding:20px 22px;margin:20px 0;text-align:center">'
  + '      <p style="font-size:20px;margin:0 0 4px"></p>'
  + '      <h3 style="color:#1f4e3d;margin:0 0 2px;font-size:18px;letter-spacing:1px">KINGS EQUESTRIAN</h3>'
  + '      <p style="font-size:13px;color:#2c5f2d;font-style:italic;margin:0 0 10px">Where Nature Connects. Riders Transform.</p>'
  + '      <hr style="border:none;border-top:1px solid #c8e6c9;margin:10px 0">'

  + '      <p style="font-size:13px;color:#444;margin:10px 0 4px"> <strong>Begin Your Ride</strong></p>'
  + '      <p style="font-size:13px;color:#555;margin:0 0 6px">A premium equestrian experience near Sarjapur —<br>perfect for beginners, families &amp; riders.</p>'
  + '      <p style="font-size:13px;color:#2c5f2d;font-style:italic;margin:0 0 10px"> Calm. Confidence. Connection.</p>'
  + '      <hr style="border:none;border-top:1px solid #c8e6c9;margin:10px 0">'

  + '      <p style="font-size:13px;color:#555;margin:10px 0 4px"> <strong>Location</strong><br>'
  + '      <a href="https://maps.app.goo.gl/LKXQ8VTbYhrw1PDb9" style="color:#1565c0" target="_blank">https://maps.app.goo.gl/LKXQ8VTbYhrw1PDb9</a></p>'
  + '      <p style="font-size:13px;color:#555;margin:4px 0 10px"> <a href="https://www.instagram.com/kingsequestrianfoundation" style="color:#1565c0" target="_blank">@kingsequestrianfoundation</a></p>'
  + '      <hr style="border:none;border-top:1px solid #c8e6c9;margin:10px 0">'

  + '      <p style="font-size:13px;color:#444;margin:10px 0 8px"> <strong>Horse Safari Experiences</strong></p>'
  + '      <table style="width:80%;border-collapse:collapse;font-size:13px;margin:0 auto">'
  + '        <tbody>'
  + '          <tr style="background:#fff">'
  + '            <td style="padding:8px 14px;border-bottom:1px solid #e0e0e0;text-align:left">Short Safari <span style="color:#777">(30 mins)</span></td>'
  + '            <td style="padding:8px 14px;border-bottom:1px solid #e0e0e0;text-align:right;font-weight:bold;color:#1f4e3d">₹1,500</td>'
  + '          </tr>'
  + '          <tr style="background:#f9f9f9">'
  + '            <td style="padding:8px 14px;text-align:left">Long Safari <span style="color:#777">(1 hour)</span></td>'
  + '            <td style="padding:8px 14px;text-align:right;font-weight:bold;color:#1f4e3d">₹2,500</td>'
  + '          </tr>'
  + '        </tbody>'
  + '      </table>'
  + '      <hr style="border:none;border-top:1px solid #c8e6c9;margin:10px 0">'

  + '      <p style="font-size:12px;color:#555;margin:10px 0;background:#fffde7;border-left:3px solid #f9a825;padding:10px 14px;text-align:left;border-radius:4px">'
  + '         Please refer to the <strong>Kings Equestrian brochure</strong> attached to this email for comprehensive services and detailed pricing.'
  + '      </p>'
  + '      <hr style="border:none;border-top:1px solid #c8e6c9;margin:10px 0">'

  + '      <p style="font-size:13px;color:#2c5f2d;font-style:italic;margin:10px 0 0">Not just a ride — an experience you\'ll return to. </p>'
  + '    </div>'
  // ── End Experience Block ───────────────────────────────────────────────────
  + '    ' + servicesBrochureBlock
  + '    <div style="background:#e8f5e9;border:2px solid #4caf50;padding:20px;border-radius:8px;margin:20px 0">'
  + '      <h3 style="color:#2e7d32;margin:0 0 12px">Pay Advance - Rs.' + Number(d.amount).toLocaleString('en-IN') + '</h3>'
  + '      <p style="font-size:13px;color:#555;margin:0 0 16px">Scan the QR code below and then submit the payment confirmation form.</p>'
  + '      <div style="text-align:center;margin:16px 0">'
  + '        <img src="' + d.qrCode + '" style="width:160px;height:160px;border:2px solid #e0e0e0;border-radius:6px">'
  + '      </div>'
  + '      <div style="text-align:center;margin-top:14px">'
  + '        <a href="' + CONFIG.PAYMENT_FORM_LINK + '" style="background:#1f4e3d;color:#fff;padding:12px 26px;text-decoration:none;border-radius:6px;font-weight:bold;font-size:14px;display:inline-block">Submit Payment</a>'
  + '      </div>'
  + '      <p style="font-size:11px;color:#777;margin:12px 0 0;text-align:center">After paying, click the button to upload your screenshot and select date/time</p>'
  + '    </div>'
  + '    <div style="background:#f9f9f9;padding:16px;border-radius:8px">'
  + '      <h4 style="color:#1f4e3d;margin:0 0 10px">What\'s Next</h4>'
  + '      <ol style="margin:0;padding-left:20px;font-size:13px;color:#555;line-height:1.9">'
  + '        <li>Pay Rs.' + Number(d.amount).toLocaleString('en-IN') + ' advance via the QR code above</li>'
  + '        <li>Submit payment via the form and choose your date and time</li>'
  + '        <li>Review the attached Terms and Conditions and Consent Form</li>'
  + '        <li>Await your payment receipt and confirmation email</li>'
  + '        <li>Arrive 15 min before your slot — wear comfortable shoes!</li>'
  + '      </ol>'
  + '    </div>'
  + '  </div>'
  + '  <div style="background:#1f4e3d;color:#fff;padding:18px 30px;text-align:center;font-size:12px">'
  + '    <strong>Kings Equestrian Foundation</strong><br>Karnataka, India<br>+91-9980895533 | info@kingsequestrian.com'
  + '  </div>'
  + '</div>'
  + '</body></html>';

  const ccEmails = getCCRecipients('welcome');
  // FIX #6: Use GmailApp (not MailApp) — requires Gmail send permission
  GmailApp.sendEmail(d.email, subject, '', {
    htmlBody    : htmlBody,
    attachments : attachments,
    cc          : ccEmails.join(','),
    name        : 'Kings Equestrian Foundation'
  });

  if (d.sheet && d.row) {
    d.sheet.getRange(d.row, CONFIG.BOOKING_COLS.WELCOME_SENT + 1).setValue('Yes').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    d.sheet.getRange(d.row, CONFIG.BOOKING_COLS.WELCOME_AT + 1).setValue(new Date()).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
  }

  Logger.log('Welcome email sent to ' + d.email + ' (KE: ' + d.keNo + ')');
}

// ────────────────────────────────────────────────────────────
//  RECEIPT EMAIL HTML BUILDER
// ────────────────────────────────────────────────────────────

function buildReceiptEmailHTML(d) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:0;color:#333">'
    + '<div style="max-width:620px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.1)">'
    + '  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;text-align:center;color:#fff">'
    + '   <img src="https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb"'
+  '   style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px"> '
    + '    <h1 style="margin:0;font-size:22px">Kings Equestrian Foundation</h1>'
    + '  </div>'
    + '  <div style="padding:28px 30px">'
    + '    <p style="font-size:16px">Dear <strong>' + d.name + '</strong>,</p>'
    + '    <div style="background:#d4edda;border-left:4px solid #28a745;padding:16px;border-radius:6px;margin:18px 0;text-align:center">'
    + '      <div style="font-size:20px;font-weight:bold;color:#155724">Payment Confirmed!</div>'
    + '      <p style="margin:6px 0 0;font-size:13px">Your 80G receipt is attached to this email.</p>'
    + '    </div>'
    + '    <table style="width:100%;border-collapse:collapse;font-size:13px;margin:16px 0">'
    + '      <tr><td style="padding:7px 0;color:#666;width:40%">KE Number</td><td style="padding:7px 0;font-weight:600;color:#1f4e3d">' + d.keNo + '</td></tr>'
    + '      <tr><td style="padding:7px 0;color:#666">Receipt No</td><td style="padding:7px 0;font-weight:600">' + d.receiptNo + '</td></tr>'
    + '      <tr><td style="padding:7px 0;color:#666">Amount Paid</td><td style="padding:7px 0;font-size:18px;font-weight:700;color:#28a745">Rs.' + Number(d.amount).toLocaleString('en-IN') + '</td></tr>'
    + (d.txnRef ? '      <tr><td style="padding:7px 0;color:#666">Transaction ID</td><td style="padding:7px 0;font-weight:600">' + d.txnRef + '</td></tr>' : '')
    + (d.payDate && fmtDate(d.payDate) ? '      <tr><td style="padding:7px 0;color:#666">Payment Date</td><td style="padding:7px 0">' + fmtDate(d.payDate) + '</td></tr>' : '')
    + '    </table>'
    + '    <div style="background:#fff8e1;border-left:4px solid #ffc107;padding:13px;border-radius:4px;font-size:12px;color:#856404">'
    + '      This receipt is eligible for deduction under Section 80G of the Income Tax Act, 1961.'
    + '    </div>'
    + '  </div>'
    + '  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">'
    + '    <strong>Kings Equestrian Foundation</strong><br>Karnataka, India | +91-9980895533 | info@kingsequestrian.com'
    + '  </div>'
    + '</div>'
    + '</body></html>';
}

// ────────────────────────────────────────────────────────────
//  BOOKING CONFIRMATION EMAIL (portal bookings)
//  FIX #1: No emojis in subject
// ────────────────────────────────────────────────────────────

function sendBookingConfirmationEmail(d) {
  if (!d.email || !d.email.includes('@')) return;

  const sessionRows = (d.added || []).map(s =>
    '<tr style="border-bottom:1px solid #e8f0e8">'
    + '<td style="padding:9px 12px;font-weight:600;color:#1f4e3d">' + s.service + '</td>'
    + '<td style="padding:9px 12px">' + s.date + '</td>'
    + '<td style="padding:9px 12px;color:#555">' + (s.timeSlot || 'Time TBD') + '</td>'
    + '</tr>'
  ).join('');

  const htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:0;color:#333">'
    + '<div style="max-width:620px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.1)">'
    + '  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:26px 30px;text-align:center;color:#fff">'
    + '    <img src="https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb"'
+  '   style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px"> '
    + '    <h1 style="margin:0;font-size:22px">Sessions Booked</h1>'
    + '    <p style="margin:6px 0 0;font-size:13px;opacity:.9">Kings Equestrian Foundation</p>'
    + '  </div>'
    + '  <div style="padding:26px 30px">'
    + '    <p style="font-size:15px">Dear <strong>' + d.name + '</strong>,</p>'
    + '    <div style="background:#d4edda;border-left:4px solid #28a745;padding:14px 18px;border-radius:6px;margin:16px 0">'
    + '      <strong style="color:#155724">Your ' + d.added.length + ' session' + (d.added.length !== 1 ? 's have' : ' has') + ' been confirmed!</strong><br>'
    + '      <span style="font-size:12px;color:#1e7e34">KE No: ' + d.keNo + '</span>'
    + '    </div>'
    + '    <table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">'
    + '      <thead><tr style="background:#1f4e3d;color:#fff">'
    + '        <th style="padding:9px 12px;text-align:left">Service</th>'
    + '        <th style="padding:9px 12px;text-align:left">Date</th>'
    + '        <th style="padding:9px 12px;text-align:left">Time Slot</th>'
    + '      </tr></thead>'
    + '      <tbody>' + sessionRows + '</tbody>'
    + '    </table>'
    + '    <div style="background:#fff8e6;border-left:4px solid #f0a500;padding:13px;border-radius:4px;font-size:12px;color:#7a5000;margin-top:16px">'
    + '      <strong>Reminder:</strong> Please ensure your advance payment is up to date. Arrive 15 minutes before your slot.'
    + '    </div>'
    + '    <div style="text-align:center;margin-top:20px">'
    + '      <a href="' + CONFIG.PAYMENT_FORM_LINK + '" style="background:#1f4e3d;color:#fff;padding:12px 26px;text-decoration:none;border-radius:6px;font-weight:bold;font-size:13px;display:inline-block">Submit Payment</a>'
    + '    </div>'
    + (d.errors && d.errors.length ? '<p style="font-size:11px;color:#c62828;margin-top:12px">Note: Some requests could not be processed — ' + d.errors.join(', ') + '</p>' : '')
    + '  </div>'
    + '  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">'
    + '    <strong>Kings Equestrian Foundation</strong><br>Karnataka, India | +91-9980895533 | info@kingsequestrian.com'
    + '  </div>'
    + '</div>'
    + '</body></html>';

  const ccEmails = getCCRecipients('welcome');
  // FIX #6: GmailApp
  GmailApp.sendEmail(d.email, 'Sessions Booked - Kings Equestrian (' + d.keNo + ')', '', {
    htmlBody : htmlBody,
    cc       : ccEmails.join(','),
    name     : 'Kings Equestrian Foundation'
  });
  Logger.log('Booking confirmation sent to ' + d.email);
}

// ────────────────────────────────────────────────────────────
//  PRESENT EMAIL
//  FIX #1: No emoji in subject. FIX #6: GmailApp
// ────────────────────────────────────────────────────────────

function sendPresentEmail(d) {
  const cleanEmail = String(d.email || '').trim();
  if (!cleanEmail || !cleanEmail.includes('@')) {
    Logger.log('sendPresentEmail skipped: invalid email [' + d.email + ']');
    return false;
  }

  const htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Georgia,serif;background:#f4f6f4;margin:0;padding:0;color:#333">'
    + '<div style="max-width:580px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,.08)">'
    + '  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;text-align:center;color:#fff">'
    + '    <img src="https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb"'
+  '   style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px"> '
    + '    <h1 style="margin:0;font-size:22px;font-family:Georgia,serif">Kings Equestrian Foundation</h1>'
    + '  </div>'
    + '  <div style="padding:32px 34px;line-height:1.9">'
    + '    <p style="font-size:15px;margin:0 0 18px">Dear <strong>' + d.name + '</strong>,</p>'
    + '    <p style="font-size:14px;margin:0 0 16px">It was truly a pleasure having you with us today.</p>'
    + '    <p style="font-size:14px;margin:0 0 16px">We hope your time with the horses brought calm, joy, and a beautiful sense of connection.</p>'
    + '    <p style="font-size:14px;margin:0 0 24px">Thank you for being part of our space — we look forward to welcoming you again soon.</p>'
    + '    <p style="font-size:14px;margin:0;color:#1f4e3d;font-style:italic">Warm regards,<br><strong>Kings Equestrian Foundation</strong></p>'
    + '  </div>'
    + '  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">'
    + '    Kings Equestrian Foundation | Karnataka | +91-9980895533'
    + '  </div>'
    + '</div>'
    + '</body></html>';

  const ccEmails = getCCRecipients('welcome');
  GmailApp.sendEmail(cleanEmail, 'Thank you for riding with us today - Kings Equestrian', '', {
    htmlBody : htmlBody,
    cc       : ccEmails.join(','),
    name     : 'Kings Equestrian Foundation'
  });
  Logger.log('Present email sent to ' + cleanEmail);
  return true;
}

// ────────────────────────────────────────────────────────────
//  NO-SHOW EMAIL
//  FIX #1: No emoji in subject. FIX #6: GmailApp
// ────────────────────────────────────────────────────────────

function sendNoShowEmail(d) {
  const cleanEmail = String(d.email || '').trim();
  if (!cleanEmail || !cleanEmail.includes('@')) {
    Logger.log('sendNoShowEmail skipped: invalid email [' + d.email + ']');
    return false;
  }

  const htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Georgia,serif;background:#f4f6f4;margin:0;padding:0;color:#333">'
    + '<div style="max-width:580px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,.08)">'
    + '  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;text-align:center;color:#fff">'
    + '    <img src="https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb"'
+  '   style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px"> '
    + '    <h1 style="margin:0;font-size:22px;font-family:Georgia,serif">Kings Equestrian Foundation</h1>'
    + '  </div>'
    + '  <div style="padding:32px 34px;line-height:1.9">'
    + '    <p style="font-size:15px;margin:0 0 18px">Dear <strong>' + d.name + '</strong>,</p>'
    + '    <p style="font-size:14px;margin:0 0 16px">We missed having you with us today and hope everything is well.</p>'
    + '    <p style="font-size:14px;margin:0 0 16px">Whenever you feel ready, we\'ll be happy to welcome you back.</p>'
    + '    <p style="font-size:14px;margin:0 0 24px">Wishing you ease and well-being,</p>'
    + '    <p style="font-size:14px;margin:0;color:#1f4e3d;font-style:italic"><strong>Kings Equestrian Foundation</strong></p>'
    + '  </div>'
    + '  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">'
    + '    Kings Equestrian Foundation | Karnataka | +91-9980895533'
    + '  </div>'
    + '</div>'
    + '</body></html>';

  const ccEmails = getCCRecipients('welcome');
  GmailApp.sendEmail(cleanEmail, 'We missed you today - Kings Equestrian', '', {
    htmlBody : htmlBody,
    cc       : ccEmails.join(','),
    name     : 'Kings Equestrian Foundation'
  });
  Logger.log('No-show email sent to ' + cleanEmail);
  return true;
}

// ────────────────────────────────────────────────────────────
//  80G RECEIPT PDF GENERATOR
// ────────────────────────────────────────────────────────────

function generate80GReceipt(riderName, pan, amount, txnRef, receiptNo) {
  const logoB64  = imgBase64FromUrl('https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb');
  const stampB64 = imgBase64FromDrive('1fQVqA1ABWCaTJs4uJVxiNqIGhl5iWugJ');
  const signB64  = imgBase64FromDrive('1CI6H0JgysxanA0RimUwu7QwSSRospSwc');
  const dateStr  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const words    = numberToWords(amount);

  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<style>'
    + '@page{size:A4;margin:0}'
    + 'body{font-family:"Times New Roman",serif;margin:0;padding:24px;background:#fff}'
    + '.box{border:2px solid #000;border-radius:32px;padding:24px 28px;max-width:800px;margin:auto;position:relative}'
    + '.hdr{display:flex;align-items:flex-start;gap:14px}'
    + '.logo{width:100px;text-align:center}'
    + '.logo img{width:90px}'
    + '.hdr-c{flex:1;text-align:center}'
    + '.org{font-size:26px;font-weight:bold}'
    + '.reg{font-size:12px;margin-top:4px}'
    + '.sub{font-size:12px;margin-top:3px}'
    + '.tagline{margin-top:8px;font-style:italic;font-weight:bold;text-decoration:underline;font-size:12px}'
    + '.rno{position:absolute;right:28px;top:14px;font-size:15px;font-weight:bold;color:red}'
    + '.rbox{border:2px solid #000;border-radius:10px;text-align:center;padding:8px;margin:18px 0 8px}'
    + '.rtitle{font-size:19px;font-weight:bold}'
    + '.rsub{font-size:11px}'
    + '.date-r{text-align:right;font-size:13px;margin-bottom:8px}'
    + '.mcols{display:flex;gap:28px;margin-top:8px}'
    + '.col{flex:1;font-size:13px}'
    + '.cb{display:inline-block;width:12px;height:12px;border:1px solid #000;margin-right:5px;vertical-align:middle}'
    + '.cb.on{background:#000;position:relative}'
    + '.cb.on::after{content:"v";color:#fff;font-size:10px;position:absolute;left:0px;top:-3px}'
    + '.dr{margin:6px 0}'
    + '.dl{font-weight:bold}'
    + '.amt-box{border:2px solid #000;margin:18px 0;padding:16px;position:relative;text-align:center}'
    + '.rs{position:absolute;left:18px;top:50%;transform:translateY(-50%);font-size:38px;color:goldenrod;font-weight:bold}'
    + '.av{font-size:32px;font-weight:bold}'
    + '.pm{font-size:13px;margin-top:8px}'
    + '.decl{margin-top:14px;font-size:12px;text-align:justify}'
    + '.sig{margin-top:32px;text-align:right}'
    + '.org-lbl{font-weight:bold;margin-bottom:4px}'
    + '.sign-area{position:relative;height:110px}'
    + '.sign-area img.sig-img{width:100px}'
    + '.sign-area img.stp-img{width:110px}'
    + '</style></head>'
    + '<body>'
    + '<div class="box">'
    + '  <div class="rno">' + receiptNo + '</div>'
    + '  <div class="hdr">'
    + '    <div class="logo"><img src="' + logoB64 + '"></div>'
    + '    <div class="hdr-c">'
    + '      <div class="org">Kings Equestrian Foundation</div>'
    + '      <div class="reg">Registered u/s 80G | Reg No: AAJCK7191GE20231 | PAN: AAJCK7191G</div>'
    + '      <div class="sub">K202, Tower-6, Jacaranda Block, Devarabisanahalli, Bellandur S.O, Bengaluru - 560103 Karnataka, India</div>'
    + '      <div class="sub">kingsequestrianfoundation@gmail.com | kingsequestrianfoundation.com</div>'
    + '      <div class="tagline">We gratefully acknowledge your generous contribution in support of our programmes.</div>'
    + '    </div>'
    + '  </div>'
    + '  <div class="rbox">'
    + '    <div class="rtitle">Receipt</div>'
    + '    <div class="rsub">Issued in compliance with Rule 18AB and Form 10BD requirements</div>'
    + '  </div>'
    + '  <div class="date-r"><strong>Date:</strong> ' + dateStr + '</div>'
    + '  <div class="mcols">'
    + '    <div class="col">'
    + '      <div style="font-weight:bold;margin-bottom:8px">Donor Category (tick Applicable)</div>'
    + '      <div><span class="cb on"></span> Resident Indian Donor</div>'
    + '      <div style="margin-top:5px"><span class="cb"></span> Non-Resident Indian (NRI)</div>'
    + '    </div>'
    + '    <div class="col">'
    + '      <div style="font-weight:bold;margin-bottom:8px">Donor Details</div>'
    + '      <div class="dr"><span class="dl">Name:</span> ' + riderName + '</div>'
    + '      <div class="dr"><span class="dl">PAN / Aadhaar:</span> ' + (pan || 'N/A') + '</div>'
    + '      <div class="dr"><span class="dl">Amount in Words:</span> ' + words + '</div>'
    + '    </div>'
    + '  </div>'
    + '  <div class="amt-box">'
    + '    <span class="rs">Rs.</span>'
    + '    <div class="av">' + Number(amount).toLocaleString('en-IN') + '</div>'
    + '  </div>'
    + '  <div class="pm">'
    + '    <strong>Mode of Payment:</strong> NEFT / RTGS / UPI (Cash not eligible u/s 80G)<br><br>'
    + (txnRef && txnRef !== 'N/A' ? '    <strong>Transaction Reference:</strong> ' + txnRef + '<br><br>' : '')
    + '    <strong>Amount in Words:</strong> ' + words
    + '  </div>'
    + '  <div class="decl">'
    + '    Certified that the above donation is received by trust for charitable purposes only.'
    + '    This donation is eligible for deduction under Section 80G of the Income Tax Act, 1961.'
    + '    This receipt will be reported in Form 10BD and Form 10BE will be issued to the donor.'
    + '  </div>'
    + '  <div class="sig">'
    + '    <div class="org-lbl">For Kings Equestrian Foundation</div>'
    + '    <div class="sign-area">'
    + '      <img class="sig-img" src="' + signB64 + '">'
    + '      <img class="stp-img" src="' + stampB64 + '">'
    + '    </div>'
    + '  </div>'
    + '</div>'
    + '</body></html>';

  const tmp  = DriveApp.createFile('receipt_temp_' + Date.now() + '.html', html, MimeType.HTML);
  const blob = tmp.getAs('application/pdf');
  blob.setName('80G_Receipt_' + riderName.replace(/\s+/g,'_') + '_' + receiptNo.replace(/\//g,'-') + '.pdf');
  tmp.setTrashed(true);
  return blob;
}